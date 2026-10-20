"""Данные для самостоятельного приложения (без бота): расписания всех групп ЧелГУ маленькими JSON-файлами.

Запуск: python tools/build_app_data.py [папка]
Обычно его запускает сам GitHub раз в несколько часов (.github/workflows/schedule.yml в репозитории сайта,
кладёт его туда tools/publish_app.py), вручную — python tools/publish_app.py <репозиторий> --with-data.

Что получается в папке:
  groups.json        — {"v":1, "u": "обновлено", "s": начало семестра, "list": [[ключ, шапка группы, факультет], ...]}
  g/<ключ>.json      — {"v":1, "g": шапка, "f": факультет, "s": начало семестра, "p": длительность пары, "u": обновлено,
                        "r": [[день недели 0–6, недели, "ЧЧ:ММ", предмет, аудитория, преподаватель, подгруппа, ссылка,
                               до ММ-ДД, с ММ-ДД, дата ГГГГ-ММ-ДД, вид l/b/p/s], ...]}
Приложение само считает по этим правилам расписание на любой день (чётность недель, «до/с», разовые даты) —
так же, как бот (`lessons_from_rows`, тест app_engine_matches_bot_schedule). Только публичное расписание с сайта
университета, никаких данных бота. Нужны только csu_schedule.py и openpyxl — код бота сюда не попадает.

Осторожность при обновлении (сайт ЧелГУ бывает нестабилен):
  * файл не изменился по существу — не переписывается (дата «u» прежняя, лишних коммитов нет);
  * какой-то факультет не скачался — его группы остаются прежними, ничего не удаляется;
  * группы удаляются, только если скачалось всё и группы действительно пропали с сайта.
"""
import hashlib
import io
import json
import os
import re
import sys
import time
import urllib.request
import zipfile
from datetime import datetime, timedelta, timezone

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
import csu_schedule as cs  # noqa: E402

DEFAULT_PAGE = os.getenv("DEFAULT_PAGE", "https://www.csu.ru/students/timetable/zan/")
SEMESTER = os.getenv("SEMESTER_START", "2026-09-01")
PAIR_MINUTES = int(os.getenv("PAIR_MINUTES", "90"))
HTTP_HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) Chrome/124 Safari/537.36"}
KIND_CODES = (("лекц", "l"), ("лаб", "b"), ("практ", "p"), ("сем", "s"))
MAX_BYTES = 25 * 1024 * 1024
MAX_UNCOMPRESSED = 60 * 1024 * 1024
TZ = timezone(timedelta(hours=5))   # Челябинск, UTC+5 круглый год


def fetch(url: str) -> bytes:
    if not url.startswith(("http://", "https://")):
        raise ValueError(f"не http-ссылка: {url}")
    last = None
    for attempt in (1, 2, 3):
        try:
            req = urllib.request.Request(url, headers=HTTP_HEADERS)
            with urllib.request.urlopen(req, timeout=60) as r:
                data = r.read(MAX_BYTES + 1)
            if len(data) > MAX_BYTES:
                raise ValueError("файл слишком большой")
            return data
        except Exception as e:   # сайт ЧелГУ бывает нестабилен — пробуем ещё
            last = e
            time.sleep(3 * attempt)
    raise last


def check_xlsx(data: bytes) -> None:
    """Защита от «zip-бомбы»: .xlsx — это zip, маленький файл может распаковаться в гигабайты."""
    with zipfile.ZipFile(io.BytesIO(data)) as z:
        if sum(i.file_size for i in z.infolist()) > MAX_UNCOMPRESSED:
            raise ValueError("файл распаковывается в слишком большой объём")


def group_key(title: str, header: str) -> str:
    return hashlib.sha1(f"{title}|{header}".encode("utf-8")).hexdigest()[:10]


def row_json(r: dict, start) -> list:
    raw = r["subject"]
    low = raw.lower()
    kind = next((code for key, code in KIND_CODES if f"({key}" in low), "")
    name = re.sub(r"\s*\((?:лекц|лаб|практ|сем)[^)]*\)", "", cs.pretty_subject(raw), flags=re.I).strip()
    teacher = cs.pretty_teacher(cs.teacher_short(r["teacher"])) if r.get("teacher") else ""
    weekday = r.get("weekday")
    return [weekday if weekday is not None else -1, r.get("week") or "", start.strftime("%H:%M"), name, r.get("room") or "",
            teacher, r.get("subgroup") or "", r.get("link") or "", r.get("until") or "", r.get("since") or "",
            r["date"].isoformat() if r.get("date") else "", kind]


def _write(path: str, obj: dict) -> bool:
    """Пишет JSON, если он изменился по существу (без учёта даты «u»). True — файл переписан."""
    try:
        old = json.load(open(path, encoding="utf-8"))
        if {k: v for k, v in old.items() if k != "u"} == {k: v for k, v in obj.items() if k != "u"}:
            return False
    except (OSError, ValueError):
        pass
    tmp = path + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(obj, f, ensure_ascii=False, separators=(",", ":"))
    os.replace(tmp, path)
    return True


def build_from_books(books: list[tuple[str, bytes]], out_dir: str, semester: str | None = None,
                     failed_titles: tuple = ()) -> dict:
    """books — [(название факультета, содержимое .xlsx)]. Обновляет groups.json и g/*.json в out_dir. Возвращает сводку.
    failed_titles — факультеты, которые сейчас не скачались: их прежние группы сохраняются как есть."""
    semester = semester or SEMESTER
    stamp = datetime.now(TZ).strftime("%d.%m.%Y %H:%M")
    os.makedirs(os.path.join(out_dir, "g"), exist_ok=True)
    try:
        old_list = json.load(open(os.path.join(out_dir, "groups.json"), encoding="utf-8")).get("list", [])
    except (OSError, ValueError):
        old_list = []
    index, lessons, changed = [], 0, 0
    for title, data in books:
        for header, rows in cs.parse_book(data).items():
            out = [row_json(r, cs.PAIR_TIMES[r["pair"]]) for r in rows if r.get("pair") in cs.PAIR_TIMES]
            if not out:
                continue
            key = group_key(title, header)
            index.append([key, header, title])
            lessons += len(out)
            changed += _write(os.path.join(out_dir, "g", f"{key}.json"),
                              {"v": 1, "g": header, "f": title, "s": semester, "p": PAIR_MINUTES, "u": stamp, "r": out})
    keys = {k for k, _, _ in index}
    for item in old_list:   # не скачался факультет — его группы остаются
        if item[2] in failed_titles and item[0] not in keys and os.path.exists(os.path.join(out_dir, "g", f"{item[0]}.json")):
            index.append(item)
            keys.add(item[0])
    removed = 0
    for name in os.listdir(os.path.join(out_dir, "g")):
        if name.endswith(".json") and name[:-5] not in keys:
            os.remove(os.path.join(out_dir, "g", name))
            removed += 1
    index.sort(key=lambda x: (cs.norm_group(cs.FAC_PREFIX.sub("", x[1])), x[2]))
    changed += _write(os.path.join(out_dir, "groups.json"), {"v": 1, "u": stamp, "s": semester, "list": index})
    return {"groups": len(index), "lessons": lessons, "faculties": len(books), "changed": changed, "removed": removed}


def build(out_dir: str, page: str | None = None) -> dict:
    page = page or DEFAULT_PAGE
    html = fetch(page).decode("utf-8", "replace")
    links = cs.discover_links(html, page)
    if not links:
        raise RuntimeError("на странице расписаний не нашлось файлов")
    books, failed = [], []
    for title, url in links:
        try:
            data = fetch(url)
            check_xlsx(data)
            books.append((title, data))
        except Exception as e:
            failed.append((title, str(e)))
    if not books:
        raise RuntimeError("ни один файл не скачался: " + "; ".join(f"{t}: {e}" for t, e in failed))
    summary = build_from_books(books, out_dir, failed_titles=tuple(t for t, _ in failed))
    summary["failed"] = [f"{t}: {e}" for t, e in failed]
    return summary


if __name__ == "__main__":
    out = sys.argv[1] if len(sys.argv) > 1 else os.path.join(ROOT, "build", "app_data")
    s = build(out)
    print(f"Готово: {s['groups']} групп, {s['lessons']} занятий, факультетов {s['faculties']}, "
          f"изменено файлов {s['changed']}, удалено {s['removed']} → {out}")
    if s["failed"]:
        print("Не скачались:", "; ".join(s["failed"]))
