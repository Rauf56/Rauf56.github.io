// Офлайн-оболочка приложения: страница открывается без сети (расписание берётся из кэша / localStorage).
// Оболочка (страница, иконки) — «сначала кэш, обновление в фоне»: свежая версия подтянется при следующем открытии.
// Данные расписания (data/…) — «сначала сеть»: если интернет есть, всегда свежие; нет — последняя сохранённая копия.
const CACHE = "chelgu-app-v4";
const SHELL = ["./", "./index.html", "./manifest.webmanifest", "./icons/icon-192.png", "./icons/icon-512.png",
  "./icons/apple-touch-icon.png", "./icons/favicon-32.png"];
const DATA_TIMEOUT = 6000;

self.addEventListener("install", (e) => {
  e.waitUntil(caches.open(CACHE).then((c) => c.addAll(SHELL)).then(() => self.skipWaiting()));
});

self.addEventListener("activate", (e) => {
  e.waitUntil(caches.keys()
    .then((keys) => Promise.all(keys.filter((k) => k !== CACHE).map((k) => caches.delete(k))))
    .then(() => self.clients.claim()));
});

async function networkFirst(req, url) {
  const cache = await caches.open(CACHE);
  const key = url.origin + url.pathname;
  try {
    const r = await Promise.race([
      fetch(req),
      new Promise((_, rej) => setTimeout(() => rej(new Error("timeout")), DATA_TIMEOUT)),
    ]);
    if (r.ok) cache.put(key, r.clone());
    return r;
  } catch (err) {
    const cached = await cache.match(key);
    return cached || new Response(JSON.stringify({ error: "offline" }), { status: 503, headers: { "Content-Type": "application/json" } });
  }
}

self.addEventListener("fetch", (e) => {
  const req = e.request;
  if (req.method !== "GET") return;
  const url = new URL(req.url);
  if (url.origin !== location.origin) return;   // telegram-web-app.js и всё внешнее — мимо кэша
  if (url.pathname.includes("/data/")) {
    e.respondWith(networkFirst(req, url));
    return;
  }
  e.respondWith((async () => {
    const cache = await caches.open(CACHE);
    const cached = await cache.match(req, { ignoreSearch: true });   // ?d=… / ?g=… у страницы разный, оболочка одна
    const key = req.mode === "navigate" ? new Request(url.origin + url.pathname) : req;
    const fresh = fetch(req).then((r) => { if (r.ok) cache.put(key, r.clone()); return r; }).catch(() => null);
    return cached || (await fresh) || new Response("Нет сети", { status: 503, headers: { "Content-Type": "text/plain; charset=utf-8" } });
  })());
});
