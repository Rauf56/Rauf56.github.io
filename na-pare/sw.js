// Офлайн-оболочка приложения: страница открывается без сети (расписание берётся из localStorage).
// «Сначала кэш, обновление в фоне»: свежая версия страницы подтянется при следующем открытии.
const CACHE = "chelgu-app-v2";
const SHELL = ["./", "./index.html", "./manifest.webmanifest", "./icons/icon-192.png", "./icons/icon-512.png",
  "./icons/apple-touch-icon.png", "./icons/favicon-32.png"];

self.addEventListener("install", (e) => {
  e.waitUntil(caches.open(CACHE).then((c) => c.addAll(SHELL)).then(() => self.skipWaiting()));
});

self.addEventListener("activate", (e) => {
  e.waitUntil(caches.keys()
    .then((keys) => Promise.all(keys.filter((k) => k !== CACHE).map((k) => caches.delete(k))))
    .then(() => self.clients.claim()));
});

self.addEventListener("fetch", (e) => {
  const req = e.request;
  if (req.method !== "GET") return;
  const url = new URL(req.url);
  if (url.origin !== location.origin) return;   // telegram-web-app.js и всё внешнее — мимо кэша
  e.respondWith((async () => {
    const cache = await caches.open(CACHE);
    const cached = await cache.match(req, { ignoreSearch: true });   // ?d=… у страницы разный, оболочка одна
    const key = req.mode === "navigate" ? new Request(url.origin + url.pathname) : req;
    const fresh = fetch(req).then((r) => { if (r.ok) cache.put(key, r.clone()); return r; }).catch(() => null);
    return cached || (await fresh) || new Response("Нет сети", { status: 503, headers: { "Content-Type": "text/plain; charset=utf-8" } });
  })());
});
