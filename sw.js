/* Timenest service worker — app shell cache. API calls (Apps Script POST) are never cached. */
const CACHE = "timenest-v1";
const SHELL = ["./", "./index.html", "./manifest.webmanifest",
  "./icons/icon-192.png", "./icons/icon-512.png", "./icons/maskable-512.png", "./icons/apple-touch-icon.png"];

self.addEventListener("install", e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(SHELL)).then(() => self.skipWaiting()));
});
self.addEventListener("activate", e => {
  e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => k !== CACHE).map(k => caches.delete(k)))).then(() => self.clients.claim()));
});
self.addEventListener("fetch", e => {
  const r = e.request;
  if (r.method !== "GET") return;                       // POSTs to Apps Script go straight to network
  const u = new URL(r.url);
  if (u.hostname === "script.google.com" || u.hostname === "script.googleusercontent.com") return;
  if (r.mode === "navigate") {                          // network-first so updates show up; cache as offline fallback
    e.respondWith(fetch(r).then(res => { caches.open(CACHE).then(c => c.put("./index.html", res.clone())); return res; })
      .catch(() => caches.match("./index.html")));
    return;
  }
  e.respondWith(caches.match(r).then(hit => hit || fetch(r).then(res => {   // static + fonts: cache-first
    if (res.ok && (u.origin === location.origin || u.hostname.endsWith("gstatic.com") || u.hostname.endsWith("googleapis.com")))
      caches.open(CACHE).then(c => c.put(r, res.clone()));
    return res;
  })));
});
