// Service Worker mínimo — necessário apenas para o navegador considerar
// o app "instalável". Não faz cache agressivo: sempre busca a rede
// primeiro (os dados do painel vêm de um Web App e precisam ser atuais),
// caindo para o cache só se a rede falhar (ex: sem internet).
const CACHE = 'onco-dashboard-v1';
const ASSETS = ['/onco/', '/onco/index.html'];

self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(CACHE).then((cache) => cache.addAll(ASSETS))
  );
  self.skipWaiting();
});

self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys().then((keys) =>
      Promise.all(keys.filter((k) => k !== CACHE).map((k) => caches.delete(k)))
    )
  );
  self.clients.claim();
});

self.addEventListener('fetch', (event) => {
  if (event.request.method !== 'GET') return;
  event.respondWith(
    fetch(event.request)
      .then((resp) => {
        const copy = resp.clone();
        caches.open(CACHE).then((cache) => cache.put(event.request, copy));
        return resp;
      })
      .catch(() => caches.match(event.request))
  );
});
