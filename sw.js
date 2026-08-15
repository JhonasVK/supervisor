// Service worker de la PWA "Supervisor". Cachea lo que el usuario visita para que
// los informes ya vistos abran sin conexion. No usa una lista fija de archivos
// fechados (Dashboard_*_YYYY-MM.html) porque esa lista crece cada mes: en vez de
// eso, cachea cada pagina la primera vez que se visita (con conexion).

const CACHE_VERSION = 'supervisor-v1';
const NUCLEO = ['./index.html', './manifest.json', './logo-cobra.png', './icon-192.png', './icon-512.png'];

self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(CACHE_VERSION).then((cache) => cache.addAll(NUCLEO)).catch(() => {})
  );
  self.skipWaiting();
});

self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys().then((nombres) =>
      Promise.all(nombres.filter((n) => n !== CACHE_VERSION).map((n) => caches.delete(n)))
    )
  );
  self.clients.claim();
});

self.addEventListener('fetch', (event) => {
  const req = event.request;
  if (req.method !== 'GET') return;

  const esMismoOrigen = new URL(req.url).origin === self.location.origin;

  if (esMismoOrigen) {
    // Paginas y datos propios: red primero (para traer lo mas nuevo), y si no hay
    // conexion, se usa lo que haya quedado guardado de la ultima visita.
    event.respondWith(
      fetch(req)
        .then((res) => {
          const copia = res.clone();
          caches.open(CACHE_VERSION).then((cache) => cache.put(req, copia));
          return res;
        })
        .catch(() => caches.match(req))
    );
  } else {
    // Recursos externos (Chart.js por CDN): usar la copia en cache de inmediato
    // si existe, y refrescarla en segundo plano para la proxima vez.
    event.respondWith(
      caches.match(req).then((cacheada) => {
        const redFetch = fetch(req)
          .then((res) => {
            caches.open(CACHE_VERSION).then((cache) => cache.put(req, res.clone()));
            return res;
          })
          .catch(() => cacheada);
        return cacheada || redFetch;
      })
    );
  }
});
