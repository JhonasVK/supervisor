// Service worker de LIMPIEZA (09-10-2026). La app antigua "Supervisor" guardaba
// en el celular copias de los informes sin clave. Desde que el portal va cifrado
// detras del ingreso, este archivo reemplaza al anterior: al actualizarse en cada
// equipo borra esas copias, se da de baja y recarga las paginas abiertas.
self.addEventListener('install', () => self.skipWaiting());
self.addEventListener('activate', (event) => {
  event.waitUntil((async () => {
    const nombres = await caches.keys();
    await Promise.all(nombres.map((n) => caches.delete(n)));
    await self.registration.unregister();
    const ventanas = await self.clients.matchAll({ type: 'window' });
    ventanas.forEach((v) => v.navigate(v.url));
  })());
});
