// Concept Time Mobile has moved to CG HUBS (https://conceptgrouppng.github.io/CG_HUBS/mobile/).
// This replacement service worker removes the old one: it deletes the
// cached copy of the old app, unregisters itself, and reloads any open
// window so it shows the "This Hub has moved" page.
self.addEventListener("install", () => self.skipWaiting());
self.addEventListener("activate", (event) => {
  event.waitUntil((async () => {
    const names = await caches.keys();
    await Promise.all(names.map((n) => caches.delete(n)));
    await self.registration.unregister();
    const wins = await self.clients.matchAll({ type: "window" });
    wins.forEach((w) => w.navigate(w.url));
  })());
});
