const LEGACY_CACHE_PREFIX = "ao-dai-";

self.addEventListener("install", () => {
  self.skipWaiting();
});

self.addEventListener("activate", (event) => {
  event.waitUntil((async () => {
    if ("caches" in self) {
      const keys = await caches.keys();
      await Promise.all(
        keys
          .filter((key) => key.startsWith(LEGACY_CACHE_PREFIX))
          .map((key) => caches.delete(key))
      );
    }

    await self.registration.unregister();
    await self.clients.claim();
  })());
});
