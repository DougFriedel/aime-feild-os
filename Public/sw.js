// AIME Field Pro — service worker
//
// Deliberately does NOT cache the app. Browsers require a service worker with a
// fetch handler before they'll offer "Install", and that's all this provides.
//
// Caching the app shell here would be actively harmful: every deploy would leave
// crews running a stale build with no obvious way to update, which is exactly the
// class of problem that's already cost time on this project. The app has its own
// offline queue for field work, so it doesn't need the cache.

const VERSION = 'aime-v1';

self.addEventListener('install', (e) => {
  self.skipWaiting();               // take over immediately on deploy
});

self.addEventListener('activate', (e) => {
  e.waitUntil((async () => {
    // clear anything a previous version may have cached
    const keys = await caches.keys();
    await Promise.all(keys.map((k) => caches.delete(k)));
    await self.clients.claim();
  })());
});

self.addEventListener('fetch', (e) => {
  // Straight to the network. Present only to satisfy the install requirement.
  return;
});
