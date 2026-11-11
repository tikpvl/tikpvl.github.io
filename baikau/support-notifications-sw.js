self.addEventListener('install', event => {
  event.waitUntil(self.skipWaiting());
});

self.addEventListener('notificationclick', event => {
  event.notification.close();
  const data = event.notification.data || {};

  event.waitUntil((async () => {
    if (typeof data.url !== 'string') return;

    const target = new URL(data.url, self.location.origin);
    if (target.origin !== self.location.origin) return;

    const windows = await self.clients.matchAll({
      type: 'window',
      includeUncontrolled: true
    });

    const existing = windows.find(client =>
      new URL(client.url).pathname === target.pathname
    );

    if (existing) {
      existing.postMessage({
        type: 'support-chat-open',
        room: data.room
      });
      await existing.focus();
    } else {
      await self.clients.openWindow(target.href);
    }
  })());
});