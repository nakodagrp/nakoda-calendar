// Service Worker for Team Calendar & Task Manager PWA
const CACHE_NAME = 'team-cal-v9';
const ASSETS = [
  'index.html',
  'TeamCalendar.html',
  'manifest.json',
  'https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap',
  'https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css'
];

// Install
self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then(cache => cache.addAll(ASSETS)).then(() => self.skipWaiting())
  );
});

// Activate
self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys().then(keys =>
      Promise.all(keys.filter(k => k !== CACHE_NAME).map(k => caches.delete(k)))
    ).then(() => self.clients.claim())
  );
});

// Fetch - Network first, fallback to cache
self.addEventListener('fetch', (event) => {
  if (event.request.method !== 'GET' || event.request.url.includes('script.google.com')) {
    return;
  }

  event.respondWith(
    fetch(event.request)
      .then(response => {
        const clone = response.clone();
        caches.open(CACHE_NAME).then(cache => cache.put(event.request, clone));
        return response;
      })
      .catch(() => caches.match(event.request))
  );
});

// Message from main app (notification fallback for mobile)
self.addEventListener('message', (event) => {
  if (event.data && event.data.type === 'notify') {
    self.registration.showNotification(event.data.title, {
      body: event.data.body,
      icon: 'icon-192.png',
      badge: 'icon-192.png',
      vibrate: [200, 100, 200],
      tag: event.data.title,
      renotify: true
    });
  }
});

// Push notifications
self.addEventListener('push', (event) => {
  const data = event.data ? event.data.json() : { title: 'Task Alert', body: 'A task needs your attention!' };
  event.waitUntil(
    self.registration.showNotification(data.title, {
      body: data.body,
      icon: '📋',
      badge: '🚨',
      vibrate: [200, 100, 200],
      tag: 'task-alert',
      renotify: true
    })
  );
});

// Notification click - open the app
self.addEventListener('notificationclick', (event) => {
  event.notification.close();
  event.waitUntil(
    clients.matchAll({ type: 'window' }).then(list => {
      if (list.length > 0) {
        list[0].focus();
      } else {
        clients.openWindow('TeamCalendar.html');
      }
    })
  );
});
