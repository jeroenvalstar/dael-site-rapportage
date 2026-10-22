// Offline-cache zodat de app ook werkt op locaties met slecht bereik
const CACHE = 'dael-rapportage-v3';
const ASSETS = [
  './',
  'index.html',
  'css/app.css',
  'js/ooxml.js',
  'js/vf-data.js',
  'js/onderaannemers.js',
  'js/forms.js',
  'js/app.js',
  'img/icon.svg',
  'manifest.webmanifest',
  'templates/vodafone-tss-4.20.xlsx',
  'templates/kpn-rct-nieuwbouw-daklocatie.docx',
  'templates/odido-vg-plan-monet-6.2.1-rooftop.docx',
  'templates/cellnex-tssr-reworks.docx',
  'https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js',
];

self.addEventListener('install', (e) => {
  e.waitUntil(caches.open(CACHE).then((c) => c.addAll(ASSETS)).then(() => self.skipWaiting()));
});

self.addEventListener('activate', (e) => {
  e.waitUntil(
    caches.keys()
      .then((keys) => Promise.all(keys.filter((k) => k !== CACHE).map((k) => caches.delete(k))))
      .then(() => self.clients.claim()),
  );
});

// Eerst netwerk (altijd de nieuwste versie), bij geen verbinding de cache
self.addEventListener('fetch', (e) => {
  if (e.request.method !== 'GET') return;
  e.respondWith(
    fetch(e.request)
      .then((res) => {
        if (res.ok && (e.request.url.startsWith(self.location.origin) || e.request.url.includes('cdnjs'))) {
          const copy = res.clone();
          caches.open(CACHE).then((c) => c.put(e.request, copy));
        }
        return res;
      })
      .catch(() => caches.match(e.request, { ignoreSearch: true })),
  );
});
