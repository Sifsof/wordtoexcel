/*
 * sw.js — service worker de WordToExcel (GitHub Pages).
 * Toutes les adresses sont relatives à l'emplacement de ce fichier :
 * l'application fonctionne aussi bien à la racine qu'en sous-dossier (/wordtoexcel/).
 *
 * Mise à jour : changer CACHE_VERSION à chaque publication. Le nouveau service worker
 * attend que l'utilisateur accepte « Mettre à jour » dans l'application (aucun rechargement surprise).
 */
const CACHE_VERSION = 'wordtoexcel-gh-1.0.1';

const PRECACHE = [
  './',
  './index.html',
  './manifest.webmanifest',
  './favicon.ico',
  './icon-192.png',
  './icon-512.png',
  './icon-maskable-512.png',
  './style.css',
  './responsive.css',
  './config.js',
  './app.js',
  './docx-parser.js',
  './project-model.js',
  './excel-export.js',
  './jszip.min.js',
  './xlsx.bundle.js'
];

const SCOPE_PATH = new URL('./', self.location.href).pathname;

self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(CACHE_VERSION).then((cache) =>
      cache.addAll(PRECACHE.map((p) => new URL(p, self.location.href).href))
    )
  );
});

self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys()
      .then((keys) => Promise.all(keys.filter((k) => k !== CACHE_VERSION).map((k) => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

// Message envoyé par la page quand l'utilisateur accepte la mise à jour.
self.addEventListener('message', (event) => {
  if (event.data && event.data.type === 'SKIP_WAITING') self.skipWaiting();
});

// Réseau d'abord (versions à jour quand Internet est disponible), cache ensuite (hors ligne).
async function networkFirst(request) {
  const cache = await caches.open(CACHE_VERSION);
  try {
    const response = await fetch(request);
    if (response && response.ok && response.type === 'basic') {
      cache.put(request, response.clone());
    }
    return response;
  } catch (err) {
    const cached = await cache.match(request, { ignoreSearch: true });
    if (cached) return cached;
    if (request.mode === 'navigate') {
      const home = await cache.match(new URL('./index.html', self.location.href).href);
      if (home) return home;
    }
    return Response.error();
  }
}

self.addEventListener('fetch', (event) => {
  const request = event.request;
  if (request.method !== 'GET') return;
  const url = new URL(request.url);
  if (url.origin !== self.location.origin) return;
  if (!url.pathname.startsWith(SCOPE_PATH)) return;
  event.respondWith(networkFirst(request));
});
