const CACHE_NAME = 'reportgen-v3';
const ASSETS = [
  './',
  './index.html',
  './style.css',
  './js/report-generator.js',
  './js/app.js',
  './manifest.json',
  './Templates/MVPanel.docx',
  './Templates/Ct1.docx',
  'https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js',
  'https://cdn.jsdelivr.net/npm/pizzip@3.1.7/dist/pizzip.min.js',
  'https://cdn.jsdelivr.net/npm/docxtemplater@3.49.1/build/docxtemplater.js',
  'https://cdn.jsdelivr.net/npm/jszip@3.10.1/dist/jszip.min.js',
  'https://cdn.jsdelivr.net/npm/file-saver@2.0.5/dist/FileSaver.min.js'
];

self.addEventListener('install', (e) => {
  e.waitUntil(caches.open(CACHE_NAME).then((cache) => cache.addAll(ASSETS)));
});

self.addEventListener('fetch', (e) => {
  e.respondWith(
    caches.match(e.request).then((res) => res || fetch(e.request))
  );
});