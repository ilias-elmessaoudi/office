// Offline-Speicher von KXD Office – nur fuer den eigenen Unterordner (/office/).
// Ersetzt den Speicher des Upstream-Projekts, der fest an der Wurzel der Adresse haengt.
// Editor-Dateien (sdkjs, web-apps, fonts) liegen in einem Zwischenspeicher, der nur bei neuem Editor-Kern
// wechselt; alles andere in einem, der bei jeder Veroeffentlichung wechselt. Beide Namen setzt office-a-paket.py.
const KERN = 'kxd-office-kern-e975654';
const APP = 'kxd-office-app-20261101132437';
const BASIS = new URL('./', self.location.href).pathname;          // z. B. /office/
const GRUNDAUSSTATTUNG = ['./', 'index.html', 'editor.html', 'manifest.json', 'kxd/app.js', 'kxd/app.css',
  'kxd/speicher.js', 'kxd/abgleich.js', 'kxd/vendor/supabase.min.js', 'kxd/icon.svg', 'kxd/icon-192.png',
  'kxd/vorlagen/leer.docx', 'kxd/vorlagen/leer.xlsx', 'kxd/vorlagen/leer.pptx'];

self.addEventListener('install', (e) => {
  e.waitUntil(caches.open(APP).then((c) => c.addAll(GRUNDAUSSTATTUNG)).then(() => self.skipWaiting()));
});
self.addEventListener('activate', (e) => {
  e.waitUntil(caches.keys()
    .then((namen) => Promise.all(namen.filter((n) => n.startsWith('kxd-office-') && n !== KERN && n !== APP).map((n) => caches.delete(n))))
    .then(() => self.clients.claim()));
});

const istKern = (pfad) => /^(sdkjs|web-apps|fonts)\//.test(pfad);

self.addEventListener('fetch', (e) => {
  const anfrage = e.request;
  if (anfrage.method !== 'GET') return;
  const url = new URL(anfrage.url);
  if (url.origin !== self.location.origin || !url.pathname.startsWith(BASIS)) return;
  if (url.searchParams.has('src') || url.searchParams.has('file')) return;
  // Rechtschreib-Baustein nicht abfangen: haengt sonst beim allerersten Start (Erfahrung des Upstream-Projekts)
  if (url.pathname.includes('/sdkjs/common/spell/')) return;
  const pfad = url.pathname.slice(BASIS.length);
  const seite = anfrage.mode === 'navigate' || pfad === '' || pfad.endsWith('.html');

  if (seite) {
    // Seiten: erst Netz (immer neueste Fassung), ohne Netz aus dem Speicher
    e.respondWith(fetch(anfrage).then((antwort) => {
      if (antwort.ok) { const kopie = antwort.clone(); caches.open(APP).then((c) => c.put(anfrage, kopie)); }
      return antwort;
    }).catch(() => caches.match(anfrage, { ignoreSearch: true }).then((r) => r || caches.match(BASIS + 'index.html'))));
    return;
  }
  // Alles andere: erst Speicher, sonst Netz und merken
  e.respondWith(caches.match(anfrage, { ignoreSearch: true }).then((r) => r || fetch(anfrage).then((antwort) => {
    if (antwort.ok && antwort.type === 'basic') {
      const kopie = antwort.clone();
      caches.open(istKern(pfad) ? KERN : APP).then((c) => c.put(anfrage, kopie));
    }
    return antwort;
  })));
});
