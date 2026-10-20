/* Service worker da app da Comenda de Braga.
 * - Páginas: rede primeiro; sem rede, mostra a última versão guardada.
 * - Ícones e ficheiros estáticos: cache primeiro.
 * - Pedidos ao servidor (Google Apps Script) nunca passam pela cache.
 * Ao alterar este ficheiro, aumentar a VERSAO para os telemóveis atualizarem.
 */
const VERSAO = 'comenda-v1';
const ESSENCIAIS = ['./', './pwa.js', './manifest.webmanifest', './icon-192.png', './apple-touch-icon.png'];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(VERSAO).then(c => c.addAll(ESSENCIAIS)).then(() => self.skipWaiting()));
});

self.addEventListener('activate', e => {
  e.waitUntil(
    caches.keys()
      .then(ks => Promise.all(ks.filter(k => k !== VERSAO).map(k => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', e => {
  const req = e.request;
  if (req.method !== 'GET') return;
  const url = new URL(req.url);
  if (url.origin !== location.origin) return; // Apps Script, Google Fonts, etc.: sempre pela rede

  if (req.mode === 'navigate') {
    e.respondWith(
      fetch(req)
        .then(res => {
          const copia = res.clone();
          caches.open(VERSAO).then(c => c.put('./', copia));
          return res;
        })
        .catch(() => caches.match('./'))
    );
    return;
  }

  e.respondWith(
    caches.match(req).then(hit => hit || fetch(req).then(res => {
      if (res.ok) { const copia = res.clone(); caches.open(VERSAO).then(c => c.put(req, copia)); }
      return res;
    }))
  );
});
