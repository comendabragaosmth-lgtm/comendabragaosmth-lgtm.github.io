/* Service worker da app da Comenda de Braga.
 * - Páginas: rede primeiro; sem rede, mostra a última versão guardada.
 * - Ícones: cache primeiro; pwa.js e manifesto: rede primeiro (as atualizações chegam logo).
 * - Pedidos ao servidor (Google Apps Script) nunca passam pela cache.
 * Ao alterar este ficheiro, aumentar a VERSAO para os telemóveis atualizarem.
 */
const VERSAO = 'comenda-v3';
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
    // Página da app: abre logo a versão guardada e vai buscar a nova em segundo plano (fica para a próxima abertura).
    const raiz = /\/(index\.html)?$/.test(url.pathname);
    const daRede = fetch(req).then(res => {
      if (res.ok) { const copia = res.clone(); caches.open(VERSAO).then(c => c.put('./', copia)); }
      return res;
    });
    if (!raiz) { e.respondWith(daRede.catch(() => caches.match('./'))); return; }
    e.waitUntil(daRede.catch(() => {}));
    e.respondWith(caches.match('./').then(hit => hit || daRede).catch(() => daRede));
    return;
  }

  // Imagens: cache primeiro. Restantes ficheiros (pwa.js, manifesto…): rede primeiro, cache sem rede.
  const imagem = /\.(png|jpe?g|svg|ico)$/i.test(url.pathname);
  const daRede = () => fetch(req).then(res => {
    if (res.ok) { const copia = res.clone(); caches.open(VERSAO).then(c => c.put(req, copia)); }
    return res;
  });
  e.respondWith(imagem
    ? caches.match(req).then(hit => hit || daRede())
    : daRede().catch(() => caches.match(req)));
});

/* ---------- notificações (enviadas pelo servidor através do Firebase) ---------- */
self.addEventListener('push', e => {
  let d = {};
  try { d = e.data ? e.data.json() : {}; } catch (err) { d = { data: { corpo: e.data && e.data.text() } }; }
  const m = d.data || {};
  const n = d.notification || {};
  const titulo = m.titulo || n.title || 'Comenda de Braga';
  e.waitUntil(self.registration.showNotification(titulo, {
    body: m.corpo || n.body || '',
    icon: 'icon-192.png',
    badge: 'favicon-32.png',
    tag: m.tag || undefined,
    data: { url: m.url || './' }
  }));
});

self.addEventListener('notificationclick', e => {
  e.notification.close();
  const url = new URL((e.notification.data && e.notification.data.url) || './', self.location.href).href;
  e.waitUntil((async () => {
    const lista = await self.clients.matchAll({ type: 'window', includeUncontrolled: true });
    for (const c of lista) {
      if (new URL(c.url).origin === self.location.origin && 'focus' in c) {
        await c.focus();
        if ('navigate' in c) { try { await c.navigate(url); } catch (err) {} }
        return;
      }
    }
    if (self.clients.openWindow) await self.clients.openWindow(url);
  })());
});
