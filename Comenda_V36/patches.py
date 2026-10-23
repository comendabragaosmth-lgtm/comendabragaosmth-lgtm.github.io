# Alterações V36 (velocidade). Cada par [antigo, novo] tem de aparecer exatamente uma vez.
import json

SERVIDOR = [
[r"""function cacheTab_(nome) { return nome === 'Fila' ? null : 'tab_' + nome; }
function limparCache_(nome) { var k = cacheTab_(nome); if (k) { try { CacheService.getScriptCache().remove(k); } catch (e) {} } }
function ler_(nome) {
  var ck = cacheTab_(nome);
  if (ck) { try { var c = CacheService.getScriptCache().get(ck); if (c) return JSON.parse(c); } catch (e) {} }
  var out = lerFolha_(nome);
  if (ck) { try { var j = JSON.stringify(out); if (j.length < 95000) CacheService.getScriptCache().put(ck, j, 300); } catch (e) {} }
  return out;
}""",
r"""/* Cache das folhas (V36): cada folha fica guardada em blocos durante até 6 h e é limpa sempre que a app escreve nela.
 * Se alguém editar a Folha Google à mão, a cache é limpa no máximo 1 minuto depois (data de alteração no Drive). */
var CACHE_SEG = 21600, CACHE_BLOCO = 32000, _mem = {};
function cacheTab_(nome) { return nome === 'Fila' ? null : 'tab_' + nome; }
function limparCache_(nome) {
  var k = cacheTab_(nome); if (!k) return;
  delete _mem[k];
  try { CacheService.getScriptCache().remove(k); } catch (e) {}
}
function verificarFolha_() {
  if (_mem._verif) return; _mem._verif = 1;
  try {
    var c = CacheService.getScriptCache(); if (c.get('folha_chk')) return;
    c.put('folha_chk', '1', 60);
    var t = String(DriveApp.getFileById(PropertiesService.getScriptProperties().getProperty('SHEET_ID')).getLastUpdated().getTime()), ant = c.get('folha_alt');
    if (ant && ant !== t) { c.removeAll(Object.keys(TABELAS).map(cacheTab_).filter(Boolean)); _mem = { _verif: 1 }; }
    c.put('folha_alt', t, CACHE_SEG);
  } catch (e) {}
}
function lerCache_(ck) {
  try {
    var c = CacheService.getScriptCache(), cab = c.get(ck);
    if (!cab || cab.indexOf('n:') !== 0) return null;
    var n = +cab.slice(2), ks = [];
    for (var i = 0; i < n; i++) ks.push(ck + '_' + i);
    var v = c.getAll(ks), out = '';
    for (var k = 0; k < n; k++) { if (v[ks[k]] == null) return null; out += v[ks[k]]; }
    return out;
  } catch (e) { return null; }
}
function guardarCache_(ck, j) {
  try {
    var n = Math.ceil(j.length / CACHE_BLOCO) || 1, o = {};
    if (n > 100) return;
    for (var i = 0; i < n; i++) o[ck + '_' + i] = j.slice(i * CACHE_BLOCO, (i + 1) * CACHE_BLOCO);
    var c = CacheService.getScriptCache(); c.putAll(o, CACHE_SEG); c.put(ck, 'n:' + n, CACHE_SEG);
  } catch (e) {}
}
function ler_(nome) {
  var ck = cacheTab_(nome);
  if (!ck) return lerFolha_(nome);
  verificarFolha_();
  var j = _mem[ck];
  if (j == null) j = lerCache_(ck);
  if (j == null) { j = JSON.stringify(lerFolha_(nome)); guardarCache_(ck, j); }
  _mem[ck] = j;
  return JSON.parse(j);
}"""],
[r"""function apiDados(token) {
  var subs = ler_('Inscritos');
  var me = token ? subs.filter(""",
r"""function apiDados(token) {
  var me = token ? ler_('Inscritos').filter("""],
[r"""function processarFila(max) {
  if (typeof max !== 'number') { try { notifAgendadas_(); } catch (err) {} }""",
r"""function processarFila(max) {
  if (typeof max !== 'number') { try { notifAgendadas_(); } catch (err) {} try { ['Definicoes', 'Eventos', 'Inscritos', 'Inscricoes', 'Convidados', 'Recusas', 'Membros', 'Dispositivos'].forEach(ler_); } catch (err) {} }"""],
]

PAGINA = [
[r"""const lsS=(k,v)=>{try{v?localStorage.setItem(k,v):localStorage.removeItem(k)}catch(e){}};""",
r"""const lsS=(k,v)=>{try{v?localStorage.setItem(k,v):localStorage.removeItem(k)}catch(e){}};
const DK='comenda_dados';
const dGet=()=>{try{const x=JSON.parse(localStorage.getItem(DK)||'null');return x&&x.tk===(B.token||'')&&x.d&&Array.isArray(x.d.eventos)?x.d:null}catch(e){return null}};
const dSet=d=>{try{d?localStorage.setItem(DK,JSON.stringify({tk:B.token||'',d})):localStorage.removeItem(DK)}catch(e){}};
const syncOn=v=>{try{if(v)document.body.dataset.sync=L('A atualizar…','Updating…');else delete document.body.dataset.sync}catch(e){}};"""],
[r"""async function loadPublic(){try{D=await call('apiDados',B.token||'')}catch(e){D={org:B.org,hoje:iso(new Date()),eventos:[],me:null};toast(errMsg(e))}""",
r"""async function loadPublic(){let ok=false;try{D=await call('apiDados',B.token||'');ok=true}catch(e){D=dGet()||{org:B.org,hoje:iso(new Date()),eventos:[],me:null};toast(errMsg(e))}"""],
[r"""  if(B.token&&!D.me&&ui.mode!=='sair'){if(B.token===lsG(TK))lsS(TK,'');B.token=''}
}""",
r"""  if(B.token&&!D.me&&ui.mode!=='sair'){if(B.token===lsG(TK))lsS(TK,'');B.token='';dSet(null)}
  else if(ok)dSet(D);
}"""],
[r"""  else await loadPublic();
  render();""",
r"""  else{const c=!B.confirmar&&!B.insc&&ui.mode!=='sair'?dGet():null;
    if(c){D=c;D.hoje=iso(new Date());if(!ui.month)ui.month=D.hoje.slice(0,7);syncOn(true);render()}
    await loadPublic();syncOn(false)}
  render();"""],
[r"""case 'sair': r=await act('apiSair',B.token);if(r){ui.msg=r;D&&(D.me=null)}break;""",
r"""case 'sair': r=await act('apiSair',B.token);if(r){ui.msg=r;D&&(D.me=null);dSet(null)}break;"""],
[r"""body{background:var(--bg);color:var(--ink);font-family:var(--f-body);font-size:16px;line-height:1.5}""",
r"""body{background:var(--bg);color:var(--ink);font-family:var(--f-body);font-size:16px;line-height:1.5}
body[data-sync]::after{content:attr(data-sync);position:fixed;left:50%;bottom:14px;transform:translateX(-50%);background:var(--ink);color:#fff;font-size:12px;padding:5px 12px;border-radius:20px;opacity:.85;z-index:60;pointer-events:none}"""],
]

SW = [
[r"""const VERSAO = 'comenda-v2';""", r"""const VERSAO = 'comenda-v3';"""],
[r"""  if (req.mode === 'navigate') {
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
  }""",
r"""  if (req.mode === 'navigate') {
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
  }"""],
]

def aplicar(texto, pares, nome):
    for a, n in pares:
        c = texto.count(a)
        if c != 1:
            raise SystemExit('%s: âncora aparece %d vezes: %r' % (nome, c, a[:80]))
        texto = texto.replace(a, n)
    return texto

if __name__ == '__main__':
    import sys
    o = 'orig/'
    open('real/Code.gs', 'w', encoding='utf-8').write(aplicar(open(o + 'Code.gs', encoding='utf-8').read(), SERVIDOR, 'Code.gs'))
    open('real/Index.html', 'w', encoding='utf-8').write(aplicar(open(o + 'Index.html', encoding='utf-8').read(), PAGINA, 'Index.html'))
    open('real/gh_index.html', 'w', encoding='utf-8').write(aplicar(open(o + 'gh_index.html', encoding='utf-8').read(), PAGINA, 'gh_index.html'))
    open('real/sw.js', 'w', encoding='utf-8').write(aplicar(open(o + 'sw.js', encoding='utf-8').read(), SW, 'sw.js'))
    json.dump({'servidor': SERVIDOR, 'pagina': PAGINA, 'sw': SW}, open('patches.json', 'w', encoding='utf-8'), ensure_ascii=False)
    print('ok')
