/**
 * App de eventos — Comenda de Braga (OSMTH)
 * Google Apps Script + Google Sheets + Gmail
 *
 * INSTALAÇÃO: correr uma vez a função «instalar» no editor (autorizar quando pedido)
 * e depois publicar como aplicação Web (Executar como: eu · Acesso: qualquer pessoa).
 */

var CONFIG = {
  NOME_ORG: 'Comenda de Braga',
  ADMINS: ['fgscarvalho@gmail.com', 'peixoto.consulting@gmail.com', 'leafar2belen@gmail.com'],
  TZ: 'Europe/Lisbon',
  HORA_LEMBRETES: 9,
  HORAS_SESSAO_ADMIN: 12,
  MIN_PASSWORD: 8,
  ENVIO_IMEDIATO_MAX: 10,
  URL_SITE: 'https://comendabragaosmth-lgtm.github.io/',
  URL_API: 'https://script.google.com/macros/s/AKfycbz1SMPBc6hPkO1bUxZsFZHUeLNSIBI9yckW5h8qL9F33TMes5HArk82c-LU79qrLDpr/exec',
  EMAIL_ORG: 'comendabraga.osmth@gmail.com',
  DURACAO_PADRAO_HORAS: 3
};

var TABELAS = {
  Inscritos: ['email', 'nome', 'telemovel', 'localidade', 'nota', 'estado', 'membro', 'token', 'criado_em', 'confirmado_em', 'ultimo_envio_confirmacao', 'n_ordem', 'lingua', 'verificar', 'cal_ligado'],
  Eventos: ['id', 'titulo', 'data', 'hora', 'local', 'descricao', 'visibilidade', 'prazo', 'lotacao', 'campos', 'atividades', 'estado', 'lembrete7', 'lembrete1', 'criado_em', 'atualizado_em', 'hora_fim', 'seq', 'album', 'agradecimento', 'titulo_en', 'descricao_en', 'atividades_en', 'externo', 'externo_link', 'pasta_id', 'pasta_url', 'morada', 'mapa', 'programa', 'traje', 'refeicao', 'programa_en', 'traje_en', 'video_estado', 'video_id', 'video_url', 'video_info', 'video_em'],
  Inscricoes: ['id', 'evento_id', 'evento', 'email', 'tipo', 'nome', 'telemovel', 'agape', 'restricoes', 'atividades', 'observacoes', 'n_convidados', 'atualizado_em', 'visita', 'nif', 'motivo_alteracao'],
  Convidados: ['inscricao_id', 'evento_id', 'evento', 'convidado_de', 'nome', 'email', 'telemovel', 'agape', 'restricoes', 'atividades', 'visita'],
  Membros: ['n_ordem', 'nome', 'email', 'telemovel', 'localidade', 'grau', 'data_admissao', 'nif', 'notas', 'criado_em', 'atualizado_em', 'direcao', 'tipo', 'cargo'],
  HistGraus: ['data', 'n_ordem', 'nome', 'campo', 'antes', 'depois', 'por', 'registado_em'],
  Notificacoes: ['id', 'criado_em', 'por', 'assunto', 'texto', 'botao_texto', 'botao_url', 'evento_id', 'grupos', 'pessoas', 'agendada_para', 'estado', 'enviada_em', 'n'],
  Recusas: ['evento_id', 'email', 'nome', 'nota', 'criado_em', 'tipo'],
  Presencas: ['evento_id', 'chave', 'tipo', 'nome', 'email', 'n_ordem', 'membro', 'localidade', 'telemovel', 'convidado_de', 'sem_inscricao', 'marcado_por', 'marcado_em'],
  PreInscricoes: ['token', 'email', 'evento_id', 'dados', 'criado_em'],
  Historico: ['data', 'email', 'nome', 'evento_id', 'evento', 'acao', 'motivo'],
  Fila: ['id', 'para', 'assunto', 'html', 'prioridade', 'criado_em', 'enviado_em', 'erro', 'ics'],
  Definicoes: ['chave', 'valor']
};

/* ============ Instalação ============ */

function instalar() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('SHEET_ID');
  var ss;
  if (id) { ss = SpreadsheetApp.openById(id); }
  else {
    ss = SpreadsheetApp.create(CONFIG.NOME_ORG + ' – App de eventos');
    props.setProperty('SHEET_ID', ss.getId());
  }
  Object.keys(TABELAS).forEach(function (nome) {
    var sh = ss.getSheetByName(nome) || ss.insertSheet(nome);
    var cab = TABELAS[nome];
    sh.getRange(1, 1, 1, cab.length).setValues([cab]).setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.getRange(1, 1, Math.max(sh.getMaxRows(), 2), cab.length).setNumberFormat('@');
  });
  var def = ss.getSheetByName('Sheet1') || ss.getSheetByName('Folha1') || ss.getSheetByName('Página1');
  if (def && ss.getSheets().length > 1) ss.deleteSheet(def);
  if (!getDef_('nome_org')) setDef_('nome_org', CONFIG.NOME_ORG);
  if (!getDef_('admins')) setDef_('admins', CONFIG.ADMINS.join(', '));
  ScriptApp.getProjectTriggers().forEach(function (t) { ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('tarefaDiaria').timeBased().everyDays(1).atHour(CONFIG.HORA_LEMBRETES).inTimezone(CONFIG.TZ).create();
  ScriptApp.newTrigger('processarFila').timeBased().everyMinutes(10).create();
  MailApp.getRemainingDailyQuota();
  Logger.log('Instalação concluída. Folha: ' + ss.getUrl());
  return ss.getUrl();
}

/* ============ Web ============ */

function doGet(e) {
  var p = (e && e.parameter) || {};
  if (p.ical) {
    return ContentService.createTextOutput(feedIcal_(String(p.ical))).setMimeType(ContentService.MimeType.ICAL);
  }
  guardarUrl_();
  if (p.pagina) {
    return ContentService.createTextOutput(HtmlService.createHtmlOutputFromFile('Index').getContent()).setMimeType(ContentService.MimeType.TEXT);
  }
  var destino = siteUrl_() + (e && e.queryString ? '?' + e.queryString : '');
  return HtmlService.createHtmlOutput('<p style="font-family:sans-serif">A abrir o calendário da Comenda… <a href="' + esc_(destino) + '" target="_top">Carregue aqui se não abrir.</a></p><script>window.top.location.href=' + JSON.stringify(destino) + ';</script>')
    .setTitle('Comenda de Braga · Calendário').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

var API_PUBLICA = ['apiDados', 'apiSubscrever', 'apiInscrever', 'apiAnular', 'apiSair', 'apiConfirmar', 'apiAdminLogin', 'apiAdminPedirPassword', 'apiAdminDefinirPassword', 'apiAdminSair', 'apiAdminDados', 'apiAdminGuardarEvento', 'apiAdminCancelar', 'apiAdminNovidade', 'apiAdminLembrete', 'apiAdminMembro', 'apiAdminRemover', 'apiAdminConvidar', 'apiAdminReenviar', 'apiAdminDefinicoes', 'apiAdminGuardarMembro', 'apiAdminRemoverMembro', 'apiAdminImportarMembros', 'apiAdminTornarAdmin', 'apiAdminPublicar', 'apiAdminOcultar', 'apiAdminEliminar', 'apiEventoLink', 'apiInscreverPublico', 'apiConfirmarInscricao', 'apiLingua', 'apiAdminNotificar', 'apiAdminPreverImportacao', 'apiVideo', 'apiAdminVideo', 'apiAdminChaveVideo', 'apiAdminEnviarAbertos', 'apiAdminEditarInscrito', 'apiRecusar', 'apiAdminInscreverPessoas', 'apiAdminPresencas', 'apiAdminMudarGrau', 'apiAdminConvidadosBase', 'apiAdminNotifPrever', 'apiAdminNotifEnviar', 'apiAdminNotifCancelar', 'apiEntrarPedirCodigo', 'apiEntrarCodigo', 'apiPushConfig', 'apiPushRegistar', 'apiPushRemover'];
function doPost(e) {
  var out;
  try {
    var req = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (API_PUBLICA.indexOf(req.fn) < 0) throw new Error('Pedido inválido.');
    var fn = globalThis[req.fn];
    out = { ok: true, data: fn.apply(null, Array.isArray(req.args) ? req.args : []) };
  } catch (err) {
    out = { ok: false, error: String(err && err.message || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}
function apiConfirmar(token) { return confirmar_(String(token || '')); }
function siteUrl_() { return getDef_('url_site') || (CONFIG.URL_SITE.indexOf('__') === 0 ? '' : CONFIG.URL_SITE) || getDef_('url_app') || ''; }

function guardarUrl_() {
  try {
    var u = ScriptApp.getService().getUrl();
    if (u && u !== getDef_('url_app')) setDef_('url_app', u);
  } catch (err) {}
}
function url_(params) {
  var base = siteUrl_();
  var q = Object.keys(params).map(function (k) { return k + '=' + encodeURIComponent(params[k]); }).join('&');
  return base + (q ? '?' + q : '');
}

/* ============ Folha: acesso aos dados ============ */

var _ss = null;
function ss_() {
  if (!_ss) _ss = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('SHEET_ID'));
  return _ss;
}
function folha_(nome) {
  var ss = ss_(), sh = ss.getSheetByName(nome);
  if (!sh) {
    sh = ss.insertSheet(nome); var cab = TABELAS[nome];
    sh.getRange(1, 1, 1, cab.length).setValues([cab]).setFontWeight('bold'); sh.setFrozenRows(1);
    sh.getRange(1, 1, Math.max(sh.getMaxRows(), 2), cab.length).setNumberFormat('@');
  }
  return sh;
}
/* Cache das folhas (V36): cada folha fica guardada em blocos durante até 6 h e é limpa sempre que a app escreve nela.
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
}
function lerFolha_(nome) {
  var sh = folha_(nome);
  var v = sh.getDataRange().getValues();
  var cab = TABELAS[nome];
  var out = [];
  for (var i = 1; i < v.length; i++) {
    if (v[i].join('') === '') continue;
    var o = { _linha: i + 1 };
    cab.forEach(function (h, j) { o[h] = texto_(v[i][j]); });
    out.push(o);
  }
  return out;
}
function texto_(x) {
  if (x instanceof Date) return Utilities.formatDate(x, CONFIG.TZ, 'yyyy-MM-dd');
  return x == null ? '' : String(x);
}
function cabecalho_(sh, cab) {
  if (sh.getLastColumn() < cab.length) sh.getRange(1, 1, 1, cab.length).setValues([cab]).setFontWeight('bold');
}
function acrescentar_(nome, o) {
  var sh = folha_(nome);
  var cab = TABELAS[nome];
  cabecalho_(sh, cab);
  limparCache_(nome);
  var linha = sh.getLastRow() + 1;
  var rg = sh.getRange(linha, 1, 1, cab.length);
  rg.setNumberFormat('@');
  rg.setValues([cab.map(function (h) { return o[h] == null ? '' : String(o[h]); })]);
  SpreadsheetApp.flush(); limparCache_(nome);
}
function atualizar_(nome, o) {
  var sh = folha_(nome);
  var cab = TABELAS[nome];
  cabecalho_(sh, cab);
  limparCache_(nome);
  sh.getRange(o._linha, 1, 1, cab.length).setValues([cab.map(function (h) { return o[h] == null ? '' : String(o[h]); })]);
  SpreadsheetApp.flush(); limparCache_(nome);
}
function apagar_(nome, filtro) {
  var sh = folha_(nome);
  limparCache_(nome);
  lerFolha_(nome).filter(filtro).map(function (o) { return o._linha; }).sort(function (a, b) { return b - a; })
    .forEach(function (l) { sh.deleteRow(l); });
  SpreadsheetApp.flush(); limparCache_(nome);
}
function getDef_(k) {
  var r = ler_('Definicoes').filter(function (o) { return o.chave === k; })[0];
  return r ? r.valor : '';
}
function setDef_(k, v) {
  var r = ler_('Definicoes').filter(function (o) { return o.chave === k; })[0];
  if (r) { r.valor = v; atualizar_('Definicoes', r); } else acrescentar_('Definicoes', { chave: k, valor: v });
}

/* ============ Utilitários ============ */

function hoje_() { return Utilities.formatDate(new Date(), CONFIG.TZ, 'yyyy-MM-dd'); }
function agora_() { return Utilities.formatDate(new Date(), CONFIG.TZ, 'yyyy-MM-dd HH:mm'); }
function novoToken_() { return Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '').slice(0, 8); }
function novoId_() { return Utilities.getUuid().slice(0, 8); }
function limpo_(s, max) { return String(s == null ? '' : s).trim().slice(0, max || 500); }
function emailOk_(e) { return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(e); }
function diasAte_(data) {
  var a = hoje_().split('-'), b = String(data).split('-');
  return Math.round((Date.UTC(+b[0], b[1] - 1, +b[2]) - Date.UTC(+a[0], a[1] - 1, +a[2])) / 864e5);
}
function sim_(b) { return b ? 'SIM' : 'NÃO'; }
function ehSim_(s) { return String(s).toUpperCase() === 'SIM'; }
function jsonOu_(s, d) { try { return s ? JSON.parse(s) : d; } catch (e) { return d; } }
function listaAdmins_() {
  return (getDef_('admins') || '').split(/[\s,;]+/).map(function (x) { return x.toLowerCase().trim(); }).filter(Boolean);
}
function esc_(s) { return String(s == null ? '' : s).replace(/[&<>"']/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]; }); }
var MESES = ['janeiro', 'fevereiro', 'março', 'abril', 'maio', 'junho', 'julho', 'agosto', 'setembro', 'outubro', 'novembro', 'dezembro'];
var DIAS = ['domingo', 'segunda-feira', 'terça-feira', 'quarta-feira', 'quinta-feira', 'sexta-feira', 'sábado'];
var MESES_EN = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
var DIAS_EN = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
function en_(x) { return (x && typeof x === 'object' ? x.lingua : x) === 'en'; }
function tr_(x, pt, en) { return en_(x) ? en : pt; }
function titE_(e, x) { return en_(x) && e.titulo_en ? e.titulo_en : e.titulo; }
function descE_(e, x) { return en_(x) && e.descricao_en ? e.descricao_en : e.descricao; }
function actsE_(e, lista, x) { var pt = jsonOu_(e.atividades, []), en = jsonOu_(e.atividades_en, []); return (lista || []).map(function (a) { var i = pt.indexOf(a); return en_(x) && i >= 0 && en[i] ? en[i] : a; }); }
function dataLonga_(s, x) {
  var p = String(s).split('-'); var d = new Date(Date.UTC(+p[0], p[1] - 1, +p[2]));
  if (en_(x)) return DIAS_EN[d.getUTCDay()] + ', ' + (+p[2]) + ' ' + MESES_EN[p[1] - 1] + ' ' + p[0];
  return DIAS[d.getUTCDay()] + ', ' + (+p[2]) + ' de ' + MESES[p[1] - 1] + ' de ' + p[0];
}
function saud_(nome, x) { if (en_(x)) return nome ? 'Dear ' + String(nome).split(' ')[0] + ',' : 'Dear Brother/Sister,'; return nome ? 'Caro(a) ' + String(nome).split(' ')[0] + ',' : 'Caro(a) Irmão(ã),'; }
function lock_(fn) {
  var l = LockService.getScriptLock();
  l.waitLock(20000);
  try { return fn(); } finally { l.releaseLock(); }
}

/* ============ Emails ============ */

function htmlEmail_(o) {
  var org = esc_(getDef_('nome_org') || CONFIG.NOME_ORG);
  var corpo = (o.linhas || []).filter(Boolean).map(function (l) {
    if (l.indexOf('**') === 0) return '<p style="margin:0 0 10px;font-size:17px;font-weight:bold;color:#1d1a1b">' + esc_(l.replace(/\*\*/g, '')) + '</p>';
    return '<p style="margin:0 0 10px">' + esc_(l) + '</p>';
  }).join('');
  var L = o.lang;
  var cal = o.calendario ? '<div style="margin:16px 0 4px;padding:12px 14px;background:#faf8f8;border:1px solid #e3dcdd;border-radius:10px;font-size:14px"><b>📅 ' + tr_(L, 'Adicionar ao calendário', 'Add to calendar') + '</b><br>' + tr_(L, 'Abra o ficheiro <b>convite.ics</b> em anexo para juntar este evento ao calendário do telemóvel. Se o evento mudar, receberá um convite atualizado.', 'Open the attached <b>convite.ics</b> file to add this event to your phone calendar. If the event changes, you will receive an updated invitation.') + '<br><a href="' + esc_(o.calendario) + '" style="color:#b0001e;font-weight:bold">' + tr_(L, 'Subscrever o meu calendário da Comenda', 'Subscribe to my Comenda calendar') + '</a> ' + tr_(L, '(atualiza-se sozinho, iPhone e Android)', '(updates automatically, iPhone and Android)') + '</div>' : '';
  var botao = o.botao ? '<p style="margin:18px 0 6px"><a href="' + esc_(o.botao.url) + '" style="background:#b0001e;color:#ffffff;text-decoration:none;padding:11px 18px;border-radius:8px;font-weight:bold;display:inline-block">' + esc_(o.botao.texto) + '</a></p>' : '';
  if (o.recusar || o.talvez) botao += '<p style="margin:4px 0 6px;font-size:14px">' + [o.talvez, o.recusar].filter(Boolean).map(function (x) { return '<a href="' + esc_(x.url) + '" style="color:#6a6264">' + esc_(x.texto) + '</a>'; }).join(' &nbsp;·&nbsp; ') + '</p>';
  var rodape = o.sair ? (o.porAdmin ? tr_(L, 'Recebe este email porque a ' + org + ' o/a incluiu nas suas notificações. <a href="' + esc_(o.sair) + '" style="color:#b0001e">Deixar de receber</a>.', 'You receive this email because ' + org + ' added you to its notifications. <a href="' + esc_(o.sair) + '" style="color:#b0001e">Unsubscribe</a>.') : tr_(L, 'Recebe este email porque se inscreveu nas notificações da ' + org + '. <a href="' + esc_(o.sair) + '" style="color:#b0001e">Deixar de receber</a>.', 'You receive this email because you subscribed to notifications from ' + org + '. <a href="' + esc_(o.sair) + '" style="color:#b0001e">Unsubscribe</a>.')) : org + ' · OSMTH';
  return '<div style="background:#f5f4f4;padding:20px 10px;font-family:Arial,Helvetica,sans-serif;color:#2b2627;font-size:15px;line-height:1.5">' +
    '<div style="max-width:560px;margin:0 auto;background:#ffffff;border:1px solid #e3dcdd;border-radius:12px;overflow:hidden">' +
    '<div style="background:#b0001e;color:#ffffff;padding:14px 18px;font-family:Georgia,serif;font-size:18px"><img src="cid:selo" width="36" height="36" alt="" style="vertical-align:middle;border-radius:50%;background:#fff;margin-right:10px">' + org + '</div>' +
    '<div style="padding:20px 18px">' + corpo + (o.extra || '') + botao + cal + '</div>' +
    '<div style="padding:12px 18px;border-top:1px solid #e3dcdd;font-size:12px;color:#6a6264">' + rodape + '</div></div></div>';
}
/** prioridade: 1 confirmações/entrada · 2 inscrições/avisos admin · 3 lembretes/alterações · 4 convites/novidades */
function enfileirar_(para, assunto, o, prioridade) {
  acrescentar_('Fila', { id: novoId_(), para: para, assunto: assunto, html: htmlEmail_(o), prioridade: prioridade || 4, criado_em: new Date().toISOString(), enviado_em: '', erro: '', ics: o.ics || '' });
  pushPendente_(para, assunto, o, prioridade);
}
function avisarAdmins_(assunto, linhas) {
  var url = url_({ v: 'admin' });
  listaAdmins_().forEach(function (a) { enfileirar_(a, assunto, { linhas: linhas, botao: { texto: 'Abrir painel', url: url } }, 2); });
}
function processarFila(max) {
  if (typeof max !== 'number') { try { notifAgendadas_(); } catch (err) {} try { ['Definicoes', 'Eventos', 'Inscritos', 'Inscricoes', 'Convidados', 'Recusas', 'Membros', 'Dispositivos'].forEach(ler_); } catch (err) {} }
  try { enviarPushPendentes_(); } catch (e) {}
  var lk = LockService.getScriptLock();
  if (!lk.tryLock(5000)) return 0;
  var enviados = 0;
  try {
    var inicio = Date.now();
    var limite = typeof max === 'number' ? max : 1e9;
    var quota = MailApp.getRemainingDailyQuota();
    var pend = ler_('Fila').filter(function (m) { return !m.enviado_em; })
      .sort(function (a, b) { return (a.prioridade - b.prioridade) || (a.criado_em < b.criado_em ? -1 : 1); });
    var selo = Utilities.newBlob(Utilities.base64Decode(SELO_PNG), 'image/jpeg', 'selo.jpg');
    var nome = getDef_('nome_org') || CONFIG.NOME_ORG;
    for (var i = 0; i < pend.length && quota > 0 && enviados < limite && Date.now() - inicio < 270000; i++) {
      var m = pend[i];
      try {
        var msg = { to: m.para, subject: m.assunto, htmlBody: m.html, name: nome, inlineImages: { selo: selo } };
        if (m.ics) { var met = /METHOD:CANCEL/.test(m.ics) ? 'CANCEL' : 'REQUEST'; msg.attachments = [Utilities.newBlob(m.ics, 'text/calendar; charset=UTF-8; method=' + met, met === 'CANCEL' ? 'cancelamento.ics' : 'convite.ics')]; }
        MailApp.sendEmail(msg);
        m.enviado_em = agora_(); m.erro = '';
        quota--; enviados++;
      } catch (err) {
        m.erro = String(err).slice(0, 200);
        if (/quota|limit|too many/i.test(m.erro)) { atualizar_('Fila', m); break; }
        m.enviado_em = 'ERRO';
      }
      atualizar_('Fila', m);
    }
  } finally { lk.releaseLock(); }
  return enviados;
}
function enviarJa_() { try { enviarPushPendentes_(); } catch (e) {} try { processarFila(CONFIG.ENVIO_IMEDIATO_MAX); } catch (e) {} }

function tarefaDiaria() {
  lock_(function () {
    var subs = ler_('Inscritos');
    ler_('Eventos').forEach(function (e) {
      if (e.estado !== 'ativo' || e.externo) return;
      var d = diasAte_(e.data);
      try { if (marcarVideo_(e, d)) atualizar_('Eventos', e); } catch (err) {}
      if (d < 0 && d >= -3 && !e.agradecimento) {
        agradecimentos_(e);
        e.agradecimento = agora_();
        atualizar_('Eventos', e);
        return;
      }
      if ((d === 7 || d === 1) && !e['lembrete' + d]) {
        lembretes_(e, d, subs);
        e['lembrete' + d] = agora_();
        atualizar_('Eventos', e);
      }
    });
    try { copiarFotos_(); } catch (err) {}
    try { lembreteSemanal_(); } catch (err) {}
    // limpar emails enviados há mais de 60 dias
    var lim = new Date(Date.now() - 60 * 864e5).toISOString();
    apagar_('Fila', function (m) { return m.enviado_em && m.criado_em < lim; });
    // limpar sessões de administrador expiradas
    var props = PropertiesService.getScriptProperties(), all = props.getProperties();
    Object.keys(all).forEach(function (k) {
      if (k.indexOf('adm_') === 0 || k.indexOf('rst_') === 0) { var s = jsonOu_(all[k], {}); if (!s.exp || s.exp < Date.now()) props.deleteProperty(k); }
    });
  });
  processarFila();
}

/* ============ Regras de negócio ============ */

/* ---- visibilidade ---- */
var _dir = null;
function direcaoSet_() {
  if (_dir) return _dir;
  _dir = { email: {}, num: {} };
  ler_('Membros').forEach(function (m) { if (ehSim_(m.direcao)) { if (m.email) _dir.email[m.email.toLowerCase()] = 1; if (m.n_ordem) _dir.num[normNum_(m.n_ordem)] = 1; } });
  listaAdmins_().forEach(function (a) { _dir.email[a] = 1; });
  return _dir;
}
function ehDirecao_(s) { if (!s) return false; var d = direcaoSet_(); return !!(d.email[s.email] || (s.n_ordem && d.num[normNum_(s.n_ordem)])); }
function podeVer_(e, s) {
  if (e.visibilidade === 'publico') return true;
  if (!s) return false;
  if (e.visibilidade === 'direcao') return ehDirecao_(s);
  return ehSim_(s.membro) || ehDirecao_(s);
}

/* ---- calendário (iCalendar) ---- */
function icsEsc_(t) { return String(t == null ? '' : t).replace(/\\/g, '\\\\').replace(/;/g, '\\;').replace(/,/g, '\\,').replace(/\r?\n/g, '\\n'); }
function icsFold_(l) { var out = [], i = 0; while (l.length - i > 72) { out.push((i ? ' ' : '') + l.slice(i, i + 72)); i += 72; } out.push((i ? ' ' : '') + l.slice(i)); return out.join('\r\n'); }
function icsUtc_(dt) { return Utilities.formatDate(dt, 'UTC', "yyyyMMdd'T'HHmmss'Z'"); }
function inicioFim_(e) {
  var ini = Utilities.parseDate(e.data + ' ' + (e.hora || '00:00'), CONFIG.TZ, 'yyyy-MM-dd HH:mm');
  var fim = e.hora_fim && e.hora_fim > (e.hora || '') ? Utilities.parseDate(e.data + ' ' + e.hora_fim, CONFIG.TZ, 'yyyy-MM-dd HH:mm') : new Date(ini.getTime() + CONFIG.DURACAO_PADRAO_HORAS * 3600000);
  return [ini, fim];
}
function vevento_(e, extra, cancelado, x) {
  var t = inicioFim_(e), org = getDef_('nome_org') || CONFIG.NOME_ORG;
  var L = ['BEGIN:VEVENT', 'UID:' + e.id + '@comendabraga-osmth', 'SEQUENCE:' + (+e.seq || 0), 'DTSTAMP:' + icsUtc_(new Date()),
    'DTSTART:' + icsUtc_(t[0]), 'DTEND:' + icsUtc_(t[1]), 'SUMMARY:' + icsEsc_(titE_(e, x)), 'LOCATION:' + icsEsc_(e.local + (e.morada ? ', ' + e.morada : '')),
    'DESCRIPTION:' + icsEsc_((descE_(e, x) ? descE_(e, x) + '\n\n' : '') + org + ' · ' + url_({ ev: e.id })), 'URL:' + url_({ ev: e.id }),
    'STATUS:' + (cancelado ? 'CANCELLED' : 'CONFIRMED'), 'TRANSP:OPAQUE'].concat(extra || []);
  if (!cancelado) L.push('BEGIN:VALARM', 'ACTION:DISPLAY', 'DESCRIPTION:' + icsEsc_(titE_(e, x)), 'TRIGGER:-P1D', 'END:VALARM');
  L.push('END:VEVENT');
  return L;
}
function calendario_(metodo, eventos, nomeCal) {
  var L = ['BEGIN:VCALENDAR', 'VERSION:2.0', 'PRODID:-//Comenda de Braga OSMTH//Calendario//PT', 'CALSCALE:GREGORIAN', 'METHOD:' + metodo];
  if (nomeCal) L.push('X-WR-CALNAME:' + icsEsc_(nomeCal), 'X-WR-TIMEZONE:' + CONFIG.TZ, 'REFRESH-INTERVAL;VALUE=DURATION:PT1H', 'X-PUBLISHED-TTL:PT1H');
  eventos.forEach(function (x) { L = L.concat(x); });
  L.push('END:VCALENDAR');
  return L.map(icsFold_).join('\r\n') + '\r\n';
}
// Convite individual (REQUEST) ou cancelamento (CANCEL) para um inscrito
function conviteIcs_(e, s, cancelar) {
  var org = getDef_('nome_org') || CONFIG.NOME_ORG;
  var extra = ['ORGANIZER;CN=' + icsEsc_(org) + ':mailto:' + CONFIG.EMAIL_ORG, 'ATTENDEE;CN=' + icsEsc_(s.nome || s.email) + ';ROLE=REQ-PARTICIPANT;PARTSTAT=ACCEPTED;RSVP=FALSE:mailto:' + s.email];
  return calendario_(cancelar ? 'CANCEL' : 'REQUEST', [vevento_(e, extra, cancelar, s)]);
}
// Calendário subscrito: eventos ativos em que a pessoa está inscrita
function feedIcal_(token) {
  var s = ler_('Inscritos').filter(function (x) { return x.token === token && x.estado === 'confirmado'; })[0];
  var org = getDef_('nome_org') || CONFIG.NOME_ORG;
  if (!s) return calendario_('PUBLISH', [], org);
  marcarCal_(s);
  var ids = {}; ler_('Inscricoes').forEach(function (r) { if (r.email === s.email) ids[r.evento_id] = 1; });
  var evs = ler_('Eventos').filter(function (e) { return ids[e.id] && e.estado === 'ativo' && podeVer_(e, s); }).map(function (e) { return vevento_(e, [], false, s); });
  return calendario_('PUBLISH', evs, org);
}
// O calendário do telemóvel vem buscar o feed periodicamente: guardamos a data (no máx. 1 escrita por dia e pessoa).
var CAL_DIAS = 14;
function calLigado_(s) { return !!(s && s.cal_ligado && diasAte_(s.cal_ligado) >= -CAL_DIAS); }
function marcarCal_(s) {
  var hoje = hoje_(); if (s.cal_ligado === hoje) return;
  try {
    var l = LockService.getScriptLock(); if (!l.tryLock(5000)) return;
    try {
      var sh = folha_('Inscritos'), cab = TABELAS.Inscritos; cabecalho_(sh, cab);
      var col = cab.indexOf('cal_ligado') + 1, tk = cab.indexOf('token') + 1;
      if (String(sh.getRange(s._linha, tk).getValue()) !== s.token) return;
      sh.getRange(s._linha, col).setNumberFormat('@').setValue(hoje); SpreadsheetApp.flush(); limparCache_('Inscritos');
    } finally { l.releaseLock(); }
  } catch (err) {}
}
function urlCal_(s) { return url_({ t: s.token, cal: 1 }); }
function urlFeed_(s) { return (CONFIG.URL_API || getDef_('url_app')) + '?ical=' + encodeURIComponent(s.token); }
// Envia convite atualizado (ou cancelamento) a todos os inscritos no evento
function atualizarCalendarios_(e, cancelar, jaAvisados) {
  var subs = ler_('Inscritos'), n = 0;
  ler_('Inscricoes').filter(function (r) { return r.evento_id === e.id; }).forEach(function (r) {
    var s = subs.filter(function (x) { return x.email === r.email && x.estado === 'confirmado'; })[0];
    if (!s || (jaAvisados && jaAvisados[s.email])) return;
    var lig = calLigado_(s);
    enfileirar_(s.email, (cancelar ? tr_(s, 'Cancelado: ', 'Cancelled: ') : tr_(s, 'Atualização do calendário: ', 'Calendar update: ')) + titE_(e, s), {
      lang: s.lingua,
      linhas: [saud_(s.nome, s), cancelar ? tr_(s, 'O seguinte evento foi cancelado e vai ser retirado do seu calendário:', 'The following event has been cancelled and will be removed from your calendar:') : lig ? tr_(s, 'O seguinte evento, em que está inscrito, foi alterado:', 'The following event you registered for has changed:') : tr_(s, 'O seguinte evento, em que está inscrito, foi alterado. O convite em anexo atualiza o seu calendário:', 'The following event you registered for has changed. The attached invitation updates your calendar:'), '**' + titE_(e, s), resumo_(e, s), lig ? tr_(s, 'O seu calendário da Comenda atualiza-se sozinho.', 'Your Comenda calendar updates automatically.') : ''],
      botao: cancelar ? null : { texto: tr_(s, 'Ver o evento', 'View the event'), url: url_({ t: s.token, ev: e.id }) }, sair: url_({ sair: s.token }), ics: lig ? '' : conviteIcs_(e, s, cancelar)
    }, 3);
    n++;
  });
  return n;
}

function evPublico_(e) {
  return { id: e.id, titulo: e.titulo, titulo_en: e.titulo_en, descricao_en: e.descricao_en, atividades_en: jsonOu_(e.atividades_en, []), data: e.data, hora: e.hora, hora_fim: e.hora_fim, local: e.local, descricao: e.descricao, visibilidade: e.visibilidade, prazo: e.prazo, lotacao: e.lotacao, campos: jsonOu_(e.campos, {}), atividades: jsonOu_(e.atividades, []), estado: e.estado, externo: e.externo || '', externo_link: e.externo_link || '', morada: e.morada || '', mapa: e.mapa || '', programa: e.programa || '', programa_en: e.programa_en || '', traje: e.traje || '', traje_en: e.traje_en || '', refeicao: refPublica_(e) };
}
function destinatarios_(e, subs) {
  return (subs || ler_('Inscritos')).filter(function (s) { return s.estado === 'confirmado' && podeVer_(e, s); });
}
function resumo_(e, x) { return dataLonga_(e.data, x) + tr_(x, ' às ', ' at ') + e.hora + ' · ' + e.local; }
function quando_(e, x) { return dataLonga_(e.data, x) + tr_(x, ' às ', ' at ') + e.hora + (e.hora_fim ? ' – ' + e.hora_fim : ''); }
/* ---- Locais (com link para o mapa), ordem dos trabalhos, almoço/jantar, pagamento e traje ---- */
function urlMapa_(local, morada, link) {
  if (/^https?:\/\/\S+$/.test(link || '')) return link;
  var q = [local, morada].filter(Boolean).join(', ');
  return q ? 'https://www.google.com/maps/search/?api=1&query=' + encodeURIComponent(q) : '';
}
function refeicao_(e) { var r = jsonOu_(e.refeicao, null); return r && r.tipo ? r : null; }
function refPublica_(e) { var r = refeicao_(e); return r ? { tipo: r.tipo, nome: r.nome || '', hora: r.hora || '', local: r.local || '', morada: r.morada || '', mapa: r.mapa || '', preco: r.preco || '', preco_ate: r.preco_ate || '', preco2: r.preco2 || '' } : null; }
function refNome_(r, x) { return !r ? '' : r.nome || (r.tipo === 'almoco' ? tr_(x, 'Almoço', 'Lunch') : tr_(x, 'Jantar', 'Dinner')); }
function num_(v) { var n = parseFloat(String(v == null ? '' : v).replace(',', '.')); return isFinite(n) && n > 0 ? Math.round(n * 100) / 100 : 0; }
function euro_(n) { return n.toFixed(2).replace('.', ',') + ' €'; }
function precoHoje_(r) { var p2 = num_(r.preco2); return r.preco_ate && p2 && hoje_() > r.preco_ate ? p2 : num_(r.preco); }
function precoTexto_(r, x) {
  var p1 = num_(r.preco), p2 = num_(r.preco2);
  if (!p1) return '';
  return euro_(p1) + tr_(x, ' por pessoa', ' per person') + (r.preco_ate && p2 ? tr_(x, ' até ', ' until ') + dataLonga_(r.preco_ate, x) + tr_(x, '; depois, ', '; after that, ') + euro_(p2) : '');
}
function lk_(url, txt) { return '<a href="' + esc_(url) + '" style="color:#b0001e;font-weight:bold">' + esc_(txt) + '</a>'; }
function caixa_(titulo, corpo, destaque) {
  return '<div style="margin:12px 0;padding:12px 14px;background:' + (destaque ? '#fdf2f3' : '#faf8f8') + ';border:1px solid ' + (destaque ? '#b0001e' : '#e3dcdd') + ';border-radius:10px;font-size:14px"><b>' + titulo + '</b><br>' + corpo + '</div>';
}
function linhasTxt_(t) { return String(t || '').split('\n').map(function (l) { return l.trim(); }).filter(Boolean); }
function detalhes_(e, x, pag) {
  var h = '', m = urlMapa_(e.local, e.morada, e.mapa), abrir = tr_(x, 'Abrir no mapa', 'Open in maps');
  h += caixa_('📍 ' + tr_(x, 'Local', 'Venue'), lk_(m, e.local) + (e.morada ? '<br>' + esc_(e.morada) : '') + '<br>' + lk_(m, abrir + ' ↗'));
  var prog = linhasTxt_(en_(x) && e.programa_en ? e.programa_en : e.programa);
  if (prog.length) h += caixa_('📜 ' + tr_(x, 'Ordem dos trabalhos', 'Agenda'), '<ul style="margin:4px 0 0;padding-left:18px">' + prog.map(function (l) { return '<li>' + esc_(l) + '</li>'; }).join('') + '</ul>');
  var r = refeicao_(e);
  if (r) {
    var mr = urlMapa_(r.local, r.morada, r.mapa), pt = precoTexto_(r, x);
    h += caixa_('🍽️ ' + esc_(refNome_(r, x)) + (r.hora ? tr_(x, ' · às ', ' · at ') + esc_(r.hora) : ''), (mr ? lk_(mr, r.local) : esc_(r.local)) + (r.morada ? '<br>' + esc_(r.morada) : '') + (mr ? '<br>' + lk_(mr, abrir + ' ↗') : '') + (pt ? '<br>' + tr_(x, 'Valor: ', 'Price: ') + esc_(pt) : ''));
  }
  var traje = en_(x) && e.traje_en ? e.traje_en : e.traje;
  if (traje) h += caixa_('👔 ' + tr_(x, 'Traje obrigatório', 'Mandatory dress code'), esc_(traje).replace(/\n/g, '<br>'), true);
  return h + (pag || '');
}
// Bloco de pagamento: só para quem escolheu a refeição (n = nº de pessoas com refeição)
function pagamento_(e, x, n) {
  var r = refeicao_(e); if (!r || !n) return '';
  var p = precoHoje_(r), L = [], p2 = num_(r.preco2);
  if (p) L.push(tr_(x, 'Total a pagar: ', 'Amount due: ') + '<b>' + n + ' × ' + euro_(p) + ' = ' + euro_(n * p) + '</b>' + (r.preco_ate && p2 && hoje_() <= r.preco_ate ? '<br><span style="color:#6a6264">' + esc_(tr_(x, 'Valor para pagamentos até ', 'Price for payments made by ') + dataLonga_(r.preco_ate, x) + tr_(x, '. Depois dessa data: ', '. After that date: ') + euro_(p2) + tr_(x, ' por pessoa.', ' per person.')) + '</span>' : ''));
  if (r.iban) L.push('IBAN: <b>' + esc_(r.iban) + '</b>');
  if (r.mbway) L.push('MB Way: <b>' + esc_(r.mbway) + '</b>');
  var em = String(r.emails || '').split(/[,;\s]+/).filter(emailOk_);
  if (em.length) L.push(tr_(x, 'Envie o comprovativo de pagamento para: ', 'Please send the proof of payment to: ') + em.map(function (a) { return lk_('mailto:' + a + '?subject=' + encodeURIComponent(tr_(x, 'Comprovativo de pagamento: ', 'Proof of payment: ') + titE_(e, x)), a); }).join(', '));
  if (r.nota) L.push(esc_(r.nota).replace(/\n/g, '<br>'));
  if (!L.length) return '';
  return caixa_('💳 ' + tr_(x, 'Pagamento', 'Payment') + ' · ' + esc_(refNome_(r, x)) + ' (' + n + ' ' + (n === 1 ? tr_(x, 'pessoa', 'person') : tr_(x, 'pessoas', 'people')) + ')', L.join('<br>'), true);
}
function convites_(e, subs) {
  if (e.externo) return informativos_(e, subs);
  destinatarios_(e, subs).forEach(function (s) {
    var org = getDef_('nome_org') || CONFIG.NOME_ORG;
    enfileirar_(s.email, tr_(s, 'Novo evento: ', 'New event: ') + titE_(e, s), {
      lang: s.lingua,
      linhas: [saud_(s.nome, s), tr_(s, 'Foi marcado um novo evento da ' + org + ':', 'A new event has been scheduled by ' + org + ':'), '**' + titE_(e, s), quando_(e, s), descE_(e, s), tr_(s, 'Inscrições até ', 'Registration closes on ') + dataLonga_(e.prazo, s) + '.'],
      extra: detalhes_(e, s),
      botao: { texto: tr_(s, 'Inscrever-me neste evento', 'Register for this event'), url: url_({ t: s.token, ev: e.id }) }, recusar: linkRecusar_(e, s), talvez: linkTalvez_(e, s), sair: url_({ sair: s.token })
    }, 4);
  });
}
// Evento de outra comenda: um único email informativo, sem inscrição
function informativos_(e, subs) {
  var org = getDef_('nome_org') || CONFIG.NOME_ORG, link = e.externo_link || '', url = /^https?:\/\/\S+$/.test(link);
  destinatarios_(e, subs).forEach(function (s) {
    enfileirar_(s.email, tr_(s, 'Para sua informação: ', 'For your information: ') + titE_(e, s) + ' (' + e.externo + ')', {
      lang: s.lingua,
      linhas: [saud_(s.nome, s), tr_(s, 'A ' + org + ' partilha consigo o seguinte evento, organizado pela ' + e.externo + ':', org + ' would like to share the following event, organised by ' + e.externo + ':'), '**' + titE_(e, s), quando_(e, s), descE_(e, s),
        tr_(s, 'Este email é apenas informativo: a inscrição não é feita através da ' + org + '.', 'This email is for information only: registration is not handled by ' + org + '.'),
        link && !url ? tr_(s, 'Mais informações: ', 'More information: ') + link : ''],
      extra: detalhes_(e, s),
      botao: url ? { texto: tr_(s, 'Mais informações', 'More information'), url: link } : { texto: tr_(s, 'Ver no calendário', 'View in the calendar'), url: url_({ t: s.token, ev: e.id }) }, sair: url_({ sair: s.token })
    }, 4);
  });
}
function lembretes_(e, d, subs) {
  var regs = ler_('Inscricoes').filter(function (r) { return r.evento_id === e.id; }), rec = recusasEvento_(e.id);
  destinatarios_(e, subs).forEach(function (s) {
    var r = regs.filter(function (x) { return x.email === s.email; })[0];
    var tv = !r && ehTalvez_(rec[s.email]);
    if (!r && rec[s.email] && !tv) return;
    enfileirar_(s.email, (d === 7 ? tr_(s, 'Faltam 7 dias: ', '7 days to go: ') : tr_(s, 'É amanhã: ', 'Tomorrow: ')) + titE_(e, s), {
      lang: s.lingua,
      linhas: [saud_(s.nome, s), d === 7 ? tr_(s, 'Lembramos que o evento é daqui a uma semana.', 'A reminder that the event is one week away.') : tr_(s, 'Lembramos que o evento é amanhã.', 'A reminder that the event is tomorrow.'), '**' + titE_(e, s), quando_(e, s),
        r ? tr_(s, 'Já está inscrito' + (+r.n_convidados ? ' com ' + r.n_convidados + ' convidado(s)' : '') + '.', 'You are registered' + (+r.n_convidados ? ' with ' + r.n_convidados + ' guest(s)' : '') + '.') : (hoje_() <= e.prazo ? (tv ? tr_(s, 'Indicou que talvez vá. Se já decidiu, inscreva-se ou diga-nos que não vai. Inscrições até ', 'You said you might attend. If you have decided, please register or let us know you will not attend. Registration closes on ') : tr_(s, 'Ainda não se inscreveu. Inscrições até ', 'You have not registered yet. Registration closes on ')) + dataLonga_(e.prazo, s) + '.' : '')],
      extra: detalhes_(e, s),
      botao: { texto: r ? tr_(s, 'Ver ou alterar inscrição', 'View or change registration') : tr_(s, 'Inscrever-me', 'Register'), url: url_({ t: s.token, ev: e.id }) }, recusar: !r && hoje_() <= e.prazo ? linkRecusar_(e, s) : null, talvez: !r && !tv && hoje_() <= e.prazo ? linkTalvez_(e, s) : null, sair: url_({ sair: s.token })
    }, 3);
  });
}
// Dia seguinte ao evento: agradecimento + pedido de fotos para o álbum da Comenda
/* ---- Pastas de fotos no Google Drive (automáticas) ---- */
function pastaRaiz_(chave, nome) {
  var props = PropertiesService.getScriptProperties(), id = props.getProperty(chave);
  if (id) { try { var f = DriveApp.getFolderById(id); if (!f.isTrashed()) return f; } catch (err) {} }
  var it = DriveApp.getFoldersByName(nome), f2 = it.hasNext() ? it.next() : DriveApp.createFolder(nome);
  props.setProperty(chave, f2.getId());
  return f2;
}
function nomePasta_(e) { return e.data + ' – ' + String(e.titulo || '').replace(/[\\/:*?"<>|]/g, ' ').replace(/\s+/g, ' ').trim(); }
// Cria (uma vez) a pasta partilhada do evento: qualquer pessoa com o link pode adicionar fotos
function garantirPasta_(e) {
  if (e.pasta_id) { try { var f = DriveApp.getFolderById(e.pasta_id); if (!f.isTrashed()) return e.pasta_url || f.getUrl(); } catch (err) {} }
  var org = getDef_('nome_org') || CONFIG.NOME_ORG;
  var raiz = pastaRaiz_('pasta_fotos', 'Fotos dos eventos – ' + org);
  var p = raiz.createFolder(nomePasta_(e));
  p.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.EDIT);
  p.setDescription('Fotos de «' + e.titulo + '» (' + e.data + '). Partilhe aqui as suas fotos.');
  e.pasta_id = p.getId(); e.pasta_url = p.getUrl();
  return e.pasta_url;
}
// Copia as fotos novas das pastas partilhadas para uma pasta privada (cópia de segurança)
function copiarFotos_() {
  var org = getDef_('nome_org') || CONFIG.NOME_ORG, privada = null, n = 0, lim = Date.now() + 4 * 60 * 1000;
  ler_('Eventos').forEach(function (e) {
    if (!e.pasta_id || diasAte_(e.data) < -120 || Date.now() > lim) return;
    var src; try { src = DriveApp.getFolderById(e.pasta_id); } catch (err) { return; }
    var files = src.getFiles(); if (!files.hasNext()) return;
    if (!privada) privada = pastaRaiz_('pasta_fotos_copia', 'Fotos dos eventos (cópia privada) – ' + org);
    var nome = nomePasta_(e), itd = privada.getFoldersByName(nome), dst = itd.hasNext() ? itd.next() : privada.createFolder(nome);
    var ja = {}, it = dst.getFiles(); while (it.hasNext()) { var d = it.next().getDescription(); if (d) ja[d] = 1; }
    while (files.hasNext() && Date.now() < lim) {
      var f = files.next(), tag = 'orig:' + f.getId();
      if (ja[tag]) continue;
      try { var c = f.makeCopy(f.getName(), dst); c.setDescription(tag); n++; } catch (err) {}
    }
  });
  return n;
}
// Correr uma vez no editor (Executar) para autorizar o acesso ao Google Drive
function autorizarDrive() {
  var org = getDef_('nome_org') || CONFIG.NOME_ORG;
  var a = pastaRaiz_('pasta_fotos', 'Fotos dos eventos – ' + org), b = pastaRaiz_('pasta_fotos_copia', 'Fotos dos eventos (cópia privada) – ' + org);
  Logger.log('Google Drive autorizado. Pastas: ' + a.getUrl() + ' e ' + b.getUrl());
}
function agradecimentos_(e) {
  var subs = ler_('Inscritos'), enviados = {}, n = 0, org = getDef_('nome_org') || CONFIG.NOME_ORG;
  var fotos = e.album || '', drive = false;
  if (!fotos) { try { fotos = garantirPasta_(e); drive = !!fotos; } catch (err) { fotos = ''; avisarAdmins_('Pasta de fotos não criada: ' + e.titulo, ['Não foi possível criar a pasta de fotos no Google Drive para este evento (' + err.message + ').', 'O agradecimento foi enviado sem o pedido de fotos. Confirme que a autorização do Google Drive foi dada no editor do Apps Script.']); } }
  var corpo = function (nome, x) {
    var L = [saud_(nome, x), tr_(x, 'Obrigado por ter estado connosco em «' + titE_(e, x) + '» (' + dataLonga_(e.data, x) + '). A sua presença tornou este encontro ainda mais especial.', 'Thank you for joining us at «' + titE_(e, x) + '» (' + dataLonga_(e.data, x) + '). Your presence made this gathering even more special.')];
    if (fotos && !drive) L.push(tr_(x, 'Se tirou fotografias, pedimos-lhe que as partilhe no álbum da ' + org + ' no Google Fotos, para que todos possam recordar este momento.', 'If you took any photos, please share them in the ' + org + ' album on Google Photos, so that everyone can remember this moment.'), tr_(x, 'Basta carregar no botão abaixo, entrar com a sua conta Google e escolher «Adicionar fotos».', 'Just click the button below, sign in with your Google account and choose «Add photos».'));
    if (fotos && drive) L.push(tr_(x, 'Se tirou fotografias, pedimos-lhe que as partilhe na pasta de fotos deste evento, no Google Drive da ' + org + ', para que todos possam recordar este momento.', 'If you took any photos, please share them in this event\'s photo folder on the ' + org + ' Google Drive, so that everyone can remember this moment.'), tr_(x, 'Carregue no botão abaixo, entre com a sua conta Google e arraste as fotos para a pasta (no telemóvel: botão «+» → «Carregar»).', 'Click the button below, sign in with your Google account and drag your photos into the folder (on a phone: «+» button → «Upload»).'));
    L.push(tr_(x, 'Até breve.', 'See you soon.'));
    return L;
  };
  var btn = function (x) { return fotos ? (drive ? { texto: '📁 ' + tr_(x, 'Enviar as minhas fotos', 'Upload my photos'), url: fotos } : { texto: '📷 ' + tr_(x, 'Partilhar as minhas fotos', 'Share my photos'), url: fotos }) : null; };
  var pres = presentesEmail_(e);
  if (pres) {
    pres.forEach(function (p) {
      n++;
      enfileirar_(p.email, tr_(p.s, 'Obrigado pela sua presença: ', 'Thank you for coming: ') + titE_(e, p.s), { lang: p.s && p.s.lingua, linhas: corpo(p.nome, p.s).concat(p.de && !p.s ? [tr_(p.s, '(Recebe este email porque esteve presente como convidado/a de ' + p.de + '.)', '(You receive this email because you attended as a guest of ' + p.de + '.)')] : []), botao: btn(p.s), sair: p.s ? url_({ sair: p.s.token }) : '' }, 4);
    });
    return n;
  }
  var porEmail = {};
  ler_('Inscricoes').filter(function (r) { return r.evento_id === e.id; }).forEach(function (r) {
    var s = subs.filter(function (x) { return x.email === r.email && x.estado === 'confirmado'; })[0];
    if (!s) return;
    porEmail[r.id] = s;
    if (enviados[s.email]) return;
    enviados[s.email] = 1; n++;
    enfileirar_(s.email, tr_(s, 'Obrigado pela sua presença: ', 'Thank you for coming: ') + titE_(e, s), { lang: s.lingua, linhas: corpo(r.nome || s.nome, s), botao: btn(s), sair: url_({ sair: s.token }) }, 4);
  });
  ler_('Convidados').filter(function (g) { return g.evento_id === e.id && g.email; }).forEach(function (g) {
    var em = String(g.email).toLowerCase(), x = porEmail[g.inscricao_id] || null;
    if (!emailOk_(em) || enviados[em]) return;
    enviados[em] = 1; n++;
    enfileirar_(em, tr_(x, 'Obrigado pela sua presença: ', 'Thank you for coming: ') + titE_(e, x), { lang: x && x.lingua, linhas: corpo(g.nome, x).concat([tr_(x, '(Recebe este email porque ' + g.convidado_de + ' o/a inscreveu como convidado/a.)', '(You receive this email because ' + g.convidado_de + ' registered you as a guest.)')]), botao: btn(x) }, 4);
  });
  return n;
}
/* ---- Vídeo automático do evento (montado no GitHub a partir das fotos da pasta do Drive) ---- */
var VIDEO_DIAS = 4, VIDEO_JANELA = 14, VIDEO_MIN_FOTOS = 5;
function infoVideo_(e) { return jsonOu_(e.video_info, {}); }
function fotosPasta_(e) {
  if (!e.pasta_id) return [];
  var it = DriveApp.getFolderById(e.pasta_id).getFiles(), out = [];
  while (it.hasNext()) {
    var f = it.next(), mime = f.getMimeType() || '';
    if (mime.indexOf('image/') === 0 && f.getSize() < 45e6) out.push({ id: f.getId(), nome: f.getName(), mime: mime, tamanho: f.getSize(), criado: f.getDateCreated().toISOString() });
  }
  return out;
}
// Chamado pela tarefa diária: ao 4.º dia após o evento, marca o vídeo como pendente (se houver fotos)
function marcarVideo_(e, d) {
  if (e.video_estado || e.estado !== 'ativo' || e.externo || d > -VIDEO_DIAS || d < -VIDEO_JANELA) return false;
  if (!e.pasta_id) {
    e.video_estado = 'sem_pasta';
    avisarAdmins_('Vídeo não gerado: ' + e.titulo, [e.album ? 'Este evento usa um álbum do Google Fotos, que a app não consegue ler.' : 'Este evento não tem pasta de fotos no Google Drive.', 'O vídeo automático só funciona com a pasta de fotos do Google Drive.']);
    return true;
  }
  var n = 0; try { n = fotosPasta_(e).length; } catch (err) {}
  if (n >= VIDEO_MIN_FOTOS) { e.video_estado = 'pendente'; e.video_em = agora_(); return true; }
  e.video_estado = 'sem_fotos';
  avisarAdmins_('Vídeo não gerado: ' + e.titulo, ['A pasta de fotos do evento só tem ' + n + ' foto(s) (são precisas pelo menos ' + VIDEO_MIN_FOTOS + ').', 'Se chegarem mais fotos, abra o painel e carregue em «Vídeo» → «Gerar vídeo» no cartão do evento.']);
  return true;
}
function pastaVideos_() { return pastaRaiz_('pasta_videos', 'Comenda de Braga – vídeos dos eventos'); }
function chaveVideo_() { return PropertiesService.getScriptProperties().getProperty('VIDEO_CHAVE') || ''; }
// API usada só pelo gerador no GitHub (protegida pela chave secreta)
function apiVideo(chave, op, d) {
  var k = chaveVideo_();
  if (!k || String(chave || '') !== k) throw new Error('Chave do gerador de vídeo inválida.');
  d = d || {};
  if (op === 'contar') {
    var lim0 = new Date(Date.now() - 3 * 3600e3);
    return ler_('Eventos').filter(function (e) { return e.video_estado === 'pendente' || (e.video_estado === 'a_gerar' && (!e.video_em || new Date(String(e.video_em).replace(' ', 'T')) < lim0)); }).length;
  }
  if (op === 'lista') {
    return lock_(function () {
      var lim = new Date(Date.now() - 3 * 3600e3), out = [];
      ler_('Eventos').forEach(function (e) {
        var preso = e.video_estado === 'a_gerar' && (!e.video_em || new Date(String(e.video_em).replace(' ', 'T')) < lim);
        if (e.video_estado !== 'pendente' && !preso) return;
        var fotos = []; try { fotos = fotosPasta_(e); } catch (err) {}
        e.video_estado = 'a_gerar'; e.video_em = agora_(); atualizar_('Eventos', e);
        out.push({ id: e.id, titulo: e.titulo, data: e.data, local: e.local, fotos: fotos, incluir: infoVideo_(e).incluir || [] });
      });
      return { org: getDef_('nome_org') || CONFIG.NOME_ORG, eventos: out };
    });
  }
  var e = evento_(String(d.evento || ''));
  if (op === 'foto') {
    var f = DriveApp.getFileById(String(d.id || '')), dentro = false, ps = f.getParents();
    while (ps.hasNext()) if (ps.next().getId() === e.pasta_id) dentro = true;
    if (!dentro) throw new Error('Esta foto não pertence à pasta do evento.');
    return { dados: Utilities.base64Encode(f.getBlob().getBytes()) };
  }
  if (op === 'entregar') {
    var blob = Utilities.newBlob(Utilities.base64Decode(String(d.video || '')), 'video/mp4', 'Vídeo – ' + nomePasta_(e) + '.mp4');
    var fich = pastaVideos_().createFile(blob);
    var info = infoVideo_(e);
    info = { incluir: info.incluir || [], usadas: d.usadas || [], excluidas: (d.excluidas || []).slice(0, 200), total: +d.total || 0, duracao: +d.duracao || 0, musica: limpo_(d.musica, 300), anteriores: (info.anteriores || []).concat(e.video_id ? [e.video_id] : []).slice(-5) };
    lock_(function () { var x = evento_(e.id); x.video_estado = 'pronto'; x.video_id = fich.getId(); x.video_url = fich.getUrl(); x.video_info = JSON.stringify(info); x.video_em = agora_(); atualizar_('Eventos', x); });
    var ex = info.excluidas.slice(0, 40).map(function (f) { return '<li>' + esc_(f.nome) + ' — ' + esc_(f.motivo) + '</li>'; }).join('');
    var extra = caixa_('🎬 ' + 'Resumo', 'Duração: ' + info.duracao + ' s · ' + info.usadas.length + ' de ' + info.total + ' fotos usadas' + (info.musica ? '<br>' + esc_(info.musica) : '')) +
      (ex ? caixa_('Fotos que ficaram de fora (' + info.excluidas.length + ')', '<ul style="margin:4px 0 0;padding-left:18px">' + ex + '</ul>' + (info.excluidas.length > 40 ? '…' : '')) : '') +
      '<p style="margin:12px 0 0">' + lk_(url_({ v: 'admin' }), 'Abrir o painel para aprovar ou gerar de novo') + '</p>';
    listaAdmins_().forEach(function (a) {
      enfileirar_(a, 'Vídeo pronto para aprovar: ' + e.titulo, { linhas: ['O vídeo do evento «' + e.titulo + '» está pronto.', 'Veja-o e, se estiver bem, aprove-o no painel (cartão do evento → «Vídeo» → «Aprovar e enviar»). Os participantes só o recebem depois de aprovado.', 'Se uma foto boa ficou de fora, marque-a no painel e carregue em «Gerar de novo».'], extra: extra, botao: { texto: '▶ Ver o vídeo', url: fich.getUrl() } }, 2);
    });
    enviarJa_();
    return 'ok';
  }
  if (op === 'falhou') {
    var semFotos = /foto/i.test(String(d.motivo || '')) && !/técnico/.test(String(d.motivo || ''));
    lock_(function () { var x = evento_(e.id), inf = infoVideo_(x); inf.motivo = limpo_(d.motivo, 400); if (d.excluidas) inf.excluidas = d.excluidas.slice(0, 200); if (d.total) inf.total = +d.total; x.video_estado = semFotos ? 'sem_fotos' : 'erro'; x.video_info = JSON.stringify(inf); x.video_em = agora_(); atualizar_('Eventos', x); });
    avisarAdmins_('Vídeo não gerado: ' + e.titulo, [limpo_(d.motivo, 400), semFotos ? 'Se chegarem mais fotos, carregue em «Vídeo» → «Gerar vídeo» no cartão do evento.' : 'Pode tentar de novo no painel (cartão do evento → «Vídeo» → «Gerar de novo»).']);
    return 'ok';
  }
  throw new Error('Operação desconhecida.');
}
function apiAdminChaveVideo(tok, nova) {
  admin_(tok);
  var p = PropertiesService.getScriptProperties();
  if (nova || !p.getProperty('VIDEO_CHAVE')) p.setProperty('VIDEO_CHAVE', novoToken_() + novoToken_());
  return { chave: p.getProperty('VIDEO_CHAVE'), api: CONFIG.URL_API || getDef_('url_app') || '' };
}
function apiAdminVideo(tok, id, acao, incluir) {
  admin_(tok);
  var r = lock_(function () {
    var e = evento_(id), info = infoVideo_(e);
    if (acao === 'gerar') {
      if (!e.pasta_id) throw new Error('Este evento não tem pasta de fotos no Google Drive.');
      if (e.video_estado === 'a_gerar') throw new Error('O vídeo já está a ser gerado. Aguarde o email.');
      info.incluir = (incluir || []).map(String).slice(0, 50);
      e.video_estado = 'pendente'; e.video_em = agora_(); e.video_info = JSON.stringify(info); atualizar_('Eventos', e);
      return { msg: 'Pedido registado. O vídeo é gerado na próxima hora e recebe um email quando estiver pronto.' };
    }
    if (acao === 'aprovar') {
      if (e.video_estado !== 'pronto' || !e.video_id) throw new Error('Não há nenhum vídeo pronto para aprovar.');
      DriveApp.getFileById(e.video_id).setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      e.video_estado = 'aprovado'; e.video_em = agora_(); atualizar_('Eventos', e);
      return { e: e, n: enviarVideo_(e) };
    }
    throw new Error('Ação desconhecida.');
  });
  enviarJa_();
  return r.msg || ('Vídeo aprovado e enviado a ' + r.n + ' pessoa(s).');
}
// Envia o link do vídeo a quem participou (inscritos e convidados com email)
function enviarVideo_(e) {
  var subs = ler_('Inscritos'), enviados = {}, n = 0, porId = {};
  var pres = presentesEmail_(e);
  var corpo = function (nome, x) { return [saud_(nome, x), tr_(x, 'Já pode ver o vídeo de «' + titE_(e, x) + '», feito com as fotos partilhadas pelos participantes.', 'The video of «' + titE_(e, x) + '», made with the photos shared by those who took part, is now available.'), quando_(e, x), tr_(x, 'Obrigado a todos os que enviaram fotografias.', 'Thank you to everyone who sent photos.')]; };
  if (pres) {
    pres.forEach(function (p) { n++; enfileirar_(p.email, tr_(p.s, 'O vídeo do evento: ', 'The video of the event: ') + titE_(e, p.s), { lang: p.s && p.s.lingua, linhas: corpo(p.nome, p.s), botao: { texto: tr_(p.s, '▶ Ver o vídeo', '▶ Watch the video'), url: e.video_url }, sair: p.s ? url_({ sair: p.s.token }) : '' }, 4); });
    return n;
  }
  ler_('Inscricoes').filter(function (r) { return r.evento_id === e.id; }).forEach(function (r) {
    var s = subs.filter(function (x) { return x.email === r.email && x.estado === 'confirmado'; })[0];
    if (!s) return; porId[r.id] = s;
    if (enviados[s.email]) return; enviados[s.email] = 1; n++;
    enfileirar_(s.email, tr_(s, 'O vídeo do evento: ', 'The video of the event: ') + titE_(e, s), { lang: s.lingua, linhas: corpo(r.nome || s.nome, s), botao: { texto: tr_(s, '▶ Ver o vídeo', '▶ Watch the video'), url: e.video_url }, sair: url_({ sair: s.token }) }, 4);
  });
  ler_('Convidados').filter(function (g) { return g.evento_id === e.id && g.email; }).forEach(function (g) {
    var em = String(g.email).toLowerCase(), x = porId[g.inscricao_id] || null;
    if (!emailOk_(em) || enviados[em]) return; enviados[em] = 1; n++;
    enfileirar_(em, tr_(x, 'O vídeo do evento: ', 'The video of the event: ') + titE_(e, x), { lang: x && x.lingua, linhas: corpo(g.nome, x), botao: { texto: tr_(x, '▶ Ver o vídeo', '▶ Watch the video'), url: e.video_url } }, 4);
  });
  return n;
}
/* ---- Presenças (registadas pelos administradores no dia e até 7 dias depois) ---- */
var PRESENCAS_DIAS = 7;
// Quem esteve presente e tem email (para agradecimento e vídeo). null = evento sem presenças registadas.
function presentesEmail_(e) {
  var P = ler_('Presencas').filter(function (p) { return p.evento_id === e.id; });
  if (!P.length) return null;
  var subs = ler_('Inscritos'), out = [], vistos = {};
  P.forEach(function (p) {
    var em = String(p.email || '').toLowerCase();
    if (!emailOk_(em) || vistos[em]) return; vistos[em] = 1;
    var s = subs.filter(function (x) { return x.email === em && x.estado === 'confirmado'; })[0] || null;
    out.push({ email: em, nome: p.nome, s: s, de: p.convidado_de });
  });
  return out;
}
function janelaPresencas_(e) { var d = diasAte_(e.data); return d <= 0 && d >= -PRESENCAS_DIAS; }
// ops: [{chave, presente, dados?}] — em lote, para a fila do telemóvel (sem rede) poder ser enviada de uma vez
function apiAdminPresencas(tok, eventoId, ops) {
  var quem = admin_(tok);
  var res = lock_(function () {
    var e = evento_(eventoId);
    if (e.externo) throw new Error('Os eventos de outra comenda não têm presenças.');
    if (!janelaPresencas_(e)) throw new Error(diasAte_(e.data) > 0 ? 'As presenças só podem ser marcadas a partir do dia do evento.' : 'Já passaram mais de ' + PRESENCAS_DIAS + ' dias sobre o evento: as presenças estão fechadas.');
    var regs = ler_('Inscricoes').filter(function (r) { return r.evento_id === e.id; });
    var conv = ler_('Convidados').filter(function (g) { return g.evento_id === e.id; });
    var I = ler_('Inscritos'), avisos = [], convites = 0;
    (ops || []).slice(0, 500).forEach(function (o) {
      var chave = limpo_(o && o.chave, 200); if (!chave) return;
      var atual = ler_('Presencas').filter(function (p) { return p.evento_id === e.id && p.chave === chave; })[0];
      if (!o.presente) { if (atual) apagar_('Presencas', function (p) { return p.evento_id === e.id && p.chave === chave; }); return; }
      if (atual) return;
      var p = { evento_id: e.id, chave: chave, marcado_por: quem, marcado_em: agora_(), sem_inscricao: '', convidado_de: '' }, t = chave.slice(0, 2), id = chave.slice(2);
      if (t === 'r:') {
        var r = regs.filter(function (x) { return x.id === id; })[0]; if (!r) { avisos.push('Inscrição não encontrada (' + id + ').'); return; }
        var s = I.filter(function (x) { return x.email === r.email; })[0] || {};
        p.tipo = ehSim_(s.membro) ? 'membro' : 'inscrito'; p.nome = r.nome; p.email = r.email; p.n_ordem = ehSim_(s.membro) ? s.n_ordem : ''; p.membro = sim_(ehSim_(s.membro)); p.localidade = s.localidade || ''; p.telemovel = r.telemovel || s.telemovel || '';
      } else if (t === 'g:') {
        var k = id.indexOf(':'), rid = id.slice(0, k), gn = id.slice(k + 1);
        var g = conv.filter(function (x) { return x.inscricao_id === rid && x.nome === gn; })[0]; if (!g) { avisos.push('Convidado não encontrado (' + gn + ').'); return; }
        p.tipo = 'convidado'; p.nome = g.nome; p.email = g.email || ''; p.membro = 'NÃO'; p.telemovel = g.telemovel || ''; p.convidado_de = g.convidado_de;
      } else if (t === 'e:') {
        var s2 = I.filter(function (x) { return x.email === id.toLowerCase(); })[0]; if (!s2) { avisos.push('Inscrito não encontrado (' + id + ').'); return; }
        p.tipo = ehSim_(s2.membro) ? 'membro' : 'inscrito'; p.nome = s2.nome; p.email = s2.email; p.n_ordem = ehSim_(s2.membro) ? s2.n_ordem : ''; p.membro = sim_(ehSim_(s2.membro)); p.localidade = s2.localidade || ''; p.telemovel = s2.telemovel || ''; p.sem_inscricao = 'SIM';
      } else if (t === 'm:') {
        var mk = normChave_(id), m = ler_('Membros').filter(function (x) { return chaveMembro_(x) === mk; })[0]; if (!m) { avisos.push('Membro não encontrado (' + id + ').'); return; }
        var cv = ehConvidado_(m);
        p.tipo = cv ? 'convidado' : 'membro'; p.nome = m.nome; p.email = String(m.email || '').toLowerCase(); p.n_ordem = cv ? '' : m.n_ordem; p.membro = sim_(!cv); p.localidade = m.localidade || ''; p.telemovel = m.telemovel || ''; p.sem_inscricao = 'SIM';
      } else if (t === 'x:') {
        var d = o.dados || {}, nome = limpo_(d.nome, 120), email = limpo_(d.email, 120).toLowerCase(), num = limpo_(d.n_ordem, 30);
        if (!nome || !emailOk_(email) || !limpo_(d.telemovel, 40) || !limpo_(d.localidade, 80)) { avisos.push('Dados incompletos para ' + (nome || 'uma pessoa') + ': nome, localidade, contacto e email são obrigatórios.'); return; }
        var mm = num ? membroPor_(email, num).num : null;
        p.tipo = num ? 'membro' : 'nao_inscrito'; p.nome = nome; p.email = email; p.n_ordem = num; p.membro = sim_(!!num); p.localidade = limpo_(d.localidade, 80); p.telemovel = limpo_(d.telemovel, 40); p.sem_inscricao = 'SIM';
        if (num && !mm) avisos.push(nome + ' indicou o nº da Ordem ' + num + ', que não existe na base de membros.');
        // convite para receber as notificações (a pessoa confirma por email)
        var ex = I.filter(function (x) { return x.email === email; })[0];
        if (!ex) {
          var ns = { email: email, nome: nome, telemovel: p.telemovel, localidade: p.localidade, nota: 'Presença em ' + e.data + ' · ' + e.titulo, estado: 'pendente', membro: 'NÃO', token: novoToken_(), criado_em: agora_(), confirmado_em: '', ultimo_envio_confirmacao: new Date().toISOString(), n_ordem: '', lingua: '' };
          identificar_(ns, num); acrescentar_('Inscritos', ns); I.push(ns);
          var org = getDef_('nome_org') || CONFIG.NOME_ORG;
          enfileirar_(email, 'Obrigado pela sua presença — receba os convites da ' + org, { linhas: [saud_(nome), 'Foi um gosto tê-lo(a) connosco em «' + e.titulo + '» (' + dataLonga_(e.data) + ').', 'Se quiser receber os convites para os próximos eventos da ' + org + ', confirme o seu email no botão abaixo. Se não confirmar, não lhe enviamos mais nada.'], botao: { texto: 'Quero receber os convites', url: url_({ confirmar: ns.token }) } }, 2);
          convites++;
        }
      } else return;
      acrescentar_('Presencas', p);
    });
    return { presencas: ler_('Presencas').filter(function (p) { return p.evento_id === e.id; }).map(presencaPub_), avisos: avisos, convites: convites };
  });
  if (res.convites) enviarJa_();
  return res;
}
function presencaPub_(p) { return { evento_id: p.evento_id, chave: p.chave, tipo: p.tipo, nome: p.nome, email: p.email, n_ordem: p.n_ordem, membro: ehSim_(p.membro), localidade: p.localidade, telemovel: p.telemovel, convidado_de: p.convidado_de, sem_inscricao: ehSim_(p.sem_inscricao), marcado_por: p.marcado_por, marcado_em: p.marcado_em }; }
/* ---- Eventos abertos para quem chega depois do convite ---- */
function eventosAbertos_(s) {
  var hoje = hoje_(), regs = {};
  ler_('Inscricoes').forEach(function (r) { if (r.email === s.email) regs[r.evento_id] = 1; });
  ler_('Recusas').forEach(function (r) { if (r.email === s.email) regs[r.evento_id] = 1; });
  return ler_('Eventos').filter(function (e) { return e.estado === 'ativo' && !e.externo && e.prazo >= hoje && !regs[e.id] && podeVer_(e, s); })
    .sort(function (a, b) { return (a.data + a.hora).localeCompare(b.data + b.hora); });
}
// Envia UM email com os eventos que ainda aceitam inscrições (para quem se inscreveu depois dos convites)
function convitesTardios_(s) {
  if (!s || s.estado !== 'confirmado') return 0;
  var evs = eventosAbertos_(s); if (!evs.length) return 0;
  var org = getDef_('nome_org') || CONFIG.NOME_ORG;
  var extra = caixasEventos_(evs, s);
  enfileirar_(s.email, evs.length === 1 ? tr_(s, 'Evento com inscrições abertas: ', 'Event open for registration: ') + titE_(evs[0], s) : tr_(s, 'Eventos com inscrições abertas na ', 'Events open for registration at ') + org, {
    lang: s.lingua,
    linhas: [saud_(s.nome, s), evs.length === 1 ? tr_(s, 'Ainda está a tempo de se inscrever neste evento da ' + org + ':', 'You can still register for this event by ' + org + ':') : tr_(s, 'Ainda está a tempo de se inscrever nestes ' + evs.length + ' eventos da ' + org + ':', 'You can still register for these ' + evs.length + ' events by ' + org + ':')],
    extra: extra,
    botao: { texto: tr_(s, 'Ver o calendário', 'View the calendar'), url: url_({ t: s.token }) }, sair: url_({ sair: s.token })
  }, 3);
  return evs.length;
}
function ehTalvez_(r) { return !!(r && r.tipo === 'talvez'); }
function linkTalvez_(e, s) { return { texto: tr_(s, 'Talvez vá', 'I might attend'), url: url_({ t: s.token, ev: e.id, talvez: 1 }) }; }
function linkRecusar_(e, s) { return { texto: tr_(s, 'Não quero participar', 'I will not attend'), url: url_({ t: s.token, ev: e.id, nao: 1 }) }; }
function recusasEvento_(eid) { var o = {}; ler_('Recusas').forEach(function (r) { if (r.evento_id === eid) o[r.email] = r; }); return o; }
// Caixas (uma por evento) com «Inscrever-me» e «Não quero participar»
function caixasEventos_(evs, s, talvez) {
  return evs.map(function (e) {
    var tv = talvez && talvez[e.id];
    return caixa_(esc_(titE_(e, s)), esc_(quando_(e, s)) + '<br>' + esc_(e.local) + '<br>' + esc_(tr_(s, 'Inscrições até ', 'Registration closes on ') + dataLonga_(e.prazo, s)) +
      '<br><a href="' + esc_(url_({ t: s.token, ev: e.id })) + '" style="display:inline-block;margin-top:8px;background:#b0001e;color:#ffffff;text-decoration:none;padding:8px 14px;border-radius:8px;font-weight:bold">' + tr_(s, 'Inscrever-me', 'Register') + '</a>' +
      (tv ? '' : ' &nbsp;<a href="' + esc_(url_({ t: s.token, ev: e.id, talvez: 1 })) + '" style="display:inline-block;margin-top:8px;color:#6a6264;font-size:14px">' + tr_(s, 'Talvez vá', 'I might attend') + '</a>') +
      ' &nbsp;<a href="' + esc_(url_({ t: s.token, ev: e.id, nao: 1 })) + '" style="display:inline-block;margin-top:8px;color:#6a6264;font-size:14px">' + tr_(s, 'Não quero participar', 'I will not attend') + '</a>' +
      (tv ? '<br><span style="font-size:13px;color:#6a6264">' + tr_(s, 'A sua resposta: talvez vá.', 'Your reply: you might attend.') + '</span>' : ''));
  }).join('');
}
// Terças-feiras (na tarefa diária das 9h): um email por pessoa com os eventos abertos a que ainda não respondeu
function lembreteSemanal_(forcar) {
  var hoje = hoje_();
  if (!forcar && (new Date(hoje + 'T12:00:00').getDay() !== 2 || getDef_('lembrete_semanal') === hoje)) return 0;
  setDef_('lembrete_semanal', hoje);
  var resp = {};
  ler_('Inscricoes').forEach(function (r) { resp[r.evento_id + '|' + r.email] = 1; });
  var tvs = {};
  ler_('Recusas').forEach(function (r) { if (ehTalvez_(r)) tvs[r.evento_id + '|' + r.email] = 1; else resp[r.evento_id + '|' + r.email] = 1; });
  var evs = ler_('Eventos').filter(function (e) {
    if (e.estado !== 'ativo' || e.externo || e.prazo < hoje) return false;
    var d = diasAte_(e.data); if (d === 7 || d === 1) return false; // nesses dias já sai o lembrete normal
    if (String(e.criado_em || '').slice(0, 10) === hoje) return false;  // convite acabado de enviar
    return !(+e.lotacao && pessoasEvento_(e.id) >= +e.lotacao);
  }).sort(function (a, b) { return (a.data + a.hora).localeCompare(b.data + b.hora); });
  if (!evs.length) return 0;
  var org = getDef_('nome_org') || CONFIG.NOME_ORG, n = 0;
  ler_('Inscritos').forEach(function (s) {
    if (s.estado !== 'confirmado') return;
    var meus = evs.filter(function (e) { return !resp[e.id + '|' + s.email] && podeVer_(e, s); });
    if (!meus.length) return;
    var meusTv = {}; meus.forEach(function (e) { if (tvs[e.id + '|' + s.email]) meusTv[e.id] = 1; });
    enfileirar_(s.email, meus.length === 1 ? tr_(s, 'Lembrete: ', 'Reminder: ') + titE_(meus[0], s) : tr_(s, 'Lembrete: ' + meus.length + ' eventos aguardam a sua resposta', 'Reminder: ' + meus.length + ' events awaiting your reply'), {
      lang: s.lingua,
      linhas: [saud_(s.nome, s), meus.length === 1 ? tr_(s, 'Ainda não confirmou a sua presença neste evento da ' + org + ':', 'You have not yet confirmed whether you will attend this event by ' + org + ':') : tr_(s, 'Ainda não confirmou a sua presença nestes eventos da ' + org + ':', 'You have not yet confirmed whether you will attend these events by ' + org + ':'),
        tr_(s, 'Inscreva-se ou, se não puder ir, carregue em «Não quero participar» para deixar de receber este lembrete.', 'Register or, if you cannot attend, click «I will not attend» to stop receiving this reminder.')],
      extra: caixasEventos_(meus, s, meusTv),
      botao: { texto: tr_(s, 'Ver o calendário', 'View the calendar'), url: url_({ t: s.token }) }, sair: url_({ sair: s.token })
    }, 4);
    n++;
  });
  return n;
}
// «Não quero participar» (com nota opcional)
function apiRecusar(token, eventoId, nota, lang, tipo) {
  var x = lang, tv = tipo === 'talvez';
  lock_(function () {
    var s = ler_('Inscritos').filter(function (y) { return y.token === token && y.estado === 'confirmado'; })[0];
    if (!s) throw new Error(tr_(lang, 'O seu acesso não é válido. Abra o link do último email recebido.', 'Your access is not valid. Please open the link in the latest email you received.'));
    x = s;
    var e = ler_('Eventos').filter(function (y) { return y.id === eventoId; })[0];
    if (!e || e.estado !== 'ativo' || e.externo || !podeVer_(e, s)) throw new Error(tr_(s, 'Este evento já não está disponível.', 'This event is no longer available.'));
    if (ler_('Inscricoes').some(function (r) { return r.evento_id === e.id && r.email === s.email; })) throw new Error(tr_(s, 'Está inscrito neste evento. Para não ir, abra a inscrição e use «Anular inscrição».', 'You are registered for this event. To withdraw, open your registration and use «Cancel registration».'));
    var r = recusasEvento_(e.id)[s.email], nt = limpo_(nota, 500);
    if (r) { r.nota = nt; r.nome = s.nome; r.criado_em = agora_(); r.tipo = tv ? 'talvez' : ''; atualizar_('Recusas', r); }
    else acrescentar_('Recusas', { evento_id: e.id, email: s.email, nome: s.nome, nota: nt, criado_em: agora_(), tipo: tv ? 'talvez' : '' });
    acrescentar_('Historico', { data: agora_(), email: s.email, nome: s.nome, evento_id: e.id, evento: e.data + ' · ' + e.titulo, acao: tv ? 'Talvez vá' : 'Não participa', motivo: nt });
  });
  if (tv) return tr_(x, 'Obrigado pela resposta. Registámos que talvez vá. Vamos continuar a enviar-lhe lembretes até decidir.', 'Thank you. We have noted that you might attend and will keep reminding you until you decide.');
  return tr_(x, 'Obrigado pela resposta. Registámos que não vai participar. Se mudar de ideias, ainda se pode inscrever.', 'Thank you for letting us know. If you change your mind, you can still register.');
}
// Administradores inscrevem uma ou mais pessoas (inscritos e/ou membros da base)
function apiAdminInscreverPessoas(tok, eventoId, chaves, o) {
  var quem = admin_(tok); o = o || {};
  var res = lock_(function () {
    var e = evento_(eventoId), campos = jsonOu_(e.campos, {});
    if (e.externo) throw new Error('Os eventos de outra comenda não têm inscrições.');
    if (e.estado !== 'ativo' && e.estado !== 'oculto') throw new Error('Só é possível inscrever pessoas em eventos publicados.');
    var nConv = campos.convidados ? Math.max(0, Math.min(10, parseInt(o.convidados, 10) || 0)) : 0;
    var agape = !!(campos.agape && o.agape);
    var ok = [], falhas = [];
    (chaves || []).forEach(function (c) {
      c = String(c || ''); var s = null, I = ler_('Inscritos');
      try {
        if (c.indexOf('e:') === 0) {
          s = I.filter(function (x) { return x.email === c.slice(2).toLowerCase(); })[0];
          if (!s) throw new Error('não encontrado');
          if (s.estado !== 'confirmado') throw new Error('ainda não confirmou o email');
        } else if (c.indexOf('m:') === 0) {
          var k = normChave_(c.slice(2)), m = ler_('Membros').filter(function (x) { return chaveMembro_(x) === k; })[0];
          if (!m) throw new Error('membro não encontrado');
          var email = String(m.email || '').toLowerCase();
          if (!emailOk_(email)) throw new Error('sem email na base de membros');
          s = I.filter(function (x) { return x.email === email; })[0];
          if (s && s.estado !== 'confirmado') { s.estado = 'confirmado'; s.confirmado_em = agora_(); atualizar_('Inscritos', s); }
          if (!s) {
            s = { email: email, nome: m.nome, telemovel: m.telemovel || '', localidade: m.localidade || '', nota: '', estado: 'confirmado', membro: ehConvidado_(m) ? 'NÃO' : 'SIM', token: novoToken_(), criado_em: agora_(), confirmado_em: agora_(), ultimo_envio_confirmacao: '', n_ordem: ehConvidado_(m) ? '' : m.n_ordem, lingua: '' };
            acrescentar_('Inscritos', s);
            s = ler_('Inscritos').filter(function (x) { return x.email === email; })[0];
          }
        } else throw new Error('pessoa inválida');
        if (ler_('Inscricoes').some(function (r) { return r.evento_id === e.id && r.email === s.email; })) throw new Error('já estava inscrito(a)');
        var conv = [];
        for (var i = 1; i <= nConv; i++) conv.push({ nome: tr_(s, 'Convidado(a) ' + i + ' de ', 'Guest ' + i + ' of ') + s.nome, agape: agape });
        inscrever_(s, e.id, { nome: s.nome, telemovel: s.telemovel, agape: agape, convidados: conv, observacoes: limpo_(o.nota, 1000) }, false, quem);
        ok.push(s.nome);
      } catch (err) { falhas.push((s && s.nome || c.slice(2)) + ': ' + String(err.message || err).replace(/^Error: /, '')); }
    });
    return { ok: ok, falhas: falhas };
  });
  enviarJa_();
  return res;
}
function apiAdminEnviarAbertos(tok, email) {
  admin_(tok);
  var n = lock_(function () {
    var s = ler_('Inscritos').filter(function (x) { return x.email === String(email || '').toLowerCase(); })[0];
    if (!s) throw new Error('Inscrito não encontrado.');
    if (s.estado !== 'confirmado') throw new Error('Esta pessoa ainda não confirmou o email.');
    return convitesTardios_(s);
  });
  enviarJa_();
  return n ? 'Enviado um email com ' + n + ' evento(s) com inscrições abertas.' : 'Não há eventos com inscrições abertas para esta pessoa (ou já está inscrita em todos).';
}
// Editar um inscrito (incluindo o email: as inscrições em eventos passam para o email novo)
function apiAdminEditarInscrito(tok, emailOriginal, d) {
  admin_(tok); d = d || {};
  var msg = lock_(function () {
    var antigo = String(emailOriginal || '').toLowerCase();
    var s = ler_('Inscritos').filter(function (x) { return x.email === antigo; })[0];
    if (!s) throw new Error('Inscrito não encontrado.');
    var email = limpo_(d.email, 120).toLowerCase(), nome = limpo_(d.nome, 120);
    if (!nome) throw new Error('Indique o nome.');
    if (!emailOk_(email)) throw new Error('Indique um email válido.');
    if (email !== antigo && ler_('Inscritos').some(function (x) { return x.email === email; })) throw new Error('Já existe um inscrito com o email ' + email + '.');
    s.nome = nome; s.telemovel = limpo_(d.telemovel, 40); s.localidade = limpo_(d.localidade, 80); s.nota = limpo_(d.nota, 1000);
    s.lingua = d.lingua === 'en' ? 'en' : (d.lingua === 'pt' ? 'pt' : '');
    var membro = !!d.membro; if (membro && !limpo_(d.n_ordem, 30)) throw new Error('Para ser membro tem de indicar o nº da Ordem. Sem nº, a pessoa fica como convidado.');
    s.membro = sim_(membro); s.n_ordem = membro ? limpo_(d.n_ordem, 30) : ''; s.verificar = '';
    var mudou = email !== antigo;
    if (mudou) {
      s.email = email;
      ler_('Inscricoes').forEach(function (r) { if (r.email === antigo) { r.email = email; atualizar_('Inscricoes', r); } });
      ler_('Recusas').forEach(function (r) { if (r.email === antigo) { r.email = email; atualizar_('Recusas', r); } });
      try { ler_('Dispositivos').forEach(function (d) { if (d.email === antigo) { d.email = email; atualizar_('Dispositivos', d); } }); } catch (e) {}
      apagar_('PreInscricoes', function (x) { return x.email === antigo; });
    }
    atualizar_('Inscritos', s);
    if (mudou && s.estado === 'confirmado') {
      var org = getDef_('nome_org') || CONFIG.NOME_ORG;
      enfileirar_(email, tr_(s, 'O seu email foi atualizado na ', 'Your email was updated at ') + org, { lang: s.lingua, linhas: [saud_(s.nome, s), tr_(s, 'A ' + org + ' atualizou o seu email de contacto para este endereço. A partir de agora recebe aqui os convites, os lembretes e as novidades.', org + ' updated your contact email to this address. From now on you will receive invitations, reminders and news here.'), tr_(s, 'As suas inscrições em eventos mantêm-se.', 'Your event registrations remain unchanged.')], botao: { texto: tr_(s, 'Abrir o calendário', 'Open the calendar'), url: url_({ t: s.token }) }, sair: url_({ sair: s.token }) }, 1);
    }
    return 'Dados atualizados.' + (mudou ? ' O email passou para ' + email + (s.estado === 'confirmado' ? ' (foi enviado um aviso para o email novo).' : '.') : '');
  });
  enviarJa_();
  return msg;
}
function pessoasEvento_(eid) {
  var regs = ler_('Inscricoes').filter(function (r) { return r.evento_id === eid; });
  var conv = ler_('Convidados').filter(function (g) { return g.evento_id === eid; });
  return regs.length + conv.length;
}

function confirmar_(token) {
  return lock_(function () {
    var s = ler_('Inscritos').filter(function (x) { return x.token === token; })[0];
    if (!s) return { ok: false, msg: 'Este link já não é válido. Peça novamente para receber as notificações. / This link is no longer valid. Please subscribe again.' };
    if (s.estado === 'confirmado') return { ok: true, msg: tr_(s, 'A sua inscrição já estava confirmada.', 'Your subscription was already confirmed.') };
    s.estado = 'confirmado'; s.confirmado_em = agora_();
    atualizar_('Inscritos', s);
    enfileirar_(s.email, tr_(s, 'Inscrição confirmada', 'Subscription confirmed'), {
      lang: s.lingua,
      linhas: [saud_(s.nome, s), tr_(s, 'A sua inscrição está confirmada. A partir de agora recebe os convites para os eventos, com o formulário de inscrição de cada um.', 'Your subscription is confirmed. From now on you will receive invitations to our events, each with its registration form.'), tr_(s, 'Guarde este email: o botão abaixo abre o calendário com o seu acesso pessoal.', 'Keep this email: the button below opens the calendar with your personal access.')],
      botao: { texto: tr_(s, 'Abrir o calendário', 'Open the calendar'), url: url_({ t: s.token }) }, sair: url_({ sair: s.token })
    }, 1);
    try { convitesTardios_(s); } catch (err) {}
    avisarAdmins_('Novo inscrito confirmado: ' + s.nome, [s.nome + ' (' + s.email + ') confirmou a inscrição nas notificações.', s.telemovel ? 'Contacto: ' + s.telemovel : '', s.localidade ? 'Localidade: ' + s.localidade : '', s.nota ? 'Nota: ' + s.nota : '', 'Se for membro, marque-o como «Membro» no painel (separador Inscritos).']);
    enviarJa_();
    return { ok: true, msg: tr_(s, 'Inscrição confirmada. Bem-vindo(a)!', 'Subscription confirmed. Welcome!') };
  });
}

/* ============ API pública (chamada pela app) ============ */

function apiDados(token) {
  var me = token ? ler_('Inscritos').filter(function (s) { return s.token === token && s.estado === 'confirmado'; })[0] : null;
  var membro = me && ehSim_(me.membro);
  var eventos = ler_('Eventos').filter(function (e) { return (e.estado === 'ativo' || e.estado === 'cancelado') && podeVer_(e, me); }).map(evPublico_);
  var out = { org: getDef_('nome_org') || CONFIG.NOME_ORG, hoje: hoje_(), eventos: eventos, me: null };
  if (me) {
    try { marcarVistoApp_(me.email); } catch (e) {}
    var regs = ler_('Inscricoes').filter(function (r) { return r.email === me.email; });
    var conv = ler_('Convidados');
    out.me = {
      recusas: ler_('Recusas').filter(function (r) { return r.email === me.email; }).map(function (r) { return { evento_id: r.evento_id, nota: r.nota, talvez: ehTalvez_(r) }; }),
      email: me.email, nome: me.nome, telemovel: me.telemovel, membro: membro, n_ordem: membro ? me.n_ordem : '', direcao: ehDirecao_(me), ical: urlFeed_(me), calLigado: calLigado_(me), lingua: me.lingua || '',
      nif: (function () { if (!membro) return ''; var mm = membroPor_(me.email, me.n_ordem); var x = mm.email || mm.num; return x ? x.nif : ''; })(),
      inscricoes: regs.map(function (r) {
        return {
          evento_id: r.evento_id, nome: r.nome, telemovel: r.telemovel, agape: ehSim_(r.agape), restricoes: r.restricoes, atividades: jsonOu_(r.atividades, []), observacoes: r.observacoes, visita: ehSim_(r.visita), nif: r.nif,
          convidados: conv.filter(function (g) { return g.inscricao_id === r.id; }).map(function (g) { return { nome: g.nome, email: g.email, telemovel: g.telemovel, agape: ehSim_(g.agape), restricoes: g.restricoes, atividades: jsonOu_(g.atividades, []), visita: ehSim_(g.visita) }; })
        };
      })
    };
  }
  return out;
}

/* ============ Membros ============ */
function ehConvidado_(m) { return !!m && (m.tipo === 'CONVIDADO' || !m.n_ordem); }
function chaveMembro_(m) { return m.n_ordem ? normNum_(m.n_ordem) : '@' + String(m.email || '').toLowerCase(); }
function normChave_(c) { c = String(c || '').trim(); return c.indexOf('@') >= 0 ? '@' + c.replace(/^@/, '').toLowerCase() : normNum_(c); }
function normNum_(n) { return limpo_(n, 30).replace(/\s+/g, '').toUpperCase(); }
function membroPor_(email, num) {
  var M = ler_('Membros'); email = (email || '').toLowerCase(); num = normNum_(num);
  M = M.filter(function (m) { return !ehConvidado_(m); });
  var porEmail = email ? M.filter(function (m) { return m.email && m.email.toLowerCase() === email; })[0] : null;
  var porNum = num ? M.filter(function (m) { return m.n_ordem && normNum_(m.n_ordem) === num; })[0] : null;
  return { email: porEmail, num: porNum };
}
// Marca o inscrito como membro se o email ou o nº coincidir. Devolve avisos para os admins.
// Nome comparável: sem acentos nem maiúsculas, sem partículas (de, da, dos…); conta o primeiro e o último nome
function chaveNome_(n) {
  var p = String(n || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase().replace(/[^a-z\s'-]/g, ' ').split(/\s+/)
    .filter(function (x) { return x && ['de', 'da', 'do', 'das', 'dos', 'e', "d'"].indexOf(x) < 0; });
  return p.length ? [p[0], p[p.length - 1]] : [];
}
function nomeIgual_(a, b) {
  var x = chaveNome_(a), y = chaveNome_(b);
  return x.length > 0 && y.length > 0 && x[0] === y[0] && x[1] === y[1];
}
// Reconhece um membro só se email, nome (primeiro e último) e nº da Ordem (se indicado) coincidirem com a mesma ficha.
// Se só parte coincidir, fica pendente («verificar») para os admins confirmarem.
function identificar_(s, numDado) {
  var r = membroPor_(s.email, numDado), m = r.email || r.num, avisos = [];
  numDado = limpo_(numDado, 30);
  if (!m) {
    if (numDado) avisos.push(s.nome + ' (' + s.email + ') indicou o nº da Ordem ' + numDado + ', que não existe na base de membros. Não foi marcado como membro.');
    return avisos;
  }
  var dif = [];
  if (!m.email || m.email.toLowerCase() !== String(s.email || '').toLowerCase()) dif.push('email: indicou ' + s.email + (m.email ? ', na base está ' + m.email : ', a ficha não tem email'));
  if (!nomeIgual_(m.nome, s.nome)) dif.push('nome: indicou «' + s.nome + '», na base está «' + m.nome + '»');
  if (numDado && normNum_(numDado) !== normNum_(m.n_ordem)) dif.push('nº da Ordem: indicou ' + numDado + ', na base está ' + m.n_ordem);
  if (r.email && r.num && r.email !== r.num) dif.push('o email pertence ao membro nº ' + r.email.n_ordem + ' e o nº indicado ao membro ' + r.num.nome);
  if (!dif.length) {
    s.membro = 'SIM'; s.n_ordem = m.n_ordem; s.verificar = '';
    if (!s.telemovel && m.telemovel) s.telemovel = m.telemovel; if (!s.localidade && m.localidade) s.localidade = m.localidade;
    return avisos;
  }
  if (ehSim_(s.membro)) return avisos;  // já era membro (confirmado antes): não muda
  s.verificar = JSON.stringify({ n_ordem: m.n_ordem, nome: m.nome, dif: dif });
  avisos.push(s.nome + ' (' + s.email + ') parece ser o membro nº ' + m.n_ordem + ' (' + m.nome + '), mas os dados não coincidem todos:', '• ' + dif.join('\n• '), 'Ficou inscrito como não membro. No painel, separador Inscritos, carregue em «Confirmar como membro» se for a mesma pessoa.');
  return avisos;
}
function sincronizarInscritos_() {
  var M = ler_('Membros'), porEmail = {};
  M.forEach(function (m) { if (m.email && !ehConvidado_(m)) porEmail[m.email.toLowerCase()] = m; });
  var n = 0;
  ler_('Inscritos').forEach(function (s) {
    var m = porEmail[s.email];
    if (m && nomeIgual_(m.nome, s.nome) && (s.membro !== 'SIM' || s.n_ordem !== m.n_ordem || s.verificar)) { s.membro = 'SIM'; s.n_ordem = m.n_ordem; s.verificar = ''; atualizar_('Inscritos', s); n++; }
  });
  return n;
}
var CAMPOS_MEMBRO = ['n_ordem', 'nome', 'email', 'telemovel', 'localidade', 'grau', 'cargo', 'data_admissao', 'nif', 'notas', 'direcao', 'tipo'];
var GRAUS_PADRAO = ['Postulante', 'Escudeiro', 'Cavaleiro / Dama', 'Comendador / Comendadora', 'Grã-Cruz'];
var CARGOS_PADRAO = ['Comendador da Comenda', 'Vice-Comendador', 'Chanceler', 'Secretário', 'Tesoureiro', 'Capelão', 'Mestre de Cerimónias', 'Hospitalário'];
function listaDef_(k, padrao) { var v = getDef_(k); return v ? String(v).split('\n').map(function (x) { return x.trim(); }).filter(Boolean) : padrao.slice(); }
function registarGraus_(hist) { hist.forEach(function (h) { acrescentar_('HistGraus', h); }); }
// Muda o grau ou o cargo de vários membros de uma vez (fica no histórico)
function apiAdminMudarGrau(tok, chaves, campo, valor, data) {
  var quem = admin_(tok);
  if (campo !== 'grau' && campo !== 'cargo') throw new Error('Campo inválido.');
  valor = limpo_(valor, 150); data = dataIso_(data) || hoje_();
  if (valor && listaDef_(campo === 'grau' ? 'graus' : 'cargos', campo === 'grau' ? GRAUS_PADRAO : CARGOS_PADRAO).indexOf(valor) < 0) throw new Error('Esse ' + campo + ' não está na lista das Definições.');
  var n = lock_(function () {
    var set = {}; (chaves || []).forEach(function (c) { set[normChave_(c)] = 1; });
    var M = ler_('Membros').map(function (m) { var c = {}; for (var k in m) c[k] = m[k]; return c; }), hist = [], agora = agora_();
    M.forEach(function (m) {
      if (!set[chaveMembro_(m)]) return;
      var antes = m[campo] || '';
      if (antes === valor) return;
      m[campo] = valor; m.atualizado_em = agora;
      hist.push({ data: data, n_ordem: m.n_ordem, nome: m.nome, campo: campo, antes: antes, depois: valor, por: quem, registado_em: agora });
    });
    if (hist.length) { escreverMembros_(M); registarGraus_(hist); }
    return hist.length;
  });
  return n ? (campo === 'grau' ? 'Grau' : 'Cargo') + ' atualizado em ' + n + ' membro(s).' : 'Nada a mudar: os membros escolhidos já tinham esse ' + campo + '.';
}
function limparMembro_(m) {
  var o = {};
  CAMPOS_MEMBRO.forEach(function (k) { o[k] = limpo_(m[k], k === 'notas' ? 1000 : 150); });
  o.n_ordem = limpo_(o.n_ordem, 30); o.email = o.email.toLowerCase();
  o.nif = o.nif.replace(/\s/g, ''); o.data_admissao = dataIso_(o.data_admissao);
  o.direcao = /^(sim|s|x|1|true|yes|y|verdadeiro|✓|✔)$/i.test(String(m.direcao === true ? 'sim' : m.direcao || '').trim()) ? 'SIM' : '';
  // Convidado: indicado no campo Tipo, ou sem nº da Ordem
  o.tipo = (/convid|guest/i.test(String(m.tipo || '')) || !o.n_ordem) ? 'CONVIDADO' : '';
  return o;
}
function dataIso_(d) {
  d = String(d || '').trim(); if (!d) return '';
  var m = d.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/); if (m) return m[1] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[3]).slice(-2);
  m = d.match(/^(\d{1,2})[\/.\-](\d{1,2})[\/.\-](\d{4})$/); if (m) return m[3] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[1]).slice(-2);
  m = d.match(/^(\d{4})$/); if (m) return m[1];
  return d;
}
function inscritoDeBase_(m, nota) {
  return { email: m.email, nome: m.nome, telemovel: m.telemovel || '', localidade: m.localidade || '', nota: [nota, m.notas].filter(Boolean).join(' · ').slice(0, 1000), estado: 'pendente', membro: 'NÃO', token: novoToken_(), criado_em: agora_(), confirmado_em: '', ultimo_envio_confirmacao: '', n_ordem: '', lingua: '' };
}
// Aplica a regra à base existente: quem não tem nº da Ordem (ou é convidado) sai da base e fica só nos Inscritos.
// executar=false: só mostra o que ia acontecer. Sem email: fica na base até ser completado ou removido à mão.
function apiAdminConvidadosBase(tok, executar) {
  admin_(tok);
  return lock_(function () {
    var M = ler_('Membros'), I = ler_('Inscritos'), por = {};
    I.forEach(function (x) { por[x.email] = x; });
    var r = { jaInscritos: [], novosPendentes: [], semEmail: [], membrosSemNum: I.filter(function (x) { return ehSim_(x.membro) && !x.n_ordem; }).map(function (x) { return { nome: x.nome, email: x.email }; }) }, ficam = [];
    M.forEach(function (m) {
      if (!ehConvidado_(m)) { ficam.push(m); return; }
      var em = String(m.email || '').toLowerCase(), it = { nome: m.nome, email: em, n_ordem: m.n_ordem || '' };
      if (!emailOk_(em)) { r.semEmail.push(it); ficam.push(m); return; }
      (por[em] ? r.jaInscritos : r.novosPendentes).push(it);
    });
    if (executar && r.membrosSemNum.length) {
      I.forEach(function (x) { if (ehSim_(x.membro) && !x.n_ordem) { x.membro = 'NÃO'; atualizar_('Inscritos', x); } });
      r.feito = true;
    }
    if (executar && (r.jaInscritos.length || r.novosPendentes.length)) {
      r.novosPendentes.forEach(function (it) { var m = M.filter(function (x) { return String(x.email || '').toLowerCase() === it.email; })[0]; acrescentar_('Inscritos', inscritoDeBase_({ email: it.email, nome: m.nome, telemovel: m.telemovel, localidade: m.localidade, notas: m.notas }, 'Vindo da base de membros (convidado)')); });
      escreverMembros_(ficam);
      r.feito = true;
    }
    return r;
  });
}
function validarMembro_(o) {
  if (!o.nome) return 'falta o nome';
  if (o.tipo === 'CONVIDADO' && !o.email) return 'sem nº da Ordem e sem email: não pode ficar nos Inscritos';
  if (o.email && !emailOk_(o.email)) return 'email inválido (' + o.email + ')';
  return '';
}
function escreverMembros_(lista) {
  _dir = null;
  var sh = folha_('Membros'), cab = TABELAS.Membros;
  cabecalho_(sh, cab);
  limparCache_('Membros');
  var ult = sh.getLastRow();
  if (ult > 1) sh.getRange(2, 1, ult - 1, cab.length).clearContent();
  if (lista.length) { var rg = sh.getRange(2, 1, lista.length, cab.length); rg.setNumberFormat('@'); rg.setValues(lista.map(function (o) { return cab.map(function (h) { return o[h] == null ? '' : String(o[h]); }); })); }
  SpreadsheetApp.flush(); limparCache_('Membros');
}
// Junta membros à base (atualiza pelo nº da Ordem, ou pelo email). Devolve {novos, atualizados, iguais, erros, plano}
// opt.manterVazios: células vazias não apagam dados existentes (importação). opt.simular: não grava (pré-visualização).
function juntarMembros_(entradas, chaveOriginal, opt) {
  opt = opt || {};
  var M = ler_('Membros').map(function (m) { var c = {}; for (var k in m) c[k] = m[k]; return c; });
  var res = { novos: 0, atualizados: 0, iguais: 0, erros: [], plano: [], paraInscritos: 0 }, agora = agora_(), mexidos = {}, hist = [], inscr = {};
  var idx = function (o) {
    if (chaveOriginal) { var c = normChave_(chaveOriginal); for (var z = 0; z < M.length; z++) if (chaveMembro_(M[z]) === c) return z; }
    var n = normNum_(o.n_ordem);
    if (n) for (var i = 0; i < M.length; i++) if (normNum_(M[i].n_ordem) === n) return i;
    if (o.email) for (var j = 0; j < M.length; j++) if (M[j].email && M[j].email.toLowerCase() === o.email) return j;
    return -1;
  };
  var vazio = function (v) { return v == null || String(v).trim() === ''; };
  entradas.forEach(function (e) {
    var bruto = e.dados || {};
    if (opt.manterVazios) {
      var i0 = idx(limparMembro_(bruto));
      if (i0 >= 0) {
        var j = {};
        CAMPOS_MEMBRO.forEach(function (k) { j[k] = vazio(bruto[k]) ? M[i0][k] : bruto[k]; });
        if (vazio(bruto.tipo) && !vazio(bruto.n_ordem)) j.tipo = '';
        bruto = j;
      }
    }
    var o = limparMembro_(bruto), err = validarMembro_(o);
    if (err) { res.erros.push({ linha: e.linha, motivo: err }); res.plano.push({ linha: e.linha, acao: 'erro', nome: o.nome || limpo_(bruto.nome, 120), motivo: err }); return; }
    // Regra: só entra na base quem tem nº da Ordem e não é convidado. Os restantes ficam só na lista de Inscritos.
    if (o.tipo === 'CONVIDADO') {
      if (!opt.paraInscritos) { var m0 = 'sem nº da Ordem: quem não tem nº (ou é convidado) fica só na lista de Inscritos'; res.erros.push({ linha: e.linha, motivo: m0 }); res.plano.push({ linha: e.linha, acao: 'erro', nome: o.nome, motivo: m0 }); return; }
      if (!o.email) { var m3 = 'sem nº da Ordem e sem email: não pode ficar nos Inscritos'; res.erros.push({ linha: e.linha, motivo: m3 }); res.plano.push({ linha: e.linha, acao: 'erro', nome: o.nome, motivo: m3 }); return; }
      var ja = ler_('Inscritos').some(function (x) { return x.email === o.email; }) || inscr[o.email];
      res.plano.push({ linha: e.linha, acao: 'inscrito', nome: o.nome, email: o.email, motivo: ja ? 'já está na lista de Inscritos' : 'entra nos Inscritos como pendente' });
      if (!ja) { res.paraInscritos = (res.paraInscritos || 0) + 1; inscr[o.email] = o; }
      return;
    }
    var i = idx(o);
    // o nº não pode pertencer a outro membro
    var dono = -1; if (o.n_ordem) for (var k = 0; k < M.length; k++) if (normNum_(M[k].n_ordem) === normNum_(o.n_ordem)) { dono = k; break; }
    if (dono >= 0 && i >= 0 && dono !== i) { var m1 = 'o nº ' + o.n_ordem + ' já pertence a ' + M[dono].nome; res.erros.push({ linha: e.linha, motivo: m1 }); res.plano.push({ linha: e.linha, acao: 'erro', nome: o.nome, motivo: m1 }); return; }
    if (o.email) { for (var q = 0; q < M.length; q++) if (q !== i && M[q].email && M[q].email.toLowerCase() === o.email) { var m2 = 'o email ' + o.email + ' já pertence a ' + M[q].nome + (M[q].n_ordem ? ' (nº ' + M[q].n_ordem + ')' : ''); res.erros.push({ linha: e.linha, motivo: m2 }); res.plano.push({ linha: e.linha, acao: 'erro', nome: o.nome, motivo: m2 }); return; } }
    if (i >= 0) {
      var mud = [];
      CAMPOS_MEMBRO.forEach(function (k) { var antes = String(M[i][k] == null ? '' : M[i][k]); if (antes !== o[k]) mud.push({ campo: k, antes: antes, depois: o[k] }); });
      var rep = mexidos[i];
      if (!mud.length) { if (!rep) res.iguais++; res.plano.push({ linha: e.linha, acao: rep ? 'repetida' : 'igual', nome: M[i].nome, n_ordem: M[i].n_ordem, com: rep || '' }); return; }
      mud.forEach(function (x) { if (x.campo === 'grau' || x.campo === 'cargo') hist.push({ data: hoje_(), n_ordem: o.n_ordem, nome: o.nome, campo: x.campo, antes: x.antes, depois: x.depois, por: opt.quem || '', registado_em: agora }); });
      CAMPOS_MEMBRO.forEach(function (k) { M[i][k] = o[k]; }); M[i].atualizado_em = agora;
      if (!rep) res.atualizados++;
      res.plano.push({ linha: e.linha, acao: rep ? 'repetida' : 'atualizar', nome: o.nome, n_ordem: o.n_ordem, mudancas: mud, com: rep || '' });
      if (!rep) mexidos[i] = e.linha;
    } else {
      o.criado_em = agora; o.atualizado_em = agora; M.push(o); res.novos++;
      mexidos[M.length - 1] = e.linha;
      res.plano.push({ linha: e.linha, acao: 'novo', nome: o.nome, n_ordem: o.n_ordem, email: o.email, tipo: o.tipo });
    }
  });
  if (opt.simular) return res;
  M.sort(function (a, b) { if (!a.n_ordem !== !b.n_ordem) return a.n_ordem ? -1 : 1; if (!a.n_ordem) return String(a.nome).localeCompare(String(b.nome)); var x = parseFloat(a.n_ordem), y = parseFloat(b.n_ordem); if (!isNaN(x) && !isNaN(y) && x !== y) return x - y; return String(a.n_ordem).localeCompare(String(b.n_ordem)); });
  if (res.novos || res.atualizados) escreverMembros_(M);
  if (hist.length) registarGraus_(hist);
  Object.keys(inscr).forEach(function (em) { acrescentar_('Inscritos', inscritoDeBase_(inscr[em], 'Importado sem nº da Ordem')); });
  res.inscritosMarcados = sincronizarInscritos_();
  delete res.plano;
  return res;
}
function apiAdminGuardarMembro(tok, m, chaveOriginal) {
  var quem = admin_(tok);
  var r = lock_(function () {
    if (!chaveOriginal) {
      var n = normNum_(m && m.n_ordem), ex = n ? ler_('Membros').filter(function (x) { return normNum_(x.n_ordem) === n; })[0] : null;
      if (ex) throw new Error('Já existe um membro com o nº ' + ex.n_ordem + ' (' + ex.nome + '). Use «Editar» para o alterar.');
      var em = String(m && m.email || '').trim().toLowerCase(), ex2 = em ? ler_('Membros').filter(function (x) { return x.email && x.email.toLowerCase() === em; })[0] : null;
      if (ex2) throw new Error('Já existe na base alguém com o email ' + em + ' (' + ex2.nome + '). Use «Editar» para o alterar.');
    }
    return juntarMembros_([{ linha: 1, dados: m || {} }], chaveOriginal || '', { quem: quem });
  });
  if (r.erros.length) throw new Error('Não foi guardado: ' + r.erros[0].motivo + '.');
  return (r.novos ? (ehConvidado_(limparMembro_(m || {})) ? 'Convidado adicionado.' : 'Membro adicionado.') : 'Dados atualizados.') + (r.inscritosMarcados ? ' Já estava inscrito nas notificações e passou a membro.' : '');
}
// Liga/desliga as notificações de pessoas da base de membros (chaves = nº da Ordem ou email; 'todos' = todos com email)
function apiAdminNotificar(tok, chaves, ligar) {
  admin_(tok);
  var org = getDef_('nome_org') || CONFIG.NOME_ORG, res = { ativados: 0, ja: 0, semEmail: 0, desativados: 0 };
  lock_(function () {
    var M = ler_('Membros'), alvo;
    if (chaves === 'todos') alvo = M;
    else { var set = {}; (chaves || []).forEach(function (c) { set[normChave_(c)] = 1; }); alvo = M.filter(function (m) { return set[chaveMembro_(m)]; }); }
    if (!alvo.length) throw new Error('Não encontrámos esse membro na base.');
    var I = ler_('Inscritos');
    alvo.forEach(function (m) {
      var email = String(m.email || '').toLowerCase();
      if (!emailOk_(email)) { res.semEmail++; return; }
      var s = I.filter(function (x) { return x.email === email; })[0];
      if (!ligar) { if (s) { apagar_('Inscritos', function (x) { return x.email === email; }); res.desativados++; } return; }
      if (s && s.estado === 'confirmado') { res.ja++; return; }
      if (s) { s.estado = 'confirmado'; s.confirmado_em = agora_(); if (!s.nome) s.nome = m.nome; if (!ehConvidado_(m)) { s.membro = 'SIM'; s.n_ordem = m.n_ordem; s.verificar = ''; } atualizar_('Inscritos', s); }
      else {
        s = { email: email, nome: m.nome, telemovel: m.telemovel || '', localidade: m.localidade || '', nota: '', estado: 'confirmado', membro: 'NÃO', token: novoToken_(), criado_em: agora_(), confirmado_em: agora_(), ultimo_envio_confirmacao: '', n_ordem: '', lingua: '' };
        identificar_(s, ehConvidado_(m) ? '' : m.n_ordem); acrescentar_('Inscritos', s);
      }
      enfileirar_(email, 'Passa a receber as notificações da ' + org, { linhas: [saud_(s.nome, s), 'A ' + org + ' ativou as notificações para o seu email. A partir de agora recebe os convites para os eventos, os lembretes e as novidades.', 'Pode ver o calendário e inscrever-se nos eventos no botão abaixo. Se não quiser receber estes emails, use o link no fim da mensagem.'], botao: { texto: 'Ver o calendário', url: url_({ t: s.token }) }, sair: url_({ sair: s.token }), porAdmin: true }, 3);
      try { convitesTardios_(s); } catch (e3) {}
      res.ativados++;
    });
  });
  enviarJa_();
  if (!ligar) return res.desativados ? 'Deixou de receber notificações.' : 'Não estava a receber notificações.';
  var p = [];
  if (res.ativados) p.push(res.ativados === 1 ? 'Notificações ativadas; foi enviado um email de boas-vindas.' : res.ativados + ' pessoas passaram a receber notificações (email de boas-vindas enviado).');
  if (res.ja) p.push(res.ja + ' já recebia' + (res.ja > 1 ? 'm' : '') + '.');
  if (res.semEmail) p.push(res.semEmail + ' sem email (não é possível notificar).');
  return p.join(' ') || 'Nada a ativar.';
}
function entradasImport_(linhas) {
  if (!linhas || !linhas.length) throw new Error('Não há linhas para importar.');
  if (linhas.length > 2000) throw new Error('O ficheiro tem demasiadas linhas (máximo 2000).');
  return linhas.map(function (l, i) { return { linha: l.linha || i + 2, dados: l.dados || l }; });
}
// Pré-visualização: o que a importação vai fazer, sem gravar
function apiAdminPreverImportacao(tok, linhas) {
  admin_(tok);
  var r = juntarMembros_(entradasImport_(linhas), '', { manterVazios: true, simular: true, paraInscritos: true });
  return { novos: r.novos, atualizados: r.atualizados, iguais: r.iguais, erros: r.erros.length, plano: r.plano, paraInscritos: r.paraInscritos };
}
function apiAdminImportarMembros(tok, linhas) {
  var quem = admin_(tok);
  return lock_(function () { return juntarMembros_(entradasImport_(linhas), '', { manterVazios: true, paraInscritos: true, quem: quem + ' (importação)' }); });
}
function apiAdminRemoverMembro(tok, num) {
  admin_(tok);
  lock_(function () {
    var n = normChave_(num), M = ler_('Membros'), fica = M.filter(function (m) { return chaveMembro_(m) !== n; });
    if (fica.length === M.length) throw new Error('Membro não encontrado.');
    escreverMembros_(fica);
  });
  return 'Membro removido da base. (Se estava inscrito nas notificações, continua inscrito; pode desmarcar «Membro» em Inscritos.)';
}

/* ============ Link do evento: inscrição direta ============ */
function apiEventoLink(id) {
  var e = ler_('Eventos').filter(function (x) { return x.id === String(id || ''); })[0];
  return e && e.estado === 'ativo' ? evPublico_(e) : null;
}
function apiLingua(token, lang) {
  lang = lang === 'en' ? 'en' : 'pt';
  lock_(function () { var s = ler_('Inscritos').filter(function (x) { return x.token === token; })[0]; if (s && s.lingua !== lang) { s.lingua = lang; atualizar_('Inscritos', s); } });
  return true;
}
function apiInscreverPublico(eventoId, p, f, lang) {
  lang = lang === 'en' ? 'en' : 'pt'; p = p || {}; f = f || {};
  var email = limpo_(p.email, 120).toLowerCase(), nome = limpo_(p.nome, 120), tel = limpo_(p.telemovel, 40), loc = limpo_(p.localidade, 80), num = limpo_(p.n_ordem, 30);
  if (!emailOk_(email)) throw new Error(tr_(lang, 'Indique um email válido.', 'Please enter a valid email.'));
  if (!nome || !tel || !loc) throw new Error(tr_(lang, 'Preencha o nome, o contacto e a localidade.', 'Please fill in your name, phone and town.'));
  lock_(function () {
    var e = ler_('Eventos').filter(function (x) { return x.id === eventoId; })[0];
    if (!e || e.estado !== 'ativo') throw new Error(tr_(lang, 'Este evento já não está disponível.', 'This event is no longer available.'));
    if (e.externo) throw new Error(tr_(lang, 'Este evento é organizado pela ' + e.externo + '; a inscrição não é feita aqui.', 'This event is organised by ' + e.externo + '; registration is not handled here.'));
    if (hoje_() > e.prazo) throw new Error(tr_(lang, 'As inscrições para este evento já fecharam.', 'Registration for this event has closed.'));
    var s = ler_('Inscritos').filter(function (x) { return x.email === email; })[0];
    var tmp = { email: email, nome: nome, membro: s ? s.membro : 'NÃO', n_ordem: s ? s.n_ordem : '' };
    identificar_(tmp, num);
    if (!podeVer_(e, tmp)) throw new Error(e.visibilidade === 'direcao' ? tr_(lang, 'Este evento é reservado à Direção. Use o email registado na Comenda.', 'This event is reserved for the board. Please use the email registered with the Comenda.') : tr_(lang, 'Este evento é reservado a membros. Use o email registado na Comenda ou indique o seu nº da Ordem.', 'This event is reserved for members. Please use the email registered with the Comenda or enter your Order number.'));
    if (ler_('Inscricoes').some(function (r) { return r.evento_id === e.id && r.email === email; })) throw new Error(tr_(lang, 'Este email já está inscrito neste evento. Para alterar, use o link pessoal que recebeu por email.', 'This email is already registered for this event. To make changes, use the personal link you received by email.'));
    var tok = novoToken_();
    acrescentar_('PreInscricoes', { token: tok, email: email, evento_id: e.id, dados: JSON.stringify({ p: { email: email, nome: nome, telemovel: tel, localidade: loc, n_ordem: num }, f: f, lang: lang }), criado_em: new Date().toISOString() });
    enfileirar_(email, tr_(lang, 'Confirme a sua inscrição: ', 'Confirm your registration: ') + titE_(e, lang), {
      lang: lang,
      linhas: [saud_(nome, lang), tr_(lang, 'Recebemos o seu pedido de inscrição em:', 'We received your registration request for:'), '**' + titE_(e, lang), resumo_(e, lang), tr_(lang, 'Para concluir, carregue no botão abaixo. Se não fez este pedido, ignore este email.', 'To complete it, click the button below. If you did not make this request, please ignore this email.')],
      botao: { texto: tr_(lang, 'Confirmar inscrição', 'Confirm registration'), url: url_({ insc: tok }) }
    }, 1);
  });
  enviarJa_();
  return tr_(lang, 'Quase feito! Enviámos um email para ' + email + '. Abra-o e carregue em «Confirmar inscrição» para concluir.', 'Almost done! We sent an email to ' + email + '. Open it and click «Confirm registration» to finish.');
}
function apiConfirmarInscricao(tok) {
  var out = lock_(function () {
    var pre = ler_('PreInscricoes').filter(function (x) { return x.token === String(tok || ''); })[0];
    if (!pre) return { ok: false, msg: 'Este link já não é válido ou já foi usado. / This link is no longer valid or has already been used.' };
    var d = jsonOu_(pre.dados, {}), p = d.p || {}, lang = d.lang === 'en' ? 'en' : 'pt';
    var s = ler_('Inscritos').filter(function (x) { return x.email === pre.email; })[0], av, novo = false;
    if (!s) {
      s = { email: pre.email, nome: p.nome, telemovel: p.telemovel, localidade: p.localidade, nota: '', estado: 'confirmado', membro: 'NÃO', token: novoToken_(), criado_em: agora_(), confirmado_em: agora_(), ultimo_envio_confirmacao: '', n_ordem: '', lingua: lang };
      av = identificar_(s, p.n_ordem); acrescentar_('Inscritos', s); novo = true;
      s = ler_('Inscritos').filter(function (x) { return x.email === pre.email; })[0];
    } else {
      if (s.estado !== 'confirmado') { s.estado = 'confirmado'; s.confirmado_em = agora_(); novo = true; }
      if (!s.lingua) s.lingua = lang;
      av = identificar_(s, p.n_ordem); atualizar_('Inscritos', s);
    }
    apagar_('PreInscricoes', function (x) { return x.token === pre.token; });
    if (av.length) avisarAdmins_('Verificar membro: ' + s.nome, av);
    if (novo) avisarAdmins_('Novo inscrito (pelo link do evento): ' + s.nome, [s.nome + ' (' + s.email + ') inscreveu-se pelo link de um evento e passou a receber notificações.', s.telemovel ? 'Contacto: ' + s.telemovel : '', s.localidade ? 'Localidade: ' + s.localidade : '']);
    try { var m1 = inscrever_(s, pre.evento_id, d.f || {}, true); if (novo) { try { convitesTardios_(s); } catch (e2) {} } return { ok: true, msg: m1, token: s.token, ev: pre.evento_id }; }
    catch (err) { return { ok: false, msg: String(err && err.message || err), token: s.token, ev: pre.evento_id }; }
  });
  enviarJa_();
  return out;
}

function apiSubscrever(f) {
  var lang = f.lang === 'en' ? 'en' : 'pt';
  var email = limpo_(f.email, 120).toLowerCase(), nome = limpo_(f.nome, 120), tel = limpo_(f.telemovel, 40), loc = limpo_(f.localidade, 80), nota = limpo_(f.nota, 1000), num = limpo_(f.n_ordem, 30);
  if (!emailOk_(email)) throw new Error(tr_(lang, 'Indique um email válido.', 'Please enter a valid email.'));
  if (!nome || !tel || !loc) throw new Error(tr_(lang, 'Preencha o nome, o contacto e a localidade.', 'Please fill in your name, phone and town.'));
  var r = lock_(function () {
    var s = ler_('Inscritos').filter(function (x) { return x.email === email; })[0];
    if (s && s.estado === 'confirmado') return 'ja';
    if (s && s.ultimo_envio_confirmacao && Date.now() - new Date(s.ultimo_envio_confirmacao).getTime() < 2 * 60000) return 'enviado';
    var av;
    if (!s) {
      s = { email: email, nome: nome, telemovel: tel, localidade: loc, nota: nota, estado: 'pendente', membro: 'NÃO', token: novoToken_(), criado_em: agora_(), confirmado_em: '', ultimo_envio_confirmacao: new Date().toISOString(), n_ordem: '', lingua: lang };
      av = identificar_(s, num);
      acrescentar_('Inscritos', s);
    } else {
      s.nome = nome; s.telemovel = tel; s.localidade = loc; s.nota = nota; s.lingua = lang; s.ultimo_envio_confirmacao = new Date().toISOString();
      av = identificar_(s, num);
      atualizar_('Inscritos', s);
    }
    if (av.length) avisarAdmins_('Verificar membro: ' + nome, av);
    var org = getDef_('nome_org') || CONFIG.NOME_ORG;
    enfileirar_(email, tr_(lang, 'Confirme a sua inscrição nas notificações', 'Please confirm your subscription'), {
      lang: lang,
      linhas: [saud_(nome, lang), tr_(lang, 'Recebemos um pedido para receber as notificações de eventos da ' + org + ' neste email.', 'We received a request to send event notifications from ' + org + ' to this email.'), tr_(lang, 'Para confirmar, carregue no botão abaixo. Se não fez este pedido, ignore esta mensagem.', 'To confirm, click the button below. If you did not make this request, please ignore this message.')],
      botao: { texto: tr_(lang, 'Confirmar inscrição', 'Confirm subscription'), url: url_({ confirmar: s.token }) }
    }, 1);
    return 'enviado';
  });
  enviarJa_();
  return r === 'ja' ? tr_(lang, 'Este email já está inscrito. Se perdeu o seu link pessoal, contacte a Comenda.', 'This email is already subscribed. If you lost your personal link, please contact the Comenda.') : tr_(lang, 'Enviámos um email para ' + email + '. Abra-o e carregue em «Confirmar inscrição».', 'We sent an email to ' + email + '. Open it and click «Confirm subscription».');
}

function apiInscrever(token, eventoId, f) {
  var res = lock_(function () {
    var s = ler_('Inscritos').filter(function (x) { return x.token === token && x.estado === 'confirmado'; })[0];
    if (!s) throw new Error(tr_(f && f.lang, 'O seu acesso não é válido. Abra o link do último email recebido.', 'Your access is not valid. Please open the link in the latest email you received.'));
    return inscrever_(s, eventoId, f, false);
  });
  enviarJa_();
  return res;
}
function inscrever_(s, eventoId, f, peloLink, porAdmin) {
  {
    var e = ler_('Eventos').filter(function (x) { return x.id === eventoId; })[0];
    if (!e || e.estado !== 'ativo') throw new Error(tr_(s, 'Este evento já não está disponível.', 'This event is no longer available.'));
    if (e.externo) throw new Error(tr_(s, 'Este evento é organizado pela ' + e.externo + '; a inscrição não é feita aqui.', 'This event is organised by ' + e.externo + '; registration is not handled here.'));
    if (!porAdmin && !podeVer_(e, s)) throw new Error(e.visibilidade === 'direcao' ? tr_(s, 'Este evento é reservado à Direção.', 'This event is reserved for the board.') : tr_(s, 'Este evento é reservado a membros.', 'This event is reserved for members.'));
    if (!porAdmin && hoje_() > e.prazo) throw new Error(tr_(s, 'As inscrições para este evento já fecharam.', 'Registration for this event has closed.'));
    var campos = jsonOu_(e.campos, {}), acts = jsonOu_(e.atividades, []);
    var nome = limpo_(f.nome, 120); if (!nome) throw new Error(tr_(s, 'Indique o seu nome.', 'Please enter your name.'));
    var filtraActs = function (a) { return (a || []).filter(function (x) { return acts.indexOf(x) >= 0; }); };
    var conv = campos.convidados ? (f.convidados || []).slice(0, 20).map(function (g) {
      return { nome: limpo_(g.nome, 120), email: limpo_(g.email, 120).toLowerCase(), telemovel: limpo_(g.telemovel, 40), agape: campos.agape && !!g.agape, restricoes: campos.restricoes ? limpo_(g.restricoes, 200) : '', atividades: filtraActs(g.atividades), visita: campos.visita && !!g.visita };
    }) : [];
    if (conv.some(function (g) { return !g.nome; })) throw new Error(tr_(s, 'Indique o nome de cada convidado.', 'Please enter the name of each guest.'));
    var regs = ler_('Inscricoes'), existente = regs.filter(function (r) { return r.evento_id === e.id && r.email === s.email; })[0];
    if (existente && peloLink) throw new Error(tr_(s, 'Já estava inscrito neste evento. Para alterar, use o link pessoal dos seus emails.', 'You were already registered for this event. To make changes, use the personal link in your emails.'));
    var motivo = limpo_(f.motivo, 500);
    if (existente && !motivo) throw new Error(tr_(s, 'Indique o motivo da alteração da inscrição.', 'Please give the reason for changing your registration.'));
    if (+e.lotacao) {
      var usados = pessoasEvento_(e.id);
      if (existente) usados -= 1 + (+existente.n_convidados || 0);
      if (usados + 1 + conv.length > +e.lotacao) throw new Error(tr_(s, 'Lotação esgotada: restam ' + Math.max(0, e.lotacao - usados) + ' lugares.', 'Not enough places left: ' + Math.max(0, e.lotacao - usados) + ' remaining.'));
    }
    var r = existente || { id: novoId_(), evento_id: e.id, email: s.email };
    r.evento = e.data + ' · ' + e.titulo; r.tipo = ehSim_(s.membro) ? 'Membro' : 'Inscrito não membro';
    r.nome = nome; r.telemovel = campos.telemovel ? limpo_(f.telemovel, 40) : '';
    r.agape = campos.agape ? sim_(!!f.agape) : ''; r.restricoes = campos.restricoes ? limpo_(f.restricoes, 200) : '';
    r.atividades = JSON.stringify(filtraActs(f.atividades)); r.observacoes = campos.observacoes ? limpo_(f.observacoes, 1000) : '';
    r.visita = campos.visita ? sim_(!!f.visita) : '';
    var nif = campos.nif ? String(f.nif || '').replace(/\s/g, '').slice(0, 20) : '';
    if (nif && !/^[A-Za-z0-9]{5,20}$/.test(nif)) throw new Error(tr_(s, 'O NIF indicado não é válido.', 'The tax number (NIF) is not valid.'));
    r.nif = nif;
    r.n_convidados = conv.length; r.atualizado_em = agora_();
    if (existente) r.motivo_alteracao = motivo;
    if (existente) atualizar_('Inscricoes', r); else acrescentar_('Inscricoes', r);
    apagar_('Recusas', function (x) { return x.evento_id === e.id && x.email === s.email; });
    apagar_('Convidados', function (g) { return g.inscricao_id === r.id; });
    conv.forEach(function (g) {
      acrescentar_('Convidados', { inscricao_id: r.id, evento_id: e.id, evento: r.evento, convidado_de: nome, nome: g.nome, email: g.email, telemovel: g.telemovel, agape: campos.agape ? sim_(g.agape) : '', restricoes: g.restricoes, atividades: JSON.stringify(g.atividades), visita: campos.visita ? sim_(g.visita) : '' });
    });
    s.nome = nome; if (r.telemovel) s.telemovel = r.telemovel; atualizar_('Inscritos', s);
    var sn = function (v) { return ehSim_(v) ? tr_(s, 'sim', 'yes') : tr_(s, 'não', 'no'); };
    var ref = refeicao_(e), refN = ref ? refNome_(ref, s) : tr_(s, 'Ágape', 'Lunch');
    var ga = conv.length ? tr_(s, 'Convidados: ', 'Guests: ') + conv.map(function (g) { return g.nome + (g.agape ? ' (' + refN.toLowerCase() + ')' : '') + (g.visita ? tr_(s, ' (visita)', ' (visit)') : ''); }).join(', ') : (campos.convidados ? tr_(s, 'Sem convidados.', 'No guests.') : '');
    var nRef = campos.agape ? (ehSim_(r.agape) ? 1 : 0) + conv.filter(function (g) { return g.agape; }).length : 0;
    enfileirar_(s.email, (existente ? tr_(s, 'Inscrição atualizada: ', 'Registration updated: ') : tr_(s, 'Inscrição confirmada: ', 'Registration confirmed: ')) + titE_(e, s), {
      lang: s.lingua,
      linhas: [saud_(nome, s), porAdmin ? tr_(s, 'A ' + (getDef_('nome_org') || CONFIG.NOME_ORG) + ' fez a sua inscrição no seguinte evento. Pode confirmar ou alterar os dados (refeição, convidados) no botão abaixo.', (getDef_('nome_org') || CONFIG.NOME_ORG) + ' has registered you for the following event. You can check or change the details (meal, guests) using the button below.') : tr_(s, 'Registámos a sua inscrição:', 'We have recorded your registration:'), '**' + titE_(e, s), quando_(e, s), campos.agape ? refN + ': ' + sn(r.agape) : '', campos.visita ? tr_(s, 'Visita: ', 'Visit: ') + sn(r.visita) : '', JSON.parse(r.atividades).length ? tr_(s, 'Atividades: ', 'Activities: ') + actsE_(e, JSON.parse(r.atividades), s).join('; ') : '', ga, calLigado_(s) ? '✓ ' + tr_(s, 'Este evento já está no seu calendário da Comenda.', 'This event is already in your Comenda calendar.') : ''],
      extra: detalhes_(e, s, pagamento_(e, s, nRef)),
      botao: { texto: tr_(s, 'Ver ou alterar inscrição', 'View or change registration'), url: url_({ t: s.token, ev: e.id }) }, sair: url_({ sair: s.token }),
      calendario: calLigado_(s) ? '' : urlCal_(s), ics: calLigado_(s) ? '' : conviteIcs_(e, s, false)
    }, 2);
    acrescentar_('Historico', { data: agora_(), email: s.email, nome: nome, evento_id: e.id, evento: r.evento, acao: porAdmin ? 'Inscrição pelo admin' : existente ? 'Alteração' : 'Inscrição', motivo: porAdmin ? 'Por ' + porAdmin : existente ? motivo : '' });
    if (!porAdmin) avisarAdmins_((existente ? 'Inscrição alterada: ' : 'Nova inscrição: ') + e.titulo, [nome + ' (' + s.email + ')' + (conv.length ? ' com ' + conv.length + ' convidado(s): ' + conv.map(function (g) { return g.nome; }).join(', ') : '') + '.', existente ? 'Motivo da alteração: ' + motivo : '', 'Total de pessoas no evento: ' + pessoasEvento_(e.id) + (+e.lotacao ? ' / ' + e.lotacao : '') + '.']);
    return existente ? tr_(s, 'Inscrição atualizada.', 'Registration updated.') : tr_(s, 'Inscrição registada. Vai receber um email de confirmação.', 'You are registered. You will receive a confirmation email.');
  }
}

function apiAnular(token, eventoId, motivo, lang) {
  motivo = limpo_(motivo, 500);
  if (!motivo) throw new Error(tr_(lang, 'Indique o motivo da anulação.', 'Please give the reason for cancelling.'));
  var x = lang;
  lock_(function () {
    var s = ler_('Inscritos').filter(function (y) { return y.token === token; })[0];
    if (!s) throw new Error(tr_(lang, 'O seu acesso não é válido.', 'Your access is not valid.'));
    x = s;
    var r = ler_('Inscricoes').filter(function (y) { return y.evento_id === eventoId && y.email === s.email; })[0];
    if (!r) return;
    apagar_('Convidados', function (g) { return g.inscricao_id === r.id; });
    apagar_('Inscricoes', function (y) { return y.id === r.id; });
    var ev = ler_('Eventos').filter(function (y) { return y.id === eventoId; })[0];
    if (ev) enfileirar_(s.email, tr_(s, 'Inscrição anulada: ', 'Registration cancelled: ') + titE_(ev, s), { lang: s.lingua, linhas: [saud_(s.nome, s), calLigado_(s) ? tr_(s, 'A sua inscrição no seguinte evento foi anulada. Vai ser retirado do seu calendário da Comenda automaticamente.', 'Your registration for the following event has been cancelled. It will be removed from your Comenda calendar automatically.') : tr_(s, 'A sua inscrição no seguinte evento foi anulada. O convite em anexo retira-o do seu calendário.', 'Your registration for the following event has been cancelled. The attached file removes it from your calendar.'), '**' + titE_(ev, s), resumo_(ev, s)], sair: url_({ sair: s.token }), ics: calLigado_(s) ? '' : conviteIcs_(ev, s, true) }, 2);
    acrescentar_('Historico', { data: agora_(), email: s.email, nome: r.nome, evento_id: eventoId, evento: r.evento, acao: 'Anulação', motivo: motivo });
    avisarAdmins_('Inscrição anulada: ' + r.evento, [r.nome + ' (' + s.email + ') anulou a inscrição' + (+r.n_convidados ? ' (e ' + r.n_convidados + ' convidado(s))' : '') + '.', 'Motivo: ' + motivo]);
  });
  enviarJa_();
  return tr_(x, 'Inscrição anulada.', 'Registration cancelled.');
}

function apiSair(token) {
  var lg = '';
  lock_(function () {
    var s = ler_('Inscritos').filter(function (x) { return x.token === token; })[0];
    if (!s) return;
    apagar_('Inscritos', function (x) { return x.token === token; });
    lg = s.lingua;
    avisarAdmins_('Cancelou a subscrição: ' + s.nome, [s.nome + ' (' + s.email + ') deixou de receber notificações.']);
  });
  enviarJa_();
  return tr_(lg, 'Deixou de receber as notificações. Pode voltar a inscrever-se quando quiser.', 'You will no longer receive notifications. You can subscribe again at any time.');
}

/* ============ API de administração ============ */

function hashPw_(pw, salt) {
  var h = salt + ':' + pw;
  for (var i = 0; i < 2000; i++) {
    h = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, h + salt, Utilities.Charset.UTF_8)
      .map(function (b) { return ('0' + (b & 255).toString(16)).slice(-2); }).join('');
  }
  return h;
}
function novaSessao_(email) {
  var tok = novoToken_();
  PropertiesService.getScriptProperties().setProperty('adm_' + tok, JSON.stringify({ email: email, exp: Date.now() + CONFIG.HORAS_SESSAO_ADMIN * 36e5 }));
  return tok;
}
function apiAdminLogin(email, pw) {
  email = limpo_(email, 120).toLowerCase(); pw = String(pw || '');
  var cache = CacheService.getScriptCache(), ck = 'falhas_' + email;
  var falhas = +(cache.get(ck) || 0);
  if (falhas >= 5) throw new Error('Demasiadas tentativas. Aguarde 15 minutos e tente de novo.');
  var props = PropertiesService.getScriptProperties();
  var reg = listaAdmins_().indexOf(email) >= 0 ? jsonOu_(props.getProperty('pw_' + email), null) : null;
  if (!reg || hashPw_(pw, reg.salt) !== reg.hash) {
    cache.put(ck, String(falhas + 1), 900);
    throw new Error(reg || listaAdmins_().indexOf(email) < 0 ? 'Email ou palavra-passe incorretos.' : 'Ainda não criou a sua palavra-passe. Use «Primeira vez ou esqueceu-se da palavra-passe?».');
  }
  cache.remove(ck);
  return novaSessao_(email);
}
function linkPassword_(email, novo, nome) {
  var tok = novoToken_(), org = getDef_('nome_org') || CONFIG.NOME_ORG;
  PropertiesService.getScriptProperties().setProperty('rst_' + tok, JSON.stringify({ email: email, exp: Date.now() + 2 * 36e5 }));
  enfileirar_(email, novo ? 'Foi nomeado administrador da ' + org : 'Criar ou alterar a palavra-passe de administrador', {
    linhas: novo ? [saud_(nome), 'Passou a ser administrador da app de eventos da ' + org + '.', 'Para entrar no painel, crie a sua palavra-passe no botão abaixo. O link é válido durante 2 horas; depois pode pedir outro em «Primeira vez ou esqueceu-se da palavra-passe?», na área do administrador.']
      : ['Recebemos um pedido para criar ou alterar a sua palavra-passe de administrador da ' + org + '.', 'O link é válido durante 2 horas. Se não fez este pedido, ignore este email: a palavra-passe atual continua válida.'],
    botao: { texto: novo ? 'Criar palavra-passe' : 'Definir palavra-passe', url: url_({ definir: tok }) }
  }, 1);
}
function apiAdminPedirPassword(email) {
  email = limpo_(email, 120).toLowerCase();
  if (listaAdmins_().indexOf(email) >= 0) {
    lock_(function () { linkPassword_(email, false); });
    enviarJa_();
  }
  return 'Se este email for de um administrador, vai receber um link para definir a palavra-passe dentro de momentos.';
}
function apiAdminDefinirPassword(resetTok, pw) {
  pw = String(pw || '');
  if (pw.length < CONFIG.MIN_PASSWORD) throw new Error('A palavra-passe tem de ter pelo menos ' + CONFIG.MIN_PASSWORD + ' caracteres.');
  var props = PropertiesService.getScriptProperties();
  var r = jsonOu_(resetTok && props.getProperty('rst_' + resetTok), null);
  if (!r || r.exp < Date.now() || listaAdmins_().indexOf(r.email) < 0) throw new Error('Este link já não é válido. Peça um novo na página de administração.');
  var salt = novoToken_();
  props.setProperty('pw_' + r.email, JSON.stringify({ salt: salt, hash: hashPw_(pw, salt), em: agora_() }));
  props.deleteProperty('rst_' + resetTok);
  return { sessao: novaSessao_(r.email), email: r.email };
}
function apiAdminSair(tok) {
  if (tok) PropertiesService.getScriptProperties().deleteProperty('adm_' + tok);
  return 'Sessão terminada.';
}
function admin_(tok) {
  var raw = tok && PropertiesService.getScriptProperties().getProperty('adm_' + tok);
  var s = jsonOu_(raw, null);
  if (!s || s.exp < Date.now() || listaAdmins_().indexOf(s.email) < 0) throw new Error('SESSAO: A sua sessão terminou. Entre de novo com o seu email e palavra-passe.');
  return s.email;
}
function apiAdminDados(tok) {
  var quem = admin_(tok);
  var conv = ler_('Convidados');
  var fila = ler_('Fila');
  var subs = ler_('Inscritos'), resp = {};
  ler_('Inscricoes').forEach(function (r) { resp[r.evento_id + '|' + r.email] = 1; });
  ler_('Recusas').forEach(function (r) { resp[r.evento_id + '|' + r.email] = 1; });
  return {
    quem: quem, hoje: hoje_(), org: getDef_('nome_org') || CONFIG.NOME_ORG, admins: listaAdmins_(), folha: ss_().getUrl(),
    fila: { pendentes: fila.filter(function (m) { return !m.enviado_em; }).length, erros: fila.filter(function (m) { return m.enviado_em === 'ERRO'; }).length, quota: MailApp.getRemainingDailyQuota() },
    eventos: ler_('Eventos').map(function (e) { var o = evPublico_(e); o.lembrete7 = e.lembrete7; o.lembrete1 = e.lembrete1; o.album = e.album; o.agradecimento = e.agradecimento; o.pasta_url = e.pasta_url || ''; o.refeicao = refeicao_(e); o.video = e.video_estado ? { estado: e.video_estado, url: e.video_url || '', info: infoVideo_(e), em: e.video_em || '' } : null; o.pasta_id = e.pasta_id ? 1 : 0; o.sem_resposta = e.externo || (e.estado !== 'ativo' && e.estado !== 'oculto') ? [] : destinatarios_(e, subs).filter(function (s) { return !resp[e.id + '|' + s.email]; }).map(function (s) { return { email: s.email, nome: s.nome }; }); return o; }),
    presencas: ler_('Presencas').map(presencaPub_), presencasDias: PRESENCAS_DIAS,
    graus: listaDef_('graus', GRAUS_PADRAO), cargos: listaDef_('cargos', CARGOS_PADRAO), histGraus: ler_('HistGraus').map(function (h) { return { data: h.data, n_ordem: h.n_ordem, nome: h.nome, campo: h.campo, antes: h.antes, depois: h.depois, por: h.por }; }),
    notificacoes: ler_('Notificacoes').slice(-200).reverse().map(function (r) { return { id: r.id, criado_em: r.criado_em, por: r.por, assunto: r.assunto, texto: r.texto, botao_texto: r.botao_texto, botao_url: r.botao_url, evento_id: r.evento_id, grupos: jsonOu_(r.grupos, []), pessoas: jsonOu_(r.pessoas, []), agendada_para: r.agendada_para, estado: r.estado, enviada_em: r.enviada_em, n: r.n }; }),
    recusas: ler_('Recusas').map(function (r) { return { evento_id: r.evento_id, email: r.email, nome: r.nome, nota: r.nota, criado_em: r.criado_em, talvez: ehTalvez_(r) }; }),
    appDisp: (function () { var dv = {}; try { ler_('Dispositivos').forEach(function (d) { if (!d.fcm || !d.email) return; var v = [d.visto_em, d.criado_em].filter(Boolean).sort().pop() || ''; var k = String(d.email).toLowerCase(); if (!dv[k] || v > dv[k]) dv[k] = v; }); } catch (e) {} return dv; })(),
    inscritos: ler_('Inscritos').map(function (s) { return { email: s.email, nome: s.nome, telemovel: s.telemovel, localidade: s.localidade, nota: s.nota, estado: s.estado, membro: ehSim_(s.membro), criado_em: s.criado_em, confirmado_em: s.confirmado_em || '', n_ordem: s.n_ordem, lingua: s.lingua || '', verificar: jsonOu_(s.verificar, null) }; }),
    historico: ler_('Historico').slice(-300).reverse(),
    membros: ler_('Membros').map(function (m) { var o = {}; CAMPOS_MEMBRO.forEach(function (k) { o[k] = m[k]; }); return o; }),
    inscricoes: ler_('Inscricoes').map(function (r) {
      return { id: r.id, evento_id: r.evento_id, email: r.email, nome: r.nome, telemovel: r.telemovel, agape: ehSim_(r.agape), restricoes: r.restricoes, atividades: jsonOu_(r.atividades, []), observacoes: r.observacoes, visita: ehSim_(r.visita), nif: r.nif, motivo: r.motivo_alteracao,
        convidados: conv.filter(function (g) { return g.inscricao_id === r.id; }).map(function (g) { return { nome: g.nome, email: g.email, telemovel: g.telemovel, agape: ehSim_(g.agape), restricoes: g.restricoes, atividades: jsonOu_(g.atividades, []), visita: ehSim_(g.visita) }; }) };
    })
  };
}
function apiAdminGuardarEvento(tok, d, avisar) {
  admin_(tok);
  var campos = d.campos || {};
  var ev = {
    titulo: limpo_(d.titulo, 150), data: limpo_(d.data, 10), hora: limpo_(d.hora, 5), local: limpo_(d.local, 200), descricao: limpo_(d.descricao, 3000),
    visibilidade: ['publico', 'membros', 'direcao'].indexOf(d.visibilidade) >= 0 ? d.visibilidade : 'membros', hora_fim: /^\d\d:\d\d$/.test(d.hora_fim || '') ? d.hora_fim : '', prazo: limpo_(d.prazo, 10), lotacao: +d.lotacao > 0 ? String(Math.floor(+d.lotacao)) : '',
    campos: JSON.stringify({ convidados: !!campos.convidados, agape: !!campos.agape || !!(!d.ext && d.refeicao && d.refeicao.tipo), restricoes: !!campos.restricoes, telemovel: !!campos.telemovel, observacoes: !!campos.observacoes, visita: !!campos.visita, nif: !!campos.nif }),
    atividades: JSON.stringify((d.atividades || []).map(function (a) { return limpo_(a, 120); }).filter(Boolean).slice(0, 15)),
    album: limpo_(d.album, 500),
    titulo_en: limpo_(d.titulo_en, 150), descricao_en: limpo_(d.descricao_en, 3000),
    atividades_en: JSON.stringify((d.atividades_en || []).map(function (a) { return limpo_(a, 120); }).slice(0, 15)),
    externo: d.ext ? limpo_(d.externo, 150) : '', externo_link: d.ext ? limpo_(d.externo_link, 300) : '',
    morada: limpo_(d.morada, 200), mapa: limpo_(d.mapa, 500), programa: limpo_(d.programa, 3000), programa_en: limpo_(d.programa_en, 3000),
    traje: limpo_(d.traje, 1000), traje_en: limpo_(d.traje_en, 1000), refeicao: d.ext ? '' : refeicaoLimpa_(d.refeicao)
  };
  if (ev.mapa && !/^https?:\/\/\S+$/.test(ev.mapa)) throw new Error('O link do mapa do local tem de começar por https:// (ou deixe-o vazio para ser criado sozinho).');
  if (!d.ext && !ev.traje) throw new Error('Indique o traje a usar: é obrigatório.');
  if (d.ext) {
    if (!ev.externo) throw new Error('Indique o nome da comenda que organiza o evento.');
    ev.prazo = ev.data; ev.lotacao = ''; ev.album = ''; ev.campos = '{}'; ev.atividades = '[]'; ev.atividades_en = '[]';
  }
  if (ev.album && !/^https:\/\/\S+$/.test(ev.album)) throw new Error('O link do álbum tem de começar por https:// (copie-o do Google Fotos em «Partilhar» → «Criar link»).');
  if (!ev.titulo || !/^\d{4}-\d\d-\d\d$/.test(ev.data) || !/^\d{4}-\d\d-\d\d$/.test(ev.prazo) || !ev.local) throw new Error('Preencha o título, a data, o local e o prazo.');
  if (ev.prazo > ev.data) throw new Error('O prazo de inscrição tem de ser antes ou no dia do evento.');
  if (ev.hora_fim && ev.hora_fim <= ev.hora) throw new Error('A hora de fim tem de ser depois da hora de início.');
  var msg = lock_(function () {
    var subs = ler_('Inscritos');
    if (d.id) {
      var e = ler_('Eventos').filter(function (x) { return x.id === d.id; })[0];
      if (!e) throw new Error('Evento não encontrado.');
      if (e.estado === 'rascunho') { Object.keys(ev).forEach(function (k) { e[k] = ev[k]; }); e.atualizado_em = agora_(); atualizar_('Eventos', e); return { msg: 'Rascunho guardado.', id: e.id, novo: false }; }
      if (e.estado === 'oculto' || ev.externo) avisar = false;
      if (e.data !== ev.data) { e.lembrete7 = ''; e.lembrete1 = ''; }
      var mudouCal = ['titulo', 'data', 'hora', 'hora_fim', 'local', 'morada', 'descricao', 'titulo_en', 'descricao_en'].some(function (k) { return String(e[k] || '') !== String(ev[k] || ''); });
      Object.keys(ev).forEach(function (k) { e[k] = ev[k]; });
      if (mudouCal) e.seq = String((+e.seq || 0) + 1);
      e.atualizado_em = agora_();
      atualizar_('Eventos', e);
      var n = 0, avisados = {};
      var inscritos = {}; if (mudouCal || avisar) ler_('Inscricoes').forEach(function (r) { if (r.evento_id === e.id) inscritos[r.email] = 1; });
      if (avisar) {
        var dest = destinatarios_(e, subs).filter(function (s) { return inscritos[s.email]; }); n = dest.length; // só quem está inscrito no evento
        dest.forEach(function (s) {
          var cal = mudouCal && inscritos[s.email];
          enfileirar_(s.email, tr_(s, 'Alteração: ', 'Change: ') + titE_(e, s), { lang: s.lingua, linhas: [saud_(s.nome, s), tr_(s, 'O seguinte evento foi alterado. Confirme os novos dados:', 'The following event has changed. Please check the new details:'), '**' + titE_(e, s), quando_(e, s), descE_(e, s), cal ? (calLigado_(s) ? tr_(s, 'O seu calendário da Comenda atualiza-se sozinho.', 'Your Comenda calendar updates automatically.') : tr_(s, 'O convite em anexo atualiza o evento no seu calendário.', 'The attached invitation updates the event in your calendar.')) : ''], extra: detalhes_(e, s), botao: { texto: tr_(s, 'Ver o evento', 'View the event'), url: url_({ t: s.token, ev: e.id }) }, sair: url_({ sair: s.token }), ics: cal && !calLigado_(s) ? conviteIcs_(e, s, false) : '' }, 3);
          avisados[s.email] = 1;
        });
      }
      var nc = mudouCal ? atualizarCalendarios_(e, false, avisados) : 0;
      return { msg: 'Evento guardado.' + (n ? ' Aviso enviado a ' + n + ' pessoa(s).' : '') + (nc ? ' Calendário atualizado para ' + nc + ' inscrito(s).' : ''), id: e.id, novo: false };
    }
    var novo = ev; novo.id = novoId_(); novo.estado = d.rascunho ? 'rascunho' : 'ativo'; novo.seq = '0'; novo.lembrete7 = ''; novo.lembrete1 = ''; novo.criado_em = agora_(); novo.atualizado_em = novo.criado_em;
    if (diasAte_(novo.data) < 7) novo.lembrete7 = 'n/a';
    acrescentar_('Eventos', novo);
    if (d.rascunho) return { msg: 'Guardado como rascunho. Ninguém o vê e não foi enviado nenhum convite. Use «Publicar» quando estiver pronto.', id: novo.id, novo: true, rascunho: true };
    convites_(novo, subs);
    return { msg: (novo.externo ? 'Evento criado. Email informativo a enviar a ' : 'Evento criado. Convite a enviar a ') + destinatarios_(novo, subs).length + ' pessoa(s).', id: novo.id, novo: true };
  });
  enviarJa_();
  return msg;
}
function refeicaoLimpa_(r) {
  if (!r || !r.tipo) return '';
  var o = { tipo: r.tipo === 'almoco' ? 'almoco' : 'jantar', nome: limpo_(r.nome, 80), hora: /^\d\d:\d\d$/.test(r.hora || '') ? r.hora : '', local: limpo_(r.local, 200), morada: limpo_(r.morada, 200), mapa: limpo_(r.mapa, 500),
    preco: num_(r.preco) ? String(num_(r.preco)) : '', preco_ate: /^\d{4}-\d\d-\d\d$/.test(r.preco_ate || '') ? r.preco_ate : '', preco2: num_(r.preco2) ? String(num_(r.preco2)) : '',
    iban: limpo_(r.iban, 50).toUpperCase().replace(/\s+/g, ' '), mbway: limpo_(r.mbway, 30), emails: limpo_(r.emails, 300), nota: limpo_(r.nota, 1000) };
  var nome = o.tipo === 'almoco' ? 'almoço' : 'jantar';
  if (!o.local || !o.hora) throw new Error('Indique a hora e o local do ' + nome + '.');
  if (o.mapa && !/^https?:\/\/\S+$/.test(o.mapa)) throw new Error('O link do mapa do ' + nome + ' tem de começar por https:// (ou deixe-o vazio para ser criado sozinho).');
  if (String(r.preco || '').trim() && !o.preco) throw new Error('O preço do ' + nome + ' não é válido.');
  if (o.preco2 && !o.preco_ate) throw new Error('Indique até que data vale o primeiro preço do ' + nome + '.');
  if (o.preco_ate && !o.preco2) throw new Error('Indique o preço do ' + nome + ' depois de ' + o.preco_ate.split('-').reverse().join('/') + '.');
  if (o.preco_ate && !o.preco) throw new Error('Indique o preço do ' + nome + '.');
  var em = o.emails ? o.emails.split(/[,;\s]+/).filter(Boolean) : [];
  var mau = em.filter(function (a) { return !emailOk_(a); });
  if (mau.length) throw new Error('Email do comprovativo inválido: ' + mau.join(', '));
  o.emails = em.join(', ');
  return JSON.stringify(o);
}
function apiAdminCancelar(tok, id) {
  admin_(tok);
  var n = lock_(function () {
    var e = ler_('Eventos').filter(function (x) { return x.id === id; })[0];
    if (!e) throw new Error('Evento não encontrado.');
    e.estado = 'cancelado'; e.seq = String((+e.seq || 0) + 1); e.atualizado_em = agora_(); atualizar_('Eventos', e);
    if (e.externo) return -1;
    var dest = destinatarios_(e), inscritos = {};
    ler_('Inscricoes').forEach(function (r) { if (r.evento_id === e.id) inscritos[r.email] = 1; });
    dest.forEach(function (s) { var cal = inscritos[s.email]; enfileirar_(s.email, tr_(s, 'Cancelado: ', 'Cancelled: ') + titE_(e, s), { lang: s.lingua, linhas: [saud_(s.nome, s), tr_(s, 'Informamos que o seguinte evento foi cancelado:', 'We regret to inform you that the following event has been cancelled:'), '**' + titE_(e, s), resumo_(e, s), tr_(s, 'Pedimos desculpa pelo incómodo.', 'We apologise for the inconvenience.'), cal ? (calLigado_(s) ? tr_(s, 'O evento vai ser retirado do seu calendário da Comenda automaticamente.', 'The event will be removed from your Comenda calendar automatically.') : tr_(s, 'O convite em anexo retira o evento do seu calendário.', 'The attached file removes the event from your calendar.')) : ''], sair: url_({ sair: s.token }), ics: cal && !calLigado_(s) ? conviteIcs_(e, s, true) : '' }, 3); });
    return dest.length;
  });
  enviarJa_();
  return n < 0 ? 'Evento cancelado. Como é de outra comenda, não foi enviado nenhum aviso.' : 'Evento cancelado. Aviso a enviar a ' + n + ' pessoa(s).';
}
/* ---- Notificações manuais (com ou sem evento, agora ou agendadas) ---- */
var NOTIF_GRUPOS_EV = { todos: 'Todos os convidados', inscritos: 'Inscritos', talvez: 'Talvez', sem: 'Sem resposta', nao: 'Não vão' };
var NOTIF_GRUPOS = { todos: 'Todos os que recebem notificações', membros: 'Membros', convidados: 'Convidados (sem nº)', direcao: 'Direção' };
function notifDados_(o) {
  o = o || {};
  var d = {
    assunto: limpo_(o.assunto, 150), texto: limpo_(o.texto, 5000), botao_texto: limpo_(o.botao_texto, 60), botao_url: limpo_(o.botao_url, 500),
    evento_id: limpo_(o.evento_id, 40), grupos: (Array.isArray(o.grupos) ? o.grupos : []).map(function (g) { return limpo_(g, 120); }).filter(Boolean).slice(0, 30),
    pessoas: (Array.isArray(o.pessoas) ? o.pessoas : []).map(function (p) { return limpo_(p, 120).toLowerCase(); }).filter(emailOk_).slice(0, 500),
    agendada_para: limpo_(o.agendada_para, 20).replace('T', ' ')
  };
  if (!d.assunto || !d.texto) throw new Error('Escreva o assunto e a mensagem.');
  if (d.botao_url && !/^https?:\/\/\S+$/.test(d.botao_url)) throw new Error('O link do botão tem de começar por http:// ou https://');
  if (d.botao_url && !d.botao_texto) d.botao_texto = 'Abrir';
  if (!d.grupos.length && !d.pessoas.length) throw new Error('Escolha pelo menos um grupo ou uma pessoa.');
  if (d.agendada_para && !/^\d{4}-\d{2}-\d{2} \d{2}:\d{2}$/.test(d.agendada_para)) throw new Error('Data e hora de envio inválidas.');
  if (d.evento_id) { var e = ler_('Eventos').filter(function (x) { return x.id === d.evento_id; })[0]; if (!e) throw new Error('Evento não encontrado.'); if (e.externo) throw new Error('Os eventos de outra comenda não têm convidados.'); }
  return d;
}
// Quem recebe: só quem está confirmado nas notificações (quem saiu já não está na lista)
function notifDest_(d) {
  var subs = ler_('Inscritos').filter(function (s) { return s.estado === 'confirmado'; }), sel = {}, G = {};
  d.grupos.forEach(function (g) { G[g] = 1; });
  if (d.evento_id) {
    var e = ler_('Eventos').filter(function (x) { return x.id === d.evento_id; })[0];
    if (!e) return [];
    var regs = {}, rec = recusasEvento_(e.id);
    ler_('Inscricoes').forEach(function (r) { if (r.evento_id === e.id) regs[r.email] = 1; });
    destinatarios_(e, subs).forEach(function (s) {
      var r = rec[s.email], tv = ehTalvez_(r);
      if (G.todos || (G.inscritos && regs[s.email]) || (G.talvez && !regs[s.email] && tv) || (G.nao && !regs[s.email] && r && !tv) || (G.sem && !regs[s.email] && !r)) sel[s.email] = s;
    });
  } else {
    var mb = null;
    subs.forEach(function (s) {
      var mem = ehSim_(s.membro) && !!s.n_ordem, ok = G.todos || (G.membros && mem) || (G.convidados && !mem) || (G.direcao && ehDirecao_(s));
      if (!ok && mem && Object.keys(G).some(function (g) { return /^(grau|cargo):/.test(g); })) {
        var p = membroPor_(s.email, s.n_ordem), m = p.num || p.email;
        if (m) ok = (m.grau && G['grau:' + m.grau]) || (m.cargo && G['cargo:' + m.cargo]);
      }
      if (ok) sel[s.email] = s;
    });
  }
  d.pessoas.forEach(function (em) { var s = subs.filter(function (x) { return x.email === em; })[0]; if (s) sel[em] = s; });
  return Object.keys(sel).map(function (k) { return sel[k]; }).sort(function (a, b) { return String(a.nome).localeCompare(String(b.nome)); });
}
// Quem tem avisos ativos e abre a app: guarda a data (no máx. 1 escrita por dia e telemóvel), para o painel mostrar «Inativos» ao fim de 60 dias sem uso
function marcarVistoApp_(email) {
  var hoje = hoje_(), D = ler_('Dispositivos').filter(function (d) { return d.email === email && String(d.visto_em || '').slice(0, 10) !== hoje; });
  if (!D.length) return;
  var l = LockService.getScriptLock(); if (!l.tryLock(3000)) return;
  try { D.forEach(function (d) { d.visto_em = agora_(); atualizar_('Dispositivos', d); }); } finally { l.releaseLock(); }
}
function apiAdminNotifPrever(tok, o) {
  admin_(tok);
  var d = notifDados_(o);
  return notifDest_(d).map(function (s) { return { nome: s.nome, email: s.email }; });
}
function notifEnfileirar_(d) {
  var e = d.evento_id ? ler_('Eventos').filter(function (x) { return x.id === d.evento_id; })[0] : null;
  var regs = {};
  if (e) ler_('Inscricoes').forEach(function (r) { if (r.evento_id === e.id) regs[r.email] = 1; });
  var dest = notifDest_(d);
  dest.forEach(function (s) {
    var linhas = [saud_(s.nome, s)].concat(d.texto.split(/\n+/));
    if (e) linhas = linhas.concat(['**' + titE_(e, s), resumo_(e, s)]);
    var botao = d.botao_url ? { texto: d.botao_texto, url: d.botao_url } : e ? { texto: regs[s.email] ? tr_(s, 'Ver ou alterar inscrição', 'View or change registration') : tr_(s, 'Ver o evento', 'View the event'), url: url_({ t: s.token, ev: e.id }) } : { texto: tr_(s, 'Abrir o calendário', 'Open the calendar'), url: url_({ t: s.token }) };
    enfileirar_(s.email, d.assunto, { lang: s.lingua, linhas: linhas, botao: botao, sair: url_({ sair: s.token }) }, 4);
  });
  return dest.length;
}
function apiAdminNotifEnviar(tok, o) {
  var quem = admin_(tok);
  var d = notifDados_(o), agora = agora_();
  var futura = d.agendada_para && d.agendada_para > agora;
  var res = lock_(function () {
    var n = futura ? 0 : notifEnfileirar_(d);
    if (!futura && !n) throw new Error('Nenhuma pessoa corresponde à escolha feita.');
    acrescentar_('Notificacoes', { id: novoId_(), criado_em: agora, por: quem, assunto: d.assunto, texto: d.texto, botao_texto: d.botao_texto, botao_url: d.botao_url, evento_id: d.evento_id, grupos: JSON.stringify(d.grupos), pessoas: JSON.stringify(d.pessoas), agendada_para: futura ? d.agendada_para : '', estado: futura ? 'agendada' : 'enviada', enviada_em: futura ? '' : agora, n: futura ? '' : n });
    return n;
  });
  if (futura) return 'Notificação agendada para ' + d.agendada_para.split(' ')[0].split('-').reverse().join('/') + ' às ' + d.agendada_para.split(' ')[1] + '.';
  enviarJa_();
  return 'Notificação a enviar a ' + res + ' pessoa(s).';
}
function apiAdminNotifCancelar(tok, id) {
  var quem = admin_(tok);
  lock_(function () {
    var r = ler_('Notificacoes').filter(function (x) { return x.id === id; })[0];
    if (!r) throw new Error('Notificação não encontrada.');
    if (r.estado !== 'agendada') throw new Error('Esta notificação já não está agendada.');
    r.estado = 'cancelada'; r.enviada_em = agora_() + ' · ' + quem; atualizar_('Notificacoes', r);
  });
  return 'Notificação agendada cancelada.';
}
// Corre na tarefa de 10 em 10 minutos (processarFila): envia as agendadas cuja hora já chegou
function notifAgendadas_() {
  var agora = agora_();
  if (!ler_('Notificacoes').some(function (r) { return r.estado === 'agendada' && r.agendada_para <= agora; })) return 0;
  return lock_(function () {
    var n = 0;
    ler_('Notificacoes').filter(function (r) { return r.estado === 'agendada' && r.agendada_para <= agora; }).forEach(function (r) {
      var d = { assunto: r.assunto, texto: r.texto, botao_texto: r.botao_texto, botao_url: r.botao_url, evento_id: r.evento_id, grupos: jsonOu_(r.grupos, []), pessoas: jsonOu_(r.pessoas, []) };
      var k = 0; try { k = notifEnfileirar_(d); } catch (err) {}
      r.estado = 'enviada'; r.enviada_em = agora_(); r.n = k; atualizar_('Notificacoes', r); n++;
    });
    return n;
  });
}
function apiAdminNovidade(tok, id, assunto, texto, soInscritos) {
  admin_(tok);
  assunto = limpo_(assunto, 150); texto = limpo_(texto, 3000);
  if (!assunto || !texto) throw new Error('Escreva o assunto e a mensagem.');
  var n = lock_(function () {
    var e = ler_('Eventos').filter(function (x) { return x.id === id; })[0];
    var subs = ler_('Inscritos');
    var regs = ler_('Inscricoes').filter(function (r) { return r.evento_id === id; }).map(function (r) { return r.email; });
    var dest = destinatarios_(e, subs).filter(function (s) { return !soInscritos || regs.indexOf(s.email) >= 0; });
    dest.forEach(function (s) {
      enfileirar_(s.email, assunto, { lang: s.lingua, linhas: [saud_(s.nome, s)].concat(texto.split(/\n+/)).concat(['**' + titE_(e, s), resumo_(e, s)]), botao: { texto: regs.indexOf(s.email) >= 0 ? tr_(s, 'Ver ou alterar inscrição', 'View or change registration') : tr_(s, 'Ver o evento', 'View the event'), url: url_({ t: s.token, ev: e.id }) }, sair: url_({ sair: s.token }) }, 4);
    });
    return dest.length;
  });
  enviarJa_();
  return 'Novidade a enviar a ' + n + ' pessoa(s).';
}
function apiAdminLembrete(tok, id, dias) {
  admin_(tok);
  lock_(function () { var e = ler_('Eventos').filter(function (x) { return x.id === id; })[0]; if (e && e.externo) throw new Error('Os eventos de outra comenda não têm lembretes.'); lembretes_(e, dias === 1 ? 1 : 7); });
  enviarJa_();
  return 'Lembrete a enviar.';
}
function apiAdminMembro(tok, email, membro) {
  admin_(tok);
  lock_(function () {
    var s = ler_('Inscritos').filter(function (x) { return x.email === email; })[0];
    if (s && membro && !s.n_ordem && !(jsonOu_(s.verificar, null) || {}).n_ordem) throw new Error('Sem nº da Ordem: esta pessoa fica como convidado. Use «Editar» para indicar o nº da Ordem.');
    if (s) { var v = jsonOu_(s.verificar, null); s.membro = sim_(membro); if (membro && v && v.n_ordem) s.n_ordem = v.n_ordem; s.verificar = ''; atualizar_('Inscritos', s); }
  });
  return membro ? 'Marcado como membro.' : 'Deixou de ser membro.';
}
function apiAdminRemover(tok, email) {
  admin_(tok);
  lock_(function () { apagar_('Inscritos', function (x) { return x.email === email; }); });
  return 'Removido da lista.';
}
function apiAdminConvidar(tok, nome, email, membro) {
  admin_(tok);
  email = limpo_(email, 120).toLowerCase(); nome = limpo_(nome, 120);
  if (!emailOk_(email) || !nome) throw new Error('Indique o nome e um email válido.');
  lock_(function () {
    var s = ler_('Inscritos').filter(function (x) { return x.email === email; })[0];
    if (s && s.estado === 'confirmado') throw new Error('Este email já está inscrito.');
    if (!s) { s = { email: email, nome: nome, telemovel: '', localidade: '', nota: '', estado: 'pendente', membro: 'NÃO', token: novoToken_(), criado_em: agora_(), confirmado_em: '', ultimo_envio_confirmacao: new Date().toISOString(), n_ordem: '' }; identificar_(s, ''); acrescentar_('Inscritos', s); }
    enfileirar_(email, 'Convite para receber as notificações da ' + (getDef_('nome_org') || CONFIG.NOME_ORG), { linhas: [saud_(nome), 'Foi convidado(a) para receber os convites dos eventos da ' + (getDef_('nome_org') || CONFIG.NOME_ORG) + '.', 'Para aceitar, confirme o seu email no botão abaixo.'], botao: { texto: 'Confirmar inscrição', url: url_({ confirmar: s.token }) } }, 1);
  });
  enviarJa_();
  return 'Convite enviado para ' + email + '.';
}
function apiAdminReenviar(tok, email) {
  admin_(tok);
  lock_(function () {
    var s = ler_('Inscritos').filter(function (x) { return x.email === email; })[0];
    if (!s) throw new Error('Email não encontrado.');
    enfileirar_(email, tr_(s, 'Confirme a sua inscrição nas notificações', 'Please confirm your subscription'), { lang: s.lingua, linhas: [saud_(s.nome, s), tr_(s, 'Para começar a receber os convites dos eventos, confirme o seu email no botão abaixo.', 'To start receiving event invitations, please confirm your email with the button below.')], botao: { texto: tr_(s, 'Confirmar inscrição', 'Confirm subscription'), url: url_({ confirmar: s.token }) } }, 1);
  });
  enviarJa_();
  return 'Email de confirmação reenviado.';
}
function apiAdminDefinicoes(tok, org, admins, graus, cargos) {
  var quem = admin_(tok);
  var lista = String(admins || '').split(/[\s,;]+/).map(function (x) { return x.toLowerCase().trim(); }).filter(emailOk_);
  if (!lista.length) throw new Error('Indique pelo menos um administrador.');
  if (lista.indexOf(quem) < 0) throw new Error('Não se pode remover a si próprio da lista de administradores.');
  var limpa = function (t) { var v = {}; return String(t || '').split('\n').map(function (x) { return limpo_(x, 150); }).filter(function (x) { if (!x || v[x]) return false; v[x] = 1; return true; }); };
  lock_(function () {
    setDef_('nome_org', limpo_(org, 100) || CONFIG.NOME_ORG); setDef_('admins', lista.join(', '));
    if (graus != null) { var g = limpa(graus); if (!g.length) throw new Error('Indique pelo menos um grau.'); setDef_('graus', g.join('\n')); }
    if (cargos != null) setDef_('cargos', limpa(cargos).join('\n'));
  });
  return 'Definições guardadas.';
}

function apiAdminTornarAdmin(tok, email, sim) {
  var quem = admin_(tok);
  email = limpo_(email, 120).toLowerCase();
  if (!emailOk_(email)) throw new Error('Este membro não tem um email válido. Adicione o email na ficha do membro.');
  var msg = lock_(function () {
    var lista = listaAdmins_();
    if (sim) {
      if (lista.indexOf(email) >= 0) return 'Já é administrador.';
      lista.push(email); setDef_('admins', lista.join(', ')); _dir = null;
      var m = ler_('Membros').filter(function (x) { return x.email && x.email.toLowerCase() === email; })[0];
      linkPassword_(email, true, m ? m.nome : '');
      return 'Passou a administrador. Foi-lhe enviado um email para criar a palavra-passe.';
    }
    if (email === quem) throw new Error('Não se pode retirar a si próprio de administrador.');
    if (lista.indexOf(email) < 0) return 'Não era administrador.';
    lista = lista.filter(function (x) { return x !== email; });
    if (!lista.length) throw new Error('Tem de ficar pelo menos um administrador.');
    setDef_('admins', lista.join(', ')); _dir = null;
    PropertiesService.getScriptProperties().deleteProperty('pw_' + email);
    return 'Deixou de ser administrador e já não consegue entrar no painel.';
  });
  enviarJa_();
  return msg;
}
function evento_(id) { var e = ler_('Eventos').filter(function (x) { return x.id === id; })[0]; if (!e) throw new Error('Evento não encontrado.'); return e; }
function apiAdminPublicar(tok, id) {
  admin_(tok);
  var msg = lock_(function () {
    var e = evento_(id);
    if (e.estado === 'oculto') { e.estado = 'ativo'; e.atualizado_em = agora_(); atualizar_('Eventos', e); return 'O evento voltou a estar visível.'; }
    if (e.estado !== 'rascunho') return 'O evento já estava publicado.';
    e.estado = 'ativo'; e.lembrete7 = diasAte_(e.data) < 7 ? 'n/a' : ''; e.lembrete1 = ''; e.atualizado_em = agora_(); atualizar_('Eventos', e);
    var subs = ler_('Inscritos'); convites_(e, subs);
    return (e.externo ? 'Evento publicado. Email informativo a enviar a ' : 'Evento publicado. Convite a enviar a ') + destinatarios_(e, subs).length + ' pessoa(s).';
  });
  enviarJa_();
  return msg;
}
function apiAdminOcultar(tok, id) {
  admin_(tok);
  return lock_(function () {
    var e = evento_(id);
    if (e.estado !== 'ativo') throw new Error('Só é possível ocultar eventos publicados.');
    e.estado = 'oculto'; e.atualizado_em = agora_(); atualizar_('Eventos', e);
    return 'Evento ocultado. Deixou de aparecer no calendário (não foi enviado nenhum email).';
  });
}
function apiAdminEliminar(tok, id) {
  admin_(tok);
  var n = lock_(function () {
    var e = evento_(id), avisados = 0;
    if ((e.estado === 'ativo' || e.estado === 'oculto') && e.data >= hoje_()) {
      e.seq = String((+e.seq || 0) + 1); e.estado = 'cancelado';
      avisados = atualizarCalendarios_(e, true);
    }
    apagar_('Convidados', function (g) { return g.evento_id === id; });
    apagar_('Inscricoes', function (r) { return r.evento_id === id; });
    apagar_('Recusas', function (r) { return r.evento_id === id; });
    apagar_('Presencas', function (r) { return r.evento_id === id; });
    apagar_('Eventos', function (x) { return x.id === id; });
    return avisados;
  });
  enviarJa_();
  return 'Evento eliminado.' + (n ? ' Aviso de cancelamento a enviar a ' + n + ' inscrito(s).' : '');
}

/* Selo OSMTH (PNG, base64) usado no cabeçalho dos emails */
var SELO_PNG = '/9j/4AAQSkZJRgABAQAAAQABAAD/2wBDAAYEBQYFBAYGBQYHBwYIChAKCgkJChQODwwQFxQYGBcUFhYaHSUfGhsjHBYWICwgIyYnKSopGR8tMC0oMCUoKSj/2wBDAQcHBwoIChMKChMoGhYaKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCj/wAARCABYAFgDASIAAhEBAxEB/8QAHwAAAQUBAQEBAQEAAAAAAAAAAAECAwQFBgcICQoL/8QAtRAAAgEDAwIEAwUFBAQAAAF9AQIDAAQRBRIhMUEGE1FhByJxFDKBkaEII0KxwRVS0fAkM2JyggkKFhcYGRolJicoKSo0NTY3ODk6Q0RFRkdISUpTVFVWV1hZWmNkZWZnaGlqc3R1dnd4eXqDhIWGh4iJipKTlJWWl5iZmqKjpKWmp6ipqrKztLW2t7i5usLDxMXGx8jJytLT1NXW19jZ2uHi4+Tl5ufo6erx8vP09fb3+Pn6/8QAHwEAAwEBAQEBAQEBAQAAAAAAAAECAwQFBgcICQoL/8QAtREAAgECBAQDBAcFBAQAAQJ3AAECAxEEBSExBhJBUQdhcRMiMoEIFEKRobHBCSMzUvAVYnLRChYkNOEl8RcYGRomJygpKjU2Nzg5OkNERUZHSElKU1RVVldYWVpjZGVmZ2hpanN0dXZ3eHl6goOEhYaHiImKkpOUlZaXmJmaoqOkpaanqKmqsrO0tba3uLm6wsPExcbHyMnK0tPU1dbX2Nna4uPk5ebn6Onq8vP09fb3+Pn6/9oADAMBAAIRAxEAPwD6poooNACOyopZ2CqOpJwBWLd+J9Otonk3vIqNtYRrluJRE3y9flYjPHAIPeub1a81rUtUithp80NxatKHhjnEltdRMBgM2AUYgEqxXAbIyeRWxYeE4hufVZmvJGBDg8LJlDGS/wDeLJsDD7pKA4BoAH8WRyakLC1t8zfahbM0koVVBVzvGN2eYnXHHzDBp3hrXLrVHSGa1kiBtEl84jPzlV3BuBg5J4wM4yO4HE+Lfitp3gXX7jR7jw7dIyASLJC0YWZW6OO/qDnuDWP/AMNE6R/0A9Q/7+x1m6sVpc9Knk+MqxU4U7p+n+Z6BoniSVLx4L2We8Mk0wjKQqPLiSYxB2IxnpuPHA59M7+ja9p2sDFlcAy7Q5hdSkgBVWBKnBxhlOf9oV5x4S+Kum+O9eh0aDw5cyGRHMkk7RskceMMT7HgY75Arrdb8MNJeW/9lp5UcjMbl2mYfLuV2VcfMC5RAWB4VMAdKqMlLVHLiMLVwsuSsrM66iuN0G51ePW1sLuSeYwRbWLr8uwNzKxznLHKoMk4Qs3JwOqs722vVla0njmEUjQuUOdrqcMp9wao5yxRRRQAVyXiq+urnULbSbJLhDKw/ex5Qkg5yrlSpC4O4HDDgrkjB3tdvGsdKuZoziYIVi+Rn+c8L8qgk8+lY3gXSP7OtJ5pVQTyuVJUoxYKSMsyBQ5J3HcVDdjyKAOgsrcwQRLK/mzrGqPMwG58dyfz/OrFFFAHhX7U+hrNo2la3Gn7y3mNrIR1KOCVz9GU/wDfVfNlfZ3xvsRf/C7XkIy0UInX2KMG/kDXxl3rjrK0j9C4ZrOpg+R/ZbX6/qfSX7LGhrDo2q63Ig824mFrGT1CIAWx9WYf9817rXDfBGxFh8LtBQLhpYTO3uXYt/Iiu5rppq0UfG5rWdfGVJvvb7tCrqdkuoWMtq808KyDBeByjgZ5ww5GenFcd4flk0zxANNd1tYY90NrYLIEghjypwGIBlkI+YBRhAdpIzk93XM+NNPV7U38KxrdRhYzIYnlbbuDKERSNzb9uFJCnJzkcVZ5501FUdEu2vNNhklI88DZMAMbZBwwxk45z3ooA534iXNuIbKzuZoo45WeVxOImiZFGDvR2UuAXU4U5yARyBXJfGXxbqPgXTvDU/h6S2WIySRPD5YMMiqq8YHTnpgjFdX4yn8/UrG2gnUfvCkqk5UY2O25fKfICkZO5QNw55480/akt4bXQPDkFtFHDCk8wWONQqqNo6AdKzqtqLaPTyelCrjYQqK6d/yZf8O/tB6Ncoqa9p11Yy95IMTR/Xsw/I13enfFDwXqAUw+IbJCe07GE/8Aj4FfFVLzXOq8kfW1uGMJUd4Nx/Ffj/mfa3jHXND1PwXrsFvq+nTmWwnVRHdI2SY2x0NfFHb8KPwFFTOfOd+V5WsujKMZXufa/g/XdD0zwZoUFzq+mwGKwgVhJdIuCI1z1NM1L4peC9PVjN4gs5CO1uTMT/3wDXxX+VLzV+3dtEeX/qtRlNynUbv2sv8AM+mPEX7QekW6Omg6bdXsvaS4Ihj+uOWP5CtT4KeK9W8eWPiSbW5o8eYkUUUcYCQqyNwAc5/HNfKVfRf7LfnDQPEn2YRmbz4tgkJCk7T1xThUlKauZZrlOFwWBnKlHXTV6vdf1oeofDyYLZXNmTtaJzIIRsCxozMVwqsdgwPutg8ZIyTRSeG7GcazcyXMN9bKjmVFku2cSliRuKoojIxjHLHpnHFFdR8SWPFUdhp5/tKezhkklHkSyT3DRRJGcFievJ2KOFy3yjpXDfHPwxq/jew8NwaFZs7NI8sjS/u1hUov3yen06+1ema9p0uo2ojivru12ncRbvsMnopYfMB/ukHnrWV4CuAdM+xlZVMOWUSRmLAYk7VVmL4Xpub73UZFTKPMrM6MJiZYWqq0N1/keXeHf2eLKNUfxDq888neKzURqPbc2SfyFd3p3wg8EWKjGiRzsP4riV5CfwJx+ld9RSVOK6HTWzbGVneVR/LT8jhPFnhHwxpPg7XLm20DSYXhsZ3V1tEypEbYIOM5zXxl2/CvtD423osPhfr8hODJCIB7l2C/yJr4w71z17XSR9Xwu5zoznNt3dtfJf8ABPszwn4R8L6t4N0O5udA0mZ5rGB2drRNzExjJJxnOaj1H4QeCL5TnRUt3P8AFbyvHj8AcfpU3wRvhf8Awu0Fw2WihMDexRiv8gK7muhRjJJtHydfE4jD15xjNqzfV9zwbxF+zxZyK7+HtYmgk6iK8USKfbcuCPyNbPwP8Nap4G07xNH4htxAEdJVlVt0ciKjZKt/kj0r2CuX8bak0K2un2yySXFw4do4pDHI0anJCEOhznA+U5A5wRSVKKd0XVzfFVqDw9WXMnbffT+upV+HVqES+u0MBjnZEUwtEwbbksx8olScseeCccj1K6DQtN/s+zVZJZp52GWkn2tJjJO0soG4Ak8nJ9zRWh5hpVxw0a70y4je1toBbQ3El7c3UeZLq6J3HZtx1OQCSx4UADkY7Gg0ANjYOisucEZ5GKdWGyXGivLJCr3Ni772VVzJHnJZv9rnn860rW/trqaSKGTMsYUuhBBXcMjP4GgDxf8Aam1tYNB0vRY3/e3U5uJAOoRBgZ+rN/47XzTX1d4++Dz+M/Ek+rXviKWLcqxxQLahhEg6KDu55JOfUmuc/wCGcrb/AKGWb/wDH/xdctSE5SvY+3ynNcDgsLGlKeu70e7+XyLX7LOtrPoWq6LIw822mFzGD1KOMHH0Zf8Ax6vc68k8A/B1vBniSDVrLxFLLtVo5YWtQolQ9VJ3ccgH6gV6bf6nb2QIkLPIFLFEGSAOpPYD6kVtSTUbM+bzepQrYqVXDyupa7Na9d/vLrMFUsc4HPAzXLWGjy3OsJqN8JpoZGFxEshMZjbB2iSLJUlcgKwAYfxZIzV2GG+1gO2or9msmwUhRsSZByGLjkHI7e1btaHmBRRRQAUUUUAFZmo6NBdt5sZNtcZB82Lhjzzn64xnriiigCJLHUrZ18nUTLHkDbNGDgZ55+mfxxT5IdWMwMdxCqCQkgjIKbwQBxkHbuB5PP6FFADI9P1KScPd6kpjDEiKOLaCOeDzzwf89at2umW1uQwQPLtVDI4BYgDAz+FFFAF2iiigAooooA//2Q==';


/* ============ Entrar na app com código por email (app instalada / PWA) ============ */
/* No iPhone, a app instalada no ecrã principal não partilha o acesso guardado no Safari.
 * O membro indica o email, recebe um código de 6 dígitos (válido 10 minutos) e a app
 * fica com o seu acesso pessoal (o mesmo token dos links dos emails). */

function codigoEntrar_() {
  var n = parseInt(Utilities.getUuid().replace(/-/g, '').slice(0, 10), 16) % 1000000;
  return ('000000' + n).slice(-6);
}

function apiEntrarPedirCodigo(email, lang) {
  lang = lang === 'en' ? 'en' : 'pt';
  email = limpo_(email, 120).toLowerCase();
  if (!emailOk_(email)) throw new Error(tr_(lang, 'Indique um email válido.', 'Please enter a valid email address.'));
  var cache = CacheService.getScriptCache(), kp = 'entrar_pedidos_' + email;
  var pedidos = +(cache.get(kp) || 0);
  if (pedidos >= 5) throw new Error(tr_(lang, 'Demasiados pedidos. Aguarde 15 minutos e tente de novo.', 'Too many requests. Please wait 15 minutes and try again.'));
  cache.put(kp, String(pedidos + 1), 900);
  var s = ler_('Inscritos').filter(function (x) { return x.email === email && x.estado === 'confirmado'; })[0];
  if (s) {
    var cod = codigoEntrar_();
    cache.put('entrar_cod_' + email, JSON.stringify({ c: cod, n: 0 }), 600);
    var L = s.lingua === 'en' || s.lingua === 'pt' ? s.lingua : lang;
    lock_(function () {
      enfileirar_(email, tr_(L, 'O seu código para entrar na app: ', 'Your code to sign in to the app: ') + cod, {
        lang: L,
        linhas: [saud_(s.nome, L),
          tr_(L, 'Use este código para entrar na app da ' + (getDef_('nome_org') || CONFIG.NOME_ORG) + ':', 'Use this code to sign in to the ' + (getDef_('nome_org') || CONFIG.NOME_ORG) + ' app:'),
          '**' + cod,
          tr_(L, 'O código é válido durante 10 minutos. Se não fez este pedido, ignore este email.', 'The code is valid for 10 minutes. If you did not request it, please ignore this email.')]
      }, 1);
    });
    enviarJa_();
  }
  return tr_(lang,
    'Se ' + email + ' estiver inscrito nas notificações, vai receber um email com o código dentro de momentos. Veja também a pasta de spam.',
    'If ' + email + ' is subscribed to the notifications, you will receive an email with the code in a moment. Please also check your spam folder.');
}

function apiEntrarCodigo(email, codigo, lang) {
  lang = lang === 'en' ? 'en' : 'pt';
  email = limpo_(email, 120).toLowerCase();
  codigo = String(codigo || '').replace(/\D/g, '');
  var cache = CacheService.getScriptCache(), k = 'entrar_cod_' + email;
  var r = jsonOu_(cache.get(k), null);
  if (!r) throw new Error(tr_(lang, 'O código expirou ou já foi usado. Peça um código novo.', 'The code has expired or was already used. Please request a new code.'));
  if (r.c !== codigo) {
    r.n = (r.n || 0) + 1;
    if (r.n >= 5) { cache.remove(k); throw new Error(tr_(lang, 'Código errado demasiadas vezes. Peça um código novo.', 'Wrong code too many times. Please request a new code.')); }
    cache.put(k, JSON.stringify(r), 600);
    throw new Error(tr_(lang, 'Código incorreto. Confirme os 6 dígitos do email.', 'Incorrect code. Please check the 6 digits in the email.'));
  }
  cache.remove(k);
  var s = ler_('Inscritos').filter(function (x) { return x.email === email && x.estado === 'confirmado'; })[0];
  if (!s) throw new Error(tr_(lang, 'Este email já não está inscrito nas notificações.', 'This email is no longer subscribed to the notifications.'));
  return { token: s.token, lingua: s.lingua || '' };
}


/* ============ Notificações push (Firebase Cloud Messaging) ============ */
/* Os avisos de prioridade 3 e 4 (convites, lembretes, alterações, cancelamentos, novidades,
 * agradecimentos, vídeos) são também enviados como notificação para os telemóveis em que
 * o membro ativou os avisos. O email continua a ser enviado como sempre.
 * Configuração (Definições do projeto → Propriedades do script):
 *   FIREBASE_WEB  configuração da app Web do Firebase (JSON com apiKey, projectId, messagingSenderId, appId)
 *   FCM_VAPID     chave pública do par de chaves «Web Push certificates»
 *   FCM_SA        conteúdo do ficheiro JSON da conta de serviço do Firebase
 * Sem estas propriedades, nada muda: a app simplesmente não mostra o botão de ativar avisos. */

TABELAS.Dispositivos = ['fcm', 'email', 'plataforma', 'criado_em', 'visto_em'];
var PUSH_PEND_ = [];

function pushCfg_() {
  var p = PropertiesService.getScriptProperties();
  var web = {}, txt = String(p.getProperty('FIREBASE_WEB') || ''), m, re = /["']?(apiKey|authDomain|projectId|messagingSenderId|appId)["']?\s*:\s*["']([^"']+)["']/g;
  while ((m = re.exec(txt))) web[m[1]] = m[2]; // aceita o JSON ou o bloco «firebaseConfig» copiado da consola
  var vapid = String(p.getProperty('FCM_VAPID') || '').replace(/\s+/g, ''), sa = jsonOu_(p.getProperty('FCM_SA'), null);
  if (!web.projectId || !web.apiKey || !web.appId || !web.messagingSenderId || !vapid || !sa || !sa.private_key || !sa.client_email) return null;
  return { web: web, vapid: vapid, sa: sa };
}

/* chamado no fim de enfileirar_ */
function pushPendente_(para, assunto, o, prioridade) {
  if ((prioridade || 4) < 3 || !para) return;
  var linhas = (o && o.linhas || []).filter(Boolean).map(function (l) { return String(l).replace(/\*\*/g, ''); });
  var corpo = linhas.slice(1).reduce(function (a, l) { return a ? a + (/:$/.test(a) ? ' ' : ' · ') + l : l; }, '').replace(/\s+/g, ' ').trim();
  if (corpo.length > 180) corpo = corpo.slice(0, 177) + '…';
  PUSH_PEND_.push({ email: String(para).toLowerCase(), titulo: String(assunto || '').slice(0, 120), corpo: corpo, url: (o && o.botao && o.botao.url) || siteUrl_() || '' });
}

function fcmAcesso_(sa) {
  var cache = CacheService.getScriptCache(), k = 'fcm_oauth';
  var t = cache.get(k);
  if (t) return t;
  var b64 = function (s) { return Utilities.base64EncodeWebSafe(s).replace(/=+$/, ''); };
  var agora = Math.floor(Date.now() / 1000);
  var cab = b64(JSON.stringify({ alg: 'RS256', typ: 'JWT' }));
  var corpo = b64(JSON.stringify({ iss: sa.client_email, scope: 'https://www.googleapis.com/auth/firebase.messaging', aud: 'https://oauth2.googleapis.com/token', iat: agora, exp: agora + 3600 }));
  var assin = Utilities.base64EncodeWebSafe(Utilities.computeRsaSha256Signature(cab + '.' + corpo, sa.private_key)).replace(/=+$/, '');
  var r = UrlFetchApp.fetch('https://oauth2.googleapis.com/token', { method: 'post', muteHttpExceptions: true, payload: { grant_type: 'urn:ietf:params:oauth:grant-type:jwt-bearer', assertion: cab + '.' + corpo + '.' + assin } });
  var j = jsonOu_(r.getContentText(), {});
  if (!j.access_token) throw new Error('FCM: não foi possível obter acesso (' + r.getResponseCode() + ') ' + r.getContentText().slice(0, 200));
  cache.put(k, j.access_token, 3000);
  return j.access_token;
}

/* chamado em enviarJa_ e processarFila: envia o que ficou pendente nesta execução */
function enviarPushPendentes_() {
  if (!PUSH_PEND_.length) return 0;
  var lista = PUSH_PEND_.splice(0, PUSH_PEND_.length);
  var enviados = 0;
  try {
    var cfg = pushCfg_();
    if (!cfg) return 0;
    var porEmail = {};
    ler_('Dispositivos').forEach(function (d) { if (d.fcm) (porEmail[d.email] = porEmail[d.email] || []).push(d); });
    var pedidos = [], alvo = [];
    lista.forEach(function (p) {
      (porEmail[p.email] || []).forEach(function (d) {
        alvo.push(d);
        pedidos.push({ msg: { message: { token: d.fcm, data: { titulo: p.titulo, corpo: p.corpo, url: p.url }, webpush: { headers: { Urgency: 'high', TTL: '172800' } } } } });
      });
    });
    if (!pedidos.length) return 0;
    var tok = fcmAcesso_(cfg.sa);
    var url = 'https://fcm.googleapis.com/v1/projects/' + cfg.web.projectId + '/messages:send';
    var maus = {};
    for (var i = 0; i < pedidos.length; i += 50) {
      var lote = pedidos.slice(i, i + 50).map(function (x) { return { url: url, method: 'post', contentType: 'application/json', headers: { Authorization: 'Bearer ' + tok }, muteHttpExceptions: true, payload: JSON.stringify(x.msg) }; });
      UrlFetchApp.fetchAll(lote).forEach(function (r, j) {
        var c = r.getResponseCode(), txt = r.getContentText();
        if (c === 200) enviados++;
        else if (c === 404 || /UNREGISTERED|registration token is not a valid|Requested entity was not found/i.test(txt)) maus[alvo[i + j].fcm] = 1;
        else console.log('FCM ' + c + ': ' + txt.slice(0, 300));
      });
    }
    if (Object.keys(maus).length) apagar_('Dispositivos', function (d) { return maus[d.fcm]; });
  } catch (e) { console.log('Push: ' + e); }
  return enviados;
}

/* ---------- API chamada pela app ---------- */
function apiPushConfig() {
  var cfg = pushCfg_();
  if (!cfg) return null;
  return { apiKey: cfg.web.apiKey, projectId: cfg.web.projectId, messagingSenderId: cfg.web.messagingSenderId, appId: cfg.web.appId, authDomain: cfg.web.authDomain || '', vapid: cfg.vapid };
}

function apiPushRegistar(token, fcm, plataforma, lang) {
  lang = lang === 'en' ? 'en' : 'pt';
  fcm = limpo_(fcm, 400); plataforma = limpo_(plataforma, 40);
  if (!fcm) throw new Error(tr_(lang, 'Não foi possível ativar os avisos neste telemóvel.', 'Could not enable notifications on this phone.'));
  var s = ler_('Inscritos').filter(function (x) { return x.token === token && x.estado === 'confirmado'; })[0];
  if (!s) throw new Error(tr_(lang, 'O seu acesso não é válido. Entre de novo com o seu email.', 'Your access is not valid. Please sign in again with your email.'));
  var novo = lock_(function () {
    var d = ler_('Dispositivos').filter(function (x) { return x.fcm === fcm; })[0];
    if (d) { d.email = s.email; d.plataforma = plataforma; d.visto_em = agora_(); atualizar_('Dispositivos', d); return false; }
    acrescentar_('Dispositivos', { fcm: fcm, email: s.email, plataforma: plataforma, criado_em: agora_(), visto_em: agora_() });
    return true;
  });
  if (novo) {
    var L = s.lingua || lang;
    PUSH_PEND_.push({ email: s.email, titulo: tr_(L, 'Avisos ativados', 'Notifications enabled'), corpo: tr_(L, 'Vai receber aqui os convites, lembretes e alterações dos eventos da Comenda.', 'You will receive Comenda event invitations, reminders and changes here.'), url: siteUrl_() || '' });
    enviarPushPendentes_();
  }
  return tr_(lang, 'Avisos ativados neste telemóvel.', 'Notifications enabled on this phone.');
}

function apiPushRemover(token, fcm) {
  fcm = limpo_(fcm, 400);
  var s = ler_('Inscritos').filter(function (x) { return x.token === token; })[0];
  if (s && fcm) lock_(function () { apagar_('Dispositivos', function (d) { return d.fcm === fcm; }); });
  return true;
}

/* Para testar no editor: escolher «testarPush» e Executar (envia para os telemóveis do primeiro administrador). */
function testarPush() {
  var emails = {};
  ler_('Dispositivos').forEach(function (d) { if (d.email) emails[d.email] = 1; });
  var lista = Object.keys(emails);
  if (lista.length > 3) { Logger.log('Há ' + lista.length + ' pessoas com avisos ativos: o teste só corre com até 3, para não incomodar os membros.'); return; }
  lista.forEach(function (email) { PUSH_PEND_.push({ email: email, titulo: 'Teste de notificação', corpo: 'Se está a ler isto no telemóvel, os avisos da Comenda estão a funcionar. Toque para abrir a app.', url: siteUrl_() || '' }); });
  Logger.log('Enviadas: ' + enviarPushPendentes_() + ' (para ' + (lista.join(', ') || 'ninguém') + ')');
}

/* Diagnóstico: confirma a configuração e o acesso ao Firebase, sem enviar nada. */
function diagnosticoPush() {
  var cfg = pushCfg_();
  if (!cfg) { Logger.log('Configuração em falta ou incompleta (FIREBASE_WEB, FCM_VAPID, FCM_SA).'); return; }
  Logger.log('Projeto: ' + cfg.web.projectId + ' · conta de serviço: ' + cfg.sa.client_email);
  CacheService.getScriptCache().remove('fcm_oauth');
  fcmAcesso_(cfg.sa);
  Logger.log('Acesso ao Firebase: OK · telemóveis registados: ' + ler_('Dispositivos').length);
}
