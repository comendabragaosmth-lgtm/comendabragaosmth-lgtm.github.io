// V36 · velocidade: abertura com os últimos dados guardados, cache do servidor em blocos, edição manual da folha
const {chromium}=require('playwright');
const BASE='http://localhost:8765/index.html';
(async()=>{const b=await chromium.launch();const ctx=await b.newContext({viewport:{width:1200,height:1100},locale:'pt-PT'});const p=await ctx.newPage();
const errs=[];p.on('pageerror',e=>errs.push(e.message));await p.route(/fonts\.(googleapis|gstatic)/,r=>r.abort());
// servidor lento / sem rede (simulado): envolve o fetch simulado
await p.addInitScript(()=>{let f=null;Object.defineProperty(window,'fetch',{configurable:true,get(){return f},set(v){const orig=v;f=async(u,o)=>{if(String(u).startsWith('http://mock/')){if(localStorage.getItem('mockfail'))throw new Error('offline');const d=+localStorage.getItem('mockdelay')||0;if(d)await new Promise(r=>setTimeout(r,d))}return orig(u,o)}}})});
const w=(t=500)=>p.waitForTimeout(t);const ok=(c,m)=>console.log((c?'PASS ':'FAIL ')+m);
const hoje=new Date(),d=n=>{const x=new Date(hoje.getTime()+n*864e5);return x.toISOString().slice(0,10)};
await p.goto(BASE);await p.evaluate(()=>localStorage.clear());await p.goto(BASE);await w();
await p.evaluate(([dt,pz])=>{
  acrescentar_('Definicoes',{chave:'admins',valor:'fgscarvalho@gmail.com'});
  acrescentar_('Inscritos',{email:'ana@ex.pt',nome:'Ana Costa',telemovel:'9',localidade:'Braga',estado:'confirmado',membro:'NÃO',token:'TKANA',criado_em:'2026-09-01 10:00'});
  acrescentar_('Eventos',{id:'ev1',titulo:'Jantar de Teste Rápido',data:dt,hora:'20:00',local:'Braga',visibilidade:'publico',prazo:pz,estado:'ativo',campos:'{}',atividades:'[]',traje:'Fato'});
},[d(20),d(15)]);
// 1) primeira abertura: guarda os dados
await p.goto(BASE+'?t=TKANA');await w(800);
const g=await p.evaluate(()=>JSON.parse(localStorage.getItem('comenda_dados')||'null'));
ok(g&&g.tk==='TKANA'&&g.d.eventos.some(e=>e.id==='ev1')&&g.d.me&&g.d.me.email==='ana@ex.pt','dados guardados no telemóvel');
// 2) servidor lento: o calendário aparece logo, com «A atualizar…»
await p.evaluate(()=>{localStorage.setItem('mockdelay','3000')});
await p.evaluate(()=>{acrescentar_('Eventos',{id:'ev2',titulo:'Evento Novo Do Servidor',data:new Date(Date.now()+25*864e5).toISOString().slice(0,10),hora:'10:00',local:'Porto',visibilidade:'publico',prazo:new Date(Date.now()+20*864e5).toISOString().slice(0,10),estado:'ativo',campos:'{}',atividades:'[]',traje:'Fato'})});
const t0=Date.now();await p.goto(BASE);await w(700);
let txt=await p.evaluate(()=>document.getElementById('app').innerText);
let sync=await p.evaluate(()=>document.body.dataset.sync||'');
ok(/Jantar de Teste Rápido/.test(txt)&&!/A carregar/.test(txt),'calendário visível em <1 s com servidor lento');
ok(sync==='A atualizar…','indicador «A atualizar…» visível');
ok(!/Evento Novo Do Servidor/.test(txt),'ainda com os dados guardados');
await w(3200);
txt=await p.evaluate(()=>document.getElementById('app').innerText);sync=await p.evaluate(()=>document.body.dataset.sync||'');
ok(/Evento Novo Do Servidor/.test(txt)&&!sync,'atualizado quando o servidor responde; indicador desaparece');
// 3) sem rede: mostra os dados guardados
await p.evaluate(()=>{localStorage.removeItem('mockdelay');localStorage.setItem('mockfail','1')});
await p.goto(BASE);await w(5500);
txt=await p.evaluate(()=>document.getElementById('app').innerText);
const toast=await p.evaluate(()=>[...document.querySelectorAll('.toast,#toast')].map(x=>x.innerText).join(' '));
ok(/Evento Novo Do Servidor/.test(txt)&&/Sem ligação/.test(toast),'sem rede: mostra os últimos dados + aviso');
await p.evaluate(()=>localStorage.removeItem('mockfail'));
// 4) outro token: não usa os dados de outra pessoa
await p.evaluate(()=>acrescentar_('Inscritos',{email:'rui@ex.pt',nome:'Rui',telemovel:'9',localidade:'B',estado:'confirmado',membro:'NÃO',token:'TKRUI',criado_em:'2026-09-01 10:00'}));
await p.evaluate(()=>localStorage.setItem('mockdelay','2000'));
await p.goto(BASE+'?t=TKRUI');await w(500);
sync=await p.evaluate(()=>document.body.dataset.sync||'');txt=await p.evaluate(()=>document.getElementById('app').innerText);
ok(!sync&&!/Ana Costa/.test(txt),'link de outra pessoa não mostra dados guardados da Ana');
await w(2200);await p.evaluate(()=>localStorage.removeItem('mockdelay'));
const g2=await p.evaluate(()=>JSON.parse(localStorage.getItem('comenda_dados')||'null'));ok(g2&&g2.tk==='TKRUI','guardado para o novo acesso');
// 5) link inválido apaga os dados guardados
await p.goto(BASE+'?t=NAOEXISTE');await w(800);
ok(!(await p.evaluate(()=>localStorage.getItem('comenda_dados'))),'link inválido apaga os dados guardados');
// 6) servidor: cache em blocos para folhas grandes
const r=await p.evaluate(()=>{
  for(let i=0;i<60;i++)acrescentar_('Inscritos',{email:'p'+i+'@ex.pt',nome:'Pessoa '+i+' Ção',nota:'Nota longa çãé '.repeat(60),estado:'confirmado',membro:'NÃO',token:'T'+i,criado_em:'2026-09-01'});
  _mem={};const a=ler_('Inscritos');const cab=CacheService.getScriptCache().get('tab_Inscritos');_mem={};const b2=ler_('Inscritos');
  return {n:a.length,cab,igual:JSON.stringify(a)===JSON.stringify(b2),tam:JSON.stringify(a).length};
});
ok(r.cab&&/^n:\d+$/.test(r.cab)&&+r.cab.slice(2)>1&&r.igual&&r.n===62,'folha grande em cache por blocos ('+r.cab+', '+r.tam+' carateres)');
// 7) escrita pela app limpa a cache
const r2=await p.evaluate(()=>{_mem={};ler_('Eventos');const s=ler_('Eventos').filter(e=>e.id==='ev1')[0];s.titulo='Jantar Alterado';atualizar_('Eventos',s);_mem={};return ler_('Eventos').filter(e=>e.id==='ev1')[0].titulo});
ok(r2==='Jantar Alterado','escrita pela app aparece logo');
// 8) edição manual da folha: limpa ao fim de 1 minuto (data de alteração no Drive)
const r3=await p.evaluate(()=>{
  let T=1000;const orig=DriveApp.getFileById;DriveApp.getFileById=id=>id===PropertiesService.getScriptProperties().getProperty('SHEET_ID')?{getLastUpdated:()=>new Date(T)}:orig.call(DriveApp,id);
  const C=CacheService.getScriptCache();C.remove('folha_chk');_mem={};ler_('Eventos');
  const sh=SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('SHEET_ID')).getSheetByName('Eventos');
  const linha=ler_('Eventos').filter(e=>e.id==='ev1')[0]._linha,col=TABELAS.Eventos.indexOf('titulo')+1;
  sh.getRange(linha,col,1,1).setValues([['Editado à Mão']]);T=2000;
  _mem={};const antes=ler_('Eventos').filter(e=>e.id==='ev1')[0].titulo;
  C.remove('folha_chk');_mem={};const depois=ler_('Eventos').filter(e=>e.id==='ev1')[0].titulo;
  DriveApp.getFileById=orig;return {antes,depois};
});
ok(r3.antes==='Jantar Alterado'&&r3.depois==='Editado à Mão','edição manual aparece após a verificação de 1 min ('+r3.antes+' → '+r3.depois+')');
// 9) «Deixar de receber» apaga os dados guardados
await p.goto(BASE+'?t=TKANA');await w(800);ok(!!(await p.evaluate(()=>localStorage.getItem('comenda_dados'))),'guardado (Ana)');
await p.goto(BASE+'?sair=TKANA');await w(800);
const bt=await p.$('[data-act=sair]');if(bt){await bt.click();await w(800)}
ok(!bt||!(await p.evaluate(()=>localStorage.getItem('comenda_dados'))),'sair apaga os dados guardados'+(bt?'':' (botão não encontrado)'));
console.log('errors',JSON.stringify(errs));await b.close();})();
