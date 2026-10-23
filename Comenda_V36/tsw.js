const {chromium}=require('playwright');const {spawn}=require('child_process');
const ok=(c,m)=>console.log((c?'PASS ':'FAIL ')+m);const sleep=t=>new Promise(r=>setTimeout(r,t));
const srv=()=>spawn('python3',['-m','http.server','8770'],{cwd:'/home/claude/comenda/pkg/swt',stdio:'ignore'});
(async()=>{let s=srv();await sleep(800);const b=await chromium.launch();const ctx=await b.newContext();const p=await ctx.newPage();
await p.route(/script\.google|fonts\./,r=>r.abort());
await p.goto('http://localhost:8770/');await p.waitForFunction(()=>navigator.serviceWorker&&navigator.serviceWorker.controller,null,{timeout:15000}).catch(()=>{});
await p.reload();await sleep(1000);
const ctl=await p.evaluate(()=>!!navigator.serviceWorker.controller);ok(ctl,'service worker ativo');
const ks=await p.evaluate(async()=>(await caches.keys()));ok(ks.includes('comenda-v3'),'cache comenda-v3 '+ks);
// versão nova no servidor: a abertura seguinte ainda mostra a guardada, a outra já mostra a nova
const fs=require('fs');const f='/home/claude/comenda/pkg/swt/index.html';const orig=fs.readFileSync(f,'utf8');fs.writeFileSync(f,orig.replace('</body>','<div id="marca">V-NOVA</div></body>'));
await p.goto('http://localhost:8770/?t=X');await sleep(1200);let m=await p.evaluate(()=>!!document.getElementById('marca'));ok(!m,'abre logo a versão guardada');
await p.goto('http://localhost:8770/');await sleep(800);m=await p.evaluate(()=>!!document.getElementById('marca'));ok(m,'na abertura seguinte já usa a versão nova');
fs.writeFileSync(f,orig);
// sem servidor: continua a abrir
s.kill();await sleep(500);
const r=await p.goto('http://localhost:8770/').catch(e=>null);await sleep(800);const t=await p.evaluate(()=>document.title).catch(()=>'');ok(!!r&&/Comenda/.test(t),'sem rede abre a página guardada');
await b.close();})();
