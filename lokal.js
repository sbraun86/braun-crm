/* Braun CRM – Microsoft-365-Version (läuft ohne PC, z. B. auf iPhone und iPad).
   Ersetzt lokal.js: wird in der Web-Version unter dem Namen „lokal.js“ ausgeliefert.
   - Anmeldung mit dem Microsoft-365-Konto (MSAL, Weiterleitung statt Popup – funktioniert auch als Home-Bildschirm-App)
   - Daten: dieselben JSON-Dateien wie die PC-Version, im OneDrive unter 20 Sven Braun Executive Search/CRM/daten
   - Speichern nur geänderter Einträge; bei gleichzeitigen Änderungen wird zusammengeführt (ETag / If-Match)
   - E-Mail-Entwürfe, Termine und Mail-Abgleich über Microsoft Graph */
(function(){
"use strict";
// Aus der App-Registrierung in Microsoft Entra (keine Geheimnisse)
const CONFIG={
  clientId:'253f8281-a04d-441a-8f08-f6f2bcab2035',   // Anwendungs-ID (Client-ID) „Braun CRM“
  tenantId:'d94e24d7-1070-4ab4-9aa5-7da8ce563aab'    // Verzeichnis-ID (Mandanten-ID) Sven Braun Executive Search
};
const MSAL_URL='https://cdn.jsdelivr.net/npm/@azure/msal-browser@3.27.0/lib/msal-browser.min.js';
const SCOPES=['User.Read','Files.ReadWrite','Mail.ReadWrite','Calendars.ReadWrite'];
const GRAPH='https://graph.microsoft.com/v1.0';
const DIR='/me/drive/root:/'+['20 Sven Braun Executive Search','CRM','daten'].map(encodeURIComponent).join('/');
const COLLS=['leads','contacts','accounts','deals','activities','meta'];
const cache={}, files={}, pending={}, timers={}, waiters={};   // files[c] = {id, eTag}
const subs={}, docSubs=[];
const clone=x=>x==null?x:JSON.parse(JSON.stringify(x));
const anyPending=()=>COLLS.some(c=>pending[c]&&pending[c].size);
window.CRM_REMOTE=true;   // keine Outlook-Desktop-Funktionen; Mail-App bzw. Outlook im Web
window.CRM_M365=true;

function setStatus(ok){try{if(typeof S!=='undefined'){S.conn=ok?'ok':'err';render()}}catch(e){}}

/* ---------- Anmeldung ---------- */
let pca=null, account=null;
function loadScript(src){return new Promise((res,rej)=>{const s=document.createElement('script');s.src=src;s.onload=res;s.onerror=()=>rej({text:'Programmbaustein konnte nicht geladen werden (Internetverbindung?)'});document.head.appendChild(s)})}
function gate(html){
  let g=document.getElementById('m365-gate');
  if(!g){g=document.createElement('div');g.id='m365-gate';g.style.cssText='position:fixed;inset:0;z-index:200;display:grid;place-items:center;background:var(--bg,#f4f6f8);font-family:system-ui,-apple-system,sans-serif;padding:16px';document.body.appendChild(g)}
  g.innerHTML=`<div style="background:var(--surface,#fff);border:1px solid var(--line,#e2e6eb);border-radius:12px;padding:28px 24px;width:min(360px,100%);text-align:center;color:var(--fg,#1d2330)"><div style="font-size:22px;font-weight:700;margin-bottom:6px">Braun CRM</div>${html}</div>`;
  return g;
}
const ready=(async()=>{
  await new Promise(r=>document.readyState==='loading'?document.addEventListener('DOMContentLoaded',r):r());
  if(!CONFIG.clientId||!CONFIG.tenantId){gate('<p style="color:#646c7a">Die Microsoft-365-Verbindung ist noch nicht eingerichtet (Client-ID und Mandanten-ID fehlen).</p>');throw new Error('nicht eingerichtet')}
  await loadScript(MSAL_URL);
  pca=new msal.PublicClientApplication({auth:{clientId:CONFIG.clientId,authority:'https://login.microsoftonline.com/'+CONFIG.tenantId,redirectUri:location.origin+location.pathname,navigateToLoginRequestUrl:false},cache:{cacheLocation:'localStorage'}});
  await pca.initialize();
  const r=await pca.handleRedirectPromise();
  account=(r&&r.account)||pca.getActiveAccount()||pca.getAllAccounts()[0]||null;
  if(!account){
    const g=gate('<p style="color:#646c7a;margin:0 0 18px">Bitte mit deinem Microsoft-365-Konto anmelden.</p><button id="m365-login" style="width:100%;font-size:16px;font-weight:600;padding:12px;border:0;border-radius:8px;background:#0a6f78;color:#fff">Mit Microsoft anmelden</button>');
    await new Promise(()=>{g.querySelector('#m365-login').onclick=()=>pca.loginRedirect({scopes:SCOPES})});   // Seite wird verlassen
  }
  pca.setActiveAccount(account);
  const g=document.getElementById('m365-gate'); if(g) g.remove();
})();
async function token(){
  await ready;
  try{return (await pca.acquireTokenSilent({scopes:SCOPES,account})).accessToken}
  catch(e){if(e instanceof msal.InteractionRequiredAuthError){await pca.acquireTokenRedirect({scopes:SCOPES,account});return new Promise(()=>{})}throw e}
}
async function graph(path,opt={}){
  const t=await token();
  const r=await fetch(/^https:/.test(path)?path:GRAPH+path,{...opt,headers:{Authorization:'Bearer '+t,...(opt.headers||{})}});
  return r;
}
async function graphJson(path,opt){const r=await graph(path,opt);if(!r.ok){let m='';try{m=(await r.json()).error.message}catch(e){}throw {status:r.status,text:m||('Microsoft 365 meldet Fehler '+r.status)}}return r.status===204?null:r.json()}

/* ---------- Daten im OneDrive ---------- */
async function listFiles(){
  const j=await graphJson(DIR+':/children?$select=id,name,eTag&$top=200');
  const out={}; for(const f of j.value||[]){const m=f.name.match(/^([a-z]+)\.json$/);if(m&&COLLS.includes(m[1]))out[m[1]]={id:f.id,eTag:f.eTag}}
  return out;
}
async function readColl(c){
  const f=files[c]; if(!f) return {};
  const r=await graph('/me/drive/items/'+f.id+'/content',{cache:'no-store'});
  if(!r.ok) throw {status:r.status,text:'Datei '+c+'.json konnte nicht gelesen werden'};
  const t=(await r.text()).trim(); if(!t) return {};
  const d=JSON.parse(t); return d&&typeof d==='object'&&!Array.isArray(d)?d:{};
}
let loading=null;
function load(){
  if(!loading) loading=(async()=>{
    Object.assign(files,await listFiles());
    await Promise.all(COLLS.map(async c=>{cache[c]=await readColl(c)}));
  })();
  return loading;
}
function collSnap(c){return {docs:Object.keys(cache[c]).map(id=>({id,data:()=>clone(cache[c][id])})),metadata:{fromCache:false}}}
function docSnap(c,id){const d=cache[c][id];return {exists:d!==undefined,data:()=>clone(d),metadata:{fromCache:false}}}
function notify(c){(subs[c]||[]).forEach(cb=>{try{cb(collSnap(c))}catch(e){}});docSubs.filter(x=>x.c===c).forEach(x=>{try{x.cb(docSnap(c,x.id))}catch(e){}})}
// Geänderte Einträge auf den aktuellen Stand der Datei anwenden und mit If-Match schreiben;
// hat ein anderes Gerät dazwischen gespeichert (412), neu lesen und erneut zusammenführen.
async function writeColl(c,p){
  for(let attempt=0;attempt<4;attempt++){
    let base;
    if(attempt===0) base=clone(cache[c]);
    else {Object.assign(files,await listFiles());base=await readColl(c);for(const [id,op] of p){if(op==='del')delete base[id];else if(cache[c][id]!==undefined)base[id]=clone(cache[c][id])}}
    const body=JSON.stringify(base);
    const f=files[c];
    const url=f?'/me/drive/items/'+f.id+'/content':DIR+'/'+c+'.json:/content';
    const r=await graph(url,{method:'PUT',headers:{'Content-Type':'text/plain',...(f?{'If-Match':f.eTag}:{})},body});
    if(r.status===412||r.status===409) continue;
    if(!r.ok) throw {status:r.status,text:'Speichern bei Microsoft 365 fehlgeschlagen ('+r.status+')'};
    const item=await r.json(); files[c]={id:item.id,eTag:item.eTag};
    if(attempt>0){cache[c]=base;notify(c)}   // Änderungen des anderen Geräts übernommen
    return;
  }
  throw {text:'Gleichzeitige Änderung konnte nicht zusammengeführt werden. Bitte erneut versuchen.'};
}
function flush(c){
  clearTimeout(timers[c]); timers[c]=null;
  const ws=waiters[c]||[]; waiters[c]=[];
  const p=pending[c]; pending[c]=new Map(); if(!p||!p.size){ws.forEach(w=>w.res());return Promise.resolve()}
  return writeColl(c,p).then(()=>{ws.forEach(w=>w.res());setStatus(true)},
    e=>{for(const [id,op] of p)if(!pending[c].has(id))pending[c].set(id,op);ws.forEach(w=>w.rej(e));setStatus(false)});
}
function schedule(c,id,op){
  (pending[c]=pending[c]||new Map()).set(id,op);
  return new Promise((res,rej)=>{(waiters[c]=waiters[c]||[]).push({res,rej});clearTimeout(timers[c]);timers[c]=setTimeout(()=>flush(c),600)});
}
// Änderungen anderer Geräte (oder der PC-Version) übernehmen
let refreshing=false;
async function refresh(){
  if(refreshing||!loading||anyPending()) return;
  refreshing=true;
  try{
    const now=await listFiles();
    const changed=COLLS.filter(c=>now[c]&&(!files[c]||now[c].eTag!==files[c].eTag));
    for(const c of changed){files[c]=now[c];const d=await readColl(c);if(anyPending())break;cache[c]=d;notify(c)}
  }catch(e){}finally{refreshing=false}
}
setInterval(()=>{if(document.visibilityState==='visible')refresh()},15000);
document.addEventListener('visibilitychange',()=>{if(document.visibilityState==='visible')refresh()});

function docRef(c,id){
  return {
    set(data){return load().then(()=>{cache[c][id]=clone(data);return schedule(c,id,'set')})},
    delete(){return load().then(()=>{delete cache[c][id];return schedule(c,id,'del')})},
    onSnapshot(cb,err){docSubs.push({c,id,cb});load().then(()=>cb(docSnap(c,id)),err||(()=>{}));return ()=>{}}
  };
}
const db={
  collection(c){return {doc:id=>docRef(c,id),onSnapshot(cb,err){(subs[c]=subs[c]||[]).push(cb);load().then(()=>cb(collSnap(c)),err||(()=>{}));return ()=>{}}}},
  doc(path){const [c,id]=String(path).split('/');return docRef(c,id)}
};
const downloads={
  save({filename,data}){
    const blob=data instanceof Blob?data:new Blob([data],{type:/\.json$/i.test(filename)?'application/json':'text/csv;charset=utf-8'});
    const a=document.createElement('a');a.href=URL.createObjectURL(blob);a.download=filename;document.body.appendChild(a);a.click();a.remove();setTimeout(()=>URL.revokeObjectURL(a.href),5000);
    return Promise.resolve();
  }
};

/* ---------- Outlook über Microsoft Graph ---------- */
const addr=list=>(list||[]).filter(Boolean).map(a=>({emailAddress:{address:String(a)}}));
const pad=n=>String(n).padStart(2,'0');
const localStamp=iso=>{const d=new Date(iso);return d.getFullYear()+'-'+pad(d.getMonth()+1)+'-'+pad(d.getDate())+'T'+pad(d.getHours())+':'+pad(d.getMinutes())};
const outlook={
  outlook:true,
  // Entwurf anlegen; 'display' öffnet ihn zusätzlich in Outlook im Web
  async mail(d){
    const m=await graphJson('/me/messages',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({subject:d.subject||'',body:{contentType:'HTML',content:'<div style="font-family:Aptos,Calibri,Arial,sans-serif;font-size:11pt">'+(d.html||'')+'</div>'},toRecipients:addr(d.to),ccRecipients:addr(d.cc),bccRecipients:addr(d.bcc)})});
    if(d.mode!=='draft'&&m&&m.webLink) window.open(m.webLink,'_blank');
    return {ok:true};
  },
  async termin(d){
    const start=new Date(d.start), end=d.end?new Date(d.end):new Date(start.getTime()+3600000);
    const f=x=>localStamp(x.toISOString())+':00';
    await graphJson('/me/events',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({subject:d.subject||'',body:{contentType:'text',content:d.body||''},start:{dateTime:f(start),timeZone:'Europe/Berlin'},end:{dateTime:f(end),timeZone:'Europe/Berlin'},location:d.location?{displayName:d.location}:undefined,isReminderOn:true,reminderMinutesBeforeStart:15})});
    return {ok:true};
  },
  // Posteingang und Gesendete Elemente seit d.since nach Adressen der Leads/Kontakte durchsuchen
  async sync(d){
    const want=new Set((d.addrs||[]).map(a=>String(a).toLowerCase()));
    const since=new Date(d.since).toISOString();
    const out=[], seen=new Set();
    for(const [folder,dir,field] of [['inbox','in','receivedDateTime'],['sentitems','out','sentDateTime']]){
      let url=`/me/mailFolders/${folder}/messages?$filter=${encodeURIComponent(field+' ge '+since)}&$orderby=${encodeURIComponent(field+' desc')}&$top=100&$select=internetMessageId,subject,from,toRecipients,ccRecipients,bccRecipients,receivedDateTime,sentDateTime,body`;
      for(let page=0;url&&page<10;page++){
        const j=await graphJson(url,{headers:{Prefer:'outlook.body-content-type="text"'}});
        for(const m of j.value||[]){
          const from=String(((m.from||{}).emailAddress||{}).address||'').toLowerCase();
          const to=[...(m.toRecipients||[]),...(m.ccRecipients||[]),...(m.bccRecipients||[])].map(r=>String((r.emailAddress||{}).address||'').toLowerCase());
          if(dir==='in'?!want.has(from):!to.some(a=>want.has(a))) continue;
          const id=m.internetMessageId||m.id; if(seen.has(id)) continue; seen.add(id);
          let text=String((m.body||{}).content||'').replace(/\r/g,'').replace(/(\n[ \t]*){3,}/g,'\n\n').trim(); if(text.length>1500) text=text.slice(0,1500)+' ...';
          out.push({id,dir,date:localStamp(m[field]),subject:m.subject||'',from,to,text});
        }
        url=j['@odata.nextLink']||null;
      }
    }
    return {ok:true,mails:out};
  }
};

/* ---------- Einlesen im Browser (Visitenkarte, Lebenslauf) ----------
   Die Bausteine werden erst beim ersten Einlesen geladen (Texterkennung inkl. deutschem Sprachpaket ca. 4 MB). */
const LIB={
  tess:'https://cdn.jsdelivr.net/npm/tesseract.js@5.1.1/dist/tesseract.min.js',
  tessWorker:'https://cdn.jsdelivr.net/npm/tesseract.js@5.1.1/dist/worker.min.js',
  tessCore:'https://cdn.jsdelivr.net/npm/tesseract.js-core@5.1.1',
  pdf:'https://cdn.jsdelivr.net/npm/pdfjs-dist@3.11.174/build/pdf.min.js',
  pdfWorker:'https://cdn.jsdelivr.net/npm/pdfjs-dist@3.11.174/build/pdf.worker.min.js',
  zip:'https://cdn.jsdelivr.net/npm/jszip@3.10.1/dist/jszip.min.js'
};
const libs={};
const lib=(k,glob)=>libs[k]||(libs[k]=loadScript(LIB[k]).then(()=>window[glob]));
let ocrWorker=null;
// Foto auf eine für die Texterkennung gute Größe bringen (ca. 2000 px); Farbe bleibt, Tesseract binarisiert selbst
async function prepImage(img){
  const src=img instanceof HTMLCanvasElement?img:await createImageBitmap(img);
  const s=Math.min(2.5,2000/Math.max(src.width,src.height));
  if(Math.abs(s-1)<0.05&&src instanceof HTMLCanvasElement) return src;
  const c=document.createElement('canvas'); c.width=Math.round(src.width*s); c.height=Math.round(src.height*s);
  c.getContext('2d').drawImage(src,0,0,c.width,c.height);
  return c;
}
// Schräg fotografierte Karten automatisch gerade richten (rotateAuto); bei unsicherem Ergebnis zweiter Lesemodus
async function ocrImage(img,opt={}){
  const T=await lib('tess','Tesseract');
  if(!ocrWorker) ocrWorker=T.createWorker(['deu','eng'],1,{workerPath:LIB.tessWorker,corePath:LIB.tessCore});
  const w=await ocrWorker, c=await prepImage(img);
  let best=null;
  for(const psm of opt.page?['3']:['3','11']){
    await w.setParameters({tessedit_pageseg_mode:psm});
    const {data}=await w.recognize(c,{rotateAuto:true});
    // Zeilen mit sehr geringer Sicherheit (Logos, Muster) weglassen
    const lines=(data.lines||[]).filter(l=>l.confidence>=35&&/[A-Za-zÄÖÜäöü0-9@]{2}/.test(l.text)).map(l=>l.text.trim());
    const r={conf:data.confidence||0,text:lines.length?lines.join('\n'):(data.text||'')};
    if(!best||r.conf>best.conf) best=r;
    if(best.conf>=75) break;
  }
  return best.text;
}
// PDF: Textebene lesen (Zeilen nach Höhe zusammensetzen); Seiten ohne Text (Scan) per Texterkennung
async function pdfText(file){
  const pdfjs=await lib('pdf','pdfjsLib'); pdfjs.GlobalWorkerOptions.workerSrc=LIB.pdfWorker;
  const doc=await pdfjs.getDocument({data:await file.arrayBuffer()}).promise;
  const out=[];
  for(let i=1;i<=Math.min(doc.numPages,3);i++){
    const page=await doc.getPage(i); const tc=await page.getTextContent(); const rows={};
    for(const it of tc.items){if(!it.str)continue;const y=Math.round(it.transform[5]/3);(rows[y]=rows[y]||[]).push(it)}
    const txt=Object.keys(rows).map(Number).sort((a,b)=>b-a).map(y=>rows[y].sort((a,b)=>a.transform[4]-b.transform[4]).map(x=>x.str).join(' ').replace(/\s+/g,' ').trim()).filter(Boolean).join('\n');
    if(txt.replace(/\s/g,'').length>40){out.push(txt);continue}
    const vp=page.getViewport({scale:2}); const c=document.createElement('canvas'); c.width=vp.width; c.height=vp.height;
    await page.render({canvasContext:c.getContext('2d'),viewport:vp}).promise; out.push(await ocrImage(c,{page:true}));
  }
  return out.join('\n\n');
}
// DOCX: Kopfzeilen, Dokument, Fußzeilen als Text
async function docxText(file){
  const Z=await lib('zip','JSZip'); const z=await Z.loadAsync(await file.arrayBuffer());
  const rank=n=>/header/.test(n)?0:/document/.test(n)?1:2;
  const names=Object.keys(z.files).filter(n=>/^word\/(header\d*|document|footer\d*)\.xml$/.test(n)).sort((a,b)=>rank(a)-rank(b));
  const out=[];
  for(const n of names){const x=await z.file(n).async('string');const t=document.createElement('textarea');t.innerHTML=x.replace(/<\/w:p>/g,'\n').replace(/<w:tab\/>/g,'\t').replace(/<w:br\/>/g,'\n').replace(/<[^>]+>/g,'');out.push(t.value)}
  return out.join('\n');
}
window.crmTools={
  ocr:async blob=>({ok:true,text:await ocrImage(blob)}),
  doc:async file=>{
    const ext=file.name.split('.').pop().toLowerCase();
    if(ext==='pdf') return {ok:true,text:await pdfText(file)};
    if(ext==='docx') return {ok:true,text:await docxText(file)};
    throw {text:'Dieses Format geht in der App nicht. Bitte als PDF oder .docx speichern.'};
  },
  msg:()=>Promise.reject({text:'.msg-Dateien gehen in der App nicht. Bitte die Mail öffnen, Text kopieren und einfügen.'}),
  selected:()=>Promise.reject({text:'In der App nicht verfügbar.'}),
  auth:{
    status:async()=>{await ready;return {m365:true,local:false,authed:true,user:account&&(account.name||account.username),username:account&&account.username}},
    logout:async()=>{await ready;pca.logoutRedirect({account})}
  }
};

window.claude={use(n){
  if(n==='db') return load().then(()=>db);
  if(n==='downloads') return Promise.resolve(downloads);
  if(n==='mcp') return ready.then(()=>outlook);
  return Promise.reject({code:'not_available'});
}};

// Ungespeicherte Änderungen beim Verlassen: so schnell wie möglich schreiben
document.addEventListener('visibilitychange',()=>{if(document.visibilityState==='hidden')COLLS.forEach(c=>{if(pending[c]&&pending[c].size)flush(c)})});
})();
