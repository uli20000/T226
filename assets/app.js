
(()=>{'use strict';
const API_BASE='https://qhztjliketfujfiwemxh.supabase.co/functions/v1/wl226-api/api/';
let sessionToken='';try{sessionToken=sessionStorage.getItem('wl226.session')||''}catch{}
function rememberSession(token){sessionToken=token||'';try{if(sessionToken)sessionStorage.setItem('wl226.session',sessionToken);else sessionStorage.removeItem('wl226.session')}catch{}}
function apiFetch(path,options={}){return fetch(API_BASE+path,{...options,credentials:'omit',headers:{...(/^(preview|import|import-batch)(?:\?|$)/.test(path)?{'X-WL-Client-Version':'6'}:{}),...(sessionToken?{Authorization:'Bearer '+sessionToken}:{}),...options.headers}})}

if(window.wlBoot)window.wlBoot.began=true;
const $=id=>document.getElementById(id),PRODUCTS=[['tart','撻皮'],['meat','雞腿堡肉片'],['strip','腿條'],['chicken','九塊雞'],['flour','麵粉']];
const state={period:'day',summary:null,records:[],reports:[],uploads:[],loading:false,setup:false,token:location.hash.slice(1),timers:new Map(),queues:new Map(),dirty:new Set(),busy:false};
if(state.token)history.replaceState(null,'',location.pathname);
const fmt=n=>Number(n).toLocaleString('en-US',{minimumFractionDigits:1,maximumFractionDigits:1});
const signed=n=>(n>0?'+':n<0?'−':'')+fmt(Math.abs(n));
const esc=s=>String(s??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const today=()=>new Intl.DateTimeFormat('en-CA',{timeZone:'Asia/Taipei',year:'numeric',month:'2-digit',day:'2-digit'}).format(new Date());
const dateString=d=>d.getFullYear()+'-'+String(d.getMonth()+1).padStart(2,'0')+'-'+String(d.getDate()).padStart(2,'0');
function message(text,error=false){$('message').classList.toggle('error-box',error);$('message').replaceChildren();if(text){const s=document.createElement('span');s.textContent=text;$('message').append(s)}}
async function api(path,options={}){const readOnly=!options.method||options.method==='GET';const controller=readOnly?new AbortController():null;const timer=controller?setTimeout(()=>controller.abort(),20000):null;try{const r=await apiFetch(path,{...options,...(controller?{signal:controller.signal}:{}),headers:{...(options.body&&!(options.body instanceof FormData)?{'Content-Type':'application/json'}:{}),...options.headers}});let b;try{b=await r.json()}catch{throw Error('網站回應異常，請聯絡資訊人員。')}const version=r.headers.get('X-WL-API-Version');if(version&&version!=='6')throw Error('網站已更新，請重新整理後再操作');if(!r.ok){if(r.status===401&&path!=='login'){rememberSession('');state.dataFingerprint=null;$('app').hidden=true;$('auth').hidden=false}throw Error(b.error||'操作失敗')}if(path==='login'&&b.sessionToken)rememberSession(b.sessionToken);if(path==='logout'){rememberSession('');state.dataFingerprint=null}return b}catch(e){if(e.name==='AbortError')throw Error('連線逾時，請重新載入。');if(e instanceof TypeError)throw Error('無法連上網站，請檢查網路。');throw e}finally{if(timer)clearTimeout(timer)}}
function post(path,body){return api(path,{method:'POST',body:JSON.stringify(body)})}
let excelPromise;function loadExcel(){if(window.XLSX)return Promise.resolve();if(!excelPromise)excelPromise=new Promise((resolve,reject)=>{const script=document.createElement('script');script.src='./assets/xlsx.js';script.onload=resolve;script.onerror=()=>{excelPromise=null;script.remove();reject(Error('Excel 元件載入失敗，請重試'))};document.head.append(script)});return excelPromise}
function range(){if(state.period==='custom')return[$('start').value,$('end').value];const d=new Date($('anchor').value+'T12:00:00');let a=new Date(d),b=new Date(d);if(state.period==='week'){a.setDate(a.getDate()-(a.getDay()+6)%7);b=new Date(a);b.setDate(b.getDate()+6)}if(state.period==='month'){a.setDate(1);b=new Date(d.getFullYear(),d.getMonth()+1,0,12)}return[dateString(a),dateString(b)]}
function updateRange(){const[a,b]=range();const label={day:'日',week:'週',month:'月',custom:'區間'}[state.period];$('cumulative-title').textContent=label+'累計差異';$('material-scope').textContent=label+'查詢：'+(a===b?a:a+' – '+b);$('range-label').textContent=a===b?a.replaceAll('-','/'):a.replaceAll('-','/')+' – '+b.replaceAll('-','/');$('custom-range').hidden=state.period!=='custom';$('date-nav').hidden=state.period==='custom';document.querySelectorAll('[data-period]').forEach(x=>x.setAttribute('aria-pressed',String(x.dataset.period===state.period)))}
let activeProduct = null;
let activeProductRange = '';
let activeDate = null;
function needsReason(row) {
  return row.difference !== 0 && !row.reason.trim();
}
function badge() {
  const rows = (state.period === 'week' || state.period === 'month')
    ? state.records.filter(row => row.id === activeProduct && row.date === activeDate)
    : state.records;
  const count = rows.filter(needsReason).length;
  $('pending-count').textContent = !state.records.length ? '尚無資料' : count ? '待填原因 ' + count + ' 筆' : '原因已齊全';
}
function render() {
  const scope = state.period + ':' + range().join();
  if (activeProductRange !== scope) {
    activeProduct = null;
    activeDate = null;
    activeProductRange = scope;
  }
  if (materialQuery && (materialRange !== range().join() || materialFingerprint !== state.reports.map(row => row.date + ':' + row.uploadedAt).sort().join())) {
    clearMaterialResults(false);
    if (materialBusy) materialReload = true;
    else searchMaterials(0);
  }
  renderSummary();
  badge();
  const periodDrilldown = state.period === 'week' || state.period === 'month';
  const reasonPanel = document.querySelector('.layout-reasons');
  if (reasonPanel) reasonPanel.hidden = periodDrilldown;
  $('daily-details').textContent = '差異原因';
  const list = $('records');
  if (reasonPanel) reasonPanel.append(list);
  list.className = periodDrilldown ? 'list inline-date-reason' : 'list';
  list.replaceChildren();
  $('product-daily-details').replaceChildren();
  const labels = document.querySelector('.header-labels');
  if (labels) labels.hidden = true;
  if (!state.records.length) {
    const notice = document.createElement('p');
    notice.className = 'notice';
    notice.textContent = '此區間尚無報表，請至報表管理查看。';
    list.append(notice);
    return;
  }
  if ((state.period === 'week' || state.period === 'month') && activeProduct) {
    const section = document.createElement('section');
    section.className = 'panel list';
    section.setAttribute('aria-label', '產品逐日差異');
    const heading = document.createElement('h2');
    heading.textContent = PRODUCTS.find(([id]) => id === activeProduct)[1] + ' · 逐日差異';
    heading.style.textAlign = 'center';
    section.append(heading);
    const rows = state.records.filter(row => row.id === activeProduct).sort((a, b) => b.date.localeCompare(a.date));
    for (const row of rows) {
      const line = document.createElement('button');
      line.type = 'button';
      line.className = 'detail-cell';
      line.dataset.date = row.date;
      line.setAttribute('aria-expanded', String(activeDate === row.date));
      line.setAttribute('aria-controls', 'records');
      line.style.outline = activeDate === row.date ? '2px solid #245b49' : '';
      line.onclick = async () => {
        try {
          await flush();
          activeDate = activeDate === row.date ? null : row.date;
          render();
        } catch (error) {message(error.message, true);}
      };
      line.style.display = 'grid';
      line.style.gridTemplateColumns = '1fr 1fr';
      const date = document.createElement('span');
      date.textContent = row.date.replaceAll('-', '/');
      const value = document.createElement('strong');
      value.className = row.difference < 0 ? 'negative' : 'positive';
      value.textContent = signed(row.difference);
      line.append(date, value);
      section.append(line);
      if (activeDate === row.date) section.append(list);
    }
    if (!rows.length) {
      const notice = document.createElement('p');
      notice.textContent = '此品項在本期間尚無資料';
      section.append(notice);
    }
    $('product-daily-details').append(section);
  }
  if (periodDrilldown) {
    const selected = state.records.find(row => row.id === activeProduct && row.date === activeDate);
    if (selected) {
      $('daily-details').textContent = selected.name + ' · ' + selected.date.replaceAll('-', '/') + ' 差異原因';
      if (needsReason(selected)) appendPendingReason(selected, list, true);
      else {
        const reason = document.createElement('p');
        reason.className = 'notice';
        reason.textContent = selected.difference === 0 ? '當日無差異，免填原因' : selected.reason;
        list.append(reason);
      }
    }
    return;
  }
  if (state.period === 'day') {
    for (const [id] of PRODUCTS) {
      const row = state.records.find(row => row.id === id);
      if (!row) continue;
      if (needsReason(row)) appendPendingReason(row, list);
      else list.append(savedReason(row));
    }
    return;
  }
  const pending = state.records.filter(needsReason).sort((a, b) => b.date.localeCompare(a.date) || PRODUCTS.findIndex(([id]) => id === a.id) - PRODUCTS.findIndex(([id]) => id === b.id));
  if (pending.length) {
    for (const row of pending) appendPendingReason(row, list);
  }
}
function savedReason(row) {
  const article = document.createElement('article');
  article.className = 'record';
  article.dataset.key = row.date + '/' + row.id;
  article.innerHTML = '<div class="record-title"><h3>' + esc(row.name) + '</h3><small>' + esc(row.date.replaceAll('-', '/')) + '</small></div><div class="record-difference ' + (row.difference < 0 ? 'negative' : 'positive') + '">' + signed(row.difference) + '</div>';
  const reason = document.createElement('div');
  reason.className = 'reason';
  const label = document.createElement('span');
  label.textContent = '差異原因';
  const text = document.createElement('p');
  text.className = 'notice';
  text.style.margin = '0';
  text.textContent = row.difference === 0 ? '當日無差異，免填原因' : row.reason;
  reason.append(label, text);
  article.append(reason);
  return article;
}
function appendPendingReason(row, container, compact = false) {
  const key = row.date + '/' + row.id;
  const article = document.createElement('article');
  article.className = compact ? 'inline-reason' : 'record';
  article.dataset.key = key;
  article.innerHTML = (compact ? '' : '<div class="record-title"><h3>' + esc(row.name) + '</h3><small>' + esc(row.date.replaceAll('-', '/')) + '</small></div><div class="record-difference ' + (row.difference < 0 ? 'negative' : 'positive') + '">' + signed(row.difference) + '</div>') + '<label class="reason">差異原因<input type="text" maxlength="2000" placeholder="輸入差異原因" aria-label="' + esc(row.date + ' ' + row.name + '差異原因') + '"><span class="save-status" aria-live="polite">原因待填</span></label>';
  const input = article.querySelector('input');
  const status = article.querySelector('.save-status');
  input.value = row.reason;
  input.oninput = () => {
    row.reason = input.value;
    state.dirty.add(key);
    status.textContent = '等待儲存…';
    status.classList.remove('error');
    clearTimeout(state.timers.get(key));
    state.timers.set(key, setTimeout(() => queueSave(key, row, input, status), 600));
    badge();
  };
  container.append(article);
}
function queueSave(k,r,input,status){clearTimeout(state.timers.get(k));state.timers.delete(k);state.dirty.add(k);const value=input.value;status.textContent='儲存中…';const job=(state.queues.get(k)||Promise.resolve()).catch(()=>{}).then(async()=>{try{await post('reason',{date:r.date,id:r.id,reason:value,difference:r.difference});r.review=false;if(input.value===value){state.dirty.delete(k);status.textContent=value.trim()?'已儲存':'原因待填';status.classList.remove('error');if(value.trim()){if(state.period==='day')input.closest('article')?.replaceWith(savedReason(r));else if((state.period==='week'||state.period==='month')&&activeProduct===r.id&&activeDate===r.date)render();else input.closest('article')?.remove()}}badge()}catch(e){status.textContent='儲存失敗，點此重試';status.classList.add('error');status.onclick=()=>queueSave(k,r,input,status);message(e.message,true);throw e}});state.queues.set(k,job);job.catch(()=>{});return job}
async function flush(){for(const[k]of state.timers){const[date,id]=k.split('/');const r=state.records.find(x=>x.date===date&&x.id===id);const inputs=[...$('records').querySelectorAll('input')];const input=inputs.find(x=>x.getAttribute('aria-label')===date+' '+r.name+'差異原因');if(input)queueSave(k,r,input,input.parentElement.querySelector('.save-status'))}await Promise.all([...state.queues.values()]);if(state.dirty.size)throw Error('仍有原因未儲存，請重試後繼續')}
function renderSummary() {
  const summary = state.summary;
  if (!summary) return;
  $('cumulative-count').textContent = '本期間 · ' + summary.days + ' 日資料';
  $('cumulative-range').textContent = summary.days ? range().map(date => date.replaceAll('-', '/')).join(' – ') : '本期間尚無報表';
  const container = $('cumulative');
  container.innerHTML = summary.items.map(row => '<div class="product" data-product="' + esc(row.id) + '"><div class="product-name">' + esc(row.name) + '</div><div class="product-number ' + (row.difference < 0 ? 'negative' : 'positive') + '">' + (row.days ? signed(row.difference) : '—') + '</div><small>' + (row.days ? '累計 ' + row.days + ' 日' : '尚無資料') + '</small></div>').join('');
  if (state.period !== 'week' && state.period !== 'month') return;
  for (const card of container.querySelectorAll('[data-product]')) {
    card.setAttribute('role', 'button');
    card.setAttribute('tabindex', '0');
    card.setAttribute('aria-expanded', String(activeProduct === card.dataset.product));
    card.style.cursor = 'pointer';
    card.style.outline = activeProduct === card.dataset.product ? '2px solid #245b49' : '';
    const select = async () => {
      try {
        await flush();
        activeProduct = activeProduct === card.dataset.product ? null : card.dataset.product;
        activeDate = null;
        render();
      } catch (error) {message(error.message, true);}
    };
    card.onclick = select;
    card.onkeydown = event => {
      if (event.key === 'Enter' || event.key === ' ') {
        event.preventDefault();
        select();
      }
    };
  }
}

async function load(){if(state.loading){state.reload=true;return false}state.loading=true;let success=false;try{await flush();updateRange();const[a,b]=range();if(!a||!b||a>b)throw Error('請選擇有效的日期區間');$('sync-status').textContent='正在讀取所選日期…';const data=await api('records?start='+a+'&end='+b);const summary=data.summary;if(range().join()!==[a,b].join()){state.reload=true;return false}const fingerprint=JSON.stringify(data);const changed=state.dataFingerprint!==fingerprint;state.dataFingerprint=fingerprint;state.summary=summary;if(changed){state.records=data.records;state.reports=data.reports;state.queues.clear()}if(changed||state.renderedRange!==state.period+':'+[a,b].join()){state.renderedRange=state.period+':'+[a,b].join();render()}success=true;$('sync-status').textContent='同步於 '+new Date().toLocaleTimeString('zh-TW',{hour:'2-digit',minute:'2-digit'});}catch(e){message(e.message,true)}finally{state.loading=false;if(state.reload){state.reload=false;load()}}return success}
async function start(){state.setup=false;$('login-button').disabled=false;if(sessionToken){$('auth').hidden=true;$('app').hidden=false;window.wlBoot?.ready();const loaded=await load();if(loaded&&sessionToken&&!$('app').hidden)await cloudPending.process();return}$('auth').hidden=false;$('app').hidden=true;$('auth-title').textContent='登入每日差異';$('confirm-label').hidden=true;$('password').autocomplete='current-password';$('password').minLength=1;$('login-button').textContent='登入';$('login-button').disabled=false;$('auth-note').textContent='僅供授權同事共同使用';window.wlBoot?.ready()}
$('auth-form').onsubmit=async e=>{e.preventDefault();$('auth-error').textContent='';const button=$('login-button');button.disabled=true;try{const password=$('password').value;if(state.setup){if(password!==$('confirm-password').value)throw Error('兩次密碼不一致');await post('setup',{token:state.token,password});state.token=''}else await post('login',{account:$('account').value,password});$('password').value='';$('confirm-password').value='';await start()}catch(e){$('auth-error').textContent=e.message}finally{button.disabled=false}};
$('anchor').value=today();$('start').value=today();$('end').value=today();document.querySelectorAll('[data-period]').forEach(button=>button.onclick=async()=>{try{await flush();state.period=button.dataset.period;await load()}catch(e){message(e.message,true)}});$('anchor').onchange=load;$('start').onchange=load;$('end').onchange=load;
function shift(n){const d=new Date($('anchor').value+'T12:00:00');if(state.period==='month'){d.setDate(1);d.setMonth(d.getMonth()+n)}else d.setDate(d.getDate()+n*(state.period==='week'?7:1));$('anchor').value=dateString(d);load()}
$('previous').onclick=()=>shift(-1);$('next').onclick=()=>shift(1);$('logout').onclick=async()=>{try{await flush();await post('logout',{});state.records=[];await start()}catch(e){message(e.message,true)}};
// Candidate bridge. Runs on the existing W+L origin only; never sends sessions.
const DRIVE_ORIGIN='https://n-odhjluzlngrs2uj42cxlp43q43bxgnhh75xv2gq-0lu-script.googleusercontent.com';
function validEnvelope(event,{opener,nonce}){
 return !!opener&&event.source===opener&&event.origin===DRIVE_ORIGIN&&
 /^[a-f0-9-]{36}$/.test(nonce)&&event.data?.type==='wl226-files'&&event.data.nonce===nonce;
}
async function decodeFiles(items,{crypto,File,atob}){
 if(!Array.isArray(items)||items.length<1||items.length>31)throw Error('雲端檔案須為1至31份');
 let total=0;const seen=new Set(),files=[];
 for(const item of items){
  if(!item||typeof item.name!=='string'||!/^.{1,200}\.xlsx$/i.test(item.name)||/^~\$/.test(item.name)||/[\/\\\x00-\x1f]/.test(item.name))throw Error('雲端檔名不正確');
  if(typeof item.base64!=='string'||item.base64.length>28*1024*1024||!Number.isInteger(item.size)||item.size<1||item.size>20*1024*1024||!/^[a-f0-9]{64}$/.test(item.sha256))throw Error('雲端檔案超出限制');
  total+=item.size;if(total>50*1024*1024)throw Error('本批超過50MB');
  const bytes=Uint8Array.from(atob(item.base64),c=>c.charCodeAt(0));
  const hash=Array.from(new Uint8Array(await crypto.subtle.digest('SHA-256',bytes)),b=>b.toString(16).padStart(2,'0')).join('');
  if(bytes.length!==item.size||hash!==item.sha256)throw Error('雲端檔案傳輸核對失敗');
  if(seen.has(hash))continue;seen.add(hash);files.push(new File([bytes],item.name,{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'}));
 }
 return files;
}
function installReceiver({window,onFiles,onError}){
 const nonce=new URL(window.location.href).searchParams.get('cloudImport');
 const opener=window.opener;if(!opener||!nonce||!/^[a-f0-9-]{36}$/.test(nonce))return ()=>{};
 let used=false;const listener=async event=>{
  if(used||!validEnvelope(event,{opener,nonce}))return;
  used=true;window.removeEventListener('message',listener);clearInterval(timer);clearTimeout(expiry);
  try{const files=await decodeFiles(event.data.files,{crypto:window.crypto,File:window.File,atob:window.atob.bind(window)});await onFiles(files)}catch(error){onError(error)}
 };
 const ready=()=>opener.postMessage({type:'wl226-ready',nonce},DRIVE_ORIGIN);
 window.addEventListener('message',listener);const timer=setInterval(ready,1000);
 const expiry=setTimeout(()=>{used=true;window.removeEventListener('message',listener);clearInterval(timer);onError(Error('雲端傳送逾時，請重新讀取'))},120000);ready();
 return ()=>{used=true;window.removeEventListener('message',listener);clearInterval(timer);clearTimeout(expiry)};
}

// In-memory pending files survive auth/flush/preview failures; never retries a DB write.
function createPendingImport({isReady,flush,load,preview,notify}){
 let pending=null,busy=false;
 const report=(status,error)=>notify({status,pending:!!pending,error});
 async function process(){
  if(!pending||busy)return false;
  if(!isReady()){report('waiting-login');return false}
  busy=true;const files=pending;
  try{
   await flush();
   if(!await load()||!isReady()){report('waiting-login');return false}
   if(await preview(files)!==true){report('preview-failed');return false}
   if(pending===files)pending=null;
   report('preview-ready');return true;
  }catch(error){report('retry',error.message||String(error));return false}
  finally{busy=false}
 }
 return {enqueue:async files=>{if(pending)throw Error('仍有雲端報表待處理，請先重試');pending=files;report('received');return process()},process,hasPending:()=>!!pending};
}


for(const button of document.querySelectorAll('.report-actions button'))button.addEventListener('click',()=>{
 for(const action of button.parentElement.children)action.classList.toggle('is-selected',action===button);
},true);
const cloudPending=createPendingImport({isReady:()=>!!sessionToken&&!$('app').hidden&&!state.busy&&$('upload-modal').hidden,flush,load,preview:files=>previewFiles(files,{latestOnly:true}),notify:({status,pending,error})=>{const labels={'received':'已收到雲端報表','waiting-login':'雲端報表已保留，請登入或稍後重試','preview-failed':'雲端報表已保留，請核對預覽錯誤後重試','retry':'雲端報表已保留：'+error,'preview-ready':''};message(labels[status]??status,status==='retry'||status==='preview-failed');if(window.opener)window.opener.postMessage({type:'wl226-status',nonce:new URL(location.href).searchParams.get('cloudImport'),status},DRIVE_ORIGIN)}});
const driveButton=$('read-drive-button');
driveButton.onclick=async()=>{
 if(driveButton.disabled||state.busy||!$('upload-modal').hidden)return;
 driveButton.disabled=true;
 driveButton.textContent='讀取中…';
 try{
  if(cloudPending.hasPending()){await cloudPending.process();return}
  await flush();
  const {files}=await api('drive-reports');
  const reports=await decodeFiles(files,{crypto:window.crypto,File:window.File,atob:window.atob.bind(window)});
  await cloudPending.enqueue(reports);
 }catch(error){message(error.message,true)}
 finally{driveButton.disabled=false;driveButton.textContent='讀取資料'}
};
installReceiver({window,onFiles:files=>cloudPending.enqueue(files),onError:e=>message(e.message,true)});

let focusBefore;function closeModal(){if(state.busy)return;$('upload-modal').hidden=true;state.uploads=[];$('files').value='';focusBefore?.focus()}
$('upload-button').onclick=()=>{if(!managerBusy)$('files').click()};$('close-modal').onclick=closeModal;$('cancel-import').onclick=closeModal;
function uploadForm(file,date='',overwrite=false,sheet=''){const f=new FormData();f.append('file',file);if(sheet)f.append('sheet',sheet);if(date)f.append('date',date);if(overwrite)f.append('overwrite','yes');return f}
$('files').onchange=()=>{if(!$('files').files.length)return;closeManager();previewFiles([...$('files').files])};
async function previewFiles(files,{latestOnly=false}={}){if(!files.length)return;if(files.length>31){message('一次最多上傳 31 個檔案',true);return}focusBefore=document.activeElement;$('upload-date-options').hidden=latestOnly;$('upload-modal').hidden=false;$('upload-list').innerHTML='<div class="loading">正在讀取報表…</div>';$('upload-error').textContent='';$('confirm-import').disabled=true;state.uploads=[];state.busy=true;
 for(const file of files){try{const result=await api('preview',{method:'POST',body:uploadForm(file)});if(!Array.isArray(result.sheets)&&!Array.isArray(result.items))throw Error('預覽格式已更新，請重新整理後再選檔');for(const item of result.sheets||[result])state.uploads.push({file,result:item,sheet:item.sheet,date:item.date||'',error:item.error||(!Array.isArray(item.items)?'此分頁沒有有效報表資料':null),selected:!item.error,overwrite:false})}catch(e){state.uploads.push({file,error:e.message})}}state.busy=false;
 const latest=latestOnly?state.uploads.reduce((date,row)=>!row.error&&row.date>date?row.date:date,''):'';
 if(latestOnly){
  const candidates=state.uploads.filter(row=>!row.error&&row.date===latest&&latest);
  const hasError=state.uploads.some(row=>row.error||!row.date);
  const issue=hasError||!latest?'有報表無法判定日期或內容，請檢查雲端檔案後重新讀取。':
   candidates.length!==1?'同一天有多份報表，請整理雲端檔案後重新讀取。':'';
  state.uploads=issue?[]:candidates;
  $('upload-list').replaceChildren();
  $('upload-skipped').textContent=latest?'報表日期 '+latest:'';
  $('upload-error').textContent=issue;
  updateImportButton();
  $('close-modal').focus();
  return true;
 }
 const months={};for(const u of state.uploads)if(u.date&&!u.error)months[u.date.slice(0,7)]=(months[u.date.slice(0,7)]||0)+1;const month=Object.keys(months).sort((a,b)=>months[b]-months[a])[0];if(month){$('upload-month').value=month;for(const u of state.uploads)u.selected=!u.error&&u.date.startsWith(month)}$('upload-skipped').textContent='依分頁日期匯入，摘要頁不納入。';renderUploads();$('close-modal').focus();return !state.uploads.some(u=>u.error)}
function renderUploads(){$('upload-list').replaceChildren();if($('upload-date-options').hidden){updateImportButton();return}for(const u of state.uploads){const card=document.createElement('div');card.className='upload-card';const title=document.createElement('div');title.className='file-name';title.textContent=u.file.name+(u.sheet?' · '+u.sheet:'');title.setAttribute('title',u.file.name);const selection=document.createElement('label');selection.className='upload-selection';const chosen=document.createElement('input');chosen.type='checkbox';chosen.checked=!!u.selected;chosen.disabled=!!u.error;chosen.setAttribute('aria-label','選取 '+u.file.name+' '+(u.sheet||''));chosen.onchange=()=>{u.selected=chosen.checked;updateImportButton()};selection.append(chosen,title);card.append(selection);if(u.error){const e=document.createElement('p');e.className='error-text';e.textContent=u.error;card.append(e)}else{const facts=document.createElement('div');facts.className='upload-facts';for(const text of ['226 店',u.result.items.length+' 個品項',...(u.result.warnings||[])]){const span=document.createElement('span');span.textContent=text;facts.append(span)}card.append(facts);const label=document.createElement('label');label.textContent=u.date?'報表日期（可更正）':'日期異常，請手動指定';const input=document.createElement('input');input.type='date';input.value=u.date;input.onchange=()=>{u.date=input.value;updateImportButton()};label.append(input);card.append(label);const check=document.createElement('label');check.className='check';const box=document.createElement('input');box.type='checkbox';box.checked=u.overwrite;box.onchange=()=>u.overwrite=box.checked;check.append(box,document.createTextNode(u.result.exists?'已有同日資料，確認覆蓋':'若日期重複，同意覆蓋既有資料'));card.append(check)}$('upload-list').append(card)}updateImportButton()}
function updateImportButton(){$('confirm-import').disabled=state.busy||!state.uploads.some(u=>u.selected&&!u.error&&u.date)}
const batchJobs=new WeakMap();
async function uploadBatch(file,rows,onProgress){const selection=rows.map(u=>({date:u.date,sheet:u.sheet||u.result.sheet,overwrite:!!u.overwrite}));const signature=JSON.stringify(selection);let job=batchJobs.get(file);if(!job||job.signature!==signature){job={id:crypto.randomUUID(),signature};batchJobs.set(file,job)}const form=new FormData();form.append('file',file);form.append('selection',signature);form.append('jobId',job.id);const response=await apiFetch('import-batch',{method:'POST',headers:{'X-WL-Client-Version':'6'},body:form});if(!response.ok){const error=await response.json();throw Error(error.error||'批次匯入失敗')}if(response.headers.get('Content-Type')?.includes('application/json')){const result=await response.json();if(!result.ok)throw Error(result.error||'批次匯入失敗');return result}const reader=response.body.getReader(),decoder=new TextDecoder();let buffer='',complete=null;function consume(line){if(!line.trim())return;const event=JSON.parse(line);if(event.type==='progress')onProgress(event);if(event.type==='complete')complete=event}for(;;){const{value,done}=await reader.read();if(done)break;buffer+=decoder.decode(value,{stream:true});let end;while((end=buffer.indexOf('\n'))>=0){consume(buffer.slice(0,end));buffer=buffer.slice(end+1)}}buffer+=decoder.decode();if(buffer.trim())consume(buffer);if(!complete)throw Error('連線中斷，部分日期可能已完成；請稍後重試同一批次');if(!complete.ok){batchJobs.delete(file);throw Error(complete.error||'批次匯入失敗')}return complete}
$('confirm-import').onclick=async()=>{if(state.busy)return;const selected=state.uploads.filter(u=>u.selected&&!u.error);if(!selected.length)return;if(selected.some(u=>!u.date)){$('upload-error').textContent='請先指定日期';return}if(new Set(selected.map(u=>u.date)).size!==selected.length){$('upload-error').textContent='本批次有相同日期，請分批匯入';return}const groups=new Map();for(const u of selected){if(!groups.has(u.file))groups.set(u.file,[]);groups.get(u.file).push(u)}if([...groups.values()].some(rows=>rows.length>31)){$('upload-error').textContent='請一次選取最多 31 天，按月份分批匯入';return}state.busy=true;updateImportButton();$('upload-error').textContent='';const started=Date.now();let processed=0,imported=0,skipped=0;const errors=[];const progress=()=>{$('confirm-import').textContent='已處理 '+processed+'/'+selected.length;$('upload-skipped').textContent='批次匯入中 · '+Math.floor((Date.now()-started)/1000)+' 秒'};progress();const timer=setInterval(progress,1000);try{for(const[file,rows]of groups){const before=processed;try{const result=await uploadBatch(file,rows,event=>{processed=before+event.processed;progress()});processed=before+rows.length;for(const item of result.results){const u=rows.find(u=>u.date===item.date&&u.sheet===item.sheet);if(item.status==='imported'){if(u)u.done=true;imported++}else if(item.status==='skipped'){if(u)u.done=true;skipped++}else errors.push(item.date+'：'+item.error)}}catch(e){errors.push(file.name+'：'+e.message)}}}finally{clearInterval(timer);state.busy=false;$('confirm-import').textContent='確認匯入'}const completed=selected.filter(u=>u.done).sort((a,b)=>b.date.localeCompare(a.date))[0];if(completed){state.period='day';$('anchor').value=completed.date;await load()}const elapsed=((Date.now()-started)/1000).toFixed(1);if(errors.length){state.uploads=state.uploads.filter(u=>!u.done);renderUploads();$('upload-error').textContent=errors.join('；')}else{closeModal();if(completed)$('daily-details').focus()}message('已匯入 '+imported+' 天'+(skipped?'，已存在 '+skipped+' 天略過':'')+' · '+elapsed+' 秒'+(errors.length?'，其餘請重試':''),!!errors.length)};
$('export-button').onclick=async()=>{const button=$('export-button');button.disabled=true;try{await flush();if(!await load())throw Error('同步失敗，請確認連線後再匯出');if(!state.records.length)throw Error('此日期區間沒有可匯出的資料');await loadExcel();const header=['日期','店號','產品','單位','當日報表期初','進貨','退貨','調入','調出','損耗','銷售','理論庫存','實際盤點','差異','原因','狀態','報表期初','期初來源','期初核對'];const rows=state.records.slice().sort((a,b)=>a.date.localeCompare(b.date)||PRODUCTS.findIndex(p=>p[0]===a.id)-PRODUCTS.findIndex(p=>p[0]===b.id)).map(r=>[r.date,'226',r.name,r.unit,...['opening','purchase','returns','transferIn','transferOut','loss','sales','theoretical','actual','difference'].map(k=>Number(r[k].toFixed(1))),r.difference?(r.reason.trim()||'原因待填'):'',!r.difference?'無差異':r.reason.trim()?'已填寫':'原因待填',Number((r.reportOpening??r.opening).toFixed(1)),'本日報表期初','']);const ws=XLSX.utils.aoa_to_sheet([header,...rows]);ws['!cols']=header.map((_,i)=>({wch:i===14?40:i===2?18:14}));ws['!autofilter']={ref:ws['!ref']};for(const k of Object.keys(ws)){if(!k.startsWith('!')&&ws[k].t==='n')ws[k].z='0.0'}const notes=[['日期','來源檔案','缺少品項／注意事項','原始檔保存'],...state.reports.map(r=>[r.date,r.filename,(r.warnings||[]).join('；'),'網站雲端，OneDrive 待連接'])];const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,ws,'每日差異');XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(notes),'匯入紀錄');const[a,b]=range();XLSX.writeFile(wb,'226-WL-'+a+'-'+b+'.xlsx');message('Excel 已下載')}catch(e){message(e.message,true)}finally{button.disabled=false}};
window.addEventListener('beforeunload',e=>{if(state.dirty.size){e.preventDefault();e.returnValue=''}});document.addEventListener('keydown',e=>{const modal=!$('manager-modal').hidden?$('manager-modal'):!$('upload-modal').hidden?$('upload-modal'):null;if(!modal)return;if(e.key==='Escape'){if(modal.id==='manager-modal')$('close-manager').click();else closeModal()}if(e.key==='Tab'){const nodes=[...modal.querySelectorAll('button:not(:disabled),input:not(:disabled),a[href]')].filter(node=>node.getClientRects().length);const first=nodes[0],last=nodes.at(-1);if(e.shiftKey&&document.activeElement===first){e.preventDefault();last?.focus()}else if(!e.shiftKey&&document.activeElement===last){e.preventDefault();first?.focus()}}});
$('select-month').onclick=()=>{const month=$('upload-month').value;if(!month)return;for(const u of state.uploads)u.selected=!u.error&&u.date.startsWith(month);renderUploads()};
let materialQuery='',materialOffset=0,materialRange='',materialFingerprint='',materialBusy=false,materialReload=false;
function clearMaterialResults(reset=true){if(reset)materialQuery='';materialOffset=0;$('material-list').replaceChildren();$('material-breakdown').hidden=true;$('material-totals').replaceChildren();$('material-pages').hidden=true;$('material-caption').textContent='正在依日期更新查詢…'}
async function searchMaterials(offset=0){if(materialBusy)return;const query=$('material-query').value.trim();if(!query)return;const[start,end]=range();const requestedRange=[start,end].join();materialBusy=true;$('material-search-button').disabled=true;$('material-previous').disabled=true;$('material-next').disabled=true;$('material-caption').textContent='正在查詢原始報表…';try{const p=new URLSearchParams({q:query,start,end,offset:String(offset)});const result=await api('material-search?'+p);if(range().join()!==requestedRange){clearMaterialResults(false);materialReload=true;return}materialQuery=query;materialOffset=offset;materialRange=requestedRange;materialFingerprint=state.reports.map(r=>r.date+':'+r.uploadedAt).sort().join();$('material-breakdown').hidden=!result.results.length;const totals=$('material-totals');totals.replaceChildren();for(const r of result.totals||[]){const card=document.createElement('div');card.className='product';const name=document.createElement('div');name.className='product-name';name.textContent=r.name;name.title=r.name;const number=document.createElement('div');number.className='product-number '+(r.difference<0?'negative':'positive');number.textContent=signed(r.difference);const label=document.createElement('small');label.textContent=({day:'日',week:'週',month:'月',custom:'區間'}[state.period])+'差異 · '+r.days+' 日資料';card.append(name,number,label);totals.append(card)}const list=$('material-list');list.replaceChildren();for(const r of result.results){const card=document.createElement('article');card.className='material-card';const head=document.createElement('div');head.className='material-heading';const title=document.createElement('h3');title.textContent=r.name;title.title=r.name;const delta=document.createElement('div');delta.className='product-number '+(r.difference<0?'negative':'positive');delta.textContent=r.error?'—':signed(r.difference);head.append(title,delta);const sub=document.createElement('div');sub.className='material-sub';for(const text of [r.date,r.code,r.unit||'數量']){const span=document.createElement('span');span.textContent=text;sub.append(span)}card.append(head,sub);if(r.error){const error=document.createElement('p');error.className='error-text';error.textContent='數字格式異常，請查看原檔';card.append(error)}else{const detail=document.createElement('details');const summary=document.createElement('summary');summary.textContent='查看數量明細';detail.append(summary);const grid=document.createElement('div');grid.className='detail-grid';for(const[label,value]of [['報表期初',r.reportOpening??r.opening],['進貨',r.purchase],['退貨',r.returns],['調入',r.transferIn],['調出',r.transferOut],['損耗',r.loss],['銷售',r.sales],['理論庫存',r.theoretical],['實際盤點',r.actual]]){const cell=document.createElement('div');cell.className='detail-cell';const name=document.createElement('span');name.textContent=label;const number=document.createElement('strong');number.textContent=fmt(value);cell.append(name,number);grid.append(cell)}const note=document.createElement('div');note.className='carry-note';const source=document.createElement('span');source.textContent=r.openingSource==='previous'?'採用前一天實際盤點':'前日無資料，採報表期初';note.append(source);if(r.openingMismatch){const warning=document.createElement('span');warning.className='danger';warning.textContent='前日盤點與報表期初不同';note.append(warning)}detail.append(grid,note);card.append(detail)}list.append(card)}const caption=$('material-caption');caption.replaceChildren();const rangeLabel=document.createElement('span');rangeLabel.textContent=start+' – '+end;const count=document.createElement('span');count.textContent=result.total?' · '+(offset+1)+'–'+Math.min(offset+100,result.total)+'／'+result.total+' 筆':' · 查無符合品項';caption.append(rangeLabel,count);for(const w of result.warnings){const p=document.createElement('div');p.className='error-text';p.textContent=w.date+'：'+w.message;caption.append(p)}$('material-pages').hidden=result.total<=100;$('material-previous').disabled=offset===0;$('material-next').disabled=offset+100>=result.total}catch(e){$('material-caption').textContent=e.message;$('material-list').replaceChildren();$('material-breakdown').hidden=true;$('material-totals').replaceChildren();$('material-pages').hidden=true}finally{materialBusy=false;$('material-search-button').disabled=false;if(materialReload){materialReload=false;searchMaterials(0)}}}
$('material-form').onsubmit=e=>{e.preventDefault();searchMaterials(0)};$('material-previous').onclick=()=>searchMaterials(Math.max(0,materialOffset-100));$('material-next').onclick=()=>searchMaterials(materialOffset+100);
let managerMode='active',managerBusy=false;
function closeManager(){if(managerBusy)return;$('manager-modal').hidden=true;$('manage-button').focus()}
$('close-manager').onclick=closeManager;
async function manage(){if(managerBusy)return;managerBusy=true;$('manager-error').textContent='';try{const data=await api('reports');$('show-active').classList.toggle('primary',managerMode==='active');$('show-trash').classList.toggle('primary',managerMode==='trash');const rows=managerMode==='active'?data.reports:data.trash;const list=$('manager-list');list.replaceChildren();if(!rows.length){const p=document.createElement('p');p.className='muted';p.textContent=managerMode==='active'?'尚無已匯入報表':'回收區沒有報表';list.append(p)}for(const r of rows){const card=document.createElement('div');card.className='manager-card';const title=document.createElement('h3');title.textContent=r.date;const filename=document.createElement('div');filename.className='file-name';filename.textContent=r.filename+(r.sheet?' · '+r.sheet:'');filename.title=filename.textContent;const meta=document.createElement('div');meta.className='manager-meta';const time=document.createElement('span');time.textContent=(managerMode==='active'?'上傳：':'移除：')+new Date(r.uploadedAt||r.deletedAt).toLocaleString('zh-TW',{month:'2-digit',day:'2-digit',hour:'2-digit',minute:'2-digit',hour12:false});meta.append(time);const actions=document.createElement('div');actions.className='manager-actions';if(managerMode==='active'){const download=document.createElement('a');download.className='source-link';download.href='#';download.onclick=event=>{event.preventDefault();downloadOriginal(r.date)};download.textContent='下載原檔';const remove=document.createElement('button');remove.className='danger';remove.textContent='移至回收區';remove.onclick=async()=>{if(!confirm('將 '+r.date+' 的報表、數字與原因移至回收區？\n該日不再計入累計，次日期初也會重新核對。'))return;remove.disabled=true;await managerAction('report-delete',{date:r.date,confirm:r.date,uploadedAt:r.uploadedAt})};actions.append(download,remove)}else{const restore=document.createElement('button');restore.textContent='還原報表';restore.onclick=()=>managerAction('report-restore',{id:r.id});const purge=document.createElement('button');purge.className='danger';purge.textContent='永久刪除';purge.onclick=()=>{const text=prompt('此操作無法還原。\n請輸入：'+r.date+'永久刪除');if(text!==r.date+'永久刪除')return;managerAction('report-purge',{id:r.id,confirm:text})};actions.append(restore,purge)}card.append(title,filename,meta,actions);list.append(card)}}catch(e){$('manager-error').textContent=e.message}finally{managerBusy=false}}
async function managerAction(path,data){if(managerBusy)return;managerBusy=true;try{await post(path,data);await load()}catch(e){$('manager-error').textContent=e.message;managerBusy=false;return}managerBusy=false;await manage()}
$('manage-button').onclick=async()=>{try{await flush();$('manager-modal').hidden=false;managerMode='active';await manage();$('close-manager').focus()}catch(e){message(e.message,true)}};
$('show-active').onclick=()=>{managerMode='active';manage()};$('show-trash').onclick=()=>{managerMode='trash';manage()};
setInterval(()=>{if(document.visibilityState==='visible'&&!$('app').hidden&&!state.dirty.size&&!state.busy&&document.activeElement?.tagName!=='INPUT'&&$('upload-modal').hidden&&$('manager-modal').hidden)load()},30000);updateRange();start();
})();

