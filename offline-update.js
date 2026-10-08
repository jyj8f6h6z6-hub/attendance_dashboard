(()=>{
'use strict';
const CURRENT='1.1.12';
if(!('serviceWorker' in navigator) || !/^https?:$/.test(location.protocol))return;
let registration=null, pendingVersion='', dismissed='';
const notice=document.getElementById('update-notice');
const title=document.getElementById('update-notice-title');
const now=document.getElementById('update-now');
const later=document.getElementById('update-later');
const show=(version)=>{
 if(!version || version===CURRENT || version===dismissed)return;
 pendingVersion=version;title.textContent='發現新版本 v'+version;notice.hidden=false;
};
const check=async()=>{
 if(!navigator.onLine)return;
 try{
  const url=new URL('version.json',document.baseURI);
  url.searchParams.set('_',String(Date.now()));
  const r=await fetch(url.href,{cache:'no-store'});
  if(!r.ok)return;
  const data=await r.json();
  if(data.version && data.version!==CURRENT){
    await registration?.update();
    show(data.version);
  }
 }catch(e){}
};
now.addEventListener('click',async()=>{
 now.disabled=true;
 try{
  if(registration){await registration.update();
   if(registration.waiting){
    let done=false;
    navigator.serviceWorker.addEventListener('controllerchange',()=>{if(!done){done=true;location.reload();}},{once:true});
    registration.waiting.postMessage({type:'SKIP_WAITING'});
    return;
   }
  }
  location.reload();
 }catch(e){now.disabled=false;location.reload();}
});
later.addEventListener('click',()=>{dismissed=pendingVersion;notice.hidden=true;});
window.addEventListener('load',async()=>{
 try{
  registration=await navigator.serviceWorker.register('./service-worker.js');
  if(registration.waiting)await check();
  await check();
  setInterval(check,30*60*1000);
  window.addEventListener('online',check);
  document.addEventListener('visibilitychange',()=>{if(!document.hidden)check();});
 }catch(e){console.warn('離線功能無法啟用',e);}
});
})();
