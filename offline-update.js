(()=>{
'use strict';
if(!('serviceWorker' in navigator) || !/^https?:$/.test(location.protocol))return;
let registration=null, installedVersion='', installedLabel='', pendingVersion='', dismissed='';
const notice=document.getElementById('update-notice');
const title=document.getElementById('update-notice-title');
const now=document.getElementById('update-now');
const later=document.getElementById('update-later');
const footer=document.querySelector('.footer-version');
const message=(worker,type)=>new Promise((resolve,reject)=>{
  if(!worker){reject(Error('No service worker'));return;}
  const channel=new MessageChannel();
  const timeout=setTimeout(()=>reject(Error('Worker response timeout')),4000);
  channel.port1.onmessage=e=>{clearTimeout(timeout);resolve(e.data);};
  worker.postMessage({type},[channel.port2]);
});
const remoteVersion=async()=>{
  const url=new URL('version.json',document.baseURI);
  url.searchParams.set('_',String(Date.now()));
  const response=await fetch(url.href,{cache:'no-store'});
  if(!response.ok)throw Error('Version lookup failed');
  return response.json();
};
const currentVersion=async()=>{
  try{
    const installed=await message(navigator.serviceWorker.controller||registration?.active,'GET_INSTALLED_VERSION');
    if(installed.version){installedLabel=installed.label||'';return installed.version;}
  }catch(e){}
  const latest=await remoteVersion();
  installedLabel=latest.label||'';
  return latest.version;
};
const check=async()=>{
  if(!navigator.onLine)return;
  try{
    const data=await remoteVersion();
    if(!installedVersion)installedVersion=await currentVersion();
    if(data.version && installedVersion && data.version!==installedVersion && data.version!==dismissed){
      pendingVersion=data.version;
      title.textContent='發現新版本 v'+data.version;
      notice.hidden=false;
    }else{notice.hidden=true;}
  }catch(e){}
};
now.addEventListener('click',async()=>{
  now.disabled=true;
  try{
    const worker=navigator.serviceWorker.controller||registration?.active;
    const result=await message(worker,'REFRESH_ASSETS');
    if(!result.ok)throw Error(result.error||'更新失敗');
    location.reload();
  }catch(e){now.disabled=false;alert('更新未完成，請保持網路連線後再試。原本資料及程式不會被清除。');}
});
later.addEventListener('click',()=>{dismissed=pendingVersion;notice.hidden=true;});
window.addEventListener('load',async()=>{
  try{
    registration=await navigator.serviceWorker.register('./service-worker.js');
    await navigator.serviceWorker.ready;
    installedVersion=await currentVersion();
    // Footer reads the installed version, not a hard-coded label.
    if(footer)footer.textContent=installedLabel||('v'+installedVersion);
    await check();
    setInterval(check,30*60*1000);
    window.addEventListener('online',check);
    document.addEventListener('visibilitychange',()=>{if(!document.hidden)check();});
  }catch(e){console.warn('離線更新功能無法啟用',e);}
});
})();
