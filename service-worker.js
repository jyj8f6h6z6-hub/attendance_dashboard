// Version is defined only in version.json. Cache refresh occurs when the user confirms an update.
const CACHE = 'attendance-static';
const STAGING = 'attendance-staging';
const CORE = ['./','./index.html','./style.css','./age-analysis.css','./script.js','./visitation-export.js','./age-analysis.js','./offline-update.js','./site.webmanifest','./version.json','./xlsx.full.min.js','./icon-192.png','./icon-512.png'];
async function downloadAssets(){
  const stage = await caches.open(STAGING);
  await Promise.all(CORE.map(async path => {
    const request = new Request(new URL(path, self.registration.scope), {cache:'no-store'});
    const response = await fetch(request);
    if(!response.ok) throw Error('Failed to refresh: '+path);
    await stage.put(new URL(path, self.registration.scope), response);
  }));
  return stage;
}
async function activateAssets(){
  const stage=await downloadAssets();
  const target=await caches.open(CACHE);
  for(const req of await stage.keys()) await target.put(req, await stage.match(req));
  await caches.delete(STAGING);
}
self.addEventListener('install', event => {
  event.waitUntil((async()=>{await activateAssets();await self.skipWaiting();})());
});
self.addEventListener('activate',event=>{
  event.waitUntil((async()=>{
    const keys=await caches.keys();
    await Promise.all(keys.filter(k=>k.startsWith('attendance-static-')||k===STAGING).map(k=>caches.delete(k)));
    await self.clients.claim();
  })());
});
self.addEventListener('message',event=>{
  if(event.data?.type==='REFRESH_ASSETS'){
    event.waitUntil((async()=>{
      try {await activateAssets();event.ports[0]?.postMessage({ok:true});}
      catch(error){await caches.delete(STAGING);event.ports[0]?.postMessage({ok:false,error:String(error)});}
    })());
  }
  if(event.data?.type==='GET_INSTALLED_VERSION'){
    event.waitUntil((async()=>{
      try {
        const response=await (await caches.open(CACHE)).match(new URL('./version.json',self.registration.scope));
        const data=response?await response.json():null;
        event.ports[0]?.postMessage({version:data?.version||null,label:data?.label||null});
      }catch(e){event.ports[0]?.postMessage({version:null});}
    })());
  }
});
self.addEventListener('fetch',event=>{
  const req=event.request;
  if(req.method!=='GET')return;
  const url=new URL(req.url);
  if(url.origin!==self.location.origin)return;
  if(url.pathname.endsWith('/version.json')){
    event.respondWith(fetch(new Request(req,{cache:'no-store'})).catch(()=>caches.match(req)));
    return;
  }
  event.respondWith((async()=>{
    const cached=await caches.match(req);
    if(cached)return cached;
    try{
      const response=await fetch(req);
      if(response.ok){const cache=await caches.open(CACHE);await cache.put(req,response.clone());}
      return response;
    }catch(e){
      if(req.mode==='navigate')return (await caches.match(new URL('./index.html',self.registration.scope)))||Response.error();
      return Response.error();
    }
  })());
});
