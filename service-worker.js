const VERSION = '1.1.8';
const CACHE = 'attendance-static-' + VERSION;
const CORE = ['./','./index.html','./style.css','./age-analysis.css','./script.js','./age-analysis.js','./offline-update.js','./site.webmanifest','./version.json','./icon-192.png','./icon-512.png'];
const XLSX_CDN = 'https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js';
self.addEventListener('install',event=>{
  event.waitUntil((async()=>{
    const cache=await caches.open(CACHE);
    await cache.addAll(CORE);
    // Optional library: cache when online; lack of CDN must not break offline shell installation.
    try {await cache.add(new Request(XLSX_CDN,{mode:'cors'}));}catch(e){}
  })());
});
self.addEventListener('activate',event=>{
  event.waitUntil((async()=>{
    const keys=await caches.keys();
    await Promise.all(keys.filter(k=>k.startsWith('attendance-static-')&&k!==CACHE).map(k=>caches.delete(k)));
    await self.clients.claim();
  })());
});
self.addEventListener('message',event=>{if(event.data?.type==='SKIP_WAITING')self.skipWaiting();});
self.addEventListener('fetch',event=>{
  const req=event.request;
  if(req.method!=='GET')return;
  const url=new URL(req.url);
  if(url.origin!==self.location.origin && url.href!==XLSX_CDN)return;
  if(url.pathname.endsWith('/version.json')){
    event.respondWith(fetch(new Request(req,{cache:'no-store'})).catch(()=>caches.match(req)));
    return;
  }
  // Cache first ensures the app and its saved data remain usable offline.
  event.respondWith((async()=>{
    const cached=await caches.match(req);
    if(cached)return cached;
    try{
      const response=await fetch(req);
      if(response.ok && (url.origin===self.location.origin || url.href===XLSX_CDN)){
        const cache=await caches.open(CACHE);await cache.put(req,response.clone());
      }
      return response;
    }catch(e){
      if(req.mode==='navigate')return (await caches.match('./index.html'))||Response.error();
      return Response.error();
    }
  })());
});
