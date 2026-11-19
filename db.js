/* Small IndexedDB storage wrapper. No server requests. */
(function(root){
 'use strict';
 const DB='document-constructor-v1';let connection;
 const open=()=>connection||(connection=new Promise((resolve,reject)=>{
   if(!('indexedDB' in root))return reject(new Error('Нет доступа к IndexedDB.'));
   const req=indexedDB.open(DB,1);
   req.onupgradeneeded=()=>{const db=req.result;for(const n of ['templates','drafts','session'])if(!db.objectStoreNames.contains(n))db.createObjectStore(n);};
   req.onsuccess=()=>resolve(req.result);req.onerror=()=>reject(req.error);
 }));
 async function op(store,action,key,value){const db=await open();return new Promise((resolve,reject)=>{
   const tx=db.transaction(store,action==='get'?'readonly':'readwrite'),s=tx.objectStore(store);
   const req=action==='get'?s.get(key):action==='put'?s.put(value,key):s.delete(key);
   req.onsuccess=()=>resolve(req.result);req.onerror=()=>reject(req.error);
 });}
 const storage={get:(s,k)=>op(s,'get',k),put:(s,k,v)=>op(s,'put',k,v),remove:(s,k)=>op(s,'delete',k)};
 root.LocalDB=storage;
})(typeof window!=='undefined'?window:globalThis);
