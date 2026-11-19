/* Small dependency-free ZIP reader/writer for modern browsers. The reader uses
   native DecompressionStream('deflate-raw'); ZIP writer uses STORE entries. */
(function(root){
  'use strict';
  const encoder = new TextEncoder(), decoder = new TextDecoder('utf-8');
  const u16=(v,p)=>v.getUint16(p,true),u32=(v,p)=>v.getUint32(p,true);
  function readZip(buffer){
    const bytes=buffer instanceof Uint8Array?buffer:new Uint8Array(buffer);
    const view=new DataView(bytes.buffer,bytes.byteOffset,bytes.byteLength);
    let eocd=-1;
    for(let p=bytes.length-22;p>=Math.max(0,bytes.length-65557);p--){if(u32(view,p)===0x06054b50){eocd=p;break;}}
    if(eocd<0) throw new Error('Файл не является корректным ZIP/XLSX.');
    const entries=new Map(); let pos=u32(view,eocd+16),count=u16(view,eocd+10);
    if(count===65535||pos===0xFFFFFFFF) throw new Error('ZIP64 пока не поддерживается.');
    for(let n=0;n<count;n++){
      if(pos+46>bytes.length||u32(view,pos)!==0x02014b50) throw new Error('Повреждена ZIP-структура.');
      const flags=u16(view,pos+8),method=u16(view,pos+10),size=u32(view,pos+20),len=u16(view,pos+28),extra=u16(view,pos+30),comment=u16(view,pos+32),local=u32(view,pos+42);
      if(flags&1) throw new Error('Парольная защита ZIP/XLSX не поддерживается.');
      const name=decoder.decode(bytes.subarray(pos+46,pos+46+len));
      entries.set(name,{method,size,local});pos+=46+len+extra+comment;
    }
    async function get(name){
      const item=entries.get(name); if(!item) return null;
      let p=item.local;
      if(p+30>bytes.length||u32(view,p)!==0x04034b50) throw new Error('Некорректная запись в ZIP.');
      p+=30+u16(view,p+26)+u16(view,p+28);
      if(p+item.size>bytes.length) throw new Error('ZIP-файл обрезан.');
      const raw=bytes.subarray(p,p+item.size);
      if(item.method===0) return new Uint8Array(raw);
      if(item.method!==8) throw new Error('Неподдерживаемый метод сжатия ZIP.');
      if(!('DecompressionStream' in root)) throw new Error('Для чтения XLSX требуется современная версия Safari.');
      const stream=new Blob([raw]).stream().pipeThrough(new DecompressionStream('deflate-raw'));
      const uncompressed=new Uint8Array(await new Response(stream).arrayBuffer());
      if(uncompressed.byteLength>150*1024*1024) throw new Error('Слишком большой распакованный элемент XLSX.');
      return uncompressed;
    }
    return {entries:[...entries.keys()],get,async text(name){const b=await get(name);return b?decoder.decode(b):null;}};
  }
  const crcTable=new Uint32Array(256);
  for(let i=0;i<256;i++){let c=i;for(let j=0;j<8;j++)c=(c&1)?0xEDB88320^(c>>>1):c>>>1;crcTable[i]=c>>>0;}
  function crc32(bytes){let c=0xFFFFFFFF;for(let i=0;i<bytes.length;i++)c=crcTable[(c^bytes[i])&255]^(c>>>8);return (c^0xFFFFFFFF)>>>0;}
  function writeZip(files){
    const locals=[],central=[];let offset=0;const now=new Date();
    const dosTime=(now.getHours()<<11)|(now.getMinutes()<<5)|Math.floor(now.getSeconds()/2);
    const dosDate=((Math.max(1980,now.getFullYear())-1980)<<9)|((now.getMonth()+1)<<5)|now.getDate();
    for(const [name,content] of Object.entries(files)){
      const nameBytes=encoder.encode(name),data=content instanceof Uint8Array?content:encoder.encode(String(content)),c=crc32(data);
      const local=new Uint8Array(30+nameBytes.length+data.length),lv=new DataView(local.buffer);
      lv.setUint32(0,0x04034b50,true);lv.setUint16(4,20,true);lv.setUint16(6,0x800,true);lv.setUint16(8,0,true);
      lv.setUint16(10,dosTime,true);lv.setUint16(12,dosDate,true);lv.setUint32(14,c,true);lv.setUint32(18,data.length,true);lv.setUint32(22,data.length,true);
      lv.setUint16(26,nameBytes.length,true);local.set(nameBytes,30);local.set(data,30+nameBytes.length);locals.push(local);
      const cent=new Uint8Array(46+nameBytes.length),cv=new DataView(cent.buffer);
      cv.setUint32(0,0x02014b50,true);cv.setUint16(4,20,true);cv.setUint16(6,20,true);cv.setUint16(8,0x800,true);
      cv.setUint16(10,0,true);cv.setUint16(12,dosTime,true);cv.setUint16(14,dosDate,true);cv.setUint32(16,c,true);
      cv.setUint32(20,data.length,true);cv.setUint32(24,data.length,true);cv.setUint16(28,nameBytes.length,true);cv.setUint32(42,offset,true);
      cent.set(nameBytes,46);central.push(cent);offset+=local.length;
    }
    const centralLength=central.reduce((s,b)=>s+b.length,0);const tail=new Uint8Array(22),tv=new DataView(tail.buffer);
    tv.setUint32(0,0x06054b50,true);tv.setUint16(8,central.length,true);tv.setUint16(10,central.length,true);
    tv.setUint32(12,centralLength,true);tv.setUint32(16,offset,true);
    return new Blob([...locals,...central,tail],{type:'application/zip'});
  }
  root.ZipTools={readZip,writeZip};if(typeof module!=='undefined')module.exports=root.ZipTools;
})(typeof window!=='undefined'?window:globalThis);
