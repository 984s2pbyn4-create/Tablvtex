/* XLSX (OOXML) and CSV/TSV import without remote CDNs. Files stay on-device. */
(function(root){
 'use strict';
 const {readZip}=root.ZipTools;
 const xml=s=>{const d=new DOMParser().parseFromString(s,'application/xml');if(d.getElementsByTagName('parsererror').length)throw new Error('Ошибка структуры XML в файле Excel.');return d;};
 const localName=e=>e.localName||e.nodeName.split(':').pop();
 const elements=(node,name)=>[...node.getElementsByTagName('*')].filter(n=>localName(n)===name);
 const child=(node,name)=>[...node.children].find(n=>localName(n)===name);
 const str=n=>n?n.textContent||'':'';
 const colIndex=ref=>{let n=0;for(const c of (String(ref||'').match(/^[A-Z]+/i)||[''])[0].toUpperCase())n=n*26+c.charCodeAt(0)-64;return n-1;};
 const colLetter=index=>{let out='';for(let n=index+1;n>0;n=Math.floor((n-1)/26))out=String.fromCharCode(65+(n-1)%26)+out;return out;};
 function parseCSV(text){
   text=text.replace(/^\uFEFF/,'');const first=text.split(/\r?\n/,8).filter(Boolean).slice(0,5).join('\n');
   const count=sep=>(first.match(new RegExp(sep==='\t'?'\t':`\\${sep}`,'g'))||[]).length;
   const sep=[';',',','\t'].sort((a,b)=>count(b)-count(a))[0];
   const out=[];let row=[],val='',inQuote=false;
   for(let i=0;i<text.length;i++){
     let c=text[i];if(c==='"'){if(inQuote&&text[i+1]==='"'){val+='"';i++;}else inQuote=!inQuote;continue;}
     if(!inQuote&&(c===sep||c==='\n'||c==='\r')){
       if(c===sep){row.push(val);val='';continue;}
       if(c==='\r'&&text[i+1]==='\n')i++;
       row.push(val);out.push(row);row=[];val='';continue;
     }val+=c;
   }
   row.push(val);if(row.some(x=>x!==''))out.push(row);
   return {sheets:[{name:'Таблица',rows:out}],warnings:[]};
 }
 function formatNumber(raw,style,formats){
    if(!style)return raw;
    const fmt=formats[style];if(!fmt)return raw;
    const clean=fmt.replace(/"[^"]*"|\\.|\[[^\]]*\]/g,'');
    const simple=clean.match(/^(0{2,})(?:;.*)?$/);
    if(simple&&Number.isFinite(Number(raw)))return String(Math.trunc(Number(raw))).padStart(simple[1].length,'0');
    return raw;
 }
 async function parseXLSX(buffer){
   const zip=readZip(buffer);
   const bookText=await zip.text('xl/workbook.xml');if(!bookText)throw new Error('В XLSX отсутствует xl/workbook.xml.');
   const doc=xml(bookText),relsDoc=xml(await zip.text('xl/_rels/workbook.xml.rels')||'<Relationships/>');
   const rels=new Map(elements(relsDoc,'Relationship').map(e=>[e.getAttribute('Id'),e.getAttribute('Target')]));
   const sheetDefs=elements(doc,'sheet');const sharedDoc=await zip.text('xl/sharedStrings.xml');const shared=[];
   if(sharedDoc){for(const si of elements(xml(sharedDoc),'si'))shared.push(elements(si,'t').map(str).join(''));}
   const stylesText=await zip.text('xl/styles.xml');const numFormats={};let styleFormats=[];
   if(stylesText){
     const styles=xml(stylesText);for(const fmt of elements(styles,'numFmt'))numFormats[fmt.getAttribute('numFmtId')]=fmt.getAttribute('formatCode')||'';
     const xfs=elements(styles,'cellXfs')[0];if(xfs)styleFormats=[...xfs.children].filter(e=>localName(e)==='xf').map(e=>numFormats[e.getAttribute('numFmtId')]||'');
   }
   const sheets=[],warnings=[];
   function toSheetPath(target){
     if(!target)return null;if(target.startsWith('/'))return target.replace(/^\//,'');
     const parts=['xl',...target.split('/')],stack=[];
     for(const part of parts){if(part==='..')stack.pop();else if(part&&part!=='.')stack.push(part);}
     return stack.join('/');
   }
   for(const [idx,def] of sheetDefs.entries()){
     const id=def.getAttribute('r:id')||def.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships','id');
     const relTarget=rels.get(id),path=toSheetPath(relTarget)||`xl/worksheets/sheet${idx+1}.xml`;
     const sheetText=await zip.text(path);if(!sheetText){warnings.push(`Не найден лист «${def.getAttribute('name')}».`);continue;}
     const sheetDoc=xml(sheetText),rows=[];
     for(const r of elements(sheetDoc,'row')){
       const rowIndex=Number(r.getAttribute('r')||rows.length+1)-1;
       if(rowIndex>200000)throw new Error('Слишком много строк. Максимум 200 000.');
       const values=[];
       for(const cell of [...r.children].filter(e=>localName(e)==='c')){
         const at=colIndex(cell.getAttribute('r'));
         if(at<0||at>3000)continue;
         const type=cell.getAttribute('t'),value=str(child(cell,'v'));
         let v='';
         if(type==='s')v=shared[Number(value)]??'';
         else if(type==='inlineStr')v=elements(cell,'t').map(str).join('');
         else if(type==='b')v=value==='1'?'TRUE':'FALSE';
         else if(type==='e')v=value;
         else v=formatNumber(value,Number(cell.getAttribute('s')||0),styleFormats);
         values[at]=v;
       }
       rows[rowIndex]=values;
     }
     sheets.push({name:def.getAttribute('name')||`Лист ${idx+1}`,rows:rows.map(r=>r||[])});
   }
   if(!sheets.length)throw new Error('Не удалось обнаружить доступные листы XLSX.');
   return {sheets,warnings};
 }
 async function loadFile(file){
   if(file.size>35*1024*1024)throw new Error('Файл превышает допустимые 35 МБ.');
   const name=file.name.toLowerCase();
   if(/\.(csv|tsv)$/.test(name))return parseCSV(await file.text());
   if(/\.(xlsx|xlsm)$/.test(name))return parseXLSX(await file.arrayBuffer());
   if(name.endsWith('.xls'))throw new Error('Старый двоичный формат .xls не поддерживается. Сохраните файл как .xlsx или CSV.');
   throw new Error('Поддерживаются XLSX, XLSM, CSV и TSV.');
 }
 root.ExcelTools={loadFile,parseCSV,parseXLSX,colIndex,colLetter};
 if(typeof module!=='undefined')module.exports=root.ExcelTools;
})(typeof window!=='undefined'?window:globalThis);
