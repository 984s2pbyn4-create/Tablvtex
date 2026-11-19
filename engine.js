/* Pure converter engine, usable in browser, Web Worker and Node tests. */
(function(root){
 'use strict';
 const ROLES=['groupKey','locality','comment','square','targetId','targetType'];
 const TOKENS=new Set(['groupKey','prefix','locality','comment','square','squares','targetId','targetIds','targetType','targetTypes','targetPrefix']);
 const norm=v=>String(v??'').trim().replace(/\s+/g,' ');
 const stableUnique=values=>[...new Set(values.map(norm).filter(Boolean))];
 const colIndex=v=>{if(Number.isInteger(v))return v;if(typeof v==='number')return -1;const x=String(v??'').trim().toUpperCase();if(/^\d+$/.test(x))return Number(x);if(!/^[A-Z]+$/.test(x))return -1;let n=0;for(const c of x)n=n*26+c.charCodeAt(0)-64;return n-1;};
 function templateDefault(){return {
   version:1,name:'Новый шаблон',columnMapping:Object.fromEntries(ROLES.map(x=>[x,''])),
   headerRow:1,sheetIndex:0,
   filters:{mode:'all',rules:[]},
   normalization:{trim:true,collapseSpaces:true,caseInsensitiveGrouping:false},
   groupOrder:[],groupSort:'source',unknownGroups:'append',groupTemplates:Object.create(null),
   groupHeader:'{prefix} {groupKey}',defaultPrefix:'',branchOrder:['emptyComment','filledComment'],
   branches:{
     emptyComment:{enabled:true,template:'в районе н.п. {locality} ({squares}) — {targetPrefix}{targetIds}',separator:', ',uniqueSquares:true,uniqueTypes:true,uniqueIds:true,sort:'source'},
     filledComment:{enabled:true,template:'«{comment}» ({square}) — {targetPrefix}{targetIds}',separator:', ',stripPrefix:'',groupBy:['comment','square'],uniqueIds:true,sort:'source'}
   },
   output:{groupSeparator:'\n\n',branchSeparator:'; ',itemSeparator:', ',valueSeparator:', ',targetPrefix:'Ц'},
   validation:{missingGroup:'warn-and-skip',missingId:'warn-and-skip',missingLocality:'warn-and-skip',missingSquare:'warn',conflicts:'warn'}
 };}
 function ensureTemplate(data){
   const d=templateDefault(),a=data&&typeof data==='object'?data:{};
   const obj={...d,...a,columnMapping:{...d.columnMapping,...a.columnMapping},
     filters:{...d.filters,...a.filters},normalization:{...d.normalization,...a.normalization},
     branches:{emptyComment:{...d.branches.emptyComment,...a.branches?.emptyComment},filledComment:{...d.branches.filledComment,...a.branches?.filledComment}},
     output:{...d.output,...a.output},validation:{...d.validation,...a.validation},groupTemplates:Object.assign(Object.create(null),a.groupTemplates||{})};
   obj.branchOrder=Array.isArray(a.branchOrder)?a.branchOrder.filter(x=>x==='emptyComment'||x==='filledComment'):d.branchOrder;
   if(obj.branchOrder.length!==2)obj.branchOrder=d.branchOrder;
   if(!Array.isArray(obj.groupOrder))obj.groupOrder=[];
   if(!Array.isArray(obj.filters.rules))obj.filters.rules=[];
   return obj;
 }
 function normalizeValue(value,config){let s=String(value??'');if(config.trim!==false)s=s.trim();if(config.collapseSpaces!==false)s=s.replace(/\s+/g,' ');return s;}
 function removePrefix(text,prefix){if(!prefix)return text;const t=norm(text),p=norm(prefix);return t.startsWith(p)?norm(t.slice(p.length)):t;}
 function matchRule(row,rule){
   const a=String(row.original?.[colIndex(rule.column)]??'').trim(),b=String(rule.value??'').trim();
   switch(rule.op){case 'empty':return a==='';case 'notEmpty':return a!=='';case 'eq':return a.toLocaleLowerCase()===b.toLocaleLowerCase();case 'neq':return a.toLocaleLowerCase()!==b.toLocaleLowerCase();
    case 'contains':return a.toLocaleLowerCase().includes(b.toLocaleLowerCase());case 'notContains':return !a.toLocaleLowerCase().includes(b.toLocaleLowerCase());
    case 'gt':return a!==''&&b!==''&&Number(a)>Number(b);case 'lt':return a!==''&&b!==''&&Number(a)<Number(b);default:return true;}
 }
 function maskVars(mask){return [...String(mask||'').matchAll(/\{([^{}]+)\}/g)].map(m=>m[1]);}
 function interpolate(mask,vars){return String(mask||'').replace(/\{([^{}]+)\}/g,(match,k)=>Object.prototype.hasOwnProperty.call(vars,k)?String(vars[k]??''):match).trim().replace(/ {2,}/g,' ');}
 function checkMask(mask,allowed,where,issues){
   for(const t of maskVars(mask))if(!allowed.has(t))issues.push({severity:'error',code:'UNKNOWN_TOKEN',message:`${where}: неизвестная переменная {${t}}.`});
 }
 function keyFor(v,conf){const x=norm(v);return conf.normalization.caseInsensitiveGrouping?x.toLocaleLowerCase():x;}
 function sortedBuckets(entries,kind){const arr=[...entries];if(kind==='alphaAsc')arr.sort((a,b)=>a.label.localeCompare(b.label,'ru',{numeric:true}));if(kind==='alphaDesc')arr.sort((a,b)=>b.label.localeCompare(a.label,'ru',{numeric:true}));return arr;}
 function firstBy(rows,field){return rows[0]?.[field]||'';}
 function render(rawRows,config,options={}){
   const cfg=ensureTemplate(config),issues=[],stats={read:0,kept:0,filtered:0,groups:0,items:0,skipped:0};
   const mappings=Object.fromEntries(ROLES.map(k=>[k,colIndex(cfg.columnMapping[k])]));
   for(const role of ROLES){if(mappings[role]<0)issues.push({severity:'error',code:'MAPPING',message:`Не указан столбец для роли ${role}.`});}
   const used=new Map();for(const role of ROLES){const idx=mappings[role];if(idx<0)continue;if(used.has(idx))issues.push({severity:'warning',code:'DUPLICATE_MAPPING',message:`Роли ${used.get(idx)} и ${role} используют один исходный столбец.`});else used.set(idx,role);}
   const maxCols=Number(options.maxColumns||0);if(maxCols>0)for(const role of ROLES){if(mappings[role]>=maxCols)issues.push({severity:'error',code:'MAPPING_OUT_OF_RANGE',message:`Роль ${role} указывает на отсутствующий столбец.`});}
   checkMask(cfg.groupHeader,TOKENS,'Заголовок группы',issues);
   for(const [key,group] of Object.entries(cfg.groupTemplates))checkMask(group.header??cfg.groupHeader,TOKENS,`Заголовок «${key}»`,issues);
   for(const branch of ['emptyComment','filledComment']){
     const mask=cfg.branches[branch].template;checkMask(mask,TOKENS,`Ветка ${branch}`,issues);
     if(cfg.branches[branch].enabled){
       const marker=branch==='emptyComment'?'{locality}':'{comment}';
       const ids=mask.indexOf('{targetIds}');const single=mask.indexOf('{targetId}');const pos=ids>=0?ids:single;
       if(pos<0)issues.push({severity:'error',code:'MISSING_IDS_TOKEN',message:`Ветка ${branch}: добавьте {targetIds} в маску.`});
       else if(mask.indexOf(marker)<0||pos<mask.indexOf(marker))issues.push({severity:'error',code:'IDS_ORDER',message:`Ветка ${branch}: идентификаторы должны идти после ${marker}.`});
       if(branch==='filledComment'&&mask.includes('{locality}'))issues.push({severity:'error',code:'LOCALITY_IN_FILLED',message:'В заполненной ветке населённый пункт не выводится. Уберите {locality}.'});
     }
   }
   if(issues.some(i=>i.severity==='error'))return {text:'',issues,stats,groupKeys:[]};
   const data=[];const rules=cfg.filters.rules.filter(r=>r.column!=null&&r.op);
   for(let i=0;i<rawRows.length;i++){
     const orig=rawRows[i]||[];if(orig.every(c=>norm(c)===''))continue;stats.read++;
     const r={original:orig,rowIndex:i+Number(options.dataStartRow||2)};
     for(const k of ROLES)r[k]=normalizeValue(orig[mappings[k]],cfg.normalization);
     if(r.comment!=='')r.cleanComment=removePrefix(r.comment,cfg.branches.filledComment.stripPrefix);else r.cleanComment='';
     if(!r.groupKey){issues.push({severity:'warning',code:'EMPTY_GROUP',row:r.rowIndex,message:`Строка ${r.rowIndex}: отсутствует ключ группы, запись пропущена.`});stats.skipped++;continue;}
     if(!r.targetId){issues.push({severity:'warning',code:'EMPTY_ID',row:r.rowIndex,message:`Строка ${r.rowIndex}: отсутствует идентификатор цели, запись пропущена.`});stats.skipped++;continue;}
     if(r.comment===''&&!r.locality){issues.push({severity:'warning',code:'EMPTY_LOCALITY',row:r.rowIndex,message:`Строка ${r.rowIndex}: нет населённого пункта для ветки с пустым комментарием.`});stats.skipped++;continue;}
     if(!r.square)issues.push({severity:'warning',code:'EMPTY_SQUARE',row:r.rowIndex,message:`Строка ${r.rowIndex}: квадрат не заполнен.`});
     if(r.comment!==''&&!r.cleanComment)issues.push({severity:'warning',code:'STRIPPED_COMMENT',row:r.rowIndex,message:`Строка ${r.rowIndex}: после удаления префикса комментарий пуст.`});
     if(rules.length){const outcomes=rules.map(rule=>matchRule(r,rule)),ok=cfg.filters.mode==='any'?outcomes.some(Boolean):outcomes.every(Boolean);if(!ok){stats.filtered++;continue;}}
     data.push(r);
   }
   stats.kept=data.length;
   const groupMap=new Map();for(const r of data){const k=keyFor(r.groupKey,cfg);if(!groupMap.has(k))groupMap.set(k,{key:r.groupKey,rows:[]});groupMap.get(k).rows.push(r);}
   let groupKeys=[...groupMap.values()].map(g=>g.key);
   if(cfg.groupSort==='alphaAsc')groupKeys.sort((a,b)=>a.localeCompare(b,'ru',{numeric:true}));
   if(cfg.groupSort==='alphaDesc')groupKeys.sort((a,b)=>b.localeCompare(a,'ru',{numeric:true}));
   if(cfg.groupSort==='manual'&&cfg.groupOrder.length){const rank=new Map(cfg.groupOrder.map((v,i)=>[keyFor(v,cfg),i]));groupKeys.sort((a,b)=>(rank.get(keyFor(a,cfg))??Infinity)-(rank.get(keyFor(b,cfg))??Infinity));}
   const paragraphs=[];
   for(const gk of groupKeys){
     const rows=groupMap.get(keyFor(gk,cfg)).rows,parts=[];
     for(const branchName of cfg.branchOrder){
       const branch=cfg.branches[branchName];if(!branch.enabled)continue;
       const branchRows=rows.filter(r=>branchName==='emptyComment'?r.comment==='':r.comment!=='');
       if(!branchRows.length)continue;
       const buckets=new Map();
       for(const r of branchRows){
         const k=branchName==='emptyComment'?keyFor(r.locality,cfg):JSON.stringify([keyFor(r.cleanComment,cfg),keyFor(r.square,cfg)]);
         if(!buckets.has(k))buckets.set(k,{label:branchName==='emptyComment'?r.locality:`${r.cleanComment} ${r.square}`,rows:[]});
         buckets.get(k).rows.push(r);
       }
       const items=[];
       for(const bucket of sortedBuckets(buckets.values(),branch.sort)){
         const rr=bucket.rows,ids=stableUnique(rr.map(r=>r.targetId)),squares=stableUnique(rr.map(r=>r.square)),types=stableUnique(rr.map(r=>r.targetType));
         const common={groupKey:gk,prefix:'',locality:firstBy(rr,'locality'),comment:firstBy(rr,'cleanComment'),square:firstBy(rr,'square'),squares:squares.join(cfg.output.valueSeparator),targetId:ids[0]||'',targetIds:ids.join(cfg.output.valueSeparator),targetType:types[0]||'',targetTypes:types.join(cfg.output.valueSeparator),targetPrefix:cfg.output.targetPrefix};
         const s=interpolate(branch.template,common);if(s){items.push(s);stats.items++;}
         if(ids.length<rr.length)issues.push({severity:'info',code:'DUPLICATE_ID',message:`Группа «${gk}»: повторяющиеся идентификаторы объединены.`});
         const typeById=new Map();for(const r of rr){const prev=typeById.get(r.targetId);if(prev&&r.targetType&&prev!==r.targetType)issues.push({severity:'warning',code:'CONFLICT',row:r.rowIndex,message:`Строка ${r.rowIndex}: у одного идентификатора обнаружены разные характеры цели.`});else if(r.targetType)typeById.set(r.targetId,r.targetType);}
       }
       if(items.length)parts.push(items.join(branch.separator??cfg.output.itemSeparator));
     }
     if(!parts.length)continue;
     const custom=cfg.groupTemplates[gk]||{},head=interpolate(custom.header??cfg.groupHeader,{groupKey:gk,prefix:custom.prefix??cfg.defaultPrefix});
     paragraphs.push([head,parts.join(cfg.output.branchSeparator)].filter(Boolean).join(' '));stats.groups++;
   }
   if(stats.kept===0&&stats.read>0)issues.push({severity:'warning',code:'NO_DATA',message:'После проверки и фильтрации не осталось записей для вывода.'});
   return {text:paragraphs.join(cfg.output.groupSeparator),issues:issues.slice(0,400),stats,groupKeys};
 }
 root.ConverterEngine={render,templateDefault,ensureTemplate,maskVars,stableUnique,removePrefix};
 if(typeof module!=='undefined')module.exports=root.ConverterEngine;
})(typeof self!=='undefined'?self:globalThis);
