(function(){
 'use strict';
 const $=id=>document.getElementById(id), $$=q=>[...document.querySelectorAll(q)],
 E=window.ConverterEngine,X=window.ExcelTools,DB=window.LocalDB;
 const ROLE_LABELS={groupKey:['Ключ группы','Первый уровень группировки'],locality:['Населённый пункт','Для ветки с пустым комментарием'],comment:['Комментарий','Определяет ветку и подгруппу'],square:['Квадрат','Уникальные значения внутри подгруппы'],targetId:['Номер цели','Уникальные идентификаторы'],targetType:['Характер цели','Уникальные значения, если используются в маске']};
 const state={workbook:null,config:E.templateDefault(),rawRows:[],columns:[],issues:[],result:null,activeView:'import',mode:'preview',dirty:false,draft:'',templates:[],loadedTemplateId:null,lastFocusedMask:null,worker:null,renderSeq:0,persistSession:false};
 let renderTimer,toastTimer,configTimer;
 const escapeHtml=s=>String(s??'').replace(/[&<>"']/g,ch=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[ch]));
 const colDisplay=(c)=>`${X.colLetter(c.index)} · ${c.title}`;
 const getSheet=()=>state.workbook?.sheets?.[state.config.sheetIndex]||null;
 const headerIndex=()=>Math.max(0,(Number(state.config.headerRow)||1)-1);
 function buildColumns(){
   const sheet=getSheet();if(!sheet){state.columns=[];state.rawRows=[];return;}
   const hdr=headerIndex(),rows=sheet.rows;state.rawRows=rows.slice(hdr+1);
   const width=Math.max(rows[hdr]?.length||0,...rows.slice(hdr+1,hdr+24).map(r=>r.length),0);
   state.columns=Array.from({length:Math.min(width,3000)},(_,index)=>({index,letter:X.colLetter(index),title:String(rows[hdr]?.[index]??'').trim()||'Без заголовка'}));
 }
 function toast(message,error=false){const el=$('toast');el.textContent=message;el.classList.toggle('error',error);el.classList.add('visible');clearTimeout(toastTimer);toastTimer=setTimeout(()=>el.classList.remove('visible'),4800);}
 function showView(name){state.activeView=name;$$('.nav-link').forEach(b=>b.classList.toggle('active',b.dataset.view===name));$$('.view').forEach(v=>v.classList.toggle('active',v.id==='view-'+name));if(name==='templates')showTemplates();}
 function setMode(mode){state.mode=mode;$('tabPreview').classList.toggle('active',mode==='preview');$('tabEdit').classList.toggle('active',mode==='edit');$('renderedText').hidden=mode==='edit';$('manualText').hidden=mode!=='edit';if(mode==='edit'){
   if(!state.dirty){state.draft=state.result?.text||'';$('manualText').value=state.draft;}
   $('manualText').focus({preventScroll:true});
 } updateDraftLine();}
 function updateDraftLine(){
   $('editDot').hidden=!state.dirty;
   $('draftLine').textContent=state.dirty?'Есть ручные изменения. При экспорте используется исправленный текст.':'Экспортирует автоматически сформированный текст';
 }
 function saveConfig(){clearTimeout(configTimer);configTimer=setTimeout(()=>DB.put('session','config',state.config).catch(()=>{}),350);}
 function changed(opts={}){saveConfig();if(opts.groups)updateGroupsFromData();scheduleRender();}
 function columnOptions(current){return `<option value="">Выберите столбец…</option>`+state.columns.map(c=>`<option value="${escapeHtml(c.letter)}" ${current===c.letter?'selected':''}>${escapeHtml(colDisplay(c))}</option>`).join('');}
 function drawMapping(){const target=$('mappingFields');target.innerHTML='';for(const [key,[name,desc]] of Object.entries(ROLE_LABELS)){
   const d=document.createElement('div');d.className='mapping-item';d.innerHTML=`<div><strong>${escapeHtml(name)}</strong><small>${escapeHtml(desc)}</small></div><select class="mapping-select" data-role="${key}" aria-label="${escapeHtml(name)}">${columnOptions(state.config.columnMapping[key])}</select>`;target.append(d);
   d.querySelector('select').addEventListener('change',ev=>{state.config.columnMapping[key]=ev.target.value;changed({groups:key==='groupKey'});});
 }}
 function drawSource(){
   const sheet=getSheet();$('fileMeta').hidden=!state.workbook;
   if(!sheet){$('sheetSelect').disabled=true;$('sheetSelect').innerHTML='<option>Сначала загрузите файл</option>';$('sourcePreview').innerHTML='<div class="empty-state">Здесь появятся столбцы и первые строки таблицы.</div>';$('tableSize').textContent='НЕТ ДАННЫХ';$('fileName').textContent='';return;}
   $('fileName').textContent=state.workbook.fileName||'Локальная таблица';
   $('fileDetails').textContent=`${state.workbook.sheets.length} лист(ов) · ${state.rawRows.filter(r=>r.some(c=>String(c??'').trim())).length} строк данных`;
   $('sheetSelect').disabled=false;$('sheetSelect').innerHTML=state.workbook.sheets.map((s,i)=>`<option value="${i}" ${i===state.config.sheetIndex?'selected':''}>${escapeHtml(s.name)}</option>`).join('');
   $('headerRowInput').value=state.config.headerRow;
   $('tableSize').textContent=`${state.columns.length} СТОЛБЦОВ`;
   let html='<table class="data-table"><thead><tr><th>№</th>'+state.columns.slice(0,20).map(c=>`<th>${escapeHtml(colDisplay(c))}</th>`).join('')+'</tr></thead><tbody>';
   state.rawRows.slice(0,9).forEach((r,i)=>{html+=`<tr><th>${headerIndex()+2+i}</th>`+state.columns.slice(0,20).map(c=>`<td title="${escapeHtml(r[c.index]??'')}">${escapeHtml(r[c.index]??'')}</td>`).join('')+'</tr>';});
   $('sourcePreview').innerHTML=html+'</tbody></table>';
 }
 function updateGroupsFromData(){
   const role=state.config.columnMapping.groupKey,idx=X.colIndex(role);if(idx<0||!state.workbook){drawGroups([]);return;}
   const keys=[],seen=new Set();for(const row of state.rawRows){let k=String(row[idx]??'').trim().replace(/\s+/g,' ');if(!k||seen.has(k))continue;seen.add(k);keys.push(k);}
   const old=state.config.groupOrder;state.config.groupOrder=[...old.filter(k=>seen.has(k)),...keys.filter(k=>!old.includes(k))];
   drawGroups(keys);
 }
 function drawGroups(sourceKeys=[]){
   const target=$('groupsList');$('groupCount').textContent=String(sourceKeys.length);
   if(!sourceKeys.length){target.innerHTML='<div class="empty-state">Загрузите таблицу и назначьте ключ группировки.</div>';return;}
   const sorted=state.config.groupOrder.filter(k=>sourceKeys.includes(k));target.innerHTML='';
   sorted.forEach((key,index)=>{
     const custom=state.config.groupTemplates[key]||{};
     const row=document.createElement('div');row.className='group-row';row.dataset.key=key;
     row.innerHTML=`<div class="group-row-head"><span class="group-index">${String(index+1).padStart(2,'0')}</span><strong>${escapeHtml(key)}</strong><div class="small-buttons"><button class="btn btn-secondary btn-sm" data-action="up" data-index="${index}" title="Выше" ${index===0?'disabled':''}>↑</button><button class="btn btn-secondary btn-sm" data-action="down" data-index="${index}" title="Ниже" ${index===sorted.length-1?'disabled':''}>↓</button></div></div><div class="group-row-fields"><label class="field"><span>Префикс группы</span><input data-group-input="prefix" data-index="${index}" value="${escapeHtml(custom.prefix??'')}" placeholder="Общий префикс"></label><label class="field"><span>Свой заголовок (необязательно)</span><input data-group-input="header" data-index="${index}" value="${escapeHtml(custom.header??'')}" placeholder="Общая маска заголовка"></label></div>`;
     target.append(row);
   });
 }
 function drawFilters(){
   $('filterMode').value=state.config.filters.mode||'all';$('filterList').innerHTML='';
   state.config.filters.rules.forEach((rule,i)=>{
     const row=document.createElement('div');row.className='filter-item';row.dataset.index=i;
     const options=columnOptions(rule.column);
     row.innerHTML=`<select class="filter-select" data-filter="column" aria-label="Столбец условия">${options}</select><select class="filter-select" data-filter="op" aria-label="Оператор"><option value="eq">Равно</option><option value="neq">Не равно</option><option value="contains">Содержит</option><option value="notContains">Не содержит</option><option value="empty">Пусто</option><option value="notEmpty">Заполнено</option><option value="gt">Больше</option><option value="lt">Меньше</option></select><input class="filter-input" data-filter="value" value="${escapeHtml(rule.value||'')}" placeholder="Значение" aria-label="Значение условия"><button class="btn btn-quiet btn-sm" data-remove-filter="${i}" aria-label="Удалить условие">×</button>`;
     $('filterList').append(row);row.querySelector('[data-filter="op"]').value=rule.op||'eq';
     row.querySelector('[data-filter="value"]').disabled=['empty','notEmpty'].includes(rule.op);
   });
 }
 function loadForm(){
   const c=state.config;
   $('normTrim').checked=c.normalization.trim!==false;$('normCollapse').checked=c.normalization.collapseSpaces!==false;$('normCase').checked=c.normalization.caseInsensitiveGrouping===true;
   $('groupHeader').value=c.groupHeader;$('defaultPrefix').value=c.defaultPrefix;$('groupSort').value=c.groupSort;
   $('emptyEnabled').checked=c.branches.emptyComment.enabled;$('filledEnabled').checked=c.branches.filledComment.enabled;
   $('emptyMask').value=c.branches.emptyComment.template;$('filledMask').value=c.branches.filledComment.template;
   $('emptySeparator').value=c.branches.emptyComment.separator;$('filledSeparator').value=c.branches.filledComment.separator;
   $('emptySort').value=c.branches.emptyComment.sort;$('filledSort').value=c.branches.filledComment.sort;
   $('stripPrefix').value=c.branches.filledComment.stripPrefix;
   $('branchFirst').value=c.branchOrder[0];$('targetPrefix').value=c.output.targetPrefix;
   $('valueSeparator').value=c.output.valueSeparator;$('branchSeparator').value=c.output.branchSeparator;
   $('groupSeparator').value=({'\n\n':'double','\n':'single','; ':'semicolon'}[c.output.groupSeparator]||'double');$('templateName').value=c.name;
   $('headerRowInput').value=c.headerRow;$('persistSession').checked=state.persistSession;
   drawMapping();drawFilters();updateGroupsFromData();drawSource();
 }
 function syncFile(file,book){
   state.workbook={...book,fileName:file?.name||book.fileName||'Локальная таблица'};
   state.config.sheetIndex=Math.min(Math.max(Number(state.config.sheetIndex)||0,0),state.workbook.sheets.length-1);
   buildColumns();loadForm();changed({groups:true});
   if(state.persistSession)DB.put('session','workbook',state.workbook).catch(()=>toast('Не удалось сохранить таблицу в память iPad.',true));
 }
 async function fileSelected(file){if(!file)return;
   if(state.dirty&&!confirm('При загрузке новой таблицы ручные правки будут удалены. Продолжить?'))return;
   try{
   $('fileName').textContent=file.name;toast('Читаю таблицу…');
   const book=await X.loadFile(file);state.config.sheetIndex=0;state.dirty=false;state.draft='';DB.remove('drafts','current').catch(()=>{});
   syncFile(file,book);toast(`Файл загружен: ${book.sheets.length} лист(ов).`);
   if(book.warnings?.length)toast(book.warnings[0],true);
   showView('mapping');
 }catch(error){toast(error.message||String(error),true);}}
 function scheduleRender(){clearTimeout(renderTimer);renderTimer=setTimeout(runRender,130);}
 function updateStats(){
   const stats=state.result?.stats||{};
   $('statRead').textContent=stats.read??0;$('statGroups').textContent=stats.groups??0;$('statIssues').textContent=state.issues.filter(i=>i.severity!=='info').length;
   $('warningCount').textContent=String(state.issues.length);
 }
 function drawIssues(){const container=$('warningsList');container.innerHTML='';if(!state.issues.length){const p=document.createElement('p');p.className='subtle';p.textContent=state.result?.text?'Ошибок и предупреждений не обнаружено.':'Замечания появятся здесь.';container.append(p);return;}
   for(const item of state.issues.slice(0,22)){const div=document.createElement('div');div.className='issue '+(item.severity||'warning');const ico=document.createElement('span');ico.textContent=item.severity==='error'?'✕':item.severity==='info'?'ℹ':'⚠';const text=document.createElement('span');text.textContent=item.message;div.append(ico,text);container.append(div);}
   if(state.issues.length>22){const p=document.createElement('p');p.className='subtle';p.textContent=`И ещё ${state.issues.length-22} замечаний.`;container.append(p);}
 }
 function applyRender(result){state.result=result;state.issues=result.issues||[];
   const text=result.text||'';$('renderedText').textContent=text||'Нет текста для вывода. Проверьте сопоставление столбцов и правила.';
   $('renderedText').classList.toggle('empty-result',!text);
   const broken=state.issues.some(i=>i.severity==='error');$('renderBadge').textContent=broken?'ПРОВЕРЬТЕ':text?'СФОРМИРОВАН':'НЕТ ДАННЫХ';
   $('renderBadge').classList.toggle('ok',!!text&&!broken);$('renderBadge').classList.toggle('bad',broken);
   if(!state.dirty){state.draft=text;$('manualText').value=text;}
   $('exportTxt').disabled=!text&&!(state.dirty&&state.draft);
   $('exportDocx').disabled=$('exportTxt').disabled;
   updateStats();drawIssues();updateDraftLine();
 }
 function runRender(){
   const valid=state.workbook&&state.rawRows.length&&Object.values(state.config.columnMapping).every(Boolean);
   if(!state.workbook){applyRender({text:'',stats:{},issues:[],groupKeys:[]});return;}
   const seq=++state.renderSeq;
   if(state.worker&&state.rawRows.length>750){
     $('renderBadge').textContent='ОБРАБОТКА';state.worker.postMessage({id:seq,rows:state.rawRows,config:state.config,dataStartRow:headerIndex()+2,maxColumns:state.columns.length});return;
   }
   try{const result=E.render(state.rawRows,state.config,{dataStartRow:headerIndex()+2,maxColumns:state.columns.length});if(seq===state.renderSeq)applyRender(result);}catch(e){toast('Ошибка формирования: '+e.message,true);}
 }
 function download(blob,filename){
   const url=URL.createObjectURL(blob),a=document.createElement('a');a.href=url;a.download=filename;
   document.body.append(a);a.click();a.remove();setTimeout(()=>URL.revokeObjectURL(url),30000);
 }
 function safeName(name){return String(name||'документ').trim().replace(/[\\/:*?"<>|\x00-\x1f]/g,'_').slice(0,75)||'документ';}
 function exportDocument(type){const text=state.dirty?state.draft:state.result?.text;if(!text){toast('Нет текста для экспорта.',true);return;}
   const base=safeName(state.config.name);
   try{if(type==='txt')download(new Blob(['\uFEFF',text],{type:'text/plain;charset=utf-8'}),base+'.txt');
   else download(window.DocxTools.makeDocx(text),base+'.docx');toast('Документ подготовлен для сохранения в «Файлы».');}
   catch(error){toast('Ошибка экспорта: '+error.message,true);}
 }
 function exportTemplate(){const payload={...state.config,version:1};const json=JSON.stringify(payload,null,2);download(new Blob([json],{type:'application/json;charset=utf-8'}),safeName(state.config.name)+'.json');}
 async function loadTemplateFile(file){if(!file)return;try{const obj=JSON.parse(await file.text());if(obj.version!==1||!obj.columnMapping||!obj.branches)throw new Error('Формат шаблона неизвестен или повреждён.');
   state.config=E.ensureTemplate(obj);state.loadedTemplateId=null;buildColumns();loadForm();changed({groups:true});toast('Шаблон импортирован. Проверьте сопоставление столбцов.');}
   catch(e){toast('Не удалось импортировать шаблон: '+e.message,true);}}
 async function saveTemplates(){await DB.put('templates','library',state.templates);}
 async function saveTemplate(){const name=$('templateName').value.trim();if(!name){toast('Укажите название шаблона.',true);return;}
   state.config.name=name;const existing=state.templates.find(t=>t.id===state.loadedTemplateId)||state.templates.find(t=>t.name===name);
   const id=existing?.id||('t-'+Date.now()+'-'+Math.random().toString(36).slice(2,7));
   const entry={id,name,updated:new Date().toISOString(),config:JSON.parse(JSON.stringify(state.config))};
   const at=state.templates.findIndex(t=>t.id===id);if(at>=0)state.templates[at]=entry;else state.templates.unshift(entry);
   state.loadedTemplateId=id;
   try{await saveTemplates();showTemplates();toast('Шаблон сохранён локально. Для резервной копии скачайте JSON.');}
   catch(e){toast('Не удалось сохранить шаблон: '+e.message,true);}}
 function showTemplates(){const target=$('savedTemplates');$('templateCount').textContent=String(state.templates.length);target.innerHTML='';
   if(!state.templates.length){target.innerHTML='<div class="empty-state">Пока нет сохранённых шаблонов.</div>';return;}
   for(const entry of state.templates){const div=document.createElement('div');div.className='saved-item';
     const info=document.createElement('div');info.style.flex='1';info.style.minWidth='0';const name=document.createElement('strong');name.textContent=entry.name;const date=document.createElement('small');date.textContent=new Date(entry.updated).toLocaleString('ru-RU');info.append(name,date);
     const load=document.createElement('button');load.className='btn btn-secondary btn-sm';load.textContent='Открыть';load.onclick=()=>{state.config=E.ensureTemplate(JSON.parse(JSON.stringify(entry.config)));state.loadedTemplateId=entry.id;buildColumns();loadForm();changed({groups:true});showView('mapping');toast('Шаблон открыт.');};
     const del=document.createElement('button');del.className='btn btn-quiet btn-sm';del.textContent='✕';del.title='Удалить шаблон';del.onclick=async()=>{if(!confirm(`Удалить шаблон «${entry.name}»?`))return;state.templates=state.templates.filter(t=>t.id!==entry.id);if(state.loadedTemplateId===entry.id)state.loadedTemplateId=null;try{await saveTemplates();showTemplates();toast('Шаблон удалён.');}catch(e){toast(e.message,true);}};
     div.append(info,load,del);target.append(div);
   }
 }
 function setNetworkStatus(){const p=$('networkState');p.classList.remove('offline');if(!navigator.onLine){p.classList.add('offline');p.querySelector('span').textContent='Без интернета';return;}
   p.querySelector('span').textContent=p.classList.contains('ready')?'Автономный режим готов':'Сеть доступна';}
 async function initServiceWorker(){
   setNetworkStatus();window.addEventListener('online',setNetworkStatus);window.addEventListener('offline',setNetworkStatus);
   if(!('serviceWorker' in navigator)||location.protocol==='file:'){ $('networkState').querySelector('span').textContent='Откройте через HTTPS';return;}
   try{await navigator.serviceWorker.register('./sw.js',{scope:'./'});await navigator.serviceWorker.ready;
     $('networkState').classList.add('ready');setNetworkStatus();
   }catch(e){$('networkState').querySelector('span').textContent='Офлайн не настроен';toast('Сервис-воркер не установлен. Проверьте подключение HTTPS.',true);}
 }
 function bindUI(){
   $$('.nav-link').forEach(b=>b.addEventListener('click',()=>showView(b.dataset.view)));
   $('fileInput').addEventListener('change',e=>{fileSelected(e.target.files[0]);e.target.value='';});
   const drop=$('dropZone');drop.addEventListener('dragover',e=>{e.preventDefault();drop.classList.add('dragging');});
   drop.addEventListener('dragleave',()=>drop.classList.remove('dragging'));
   drop.addEventListener('drop',e=>{e.preventDefault();drop.classList.remove('dragging');fileSelected(e.dataTransfer.files[0]);});
   $('clearFile').addEventListener('click',async()=>{state.workbook=null;state.rawRows=[];state.columns=[];state.dirty=false;state.draft='';DB.remove('drafts','current').catch(()=>{});await DB.remove('session','workbook').catch(()=>{});loadForm();scheduleRender();toast('Таблица удалена из рабочего пространства.');});
   $('sheetSelect').addEventListener('change',e=>{state.config.sheetIndex=Number(e.target.value);buildColumns();loadForm();changed({groups:true});});
   $('headerRowInput').addEventListener('change',e=>{state.config.headerRow=Math.max(1,Math.min(10000,Number(e.target.value)||1));buildColumns();loadForm();changed({groups:true});});
   $('persistSession').addEventListener('change',async e=>{state.persistSession=e.target.checked;await DB.put('session','saveFile',state.persistSession).catch(()=>{});
     if(state.persistSession&&state.workbook)DB.put('session','workbook',state.workbook).then(()=>toast('Таблица сохранена на устройстве.')).catch(()=>toast('Не удалось сохранить таблицу.',true));
     if(!state.persistSession)DB.remove('session','workbook').catch(()=>{});
   });
   const configBinds={normTrim:['normalization','trim','checked'],normCollapse:['normalization','collapseSpaces','checked'],normCase:['normalization','caseInsensitiveGrouping','checked'],
      groupHeader:[null,'groupHeader','value'],defaultPrefix:[null,'defaultPrefix','value'],groupSort:[null,'groupSort','value'],
      emptyEnabled:['branches','emptyComment','checked','enabled'],filledEnabled:['branches','filledComment','checked','enabled'],
      emptyMask:['branches','emptyComment','value','template'],filledMask:['branches','filledComment','value','template'],
      emptySeparator:['branches','emptyComment','value','separator'],filledSeparator:['branches','filledComment','value','separator'],
      emptySort:['branches','emptyComment','value','sort'],filledSort:['branches','filledComment','value','sort'],
      stripPrefix:['branches','filledComment','value','stripPrefix'],targetPrefix:['output','targetPrefix','value'],valueSeparator:['output','valueSeparator','value'],
      branchSeparator:['output','branchSeparator','value']};
   for(const [id,path] of Object.entries(configBinds)){
      $(id).addEventListener(['checked'].includes(path[2])?'change':'input',e=>{
        const val=path[2]==='checked'?e.target.checked:e.target.value;
        if(path[0]==='branches')state.config.branches[path[1]][path[3]]=val;
        else if(path[0])state.config[path[0]][path[1]]=val;
        else state.config[path[1]]=val;
        changed();
      });
   }
   $('groupSeparator').addEventListener('change',e=>{state.config.output.groupSeparator=({double:'\n\n',single:'\n',semicolon:'; '})[e.target.value]||'\n\n';changed();});
   $('branchFirst').addEventListener('change',e=>{const f=e.target.value;state.config.branchOrder=[f,f==='emptyComment'?'filledComment':'emptyComment'];changed();});
   $('groupsList').addEventListener('input',e=>{
     const what=e.target.dataset.groupInput;if(!what)return;
     const i=Number(e.target.dataset.index),key=state.config.groupOrder[i];if(!key)return;
     const obj=state.config.groupTemplates[key]||{};
     if(e.target.value)obj[what]=e.target.value;else delete obj[what];
     if(Object.keys(obj).length)state.config.groupTemplates[key]=obj;else delete state.config.groupTemplates[key];changed();
   });
   $('groupsList').addEventListener('click',e=>{const b=e.target.closest('[data-action]');if(!b)return;
      const i=Number(b.dataset.index),j=i+(b.dataset.action==='up'?-1:1),arr=state.config.groupOrder;
      if(j<0||j>=arr.length)return;[arr[i],arr[j]]=[arr[j],arr[i]];state.config.groupSort='manual';$('groupSort').value='manual';updateGroupsFromData();changed();
   });
   $('addFilter').addEventListener('click',()=>{state.config.filters.rules.push({column:state.columns[0]?.letter||'',op:'eq',value:''});drawFilters();changed();});
   $('filterMode').addEventListener('change',e=>{state.config.filters.mode=e.target.value;changed();});
   $('filterList').addEventListener('input',e=>{
      const el=e.target,field=el.dataset.filter;if(!field)return;
      const index=Number(el.closest('.filter-item').dataset.index);
      state.config.filters.rules[index][field]=el.value;
      if(field==='op')el.closest('.filter-item').querySelector('[data-filter="value"]').disabled=['empty','notEmpty'].includes(el.value);
      changed();
   });
   $('filterList').addEventListener('change',e=>{if(e.target.dataset.filter==='op')e.target.dispatchEvent(new Event('input',{bubbles:true}));});
   $('filterList').addEventListener('click',e=>{const b=e.target.closest('[data-remove-filter]');if(!b)return;state.config.filters.rules.splice(Number(b.dataset.removeFilter),1);drawFilters();changed();});
   $('tokenChips').addEventListener('click',e=>{const btn=e.target.closest('[data-token]');if(!btn)return;const token=btn.dataset.token;
      const target=state.lastFocusedMask||$('emptyMask');const start=target.selectionStart??target.value.length,end=target.selectionEnd??start;
      target.value=target.value.slice(0,start)+token+target.value.slice(end);target.focus();target.setSelectionRange(start+token.length,start+token.length);target.dispatchEvent(new Event('input',{bubbles:true}));
   });
   for(const id of ['emptyMask','filledMask'])$(id).addEventListener('focus',e=>state.lastFocusedMask=e.target);
   $('tabPreview').addEventListener('click',()=>setMode('preview'));
   $('tabEdit').addEventListener('click',()=>setMode('edit'));
   $('manualText').addEventListener('input',e=>{state.dirty=true;state.draft=e.target.value;updateDraftLine();DB.put('drafts','current',state.draft).catch(()=>{});});
   $('regenerateButton').addEventListener('click',()=>{if(state.dirty&&!confirm('Удалить ручные изменения и восстановить текст из правил?'))return;state.dirty=false;DB.remove('drafts','current').catch(()=>{});state.draft=state.result?.text||'';$('manualText').value=state.draft;updateDraftLine();toast('Текст восстановлен из правил.');});
   $('exportTxt').addEventListener('click',()=>exportDocument('txt'));
   $('exportDocx').addEventListener('click',()=>exportDocument('docx'));
   $('saveTemplate').addEventListener('click',saveTemplate);
   $('exportTemplate').addEventListener('click',exportTemplate);
   $('templateInput').addEventListener('change',e=>{loadTemplateFile(e.target.files[0]);e.target.value='';});
   $('deleteLocal').addEventListener('click',async()=>{if(!confirm('Удалить сохранённую таблицу и черновик с устройства? Сохранённые шаблоны останутся.'))return;
     await Promise.all([DB.remove('session','workbook'),DB.remove('drafts','current')].map(p=>p.catch(()=>{})));
     state.workbook=null;state.rawRows=[];state.columns=[];state.dirty=false;state.draft='';state.persistSession=false;await DB.put('session','saveFile',false).catch(()=>{});loadForm();scheduleRender();toast('Локальная таблица и черновик удалены.');
   });
   $('helpButton').addEventListener('click',()=>$('helpDialog').showModal());
   $('closeHelp').addEventListener('click',()=>$('helpDialog').close());
 }
 async function restore(){
   try{const config=await DB.get('session','config');if(config&&config.version===1)state.config=E.ensureTemplate(config);
     state.templates=(await DB.get('templates','library'))||[];state.persistSession=!!(await DB.get('session','saveFile'));
     if(state.persistSession){const book=await DB.get('session','workbook');if(book?.sheets)state.workbook=book;}
   }catch(e){toast('Локальное хранилище недоступно. Шаблоны можно сохранять в JSON.',true);}
   buildColumns();loadForm();showTemplates();runRender();
   if(state.workbook){try{const draft=await DB.get('drafts','current');if(typeof draft==='string'&&draft!==(state.result?.text||'')){state.dirty=true;state.draft=draft;$('manualText').value=draft;updateDraftLine();}}catch(e){}}
 }
 function startWorker(){try{state.worker=new Worker('./worker.js');state.worker.onmessage=e=>{const {id,ok,result,error}=e.data;if(id!==state.renderSeq)return;if(ok)applyRender(result);else{state.worker.terminate();state.worker=null;toast('Переключаюсь на обычную обработку: '+error,true);runRender();}};state.worker.onerror=()=>{state.worker.terminate();state.worker=null;runRender();};}catch(e){state.worker=null;}}
 async function boot(){bindUI();startWorker();$('tokenChips').innerHTML=['{groupKey}','{prefix}','{locality}','{comment}','{square}','{squares}','{targetIds}','{targetTypes}','{targetPrefix}'].map(token=>`<button type="button" class="token-chip" data-token="${token}">${token}</button>`).join('');
   await restore();initServiceWorker();}
 boot();
})();
