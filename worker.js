importScripts('./engine.js');
self.onmessage=event=>{
 const {id,rows,config,dataStartRow,maxColumns}=event.data;
 try{const result=self.ConverterEngine.render(rows,config,{dataStartRow,maxColumns});self.postMessage({id,ok:true,result});}
 catch(error){self.postMessage({id,ok:false,error:String(error.message||error)});}
};
