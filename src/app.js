// Check library availability
(function(){
    var m=[];
    if(typeof JSZip==='undefined')m.push('JSZip');
    if(typeof cytoscape==='undefined')m.push('Cytoscape');
    if(m.length){document.getElementById('libMissing').classList.add('visible');document.getElementById('dropZone').style.pointerEvents='none';document.getElementById('dropZone').style.opacity='0.4';}
})();

const MAX_FILE_SIZE=150*1024*1024;
const PROMPT_TEMPLATES={
    analyze:'Analyze these Power Query M scripts. For each query:\n1. Identify what data source it connects to\n2. Map all dependencies between queries\n3. Suggest which queries could be consolidated\n4. Identify any circular dependencies or issues\n\nHere are the queries:\n\n',
    optimize:'Review these Power Query M scripts for performance optimizations:\n1. Identify any inefficient patterns (e.g., multiple source reads, unnecessary type conversions)\n2. Suggest query folding opportunities\n3. Recommend step reordering for better performance\n4. Flag any operations that might cause full data loads\n\nHere are the queries:\n\n',
    document:'Generate documentation for these Power Query M scripts:\n1. Create a summary of each query\'s purpose\n2. Document the data flow from sources to final outputs\n3. List all parameters and their expected values\n4. Create a dependency diagram in Mermaid format\n\nHere are the queries:\n\n',
    errors:'Review these Power Query M scripts for potential errors and issues:\n1. Check for hardcoded values that should be parameters\n2. Identify missing error handling\n3. Flag potential null/empty value issues\n4. Check for type mismatches\n5. Identify queries that might fail with data changes\n\nHere are the queries:\n\n'
};
const FILE_COLORS=['#60c0a0','#4c86c8','#7a67c7','#c88a36','#c65364','#3e9b6c','#6f7f95','#9c6f4f','#5f8f8a','#8d6fae'];
let appState={files:[],queries:[],errors:[],selectedFiles:[],cyInstance:null,worksheets:[],activeSheet:null,dataProfile:null};

function escapeHtml(t){return typeof t!=='string'?'':t.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;').replace(/'/g,'&#39;');}

// ═══ MS-QDEFF DataMashup Parser ═══
function readU32(b,o){return(b[o]|(b[o+1]<<8)|(b[o+2]<<16)|(b[o+3]<<24))>>>0;}

async function parseXlsxFile(file){
    const zip=await JSZip.loadAsync(file);
    const queries=[],seen=new Set(),errors=[];
    function addUnique(nq){for(const q of nq){const k=q.fileName+'::'+q.name;if(!seen.has(k)){seen.add(k);queries.push(q);}}}
    for(const[p,ze]of Object.entries(zip.files)){
        if(p.startsWith('customXml/')&&p.endsWith('.bin')){
            try{addUnique(await extractFromDataMashup(await ze.async('arraybuffer'),file.name));}
            catch(e){errors.push(p+': '+e.message);}
        }
    }
    for(const[p,ze]of Object.entries(zip.files)){
        if(p.match(/customXml\/item\d*\.xml$/i)){
            try{
                const xml=decodeXmlBytes(await ze.async('arraybuffer'));
                addUnique(await extractFromCustomXml(xml,file.name));
            }catch(e){}
        }
    }
    if(queries.length===0&&zip.files['xl/connections.xml']){
        try{addUnique(parseConnectionsXml(await zip.files['xl/connections.xml'].async('string'),file.name));}
        catch(e){errors.push('connections.xml: '+e.message);}
    }
    return{queries,errors};
}

async function parsePbixFile(file){
    const zip=await JSZip.loadAsync(file);
    const queries=[],seen=new Set(),errors=[];
    function addUnique(nq){for(const q of nq){const k=q.fileName+'::'+q.name;if(!seen.has(k)){seen.add(k);queries.push(q);}}}
    // Path 1: DataMashup as a top-level file (MS-QDEFF binary) — works for .pbit and older .pbix
    const dmEntry=zip.files['DataMashup']||zip.files['datamashup'];
    if(dmEntry){
        try{addUnique(await extractFromDataMashup(await dmEntry.async('arraybuffer'),file.name));}
        catch(e){errors.push('DataMashup: '+e.message);}
    }
    // Path 2: DataModelSchema JSON — works for .pbit and enhanced metadata .pbix
    if(!queries.length){
        const schemaEntry=zip.files['DataModelSchema'];
        if(schemaEntry){
            try{addUnique(extractFromDataModelSchema(await schemaEntry.async('arraybuffer'),file.name));}
            catch(e){errors.push('DataModelSchema: '+e.message);}
        }
    }
    // Path 3: DataModel (XPress9-compressed ABF) — decompress and extract M code
    if(!queries.length&&zip.files['DataModel']){
        if(typeof extractFromDataModel==='function'){
            try{
                const dmData=await zip.files['DataModel'].async('arraybuffer');
                const dmQueries=await extractFromDataModel(dmData,file.name);
                addUnique(dmQueries);
            }catch(e){errors.push('DataModel: '+e.message);}
        }else{
            errors.push('DataModel decoder not available. This file uses compressed V3 format.');
        }
    }
    return{queries,errors};
}

function extractFromDataModelSchema(buf,fn){
    const q=[];
    // DataModelSchema is UTF-16LE JSON
    let raw=new Uint8Array(buf instanceof ArrayBuffer?buf:buf.buffer||buf);
    let text='';
    try{text=new TextDecoder('utf-16le').decode(raw).replace(/^\uFEFF/,'');}
    catch(e){text=new TextDecoder('utf-8').decode(raw).replace(/^\uFEFF/,'');}
    const model=JSON.parse(text);
    const tables=(model.model||model).tables||[];
    const seen=new Set();
    // Extract from table partitions (source type 'm')
    for(const t of tables){
        for(const p of(t.partitions||[])){
            const src=p.source||{};
            if(src.type!=='m')continue;
            let expr=src.expression;
            if(Array.isArray(expr))expr=expr.join('\n');
            if(!expr)continue;
            const name=t.name;
            if(seen.has(name))continue;seen.add(name);
            q.push({name,code:expr.trim(),fileName:fn,dependencies:findDeps(stripCS(expr),name),externalRefs:findExternalFileRefs(expr)});
        }
    }
    // Extract from top-level expressions (parameters, functions, standalone queries)
    for(const e of((model.model||model).expressions||[])){
        let expr=e.expression;
        if(Array.isArray(expr))expr=expr.join('\n');
        if(!expr)continue;
        const name=e.name;
        if(seen.has(name))continue;seen.add(name);
        q.push({name,code:expr.trim(),fileName:fn,dependencies:findDeps(stripCS(expr),name),externalRefs:findExternalFileRefs(expr)});
    }
    return q;
}

function isPbixFile(name){return/\.pbi[xt]$/i.test(name);}
function isSupportedFile(name){return/\.(xlsx|pbi[xt])$/i.test(name);}

async function extractFromDataMashup(buf,fn){
    const q=[],b=new Uint8Array(buf);
    if(b.length<8)return q;
    // Structured MS-QDEFF: version(4) + pkgLen(4) + ZIP(pkgLen)
    try{const v=readU32(b,0),n=readU32(b,4);
        if(v===0&&n>0&&n<=b.length-8){q.push(...await extractMCodeFromZip(await JSZip.loadAsync(b.slice(8,8+n)),fn));}
    }catch(e){}
    // Fallback: scan for PK signature
    if(!q.length){const s=findPK(b);if(s>=0){try{q.push(...await extractMCodeFromZip(await JSZip.loadAsync(b.slice(s)),fn));}catch(e){}}}
    return q;
}

function findPK(b){for(let i=0;i<b.length-4;i++){if(b[i]===0x50&&b[i+1]===0x4B&&b[i+2]===0x03&&b[i+3]===0x04)return i;}return-1;}

async function extractMCodeFromZip(z,fn){
    const q=[],seen=new Set();
    for(const[p,e]of Object.entries(z.files)){
        if(p.startsWith('Formulas/')&&p.endsWith('.m')&&!e.dir){
            try{for(const x of parseMCodeFile(await e.async('string'),fn)){const k=x.fileName+'::'+x.name;if(!seen.has(k)){seen.add(k);q.push(x);}}}catch(e){}
        }
    }
    return q;
}

async function extractFromCustomXml(xml,fn){
    const q=[],sx=(xml||'').replace(/\u0000/g,'');
    const mm=sx.match(/<(?:\w+:)?DataMashup\b[^>]*>([\s\S]*?)<\/(?:\w+:)?DataMashup>/i);
    if(mm){try{q.push(...await extractFromDataMashup(b64ToBytes(mm[1]).buffer,fn));}catch(e){}}
    if(!q.length){const ms=sx.match(/[A-Za-z0-9+/]{200,}={0,2}/g);
        if(ms)for(const b of ms){try{q.push(...await extractFromDataMashup(b64ToBytes(b).buffer,fn));}catch(e){}}
    }
    return q;
}

function decodeXmlBytes(buf){
    const b=buf instanceof Uint8Array?buf:new Uint8Array(buf);
    if(!b.length)return'';
    try{
        if(b.length>=2&&b[0]===0xFF&&b[1]===0xFE)return new TextDecoder('utf-16le').decode(b).replace(/^\uFEFF/,'');
        if(b.length>=2&&b[0]===0xFE&&b[1]===0xFF)return new TextDecoder('utf-16be').decode(b).replace(/^\uFEFF/,'');
        if(b.length>=3&&b[0]===0xEF&&b[1]===0xBB&&b[2]===0xBF)return new TextDecoder('utf-8').decode(b).replace(/^\uFEFF/,'');
        let evenNul=0,oddNul=0,sample=Math.min(b.length,1024);
        for(let i=0;i<sample;i++)if(b[i]===0){if(i%2)oddNul++;else evenNul++;}
        if(oddNul>sample/10)return new TextDecoder('utf-16le').decode(b).replace(/^\uFEFF/,'');
        if(evenNul>sample/10)return new TextDecoder('utf-16be').decode(b).replace(/^\uFEFF/,'');
    }catch(e){}
    return new TextDecoder('utf-8').decode(b).replace(/^\uFEFF/,'');
}

function b64ToBytes(s){
    let x=(s||'').replace(/[^A-Za-z0-9+/=]/g,'');
    if(!x.length)return new Uint8Array(0);
    const r=x.length%4;
    if(r)x+='='.repeat(4-r);
    const d=atob(x),b=new Uint8Array(d.length);
    for(let i=0;i<d.length;i++)b[i]=d.charCodeAt(i);
    return b;
}

function parseConnectionsXml(xml,fn){
    const q=[],d=new DOMParser().parseFromString(xml,'text/xml');
    for(const c of d.getElementsByTagName('connection')){
        const n=c.getAttribute('name');
        if(n&&n.startsWith('Query -'))q.push({name:n.replace('Query - ',''),code:'// Query code not available \u2014 connections.xml only',fileName:fn,dependencies:[]});
    }
    return q;
}

// ═══ M Code Parser ═══
function stripCS(code){
    let r='',i=0;
    while(i<code.length){
        if(code[i]==='#'&&code[i+1]==='"'){r+='#"';i+=2;while(i<code.length){if(code[i]==='"'&&code[i+1]==='"'){r+='""';i+=2;}else if(code[i]==='"'){r+='"';i++;break;}else{r+=code[i];i++;}}continue;}
        if(code[i]==='"'){r+=' ';i++;while(i<code.length){if(code[i]==='"'&&code[i+1]==='"'){r+='  ';i+=2;}else if(code[i]==='"'){r+=' ';i++;break;}else{r+=code[i]==='\n'?'\n':' ';i++;}}continue;}
        if(code[i]==='/'&&code[i+1]==='*'){r+='  ';i+=2;while(i<code.length){if(code[i]==='*'&&code[i+1]==='/'){r+='  ';i+=2;break;}r+=code[i]==='\n'?'\n':' ';i++;}continue;}
        if(code[i]==='/'&&code[i+1]==='/'){while(i<code.length&&code[i]!=='\n'){r+=' ';i++;}continue;}
        r+=code[i];i++;
    }
    return r;
}

function parseMCodeFile(mCode,fn){
    const q=[];
    let code=mCode.replace(/^\uFEFF/,'');
    const stripped=stripCS(code);
    const re=/\bshared\s+(?:#"(?:[^"]*(?:""[^"]*)*)"|[\w_][\w_]*)\s*=/gi;
    const pos=[];let m;
    while((m=re.exec(stripped))!==null)pos.push(m.index);

    if(pos.length>0){
        for(let i=0;i<pos.length;i++){
            const s=pos[i],e=i+1<pos.length?pos[i+1]:code.length;
            const chunk=code.substring(s,e),sc=stripped.substring(s,e);
            const nm=chunk.match(/^shared\s+(#"(?:[^"]*(?:""[^"]*)*)"|[\w_][\w_]*)\s*=/i);
            if(!nm)continue;
            let name=nm[1];
            if(name.startsWith('#"'))name=name.slice(2,-1).replace(/""/g,'"');
            const eq=chunk.indexOf('=',chunk.indexOf(nm[1])+nm[1].length);
            const sb=sc.substring(eq+1),ls=sb.lastIndexOf(';');
            const qc=ls!==-1?chunk.substring(eq+1,(eq+1)+ls).trim():chunk.substring(eq+1).trim();
            q.push({name,code:qc,fileName:fn,dependencies:findDeps(stripCS(qc),name),externalRefs:findExternalFileRefs(qc)});
        }
    }
    if(!q.length&&/\blet\b/i.test(stripped)&&/\bin\b/i.test(stripped))
        q.push({name:'Query1',code:mCode.trim(),fileName:fn,dependencies:findDeps(stripped,'Query1'),externalRefs:findExternalFileRefs(mCode)});
    return q;
}

const MKW=new Set(['let','in','if','then','else','true','false','null','and','or','not','each','error','try','otherwise','type','is','as','meta','section','shared','Table','List','Record','Text','Number','Date','DateTime','DateTimeZone','Duration','Time','Function','Binary','Csv','Json','Xml','Excel','Sql','OData','Web','File','Folder','Lines','Splitter','Comparer','Combiner','Replacer','Expression','Error','Value','Type','Action','Uri','Byte','Currency','Percentage','Int8','Int16','Int32','Int64','Single','Double','Decimal','Logical','Access','ActiveDirectory','AdobeAnalytics','AdoDotNet','AnalysisServices','AzureStorage','DB2','Cube','Diagnostics','Exchange','Facebook','GoogleAnalytics','Hdfs','Informix','MySQL','Odbc','OleDb','Oracle','PDF','PostgreSQL','Salesforce','SapBusinessWarehouse','SapHana','SharePoint','Sybase','Teradata','Power','Any','Source','Navigation','Data','Schema','Item']);

function findDeps(sc,self){
    // Ignore plain-word token extraction inside #"..."
    // to avoid false splits like #"FactOnlineSales Agg" -> FactOnlineSales + Agg.
    let scan='';{
        let i=0;
        while(i<sc.length){
            if(sc[i]==='#'&&sc[i+1]==='"'){
                scan+='#"';i+=2;
                while(i<sc.length){
                    if(sc[i]==='"'&&sc[i+1]==='"'){scan+='  ';i+=2;}
                    else if(sc[i]==='"'){scan+='"';i++;break;}
                    else{scan+=sc[i]==='\n'?'\n':' ';i++;}
                }
                continue;
            }
            scan+=sc[i];i++;
        }
    }
    const d=new Set();let m;
    const ip=/\b([A-Za-z_][\w]*)\b/g;
    while((m=ip.exec(scan))!==null){const id=m[1];if(id!==self&&!MKW.has(id)&&!/^\d/.test(id)&&scan[m.index+id.length]!=='.')d.add(id);}
    const qp=/#"((?:[^"]*(?:""[^"]*)*)?)"/g;
    while((m=qp.exec(sc))!==null){const id=m[1].replace(/""/g,'"');if(id!==self)d.add(id);}
    return Array.from(d);
}

function findExternalFileRefs(code){
    const refs=new Set();let m;
    const re=/(?:File|Folder)\.Contents\s*\(\s*(#"(?:[^"]*(?:""[^"]*)*)"|"(?:[^"]|"")*")\s*\)/gi;
    while((m=re.exec(code))!==null){
        const lit=m[1];
        let p='';
        if(lit.startsWith('#"'))p=lit.slice(2,-1).replace(/""/g,'"');
        else p=lit.slice(1,-1).replace(/""/g,'"');
        const s=p.replace(/[?#].*$/,'').replace(/[\\\/]+$/,'').split(/[\\\/]/);
        const base=s[s.length-1]||p;
        if(base)refs.add(base);
    }
    return Array.from(refs);
}

// ═══ Syntax Highlighting ═══
function hlM(c){
    const e=escapeHtml(c);
    return e.replace(/(\/\*[\s\S]*?\*\/)/g,'<span class="cm">$1</span>')
        .replace(/(\/\/[^\n]*)/g,'<span class="cm">$1</span>')
        .replace(/(&quot;(?:[^&]|&(?!quot;))*?&quot;)/g,'<span class="str">$1</span>')
        .replace(/\b(\d+(?:\.\d+)?)\b/g,'<span class="num">$1</span>')
        .replace(/\b(let|in|if|then|else|true|false|null|and|or|not|each|error|try|otherwise|type|is|as|meta|shared|section)\b/g,'<span class="kw">$1</span>')
        .replace(/\b([A-Z][a-zA-Z]+\.[A-Z][a-zA-Z]+)\b/g,'<span class="fn">$1</span>')
        .replace(/(#&quot;[^&]*?&quot;)/g,'<span class="fn">$1</span>');
}

// ═══ UI ═══
function showToast(msg,type){
    const t=document.getElementById('toast');t.textContent=msg;
    t.className='toast show'+(type&&type!=='info'?' '+type:'');
    clearTimeout(t._t);t._t=setTimeout(()=>t.classList.remove('show'),4000);
}
function showBottomNotice(msg){
    const n=document.getElementById('bottomNotice');
    n.textContent=msg;
    n.classList.add('visible');
    clearTimeout(n._t);
    n._t=setTimeout(()=>n.classList.remove('visible'),4500);
}
function hideBottomNotice(){document.getElementById('bottomNotice').classList.remove('visible');}

function getActiveFileSet(){return new Set(appState.selectedFiles);}
function getActiveFiles(){const s=getActiveFileSet();return appState.files.filter(f=>s.has(f));}
function getActiveQueries(){const s=getActiveFileSet();return appState.queries.filter(q=>s.has(q.fileName));}
function applyFileSelection(){updateStats();renderFileList();renderGraph();renderCodePanel();renderDataPanel();}

function updateStats(){
    const q=getActiveQueries();
    const qByFile=new Map();for(const x of q){if(!qByFile.has(x.fileName))qByFile.set(x.fileName,new Set());qByFile.get(x.fileName).add(x.name);}
    const qNames=new Set(q.map(x=>x.name));
    const deps=q.reduce((s,x)=>s+x.dependencies.filter(d=>(qByFile.get(x.fileName)?.has(d))||qNames.has(d)).length,0);
    const chars=q.reduce((s,x)=>s+x.code.length+x.name.length+30,0);
    const nf=getActiveFiles().length,nq=q.length,tok='~'+Math.round(chars/3.5).toLocaleString();
    document.getElementById('statFiles').textContent=nf;
    document.getElementById('statQueries').textContent=nq;
    document.getElementById('statDeps').textContent=deps;
    document.getElementById('statTokens').textContent=tok;
    document.getElementById('headerStats').textContent=nf+' files | '+nq+' queries | '+deps+' deps | '+tok+' tokens';
}

function renderFileList(){
    const list=document.getElementById('fileList');
    const bf={};for(const q of appState.queries)bf[q.fileName]=(bf[q.fileName]||0)+1;
    const active=getActiveFileSet();
    list.innerHTML='<button class="file-select-all" id="fileSelectAll">Select All Files</button><span class="file-list-hint">Ctrl/Cmd+click to select multiple files</span>'+
        appState.files.map((f,i)=>{
            const isActive=active.has(f);
            return '<div class="file-chip legend-item '+(isActive?'active':'inactive')+'" data-file="'+escapeHtml(f)+'"><span class="dot" style="background:'+FILE_COLORS[i%FILE_COLORS.length]+'"></span>'+escapeHtml(f)+'<span class="qc">'+(bf[f]||0)+'</span></div>';
        }).join('');
    const selAll=document.getElementById('fileSelectAll');
    if(selAll)selAll.addEventListener('click',()=>{
        appState.selectedFiles=[...appState.files];
        applyFileSelection();
    });
    list.querySelectorAll('.file-chip').forEach(chip=>chip.addEventListener('click',e=>{
        const file=chip.dataset.file;
        const multi=e.ctrlKey||e.metaKey;
        const next=new Set(appState.selectedFiles);
        if(multi){
            if(next.has(file))next.delete(file);else next.add(file);
        }else{
            next.clear();
            next.add(file);
        }
        appState.selectedFiles=appState.files.filter(f=>next.has(f));
        applyFileSelection();
    }));
}

function renderGraph(){
    const ct=document.getElementById('graph-container');
    if(appState.cyInstance){appState.cyInstance.destroy();appState.cyInstance=null;}
    if(!appState.queries.length){ct.innerHTML='<div class="empty-state">No queries to visualize</div>';return;}
    const activeQueries=getActiveQueries();
    if(!activeQueries.length){ct.innerHTML='<div class="empty-state">No files selected. Use the file legend above.</div>';return;}
    const cm={};appState.files.forEach((f,i)=>{cm[f]=FILE_COLORS[i%FILE_COLORS.length];});
    const byFileAndName=new Map();for(const q of activeQueries)byFileAndName.set(q.fileName+'::'+q.name,q);
    const byName=new Map();for(const q of activeQueries)if(!byName.has(q.name))byName.set(q.name,q);
    const nodes=activeQueries.map(q=>({data:{id:q.fileName+'::'+q.name,label:q.name,fileName:q.fileName,color:cm[q.fileName]||'#8c99a8'}}));
    const nids=new Set(nodes.map(n=>n.data.id)),edges=[];
    for(const q of activeQueries)for(const dep of q.dependencies){
        // Only link dependencies within the same file
        const dq=byFileAndName.get(q.fileName+'::'+dep);
        if(dq){
            const depId=dq.fileName+'::'+dq.name,queryId=q.fileName+'::'+q.name;
            if(nids.has(depId)&&nids.has(queryId)&&depId!==queryId)edges.push({data:{source:depId,target:queryId}});
        }
    }
    appState.cyInstance=cytoscape({container:ct,elements:[...nodes,...edges],
        style:[
            {selector:'node',style:{'background-color':'data(color)','label':'data(label)','color':'#d5deea','font-size':'11px','font-family':"'Cascadia Code','Fira Code','Consolas',monospace",'text-valign':'bottom','text-margin-y':8,'width':24,'height':24,'border-width':2,'border-color':'data(color)','border-opacity':0.34,'background-opacity':0.86,'text-outline-color':'#111418','text-outline-width':2,'text-max-width':'150px','text-wrap':'ellipsis'}},
            {selector:'edge',style:{'width':1.8,'line-color':'#3e4b5d','target-arrow-color':'#3e4b5d','target-arrow-shape':'triangle','curve-style':'bezier','arrow-scale':0.75,'opacity':0.66}},
            {selector:'node:selected',style:{'border-width':3,'border-color':'#60c0a0','border-opacity':1,'background-opacity':1}},
            {selector:'node.highlighted',style:{'border-width':3,'border-color':'#c88a36','border-opacity':1,'width':34,'height':34,'font-size':'13px','z-index':999}},
            {selector:'node.dimmed',style:{'opacity':0.15}},
            {selector:'edge.dimmed',style:{'opacity':0.05}}
        ],
        layout:{name:'cose',animate:false,nodeRepulsion:function(){return 9000;},idealEdgeLength:function(){return 105;},gravity:0.45,numIter:600,padding:24},
        minZoom:0.2,maxZoom:4
    });
}

function renderCodePanel(){
    const content=document.getElementById('codeContent'),filters=document.getElementById('codeFilters');
    const activeFiles=getActiveFiles(),activeQueries=getActiveQueries();
    if(!activeFiles.length){
        filters.innerHTML='';
        content.innerHTML='<div class="empty-state">No files selected. Use the file legend above.</div>';
        updSel();
        return;
    }
    filters.innerHTML=activeFiles.map(f=>{
        const idx=appState.files.indexOf(f);
        return '<label class="filter-checkbox"><input type="checkbox" checked data-file="'+escapeHtml(f)+'"><span style="color:'+FILE_COLORS[(idx<0?0:idx)%FILE_COLORS.length]+'">'+escapeHtml(f)+'</span></label>';
    }).join('');
    const bf={};for(const q of activeQueries){if(!bf[q.fileName])bf[q.fileName]=[];bf[q.fileName].push(q);}
    content.innerHTML=activeFiles.map(f=>{
        const qs=bf[f]||[],idx=appState.files.indexOf(f),c=FILE_COLORS[(idx<0?0:idx)%FILE_COLORS.length];
        return '<div class="file-section" data-file="'+escapeHtml(f)+'"><div class="file-header" onclick="toggleFS(this)"><svg class="file-chevron" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="6,9 12,15 18,9"/></svg><span class="file-name" style="color:'+c+'">'+escapeHtml(f)+'</span><span class="file-badge">'+qs.length+' quer'+(qs.length===1?'y':'ies')+'</span></div><div class="file-queries">'+qs.map(q=>
            '<div class="query-block"><div class="query-header"><input type="checkbox" class="query-checkbox" checked data-query="'+escapeHtml(q.name)+'" data-file="'+escapeHtml(q.fileName)+'"><span class="query-name">'+escapeHtml(q.name)+'</span>'+(q.dependencies.length?'<span class="query-deps">\u2192 '+q.dependencies.map(d=>escapeHtml(d)).join(', ')+'</span>':'')+((q.externalRefs&&q.externalRefs.length)?'<span class="query-files'+(q.dependencies.length?'':' only')+'">\u2197 '+q.externalRefs.map(d=>escapeHtml(d)).join(', ')+'</span>':'')+'</div><pre class="query-code">'+hlM(q.code)+'</pre></div>'
        ).join('')+'</div></div>';
    }).join('');
    filters.querySelectorAll('input[type="checkbox"]').forEach(cb=>{
        cb.addEventListener('change',e=>{
            const f=e.target.dataset.file,sec=content.querySelector('.file-section[data-file="'+CSS.escape(f)+'"]');
            if(sec){sec.style.display=e.target.checked?'block':'none';sec.querySelectorAll('.query-checkbox').forEach(c=>{c.checked=e.target.checked;});}
            updSel();
        });
    });
    content.querySelectorAll('.query-checkbox').forEach(cb=>{cb.addEventListener('change',updSel);});
    updSel();
}

function toggleFS(h){h.parentElement.classList.toggle('collapsed');}
function updSel(){document.getElementById('copyBtnText').textContent='Copy Selected ('+document.querySelectorAll('.query-checkbox:checked').length+')';}

function getSelCode(prompt){
    const sel=document.querySelectorAll('.query-checkbox:checked'),bf={};
    sel.forEach(cb=>{const f=cb.dataset.file,n=cb.dataset.query,q=appState.queries.find(x=>x.name===n&&x.fileName===f);if(q){if(!bf[f])bf[f]=[];bf[f].push(q);}});
    let o=prompt||'';
    for(const[f,qs]of Object.entries(bf)){o+='// ========== '+f+' ==========\n\n';for(const q of qs){o+='// --- Query: '+q.name+' ---\n';if(q.dependencies.length)o+='// Dependencies: '+q.dependencies.join(', ')+'\n';if(q.externalRefs&&q.externalRefs.length)o+='// External file refs: '+q.externalRefs.join(', ')+'\n';o+=q.code+'\n\n';}}
    return o.trim();
}

async function copyClip(t,msg){
    if(!t){showToast('No queries selected','warning');return;}
    try{await navigator.clipboard.writeText(t);}
    catch(e){const ta=document.createElement('textarea');ta.value=t;ta.style.cssText='position:fixed;opacity:0';document.body.appendChild(ta);ta.select();document.execCommand('copy');document.body.removeChild(ta);}
    showToast(msg||'Copied to clipboard!','success');
}

// ═══ File Processing ═══
async function processFiles(files){
    document.getElementById('dropZone').style.display='none';
    document.getElementById('loading').classList.add('active');
    document.getElementById('errorLog').classList.remove('visible');
    hideBottomNotice();
    const prog=document.getElementById('fileProgress');
    if(appState.cyInstance){appState.cyInstance.destroy();}
    appState={files:[],queries:[],errors:[],selectedFiles:[],cyInstance:null,worksheets:[],activeSheet:null,dataProfile:null};

    for(let i=0;i<files.length;i++){
        const f=files[i];prog.textContent=(i+1)+' / '+files.length+': '+f.name;
        if(f.size>MAX_FILE_SIZE){appState.errors.push({file:f.name,msg:'Too large (>'+Math.round(MAX_FILE_SIZE/1024/1024)+'MB)'});continue;}
        try{
            const r=isPbixFile(f.name)?await parsePbixFile(f):await parseXlsxFile(f);
            if(r.errors.length)appState.errors.push(...r.errors.map(e=>({file:f.name,msg:e})));
            if(r.queries.length>0){appState.files.push(f.name);appState.queries.push(...r.queries);}
            else appState.errors.push({file:f.name,msg:'No Power Query code found'});
            // Extract worksheet/table data
            if(!isPbixFile(f.name)){try{const ws=await extractWorksheetData(f);appState.worksheets.push(...ws);}catch(e){}}
            else{try{const ws=await extractPbixTableData(f);appState.worksheets.push(...ws);}catch(e){}}
        }catch(e){appState.errors.push({file:f.name,msg:e.message||'Failed to parse'});}
    }

    document.getElementById('loading').classList.remove('active');
    const noQueryErrors=appState.errors.filter(e=>e.msg==='No Power Query code found');
    const otherErrors=appState.errors.filter(e=>e.msg!=='No Power Query code found');
    if(otherErrors.length){
        document.getElementById('errorList').innerHTML=otherErrors.map(e=>'<li>'+escapeHtml(e.file)+': '+escapeHtml(e.msg)+'</li>').join('');
        document.getElementById('errorLog').classList.add('visible');
    }
    if(noQueryErrors.length){
        if(noQueryErrors.length===1)showBottomNotice(noQueryErrors[0].file+': no Power Query code found');
        else showBottomNotice(noQueryErrors.length+' files had no Power Query code');
    }
    if(appState.queries.length>0){
        document.getElementById('mainContent').classList.add('active');
        document.getElementById('resetBtn').classList.add('visible');
        document.querySelector('.container').classList.add('compact');
        appState.selectedFiles=[...appState.files];
        // Show data tab if worksheets found
        if(appState.worksheets.length>0){
            document.getElementById('dataTabBtn').style.display='';
            document.getElementById('includeProfileWrap').style.display='';
            document.getElementById('headerProfileWrap').style.display='';
        }
        applyFileSelection();
    }else{
        document.getElementById('dropZone').style.display='block';
        if(!noQueryErrors.length)showToast('No Power Query code found. For Excel, check Data > Queries.','warning');
    }
}

function resetApp(){
    if(appState.cyInstance){appState.cyInstance.destroy();}
    appState={files:[],queries:[],errors:[],selectedFiles:[],cyInstance:null,worksheets:[],activeSheet:null,dataProfile:null};
    document.getElementById('mainContent').classList.remove('active');
    document.getElementById('resetBtn').classList.remove('visible');
    document.getElementById('errorLog').classList.remove('visible');
    document.querySelector('.container').classList.remove('compact');
    document.getElementById('dataTabBtn').style.display='none';
    document.getElementById('includeProfileWrap').style.display='none';
    document.getElementById('includeProfileCb').checked=false;
    hideBottomNotice();
    document.getElementById('dropZone').style.display='block';
    document.getElementById('fileInput').value='';
}

// Directory readers (legacy webkit entries + modern file system handles)
async function readDir(de,files){
    const rd=de.createReader(),all=[];
    await new Promise(res=>{(function rb(){rd.readEntries(ents=>{if(!ents.length)return res();all.push(...ents);rb();});})();});
    for(const e of all){
        if(e.isDirectory)await readDir(e,files);
        else if(isSupportedFile(e.name))files.push(await new Promise(r=>e.file(r)));
    }
}

async function readHandle(h,files){
    if(!h)return;
    if(h.kind==='file'){
        const f=await h.getFile();
        if(f&&isSupportedFile(f.name))files.push(f);
        return;
    }
    if(h.kind==='directory'){
        for await(const ch of h.values())await readHandle(ch,files);
    }
}

function uniqFiles(files){
    const m=new Map();
    for(const f of files){
        const k=[f.name,f.size,f.lastModified].join('::');
        if(!m.has(k))m.set(k,f);
    }
    return Array.from(m.values());
}

// ═══ Event Handlers ═══
const dz=document.getElementById('dropZone'),fi=document.getElementById('fileInput');
dz.addEventListener('dragover',e=>{e.preventDefault();e.stopPropagation();dz.classList.add('drag-over');});
dz.addEventListener('dragleave',e=>{e.preventDefault();e.stopPropagation();dz.classList.remove('drag-over');});
dz.addEventListener('drop',async e=>{
    e.preventDefault();e.stopPropagation();dz.classList.remove('drag-over');
    const dt=e.dataTransfer,items=dt?.items,files=[],pending=[];
    if(items)for(const it of items){
        if(it.kind!=='file')continue;
        let usedHandleApi=false;
        if(typeof it.getAsFileSystemHandle==='function'){
            try{
                // Collect handle promises in the same drop-event tick.
                // Some platforms only allow grabbing all handles synchronously.
                const hp=it.getAsFileSystemHandle();
                if(hp!=null){
                    usedHandleApi=true;
                    pending.push(Promise.resolve(hp).then(async h=>{if(h)await readHandle(h,files);}).catch(()=>{}));
                }
            }catch(err){}
        }
        if(usedHandleApi)continue;
        const en=it.webkitGetAsEntry?.();
        if(en){
            if(en.isDirectory)pending.push(readDir(en,files).catch(()=>{}));
            else if(isSupportedFile(en.name)){const f=it.getAsFile();if(f)files.push(f);}
            continue;
        }
        const f=it.getAsFile();
        if(f&&isSupportedFile(f.name))files.push(f);
    }
    if(dt?.files){
        for(const f of Array.from(dt.files))if(isSupportedFile(f.name))files.push(f);
    }
    if(pending.length)await Promise.all(pending);
    const finalFiles=uniqFiles(files);
    if(finalFiles.length)await processFiles(finalFiles);else showToast('No supported files found (.xlsx, .pbix, .pbit)','warning');
});
dz.addEventListener('click',e=>{if(e.target.closest('.browse-btn')||e.target===fi)return;fi.click();});
fi.addEventListener('change',async e=>{const f=Array.from(e.target.files).filter(f=>isSupportedFile(f.name));if(f.length)await processFiles(f);});

document.getElementById('resetBtn').addEventListener('click',resetApp);

document.querySelectorAll('.tab').forEach(tab=>{tab.addEventListener('click',()=>{
    document.querySelectorAll('.tab').forEach(t=>t.classList.remove('active'));tab.classList.add('active');
    const t=tab.dataset.tab;
    document.getElementById('graphPanel').classList.toggle('active',t==='graph');
    document.getElementById('codePanel').classList.toggle('active',t==='code');
    document.getElementById('limitationsPanel').classList.toggle('active',t==='limitations');
    document.getElementById('dataPanel').classList.toggle('active',t==='data');
});});

document.getElementById('copyBtn').addEventListener('click',async()=>{
    const profile=document.getElementById('includeProfileCb').checked;
    let extra='';
    if(profile&&appState.worksheets.length>0){
        const btn=document.getElementById('copyBtn');
        const origText=document.getElementById('copyBtnText').textContent;
        document.getElementById('copyBtnText').textContent='Computing profile...';
        btn.disabled=true;
        try{extra=buildDataProfile(appState.worksheets);}finally{document.getElementById('copyBtnText').textContent=origText;btn.disabled=false;}
    }
    copyClip(getSelCode()+(extra?'\n\n'+extra:''));
});

document.getElementById('selectAllBtn').addEventListener('click',()=>{
    const cbs=document.querySelectorAll('.query-checkbox'),all=Array.from(cbs).every(c=>c.checked);
    cbs.forEach(c=>c.checked=!all);
    document.querySelector('#selectAllBtn span').textContent=all?'Select All':'Deselect All';
    updSel();
});

document.querySelectorAll('.prompt-template').forEach(b=>{b.addEventListener('click',async()=>{
    const profile=document.getElementById('includeProfileCb').checked;
    let extra='';
    if(profile&&appState.worksheets.length>0){try{extra=buildDataProfile(appState.worksheets);}catch(e){}}
    copyClip(getSelCode(PROMPT_TEMPLATES[b.dataset.prompt])+(extra?'\n\n'+extra:''),'Prompt and selected M code copied');
});});

// Header Copy All button (compact mode)
document.getElementById('copyAllBtn').addEventListener('click',()=>{
    const key=document.getElementById('promptDropdown').value;
    const prefix=key?PROMPT_TEMPLATES[key]:'';
    const profile=document.getElementById('headerProfileCb').checked;
    let extra='';
    if(profile&&appState.worksheets.length>0){try{extra=buildDataProfile(appState.worksheets);}catch(e){}}
    const code=getSelCode(prefix);
    copyClip(code+(extra?'\n\n'+extra:''),prefix?'Prompt and selected M code copied':'Selected M code copied');
});

// Sync profile checkboxes between header and code panel
document.getElementById('headerProfileCb').addEventListener('change',e=>{document.getElementById('includeProfileCb').checked=e.target.checked;});
document.getElementById('includeProfileCb').addEventListener('change',e=>{document.getElementById('headerProfileCb').checked=e.target.checked;});

// Graph controls
document.getElementById('graphZoomIn').addEventListener('click',()=>{if(appState.cyInstance)appState.cyInstance.zoom(appState.cyInstance.zoom()*1.3);});
document.getElementById('graphZoomOut').addEventListener('click',()=>{if(appState.cyInstance)appState.cyInstance.zoom(appState.cyInstance.zoom()*0.7);});
document.getElementById('graphFit').addEventListener('click',()=>{if(appState.cyInstance)appState.cyInstance.fit(null,40);});
document.getElementById('graphRelayout').addEventListener('click',()=>{if(appState.cyInstance)appState.cyInstance.layout({name:'cose',animate:true,animationDuration:500,nodeRepulsion:function(){return 9000;},idealEdgeLength:function(){return 105;},gravity:0.45,numIter:600,padding:24}).run();});

let _gs;
document.getElementById('graphSearchInput').addEventListener('input',e=>{
    clearTimeout(_gs);_gs=setTimeout(()=>{
        const q=e.target.value.toLowerCase().trim(),cy=appState.cyInstance;if(!cy)return;
        cy.elements().removeClass('highlighted dimmed');if(!q)return;
        const m=cy.nodes().filter(n=>n.data('label').toLowerCase().includes(q));
        if(m.length){cy.elements().addClass('dimmed');m.removeClass('dimmed').addClass('highlighted');m.connectedEdges().removeClass('dimmed');m.neighborhood().nodes().removeClass('dimmed');cy.animate({center:{eles:m.first()},duration:300});}
    },200);
});

// Keyboard shortcuts
document.addEventListener('keydown',e=>{
    if(e.target.tagName==='INPUT'||e.target.tagName==='TEXTAREA')return;
    const cp=document.getElementById('codePanel').classList.contains('active');
    if((e.ctrlKey||e.metaKey)&&e.key==='a'&&cp){e.preventDefault();document.querySelectorAll('.query-checkbox').forEach(c=>c.checked=true);updSel();}
    if((e.ctrlKey||e.metaKey)&&e.key==='c'&&cp){e.preventDefault();copyClip(getSelCode());}
    if(e.key==='Escape'&&document.getElementById('mainContent').classList.contains('active'))resetApp();
});
// ═══ Worksheet Data Extraction (xlsx) ═══
async function extractWorksheetData(file){
    const zip=await JSZip.loadAsync(file);
    const sheets=[];
    // Read shared strings
    const ssEntry=zip.files['xl/sharedStrings.xml'];
    const sharedStrings=[];
    if(ssEntry){
        const ssXml=await ssEntry.async('string');
        const ssDoc=new DOMParser().parseFromString(ssXml,'text/xml');
        for(const si of ssDoc.getElementsByTagName('si')){
            const tEls=si.getElementsByTagName('t');
            let text='';for(const t of tEls)text+=t.textContent||'';
            sharedStrings.push(text);
        }
    }
    // Read workbook.xml for sheet names
    const wbEntry=zip.files['xl/workbook.xml'];
    const sheetNames=[];
    if(wbEntry){
        const wbXml=await wbEntry.async('string');
        const wbDoc=new DOMParser().parseFromString(wbXml,'text/xml');
        for(const s of wbDoc.getElementsByTagName('sheet'))sheetNames.push(s.getAttribute('name')||'Sheet');
    }
    // Parse each worksheet
    const MAX_ROWS=10000;
    for(const[path,entry]of Object.entries(zip.files)){
        const m=path.match(/^xl\/worksheets\/sheet(\d+)\.xml$/);
        if(!m||entry.dir)continue;
        const sheetIdx=parseInt(m[1])-1;
        const sheetName=sheetNames[sheetIdx]||('Sheet'+m[1]);
        const xml=await entry.async('string');
        const doc=new DOMParser().parseFromString(xml,'text/xml');
        const rowEls=doc.getElementsByTagName('row');
        const rows=[];let truncated=false,maxCol=0;
        for(const rowEl of rowEls){
            if(rows.length>=MAX_ROWS){truncated=true;break;}
            const cells=rowEl.getElementsByTagName('c');
            const row=[];
            for(const cell of cells){
                const ref=cell.getAttribute('r')||'';
                const colIdx=colRefToIndex(ref);
                const type=cell.getAttribute('t');
                const vEl=cell.getElementsByTagName('v')[0];
                const val=vEl?vEl.textContent:'';
                let resolved=val;
                if(type==='s')resolved=sharedStrings[parseInt(val)]||val;
                else if(type==='b')resolved=val==='1'?'TRUE':'FALSE';
                else if(type==='inlineStr'){const is=cell.getElementsByTagName('is')[0];if(is){const t=is.getElementsByTagName('t')[0];resolved=t?t.textContent:'';}}
                while(row.length<=colIdx)row.push('');
                row[colIdx]=resolved;
                if(colIdx+1>maxCol)maxCol=colIdx+1;
            }
            rows.push(row);
        }
        for(const row of rows)while(row.length<maxCol)row.push('');
        if(rows.length===0)continue;
        const headers=rows[0];
        const dataRows=rows.length>1?rows.slice(1):[];
        sheets.push({fileName:file.name,sheetName,headers,rows:dataRows,totalRows:dataRows.length,truncated});
    }
    return sheets;
}

function colRefToIndex(ref){
    const letters=ref.replace(/[0-9]/g,'');
    let idx=0;
    for(let i=0;i<letters.length;i++)idx=idx*26+(letters.charCodeAt(i)-64);
    return idx-1;
}

// ═══ Data Panel Rendering ═══
function renderDataPanel(){
    const sheetList=document.getElementById('dataSheetList');
    const preview=document.getElementById('dataPreview');
    if(!appState.worksheets.length){
        sheetList.innerHTML='';
        preview.innerHTML='<div class="empty-state">Drop an .xlsx file to preview worksheet data</div>';
        return;
    }
    const activeFiles=getActiveFileSet();
    const filtered=appState.worksheets.filter(ws=>activeFiles.has(ws.fileName));
    if(!filtered.length){
        sheetList.innerHTML='';
        preview.innerHTML='<div class="empty-state">No table data for the selected file(s)</div>';
        return;
    }
    let chips='';
    for(const ws of filtered){
        const isActive=appState.activeSheet&&appState.activeSheet.fileName===ws.fileName&&appState.activeSheet.sheetName===ws.sheetName;
        chips+='<button class="data-sheet-chip'+(isActive?' active':'')+'" data-file="'+escapeHtml(ws.fileName)+'" data-sheet="'+escapeHtml(ws.sheetName)+'">'+escapeHtml(ws.sheetName);
        if(filtered.length>1)chips+=' <span style="opacity:0.5;font-size:10px">'+escapeHtml(ws.fileName)+'</span>';
        chips+='</button>';
    }
    sheetList.innerHTML=chips;
    sheetList.querySelectorAll('.data-sheet-chip').forEach(chip=>{
        chip.addEventListener('click',()=>{
            const ws=appState.worksheets.find(w=>w.fileName===chip.dataset.file&&w.sheetName===chip.dataset.sheet);
            if(ws)renderWorksheetPreview(ws);
            sheetList.querySelectorAll('.data-sheet-chip').forEach(c=>c.classList.remove('active'));
            chip.classList.add('active');
        });
    });
    const activeStillVisible=appState.activeSheet&&filtered.some(ws=>ws.fileName===appState.activeSheet.fileName&&ws.sheetName===appState.activeSheet.sheetName);
    if(!activeStillVisible&&filtered.length>0){
        const first=filtered[0];
        appState.activeSheet={fileName:first.fileName,sheetName:first.sheetName};
        renderWorksheetPreview(first);
        const firstChip=sheetList.querySelector('.data-sheet-chip');
        if(firstChip)firstChip.classList.add('active');
    }
}

function renderWorksheetPreview(ws){
    const preview=document.getElementById('dataPreview');
    const DISPLAY_LIMIT=500;
    const displayRows=ws.rows.slice(0,DISPLAY_LIMIT);
    let html='<table><thead><tr>';
    for(const h of ws.headers)html+='<th>'+escapeHtml(h||'')+'</th>';
    html+='</tr></thead><tbody>';
    for(const row of displayRows){
        html+='<tr>';
        for(let i=0;i<ws.headers.length;i++)html+='<td>'+escapeHtml(row[i]!=null?String(row[i]):'')+'</td>';
        html+='</tr>';
    }
    html+='</tbody></table>';
    html+='<div class="data-row-info">Showing '+displayRows.length+' of '+ws.totalRows+' rows'+(ws.truncated?' (capped at 10,000 rows)':'')+'</div>';
    preview.innerHTML=html;
    appState.activeSheet={fileName:ws.fileName,sheetName:ws.sheetName};
    document.getElementById('exportCsvBtn').disabled=false;
    document.getElementById('exportParquetBtn').disabled=false;
}

// ═══ Streaming CSV Export ═══
function escapeCSVField(v){
    if(v==null)return'';
    const s=String(v);
    if(s.includes(',')||s.includes('\n')||s.includes('"')||s.includes('\r'))return'"'+s.replace(/"/g,'""')+'"';
    return s;
}

async function exportCSV(ws){
    if(!ws){showToast('No table selected','warning');return;}
    const CHUNK=5000;
    const parts=[];
    // Header row
    parts.push(ws.headers.map(escapeCSVField).join(',')+'\n');
    // Stream rows in chunks to avoid blocking
    for(let start=0;start<ws.rows.length;start+=CHUNK){
        const end=Math.min(start+CHUNK,ws.rows.length);
        const lines=[];
        for(let r=start;r<end;r++){
            const fields=[];
            for(let c=0;c<ws.headers.length;c++)fields.push(escapeCSVField(ws.rows[r][c]));
            lines.push(fields.join(','));
        }
        parts.push(lines.join('\n')+'\n');
        if(end<ws.rows.length)await new Promise(r=>setTimeout(r,0));
    }
    const blob=new Blob(['\uFEFF',...parts],{type:'text/csv;charset=utf-8'});
    downloadBlob(blob,ws.sheetName.replace(/[^a-zA-Z0-9_-]/g,'_')+'.csv');
    showToast('Exported '+ws.sheetName+'.csv','success');
}

// ═══ Parquet Export (columnar streaming) ═══
async function exportParquet(ws){
    if(!ws){showToast('No table selected','warning');return;}
    // Build columnar data streaming one column at a time
    const CHUNK=5000;
    const colArrays=[];
    for(let c=0;c<ws.headers.length;c++){
        const parts=[];
        for(let start=0;start<ws.rows.length;start+=CHUNK){
            const end=Math.min(start+CHUNK,ws.rows.length);
            for(let r=start;r<end;r++)parts.push(ws.rows[r][c]!=null?String(ws.rows[r][c]):'');
            if(end<ws.rows.length)await new Promise(r=>setTimeout(r,0));
        }
        colArrays.push(parts);
    }
    // Build a minimal Parquet file
    const buf=buildParquetBuffer(ws.headers,colArrays);
    const blob=new Blob([buf],{type:'application/octet-stream'});
    downloadBlob(blob,ws.sheetName.replace(/[^a-zA-Z0-9_-]/g,'_')+'.parquet');
    showToast('Exported '+ws.sheetName+'.parquet','success');
}

// Minimal Parquet writer (no external dependency, string-only columns)
function buildParquetBuffer(headers,colArrays){
    const numRows=colArrays[0]?colArrays[0].length:0;
    const numCols=headers.length;
    const enc=new TextEncoder();
    // Thrift compact protocol helpers
    function writeVarint(arr,v){v=v>>>0;while(v>0x7f){arr.push((v&0x7f)|0x80);v>>>=7;}arr.push(v&0x7f);}
    function writeZigzag(arr,v){writeVarint(arr,(v<<1)^(v>>31));}
    function writeBinary(arr,buf){writeVarint(arr,buf.length);for(const b of buf)arr.push(b);}
    function writeFieldHeader(arr,delta,type){if(delta>0&&delta<=15)arr.push((delta<<4)|type);else{arr.push(type);writeZigzag(arr,delta);/* not used for small schemas */}}
    // Build column chunks (plain encoding, no compression)
    const columnChunks=[];
    const columnMetas=[];
    for(let c=0;c<numCols;c++){
        // Data page: plain encoding, byte_array type
        const pageData=[];
        // Definition levels (all present = all 1s, RLE encoded)
        // For required columns, no definition levels needed in PLAIN
        // Values: length-prefixed byte arrays
        for(let r=0;r<numRows;r++){
            const val=enc.encode(colArrays[c][r]||'');
            const lenBuf=new Uint8Array(4);
            new DataView(lenBuf.buffer).setInt32(0,val.length,true);
            for(const b of lenBuf)pageData.push(b);
            for(const b of val)pageData.push(b);
        }
        // Page header (Thrift)
        const pageHeader=[];
        writeFieldHeader(pageHeader,1,5);writeZigzag(pageHeader,0);// type=DATA_PAGE(0)
        writeFieldHeader(pageHeader,1,5);writeZigzag(pageHeader,pageData.length);// uncompressed_page_size
        writeFieldHeader(pageHeader,1,5);writeZigzag(pageHeader,pageData.length);// compressed_page_size
        // DataPageHeader
        writeFieldHeader(pageHeader,2,12);// field 5 -> data_page_header (struct)
        writeFieldHeader(pageHeader,1,5);writeZigzag(pageHeader,numRows);// num_values
        writeFieldHeader(pageHeader,1,5);writeZigzag(pageHeader,0);// encoding=PLAIN
        writeFieldHeader(pageHeader,1,5);writeZigzag(pageHeader,0);// def_level_encoding=RLE
        writeFieldHeader(pageHeader,1,5);writeZigzag(pageHeader,0);// rep_level_encoding=RLE
        pageHeader.push(0);// stop DataPageHeader struct
        pageHeader.push(0);// stop PageHeader struct
        const chunk=new Uint8Array(pageHeader.length+pageData.length);
        chunk.set(pageHeader);
        chunk.set(new Uint8Array(pageData),pageHeader.length);
        columnChunks.push(chunk);
        columnMetas.push({name:headers[c],offset:0,size:chunk.length,numValues:numRows});
    }
    // Assemble file
    const magic=enc.encode('PAR1');
    // Calculate offsets
    let offset=4;// after magic
    for(let c=0;c<numCols;c++){columnMetas[c].offset=offset;offset+=columnChunks[c].length;}
    // Build FileMetaData (Thrift)
    const meta=[];
    writeFieldHeader(meta,1,5);writeZigzag(meta,1);// version
    // Schema
    writeFieldHeader(meta,1,12);// schema (list of SchemaElement)
    // Write list header: field type 12 (struct), count
    meta.push(0x0c);// type=struct in list
    writeVarint(meta,numCols+1);// root + columns
    // Root element
    writeFieldHeader(meta,1,8);writeBinary(meta,enc.encode('schema'));// name
    writeFieldHeader(meta,2,5);writeZigzag(meta,numCols);// num_children
    meta.push(0);// stop root
    for(let c=0;c<numCols;c++){
        writeFieldHeader(meta,1,8);writeBinary(meta,enc.encode(headers[c]));// name
        writeFieldHeader(meta,2,5);writeZigzag(meta,6);// type=BYTE_ARRAY
        writeFieldHeader(meta,2,5);writeZigzag(meta,1);// repetition=REQUIRED
        meta.push(0);// stop
    }
    writeFieldHeader(meta,1,5);writeZigzag(meta,numRows);// num_rows
    // Row groups
    writeFieldHeader(meta,1,12);// row_groups list
    meta.push(0x0c);writeVarint(meta,1);// one row group
    // RowGroup
    writeFieldHeader(meta,1,12);// columns list
    meta.push(0x0c);writeVarint(meta,numCols);
    for(let c=0;c<numCols;c++){
        // ColumnChunk
        writeFieldHeader(meta,1,5);writeZigzag(meta,columnMetas[c].offset);// file_offset
        writeFieldHeader(meta,1,12);// meta_data struct
        writeFieldHeader(meta,1,5);writeZigzag(meta,6);// type=BYTE_ARRAY
        writeFieldHeader(meta,1,12);// encodings list
        meta.push(0x05);writeVarint(meta,1);writeZigzag(meta,0);// PLAIN
        writeFieldHeader(meta,1,12);// path_in_schema list
        meta.push(0x08);writeVarint(meta,1);writeBinary(meta,enc.encode(headers[c]));
        writeFieldHeader(meta,1,5);writeZigzag(meta,0);// codec=UNCOMPRESSED
        writeFieldHeader(meta,1,5);writeZigzag(meta,numRows);// num_values
        writeFieldHeader(meta,1,5);writeZigzag(meta,columnChunks[c].length);// total_uncompressed_size
        writeFieldHeader(meta,1,5);writeZigzag(meta,columnChunks[c].length);// total_compressed_size
        writeFieldHeader(meta,1,5);writeZigzag(meta,columnMetas[c].offset);// data_page_offset
        meta.push(0);// stop meta_data
        meta.push(0);// stop ColumnChunk
    }
    writeFieldHeader(meta,1,5);writeZigzag(meta,offset-4);// total_byte_size
    writeFieldHeader(meta,1,5);writeZigzag(meta,numRows);// num_rows
    meta.push(0);// stop RowGroup
    meta.push(0);// stop FileMetaData
    const metaBytes=new Uint8Array(meta);
    const metaLen=new Uint8Array(4);
    new DataView(metaLen.buffer).setInt32(0,metaBytes.length,true);
    // Final file: magic + chunks + metadata + metadata_len + magic
    const totalSize=4+offset-4+metaBytes.length+4+4;
    const result=new Uint8Array(totalSize);
    let pos=0;
    result.set(magic,pos);pos+=4;
    for(const chunk of columnChunks){result.set(chunk,pos);pos+=chunk.length;}
    result.set(metaBytes,pos);pos+=metaBytes.length;
    result.set(metaLen,pos);pos+=4;
    result.set(magic,pos);
    return result;
}

function downloadBlob(blob,filename){
    const url=URL.createObjectURL(blob);
    const a=document.createElement('a');
    a.href=url;a.download=filename;
    document.body.appendChild(a);a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
}

// ═══ Data Profile for LLM Copy ═══
function buildDataProfile(worksheets){
    const lines=['// ========== Data Profile ==========',''];
    for(const ws of worksheets){
        lines.push('// --- '+ws.fileName+' / '+ws.sheetName+' ('+ws.totalRows+' rows) ---');
        for(let c=0;c<ws.headers.length;c++){
            const colName=ws.headers[c]||('Column'+(c+1));
            const values=ws.rows.map(r=>r[c]);
            const stats=computeColumnStats(colName,values);
            let line='// '+stats.name+': '+stats.distinct+' distinct, '+stats.nulls+' nulls/empty';
            if(stats.isNumeric&&stats.numCount>0){
                line+=', min='+stats.min+', max='+stats.max+', avg='+Math.round(stats.avg*100)/100;
            }
            if(stats.top&&stats.top.length>0){
                line+=', top=['+stats.top.map(t=>'"'+t.value+'"('+t.count+')').join(', ')+']';
            }
            lines.push(line);
        }
        lines.push('');
    }
    return lines.join('\n');
}

function computeColumnStats(name,data){
    const stat={name,distinct:0,nulls:0,isNumeric:false,numCount:0,min:Infinity,max:-Infinity,avg:0,top:null};
    const seen=new Set();const freq=new Map();let numSum=0;
    for(const v of data){
        if(v==null||v===''){stat.nulls++;continue;}
        seen.add(v);
        if(freq.size<5000)freq.set(v,(freq.get(v)||0)+1);
        const n=Number(v);
        if(!isNaN(n)&&v!==''){stat.isNumeric=true;stat.numCount++;numSum+=n;if(n<stat.min)stat.min=n;if(n>stat.max)stat.max=n;}
    }
    stat.distinct=seen.size;
    if(stat.numCount>0)stat.avg=numSum/stat.numCount;
    else{stat.min=0;stat.max=0;}
    if(!stat.isNumeric&&freq.size>0&&freq.size<=500){
        stat.top=[...freq.entries()].sort((a,b)=>b[1]-a[1]).slice(0,5).map(([v,c])=>({value:String(v).substring(0,40),count:c}));
    }
    return stat;
}

// ═══ PBIX DataModel Table Data Extraction ═══

function buildSchemaFromSQLite(db) {
    const tableRows = db.getTableRows('Table');
    const columnRows = db.getTableRows('Column');
    const columnStorageRows = db.getTableRows('ColumnStorage');
    const columnPartitionStorageRows = db.getTableRows('ColumnPartitionStorage');
    const dictionaryStorageRows = db.getTableRows('DictionaryStorage');
    const storageFileRows = db.getTableRows('StorageFile');
    const attrHierRows = db.getTableRows('AttributeHierarchy');
    const attrHierStorageRows = db.getTableRows('AttributeHierarchyStorage');

    const tables = new Map();
    for (const r of tableRows) tables.set(r.rowid, { name: r.values[2] });

    const storageFiles = new Map();
    for (const r of storageFileRows) storageFiles.set(r.rowid, r.values[4]);

    const colStorageMap = new Map();
    for (const r of columnStorageRows) {
        colStorageMap.set(r.rowid, {
            dictStorageId: r.values[4],
            cardinality: r.values[11]
        });
    }

    const dictStorageMap = new Map();
    for (const r of dictionaryStorageRows) {
        dictStorageMap.set(r.rowid, {
            baseId: r.values[5],
            magnitude: r.values[6],
            isNullable: r.values[8],
            storageFileId: r.values[12]
        });
    }

    const colPartStorageMap = new Map();
    for (const r of columnPartitionStorageRows) colPartStorageMap.set(r.values[1], r.values[6]);

    const attrHierMap = new Map();
    for (const r of attrHierRows) {
        const colId = r.values[1];
        const ahsId = r.values[3];
        if (colId != null) attrHierMap.set(colId, ahsId);
    }

    const attrHierStorageMap = new Map();
    for (const r of attrHierStorageRows) attrHierStorageMap.set(r.rowid, r.values[9]);

    const result = new Map();

    for (const r of columnRows) {
        const colId = r.rowid;
        const tableId = r.values[1];
        const colName = r.values[2];
        const dataType = r.values[4];
        const colStorageId = r.values[18];
        const colType = r.values[19];

        if (colType !== 1 && colType !== 2) continue;

        const tableInfo = tables.get(tableId);
        if (!tableInfo) continue;
        const tableName = tableInfo.name;

        if (/^(LocalDateTable_|DateTableTemplate_|H\$|R\$|U\$)/.test(tableName)) continue;

        const cs = colStorageMap.get(colStorageId);
        if (!cs) continue;

        const idfSfId = colPartStorageMap.get(colStorageId);
        const idfFile = idfSfId != null ? storageFiles.get(idfSfId) : null;
        if (!idfFile) continue;

        let dictFile = null, baseId = 0, magnitude = 1, isNullable = false;
        if (cs.dictStorageId != null) {
            const ds = dictStorageMap.get(cs.dictStorageId);
            if (ds) {
                dictFile = ds.storageFileId != null ? storageFiles.get(ds.storageFileId) : null;
                baseId = ds.baseId || 0;
                magnitude = ds.magnitude || 1;
                isNullable = !!ds.isNullable;
            }
        }

        let hidxFile = null;
        const ahsId = attrHierMap.get(colId);
        if (ahsId != null) {
            const sfId = attrHierStorageMap.get(ahsId);
            if (sfId != null) hidxFile = storageFiles.get(sfId);
        }

        if (!result.has(tableName)) result.set(tableName, { columns: [] });

        result.get(tableName).columns.push({
            name: colName, idf: idfFile, dictionary: dictFile, hidx: hidxFile,
            dataType, baseId, magnitude, isNullable, cardinality: cs.cardinality
        });
    }

    return result;
}

function _buildFileCache(schema, abf) {
    const cache = new Map();
    const _get = (name) => {
        if (cache.has(name)) return;
        try { cache.set(name, getDataSlice(abf, name)); } catch (e) { /* skip missing */ }
    };
    for (const [, tableInfo] of schema) {
        for (const col of tableInfo.columns) {
            _get(col.idf);
            _get(col.idf + 'meta');
            if (col.dictionary) _get(col.dictionary);
        }
    }
    return cache;
}

function readIdfMeta(buf) {
    const dv = new DataView(buf.buffer, buf.byteOffset, buf.byteLength);
    let pos = 0;

    pos += 6; // CP tag
    pos += 8; // version_one
    pos += 6; // CS tag
    pos += 8; // records
    pos += 8; // one
    const aba5a = dv.getUint32(pos, true); pos += 4;
    const iterator = dv.getUint32(pos, true); pos += 4;
    pos += 8; // bookmark_bits
    pos += 8; // storage_alloc_size
    pos += 8; // storage_used_size
    pos += 1; // segment_needs_resizing
    pos += 4; // compression_info

    pos += 6; // SS tag
    pos += 8; // distinct_states
    const minDataId = dv.getUint32(pos, true); pos += 4;
    pos += 4; // max_data_id
    pos += 4; // original_min_segment_data_id
    pos += 8; // rle_sort_order
    const rowCount = Number(dv.getBigUint64(pos, true)); pos += 8;
    pos += 1; // has_nulls
    pos += 8; // rle_runs
    pos += 8; // others_rle_runs
    pos += 6; // SS end tag

    pos += 1; // has_bit_packed_sub_seg
    pos += 6; // CS1 tag
    const countBitPacked = Number(dv.getBigUint64(pos, true));

    const bitWidth = (36 - aba5a) + iterator;

    return { minDataId, countBitPacked, bitWidth, rowCount };
}

function readIdf(buf) {
    const dv = new DataView(buf.buffer, buf.byteOffset, buf.byteLength);
    let pos = 0;

    const primarySize = Number(dv.getBigUint64(pos, true)); pos += 8;
    const primarySegment = [];
    for (let i = 0; i < primarySize; i++) {
        const dataValue = dv.getUint32(pos, true); pos += 4;
        const repeatValue = dv.getUint32(pos, true); pos += 4;
        primarySegment.push({ dataValue, repeatValue });
    }

    const subSegSize = Number(dv.getBigUint64(pos, true)); pos += 8;
    const subSegment = [];
    for (let i = 0; i < subSegSize; i++) {
        subSegment.push(dv.getBigUint64(pos, true)); pos += 8;
    }

    return { primarySegment, subSegment };
}

function decodeRleBitPackedHybrid(idfData, meta) {
    const { primarySegment, subSegment } = idfData;
    const { minDataId, countBitPacked, bitWidth } = meta;

    let bitpackedValues = [];
    if (countBitPacked > 0 && subSegment.length > 0) {
        if (subSegment.length === 1 && subSegment[0] === 0n) {
            bitpackedValues = new Array(countBitPacked).fill(minDataId);
        } else {
            const mask = BigInt((1 << bitWidth) - 1);
            const minId = BigInt(minDataId);
            const bw = BigInt(bitWidth);
            for (const u64 of subSegment) {
                let val = u64;
                const count = 64 / bitWidth;
                for (let j = 0; j < count; j++) {
                    bitpackedValues.push(Number(minId + (val & mask)));
                    val >>= bw;
                }
            }
        }
    }

    const vector = [];
    let bpOffset = 0;

    for (const entry of primarySegment) {
        if ((entry.dataValue + bpOffset) === 0xFFFFFFFF) {
            const count = entry.repeatValue;
            for (let i = 0; i < count && bpOffset + i < bitpackedValues.length; i++) {
                vector.push(bitpackedValues[bpOffset + i]);
            }
            bpOffset += count;
        } else {
            for (let i = 0; i < entry.repeatValue; i++) {
                vector.push(entry.dataValue);
            }
        }
    }

    return vector;
}

function decompressEncodeArray(compressed) {
    const full = new Array(256).fill(0);
    for (let i = 0; i < 128; i++) {
        full[2 * i] = compressed[i] & 0x0F;
        full[2 * i + 1] = (compressed[i] >> 4) & 0x0F;
    }
    return full;
}

function buildHuffmanTree(encodeArray) {
    const sorted = [];
    for (let i = 0; i < 256; i++) {
        if (encodeArray[i] !== 0) sorted.push([encodeArray[i], i]);
    }
    sorted.sort((a, b) => a[0] - b[0] || a[1] - b[1]);

    const root = { left: null, right: null, c: 0 };
    let code = 0, lastLen = 0;
    for (const [len, ch] of sorted) {
        if (lastLen !== len) {
            code <<= (len - lastLen);
            lastLen = len;
        }
        let node = root;
        for (let bit = len - 1; bit >= 0; bit--) {
            if (code & (1 << bit)) {
                if (!node.right) node.right = { left: null, right: null, c: 0 };
                node = node.right;
            } else {
                if (!node.left) node.left = { left: null, right: null, c: 0 };
                node = node.left;
            }
        }
        node.c = ch;
        code++;
    }
    return root;
}

function decodeHuffmanString(bitstream, tree, startBit, endBit) {
    let result = '';
    let node = tree;
    const totalBits = endBit - startBit;

    for (let i = 0; i < totalBits; i++) {
        let bitPos = startBit + i;
        let bytePos = bitPos >> 3;
        let bitOffset = bitPos & 7;
        bytePos = (bytePos & ~1) + (1 - (bytePos & 1));

        if (!node.left && !node.right) {
            result += String.fromCharCode(node.c);
            node = tree;
        }

        if (bitstream[bytePos] & (1 << (7 - bitOffset))) {
            node = node.right;
        } else {
            node = node.left;
        }
    }

    if (!node.left && !node.right) {
        result += String.fromCharCode(node.c);
    }

    return result;
}

function readDictionary(buf, minDataId) {
    const dv = new DataView(buf.buffer, buf.byteOffset, buf.byteLength);
    let pos = 0;

    const dictType = dv.getInt32(pos, true); pos += 4;
    pos += 24; // hash_information

    if (dictType === 2) return readStringDictionary(buf, pos, minDataId);
    else if (dictType === 0 || dictType === 1) return readNumericDictionary(buf, pos, minDataId, dictType);
    return new Map();
}

function readStringDictionary(buf, pos, minDataId) {
    const dv = new DataView(buf.buffer, buf.byteOffset, buf.byteLength);
    const dict = new Map();

    pos += 8; // store_string_count
    pos += 1; // f_store_compressed
    pos += 8; // store_longest_string
    const pageCount = Number(dv.getBigInt64(pos, true)); pos += 8;

    const pages = [];
    for (let p = 0; p < pageCount; p++) {
        const page = {};
        pos += 8; // page_mask
        pos += 1; // page_contains_nulls
        pos += 8; // page_start_index
        page.stringCount = Number(dv.getBigUint64(pos, true)); pos += 8;
        page.compressed = buf[pos]; pos += 1;
        pos += 4; // string_store_begin_mark

        if (page.compressed) {
            page.storeTotalBits = dv.getUint32(pos, true); pos += 4;
            page.charSetTypeId = dv.getUint32(pos, true); pos += 4;
            const charSetTypeId = page.charSetTypeId;
            const allocSize = Number(dv.getBigUint64(pos, true)); pos += 8;
            if (charSetTypeId !== 703122) pos += 1; // character_set_used (absent in type 703122)
            pos += 4; // ui_decode_bits
            page.encodeArray = new Uint8Array(buf.buffer, buf.byteOffset + pos, 128);
            pos += 128;
            pos += 8; // ui64_buffer_size
            page.compressedBuffer = new Uint8Array(buf.buffer, buf.byteOffset + pos, allocSize);
            pos += allocSize;
        } else {
            pos += 8; // remaining_store_available
            pos += 8; // buffer_used_characters
            const allocSize = Number(dv.getBigUint64(pos, true)); pos += 8;
            const textBytes = new Uint8Array(buf.buffer, buf.byteOffset + pos, allocSize);
            page.text = new TextDecoder('utf-16le').decode(textBytes);
            pos += allocSize;
        }

        // Verify end mark (0xABCDABCD); self-correct if format variant shifted pos
        if (pos + 4 <= buf.byteLength && dv.getUint32(pos, true) !== 0xABCDABCD) {
            for (let adj = -2; adj <= 2; adj++) {
                if (adj !== 0 && pos + adj >= 0 && pos + adj + 4 <= buf.byteLength && dv.getUint32(pos + adj, true) === 0xABCDABCD) { pos += adj; break; }
            }
        }
        pos += 4; // string_store_end_mark
        pages.push(page);
    }

    const handleCount = Number(dv.getBigUint64(pos, true)); pos += 8;
    pos += 4; // element_size

    const handles = [];
    for (let i = 0; i < handleCount; i++) {
        const offset = dv.getUint32(pos, true); pos += 4;
        const pageId = dv.getUint32(pos, true); pos += 4;
        handles.push({ offset, pageId });
    }

    const handlesByPage = new Map();
    for (const h of handles) {
        if (!handlesByPage.has(h.pageId)) handlesByPage.set(h.pageId, []);
        handlesByPage.get(h.pageId).push(h.offset);
    }

    let index = minDataId;
    for (let pageId = 0; pageId < pages.length; pageId++) {
        const page = pages[pageId];

        if (page.compressed) {
            const fullEncode = decompressEncodeArray(page.encodeArray);
            const tree = buildHuffmanTree(fullEncode);
            const offsets = handlesByPage.get(pageId) || [];

            const isUtf16 = page.charSetTypeId === 703122;
            for (let i = 0; i < offsets.length; i++) {
                const startBit = offsets[i];
                const endBit = (i + 1 < offsets.length) ? offsets[i + 1] : page.storeTotalBits;
                let decoded = decodeHuffmanString(page.compressedBuffer, tree, startBit, endBit);
                if (isUtf16 && decoded.length >= 2) {
                    const bytes = new Uint8Array(decoded.length);
                    for (let j = 0; j < decoded.length; j++) bytes[j] = decoded.charCodeAt(j);
                    decoded = new TextDecoder('utf-16le').decode(bytes);
                }
                dict.set(index, decoded);
                index++;
            }
        } else {
            const strings = page.text.split('\0');
            if (strings.length > 0 && strings[strings.length - 1] === '') strings.pop();
            for (const s of strings) {
                dict.set(index, s);
                index++;
            }
        }
    }

    return dict;
}

function readNumericDictionary(buf, pos, minDataId, dictType) {
    const dv = new DataView(buf.buffer, buf.byteOffset, buf.byteLength);
    const dict = new Map();

    const count = Number(dv.getBigUint64(pos, true)); pos += 8;
    const elemSize = dv.getUint32(pos, true); pos += 4;

    for (let i = 0; i < count; i++) {
        let val;
        if (elemSize === 4) {
            val = dv.getInt32(pos, true); pos += 4;
        } else if (elemSize === 8 && dictType === 0) {
            val = Number(dv.getBigInt64(pos, true)); pos += 8;
        } else {
            val = dv.getFloat64(pos, true); pos += 8;
        }
        dict.set(minDataId + i, val);
    }

    return dict;
}

function convertColumnValue(value, dataType) {
    if (value == null) return null;
    switch (dataType) {
        case 9: return new Date((value - 25569) * 86400000);
        case 10: return value / 10000;
        default: return value;
    }
}

function _extractColumn(col, fileCache) {
    let meta;
    try {
        const metaBuf = fileCache.get(col.idf + 'meta');
        if (!metaBuf) return null;
        meta = readIdfMeta(metaBuf);
    } catch (e) { return null; }

    const idfBuf = fileCache.get(col.idf);
    if (!idfBuf) return null;

    const indices = decodeRleBitPackedHybrid(readIdf(idfBuf), meta);

    if (col.dictionary) {
        try {
            const dictBuf = fileCache.get(col.dictionary);
            if (!dictBuf) return indices;
            const dict = readDictionary(dictBuf, meta.minDataId);
            return indices.map(idx => {
                const v = dict.get(idx);
                return v !== undefined ? convertColumnValue(v, col.dataType) : null;
            });
        } catch (e) { return indices; }
    } else if (col.hidx) {
        return indices.map(idx => convertColumnValue((idx + col.baseId) / col.magnitude, col.dataType));
    }
    return indices;
}

function extractTableData(tableName, schema, fileCache) {
    const tableSchema = schema.get(tableName);
    if (!tableSchema) throw new Error('Table not found: ' + tableName);

    const columns = [];
    const columnData = [];

    for (const col of tableSchema.columns) {
        const values = _extractColumn(col, fileCache);
        if (values === null) continue;
        columns.push(col.name);
        columnData.push(values);
    }

    const rowCount = columnData.reduce((max, c) => Math.max(max, c.length), 0);
    return { columns, columnData, rowCount };
}

function formatPbixValue(val) {
    if (val == null) return '';
    if (val instanceof Date) {
        if (isNaN(val.getTime())) return '';
        return val.toISOString().replace('T', ' ').replace(/\.000Z$/, '');
    }
    return String(val);
}

async function extractPbixTableData(file) {
    const zip = await JSZip.loadAsync(file);
    const dmEntry = zip.files['DataModel'] || zip.files['datamodel'];
    if (!dmEntry) return [];

    if (typeof parseABF !== 'function' || typeof getDataSlice !== 'function' || typeof readSQLiteTables !== 'function') return [];

    const dmData = await dmEntry.async('arraybuffer');
    const decompressed = await decompressXpress9(dmData);
    const abf = parseABF(decompressed);
    const sqliteBuf = getDataSlice(abf, 'metadata.sqlitedb');
    const db = readSQLiteTables(sqliteBuf);
    const schema = buildSchemaFromSQLite(db);
    const fileCache = _buildFileCache(schema, abf);

    const MAX_ROWS = 10000;
    const sheets = [];

    for (const tableName of Array.from(schema.keys()).sort()) {
        try {
            const td = extractTableData(tableName, schema, fileCache);
            if (td.columns.length === 0) continue;

            const headers = td.columns;
            const totalRows = td.rowCount;
            const rowLimit = Math.min(totalRows, MAX_ROWS);
            const rows = [];

            for (let r = 0; r < rowLimit; r++) {
                const row = [];
                for (let c = 0; c < td.columnData.length; c++) {
                    row.push(formatPbixValue(td.columnData[c][r]));
                }
                rows.push(row);
            }

            sheets.push({
                fileName: file.name,
                sheetName: tableName,
                headers,
                rows,
                totalRows,
                truncated: totalRows > MAX_ROWS
            });
        } catch (e) { /* skip tables that fail to decode */ }
    }

    return sheets;
}

// ═══ Export Button Handlers ═══
document.getElementById('exportCsvBtn').addEventListener('click',()=>{
    const ws=appState.worksheets.find(w=>appState.activeSheet&&w.fileName===appState.activeSheet.fileName&&w.sheetName===appState.activeSheet.sheetName);
    exportCSV(ws);
});
document.getElementById('exportParquetBtn').addEventListener('click',()=>{
    const ws=appState.worksheets.find(w=>appState.activeSheet&&w.fileName===appState.activeSheet.fileName&&w.sheetName===appState.activeSheet.sheetName);
    exportParquet(ws);
});
