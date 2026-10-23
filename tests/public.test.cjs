const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const XLSX = require('xlsx');

const scriptPath = path.join(__dirname, '..', 'case_review_public_3.3.7.user.js');
const source = fs.readFileSync(scriptPath, 'utf8');
const plain = value => JSON.parse(JSON.stringify(value));
const names = ['CLASSIFICATIONS','PUBLIC_SETTINGS','CLEAN_BASE_COLUMNS','CLEAN_FIRST_COLUMNS','CLEAN_FINAL_COLUMNS','CLEAN_INITIAL','CLEAN_FINAL','CLEAN_SCRIPT_KEYWORDS','CLEAN_AI_KEYWORDS','SHEET_NAME','compilePrinciple','executePrinciple','classificationRules','rulesSignature','classifyRecord','filterTopics','canonicalReviewTopics','cleanTopicSignature','buildSharedRulebook','validateCleanFinalRow','readCleanState','writeCleanState','rememberCleanInitial','makeCleanWorkbook','cleanExportRows','parseExcel','runImport','importCleanFinalRows','runSecondImport','runExport','buildDashboardData','buildReportHTMLV2','getFieldMap','fetchReviewContexts','listRoot','injectButtons'];

function environment({items=[],writeFailure=false}={}) {
    const store=new Map(),requests=[],downloads=[],notices=[],consoleErrors=[];
    const state={items:structuredClone(items),nextId:100};
    const fields=[
        {Title:'case subject',InternalName:'SubjectText',TypeAsString:'Text'},
        {Title:'classification',InternalName:'ClassificationText',TypeAsString:'Note'},
        {Title:'customer name',InternalName:'CustomerText',TypeAsString:'Text'},
        {Title:'processor',InternalName:'ProcessorText',TypeAsString:'Text'},
        {Title:'created by',InternalName:'CreatedText',TypeAsString:'Text'},
        {Title:'updated by',InternalName:'UpdatedText',TypeAsString:'Text'},
        {Title:'reopen count',InternalName:'ReopenNumber',TypeAsString:'Number'},
    ];
    const elements=new Map();
    function element(tag='div') {
        const el={tagName:tag,style:{},children:[],appendChild(child){this.children.push(child);if(child.id)elements.set(child.id,child);},remove(){if(this.id)elements.delete(this.id);},click(){if(this.onclick)this.onclick();},setAttribute(){},textContent:'',disabled:false};
        Object.defineProperty(el,'innerHTML',{set(html){this.html=html;for(const match of html.matchAll(/id="([^"]+)"/g))elements.set(match[1],element());},get(){return this.html;}});
        return el;
    }
    const body=element('body');
    const document={body,createElement:element,getElementById:id=>elements.get(id)||null,querySelector(selector){return selector==='.sp-toast'?null:null;}};
    const sandbox={URL,Blob,TextEncoder,TextDecoder,Uint8Array,ArrayBuffer,Date,JSON,Map,Set,Promise,document,location:{href:'https://example.sharepoint.com/sites/review/Lists/Cases'},window:{},navigator:{},
        MutationObserver:class{observe(){}},setTimeout:()=>0,clearTimeout(){},
        GM_addStyle(){},GM_getValue:(key,otherwise)=>store.get(key)??otherwise,GM_setValue:(key,value)=>store.set(key,value),
        console:{log(){},table(){},warn(...args){notices.push(args.join(' '));},error(...args){consoleErrors.push(args.join(' '));}},
        XLSX:{...XLSX,writeFile(book,filename){downloads.push({book,filename});}},
        GM_xmlhttpRequest(options){
            requests.push({method:options.method,url:options.url,headers:options.headers,data:options.data});
            setImmediate(()=>{
                const url=new URL(options.url);let status=200,response={};
                if(url.hostname!=='example.sharepoint.com')return options.onerror?.(new Error('Unexpected destination'));
                if(url.pathname.endsWith('/contextinfo'))response={FormDigestValue:'test-digest'};
                else if(url.pathname.endsWith('/fields'))response={value:fields};
                else if(options.method==='GET')response={value:state.items};
                else if(writeFailure){status=403;response={error:{message:'Test write denied'}};}
                else if(options.headers['X-HTTP-Method']==='MERGE'){
                    const id=Number(url.pathname.match(/items\((\d+)\)/)?.[1]);const item=state.items.find(i=>i.Id===id);
                    if(!item){status=404;}else{Object.assign(item,JSON.parse(options.data));status=204;}
                }else if(options.headers['X-HTTP-Method']==='DELETE'){
                    const id=Number(url.pathname.match(/items\((\d+)\)/)?.[1]);state.items=state.items.filter(i=>i.Id!==id);status=204;
                }else{const item={Id:state.nextId++,...JSON.parse(options.data)};state.items.push(item);status=201;response={Id:item.Id};}
                options.onload({status,responseText:JSON.stringify(response)});
            });
        },
    };
    const code=source.replace("siteUrl: ''","siteUrl: 'https://example.sharepoint.com/sites/review'").replace("listName: ''","listName: 'Cases'");
    const index=code.lastIndexOf('})();');
    vm.runInNewContext(code.slice(0,index)+`globalThis.api={${names.join(',')}};`+code.slice(index),sandbox,{timeout:5000});
    return {...sandbox.api,state,requests,downloads,notices,consoleErrors,store,document};
}
function fileForRows(rows,headers,sheet='Raw data'){
    const book=XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(book,XLSX.utils.json_to_sheet(rows,headers?{header:headers}:undefined),sheet);
    const buffer=XLSX.write(book,{bookType:'xlsx',type:'buffer'});
    return {arrayBuffer:async()=>buffer.buffer.slice(buffer.byteOffset,buffer.byteOffset+buffer.byteLength)};
}
function rawRows(){return [
    {'Number':'DEMO-001','Subject':'请提供使用指南','Name':'示例团队A'},
    {'Number':'DEMO-002','Subject':'请为协作者开通权限','Name':'示例团队B'},
    {'Number':'DEMO-003','Subject':'资料已收到但无法打开','Name':'示例团队C'},
    {'Number':'DEMO-004','Subject':'请提供操作说明','Name':'示例团队D'},
    {'Number':'DEMO-005','Subject':'还是和之前一样','Name':'示例团队E'},
    {'Number':'DEMO-006','Subject':'请提供使用指南并开通权限','Name':'示例团队F'},
];}
const progress={update(){}};

test('only fully explicit literal principles execute; semantic or mixed instructions go to AI',()=>{
    const e=environment();
    assert.deepEqual(plain(e.compilePrinciple('标题包含“使用指南”或“操作说明”')),{source:'subject',operation:'contains',phrases:['使用指南','操作说明']});
    assert.deepEqual(plain(e.compilePrinciple('描述等于“申请帮助”')),{source:'description',operation:'equals',phrases:['申请帮助']});
    for(const text of ['已经交付的资料无法使用，需要进一步处理。','标题包含“退款”但不包括历史问题','标题包含“退款”或','标题包含“退款”；忽略其他条件','标题包含“”'])assert.equal(e.compilePrinciple(text),null,text);
    assert.deepEqual(plain(e.classifyRecord('资料已收到但无法打开').topics),[]);
    assert.deepEqual(plain(e.classifyRecord('请提供使用指南').topics),['资料查询']);
    assert.equal(e.classifyRecord('请提供使用指南').needsAI,true);
});
test('priority only selects matched candidates; it never fills unmatched records',()=>{
    const e=environment();
    assert.deepEqual(plain(e.classifyRecord('请提供使用指南并开通权限').topics),['资料查询']);
    e.CLASSIFICATIONS[1].priority=1;e.CLASSIFICATIONS[0].priority=3;
    assert.deepEqual(plain(e.classifyRecord('请提供使用指南并开通权限').topics),['权限申请']);
    assert.deepEqual(plain(e.classifyRecord('信息不足').topics),[]);
});
test('only explicitly configured combinations are allowed',()=>{
    const e=environment();
    assert.deepEqual(plain(e.filterTopics(['资料查询','权限申请'])),['资料查询']);
    e.CLASSIFICATIONS[0].combinations=['权限申请'];e.CLASSIFICATIONS[1].combinations=['资料查询'];
    assert.deepEqual(plain(e.filterTopics(['权限申请','资料查询'])),['资料查询','权限申请']);
    e.CLASSIFICATIONS[1].combinations=[];
    assert.deepEqual(plain(e.filterTopics(['资料查询','权限申请'])),['资料查询']);
});
test('AI handoff exports the same standards and only the initial/final schema',()=>{
    const e=environment(),book=plain(e.buildSharedRulebook());
    assert.equal(book.classifications.length,3);
    assert.equal(book.automaticPrinciples[2].compiled,null);
    assert.equal(book.outputFormat.initialColumns.length,9);
    assert.equal(book.outputFormat.finalColumns.length,11);
    assert.ok(book.instructions.join(' ').includes('MCP'));
    assert.ok(!JSON.stringify(book).includes('example.sharepoint.com'));
});
test('real XLSX import → initial export → mocked AI final import → idempotent import → report',async()=>{
    const e=environment();
    await e.runImport(fileForRows(rawRows()));
    assert.equal(e.state.items.length,6);
    assert.deepEqual(e.state.items.map(i=>i.ClassificationText),['资料查询','权限申请','','资料查询','','资料查询']);
    assert.equal(e.consoleErrors.length,0);
    const map=await e.getFieldMap(),initial=e.cleanExportRows(e.state.items,map);
    assert.equal(initial.final,false);
    const initialBook=e.makeCleanWorkbook(initial.rows,false);
    const firstRound=XLSX.read(XLSX.write(initialBook,{type:'buffer',bookType:'xlsx'}));
    assert.deepEqual(XLSX.utils.sheet_to_json(firstRound.Sheets['Case Review'],{header:1})[0],plain(e.CLEAN_FIRST_COLUMNS));
    const finalRows=initial.rows.map(row=>({...row}));
    finalRows.find(r=>r['Case Number']==='DEMO-003')[e.CLEAN_FINAL]='交付跟进';
    finalRows.find(r=>r['Case Number']==='DEMO-003')[e.CLEAN_AI_KEYWORDS]='DEMO-003正文：已交付的资料无法打开';
    const finalBook=e.makeCleanWorkbook(finalRows,true),data=XLSX.write(finalBook,{type:'buffer',bookType:'xlsx'});
    const finalWorkbook=XLSX.read(data),rows=XLSX.utils.sheet_to_json(finalWorkbook.Sheets['Case Review'],{defval:'',raw:false});
    const result=await e.importCleanFinalRows(rows,progress);
    assert.equal(result.updated,1);assert.equal(result.unchanged,5);assert.equal(result.held.length,0);
    const again=await e.importCleanFinalRows(rows,progress);
    assert.equal(again.updated,0);assert.equal(again.unchanged,6);
    const after=e.cleanExportRows(e.state.items,map);assert.equal(after.final,true);
    const changed=after.rows.find(r=>r['Case Number']==='DEMO-003');
    assert.equal(changed[e.CLEAN_INITIAL],'');assert.equal(changed[e.CLEAN_FINAL],'交付跟进');
    assert.equal(changed[e.CLEAN_SCRIPT_KEYWORDS],'');assert.match(changed[e.CLEAN_AI_KEYWORDS],/DEMO-003/);
    const stats=e.buildDashboardData(e.state.items.map(i=>({'case number':i.Title,'case subject':i.SubjectText,classification:i.ClassificationText})));
    assert.equal(stats.total,6);assert.equal(stats.unresolvedTopicCount,1);
    assert.equal(stats.topicSorted.find(([n])=>n==='交付跟进')[1],1);
    assert.ok(e.requests.every(r=>new URL(r.url).hostname==='example.sharepoint.com'));
});
test('changed AI result requires evidence; unknown labels, tampered baseline and stale data are rejected',()=>{
    const e=environment();const saved={subject:'请提供使用指南',initial:'资料查询',scriptKeywords:'使用指南',rules:e.rulesSignature(),reviewed:false};
    const context={subject:saved.subject,customer:'示例团队',topics:'资料查询'};
    const row={'case subject':saved.subject,'customer name':'示例团队',[e.CLEAN_INITIAL]:'资料查询',[e.CLEAN_FINAL]:'交付跟进',[e.CLEAN_SCRIPT_KEYWORDS]:'使用指南',[e.CLEAN_AI_KEYWORDS]:''};
    assert.throws(()=>e.validateCleanFinalRow(row,context,saved),/AI evidence/);
    row[e.CLEAN_AI_KEYWORDS]='实际已交付资料无法使用';assert.equal(e.validateCleanFinalRow(row,context,saved).changed,true);
    assert.throws(()=>e.validateCleanFinalRow({...row,[e.CLEAN_FINAL]:'未定义标签'},context,saved),/Unknown/);
    assert.throws(()=>e.validateCleanFinalRow({...row,[e.CLEAN_INITIAL]:''},context,saved),/Initial/);
    assert.throws(()=>e.validateCleanFinalRow({...row,[e.CLEAN_SCRIPT_KEYWORDS]:'伪造'},context,saved),/Script evidence/);
    assert.throws(()=>e.validateCleanFinalRow(row,{...context,topics:'权限申请'},saved),/changed since/);
    e.CLASSIFICATIONS[0].principle='标题包含“新指南”';assert.throws(()=>e.validateCleanFinalRow(row,context,saved),/principles changed/);
});
test('explicit blank in final sheet clears a previously assigned classification',async()=>{
    const e=environment();await e.runImport(fileForRows([rawRows()[0]]));
    const map=await e.getFieldMap(),initial=e.cleanExportRows(e.state.items,map),row=initial.rows[0];row[e.CLEAN_FINAL]='';row[e.CLEAN_AI_KEYWORDS]='';
    const result=await e.importCleanFinalRows([Object.fromEntries(e.CLEAN_FINAL_COLUMNS.map(n=>[n,row[n]??'']))],progress);
    assert.equal(result.updated,1);assert.equal(e.state.items[0].ClassificationText,'');
});
test('wrong template and duplicate final IDs fail before any write',async()=>{
    const e=environment();await e.runImport(fileForRows([rawRows()[0]]));const map=await e.getFieldMap(),initial=e.cleanExportRows(e.state.items,map);
    const before=e.requests.length;
    await assert.rejects(e.importCleanFinalRows([Object.fromEntries(e.CLEAN_FIRST_COLUMNS.map(n=>[n,initial.rows[0][n]]))],progress),/final template/);
    const row=Object.fromEntries(e.CLEAN_FINAL_COLUMNS.map(n=>[n,initial.rows[0][n]??'']));
    await assert.rejects(e.importCleanFinalRows([row,row],progress),/duplicate/);
    assert.equal(e.requests.length,before);
});
test('raw duplicate IDs use last occurrence; missing IDs are omitted',async()=>{
    const e=environment();await e.runImport(fileForRows([rawRows()[0],{Number:'',Subject:'缺编号'},{...rawRows()[0],Subject:'请开通权限'}]));
    assert.equal(e.state.items.length,1);assert.equal(e.state.items[0].ClassificationText,'权限申请');
});
test('failed SharePoint write does not persist a successful classification baseline',async()=>{
    const e=environment({writeFailure:true});await e.runImport(fileForRows([rawRows()[0]]));
    assert.equal(e.state.items.length,0);assert.deepEqual(plain(e.readCleanState().records),{});
});
test('counting modes preserve raw labels and distinguish unknown / unclassified values',()=>{
    const e=environment(),rows=[{'case number':'a',classification:'资料查询, 权限申请'},{'case number':'b',classification:''},{'case number':'c',classification:'非法标签'}];
    const before=JSON.stringify(rows),primary=e.buildDashboardData(rows,'primary'),all=e.buildDashboardData(rows,'all');
    assert.equal(primary.topicSorted.reduce((n,[,count])=>n+count,0),3);
    assert.equal(all.topicSorted.reduce((n,[,count])=>n+count,0),4);
    assert.equal(all.unresolvedTopicCount,2);assert.equal(all.invalidTopicCases,1);assert.equal(JSON.stringify(rows),before);
});
test('dashboard keeps generic chart and escapes user category text',()=>{
    const e=environment();e.CLASSIFICATIONS[0].name='<img src=x onerror=alert(1)>';
    const stats=e.buildDashboardData([{classification:e.CLASSIFICATIONS[0].name}]);const html=e.buildReportHTMLV2(stats);
    assert.ok(!html.includes('<img src=x'));assert.ok(html.includes('&lt;img'));
    assert.equal((html.match(/class="report-page"/g)||[]).length,1);
    assert.ok(html.includes('v2-hbar'));assert.ok(!html.includes('v2-init'));assert.ok(!html.includes('v2-act'));
});
test('public script contains no company rule packages, case references, region or source/activity classifiers',()=>{
    for(const pattern of [/sap/i,/SNOW/,/PC3/,/CDM/i,/non.?apj/i,/CASE_REVIEW_REFERENCES/,/UNIFIED_REVIEW_RULES/,/classifyIntExt/,/classifyCaseActivity/,/RITM/,/PTASK/,/CIC\d+/,/internal\/external/,/FFFFAD49/,/FFFFFF00/])assert.ok(!pattern.test(source),String(pattern));
    assert.ok(source.includes("siteUrl: ''"));assert.ok(source.includes("listName: ''"));
});
