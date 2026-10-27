const {test}=require('node:test');
const assert=require('node:assert/strict');
const fs=require('node:fs');
const vm=require('node:vm');
const path=require('node:path');
const code=fs.readFileSync(path.join(__dirname,'../app.js'),'utf8');
const key='family-budget.rules.v1';
function boot(storage=new Map(),blocked=false){
 const nodes=new Map();
 function element(){return {value:'',textContent:'',children:[],style:{},handlers:{},files:[],append(...children){this.children.push(...children);},add(child){this.append(child);},replaceChildren(...children){this.children=children;},addEventListener(event,fn){this.handlers[event]=fn;},click(){}};}
 const get=id=>{if(!nodes.has(id))nodes.set(id,element());return nodes.get(id);};
 let blob;
 const context=vm.createContext({document:{getElementById:get,createElement:element},localStorage:{getItem(k){if(blocked)throw Error('blocked');return storage.get(k)||null;},setItem(k,v){if(blocked)throw Error('blocked');storage.set(k,v);}},Option:function(text,value){return {textContent:text,value};},Blob,URL:{createObjectURL(b){blob=b;return 'blob:test';},revokeObjectURL(){}},setTimeout:f=>f(),console:{error(){}}});
 vm.runInContext(code,context);
 get('month').value='2026-10';get('filter').value='all';get('reason-filter').value='all';get('transaction-sort').value='date-desc';
 return {get,storage,run:source=>vm.runInContext(source,context),csv:()=>blob.text()};
}
function transaction(name='상점',extra={}){return {date:'2026-10-01',name,source:'신한카드',amount:10000,category:'확인 필요',reason:'분류 근거 부족',...extra};}
function setData(app,rows){app.run(`data=${JSON.stringify(rows)}`);}
function changeFirst(app,category='식비'){
 app.run('renderRows()');
 const select=app.get('transactions').children[0].children[4].children[0];
 select.value=category;select.handlers.change();
}
test('manual change persists only the rule, restores on reload and protects excluded transactions',()=>{
 const app=boot();setData(app,[transaction(),transaction(),transaction('상점',{source:'토스뱅크'}),transaction('상점',{reason:'카드 취소 표시',category:'비용 제외'})]);changeFirst(app);
 assert.equal(app.run('data[1].category'),'식비');
 assert.equal(app.run('data[2].category'),'확인 필요');
 assert.equal(app.run('data[3].category'),'비용 제외');
 assert.equal(app.run('data[0].updatedAt'),app.run('data[1].updatedAt'));
 const saved=JSON.parse(app.storage.get(key));
 assert.deepEqual(Object.keys(saved[0][1]).sort(),['category','updatedAt']);
 assert.equal(saved[0][0],'신한카드|상점');
 const next=boot(app.storage);setData(next,[transaction(),transaction('상점',{reason:'입금·환급',category:'비용 제외'})]);next.run('applyOverrides()');
 assert.equal(next.run('data[0].category'),'식비');assert.equal(next.run('data[1].category'),'비용 제외');
 assert.equal(next.run('data[0].updatedAt'),app.run('data[0].updatedAt'));
 next.get('clear').handlers.click();assert.equal(next.run('data.length'),0);assert.equal(next.storage.get(key),app.storage.get(key));
});
test('real import automatically applies saved rules and time',async()=>{
 const storage=new Map([[key,JSON.stringify([['신한카드|상점',{category:'생활비',updatedAt:'2026-10-01T01:02:03.000Z'}]])]]);
 const app=boot(storage);
 app.run('unzipEntries=async b=>b;rowsFromXML=b=>b;');
 const rows=[['거래일','가맹점명','금액'],['2026-10-01','상점','10000']];
 app.get('statements').files=[{name:'statement.xlsx',size:100,arrayBuffer:async()=>rows}];
 await app.get('analyze').handlers.click();
 assert.equal(app.run('data[0].category'),'생활비');assert.equal(app.run('data[0].updatedAt'),'2026-10-01T01:02:03.000Z');
 assert.match(app.get('status').textContent,/확인 필요 0건/);
});
test('both timestamp directions keep unedited rows last; recent button clears filters',()=>{
 const app=boot();setData(app,[transaction('없음'),transaction('이전',{reason:'사용자 지정',updatedAt:'2026-10-01T01:00:00Z'}),transaction('최근',{reason:'사용자 지정',updatedAt:'2026-10-02T01:00:00Z'})]);
 app.get('transaction-sort').value='updated-desc';assert.equal(app.run('sortedTransactions().map(x=>x.name).join()'),'최근,이전,없음');
 app.get('transaction-sort').value='updated-asc';assert.equal(app.run('sortedTransactions().map(x=>x.name).join()'),'이전,최근,없음');
 app.get('search').value='찾을수없음';app.get('reason-filter').value='구매 목적 불명';app.get('filter').value='식비';
 app.get('recent-edits').handlers.click();assert.equal(app.run('sortedTransactions().map(x=>x.name).join()'),'최근,이전');
});
test('CSV contains the displayed timestamps in selected order',async()=>{
 const app=boot();setData(app,[transaction('이전',{updatedAt:'2026-10-01T01:00:00Z'}),transaction('최근',{updatedAt:'2026-10-02T01:00:00Z'})]);
 app.get('transaction-sort').value='updated-desc';app.run('exportCSV()');const csv=await app.csv();
 assert.ok(csv.includes('분류 수정 일시'));assert.ok(csv.indexOf('최근')<csv.indexOf('이전'));
 assert.ok(csv.includes(app.run("formatUpdatedAt('2026-10-02T01:00:00Z')")));
});
test('unavailable or corrupt storage does not break classification',()=>{
 for(const app of [boot(new Map(),true),boot(new Map([[key,'broken JSON']])),boot(new Map([[key,'[["bad",{"category":"invalid","updatedAt":"bad"}]]']]))]){
  assert.match(app.get('rules-status').textContent,/읽을 수 없습니다/);
  setData(app,[transaction()]);changeFirst(app);assert.equal(app.run('data[0].category'),'식비');
 }
 const blocked=boot(new Map(),true);setData(blocked,[transaction()]);changeFirst(blocked);assert.match(blocked.get('rules-status').textContent,/저장에 실패/);
});
test('demo changes never overwrite saved real rules',()=>{
 const storage=new Map([[key,JSON.stringify([['신한카드|상점',{category:'생활비',updatedAt:'2026-10-01T01:02:03Z'}]])]]);
 const before=storage.get(key),app=boot(storage);app.get('demo').handlers.click();changeFirst(app,'교통비');
 assert.equal(storage.get(key),before);assert.match(app.get('rules-status').textContent,/저장하지 않습니다/);
 app.get('clear').handlers.click();assert.equal(app.run("overrides.get('신한카드|상점').category"),'생활비');
});
