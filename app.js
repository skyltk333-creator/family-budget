'use strict';
// No network requests, external scripts, or backend. XLSX ZIP decompression uses browser DecompressionStream.
const $=id=>document.getElementById(id), KRW=n=>Math.round(n).toLocaleString('ko-KR')+'원';
const CATS=['식비','생활비','교통비','관리비','문화생활비','확인 필요','비용 제외'];
let data=[], overrides=new Map();
const clean=s=>String(s??'').trim();
const money=s=>Number(clean(s).replace(/[₩원,\s]/g,''))||0;
function classify(name,type='',source=''){
 const s=(name+' '+type).toLowerCase().replace(/\s/g,'');
 if(source==='토스뱅크' && /입금|환급|이자수입/.test(type))return ['비용 제외','입금'];
 if(/카드대금|카드결제대금|신한카드대금|자동이체.*카드|결제대금|본인계좌|적금|예금|증권|주식|펀드|보험료대출|대출상환|원금상환/.test(s))return ['비용 제외','이체·중복 후보'];
 if(/취소|환불/.test(type))return ['비용 제외','취소·환불'];
 if(/쿠팡|네이버페이|카카오페이|토스페이|페이코|스마트스토어|11번가|지마켓|옥션|마켓컬리/.test(s))return ['확인 필요','구매 목적 불명'];
 if(/주유|gs칼텍스|s-oil|에쓰오일|sk에너지|하이패스|주차|파킹|택시|카카오t|카카오택시|대리운전/.test(s))return ['교통비','자동 분류'];
 if(/도시가스|한국전력|전기요금|수도요금|상하수도|관리비|통신요금|sk텔레콤|kt통신|lg유플러스|인터넷요금/.test(s))return ['관리비','자동 분류'];
 if(/약국|의원|병원|치과|한의원|올리브영|다이소|미용실|헤어|세탁|정비|세차|자동차보험/.test(s))return ['생활비','자동 분류'];
 if(/영화|cgv|메가박스|롯데시네마|노래방|pc방|game|게임|볼링|골프|숙박|호텔|펜션|여행|티켓|공연|놀이공원/.test(s))return ['문화생활비','자동 분류'];
 if(/스타벅스|투썸|이디야|메가커피|컴포즈|빽다방|카페|커피|베이커리|파리바게뜨|뚜레쥬르|마트|이마트|홈플러스|롯데마트|하나로마트|식자재|편의점|gs25|cu|세븐일레븐|식당|김밥|치킨|피자|버거|짬뽕|국밥|배달|요기요|배민|쿠팡이츠|맥도날드|맘스터치|배스킨|분식|초밥|횟집|고기집/.test(s))return ['식비','자동 분류'];
 return ['확인 필요','분류 근거 부족'];
}
// Minimal XLSX reader: parses Office Open XML ZIP without sending files to a server.
async function unzipEntries(buffer){
 const bytes=new Uint8Array(buffer),view=new DataView(buffer),decoder=new TextDecoder(),out=new Map();
 let end=-1;for(let p=bytes.length-22;p>=Math.max(0,bytes.length-65557);p--){if(view.getUint32(p,true)===0x06054b50){end=p;break;}}
 if(end<0)throw Error('정상적인 XLSX 파일이 아닙니다.');
 let pos=view.getUint32(end+16,true),count=view.getUint16(end+10,true);
 for(let i=0;i<count;i++){
  if(view.getUint32(pos,true)!==0x02014b50)throw Error('ZIP 목차를 읽지 못했습니다.');
  const method=view.getUint16(pos+10,true),compressed=view.getUint32(pos+20,true),n=view.getUint16(pos+28,true),ex=view.getUint16(pos+30,true),comment=view.getUint16(pos+32,true),offset=view.getUint32(pos+42,true),name=decoder.decode(bytes.slice(pos+46,pos+46+n));
  if(name==='xl/sharedStrings.xml'||/^xl\/worksheets\/sheet\d+\.xml$/.test(name)){
   const start=offset+30+view.getUint16(offset+26,true)+view.getUint16(offset+28,true),chunk=bytes.slice(start,start+compressed);
   let content;
   if(method===0)content=chunk;
   else if(method===8){if(typeof DecompressionStream==='undefined')throw Error('이 브라우저는 XLSX 압축 해제를 지원하지 않습니다. 최신 Chrome 또는 Edge를 이용해 주세요.');const stream=new Blob([chunk]).stream().pipeThrough(new DecompressionStream('deflate-raw'));content=new Uint8Array(await new Response(stream).arrayBuffer());}
   else throw Error('지원하지 않는 엑셀 압축 방식입니다.');
   out.set(name,decoder.decode(content));
  }
  pos+=46+n+ex+comment;
 }
 return out;
}
function parseXML(s){const doc=new DOMParser().parseFromString(s,'application/xml');if(doc.querySelector('parsererror'))throw Error('엑셀 XML 구조를 읽지 못했습니다.');return doc;}
function textNodes(el,tag){return Array.from(el.getElementsByTagName(tag)).map(x=>x.textContent).join('');}
function rowsFromXML(entries){
 const shared=entries.has('xl/sharedStrings.xml')?Array.from(parseXML(entries.get('xl/sharedStrings.xml')).getElementsByTagName('si')).map(x=>textNodes(x,'t')):[];
 const sheet=entries.get('xl/worksheets/sheet1.xml')||[...entries].find(([k])=>/^xl\/worksheets\/sheet\d+\.xml$/.test(k))?.[1];
 if(!sheet)throw Error('엑셀에서 시트를 찾지 못했습니다.');
 const doc=parseXML(sheet),rows=[];
 for(const r of doc.getElementsByTagName('row')){
  const a=[];for(const c of r.getElementsByTagName('c')){
   const ref=c.getAttribute('r')||'A1',letters=ref.match(/^[A-Z]+/)?.[0]||'A';let idx=0;for(const letter of letters)idx=idx*26+letter.charCodeAt(0)-64;idx--;
   const v=c.getElementsByTagName('v')[0]?.textContent||'',type=c.getAttribute('t');a[idx]=type==='s'?shared[Number(v)]||'':type==='inlineStr'?textNodes(c,'t'):v;
  }rows.push(a);
 }return rows;
}
function dateValue(s){let x=clean(s);if(/^\d{4}[.\-/]\d{2}[.\-/]\d{2}/.test(x))return x.slice(0,10).replace(/[./]/g,'-');if(/^\d{5}(\.\d+)?$/.test(x)){const d=new Date(Date.UTC(1899,11,30)+Number(x)*86400000);return d.toISOString().slice(0,10);}return '';}
function readShinhan(rows){const h=rows.findIndex(r=>r.includes('거래일')&&r.includes('가맹점명'));if(h<0)throw Error('신한카드 파일의 제목 행을 찾지 못했습니다.');const head=rows[h],at=n=>head.indexOf(n);return rows.slice(h+1).map((r,i)=>{
 const date=dateValue(r[at('거래일')]),name=clean(r[at('가맹점명')]),amount=Math.abs(money(r[at('금액')])),cancel=clean(r[at('취소상태')]);if(!date||!name||!amount)return null;
 const [cat,reason]=/취소/.test(cancel)?['비용 제외','카드 취소 표시']:classify(name,clean(r[at('이용구분')]),'신한카드');return {id:`S:${date}:${clean(r[at('승인번호')])||i}:${amount}`,date,name,amount,source:'신한카드',category:cat,reason,rawType:cancel};
 }).filter(Boolean);}
function readToss(rows){const h=rows.findIndex(r=>r.includes('거래 일시')&&r.includes('거래 금액'));if(h<0)throw Error('토스뱅크 파일의 제목 행을 찾지 못했습니다.');const head=rows[h],at=n=>head.indexOf(n);return rows.slice(h+1).map((r,i)=>{
 const date=dateValue(r[at('거래 일시')]),name=clean(r[at('적요')]),signed=money(r[at('거래 금액')]),type=clean(r[at('거래 유형')]);if(!date||!name||!signed)return null;
 let [category,reason]=signed>0?['비용 제외','입금·환급']:classify(name,type,'토스뱅크');
 if(signed<0&&/카드대금|신한카드|카드결제금액|카드이용대금/.test(name+type)){category='비용 제외';reason='카드 청구 중복 후보';}
 if(signed<0&&/이체|송금|출금/.test(type)&&category==='확인 필요')reason='계좌이체 목적 확인';
 return {id:`T:${date}:${i}:${Math.abs(signed)}`,date,name,amount:Math.abs(signed),source:'토스뱅크',category,reason,rawType:type};
 }).filter(Boolean);}
function dedupe(list){const seen=new Set();return list.filter(x=>{if(seen.has(x.id))return false;seen.add(x.id);return true;});}
function applyOverrides(){for(const x of data){if(overrides.has(x.source+'|'+x.name)&&x.reason!=='카드 취소 표시'&&x.reason!=='입금·환급'){x.category=overrides.get(x.source+'|'+x.name);x.reason='사용자 지정';}}}
function months(){return [...new Set(data.map(x=>x.date.slice(0,7)))].sort().reverse();}
function selected(){const m=$('month').value;return data.filter(x=>x.date.startsWith(m));}
function render(){if(!data.length){$('results').hidden=true;return;}$('results').hidden=false;const mm=months(),old=$('month').value;$('month').replaceChildren(...mm.map(m=>new Option(m,m)));$('month').value=mm.includes(old)?old:mm[0];renderMonth();}
function renderMonth(){const items=selected(),expense=items.filter(x=>!['확인 필요','비용 제외'].includes(x.category)),spent=expense.reduce((a,x)=>a+x.amount,0),pending=items.filter(x=>x.category==='확인 필요').reduce((a,x)=>a+x.amount,0);$('spent').textContent=KRW(spent);$('pending').textContent=KRW(pending);$('excluded').textContent=items.filter(x=>x.category==='비용 제외').length+'건';$('count').textContent=items.length+'건';
 $('bars').replaceChildren(...CATS.filter(x=>!['확인 필요','비용 제외'].includes(x)).map(c=>{const n=expense.filter(x=>x.category===c).reduce((a,x)=>a+x.amount,0),r=document.createElement('div');r.className='barrow';const title=document.createElement('span');title.textContent=c;const track=document.createElement('div');track.className='track';const fill=document.createElement('div');fill.className='fill';fill.style.width=(spent?100*n/spent:0)+'%';track.append(fill);const val=document.createElement('span');val.textContent=KRW(n);r.append(title,track,val);return r;}));
 const monthly=months().slice().reverse(),values=monthly.map(m=>data.filter(x=>x.date.startsWith(m)&&!['확인 필요','비용 제외'].includes(x.category)).reduce((a,x)=>a+x.amount,0)),max=Math.max(...values,1);$('trend').replaceChildren(...monthly.map((m,i)=>{const e=document.createElement('div');e.className='trenditem';const t=document.createElement('div');t.textContent=(values[i]/10000).toFixed(0)+'만';const bar=document.createElement('div');bar.className='trendbar';bar.style.height=Math.max(2,90*values[i]/max)+'px';const label=document.createElement('div');label.textContent=m;e.append(t,bar,label);return e;}));renderRows();}
function renderRows(){const query=$('search').value.toLowerCase(),filter=$('filter').value;const list=selected().filter(x=>(filter==='all'||x.category===filter)&&x.name.toLowerCase().includes(query)).sort((a,b)=>b.date.localeCompare(a.date));const tbody=$('transactions');tbody.replaceChildren();for(const x of list.slice(0,350)){
 const tr=document.createElement('tr');for(const v of [x.date,x.source,x.name,KRW(x.amount)]){const td=document.createElement('td');td.textContent=v;if(v===KRW(x.amount))td.className='amount';tr.append(td);}const td=document.createElement('td'),select=document.createElement('select');for(const c of CATS)select.add(new Option(c,c));select.value=x.category;select.addEventListener('change',()=>{overrides.set(x.source+'|'+x.name,select.value);applyOverrides();renderMonth();});td.append(select);tr.append(td);const status=document.createElement('td');status.textContent=x.reason;status.className=x.category==='확인 필요'?'flag':'muted';tr.append(status);tbody.append(tr);
 }$('shown').textContent=`${list.length}건 중 최대 350건 표시 · CSV에는 필터된 전체 거래가 포함됩니다.`;}
function exportCSV(){const filter=$('filter').value,query=$('search').value.toLowerCase();const rows=selected().filter(x=>(filter==='all'||x.category===filter)&&x.name.toLowerCase().includes(query));const cell=v=>'"'+String(v??'').replace(/"/g,'""')+'"';const csv='\ufeff'+[['날짜','출처','거래처','금액(원)','분류','상태'],...rows.map(x=>[x.date,x.source,x.name,x.amount,x.category,x.reason])].map(r=>r.map(cell).join(',')).join('\r\n');const url=URL.createObjectURL(new Blob([csv],{type:'text/csv;charset=utf-8'}));const a=document.createElement('a');a.href=url;a.download=`우리집가계부_${$('month').value}.csv`;a.click();setTimeout(()=>URL.revokeObjectURL(url),1000);}
$('analyze').addEventListener('click',async()=>{const s=$('shinhan').files[0],t=$('toss').files[0];if(!s&&!t){$('status').textContent='신한카드 또는 토스뱅크 파일을 선택해 주세요.';return;}try{$('status').textContent='엑셀을 분석하고 있습니다…';let all=[];for(const [f,reader] of [[s,readShinhan],[t,readToss]])if(f){if(!f.name.toLowerCase().endsWith('.xlsx'))throw Error('.xlsx 형식만 지원합니다.');if(f.size>15*1024*1024)throw Error('파일 크기는 15MB 이하만 지원합니다.');const entries=await unzipEntries(await f.arrayBuffer());all.push(...reader(rowsFromXML(entries)));}data=dedupe(all);overrides.clear();$('filter').replaceChildren(new Option('전체 분류','all'),...CATS.map(c=>new Option(c,c)));$('status').textContent=`분석 완료: ${data.length.toLocaleString()}건 · 확인 필요 ${data.filter(x=>x.category==='확인 필요').length}건. 파일 원본은 서버로 전송되지 않았습니다.`;render();}catch(e){$('status').textContent='분석 실패: '+e.message;console.error(e);}});
$('demo').addEventListener('click',()=>{const names=['식자재마트','스타벅스','메트로주유소','쿠팡페이','상가 관리비','CGV','신한카드대금','다이소'];data=Array.from({length:90},(_,i)=>{const month=String(7+Math.floor(i/30)).padStart(2,'0'),name=names[i%names.length],amount=10000+(i%11)*6900,[category,reason]=classify(name,'','토스뱅크');return {id:'D'+i,date:`2026-${month}-${String(i%28+1).padStart(2,'0')}`,name,amount,source:i%2?'토스뱅크':'신한카드',category,reason};});overrides.clear();$('filter').replaceChildren(new Option('전체 분류','all'),...CATS.map(c=>new Option(c,c)));$('status').textContent='가상 데이터 90건을 불러왔습니다. 실제 금융정보가 아닙니다.';render();});
$('clear').addEventListener('click',()=>{data=[];overrides.clear();$('shinhan').value='';$('toss').value='';$('results').hidden=true;$('status').textContent='분석 데이터를 초기화했습니다.';});
$('month').addEventListener('change',renderMonth);$('search').addEventListener('input',renderRows);$('filter').addEventListener('change',renderRows);$('export').addEventListener('click',exportCSV);
