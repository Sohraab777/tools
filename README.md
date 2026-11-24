<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">
<title>Result Cross-Verification Tool</title>
<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/exceljs/4.3.0/exceljs.min.js"></script>
<link href="https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;500&family=IBM+Plex+Sans:wght@400;500;600;700&family=Fraunces:ital,opsz,wght@0,9..144,500;0,9..144,600;1,9..144,400&display=swap" rel="stylesheet">
<style>
:root{--paper:#eceae3;--card:#f8f7f3;--card2:#fdfdfb;--ink:#111317;--soft:#3a3d44;--muted:#6f7078;--line:#d7d5cc;--strong:#b9b6aa;--acc:#1436d1;--wash:rgba(20,54,209,.07);
--ok:#2f6b3c;--okbg:#e2ece1;--bad:#a52f16;--badbg:#f7e3de;--warn:#836517;--warnbg:#f6ecd4;
--mono:"IBM Plex Mono",ui-monospace,Menlo,Consolas,monospace;--sans:"IBM Plex Sans",system-ui,-apple-system,"Segoe UI",Roboto,Arial,sans-serif;color-scheme:light}
@media(prefers-color-scheme:dark){:root:not([data-theme="light"]){--paper:#0f1114;--card:#16191d;--card2:#1c2025;--ink:#e9e7e1;--soft:#c3c1bb;--muted:#8b9098;--line:#282d34;--strong:#3d434c;--acc:#7d97ff;--wash:rgba(125,151,255,.1);--ok:#7ad392;--okbg:#16281c;--bad:#ff8b6c;--badbg:#2e1714;--warn:#e8c063;--warnbg:#2b2313;color-scheme:dark}}
:root[data-theme="dark"]{--paper:#0f1114;--card:#16191d;--card2:#1c2025;--ink:#e9e7e1;--soft:#c3c1bb;--muted:#8b9098;--line:#282d34;--strong:#3d434c;--acc:#7d97ff;--wash:rgba(125,151,255,.1);--ok:#7ad392;--okbg:#16281c;--bad:#ff8b6c;--badbg:#2e1714;--warn:#e8c063;--warnbg:#2b2313;color-scheme:dark}
:root{box-sizing:border-box;padding-top:env(safe-area-inset-top,0px);padding-bottom:env(safe-area-inset-bottom,0px)}
html{scroll-padding-top:env(safe-area-inset-top,0px)}
*{box-sizing:border-box}
body{margin:0;background:var(--paper);color:var(--ink);font:14px/1.5 var(--sans);-webkit-font-smoothing:antialiased}
.wrap{max-width:1280px;margin:0 auto;padding:0 26px 90px}
header{border-bottom:1px solid var(--ink);padding:46px 0 30px}
.idx{font:600 12px Fraunces,Georgia,"Times New Roman",serif;letter-spacing:.2em;text-transform:uppercase;color:var(--acc);margin:0 0 18px}
h1{margin:0;font-family:Fraunces,Georgia,"Times New Roman",serif;font-size:clamp(32px,5vw,58px);line-height:1.02;letter-spacing:-.02em;font-weight:600;max-width:18ch}
h1 em{font-family:Fraunces,Georgia,serif;font-style:italic;font-weight:400;color:var(--acc)}
.lede{max-width:62ch;margin:20px 0 0;color:var(--soft);font-size:15px}
.grid{display:grid;grid-template-columns:1fr 1fr;gap:1px;background:var(--line);border:1px solid var(--line);margin-top:34px}
.card{background:var(--card);padding:24px}
.step{font:500 10px var(--mono);letter-spacing:.2em;text-transform:uppercase;color:var(--acc)}
.card h2{margin:10px 0 16px;font-family:Fraunces,Georgia,"Times New Roman",serif;font-size:22px;font-weight:600;letter-spacing:-.01em}
.drop{display:block;border:1px dashed var(--strong);background:var(--card2);padding:20px 18px;cursor:pointer;transition:.18s}
.drop:hover,.drop.hot{border-color:var(--acc);background:var(--wash)}
.drop strong{display:block;font-size:15px}.drop small{display:block;margin-top:6px;color:var(--muted);font:400 11px/1.5 var(--mono)}
input[type=file]{display:none}
.pill{display:none;margin-top:12px;padding:7px 12px;border:1px solid var(--strong);background:var(--card2);font:500 11.5px var(--mono)}
.pill.on{display:inline-block}.pill.err{color:var(--bad);border-color:var(--bad)}
.bar{display:flex;gap:10px;flex-wrap:wrap;align-items:center;margin:22px 0 0}
button{font:500 11.5px var(--mono);letter-spacing:.14em;text-transform:uppercase;padding:14px 20px;border:1px solid var(--ink);cursor:pointer;border-radius:0;transition:.16s}
button.p{background:var(--ink);color:var(--paper)}button.p:hover:not(:disabled){background:var(--acc);border-color:var(--acc);color:#fff}
button.g{background:transparent;color:var(--ink);border-color:var(--strong)}button.g:hover:not(:disabled){border-color:var(--acc);color:var(--acc)}
button:disabled{opacity:.4;cursor:not-allowed}
#msg{font:400 11.5px var(--mono);color:var(--muted)}
.note{border-left:2px solid var(--acc);padding:2px 0 2px 14px;margin:26px 0 0;color:var(--soft);font-size:12.5px}
#res{display:none;margin-top:44px}
.kpis{display:grid;grid-template-columns:repeat(6,1fr);border-top:1px solid var(--ink);border-bottom:1px solid var(--ink);background:var(--card)}
.kpi{padding:16px;border-right:1px solid var(--line)}.kpi:last-child{border:0}
.kpi b{display:block;font:500 28px var(--mono);letter-spacing:-.03em}.kpi span{font:400 9.5px var(--mono);letter-spacing:.14em;text-transform:uppercase;color:var(--muted)}
.kpi.ok b{color:var(--ok)}.kpi.bad b{color:var(--bad)}.kpi.warn b{color:var(--warn)}
.tabs{display:flex;flex-wrap:wrap;border-bottom:1px solid var(--ink);margin-top:26px}
.tab{padding:12px 15px;font:500 10.5px var(--mono);letter-spacing:.12em;text-transform:uppercase;color:var(--muted);cursor:pointer;border-bottom:2px solid transparent;margin-bottom:-1px}
.tab.on{color:var(--ink);border-color:var(--acc)}
.scroll{max-height:520px;overflow:auto;border:1px solid var(--line);border-top:0;background:var(--card)}
table{border-collapse:collapse;width:100%;font-size:12.5px}
th{position:sticky;top:0;background:var(--card);text-align:left;padding:10px 12px;font:500 9.5px var(--mono);letter-spacing:.12em;text-transform:uppercase;border-bottom:1.5px solid var(--ink);white-space:nowrap}
td{padding:8px 12px;border-bottom:1px solid var(--line);white-space:nowrap}
tr:hover td{background:var(--wash)}
td.ok{background:var(--okbg);color:var(--ok)}td.bad{background:var(--badbg);color:var(--bad);font-weight:600}td.warn{background:var(--warnbg);color:var(--warn)}
@media(max-width:860px){.grid{grid-template-columns:1fr}.kpis{grid-template-columns:repeat(2,1fr)}.kpi{border-bottom:1px solid var(--line)}}
</style>
</head>
<body>
<div class="wrap">
<header>
<p class="idx">Cross-verification</p>
<h1>Match two result sheets, <em>field by field.</em></h1>
<p class="lede">Load the Consolidated Report and the TRU student-transcript sheet. Students are matched by roll number, then names, grades, marks, credits, SGPA and CGPA are compared. Everything runs in your browser; no file leaves this device.</p>
<div class="grid">
<div class="card"><div class="step">Step 1</div><h2>Consolidated Report</h2>
<label class="drop" id="d1"><strong>Choose the Consolidated Report (.xlsx)</strong><small>One row per student per course. Needs a “Registration ID” column.</small><input type="file" id="f1" accept=".xlsx,.xls"></label><div class="pill" id="p1"></div></div>
<div class="card"><div class="step">Step 2</div><h2>TRU student transcripts</h2>
<label class="drop" id="d2"><strong>Choose the TRU sheet (.xlsx)</strong><small>One row per student, subjects across. Needs an “ENROLLMENT_NO” column.</small><input type="file" id="f2" accept=".xlsx,.xls"></label><div class="pill" id="p2"></div></div>
</div>
<div class="bar"><button class="p" id="run" disabled>Run verification</button><button class="g" id="dl" disabled>Download Excel report</button><button class="g" id="clr">Clear</button><span id="msg">Load both sheets to begin.</span></div>
<p class="note">Rules: match key is the roll number · grade, grade point and total marks must be equal · SGPA may differ only by rounding (two decimals against one) · credit points are grade point × credits (credits come from the TRU “3-0-0” structure) · percentage is the average of the subject totals · a blank CGPA is reported as not comparable.</p>
</header>
<div id="res">
<div class="kpis" id="kpis"></div>
<div class="tabs" id="tabs"></div>
<div id="panel"></div>
</div>
</div>
<script>
const $=s=>document.querySelector(s);
const esc=v=>String(v==null?'':v).replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const key=v=>{let s=String(v==null?'':v).trim();if(/^\d+(\.0+)?$/.test(s))s=s.replace(/\.0+$/,'');return s};
const norm=s=>String(s||'').toLowerCase().replace(/[^a-z0-9]/g,'');
const num=v=>{const n=parseFloat(String(v).trim());return isNaN(n)?null:n};
const r1=x=>Math.round((x+1e-9)*10)/10;
const cred=s=>String(s||'').split('-').reduce((a,b)=>a+(parseInt(b,10)||0),0);
let C=null,T=null,OUT=null;
async function readSheet(f){const wb=XLSX.read(await f.arrayBuffer(),{type:'array'});
 const rows=XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]],{defval:''});
 return rows.map(r=>{const o={};for(const k in r)o[k.trim()]=r[k];return o})}
function wire(n,need,label,set){
 const inp=$('#f'+n),pill=$('#p'+n),drop=$('#d'+n);
 const load=async f=>{if(!f)return;pill.className='pill on';pill.textContent='Reading '+f.name+'…';
  try{const rows=await readSheet(f);
   if(!rows.length||!(need in rows[0])){pill.className='pill on err';pill.textContent='This does not look like the '+label+': column “'+need+'” not found.';set(null);}
   else{set(rows);pill.className='pill on';pill.textContent=f.name+' · '+rows.length+' rows';}}
  catch(e){pill.className='pill on err';pill.textContent='Could not read this file.';set(null)}
  ready()};
 inp.onchange=e=>load(e.target.files[0]);
 drop.ondragover=e=>{e.preventDefault();drop.classList.add('hot')};drop.ondragleave=()=>drop.classList.remove('hot');
 drop.ondrop=e=>{e.preventDefault();drop.classList.remove('hot');load(e.dataTransfer.files[0])}}
wire(1,'Registration ID','Consolidated Report',v=>C=v);wire(2,'ENROLLMENT_NO','TRU sheet',v=>T=v);
function ready(){$('#run').disabled=!(C&&T);$('#msg').textContent=(C&&T)?'Ready: '+C.length+' consolidated rows · '+T.length+' TRU students.':'Load both sheets to begin.'}

function compare(C,T){
 const cs={};C.forEach(r=>{(cs[key(r['Registration ID'])]=cs[key(r['Registration ID'])]||[]).push(r)});
 const pads=Object.keys(T[0]).map(k=>(k.match(/^SUBJ_CODE__(\d+)$/)||[])[1]).filter(Boolean);
 const U=[],S=[],used=new Set();
 T.forEach(t=>{
  const rid=key(t.ENROLLMENT_NO),rows=cs[rid];
  if(!rows){S.push([rid,t.NAME,'NOT IN CONSOLIDATED'].concat(Array(25).fill('')));return}
  used.add(rid);let cp=0,tcp=0,tc=0,tm=[],miss=0;
  pads.forEach(p=>{
   const code=String(t['SUBJ_CODE__'+p]).trim();if(!code)return;
   const s=rows.find(r=>String(r['Course Code']).trim()===code),cr=cred(t['MAX_CREDS_TH__'+p]);tc+=cr;
   if(!s){miss++;U.push([rid,t.NAME,code,t['SUBJ_NAME__'+p],'','MISSING'].concat(Array(19).fill('')));return}
   const cm=num(s['Total Marks']),tmk=num(t['OBT_CREDS_TH__'+p]),cg=num(s['Grade Points']),tg=num(t['GRD_POINT_TH__'+p]),tcpv=num(t['CRD_POINT_TH__'+p]);
   const sn=String(t['SUBJ_NAME__'+p]).trim(),cn=String(s['Course Name']).trim(),mx=num(t['MAX_MARKS_TH__'+p]);
   cp+=cg*cr;tcp+=tcpv;tm.push(cm);
   U.push([rid,t.NAME,code,sn,cn,sn===cn?'Match':norm(sn)===norm(cn)?'Match (case/spacing differs)':'DIFFERENT WORDING',
    mx,'100 (scheme)',mx===100?'Match':'MISMATCH',cm,tmk,cm===tmk?'Match':Math.round(cm)===tmk?'Match (rounded)':'MISMATCH',
    s.Grade,t['GRD_LETTR_TH__'+p],s.Grade===t['GRD_LETTR_TH__'+p]?'Match':'MISMATCH',cg,tg,cg===tg?'Match':'MISMATCH',
    cr,t['MAX_CREDS_TH__'+p],cg*cr,tcpv,cg*cr===tcpv?'Match':'MISMATCH',s['Is Failed'],s['Is Failed']==='NO'?'Match':'MISMATCH'])});
  const f=rows[0],cn=String(f['Student Name']).trim(),tn=String(t.NAME).trim();
  const avg=tm.reduce((a,b)=>a+b,0)/(tm.length||1),tp=num(t.PERCENTAGE),csg=num(f.SGPA),tsg=num(t.SGPA),calc=tc?cp/tc:0;
  const gt=num(t.GRAND_TTL_MARKS),tcr=num(t.TOTAL_OBT_CREDITS),tcgp=String(t.CGPA).trim();
  const cgc=num(f.CGPA),tcg=num(tcgp);
  S.push([rid,t.NAME,cn,tn===cn?'Match':norm(tn)===norm(cn)?'Match (case/spacing/punctuation differs)':'MISMATCH',
   gt,tm.length*100,gt===tm.length*100?'Match':'MISMATCH',
   tm.reduce((a,b)=>a+b,0),+avg.toFixed(2),tp,r1(avg)===tp?'Match':Math.abs(avg-tp)<.1?'Match (rounded)':'MISMATCH',
   tcr,tc,tcr===tc?'Match':'MISMATCH',tcp,cp,tcp===cp?'Match':'MISMATCH',
   csg,tsg,+calc.toFixed(2),csg===tsg?'Match':r1(csg)===tsg?'Match (rounded)':'MISMATCH',Math.abs(calc-csg)<.005?'Match':'MISMATCH',
   f.CGPA,tcgp,tcgp===''?'Cannot compare (TRU CGPA blank)':(cgc===tcg?'Match':cgc!=null&&r1(cgc)===tcg?'Match (rounded)':'MISMATCH'),
   f['Programme Name'],f['Batch Year'],rows.every(r=>r['Is Failed']==='NO')?'All pass (Consolidated)':'FAIL IN CONSOLIDATED'])});
 const N=[];Object.keys(cs).filter(k=>!used.has(k)).forEach(k=>{const r=cs[k];N.push([k,String(r[0]['Student Name']).trim(),r[0].SGPA,r[0].CGPA,r.filter(x=>x['Is Failed']==='YES').length])});
 const ok=S.filter(s=>s[2]!=='NOT IN CONSOLIDATED'),nOK=ok.length;
 const tally=(label,rows,i,n,note)=>{let e=0,s=0,b=0,x=0;rows.forEach(r=>{const v=String(r[i]);if(v==='Match')e++;else if(v.startsWith('Match ('))s++;else if(v.startsWith('Cannot'))x++;else if(v)b++});
  return [label,n==null?rows.length:n,e,s,b,x?'Not compared: '+x+' record(s) have no value in the TRU file':note||'']};
 const ur=U.filter(r=>r[5]!=='MISSING');
 let ie=0,is=0,ib=0;C.forEach(r=>{const g=num(r['Grace Marks'])||0,d=num(r['Total Marks'])-(num(r['Internal Marks'])+num(r['External Marks'])+g);Math.abs(d)<.01?ie++:Math.abs(d-.5)<.01?is++:ib++});
 const M=[tally('Roll number (key)',[],0,T.length,'')];M[0][2]=nOK;M[0][4]=T.length-nOK;M[0][5]=T.length-nOK?'Roll numbers in TRU but absent from the Consolidated Report':'All TRU roll numbers found in the Consolidated Report';
 M.push(tally('Student name',ok,3),tally('Course code',[],0,U.length));M[2][2]=U.length-miss0(U);M[2][4]=miss0(U);M[2][5]='Subjects found in Consolidated by course code';
 M.push(tally('Course name',U,5,null,'Wording differs between the two files for the course name'),
  tally('Maximum marks (100)',ur,8),tally('Total marks per subject',ur,11,null,'TRU column OBT_CREDS_TH holds marks'),
  tally('Grade letter',ur,14),tally('Grade point',ur,17),tally('Credit points per subject (GP × credits)',ur,22),tally('Pass/Fail per subject',ur,24,null,'TRU result column is blank; Consolidated “Is Failed” is used'),
  tally('Grand total marks',ok,6),tally('Overall percentage',ok,10,null,'TRU % against the average of the subject totals'),
  tally('Total credits',ok,13),tally('Total credit points',ok,16),tally('SGPA',ok,20,null,'Soft = Consolidated 2 dp against TRU 1 dp rounding'),
  tally('SGPA recomputed from grade points × credits',ok,21),tally('CGPA',ok,24),
  ['Consolidated check: Internal + External vs Total marks',C.length,ie,is,ib,'Soft = Total rounded up by 0.5; no real mismatch if the last count is 0']);
 return {M,S,U,N}}
const miss0=U=>U.filter(r=>r[5]==='MISSING').length;
const SH={
 M:['Field compared','Records checked','Exact match','Match after rounding/format','Mismatch / different','Note'],
 S:['Roll No','Name (TRU)','Name (Consolidated)','Name Status','Grand Total Marks (TRU)','Max Marks (derived)','Grand Total Status','Sum of Total Marks (Consol.)','Average % (derived)','Percentage (TRU)','Percentage Status','Total Credits (TRU)','Total Credits (derived)','Credits Status','Total Credit Points (TRU)','Credit Points (Consol. GP × credits)','Credit Points Status','SGPA (Consol.)','SGPA (TRU)','SGPA recomputed','SGPA Status (Consol. vs TRU)','SGPA Recalc Check','CGPA (Consol.)','CGPA (TRU)','CGPA Status','Programme (Consol.)','Batch (Consol.)','Pass/Fail'],
 U:['Roll No','Name','Course Code','Course Name (TRU)','Course Name (Consol.)','Course Name Status','Max Marks (TRU)','Max Marks (scheme)','Max Marks Status','Total Marks (Consol.)','Total Marks (TRU)','Total Marks Status','Grade (Consol.)','Grade (TRU)','Grade Status','Grade Point (Consol.)','Grade Point (TRU)','Grade Point Status','Credits (derived)','Credit Structure (TRU)','Credit Points (Consol.)','Credit Points (TRU)','Credit Points Status','Is Failed (Consol.)','Pass Status'],
 N:['Roll No','Student Name','SGPA','CGPA','Failed Subjects']};
const FM=[['Roll number','Registration ID','ENROLLMENT_NO','Compared (join key)'],['Student name','Student Name','NAME','Compared'],['Course code','Course Code','SUBJ_CODE__xx','Compared'],['Course name','Course Name','SUBJ_NAME__xx','Compared'],['Max marks','Course Schema (100)','MAX_MARKS_TH__xx','Compared'],['Total marks','Total Marks','OBT_CREDS_TH__xx (header is shifted; holds marks)','Compared'],['Grade letter','Grade','GRD_LETTR_TH__xx','Compared'],['Grade point','Grade Points','GRD_POINT_TH__xx','Compared'],['Credits','(not in Consolidated)','MAX_CREDS_TH__xx e.g. 3-0-0','Derived from TRU'],['Credit points','Grade Points × credits (derived)','CRD_POINT_TH__xx','Compared'],['Grand total marks','subjects × 100 (derived)','GRAND_TTL_MARKS','Compared'],['Percentage','Average of Total Marks (derived)','PERCENTAGE','Compared'],['Total credits','Σ credits (derived)','TOTAL_OBT_CREDITS','Compared'],['SGPA','SGPA','SGPA','Compared and recomputed'],['CGPA','CGPA','CGPA','Compared when present in TRU'],['Pass/Fail','Is Failed','RESULT (blank in TRU)','Consolidated only'],['Internal / external / grace / re-exam marks, programme, batch, attendance','Present','Not in TRU','Not comparable'],['Email, hall admit no.','Not in Consolidated','Present','Not comparable']];
const FMH=['Data item','Consolidated Report column','TRU file column','Treatment'];
const TABS=[['Summary','M',SH.M],['Student-wise','S',SH.S],['Subject-wise','U',SH.U],['Field Mapping','F',FMH],['Only in Consolidated','N',SH.N]];
const cls=v=>{v=String(v);return/^(MISMATCH|DIFFERENT|MISSING|NOT IN|FAIL)/.test(v)?'bad':/^(Match \(|Cannot)/.test(v)?'warn':v==='Match'?'ok':''};
function table(h,rows,color){return '<div class="scroll"><table><thead><tr>'+h.map(x=>'<th>'+esc(x)+'</th>').join('')+'</tr></thead><tbody>'+
 rows.map(r=>'<tr>'+r.map(c=>'<td class="'+(color?cls(c):'')+'">'+esc(c)+'</td>').join('')+'</tr>').join('')+'</tbody></table></div>'}
function show(i){document.querySelectorAll('.tab').forEach((t,j)=>t.classList.toggle('on',i===j));
 const [,k,h]=TABS[i];const d=k==='F'?FM:OUT[k];$('#panel').innerHTML=table(h,d,k==='S'||k==='U')}
function render(){const {M,S,U,N}=OUT;const all=[].concat(...[S.map(r=>r.slice(3)),U.map(r=>r.slice(3))]).map(String);
 const bad=all.filter(v=>/^(MISMATCH|DIFFERENT|MISSING)/.test(v)).length,soft=all.filter(v=>v.startsWith('Match (')).length;
 const k=[['ok',S.length,'Students compared'],['',U.length,'Subject rows'],['bad',bad,'Real mismatches'],['warn',soft,'Rounding / format only'],['warn',N.length,'Only in Consolidated'],['',M[M.length-1]&&M.find(r=>r[0]==='CGPA')[5].startsWith('Not')?'n/a':'ok','CGPA check']];
 $('#kpis').innerHTML=k.map(x=>'<div class="kpi '+x[0]+'"><b>'+esc(x[1])+'</b><span>'+esc(x[2])+'</span></div>').join('');
 $('#tabs').innerHTML=TABS.map((t,i)=>'<div class="tab" data-i="'+i+'">'+t[0]+'</div>').join('');
 document.querySelectorAll('.tab').forEach(t=>t.onclick=()=>show(+t.dataset.i));
 $('#res').style.display='block';show(0)}
$('#run').onclick=()=>{try{OUT=compare(C,T);render();$('#dl').disabled=false;build().then(b=>{if(OUT)OUT.buf=b}).catch(()=>{});$('#msg').textContent='Done. '+OUT.S.length+' students checked.'}catch(e){$('#msg').textContent='Could not compare these sheets: '+e.message}};
$('#clr').onclick=()=>{C=T=OUT=null;['1','2'].forEach(n=>{$('#f'+n).value='';$('#p'+n).className='pill'});$('#res').style.display='none';$('#dl').disabled=true;ready()};
const FILL={ok:'FFC6EFCE',bad:'FFFFC7CE',warn:'FFFFEB9C'};
async function build(){const wb=new ExcelJS.Workbook();
 TABS.forEach(([name,k,h])=>{const ws=wb.addWorksheet(name==='Summary'?'Summary':name==='Student-wise'?'Student-wise Comparison':name==='Subject-wise'?'Subject-wise Comparison':name);
  const d=k==='F'?FM:OUT[k];ws.addRow(h);
  ws.getRow(1).eachCell(c=>{c.font={bold:true,color:{argb:'FFFFFFFF'}};c.fill={type:'pattern',pattern:'solid',fgColor:{argb:'FF305496'}};c.alignment={wrapText:true,vertical:'top'}});
  d.forEach(r=>{const row=ws.addRow(r);if(k==='S'||k==='U')row.eachCell(c=>{const t=cls(c.value);if(t)c.fill={type:'pattern',pattern:'solid',fgColor:{argb:FILL[t]}}})});
  if(k==='M')ws.eachRow((row,i)=>{if(i>1&&Number(row.getCell(5).value)>0)row.getCell(5).fill={type:'pattern',pattern:'solid',fgColor:{argb:FILL.bad}}});
  ws.views=[{state:'frozen',ySplit:1,xSplit:k==='S'||k==='U'?2:0}];
  h.forEach((_,i)=>{ws.getColumn(i+1).width=Math.min(45,Math.max(10,...[h[i]].concat(d.map(r=>r[i])).map(v=>String(v==null?'':v).length+2)))})});
 return wb.xlsx.writeBuffer()}
const XMIME='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',XNAME='Full_Cross_Verification_Report.xlsx';
function saveBlob(b,n){const a=document.createElement('a'),u=URL.createObjectURL(b);a.href=u;a.download=n;a.style.display='none';document.body.appendChild(a);a.click();setTimeout(()=>{URL.revokeObjectURL(u);a.remove()},2000)}
$('#dl').onclick=async()=>{const m=$('#msg');
 try{const dl=typeof claude!=='undefined'?await claude.use('downloads'):null;if(!dl){if(!OUT.buf)OUT.buf=await build();saveBlob(new Blob([OUT.buf],{type:XMIME}),XNAME);m.textContent='Report downloaded. Check your Downloads folder.';return}
  if(!OUT.buf)OUT.buf=await build();await dl.save({filename:'Full_Cross_Verification_Report.xlsx',data:new Blob([OUT.buf],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'})});m.textContent='Report saved.'}
 catch(e){m.textContent=e&&e.code==='declined'?'Save cancelled.':'Could not save the report ('+(e&&(e.code||e.message)||'error')+').'}};
</script>
</body>
</html>
