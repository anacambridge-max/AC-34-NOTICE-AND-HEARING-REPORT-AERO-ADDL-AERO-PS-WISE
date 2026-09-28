"use client";

import { useEffect, useMemo, useState } from "react";
import * as XLSX from "xlsx";

type Row = {
  ps:number; aero:string; designation:string; aeroMobile:string; bloName:string; bloMobile:string; supervisor:string;
  noticeGenerated:number; prevDelivered:number; latestDelivered:number; prevHearing:number; latestHearing:number;
  hearingLapse:number; discrepancyDelivered:number; bloDocs:number; bloLetter:number;
};

const aliases = {
  ps:["ps","ps no","ps number","p.s.","p.s. no","p.s. number","part no","part number","partno","polling station","part","part no of polling station"],
  generated:["notice generated","notices generated","notice gen","no of notices generated","total notices generated"],
  delivered:["notice delivered","notices delivered","delivered","notices delivered to electors","notice delivery"],
  hearing:["hearings held","hearing held","hearing done","hearing","hearing held by blo","no of hearings held"],
  lapse:["hearing date lapsed","hearing lapsed","lapsed","hearing lapse","date lapsed"],
  docs:["documents uploaded by blo","documents uploaded by blo (no mapping)","documents uploaded","docs uploaded by blo","blo docs uploaded","blo docs uploaded no mapping","blo documents","blo docs","documents uploaded by blo no mapping","no of documents uploaded by blo","document uploaded by blo"],
  discrepancy:["discrep notices delivered","discrepancy notices delivered","discrep notices delivery","anomaly notices delivered","anomaly delivered","discrepancy delivered","discrepancy/anomaly delivered","discrepancy + anomaly delivered","no of discrepancy notices delivered","no of anomaly notices delivered","discrepancy notice delivered"],
  letter:["blo letter uploaded","blo letters uploaded","blo letter","blo letter uploaded by blo","letters uploaded by blo","no of blo letter uploaded","blo letters"]
};

function norm(v:any){return String(v??"").trim().toLowerCase().replace(/[^a-z0-9]+/g," ");}
function num(v:any){const n=Number(String(v??"").replace(/,/g,"").replace(/%/g,""));return Number.isFinite(n)?n:0;}
function findCol(headers:string[], names:string[]){
  const hs=headers.map(norm);
  for(const n of names){const target=norm(n);const i=hs.indexOf(target);if(i>=0)return i;}
  for(const n of names){
    const target=norm(n);
    if(target.length<4) continue;
    const i=hs.findIndex(h=>h===target || h.includes(target) || target.includes(h));
    if(i>=0)return i;
  }
  return -1;
}

async function readRows(file:File){
  const buf=await file.arrayBuffer();
  const wb=XLSX.read(buf,{type:"array",cellDates:true});
  const ws=wb.Sheets[wb.SheetNames[0]];
  const raw=XLSX.utils.sheet_to_json<any[]>(ws,{header:1,defval:""});
  let headerIndex=0;
  for(let i=0;i<Math.min(raw.length,40);i++){
    const h=(raw[i]||[]).map(norm).join(" | ");
    if(h.includes("part no")||h.includes("p s no")||h.includes("ps no")||h.includes("ps number")||h.includes("notice generated")||h.includes("hearings held")||h.includes("discrepancy")||h.includes("blo docs")||h.includes("blo letter")){headerIndex=i;break}
  }
  const headers=(raw[headerIndex]||[]).map((x:any)=>String(x??""));
  return {headers, rows:raw.slice(headerIndex+1)};
}

function parseMetricFile(parsed:{headers:string[];rows:any[][]}, kind:"eci"|"blo"){
  const {headers,rows}=parsed;
  const psI=findCol(headers,aliases.ps);
  const genI=findCol(headers,aliases.generated);
  const delI=findCol(headers,aliases.delivered);
  const hearI=findCol(headers,aliases.hearing);
  const lapseI=findCol(headers,aliases.lapse);
  const docsI=findCol(headers,aliases.docs);
  const discI=findCol(headers,aliases.discrepancy);
  const letterI=findCol(headers,aliases.letter);
  const out=new Map<number,any>();
  for(const r of rows){
    if(psI<0) continue;
    const ps=Math.round(num(r[psI])); if(!ps) continue;
    const item:any={};
    if(genI>=0)item.noticeGenerated=num(r[genI]);
    if(delI>=0)item.latestDelivered=num(r[delI]);
    if(hearI>=0)item.latestHearing=num(r[hearI]);
    if(lapseI>=0)item.hearingLapse=num(r[lapseI]);
    if(docsI>=0)item.bloDocs=num(r[docsI]);
    if(discI>=0)item.discrepancyDelivered=num(r[discI]);
    if(letterI>=0)item.bloLetter=num(r[letterI]);
    out.set(ps,item);
  }
  return out;
}

function parseMapping(parsed:{headers:string[];rows:any[][]}){
  const {headers,rows}=parsed;
  const psI=findCol(headers,aliases.ps);
  const get=(names:string[])=>findCol(headers,names);
  const aeroI=get(["aero","ad aero","aero / ad aero","officer"]);
  const desI=get(["designation","type"]);
  const aeroMobI=get(["aero mobile","aero mobile no","officer mobile"]);
  const bloI=get(["blo name","blo"]);
  const bloMobI=get(["blo mobile","blo mobile no","mobile no"]);
  const supI=get(["blo supervisor","supervisor"]);
  const out=new Map<number,any>();
  for(const r of rows){const ps=psI>=0?Math.round(num(r[psI])):0;if(!ps)continue;out.set(ps,{aero:aeroI>=0?String(r[aeroI]||""):"",designation:desI>=0?String(r[desI]||""):"",aeroMobile:aeroMobI>=0?String(r[aeroMobI]||""):"",bloName:bloI>=0?String(r[bloI]||""):"",bloMobile:bloMobI>=0?String(r[bloMobI]||""):"",supervisor:supI>=0?String(r[supI]||""):""});} 
  return out;
}

function merge(master:Map<number,any>, previous:Map<number,any>, latest:Map<number,any>, blo:Map<number,any>){
  return Array.from(master.entries()).sort((a,b)=>a[0]-b[0]).map(([ps,m])=>{
    const p=previous.get(ps)||{}, l=latest.get(ps)||{}, b=blo.get(ps)||{};
    const latestDelivered=l.latestDelivered??m.latestDelivered??0;
    const prevDelivered=p.latestDelivered??m.prevDelivered??0;
    const latestHearing=l.latestHearing??m.latestHearing??0;
    const prevHearing=p.latestHearing??m.prevHearing??0;
    return {...m,ps,noticeGenerated:l.noticeGenerated??m.noticeGenerated??0,prevDelivered,latestDelivered,prevHearing,latestHearing,hearingLapse:l.hearingLapse??m.hearingLapse??0,discrepancyDelivered:b.discrepancyDelivered??l.discrepancyDelivered??m.discrepancyDelivered??0,bloDocs:b.bloDocs??l.bloDocs??m.bloDocs??0,bloLetter:b.bloLetter??l.bloLetter??m.bloLetter??0};
  });
}

function printReport(allRows:Row[]) {
  const groups=new Map<string,Row[]>();
  for(const r of allRows){const k=r.aero||"Unmapped";const a=groups.get(k)||[];a.push(r);groups.set(k,a);}
  const sum=(rs:Row[],k:keyof Row)=>rs.reduce((s,r)=>s+Number(r[k]||0),0);
  const n=(x:number)=>x.toLocaleString("en-IN");
  const pct=(a:number,b:number)=>b?((a/b)*100).toFixed(1)+"%":"0.0%";
  const esc=(x:any)=>String(x??"").replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;");
  const groupRows=Array.from(groups.entries()).map(([aero,rs])=>{
    const x={ps:rs.length,gen:sum(rs,"noticeGenerated"),pd:sum(rs,"prevDelivered"),ld:sum(rs,"latestDelivered"),ph:sum(rs,"prevHearing"),lh:sum(rs,"latestHearing"),lapse:sum(rs,"hearingLapse"),disc:sum(rs,"discrepancyDelivered"),docs:sum(rs,"bloDocs"),letter:sum(rs,"bloLetter")};
    return {aero,rs,x,dispose:x.lh+x.lapse?x.lh/(x.lh+x.lapse)*100:0};
  }).sort((a,b)=>a.dispose-b.dispose);
  const under=new Set(groupRows.slice(0,3).map(x=>x.aero));
  const bloMap=new Map<string,{aero:string,blo:string,ps:number,hearing:number,lapse:number}>();
  for(const r of allRows){const key=`${r.aero}|${r.bloName}|${r.ps}`;const x=bloMap.get(key)||{aero:r.aero,blo:r.bloName,ps:r.ps,hearing:0,lapse:0};x.hearing+=r.latestHearing;x.lapse+=r.hearingLapse;bloMap.set(key,x);}
  const underBlo=new Set(Array.from(bloMap.values()).sort((a,b)=>(a.hearing+a.lapse?a.hearing/(a.hearing+a.lapse)*100:0)-(b.hearing+b.lapse?b.hearing/(b.hearing+b.lapse)*100:0)).slice(0,3).map(x=>`${x.aero}|${x.blo}|${x.ps}`));
  let h="<!doctype html><html><head><title>AC-34 Notice & Hearing Report</title><style>@page{size:A4 landscape;margin:7mm}body{font-family:Arial,sans-serif;font-size:10px;color:#111}h1{text-align:center;font-size:18px;margin:0 0 5px}h2{font-size:14px;margin:13px 0 6px}.meta{text-align:center;font-size:10px;margin-bottom:9px}table{border-collapse:collapse;width:100%;margin-bottom:12px;page-break-inside:auto}tr{page-break-inside:avoid}th,td{border:1px solid #222;padding:5px 4px;vertical-align:middle}th{background:#dfe6ee;font-weight:700;font-size:9px}td.num{text-align:right;font-size:10px;font-weight:600}.zero{background:#ffd6d6!important;color:#9b0000!important;font-weight:700}.under{background:#b9d9ff!important;color:#063b73!important;font-weight:700}.grand{background:#d9d9d9;font-weight:700}.aero{background:#dfe6ee;font-weight:700;padding:7px;border:1px solid #222;font-size:11px}.badge{background:#b9d9ff;border:1px solid #5d91c9;padding:2px 5px;margin-left:5px}.small{font-size:9px}</style></head><body>";
  h+="<h1>OFFICE OF THE ELECTORAL REGISTRATION OFFICER, AC-34, MATIALA</h1><div class=meta>SIR-2026 — PS-WISE (AERO-WISE) NOTICE & HEARING REPORT — Previous ECI vs Latest ECI</div>";
  h+="<h2>AERO / Ad.AERO WISE CONSOLIDATED REPORT</h2><table><tr><th>S.No.</th><th>AERO / Ad.AERO</th><th>Designation</th><th>PS</th><th>Notice Generated</th><th>Delivered Prev</th><th>Delivered Latest</th><th>Diff Delivered</th><th>Hearing Prev</th><th>Hearing Latest</th><th>Diff Hearing</th><th>Lapse</th><th>Dispose %</th><th>Discrepancy Delivered</th><th>BLO Docs</th><th>BLO Letter</th></tr>";
  let gi=0; const grand={ps:0,gen:0,pd:0,ld:0,ph:0,lh:0,lapse:0,disc:0,docs:0,letter:0};
  for(const g of groupRows){
    const {aero,rs,x}=g;
    grand.ps+=x.ps;grand.gen+=x.gen;grand.pd+=x.pd;grand.ld+=x.ld;grand.ph+=x.ph;grand.lh+=x.lh;grand.lapse+=x.lapse;grand.disc+=x.disc;grand.docs+=x.docs;grand.letter+=x.letter;
    gi++;
    h+="<tr class="+(under.has(aero)?"under":"")+"><td>"+gi+"</td><td>"+esc(aero)+(under.has(aero)?" <span class=badge>TOP 3 UNDERPERFORMER</span>":"")+"</td><td>"+esc(rs[0]?.designation)+"</td><td class=num>"+n(x.ps)+"</td><td class=num>"+n(x.gen)+"</td><td class=num>"+n(x.pd)+"</td><td class=num>"+n(x.ld)+"</td><td class=num>"+n(x.ld-x.pd)+"</td><td class=num>"+n(x.ph)+"</td><td class=num>"+n(x.lh)+"</td><td class=num>"+n(x.lh-x.ph)+"</td><td class=num>"+n(x.lapse)+"</td><td class=num>"+pct(x.lh,x.lh+x.lapse)+"</td><td class=num>"+n(x.disc)+"</td><td class=num>"+n(x.docs)+"</td><td class=num>"+n(x.letter)+"</td></tr>";
  }
  h+="<tr class=grand><td colspan=3>GRAND TOTAL</td><td class=num>"+n(grand.ps)+"</td><td class=num>"+n(grand.gen)+"</td><td class=num>"+n(grand.pd)+"</td><td class=num>"+n(grand.ld)+"</td><td class=num>"+n(grand.ld-grand.pd)+"</td><td class=num>"+n(grand.ph)+"</td><td class=num>"+n(grand.lh)+"</td><td class=num>"+n(grand.lh-grand.ph)+"</td><td class=num>"+n(grand.lapse)+"</td><td class=num>"+pct(grand.lh,grand.lh+grand.lapse)+"</td><td class=num>"+n(grand.disc)+"</td><td class=num>"+n(grand.docs)+"</td><td class=num>"+n(grand.letter)+"</td></tr></table>";
  h+="<div class=small>Blue rows = 3 AERO / Ad.AERO with the lowest disposal percentage. Red PS rows = Latest Hearing Held = 0.</div>";
  h+="<h2>PS-WISE AERO-WISE REPORT</h2>";
  for(const g of groupRows){
    const {aero,rs,x}=g;
    h+="<div class=aero>"+esc(aero)+" — "+esc(rs[0]?.designation)+" | "+rs.length+" PS | Notice Generated: "+n(x.gen)+" | Hearings Held: "+n(x.lh)+(under.has(aero)?" <span class=badge>TOP 3 UNDERPERFORMER</span>":"")+"</div>";
    h+="<table><tr><th>S.No.</th><th>PS No.</th><th>BLO Name</th><th>BLO Mobile</th><th>BLO Supervisor</th><th>Notice Generated</th><th>Delivered Prev</th><th>Delivered Latest</th><th>Diff Delivered</th><th>Hearing Prev</th><th>Hearing Latest</th><th>Diff Hearing</th><th>Lapse</th><th>Discrepancy Delivered</th><th>BLO Docs</th><th>BLO Letter</th></tr>";
    rs.slice().sort((a,b)=>a.ps-b.ps).forEach((r,i)=>{
      h+="<tr class="+(r.latestHearing===0?"zero":(underBlo.has(`${r.aero}|${r.bloName}|${r.ps}`)?"under":""))+"><td>"+(i+1)+"</td><td>"+r.ps+"</td><td>"+esc(r.bloName)+"</td><td>"+esc(r.bloMobile)+"</td><td>"+esc(r.supervisor)+"</td><td class=num>"+n(r.noticeGenerated)+"</td><td class=num>"+n(r.prevDelivered)+"</td><td class=num>"+n(r.latestDelivered)+"</td><td class=num>"+n(r.latestDelivered-r.prevDelivered)+"</td><td class=num>"+n(r.prevHearing)+"</td><td class=num>"+n(r.latestHearing)+(r.latestHearing===0?" <span class=badge>ZERO HEARING HELD</span>":"")+"</td><td class=num>"+n(r.latestHearing-r.prevHearing)+"</td><td class=num>"+n(r.hearingLapse)+"</td><td class=num>"+n(r.discrepancyDelivered)+"</td><td class=num>"+n(r.bloDocs)+"</td><td class=num>"+n(r.bloLetter)+"</td></tr>";
    });
    h+="</table>";
  }
  h+="</body></html>";
  const w=window.open("","_blank","width=1600,height=1100");
  if(!w){alert("Please allow pop-ups for PDF export.");return;}
  w.document.write(h);w.document.close();setTimeout(()=>w.print(),700);
}
export default function Page(){
  const [master,setMaster]=useState<Map<number,any>>(new Map());
  const [rows,setRows]=useState<Row[]>([]);
  const [previous,setPrevious]=useState<Map<number,any>>(new Map());
  const [latest,setLatest]=useState<Map<number,any>>(new Map());
  const [blo,setBlo]=useState<Map<number,any>>(new Map());
  const [filterAero,setFilterAero]=useState("ALL");
  const [filterSup,setFilterSup]=useState("ALL");
  const [search,setSearch]=useState("");
  const [error,setError]=useState("");

  useEffect(()=>{try{const saved=localStorage.getItem("ac34-master-mapping");if(saved){const arr=JSON.parse(saved);setMaster(new Map(arr));}}catch{}},[]);
  useEffect(()=>{if(master.size)localStorage.setItem("ac34-master-mapping",JSON.stringify(Array.from(master.entries())));},[master]);

  const load=(setter:any,kind:"eci"|"blo")=>(e:React.ChangeEvent<HTMLInputElement>)=>{const f=e.target.files?.[0];if(!f)return;setError("");readRows(f).then(p=>{const parsed=kind==="eci"?parseMetricFile(p,"eci"):parseMetricFile(p,"blo");setter(parsed);}).catch(err=>setError(String(err)));};

  useEffect(()=>{if(master.size) setRows(merge(master,previous,latest,blo));},[master,previous,latest,blo]);

  const aeros=useMemo(()=>["ALL",...Array.from(new Set(rows.map(r=>r.aero).filter(Boolean)))],[rows]);
  const sups=useMemo(()=>["ALL",...Array.from(new Set(rows.map(r=>r.supervisor).filter(Boolean)))],[rows]);
  const shown=useMemo(()=>rows.filter(r=>(filterAero==="ALL"||r.aero===filterAero)&&(filterSup==="ALL"||r.supervisor===filterSup)&&((r.ps+" "+r.bloName+" "+r.supervisor+" "+r.aero).toLowerCase().includes(search.toLowerCase()))),[rows,filterAero,filterSup,search]);
  const totals=useMemo(()=>rows.reduce((a,r)=>({gen:a.gen+r.noticeGenerated,del:a.del+r.latestDelivered,ph:a.ph+r.prevHearing,lh:a.lh+r.latestHearing,zero:a.zero+(r.latestHearing===0?1:0)}),{gen:0,del:0,ph:0,lh:0,zero:0}),[rows]);
  const officerSummary=useMemo(()=>{const m=new Map<string,any>();for(const r of rows){const k=r.aero||"Unmapped";const x=m.get(k)||{aero:k,designation:r.designation,ps:0,gen:0,del:0,dh:0,zero:0};x.ps++;x.gen+=r.noticeGenerated;x.del+=r.latestDelivered;x.dh+=r.latestHearing;x.zero+=r.latestHearing===0?1:0;m.set(k,x);}return Array.from(m.values()).sort((a,b)=>a.aero.localeCompare(b.aero));},[rows]);
  const underperformers=useMemo(()=>{const m=officerSummary.map(x=>{const rs=rows.filter(r=>r.aero===x.aero);const hearing=rs.reduce((s,r)=>s+r.latestHearing,0);const lapse=rs.reduce((s,r)=>s+r.hearingLapse,0);return {aero:x.aero,dispose:(hearing+lapse)?hearing/(hearing+lapse)*100:0};}).sort((a,b)=>a.dispose-b.dispose);return new Set(m.slice(0,3).map(x=>x.aero));},[officerSummary,rows]);
  const bloUnderperformers=useMemo(()=>{const m=new Map<string,any>();for(const r of rows){const key=`${r.aero}|${r.bloName}|${r.ps}`;const x=m.get(key)||{aero:r.aero,blo:r.bloName,ps:r.ps,hearing:0,lapse:0};x.hearing+=r.latestHearing;x.lapse+=r.hearingLapse;m.set(key,x);}return new Set(Array.from(m.values()).sort((a,b)=>((a.hearing+a.lapse)?a.hearing/(a.hearing+a.lapse)*100:0)-((b.hearing+b.lapse)?b.hearing/(b.hearing+b.lapse)*100:0)).slice(0,3).map(x=>`${x.aero}|${x.blo}|${x.ps}`));},[rows]);
  const supervisorSummary=useMemo(()=>{const m=new Map<string,any>();for(const r of rows){const k=r.supervisor||"Unmapped";const x=m.get(k)||{name:k,ps:0,hearing:0,zero:0};x.ps++;x.hearing+=r.latestHearing;x.zero+=r.latestHearing===0?1:0;m.set(k,x);}return Array.from(m.values()).sort((a,b)=>a.name.localeCompare(b.name));},[rows]);

  const exportCsv=()=>{const data=shown.map(r=>({PS:r.ps,AERO:r.aero,Designation:r.designation,"BLO Supervisor":r.supervisor,"BLO Name":r.bloName,"BLO Mobile":r.bloMobile,"Notice Generated":r.noticeGenerated,"Delivered Previous":r.prevDelivered,"Delivered Latest":r.latestDelivered,"Delivered Difference":r.latestDelivered-r.prevDelivered,"Hearing Previous":r.prevHearing,"Hearing Latest":r.latestHearing,"Hearing Difference":r.latestHearing-r.prevHearing,"Hearing Lapse":r.hearingLapse,"Discrepancy Delivered":r.discrepancyDelivered,"BLO Docs":r.bloDocs,"BLO Letter":r.bloLetter}));const ws=XLSX.utils.json_to_sheet(data);const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,ws,"Dashboard");XLSX.writeFile(wb,"AC34_Dashboard_Export.xlsx");};

  return <main className="wrap">
    <div className="top"><div><div className="title">AC-34 MATIALA — NOTICE & HEARING DASHBOARD</div><div className="sub">SIR-2026 • Fixed PS → AERO/Ad.AERO → BLO Supervisor → BLO mapping • Latest ECI comparison</div></div><div className="badge">430 PS MASTER STRUCTURE</div></div>
    <div className="grid">
      <div className="card"><h3>① Previous ECI Excel</h3><p>Baseline report used for comparison.</p><input className="input" type="file" accept=".xlsx,.xls,.csv" onChange={load(setPrevious,"eci")}/></div>
      <div className="card"><h3>② Latest ECI Excel</h3><p>Current report. Notice Generated + latest hearing status come from this file.</p><input className="input" type="file" accept=".xlsx,.xls,.csv" onChange={load(setLatest,"eci")}/></div>
      <div className="card"><h3>③ BLO / Other Excel</h3><p>Documents uploaded, discrepancy/anomaly delivery and BLO-letter data are merged by PS.</p><input className="input" type="file" accept=".xlsx,.xls,.csv" onChange={load(setBlo,"blo")}/></div>
    </div>
    {error&&<div className="error">{error}</div>}
    <div className="section card"><h3>One-time Master Mapping</h3><div className="note">The app has a fixed master mapping slot. If you need to replace it, upload an Excel containing PS No, AERO/Ad.AERO, BLO Supervisor, BLO Name and mobiles. The mapping is used on every subsequent ECI upload.</div><div style={{marginTop:10}}><input className="input" type="file" accept=".xlsx,.xls,.csv" onChange={e=>{const f=e.target.files?.[0];if(!f)return;readRows(f).then(p=>{const m=parseMapping(p);if(m.size<100){setError("Master mapping file was not detected. The PS column must be named PS No / PS Number / Part No, and the file must contain AERO/Ad.AERO, BLO Supervisor and BLO Name.");return;}setMaster(m);setError("");}).catch(err=>setError(String(err)));}}/></div></div>
    {master.size===0&&<div className="error">Master mapping is not loaded yet. Upload the one-time mapping Excel below. It must contain PS No, AERO/Ad.AERO, BLO Supervisor and BLO Name. After upload it is saved in this browser for future ECI updates.</div>}
    <div className="stats">
      <div className="stat"><span>PS shown</span><b>{shown.length}</b></div><div className="stat"><span>Notice Generated</span><b>{totals.gen.toLocaleString()}</b></div><div className="stat"><span>Latest Delivered</span><b>{totals.del.toLocaleString()}</b></div><div className="stat"><span>Previous Hearing</span><b>{totals.ph.toLocaleString()}</b></div><div className="stat"><span>Latest Hearing</span><b>{totals.lh.toLocaleString()}</b></div><div className="stat"><span>Zero Hearing PS</span><b className="neg">{totals.zero}</b></div>
    </div>
    <div className="toolbar">
      <select className="select" value={filterAero} onChange={e=>setFilterAero(e.target.value)}>{aeros.map(x=><option key={x}>{x}</option>)}</select>
      <select className="select" value={filterSup} onChange={e=>setFilterSup(e.target.value)}>{sups.map(x=><option key={x}>{x}</option>)}</select>
      <input className="search" placeholder="Search PS / BLO / Supervisor / AERO" value={search} onChange={e=>setSearch(e.target.value)}/>
      <button className="btn" onClick={exportCsv}>Export Excel</button>
      <button className="btn primary" onClick={()=>printReport(rows)}>Download PDF</button>
    </div>
    <div className="section card"><h2>AERO / Ad.AERO Wise Consolidated Report</h2><div className="tableWrap"><table className="table"><thead><tr><th>S.No.</th><th>AERO / Ad.AERO</th><th>Designation</th><th>PS</th><th>Notice Generated</th><th>Delivered Prev</th><th>Delivered Latest</th><th>Δ Delivered</th><th>Hearing Prev</th><th>Hearing Latest</th><th>Δ Hearing</th><th>Lapse</th><th>Discrepancy</th><th>BLO Docs</th><th>BLO Letter</th></tr></thead><tbody>{officerSummary.map((x,i)=>{const rs=rows.filter(r=>r.aero===x.aero);return <tr key={x.aero} className={underperformers.has(x.aero)?"under":""}><td>{i+1}</td><td>{x.aero}{underperformers.has(x.aero)&&<span className="underBadge"> TOP 3 UNDERPERFORMER</span>}</td><td>{x.designation}</td><td className="num">{x.ps}</td><td className="num">{x.gen.toLocaleString()}</td><td className="num">{rs.reduce((s,r)=>s+r.prevDelivered,0).toLocaleString()}</td><td className="num">{x.del.toLocaleString()}</td><td className="num">{rs.reduce((s,r)=>s+r.latestDelivered-r.prevDelivered,0).toLocaleString()}</td><td className="num">{rs.reduce((s,r)=>s+r.prevHearing,0).toLocaleString()}</td><td className="num">{x.dh.toLocaleString()}</td><td className="num">{rs.reduce((s,r)=>s+r.latestHearing-r.prevHearing,0).toLocaleString()}</td><td className="num">{rs.reduce((s,r)=>s+r.hearingLapse,0).toLocaleString()}</td><td className="num">{rs.reduce((s,r)=>s+r.discrepancyDelivered,0).toLocaleString()}</td><td className="num">{rs.reduce((s,r)=>s+r.bloDocs,0).toLocaleString()}</td><td className="num">{rs.reduce((s,r)=>s+r.bloLetter,0).toLocaleString()}</td></tr>})}</tbody></table></div></div>
    <div className="note" style={{marginBottom:8}}>🔴 Red rows = <b>Latest ECI Hearings Held = 0</b>. They remain highlighted on every refresh until the latest report shows at least one hearing held.</div>
    <div className="tableWrap"><table className="table"><thead><tr>
      {["PS","AERO / Ad.AERO","BLO Supervisor","BLO Name","BLO Mobile","Notice Generated","Delivered Prev","Delivered Latest","Δ Delivered","Hearing Prev","Hearing Latest","Δ Hearing","Hearing Lapse","Discrepancy Delivered","BLO Docs","BLO Letter"].map(h=><th key={h}>{h}</th>)}
    </tr></thead><tbody>{shown.map(r=><tr key={r.ps} className={r.latestHearing===0?"zero":(bloUnderperformers.has(`${r.aero}|${r.bloName}|${r.ps}`)?"under":"")}>
      <td>{r.ps}</td><td>{r.aero}</td><td>{r.supervisor}</td><td>{r.bloName}{bloUnderperformers.has(`${r.aero}|${r.bloName}|${r.ps}`)&&<span className="underBadge"> TOP 3 UNDERPERFORMER</span>}</td><td>{r.bloMobile}</td><td className="num">{r.noticeGenerated.toLocaleString()}</td><td className="num">{r.prevDelivered.toLocaleString()}</td><td className="num">{r.latestDelivered.toLocaleString()}</td><td className="num">{(r.latestDelivered-r.prevDelivered).toLocaleString()}</td><td className="num">{r.prevHearing.toLocaleString()}</td><td className="num">{r.latestHearing.toLocaleString()}{r.latestHearing===0&&<span className="zeroBadge"> ZERO HEARING HELD</span>}</td><td className="num">{(r.latestHearing-r.prevHearing).toLocaleString()}</td><td className="num">{r.hearingLapse.toLocaleString()}</td><td className="num">{r.discrepancyDelivered.toLocaleString()}</td><td className="num">{r.bloDocs.toLocaleString()}</td><td className="num">{r.bloLetter.toLocaleString()}</td>
    </tr>)}</tbody></table></div>
    <div className="section card"><h2>How the update works</h2><div className="note">Previous ECI is the comparison baseline. Latest ECI replaces current metrics. Difference columns are Latest − Previous. Notice Generated is taken from Latest. BLO/Other data is joined by PS. The master mapping is not recalculated from the ECI report, so officer/BLO relationships stay fixed while daily numbers change.</div></div>
  </main>;
}
