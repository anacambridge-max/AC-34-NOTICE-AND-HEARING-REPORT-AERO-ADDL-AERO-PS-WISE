"use client";

import { useEffect, useMemo, useState } from "react";
import * as XLSX from "xlsx";

type Row = {
  ps:number; aero:string; designation:string; aeroMobile:string; bloName:string; bloMobile:string; supervisor:string;
  noticeGenerated:number; prevDelivered:number; latestDelivered:number; prevHearing:number; latestHearing:number;
  hearingLapse:number; discrepancyDelivered:number; bloDocs:number; bloLetter:number;
};

const aliases = {
  ps:["ps","p.s.","part no","part number","partno","polling station","part"],
  generated:["notice generated","notices generated","notice gen"],
  delivered:["notice delivered","notices delivered","delivered"],
  hearing:["hearings held","hearing held","hearing done","hearing"],
  lapse:["hearing date lapsed","hearing lapsed","lapsed"],
  docs:["documents uploaded by blo","docs uploaded by blo","blo docs uploaded","blo documents","documents uploaded"],
  discrepancy:["discrep notices delivered","discrepancy notices delivered","anomaly notices delivered","discrepancy delivered"],
  letter:["blo letter uploaded","blo letters uploaded","blo letter"]
};

function norm(v:any){return String(v??"").trim().toLowerCase().replace(/[^a-z0-9]+/g," ");}
function num(v:any){const n=Number(String(v??"").replace(/,/g,"").replace(/%/g,""));return Number.isFinite(n)?n:0;}
function findCol(headers:string[], names:string[]){const hs=headers.map(norm); for(const n of names){const i=hs.indexOf(norm(n));if(i>=0)return i} return -1;}

async function readRows(file:File){
  const buf=await file.arrayBuffer();
  const wb=XLSX.read(buf,{type:"array",cellDates:true});
  const ws=wb.Sheets[wb.SheetNames[0]];
  const raw=XLSX.utils.sheet_to_json<any[]>(ws,{header:1,defval:""});
  let headerIndex=0;
  for(let i=0;i<Math.min(raw.length,40);i++){
    const h=(raw[i]||[]).map(norm).join(" | ");
    if(h.includes("part no")||h.includes("p s no")||h.includes("notice generated")||h.includes("hearings held")){headerIndex=i;break}
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
    <div className="section card"><h3>One-time Master Mapping</h3><div className="note">The app has a fixed master mapping slot. If you need to replace it, upload an Excel containing PS No, AERO/Ad.AERO, BLO Supervisor, BLO Name and mobiles. The mapping is used on every subsequent ECI upload.</div><div style={{marginTop:10}}><input className="input" type="file" accept=".xlsx,.xls,.csv" onChange={e=>{const f=e.target.files?.[0];if(!f)return;readRows(f).then(p=>{const m=parseMapping(p);if(m.size<100){setError("Master mapping file was not detected. Required: PS No + AERO/Ad.AERO + BLO Supervisor + BLO Name.");return;}setMaster(m);setError("");}).catch(err=>setError(String(err)));}}/></div></div>
    {master.size===0&&<div className="error">Master mapping is not loaded yet. Upload the one-time mapping Excel below. It must contain PS No, AERO/Ad.AERO, BLO Supervisor and BLO Name. After upload it is saved in this browser for future ECI updates.</div>}
    <div className="stats">
      <div className="stat"><span>PS shown</span><b>{shown.length}</b></div><div className="stat"><span>Notice Generated</span><b>{totals.gen.toLocaleString()}</b></div><div className="stat"><span>Latest Delivered</span><b>{totals.del.toLocaleString()}</b></div><div className="stat"><span>Previous Hearing</span><b>{totals.ph.toLocaleString()}</b></div><div className="stat"><span>Latest Hearing</span><b>{totals.lh.toLocaleString()}</b></div><div className="stat"><span>Zero Hearing PS</span><b className="neg">{totals.zero}</b></div>
    </div>
    <div className="toolbar">
      <select className="select" value={filterAero} onChange={e=>setFilterAero(e.target.value)}>{aeros.map(x=><option key={x}>{x}</option>)}</select>
      <select className="select" value={filterSup} onChange={e=>setFilterSup(e.target.value)}>{sups.map(x=><option key={x}>{x}</option>)}</select>
      <input className="search" placeholder="Search PS / BLO / Supervisor / AERO" value={search} onChange={e=>setSearch(e.target.value)}/>
      <button className="btn" onClick={exportCsv}>Export Excel</button>
    </div>
    <div className="note" style={{marginBottom:8}}>🔴 Red rows = <b>Latest ECI Hearings Held = 0</b>. They remain highlighted on every refresh until the latest report shows at least one hearing held.</div>
    <div className="tableWrap"><table className="table"><thead><tr>
      {["PS","AERO / Ad.AERO","BLO Supervisor","BLO Name","BLO Mobile","Notice Generated","Delivered Prev","Delivered Latest","Δ Delivered","Hearing Prev","Hearing Latest","Δ Hearing","Hearing Lapse","Discrepancy Delivered","BLO Docs","BLO Letter"].map(h=><th key={h}>{h}</th>)}
    </tr></thead><tbody>{shown.map(r=><tr key={r.ps} className={r.latestHearing===0?"zero":""}>
      <td>{r.ps}</td><td>{r.aero}</td><td>{r.supervisor}</td><td>{r.bloName}</td><td>{r.bloMobile}</td><td className="num">{r.noticeGenerated.toLocaleString()}</td><td className="num">{r.prevDelivered.toLocaleString()}</td><td className="num">{r.latestDelivered.toLocaleString()}</td><td className="num">{(r.latestDelivered-r.prevDelivered).toLocaleString()}</td><td className="num">{r.prevHearing.toLocaleString()}</td><td className="num">{r.latestHearing.toLocaleString()}</td><td className="num">{(r.latestHearing-r.prevHearing).toLocaleString()}</td><td className="num">{r.hearingLapse.toLocaleString()}</td><td className="num">{r.discrepancyDelivered.toLocaleString()}</td><td className="num">{r.bloDocs.toLocaleString()}</td><td className="num">{r.bloLetter.toLocaleString()}</td>
    </tr>)}</tbody></table></div>
    <div className="section card"><h2>How the update works</h2><div className="note">Previous ECI is the comparison baseline. Latest ECI replaces current metrics. Difference columns are Latest − Previous. Notice Generated is taken from Latest. BLO/Other data is joined by PS. The master mapping is not recalculated from the ECI report, so officer/BLO relationships stay fixed while daily numbers change.</div></div>
  </main>;
}
