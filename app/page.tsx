"use client";

// AC-34 build-fix checkpoint: keep the deployed page in sync with the latest clean report layout.

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
  discrepancy:["discrep notices delivered","discrepancy notices delivered","discrep notices delivery","anomaly notices delivered","anomaly delivered","discrepancy delivered","discrepancy/anomaly delivered","discrepancy + anomaly delivered","no of discrepancy notices delivered","no of anomaly notices delivered","discrepancy notice delivered","discrep notices delivered","notice delivered by blo of discrepancies","notice delivered by blo of discrepancy","notice delivered by blo"],
  letter:["blo letter uploaded","blo letters uploaded","blo letter","blo letter uploaded by blo","letters uploaded by blo","no of blo letter uploaded","blo letters","blo letter uploaded no mapping"]
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

function findDataCol(headers:string[], names:string[], exclude:number[]=[]){
  const hs=headers.map(norm);
  const blocked=new Set(exclude);
  // For metric columns, prefer exact header matches only. This prevents
  // fuzzy matching from accidentally selecting the PS No. column.
  for(const n of names){
    const target=norm(n);
    const i=hs.findIndex((h,idx)=>!blocked.has(idx) && h===target);
    if(i>=0)return i;
  }
  // Controlled fuzzy fallback, never allowed to use an excluded column.
  for(const n of names){
    const target=norm(n);
    if(target.length<5) continue;
    const i=hs.findIndex((h,idx)=>!blocked.has(idx) && (h.includes(target) || target.includes(h)));
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
  let bestScore=-1;
  for(let i=0;i<Math.min(raw.length,60);i++){
    const h=(raw[i]||[]).map(norm);
    const joined=h.join(" | ");
    let score=0;
    if(h.some(x=>["p s no","ps no","ps number","part no","part number"].includes(x)))score+=5;
    if(joined.includes("notice generated"))score+=3;
    if(joined.includes("notice delivered")||joined.includes("notices delivered"))score+=2;
    if(joined.includes("hearings held")||joined.includes("hearing held"))score+=2;
    if(joined.includes("discrep"))score+=2;
    if(joined.includes("blo docs")||joined.includes("documents uploaded by blo"))score+=2;
    if(joined.includes("blo letter"))score+=2;
    if(score>bestScore){bestScore=score;headerIndex=i;}
  }
  const headers=(raw[headerIndex]||[]).map((x:any)=>String(x??""));
  return {headers, rows:raw.slice(headerIndex+1)};
}

function parseAcTotals(headers:string[], rows:any[][]){
  const psI=findCol(headers,aliases.ps);
  const docsI=findDataCol(headers,aliases.docs,[psI]);
  const discI=findDataCol(headers,aliases.discrepancy,[psI]);
  const letterI=findDataCol(headers,aliases.letter,[psI]);

  // Prefer the explicit AC TOTAL / GRAND TOTAL row from the uploaded
  // BLO/Letter workbook. This is the authoritative total for these metrics.
  const totalRow=rows.find(r=>{
    const text=r.map((v:any)=>norm(v)).join(" | ");
    return text.includes("ac total") || text.includes("grand total") || text==="total" || text.startsWith("total ");
  });

  const value=(row:any[], idx:number)=>{
    if(idx<0 || !row) return undefined;
    const v=num(row[idx]);
    return Number.isFinite(v)?v:undefined;
  };

  if(totalRow){
    return {
      totalDocs:value(totalRow,docsI),
      totalDisc:value(totalRow,discI),
      totalLetter:value(totalRow,letterI)
    };
  }

  // Fallback: if the workbook has no explicit total row, sum the PS-wise
  // values rather than using an unrelated ECI total.
  let totalDocs=0,totalDisc=0,totalLetter=0;
  for(const r of rows){
    const ps=psI>=0?Math.round(num(r[psI])):0;
    if(!ps) continue;
    if(docsI>=0) totalDocs+=num(r[docsI]);
    if(discI>=0) totalDisc+=num(r[discI]);
    if(letterI>=0) totalLetter+=num(r[letterI]);
  }
  return {totalDocs,totalDisc,totalLetter};
}

function parseMetricFile(parsed:{headers:string[];rows:any[][]}, kind:"eci"|"blo"){
  const {headers,rows}=parsed;
  const psI=findCol(headers,aliases.ps);
  const genI=findCol(headers,aliases.generated);
  const delI=findCol(headers,aliases.delivered);
  const hearI=findCol(headers,aliases.hearing);
  const lapseI=findCol(headers,aliases.lapse);
  // BLO/Discrepancy metrics must NEVER resolve to the PS No. column.
  const docsI=findDataCol(headers,aliases.docs,[psI]);
  const discI=findDataCol(headers,aliases.discrepancy,[psI]);
  const letterI=findDataCol(headers,aliases.letter,[psI]);
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
  (out as any).acTotals=parseAcTotals(headers,rows);
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
    // The uploaded BLO/Other workbook is the authoritative PS-wise source
    // for BLO document and BLO letter metrics. ECI remains the source for
    // notice/hearing metrics.
    return {...m,ps,noticeGenerated:l.noticeGenerated??m.noticeGenerated??0,prevDelivered,latestDelivered,prevHearing,latestHearing,hearingLapse:l.hearingLapse??m.hearingLapse??0,
      discrepancyDelivered:b.discrepancyDelivered??l.discrepancyDelivered??m.discrepancyDelivered??0,
      bloDocs:b.bloDocs??l.bloDocs??0,
      bloLetter:b.bloLetter??l.bloLetter??m.bloLetter??0};
  });
}

const n=(x:number)=>x.toLocaleString("en-IN");
const pct=(a:number,b:number)=>b?((a/b)*100).toFixed(2)+"%":"0.00%";
const esc=(x:any)=>String(x??"").replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;");

const REPORT_AERO_ORDER=[
  "Smt. Anuja Trivedi","Smt. Parul Gupta","Sh. Parveen Kumar","Smt. Shashi Bala",
  "Sh. Subhashish Boss","Sh. Virender Singh","Sh. Ajay Kumar","Smt. Ranjana Sharma",
  "Smt. Mamta Meena","Sh. Rajesh Shriwastav","Smt. Saroj Meena","Sh. Hemvir",
  "Sh. Rakesh Yadav","Sh. Mohit","Smt. Vandana Bansal","Sh. Dharamvir Singh","Sh. Manoj Kumar"
];
function reportAeroIndex(name:string){
  const i=REPORT_AERO_ORDER.indexOf(name);
  return i>=0?i:999;
}

function buildGroups(rows:Row[]){
  const groups=new Map<string,Row[]>();
  for(const r of rows){const k=r.aero||"Unmapped";const a=groups.get(k)||[];a.push(r);groups.set(k,a);}
  const sum=(rs:Row[],k:keyof Row)=>rs.reduce((s,r)=>s+Number(r[k]||0),0);
  return Array.from(groups.entries()).map(([aero,rs])=>{
    const x={ps:rs.length,gen:sum(rs,"noticeGenerated"),pd:sum(rs,"prevDelivered"),ld:sum(rs,"latestDelivered"),ph:sum(rs,"prevHearing"),lh:sum(rs,"latestHearing"),lapse:sum(rs,"hearingLapse"),disc:sum(rs,"discrepancyDelivered"),docs:sum(rs,"bloDocs"),letter:sum(rs,"bloLetter")};
    return {aero,rs,x,dispose:x.lh+x.lapse?x.lh/(x.lh+x.lapse)*100:0};
  }).sort((a,b)=>reportAeroIndex(a.aero)-reportAeroIndex(b.aero));
}

function printReport(rows:Row[], sourceTotals?:{totalDocs?:number,totalDisc?:number,totalLetter?:number}){
  // PDF-only presentation layer. Dashboard state/calculations remain untouched.
  const exportRows:Row[]=rows.map(r=>({...r}));
  const groups=buildGroups(exportRows);
  const grand=groups.reduce((g,z)=>({
    ps:g.ps+z.x.ps,gen:g.gen+z.x.gen,pd:g.pd+z.x.pd,ld:g.ld+z.x.ld,
    ph:g.ph+z.x.ph,lh:g.lh+z.x.lh,lapse:g.lapse+z.x.lapse,
    disc:g.disc+z.x.disc,docs:g.docs+z.x.docs,letter:g.letter+z.x.letter
  }),{ps:0,gen:0,pd:0,ld:0,ph:0,lh:0,lapse:0,disc:0,docs:0,letter:0});
  if(sourceTotals?.totalDocs!=null) grand.docs=sourceTotals.totalDocs;
  if(sourceTotals?.totalDisc!=null) grand.disc=sourceTotals.totalDisc;
  if(sourceTotals?.totalLetter!=null) grand.letter=sourceTotals.totalLetter;

  const now=new Date();
  const stamp=now.toLocaleString("en-IN",{day:"2-digit",month:"2-digit",year:"numeric",hour:"2-digit",minute:"2-digit",second:"2-digit",hour12:false});
  const zeroRows=exportRows.filter(r=>r.latestHearing===0).sort((a,b)=>a.ps-b.ps);
  const zeroPct=grand.ps?zeroRows.length/grand.ps*100:0;
  const hearingPct=grand.gen?grand.lh/grand.gen*100:0;
  const disposalPct=grand.lh+grand.lapse?grand.lh/(grand.lh+grand.lapse)*100:0;

  const card=(label:string,value:string,sub:string,cls:string)=>{
    return '<div class="kpi '+cls+'"><div class="kpiLabel">'+esc(label)+'</div><div class="kpiValue">'+esc(value)+'</div><div class="kpiSub">'+esc(sub)+'</div></div>';
  };
  const metric=(label:string,prev:number,latest:number)=>{
    const diff=latest-prev;
    return '<div class="metricBox"><div class="metricLabel">'+esc(label)+'</div><div class="metricLine"><span>'+n(prev)+'</span><b>→</b><span>'+n(latest)+'</span></div><div class="metricDiff '+(diff>=0?'positive':'negative')+'">'+(diff>=0?'+':'')+n(diff)+'</div><div class="metricCaption">Previous&nbsp;&nbsp;→&nbsp;&nbsp;Latest&nbsp;&nbsp;|&nbsp;&nbsp;Difference</div></div>';
  };

  let h='<!doctype html><html><head><meta charset="utf-8"><title>AC-34 MATIALA - NOTICE & HEARING MONITORING REPORT</title>';
  h+='<style>';
  h+='@page{size:A4 landscape;margin:12mm 10mm 15mm 10mm}';
  h+='*{box-sizing:border-box}html,body{margin:0;padding:0;background:#fff;color:#17212b;font-family:Arial,Helvetica,sans-serif;font-size:9px}body{counter-reset:page}';
  h+='.page{page-break-after:always;position:relative;min-height:170mm;padding-bottom:13mm}.page:last-child{page-break-after:auto}';
  h+='.cover{min-height:180mm;display:flex;flex-direction:column;justify-content:center;padding:15mm 18mm;background:linear-gradient(135deg,#f4f8fb 0%,#fff 55%,#eef5fa 100%);border:1px solid #cbd8e3}';
  h+='.brand{font-size:12px;font-weight:800;letter-spacing:1.5px;color:#0e4d7e;text-transform:uppercase;margin-bottom:12px}.coverTitle{font-size:30px;line-height:1.12;font-weight:900;color:#123d5d;margin:0}.coverSub{font-size:16px;font-weight:700;color:#465866;margin-top:8px}.coverRule{width:90px;height:4px;background:#0e4d7e;margin:18px 0}.coverMeta{font-size:11px;color:#526574;line-height:1.8}.coverMeta b{color:#17212b}.coverNote{margin-top:20px;padding:10px 12px;border-left:4px solid #0e4d7e;background:#fff;border:1px solid #d7e1e8;font-size:9.5px;color:#52606b}';
  h+='.sectionTitle{font-size:16px;font-weight:900;color:#0e4d7e;border-bottom:3px solid #0e4d7e;padding:0 0 5px;margin:0 0 9px}.sectionSub{font-size:9px;color:#60717d;margin:-4px 0 8px}.kpiGrid{display:grid;grid-template-columns:repeat(6,1fr);gap:6px;margin-bottom:10px}.kpi{border:1px solid #cbd8e3;border-top:4px solid #0e4d7e;border-radius:3px;padding:7px 8px;min-height:58px;background:#fff}.kpi.green{border-top-color:#4f8a55}.kpi.orange{border-top-color:#d18b32}.kpi.red{border-top-color:#b73737}.kpi.blue{border-top-color:#0e4d7e}.kpiLabel{font-size:7.5px;text-transform:uppercase;font-weight:800;color:#61727f;letter-spacing:.3px}.kpiValue{font-size:16px;font-weight:900;color:#152d3e;margin-top:3px}.kpiSub{font-size:7px;color:#74818a;margin-top:2px}.metricGrid{display:grid;grid-template-columns:1fr 1fr;gap:8px;margin:8px 0 12px}.metricBox{border:1px solid #cbd8e3;border-radius:3px;padding:8px 10px;background:#f8fbfd}.metricLabel{font-size:9px;font-weight:900;color:#35566d;text-transform:uppercase}.metricLine{display:flex;align-items:center;gap:14px;font-size:18px;font-weight:900;margin-top:4px}.metricLine b{color:#0e4d7e}.metricDiff{font-size:11px;font-weight:900;margin-top:2px}.metricDiff.positive{color:#2f6d38}.metricDiff.negative{color:#a02f2f}.metricCaption{font-size:7px;color:#74818a;margin-top:3px}';
  h+='.legend{display:flex;gap:18px;align-items:center;border:1px solid #d6e0e6;background:#fafcfd;padding:6px 8px;margin:8px 0}.legendItem{font-weight:800;font-size:8px}.swatch{display:inline-block;width:13px;height:10px;border:1px solid #aaa;margin-right:5px;vertical-align:-1px}.swatch.red{background:#ffd9d9}.swatch.blue{background:#d9eaff}.swatch.green{background:#dcebd5}';
  h+='table{width:100%;border-collapse:collapse;table-layout:fixed;margin-top:5px}th{background:#0e4d7e;color:#fff;font-weight:900;text-align:center;border:1px solid #17486b;padding:4px 3px;font-size:7.5px}td{border:1px solid #bfcbd3;padding:3px 3px;vertical-align:middle;font-size:7.5px;line-height:1.2;overflow-wrap:anywhere}td.num{text-align:right;white-space:nowrap}td.center{text-align:center}.wrap{white-space:normal;overflow-wrap:anywhere}.total td{background:#e8f0f5;font-weight:900;border-top:2px solid #7890a0}.rate{background:#dcebd5!important;color:#17212b;font-weight:900!important}.diff{font-weight:900}.zero td{background:#ffd9d9!important;color:#6e1111!important}.zeroBadge{display:inline-block;background:#b73737;color:#fff;font-size:6.5px;font-weight:900;padding:2px 4px;border-radius:2px;white-space:nowrap}.statusZero{font-weight:900;color:#a61f1f}.statusNormal{font-weight:800;color:#35566d}';
  h+='.officerHead{margin:11px 0 5px;padding:8px 10px;background:#e8f2f8;border-left:5px solid #0e4d7e;display:flex;justify-content:space-between;align-items:center;page-break-after:avoid}.officerName{font-size:14px;font-weight:900;color:#123d5d}.officerDesignation{font-size:9px;font-weight:800;color:#536873}.officerPs{font-size:10px;font-weight:900;color:#0e4d7e}.miniGrid{display:grid;grid-template-columns:repeat(8,1fr);gap:4px;margin-bottom:6px}.mini{border:1px solid #d0dbe2;background:#fbfdfe;padding:5px}.mini span{display:block;font-size:6.5px;color:#687983;text-transform:uppercase;font-weight:800}.mini b{display:block;font-size:10px;margin-top:2px;color:#203b4d}.subTitle{font-size:10px;font-weight:900;color:#35566d;margin:8px 0 4px}.avoidBreak{break-inside:avoid;page-break-inside:avoid}.zeroSummary{margin-bottom:9px}.footer{position:fixed;bottom:3mm;left:10mm;right:10mm;border-top:1px solid #ccd6dc;padding-top:3px;color:#687780;font-size:7px;display:flex;justify-content:space-between;z-index:10;background:#fff}.footer .pageNo:after{content:counter(page)}.footer .pagesNo:after{content:counter(pages)}';
  h+='@media print{body{counter-reset:page}.noPrint{display:none!important}.footer{position:fixed}.page{page-break-after:always}.page:last-child{page-break-after:auto}}';
  h+='</style></head><body>';

  // PAGE 1 — Executive Summary
  h+='<div class="page"><div class="cover">';
  h+='<div class="brand">Government Report • AC-34 Matiala • SIR-2026</div>';
  h+='<h1 class="coverTitle">AC-34 MATIALA<br>SIR-2026<br>NOTICE & HEARING<br>MONITORING REPORT</h1>';
  h+='<div class="coverSub">Previous ECI vs Latest ECI</div><div class="coverRule"></div>';
  h+='<div class="coverMeta"><b>Report Generated:</b> '+esc(stamp)+'<br><b>Data Scope:</b> '+n(grand.ps)+' PS &nbsp;|&nbsp; '+n(groups.length)+' AERO / Ad.AERO</div>';
  h+='<div class="coverNote">This PDF is generated directly from the processed dashboard data at the time the existing <b>Download PDF</b> action is used. The dashboard UI and upload workflow remain unchanged.</div>';
  h+='</div>';
  h+='<div style="margin-top:10px"><div class="sectionTitle">EXECUTIVE SUMMARY</div>';
  h+='<div class="kpiGrid">';
  h+=card("Total PS",n(grand.ps),"Polling Stations","blue");
  h+=card("Notice Generated",n(grand.gen),"Current dataset","blue");
  h+=card("Previous Notice Delivered",n(grand.pd),"Previous ECI","blue");
  h+=card("Latest Notice Delivered",n(grand.ld),"Latest ECI","green");
  h+=card("Delivery Difference",(grand.ld-grand.pd>=0?"+":"")+n(grand.ld-grand.pd),"Latest − Previous","green");
  h+=card("Previous Hearing Held",n(grand.ph),"Previous ECI","blue");
  h+=card("Latest Hearing Held",n(grand.lh),"Latest ECI","green");
  h+=card("Hearing Difference",(grand.lh-grand.ph>=0?"+":"")+n(grand.lh-grand.ph),"Latest − Previous","green");
  h+=card("% Hearing Held",hearingPct.toFixed(2)+"%","Latest Hearing / Notice Generated","green");
  h+=card("Hearing Lapsed",n(grand.lapse),"Latest ECI","orange");
  h+=card("Zero Hearing PS",n(zeroRows.length),"Latest Hearing Held = 0","red");
  h+=card("Zero Hearing %",zeroPct.toFixed(2)+"%","Of total PS","red");
  h+='</div>';
  h+='<div class="metricGrid">'+metric("Notice Delivered",grand.pd,grand.ld)+metric("Hearing Held",grand.ph,grand.lh)+'</div>';
  h+='<div class="legend"><span class="legendItem"><span class="swatch red"></span>RED = ZERO HEARING HELD</span><span class="legendItem"><span class="swatch green"></span>GREEN = POSITIVE LATEST / RATE</span><span class="legendItem"><span class="swatch blue"></span>BLUE = REPORT / SECTION HIGHLIGHT</span></div>';
  h+='</div></div>';

  // PAGE 2 — Consolidated report
  h+='<div class="page"><div class="sectionTitle">AERO / Ad.AERO CONSOLIDATED REPORT</div><div class="sectionSub">Officer-wise consolidation using the same rows and calculations already processed by the dashboard.</div>';
  h+='<table><thead><tr><th style="width:3%">S.No.</th><th style="width:14%">AERO / Ad.AERO</th><th style="width:7%">Designation</th><th style="width:4%">PS</th><th>Notice Generated</th><th>Prev Delivered</th><th>Latest Delivered</th><th>Delivery Diff</th><th>Prev Hearing</th><th>Latest Hearing</th><th>Hearing Diff</th><th>% Held</th><th>Hearing Lapsed</th><th>Zero Hearing PS</th><th>Disposal %</th></tr></thead><tbody>';
  groups.forEach((g,i)=>{
    const zero=g.rs.filter(r=>r.latestHearing===0).length;
    h+='<tr><td class="center">'+(i+1)+'</td><td class="wrap"><b>'+esc(g.aero)+'</b></td><td class="center">'+esc(g.rs[0]?.designation||"")+'</td><td class="center">'+g.x.ps+'</td><td class="num">'+n(g.x.gen)+'</td><td class="num">'+n(g.x.pd)+'</td><td class="num">'+n(g.x.ld)+'</td><td class="num diff">'+(g.x.ld-g.x.pd>=0?"+":"")+n(g.x.ld-g.x.pd)+'</td><td class="num">'+n(g.x.ph)+'</td><td class="num">'+n(g.x.lh)+'</td><td class="num diff">'+(g.x.lh-g.x.ph>=0?"+":"")+n(g.x.lh-g.x.ph)+'</td><td class="num rate">'+pct(g.x.lh,g.x.gen)+'</td><td class="num">'+n(g.x.lapse)+'</td><td class="center">'+zero+'</td><td class="num rate">'+g.dispose.toFixed(2)+'%</td></tr>';
  });
  h+='<tr class="total"><td colspan="4">GRAND TOTAL</td><td class="num">'+n(grand.gen)+'</td><td class="num">'+n(grand.pd)+'</td><td class="num">'+n(grand.ld)+'</td><td class="num diff">'+(grand.ld-grand.pd>=0?"+":"")+n(grand.ld-grand.pd)+'</td><td class="num">'+n(grand.ph)+'</td><td class="num">'+n(grand.lh)+'</td><td class="num diff">'+(grand.lh-grand.ph>=0?"+":"")+n(grand.lh-grand.ph)+'</td><td class="num rate">'+hearingPct.toFixed(2)+'%</td><td class="num">'+n(grand.lapse)+'</td><td class="center">'+zeroRows.length+'</td><td class="num rate">'+disposalPct.toFixed(2)+'%</td></tr></tbody></table>';
  h+='<div class="subTitle">HEARING HELD % — OFFICER-WISE FACTUAL TABLE</div><table><thead><tr><th style="width:4%">S.No.</th><th>Officer</th><th>Designation</th><th>PS</th><th>Hearing Held</th><th>Notice Generated</th><th>% Hearing Held</th></tr></thead><tbody>';
  groups.forEach((g,i)=>{h+='<tr><td class="center">'+(i+1)+'</td><td class="wrap">'+esc(g.aero)+'</td><td class="center">'+esc(g.rs[0]?.designation||"")+'</td><td class="center">'+g.x.ps+'</td><td class="num">'+n(g.x.lh)+'</td><td class="num">'+n(g.x.gen)+'</td><td class="num rate">'+pct(g.x.lh,g.x.gen)+'</td></tr>';});
  h+='</tbody></table></div>';

  // PAGE 3+ — detailed officer sections
  groups.forEach(g=>{
    h+='<div class="page"><div class="officerHead"><div><div class="officerName">'+esc(g.aero)+'</div><div class="officerDesignation">'+esc(g.rs[0]?.designation||"")+'</div></div><div class="officerPs">'+g.x.ps+' PS</div></div>';
    h+='<div class="miniGrid">';
    h+=card("Notice Generated",n(g.x.gen),"","blue");
    h+=card("Previous Delivered",n(g.x.pd),"","blue");
    h+=card("Latest Delivered",n(g.x.ld),"","green");
    h+=card("Previous Hearing",n(g.x.ph),"","blue");
    h+=card("Latest Hearing",n(g.x.lh),"","green");
    h+=card("% Held",pct(g.x.lh,g.x.gen),"","green");
    h+=card("Hearing Lapse",n(g.x.lapse),"","orange");
    h+=card("Zero Hearing PS",n(g.rs.filter(r=>r.latestHearing===0).length),"","red");
    h+='</div>';
    h+='<div class="metricGrid">'+metric("Notice Delivered",g.x.pd,g.x.ld)+metric("Hearing Held",g.x.ph,g.x.lh)+'</div>';
    h+='<div class="subTitle">PS-WISE DETAIL TABLE</div>';
    h+='<table><thead><tr><th style="width:3%">S.No.</th><th style="width:4%">PS</th><th style="width:14%">BLO Name</th><th style="width:9%">BLO Mobile</th><th style="width:15%">BLO Supervisor</th><th>Notice Generated</th><th>Prev Delivered</th><th>Latest Delivered</th><th>Delivery Diff</th><th>Prev Hearing</th><th>Latest Hearing</th><th>Hearing Diff</th><th>% Held</th><th>Hearing Lapsed</th><th>Status</th></tr></thead><tbody>';
    g.rs.slice().sort((a,b)=>a.ps-b.ps).forEach((r,i)=>{
      const isZero=r.latestHearing===0;
      h+='<tr class="'+(isZero?'zero':'')+'"><td class="center">'+(i+1)+'</td><td class="center">'+r.ps+'</td><td class="wrap"><b>'+esc(r.bloName)+'</b></td><td class="center">'+esc(r.bloMobile)+'</td><td class="wrap">'+esc(r.supervisor)+'</td><td class="num">'+n(r.noticeGenerated)+'</td><td class="num">'+n(r.prevDelivered)+'</td><td class="num">'+n(r.latestDelivered)+'</td><td class="num diff">'+(r.latestDelivered-r.prevDelivered>=0?"+":"")+n(r.latestDelivered-r.prevDelivered)+'</td><td class="num">'+n(r.prevHearing)+'</td><td class="num">'+n(r.latestHearing)+'</td><td class="num diff">'+(r.latestHearing-r.prevHearing>=0?"+":"")+n(r.latestHearing-r.prevHearing)+'</td><td class="num rate">'+pct(r.latestHearing,r.noticeGenerated)+'</td><td class="num">'+n(r.hearingLapse)+'</td><td class="center '+(isZero?'statusZero':'statusNormal')+'">'+(isZero?'ZERO HEARING':'ACTIVE')+'</td></tr>';
    });
    h+='<tr class="total"><td colspan="5">TOTAL — '+esc(g.aero)+' ('+g.x.ps+' PS)</td><td class="num">'+n(g.x.gen)+'</td><td class="num">'+n(g.x.pd)+'</td><td class="num">'+n(g.x.ld)+'</td><td class="num diff">'+(g.x.ld-g.x.pd>=0?"+":"")+n(g.x.ld-g.x.pd)+'</td><td class="num">'+n(g.x.ph)+'</td><td class="num">'+n(g.x.lh)+'</td><td class="num diff">'+(g.x.lh-g.x.ph>=0?"+":"")+n(g.x.lh-g.x.ph)+'</td><td class="num rate">'+pct(g.x.lh,g.x.gen)+'</td><td class="num">'+n(g.x.lapse)+'</td><td class="center">'+g.rs.filter(r=>r.latestHearing===0).length+' Zero</td></tr></tbody></table>';
    h+='</div>';
  });

  // Final section — zero hearing
  h+='<div class="page"><div class="sectionTitle">ZERO HEARING PS — ACTION REQUIRED</div><div class="sectionSub">All PS where Latest Hearing Held = 0. The list is generated dynamically from the latest dashboard data.</div>';
  h+='<div class="kpiGrid" style="grid-template-columns:repeat(4,1fr)">'+card("Total Zero Hearing PS",n(zeroRows.length),"Latest Hearing = 0","red")+card("Total PS",n(grand.ps),"Current dataset","blue")+card("Zero Hearing %",zeroPct.toFixed(2)+"%","Of total PS","red")+card("Latest Hearing Held",n(grand.lh),"Current dataset","green")+'</div>';
  h+='<div class="subTitle">AERO / Ad.AERO ZERO HEARING SUMMARY</div><table class="zeroSummary"><thead><tr><th>AERO / Ad.AERO</th><th>Designation</th><th>Total PS</th><th>Zero Hearing PS</th><th>Zero Hearing %</th></tr></thead><tbody>';
  groups.forEach(g=>{const z=g.rs.filter(r=>r.latestHearing===0).length;h+='<tr><td class="wrap">'+esc(g.aero)+'</td><td class="center">'+esc(g.rs[0]?.designation||"")+'</td><td class="center">'+g.x.ps+'</td><td class="center statusZero">'+z+'</td><td class="num">'+(g.x.ps?((z/g.x.ps)*100).toFixed(2):"0.00")+'%</td></tr>';});
  h+='</tbody></table>';
  h+='<div class="subTitle">ZERO HEARING PS — DETAILED LIST</div><table><thead><tr><th style="width:3%">S.No.</th><th style="width:15%">AERO / Ad.AERO</th><th style="width:15%">Supervisor</th><th style="width:4%">PS No.</th><th style="width:15%">BLO Name</th><th>Notice Generated</th><th>Notice Delivered</th><th>Hearing Held</th><th>Hearing Lapsed</th><th>Status</th></tr></thead><tbody>';
  zeroRows.forEach((r,i)=>{h+='<tr class="zero"><td class="center">'+(i+1)+'</td><td class="wrap">'+esc(r.aero)+'</td><td class="wrap">'+esc(r.supervisor)+'</td><td class="center">'+r.ps+'</td><td class="wrap">'+esc(r.bloName)+'</td><td class="num">'+n(r.noticeGenerated)+'</td><td class="num">'+n(r.latestDelivered)+'</td><td class="num">0</td><td class="num">'+n(r.hearingLapse)+'</td><td class="center statusZero">ZERO HEARING</td></tr>';});
  h+='<tr class="total"><td colspan="5">TOTAL ZERO HEARING PS</td><td colspan="5" class="center">'+n(zeroRows.length)+' / '+n(grand.ps)+' &nbsp; | &nbsp; '+zeroPct.toFixed(2)+'%</td></tr></tbody></table>';
  h+='<div class="legend"><span class="legendItem"><span class="swatch red"></span>RED = ZERO HEARING HELD</span><span class="legendItem">Report timestamp: '+esc(stamp)+'</span></div></div>';

  h+='<div class="footer"><span>AC-34 MATIALA | SIR-2026 | NOTICE & HEARING MONITORING REPORT</span><span>Report Generated: '+esc(stamp)+'</span><span>Page <span class="pageNo"></span> of <span class="pagesNo"></span></span></div>';
  h+='</body></html>';

  const w=window.open("","_blank","width=1800,height=1200");
  if(!w){alert("Please allow pop-ups for PDF export.");return;}
  w.document.write(h); w.document.close(); w.focus();
  setTimeout(()=>{w.focus();w.print();},900);
}
export default function Page(){
  const [master,setMaster]=useState<Map<number,any>>(new Map());
  const [rows,setRows]=useState<Row[]>([]);
  const [previous,setPrevious]=useState<Map<number,any>>(new Map());
  const [latest,setLatest]=useState<Map<number,any>>(new Map());
  const [blo,setBlo]=useState<Map<number,any>>(new Map());
  const [sourceTotals,setSourceTotals]=useState<{totalDocs?:number,totalDisc?:number,totalLetter?:number}>({});
  const [filterAero,setFilterAero]=useState("ALL");
  const [filterSup,setFilterSup]=useState("ALL");
  const [search,setSearch]=useState("");
  const [error,setError]=useState("");
  const [filterDesignation,setFilterDesignation]=useState("ALL");
  const [filterBlo,setFilterBlo]=useState("ALL");
  const [filterHearing,setFilterHearing]=useState("ALL");
  const [filterZero,setFilterZero]=useState("ALL");
  const [filterHeld,setFilterHeld]=useState("ALL");
  const [filterLapse,setFilterLapse]=useState("ALL");
  const [minPct,setMinPct]=useState("");
  const [maxPct,setMaxPct]=useState("");
  const [selectedPs,setSelectedPs]=useState<number|null>(null);
  const [sortField,setSortField]=useState<string>("ps");
  const [sortDir,setSortDir]=useState<"asc"|"desc">("asc");
  const [columns,setColumns]=useState<Record<string,boolean>>({ps:true,aero:true,designation:true,supervisor:true,bloName:true,bloMobile:true,gen:true,pd:true,ld:true,diffDel:true,ph:true,lh:true,diffHear:true,pct:true,lapse:true});

  useEffect(()=>{try{const saved=localStorage.getItem("ac34-master-mapping");if(saved){const arr=JSON.parse(saved);setMaster(new Map(arr));}}catch{}},[]);
  useEffect(()=>{if(master.size)localStorage.setItem("ac34-master-mapping",JSON.stringify(Array.from(master.entries())));},[master]);

  const load=(setter:any,kind:"eci"|"blo")=>(e:React.ChangeEvent<HTMLInputElement>)=>{const f=e.target.files?.[0];if(!f)return;setError("");readRows(f).then(p=>{const parsed:any=kind==="eci"?parseMetricFile(p,"eci"):parseMetricFile(p,"blo");if(kind==="blo" && parsed.acTotals)setSourceTotals(parsed.acTotals);setter(parsed);}).catch(err=>setError(String(err)));};

  useEffect(()=>{if(master.size) setRows(merge(master,previous,latest,blo));},[master,previous,latest,blo]);

  const aeros=useMemo(()=>["ALL",...Array.from(new Set(rows.map(r=>r.aero).filter(Boolean)))],[rows]);
  const sups=useMemo(()=>["ALL",...Array.from(new Set(rows.map(r=>r.supervisor).filter(Boolean)))],[rows]);
  const designations=useMemo(()=>["ALL",...Array.from(new Set(rows.map(r=>r.designation).filter(Boolean))).sort()],[rows]);
  const blos=useMemo(()=>["ALL",...Array.from(new Set(rows.map(r=>r.bloName).filter(Boolean))).sort()],[rows]);
  const pendingFor=(r:Row)=>Math.max(r.latestDelivered-r.latestHearing-r.hearingLapse,0);
  const disposePct=(r:Row)=>(r.latestHearing+r.hearingLapse)?r.latestHearing/(r.latestHearing+r.hearingLapse)*100:0;
  const hearingStatus=(r:Row)=>r.latestHearing===0?"ZERO HEARING":pendingFor(r)>0?"PENDING":"DISPOSED";
  const shown=useMemo(()=>rows.filter(r=>{
    const hp=r.noticeGenerated?(r.latestHearing/r.noticeGenerated)*100:0;
    const q=search.toLowerCase();
    if(filterAero!=="ALL"&&r.aero!==filterAero)return false;
    if(filterSup!=="ALL"&&r.supervisor!==filterSup)return false;
    if(filterDesignation!=="ALL"&&r.designation!==filterDesignation)return false;
    if(filterBlo!=="ALL"&&r.bloName!==filterBlo)return false;
    if(filterHearing!=="ALL"&&hearingStatus(r)!==filterHearing)return false;
    if(filterZero==="YES"&&r.latestHearing!==0)return false;
    if(filterZero==="NO"&&r.latestHearing===0)return false;
    if(filterHeld==="YES"&&r.latestHearing<=0)return false;
    if(filterHeld==="NO"&&r.latestHearing>0)return false;
    if(filterLapse==="YES"&&r.hearingLapse<=0)return false;
    if(filterLapse==="NO"&&r.hearingLapse>0)return false;
    if(minPct!==""&&hp<Number(minPct))return false;
    if(maxPct!==""&&hp>Number(maxPct))return false;
    if(q&&!((r.ps+" "+r.bloName+" "+r.supervisor+" "+r.aero+" "+r.designation).toLowerCase().includes(q)))return false;
    return true;
  }).sort((a,b)=>{
    const av:any=sortField==="ps"?a.ps:sortField==="aero"?a.aero:sortField==="supervisor"?a.supervisor:sortField==="bloName"?a.bloName:(a as any)[sortField];
    const bv:any=sortField==="ps"?b.ps:sortField==="aero"?b.aero:sortField==="supervisor"?b.supervisor:sortField==="bloName"?b.bloName:(b as any)[sortField];
    const aa=typeof av==="number"?av:String(av??"").toLowerCase(),bb=typeof bv==="number"?bv:String(bv??"").toLowerCase();
    const cmp=aa<bb?-1:aa>bb?1:0;return sortDir==="asc"?cmp:-cmp;
  }),[rows,filterAero,filterSup,filterDesignation,filterBlo,filterHearing,filterZero,filterHeld,filterLapse,minPct,maxPct,search,sortField,sortDir]);
  const totals=useMemo(()=>rows.reduce((a,r)=>({gen:a.gen+r.noticeGenerated,pd:a.pd+r.prevDelivered,del:a.del+r.latestDelivered,ph:a.ph+r.prevHearing,lh:a.lh+r.latestHearing,lapse:a.lapse+r.hearingLapse,zero:a.zero+(r.latestHearing===0?1:0)}),{gen:0,pd:0,del:0,ph:0,lh:0,lapse:0,zero:0}),[rows]);
  const officerSummary=useMemo(()=>{const m=new Map<string,any>();for(const r of rows){const k=r.aero||"Unmapped";const x=m.get(k)||{aero:k,designation:r.designation,ps:0,gen:0,del:0,dh:0,zero:0};x.ps++;x.gen+=r.noticeGenerated;x.del+=r.latestDelivered;x.dh+=r.latestHearing;x.zero+=r.latestHearing===0?1:0;m.set(k,x);}return Array.from(m.values()).sort((a,b)=>reportAeroIndex(a.aero)-reportAeroIndex(b.aero));},[rows]);
  const supervisorSummary=useMemo(()=>{const m=new Map<string,any>();for(const r of rows){const k=r.supervisor||"Unmapped";const x=m.get(k)||{name:k,aeros:new Set<string>(),ps:0,blos:new Set<string>(),gen:0,del:0,hearing:0,lapse:0,zero:0};x.aeros.add(r.aero);x.ps++;x.blos.add(r.bloName);x.gen+=r.noticeGenerated;x.del+=r.latestDelivered;x.hearing+=r.latestHearing;x.lapse+=r.hearingLapse;x.zero+=r.latestHearing===0?1:0;m.set(k,x);}return Array.from(m.values()).map(x=>({...x,aero:Array.from(x.aeros).join(", "),blosCount:x.blos.size,pending:Math.max(x.del-x.hearing-x.lapse,0),heldPct:x.gen?x.hearing/x.gen*100:0,dispose:x.hearing+x.lapse?x.hearing/(x.hearing+x.lapse)*100:0})).sort((a,b)=>a.name.localeCompare(b.name));},[rows]);
  const zeroRows=useMemo(()=>rows.filter(r=>r.latestHearing===0),[rows]);

  const exportWorkbook=()=>{
    const wb=XLSX.utils.book_new();
    const add=(name:string,data:any[])=>{const ws=XLSX.utils.json_to_sheet(data);ws["!cols"]=Object.keys(data[0]||{}).map(k=>({wch:Math.min(34,Math.max(12,k.length+3))}));ws["!freeze"]={xSplit:0,ySplit:1};if(ws["!ref"])ws["!autofilter"]={ref:ws["!ref"]};XLSX.utils.book_append_sheet(wb,ws,name.slice(0,31));};
    add("Executive Summary",[
      {Metric:"Report Generated",Value:new Date().toLocaleString("en-IN")},
      {Metric:"Total PS",Value:rows.length},{Metric:"Notice Generated",Value:totals.gen},
      {Metric:"Notice Delivered Previous",Value:totals.pd},{Metric:"Notice Delivered Latest",Value:totals.del},{Metric:"Delivery Difference",Value:totals.del-totals.pd},
      {Metric:"Hearing Held Previous",Value:totals.ph},{Metric:"Hearing Held Latest",Value:totals.lh},{Metric:"Hearing Difference",Value:totals.lh-totals.ph},
      {Metric:"Overall Hearing Held %",Value:totals.gen?(totals.lh/totals.gen*100).toFixed(2)+"%":"0.00%"},{Metric:"Hearing Lapsed",Value:totals.lapse},
      {Metric:"Zero Hearing PS",Value:totals.zero}
    ]);
    add("AERO-wise Report",officerSummary.map((x:any)=>({AERO:x.aero,Designation:x.designation,"No. of PS":x.ps,"Notice Generated":x.gen,"Delivered Previous":x.pd,"Delivered Latest":x.del,"Delivery Difference":x.del-x.pd,"Hearing Previous":x.ph,"Hearing Latest":x.lh,"Hearing Difference":x.lh-x.ph,"Hearing %":x.gen?(x.lh/x.gen*100).toFixed(2)+"%":"0.00%","Hearing Lapsed":x.lapse,"Zero Hearing PS":x.zero,"Pending Hearing PS":Math.max(x.del-x.lh-x.lapse,0),"Disposal %":x.lh+x.lapse?(x.lh/(x.lh+x.lapse)*100).toFixed(2)+"%":"0.00%"})));
    add("Supervisor-wise",supervisorSummary.map((x:any)=>({Supervisor:x.name,"AERO / Ad.AERO":x.aero,"Total PS":x.ps,"Total BLOs":x.blosCount,"Notice Generated":x.gen,"Notice Delivered":x.del,"Hearing Held":x.hearing,"Hearing Pending":x.pending,"Hearing Lapsed":x.lapse,"Zero Hearing PS":x.zero,"Hearing %":x.heldPct.toFixed(2)+"%","Disposal %":x.dispose.toFixed(2)+"%"})));
    add("BLO-wise",rows.map(r=>({AERO:r.aero,Designation:r.designation,Supervisor:r.supervisor,"BLO Name":r.bloName,"PS No.":r.ps,"Notice Generated":r.noticeGenerated,"Notice Delivered":r.latestDelivered,"Hearing Held":r.latestHearing,"Hearing Lapsed":r.hearingLapse,"Pending Hearing":pendingFor(r),"Hearing %":r.noticeGenerated?(r.latestHearing/r.noticeGenerated*100).toFixed(2)+"%":"0.00%","Disposal %":disposePct(r).toFixed(2)+"%"})));
    add("Zero Hearing PS",zeroRows.map(r=>({AERO:r.aero,Designation:r.designation,Supervisor:r.supervisor,"PS No.":r.ps,"BLO Name":r.bloName,"Notice Generated":r.noticeGenerated,"Notice Delivered":r.latestDelivered,"Hearing Held":r.latestHearing,"Hearing Lapsed":r.hearingLapse,Status:"ZERO HEARING"})));
    add("PS-wise Detailed",rows.map(r=>({PS:r.ps,AERO:r.aero,Designation:r.designation,Supervisor:r.supervisor,"BLO Name":r.bloName,"BLO Mobile":r.bloMobile,"Notice Generated":r.noticeGenerated,"Delivered Previous":r.prevDelivered,"Delivered Latest":r.latestDelivered,"Delivery Difference":r.latestDelivered-r.prevDelivered,"Hearing Previous":r.prevHearing,"Hearing Latest":r.latestHearing,"Hearing Difference":r.latestHearing-r.prevHearing,"Hearing %":r.noticeGenerated?(r.latestHearing/r.noticeGenerated*100).toFixed(2)+"%":"0.00%","Hearing Lapsed":r.hearingLapse,"Pending Hearing":pendingFor(r),"Disposal %":disposePct(r).toFixed(2)+"%","Status":hearingStatus(r)})));
    XLSX.writeFile(wb,"AC34_SIR_Monitoring_Control_Report_"+new Date().toISOString().slice(0,10)+".xlsx");
  };
  const exportZeroExcel=()=>{const wb=XLSX.utils.book_new();const data=zeroRows.map(r=>({AERO:r.aero,Designation:r.designation,Supervisor:r.supervisor,"PS No.":r.ps,"BLO Name":r.bloName,"Notice Generated":r.noticeGenerated,"Notice Delivered":r.latestDelivered,"Hearing Held":r.latestHearing,"Hearing Lapsed":r.hearingLapse,Status:"ZERO HEARING"}));const ws=XLSX.utils.json_to_sheet(data);ws["!cols"]=Object.keys(data[0]||{}).map(k=>({wch:Math.max(14,k.length+3)}));ws["!freeze"]={xSplit:0,ySplit:1};XLSX.utils.book_append_sheet(wb,ws,"Zero Hearing PS");XLSX.writeFile(wb,"AC34_Zero_Hearing_PS_"+new Date().toISOString().slice(0,10)+".xlsx");};
  const resetFilters=()=>{setFilterAero("ALL");setFilterSup("ALL");setFilterDesignation("ALL");setFilterBlo("ALL");setFilterHearing("ALL");setFilterZero("ALL");setFilterHeld("ALL");setFilterLapse("ALL");setMinPct("");setMaxPct("");setSearch("");};

  return <main className="wrap">
    <div className="top"><div><div className="title">AC-34 MATIALA — NOTICE & HEARING DASHBOARD</div><div className="sub">SIR-2026 • Fixed PS → AERO/Ad.AERO → BLO Supervisor → BLO mapping • Latest ECI comparison</div></div><div className="badge">430 PS MASTER STRUCTURE</div></div>
    <div className="grid">
      <div className="card"><h3>① Previous ECI Excel</h3><p>Baseline report used for comparison.</p><input className="input" type="file" accept=".xlsx,.xls,.csv" onChange={load(setPrevious,"eci")}/></div>
      <div className="card"><h3>② Latest ECI Excel</h3><p>Current report. Notice Generated + latest hearing status come from this file.</p><input className="input" type="file" accept=".xlsx,.xls,.csv" onChange={load(setLatest,"eci")}/></div>
    </div>
    {error&&<div className="error">{error}</div>}
    <div className="section card"><h3>One-time Master Mapping</h3><div className="note">Authoritative mapping: PS → AERO/Ad.AERO → BLO Supervisor → BLO. A replacement is accepted only when the 430-PS structure is present and required mapping fields are populated.</div><div style={{marginTop:10}}><input className="input" type="file" accept=".xlsx,.xls,.csv" onChange={e=>{const f=e.target.files?.[0];if(!f)return;setError("");readRows(f).then(p=>{const m=parseMapping(p);const missing=Array.from({length:430},(_,i)=>i+1).filter(ps=>!m.has(ps));const bad=Array.from(m.values()).filter(v=>!v.aero||!v.bloName||!v.supervisor);if(m.size!==430||missing.length||bad.length){setError("Master mapping validation failed: the file must contain exactly the 430 PS records with AERO, BLO Name and BLO Supervisor populated.");return;}setMaster(m);setError("");}).catch(err=>setError(String(err)));}}/></div></div>
    {master.size===0&&<div className="error">Master mapping is not loaded yet. Upload the one-time mapping Excel below. It must contain PS No, AERO/Ad.AERO, BLO Supervisor and BLO Name. After upload it is saved in this browser for future ECI updates.</div>}
    <div className="section execPanel">
      <div className="sectionHead"><div><h2>Executive Summary</h2><div className="note">Dynamic values from the uploaded Previous and Latest ECI reports.</div></div><div className="buttonRow"><button className="btn" onClick={exportWorkbook}>Export Full Excel</button><button className="btn primary" onClick={()=>printReport(rows.map(r=>({...r})))}>Download Full PDF</button></div></div>
      <div className="stats execStats">
        {[
          ["Total PS",rows.length,"info"],["Notice Generated",totals.gen,"info"],["Delivered Previous",totals.pd,"info"],["Delivered Latest",totals.del,"good"],
          ["Delivery Difference",totals.del-totals.pd,"good"],["Hearing Previous",totals.ph,"info"],["Hearing Latest",totals.lh,"good"],["Hearing Difference",totals.lh-totals.ph,"good"],
          ["Overall Hearing %",totals.gen?(totals.lh/totals.gen*100).toFixed(2)+"%":"0.00%","good"],["Hearing Lapsed",totals.lapse,"attention"],["Zero Hearing PS",totals.zero,"critical"],["Zero Hearing %",rows.length?(totals.zero/rows.length*100).toFixed(2)+"%":"0.00%","critical"]
        ].map(([k,v,t])=><div className={"stat metric "+t} key={String(k)}><span>{k}</span><b>{typeof v==="number"?v.toLocaleString("en-IN"):v}</b></div>)}
      </div>
      <div className="compareGrid"><div className="compareCard"><b>Notice Delivered</b><div><span>{totals.pd.toLocaleString("en-IN")}</span><span>→</span><strong>{totals.del.toLocaleString("en-IN")}</strong><em className="pos">{totals.del>=totals.pd?"+":""}{(totals.del-totals.pd).toLocaleString("en-IN")}</em></div><small>% Change: {totals.pd?((totals.del-totals.pd)/totals.pd*100).toFixed(2):"0.00"}%</small></div><div className="compareCard"><b>Hearing Held</b><div><span>{totals.ph.toLocaleString("en-IN")}</span><span>→</span><strong>{totals.lh.toLocaleString("en-IN")}</strong><em className="pos">{totals.lh>=totals.ph?"+":""}{(totals.lh-totals.ph).toLocaleString("en-IN")}</em></div><small>% Change: {totals.ph?((totals.lh-totals.ph)/totals.ph*100).toFixed(2):"0.00"}%</small></div><div className="compareCard"><b>Hearing Lapsed</b><div><span>—</span><span>→</span><strong>{totals.lapse.toLocaleString("en-IN")}</strong><em className="attentionText">CURRENT</em></div><small>Latest snapshot metric</small></div></div>
    </div>
    <div className="section card">
      <div className="sectionHead"><div><h2>Advanced Filters & Search</h2><div className="note">All filters work together. Search covers PS, BLO, Supervisor, AERO and Designation.</div></div><button className="btn" onClick={resetFilters}>Reset Filters</button></div>
      <div className="filters">
        <select className="select" value={filterAero} onChange={e=>setFilterAero(e.target.value)}>{aeros.map(x=><option key={x}>{x}</option>)}</select>
        <select className="select" value={filterDesignation} onChange={e=>setFilterDesignation(e.target.value)}>{designations.map(x=><option key={x}>{x}</option>)}</select>
        <select className="select" value={filterSup} onChange={e=>setFilterSup(e.target.value)}>{sups.map(x=><option key={x}>{x}</option>)}</select>
        <select className="select" value={filterBlo} onChange={e=>setFilterBlo(e.target.value)}>{blos.map(x=><option key={x}>{x}</option>)}</select>
        <select className="select" value={filterHearing} onChange={e=>setFilterHearing(e.target.value)}><option>ALL</option><option>ZERO HEARING</option><option>PENDING</option><option>DISPOSED</option></select>
        <select className="select" value={filterZero} onChange={e=>setFilterZero(e.target.value)}><option>ALL</option><option>YES</option><option>NO</option></select>
        <select className="select" value={filterHeld} onChange={e=>setFilterHeld(e.target.value)}><option value="ALL">HEARING HELD: ALL</option><option value="YES">HEARING HELD: YES</option><option value="NO">HEARING HELD: NO</option></select>
        <select className="select" value={filterLapse} onChange={e=>setFilterLapse(e.target.value)}><option value="ALL">HEARING LAPSED: ALL</option><option value="YES">HEARING LAPSED: YES</option><option value="NO">HEARING LAPSED: NO</option></select>
        <input className="search" placeholder="Search PS / BLO / Supervisor / AERO / Designation" value={search} onChange={e=>setSearch(e.target.value)}/>
        <input className="select range" type="number" min="0" placeholder="Min Hearing %" value={minPct} onChange={e=>setMinPct(e.target.value)}/>
        <input className="select range" type="number" min="0" placeholder="Max Hearing %" value={maxPct} onChange={e=>setMaxPct(e.target.value)}/>
      </div>
    </div>
    <div className="section card"><h2>AERO / Ad.AERO Wise Consolidated Report</h2><div className="note" style={{marginBottom:8}}>🔴 Red = ZERO HEARING HELD. 🟢 Green = 20%+ hearings held of notice generated. 🟠 Orange = below 10%.</div><div className="tableWrap"><table className="table consolidated"><thead><tr>
      <th>S.No.</th><th>AERO / Ad.AERO</th><th>Designation</th><th>No. of PS</th><th>Notice Generated</th><th>Notice Deliv. Prev</th><th>Notice Deliv. Latest</th><th>Diff.</th><th>Hearing Held Prev</th><th>Hearing Held Latest</th><th>Diff.</th><th>% Held (of Gen.)</th><th>Hearing Lapse</th>
    </tr></thead><tbody>{officerSummary.map((x,i)=>{const rs=rows.filter(r=>r.aero===x.aero);const pd=rs.reduce((s,r)=>s+r.prevDelivered,0);const ph=rs.reduce((s,r)=>s+r.prevHearing,0);const lapse=rs.reduce((s,r)=>s+r.hearingLapse,0);return <tr key={x.aero} className={""}>
      <td>{i+1}</td><td>{x.aero}</td><td>{x.designation}</td><td className="num">{x.ps}</td><td className="num">{x.gen.toLocaleString()}</td><td className="num">{pd.toLocaleString()}</td><td className="num">{x.del.toLocaleString()}</td><td className="num">{(x.del-pd).toLocaleString()}</td><td className="num">{ph.toLocaleString()}</td><td className="num">{x.dh.toLocaleString()}</td><td className="num">{(x.dh-ph).toLocaleString()}</td><td className={"num "+(x.gen&&x.dh/x.gen<0.10?"attention":x.gen&&x.dh/x.gen>=0.20?"good":"")}>{x.gen?((x.dh/x.gen)*100).toFixed(2)+"%":"0.00%"}</td><td className="num">{lapse.toLocaleString()}</td>
    </tr>})}</tbody></table></div></div>
    <div className="section card">
      <div className="sectionHead"><div><h2>ZERO HEARING PS</h2><div className="note">All PS where Latest ECI Hearing Held = 0. Click a row for PS detail.</div></div><div className="buttonRow"><b className="criticalPill">{zeroRows.length} / {rows.length} PS</b><button className="btn" onClick={exportZeroExcel}>Export Zero Hearing Excel</button><button className="btn" onClick={()=>printReport(zeroRows.map(r=>({...r})))}>Download Zero Hearing PDF</button></div></div>
      <div className="tableWrap"><table className="table"><thead><tr><th>AERO / Ad.AERO</th><th>Designation</th><th>Supervisor</th><th>PS No.</th><th>BLO Name</th><th>Notice Generated</th><th>Notice Delivered</th><th>Hearing Held</th><th>Hearing Lapsed</th><th>Status</th></tr></thead><tbody>{zeroRows.map(r=><tr key={r.ps} className="zero clickable" onClick={()=>setSelectedPs(r.ps)}><td>{r.aero}</td><td>{r.designation}</td><td>{r.supervisor}</td><td>{r.ps}</td><td>{r.bloName}</td><td className="num">{r.noticeGenerated.toLocaleString("en-IN")}</td><td className="num">{r.latestDelivered.toLocaleString("en-IN")}</td><td className="num">0</td><td className="num">{r.hearingLapse.toLocaleString("en-IN")}</td><td><span className="zeroBadge">ZERO HEARING</span></td></tr>)}</tbody></table></div>
    </div>
    <div className="section card">
      <div className="sectionHead"><div><h2>ZERO HEARING — AERO / Ad.AERO WISE</h2><div className="note">Click an AERO row to filter the Zero Hearing module.</div></div></div>
      <div className="tableWrap"><table className="table"><thead><tr><th>S.No.</th><th>AERO / Ad.AERO</th><th>Total PS</th><th>Zero Hearing PS</th><th>Zero Hearing %</th></tr></thead><tbody>{officerSummary.map((x:any,i:number)=>{const z=rows.filter(r=>r.aero===x.aero&&r.latestHearing===0).length;return <tr key={x.aero} className={z?"zero clickable":""} onClick={()=>{if(z){setFilterAero(x.aero);setFilterZero("YES");}}}><td>{i+1}</td><td>{x.aero}</td><td>{x.ps}</td><td>{z}</td><td>{x.ps?(z/x.ps*100).toFixed(2)+"%":"0.00%"}</td></tr>})}<tr className="total"><td colSpan={2}>TOTAL AC-34</td><td>{rows.length}</td><td>{zeroRows.length}</td><td>{rows.length?(zeroRows.length/rows.length*100).toFixed(2)+"%":"0.00%"}</td></tr></tbody></table></div>
    </div>
    <div className="section card">
      <div className="sectionHead"><div><h2>Supervisor Dashboard</h2><div className="note">Click a supervisor to filter the PS-wise data below. Mapping is taken from Master Mapping.</div></div></div>
      <div className="tableWrap"><table className="table"><thead><tr><th>Supervisor</th><th>AERO / Ad.AERO</th><th>Total PS</th><th>Total BLOs</th><th>Notice Generated</th><th>Notice Delivered</th><th>Hearing Held</th><th>Hearing Pending</th><th>Hearing Lapsed</th><th>Zero Hearing PS</th><th>Hearing %</th><th>Disposal %</th></tr></thead><tbody>{supervisorSummary.map((x:any)=><tr key={x.name} className="clickable" onClick={()=>{setFilterSup(x.name);setFilterAero("ALL");}}><td>{x.name}</td><td>{x.aero}</td><td>{x.ps}</td><td>{x.blosCount}</td><td className="num">{x.gen.toLocaleString("en-IN")}</td><td className="num">{x.del.toLocaleString("en-IN")}</td><td className="num">{x.hearing.toLocaleString("en-IN")}</td><td className="num">{x.pending.toLocaleString("en-IN")}</td><td className="num">{x.lapse.toLocaleString("en-IN")}</td><td className={x.zero?"criticalText":"num"}>{x.zero}</td><td>{x.heldPct.toFixed(2)}%</td><td>{x.dispose.toFixed(2)}%</td></tr>)}</tbody></table></div>
    </div>
    <div className="section card">
      <div className="sectionHead"><div><h2>AERO → Supervisor → BLO → PS Drill Down</h2><div className="note">Current selection: {filterAero} → {filterSup} → {shown.length} PS.</div></div></div>
      <div className="tableWrap"><table className="table"><thead><tr><th>AERO</th><th>Supervisor</th><th>BLO Name</th><th>PS</th><th>Notice Generated</th><th>Delivered</th><th>Hearing Held</th><th>Lapsed</th><th>Pending</th><th>Hearing %</th><th>Disposal %</th></tr></thead><tbody>{shown.map(r=><tr key={"drill-"+r.ps} className="clickable" onClick={()=>setSelectedPs(r.ps)}><td>{r.aero}</td><td>{r.supervisor}</td><td>{r.bloName}</td><td>{r.ps}</td><td>{r.noticeGenerated.toLocaleString("en-IN")}</td><td>{r.latestDelivered.toLocaleString("en-IN")}</td><td>{r.latestHearing.toLocaleString("en-IN")}</td><td>{r.hearingLapse.toLocaleString("en-IN")}</td><td>{pendingFor(r).toLocaleString("en-IN")}</td><td>{r.noticeGenerated?(r.latestHearing/r.noticeGenerated*100).toFixed(2)+"%":"0.00%"}</td><td>{disposePct(r).toFixed(2)}%</td></tr>)}</tbody></table></div>
    </div>
    <div className="note" style={{marginBottom:8}}>🔴 Red rows = <b>Latest ECI Hearings Held = 0</b>. They remain highlighted on every refresh until the latest report shows at least one hearing held.</div>
    <div className="sectionHead"><div><h2>PS-wise Detailed Report</h2><div className="note">{shown.length} PS matching current filters. Click any PS for the detail view.</div></div><div className="columnBox">{Object.entries({ps:"PS",aero:"AERO",designation:"Designation",supervisor:"Supervisor",bloName:"BLO Name",bloMobile:"BLO Mobile",gen:"Notice Generated",pd:"Delivered Prev",ld:"Delivered Latest",diffDel:"Δ Delivered",ph:"Hearing Prev",lh:"Hearing Latest",diffHear:"Δ Hearing",pct:"% Held",lapse:"Hearing Lapsed"}).map(([k,label])=><label key={k}><input type="checkbox" checked={columns[k]} onChange={()=>setColumns(c=>({...c,[k]:!c[k]}))}/>{label}</label>)}</div></div>
    <div className="tableWrap"><table className="table detailUi"><thead><tr>
      {columns.ps&&<th onClick={()=>{setSortField("ps");setSortDir(d=>sortField==="ps"?(d==="asc"?"desc":"asc"):"asc")}}>PS ↕</th>}{columns.aero&&<th onClick={()=>{setSortField("aero");setSortDir(d=>sortField==="aero"?(d==="asc"?"desc":"asc"):"asc")}}>AERO ↕</th>}{columns.designation&&<th onClick={()=>{setSortField("designation");setSortDir(d=>sortField==="designation"?(d==="asc"?"desc":"asc"):"asc")}}>Designation ↕</th>}{columns.supervisor&&<th onClick={()=>{setSortField("supervisor");setSortDir(d=>sortField==="supervisor"?(d==="asc"?"desc":"asc"):"asc")}}>Supervisor ↕</th>}{columns.bloName&&<th onClick={()=>{setSortField("bloName");setSortDir(d=>sortField==="bloName"?(d==="asc"?"desc":"asc"):"asc")}}>BLO Name ↕</th>}{columns.bloMobile&&<th>BLO Mobile</th>}{columns.gen&&<th onClick={()=>{setSortField("noticeGenerated");setSortDir(d=>sortField==="noticeGenerated"?(d==="asc"?"desc":"asc"):"asc")}}>Notice Generated ↕</th>}{columns.pd&&<th onClick={()=>{setSortField("prevDelivered");setSortDir(d=>sortField==="prevDelivered"?(d==="asc"?"desc":"asc"):"asc")}}>Delivered Prev ↕</th>}{columns.ld&&<th onClick={()=>{setSortField("latestDelivered");setSortDir(d=>sortField==="latestDelivered"?(d==="asc"?"desc":"asc"):"asc")}}>Delivered Latest ↕</th>}{columns.diffDel&&<th>Δ Delivered</th>}{columns.ph&&<th onClick={()=>{setSortField("prevHearing");setSortDir(d=>sortField==="prevHearing"?(d==="asc"?"desc":"asc"):"asc")}}>Hearing Prev ↕</th>}{columns.lh&&<th onClick={()=>{setSortField("latestHearing");setSortDir(d=>sortField==="latestHearing"?(d==="asc"?"desc":"asc"):"asc")}}>Hearing Latest ↕</th>}{columns.diffHear&&<th>Δ Hearing</th>}{columns.pct&&<th>% Held</th>}{columns.lapse&&<th onClick={()=>{setSortField("hearingLapse");setSortDir(d=>sortField==="hearingLapse"?(d==="asc"?"desc":"asc"):"asc")}}>Hearing Lapsed ↕</th>}
    </tr></thead><tbody>{shown.map(r=><tr key={r.ps} className={r.latestHearing===0?"zero clickable":"clickable"} onClick={()=>setSelectedPs(r.ps)}>
      {columns.ps&&<td>{r.ps}</td>}{columns.aero&&<td>{r.aero}</td>}{columns.designation&&<td>{r.designation}</td>}{columns.supervisor&&<td>{r.supervisor}</td>}{columns.bloName&&<td>{r.bloName}</td>}{columns.bloMobile&&<td>{r.bloMobile}</td>}{columns.gen&&<td className="num">{r.noticeGenerated.toLocaleString("en-IN")}</td>}{columns.pd&&<td className="num">{r.prevDelivered.toLocaleString("en-IN")}</td>}{columns.ld&&<td className="num">{r.latestDelivered.toLocaleString("en-IN")}</td>}{columns.diffDel&&<td className="num pos">{(r.latestDelivered-r.prevDelivered>=0?"+":"")+(r.latestDelivered-r.prevDelivered).toLocaleString("en-IN")}</td>}{columns.ph&&<td className="num">{r.prevHearing.toLocaleString("en-IN")}</td>}{columns.lh&&<td className="num">{r.latestHearing.toLocaleString("en-IN")}{r.latestHearing===0&&<span className="zeroBadge"> ZERO</span>}</td>}{columns.diffHear&&<td className="num pos">{(r.latestHearing-r.prevHearing>=0?"+":"")+(r.latestHearing-r.prevHearing).toLocaleString("en-IN")}</td>}{columns.pct&&<td className={"num "+(r.noticeGenerated&&r.latestHearing/r.noticeGenerated<.10?"attention":r.noticeGenerated&&r.latestHearing/r.noticeGenerated>=.20?"good":"")}>{r.noticeGenerated?(r.latestHearing/r.noticeGenerated*100).toFixed(2)+"%":"0.00%"}</td>}{columns.lapse&&<td className="num">{r.hearingLapse.toLocaleString("en-IN")}</td>}
    </tr>)}</tbody></table></div>
    <div className="section card"><h2>How the update works</h2><div className="note">Previous ECI is the comparison baseline. Latest ECI replaces current metrics. Difference columns are Latest − Previous. Notice Generated is taken from Latest. BLO/Other data is joined by PS. The master mapping is not recalculated from the ECI report, so officer/BLO relationships stay fixed while daily numbers change.</div></div>
    {selectedPs!==null&&(()=>{const r=rows.find(x=>x.ps===selectedPs);if(!r)return null;return <div className="modalBack" onClick={()=>setSelectedPs(null)}><div className="modal" onClick={e=>e.stopPropagation()}><div className="sectionHead"><div><h2>PS {r.ps} — Detail View</h2><div className="note">Master Mapping + Previous ECI + Latest ECI</div></div><button className="btn" onClick={()=>setSelectedPs(null)}>Close</button></div><div className="detailGrid">{[["PS Number",r.ps],["AERO / Ad.AERO",r.aero],["Designation",r.designation],["Supervisor",r.supervisor],["BLO Name",r.bloName],["BLO Mobile",r.bloMobile],["Notice Generated",r.noticeGenerated.toLocaleString("en-IN")],["Notice Delivered Previous",r.prevDelivered.toLocaleString("en-IN")],["Notice Delivered Latest",r.latestDelivered.toLocaleString("en-IN")],["Delivery Difference",(r.latestDelivered-r.prevDelivered).toLocaleString("en-IN")],["Hearing Held Previous",r.prevHearing.toLocaleString("en-IN")],["Hearing Held Latest",r.latestHearing.toLocaleString("en-IN")],["Hearing Difference",(r.latestHearing-r.prevHearing).toLocaleString("en-IN")],["Hearing Lapsed",r.hearingLapse.toLocaleString("en-IN")],["Pending Hearing",pendingFor(r).toLocaleString("en-IN")],["Hearing Status",hearingStatus(r)],["Hearing %",r.noticeGenerated?(r.latestHearing/r.noticeGenerated*100).toFixed(2)+"%":"0.00%"],["Disposal %",disposePct(r).toFixed(2)+"%"]].map(([k,v])=><div className="detailItem" key={String(k)}><span>{k}</span><b>{String(v)}</b></div>)}</div></div></div>})()}
  </main>;
}
