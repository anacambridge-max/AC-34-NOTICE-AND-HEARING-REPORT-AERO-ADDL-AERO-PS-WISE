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
  // Snapshot the CURRENT dashboard state at the exact moment PDF export is clicked.
  // This prevents the print window from ever depending on a stale reference after an Excel refresh.
  const exportRows:Row[]=rows.map(r=>({...r}));
  const groups=buildGroups(exportRows);
  const grand=groups.reduce((g,z)=>({ps:g.ps+z.x.ps,gen:g.gen+z.x.gen,pd:g.pd+z.x.pd,ld:g.ld+z.x.ld,ph:g.ph+z.x.ph,lh:g.lh+z.x.lh,lapse:g.lapse+z.x.lapse,disc:g.disc+z.x.disc,docs:g.docs+z.x.docs,letter:g.letter+z.x.letter}),{ps:0,gen:0,pd:0,ld:0,ph:0,lh:0,lapse:0,disc:0,docs:0,letter:0});
  if(sourceTotals?.totalDocs!=null) grand.docs=sourceTotals.totalDocs;
  if(sourceTotals?.totalDisc!=null) grand.disc=sourceTotals.totalDisc;
  if(sourceTotals?.totalLetter!=null) grand.letter=sourceTotals.totalLetter;
  let h='<!doctype html><html><head><title>AC-34 MATIALA - NOTICE & HEARING REPORT - CURRENT DATA</title><style>';
  h+='@page{size:A4 landscape !important;margin:4mm 4mm 5mm 4mm}*{box-sizing:border-box;-webkit-print-color-adjust:exact!important;print-color-adjust:exact!important}html{width:100%;background:#fff}body{font-family:Arial,Helvetica,sans-serif;color:#111;background:#fff;margin:0;width:100%;font-size:9px;-webkit-print-color-adjust:exact!important;print-color-adjust:exact!important}thead{display:table-header-group}tbody{display:table-row-group}@media print{html,body{width:100%;height:auto;background:#fff;-webkit-print-color-adjust:exact!important;print-color-adjust:exact!important}.reportPage{page-break-before:always;break-before:page}.summaryEnd{page-break-after:always;break-after:page}}';
  h+='h1{text-align:center;font-size:13px;margin:0 0 1px}h2{text-align:center;font-size:8px;margin:0 0 3px}.overall{border:1px solid #5c6f55;background:#e7f2df;text-align:center;padding:3px;font-weight:700;margin:3px 0 4px}.detailTitle{border:1px solid #5c6f55;background:#e7f2df;text-align:center;padding:3px;font-weight:800;margin:2px 0 2px}.detailStats{border:1px solid #5c6f55;background:#e7f2df;text-align:center;padding:3px;font-weight:700;margin:0 0 4px}.reportPage{page-break-before:always;break-before:page}.officer{font-size:9px;font-weight:700;margin:3px 0 2px}';
  h+='table{border-collapse:collapse;width:100%;table-layout:fixed;margin:0 0 3px;page-break-inside:auto}tr{page-break-inside:avoid}th,td{border:1px solid #333;padding:2px 2.5px;text-align:center;vertical-align:middle;overflow:hidden;font-size:10.5px;line-height:1.08}th{font-size:8.5px;line-height:1.08;font-weight:800}.num{white-space:nowrap}.summaryTable th:nth-child(1),.summaryTable td:nth-child(1){width:3%}.summaryTable th:nth-child(2),.summaryTable td:nth-child(2){width:12%}.summaryTable th:nth-child(3),.summaryTable td:nth-child(3){width:7%}.summaryTable th:nth-child(4),.summaryTable td:nth-child(4){width:5%}.summaryTable th:nth-child(5),.summaryTable td:nth-child(5){width:8%}.summaryTable th:nth-child(6),.summaryTable td:nth-child(6),.summaryTable th:nth-child(7),.summaryTable td:nth-child(7){width:7%}.summaryTable th:nth-child(8),.summaryTable td:nth-child(8){width:5%}.summaryTable th:nth-child(9),.summaryTable td:nth-child(9),.summaryTable th:nth-child(10),.summaryTable td:nth-child(10){width:7%}.summaryTable th:nth-child(11),.summaryTable td:nth-child(11){width:5%}.summaryTable th:nth-child(12),.summaryTable td:nth-child(12){width:6%}.summaryTable th:nth-child(13),.summaryTable td:nth-child(13){width:6%}.summaryTable th:nth-child(14),.summaryTable td:nth-child(14){width:7%}.summaryTable th:nth-child(15),.summaryTable td:nth-child(15){width:8%}.summaryTable th:nth-child(16),.summaryTable td:nth-child(16){width:7%}.detailTable th:nth-child(1),.detailTable td:nth-child(1){width:3%}.detailTable th:nth-child(2),.detailTable td:nth-child(2){width:4%}.detailTable th:nth-child(3),.detailTable td:nth-child(3){width:12%}.detailTable th:nth-child(4),.detailTable td:nth-child(4){width:8%}.detailTable th:nth-child(5),.detailTable td:nth-child(5){width:9%}.detailTable th:nth-child(6),.detailTable td:nth-child(6){width:7%}.detailTable th:nth-child(7),.detailTable td:nth-child(7),.detailTable th:nth-child(8),.detailTable td:nth-child(8){width:6%}.detailTable th:nth-child(9),.detailTable td:nth-child(9){width:4%}.detailTable th:nth-child(10),.detailTable td:nth-child(10),.detailTable th:nth-child(11),.detailTable td:nth-child(11){width:6%}.detailTable th:nth-child(12),.detailTable td:nth-child(12){width:4%}.detailTable th:nth-child(13),.detailTable td:nth-child(13){width:6%}.detailTable th:nth-child(14),.detailTable td:nth-child(14){width:6%}.detailTable th:nth-child(15),.detailTable td:nth-child(15){width:6%}.detailTable th:nth-child(16),.detailTable td:nth-child(16){width:7%}.detailTable th:nth-child(17),.detailTable td:nth-child(17){width:6%}.detailTable th,.detailTable td{font-size:9.5px;padding:1.8px 1.8px;line-height:1.05;overflow-wrap:anywhere}.detailTable td:nth-child(3),.detailTable td:nth-child(4),.detailTable td:nth-child(5){white-space:normal;word-break:break-word}.detailTable th{font-size:8.2px}.summaryTable th:nth-child(12),.summaryTable td:nth-child(12),.detailTable th:nth-child(13),.detailTable td:nth-child(13){background:#dcebd5!important}.officer{page-break-inside:avoid;break-inside:avoid-page;page-break-after:avoid;break-after:avoid-page}.name{text-align:left}.num{text-align:right}';
  h+='.dark{background:#333;color:#fff}.blueHead{background:#0e4d7e;color:#fff}.greenHead{background:#dcebd5}.orangeHead{background:#9a5a00;color:#fff}.under{background:#b9d9ff!important;color:#063b73!important;font-weight:800}.underBadge{background:#2f73b8;color:#fff;padding:1px 3px;border-radius:3px;font-size:6.5px;font-weight:800;white-space:nowrap}.zero{background:#f4b7bd!important;color:#8b0000!important;font-weight:800}.good{background:#dcebd5!important;color:#14532d!important;font-weight:800}.attention{background:#ffe0b2!important;color:#8a4b08!important;font-weight:800}.zeroBadge{background:#9b0000;color:#fff;padding:1px 3px;border-radius:3px;font-size:6.5px;font-weight:800;white-space:nowrap}.total{background:#fff200!important;font-weight:800} .small{font-size:7px}.summaryEnd{page-break-after:always;break-after:page}</style></head><body>';
  h+='<h1>OFFICE OF THE ELECTORAL REGISTRATION OFFICER, AC-34, MATIALA</h1>';
  h+='<h2>SIR-2026 - PS-WISE (AERO-WISE) NOTICE & HEARING COMPARISON - Previous ECI vs Latest ECI</h2>';
  h+='<div class="overall">OVERALL AC-34 MATIALA - Notice Generated: '+n(grand.gen)+' | Notices Delivered: +'+n(grand.ld-grand.pd)+' | Hearings Held: +'+n(grand.lh-grand.ph)+' | % Held: '+pct(grand.lh,grand.gen)+'</div>';
  h+='<h2>AERO/Ad.AERO-wise Summary</h2><table class="summaryTable"><thead><tr><th class="dark">S.No.</th><th class="dark">AERO / Ad.AERO</th><th class="dark">Designation</th><th class="dark">No. of PS</th><th class="dark">Notice Generated</th><th class="blueHead">Notice Deliv.<br>Prev</th><th class="blueHead">Notice Deliv.<br>Latest</th><th class="blueHead">Diff.</th><th class="blueHead">Hearing Held<br>Prev</th><th class="blueHead">Hearing Held<br>Latest</th><th class="blueHead">Diff.</th><th class="greenHead">% Held<br>(of Gen.)</th><th class="dark">Hearing Lapse</th></tr>';
  groups.forEach((g,i)=>{
    h+='<tr class="'+''+'"><td>'+String(i+1)+'</td><td class="name">'+esc(g.aero)+''+'</td><td>'+esc(g.rs[0]?.designation||'')+'</td><td>'+n(g.x.ps)+'</td><td class="num">'+n(g.x.gen)+'</td><td class="num">'+n(g.x.pd)+'</td><td class="num">'+n(g.x.ld)+'</td><td class="num">'+n(g.x.ld-g.x.pd)+'</td><td class="num">'+n(g.x.ph)+'</td><td class="num">'+n(g.x.lh)+'</td><td class="num">'+n(g.x.lh-g.x.ph)+'</td><td class="num '+(g.x.gen&&g.x.lh/g.x.gen<0.10?'attention':g.x.gen&&g.x.lh/g.x.gen>=0.20?'good':'')+'">'+pct(g.x.lh,g.x.gen)+'</td><td class="num">'+n(g.x.lapse)+'</td></tr>';
  });
  h+='<tr class="total"><td colspan="3">GRAND TOTAL - AC-34 MATIALA</td><td>'+n(grand.ps)+'</td><td>'+n(grand.gen)+'</td><td>'+n(grand.pd)+'</td><td>'+n(grand.ld)+'</td><td>'+n(grand.ld-grand.pd)+'</td><td>'+n(grand.ph)+'</td><td>'+n(grand.lh)+'</td><td>'+n(grand.lh-grand.ph)+'</td><td>'+pct(grand.lh,grand.gen)+'</td><td>'+n(grand.lapse)+'</td></tr></table>';
  const zeroGroups=groups.map(g=>({aero:g.aero,ps:g.x.ps,zero:g.rs.filter(r=>r.latestHearing===0).length})).filter(x=>x.zero>0);
  h+='<h2 style="margin-top:6px">ZERO HEARING HELD - AERO/Ad.AERO WISE</h2><table class="summaryTable"><thead><tr><th class="dark">S.No.</th><th class="dark">AERO / Ad.AERO</th><th class="dark">Total PS</th><th class="dark">Zero Hearing PS</th><th class="dark">Zero Hearing %</th></tr></thead><tbody>';
  zeroGroups.forEach((z,i)=>{h+='<tr class="zero"><td>'+String(i+1)+'</td><td class="name">'+esc(z.aero)+'</td><td>'+n(z.ps)+'</td><td>'+n(z.zero)+'</td><td>'+pct(z.zero,z.ps)+'</td></tr>';});
  h+='<tr class="total"><td colspan="2">TOTAL</td><td>'+n(grand.ps)+'</td><td>'+n(exportRows.filter(r=>r.latestHearing===0).length)+'</td><td>'+pct(exportRows.filter(r=>r.latestHearing===0).length,grand.ps)+'</td></tr></tbody></table>';
  h+='<div class="small">Report generated from CURRENT dashboard data at '+new Date().toLocaleString("en-IN")+'. Red rows = ZERO HEARING HELD.</div><div class="summaryEnd"></div>';

  groups.forEach((g,gi)=>{
    h+='<div class="reportPage"><h1>OFFICE OF THE ELECTORAL REGISTRATION OFFICER, AC-34, MATIALA</h1><h2>SIR-2026 - PS-WISE (AERO-WISE) NOTICE & HEARING COMPARISON - Previous ECI vs Latest ECI</h2><div class="detailTitle">'+String(gi+1)+'. '+esc(g.aero)+', '+esc(g.rs[0]?.designation||'')+' (Mobile: '+esc(g.rs[0]?.aeroMobile||'')+') - '+g.x.ps+' PS | Notice Generated: '+n(g.x.gen)+'</div><div class="detailStats">Notices Delivered: +'+n(g.x.ld-g.x.pd)+' (from '+n(g.x.pd)+' to '+n(g.x.ld)+') = '+pct(g.x.ld-g.x.pd,g.x.gen)+' of Notice Generated | Hearings Held: +'+n(g.x.lh-g.x.ph)+' (from '+n(g.x.ph)+' to '+n(g.x.lh)+') = '+pct(g.x.lh-g.x.ph,g.x.gen)+' of Notice Generated'+''+'</div>';
    h+='<table class="detailTable"><thead><tr><th class="dark">S.No.</th><th class="dark">P.S. No.</th><th class="dark">BLO Name</th><th class="dark">BLO Mobile No.</th><th class="dark">BLO Supervisor</th><th class="dark">Notice<br>Generated</th><th class="blueHead">Notice Deliv.<br>Prev</th><th class="blueHead">Notice Deliv.<br>Latest</th><th class="blueHead">Diff.</th><th class="blueHead">Hearing Held<br>Prev</th><th class="blueHead">Hearing Held<br>Latest</th><th class="blueHead">Diff.</th><th class="greenHead">% Held<br>(of Gen.)</th><th class="dark">Hearing Lapse</th></tr>';
    g.rs.slice().sort((a,b)=>a.ps-b.ps).forEach((r,i)=>{
      const isZero=r.latestHearing===0;
      h+='<tr class="'+(isZero?'zero':'')+'"><td>'+String(i+1)+'</td><td>'+r.ps+'</td><td class="name">'+esc(r.bloName)+(isZero?' <span class="zeroBadge">ZERO HEARING HELD</span>':'')+'</td><td>'+esc(r.bloMobile)+'</td><td class="name">'+esc(r.supervisor)+'</td><td>'+n(r.noticeGenerated)+'</td><td>'+n(r.prevDelivered)+'</td><td>'+n(r.latestDelivered)+'</td><td>'+n(r.latestDelivered-r.prevDelivered)+'</td><td>'+n(r.prevHearing)+'</td><td>'+n(r.latestHearing)+'</td><td>'+n(r.latestHearing-r.prevHearing)+'</td><td>'+pct(r.latestHearing,r.noticeGenerated)+'</td><td>'+n(r.hearingLapse)+'</td></tr>';
    });
    h+='<tr class="total"><td colspan="5">Total - '+esc(g.aero)+', '+esc(g.rs[0]?.designation||'')+' ('+g.x.ps+' PS)</td><td>'+n(g.x.gen)+'</td><td>'+n(g.x.pd)+'</td><td>'+n(g.x.ld)+'</td><td>'+n(g.x.ld-g.x.pd)+'</td><td>'+n(g.x.ph)+'</td><td>'+n(g.x.lh)+'</td><td>'+n(g.x.lh-g.x.ph)+'</td><td>'+pct(g.x.lh,g.x.gen)+'</td><td>'+n(g.x.lapse)+'</td></tr></table></div>';
  });
  h+='</body></html>';
  const w=window.open("","_blank","width=1800,height=1200"); if(!w){alert("Please allow pop-ups for PDF export.");return;} w.document.write(h); w.document.close(); w.focus(); setTimeout(()=>{w.focus();w.print();},1200);
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

  useEffect(()=>{try{const saved=localStorage.getItem("ac34-master-mapping");if(saved){const arr=JSON.parse(saved);setMaster(new Map(arr));}}catch{}},[]);
  useEffect(()=>{if(master.size)localStorage.setItem("ac34-master-mapping",JSON.stringify(Array.from(master.entries())));},[master]);

  const load=(setter:any,kind:"eci"|"blo")=>(e:React.ChangeEvent<HTMLInputElement>)=>{const f=e.target.files?.[0];if(!f)return;setError("");readRows(f).then(p=>{const parsed:any=kind==="eci"?parseMetricFile(p,"eci"):parseMetricFile(p,"blo");if(kind==="blo" && parsed.acTotals)setSourceTotals(parsed.acTotals);setter(parsed);}).catch(err=>setError(String(err)));};

  useEffect(()=>{if(master.size) setRows(merge(master,previous,latest,blo));},[master,previous,latest,blo]);

  const aeros=useMemo(()=>["ALL",...Array.from(new Set(rows.map(r=>r.aero).filter(Boolean)))],[rows]);
  const sups=useMemo(()=>["ALL",...Array.from(new Set(rows.map(r=>r.supervisor).filter(Boolean)))],[rows]);
  const shown=useMemo(()=>rows.filter(r=>(filterAero==="ALL"||r.aero===filterAero)&&(filterSup==="ALL"||r.supervisor===filterSup)&&((r.ps+" "+r.bloName+" "+r.supervisor+" "+r.aero).toLowerCase().includes(search.toLowerCase()))),[rows,filterAero,filterSup,search]);
  const totals=useMemo(()=>rows.reduce((a,r)=>({gen:a.gen+r.noticeGenerated,del:a.del+r.latestDelivered,ph:a.ph+r.prevHearing,lh:a.lh+r.latestHearing,zero:a.zero+(r.latestHearing===0?1:0)}),{gen:0,del:0,ph:0,lh:0,zero:0}),[rows]);
  const officerSummary=useMemo(()=>{const m=new Map<string,any>();for(const r of rows){const k=r.aero||"Unmapped";const x=m.get(k)||{aero:k,designation:r.designation,ps:0,gen:0,del:0,dh:0,zero:0};x.ps++;x.gen+=r.noticeGenerated;x.del+=r.latestDelivered;x.dh+=r.latestHearing;x.zero+=r.latestHearing===0?1:0;m.set(k,x);}return Array.from(m.values()).sort((a,b)=>reportAeroIndex(a.aero)-reportAeroIndex(b.aero));},[rows]);
  const supervisorSummary=useMemo(()=>{const m=new Map<string,any>();for(const r of rows){const k=r.supervisor||"Unmapped";const x=m.get(k)||{name:k,ps:0,hearing:0,zero:0};x.ps++;x.hearing+=r.latestHearing;x.zero+=r.latestHearing===0?1:0;m.set(k,x);}return Array.from(m.values()).sort((a,b)=>a.name.localeCompare(b.name));},[rows]);

  const exportCsv=()=>{const data=shown.map(r=>({PS:r.ps,AERO:r.aero,Designation:r.designation,"BLO Supervisor":r.supervisor,"BLO Name":r.bloName,"BLO Mobile":r.bloMobile,"Notice Generated":r.noticeGenerated,"Delivered Previous":r.prevDelivered,"Delivered Latest":r.latestDelivered,"Delivered Difference":r.latestDelivered-r.prevDelivered,"Hearing Previous":r.prevHearing,"Hearing Latest":r.latestHearing,"Hearing Difference":r.latestHearing-r.prevHearing,"Hearing Lapse":r.hearingLapse,"Discrepancy Delivered":r.discrepancyDelivered,"BLO Docs":r.bloDocs,"BLO Letter":r.bloLetter}));const ws=XLSX.utils.json_to_sheet(data);const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,ws,"Dashboard");XLSX.writeFile(wb,"AC34_Dashboard_Export.xlsx");};

  return <main className="wrap">
    <div className="top"><div><div className="title">AC-34 MATIALA — NOTICE & HEARING DASHBOARD</div><div className="sub">SIR-2026 • Fixed PS → AERO/Ad.AERO → BLO Supervisor → BLO mapping • Latest ECI comparison</div></div><div className="badge">430 PS MASTER STRUCTURE</div></div>
    <div className="grid">
      <div className="card"><h3>① Previous ECI Excel</h3><p>Baseline report used for comparison.</p><input className="input" type="file" accept=".xlsx,.xls,.csv" onChange={load(setPrevious,"eci")}/></div>
      <div className="card"><h3>② Latest ECI Excel</h3><p>Current report. Notice Generated + latest hearing status come from this file.</p><input className="input" type="file" accept=".xlsx,.xls,.csv" onChange={load(setLatest,"eci")}/></div>
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
      <button className="btn primary" onClick={()=>printReport(rows.map(r=>({...r})))}>Download PDF</button>
    </div>
    <div className="section card"><h2>AERO / Ad.AERO Wise Consolidated Report</h2><div className="note" style={{marginBottom:8}}>🔴 Red = ZERO HEARING HELD. 🟢 Green = 20%+ hearings held of notice generated. 🟠 Orange = below 10%.</div><div className="tableWrap"><table className="table consolidated"><thead><tr>
      <th>S.No.</th><th>AERO / Ad.AERO</th><th>Designation</th><th>No. of PS</th><th>Notice Generated</th><th>Notice Deliv. Prev</th><th>Notice Deliv. Latest</th><th>Diff.</th><th>Hearing Held Prev</th><th>Hearing Held Latest</th><th>Diff.</th><th>% Held (of Gen.)</th><th>Hearing Lapse</th>
    </tr></thead><tbody>{officerSummary.map((x,i)=>{const rs=rows.filter(r=>r.aero===x.aero);const pd=rs.reduce((s,r)=>s+r.prevDelivered,0);const ph=rs.reduce((s,r)=>s+r.prevHearing,0);const lapse=rs.reduce((s,r)=>s+r.hearingLapse,0);return <tr key={x.aero} className={""}>
      <td>{i+1}</td><td>{x.aero}</td><td>{x.designation}</td><td className="num">{x.ps}</td><td className="num">{x.gen.toLocaleString()}</td><td className="num">{pd.toLocaleString()}</td><td className="num">{x.del.toLocaleString()}</td><td className="num">{(x.del-pd).toLocaleString()}</td><td className="num">{ph.toLocaleString()}</td><td className="num">{x.dh.toLocaleString()}</td><td className="num">{(x.dh-ph).toLocaleString()}</td><td className={"num "+(x.gen&&x.dh/x.gen<0.10?"attention":x.gen&&x.dh/x.gen>=0.20?"good":"")}>{x.gen?((x.dh/x.gen)*100).toFixed(2)+"%":"0.00%"}</td><td className="num">{lapse.toLocaleString()}</td>
    </tr>})}</tbody></table></div></div>
    <div className="section card"><h2>ZERO HEARING HELD — AERO / Ad.AERO WISE</h2><div className="tableWrap"><table className="table"><thead><tr><th>S.No.</th><th>AERO / Ad.AERO</th><th>Total PS</th><th>Zero Hearing PS</th><th>Zero Hearing %</th></tr></thead><tbody>{officerSummary.map((x,i)=>{const rs=rows.filter(r=>r.aero===x.aero);const z=rs.filter(r=>r.latestHearing===0).length;return z>0?<tr key={x.aero} className="zero"><td>{i+1}</td><td>{x.aero}</td><td>{x.ps}</td><td>{z}</td><td>{x.ps?((z/x.ps)*100).toFixed(2)+"%":"0.00%"}</td></tr>:null;})}<tr className="total"><td colSpan={2}><b>TOTAL AC-34</b></td><td><b>{rows.length}</b></td><td><b>{rows.filter(r=>r.latestHearing===0).length}</b></td><td><b>{rows.length?((rows.filter(r=>r.latestHearing===0).length/rows.length)*100).toFixed(2)+"%":"0.00%"}</b></td></tr></tbody></table></div></div>
    <div className="note" style={{marginBottom:8}}>🔴 Red rows = <b>Latest ECI Hearings Held = 0</b>. They remain highlighted on every refresh until the latest report shows at least one hearing held.</div>
    <div className="tableWrap"><table className="table"><thead><tr>
      {["PS","AERO / Ad.AERO","BLO Supervisor","BLO Name","BLO Mobile","Notice Generated","Delivered Prev","Delivered Latest","Δ Delivered","Hearing Prev","Hearing Latest","Δ Hearing","% Held (of Gen.)","Hearing Lapse"].map(h=><th key={h}>{h}</th>)}
    </tr></thead><tbody>{shown.map(r=><tr key={r.ps} className={r.latestHearing===0?"zero":""}>
      <td>{r.ps}</td><td>{r.aero}</td><td>{r.supervisor}</td><td>{r.bloName}</td><td>{r.bloMobile}</td><td className="num">{r.noticeGenerated.toLocaleString()}</td><td className="num">{r.prevDelivered.toLocaleString()}</td><td className="num">{r.latestDelivered.toLocaleString()}</td><td className="num">{(r.latestDelivered-r.prevDelivered).toLocaleString()}</td><td className="num">{r.prevHearing.toLocaleString()}</td><td className="num">{r.latestHearing.toLocaleString()}{r.latestHearing===0&&<span className="zeroBadge"> ZERO HEARING HELD</span>}</td><td className="num">{(r.latestHearing-r.prevHearing).toLocaleString()}</td><td className={"num "+(r.noticeGenerated&&r.latestHearing/r.noticeGenerated<0.10?"attention":r.noticeGenerated&&r.latestHearing/r.noticeGenerated>=0.20?"good":"")}>{r.noticeGenerated?((r.latestHearing/r.noticeGenerated)*100).toFixed(2)+"%":"0.00%"}</td><td className="num">{r.hearingLapse.toLocaleString()}</td>
    </tr>)}</tbody></table></div>
    <div className="section card"><h2>How the update works</h2><div className="note">Previous ECI is the comparison baseline. Latest ECI replaces current metrics. Difference columns are Latest − Previous. Notice Generated is taken from Latest. BLO/Other data is joined by PS. The master mapping is not recalculated from the ECI report, so officer/BLO relationships stay fixed while daily numbers change.</div></div>
  </main>;
}
