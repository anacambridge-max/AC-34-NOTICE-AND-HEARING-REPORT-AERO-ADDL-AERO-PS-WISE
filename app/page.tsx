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
  const exportRows:Row[]=rows.map(r=>({...r}));
  const groups=buildGroups(exportRows);
  const grand=groups.reduce((g,z)=>({ps:g.ps+z.x.ps,gen:g.gen+z.x.gen,pd:g.pd+z.x.pd,ld:g.ld+z.x.ld,ph:g.ph+z.x.ph,lh:g.lh+z.x.lh,lapse:g.lapse+z.x.lapse,disc:g.disc+z.x.disc,docs:g.docs+z.x.docs,letter:g.letter+z.x.letter}),{ps:0,gen:0,pd:0,ld:0,ph:0,lh:0,lapse:0,disc:0,docs:0,letter:0});
  if(sourceTotals?.totalDocs!=null) grand.docs=sourceTotals.totalDocs;
  if(sourceTotals?.totalDisc!=null) grand.disc=sourceTotals.totalDisc;
  if(sourceTotals?.totalLetter!=null) grand.letter=sourceTotals.totalLetter;
  const now=new Date();
  const stamp=now.toLocaleString("en-IN",{day:"2-digit",month:"2-digit",year:"numeric",hour:"2-digit",minute:"2-digit",second:"2-digit",hour12:false});
  const zeroRows=exportRows.filter(r=>r.latestHearing===0).sort((a,b)=>a.ps-b.ps);
  const zeroPct=grand.ps?zeroRows.length/grand.ps*100:0;
  const hearingPct=grand.gen?grand.lh/grand.gen*100:0;
  const reviewPsByOfficer=new Map<string,Set<number>>();
  groups.forEach(g=>{
    const selected=g.rs.slice().sort((a,b)=>{
      const az=a.latestHearing===0?0:1, bz=b.latestHearing===0?0:1;
      return az-bz || a.ps-b.ps;
    }).slice(0,3);
    reviewPsByOfficer.set(g.aero,new Set(selected.map(r=>r.ps)));
  });
  const EMBLEM_DATA="data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAC0AAAA8CAIAAAB98qTzAAAScklEQVR42pVZd3gTV7a/586ouUm2ZdmSkY2FjSu2MTbY1IDpLQ41BEJCQkgCWRJesg942X152c3u+7aFt8tjv002ixMCSWBDIBA6oQRjm+4CuMi9y7Lc1DWae98fKpZkk+y7nyzPaGbuPXPK7/zOufBdbd+y5EhKKMJAKUKUIkCIIooQpRT5DM8ppQgBQggBgPcb+d/7/x4sxq4pPYt5DgBR5LMARSPrgM8/8B7Cj65DvV+Bw/Ucy4Dv3NR1CBghhKlLCYAopUCBYtdk2DMBBYQwQggwIEp/RBaXjl0KpZ41XN8UYUAIAcu4FBI4BXhuG/kBEAQoxMcuAD+hiwBlesyJ3EbGAgZ+0njgMQH4TeSW7l8cbvv6uxwCt5QsA/CvTOFelVLwFW/kWeqedSzdBijapQG3t7tnRhiPJQd1D6/QCFGAEWX6LRUQVL6vTCklhPjpgSKg1C22Z37q8rOA2egoZSK3myEfS4CPar130ieoc0Q0AI8eKHJ/EEIUsXSsh6lXa76+5BUFIPCCRzwPlnhCDcDXHwGAEAI+5vW+DSY/hT/0R5QEfnEAoyw74keA/Vwe3B+XohFCLPHcTQPg023DkQimnpVdr0g9BqZ+cejjwwCUUoyxtqEBY5igmUAIAXCr2Xu/ax6W+kpAnxz3XuyBAOAdbdCRNVyG+HDfPt7h+NtHH2GM3W7rp0igiLIuB6eIus7BFxyoL3hT6hXC4zm+mne9DCXEdYoxRgi1tDR//fXxi+fOcRw5cGD/po2bZeEyQghQoEC9DyKE0JXmASdPOJ538ryTdw9CeN714QnPE8IT30Gp+4D3H06es9vtVqvVbrdTSowm4+JFS0NDQqZPz8/OmQwYvb3rHVckcxxnt9vtdruTc/I8TwhB3zf1c245yIgc7r9AOSgllNIAOZye4XDYLRaL0Wg0m808z9fV1UkjI377wQft7R2PH9ds374jIWG81W6nlDocDrPZbLFYOAfnkoP1pNMA3w9IoQEpBEahl8ssbnNQBIQQWbhUEa2KjlHGxETL5RFr1q9ta29xGQ5jLBQIEbg9iRDCUrerwGj0/XGEpl5CQAhmGIwxIRSBUygSusIqIlw+JTP788OHe3W9ut7ue3fuzVswXyIWI4QYhkHMyFQcx7HIi7Ij6Wwki/lK5UkFI0HhUgMrEFhtNr1eX19bp9frRWKRShmbkpocFha6tGjJz17fXnb7DnE6GQZe2/4zhBChpPrhw9bWNs5mT0lL1SQkSCQS1g81YBSJoH60xzdAXKcMy3R2dn519OubJTcEDB4Y6L9+9QepVLrt9Vde3/GazWxOT5sYE6vSPm6OUkSu3bjGarOdv3Bu7549w/1DM2fPFgWJU1JSiopWoovaPo4nLrflRyLCNQglhFLq/cHrnu6rlDY3N8+dPTc0SLZ80YqW1tY+g27btq0Txk+ICJEdPFj8ytbX4xMmLJy/QK2MW7Z06YOqit//7verip5Rq+N27HiDUqrVamfMmpGYkIRd2Aq+sAz+YO0PZy5RnLwTITCZTS+/8krlg4fLi1bs3LVDKBJwdmduXvbyZUuk0kixWLT22dXSUJnJaN/9y59HR8csfWrJjevX9r63RxOvMRnNBoMhTh33/i/f12g0LEIUwE1GvahKfVIH9U1HHq0ABkrpn/fta29t/eB377c2twjDBPXax831baePnn1p56aGFm1IuHjmjLzcqVk9Xb1FzyzRG/Qz5sz8zR9+a7IOKuSRfXrdnr27lVGxq9at+PTwx8zmnbsTIiQB6dydBahHJx5v8ArHMmxt7eNf/MfenW/tXLtx1eGDXz2qrLGYjAyDL56/GqmS1VTU6HWDQ0P9Vy9e09bV37l1r0Xb+NGnB+pq6p9f9yJgtP75VXfLqx5U3h+niUlKSGR9Euooigq+GY96BWQYBjB8ffKkzUnNQ8aykrLElITMyVk/XLyWvCyu6LmFx7/8TiwSX7l67dCnh9XjY4LCgm6U3FLHqi6cOXv5+6sIYYFIgAVo2bJFkeqoeyWVTdV/x+4goWMmfA8tG6s44Tknw+J7dys+/+gwRY6ERNXQwFBtTRPn4BprW8OkobML8xMmaCTBId1tOokkqNeg/+BXf7hTXpGcHm8cNu//48GL568RCy8AUeyEBOaFnf+eECEB/6AFt9+OTRld6Ts+Lr60tIQVMLlTc1sbu+wWTiYL+/zgUZvJEhIiSUnRbHl1U2N9U1hwaOy4eJ44wmVhvIOKg4LiE1UG3YDZbGeFwue2rE7LTmIYjBEAjF6SjsG1XcDlzXaaCZqs7Oy+Pr0sKlShUKRnZaTnpfQPDEbHK7PyU3rbe8ou3L92vvTZdasOHz2wbdvmKZnZ8vAI87D18QOtgMWxagUG6NX3N9Y0drW0Y/AlmSPO4kuUPabzsD2XNJRSm9liNVu6WrsnT89SjYt5cKtyvEbV191//eztlOwMRbxCpYqKlIV3NnXe/aE8YXzMH//3V2/tfpURiAkIpBJJVnbK4b8eO3boW7FEhBEE5hbqz8foGCWum3tihu3WGx5V15tNw+Yh073S6tmFBf36AZ1hMFIh54h944vrNMlJxz47de3qbblCIZUGhYiFZoOxt2swJSNl9calQWJ2Uk5G3eMG5sW3dieES55ECMcqDNwZkGGw2Ww6depUXm5OctrE0OBQbW1dzowMntiiIuU8IRdPX82ZmpmZleGgfFCwhHei4r9/oa2rkasiZBERSakT6ysaI6LClqwuPHPyLPPim7vHh0vAr+Rz52Bfxuk1ituLATCAOl5d/Fmxw8YJcVB4uNRiMd64UBanjEmeqK59pH38UNvc1NxQ1/LsxlVRyogjn311r7x6ZuHUqNiojtqulLRsh50QSojdceb0eUyRv4O4JfDj4v6pEAAAAxBCQoPDZs6Y0djU0GfQRyhk+bPyW9t6wqTRddXNlberJibFDg0Nl5SWWc3m3h6d0TKYWzBpeMhRdvleRm7GoqKClc8WdrXoznz7PQ8U07HL4LH7Az4qch9sWL+xf2Cw7GbZ3dv34+Li5i2cPe2paYnp6dJw6eafPRerUvX09DY0td8pq9I+arY6zHfLy41me9H6FbKQ4KrblYpxMk1q4qO6WnZUueDHXj31j9snfO51m2nu3Dlpk9I72jpOfftdsmbiy9ueV6pigsWiuppHwz2m/IKciKiIPn2v3WmJipEP9w/PKpyRnz9rqHfo1P2HPZ09i58uPP7NaafdAZca9IWaSG/55n1rF/12+YNfJYpcAOuWn2XZs+fObX1lq0QkTktMee2NrbHq2Ae3KstKblIOTV8wlRULdG2GgQEDK8Yh4vDgkCDDsP7y2SsmszkyPDw1PfX46VOIAlzU6gsnRHpTrKfc8utBjVTcniqWEkIIRQixDAMYXtiy5drVq8ECsVQWJmBFDrtDJguLT9I47ObhQeNA33CwVOJw2ABwW3O7xW42mSysUBglk3MOrkvXfyzpEZYSv5YRBQS+fgkBhAQodd3gbYtRDMyLmzdfvnRhxeo1c58qrK6pMRgN1y9cr6qo4mwOm8Nm0BskUvFwn1Emk9ocVp4neVOm7dy1I0gSsmHjhsVLVjy9cglL/ZXusgGMnfvAqyuMgXpaUZTS9Iy0qCiFWh23dOXyxSsW2zjb6fQpt++VGwfN0TEKi9USGh5qtdrKfrgZLBa99sbr+dPyIyLldTUPhULBcxs2AAZffgp+WOqqg6lfMelFUl+KSikNEgcziGmoa3TyDs7hFInEa9atjFSG6Dr0K1c9jRDGwDCYnjpxMiNzUmJSssu2N0pKKWESNQmUoJG4BQ9AjdGSGlVA+2IcRVQsFkaropEYA2Cbw2E0mimCVm17dXUtRbyhr59zOFlGkJyRqh4fTyh1OglC0NHdrVQq5YpIT1NwNC0FN66Cp/sE/sOPoRDq5PnBgaH2pg6jyciyWCRhKdDHtY2tTR2dPd0Oh1UmC+vo1p385pyT4xkGA2AAsJptAwNDNosNAWA/KB1drdGRysYX8dwplxBCCEKUwUx0tILjOZPRJGCYR5U13505x7AkJUtzuPird9/+z5Kyqwf+Z//pk6fbOhoFrBBjhADmFRYKRGx7VycAwu50MmbFBl7G7o92ns6fSwiMoeJRRWNTg1QagoXQ2Nbw3rvvnT1+ISQ4KHZcpBCzVovz47/8QxLMLlwy793dv/rb/r8OmwdZll26dNHixQv++c9jApZhXb0GOrp16kdQ/bssgAABxphSBAhulpV/8Ntfq5Sq9etWX7v0w4d//HDu/BlFzzxT/I8jdbV1C+cvjIlV3iq7nZKdLgsOu1tacfby+duVt9WquKTECZTjT3xzYmpBPrNx+zuJUcEj/Vjkx9HdbSAIoB8ud0EMwxCnc+ebux49qJleUGA0mu5X3JdKZFtefb6tvbX4ky9sJv65l9fMK3yqvqH55vmSOHVsVm7WvCWzq2496tP367p1mZMz7XbbkUNH2DEzGx2jcQ/utqdvH4vS0lvl5XfL5ubP3rB5g8VsXhe/lrPznx787Mb1q8kpcTq9qb6hOTU1Sd/WFTVOIVWGHi0+Nnfx/DlzZufNzA2SBEVFyePi427efJFF4OcfdOyU6xOlhBBKMQBmGAD485/2ETu/bfu2yZOzKU/v3Lm7b/8fKsqqhIxk+oLpzSev/9fuX1eWV5SWlJpsdgJ8V4/uwL6/S4NClAlRiQnJdVWNKWnJhUsWsjAKIXyVEXDVVdQSQgiAiGG+O3f2wqVLC+bPn5o/TcAKWzqatr60rX/QEKuMkYfHDPdac3Kz+/XDlfdqlm9Y8biy/vyJ7wvm5YTLzL2GvssXroWtiRinHtfc3DSsG2Q9JcJPb5+4O2mUEkIwxoSSI4c+dzi47JwpIpHIwdn37fvQSa2ZWelGg2nXu7sYEN66fvOdvSvvlVYXrVuMCD7w509OnjwlkwWr41Qnjn2bmpaRmZWm14UoomJZ7K4h4UkIErD74hIFIeRwOAwGfUx0tFIZDQi09drjX56Kjo4QYJqYpJlRkAeItRstTo5MTNP0dOmlYSHLVs1vbmns6Oyqq2pKSku8OZiempatCJy2TOLMINhtDeiQNz06UCjkV5xUFCI3Wbr7dIZh4xCkWjHru3jNPG9PcP1D7UGvT5CFpmZmxkhVzQ0tG7b9Pr3565fu3xdER06d8Gszg6DrqvPZDGKggUWi3Xhwqdcuy90jK0OQKNch/r2JFmGjVZG2zh7pEIOAhyjjPn5O/82b9G8SIU8PjHhtVd3vv3GOy1NjcooRYRc6kBcZm5GxqRJ1dX1eXnZazYti09QJ6cmOzh7Y32zcdiEPc2ogA2NwGolsFGMEAJITppos9pamlvCw6Qmk7W7s1cdNR4Q+sV/705OnPjJJ8XHjhzr6Gy6faOMOLlxCdGT87Ka6jtqaxrXr18hCw27df1u8f7P5YrwY0eOM1ve3B0fLvG+u7db7w/jgarBgBkGh4WGHf7yiK7HMCkzPTQsmHM6k5ISW9taTEP9TxcVTcxKunvn3vGjJ+vqa41GU0Nd88OKquqq6o62zmZtS+nN8pxpk7e8/AII4eDHxcyWt/bEy8ToCWW1t2ryvYYxxhgQQnK53Ggyf/vtqd7Ovs72DgxkWn5BcmryXz/8CBCsXrtc16u7X1K9dO0CZWzMlfOloVLZlOnZDTVNEQrFouWLdr69UyQU/Pr939wsu8lseXN3vEwMTwKwsRzHRed5njAMzp2S097R8c2JE5Ty2rr64cGBOfNmqWLjJLKgAYPhy0NH+/sGpLKQgT69k3NsfmnzihVLec655tlnlixbfOnClb179l66dHlqXj77hG2xEcsAAgqBuzQ8zzscHKFseHjEvn1/mpiStP8v+wHhoeGh6qqHOVOmzFk8x2a0T5tWMLnA/sO5UsOAftOWdXPmzbQaLSzDdnd1f/XFl8UHDxGCXtryyp69P4cbLQMz46R0VLB4NjnB0+WnIzZCiGFcfWPkdDoFLGt32K9cufy733/4qKp6wDCgUsckqOPz8wuS01IypqScOfG9tr5h5qxp4+JUZ06eLSkpN/QbEJD8/Jlbt22ZmpenUCi8cgTYBbwUKGBfh1CEARobGj/628c5OVlr1q6lAAyAQChoamoqv3XrVvmd/mHDnZK7HLE6nZw6To0IBgyGPgMAtgya1alx8rCYTS9syJ0yJValdDF+1oMeMNYuFIy53wgIvjlx/Pg3x2fMKhCJRV7mrNFoNBrN8mXLLFZrb6+usVX7xaF/tnW2d7d1AKI5U3MXLlo4LlqZkDQ+XCpXqmIwYI7jODvX09f5f6aEOmEXYxozAAAAAElFTkSuQmCC";
  const MAP_DATA="data:image/png;base64,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";
  const card=(label:string,value:string,kind:"blue"|"green"|"yellow"|"red"="blue")=>'<div class="kpi '+kind+'"><div class="kpiLabel">'+esc(label)+'</div><div class="kpiValue">'+esc(value)+'</div></div>';
  const flowCard=(label:string,prev:number,latest:number,kind:"blue"|"green"|"yellow"|"red"="blue")=>{const d=latest-prev;return '<div class="kpi '+kind+'"><div class="kpiLabel">'+esc(label)+'</div><div class="flow"><span>'+n(prev)+'</span><b>→</b><span>'+n(latest)+'</span></div><div class="flowDiff">('+(d>=0?"+":"")+n(d)+')</div></div>';};
  let h='<!doctype html><html><head><meta charset="utf-8"><title>AC-34 MATIALA - NOTICE & HEARING MONITORING REPORT</title><style>';
  h+='@page{size:A4 landscape;margin:7mm 8mm 10mm}*{box-sizing:border-box}html,body{margin:0;padding:0;background:#fff;color:#11243a;font-family:Arial,Helvetica,sans-serif;font-size:7px}body{counter-reset:page}';
  h+='.page{position:relative;min-height:193mm;padding-bottom:10mm;page-break-after:always}.page:last-child{page-break-after:auto}';
  h+='.mast{height:25mm;border-bottom:2px solid #0b4f83;position:relative;display:grid;grid-template-columns:24mm 1fr 45mm;align-items:center;margin-bottom:4mm}.emblem{width:16mm;height:19mm;justify-self:start;display:flex;align-items:center}.emblem img{width:12mm;height:16mm;object-fit:contain}.mapMark{font-size:7px;font-weight:900;color:#58758c;margin-top:3mm;display:flex;align-items:center;justify-content:flex-end;gap:1mm}.mapMark img{width:12mm;height:8mm;object-fit:contain}.title{text-align:center;color:#0c3152}.title h1{font-size:17px;line-height:1;margin:0;font-weight:900}.title h2,.title h3{font-size:8px;line-height:1.1;margin:2px 0 0;font-weight:900}.meta{text-align:right;color:#5b7182;font-size:6px;line-height:1.35;align-self:start;padding-top:2mm}';
  h+='.bar{height:7mm;background:#075184;color:#fff;border-radius:1.5mm 1.5mm 0 0;padding:1.4mm 2.2mm;font-size:9px;font-weight:900;letter-spacing:.15px;margin-bottom:2mm}.bar.red{background:#a92121}';
  h+='.kpiGrid{display:grid;grid-template-columns:repeat(6,1fr);gap:2mm;margin-bottom:3mm}.kpiGrid.row5{grid-template-columns:repeat(5,1fr)}.kpiSpacer{display:none;height:0}.reviewRow td{background:#d8ebff!important;color:#123d5d!important;font-weight:700}.reviewRow td:nth-child(12){background:#dcebd5!important;color:#17212b!important;font-weight:900}.reviewBadge{display:inline-block;background:#0b4f83;color:#fff;font-size:5.2px;font-weight:900;padding:.5mm 1mm;border-radius:1mm;margin-left:1mm;white-space:nowrap}.kpi{height:16mm;border:1px solid #9cc6e7;border-radius:1.5mm;background:#eef7ff;text-align:center;padding:1.7mm 1mm}.kpi.green{border-color:#91cf9b;background:#f0faee}.kpi.yellow{border-color:#e6c96d;background:#fff9e2}.kpi.red{border-color:#efaaaa;background:#fff0f0}.kpiLabel{font-size:6.2px;font-weight:900;color:#2b4c62;line-height:1.05;min-height:5mm}.kpiValue{font-size:11px;font-weight:900;color:#102a40;line-height:1.1;margin-top:1mm}.flow{display:flex;justify-content:center;align-items:center;gap:1.4mm;font-size:9.3px;font-weight:900;margin-top:1.3mm}.flow b{color:#1a5a87}.flowDiff{font-size:7px;font-weight:900;color:#2f7138;margin-top:.6mm}';
  h+='.summaryTable,.detailTable,.zeroTable,.zeroSummary{width:100%;border-collapse:collapse;table-layout:fixed}.summaryTable th{background:#0b4f83;color:#fff;border:1px solid #0a3e65;padding:1.35mm .7mm;font-size:6.2px;line-height:1.05;text-align:center;font-weight:900}.summaryTable td{border:1px solid #6f7e89;padding:1.05mm .7mm;font-size:6.2px;line-height:1.05;text-align:center;vertical-align:middle}.summaryTable td.name{text-align:left;padding-left:1.2mm;font-weight:800}.rate{background:#dff0d8!important;font-weight:900}.summaryTable tr.total td,.detailTable tr.total td,.zeroTable tr.total td{background:#ffe95a;font-weight:900;border-top:2px solid #7b6a00}.num,.nowrap{white-space:nowrap}.wrap{overflow-wrap:anywhere}.zeroRow td{background:#ffd7d7!important;color:#8d1717;font-weight:800}.zeroRow td.rate{background:#ffd7d7!important}';
  h+='.officerHead{display:flex;align-items:center;justify-content:space-between;margin:0 0 2mm;padding:1.5mm 2mm;background:#edf7ff;border:1px solid #a8cbe5;border-radius:1.5mm}.officerTitle{font-size:10px;font-weight:900;color:#0d4169}.numCircle{display:inline-flex;width:5mm;height:5mm;border-radius:50%;background:#0b5a91;color:#fff;align-items:center;justify-content:center;font-size:7px;margin-right:1.2mm}.officerPS{font-size:7.5px;font-weight:900;color:#0d4169}.miniGrid{display:grid;grid-template-columns:repeat(6,1fr);gap:1.8mm;margin-bottom:2.3mm}.miniCard{height:14mm;border:1px solid #9cc6e7;border-radius:1.4mm;background:#f4faff;text-align:center;padding:1.2mm .8mm}.miniLabel{font-size:5.8px;font-weight:900;color:#2e4d61;min-height:4mm}.miniValue{font-size:8px;font-weight:900;color:#152e42;margin-top:.8mm}.miniDiff{font-size:6.4px;font-weight:900;color:#2f7138;margin-top:.3mm}.miniGrid .miniCard:nth-child(4){background:#f0faee;border-color:#91cf9b}.miniGrid .miniCard:nth-child(5){background:#fff9e2;border-color:#e6c96d}.miniGrid .miniCard:nth-child(6){background:#fff0f0;border-color:#efaaaa}';
  h+='.reviewCaption{background:#0b4f83;color:#fff;border:1px solid #2563eb;text-align:center;font-size:6.5px;font-weight:900;padding:1.2mm;margin:0;border-radius:1mm 1mm 0 0}.detailTable thead,.zeroTable thead,.summaryTable thead{display:table-header-group}.detailTable th{background:#0b4f83;color:#fff;border:1px solid #0a3e65;padding:1.15mm .55mm;font-size:5.8px;line-height:1;text-align:center;font-weight:900}.detailTable td{border:1px solid #6f7e89;padding:.85mm .55mm;font-size:5.8px;line-height:1.05;vertical-align:middle;text-align:center}.detailTable td.left{text-align:left}.detailTable tr{break-inside:avoid;page-break-inside:avoid}.zeroIntro{font-size:7px;color:#526b7c;margin:-1mm 0 2mm}.zeroTable th{background:#a92121;color:#fff;border:1px solid #7e1818;padding:1.35mm .7mm;font-size:6.2px;font-weight:900}.zeroTable td{border:1px solid #a65d5d;background:#fff1f1;padding:1.1mm .7mm;font-size:6.2px;text-align:center;font-weight:700}.zeroSummary{margin:2mm 0 3mm}.zeroSummary th{background:#0b4f83;color:#fff;border:1px solid #0a3e65;padding:1.2mm;font-size:6px}.zeroSummary td{border:1px solid #6f7e89;padding:1mm;font-size:6px;text-align:center}.zeroSummary td:first-child{text-align:left;font-weight:800}.zeroSummary td.zero{color:#a92121;font-weight:900}.statusPill{display:inline-block;border:1px solid #a92121;border-radius:2mm;padding:.5mm 1.2mm;font-size:5.8px;font-weight:900;color:#8d1717;background:#fff}';
  h+='.footer{position:absolute;left:8mm;right:8mm;bottom:0;height:7mm;border-top:1px solid #b9c6cf;background:#fff;display:grid;grid-template-columns:1fr 1.3fr auto;align-items:center;color:#425a6b;font-size:5.8px;z-index:100}.legend{font-weight:900;white-space:nowrap}.blueBox{display:inline-block;width:4mm;height:2.5mm;background:#d8ebff;border:1px solid #9cc6e7;vertical-align:-.4mm;margin-right:1mm}.redBox{display:inline-block;width:4mm;height:2.5mm;background:#ffd7d7;border:1px solid #e3a1a1;vertical-align:-.4mm;margin:0 1mm 0 2mm}.pageNo{font-weight:900;color:#24455c;border:1px solid #a9c6d8;border-radius:1mm;padding:1.2mm 2mm;background:#f2f8fc;white-space:nowrap}';
  h+='@media print{.page{page-break-after:always}.page:last-child{page-break-after:auto}}';
  h+='</style></head><body>';

  h+='<section class="page"><div class="mast"><div class="emblem"><img src="'+EMBLEM_DATA+'" alt="" /></div><div class="title"><h1>AC-34 MATIALA</h1><h2>SIR-2026</h2><h3>NOTICE &amp; HEARING MONITORING REPORT</h3><h3>Previous ECI vs Latest ECI</h3></div><div class="meta">Report Generated: '+esc(stamp)+'<div class="mapMark"><img src="'+MAP_DATA+'" alt="" />AC-34 MATIALA</div></div></div>';
  h+='<div class="bar">EXECUTIVE SUMMARY</div><div class="kpiGrid row5">';
  h+=card("Total PS",n(grand.ps),"blue");h+=card("Notice Generated",n(grand.gen),"yellow");h+=card("Previous Delivered",n(grand.pd),"red");h+=card("Latest Delivered",n(grand.ld),"blue");h+=card("Delivery Increase",(grand.ld-grand.pd>=0?"+":"")+n(grand.ld-grand.pd),"green");h+='</div><div class="kpiGrid">';
  h+=card("Previous Hearing",n(grand.ph),"blue");h+=card("Latest Hearing",n(grand.lh),"green");h+=card("Hearing Increase",(grand.lh-grand.ph>=0?"+":"")+n(grand.lh-grand.ph),"green");h+=card("% Hearing Held",hearingPct.toFixed(2)+"%","blue");h+=card("Hearing Lapse",n(grand.lapse),"yellow");h+=card("Zero Hearing PS",n(zeroRows.length),"red");h+='</div>';
  h+='<div class="bar">AERO/Ad.AERO-wise Summary</div><table class="summaryTable"><thead><tr><th style="width:3.5%">S.No.</th><th style="width:15%">AERO / Ad.AERO</th><th style="width:7%">Designation</th><th style="width:4.5%">No. of PS</th><th style="width:7%">Notice<br>Generated</th><th style="width:7%">Notice<br>Deliv. Prev</th><th style="width:7%">Notice<br>Deliv. Latest</th><th style="width:6%">Diff.</th><th style="width:7%">Hearing<br>Held Prev</th><th style="width:7%">Hearing<br>Held Latest</th><th style="width:6%">Diff.</th><th style="width:6%">% Held<br>(of Gen.)</th><th style="width:7%">Hearing<br>Lapse</th></tr></thead><tbody>';
  groups.forEach((g,i)=>{h+='<tr><td>'+n(i+1)+'</td><td class="name wrap">'+esc(g.aero)+'</td><td class="nowrap">'+esc(g.rs[0]?.designation||"")+'</td><td>'+n(g.x.ps)+'</td><td>'+n(g.x.gen)+'</td><td>'+n(g.x.pd)+'</td><td>'+n(g.x.ld)+'</td><td class="num">'+(g.x.ld-g.x.pd>=0?"+":"")+n(g.x.ld-g.x.pd)+'</td><td>'+n(g.x.ph)+'</td><td>'+n(g.x.lh)+'</td><td class="num">'+(g.x.lh-g.x.ph>=0?"+":"")+n(g.x.lh-g.x.ph)+'</td><td class="rate">'+pct(g.x.lh,g.x.gen)+'</td><td>'+n(g.x.lapse)+'</td></tr>';});
  h+='<tr class="total"><td colspan="4">GRAND TOTAL</td><td>'+n(grand.gen)+'</td><td>'+n(grand.pd)+'</td><td>'+n(grand.ld)+'</td><td>'+ (grand.ld-grand.pd>=0?"+":"")+n(grand.ld-grand.pd)+'</td><td>'+n(grand.ph)+'</td><td>'+n(grand.lh)+'</td><td>'+ (grand.lh-grand.ph>=0?"+":"")+n(grand.lh-grand.ph)+'</td><td>'+hearingPct.toFixed(2)+'%</td><td>'+n(grand.lapse)+'</td></tr></tbody></table>';
  h+='<div class="footer"><span>AC-34 MATIALA | SIR-2026 | NOTICE &amp; HEARING MONITORING REPORT</span><span class="legend"><span class="blueBox"></span>Blue rows = OPERATIONAL REVIEW <span class="redBox"></span>Red rows = ZERO HEARING HELD</span><span class="pageNo">Page 1</span></div></section>';

  groups.forEach((g,gi)=>{
    h+='<section class="page"><div class="officerHead"><div class="officerTitle"><span class="numCircle">'+n(gi+1)+'</span>'+esc(g.aero)+' <span style="font-size:7.5px;font-weight:800">— '+esc(g.rs[0]?.designation||"")+' — '+n(g.x.ps)+' PS</span></div><div class="officerPS">Notice Generated: '+n(g.x.gen)+'</div></div>';
    h+='<div class="miniGrid">'+card("Notice Generated",n(g.x.gen),"blue")+flowCard("Delivered (Prev → Latest)",g.x.pd,g.x.ld,"blue")+flowCard("Hearing Held (Prev → Latest)",g.x.ph,g.x.lh,"blue")+'<div class="miniCard" style="background:#f0faee;border-color:#91cf9b"><div class="miniLabel">% Held</div><div class="miniValue">'+pct(g.x.lh,g.x.gen)+'</div></div><div class="miniCard" style="background:#fff9e2;border-color:#e6c96d"><div class="miniLabel">Hearing Lapse</div><div class="miniValue">'+n(g.x.lapse)+'</div></div><div class="miniCard" style="background:#fff0f0;border-color:#efaaaa"><div class="miniLabel">Zero Hearing PS</div><div class="miniValue">'+n(g.rs.filter(r=>r.latestHearing===0).length)+'</div></div></div>';
    h+='<div class="reviewCaption">OPERATIONAL REVIEW — 3 PS HIGHLIGHTED</div><table class="detailTable"><thead><tr><th style="width:3%">S.No.</th><th style="width:4%">P.S. No.</th><th style="width:14%">BLO Name</th><th style="width:10%">BLO Mobile No.</th><th style="width:7%">Notice<br>Generated</th><th style="width:7%">Notice<br>Deliv. Prev</th><th style="width:7%">Notice<br>Deliv. Latest</th><th style="width:5%">Diff.</th><th style="width:7%">Hearing<br>Held Prev</th><th style="width:7%">Hearing<br>Held Latest</th><th style="width:5%">Diff.</th><th style="width:7%">% Held</th><th style="width:7%">Hearing<br>Lapse</th></tr></thead><tbody>';
    g.rs.slice().sort((a,b)=>a.ps-b.ps).forEach((r,i)=>{const isZero=r.latestHearing===0;const isReview=reviewPsByOfficer.get(g.aero)?.has(r.ps)===true;const rowClass=isZero?'zeroRow':(isReview?'reviewRow':'');const badge=isReview?'<span class="reviewBadge">OPERATIONAL REVIEW</span>':'';h+='<tr class="'+rowClass+'"><td>'+n(i+1)+'</td><td>'+n(r.ps)+'</td><td class="left wrap">'+esc(r.bloName)+badge+(isZero?' <span class="statusPill">ZERO HEARING</span>':'')+'</td><td class="nowrap">'+esc(r.bloMobile)+'</td><td>'+n(r.noticeGenerated)+'</td><td>'+n(r.prevDelivered)+'</td><td>'+n(r.latestDelivered)+'</td><td>'+(r.latestDelivered-r.prevDelivered>=0?"+":"")+n(r.latestDelivered-r.prevDelivered)+'</td><td>'+n(r.prevHearing)+'</td><td>'+n(r.latestHearing)+'</td><td>'+(r.latestHearing-r.prevHearing>=0?"+":"")+n(r.latestHearing-r.prevHearing)+'</td><td class="rate">'+pct(r.latestHearing,r.noticeGenerated)+'</td><td>'+n(r.hearingLapse)+'</td></tr>';});
    h+='<tr class="total"><td colspan="4">TOTAL — '+esc(g.aero)+' ('+n(g.x.ps)+' PS)</td><td>'+n(g.x.gen)+'</td><td>'+n(g.x.pd)+'</td><td>'+n(g.x.ld)+'</td><td>'+(g.x.ld-g.x.pd>=0?"+":"")+n(g.x.ld-g.x.pd)+'</td><td>'+n(g.x.ph)+'</td><td>'+n(g.x.lh)+'</td><td>'+(g.x.lh-g.x.ph>=0?"+":"")+n(g.x.lh-g.x.ph)+'</td><td>'+pct(g.x.lh,g.x.gen)+'</td><td>'+n(g.x.lapse)+'</td><td>'+n(g.rs.filter(r=>r.latestHearing===0).length)+' Zero</td></tr></tbody></table>';
    h+='<div class="footer"><span>AC-34 MATIALA | SIR-2026 | NOTICE &amp; HEARING MONITORING REPORT</span><span class="legend"><span class="blueBox"></span>Blue = Report Highlight <span class="redBox"></span>Red rows = ZERO HEARING HELD</span><span class="pageNo">Page '+n(gi+2)+'</span></div></section>';
  });

  h+='<section class="page"><div class="bar red">ZERO HEARING PS — ACTION REQUIRED</div><div class="zeroIntro">All PS where Latest Hearing Held = 0. This list is generated dynamically from the latest dashboard data.</div><div class="bar">AERO / Ad.AERO-wise Zero Hearing Summary</div><table class="zeroSummary"><thead><tr><th>AERO / Ad.AERO</th><th>Designation</th><th>Total PS</th><th>Zero Hearing PS</th><th>Zero Hearing %</th></tr></thead><tbody>';
  groups.forEach(g=>{const z=g.rs.filter(r=>r.latestHearing===0).length;h+='<tr><td>'+esc(g.aero)+'</td><td>'+esc(g.rs[0]?.designation||"")+'</td><td>'+n(g.x.ps)+'</td><td class="zero">'+n(z)+'</td><td>'+(g.x.ps?(z/g.x.ps*100).toFixed(2):"0.00")+'%</td></tr>';});
  h+='</tbody></table><div class="bar red">ZERO HEARING PS — DETAILED LIST</div><table class="zeroTable"><thead><tr><th style="width:4%">S.No.</th><th style="width:16%">AERO / Ad.AERO</th><th style="width:17%">Supervisor</th><th style="width:6%">PS No.</th><th style="width:15%">BLO Name</th><th>Notice Gen.</th><th>Delivered</th><th>Hearing Held</th><th>Hearing Lapse</th><th style="width:9%">Status</th></tr></thead><tbody>';
  zeroRows.forEach((r,i)=>{h+='<tr><td>'+n(i+1)+'</td><td>'+esc(r.aero)+'</td><td>'+esc(r.supervisor)+'</td><td>'+n(r.ps)+'</td><td>'+esc(r.bloName)+'</td><td>'+n(r.noticeGenerated)+'</td><td>'+n(r.latestDelivered)+'</td><td>0</td><td>'+n(r.hearingLapse)+'</td><td><span class="statusPill">ZERO</span></td></tr>';});
  h+='<tr class="total"><td colspan="5">TOTAL ZERO HEARING PS</td><td colspan="5">'+n(zeroRows.length)+' / '+n(grand.ps)+' ('+zeroPct.toFixed(2)+'%)</td></tr></tbody></table><div class="footer"><span>AC-34 MATIALA | SIR-2026 | NOTICE &amp; HEARING MONITORING REPORT</span><span class="legend"><span class="blueBox"></span>Blue = Report Highlight <span class="redBox"></span>Red rows = ZERO HEARING HELD</span><span class="pageNo">Page '+n(groups.length+2)+'</span></div></section>';
  h+='</body></html>';

  const w=window.open("","_blank","width=1800,height=1200");
  if(!w){alert("Please allow pop-ups for PDF export.");return;}
  w.document.write(h);w.document.close();w.focus();setTimeout(()=>{w.focus();w.print();},900);
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
