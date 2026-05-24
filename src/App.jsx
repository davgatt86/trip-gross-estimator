import React, { useState, useMemo, useRef } from "react";
import * as XLSX from "xlsx";

/* ============================================================
   Trip Gross Estimator v2 — Peterhead vs Hanstholm
   - Upload boat totals file (csv/tsv text or PDF) -> auto tally
   - Upload PD + DK price sheets (vision) -> review grids
   - Name-based grade mapping (species + size name), editable
   - Estimated gross per market, side by side, CSV export
   ============================================================ */

const FONT = `"DM Sans", -apple-system, sans-serif`;
const MONO = `"DM Mono", ui-monospace, monospace`;
const C = {
  bg: "#0f1419", panel: "#1a2129", panel2: "#222b35", line: "#33404d",
  ink: "#e8edf2", dim: "#8b9aa8", pd: "#4a9eff", dk: "#ff7a45",
  good: "#3ddc84", warn: "#ffcc44", bad: "#ff5470",
};

/* Name-based grade dictionary: species -> size token -> {pd grade, dk sort, conf}
   Direction confirmed: 1 = smallest; names are authoritative.
   PD A1(big)..A5(small)/U9 ; DK sort 1(big), higher=smaller, 0=above 1. */
const GRADE_DICT = {
  CAT:{ SMALL:{pd:"U9",dk:"2",c:"med"}, MEDIUM:{pd:"U9",dk:"1",c:"med"}, LARGE:{pd:"U9",dk:"1",c:"med"} },
  COD:{ ROBBIE:{pd:"A5",dk:"5",c:"high"}, BABY:{pd:"A5",dk:"5",c:"high"}, "BIG BABY":{pd:"A4",dk:"4",c:"high"},
        MEDIUM:{pd:"A3",dk:"3",c:"high"}, SPRAG:{pd:"A2",dk:"2",c:"high"}, COD:{pd:"A1",dk:"1",c:"high"}, LARGE:{pd:"A1",dk:"1",c:"high"}, XL:{pd:"A1",dk:"0",c:"med"}, "X LARGE":{pd:"A1",dk:"0",c:"med"} },
  HADDOCK:{ "MINI METRO":{pd:"A4",dk:"3",c:"high"}, METRO:{pd:"A4",dk:"3",c:"med"}, CHIPPER:{pd:"A4",dk:"2",c:"med"},
        SEED:{pd:"A3",dk:"2",c:"high"}, "GOOD SEED":{pd:"A2",dk:"1",c:"high"}, PINGER:{pd:"A2",dk:"1",c:"high"}, CHAT:{pd:"A1",dk:"1",c:"high"} },
  HAKE:{ "X SMALL":{pd:"A4",dk:"4",c:"high"}, SMALL:{pd:"A3",dk:"3",c:"high"}, SEL:{pd:"A2",dk:"2",c:"high"},
        SELECTED:{pd:"A2",dk:"2",c:"high"}, MEDIUM:{pd:"A1",dk:"1",c:"high"}, LARGE:{pd:"A1",dk:"1",c:"high"}, XL:{pd:"A1",dk:"0",c:"med"}, "X LARGE":{pd:"A1",dk:"0",c:"med"} },
  HALIBUT:{ "SIZE 1":{pd:"U9",dk:"1",c:"low"}, SMALL:{pd:"U9",dk:"2",c:"med"}, LARGE:{pd:"U9",dk:"1",c:"med"} },
  "LEMON SOLE":{ SMALL:{pd:"A3",dk:"3",c:"low"}, MEDIUM:{pd:"A2",dk:"2",c:"low"}, LARGE:{pd:"A1",dk:"2",c:"low"} },
  LING:{ SMALL:{pd:"A3",dk:"3",c:"high"}, MEDIUM:{pd:"A2",dk:"2",c:"high"}, LARGE:{pd:"A1",dk:"1",c:"high"} },
  LYTHE:{ SMALL:{pd:"A4",dk:"4",c:"high"}, SEL:{pd:"A3",dk:"3",c:"high"}, SELECTED:{pd:"A3",dk:"3",c:"high"},
        MEDIUM:{pd:"A2",dk:"2",c:"high"}, LARGE:{pd:"A2",dk:"1",c:"med"} },
  MEGRIM:{ "X SMALL":{pd:"A3",dk:"3",c:"med"}, SMALL:{pd:"A3",dk:"3",c:"high"}, MEDIUM:{pd:"A2",dk:"2",c:"med"}, LARGE:{pd:"A2",dk:"1",c:"med"} },
  MONKFISH:{ FROGS:{pd:"A5",dk:"5",c:"high"}, SMALL:{pd:"A4",dk:"4",c:"high"}, SEL:{pd:"A3",dk:"3",c:"high"},
        SELECTED:{pd:"A3",dk:"3",c:"high"}, MEDIUM:{pd:"A2",dk:"2",c:"high"}, LARGE:{pd:"A1",dk:"1",c:"high"}, XL:{pd:"A1",dk:"1",c:"med"} },
  SAITHE:{ "X SMALL":{pd:"A4",dk:"4",c:"high"}, SMALL:{pd:"A4",dk:"4",c:"high"}, SEL:{pd:"A3",dk:"3",c:"high"},
        SELECTED:{pd:"A3",dk:"3",c:"high"}, MEDIUM:{pd:"A2",dk:"2",c:"high"}, LARGE:{pd:"A1",dk:"1",c:"high"} },
  SQUID:{ SEL:{pd:"ANY",dk:"2",c:"low"}, SELECTED:{pd:"ANY",dk:"2",c:"low"}, MEDIUM:{pd:"ANY",dk:"2",c:"low"}, SMALL:{pd:"ANY",dk:"2",c:"low"} },
  TURBOT:{ "SIZE 1":{pd:"U9",dk:"—",c:"low"}, LARGE:{pd:"U9",dk:"—",c:"low"} },
  TUSK:{ "SIZE 1":{pd:"U9",dk:"1",c:"low"}, MIX:{pd:"U9",dk:"1",c:"med"} },
  WHITING:{ "SMALL ROUND":{pd:"A4r",dk:"2",c:"med"}, ROUND:{pd:"A4r",dk:"2",c:"med"}, SMALL:{pd:"A2",dk:"1",c:"med"}, MEDIUM:{pd:"A2",dk:"1",c:"high"} },
  WITCH:{ "SIZE 3":{pd:"U9",dk:"3",c:"low"}, ROUND:{pd:"U9",dk:"2",c:"low"} },
};

const SP_TO_PD = { CAT:"Catfish", COD:"Cod", HADDOCK:"Haddock", HAKE:"Hake", HALIBUT:"Halibut",
  "LEMON SOLE":"Lemons", LING:"Ling", LYTHE:"Lythe", MEGRIM:"Megrim", MONKFISH:"Monks",
  SAITHE:"Coley", SQUID:"Squid", TURBOT:"Turbot", TUSK:"Tusk", WHITING:"Whiting", WITCH:"Witch" };
const SP_TO_DK = { CAT:"Catfishes", COD:"Atlantic Cod", HADDOCK:"Haddock", HAKE:"European Hake",
  HALIBUT:"Atlantic Halibut", "LEMON SOLE":"Lemon Sole", LING:"Ling", LYTHE:"Pollack", MEGRIM:"Megrim",
  MONKFISH:"Monkfish", SAITHE:"Saithe", SQUID:"Squid", TURBOT:"—", TUSK:"Tusk", WHITING:"Whiting", WITCH:"Witch Flounder" };

const DEFAULT_PD = {
  Cod:{A1:7.08,A2:7.26,A3:7.41}, Haddock:{A1:6.75,A2:6.98,A3:4.77,A4:2.42,A4c:2.50,A4m:2.50,A4ma:2.20},
  Whiting:{A2:4.08,A4r:2.10}, Catfish:{U9:2.06}, Ling:{A1:2.55,A2:2.55,A3:2.55},
  Coley:{A1:3.36,A2:3.10,A3:2.58,A4:2.50}, Monks:{A1:4.78,A2:5.00,A3:5.20,A4:4.60,A5:3.57},
  Lythe:{A1:6.25,A2:5.32,A4:4.45}, Lemons:{A2:10.80,A3:1.86}, Plaice:{A2:2.50,A4:0.75},
  Megrim:{A1:2.67,A2:1.33,A3:1.00,A4:0.90}, Hake:{A1:10.00,A2:8.33,A3:6.25,A4:4.74,A5:1.40},
  Turbot:{U9:11.55}, Halibut:{U9:9.27}, Witch:{U9:1.66}, Skate:{U9:2.25},
};
const DEFAULT_DK = {
  Squid:{"2":0.58,"9":0.58}, "Blue Ling":{"1":2.90,"2":2.90,"3":2.55,"9":1.39}, Tusk:{"2":3.91,"9":3.72},
  Megrim:{"2":7.14}, Catfishes:{"1":3.94,"2":4.17,"3":1.49},
  Monkfish:{"1":6.03,"2":5.76,"3":5.65,"4":5.47,"5":2.91},
  "Atlantic Halibut":{"0":11.02,"1":9.62,"2":8.60,"3":8.59,"4":8.67,"5":7.03,"9":8.70},
  Whiting:{"1":3.55,"2":2.32}, "Common Dab":{"1":1.16,"2":0.71}, "Norway Lobster":{"1":14.16,"2":8.23,"3":2.69},
  Haddock:{"1":6.52,"2":3.36,"3":2.12,"4":0.49,"9":0.73},
  "European Hake":{"0":8.72,"1":7.70,"2":6.12,"3":4.36,"4":1.61}, Ling:{"1":4.24,"2":4.11,"3":4.13},
  Pollack:{"2":9.72,"3":6.26,"4":5.20}, Mackerel:{"1":3.55,"2":3.15},
  Saithe:{"1":5.11,"2":5.65,"3":5.55,"4":3.83,"9":1.74},
  Turbot:{"0":23.21,"1":21.14,"2":16.93,"3":18.27,"4":13.09},
  "European Plaice":{"0":3.51,"1":3.30,"2":3.80,"3":2.64,"4":1.48,"9":0.93},
  "Lemon Sole":{"1":6.96,"2":5.80,"3":3.48}, "Greater Forkbeard":{"1":3.88,"2":3.27,"9":2.63},
  "Witch Flounder":{"1":6.61,"2":4.51,"3":1.62},
  "Atlantic Cod":{"0":8.64,"1":8.15,"2":7.79,"3":7.61,"4":6.26,"5":4.55},
};

const DEFAULT_TALLY = [
  ["CAT","1. Small (U9b)",12,488.3],["CAT","2. Large (U9a)",19,776.4],
  ["COD","1. Robbie (5b)",2,60.78],["COD","2. Baby (5a)",3,91.31],["COD","3. Big Baby (4)",7,212.66],["COD","4. Medium (3)",9,278.26],["COD","5. Sprag (2)",30,923.34],["COD","6. Large (1b)",11,331.45],["COD","7. X Large (1a)",2,62.25],
  ["HADDOCK","1. Mini Metro (4)",38,1537.31],["HADDOCK","2. Metro (3)",39,1575.73],["HADDOCK","3. Chipper (2b)",29,1176.3],["HADDOCK","4. Seed (2a)",19,773.33],["HADDOCK","5. Good Seed (1d)",20,816.18],["HADDOCK","6. Pinger (1c)",2,80.67],
  ["HAKE","2. X Small (4)",3,91.8],["HAKE","3. Small (3)",7,212.17],["HAKE","4. Selected (2)",33,1012.88],["HAKE","5. Medium (1c)",18,554.17],["HAKE","6. Large (1b)",8,246.25],["HAKE","7. X Large (1a)",1,30.44],
  ["HALIBUT","1. Small (U9b)",3,62.0],["HALIBUT","2. Large (U9a)",9,199.0],
  ["LEMON SOLE","1. Small (3)",1,30.5],["LEMON SOLE","2. Medium (2)",1,30.1],["LEMON SOLE","3. Large (1b)",4,123.8],
  ["LING","1. Small (3)",12,492.5],["LING","2. Medium (2)",22,901.8],["LING","3. Large (1)",57,2259.9],
  ["LYTHE","1. Small (4)",8,324.86],["LYTHE","2. Selected (3)",15,610.02],["LYTHE","3. Medium (2)",19,777.77],["LYTHE","4. Large (1)",23,937.44],
  ["MEGRIM","2. Small (3)",2,60.4],["MEGRIM","4. Medium (1b)",2,62.1],["MEGRIM","5. Large (1a)",5,155.0],
  ["MONKFISH","1. Frogs (5)",3,120.65],["MONKFISH","2. Small (4)",3,122.36],["MONKFISH","3. Selected (3)",11,452.49],["MONKFISH","4. Medium (2)",12,493.91],["MONKFISH","5. Large (1)",23,952.0],
  ["SAITHE","1. X Small (4b)",86,3482.12],["SAITHE","2. Small (4a)",73,2961.05],["SAITHE","3. Selected (3)",64,2608.42],["SAITHE","4. Medium (2)",20,816.38],["SAITHE","5. Large (1)",6,246.45],
  ["SQUID","2. Small (U9e)",10,252.45],["SQUID","3. Selected (U9d)",1,25.34],
  ["TURBOT","1. Small (U9b)",1,22.4],
  ["TUSK","1. Mix (U9a)",6,246.0],
  ["WHITING","1. Small Round (4)",2,80.97],["WHITING","2. Round (3)",25,1013.87],["WHITING","3. Small (2)",8,324.36],["WHITING","4. Medium (1b)",7,284.18],
].map((r,i)=>({id:i,sp:r[0],size:r[1],boxes:r[2],wt:r[3],avgBox:+(r[3]/r[2]).toFixed(1)}));

const fmtGBP=(n)=>n==null||isNaN(n)?"—":"£"+n.toLocaleString("en-GB",{minimumFractionDigits:2,maximumFractionDigits:2});
const fmtKg=(n)=>n==null||isNaN(n)?"—":n.toLocaleString("en-GB",{maximumFractionDigits:1})+" kg";
const GRADE_LABEL = { A4c:"A4 Chipper", A4m:"A4 Metro (high)", A4ma:"A4 Metro (avg)", A4r:"A4 Round" };

const norm=(s)=>String(s||"").toUpperCase().replace(/^\d+\.?\s*/,"").replace(/\s*\([^)]*\)\s*$/,"").trim();

// Extract the bracket code, e.g. "1. Robbie (5b)" -> "5B". This is the TRUE grade;
// the leading number before the name is just the scales touchscreen button position.
const bracket=(s)=>{const m=String(s||"").match(/\(([^)]*)\)\s*$/);return m?m[1].toUpperCase().trim():"";};

// Per-species bracket-code -> mapping. Currently Haddock only (per user).
// Haddock A4 has THREE named PD prices, so we use distinct pd keys: A4c=Chipper, A4m=Metro(high), A4ma=Metro(avg).
const BRACKET_DICT = {
  COD:{
    "5B":{pd:"A5",dk:"5",c:"high"},
    "5A":{pd:"A5",dk:"5",c:"high"},
    "4":{pd:"A4",dk:"4",c:"high"},
    "3":{pd:"A3",dk:"3",c:"high"},
    "2":{pd:"A2",dk:"2",c:"high"},
    "1B":{pd:"A1",dk:"1",c:"high"},
    "1A":{pd:"A1",dk:"0",c:"med"},   // 1a = top end -> DK Sort 0 (above grade)
  },
  HADDOCK:{
    "1A":{pd:"A1",dk:"1",c:"high"},
    "1B":{pd:"A1",dk:"1",c:"high"},
    "1C":{pd:"A1",dk:"1",c:"high"},   // Pinger
    "1D":{pd:"A2",dk:"1",c:"high"},   // Good Seed
    "2A":{pd:"A3",dk:"2",c:"high"},   // Seed
    "2B":{pd:"A4c",dk:"2",c:"high"},  // Chipper -> A4 Chipper price
    "3":{pd:"A4m",dk:"3",c:"high"},   // Metro -> A4 Metro HIGH column
    "4":{pd:"A4ma",dk:"3",c:"high"},  // Mini Metro -> A4 Metro AVE column
  },
};

function buildMapping(tally){
  return tally.map((r)=>{
    const code=bracket(r.size);
    // Prefer per-species bracket-code mapping when available
    const bd=BRACKET_DICT[r.sp]&&code&&BRACKET_DICT[r.sp][code];
    if(bd) return { pdSp:SP_TO_PD[r.sp]||"—", pdGr:bd.pd, dkSp:SP_TO_DK[r.sp]||"—", dkSort:bd.dk, conf:bd.c };
    // Fallback: existing name-based mapping
    const tok=norm(r.size);
    const d=(GRADE_DICT[r.sp]&&GRADE_DICT[r.sp][tok])||{pd:"—",dk:"—",c:"low"};
    return { pdSp:SP_TO_PD[r.sp]||"—", pdGr:d.pd, dkSp:SP_TO_DK[r.sp]||"—", dkSort:d.dk, conf:d.c };
  });
}

export default function App(){
  const [step,setStep]=useState(0);
  const [pd,setPd]=useState(DEFAULT_PD);
  const [dk,setDk]=useState(DEFAULT_DK);
  const [tally,setTally]=useState(DEFAULT_TALLY);
  const [map,setMap]=useState(()=>buildMapping(DEFAULT_TALLY));
  const [tallyMode,setTallyMode]=useState("weight");
  const [fillMissing,setFillMissing]=useState(true);
  const [busy,setBusy]=useState({pd:false,dk:false,boat:false});
  const [msg,setMsg]=useState({pd:"",dk:"",boat:"Loaded Trip 54 (so far) — prices default to 22/05/26. Upload to replace either."});

  const effW=(r)=>tallyMode==="boxes"?r.boxes*r.avgBox:r.wt;
  const pdPrice=(sp,gr)=>(pd[sp]&&pd[sp][gr]!=null?pd[sp][gr]:0);
  const dkPrice=(sp,so)=>(dk[sp]&&dk[sp][so]!=null?dk[sp][so]:0);

  const rows=useMemo(()=>tally.map((r,i)=>{
    const m=map[i]||{}; const w=effW(r);
    let pdN=pdPrice(m.pdSp,m.pdGr), dkN=dkPrice(m.dkSp,m.dkSort);
    let pu=pdN,du=dkN,note="";
    if(fillMissing){
      if(pdN===0&&dkN===0)note="No price either market";
      else if(pdN===0){pu=dkN;note="PD ← Hanstholm";}
      else if(dkN===0){du=pdN;note="DK ← Peterhead";}
    }
    return {...r,m,w,pdPrice:pu,dkPrice:du,pdTotal:w*pu,dkTotal:w*du,diff:w*du-w*pu,note};
  }),[tally,map,pd,dk,tallyMode,fillMissing]);

  const totals=useMemo(()=>{const t=rows.reduce((a,r)=>({w:a.w+r.w,pd:a.pd+r.pdTotal,dk:a.dk+r.dkTotal}),{w:0,pd:0,dk:0});return{...t,diff:t.dk-t.pd};},[rows]);
  const summary=useMemo(()=>{const m={};rows.forEach((r)=>{if(!m[r.sp])m[r.sp]={sp:r.sp,w:0,pd:0,dk:0};m[r.sp].w+=r.w;m[r.sp].pd+=r.pdTotal;m[r.sp].dk+=r.dkTotal;});return Object.values(m).map((s)=>({...s,diff:s.dk-s.pd}));},[rows]);

  const fileToB64=(f)=>new Promise((res,rej)=>{const r=new FileReader();r.onload=()=>res(r.result.split(",")[1]);r.onerror=rej;r.readAsDataURL(f);});
  const fileToText=(f)=>new Promise((res,rej)=>{const r=new FileReader();r.onload=()=>res(r.result);r.onerror=rej;r.readAsText(f);});
  const fileToBuf=(f)=>new Promise((res,rej)=>{const r=new FileReader();r.onload=()=>res(r.result);r.onerror=rej;r.readAsArrayBuffer(f);});

  // Parse boat xlsx: rows are [Species, Size, Boxes, Weight]. Species header has '*' in Size col;
  // size sub-rows have blank Species; 'Total' row ends it. Weight may carry a 'kg' suffix.
  function parseBoatRows(rows){
    const out=[];let curSp=null;
    const numOf=(v)=>{if(v==null)return 0;const m=String(v).replace(/[^0-9.\-]/g,"");return +m||0;};
    rows.forEach((cells)=>{
      if(!cells||!cells.length)return;
      const first=(cells[0]==null?"":String(cells[0])).trim();
      const second=(cells[1]==null?"":String(cells[1])).trim();
      if(first&&/^[A-Z][A-Z \/]+$/.test(first)&&(second==="*"||second===""))  {curSp=first;return;}
      if(/^total$/i.test(first))return;
      if(!first&&curSp&&second){
        const boxes=numOf(cells[2]);const wt=numOf(cells[3]);
        if(wt)out.push({sp:curSp,size:second,boxes,wt});
      }
    });
    return out;
  }

  async function parsePrice(file,which){
    setBusy((b)=>({...b,[which]:true}));setMsg((m)=>({...m,[which]:"Reading…"}));
    try{
      const b64=await fileToB64(file);const mt=file.type||"image/png";
      const prompt=which==="pd"
        ?`Peterhead fish market sheet with LOW, HIGH, AVE columns. Extract species + grade (A1..A5,U9) and the AVE price (GBP). Respond with ONLY a JSON object and nothing else — no explanation, no markdown. Shape: {"Species":{"A1":7.02}}. Use AVE column, skip blank cells. SPECIAL CASE Haddock A4 (rows Chipper, Metro, Round): keys "A4c"=Chipper AVE, "A4m"=Metro HIGH column, "A4ma"=Metro AVE, "A4"=Chipper AVE. Example {"Haddock":{"A1":6.75,"A4c":2.50,"A4m":2.50,"A4ma":2.20,"A4":2.50}}.`
        :`Hanstholm Danish auction sheet. Extract species + sort number (0,1,2,3,4,5,9) and the Avg. price (the second price on each row; the first is Max). Respond with ONLY a JSON object, no explanation, no markdown. Shape: {"Species":{"1":5.17}}. Sorts are strings. Keep names like "Atlantic Cod","European Hake".`;
      const resp=await fetch("/.netlify/functions/parse",{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify({media:b64,mediaType:mt,prompt})});
      const data=await resp.json();
      if(!resp.ok){throw new Error((data&&data.error)||`server ${resp.status}`);}
      let t=data.text||"";
      if(!t||!t.trim()){throw new Error("empty reply");}
      // Robust: pull the JSON object from the first { to the last }
      const a=t.indexOf("{"),z=t.lastIndexOf("}");
      if(a<0||z<0||z<=a){throw new Error("no JSON found");}
      const jsonStr=t.slice(a,z+1);
      const parsed=JSON.parse(jsonStr);
      if(!parsed||typeof parsed!=="object"||Array.isArray(parsed)||!Object.keys(parsed).length){throw new Error("no species parsed");}
      if(which==="pd"&&parsed.Haddock){const h=parsed.Haddock;if(h.A4!=null){if(h.A4c==null)h.A4c=h.A4;if(h.A4m==null)h.A4m=h.A4;if(h.A4ma==null)h.A4ma=h.A4;}}
      if(which==="pd")setPd(parsed);else setDk(parsed);
      setMsg((m)=>({...m,[which]:`Parsed ${Object.keys(parsed).length} species — review & correct below.`}));
    }catch(e){setMsg((m)=>({...m,[which]:`Auto-parse failed (${e.message}). Your current prices are kept — edit by hand, or try a clearer photo / the PDF.`}));}
    finally{setBusy((b)=>({...b,[which]:false}));}
  }

  function parseBoatText(txt){
    const out=[];let curSp=null;
    txt.split(/\r?\n/).forEach((line)=>{
      const cells=line.split(/[\t,]/).map((c)=>c.trim().replace(/^"|"$/g,""));
      if(cells.length<2)return;
      const first=cells[0];
      if(first&&first===first.toUpperCase()&&/[A-Z]/.test(first)&&cells[1]==="*"){curSp=first;return;}
      if(/^total$/i.test(first))return;
      if(!first&&curSp){
        const size=cells[1];const boxes=+cells[2]||0;const wt=+cells[3]||0;
        if(size&&wt)out.push({sp:curSp,size,boxes,wt});
      }
    });
    return out;
  }

  async function parseBoat(file){
    setBusy((b)=>({...b,boat:true}));setMsg((m)=>({...m,boat:"Reading boat file…"}));
    try{
      const name=(file.name||"").toLowerCase();let parsed=[];
      if(name.endsWith(".csv")||name.endsWith(".tsv")||name.endsWith(".txt")){
        parsed=parseBoatText(await fileToText(file));
      }else if(name.endsWith(".pdf")){
        const b64=await fileToB64(file);
        const prompt=`Boat landings/totals report. For every SIZE sub-row (NOT species summary '*' rows, NOT the grand Total), extract species, size label, boxes, weight kg. Respond with ONLY a JSON array, no explanation: [{"sp":"COD","size":"5. Sprag (2)","boxes":8,"wt":265.89}]. Species in CAPS, keep the size label exactly as printed including the bracket code.`;
        const resp=await fetch("/.netlify/functions/parse",{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify({media:b64,mediaType:"application/pdf",prompt})});
        const data=await resp.json();
        if(!resp.ok){throw new Error((data&&data.error)||`server ${resp.status}`);}
        let t=data.text||"";
        const a=t.indexOf("["),z=t.lastIndexOf("]");
        if(a<0||z<0||z<=a){throw new Error("no JSON array found");}
        parsed=JSON.parse(t.slice(a,z+1)).map((r)=>({sp:String(r.sp).toUpperCase(),size:r.size,boxes:+r.boxes||0,wt:+r.wt||0}));
      }else if(name.endsWith(".xlsx")||name.endsWith(".xls")||name.endsWith(".xlsm")){
        const buf=await fileToBuf(file);
        const wb=XLSX.read(buf,{type:"array"});
        const ws=wb.Sheets[wb.SheetNames[0]];
        const rows=XLSX.utils.sheet_to_json(ws,{header:1,blankrows:false,defval:""});
        parsed=parseBoatRows(rows);
      }else{
        setMsg((m)=>({...m,boat:"Unsupported file. Upload the boat .xlsx, a CSV, or the boat PDF."}));
        setBusy((b)=>({...b,boat:false}));return;
      }
      if(!parsed.length)throw new Error("no rows found");
      const t=parsed.map((r,i)=>({id:i,sp:r.sp,size:r.size,boxes:r.boxes,wt:r.wt,avgBox:r.boxes?+(r.wt/r.boxes).toFixed(1):0}));
      setTally(t);setMap(buildMapping(t));
      setMsg((m)=>({...m,boat:`Loaded ${t.length} size lines across ${new Set(t.map((x)=>x.sp)).size} species. Mapping auto-built — check step 4.`}));
    }catch(e){setMsg((m)=>({...m,boat:`Couldn't read boat file (${e.message}). Sample tally still loaded.`}));}
    finally{setBusy((b)=>({...b,boat:false}));}
  }

  function exportCSV(){
    const head=["Species","Size","Weight kg","PD Sp","PD Gr","PD £/kg","PD Total","DK Sp","DK Sort","DK £/kg","DK Total","Diff","Note"];
    const L=rows.map((r)=>[r.sp,r.size,r.w.toFixed(1),r.m.pdSp,r.m.pdGr,r.pdPrice.toFixed(2),r.pdTotal.toFixed(2),r.m.dkSp,r.m.dkSort,r.dkPrice.toFixed(2),r.dkTotal.toFixed(2),r.diff.toFixed(2),r.note]);
    L.push([]);L.push(["TOTAL","",totals.w.toFixed(1),"","","",totals.pd.toFixed(2),"","","",totals.dk.toFixed(2),totals.diff.toFixed(2),""]);
    const csv=[head,...L].map((r)=>r.map((c)=>`"${c??""}"`).join(",")).join("\n");
    const a=document.createElement("a");a.href=URL.createObjectURL(new Blob([csv],{type:"text/csv"}));a.download="trip_gross_estimate.csv";a.click();
  }

  const STEPS=["Boat tally","Peterhead prices","Hanstholm prices","Check mapping","Estimated gross"];
  return(
    <div style={{minHeight:"100vh",background:C.bg,color:C.ink,fontFamily:FONT,paddingBottom:80}}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;700;800&family=DM+Mono:wght@400;500&display=swap');
        *{box-sizing:border-box} input,select{font-family:${MONO}}
        table{border-collapse:collapse;width:100%}
        th,td{padding:7px 10px;border-bottom:1px solid ${C.line};text-align:left;font-size:13px}
        th{color:${C.dim};font-weight:600;font-size:11px;text-transform:uppercase;letter-spacing:.04em}
        td.r,th.r{text-align:right}
        .ci{width:70px;background:${C.panel2};border:1px solid ${C.line};color:${C.ink};padding:4px 6px;border-radius:5px;text-align:right}
        .si{background:${C.panel2};border:1px solid ${C.line};color:${C.ink};padding:4px 6px;border-radius:5px}
        tr:hover td{background:rgba(255,255,255,.02)}
      `}</style>

      <div style={{borderBottom:`1px solid ${C.line}`,padding:"18px 22px",background:C.panel,position:"sticky",top:0,zIndex:10}}>
        <div style={{display:"flex",alignItems:"baseline",gap:14,flexWrap:"wrap"}}>
          <div style={{fontSize:22,fontWeight:800,letterSpacing:"-.02em"}}>Trip Gross Estimator</div>
          <div style={{color:C.dim,fontSize:13}}>Peterhead vs Hanstholm</div>
          <div style={{marginLeft:"auto",display:"flex",gap:18}}>
            <Tot label="Peterhead" val={fmtGBP(totals.pd)} col={C.pd}/>
            <Tot label="Hanstholm" val={fmtGBP(totals.dk)} col={C.dk}/>
            <Tot label="Difference" val={(totals.diff>=0?"+":"")+fmtGBP(totals.diff)} col={totals.diff>=0?C.dk:C.pd}/>
          </div>
        </div>
        <div style={{display:"flex",gap:8,marginTop:16,flexWrap:"wrap"}}>
          {STEPS.map((s,i)=>(<button key={i} onClick={()=>setStep(i)} style={{background:i===step?C.ink:"transparent",color:i===step?C.bg:C.dim,border:`1px solid ${i===step?C.ink:C.line}`,padding:"6px 13px",borderRadius:20,fontSize:12.5,fontWeight:600,cursor:"pointer",fontFamily:FONT}}>{i+1}. {s}</button>))}
        </div>
      </div>

      <div style={{maxWidth:1180,margin:"0 auto",padding:"26px 22px"}}>
        {step===0&&<BoatStep tally={tally} setTally={setTally} mode={tallyMode} setMode={setTallyMode} effW={effW} onUpload={parseBoat} busy={busy.boat} msg={msg.boat} next={()=>setStep(1)}/>}
        {step===1&&<PriceStep which="pd" title="Peterhead price sheet" accent={C.pd} prices={pd} setPrices={setPd} onUpload={(f)=>parsePrice(f,"pd")} busy={busy.pd} msg={msg.pd} next={()=>setStep(2)}/>}
        {step===2&&<PriceStep which="dk" title="Hanstholm price sheet" accent={C.dk} prices={dk} setPrices={setDk} onUpload={(f)=>parsePrice(f,"dk")} busy={busy.dk} msg={msg.dk} next={()=>setStep(3)}/>}
        {step===3&&<MapStep tally={tally} map={map} setMap={setMap} pd={pd} dk={dk} next={()=>setStep(4)}/>}
        {step===4&&<ResultStep rows={rows} totals={totals} summary={summary} fillMissing={fillMissing} setFillMissing={setFillMissing} tallyMode={tallyMode} exportCSV={exportCSV}/>}
      </div>
    </div>
  );
}

function Tot({label,val,col}){return(<div style={{textAlign:"right"}}><div style={{fontSize:10.5,color:C.dim,textTransform:"uppercase",letterSpacing:".05em"}}>{label}</div><div style={{fontSize:16,fontWeight:600,color:col,fontFamily:MONO}}>{val}</div></div>);}
function Card({children,accent}){return(<div style={{background:C.panel,border:`1px solid ${C.line}`,borderRadius:12,padding:22,borderTop:accent?`3px solid ${accent}`:undefined}}>{children}</div>);}
function NextBtn({onClick,label="Next →"}){return(<button onClick={onClick} style={{background:C.good,color:"#06281a",border:"none",padding:"11px 22px",borderRadius:8,fontSize:14,fontWeight:700,cursor:"pointer",fontFamily:FONT,marginTop:20}}>{label}</button>);}
function UploadBtn({onUpload,busy,accent,label}){const ref=useRef();return(<><input ref={ref} type="file" accept="image/*,application/pdf,.csv,.tsv,.txt,.xlsx,.xls,.xlsm" style={{display:"none"}} onChange={(e)=>e.target.files[0]&&onUpload(e.target.files[0])}/><button onClick={()=>ref.current.click()} disabled={busy} style={{background:accent,color:"#031018",border:"none",padding:"10px 18px",borderRadius:8,fontWeight:700,cursor:busy?"wait":"pointer",fontFamily:FONT,fontSize:13.5}}>{busy?"Reading…":label}</button></>);}

function BoatStep({tally,setTally,mode,setMode,effW,onUpload,busy,msg,next}){
  function upd(id,f,v){setTally((p)=>p.map((r)=>r.id===id?{...r,[f]:v===""?0:+v}:r));}
  const totW=tally.reduce((a,r)=>a+effW(r),0);
  return(<Card>
    <div style={{display:"flex",alignItems:"center",gap:14,flexWrap:"wrap"}}>
      <div style={{fontSize:18,fontWeight:700}}>Boat tally</div>
      <UploadBtn onUpload={onUpload} busy={busy} accent={C.good} label="⬆ Upload boat file"/>
      <div style={{display:"flex",gap:6,background:C.panel2,padding:3,borderRadius:8}}>
        {[["weight","Weight"],["boxes","Boxes × avg"]].map(([m,l])=>(<button key={m} onClick={()=>setMode(m)} style={{background:mode===m?C.ink:"transparent",color:mode===m?C.bg:C.dim,border:"none",padding:"6px 12px",borderRadius:6,fontSize:12.5,fontWeight:600,cursor:"pointer",fontFamily:FONT}}>{l}</button>))}
      </div>
      <div style={{marginLeft:"auto",color:C.dim,fontSize:14,fontFamily:MONO}}>Total: <b style={{color:C.ink}}>{fmtKg(totW)}</b></div>
    </div>
    {msg&&<div style={{fontSize:12.5,color:busy?C.warn:C.dim,marginTop:10}}>{msg}</div>}
    <div style={{fontSize:11.5,color:C.dim,marginTop:6}}>Accepts the boat <b>.xlsx</b> directly, or CSV/PDF. Skips “*” summary and Total rows automatically.</div>
    <div style={{marginTop:16,overflowX:"auto"}}>
      <table><thead><tr><th>Species</th><th>Size</th>{mode==="boxes"?<><th className="r">Boxes</th><th className="r">Avg box</th><th className="r">Weight</th></>:<th className="r">Weight kg</th>}</tr></thead>
      <tbody>{tally.map((r)=>(<tr key={r.id}><td style={{fontWeight:600}}>{r.sp}</td><td style={{color:C.dim}}>{r.size}</td>
        {mode==="boxes"?<><td className="r"><input className="ci" style={{width:56}} type="number" value={r.boxes} onChange={(e)=>upd(r.id,"boxes",e.target.value)}/></td><td className="r"><input className="ci" style={{width:60}} type="number" step="0.1" value={r.avgBox} onChange={(e)=>upd(r.id,"avgBox",e.target.value)}/></td><td className="r" style={{color:C.dim,fontFamily:MONO}}>{(r.boxes*r.avgBox).toFixed(1)}</td></>
        :<td className="r"><input className="ci" style={{width:80}} type="number" step="0.01" value={r.wt} onChange={(e)=>upd(r.id,"wt",e.target.value)}/></td>}
      </tr>))}</tbody></table>
    </div>
    <NextBtn onClick={next}/>
  </Card>);
}

function PriceStep({which,title,accent,prices,setPrices,onUpload,busy,msg,next}){
  function setCell(sp,gr,v){setPrices((p)=>({...p,[sp]:{...p[sp],[gr]:v===""?null:+v}}));}
  return(<Card accent={accent}>
    <div style={{display:"flex",alignItems:"center",gap:12,flexWrap:"wrap"}}>
      <div style={{fontSize:18,fontWeight:700}}>{title}</div>
      <UploadBtn onUpload={onUpload} busy={busy} accent={accent} label="⬆ Upload sheet"/>
      {msg&&<div style={{fontSize:12.5,color:busy?C.warn:C.dim}}>{msg}</div>}
    </div>
    <div style={{marginTop:18,display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(260px,1fr))",gap:14}}>
      {Object.keys(prices).map((sp)=>(<div key={sp} style={{background:C.panel2,border:`1px solid ${C.line}`,borderRadius:9,padding:"10px 12px"}}>
        <div style={{fontWeight:700,fontSize:13.5,marginBottom:7}}>{sp}</div>
        <div style={{display:"flex",flexWrap:"wrap",gap:7}}>{Object.keys(prices[sp]).map((gr)=>(<div key={gr} style={{display:"flex",alignItems:"center",gap:4}}><span style={{fontSize:11,color:C.dim,fontFamily:MONO}}>{gr}</span><input className="ci" style={{width:56}} type="number" step="0.01" value={prices[sp][gr]??""} onChange={(e)=>setCell(sp,gr,e.target.value)}/></div>))}</div>
      </div>))}
    </div>
    <NextBtn onClick={next}/>
  </Card>);
}

function MapStep({tally,map,setMap,pd,dk,next}){
  const [editSp,setEditSp]=useState(false);
  function upd(i,f,v){setMap((p)=>p.map((m,idx)=>idx===i?{...m,[f]:v}:m));}
  const pdSp=Object.keys(pd).concat("—"),dkSp=Object.keys(dk).concat("—");
  const cc={high:C.good,med:C.warn,low:C.bad};
  const px=(obj,sp,k)=>(obj[sp]&&obj[sp][k]!=null?obj[sp][k]:null);
  return(<Card>
    <div style={{display:"flex",alignItems:"center",gap:10,flexWrap:"wrap"}}>
      <div style={{fontSize:18,fontWeight:700}}>Check grade mapping</div>
      <label style={{marginLeft:"auto",fontSize:12,color:C.dim,display:"flex",alignItems:"center",gap:6,cursor:"pointer"}}>
        <input type="checkbox" checked={editSp} onChange={(e)=>setEditSp(e.target.checked)}/>Edit species too
      </label>
    </div>
    <div style={{color:C.dim,fontSize:12.5,marginTop:4,display:"flex",gap:14,flexWrap:"wrap"}}>
      <span><Dot c={C.good}/> good</span><span><Dot c={C.warn}/> check</span><span><Dot c={C.bad}/> best-guess</span>
    </div>

    <div style={{marginTop:14,display:"grid",gap:10}}>
      {tally.map((r,i)=>{
        const m=map[i];
        const pg=(pd[m.pdSp]?Object.keys(pd[m.pdSp]):[]).concat("ANY","—");
        const ds=(dk[m.dkSp]?Object.keys(dk[m.dkSp]):[]).concat("—");
        const pdVal=px(pd,m.pdSp,m.pdGr), dkVal=px(dk,m.dkSp,m.dkSort);
        return(
          <div key={r.id} style={{background:C.panel2,border:`1px solid ${C.line}`,borderLeft:`4px solid ${cc[m.conf]}`,borderRadius:10,padding:"11px 13px"}}>
            {/* line header */}
            <div style={{display:"flex",alignItems:"baseline",gap:8}}>
              <span style={{fontWeight:800,fontSize:15}}>{r.sp}</span>
              <span style={{color:C.dim,fontSize:13}}>{r.size}</span>
              <span style={{marginLeft:"auto",fontSize:12,color:C.dim,fontFamily:MONO}}>{r.wt.toFixed(0)} kg</span>
            </div>
            {/* two columns: PD | DK */}
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginTop:9}}>
              <div style={{background:pdVal==null?"rgba(255,204,68,.08)":C.panel,borderRadius:8,padding:"8px 9px",borderTop:`2px solid ${pdVal==null?C.warn:C.pd}`,border:pdVal==null?`1px solid ${C.warn}`:"1px solid transparent"}}>
                <div style={{display:"flex",alignItems:"center",gap:6,marginBottom:5}}>
                  <span style={{fontSize:10.5,color:pdVal==null?C.warn:C.pd,fontWeight:700,letterSpacing:".04em"}}>PETERHEAD</span>
                  {pdVal==null&&<span style={{marginLeft:"auto",fontSize:9,fontWeight:800,color:"#06281a",background:C.warn,borderRadius:4,padding:"1px 5px",letterSpacing:".03em"}}>{dkVal!=null?"WILL SUB":"NO PRICE"}</span>}
                </div>
                {editSp&&<select className="si" style={{width:"100%",marginBottom:5,fontSize:12}} value={m.pdSp} onChange={(e)=>upd(i,"pdSp",e.target.value)}>{pdSp.map((s)=><option key={s}>{s}</option>)}</select>}
                <div style={{display:"flex",alignItems:"center",gap:8}}>
                  <select className="si" style={{flex:1,fontSize:14,padding:"7px 8px",color:pdVal==null?C.warn:C.ink}} value={m.pdGr} onChange={(e)=>upd(i,"pdGr",e.target.value)}>{pdVal==null&&!pg.includes(m.pdGr)&&<option value={m.pdGr}>{(GRADE_LABEL[m.pdGr]||m.pdGr)} · no price</option>}{[...new Set(pg)].map((g)=><option key={g} value={g}>{GRADE_LABEL[g]||g}</option>)}</select>
                  <span style={{fontSize:13,fontFamily:MONO,color:pdVal!=null?C.ink:C.warn,minWidth:46,textAlign:"right"}}>{pdVal!=null?fmtGBP(pdVal):(dkVal!=null?"("+fmtGBP(dkVal)+")":"—")}</span>
                </div>
              </div>
              <div style={{background:dkVal==null?"rgba(255,204,68,.08)":C.panel,borderRadius:8,padding:"8px 9px",borderTop:`2px solid ${dkVal==null?C.warn:C.dk}`,border:dkVal==null?`1px solid ${C.warn}`:"1px solid transparent"}}>
                <div style={{display:"flex",alignItems:"center",gap:6,marginBottom:5}}>
                  <span style={{fontSize:10.5,color:dkVal==null?C.warn:C.dk,fontWeight:700,letterSpacing:".04em"}}>HANSTHOLM</span>
                  {dkVal==null&&<span style={{marginLeft:"auto",fontSize:9,fontWeight:800,color:"#06281a",background:C.warn,borderRadius:4,padding:"1px 5px",letterSpacing:".03em"}}>{pdVal!=null?"WILL SUB":"NO PRICE"}</span>}
                </div>
                {editSp&&<select className="si" style={{width:"100%",marginBottom:5,fontSize:12}} value={m.dkSp} onChange={(e)=>upd(i,"dkSp",e.target.value)}>{dkSp.map((s)=><option key={s}>{s}</option>)}</select>}
                <div style={{display:"flex",alignItems:"center",gap:8}}>
                  <select className="si" style={{flex:1,fontSize:14,padding:"7px 8px",color:dkVal==null?C.warn:C.ink}} value={m.dkSort} onChange={(e)=>upd(i,"dkSort",e.target.value)}>{dkVal==null&&!ds.includes(m.dkSort)&&<option value={m.dkSort}>{m.dkSort==="—"?"— none —":m.dkSort+" · no price"}</option>}{[...new Set(ds)].map((g)=><option key={g}>{g}</option>)}</select>
                  <span style={{fontSize:13,fontFamily:MONO,color:dkVal!=null?C.ink:C.warn,minWidth:46,textAlign:"right"}}>{dkVal!=null?fmtGBP(dkVal):(pdVal!=null?"("+fmtGBP(pdVal)+")":"—")}</span>
                </div>
              </div>
            </div>
          </div>
        );
      })}
    </div>
    <NextBtn onClick={next} label="See gross →"/>
  </Card>);
}
function Dot({c}){return(<span style={{display:"inline-block",width:9,height:9,borderRadius:9,background:c,verticalAlign:"middle"}}/>);}

function ResultStep({rows,totals,summary,fillMissing,setFillMissing,tallyMode,exportCSV}){
  const winner=totals.diff>=0?"Hanstholm":"Peterhead";const wc=totals.diff>=0?C.dk:C.pd;
  const pct=Math.min(totals.pd,totals.dk)>0?(Math.abs(totals.diff)/Math.min(totals.pd,totals.dk)*100).toFixed(1):"0";
  return(<div style={{display:"grid",gap:20}}>
    <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(220px,1fr))",gap:16}}>
      <Big label="Peterhead gross" val={fmtGBP(totals.pd)} col={C.pd}/>
      <Big label="Hanstholm gross" val={fmtGBP(totals.dk)} col={C.dk}/>
      <Big label={`${winner} better by`} val={(totals.diff>=0?"+":"")+fmtGBP(totals.diff)} col={wc} sub={`${pct}% · ${fmtKg(totals.w)} landed`}/>
    </div>
    <Card>
      <div style={{display:"flex",alignItems:"center",gap:12,flexWrap:"wrap",marginBottom:10}}>
        <div style={{fontSize:16,fontWeight:700}}>By species</div>
        <button onClick={exportCSV} style={{marginLeft:"auto",background:C.ink,color:C.bg,border:"none",padding:"8px 16px",borderRadius:7,fontWeight:700,cursor:"pointer",fontFamily:FONT,fontSize:13}}>⬇ Export CSV</button>
      </div>
      <label style={{fontSize:12.5,color:C.dim,display:"flex",alignItems:"center",gap:7,cursor:"pointer",marginBottom:12}}><input type="checkbox" checked={fillMissing} onChange={(e)=>setFillMissing(e.target.checked)}/>Fill missing prices from other market</label>

      <div style={{display:"grid",gap:8}}>
        {summary.map((s)=>{
          const win=Math.abs(s.diff)<0.01?null:s.diff>0?C.dk:C.pd;
          return(
            <div key={s.sp} style={{background:C.panel2,border:`1px solid ${C.line}`,borderLeft:`4px solid ${win||C.line}`,borderRadius:9,padding:"10px 12px"}}>
              <div style={{display:"flex",alignItems:"baseline",gap:8}}>
                <span style={{fontWeight:800,fontSize:14.5}}>{s.sp}</span>
                <span style={{fontSize:12,color:C.dim,fontFamily:MONO}}>{s.w.toFixed(0)} kg</span>
                <span style={{marginLeft:"auto",fontSize:13,fontFamily:MONO,fontWeight:700,color:win||C.dim}}>{Math.abs(s.diff)<0.01?"tie":(s.diff>0?"+":"")+fmtGBP(s.diff)}</span>
              </div>
              <div style={{display:"flex",gap:16,marginTop:7}}>
                <div style={{flex:1}}><div style={{fontSize:10,color:C.pd,fontWeight:700,letterSpacing:".04em"}}>PETERHEAD</div><div style={{fontSize:14,fontFamily:MONO,marginTop:2}}>{fmtGBP(s.pd)}</div></div>
                <div style={{flex:1}}><div style={{fontSize:10,color:C.dk,fontWeight:700,letterSpacing:".04em"}}>HANSTHOLM</div><div style={{fontSize:14,fontFamily:MONO,marginTop:2}}>{fmtGBP(s.dk)}</div></div>
              </div>
            </div>
          );
        })}
      </div>

      <div style={{background:C.panel,border:`1px solid ${C.line}`,borderRadius:9,padding:"12px 14px",marginTop:12,display:"flex",alignItems:"center",gap:12,flexWrap:"wrap"}}>
        <span style={{fontWeight:800,fontSize:15}}>TOTAL</span>
        <span style={{fontSize:12,color:C.dim,fontFamily:MONO}}>{totals.w.toFixed(0)} kg</span>
        <div style={{marginLeft:"auto",display:"flex",gap:16,alignItems:"baseline"}}>
          <div><div style={{fontSize:10,color:C.pd,fontWeight:700}}>PD</div><div style={{fontFamily:MONO,fontWeight:700,color:C.pd}}>{fmtGBP(totals.pd)}</div></div>
          <div><div style={{fontSize:10,color:C.dk,fontWeight:700}}>DK</div><div style={{fontFamily:MONO,fontWeight:700,color:C.dk}}>{fmtGBP(totals.dk)}</div></div>
        </div>
      </div>
    </Card>
    {rows.some((r)=>r.note)&&<Card><div style={{fontSize:14,fontWeight:700,marginBottom:8,color:C.warn}}>Substituted / flagged lines</div><div style={{display:"grid",gap:5}}>{rows.filter((r)=>r.note).map((r)=>(<div key={r.id} style={{fontSize:12.5,color:C.dim,fontFamily:MONO}}><span style={{color:C.ink}}>{r.sp} {r.size}</span> — {r.note}</div>))}</div></Card>}
    <div style={{fontSize:12,color:C.dim,textAlign:"center"}}>Estimate only — uses {tallyMode==="boxes"?"boxes × avg box weight":"entered weights"} and the prices/mappings you confirmed. Not a settlement.</div>
  </div>);
}
function Big({label,val,col,sub}){return(<div style={{background:C.panel,border:`1px solid ${C.line}`,borderTop:`3px solid ${col}`,borderRadius:12,padding:"20px 22px"}}><div style={{fontSize:11.5,color:C.dim,textTransform:"uppercase",letterSpacing:".06em"}}>{label}</div><div style={{fontSize:30,fontWeight:700,color:col,marginTop:6,letterSpacing:"-.02em",fontFamily:MONO}}>{val}</div>{sub&&<div style={{fontSize:12,color:C.dim,marginTop:4}}>{sub}</div>}</div>);}
