import { useState, useEffect, useRef } from "react";
import * as XLSX from "xlsx";

/* ═══════════════════════════════════════════════════════════════
   SUPABASE
═══════════════════════════════════════════════════════════════ */
const SUPA_URL = "https://uicmfyudiullulbbwzmh.supabase.co";
const SUPA_KEY = "sb_publishable_9h9AyvXpkp9glLxDVWRuGw_1eKVS7sE";

async function supa(path, { method="GET", body, prefer }={}) {
  const r = await fetch(`${SUPA_URL}/rest/v1${path}`, {
    method,
    headers: {
      "apikey": SUPA_KEY,
      "Authorization": `Bearer ${SUPA_KEY}`,
      "Content-Type": "application/json",
      ...(prefer ? { "Prefer": prefer } : {}),
    },
    ...(body !== undefined ? { body: JSON.stringify(body) } : {}),
  });
  if (!r.ok) throw new Error(await r.text() || `HTTP ${r.status}`);
  const t = await r.text();
  return t ? JSON.parse(t) : null;
}

const API = {
  projects: {
    list:   ()       => supa("/projects?select=*&order=created_at.desc"),
    create: (d)      => supa("/projects", { method:"POST", body:d, prefer:"return=representation" }),
    update: (id,d)   => supa(`/projects?id=eq.${id}`, { method:"PATCH", body:d, prefer:"return=representation" }),
    remove: (id)     => supa(`/projects?id=eq.${id}`, { method:"DELETE" }),
  },
  reports: {
    forProject: (pid)    => supa(`/daily_reports?project_id=eq.${pid}&order=date.desc`),
    all:        ()       => supa("/daily_reports?select=*,projects(id,name)&order=date.desc&limit=300"),
    pending:    ()       => supa("/daily_reports?status=eq.submitted&select=*,projects(id,name)&order=created_at.desc"),
    create:     (d)      => supa("/daily_reports", { method:"POST", body:d, prefer:"return=representation" }),
    update:     (id,d)   => supa(`/daily_reports?id=eq.${id}`, { method:"PATCH", body:d, prefer:"return=representation" }),
    remove:     (id)     => supa(`/daily_reports?id=eq.${id}`, { method:"DELETE" }),
  },
  safety: {
    forProject: (pid) => supa(`/safety_logs?project_id=eq.${pid}&order=created_at.desc`),
    create:     (d)   => supa("/safety_logs", { method:"POST", body:d, prefer:"return=representation" }),
    remove:     (id)  => supa(`/safety_logs?id=eq.${id}`, { method:"DELETE" }),
  },
  photos: {
    forProject: (pid) => supa(`/project_photos?project_id=eq.${pid}&order=created_at.desc`),
    create:     (d)   => supa("/project_photos", { method:"POST", body:d, prefer:"return=representation" }),
    remove:     (id)  => supa(`/project_photos?id=eq.${id}`, { method:"DELETE" }),
  },
  timeCards: {
    forProject: (pid)       => supa(`/time_cards?project_id=eq.${pid}&order=date.desc,created_at.desc`),
    forDate:    (pid,date)  => supa(`/time_cards?project_id=eq.${pid}&date=eq.${date}&order=created_at`),
    create:     (d)         => supa("/time_cards", { method:"POST", body:d, prefer:"return=representation" }),
    remove:     (id)        => supa(`/time_cards?id=eq.${id}`, { method:"DELETE" }),
  },
  weather: {
    forProject: (pid) => supa(`/weather_logs?project_id=eq.${pid}&order=date.desc&limit=14`),
    upsert:     (d)   => supa("/weather_logs", { method:"POST", body:d, prefer:"return=representation,resolution=merge-duplicates" }),
    remove:     (id)  => supa(`/weather_logs?id=eq.${id}`, { method:"DELETE" }),
  },
  equipment: {
    forProject: (pid)      => supa(`/equipment_on_site?project_id=eq.${pid}&order=date.desc,created_at.desc`),
    create:     (d)        => supa("/equipment_on_site", { method:"POST", body:d, prefer:"return=representation" }),
    remove:     (id)       => supa(`/equipment_on_site?id=eq.${id}`, { method:"DELETE" }),
  },
  subs: {
    forProject: (pid) => supa(`/subcontractors?project_id=eq.${pid}&order=date.desc,created_at.desc`),
    create:     (d)   => supa("/subcontractors", { method:"POST", body:d, prefer:"return=representation" }),
    remove:     (id)  => supa(`/subcontractors?id=eq.${id}`, { method:"DELETE" }),
  },
};

/* ═══════════════════════════════════════════════════════════════
   THEME
═══════════════════════════════════════════════════════════════ */
const T = {
  bg:"#09090B", surface:"#111113", card:"#18181B", border:"#27272A",
  orange:"#F97316", orangeLow:"#F9731614", orangeMid:"#F9731630",
  green:"#22C55E",  greenLow:"#22C55E14",
  red:"#EF4444",    redLow:"#EF444414",
  yellow:"#EAB308", yellowLow:"#EAB30814",
  blue:"#3B82F6",   blueLow:"#3B82F614",
  purple:"#A855F7",
  text:"#FAFAFA", sub:"#A1A1AA", muted:"#52525B",
};
const inp = {
  width:"100%", boxSizing:"border-box", background:"#0C0C0F",
  border:`1px solid ${T.border}`, borderRadius:12, color:T.text,
  fontSize:15, padding:"13px 14px", outline:"none",
  fontFamily:"inherit", appearance:"none", WebkitAppearance:"none",
};
const lbl = { display:"block", fontSize:11, fontWeight:700, color:T.muted, letterSpacing:"1px", textTransform:"uppercase", marginBottom:6 };
const cardS = { background:T.card, border:`1px solid ${T.border}`, borderRadius:16, padding:"16px" };
const pill = (c) => ({ display:"inline-flex", alignItems:"center", background:c+"20", color:c, borderRadius:20, padding:"3px 10px", fontSize:11, fontWeight:700 });
const primBtn = { width:"100%", background:T.orange, color:"#09090B", border:"none", borderRadius:14, padding:"16px", fontSize:16, fontWeight:800, cursor:"pointer", fontFamily:"inherit", display:"flex", alignItems:"center", justifyContent:"center", gap:8 };
const ghostBtn = { background:"transparent", border:`1px solid ${T.border}`, borderRadius:12, padding:"12px 16px", color:T.sub, fontSize:14, cursor:"pointer", fontFamily:"inherit", fontWeight:600 };
const dangerBtn = { background:T.redLow, border:`1px solid ${T.red}30`, borderRadius:12, padding:"12px 16px", color:T.red, fontSize:14, cursor:"pointer", fontFamily:"inherit", fontWeight:600, width:"100%", textAlign:"center" };

/* ═══════════════════════════════════════════════════════════════
   DATA
═══════════════════════════════════════════════════════════════ */
const POSITIONS = [
  {name:"Project Manager",rate:64.50},{name:"Foreman",rate:63.25},
  {name:"Technician",rate:60.75},{name:"Inspector",rate:53.75},
  {name:"Certified Welder",rate:60.75},{name:"Fitter",rate:58.50},
  {name:"Mechanic",rate:58.50},{name:"Operator",rate:58.50},
  {name:"Truck Driver",rate:58.50},{name:"Helper (Welder)",rate:57.25},
  {name:"Laborer",rate:51.00},{name:"Foreman (Elect)",rate:82.25},
  {name:"Electrician",rate:82.25},{name:"Helper (Elect)",rate:45.50},
  {name:"Per Diem",rate:190.00,flat:true},
];
const NAMES = [
  "Aden Walter","Alan Fairbrother","Alan Robinson","Artie Pugh","Ben Torres",
  "Brandon Milano","Brandon Stach","Charles Dovel","Chris Utz","Christopher Dean",
  "Chuck Dean","Clay Lau","Connor Kestner","Derrick Spriggs","Eric Bowens",
  "Gary Czajkowski","Gavin Bowens","Gavin Kight","Howard Lau","James Gaskin",
  "Jeff Lau","Jeff White","Jesse Harbarger","Jessica Vance","John Baier",
  "John P. Cosner Jr.","Jordan Gorwell","Joseph Lau","Josh Gladhill","Julian Torres",
  "Kevin Gabrish","Kevin Holmes","Kurt Batterton","Leo Velez","Logan Robinson",
  "Mark Hamilton","Matt Kilby","Matt Pearlingi","Matthew Linton","Mike Gamble",
  "Mike McBride","Mike Meridith","Mike Peterson","Mike Seiler","Nate Nemire",
  "Owen Cross","Pat Gorman","Paul Howard","Rich Raborg","Robert Neslein",
  "Roland Long","Shane Hower","Steve Kestner","Theodoro Torres","Thomas Britton",
  "Tom Hatfield","Troy Strother","Tyrone Davis","Valentine Mbah","Walter Chicas-Luna",
  "Will Wychulis","Wyatt Gill",
].sort();
const EQUIP_LIST = [
  {section:"Trucks & Trailers"},
  {name:"Truck - 1 Ton",rate:21.50,unit:"Hours"},{name:"Truck - 3/4 Ton w/ Snow Plow",rate:350,unit:"Days"},
  {name:"Truck - 1/2 Ton",rate:18.50,unit:"Hours"},{name:"Truck - Boom (20-29 Ton)",rate:65,unit:"Hours"},
  {name:"Truck - Bucket",rate:45,unit:"Hours"},{name:"Truck - Dump Truck (3 Axle)",rate:35,unit:"Hours"},
  {name:"Truck - Haul Truck - No Trailer",rate:70,unit:"Hours"},{name:"Truck - Tru-Vac",rate:13500,unit:"Month"},
  {name:"Truck - Welding Rig",rate:35,unit:"Hours"},{name:"Trailer - Electrical - Colonial",rate:147,unit:"Month"},
  {name:"Trailer - Lowboy - 2 Axle",rate:28,unit:"Hours"},{name:"Trailer - Tag Along",rate:50,unit:"Days"},
  {name:"Trailer - Tool Trailer - 18-25'",rate:175,unit:"Days"},{name:"Trailer - Tool Trailer - 26-40'",rate:200,unit:"Days"},
  {section:"Earthmoving & ROW"},
  {name:"ATV - 4 Wheel",rate:125,unit:"Days"},{name:"Backhoe Loader - 80-105 HP",rate:62.45,unit:"Hours"},
  {name:"Excavator - Mini - 2-8K LB",rate:299,unit:"Days"},{name:"Excavator - Mini - 9K LB",rate:335,unit:"Days"},
  {name:"Excavator - Mini - 12-16K LB",rate:475,unit:"Days"},{name:"Excavator - Small - 21-29K LB",rate:565,unit:"Days"},
  {name:"Excavator - Small - 30-33K LB",rate:632,unit:"Days"},{name:"Excavator - Medium - 48-55K LB",rate:852,unit:"Days"},
  {name:"Excavator - Large - 80-89K LB",rate:1050,unit:"Days"},{name:"Excavator - Large - 90-119K LB",rate:1350,unit:"Days"},
  {name:"Skidsteer Loader - 70-80 HP",rate:440,unit:"Days"},{name:"Skidsteer Loader - 81-100 HP",rate:475,unit:"Days"},
  {name:"Tractor - 50 HP 4x4 w/ Bush Hog",rate:36.50,unit:"Hours"},{name:"Mower - Riding/Zero Turn",rate:175,unit:"Days"},
  {section:"Air, Compressors & Blast"},
  {name:"Air Compressor - 185 CFM",rate:195,unit:"Days"},{name:"Air Compressor - 375 CFM",rate:275,unit:"Days"},
  {name:"Air Impact Wrench - 1\"",rate:50,unit:"Days"},{name:"Air Spade / Knife",rate:55,unit:"Days"},
  {name:"Blast Rig - 4 Bag Pot w/ 185 CFM AC",rate:55.50,unit:"Hours"},{name:"Blast Rig - 1 Pot w/ 375 CFM AC",rate:500,unit:"Days"},
  {section:"Testing & Misc. Tools"},
  {name:"Holiday Detector / Pipe Jeep",rate:72,unit:"Days"},{name:"Hydraulic Torque",rate:200,unit:"Days"},
  {name:"Hydro Test Pump",rate:60,unit:"Days"},{name:"Hydrotest - High Pressure",rate:3800,unit:"Days"},
  {name:"Jack Hammer",rate:72,unit:"Days"},{name:"LEL/Gas Monitor - 4 Gas",rate:50,unit:"Days"},
  {name:"Line Locator",rate:50,unit:"Days"},{name:"HEPA Vacuum",rate:100,unit:"Days"},
  {name:"Torque Wrench w/Sockets Hyd/Pneu",rate:195,unit:"Days"},{name:"Pipe Beveling Machine 16-22\"",rate:100,unit:"Days"},
];
const WMO = {
  0:["Clear Sky","☀️"],1:["Mainly Clear","🌤️"],2:["Partly Cloudy","⛅"],3:["Overcast","☁️"],
  45:["Foggy","🌫️"],48:["Icy Fog","🌫️"],
  51:["Light Drizzle","🌦️"],53:["Drizzle","🌦️"],55:["Heavy Drizzle","🌦️"],
  61:["Light Rain","🌧️"],63:["Rain","🌧️"],65:["Heavy Rain","🌧️"],
  71:["Light Snow","🌨️"],73:["Snow","🌨️"],75:["Heavy Snow","❄️"],77:["Snow Grains","❄️"],
  80:["Light Showers","🌦️"],81:["Showers","🌦️"],82:["Violent Showers","⛈️"],
  85:["Snow Showers","🌨️"],86:["Heavy Snow Showers","🌨️"],
  95:["Thunderstorm","⛈️"],96:["Thunderstorm + Hail","⛈️"],99:["Severe Thunderstorm","⛈️"],
};

/* ═══════════════════════════════════════════════════════════════
   HELPERS
═══════════════════════════════════════════════════════════════ */
const uid       = () => Math.random().toString(36).slice(2,9);
const today     = () => new Date().toISOString().split("T")[0];
const fmt       = (n) => Number(n||0).toLocaleString("en-US",{minimumFractionDigits:2,maximumFractionDigits:2});
const fmtDate   = (d) => d ? new Date(d+"T12:00:00").toLocaleDateString("en-US",{month:"short",day:"numeric",year:"numeric"}) : "—";
const fmtShort  = (d) => d ? new Date(d+"T12:00:00").toLocaleDateString("en-US",{weekday:"short",month:"short",day:"numeric"}) : "—";

function laborAmt(r) {
  const p = POSITIONS.find(x=>x.name===r.classification);
  if (!p) return 0;
  if (p.flat) return p.rate;
  return p.rate*((parseFloat(r.regHrs)||0)+(parseFloat(r.otHrs)||0)*1.5+(parseFloat(r.travelHrs)||0));
}
function equipAmt(r) { return (parseFloat(r.rate)||0)*(parseFloat(r.qty)||0)*(parseFloat(r.usage)||0); }
function reportTotals(r) {
  const labor=(r.labor||[]).reduce((s,x)=>s+laborAmt(x),0);
  const equip=(r.equipment||[]).reduce((s,x)=>s+equipAmt(x),0);
  const mats=(r.materials||[]).reduce((s,x)=>s+(parseFloat(x.amount)||0),0);
  return { labor, equip, mats, grand:labor+equip+mats };
}
function calcHours(ci, co) {
  if (!ci||!co) return 0;
  const [ih,im]=ci.split(":").map(Number);
  const [oh,om]=co.split(":").map(Number);
  const diff=(oh*60+om)-(ih*60+im);
  return diff>0 ? Math.round(diff/60*100)/100 : 0;
}
function getWeekStart() {
  const d=new Date(); const day=d.getDay();
  d.setDate(d.getDate()-(day===0?6:day-1));
  return d.toISOString().split("T")[0];
}
async function compressImg(file,maxW=900,q=0.65) {
  return new Promise(res=>{
    const rd=new FileReader();
    rd.onload=ev=>{
      const img=new Image();
      img.onload=()=>{
        const sc=Math.min(1,maxW/img.width);
        const c=document.createElement("canvas");
        c.width=Math.round(img.width*sc); c.height=Math.round(img.height*sc);
        c.getContext("2d").drawImage(img,0,0,c.width,c.height);
        res(c.toDataURL("image/jpeg",q));
      };
      img.src=ev.target.result;
    };
    rd.readAsDataURL(file);
  });
}
async function fetchWeather(location) {
  const geoR=await fetch(`https://geocoding-api.open-meteo.com/v1/search?name=${encodeURIComponent(location)}&count=1&language=en&format=json`);
  const geoD=await geoR.json();
  if (!geoD.results?.length) throw new Error(`Could not find location: "${location}"`);
  const {latitude:lat,longitude:lon,name,admin1}=geoD.results[0];
  const wR=await fetch(`https://api.open-meteo.com/v1/forecast?latitude=${lat}&longitude=${lon}&current=temperature_2m,apparent_temperature,weathercode,windspeed_10m,precipitation,relativehumidity_2m&daily=temperature_2m_max,temperature_2m_min,precipitation_sum,weathercode&timezone=auto&temperature_unit=fahrenheit&windspeed_unit=mph&precipitation_unit=inch&forecast_days=1`);
  const wD=await wR.json();
  return { ...wD, locationName:`${name}, ${admin1}` };
}

/* ═══════════════════════════════════════════════════════════════
   SHARED UI
═══════════════════════════════════════════════════════════════ */
function Spinner() {
  return (
    <div style={{display:"flex",justifyContent:"center",padding:"48px 0"}}>
      <div style={{width:32,height:32,border:`3px solid ${T.border}`,borderTopColor:T.orange,borderRadius:"50%",animation:"spin 0.7s linear infinite"}}/>
      <style>{`@keyframes spin{to{transform:rotate(360deg)}}`}</style>
    </div>
  );
}
function ErrBanner({msg,onDismiss}) {
  if (!msg) return null;
  return (
    <div style={{background:T.redLow,border:`1px solid ${T.red}40`,borderRadius:12,padding:"12px 14px",marginBottom:14,display:"flex",justifyContent:"space-between",alignItems:"center"}}>
      <span style={{fontSize:13,color:T.red}}>⚠️ {msg}</span>
      <button onClick={onDismiss} style={{background:"none",border:"none",color:T.red,cursor:"pointer",fontSize:18,padding:"0 0 0 10px"}}>×</button>
    </div>
  );
}
function TopBar({title,sub,onBack,right,noBorder}) {
  return (
    <div style={{background:T.surface,borderBottom:noBorder?"none":`1px solid ${T.border}`,padding:"14px 16px",position:"sticky",top:0,zIndex:50}}>
      {onBack&&<button onClick={onBack} style={{background:"none",border:"none",color:T.sub,fontSize:13,cursor:"pointer",marginBottom:8,padding:0,fontFamily:"inherit"}}>← Back</button>}
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start"}}>
        <div style={{flex:1,minWidth:0}}>
          <div style={{fontSize:20,fontWeight:900,color:T.text,letterSpacing:"-0.5px"}}>{title}</div>
          {sub&&<div style={{fontSize:12,color:T.muted,marginTop:2}}>{sub}</div>}
        </div>
        {right&&<div style={{flexShrink:0,marginLeft:12}}>{right}</div>}
      </div>
    </div>
  );
}
function Lightbox({src,onClose}) {
  if (!src) return null;
  return (
    <div onClick={onClose} style={{position:"fixed",inset:0,zIndex:999,background:"rgba(0,0,0,0.95)",display:"flex",alignItems:"center",justifyContent:"center",padding:16}}>
      <img src={src} alt="" style={{maxWidth:"100%",maxHeight:"90vh",borderRadius:12}} onClick={e=>e.stopPropagation()}/>
      <button onClick={onClose} style={{position:"absolute",top:16,right:16,background:"#333",border:"none",color:"#fff",borderRadius:"50%",width:36,height:36,fontSize:18,cursor:"pointer"}}>×</button>
    </div>
  );
}
function DashedAdd({label,onClick,color}) {
  const c=color||T.muted;
  return (
    <button onClick={onClick} style={{width:"100%",border:`2px dashed ${c}50`,background:c+"08",color:c,borderRadius:14,padding:"14px",fontSize:15,cursor:"pointer",fontFamily:"inherit",fontWeight:600}}>
      {label}
    </button>
  );
}
function StatBar({items}) {
  return (
    <div style={{display:"grid",gridTemplateColumns:`repeat(${items.length},1fr)`,gap:8}}>
      {items.map(({label,val,color})=>(
        <div key={label} style={{background:T.card,border:`1px solid ${T.border}`,borderRadius:12,padding:"10px 8px",textAlign:"center"}}>
          <div style={{fontSize:16,fontWeight:900,color:color||T.text}}>{val}</div>
          <div style={{fontSize:9,color:T.muted,textTransform:"uppercase",letterSpacing:"0.7px",marginTop:2}}>{label}</div>
        </div>
      ))}
    </div>
  );
}

/* ═══════════════════════════════════════════════════════════════
   FORM CARDS
═══════════════════════════════════════════════════════════════ */
function LaborCard({row,onChange,onRemove}) {
  const pos=POSITIONS.find(p=>p.name===row.classification);
  const amt=laborAmt(row);
  const set=(k,v)=>{const u={...row,[k]:v};if(k==="classification"){const p=POSITIONS.find(x=>x.name===v);u.rate=p?p.rate:"";}onChange(u);};
  return (
    <div style={{...cardS,marginBottom:10,borderLeft:`3px solid ${T.orange}`}}>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8,marginBottom:8}}>
        <div style={{gridColumn:"1/-1"}}><label style={lbl}>Name</label>
          <select value={row.name||""} onChange={e=>set("name",e.target.value)} style={inp}>
            <option value="">— Select —</option>{NAMES.map(n=><option key={n}>{n}</option>)}
          </select></div>
        <div style={{gridColumn:"1/-1"}}><label style={lbl}>Classification</label>
          <select value={row.classification||""} onChange={e=>set("classification",e.target.value)} style={inp}>
            <option value="">— Select —</option>{POSITIONS.map(p=><option key={p.name}>{p.name}</option>)}
          </select></div>
      </div>
      {pos&&!pos.flat&&(
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:8,marginBottom:8}}>
          {[["regHrs","Reg Hrs"],["otHrs","OT Hrs"],["travelHrs","Travel"]].map(([k,l])=>(
            <div key={k}><label style={lbl}>{l}</label>
              <input type="number" min="0" step="0.5" placeholder="0" value={row[k]||""} onChange={e=>set(k,e.target.value)} style={inp}/>
            </div>
          ))}
        </div>
      )}
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",paddingTop:8,borderTop:`1px solid ${T.border}`}}>
        <span style={{fontSize:11,color:T.muted}}>{pos?`$${pos.rate.toFixed(2)}${pos.flat?" flat":"/hr"}`:""}</span>
        <div style={{display:"flex",alignItems:"center",gap:10}}>
          {amt>0&&<span style={{fontSize:16,fontWeight:800,color:T.green}}>${fmt(amt)}</span>}
          <button onClick={onRemove} style={{background:"none",border:"none",color:T.red,cursor:"pointer",fontSize:20,padding:0}}>×</button>
        </div>
      </div>
    </div>
  );
}
function EquipCard({row,onChange,onRemove}) {
  const eq=EQUIP_LIST.find(e=>!e.section&&e.name===row.description);
  const amt=equipAmt(row);
  const set=(k,v)=>{const u={...row,[k]:v};if(k==="description"){const e=EQUIP_LIST.find(x=>!x.section&&x.name===v);u.rate=e?e.rate:"";u.unit=e?e.unit:"";}onChange(u);};
  return (
    <div style={{...cardS,marginBottom:10,borderLeft:`3px solid ${T.yellow}`}}>
      <div style={{marginBottom:8}}><label style={lbl}>Equipment</label>
        <select value={row.description||""} onChange={e=>set("description",e.target.value)} style={inp}>
          <option value="">— Select —</option>
          {EQUIP_LIST.map((e,i)=>e.section?<option key={i} disabled>── {e.section} ──</option>:<option key={i} value={e.name}>{e.name}</option>)}
        </select>
      </div>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8,marginBottom:8}}>
        <div><label style={lbl}>Qty</label><input type="number" min="0" placeholder="0" value={row.qty||""} onChange={e=>set("qty",e.target.value)} style={inp}/></div>
        <div><label style={lbl}>{eq?eq.unit:"Hrs / Days"}</label><input type="number" min="0" step="0.5" placeholder="0" value={row.usage||""} onChange={e=>set("usage",e.target.value)} style={inp}/></div>
      </div>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",paddingTop:8,borderTop:`1px solid ${T.border}`}}>
        <span style={{fontSize:11,color:T.muted}}>{eq?`$${eq.rate.toLocaleString()}/${eq.unit}`:""}</span>
        <div style={{display:"flex",alignItems:"center",gap:10}}>
          {amt>0&&<span style={{fontSize:16,fontWeight:800,color:T.green}}>${fmt(amt)}</span>}
          <button onClick={onRemove} style={{background:"none",border:"none",color:T.red,cursor:"pointer",fontSize:20,padding:0}}>×</button>
        </div>
      </div>
    </div>
  );
}
function MatCard({row,onChange,onRemove}) {
  const fileRef=useRef(null);
  const receipts=row.receipts||[];
  async function handleFiles(files){
    const n=[];
    for(const f of files){if(!f.type.startsWith("image/"))continue;const src=await compressImg(f,800,0.6);n.push({id:uid(),src});}
    onChange({...row,receipts:[...receipts,...n]});
  }
  return (
    <div style={{...cardS,marginBottom:10,borderLeft:`3px solid ${T.blue}`}}>
      <div style={{display:"grid",gridTemplateColumns:"56px 1fr 88px",gap:8,marginBottom:10}}>
        <div><label style={lbl}>Qty</label><input type="number" min="0" placeholder="0" value={row.qty||""} onChange={e=>onChange({...row,qty:e.target.value})} style={inp}/></div>
        <div><label style={lbl}>Description</label><input type="text" placeholder="Item / material" value={row.description||""} onChange={e=>onChange({...row,description:e.target.value})} style={inp}/></div>
        <div><label style={lbl}>Amount</label><input type="number" min="0" placeholder="0.00" value={row.amount||""} onChange={e=>onChange({...row,amount:e.target.value})} style={inp}/></div>
      </div>
      <div style={{borderTop:`1px solid ${T.border}`,paddingTop:10}}>
        <label style={{...lbl,marginBottom:8}}>📎 Receipts</label>
        <div style={{display:"flex",gap:8,flexWrap:"wrap",alignItems:"center"}}>
          {receipts.map(r=>(
            <div key={r.id} style={{position:"relative"}}>
              <img src={r.src} alt="" style={{width:60,height:60,objectFit:"cover",borderRadius:10,border:`2px solid ${T.blue}40`,display:"block"}}/>
              <button onClick={()=>onChange({...row,receipts:receipts.filter(x=>x.id!==r.id)})} style={{position:"absolute",top:-5,right:-5,width:18,height:18,borderRadius:"50%",background:T.red,border:"none",color:"#fff",fontSize:11,cursor:"pointer",display:"flex",alignItems:"center",justifyContent:"center"}}>×</button>
            </div>
          ))}
          <button onClick={()=>fileRef.current?.click()} style={{width:60,height:60,borderRadius:10,border:`2px dashed ${T.blue}40`,background:T.blueLow,color:T.blue,cursor:"pointer",display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",fontSize:18,gap:2}}>
            <span>📷</span><span style={{fontSize:9,fontWeight:700}}>ADD</span>
          </button>
          <input ref={fileRef} type="file" accept="image/*" capture="environment" multiple style={{display:"none"}} onChange={e=>{handleFiles(Array.from(e.target.files));e.target.value="";}}/>
        </div>
      </div>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginTop:10}}>
        {(row.amount>0||row.description)&&<span style={{fontSize:14,fontWeight:700,color:T.green}}>${fmt(parseFloat(row.amount)||0)}</span>}
        <button onClick={onRemove} style={{background:"none",border:"none",color:T.red,cursor:"pointer",fontSize:12,fontWeight:700,fontFamily:"inherit",marginLeft:"auto"}}>Remove</button>
      </div>
    </div>
  );
}

/* ═══════════════════════════════════════════════════════════════
   LOGIN
═══════════════════════════════════════════════════════════════ */
function LoginScreen({onLogin}) {
  const [name,setName]=useState(""); const [role,setRole]=useState("crew"); const [pin,setPin]=useState("");
  function go(){if(!name)return;if(role==="pm"&&pin!=="1234"){alert("Incorrect PIN");return;}onLogin({name:name.trim(),role});}
  return (
    <div style={{minHeight:"100vh",background:T.bg,display:"flex",flexDirection:"column",justifyContent:"center",padding:24,fontFamily:"inherit"}}>
      <div style={{textAlign:"center",marginBottom:40}}>
        <div style={{display:"inline-flex",alignItems:"center",gap:12,marginBottom:10}}>
          <div style={{width:56,height:56,background:T.orange,borderRadius:18,display:"flex",alignItems:"center",justifyContent:"center",fontSize:28,fontWeight:900,color:"#09090B"}}>A</div>
          <div style={{textAlign:"left"}}>
            <div style={{fontSize:26,fontWeight:900,color:T.text,letterSpacing:"-1px",lineHeight:1.1}}>AIME Field OS</div>
            <div style={{fontSize:11,color:T.muted,letterSpacing:"2px",textTransform:"uppercase"}}>Colonial Pipeline</div>
          </div>
        </div>
        <div style={{fontSize:13,color:T.muted}}>Daily reporting · Time cards · Safety · PM dashboard</div>
      </div>
      <div style={{...cardS,maxWidth:400,margin:"0 auto",width:"100%"}}>
        <div style={{marginBottom:14}}><label style={lbl}>Your Name</label>
          <select value={name} onChange={e=>setName(e.target.value)} style={inp}>
            <option value="">— Select your name —</option>
            {NAMES.map(n=><option key={n}>{n}</option>)}
          </select>
        </div>
        <div style={{marginBottom:14}}><label style={lbl}>I am a…</label>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8}}>
            {["crew","pm"].map(r=>(
              <button key={r} onClick={()=>setRole(r)} style={{padding:"14px",borderRadius:12,border:`2px solid ${role===r?T.orange:T.border}`,background:role===r?T.orangeLow:T.surface,color:role===r?T.orange:T.sub,fontWeight:700,fontSize:14,cursor:"pointer",fontFamily:"inherit"}}>
                {r==="crew"?"👷 Field Crew":"📋 Project Manager"}
              </button>
            ))}
          </div>
        </div>
        {role==="pm"&&<div style={{marginBottom:14}}><label style={lbl}>PM PIN</label><input type="password" maxLength={4} placeholder="4-digit PIN (default: 1234)" value={pin} onChange={e=>setPin(e.target.value)} style={inp}/></div>}
        <button onClick={go} style={{...primBtn,opacity:name?1:0.45}}>Sign In →</button>
        {role==="pm"&&<p style={{fontSize:11,color:T.muted,textAlign:"center",marginTop:8}}>Default PIN: 1234</p>}
      </div>
    </div>
  );
}

/* ═══════════════════════════════════════════════════════════════
   PROJECTS HOME
═══════════════════════════════════════════════════════════════ */
function ProjectsHome({user,projects,loading,onSelect,onNew,onLogout,onDash}) {
  const active=projects.filter(p=>p.status==="active");
  const archived=projects.filter(p=>p.status!=="active");
  const totalBilled=projects.reduce((s,p)=>s+(p._billed||0),0);
  const totalReports=projects.reduce((s,p)=>s+(p._reports||0),0);
  return (
    <div style={{background:T.bg,minHeight:"100vh",fontFamily:"inherit"}}>
      <div style={{background:T.surface,borderBottom:`1px solid ${T.border}`,padding:"16px"}}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:14}}>
          <div style={{display:"flex",alignItems:"center",gap:10}}>
            <div style={{width:36,height:36,background:T.orange,borderRadius:10,display:"flex",alignItems:"center",justifyContent:"center",fontWeight:900,fontSize:18,color:"#09090B"}}>A</div>
            <div><div style={{fontSize:16,fontWeight:800,color:T.text}}>AIME Field OS</div><div style={{fontSize:11,color:T.muted}}>👋 {user.name} · {user.role==="pm"?"Project Manager":"Field Crew"}</div></div>
          </div>
          <button onClick={onLogout} style={{...ghostBtn,padding:"8px 12px",fontSize:12}}>Sign Out</button>
        </div>
        <StatBar items={[
          {label:"Projects",val:active.length,color:T.orange},
          {label:"Reports",val:totalReports,color:T.green},
          {label:"Billed",val:"$"+(totalBilled>=1000?(totalBilled/1000).toFixed(1)+"k":fmt(totalBilled)),color:T.blue},
        ]}/>
      </div>
      <div style={{padding:"14px 16px 100px"}}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:12}}>
          <span style={{fontSize:12,fontWeight:700,color:T.muted,textTransform:"uppercase",letterSpacing:"1px"}}>Active Projects</span>
          {user.role==="pm"&&<button onClick={onNew} style={{background:T.orange,color:"#09090B",border:"none",borderRadius:10,padding:"8px 14px",fontSize:13,fontWeight:800,cursor:"pointer",fontFamily:"inherit"}}>+ New Project</button>}
        </div>
        {loading&&<Spinner/>}
        {!loading&&active.length===0&&<div style={{textAlign:"center",padding:"40px 0",color:T.muted}}><div style={{fontSize:36,marginBottom:8}}>🏗️</div><div>{user.role==="pm"?"No projects yet. Tap + New Project.":"No active projects. Check back soon."}</div></div>}
        {active.map(p=><ProjectCard key={p.id} p={p} onSelect={onSelect}/>)}
        {archived.length>0&&<>
          <div style={{fontSize:12,fontWeight:700,color:T.muted,textTransform:"uppercase",letterSpacing:"1px",marginTop:24,marginBottom:10}}>Archived</div>
          {archived.map(p=><ProjectCard key={p.id} p={p} onSelect={onSelect} dim/>)}
        </>}
      </div>
      {user.role==="pm"&&(
        <div style={{position:"fixed",bottom:0,left:"50%",transform:"translateX(-50%)",width:"100%",maxWidth:480,background:T.surface,borderTop:`1px solid ${T.border}`,padding:"10px 16px",display:"flex",gap:8,zIndex:50}}>
          <button style={{flex:1,background:T.card,border:`1px solid ${T.border}`,borderRadius:12,padding:"13px",color:T.text,fontSize:13,fontWeight:700,cursor:"pointer",fontFamily:"inherit"}}>🏗️ Projects</button>
          <button onClick={onDash} style={{flex:1,background:T.orange,border:"none",borderRadius:12,padding:"13px",color:"#09090B",fontSize:13,fontWeight:800,cursor:"pointer",fontFamily:"inherit"}}>📊 PM Dashboard</button>
        </div>
      )}
    </div>
  );
}
function ProjectCard({p,onSelect,dim}) {
  return (
    <div onClick={()=>onSelect(p)} style={{...cardS,marginBottom:10,cursor:"pointer",borderLeft:`3px solid ${dim?T.border:T.orange}`,opacity:dim?0.55:1,display:"flex",justifyContent:"space-between",alignItems:"center"}}>
      <div style={{flex:1,minWidth:0}}>
        <div style={{fontSize:15,fontWeight:800,color:T.text}}>{p.name}</div>
        <div style={{fontSize:12,color:T.sub,marginTop:2}}>{p.client||"No client"}{p.location?" · "+p.location:""}</div>
        <div style={{display:"flex",gap:6,marginTop:7,flexWrap:"wrap"}}>
          <span style={pill(T.orange)}>{p._reports||0} report{p._reports!==1?"s":""}</span>
          {p.afe&&<span style={pill(T.muted)}>AFE: {p.afe}</span>}
          <span style={pill(p.status==="active"?T.green:T.muted)}>{p.status}</span>
        </div>
      </div>
      <div style={{textAlign:"right",flexShrink:0,marginLeft:12}}>
        <div style={{fontSize:18,fontWeight:900,color:T.green}}>${(p._billed||0)>=1000?((p._billed||0)/1000).toFixed(1)+"k":fmt(p._billed||0)}</div>
        <div style={{fontSize:9,color:T.muted,textTransform:"uppercase",marginTop:2}}>Billed</div>
      </div>
    </div>
  );
}

/* ═══════════════════════════════════════════════════════════════
   PROJECT FORM (new / edit)
═══════════════════════════════════════════════════════════════ */
function ProjectForm({initial,onSave,onCancel,saving}) {
  const [f,setF]=useState(initial||{name:"",client:"",location:"",afe:"",work_order:"",start_date:today(),notes:"",status:"active"});
  const set=(k,v)=>setF(x=>({...x,[k]:v}));
  return (
    <div style={{background:T.bg,minHeight:"100vh",fontFamily:"inherit"}}>
      <TopBar title={initial?"Edit Project":"New Project"} onBack={onCancel}/>
      <div style={{padding:"16px 16px 100px"}}>
        {[{k:"name",l:"Project Name *",ph:"e.g. HDD Crossing – Station 42"},{k:"client",l:"Client",ph:"Colonial Pipeline"},{k:"location",l:"Location",ph:"City, State or Milepost"},{k:"afe",l:"AFE No.",ph:"AFE #"},{k:"work_order",l:"Work Order #",ph:"WO #"}].map(({k,l,ph})=>(
          <div key={k} style={{marginBottom:12}}><label style={lbl}>{l}</label><input type="text" placeholder={ph} value={f[k]||""} onChange={e=>set(k,e.target.value)} style={inp}/></div>
        ))}
        <div style={{marginBottom:12}}><label style={lbl}>Start Date</label><input type="date" value={f.start_date||today()} onChange={e=>set("start_date",e.target.value)} style={inp}/></div>
        <div style={{marginBottom:20}}><label style={lbl}>Notes</label><textarea placeholder="Project notes, scope, special instructions…" value={f.notes||""} onChange={e=>set("notes",e.target.value)} rows={3} style={{...inp,resize:"vertical",lineHeight:1.5}}/></div>
        <button onClick={()=>f.name.trim()&&!saving&&onSave(f)} style={{...primBtn,opacity:f.name.trim()&&!saving?1:0.5}}>
          {saving?"Saving…":initial?"Save Changes":"Create Project"}
        </button>
      </div>
    </div>
  );
}

/* ═══════════════════════════════════════════════════════════════
   DAILY REPORT FORM
═══════════════════════════════════════════════════════════════ */
const RSTEPS=["Job Info","Labor","Equipment","Materials","Review"];
function DailyReportForm({user,project,onSave,onCancel}) {
  const [step,setStep]=useState(1); const [saving,setSaving]=useState(false);
  const [rpt,setRpt]=useState({date:today(),description:"",report_no:"",labor:[],equipment:[],materials:[]});
  const topRef=useRef(null);
  const setR=(k,v)=>setRpt(r=>({...r,[k]:v}));
  function addLabor(){setR("labor",[...rpt.labor,{id:uid(),name:"",classification:"",regHrs:"",otHrs:"",travelHrs:""}]);}
  function addEquip(){setR("equipment",[...rpt.equipment,{id:uid(),description:"",qty:"",usage:"",rate:"",unit:""}]);}
  function addMat(){setR("materials",[...rpt.materials,{id:uid(),qty:"",description:"",amount:"",receipts:[]}]);}
  function upd(key,i,row){const a=[...rpt[key]];a[i]=row;setR(key,a);}
  function del(key,i){setR(key,rpt[key].filter((_,j)=>j!==i));}
  const tot=reportTotals(rpt);
  async function submit(){setSaving(true);await onSave({...rpt,submitted_by:user.name,status:"submitted"});setSaving(false);}
  const scroll=()=>topRef.current?.scrollIntoView({behavior:"smooth"});
  return (
    <div ref={topRef} style={{background:T.bg,minHeight:"100vh",fontFamily:"inherit"}}>
      <div style={{background:T.surface,borderBottom:`1px solid ${T.border}`,padding:"14px 16px",position:"sticky",top:0,zIndex:50}}>
        <div style={{display:"flex",justifyContent:"space-between",marginBottom:12}}>
          <div style={{fontSize:16,fontWeight:800}}>New Daily Report</div>
          <button onClick={onCancel} style={{background:"none",border:"none",color:T.sub,cursor:"pointer",fontSize:13,fontFamily:"inherit"}}>Cancel</button>
        </div>
        <div style={{display:"flex",alignItems:"center"}}>
          {RSTEPS.map((s,i)=>(
            <div key={i} style={{display:"flex",alignItems:"center",flex:i<RSTEPS.length-1?1:undefined}}>
              <div style={{display:"flex",flexDirection:"column",alignItems:"center",gap:3}}>
                <div style={{width:26,height:26,borderRadius:"50%",background:i+1<step?T.green:i+1===step?T.orange:T.border,color:i+1<=step?"#09090B":T.muted,display:"flex",alignItems:"center",justifyContent:"center",fontSize:11,fontWeight:800}}>{i+1<step?"✓":i+1}</div>
                <div style={{fontSize:8,color:i+1===step?T.orange:T.muted,fontWeight:i+1===step?700:400,whiteSpace:"nowrap"}}>{s}</div>
              </div>
              {i<RSTEPS.length-1&&<div style={{flex:1,height:2,background:i+1<step?T.green:T.border,margin:"0 3px",marginBottom:14}}/>}
            </div>
          ))}
        </div>
      </div>
      <div style={{padding:"14px 16px 100px"}}>
        {step===1&&(
          <div>
            <div style={{...cardS,marginBottom:14,borderLeft:`3px solid ${T.orange}`}}>
              <div style={{fontSize:12,color:T.muted,marginBottom:3}}>Project</div>
              <div style={{fontSize:15,fontWeight:700}}>{project.name}</div>
              {project.afe&&<div style={{fontSize:12,color:T.sub}}>AFE: {project.afe}{project.work_order?" · WO: "+project.work_order:""}</div>}
            </div>
            <div style={{marginBottom:12}}><label style={lbl}>Report Date</label><input type="date" value={rpt.date} onChange={e=>setR("date",e.target.value)} style={inp}/></div>
            <div style={{marginBottom:12}}><label style={lbl}>Report No.</label><input type="text" placeholder="Report #" value={rpt.report_no||""} onChange={e=>setR("report_no",e.target.value)} style={inp}/></div>
            <div style={{marginBottom:12}}><label style={lbl}>Description of Work Done</label><textarea placeholder="Describe the work performed today…" value={rpt.description||""} onChange={e=>setR("description",e.target.value)} rows={4} style={{...inp,resize:"vertical",lineHeight:1.5}}/></div>
          </div>
        )}
        {step===2&&(
          <div>
            <div style={{display:"flex",justifyContent:"space-between",marginBottom:12}}><div style={{fontSize:17,fontWeight:800}}>👷 Labor</div>{tot.labor>0&&<div style={{fontSize:16,fontWeight:800,color:T.green}}>${fmt(tot.labor)}</div>}</div>
            {rpt.labor.map((row,i)=><LaborCard key={row.id} row={row} onChange={r=>upd("labor",i,r)} onRemove={()=>del("labor",i)}/>)}
            <DashedAdd label="+ Add Worker" onClick={addLabor} color={T.orange}/>
          </div>
        )}
        {step===3&&(
          <div>
            <div style={{display:"flex",justifyContent:"space-between",marginBottom:12}}><div style={{fontSize:17,fontWeight:800}}>🚜 Equipment</div>{tot.equip>0&&<div style={{fontSize:16,fontWeight:800,color:T.green}}>${fmt(tot.equip)}</div>}</div>
            {rpt.equipment.map((row,i)=><EquipCard key={row.id} row={row} onChange={r=>upd("equipment",i,r)} onRemove={()=>del("equipment",i)}/>)}
            <DashedAdd label="+ Add Equipment" onClick={addEquip} color={T.yellow}/>
          </div>
        )}
        {step===4&&(
          <div>
            <div style={{fontSize:17,fontWeight:800,marginBottom:12}}>📦 Materials & Misc.</div>
            {rpt.materials.map((row,i)=><MatCard key={row.id} row={row} onChange={r=>upd("materials",i,r)} onRemove={()=>del("materials",i)}/>)}
            <DashedAdd label="+ Add Material / Item" onClick={addMat} color={T.blue}/>
          </div>
        )}
        {step===5&&(
          <div>
            <div style={{fontSize:17,fontWeight:800,marginBottom:12}}>✅ Review & Submit</div>
            <div style={{...cardS,marginBottom:12}}>
              <div style={{fontSize:11,color:T.orange,fontWeight:700,textTransform:"uppercase",letterSpacing:"1px",marginBottom:10}}>Summary</div>
              {[["Project",project.name],["Date",fmtDate(rpt.date)],["Report No.",rpt.report_no||"—"],["Workers",rpt.labor.length],["Equipment",rpt.equipment.length+" items"],["Materials",rpt.materials.length+" items"]].map(([l,v])=>(
                <div key={l} style={{display:"flex",justifyContent:"space-between",padding:"7px 0",borderBottom:`1px solid ${T.border}`}}>
                  <span style={{fontSize:13,color:T.sub}}>{l}</span><span style={{fontSize:13,fontWeight:600}}>{v}</span>
                </div>
              ))}
              {[["Labor",tot.labor,T.green],["Equipment",tot.equip,T.green],["Materials",tot.mats,T.green]].map(([l,v,c])=>(
                <div key={l} style={{display:"flex",justifyContent:"space-between",padding:"7px 0",borderBottom:`1px solid ${T.border}`}}>
                  <span style={{fontSize:13,color:T.sub}}>{l}</span><span style={{fontSize:13,fontWeight:700,color:c}}>${fmt(v)}</span>
                </div>
              ))}
              <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",paddingTop:14}}>
                <span style={{fontSize:16,fontWeight:800}}>Grand Total</span>
                <span style={{fontSize:26,fontWeight:900,color:T.orange,letterSpacing:"-1px"}}>${fmt(tot.grand)}</span>
              </div>
            </div>
          </div>
        )}
      </div>
      <div style={{position:"fixed",bottom:0,left:"50%",transform:"translateX(-50%)",width:"100%",maxWidth:480,background:T.bg+"EE",backdropFilter:"blur(12px)",borderTop:`1px solid ${T.border}`,padding:"12px 16px",display:"flex",gap:10}}>
        {step>1&&<button onClick={()=>{setStep(s=>s-1);scroll();}} style={{...ghostBtn,flex:1}}>← Back</button>}
        {step<5?<button onClick={()=>{setStep(s=>s+1);scroll();}} style={{...primBtn,flex:2,borderRadius:12}}>{step===4?"Review →":"Next →"}</button>
          :<button onClick={submit} style={{...primBtn,flex:2,borderRadius:12,opacity:saving?0.6:1}}>{saving?"Saving…":"💾 Save Report"}</button>}
      </div>
    </div>
  );
}

/* ═══════════════════════════════════════════════════════════════
   REPORT DETAIL  +  EXCEL EXPORT
═══════════════════════════════════════════════════════════════ */
function ReportDetail({report,project,user,onBack,onDelete,onApprove,onFlag}) {
  const [lb,setLb]=useState(null); const [flagNote,setFlagNote]=useState(""); const [flagging,setFlagging]=useState(false); const [deleting,setDeleting]=useState(false);
  const tot=reportTotals(report);
  const statusColor={submitted:T.yellow,approved:T.green,flagged:T.red}[report.status]||T.muted;
  function exportXLSX(){
    const wb=XLSX.utils.book_new();const rows=[];const blank=()=>Array(11).fill(null);
    const r1=blank();r1[1]="COLONIAL PIPELINE COMPANY";rows.push(r1);
    const r2=blank();r2[1]="DAILY REPORT-WORK PERFORMED BY CONTRACTOR";rows.push(r2);
    const r3=blank();r3[1]="LOCATION";r3[2]="LOCATION";r3[3]="AFE NO.";r3[4]="WORK ORDER\nNUMBER";r3[6]="REPORT DATE";r3[8]="REPORT\nNO.";rows.push(r3);
    const [yr,mo,dy]=(report.date||"").split("-");
    const r4=blank();r4[1]=project.location||"";r4[3]=project.afe||"";r4[4]=project.work_order||"";r4[6]=`${mo}/${dy}/${yr}`;r4[8]=report.report_no||"";rows.push(r4);
    const r5=blank();r5[1]="CONTRACTOR:";r5[2]="AIME";r5[3]="CONTRACTOR NO:";r5[8]="CONTRACTOR DATE:";r5[9]=report.date||"";rows.push(r5);
    const r6=blank();r6[1]="DESCRIPTION OF WORK DONE:";rows.push(r6);
    const r7=blank();r7[1]=report.description||"";rows.push(r7);
    const r8=blank();r8[1]="LABOR";rows.push(r8);
    const r9=blank();r9[1]="NAME";r9[3]="CLASSIFICATION";r9[5]="REG. HRS.";r9[6]="O.T. HRS.";r9[7]="TRAVEL HRS.";r9[8]="REGULAR RATE";r9[9]="AMOUNT";rows.push(r9);
    const labRows=[...(report.labor||[])];while(labRows.length<14)labRows.push(null);
    labRows.forEach(lr=>{const row=blank();if(lr){const pos=POSITIONS.find(p=>p.name===lr.classification);row[1]=lr.name||"";row[3]=lr.classification||"";if(pos&&!pos.flat){row[5]=parseFloat(lr.regHrs)||0;row[6]=parseFloat(lr.otHrs)||0;row[7]=parseFloat(lr.travelHrs)||0;}row[8]=pos?pos.rate:"";row[9]=laborAmt(lr);}else row[9]=0;rows.push(row);});
    const tlRow=blank();tlRow[8]="TOTAL LABOR";tlRow[9]=tot.labor;rows.push(tlRow);rows.push(blank());
    const ehRow=blank();ehRow[1]="EQUIPMENT";rows.push(ehRow);
    const ecRow=blank();ecRow[1]="DESCRIPTION";ecRow[6]="Quantity";ecRow[7]="Hours/Days";ecRow[8]="RATE";ecRow[9]="AMOUNT";rows.push(ecRow);
    const eqRows=[...(report.equipment||[])];while(eqRows.length<15)eqRows.push(null);
    eqRows.forEach(er=>{const row=blank();if(er){row[1]=er.description||"";row[6]=parseFloat(er.qty)||0;row[7]=parseFloat(er.usage)||0;row[8]=parseFloat(er.rate)||0;row[9]=equipAmt(er);}else row[9]=0;rows.push(row);});
    const teRow=blank();teRow[8]="TOTAL EQUIPMENT";teRow[9]=tot.equip;rows.push(teRow);rows.push(blank());
    const mhRow=blank();mhRow[2]="MATERIAL & MISCELLANEOUS";rows.push(mhRow);
    const mcRow=blank();mcRow[1]="QUANTITY";mcRow[2]="DESCRIPTION";mcRow[5]="AMOUNT";rows.push(mcRow);
    (report.materials||[]).forEach(m=>{const row=blank();row[1]=m.qty||"";row[2]=m.description||"";row[5]=parseFloat(m.amount)||0;rows.push(row);});
    const gtRow=blank();gtRow[8]="GRAND TOTAL";gtRow[9]=tot.grand;rows.push(gtRow);
    const sgRow=blank();sgRow[1]="VERIFIED AND ACCEPTED BY CO. REP";sgRow[3]="DATE";sgRow[4]="CERTIFIED AS CORRECT BY CONTRACTOR'S REP";rows.push(sgRow);
    const ws=XLSX.utils.aoa_to_sheet(rows);
    ws["!cols"]=[{wch:5.7},{wch:15},{wch:11.7},{wch:17.1},{wch:12.7},{wch:13.1},{wch:10},{wch:10},{wch:24.9},{wch:23.7},{wch:10.1}];
    const rng=XLSX.utils.decode_range(ws["!ref"]);
    for(let r=0;r<=rng.e.r;r++){const a=XLSX.utils.encode_cell({r,c:9});if(ws[a]&&typeof ws[a].v==="number")ws[a].z='"$"#,##0.00';const b=XLSX.utils.encode_cell({r,c:5});if(ws[b]&&typeof ws[b].v==="number")ws[b].z='"$"#,##0.00';}
    XLSX.utils.book_append_sheet(wb,ws,"Daily Report");
    XLSX.writeFile(wb,`AIME_${project.name.replace(/\s+/g,"_")}_${(report.date||"").replace(/-/g,"")}.xlsx`);
  }
  return (
    <div style={{background:T.bg,minHeight:"100vh",padding:16,fontFamily:"inherit"}}>
      <Lightbox src={lb} onClose={()=>setLb(null)}/>
      <button onClick={onBack} style={{...ghostBtn,marginBottom:14}}>← Reports</button>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:4}}>
        <div style={{fontSize:20,fontWeight:900,letterSpacing:"-0.5px"}}>{fmtDate(report.date)}</div>
        <span style={pill(statusColor)}>{(report.status||"submitted").toUpperCase()}</span>
      </div>
      {report.submitted_by&&<div style={{fontSize:12,color:T.muted,marginBottom:report.pm_notes?6:14}}>by {report.submitted_by}</div>}
      {report.pm_notes&&<div style={{...cardS,marginBottom:14,borderLeft:`3px solid ${T.red}`,background:T.redLow}}><div style={{fontSize:11,color:T.red,fontWeight:700,marginBottom:4}}>🚩 PM NOTE</div><div style={{fontSize:13,color:T.sub}}>{report.pm_notes}</div></div>}
      {report.description&&<div style={{...cardS,marginBottom:12,borderLeft:`3px solid ${T.blue}`}}><div style={{fontSize:11,color:T.muted,textTransform:"uppercase",letterSpacing:"1px",marginBottom:6}}>Work Done</div><div style={{fontSize:14,color:T.sub,lineHeight:1.6}}>{report.description}</div></div>}
      {(report.labor||[]).length>0&&<div style={{...cardS,marginBottom:12}}>
        <div style={{fontSize:12,color:T.orange,fontWeight:700,textTransform:"uppercase",letterSpacing:"1px",marginBottom:10}}>Labor · <span style={{color:T.green}}>${fmt(tot.labor)}</span></div>
        {report.labor.map((r,i)=>(
          <div key={i} style={{display:"flex",justifyContent:"space-between",padding:"8px 0",borderBottom:i<report.labor.length-1?`1px solid ${T.border}`:"none"}}>
            <div><div style={{fontSize:14,fontWeight:600}}>{r.name||"—"}</div><div style={{fontSize:11,color:T.muted}}>{r.classification} · {r.regHrs||0}reg {r.otHrs||0}OT {r.travelHrs||0}tr</div></div>
            <div style={{fontSize:14,fontWeight:800,color:T.green}}>${fmt(laborAmt(r))}</div>
          </div>
        ))}
      </div>}
      {(report.equipment||[]).length>0&&<div style={{...cardS,marginBottom:12}}>
        <div style={{fontSize:12,color:T.orange,fontWeight:700,textTransform:"uppercase",letterSpacing:"1px",marginBottom:10}}>Equipment · <span style={{color:T.green}}>${fmt(tot.equip)}</span></div>
        {report.equipment.map((r,i)=>(
          <div key={i} style={{display:"flex",justifyContent:"space-between",padding:"8px 0",borderBottom:i<report.equipment.length-1?`1px solid ${T.border}`:"none"}}>
            <div style={{flex:1,paddingRight:10}}><div style={{fontSize:13,fontWeight:600}}>{r.description}</div><div style={{fontSize:11,color:T.muted}}>Qty {r.qty} × {r.usage} {r.unit}</div></div>
            <div style={{fontSize:14,fontWeight:800,color:T.green}}>${fmt(equipAmt(r))}</div>
          </div>
        ))}
      </div>}
      {(report.materials||[]).length>0&&<div style={{...cardS,marginBottom:12}}>
        <div style={{fontSize:12,color:T.orange,fontWeight:700,textTransform:"uppercase",letterSpacing:"1px",marginBottom:10}}>Materials · <span style={{color:T.green}}>${fmt(tot.mats)}</span></div>
        {report.materials.map((r,i)=>(
          <div key={i} style={{padding:"8px 0",borderBottom:i<report.materials.length-1?`1px solid ${T.border}`:"none"}}>
            <div style={{display:"flex",justifyContent:"space-between",marginBottom:r.receipts?.length>0?8:0}}>
              <span style={{fontSize:13}}>{r.qty?`${r.qty}× `:""}{r.description}</span>
              <span style={{fontSize:13,fontWeight:700,color:T.green}}>${fmt(parseFloat(r.amount)||0)}</span>
            </div>
            {r.receipts?.length>0&&<div style={{display:"flex",gap:6,flexWrap:"wrap"}}>{r.receipts.map(rc=><img key={rc.id} src={rc.src} alt="" onClick={()=>setLb(rc.src)} style={{width:56,height:56,objectFit:"cover",borderRadius:8,cursor:"pointer"}}/>)}</div>}
          </div>
        ))}
      </div>}
      <div style={{...cardS,background:T.orangeLow,border:`1px solid ${T.orange}40`,marginBottom:12,display:"flex",justifyContent:"space-between",alignItems:"center"}}>
        <span style={{fontSize:15,fontWeight:800}}>Grand Total</span>
        <span style={{fontSize:26,fontWeight:900,color:T.orange,letterSpacing:"-1px"}}>${fmt(tot.grand)}</span>
      </div>
      {user?.role==="pm"&&report.status==="submitted"&&(
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:10}}>
          <button onClick={()=>onApprove&&onApprove(report.id)} style={{...primBtn,background:T.greenLow,color:T.green,border:`1px solid ${T.green}40`,borderRadius:12}}>✓ Approve</button>
          <button onClick={()=>setFlagging(!flagging)} style={{...primBtn,background:T.redLow,color:T.red,border:`1px solid ${T.red}40`,borderRadius:12}}>🚩 Flag</button>
        </div>
      )}
      {flagging&&<div style={{...cardS,marginBottom:10}}>
        <label style={lbl}>Flag Note for Crew</label>
        <textarea value={flagNote} onChange={e=>setFlagNote(e.target.value)} rows={3} placeholder="Explain what needs to be corrected…" style={{...inp,resize:"vertical",marginBottom:10}}/>
        <button onClick={()=>{onFlag&&onFlag(report.id,flagNote);setFlagging(false);}} style={{...primBtn,borderRadius:12}}>Send Flag</button>
      </div>}
      <button onClick={exportXLSX} style={{...primBtn,background:T.orangeLow,color:T.orange,border:`1px solid ${T.orange}40`,marginBottom:10,borderRadius:14}}>📥 Export to Excel (.xlsx)</button>
      <button onClick={async()=>{if(!window.confirm("Delete this report?"))return;setDeleting(true);await onDelete(report.id);setDeleting(false);}} style={dangerBtn}>{deleting?"Deleting…":"🗑 Delete Report"}</button>
    </div>
  );
}

/* ═══════════════════════════════════════════════════════════════
   TIME CARDS TAB
═══════════════════════════════════════════════════════════════ */
function TimeCardsTab({projectId,user,onErr}) {
  const [cards,setCards]=useState([]); const [loading,setLoading]=useState(true);
  const [showForm,setShowForm]=useState(false); const [saving,setSaving]=useState(false);
  const [f,setF]=useState({worker_name:user.name,date:today(),clock_in:"07:00",clock_out:"",notes:""});
  const weekStart=getWeekStart();

  async function load(){setLoading(true);try{setCards(await API.timeCards.forProject(projectId)||[]);}catch(e){onErr(e.message);}setLoading(false);}
  useEffect(()=>{load();},[projectId]);

  async function save(){
    if(!f.worker_name||!f.date)return;
    setSaving(true);
    const total_hours=calcHours(f.clock_in,f.clock_out);
    const ot_hours=Math.max(0,total_hours-8);
    try{await API.timeCards.create({...f,project_id:projectId,total_hours,ot_hours});await load();setShowForm(false);setF({worker_name:user.name,date:today(),clock_in:"07:00",clock_out:"",notes:""});}
    catch(e){onErr(e.message);}
    setSaving(false);
  }
  async function remove(id){try{await API.timeCards.remove(id);await load();}catch(e){onErr(e.message);}}

  // Weekly summary
  const weekCards=cards.filter(c=>c.date>=weekStart);
  const byWorker={};
  weekCards.forEach(c=>{
    if(!byWorker[c.worker_name])byWorker[c.worker_name]={name:c.worker_name,reg:0,ot:0,total:0};
    byWorker[c.worker_name].total+=(c.total_hours||0);
    byWorker[c.worker_name].ot+=(c.ot_hours||0);
    byWorker[c.worker_name].reg+=Math.min(c.total_hours||0,(c.total_hours||0)-(c.ot_hours||0));
  });
  const workerRows=Object.values(byWorker).sort((a,b)=>b.total-a.total);

  const todayCards=cards.filter(c=>c.date===today());
  const recentCards=cards.slice(0,40);

  return (
    <div>
      <button onClick={()=>setShowForm(!showForm)} style={{...primBtn,marginBottom:14,borderRadius:14}}>{showForm?"✕ Cancel":"⏱️ Log Time"}</button>
      {showForm&&(
        <div style={{...cardS,marginBottom:14,borderLeft:`3px solid ${T.green}`}}>
          <div style={{marginBottom:10}}><label style={lbl}>Worker</label>
            <select value={f.worker_name} onChange={e=>setF(x=>({...x,worker_name:e.target.value}))} style={inp}>
              {NAMES.map(n=><option key={n}>{n}</option>)}
            </select>
          </div>
          <div style={{marginBottom:10}}><label style={lbl}>Date</label><input type="date" value={f.date} onChange={e=>setF(x=>({...x,date:e.target.value}))} style={inp}/></div>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:10}}>
            <div><label style={lbl}>Clock In</label><input type="time" value={f.clock_in} onChange={e=>setF(x=>({...x,clock_in:e.target.value}))} style={inp}/></div>
            <div><label style={lbl}>Clock Out</label><input type="time" value={f.clock_out} onChange={e=>setF(x=>({...x,clock_out:e.target.value}))} style={inp}/></div>
          </div>
          {f.clock_in&&f.clock_out&&(()=>{const h=calcHours(f.clock_in,f.clock_out);const ot=Math.max(0,h-8);return h>0&&(
            <div style={{background:T.greenLow,borderRadius:10,padding:"10px 12px",marginBottom:10,display:"flex",gap:16}}>
              <div><div style={{fontSize:18,fontWeight:900,color:T.green}}>{h.toFixed(2)}h</div><div style={{fontSize:10,color:T.muted,textTransform:"uppercase"}}>Total</div></div>
              <div><div style={{fontSize:18,fontWeight:900,color:ot>0?T.yellow:T.muted}}>{Math.min(h,8).toFixed(2)}h</div><div style={{fontSize:10,color:T.muted,textTransform:"uppercase"}}>Regular</div></div>
              {ot>0&&<div><div style={{fontSize:18,fontWeight:900,color:T.yellow}}>{ot.toFixed(2)}h</div><div style={{fontSize:10,color:T.muted,textTransform:"uppercase"}}>Overtime</div></div>}
            </div>
          );})()}
          <div style={{marginBottom:10}}><label style={lbl}>Notes</label><input type="text" placeholder="Optional notes…" value={f.notes} onChange={e=>setF(x=>({...x,notes:e.target.value}))} style={inp}/></div>
          <button onClick={save} style={{...primBtn,background:T.green,color:"#09090B",borderRadius:12}}>{saving?"Saving…":"Save Time Card"}</button>
        </div>
      )}

      {loading&&<Spinner/>}
      {!loading&&(
        <>
          {/* Weekly Summary */}
          {workerRows.length>0&&(
            <div style={{...cardS,marginBottom:16}}>
              <div style={{fontSize:12,fontWeight:700,color:T.green,textTransform:"uppercase",letterSpacing:"1px",marginBottom:10}}>⏱️ This Week's Hours</div>
              {workerRows.map(w=>(
                <div key={w.name} style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"8px 0",borderBottom:`1px solid ${T.border}`}}>
                  <span style={{fontSize:14,fontWeight:600}}>{w.name}</span>
                  <div style={{display:"flex",gap:12,alignItems:"center"}}>
                    <span style={{fontSize:12,color:T.muted}}>{w.reg.toFixed(1)}reg</span>
                    {w.ot>0&&<span style={{fontSize:12,color:T.yellow}}>{w.ot.toFixed(1)}OT</span>}
                    <span style={{fontSize:15,fontWeight:800,color:T.green}}>{w.total.toFixed(1)}h</span>
                  </div>
                </div>
              ))}
            </div>
          )}

          {/* Today */}
          {todayCards.length>0&&(
            <div style={{fontSize:12,fontWeight:700,color:T.muted,textTransform:"uppercase",letterSpacing:"1px",marginBottom:10}}>Today On Site</div>
          )}
          {todayCards.map(c=><TimeCardRow key={c.id} c={c} onRemove={()=>remove(c.id)}/>)}

          {/* Recent */}
          {recentCards.filter(c=>c.date!==today()).length>0&&(
            <div style={{fontSize:12,fontWeight:700,color:T.muted,textTransform:"uppercase",letterSpacing:"1px",margin:"14px 0 10px"}}>Recent</div>
          )}
          {recentCards.filter(c=>c.date!==today()).map(c=><TimeCardRow key={c.id} c={c} onRemove={()=>remove(c.id)}/>)}
          {cards.length===0&&!showForm&&<div style={{textAlign:"center",padding:"40px 0",color:T.muted}}><div style={{fontSize:32,marginBottom:8}}>⏱️</div><div>No time cards yet. Tap Log Time to add the first entry.</div></div>}
        </>
      )}
    </div>
  );
}
function TimeCardRow({c,onRemove}) {
  const ot=c.ot_hours||0;
  return (
    <div style={{...cardS,marginBottom:8,display:"flex",justifyContent:"space-between",alignItems:"center"}}>
      <div>
        <div style={{fontSize:14,fontWeight:700}}>{c.worker_name}</div>
        <div style={{fontSize:11,color:T.muted,marginTop:3}}>
          {fmtShort(c.date)}{c.clock_in?" · "+c.clock_in:""}{c.clock_out?" → "+c.clock_out:""}
        </div>
        {c.notes&&<div style={{fontSize:11,color:T.sub,marginTop:2}}>{c.notes}</div>}
      </div>
      <div style={{display:"flex",alignItems:"center",gap:10}}>
        <div style={{textAlign:"right"}}>
          <div style={{fontSize:16,fontWeight:800,color:T.green}}>{(c.total_hours||0).toFixed(1)}h</div>
          {ot>0&&<div style={{fontSize:10,color:T.yellow}}>{ot.toFixed(1)} OT</div>}
        </div>
        <button onClick={onRemove} style={{background:"none",border:"none",color:T.muted,cursor:"pointer",fontSize:16,padding:0}}>🗑</button>
      </div>
    </div>
  );
}

/* ═══════════════════════════════════════════════════════════════
   CREW TAB  (who's on site + equipment on site)
═══════════════════════════════════════════════════════════════ */
function CrewTab({projectId,user,onErr}) {
  const [equip,setEquip]=useState([]); const [loading,setLoading]=useState(true);
  const [showForm,setShowForm]=useState(false); const [saving,setSaving]=useState(false);
  const [f,setF]=useState({equipment_name:"",quantity:1,operator_name:"",hours_used:"",notes:"",date:today()});

  async function load(){setLoading(true);try{setEquip(await API.equipment.forProject(projectId)||[]);}catch(e){onErr(e.message);}setLoading(false);}
  useEffect(()=>{load();},[projectId]);
  async function save(){
    if(!f.equipment_name)return;setSaving(true);
    try{await API.equipment.create({...f,project_id:projectId});await load();setShowForm(false);setF({equipment_name:"",quantity:1,operator_name:"",hours_used:"",notes:"",date:today()});}
    catch(e){onErr(e.message);}setSaving(false);
  }
  async function remove(id){try{await API.equipment.remove(id);await load();}catch(e){onErr(e.message);}}

  const todayEquip=equip.filter(e=>e.date===today());
  const prevEquip=equip.filter(e=>e.date!==today()).slice(0,20);

  return (
    <div>
      <button onClick={()=>setShowForm(!showForm)} style={{...primBtn,marginBottom:14,borderRadius:14}}>{showForm?"✕ Cancel":"🚜 Log Equipment On Site"}</button>
      {showForm&&(
        <div style={{...cardS,marginBottom:14,borderLeft:`3px solid ${T.yellow}`}}>
          <div style={{marginBottom:10}}><label style={lbl}>Equipment Name</label>
            <select value={f.equipment_name} onChange={e=>setF(x=>({...x,equipment_name:e.target.value}))} style={inp}>
              <option value="">— Select or type below —</option>
              {EQUIP_LIST.filter(e=>!e.section).map(e=><option key={e.name} value={e.name}>{e.name}</option>)}
            </select>
          </div>
          {!EQUIP_LIST.find(e=>!e.section&&e.name===f.equipment_name)&&(
            <div style={{marginBottom:10}}><label style={lbl}>Or Enter Custom Equipment</label><input type="text" placeholder="Equipment name…" value={f.equipment_name} onChange={e=>setF(x=>({...x,equipment_name:e.target.value}))} style={inp}/></div>
          )}
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:10}}>
            <div><label style={lbl}>Quantity</label><input type="number" min="1" value={f.quantity} onChange={e=>setF(x=>({...x,quantity:e.target.value}))} style={inp}/></div>
            <div><label style={lbl}>Hours Used</label><input type="number" min="0" step="0.5" placeholder="0" value={f.hours_used} onChange={e=>setF(x=>({...x,hours_used:e.target.value}))} style={inp}/></div>
          </div>
          <div style={{marginBottom:10}}><label style={lbl}>Operator</label>
            <select value={f.operator_name} onChange={e=>setF(x=>({...x,operator_name:e.target.value}))} style={inp}>
              <option value="">— Operator (optional) —</option>
              {NAMES.map(n=><option key={n}>{n}</option>)}
            </select>
          </div>
          <div style={{marginBottom:10}}><label style={lbl}>Date</label><input type="date" value={f.date} onChange={e=>setF(x=>({...x,date:e.target.value}))} style={inp}/></div>
          <div style={{marginBottom:10}}><label style={lbl}>Notes</label><input type="text" placeholder="Condition, issues, etc." value={f.notes} onChange={e=>setF(x=>({...x,notes:e.target.value}))} style={inp}/></div>
          <button onClick={save} style={{...primBtn,background:T.yellow,color:"#09090B",borderRadius:12}}>{saving?"Saving…":"Save Entry"}</button>
        </div>
      )}
      {loading&&<Spinner/>}
      {!loading&&(
        <>
          {todayEquip.length>0&&<div style={{fontSize:12,fontWeight:700,color:T.muted,textTransform:"uppercase",letterSpacing:"1px",marginBottom:10}}>On Site Today</div>}
          {todayEquip.map(e=><EquipSiteRow key={e.id} e={e} onRemove={()=>remove(e.id)}/>)}
          {prevEquip.length>0&&<div style={{fontSize:12,fontWeight:700,color:T.muted,textTransform:"uppercase",letterSpacing:"1px",margin:"16px 0 10px"}}>Previous</div>}
          {prevEquip.map(e=><EquipSiteRow key={e.id} e={e} onRemove={()=>remove(e.id)}/>)}
          {equip.length===0&&!showForm&&<div style={{textAlign:"center",padding:"40px 0",color:T.muted}}><div style={{fontSize:32,marginBottom:8}}>🚜</div><div>No equipment logged. Tap above to log equipment on site.</div></div>}
        </>
      )}
    </div>
  );
}
function EquipSiteRow({e,onRemove}) {
  return (
    <div style={{...cardS,marginBottom:8,borderLeft:`3px solid ${T.yellow}`,display:"flex",justifyContent:"space-between",alignItems:"center"}}>
      <div>
        <div style={{fontSize:14,fontWeight:700}}>{e.equipment_name}</div>
        <div style={{fontSize:11,color:T.muted,marginTop:3}}>
          {fmtShort(e.date)} · Qty {e.quantity||1}{e.operator_name?" · "+e.operator_name:""}{e.hours_used?" · "+e.hours_used+"h":""}
        </div>
        {e.notes&&<div style={{fontSize:11,color:T.sub,marginTop:2}}>{e.notes}</div>}
      </div>
      <button onClick={onRemove} style={{background:"none",border:"none",color:T.muted,cursor:"pointer",fontSize:16,padding:0,marginLeft:12}}>🗑</button>
    </div>
  );
}

/* ═══════════════════════════════════════════════════════════════
   SUBCONTRACTORS TAB
═══════════════════════════════════════════════════════════════ */
function SubsTab({projectId,user,onErr}) {
  const [subs,setSubs]=useState([]); const [loading,setLoading]=useState(true);
  const [showForm,setShowForm]=useState(false); const [saving,setSaving]=useState(false);
  const [f,setF]=useState({date:today(),company_name:"",trade:"",contact_name:"",contact_phone:"",workers_count:1,hours_worked:"",work_description:""});
  const set=(k,v)=>setF(x=>({...x,[k]:v}));

  async function load(){setLoading(true);try{setSubs(await API.subs.forProject(projectId)||[]);}catch(e){onErr(e.message);}setLoading(false);}
  useEffect(()=>{load();},[projectId]);
  async function save(){
    if(!f.company_name)return;setSaving(true);
    try{await API.subs.create({...f,project_id:projectId,created_by:user.name});await load();setShowForm(false);setF({date:today(),company_name:"",trade:"",contact_name:"",contact_phone:"",workers_count:1,hours_worked:"",work_description:""});}
    catch(e){onErr(e.message);}setSaving(false);
  }
  async function remove(id){try{await API.subs.remove(id);await load();}catch(e){onErr(e.message);}}

  const trades=["Electrical","Mechanical","Civil","Welding","Coating","Survey","Inspection","HDD","Boring","Concrete","Other"];

  return (
    <div>
      <button onClick={()=>setShowForm(!showForm)} style={{...primBtn,marginBottom:14,borderRadius:14}}>{showForm?"✕ Cancel":"🏢 Log Subcontractor"}</button>
      {showForm&&(
        <div style={{...cardS,marginBottom:14,borderLeft:`3px solid ${T.purple}`}}>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:10}}>
            <div style={{gridColumn:"1/-1"}}><label style={lbl}>Company Name *</label><input type="text" placeholder="Sub company name" value={f.company_name} onChange={e=>set("company_name",e.target.value)} style={inp}/></div>
            <div><label style={lbl}>Trade</label>
              <select value={f.trade} onChange={e=>set("trade",e.target.value)} style={inp}>
                <option value="">— Select —</option>
                {trades.map(t=><option key={t}>{t}</option>)}
              </select>
            </div>
            <div><label style={lbl}>Date</label><input type="date" value={f.date} onChange={e=>set("date",e.target.value)} style={inp}/></div>
            <div><label style={lbl}>Contact Name</label><input type="text" placeholder="Foreman / Contact" value={f.contact_name} onChange={e=>set("contact_name",e.target.value)} style={inp}/></div>
            <div><label style={lbl}>Contact Phone</label><input type="tel" placeholder="555-555-5555" value={f.contact_phone} onChange={e=>set("contact_phone",e.target.value)} style={inp}/></div>
            <div><label style={lbl}>Workers on Site</label><input type="number" min="0" value={f.workers_count} onChange={e=>set("workers_count",e.target.value)} style={inp}/></div>
            <div><label style={lbl}>Hours Worked</label><input type="number" min="0" step="0.5" placeholder="0" value={f.hours_worked} onChange={e=>set("hours_worked",e.target.value)} style={inp}/></div>
          </div>
          <div style={{marginBottom:10}}><label style={lbl}>Work Description</label><textarea placeholder="What work was performed?" value={f.work_description} onChange={e=>set("work_description",e.target.value)} rows={3} style={{...inp,resize:"vertical"}}/></div>
          <button onClick={save} style={{...primBtn,background:T.purple,color:"#fff",borderRadius:12}}>{saving?"Saving…":"Save Sub Entry"}</button>
        </div>
      )}
      {loading&&<Spinner/>}
      {!loading&&subs.map(s=>(
        <div key={s.id} style={{...cardS,marginBottom:10,borderLeft:`3px solid ${T.purple}`}}>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start"}}>
            <div style={{flex:1}}>
              <div style={{fontSize:15,fontWeight:800}}>{s.company_name}</div>
              <div style={{display:"flex",gap:8,marginTop:5,flexWrap:"wrap"}}>
                {s.trade&&<span style={pill(T.purple)}>{s.trade}</span>}
                <span style={pill(T.muted)}>{fmtShort(s.date)}</span>
                {s.workers_count>0&&<span style={pill(T.blue)}>👷 {s.workers_count} workers</span>}
                {s.hours_worked>0&&<span style={pill(T.green)}>{s.hours_worked}h</span>}
              </div>
              {s.contact_name&&<div style={{fontSize:12,color:T.sub,marginTop:6}}>📞 {s.contact_name}{s.contact_phone?" · "+s.contact_phone:""}</div>}
              {s.work_description&&<div style={{fontSize:12,color:T.sub,marginTop:4,lineHeight:1.5}}>{s.work_description}</div>}
            </div>
            <button onClick={()=>remove(s.id)} style={{background:"none",border:"none",color:T.muted,cursor:"pointer",fontSize:16,padding:"0 0 0 10px"}}>🗑</button>
          </div>
        </div>
      ))}
      {!loading&&subs.length===0&&!showForm&&<div style={{textAlign:"center",padding:"40px 0",color:T.muted}}><div style={{fontSize:32,marginBottom:8}}>🏢</div><div>No subcontractors logged. Tap above to add a sub entry.</div></div>}
    </div>
  );
}

/* ═══════════════════════════════════════════════════════════════
   SAFETY TAB
═══════════════════════════════════════════════════════════════ */
function SafetyTab({projectId,safety,user,onRefresh,onErr}) {
  const [showForm,setShowForm]=useState(false); const [saving,setSaving]=useState(false);
  const [type,setType]=useState("toolbox");
  const [f,setF]=useState({date:today(),topic:"",notes:"",severity:"low"});
  const TC={toolbox:T.blue,observation:T.yellow,incident:T.red,nearmiss:T.orange,jsa:T.purple};
  const TL={toolbox:"🛠 Toolbox Talk",observation:"👁 Observation",incident:"🚨 Incident",nearmiss:"⚠️ Near Miss",jsa:"📋 JSA"};
  async function save(){
    if(!f.topic.trim())return;setSaving(true);
    try{await API.safety.create({...f,type,project_id:projectId,created_by:user.name});await onRefresh();setShowForm(false);setF({date:today(),topic:"",notes:"",severity:"low"});}
    catch(e){onErr(e.message);}setSaving(false);
  }
  async function del(id){try{await API.safety.remove(id);await onRefresh();}catch(e){onErr(e.message);}}
  return (
    <div>
      <button onClick={()=>setShowForm(!showForm)} style={{...primBtn,marginBottom:14,borderRadius:14}}>{showForm?"✕ Cancel":"⛑️ Log Safety Entry"}</button>
      {showForm&&(
        <div style={{...cardS,marginBottom:14,borderLeft:`3px solid ${T.yellow}`}}>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8,marginBottom:12}}>
            {Object.entries(TL).map(([k,v])=>(
              <button key={k} onClick={()=>setType(k)} style={{padding:"10px",borderRadius:10,border:`2px solid ${type===k?TC[k]:T.border}`,background:type===k?TC[k]+"20":T.surface,color:type===k?TC[k]:T.sub,fontSize:12,fontWeight:700,cursor:"pointer",fontFamily:"inherit"}}>{v}</button>
            ))}
          </div>
          <div style={{marginBottom:10}}><label style={lbl}>Date</label><input type="date" value={f.date} onChange={e=>setF(x=>({...x,date:e.target.value}))} style={inp}/></div>
          <div style={{marginBottom:10}}><label style={lbl}>{type==="toolbox"?"Topic / Title":type==="jsa"?"Job / Task Name":"Description"}</label><input type="text" placeholder="Describe…" value={f.topic} onChange={e=>setF(x=>({...x,topic:e.target.value}))} style={inp}/></div>
          <div style={{marginBottom:10}}><label style={lbl}>Notes / Corrective Action</label><textarea rows={3} placeholder="Additional details…" value={f.notes} onChange={e=>setF(x=>({...x,notes:e.target.value}))} style={{...inp,resize:"vertical"}}/></div>
          {(type==="incident"||type==="nearmiss")&&<div style={{marginBottom:10}}><label style={lbl}>Severity</label>
            <select value={f.severity} onChange={e=>setF(x=>({...x,severity:e.target.value}))} style={inp}>
              <option value="low">Low – First Aid Only</option><option value="medium">Medium – Recordable</option><option value="high">High – Lost Time Injury</option>
            </select>
          </div>}
          <button onClick={save} style={{...primBtn,background:T.yellow,color:"#09090B",borderRadius:12}}>{saving?"Saving…":"Save Entry"}</button>
        </div>
      )}
      {safety.length===0&&!showForm&&<div style={{textAlign:"center",padding:"40px 0",color:T.muted}}><div style={{fontSize:32,marginBottom:8}}>⛑️</div><div>No safety entries. Log toolbox talks, JSAs, observations and incidents here.</div></div>}
      {[...safety].map(s=>(
        <div key={s.id} style={{...cardS,marginBottom:9,borderLeft:`3px solid ${TC[s.type]||T.border}`}}>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start"}}>
            <div style={{flex:1}}>
              <span style={{...pill(TC[s.type]||T.muted),marginBottom:6,display:"inline-flex"}}>{TL[s.type]||s.type}</span>
              <div style={{fontSize:14,fontWeight:700,marginTop:4}}>{s.topic}</div>
              {s.notes&&<div style={{fontSize:12,color:T.sub,marginTop:4,lineHeight:1.5}}>{s.notes}</div>}
              <div style={{fontSize:11,color:T.muted,marginTop:6}}>{fmtDate(s.date)} · {s.created_by}</div>
            </div>
            <div style={{display:"flex",flexDirection:"column",alignItems:"flex-end",gap:6}}>
              {(s.type==="incident"||s.type==="nearmiss")&&<span style={pill(s.severity==="high"?T.red:s.severity==="medium"?T.yellow:T.green)}>{(s.severity||"low").toUpperCase()}</span>}
              <button onClick={()=>del(s.id)} style={{background:"none",border:"none",color:T.muted,cursor:"pointer",fontSize:16,padding:0}}>🗑</button>
            </div>
          </div>
        </div>
      ))}
    </div>
  );
}

/* ═══════════════════════════════════════════════════════════════
   PHOTOS TAB
═══════════════════════════════════════════════════════════════ */
function PhotosTab({projectId,photos,onRefresh,onErr}) {
  const [caption,setCaption]=useState(""); const [saving,setSaving]=useState(false); const [lb,setLb]=useState(null);
  const fileRef=useRef(null);
  async function handleFiles(files){
    setSaving(true);
    try{for(const f of files){if(!f.type.startsWith("image/"))continue;const src=await compressImg(f,1100,0.72);await API.photos.create({project_id:projectId,src,caption,date:today()});}await onRefresh();setCaption("");}
    catch(e){onErr(e.message);}setSaving(false);
  }
  async function del(id){try{await API.photos.remove(id);await onRefresh();}catch(e){onErr(e.message);}}
  return (
    <div>
      <Lightbox src={lb} onClose={()=>setLb(null)}/>
      <div style={{...cardS,marginBottom:14,borderStyle:"dashed",borderColor:T.orange+"44"}}>
        <div style={{marginBottom:10}}><label style={lbl}>Caption (optional)</label><input type="text" placeholder="What's this photo of?" value={caption} onChange={e=>setCaption(e.target.value)} style={inp}/></div>
        <button onClick={()=>fileRef.current?.click()} style={{...primBtn}}>{saving?"Uploading…":"📷 Add Site Photos"}</button>
        <input ref={fileRef} type="file" accept="image/*" capture="environment" multiple style={{display:"none"}} onChange={e=>{handleFiles(Array.from(e.target.files));e.target.value="";}}/>
      </div>
      {photos.length===0&&<div style={{textAlign:"center",padding:"40px 0",color:T.muted}}><div style={{fontSize:32,marginBottom:8}}>📷</div><div>No photos yet.</div></div>}
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}>
        {photos.map(p=>(
          <div key={p.id} style={{position:"relative",borderRadius:12,overflow:"hidden",aspectRatio:"4/3",background:T.card}}>
            <img src={p.src} alt={p.caption} onClick={()=>setLb(p.src)} style={{width:"100%",height:"100%",objectFit:"cover",cursor:"pointer",display:"block"}}/>
            <button onClick={()=>del(p.id)} style={{position:"absolute",top:6,right:6,background:"rgba(0,0,0,0.75)",border:"none",color:"#fff",borderRadius:"50%",width:24,height:24,fontSize:12,cursor:"pointer"}}>×</button>
            {p.caption&&<div style={{position:"absolute",bottom:0,left:0,right:0,background:"rgba(0,0,0,0.65)",color:"#fff",fontSize:11,padding:"4px 8px"}}>{p.caption}</div>}
            <div style={{position:"absolute",top:6,left:6,background:"rgba(0,0,0,0.6)",borderRadius:6,padding:"2px 6px",fontSize:9,color:"#fff"}}>{fmtShort(p.date)}</div>
          </div>
        ))}
      </div>
    </div>
  );
}

/* ═══════════════════════════════════════════════════════════════
   WEATHER TAB
═══════════════════════════════════════════════════════════════ */
function WeatherTab({projectId,project,weather,onRefresh,onErr}) {
  const [fetching,setFetching]=useState(false); const [liveWeather,setLiveWeather]=useState(null);
  const [manualNote,setManualNote]=useState(""); const [saving,setSaving]=useState(false);

  async function autoFetch(){
    if(!project.location){onErr("Add a location to the project first.");return;}
    setFetching(true);setLiveWeather(null);
    try{
      const w=await fetchWeather(project.location);
      setLiveWeather(w);
    }catch(e){onErr(e.message);}
    setFetching(false);
  }

  async function logWeather(){
    if(!liveWeather)return;setSaving(true);
    const c=liveWeather.current;
    const [desc]=WMO[c.weathercode]||["Unknown","🌡️"];
    try{
      await API.weather.upsert({
        project_id:projectId, date:today(),
        temp_high:liveWeather.daily?.temperature_2m_max?.[0]||c.temperature_2m,
        temp_low:liveWeather.daily?.temperature_2m_min?.[0]||c.temperature_2m,
        conditions:desc, wind_speed:c.windspeed_10m,
        precipitation:liveWeather.daily?.precipitation_sum?.[0]||0,
        notes:manualNote,
      });
      await onRefresh();setLiveWeather(null);setManualNote("");
    }catch(e){onErr(e.message);}setSaving(false);
  }

  async function del(id){try{await API.weather.remove(id);await onRefresh();}catch(e){onErr(e.message);}}

  return (
    <div>
      <button onClick={autoFetch} style={{...primBtn,marginBottom:14,borderRadius:14,opacity:fetching?0.6:1}}>{fetching?"🌐 Fetching weather…":"🌤️ Auto-Fetch Today's Weather"}</button>
      {!project.location&&<div style={{...cardS,marginBottom:14,background:T.yellowLow,border:`1px solid ${T.yellow}40`}}><div style={{fontSize:13,color:T.yellow}}>⚠️ Set a location on this project (Info tab) to auto-fetch weather.</div></div>}

      {liveWeather&&(()=>{
        const c=liveWeather.current;
        const [desc,icon]=WMO[c.weathercode]||["Unknown","🌡️"];
        return (
          <div style={{...cardS,marginBottom:14,borderLeft:`3px solid ${T.blue}`}}>
            <div style={{fontSize:11,color:T.blue,fontWeight:700,textTransform:"uppercase",letterSpacing:"1px",marginBottom:10}}>Live Weather · {liveWeather.locationName}</div>
            <div style={{display:"flex",alignItems:"center",gap:12,marginBottom:12}}>
              <span style={{fontSize:48}}>{icon}</span>
              <div>
                <div style={{fontSize:28,fontWeight:900,color:T.text,letterSpacing:"-1px"}}>{Math.round(c.temperature_2m)}°F</div>
                <div style={{fontSize:14,color:T.sub}}>{desc}</div>
                <div style={{fontSize:12,color:T.muted}}>Feels like {Math.round(c.apparent_temperature)}°F</div>
              </div>
            </div>
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:8,marginBottom:12}}>
              {[["💨 Wind",Math.round(c.windspeed_10m)+" mph"],["🌡️ High",Math.round(liveWeather.daily?.temperature_2m_max?.[0]||c.temperature_2m)+"°F"],["🌧️ Precip",(liveWeather.daily?.precipitation_sum?.[0]||0).toFixed(2)+'"']].map(([l,v])=>(
                <div key={l} style={{background:T.surface,borderRadius:10,padding:"8px",textAlign:"center"}}>
                  <div style={{fontSize:13,fontWeight:700,color:T.text}}>{v}</div>
                  <div style={{fontSize:10,color:T.muted}}>{l}</div>
                </div>
              ))}
            </div>
            <div style={{marginBottom:10}}><label style={lbl}>Field Notes (optional)</label><input type="text" placeholder="Work impacted by weather?" value={manualNote} onChange={e=>setManualNote(e.target.value)} style={inp}/></div>
            <button onClick={logWeather} style={{...primBtn,background:T.blue,borderRadius:12}}>{saving?"Saving…":"💾 Log This Weather"}</button>
          </div>
        );
      })()}

      {/* History */}
      {weather.length>0&&<div style={{fontSize:12,fontWeight:700,color:T.muted,textTransform:"uppercase",letterSpacing:"1px",marginBottom:10}}>Weather History</div>}
      {weather.map(w=>{
        const [,icon]=WMO[Object.keys(WMO).find(k=>WMO[k][0]===w.conditions)]||["","🌡️"];
        return (
          <div key={w.id} style={{...cardS,marginBottom:8,display:"flex",justifyContent:"space-between",alignItems:"center"}}>
            <div>
              <div style={{display:"flex",alignItems:"center",gap:8}}>
                <span style={{fontSize:20}}>{icon}</span>
                <div>
                  <div style={{fontSize:14,fontWeight:700}}>{w.conditions||"Logged"}</div>
                  <div style={{fontSize:11,color:T.muted}}>{fmtShort(w.date)}{w.wind_speed?" · "+Math.round(w.wind_speed)+"mph wind":""}{w.precipitation>0?" · "+w.precipitation+'" precip':""}</div>
                </div>
              </div>
              {w.notes&&<div style={{fontSize:12,color:T.sub,marginTop:4}}>{w.notes}</div>}
            </div>
            <div style={{display:"flex",alignItems:"center",gap:8}}>
              {w.temp_high&&<div style={{textAlign:"right"}}><div style={{fontSize:15,fontWeight:800,color:T.orange}}>{Math.round(w.temp_high)}°</div><div style={{fontSize:10,color:T.muted}}>{w.temp_low?Math.round(w.temp_low)+"° lo":""}</div></div>}
              <button onClick={()=>del(w.id)} style={{background:"none",border:"none",color:T.muted,cursor:"pointer",fontSize:16,padding:0}}>🗑</button>
            </div>
          </div>
        );
      })}
      {weather.length===0&&!liveWeather&&<div style={{textAlign:"center",padding:"40px 0",color:T.muted}}><div style={{fontSize:32,marginBottom:8}}>🌤️</div><div>No weather logs yet. Tap Auto-Fetch to pull today's weather.</div></div>}
    </div>
  );
}

/* ═══════════════════════════════════════════════════════════════
   INFO TAB
═══════════════════════════════════════════════════════════════ */
function InfoTab({project,user,onEdit,onArchive}) {
  return (
    <div>
      <div style={cardS}>
        {[["Client",project.client],["Location",project.location],["AFE No.",project.afe],["Work Order",project.work_order],["Start Date",fmtDate(project.start_date)],["Status",project.status],["Created By",project.created_by]].map(([l,v])=>v?(
          <div key={l} style={{display:"flex",justifyContent:"space-between",padding:"11px 0",borderBottom:`1px solid ${T.border}`}}>
            <span style={{fontSize:13,color:T.muted}}>{l}</span>
            <span style={{fontSize:13,fontWeight:600}}>{v}</span>
          </div>
        ):null)}
      </div>
      {project.notes&&<div style={{...cardS,marginTop:12}}><div style={{fontSize:11,color:T.muted,textTransform:"uppercase",letterSpacing:"1px",marginBottom:8}}>Notes</div><div style={{fontSize:14,color:T.sub,lineHeight:1.6}}>{project.notes}</div></div>}
      {user.role==="pm"&&<div style={{marginTop:16,display:"flex",flexDirection:"column",gap:10}}>
        <button onClick={onEdit} style={{...ghostBtn,width:"100%",textAlign:"center"}}>✏️ Edit Project</button>
        <button onClick={onArchive} style={{...ghostBtn,width:"100%",textAlign:"center",color:T.muted}}>{project.status==="active"?"📦 Archive Project":"♻️ Restore Project"}</button>
      </div>}
    </div>
  );
}

/* ═══════════════════════════════════════════════════════════════
   PROJECT DETAIL  (orchestrator)
═══════════════════════════════════════════════════════════════ */
const PTABS=[
  {id:"reports",icon:"📋",label:"Reports"},
  {id:"time",icon:"⏱️",label:"Time"},
  {id:"crew",icon:"🚜",label:"Crew"},
  {id:"subs",icon:"🏢",label:"Subs"},
  {id:"safety",icon:"⛑️",label:"Safety"},
  {id:"photos",icon:"📷",label:"Photos"},
  {id:"weather",icon:"🌤️",label:"Weather"},
  {id:"info",icon:"ℹ️",label:"Info"},
];

function ProjectDetail({project:initP,user,onBack,onProjectUpdated}) {
  const [project,setProject]=useState(initP);
  const [reports,setReports]=useState([]); const [safety,setSafety]=useState([]);
  const [photos,setPhotos]=useState([]); const [weather,setWeather]=useState([]);
  const [tab,setTab]=useState("reports"); const [loading,setLoading]=useState(true);
  const [err,setErr]=useState(""); const [screen,setScreen]=useState("detail");
  const [activeReport,setActiveReport]=useState(null); const [editProject,setEditProject]=useState(false);

  async function load(silent=false){
    if(!silent)setLoading(true);
    try{
      const [reps,saf,phs,wx]=await Promise.all([
        API.reports.forProject(project.id),
        API.safety.forProject(project.id),
        API.photos.forProject(project.id),
        API.weather.forProject(project.id),
      ]);
      setReports(reps||[]);setSafety(saf||[]);setPhotos(phs||[]);setWeather(wx||[]);
    }catch(e){setErr(e.message);}
    if(!silent)setLoading(false);
  }
  useEffect(()=>{load();},[project.id]);

  async function saveReport(d){try{await API.reports.create({...d,project_id:project.id});await load(true);setScreen("detail");}catch(e){setErr(e.message);}}
  async function deleteReport(id){try{await API.reports.remove(id);setActiveReport(null);await load(true);setScreen("detail");}catch(e){setErr(e.message);}}
  async function approveReport(id){try{await API.reports.update(id,{status:"approved",approved_by:user.name,approved_at:new Date().toISOString()});setActiveReport(r=>({...r,status:"approved",approved_by:user.name}));await load(true);}catch(e){setErr(e.message);}}
  async function flagReport(id,pm_notes){try{await API.reports.update(id,{status:"flagged",pm_notes});setActiveReport(r=>({...r,status:"flagged",pm_notes}));await load(true);}catch(e){setErr(e.message);}}
  async function updateProject(data){try{const [u]=await API.projects.update(project.id,data);setProject(u);onProjectUpdated(u);setEditProject(false);}catch(e){setErr(e.message);}}
  async function archiveProject(){
    if(!window.confirm(project.status==="active"?"Archive this project?":"Restore this project?"))return;
    await updateProject({status:project.status==="active"?"archived":"active"});onBack();
  }

  const tot=reports.reduce((s,r)=>{const t=reportTotals(r);return{l:s.l+t.labor,e:s.e+t.equip,m:s.m+t.mats,g:s.g+t.grand};},{l:0,e:0,m:0,g:0});

  if(screen==="newReport") return <DailyReportForm user={user} project={project} onSave={saveReport} onCancel={()=>setScreen("detail")}/>;
  if(screen==="reportDetail"&&activeReport) return <ReportDetail report={activeReport} project={project} user={user} onBack={()=>setScreen("detail")} onDelete={deleteReport} onApprove={approveReport} onFlag={flagReport}/>;
  if(editProject) return <ProjectForm initial={project} onSave={updateProject} onCancel={()=>setEditProject(false)}/>;

  return (
    <div style={{background:T.bg,minHeight:"100vh",fontFamily:"inherit"}}>
      {/* Sticky header */}
      <div style={{background:T.surface,borderBottom:`1px solid ${T.border}`,padding:"14px 16px",position:"sticky",top:0,zIndex:50}}>
        <button onClick={onBack} style={{background:"none",border:"none",color:T.sub,fontSize:13,cursor:"pointer",marginBottom:8,padding:0,fontFamily:"inherit"}}>← Projects</button>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:12}}>
          <div style={{flex:1,minWidth:0}}>
            <div style={{fontSize:19,fontWeight:900,color:T.text,letterSpacing:"-0.4px",lineHeight:1.2}}>{project.name}</div>
            <div style={{fontSize:12,color:T.sub,marginTop:2}}>{project.client}{project.location?" · "+project.location:""}</div>
          </div>
          <span style={pill(project.status==="active"?T.green:T.muted)}>{project.status}</span>
        </div>
        <StatBar items={[
          {label:"Reports",val:reports.length,color:T.orange},
          {label:"Labor",val:"$"+(tot.l>=1000?(tot.l/1000).toFixed(1)+"k":fmt(tot.l)),color:T.green},
          {label:"Equip",val:"$"+(tot.e>=1000?(tot.e/1000).toFixed(1)+"k":fmt(tot.e)),color:T.yellow},
          {label:"Total",val:"$"+(tot.g>=1000?(tot.g/1000).toFixed(1)+"k":fmt(tot.g)),color:T.blue},
        ]}/>
        {/* Scrollable tab bar */}
        <div style={{display:"flex",gap:4,marginTop:12,overflowX:"auto",paddingBottom:2,WebkitOverflowScrolling:"touch"}}>
          {PTABS.map(t=>(
            <button key={t.id} onClick={()=>setTab(t.id)} style={{flexShrink:0,background:tab===t.id?T.orange:"transparent",border:tab===t.id?"none":`1px solid ${T.border}`,borderRadius:10,padding:"8px 10px",fontSize:11,fontWeight:tab===t.id?800:500,cursor:"pointer",color:tab===t.id?"#09090B":T.sub,fontFamily:"inherit",whiteSpace:"nowrap"}}>
              {t.icon} {t.label}
            </button>
          ))}
        </div>
      </div>

      <div style={{padding:"14px 16px 80px"}}>
        <ErrBanner msg={err} onDismiss={()=>setErr("")}/>
        {loading&&<Spinner/>}

        {!loading&&tab==="reports"&&(
          <div>
            <button onClick={()=>setScreen("newReport")} style={{...primBtn,marginBottom:14,borderRadius:14}}>+ New Daily Report</button>
            {reports.length===0&&<div style={{textAlign:"center",padding:"32px 0",color:T.muted}}><div style={{fontSize:32,marginBottom:8}}>📋</div><div>No reports yet.</div></div>}
            {reports.map(r=>{
              const t=reportTotals(r);
              const sc={submitted:T.yellow,approved:T.green,flagged:T.red}[r.status||"submitted"]||T.muted;
              return (
                <div key={r.id} onClick={()=>{setActiveReport(r);setScreen("reportDetail");}} style={{...cardS,marginBottom:9,cursor:"pointer",borderLeft:`3px solid ${sc}`,display:"flex",justifyContent:"space-between",alignItems:"center"}}>
                  <div>
                    <div style={{display:"flex",alignItems:"center",gap:8}}>
                      <div style={{fontSize:15,fontWeight:700}}>{fmtShort(r.date)}</div>
                      <span style={pill(sc)}>{(r.status||"submitted").toUpperCase()}</span>
                    </div>
                    <div style={{fontSize:11,color:T.muted,marginTop:4,display:"flex",gap:8}}>
                      {(r.labor||[]).length>0&&<span>👷 {r.labor.length}</span>}
                      {(r.equipment||[]).length>0&&<span>🚜 {r.equipment.length}</span>}
                      {r.submitted_by&&<span>by {r.submitted_by}</span>}
                      {r.status==="flagged"&&<span style={{color:T.red}}>🚩 Flagged</span>}
                    </div>
                  </div>
                  <div style={{textAlign:"right"}}>
                    <div style={{fontSize:17,fontWeight:900,color:T.green}}>${fmt(t.grand)}</div>
                    <div style={{fontSize:9,color:T.muted}}>TOTAL</div>
                  </div>
                </div>
              );
            })}
          </div>
        )}

        {!loading&&tab==="time"&&<TimeCardsTab projectId={project.id} user={user} onErr={setErr}/>}
        {!loading&&tab==="crew"&&<CrewTab projectId={project.id} user={user} onErr={setErr}/>}
        {!loading&&tab==="subs"&&<SubsTab projectId={project.id} user={user} onErr={setErr}/>}
        {!loading&&tab==="safety"&&<SafetyTab projectId={project.id} safety={safety} user={user} onRefresh={()=>load(true)} onErr={setErr}/>}
        {!loading&&tab==="photos"&&<PhotosTab projectId={project.id} photos={photos} onRefresh={()=>load(true)} onErr={setErr}/>}
        {!loading&&tab==="weather"&&<WeatherTab projectId={project.id} project={project} weather={weather} onRefresh={()=>load(true)} onErr={setErr}/>}
        {!loading&&tab==="info"&&<InfoTab project={project} user={user} onEdit={()=>setEditProject(true)} onArchive={archiveProject}/>}
      </div>
    </div>
  );
}

/* ═══════════════════════════════════════════════════════════════
   PM DASHBOARD
═══════════════════════════════════════════════════════════════ */
function PMDashboard({onBack,user}) {
  const [projects,setProjects]=useState([]); const [reports,setReports]=useState([]);
  const [pending,setPending]=useState([]); const [loading,setLoading]=useState(true);
  const [err,setErr]=useState(""); const [pmTab,setPmTab]=useState("overview");
  const [activeReport,setActiveReport]=useState(null); const [activeProject,setActiveProject]=useState(null);

  async function load(){
    setLoading(true);setErr("");
    try{const [projs,reps,pend]=await Promise.all([API.projects.list(),API.reports.all(),API.reports.pending()]);setProjects(projs||[]);setReports(reps||[]);setPending(pend||[]);}
    catch(e){setErr(e.message);}setLoading(false);
  }
  useEffect(()=>{load();},[]);

  async function approve(id){
    try{await API.reports.update(id,{status:"approved",approved_by:user.name,approved_at:new Date().toISOString()});await load();}
    catch(e){setErr(e.message);}
  }
  async function flag(id,notes){
    try{await API.reports.update(id,{status:"flagged",pm_notes:notes});await load();}
    catch(e){setErr(e.message);}
  }

  if(activeReport&&activeProject) return (
    <ReportDetail report={activeReport} project={activeProject} user={user}
      onBack={()=>{setActiveReport(null);setActiveProject(null);load();}}
      onDelete={async(id)=>{await API.reports.remove(id);setActiveReport(null);setActiveProject(null);load();}}
      onApprove={approve} onFlag={flag}
    />
  );

  const allTot=reports.reduce((s,r)=>{const t=reportTotals(r);return{l:s.l+t.labor,e:s.e+t.equip,m:s.m+t.mats,g:s.g+t.grand};},{l:0,e:0,m:0,g:0});
  const thisWeek=reports.filter(r=>{const d=new Date(r.date+"T12:00:00");return(Date.now()-d.getTime())/86400000<=7;});

  // Worker hours summary from all reports this month
  const monthStart=new Date(); monthStart.setDate(1); const ms=monthStart.toISOString().split("T")[0];
  const monthReports=reports.filter(r=>r.date>=ms);
  const workerHours={};
  monthReports.forEach(r=>(r.labor||[]).forEach(l=>{if(!l.name)return;if(!workerHours[l.name])workerHours[l.name]={name:l.name,reg:0,ot:0,travel:0,pay:0};workerHours[l.name].reg+=parseFloat(l.regHrs)||0;workerHours[l.name].ot+=parseFloat(l.otHrs)||0;workerHours[l.name].travel+=parseFloat(l.travelHrs)||0;workerHours[l.name].pay+=laborAmt(l);}));
  const workerRows=Object.values(workerHours).sort((a,b)=>b.pay-a.pay);

  // Project billing breakdown
  const projMap={};
  projects.forEach(p=>{projMap[p.id]={...p,labor:0,equip:0,mats:0,grand:0,count:0};});
  reports.forEach(r=>{if(!projMap[r.project_id])return;const t=reportTotals(r);projMap[r.project_id].labor+=t.labor;projMap[r.project_id].equip+=t.equip;projMap[r.project_id].mats+=t.mats;projMap[r.project_id].grand+=t.grand;projMap[r.project_id].count++;});
  const projRows=Object.values(projMap).filter(p=>p.status==="active").sort((a,b)=>b.grand-a.grand);

  const DMTABS=[{id:"overview",l:"📊 Overview"},{id:"approvals",l:"✅ Approvals"},{id:"workers",l:"👷 Workers"},{id:"billing",l:"💰 Billing"}];

  return (
    <div style={{background:T.bg,minHeight:"100vh",fontFamily:"inherit"}}>
      <div style={{background:T.surface,borderBottom:`1px solid ${T.border}`,padding:"14px 16px",position:"sticky",top:0,zIndex:50}}>
        <button onClick={onBack} style={{background:"none",border:"none",color:T.sub,fontSize:13,cursor:"pointer",marginBottom:8,padding:0,fontFamily:"inherit"}}>← Projects</button>
        <div style={{fontSize:20,fontWeight:900,letterSpacing:"-0.5px",marginBottom:12}}>PM Dashboard</div>
        <div style={{display:"flex",gap:6,overflowX:"auto"}}>
          {DMTABS.map(t=>(
            <button key={t.id} onClick={()=>setPmTab(t.id)} style={{flexShrink:0,background:pmTab===t.id?T.orange:"transparent",border:pmTab===t.id?"none":`1px solid ${T.border}`,borderRadius:10,padding:"9px 12px",fontSize:12,fontWeight:pmTab===t.id?800:500,cursor:"pointer",color:pmTab===t.id?"#09090B":T.sub,fontFamily:"inherit"}}>
              {t.l}{t.id==="approvals"&&pending.length>0?` (${pending.length})`:""}
            </button>
          ))}
        </div>
      </div>

      <div style={{padding:"14px 16px 80px"}}>
        <ErrBanner msg={err} onDismiss={()=>setErr("")}/>
        {loading&&<Spinner/>}
        {!loading&&(
          <>
            {pmTab==="overview"&&(
              <div>
                <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:20}}>
                  {[
                    {icon:"💰",label:"Total Billed (All Time)",val:"$"+fmt(allTot.g),color:T.green},
                    {icon:"📋",label:"Reports This Week",val:thisWeek.length,color:T.orange},
                    {icon:"🏗️",label:"Active Projects",val:projects.filter(p=>p.status==="active").length,color:T.blue},
                    {icon:"✅",label:"Pending Approvals",val:pending.length,color:pending.length>0?T.yellow:T.muted},
                  ].map(k=>(
                    <div key={k.label} style={cardS}>
                      <div style={{fontSize:24,marginBottom:6}}>{k.icon}</div>
                      <div style={{fontSize:22,fontWeight:900,color:k.color,letterSpacing:"-0.5px"}}>{k.val}</div>
                      <div style={{fontSize:11,color:T.muted,marginTop:3}}>{k.label}</div>
                    </div>
                  ))}
                </div>
                <div style={{fontSize:12,fontWeight:700,color:T.muted,textTransform:"uppercase",letterSpacing:"1px",marginBottom:12}}>Recent Reports</div>
                {reports.slice(0,15).map(r=>{
                  const t=reportTotals(r);const p=projects.find(x=>x.id===r.project_id);const sc={submitted:T.yellow,approved:T.green,flagged:T.red}[r.status||"submitted"]||T.muted;
                  return(
                    <div key={r.id} onClick={()=>{setActiveReport(r);setActiveProject(p||{id:r.project_id,name:r.projects?.name||"Unknown",...(p||{})});}} style={{...cardS,marginBottom:8,cursor:"pointer",display:"flex",justifyContent:"space-between",alignItems:"center"}}>
                      <div><div style={{fontSize:13,fontWeight:700}}>{fmtShort(r.date)} · {r.projects?.name||"Unknown"}</div><div style={{fontSize:11,color:T.muted}}>{r.submitted_by||"Unknown"} · <span style={{color:sc}}>{(r.status||"submitted").toUpperCase()}</span></div></div>
                      <div style={{fontSize:15,fontWeight:800,color:T.green}}>${fmt(t.grand)}</div>
                    </div>
                  );
                })}
              </div>
            )}

            {pmTab==="approvals"&&(
              <div>
                {pending.length===0&&<div style={{textAlign:"center",padding:"40px 0",color:T.muted}}><div style={{fontSize:40,marginBottom:10}}>✅</div><div style={{fontSize:16,fontWeight:700}}>All caught up!</div><div style={{fontSize:13,marginTop:4}}>No reports pending approval.</div></div>}
                {pending.map(r=>{
                  const t=reportTotals(r);const p=projects.find(x=>x.id===r.project_id);
                  return(
                    <div key={r.id} style={{...cardS,marginBottom:12,borderLeft:`3px solid ${T.yellow}`}}>
                      <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:10}}>
                        <div>
                          <div style={{fontSize:15,fontWeight:800}}>{fmtShort(r.date)}</div>
                          <div style={{fontSize:12,color:T.sub}}>{r.projects?.name||"Unknown Project"}</div>
                          <div style={{fontSize:12,color:T.muted}}>by {r.submitted_by||"Unknown"}</div>
                          {r.description&&<div style={{fontSize:12,color:T.sub,marginTop:4,lineHeight:1.4}}>{r.description.slice(0,120)}{r.description.length>120?"…":""}</div>}
                        </div>
                        <div style={{textAlign:"right",flexShrink:0,marginLeft:10}}>
                          <div style={{fontSize:18,fontWeight:900,color:T.green}}>${fmt(t.grand)}</div>
                          <div style={{display:"flex",gap:4,marginTop:6}}>
                            {(r.labor||[]).length>0&&<span style={pill(T.orange)}>👷{r.labor.length}</span>}
                            {(r.equipment||[]).length>0&&<span style={pill(T.yellow)}>🚜{r.equipment.length}</span>}
                          </div>
                        </div>
                      </div>
                      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8}}>
                        <button onClick={()=>approve(r.id)} style={{...primBtn,background:T.greenLow,color:T.green,border:`1px solid ${T.green}40`,borderRadius:10,padding:"12px"}}>✓ Approve</button>
                        <button onClick={()=>{const n=window.prompt("Flag note for crew:");if(n!==null)flag(r.id,n);}} style={{...primBtn,background:T.redLow,color:T.red,border:`1px solid ${T.red}40`,borderRadius:10,padding:"12px"}}>🚩 Flag</button>
                      </div>
                      <button onClick={()=>{setActiveReport(r);setActiveProject(p||{id:r.project_id,name:r.projects?.name||"Unknown",...(p||{})});}} style={{...ghostBtn,width:"100%",textAlign:"center",marginTop:8,padding:"10px"}}>View Full Report</button>
                    </div>
                  );
                })}
              </div>
            )}

            {pmTab==="workers"&&(
              <div>
                <div style={{fontSize:12,color:T.muted,marginBottom:14}}>Hours from daily reports this month</div>
                {workerRows.length===0&&<div style={{textAlign:"center",padding:"40px 0",color:T.muted}}><div style={{fontSize:32,marginBottom:8}}>👷</div><div>No labor entries this month yet.</div></div>}
                {workerRows.map(w=>(
                  <div key={w.name} style={{...cardS,marginBottom:9}}>
                    <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start"}}>
                      <div>
                        <div style={{fontSize:15,fontWeight:700}}>{w.name}</div>
                        <div style={{display:"flex",gap:10,marginTop:6}}>
                          <span style={{fontSize:12,color:T.sub}}>{w.reg.toFixed(1)} reg</span>
                          {w.ot>0&&<span style={{fontSize:12,color:T.yellow}}>{w.ot.toFixed(1)} OT</span>}
                          {w.travel>0&&<span style={{fontSize:12,color:T.blue}}>{w.travel.toFixed(1)} travel</span>}
                        </div>
                      </div>
                      <div style={{textAlign:"right"}}>
                        <div style={{fontSize:17,fontWeight:900,color:T.green}}>${fmt(w.pay)}</div>
                        <div style={{fontSize:10,color:T.muted}}>{(w.reg+w.ot+w.travel).toFixed(1)} total hrs</div>
                      </div>
                    </div>
                  </div>
                ))}
              </div>
            )}

            {pmTab==="billing"&&(
              <div>
                <div style={{...cardS,marginBottom:16,background:T.orangeLow,border:`1px solid ${T.orange}40`}}>
                  <div style={{fontSize:12,color:T.orange,fontWeight:700,textTransform:"uppercase",letterSpacing:"1px",marginBottom:10}}>All-Time Totals</div>
                  {[["Labor",allTot.l,T.green],["Equipment",allTot.e,T.yellow],["Materials",allTot.m,T.blue],["Grand Total",allTot.g,T.orange]].map(([l,v,c])=>(
                    <div key={l} style={{display:"flex",justifyContent:"space-between",padding:"8px 0",borderBottom:`1px solid ${T.border}`}}>
                      <span style={{fontSize:14,color:T.sub,fontWeight:l==="Grand Total"?800:400}}>{l}</span>
                      <span style={{fontSize:l==="Grand Total"?20:14,fontWeight:800,color:c}}>${fmt(v)}</span>
                    </div>
                  ))}
                </div>
                <div style={{fontSize:12,fontWeight:700,color:T.muted,textTransform:"uppercase",letterSpacing:"1px",marginBottom:12}}>By Project</div>
                {projRows.map(p=>(
                  <div key={p.id} style={{...cardS,marginBottom:10}}>
                    <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:10}}>
                      <div><div style={{fontSize:15,fontWeight:700}}>{p.name}</div><div style={{fontSize:11,color:T.muted}}>{p.count} reports · {p.location||"No location"}</div></div>
                      <div style={{fontSize:18,fontWeight:900,color:T.green}}>${p.grand>=1000?(p.grand/1000).toFixed(1)+"k":fmt(p.grand)}</div>
                    </div>
                    <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:6}}>
                      {[["Labor",p.labor,T.orange],["Equip",p.equip,T.yellow],["Mats",p.mats,T.blue]].map(([l,v,c])=>(
                        <div key={l} style={{background:T.surface,borderRadius:8,padding:"8px",textAlign:"center"}}>
                          <div style={{fontSize:13,fontWeight:700,color:c}}>${v>=1000?(v/1000).toFixed(1)+"k":fmt(v)}</div>
                          <div style={{fontSize:9,color:T.muted,textTransform:"uppercase",letterSpacing:"0.5px"}}>{l}</div>
                        </div>
                      ))}
                    </div>
                  </div>
                ))}
              </div>
            )}
          </>
        )}
      </div>
    </div>
  );
}

/* ═══════════════════════════════════════════════════════════════
   ROOT APP
═══════════════════════════════════════════════════════════════ */
export default function App() {
  const [user,setUser]     = useState(null);
  const [projects,setProjects] = useState([]);
  const [loading,setLoading]   = useState(false);
  const [err,setErr]           = useState("");
  const [screen,setScreen]     = useState("projects"); // projects|newProject|projectDetail|pmDashboard
  const [activeProject,setActiveProject] = useState(null);

  useEffect(()=>{
    const link=document.createElement("link");
    link.rel="stylesheet";
    link.href="https://fonts.googleapis.com/css2?family=DM+Sans:ital,opsz,wght@0,9..40,400;0,9..40,500;0,9..40,600;0,9..40,700;0,9..40,800;0,9..40,900&display=swap";
    document.head.appendChild(link);
  },[]);

  async function loadProjects() {
    setLoading(true);setErr("");
    try{
      const projs=await API.projects.list();
      const enriched=await Promise.all((projs||[]).map(async p=>{
        try{
          const reps=await API.reports.forProject(p.id);
          const billed=(reps||[]).reduce((s,r)=>s+reportTotals(r).grand,0);
          return {...p,_reports:(reps||[]).length,_billed:billed};
        }catch{return {...p,_reports:0,_billed:0};}
      }));
      setProjects(enriched);
    }catch(e){setErr(e.message);}
    setLoading(false);
  }

  useEffect(()=>{if(user)loadProjects();},[user]);

  async function handleNewProject(data) {
    try{await API.projects.create({...data,created_by:user.name});await loadProjects();setScreen("projects");}
    catch(e){setErr(e.message);}
  }

  function handleProjectUpdated(updated) {
    setActiveProject(p=>({...p,...updated}));
    loadProjects();
  }

  function selectProject(p) { setActiveProject(p); setScreen("projectDetail"); }

  if(!user) return (
    <div style={{maxWidth:480,margin:"0 auto",fontFamily:"'DM Sans',system-ui,sans-serif"}}>
      <LoginScreen onLogin={u=>setUser(u)}/>
    </div>
  );

  return (
    <div style={{maxWidth:480,margin:"0 auto",fontFamily:"'DM Sans',system-ui,sans-serif",background:T.bg,minHeight:"100vh"}}>
      {err&&<div style={{position:"fixed",top:0,left:"50%",transform:"translateX(-50%)",width:"100%",maxWidth:480,zIndex:200,padding:"0 16px"}}>
        <ErrBanner msg={err} onDismiss={()=>setErr("")}/>
      </div>}

      {screen==="pmDashboard"&&<PMDashboard user={user} onBack={()=>setScreen("projects")}/>}

      {screen==="newProject"&&<ProjectForm onSave={handleNewProject} onCancel={()=>setScreen("projects")}/>}

      {screen==="projectDetail"&&activeProject&&(
        <ProjectDetail
          project={projects.find(p=>p.id===activeProject.id)||activeProject}
          user={user}
          onBack={()=>{setScreen("projects");loadProjects();}}
          onProjectUpdated={handleProjectUpdated}
        />
      )}

      {screen==="projects"&&(
        <ProjectsHome
          user={user} projects={projects} loading={loading}
          onSelect={selectProject}
          onNew={()=>setScreen("newProject")}
          onLogout={()=>{setUser(null);setProjects([]);setScreen("projects");}}
          onDash={()=>setScreen("pmDashboard")}
        />
      )}
    </div>
  );
}
