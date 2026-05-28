import { useState, useEffect, useRef } from "react";
import * as XLSX from "xlsx";

/* ── SUPABASE ───────────────────────────────────────────────── */
const SUPA_URL = "https://uicmfyudiullulbbwzmh.supabase.co";
const SUPA_KEY = "sb_publishable_9h9AyvXpkp9glLxDVWRuGw_1eKVS7sE";
async function supa(path,{method="GET",body,prefer}={}){
  const r=await fetch(`${SUPA_URL}/rest/v1${path}`,{method,headers:{"apikey":SUPA_KEY,"Authorization":`Bearer ${SUPA_KEY}`,"Content-Type":"application/json",...(prefer?{"Prefer":prefer}:{}),...(body!==undefined?{body:JSON.stringify(body)}:{}),...(body!==undefined?{}:{})},body:body!==undefined?JSON.stringify(body):undefined});
  if(!r.ok)throw new Error(await r.text()||`HTTP ${r.status}`);
  const t=await r.text();return t?JSON.parse(t):null;
}
// fix: proper fetch helper
async function sb(path,opts={}){
  const{method="GET",body,prefer}=opts;
  const headers={"apikey":SUPA_KEY,"Authorization":`Bearer ${SUPA_KEY}`,"Content-Type":"application/json"};
  if(prefer)headers["Prefer"]=prefer;
  const res=await fetch(`${SUPA_URL}/rest/v1${path}`,{method,headers,...(body!==undefined?{body:JSON.stringify(body)}:{})});
  if(!res.ok)throw new Error(await res.text()||`HTTP ${res.status}`);
  const t=await res.text();return t?JSON.parse(t):null;
}

const API={
  projects:{
    list:()=>sb("/projects?select=*&order=created_at.desc"),
    byDivision:(div)=>sb(`/projects?division=eq.${encodeURIComponent(div)}&select=*&order=created_at.desc`),
    create:(d)=>sb("/projects",{method:"POST",body:d,prefer:"return=representation"}),
    update:(id,d)=>sb(`/projects?id=eq.${id}`,{method:"PATCH",body:d,prefer:"return=representation"}),
    remove:(id)=>sb(`/projects?id=eq.${id}`,{method:"DELETE"}),
  },
  reports:{
    forProject:(pid)=>sb(`/daily_reports?project_id=eq.${pid}&order=date.desc`),
    all:()=>sb("/daily_reports?select=*,projects(id,name,division)&order=date.desc&limit=300"),
    pending:()=>sb("/daily_reports?status=eq.submitted&select=*,projects(id,name,division)&order=created_at.desc"),
    create:(d)=>sb("/daily_reports",{method:"POST",body:d,prefer:"return=representation"}),
    update:(id,d)=>sb(`/daily_reports?id=eq.${id}`,{method:"PATCH",body:d,prefer:"return=representation"}),
    remove:(id)=>sb(`/daily_reports?id=eq.${id}`,{method:"DELETE"}),
  },
  safety:   {forProject:(pid)=>sb(`/safety_logs?project_id=eq.${pid}&order=created_at.desc`),create:(d)=>sb("/safety_logs",{method:"POST",body:d,prefer:"return=representation"}),remove:(id)=>sb(`/safety_logs?id=eq.${id}`,{method:"DELETE"})},
  photos:   {forProject:(pid)=>sb(`/project_photos?project_id=eq.${pid}&order=created_at.desc`),create:(d)=>sb("/project_photos",{method:"POST",body:d,prefer:"return=representation"}),remove:(id)=>sb(`/project_photos?id=eq.${id}`,{method:"DELETE"})},
  timeCards:{forProject:(pid)=>sb(`/time_cards?project_id=eq.${pid}&order=date.desc,created_at.desc`),create:(d)=>sb("/time_cards",{method:"POST",body:d,prefer:"return=representation"}),remove:(id)=>sb(`/time_cards?id=eq.${id}`,{method:"DELETE"})},
  weather:  {forProject:(pid)=>sb(`/weather_logs?project_id=eq.${pid}&order=date.desc&limit=14`),upsert:(d)=>sb("/weather_logs",{method:"POST",body:d,prefer:"return=representation,resolution=merge-duplicates"}),remove:(id)=>sb(`/weather_logs?id=eq.${id}`,{method:"DELETE"})},
  equipment:{forProject:(pid)=>sb(`/equipment_on_site?project_id=eq.${pid}&order=date.desc,created_at.desc`),create:(d)=>sb("/equipment_on_site",{method:"POST",body:d,prefer:"return=representation"}),remove:(id)=>sb(`/equipment_on_site?id=eq.${id}`,{method:"DELETE"})},
  subs:     {forProject:(pid)=>sb(`/subcontractors?project_id=eq.${pid}&order=date.desc,created_at.desc`),create:(d)=>sb("/subcontractors",{method:"POST",body:d,prefer:"return=representation"}),remove:(id)=>sb(`/subcontractors?id=eq.${id}`,{method:"DELETE"})},
  docs:     {forProject:(pid)=>sb(`/documents?project_id=eq.${pid}&order=created_at.desc`),create:(d)=>sb("/documents",{method:"POST",body:d,prefer:"return=representation"}),remove:(id)=>sb(`/documents?id=eq.${id}`,{method:"DELETE"})},
  milestones:{forProject:(pid)=>sb(`/milestones?project_id=eq.${pid}&order=sort_order.asc,target_date.asc`),create:(d)=>sb("/milestones",{method:"POST",body:d,prefer:"return=representation"}),update:(id,d)=>sb(`/milestones?id=eq.${id}`,{method:"PATCH",body:d,prefer:"return=representation"}),remove:(id)=>sb(`/milestones?id=eq.${id}`,{method:"DELETE"})},
  crew:     {list:()=>sb("/crew_members?order=name.asc"),create:(d)=>sb("/crew_members",{method:"POST",body:d,prefer:"return=representation"}),update:(id,d)=>sb(`/crew_members?id=eq.${id}`,{method:"PATCH",body:d,prefer:"return=representation"}),remove:(id)=>sb(`/crew_members?id=eq.${id}`,{method:"DELETE"})},
  notifications:{list:()=>sb("/notifications?order=created_at.desc&limit=50"),unread:()=>sb("/notifications?read=eq.false&order=created_at.desc"),markRead:(id)=>sb(`/notifications?id=eq.${id}`,{method:"PATCH",body:{read:true}}),markAllRead:()=>sb("/notifications?read=eq.false",{method:"PATCH",body:{read:true}}),create:(d)=>sb("/notifications",{method:"POST",body:d,prefer:"return=representation"})},
  notifSettings:{get:(name)=>sb(`/notification_settings?pm_name=eq.${encodeURIComponent(name)}&limit=1`),upsert:(d)=>sb("/notification_settings",{method:"POST",body:d,prefer:"return=representation,resolution=merge-duplicates"})},
  userProfiles:{
    list:()=>sb("/user_profiles?order=name.asc"),
    getByName:(name)=>sb(`/user_profiles?name=eq.${encodeURIComponent(name)}&limit=1`),
    create:(d)=>sb("/user_profiles",{method:"POST",body:d,prefer:"return=representation"}),
    update:(id,d)=>sb(`/user_profiles?id=eq.${id}`,{method:"PATCH",body:d,prefer:"return=representation"}),
    upsert:(d)=>sb("/user_profiles",{method:"POST",body:d,prefer:"return=representation,resolution=merge-duplicates"}),
    remove:(id)=>sb(`/user_profiles?id=eq.${id}`,{method:"DELETE"}),
  },
};

/* ── THEME ──────────────────────────────────────────────────── */
const T={bg:"#09090B",surface:"#111113",card:"#18181B",border:"#27272A",orange:"#F97316",orangeLow:"#F9731614",orangeMid:"#F9731630",green:"#22C55E",greenLow:"#22C55E14",red:"#EF4444",redLow:"#EF444414",yellow:"#EAB308",yellowLow:"#EAB30814",blue:"#3B82F6",blueLow:"#3B82F614",purple:"#A855F7",teal:"#14B8A6",text:"#FAFAFA",sub:"#A1A1AA",muted:"#52525B"};
const inp={width:"100%",boxSizing:"border-box",background:"#0C0C0F",border:`1px solid ${T.border}`,borderRadius:12,color:T.text,fontSize:15,padding:"13px 14px",outline:"none",fontFamily:"inherit",appearance:"none",WebkitAppearance:"none"};
const lbl={display:"block",fontSize:11,fontWeight:700,color:T.muted,letterSpacing:"1px",textTransform:"uppercase",marginBottom:6};
const cardS={background:T.card,border:`1px solid ${T.border}`,borderRadius:16,padding:"16px"};
const pill=(c)=>({display:"inline-flex",alignItems:"center",background:c+"20",color:c,borderRadius:20,padding:"3px 10px",fontSize:11,fontWeight:700});
const primBtn={width:"100%",background:T.orange,color:"#09090B",border:"none",borderRadius:14,padding:"16px",fontSize:16,fontWeight:800,cursor:"pointer",fontFamily:"inherit",display:"flex",alignItems:"center",justifyContent:"center",gap:8};
const ghostBtn={background:"transparent",border:`1px solid ${T.border}`,borderRadius:12,padding:"12px 16px",color:T.sub,fontSize:14,cursor:"pointer",fontFamily:"inherit",fontWeight:600};
const dangerBtn={background:T.redLow,border:`1px solid ${T.red}30`,borderRadius:12,padding:"12px 16px",color:T.red,fontSize:14,cursor:"pointer",fontFamily:"inherit",fontWeight:600,width:"100%",textAlign:"center"};

/* ── DATA ───────────────────────────────────────────────────── */
const POSITIONS=[{name:"Project Manager",rate:64.50},{name:"Foreman",rate:63.25},{name:"Technician",rate:60.75},{name:"Inspector",rate:53.75},{name:"Certified Welder",rate:60.75},{name:"Fitter",rate:58.50},{name:"Mechanic",rate:58.50},{name:"Operator",rate:58.50},{name:"Truck Driver",rate:58.50},{name:"Helper (Welder)",rate:57.25},{name:"Laborer",rate:51.00},{name:"Foreman (Elect)",rate:82.25},{name:"Electrician",rate:82.25},{name:"Helper (Elect)",rate:45.50},{name:"Per Diem",rate:190.00,flat:true}];
const NAMES=["Aden Walter","Alan Fairbrother","Alan Robinson","Brandon Milano","Brandon Stach","Charles Dovel","Chris Utz","Christopher Dean","Chuck Dean","Clay Lau","Connor Kestner","Eric Bowens","Howard Lau","Jeff White","Jessica Vance","John Baier","John P. Cosner Jr.","Jordan Gorwell","Joseph Lau","Josh Gladhill","Kevin Gabrish","Kurt Batterton","Leo Velez","Mark Hamilton","Amanda Harvey","Thomas Hatfield","Edgardo Ruiz","Jaden Pugh","Matthew Linton","Mike Gamble","Mike McBride","Mike Meridith","Mike Seiler","Pat Gorman","Paul Howard","Rich Raborg","Robert Neslein","Roland Long","Shane Hower","Steve Kestner","Tom Hatfield","Troy Strother","Doug Friedel","Charles Acree","Jackson Fama","Morgan Schramek","Eric Shumate","Tyrone Davis","Walter Chicas-Luna","Will Wychulis","Wyatt Gill"}];
const EQUIP_LIST=[{section:"Trucks & Trailers"},{name:"Truck - 1 Ton",rate:21.50,unit:"Hours"},{name:"Truck - 3/4 Ton w/ Snow Plow",rate:350,unit:"Days"},{name:"Truck - 1/2 Ton",rate:18.50,unit:"Hours"},{name:"Truck - Boom (20-29 Ton)",rate:65,unit:"Hours"},{name:"Truck - Bucket",rate:45,unit:"Hours"},{name:"Truck - Dump Truck (3 Axle)",rate:35,unit:"Hours"},{name:"Truck - Haul Truck - No Trailer",rate:70,unit:"Hours"},{name:"Truck - Tru-Vac",rate:13500,unit:"Month"},{name:"Truck - Welding Rig",rate:35,unit:"Hours"},{name:"Trailer - Electrical - Colonial",rate:147,unit:"Month"},{name:"Trailer - Lowboy - 2 Axle",rate:28,unit:"Hours"},{name:"Trailer - Tag Along",rate:50,unit:"Days"},{name:"Trailer - Tool Trailer - 18-25ft",rate:175,unit:"Days"},{name:"Trailer - Tool Trailer - 26-40ft",rate:200,unit:"Days"},{section:"Earthmoving & ROW"},{name:"ATV - 4 Wheel",rate:125,unit:"Days"},{name:"Backhoe Loader - 80-105 HP",rate:62.45,unit:"Hours"},{name:"Excavator - Mini - 2-8K LB",rate:299,unit:"Days"},{name:"Excavator - Mini - 9K LB",rate:335,unit:"Days"},{name:"Excavator - Mini - 12-16K LB",rate:475,unit:"Days"},{name:"Excavator - Small - 21-29K LB",rate:565,unit:"Days"},{name:"Excavator - Small - 30-33K LB",rate:632,unit:"Days"},{name:"Excavator - Medium - 48-55K LB",rate:852,unit:"Days"},{name:"Excavator - Large - 80-89K LB",rate:1050,unit:"Days"},{name:"Excavator - Large - 90-119K LB",rate:1350,unit:"Days"},{name:"Skidsteer Loader - 70-80 HP",rate:440,unit:"Days"},{name:"Skidsteer Loader - 81-100 HP",rate:475,unit:"Days"},{name:"Tractor - 50 HP 4x4 w/ Bush Hog",rate:36.50,unit:"Hours"},{name:"Mower - Riding/Zero Turn",rate:175,unit:"Days"},{section:"Air, Compressors & Blast"},{name:"Air Compressor - 185 CFM",rate:195,unit:"Days"},{name:"Air Compressor - 375 CFM",rate:275,unit:"Days"},{name:"Air Impact Wrench - 1in",rate:50,unit:"Days"},{name:"Air Spade / Knife",rate:55,unit:"Days"},{name:"Blast Rig - 4 Bag Pot w/ 185 CFM AC",rate:55.50,unit:"Hours"},{name:"Blast Rig - 1 Pot w/ 375 CFM AC",rate:500,unit:"Days"},{section:"Testing & Misc. Tools"},{name:"Holiday Detector / Pipe Jeep",rate:72,unit:"Days"},{name:"Hydraulic Torque",rate:200,unit:"Days"},{name:"Hydro Test Pump",rate:60,unit:"Days"},{name:"Hydrotest - High Pressure",rate:3800,unit:"Days"},{name:"Jack Hammer",rate:72,unit:"Days"},{name:"LEL/Gas Monitor - 4 Gas",rate:50,unit:"Days"},{name:"Line Locator",rate:50,unit:"Days"},{name:"HEPA Vacuum",rate:100,unit:"Days"},{name:"Torque Wrench w/Sockets Hyd/Pneu",rate:195,unit:"Days"},{name:"Pipe Beveling Machine 16-22in",rate:100,unit:"Days"}];
const WMO={0:["Clear Sky","☀️"],1:["Mainly Clear","🌤️"],2:["Partly Cloudy","⛅"],3:["Overcast","☁️"],45:["Foggy","🌫️"],48:["Icy Fog","🌫️"],51:["Light Drizzle","🌦️"],53:["Drizzle","🌦️"],55:["Heavy Drizzle","🌦️"],61:["Light Rain","🌧️"],63:["Rain","🌧️"],65:["Heavy Rain","🌧️"],71:["Light Snow","🌨️"],73:["Snow","🌨️"],75:["Heavy Snow","❄️"],80:["Light Showers","🌦️"],81:["Showers","🌦️"],82:["Violent Showers","⛈️"],95:["Thunderstorm","⛈️"],96:["Thunderstorm + Hail","⛈️"],99:["Severe Thunderstorm","⛈️"]};
const DIVISIONS=["Mechanical","Pipeline","Structural"];
const DIV_META={Mechanical:{icon:"⚙️",color:"#F97316",desc:"Mechanical projects and equipment"},Pipeline:{icon:"🔧",color:"#3B82F6",desc:"Pipeline construction and maintenance"},Structural:{icon:"🏗️",color:"#22C55E",desc:"Structural steel and civil work"}};
const ROLES=["crew","foreman","pm","admin"];
const ROLE_META={crew:{label:"Field Crew",color:T.green,desc:"Submit daily reports and time cards"},foreman:{label:"Foreman",color:T.yellow,desc:"Reports, time, safety, equipment, docs, schedule"},pm:{label:"Project Manager",color:T.orange,desc:"Approve reports, manage jobs, PM dashboard"},admin:{label:"Admin",color:T.red,desc:"Full access, user management"}};

/* ── PERMISSIONS ────────────────────────────────────────────── */
const PERMS={
  admin:  ["manage_users","create_job","edit_job","archive_job","approve_report","flag_report","view_dashboard","submit_report","time_card","safety","photos","docs","schedule","weather","subs","crew_equip","crew_directory","custom_reports","notifications"],
  pm:     ["create_job","edit_job","archive_job","approve_report","flag_report","view_dashboard","submit_report","time_card","safety","photos","docs","schedule","weather","subs","crew_equip","crew_directory","custom_reports","notifications"],
  foreman:["submit_report","time_card","safety","photos","docs","schedule","weather","subs","crew_equip"],
  crew:   ["submit_report","time_card","photos"],
};
const can=(user,action)=>(PERMS[user?.role]||PERMS.crew).includes(action);

/* ── HELPERS ────────────────────────────────────────────────── */
const uid=()=>Math.random().toString(36).slice(2,9);
const today=()=>new Date().toISOString().split("T")[0];
const fmt=(n)=>Number(n||0).toLocaleString("en-US",{minimumFractionDigits:2,maximumFractionDigits:2});
const fmtDate=(d)=>d?new Date(d+"T12:00:00").toLocaleDateString("en-US",{month:"short",day:"numeric",year:"numeric"}):"—";
const fmtShort=(d)=>d?new Date(d+"T12:00:00").toLocaleDateString("en-US",{weekday:"short",month:"short",day:"numeric"}):"—";
const daysUntil=(d)=>{if(!d)return null;const diff=new Date(d+"T12:00:00")-new Date();return Math.ceil(diff/86400000);};
function laborAmt(r){const p=POSITIONS.find(x=>x.name===r.classification);if(!p)return 0;if(p.flat)return p.rate;return p.rate*((parseFloat(r.regHrs)||0)+(parseFloat(r.otHrs)||0)*1.5+(parseFloat(r.travelHrs)||0));}
function equipAmt(r){return(parseFloat(r.rate)||0)*(parseFloat(r.qty)||0)*(parseFloat(r.usage)||0);}
function reportTotals(r){const labor=(r.labor||[]).reduce((s,x)=>s+laborAmt(x),0);const equip=(r.equipment||[]).reduce((s,x)=>s+equipAmt(x),0);const mats=(r.materials||[]).reduce((s,x)=>s+(parseFloat(x.amount)||0),0);return{labor,equip,mats,grand:labor+equip+mats};}
function calcHours(ci,co){if(!ci||!co)return 0;const[ih,im]=ci.split(":").map(Number);const[oh,om]=co.split(":").map(Number);const diff=(oh*60+om)-(ih*60+im);return diff>0?Math.round(diff/60*100)/100:0;}
function getWeekStart(){const d=new Date();const day=d.getDay();d.setDate(d.getDate()-(day===0?6:day-1));return d.toISOString().split("T")[0];}
async function compressImg(file,maxW=900,q=0.65){return new Promise(res=>{const rd=new FileReader();rd.onload=ev=>{const img=new Image();img.onload=()=>{const sc=Math.min(1,maxW/img.width);const c=document.createElement("canvas");c.width=Math.round(img.width*sc);c.height=Math.round(img.height*sc);c.getContext("2d").drawImage(img,0,0,c.width,c.height);res(c.toDataURL("image/jpeg",q));};img.src=ev.target.result;};rd.readAsDataURL(file);});}
async function fetchWeather(location){const gR=await fetch(`https://geocoding-api.open-meteo.com/v1/search?name=${encodeURIComponent(location)}&count=1&language=en&format=json`);const gD=await gR.json();if(!gD.results?.length)throw new Error(`Cannot find: "${location}"`);const{latitude:lat,longitude:lon,name,admin1}=gD.results[0];const wR=await fetch(`https://api.open-meteo.com/v1/forecast?latitude=${lat}&longitude=${lon}&current=temperature_2m,apparent_temperature,weathercode,windspeed_10m,precipitation&daily=temperature_2m_max,temperature_2m_min,precipitation_sum,weathercode&timezone=auto&temperature_unit=fahrenheit&windspeed_unit=mph&precipitation_unit=inch&forecast_days=1`);const wD=await wR.json();return{...wD,locationName:`${name}, ${admin1}`};}
async function notify(type,title,body,extra={}){try{await API.notifications.create({type,title,body,...extra});}catch{}}

/* ── SHARED UI ──────────────────────────────────────────────── */
function Spinner(){return(<div style={{display:"flex",justifyContent:"center",padding:"48px 0"}}><div style={{width:32,height:32,border:`3px solid ${T.border}`,borderTopColor:T.orange,borderRadius:"50%",animation:"spin 0.7s linear infinite"}}/><style>{`@keyframes spin{to{transform:rotate(360deg)}}`}</style></div>);}
function ErrBanner({msg,onDismiss}){if(!msg)return null;return(<div style={{background:T.redLow,border:`1px solid ${T.red}40`,borderRadius:12,padding:"12px 14px",marginBottom:14,display:"flex",justifyContent:"space-between",alignItems:"center"}}><span style={{fontSize:13,color:T.red}}>⚠️ {msg}</span><button onClick={onDismiss} style={{background:"none",border:"none",color:T.red,cursor:"pointer",fontSize:18,padding:"0 0 0 10px"}}>×</button></div>);}
function Lightbox({src,onClose}){if(!src)return null;return(<div onClick={onClose} style={{position:"fixed",inset:0,zIndex:999,background:"rgba(0,0,0,0.95)",display:"flex",alignItems:"center",justifyContent:"center",padding:16}}><img src={src} alt="" style={{maxWidth:"100%",maxHeight:"90vh",borderRadius:12}} onClick={e=>e.stopPropagation()}/><button onClick={onClose} style={{position:"absolute",top:16,right:16,background:"#333",border:"none",color:"#fff",borderRadius:"50%",width:36,height:36,fontSize:18,cursor:"pointer"}}>×</button></div>);}
function DashedAdd({label,onClick,color}){const c=color||T.muted;return(<button onClick={onClick} style={{width:"100%",border:`2px dashed ${c}50`,background:c+"08",color:c,borderRadius:14,padding:"14px",fontSize:15,cursor:"pointer",fontFamily:"inherit",fontWeight:600}}>{label}</button>);}
function StatBar({items}){return(<div style={{display:"grid",gridTemplateColumns:`repeat(${items.length},1fr)`,gap:8}}>{items.map(({label,val,color})=>(<div key={label} style={{background:T.card,border:`1px solid ${T.border}`,borderRadius:12,padding:"10px 8px",textAlign:"center"}}><div style={{fontSize:16,fontWeight:900,color:color||T.text}}>{val}</div><div style={{fontSize:9,color:T.muted,textTransform:"uppercase",letterSpacing:"0.7px",marginTop:2}}>{label}</div></div>))}</div>);}
function TopBar({title,sub,onBack,right}){return(<div style={{background:T.surface,borderBottom:`1px solid ${T.border}`,padding:"14px 16px",position:"sticky",top:0,zIndex:50}}>{onBack&&<button onClick={onBack} style={{background:"none",border:"none",color:T.sub,fontSize:13,cursor:"pointer",marginBottom:8,padding:0,fontFamily:"inherit"}}>← Back</button>}<div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start"}}><div style={{flex:1,minWidth:0}}><div style={{fontSize:20,fontWeight:900,color:T.text,letterSpacing:"-0.5px"}}>{title}</div>{sub&&<div style={{fontSize:12,color:T.muted,marginTop:2}}>{sub}</div>}</div>{right&&<div style={{flexShrink:0,marginLeft:12}}>{right}</div>}</div></div>);}

/* ── FORM CARDS ─────────────────────────────────────────────── */
function LaborCard({row,onChange,onRemove}){const pos=POSITIONS.find(p=>p.name===row.classification);const amt=laborAmt(row);const set=(k,v)=>{const u={...row,[k]:v};if(k==="classification"){const p=POSITIONS.find(x=>x.name===v);u.rate=p?p.rate:"";}onChange(u);};return(<div style={{...cardS,marginBottom:10,borderLeft:`3px solid ${T.orange}`}}><div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8,marginBottom:8}}><div style={{gridColumn:"1/-1"}}><label style={lbl}>Name</label><select value={row.name||""} onChange={e=>set("name",e.target.value)} style={inp}><option value="">— Select —</option>{NAMES.map(n=><option key={n}>{n}</option>)}</select></div><div style={{gridColumn:"1/-1"}}><label style={lbl}>Classification</label><select value={row.classification||""} onChange={e=>set("classification",e.target.value)} style={inp}><option value="">— Select —</option>{POSITIONS.map(p=><option key={p.name}>{p.name}</option>)}</select></div></div>{pos&&!pos.flat&&(<div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:8,marginBottom:8}}>{[["regHrs","Reg Hrs"],["otHrs","OT Hrs"],["travelHrs","Travel"]].map(([k,l])=>(<div key={k}><label style={lbl}>{l}</label><input type="number" min="0" step="0.5" placeholder="0" value={row[k]||""} onChange={e=>set(k,e.target.value)} style={inp}/></div>))}</div>)}<div style={{display:"flex",justifyContent:"space-between",alignItems:"center",paddingTop:8,borderTop:`1px solid ${T.border}`}}><span style={{fontSize:11,color:T.muted}}>{pos?`$${pos.rate.toFixed(2)}${pos.flat?" flat":"/hr"}`:""}</span><div style={{display:"flex",alignItems:"center",gap:10}}>{amt>0&&<span style={{fontSize:16,fontWeight:800,color:T.green}}>${fmt(amt)}</span>}<button onClick={onRemove} style={{background:"none",border:"none",color:T.red,cursor:"pointer",fontSize:20,padding:0}}>×</button></div></div></div>);}
function EquipCard({row,onChange,onRemove}){const eq=EQUIP_LIST.find(e=>!e.section&&e.name===row.description);const amt=equipAmt(row);const set=(k,v)=>{const u={...row,[k]:v};if(k==="description"){const e=EQUIP_LIST.find(x=>!x.section&&x.name===v);u.rate=e?e.rate:"";u.unit=e?e.unit:"";}onChange(u);};return(<div style={{...cardS,marginBottom:10,borderLeft:`3px solid ${T.yellow}`}}><div style={{marginBottom:8}}><label style={lbl}>Equipment</label><select value={row.description||""} onChange={e=>set("description",e.target.value)} style={inp}><option value="">— Select —</option>{EQUIP_LIST.map((e,i)=>e.section?<option key={i} disabled>── {e.section} ──</option>:<option key={i} value={e.name}>{e.name}</option>)}</select></div><div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8,marginBottom:8}}><div><label style={lbl}>Qty</label><input type="number" min="0" placeholder="0" value={row.qty||""} onChange={e=>set("qty",e.target.value)} style={inp}/></div><div><label style={lbl}>{eq?eq.unit:"Hrs/Days"}</label><input type="number" min="0" step="0.5" placeholder="0" value={row.usage||""} onChange={e=>set("usage",e.target.value)} style={inp}/></div></div><div style={{display:"flex",justifyContent:"space-between",alignItems:"center",paddingTop:8,borderTop:`1px solid ${T.border}`}}><span style={{fontSize:11,color:T.muted}}>{eq?`$${eq.rate.toLocaleString()}/${eq.unit}`:""}</span><div style={{display:"flex",alignItems:"center",gap:10}}>{amt>0&&<span style={{fontSize:16,fontWeight:800,color:T.green}}>${fmt(amt)}</span>}<button onClick={onRemove} style={{background:"none",border:"none",color:T.red,cursor:"pointer",fontSize:20,padding:0}}>×</button></div></div></div>);}
function MatCard({row,onChange,onRemove}){const fileRef=useRef(null);const receipts=row.receipts||[];async function handleFiles(files){const n=[];for(const f of files){if(!f.type.startsWith("image/"))continue;const src=await compressImg(f,800,0.6);n.push({id:uid(),src});}onChange({...row,receipts:[...receipts,...n]});}return(<div style={{...cardS,marginBottom:10,borderLeft:`3px solid ${T.blue}`}}><div style={{display:"grid",gridTemplateColumns:"56px 1fr 88px",gap:8,marginBottom:10}}><div><label style={lbl}>Qty</label><input type="number" min="0" placeholder="0" value={row.qty||""} onChange={e=>onChange({...row,qty:e.target.value})} style={inp}/></div><div><label style={lbl}>Description</label><input type="text" placeholder="Item / material" value={row.description||""} onChange={e=>onChange({...row,description:e.target.value})} style={inp}/></div><div><label style={lbl}>Amount</label><input type="number" min="0" placeholder="0.00" value={row.amount||""} onChange={e=>onChange({...row,amount:e.target.value})} style={inp}/></div></div><div style={{borderTop:`1px solid ${T.border}`,paddingTop:10}}><label style={{...lbl,marginBottom:8}}>📎 Receipts</label><div style={{display:"flex",gap:8,flexWrap:"wrap",alignItems:"center"}}>{receipts.map(r=>(<div key={r.id} style={{position:"relative"}}><img src={r.src} alt="" style={{width:60,height:60,objectFit:"cover",borderRadius:10,border:`2px solid ${T.blue}40`,display:"block"}}/><button onClick={()=>onChange({...row,receipts:receipts.filter(x=>x.id!==r.id)})} style={{position:"absolute",top:-5,right:-5,width:18,height:18,borderRadius:"50%",background:T.red,border:"none",color:"#fff",fontSize:11,cursor:"pointer",display:"flex",alignItems:"center",justifyContent:"center"}}>×</button></div>))}<button onClick={()=>fileRef.current?.click()} style={{width:60,height:60,borderRadius:10,border:`2px dashed ${T.blue}40`,background:T.blueLow,color:T.blue,cursor:"pointer",display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",fontSize:18,gap:2}}><span>📷</span><span style={{fontSize:9,fontWeight:700}}>ADD</span></button><input ref={fileRef} type="file" accept="image/*" capture="environment" multiple style={{display:"none"}} onChange={e=>{handleFiles(Array.from(e.target.files));e.target.value="";}} /></div></div><div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginTop:10}}>{row.amount>0&&<span style={{fontSize:14,fontWeight:700,color:T.green}}>${fmt(parseFloat(row.amount)||0)}</span>}<button onClick={onRemove} style={{background:"none",border:"none",color:T.red,cursor:"pointer",fontSize:12,fontWeight:700,fontFamily:"inherit",marginLeft:"auto"}}>Remove</button></div></div>);}

/* ── LOGIN SCREEN ───────────────────────────────────────────── */
function LoginScreen({onLogin}){
  const [name,setName]=useState("");
  const [pin,setPin]=useState("");
  const [profile,setProfile]=useState(null);
  const [loading,setLoading]=useState(false);
  const [err,setErr]=useState("");

  async function handleNameChange(n){
    setName(n); setPin(""); setErr(""); setProfile(null);
    if(!n) return;
    setLoading(true);
    try{
      const rows=await API.userProfiles.getByName(n);
      setProfile(rows&&rows.length>0?rows[0]:{name:n,role:"crew",division:null});
    }catch{setProfile({name:n,role:"crew",division:null});}
    setLoading(false);
  }

  async function handleLogin(){
    if(!name||!profile)return;
    if((profile.role==="admin"||profile.role==="pm")&&pin!==profile.pin){
      setErr("Incorrect PIN"); return;
    }
    onLogin(profile);
  }

  const needsPin=profile&&(profile.role==="admin"||profile.role==="pm");
  const roleM=profile?ROLE_META[profile.role]:null;

  return(
    <div style={{minHeight:"100vh",background:T.bg,display:"flex",flexDirection:"column",justifyContent:"center",padding:24,fontFamily:"inherit"}}>
      <div style={{textAlign:"center",marginBottom:40}}>
        <div style={{display:"inline-flex",alignItems:"center",gap:12,marginBottom:12}}>
          <div style={{width:60,height:60,background:T.orange,borderRadius:20,display:"flex",alignItems:"center",justifyContent:"center",fontSize:30,fontWeight:900,color:"#09090B"}}>A</div>
          <div style={{textAlign:"left"}}>
            <div style={{fontSize:28,fontWeight:900,color:T.text,letterSpacing:"-1px",lineHeight:1.1}}>AIME</div>
            <div style={{fontSize:12,color:T.muted,letterSpacing:"3px",textTransform:"uppercase"}}>Field OS</div>
          </div>
        </div>
        <div style={{fontSize:13,color:T.muted}}>Colonial Pipeline · Field Management</div>
      </div>

      <div style={{...cardS,maxWidth:400,margin:"0 auto",width:"100%"}}>
        <ErrBanner msg={err} onDismiss={()=>setErr("")}/>
        <div style={{marginBottom:14}}>
          <label style={lbl}>Select Your Name</label>
          <select value={name} onChange={e=>handleNameChange(e.target.value)} style={inp}>
            <option value="">— Select your name —</option>
            {NAMES.map(n=><option key={n}>{n}</option>)}
            <option value="Admin">Admin</option>
          </select>
        </div>

        {loading&&<div style={{textAlign:"center",padding:"10px 0",color:T.muted,fontSize:13}}>Looking up profile…</div>}

        {profile&&!loading&&(
          <div style={{...cardS,marginBottom:14,background:T.surface,borderLeft:`3px solid ${roleM?.color||T.border}`}}>
            <div style={{display:"flex",alignItems:"center",gap:10}}>
              <div style={{width:36,height:36,borderRadius:10,background:(roleM?.color||T.muted)+"20",display:"flex",alignItems:"center",justifyContent:"center",fontSize:18}}>{profile.role==="admin"?"🔴":profile.role==="pm"?"🟠":profile.role==="foreman"?"🟡":"🟢"}</div>
              <div>
                <div style={{fontSize:14,fontWeight:700,color:T.text}}>{roleM?.label||"Field Crew"}</div>
                <div style={{fontSize:12,color:T.muted}}>{profile.division||"All Divisions"}</div>
              </div>
            </div>
          </div>
        )}

        {needsPin&&(
          <div style={{marginBottom:14}}>
            <label style={lbl}>{profile.role==="admin"?"Admin PIN":"PM PIN"}</label>
            <input type="password" maxLength={6} placeholder="Enter PIN" value={pin} onChange={e=>setPin(e.target.value)} onKeyDown={e=>e.key==="Enter"&&handleLogin()} style={inp}/>
          </div>
        )}

        <button onClick={handleLogin} style={{...primBtn,opacity:name&&!loading?1:0.45}} disabled={!name||loading}>
          {loading?"Loading…":"Sign In →"}
        </button>
      </div>
    </div>
  );
}

/* ── DIVISION SELECTION SCREEN ──────────────────────────────── */
function DivisionScreen({user,projects,onSelect,onLogout,onCrew,onDash}){
  const divStats=DIVISIONS.map(div=>{
    const divProjects=projects.filter(p=>p.division===div&&p.status==="active");
    const totalBilled=divProjects.reduce((s,p)=>s+(p._billed||0),0);
    const totalReports=divProjects.reduce((s,p)=>s+(p._reports||0),0);
    return{div,count:divProjects.length,billed:totalBilled,reports:totalReports};
  });

  return(
    <div style={{background:T.bg,minHeight:"100vh",fontFamily:"inherit"}}>
      {/* Header */}
      <div style={{background:T.surface,borderBottom:`1px solid ${T.border}`,padding:"16px"}}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:6}}>
          <div style={{display:"flex",alignItems:"center",gap:10}}>
            <div style={{width:40,height:40,background:T.orange,borderRadius:12,display:"flex",alignItems:"center",justifyContent:"center",fontWeight:900,fontSize:20,color:"#09090B"}}>A</div>
            <div>
              <div style={{fontSize:16,fontWeight:800,color:T.text}}>AIME Field OS</div>
              <div style={{fontSize:11,color:T.muted}}>
                {user.role==="admin"?"🔴":user.role==="pm"?"🟠":user.role==="foreman"?"🟡":"🟢"} {user.name} · {ROLE_META[user.role]?.label}
              </div>
            </div>
          </div>
          <div style={{display:"flex",gap:8}}>
            {can(user,"view_dashboard")&&<button onClick={onDash} style={{background:T.orangeLow,border:`1px solid ${T.orange}40`,borderRadius:10,padding:"8px 12px",color:T.orange,fontSize:12,fontWeight:700,cursor:"pointer",fontFamily:"inherit"}}>📊</button>}
            {can(user,"crew_directory")&&<button onClick={onCrew} style={{background:T.blueLow,border:`1px solid ${T.blue}40`,borderRadius:10,padding:"8px 12px",color:T.blue,fontSize:12,fontWeight:700,cursor:"pointer",fontFamily:"inherit"}}>👥</button>}
            <button onClick={onLogout} style={{...ghostBtn,padding:"8px 12px",fontSize:12}}>Out</button>
          </div>
        </div>
      </div>

      {/* Division picker */}
      <div style={{padding:"20px 16px 80px"}}>
        <div style={{marginBottom:24}}>
          <div style={{fontSize:22,fontWeight:900,color:T.text,letterSpacing:"-0.5px",marginBottom:4}}>Select Division</div>
          <div style={{fontSize:13,color:T.muted}}>Choose the division you are working in today</div>
        </div>

        {DIVISIONS.map((div,i)=>{
          const meta=DIV_META[div];
          const stats=divStats.find(s=>s.div===div);
          const divColor=meta.color;
          return(
            <div key={div} onClick={()=>onSelect(div)} style={{
              background:T.card,
              border:`1px solid ${T.border}`,
              borderRadius:20,
              marginBottom:14,
              cursor:"pointer",
              overflow:"hidden",
              transition:"transform 0.1s",
            }}>
              {/* Top gradient bar */}
              <div style={{height:4,background:`linear-gradient(90deg,${divColor},${divColor}88)`}}/>
              <div style={{padding:"20px 20px 16px"}}>
                <div style={{display:"flex",alignItems:"center",gap:14,marginBottom:14}}>
                  <div style={{width:56,height:56,borderRadius:16,background:divColor+"20",border:`2px solid ${divColor}40`,display:"flex",alignItems:"center",justifyContent:"center",fontSize:28,flexShrink:0}}>
                    {meta.icon}
                  </div>
                  <div style={{flex:1}}>
                    <div style={{fontSize:22,fontWeight:900,color:T.text,letterSpacing:"-0.5px"}}>{div}</div>
                    <div style={{fontSize:13,color:T.sub,marginTop:2}}>{meta.desc}</div>
                  </div>
                  <div style={{fontSize:22,color:divColor}}>→</div>
                </div>
                {/* Stats */}
                <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:8,paddingTop:12,borderTop:`1px solid ${T.border}`}}>
                  {[["Active Jobs",stats?.count||0,divColor],["Reports",stats?.reports||0,T.green],["Billed","$"+(stats?.billed>=1000?(stats.billed/1000).toFixed(1)+"k":fmt(stats?.billed||0)),T.blue]].map(([l,v,c])=>(
                    <div key={l} style={{textAlign:"center"}}>
                      <div style={{fontSize:15,fontWeight:800,color:c}}>{v}</div>
                      <div style={{fontSize:9,color:T.muted,textTransform:"uppercase",letterSpacing:"0.6px",marginTop:2}}>{l}</div>
                    </div>
                  ))}
                </div>
              </div>
            </div>
          );
        })}
      </div>
    </div>
  );
}

/* ── JOB BOARD (per division) ───────────────────────────────── */
function JobBoard({user,division,projects,loading,onSelect,onNew,onBack}){
  const [search,setSearch]=useState("");
  const [filter,setFilter]=useState("active");
  const meta=DIV_META[division]||{icon:"🏗️",color:T.orange};

  const divProjects=projects.filter(p=>p.division===division);
  const filtered=divProjects.filter(p=>{
    const ms=filter==="all"?true:p.status===filter;
    const q=search.toLowerCase();
    const ms2=!q||p.name?.toLowerCase().includes(q)||p.location?.toLowerCase().includes(q)||p.afe?.toLowerCase().includes(q)||p.client?.toLowerCase().includes(q);
    return ms&&ms2;
  });
  const active=divProjects.filter(p=>p.status==="active");

  return(
    <div style={{background:T.bg,minHeight:"100vh",fontFamily:"inherit"}}>
      <div style={{background:T.surface,borderBottom:`1px solid ${T.border}`,padding:"14px 16px 0"}}>
        <button onClick={onBack} style={{background:"none",border:"none",color:T.sub,fontSize:13,cursor:"pointer",marginBottom:10,padding:0,fontFamily:"inherit"}}>← Divisions</button>
        <div style={{display:"flex",alignItems:"center",gap:12,marginBottom:14}}>
          <div style={{width:44,height:44,borderRadius:14,background:meta.color+"20",border:`2px solid ${meta.color}40`,display:"flex",alignItems:"center",justifyContent:"center",fontSize:22,flexShrink:0}}>{meta.icon}</div>
          <div style={{flex:1}}>
            <div style={{fontSize:20,fontWeight:900,color:T.text,letterSpacing:"-0.5px"}}>{division}</div>
            <div style={{fontSize:11,color:T.muted}}>{active.length} active job{active.length!==1?"s":""}</div>
          </div>
          {can(user,"create_job")&&<button onClick={onNew} style={{background:T.orange,color:"#09090B",border:"none",borderRadius:12,padding:"10px 16px",fontSize:13,fontWeight:800,cursor:"pointer",fontFamily:"inherit"}}>+ New Job</button>}
        </div>
        {/* Search */}
        <div style={{position:"relative",marginBottom:10}}>
          <span style={{position:"absolute",left:13,top:"50%",transform:"translateY(-50%)",fontSize:14,pointerEvents:"none"}}>🔍</span>
          <input type="text" placeholder="Search jobs…" value={search} onChange={e=>setSearch(e.target.value)} style={{...inp,paddingLeft:38,borderRadius:12,fontSize:14}}/>
          {search&&<button onClick={()=>setSearch("")} style={{position:"absolute",right:12,top:"50%",transform:"translateY(-50%)",background:"none",border:"none",color:T.muted,cursor:"pointer",fontSize:18,padding:0}}>×</button>}
        </div>
        {/* Tabs */}
        <div style={{display:"flex",gap:6}}>
          {[["active","Active"],["archived","Archived"],["all","All"]].map(([v,l])=>(<button key={v} onClick={()=>setFilter(v)} style={{padding:"8px 14px",borderRadius:"10px 10px 0 0",background:filter===v?T.bg:"transparent",border:filter===v?`1px solid ${T.border}`:"1px solid transparent",borderBottom:filter===v?`1px solid ${T.bg}`:"none",color:filter===v?T.text:T.muted,fontSize:13,fontWeight:filter===v?700:500,cursor:"pointer",fontFamily:"inherit",position:"relative",zIndex:filter===v?1:0,marginBottom:filter===v?-1:0}}>{l}{v==="active"&&active.length>0&&<span style={{marginLeft:5,background:meta.color+"25",color:meta.color,borderRadius:20,padding:"1px 6px",fontSize:10,fontWeight:800}}>{active.length}</span>}</button>))}
        </div>
      </div>

      <div style={{padding:"12px 16px 80px"}}>
        {loading&&<Spinner/>}
        {!loading&&filtered.length===0&&(
          <div style={{textAlign:"center",padding:"60px 20px",color:T.muted}}>
            <div style={{fontSize:48,marginBottom:12}}>{meta.icon}</div>
            <div style={{fontSize:17,fontWeight:700,color:T.sub,marginBottom:6}}>{search?`No jobs matching "${search}"`:filter==="archived"?"No archived jobs":"No active jobs in "+division}</div>
            {!search&&filter==="active"&&can(user,"create_job")&&<div style={{fontSize:13}}>Tap + New Job to create one.</div>}
          </div>
        )}
        {!loading&&filtered.map(p=><JobCard key={p.id} p={p} onSelect={onSelect} divColor={meta.color}/>)}
      </div>
    </div>
  );
}

function JobCard({p,onSelect,divColor}){
  const isArchived=p.status!=="active";
  const daysSince=p._lastReport?Math.floor((Date.now()-new Date(p._lastReport+"T12:00:00").getTime())/86400000):null;
  const actColor=daysSince===null?T.muted:daysSince===0?T.green:daysSince<=2?T.orange:T.red;
  const actLabel=daysSince===null?"No reports yet":daysSince===0?"Reported today":daysSince===1?"Reported yesterday":`${daysSince}d since last report`;
  const c=divColor||T.orange;
  return(
    <div onClick={()=>onSelect(p)} style={{background:T.card,border:`1px solid ${T.border}`,borderRadius:16,marginBottom:10,cursor:"pointer",overflow:"hidden",opacity:isArchived?0.55:1}}>
      <div style={{height:3,background:isArchived?T.border:`linear-gradient(90deg,${c},${c}88)`}}/>
      <div style={{padding:"14px 16px"}}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:8}}>
          <div style={{flex:1,minWidth:0,paddingRight:12}}>
            <div style={{fontSize:17,fontWeight:900,color:T.text,letterSpacing:"-0.3px",lineHeight:1.2}}>{p.name}</div>
            <div style={{fontSize:12,color:T.sub,marginTop:3}}>{[p.client,p.location].filter(Boolean).join(" · ")||"No details"}</div>
          </div>
          <div style={{textAlign:"right",flexShrink:0}}>
            <div style={{fontSize:20,fontWeight:900,color:T.green,letterSpacing:"-0.5px"}}>${(p._billed||0)>=1000?((p._billed||0)/1000).toFixed(1)+"k":fmt(p._billed||0)}</div>
            <div style={{fontSize:9,color:T.muted,textTransform:"uppercase",letterSpacing:"0.5px",marginTop:1}}>Total Billed</div>
          </div>
        </div>
        <div style={{display:"flex",gap:6,flexWrap:"wrap",marginBottom:10}}>
          {p.afe&&<span style={pill(T.muted)}>AFE: {p.afe}</span>}
          {p.work_order&&<span style={pill(T.muted)}>WO: {p.work_order}</span>}
          <span style={pill(isArchived?T.muted:T.green)}>{isArchived?"Archived":"Active"}</span>
        </div>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",paddingTop:10,borderTop:`1px solid ${T.border}`}}>
          <div style={{display:"flex",gap:16}}>
            <div><div style={{fontSize:16,fontWeight:800,color:c}}>{p._reports||0}</div><div style={{fontSize:9,color:T.muted,textTransform:"uppercase",letterSpacing:"0.5px"}}>Reports</div></div>
            <div><div style={{fontSize:11,fontWeight:700,color:actColor,marginTop:2}}>{actLabel}</div><div style={{fontSize:9,color:T.muted,textTransform:"uppercase",letterSpacing:"0.5px",marginTop:1}}>Last Activity</div></div>
          </div>
          <div style={{display:"flex",alignItems:"center",gap:6,background:c+"15",border:`1px solid ${c}40`,borderRadius:10,padding:"8px 14px"}}>
            <span style={{fontSize:13,fontWeight:700,color:c}}>Enter Job</span>
            <span style={{fontSize:16,color:c}}>→</span>
          </div>
        </div>
      </div>
    </div>
  );
}

/* ── PROJECT FORM ───────────────────────────────────────────── */
function ProjectForm({initial,onSave,onCancel,saving,defaultDivision}){
  const [f,setF]=useState(initial||{name:"",client:"",location:"",afe:"",work_order:"",start_date:today(),notes:"",status:"active",division:defaultDivision||"Pipeline"});
  const set=(k,v)=>setF(x=>({...x,[k]:v}));
  return(
    <div style={{background:T.bg,minHeight:"100vh",fontFamily:"inherit"}}>
      <TopBar title={initial?"Edit Job":"New Job"} onBack={onCancel}/>
      <div style={{padding:"16px 16px 100px"}}>
        {[{k:"name",l:"Job Name *",ph:"e.g. HDD Crossing – Station 42"},{k:"client",l:"Client",ph:"Colonial Pipeline"},{k:"location",l:"Location",ph:"City, State or Milepost"},{k:"afe",l:"AFE No.",ph:"AFE #"},{k:"work_order",l:"Work Order #",ph:"WO #"}].map(({k,l,ph})=>(<div key={k} style={{marginBottom:12}}><label style={lbl}>{l}</label><input type="text" placeholder={ph} value={f[k]||""} onChange={e=>set(k,e.target.value)} style={inp}/></div>))}
        <div style={{marginBottom:12}}>
          <label style={lbl}>Division</label>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:8}}>
            {DIVISIONS.map(div=>{const m=DIV_META[div];return(<button key={div} onClick={()=>set("division",div)} style={{padding:"12px 8px",borderRadius:12,border:`2px solid ${f.division===div?m.color:T.border}`,background:f.division===div?m.color+"20":T.surface,color:f.division===div?m.color:T.sub,fontWeight:700,fontSize:12,cursor:"pointer",fontFamily:"inherit",textAlign:"center"}}><div style={{fontSize:20,marginBottom:4}}>{m.icon}</div>{div}</button>);})}
          </div>
        </div>
        <div style={{marginBottom:12}}><label style={lbl}>Start Date</label><input type="date" value={f.start_date||today()} onChange={e=>set("start_date",e.target.value)} style={inp}/></div>
        <div style={{marginBottom:20}}><label style={lbl}>Notes</label><textarea placeholder="Project notes, scope, special instructions…" value={f.notes||""} onChange={e=>set("notes",e.target.value)} rows={3} style={{...inp,resize:"vertical",lineHeight:1.5}}/></div>
        <button onClick={()=>f.name.trim()&&!saving&&onSave(f)} style={{...primBtn,opacity:f.name.trim()&&!saving?1:0.5}}>{saving?"Saving…":initial?"Save Changes":"Create Job"}</button>
      </div>
    </div>
  );
}

/* ── DAILY REPORT FORM ──────────────────────────────────────── */
const RSTEPS=["Job Info","Labor","Equipment","Materials","Review"];
function DailyReportForm({user,project,onSave,onCancel}){
  const [step,setStep]=useState(1);const [saving,setSaving]=useState(false);
  const [rpt,setRpt]=useState({date:today(),description:"",report_no:"",labor:[],equipment:[],materials:[]});
  const topRef=useRef(null);
  const setR=(k,v)=>setRpt(r=>({...r,[k]:v}));
  function add(key,item){setR(key,[...rpt[key],item]);}
  function upd(key,i,row){const a=[...rpt[key]];a[i]=row;setR(key,a);}
  function del(key,i){setR(key,rpt[key].filter((_,j)=>j!==i));}
  const tot=reportTotals(rpt);
  async function submit(){setSaving(true);await onSave({...rpt,submitted_by:user.name,status:"submitted"});await notify("report_submitted","New Report Submitted",`${user.name} submitted a report for ${project.name}`,{project_id:project.id});setSaving(false);}
  const scroll=()=>topRef.current?.scrollIntoView({behavior:"smooth"});
  const divMeta=DIV_META[project.division]||{color:T.orange};
  return(
    <div ref={topRef} style={{background:T.bg,minHeight:"100vh",fontFamily:"inherit"}}>
      <div style={{background:T.surface,borderBottom:`1px solid ${T.border}`,padding:"14px 16px",position:"sticky",top:0,zIndex:50}}>
        <div style={{display:"flex",justifyContent:"space-between",marginBottom:12}}><div style={{fontSize:16,fontWeight:800}}>New Daily Report</div><button onClick={onCancel} style={{background:"none",border:"none",color:T.sub,cursor:"pointer",fontSize:13,fontFamily:"inherit"}}>Cancel</button></div>
        <div style={{display:"flex",alignItems:"center"}}>{RSTEPS.map((s,i)=>(<div key={i} style={{display:"flex",alignItems:"center",flex:i<RSTEPS.length-1?1:undefined}}><div style={{display:"flex",flexDirection:"column",alignItems:"center",gap:3}}><div style={{width:26,height:26,borderRadius:"50%",background:i+1<step?T.green:i+1===step?divMeta.color:T.border,color:i+1<=step?"#09090B":T.muted,display:"flex",alignItems:"center",justifyContent:"center",fontSize:11,fontWeight:800}}>{i+1<step?"✓":i+1}</div><div style={{fontSize:8,color:i+1===step?divMeta.color:T.muted,fontWeight:i+1===step?700:400,whiteSpace:"nowrap"}}>{s}</div></div>{i<RSTEPS.length-1&&<div style={{flex:1,height:2,background:i+1<step?T.green:T.border,margin:"0 3px",marginBottom:14}}/>}</div>))}</div>
      </div>
      <div style={{padding:"14px 16px 100px"}}>
        {step===1&&(<div>
          <div style={{...cardS,marginBottom:14,borderLeft:`3px solid ${divMeta.color}`}}><div style={{fontSize:11,color:T.muted,marginBottom:2}}>Project · {project.division}</div><div style={{fontSize:15,fontWeight:700}}>{project.name}</div>{project.afe&&<div style={{fontSize:12,color:T.sub}}>AFE: {project.afe}{project.work_order?" · WO: "+project.work_order:""}</div>}</div>
          <div style={{marginBottom:12}}><label style={lbl}>Date</label><input type="date" value={rpt.date} onChange={e=>setR("date",e.target.value)} style={inp}/></div>
          <div style={{marginBottom:12}}><label style={lbl}>Report No.</label><input type="text" placeholder="Report #" value={rpt.report_no||""} onChange={e=>setR("report_no",e.target.value)} style={inp}/></div>
          <div style={{marginBottom:12}}><label style={lbl}>Description of Work Done</label><textarea placeholder="Describe the work performed today…" value={rpt.description||""} onChange={e=>setR("description",e.target.value)} rows={4} style={{...inp,resize:"vertical",lineHeight:1.5}}/></div>
        </div>)}
        {step===2&&(<div><div style={{display:"flex",justifyContent:"space-between",marginBottom:12}}><div style={{fontSize:17,fontWeight:800}}>👷 Labor</div>{tot.labor>0&&<div style={{fontSize:16,fontWeight:800,color:T.green}}>${fmt(tot.labor)}</div>}</div>{rpt.labor.map((row,i)=><LaborCard key={row.id} row={row} onChange={r=>upd("labor",i,r)} onRemove={()=>del("labor",i)}/>)}<DashedAdd label="+ Add Worker" onClick={()=>add("labor",{id:uid(),name:"",classification:"",regHrs:"",otHrs:"",travelHrs:""})} color={T.orange}/></div>)}
        {step===3&&(<div><div style={{display:"flex",justifyContent:"space-between",marginBottom:12}}><div style={{fontSize:17,fontWeight:800}}>🚜 Equipment</div>{tot.equip>0&&<div style={{fontSize:16,fontWeight:800,color:T.green}}>${fmt(tot.equip)}</div>}</div>{rpt.equipment.map((row,i)=><EquipCard key={row.id} row={row} onChange={r=>upd("equipment",i,r)} onRemove={()=>del("equipment",i)}/>)}<DashedAdd label="+ Add Equipment" onClick={()=>add("equipment",{id:uid(),description:"",qty:"",usage:"",rate:"",unit:""})} color={T.yellow}/></div>)}
        {step===4&&(<div><div style={{fontSize:17,fontWeight:800,marginBottom:12}}>📦 Materials & Misc.</div>{rpt.materials.map((row,i)=><MatCard key={row.id} row={row} onChange={r=>upd("materials",i,r)} onRemove={()=>del("materials",i)}/>)}<DashedAdd label="+ Add Material / Item" onClick={()=>add("materials",{id:uid(),qty:"",description:"",amount:"",receipts:[]})} color={T.blue}/></div>)}
        {step===5&&(<div>
          <div style={{fontSize:17,fontWeight:800,marginBottom:12}}>✅ Review & Submit</div>
          <div style={{...cardS,marginBottom:12}}>
            <div style={{fontSize:11,color:divMeta.color,fontWeight:700,textTransform:"uppercase",letterSpacing:"1px",marginBottom:10}}>Summary</div>
            {[["Project",project.name],["Division",project.division],["Date",fmtDate(rpt.date)],["Report No.",rpt.report_no||"—"],["Workers",rpt.labor.length],["Equipment",rpt.equipment.length+" items"],["Materials",rpt.materials.length+" items"]].map(([l,v])=>(<div key={l} style={{display:"flex",justifyContent:"space-between",padding:"7px 0",borderBottom:`1px solid ${T.border}`}}><span style={{fontSize:13,color:T.sub}}>{l}</span><span style={{fontSize:13,fontWeight:600}}>{v}</span></div>))}
            {[["Labor",tot.labor,T.green],["Equipment",tot.equip,T.green],["Materials",tot.mats,T.green]].map(([l,v,c])=>(<div key={l} style={{display:"flex",justifyContent:"space-between",padding:"7px 0",borderBottom:`1px solid ${T.border}`}}><span style={{fontSize:13,color:T.sub}}>{l}</span><span style={{fontSize:13,fontWeight:700,color:c}}>${fmt(v)}</span></div>))}
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",paddingTop:14}}><span style={{fontSize:16,fontWeight:800}}>Grand Total</span><span style={{fontSize:26,fontWeight:900,color:divMeta.color,letterSpacing:"-1px"}}>${fmt(tot.grand)}</span></div>
          </div>
        </div>)}
      </div>
      <div style={{position:"fixed",bottom:0,left:"50%",transform:"translateX(-50%)",width:"100%",maxWidth:480,background:T.bg+"EE",backdropFilter:"blur(12px)",borderTop:`1px solid ${T.border}`,padding:"12px 16px",display:"flex",gap:10}}>
        {step>1&&<button onClick={()=>{setStep(s=>s-1);scroll();}} style={{...ghostBtn,flex:1}}>← Back</button>}
        {step<5?<button onClick={()=>{setStep(s=>s+1);scroll();}} style={{...primBtn,flex:2,borderRadius:12,background:divMeta.color}}>{step===4?"Review →":"Next →"}</button>:<button onClick={submit} style={{...primBtn,flex:2,borderRadius:12,background:divMeta.color,opacity:saving?0.6:1}}>{saving?"Saving…":"💾 Save Report"}</button>}
      </div>
    </div>
  );
}

/* ── REPORT DETAIL ──────────────────────────────────────────── */
function ReportDetail({report,project,user,onBack,onDelete,onApprove,onFlag}){
  const [lb,setLb]=useState(null);const [flagNote,setFlagNote]=useState("");const [flagging,setFlagging]=useState(false);
  const tot=reportTotals(report);
  const sc={submitted:T.yellow,approved:T.green,flagged:T.red}[report.status]||T.muted;
  const divColor=DIV_META[project.division]?.color||T.orange;
  function exportXLSX(){
    const wb=XLSX.utils.book_new();const rows=[];const blank=()=>Array(11).fill(null);
    const r1=blank();r1[1]="COLONIAL PIPELINE COMPANY";rows.push(r1);
    const r2=blank();r2[1]="DAILY REPORT-WORK PERFORMED BY CONTRACTOR";rows.push(r2);
    const r3=blank();r3[1]="LOCATION";r3[3]="AFE NO.";r3[4]="WORK ORDER\nNUMBER";r3[6]="REPORT DATE";r3[8]="REPORT\nNO.";rows.push(r3);
    const[yr,mo,dy]=(report.date||"").split("-");
    const r4=blank();r4[1]=project.location||"";r4[3]=project.afe||"";r4[4]=project.work_order||"";r4[6]=`${mo}/${dy}/${yr}`;r4[8]=report.report_no||"";rows.push(r4);
    const r5=blank();r5[1]="CONTRACTOR:";r5[2]="AIME";r5[8]="CONTRACTOR DATE:";r5[9]=report.date||"";rows.push(r5);
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
    const sgRow=blank();sgRow[1]="VERIFIED AND ACCEPTED BY CO. REP";sgRow[3]="DATE";sgRow[4]="CERTIFIED AS CORRECT BY CONTRACTOR";rows.push(sgRow);
    const ws=XLSX.utils.aoa_to_sheet(rows);
    ws["!cols"]=[{wch:5.7},{wch:15},{wch:11.7},{wch:17.1},{wch:12.7},{wch:13.1},{wch:10},{wch:10},{wch:24.9},{wch:23.7},{wch:10.1}];
    const rng=XLSX.utils.decode_range(ws["!ref"]);
    for(let r=0;r<=rng.e.r;r++){const a=XLSX.utils.encode_cell({r,c:9});if(ws[a]&&typeof ws[a].v==="number")ws[a].z='"$"#,##0.00';const b=XLSX.utils.encode_cell({r,c:5});if(ws[b]&&typeof ws[b].v==="number")ws[b].z='"$"#,##0.00';}
    XLSX.utils.book_append_sheet(wb,ws,"Daily Report");
    XLSX.writeFile(wb,`AIME_${project.name.replace(/\s+/g,"_")}_${(report.date||"").replace(/-/g,"")}.xlsx`);
  }
  return(
    <div style={{background:T.bg,minHeight:"100vh",padding:16,fontFamily:"inherit"}}>
      <Lightbox src={lb} onClose={()=>setLb(null)}/>
      <button onClick={onBack} style={{...ghostBtn,marginBottom:14}}>← Reports</button>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:4}}><div style={{fontSize:20,fontWeight:900,letterSpacing:"-0.5px"}}>{fmtDate(report.date)}</div><span style={pill(sc)}>{(report.status||"submitted").toUpperCase()}</span></div>
      {report.submitted_by&&<div style={{fontSize:12,color:T.muted,marginBottom:14}}>by {report.submitted_by}</div>}
      {report.pm_notes&&<div style={{...cardS,marginBottom:14,borderLeft:`3px solid ${T.red}`,background:T.redLow}}><div style={{fontSize:11,color:T.red,fontWeight:700,marginBottom:4}}>🚩 PM NOTE</div><div style={{fontSize:13,color:T.sub}}>{report.pm_notes}</div></div>}
      {report.description&&<div style={{...cardS,marginBottom:12,borderLeft:`3px solid ${T.blue}`}}><div style={{fontSize:11,color:T.muted,textTransform:"uppercase",letterSpacing:"1px",marginBottom:6}}>Work Done</div><div style={{fontSize:14,color:T.sub,lineHeight:1.6}}>{report.description}</div></div>}
      {(report.labor||[]).length>0&&<div style={{...cardS,marginBottom:12}}><div style={{fontSize:12,color:divColor,fontWeight:700,textTransform:"uppercase",letterSpacing:"1px",marginBottom:10}}>Labor · <span style={{color:T.green}}>${fmt(tot.labor)}</span></div>{report.labor.map((r,i)=>(<div key={i} style={{display:"flex",justifyContent:"space-between",padding:"8px 0",borderBottom:i<report.labor.length-1?`1px solid ${T.border}`:"none"}}><div><div style={{fontSize:14,fontWeight:600}}>{r.name||"—"}</div><div style={{fontSize:11,color:T.muted}}>{r.classification} · {r.regHrs||0}reg {r.otHrs||0}OT {r.travelHrs||0}tr</div></div><div style={{fontSize:14,fontWeight:800,color:T.green}}>${fmt(laborAmt(r))}</div></div>))}</div>}
      {(report.equipment||[]).length>0&&<div style={{...cardS,marginBottom:12}}><div style={{fontSize:12,color:divColor,fontWeight:700,textTransform:"uppercase",letterSpacing:"1px",marginBottom:10}}>Equipment · <span style={{color:T.green}}>${fmt(tot.equip)}</span></div>{report.equipment.map((r,i)=>(<div key={i} style={{display:"flex",justifyContent:"space-between",padding:"8px 0",borderBottom:i<report.equipment.length-1?`1px solid ${T.border}`:"none"}}><div style={{flex:1,paddingRight:10}}><div style={{fontSize:13,fontWeight:600}}>{r.description}</div><div style={{fontSize:11,color:T.muted}}>Qty {r.qty} x {r.usage} {r.unit}</div></div><div style={{fontSize:14,fontWeight:800,color:T.green}}>${fmt(equipAmt(r))}</div></div>))}</div>}
      {(report.materials||[]).length>0&&<div style={{...cardS,marginBottom:12}}><div style={{fontSize:12,color:divColor,fontWeight:700,textTransform:"uppercase",letterSpacing:"1px",marginBottom:10}}>Materials · <span style={{color:T.green}}>${fmt(tot.mats)}</span></div>{report.materials.map((r,i)=>(<div key={i} style={{padding:"8px 0",borderBottom:i<report.materials.length-1?`1px solid ${T.border}`:"none"}}><div style={{display:"flex",justifyContent:"space-between",marginBottom:r.receipts?.length>0?8:0}}><span style={{fontSize:13}}>{r.qty?`${r.qty}x `:""}{r.description}</span><span style={{fontSize:13,fontWeight:700,color:T.green}}>${fmt(parseFloat(r.amount)||0)}</span></div>{r.receipts?.length>0&&<div style={{display:"flex",gap:6,flexWrap:"wrap"}}>{r.receipts.map(rc=><img key={rc.id} src={rc.src} alt="" onClick={()=>setLb(rc.src)} style={{width:56,height:56,objectFit:"cover",borderRadius:8,cursor:"pointer"}}/>)}</div>}</div>))}</div>}
      <div style={{...cardS,background:divColor+"12",border:`1px solid ${divColor}40`,marginBottom:12,display:"flex",justifyContent:"space-between",alignItems:"center"}}><span style={{fontSize:15,fontWeight:800}}>Grand Total</span><span style={{fontSize:26,fontWeight:900,color:divColor,letterSpacing:"-1px"}}>${fmt(tot.grand)}</span></div>
      {can(user,"approve_report")&&report.status==="submitted"&&(<div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:10}}><button onClick={()=>onApprove&&onApprove(report.id)} style={{...primBtn,background:T.greenLow,color:T.green,border:`1px solid ${T.green}40`,borderRadius:12}}>✓ Approve</button><button onClick={()=>setFlagging(!flagging)} style={{...primBtn,background:T.redLow,color:T.red,border:`1px solid ${T.red}40`,borderRadius:12}}>🚩 Flag</button></div>)}
      {flagging&&<div style={{...cardS,marginBottom:10}}><label style={lbl}>Flag Note for Crew</label><textarea value={flagNote} onChange={e=>setFlagNote(e.target.value)} rows={3} placeholder="What needs to be corrected…" style={{...inp,resize:"vertical",marginBottom:10}}/><button onClick={()=>{onFlag&&onFlag(report.id,flagNote);setFlagging(false);}} style={{...primBtn,borderRadius:12}}>Send Flag</button></div>}
      <button onClick={exportXLSX} style={{...primBtn,background:divColor+"15",color:divColor,border:`1px solid ${divColor}40`,marginBottom:10,borderRadius:14}}>📥 Export to Excel (.xlsx)</button>
      <button onClick={()=>window.confirm("Delete this report?")&&onDelete(report.id)} style={dangerBtn}>🗑 Delete Report</button>
    </div>
  );
}

/* ── TIME CARDS TAB ─────────────────────────────────────────── */
function TimeCardsTab({projectId,user,onErr}){
  const [cards,setCards]=useState([]);const [loading,setLoading]=useState(true);const [showForm,setShowForm]=useState(false);const [saving,setSaving]=useState(false);
  const [f,setF]=useState({worker_name:user.name,date:today(),clock_in:"07:00",clock_out:"",notes:""});
  const weekStart=getWeekStart();
  async function load(){setLoading(true);try{setCards(await API.timeCards.forProject(projectId)||[]);}catch(e){onErr(e.message);}setLoading(false);}
  useEffect(()=>{load();},[projectId]);
  async function save(){if(!f.worker_name||!f.date)return;setSaving(true);const total_hours=calcHours(f.clock_in,f.clock_out);const ot_hours=Math.max(0,total_hours-8);try{await API.timeCards.create({...f,project_id:projectId,total_hours,ot_hours});await load();setShowForm(false);setF({worker_name:user.name,date:today(),clock_in:"07:00",clock_out:"",notes:""});}catch(e){onErr(e.message);}setSaving(false);}
  async function remove(id){try{await API.timeCards.remove(id);await load();}catch(e){onErr(e.message);}}
  const weekCards=cards.filter(c=>c.date>=weekStart);
  const byWorker={};weekCards.forEach(c=>{if(!byWorker[c.worker_name])byWorker[c.worker_name]={name:c.worker_name,reg:0,ot:0,total:0};byWorker[c.worker_name].total+=(c.total_hours||0);byWorker[c.worker_name].ot+=(c.ot_hours||0);byWorker[c.worker_name].reg+=Math.max(0,(c.total_hours||0)-(c.ot_hours||0));});
  const workerRows=Object.values(byWorker).sort((a,b)=>b.total-a.total);
  const todayCards=cards.filter(c=>c.date===today());
  const recentCards=cards.filter(c=>c.date!==today()).slice(0,30);
  return(<div>
    <button onClick={()=>setShowForm(!showForm)} style={{...primBtn,marginBottom:14,borderRadius:14}}>{showForm?"✕ Cancel":"⏱️ Log Time"}</button>
    {showForm&&<div style={{...cardS,marginBottom:14,borderLeft:`3px solid ${T.green}`}}>
      <div style={{marginBottom:10}}><label style={lbl}>Worker</label><select value={f.worker_name} onChange={e=>setF(x=>({...x,worker_name:e.target.value}))} style={inp}>{NAMES.map(n=><option key={n}>{n}</option>)}</select></div>
      <div style={{marginBottom:10}}><label style={lbl}>Date</label><input type="date" value={f.date} onChange={e=>setF(x=>({...x,date:e.target.value}))} style={inp}/></div>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:10}}><div><label style={lbl}>Clock In</label><input type="time" value={f.clock_in} onChange={e=>setF(x=>({...x,clock_in:e.target.value}))} style={inp}/></div><div><label style={lbl}>Clock Out</label><input type="time" value={f.clock_out} onChange={e=>setF(x=>({...x,clock_out:e.target.value}))} style={inp}/></div></div>
      {f.clock_in&&f.clock_out&&(()=>{const h=calcHours(f.clock_in,f.clock_out);const ot=Math.max(0,h-8);return h>0&&(<div style={{background:T.greenLow,borderRadius:10,padding:"10px 12px",marginBottom:10,display:"flex",gap:16}}><div><div style={{fontSize:18,fontWeight:900,color:T.green}}>{h.toFixed(2)}h</div><div style={{fontSize:10,color:T.muted,textTransform:"uppercase"}}>Total</div></div><div><div style={{fontSize:18,fontWeight:900,color:ot>0?T.yellow:T.muted}}>{Math.min(h,8).toFixed(2)}h</div><div style={{fontSize:10,color:T.muted,textTransform:"uppercase"}}>Regular</div></div>{ot>0&&<div><div style={{fontSize:18,fontWeight:900,color:T.yellow}}>{ot.toFixed(2)}h</div><div style={{fontSize:10,color:T.muted,textTransform:"uppercase"}}>OT</div></div>}</div>);})()}
      <div style={{marginBottom:10}}><label style={lbl}>Notes</label><input type="text" placeholder="Optional…" value={f.notes} onChange={e=>setF(x=>({...x,notes:e.target.value}))} style={inp}/></div>
      <button onClick={save} style={{...primBtn,background:T.green,color:"#09090B",borderRadius:12}}>{saving?"Saving…":"Save Time Card"}</button>
    </div>}
    {loading&&<Spinner/>}
    {!loading&&<>{workerRows.length>0&&<div style={{...cardS,marginBottom:14}}><div style={{fontSize:12,fontWeight:700,color:T.green,textTransform:"uppercase",letterSpacing:"1px",marginBottom:10}}>This Week</div>{workerRows.map(w=>(<div key={w.name} style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"8px 0",borderBottom:`1px solid ${T.border}`}}><span style={{fontSize:14,fontWeight:600}}>{w.name}</span><div style={{display:"flex",gap:12,alignItems:"center"}}><span style={{fontSize:12,color:T.muted}}>{w.reg.toFixed(1)}reg</span>{w.ot>0&&<span style={{fontSize:12,color:T.yellow}}>{w.ot.toFixed(1)}OT</span>}<span style={{fontSize:15,fontWeight:800,color:T.green}}>{w.total.toFixed(1)}h</span></div></div>))}</div>}
    {todayCards.length>0&&<div style={{fontSize:12,fontWeight:700,color:T.muted,textTransform:"uppercase",letterSpacing:"1px",marginBottom:10}}>Today</div>}
    {todayCards.map(c=><div key={c.id} style={{...cardS,marginBottom:8,display:"flex",justifyContent:"space-between",alignItems:"center"}}><div><div style={{fontSize:14,fontWeight:700}}>{c.worker_name}</div><div style={{fontSize:11,color:T.muted,marginTop:3}}>{fmtShort(c.date)}{c.clock_in?" · "+c.clock_in:""}{c.clock_out?" → "+c.clock_out:""}</div>{c.notes&&<div style={{fontSize:11,color:T.sub,marginTop:2}}>{c.notes}</div>}</div><div style={{display:"flex",alignItems:"center",gap:10}}><div style={{textAlign:"right"}}><div style={{fontSize:16,fontWeight:800,color:T.green}}>{(c.total_hours||0).toFixed(1)}h</div>{(c.ot_hours||0)>0&&<div style={{fontSize:10,color:T.yellow}}>{c.ot_hours.toFixed(1)} OT</div>}</div><button onClick={()=>remove(c.id)} style={{background:"none",border:"none",color:T.muted,cursor:"pointer",fontSize:16,padding:0}}>🗑</button></div></div>)}
    {recentCards.length>0&&<div style={{fontSize:12,fontWeight:700,color:T.muted,textTransform:"uppercase",letterSpacing:"1px",margin:"14px 0 10px"}}>Recent</div>}
    {recentCards.map(c=><div key={c.id} style={{...cardS,marginBottom:8,display:"flex",justifyContent:"space-between",alignItems:"center"}}><div><div style={{fontSize:14,fontWeight:700}}>{c.worker_name}</div><div style={{fontSize:11,color:T.muted,marginTop:3}}>{fmtShort(c.date)}{c.clock_in?" · "+c.clock_in:""}{c.clock_out?" → "+c.clock_out:""}</div></div><div style={{display:"flex",alignItems:"center",gap:10}}><div style={{fontSize:16,fontWeight:800,color:T.green}}>{(c.total_hours||0).toFixed(1)}h</div><button onClick={()=>remove(c.id)} style={{background:"none",border:"none",color:T.muted,cursor:"pointer",fontSize:16,padding:0}}>🗑</button></div></div>)}
    {cards.length===0&&!showForm&&<div style={{textAlign:"center",padding:"40px 0",color:T.muted}}><div style={{fontSize:32,marginBottom:8}}>⏱️</div><div>No time cards yet.</div></div>}</>}
  </div>);
}

/* ── CREW/EQUIP ON SITE TAB ─────────────────────────────────── */
function CrewEquipTab({projectId,user,onErr}){
  const [equip,setEquip]=useState([]);const [loading,setLoading]=useState(true);const [showForm,setShowForm]=useState(false);const [saving,setSaving]=useState(false);
  const [f,setF]=useState({equipment_name:"",quantity:1,operator_name:"",hours_used:"",notes:"",date:today()});
  async function load(){setLoading(true);try{setEquip(await API.equipment.forProject(projectId)||[]);}catch(e){onErr(e.message);}setLoading(false);}
  useEffect(()=>{load();},[projectId]);
  async function save(){if(!f.equipment_name)return;setSaving(true);try{await API.equipment.create({...f,project_id:projectId});await load();setShowForm(false);setF({equipment_name:"",quantity:1,operator_name:"",hours_used:"",notes:"",date:today()});}catch(e){onErr(e.message);}setSaving(false);}
  async function remove(id){try{await API.equipment.remove(id);await load();}catch(e){onErr(e.message);}}
  const todayEquip=equip.filter(e=>e.date===today());const prevEquip=equip.filter(e=>e.date!==today()).slice(0,20);
  return(<div>
    <button onClick={()=>setShowForm(!showForm)} style={{...primBtn,marginBottom:14,borderRadius:14}}>{showForm?"✕ Cancel":"🚜 Log Equipment On Site"}</button>
    {showForm&&<div style={{...cardS,marginBottom:14,borderLeft:`3px solid ${T.yellow}`}}>
      <div style={{marginBottom:10}}><label style={lbl}>Equipment</label><select value={f.equipment_name} onChange={e=>setF(x=>({...x,equipment_name:e.target.value}))} style={inp}><option value="">— Select —</option>{EQUIP_LIST.filter(e=>!e.section).map(e=><option key={e.name} value={e.name}>{e.name}</option>)}</select></div>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:10}}><div><label style={lbl}>Quantity</label><input type="number" min="1" value={f.quantity} onChange={e=>setF(x=>({...x,quantity:e.target.value}))} style={inp}/></div><div><label style={lbl}>Hours Used</label><input type="number" min="0" step="0.5" placeholder="0" value={f.hours_used} onChange={e=>setF(x=>({...x,hours_used:e.target.value}))} style={inp}/></div></div>
      <div style={{marginBottom:10}}><label style={lbl}>Operator</label><select value={f.operator_name} onChange={e=>setF(x=>({...x,operator_name:e.target.value}))} style={inp}><option value="">— Optional —</option>{NAMES.map(n=><option key={n}>{n}</option>)}</select></div>
      <div style={{marginBottom:10}}><label style={lbl}>Date</label><input type="date" value={f.date} onChange={e=>setF(x=>({...x,date:e.target.value}))} style={inp}/></div>
      <div style={{marginBottom:10}}><label style={lbl}>Notes</label><input type="text" placeholder="Condition, issues…" value={f.notes} onChange={e=>setF(x=>({...x,notes:e.target.value}))} style={inp}/></div>
      <button onClick={save} style={{...primBtn,background:T.yellow,color:"#09090B",borderRadius:12}}>{saving?"Saving…":"Save Entry"}</button>
    </div>}
    {loading&&<Spinner/>}
    {!loading&&<>{todayEquip.length>0&&<div style={{fontSize:12,fontWeight:700,color:T.muted,textTransform:"uppercase",letterSpacing:"1px",marginBottom:10}}>On Site Today</div>}{todayEquip.map(e=><div key={e.id} style={{...cardS,marginBottom:8,borderLeft:`3px solid ${T.yellow}`,display:"flex",justifyContent:"space-between",alignItems:"center"}}><div><div style={{fontSize:14,fontWeight:700}}>{e.equipment_name}</div><div style={{fontSize:11,color:T.muted,marginTop:3}}>{fmtShort(e.date)} · Qty {e.quantity||1}{e.operator_name?" · "+e.operator_name:""}{e.hours_used?" · "+e.hours_used+"h":""}</div></div><button onClick={()=>remove(e.id)} style={{background:"none",border:"none",color:T.muted,cursor:"pointer",fontSize:16,padding:0,marginLeft:12}}>🗑</button></div>)}
    {prevEquip.length>0&&<div style={{fontSize:12,fontWeight:700,color:T.muted,textTransform:"uppercase",letterSpacing:"1px",margin:"14px 0 10px"}}>Previous</div>}{prevEquip.map(e=><div key={e.id} style={{...cardS,marginBottom:8,borderLeft:`3px solid ${T.yellow}40`,display:"flex",justifyContent:"space-between",alignItems:"center"}}><div><div style={{fontSize:13,fontWeight:700}}>{e.equipment_name}</div><div style={{fontSize:11,color:T.muted}}>{fmtShort(e.date)} · Qty {e.quantity||1}{e.operator_name?" · "+e.operator_name:""}</div></div><button onClick={()=>remove(e.id)} style={{background:"none",border:"none",color:T.muted,cursor:"pointer",fontSize:16,padding:0,marginLeft:12}}>🗑</button></div>)}
    {equip.length===0&&!showForm&&<div style={{textAlign:"center",padding:"40px 0",color:T.muted}}><div style={{fontSize:32,marginBottom:8}}>🚜</div><div>No equipment logged.</div></div>}</>}
  </div>);
}

/* ── SUBS TAB ───────────────────────────────────────────────── */
function SubsTab({projectId,user,onErr}){
  const [subs,setSubs]=useState([]);const [loading,setLoading]=useState(true);const [showForm,setShowForm]=useState(false);const [saving,setSaving]=useState(false);
  const [f,setF]=useState({date:today(),company_name:"",trade:"",contact_name:"",contact_phone:"",workers_count:1,hours_worked:"",work_description:""});
  const set=(k,v)=>setF(x=>({...x,[k]:v}));
  async function load(){setLoading(true);try{setSubs(await API.subs.forProject(projectId)||[]);}catch(e){onErr(e.message);}setLoading(false);}
  useEffect(()=>{load();},[projectId]);
  async function save(){if(!f.company_name)return;setSaving(true);try{await API.subs.create({...f,project_id:projectId,created_by:user.name});await load();setShowForm(false);setF({date:today(),company_name:"",trade:"",contact_name:"",contact_phone:"",workers_count:1,hours_worked:"",work_description:""});}catch(e){onErr(e.message);}setSaving(false);}
  async function remove(id){try{await API.subs.remove(id);await load();}catch(e){onErr(e.message);}}
  const trades=["Electrical","Mechanical","Civil","Welding","Coating","Survey","Inspection","HDD","Boring","Concrete","Other"];
  return(<div>
    <button onClick={()=>setShowForm(!showForm)} style={{...primBtn,marginBottom:14,borderRadius:14}}>{showForm?"✕ Cancel":"🏢 Log Subcontractor"}</button>
    {showForm&&<div style={{...cardS,marginBottom:14,borderLeft:`3px solid ${T.purple}`}}>
      <div style={{marginBottom:10}}><label style={lbl}>Company Name *</label><input type="text" placeholder="Sub company name" value={f.company_name} onChange={e=>set("company_name",e.target.value)} style={inp}/></div>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:10}}><div><label style={lbl}>Trade</label><select value={f.trade} onChange={e=>set("trade",e.target.value)} style={inp}><option value="">— Select —</option>{trades.map(t=><option key={t}>{t}</option>)}</select></div><div><label style={lbl}>Date</label><input type="date" value={f.date} onChange={e=>set("date",e.target.value)} style={inp}/></div><div><label style={lbl}>Contact</label><input type="text" placeholder="Foreman" value={f.contact_name} onChange={e=>set("contact_name",e.target.value)} style={inp}/></div><div><label style={lbl}>Phone</label><input type="tel" placeholder="555-555-5555" value={f.contact_phone} onChange={e=>set("contact_phone",e.target.value)} style={inp}/></div><div><label style={lbl}>Workers</label><input type="number" min="0" value={f.workers_count} onChange={e=>set("workers_count",e.target.value)} style={inp}/></div><div><label style={lbl}>Hours</label><input type="number" min="0" step="0.5" placeholder="0" value={f.hours_worked} onChange={e=>set("hours_worked",e.target.value)} style={inp}/></div></div>
      <div style={{marginBottom:10}}><label style={lbl}>Work Description</label><textarea placeholder="What work was performed?" value={f.work_description} onChange={e=>set("work_description",e.target.value)} rows={3} style={{...inp,resize:"vertical"}}/></div>
      <button onClick={save} style={{...primBtn,background:T.purple,color:"#fff",borderRadius:12}}>{saving?"Saving…":"Save Sub Entry"}</button>
    </div>}
    {loading&&<Spinner/>}
    {!loading&&subs.map(s=>(<div key={s.id} style={{...cardS,marginBottom:10,borderLeft:`3px solid ${T.purple}`}}><div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start"}}><div style={{flex:1}}><div style={{fontSize:15,fontWeight:800}}>{s.company_name}</div><div style={{display:"flex",gap:6,marginTop:5,flexWrap:"wrap"}}>{s.trade&&<span style={pill(T.purple)}>{s.trade}</span>}<span style={pill(T.muted)}>{fmtShort(s.date)}</span>{s.workers_count>0&&<span style={pill(T.blue)}>👷 {s.workers_count}</span>}{s.hours_worked>0&&<span style={pill(T.green)}>{s.hours_worked}h</span>}</div>{s.contact_name&&<div style={{fontSize:12,color:T.sub,marginTop:6}}>📞 {s.contact_name}{s.contact_phone?" · "+s.contact_phone:""}</div>}{s.work_description&&<div style={{fontSize:12,color:T.sub,marginTop:4,lineHeight:1.5}}>{s.work_description}</div>}</div><button onClick={()=>remove(s.id)} style={{background:"none",border:"none",color:T.muted,cursor:"pointer",fontSize:16,padding:"0 0 0 10px"}}>🗑</button></div></div>))}
    {!loading&&subs.length===0&&!showForm&&<div style={{textAlign:"center",padding:"40px 0",color:T.muted}}><div style={{fontSize:32,marginBottom:8}}>🏢</div><div>No subcontractors logged.</div></div>}
  </div>);
}

/* ── SAFETY TAB ─────────────────────────────────────────────── */
function SafetyTab({projectId,safety,user,onRefresh,onErr}){
  const [showForm,setShowForm]=useState(false);const [saving,setSaving]=useState(false);const [type,setType]=useState("toolbox");
  const [f,setF]=useState({date:today(),topic:"",notes:"",severity:"low"});
  const TC={toolbox:T.blue,observation:T.yellow,incident:T.red,nearmiss:T.orange,jsa:T.purple};
  const TL={toolbox:"🛠 Toolbox Talk",observation:"👁 Observation",incident:"🚨 Incident",nearmiss:"⚠️ Near Miss",jsa:"📋 JSA"};
  async function save(){if(!f.topic.trim())return;setSaving(true);try{await API.safety.create({...f,type,project_id:projectId,created_by:user.name});await onRefresh();setShowForm(false);setF({date:today(),topic:"",notes:"",severity:"low"});}catch(e){onErr(e.message);}setSaving(false);}
  async function del(id){try{await API.safety.remove(id);await onRefresh();}catch(e){onErr(e.message);}}
  return(<div>
    <button onClick={()=>setShowForm(!showForm)} style={{...primBtn,marginBottom:14,borderRadius:14}}>{showForm?"✕ Cancel":"⛑️ Log Safety Entry"}</button>
    {showForm&&<div style={{...cardS,marginBottom:14,borderLeft:`3px solid ${T.yellow}`}}>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8,marginBottom:12}}>{Object.entries(TL).map(([k,v])=>(<button key={k} onClick={()=>setType(k)} style={{padding:"10px",borderRadius:10,border:`2px solid ${type===k?TC[k]:T.border}`,background:type===k?TC[k]+"20":T.surface,color:type===k?TC[k]:T.sub,fontSize:12,fontWeight:700,cursor:"pointer",fontFamily:"inherit"}}>{v}</button>))}</div>
      <div style={{marginBottom:10}}><label style={lbl}>Date</label><input type="date" value={f.date} onChange={e=>setF(x=>({...x,date:e.target.value}))} style={inp}/></div>
      <div style={{marginBottom:10}}><label style={lbl}>{type==="toolbox"?"Topic":type==="jsa"?"Job / Task Name":"Description"}</label><input type="text" placeholder="Describe…" value={f.topic} onChange={e=>setF(x=>({...x,topic:e.target.value}))} style={inp}/></div>
      <div style={{marginBottom:10}}><label style={lbl}>Notes / Corrective Action</label><textarea rows={3} placeholder="Additional details…" value={f.notes} onChange={e=>setF(x=>({...x,notes:e.target.value}))} style={{...inp,resize:"vertical"}}/></div>
      {(type==="incident"||type==="nearmiss")&&<div style={{marginBottom:10}}><label style={lbl}>Severity</label><select value={f.severity} onChange={e=>setF(x=>({...x,severity:e.target.value}))} style={inp}><option value="low">Low – First Aid</option><option value="medium">Medium – Recordable</option><option value="high">High – Lost Time</option></select></div>}
      <button onClick={save} style={{...primBtn,background:T.yellow,color:"#09090B",borderRadius:12}}>{saving?"Saving…":"Save Entry"}</button>
    </div>}
    {safety.length===0&&!showForm&&<div style={{textAlign:"center",padding:"40px 0",color:T.muted}}><div style={{fontSize:32,marginBottom:8}}>⛑️</div><div>No safety entries yet.</div></div>}
    {[...safety].map(s=>(<div key={s.id} style={{...cardS,marginBottom:9,borderLeft:`3px solid ${TC[s.type]||T.border}`}}><div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start"}}><div style={{flex:1}}><span style={{...pill(TC[s.type]||T.muted),marginBottom:6,display:"inline-flex"}}>{TL[s.type]||s.type}</span><div style={{fontSize:14,fontWeight:700,marginTop:4}}>{s.topic}</div>{s.notes&&<div style={{fontSize:12,color:T.sub,marginTop:4,lineHeight:1.5}}>{s.notes}</div>}<div style={{fontSize:11,color:T.muted,marginTop:6}}>{fmtDate(s.date)} · {s.created_by}</div></div><div style={{display:"flex",flexDirection:"column",alignItems:"flex-end",gap:6}}>{(s.type==="incident"||s.type==="nearmiss")&&<span style={pill(s.severity==="high"?T.red:s.severity==="medium"?T.yellow:T.green)}>{(s.severity||"low").toUpperCase()}</span>}<button onClick={()=>del(s.id)} style={{background:"none",border:"none",color:T.muted,cursor:"pointer",fontSize:16,padding:0}}>🗑</button></div></div></div>))}
  </div>);
}

/* ── DOCS TAB ───────────────────────────────────────────────── */
function DocsTab({projectId,user,onErr}){
  const [docs,setDocs]=useState([]);const [loading,setLoading]=useState(true);const [showForm,setShowForm]=useState(false);const [saving,setSaving]=useState(false);
  const [f,setF]=useState({name:"",doc_type:"Drawing",file_url:"",notes:""});
  const docTypes=["Drawing","Specification","Manual","Permit","Contract","As-Built","ITP","Procedure","Safety Plan","Other"];
  const docIcons={Drawing:"📐",Specification:"📄",Manual:"📗",Permit:"🗂️",Contract:"📝","As-Built":"🗺️",ITP:"✅",Procedure:"📋","Safety Plan":"⛑️",Other:"📁"};
  async function load(){setLoading(true);try{setDocs(await API.docs.forProject(projectId)||[]);}catch(e){onErr(e.message);}setLoading(false);}
  useEffect(()=>{load();},[projectId]);
  async function save(){if(!f.name.trim())return;setSaving(true);try{await API.docs.create({...f,project_id:projectId,uploaded_by:user.name});await load();setShowForm(false);setF({name:"",doc_type:"Drawing",file_url:"",notes:""});}catch(e){onErr(e.message);}setSaving(false);}
  async function remove(id){try{await API.docs.remove(id);await load();}catch(e){onErr(e.message);}}
  const grouped={};docs.forEach(d=>{const t=d.doc_type||"Other";if(!grouped[t])grouped[t]=[];grouped[t].push(d);});
  return(<div>
    <button onClick={()=>setShowForm(!showForm)} style={{...primBtn,marginBottom:14,borderRadius:14}}>{showForm?"✕ Cancel":"📁 + Add Document"}</button>
    {showForm&&<div style={{...cardS,marginBottom:14,borderLeft:`3px solid ${T.teal}`}}>
      <div style={{marginBottom:10}}><label style={lbl}>Document Name *</label><input type="text" placeholder="e.g. P&ID Drawing Rev 3" value={f.name} onChange={e=>setF(x=>({...x,name:e.target.value}))} style={inp}/></div>
      <div style={{marginBottom:10}}><label style={lbl}>Type</label><select value={f.doc_type} onChange={e=>setF(x=>({...x,doc_type:e.target.value}))} style={inp}>{docTypes.map(t=><option key={t}>{t}</option>)}</select></div>
      <div style={{marginBottom:10}}><label style={lbl}>Link (Google Drive, Dropbox, SharePoint…)</label><input type="url" placeholder="https://drive.google.com/…" value={f.file_url} onChange={e=>setF(x=>({...x,file_url:e.target.value}))} style={inp}/><div style={{fontSize:11,color:T.muted,marginTop:4}}>Paste a shareable link. Store files in Google Drive and share the link here.</div></div>
      <div style={{marginBottom:10}}><label style={lbl}>Notes / Revision</label><input type="text" placeholder="e.g. Rev 3 – Approved" value={f.notes} onChange={e=>setF(x=>({...x,notes:e.target.value}))} style={inp}/></div>
      <button onClick={save} style={{...primBtn,background:T.teal,color:"#09090B",borderRadius:12}}>{saving?"Saving…":"Save Document"}</button>
    </div>}
    {loading&&<Spinner/>}
    {!loading&&docs.length===0&&!showForm&&<div style={{textAlign:"center",padding:"40px 0",color:T.muted}}><div style={{fontSize:32,marginBottom:8}}>📁</div><div style={{marginBottom:4}}>No documents yet.</div><div style={{fontSize:12}}>Add links to drawings, specs, manuals, permits.</div></div>}
    {!loading&&Object.entries(grouped).map(([type,typeDocs])=>(<div key={type} style={{marginBottom:16}}><div style={{fontSize:11,fontWeight:700,color:T.muted,textTransform:"uppercase",letterSpacing:"1px",marginBottom:8}}>{docIcons[type]||"📁"} {type} ({typeDocs.length})</div>{typeDocs.map(doc=>(<div key={doc.id} style={{...cardS,marginBottom:8,borderLeft:`3px solid ${T.teal}`}}><div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start"}}><div style={{flex:1,minWidth:0}}><div style={{fontSize:14,fontWeight:700,marginBottom:2}}>{doc.name}</div>{doc.notes&&<div style={{fontSize:12,color:T.sub,marginBottom:4}}>{doc.notes}</div>}<div style={{fontSize:11,color:T.muted}}>{doc.uploaded_by}</div></div><div style={{display:"flex",gap:8,flexShrink:0,marginLeft:10}}>{doc.file_url&&<button onClick={()=>window.open(doc.file_url,"_blank")} style={{background:T.teal+"20",border:`1px solid ${T.teal}40`,borderRadius:8,padding:"6px 12px",color:T.teal,fontSize:12,fontWeight:700,cursor:"pointer",fontFamily:"inherit"}}>Open →</button>}<button onClick={()=>remove(doc.id)} style={{background:"none",border:"none",color:T.muted,cursor:"pointer",fontSize:16,padding:0}}>🗑</button></div></div></div>))}</div>))}
  </div>);
}

/* ── SCHEDULE TAB ───────────────────────────────────────────── */
function ScheduleTab({projectId,user,onErr}){
  const [milestones,setMilestones]=useState([]);const [loading,setLoading]=useState(true);const [showForm,setShowForm]=useState(false);const [saving,setSaving]=useState(false);
  const [f,setF]=useState({title:"",description:"",target_date:"",status:"pending"});
  const set=(k,v)=>setF(x=>({...x,[k]:v}));
  async function load(){setLoading(true);try{setMilestones(await API.milestones.forProject(projectId)||[]);}catch(e){onErr(e.message);}setLoading(false);}
  useEffect(()=>{load();},[projectId]);
  async function save(){if(!f.title.trim())return;setSaving(true);try{await API.milestones.create({...f,project_id:projectId,sort_order:milestones.length});await load();setShowForm(false);setF({title:"",description:"",target_date:"",status:"pending"});}catch(e){onErr(e.message);}setSaving(false);}
  async function toggleStatus(m){const next={pending:"in_progress",in_progress:"completed",completed:"pending"};const completed_date=next[m.status]==="completed"?today():null;try{const[u]=await API.milestones.update(m.id,{status:next[m.status],completed_date});setMilestones(ms=>ms.map(x=>x.id===m.id?u:x));}catch(e){onErr(e.message);}}
  async function del(id){try{await API.milestones.remove(id);await load();}catch(e){onErr(e.message);}}
  const statusColor={pending:T.muted,in_progress:T.orange,completed:T.green,delayed:T.red};
  const statusIcon={pending:"○",in_progress:"◐",completed:"●",delayed:"⚠️"};
  const completed=milestones.filter(m=>m.status==="completed").length;const total=milestones.length;
  return(<div>
    <button onClick={()=>setShowForm(!showForm)} style={{...primBtn,marginBottom:14,borderRadius:14}}>{showForm?"✕ Cancel":"📅 + Add Milestone"}</button>
    {showForm&&<div style={{...cardS,marginBottom:14,borderLeft:`3px solid ${T.blue}`}}>
      <div style={{marginBottom:10}}><label style={lbl}>Milestone Title *</label><input type="text" placeholder="e.g. HDD Bore Complete" value={f.title} onChange={e=>set("title",e.target.value)} style={inp}/></div>
      <div style={{marginBottom:10}}><label style={lbl}>Description</label><input type="text" placeholder="Optional details…" value={f.description} onChange={e=>set("description",e.target.value)} style={inp}/></div>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:10}}><div><label style={lbl}>Target Date</label><input type="date" value={f.target_date} onChange={e=>set("target_date",e.target.value)} style={inp}/></div><div><label style={lbl}>Status</label><select value={f.status} onChange={e=>set("status",e.target.value)} style={inp}><option value="pending">Pending</option><option value="in_progress">In Progress</option><option value="completed">Completed</option><option value="delayed">Delayed</option></select></div></div>
      <button onClick={save} style={{...primBtn,background:T.blue,borderRadius:12}}>{saving?"Saving…":"Save Milestone"}</button>
    </div>}
    {loading&&<Spinner/>}
    {!loading&&total>0&&<div style={{...cardS,marginBottom:14}}><div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:8}}><div style={{fontSize:13,fontWeight:700}}>Progress</div><div style={{fontSize:13,fontWeight:700,color:T.green}}>{completed}/{total} complete</div></div><div style={{height:8,background:T.border,borderRadius:4,overflow:"hidden"}}><div style={{height:"100%",background:`linear-gradient(90deg,${T.orange},${T.green})`,borderRadius:4,width:`${(completed/total)*100}%`,transition:"width 0.4s"}}/></div></div>}
    {!loading&&milestones.length===0&&!showForm&&<div style={{textAlign:"center",padding:"40px 0",color:T.muted}}><div style={{fontSize:32,marginBottom:8}}>📅</div><div>No milestones yet.</div></div>}
    {!loading&&milestones.map(m=>{const du=daysUntil(m.target_date);const overdue=du!==null&&du<0&&m.status!=="completed";const dueSoon=du!==null&&du>=0&&du<=7&&m.status!=="completed";return(<div key={m.id} style={{...cardS,marginBottom:10,borderLeft:`3px solid ${statusColor[m.status]||T.border}`,opacity:m.status==="completed"?0.7:1}}><div style={{display:"flex",alignItems:"flex-start",gap:12}}><button onClick={()=>toggleStatus(m)} style={{width:32,height:32,borderRadius:"50%",border:`2px solid ${statusColor[m.status]||T.border}`,background:m.status==="completed"?T.green:T.surface,color:m.status==="completed"?"#09090B":statusColor[m.status]||T.muted,cursor:"pointer",display:"flex",alignItems:"center",justifyContent:"center",fontSize:14,fontWeight:800,flexShrink:0,marginTop:2}}>{statusIcon[m.status]||"○"}</button><div style={{flex:1,minWidth:0}}><div style={{fontSize:15,fontWeight:700,textDecoration:m.status==="completed"?"line-through":"none",color:m.status==="completed"?T.muted:T.text}}>{m.title}</div>{m.description&&<div style={{fontSize:12,color:T.sub,marginTop:2}}>{m.description}</div>}<div style={{display:"flex",gap:8,marginTop:6,flexWrap:"wrap"}}><span style={pill(statusColor[m.status]||T.muted)}>{m.status.replace("_"," ").toUpperCase()}</span>{m.target_date&&<span style={pill(overdue?T.red:dueSoon?T.yellow:T.muted)}>{overdue?`${Math.abs(du)}d overdue`:dueSoon?`Due in ${du}d`:`Target: ${fmtDate(m.target_date)}`}</span>}{m.completed_date&&<span style={pill(T.green)}>Done: {fmtDate(m.completed_date)}</span>}</div></div><button onClick={()=>del(m.id)} style={{background:"none",border:"none",color:T.muted,cursor:"pointer",fontSize:16,padding:0,flexShrink:0}}>🗑</button></div></div>);})}
  </div>);
}

/* ── PHOTOS TAB ─────────────────────────────────────────────── */
function PhotosTab({projectId,photos,onRefresh,onErr}){
  const [caption,setCaption]=useState("");const [saving,setSaving]=useState(false);const [lb,setLb]=useState(null);
  const fileRef=useRef(null);
  async function handleFiles(files){setSaving(true);try{for(const f of files){if(!f.type.startsWith("image/"))continue;const src=await compressImg(f,1100,0.72);await API.photos.create({project_id:projectId,src,caption,date:today()});}await onRefresh();setCaption("");}catch(e){onErr(e.message);}setSaving(false);}
  async function del(id){try{await API.photos.remove(id);await onRefresh();}catch(e){onErr(e.message);}}
  return(<div>
    <Lightbox src={lb} onClose={()=>setLb(null)}/>
    <div style={{...cardS,marginBottom:14,borderStyle:"dashed",borderColor:T.orange+"44"}}><div style={{marginBottom:10}}><label style={lbl}>Caption (optional)</label><input type="text" placeholder="What is this photo of?" value={caption} onChange={e=>setCaption(e.target.value)} style={inp}/></div><button onClick={()=>fileRef.current?.click()} style={{...primBtn}}>{saving?"Uploading…":"📷 Add Site Photos"}</button><input ref={fileRef} type="file" accept="image/*" capture="environment" multiple style={{display:"none"}} onChange={e=>{handleFiles(Array.from(e.target.files));e.target.value="";}} /></div>
    {photos.length===0&&<div style={{textAlign:"center",padding:"40px 0",color:T.muted}}><div style={{fontSize:32,marginBottom:8}}>📷</div><div>No photos yet.</div></div>}
    <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}>{photos.map(p=>(<div key={p.id} style={{position:"relative",borderRadius:12,overflow:"hidden",aspectRatio:"4/3",background:T.card}}><img src={p.src} alt={p.caption} onClick={()=>setLb(p.src)} style={{width:"100%",height:"100%",objectFit:"cover",cursor:"pointer",display:"block"}}/><button onClick={()=>del(p.id)} style={{position:"absolute",top:6,right:6,background:"rgba(0,0,0,0.75)",border:"none",color:"#fff",borderRadius:"50%",width:24,height:24,fontSize:12,cursor:"pointer"}}>×</button>{p.caption&&<div style={{position:"absolute",bottom:0,left:0,right:0,background:"rgba(0,0,0,0.65)",color:"#fff",fontSize:11,padding:"4px 8px"}}>{p.caption}</div>}<div style={{position:"absolute",top:6,left:6,background:"rgba(0,0,0,0.6)",borderRadius:6,padding:"2px 6px",fontSize:9,color:"#fff"}}>{fmtShort(p.date)}</div></div>))}</div>
  </div>);
}

/* ── WEATHER TAB ────────────────────────────────────────────── */
function WeatherTab({projectId,project,weather,onRefresh,onErr}){
  const [fetching,setFetching]=useState(false);const [liveWeather,setLiveWeather]=useState(null);const [manualNote,setManualNote]=useState("");const [saving,setSaving]=useState(false);
  async function autoFetch(){if(!project.location){onErr("Add a location to this job (Info tab).");return;}setFetching(true);setLiveWeather(null);try{setLiveWeather(await fetchWeather(project.location));}catch(e){onErr(e.message);}setFetching(false);}
  async function logWeather(){if(!liveWeather)return;setSaving(true);const c=liveWeather.current;const[desc]=WMO[c.weathercode]||["Unknown"];try{await API.weather.upsert({project_id:projectId,date:today(),temp_high:liveWeather.daily?.temperature_2m_max?.[0]||c.temperature_2m,temp_low:liveWeather.daily?.temperature_2m_min?.[0]||c.temperature_2m,conditions:desc,wind_speed:c.windspeed_10m,precipitation:liveWeather.daily?.precipitation_sum?.[0]||0,notes:manualNote});await onRefresh();setLiveWeather(null);setManualNote("");}catch(e){onErr(e.message);}setSaving(false);}
  async function del(id){try{await API.weather.remove(id);await onRefresh();}catch(e){onErr(e.message);}}
  return(<div>
    <button onClick={autoFetch} style={{...primBtn,marginBottom:14,borderRadius:14,opacity:fetching?0.6:1}}>{fetching?"🌐 Fetching…":"🌤️ Auto-Fetch Today's Weather"}</button>
    {!project.location&&<div style={{...cardS,marginBottom:14,background:T.yellowLow,border:`1px solid ${T.yellow}40`}}><div style={{fontSize:13,color:T.yellow}}>⚠️ Add a location to this job (Info tab) to auto-fetch weather.</div></div>}
    {liveWeather&&(()=>{const c=liveWeather.current;const[desc,icon]=WMO[c.weathercode]||["Unknown","🌡️"];return(<div style={{...cardS,marginBottom:14,borderLeft:`3px solid ${T.blue}`}}><div style={{fontSize:11,color:T.blue,fontWeight:700,textTransform:"uppercase",letterSpacing:"1px",marginBottom:10}}>Live · {liveWeather.locationName}</div><div style={{display:"flex",alignItems:"center",gap:12,marginBottom:12}}><span style={{fontSize:44}}>{icon}</span><div><div style={{fontSize:28,fontWeight:900,letterSpacing:"-1px"}}>{Math.round(c.temperature_2m)}°F</div><div style={{fontSize:14,color:T.sub}}>{desc}</div><div style={{fontSize:12,color:T.muted}}>Feels {Math.round(c.apparent_temperature)}°F</div></div></div><div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:8,marginBottom:12}}>{[["Wind",Math.round(c.windspeed_10m)+" mph"],["High",Math.round(liveWeather.daily?.temperature_2m_max?.[0]||c.temperature_2m)+"°F"],["Precip",(liveWeather.daily?.precipitation_sum?.[0]||0).toFixed(2)+"in"]].map(([l,v])=>(<div key={l} style={{background:T.surface,borderRadius:10,padding:"8px",textAlign:"center"}}><div style={{fontSize:13,fontWeight:700}}>{v}</div><div style={{fontSize:10,color:T.muted}}>{l}</div></div>))}</div><div style={{marginBottom:10}}><label style={lbl}>Field Notes</label><input type="text" placeholder="Work impacted by weather?" value={manualNote} onChange={e=>setManualNote(e.target.value)} style={inp}/></div><button onClick={logWeather} style={{...primBtn,background:T.blue,borderRadius:12}}>{saving?"Saving…":"💾 Log This Weather"}</button></div>);})()}
    {weather.length>0&&<div style={{fontSize:12,fontWeight:700,color:T.muted,textTransform:"uppercase",letterSpacing:"1px",marginBottom:10}}>History</div>}
    {weather.map(w=>{const entry=Object.entries(WMO).find(([,v])=>v[0]===w.conditions);const icon=entry?entry[1][1]:"🌡️";return(<div key={w.id} style={{...cardS,marginBottom:8,display:"flex",justifyContent:"space-between",alignItems:"center"}}><div><div style={{display:"flex",alignItems:"center",gap:8}}><span style={{fontSize:20}}>{icon}</span><div><div style={{fontSize:14,fontWeight:700}}>{w.conditions||"Logged"}</div><div style={{fontSize:11,color:T.muted}}>{fmtShort(w.date)}{w.wind_speed?" · "+Math.round(w.wind_speed)+"mph":""}{w.precipitation>0?" · "+w.precipitation+"in":""}</div></div></div>{w.notes&&<div style={{fontSize:12,color:T.sub,marginTop:4}}>{w.notes}</div>}</div><div style={{display:"flex",alignItems:"center",gap:8}}>{w.temp_high&&<div style={{textAlign:"right"}}><div style={{fontSize:15,fontWeight:800,color:T.orange}}>{Math.round(w.temp_high)}°</div>{w.temp_low&&<div style={{fontSize:10,color:T.muted}}>{Math.round(w.temp_low)}° lo</div>}</div>}<button onClick={()=>del(w.id)} style={{background:"none",border:"none",color:T.muted,cursor:"pointer",fontSize:16,padding:0}}>🗑</button></div></div>);})}
    {weather.length===0&&!liveWeather&&<div style={{textAlign:"center",padding:"40px 0",color:T.muted}}><div style={{fontSize:32,marginBottom:8}}>🌤️</div><div>No weather logs yet.</div></div>}
  </div>);
}

/* ── INFO TAB ───────────────────────────────────────────────── */
function InfoTab({project,user,onEdit,onArchive}){
  return(<div>
    <div style={cardS}>{[["Division",project.division],["Client",project.client],["Location",project.location],["AFE No.",project.afe],["Work Order",project.work_order],["Start Date",fmtDate(project.start_date)],["Status",project.status],["Created By",project.created_by]].map(([l,v])=>v?(<div key={l} style={{display:"flex",justifyContent:"space-between",padding:"11px 0",borderBottom:`1px solid ${T.border}`}}><span style={{fontSize:13,color:T.muted}}>{l}</span><span style={{fontSize:13,fontWeight:600}}>{v}</span></div>):null)}</div>
    {project.notes&&<div style={{...cardS,marginTop:12}}><div style={{fontSize:11,color:T.muted,textTransform:"uppercase",letterSpacing:"1px",marginBottom:8}}>Notes</div><div style={{fontSize:14,color:T.sub,lineHeight:1.6}}>{project.notes}</div></div>}
    {can(user,"edit_job")&&<div style={{marginTop:16,display:"flex",flexDirection:"column",gap:10}}><button onClick={onEdit} style={{...ghostBtn,width:"100%",textAlign:"center"}}>✏️ Edit Job</button><button onClick={onArchive} style={{...ghostBtn,width:"100%",textAlign:"center",color:T.muted}}>{project.status==="active"?"📦 Archive Job":"♻️ Restore Job"}</button></div>}
  </div>);
}

/* ── USER MANAGEMENT SCREEN (Admin only) ───────────────────── */
function UserManagementScreen({onBack,currentUser}){
  const [profiles,setProfiles]=useState([]);const [loading,setLoading]=useState(true);const [err,setErr]=useState("");
  const [mode,setMode]=useState("list");// list | edit
  const [active,setActive]=useState(null);const [saving,setSaving]=useState(false);
  const blank={name:"",role:"crew",division:null,pin:"",active:true};
  const [f,setF]=useState({...blank});
  const set=(k,v)=>setF(x=>({...x,[k]:v}));

  async function load(){setLoading(true);try{setProfiles(await API.userProfiles.list()||[]);}catch(e){setErr(e.message);}setLoading(false);}
  useEffect(()=>{load();},[]);

  async function save(){
    if(!f.name.trim())return;setSaving(true);
    try{
      if(active){await API.userProfiles.update(active.id,{role:f.role,division:f.division,pin:f.pin,active:f.active});}
      else{await API.userProfiles.upsert({name:f.name,role:f.role,division:f.division||null,pin:f.pin||null,active:true});}
      await load();setMode("list");setActive(null);setF({...blank});
    }catch(e){setErr(e.message);}setSaving(false);
  }

  async function remove(id){if(!window.confirm("Remove this user profile?"))return;try{await API.userProfiles.remove(id);await load();}catch(e){setErr(e.message);}}

  // Build map of who has a profile
  const profileMap={};profiles.forEach(p=>profileMap[p.name]=p);
  // Everyone in NAMES list + any extra profiles
  const allNames=[...new Set([...NAMES,...profiles.map(p=>p.name)])].sort();

  if(mode==="edit") return(
    <div style={{background:T.bg,minHeight:"100vh",fontFamily:"inherit"}}>
      <TopBar title={active?"Edit User":"Add User"} onBack={()=>{setMode("list");setActive(null);setF({...blank});}}/>
      <div style={{padding:"16px 16px 100px"}}>
        <ErrBanner msg={err} onDismiss={()=>setErr("")}/>
        {!active&&<div style={{marginBottom:14}}><label style={lbl}>Name</label><select value={f.name} onChange={e=>set("name",e.target.value)} style={inp}><option value="">— Select —</option>{NAMES.map(n=><option key={n}>{n}</option>)}</select></div>}
        {active&&<div style={{...cardS,marginBottom:14}}><div style={{fontSize:16,fontWeight:800}}>{active.name}</div><div style={{fontSize:12,color:T.muted}}>Editing permissions</div></div>}

        <div style={{marginBottom:14}}>
          <label style={lbl}>Permission Level</label>
          <div style={{display:"flex",flexDirection:"column",gap:8}}>
            {ROLES.map(role=>{const m=ROLE_META[role];return(<button key={role} onClick={()=>set("role",role)} style={{display:"flex",alignItems:"center",gap:12,padding:"14px",borderRadius:12,border:`2px solid ${f.role===role?m.color:T.border}`,background:f.role===role?m.color+"18":T.surface,cursor:"pointer",fontFamily:"inherit",textAlign:"left"}}>
              <div style={{width:12,height:12,borderRadius:"50%",background:m.color,flexShrink:0}}/>
              <div style={{flex:1}}><div style={{fontSize:14,fontWeight:700,color:f.role===role?m.color:T.text}}>{m.label}</div><div style={{fontSize:12,color:T.muted}}>{m.desc}</div></div>
              {f.role===role&&<div style={{fontSize:16,color:m.color}}>✓</div>}
            </button>);})}
          </div>
        </div>

        <div style={{marginBottom:14}}>
          <label style={lbl}>Assigned Division (optional)</label>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8}}>
            <button onClick={()=>set("division",null)} style={{padding:"12px",borderRadius:12,border:`2px solid ${f.division===null?T.orange:T.border}`,background:f.division===null?T.orangeLow:T.surface,color:f.division===null?T.orange:T.sub,fontWeight:600,fontSize:13,cursor:"pointer",fontFamily:"inherit"}}>All Divisions</button>
            {DIVISIONS.map(div=>{const m=DIV_META[div];return(<button key={div} onClick={()=>set("division",div)} style={{padding:"12px",borderRadius:12,border:`2px solid ${f.division===div?m.color:T.border}`,background:f.division===div?m.color+"18":T.surface,color:f.division===div?m.color:T.sub,fontWeight:600,fontSize:13,cursor:"pointer",fontFamily:"inherit"}}>{m.icon} {div}</button>);})}
          </div>
        </div>

        {(f.role==="pm"||f.role==="admin")&&<div style={{marginBottom:14}}><label style={lbl}>{f.role==="admin"?"Admin PIN":"PM PIN"} (required)</label><input type="text" maxLength={6} placeholder="Set a PIN (numbers)" value={f.pin||""} onChange={e=>set("pin",e.target.value)} style={inp}/><div style={{fontSize:11,color:T.muted,marginTop:4}}>This person will need to enter this PIN to sign in.</div></div>}

        {active&&<div style={{marginBottom:14}}><label style={lbl}>Status</label><div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8}}>{[true,false].map(v=>(<button key={String(v)} onClick={()=>set("active",v)} style={{padding:"12px",borderRadius:12,border:`2px solid ${f.active===v?(v?T.green:T.red):T.border}`,background:f.active===v?(v?T.greenLow:T.redLow):T.surface,color:f.active===v?(v?T.green:T.red):T.sub,fontWeight:700,fontSize:14,cursor:"pointer",fontFamily:"inherit"}}>{v?"✅ Active":"❌ Inactive"}</button>))}</div></div>}

        <button onClick={save} style={{...primBtn,opacity:f.name&&!saving?1:0.5}}>{saving?"Saving…":active?"Save Changes":"Add User"}</button>
      </div>
    </div>
  );

  return(
    <div style={{background:T.bg,minHeight:"100vh",fontFamily:"inherit"}}>
      <div style={{background:T.surface,borderBottom:`1px solid ${T.border}`,padding:"14px 16px",position:"sticky",top:0,zIndex:50}}>
        <button onClick={onBack} style={{background:"none",border:"none",color:T.sub,fontSize:13,cursor:"pointer",marginBottom:8,padding:0,fontFamily:"inherit"}}>← Back</button>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}>
          <div><div style={{fontSize:20,fontWeight:900,letterSpacing:"-0.5px"}}>👤 User Management</div><div style={{fontSize:12,color:T.muted}}>Set permission levels for your crew</div></div>
          <button onClick={()=>{setF({...blank});setMode("edit");}} style={{background:T.orange,color:"#09090B",border:"none",borderRadius:10,padding:"8px 14px",fontSize:13,fontWeight:800,cursor:"pointer",fontFamily:"inherit"}}>+ Add</button>
        </div>
      </div>
      <div style={{padding:"14px 16px 80px"}}>
        <ErrBanner msg={err} onDismiss={()=>setErr("")}/>
        {loading&&<Spinner/>}
        {!loading&&<>
          {/* Users with profiles */}
          <div style={{fontSize:12,fontWeight:700,color:T.muted,textTransform:"uppercase",letterSpacing:"1px",marginBottom:12}}>Configured Users ({profiles.length})</div>
          {profiles.length===0&&<div style={{...cardS,marginBottom:14,background:T.yellowLow,border:`1px solid ${T.yellow}40`}}><div style={{fontSize:13,color:T.yellow}}>⚠️ No user profiles set. All users will sign in as Field Crew until you configure them.</div></div>}
          {profiles.map(p=>{
            const m=ROLE_META[p.role]||ROLE_META.crew;
            return(<div key={p.id} style={{...cardS,marginBottom:8,borderLeft:`3px solid ${m.color}`,opacity:p.active?1:0.5}}>
              <div style={{display:"flex",alignItems:"center",gap:12}}>
                <div style={{width:10,height:10,borderRadius:"50%",background:m.color,flexShrink:0}}/>
                <div style={{flex:1}}>
                  <div style={{fontSize:14,fontWeight:700}}>{p.name}</div>
                  <div style={{display:"flex",gap:6,marginTop:4,flexWrap:"wrap"}}>
                    <span style={pill(m.color)}>{m.label}</span>
                    {p.division&&<span style={pill(DIV_META[p.division]?.color||T.muted)}>{DIV_META[p.division]?.icon} {p.division}</span>}
                    {!p.active&&<span style={pill(T.red)}>INACTIVE</span>}
                  </div>
                </div>
                <div style={{display:"flex",gap:6}}>
                  <button onClick={()=>{setActive(p);setF({name:p.name,role:p.role,division:p.division||null,pin:p.pin||"",active:p.active});setMode("edit");}} style={{background:T.orangeLow,border:`1px solid ${T.orange}40`,borderRadius:8,padding:"6px 12px",color:T.orange,fontSize:12,fontWeight:700,cursor:"pointer",fontFamily:"inherit"}}>Edit</button>
                  {p.name!==currentUser.name&&<button onClick={()=>remove(p.id)} style={{background:"none",border:"none",color:T.muted,cursor:"pointer",fontSize:16,padding:0}}>🗑</button>}
                </div>
              </div>
            </div>);
          })}

          {/* Everyone else defaults to Crew */}
          <div style={{fontSize:12,fontWeight:700,color:T.muted,textTransform:"uppercase",letterSpacing:"1px",margin:"20px 0 10px"}}>Unconfigured (Default: Field Crew)</div>
          {NAMES.filter(n=>!profileMap[n]).slice(0,15).map(n=>(
            <div key={n} style={{...cardS,marginBottom:6,display:"flex",alignItems:"center",justifyContent:"space-between",opacity:0.5}}>
              <div style={{display:"flex",alignItems:"center",gap:10}}><div style={{width:10,height:10,borderRadius:"50%",background:T.green}}/><span style={{fontSize:13}}>{n}</span><span style={pill(T.green)}>Field Crew</span></div>
              <button onClick={()=>{setF({...blank,name:n});setMode("edit");}} style={{background:T.surface,border:`1px solid ${T.border}`,borderRadius:8,padding:"5px 10px",color:T.sub,fontSize:12,cursor:"pointer",fontFamily:"inherit"}}>Set Role</button>
            </div>
          ))}
          {NAMES.filter(n=>!profileMap[n]).length>15&&<div style={{fontSize:12,color:T.muted,textAlign:"center",padding:"8px 0"}}>+ {NAMES.filter(n=>!profileMap[n]).length-15} more (tap + Add to configure)</div>}
        </>}
      </div>
    </div>
  );
}

/* ── CREW DIRECTORY SCREEN ──────────────────────────────────── */
function CrewDirectoryScreen({onBack,user}){
  const [members,setMembers]=useState([]);const [loading,setLoading]=useState(true);const [err,setErr]=useState("");
  const [mode,setMode]=useState("list");const [active,setActive]=useState(null);const [saving,setSaving]=useState(false);
  const blank={name:"",classification:"",phone:"",email:"",emergency_contact_name:"",emergency_contact_phone:"",certifications:[],notes:"",active:true};
  const [f,setF]=useState({...blank});const set=(k,v)=>setF(x=>({...x,[k]:v}));
  async function load(){setLoading(true);try{setMembers(await API.crew.list()||[]);}catch(e){setErr(e.message);}setLoading(false);}
  useEffect(()=>{load();},[]);
  async function save(){if(!f.name.trim())return;setSaving(true);try{if(active){await API.crew.update(active.id,f);setActive({...active,...f});}else{await API.crew.create(f);}await load();setMode("list");setActive(null);setF({...blank});}catch(e){setErr(e.message);}setSaving(false);}
  async function remove(id){if(!window.confirm("Remove crew member?"))return;try{await API.crew.remove(id);await load();setMode("list");setActive(null);}catch(e){setErr(e.message);}}
  function addCert(){set("certifications",[...(f.certifications||[]),{id:uid(),name:"",expiry:"",cert_number:""}]);}
  function updateCert(i,k,v){const c=[...(f.certifications||[])];c[i]={...c[i],[k]:v};set("certifications",c);}
  function removeCert(i){set("certifications",(f.certifications||[]).filter((_,j)=>j!==i));}
  const CERT_TYPES=["OSHA 10","OSHA 30","First Aid / CPR","Confined Space Entry","Crane Operator","Welding Certification","Pipeline Operator Qualification","Hydro Test Operator","Excavation Competent Person","H2S Safety","Driver CDL","Other"];
  const active_m=members.filter(m=>m.active);const inactive_m=members.filter(m=>!m.active);

  if(mode==="new"||mode==="edit") return(
    <div style={{background:T.bg,minHeight:"100vh",fontFamily:"inherit"}}>
      <TopBar title={mode==="edit"?"Edit Member":"Add Member"} onBack={()=>{setMode("list");setActive(null);setF({...blank});}}/>
      <div style={{padding:"16px 16px 100px"}}>
        <ErrBanner msg={err} onDismiss={()=>setErr("")}/>
        <div style={{marginBottom:12}}><label style={lbl}>Full Name *</label><select value={f.name} onChange={e=>set("name",e.target.value)} style={inp}><option value="">— Select —</option>{NAMES.map(n=><option key={n}>{n}</option>)}</select></div>
        <div style={{marginBottom:12}}><label style={lbl}>Classification</label><select value={f.classification} onChange={e=>set("classification",e.target.value)} style={inp}><option value="">— Select —</option>{POSITIONS.map(p=><option key={p.name}>{p.name}</option>)}</select></div>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:12}}><div><label style={lbl}>Cell Phone</label><input type="tel" placeholder="555-555-5555" value={f.phone} onChange={e=>set("phone",e.target.value)} style={inp}/></div><div><label style={lbl}>Email</label><input type="email" placeholder="email@example.com" value={f.email} onChange={e=>set("email",e.target.value)} style={inp}/></div></div>
        <div style={{...cardS,marginBottom:12}}><div style={{fontSize:12,fontWeight:700,color:T.red,textTransform:"uppercase",letterSpacing:"1px",marginBottom:10}}>🆘 Emergency Contact</div><div style={{marginBottom:10}}><label style={lbl}>Name</label><input type="text" placeholder="Spouse, parent…" value={f.emergency_contact_name} onChange={e=>set("emergency_contact_name",e.target.value)} style={inp}/></div><div><label style={lbl}>Phone</label><input type="tel" placeholder="555-555-5555" value={f.emergency_contact_phone} onChange={e=>set("emergency_contact_phone",e.target.value)} style={inp}/></div></div>
        <div style={{...cardS,marginBottom:12}}><div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:10}}><div style={{fontSize:12,fontWeight:700,color:T.blue,textTransform:"uppercase",letterSpacing:"1px"}}>🎖️ Certifications</div><button onClick={addCert} style={{background:T.blueLow,border:`1px solid ${T.blue}40`,borderRadius:8,padding:"6px 12px",color:T.blue,fontSize:12,fontWeight:700,cursor:"pointer",fontFamily:"inherit"}}>+ Add</button></div>{(f.certifications||[]).length===0&&<div style={{fontSize:13,color:T.muted,textAlign:"center",padding:"10px 0"}}>No certifications added.</div>}{(f.certifications||[]).map((cert,i)=>(<div key={cert.id} style={{borderTop:`1px solid ${T.border}`,paddingTop:10,marginTop:i>0?10:0}}><div style={{marginBottom:8}}><label style={lbl}>Certification</label><select value={cert.name} onChange={e=>updateCert(i,"name",e.target.value)} style={inp}><option value="">— Select —</option>{CERT_TYPES.map(t=><option key={t}>{t}</option>)}</select></div><div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8,marginBottom:8}}><div><label style={lbl}>Cert Number</label><input type="text" placeholder="Optional" value={cert.cert_number} onChange={e=>updateCert(i,"cert_number",e.target.value)} style={inp}/></div><div><label style={lbl}>Expiry Date</label><input type="date" value={cert.expiry} onChange={e=>updateCert(i,"expiry",e.target.value)} style={inp}/></div></div><button onClick={()=>removeCert(i)} style={{background:"none",border:"none",color:T.red,cursor:"pointer",fontSize:12,fontWeight:700,fontFamily:"inherit"}}>Remove</button></div>))}</div>
        <div style={{marginBottom:20}}><label style={lbl}>Notes</label><textarea placeholder="Skills, notes, restrictions…" value={f.notes} onChange={e=>set("notes",e.target.value)} rows={3} style={{...inp,resize:"vertical"}}/></div>
        <button onClick={save} style={{...primBtn,opacity:f.name&&!saving?1:0.5}}>{saving?"Saving…":mode==="edit"?"Save Changes":"Add Member"}</button>
      </div>
    </div>
  );

  if(mode==="view"&&active) return(
    <div style={{background:T.bg,minHeight:"100vh",padding:16,fontFamily:"inherit"}}>
      <button onClick={()=>{setMode("list");setActive(null);}} style={{...ghostBtn,marginBottom:14}}>← Directory</button>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:14}}><div><div style={{fontSize:22,fontWeight:900,letterSpacing:"-0.5px"}}>{active.name}</div>{active.classification&&<div style={{fontSize:14,color:T.sub,marginTop:2}}>{active.classification}</div>}</div><button onClick={()=>{setF({...blank,...active,certifications:active.certifications||[]});setMode("edit");}} style={{background:T.orangeLow,border:`1px solid ${T.orange}40`,borderRadius:10,padding:"8px 14px",color:T.orange,fontSize:13,fontWeight:700,cursor:"pointer",fontFamily:"inherit"}}>✏️ Edit</button></div>
      {(active.phone||active.email)&&<div style={{...cardS,marginBottom:12}}><div style={{fontSize:12,fontWeight:700,color:T.blue,textTransform:"uppercase",letterSpacing:"1px",marginBottom:10}}>Contact</div>{active.phone&&<div style={{display:"flex",justifyContent:"space-between",padding:"8px 0",borderBottom:`1px solid ${T.border}`}}><span style={{fontSize:13,color:T.muted}}>Cell</span><a href={`tel:${active.phone}`} style={{fontSize:13,fontWeight:600,color:T.blue,textDecoration:"none"}}>{active.phone}</a></div>}{active.email&&<div style={{display:"flex",justifyContent:"space-between",padding:"8px 0"}}><span style={{fontSize:13,color:T.muted}}>Email</span><a href={`mailto:${active.email}`} style={{fontSize:13,fontWeight:600,color:T.blue,textDecoration:"none"}}>{active.email}</a></div>}</div>}
      {(active.emergency_contact_name||active.emergency_contact_phone)&&<div style={{...cardS,marginBottom:12,borderLeft:`3px solid ${T.red}`}}><div style={{fontSize:12,fontWeight:700,color:T.red,textTransform:"uppercase",letterSpacing:"1px",marginBottom:10}}>🆘 Emergency Contact</div>{active.emergency_contact_name&&<div style={{fontSize:14,fontWeight:700,marginBottom:4}}>{active.emergency_contact_name}</div>}{active.emergency_contact_phone&&<a href={`tel:${active.emergency_contact_phone}`} style={{fontSize:14,color:T.red,textDecoration:"none",fontWeight:700}}>📞 {active.emergency_contact_phone}</a>}</div>}
      {(active.certifications||[]).length>0&&<div style={{...cardS,marginBottom:12}}><div style={{fontSize:12,fontWeight:700,color:T.blue,textTransform:"uppercase",letterSpacing:"1px",marginBottom:10}}>🎖️ Certifications</div>{(active.certifications||[]).map((cert,i)=>{const exp=cert.expiry?daysUntil(cert.expiry):null;const expired=exp!==null&&exp<0;const expiring=exp!==null&&exp>=0&&exp<=30;return(<div key={cert.id||i} style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"8px 0",borderBottom:i<active.certifications.length-1?`1px solid ${T.border}`:"none"}}><div><div style={{fontSize:13,fontWeight:600}}>{cert.name}</div>{cert.cert_number&&<div style={{fontSize:11,color:T.muted}}>#{cert.cert_number}</div>}</div>{cert.expiry&&<span style={pill(expired?T.red:expiring?T.yellow:T.green)}>{expired?"EXPIRED":expiring?`Exp ${exp}d`:fmtDate(cert.expiry)}</span>}</div>);})}</div>}
      {active.notes&&<div style={{...cardS,marginBottom:12}}><div style={{fontSize:11,color:T.muted,textTransform:"uppercase",letterSpacing:"1px",marginBottom:6}}>Notes</div><div style={{fontSize:14,color:T.sub,lineHeight:1.6}}>{active.notes}</div></div>}
      {can(user,"crew_directory")&&<button onClick={()=>remove(active.id)} style={{...dangerBtn,marginTop:8}}>Remove from Directory</button>}
    </div>
  );

  return(
    <div style={{background:T.bg,minHeight:"100vh",fontFamily:"inherit"}}>
      <div style={{background:T.surface,borderBottom:`1px solid ${T.border}`,padding:"14px 16px",position:"sticky",top:0,zIndex:50}}>
        <button onClick={onBack} style={{background:"none",border:"none",color:T.sub,fontSize:13,cursor:"pointer",marginBottom:8,padding:0,fontFamily:"inherit"}}>← Back</button>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}><div style={{fontSize:20,fontWeight:900,letterSpacing:"-0.5px"}}>👥 Crew Directory</div><button onClick={()=>{setF({...blank});setMode("new");}} style={{background:T.orange,color:"#09090B",border:"none",borderRadius:10,padding:"8px 14px",fontSize:13,fontWeight:800,cursor:"pointer",fontFamily:"inherit"}}>+ Add</button></div>
      </div>
      <div style={{padding:"14px 16px 80px"}}>
        <ErrBanner msg={err} onDismiss={()=>setErr("")}/>
        {loading&&<Spinner/>}
        {!loading&&<>
          {active_m.length===0&&<div style={{textAlign:"center",padding:"40px 0",color:T.muted}}><div style={{fontSize:36,marginBottom:8}}>👥</div><div>No crew members yet.</div></div>}
          {active_m.map(m=>{const expiredCerts=(m.certifications||[]).filter(c=>c.expiry&&daysUntil(c.expiry)<0);const expiringSoon=(m.certifications||[]).filter(c=>c.expiry&&daysUntil(c.expiry)>=0&&daysUntil(c.expiry)<=30);return(<div key={m.id} onClick={()=>{setActive(m);setMode("view");}} style={{...cardS,marginBottom:10,cursor:"pointer",display:"flex",alignItems:"center",gap:12}}><div style={{width:44,height:44,borderRadius:12,background:T.orangeLow,border:`2px solid ${T.orange}40`,display:"flex",alignItems:"center",justifyContent:"center",fontSize:16,fontWeight:800,color:T.orange,flexShrink:0}}>{m.name.split(" ").map(w=>w[0]).slice(0,2).join("")}</div><div style={{flex:1,minWidth:0}}><div style={{fontSize:15,fontWeight:700}}>{m.name}</div><div style={{fontSize:12,color:T.sub}}>{m.classification||"No classification"}{m.phone?" · "+m.phone:""}</div><div style={{display:"flex",gap:6,marginTop:4,flexWrap:"wrap"}}>{(m.certifications||[]).length>0&&<span style={pill(T.blue)}>{m.certifications.length} certs</span>}{expiredCerts.length>0&&<span style={pill(T.red)}>{expiredCerts.length} expired</span>}{expiringSoon.length>0&&<span style={pill(T.yellow)}>{expiringSoon.length} expiring</span>}</div></div><span style={{fontSize:16,color:T.muted}}>›</span></div>);})}
          {inactive_m.length>0&&<><div style={{fontSize:12,fontWeight:700,color:T.muted,textTransform:"uppercase",letterSpacing:"1px",margin:"20px 0 10px"}}>Inactive</div>{inactive_m.map(m=>(<div key={m.id} onClick={()=>{setActive(m);setMode("view");}} style={{...cardS,marginBottom:8,cursor:"pointer",opacity:0.5,display:"flex",alignItems:"center",gap:12}}><div style={{width:36,height:36,borderRadius:10,background:T.surface,display:"flex",alignItems:"center",justifyContent:"center",fontSize:13,fontWeight:700,color:T.muted,flexShrink:0}}>{m.name.split(" ").map(w=>w[0]).slice(0,2).join("")}</div><div><div style={{fontSize:14,fontWeight:600}}>{m.name}</div><div style={{fontSize:12,color:T.muted}}>{m.classification}</div></div></div>))}</>}
        </>}
      </div>
    </div>
  );
}

/* ── PROJECT DETAIL (ORCHESTRATOR) ─────────────────────────── */
const PTABS=[
  {id:"reports",icon:"📋",label:"Reports",perm:"submit_report"},
  {id:"time",icon:"⏱️",label:"Time",perm:"time_card"},
  {id:"crew",icon:"🚜",label:"Crew",perm:"crew_equip"},
  {id:"subs",icon:"🏢",label:"Subs",perm:"subs"},
  {id:"safety",icon:"⛑️",label:"Safety",perm:"safety"},
  {id:"docs",icon:"📁",label:"Docs",perm:"docs"},
  {id:"schedule",icon:"📅",label:"Schedule",perm:"schedule"},
  {id:"photos",icon:"📷",label:"Photos",perm:"photos"},
  {id:"weather",icon:"🌤️",label:"Weather",perm:"weather"},
  {id:"info",icon:"ℹ️",label:"Info",perm:null},
];

function ProjectDetail({project:initP,user,onBack,onProjectUpdated}){
  const [project,setProject]=useState(initP);
  const [reports,setReports]=useState([]);const [safety,setSafety]=useState([]);const [photos,setPhotos]=useState([]);const [weather,setWeather]=useState([]);
  const [tab,setTab]=useState("reports");const [loading,setLoading]=useState(true);const [err,setErr]=useState("");
  const [screen,setScreen]=useState("detail");const [activeReport,setActiveReport]=useState(null);const [editProject,setEditProject]=useState(false);
  const visibleTabs=PTABS.filter(t=>!t.perm||can(user,t.perm));
  const divMeta=DIV_META[project.division]||{color:T.orange,icon:"🏗️"};

  async function load(silent=false){if(!silent)setLoading(true);try{const[reps,saf,phs,wx]=await Promise.all([API.reports.forProject(project.id),API.safety.forProject(project.id),API.photos.forProject(project.id),API.weather.forProject(project.id)]);setReports(reps||[]);setSafety(saf||[]);setPhotos(phs||[]);setWeather(wx||[]);}catch(e){setErr(e.message);}if(!silent)setLoading(false);}
  useEffect(()=>{load();const firstTab=visibleTabs[0]?.id||"info";setTab(firstTab);},[project.id]);

  async function saveReport(d){try{await API.reports.create({...d,project_id:project.id});await load(true);setScreen("detail");}catch(e){setErr(e.message);}}
  async function deleteReport(id){try{await API.reports.remove(id);setActiveReport(null);await load(true);setScreen("detail");}catch(e){setErr(e.message);}}
  async function approveReport(id){try{await API.reports.update(id,{status:"approved",approved_by:user.name,approved_at:new Date().toISOString()});setActiveReport(r=>({...r,status:"approved"}));await load(true);}catch(e){setErr(e.message);}}
  async function flagReport(id,pm_notes){try{await API.reports.update(id,{status:"flagged",pm_notes});setActiveReport(r=>({...r,status:"flagged",pm_notes}));await notify("report_flagged","Report Flagged",pm_notes,{project_id:project.id,report_id:id});await load(true);}catch(e){setErr(e.message);}}
  async function updateProject(data){try{const[u]=await API.projects.update(project.id,data);setProject(u);onProjectUpdated(u);setEditProject(false);}catch(e){setErr(e.message);}}
  async function archiveProject(){if(!window.confirm(project.status==="active"?"Archive this job?":"Restore?"))return;await updateProject({status:project.status==="active"?"archived":"active"});onBack();}

  const tot=reports.reduce((s,r)=>{const t=reportTotals(r);return{l:s.l+t.labor,e:s.e+t.equip,m:s.m+t.mats,g:s.g+t.grand};},{l:0,e:0,m:0,g:0});

  if(screen==="newReport"&&can(user,"submit_report")) return <DailyReportForm user={user} project={project} onSave={saveReport} onCancel={()=>setScreen("detail")}/>;
  if(screen==="reportDetail"&&activeReport) return <ReportDetail report={activeReport} project={project} user={user} onBack={()=>setScreen("detail")} onDelete={deleteReport} onApprove={approveReport} onFlag={flagReport}/>;
  if(editProject&&can(user,"edit_job")) return <ProjectForm initial={project} onSave={updateProject} onCancel={()=>setEditProject(false)} defaultDivision={project.division}/>;

  return(
    <div style={{background:T.bg,minHeight:"100vh",fontFamily:"inherit"}}>
      <div style={{background:T.surface,borderBottom:`1px solid ${T.border}`,padding:"14px 16px",position:"sticky",top:0,zIndex:50}}>
        <button onClick={onBack} style={{background:"none",border:"none",color:T.sub,fontSize:13,cursor:"pointer",marginBottom:8,padding:0,fontFamily:"inherit"}}>← {project.division}</button>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:12}}>
          <div style={{flex:1,minWidth:0,paddingRight:10}}>
            <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:2}}>
              <span style={{fontSize:16}}>{divMeta.icon}</span>
              <div style={{fontSize:19,fontWeight:900,color:T.text,letterSpacing:"-0.4px",lineHeight:1.2}}>{project.name}</div>
            </div>
            <div style={{fontSize:12,color:T.sub}}>{[project.client,project.location].filter(Boolean).join(" · ")||"No details"}</div>
          </div>
          <span style={pill(project.status==="active"?T.green:T.muted)}>{project.status}</span>
        </div>
        <StatBar items={[{label:"Reports",val:reports.length,color:divMeta.color},{label:"Labor",val:"$"+(tot.l>=1000?(tot.l/1000).toFixed(1)+"k":fmt(tot.l)),color:T.green},{label:"Equip",val:"$"+(tot.e>=1000?(tot.e/1000).toFixed(1)+"k":fmt(tot.e)),color:T.yellow},{label:"Total",val:"$"+(tot.g>=1000?(tot.g/1000).toFixed(1)+"k":fmt(tot.g)),color:T.blue}]}/>
        <div style={{display:"flex",gap:4,marginTop:12,overflowX:"auto",paddingBottom:2,WebkitOverflowScrolling:"touch"}}>
          {visibleTabs.map(t=>(<button key={t.id} onClick={()=>setTab(t.id)} style={{flexShrink:0,background:tab===t.id?divMeta.color:"transparent",border:tab===t.id?"none":`1px solid ${T.border}`,borderRadius:10,padding:"8px 10px",fontSize:11,fontWeight:tab===t.id?800:500,cursor:"pointer",color:tab===t.id?"#09090B":T.sub,fontFamily:"inherit",whiteSpace:"nowrap"}}>{t.icon} {t.label}</button>))}
        </div>
      </div>
      <div style={{padding:"14px 16px 80px"}}>
        <ErrBanner msg={err} onDismiss={()=>setErr("")}/>
        {loading&&<Spinner/>}
        {!loading&&tab==="reports"&&(<div>
          {can(user,"submit_report")&&<button onClick={()=>setScreen("newReport")} style={{...primBtn,marginBottom:16,borderRadius:14,padding:"18px",fontSize:17,background:divMeta.color}}>📋 + New Daily Report</button>}
          {reports.length===0&&<div style={{textAlign:"center",padding:"32px 0",color:T.muted}}><div style={{fontSize:36,marginBottom:8}}>📋</div><div style={{fontSize:15,fontWeight:600,color:T.sub,marginBottom:4}}>No reports yet</div></div>}
          {reports.map(r=>{const t=reportTotals(r);const sc={submitted:T.yellow,approved:T.green,flagged:T.red}[r.status||"submitted"]||T.muted;return(<div key={r.id} onClick={()=>{setActiveReport(r);setScreen("reportDetail");}} style={{...cardS,marginBottom:9,cursor:"pointer",borderLeft:`3px solid ${sc}`,display:"flex",justifyContent:"space-between",alignItems:"center"}}><div><div style={{display:"flex",alignItems:"center",gap:8}}><div style={{fontSize:15,fontWeight:700}}>{fmtShort(r.date)}</div><span style={pill(sc)}>{(r.status||"submitted").toUpperCase()}</span></div><div style={{fontSize:11,color:T.muted,marginTop:4,display:"flex",gap:8}}>{(r.labor||[]).length>0&&<span>👷 {r.labor.length}</span>}{(r.equipment||[]).length>0&&<span>🚜 {r.equipment.length}</span>}{r.submitted_by&&<span>by {r.submitted_by}</span>}</div></div><div style={{textAlign:"right"}}><div style={{fontSize:17,fontWeight:900,color:T.green}}>${fmt(t.grand)}</div><div style={{fontSize:9,color:T.muted}}>TOTAL</div></div></div>);})}
        </div>)}
        {!loading&&tab==="time"     &&can(user,"time_card")   &&<TimeCardsTab projectId={project.id} user={user} onErr={setErr}/>}
        {!loading&&tab==="crew"     &&can(user,"crew_equip")  &&<CrewEquipTab projectId={project.id} user={user} onErr={setErr}/>}
        {!loading&&tab==="subs"     &&can(user,"subs")        &&<SubsTab projectId={project.id} user={user} onErr={setErr}/>}
        {!loading&&tab==="safety"   &&can(user,"safety")      &&<SafetyTab projectId={project.id} safety={safety} user={user} onRefresh={()=>load(true)} onErr={setErr}/>}
        {!loading&&tab==="docs"     &&can(user,"docs")        &&<DocsTab projectId={project.id} user={user} onErr={setErr}/>}
        {!loading&&tab==="schedule" &&can(user,"schedule")    &&<ScheduleTab projectId={project.id} user={user} onErr={setErr}/>}
        {!loading&&tab==="photos"   &&can(user,"photos")      &&<PhotosTab projectId={project.id} photos={photos} onRefresh={()=>load(true)} onErr={setErr}/>}
        {!loading&&tab==="weather"  &&can(user,"weather")     &&<WeatherTab projectId={project.id} project={project} weather={weather} onRefresh={()=>load(true)} onErr={setErr}/>}
        {!loading&&tab==="info"     &&<InfoTab project={project} user={user} onEdit={()=>setEditProject(true)} onArchive={archiveProject}/>}
      </div>
    </div>
  );
}

/* ── PM DASHBOARD ───────────────────────────────────────────── */
function PMDashboard({onBack,user}){
  const [projects,setProjects]=useState([]);const [reports,setReports]=useState([]);const [pending,setPending]=useState([]);const [loading,setLoading]=useState(true);const [err,setErr]=useState("");const [pmTab,setPmTab]=useState("overview");
  const [activeReport,setActiveReport]=useState(null);const [activeProject,setActiveProject]=useState(null);const [unread,setUnread]=useState(0);const [showNotifs,setShowNotifs]=useState(false);

  async function load(){setLoading(true);setErr("");try{const[projs,reps,pend,notifs]=await Promise.all([API.projects.list(),API.reports.all(),API.reports.pending(),API.notifications.unread()]);setProjects(projs||[]);setReports(reps||[]);setPending(pend||[]);setUnread((notifs||[]).length);}catch(e){setErr(e.message);}setLoading(false);}
  useEffect(()=>{load();},[]);

  async function approve(id){try{await API.reports.update(id,{status:"approved",approved_by:user.name,approved_at:new Date().toISOString()});await load();}catch(e){setErr(e.message);}}
  async function flag(id,notes){try{await API.reports.update(id,{status:"flagged",pm_notes:notes});await notify("report_flagged","Report Flagged",notes,{report_id:id});await load();}catch(e){setErr(e.message);}}

  if(showNotifs) return(
    <div style={{background:T.bg,minHeight:"100vh",fontFamily:"inherit"}}>
      <div style={{background:T.surface,borderBottom:`1px solid ${T.border}`,padding:"14px 16px",position:"sticky",top:0,zIndex:50}}>
        <button onClick={()=>{setShowNotifs(false);load();}} style={{background:"none",border:"none",color:T.sub,fontSize:13,cursor:"pointer",marginBottom:8,padding:0,fontFamily:"inherit"}}>← Back</button>
        <div style={{fontSize:20,fontWeight:900}}>🔔 Notifications</div>
      </div>
      <NotificationsPanel onClose={()=>{setShowNotifs(false);load();}}/>
    </div>
  );

  if(activeReport&&activeProject) return(<ReportDetail report={activeReport} project={activeProject} user={user} onBack={()=>{setActiveReport(null);setActiveProject(null);load();}} onDelete={async(id)=>{await API.reports.remove(id);setActiveReport(null);setActiveProject(null);load();}} onApprove={approve} onFlag={flag}/>);

  const allTot=reports.reduce((s,r)=>{const t=reportTotals(r);return{l:s.l+t.labor,e:s.e+t.equip,m:s.m+t.mats,g:s.g+t.grand};},{l:0,e:0,m:0,g:0});
  const thisWeek=reports.filter(r=>{const d=new Date(r.date+"T12:00:00");return(Date.now()-d.getTime())/86400000<=7;});
  const monthStart=new Date();monthStart.setDate(1);const ms=monthStart.toISOString().split("T")[0];
  const workerHours={};
  reports.filter(r=>r.date>=ms).forEach(r=>(r.labor||[]).forEach(l=>{if(!l.name)return;if(!workerHours[l.name])workerHours[l.name]={name:l.name,reg:0,ot:0,travel:0,pay:0};workerHours[l.name].reg+=parseFloat(l.regHrs)||0;workerHours[l.name].ot+=parseFloat(l.otHrs)||0;workerHours[l.name].travel+=parseFloat(l.travelHrs)||0;workerHours[l.name].pay+=laborAmt(l);}));
  const workerRows=Object.values(workerHours).sort((a,b)=>b.pay-a.pay);
  const projMap={};projects.forEach(p=>{projMap[p.id]={...p,labor:0,equip:0,mats:0,grand:0,count:0};});
  reports.forEach(r=>{if(!projMap[r.project_id])return;const t=reportTotals(r);projMap[r.project_id].labor+=t.labor;projMap[r.project_id].equip+=t.equip;projMap[r.project_id].mats+=t.mats;projMap[r.project_id].grand+=t.grand;projMap[r.project_id].count++;});
  const projRows=Object.values(projMap).filter(p=>p.status==="active").sort((a,b)=>b.grand-a.grand);

  const DMTABS=[{id:"overview",l:"📊 Overview"},{id:"approvals",l:`✅ Approvals${pending.length>0?" ("+pending.length+")":""}`},{id:"workers",l:"👷 Workers"},{id:"billing",l:"💰 Billing"},{id:"reports",l:"📄 Reports"},{id:"users",l:"👤 Users"}];

  function CustomReports(){
    const [range,setRange]=useState("month");const [selProj,setSelProj]=useState("all");
    function getStart(){const d=new Date();if(range==="week"){d.setDate(d.getDate()-7);}else if(range==="month"){d.setDate(1);}else if(range==="quarter"){const q=Math.floor(d.getMonth()/3);d.setMonth(q*3);d.setDate(1);}else if(range==="year"){d.setMonth(0);d.setDate(1);}return d.toISOString().split("T")[0];}
    const fr=reports.filter(r=>{const start=range==="all"?"2000-01-01":getStart();const md=r.date>=start;const mp=selProj==="all"||r.project_id===selProj;return md&&mp;});
    const tot=fr.reduce((s,r)=>{const t=reportTotals(r);return{l:s.l+t.labor,e:s.e+t.equip,m:s.m+t.mats,g:s.g+t.grand};},{l:0,e:0,m:0,g:0});
    const wm={};fr.forEach(r=>(r.labor||[]).forEach(l=>{if(!l.name)return;if(!wm[l.name])wm[l.name]={name:l.name,reg:0,ot:0,travel:0,pay:0};wm[l.name].reg+=parseFloat(l.regHrs)||0;wm[l.name].ot+=parseFloat(l.otHrs)||0;wm[l.name].travel+=parseFloat(l.travelHrs)||0;wm[l.name].pay+=laborAmt(l);}));
    const wr=Object.values(wm).sort((a,b)=>b.pay-a.pay);
    const pm2={};projects.forEach(p=>{pm2[p.id]={name:p.name,division:p.division,labor:0,equip:0,mats:0,grand:0,count:0};});
    fr.forEach(r=>{if(!pm2[r.project_id])return;const t=reportTotals(r);pm2[r.project_id].labor+=t.labor;pm2[r.project_id].equip+=t.equip;pm2[r.project_id].mats+=t.mats;pm2[r.project_id].grand+=t.grand;pm2[r.project_id].count++;});
    const pr=Object.values(pm2).filter(p=>p.count>0).sort((a,b)=>b.grand-a.grand);
    function exportReport(){
      const wb=XLSX.utils.book_new();
      const sumRows=[["AIME Field OS — Custom Report"],[`Period: ${range} · Project: ${selProj==="all"?"All":projects.find(p=>p.id===selProj)?.name||"—"}`],[`Generated: ${new Date().toLocaleString()}`],[],["TOTALS"],["Labor","Equipment","Materials","Grand Total"],[tot.l,tot.e,tot.m,tot.g],[],["WORKERS"],["Name","Reg Hrs","OT Hrs","Travel Hrs","Total Pay"],...wr.map(w=>[w.name,w.reg,w.ot,w.travel,w.pay]),[],["PROJECTS"],["Project","Division","Reports","Labor","Equip","Mats","Total"],...pr.map(p=>[p.name,p.division||"",p.count,p.labor,p.equip,p.mats,p.grand])];
      const ws1=XLSX.utils.aoa_to_sheet(sumRows);ws1["!cols"]=[{wch:30},{wch:15},{wch:15},{wch:15},{wch:15},{wch:15},{wch:15}];XLSX.utils.book_append_sheet(wb,ws1,"Summary");
      const dRows=[["Date","Project","Division","By","Status","Labor","Equip","Mats","Total"]];fr.forEach(r=>{const t=reportTotals(r);const p=projects.find(x=>x.id===r.project_id);dRows.push([r.date,p?.name||"",p?.division||"",r.submitted_by||"",r.status||"",t.labor,t.equip,t.mats,t.grand]);});
      const ws2=XLSX.utils.aoa_to_sheet(dRows);XLSX.utils.book_append_sheet(wb,ws2,"Detail");
      XLSX.writeFile(wb,`AIME_Report_${range}_${today()}.xlsx`);
    }
    return(<div>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:14}}><div><label style={lbl}>Time Range</label><select value={range} onChange={e=>setRange(e.target.value)} style={inp}><option value="week">This Week</option><option value="month">This Month</option><option value="quarter">This Quarter</option><option value="year">This Year</option><option value="all">All Time</option></select></div><div><label style={lbl}>Project</label><select value={selProj} onChange={e=>setSelProj(e.target.value)} style={inp}><option value="all">All Projects</option>{projects.map(p=><option key={p.id} value={p.id}>{p.name}</option>)}</select></div></div>
      <div style={{...cardS,marginBottom:14,background:T.orangeLow,border:`1px solid ${T.orange}40`}}><div style={{fontSize:11,color:T.orange,fontWeight:700,textTransform:"uppercase",letterSpacing:"1px",marginBottom:10}}>{fr.length} Reports</div>{[["Labor",tot.l,T.green],["Equipment",tot.e,T.yellow],["Materials",tot.m,T.blue],["Grand Total",tot.g,T.orange]].map(([l,v,c])=>(<div key={l} style={{display:"flex",justifyContent:"space-between",padding:"7px 0",borderBottom:`1px solid ${T.border}`}}><span style={{fontSize:13,color:T.sub,fontWeight:l==="Grand Total"?700:400}}>{l}</span><span style={{fontSize:l==="Grand Total"?18:13,fontWeight:800,color:c}}>${fmt(v)}</span></div>))}</div>
      {/* By Division */}
      {["Mechanical","Pipeline","Structural"].map(div=>{const dr=fr.filter(r=>r.projects?.division===div||projects.find(p=>p.id===r.project_id)?.division===div);if(dr.length===0)return null;const dt=dr.reduce((s,r)=>{const t=reportTotals(r);return{g:s.g+t.grand};},{g:0});return(<div key={div} style={{...cardS,marginBottom:8,borderLeft:`3px solid ${DIV_META[div]?.color||T.border}`}}><div style={{display:"flex",justifyContent:"space-between"}}><div style={{fontSize:14,fontWeight:700}}>{DIV_META[div]?.icon} {div}</div><div style={{fontSize:15,fontWeight:800,color:T.green}}>${dt.g>=1000?(dt.g/1000).toFixed(1)+"k":fmt(dt.g)}</div></div><div style={{fontSize:11,color:T.muted,marginTop:2}}>{dr.length} reports</div></div>);})}
      {pr.length>0&&<div style={{...cardS,marginBottom:14}}><div style={{fontSize:12,fontWeight:700,color:T.muted,textTransform:"uppercase",letterSpacing:"1px",marginBottom:10}}>By Job</div>{pr.map(p=>(<div key={p.name} style={{display:"flex",justifyContent:"space-between",padding:"8px 0",borderBottom:`1px solid ${T.border}`}}><div><div style={{fontSize:13,fontWeight:700}}>{p.name}</div><div style={{fontSize:11,color:T.muted}}>{p.division} · {p.count} reports</div></div><div style={{fontSize:15,fontWeight:800,color:T.green}}>${fmt(p.grand)}</div></div>))}</div>}
      {wr.length>0&&<div style={{...cardS,marginBottom:14}}><div style={{fontSize:12,fontWeight:700,color:T.muted,textTransform:"uppercase",letterSpacing:"1px",marginBottom:10}}>By Worker</div>{wr.map(w=>(<div key={w.name} style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"8px 0",borderBottom:`1px solid ${T.border}`}}><div><div style={{fontSize:13,fontWeight:700}}>{w.name}</div><div style={{fontSize:11,color:T.muted}}>{w.reg.toFixed(1)}reg {w.ot.toFixed(1)}OT {w.travel.toFixed(1)}tr hrs</div></div><div style={{fontSize:14,fontWeight:800,color:T.green}}>${fmt(w.pay)}</div></div>))}</div>}
      <button onClick={exportReport} style={{...primBtn,borderRadius:14}}>📥 Export Full Report (.xlsx)</button>
    </div>);
  }

  function NotificationsPanel(){
    const [notifs,setNotifs]=useState([]);const [nl,setNl]=useState(true);
    async function loadN(){setNl(true);try{setNotifs(await API.notifications.list()||[]);}catch{}setNl(false);}
    useEffect(()=>{loadN();},[]);
    const typeIcon={report_submitted:"📋",report_flagged:"🚩",report_approved:"✅"};
    return(<div style={{padding:"14px 16px 80px"}}>
      {unread>0&&<button onClick={async()=>{await API.notifications.markAllRead();setUnread(0);await loadN();}} style={{...ghostBtn,width:"100%",textAlign:"center",marginBottom:14}}>Mark all read</button>}
      {nl&&<Spinner/>}
      {!nl&&notifs.length===0&&<div style={{textAlign:"center",padding:"40px 0",color:T.muted}}><div style={{fontSize:36,marginBottom:8}}>🔔</div><div>No notifications yet.</div></div>}
      {!nl&&notifs.map(n=>(<div key={n.id} onClick={async()=>{if(!n.read){await API.notifications.markRead(n.id);setUnread(u=>Math.max(0,u-1));await loadN();}}} style={{...cardS,marginBottom:8,borderLeft:`3px solid ${n.read?T.border:T.orange}`,opacity:n.read?0.6:1,cursor:n.read?"default":"pointer"}}><div style={{display:"flex",gap:10,alignItems:"flex-start"}}><span style={{fontSize:18,flexShrink:0}}>{typeIcon[n.type]||"📬"}</span><div style={{flex:1}}><div style={{fontSize:14,fontWeight:700}}>{n.title}</div>{n.body&&<div style={{fontSize:12,color:T.sub,marginTop:2}}>{n.body}</div>}<div style={{fontSize:11,color:T.muted,marginTop:4}}>{n.created_at?new Date(n.created_at).toLocaleString():""}</div></div>{!n.read&&<div style={{width:8,height:8,borderRadius:"50%",background:T.orange,flexShrink:0,marginTop:4}}/>}</div></div>))}
    </div>);
  }

  return(
    <div style={{background:T.bg,minHeight:"100vh",fontFamily:"inherit"}}>
      <div style={{background:T.surface,borderBottom:`1px solid ${T.border}`,padding:"14px 16px",position:"sticky",top:0,zIndex:50}}>
        <button onClick={onBack} style={{background:"none",border:"none",color:T.sub,fontSize:13,cursor:"pointer",marginBottom:8,padding:0,fontFamily:"inherit"}}>← Divisions</button>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:12}}>
          <div style={{fontSize:20,fontWeight:900,letterSpacing:"-0.5px"}}>PM Dashboard</div>
          <button onClick={()=>setShowNotifs(true)} style={{background:unread>0?T.orangeLow:T.card,border:`1px solid ${unread>0?T.orange:T.border}`,borderRadius:10,padding:"8px 12px",color:unread>0?T.orange:T.muted,fontSize:13,fontWeight:700,cursor:"pointer",fontFamily:"inherit",position:"relative"}}>
            🔔{unread>0&&<span style={{position:"absolute",top:-4,right:-4,width:16,height:16,borderRadius:"50%",background:T.orange,color:"#09090B",fontSize:9,fontWeight:900,display:"flex",alignItems:"center",justifyContent:"center"}}>{unread}</span>}
          </button>
        </div>
        <div style={{display:"flex",gap:6,overflowX:"auto"}}>
          {DMTABS.map(t=>(<button key={t.id} onClick={()=>setPmTab(t.id)} style={{flexShrink:0,background:pmTab===t.id?T.orange:"transparent",border:pmTab===t.id?"none":`1px solid ${T.border}`,borderRadius:10,padding:"9px 12px",fontSize:12,fontWeight:pmTab===t.id?800:500,cursor:"pointer",color:pmTab===t.id?"#09090B":T.sub,fontFamily:"inherit"}}>{t.l}</button>))}
        </div>
      </div>
      <div style={{padding:"14px 16px 80px"}}>
        <ErrBanner msg={err} onDismiss={()=>setErr("")}/>
        {loading&&<Spinner/>}
        {!loading&&(<>
          {pmTab==="overview"&&(<div>
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:20}}>
              {[{icon:"💰",label:"Total Billed",val:"$"+fmt(allTot.g),color:T.green},{icon:"📋",label:"This Week",val:thisWeek.length,color:T.orange},{icon:"🏗️",label:"Active Jobs",val:projects.filter(p=>p.status==="active").length,color:T.blue},{icon:"✅",label:"Pending",val:pending.length,color:pending.length>0?T.yellow:T.muted}].map(k=>(<div key={k.label} style={cardS}><div style={{fontSize:24,marginBottom:6}}>{k.icon}</div><div style={{fontSize:22,fontWeight:900,color:k.color,letterSpacing:"-0.5px"}}>{k.val}</div><div style={{fontSize:11,color:T.muted,marginTop:3}}>{k.label}</div></div>))}
            </div>
            {/* Division breakdown */}
            <div style={{fontSize:12,fontWeight:700,color:T.muted,textTransform:"uppercase",letterSpacing:"1px",marginBottom:12}}>By Division</div>
            {DIVISIONS.map(div=>{const dp=projects.filter(p=>p.division===div);const dr=reports.filter(r=>dp.find(p=>p.id===r.project_id));const dt=dr.reduce((s,r)=>{const t=reportTotals(r);return{g:s.g+t.grand};},{g:0});const m=DIV_META[div];return(<div key={div} style={{...cardS,marginBottom:10,borderLeft:`3px solid ${m.color}`}}><div style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}><div style={{display:"flex",alignItems:"center",gap:8}}><span style={{fontSize:20}}>{m.icon}</span><div><div style={{fontSize:15,fontWeight:700}}>{div}</div><div style={{fontSize:11,color:T.muted}}>{dp.filter(p=>p.status==="active").length} active jobs · {dr.length} reports</div></div></div><div style={{fontSize:18,fontWeight:900,color:T.green}}>${dt.g>=1000?(dt.g/1000).toFixed(1)+"k":fmt(dt.g)}</div></div></div>);})}
            <div style={{fontSize:12,fontWeight:700,color:T.muted,textTransform:"uppercase",letterSpacing:"1px",margin:"20px 0 12px"}}>Recent Reports</div>
            {reports.slice(0,15).map(r=>{const t=reportTotals(r);const p=projects.find(x=>x.id===r.project_id);const sc={submitted:T.yellow,approved:T.green,flagged:T.red}[r.status||"submitted"]||T.muted;return(<div key={r.id} onClick={()=>{setActiveReport(r);setActiveProject(p||{id:r.project_id,name:r.projects?.name||"Unknown",...(p||{})});}} style={{...cardS,marginBottom:8,cursor:"pointer",display:"flex",justifyContent:"space-between",alignItems:"center"}}><div><div style={{fontSize:13,fontWeight:700}}>{fmtShort(r.date)} · {r.projects?.name||"Unknown"}</div><div style={{fontSize:11,color:T.muted}}>{r.projects?.division||""} · {r.submitted_by||"Unknown"} · <span style={{color:sc}}>{(r.status||"submitted").toUpperCase()}</span></div></div><div style={{fontSize:15,fontWeight:800,color:T.green}}>${fmt(t.grand)}</div></div>);})}
            {reports.length===0&&<div style={{textAlign:"center",padding:"32px 0",color:T.muted}}>No reports yet.</div>}
          </div>)}
          {pmTab==="approvals"&&(<div>
            {pending.length===0&&<div style={{textAlign:"center",padding:"40px 0",color:T.muted}}><div style={{fontSize:40,marginBottom:10}}>✅</div><div style={{fontSize:16,fontWeight:700}}>All caught up!</div></div>}
            {pending.map(r=>{const t=reportTotals(r);const p=projects.find(x=>x.id===r.project_id);return(<div key={r.id} style={{...cardS,marginBottom:12,borderLeft:`3px solid ${T.yellow}`}}><div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:10}}><div><div style={{fontSize:15,fontWeight:800}}>{fmtShort(r.date)}</div><div style={{fontSize:12,color:T.sub}}>{r.projects?.name||"Unknown"} · {r.projects?.division||""}</div><div style={{fontSize:12,color:T.muted}}>by {r.submitted_by||"Unknown"}</div>{r.description&&<div style={{fontSize:12,color:T.sub,marginTop:4,lineHeight:1.4}}>{r.description.slice(0,100)}{r.description.length>100?"…":""}</div>}</div><div style={{textAlign:"right",flexShrink:0,marginLeft:10}}><div style={{fontSize:18,fontWeight:900,color:T.green}}>${fmt(t.grand)}</div><div style={{display:"flex",gap:4,marginTop:6,justifyContent:"flex-end"}}>{(r.labor||[]).length>0&&<span style={pill(T.orange)}>👷{r.labor.length}</span>}{(r.equipment||[]).length>0&&<span style={pill(T.yellow)}>🚜{r.equipment.length}</span>}</div></div></div><div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8,marginBottom:8}}><button onClick={()=>approve(r.id)} style={{...primBtn,background:T.greenLow,color:T.green,border:`1px solid ${T.green}40`,borderRadius:10,padding:"12px"}}>✓ Approve</button><button onClick={()=>{const n=window.prompt("Flag note:");if(n!==null)flag(r.id,n);}} style={{...primBtn,background:T.redLow,color:T.red,border:`1px solid ${T.red}40`,borderRadius:10,padding:"12px"}}>🚩 Flag</button></div><button onClick={()=>{setActiveReport(r);setActiveProject(p||{id:r.project_id,name:r.projects?.name||"Unknown",...(p||{})});}} style={{...ghostBtn,width:"100%",textAlign:"center",padding:"10px"}}>View Full Report →</button></div>);})}
          </div>)}
          {pmTab==="workers"&&(<div>
            <div style={{fontSize:12,color:T.muted,marginBottom:14}}>Hours from reports this month</div>
            {workerRows.length===0&&<div style={{textAlign:"center",padding:"40px 0",color:T.muted}}><div style={{fontSize:32,marginBottom:8}}>👷</div><div>No labor entries this month.</div></div>}
            {workerRows.map(w=>(<div key={w.name} style={{...cardS,marginBottom:9}}><div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start"}}><div><div style={{fontSize:15,fontWeight:700}}>{w.name}</div><div style={{display:"flex",gap:10,marginTop:6}}><span style={{fontSize:12,color:T.sub}}>{w.reg.toFixed(1)} reg</span>{w.ot>0&&<span style={{fontSize:12,color:T.yellow}}>{w.ot.toFixed(1)} OT</span>}{w.travel>0&&<span style={{fontSize:12,color:T.blue}}>{w.travel.toFixed(1)} travel</span>}</div></div><div style={{textAlign:"right"}}><div style={{fontSize:17,fontWeight:900,color:T.green}}>${fmt(w.pay)}</div><div style={{fontSize:10,color:T.muted}}>{(w.reg+w.ot+w.travel).toFixed(1)} total hrs</div></div></div></div>))}
          </div>)}
          {pmTab==="billing"&&(<div>
            <div style={{...cardS,marginBottom:16,background:T.orangeLow,border:`1px solid ${T.orange}40`}}><div style={{fontSize:12,color:T.orange,fontWeight:700,textTransform:"uppercase",letterSpacing:"1px",marginBottom:10}}>All-Time Totals</div>{[["Labor",allTot.l,T.green],["Equipment",allTot.e,T.yellow],["Materials",allTot.m,T.blue],["Grand Total",allTot.g,T.orange]].map(([l,v,c])=>(<div key={l} style={{display:"flex",justifyContent:"space-between",padding:"8px 0",borderBottom:`1px solid ${T.border}`}}><span style={{fontSize:14,color:T.sub,fontWeight:l==="Grand Total"?800:400}}>{l}</span><span style={{fontSize:l==="Grand Total"?20:14,fontWeight:800,color:c}}>${fmt(v)}</span></div>))}</div>
            <div style={{fontSize:12,fontWeight:700,color:T.muted,textTransform:"uppercase",letterSpacing:"1px",marginBottom:12}}>By Job</div>
            {projRows.map(p=>(<div key={p.id} style={{...cardS,marginBottom:10}}><div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:10}}><div><div style={{fontSize:15,fontWeight:700}}>{p.name}</div><div style={{fontSize:11,color:T.muted}}>{p.division} · {p.count} reports</div></div><div style={{fontSize:18,fontWeight:900,color:T.green}}>${p.grand>=1000?(p.grand/1000).toFixed(1)+"k":fmt(p.grand)}</div></div><div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:6}}>{[["Labor",p.labor,T.orange],["Equip",p.equip,T.yellow],["Mats",p.mats,T.blue]].map(([l,v,c])=>(<div key={l} style={{background:T.surface,borderRadius:8,padding:"8px",textAlign:"center"}}><div style={{fontSize:13,fontWeight:700,color:c}}>${v>=1000?(v/1000).toFixed(1)+"k":fmt(v)}</div><div style={{fontSize:9,color:T.muted,textTransform:"uppercase",letterSpacing:"0.5px"}}>{l}</div></div>))}</div></div>))}
          </div>)}
          {pmTab==="reports"&&<CustomReports/>}
          {pmTab==="users"&&(can(user,"manage_users")?<UserManagement/>:<div style={{textAlign:"center",padding:"40px 0",color:T.muted}}><div style={{fontSize:36,marginBottom:8}}>🔒</div><div>Admin access required.</div></div>)}
        </>)}
      </div>
    </div>
  );

  function UserManagement(){
    const [profs,setProfs]=useState([]);const [ul,setUl]=useState(true);const [uerr,setUerr]=useState("");
    async function loadU(){setUl(true);try{setProfs(await API.userProfiles.list()||[]);}catch(e){setUerr(e.message);}setUl(false);}
    useEffect(()=>{loadU();},[]);
    return(<div>
      <ErrBanner msg={uerr} onDismiss={()=>setUerr("")}/>
      <div style={{...cardS,marginBottom:14,background:T.blueLow,border:`1px solid ${T.blue}40`}}><div style={{fontSize:13,color:T.blue}}>👤 Go to the full User Management screen for complete control.</div></div>
      {ul&&<Spinner/>}
      {!ul&&profs.map(p=>{const m=ROLE_META[p.role]||ROLE_META.crew;return(<div key={p.id} style={{...cardS,marginBottom:8,borderLeft:`3px solid ${m.color}`}}><div style={{display:"flex",alignItems:"center",gap:10}}><div style={{width:10,height:10,borderRadius:"50%",background:m.color}}/><div style={{flex:1}}><div style={{fontSize:14,fontWeight:700}}>{p.name}</div><div style={{display:"flex",gap:6,marginTop:4}}><span style={pill(m.color)}>{m.label}</span>{p.division&&<span style={pill(DIV_META[p.division]?.color||T.muted)}>{p.division}</span>}</div></div></div></div>);})}
    </div>);
  }
}

/* ── ROOT APP ───────────────────────────────────────────────── */
export default function App(){
  const [user,setUser]           = useState(null);
  const [projects,setProjects]   = useState([]);
  const [loading,setLoading]     = useState(false);
  const [err,setErr]             = useState("");
  const [screen,setScreen]       = useState("division"); // division|jobboard|projectDetail|pmDashboard|crewDirectory|userManagement|newProject
  const [division,setDivision]   = useState(null);
  const [activeProject,setActiveProject] = useState(null);

  useEffect(()=>{
    const link=document.createElement("link");
    link.rel="stylesheet";
    link.href="https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600;700;800;900&display=swap";
    document.head.appendChild(link);
  },[]);

  async function loadProjects(){
    setLoading(true); setErr("");
    try{
      const projs = await API.projects.list();
      const enriched = await Promise.all((projs||[]).map(async p=>{
        try{
          const reps = await API.reports.forProject(p.id);
          const billed = (reps||[]).reduce((s,r)=>s+reportTotals(r).grand,0);
          const sorted = [...(reps||[])].sort((a,b)=>b.date>a.date?1:-1);
          return {...p,_reports:(reps||[]).length,_billed:billed,_lastReport:sorted[0]?.date||null};
        }catch{ return {...p,_reports:0,_billed:0,_lastReport:null}; }
      }));
      setProjects(enriched);
    }catch(e){ setErr(e.message); }
    setLoading(false);
  }

  useEffect(()=>{ if(user) loadProjects(); },[user]);

  async function handleNewProject(data){
    try{
      await API.projects.create({...data,created_by:user.name});
      await loadProjects();
      setScreen("jobboard");
    }catch(e){ setErr(e.message); }
  }

  function handleLogin(profile){
    setUser(profile);
    setScreen("division");
  }

  function handleLogout(){
    setUser(null); setProjects([]);
    setDivision(null); setScreen("division");
  }

  function handleDivisionSelect(div){
    setDivision(div);
    setScreen("jobboard");
  }

  function handleSelectProject(p){
    setActiveProject(p);
    setScreen("projectDetail");
  }

  if(!user) return(
    <div style={{maxWidth:480,margin:"0 auto",fontFamily:"'DM Sans',system-ui,sans-serif"}}>
      <LoginScreen onLogin={handleLogin}/>
    </div>
  );

  return(
    <div style={{maxWidth:480,margin:"0 auto",fontFamily:"'DM Sans',system-ui,sans-serif",background:T.bg,minHeight:"100vh"}}>
      {err&&(
        <div style={{position:"fixed",top:0,left:"50%",transform:"translateX(-50%)",width:"100%",maxWidth:480,zIndex:200,padding:"0 16px"}}>
          <ErrBanner msg={err} onDismiss={()=>setErr("")}/>
        </div>
      )}

      {/* DIVISION SELECTION */}
      {screen==="division"&&(
        <DivisionScreen
          user={user}
          projects={projects}
          onSelect={handleDivisionSelect}
          onLogout={handleLogout}
          onCrew={()=>setScreen("crewDirectory")}
          onDash={()=>setScreen("pmDashboard")}
        />
      )}

      {/* JOB BOARD for selected division */}
      {screen==="jobboard"&&division&&(
        <JobBoard
          user={user}
          division={division}
          projects={projects}
          loading={loading}
          onSelect={handleSelectProject}
          onNew={()=>setScreen("newProject")}
          onBack={()=>setScreen("division")}
        />
      )}

      {/* NEW PROJECT FORM */}
      {screen==="newProject"&&(
        <ProjectForm
          onSave={handleNewProject}
          onCancel={()=>setScreen("jobboard")}
          defaultDivision={division}
        />
      )}

      {/* PROJECT DETAIL */}
      {screen==="projectDetail"&&activeProject&&(
        <ProjectDetail
          project={projects.find(p=>p.id===activeProject.id)||activeProject}
          user={user}
          onBack={()=>{ setScreen("jobboard"); loadProjects(); }}
          onProjectUpdated={updated=>{ setActiveProject(p=>({...p,...updated})); loadProjects(); }}
        />
      )}

      {/* PM DASHBOARD */}
      {screen==="pmDashboard"&&(
        <PMDashboard
          user={user}
          onBack={()=>setScreen("division")}
        />
      )}

      {/* CREW DIRECTORY */}
      {screen==="crewDirectory"&&(
        <CrewDirectoryScreen
          user={user}
          onBack={()=>setScreen("division")}
        />
      )}

      {/* USER MANAGEMENT (admin only, also accessible from division screen) */}
      {screen==="userManagement"&&can(user,"manage_users")&&(
        <UserManagementScreen
          currentUser={user}
          onBack={()=>setScreen("division")}
        />
      )}
    </div>
  );
}
