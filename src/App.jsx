import { useState, useEffect, useRef } from "react";
import * as XLSX from "xlsx";

/* ═══════════════════════════════════════════════════════════════
   SUPABASE
═══════════════════════════════════════════════════════════════ */
const SUPA_URL = "https://uicmfyudiullulbbwzmh.supabase.co";
const SUPA_KEY = "sb_publishable_9h9AyvXpkp9glLxDVWRuGw_1eKVS7sE";

async function supa(path, { method = "GET", body, prefer } = {}) {
  const res = await fetch(`${SUPA_URL}/rest/v1${path}`, {
    method,
    headers: {
      "apikey": SUPA_KEY,
      "Authorization": `Bearer ${SUPA_KEY}`,
      "Content-Type": "application/json",
      ...(prefer ? { Prefer: prefer } : {}),
    },
    ...(body !== undefined ? { body: JSON.stringify(body) } : {}),
  });
  if (!res.ok) throw new Error(await res.text());
  const t = await res.text();
  return t ? JSON.parse(t) : null;
}

const API = {
  projects: {
    list:   ()      => supa("/projects?select=*&order=created_at.desc"),
    create: (d)     => supa("/projects", { method: "POST", body: d, prefer: "return=representation" }),
    update: (id, d) => supa(`/projects?id=eq.${id}`, { method: "PATCH", body: d, prefer: "return=representation" }),
    remove: (id)    => supa(`/projects?id=eq.${id}`, { method: "DELETE" }),
  },
  reports: {
    forProject: (pid) => supa(`/daily_reports?project_id=eq.${pid}&order=date.desc`),
    all:        ()    => supa("/daily_reports?select=*,projects(id,name)&order=date.desc&limit=500"),
    create:     (d)   => supa("/daily_reports", { method: "POST", body: d, prefer: "return=representation" }),
    remove:     (id)  => supa(`/daily_reports?id=eq.${id}`, { method: "DELETE" }),
  },
  safety: {
    forProject: (pid) => supa(`/safety_logs?project_id=eq.${pid}&order=created_at.desc`),
    create:     (d)   => supa("/safety_logs", { method: "POST", body: d, prefer: "return=representation" }),
    remove:     (id)  => supa(`/safety_logs?id=eq.${id}`, { method: "DELETE" }),
  },
  photos: {
    forProject: (pid) => supa(`/project_photos?project_id=eq.${pid}&order=created_at.desc`),
    create:     (d)   => supa("/project_photos", { method: "POST", body: d, prefer: "return=representation" }),
    remove:     (id)  => supa(`/project_photos?id=eq.${id}`, { method: "DELETE" }),
  },
};

/* ═══════════════════════════════════════════════════════════════
   THEME
═══════════════════════════════════════════════════════════════ */
const T = {
  bg: "#09090B", surface: "#111113", card: "#18181B",
  border: "#27272A", orange: "#F97316", orangeLow: "#F9731615",
  green: "#22C55E", greenLow: "#22C55E15", red: "#EF4444",
  redLow: "#EF444415", yellow: "#EAB308", blue: "#3B82F6",
  text: "#FAFAFA", sub: "#A1A1AA", muted: "#52525B",
};
const inp = {
  width:"100%", boxSizing:"border-box", background:"#0C0C0E",
  border:`1px solid ${T.border}`, borderRadius:12, color:T.text,
  fontSize:15, padding:"13px 14px", outline:"none",
  fontFamily:"inherit", appearance:"none", WebkitAppearance:"none",
};
const lbl = {
  display:"block", fontSize:11, fontWeight:700, color:T.muted,
  letterSpacing:"1px", textTransform:"uppercase", marginBottom:6,
};
const card = {
  background:T.card, border:`1px solid ${T.border}`, borderRadius:16, padding:"16px",
};
const pill = (c) => ({
  display:"inline-flex", alignItems:"center", background:c+"20",
  color:c, borderRadius:20, padding:"3px 10px", fontSize:11, fontWeight:700,
});
const primaryBtn = {
  width:"100%", background:T.orange, color:"#09090B", border:"none",
  borderRadius:14, padding:"16px", fontSize:16, fontWeight:800,
  cursor:"pointer", fontFamily:"inherit", letterSpacing:"-0.3px",
  display:"flex", alignItems:"center", justifyContent:"center", gap:8,
};
const ghostBtn = {
  background:"transparent", border:`1px solid ${T.border}`, borderRadius:12,
  padding:"12px 16px", color:T.sub, fontSize:14, cursor:"pointer",
  fontFamily:"inherit", fontWeight:600,
};

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
  {name:"Excavator - Mini - 10-11K LB",rate:335,unit:"Days"},{name:"Excavator - Mini - 12-16K LB",rate:475,unit:"Days"},
  {name:"Excavator - Small - 21-29K LB",rate:565,unit:"Days"},{name:"Excavator - Small - 30-33K LB",rate:632,unit:"Days"},
  {name:"Excavator - Medium - 48-55K LB",rate:852,unit:"Days"},{name:"Excavator - Large - 80-89K LB",rate:1050,unit:"Days"},
  {name:"Excavator - Large - 90-119K LB",rate:1350,unit:"Days"},{name:"Excavator - Large - 120-175K LB",rate:1750,unit:"Days"},
  {name:"Skidsteer Loader - 70-80 HP",rate:440,unit:"Days"},{name:"Skidsteer Loader - 81-100 HP",rate:475,unit:"Days"},
  {name:"Tractor - 50 HP 4x4 w/ Bush Hog",rate:36.50,unit:"Hours"},{name:"Mower - Riding/Zero Turn",rate:175,unit:"Days"},
  {section:"Air, Compressors & Blast"},
  {name:"Air Compressor - 185 CFM",rate:195,unit:"Days"},{name:"Air Compressor - 375 CFM",rate:275,unit:"Days"},
  {name:'Air Impact Wrench - 1"',rate:50,unit:"Days"},{name:"Air Spade / Knife",rate:55,unit:"Days"},
  {name:"Blast Rig - 4 Bag Pot w/ 185 CFM AC",rate:55.50,unit:"Hours"},{name:"Blast Rig - 1 Pot w/ 375 CFM AC",rate:500,unit:"Days"},
  {section:"Testing & Misc. Tools"},
  {name:"Holiday Detector / Pipe Jeep",rate:72,unit:"Days"},{name:"Hydraulic Torque",rate:200,unit:"Days"},
  {name:"Hydro Test Pump",rate:60,unit:"Days"},{name:"Hydrotest - High Pressure",rate:3800,unit:"Days"},
  {name:"Jack Hammer",rate:72,unit:"Days"},{name:"LEL/Gas Monitor - 4 Gas",rate:50,unit:"Days"},
  {name:"Line Locator",rate:50,unit:"Days"},{name:"Torque Wrench w/Sockets Hyd/Pneu",rate:195,unit:"Days"},
  {name:"Pipe Beveling Machine 16-22\"",rate:100,unit:"Days"},{name:"HEPA Vacuum",rate:100,unit:"Days"},
];

/* ═══════════════════════════════════════════════════════════════
   HELPERS
═══════════════════════════════════════════════════════════════ */
const uid = () => Math.random().toString(36).slice(2,9);
const today = () => new Date().toISOString().split("T")[0];
const fmt = (n) => Number(n||0).toLocaleString("en-US",{minimumFractionDigits:2,maximumFractionDigits:2});
const fmtDate = (d) => d ? new Date(d+"T12:00:00").toLocaleDateString("en-US",{month:"short",day:"numeric",year:"numeric"}) : "—";
const fmtShort = (d) => d ? new Date(d+"T12:00:00").toLocaleDateString("en-US",{weekday:"short",month:"short",day:"numeric"}) : "—";

function laborAmt(r) {
  const p = POSITIONS.find(x => x.name === r.classification);
  if (!p) return 0;
  if (p.flat) return p.rate;
  return p.rate * ((parseFloat(r.regHrs)||0) + (parseFloat(r.otHrs)||0)*1.5 + (parseFloat(r.travelHrs)||0));
}
function equipAmt(r) { return (parseFloat(r.rate)||0)*(parseFloat(r.qty)||0)*(parseFloat(r.usage)||0); }
function reportTotals(r) {
  const labor = (r.labor||[]).reduce((s,x)=>s+laborAmt(x),0);
  const equip = (r.equipment||[]).reduce((s,x)=>s+equipAmt(x),0);
  const mats  = (r.materials||[]).reduce((s,x)=>s+(parseFloat(x.amount)||0),0);
  return {labor,equip,mats,grand:labor+equip+mats};
}

async function compressImg(file, maxW=900, q=0.65) {
  return new Promise(res => {
    const reader = new FileReader();
    reader.onload = ev => {
      const img = new Image();
      img.onload = () => {
        const sc = Math.min(1, maxW/img.width);
        const c  = document.createElement("canvas");
        c.width  = Math.round(img.width*sc);
        c.height = Math.round(img.height*sc);
        c.getContext("2d").drawImage(img,0,0,c.width,c.height);
        res(c.toDataURL("image/jpeg",q));
      };
      img.src = ev.target.result;
    };
    reader.readAsDataURL(file);
  });
}

/* ═══════════════════════════════════════════════════════════════
   SHARED UI ATOMS
═══════════════════════════════════════════════════════════════ */
function Spinner() {
  return (
    <div style={{display:"flex",justifyContent:"center",padding:"40px 0"}}>
      <div style={{width:32,height:32,border:`3px solid ${T.border}`,borderTopColor:T.orange,borderRadius:"50%",animation:"spin 0.7s linear infinite"}} />
      <style>{`@keyframes spin{to{transform:rotate(360deg)}}`}</style>
    </div>
  );
}

function ErrBanner({msg,onDismiss}) {
  if(!msg) return null;
  return (
    <div style={{background:T.redLow,border:`1px solid ${T.red}44`,borderRadius:12,padding:"12px 14px",marginBottom:14,display:"flex",justifyContent:"space-between",alignItems:"center"}}>
      <span style={{fontSize:13,color:T.red}}>⚠️ {msg}</span>
      <button onClick={onDismiss} style={{background:"none",border:"none",color:T.red,cursor:"pointer",fontSize:16}}>×</button>
    </div>
  );
}

function TopBar({title,sub,onBack,right}) {
  return (
    <div style={{background:T.surface,borderBottom:`1px solid ${T.border}`,padding:"14px 16px",position:"sticky",top:0,zIndex:50}}>
      {onBack && <button onClick={onBack} style={{background:"none",border:"none",color:T.sub,fontSize:13,cursor:"pointer",marginBottom:8,padding:0,fontFamily:"inherit"}}>← Back</button>}
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start"}}>
        <div><div style={{fontSize:20,fontWeight:900,color:T.text,letterSpacing:"-0.5px"}}>{title}</div>{sub&&<div style={{fontSize:12,color:T.muted,marginTop:2}}>{sub}</div>}</div>
        {right}
      </div>
    </div>
  );
}

function Lightbox({src,onClose}) {
  if(!src) return null;
  return (
    <div onClick={onClose} style={{position:"fixed",inset:0,zIndex:999,background:"rgba(0,0,0,0.95)",display:"flex",alignItems:"center",justifyContent:"center",padding:16}}>
      <img src={src} alt="" style={{maxWidth:"100%",maxHeight:"90vh",borderRadius:12}} onClick={e=>e.stopPropagation()} />
      <button onClick={onClose} style={{position:"absolute",top:16,right:16,background:"#333",border:"none",color:"#fff",borderRadius:"50%",width:36,height:36,fontSize:18,cursor:"pointer"}}>×</button>
    </div>
  );
}

/* ═══════════════════════════════════════════════════════════════
   LOGIN
═══════════════════════════════════════════════════════════════ */
function LoginScreen({onLogin}) {
  const [name,setName] = useState("");
  const [role,setRole] = useState("crew");
  const [pin,setPin]   = useState("");
  return (
    <div style={{minHeight:"100vh",background:T.bg,display:"flex",flexDirection:"column",justifyContent:"center",padding:24,fontFamily:"inherit"}}>
      <div style={{textAlign:"center",marginBottom:40}}>
        <div style={{display:"inline-flex",alignItems:"center",gap:12,marginBottom:10}}>
          <div style={{width:52,height:52,background:T.orange,borderRadius:16,display:"flex",alignItems:"center",justifyContent:"center",fontSize:26,fontWeight:900,color:"#09090B"}}>A</div>
          <div style={{textAlign:"left"}}>
            <div style={{fontSize:24,fontWeight:900,color:T.text,letterSpacing:"-1px"}}>AIME Field OS</div>
            <div style={{fontSize:11,color:T.muted,letterSpacing:"2px",textTransform:"uppercase"}}>Colonial Pipeline</div>
          </div>
        </div>
      </div>
      <div style={{...card,maxWidth:400,margin:"0 auto",width:"100%"}}>
        <div style={{marginBottom:14}}>
          <label style={lbl}>Your Name</label>
          <select value={name} onChange={e=>setName(e.target.value)} style={inp}>
            <option value="">— Select your name —</option>
            {NAMES.map(n=><option key={n} value={n}>{n}</option>)}
          </select>
        </div>
        <div style={{marginBottom:14}}>
          <label style={lbl}>I am a…</label>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8}}>
            {["crew","pm"].map(r=>(
              <button key={r} onClick={()=>setRole(r)} style={{padding:"14px",borderRadius:12,border:`2px solid ${role===r?T.orange:T.border}`,background:role===r?T.orangeLow:T.surface||"#111",color:role===r?T.orange:T.sub,fontWeight:700,fontSize:14,cursor:"pointer",fontFamily:"inherit"}}>
                {r==="crew"?"👷 Field Crew":"📋 Project Manager"}
              </button>
            ))}
          </div>
        </div>
        {role==="pm"&&(
          <div style={{marginBottom:14}}>
            <label style={lbl}>PM PIN</label>
            <input type="password" maxLength={4} placeholder="4-digit PIN (default: 1234)" value={pin} onChange={e=>setPin(e.target.value)} style={inp} />
          </div>
        )}
        <button onClick={()=>{if(!name)return;if(role==="pm"&&pin!=="1234"){alert("Incorrect PIN");return;}onLogin({name:name.trim(),role});}} style={{...primaryBtn,opacity:name?1:0.5}}>
          Sign In →
        </button>
        {role==="pm"&&<div style={{fontSize:11,color:T.muted,textAlign:"center",marginTop:8}}>Default PIN: 1234</div>}
      </div>
    </div>
  );
}

/* ═══════════════════════════════════════════════════════════════
   PROJECTS HOME
═══════════════════════════════════════════════════════════════ */
function ProjectsHome({user,projects,loading,onSelect,onNew,onLogout,onDash}) {
  const active   = projects.filter(p=>p.status==="active");
  const archived = projects.filter(p=>p.status!=="active");
  const totalReports = projects.reduce((s,p)=>s+(p._reportCount||0),0);
  const totalBilled  = projects.reduce((s,p)=>s+(p._totalBilled||0),0);

  return (
    <div style={{fontFamily:"inherit",background:T.bg,minHeight:"100vh"}}>
      <div style={{background:T.surface,borderBottom:`1px solid ${T.border}`,padding:"16px"}}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:14}}>
          <div style={{display:"flex",alignItems:"center",gap:10}}>
            <div style={{width:36,height:36,background:T.orange,borderRadius:10,display:"flex",alignItems:"center",justifyContent:"center",fontWeight:900,fontSize:18,color:"#09090B"}}>A</div>
            <div>
              <div style={{fontSize:16,fontWeight:800,color:T.text}}>AIME Field OS</div>
              <div style={{fontSize:11,color:T.muted}}>👋 {user.name} · {user.role==="pm"?"PM":"Field Crew"}</div>
            </div>
          </div>
          <button onClick={onLogout} style={{...ghostBtn,padding:"8px 12px",fontSize:12}}>Sign Out</button>
        </div>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:8}}>
          {[
            {label:"Projects",val:active.length,color:T.orange},
            {label:"Reports",val:totalReports,color:T.green},
            {label:"Billed",val:"$"+(totalBilled>=1000?(totalBilled/1000).toFixed(1)+"k":fmt(totalBilled)),color:T.blue},
          ].map(s=>(
            <div key={s.label} style={{background:T.card,border:`1px solid ${T.border}`,borderRadius:12,padding:"12px 8px",textAlign:"center"}}>
              <div style={{fontSize:18,fontWeight:900,color:s.color}}>{s.val}</div>
              <div style={{fontSize:9,color:T.muted,textTransform:"uppercase",letterSpacing:"0.8px",marginTop:2}}>{s.label}</div>
            </div>
          ))}
        </div>
      </div>

      <div style={{padding:"14px 16px 100px"}}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:12}}>
          <div style={{fontSize:12,fontWeight:700,color:T.muted,textTransform:"uppercase",letterSpacing:"1px"}}>Active Projects</div>
          {user.role==="pm"&&<button onClick={onNew} style={{background:T.orange,color:"#09090B",border:"none",borderRadius:10,padding:"8px 14px",fontSize:13,fontWeight:800,cursor:"pointer",fontFamily:"inherit"}}>+ New Project</button>}
        </div>

        {loading&&<Spinner/>}
        {!loading&&active.length===0&&<div style={{textAlign:"center",padding:"40px 0",color:T.muted}}><div style={{fontSize:36,marginBottom:8}}>🏗️</div><div>No active projects. {user.role==="pm"?"Tap + New Project to get started.":"Check back soon."}</div></div>}
        {active.map(p=><ProjectCard key={p.id} project={p} onSelect={onSelect}/>)}

        {archived.length>0&&(
          <>
            <div style={{fontSize:12,fontWeight:700,color:T.muted,textTransform:"uppercase",letterSpacing:"1px",marginTop:24,marginBottom:10}}>Archived</div>
            {archived.map(p=><ProjectCard key={p.id} project={p} onSelect={onSelect} dim/>)}
          </>
        )}
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

function ProjectCard({project:p,onSelect,dim}) {
  return (
    <div onClick={()=>onSelect(p)} style={{...card,marginBottom:10,cursor:"pointer",borderLeft:`3px solid ${dim?T.border:T.orange}`,opacity:dim?0.5:1,display:"flex",justifyContent:"space-between",alignItems:"center"}}>
      <div style={{flex:1,minWidth:0}}>
        <div style={{fontSize:15,fontWeight:800,color:T.text}}>{p.name}</div>
        <div style={{fontSize:12,color:T.sub,marginTop:2}}>{p.client||"No client"}{p.location?" · "+p.location:""}</div>
        <div style={{display:"flex",gap:6,marginTop:7,flexWrap:"wrap"}}>
          <span style={pill(T.orange)}>{p._reportCount||0} report{p._reportCount!==1?"s":""}</span>
          {p.afe&&<span style={pill(T.muted)}>AFE: {p.afe}</span>}
          <span style={pill(p.status==="active"?T.green:T.muted)}>{p.status}</span>
        </div>
      </div>
      <div style={{textAlign:"right",flexShrink:0,marginLeft:12}}>
        <div style={{fontSize:18,fontWeight:900,color:T.green}}>${p._totalBilled>=1000?(p._totalBilled/1000).toFixed(1)+"k":fmt(p._totalBilled||0)}</div>
        <div style={{fontSize:9,color:T.muted,textTransform:"uppercase",marginTop:2}}>Billed</div>
      </div>
    </div>
  );
}

/* ═══════════════════════════════════════════════════════════════
   NEW / EDIT PROJECT FORM
═══════════════════════════════════════════════════════════════ */
function ProjectForm({initial,onSave,onCancel,saving}) {
  const [f,setF] = useState(initial||{name:"",client:"",location:"",afe:"",work_order:"",start_date:today(),notes:"",status:"active"});
  const set = (k,v) => setF(x=>({...x,[k]:v}));
  return (
    <div style={{fontFamily:"inherit",background:T.bg,minHeight:"100vh"}}>
      <TopBar title={initial?"Edit Project":"New Project"} onBack={onCancel}/>
      <div style={{padding:"16px 16px 100px"}}>
        {[
          {k:"name",l:"Project Name *",ph:"e.g. HDD Crossing – Station 42"},
          {k:"client",l:"Client",ph:"Colonial Pipeline"},
          {k:"location",l:"Location",ph:"City, State or Milepost"},
          {k:"afe",l:"AFE No.",ph:"AFE #"},
          {k:"work_order",l:"Work Order #",ph:"WO #"},
        ].map(({k,l,ph})=>(
          <div key={k} style={{marginBottom:12}}>
            <label style={lbl}>{l}</label>
            <input type="text" placeholder={ph} value={f[k]||""} onChange={e=>set(k,e.target.value)} style={inp}/>
          </div>
        ))}
        <div style={{marginBottom:12}}>
          <label style={lbl}>Start Date</label>
          <input type="date" value={f.start_date||today()} onChange={e=>set("start_date",e.target.value)} style={inp}/>
        </div>
        <div style={{marginBottom:20}}>
          <label style={lbl}>Notes</label>
          <textarea placeholder="Project notes, scope, special instructions…" value={f.notes||""} onChange={e=>set("notes",e.target.value)} rows={3} style={{...inp,resize:"vertical",lineHeight:1.5}}/>
        </div>
        <button onClick={()=>f.name.trim()&&onSave(f)} style={{...primaryBtn,opacity:f.name.trim()&&!saving?1:0.5}}>
          {saving?"Saving…":(initial?"Save Changes":"Create Project")}
        </button>
      </div>
    </div>
  );
}

/* ═══════════════════════════════════════════════════════════════
   PROJECT DETAIL
═══════════════════════════════════════════════════════════════ */
function ProjectDetail({project:initProject,user,onBack,onProjectUpdated}) {
  const [project,setProject] = useState(initProject);
  const [reports,setReports] = useState([]);
  const [safety,setSafety]   = useState([]);
  const [photos,setPhotos]   = useState([]);
  const [tab,setTab]         = useState("reports");
  const [loading,setLoading] = useState(true);
  const [err,setErr]         = useState("");
  const [screen,setScreen]   = useState("detail"); // detail | newReport | reportDetail | editProject
  const [activeReport,setActiveReport] = useState(null);

  async function load() {
    setLoading(true); setErr("");
    try {
      const [reps,saf,phs] = await Promise.all([
        API.reports.forProject(project.id),
        API.safety.forProject(project.id),
        API.photos.forProject(project.id),
      ]);
      setReports(reps||[]);
      setSafety(saf||[]);
      setPhotos(phs||[]);
    } catch(e) { setErr(e.message); }
    setLoading(false);
  }

  useEffect(()=>{ load(); },[project.id]);

  async function saveReport(reportData) {
    try {
      await API.reports.create({...reportData,project_id:project.id});
      await load();
      setScreen("detail");
    } catch(e) { setErr(e.message); }
  }

  async function deleteReport(id) {
    try { await API.reports.remove(id); setActiveReport(null); await load(); setScreen("detail"); }
    catch(e) { setErr(e.message); }
  }

  async function updateProject(data) {
    try {
      const [updated] = await API.projects.update(project.id,data);
      setProject(updated);
      onProjectUpdated(updated);
      setScreen("detail");
    } catch(e) { setErr(e.message); }
  }

  async function archiveProject() {
    if(!window.confirm(project.status==="active"?"Archive this project?":"Restore this project?")) return;
    await updateProject({status:project.status==="active"?"archived":"active"});
    onBack();
  }

  const tot = reports.reduce((s,r)=>{const t=reportTotals(r);return{labor:s.labor+t.labor,equip:s.equip+t.equip,mats:s.mats+t.mats,grand:s.grand+t.grand};},{labor:0,equip:0,mats:0,grand:0});

  if(screen==="newReport") return <DailyReportForm user={user} project={project} onSave={saveReport} onCancel={()=>setScreen("detail")}/>;
  if(screen==="reportDetail"&&activeReport) return <ReportDetail report={activeReport} project={project} onBack={()=>setScreen("detail")} onDelete={deleteReport}/>;
  if(screen==="editProject") return <ProjectForm initial={project} onSave={updateProject} onCancel={()=>setScreen("detail")}/>;

  const TABS=[{id:"reports",icon:"📋",label:"Reports"},{id:"safety",icon:"⛑️",label:"Safety"},{id:"photos",icon:"📷",label:"Photos"},{id:"info",icon:"ℹ️",label:"Info"}];

  return (
    <div style={{fontFamily:"inherit",background:T.bg,minHeight:"100vh"}}>
      <div style={{background:T.surface,borderBottom:`1px solid ${T.border}`,padding:"14px 16px",position:"sticky",top:0,zIndex:50}}>
        <button onClick={onBack} style={{background:"none",border:"none",color:T.sub,fontSize:13,cursor:"pointer",marginBottom:8,padding:0,fontFamily:"inherit"}}>← Projects</button>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start"}}>
          <div><div style={{fontSize:19,fontWeight:900,color:T.text,letterSpacing:"-0.4px"}}>{project.name}</div><div style={{fontSize:12,color:T.sub,marginTop:2}}>{project.client}{project.location?" · "+project.location:""}</div></div>
          <span style={pill(project.status==="active"?T.green:T.muted)}>{project.status}</span>
        </div>
        <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:6,marginTop:12}}>
          {[["Reports",reports.length,T.orange],["Labor","$"+(tot.labor>=1000?(tot.labor/1000).toFixed(1)+"k":fmt(tot.labor)),T.green],["Equip","$"+(tot.equip>=1000?(tot.equip/1000).toFixed(1)+"k":fmt(tot.equip)),T.yellow],["Total","$"+(tot.grand>=1000?(tot.grand/1000).toFixed(1)+"k":fmt(tot.grand)),T.blue]].map(([l,v,c])=>(
            <div key={l} style={{background:T.card,borderRadius:10,padding:"9px 6px",textAlign:"center"}}>
              <div style={{fontSize:13,fontWeight:800,color:c}}>{v}</div>
              <div style={{fontSize:9,color:T.muted,textTransform:"uppercase",letterSpacing:"0.6px",marginTop:1}}>{l}</div>
            </div>
          ))}
        </div>
        <div style={{display:"flex",gap:4,marginTop:12}}>
          {TABS.map(t=>(
            <button key={t.id} onClick={()=>setTab(t.id)} style={{flex:1,background:tab===t.id?T.orange:"transparent",border:tab===t.id?"none":`1px solid ${T.border}`,borderRadius:10,padding:"9px 4px",fontSize:11,fontWeight:tab===t.id?800:500,cursor:"pointer",color:tab===t.id?"#09090B":T.sub,fontFamily:"inherit"}}>
              {t.icon} {t.label}
            </button>
          ))}
        </div>
      </div>

      <div style={{padding:"14px 16px 80px"}}>
        <ErrBanner msg={err} onDismiss={()=>setErr("")}/>
        {loading&&<Spinner/>}

        {!loading&&tab==="reports"&&(
          <>
            <button onClick={()=>setScreen("newReport")} style={{...primaryBtn,marginBottom:14,borderRadius:14}}>+ New Daily Report</button>
            {reports.length===0&&<div style={{textAlign:"center",padding:"32px 0",color:T.muted}}><div style={{fontSize:32,marginBottom:8}}>📋</div><div>No reports yet.</div></div>}
            {reports.map(r=>{
              const t=reportTotals(r);
              return(
                <div key={r.id} onClick={()=>{setActiveReport(r);setScreen("reportDetail");}} style={{...card,marginBottom:9,cursor:"pointer",borderLeft:`3px solid ${T.orange}`,display:"flex",justifyContent:"space-between",alignItems:"center"}}>
                  <div>
                    <div style={{fontSize:15,fontWeight:700}}>{fmtShort(r.date)}</div>
                    <div style={{fontSize:11,color:T.muted,marginTop:3,display:"flex",gap:8}}>
                      {(r.labor||[]).length>0&&<span>👷 {r.labor.length} workers</span>}
                      {(r.equipment||[]).length>0&&<span>🚜 {r.equipment.length} equip</span>}
                      {r.submitted_by&&<span>by {r.submitted_by}</span>}
                    </div>
                  </div>
                  <div style={{textAlign:"right"}}>
                    <div style={{fontSize:17,fontWeight:900,color:T.green}}>${fmt(t.grand)}</div>
                    <div style={{fontSize:9,color:T.muted}}>TOTAL</div>
                  </div>
                </div>
              );
            })}
          </>
        )}

        {!loading&&tab==="safety"&&<SafetyTab projectId={project.id} safety={safety} user={user} onRefresh={load} onErr={setErr}/>}
        {!loading&&tab==="photos"&&<PhotosTab projectId={project.id} photos={photos} onRefresh={load} onErr={setErr}/>}
        {!loading&&tab==="info"&&(
          <div>
            {[["Client",project.client],["Location",project.location],["AFE No.",project.afe],["Work Order",project.work_order],["Start Date",fmtDate(project.start_date)]].map(([l,v])=>v?(
              <div key={l} style={{display:"flex",justifyContent:"space-between",padding:"11px 0",borderBottom:`1px solid ${T.border}`}}>
                <span style={{fontSize:13,color:T.muted}}>{l}</span>
                <span style={{fontSize:13,fontWeight:600}}>{v}</span>
              </div>
            ):null)}
            {project.notes&&<div style={{...card,marginTop:14}}><div style={{fontSize:11,color:T.muted,textTransform:"uppercase",letterSpacing:"1px",marginBottom:6}}>Notes</div><div style={{fontSize:14,color:T.sub,lineHeight:1.6}}>{project.notes}</div></div>}
            {user.role==="pm"&&(
              <div style={{marginTop:18,display:"flex",flexDirection:"column",gap:8}}>
                <button onClick={()=>setScreen("editProject")} style={{...ghostBtn,width:"100%",textAlign:"center"}}>✏️ Edit Project</button>
                <button onClick={archiveProject} style={{...ghostBtn,width:"100%",textAlign:"center",color:T.muted}}>{project.status==="active"?"📦 Archive Project":"♻️ Restore Project"}</button>
              </div>
            )}
          </div>
        )}
      </div>
    </div>
  );
}

/* ═══════════════════════════════════════════════════════════════
   SAFETY TAB
═══════════════════════════════════════════════════════════════ */
function SafetyTab({projectId,safety,user,onRefresh,onErr}) {
  const [showForm,setShowForm] = useState(false);
  const [saving,setSaving]     = useState(false);
  const [type,setType]         = useState("toolbox");
  const [f,setF]               = useState({date:today(),topic:"",notes:"",severity:"low"});
  const TC={toolbox:T.blue,observation:T.yellow,incident:T.red,nearmiss:T.orange};
  const TL={toolbox:"🛠 Toolbox Talk",observation:"👁 Observation",incident:"🚨 Incident",nearmiss:"⚠️ Near Miss"};

  async function save() {
    if(!f.topic.trim()) return;
    setSaving(true);
    try { await API.safety.create({...f,type,project_id:projectId,created_by:user.name}); await onRefresh(); setShowForm(false); setF({date:today(),topic:"",notes:"",severity:"low"}); }
    catch(e) { onErr(e.message); }
    setSaving(false);
  }
  async function del(id) {
    try { await API.safety.remove(id); await onRefresh(); } catch(e) { onErr(e.message); }
  }

  return (
    <div>
      <button onClick={()=>setShowForm(!showForm)} style={{...primaryBtn,marginBottom:14,borderRadius:14}}>{showForm?"✕ Cancel":"+ Log Safety Entry"}</button>
      {showForm&&(
        <div style={{...card,marginBottom:14,borderLeft:`3px solid ${T.yellow}`}}>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8,marginBottom:12}}>
            {Object.entries(TL).map(([k,v])=>(
              <button key={k} onClick={()=>setType(k)} style={{padding:"10px",borderRadius:10,border:`2px solid ${type===k?TC[k]:T.border}`,background:type===k?TC[k]+"20":T.surface||"#111",color:type===k?TC[k]:T.sub,fontSize:12,fontWeight:700,cursor:"pointer",fontFamily:"inherit"}}>{v}</button>
            ))}
          </div>
          <div style={{marginBottom:10}}><label style={lbl}>Date</label><input type="date" value={f.date} onChange={e=>setF(x=>({...x,date:e.target.value}))} style={inp}/></div>
          <div style={{marginBottom:10}}><label style={lbl}>{type==="toolbox"?"Topic / Title":"Description"}</label><input type="text" placeholder="Describe…" value={f.topic} onChange={e=>setF(x=>({...x,topic:e.target.value}))} style={inp}/></div>
          <div style={{marginBottom:10}}><label style={lbl}>Notes / Corrective Action</label><textarea rows={3} placeholder="Additional details…" value={f.notes} onChange={e=>setF(x=>({...x,notes:e.target.value}))} style={{...inp,resize:"vertical"}}/></div>
          {(type==="incident"||type==="nearmiss")&&(
            <div style={{marginBottom:10}}><label style={lbl}>Severity</label>
              <select value={f.severity} onChange={e=>setF(x=>({...x,severity:e.target.value}))} style={inp}>
                <option value="low">Low – First Aid</option><option value="medium">Medium – Recordable</option><option value="high">High – Lost Time</option>
              </select>
            </div>
          )}
          <button onClick={save} style={{...primaryBtn,background:T.yellow,color:"#09090B"}}>{saving?"Saving…":"Save Entry"}</button>
        </div>
      )}
      {safety.length===0&&!showForm&&<div style={{textAlign:"center",padding:"40px 0",color:T.muted}}><div style={{fontSize:32,marginBottom:8}}>⛑️</div><div>No safety entries. Log toolbox talks, observations and incidents here.</div></div>}
      {[...safety].map(s=>(
        <div key={s.id} style={{...card,marginBottom:9,borderLeft:`3px solid ${TC[s.type]||T.border}`}}>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start"}}>
            <div>
              <span style={{...pill(TC[s.type]||T.muted),marginBottom:6,display:"inline-flex"}}>{TL[s.type]||s.type}</span>
              <div style={{fontSize:14,fontWeight:700,marginTop:4}}>{s.topic}</div>
              {s.notes&&<div style={{fontSize:12,color:T.sub,marginTop:4}}>{s.notes}</div>}
              <div style={{fontSize:11,color:T.muted,marginTop:6}}>{fmtDate(s.date)} · {s.created_by}</div>
            </div>
            <div style={{display:"flex",flexDirection:"column",alignItems:"flex-end",gap:6}}>
              {(s.type==="incident"||s.type==="nearmiss")&&<span style={pill(s.severity==="high"?T.red:s.severity==="medium"?T.yellow:T.green)}>{(s.severity||"").toUpperCase()}</span>}
              {user.role==="pm"&&<button onClick={()=>del(s.id)} style={{background:"none",border:"none",color:T.muted,cursor:"pointer",fontSize:16,padding:0}}>🗑</button>}
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
  const [caption,setCaption] = useState("");
  const [saving,setSaving]   = useState(false);
  const [lb,setLb]           = useState(null);
  const fileRef = useRef(null);

  async function handleFiles(files) {
    setSaving(true);
    try {
      for(const f of files) {
        if(!f.type.startsWith("image/")) continue;
        const src = await compressImg(f,1000,0.7);
        await API.photos.create({project_id:projectId,src,caption,date:today()});
      }
      await onRefresh(); setCaption("");
    } catch(e) { onErr(e.message); }
    setSaving(false);
  }

  async function del(id) {
    try { await API.photos.remove(id); await onRefresh(); } catch(e) { onErr(e.message); }
  }

  return (
    <div>
      <Lightbox src={lb} onClose={()=>setLb(null)}/>
      <div style={{...card,marginBottom:14,borderStyle:"dashed",borderColor:T.orange+"44"}}>
        <div style={{marginBottom:10}}><label style={lbl}>Caption (optional)</label><input type="text" placeholder="What's this photo of?" value={caption} onChange={e=>setCaption(e.target.value)} style={inp}/></div>
        <button onClick={()=>fileRef.current?.click()} style={{...primaryBtn}}>{saving?"Uploading…":"📷 Add Site Photos"}</button>
        <input ref={fileRef} type="file" accept="image/*" capture="environment" multiple style={{display:"none"}} onChange={e=>{handleFiles(Array.from(e.target.files));e.target.value="";}}/>
      </div>
      {photos.length===0&&<div style={{textAlign:"center",padding:"40px 0",color:T.muted}}><div style={{fontSize:32,marginBottom:8}}>📷</div><div>No photos yet.</div></div>}
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}>
        {photos.map(p=>(
          <div key={p.id} style={{position:"relative",borderRadius:12,overflow:"hidden",aspectRatio:"4/3",background:T.card}}>
            <img src={p.src} alt={p.caption} onClick={()=>setLb(p.src)} style={{width:"100%",height:"100%",objectFit:"cover",cursor:"pointer",display:"block"}}/>
            <button onClick={()=>del(p.id)} style={{position:"absolute",top:6,right:6,background:"rgba(0,0,0,0.7)",border:"none",color:"#fff",borderRadius:"50%",width:24,height:24,fontSize:12,cursor:"pointer"}}>×</button>
            {p.caption&&<div style={{position:"absolute",bottom:0,left:0,right:0,background:"rgba(0,0,0,0.65)",color:"#fff",fontSize:11,padding:"4px 8px"}}>{p.caption}</div>}
          </div>
        ))}
      </div>
    </div>
  );
}

/* ═══════════════════════════════════════════════════════════════
   DAILY REPORT FORM
═══════════════════════════════════════════════════════════════ */
const STEPS=["Job Info","Labor","Equipment","Materials","Review"];

function DailyReportForm({user,project,onSave,onCancel}) {
  const [step,setStep] = useState(1);
  const [saving,setSaving] = useState(false);
  const [report,setReport] = useState({date:today(),description:"",report_no:"",labor:[],equipment:[],materials:[]});
  const topRef = useRef(null);
  const setR = (k,v) => setReport(r=>({...r,[k]:v}));

  function addLabor()  { setR("labor",[...report.labor,{id:uid(),name:"",classification:"",regHrs:"",otHrs:"",travelHrs:""}]); }
  function updLabor(i,row) { const a=[...report.labor]; a[i]=row; setR("labor",a); }
  function delLabor(i) { setR("labor",report.labor.filter((_,j)=>j!==i)); }
  function addEquip()  { setR("equipment",[...report.equipment,{id:uid(),description:"",qty:"",usage:"",rate:"",unit:""}]); }
  function updEquip(i,row) { const a=[...report.equipment]; a[i]=row; setR("equipment",a); }
  function delEquip(i) { setR("equipment",report.equipment.filter((_,j)=>j!==i)); }
  function addMat()    { setR("materials",[...report.materials,{id:uid(),qty:"",description:"",amount:"",receipts:[]}]); }
  function updMat(i,row) { const a=[...report.materials]; a[i]=row; setR("materials",a); }
  function delMat(i)   { setR("materials",report.materials.filter((_,j)=>j!==i)); }

  const tot = reportTotals(report);

  async function submit() {
    setSaving(true);
    await onSave({...report,submitted_by:user.name});
    setSaving(false);
  }

  return (
    <div ref={topRef} style={{fontFamily:"inherit",background:T.bg,minHeight:"100vh"}}>
      <div style={{background:T.surface,borderBottom:`1px solid ${T.border}`,padding:"14px 16px",position:"sticky",top:0,zIndex:50}}>
        <div style={{display:"flex",justifyContent:"space-between",marginBottom:12}}>
          <div style={{fontSize:16,fontWeight:800}}>New Daily Report</div>
          <button onClick={onCancel} style={{background:"none",border:"none",color:T.sub,cursor:"pointer",fontSize:13,fontFamily:"inherit"}}>Cancel</button>
        </div>
        <div style={{display:"flex",alignItems:"center"}}>
          {STEPS.map((s,i)=>(
            <div key={i} style={{display:"flex",alignItems:"center",flex:i<STEPS.length-1?1:undefined}}>
              <div style={{display:"flex",flexDirection:"column",alignItems:"center",gap:3}}>
                <div style={{width:26,height:26,borderRadius:"50%",background:i+1<step?T.green:i+1===step?T.orange:T.border,color:i+1<=step?"#09090B":T.muted,display:"flex",alignItems:"center",justifyContent:"center",fontSize:11,fontWeight:800}}>{i+1<step?"✓":i+1}</div>
                <div style={{fontSize:8,color:i+1===step?T.orange:T.muted,fontWeight:i+1===step?700:400,whiteSpace:"nowrap"}}>{s}</div>
              </div>
              {i<STEPS.length-1&&<div style={{flex:1,height:2,background:i+1<step?T.green:T.border,margin:"0 3px",marginBottom:14}}/>}
            </div>
          ))}
        </div>
      </div>

      <div style={{padding:"14px 16px 100px"}}>
        {step===1&&(
          <div>
            <div style={{...card,marginBottom:14,borderLeft:`3px solid ${T.orange}`}}>
              <div style={{fontSize:12,color:T.muted,marginBottom:3}}>Project</div>
              <div style={{fontSize:15,fontWeight:700}}>{project.name}</div>
              {project.afe&&<div style={{fontSize:12,color:T.sub}}>AFE: {project.afe}{project.work_order?" · WO: "+project.work_order:""}</div>}
            </div>
            {[{k:"date",l:"Report Date",t:"date"},{k:"report_no",l:"Report No.",t:"text",ph:"Report #"}].map(({k,l,t,ph})=>(
              <div key={k} style={{marginBottom:12}}><label style={lbl}>{l}</label><input type={t} placeholder={ph||""} value={report[k]||""} onChange={e=>setR(k,e.target.value)} style={inp}/></div>
            ))}
            <div style={{marginBottom:12}}><label style={lbl}>Description of Work Done</label><textarea placeholder="Describe the work performed today…" value={report.description||""} onChange={e=>setR("description",e.target.value)} rows={4} style={{...inp,resize:"vertical",lineHeight:1.5}}/></div>
          </div>
        )}

        {step===2&&(
          <div>
            <div style={{display:"flex",justifyContent:"space-between",marginBottom:12}}><div style={{fontSize:17,fontWeight:800}}>👷 Labor</div>{tot.labor>0&&<div style={{fontSize:16,fontWeight:800,color:T.green}}>${fmt(tot.labor)}</div>}</div>
            {report.labor.map((row,i)=><LaborCard key={row.id} row={row} onChange={r=>updLabor(i,r)} onRemove={()=>delLabor(i)}/>)}
            <DashedAdd label="+ Add Worker" onClick={addLabor}/>
          </div>
        )}

        {step===3&&(
          <div>
            <div style={{display:"flex",justifyContent:"space-between",marginBottom:12}}><div style={{fontSize:17,fontWeight:800}}>🚜 Equipment</div>{tot.equip>0&&<div style={{fontSize:16,fontWeight:800,color:T.green}}>${fmt(tot.equip)}</div>}</div>
            {report.equipment.map((row,i)=><EquipCard key={row.id} row={row} onChange={r=>updEquip(i,r)} onRemove={()=>delEquip(i)}/>)}
            <DashedAdd label="+ Add Equipment" onClick={addEquip}/>
          </div>
        )}

        {step===4&&(
          <div>
            <div style={{fontSize:17,fontWeight:800,marginBottom:12}}>📦 Materials & Misc.</div>
            {report.materials.map((row,i)=><MatCard key={row.id} row={row} onChange={r=>updMat(i,r)} onRemove={()=>delMat(i)}/>)}
            <DashedAdd label="+ Add Material / Item" onClick={addMat}/>
          </div>
        )}

        {step===5&&(
          <div>
            <div style={{fontSize:17,fontWeight:800,marginBottom:12}}>✅ Review & Submit</div>
            <div style={{...card,marginBottom:12}}>
              <div style={{fontSize:11,color:T.orange,fontWeight:700,textTransform:"uppercase",letterSpacing:"1px",marginBottom:10}}>Report Summary</div>
              {[["Project",project.name],["Date",fmtDate(report.date)],["Report No.",report.report_no||"—"],["Workers",report.labor.length],["Equipment",report.equipment.length+" items"],["Materials",report.materials.length+" items"]].map(([l,v])=>(
                <div key={l} style={{display:"flex",justifyContent:"space-between",padding:"7px 0",borderBottom:`1px solid ${T.border}`}}><span style={{fontSize:13,color:T.sub}}>{l}</span><span style={{fontSize:13,fontWeight:600}}>{v}</span></div>
              ))}
              {[["Labor",tot.labor,T.green],["Equipment",tot.equip,T.green],["Materials",tot.mats,T.green]].map(([l,v,c])=>(
                <div key={l} style={{display:"flex",justifyContent:"space-between",padding:"7px 0",borderBottom:`1px solid ${T.border}`}}><span style={{fontSize:13,color:T.sub}}>{l}</span><span style={{fontSize:13,fontWeight:700,color:c}}>${fmt(v)}</span></div>
              ))}
              <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",paddingTop:14}}>
                <span style={{fontSize:16,fontWeight:800}}>Grand Total</span>
                <span style={{fontSize:26,fontWeight:900,color:T.orange,letterSpacing:"-1px"}}>${fmt(tot.grand)}</span>
              </div>
            </div>
            {report.description&&<div style={{...card,marginBottom:12,fontSize:13,color:T.sub,lineHeight:1.6}}><div style={{fontSize:11,color:T.muted,textTransform:"uppercase",letterSpacing:"1px",marginBottom:6}}>Work Done</div>{report.description}</div>}
          </div>
        )}
      </div>

      <div style={{position:"fixed",bottom:0,left:"50%",transform:"translateX(-50%)",width:"100%",maxWidth:480,background:T.bg+"EE",backdropFilter:"blur(12px)",borderTop:`1px solid ${T.border}`,padding:"12px 16px",display:"flex",gap:10}}>
        {step>1&&<button onClick={()=>{setStep(s=>s-1);topRef.current?.scrollIntoView();}} style={{...ghostBtn,flex:1}}>← Back</button>}
        {step<5?<button onClick={()=>{setStep(s=>s+1);topRef.current?.scrollIntoView();}} style={{...primaryBtn,flex:2,borderRadius:12}}>{step===4?"Review →":"Next →"}</button>
          :<button onClick={submit} style={{...primaryBtn,flex:2,borderRadius:12,opacity:saving?0.7:1}}>{saving?"Saving…":"💾 Save Report"}</button>}
      </div>
    </div>
  );
}

function DashedAdd({label,onClick}) {
  return <button onClick={onClick} style={{width:"100%",border:`2px dashed ${T.border}`,background:"transparent",color:T.muted,borderRadius:14,padding:"14px",fontSize:15,cursor:"pointer",fontFamily:"inherit"}}>{label}</button>;
}

function LaborCard({row,onChange,onRemove}) {
  const pos=POSITIONS.find(p=>p.name===row.classification);
  const amt=laborAmt(row);
  const set=(k,v)=>{const u={...row,[k]:v};if(k==="classification"){const p=POSITIONS.find(x=>x.name===v);u.rate=p?p.rate:"";}onChange(u);};
  return(
    <div style={{...card,marginBottom:10,borderLeft:`3px solid ${T.orange}`}}>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8,marginBottom:8}}>
        <div style={{gridColumn:"1/-1"}}><label style={lbl}>Name</label><select value={row.name||""} onChange={e=>set("name",e.target.value)} style={inp}><option value="">— Select —</option>{NAMES.map(n=><option key={n} value={n}>{n}</option>)}</select></div>
        <div style={{gridColumn:"1/-1"}}><label style={lbl}>Classification</label><select value={row.classification||""} onChange={e=>set("classification",e.target.value)} style={inp}><option value="">— Select —</option>{POSITIONS.map(p=><option key={p.name} value={p.name}>{p.name}</option>)}</select></div>
      </div>
      {pos&&!pos.flat&&(
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:8,marginBottom:8}}>
          {[["regHrs","Reg Hrs"],["otHrs","OT Hrs"],["travelHrs","Travel"]].map(([k,l])=>(
            <div key={k}><label style={lbl}>{l}</label><input type="number" min="0" step="0.5" placeholder="0" value={row[k]||""} onChange={e=>set(k,e.target.value)} style={inp}/></div>
          ))}
        </div>
      )}
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",paddingTop:8,borderTop:`1px solid ${T.border}`}}>
        <div style={{fontSize:11,color:T.muted}}>{pos?`$${pos.rate.toFixed(2)}${pos.flat?" flat":"/hr"}`:""}</div>
        <div style={{display:"flex",alignItems:"center",gap:10}}>
          {amt>0&&<div style={{fontSize:16,fontWeight:800,color:T.green}}>${fmt(amt)}</div>}
          <button onClick={onRemove} style={{background:"none",border:"none",color:T.red,cursor:"pointer",fontSize:20,padding:0}}>×</button>
        </div>
      </div>
    </div>
  );
}

function EquipCard({row,onChange,onRemove}) {
  const amt=equipAmt(row);
  const eq=EQUIP_LIST.find(e=>!e.section&&e.name===row.description);
  const set=(k,v)=>{const u={...row,[k]:v};if(k==="description"){const e=EQUIP_LIST.find(x=>!x.section&&x.name===v);u.rate=e?e.rate:"";u.unit=e?e.unit:"";}onChange(u);};
  return(
    <div style={{...card,marginBottom:10,borderLeft:`3px solid ${T.yellow}`}}>
      <div style={{marginBottom:8}}><label style={lbl}>Equipment</label>
        <select value={row.description||""} onChange={e=>set("description",e.target.value)} style={inp}>
          <option value="">— Select Equipment —</option>
          {EQUIP_LIST.map((e,i)=>e.section?<option key={i} disabled>── {e.section} ──</option>:<option key={i} value={e.name}>{e.name}</option>)}
        </select>
      </div>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8,marginBottom:8}}>
        <div><label style={lbl}>Qty</label><input type="number" min="0" placeholder="0" value={row.qty||""} onChange={e=>set("qty",e.target.value)} style={inp}/></div>
        <div><label style={lbl}>{eq?eq.unit:"Hrs/Days"}</label><input type="number" min="0" step="0.5" placeholder="0" value={row.usage||""} onChange={e=>set("usage",e.target.value)} style={inp}/></div>
      </div>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",paddingTop:8,borderTop:`1px solid ${T.border}`}}>
        <div style={{fontSize:11,color:T.muted}}>{eq?`$${eq.rate.toLocaleString()}/${eq.unit}`:""}</div>
        <div style={{display:"flex",alignItems:"center",gap:10}}>
          {amt>0&&<div style={{fontSize:16,fontWeight:800,color:T.green}}>${fmt(amt)}</div>}
          <button onClick={onRemove} style={{background:"none",border:"none",color:T.red,cursor:"pointer",fontSize:20,padding:0}}>×</button>
        </div>
      </div>
    </div>
  );
}

function MatCard({row,onChange,onRemove}) {
  const fileRef=useRef(null);
  const receipts=row.receipts||[];
  async function handleFiles(files){const n=[];for(const f of files){if(!f.type.startsWith("image/"))continue;const src=await compressImg(f,800,0.6);n.push({id:uid(),src});}onChange({...row,receipts:[...receipts,...n]});}
  return(
    <div style={{...card,marginBottom:10,borderLeft:`3px solid ${T.blue}`}}>
      <div style={{display:"grid",gridTemplateColumns:"60px 1fr 90px",gap:8,marginBottom:10}}>
        <div><label style={lbl}>Qty</label><input type="number" min="0" placeholder="0" value={row.qty||""} onChange={e=>onChange({...row,qty:e.target.value})} style={inp}/></div>
        <div><label style={lbl}>Description</label><input type="text" placeholder="Item description" value={row.description||""} onChange={e=>onChange({...row,description:e.target.value})} style={inp}/></div>
        <div><label style={lbl}>Amount</label><input type="number" min="0" placeholder="0.00" value={row.amount||""} onChange={e=>onChange({...row,amount:e.target.value})} style={inp}/></div>
      </div>
      <div style={{borderTop:`1px solid ${T.border}`,paddingTop:10}}>
        <label style={{...lbl,marginBottom:8}}>📎 Receipts</label>
        <div style={{display:"flex",gap:8,flexWrap:"wrap",alignItems:"center"}}>
          {receipts.map(r=>(
            <div key={r.id} style={{position:"relative"}}>
              <img src={r.src} alt="" style={{width:64,height:64,objectFit:"cover",borderRadius:10,border:`2px solid ${T.blue}44`,display:"block"}}/>
              <button onClick={()=>onChange({...row,receipts:receipts.filter(x=>x.id!==r.id)})} style={{position:"absolute",top:-5,right:-5,width:18,height:18,borderRadius:"50%",background:T.red,border:"none",color:"#fff",fontSize:11,cursor:"pointer",display:"flex",alignItems:"center",justifyContent:"center"}}>×</button>
            </div>
          ))}
          <button onClick={()=>fileRef.current?.click()} style={{width:64,height:64,borderRadius:10,border:`2px dashed ${T.blue}44`,background:T.blue+"0A",color:T.blue,cursor:"pointer",display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",fontSize:20,gap:2}}>
            <span>📷</span><span style={{fontSize:9,fontWeight:700}}>ADD</span>
          </button>
          <input ref={fileRef} type="file" accept="image/*" capture="environment" multiple style={{display:"none"}} onChange={e=>{handleFiles(Array.from(e.target.files));e.target.value="";}}/>
        </div>
      </div>
      <div style={{display:"flex",justifyContent:"flex-end",marginTop:8}}>
        <button onClick={onRemove} style={{background:"none",border:"none",color:T.red,cursor:"pointer",fontSize:12,fontWeight:700,fontFamily:"inherit"}}>Remove</button>
      </div>
    </div>
  );
}

/* ═══════════════════════════════════════════════════════════════
   REPORT DETAIL  +  EXCEL EXPORT
═══════════════════════════════════════════════════════════════ */
function ReportDetail({report,project,onBack,onDelete}) {
  const [lb,setLb]=useState(null);
  const tot=reportTotals(report);

  function exportXLSX(){
    const wb=XLSX.utils.book_new();
    const rows=[];
    const blank=()=>Array(11).fill(null);
    const r1=blank();r1[1]="COLONIAL PIPELINE COMPANY";rows.push(r1);
    const r2=blank();r2[1]="DAILY REPORT-WORK PERFORMED BY CONTRACTOR";rows.push(r2);
    const r3=blank();r3[1]="LOCATION";r3[2]="LOCATION";r3[3]="AFE NO.";r3[4]="WORK ORDER\nNUMBER";r3[6]="REPORT DATE";r3[8]="REPORT\nNO.";r3[9]="REGION";rows.push(r3);
    const [yr,mo,dy]=(report.date||"").split("-");
    const r4=blank();r4[1]=project.location||"";r4[3]=project.afe||"";r4[4]=project.work_order||"";r4[6]=`${mo}/${dy}/${yr}`;r4[8]=report.report_no||"";rows.push(r4);
    const r5=blank();r5[1]="CONTRACTOR:";r5[2]="AIME";r5[3]="CONTRACTOR NO:";r5[8]="CONTRACTOR DATE:";r5[9]=report.date||"";rows.push(r5);
    const r6=blank();r6[1]="DESCRIPTION OF WORK DONE:";rows.push(r6);
    const r7=blank();r7[1]=report.description||"";rows.push(r7);
    const r8=blank();r8[1]="LABOR";rows.push(r8);
    const r9=blank();r9[1]="NAME";r9[3]="CLASSIFICATION";r9[5]="REG. HRS.";r9[6]="O.T. HRS.";r9[7]="TRAVEL HRS.";r9[8]="REGULAR RATE";r9[9]="AMOUNT";rows.push(r9);
    const labRows=[...(report.labor||[])];while(labRows.length<14)labRows.push(null);
    labRows.forEach(lr=>{const row=blank();if(lr){const pos=POSITIONS.find(p=>p.name===lr.classification);row[1]=lr.name||"";row[3]=lr.classification||"";if(pos&&!pos.flat){row[5]=parseFloat(lr.regHrs)||0;row[6]=parseFloat(lr.otHrs)||0;row[7]=parseFloat(lr.travelHrs)||0;}row[8]=pos?pos.rate:"";row[9]=laborAmt(lr);}else row[9]=0;rows.push(row);});
    const pdRow=blank();pdRow[3]="Per Diem";pdRow[8]=190;pdRow[9]=0;rows.push(pdRow);
    const tlRow=blank();tlRow[8]="TOTAL LABOR";tlRow[9]=tot.labor;rows.push(tlRow);
    rows.push(blank());
    const ehRow=blank();ehRow[1]="EQUIPMENT";rows.push(ehRow);
    const ecRow=blank();ecRow[1]="DESCRIPTION";ecRow[6]="Quantity";ecRow[7]="Hours/Days";ecRow[8]="RATE";ecRow[9]="AMOUNT";rows.push(ecRow);
    const eqRows=[...(report.equipment||[])];while(eqRows.length<15)eqRows.push(null);
    eqRows.forEach(er=>{const row=blank();if(er){row[1]=er.description||"";row[6]=parseFloat(er.qty)||0;row[7]=parseFloat(er.usage)||0;row[8]=parseFloat(er.rate)||0;row[9]=equipAmt(er);}else row[9]=0;rows.push(row);});
    const teRow=blank();teRow[8]="TOTAL EQUIPMENT";teRow[9]=tot.equip;rows.push(teRow);
    rows.push(blank());
    const rlRow=blank();rlRow[4]="RENTAL EQUIPMENT";rows.push(rlRow);
    const mhRow=blank();mhRow[2]="MATERIAL & MISCELLANEOUS — LIST OF MATERIAL & ATTACH SUPPORTING INVOICES";rows.push(mhRow);
    const mcRow=blank();mcRow[1]="QUANTITY";mcRow[2]="DESCRIPTION";mcRow[5]="AMOUNT";mcRow[6]="QUANTITY";mcRow[7]="DESCRIPTION";mcRow[9]="AMOUNT";rows.push(mcRow);
    const mats=report.materials||[];const pairs=[];for(let i=0;i<mats.length;i+=2)pairs.push([mats[i]||null,mats[i+1]||null]);if(pairs.length===0)pairs.push([null,null]);
    pairs.forEach(([l,r])=>{const ir=blank();if(l){ir[1]=l.qty||"";ir[2]=l.description||"";ir[5]=parseFloat(l.amount)||0;}else ir[5]=0;if(r){ir[6]=r.qty||"";ir[7]=r.description||"";ir[9]=parseFloat(r.amount)||0;}else ir[9]=0;rows.push(ir);const tx=blank();tx[2]="Tax";tx[7]="Tax";rows.push(tx);const to=blank();to[2]="Total";to[5]=parseFloat(l?.amount)||0;to[7]="Total";to[9]=parseFloat(r?.amount)||0;rows.push(to);});
    const trRow=blank();trRow[8]="TOTAL RENTAL EQUIPMENT";trRow[9]=tot.mats;rows.push(trRow);
    const gtRow=blank();gtRow[8]="GRAND TOTAL";gtRow[9]=tot.grand;rows.push(gtRow);
    const sgRow=blank();sgRow[1]="VERIFIED AND\nACCEPTED BY CO. REP";sgRow[3]="DATE";sgRow[4]="CERTIFIED AS CORRECT\nBY CONTRACTOR'S REP";sgRow[8]="REG. OFF";sgRow[9]="ACCT.DEPT";rows.push(sgRow);
    const ws=XLSX.utils.aoa_to_sheet(rows);
    ws["!cols"]=[{wch:5.7},{wch:15},{wch:11.7},{wch:17.1},{wch:12.7},{wch:13.1},{wch:10},{wch:10},{wch:24.9},{wch:23.7},{wch:10.1}];
    const rng=XLSX.utils.decode_range(ws["!ref"]);
    for(let r=0;r<=rng.e.r;r++){const a=XLSX.utils.encode_cell({r,c:9});if(ws[a]&&typeof ws[a].v==="number")ws[a].z='"$"#,##0.00';const b=XLSX.utils.encode_cell({r,c:5});if(ws[b]&&typeof ws[b].v==="number")ws[b].z='"$"#,##0.00';}
    XLSX.utils.book_append_sheet(wb,ws,"Daily Report");
    XLSX.writeFile(wb,`AIME_${project.name.replace(/\s+/g,"_")}_${(report.date||"").replace(/-/g,"")}.xlsx`);
  }

  return (
    <div style={{fontFamily:"inherit",background:T.bg,minHeight:"100vh",padding:16}}>
      <Lightbox src={lb} onClose={()=>setLb(null)}/>
      <button onClick={onBack} style={{...ghostBtn,marginBottom:16}}>← Reports</button>
      <div style={{fontSize:20,fontWeight:900,letterSpacing:"-0.5px",marginBottom:2}}>{fmtDate(report.date)}</div>
      {report.submitted_by&&<div style={{fontSize:12,color:T.muted,marginBottom:14}}>Submitted by {report.submitted_by}</div>}
      {report.description&&<div style={{...card,marginBottom:12,borderLeft:`3px solid ${T.blue}`}}><div style={{fontSize:11,color:T.muted,textTransform:"uppercase",letterSpacing:"1px",marginBottom:6}}>Work Done</div><div style={{fontSize:14,color:T.sub,lineHeight:1.6}}>{report.description}</div></div>}
      {(report.labor||[]).length>0&&(
        <div style={{...card,marginBottom:12}}>
          <div style={{fontSize:12,color:T.orange,fontWeight:700,textTransform:"uppercase",letterSpacing:"1px",marginBottom:10}}>Labor · <span style={{color:T.green}}>${fmt(tot.labor)}</span></div>
          {report.labor.map((r,i)=>(
            <div key={i} style={{display:"flex",justifyContent:"space-between",padding:"8px 0",borderBottom:i<report.labor.length-1?`1px solid ${T.border}`:"none"}}>
              <div><div style={{fontSize:14,fontWeight:600}}>{r.name||"—"}</div><div style={{fontSize:11,color:T.muted}}>{r.classification} · {r.regHrs||0}reg {r.otHrs||0}OT {r.travelHrs||0}tr</div></div>
              <div style={{fontSize:14,fontWeight:800,color:T.green}}>${fmt(laborAmt(r))}</div>
            </div>
          ))}
        </div>
      )}
      {(report.equipment||[]).length>0&&(
        <div style={{...card,marginBottom:12}}>
          <div style={{fontSize:12,color:T.orange,fontWeight:700,textTransform:"uppercase",letterSpacing:"1px",marginBottom:10}}>Equipment · <span style={{color:T.green}}>${fmt(tot.equip)}</span></div>
          {report.equipment.map((r,i)=>(
            <div key={i} style={{display:"flex",justifyContent:"space-between",padding:"8px 0",borderBottom:i<report.equipment.length-1?`1px solid ${T.border}`:"none"}}>
              <div style={{flex:1,paddingRight:10}}><div style={{fontSize:13,fontWeight:600}}>{r.description}</div><div style={{fontSize:11,color:T.muted}}>Qty {r.qty} × {r.usage} {r.unit}</div></div>
              <div style={{fontSize:14,fontWeight:800,color:T.green}}>${fmt(equipAmt(r))}</div>
            </div>
          ))}
        </div>
      )}
      {(report.materials||[]).length>0&&(
        <div style={{...card,marginBottom:12}}>
          <div style={{fontSize:12,color:T.orange,fontWeight:700,textTransform:"uppercase",letterSpacing:"1px",marginBottom:10}}>Materials · <span style={{color:T.green}}>${fmt(tot.mats)}</span></div>
          {report.materials.map((r,i)=>(
            <div key={i} style={{padding:"8px 0",borderBottom:i<report.materials.length-1?`1px solid ${T.border}`:"none"}}>
              <div style={{display:"flex",justifyContent:"space-between",marginBottom:r.receipts?.length>0?8:0}}>
                <span style={{fontSize:13}}>{r.qty?`${r.qty}× `:""}{r.description}</span>
                <span style={{fontSize:13,fontWeight:700,color:T.green}}>${fmt(parseFloat(r.amount)||0)}</span>
              </div>
              {r.receipts?.length>0&&<div style={{display:"flex",gap:6,flexWrap:"wrap"}}>{r.receipts.map(rc=><img key={rc.id} src={rc.src} alt="" onClick={()=>setLb(rc.src)} style={{width:64,height:64,objectFit:"cover",borderRadius:8,cursor:"pointer"}}/>)}</div>}
            </div>
          ))}
        </div>
      )}
      <div style={{...card,background:T.orangeLow,border:`1px solid ${T.orange}44`,marginBottom:12}}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}>
          <div style={{fontSize:15,fontWeight:800}}>Grand Total</div>
          <div style={{fontSize:28,fontWeight:900,color:T.orange,letterSpacing:"-1px"}}>${fmt(tot.grand)}</div>
        </div>
      </div>
      <button onClick={exportXLSX} style={{...primaryBtn,background:T.orangeLow,color:T.orange,border:`1px solid ${T.orange}44`,marginBottom:10,borderRadius:14}}>📥 Export to Excel (.xlsx)</button>
      <button onClick={()=>window.confirm("Delete this report?")&&onDelete(report.id)} style={{...ghostBtn,width:"100%",textAlign:"center",color:T.red,borderColor:T.redLow}}>🗑 Delete Report</button>
    </div>
  );
}

/* ═══════════════════════════════════════════════════════════════
   PM DASHBOARD
═══════════════════════════════════════════════════════════════ */
function PMDashboard({onBack}) {
  const [projects,setProjects] = useState([]);
  const [reports,setReports]   = useState([]);
  const [loading,setLoading]   = useState(true);
  const [err,setErr]           = useState("");

  useEffect(()=>{
    (async()=>{
      try {
        const [projs,reps]=await Promise.all([API.projects.list(),API.reports.all()]);
        setProjects(projs||[]);
        setReports(reps||[]);
      } catch(e){setErr(e.message);}
      setLoading(false);
    })();
  },[]);

  const allTot = reports.reduce((s,r)=>{const t=reportTotals(r);return{labor:s.labor+t.labor,equip:s.equip+t.equip,mats:s.mats+t.mats,grand:s.grand+t.grand};},{labor:0,equip:0,mats:0,grand:0});
  const thisWeek = reports.filter(r=>{const d=new Date(r.date+"T12:00:00");return(Date.now()-d.getTime())/86400000<=7;});
  const uniqueWorkers = [...new Set(reports.flatMap(r=>(r.labor||[]).map(l=>l.name).filter(Boolean)))];

  const projStats = projects.map(p=>{
    const pr=reports.filter(r=>r.project_id===p.id);
    const pt=pr.reduce((s,r)=>{const t=reportTotals(r);return{labor:s.labor+t.labor,equip:s.equip+t.equip,mats:s.mats+t.mats,grand:s.grand+t.grand};},{labor:0,equip:0,mats:0,grand:0});
    return{...p,reportCount:pr.length,...pt};
  });

  const recent=[...reports].slice(0,30);

  return (
    <div style={{fontFamily:"inherit",background:T.bg,minHeight:"100vh"}}>
      <TopBar title="PM Dashboard" sub="All projects · live data" onBack={onBack}/>
      <div style={{padding:"14px 16px 80px"}}>
        <ErrBanner msg={err} onDismiss={()=>setErr("")}/>
        {loading&&<Spinner/>}
        {!loading&&(
          <>
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:20}}>
              {[
                {icon:"💰",label:"Total Billed (All Time)",val:"$"+fmt(allTot.grand),color:T.green},
                {icon:"📋",label:"Reports This Week",val:thisWeek.length,color:T.orange},
                {icon:"🏗️",label:"Active Projects",val:projects.filter(p=>p.status==="active").length,color:T.blue},
                {icon:"👷",label:"Unique Workers",val:uniqueWorkers.length,color:T.yellow},
              ].map(k=>(
                <div key={k.label} style={card}>
                  <div style={{fontSize:22,marginBottom:6}}>{k.icon}</div>
                  <div style={{fontSize:22,fontWeight:900,color:k.color,letterSpacing:"-0.5px"}}>{k.val}</div>
                  <div style={{fontSize:11,color:T.muted,marginTop:3}}>{k.label}</div>
                </div>
              ))}
            </div>

            <div style={{fontSize:12,fontWeight:700,color:T.muted,textTransform:"uppercase",letterSpacing:"1px",marginBottom:12}}>Projects Breakdown</div>
            {projStats.map(p=>(
              <div key={p.id} style={{...card,marginBottom:10}}>
                <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:10}}>
                  <div><div style={{fontSize:15,fontWeight:700}}>{p.name}</div><div style={{fontSize:11,color:T.muted}}>{p.reportCount} reports · {p.location||"No location"}</div></div>
                  <div style={{fontSize:18,fontWeight:900,color:T.green}}>${p.grand>=1000?(p.grand/1000).toFixed(1)+"k":fmt(p.grand)}</div>
                </div>
                <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:6}}>
                  {[["Labor",p.labor,T.orange],["Equip",p.equip,T.yellow],["Mats",p.mats,T.blue]].map(([l,v,c])=>(
                    <div key={l} style={{background:T.surface||"#111",borderRadius:8,padding:"8px",textAlign:"center"}}>
                      <div style={{fontSize:13,fontWeight:700,color:c}}>${v>=1000?(v/1000).toFixed(1)+"k":fmt(v)}</div>
                      <div style={{fontSize:9,color:T.muted,textTransform:"uppercase",letterSpacing:"0.5px"}}>{l}</div>
                    </div>
                  ))}
                </div>
              </div>
            ))}

            <div style={{fontSize:12,fontWeight:700,color:T.muted,textTransform:"uppercase",letterSpacing:"1px",margin:"20px 0 12px"}}>Recent Reports — All Projects</div>
            {recent.map(r=>{
              const t=reportTotals(r);
              const pname=r.projects?.name||"Unknown Project";
              return(
                <div key={r.id} style={{...card,marginBottom:8,display:"flex",justifyContent:"space-between",alignItems:"center"}}>
                  <div><div style={{fontSize:13,fontWeight:700}}>{fmtShort(r.date)}</div><div style={{fontSize:11,color:T.muted}}>{pname} · {r.submitted_by||"Unknown"}</div></div>
                  <div style={{fontSize:15,fontWeight:800,color:T.green}}>${fmt(t.grand)}</div>
                </div>
              );
            })}
            {recent.length===0&&<div style={{textAlign:"center",padding:"32px 0",color:T.muted}}>No reports submitted yet.</div>}
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
  const [user,setUser]       = useState(null);
  const [projects,setProjects] = useState([]);
  const [loading,setLoading] = useState(false);
  const [err,setErr]         = useState("");
  const [screen,setScreen]   = useState("projects"); // projects | newProject | projectDetail | pmDashboard
  const [activeProject,setActiveProject] = useState(null);

  useEffect(()=>{
    const link=document.createElement("link");
    link.rel="stylesheet";
    link.href="https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600;700;800;900&display=swap";
    document.head.appendChild(link);
  },[]);

  async function loadProjects() {
    setLoading(true); setErr("");
    try {
      const projs = await API.projects.list();
      // For each project, load report count + total billed
      const enriched = await Promise.all((projs||[]).map(async p=>{
        try {
          const reps = await API.reports.forProject(p.id);
          const totalBilled = (reps||[]).reduce((s,r)=>s+reportTotals(r).grand,0);
          return {...p,_reportCount:(reps||[]).length,_totalBilled:totalBilled};
        } catch { return {...p,_reportCount:0,_totalBilled:0}; }
      }));
      setProjects(enriched);
    } catch(e) { setErr(e.message); }
    setLoading(false);
  }

  useEffect(()=>{ if(user) loadProjects(); },[user]);

  async function handleNewProject(data) {
    try {
      const [created]=await API.projects.create({...data,created_by:user.name});
      await loadProjects();
      setScreen("projects");
    } catch(e) { setErr(e.message); }
  }

  function handleProjectUpdated(updated) {
    setActiveProject(updated);
    loadProjects();
  }

  if(!user) return (
    <div style={{maxWidth:480,margin:"0 auto",fontFamily:"'DM Sans',system-ui,sans-serif"}}>
      <LoginScreen onLogin={u=>{setUser(u);}}/>
    </div>
  );

  return (
    <div style={{maxWidth:480,margin:"0 auto",fontFamily:"'DM Sans',system-ui,sans-serif",background:T.bg,minHeight:"100vh"}}>
      {err&&<div style={{position:"fixed",top:0,left:"50%",transform:"translateX(-50%)",width:"100%",maxWidth:480,zIndex:200,padding:"12px 16px",background:T.redLow,borderBottom:`1px solid ${T.red}44`}}><ErrBanner msg={err} onDismiss={()=>setErr("")}/></div>}

      {screen==="pmDashboard"&&<PMDashboard onBack={()=>setScreen("projects")}/>}

      {screen==="newProject"&&<ProjectForm onSave={handleNewProject} onCancel={()=>setScreen("projects")}/>}

      {screen==="projectDetail"&&activeProject&&(
        <ProjectDetail
          project={activeProject}
          user={user}
          onBack={()=>{ setScreen("projects"); loadProjects(); }}
          onProjectUpdated={handleProjectUpdated}
        />
      )}

      {screen==="projects"&&(
        <ProjectsHome
          user={user}
          projects={projects}
          loading={loading}
          onSelect={p=>{setActiveProject(p);setScreen("projectDetail");}}
          onNew={()=>setScreen("newProject")}
          onLogout={()=>{setUser(null);setProjects([]);setScreen("projects");}}
          onDash={()=>setScreen("pmDashboard")}
        />
      )}
    </div>
  );
}
