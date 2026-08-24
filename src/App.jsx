import React, { useState, useEffect, useRef } from "react";
import * as XLSX from "xlsx";
const LOGIN_BG="data:image/webp;base64,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";

class ErrorBoundary extends React.Component {
  constructor(props){super(props);this.state={error:null,info:null,copied:false};}
  static getDerivedStateFromError(error){return{error};}
  componentDidCatch(error,info){console.error("App Error:",error,info);this.setState({info});}
  render(){
    if(this.state.error){
      return(
        <div style={{background:"#0D0D0F",minHeight:"100vh",display:"flex",alignItems:"center",justifyContent:"center",fontFamily:"system-ui,sans-serif",padding:20}}>
          <div style={{background:"#1A1A20",border:"1px solid #EF4444",borderRadius:16,padding:24,maxWidth:480,width:"100%"}}>
            <div style={{fontSize:24,marginBottom:8}}>⚠️</div>
            <div style={{fontSize:16,fontWeight:700,color:"#FC8181",marginBottom:8}}>App Error</div>
            <div style={{fontSize:12,color:"#A1A1AA",marginBottom:16,wordBreak:"break-all",background:"#0C0C0F",padding:12,borderRadius:8,lineHeight:1.6}}>
              {this.state.error?.message||String(this.state.error)}
            </div>
            {this.state.info?.componentStack&&(
              <div style={{marginBottom:16}}>
                <div style={{fontSize:10,color:"#71717A",textTransform:"uppercase",letterSpacing:"1px",marginBottom:5}}>Where it happened</div>
                <div style={{fontSize:11,color:"#A1A1AA",background:"#0C0C0F",padding:10,borderRadius:8,lineHeight:1.7,maxHeight:150,overflow:"auto",whiteSpace:"pre-wrap",fontFamily:"ui-monospace,monospace"}}>
                  {this.state.info.componentStack.trim().split("\n").slice(0,8).join("\n")}
                </div>
              </div>
            )}
            <button onClick={()=>{
              const txt=[
                this.state.error?.message||String(this.state.error),
                "",
                this.state.error?.stack||"",
                "",
                this.state.info?.componentStack||"",
              ].join("\n");
              const done=()=>{this.setState({copied:true});setTimeout(()=>this.setState({copied:false}),2500);};
              if(navigator.clipboard) navigator.clipboard.writeText(txt).then(done).catch(done);
              else done();
            }} style={{background:"transparent",color:"#60A5FA",border:"1px solid #26262E",borderRadius:10,padding:"9px 16px",fontSize:12,fontWeight:700,cursor:"pointer",fontFamily:"inherit",width:"100%",marginBottom:10}}>
              {this.state.copied?"✓ Copied — paste it to your developer":"📋 Copy Full Error Details"}
            </button>
            <button onClick={()=>{this.setState({error:null});window.location.reload();}}
              style={{background:"#60A5FA",color:"#000",border:"none",borderRadius:10,padding:"10px 20px",fontSize:14,fontWeight:700,cursor:"pointer",fontFamily:"inherit",width:"100%"}}>
              🔄 Reload App
            </button>
          </div>
        </div>
      );
    }
    return this.props.children;
  }
}

const QUEUE_KEY   = 'aime_pending_queue';
const DRAFT_PREFIX = 'aime_draft_';

function getQueue(){
  try{return JSON.parse(localStorage.getItem(QUEUE_KEY)||'[]');}
  catch{return [];}
}
function addToQueue(item){
  const q=getQueue();
  q.push({...item,qid:Math.random().toString(36).slice(2),queued_at:new Date().toISOString()});
  localStorage.setItem(QUEUE_KEY,JSON.stringify(q));
}
function removeFromQueue(qid){
  localStorage.setItem(QUEUE_KEY,JSON.stringify(getQueue().filter(i=>i.qid!==qid)));
}
function saveDraft(key,data){
  try{localStorage.setItem(DRAFT_PREFIX+key,JSON.stringify({data,saved_at:new Date().toISOString()}));}
  catch(e){console.warn('Draft save failed:',e);}
}
function loadDraft(key){
  try{const s=localStorage.getItem(DRAFT_PREFIX+key);return s?JSON.parse(s):null;}
  catch{return null;}
}
function clearDraft(key){
  try{localStorage.removeItem(DRAFT_PREFIX+key);}catch{}
}

const SUPA_URL = "https://uicmfyudiullulbbwzmh.supabase.co";
const SUPA_KEY = "sb_publishable_9h9AyvXpkp9glLxDVWRuGw_1eKVS7sE";
async function supa(path,{method="GET",body,prefer}={}){
  const r=await fetch(`${SUPA_URL}/rest/v1${path}`,{method,headers:{"apikey":SUPA_KEY,"Authorization":`Bearer ${SUPA_KEY}`,"Content-Type":"application/json",...(prefer?{"Prefer":prefer}:{}),...(body!==undefined?{body:JSON.stringify(body)}:{}),...(body!==undefined?{}:{})},body:body!==undefined?JSON.stringify(body):undefined});
  if(!r.ok)throw new Error(await r.text()||`HTTP ${r.status}`);
  const t=await r.text();return t?JSON.parse(t):null;
}
async function sb(path,opts={}){
  const{method="GET",body,prefer}=opts;
  const headers={"apikey":SUPA_KEY,"Authorization":`Bearer ${SUPA_KEY}`,"Content-Type":"application/json"};
  if(prefer)headers["Prefer"]=prefer;
  const res=await fetch(`${SUPA_URL}/rest/v1${path}`,{method,headers,...(body!==undefined?{body:JSON.stringify(body)}:{})});
  if(!res.ok)throw new Error(await res.text()||`HTTP ${res.status}`);
  const t=await res.text();return t?JSON.parse(t):null;
}

/* Supabase Storage over plain REST — the app has no supabase-js client. */
async function storageUpload(bucket,path,file,contentType){
  const res=await fetch(`${SUPA_URL}/storage/v1/object/${bucket}/${path}`,{
    method:"POST",
    headers:{apikey:SUPA_KEY,Authorization:`Bearer ${SUPA_KEY}`,"x-upsert":"true",
      ...(contentType?{"Content-Type":contentType}:{})},
    body:file,
  });
  if(!res.ok)throw new Error((await res.text())||`Upload failed (${res.status})`);
  return res.json().catch(()=>({}));
}
function storagePublicUrl(bucket,path){
  return `${SUPA_URL}/storage/v1/object/public/${bucket}/${path}`;
}
async function storageRemove(bucket,path){
  const res=await fetch(`${SUPA_URL}/storage/v1/object/${bucket}/${path}`,{
    method:"DELETE",
    headers:{apikey:SUPA_KEY,Authorization:`Bearer ${SUPA_KEY}`},
  });
  if(!res.ok&&res.status!==404)throw new Error((await res.text())||`Delete failed (${res.status})`);
}

// Wider shell on desktop for list and dashboard screens only. maxWidth is a cap,
// so phones (viewport under 480) are completely unaffected by these values.
// Data-entry screens stay narrow — a form field stretched across a 27" monitor
// is worse, not better.
// Estimating is under construction — restricted to this account regardless of
// role. Clear ESTIMATING_OWNER (set it to "") to open it up to the normal
// `estimating` permission again.
const ESTIMATING_OWNER="Doug Friedel";
const canEstimate=(u)=>!!u&&can(u,"estimating")&&(!ESTIMATING_OWNER||u.name===ESTIMATING_OWNER);

const WIDE_SCREENS=new Set(["pmDashboard","timeCards","crewDirectory","userManagement","estimating","jobs"]);
const shellMax=(screen)=>WIDE_SCREENS.has(screen)?1180:480;

/* ── Session ────────────────────────────────────────────────
   Keeps you signed in across refreshes. Stores only the profile name and a
   timestamp — never the PIN — and re-reads the profile from the database on
   restore, so a role change or deactivation takes effect immediately. */
const SESSION_KEY="aime_session";
const SESSION_MAX_AGE=1000*60*60*24*7; // 7 days

function saveSession(profile){
  try{localStorage.setItem(SESSION_KEY,JSON.stringify({name:profile.name,at:Date.now()}));}catch{}
}
function clearSession(){
  try{localStorage.removeItem(SESSION_KEY);}catch{}
}
async function restoreSession(){
  let s;
  try{s=JSON.parse(localStorage.getItem(SESSION_KEY)||"null");}catch{return null;}
  if(!s||!s.name)return null;
  if(Date.now()-(s.at||0)>SESSION_MAX_AGE){clearSession();return null;}
  try{
    const rows=await API.userProfiles.getByName(s.name);
    const p=rows&&rows.length>0?rows[0]:null;
    if(!p||p.active===false){clearSession();return null;}
    saveSession(p);            // slide the window forward on each use
    return p;
  }catch{ return null; }       // offline: don't wipe the session, just stay logged out this load
}

const API={
  drawings:{
    forProject:(pid)=>sb(`/drawings?project_id=eq.${pid}&select=*&order=created_at.desc`),
    create:(d)=>sb('/drawings',{method:'POST',body:d,prefer:'return=representation'}),
    update:(id,d)=>sb(`/drawings?id=eq.${id}`,{method:'PATCH',body:d,prefer:'return=representation'}),
    remove:(id)=>sb(`/drawings?id=eq.${id}`,{method:'DELETE'}),
  },
  markups:{
    forDrawing:(did)=>sb(`/drawing_markups?drawing_id=eq.${did}&select=*`),
    create:(d)=>sb('/drawing_markups',{method:'POST',body:d,prefer:'return=representation'}),
    update:(id,d)=>sb(`/drawing_markups?id=eq.${id}`,{method:'PATCH',body:d,prefer:'return=representation'}),
  },
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
    update:(id,d)=>sb(`/daily_reports?id=eq.${id}`,{method:"PATCH",body:d,prefer:"return=representation"}),count:(id)=>sb(`/daily_reports?id=eq.${id}&select=id`),
    remove:(id)=>sb(`/daily_reports?id=eq.${id}`,{method:"DELETE"}),
  },
  tmTickets:{
    forProject:(pid)=>sb(`/tm_tickets?project_id=eq.${pid}&order=created_at.desc`),
    byId:(id)=>sb(`/tm_tickets?id=eq.${id}&limit=1`),
    create:(d)=>sb('/tm_tickets',{method:'POST',body:d,prefer:'return=representation'}),
    update:(id,d)=>sb(`/tm_tickets?id=eq.${id}`,{method:'PATCH',body:d,prefer:'return=representation'}),
    remove:(id)=>sb(`/tm_tickets?id=eq.${id}`,{method:'DELETE'}),
  },
  safety:   {forProject:(pid)=>sb(`/safety_logs?project_id=eq.${pid}&order=created_at.desc`),create:(d)=>sb("/safety_logs",{method:"POST",body:d,prefer:"return=representation"}),remove:(id)=>sb(`/safety_logs?id=eq.${id}`,{method:"DELETE"})},
  photos:   {forProject:(pid)=>sb(`/project_photos?project_id=eq.${pid}&order=created_at.desc`),create:(d)=>sb("/project_photos",{method:"POST",body:d,prefer:"return=representation"}),remove:(id)=>sb(`/project_photos?id=eq.${id}`,{method:"DELETE"})},
  timeCards:{forProject:(pid)=>sb(`/time_cards?project_id=eq.${pid}&order=date.desc,created_at.desc`),all:()=>sb("/time_cards?order=date.desc,created_at.desc&limit=500"),byDate:(date)=>sb(`/time_cards?date=eq.${date}&order=worker_name.asc`),byRange:(from,to)=>sb(`/time_cards?date=gte.${from}&date=lte.${to}&order=date.desc,worker_name.asc`),find:(name,date,pid)=>sb(`/time_cards?worker_name=eq.${encodeURIComponent(name)}&date=eq.${date}&project_id=eq.${pid}&limit=1`),create:(d)=>sb("/time_cards",{method:"POST",body:d,prefer:"return=representation"}),update:(id,d)=>sb(`/time_cards?id=eq.${id}`,{method:"PATCH",body:d,prefer:"return=representation"}),remove:(id)=>sb(`/time_cards?id=eq.${id}`,{method:"DELETE"})},
  weather:  {forProject:(pid)=>sb(`/weather_logs?project_id=eq.${pid}&order=date.desc&limit=14`),upsert:(d)=>sb("/weather_logs",{method:"POST",body:d,prefer:"return=representation,resolution=merge-duplicates"}),remove:(id)=>sb(`/weather_logs?id=eq.${id}`,{method:"DELETE"})},
  equipment:{forProject:(pid)=>sb(`/equipment_on_site?project_id=eq.${pid}&order=date.desc,created_at.desc`),create:(d)=>sb("/equipment_on_site",{method:"POST",body:d,prefer:"return=representation"}),remove:(id)=>sb(`/equipment_on_site?id=eq.${id}`,{method:"DELETE"})},
  subs:     {forProject:(pid)=>sb(`/subcontractors?project_id=eq.${pid}&order=date.desc,created_at.desc`),create:(d)=>sb("/subcontractors",{method:"POST",body:d,prefer:"return=representation"}),remove:(id)=>sb(`/subcontractors?id=eq.${id}`,{method:"DELETE"})},
  mfg:{
    jobs:{list:()=>sb('/mfg_jobs?order=created_at.desc'),create:(d)=>sb('/mfg_jobs',{method:'POST',body:d,prefer:'return=representation'}),update:(id,d)=>sb(`/mfg_jobs?id=eq.${id}`,{method:'PATCH',body:d}),remove:(id)=>sb(`/mfg_jobs?id=eq.${id}`,{method:'DELETE'})},
    parts:{forJob:(jid)=>sb(`/mfg_parts?job_id=eq.${jid}&order=part_number.asc`),create:(d)=>sb('/mfg_parts',{method:'POST',body:d,prefer:'return=representation'}),update:(id,d)=>sb(`/mfg_parts?id=eq.${id}`,{method:'PATCH',body:d}),remove:(id)=>sb(`/mfg_parts?id=eq.${id}`,{method:'DELETE'})},
    bom:{forPart:(pid)=>sb(`/mfg_bom?part_id=eq.${pid}`),create:(d)=>sb('/mfg_bom',{method:'POST',body:d,prefer:'return=representation'}),update:(id,d)=>sb(`/mfg_bom?id=eq.${id}`,{method:'PATCH',body:d}),remove:(id)=>sb(`/mfg_bom?id=eq.${id}`,{method:'DELETE'})},
    receipts:{forPart:(pid)=>sb(`/mfg_receipts?part_id=eq.${pid}&order=created_at.desc`),create:(d)=>sb('/mfg_receipts',{method:'POST',body:d,prefer:'return=representation'})},
    travelers:{forPart:(pid)=>sb(`/mfg_travelers?part_id=eq.${pid}&limit=1`),upsert:(d)=>sb('/mfg_travelers',{method:'POST',body:d,prefer:'return=representation',resolution:'merge-duplicates'})},
    stageLog:{forPart:(pid)=>sb(`/mfg_stage_log?part_id=eq.${pid}&order=created_at.desc`),create:(d)=>sb('/mfg_stage_log',{method:'POST',body:d,prefer:'return=representation'}),remove:(id)=>sb(`/mfg_stage_log?id=eq.${id}`,{method:'DELETE'})},
    assemblyLog:{forPart:(pid)=>sb(`/mfg_assembly_log?part_id=eq.${pid}&order=completion_date.desc`),forJob:(jid)=>sb(`/mfg_assembly_log?job_id=eq.${jid}&order=completion_date.desc`),create:(d)=>sb('/mfg_assembly_log',{method:'POST',body:d,prefer:'return=representation'})},
    shippingLog:{forPart:(pid)=>sb(`/mfg_shipping_log?part_id=eq.${pid}&order=ship_date.desc`),forJob:(jid)=>sb(`/mfg_shipping_log?job_id=eq.${jid}&order=ship_date.desc`),create:(d)=>sb('/mfg_shipping_log',{method:'POST',body:d,prefer:'return=representation'})},
    packingSlips:{forJob:(jid)=>sb(`/mfg_packing_slips?job_id=eq.${jid}&order=created_at.desc`),create:(d)=>sb('/mfg_packing_slips',{method:'POST',body:d,prefer:'return=representation'}),update:(id,d)=>sb(`/mfg_packing_slips?id=eq.${id}`,{method:'PATCH',body:d,prefer:'return=representation'}),remove:(id)=>sb(`/mfg_packing_slips?id=eq.${id}`,{method:'DELETE'})},
    labor:{forJob:(jid)=>sb(`/mfg_labor?job_id=eq.${jid}&order=work_date.desc`),forPart:(pid)=>sb(`/mfg_labor?part_id=eq.${pid}&order=work_date.desc`),create:(d)=>sb('/mfg_labor',{method:'POST',body:d,prefer:'return=representation'}),remove:(id)=>sb(`/mfg_labor?id=eq.${id}`,{method:'DELETE'})},
    ncr:{forJob:(jid)=>sb(`/mfg_ncr?job_id=eq.${jid}&order=created_at.desc`),forPart:(pid)=>sb(`/mfg_ncr?part_id=eq.${pid}&order=created_at.desc`),create:(d)=>sb('/mfg_ncr',{method:'POST',body:d,prefer:'return=representation'}),update:(id,d)=>sb(`/mfg_ncr?id=eq.${id}`,{method:'PATCH',body:d})},
  },
  docFolders:{forProject:(pid)=>sb(`/document_folders?project_id=eq.${pid}&order=name.asc`),create:(d)=>sb("/document_folders",{method:"POST",body:d,prefer:"return=representation"}),update:(id,d)=>sb(`/document_folders?id=eq.${id}`,{method:"PATCH",body:d}),remove:(id)=>sb(`/document_folders?id=eq.${id}`,{method:"DELETE"})},
  docs:     {forProject:(pid)=>sb(`/documents?project_id=eq.${pid}&order=created_at.desc`),create:(d)=>sb("/documents",{method:"POST",body:d,prefer:"return=representation"}),update:(id,d)=>sb(`/documents?id=eq.${id}`,{method:"PATCH",body:d,prefer:"return=representation"}),remove:(id)=>sb(`/documents?id=eq.${id}`,{method:"DELETE"})},
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
  catalog:{
    list:(div,cat,search)=>{
      let url="/cost_catalog?active=eq.true&order=name.asc&limit=200";
      if(div) url+=`&division=eq.${encodeURIComponent(div)}`;
      if(cat) url+=`&category=eq.${encodeURIComponent(cat)}`;
      if(search) url+=`&name=ilike.${encodeURIComponent("*"+search+"*")}`;
      return sb(url);
    },
    categories:(div)=>sb(`/cost_catalog?active=eq.true&division=eq.${encodeURIComponent(div)}&select=category&order=category.asc`),
    update:(id,d)=>sb(`/cost_catalog?id=eq.${id}`,{method:"PATCH",body:d,prefer:"return=representation"}),
    create:(d)=>sb("/cost_catalog",{method:"POST",body:d,prefer:"return=representation"}),
    remove:(id)=>sb(`/cost_catalog?id=eq.${id}`,{method:"PATCH",body:{active:false}}),
  },
  estimates:{
    list:()=>sb("/estimates?order=created_at.desc"),
    create:(d)=>sb("/estimates",{method:"POST",body:d,prefer:"return=representation"}),
    update:(id,d)=>sb(`/estimates?id=eq.${id}`,{method:"PATCH",body:d,prefer:"return=representation"}),
    remove:(id)=>sb(`/estimates?id=eq.${id}`,{method:"DELETE"}),
    items:(eid)=>sb(`/estimate_items?estimate_id=eq.${eid}&order=sort_order.asc,category.asc`),
    addItem:(d)=>sb("/estimate_items",{method:"POST",body:d,prefer:"return=representation"}),
    updateItem:(id,d)=>sb(`/estimate_items?id=eq.${id}`,{method:"PATCH",body:d}),
    removeItem:(id)=>sb(`/estimate_items?id=eq.${id}`,{method:"DELETE"}),
  },
  changeOrders:{
    forProject:(pid)=>sb(`/change_orders?project_id=eq.${pid}&order=date_submitted.asc`),
    create:(d)=>sb("/change_orders",{method:"POST",body:d,prefer:"return=representation"}),
    update:(id,d)=>sb(`/change_orders?id=eq.${id}`,{method:"PATCH",body:d,prefer:"return=representation"}),
    remove:(id)=>sb(`/change_orders?id=eq.${id}`,{method:"DELETE"}),
  },
  rfis:{
    forProject:(pid)=>sb(`/rfis?project_id=eq.${pid}&order=date_submitted.asc`),
    create:(d)=>sb("/rfis",{method:"POST",body:d,prefer:"return=representation"}),
    update:(id,d)=>sb(`/rfis?id=eq.${id}`,{method:"PATCH",body:d,prefer:"return=representation"}),
    remove:(id)=>sb(`/rfis?id=eq.${id}`,{method:"DELETE"}),
  },
};

const T={bg:"#0D0D0F",surface:"#141418",card:"#1A1A20",border:"#26262E",orange:"#60A5FA",orangeLow:"#60A5FA14",orangeMid:"#60A5FA30",green:"#34D399",greenLow:"#34D39914",red:"#FC8181",redLow:"#FC818114",yellow:"#FBBF24",yellowLow:"#FBBF2414",blue:"#60A5FA",blueLow:"#60A5FA14",purple:"#A78BFA",purpleLow:"#A78BFA14",teal:"#2DD4BF",text:"#F0F4FF",sub:"#C8D4F0",muted:"#7080A0"};
const inp={width:"100%",boxSizing:"border-box",background:"#0C0C0F",border:`1px solid ${T.border}`,borderRadius:12,color:T.text,fontSize:15,padding:"13px 14px",outline:"none",fontFamily:"inherit",appearance:"none",WebkitAppearance:"none"};
const inpSel={...inp,color:T.orange,background:T.card,colorScheme:"dark"};
const lbl={display:"block",fontSize:11,fontWeight:700,color:"#D4D4D8",letterSpacing:"1px",textTransform:"uppercase",marginBottom:6};
const cardS={background:T.card,border:`1px solid ${T.border}`,borderRadius:16,padding:"16px"};
const pill=(c)=>({display:"inline-flex",alignItems:"center",background:c+"20",color:c,borderRadius:20,padding:"3px 10px",fontSize:11,fontWeight:700});
const primBtn={width:"100%",background:T.orange,color:"#FFFFFF",border:"none",borderRadius:14,padding:"16px",fontSize:16,fontWeight:800,cursor:"pointer",fontFamily:"inherit",display:"flex",alignItems:"center",justifyContent:"center",gap:8};
const ghostBtn={background:"transparent",border:`1px solid ${T.border}`,borderRadius:12,padding:"12px 16px",color:T.sub,fontSize:14,cursor:"pointer",fontFamily:"inherit",fontWeight:600};
const dangerBtn={background:T.redLow,border:`1px solid ${T.red}30`,borderRadius:12,padding:"12px 16px",color:T.red,fontSize:14,cursor:"pointer",fontFamily:"inherit",fontWeight:600,width:"100%",textAlign:"center"};

const POSITIONS_PIPELINE=[{name:"Project Manager",rate:64.50},{name:"Foreman",rate:63.25},{name:"Technician",rate:60.75},{name:"Inspector",rate:53.75},{name:"Certified Welder",rate:60.75},{name:"Fitter",rate:58.50},{name:"Mechanic",rate:58.50},{name:"Operator",rate:58.50},{name:"Truck Driver",rate:58.50},{name:"Helper (Welder)",rate:57.25},{name:"Laborer",rate:51.00},{name:"Foreman (Elect)",rate:82.25},{name:"Electrician",rate:82.25},{name:"Helper (Elect)",rate:45.50},{name:"Per Diem",rate:190.00,flat:true}];
const NAMES=["Alan Fairbrother","Alan Robinson","Doug Friedel","Jaden Pugh","Brandon Milano","Charles Acree","Charles Dovel","Chris Utz","Christopher Dean","Chuck Dean","Clay Lau","Connor Kestner","Morgan Schramek","Eric Bowens","Eric Shumate","Jackson Fama","Howard Lau","Jeff White","Jessica Vance","John Baier","John P. Cosner Jr.","Jordan Gorwell","Joseph Lau","Josh Gladhill","Kevin Gabrish","Kurt Batterton","Leo Velez","Edgrado Ruiz","Mark Hamilton","Alejandro Figueroa","Matthew Linton","Mike Gamble","Mike Gamble III","Mike Seiler","Pat Gorman","Paul Howard","Rich Raborg","Robert Neslein","Roland Long","Shane Hower","Steve Kestner","Tom Hatfield","Troy Strother","Tyrone Davis","Walter Chicas-Luna","Will Wychulis","Wyatt Gill"].sort();
const EQUIP_LIST_PIPELINE=[{section:"Trucks & Trailers"},{name:"Truck - 1 Ton",rate:21.50,unit:"Hours"},{name:"Truck - 3/4 Ton w/ Snow Plow",rate:350,unit:"Days"},{name:"Truck - 1/2 Ton",rate:18.50,unit:"Hours"},{name:"Truck - Boom (20-29 Ton)",rate:65,unit:"Hours"},{name:"Truck - Bucket",rate:45,unit:"Hours"},{name:"Truck - Dump Truck (3 Axle)",rate:35,unit:"Hours"},{name:"Truck - Haul Truck - No Trailer",rate:70,unit:"Hours"},{name:"Truck - Tru-Vac",rate:13500,unit:"Month"},{name:"Truck - Welding Rig",rate:35,unit:"Hours"},{name:"Trailer - Electrical - Colonial",rate:147,unit:"Month"},{name:"Trailer - Lowboy - 2 Axle",rate:28,unit:"Hours"},{name:"Trailer - Tag Along",rate:50,unit:"Days"},{name:"Trailer - Tool Trailer - 18-25ft",rate:175,unit:"Days"},{name:"Trailer - Tool Trailer - 26-40ft",rate:200,unit:"Days"},{section:"Earthmoving & ROW"},{name:"ATV - 4 Wheel",rate:125,unit:"Days"},{name:"Backhoe Loader - 80-105 HP",rate:62.45,unit:"Hours"},{name:"Excavator - Mini - 2-8K LB",rate:299,unit:"Days"},{name:"Excavator - Mini - 9K LB",rate:335,unit:"Days"},{name:"Excavator - Mini - 12-16K LB",rate:475,unit:"Days"},{name:"Excavator - Small - 21-29K LB",rate:565,unit:"Days"},{name:"Excavator - Small - 30-33K LB",rate:632,unit:"Days"},{name:"Excavator - Medium - 48-55K LB",rate:852,unit:"Days"},{name:"Excavator - Large - 80-89K LB",rate:1050,unit:"Days"},{name:"Excavator - Large - 90-119K LB",rate:1350,unit:"Days"},{name:"Skidsteer Loader - 70-80 HP",rate:440,unit:"Days"},{name:"Skidsteer Loader - 81-100 HP",rate:475,unit:"Days"},{name:"Tractor - 50 HP 4x4 w/ Bush Hog",rate:36.50,unit:"Hours"},{name:"Mower - Riding/Zero Turn",rate:175,unit:"Days"},{section:"Air, Compressors & Blast"},{name:"Air Compressor - 185 CFM",rate:195,unit:"Days"},{name:"Air Compressor - 375 CFM",rate:275,unit:"Days"},{name:"Air Impact Wrench - 1in",rate:50,unit:"Days"},{name:"Air Spade / Knife",rate:55,unit:"Days"},{name:"Blast Rig - 4 Bag Pot w/ 185 CFM AC",rate:55.50,unit:"Hours"},{name:"Blast Rig - 1 Pot w/ 375 CFM AC",rate:500,unit:"Days"},{section:"Testing & Misc. Tools"},{name:"Holiday Detector / Pipe Jeep",rate:72,unit:"Days"},{name:"Hydraulic Torque",rate:200,unit:"Days"},{name:"Hydro Test Pump",rate:60,unit:"Days"},{name:"Hydrotest - High Pressure",rate:3800,unit:"Days"},{name:"Jack Hammer",rate:72,unit:"Days"},{name:"LEL/Gas Monitor - 4 Gas",rate:50,unit:"Days"},{name:"Line Locator",rate:50,unit:"Days"},{name:"HEPA Vacuum",rate:100,unit:"Days"},{name:"Torque Wrench w/Sockets Hyd/Pneu",rate:195,unit:"Days"},{name:"Pipe Beveling Machine 16-22in",rate:100,unit:"Days"}];

const POSITIONS_MECHANICAL=[
  {name:"Project Manager",rate:93.00},
  {name:"Foreman",rate:72.50},
  {name:"Inspector",rate:140.00},
  {name:"Certified Welder",rate:69.50},
  {name:"Mechanic",rate:69.50},
  {name:"Operator",rate:69.50},
  {name:"Truck Driver",rate:69.50},
  {name:"Dock Watch",rate:70.00},
  {name:"Per Diem",rate:190.00,flat:true},
];

const EQUIP_LIST_MECHANICAL=[
  {section:"Trucks & Trailers"},
  {name:"Truck - Tool Truck",rate:40,unit:"Hours"},
  {name:"Truck - 3/4 Ton w/ Snow Plow",rate:100,unit:"Hours"},
  {name:"Truck - Boom (20-29 Ton)",rate:98,unit:"Hours"},
  {name:"Truck - Dump Truck",rate:40,unit:"Hours"},
  {name:"Truck - Welding Rig",rate:40,unit:"Hours"},
  {name:"Trailer - Lowboy - 2 Axle",rate:29,unit:"Hours"},
  {name:"Trailer - Tag Along",rate:50,unit:"Days"},
  {name:"Trailer - Tool Trailer - 18-25ft",rate:175,unit:"Days"},
  {section:"Earthmoving"},
  {name:"Backhoe",rate:450,unit:"Days"},
  {name:"Bobcat Skidsteer",rate:25,unit:"Hours"},
  {section:"Air, Compressors & Blast"},
  {name:"Air Compressor - 185 CFM",rate:195,unit:"Days"},
  {name:"Air Impact Wrench - 1in",rate:50,unit:"Days"},
  {name:"Air Spade / Knife",rate:55,unit:"Days"},
  {name:"Blast Rig - 4 Bag Pot w/ 185 CFM AC",rate:55.50,unit:"Hours"},
  {name:"Blast Rig - 1 Pot w/ 375 CFM AC",rate:500,unit:"Days"},
  {section:"Tools & Testing"},
  {name:"Bench & Volt Meter",rate:875.50,unit:"Month"},
  {name:"Beveling Band - 30in",rate:25,unit:"Days"},
  {name:"Dearman Pipe Clamps",rate:25,unit:"Days"},
  {name:"Gasoline Emergency Response Equipment",rate:5000,unit:"Week"},
  {name:"HEPA Vacuum",rate:100,unit:"Days"},
  {name:"Holiday Detector / Pipe Jeep",rate:72,unit:"Days"},
  {name:"Hydraulic Torque",rate:200,unit:"Days"},
  {name:"German Air Saw (plus Blades)",rate:125,unit:"Days"},
  {name:"Hydrotest / Pressure Test Equipment",rate:200,unit:"Days"},
  {name:"Jack Hammer",rate:72,unit:"Days"},
  {name:"Laser Pump Aligner",rate:275,unit:"Days"},
  {name:"LEL/Gas Monitor - 4 Gas",rate:50,unit:"Days"},
  {name:"Line Locator",rate:50,unit:"Days"},
  {name:"Pipe Band Crawler",rate:25,unit:"Days"},
  {name:"Pipe Beveling Machine 1.5-3in",rate:25,unit:"Days"},
  {name:"Pipe Beveling Machine 10-14in",rate:40,unit:"Days"},
  {name:"Pipe Beveling Machine 16-22in",rate:100,unit:"Days"},
  {name:"Tap Machine - 2in",rate:190,unit:"Days"},
  {name:"Torque Wrench - Pneumatic J5",rate:175,unit:"Days"},
  {name:"Torque Wrench w/Multiplier Hand",rate:25,unit:"Days"},
  {name:"Torque Wrench w/Sockets Hyd/Pneu",rate:195,unit:"Days"},
  {name:"Wach Pipe Cutting Saw",rate:30,unit:"Hours"},
  {name:"Cold Cutters 0.5-2in",rate:20,unit:"Days"},
  {name:"Cold Cutters 2-4in",rate:30,unit:"Days"},
  {name:"Cold Cutters 4-8in",rate:40,unit:"Days"},
  {name:"Cold Cutters 8-12in",rate:60,unit:"Days"},
  {name:"Cold Cutters 12-14in",rate:75,unit:"Days"},
  {name:"Concrete Saw (plus blades)",rate:50,unit:"Days"},
  {name:"Confined Space Equipment",rate:375,unit:"Days"},
  {name:"Hydrotest Pump",rate:125,unit:"Days"},
  {section:"Weld Rates"},
  {name:"1G Weld",rate:8,unit:"Ft"},
  {name:"3G Weld",rate:8.50,unit:"Ft"},
  {name:"4G Weld",rate:9,unit:"Ft"},
];

function getPositions(division){ return division==="Mechanical"?POSITIONS_MECHANICAL:POSITIONS_PIPELINE; }
function getEquipList(division){ return division==="Mechanical"?EQUIP_LIST_MECHANICAL:EQUIP_LIST_PIPELINE; }
const POSITIONS=POSITIONS_PIPELINE;
const EQUIP_LIST=EQUIP_LIST_PIPELINE;
function getAllPositions(){
  const all=[...POSITIONS_PIPELINE,...POSITIONS_MECHANICAL];
  const seen=new Set();
  return all.filter(p=>{if(seen.has(p.name))return false;seen.add(p.name);return true;});
}

const WMO={0:["Clear Sky","☀️"],1:["Mainly Clear","🌤️"],2:["Partly Cloudy","⛅"],3:["Overcast","☁️"],45:["Foggy","🌫️"],48:["Icy Fog","🌫️"],51:["Light Drizzle","🌦️"],53:["Drizzle","🌦️"],55:["Heavy Drizzle","🌦️"],61:["Light Rain","🌧️"],63:["Rain","🌧️"],65:["Heavy Rain","🌧️"],71:["Light Snow","🌨️"],73:["Snow","🌨️"],75:["Heavy Snow","❄️"],80:["Light Showers","🌦️"],81:["Showers","🌦️"],82:["Violent Showers","⛈️"],95:["Thunderstorm","⛈️"],96:["Thunderstorm + Hail","⛈️"],99:["Severe Thunderstorm","⛈️"]};
const DIVISIONS=["Mechanical","Pipeline","Structural","Manufacturing"];
const DIV_META={Mechanical:{icon:"⚙️",color:"#60A5FA",desc:"Mechanical projects and equipment"},Pipeline:{icon:"🔧",color:"#3B82F6",desc:"Pipeline construction and maintenance"},Structural:{icon:"🏗️",color:"#34D399",desc:"Structural steel and civil work"},Manufacturing:{icon:"🏭",color:"#8B5CF6",desc:"Shop fabrication & production"}};
const ROLES=["crew","foreman","pm","admin"];
const ROLE_META={crew:{label:"Field Crew",color:T.green,desc:"Submit daily reports and time cards"},foreman:{label:"Foreman",color:T.yellow,desc:"Reports, time, safety, equipment, docs, schedule"},pm:{label:"Project Manager",color:T.orange,desc:"Approve reports, manage jobs, PM dashboard"},estimator:{label:"Estimator",color:T.purple,desc:"Estimating platform access only"},admin:{label:"Admin",color:T.red,desc:"Full access, user management"}};

const PERMS={
  admin:     ["manage_users","create_job","edit_job","archive_job","approve_report","flag_report","view_dashboard","submit_report","time_card","safety","photos","docs","schedule","weather","subs","crew_equip","crew_directory","custom_reports","notifications","estimating"],
  pm:        ["create_job","edit_job","archive_job","approve_report","flag_report","view_dashboard","submit_report","time_card","safety","photos","docs","schedule","weather","subs","crew_equip","crew_directory","custom_reports","notifications"],
  estimator: ["estimating","view_dashboard","crew_directory"],
  foreman:   ["submit_report","time_card","safety","photos","docs","schedule","weather","subs","crew_equip","crew_directory"],
  crew:      ["submit_report","time_card","photos","crew_directory"],
};
const can=(user,action)=>(PERMS[user?.role]||PERMS.crew).includes(action);

const uid=()=>Math.random().toString(36).slice(2,9);
const today=()=>new Date().toISOString().split("T")[0];
const fmt=(n)=>Number(n||0).toLocaleString("en-US",{minimumFractionDigits:2,maximumFractionDigits:2});
const fmtDate=(d)=>d?new Date(d+"T12:00:00").toLocaleDateString("en-US",{month:"short",day:"numeric",year:"numeric"}):"—";
const fmtShort=(d)=>d?new Date(d+"T12:00:00").toLocaleDateString("en-US",{weekday:"short",month:"short",day:"numeric"}):"—";
const daysUntil=(d)=>{if(!d)return null;const diff=new Date(d+"T12:00:00")-new Date();return Math.ceil(diff/86400000);};
function laborAmt(r,division){const positions=getPositions(division);const p=positions.find(x=>x.name===r.classification);if(!p)return 0;if(p.flat)return p.rate;return p.rate*((parseFloat(r.regHrs)||0)+(parseFloat(r.otHrs)||0)*1.5+(parseFloat(r.travelHrs)||0));}
function equipAmt(r,division){
  let rate=parseFloat(r.rate)||0;
  if(!rate&&r.description){
    const eq=getEquipList(division).find(e=>!e.section&&e.name===r.description);
    if(eq)rate=eq.rate;
  }
  const qty=parseFloat(r.qty)||0;
  const usage=parseFloat(r.usage)||0;
  return qty*rate*(usage||1);
}
// Colonial-style billing: each material / rental line carries a manually entered
// markup percentage and tax amount. Blank fields contribute nothing, so reports
// created before these existed total exactly as they did before.
function matLineTotal(x){
  const base=parseFloat(x.amount)||0;
  const mk=base*((parseFloat(x.markup_pct)||0)/100);
  return base+mk+(parseFloat(x.tax_amount)||0);
}
function rentalLineTotal(x){
  const base=(parseFloat(x.qty)||0)*(parseFloat(x.rate)||0)*(parseFloat(x.usage)||1);
  const mk=base*((parseFloat(x.markup_pct)||0)/100);
  return base+mk+(parseFloat(x.tax_amount)||0);
}

function reportTotals(r,division){const labor=(r.labor||[]).reduce((s,x)=>s+laborAmt(x,division),0);const equip=(r.equipment||[]).reduce((s,x)=>s+equipAmt(x,division),0);const rental=(r.rental_equipment||[]).reduce((s,x)=>s+rentalLineTotal(x),0);const mats=(r.materials||[]).reduce((s,x)=>s+matLineTotal(x),0);const labor_hrs=(r.labor||[]).reduce((s,x)=>s+(parseFloat(x.regHrs)||0)+(parseFloat(x.otHrs)||0)+(parseFloat(x.travelHrs)||0),0);return{labor,equip,rental,mats,labor_hrs,grand:labor+equip+rental+mats};}
function calcHours(ci,co){if(!ci||!co)return 0;const[ih,im]=ci.split(":").map(Number);const[oh,om]=co.split(":").map(Number);const diff=(oh*60+om)-(ih*60+im);return diff>0?Math.round(diff/60*100)/100:0;}
function getWeekStart(){const d=new Date();const day=d.getDay();d.setDate(d.getDate()-(day===0?6:day-1));return d.toISOString().split("T")[0];}
async function compressImg(file,maxW=900,q=0.65){return new Promise(res=>{const rd=new FileReader();rd.onload=ev=>{const img=new Image();img.onload=()=>{const sc=Math.min(1,maxW/img.width);const c=document.createElement("canvas");c.width=Math.round(img.width*sc);c.height=Math.round(img.height*sc);c.getContext("2d").drawImage(img,0,0,c.width,c.height);res(c.toDataURL("image/jpeg",q));};img.src=ev.target.result;};rd.readAsDataURL(file);});}
async function fetchWeather(location){const gR=await fetch(`https://geocoding-api.open-meteo.com/v1/search?name=${encodeURIComponent(location)}&count=1&language=en&format=json`);const gD=await gR.json();if(!gD.results?.length)throw new Error(`Cannot find: "${location}"`);const{latitude:lat,longitude:lon,name,admin1}=gD.results[0];const wR=await fetch(`https://api.open-meteo.com/v1/forecast?latitude=${lat}&longitude=${lon}&current=temperature_2m,apparent_temperature,weathercode,windspeed_10m,precipitation&daily=temperature_2m_max,temperature_2m_min,precipitation_sum,weathercode&timezone=auto&temperature_unit=fahrenheit&windspeed_unit=mph&precipitation_unit=inch&forecast_days=1`);const wD=await wR.json();return{...wD,locationName:`${name}, ${admin1}`};}
async function notify(type,title,body,extra={}){try{await API.notifications.create({type,title,body,...extra});}catch{}}

function Spinner(){return(<div style={{display:"flex",justifyContent:"center",padding:"48px 0"}}><div style={{width:32,height:32,border:`3px solid ${T.border}`,borderTopColor:T.orange,borderRadius:"50%",animation:"spin 0.7s linear infinite"}}/><style>{`@keyframes spin{to{transform:rotate(360deg)}} select{color-scheme:dark;} select{background:#1A1A20 !important;color:#F0F4FF !important;border-color:#26262E !important;} select option{background:#1A1A20 !important;color:#F0F4FF !important;} select option:hover{background:#26262E !important;} select:focus{outline:none !important;} select *{background:#1A1A20 !important;color:#F0F4FF !important;}`}</style></div>);}
function ErrBanner({msg,onDismiss}){if(!msg)return null;return(<div style={{background:T.redLow,border:`1px solid ${T.red}40`,borderRadius:12,padding:"12px 14px",marginBottom:14,display:"flex",justifyContent:"space-between",alignItems:"center"}}><span style={{fontSize:13,color:T.red}}>⚠️ {msg}</span><button onClick={onDismiss} style={{background:"none",border:"none",color:T.red,cursor:"pointer",fontSize:18,padding:"0 0 0 10px"}}>×</button></div>);}
function Lightbox({src,onClose}){if(!src)return null;return(<div onClick={onClose} style={{position:"fixed",inset:0,zIndex:999,background:"rgba(0,0,0,0.95)",display:"flex",alignItems:"center",justifyContent:"center",padding:16}}><img src={src} alt="" style={{maxWidth:"100%",maxHeight:"90vh",borderRadius:12}} onClick={e=>e.stopPropagation()}/><button onClick={onClose} style={{position:"absolute",top:16,right:16,background:"#D4D4D8",border:"none",color:"#fff",borderRadius:"50%",width:36,height:36,fontSize:18,cursor:"pointer"}}>×</button></div>);}
function DashedAdd({label,onClick,color}){const c=color||T.muted;return(<button onClick={onClick} style={{width:"100%",border:`2px dashed ${c}50`,background:c+"08",color:c,borderRadius:14,padding:"14px",fontSize:15,cursor:"pointer",fontFamily:"inherit",fontWeight:600}}>{label}</button>);}
function StatBar({items}){return(<div style={{display:"grid",gridTemplateColumns:`repeat(${items.length},1fr)`,gap:8}}>{items.map(({label,val,color})=>(<div key={label} style={{background:T.card,border:`1px solid ${T.border}`,borderRadius:12,padding:"10px 8px",textAlign:"center"}}><div style={{fontSize:16,fontWeight:900,color:color||T.text}}>{val}</div><div style={{fontSize:9,color:T.muted,textTransform:"uppercase",letterSpacing:"0.7px",marginTop:2}}>{label}</div></div>))}</div>);}
function TopBar({title,sub,onBack,right}){return(<div style={{background:T.surface,borderBottom:`1px solid ${T.border}`,padding:"14px 16px",position:"sticky",top:0,zIndex:50}}>{onBack&&<button onClick={onBack} style={{background:"none",border:"none",color:T.sub,fontSize:13,cursor:"pointer",marginBottom:8,padding:0,fontFamily:"inherit"}}>← Back</button>}<div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start"}}><div style={{flex:1,minWidth:0}}><div style={{fontSize:20,fontWeight:900,color:T.text,letterSpacing:"-0.5px"}}>{title}</div>{sub&&<div style={{fontSize:12,color:T.muted,marginTop:2}}>{sub}</div>}</div>{right&&<div style={{flexShrink:0,marginLeft:12}}>{right}</div>}</div></div>);}

function LaborCard({row,onChange,onRemove,division}){const positions=getPositions(division);const pos=positions.find(p=>p.name===row.classification);const amt=laborAmt(row,division);const set=(k,v)=>{const u={...row,[k]:v};if(k==="classification"){const p=getPositions(division).find(x=>x.name===v);u.rate=p?p.rate:"";}onChange(u);};return(<div style={{...cardS,marginBottom:10,borderLeft:`3px solid ${T.orange}`}}><div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8,marginBottom:8}}><div style={{gridColumn:"1/-1"}}><label style={lbl}>Name</label><select value={row.name||""} onChange={e=>set("name",e.target.value)} style={inpSel}><option value="">— Select —</option>{NAMES.map(n=><option key={n}>{n}</option>)}</select></div><div style={{gridColumn:"1/-1"}}><label style={lbl}>Classification</label><select value={row.classification||""} onChange={e=>set("classification",e.target.value)} style={inpSel}><option value="">— Select —</option>{getAllPositions().map(p=><option key={p.name}>{p.name}</option>)}</select></div></div>{pos&&!pos.flat&&(<div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:8,marginBottom:8}}>{[["regHrs","Reg Hrs"],["otHrs","OT Hrs"],["travelHrs","Travel"]].map(([k,l])=>(<div key={k}><label style={lbl}>{l}</label><input type="number" min="0" step="0.5" placeholder="0" value={row[k]||""} onChange={e=>set(k,e.target.value)} style={inp}/></div>))}</div>)}<div style={{display:"flex",justifyContent:"space-between",alignItems:"center",paddingTop:8,borderTop:`1px solid ${T.border}`}}><span style={{fontSize:11,color:T.muted}}>{pos?`$${pos.rate.toFixed(2)}${pos.flat?" flat":"/hr"}`:""}</span><div style={{display:"flex",alignItems:"center",gap:10}}>{amt>0&&<span style={{fontSize:16,fontWeight:800,color:T.green}}>${fmt(amt)}</span>}<button onClick={onRemove} style={{background:"none",border:"none",color:T.red,cursor:"pointer",fontSize:20,padding:0}}>×</button></div></div></div>);}
function EquipCard({row,onChange,onRemove,division}){const eqList=getEquipList(division);const eq=eqList.find(e=>!e.section&&e.name===row.description);const amt=equipAmt(row);const set=(k,v)=>{const u={...row,[k]:v};if(k==="description"){const e=eqList.find(x=>!x.section&&x.name===v);u.rate=e?e.rate:"";u.unit=e?e.unit:"";}onChange(u);};return(<div style={{...cardS,marginBottom:10,borderLeft:`3px solid ${T.yellow}`}}><div style={{marginBottom:8}}><label style={lbl}>Equipment</label><select value={row.description||""} onChange={e=>set("description",e.target.value)} style={inpSel}><option value="">— Select —</option>{eqList.map((e,i)=>e.section?<option key={i} disabled>── {e.section} ──</option>:<option key={i} value={e.name}>{e.name}</option>)}</select></div><div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8,marginBottom:8}}><div><label style={lbl}>Qty</label><input type="number" min="0" placeholder="0" value={row.qty||""} onChange={e=>set("qty",e.target.value)} style={inp}/></div><div>
              <label style={lbl}>
                {eq
                  ?<span>{eq.unit==="Hours"?"⏱️ Hours":eq.unit==="Days"?"📅 Days":eq.unit==="Ft"?"📏 Feet":eq.unit==="Week"?"📅 Weeks":eq.unit==="Month"?"📅 Months":"📊 "+eq.unit}</span>
                  :"Hrs / Days"}
              </label>
              <input type="number" min="0" step="0.5" placeholder="0" value={row.usage||""} onChange={e=>set("usage",e.target.value)} style={inp}/>
            </div></div><div style={{display:"flex",justifyContent:"space-between",alignItems:"center",paddingTop:8,borderTop:`1px solid ${T.border}`}}><span style={{fontSize:11,color:T.muted}}>{eq?`$${eq.rate.toLocaleString()}/${eq.unit}`:""}</span><div style={{display:"flex",alignItems:"center",gap:10}}>{amt>0&&<span style={{fontSize:16,fontWeight:800,color:T.green}}>${fmt(amt)}</span>}<button onClick={onRemove} style={{background:"none",border:"none",color:T.red,cursor:"pointer",fontSize:20,padding:0}}>×</button></div></div></div>);}
function RentedEquipCard({row,onChange,onRemove}){
  const set=(k,v)=>onChange({...row,[k]:v});
  const amt=(parseFloat(row.qty)||0)*(parseFloat(row.rate)||0)*(parseFloat(row.usage)||1);
  return(
    <div style={{...cardS,marginBottom:10,borderLeft:`3px solid ${T.purple}`}}>
      <div style={{marginBottom:8}}>
        <label style={lbl}>Equipment Name / Description</label>
        <input type="text" placeholder="e.g. 40-Ton Crane, Scissor Lift, Generator…"
          value={row.description||""} onChange={e=>set("description",e.target.value)} style={inp}/>
      </div>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:8,marginBottom:8}}>
        <div><label style={lbl}>Qty</label>
          <input type="number" min="0" placeholder="0" value={row.qty||""} onChange={e=>set("qty",e.target.value)} style={inp}/></div>
        <div><label style={lbl}>Days / Hrs</label>
          <input type="number" min="0" step="0.5" placeholder="0" value={row.usage||""} onChange={e=>set("usage",e.target.value)} style={inp}/></div>
        <div><label style={lbl}>Rate / Unit</label>
          <input type="number" min="0" placeholder="0.00" value={row.rate||""} onChange={e=>set("rate",e.target.value)} style={inp}/></div>
      </div>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:8,marginBottom:8}}>
        <div><label style={lbl}>Markup %</label>
          <input type="number" min="0" step="0.1" placeholder="10" value={row.markup_pct||""} onChange={e=>set("markup_pct",e.target.value)} style={inp}/></div>
        <div><label style={lbl}>Tax ($)</label>
          <input type="number" min="0" step="0.01" placeholder="0.00" value={row.tax_amount||""} onChange={e=>set("tax_amount",e.target.value)} style={inp}/></div>
        <div><label style={lbl}>Line Total</label>
          <div style={{...inp,display:"flex",alignItems:"center",color:T.green,fontWeight:800}}>${fmt(rentalLineTotal(row))}</div></div>
      </div>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",paddingTop:8,borderTop:`1px solid ${T.border}`}}>
        <span style={{fontSize:11,color:T.muted}}>Qty × Rate × Days/Hrs{(parseFloat(row.markup_pct)||0)>0?` + ${row.markup_pct}%`:""}{(parseFloat(row.tax_amount)||0)>0?" + tax":""}</span>
        <div style={{display:"flex",alignItems:"center",gap:10}}>
          {rentalLineTotal(row)>0&&<span style={{fontSize:16,fontWeight:800,color:T.green}}>${fmt(rentalLineTotal(row))}</span>}
          <button onClick={onRemove} style={{background:"none",border:"none",color:T.red,cursor:"pointer",fontSize:20,padding:0}}>×</button>
        </div>
      </div>
    </div>
  );
}

function MatCard({row,onChange,onRemove}){const fileRef=useRef(null);const receipts=row.receipts||[];async function handleFiles(files){const n=[];for(const f of files){if(!f.type.startsWith("image/"))continue;const src=await compressImg(f,800,0.6);n.push({id:uid(),src});}onChange({...row,receipts:[...receipts,...n]});}return(<div style={{...cardS,marginBottom:10,borderLeft:`3px solid ${T.blue}`}}><div style={{display:"grid",gridTemplateColumns:"56px 1fr 88px",gap:8,marginBottom:10}}><div><label style={lbl}>Qty</label><input type="number" min="0" placeholder="0" value={row.qty||""} onChange={e=>onChange({...row,qty:e.target.value})} style={inp}/></div><div><label style={lbl}>Description</label><input type="text" placeholder="Item / material" value={row.description||""} onChange={e=>onChange({...row,description:e.target.value})} style={inp}/></div><div><label style={lbl}>Amount</label><input type="number" min="0" placeholder="0.00" value={row.amount||""} onChange={e=>onChange({...row,amount:e.target.value})} style={inp}/></div></div><div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:8,marginBottom:10}}><div><label style={lbl}>Markup %</label><input type="number" min="0" step="0.1" placeholder="12" value={row.markup_pct||""} onChange={e=>onChange({...row,markup_pct:e.target.value})} style={inp}/></div><div><label style={lbl}>Tax ($)</label><input type="number" min="0" step="0.01" placeholder="0.00" value={row.tax_amount||""} onChange={e=>onChange({...row,tax_amount:e.target.value})} style={inp}/></div><div><label style={lbl}>Line Total</label><div style={{...inp,display:"flex",alignItems:"center",color:T.green,fontWeight:800}}>${fmt(matLineTotal(row))}</div></div></div><div style={{borderTop:`1px solid ${T.border}`,paddingTop:10}}><label style={{...lbl,marginBottom:8}}>📎 Receipts</label><div style={{display:"flex",gap:8,flexWrap:"wrap",alignItems:"center"}}>{receipts.map(r=>(<div key={r.id} style={{position:"relative"}}><img src={r.src} alt="" style={{width:60,height:60,objectFit:"cover",borderRadius:10,border:`2px solid ${T.blue}40`,display:"block"}}/><button onClick={()=>onChange({...row,receipts:receipts.filter(x=>x.id!==r.id)})} style={{position:"absolute",top:-5,right:-5,width:18,height:18,borderRadius:"50%",background:T.red,border:"none",color:"#fff",fontSize:11,cursor:"pointer",display:"flex",alignItems:"center",justifyContent:"center"}}>×</button></div>))}<button onClick={()=>fileRef.current?.click()} style={{width:60,height:60,borderRadius:10,border:`2px dashed ${T.blue}40`,background:T.blueLow,color:T.blue,cursor:"pointer",display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",fontSize:18,gap:2}}><span>📷</span><span style={{fontSize:9,fontWeight:700}}>ADD</span></button><input ref={fileRef} type="file" accept="image/*" capture="environment" multiple style={{display:"none"}} onChange={e=>{handleFiles(Array.from(e.target.files));e.target.value="";}} /></div></div><div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginTop:10}}>{matLineTotal(row)>0&&<span style={{fontSize:14,fontWeight:700,color:T.green}}>${fmt(matLineTotal(row))}</span>}<button onClick={onRemove} style={{background:"none",border:"none",color:T.red,cursor:"pointer",fontSize:12,fontWeight:700,fontFamily:"inherit",marginLeft:"auto"}}>Remove</button></div></div>);}

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
      setProfile(rows&&rows.length>0?rows[0]:{name:n,role:"crew",division:null,pin:null});
    }catch{setProfile({name:n,role:"crew",division:null,pin:null});}
    setLoading(false);
  }

  async function handleLogin(){
    if(!name||!profile)return;
    if(profile.pin){
      if(pin!==profile.pin){ setErr("Incorrect PIN"); return; }
    } else if(name==="Admin"){
      if(pin!=="1234"){ setErr("Incorrect PIN (default: 1234)"); return; }
    }
    onLogin(profile);
  }

  const roleM=profile?ROLE_META[profile.role]:null;
  const pinIsSet=profile&&(profile.pin||name==="Admin");

  return(
    <div style={{minHeight:"100vh",backgroundImage:`url(${LOGIN_BG})`,backgroundSize:"cover",backgroundPosition:"center",backgroundRepeat:"no-repeat",display:"flex",flexDirection:"column",justifyContent:"center",padding:24,fontFamily:"inherit",position:"relative"}}>
      {/* Dark overlay for readability */}
      <div style={{position:"absolute",inset:0,background:"rgba(5,5,10,0.55)",backdropFilter:"blur(1px)",pointerEvents:"none"}}/>
      <div style={{position:"relative",zIndex:1,textAlign:"center",marginBottom:40}}>
        <div style={{display:"inline-flex",alignItems:"center",gap:12,marginBottom:12}}>
          <div style={{textAlign:"center"}}>
            <div style={{height:72,display:"flex",alignItems:"center",justifyContent:"center",marginBottom:6}}><span style={{fontSize:42,fontWeight:900,color:"#60A5FA",letterSpacing:2,fontFamily:"Arial Black,sans-serif",textShadow:"0 2px 8px rgba(96,165,250,0.4)"}}>AIME</span></div>
            <div style={{fontSize:13,color:T.muted,letterSpacing:"3px",textTransform:"uppercase",fontWeight:600}}>Field Pro</div>
          </div>
        </div>

      </div>

      <div style={{...cardS,maxWidth:420,margin:"0 auto",width:"100%",background:"rgba(8,10,18,0.88)",backdropFilter:"blur(12px)",WebkitBackdropFilter:"blur(12px)",borderRadius:20,border:"1px solid rgba(96,165,250,0.25)",padding:"28px 24px",boxShadow:"0 8px 32px rgba(0,0,0,0.6),0 0 0 1px rgba(255,255,255,0.05)"}}>
        <ErrBanner msg={err} onDismiss={()=>setErr("")}/>
        <div style={{marginBottom:14}}>
          <label style={{...lbl,color:"#C8D4F0"}}>Select Your Name</label>
          <select value={name} onChange={e=>handleNameChange(e.target.value)} style={{...inp,color:T.orange,background:T.card}}>
            <option value="" style={{background:T.card,color:T.muted}}>— Select your name —</option>
            {NAMES.map(n=><option key={n} style={{background:T.card,color:T.text}}>{n}</option>)}
            <option value="Admin" style={{background:T.card,color:T.orange}}>Admin</option>
          </select>
        </div>

        {loading&&<div style={{textAlign:"center",padding:"10px 0",color:T.muted,fontSize:13}}>Looking up profile…</div>}

        {profile&&!loading&&(
          <div style={{...cardS,marginBottom:14,background:T.surface,borderLeft:`3px solid ${roleM?.color||T.border}`}}>
            <div style={{display:"flex",alignItems:"center",gap:10}}>
              <div style={{width:36,height:36,borderRadius:10,background:(roleM?.color||T.muted)+"20",display:"flex",alignItems:"center",justifyContent:"center",fontSize:18}}>
                {profile.role==="admin"?"🔴":profile.role==="pm"?"🟠":profile.role==="foreman"?"🟡":"🟢"}
              </div>
              <div>
                <div style={{fontSize:14,fontWeight:700,color:T.text}}>{roleM?.label||"Field Crew"}</div>
                <div style={{fontSize:12,color:T.muted}}>{profile.division||"All Divisions"}</div>
              </div>
            </div>
          </div>
        )}

        {profile&&!loading&&(
          <div style={{marginBottom:14}}>
            <label style={lbl}>PIN</label>
            <input
              type="password" maxLength={6}
              placeholder={pinIsSet?"Enter your PIN":"No PIN set — contact admin"}
              value={pin}
              onChange={e=>setPin(e.target.value)}
              onKeyDown={e=>e.key==="Enter"&&handleLogin()}
              style={{...inp,opacity:pinIsSet?1:0.5}}
              disabled={!pinIsSet}
            />
            {!pinIsSet&&profile&&<div style={{fontSize:11,color:T.yellow,marginTop:4}}>⚠️ Ask your admin to set your PIN in PM Dashboard → Users</div>}
          </div>
        )}

        <button onClick={handleLogin} style={{...primBtn,opacity:name&&!loading&&(pinIsSet?pin.length>0:false)?1:0.3,boxShadow:name&&!loading&&(pinIsSet?pin.length>0:false)?"0 4px 20px rgba(96,165,250,0.4)":"none"}} disabled={!name||loading||!pinIsSet||pin.length===0}>
          {loading?"Loading…":"Sign In →"}
        </button>
      </div>
    </div>
  );
}

function DivisionScreen({user,projects,onSelect,onLogout,onCrew,onDash,onTimeCards,onEstimating,isOnline,pendingCount,onSync}){
  const divStats=DIVISIONS.map(div=>{
    const divProjects=projects.filter(p=>p.division===div&&p.status==="active");
    const totalBilled=divProjects.reduce((s,p)=>s+(p._billed||0),0);
    const totalReports=divProjects.reduce((s,p)=>s+(p._reports||0),0);
    return{div,count:divProjects.length,billed:totalBilled,reports:totalReports};
  });

  return(
    <div style={{background:T.bg,minHeight:"100vh",fontFamily:"inherit",color:T.text}}>
      {/* Header */}
      <div style={{background:T.surface,borderBottom:`1px solid ${T.border}`,padding:"16px"}}>
        {/* Offline / pending banner */}
        {!isOnline&&<div style={{background:"#7c2d12",borderRadius:10,padding:"8px 12px",marginBottom:10,display:"flex",alignItems:"center",gap:8}}><span style={{fontSize:16}}>📡</span><div><div style={{fontSize:13,fontWeight:700,color:"#fed7aa"}}>No Connection</div><div style={{fontSize:11,color:"#fdba74"}}>{pendingCount>0?`${pendingCount} report${pendingCount!==1?'s':''} will sync when back online`:"Reports will save locally until reconnected"}</div></div></div>}
        {isOnline&&pendingCount>0&&<div style={{background:T.greenLow,border:`1px solid ${T.green}40`,borderRadius:10,padding:"8px 12px",marginBottom:10,display:"flex",justifyContent:"space-between",alignItems:"center"}}><div style={{display:"flex",alignItems:"center",gap:8}}><span style={{fontSize:16}}>⏳</span><div style={{fontSize:13,fontWeight:700,color:T.green}}>{pendingCount} pending — tap to sync</div></div><button onClick={onSync} style={{background:T.green,color:"#0D0D0F",border:"none",borderRadius:8,padding:"5px 12px",fontSize:12,fontWeight:700,cursor:"pointer",fontFamily:"inherit"}}>Sync Now</button></div>}
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:6}}>
          <div style={{display:"flex",alignItems:"center",gap:10}}>
            <div style={{display:"flex",flexDirection:"column",gap:2}}>
              <div style={{fontSize:22,fontWeight:900,color:T.orange,letterSpacing:"1px"}}>AIME</div>
              <div style={{fontSize:9,color:T.muted,letterSpacing:"2.5px",textTransform:"uppercase",fontWeight:700,marginTop:1}}>Field Pro</div>
            </div>
            <div style={{fontSize:11,color:T.muted}}>
              {user.role==="admin"?"🔴":user.role==="pm"?"🟠":user.role==="foreman"?"🟡":"🟢"} {user.name} · {ROLE_META[user.role]?.label} {isOnline?<span style={{color:T.green,fontSize:10}}>● online</span>:<span style={{color:"#f97316",fontSize:10}}>● offline</span>}
            </div>
          </div>
          <div style={{display:"flex",gap:8}}>
            {can(user,"view_dashboard")&&<button onClick={onDash} style={{background:T.orangeLow,border:`1px solid ${T.orange}40`,borderRadius:10,padding:"8px 12px",color:T.orange,fontSize:12,fontWeight:700,cursor:"pointer",fontFamily:"inherit"}}>📊</button>}
            {(user.role==="admin"||user.role==="pm")&&<button onClick={onTimeCards} style={{background:T.greenLow,border:`1px solid ${T.green}40`,borderRadius:10,padding:"8px 12px",color:T.green,fontSize:12,fontWeight:700,cursor:"pointer",fontFamily:"inherit"}}>⏱️</button>}
            {can(user,"crew_directory")&&<button onClick={onCrew} style={{background:T.blueLow,border:`1px solid ${T.blue}40`,borderRadius:10,padding:"8px 12px",color:T.blue,fontSize:12,fontWeight:700,cursor:"pointer",fontFamily:"inherit"}}>👥</button>}
            {canEstimate(user)&&<button onClick={onEstimating} style={{background:`${T.purple}15`,border:`1px solid ${T.purple}40`,borderRadius:10,padding:"8px 12px",color:T.purple,fontSize:12,fontWeight:700,cursor:"pointer",fontFamily:"inherit"}}>📊</button>}
            <button onClick={onLogout} style={{...ghostBtn,padding:"8px 12px",fontSize:12}}>Out</button>
          </div>
        </div>
      </div>
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
                  <div style={{display:"flex",flexDirection:"column",alignItems:"flex-end",gap:4,flexShrink:0}}>
                    <div style={{fontSize:22,color:divColor}}>→</div>
                    {stats?.count>0&&<div style={{background:divColor+"20",border:`1px solid ${divColor}40`,borderRadius:20,padding:"2px 10px",fontSize:11,fontWeight:800,color:divColor}}>{stats.count} job{stats.count!==1?"s":""}</div>}
                    {stats?.count===0&&<div style={{background:T.border,borderRadius:20,padding:"2px 10px",fontSize:10,color:T.muted}}>No jobs</div>}
                  </div>
                </div>
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

function PullToRefresh({onRefresh,children}){
  const [pulling,setPulling]=React.useState(false);
  const [pullY,setPullY]=React.useState(0);
  const [refreshing,setRefreshing]=React.useState(false);
  const startY=React.useRef(null);
  const containerRef=React.useRef(null);
  const THRESHOLD=70;

  function onTouchStart(e){
    const container=containerRef.current;
    if(!container)return;
    if(container.scrollTop>0)return; // only trigger at top
    startY.current=e.touches[0].clientY;
  }
  function onTouchMove(e){
    if(startY.current===null)return;
    const dy=e.touches[0].clientY-startY.current;
    if(dy<=0){startY.current=null;return;}
    setPulling(true);
    setPullY(Math.min(dy*0.4,THRESHOLD+20));
  }
  async function onTouchEnd(){
    if(!pulling){startY.current=null;return;}
    if(pullY>=THRESHOLD){
      setRefreshing(true);
      setPullY(THRESHOLD);
      try{await onRefresh();}catch(e){}
      setRefreshing(false);
    }
    setPulling(false);
    setPullY(0);
    startY.current=null;
  }

  return(
    <div ref={containerRef} style={{overflowY:"auto",height:"100%",WebkitOverflowScrolling:"touch",position:"relative"}}
      onTouchStart={onTouchStart} onTouchMove={onTouchMove} onTouchEnd={onTouchEnd}>
      {/* Pull indicator */}
      <div style={{
        height:pullY>0?pullY:0,overflow:"hidden",display:"flex",alignItems:"center",justifyContent:"center",
        transition:pulling?"none":"height 0.3s ease",color:T.orange,fontSize:13,fontWeight:700,gap:6
      }}>
        {pullY>0&&<>
          <div style={{
            width:20,height:20,border:`2px solid ${T.orange}`,borderTopColor:"transparent",borderRadius:"50%",
            animation:refreshing?"spin 0.7s linear infinite":undefined,
            transform:refreshing?"none":`rotate(${Math.min(pullY/THRESHOLD,1)*360}deg)`,transition:"transform 0.1s"
          }}/>
          {pullY>=THRESHOLD?"Release to refresh":"Pull to refresh"}
        </>}
      </div>
      {children}
    </div>
  );
}

function JobBoard({user,division,projects,loading,onSelect,onNew,onBack,onRefresh}){
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
  const canCreate=user.role==="admin"||user.role==="pm"||can(user,"create_job");

  return(
    <div style={{background:T.bg,minHeight:"100vh",fontFamily:"inherit",color:T.text}}>
      <div style={{background:T.surface,borderBottom:`1px solid ${T.border}`,padding:"14px 16px 0"}}>
        <button onClick={onBack} style={{background:"none",border:"none",color:T.sub,fontSize:13,cursor:"pointer",marginBottom:10,padding:0,fontFamily:"inherit"}}>← Divisions</button>
        <div style={{display:"flex",alignItems:"center",gap:12,marginBottom:14}}>
          <div style={{width:44,height:44,borderRadius:14,background:meta.color+"20",border:`2px solid ${meta.color}40`,display:"flex",alignItems:"center",justifyContent:"center",fontSize:22,flexShrink:0}}>{meta.icon}</div>
          <div style={{flex:1}}>
            <div style={{fontSize:20,fontWeight:900,color:T.text,letterSpacing:"-0.5px"}}>{division}</div>
            <div style={{fontSize:11,color:T.muted}}>{active.length} active job{active.length!==1?"s":""}</div>
          </div>
          {canCreate&&<button onClick={onNew} style={{background:T.orange,color:"#0D0D0F",border:"none",borderRadius:12,padding:"10px 16px",fontSize:13,fontWeight:800,cursor:"pointer",fontFamily:"inherit"}}>+ New Job</button>}
        </div>
        <div style={{position:"relative",marginBottom:10}}>
          <span style={{position:"absolute",left:13,top:"50%",transform:"translateY(-50%)",fontSize:14,pointerEvents:"none"}}>🔍</span>
          <input type="text" placeholder="Search jobs…" value={search} onChange={e=>setSearch(e.target.value)} style={{...inp,paddingLeft:38,borderRadius:12,fontSize:14}}/>
          {search&&<button onClick={()=>setSearch("")} style={{position:"absolute",right:12,top:"50%",transform:"translateY(-50%)",background:"none",border:"none",color:T.muted,cursor:"pointer",fontSize:18,padding:0}}>×</button>}
        </div>
        <div style={{display:"flex",gap:6}}>
          {[["active","Active"],["archived","Archived"],["all","All"]].map(([v,l])=>(<button key={v} onClick={()=>setFilter(v)} style={{padding:"8px 14px",borderRadius:"10px 10px 0 0",background:filter===v?T.bg:"transparent",border:filter===v?`1px solid ${T.border}`:"1px solid transparent",borderBottom:filter===v?`1px solid ${T.bg}`:"none",color:filter===v?T.text:T.muted,fontSize:13,fontWeight:filter===v?700:500,cursor:"pointer",fontFamily:"inherit",position:"relative",zIndex:filter===v?1:0,marginBottom:filter===v?-1:0}}>{l}{v==="active"&&active.length>0&&<span style={{marginLeft:5,background:meta.color+"25",color:meta.color,borderRadius:20,padding:"1px 6px",fontSize:10,fontWeight:800}}>{active.length}</span>}</button>))}
        </div>
      </div>

      <PullToRefresh onRefresh={async()=>onRefresh&&await onRefresh()}>
      <div style={{padding:"12px 16px 80px"}}>
        {loading&&<Spinner/>}
      {canCreate&&<div style={{position:"fixed",bottom:20,right:"max(16px,calc(50vw - 224px))",zIndex:100}}><button onClick={onNew} style={{background:T.orange,color:"#0D0D0F",border:"none",borderRadius:50,padding:"14px 22px",fontSize:15,fontWeight:900,cursor:"pointer",fontFamily:"inherit",boxShadow:"0 4px 24px rgba(249,115,22,0.5)",display:"flex",alignItems:"center",gap:8}}>＋ New Job</button></div>}
        {!loading&&filtered.length===0&&(
          <div style={{textAlign:"center",padding:"60px 20px",color:T.muted}}>
            <div style={{fontSize:48,marginBottom:12}}>{meta.icon}</div>
            <div style={{fontSize:17,fontWeight:700,color:T.sub,marginBottom:6}}>{search?`No jobs matching "${search}"`:filter==="archived"?"No archived jobs":"No active jobs in "+division}</div>
            {!search&&filter==="active"&&canCreate&&<div style={{fontSize:13}}>Tap + New Job to create one.</div>}
          </div>
        )}
        {!loading&&filtered.map(p=><JobCard key={p.id} p={p} onSelect={onSelect} divColor={meta.color}/>)}
      </div>
      </PullToRefresh>
    </div>
  );
}

function JobCard({p,onSelect,divColor}){
  const isArchived=p.status!=="active";
  const daysSince=p._lastReport?Math.floor((Date.now()-new Date(p._lastReport+"T12:00:00").getTime())/86400000):null;
  const actColor=daysSince===null?T.muted:daysSince===0?T.green:daysSince<=2?T.orange:T.red;
  const actLabel=daysSince===null?"No reports yet":daysSince===0?"Today":daysSince===1?"Yesterday":`${daysSince}d ago`;
  const c=divColor||T.orange;
  const billedAmt=p._billed||0;
  const billedFmt=billedAmt>=1000?(billedAmt/1000).toFixed(1)+"k":fmt(billedAmt);
  return(
    <div onClick={()=>onSelect(p)} style={{background:T.card,border:`1px solid ${T.border}`,borderRadius:16,marginBottom:10,cursor:"pointer",overflow:"hidden",opacity:isArchived?0.55:1}}>
      <div style={{height:3,background:isArchived?T.border:`linear-gradient(90deg,${c},${c}88)`}}/>
      <div style={{padding:"14px 16px"}}>

        {/* Job name + billed */}
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:8}}>
          <div style={{flex:1,minWidth:0,paddingRight:12}}>
            <div style={{fontSize:17,fontWeight:900,color:T.text,letterSpacing:"-0.3px",lineHeight:1.2}}>{p.name}</div>
            <div style={{fontSize:12,color:T.sub,marginTop:3}}>{[p.client,p.location].filter(Boolean).join(" · ")||"No details"}</div>
          </div>
          <div style={{textAlign:"right",flexShrink:0}}>
            <div style={{fontSize:20,fontWeight:900,color:T.green,letterSpacing:"-0.5px"}}>${billedFmt}</div>
            <div style={{fontSize:9,color:T.muted,textTransform:"uppercase",letterSpacing:"0.5px",marginTop:1}}>Billed</div>
          </div>
        </div>
        <div style={{display:"flex",gap:6,flexWrap:"wrap",marginBottom:10}}>
          {p.afe&&<span style={pill(T.muted)}>AFE: {p.afe}</span>}
          {p.work_order&&<span style={pill(T.muted)}>PO: {p.work_order}</span>}
          <span style={pill(p.job_type==="Contract"?T.blue:T.orange)}>{p.job_type||"T&M"}</span>
          <span style={pill(isArchived?T.muted:T.green)}>{isArchived?"Archived":"Active"}</span>
        </div>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr 1fr",gap:6,marginBottom:10}}>
          {[
            [p._reports||0,"📋","Reports",c],
            [p._photos||0,"📷","Photos",T.blue],
            [p._openRfis||0,"📝","Open RFIs",p._openRfis>0?T.yellow:T.muted],
            [p._pendingCOs||0,"📋","Pending COs",p._pendingCOs>0?T.orange:T.muted],
          ].map(([val,icon,label,color])=>(
            <div key={label} style={{background:T.surface,borderRadius:8,padding:"6px 4px",textAlign:"center",border:val>0&&(label==="Open RFIs"||label==="Pending COs")?`1px solid ${color}40`:`1px solid ${T.border}`}}>
              <div style={{fontSize:14,fontWeight:900,color:val>0?color:T.muted}}>{val}</div>
              <div style={{fontSize:8,color:T.muted,textTransform:"uppercase",letterSpacing:"0.3px",marginTop:1}}>{label}</div>
            </div>
          ))}
        </div>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",paddingTop:10,borderTop:`1px solid ${T.border}`}}>
          <div style={{display:"flex",alignItems:"center",gap:6}}>
            <div style={{width:6,height:6,borderRadius:"50%",background:actColor,flexShrink:0}}/>
            <span style={{fontSize:11,fontWeight:600,color:actColor}}>{actLabel}</span>
            {p._pendingCOTotal>0&&<span style={{fontSize:10,color:T.orange,fontWeight:700,marginLeft:4}}>+${(p._pendingCOTotal/1000).toFixed(1)}k pending</span>}
          </div>
          <div style={{display:"flex",alignItems:"center",gap:6,background:c+"15",border:`1px solid ${c}40`,borderRadius:10,padding:"7px 12px"}}>
            <span style={{fontSize:12,fontWeight:700,color:c}}>Enter Job</span>
            <span style={{fontSize:14,color:c}}>→</span>
          </div>
        </div>
      </div>
    </div>
  );
}

function ProjectForm({initial,onSave,onCancel,saving,defaultDivision,externalErr,onClearErr}){
  const [f,setF]=useState(initial||{name:"",client:"",location:"",afe:"",work_order:"",start_date:today(),notes:"",status:"active",division:defaultDivision||"Pipeline",job_type:"T&M",contract_value:"",contract_hours:"",estimated_budget:""});
  const set=(k,v)=>setF(x=>({...x,[k]:v}));
  return(
    <div style={{background:T.bg,minHeight:"100vh",fontFamily:"inherit",color:T.text}}>
      <TopBar title={initial?"Edit Job":"New Job"} onBack={onCancel}/>
      <div style={{padding:"16px 16px 100px"}}>
        <ErrBanner msg={externalErr} onDismiss={onClearErr}/>
        {[{k:"name",l:"Job Number *",ph:"e.g. HDD-2026-001"},{k:"client",l:"Client",ph:"Colonial Pipeline"},{k:"location",l:"Location",ph:"City, State or Milepost"},{k:"afe",l:"AFE No.",ph:"AFE #"},{k:"work_order",l:"PO #",ph:"PO #"}].map(({k,l,ph})=>(<div key={k} style={{marginBottom:12}}><label style={lbl}>{l}</label><input type="text" placeholder={ph} value={f[k]||""} onChange={e=>set(k,e.target.value)} style={inp}/></div>))}

        {/* Job Type */}
        <div style={{marginBottom:12}}>
          <label style={lbl}>Job Type</label>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8}}>
            {[["T&M","⏱️ Time & Material"],["Contract","📋 Contract"]].map(([val,label])=>(
              <button key={val} onClick={()=>set("job_type",val)} style={{
                padding:"14px 10px",borderRadius:12,
                border:`2px solid ${f.job_type===val?T.orange:T.border}`,
                background:f.job_type===val?T.orangeLow:T.surface,
                color:f.job_type===val?T.orange:T.sub,
                fontWeight:700,fontSize:14,cursor:"pointer",fontFamily:"inherit",textAlign:"center"
              }}>{label}</button>
            ))}
          </div>
        </div>
        {f.job_type==="Contract"&&(<>
          <div style={{marginBottom:12}}>
            <label style={lbl}>Contract Total Value ($)</label>
            <input
              type="number"
              min="0"
              step="0.01"
              placeholder="e.g. 250000.00"
              value={f.contract_value||""}
              onChange={e=>set("contract_value",e.target.value)}
              style={inp}
            />
          </div>
          <div style={{marginBottom:12}}>
            <label style={lbl}>Contract Total Hours</label>
            <input
              type="number"
              min="0"
              step="0.5"
              placeholder="e.g. 2000 — total hours budgeted for this contract"
              value={f.contract_hours||""}
              onChange={e=>set("contract_hours",e.target.value)}
              style={inp}
            />
            <div style={{fontSize:11,color:T.muted,marginTop:4}}>Track hours used vs. contract hours — shows a progress bar on the job.</div>
          </div>
        </>)}
        {/* Estimated budget for T&M jobs */}
        {f.job_type==="T&M"&&(
          <div style={{marginBottom:12}}>
            <label style={lbl}>Estimated Budget (optional, $)</label>
            <input
              type="number"
              min="0"
              step="0.01"
              placeholder="e.g. 50000.00 — leave blank if open-ended"
              value={f.estimated_budget||""}
              onChange={e=>set("estimated_budget",e.target.value)}
              style={inp}
            />
            <div style={{fontSize:11,color:T.muted,marginTop:4}}>Used to show budget vs. actual billing progress on the job.</div>
          </div>
        )}

        <div style={{marginBottom:12}}>
          <label style={lbl}>Division</label>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:8}}>
            {DIVISIONS.map(div=>{const m=DIV_META[div];return(<button key={div} onClick={()=>set("division",div)} style={{padding:"12px 8px",borderRadius:12,border:`2px solid ${f.division===div?m.color:T.border}`,background:f.division===div?m.color+"20":T.surface,color:f.division===div?m.color:T.sub,fontWeight:700,fontSize:12,cursor:"pointer",fontFamily:"inherit",textAlign:"center"}}><div style={{fontSize:20,marginBottom:4}}>{m.icon}</div>{div}</button>);})}
          </div>
        </div>
        <div style={{marginBottom:12}}><label style={lbl}>Start Date</label><input type="date" value={f.start_date||today()} onChange={e=>set("start_date",e.target.value)} style={inp}/></div>
        <div style={{marginBottom:20}}><label style={lbl}>Notes</label><textarea placeholder="Project notes, scope, special instructions…" value={f.notes||""} onChange={e=>set("notes",e.target.value)} rows={3} style={{...inp,resize:"vertical",lineHeight:1.5}}/></div>
        <button onClick={()=>f.name.trim()&&!saving&&onSave({...f,contract_value:f.contract_value&&f.contract_value!==""?parseFloat(f.contract_value):null,contract_hours:f.contract_hours&&f.contract_hours!==""?parseFloat(f.contract_hours):null,estimated_budget:f.estimated_budget&&f.estimated_budget!==""?parseFloat(f.estimated_budget):null})} style={{...primBtn,opacity:f.name.trim()&&!saving?1:0.5}}>{saving?"Saving…":initial?"Save Changes":"Create Job"}</button>
      </div>
    </div>
  );
}

const RSTEPS=["Job Info","Labor","Equipment","Materials","Site Notes","Review"];

async function autoPopulateTimeCards(report, project){
  const labor=(report.labor||[]).filter(l=>l.name&&l.name.trim());
  if(!labor.length) return {created:0,updated:0};
  let created=0,updated=0;
  for(const entry of labor){
    const reg=parseFloat(entry.regHrs)||0;
    const ot=parseFloat(entry.otHrs)||0;
    const travel=parseFloat(entry.travelHrs)||0;
    if(reg+ot+travel===0) continue;
    try{
      const existing=await API.timeCards.find(entry.name,report.date,project.id);
      const card=Array.isArray(existing)?existing[0]:null;
      if(card){
        const newReg=(parseFloat(card.reg_hours)||0)+reg;
        const newOT=(parseFloat(card.ot_hours)||0)+ot;
        const newTravel=(parseFloat(card.travel_hours)||0)+travel;
        await API.timeCards.update(card.id,{
          reg_hours:newReg,
          ot_hours:newOT,
          travel_hours:newTravel,
          total_hours:newReg+newOT+newTravel,
        });
        updated++;
      }else{
        await API.timeCards.create({
          worker_name:entry.name,
          date:report.date,
          project_id:project.id,
          division:project.division,
          classification:entry.classification||"",
          reg_hours:reg,
          ot_hours:ot,
          travel_hours:travel,
          total_hours:reg+ot+travel,
          notes:`Auto-filled from daily report${report.report_no?" #"+report.report_no:""} · ${project.name}`,
        });
        created++;
      }
    }catch(e){

    }
  }
  return{created,updated};
}

function VisitorAddRow({onAdd}){
  const [show,setShow]=useState(false);
  const [f,setF]=useState({name:"",company:"",type:"Inspector",notes:""});
  if(!show) return <button onClick={()=>setShow(true)} style={{...ghostBtn,width:"100%",textAlign:"center",fontSize:13,marginTop:4}}>+ Add Visitor</button>;
  return(
    <div style={{background:"#0C0C0F",borderRadius:10,padding:12,marginTop:8}}>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8,marginBottom:8}}>
        <div><label style={lbl}>Name *</label><input value={f.name} onChange={e=>setF(x=>({...x,name:e.target.value}))} placeholder="Full name" style={inp}/></div>
        <div><label style={lbl}>Company</label><input value={f.company} onChange={e=>setF(x=>({...x,company:e.target.value}))} placeholder="Company" style={inp}/></div>
      </div>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8,marginBottom:8}}>
        <div><label style={lbl}>Type</label>
          <select value={f.type} onChange={e=>setF(x=>({...x,type:e.target.value}))} style={inpSel}>
            {["Inspector","Client","Engineer","Vendor","Safety Officer","Owner Rep","Subcontractor","Other"].map(t=><option key={t}>{t}</option>)}
          </select>
        </div>
        <div><label style={lbl}>Notes</label><input value={f.notes} onChange={e=>setF(x=>({...x,notes:e.target.value}))} placeholder="Purpose of visit" style={inp}/></div>
      </div>
      <div style={{display:"flex",gap:8}}>
        <button onClick={()=>{if(!f.name.trim())return;onAdd({...f});setF({name:"",company:"",type:"Inspector",notes:""});setShow(false);}}
          disabled={!f.name.trim()} style={{...primBtn,flex:2,borderRadius:10,fontSize:13,opacity:f.name.trim()?1:0.5}}>Add Visitor</button>
        <button onClick={()=>setShow(false)} style={{...ghostBtn,flex:1,textAlign:"center",fontSize:13}}>Cancel</button>
      </div>
    </div>
  );
}

function DelayAddRow({onAdd}){
  const [show,setShow]=useState(false);
  const CAUSES=["Weather","Material Delay","Equipment Breakdown","Owner/Client","Subcontractor","Labor","Design/Engineering","Permitting","Safety Stop","Other"];
  const [f,setF]=useState({cause:"Weather",description:"",hours:0,impact:""});
  if(!show) return <button onClick={()=>setShow(true)} style={{...ghostBtn,width:"100%",textAlign:"center",fontSize:13,marginTop:4,color:T.red,border:`1px solid ${T.red}30`}}>+ Log Delay / Issue</button>;
  return(
    <div style={{background:"#0C0C0F",borderRadius:10,padding:12,marginTop:8,border:`1px solid ${T.red}30`}}>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8,marginBottom:8}}>
        <div><label style={lbl}>Cause *</label>
          <select value={f.cause} onChange={e=>setF(x=>({...x,cause:e.target.value}))} style={{...inpSel,color:T.red}}>
            {CAUSES.map(c=><option key={c}>{c}</option>)}
          </select>
        </div>
        <div><label style={lbl}>Hours Lost</label><input type="number" value={f.hours} onChange={e=>setF(x=>({...x,hours:e.target.value}))} min={0} step={0.5} style={inp}/></div>
      </div>
      <div style={{marginBottom:8}}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:6}}>
          <label style={lbl}>Description *</label>
        </div>
        <input value={f.description} onChange={e=>setF(x=>({...x,description:e.target.value}))} placeholder="What happened?" style={inp}/>
      </div>
      <div style={{marginBottom:10}}><label style={lbl}>Impact / Action Taken</label>
        <input value={f.impact} onChange={e=>setF(x=>({...x,impact:e.target.value}))} placeholder="How did it affect work?" style={inp}/></div>
      <div style={{display:"flex",gap:8}}>
        <button onClick={()=>{if(!f.description.trim())return;onAdd({...f,hours:parseFloat(f.hours)||0});setF({cause:"Weather",description:"",hours:0,impact:""});setShow(false);}}
          disabled={!f.description.trim()} style={{...primBtn,flex:2,borderRadius:10,fontSize:13,background:T.red,opacity:f.description.trim()?1:0.5}}>Log Delay</button>
        <button onClick={()=>setShow(false)} style={{...ghostBtn,flex:1,textAlign:"center",fontSize:13}}>Cancel</button>
      </div>
    </div>
  );
}

function DailyReportForm({user,project,onSave,onCancel,isOnline}){
  const draftKey=`${project.id}_${user.name}`;
  const existingDraft=loadDraft(draftKey);

  const [step,setStep]=useState(1);const [saving,setSaving]=useState(false);
  const [draftSaved,setDraftSaved]=useState(false);
  const [showDraftBanner,setShowDraftBanner]=useState(!!existingDraft);
  const [rpt,setRpt]=useState(existingDraft?.data||{date:today(),description:"",report_no:"",labor:[],equipment:[],rental_equipment:[],materials:[],visitor_log:[],delays:[],site_conditions:""});
  const topRef=useRef(null);
  const setR=(k,v)=>setRpt(r=>({...r,[k]:v}));
  function add(key,item){setR(key,[...rpt[key],item]);}
  function upd(key,i,row){const a=[...rpt[key]];a[i]=row;setR(key,a);}
  function del(key,i){setR(key,rpt[key].filter((_,j)=>j!==i));}

  const [listening,setListening]=useState(false);
  const [listenTarget,setListenTarget]=useState(null);
  function startVoice(key){
    const SR=window.SpeechRecognition||window.webkitSpeechRecognition;
    if(!SR){alert("Voice-to-text not supported in this browser. Try Chrome.");return;}
    const r=new SR();r.lang="en-US";r.continuous=false;r.interimResults=false;
    setListening(true);setListenTarget(key);
    r.onresult=e=>{
      const txt=e.results[0][0].transcript;
      setR(key,(rpt[key]||"")+(rpt[key]?" ":"")+txt);
      setListening(false);setListenTarget(null);
    };
    r.onerror=()=>{setListening(false);setListenTarget(null);};
    r.onend=()=>{setListening(false);setListenTarget(null);};
    r.start();
  }

  const [weatherFilling,setWeatherFilling]=useState(false);
  async function autoFillWeather(){
    const WMO={0:"Clear Sky",1:"Mainly Clear",2:"Partly Cloudy",3:"Overcast",45:"Foggy",48:"Icy Fog",51:"Light Drizzle",53:"Drizzle",55:"Heavy Drizzle",61:"Light Rain",63:"Rain",65:"Heavy Rain",71:"Light Snow",73:"Snow",75:"Heavy Snow",80:"Rain Showers",81:"Heavy Showers",82:"Violent Showers",95:"Thunderstorm",99:"Thunderstorm w/Hail"};
    let location=project.location?.trim();
    if(!location){
      location=window.prompt("Enter city or zip code for weather lookup:\n(e.g. \"Houston, TX\" or \"77002\")\n\nTip: Add a Location to this job to skip this step next time.","");
      if(!location||!location.trim())return;
      location=location.trim();
    }
    setWeatherFilling(true);
    try{
      const w=await fetchWeather(location);
      const cur=w.current||{};
      const temp=Math.round(cur.temperature_2m||0);
      const wind=Math.round(cur.windspeed_10m||0);
      const precip=cur.precipitation||0;
      const code=cur.weathercode||0;
      const condition=WMO[code]||`Code ${code}`;
      setR("site_conditions",`${condition} · ${temp}°F · Wind: ${wind}mph${precip>0?` · Precip: ${precip}in`:""}${w.locationName?` · ${w.locationName}`:""}`);
    }catch(e){alert("Could not fetch weather: "+e.message+"\n\nTry a city name like \"Houston, TX\" or a zip code.");}
    setWeatherFilling(false);
  }
  const tot=reportTotals(rpt,project.division);

  useEffect(()=>{
    if(!rpt.report_no){
      (async()=>{
        try{
          const existing=await API.reports.forProject(project.id);
          const nums=(existing||[]).map(r=>{
            const n=parseInt((r.report_no||"0").replace(/\D/g,""));
            return isNaN(n)?0:n;
          });
          const nextNum=(nums.length>0?Math.max(...nums):0)+1;
          setR("report_no",String(nextNum).padStart(3,"0"));
        }catch(e){}
      })();
    }
  },[]);

  useEffect(()=>{
    const t=setTimeout(()=>{
      saveDraft(draftKey,rpt);
      setDraftSaved(true);
      setTimeout(()=>setDraftSaved(false),2000);
    },800);
    return()=>clearTimeout(t);
  },[rpt]);

  async function submit(){
    setSaving(true);
    const{rental_equipment,...rptClean}=rpt;
    const reportData={...rptClean,submitted_by:user.name,status:"submitted",project_id:project.id,rental_equipment:rental_equipment||[]};

    if(isOnline){
      try{
        const existing=await API.reports.forProject(project.id);
        const dupe=(existing||[]).find(r=>r.date===rpt.date);
        if(dupe){
          const proceed=window.confirm(
            `⚠️ A report for ${fmtDate(rpt.date)} already exists on this job (submitted by ${dupe.submitted_by||"someone"}).\n\nDo you still want to submit a second report for this date?`
          );
          if(!proceed){setSaving(false);return;}
        }
      }catch{/* ignore — if check fails just proceed */}
    }

    if(!isOnline){
      addToQueue({type:'report',data:reportData});
      clearDraft(draftKey);
      setSaving(false);
      alert("No connection — report saved and will sync automatically when you're back online.");
      onCancel();
      return;
    }
    try{
      await onSave(reportData);
      clearDraft(draftKey);
      try{
        const tcResult=await autoPopulateTimeCards(reportData,project);

      }catch(e){}
      await notify("report_submitted","New Report Submitted",`${user.name} submitted a report for ${project.name}`,{project_id:project.id});
    }catch(e){
      addToQueue({type:'report',data:reportData});
      clearDraft(draftKey);
      alert("Couldn't reach server — report queued and will sync when reconnected.");
      onCancel();
    }
    setSaving(false);
  }
  const scroll=()=>topRef.current?.scrollIntoView({behavior:"smooth"});
  const divMeta=DIV_META[project.division]||{color:T.orange};
  return(
    <div ref={topRef} style={{background:T.bg,minHeight:"100vh",fontFamily:"inherit"}}>
      <div style={{background:T.surface,borderBottom:`1px solid ${T.border}`,padding:"14px 16px",position:"sticky",top:0,zIndex:50}}>
        {/* Offline indicator */}
        {!isOnline&&<div style={{background:"#7c2d12",borderRadius:8,padding:"6px 10px",marginBottom:8,fontSize:12,color:"#fed7aa",display:"flex",alignItems:"center",gap:6}}><span>📡</span>Offline — report will save locally and sync when reconnected</div>}
        {/* Draft restore banner */}
        {showDraftBanner&&<div style={{background:T.blueLow,border:`1px solid ${T.blue}40`,borderRadius:8,padding:"8px 10px",marginBottom:8,display:"flex",justifyContent:"space-between",alignItems:"center"}}><div style={{fontSize:12,color:T.blue}}><strong>Draft restored</strong> · saved {existingDraft?.saved_at?new Date(existingDraft.saved_at).toLocaleTimeString():""}</div><button onClick={()=>setShowDraftBanner(false)} style={{background:"none",border:"none",color:T.blue,cursor:"pointer",fontSize:16,padding:0}}>×</button></div>}
        <div style={{display:"flex",justifyContent:"space-between",marginBottom:12}}>
          <div style={{display:"flex",alignItems:"center",gap:8}}><div style={{fontSize:16,fontWeight:800}}>New Daily Report</div>{draftSaved&&<span style={{fontSize:10,color:T.green,fontWeight:600}}>✓ Draft saved</span>}</div>
          <button onClick={onCancel} style={{background:"none",border:"none",color:T.sub,cursor:"pointer",fontSize:13,fontFamily:"inherit"}}>Cancel</button>
        </div>
        <div style={{display:"flex",alignItems:"center"}}>{RSTEPS.map((s,i)=>(<div key={i} style={{display:"flex",alignItems:"center",flex:i<RSTEPS.length-1?1:undefined}}><div style={{display:"flex",flexDirection:"column",alignItems:"center",gap:3}}><div style={{width:26,height:26,borderRadius:"50%",background:i+1<step?T.green:i+1===step?divMeta.color:T.border,color:i+1<=step?"#0D0D0F":T.muted,display:"flex",alignItems:"center",justifyContent:"center",fontSize:11,fontWeight:800}}>{i+1<step?"✓":i+1}</div><div style={{fontSize:8,color:i+1===step?divMeta.color:T.muted,fontWeight:i+1===step?700:400,whiteSpace:"nowrap"}}>{s}</div></div>{i<RSTEPS.length-1&&<div style={{flex:1,height:2,background:i+1<step?T.green:T.border,margin:"0 3px",marginBottom:14}}/>}</div>))}</div>
      </div>
      <div style={{padding:"14px 16px 100px"}}>
        {step===1&&(<div>
          <div style={{...cardS,marginBottom:14,borderLeft:`3px solid ${divMeta.color}`}}><div style={{fontSize:11,color:T.muted,marginBottom:2}}>Project · {project.division}</div><div style={{fontSize:15,fontWeight:700}}>{project.name}</div>{project.afe&&<div style={{fontSize:12,color:T.sub}}>AFE: {project.afe}{project.work_order?" · PO: "+project.work_order:""}</div>}<div style={{marginTop:6,display:"flex",gap:6}}><span style={pill(divMeta.color)}>{divMeta.icon} {project.division} Rates</span></div></div>
          <div style={{marginBottom:12}}><label style={lbl}>Date</label><input type="date" value={rpt.date} onChange={e=>setR("date",e.target.value)} style={inp}/></div>
          <div style={{marginBottom:12}}><label style={lbl}>Report No.</label><input type="text" placeholder="Report #" value={rpt.report_no||""} onChange={e=>setR("report_no",e.target.value)} style={inp}/></div>
          <div style={{...cardS,marginBottom:12,border:`1px solid ${T.blue}30`,padding:"12px 14px"}}>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:8}}>
              <label style={lbl}>🌤️ Site Conditions / Weather</label>
              <button onClick={autoFillWeather} disabled={weatherFilling}
                style={{background:T.blueLow,border:`1px solid ${T.blue}40`,borderRadius:8,padding:"5px 10px",color:T.blue,fontSize:11,fontWeight:700,cursor:"pointer",fontFamily:"inherit"}}>
                {weatherFilling?"Fetching…":"⚡ Auto-Fill"}
              </button>
            </div>
            <input value={rpt.site_conditions||""} onChange={e=>setR("site_conditions",e.target.value)}
              placeholder="e.g. Clear · 72°F · Wind 5mph · Humidity 45%" style={inp}/>
          </div>
          <div style={{marginBottom:12}}>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:6}}>
              <label style={lbl}>Description of Work Done</label>
              <button onClick={()=>startVoice("description")} title="Voice to text"
                style={{background:listening&&listenTarget==="description"?T.redLow:T.card,border:`1px solid ${listening&&listenTarget==="description"?T.red:T.border}`,borderRadius:8,padding:"5px 10px",color:listening&&listenTarget==="description"?T.red:T.muted,fontSize:13,cursor:"pointer",fontFamily:"inherit",animation:listening&&listenTarget==="description"?"pulse 1s infinite":undefined}}>
                {listening&&listenTarget==="description"?"🔴 Listening…":"🎤"}
              </button>
            </div>
            <textarea placeholder="Describe the work performed today… or tap 🎤 to speak" value={rpt.description||""} onChange={e=>setR("description",e.target.value)} rows={4} style={{...inp,resize:"vertical",lineHeight:1.5}}/>
          </div>
        </div>)}
        {step===2&&(<div><div style={{display:"flex",justifyContent:"space-between",marginBottom:12}}><div style={{fontSize:17,fontWeight:800}}>👷 Labor</div>{tot.labor>0&&can(user,"view_dashboard")&&<div style={{fontSize:16,fontWeight:800,color:T.green}}>${fmt(tot.labor)}</div>}</div>{rpt.labor.map((row,i)=><LaborCard key={row.id} row={row} onChange={r=>upd("labor",i,r)} onRemove={()=>del("labor",i)} division={project.division}/>)}<DashedAdd label="+ Add Worker" onClick={()=>add("labor",{id:uid(),name:"",classification:"",regHrs:"",otHrs:"",travelHrs:""})} color={T.orange}/></div>)}
        {step===3&&(<div>
          {/* Company Equipment */}
          <div style={{display:"flex",justifyContent:"space-between",marginBottom:10}}>
            <div style={{fontSize:17,fontWeight:800}}>🚜 Equipment</div>
            {tot.equip>0&&can(user,"view_dashboard")&&<div style={{fontSize:16,fontWeight:800,color:T.green}}>${fmt(tot.equip)}</div>}
          </div>
          {rpt.equipment.map((row,i)=><EquipCard key={row.id} row={row} onChange={r=>upd("equipment",i,r)} onRemove={()=>del("equipment",i)} division={project.division}/>)}
          <DashedAdd label="+ Add Company Equipment" onClick={()=>add("equipment",{id:uid(),description:"",qty:"",usage:"",rate:"",unit:""})} color={T.yellow}/>

          {/* Rented Equipment */}
          <div style={{marginTop:18,marginBottom:10,paddingTop:14,borderTop:`2px dashed ${T.border}`}}>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}>
              <div style={{fontSize:17,fontWeight:800}}>🔑 Rented Equipment</div>
              {(rpt.rental_equipment||[]).length>0&&can(user,"view_dashboard")&&
                <div style={{fontSize:16,fontWeight:800,color:T.green}}>
                  ${fmt((rpt.rental_equipment||[]).reduce((s,r)=>{const a=(parseFloat(r.qty)||0)*(parseFloat(r.rate)||0)*(parseFloat(r.usage)||1);return s+a;},0))}
                </div>}
            </div>
            <div style={{fontSize:12,color:T.muted,marginTop:2,marginBottom:10}}>Equipment you don't own — type any description</div>
          </div>
          {(rpt.rental_equipment||[]).map((row,i)=><RentedEquipCard key={row.id} row={row} onChange={r=>upd("rental_equipment",i,r)} onRemove={()=>del("rental_equipment",i)}/>)}
          <DashedAdd label="+ Add Rented Equipment" onClick={()=>add("rental_equipment",{id:uid(),description:"",qty:"",usage:"",rate:""})} color={T.purple}/>
        </div>)}
        {step===4&&(<div><div style={{fontSize:17,fontWeight:800,marginBottom:12}}>📦 Materials & Misc.</div>{rpt.materials.map((row,i)=><MatCard key={row.id} row={row} onChange={r=>upd("materials",i,r)} onRemove={()=>del("materials",i)}/>)}<DashedAdd label="+ Add Material / Item" onClick={()=>add("materials",{id:uid(),qty:"",description:"",amount:"",receipts:[]})} color={T.blue}/></div>)}
        {step===5&&(<div>
          <div style={{fontSize:17,fontWeight:800,marginBottom:16}}>📋 Site Notes</div>
          <div style={{...cardS,marginBottom:12,background:T.blueLow,border:`1px solid ${T.blue}30`}}>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}>
              <div>
                <div style={{fontSize:12,fontWeight:700,color:T.blue,textTransform:"uppercase",letterSpacing:"1px"}}>👷 Manpower On Site</div>
                <div style={{fontSize:28,fontWeight:900,color:T.text,marginTop:2}}>{rpt.labor?.length||0}</div>
                <div style={{fontSize:11,color:T.muted}}>workers from labor log</div>
              </div>
              <div style={{textAlign:"right"}}>
                <div style={{fontSize:11,color:T.muted}}>Total Hrs</div>
                <div style={{fontSize:22,fontWeight:800,color:T.green}}>{(rpt.labor||[]).reduce((s,l)=>(s+(parseFloat(l.regHrs)||0)+(parseFloat(l.otHrs)||0)+(parseFloat(l.travelHrs)||0)),0).toFixed(1)}</div>
              </div>
            </div>
          </div>
          <div style={{...cardS,marginBottom:12}}>
            <div style={{fontSize:14,fontWeight:800,color:T.text,marginBottom:12}}>🏗️ Visitor Log</div>
            {(rpt.visitor_log||[]).map((v,i)=>(
              <div key={i} style={{background:T.surface,borderRadius:10,padding:"10px 12px",marginBottom:8,display:"flex",justifyContent:"space-between",alignItems:"flex-start"}}>
                <div>
                  <div style={{fontSize:13,fontWeight:700,color:T.text}}>{v.name}</div>
                  <div style={{fontSize:11,color:T.muted}}>{v.company&&v.company+" · "}{v.type}</div>
                  {v.notes&&<div style={{fontSize:11,color:T.sub,marginTop:2,fontStyle:"italic"}}>{v.notes}</div>}
                </div>
                <button onClick={()=>del("visitor_log",i)} style={{background:"none",border:"none",color:T.red,cursor:"pointer",fontSize:16,padding:0}}>✕</button>
              </div>
            ))}
            <VisitorAddRow onAdd={v=>add("visitor_log",v)}/>
          </div>
          <div style={{...cardS,marginBottom:12}}>
            <div style={{fontSize:14,fontWeight:800,color:T.text,marginBottom:12}}>⚠️ Delays & Issues</div>
            {(rpt.delays||[]).map((d,i)=>(
              <div key={i} style={{background:T.surface,borderRadius:10,padding:"10px 12px",marginBottom:8,borderLeft:`3px solid ${T.red}`,display:"flex",justifyContent:"space-between",alignItems:"flex-start"}}>
                <div>
                  <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:2}}>
                    <span style={{...pill(T.red),fontSize:10}}>{d.cause}</span>
                    {d.hours>0&&<span style={{fontSize:11,color:T.muted}}>{d.hours}h delay</span>}
                  </div>
                  <div style={{fontSize:13,color:T.text}}>{d.description}</div>
                  {d.impact&&<div style={{fontSize:11,color:T.muted,marginTop:2}}>Impact: {d.impact}</div>}
                </div>
                <button onClick={()=>del("delays",i)} style={{background:"none",border:"none",color:T.red,cursor:"pointer",fontSize:16,padding:0}}>✕</button>
              </div>
            ))}
            <DelayAddRow onAdd={d=>add("delays",d)}/>
          </div>
        </div>)}

        {step===6&&(<div>
          <div style={{fontSize:17,fontWeight:800,marginBottom:12}}>✅ Review & Submit</div>
          <div style={{...cardS,marginBottom:12}}>
            <div style={{fontSize:11,color:divMeta.color,fontWeight:700,textTransform:"uppercase",letterSpacing:"1px",marginBottom:10}}>Summary</div>
            {[["Project",project.name],["Division",project.division],["Date",fmtDate(rpt.date)],["Report No.",rpt.report_no||"—"],["Workers",rpt.labor.length],["Equipment",rpt.equipment.length+" items"],["Rented Equip",(rpt.rental_equipment||[]).length+" items"],["Materials",rpt.materials.length+" items"]].map(([l,v])=>(<div key={l} style={{display:"flex",justifyContent:"space-between",padding:"7px 0",borderBottom:`1px solid ${T.border}`}}><span style={{fontSize:13,color:T.sub}}>{l}</span><span style={{fontSize:13,fontWeight:600}}>{v}</span></div>))}
            {[["Labor",tot.labor,T.green],["Equipment",tot.equip,T.green],["Materials",tot.mats,T.green]].map(([l,v,c])=>(<div key={l} style={{display:"flex",justifyContent:"space-between",padding:"7px 0",borderBottom:`1px solid ${T.border}`}}><span style={{fontSize:13,color:T.sub}}>{l}</span><span style={{fontSize:13,fontWeight:700,color:c}}>${fmt(v)}</span></div>))}
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",paddingTop:14}}><span style={{fontSize:16,fontWeight:800}}>Grand Total</span><span style={{fontSize:26,fontWeight:900,color:divMeta.color,letterSpacing:"-1px"}}>${fmt(tot.grand)}</span></div>
          </div>
        </div>)}
      </div>
      <div style={{position:"fixed",bottom:0,left:"50%",transform:"translateX(-50%)",width:"100%",maxWidth:480,background:T.bg+"EE",backdropFilter:"blur(12px)",borderTop:`1px solid ${T.border}`,padding:"12px 16px",display:"flex",gap:10}}>
        {step>1&&<button onClick={()=>{setStep(s=>s-1);scroll();}} style={{...ghostBtn,flex:1}}>← Back</button>}
        {step<6?<button onClick={()=>{setStep(s=>s+1);scroll();}} style={{...primBtn,flex:2,borderRadius:12,background:divMeta.color}}>{step===4?"Review →":"Next →"}</button>:<button onClick={submit} style={{...primBtn,flex:2,borderRadius:12,background:divMeta.color,opacity:saving?0.6:1}}>{saving?"Saving…":"💾 Save Report"}</button>}
      </div>
    </div>
  );
}

function printReport(report, project){
  const positions = getPositions(project.division);
  const tot = reportTotals(report, project.division);
  const [yr,mo,dy] = (report.date||'').split('-');
  const dateStr = `${mo}/${dy}/${yr}`;
  const fmt2 = n => (parseFloat(n)||0).toLocaleString('en-US',{style:'currency',currency:'USD',minimumFractionDigits:2});

  const laborRows = [...(report.labor||[]).filter(l=>l.classification!=='Per Diem')];
  while(laborRows.length<14) laborRows.push(null);
  const perDiemEntry = (report.labor||[]).find(l=>l.classification==='Per Diem');
  const equipRows = [...(report.equipment||[])];
  while(equipRows.length<15) equipRows.push(null);
  const mats = report.materials||[];
  const rentals = (report.rental_equipment||[]).map(r=>({
    qty:r.qty||'',
    description:r.description||'',
    amount:(parseFloat(r.qty)||0)*(parseFloat(r.rate)||0)*(parseFloat(r.usage)||1)
  }));

  function laborRow(lr,i){
    if(!lr) return `<tr><td></td><td colspan="2"></td><td></td><td></td><td></td><td></td><td></td></tr>`;
    const pos = positions.find(p=>p.name===lr.classification);
    return `<tr>
      <td>${lr.name||''}</td>
      <td colspan="2">${lr.classification||''}</td>
      <td class="num">${pos&&!pos.flat?lr.regHrs||0:''}</td>
      <td class="num">${pos&&!pos.flat?lr.otHrs||0:''}</td>
      <td class="num">${pos&&!pos.flat?lr.travelHrs||0:''}</td>
      <td class="num">${pos?pos.rate:''}</td>
      <td class="num amt">${fmt2(laborAmt(lr,project.division))}</td>
    </tr>`;
  }

  function equipRow(er){
    if(!er) return `<tr><td colspan="4"></td><td></td><td></td><td></td><td class="num amt">$0.00</td></tr>`;
    return `<tr>
      <td colspan="4">${er.description||''}</td>
      <td class="num">${er.qty||''}</td>
      <td class="num">${er.usage||''}</td>
      <td class="num">${er.rate||''}</td>
      <td class="num amt">${fmt2(equipAmt(er))}</td>
    </tr>`;
  }

  function matPair(left,right){
    return `
    <tr>
      <td class="num">${left?left.qty||'':''}</td>
      <td colspan="3">${left?left.description||'':''}</td>
      <td class="num amt">${left?fmt2(parseFloat(left.amount)||0):'$0.00'}</td>
      <td class="num">${right?right.qty||'':''}</td>
      <td colspan="2">${right?right.description||'':''}</td>
      <td class="num amt">${right?fmt2(parseFloat(right.amount)||0):'$0.00'}</td>
    </tr>
    <tr class="sub-row">
      <td></td><td colspan="3" class="sub-label">Tax</td><td class="num">$0.00</td>
      <td></td><td colspan="2" class="sub-label">Tax</td><td class="num">$0.00</td>
    </tr>
    <tr class="sub-row">
      <td></td><td colspan="3" class="sub-label">Total</td>
      <td class="num">${left?fmt2(parseFloat(left.amount)||0):'$0.00'}</td>
      <td></td><td colspan="2" class="sub-label">Total</td>
      <td class="num">${right?fmt2(parseFloat(right.amount)||0):'$0.00'}</td>
    </tr>`;
  }

  const html = `<!DOCTYPE html><html><head><meta charset="utf-8">
<title>AIME Daily Report — ${project.name} — ${dateStr}</title>
<style>
  @page { size: letter portrait; margin: 0.35in 0.3in; }
  * { box-sizing: border-box; margin: 0; padding: 0; font-family: Arial, sans-serif; }
  body { font-size: 7.5pt; color: #000; }
  table { width: 100%; border-collapse: collapse; }
  td, th { border: 1px solid #000; padding: 2px 3px; vertical-align: middle; font-size: 7.5pt; }
  .num { text-align: right; }
  .amt { font-weight: 600; }
  .section-header td { background: #d9e1f2; font-weight: bold; font-size: 8.5pt; text-align: center; padding: 3px; border: 2px solid #000; }
  .col-header td { background: #f2f2f2; font-weight: bold; font-size: 7pt; text-align: center; border: 2px solid #000; }
  .total-row td { background: #fff2cc; font-weight: bold; border: 2px solid #000; font-size: 8pt; }
  .grand-total td { background: #ffd966; font-weight: bold; border: 2px solid #000; font-size: 9pt; }
  .sub-row td { background: #fafafa; font-size: 6.5pt; }
  .sub-label { font-style: italic; color: #555; }
  .header-table td { border: 1px solid #000; padding: 3px 4px; }
  .title-row td { background: #1f3864; color: #fff; font-size: 11pt; font-weight: bold; text-align: center; padding: 5px; border: 2px solid #000; letter-spacing: 1px; }
  .sig-box { min-height: 40px; vertical-align: top; }
  @media print {
    body { -webkit-print-color-adjust: exact; print-color-adjust: exact; }
  }
</style></head><body>

<table class="header-table">
  <!-- Title -->
  <tr class="title-row"><td colspan="9">AIME DAILY REPORT</td></tr>

  <!-- Row 2: Customer / PO / Date -->
  <tr>
    <td style="width:8%;font-weight:bold">Customer:</td>
    <td colspan="2" style="width:22%">${project.client||''}</td>
    <td style="width:13%;font-weight:bold;text-align:center">Work Order / PO #</td>
    <td colspan="2" style="width:17%">${project.work_order||''}</td>
    <td style="width:10%;font-weight:bold;text-align:center">Report Date:</td>
    <td colspan="2" style="width:14%">${dateStr}</td>
  </tr>

  <!-- Row 3: Location -->
  <tr>
    <td style="font-weight:bold">Job Location:</td>
    <td colspan="8">${project.location||''}</td>
  </tr>

  <!-- Row 4: Job Number / Report # -->
  <tr>
    <td style="font-weight:bold">Job Number:</td>
    <td colspan="2">${project.name||''}</td>
    <td style="font-weight:bold;text-align:center">Report #:</td>
    <td colspan="5">${report.report_no||''}</td>
  </tr>

  <!-- Description -->
  <tr>
    <td style="font-weight:bold">Description of Work Done:</td>
    <td colspan="8" style="min-height:16px">${report.description||''}</td>
  </tr>
</table>

<br style="line-height:3px">

<!-- LABOR -->
<table>
  <tr class="section-header"><td colspan="8">LABOR</td></tr>
  <tr class="col-header">
    <td style="width:18%">NAME</td>
    <td colspan="2" style="width:22%">CLASSIFICATION</td>
    <td style="width:8%">REG. HRS.</td>
    <td style="width:8%">O.T. HRS.</td>
    <td style="width:8%">TRAVEL HRS.</td>
    <td style="width:12%">REGULAR RATE</td>
    <td style="width:12%">AMOUNT</td>
  </tr>
  ${laborRows.map((lr,i)=>laborRow(lr,i)).join('')}
  <!-- Per Diem row -->
  <tr>
    <td></td>
    <td colspan="2" style="font-style:italic">Per Diem</td>
    <td></td><td></td><td></td><td></td>
    <td class="num amt">${fmt2(perDiemEntry?laborAmt(perDiemEntry,project.division):0)}</td>
  </tr>
  <tr class="total-row">
    <td colspan="7" style="text-align:right;padding-right:6px">TOTAL LABOR</td>
    <td class="num">${fmt2(tot.labor)}</td>
  </tr>
</table>

<br style="line-height:3px">

<!-- EQUIPMENT -->
<table>
  <tr class="section-header"><td colspan="8">EQUIPMENT</td></tr>
  <tr class="col-header">
    <td colspan="4" style="width:46%">DESCRIPTION</td>
    <td style="width:8%">QUANTITY</td>
    <td style="width:8%">HOURS/DAYS</td>
    <td style="width:12%">RATE</td>
    <td style="width:12%">AMOUNT</td>
  </tr>
  ${equipRows.map(er=>equipRow(er)).join('')}
  <tr class="total-row">
    <td colspan="7" style="text-align:right;padding-right:6px">TOTAL EQUIPMENT</td>
    <td class="num">${fmt2(tot.equip)}</td>
  </tr>
</table>

<br style="line-height:3px">

<!-- RENTAL EQUIPMENT / MATERIALS -->
<table>
  <tr class="section-header"><td colspan="9">RENTAL EQUIPMENT</td></tr>
  <tr><td colspan="9" style="font-size:6.5pt;font-style:italic;background:#f9f9f9;border:1px solid #000">MATERIAL &amp; MISCELLANEOUS — LIST OF MATERIAL &amp; ATTACH SUPPORTING INVOICES</td></tr>
  <tr class="col-header">
    <td style="width:5%">QTY</td>
    <td colspan="3" style="width:27%">DESCRIPTION</td>
    <td style="width:10%">AMOUNT</td>
    <td style="width:5%">QTY</td>
    <td colspan="2" style="width:27%">DESCRIPTION</td>
    <td style="width:10%">AMOUNT</td>
  </tr>
  ${(rentals[0]||rentals[1])?matPair(rentals[0]||null,rentals[1]||null):''}
  ${(rentals[2]||rentals[3])?matPair(rentals[2]||null,rentals[3]||null):''}
  ${(rentals[4]||rentals[5])?matPair(rentals[4]||null,rentals[5]||null):''}
  ${(!rentals[0]&&!rentals[1]&&!rentals[2]&&!rentals[3]&&!rentals[4]&&!rentals[5])?matPair(null,null):''}
  <tr class="total-row">
    <td colspan="8" style="text-align:right;padding-right:6px">TOTAL RENTAL EQUIPMENT</td>
    <td class="num">${fmt2(tot.mats)}</td>
  </tr>
  <tr class="grand-total">
    <td colspan="8" style="text-align:right;padding-right:6px">GRAND TOTAL</td>
    <td class="num">${fmt2(tot.grand)}</td>
  </tr>
</table>

<br style="line-height:3px">

<!-- SIGNATURE -->
<table>
  <tr class="col-header">
    <td style="width:28%">ACCEPTED BY</td>
    <td style="width:12%">DATE</td>
    <td style="width:28%">CERTIFIED AS CORRECT BY CONTRACTOR'S REP</td>
    <td style="width:12%">DATE</td>
    <td style="width:20%">INSPECTOR SIGNATURE</td>
  </tr>
  <tr>
    <td class="sig-box" style="height:50px">${report.inspector_name||''}</td>
    <td class="sig-box">${report.inspector_signed_at?new Date(report.inspector_signed_at).toLocaleDateString():''}</td>
    <td class="sig-box"></td>
    <td class="sig-box"></td>
    <td class="sig-box" style="text-align:center">${report.inspector_signature?`<img src="${report.inspector_signature}" style="max-height:44px;max-width:100%;object-fit:contain;background:#fff;">`:''}</td>
  </tr>
</table>

</body></html>`;

  const win = window.open('','_blank','width=900,height=700');
  win.document.write(html);
  win.document.close();
  win.focus();
  setTimeout(()=>{win.focus();win.print();},1200);
}

function printReportWithOptions(report, project, sections, photos, photoLayout){
  const division = project.division;
  const positions = getPositions(division);
  const tot = reportTotals(report, division);
  const [yr,mo,dy] = (report.date||'').split('-');
  const dateStr = `${mo}/${dy}/${yr}`;
  const fmt2 = n => (parseFloat(n)||0).toLocaleString('en-US',{style:'currency',currency:'USD',minimumFractionDigits:2});
  const fmtH = n => `${(parseFloat(n)||0).toFixed(1)}h`;
  const visitors = report.visitor_log||[];
  const delays = report.delays||[];

  let photoHTML = '';
  if(photos&&photos.length>0){
    if(photoLayout==='full'){
      photoHTML = photos.map(ph=>`
        <div style="page-break-before:always;padding:20px;">
          <div style="font-size:9pt;color:#555;margin-bottom:8px;font-weight:700">${ph.category||'Photo'} ${ph.date?'· '+ph.date:''} ${ph.caption?'· '+ph.caption:''}</div>
          <img src="${ph.src}" style="width:100%;max-height:650px;object-fit:contain;display:block;border-radius:6px;border:1px solid #e5e7eb"/>
          ${ph.lat?`<div style="font-size:8pt;color:#6b7280;margin-top:6px">📍 GPS: ${ph.lat.toFixed(5)}, ${ph.lng.toFixed(5)}</div>`:''}
        </div>`).join('');
    }else{
      const pages = [];
      for(let i=0;i<photos.length;i+=4) pages.push(photos.slice(i,i+4));
      photoHTML = pages.map((pg,pi)=>`
        <div style="${pi>0?'page-break-before:always;':''}padding:20px;">
          <div style="font-size:10pt;font-weight:800;color:#1f3864;margin-bottom:10px;">📷 Site Photos</div>
          <div style="display:grid;grid-template-columns:1fr 1fr;gap:10px;">
            ${pg.map(ph=>`<div>
              <img src="${ph.src}" style="width:100%;height:200px;object-fit:cover;border-radius:6px;display:block;border:1px solid #e5e7eb"/>
              <div style="font-size:8pt;color:#555;margin-top:4px;">${ph.category||''} ${ph.caption?'· '+ph.caption:''}</div>
            </div>`).join('')}
          </div>
        </div>`).join('');
    }
  }

  const html = `<!DOCTYPE html><html><head><meta charset="utf-8">
<title>Daily Report — ${project.name} — ${dateStr}</title>
<style>
@page{size:letter portrait;margin:0.5in;}
*{box-sizing:border-box;margin:0;padding:0;font-family:Arial,sans-serif;}
body{font-size:10pt;color:#000;}
.header{display:flex;justify-content:space-between;align-items:flex-start;padding-bottom:12px;border-bottom:3px solid #1f3864;margin-bottom:16px;}
.co{font-size:20pt;font-weight:900;color:#1f3864;}
.co-sub{font-size:9pt;color:#555;margin-top:3px;}
.doc-title{text-align:right;}
.doc-title h1{font-size:18pt;font-weight:900;color:#1f3864;}
.proj-box{background:#f0f4ff;border:1px solid #c7d2fe;border-radius:6px;padding:10px 14px;margin-bottom:14px;display:grid;grid-template-columns:1fr 1fr 1fr;gap:10px;}
.fl{font-size:7.5pt;font-weight:700;text-transform:uppercase;letter-spacing:0.5px;color:#6b7280;margin-bottom:2px;}
.fv{font-size:10pt;font-weight:600;color:#111;}
.section{margin-bottom:14px;}
.section h2{font-size:9pt;font-weight:800;text-transform:uppercase;letter-spacing:1px;color:#1f3864;border-bottom:1.5px solid #1f3864;padding-bottom:5px;margin-bottom:3px;}
.desc-box{background:#f9fafb;border:1px solid #e5e7eb;border-radius:6px;padding:10px;font-size:10pt;line-height:1.7;min-height:48px;}
table{width:100%;border-collapse:collapse;margin-bottom:6px;font-size:9pt;}
th{background:#1f3864;color:#fff;padding:5px 8px;text-align:left;font-size:8pt;}
td{padding:5px 8px;border-bottom:1px solid #e5e7eb;}
tr:nth-child(even) td{background:#f9fafb;}
.total-row td{background:#EEF2FF;font-weight:700;border-top:1.5px solid #1f3864;}
.badge{display:inline-block;padding:2px 8px;border-radius:10px;font-size:8pt;font-weight:700;}
.badge-red{background:#fee2e2;color:#991b1b;}
.badge-yellow{background:#fef9c3;color:#713f12;}
.badge-blue{background:#dbeafe;color:#1e40af;}
.visitor-row{padding:7px;border-bottom:1px solid #e5e7eb;display:flex;gap:10px;}
.delay-row{padding:7px;border-bottom:1px solid #e5e7eb;border-left:3px solid #ef4444;padding-left:10px;}
.sig-section{margin-top:20px;display:grid;grid-template-columns:1fr 1fr;gap:20px;}
.sig-box{border-top:1.5px solid #000;padding-top:8px;}
.sig-label{font-size:8pt;color:#555;text-transform:uppercase;}
.footer{margin-top:16px;padding-top:8px;border-top:1px solid #e5e7eb;font-size:7.5pt;color:#9ca3af;display:flex;justify-content:space-between;}
@media print{body{-webkit-print-color-adjust:exact;print-color-adjust:exact;}}
</style></head><body>

<div class="header">
  <div><div class="co">AIME</div><div class="co-sub">Atlantic Industrial Mechanical & Environmental Inc.<br/>${project.division||''} Division</div></div>
  <div class="doc-title"><h1>Daily Field Report</h1><div style="font-size:11pt;font-weight:700;color:#555">Report #${report.report_no||'—'}</div></div>
</div>

<div class="proj-box">
  <div><div class="fl">Project</div><div class="fv">${project.name||'—'}</div></div>
  <div><div class="fl">Client</div><div class="fv">${project.client||'—'}</div></div>
  <div><div class="fl">Date</div><div class="fv">${dateStr}</div></div>
  <div><div class="fl">AFE / PO</div><div class="fv">${project.afe||'—'}${project.work_order?' / '+project.work_order:''}</div></div>
  <div><div class="fl">Submitted By</div><div class="fv">${report.submitted_by||'—'}</div></div>
  <div><div class="fl">Status</div><div class="fv"><span class="badge ${report.status==='approved'?'badge-blue':report.status==='flagged'?'badge-red':'badge-yellow'}">${(report.status||'submitted').toUpperCase()}</span></div></div>
</div>

${sections.weather&&report.site_conditions?`<div class="section"><h2>🌤️ Site Conditions / Weather</h2><div class="desc-box">${report.site_conditions}</div></div>`:''}

${sections.description&&report.description?`<div class="section"><h2>📝 Description of Work</h2><div class="desc-box">${report.description.replace(/\n/g,'<br/>')}</div></div>`:''}

${sections.labor&&(report.labor||[]).length>0?`<div class="section"><h2>👷 Labor — ${(report.labor||[]).length} Workers · ${fmtH(tot.labor_hrs||0)} Total Hrs${tot.labor>0?' · '+fmt2(tot.labor):''}</h2>
<table><thead><tr><th>Name</th><th>Classification</th><th style="text-align:center">Reg Hrs</th><th style="text-align:center">OT Hrs</th><th style="text-align:center">Travel</th><th style="text-align:right">Amount</th></tr></thead>
<tbody>${(report.labor||[]).filter(l=>l.name||l.classification||parseFloat(l.regHrs)||parseFloat(l.otHrs)||parseFloat(l.travelHrs)).map(l=>`<tr><td>${l.name||'—'}</td><td>${l.classification||'—'}</td><td style="text-align:center">${l.regHrs||0}</td><td style="text-align:center">${l.otHrs||0}</td><td style="text-align:center">${l.travelHrs||0}</td><td style="text-align:right">${fmt2(laborAmt(l,division))}</td></tr>`).join('')}
</tbody><tfoot><tr class="total-row"><td colspan="5"><strong>TOTAL LABOR</strong></td><td style="text-align:right"><strong>${fmt2(tot.labor||0)}</strong></td></tr></tfoot></table></div>`:''}

${sections.equipment&&(report.equipment||[]).length>0?`<div class="section"><h2>🚜 Equipment — ${(report.equipment||[]).length} Items${tot.equip>0?' · '+fmt2(tot.equip):''}</h2>
<table><thead><tr><th>Equipment</th><th>Unit</th><th style="text-align:center">Qty</th><th style="text-align:center">Hrs / Days</th><th style="text-align:right">Rate</th><th style="text-align:right">Amount</th></tr></thead>
<tbody>${(report.equipment||[]).filter(e=>e.description||parseFloat(e.qty)).map(e=>{const rate=parseFloat(e.rate)||(getEquipList(division).find(x=>!x.section&&x.name===e.description)||{}).rate||0;return `<tr><td>${e.description||'—'}</td><td>${e.unit||'—'}</td><td style="text-align:center">${e.qty||0}</td><td style="text-align:center">${e.usage||'—'}</td><td style="text-align:right">${rate?fmt2(rate):'—'}</td><td style="text-align:right">${fmt2(equipAmt(e,division))}</td></tr>`;}).join('')}
</tbody><tfoot><tr class="total-row"><td colspan="5"><strong>TOTAL EQUIPMENT</strong></td><td style="text-align:right"><strong>${fmt2(tot.equip||0)}</strong></td></tr></tfoot></table></div>`:''}

${sections.rental&&(report.rental_equipment||[]).length>0?`<div class="section"><h2>🔧 Rental Equipment</h2>
<table><thead><tr><th>Description</th><th>Company</th><th style="text-align:center">Hours</th><th style="text-align:right">Amount</th></tr></thead>
<tbody>${(report.rental_equipment||[]).map(r=>`<tr><td>${r.description||'—'}</td><td>${r.vendor||'—'}</td><td style="text-align:center">${r.hours||'—'}</td><td style="text-align:right">${r.amount?fmt2(r.amount):'—'}</td></tr>`).join('')}
</tbody></table></div>`:''}

${sections.materials&&(report.materials||[]).length>0?`<div class="section"><h2>📦 Materials & Misc.</h2>
<table><thead><tr><th>Description</th><th style="text-align:center">Qty</th><th style="text-align:right">Amount</th></tr></thead>
<tbody>${(report.materials||[]).map(m=>`<tr><td>${m.description||'—'}</td><td style="text-align:center">${m.qty||'—'}</td><td style="text-align:right">${m.amount?fmt2(m.amount):'—'}</td></tr>`).join('')}
</tbody></table></div>`:''}

${sections.visitors&&visitors.length>0?`<div class="section"><h2>🏗️ Visitor Log — ${visitors.length} Visitor${visitors.length!==1?'s':''}</h2>
${visitors.map(v=>`<div class="visitor-row"><div style="flex:1"><strong>${v.name||'—'}</strong>${v.company?' · '+v.company:''}</div><div style="font-size:8pt;color:#555">${v.type||''}</div>${v.notes?`<div style="font-size:8pt;color:#6b7280;font-style:italic">${v.notes}</div>`:''}</div>`).join('')}</div>`:''}

${sections.delays&&delays.length>0?`<div class="section"><h2>⚠️ Delays & Issues — ${delays.length} Item${delays.length!==1?'s':''}</h2>
${delays.map(d=>`<div class="delay-row"><div style="display:flex;gap:10px;align-items:center;margin-bottom:3px"><strong>${d.cause||'—'}</strong>${d.hours>0?`<span style="font-size:8pt;color:#ef4444">${d.hours}h delay</span>`:''}</div><div>${d.description||''}</div>${d.impact?`<div style="font-size:8pt;color:#555">Impact: ${d.impact}</div>`:''}</div>`).join('')}</div>`:''}

${sections.signature&&report.inspector_signature?`<div class="section"><h2>✍️ Inspector Sign-Off</h2>
<div style="background:#fff;border:1px solid #86efac;border-radius:6px;padding:10px;display:flex;align-items:center;gap:16px">
<div style="background:#ffffff;border:1px solid #ccc;border-radius:4px;padding:6px;display:inline-block;overflow:hidden;">
<img src="${report.inspector_signature}" style="max-height:70px;max-width:240px;object-fit:contain;display:block;filter:none;"/>
</div>
<div><div style="font-weight:700">${report.inspector_name||'Inspector'}</div><div style="font-size:9pt;color:#555">${report.inspector_signed_at?new Date(report.inspector_signed_at).toLocaleString():''}</div></div>
</div></div>`:''}

<div class="sig-section">
  <div class="sig-box"><div style="height:50px"></div><div class="sig-label">Foreman / Submitted By</div><div style="font-weight:700;margin-top:4px">${report.submitted_by||''}</div></div>
  <div class="sig-box"><div style="height:50px"></div><div class="sig-label">PM / Reviewed By</div><div style="margin-top:4px">Date: ______________</div></div>
</div>

<div class="footer"><span>AIME Field Pro · ${project.name} · Report #${report.report_no||'—'} · ${dateStr}</span><span>Generated: ${new Date().toLocaleString()}</span></div>

${photoHTML}

</body></html>`;

  const win = window.open('','_blank','width=950,height=800');
  if(!win){alert('Popup blocked — please allow popups and try again.');return;}
  win.document.write(html);
  win.document.close();
  win.focus();
  setTimeout(()=>{win.focus();win.print();},1200);
}

function SignaturePad({onSave,onCancel,reportName}){
  const [name,setName]=useState("");
  const [company,setCompany]=useState("");
  const [role,setRole]=useState("");
  const [font,setFont]=useState("'Dancing Script', cursive");
  const [saving,setSaving]=useState(false);
  const [err,setErr]=useState("");
  const previewRef=useRef(null);

  const FONTS=[
    {label:"Classic",value:"'Dancing Script', cursive"},
    {label:"Elegant",value:"'Great Vibes', cursive"},
    {label:"Bold",value:"'Pacifico', cursive"},
    {label:"Formal",value:"'Pinyon Script', cursive"},
  ];

  // Load Google Fonts
  useEffect(()=>{
    const link=document.createElement("link");
    link.rel="stylesheet";
    link.href="https://fonts.googleapis.com/css2?family=Dancing+Script:wght@700&family=Great+Vibes&family=Pacifico&family=Pinyon+Script&display=swap";
    document.head.appendChild(link);
    return()=>document.head.removeChild(link);
  },[]);

  async function save(){
    if(!name.trim()){setErr("Please enter your name.");return;}
    setSaving(true);setErr("");
    try{
      // Render the typed signature to a canvas for storage
      const canvas=document.createElement("canvas");
      canvas.width=600;canvas.height=200;
      const ctx=canvas.getContext("2d");
      // White background
      ctx.fillStyle="#ffffff";
      ctx.fillRect(0,0,600,200);
      // Draw name in chosen font
      ctx.fillStyle="#000000";
      ctx.font=`bold 72px ${font}`;
      ctx.textAlign="center";
      ctx.textBaseline="middle";
      ctx.fillText(name.trim(),300,100);
      const sigData=canvas.toDataURL("image/png");
      await onSave(`${name.trim()}${company?" · "+company:""}${role?" · "+role:""}`,sigData);
    }catch(e){setErr("Error: "+e.message);}
    setSaving(false);
  }

  return(
    <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.9)",zIndex:300,display:"flex",alignItems:"center",justifyContent:"center",padding:16,fontFamily:"inherit"}}>
      <div style={{background:"#fff",borderRadius:16,padding:24,width:"100%",maxWidth:440,color:"#111"}}>
        <div style={{fontSize:16,fontWeight:900,color:"#1f3864",marginBottom:4}}>✍️ Inspector Sign-Off</div>
        <div style={{fontSize:12,color:"#666",marginBottom:20}}>{reportName||"Daily Field Report"}</div>

        {err&&<div style={{background:"#fee2e2",border:"1px solid #fca5a5",borderRadius:8,padding:"8px 12px",marginBottom:12,fontSize:12,color:"#dc2626"}}>{err}</div>}

        {/* Name input */}
        <div style={{marginBottom:12}}>
          <label style={{display:"block",fontSize:11,fontWeight:700,color:"#555",textTransform:"uppercase",letterSpacing:"1px",marginBottom:6}}>Full Name *</label>
          <input value={name} onChange={e=>setName(e.target.value)}
            placeholder="e.g. John Smith"
            style={{width:"100%",padding:"10px 12px",fontSize:15,border:"2px solid #e5e7eb",borderRadius:8,fontFamily:"inherit",boxSizing:"border-box",outline:"none"}}/>
        </div>

        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:16}}>
          <div>
            <label style={{display:"block",fontSize:11,fontWeight:700,color:"#555",textTransform:"uppercase",letterSpacing:"1px",marginBottom:6}}>Company</label>
            <input value={company} onChange={e=>setCompany(e.target.value)}
              placeholder="Optional"
              style={{width:"100%",padding:"8px 10px",fontSize:13,border:"2px solid #e5e7eb",borderRadius:8,fontFamily:"inherit",boxSizing:"border-box"}}/>
          </div>
          <div>
            <label style={{display:"block",fontSize:11,fontWeight:700,color:"#555",textTransform:"uppercase",letterSpacing:"1px",marginBottom:6}}>Title / Role</label>
            <input value={role} onChange={e=>setRole(e.target.value)}
              placeholder="Optional"
              style={{width:"100%",padding:"8px 10px",fontSize:13,border:"2px solid #e5e7eb",borderRadius:8,fontFamily:"inherit",boxSizing:"border-box"}}/>
          </div>
        </div>

        {/* Font picker */}
        <div style={{marginBottom:14}}>
          <label style={{display:"block",fontSize:11,fontWeight:700,color:"#555",textTransform:"uppercase",letterSpacing:"1px",marginBottom:8}}>Signature Style</label>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8}}>
            {FONTS.map(f=>(
              <button key={f.value} onClick={()=>setFont(f.value)}
                style={{padding:"8px 10px",border:`2px solid ${font===f.value?"#1f3864":"#e5e7eb"}`,borderRadius:8,background:font===f.value?"#f0f4ff":"#fff",cursor:"pointer",fontFamily:f.value,fontSize:18,color:"#111",textAlign:"center"}}>
                {name||"Signature"}
              </button>
            ))}
          </div>
        </div>

        {/* Live preview */}
        {name&&<div style={{background:"#f9fafb",border:"2px solid #e5e7eb",borderRadius:10,padding:"16px 20px",marginBottom:16,textAlign:"center"}}>
          <div style={{fontFamily:font,fontSize:48,color:"#000",lineHeight:1.2,minHeight:60}}>{name}</div>
          <div style={{width:"80%",height:2,background:"#000",margin:"8px auto 0"}}/>
          <div style={{fontSize:11,color:"#666",marginTop:4}}>{name.trim()}{company?" · "+company:""}{role?" · "+role:""}</div>
        </div>}

        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}>
          <button onClick={save} disabled={!name.trim()||saving}
            style={{padding:"12px",background:name.trim()&&!saving?"#1f3864":"#9ca3af",color:"#fff",border:"none",borderRadius:10,fontSize:14,fontWeight:700,cursor:name.trim()&&!saving?"pointer":"default",fontFamily:"inherit"}}>
            {saving?"Saving…":"✓ Sign Report"}
          </button>
          <button onClick={onCancel}
            style={{padding:"12px",background:"#f3f4f6",color:"#374151",border:"2px solid #e5e7eb",borderRadius:10,fontSize:14,fontWeight:700,cursor:"pointer",fontFamily:"inherit"}}>
            Cancel
          </button>
        </div>
      </div>
    </div>
  );
}


function ReportDetail({report:initReport,project,user,onBack,onDelete,onApprove,onFlag}){
  const [report,setReport]=useState(initReport);
  const [lb,setLb]=useState(null);const [flagNote,setFlagNote]=useState("");const [flagging,setFlagging]=useState(false);
  const [showSigPad,setShowSigPad]=useState(false);const [sigSaving,setSigSaving]=useState(false);
  const [showInspectorShare,setShowInspectorShare]=useState(false);
  const [inspLinkCopied,setInspLinkCopied]=useState(false);
  const [showEsigModal,setShowEsigModal]=useState(false);
  const [esigEmail,setEsigEmail]=useState("");
  const [esigName,setEsigName]=useState("");
  const [esigSending,setEsigSending]=useState(false);
  const [esigError,setEsigError]=useState("");
  const [esigSent,setEsigSent]=useState(false);
  const [showPrintModal,setShowPrintModal]=useState(false);
  const [reportPhotos,setReportPhotos]=useState([]);
  const [selectedPhotos,setSelectedPhotos]=useState([]);
  const [photosLoading,setPhotosLoading]=useState(false);
  const [printSections,setPrintSections]=useState({
    weather:true,description:true,labor:true,equipment:true,
    rental:true,materials:true,visitors:true,delays:true,signature:true
  });
  const [photoLayout,setPhotoLayout]=useState("grid"); // grid | full
  const [editing,setEditing]=useState(false);const [editErr,setEditErr]=useState("");const [editSaving,setEditSaving]=useState(false);
  const [editData,setEditData]=useState(null);

  async function saveEdit(updated){
    setEditSaving(true);setEditErr("");
    try{
      await API.reports.update(report.id,{...updated,status:"submitted",updated_at:new Date().toISOString()});
      setReport(r=>({...r,...updated,status:"submitted"}));
      setEditing(false);
    }catch(e){setEditErr(e.message);}
    setEditSaving(false);
  }
  const tot=reportTotals(report,project.division);
  const sc={submitted:T.yellow,approved:T.green,flagged:T.red,signed:T.green}[report.status]||T.muted;
  const divColor=DIV_META[project.division]?.color||T.orange;

  async function loadPhotosForPrint(){
    setPhotosLoading(true);
    try{
      const p=await API.photos.forProject(project.id);
      const list=Array.isArray(p)?p:[];
      const onDate=list.filter(ph=>ph.date===report.date);
      const others=list.filter(ph=>ph.date!==report.date);
      setReportPhotos([...onDate,...others]);
      setSelectedPhotos(onDate.map(ph=>ph.id));
    }catch(e){}
    setPhotosLoading(false);
  }

  async function sendEsig(){
    if(!esigEmail.trim()||!esigName.trim()){setEsigError("Please enter inspector name and email.");return;}
    setEsigSending(true);setEsigError("");
    try{
      const tot=reportTotals(report,project.division);
      const div=project.division;
      const m=(n)=>"$"+Number(n||0).toLocaleString("en-US",{minimumFractionDigits:2,maximumFractionDigits:2});
      const res=await fetch("/.netlify/functions/box-sign-create",{
        method:"POST",
        headers:{"Content-Type":"application/json"},
        body:JSON.stringify({
          docType:"daily",
          reportId:report.id,
          inspectorEmail:esigEmail.trim(),
          inspectorName:esigName.trim(),
          projectName:project.name,
          customer:project.client||"",
          location:project.location||"",
          description:report.description||"",
          reportNo:report.report_no||"",
          reportDate:report.date||"",
          submittedBy:report.submitted_by||"",
          lineItems:{
            labor:(report.labor||[]).map(l=>({
              name:l.name||"",
              classification:l.classification||"",
              hours:((parseFloat(l.regHrs)||0)+(parseFloat(l.otHrs)||0)+(parseFloat(l.travelHrs)||0)).toFixed(1),
              rate:"",
              amount:m(laborAmt(l,div)),
            })),
            equipment:(report.equipment||[]).map(e=>({
              description:e.description||"",
              unit:e.unit||"",
              qty:e.qty||0,
              rate:e.rate?m(e.rate):"",
              amount:m(equipAmt(e,div)),
            })),
            rental:(report.rental_equipment||[]).map(r=>({
              description:r.description||"",
              qty:r.qty||0,
              rate:m(r.rate),
              amount:m((parseFloat(r.qty)||0)*(parseFloat(r.rate)||0)*(parseFloat(r.usage)||1)),
            })),
            materials:(report.materials||[]).map(x=>({
              description:x.description||"",
              qty:x.qty||"",
              unit_price:"",
              amount:m(x.amount),
            })),
            other:[],
          },
          laborTotal:m(tot.labor),
          equipmentTotal:m(tot.equip),
          rentalTotal:m(tot.rental),
          materialsTotal:m(tot.mats),
          grandTotal:m(tot.grand),
        }),
      });
      const data=await res.json();
      if(!res.ok)throw new Error(data.error||"Failed to send");
      // Save request ID to report
      await API.reports.update(report.id,{hellosign_request_id:data.requestId,hellosign_status:"pending"});
      setEsigSent(true);
      setShowEsigModal(false);
    }catch(e){setEsigError("Error: "+e.message);}
    setEsigSending(false);
  }

  async function saveSignature(inspectorName,sigData){
    setSigSaving(true);
    try{
      if(!report||!report.id){
        setSigSaving(false);
        throw new Error("Report ID missing. Please go back and reopen this report.");
      }
      const compressSig=()=>new Promise(res=>{
        const img=new Image();
        img.onload=()=>{
          const c=document.createElement("canvas");
          c.width=Math.min(img.width,800);
          c.height=Math.round(img.height*(c.width/img.width));
          c.getContext("2d").drawImage(img,0,0,c.width,c.height);
          res(c.toDataURL("image/jpeg",0.5));
        };
        img.src=sigData;
      });
      const compressedSig=await compressSig();
      await API.reports.update(report.id,{
        inspector_name:inspectorName,
        inspector_signature:compressedSig,
        inspector_signed_at:new Date().toISOString(),
        status:"signed",
      });
      setReport(r=>({...r,inspector_name:inspectorName,inspector_signature:compressedSig,inspector_signed_at:new Date().toISOString(),status:"signed"}));
      setShowSigPad(false);
    }catch(e){
      const msg=e.message||"Could not save signature.";
      setErr("Signature save failed: "+msg);
      throw new Error(msg);
    }finally{
      setSigSaving(false);
    }
  }

  function exportXLSX(){
    const wb = XLSX.read(DAILY_REPORT_TEMPLATE_B64, {type:'base64', cellStyles:true, cellFormula:true});
    const ws = wb.Sheets['3-24-2026'];

    Object.keys(ws).forEach(addr=>{
      if(addr.startsWith('!'))return;
      const cell=ws[addr];
      if(cell&&cell.f){delete cell.f; cell.v=cell.v||0;}
    });

    function sc(addr, val){
      const existing=ws[addr]||{};
      const s=existing.s||{};
      if(typeof val==='number'){
        ws[addr]={...existing, s, t:'n', v:val, w:undefined, f:undefined};
      } else {
        ws[addr]={...existing, s, t:'s', v:val==null?'':String(val), w:undefined, f:undefined};
      }
    }
    function scn(addr, val){
      const existing=ws[addr]||{};
      const s=existing.s||{};
      const n=parseFloat(val)||0;
      ws[addr]={...existing, s, t:'n', v:n, z:'"$"#,##0.00', w:undefined, f:undefined};
    }
    function scn0(addr){
      const existing=ws[addr]||{};
      if(!existing.v) scn(addr,0);
    }

    const positions=getPositions(project.division);
    const[yr,mo,dy]=(report.date||'').split('-');
    const dateStr=`${mo}/${dy}/${yr}`;
    const tot=reportTotals(report,project.division);

    sc('D3', project.client||'');
    sc('H3', project.work_order||'');
    sc('J3', dateStr);
    sc('D4', project.location||'');
    sc('C5', project.name||'');
    sc('F5', report.report_no||'');
    sc('B7', report.description||'');

    const laborRows=[...(report.labor||[]).filter(l=>l.classification!=='Per Diem')];
    while(laborRows.length<14) laborRows.push(null);
    laborRows.slice(0,14).forEach((lr,i)=>{
      const row=10+i;
      if(lr){
        const pos=positions.find(p=>p.name===lr.classification);
        sc(`B${row}`, lr.name||'');
        sc(`D${row}`, lr.classification||'');
        if(pos&&!pos.flat){
          sc(`F${row}`, parseFloat(lr.regHrs)||0);
          sc(`G${row}`, parseFloat(lr.otHrs)||0);
          sc(`H${row}`, parseFloat(lr.travelHrs)||0);
        } else {
          sc(`F${row}`, 0); sc(`G${row}`, 0); sc(`H${row}`, 0);
        }
        sc(`I${row}`, pos?pos.rate:0);
        scn(`J${row}`, laborAmt(lr,project.division));
      } else {
        sc(`B${row}`,''); sc(`D${row}`,'');
        sc(`F${row}`,0); sc(`G${row}`,0); sc(`H${row}`,0);
        sc(`I${row}`,0); scn(`J${row}`,0);
      }
    });

    const perDiemEntry=(report.labor||[]).find(l=>l.classification==='Per Diem');
    scn('J24', perDiemEntry?laborAmt(perDiemEntry,project.division):0);

    scn('J25', tot.labor);

    const equipRows=[...(report.equipment||[])];
    while(equipRows.length<15) equipRows.push(null);
    equipRows.slice(0,15).forEach((er,i)=>{
      const row=29+i;
      if(er){
        sc(`B${row}`, er.description||'');
        sc(`G${row}`, parseFloat(er.qty)||0);
        sc(`H${row}`, parseFloat(er.usage)||0);
        sc(`I${row}`, parseFloat(er.rate)||0);
        scn(`J${row}`, equipAmt(er));
      } else {
        sc(`B${row}`,'');
        sc(`G${row}`,0); sc(`H${row}`,0); sc(`I${row}`,0);
        scn(`J${row}`,0);
      }
    });

    scn('J44', tot.equip);

    const mats=report.materials||[];
    const m0=mats[0]||null;const m1=mats[1]||null;
    sc('B49',m0?m0.qty||'':''); sc('C49',m0?m0.description||'':'');
    scn('F49',m0?parseFloat(m0.amount)||0:0); scn('F52',m0?parseFloat(m0.amount)||0:0);
    sc('G49',m1?m1.qty||'':''); sc('H49',m1?m1.description||'':'');
    scn('J49',m1?parseFloat(m1.amount)||0:0); scn('J52',m1?parseFloat(m1.amount)||0:0);
    scn0('F50'); scn0('J50');
    const m2=mats[2]||null;const m3=mats[3]||null;
    sc('B53',m2?m2.qty||'':''); sc('C53',m2?m2.description||'':'');
    scn('F54',m2?parseFloat(m2.amount)||0:0); scn('F56',m2?parseFloat(m2.amount)||0:0);
    sc('G53',m3?m3.qty||'':''); sc('H53',m3?m3.description||'':'');
    scn('J54',m3?parseFloat(m3.amount)||0:0); scn('J56',m3?parseFloat(m3.amount)||0:0);
    scn0('F53'); scn0('J53');
    const m4=mats[4]||null;const m5=mats[5]||null;
    sc('B57',m4?m4.qty||'':''); sc('C57',m4?m4.description||'':'');
    scn('F58',m4?parseFloat(m4.amount)||0:0); scn('F60',m4?parseFloat(m4.amount)||0:0);
    sc('G57',m5?m5.qty||'':''); sc('H57',m5?m5.description||'':'');
    scn('J58',m5?parseFloat(m5.amount)||0:0); scn('J60',m5?parseFloat(m5.amount)||0:0);
    scn0('F57'); scn0('J57');

    scn('J61', tot.mats);
    scn('J62', tot.grand);

    sc('D64', dateStr);
    if(report.inspector_name){
      sc('B64', report.inspector_name);
      sc('G64', report.inspector_signed_at?new Date(report.inspector_signed_at).toLocaleDateString():'');
    }

    const shIdx=wb.SheetNames.indexOf('3-24-2026');
    if(shIdx>=0){wb.SheetNames[shIdx]='Daily Report';wb.Sheets['Daily Report']=ws;delete wb.Sheets['3-24-2026'];}

    if(ws['!pageSetup']){
      ws['!pageSetup'].fitToPage=true;
      ws['!pageSetup'].fitToWidth=1;
      ws['!pageSetup'].fitToHeight=1;
      ws['!pageSetup'].scale=undefined;
      ws['!pageSetup'].orientation='portrait';
      ws['!pageSetup'].paperSize=1;
    } else {
      ws['!pageSetup']={fitToPage:true,fitToWidth:1,fitToHeight:1,orientation:'portrait',paperSize:1};
    }
    ws['!sheetPr']={...(ws['!sheetPr']||{}),pageSetUpPr:{fitToPage:true}};
    ws['!margins']={left:0.25,right:0.25,top:0.5,bottom:0.5,header:0.3,footer:0.3};
    XLSX.writeFile(wb, `AIME_${(project.name||'').replace(/\s+/g,'_')}_${(report.date||'').replace(/-/g,'')}.xlsx`,
      {cellStyles:true, bookSST:false});
  }
  if(editing&&editData){
    return(
      <div style={{background:T.bg,minHeight:"100vh",fontFamily:"inherit",color:T.text}}>
        <TopBar title="Edit Report" onBack={()=>setEditing(false)}/>
        <div style={{padding:"14px 16px 100px"}}>
          <ErrBanner msg={editErr} onDismiss={()=>setEditErr("")}/>
          <div style={{...cardS,marginBottom:12,background:T.yellowLow,border:`1px solid ${T.yellow}40`}}>
            <div style={{fontSize:12,color:T.yellow}}>⚠️ Editing this report will reset its status to Submitted and require re-approval.</div>
          </div>
          <div style={{marginBottom:12}}><label style={lbl}>Date</label><input type="date" value={editData.date||""} onChange={e=>setEditData(d=>({...d,date:e.target.value}))} style={inp}/></div>
          <div style={{marginBottom:12}}><label style={lbl}>Report No.</label><input type="text" value={editData.report_no||""} onChange={e=>setEditData(d=>({...d,report_no:e.target.value}))} style={inp}/></div>
          <div style={{marginBottom:20}}><label style={lbl}>Description of Work Done</label><textarea rows={4} value={editData.description||""} onChange={e=>setEditData(d=>({...d,description:e.target.value}))} style={{...inp,resize:"vertical",lineHeight:1.5}}/></div>
          <div style={{fontSize:13,color:T.muted,marginBottom:12}}>To edit labor, equipment, or materials in detail — delete this report and create a new one.</div>
          <button onClick={()=>saveEdit(editData)} style={{...primBtn,opacity:editSaving?0.6:1,borderRadius:14}}>{editSaving?"Saving…":"Save Changes"}</button>
        </div>
      </div>
    );
  }

  if(showPrintModal){
    const SECTIONS=[
      {key:"description",label:"📝 Work Description"},
      {key:"weather",label:"🌤️ Site Conditions / Weather"},
      {key:"labor",label:"👷 Labor"},
      {key:"equipment",label:"🚜 Equipment"},
      {key:"rental",label:"🔧 Rental Equipment"},
      {key:"materials",label:"📦 Materials"},
      {key:"visitors",label:"🏗️ Visitor Log"},
      {key:"delays",label:"⚠️ Delays & Issues"},
      {key:"signature",label:"✍️ Inspector Signature"},
    ];
    const toggleSection=(k)=>setPrintSections(s=>({...s,[k]:!s[k]}));
    const togglePhoto=(id)=>setSelectedPhotos(s=>s.includes(id)?s.filter(x=>x!==id):[...s,id]);
    const selAll=()=>setSelectedPhotos(reportPhotos.map(p=>p.id));
    const selNone=()=>setSelectedPhotos([]);
    const selDateOnly=()=>setSelectedPhotos(reportPhotos.filter(p=>p.date===report.date).map(p=>p.id));

    const catColor={Progress:T.blue,Safety:T.red,Equipment:T.yellow,"Issue/Deficiency":T.red,Before:T.purple,After:T.green,Inspection:T.orange,Other:T.muted};

    function generate(){
      const photos=reportPhotos.filter(p=>selectedPhotos.includes(p.id));
      printReportWithOptions(report,project,printSections,photos,photoLayout);
      setShowPrintModal(false);
    }

    return(
      <div style={{background:T.bg,minHeight:"100vh",fontFamily:"inherit"}}>
        <div style={{background:T.surface,borderBottom:`1px solid ${T.border}`,padding:"14px 16px",position:"sticky",top:0,zIndex:50,display:"flex",justifyContent:"space-between",alignItems:"center"}}>
          <button onClick={()=>setShowPrintModal(false)} style={{background:"none",border:"none",color:T.sub,fontSize:13,cursor:"pointer",fontFamily:"inherit"}}>← Back</button>
          <div style={{fontSize:15,fontWeight:800,color:T.text}}>🖨️ Print / Export PDF</div>
          <button onClick={generate} style={{background:T.blue,color:"#fff",border:"none",borderRadius:10,padding:"8px 16px",fontSize:13,fontWeight:700,cursor:"pointer",fontFamily:"inherit"}}>
            Generate PDF
          </button>
        </div>

        <div style={{padding:"14px 16px 100px"}}>
          {/* Report summary */}
          <div style={{...cardS,marginBottom:14,background:T.blueLow,border:`1px solid ${T.blue}30`}}>
            <div style={{fontSize:13,fontWeight:800,color:T.text}}>{project.name}</div>
            <div style={{fontSize:11,color:T.muted}}>{report.date} · Report #{report.report_no} · {report.submitted_by}</div>
          </div>
          <div style={{...cardS,marginBottom:14}}>
            <div style={{fontSize:12,fontWeight:800,color:T.text,textTransform:"uppercase",letterSpacing:"1px",marginBottom:12}}>Report Sections</div>
            <div style={{display:"flex",justifyContent:"flex-end",gap:8,marginBottom:10}}>
              <button onClick={()=>setPrintSections(Object.fromEntries(Object.keys(printSections).map(k=>[k,true])))}
                style={{...ghostBtn,fontSize:11,padding:"4px 10px"}}>All On</button>
              <button onClick={()=>setPrintSections(Object.fromEntries(Object.keys(printSections).map(k=>[k,false])))}
                style={{...ghostBtn,fontSize:11,padding:"4px 10px"}}>All Off</button>
            </div>
            {SECTIONS.map(s=>(
              <div key={s.key} onClick={()=>toggleSection(s.key)}
                style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"10px 0",borderBottom:`1px solid ${T.border}`,cursor:"pointer"}}>
                <span style={{fontSize:13,color:printSections[s.key]?T.text:T.muted}}>{s.label}</span>
                <div style={{width:44,height:24,borderRadius:12,background:printSections[s.key]?T.green:T.border,position:"relative",transition:"background 0.2s",flexShrink:0}}>
                  <div style={{position:"absolute",top:2,left:printSections[s.key]?20:2,width:20,height:20,borderRadius:"50%",background:"#fff",transition:"left 0.2s",boxShadow:"0 1px 3px rgba(0,0,0,0.3)"}}/>
                </div>
              </div>
            ))}
          </div>
          <div style={{...cardS,marginBottom:14}}>
            <div style={{fontSize:12,fontWeight:800,color:T.text,textTransform:"uppercase",letterSpacing:"1px",marginBottom:4}}>
              📷 Photos ({selectedPhotos.length}/{reportPhotos.length} selected)
            </div>
            <div style={{fontSize:11,color:T.muted,marginBottom:12}}>Tap photos to include/exclude from PDF</div>

            {photosLoading&&<div style={{textAlign:"center",padding:"20px 0",color:T.muted}}>Loading photos...</div>}

            {!photosLoading&&reportPhotos.length===0&&(
              <div style={{textAlign:"center",padding:"20px 0",color:T.muted,fontSize:12}}>No photos on this job yet</div>
            )}

            {!photosLoading&&reportPhotos.length>0&&<>
              <div style={{display:"flex",gap:8,marginBottom:12,flexWrap:"wrap"}}>
                <button onClick={selAll} style={{...ghostBtn,fontSize:11,padding:"4px 10px",color:T.green,border:`1px solid ${T.green}40`}}>✓ Select All ({reportPhotos.length})</button>
                <button onClick={selDateOnly} style={{...ghostBtn,fontSize:11,padding:"4px 10px",color:T.blue,border:`1px solid ${T.blue}40`}}>📅 This Date ({reportPhotos.filter(p=>p.date===report.date).length})</button>
                <button onClick={selNone} style={{...ghostBtn,fontSize:11,padding:"4px 10px",color:T.red,border:`1px solid ${T.red}40`}}>✕ None</button>
              </div>

              <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:6,marginBottom:12}}>
                {reportPhotos.map(ph=>{
                  const sel=selectedPhotos.includes(ph.id);
                  return(
                    <div key={ph.id} onClick={()=>togglePhoto(ph.id)}
                      style={{position:"relative",borderRadius:8,overflow:"hidden",aspectRatio:"4/3",cursor:"pointer",border:`2px solid ${sel?T.green:T.border}`,transition:"border-color 0.15s"}}>
                      <img src={ph.src} alt={ph.caption} style={{width:"100%",height:"100%",objectFit:"cover",display:"block",opacity:sel?1:0.4,transition:"opacity 0.15s"}}/>
                      {sel&&<div style={{position:"absolute",top:4,right:4,background:T.green,borderRadius:"50%",width:20,height:20,display:"flex",alignItems:"center",justifyContent:"center",fontSize:11,color:"#000",fontWeight:800}}>✓</div>}
                      {ph.category&&<div style={{position:"absolute",bottom:0,left:0,right:0,background:"rgba(0,0,0,0.7)",padding:"3px 5px",fontSize:8,color:catColor[ph.category]||T.muted,fontWeight:700}}>{ph.category}{ph.date===report.date?" · Today":ph.date?" · "+ph.date:""}</div>}
                    </div>
                  );
                })}
              </div>
              {selectedPhotos.length>0&&<div>
                <div style={{fontSize:11,fontWeight:700,color:T.muted,textTransform:"uppercase",letterSpacing:"1px",marginBottom:8}}>Photo Layout in PDF</div>
                <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8}}>
                  {[["grid","2×2 Grid (compact)"],["full","Full Page (large)"]].map(([val,label])=>(
                    <button key={val} onClick={()=>setPhotoLayout(val)}
                      style={{...ghostBtn,padding:"10px",textAlign:"center",fontSize:12,fontWeight:700,
                        background:photoLayout===val?T.blueLow:T.surface,
                        color:photoLayout===val?T.blue:T.muted,
                        border:`1px solid ${photoLayout===val?T.blue:T.border}`}}>
                      {label}
                    </button>
                  ))}
                </div>
              </div>}
            </>}
          </div>
          <button onClick={generate}
            style={{...primBtn,borderRadius:14,background:T.blue,color:"#fff"}}>
            🖨️ Generate PDF ({Object.values(printSections).filter(Boolean).length} sections{selectedPhotos.length>0?`, ${selectedPhotos.length} photos`:""})
          </button>
        </div>
      </div>
    );
  }

  if(showInspectorShare){
    const link=`${window.location.origin}${window.location.pathname}?inspect=${report.id}`;
    return(
      <div style={{background:T.bg,minHeight:"100vh",fontFamily:"inherit"}}>
        <TopBar title="📤 Send Inspector Link" onBack={()=>setShowInspectorShare(false)}/>
        <div style={{padding:"20px 16px"}}>
          <div style={{...cardS,marginBottom:14,background:T.blueLow,border:`1px solid ${T.blue}40`}}>
            <div style={{fontSize:14,fontWeight:800,color:T.blue,marginBottom:4}}>How it works</div>
            <div style={{fontSize:12,color:T.sub,lineHeight:1.7}}>
              1. Copy the link below and text or email it to the inspector<br/>
              2. They open it on their phone — no login needed<br/>
              3. They enter their name and draw their signature<br/>
              4. Saves directly to this report automatically
            </div>
          </div>
          <div style={{...cardS,marginBottom:12,padding:"12px 14px"}}>
            <div style={{fontSize:11,fontWeight:700,color:T.muted,marginBottom:6}}>SIGNING LINK</div>
            <div style={{fontSize:12,color:T.sub,wordBreak:"break-all",lineHeight:1.6,background:T.surface,borderRadius:8,padding:"10px 12px"}}>{link}</div>
          </div>
          <button onClick={()=>{navigator.clipboard.writeText(link).then(()=>{setInspLinkCopied(true);setTimeout(()=>setInspLinkCopied(false),3000);}).catch(()=>{const el=document.createElement("textarea");el.value=link;document.body.appendChild(el);el.select();document.execCommand("copy");document.body.removeChild(el);setInspLinkCopied(true);setTimeout(()=>setInspLinkCopied(false),3000);});}}
            style={{...primBtn,borderRadius:14,marginBottom:10,background:inspLinkCopied?T.green:T.orange,transition:"background 0.3s"}}>
            {inspLinkCopied?"✅ Copied! Paste into a text or email":"📋 Copy Link to Clipboard"}
          </button>
          {inspLinkCopied&&<div style={{background:T.greenLow,border:`1px solid ${T.green}40`,borderRadius:10,padding:"10px 14px",fontSize:13,color:T.green,textAlign:"center",marginBottom:10}}>✓ Link copied — paste it into a text or email to the inspector</div>}
          <div style={{display:"flex",alignItems:"center",gap:10,marginBottom:10}}>
            <div style={{flex:1,height:1,background:T.border}}/><span style={{fontSize:11,color:T.muted}}>OR</span><div style={{flex:1,height:1,background:T.border}}/>
          </div>
          <button onClick={()=>{
            const subj=`Daily Report Sign-Off — ${project.name} — ${report.date}`;
            const body=`Please sign off on the daily report for ${report.date} on ${project.name}.%0D%0A%0D%0AOpen this link — no login required:%0D%0A%0D%0A${link}%0D%0A%0D%0AThank you`;
            window.location.href=`mailto:?subject=${encodeURIComponent(subj)}&body=${encodeURIComponent(body).replace(/%250D%250A/g,'%0D%0A')}`;
          }} style={{...ghostBtn,width:"100%",textAlign:"center",fontSize:14}}>
            📧 Open Email Draft
          </button>
        </div>
      </div>
    );
  }

  return(
    <div style={{background:T.bg,minHeight:"100vh",padding:16,fontFamily:"inherit"}}>
      {showSigPad&&<SignaturePad reportName={`${project.name} · ${fmtDate(report.date)}`} onSave={saveSignature} onCancel={()=>setShowSigPad(false)}/>}
      <Lightbox src={lb} onClose={()=>setLb(null)}/>
      <button onClick={onBack} style={{...ghostBtn,marginBottom:14}}>← Reports</button>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:4}}><div style={{fontSize:20,fontWeight:900,letterSpacing:"-0.5px"}}>{fmtDate(report.date)}</div><span style={pill(sc)}>{(report.status||"submitted").toUpperCase()}</span></div>
      {report.submitted_by&&<div style={{fontSize:12,color:T.muted,marginBottom:14}}>by {report.submitted_by}</div>}
      {report.pm_notes&&<div style={{...cardS,marginBottom:14,borderLeft:`3px solid ${T.red}`,background:T.redLow}}><div style={{fontSize:11,color:T.red,fontWeight:700,marginBottom:4}}>🚩 PM NOTE</div><div style={{fontSize:13,color:T.sub}}>{report.pm_notes}</div></div>}
      {report.description&&<div style={{...cardS,marginBottom:12,borderLeft:`3px solid ${T.blue}`}}><div style={{fontSize:11,color:T.muted,textTransform:"uppercase",letterSpacing:"1px",marginBottom:6}}>Work Done</div><div style={{fontSize:14,color:T.sub,lineHeight:1.6}}>{report.description}</div></div>}
      {(report.labor||[]).length>0&&<div style={{...cardS,marginBottom:12}}><div style={{fontSize:12,color:divColor,fontWeight:700,textTransform:"uppercase",letterSpacing:"1px",marginBottom:10}}>Labor{can(user,"view_dashboard")&&<span style={{color:T.green}}> · ${fmt(tot.labor)}</span>}</div>{report.labor.map((r,i)=>(<div key={i} style={{display:"flex",justifyContent:"space-between",padding:"8px 0",borderBottom:i<report.labor.length-1?`1px solid ${T.border}`:"none"}}><div><div style={{fontSize:14,fontWeight:600,color:T.text}}>{r.name||"—"}</div><div style={{fontSize:11,color:T.muted}}>{r.classification} · {r.regHrs||0}reg {r.otHrs||0}OT {r.travelHrs||0}tr</div></div>{can(user,"view_dashboard")&&<div style={{fontSize:14,fontWeight:800,color:T.green}}>${fmt(laborAmt(r))}</div>}</div>))}</div>}
      {(report.equipment||[]).length>0&&<div style={{...cardS,marginBottom:12}}><div style={{fontSize:12,color:divColor,fontWeight:700,textTransform:"uppercase",letterSpacing:"1px",marginBottom:10}}>Equipment{can(user,"view_dashboard")&&<span style={{color:T.green}}> · ${fmt(tot.equip)}</span>}</div>{report.equipment.map((r,i)=>(<div key={i} style={{display:"flex",justifyContent:"space-between",padding:"8px 0",borderBottom:i<report.equipment.length-1?`1px solid ${T.border}`:"none"}}><div style={{flex:1,paddingRight:10}}><div style={{fontSize:13,fontWeight:600,color:T.text}}>{r.description}</div><div style={{fontSize:11,color:T.muted}}>Qty {r.qty} x {r.usage} {r.unit}</div></div>{can(user,"view_dashboard")&&<div style={{fontSize:14,fontWeight:800,color:T.green}}>${fmt(equipAmt(r,project.division))}</div>}</div>))}</div>}
      {(report.rental_equipment||[]).length>0&&<div style={{...cardS,marginBottom:12,borderLeft:`3px solid ${T.purple}`}}><div style={{fontSize:12,color:T.purple,fontWeight:700,textTransform:"uppercase",letterSpacing:"1px",marginBottom:10}}>🔑 Rented Equipment{can(user,"view_dashboard")&&<span style={{color:T.green}}> · ${fmt((report.rental_equipment||[]).reduce((s,r)=>s+(parseFloat(r.qty)||0)*(parseFloat(r.rate)||0)*(parseFloat(r.usage)||1),0))}</span>}</div>{(report.rental_equipment||[]).map((r,i)=>{const amt=(parseFloat(r.qty)||0)*(parseFloat(r.rate)||0)*(parseFloat(r.usage)||1);return(<div key={i} style={{display:"flex",justifyContent:"space-between",padding:"8px 0",borderBottom:i<report.rental_equipment.length-1?`1px solid ${T.border}`:"none"}}><div style={{flex:1,paddingRight:10}}><div style={{fontSize:13,fontWeight:600,color:T.text}}>{r.description}</div><div style={{fontSize:11,color:T.muted}}>Qty {r.qty||0} × {r.usage||0} days/hrs @ ${r.rate||0}</div></div>{can(user,"view_dashboard")&&<div style={{fontSize:14,fontWeight:800,color:T.green}}>${fmt(amt)}</div>}</div>);})}</div>}
      {(report.materials||[]).length>0&&<div style={{...cardS,marginBottom:12}}><div style={{fontSize:12,color:divColor,fontWeight:700,textTransform:"uppercase",letterSpacing:"1px",marginBottom:10}}>Materials{can(user,"view_dashboard")&&<span style={{color:T.green}}> · ${fmt(tot.mats)}</span>}</div>{report.materials.map((r,i)=>(<div key={i} style={{padding:"8px 0",borderBottom:i<report.materials.length-1?`1px solid ${T.border}`:"none"}}><div style={{display:"flex",justifyContent:"space-between",marginBottom:r.receipts?.length>0?8:0}}><span style={{fontSize:13}}>{r.qty?`${r.qty}x `:""}{r.description}</span>{can(user,"view_dashboard")&&<span style={{fontSize:13,fontWeight:700,color:T.green}}>${fmt(parseFloat(r.amount)||0)}</span>}</div>{r.receipts?.length>0&&<div style={{display:"flex",gap:6,flexWrap:"wrap"}}>{r.receipts.map(rc=><img key={rc.id} src={rc.src} alt="" onClick={()=>setLb(rc.src)} style={{width:56,height:56,objectFit:"cover",borderRadius:8,cursor:"pointer"}}/>)}</div>}</div>))}</div>}
      {can(user,"view_dashboard")&&<div style={{...cardS,background:divColor+"12",border:`1px solid ${divColor}40`,marginBottom:12,display:"flex",justifyContent:"space-between",alignItems:"center"}}><span style={{fontSize:15,fontWeight:800}}>Grand Total</span><span style={{fontSize:26,fontWeight:900,color:divColor,letterSpacing:"-1px"}}>${fmt(tot.grand)}</span></div>}
      {can(user,"approve_report")&&report.status==="submitted"&&(<div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:10}}><button onClick={()=>onApprove&&onApprove(report.id)} style={{...primBtn,background:T.greenLow,color:T.green,border:`1px solid ${T.green}40`,borderRadius:12}}>✓ Approve</button><button onClick={()=>setFlagging(!flagging)} style={{...primBtn,background:T.redLow,color:T.red,border:`1px solid ${T.red}40`,borderRadius:12}}>🚩 Flag</button></div>)}
      {flagging&&<div style={{...cardS,marginBottom:10}}><label style={lbl}>Flag Note for Crew</label><textarea value={flagNote} onChange={e=>setFlagNote(e.target.value)} rows={3} placeholder="What needs to be corrected…" style={{...inp,resize:"vertical",marginBottom:10}}/><button onClick={()=>{onFlag&&onFlag(report.id,flagNote);setFlagging(false);}} style={{...primBtn,borderRadius:12}}>Send Flag</button></div>}
      {/* Signature section */}
      {report.inspector_signature?(
        <div style={{...cardS,marginBottom:12,borderLeft:`3px solid ${T.green}`}}>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:8}}>
            <div>
              <div style={{fontSize:12,fontWeight:700,color:T.green,textTransform:"uppercase",letterSpacing:"1px"}}>✅ Inspector Sign-Off</div>
              <div style={{fontSize:14,fontWeight:700,color:T.orange,marginTop:2}}>{report.inspector_name}</div>
              {report.inspector_signed_at&&<div style={{fontSize:11,color:T.muted,marginTop:2}}>{new Date(report.inspector_signed_at).toLocaleString()}</div>}
            </div>
          </div>
          <div style={{background:"#fff",borderRadius:10,padding:4,marginTop:4}}>
            <img src={report.inspector_signature} alt="Inspector signature" style={{width:"100%",borderRadius:8,display:"block"}}/>
          </div>
        </div>
      ):(
        <div style={{marginBottom:12}}>
          <div style={{fontSize:11,fontWeight:700,color:T.muted,textTransform:"uppercase",letterSpacing:"1px",marginBottom:8}}>🔏 Inspector Sign-Off</div>
          {report.hellosign_status==="pending"&&<div style={{background:T.blueLow,border:`1px solid ${T.blue}40`,borderRadius:10,padding:"10px 14px",marginBottom:8,fontSize:12,color:T.blue,fontWeight:600}}>
            📦 Box Sign request sent — waiting for inspector to sign
          </div>}
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:8}}>
            <button onClick={()=>setShowSigPad(true)}
              style={{...primBtn,background:T.greenLow,color:T.green,border:`1px solid ${T.green}40`,borderRadius:12,fontSize:12}}>
              ✍️ Sign Here
            </button>
            <button onClick={()=>setShowInspectorShare(true)}
              style={{...primBtn,background:T.blueLow,color:T.blue,border:`1px solid ${T.blue}40`,borderRadius:12,fontSize:12}}>
              🔗 Send Link
            </button>
            <button onClick={()=>setShowEsigModal(true)}
              style={{...primBtn,background:"#1e3a5f",color:"#60A5FA",border:"1px solid #2563EB40",borderRadius:12,fontSize:12}}>
              ✉️ eSign
            </button>
          </div>
          <div style={{fontSize:10,color:T.muted,textAlign:"center",marginTop:5}}>
            Sign Here = on this device · Send Link = inspector's phone · eSign = Box Sign email
          </div>
        </div>
      )}

      {/* eSignature Modal */}
      {showEsigModal&&<div style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.85)",zIndex:200,display:"flex",alignItems:"center",justifyContent:"center",padding:16,fontFamily:"inherit"}}>
        <div style={{background:T.card,borderRadius:16,padding:24,width:"100%",maxWidth:400,border:`1px solid ${T.blue}40`}}>
          <div style={{fontSize:16,fontWeight:900,color:T.blue,marginBottom:4}}>📦 Send via Box Sign</div>
          <div style={{fontSize:12,color:T.muted,marginBottom:16}}>Inspector receives a Box Sign email and can sign on any device. Report auto-updates when signed.</div>
          {esigError&&<div style={{background:T.redLow,border:`1px solid ${T.red}40`,borderRadius:8,padding:"8px 12px",marginBottom:12,fontSize:12,color:T.red}}>{esigError}</div>}
          <div style={{marginBottom:10}}><label style={lbl}>Inspector Name *</label><input value={esigName} onChange={e=>setEsigName(e.target.value)} placeholder="John Smith" style={inp} autoFocus/></div>
          <div style={{marginBottom:16}}><label style={lbl}>Inspector Email *</label><input type="email" value={esigEmail} onChange={e=>setEsigEmail(e.target.value)} placeholder="inspector@company.com" style={inp}/></div>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8}}>
            <button onClick={sendEsig} disabled={esigSending||!esigEmail.trim()||!esigName.trim()}
              style={{...primBtn,borderRadius:12,background:T.blue,opacity:esigSending||!esigEmail.trim()||!esigName.trim()?0.5:1}}>
              {esigSending?"Sending…":"Send Request"}
            </button>
            <button onClick={()=>{setShowEsigModal(false);setEsigError("");}}
              style={{...ghostBtn,textAlign:"center",borderRadius:12}}>Cancel</button>
          </div>
        </div>
      </div>}

      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:10}}>
        <button onClick={async()=>{setShowPrintModal(true);await loadPhotosForPrint();}} style={{...primBtn,background:"#1f3864",color:"#fff",borderRadius:14}}>🖨️ Print / Save PDF</button>
        <button onClick={exportXLSX} style={{...primBtn,background:divColor+"15",color:divColor,border:`1px solid ${divColor}40`,borderRadius:14}}>📥 Excel (.xlsx)</button>
      </div>
      {can(user,"approve_report")&&!editing&&(
        <button onClick={()=>{setEditData({date:report.date,report_no:report.report_no||"",description:report.description||""});setEditing(true);}} style={{...ghostBtn,width:"100%",textAlign:"center",marginBottom:10}}>
          ✏️ Edit Report
        </button>
      )}
      <button onClick={()=>window.confirm("Delete this report?")&&onDelete(report.id)} style={dangerBtn}>🗑 Delete Report</button>
    </div>
  );
}

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
    <div style={{background:T.greenLow,border:`1px solid ${T.green}40`,borderRadius:12,padding:"10px 14px",marginBottom:14,fontSize:12,color:T.green,lineHeight:1.5}}>
      <strong>⚡ Auto-filled from daily reports</strong> — hours are added automatically when a foreman submits a daily report. Manual entries below for anything not on a daily.
    </div>
    <button onClick={()=>setShowForm(!showForm)} style={{...primBtn,marginBottom:14,borderRadius:14}}>{showForm?"✕ Cancel":"⏱️ Log Time"}</button>
    {showForm&&<div style={{...cardS,marginBottom:14,borderLeft:`3px solid ${T.green}`}}>
      <div style={{marginBottom:10}}><label style={lbl}>Worker</label><select value={f.worker_name} onChange={e=>setF(x=>({...x,worker_name:e.target.value}))} style={inpSel}>{NAMES.map(n=><option key={n}>{n}</option>)}</select></div>
      <div style={{marginBottom:10}}><label style={lbl}>Date</label><input type="date" value={f.date} onChange={e=>setF(x=>({...x,date:e.target.value}))} style={inp}/></div>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:10}}><div><label style={lbl}>Clock In</label><input type="time" value={f.clock_in} onChange={e=>setF(x=>({...x,clock_in:e.target.value}))} style={inp}/></div><div><label style={lbl}>Clock Out</label><input type="time" value={f.clock_out} onChange={e=>setF(x=>({...x,clock_out:e.target.value}))} style={inp}/></div></div>
      {f.clock_in&&f.clock_out&&(()=>{const h=calcHours(f.clock_in,f.clock_out);const ot=Math.max(0,h-8);return h>0&&(<div style={{background:T.greenLow,borderRadius:10,padding:"10px 12px",marginBottom:10,display:"flex",gap:16}}><div><div style={{fontSize:18,fontWeight:900,color:T.green}}>{h.toFixed(2)}h</div><div style={{fontSize:10,color:T.muted,textTransform:"uppercase"}}>Total</div></div><div><div style={{fontSize:18,fontWeight:900,color:ot>0?T.yellow:T.muted}}>{Math.min(h,8).toFixed(2)}h</div><div style={{fontSize:10,color:T.muted,textTransform:"uppercase"}}>Regular</div></div>{ot>0&&<div><div style={{fontSize:18,fontWeight:900,color:T.yellow}}>{ot.toFixed(2)}h</div><div style={{fontSize:10,color:T.muted,textTransform:"uppercase"}}>OT</div></div>}</div>);})()}
      <div style={{marginBottom:10}}><label style={lbl}>Notes</label><input type="text" placeholder="Optional…" value={f.notes} onChange={e=>setF(x=>({...x,notes:e.target.value}))} style={inp}/></div>
      <button onClick={save} style={{...primBtn,background:T.green,color:"#0D0D0F",borderRadius:12}}>{saving?"Saving…":"Save Time Card"}</button>
    </div>}
    {loading&&<Spinner/>}
    {!loading&&<>{workerRows.length>0&&<div style={{...cardS,marginBottom:14}}><div style={{fontSize:12,fontWeight:700,color:T.green,textTransform:"uppercase",letterSpacing:"1px",marginBottom:10}}>This Week</div>{workerRows.map(w=>(<div key={w.name} style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"8px 0",borderBottom:`1px solid ${T.border}`}}><span style={{fontSize:14,fontWeight:600,color:T.text}}>{w.name}</span><div style={{display:"flex",gap:12,alignItems:"center"}}><span style={{fontSize:12,color:T.muted}}>{w.reg.toFixed(1)}reg</span>{w.ot>0&&<span style={{fontSize:12,color:T.yellow}}>{w.ot.toFixed(1)}OT</span>}<span style={{fontSize:15,fontWeight:800,color:T.green}}>{w.total.toFixed(1)}h</span></div></div>))}</div>}
    {todayCards.length>0&&<div style={{fontSize:12,fontWeight:700,color:T.muted,textTransform:"uppercase",letterSpacing:"1px",marginBottom:10}}>Today</div>}
    {todayCards.map(c=><div key={c.id} style={{...cardS,marginBottom:8,display:"flex",justifyContent:"space-between",alignItems:"center"}}><div><div style={{fontSize:14,fontWeight:700,color:T.text}}>{c.worker_name}</div><div style={{fontSize:11,color:T.muted,marginTop:3}}>{fmtShort(c.date)}{c.clock_in?" · "+c.clock_in:""}{c.clock_out?" → "+c.clock_out:""}</div>{c.notes&&<div style={{fontSize:11,color:T.sub,marginTop:2}}>{c.notes}</div>}</div><div style={{display:"flex",alignItems:"center",gap:10}}><div style={{textAlign:"right"}}><div style={{fontSize:16,fontWeight:800,color:T.green}}>{(c.total_hours||0).toFixed(1)}h</div>{(c.ot_hours||0)>0&&<div style={{fontSize:10,color:T.yellow}}>{c.ot_hours.toFixed(1)} OT</div>}</div><button onClick={()=>remove(c.id)} style={{background:"none",border:"none",color:T.muted,cursor:"pointer",fontSize:16,padding:0}}>🗑</button></div></div>)}
    {recentCards.length>0&&<div style={{fontSize:12,fontWeight:700,color:T.muted,textTransform:"uppercase",letterSpacing:"1px",margin:"14px 0 10px"}}>Recent</div>}
    {recentCards.map(c=><div key={c.id} style={{...cardS,marginBottom:8,display:"flex",justifyContent:"space-between",alignItems:"center"}}><div><div style={{fontSize:14,fontWeight:700,color:T.text}}>{c.worker_name}</div><div style={{fontSize:11,color:T.muted,marginTop:3}}>{fmtShort(c.date)}{c.clock_in?" · "+c.clock_in:""}{c.clock_out?" → "+c.clock_out:""}</div></div><div style={{display:"flex",alignItems:"center",gap:10}}><div style={{fontSize:16,fontWeight:800,color:T.green}}>{(c.total_hours||0).toFixed(1)}h</div><button onClick={()=>remove(c.id)} style={{background:"none",border:"none",color:T.muted,cursor:"pointer",fontSize:16,padding:0}}>🗑</button></div></div>)}
    {cards.length===0&&!showForm&&<div style={{textAlign:"center",padding:"40px 0",color:T.muted}}><div style={{fontSize:32,marginBottom:8}}>⏱️</div><div>No time cards yet.</div></div>}</>}
  </div>);
}

function CrewEquipTab({projectId,user,onErr}){
  const [equip,setEquip]=useState([]);const [loading,setLoading]=useState(true);const [showForm,setShowForm]=useState(false);const [saving,setSaving]=useState(false);
  const [f,setF]=useState({equipment_name:"",quantity:1,operator_name:"",hours_used:"",notes:"",date:today()});
  async function load(){setLoading(true);try{setEquip(await API.equipment.forProject(projectId)||[]);}catch(e){onErr(e.message);}setLoading(false);}
  useEffect(()=>{load();},[projectId]);
  async function save(){if(!f.equipment_name)return;setSaving(true);try{await API.equipment.create({...f,project_id:projectId});await load();setShowForm(false);setF({equipment_name:"",quantity:1,operator_name:"",hours_used:"",notes:"",date:today()});}catch(e){onErr(e.message);}setSaving(false);}
  async function remove(id){try{await API.equipment.remove(id);await load();}catch(e){onErr(e.message);}}
  const todayEquip=equip.filter(e=>e.date===today());const prevEquip=equip.filter(e=>e.date!==today()).slice(0,20);
  return(<div>
    <div style={{background:T.greenLow,border:`1px solid ${T.green}40`,borderRadius:12,padding:"10px 14px",marginBottom:14,fontSize:12,color:T.green,lineHeight:1.5}}>
      <strong>⚡ Auto-filled from daily reports</strong> — hours are added automatically when a foreman submits a daily report. Manual entries below for anything not on a daily.
    </div>
    <button onClick={()=>setShowForm(!showForm)} style={{...primBtn,marginBottom:14,borderRadius:14}}>{showForm?"✕ Cancel":"🚜 Log Equipment On Site"}</button>
    {showForm&&<div style={{...cardS,marginBottom:14,borderLeft:`3px solid ${T.yellow}`}}>
      <div style={{marginBottom:10}}><label style={lbl}>Equipment</label><select value={f.equipment_name} onChange={e=>setF(x=>({...x,equipment_name:e.target.value}))} style={inpSel}><option value="">— Select —</option>{EQUIP_LIST.filter(e=>!e.section).map(e=><option key={e.name} value={e.name}>{e.name}</option>)}</select></div>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:10}}><div><label style={lbl}>Quantity</label><input type="number" min="1" value={f.quantity} onChange={e=>setF(x=>({...x,quantity:e.target.value}))} style={inp}/></div><div><label style={lbl}>Hours Used</label><input type="number" min="0" step="0.5" placeholder="0" value={f.hours_used} onChange={e=>setF(x=>({...x,hours_used:e.target.value}))} style={inp}/></div></div>
      <div style={{marginBottom:10}}><label style={lbl}>Operator</label><select value={f.operator_name} onChange={e=>setF(x=>({...x,operator_name:e.target.value}))} style={inpSel}><option value="">— Optional —</option>{NAMES.map(n=><option key={n}>{n}</option>)}</select></div>
      <div style={{marginBottom:10}}><label style={lbl}>Date</label><input type="date" value={f.date} onChange={e=>setF(x=>({...x,date:e.target.value}))} style={inp}/></div>
      <div style={{marginBottom:10}}><label style={lbl}>Notes</label><input type="text" placeholder="Condition, issues…" value={f.notes} onChange={e=>setF(x=>({...x,notes:e.target.value}))} style={inp}/></div>
      <button onClick={save} style={{...primBtn,background:T.yellow,color:"#0D0D0F",borderRadius:12}}>{saving?"Saving…":"Save Entry"}</button>
    </div>}
    {loading&&<Spinner/>}
    {!loading&&<>{todayEquip.length>0&&<div style={{fontSize:12,fontWeight:700,color:T.muted,textTransform:"uppercase",letterSpacing:"1px",marginBottom:10}}>On Site Today</div>}{todayEquip.map(e=><div key={e.id} style={{...cardS,marginBottom:8,borderLeft:`3px solid ${T.yellow}`,display:"flex",justifyContent:"space-between",alignItems:"center"}}><div><div style={{fontSize:14,fontWeight:700,color:T.text}}>{e.equipment_name}</div><div style={{fontSize:11,color:T.muted,marginTop:3}}>{fmtShort(e.date)} · Qty {e.quantity||1}{e.operator_name?" · "+e.operator_name:""}{e.hours_used?" · "+e.hours_used+"h":""}</div></div><button onClick={()=>remove(e.id)} style={{background:"none",border:"none",color:T.muted,cursor:"pointer",fontSize:16,padding:0,marginLeft:12}}>🗑</button></div>)}
    {prevEquip.length>0&&<div style={{fontSize:12,fontWeight:700,color:T.muted,textTransform:"uppercase",letterSpacing:"1px",margin:"14px 0 10px"}}>Previous</div>}{prevEquip.map(e=><div key={e.id} style={{...cardS,marginBottom:8,borderLeft:`3px solid ${T.yellow}40`,display:"flex",justifyContent:"space-between",alignItems:"center"}}><div><div style={{fontSize:13,fontWeight:700}}>{e.equipment_name}</div><div style={{fontSize:11,color:T.muted}}>{fmtShort(e.date)} · Qty {e.quantity||1}{e.operator_name?" · "+e.operator_name:""}</div></div><button onClick={()=>remove(e.id)} style={{background:"none",border:"none",color:T.muted,cursor:"pointer",fontSize:16,padding:0,marginLeft:12}}>🗑</button></div>)}
    {equip.length===0&&!showForm&&<div style={{textAlign:"center",padding:"40px 0",color:T.muted}}><div style={{fontSize:32,marginBottom:8}}>🚜</div><div>No equipment logged.</div></div>}</>}
  </div>);
}

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
    <div style={{background:T.greenLow,border:`1px solid ${T.green}40`,borderRadius:12,padding:"10px 14px",marginBottom:14,fontSize:12,color:T.green,lineHeight:1.5}}>
      <strong>⚡ Auto-filled from daily reports</strong> — hours are added automatically when a foreman submits a daily report. Manual entries below for anything not on a daily.
    </div>
    <button onClick={()=>setShowForm(!showForm)} style={{...primBtn,marginBottom:14,borderRadius:14}}>{showForm?"✕ Cancel":"🏢 Log Subcontractor"}</button>
    {showForm&&<div style={{...cardS,marginBottom:14,borderLeft:`3px solid ${T.purple}`}}>
      <div style={{marginBottom:10}}><label style={lbl}>Company Name *</label><input type="text" placeholder="Sub company name" value={f.company_name} onChange={e=>set("company_name",e.target.value)} style={inp}/></div>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:10}}><div><label style={lbl}>Trade</label><select value={f.trade} onChange={e=>set("trade",e.target.value)} style={inpSel}><option value="">— Select —</option>{trades.map(t=><option key={t}>{t}</option>)}</select></div><div><label style={lbl}>Date</label><input type="date" value={f.date} onChange={e=>set("date",e.target.value)} style={inp}/></div><div><label style={lbl}>Contact</label><input type="text" placeholder="Foreman" value={f.contact_name} onChange={e=>set("contact_name",e.target.value)} style={inp}/></div><div><label style={lbl}>Phone</label><input type="tel" placeholder="555-555-5555" value={f.contact_phone} onChange={e=>set("contact_phone",e.target.value)} style={inp}/></div><div><label style={lbl}>Workers</label><input type="number" min="0" value={f.workers_count} onChange={e=>set("workers_count",e.target.value)} style={inp}/></div><div><label style={lbl}>Hours</label><input type="number" min="0" step="0.5" placeholder="0" value={f.hours_worked} onChange={e=>set("hours_worked",e.target.value)} style={inp}/></div></div>
      <div style={{marginBottom:10}}><label style={lbl}>Work Description</label><textarea placeholder="What work was performed?" value={f.work_description} onChange={e=>set("work_description",e.target.value)} rows={3} style={{...inp,resize:"vertical"}}/></div>
      <button onClick={save} style={{...primBtn,background:T.purple,color:"#fff",borderRadius:12}}>{saving?"Saving…":"Save Sub Entry"}</button>
    </div>}
    {loading&&<Spinner/>}
    {!loading&&subs.map(s=>(<div key={s.id} style={{...cardS,marginBottom:10,borderLeft:`3px solid ${T.purple}`}}><div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start"}}><div style={{flex:1}}><div style={{fontSize:15,fontWeight:800}}>{s.company_name}</div><div style={{display:"flex",gap:6,marginTop:5,flexWrap:"wrap"}}>{s.trade&&<span style={pill(T.purple)}>{s.trade}</span>}<span style={pill(T.muted)}>{fmtShort(s.date)}</span>{s.workers_count>0&&<span style={pill(T.blue)}>👷 {s.workers_count}</span>}{s.hours_worked>0&&<span style={pill(T.green)}>{s.hours_worked}h</span>}</div>{s.contact_name&&<div style={{fontSize:12,color:T.sub,marginTop:6}}>📞 {s.contact_name}{s.contact_phone?" · "+s.contact_phone:""}</div>}{s.work_description&&<div style={{fontSize:12,color:T.sub,marginTop:4,lineHeight:1.5}}>{s.work_description}</div>}</div><button onClick={()=>remove(s.id)} style={{background:"none",border:"none",color:T.muted,cursor:"pointer",fontSize:16,padding:"0 0 0 10px"}}>🗑</button></div></div>))}
    {!loading&&subs.length===0&&!showForm&&<div style={{textAlign:"center",padding:"40px 0",color:T.muted}}><div style={{fontSize:32,marginBottom:8}}>🏢</div><div>No subcontractors logged.</div></div>}
  </div>);
}

function SafetyTab({projectId,safety,user,onRefresh,onErr}){
  const [showForm,setShowForm]=useState(false);const [saving,setSaving]=useState(false);const [type,setType]=useState("toolbox");
  const [f,setF]=useState({date:today(),topic:"",notes:"",severity:"low"});
  const TC={toolbox:T.blue,observation:T.yellow,incident:T.red,nearmiss:T.orange,jsa:T.purple};
  const TL={toolbox:"🛠 Toolbox Talk",observation:"👁 Observation",incident:"🚨 Incident",nearmiss:"⚠️ Near Miss",jsa:"📋 JSA"};
  async function save(){if(!f.topic.trim())return;setSaving(true);try{await API.safety.create({...f,type,project_id:projectId,created_by:user.name});await onRefresh();setShowForm(false);setF({date:today(),topic:"",notes:"",severity:"low"});}catch(e){onErr(e.message);}setSaving(false);}
  async function del(id){try{await API.safety.remove(id);await onRefresh();}catch(e){onErr(e.message);}}
  return(<div>
    <div style={{background:T.greenLow,border:`1px solid ${T.green}40`,borderRadius:12,padding:"10px 14px",marginBottom:14,fontSize:12,color:T.green,lineHeight:1.5}}>
      <strong>⚡ Auto-filled from daily reports</strong> — hours are added automatically when a foreman submits a daily report. Manual entries below for anything not on a daily.
    </div>
    <button onClick={()=>setShowForm(!showForm)} style={{...primBtn,marginBottom:14,borderRadius:14}}>{showForm?"✕ Cancel":"⛑️ Log Safety Entry"}</button>
    {showForm&&<div style={{...cardS,marginBottom:14,borderLeft:`3px solid ${T.yellow}`}}>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8,marginBottom:12}}>{Object.entries(TL).map(([k,v])=>(<button key={k} onClick={()=>setType(k)} style={{padding:"10px",borderRadius:10,border:`2px solid ${type===k?TC[k]:T.border}`,background:type===k?TC[k]+"20":T.surface,color:type===k?TC[k]:T.sub,fontSize:12,fontWeight:700,cursor:"pointer",fontFamily:"inherit"}}>{v}</button>))}</div>
      <div style={{marginBottom:10}}><label style={lbl}>Date</label><input type="date" value={f.date} onChange={e=>setF(x=>({...x,date:e.target.value}))} style={inp}/></div>
      <div style={{marginBottom:10}}><label style={lbl}>{type==="toolbox"?"Topic":type==="jsa"?"Job / Task Name":"Description"}</label><input type="text" placeholder="Describe…" value={f.topic} onChange={e=>setF(x=>({...x,topic:e.target.value}))} style={inp}/></div>
      <div style={{marginBottom:10}}><label style={lbl}>Notes / Corrective Action</label><textarea rows={3} placeholder="Additional details…" value={f.notes} onChange={e=>setF(x=>({...x,notes:e.target.value}))} style={{...inp,resize:"vertical"}}/></div>
      {(type==="incident"||type==="nearmiss")&&<div style={{marginBottom:10}}><label style={lbl}>Severity</label><select value={f.severity} onChange={e=>setF(x=>({...x,severity:e.target.value}))} style={inpSel}><option value="low">Low – First Aid</option><option value="medium">Medium – Recordable</option><option value="high">High – Lost Time</option></select></div>}
      <button onClick={save} style={{...primBtn,background:T.yellow,color:"#0D0D0F",borderRadius:12}}>{saving?"Saving…":"Save Entry"}</button>
    </div>}
    {safety.length===0&&!showForm&&<div style={{textAlign:"center",padding:"40px 0",color:T.muted}}><div style={{fontSize:32,marginBottom:8}}>⛑️</div><div>No safety entries yet.</div></div>}
    {[...safety].map(s=>(<div key={s.id} style={{...cardS,marginBottom:9,borderLeft:`3px solid ${TC[s.type]||T.border}`}}><div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start"}}><div style={{flex:1}}><span style={{...pill(TC[s.type]||T.muted),marginBottom:6,display:"inline-flex"}}>{TL[s.type]||s.type}</span><div style={{fontSize:14,fontWeight:700,marginTop:4}}>{s.topic}</div>{s.notes&&<div style={{fontSize:12,color:T.sub,marginTop:4,lineHeight:1.5}}>{s.notes}</div>}<div style={{fontSize:11,color:T.muted,marginTop:6}}>{fmtDate(s.date)} · {s.created_by}</div></div><div style={{display:"flex",flexDirection:"column",alignItems:"flex-end",gap:6}}>{(s.type==="incident"||s.type==="nearmiss")&&<span style={pill(s.severity==="high"?T.red:s.severity==="medium"?T.yellow:T.green)}>{(s.severity||"low").toUpperCase()}</span>}<button onClick={()=>del(s.id)} style={{background:"none",border:"none",color:T.muted,cursor:"pointer",fontSize:16,padding:0}}>🗑</button></div></div></div>))}
  </div>);
}

function DocsTab({projectId,user,onErr}){
  const canAdmin=user.role==="admin"||user.role==="pm";

  const [folders,setFolders]=useState([]);
  const [currentFolder,setCurrentFolder]=useState(null); // null = root view
  const [showNewFolder,setShowNewFolder]=useState(false);
  const [folderName,setFolderName]=useState("");
  const [folderColor,setFolderColor]=useState("#60A5FA");
  const [folderDesc,setFolderDesc]=useState("");
  const [editingFolder,setEditingFolder]=useState(null);

  const [docs,setDocs]=useState([]);
  const [loading,setLoading]=useState(true);
  const [uploading,setUploading]=useState(false);
  const [dragging,setDragging]=useState(false);
  const [editPerms,setEditPerms]=useState(null);
  const [dragDoc,setDragDoc]=useState(null);   // doc being dragged
  const [dragOverFolder,setDragOverFolder]=useState(null); // folder being hovered
  const fileRef=useRef(null);

  const FOLDER_COLORS=["#60A5FA","#34D399","#F97316","#FC8181","#FBBF24","#A78BFA","#2DD4BF","#F472B6"];
  const docIcons={Drawing:"📐",Specification:"📄",Manual:"📗",Permit:"🗂️",Contract:"📝","As-Built":"🗺️",ITP:"✅",Procedure:"📋","Safety Plan":"⛑️",Other:"📁","Fillable Form":"📝"};
  const mimeIcons={"application/pdf":"📄","image/":"🖼️","application/vnd.openxmlformats-officedocument.spreadsheetml":"📊","application/vnd.openxmlformats-officedocument.wordprocessingml":"📝","application/vnd.openxmlformats-officedocument.presentationml":"📊","video/":"🎬","text/":"📝"};
  function getMimeIcon(mime=""){for(const[k,v] of Object.entries(mimeIcons)){if(mime.startsWith(k))return v;}return"📁";}
  function fmtSize(bytes){if(!bytes)return"";if(bytes<1024)return bytes+"B";if(bytes<1048576)return(bytes/1024).toFixed(1)+"KB";return(bytes/1048576).toFixed(1)+"MB";}

  async function load(){
    setLoading(true);
    try{
      const [f,d]=await Promise.all([API.docFolders.forProject(projectId),API.docs.forProject(projectId)]);
      setFolders(Array.isArray(f)?f:[]);
      setDocs(Array.isArray(d)?d:[]);
    }catch(e){onErr(e.message);}
    setLoading(false);
  }
  useEffect(()=>{load();},[projectId]);

  async function createFolder(){
    if(!folderName.trim())return;
    try{
      await API.docFolders.create({project_id:projectId,name:folderName.trim(),description:folderDesc.trim(),color:folderColor,created_by:user.name});
      setFolderName("");setFolderDesc("");setFolderColor("#60A5FA");
      setShowNewFolder(false);await load();
    }catch(e){onErr(e.message);}
  }
  async function renameFolder(folder){
    const name=window.prompt("Rename folder:",folder.name);
    if(!name||!name.trim()||name===folder.name)return;
    try{await API.docFolders.update(folder.id,{name:name.trim()});await load();}
    catch(e){onErr(e.message);}
  }
  async function deleteFolder(folder){
    const docsInFolder=docs.filter(d=>d.folder_id===folder.id);
    const msg=docsInFolder.length>0
      ?`"${folder.name}" contains ${docsInFolder.length} document${docsInFolder.length!==1?"s":""}. Documents will be moved to the root. Delete folder?`
      :`Delete folder "${folder.name}"?`;
    if(!window.confirm(msg))return;
    try{
      await Promise.all(docsInFolder.map(d=>API.docs.update(d.id,{folder_id:null})));
      await API.docFolders.remove(folder.id);
      if(currentFolder?.id===folder.id)setCurrentFolder(null);
      await load();
    }catch(e){onErr(e.message);}
  }
  async function moveDocToFolder(doc,folderId){
    try{
      await API.docs.update(doc.id,{folder_id:folderId||null});
      await load();
    }catch(e){
      if(e.message?.includes("folder_id")||e.message?.includes("PGRST204")){
        onErr("Run AIME_docs_fix.sql in Supabase SQL Editor to enable folders, then try again.");
      }else{onErr(e.message);}
    }
  }

  function toB64(file){return new Promise((res,rej)=>{const r=new FileReader();r.onload=()=>res(r.result);r.onerror=rej;r.readAsDataURL(file);});}
  async function uploadFiles(files){
    setUploading(true);
    let uploaded=0;
    for(const file of files){
      try{
        const data=await toB64(file);
        const isFillable=file.name.toLowerCase().endsWith('.pdf');
        const payload={
          project_id:projectId,
          name:file.name.replace(/\.[^.]+$/,""),
          doc_type:isFillable?"Fillable Form":"Other",
          file:data,
          file_name:file.name,
          file_size:file.size,
          file_type:file.type,
          folder_id:currentFolder?.id||null,
          visible_to:["admin","pm","estimator","foreman","crew"],
          can_download:["admin","pm","estimator","foreman","crew"],
          uploaded_by:user.name,
          is_fillable:isFillable,
        };
        try{
          await API.docs.create(payload);
        }catch(colErr){
          if(colErr.message?.includes("PGRST204")||colErr.message?.includes("schema cache")){
            const{file_type,file_size,is_fillable,uploaded_by,folder_id,...safe}=payload;
            await API.docs.create(safe);
          }else{throw colErr;}
        }
        uploaded++;
      }catch(e){onErr("Upload failed: "+e.message+". Run AIME_docs_fix.sql in Supabase then try again.");}
    }
    setUploading(false);
    if(uploaded>0)await load();
  }
  async function removeDoc(id){
    if(!window.confirm("Delete this document?"))return;
    try{await API.docs.remove(id);await load();}catch(e){onErr(e.message);}
  }
  function downloadDoc(doc){
    if(!doc.file)return;
    const a=document.createElement("a");
    a.href=doc.file;
    a.download=doc.file_name||doc.name||"document";
    document.body.appendChild(a);a.click();document.body.removeChild(a);
  }
  function canDownload(doc){
    return(doc.can_download||[]).includes(user.role)||user.role==="admin";
  }

  const currentDocs=currentFolder
    ?docs.filter(d=>d.folder_id===currentFolder.id)
    :docs.filter(d=>!d.folder_id);

  const rootDocCount=docs.filter(d=>!d.folder_id).length;

  return(
    <div onDragOver={e=>{e.preventDefault();setDragging(true);}} onDragLeave={()=>setDragging(false)}
      onDrop={e=>{e.preventDefault();setDragging(false);uploadFiles(Array.from(e.dataTransfer.files));}}
      style={{minHeight:200}}>

      {/* Breadcrumb nav */}
      <div style={{display:"flex",alignItems:"center",gap:6,marginBottom:12,padding:"8px 0"}}>
        <button onClick={()=>setCurrentFolder(null)}
          style={{background:"none",border:"none",color:currentFolder?T.orange:T.text,fontSize:13,fontWeight:700,cursor:"pointer",fontFamily:"inherit",padding:0}}>
          📁 Documents
        </button>
        {currentFolder&&<>
          <span style={{color:T.muted,fontSize:13}}>›</span>
          <div style={{display:"flex",alignItems:"center",gap:6}}>
            <div style={{width:8,height:8,borderRadius:"50%",background:currentFolder.color||T.orange}}/>
            <span style={{fontSize:13,fontWeight:700,color:T.text}}>{currentFolder.name}</span>
          </div>
        </>}
      </div>
      {dragging&&<div style={{background:T.orangeLow,border:`2px dashed ${T.orange}`,borderRadius:12,padding:"20px",textAlign:"center",fontSize:14,color:T.orange,fontWeight:700,marginBottom:12}}>
        Drop files to upload{currentFolder?` into "${currentFolder.name}"`:""}</div>}

      {/* Action row */}
      <div style={{display:"flex",gap:8,marginBottom:16,flexWrap:"wrap"}}>
        <button onClick={()=>fileRef.current?.click()}
          style={{...primBtn,flex:1,minWidth:120,borderRadius:12,fontSize:13}}>
          {uploading?"Uploading…":"📎 Upload Files"}
        </button>
        {canAdmin&&!currentFolder&&<button onClick={()=>setShowNewFolder(true)}
          style={{...ghostBtn,flex:1,minWidth:120,textAlign:"center",fontSize:13,borderRadius:12,color:T.orange,border:`1px solid ${T.orange}40`}}>
          📁 New Folder
        </button>}
        {currentFolder&&<button onClick={()=>setCurrentFolder(null)}
          style={{...ghostBtn,flex:1,textAlign:"center",fontSize:12,borderRadius:12}}>
          ← Back to Root
        </button>}
        <input ref={fileRef} type="file" multiple accept="*/*" style={{display:"none"}}
          onChange={e=>{uploadFiles(Array.from(e.target.files));e.target.value="";}}/>
      </div>
      {showNewFolder&&<div style={{...cardS,marginBottom:14,border:`1px solid ${T.orange}40`}}>
        <div style={{fontSize:13,fontWeight:800,color:T.text,marginBottom:12}}>📁 Create New Folder</div>
        <div style={{marginBottom:10}}><label style={lbl}>Folder Name *</label>
          <input value={folderName} onChange={e=>setFolderName(e.target.value)} placeholder="e.g. As-Builts, Safety Plans, Permits" style={inp} autoFocus/></div>
        <div style={{marginBottom:10}}><label style={lbl}>Description (optional)</label>
          <input value={folderDesc} onChange={e=>setFolderDesc(e.target.value)} placeholder="What goes in this folder?" style={inp}/></div>
        <div style={{marginBottom:14}}>
          <label style={lbl}>Color</label>
          <div style={{display:"flex",gap:8,flexWrap:"wrap"}}>
            {FOLDER_COLORS.map(c=><div key={c} onClick={()=>setFolderColor(c)}
              style={{width:28,height:28,borderRadius:8,background:c,cursor:"pointer",border:folderColor===c?"3px solid #fff":"2px solid transparent",boxShadow:folderColor===c?"0 0 0 2px "+c:"none"}}/>)}
          </div>
        </div>
        <div style={{display:"flex",gap:8}}>
          <button onClick={createFolder} disabled={!folderName.trim()}
            style={{...primBtn,flex:2,borderRadius:12,opacity:folderName.trim()?1:0.5}}>Create Folder</button>
          <button onClick={()=>{setShowNewFolder(false);setFolderName("");}}
            style={{...ghostBtn,flex:1,textAlign:"center"}}>Cancel</button>
        </div>
      </div>}

      {loading&&<Spinner/>}

      {/* Root view — show folders + root docs */}
      {!loading&&!currentFolder&&<>
        {folders.length>0&&<div style={{marginBottom:16}}>
          <div style={{fontSize:11,fontWeight:700,color:T.muted,textTransform:"uppercase",letterSpacing:"1px",marginBottom:8}}>Folders ({folders.length})</div>
          {dragDoc&&<div style={{background:T.orangeLow,border:`1px solid ${T.orange}40`,borderRadius:10,padding:"8px 12px",marginBottom:8,fontSize:12,color:T.orange,fontWeight:700,textAlign:"center"}}>
            📂 Drop onto a folder to move the file
          </div>}
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8}}>
            {folders.map(folder=>{
              const count=docs.filter(d=>d.folder_id===folder.id).length;
              return(
                <div key={folder.id}
                  onClick={()=>!dragDoc&&setCurrentFolder(folder)}
                  onDragOver={e=>{e.preventDefault();setDragOverFolder(folder.id);}}
                  onDragLeave={()=>setDragOverFolder(null)}
                  onDrop={e=>{e.preventDefault();setDragOverFolder(null);if(dragDoc){moveDocToFolder(dragDoc,folder.id);setDragDoc(null);}}}
                  style={{...cardS,cursor:dragDoc?"copy":"pointer",padding:"12px 14px",
                    border:`2px solid ${dragOverFolder===folder.id?folder.color||T.orange:(folder.color||T.orange)+"30"}`,
                    background:dragOverFolder===folder.id?`${folder.color||T.orange}18`:T.card,
                    transform:dragOverFolder===folder.id?"scale(1.02)":"scale(1)",
                    transition:"all 0.15s",position:"relative"}}>
                  <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:4}}>
                    <div style={{fontSize:20}}>📁</div>
                    <div style={{flex:1,minWidth:0}}>
                      <div style={{fontSize:13,fontWeight:800,color:T.text,lineHeight:1.2,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{folder.name}</div>
                      <div style={{fontSize:10,color:T.muted,marginTop:1}}>{count} item{count!==1?"s":""}</div>
                    </div>
                    <div style={{width:8,height:8,borderRadius:"50%",background:folder.color||T.orange,flexShrink:0}}/>
                  </div>
                  {folder.description&&<div style={{fontSize:10,color:T.muted,lineHeight:1.4}}>{folder.description}</div>}
                  {canAdmin&&<div style={{display:"flex",gap:4,marginTop:8,paddingTop:8,borderTop:`1px solid ${T.border}`}}
                    onClick={e=>e.stopPropagation()}>
                    <button onClick={()=>renameFolder(folder)} style={{...ghostBtn,flex:1,textAlign:"center",fontSize:11,padding:"3px 6px"}}>✏️</button>
                    <button onClick={()=>deleteFolder(folder)} style={{...ghostBtn,flex:1,textAlign:"center",fontSize:11,padding:"3px 6px",color:T.red,border:`1px solid ${T.red}30`}}>🗑</button>
                  </div>}
                </div>
              );
            })}
          </div>
        </div>}

        {/* Root-level docs */}
        {rootDocCount>0&&<div style={{fontSize:11,fontWeight:700,color:T.muted,textTransform:"uppercase",letterSpacing:"1px",marginBottom:8}}>
          Root Files ({rootDocCount})
        </div>}
        {currentDocs.map(doc=>(<DocRow key={doc.id} doc={doc} folders={folders} user={user} canAdmin={canAdmin}
          onDownload={()=>downloadDoc(doc)} canDownload={canDownload(doc)}
          onMove={fid=>moveDocToFolder(doc,fid)} onDelete={()=>removeDoc(doc.id)}
          getMimeIcon={getMimeIcon} fmtSize={fmtSize} docIcons={docIcons}
          onDragStart={()=>setDragDoc(doc.id)} onDragEnd={()=>{setDragDoc(null);setDragOverFolder(null);}}
          isDragging={dragDoc===doc.id}/>))}
        {folders.length===0&&rootDocCount===0&&<div style={{textAlign:"center",padding:"40px 16px",color:T.muted}}>
          <div style={{fontSize:44,marginBottom:12}}>📁</div>
          <div style={{fontSize:15,fontWeight:700,color:T.sub,marginBottom:6}}>No Documents Yet</div>
          <div style={{fontSize:12,lineHeight:1.6}}>Tap <strong style={{color:T.orange}}>📁 New Folder</strong> to organize your files, or <strong style={{color:T.orange}}>📎 Upload Files</strong> to add documents directly.<br/><br/>You can upload PDFs, drawings, specs, permits, contracts — any file type.</div>
        </div>}
      </>}

      {/* Folder view — show docs inside this folder */}
      {!loading&&currentFolder&&<>
        <div style={{...cardS,marginBottom:12,background:currentFolder.color+"12",border:`1px solid ${currentFolder.color}30`,padding:"10px 14px"}}>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}>
            <div>
              <div style={{fontSize:14,fontWeight:800,color:T.text}}>{currentFolder.name}</div>
              {currentFolder.description&&<div style={{fontSize:11,color:T.muted,marginTop:2}}>{currentFolder.description}</div>}
            </div>
            <div style={{fontSize:13,fontWeight:700,color:currentFolder.color||T.orange}}>{currentDocs.length} item{currentDocs.length!==1?"s":""}</div>
          </div>
        </div>
        {currentDocs.map(doc=>(<DocRow key={doc.id} doc={doc} folders={folders} user={user} canAdmin={canAdmin}
          onDownload={()=>downloadDoc(doc)} canDownload={canDownload(doc)}
          onMove={fid=>moveDocToFolder(doc,fid)} onDelete={()=>removeDoc(doc.id)}
          getMimeIcon={getMimeIcon} fmtSize={fmtSize} docIcons={docIcons}
          onDragStart={()=>setDragDoc(doc.id)} onDragEnd={()=>{setDragDoc(null);setDragOverFolder(null);}}
          isDragging={dragDoc===doc.id}/>))}
        {currentDocs.length===0&&<div style={{textAlign:"center",padding:"40px 16px",color:T.muted}}>
          <div style={{fontSize:44,marginBottom:12}}>📁</div>
          <div style={{fontSize:14,fontWeight:700,color:T.sub,marginBottom:6}}>This folder is empty</div>
          <div style={{fontSize:12}}>Tap <strong style={{color:T.orange}}>📎 Upload Files</strong> above to add documents to <strong style={{color:T.text}}>"{currentFolder.name}"</strong>.</div>
        </div>}
      </>}
    </div>
  );
}

function DocRow({doc,folders,user,canAdmin,onDownload,canDownload,onMove,onDelete,getMimeIcon,fmtSize,docIcons,onDragStart,onDragEnd,isDragging}){
  const [showPreview,setShowPreview]=useState(false);
  const [showMove,setShowMove]=useState(false);   // must be before any conditional return

  const fname=(doc.file_name||doc.name||"").toLowerCase();
  const isImage=doc.file_type?.startsWith("image/")||(doc.file||"").startsWith("data:image")||[".jpg",".jpeg",".png",".gif",".webp",".bmp"].some(e=>fname.endsWith(e));
  const isPdf=doc.file_type==="application/pdf"||(doc.file||"").startsWith("data:application/pdf")||fname.endsWith(".pdf");
  const canPreview=isImage||isPdf;
  const hasFileData=!!(doc.file&&doc.file.length>100);

  if(showPreview){
    return(
      <div style={{position:"fixed",inset:0,zIndex:300,background:"rgba(0,0,0,0.97)",display:"flex",flexDirection:"column",fontFamily:"inherit"}}>
        {/* Header */}
        <div style={{background:"#141418",borderBottom:"1px solid #26262E",padding:"12px 16px",display:"flex",justifyContent:"space-between",alignItems:"center",flexShrink:0}}>
          <div style={{flex:1,minWidth:0}}>
            <div style={{fontSize:14,fontWeight:800,color:"#F0F4FF",overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{doc.name}</div>
            <div style={{fontSize:11,color:"#7080A0"}}>{doc.file_name||""}{doc.file_size?` · ${fmtSize(doc.file_size)}`:""}</div>
          </div>
          <div style={{display:"flex",gap:8,flexShrink:0,marginLeft:12}}>
            {canDownload&&<button onClick={onDownload}
              style={{background:"#34D399",color:"#000",border:"none",borderRadius:8,padding:"7px 14px",fontSize:12,fontWeight:700,cursor:"pointer",fontFamily:"inherit"}}>
              ⬇️ Download
            </button>}
            <button onClick={()=>setShowPreview(false)}
              style={{background:"#26262E",color:"#F0F4FF",border:"none",borderRadius:8,padding:"7px 14px",fontSize:12,fontWeight:700,cursor:"pointer",fontFamily:"inherit"}}>
              ✕ Close
            </button>
          </div>
        </div>
        <div style={{flex:1,overflow:"hidden",display:"flex",alignItems:"center",justifyContent:"center",padding:8}}>
          {isImage&&<img src={doc.file} alt={doc.name}
            style={{maxWidth:"100%",maxHeight:"100%",objectFit:"contain",borderRadius:8}}/>}

          {isPdf&&<iframe src={doc.file} title={doc.name}
            style={{width:"100%",height:"100%",border:"none",borderRadius:8,background:"#fff"}}/>}
        </div>
      </div>
    );
  }
  const ext=doc.file_name?.split(".").pop()?.toUpperCase()||"";
  const icon=getMimeIcon(doc.file_type||"")||docIcons[doc.doc_type]||"📁";
  return(
    <div draggable={!!onDragStart} onDragStart={onDragStart} onDragEnd={onDragEnd}
      style={{...cardS,marginBottom:8,padding:"10px 14px",
        opacity:isDragging?0.4:1,
        cursor:onDragStart?"grab":"default",
        border:isDragging?`2px solid ${T.orange}`:undefined,
        transition:"opacity 0.15s,border 0.15s"}}>
      {/* Drag handle hint */}
      {onDragStart&&<div style={{display:"flex",alignItems:"center",gap:4,marginBottom:6,color:T.muted,fontSize:10}}>
        <span style={{letterSpacing:"1px",fontSize:8}}>⠿⠿</span>
        <span>Drag to a folder</span>
      </div>}
      <div style={{display:"flex",alignItems:"flex-start",gap:10}}>
        <div style={{fontSize:24,flexShrink:0,marginTop:2}}>{icon}</div>
        <div style={{flex:1,minWidth:0}}>
          <div style={{fontSize:13,fontWeight:700,color:T.text,lineHeight:1.3}}>{doc.name}</div>
          <div style={{display:"flex",gap:6,flexWrap:"wrap",marginTop:4,alignItems:"center"}}>
            {ext&&<span style={{background:T.surface,borderRadius:4,padding:"1px 6px",fontSize:9,color:T.muted,fontWeight:700}}>{ext}</span>}
            {doc.doc_type&&<span style={{fontSize:10,color:T.muted}}>{doc.doc_type}</span>}
            {doc.file_size&&<span style={{fontSize:10,color:T.muted}}>{fmtSize(doc.file_size)}</span>}
            {doc.is_fillable&&<span style={{background:T.blueLow,border:`1px solid ${T.blue}40`,borderRadius:4,padding:"1px 6px",fontSize:9,color:T.blue,fontWeight:700}}>FILLABLE</span>}
          </div>
          {doc.notes&&<div style={{fontSize:11,color:T.muted,marginTop:4,fontStyle:"italic"}}>{doc.notes}</div>}
        </div>
      </div>
      <div style={{display:"flex",gap:6,marginTop:10,paddingTop:10,borderTop:`1px solid ${T.border}`,flexWrap:"wrap"}}>
        {canPreview&&<button onClick={()=>{
            if(!hasFileData){alert("File data not found — please delete and re-upload this document. (It was likely uploaded before the database was updated.)");return;}
            setShowPreview(true);
          }}
          style={{...primBtn,flex:2,borderRadius:10,fontSize:12,background:hasFileData?T.blue:"#26262E",color:hasFileData?"#fff":T.muted,padding:"8px"}}>
          👁 {hasFileData?"Preview":"Preview (re-upload)"}
        </button>}
        {canDownload&&doc.file&&<button onClick={onDownload}
          style={{...primBtn,flex:canPreview?1:2,borderRadius:10,fontSize:12,background:T.green,color:"#000",padding:"8px"}}>
          ⬇️
        </button>}
        {canAdmin&&folders.length>0&&<div style={{position:"relative",flex:1}}>
          <button onClick={()=>setShowMove(s=>!s)}
            style={{...ghostBtn,width:"100%",textAlign:"center",fontSize:12,padding:"8px"}}>📁 Move</button>
          {showMove&&<div style={{position:"absolute",bottom:"100%",left:0,right:0,background:T.card,border:`1px solid ${T.border}`,borderRadius:10,zIndex:50,overflow:"hidden",marginBottom:4}}>
            <div onClick={()=>{onMove(null);setShowMove(false);}} style={{padding:"8px 12px",fontSize:12,cursor:"pointer",color:T.muted,borderBottom:`1px solid ${T.border}`}}>📁 Root (no folder)</div>
            {folders.map(f=>(
              <div key={f.id} onClick={()=>{onMove(f.id);setShowMove(false);}}
                style={{padding:"8px 12px",fontSize:12,cursor:"pointer",color:T.text,borderBottom:`1px solid ${T.border}`,display:"flex",alignItems:"center",gap:6}}>
                <div style={{width:8,height:8,borderRadius:"50%",background:f.color||T.orange}}/>
                {f.name}
              </div>
            ))}
          </div>}
        </div>}
        {canAdmin&&<button onClick={onDelete}
          style={{...ghostBtn,fontSize:12,padding:"8px 12px",color:T.red,border:`1px solid ${T.red}30`}}>🗑</button>}
      </div>
    </div>
  );
}

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
    <div style={{background:T.greenLow,border:`1px solid ${T.green}40`,borderRadius:12,padding:"10px 14px",marginBottom:14,fontSize:12,color:T.green,lineHeight:1.5}}>
      <strong>⚡ Auto-filled from daily reports</strong> — hours are added automatically when a foreman submits a daily report. Manual entries below for anything not on a daily.
    </div>
    <button onClick={()=>setShowForm(!showForm)} style={{...primBtn,marginBottom:14,borderRadius:14}}>{showForm?"✕ Cancel":"📅 + Add Milestone"}</button>
    {showForm&&<div style={{...cardS,marginBottom:14,borderLeft:`3px solid ${T.blue}`}}>
      <div style={{marginBottom:10}}><label style={lbl}>Milestone Title *</label><input type="text" placeholder="e.g. HDD Bore Complete" value={f.title} onChange={e=>set("title",e.target.value)} style={inp}/></div>
      <div style={{marginBottom:10}}><label style={lbl}>Description</label><input type="text" placeholder="Optional details…" value={f.description} onChange={e=>set("description",e.target.value)} style={inp}/></div>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:10}}><div><label style={lbl}>Target Date</label><input type="date" value={f.target_date} onChange={e=>set("target_date",e.target.value)} style={inp}/></div><div><label style={lbl}>Status</label><select value={f.status} onChange={e=>set("status",e.target.value)} style={inpSel}><option value="pending">Pending</option><option value="in_progress">In Progress</option><option value="completed">Completed</option><option value="delayed">Delayed</option></select></div></div>
      <button onClick={save} style={{...primBtn,background:T.blue,borderRadius:12}}>{saving?"Saving…":"Save Milestone"}</button>
    </div>}
    {loading&&<Spinner/>}
    {!loading&&total>0&&<div style={{...cardS,marginBottom:14}}><div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:8}}><div style={{fontSize:13,fontWeight:700}}>Progress</div><div style={{fontSize:13,fontWeight:700,color:T.green}}>{completed}/{total} complete</div></div><div style={{height:8,background:T.border,borderRadius:4,overflow:"hidden"}}><div style={{height:"100%",background:`linear-gradient(90deg,${T.orange},${T.green})`,borderRadius:4,width:`${(completed/total)*100}%`,transition:"width 0.4s"}}/></div></div>}
    {!loading&&milestones.length===0&&!showForm&&<div style={{textAlign:"center",padding:"40px 0",color:T.muted}}><div style={{fontSize:32,marginBottom:8}}>📅</div><div>No milestones yet.</div></div>}
    {!loading&&milestones.map(m=>{const du=daysUntil(m.target_date);const overdue=du!==null&&du<0&&m.status!=="completed";const dueSoon=du!==null&&du>=0&&du<=7&&m.status!=="completed";return(<div key={m.id} style={{...cardS,marginBottom:10,borderLeft:`3px solid ${statusColor[m.status]||T.border}`,opacity:m.status==="completed"?0.7:1}}><div style={{display:"flex",alignItems:"flex-start",gap:12}}><button onClick={()=>toggleStatus(m)} style={{width:32,height:32,borderRadius:"50%",border:`2px solid ${statusColor[m.status]||T.border}`,background:m.status==="completed"?T.green:T.surface,color:m.status==="completed"?"#0D0D0F":statusColor[m.status]||T.muted,cursor:"pointer",display:"flex",alignItems:"center",justifyContent:"center",fontSize:14,fontWeight:800,flexShrink:0,marginTop:2}}>{statusIcon[m.status]||"○"}</button><div style={{flex:1,minWidth:0}}><div style={{fontSize:15,fontWeight:700,textDecoration:m.status==="completed"?"line-through":"none",color:m.status==="completed"?T.muted:T.text}}>{m.title}</div>{m.description&&<div style={{fontSize:12,color:T.sub,marginTop:2}}>{m.description}</div>}<div style={{display:"flex",gap:8,marginTop:6,flexWrap:"wrap"}}><span style={pill(statusColor[m.status]||T.muted)}>{m.status.replace("_"," ").toUpperCase()}</span>{m.target_date&&<span style={pill(overdue?T.red:dueSoon?T.yellow:T.muted)}>{overdue?`${Math.abs(du)}d overdue`:dueSoon?`Due in ${du}d`:`Target: ${fmtDate(m.target_date)}`}</span>}{m.completed_date&&<span style={pill(T.green)}>Done: {fmtDate(m.completed_date)}</span>}</div></div><button onClick={()=>del(m.id)} style={{background:"none",border:"none",color:T.muted,cursor:"pointer",fontSize:16,padding:0,flexShrink:0}}>🗑</button></div></div>);})}
  </div>);
}

function PhotosTab({projectId,photos,onRefresh,onErr}){
  const CATEGORIES=["Progress","Safety","Equipment","Issue/Deficiency","Before","After","Inspection","Other"];
  const [caption,setCaption]=useState("");
  const [category,setCategory]=useState("Progress");
  const [saving,setSaving]=useState(false);
  const [lb,setLb]=useState(null);
  const [filterCat,setFilterCat]=useState("All");
  const [annotating,setAnnotating]=useState(null); // photo src being annotated
  const fileRef=useRef(null);
  const annotCanvasRef=useRef(null);
  const [annotColor,setAnnotColor]=useState("#FF0000");
  const [annotMode,setAnnotMode]=useState("draw"); // draw | text
  const [annotText,setAnnotText]=useState("");
  const [annotDrawing,setAnnotDrawing]=useState(false);

  async function getGeoLocation(){
    return new Promise(res=>{
      if(!navigator.geolocation){res({lat:null,lng:null});return;}
      navigator.geolocation.getCurrentPosition(
        pos=>res({lat:pos.coords.latitude,lng:pos.coords.longitude}),
        ()=>res({lat:null,lng:null}),
        {timeout:5000}
      );
    });
  }

  async function handleFiles(files){
    setSaving(true);
    try{
      const {lat,lng}=await getGeoLocation();
      for(const f of files){
        if(!f.type.startsWith("image/"))continue;
        const src=await compressImg(f,1100,0.72);
        await API.photos.create({
          project_id:projectId,src,caption,date:today(),
          category,
          lat:lat||null,lng:lng||null,
        });
      }
      await onRefresh();setCaption("");
    }catch(e){onErr(e.message);}
    setSaving(false);
  }

  async function del(id){
    try{await API.photos.remove(id);await onRefresh();}
    catch(e){onErr(e.message);}
  }

  function annotStart(e){
    const c=annotCanvasRef.current;if(!c)return;
    const r=c.getBoundingClientRect();
    const x=(e.touches?e.touches[0].clientX:e.clientX)-r.left;
    const y=(e.touches?e.touches[0].clientY:e.clientY)-r.top;
    const ctx=c.getContext("2d");
    ctx.beginPath();ctx.moveTo(x*(c.width/r.width),y*(c.height/r.height));
    c._drawing=true;
  }
  function annotDraw(e){
    e.preventDefault();
    const c=annotCanvasRef.current;if(!c||!c._drawing)return;
    const r=c.getBoundingClientRect();
    const x=(e.touches?e.touches[0].clientX:e.clientX)-r.left;
    const y=(e.touches?e.touches[0].clientY:e.clientY)-r.top;
    const ctx=c.getContext("2d");
    ctx.lineWidth=3;ctx.lineCap="round";ctx.strokeStyle=annotColor;
    ctx.lineTo(x*(c.width/r.width),y*(c.height/r.height));
    ctx.stroke();ctx.beginPath();ctx.moveTo(x*(c.width/r.width),y*(c.height/r.height));
  }
  function annotEnd(e){const c=annotCanvasRef.current;if(c)c._drawing=false;}

  async function saveAnnotation(){
    const c=annotCanvasRef.current;if(!c)return;
    const img=new Image();img.crossOrigin="anonymous";
    img.onload=async()=>{
      const merge=document.createElement("canvas");
      merge.width=img.width;merge.height=img.height;
      const mctx=merge.getContext("2d");
      mctx.drawImage(img,0,0);
      mctx.drawImage(c,0,0,img.width,img.height);
      const annotated=merge.toDataURL("image/jpeg",0.85);
      const {lat,lng}=await getGeoLocation();
      await API.photos.create({project_id:projectId,src:annotated,caption:`Annotated · ${caption||"Photo"}`,date:today(),category,lat,lng});
      await onRefresh();
      setAnnotating(null);
    };
    img.src=annotating;
  }

  const filtered=filterCat==="All"?photos:photos.filter(p=>p.category===filterCat);

  const byDate={};
  filtered.forEach(p=>{
    const d=p.date||"Unknown";
    if(!byDate[d])byDate[d]=[];
    byDate[d].push(p);
  });
  const sortedDates=Object.keys(byDate).sort().reverse();

  const catColor={Progress:T.blue,Safety:T.red,Equipment:T.yellow,"Issue/Deficiency":T.red,Before:T.purple,After:T.green,Inspection:T.orange,Other:T.muted};

  if(annotating) return(
    <div style={{background:T.bg,minHeight:"100vh",fontFamily:"inherit"}}>
      <div style={{background:"#1a1a1a",padding:"12px 16px",display:"flex",justifyContent:"space-between",alignItems:"center",borderBottom:`1px solid ${T.border}`}}>
        <div style={{fontSize:15,fontWeight:800,color:T.text}}>✏️ Annotate Photo</div>
        <div style={{display:"flex",gap:8}}>
          {["#FF0000","#FFFF00","#00FF00","#FFFFFF","#000000"].map(c=>(
            <button key={c} onClick={()=>setAnnotColor(c)}
              style={{width:24,height:24,borderRadius:"50%",background:c,border:annotColor===c?"3px solid #60A5FA":"2px solid #555",cursor:"pointer"}}/>
          ))}
        </div>
      </div>
      <div style={{position:"relative",display:"inline-block",width:"100%"}}>
        <img src={annotating} style={{width:"100%",display:"block"}} alt="annotate"/>
        <canvas ref={annotCanvasRef} width={800} height={600}
          style={{position:"absolute",top:0,left:0,width:"100%",height:"100%",touchAction:"none",cursor:"crosshair"}}
          onMouseDown={annotStart} onMouseMove={annotDraw} onMouseUp={annotEnd}
          onTouchStart={annotStart} onTouchMove={annotDraw} onTouchEnd={annotEnd}
        />
      </div>
      <div style={{display:"flex",gap:10,padding:"12px 16px"}}>
        <button onClick={()=>{const c=annotCanvasRef.current;if(c)c.getContext("2d").clearRect(0,0,c.width,c.height);}}
          style={{...ghostBtn,flex:1,textAlign:"center"}}>🗑 Clear</button>
        <button onClick={()=>setAnnotating(null)} style={{...ghostBtn,flex:1,textAlign:"center"}}>Cancel</button>
        <button onClick={saveAnnotation} style={{...primBtn,flex:2,borderRadius:12}}>✅ Save Annotated</button>
      </div>
    </div>
  );

  return(
    <div>
      <Lightbox src={lb} onClose={()=>setLb(null)}/>

      {/* Upload card */}
      <div style={{...cardS,marginBottom:14,borderStyle:"dashed",borderColor:T.orange+"44"}}>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8,marginBottom:10}}>
          <div>
            <label style={lbl}>Category</label>
            <select value={category} onChange={e=>setCategory(e.target.value)} style={inpSel}>
              {CATEGORIES.map(c=><option key={c}>{c}</option>)}
            </select>
          </div>
          <div>
            <label style={lbl}>Caption</label>
            <input type="text" placeholder="Optional caption" value={caption} onChange={e=>setCaption(e.target.value)} style={inp}/>
          </div>
        </div>
        <div style={{fontSize:10,color:T.muted,marginBottom:8}}>📍 GPS location will be automatically captured</div>
        <button onClick={()=>fileRef.current?.click()} style={{...primBtn}}>
          {saving?"Uploading…":"📷 Add Site Photos"}
        </button>
        <input ref={fileRef} type="file" accept="image/*" capture="environment" multiple style={{display:"none"}}
          onChange={e=>{handleFiles(Array.from(e.target.files));e.target.value="";}}/>
      </div>
      {photos.length>0&&<div style={{display:"flex",gap:6,overflowX:"auto",paddingBottom:6,marginBottom:12,WebkitOverflowScrolling:"touch"}}>
        {["All",...CATEGORIES].map(c=>(
          <button key={c} onClick={()=>setFilterCat(c)}
            style={{flexShrink:0,padding:"5px 10px",borderRadius:8,border:`1px solid ${filterCat===c?(catColor[c]||T.orange):T.border}`,background:filterCat===c?`${catColor[c]||T.orange}15`:T.card,color:filterCat===c?(catColor[c]||T.orange):T.muted,fontSize:11,fontWeight:700,cursor:"pointer",fontFamily:"inherit",whiteSpace:"nowrap"}}>
            {c}{c!=="All"?` (${photos.filter(p=>p.category===c).length})`:`(${photos.length})`}
          </button>
        ))}
      </div>}

      {filtered.length===0&&<div style={{textAlign:"center",padding:"40px 16px",color:T.muted}}>
        <div style={{fontSize:44,marginBottom:12}}>📷</div>
        <div style={{fontSize:15,fontWeight:700,color:T.sub,marginBottom:6}}>{filterCat==="All"?"No Photos Yet":`No ${filterCat} Photos`}</div>
        <div style={{fontSize:12,color:T.muted,lineHeight:1.6}}>{filterCat==="All"?"Tap 📷 Add Site Photos above. GPS location and category are captured automatically.":"Try a different category filter or add new photos above."}</div>
      </div>}

      {/* Photos grouped by date */}
      {sortedDates.map(date=>(
        <div key={date} style={{marginBottom:20}}>
          <div style={{fontSize:11,fontWeight:700,color:T.muted,textTransform:"uppercase",letterSpacing:"1px",marginBottom:8,paddingBottom:4,borderBottom:`1px solid ${T.border}`}}>
            📅 {date}  <span style={{color:T.border}}>({byDate[date].length} photo{byDate[date].length!==1?"s":""})</span>
          </div>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8}}>
            {byDate[date].map(p=>(
              <div key={p.id} style={{position:"relative",borderRadius:12,overflow:"hidden",aspectRatio:"4/3",background:T.card}}>
                <img src={p.src} alt={p.caption} onClick={()=>setLb(p.src)}
                  style={{width:"100%",height:"100%",objectFit:"cover",cursor:"pointer",display:"block"}}/>
                {/* Category badge */}
                {p.category&&<div style={{position:"absolute",top:6,left:6,background:"rgba(0,0,0,0.75)",borderRadius:6,padding:"2px 7px",fontSize:9,fontWeight:700,color:catColor[p.category]||T.muted}}>
                  {p.category}
                </div>}
                {/* GPS badge */}
                {p.lat&&<div style={{position:"absolute",bottom:24,left:6,background:"rgba(0,0,0,0.75)",borderRadius:6,padding:"2px 6px",fontSize:9,color:"#4ade80"}}>
                  📍 GPS
                </div>}
                {p.caption&&<div style={{position:"absolute",bottom:0,left:0,right:0,background:"linear-gradient(transparent,rgba(0,0,0,0.8))",padding:"16px 8px 6px",fontSize:10,color:"#fff"}}>
                  {p.caption}
                </div>}
                {/* Action buttons */}
                <div style={{position:"absolute",top:6,right:6,display:"flex",gap:4}}>
                  <button onClick={()=>setAnnotating(p.src)}
                    style={{background:"rgba(0,0,0,0.75)",border:"none",color:"#fff",borderRadius:6,width:24,height:24,fontSize:11,cursor:"pointer",display:"flex",alignItems:"center",justifyContent:"center"}}>✏️</button>
                  <button onClick={()=>del(p.id)}
                    style={{background:"rgba(0,0,0,0.75)",border:"none",color:"#fff",borderRadius:"50%",width:24,height:24,fontSize:12,cursor:"pointer",display:"flex",alignItems:"center",justifyContent:"center"}}>✕</button>
                </div>
              </div>
            ))}
          </div>
        </div>
      ))}
    </div>
  );
}

/* ── DRAWINGS: sheet list + upload ───────────────────────────── */
/* ── SIGNATURE PACKAGE: batch several reports into one signature ── */
function SignaturePackageScreen({project,user,onBack,onErr}){
  const [dailies,setDailies]=useState([]);
  const [tickets,setTickets]=useState([]);
  const [loading,setLoading]=useState(true);
  const [sel,setSel]=useState({});           // id -> {type,row}
  const [name,setName]=useState("");
  const [email,setEmail]=useState("");
  const [sending,setSending]=useState(false);
  const [sent,setSent]=useState(false);
  const [err,setErr]=useState("");

  const div=project.division;
  const m=(n)=>"$"+Number(n||0).toLocaleString("en-US",{minimumFractionDigits:2,maximumFractionDigits:2});

  useEffect(()=>{(async()=>{
    setLoading(true);
    try{
      const [d,t]=await Promise.all([
        API.reports.forProject(project.id).catch(()=>[]),
        API.tmTickets.forProject(project.id).catch(()=>[]),
      ]);
      setDailies((d||[]).filter(r=>r.status!=="draft"));
      setTickets(t||[]);
    }catch(e){onErr&&onErr(e.message);}
    setLoading(false);
  })();},[project.id]);

  const toggle=(type,row)=>setSel(s=>{
    const n={...s};
    if(n[row.id])delete n[row.id]; else n[row.id]={type,row};
    return n;
  });

  const chosen=Object.values(sel);
  const total=chosen.reduce((s,{type,row})=>{
    if(type==="tm")return s+(parseFloat(row.grand_total)||0);
    const t=reportTotals(row,div);return s+(t.grand||0);
  },0);

  function buildItem({type,row}){
    if(type==="tm"){
      return {
        docType:"tm",projectName:project.name,customer:project.client||"",
        poNumber:project.work_order||"",afeNumber:project.afe||"",location:project.location||"",
        description:row.description||"",reportNo:row.ticket_no||"",reportDate:row.ticket_date||"",
        submittedBy:row.submitted_by||"",
        lineItems:{
          labor:(row.labor||[]).map(r=>({name:r.name||"",classification:r.classification||"",
            hours:r.hours||0,rate:m(r.rate),amount:m((parseFloat(r.hours)||0)*(parseFloat(r.rate)||0))})),
          equipment:(row.equipment||[]).map(r=>({description:r.description||"",unit:r.unit||"",
            qty:r.qty||0,rate:m(r.rate),amount:m((parseFloat(r.qty)||0)*(parseFloat(r.rate)||0))})),
          rental:[],
          materials:(row.materials||[]).map(r=>({description:r.description||"",qty:r.qty||0,
            unit:r.unit||"",unit_price:m(r.unit_price),
            amount:m((parseFloat(r.qty)||0)*(parseFloat(r.unit_price)||0))})),
          other:(row.other_charges||[]).map(r=>({description:r.description||"",amount:m(r.amount)})),
        },
        grandTotal:m(row.grand_total),
      };
    }
    const t=reportTotals(row,div);
    return {
      docType:"daily",projectName:project.name,customer:project.client||"",
      poNumber:project.work_order||"",afeNumber:project.afe||"",location:project.location||"",
      description:row.description||"",reportNo:row.report_no||"",reportDate:row.date||"",
      submittedBy:row.submitted_by||"",
      lineItems:{
        labor:(row.labor||[]).map(l=>({name:l.name||"",classification:l.classification||"",
          hours:((parseFloat(l.regHrs)||0)+(parseFloat(l.otHrs)||0)+(parseFloat(l.travelHrs)||0)).toFixed(1),
          rate:"",amount:m(laborAmt(l,div))})),
        equipment:(row.equipment||[]).map(e=>({description:e.description||"",unit:e.unit||"",
          qty:e.qty||0,rate:e.rate?m(e.rate):"",amount:m(equipAmt(e,div))})),
        rental:(row.rental_equipment||[]).map(r=>({description:r.description||"",qty:r.qty||0,
          rate:m(r.rate),amount:m((parseFloat(r.qty)||0)*(parseFloat(r.rate)||0)*(parseFloat(r.usage)||1))})),
        materials:(row.materials||[]).map(x=>({description:x.description||"",qty:x.qty||"",
          unit:"",unit_price:"",amount:m(x.amount)})),
        other:[],
      },
      laborTotal:m(t.labor),equipmentTotal:m(t.equip),rentalTotal:m(t.rental),
      materialsTotal:m(t.mats),grandTotal:m(t.grand),
    };
  }

  async function send(){
    if(!name.trim()||!email.trim()){setErr("Client name and email are required.");return;}
    if(!chosen.length){setErr("Select at least one document.");return;}
    setSending(true);setErr("");
    try{
      const items=chosen
        .sort((a,b)=>String(a.row.date||a.row.ticket_date||"").localeCompare(String(b.row.date||b.row.ticket_date||"")))
        .map(buildItem);
      // receipts attached to T&M material rows ride along as extra pages
      const attachments=chosen.flatMap(({type,row})=>
        type==="tm"
          ?(row.materials||[]).map(x=>x.attachment_url).filter(Boolean)
          :(row.materials||[]).flatMap(x=>(x.receipts||[]).map(r=>r.src)).filter(Boolean));
      const dates=items.map(i=>i.reportDate).filter(Boolean).sort();
      const label=dates.length>1?`${dates[0]}_to_${dates[dates.length-1]}`:(dates[0]||"package");

      const res=await fetch("/.netlify/functions/box-sign-create",{
        method:"POST",headers:{"Content-Type":"application/json"},
        body:JSON.stringify({
          docType:"package",reportId:project.id,
          inspectorEmail:email.trim(),inspectorName:name.trim(),
          projectName:project.name,customer:project.client||"",
          reportNo:label,reportDate:dates[dates.length-1]||today(),
          items,attachments,
        }),
      });
      const data=await res.json();
      if(!res.ok)throw new Error(data.error||"Send failed");
      setSent(true);
    }catch(e){setErr(e.message);}
    setSending(false);
  }

  const Row=({type,row,title,sub,amt})=>(
    <div onClick={()=>toggle(type,row)} style={{...cardS,marginBottom:8,display:"flex",alignItems:"center",
      gap:12,cursor:"pointer",border:sel[row.id]?`1px solid ${T.green}`:`1px solid ${T.border}`,
      background:sel[row.id]?T.greenLow:T.card}}>
      <div style={{width:20,height:20,borderRadius:6,flexShrink:0,display:"flex",alignItems:"center",
        justifyContent:"center",fontSize:13,fontWeight:900,
        background:sel[row.id]?T.green:"transparent",color:"#000",
        border:sel[row.id]?"none":`1.5px solid ${T.border}`}}>{sel[row.id]?"✓":""}</div>
      <div style={{flex:1,minWidth:0}}>
        <div style={{fontSize:13,fontWeight:700,color:T.text}}>{title}</div>
        <div style={{fontSize:11,color:T.muted,marginTop:2}}>{sub}</div>
      </div>
      <div style={{fontSize:13,fontWeight:800,color:T.green,flexShrink:0}}>{amt}</div>
    </div>
  );

  if(sent)return(
    <div style={{background:T.bg,minHeight:"100vh",fontFamily:"inherit",padding:"20px 16px"}}>
      <div style={{...cardS,textAlign:"center",padding:32,borderLeft:`3px solid ${T.green}`}}>
        <div style={{fontSize:44,marginBottom:10}}>📦</div>
        <div style={{fontSize:18,fontWeight:800,color:T.green,marginBottom:6}}>Package Sent</div>
        <div style={{fontSize:13,color:T.sub,lineHeight:1.6}}>
          {chosen.length} document{chosen.length!==1?"s":""} combined into one PDF and sent to {name} for signature.
        </div>
        <button onClick={onBack} style={{...primBtn,borderRadius:12,marginTop:18}}>Done</button>
      </div>
    </div>
  );

  return(
    <div style={{background:T.bg,minHeight:"100vh",fontFamily:"inherit"}}>
      <div style={{background:T.surface,borderBottom:`1px solid ${T.border}`,padding:"12px 16px",position:"sticky",top:0,zIndex:50}}>
        <button onClick={onBack} style={{background:"none",border:"none",color:T.sub,fontSize:13,cursor:"pointer",fontFamily:"inherit",marginBottom:4,padding:0}}>← Back</button>
        <div style={{fontSize:15,fontWeight:900,color:T.text}}>📦 Send Package for Signature</div>
        <div style={{fontSize:11,color:T.muted,marginTop:2}}>
          Pick any reports and tickets — they go out as one PDF, signed once.
        </div>
      </div>

      <div style={{padding:"16px 16px 90px"}}>
        {err&&<div style={{background:T.redLow,border:`1px solid ${T.red}40`,borderRadius:10,padding:"10px 14px",marginBottom:12,fontSize:12,color:T.red}}>{err}</div>}

        {loading?<div style={{textAlign:"center",color:T.muted,padding:24,fontSize:13}}>Loading documents…</div>:<>
          {dailies.length>0&&<>
            <div style={{fontSize:11,fontWeight:700,color:T.muted,textTransform:"uppercase",letterSpacing:"1px",margin:"4px 0 8px"}}>Daily Reports</div>
            {dailies.map(r=>(<Row key={r.id} type="daily" row={r}
              title={`${r.date||"—"}${r.report_no?`  ·  #${r.report_no}`:""}`}
              sub={r.submitted_by||""} amt={m(reportTotals(r,div).grand)}/>))}
          </>}

          {tickets.length>0&&<>
            <div style={{fontSize:11,fontWeight:700,color:T.muted,textTransform:"uppercase",letterSpacing:"1px",margin:"16px 0 8px"}}>T&amp;M Tickets</div>
            {tickets.map(t=>(<Row key={t.id} type="tm" row={t}
              title={`${t.ticket_date||"—"}  ·  #${t.ticket_no||""}`}
              sub={t.submitted_by||""} amt={m(t.grand_total)}/>))}
          </>}

          {!dailies.length&&!tickets.length&&(
            <div style={{...cardS,textAlign:"center",padding:28,color:T.muted}}>
              <div style={{fontSize:30,marginBottom:8}}>📭</div>
              <div style={{fontSize:13}}>No submitted reports or tickets on this job yet.</div>
            </div>
          )}

          <div style={{marginTop:20}}>
            <label style={lbl}>Client Contact Name *</label>
            <input value={name} onChange={e=>setName(e.target.value)} placeholder="Who signs this" style={{...inp,marginBottom:12}}/>
            <label style={lbl}>Client Email *</label>
            <input type="email" value={email} onChange={e=>setEmail(e.target.value)} placeholder="name@company.com" style={inp}/>
          </div>
        </>}
      </div>

      {!loading&&(dailies.length>0||tickets.length>0)&&(
        <div style={{position:"fixed",left:0,right:0,bottom:0,background:T.surface,borderTop:`1px solid ${T.border}`,padding:"10px 16px",zIndex:60}}>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:8}}>
            <span style={{fontSize:12,color:T.sub}}>{chosen.length} selected</span>
            <span style={{fontSize:15,fontWeight:900,color:T.green}}>{m(total)}</span>
          </div>
          <button onClick={send} disabled={sending||!chosen.length}
            style={{...primBtn,borderRadius:12,background:chosen.length?"#1f3864":T.border,
              color:chosen.length?"#fff":T.muted,opacity:sending?0.6:1}}>
            {sending?"Building package…":`📦 Send ${chosen.length||""} for Signature`}
          </button>
        </div>
      )}
    </div>
  );
}

function DrawingsTab({projectId,user,onErr}){
  const canAdmin=user.role==="admin"||user.role==="pm";
  const [drawings,setDrawings]=useState([]);
  const [loading,setLoading]=useState(true);
  const [uploading,setUploading]=useState(false);
  const [progress,setProgress]=useState("");
  const [viewing,setViewing]=useState(null);
  const [q,setQ]=useState("");
  const fileRef=useRef(null);

  const DISCIPLINES=["Mechanical","Piping","Structural","Electrical","Civil","Architectural","P&ID","Isometric","Other"];

  async function load(){
    setLoading(true);
    try{ setDrawings(await API.drawings.forProject(projectId)||[]); }
    catch(e){ onErr&&onErr(e.message); }
    setLoading(false);
  }
  useEffect(()=>{load();},[projectId]);

  async function handleFiles(files){
    if(!files||!files.length)return;
    setUploading(true);
    for(let i=0;i<files.length;i++){
      const file=files[i];
      setProgress(`Uploading ${i+1} of ${files.length}: ${file.name}`);
      try{
        if(file.type!=="application/pdf"){ onErr&&onErr(`${file.name} is not a PDF — skipped.`); continue; }
        const clean=file.name.replace(/[^A-Za-z0-9._-]/g,"_");
        const path=`${projectId}/${Date.now()}-${clean}`;
        await storageUpload("drawings",path,file,"application/pdf");
        await API.drawings.create({
          project_id:projectId,
          title:file.name.replace(/\.pdf$/i,""),
          storage_path:path,
          file_size:file.size,
          uploaded_by:user.name,
        });
      }catch(e){ onErr&&onErr(`${file.name}: ${e.message}`); }
    }
    setProgress(""); setUploading(false);
    if(fileRef.current)fileRef.current.value="";
    await load();
  }

  async function del(d){
    if(!window.confirm(`Delete "${d.title}"? This cannot be undone.`))return;
    try{
      await storageRemove("drawings",d.storage_path);
      await API.drawings.remove(d.id);
      await load();
    }catch(e){ onErr&&onErr(e.message); }
  }

  const fmtSize=(b)=>!b?"":b<1048576?(b/1024).toFixed(0)+" KB":(b/1048576).toFixed(1)+" MB";
  const filtered=drawings.filter(d=>{
    if(!q.trim())return true;
    const s=q.toLowerCase();
    return (d.title||"").toLowerCase().includes(s)||(d.sheet_number||"").toLowerCase().includes(s)||(d.discipline||"").toLowerCase().includes(s);
  });

  if(viewing) return <DrawingViewer drawing={viewing} user={user} onBack={()=>{setViewing(null);load();}} onErr={onErr}/>;

  return(
    <div>
      {canAdmin&&<>
        <input ref={fileRef} type="file" accept="application/pdf" multiple style={{display:"none"}}
          onChange={e=>handleFiles(Array.from(e.target.files||[]))}/>
        <button onClick={()=>fileRef.current&&fileRef.current.click()} disabled={uploading}
          style={{...primBtn,borderRadius:14,marginBottom:12,background:T.orange,color:"#000",opacity:uploading?0.6:1}}>
          {uploading?"Uploading…":"📐 Upload Drawings (PDF)"}
        </button>
        {progress&&<div style={{fontSize:11,color:T.muted,textAlign:"center",marginBottom:10}}>{progress}</div>}
      </>}

      {drawings.length>3&&<input value={q} onChange={e=>setQ(e.target.value)} placeholder="Search sheets…"
        style={{...inp,marginBottom:12}}/>}

      {loading?<div style={{textAlign:"center",color:T.muted,padding:24,fontSize:13}}>Loading drawings…</div>:
       filtered.length===0?(
        <div style={{...cardS,textAlign:"center",padding:28,color:T.muted}}>
          <div style={{fontSize:32,marginBottom:8}}>📐</div>
          <div style={{fontSize:14,fontWeight:700,color:T.sub,marginBottom:4}}>No Drawings</div>
          <div style={{fontSize:12}}>{canAdmin?"Upload a PDF drawing set to get started.":"No drawings have been uploaded yet."}</div>
        </div>
      ):filtered.map(d=>(
        <div key={d.id} style={{...cardS,marginBottom:10}}>
          <div onClick={()=>setViewing(d)} style={{cursor:"pointer"}}>
            <div style={{fontSize:15,fontWeight:800,color:"#60A5FA",lineHeight:1.35,wordBreak:"break-word"}}>
              {d.sheet_number?`${d.sheet_number} · `:""}{d.title}
            </div>
            <div style={{fontSize:12,color:T.sub,marginTop:6,lineHeight:1.5}}>
              {[d.discipline,d.revision?`Rev ${d.revision}`:null,d.page_count?`${d.page_count} sheet${d.page_count!==1?"s":""}`:null,fmtSize(d.file_size)]
                .filter(Boolean).join("  ·  ")||"PDF"}
            </div>
            <div style={{fontSize:11,color:T.muted,marginTop:4}}>
              {d.uploaded_by||""}{d.created_at?`  ·  ${new Date(d.created_at).toLocaleDateString()}`:""}
            </div>
          </div>
          <div style={{display:"flex",gap:8,marginTop:12}}>
            <button onClick={()=>setViewing(d)}
              style={{...primBtn,width:"auto",flex:1,padding:"11px",fontSize:13,borderRadius:10,background:"#1f3864"}}>Open</button>
            {canAdmin&&<button onClick={()=>del(d)} title="Delete"
              style={{...primBtn,width:"auto",flex:"0 0 auto",padding:"11px 16px",fontSize:13,borderRadius:10,
                background:T.redLow,color:T.red,border:`1px solid ${T.red}30`}}>🗑</button>}
          </div>
        </div>
      ))}
    </div>
  );
}

/* ── DRAWING VIEWER: PDF render + redline markup ─────────────── */
function DrawingViewer({drawing,user,onBack,onErr}){
  const [pdf,setPdf]=useState(null);
  const [page,setPage]=useState(1);
  const [pages,setPages]=useState(0);
  const [scale,setScale]=useState(1);
  const [offset,setOffset]=useState({x:0,y:0});
  const [loading,setLoading]=useState(true);
  const [err,setErr]=useState("");
  const [rendering,setRendering]=useState(false);
  const [barOpen,setBarOpen]=useState(true);

  // Markup state
  const [tool,setTool]=useState("pan");           // pan | pen | arrow | box | text | erase
  const [color,setColor]=useState("#EF4444");
  const [shapes,setShapes]=useState([]);          // current page, PDF user-space coords
  const [markupId,setMarkupId]=useState(null);    // row id for this drawing+page
  const [showMarkup,setShowMarkup]=useState(true);
  const [toolsOpen,setToolsOpen]=useState(false);
  const [dirty,setDirty]=useState(false);
  const [saving,setSaving]=useState(false);
  const [baseW,setBaseW]=useState(0);             // page width at scale 1 = the coord system

  const canvasRef=useRef(null);
  const overlayRef=useRef(null);
  const wrapRef=useRef(null);
  const renderTaskRef=useRef(null);
  const gesture=useRef({mode:null});
  const draftRef=useRef(null);
  const fitRef=useRef(1);

  const COLORS=["#EF4444","#FBBF24","#22C55E","#3B82F6","#111827"];
  const clamp=(v,lo,hi)=>Math.max(lo,Math.min(hi,v));
  const dist=(a,b)=>Math.hypot(a.clientX-b.clientX,a.clientY-b.clientY);

  /* ── load PDF ── */
  useEffect(()=>{
    let cancelled=false;
    (async()=>{
      try{
        setLoading(true);setErr("");
        const pdfjsLib=await loadPdfJs();
        const publicUrl=storagePublicUrl("drawings",drawing.storage_path);
        const doc=await pdfjsLib.getDocument({url:publicUrl}).promise;
        if(cancelled)return;
        setPdf(doc);setPages(doc.numPages);setPage(1);
        if(!drawing.page_count||drawing.page_count!==doc.numPages)
          API.drawings.update(drawing.id,{page_count:doc.numPages}).catch(()=>{});
      }catch(e){ if(!cancelled)setErr("Could not open drawing: "+e.message); }
      if(!cancelled)setLoading(false);
    })();
    return()=>{cancelled=true;};
  },[drawing.id]);

  /* ── load markup for the current page ── */
  useEffect(()=>{
    let cancelled=false;
    (async()=>{
      try{
        const rows=await API.markups.forDrawing(drawing.id);
        if(cancelled)return;
        const row=(rows||[]).find(r=>r.page===page);
        setShapes(row&&Array.isArray(row.shapes)?row.shapes:[]);
        setMarkupId(row?row.id:null);
        setDirty(false);
      }catch(e){ if(!cancelled){setShapes([]);setMarkupId(null);} }
    })();
    return()=>{cancelled=true;};
  },[drawing.id,page]);

  /* ── render PDF page ── */
  useEffect(()=>{
    if(!pdf)return;
    let cancelled=false;
    (async()=>{
      try{
        setRendering(true);
        const pg=await pdf.getPage(page);
        if(cancelled)return;
        const dpr=window.devicePixelRatio||1;
        const base=pg.getViewport({scale:1});
        const avail=(wrapRef.current?wrapRef.current.clientWidth:600)-8;
        const fit=avail/base.width;
        fitRef.current=fit;
        setBaseW(base.width);
        const vp=pg.getViewport({scale:fit*scale*dpr});
        const cv=canvasRef.current, ov=overlayRef.current;
        if(!cv)return;
        for(const c of [cv,ov]){
          if(!c)continue;
          c.width=vp.width;c.height=vp.height;
          c.style.width=(vp.width/dpr)+"px";c.style.height=(vp.height/dpr)+"px";
        }
        if(renderTaskRef.current){try{renderTaskRef.current.cancel();}catch(e){}}
        const task=pg.render({canvasContext:cv.getContext("2d"),viewport:vp});
        renderTaskRef.current=task;
        await task.promise;
      }catch(e){ if(e&&e.name!=="RenderingCancelledException"&&!cancelled)setErr(e.message); }
      if(!cancelled)setRendering(false);
    })();
    return()=>{cancelled=true;};
  },[pdf,page,scale]);

  /* ── draw markup overlay ── */
  const paint=React.useCallback((extra)=>{
    const ov=overlayRef.current;
    if(!ov||!baseW)return;
    const ctx=ov.getContext("2d");
    ctx.setTransform(1,0,0,1,0,0);
    ctx.clearRect(0,0,ov.width,ov.height);
    if(!showMarkup)return;
    const k=ov.width/baseW;                 // user-space → device px
    ctx.setTransform(k,0,0,k,0,0);
    ctx.lineCap="round";ctx.lineJoin="round";
    const all=extra?[...shapes,extra]:shapes;
    for(const s of all){
      ctx.strokeStyle=s.color||"#EF4444";
      ctx.fillStyle=s.color||"#EF4444";
      ctx.lineWidth=s.w||2;
      const p=s.points||[];
      if(s.type==="pen"&&p.length>1){
        ctx.beginPath();ctx.moveTo(p[0][0],p[0][1]);
        for(let i=1;i<p.length;i++)ctx.lineTo(p[i][0],p[i][1]);
        ctx.stroke();
      }else if(s.type==="box"&&p.length>1){
        ctx.strokeRect(p[0][0],p[0][1],p[1][0]-p[0][0],p[1][1]-p[0][1]);
      }else if(s.type==="arrow"&&p.length>1){
        const[a,b]=[p[0],p[1]];
        ctx.beginPath();ctx.moveTo(a[0],a[1]);ctx.lineTo(b[0],b[1]);ctx.stroke();
        const ang=Math.atan2(b[1]-a[1],b[0]-a[0]),h=(s.w||2)*5;
        ctx.beginPath();ctx.moveTo(b[0],b[1]);
        ctx.lineTo(b[0]-h*Math.cos(ang-0.4),b[1]-h*Math.sin(ang-0.4));
        ctx.lineTo(b[0]-h*Math.cos(ang+0.4),b[1]-h*Math.sin(ang+0.4));
        ctx.closePath();ctx.fill();
      }else if(s.type==="text"&&p.length){
        ctx.font=`${s.size||14}px Arial,sans-serif`;
        ctx.fillText(s.text||"",p[0][0],p[0][1]);
      }
    }
  },[shapes,showMarkup,baseW]);

  useEffect(()=>{paint();},[paint,rendering,scale,page]);

  /* ── coordinate mapping: screen → PDF user space ── */
  function toUser(clientX,clientY){
    const ov=overlayRef.current;
    if(!ov||!baseW)return[0,0];
    const r=ov.getBoundingClientRect();
    return[(clientX-r.left)/r.width*baseW,(clientY-r.top)/r.width*baseW];
  }
  // stroke weight stays visually consistent regardless of zoom
  const strokeW=()=>Math.max(1.2,2/(fitRef.current*scale));

  /* ── markup input ── */
  function startDraw(cx,cy){
    const pt=toUser(cx,cy);
    if(tool==="erase"){ eraseAt(pt); return; }
    if(tool==="text"){
      const t=window.prompt("Text to place:");
      if(t&&t.trim()){
        const s={type:"text",color,text:t.trim(),size:Math.max(10,16/(fitRef.current*scale)),points:[pt],w:strokeW()};
        setShapes(v=>[...v,s]);setDirty(true);
      }
      return;
    }
    draftRef.current={type:tool,color,w:strokeW(),points:[pt,pt]};
    if(tool==="pen")draftRef.current.points=[pt];
  }
  function moveDraw(cx,cy){
    const d=draftRef.current;
    if(!d)return;
    const pt=toUser(cx,cy);
    if(d.type==="pen")d.points.push(pt); else d.points[1]=pt;
    paint(d);
  }
  function endDraw(){
    const d=draftRef.current;
    draftRef.current=null;
    if(!d)return;
    if(d.type==="pen"&&d.points.length<2)return;
    setShapes(v=>[...v,d]);setDirty(true);
  }
  function eraseAt(pt){
    let best=-1,bestD=Infinity;
    const tol=Math.max(6,10/(fitRef.current*scale));
    shapes.forEach((s,i)=>{
      for(const p of s.points||[]){
        const dd=Math.hypot(p[0]-pt[0],p[1]-pt[1]);
        if(dd<bestD){bestD=dd;best=i;}
      }
    });
    if(best>=0&&bestD<tol*3){ setShapes(v=>v.filter((_,i)=>i!==best)); setDirty(true); }
  }
  const undo=()=>{setShapes(v=>v.slice(0,-1));setDirty(true);};

  async function saveMarkup(){
    setSaving(true);
    try{
      if(markupId) await API.markups.update(markupId,{shapes,author:user.name,updated_at:new Date().toISOString()});
      else{
        const r=await API.markups.create({drawing_id:drawing.id,page,shapes,author:user.name});
        const row=Array.isArray(r)?r[0]:r;
        if(row&&row.id)setMarkupId(row.id);
      }
      setDirty(false);
    }catch(e){ onErr&&onErr("Markup save failed: "+e.message); }
    setSaving(false);
  }

  /* ── gestures: pan/zoom always on two fingers; one finger draws in a tool mode ── */
  const drawMode=tool!=="pan";
  function onTouchStart(e){
    if(e.touches.length===2){
      gesture.current={mode:"zoom",startDist:dist(e.touches[0],e.touches[1]),startScale:scale};
      draftRef.current=null;
    }else if(e.touches.length===1){
      if(drawMode){ gesture.current={mode:"draw"}; startDraw(e.touches[0].clientX,e.touches[0].clientY); }
      else gesture.current={mode:"pan",startX:e.touches[0].clientX,startY:e.touches[0].clientY,startOff:{...offset}};
    }
  }
  function onTouchMove(e){
    const g=gesture.current;
    if(g.mode==="zoom"&&e.touches.length===2){
      e.preventDefault();
      setScale(clamp(g.startScale*(dist(e.touches[0],e.touches[1])/g.startDist),0.5,8));
    }else if(g.mode==="draw"&&e.touches.length===1){
      e.preventDefault();moveDraw(e.touches[0].clientX,e.touches[0].clientY);
    }else if(g.mode==="pan"&&e.touches.length===1){
      e.preventDefault();
      setOffset({x:g.startOff.x+(e.touches[0].clientX-g.startX),y:g.startOff.y+(e.touches[0].clientY-g.startY)});
    }
  }
  function onTouchEnd(){ if(gesture.current.mode==="draw")endDraw(); gesture.current={mode:null}; }

  function onMouseDown(e){
    if(drawMode){ gesture.current={mode:"draw"}; startDraw(e.clientX,e.clientY); }
    else gesture.current={mode:"pan",startX:e.clientX,startY:e.clientY,startOff:{...offset}};
  }
  function onMouseMove(e){
    const g=gesture.current;
    if(g.mode==="draw")moveDraw(e.clientX,e.clientY);
    else if(g.mode==="pan")setOffset({x:g.startOff.x+(e.clientX-g.startX),y:g.startOff.y+(e.clientY-g.startY)});
  }
  function onMouseUp(){ if(gesture.current.mode==="draw")endDraw(); gesture.current={mode:null}; }
  function onWheel(e){ if(!e.ctrlKey&&!e.metaKey)return; e.preventDefault(); setScale(s=>clamp(s*(e.deltaY<0?1.1:0.9),0.5,8)); }

  async function changePage(n){ if(dirty)await saveMarkup(); setPage(n); }
  async function goBack(){ if(dirty)await saveMarkup(); onBack(); }

  const btn={...primBtn,width:"auto",flex:"0 0 auto",padding:"6px 10px",fontSize:12,fontWeight:700,borderRadius:8,background:T.surface,border:`1px solid ${T.border}`,color:T.text,lineHeight:1.4,minWidth:0,whiteSpace:"nowrap"};
  const pill={background:"rgba(20,20,24,0.92)",border:`1px solid ${T.border}`,borderRadius:20,backdropFilter:"blur(6px)"};
  const tBtn=(id,label)=>(
    <button key={id} onClick={()=>setTool(id)} title={id}
      style={{...btn,background:tool===id?T.orange:T.surface,color:tool===id?"#000":T.text,padding:"6px 10px",fontSize:13}}>{label}</button>
  );

  return(
    <div style={{position:"fixed",inset:0,background:T.bg,zIndex:150,display:"flex",flexDirection:"column",fontFamily:"inherit"}}>
      <div style={{background:T.surface,borderBottom:`1px solid ${T.border}`,padding:"10px 14px",flexShrink:0,display:"flex",alignItems:"center",gap:10}}>
        <div style={{flex:1,minWidth:0}}>
          <button onClick={goBack} style={{background:"none",border:"none",color:T.sub,fontSize:13,cursor:"pointer",fontFamily:"inherit",marginBottom:4,padding:0}}>← Back to Drawings</button>
          <div style={{fontSize:14,fontWeight:800,color:T.text,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>
            {drawing.sheet_number?`${drawing.sheet_number} · `:""}{drawing.title}
          </div>
        </div>
        {dirty&&<span style={{fontSize:10,color:T.orange,fontWeight:700,flexShrink:0}}>● unsaved</span>}
        {saving&&<span style={{fontSize:10,color:T.muted,flexShrink:0}}>saving…</span>}
      </div>

      <div ref={wrapRef}
        onTouchStart={onTouchStart} onTouchMove={onTouchMove} onTouchEnd={onTouchEnd}
        onMouseDown={onMouseDown} onMouseMove={onMouseMove} onMouseUp={onMouseUp} onMouseLeave={onMouseUp}
        onWheel={onWheel}
        style={{flex:1,overflow:"hidden",position:"relative",background:"#3a3a42",touchAction:"none",cursor:drawMode?"crosshair":"grab"}}>
        {loading&&<div style={{position:"absolute",inset:0,display:"flex",alignItems:"center",justifyContent:"center",color:T.sub,fontSize:13}}>Loading drawing…</div>}
        {err&&<div style={{position:"absolute",inset:0,display:"flex",alignItems:"center",justifyContent:"center",padding:20}}>
          <div style={{background:T.redLow,border:`1px solid ${T.red}40`,borderRadius:12,padding:16,color:T.red,fontSize:13,maxWidth:420,textAlign:"center"}}>{err}</div>
        </div>}
        <div style={{position:"absolute",left:"50%",top:0,transform:`translateX(-50%) translate(${offset.x}px, ${offset.y}px)`,padding:"4px 0"}}>
          <div style={{position:"relative",display:"block"}}>
            <canvas ref={canvasRef} style={{display:"block",background:"#fff",boxShadow:"0 4px 24px rgba(0,0,0,0.5)"}}/>
            <canvas ref={overlayRef} style={{position:"absolute",left:0,top:0,pointerEvents:"none"}}/>
          </div>
        </div>
        {rendering&&!loading&&<div style={{position:"absolute",top:8,right:8,background:"rgba(0,0,0,0.6)",color:"#fff",fontSize:10,padding:"4px 8px",borderRadius:6}}>rendering…</div>}

        {/* Tool row */}
        {toolsOpen&&barOpen&&(
          <div style={{...pill,position:"absolute",left:"50%",bottom:56,transform:"translateX(-50%)",
            padding:"5px 8px",display:"flex",alignItems:"center",gap:4,zIndex:21,maxWidth:"96%",flexWrap:"wrap",justifyContent:"center"}}>
            {tBtn("pan","✋")}{tBtn("pen","✏️")}{tBtn("arrow","↗")}{tBtn("box","▭")}{tBtn("text","T")}{tBtn("erase","⌫")}
            <div style={{width:1,height:18,background:T.border}}/>
            {COLORS.map(c=>(
              <button key={c} onClick={()=>setColor(c)}
                style={{width:18,height:18,borderRadius:"50%",background:c,cursor:"pointer",flexShrink:0,
                  border:color===c?"2px solid #fff":`1px solid ${T.border}`,padding:0}}/>
            ))}
            <div style={{width:1,height:18,background:T.border}}/>
            <button onClick={undo} disabled={!shapes.length} style={{...btn,opacity:shapes.length?1:0.35}}>↶</button>
            <button onClick={()=>setShowMarkup(v=>!v)} style={{...btn,background:showMarkup?T.surface:T.border}}>{showMarkup?"👁":"🚫"}</button>
            <button onClick={saveMarkup} disabled={!dirty||saving}
              style={{...btn,background:dirty?T.green:T.surface,color:dirty?"#000":T.muted,opacity:saving?0.6:1}}>
              {saving?"…":"Save"}
            </button>
          </div>
        )}

        {/* Main bar */}
        {barOpen?(
          <div style={{...pill,position:"absolute",left:"50%",bottom:12,transform:"translateX(-50%)",
            padding:"6px 9px",display:"flex",alignItems:"center",gap:6,zIndex:20,maxWidth:"96%",flexWrap:"wrap",justifyContent:"center"}}>
            <button onClick={()=>changePage(Math.max(1,page-1))} disabled={page<=1} style={{...btn,opacity:page<=1?0.35:1}}>‹</button>
            <span style={{fontSize:12,color:T.text,fontWeight:700,minWidth:54,textAlign:"center",whiteSpace:"nowrap"}}>{page}/{pages||"…"}</span>
            <button onClick={()=>changePage(Math.min(pages,page+1))} disabled={page>=pages} style={{...btn,opacity:page>=pages?0.35:1}}>›</button>
            <div style={{width:1,height:18,background:T.border,flexShrink:0}}/>
            <button onClick={()=>setScale(s=>clamp(s*0.8,0.5,8))} style={btn}>−</button>
            <span style={{fontSize:12,color:T.text,fontWeight:700,minWidth:44,textAlign:"center",whiteSpace:"nowrap"}}>{Math.round(scale*100)}%</span>
            <button onClick={()=>setScale(s=>clamp(s*1.25,0.5,8))} style={btn}>+</button>
            <button onClick={()=>{setScale(1);setOffset({x:0,y:0});}} style={btn}>Fit</button>
            <div style={{width:1,height:18,background:T.border,flexShrink:0}}/>
            <button onClick={()=>setToolsOpen(v=>!v)}
              style={{...btn,background:toolsOpen?T.orange:T.surface,color:toolsOpen?"#000":T.text}}>✏️</button>
            <button onClick={()=>setBarOpen(false)} title="Hide controls"
              style={{...btn,background:"none",border:"none",color:T.muted,padding:"4px 6px",fontSize:13}}>⌄</button>
          </div>
        ):(
          <button onClick={()=>setBarOpen(true)} title="Show controls"
            style={{...pill,position:"absolute",left:"50%",bottom:12,transform:"translateX(-50%)",
              padding:"5px 12px",display:"flex",alignItems:"center",gap:7,zIndex:20,cursor:"pointer",
              color:T.sub,fontSize:11,fontWeight:700,fontFamily:"inherit"}}>
            <span style={{whiteSpace:"nowrap"}}>{page}/{pages||"…"} · {Math.round(scale*100)}%</span>
            <span style={{color:T.muted,fontSize:13}}>⌃</span>
          </button>
        )}
      </div>
    </div>
  );
}

// PDF.js loaded from CDN at runtime so package.json and the Vite build stay untouched.
let _pdfjsPromise=null;
function loadPdfJs(){
  if(window.pdfjsLib)return Promise.resolve(window.pdfjsLib);
  if(_pdfjsPromise)return _pdfjsPromise;
  _pdfjsPromise=new Promise((resolve,reject)=>{
    const s=document.createElement("script");
    s.src="https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js";
    s.onload=()=>{
      const lib=window.pdfjsLib;
      if(!lib){reject(new Error("PDF.js failed to load"));return;}
      lib.GlobalWorkerOptions.workerSrc="https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js";
      resolve(lib);
    };
    s.onerror=()=>reject(new Error("Could not load PDF.js from CDN"));
    document.head.appendChild(s);
  });
  return _pdfjsPromise;
}

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

function InfoTab({project,user,onEdit,onArchive,onDelete}){
  return(<div>
    <div style={cardS}>{[["Division",project.division],["Client",project.client],["Location",project.location],["AFE No.",project.afe],["Work Order",project.work_order],["Start Date",fmtDate(project.start_date)],["Status",project.status],["Created By",project.created_by]].map(([l,v])=>v?(<div key={l} style={{display:"flex",justifyContent:"space-between",padding:"11px 0",borderBottom:`1px solid ${T.border}`}}><span style={{fontSize:13,color:T.muted}}>{l}</span><span style={{fontSize:13,fontWeight:600}}>{v}</span></div>):null)}</div>
    {project.notes&&<div style={{...cardS,marginTop:12}}><div style={{fontSize:11,color:T.muted,textTransform:"uppercase",letterSpacing:"1px",marginBottom:8}}>Notes</div><div style={{fontSize:14,color:T.sub,lineHeight:1.6}}>{project.notes}</div></div>}
    {can(user,"edit_job")&&<div style={{marginTop:16,display:"flex",flexDirection:"column",gap:10}}><button onClick={onEdit} style={{...ghostBtn,width:"100%",textAlign:"center"}}>✏️ Edit Job</button><button onClick={onArchive} style={{...ghostBtn,width:"100%",textAlign:"center",color:T.muted}}>{project.status==="active"?"📦 Archive Job":"♻️ Restore Job"}</button><button onClick={onDelete} style={{...dangerBtn}}>🗑 Delete Job Permanently</button></div>}
  </div>);
}

const PTABS=[
  {id:"reports",icon:"📋",label:"Reports",perm:"submit_report"},
  {id:"tm",icon:"🧾",label:"T&M",perm:"submit_report"},
  {id:"time",icon:"⏱️",label:"Time",perm:"approve_report"},
  {id:"crew",icon:"🚜",label:"Crew",perm:"crew_equip"},
  {id:"subs",icon:"🏢",label:"Subs",perm:"subs"},
  {id:"safety",icon:"⛑️",label:"Safety",perm:"safety"},
  {id:"docs",icon:"📁",label:"Docs",perm:"docs"},
  {id:"drawings",icon:"📐",label:"Drawings",perm:"docs"},
  {id:"package",icon:"📦",label:"Sign Pkg",perm:"docs"},
  {id:"schedule",icon:"📅",label:"Schedule",perm:"schedule"},
  {id:"photos",icon:"📷",label:"Photos",perm:"photos"},
  {id:"weather",icon:"🌤️",label:"Weather",perm:"weather"},
  {id:"co",icon:"📋",label:"CO",perm:"subs"},
  {id:"rfi",icon:"📝",label:"RFI",perm:"subs"},
  {id:"info",icon:"ℹ️",label:"Info",perm:null},
];

function PMDashboard({onBack,user,projects:initProjects,onRefresh,onErr}){
  const [projects,setProjects]=useState(initProjects||[]);
  const [reports,setReports]=useState([]);
  const [pending,setPending]=useState([]);
  const [loading,setLoading]=useState(true);
  const [err,setErr]=useState("");
  const [pmTab,setPmTab]=useState("overview");
  const [activeReport,setActiveReport]=useState(null);
  const [activeProject,setActiveProject]=useState(null);
  const [unread,setUnread]=useState(0);
  const [showNotifs,setShowNotifs]=useState(false);

  async function load(){
    setLoading(true);setErr("");
    try{
      const[projs,reps,pend]=await Promise.all([API.projects.list(),API.reports.all(),API.reports.pending()]);
      setProjects(projs||[]);setReports(reps||[]);setPending(pend||[]);
    }catch(e){setErr(e.message);}
    setLoading(false);
  }
  useEffect(()=>{load();},[]);

  async function approve(id){try{await API.reports.update(id,{status:"approved",approved_by:user.name,approved_at:new Date().toISOString()});await load();}catch(e){setErr(e.message);}}
  async function flag(id,notes){try{await API.reports.update(id,{status:"flagged",pm_notes:notes});await load();}catch(e){setErr(e.message);}}

  const fmt=n=>"$"+Number(n||0).toLocaleString("en-US",{minimumFractionDigits:0});

  if(showNotifs)return(<div style={{background:T.bg,minHeight:"100vh",fontFamily:"inherit"}}><TopBar title="🔔 Notifications" onBack={()=>{setShowNotifs(false);load();}}/><NotificationsPanel onClose={()=>{setShowNotifs(false);load();}}/></div>);

  if(activeReport&&activeProject)return(<ReportDetail report={activeReport} project={activeProject} user={user} onBack={()=>{setActiveReport(null);setActiveProject(null);load();}} onDelete={async(id)=>{await API.reports.remove(id);setActiveReport(null);setActiveProject(null);load();}} onApprove={approve} onFlag={flag}/>);

  const DMTABS=[{id:"overview",l:"📊 Overview"},{id:"approvals",l:`✅ Approvals${pending.length>0?" ("+pending.length+")":""}`},{id:"workers",l:"👷 Workers"},{id:"billing",l:"💰 Billing"},{id:"reports",l:"📄 Reports"},{id:"users",l:"👤 Users"}];

  const allTot=reports.reduce((s,r)=>{const t=reportTotals(r);return{l:s.l+t.labor,e:s.e+t.equip,g:s.g+t.grand};},{l:0,e:0,g:0});
  const projMap={};projects.forEach(p=>{projMap[p.id]={...p,grand:0,count:0};});
  reports.forEach(r=>{if(!projMap[r.project_id])return;const t=reportTotals(r);projMap[r.project_id].grand+=t.grand;projMap[r.project_id].count++;});
  const projRows=Object.values(projMap).filter(p=>p.status==="active").sort((a,b)=>b.grand-a.grand);
  const workerHours={};
  reports.forEach(r=>(r.labor||[]).forEach(l=>{if(!l.name)return;if(!workerHours[l.name])workerHours[l.name]={name:l.name,reg:0,ot:0,pay:0};workerHours[l.name].reg+=parseFloat(l.regHrs)||0;workerHours[l.name].ot+=parseFloat(l.otHrs)||0;workerHours[l.name].pay+=laborAmt(l);}));
  const workerRows=Object.values(workerHours).sort((a,b)=>b.pay-a.pay);

  return(
    <div style={{background:T.bg,minHeight:"100vh",fontFamily:"inherit"}}>
      <div style={{background:T.surface,borderBottom:`1px solid ${T.border}`,padding:"14px 16px",position:"sticky",top:0,zIndex:50,display:"flex",justifyContent:"space-between",alignItems:"center"}}>
        <button onClick={onBack} style={{background:"none",border:"none",color:T.sub,fontSize:13,cursor:"pointer",fontFamily:"inherit"}}>← Back</button>
        <div style={{fontSize:16,fontWeight:900,color:T.orange}}>PM Dashboard</div>
        <button onClick={()=>setShowNotifs(true)} style={{background:"none",border:"none",color:T.muted,fontSize:20,cursor:"pointer"}}>{unread>0?"🔔":"🔕"}</button>
      </div>
      <ErrBanner msg={err} onDismiss={()=>setErr("")}/>

      {/* Tab bar */}
      <div style={{display:"flex",overflowX:"auto",borderBottom:`1px solid ${T.border}`,background:T.surface,WebkitOverflowScrolling:"touch"}}>
        {DMTABS.map(t=><button key={t.id} onClick={()=>setPmTab(t.id)}
          style={{flexShrink:0,padding:"12px 14px",background:"none",border:"none",borderBottom:`2px solid ${pmTab===t.id?T.orange:"transparent"}`,color:pmTab===t.id?T.orange:T.muted,fontSize:12,fontWeight:700,cursor:"pointer",fontFamily:"inherit",whiteSpace:"nowrap"}}>
          {t.l}
        </button>)}
      </div>

      <div style={{padding:"14px 16px 80px"}}>
        {loading&&<Spinner/>}

        {/* OVERVIEW */}
        {pmTab==="overview"&&!loading&&<div>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:14}}>
            {[[`${reports.length}`,"Total Reports",T.blue],[`${pending.length}`,"Pending Approval",T.yellow],[fmt(allTot.l),"Total Labor",T.green],[fmt(allTot.g),"Total Billed",T.orange]].map(([v,l,c])=>(
              <div key={l} style={{...cardS,textAlign:"center"}}><div style={{fontSize:22,fontWeight:900,color:c}}>{v}</div><div style={{fontSize:11,color:T.muted,textTransform:"uppercase",letterSpacing:"1px",marginTop:2}}>{l}</div></div>
            ))}
          </div>
          <div style={{fontSize:12,fontWeight:700,color:T.muted,textTransform:"uppercase",letterSpacing:"1px",marginBottom:8}}>Active Jobs by Billings</div>
          {projRows.slice(0,5).map(p=><div key={p.id} style={{...cardS,marginBottom:8,display:"flex",justifyContent:"space-between",alignItems:"center"}}>
            <div><div style={{fontSize:13,fontWeight:700,color:T.orange}}>{p.name}</div><div style={{fontSize:11,color:T.muted}}>{p.count} reports</div></div>
            <div style={{fontSize:15,fontWeight:800,color:T.green}}>{fmt(p.grand)}</div>
          </div>)}
        </div>}

        {/* APPROVALS */}
        {pmTab==="approvals"&&!loading&&<div>
          {pending.length===0&&<div style={{textAlign:"center",padding:"40px 0",color:T.muted}}><div style={{fontSize:32}}>✅</div><div style={{marginTop:8}}>All reports approved</div></div>}
          {pending.map(r=>{
            const proj=projects.find(p=>p.id===r.project_id)||{name:"Unknown"};
            const tot=reportTotals(r);
            return(<div key={r.id} style={{...cardS,marginBottom:10}}>
              <div style={{fontSize:14,fontWeight:700,color:T.orange}}>{proj.name}</div>
              <div style={{fontSize:12,color:T.muted,marginBottom:8}}>{r.date} · {r.submitted_by} · {fmt(tot.grand)}</div>
              <div style={{display:"flex",gap:8}}>
                <button onClick={()=>{setActiveReport(r);setActiveProject(proj);}} style={{...ghostBtn,flex:1,textAlign:"center",fontSize:13}}>👁 View</button>
                <button onClick={()=>approve(r.id)} style={{...primBtn,flex:1,fontSize:13,borderRadius:10,background:T.green}}>✓ Approve</button>
              </div>
            </div>);
          })}
        </div>}

        {/* WORKERS */}
        {pmTab==="workers"&&!loading&&<div>
          {workerRows.length===0&&<div style={{textAlign:"center",padding:"40px 0",color:T.muted}}>No labor data yet</div>}
          {workerRows.map(w=><div key={w.name} style={{...cardS,marginBottom:8,display:"flex",justifyContent:"space-between",alignItems:"center"}}>
            <div><div style={{fontSize:13,fontWeight:700,color:T.orange}}>{w.name}</div><div style={{fontSize:11,color:T.muted}}>{w.reg.toFixed(1)}h reg · {w.ot.toFixed(1)}h OT</div></div>
            <div style={{fontSize:15,fontWeight:800,color:T.green}}>{fmt(w.pay)}</div>
          </div>)}
        </div>}

        {/* BILLING */}
        {pmTab==="billing"&&!loading&&<div>
          {projRows.map(p=><div key={p.id} style={{...cardS,marginBottom:8}}>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}>
              <div><div style={{fontSize:13,fontWeight:700,color:T.orange}}>{p.name}</div><div style={{fontSize:11,color:T.muted}}>{p.client||"No client"} · {p.count} reports</div></div>
              <div style={{fontSize:15,fontWeight:800,color:T.green}}>{fmt(p.grand)}</div>
            </div>
          </div>)}
        </div>}

        {/* REPORTS */}
        {pmTab==="reports"&&!loading&&<div>
          {reports.slice(0,50).map(r=>{
            const proj=projects.find(p=>p.id===r.project_id)||{name:"Unknown"};
            const tot=reportTotals(r);
            const statusColor={approved:T.green,flagged:T.red,submitted:T.yellow}[r.status]||T.muted;
            return(<div key={r.id} style={{...cardS,marginBottom:8,borderLeft:`3px solid ${statusColor}`,cursor:"pointer"}} onClick={()=>{setActiveReport(r);setActiveProject(proj);}}>
              <div style={{display:"flex",justifyContent:"space-between"}}>
                <div><div style={{fontSize:13,fontWeight:700,color:T.orange}}>{proj.name}</div><div style={{fontSize:11,color:T.muted}}>{r.date} · {r.submitted_by}</div></div>
                <div style={{textAlign:"right"}}><div style={{fontSize:13,fontWeight:800,color:T.green}}>{fmt(tot.grand)}</div><span style={pill(statusColor)}>{r.status}</span></div>
              </div>
            </div>);
          })}
        </div>}

        {/* USERS */}
        {pmTab==="users"&&<UserManagementScreen user={user} onBack={()=>setPmTab("overview")}/>}
      </div>
    </div>
  );
}

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

  const today_d=new Date();
  const d30=new Date(today_d);d30.setDate(d30.getDate()+30);
  const d60=new Date(today_d);d60.setDate(d60.getDate()+60);
  const d90=new Date(today_d);d90.setDate(d90.getDate()+90);
  const certAlerts=[];
  members.forEach(m=>{
    (m.certifications||[]).forEach(c=>{
      if(!c.expiry)return;
      const exp=new Date(c.expiry+"T12:00:00");
      const daysLeft=Math.ceil((exp-today_d)/(1000*60*60*24));
      if(daysLeft<=90){
        certAlerts.push({worker:m.name,cert:c.name,expiry:c.expiry,daysLeft,color:daysLeft<=0?T.red:daysLeft<=30?T.red:daysLeft<=60?T.yellow:T.orange});
      }
    });
  });
  certAlerts.sort((a,b)=>a.daysLeft-b.daysLeft);

  return(
    <div style={{background:T.bg,minHeight:"100vh",fontFamily:"inherit"}}>
      {}
      {certAlerts.length>0&&<div style={{background:T.redLow,borderBottom:`1px solid ${T.red}40`,padding:"10px 16px"}}>
        <div style={{fontSize:12,fontWeight:800,color:T.red,marginBottom:6}}>⚠️ {certAlerts.filter(a=>a.daysLeft<=0).length>0?"EXPIRED":"EXPIRING SOON"} — {certAlerts.length} Cert{certAlerts.length!==1?"s":""}</div>
        {certAlerts.slice(0,3).map((a,i)=>(
          <div key={i} style={{fontSize:11,color:a.color,marginBottom:2}}>
            {a.daysLeft<=0?"🔴 EXPIRED":a.daysLeft<=30?"🔴":a.daysLeft<=60?"🟡":"🟠"} {a.worker} · {a.cert} · {a.daysLeft<=0?"Expired: "+a.expiry:"Expires: "+a.expiry+" ("+a.daysLeft+"d)"}
          </div>
        ))}
        {certAlerts.length>3&&<div style={{fontSize:10,color:T.muted,marginTop:2}}>+{certAlerts.length-3} more — view Crew Directory for full list</div>}
      </div>}
      <div style={{background:T.surface,borderBottom:`1px solid ${T.border}`,padding:"14px 16px",position:"sticky",top:0,zIndex:50}}>
        <button onClick={onBack} style={{background:"none",border:"none",color:T.sub,fontSize:13,cursor:"pointer",marginBottom:8,padding:0,fontFamily:"inherit"}}>← Back</button>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}><div style={{fontSize:20,fontWeight:900,letterSpacing:"-0.5px"}}>👥 Crew Directory</div><button onClick={()=>{setF({...blank});setMode("new");}} style={{background:T.orange,color:"#0D0D0F",border:"none",borderRadius:10,padding:"8px 14px",fontSize:13,fontWeight:800,cursor:"pointer",fontFamily:"inherit"}}>+ Add</button></div>
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

function UserManagementScreen({onBack,currentUser,user}){
  // Both call sites pass `user`; this component was written expecting
  // `currentUser`. Accept either, and tolerate neither.
  const me=currentUser||user||{};
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

  const profileMap={};profiles.forEach(p=>profileMap[p.name]=p);
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

        {/* Every role needs a PIN — the login screen requires one regardless of
            permission level, so gating this field by role locked crew out. */}
        <div style={{marginBottom:14}}>
          <label style={lbl}>{(ROLE_META[f.role]||ROLE_META.crew).label} PIN (required)</label>
          <input type="text" inputMode="numeric" maxLength={6} placeholder="Set a PIN (numbers)"
            value={f.pin||""} onChange={e=>set("pin",e.target.value.replace(/[^0-9]/g,""))} style={inp}/>
          <div style={{fontSize:11,color:T.muted,marginTop:4}}>
            This person will need to enter this PIN to sign in.
          </div>
          {!f.pin&&<div style={{fontSize:11,color:T.yellow,marginTop:4}}>
            ⚠️ Without a PIN this person cannot log in.
          </div>}
        </div>

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
          <button onClick={()=>{setF({...blank});setMode("edit");}} style={{background:T.orange,color:"#0D0D0F",border:"none",borderRadius:10,padding:"8px 14px",fontSize:13,fontWeight:800,cursor:"pointer",fontFamily:"inherit"}}>+ Add</button>
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
                    {!p.pin&&<span style={pill(T.yellow)}>NO PIN</span>}
                  </div>
                </div>
                <div style={{display:"flex",gap:6}}>
                  <button onClick={()=>{setActive(p);setF({name:p.name,role:p.role,division:p.division||null,pin:p.pin||"",active:p.active});setMode("edit");}} style={{background:T.orangeLow,border:`1px solid ${T.orange}40`,borderRadius:8,padding:"6px 12px",color:T.orange,fontSize:12,fontWeight:700,cursor:"pointer",fontFamily:"inherit"}}>Edit</button>
                  {p.name!==me.name&&<button onClick={()=>remove(p.id)} style={{background:"none",border:"none",color:T.muted,cursor:"pointer",fontSize:16,padding:0}}>🗑</button>}
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

function printEmployeeTimecards(cards,from,to,selectedJobs,projects,preOpenedWin=null){
  const filtered=cards.filter(c=>{
    if(!c.date||c.date<from||c.date>to)return false;
    if(selectedJobs&&selectedJobs.length>0&&!selectedJobs.includes(c.project_id))return false;
    return true;
  });
  const byEmployee={};
  filtered.forEach(c=>{
    const name=c.worker_name||"Unknown";
    if(!byEmployee[name])byEmployee[name]={name,entries:[],reg:0,ot:0,travel:0,total:0};
    const reg=parseFloat(c.reg_hours)||0;
    const ot=parseFloat(c.ot_hours)||0;
    const travel=parseFloat(c.travel_hours)||0;
    const total=c.total_hours?parseFloat(c.total_hours):reg+ot+travel;
    const proj=projects.find(p=>p.id===c.project_id);
    byEmployee[name].entries.push({...c,reg,ot,travel,total,projName:proj?.name||"General",projAfe:proj?.afe||""});
    byEmployee[name].reg+=reg;byEmployee[name].ot+=ot;byEmployee[name].travel+=travel;byEmployee[name].total+=total;
  });
  const employees=Object.values(byEmployee).sort((a,b)=>a.name.localeCompare(b.name));
  const fmtD=d=>{if(!d)return"";const[y,m,dy]=d.split("-");return`${m}/${dy}/${y}`;};
  const fmtN=n=>Number(n||0).toFixed(1);
  const html=`<!DOCTYPE html><html><head><meta charset="utf-8"><title>Employee Timecards ${fmtD(from)}–${fmtD(to)}</title>
<style>
@page{size:letter portrait;margin:0.5in;}
*{box-sizing:border-box;margin:0;padding:0;font-family:Arial,sans-serif;}
body{font-size:10pt;color:#000;}
.page-break{page-break-after:always;}
.header{display:flex;justify-content:space-between;align-items:center;padding-bottom:10px;border-bottom:2px solid #1f3864;margin-bottom:14px;}
.co{font-size:18pt;font-weight:900;color:#1f3864;}
.emp-header{background:#1f3864;color:#fff;border-radius:8px;padding:12px 16px;margin-bottom:12px;display:flex;justify-content:space-between;align-items:center;}
.emp-name{font-size:16pt;font-weight:900;}
.period{font-size:9pt;opacity:0.8;}
table{width:100%;border-collapse:collapse;margin-bottom:14px;}
th{background:#1f3864;color:#fff;padding:6px 8px;text-align:left;font-size:8pt;text-transform:uppercase;letter-spacing:0.5px;}
td{padding:5px 8px;border-bottom:1px solid #e5e7eb;font-size:9pt;}
tr:nth-child(even) td{background:#f9fafb;}
.totals-row td{background:#EEF2FF;font-weight:800;border-top:2px solid #1f3864;}
.summary{display:grid;grid-template-columns:1fr 1fr 1fr 1fr;gap:10px;margin-bottom:16px;}
.sum-box{border-radius:8px;padding:10px;text-align:center;}
.sum-box.reg{background:#dcfce7;border:1px solid #86efac;}
.sum-box.ot{background:#fef9c3;border:1px solid #fde047;}
.sum-box.travel{background:#dbeafe;border:1px solid #93c5fd;}
.sum-box.total{background:#1f3864;border:1px solid #1f3864;}
.sum-label{font-size:7pt;text-transform:uppercase;font-weight:700;color:#374151;}
.sum-box.total .sum-label{color:rgba(255,255,255,0.7);}
.sum-val{font-size:16pt;font-weight:900;color:#111;margin-top:2px;}
.sum-box.total .sum-val{color:#fff;font-size:20pt;}
.sigs{display:grid;grid-template-columns:1fr 1fr 1fr;gap:20px;margin-top:20px;}
.sig-box{border-top:1.5px solid #000;padding-top:8px;text-align:center;}
.sig-label{font-size:8pt;color:#555;text-transform:uppercase;}
@media print{body{-webkit-print-color-adjust:exact;print-color-adjust:exact;}}
</style></head><body>
${employees.length===0
  ? `<div style="text-align:center;padding:60px;color:#666"><h2>No entries found for ${fmtD(from)} – ${fmtD(to)}</h2></div>`
  : employees.map((emp,idx)=>`
<div>
  <div class="header">
    <div><div class="co">AIME</div><div style="font-size:8pt;color:#555">Atlantic Industrial Mechanical &amp; Environmental Inc.</div></div>
    <div style="text-align:right"><div style="font-size:14pt;font-weight:800;color:#1f3864">EMPLOYEE TIMECARD</div><div style="font-size:8pt;color:#555">${fmtD(from)} — ${fmtD(to)}</div></div>
  </div>
  <div class="emp-header">
    <div class="emp-name">${emp.name}</div>
    <div class="period">Report Period: ${fmtD(from)} — ${fmtD(to)}</div>
  </div>
  <table>
    <thead><tr><th>Date</th><th>Project</th><th>AFE/PO</th><th style="text-align:center">REG</th><th style="text-align:center">OT</th><th style="text-align:center">TRAVEL</th><th style="text-align:center">TOTAL</th><th>Notes</th></tr></thead>
    <tbody>
      ${emp.entries.sort((a,b)=>(a.date||"").localeCompare(b.date||"")).map(e=>`
      <tr><td>${fmtD(e.date)}</td><td>${e.projName}</td><td style="text-align:center">${e.projAfe||"—"}</td>
      <td style="text-align:center;color:#166534;font-weight:600">${fmtN(e.reg)}</td>
      <td style="text-align:center;color:#b45309;font-weight:600">${fmtN(e.ot)}</td>
      <td style="text-align:center;color:#1e40af;font-weight:600">${fmtN(e.travel)}</td>
      <td style="text-align:center;font-weight:800">${fmtN(e.total)}</td>
      <td style="font-size:8pt;color:#6b7280">${(e.notes||"").replace("Auto-filled from daily report","Auto")}</td></tr>`).join("")}
    </tbody>
    <tfoot><tr class="totals-row"><td colspan="3"><strong>TOTALS — ${emp.name.toUpperCase()}</strong></td>
    <td style="text-align:center">${fmtN(emp.reg)}</td><td style="text-align:center">${fmtN(emp.ot)}</td>
    <td style="text-align:center">${fmtN(emp.travel)}</td><td style="text-align:center"><strong>${fmtN(emp.total)}</strong></td><td></td></tr></tfoot>
  </table>
  <div class="summary">
    <div class="sum-box reg"><div class="sum-label">Regular</div><div class="sum-val">${fmtN(emp.reg)}h</div></div>
    <div class="sum-box ot"><div class="sum-label">Overtime</div><div class="sum-val">${fmtN(emp.ot)}h</div></div>
    <div class="sum-box travel"><div class="sum-label">Travel</div><div class="sum-val">${fmtN(emp.travel)}h</div></div>
    <div class="sum-box total"><div class="sum-label">TOTAL HOURS</div><div class="sum-val">${fmtN(emp.total)}</div></div>
  </div>
  <div class="sigs">
    <div class="sig-box"><div style="height:50px"></div><div class="sig-label">Employee Signature</div></div>
    <div class="sig-box"><div style="height:50px"></div><div class="sig-label">Supervisor Signature</div></div>
    <div class="sig-box"><div style="height:50px"></div><div class="sig-label">Date Approved</div></div>
  </div>
</div>${idx<employees.length-1?'<div class="page-break"></div>':""}`).join("")}
</body></html>`;
  const win=preOpenedWin||window.open("","_blank","width=950,height=800");
  if(!win){alert("Popup blocked — please allow popups for this site and try again.");return;}
  win.document.write(html);
  win.document.close();
  win.focus();
  setTimeout(()=>{win.focus();win.print();},1200);
}

function TimeCardsScreen({user,projects,onBack}){
  const [cards,setCards]=useState([]);const [loading,setLoading]=useState(true);
  const [err,setErr]=useState('');
  const [fromDate,setFromDate]=useState(()=>{const d=new Date();d.setDate(d.getDate()-14);return d.toISOString().slice(0,10);});
  const [toDate,setToDate]=useState(today());
  const [selectedJobs,setSelectedJobs]=useState([]);
  const [showJobFilter,setShowJobFilter]=useState(false);
  const [printing,setPrinting]=useState(false);

  useEffect(()=>{(async()=>{setLoading(true);try{const r=await API.timeCards.all();setCards(Array.isArray(r)?r:[]);}catch(e){setErr(e.message);}setLoading(false);})();},[]);

  async function remove(id){try{await API.timeCards.remove(id);setCards(c=>c.filter(x=>x.id!==id));}catch(e){setErr(e.message);}}

  function handlePrint(){
    const win=window.open("","_blank","width=900,height=750");
    if(!win){
      alert("Popup blocked! Please allow popups for this site in your browser settings, then try again.");
      return;
    }
    setPrinting(true);
    try{
      printEmployeeTimecards(cards,fromDate,toDate,selectedJobs.length>0?selectedJobs:null,projects,win);
    }catch(e){
      win.close();
      alert("Error generating report: "+e.message);
    }
    setTimeout(()=>setPrinting(false),1000);
  }

  const filtered=cards.filter(c=>c.date&&c.date>=fromDate&&c.date<=toDate&&(selectedJobs.length===0||selectedJobs.includes(c.project_id)));
  const byWorker={};
  filtered.forEach(c=>{
    const n=c.worker_name||'?';
    if(!byWorker[n])byWorker[n]={name:n,total:0,ot:0,reg:0,travel:0};
    const reg=parseFloat(c.reg_hours)||0;const ot=parseFloat(c.ot_hours)||0;const trav=parseFloat(c.travel_hours)||0;
    const tot=c.total_hours?parseFloat(c.total_hours):reg+ot+trav;
    byWorker[n].total+=tot;byWorker[n].ot+=ot;byWorker[n].reg+=reg;byWorker[n].travel+=trav;
  });
  const workerRows=Object.values(byWorker).sort((a,b)=>b.total-a.total);
  const fmt=n=>Number(n||0).toFixed(1);
  const totalHrs=workerRows.reduce((s,w)=>s+w.total,0);
  const totalOT=workerRows.reduce((s,w)=>s+w.ot,0);

  return(
    <div style={{background:T.bg,minHeight:'100vh',fontFamily:'inherit'}}>
      <TopBar title="⏱️ Time Cards" onBack={onBack}/>
      <div style={{padding:'12px 16px 100px'}}>
        <ErrBanner msg={err} onDismiss={()=>setErr('')}/>

        {/* Date Range Selector */}
        <div style={{...cardS,marginBottom:12}}>
          <div style={{fontSize:12,fontWeight:700,color:T.muted,textTransform:'uppercase',letterSpacing:'1px',marginBottom:10}}>📅 Date Range</div>
          <div style={{display:'grid',gridTemplateColumns:'1fr 1fr',gap:8,marginBottom:10}}>
            <div><label style={lbl}>From</label><input type="date" value={fromDate} onChange={e=>setFromDate(e.target.value)} style={inp}/></div>
            <div><label style={lbl}>To</label><input type="date" value={toDate} onChange={e=>setToDate(e.target.value)} style={inp}/></div>
          </div>
          <div style={{display:'flex',gap:6,flexWrap:'wrap'}}>
            {[['This Week',0,6],['Last 2 Weeks',0,13],['This Month',0,29],['Last Month',30,59]].map(([label,from,to])=>(
              <button key={label} onClick={()=>{
                const d=new Date();
                const t=new Date();t.setDate(t.getDate()-from);
                const f=new Date();f.setDate(f.getDate()-to);
                setToDate(t.toISOString().slice(0,10));
                setFromDate(f.toISOString().slice(0,10));
              }} style={{background:T.surface,border:`1px solid ${T.border}`,borderRadius:8,padding:'4px 10px',color:T.muted,fontSize:11,fontWeight:600,cursor:'pointer',fontFamily:'inherit'}}>
                {label}
              </button>
            ))}
          </div>
        </div>
        <div style={{...cardS,marginBottom:12}}>
          <button onClick={()=>setShowJobFilter(s=>!s)} style={{background:'none',border:'none',color:T.text,fontWeight:700,fontSize:13,cursor:'pointer',fontFamily:'inherit',width:'100%',textAlign:'left',display:'flex',justifyContent:'space-between'}}>
            <span>🏗️ Jobs {selectedJobs.length>0?`(${selectedJobs.length} selected)`:'(All)'}</span>
            <span style={{color:T.muted}}>{showJobFilter?'▲':'▼'}</span>
          </button>
          {showJobFilter&&<div style={{marginTop:10,display:'flex',flexDirection:'column',gap:6}}>
            {projects.filter(p=>p.status==='active').map(p=>(
              <label key={p.id} style={{display:'flex',alignItems:'center',gap:8,cursor:'pointer',fontSize:13,color:T.sub}}>
                <input type="checkbox" checked={selectedJobs.includes(p.id)}
                  onChange={e=>setSelectedJobs(s=>e.target.checked?[...s,p.id]:s.filter(x=>x!==p.id))}
                  style={{width:16,height:16,accentColor:T.orange}}/>
                {p.name}
                {p.division&&<span style={{fontSize:10,color:T.muted}}>· {p.division}</span>}
              </label>
            ))}
            {selectedJobs.length>0&&<button onClick={()=>setSelectedJobs([])} style={{...ghostBtn,fontSize:12,textAlign:'center'}}>Clear filter</button>}
          </div>}
        </div>
        {filtered.length>0&&<div style={{display:'grid',gridTemplateColumns:'1fr 1fr 1fr',gap:8,marginBottom:12}}>
          {[[totalHrs.toFixed(1),'Total Hrs',T.orange],[totalOT.toFixed(1),'OT Hrs',T.yellow],[workerRows.length,'Workers',T.blue]].map(([v,l,c])=>(
            <div key={l} style={{...cardS,textAlign:'center',padding:'10px 6px'}}><div style={{fontSize:20,fontWeight:900,color:c}}>{v}</div><div style={{fontSize:10,color:T.muted,textTransform:'uppercase',marginTop:2}}>{l}</div></div>
          ))}
        </div>}

        {/* Print button */}
        <button onClick={handlePrint} disabled={printing||filtered.length===0}
          style={{...primBtn,borderRadius:14,marginBottom:16,background:filtered.length===0?T.surface:T.green,color:filtered.length===0?T.muted:'#000',opacity:filtered.length===0?0.5:1}}>
          {printing?'Opening PDF…':`🖨️ Print Timecard Report (${filtered.length} entries · ${workerRows.length} workers)`}
        </button>

        <div style={{background:T.greenLow,border:`1px solid ${T.green}40`,borderRadius:12,padding:'10px 14px',marginBottom:14,fontSize:12,color:T.green,lineHeight:1.5}}>
          <strong>⚡ Auto-filled from daily reports</strong> — hours added automatically when a report is submitted.
        </div>

        {loading&&<Spinner/>}
        {!loading&&filtered.length===0&&<div style={{textAlign:'center',padding:'40px 0',color:T.muted}}><div style={{fontSize:32}}>⏱️</div><div style={{marginTop:8}}>No time cards in this date range</div></div>}

        {/* Worker summary cards */}
        {workerRows.map(w=>(
          <div key={w.name} style={{...cardS,marginBottom:8}}>
            <div style={{display:'flex',justifyContent:'space-between',alignItems:'center',marginBottom:6}}>
              <div style={{fontSize:14,fontWeight:800,color:T.orange}}>{w.name}</div>
              <div style={{fontSize:16,fontWeight:900,color:T.green}}>{fmt(w.total)}h</div>
            </div>
            <div style={{display:'flex',gap:12,fontSize:11,color:T.muted}}>
              <span>Reg: <strong style={{color:T.sub}}>{fmt(w.reg)}h</strong></span>
              {w.ot>0&&<span>OT: <strong style={{color:T.yellow}}>{fmt(w.ot)}h</strong></span>}
              {w.travel>0&&<span>Travel: <strong style={{color:T.blue}}>{fmt(w.travel)}h</strong></span>}
            </div>
          </div>
        ))}

        {/* Individual entries */}
        {filtered.length>0&&<>
          <div style={{fontSize:11,fontWeight:700,color:T.muted,textTransform:'uppercase',letterSpacing:'1px',margin:'14px 0 8px'}}>All Entries</div>
          {filtered.sort((a,b)=>b.date?.localeCompare(a.date)).map(c=>{
            const reg=parseFloat(c.reg_hours)||0;const ot=parseFloat(c.ot_hours)||0;const trav=parseFloat(c.travel_hours)||0;
            const tot=c.total_hours?parseFloat(c.total_hours):reg+ot+trav;
            const proj=projects.find(p=>p.id===c.project_id);
            return(<div key={c.id} style={{...cardS,marginBottom:6,display:'flex',justifyContent:'space-between',alignItems:'flex-start'}}>
              <div style={{flex:1}}>
                <div style={{fontSize:13,fontWeight:700,color:T.orange}}>{c.worker_name}</div>
                <div style={{fontSize:11,color:T.muted}}>{c.date}{proj?` · ${proj.name}`:''}</div>
                {c.notes&&<div style={{fontSize:10,color:T.muted,fontStyle:'italic',marginTop:1}}>{c.notes}</div>}
              </div>
              <div style={{display:'flex',alignItems:'center',gap:8}}>
                <div style={{textAlign:'right'}}>
                  <div style={{fontSize:15,fontWeight:800,color:T.green}}>{fmt(tot)}h</div>
                  {ot>0&&<div style={{fontSize:9,color:T.yellow}}>{fmt(ot)} OT</div>}
                </div>
                {(user.role==='admin'||user.role==='pm')&&<button onClick={()=>remove(c.id)} style={{background:'none',border:'none',color:T.red,cursor:'pointer',fontSize:14}}>🗑</button>}
              </div>
            </div>);
          })}
        </>}
      </div>
    </div>
  );
}

function EstimatingScreen({user,onBack}){
  return(
    <div style={{background:T.bg,minHeight:'100vh',fontFamily:'inherit'}}>
      <TopBar title="📊 Estimating" onBack={onBack}/>
      <div style={{display:'flex',flexDirection:'column',alignItems:'center',justifyContent:'center',padding:'60px 24px',textAlign:'center'}}>
        <div style={{fontSize:64,marginBottom:20}}>🚧</div>
        <div style={{fontSize:22,fontWeight:900,color:T.text,marginBottom:8}}>Under Maintenance</div>
        <div style={{fontSize:14,color:T.muted,lineHeight:1.7,maxWidth:320}}>The Estimating platform is currently being updated. Check back soon.</div>
      </div>
    </div>
  );
}

function EmployeeHistory({user,projects,onBack}){return <TimeCardsScreen user={user} projects={projects} onBack={onBack}/>;}
function FinancialsScreen({user,projects,onBack,onErr}){return(<div style={{background:T.bg,minHeight:'100vh'}}><TopBar title="💵 Financials" onBack={onBack}/><div style={{padding:20,color:T.muted,textAlign:'center',marginTop:40}}>Financials coming soon</div></div>);}
function TimecardReportScreen({user,projects,onBack}){return <TimeCardsScreen user={user} projects={projects} onBack={onBack}/>;}
function EmailSummaryModal({user,projects,onClose}){return(<div style={{position:'fixed',inset:0,background:'rgba(0,0,0,0.8)',display:'flex',alignItems:'center',justifyContent:'center',zIndex:200}}><div style={{background:T.card,borderRadius:16,padding:24,maxWidth:400,width:'90%'}}><div style={{fontSize:16,fontWeight:800,color:T.text,marginBottom:12}}>📧 Email Summary</div><button onClick={onClose} style={{...primBtn,borderRadius:12}}>Close</button></div></div>);}

function ProjectDetail({project:initP,user,onBack,onProjectUpdated,isOnline=true,onErr:onErrProp,onRefresh}){
  const [project,setProject]=useState(initP);
  const [reports,setReports]=useState([]);const [safety,setSafety]=useState([]);const [photos,setPhotos]=useState([]);const [weather,setWeather]=useState([]);
  const [tab,setTab]=useState("reports");
  const [showTMForm,setShowTMForm]=useState(false);
  const [editingTicket,setEditingTicket]=useState(null);const [loading,setLoading]=useState(true);const [err,setErr]=useState("");
  const [screen,setScreen]=useState("detail");const [activeReport,setActiveReport]=useState(null);const [editProject,setEditProject]=useState(false);
  const [projSaving,setProjSaving]=useState(false);
  const visibleTabs=PTABS.filter(t=>!t.perm||can(user,t.perm));
  const divMeta=DIV_META[project.division]||{color:T.orange,icon:"🏗️"};

  async function load(silent=false){if(!silent)setLoading(true);try{const[reps,saf,phs,wx]=await Promise.all([API.reports.forProject(project.id),API.safety.forProject(project.id),API.photos.forProject(project.id),API.weather.forProject(project.id)]);setReports(reps||[]);setSafety(saf||[]);setPhotos(phs||[]);setWeather(wx||[]);}catch(e){setErr(e.message);}if(!silent)setLoading(false);}
  useEffect(()=>{load();const firstTab=visibleTabs[0]?.id||"info";setTab(firstTab);},[project.id]);

  async function saveReport(d){try{const{rental_equipment,...dbData}=d;let saved=false;try{await API.reports.create({...dbData,rental_equipment,project_id:project.id});saved=true;}catch(colErr){if(colErr.message&&colErr.message.includes("rental_equipment")){await API.reports.create({...dbData,project_id:project.id});saved=true;}else{throw colErr;}}if(!saved)throw new Error("Save failed");await load(true);setScreen("detail");}catch(e){setErr(e.message);}}
  async function deleteReport(id){try{await API.reports.remove(id);setActiveReport(null);await load(true);setScreen("detail");}catch(e){setErr(e.message);}}
  async function approveReport(id){try{await API.reports.update(id,{status:"approved",approved_by:user.name,approved_at:new Date().toISOString()});setActiveReport(r=>({...r,status:"approved"}));await load(true);}catch(e){setErr(e.message);}}
  async function flagReport(id,pm_notes){try{await API.reports.update(id,{status:"flagged",pm_notes});setActiveReport(r=>({...r,status:"flagged",pm_notes}));await notify("report_flagged","Report Flagged",pm_notes,{project_id:project.id,report_id:id});await load(true);}catch(e){setErr(e.message);}}
  async function updateProject(data){
    setProjSaving(true);setErr("");
    try{
      // Projects carry computed fields (_reports, _billed, _openRfis, _photos …)
      // that have no database columns. Strip every underscore-prefixed key rather
      // than naming them one by one, so adding a new one can't break saving again.
      const dbData=Object.fromEntries(Object.entries(data).filter(([k])=>!k.startsWith("_")));
      const res=await API.projects.update(project.id,dbData);
      const u=Array.isArray(res)?res[0]:res;
      const merged=u||{...project,...dbData};
      setProject(merged);
      onProjectUpdated&&onProjectUpdated(merged);
      setEditProject(false);
    }catch(e){setErr(e.message||"Save failed");}
    setProjSaving(false);
  }
  async function archiveProject(){if(!window.confirm(project.status==="active"?"Archive this job?":"Restore?"))return;await updateProject({status:project.status==="active"?"archived":"active"});onBack();}
  async function deleteProject(){if(!window.confirm("Permanently delete this job and ALL its data? This cannot be undone."))return;if(!window.confirm("Are you sure? All reports, photos, time cards and safety logs will be deleted."))return;try{await API.projects.remove(project.id);onBack();}catch(e){setErr(e.message);}}

  const tot=reports.reduce((s,r)=>{const t=reportTotals(r);return{l:s.l+t.labor,e:s.e+t.equip,m:s.m+t.mats,g:s.g+t.grand};},{l:0,e:0,m:0,g:0});

  if(screen==="newReport"&&can(user,"submit_report")) return <DailyReportForm user={user} project={project} onSave={saveReport} onCancel={()=>setScreen("detail")} isOnline={isOnline}/>;
  if(screen==="reportDetail"&&activeReport) return <ReportDetail report={activeReport} project={project} user={user} onBack={()=>setScreen("detail")} onDelete={deleteReport} onApprove={approveReport} onFlag={flagReport}/>;
  if(editProject&&can(user,"edit_job")) return <ProjectForm initial={project} onSave={updateProject} onCancel={()=>{setEditProject(false);setErr("");}} defaultDivision={project.division} saving={projSaving} externalErr={err} onClearErr={()=>setErr("")}/>;

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
          {visibleTabs.map(t=>(<button key={t.id} onClick={()=>setTab(t.id)} style={{flexShrink:0,background:tab===t.id?divMeta.color:"transparent",border:tab===t.id?"none":`1px solid ${T.border}`,borderRadius:10,padding:"8px 10px",fontSize:11,fontWeight:tab===t.id?800:500,cursor:"pointer",color:tab===t.id?"#0D0D0F":T.sub,fontFamily:"inherit",whiteSpace:"nowrap"}}>{t.icon} {t.label}</button>))}
        </div>
      </div>
      <div style={{padding:"14px 16px 80px"}}>
        <ErrBanner msg={err} onDismiss={()=>setErr("")}/>
        {loading&&<Spinner/>}
        {!loading&&tab==="tm"&&(showTMForm||editingTicket?(
          <TMTicketForm project={project} user={user} ticket={editingTicket}
            onBack={()=>{setShowTMForm(false);setEditingTicket(null);}}
            onSaved={()=>{setShowTMForm(false);setEditingTicket(null);}}/>
        ):(
          <TMTicketList project={project} user={user}
            onNew={()=>{setEditingTicket(null);setShowTMForm(true);}}
            onOpen={(t)=>{setEditingTicket(t);setShowTMForm(true);}}/>
        ))}
        {!loading&&tab==="reports"&&(<div>
          {can(user,"submit_report")&&<button onClick={()=>setScreen("newReport")} style={{...primBtn,marginBottom:16,borderRadius:14,padding:"18px",fontSize:17,background:divMeta.color}}>📋 + New Daily Report</button>}
          {reports.length===0&&<div style={{textAlign:"center",padding:"32px 16px",color:T.muted}}>
              <div style={{fontSize:44,marginBottom:12}}>📋</div>
              <div style={{fontSize:16,fontWeight:700,color:T.sub,marginBottom:6}}>No Daily Reports Yet</div>
              <div style={{fontSize:13,color:T.muted,lineHeight:1.6,marginBottom:16}}>Tap <strong style={{color:T.orange}}>+ New Report</strong> below to submit the first daily report for this job.</div>
              <div style={{background:T.orangeLow,border:`1px solid ${T.orange}30`,borderRadius:12,padding:"10px 14px",fontSize:12,color:T.orange,textAlign:"left"}}>
                💡 Reports track labor, equipment, materials and site conditions — and automatically generate time cards for payroll.
              </div>
            </div>}
          {reports.map(r=>{const t=reportTotals(r);const sc={submitted:T.yellow,approved:T.green,flagged:T.red}[r.status||"submitted"]||T.muted;return(<div key={r.id} onClick={()=>{setActiveReport(r);setScreen("reportDetail");}} style={{...cardS,marginBottom:9,cursor:"pointer",borderLeft:`3px solid ${sc}`,display:"flex",justifyContent:"space-between",alignItems:"center"}}><div><div style={{display:"flex",alignItems:"center",gap:8}}><div style={{fontSize:15,fontWeight:700}}>{fmtShort(r.date)}</div><span style={pill(sc)}>{(r.status||"submitted").toUpperCase()}</span></div><div style={{fontSize:11,color:T.muted,marginTop:4,display:"flex",gap:8}}>{(r.labor||[]).length>0&&<span>👷 {r.labor.length}</span>}{(r.equipment||[]).length>0&&<span>🚜 {r.equipment.length}</span>}{r.submitted_by&&<span>by {r.submitted_by}</span>}</div></div><div style={{textAlign:"right"}}><div style={{fontSize:17,fontWeight:900,color:T.green}}>${fmt(t.grand)}</div><div style={{fontSize:9,color:T.muted}}>TOTAL</div></div></div>);})}
        </div>)}
        {!loading&&tab==="time"     &&can(user,"time_card")   &&<TimeCardsTab projectId={project.id} user={user} onErr={setErr}/>}
        {!loading&&tab==="crew"     &&can(user,"crew_equip")  &&<CrewEquipTab projectId={project.id} user={user} onErr={setErr}/>}
        {!loading&&tab==="subs"     &&can(user,"subs")        &&<SubsTab projectId={project.id} user={user} onErr={setErr}/>}
        {!loading&&tab==="safety"   &&can(user,"safety")      &&<SafetyTab projectId={project.id} safety={safety} user={user} onRefresh={()=>load(true)} onErr={setErr}/>}
        {!loading&&tab==="co"&&<ChangeOrdersTab project={project} user={user} onErr={setErr}/>}
        {!loading&&tab==="rfi"&&<RFIsTab project={project} user={user} onErr={setErr}/>}
        {!loading&&tab==="docs"     &&can(user,"docs")        &&<DocsTab projectId={project.id} user={user} onErr={setErr}/>}
        {!loading&&tab==="drawings" &&can(user,"docs")        &&<DrawingsTab projectId={project.id} user={user} onErr={setErr}/>}
        {!loading&&tab==="package"  &&can(user,"docs")        &&<SignaturePackageScreen project={project} user={user} onBack={()=>setTab("reports")} onErr={setErr}/>}
        {!loading&&tab==="schedule" &&can(user,"schedule")    &&<ScheduleTab projectId={project.id} user={user} onErr={setErr}/>}
        {!loading&&tab==="photos"   &&can(user,"photos")      &&<PhotosTab projectId={project.id} photos={photos} onRefresh={()=>load(true)} onErr={setErr}/>}
        {!loading&&tab==="weather"  &&can(user,"weather")     &&<WeatherTab projectId={project.id} project={project} weather={weather} onRefresh={()=>load(true)} onErr={setErr}/>}
        {!loading&&tab==="info"     &&<InfoTab project={project} user={user} onEdit={()=>setEditProject(true)} onArchive={archiveProject} onDelete={deleteProject}/>}
      </div>
    </div>
  );
}

function ChangeOrdersTab({project,user,onErr}){
  const canEdit=user.role==="admin"||user.role==="pm";
  const [cos,setCos]=useState([]);
  const [loading,setLoading]=useState(true);
  const [showForm,setShowForm]=useState(false);
  const [editing,setEditing]=useState(null);
  const [saving,setSaving]=useState(false);
  const [shareCo,setShareCo]=useState(null);
  const [coCopied,setCoCopied]=useState(false);
  const appUrl=window.location.origin+window.location.pathname;
  const blank={co_number:"",description:"",date_submitted:today(),amount:"",status:"Pending",submitted_by:user.name,approved_by:"",approved_date:"",notes:"",client_signature:null,client_signed_by:""};
  const [f,setF]=useState(blank);
  const set=(k,v)=>setF(x=>({...x,[k]:v}));
  const statusColor={Pending:T.yellow,Approved:T.green,Rejected:T.red};

  useEffect(()=>{load();},[project.id]);
  async function load(){setLoading(true);try{const r=await API.changeOrders.forProject(project.id);setCos(Array.isArray(r)?r:[]);}catch(e){onErr(e.message);}setLoading(false);}

  async function save(){
    setSaving(true);
    try{
      const toDate=v=>v&&v.trim()&&v!=="Invalid Date"?v:null;
      const payload={...f,project_id:project.id,amount:parseFloat(f.amount)||0,date_submitted:toDate(f.date_submitted),approved_date:toDate(f.approved_date)};
      if(editing){await API.changeOrders.update(editing,payload);}
      else{await API.changeOrders.create(payload);}
      setShowForm(false);setEditing(null);setF(blank);await load();
    }catch(e){onErr(e.message);}
    setSaving(false);
  }

  async function remove(id){if(!window.confirm("Delete this change order?"))return;try{await API.changeOrders.remove(id);await load();}catch(e){onErr(e.message);}}

  function copyCoLink(co){
    const link=`${appUrl}?co=${co.id}`;
    navigator.clipboard.writeText(link).then(()=>{setCoCopied(true);setTimeout(()=>setCoCopied(false),3000);}).catch(()=>{const el=document.createElement("textarea");el.value=link;document.body.appendChild(el);el.select();document.execCommand("copy");document.body.removeChild(el);setCoCopied(true);setTimeout(()=>setCoCopied(false),3000);});
  }

  function emailCO(co){
    const link=`${appUrl}?co=${co.id}`;
    const subj=`Change Order ${co.co_number} — ${project.name} — Signature Required`;
    const ln="%0D%0A";
    const body=[`Please review and sign the following Change Order:`,ln,ln,`━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━`,ln,`  📋  CHANGE ORDER — ${co.co_number}`,ln,`━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━`,ln,ln,`  Project: ${project.name}`,ln,`  Amount: $${Number(co.amount||0).toLocaleString("en-US",{minimumFractionDigits:2})}`,ln,ln,`  DESCRIPTION:`,ln,`  ${co.description||""}`,ln,ln,`  ➤  ${link}`,ln,ln,`(No login required — open in any browser, sign, and click Approve)`,ln,ln,`Thank you,`,ln,`${co.submitted_by||"AIME Field Operations"}`,ln,`Atlantic Industrial Mechanical & Environmental Inc.`,ln,].filter(Boolean).join("");
    if(navigator.clipboard) navigator.clipboard.writeText(link).catch(()=>{});
    window.location.href=`mailto:?subject=${encodeURIComponent(subj)}&body=${body}`;
  }

  function printCO(co){
    const contractVal=parseFloat(project.contract_value)||0;
    const approvedCOs=cos.filter(c=>c.status==="Approved").reduce((s,c)=>s+(parseFloat(c.amount)||0),0);
    const fmt=n=>"$"+Number(n||0).toLocaleString("en-US",{minimumFractionDigits:2});
    const html=`<!DOCTYPE html><html><head><meta charset="utf-8"><title>CO ${co.co_number}</title>
<style>@page{size:letter portrait;margin:0.6in;}*{box-sizing:border-box;margin:0;padding:0;font-family:Arial,sans-serif;}body{font-size:10pt;color:#000;}
.lh{display:flex;justify-content:space-between;align-items:flex-start;padding-bottom:14px;border-bottom:3px solid #1f3864;margin-bottom:20px;}
.co{font-size:20pt;font-weight:900;color:#1f3864;}.co-sub{font-size:9pt;color:#555;margin-top:4px;}
.doc-title{text-align:right;}.doc-title h1{font-size:20pt;font-weight:900;color:#1f3864;}
.badge{display:inline-block;padding:4px 14px;border-radius:20px;font-weight:700;font-size:10pt;margin-top:6px;background:${co.status==="Approved"?"#dcfce7":co.status==="Rejected"?"#fee2e2":"#fef9c3"};color:${co.status==="Approved"?"#166534":co.status==="Rejected"?"#991b1b":"#713f12"};}
.proj-box{background:#f0f4ff;border:1px solid #c7d2fe;border-radius:8px;padding:12px 16px;margin-bottom:20px;display:grid;grid-template-columns:1fr 1fr 1fr;gap:10px;}
.fl{font-size:7.5pt;font-weight:700;text-transform:uppercase;letter-spacing:0.5px;color:#6b7280;margin-bottom:2px;}.fv{font-size:10pt;font-weight:600;color:#111;}
.amt-box{background:#1f3864;color:#fff;border-radius:10px;padding:16px 20px;text-align:center;margin-bottom:18px;}
.desc-box{background:#f9fafb;border:1px solid #e5e7eb;border-radius:6px;padding:12px;font-size:10pt;line-height:1.7;min-height:60px;margin-bottom:18px;}
.sig-grid{display:grid;grid-template-columns:1fr 1fr;gap:20px;margin-top:24px;}
.sig-box{border-top:1.5px solid #000;padding-top:8px;}.sig-label{font-size:8pt;color:#666;text-transform:uppercase;}
.footer{margin-top:24px;padding-top:10px;border-top:1px solid #e5e7eb;font-size:7.5pt;color:#9ca3af;display:flex;justify-content:space-between;}
@media print{body{-webkit-print-color-adjust:exact;print-color-adjust:exact;}}</style></head><body>
<div class="lh"><div><div class="co">AIME</div><div class="co-sub">Atlantic Industrial Mechanical & Environmental Inc.<br>Field Operations Division</div></div>
<div class="doc-title"><h1>Change Order</h1><div>${co.co_number}</div><div><span class="badge">${co.status.toUpperCase()}</span></div></div></div>
<div class="proj-box">
<div><div class="fl">Project</div><div class="fv">${project.name||"—"}</div></div>
<div><div class="fl">Client</div><div class="fv">${project.client||"—"}</div></div>
<div><div class="fl">Division</div><div class="fv">${project.division||"—"}</div></div>
<div><div class="fl">Date</div><div class="fv">${co.date_submitted||"—"}</div></div>
<div><div class="fl">Submitted By</div><div class="fv">${co.submitted_by||"—"}</div></div>
<div><div class="fl">AFE / PO</div><div class="fv">${project.afe||"—"}</div></div></div>
<div class="amt-box"><div style="font-size:9pt;text-transform:uppercase;letter-spacing:1px;opacity:0.8">Change Order Amount</div><div style="font-size:26pt;font-weight:900;margin-top:4px">${fmt(co.amount)}</div></div>
<div><div class="fl" style="margin-bottom:6px">Description of Change</div><div class="desc-box">${co.description||"—"}</div></div>
${co.notes?`<div><div class="fl" style="margin-bottom:6px">Notes</div><div class="desc-box">${co.notes}</div></div>`:""}
${co.client_signature?`<div style="background:#fff;border:1px solid #86efac;border-radius:8px;padding:12px;margin-bottom:12px"><div class="fl" style="color:#166534;margin-bottom:6px">Client Signature — ${co.client_signed_by||""}</div><img src="${co.client_signature}" style="max-height:80px;max-width:300px;object-fit:contain"/></div>`:""}
<div class="sig-grid">
<div class="sig-box"><div style="height:48px"></div><div class="sig-label">Authorized by (AIME)</div><div style="font-size:10pt;font-weight:700;margin-top:4px">${co.submitted_by||""}</div><div style="font-size:9pt;color:#555;margin-top:4px">Date: ______________</div></div>
<div class="sig-box">${co.client_signature?`<img src="${co.client_signature}" style="max-height:60px;max-width:200px;object-fit:contain;display:block;margin-bottom:4px"/>`:`<div style="height:60px;border-bottom:1px solid #000;margin-bottom:4px"></div>`}<div class="sig-label">Accepted by (Client)</div><div style="font-size:10pt;font-weight:700;margin-top:4px">${co.client_signed_by||""}</div><div style="font-size:9pt;color:#555;margin-top:4px">Date: ${co.approved_date||"______________"}</div></div></div>
<div class="footer"><span>AIME Field Pro · CO ${co.co_number} · ${project.name}</span><span>Generated: ${new Date().toLocaleString()}</span></div>
</body></html>`;
    const win=window.open("","_blank","width=900,height=750");win.document.write(html);win.document.close();win.focus();setTimeout(()=>{win.focus();win.print();},1200);
  }

  const contractVal=parseFloat(project.contract_value)||0;
  const approvedCOs=cos.filter(c=>c.status==="Approved").reduce((s,c)=>s+(parseFloat(c.amount)||0),0);
  const pendingCOs=cos.filter(c=>c.status==="Pending").reduce((s,c)=>s+(parseFloat(c.amount)||0),0);
  const revisedContract=contractVal+approvedCOs;
  const fmt=n=>"$"+Number(n||0).toLocaleString("en-US",{minimumFractionDigits:2});

  if(showForm||editing) return(
    <div style={{padding:"14px 16px 80px"}}>
      <button onClick={()=>{setShowForm(false);setEditing(null);setF(blank);}} style={{...ghostBtn,marginBottom:14}}>← Back</button>
      <div style={{fontSize:17,fontWeight:800,marginBottom:16,color:T.text}}>{editing?"Edit":"New"} Change Order</div>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:10}}>
        <div><label style={lbl}>CO Number *</label><input value={f.co_number} onChange={e=>set("co_number",e.target.value)} placeholder="CO-001" style={inp}/></div>
        <div><label style={lbl}>Date</label><input type="date" value={f.date_submitted||""} onChange={e=>set("date_submitted",e.target.value)} style={inp}/></div>
      </div>
      <div style={{marginBottom:10}}><label style={lbl}>Description *</label><textarea rows={3} value={f.description} onChange={e=>set("description",e.target.value)} placeholder="Describe the change in scope..." style={{...inp,resize:"vertical"}}/></div>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:10}}>
        <div><label style={lbl}>Amount ($)</label><input type="number" value={f.amount} onChange={e=>set("amount",e.target.value)} placeholder="0.00" style={inp}/></div>
        <div><label style={lbl}>Status</label><select value={f.status} onChange={e=>set("status",e.target.value)} style={inpSel}>{["Pending","Approved","Rejected"].map(s=><option key={s}>{s}</option>)}</select></div>
      </div>
      {f.status==="Approved"&&<div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:10}}>
        <div><label style={lbl}>Approved By</label><input value={f.approved_by} onChange={e=>set("approved_by",e.target.value)} style={inp}/></div>
        <div><label style={lbl}>Approval Date</label><input type="date" value={f.approved_date||""} onChange={e=>set("approved_date",e.target.value)} style={inp}/></div>
      </div>}
      <div style={{marginBottom:14}}><label style={lbl}>Notes</label><textarea rows={2} value={f.notes} onChange={e=>set("notes",e.target.value)} style={{...inp,resize:"vertical"}}/></div>
      <button onClick={save} disabled={!f.co_number.trim()||saving} style={{...primBtn,opacity:f.co_number.trim()&&!saving?1:0.5,borderRadius:14}}>{saving?"Saving…":"Save Change Order"}</button>
    </div>
  );

  return(
    <div style={{padding:"14px 16px 80px"}}>
      {/* Share Modal */}
      {shareCo&&<div onClick={()=>{setShareCo(null);setCoCopied(false);}} style={{position:"fixed",inset:0,zIndex:300,background:"rgba(0,0,0,0.85)",display:"flex",alignItems:"flex-end",justifyContent:"center"}}>
        <div onClick={e=>e.stopPropagation()} style={{background:T.card,borderRadius:"20px 20px 0 0",padding:"20px 20px 40px",width:"100%",maxWidth:480}}>
          <div style={{fontSize:17,fontWeight:900,marginBottom:4,color:T.text}}>📤 Send CO {shareCo.co_number}</div>
          <div style={{fontSize:12,color:T.muted,marginBottom:16}}>Client opens the link, reviews the CO, signs it, and it saves directly to your app.</div>
          <div style={{background:T.greenLow,border:`1px solid ${T.green}40`,borderRadius:10,padding:"10px 14px",marginBottom:14,display:"flex",justifyContent:"space-between",alignItems:"center"}}>
            <span style={{fontSize:12,color:T.sub}}>{shareCo.description?.substring(0,50)}</span>
            <span style={{fontSize:16,fontWeight:800,color:T.green}}>${Number(shareCo.amount||0).toLocaleString("en-US",{minimumFractionDigits:2})}</span>
          </div>
          <div style={{background:T.surface,border:`1px solid ${T.border}`,borderRadius:12,padding:"10px 14px",marginBottom:14,wordBreak:"break-all",fontSize:12,color:T.sub}}>{`${appUrl}?co=${shareCo.id}`}</div>
          <button onClick={()=>copyCoLink(shareCo)} style={{...primBtn,borderRadius:14,marginBottom:10,background:coCopied?T.green:T.orange,transition:"background 0.2s"}}>{coCopied?"✅ Link Copied! Paste into your email":"📋 Copy Link to Clipboard"}</button>
          {coCopied&&<div style={{background:T.greenLow,border:`1px solid ${T.green}40`,borderRadius:10,padding:"10px 14px",marginBottom:10,fontSize:13,color:T.green,textAlign:"center",fontWeight:600}}>✓ Paste this link into an email or text to your client</div>}
          <div style={{display:"flex",alignItems:"center",gap:10,marginBottom:10}}><div style={{flex:1,height:1,background:T.border}}/><span style={{fontSize:11,color:T.muted}}>OR</span><div style={{flex:1,height:1,background:T.border}}/></div>
          <button onClick={()=>emailCO(shareCo)} style={{...ghostBtn,width:"100%",textAlign:"center",marginBottom:10,fontSize:14}}>📧 Open Email Draft with Link</button>
          <div style={{fontSize:11,color:T.muted,textAlign:"center",marginBottom:14}}>Note: Copy Link above gives a clickable link.</div>
          <button onClick={()=>{setShareCo(null);setCoCopied(false);}} style={{...ghostBtn,width:"100%",textAlign:"center"}}>Done</button>
        </div>
      </div>}

      {contractVal>0&&<div style={{...cardS,marginBottom:14,background:T.blueLow,border:`1px solid ${T.blue}40`}}>
        <div style={{fontSize:12,fontWeight:700,color:T.blue,textTransform:"uppercase",letterSpacing:"1px",marginBottom:10}}>Contract Summary</div>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8,marginBottom:8}}>
          <div style={{background:T.card,borderRadius:10,padding:"8px 10px"}}><div style={{fontSize:11,color:T.muted}}>Original Contract</div><div style={{fontSize:15,fontWeight:800,color:T.text}}>{fmt(contractVal)}</div></div>
          <div style={{background:T.card,borderRadius:10,padding:"8px 10px"}}><div style={{fontSize:11,color:T.muted}}>Approved COs</div><div style={{fontSize:15,fontWeight:800,color:T.green}}>+{fmt(approvedCOs)}</div></div>
        </div>
        {pendingCOs>0&&<div style={{background:T.card,borderRadius:10,padding:"8px 10px",marginBottom:8}}><div style={{fontSize:11,color:T.muted}}>Pending COs</div><div style={{fontSize:15,fontWeight:800,color:T.yellow}}>+{fmt(pendingCOs)}</div></div>}
        <div style={{background:T.card,borderRadius:10,padding:"10px",textAlign:"center",border:`1px solid ${T.blue}40`}}><div style={{fontSize:11,color:T.muted,textTransform:"uppercase",letterSpacing:"1px"}}>Revised Contract Value</div><div style={{fontSize:22,fontWeight:900,color:T.blue}}>{fmt(revisedContract)}</div></div>
      </div>}

      {canEdit&&<button onClick={()=>setShowForm(true)} style={{...primBtn,borderRadius:14,marginBottom:14}}>+ New Change Order</button>}
      {loading&&<Spinner/>}
      {!loading&&cos.length===0&&<div style={{textAlign:"center",padding:"40px 16px",color:T.muted}}>
          <div style={{fontSize:44,marginBottom:12}}>📋</div>
          <div style={{fontSize:15,fontWeight:700,color:T.sub,marginBottom:6}}>No Change Orders</div>
          <div style={{fontSize:12,color:T.muted,lineHeight:1.6}}>{canEdit?`Tap "+ New Change Order" above to document scope changes that need client approval.`:"No change orders have been submitted for this job yet."}</div>
        </div>}
      {cos.map(co=>(
        <div key={co.id} style={{...cardS,marginBottom:10,borderLeft:`3px solid ${statusColor[co.status]||T.muted}`}}>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:8}}>
            <div><div style={{display:"flex",alignItems:"center",gap:8,marginBottom:2}}><span style={{fontSize:15,fontWeight:800,color:T.orange}}>{co.co_number}</span><span style={pill(statusColor[co.status]||T.muted)}>{co.status}</span></div>
            <div style={{fontSize:12,color:T.muted}}>{co.date_submitted} · {co.submitted_by}</div></div>
            <div style={{fontSize:18,fontWeight:900,color:co.status==="Rejected"?T.red:T.green}}>{co.status==="Rejected"?"—":"+"}{fmt(co.amount)}</div>
          </div>
          {co.description&&<div style={{fontSize:13,color:T.text,marginBottom:8,lineHeight:1.5}}>{co.description}</div>}
          {co.status==="Approved"&&co.approved_by&&<div style={{fontSize:11,color:T.green}}>✓ Approved by {co.approved_by}{co.approved_date?" on "+co.approved_date:""}</div>}
          {co.client_signature&&<div style={{background:"#fff",borderRadius:8,padding:4,marginTop:6}}><div style={{fontSize:9,color:"#999",paddingLeft:4,marginBottom:2}}>CLIENT SIGNATURE — {co.client_signed_by||""}</div><img src={co.client_signature} alt="sig" style={{width:"100%",maxHeight:60,objectFit:"contain",borderRadius:6}}/></div>}
          {co.notes&&<div style={{fontSize:11,color:T.muted,marginTop:4,fontStyle:"italic"}}>{co.notes}</div>}
          <div style={{display:"flex",gap:6,marginTop:10,paddingTop:10,borderTop:`1px solid ${T.border}`,flexWrap:"wrap"}}>
            <button onClick={()=>{setShareCo(co);setCoCopied(false);}} style={{...ghostBtn,flex:1,textAlign:"center",fontSize:12,color:T.orange,border:`1px solid ${T.orange}40`,fontWeight:700,minWidth:90}}>📤 Send Link</button>
            <button onClick={()=>printCO(co)} style={{...ghostBtn,flex:1,textAlign:"center",fontSize:12,color:"#CBD5E1",border:"1px solid #27272A",minWidth:70}}>🖨️ PDF</button>
            {canEdit&&<>
              <button onClick={()=>{setEditing(co.id);setF({...co,amount:co.amount||"",client_signature:co.client_signature||null,client_signed_by:co.client_signed_by||""});}} style={{...ghostBtn,flex:1,textAlign:"center",fontSize:12,minWidth:50}}>✏️</button>
              <button onClick={()=>remove(co.id)} style={{...ghostBtn,flex:1,textAlign:"center",fontSize:12,color:T.red,border:`1px solid ${T.red}40`,minWidth:40}}>🗑</button>
            </>}
          </div>
        </div>
      ))}
    </div>
  );
}

function RFIsTab({project,user,onErr}){
  const canEdit=user.role==="admin"||user.role==="pm";
  const [rfis,setRfis]=useState([]);
  const [loading,setLoading]=useState(true);
  const [showForm,setShowForm]=useState(false);
  const [editing,setEditing]=useState(null);
  const [saving,setSaving]=useState(false);
  const [shareRfi,setShareRfi]=useState(null);
  const [copied,setCopied]=useState(false);
  const appUrl=window.location.origin+window.location.pathname;
  const blank={rfi_number:"",date_submitted:today(),submitted_by:user.name,question:"",description:"",due_date:"",response:"",responded_by:"",response_date:"",status:"Open",notes:"",ball_in_court:"",ball_in_court_email:""};
  const [f,setF]=useState(blank);
  const set=(k,v)=>setF(x=>({...x,[k]:v}));
  const statusColor={Open:T.yellow,Answered:T.blue,Closed:T.green,Overdue:T.red};

  useEffect(()=>{load();},[project.id]);
  async function load(){setLoading(true);try{const r=await API.rfis.forProject(project.id);setRfis(Array.isArray(r)?r:[]);}catch(e){onErr(e.message);}setLoading(false);}

  async function save(){
    setSaving(true);
    try{
      const toDate=v=>v&&v.trim()&&v!=="Invalid Date"?v:null;
      const payload={...f,project_id:project.id,date_submitted:toDate(f.date_submitted),due_date:toDate(f.due_date),response_date:toDate(f.response_date)};
      if(editing){await API.rfis.update(editing,payload);}else{await API.rfis.create(payload);}
      setShowForm(false);setEditing(null);setF(blank);await load();
    }catch(e){onErr(e.message);}
    setSaving(false);
  }

  async function remove(id){if(!window.confirm("Delete this RFI?"))return;try{await API.rfis.remove(id);await load();}catch(e){onErr(e.message);}}

  function copyLink(rfi){
    const link=`${appUrl}?rfi=${rfi.id}`;
    navigator.clipboard.writeText(link).then(()=>{setCopied(true);setTimeout(()=>setCopied(false),3000);}).catch(()=>{const el=document.createElement("textarea");el.value=link;document.body.appendChild(el);el.select();document.execCommand("copy");document.body.removeChild(el);setCopied(true);setTimeout(()=>setCopied(false),3000);});
  }

  function openEmailDraft(rfi){
    const link=`${appUrl}?rfi=${rfi.id}`;
    const subj=`RFI #${rfi.rfi_number} — ${project.name} — Response Required`;
    const ln="%0D%0A";const sep="━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━";
    const body=[`Hi ${rfi.ball_in_court||""},`,ln,ln,`Please review and respond to the RFI below.`,ln,ln,sep,ln,`  📋  REQUEST FOR INFORMATION — RFI #${rfi.rfi_number}`,ln,sep,ln,ln,`  Project:    ${project.name}`,ln,rfi.due_date?`  Due By:    ${rfi.due_date}${ln}`:"",ln,`  QUESTION:`,ln,`  ${rfi.question}`,ln,ln,sep,ln,ln,`To respond, click or copy this link into your browser:`,ln,ln,`  ${link}`,ln,ln,`(No login required — fill in the form and click Submit)`,ln,ln,sep,ln,ln,`Thank you,`,ln,`${rfi.submitted_by||"AIME Field Operations"}`,ln,`Atlantic Industrial Mechanical & Environmental Inc.`,ln,].filter(Boolean).join("");
    window.location.href=`mailto:${rfi.ball_in_court_email||""}?subject=${encodeURIComponent(subj)}&body=${body}`;
  }

  function printRFI(rfi){
    const isOverdue=rfi.due_date&&new Date(rfi.due_date)<new Date()&&rfi.status==="Open";
    const effStatus=isOverdue?"Overdue":rfi.status;
    const html=`<!DOCTYPE html><html><head><meta charset="utf-8"><title>RFI ${rfi.rfi_number}</title>
<style>@page{size:letter portrait;margin:0.6in;}*{box-sizing:border-box;margin:0;padding:0;font-family:Arial,sans-serif;}body{font-size:10pt;color:#000;}
.action-bar{background:#1f3864;color:#fff;padding:12px 20px;display:flex;justify-content:space-between;align-items:center;gap:12px;position:sticky;top:0;z-index:100;}
.action-bar h2{font-size:12pt;font-weight:700;margin:0;}.btns{display:flex;gap:10px;}
.btn{padding:8px 16px;border-radius:8px;border:none;font-size:11pt;font-weight:700;cursor:pointer;font-family:inherit;}
.btn-print{background:#60A5FA;color:#000;}.btn-clear{background:transparent;color:#fff;border:1px solid rgba(255,255,255,0.4);}
@media print{.action-bar{display:none!important;}.fillable{border:none!important;background:transparent!important;}}
.doc{max-width:750px;margin:0 auto;padding:24px 20px 40px;}
.lh{display:flex;justify-content:space-between;align-items:flex-start;margin-bottom:20px;padding-bottom:14px;border-bottom:3px solid #1f3864;}
.co{font-size:20pt;font-weight:900;color:#1f3864;}.co-sub{font-size:9pt;color:#555;margin-top:4px;}
.proj-box{background:#f0f4ff;border:1px solid #c7d2fe;border-radius:8px;padding:12px 16px;margin-bottom:16px;display:grid;grid-template-columns:1fr 1fr 1fr;gap:10px;}
.fl{font-size:7.5pt;font-weight:700;text-transform:uppercase;letter-spacing:0.5px;color:#6b7280;margin-bottom:2px;}.fv{font-size:10pt;font-weight:600;color:#111;}
.q-box{background:#1f3864;color:#fff;border-radius:8px;padding:14px 16px;font-size:12pt;font-weight:700;margin-bottom:12px;line-height:1.5;}
.resp-section{background:#f0fdf4;border:2px solid #22c55e;border-radius:10px;padding:16px;margin-bottom:18px;}
.resp-section h2{color:#166534;border-bottom-color:#86efac;margin-bottom:12px;font-size:11pt;font-weight:800;text-transform:uppercase;letter-spacing:0.5px;}
.fill-hint{background:#dcfce7;border:1px solid #86efac;border-radius:6px;padding:8px 12px;font-size:9pt;color:#166534;margin-bottom:12px;font-weight:600;}
.fillable{width:100%;min-height:100px;border:1.5px dashed #22c55e;border-radius:6px;padding:10px 12px;font-size:10pt;line-height:1.7;background:#fff;color:#000;font-family:Arial,sans-serif;resize:vertical;}
.fill-row{display:grid;grid-template-columns:1fr 1fr;gap:12px;margin-top:10px;}
.fill-field{display:flex;flex-direction:column;gap:4px;}.fill-label{font-size:8pt;font-weight:700;color:#166534;text-transform:uppercase;letter-spacing:0.5px;}
input.fillable{min-height:auto;height:38px;}
.sig-grid{display:grid;grid-template-columns:1fr 1fr;gap:20px;margin-top:24px;}
.sig-box{border-top:1.5px solid #000;padding-top:8px;}.sig-label{font-size:8pt;color:#666;text-transform:uppercase;letter-spacing:0.5px;}
.footer{margin-top:24px;padding-top:10px;border-top:1px solid #e5e7eb;font-size:7.5pt;color:#9ca3af;display:flex;justify-content:space-between;}
@media print{body{-webkit-print-color-adjust:exact;print-color-adjust:exact;}}
</style></head><body>
<div class="action-bar"><div><h2>📝 RFI #${rfi.rfi_number} — Fill in your response below, then Print/Save as PDF</h2></div>
<div class="btns"><button class="btn btn-print" onclick="window.print()">🖨️ Print / Save as PDF</button></div></div>
<div class="doc">
<div class="lh"><div><div class="co">AIME</div><div class="co-sub">Atlantic Industrial Mechanical & Environmental Inc.<br>Field Operations Division</div></div>
<div style="text-align:right"><h1 style="font-size:20pt;font-weight:900;color:#1f3864">Request for Information</h1><div>RFI #${rfi.rfi_number}</div></div></div>
<div class="proj-box">
<div><div class="fl">Project</div><div class="fv">${project.name||"—"}</div></div>
<div><div class="fl">Client</div><div class="fv">${project.client||"—"}</div></div>
<div><div class="fl">Division</div><div class="fv">${project.division||"—"}</div></div>
<div><div class="fl">Date Submitted</div><div class="fv">${rfi.date_submitted||"—"}</div></div>
<div><div class="fl">Submitted By</div><div class="fv">${rfi.submitted_by||"—"}</div></div>
<div><div class="fl">Response Due</div><div class="fv">${rfi.due_date||"—"}</div></div></div>
${rfi.ball_in_court?`<div style="background:#fff7ed;border:1.5px solid #fed7aa;border-radius:8px;padding:12px 16px;margin-bottom:16px;display:flex;justify-content:space-between;align-items:center"><div><div class="fl" style="color:#9a3412">🏀 Ball in Court</div><div style="font-size:13pt;font-weight:900;color:#c2410c;margin-top:2px">${rfi.ball_in_court}</div>${rfi.ball_in_court_email?`<div style="font-size:9pt;color:#9a3412">${rfi.ball_in_court_email}</div>`:""}</div></div>`:""}
<div><div class="fl" style="margin-bottom:6px">Question / Issue</div><div class="q-box">${rfi.question||"—"}</div>${rfi.description?`<div style="background:#f9fafb;border:1px solid #e5e7eb;border-radius:6px;padding:12px;font-size:10pt;line-height:1.7">${rfi.description}</div>`:""}</div><br>
${rfi.response?`<div class="resp-section"><h2>Response</h2><div style="background:#fff;border:1px solid #86efac;border-radius:8px;padding:12px"><div style="font-size:8pt;font-weight:700;color:#166534;margin-bottom:4px">From ${rfi.responded_by||"—"} · ${rfi.response_date||""}</div><div>${rfi.response}</div>${rfi.response_signature?`<img src="${rfi.response_signature}" style="max-height:60px;width:100%;object-fit:contain;display:block;margin-top:8px"/>`:""}</div></div>`:
`<div class="resp-section"><h2>Response <span style="font-weight:400;font-size:9pt;text-transform:none">(Please complete and return)</span></h2>
<div class="fill-hint">✏️ Type your response below, then click Print / Save as PDF</div>
<textarea id="resp" class="fillable" placeholder="Enter your response here..." rows="6"></textarea>
<div class="fill-row"><div class="fill-field"><div class="fill-label">Responded By</div><input id="respBy" type="text" class="fillable" placeholder="Your name"/></div>
<div class="fill-field"><div class="fill-label">Response Date</div><input id="respDate" type="date" class="fillable"/></div></div></div>`}
<div class="sig-grid">
<div class="sig-box"><div style="height:48px"></div><div class="sig-label">Submitted by (AIME)</div><div style="font-size:10pt;font-weight:700;margin-top:4px">${rfi.submitted_by||""}</div><div style="font-size:9pt;color:#555;margin-top:4px">Date: ${rfi.date_submitted||"______________"}</div></div>
<div class="sig-box"><div style="height:48px"></div><div class="sig-label">Response by (Client)</div><div style="height:24px"></div><div style="font-size:9pt;color:#555;margin-top:4px">Date: ______________</div></div></div>
<div class="footer"><span>AIME Field Pro · RFI #${rfi.rfi_number} · ${project.name}</span><span>Generated: ${new Date().toLocaleString()}</span></div>
</div></body></html>`;
    const win=window.open("","_blank","width=850,height=800");win.document.write(html);win.document.close();win.focus();
  }

  const open=rfis.filter(r=>r.status==="Open"||r.status==="Overdue").length;
  const answered=rfis.filter(r=>r.status==="Answered").length;

  if(showForm||editing) return(
    <div style={{padding:"14px 16px 80px"}}>
      <button onClick={()=>{setShowForm(false);setEditing(null);setF(blank);}} style={{...ghostBtn,marginBottom:14}}>← Back</button>
      <div style={{fontSize:17,fontWeight:800,marginBottom:16,color:T.text}}>{editing?"Edit":"New"} RFI</div>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:10}}>
        <div><label style={lbl}>RFI Number *</label><input value={f.rfi_number} onChange={e=>set("rfi_number",e.target.value)} placeholder="RFI-001" style={inp}/></div>
        <div><label style={lbl}>Date Submitted</label><input type="date" value={f.date_submitted||""} onChange={e=>set("date_submitted",e.target.value)} style={inp}/></div>
      </div>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:10}}>
        <div><label style={lbl}>Submitted By</label><input value={f.submitted_by} onChange={e=>set("submitted_by",e.target.value)} style={inp}/></div>
        <div><label style={lbl}>Response Due</label><input type="date" value={f.due_date||""} onChange={e=>set("due_date",e.target.value)} style={inp}/></div>
      </div>
      <div style={{marginBottom:10}}><label style={lbl}>Question / Issue *</label><input value={f.question} onChange={e=>set("question",e.target.value)} placeholder="Brief summary of the question..." style={inp}/></div>
      <div style={{marginBottom:10}}><label style={lbl}>Full Description</label><textarea rows={3} value={f.description} onChange={e=>set("description",e.target.value)} placeholder="Detailed description..." style={{...inp,resize:"vertical"}}/></div>
      <div style={{marginBottom:10}}><label style={lbl}>Status</label><select value={f.status} onChange={e=>set("status",e.target.value)} style={inpSel}>{["Open","Answered","Closed","Overdue"].map(s=><option key={s}>{s}</option>)}</select></div>
      <div style={{...cardS,marginBottom:10,borderLeft:`3px solid ${T.orange}`,background:T.orangeLow}}>
        <div style={{fontSize:12,fontWeight:700,color:T.orange,marginBottom:8}}>🏀 BALL IN COURT</div>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}>
          <div><label style={lbl}>Person Responsible</label><input value={f.ball_in_court} onChange={e=>set("ball_in_court",e.target.value)} placeholder="Name of person who needs to act" style={inp}/></div>
          <div><label style={lbl}>Their Email</label><input type="email" value={f.ball_in_court_email} onChange={e=>set("ball_in_court_email",e.target.value)} placeholder="email@company.com" style={inp}/></div>
        </div>
      </div>
      {(f.status==="Answered"||f.status==="Closed")&&<>
        <div style={{marginBottom:10}}><label style={lbl}>Response</label><textarea rows={3} value={f.response} onChange={e=>set("response",e.target.value)} style={{...inp,resize:"vertical"}}/></div>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:10}}>
          <div><label style={lbl}>Responded By</label><input value={f.responded_by} onChange={e=>set("responded_by",e.target.value)} style={inp}/></div>
          <div><label style={lbl}>Response Date</label><input type="date" value={f.response_date||""} onChange={e=>set("response_date",e.target.value)} style={inp}/></div>
        </div>
      </>}
      <button onClick={save} disabled={!f.rfi_number.trim()||!f.question.trim()||saving} style={{...primBtn,opacity:f.rfi_number.trim()&&f.question.trim()&&!saving?1:0.5,borderRadius:14}}>{saving?"Saving…":"Save RFI"}</button>
    </div>
  );

  return(
    <div style={{padding:"14px 16px 80px"}}>
      {/* Share Modal */}
      {shareRfi&&<div onClick={()=>setShareRfi(null)} style={{position:"fixed",inset:0,zIndex:300,background:"rgba(0,0,0,0.85)",display:"flex",alignItems:"flex-end",justifyContent:"center"}}>
        <div onClick={e=>e.stopPropagation()} style={{background:T.card,borderRadius:"20px 20px 0 0",padding:"20px 20px 40px",width:"100%",maxWidth:480}}>
          <div style={{fontSize:17,fontWeight:900,marginBottom:4,color:T.text}}>📤 Share RFI #{shareRfi.rfi_number}</div>
          <div style={{fontSize:12,color:T.muted,marginBottom:20}}>Send this link to {shareRfi.ball_in_court||"the recipient"} — they fill in their response and it saves directly to your app.</div>
          <div style={{background:T.surface,border:`1px solid ${T.border}`,borderRadius:12,padding:"12px 14px",marginBottom:14,wordBreak:"break-all",fontSize:12,color:T.sub}}>{`${appUrl}?rfi=${shareRfi.id}`}</div>
          <button onClick={()=>copyLink(shareRfi)} style={{...primBtn,borderRadius:14,marginBottom:10,background:copied?T.green:T.orange,transition:"background 0.2s"}}>{copied?"✅ Link Copied! Paste it into your email":"📋 Copy Link to Clipboard"}</button>
          {copied&&<div style={{background:T.greenLow,border:`1px solid ${T.green}40`,borderRadius:10,padding:"10px 14px",marginBottom:10,fontSize:13,color:T.green,textAlign:"center",fontWeight:600}}>✓ Link copied! Open your email app, paste it, and send.</div>}
          <div style={{display:"flex",alignItems:"center",gap:10,marginBottom:10}}><div style={{flex:1,height:1,background:T.border}}/><span style={{fontSize:11,color:T.muted}}>OR</span><div style={{flex:1,height:1,background:T.border}}/></div>
          <button onClick={()=>openEmailDraft(shareRfi)} style={{...ghostBtn,width:"100%",textAlign:"center",marginBottom:10,fontSize:14}}>📧 Open Email Draft {shareRfi.ball_in_court_email?`(to ${shareRfi.ball_in_court_email})`:""}</button>
          <div style={{fontSize:11,color:T.muted,textAlign:"center",marginBottom:14}}>Note: Copy Link gives a clickable link. Email draft is plain text only.</div>
          <button onClick={()=>setShareRfi(null)} style={{...ghostBtn,width:"100%",textAlign:"center"}}>Done</button>
        </div>
      </div>}

      {rfis.length>0&&<div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:8,marginBottom:14}}>
        {[[open,"Open / Overdue",T.yellow],[answered,"Answered",T.blue],[rfis.filter(r=>r.status==="Closed").length,"Closed",T.green]].map(([v,l,c])=>(
          <div key={l} style={{background:T.card,borderRadius:10,padding:"10px",textAlign:"center",border:`1px solid ${c}30`}}><div style={{fontSize:20,fontWeight:900,color:c}}>{v}</div><div style={{fontSize:9,color:T.muted,textTransform:"uppercase",marginTop:2}}>{l}</div></div>
        ))}
      </div>}
      {canEdit&&<button onClick={()=>setShowForm(true)} style={{...primBtn,borderRadius:14,marginBottom:14}}>+ New RFI</button>}
      {loading&&<Spinner/>}
      {!loading&&rfis.length===0&&<div style={{textAlign:"center",padding:"40px 16px",color:T.muted}}>
          <div style={{fontSize:44,marginBottom:12}}>📝</div>
          <div style={{fontSize:15,fontWeight:700,color:T.sub,marginBottom:6}}>No RFIs</div>
          <div style={{fontSize:12,color:T.muted,lineHeight:1.6}}>{canEdit?`Tap "+ New RFI" above to submit a question to Colonial Pipeline or the engineer.`:"No requests for information have been submitted yet."}</div>
        </div>}
      {rfis.map(rfi=>{
        const isOverdue=rfi.due_date&&new Date(rfi.due_date)<new Date()&&rfi.status==="Open";
        const effStatus=isOverdue?"Overdue":rfi.status;
        const sc=statusColor[effStatus]||T.muted;
        return(
          <div key={rfi.id} style={{...cardS,marginBottom:10,borderLeft:`3px solid ${sc}`}}>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:8}}>
              <div>
                <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:2}}><span style={{fontSize:15,fontWeight:800,color:T.orange}}>{rfi.rfi_number}</span><span style={pill(sc)}>{effStatus.toUpperCase()}</span></div>
                <div style={{fontSize:12,color:T.muted}}>{rfi.date_submitted} · {rfi.submitted_by}</div>
                {rfi.due_date&&<div style={{fontSize:11,color:isOverdue?T.red:T.muted,marginTop:2}}>{isOverdue?"⚠️ Overdue — ":"Due: "}{rfi.due_date}</div>}
                {rfi.ball_in_court&&<div style={{fontSize:11,color:T.orange,marginTop:2,fontWeight:600}}>🏀 Ball in Court: {rfi.ball_in_court}</div>}
              </div>
            </div>
            <div style={{fontSize:14,fontWeight:700,color:T.text,marginBottom:4}}>{rfi.question}</div>
            {rfi.description&&<div style={{fontSize:12,color:T.sub,marginBottom:8,lineHeight:1.5}}>{rfi.description}</div>}
            {rfi.response&&<div style={{background:T.greenLow,border:`1px solid ${T.green}40`,borderRadius:10,padding:"10px",marginTop:8}}>
              <div style={{fontSize:11,fontWeight:700,color:T.green,marginBottom:4}}>✓ RESPONSE — {rfi.responded_by}{rfi.response_date?" · "+rfi.response_date:""}</div>
              <div style={{fontSize:12,color:T.text,lineHeight:1.5}}>{rfi.response}</div>
              {rfi.response_signature&&<div style={{marginTop:8,background:"#fff",borderRadius:8,padding:4}}><div style={{fontSize:9,color:"#999",marginBottom:2,paddingLeft:4}}>SIGNATURE</div><img src={rfi.response_signature} alt="Signature" style={{width:"100%",maxHeight:80,objectFit:"contain",display:"block",borderRadius:6}}/></div>}
            </div>}
            <div style={{display:"flex",gap:6,marginTop:10,paddingTop:10,borderTop:`1px solid ${T.border}`,flexWrap:"wrap"}}>
              <button onClick={()=>{setShareRfi(rfi);setCopied(false);}} style={{...ghostBtn,flex:1,textAlign:"center",fontSize:12,color:T.orange,border:`1px solid ${T.orange}40`,fontWeight:700,minWidth:90}}>📤 Send Link</button>
              <button onClick={()=>printRFI(rfi)} style={{...ghostBtn,flex:1,textAlign:"center",fontSize:12,color:"#CBD5E1",border:"1px solid #27272A",minWidth:70}}>🖨️ PDF</button>
              {canEdit&&<>
                <button onClick={()=>{setEditing(rfi.id);setF({...rfi,notes:rfi.notes||"",ball_in_court:rfi.ball_in_court||"",ball_in_court_email:rfi.ball_in_court_email||""});}} style={{...ghostBtn,flex:1,textAlign:"center",fontSize:12,minWidth:60}}>✏️</button>
                <button onClick={()=>remove(rfi.id)} style={{...ghostBtn,flex:1,textAlign:"center",fontSize:12,color:T.red,border:`1px solid ${T.red}40`,minWidth:50}}>🗑</button>
              </>}
            </div>
          </div>
        );
      })}
    </div>
  );
}

function PublicRFIForm({rfiId}){
  const [rfi,setRfi]=useState(null);const [project,setProject]=useState(null);
  const [loading,setLoading]=useState(true);const [saving,setSaving]=useState(false);
  const [submitted,setSubmitted]=useState(false);const [err,setErr]=useState("");
  const [resp,setResp]=useState("");const [respBy,setRespBy]=useState("");
  const [respDate,setRespDate]=useState(today());const [respTitle,setRespTitle]=useState("");
  const [sigData,setSigData]=useState(null);const [drawing,setDrawing]=useState(false);
  const sigRef=React.useRef(null);
  useEffect(()=>{(async()=>{try{
    const d=await sb(`/rfis?id=eq.${rfiId}&limit=1`);const r=Array.isArray(d)?d[0]:d;
    if(!r){setErr("RFI not found.");setLoading(false);return;}
    setRfi(r);if(r.response){setResp(r.response);setRespBy(r.responded_by||"");setSubmitted(true);}
    const pd=await sb(`/projects?id=eq.${r.project_id}&limit=1`);
    setProject(Array.isArray(pd)?pd[0]:pd||{name:"Unknown"});
  }catch(e){setErr(e.message);}setLoading(false);})();},[rfiId]);
  function getPos(e,c){const r=c.getBoundingClientRect();const sx=c.width/r.width;const sy=c.height/r.height;if(e.touches)return{x:(e.touches[0].clientX-r.left)*sx,y:(e.touches[0].clientY-r.top)*sy};return{x:(e.clientX-r.left)*sx,y:(e.clientY-r.top)*sy};}
  function startSig(e){e.preventDefault();const c=sigRef.current;if(!c)return;const p=getPos(e,c);c.getContext("2d").beginPath();c.getContext("2d").moveTo(p.x,p.y);setDrawing(true);}
  function drawSig(e){e.preventDefault();if(!drawing)return;const c=sigRef.current;if(!c)return;const ctx=c.getContext("2d");const p=getPos(e,c);ctx.lineWidth=2.5;ctx.lineCap="round";ctx.strokeStyle="#1f3864";ctx.lineTo(p.x,p.y);ctx.stroke();ctx.beginPath();ctx.moveTo(p.x,p.y);}
  function endSig(e){e.preventDefault();setDrawing(false);const c=sigRef.current;if(c)setSigData(c.toDataURL("image/jpeg",0.7));}
  function clearSig(){const c=sigRef.current;if(c)c.getContext("2d").clearRect(0,0,c.width,c.height);setSigData(null);}
  async function submit(){
    if(!resp.trim()||!respBy.trim()||!sigData){setErr("Please complete all required fields and sign.");return;}
    setSaving(true);setErr("");
    try{await sb(`/rfis?id=eq.${rfiId}`,{method:"PATCH",body:{response:resp,responded_by:respBy+(respTitle?" ("+respTitle+")":""),response_date:respDate||today(),status:"Answered",response_signature:sigData},prefer:"return=representation"});setSubmitted(true);}
    catch(e){setErr("Failed: "+e.message);}setSaving(false);
  }
  const s={bg:"#0D0D0F",card:"#1A1A20",inp:{width:"100%",background:"#0C0C0F",border:"1px solid #27272A",borderRadius:10,color:"#fff",fontSize:14,padding:"12px 14px",fontFamily:"inherit",outline:"none"},lbl:{display:"block",fontSize:11,fontWeight:700,color:"#A1A1AA",textTransform:"uppercase",letterSpacing:"1px",marginBottom:6}};
  if(loading)return <div style={{background:s.bg,minHeight:"100vh",display:"flex",alignItems:"center",justifyContent:"center"}}><div style={{color:"#fff"}}>Loading...</div></div>;
  if(err&&!rfi)return <div style={{background:s.bg,minHeight:"100vh",display:"flex",alignItems:"center",justifyContent:"center",padding:20}}><div style={{background:s.card,borderRadius:16,padding:32,maxWidth:400,textAlign:"center"}}><div style={{fontSize:40}}>⚠️</div><div style={{color:"#fff",fontWeight:700,margin:"8px 0"}}>{err}</div></div></div>;
  if(submitted)return <div style={{background:s.bg,minHeight:"100vh",fontFamily:"system-ui",padding:20}}><div style={{maxWidth:600,margin:"0 auto",background:s.card,borderRadius:16,padding:32,textAlign:"center",border:"1px solid #22C55E40"}}><div style={{fontSize:48}}>✅</div><div style={{fontSize:20,fontWeight:800,color:"#34D399",margin:"8px 0"}}>Response Submitted!</div><div style={{color:"#A1A1AA"}}>RFI #{rfi?.rfi_number} response has been saved.</div></div></div>;
  return(
    <div style={{background:s.bg,minHeight:"100vh",fontFamily:"system-ui",color:"#fff",padding:"16px 16px 60px"}}>
      <div style={{maxWidth:600,margin:"0 auto"}}>
        <div style={{background:"#1f3864",borderRadius:16,padding:"16px 20px",marginBottom:16,display:"flex",justifyContent:"space-between",alignItems:"center"}}>
          <div><div style={{fontSize:20,fontWeight:900,color:"#60A5FA"}}>AIME</div><div style={{fontSize:11,color:"rgba(255,255,255,0.7)"}}>Atlantic Industrial Mechanical & Environmental Inc.</div></div>
          <div style={{textAlign:"right"}}><div style={{fontSize:18,fontWeight:800}}>RFI #{rfi.rfi_number}</div></div>
        </div>
        <div style={{background:"#1f3864",borderRadius:12,padding:16,marginBottom:12}}>
          <div style={{fontSize:10,color:"rgba(255,255,255,0.6)",textTransform:"uppercase",marginBottom:6}}>Question / Issue</div>
          <div style={{fontSize:16,fontWeight:700,lineHeight:1.5}}>{rfi.question}</div>
          {rfi.description&&<div style={{fontSize:13,color:"rgba(255,255,255,0.7)",marginTop:8}}>{rfi.description}</div>}
        </div>
        {err&&<div style={{background:"#EF444420",border:"1px solid #EF4444",borderRadius:10,padding:"10px 14px",marginBottom:12,fontSize:13,color:"#FC8181"}}>{err}</div>}
        <div style={{background:s.card,borderRadius:16,padding:20,border:"2px solid #22C55E40"}}>
          <div style={{fontSize:14,fontWeight:800,color:"#34D399",marginBottom:14}}>✏️ Your Response</div>
          <div style={{marginBottom:12}}><label style={s.lbl}>Response *</label><textarea rows={6} value={resp} onChange={e=>setResp(e.target.value)} placeholder="Enter your response..." style={{...s.inp,resize:"vertical",lineHeight:1.6}}/></div>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:12}}>
            <div><label style={s.lbl}>Your Name *</label><input value={respBy} onChange={e=>setRespBy(e.target.value)} placeholder="Full name" style={s.inp}/></div>
            <div><label style={s.lbl}>Date</label><input type="date" value={respDate} onChange={e=>setRespDate(e.target.value)} style={s.inp}/></div>
          </div>
          <div style={{marginBottom:14}}><label style={s.lbl}>Company / Title</label><input value={respTitle} onChange={e=>setRespTitle(e.target.value)} placeholder="e.g. Colonial Pipeline · Engineer" style={s.inp}/></div>
          <div style={{marginBottom:16}}>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:6}}><label style={s.lbl}>Signature *</label>{sigData&&<button onClick={clearSig} style={{background:"none",border:"none",color:"#FC8181",fontSize:12,cursor:"pointer"}}>Clear</button>}</div>
            <div style={{background:"#fff",borderRadius:10,overflow:"hidden",border:"2px solid #22C55E"}}>
              <canvas ref={sigRef} width={600} height={160} style={{width:"100%",height:160,display:"block",touchAction:"none",cursor:"crosshair"}}
                onMouseDown={startSig} onMouseMove={drawSig} onMouseUp={endSig} onMouseLeave={endSig}
                onTouchStart={startSig} onTouchMove={drawSig} onTouchEnd={endSig}/>
            </div>
            {!sigData&&<div style={{textAlign:"center",fontSize:11,color:"#71717A",marginTop:4}}>Draw your signature above</div>}
            {sigData&&<div style={{textAlign:"center",fontSize:11,color:"#34D399",marginTop:4}}>✓ Signature captured</div>}
          </div>
          <button onClick={submit} disabled={saving||!resp.trim()||!respBy.trim()||!sigData}
            style={{width:"100%",background:saving||!resp.trim()||!respBy.trim()||!sigData?"#26262E":"#34D399",color:saving||!resp.trim()||!respBy.trim()||!sigData?"#71717A":"#000",border:"none",borderRadius:12,padding:14,fontSize:16,fontWeight:800,cursor:"pointer",fontFamily:"inherit"}}>
            {saving?"Submitting…":"✅ Submit Signed Response"}
          </button>
        </div>
      </div>
    </div>
  );
}

function PublicCOForm({coId}){
  const [co,setCo]=useState(null);const [project,setProject]=useState(null);
  const [loading,setLoading]=useState(true);const [saving,setSaving]=useState(false);
  const [submitted,setSubmitted]=useState(false);const [err,setErr]=useState("");
  const [signerName,setSignerName]=useState("");const [signerTitle,setSignerTitle]=useState("");
  const [sigData,setSigData]=useState(null);const [drawing,setDrawing]=useState(false);
  const sigRef=React.useRef(null);
  useEffect(()=>{(async()=>{try{
    const d=await sb(`/change_orders?id=eq.${coId}&limit=1`);const c=Array.isArray(d)?d[0]:d;
    if(!c){setErr("CO not found.");setLoading(false);return;}
    setCo(c);if(c.client_signature)setSubmitted(true);
    const pd=await sb(`/projects?id=eq.${c.project_id}&limit=1`);
    setProject(Array.isArray(pd)?pd[0]:pd||{name:"Unknown"});
  }catch(e){setErr(e.message);}setLoading(false);})();},[coId]);
  function getPos(e,c){const r=c.getBoundingClientRect();const sx=c.width/r.width;const sy=c.height/r.height;if(e.touches)return{x:(e.touches[0].clientX-r.left)*sx,y:(e.touches[0].clientY-r.top)*sy};return{x:(e.clientX-r.left)*sx,y:(e.clientY-r.top)*sy};}
  function startSig(e){e.preventDefault();const c=sigRef.current;if(!c)return;const p=getPos(e,c);c.getContext("2d").beginPath();c.getContext("2d").moveTo(p.x,p.y);setDrawing(true);}
  function drawSig(e){e.preventDefault();if(!drawing)return;const c=sigRef.current;if(!c)return;const ctx=c.getContext("2d");const p=getPos(e,c);ctx.lineWidth=2.5;ctx.lineCap="round";ctx.strokeStyle="#1f3864";ctx.lineTo(p.x,p.y);ctx.stroke();ctx.beginPath();ctx.moveTo(p.x,p.y);}
  function endSig(e){e.preventDefault();setDrawing(false);const c=sigRef.current;if(c)setSigData(c.toDataURL("image/jpeg",0.7));}
  function clearSig(){const c=sigRef.current;if(c)c.getContext("2d").clearRect(0,0,c.width,c.height);setSigData(null);}
  async function submit(){
    if(!signerName.trim()||!sigData){setErr("Please enter your name and signature.");return;}
    setSaving(true);setErr("");
    try{await sb(`/change_orders?id=eq.${coId}`,{method:"PATCH",body:{client_signature:sigData,client_signed_by:signerName+(signerTitle?" ("+signerTitle+")":""),client_signed_date:today(),status:"Approved"},prefer:"return=representation"});setSubmitted(true);}
    catch(e){setErr("Failed: "+e.message);}setSaving(false);
  }
  const s={bg:"#0D0D0F",card:"#1A1A20",inp:{width:"100%",background:"#0C0C0F",border:"1px solid #27272A",borderRadius:10,color:"#fff",fontSize:14,padding:"12px 14px",fontFamily:"inherit",outline:"none"},lbl:{display:"block",fontSize:11,fontWeight:700,color:"#A1A1AA",textTransform:"uppercase",letterSpacing:"1px",marginBottom:6}};
  const fmt=n=>"$"+Number(n||0).toLocaleString("en-US",{minimumFractionDigits:2});
  if(loading)return <div style={{background:s.bg,minHeight:"100vh",display:"flex",alignItems:"center",justifyContent:"center"}}><div style={{color:"#fff"}}>Loading...</div></div>;
  if(err&&!co)return <div style={{background:s.bg,minHeight:"100vh",display:"flex",alignItems:"center",justifyContent:"center",padding:20}}><div style={{background:s.card,borderRadius:16,padding:32,maxWidth:400,textAlign:"center"}}><div style={{fontSize:40}}>⚠️</div><div style={{color:"#fff",fontWeight:700,margin:"8px 0"}}>{err}</div></div></div>;
  if(submitted)return <div style={{background:s.bg,minHeight:"100vh",fontFamily:"system-ui",padding:20}}><div style={{maxWidth:560,margin:"0 auto",background:s.card,borderRadius:16,padding:32,textAlign:"center",border:"1px solid #22C55E40"}}><div style={{fontSize:48}}>✅</div><div style={{fontSize:20,fontWeight:800,color:"#34D399",margin:"8px 0"}}>Change Order Signed!</div><div style={{color:"#A1A1AA"}}>CO {co?.co_number} has been approved.</div></div></div>;
  return(
    <div style={{background:s.bg,minHeight:"100vh",fontFamily:"system-ui",color:"#fff",padding:"16px 16px 60px"}}>
      <div style={{maxWidth:560,margin:"0 auto"}}>
        <div style={{background:"#1f3864",borderRadius:16,padding:"16px 20px",marginBottom:16,display:"flex",justifyContent:"space-between",alignItems:"center"}}>
          <div><div style={{fontSize:20,fontWeight:900,color:"#60A5FA"}}>AIME</div><div style={{fontSize:11,color:"rgba(255,255,255,0.7)"}}>Atlantic Industrial Mechanical & Environmental Inc.</div></div>
          <div><div style={{fontSize:18,fontWeight:800}}>{co.co_number}</div><div style={{fontSize:11,color:"rgba(255,255,255,0.6)"}}>Change Order</div></div>
        </div>
        <div style={{background:"#1f3864",borderRadius:12,padding:16,marginBottom:12}}>
          <div style={{fontSize:10,color:"rgba(255,255,255,0.6)",textTransform:"uppercase",marginBottom:6}}>Description of Change</div>
          <div style={{fontSize:15,fontWeight:700,lineHeight:1.5,marginBottom:12}}>{co.description}</div>
          <div style={{background:"rgba(255,255,255,0.1)",borderRadius:10,padding:12,textAlign:"center"}}>
            <div style={{fontSize:10,color:"rgba(255,255,255,0.6)",textTransform:"uppercase",letterSpacing:"1px"}}>Change Order Amount</div>
            <div style={{fontSize:28,fontWeight:900,color:"#34D399",marginTop:4}}>{fmt(co.amount)}</div>
          </div>
        </div>
        {err&&<div style={{background:"#EF444420",border:"1px solid #EF4444",borderRadius:10,padding:"10px 14px",marginBottom:12,fontSize:13,color:"#FC8181"}}>{err}</div>}
        <div style={{background:s.card,borderRadius:16,padding:20,border:"2px solid #60A5FA40"}}>
          <div style={{fontSize:14,fontWeight:800,color:"#60A5FA",marginBottom:14}}>✍️ Sign & Approve Change Order</div>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:12}}>
            <div><label style={s.lbl}>Your Name *</label><input value={signerName} onChange={e=>setSignerName(e.target.value)} placeholder="Full name" style={s.inp}/></div>
            <div><label style={s.lbl}>Title / Company</label><input value={signerTitle} onChange={e=>setSignerTitle(e.target.value)} placeholder="Your title" style={s.inp}/></div>
          </div>
          <div style={{marginBottom:16}}>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:6}}><label style={s.lbl}>Signature *</label>{sigData&&<button onClick={clearSig} style={{background:"none",border:"none",color:"#FC8181",fontSize:12,cursor:"pointer"}}>Clear</button>}</div>
            <div style={{background:"#fff",borderRadius:10,overflow:"hidden",border:"2px solid #60A5FA"}}>
              <canvas ref={sigRef} width={600} height={160} style={{width:"100%",height:160,display:"block",touchAction:"none",cursor:"crosshair"}}
                onMouseDown={startSig} onMouseMove={drawSig} onMouseUp={endSig} onMouseLeave={endSig}
                onTouchStart={startSig} onTouchMove={drawSig} onTouchEnd={endSig}/>
            </div>
            {!sigData&&<div style={{textAlign:"center",fontSize:11,color:"#71717A",marginTop:4}}>Draw your signature above</div>}
            {sigData&&<div style={{textAlign:"center",fontSize:11,color:"#34D399",marginTop:4}}>✓ Signature captured</div>}
          </div>
          <button onClick={submit} disabled={saving||!signerName.trim()||!sigData}
            style={{width:"100%",background:saving||!signerName.trim()||!sigData?"#26262E":"#60A5FA",color:saving||!signerName.trim()||!sigData?"#71717A":"#000",border:"none",borderRadius:12,padding:14,fontSize:16,fontWeight:800,cursor:"pointer",fontFamily:"inherit"}}>
            {saving?"Submitting…":"✅ Sign & Approve Change Order"}
          </button>
        </div>
      </div>
    </div>
  );
}

async function setupPushNotifications(){
  const isNative=typeof window!=="undefined"&&window.Capacitor?.isNativePlatform?.();
  if(!isNative) return;
  try{
    const cap=window.Capacitor?.Plugins;
    if(!cap) return;
    const{PushNotifications}=cap;
    if(!PushNotifications) return;
    const perm=await PushNotifications.requestPermissions();
    if(perm.receive==="granted"){
      await PushNotifications.register();
    }
  }catch(e){
  }
}

function PublicInspectorForm({reportId}){
  const [report,setReport]=useState(null);const [project,setProject]=useState(null);
  const [loading,setLoading]=useState(true);const [saving,setSaving]=useState(false);
  const [submitted,setSubmitted]=useState(false);const [err,setErr]=useState("");
  const [inspName,setInspName]=useState("");const [inspTitle,setInspTitle]=useState("");
  const [inspCo,setInspCo]=useState("");const [sigData,setSigData]=useState(null);
  const [drawing,setDrawing]=useState(false);
  const sigRef=React.useRef(null);

  useEffect(()=>{(async()=>{try{
    const d=await sb(`/daily_reports?id=eq.${reportId}&limit=1`);
    const r=Array.isArray(d)?d[0]:d;
    if(!r){setErr("Report not found.");setLoading(false);return;}
    setReport(r);
    if(r.inspector_signature){setSubmitted(true);}
    const pd=await sb(`/projects?id=eq.${r.project_id}&limit=1`);
    setProject(Array.isArray(pd)?pd[0]:pd||{name:"Unknown"});
  }catch(e){setErr(e.message);}setLoading(false);})();},[reportId]);

  function getPos(e,c){const r=c.getBoundingClientRect();const sx=c.width/r.width;const sy=c.height/r.height;if(e.touches)return{x:(e.touches[0].clientX-r.left)*sx,y:(e.touches[0].clientY-r.top)*sy};return{x:(e.clientX-r.left)*sx,y:(e.clientY-r.top)*sy};}
  function startSig(e){e.preventDefault();const c=sigRef.current;if(!c)return;const p=getPos(e,c);c.getContext("2d").beginPath();c.getContext("2d").moveTo(p.x,p.y);setDrawing(true);}
  function drawSig(e){e.preventDefault();if(!drawing)return;const c=sigRef.current;if(!c)return;const ctx=c.getContext("2d");const p=getPos(e,c);ctx.lineWidth=2.5;ctx.lineCap="round";ctx.strokeStyle="#1f3864";ctx.lineTo(p.x,p.y);ctx.stroke();ctx.beginPath();ctx.moveTo(p.x,p.y);}
  function endSig(e){e.preventDefault();setDrawing(false);const c=sigRef.current;if(c)setSigData(c.toDataURL("image/jpeg",0.7));}
  function clearSig(){const c=sigRef.current;if(c)c.getContext("2d").clearRect(0,0,c.width,c.height);setSigData(null);}

  async function submit(){
    if(!inspName.trim()||!sigData){setErr("Please enter your name and sign.");return;}
    setSaving(true);setErr("");
    try{
      const fullName=inspName.trim()+(inspTitle?` · ${inspTitle}`:"")+(inspCo?` · ${inspCo}`:"");
      await sb(`/daily_reports?id=eq.${reportId}`,{method:"PATCH",body:{
        inspector_name:fullName,
        inspector_signature:sigData,
        inspector_signed_at:new Date().toISOString(),
      },prefer:"return=representation"});
      setSubmitted(true);
    }catch(e){setErr("Failed: "+e.message);}setSaving(false);
  }

  const s={bg:"#0D0D0F",card:"#1A1A20",border:"#26262E",inp:{width:"100%",background:"#141418",border:"1px solid #26262E",borderRadius:10,color:"#F0F4FF",fontSize:14,padding:"12px 14px",fontFamily:"inherit",outline:"none"},lbl:{display:"block",fontSize:11,fontWeight:700,color:"#7080A0",textTransform:"uppercase",letterSpacing:"1px",marginBottom:6}};

  if(loading)return<div style={{background:s.bg,minHeight:"100vh",display:"flex",alignItems:"center",justifyContent:"center",fontFamily:"system-ui"}}><div style={{color:"#F0F4FF"}}>Loading report...</div></div>;
  if(err&&!report)return<div style={{background:s.bg,minHeight:"100vh",display:"flex",alignItems:"center",justifyContent:"center",fontFamily:"system-ui",padding:20}}><div style={{background:s.card,borderRadius:16,padding:32,maxWidth:400,textAlign:"center"}}><div style={{fontSize:40}}>⚠️</div><div style={{color:"#F0F4FF",fontWeight:700,margin:"8px 0"}}>{err}</div></div></div>;
  if(submitted)return(
    <div style={{background:s.bg,minHeight:"100vh",fontFamily:"system-ui",padding:20}}>
      <div style={{maxWidth:560,margin:"0 auto",background:s.card,borderRadius:16,padding:32,textAlign:"center",border:"1px solid #34D39940"}}>
        <div style={{fontSize:48,marginBottom:12}}>✅</div>
        <div style={{fontSize:20,fontWeight:800,color:"#34D399",marginBottom:8}}>Report Signed!</div>
        <div style={{color:"#7080A0",marginBottom:16}}>Daily Report for {report?.date} has been signed and saved.</div>
        {report?.inspector_signature&&<div style={{background:"#fff",borderRadius:10,padding:8,marginTop:8}}><img src={report.inspector_signature} alt="sig" style={{maxHeight:80,width:"100%",objectFit:"contain"}}/></div>}
        <div style={{fontSize:12,color:"#555",marginTop:12}}>You may close this window.</div>
      </div>
    </div>
  );

  return(
    <div style={{background:s.bg,minHeight:"100vh",fontFamily:"system-ui",color:"#F0F4FF",padding:"16px 16px 60px"}}>
      <div style={{maxWidth:560,margin:"0 auto"}}>
        {/* Header */}
        <div style={{background:"#141418",borderRadius:16,padding:"16px 20px",marginBottom:16,border:"1px solid #26262E"}}>
          <div style={{fontSize:20,fontWeight:900,color:"#60A5FA",marginBottom:2}}>AIME</div>
          <div style={{fontSize:11,color:"#7080A0"}}>Atlantic Industrial Mechanical & Environmental Inc.</div>
        </div>
        <div style={{background:s.card,borderRadius:12,padding:"14px 16px",marginBottom:14,border:"1px solid #26262E"}}>
          <div style={{fontSize:11,fontWeight:700,color:"#7080A0",textTransform:"uppercase",letterSpacing:"1px",marginBottom:10}}>Daily Report — Inspector Sign-Off</div>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}>
            {[["Project",project?.name||"—"],["Date",report?.date||"—"],["Submitted By",report?.submitted_by||"—"],["Workers On Site",(report?.labor||[]).length||report?.manpower_count||"—"]].map(([l,v])=>(
              <div key={l}><div style={{fontSize:10,color:"#7080A0",fontWeight:700,textTransform:"uppercase"}}>{l}</div><div style={{fontSize:13,fontWeight:600,color:"#F0F4FF",marginTop:2}}>{v}</div></div>
            ))}
          </div>
          {report?.site_conditions&&<div style={{marginTop:10,paddingTop:10,borderTop:"1px solid #26262E"}}>
            <div style={{fontSize:10,color:"#7080A0",fontWeight:700,textTransform:"uppercase",marginBottom:4}}>Site Conditions</div>
            <div style={{fontSize:13,color:"#C8D4F0"}}>{report.site_conditions}</div>
          </div>}
          {report?.description&&<div style={{marginTop:8}}>
            <div style={{fontSize:10,color:"#7080A0",fontWeight:700,textTransform:"uppercase",marginBottom:4}}>Work Description</div>
            <div style={{fontSize:12,color:"#C8D4F0",lineHeight:1.5}}>{report.description}</div>
          </div>}
        </div>

        {err&&<div style={{background:"#FC818120",border:"1px solid #FC8181",borderRadius:10,padding:"10px 14px",marginBottom:12,fontSize:13,color:"#FC8181"}}>{err}</div>}

        {/* Sign form */}
        <div style={{background:s.card,borderRadius:16,padding:20,border:"1px solid #26262E"}}>
          <div style={{fontSize:14,fontWeight:800,color:"#60A5FA",marginBottom:14}}>✍️ Inspector Sign-Off</div>

          <div style={{marginBottom:10}}><label style={s.lbl}>Inspector Name *</label><input value={inspName} onChange={e=>setInspName(e.target.value)} placeholder="Full name" style={s.inp}/></div>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8,marginBottom:14}}>
            <div><label style={s.lbl}>Title / Role</label><input value={inspTitle} onChange={e=>setInspTitle(e.target.value)} placeholder="e.g. Safety Inspector" style={s.inp}/></div>
            <div><label style={s.lbl}>Company</label><input value={inspCo} onChange={e=>setInspCo(e.target.value)} placeholder="e.g. Colonial Pipeline" style={s.inp}/></div>
          </div>

          <div style={{marginBottom:16}}>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:6}}><label style={s.lbl}>Signature *</label>{sigData&&<button onClick={clearSig} style={{background:"none",border:"none",color:"#FC8181",fontSize:12,cursor:"pointer"}}>Clear</button>}</div>
            <div style={{background:"#fff",borderRadius:10,overflow:"hidden",border:"2px solid #60A5FA"}}>
              <canvas ref={sigRef} width={600} height={160} style={{width:"100%",height:160,display:"block",touchAction:"none",cursor:"crosshair"}}
                onMouseDown={startSig} onMouseMove={drawSig} onMouseUp={endSig} onMouseLeave={endSig}
                onTouchStart={startSig} onTouchMove={drawSig} onTouchEnd={endSig}/>
            </div>
            {!sigData&&<div style={{textAlign:"center",fontSize:11,color:"#7080A0",marginTop:4}}>Draw your signature above</div>}
            {sigData&&<div style={{textAlign:"center",fontSize:11,color:"#34D399",marginTop:4}}>✓ Signature captured</div>}
          </div>

          <button onClick={submit} disabled={saving||!inspName.trim()||!sigData}
            style={{width:"100%",background:saving||!inspName.trim()||!sigData?"#26262E":"#34D399",color:saving||!inspName.trim()||!sigData?"#7080A0":"#000",border:"none",borderRadius:12,padding:14,fontSize:16,fontWeight:800,cursor:"pointer",fontFamily:"inherit"}}>
            {saving?"Saving…":"✅ Sign Daily Report"}
          </button>
          <div style={{fontSize:11,color:"#555",textAlign:"center",marginTop:8}}>By signing you confirm the work described above was performed.</div>
        </div>
      </div>
    </div>
  );
}

function PublicTMSignForm({ticketId}){
  const [ticket,setTicket]=useState(null);const [project,setProject]=useState(null);
  const [loading,setLoading]=useState(true);const [saving,setSaving]=useState(false);
  const [submitted,setSubmitted]=useState(false);const [err,setErr]=useState("");
  const [cName,setCName]=useState("");const [cTitle,setCTitle]=useState("");
  const [cCo,setCCo]=useState("");const [sigData,setSigData]=useState(null);
  const [drawing,setDrawing]=useState(false);
  const sigRef=React.useRef(null);

  useEffect(()=>{(async()=>{try{
    const d=await sb(`/tm_tickets?id=eq.${ticketId}&limit=1`);
    const t=Array.isArray(d)?d[0]:d;
    if(!t){setErr("Ticket not found.");setLoading(false);return;}
    setTicket(t);
    if(t.client_signature){setSubmitted(true);}
    const pd=await sb(`/projects?id=eq.${t.project_id}&limit=1`);
    setProject(Array.isArray(pd)?pd[0]:pd||{name:"Unknown"});
  }catch(e){setErr(e.message);}setLoading(false);})();},[ticketId]);

  function getPos(e,c){const r=c.getBoundingClientRect();const sx=c.width/r.width;const sy=c.height/r.height;if(e.touches)return{x:(e.touches[0].clientX-r.left)*sx,y:(e.touches[0].clientY-r.top)*sy};return{x:(e.clientX-r.left)*sx,y:(e.clientY-r.top)*sy};}
  function startSig(e){e.preventDefault();const c=sigRef.current;if(!c)return;const p=getPos(e,c);c.getContext("2d").beginPath();c.getContext("2d").moveTo(p.x,p.y);setDrawing(true);}
  function drawSig(e){e.preventDefault();if(!drawing)return;const c=sigRef.current;if(!c)return;const ctx=c.getContext("2d");const p=getPos(e,c);ctx.lineWidth=2.5;ctx.lineCap="round";ctx.strokeStyle="#1f3864";ctx.lineTo(p.x,p.y);ctx.stroke();ctx.beginPath();ctx.moveTo(p.x,p.y);}
  function endSig(e){e.preventDefault();setDrawing(false);const c=sigRef.current;if(c)setSigData(c.toDataURL("image/jpeg",0.7));}
  function clearSig(){const c=sigRef.current;if(c)c.getContext("2d").clearRect(0,0,c.width,c.height);setSigData(null);}

  async function submit(){
    if(!cName.trim()||!sigData){setErr("Please enter your name and sign.");return;}
    setSaving(true);setErr("");
    try{
      const fullName=cName.trim()+(cTitle?` · ${cTitle}`:"")+(cCo?` · ${cCo}`:"");
      await sb(`/tm_tickets?id=eq.${ticketId}`,{method:"PATCH",body:{
        client_contact:fullName,
        client_signature:sigData,
        client_signed_at:new Date().toISOString(),
      },prefer:"return=representation"});
      setSubmitted(true);
    }catch(e){setErr("Failed: "+e.message);}setSaving(false);
  }

  const money=(n)=>"$"+Number(n||0).toLocaleString("en-US",{minimumFractionDigits:2,maximumFractionDigits:2});
  const s={bg:"#0D0D0F",card:"#1A1A20",border:"#26262E",inp:{width:"100%",background:"#141418",border:"1px solid #26262E",borderRadius:10,color:"#F0F4FF",fontSize:14,padding:"12px 14px",fontFamily:"inherit",outline:"none"},lbl:{display:"block",fontSize:11,fontWeight:700,color:"#7080A0",textTransform:"uppercase",letterSpacing:"1px",marginBottom:6}};

  if(loading)return<div style={{background:s.bg,minHeight:"100vh",display:"flex",alignItems:"center",justifyContent:"center",fontFamily:"system-ui"}}><div style={{color:"#F0F4FF"}}>Loading ticket...</div></div>;
  if(err&&!ticket)return<div style={{background:s.bg,minHeight:"100vh",display:"flex",alignItems:"center",justifyContent:"center",fontFamily:"system-ui",padding:20}}><div style={{background:s.card,borderRadius:16,padding:32,maxWidth:400,textAlign:"center"}}><div style={{fontSize:40}}>⚠️</div><div style={{color:"#F0F4FF",fontWeight:700,margin:"8px 0"}}>{err}</div></div></div>;
  if(submitted)return(
    <div style={{background:s.bg,minHeight:"100vh",fontFamily:"system-ui",padding:20}}>
      <div style={{maxWidth:560,margin:"0 auto",background:s.card,borderRadius:16,padding:32,textAlign:"center",border:"1px solid #34D39940"}}>
        <div style={{fontSize:48,marginBottom:12}}>✅</div>
        <div style={{fontSize:20,fontWeight:800,color:"#34D399",marginBottom:8}}>Ticket Signed!</div>
        <div style={{color:"#7080A0",marginBottom:16}}>T&M Ticket {ticket?.ticket_no} has been signed and saved.</div>
        {ticket?.client_signature&&<div style={{background:"#fff",borderRadius:10,padding:8,marginTop:8}}><img src={ticket.client_signature} alt="sig" style={{maxHeight:80,width:"100%",objectFit:"contain"}}/></div>}
        <div style={{fontSize:12,color:"#555",marginTop:12}}>You may close this window.</div>
      </div>
    </div>
  );

  const sec=(title,rows)=>rows&&rows.length?(
    <div style={{marginTop:10,paddingTop:10,borderTop:"1px solid #26262E"}}>
      <div style={{fontSize:10,color:"#7080A0",fontWeight:700,textTransform:"uppercase",marginBottom:6}}>{title}</div>
      {rows.map((r,i)=>(
        <div key={i} style={{display:"flex",justifyContent:"space-between",gap:10,fontSize:12,color:"#C8D4F0",padding:"3px 0"}}>
          <span>{r.label}</span><span style={{fontWeight:700,whiteSpace:"nowrap"}}>{r.amt}</span>
        </div>
      ))}
    </div>
  ):null;

  return(
    <div style={{background:s.bg,minHeight:"100vh",fontFamily:"system-ui",color:"#F0F4FF",padding:"16px 16px 60px"}}>
      <div style={{maxWidth:560,margin:"0 auto"}}>
        <div style={{background:"#141418",borderRadius:16,padding:"16px 20px",marginBottom:16,border:"1px solid #26262E"}}>
          <div style={{fontSize:20,fontWeight:900,color:"#60A5FA",marginBottom:2}}>AIME</div>
          <div style={{fontSize:11,color:"#7080A0"}}>Atlantic Industrial Mechanical & Environmental Inc.</div>
        </div>
        <div style={{background:s.card,borderRadius:12,padding:"14px 16px",marginBottom:14,border:"1px solid #26262E"}}>
          <div style={{fontSize:11,fontWeight:700,color:"#7080A0",textTransform:"uppercase",letterSpacing:"1px",marginBottom:10}}>Time &amp; Materials Ticket — Client Approval</div>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}>
            {[["Project",project?.name||"—"],["T&M #",ticket?.ticket_no||"—"],["Date",ticket?.ticket_date||"—"],["Submitted By",ticket?.submitted_by||"—"]].map(([l,v])=>(
              <div key={l}><div style={{fontSize:10,color:"#7080A0",fontWeight:700,textTransform:"uppercase"}}>{l}</div><div style={{fontSize:13,fontWeight:600,color:"#F0F4FF",marginTop:2}}>{v}</div></div>
            ))}
          </div>
          {ticket?.description&&<div style={{marginTop:10,paddingTop:10,borderTop:"1px solid #26262E"}}>
            <div style={{fontSize:10,color:"#7080A0",fontWeight:700,textTransform:"uppercase",marginBottom:4}}>Work Description</div>
            <div style={{fontSize:12,color:"#C8D4F0",lineHeight:1.5}}>{ticket.description}</div>
          </div>}

          {sec("Labor",(ticket?.labor||[]).map(r=>({label:`${r.name||"—"} · ${r.classification||""} · ${r.hours||0} hr`,amt:money((parseFloat(r.hours)||0)*(parseFloat(r.rate)||0))})))}
          {sec("Equipment",(ticket?.equipment||[]).map(r=>({label:`${r.description||"—"} · ${r.qty||0} ${r.unit||""}`,amt:money((parseFloat(r.qty)||0)*(parseFloat(r.rate)||0))})))}
          {sec("Materials",(ticket?.materials||[]).map(r=>({label:`${r.description||"—"} · ${r.qty||0}`,amt:money((parseFloat(r.qty)||0)*(parseFloat(r.unit_price)||0))})))}
          {sec("Other",(ticket?.other_charges||[]).map(r=>({label:r.description||"—",amt:money(r.amount)})))}

          <div style={{marginTop:12,paddingTop:10,borderTop:"1px solid #26262E",display:"flex",justifyContent:"space-between",alignItems:"center"}}>
            <span style={{fontSize:13,fontWeight:800,color:"#F0F4FF"}}>Grand Total</span>
            <span style={{fontSize:20,fontWeight:900,color:"#34D399"}}>{money(ticket?.grand_total)}</span>
          </div>
        </div>

        {err&&<div style={{background:"#FC818120",border:"1px solid #FC8181",borderRadius:10,padding:"10px 14px",marginBottom:12,fontSize:13,color:"#FC8181"}}>{err}</div>}

        <div style={{background:s.card,borderRadius:16,padding:20,border:"1px solid #26262E"}}>
          <div style={{fontSize:14,fontWeight:800,color:"#60A5FA",marginBottom:14}}>✍️ Client Approval</div>
          <div style={{marginBottom:10}}><label style={s.lbl}>Your Name *</label><input value={cName} onChange={e=>setCName(e.target.value)} placeholder="Full name" style={s.inp}/></div>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8,marginBottom:14}}>
            <div><label style={s.lbl}>Title / Role</label><input value={cTitle} onChange={e=>setCTitle(e.target.value)} placeholder="e.g. Project Manager" style={s.inp}/></div>
            <div><label style={s.lbl}>Company</label><input value={cCo} onChange={e=>setCCo(e.target.value)} placeholder="e.g. Colonial Pipeline" style={s.inp}/></div>
          </div>
          <div style={{marginBottom:16}}>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:6}}><label style={s.lbl}>Signature *</label>{sigData&&<button onClick={clearSig} style={{background:"none",border:"none",color:"#FC8181",fontSize:12,cursor:"pointer"}}>Clear</button>}</div>
            <div style={{background:"#fff",borderRadius:10,overflow:"hidden",border:"2px solid #60A5FA"}}>
              <canvas ref={sigRef} width={600} height={160} style={{width:"100%",height:160,display:"block",touchAction:"none",cursor:"crosshair"}}
                onMouseDown={startSig} onMouseMove={drawSig} onMouseUp={endSig} onMouseLeave={endSig}
                onTouchStart={startSig} onTouchMove={drawSig} onTouchEnd={endSig}/>
            </div>
            {!sigData&&<div style={{textAlign:"center",fontSize:11,color:"#7080A0",marginTop:4}}>Draw your signature above</div>}
            {sigData&&<div style={{textAlign:"center",fontSize:11,color:"#34D399",marginTop:4}}>✓ Signature captured</div>}
          </div>
          <button onClick={submit} disabled={saving||!cName.trim()||!sigData}
            style={{width:"100%",background:saving||!cName.trim()||!sigData?"#26262E":"#34D399",color:saving||!cName.trim()||!sigData?"#7080A0":"#000",border:"none",borderRadius:12,padding:14,fontSize:16,fontWeight:800,cursor:"pointer",fontFamily:"inherit"}}>
            {saving?"Saving…":"✅ Approve T&M Ticket"}
          </button>
          <div style={{fontSize:11,color:"#555",textAlign:"center",marginTop:8}}>By signing you confirm the labor, equipment, and materials itemized above were furnished as described and are authorized for invoicing.</div>
        </div>
      </div>
    </div>
  );
}

const MFG_STAGES=[
  {id:"mat_received",  label:"Material Received",   icon:"📦", color:"#60A5FA", desc:"Log GFM material receipt"},
  {id:"mat_inspection",label:"Material Inspection", icon:"🔍", color:"#FBBF24", desc:"Inspect received materials"},
  {id:"tacked",        label:"Tacked in Jig",       icon:"🔩", color:"#F97316", desc:"All parts tacked together"},
  {id:"welded",        label:"Fully Welded",         icon:"🔥", color:"#EF4444", desc:"Complete weld-out"},
  {id:"welder_qc",     label:"Welder QC",            icon:"🟡", color:"#FCD34D", desc:"Welder self-inspection — Yellow Ribbon"},
  {id:"manager_qc",    label:"Manager QC",           icon:"⚪", color:"#F0F4FF", desc:"Manager sign-off — White Ribbon"},
  {id:"shipped",       label:"Palletized & Shipped", icon:"📦", color:"#34D399", desc:"Banded and ready to go"},
];

function ManufacturingJobBoard({user,onBack,onSelectJob}){
  const [jobs,setJobs]=useState([]);
  const [loading,setLoading]=useState(true);
  const [showNew,setShowNew]=useState(false);
  const [f,setF]=useState({job_number:"",customer:"",description:"",po_number:"",due_date:"",notes:""});
  const [saving,setSaving]=useState(false);
  const [boardTab,setBoardTab]=useState("jobs");
  const canAdmin=user.role==="admin"||user.role==="pm";
  const set=(k,v)=>setF(x=>({...x,[k]:v}));

  useEffect(()=>{load();},[]);
  async function load(){
    setLoading(true);
    try{const j=await API.mfg.jobs.list();setJobs(Array.isArray(j)?j:[]);}
    catch(e){console.error("MFG error:",e.message||e);}
    setLoading(false);
  }
  const [formErr,setFormErr]=useState("");
  async function createJob(){
    if(!f.job_number.trim())return;
    setSaving(true);setFormErr("");
    try{
      const payload={...f,created_by:user.name};
      if(!payload.due_date)payload.due_date=null;
      if(!payload.po_number)payload.po_number=null;
      await API.mfg.jobs.create(payload);
      setShowNew(false);
      setF({job_number:"",customer:"",description:"",po_number:"",due_date:"",notes:""});
      await load();
    }catch(e){
      const msg=e.message||String(e);
      if(msg.includes("mfg_jobs")||msg.includes("relation")||msg.includes("exist")){
        setFormErr("⚠️ Database table missing — please run AIME_manufacturing.sql in Supabase SQL Editor first, then try again.");
      }else{
        setFormErr("Error: "+msg);
      }
    }
    setSaving(false);
  }
  async function archiveJob(id,status){
    if(!window.confirm(status==="active"?"Put job on hold?":"Reactivate job?"))return;
    try{await API.mfg.jobs.update(id,{status:status==="active"?"on_hold":"active"});await load();}catch(e){console.error("MFG error:",e.message||e);}
  }
  async function deleteJob(id){
    if(!window.confirm("Permanently delete this job and ALL its data? This cannot be undone."))return;
    try{await API.mfg.jobs.remove(id);await load();}catch(e){alert("Error: "+e.message);}
  }

  const active=jobs.filter(j=>j.status==="active");
  const held=jobs.filter(j=>j.status==="on_hold");

  return(
    <div style={{background:T.bg,minHeight:"100vh",fontFamily:"inherit"}}>
      <div style={{background:T.surface,borderBottom:`1px solid ${T.border}`,padding:"14px 16px",display:"flex",justifyContent:"space-between",alignItems:"center",position:"sticky",top:0,zIndex:50}}>
        <button onClick={onBack} style={{background:"none",border:"none",color:T.sub,fontSize:13,cursor:"pointer",fontFamily:"inherit"}}>← Back</button>
        <div style={{display:"flex",alignItems:"center",gap:8}}>
          <span style={{fontSize:20}}>🏭</span>
          <div style={{fontSize:15,fontWeight:900,color:T.purple}}>Manufacturing</div>
        </div>
        {canAdmin&&<button onClick={()=>setShowNew(s=>!s)} style={{background:T.purple,color:"#fff",border:"none",borderRadius:10,padding:"7px 14px",fontSize:13,fontWeight:700,cursor:"pointer",fontFamily:"inherit"}}>
          + New Job
        </button>}
      </div>

      <div style={{padding:"14px 16px 80px"}}>
        {/* New job form */}
        {showNew&&<div style={{...cardS,marginBottom:16,border:`1px solid ${T.purple}40`}}>
          <div style={{fontSize:14,fontWeight:800,color:T.purple,marginBottom:14}}>🏭 New Manufacturing Job</div>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:10}}>
            <div><label style={lbl}>Job / WO # *</label><input value={f.job_number} onChange={e=>set("job_number",e.target.value)} placeholder="MFG-2026-001" style={inp}/></div>
            <div><label style={lbl}>Customer</label><input value={f.customer} onChange={e=>set("customer",e.target.value)} placeholder="Customer name" style={inp}/></div>
          </div>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:10}}>
            <div><label style={lbl}>PO Number</label><input value={f.po_number} onChange={e=>set("po_number",e.target.value)} placeholder="PO #" style={inp}/></div>
            <div><label style={lbl}>Due Date</label><input type="date" value={f.due_date} onChange={e=>set("due_date",e.target.value)} style={inp}/></div>
          </div>
          <div style={{marginBottom:10}}><label style={lbl}>Description</label><input value={f.description} onChange={e=>set("description",e.target.value)} placeholder="What are we making?" style={inp}/></div>
          <div style={{marginBottom:14}}><label style={lbl}>Notes</label><textarea value={f.notes} onChange={e=>set("notes",e.target.value)} rows={2} style={{...inp,resize:"vertical"}}/></div>
          <div style={{display:"flex",gap:8}}>
            {formErr&&<div style={{background:T.redLow,border:`1px solid ${T.red}40`,borderRadius:10,padding:"10px 12px",marginBottom:8,fontSize:12,color:T.red,lineHeight:1.6}}>{formErr}</div>}
            <button onClick={createJob} disabled={!f.job_number.trim()||saving} style={{...primBtn,flex:2,borderRadius:12,background:T.purple,opacity:f.job_number.trim()&&!saving?1:0.5}}>{saving?"Creating…":"Create Job"}</button>
            <button onClick={()=>setShowNew(false)} style={{...ghostBtn,flex:1,textAlign:"center"}}>Cancel</button>
          </div>
        </div>}

        {loading&&<Spinner/>}

        {/* Summary stats */}
        {!loading&&jobs.length>0&&<div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8,marginBottom:16}}>
          {[[active.length,"Active Jobs",T.purple],[held.length,"On Hold",T.muted]].map(([v,l,c])=>(
            <div key={l} style={{...cardS,textAlign:"center"}}><div style={{fontSize:24,fontWeight:900,color:c}}>{v}</div><div style={{fontSize:10,color:T.muted,textTransform:"uppercase",marginTop:2}}>{l}</div></div>
          ))}
        </div>}

        {/* Dashboard tab toggle */}
        {!loading&&jobs.length>0&&<div style={{display:"flex",background:T.surface,borderRadius:12,padding:4,marginBottom:14,gap:4}}>
          {[["jobs","🔩 Jobs"],["dashboard","📊 Dashboard"]].map(([id,label])=>(
            <button key={id} onClick={()=>setBoardTab(id)}
              style={{flex:1,padding:"8px",background:boardTab===id?T.purple:"none",color:boardTab===id?"#fff":T.muted,border:"none",borderRadius:10,fontSize:13,fontWeight:700,cursor:"pointer",fontFamily:"inherit",transition:"all 0.15s"}}>
              {label}
            </button>
          ))}
        </div>}

        {boardTab==="dashboard"&&jobs.length>0&&<ManufacturingDashboard jobs={jobs} user={user} onSelectJob={j=>onSelectJob(j)}/>}
        {boardTab==="jobs"&&active.length===0&&!loading&&<div style={{textAlign:"center",padding:"40px 16px",color:T.muted}}>
          <div style={{fontSize:48,marginBottom:12}}>🏭</div>
          <div style={{fontSize:15,fontWeight:700,color:T.sub,marginBottom:6}}>No Manufacturing Jobs</div>
          <div style={{fontSize:12}}>Tap <strong style={{color:T.purple}}>+ New Job</strong> to create your first production job.</div>
        </div>}

        {boardTab==="jobs"&&active.map(job=><MfgJobCard key={job.id} job={job} onSelect={()=>onSelectJob(job)} onArchive={()=>archiveJob(job.id,job.status)} onDelete={()=>deleteJob(job.id)}/>)}
        {boardTab==="jobs"&&held.length>0&&<><div style={{fontSize:11,fontWeight:700,color:T.muted,textTransform:"uppercase",letterSpacing:"1px",margin:"14px 0 8px"}}>On Hold</div>
        {held.map(job=><MfgJobCard key={job.id} job={job} onSelect={()=>onSelectJob(job)} onArchive={()=>archiveJob(job.id,job.status)} onDelete={()=>deleteJob(job.id)} dimmed/>)}</>}
      </div>
    </div>
  );
}

function MfgJobCard({job,onSelect,onArchive,onDelete,dimmed}){
  const daysLeft=job.due_date?Math.ceil((new Date(job.due_date+"T12:00:00")-new Date())/86400000):null;
  const dueCo=daysLeft===null?T.muted:daysLeft<0?T.red:daysLeft<=7?T.yellow:T.green;
  return(
    <div onClick={onSelect} style={{...cardS,marginBottom:10,cursor:"pointer",opacity:dimmed?0.6:1,borderLeft:`3px solid ${T.purple}`}}>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:8}}>
        <div>
          <div style={{fontSize:16,fontWeight:900,color:T.purple}}>{job.job_number}</div>
          <div style={{fontSize:12,color:T.sub}}>{job.customer||"No customer"}{job.po_number?" · PO: "+job.po_number:""}</div>
          {job.description&&<div style={{fontSize:12,color:T.muted,marginTop:2}}>{job.description}</div>}
        </div>
        <div style={{textAlign:"right",flexShrink:0}}>
          {daysLeft!==null&&<div style={{fontSize:11,fontWeight:700,color:dueCo}}>{daysLeft<0?`${Math.abs(daysLeft)}d OVERDUE`:daysLeft===0?"Due TODAY":`${daysLeft}d left`}</div>}
          <div style={{fontSize:9,color:T.muted,marginTop:2}}>Due: {job.due_date||"—"}</div>
        </div>
      </div>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}>
        <div style={{display:"flex",gap:6}}>
          <span style={{...pill(T.purple),fontSize:10}}>{job.status==="active"?"● Active":"⏸ On Hold"}</span>
          {job.due_date&&daysLeft<0&&<span style={{...pill(T.red),fontSize:10}}>⚠️ OVERDUE</span>}
        </div>
        <div style={{display:"flex",gap:6,alignItems:"center"}}>
          <button onClick={e=>{e.stopPropagation();onArchive&&onArchive();}} style={{background:"none",border:`1px solid ${T.border}`,borderRadius:8,padding:"4px 8px",color:T.muted,fontSize:10,cursor:"pointer",fontFamily:"inherit"}}>
            {job.status==="active"?"⏸":"▶"}
          </button>
          <button onClick={e=>{e.stopPropagation();onDelete&&onDelete();}} style={{background:"none",border:`1px solid ${T.red}40`,borderRadius:8,padding:"4px 8px",color:T.red,fontSize:10,cursor:"pointer",fontFamily:"inherit"}}>🗑</button>
          <span style={{color:T.orange,fontSize:16,fontWeight:700}}>→</span>
        </div>
      </div>
    </div>
  );
}

function ManufacturingJobDetail({job,user,onBack,onSelectPart}){
  const [parts,setParts]=useState([]);
  const [boms,setBoms]=useState({});      // keyed by part_id → array of bom items
  const [receipts,setReceipts]=useState([]); // flat list of all receipts for this job
  const [assemblyLogs,setAssemblyLogs]=useState([]);
  const [shippingLogs,setShippingLogs]=useState([]);
  const [loading,setLoading]=useState(true);
  const [tab,setTab]=useState("overview"); // overview | received | assembly | shipping
  const canAdmin=user.role==="admin"||user.role==="pm"||user.role==="foreman";

  const [selectedBomItem,setSelectedBomItem]=useState(null);
  const [rQty,setRQty]=useState("");
  const [rDate,setRDate]=useState(today());
  const [rJob,setRJob]=useState(job.job_number||"");
  const [rBol,setRBol]=useState("");
  const [rBy,setRBy]=useState(user.name);
  const [saving,setSaving]=useState(false);
  const [formErr,setFormErr]=useState("");
  const [showPackingSlip,setShowPackingSlip]=useState(false);
  const [editingSlip,setEditingSlip]=useState(null);
  const [packingSlips,setPackingSlips]=useState([]);
  const [showNewPart,setShowNewPart]=useState(false);
  const [showAddComp,setShowAddComp]=useState(false);
  const [pf,setPf]=useState({part_number:"",description:"",drawing_number:"",qty_ordered:""});
  const [newCompPartId,setNewCompPartId]=useState("");
  const [newCompNum,setNewCompNum]=useState("");
  const [newCompDesc,setNewCompDesc]=useState("");
  const [newCompQpa,setNewCompQpa]=useState("1");
  const [newCompReorder,setNewCompReorder]=useState("");

  const [showAsmForm,setShowAsmForm]=useState(false);
  const [af,setAf]=useState({part_id:"",qty:"",date:today(),by:user.name});

  const [showShipForm,setShowShipForm]=useState(false);
  const [sf,setSf]=useState({part_id:"",qty:"",date:today(),customer:job.customer||"",bol:"",by:user.name});

  useEffect(()=>{load();},[job.id]);

  async function load(){
    setLoading(true);
    try{
      const ps=await API.mfg.parts.forJob(job.id).catch(()=>[]);
      const partList=Array.isArray(ps)?ps:[];
      setParts(partList);
      const bomMap={};
      const allReceipts=[];
      await Promise.all(partList.map(async p=>{
        const [b,r]=await Promise.all([
          API.mfg.bom.forPart(p.id).catch(()=>[]),
          API.mfg.receipts.forPart(p.id).catch(()=>[]),
        ]);
        bomMap[p.id]=Array.isArray(b)?b:[];
        (Array.isArray(r)?r:[]).forEach(rx=>allReceipts.push({...rx,_partId:p.id,_partNum:p.part_number}));
      }));
      setBoms(bomMap);
      setReceipts(allReceipts.sort((a,b)=>(b.received_date||"").localeCompare(a.received_date||"")));
      const [aLog,sLog,pSlips]=await Promise.all([
        API.mfg.assemblyLog.forJob(job.id).catch(()=>[]),
        API.mfg.shippingLog.forJob(job.id).catch(()=>[]),
        API.mfg.packingSlips.forJob(job.id).catch(()=>[]),
      ]);
      setPackingSlips(Array.isArray(pSlips)?pSlips:[]);
      setAssemblyLogs(Array.isArray(aLog)?aLog:[]);
      setShippingLogs(Array.isArray(sLog)?sLog:[]);
    }catch(e){console.error(e);}
    setLoading(false);
  }

  function inv(item){
    const itxns=receipts.filter(r=>r.bom_id===item.id);
    const received=itxns.reduce((s,r)=>s+(r.qty_received||0),0);
    const damaged=itxns.reduce((s,r)=>s+(r.qty_damaged||0),0);
    const usedInAsm=itxns.filter(r=>r.transaction_type==="Issued").reduce((s,r)=>s+(r.qty_received||0),0);
    const onHand=Math.max(0,received-damaged-usedInAsm);
    const qpa=item.qty_per_assembly||1;
    return{received,damaged,usedInAsm,onHand,canBuild:Math.floor(onHand/qpa),qpa,reorderLevel:item.reorder_level||0,needsReorder:(item.reorder_level||0)>0&&onHand<=(item.reorder_level||0)};
  }

  function canBuildPart(partId){
    const bom=boms[partId]||[];
    if(!bom.length)return 0;
    return Math.min(...bom.map(item=>inv(item).canBuild));
  }

  function asmTotals(partId){
    const done=assemblyLogs.filter(a=>a.part_id===partId).reduce((s,a)=>s+(a.qty_completed||0),0);
    const shipped=shippingLogs.filter(s=>s.part_id===partId).reduce((s,a)=>s+(a.qty_shipped||0),0);
    return{done,shipped,readyToShip:done-shipped};
  }

  async function createPart(){
    if(!pf.part_number.trim())return;
    setSaving(true);
    try{
      await API.mfg.parts.create({...pf,job_id:job.id,qty_ordered:parseInt(pf.qty_ordered)||0,drawing_number:pf.drawing_number||null});
      setShowNewPart(false);setPf({part_number:"",description:"",drawing_number:"",qty_ordered:""});
      await load();
    }catch(e){alert("Error: "+e.message);}
    setSaving(false);
  }
  async function addCompPart(){
    if(!newCompNum.trim()||!newCompDesc.trim()||!newCompPartId)return;
    setSaving(true);
    try{
      await API.mfg.bom.create({part_id:newCompPartId,component_part_number:newCompNum.trim(),material:newCompDesc.trim(),qty_per_assembly:parseFloat(newCompQpa)||1,reorder_level:parseFloat(newCompReorder)||0,unit:"ea"});
      setNewCompNum("");setNewCompDesc("");setNewCompQpa("1");setNewCompReorder("");setShowAddComp(false);
      await load();
    }catch(e){alert("Error: "+e.message);}
    setSaving(false);
  }
  async function deletePart(id){
    if(!window.confirm("Delete this assembly and all its data?"))return;
    try{await API.mfg.parts.remove(id);await load();}catch(e){}
  }

  async function logReceipt(){
    if(!selectedBomItem||!rQty){setFormErr("Please select a part and enter a quantity.");return;}
    setSaving(true);setFormErr("");
    try{
      await API.mfg.receipts.create({
        bom_id:selectedBomItem.id,
        part_id:selectedBomItem.part_id,
        qty_received:parseFloat(rQty)||0,
        qty_issued:0,qty_damaged:0,
        transaction_type:"Received",
        received_date:rDate,
        received_by:rBy||user.name,
        reference_bol:rBol||null,
        job_number:rJob||null,
      });
      setSelectedBomItem(null);setRQty("");setRBol("");
      await load();
    }catch(e){setFormErr("Error: "+e.message);}
    setSaving(false);
  }

  async function logAssembly(){
    if(!af.qty||!af.part_id)return;
    setSaving(true);
    try{
      await API.mfg.assemblyLog.create({part_id:af.part_id,job_id:job.id,qty_completed:parseInt(af.qty),completion_date:af.date,entered_by:af.by});
      const bom=boms[af.part_id]||[];
      await Promise.all(bom.map(item=>API.mfg.receipts.create({bom_id:item.id,part_id:af.part_id,qty_received:parseInt(af.qty)*(item.qty_per_assembly||1),qty_issued:0,qty_damaged:0,transaction_type:"Issued",received_date:af.date,received_by:af.by,job_number:job.job_number||null,notes:"Auto: assemblies completed"}).catch(()=>{})));
      setShowAsmForm(false);setAf({part_id:"",qty:"",date:today(),by:user.name});
      await load();
    }catch(e){alert(e.message);}
    setSaving(false);
  }

  async function logShipment(){
    if(!sf.qty||!sf.part_id)return;
    setSaving(true);
    try{
      await API.mfg.shippingLog.create({part_id:sf.part_id,job_id:job.id,qty_shipped:parseInt(sf.qty),ship_date:sf.date,customer:sf.customer||null,bol_number:sf.bol||null,entered_by:sf.by});
      setShowShipForm(false);setSf({part_id:"",qty:"",date:today(),customer:job.customer||"",bol:"",by:user.name});
      await load();
    }catch(e){alert(e.message);}
    setSaving(false);
  }

  const allBomItems=parts.flatMap(p=>(boms[p.id]||[]).map(item=>({...item,part_id:p.id,_partNum:p.part_number})));
  const totalCanBuild=parts.length>0?parts.reduce((mn,p)=>Math.min(mn,canBuildPart(p.id)),9999):0;
  const totalReadyToShip=parts.reduce((s,p)=>s+asmTotals(p.id).readyToShip,0);
  const totalShipped=parts.reduce((s,p)=>s+asmTotals(p.id).shipped,0);
  const reorderNeeded=allBomItems.filter(item=>inv(item).needsReorder);

  if(showPackingSlip) return(
    <PackingSlipScreen job={job} parts={parts} user={user} existingSlip={editingSlip} onBack={()=>{setShowPackingSlip(false);setEditingSlip(null);}} onSaved={async()=>{await load();}}/>
  );

  return(
    <div style={{background:T.bg,minHeight:"100vh",fontFamily:"inherit"}}>
      {/* Header */}
      <div style={{background:T.surface,borderBottom:`1px solid ${T.border}`,padding:"12px 16px",position:"sticky",top:0,zIndex:50}}>
        <button onClick={onBack} style={{background:"none",border:"none",color:T.sub,fontSize:13,cursor:"pointer",fontFamily:"inherit",display:"block",marginBottom:4}}>← Back to Jobs</button>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}>
          <div><div style={{fontSize:18,fontWeight:900,color:T.purple}}>{job.job_number}</div>
            <div style={{fontSize:12,color:T.muted}}>{job.customer}{job.po_number?" · PO: "+job.po_number:""}</div>
          </div>
          {job.due_date&&<div style={{fontSize:11,fontWeight:700,color:T.muted}}>Due: {job.due_date}</div>}
        </div>
      </div>

      {/* Reorder alert */}
      {reorderNeeded.length>0&&<div style={{background:"#7c2d12",padding:"8px 16px",display:"flex",gap:8,alignItems:"center"}}>
        <span>🔴</span>
        <div style={{fontSize:11,color:"#fed7aa",fontWeight:700}}>REORDER: {reorderNeeded.map(i=>i.component_part_number||i.material).join(" · ")}</div>
      </div>}

      {/* Tabs */}
      <div style={{display:"flex",background:T.surface,borderBottom:`1px solid ${T.border}`}}>
        {[["overview","📊 Overview"],["received","📦 Received Parts"],["assembly","🏭 Assembly Log"],["shipping","📤 Shipping Log"]].map(([id,label])=>(
          <button key={id} onClick={()=>setTab(id)} style={{flex:1,padding:"12px 4px",background:"none",border:"none",borderBottom:`3px solid ${tab===id?T.purple:"transparent"}`,color:tab===id?T.purple:T.muted,fontSize:11,fontWeight:700,cursor:"pointer",fontFamily:"inherit",whiteSpace:"nowrap"}}>
            {label}
          </button>
        ))}
      </div>

      <div style={{padding:"14px 16px 100px"}}>
        {loading&&<Spinner/>}

        {}
        {!loading&&tab==="overview"&&<>
          {/* 3 key KPIs */}
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:10,marginBottom:16}}>
            <div style={{...cardS,textAlign:"center",padding:"16px 8px",border:`2px solid ${totalCanBuild>0?T.green:T.red}40`}}>
              <div style={{fontSize:11,color:T.muted,textTransform:"uppercase",marginBottom:4}}>Can Build</div>
              <div style={{fontSize:32,fontWeight:900,color:totalCanBuild>0?T.green:T.red,lineHeight:1}}>{totalCanBuild}</div>
              <div style={{fontSize:10,color:T.muted,marginTop:4}}>assemblies</div>
            </div>
            <div style={{...cardS,textAlign:"center",padding:"16px 8px"}}>
              <div style={{fontSize:11,color:T.muted,textTransform:"uppercase",marginBottom:4}}>Ready to Ship</div>
              <div style={{fontSize:32,fontWeight:900,color:totalReadyToShip>0?T.orange:T.muted,lineHeight:1}}>{totalReadyToShip}</div>
              <div style={{fontSize:10,color:T.muted,marginTop:4}}>completed</div>
            </div>
            <div style={{...cardS,textAlign:"center",padding:"16px 8px"}}>
              <div style={{fontSize:11,color:T.muted,textTransform:"uppercase",marginBottom:4}}>Shipped</div>
              <div style={{fontSize:32,fontWeight:900,color:T.blue,lineHeight:1}}>{totalShipped}</div>
              <div style={{fontSize:10,color:T.muted,marginTop:4}}>total</div>
            </div>
          </div>

          {/* Parts in stock — one card per component */}
          {/* Assembly parts with delete */}
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:10}}>
            <div style={{fontSize:12,fontWeight:800,color:T.text}}>Parts in Stock</div>
            {canAdmin&&parts.map(p=>(
              <div key={p.id} style={{display:"flex",alignItems:"center",gap:6}}>
                <span style={{fontSize:11,color:T.purple,fontWeight:700}}>{p.part_number}</span>
                <button onClick={()=>deletePart(p.id)} style={{background:"none",border:`1px solid ${T.red}30`,borderRadius:6,padding:"2px 8px",color:T.red,fontSize:10,cursor:"pointer",fontFamily:"inherit"}}>🗑 Delete Job Part</button>
              </div>
            ))}
          </div>
          {allBomItems.length===0&&<div style={{...cardS,textAlign:"center",padding:30,color:T.muted}}>
            <div style={{fontSize:32,marginBottom:8}}>📦</div>
            <div style={{fontWeight:700,color:T.sub,marginBottom:4}}>No component parts added yet</div>
            <div style={{fontSize:12}}>Go to the <strong style={{color:T.purple}}>Received Parts</strong> tab to add component parts and log inventory.</div>
          </div>}

          {allBomItems.map(item=>{
            const i=inv(item);
            const pct=i.received>0?Math.min(100,(i.onHand/i.received)*100):0;
            return(
              <div key={item.id} style={{...cardS,marginBottom:8,borderLeft:`4px solid ${i.needsReorder?T.red:i.onHand<=0?T.yellow:T.green}`}}>
                <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:8}}>
                  <div>
                    <div style={{fontSize:14,fontWeight:800,color:T.orange}}>{item.component_part_number||"—"}</div>
                    <div style={{fontSize:12,color:T.sub}}>{item.material}</div>
                    <div style={{fontSize:10,color:T.muted,marginTop:2}}>{i.qpa}× per assembly · Reorder at {i.reorderLevel||"—"}</div>
                  </div>
                  <div style={{textAlign:"right",display:"flex",flexDirection:"column",alignItems:"flex-end",gap:4}}>
                    {canAdmin&&<button onClick={async e=>{e.stopPropagation();if(window.confirm("Remove "+item.component_part_number+" from BOM?"))try{await API.mfg.bom.remove(item.id);await load();}catch(err){}}} style={{background:"none",border:`1px solid ${T.red}30`,borderRadius:6,padding:"2px 8px",color:T.red,fontSize:10,cursor:"pointer",fontFamily:"inherit"}}>🗑 Remove</button>}
                    <div style={{fontSize:28,fontWeight:900,color:i.needsReorder?T.red:i.onHand<=0?T.yellow:T.green,lineHeight:1}}>{i.onHand}</div>
                    <div style={{fontSize:10,color:T.muted,marginTop:2}}>on hand</div>
                    {i.needsReorder&&<div style={{background:"#7c2d12",color:"#fca5a5",borderRadius:6,padding:"2px 8px",fontSize:9,fontWeight:800,marginTop:4}}>🔴 REORDER</div>}
                  </div>
                </div>
                <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr 1fr",gap:6}}>
                  {[["Received",i.received,T.blue],["Used",i.usedInAsm,T.muted],["Damaged",i.damaged,i.damaged>0?T.red:T.muted],["Can Build",i.canBuild,i.canBuild>0?T.green:T.red]].map(([l,v,c])=>(
                    <div key={l} style={{background:T.surface,borderRadius:6,padding:"6px",textAlign:"center"}}>
                      <div style={{fontSize:15,fontWeight:800,color:c}}>{v}</div>
                      <div style={{fontSize:8,color:T.muted,textTransform:"uppercase"}}>{l}</div>
                    </div>
                  ))}
                </div>
              </div>
            );
          })}
        </>}

        {}
        {!loading&&tab==="received"&&<>

          {/* STEP 1: Create assembly if none exist */}
          <div style={{...cardS,marginBottom:12,border:`1px solid ${T.purple}40`}}>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}>
              <div>
                <div style={{fontSize:13,fontWeight:800,color:T.purple}}>Step 1 — What are you BUILDING?</div>
                <div style={{fontSize:11,color:T.muted,marginTop:2}}>
                  {parts.length===0?"Enter the finished product you're manufacturing (e.g. 1651 — Boom Pivot)":parts.map(p=>p.part_number+(p.description?" – "+p.description:"")).join(", ")}
                </div>
              </div>
              {canAdmin&&<button onClick={()=>setShowNewPart(s=>!s)}
                style={{background:parts.length===0?T.purple:T.surface,color:parts.length===0?"#fff":T.muted,border:`1px solid ${T.purple}40`,borderRadius:10,padding:"6px 12px",fontSize:12,fontWeight:700,cursor:"pointer",fontFamily:"inherit",flexShrink:0}}>
                {showNewPart?"✕ Cancel":"+ Add"}
              </button>}
            </div>
            {showNewPart&&<div style={{marginTop:12,paddingTop:12,borderTop:`1px solid ${T.border}`}}>
              <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8,marginBottom:8}}>
                <div><label style={lbl}>Finished Product Part # *</label><input value={pf.part_number} onChange={e=>setPf(x=>({...x,part_number:e.target.value}))} placeholder="e.g. 1651 (your part #)" style={inp} autoFocus/></div>
                <div><label style={lbl}>Description</label><input value={pf.description} onChange={e=>setPf(x=>({...x,description:e.target.value}))} placeholder="e.g. Boom Pivot" style={inp}/></div>
              </div>
              <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8,marginBottom:10}}>
                <div><label style={lbl}>Drawing #</label><input value={pf.drawing_number} onChange={e=>setPf(x=>({...x,drawing_number:e.target.value}))} placeholder="DWG-001" style={inp}/></div>
                <div><label style={lbl}>Qty Ordered</label><input type="number" value={pf.qty_ordered} onChange={e=>setPf(x=>({...x,qty_ordered:e.target.value}))} placeholder="e.g. 500" style={inp}/></div>
              </div>
              <button onClick={createPart} disabled={!pf.part_number.trim()||saving}
                style={{...primBtn,borderRadius:12,background:T.purple,opacity:pf.part_number.trim()&&!saving?1:0.5}}>
                {saving?"Creating…":"Create Finished Part"}
              </button>
            </div>}
          </div>

          {/* STEP 2: Add component parts */}
          {parts.length>0&&<div style={{...cardS,marginBottom:12,border:`1px solid ${T.orange}40`}}>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}>
              <div>
                <div style={{fontSize:13,fontWeight:800,color:T.orange}}>Step 2 — What parts COME IN from the customer?</div>
                <div style={{fontSize:11,color:T.muted,marginTop:2}}>
                  {allBomItems.length===0?"These are the JLG/customer part #s that arrive as raw materials (e.g. 3572922 Side Pieces)":""+allBomItems.length+" component part"+(allBomItems.length!==1?"s":"")+" configured"}
                </div>
              </div>
              {canAdmin&&<button onClick={()=>setShowAddComp(s=>!s)}
                style={{background:allBomItems.length===0?T.orange:T.surface,color:allBomItems.length===0?"#000":T.muted,border:`1px solid ${T.orange}40`,borderRadius:10,padding:"6px 12px",fontSize:12,fontWeight:700,cursor:"pointer",fontFamily:"inherit",flexShrink:0}}>
                {showAddComp?"✕ Cancel":"+ Add"}
              </button>}
            </div>
            {showAddComp&&<div style={{marginTop:12,paddingTop:12,borderTop:`1px solid ${T.border}`}}>
              <div style={{marginBottom:8}}><label style={lbl}>Assembly *</label>
                <select value={newCompPartId} onChange={e=>setNewCompPartId(e.target.value)} style={inpSel}>
                  <option value="">— Select assembly —</option>
                  {parts.map(p=><option key={p.id} value={p.id}>{p.part_number}{p.description?" — "+p.description:""}</option>)}
                </select>
              </div>
              <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8,marginBottom:8}}>
                <div><label style={lbl}>Customer (JLG) Part # *</label><input value={newCompNum} onChange={e=>setNewCompNum(e.target.value)} placeholder="e.g. 3572922" style={inp}/></div>
                <div><label style={lbl}>Description *</label><input value={newCompDesc} onChange={e=>setNewCompDesc(e.target.value)} placeholder="e.g. Side Pieces" style={inp}/></div>
              </div>
              <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8,marginBottom:10}}>
                <div><label style={lbl}>Qty Per Assembly</label><input type="number" value={newCompQpa} onChange={e=>setNewCompQpa(e.target.value)} placeholder="1" style={inp}/></div>
                <div><label style={lbl}>Reorder Level</label><input type="number" value={newCompReorder} onChange={e=>setNewCompReorder(e.target.value)} placeholder="e.g. 50" style={inp}/></div>
              </div>
              <button onClick={addCompPart} disabled={!newCompNum.trim()||!newCompDesc.trim()||!newCompPartId||saving}
                style={{...primBtn,borderRadius:12,background:T.orange,color:"#000",opacity:newCompNum.trim()&&newCompDesc.trim()&&newCompPartId&&!saving?1:0.5}}>
                {saving?"Adding…":"Add Incoming Part"}
              </button>
            </div>}
            {/* List current components */}
            {allBomItems.length>0&&<div style={{marginTop:10}}>
              {allBomItems.map(item=>(
                <div key={item.id} style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"6px 0",borderBottom:`1px solid ${T.border}`,fontSize:12}}>
                  <div><span style={{color:T.orange,fontWeight:700}}>{item.component_part_number}</span> <span style={{color:T.sub}}>{item.material}</span> <span style={{color:T.muted}}>· {item.qty_per_assembly||1}x/asm</span></div>
                  {canAdmin&&<button onClick={async()=>{if(window.confirm("Remove?"))try{await API.mfg.bom.remove(item.id);await load();}catch(e){}}} style={{background:"none",border:"none",color:T.red,cursor:"pointer",fontSize:12}}>🗑</button>}
                </div>
              ))}
            </div>}
          </div>}

          {/* STEP 3: Log receipt */}
          {allBomItems.length>0&&<div style={{...cardS,marginBottom:16,border:`1px solid ${T.green}40`}}>
            <div style={{fontSize:13,fontWeight:800,color:T.green,marginBottom:12}}>Step 3 — Log Parts Received from Customer</div>

            {formErr&&<div style={{background:T.redLow,border:`1px solid ${T.red}40`,borderRadius:8,padding:"8px 12px",marginBottom:10,fontSize:12,color:T.red}}>{formErr}</div>}

            {/* Part selector */}
            <div style={{marginBottom:10}}>
              <label style={lbl}>Select Component Part *</label>
              <select value={selectedBomItem?.id||""} onChange={e=>{const item=allBomItems.find(b=>b.id===e.target.value);setSelectedBomItem(item||null);}} style={inpSel}>
                <option value="">— Select a part —</option>
                {parts.map(p=><optgroup key={p.id} label={`Assembly: ${p.part_number}`}>
                  {(boms[p.id]||[]).map(item=>(
                    <option key={item.id} value={item.id}>{item.component_part_number} — {item.material} ({item.qty_per_assembly||1}x per asm)</option>
                  ))}
                </optgroup>)}
              </select>
            </div>

            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8,marginBottom:8}}>
              <div><label style={lbl}>Qty Received *</label><input type="number" value={rQty} onChange={e=>setRQty(e.target.value)} placeholder="0" style={inp}/></div>
              <div><label style={lbl}>Date</label><input type="date" value={rDate} onChange={e=>setRDate(e.target.value)} style={inp}/></div>
            </div>
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8,marginBottom:12}}>
              <div><label style={lbl}>BOL / Reference #</label><input value={rBol} onChange={e=>setRBol(e.target.value)} placeholder="BOL #" style={inp}/></div>
              <div><label style={lbl}>Received By</label><input value={rBy} onChange={e=>setRBy(e.target.value)} style={inp}/></div>
            </div>

            <button onClick={logReceipt} disabled={saving||!rQty||!selectedBomItem}
              style={{...primBtn,borderRadius:12,background:T.green,color:"#000",opacity:rQty&&selectedBomItem&&!saving?1:0.4}}>
              {saving?"Saving…":"✓ Log Receipt"}
            </button>
          </div>}

          {/* Running receipt list */}
          <div style={{fontSize:12,fontWeight:800,color:T.text,marginBottom:10,marginTop:4}}>Receipt History</div>
          {receipts.filter(r=>r.transaction_type==="Received").length===0&&<div style={{textAlign:"center",padding:30,color:T.muted,fontSize:12}}>No receipts logged yet. Use the form above to log incoming parts.</div>}
          {receipts.filter(r=>r.transaction_type==="Received").map((r,i)=>{
            const bomItem=allBomItems.find(b=>b.id===r.bom_id);
            return(
              <div key={i} style={{...cardS,marginBottom:8,display:"flex",justifyContent:"space-between",alignItems:"center"}}>
                <div>
                  <div style={{fontSize:13,fontWeight:700,color:T.green}}>+{r.qty_received} — <span style={{color:T.orange}}>{bomItem?.component_part_number||"—"}</span> <span style={{color:T.sub,fontWeight:400}}>{bomItem?.material||""}</span></div>
                  <div style={{fontSize:11,color:T.muted}}>{r.received_date}{r.received_by?" · "+r.received_by:""}{r.job_number?" · "+r.job_number:""}{r.reference_bol?" · BOL: "+r.reference_bol:""}</div>
                </div>
                {canAdmin&&<button onClick={async()=>{if(window.confirm("Delete this receipt?"))try{await sb(`/mfg_receipts?id=eq.${r.id}`,{method:"DELETE"});await load();}catch(e){}}} style={{background:"none",border:`1px solid ${T.red}30`,borderRadius:8,padding:"4px 8px",color:T.red,fontSize:11,cursor:"pointer",fontFamily:"inherit"}}>🗑</button>}
              </div>
            );
          })}
        </>}

        {}
        {!loading&&tab==="assembly"&&<>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:8,marginBottom:14}}>
            {[[assemblyLogs.reduce((s,a)=>s+(a.qty_completed||0),0),"Completed",T.green],[totalShipped,"Shipped",T.blue],[totalReadyToShip,"Ready to Ship",T.orange]].map(([v,l,c])=>(
              <div key={l} style={{...cardS,textAlign:"center"}}><div style={{fontSize:26,fontWeight:900,color:c}}>{v}</div><div style={{fontSize:10,color:T.muted,textTransform:"uppercase"}}>{l}</div></div>
            ))}
          </div>
          <button onClick={()=>setShowAsmForm(s=>!s)} style={{...primBtn,borderRadius:14,marginBottom:14,background:T.green,color:"#000"}}>🏭 Log Completed Assemblies</button>
          {showAsmForm&&<div style={{...cardS,marginBottom:14,border:`1px solid ${T.green}40`}}>
            <div style={{fontSize:13,fontWeight:800,color:T.green,marginBottom:12}}>Log Assembly Batch</div>
            <div style={{marginBottom:10}}><label style={lbl}>Assembly *</label>
              <select value={af.part_id} onChange={e=>setAf(x=>({...x,part_id:e.target.value}))} style={inpSel}>
                <option value="">— Select —</option>
                {parts.map(p=><option key={p.id} value={p.id}>{p.part_number}{p.description?" — "+p.description:""}</option>)}
              </select>
            </div>
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8,marginBottom:10}}>
              <div><label style={lbl}>Qty Completed *</label><input type="number" value={af.qty} onChange={e=>setAf(x=>({...x,qty:e.target.value}))} placeholder="e.g. 71" style={inp}/></div>
              <div><label style={lbl}>Date</label><input type="date" value={af.date} onChange={e=>setAf(x=>({...x,date:e.target.value}))} style={inp}/></div>
            </div>
            <div style={{marginBottom:10}}><label style={lbl}>Entered By</label><input value={af.by} onChange={e=>setAf(x=>({...x,by:e.target.value}))} style={inp}/></div>
            {af.part_id&&af.qty&&<div style={{background:T.blueLow,borderRadius:8,padding:"8px 12px",marginBottom:10,fontSize:11,color:T.blue}}>ℹ️ Will auto-deduct {af.qty} × each component from inventory</div>}
            <div style={{display:"flex",gap:8}}>
              <button onClick={logAssembly} disabled={!af.qty||!af.part_id||saving} style={{...primBtn,flex:2,borderRadius:12,background:T.green,color:"#000",opacity:af.qty&&af.part_id&&!saving?1:0.5}}>{saving?"Saving…":"Log Assemblies"}</button>
              <button onClick={()=>setShowAsmForm(false)} style={{...ghostBtn,flex:1,textAlign:"center"}}>Cancel</button>
            </div>
          </div>}
          {assemblyLogs.map(a=>{
            const part=parts.find(p=>p.id===a.part_id);
            return(<div key={a.id} style={{...cardS,marginBottom:8,borderLeft:`3px solid ${T.green}`,display:"flex",justifyContent:"space-between",alignItems:"center"}}>
              <div>
                <div style={{fontSize:14,fontWeight:800,color:T.green}}>{a.qty_completed} assemblies — {part?.part_number||"—"}</div>
                <div style={{fontSize:11,color:T.muted}}>{a.completion_date}{a.entered_by?" · "+a.entered_by:""}</div>
              </div>
              {canAdmin&&<button onClick={async()=>{if(window.confirm("Delete?"))try{await sb(`/mfg_assembly_log?id=eq.${a.id}`,{method:"DELETE"});await load();}catch(e){}}} style={{background:"none",border:`1px solid ${T.red}30`,borderRadius:8,padding:"4px 8px",color:T.red,fontSize:11,cursor:"pointer",fontFamily:"inherit"}}>🗑</button>}
            </div>);
          })}
          {assemblyLogs.length===0&&<div style={{textAlign:"center",padding:30,color:T.muted,fontSize:12}}>No assemblies logged yet.</div>}
        </>}

        {}
        {!loading&&tab==="shipping"&&<>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:8,marginBottom:14}}>
            {[[assemblyLogs.reduce((s,a)=>s+(a.qty_completed||0),0),"Completed",T.green],[totalShipped,"Shipped",T.blue],[totalReadyToShip,"Ready to Ship",T.orange]].map(([v,l,c])=>(
              <div key={l} style={{...cardS,textAlign:"center"}}><div style={{fontSize:26,fontWeight:900,color:c}}>{v}</div><div style={{fontSize:10,color:T.muted,textTransform:"uppercase"}}>{l}</div></div>
            ))}
          </div>
          <div style={{display:"flex",gap:8,marginBottom:14}}>
            <button onClick={()=>setShowShipForm(s=>!s)} style={{...primBtn,flex:2,borderRadius:12,background:T.blue}}>📤 Log Shipment</button>
            <button onClick={()=>setShowPackingSlip(true)} style={{background:"#1f3864",color:"#fff",border:"none",borderRadius:12,padding:"12px",fontSize:12,fontWeight:700,cursor:"pointer",fontFamily:"inherit",flex:1}}>📄 Packing Slip</button>
          </div>
          {showShipForm&&<div style={{...cardS,marginBottom:14,border:`1px solid ${T.blue}40`}}>
            <div style={{fontSize:13,fontWeight:800,color:T.blue,marginBottom:12}}>Log Shipment</div>
            <div style={{marginBottom:10}}><label style={lbl}>Assembly *</label>
              <select value={sf.part_id} onChange={e=>setSf(x=>({...x,part_id:e.target.value}))} style={inpSel}>
                <option value="">— Select —</option>
                {parts.map(p=><option key={p.id} value={p.id}>{p.part_number}{p.description?" — "+p.description:""}</option>)}
              </select>
            </div>
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8,marginBottom:10}}>
              <div><label style={lbl}>Qty Shipped *</label><input type="number" value={sf.qty} onChange={e=>setSf(x=>({...x,qty:e.target.value}))} placeholder="0" style={inp}/></div>
              <div><label style={lbl}>Ship Date</label><input type="date" value={sf.date} onChange={e=>setSf(x=>({...x,date:e.target.value}))} style={inp}/></div>
            </div>
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8,marginBottom:10}}>
              <div><label style={lbl}>Customer</label><input value={sf.customer} onChange={e=>setSf(x=>({...x,customer:e.target.value}))} style={inp}/></div>
              <div><label style={lbl}>BOL #</label><input value={sf.bol} onChange={e=>setSf(x=>({...x,bol:e.target.value}))} placeholder="BOL #" style={inp}/></div>
            </div>
            <div style={{display:"flex",gap:8}}>
              <button onClick={logShipment} disabled={!sf.qty||!sf.part_id||saving} style={{...primBtn,flex:2,borderRadius:12,background:T.blue,opacity:sf.qty&&sf.part_id&&!saving?1:0.5}}>{saving?"Saving…":"Log Shipment"}</button>
              <button onClick={()=>setShowShipForm(false)} style={{...ghostBtn,flex:1,textAlign:"center"}}>Cancel</button>
            </div>
          </div>}
          {shippingLogs.map(s=>{
            const part=parts.find(p=>p.id===s.part_id);
            return(<div key={s.id} style={{...cardS,marginBottom:8,borderLeft:`3px solid ${T.blue}`,display:"flex",justifyContent:"space-between",alignItems:"center"}}>
              <div>
                <div style={{fontSize:14,fontWeight:800,color:T.blue}}>{s.qty_shipped} shipped — {part?.part_number||"—"}</div>
                <div style={{fontSize:11,color:T.muted}}>{s.ship_date}{s.customer?" · "+s.customer:""}{s.bol_number?" · BOL: "+s.bol_number:""}</div>
              </div>
              {canAdmin&&<button onClick={async()=>{if(window.confirm("Delete?"))try{await sb(`/mfg_shipping_log?id=eq.${s.id}`,{method:"DELETE"});await load();}catch(e){}}} style={{background:"none",border:`1px solid ${T.red}30`,borderRadius:8,padding:"4px 8px",color:T.red,fontSize:11,cursor:"pointer",fontFamily:"inherit"}}>🗑</button>}
            </div>);
          })}
          {shippingLogs.length===0&&<div style={{textAlign:"center",padding:30,color:T.muted,fontSize:12}}>No shipments logged yet.</div>}

          {/* Saved Packing Slips */}
          {packingSlips.length>0&&<>
            <div style={{fontSize:12,fontWeight:800,color:T.text,marginTop:16,marginBottom:8}}>📄 Saved Packing Slips</div>
            {packingSlips.map(slip=>(
              <div key={slip.id} style={{...cardS,marginBottom:8,borderLeft:`3px solid #1f3864`,display:"flex",justifyContent:"space-between",alignItems:"center"}}>
                <div>
                  <div style={{fontSize:13,fontWeight:800,color:"#60A5FA"}}>Packing Slip {slip.slip_number||"(no #)"}</div>
                  <div style={{fontSize:11,color:T.muted}}>{slip.ship_date||""}{slip.created_by?" · "+slip.created_by:""} · Saved {new Date(slip.created_at).toLocaleDateString()}</div>
                </div>
                <div style={{display:"flex",gap:6}}>
                  <button onClick={()=>{setEditingSlip(slip);setShowPackingSlip(true);}}
                    style={{background:"#1f3864",color:"#fff",border:"none",borderRadius:8,padding:"5px 10px",fontSize:11,fontWeight:700,cursor:"pointer",fontFamily:"inherit"}}>
                    Open
                  </button>
                  {canAdmin&&<button onClick={async()=>{if(window.confirm("Delete this packing slip?"))try{await API.mfg.packingSlips.remove(slip.id);await load();}catch(e){}}}
                    style={{background:"none",border:`1px solid ${T.red}30`,borderRadius:8,padding:"5px 8px",color:T.red,fontSize:11,cursor:"pointer",fontFamily:"inherit"}}>🗑</button>}
                </div>
              </div>
            ))}
          </>}
        </>}
      </div>
    </div>
  );
}

function AddBomItemSection({parts,onSaved,job}){
  const [show,setShow]=useState(false);
  const [partId,setPartId]=useState("");
  const [compNum,setCompNum]=useState("");
  const [desc,setDesc]=useState("");
  const [qpa,setQpa]=useState("1");
  const [reorder,setReorder]=useState("");
  const [saving,setSaving]=useState(false);

  async function save(){
    if(!compNum.trim()||!desc.trim()||!partId)return;
    setSaving(true);
    try{
      await API.mfg.bom.create({part_id:partId,component_part_number:compNum.trim(),material:desc.trim(),qty_per_assembly:parseFloat(qpa)||1,reorder_level:parseFloat(reorder)||0,unit:"ea"});
      setCompNum("");setDesc("");setQpa("1");setReorder("");setShow(false);
      await onSaved();
    }catch(e){alert("Error: "+e.message);}
    setSaving(false);
  }

  return(
    <div style={{marginBottom:16}}>
      <button onClick={()=>setShow(s=>!s)} style={{...ghostBtn,width:"100%",textAlign:"center",color:T.purple,border:`1px solid ${T.purple}40`,fontSize:12,marginBottom:show?10:0}}>
        {show?"✕ Cancel":"+ Add Component Part Number"}
      </button>
      {show&&<div style={{...cardS,border:`1px solid ${T.purple}40`}}>
        <div style={{fontSize:12,fontWeight:800,color:T.purple,marginBottom:12}}>Add Component Part to Assembly</div>
        <div style={{marginBottom:10}}><label style={lbl}>Finished Product Part # *</label>
          <select value={partId} onChange={e=>setPartId(e.target.value)} style={inpSel}>
            <option value="">— Select assembly —</option>
            {parts.map(p=><option key={p.id} value={p.id}>{p.part_number}{p.description?" — "+p.description:""}</option>)}
          </select>
        </div>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8,marginBottom:8}}>
          <div><label style={lbl}>Customer (JLG) Part # *</label><input value={compNum} onChange={e=>setCompNum(e.target.value)} placeholder="e.g. 3572922" style={inp}/></div>
          <div><label style={lbl}>Description *</label><input value={desc} onChange={e=>setDesc(e.target.value)} placeholder="e.g. Side Pieces" style={inp}/></div>
        </div>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8,marginBottom:12}}>
          <div><label style={lbl}>Qty Per Assembly</label><input type="number" value={qpa} onChange={e=>setQpa(e.target.value)} placeholder="1" style={inp}/></div>
          <div><label style={lbl}>Reorder Level</label><input type="number" value={reorder} onChange={e=>setReorder(e.target.value)} placeholder="e.g. 50" style={inp}/></div>
        </div>
        <button onClick={save} disabled={!compNum.trim()||!desc.trim()||!partId||saving} style={{...primBtn,borderRadius:12,background:T.purple,opacity:compNum.trim()&&desc.trim()&&partId&&!saving?1:0.5}}>
          {saving?"Adding…":"Add Incoming Part"}
        </button>
      </div>}
    </div>
  );
}

function BomAddForm({partId,onSave,onCancel}){
  const [f,setF]=useState({component_part_number:"",material:"",spec:"",qty_per_assembly:"1",unit:"ea",reorder_level:""});
  const UNITS=["ea","ft","in","lbs","kg","m","mm","set","pcs"];
  return(
    <div style={{background:T.surface,borderRadius:10,padding:12,marginBottom:10,border:`1px solid ${T.purple}30`}}>
      <div style={{fontSize:11,fontWeight:700,color:T.purple,marginBottom:10,textTransform:"uppercase",letterSpacing:"1px"}}>Add Component Part</div>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8,marginBottom:8}}>
        <div><label style={lbl}>Customer (JLG) Part # *</label>
          <input value={f.component_part_number} onChange={e=>setF(x=>({...x,component_part_number:e.target.value}))} placeholder="e.g. 3572922" style={inp}/>
        </div>
        <div><label style={lbl}>Description *</label>
          <input value={f.material} onChange={e=>setF(x=>({...x,material:e.target.value}))} placeholder="e.g. Side Pieces" style={inp}/>
        </div>
      </div>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:8,marginBottom:8}}>
        <div><label style={lbl}>Qty / Assembly *</label>
          <input type="number" value={f.qty_per_assembly} onChange={e=>setF(x=>({...x,qty_per_assembly:e.target.value}))} placeholder="1" style={inp}/>
        </div>
        <div><label style={lbl}>Unit</label>
          <select value={f.unit} onChange={e=>setF(x=>({...x,unit:e.target.value}))} style={inpSel}>{UNITS.map(u=><option key={u}>{u}</option>)}</select>
        </div>
        <div><label style={lbl}>Reorder Level</label>
          <input type="number" value={f.reorder_level} onChange={e=>setF(x=>({...x,reorder_level:e.target.value}))} placeholder="e.g. 50" style={inp}/>
        </div>
      </div>
      <div style={{marginBottom:10}}><label style={lbl}>Spec (optional)</label>
        <input value={f.spec} onChange={e=>setF(x=>({...x,spec:e.target.value}))} placeholder="Material spec" style={inp}/>
      </div>
      <div style={{display:"flex",gap:8}}>
        <button onClick={()=>f.component_part_number.trim()&&f.material.trim()&&onSave({...f,qty_per_assembly:parseFloat(f.qty_per_assembly)||1,reorder_level:parseFloat(f.reorder_level)||0})}
          style={{...primBtn,flex:2,borderRadius:10,fontSize:13,background:T.purple,opacity:f.component_part_number.trim()&&f.material.trim()?1:0.5}}>
          Add Component
        </button>
        <button onClick={onCancel} style={{...ghostBtn,flex:1,textAlign:"center",fontSize:13}}>Cancel</button>
      </div>
    </div>
  );
}

function InventoryTab({parts,boms,txns,job,user,canAdmin,onRefresh}){
  const [showNewBom,setShowNewBom]=useState(null);
  const [showTxn,setShowTxn]=useState(null); // {bomItem, type}
  const [tf,setTf]=useState({qty:"",transaction_type:"Received",reference_bol:"",notes:"",received_date:today(),received_by:""});
  const [saving,setSaving]=useState(false);
  const set=(k,v)=>setTf(x=>({...x,[k]:v}));

  function calcInv(item){
    const itemTxns=txns[item.id]||[];
    const received=itemTxns.reduce((s,t)=>s+(t.qty_received||0),0);
    const issued=itemTxns.reduce((s,t)=>s+(t.qty_issued||0),0);
    const damaged=itemTxns.reduce((s,t)=>s+(t.qty_damaged||0),0);
    const usedInAsm=itemTxns.filter(t=>t.transaction_type==="Issued").reduce((s,t)=>s+(t.qty_received||0),0);
    const onHand=received-issued-damaged-usedInAsm;
    const qpa=item.qty_per_assembly||1;
    const reorderLevel=item.reorder_level||0;
    const needsReorder=reorderLevel>0&&onHand<=reorderLevel;
    return{received,issued,damaged,usedInAsm,onHand,canBuild:Math.floor(Math.max(0,onHand)/qpa),needsReorder,reorderLevel};
  }

  async function logTxn(){
    if(!tf.qty||!showTxn)return;
    setSaving(true);
    try{
      const isReceived=showTxn.type==="Received";
      const isDamaged=showTxn.type==="Damaged";
      await API.mfg.receipts.create({
        bom_id:showTxn.item.id,
        part_id:showTxn.item.part_id,
        qty_received:isReceived?parseFloat(tf.qty):0,
        qty_damaged:isDamaged?parseFloat(tf.qty):0,
        qty_issued:0,
        transaction_type:showTxn.type,
        reference_bol:tf.reference_bol||null,
        notes:tf.notes||null,
        received_date:tf.received_date,
        received_by:tf.received_by||user.name,
        heat_number:tf.heat_number||null,
        condition:isDamaged?"Damaged":"Good",
        job_number:job.job_number||null,
      });
      setShowTxn(null);
      setTf({qty:"",transaction_type:"Received",reference_bol:"",notes:"",received_date:today(),received_by:""});
      await onRefresh();
    }catch(e){alert("Error: "+e.message);}
    setSaving(false);
  }

  return(
    <div>
      <div style={{fontSize:12,color:T.muted,marginBottom:14,lineHeight:1.6,background:T.blueLow,border:`1px solid ${T.blue}30`,borderRadius:10,padding:"10px 14px"}}>
        📦 Track all GFM (Customer Furnished) component parts. Each component has its own customer part # and qty required per assembly.
      </div>

      {parts.map(part=>{
        const bom=boms[part.id]||[];
        const minCanBuild=bom.length>0?Math.min(...bom.map(item=>calcInv(item).canBuild)):0;
        return(
          <div key={part.id} style={{...cardS,marginBottom:14}}>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:10}}>
              <div>
                <div style={{fontSize:14,fontWeight:800,color:T.purple}}>{part.part_number}</div>
                {part.description&&<div style={{fontSize:11,color:T.muted}}>{part.description}</div>}
              </div>
              <div style={{textAlign:"right"}}>
                <div style={{fontSize:13,fontWeight:700,color:bom.length>0?T.green:T.muted}}>Can build: {minCanBuild}</div>
                <div style={{fontSize:10,color:T.muted}}>assemblies now</div>
              </div>
            </div>

            {canAdmin&&<button onClick={()=>setShowNewBom(showNewBom===part.id?null:part.id)}
              style={{...ghostBtn,width:"100%",textAlign:"center",fontSize:12,color:T.purple,border:`1px solid ${T.purple}40`,marginBottom:10}}>
              + Add Component Part
            </button>}

            {showNewBom===part.id&&<BomAddForm partId={part.id} onSave={async(d)=>{
              await API.mfg.bom.create({...d,part_id:part.id});
              setShowNewBom(null);await onRefresh();
            }} onCancel={()=>setShowNewBom(null)}/>}

            {/* Component list */}
            {bom.length===0&&<div style={{fontSize:12,color:T.muted,fontStyle:"italic",textAlign:"center",padding:"10px 0"}}>No components added — tap above to add customer part numbers</div>}
            {bom.map(item=>{
              const inv=calcInv(item);
              const short=inv.onHand<0;
              const itemTxns=txns[item.id]||[];
              return(
                <div key={item.id} style={{background:T.surface,borderRadius:10,padding:"10px 12px",marginBottom:8,borderLeft:`3px solid ${short?T.red:inv.onHand===0?T.yellow:T.green}`}}>
                  <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:8}}>
                    <div>
                      <div style={{display:"flex",gap:8,alignItems:"center"}}>
                        <div style={{fontSize:13,fontWeight:800,color:T.orange}}>{item.component_part_number||"—"}</div>
                        <span style={{background:T.surface,border:`1px solid ${T.border}`,borderRadius:4,padding:"1px 6px",fontSize:9,color:T.muted,fontWeight:700}}>{item.qty_per_assembly||1}× per asm</span>
                      </div>
                      <div style={{fontSize:12,color:T.sub,marginTop:1}}>{item.material}</div>
                      {item.spec&&<div style={{fontSize:10,color:T.muted}}>Spec: {item.spec}</div>}
                    </div>
                    <div style={{textAlign:"right",flexShrink:0}}>
                      <div style={{fontSize:18,fontWeight:900,color:short?T.red:inv.onHand===0?T.yellow:T.green}}>{inv.onHand}</div>
                      <div style={{fontSize:9,color:T.muted,textTransform:"uppercase"}}>On Hand</div>
                    </div>
                  </div>

                  {/* REORDER badge */}
                  {inv.needsReorder&&<div style={{background:"#7c2d12",border:"1px solid #f97316",borderRadius:8,padding:"5px 10px",fontSize:11,color:"#fed7aa",fontWeight:800,marginBottom:8}}>
                    🔴 REORDER — {inv.onHand} on hand (reorder at {inv.reorderLevel})
                  </div>}
                  {/* Stats row */}
                  <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr 1fr 1fr",gap:4,marginBottom:8}}>
                    {[["Received",inv.received,T.blue],["Used/Asm",inv.usedInAsm,T.muted],["Issued",inv.issued,T.muted],["Damaged",inv.damaged,T.red],["On Hand",inv.onHand,inv.onHand<=0?T.red:T.green]].map(([l,v,c])=>(
                      <div key={l} style={{background:T.card,borderRadius:6,padding:"4px",textAlign:"center"}}>
                        <div style={{fontSize:13,fontWeight:800,color:c}}>{v}</div>
                        <div style={{fontSize:8,color:T.muted,textTransform:"uppercase"}}>{l}</div>
                      </div>
                    ))}
                  </div>

                  {/* Transaction buttons */}
                  <div style={{display:"flex",gap:6}}>
                    {["Received","Damaged"].map(type=>(
                      <button key={type} onClick={()=>{setShowTxn({item,type});setTf({qty:"",transaction_type:type,reference_bol:"",notes:"",received_date:today(),received_by:user.name,heat_number:""});}}
                        style={{...ghostBtn,flex:1,textAlign:"center",fontSize:11,padding:"6px 4px",color:type==="Received"?T.green:T.red,border:`1px solid ${type==="Received"?T.green:T.red}30`}}>
                        {type==="Received"?"📦 Receive":"⚠️ Damage"}
                      </button>
                    ))}
                    {canAdmin&&<button onClick={async()=>{if(window.confirm("Remove?"))await API.mfg.bom.remove(item.id).then(onRefresh);}} style={{...ghostBtn,fontSize:11,padding:"6px 10px",color:T.red,border:`1px solid ${T.red}30`}}>🗑</button>}
                  </div>

                  {/* Transaction form */}
                  {showTxn?.item?.id===item.id&&<div style={{marginTop:10,background:T.card,borderRadius:10,padding:12,border:`1px solid ${showTxn.type==="Received"?T.green:T.red}30`}}>
                    <div style={{fontSize:12,fontWeight:800,color:showTxn.type==="Received"?T.green:T.red,marginBottom:10}}>
                      {showTxn.type==="Received"?"📦 Log Receipt":"⚠️ Log Damage"} — {item.component_part_number||item.material}
                    </div>
                    <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8,marginBottom:8}}>
                      <div><label style={lbl}>Qty *</label><input type="number" value={tf.qty} onChange={e=>set("qty",e.target.value)} placeholder="0" style={inp}/></div>
                      <div><label style={lbl}>Date</label><input type="date" value={tf.received_date} onChange={e=>set("received_date",e.target.value)} style={inp}/></div>
                    </div>
                    {showTxn.type==="Received"&&<div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8,marginBottom:8}}>
                      <div><label style={lbl}>Heat / Lot #</label><input value={tf.heat_number||""} onChange={e=>set("heat_number",e.target.value)} placeholder="Heat #" style={inp}/></div>
                      <div><label style={lbl}>BOL / Ref #</label><input value={tf.reference_bol} onChange={e=>set("reference_bol",e.target.value)} placeholder="BOL #" style={inp}/></div>
                    </div>}
                    <div style={{marginBottom:8}}><label style={lbl}>Received By</label><input value={tf.received_by} onChange={e=>set("received_by",e.target.value)} style={inp}/></div>
                    <div style={{marginBottom:10}}><label style={lbl}>Notes</label><input value={tf.notes} onChange={e=>set("notes",e.target.value)} placeholder="Optional" style={inp}/></div>
                    <div style={{display:"flex",gap:8}}>
                      <button onClick={logTxn} disabled={!tf.qty||saving} style={{...primBtn,flex:2,borderRadius:10,background:showTxn.type==="Received"?T.green:T.red,color:"#000",opacity:tf.qty&&!saving?1:0.5}}>{saving?"Saving…":"Confirm"}</button>
                      <button onClick={()=>setShowTxn(null)} style={{...ghostBtn,flex:1,textAlign:"center"}}>Cancel</button>
                    </div>
                  </div>}

                  {/* Recent transactions */}
                  {itemTxns.length>0&&<div style={{marginTop:8,paddingTop:8,borderTop:`1px solid ${T.border}`}}>
                    <div style={{fontSize:10,fontWeight:700,color:T.muted,textTransform:"uppercase",marginBottom:4}}>Recent Transactions</div>
                    {itemTxns.slice(0,3).map((t,i)=>(
                      <div key={i} style={{display:"flex",justifyContent:"space-between",fontSize:11,padding:"3px 0",borderBottom:`1px solid ${T.border}`}}>
                        <span style={{color:t.transaction_type==="Received"?T.green:T.red}}>{t.transaction_type}</span>
                        <span style={{color:T.sub}}>{t.received_date}</span>
                        {t.reference_bol&&<span style={{color:T.muted}}>BOL: {t.reference_bol}</span>}
                        <span style={{fontWeight:700,color:T.text}}>{t.qty_received>0?"+":""}{t.qty_received}</span>
                      </div>
                    ))}
                  </div>}
                </div>
              );
            })}
          </div>
        );
      })}
    </div>
  );
}

function AssemblyLogTab({parts,assemblyLogs,boms,txns,job,user,canAdmin,onRefresh}){
  const [showForm,setShowForm]=useState(false);
  const [f,setF]=useState({part_id:"",qty_completed:"",completion_date:today(),entered_by:user.name,notes:""});
  const [saving,setSaving]=useState(false);
  const set=(k,v)=>setF(x=>({...x,[k]:v}));

  const totalCompleted=assemblyLogs.reduce((s,a)=>s+(a.qty_completed||0),0);

  async function save(){
    if(!f.qty_completed||!f.part_id)return;
    setSaving(true);
    try{
      await API.mfg.assemblyLog.create({...f,job_id:job.id,qty_completed:parseInt(f.qty_completed)||0,part_id:f.part_id||null});
      const partBom=boms[f.part_id]||[];
      const qtyAsm=parseInt(f.qty_completed)||0;
      await Promise.all(partBom.map(item=>
        API.mfg.receipts.create({
          bom_id:item.id,part_id:f.part_id,
          qty_received:qtyAsm*(item.qty_per_assembly||1),
          qty_issued:0, qty_damaged:0,
          transaction_type:"Issued",
          received_date:f.completion_date,
          received_by:f.entered_by,
          job_number:job.job_number||null,
          notes:`Auto: ${qtyAsm} assemblies completed on ${f.completion_date}`,
        }).catch(()=>{})
      ));
      setShowForm(false);setF({part_id:"",qty_completed:"",completion_date:today(),entered_by:user.name,notes:""});
      await onRefresh();
    }catch(e){alert("Error: "+e.message);}
    setSaving(false);
  }

  return(
    <div>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8,marginBottom:14}}>
        <div style={{...cardS,textAlign:"center"}}><div style={{fontSize:26,fontWeight:900,color:T.green}}>{totalCompleted}</div><div style={{fontSize:10,color:T.muted,textTransform:"uppercase"}}>Total Completed</div></div>
        <div style={{...cardS,textAlign:"center"}}><div style={{fontSize:26,fontWeight:900,color:T.purple}}>{parts.reduce((s,p)=>s+(p.qty_ordered||0),0)}</div><div style={{fontSize:10,color:T.muted,textTransform:"uppercase"}}>Total Ordered</div></div>
      </div>

      <button onClick={()=>setShowForm(s=>!s)} style={{...primBtn,borderRadius:14,marginBottom:14,background:T.green,color:"#000"}}>🏭 Log Completed Assemblies</button>

      {showForm&&<div style={{...cardS,marginBottom:14,border:`1px solid ${T.green}40`}}>
        <div style={{fontSize:13,fontWeight:800,color:T.green,marginBottom:12}}>🏭 Log Completed Assembly Batch</div>
        <div style={{marginBottom:10}}><label style={lbl}>Finished Product Part # *</label>
          <select value={f.part_id} onChange={e=>set("part_id",e.target.value)} style={inpSel}>
            <option value="">— Select part —</option>
            {parts.map(p=><option key={p.id} value={p.id}>{p.part_number}{p.description?" — "+p.description:""}</option>)}
          </select>
        </div>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:10}}>
          <div><label style={lbl}>Assemblies Completed *</label><input type="number" value={f.qty_completed} onChange={e=>set("qty_completed",e.target.value)} placeholder="e.g. 71" style={inp}/></div>
          <div><label style={lbl}>Date</label><input type="date" value={f.completion_date} onChange={e=>set("completion_date",e.target.value)} style={inp}/></div>
        </div>
        <div style={{marginBottom:10}}><label style={lbl}>Entered By</label><input value={f.entered_by} onChange={e=>set("entered_by",e.target.value)} style={inp}/></div>
        <div style={{marginBottom:10}}><label style={lbl}>Notes</label><textarea value={f.notes} onChange={e=>set("notes",e.target.value)} rows={2} style={{...inp,resize:"vertical"}}/></div>
        {f.part_id&&f.qty_completed&&<div style={{background:T.blueLow,border:`1px solid ${T.blue}30`,borderRadius:8,padding:"8px 12px",marginBottom:10,fontSize:11,color:T.blue}}>
          ℹ️ Logging {f.qty_completed} assemblies will auto-deduct components from inventory
        </div>}
        <div style={{display:"flex",gap:8}}>
          <button onClick={save} disabled={!f.qty_completed||!f.part_id||saving} style={{...primBtn,flex:2,borderRadius:12,background:T.green,color:"#000",opacity:f.qty_completed&&f.part_id&&!saving?1:0.5}}>{saving?"Saving…":"Log Assemblies"}</button>
          <button onClick={()=>setShowForm(false)} style={{...ghostBtn,flex:1,textAlign:"center"}}>Cancel</button>
        </div>
      </div>}

      {assemblyLogs.length===0&&<div style={{textAlign:"center",padding:"40px 16px",color:T.muted}}>
        <div style={{fontSize:44,marginBottom:12}}>🏭</div>
        <div style={{fontSize:14,fontWeight:700,color:T.sub,marginBottom:6}}>No Assemblies Logged</div>
        <div style={{fontSize:12}}>Tap the button above when batches of assemblies are completed.</div>
      </div>}

      {assemblyLogs.map(a=>{
        const part=parts.find(p=>p.id===a.part_id);
        return(
          <div key={a.id} style={{...cardS,marginBottom:8,borderLeft:`3px solid ${T.green}`}}>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start"}}>
              <div>
                <div style={{fontSize:16,fontWeight:900,color:T.green}}>{a.qty_completed} assemblies</div>
                <div style={{fontSize:12,color:T.sub}}>{a.completion_date} · {part?.part_number||"—"}</div>
                {a.entered_by&&<div style={{fontSize:11,color:T.muted}}>By: {a.entered_by}</div>}
              </div>
              <div style={{display:"flex",gap:6,alignItems:"center"}}>
                <div style={{fontSize:20}}>🏭</div>
                {canAdmin&&<button onClick={async()=>{if(window.confirm("Delete this assembly log entry?"))try{await sb(`/mfg_assembly_log?id=eq.${a.id}`,{method:"DELETE"});await onRefresh();}catch(e){}}} style={{background:"none",border:`1px solid ${T.red}40`,borderRadius:8,padding:"4px 8px",color:T.red,fontSize:11,cursor:"pointer",fontFamily:"inherit"}}>🗑</button>}
              </div>
            </div>
            {a.notes&&<div style={{fontSize:11,color:T.muted,marginTop:4,fontStyle:"italic"}}>{a.notes}</div>}
          </div>
        );
      })}
    </div>
  );
}

function ShippingLogTab({parts,shippingLogs,assemblyLogs,job,user,canAdmin,onRefresh}){
  const [showForm,setShowForm]=useState(false);
  const [f,setF]=useState({part_id:"",qty_shipped:"",ship_date:today(),customer:job.customer||"",customer_po:"",bol_number:"",entered_by:user.name,notes:""});
  const [saving,setSaving]=useState(false);
  const set=(k,v)=>setF(x=>({...x,[k]:v}));

  const totalShipped=shippingLogs.reduce((s,a)=>s+(a.qty_shipped||0),0);
  const totalCompleted=assemblyLogs.reduce((s,a)=>s+(a.qty_completed||0),0);
  const onHand=totalCompleted-totalShipped;

  async function save(){
    if(!f.qty_shipped||!f.part_id)return;
    setSaving(true);
    try{
      await API.mfg.shippingLog.create({...f,job_id:job.id,qty_shipped:parseInt(f.qty_shipped)||0});
      await API.mfg.parts.update(f.part_id,{qty_shipped:(parts.find(p=>p.id===f.part_id)?.qty_shipped||0)+(parseInt(f.qty_shipped)||0)});
      setShowForm(false);setF({part_id:"",qty_shipped:"",ship_date:today(),customer:job.customer||"",customer_po:"",bol_number:"",entered_by:user.name,notes:""});
      await onRefresh();
    }catch(e){alert("Error: "+e.message);}
    setSaving(false);
  }

  return(
    <div>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:8,marginBottom:14}}>
        {[[totalCompleted,"Completed",T.green],[totalShipped,"Shipped",T.blue],[onHand,"On Hand",onHand<0?T.red:T.orange]].map(([v,l,c])=>(
          <div key={l} style={{...cardS,textAlign:"center"}}><div style={{fontSize:22,fontWeight:900,color:c}}>{v}</div><div style={{fontSize:10,color:T.muted,textTransform:"uppercase"}}>{l}</div></div>
        ))}
      </div>

      <button onClick={()=>setShowForm(s=>!s)} style={{...primBtn,borderRadius:14,marginBottom:14,background:T.blue}}>📤 Log Shipment</button>

      {showForm&&<div style={{...cardS,marginBottom:14,border:`1px solid ${T.blue}40`}}>
        <div style={{fontSize:13,fontWeight:800,color:T.blue,marginBottom:12}}>📤 Log Assembly Shipment</div>
        <div style={{marginBottom:10}}><label style={lbl}>Finished Product Part # *</label>
          <select value={f.part_id} onChange={e=>set("part_id",e.target.value)} style={inpSel}>
            <option value="">— Select part —</option>
            {parts.map(p=><option key={p.id} value={p.id}>{p.part_number}{p.description?" — "+p.description:""}</option>)}
          </select>
        </div>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:10}}>
          <div><label style={lbl}>Qty Shipped *</label><input type="number" value={f.qty_shipped} onChange={e=>set("qty_shipped",e.target.value)} placeholder="e.g. 40" style={inp}/></div>
          <div><label style={lbl}>Ship Date</label><input type="date" value={f.ship_date} onChange={e=>set("ship_date",e.target.value)} style={inp}/></div>
        </div>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:10}}>
          <div><label style={lbl}>Customer</label><input value={f.customer} onChange={e=>set("customer",e.target.value)} style={inp}/></div>
          <div><label style={lbl}>Customer PO</label><input value={f.customer_po} onChange={e=>set("customer_po",e.target.value)} placeholder="PO #" style={inp}/></div>
        </div>
        <div style={{marginBottom:10}}><label style={lbl}>BOL / Packing Slip #</label><input value={f.bol_number} onChange={e=>set("bol_number",e.target.value)} placeholder="BOL #" style={inp}/></div>
        <div style={{display:"flex",gap:8}}>
          <button onClick={save} disabled={!f.qty_shipped||!f.part_id||saving} style={{...primBtn,flex:2,borderRadius:12,background:T.blue,opacity:f.qty_shipped&&f.part_id&&!saving?1:0.5}}>{saving?"Saving…":"Log Shipment"}</button>
          <button onClick={()=>setShowForm(false)} style={{...ghostBtn,flex:1,textAlign:"center"}}>Cancel</button>
        </div>
      </div>}

      {shippingLogs.length===0&&<div style={{textAlign:"center",padding:"40px 16px",color:T.muted}}>
        <div style={{fontSize:44,marginBottom:12}}>📤</div>
        <div style={{fontSize:14,fontWeight:700,color:T.sub,marginBottom:6}}>No Shipments Logged</div>
        <div style={{fontSize:12}}>Log shipments here to track completed assemblies leaving the shop.</div>
      </div>}

      {shippingLogs.map(s=>{
        const part=parts.find(p=>p.id===s.part_id);
        return(
          <div key={s.id} style={{...cardS,marginBottom:8,borderLeft:`3px solid ${T.blue}`}}>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start"}}>
              <div>
                <div style={{fontSize:16,fontWeight:900,color:T.blue}}>{s.qty_shipped} shipped</div>
                <div style={{fontSize:12,color:T.sub}}>{s.ship_date} · {part?.part_number||"—"} · {s.customer||""}</div>
                {s.bol_number&&<div style={{fontSize:11,color:T.muted}}>BOL: {s.bol_number}</div>}
                {s.customer_po&&<div style={{fontSize:11,color:T.muted}}>PO: {s.customer_po}</div>}
              </div>
              <div style={{display:"flex",gap:6,alignItems:"center"}}>
                <div style={{fontSize:20}}>📤</div>
                {canAdmin&&<button onClick={async()=>{if(window.confirm("Delete this shipment entry?"))try{await sb(`/mfg_shipping_log?id=eq.${s.id}`,{method:"DELETE"});await onRefresh();}catch(e){}}} style={{background:"none",border:`1px solid ${T.red}40`,borderRadius:8,padding:"4px 8px",color:T.red,fontSize:11,cursor:"pointer",fontFamily:"inherit"}}>🗑</button>}
              </div>
            </div>
          </div>
        );
      })}
    </div>
  );
}

function ManufacturingTraveler({part,job,user,onBack}){
  const [logs,setLogs]=useState([]);
  const [loading,setLoading]=useState(true);
  const [openStage,setOpenStage]=useState(null); // which stage form is open
  const [sf,setSf]=useState({});
  const [saving,setSaving]=useState(false);
  const canAdmin=user.role==="admin"||user.role==="pm"||user.role==="foreman";

  const STAGES=[
    {id:"mat_received",   label:"Material Received",   icon:"📦",color:"#60A5FA", fields:[
      {k:"qty",l:"Qty Received *",t:"number",ph:"0"},
      {k:"heat_number",l:"Heat / Lot #",t:"text",ph:"Heat #"},
      {k:"entered_by",l:"Received By",t:"text",ph:user.name,def:user.name},
      {k:"notes",l:"Notes",t:"text",ph:"Condition, reference, etc."},
    ]},
    {id:"mat_inspection", label:"Material Inspection", icon:"🔍",color:"#FBBF24", fields:[
      {k:"qty",l:"Qty Inspected *",t:"number",ph:"0"},
      {k:"pass_fail",l:"Result *",t:"select",opts:["Pass","Fail"]},
      {k:"entered_by",l:"Inspector",t:"text",ph:user.name,def:user.name},
      {k:"notes",l:"Defects / Notes",t:"text",ph:"Describe any issues"},
    ]},
    {id:"tacked",         label:"Tacked in Jig",       icon:"🔩",color:"#F97316", fields:[
      {k:"qty",l:"Qty Tacked *",t:"number",ph:"0"},
      {k:"worker_name",l:"Tacker Name",t:"text",ph:user.name,def:user.name},
      {k:"jig_number",l:"Jig Number",t:"text",ph:"Jig #"},
      {k:"notes",l:"Notes",t:"text",ph:""},
    ]},
    {id:"welded",         label:"Fully Welded",         icon:"🔥",color:"#EF4444", fields:[
      {k:"qty",l:"Qty Welded *",t:"number",ph:"0"},
      {k:"worker_name",l:"Welder Name",t:"text",ph:user.name,def:user.name},
      {k:"notes",l:"Notes",t:"text",ph:""},
    ]},
    {id:"welder_qc",      label:"Welder QC 🟡",         icon:"🟡",color:"#FCD34D", fields:[
      {k:"qty",l:"Qty QC'd *",t:"number",ph:"0"},
      {k:"worker_name",l:"Welder (Self-QC)",t:"text",ph:user.name,def:user.name},
      {k:"pass_fail",l:"Result *",t:"select",opts:["Pass","Fail"]},
      {k:"notes",l:"Notes / Rejects",t:"text",ph:"Any issues found"},
    ]},
    {id:"manager_qc",     label:"Manager QC ⚪",        icon:"⚪",color:"#E2E8F0", fields:[
      {k:"qty",l:"Qty QC'd *",t:"number",ph:"0"},
      {k:"worker_name",l:"Inspector / Manager",t:"text",ph:user.name,def:user.name},
      {k:"pass_fail",l:"Result *",t:"select",opts:["Pass","Fail"]},
      {k:"notes",l:"Notes / Rejects",t:"text",ph:"Any issues found"},
    ]},
    {id:"shipped",        label:"Palletized & Shipped", icon:"📦",color:"#34D399", fields:[
      {k:"qty",l:"Qty Shipped *",t:"number",ph:"0"},
      {k:"pallet_number",l:"Pallet #",t:"text",ph:"Pallet or skid #"},
      {k:"entered_by",l:"Packed By",t:"text",ph:user.name,def:user.name},
      {k:"notes",l:"Notes",t:"text",ph:"Banding confirmed, etc."},
    ]},
  ];

  useEffect(()=>{load();},[part.id]);
  async function load(){
    setLoading(true);
    try{const r=await API.mfg.stageLog.forPart(part.id);setLogs(Array.isArray(r)?r:[]);}
    catch(e){console.error(e);}
    setLoading(false);
  }

  function getStageTotal(stageId){
    return logs.filter(l=>l.stage===stageId).reduce((s,l)=>s+(l.qty||0),0);
  }

  async function saveEntry(){
    if(!sf.qty||!openStage)return;
    setSaving(true);
    try{
      await API.mfg.stageLog.create({
        part_id:part.id,
        job_id:job.id,
        stage:openStage,
        log_date:sf.log_date||today(),
        qty:parseInt(sf.qty)||0,
        entered_by:sf.entered_by||user.name,
        worker_name:sf.worker_name||null,
        pass_fail:sf.pass_fail||null,
        jig_number:sf.jig_number||null,
        heat_number:sf.heat_number||null,
        pallet_number:sf.pallet_number||null,
        notes:sf.notes||null,
      });
      setOpenStage(null);setSf({});await load();
    }catch(e){alert("Error: "+e.message);}
    setSaving(false);
  }

  async function deleteEntry(id){
    if(!window.confirm("Delete this log entry?"))return;
    try{await API.mfg.stageLog.remove(id);await load();}catch(e){}
  }

  const ordered=part.qty_ordered||0;

  return(
    <div style={{background:T.bg,minHeight:"100vh",fontFamily:"inherit"}}>
      {/* Header */}
      <div style={{background:T.surface,borderBottom:`1px solid ${T.border}`,padding:"14px 16px",position:"sticky",top:0,zIndex:50}}>
        <button onClick={onBack} style={{background:"none",border:"none",color:T.sub,fontSize:13,cursor:"pointer",fontFamily:"inherit",marginBottom:6}}>← {job.job_number}</button>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start"}}>
          <div>
            <div style={{fontSize:18,fontWeight:900,color:T.purple}}>{part.part_number}</div>
            <div style={{fontSize:12,color:T.sub}}>{part.description||""}{part.drawing_number?" · DWG: "+part.drawing_number:""}</div>
          </div>
          <div style={{textAlign:"right"}}>
            <div style={{fontSize:18,fontWeight:900,color:T.text}}>{ordered}</div>
            <div style={{fontSize:9,color:T.muted,textTransform:"uppercase"}}>Ordered</div>
          </div>
        </div>
      </div>

      <div style={{padding:"14px 16px 80px"}}>
        {loading&&<Spinner/>}

        {/* Stage progress summary row */}
        {!loading&&<div style={{...cardS,marginBottom:16}}>
          <div style={{fontSize:11,fontWeight:700,color:T.muted,textTransform:"uppercase",letterSpacing:"1px",marginBottom:10}}>Production Progress — {ordered} ordered</div>
          <div style={{display:"grid",gridTemplateColumns:"repeat(7,1fr)",gap:4}}>
            {STAGES.map(s=>{
              const total=getStageTotal(s.id);
              const pct=ordered>0?Math.min(100,(total/ordered)*100):0;
              return(
                <div key={s.id} style={{textAlign:"center"}}>
                  <div style={{fontSize:16,marginBottom:2}}>{s.icon}</div>
                  <div style={{fontSize:13,fontWeight:900,color:total>0?s.color:T.muted}}>{total}</div>
                  <div style={{background:T.border,borderRadius:3,height:4,margin:"3px 0"}}>
                    <div style={{height:4,borderRadius:3,background:s.color,width:`${pct}%`,transition:"width 0.3s"}}/>
                  </div>
                  <div style={{fontSize:7,color:T.muted,lineHeight:1.2}}>{s.label.split(" ").slice(0,2).join(" ")}</div>
                </div>
              );
            })}
          </div>
        </div>}

        {/* Stage cards — each is independent, no locking */}
        {!loading&&STAGES.map(stage=>{
          const stageLogs=logs.filter(l=>l.stage===stage.id);
          const total=stageLogs.reduce((s,l)=>s+(l.qty||0),0);
          const isOpen=openStage===stage.id;
          const passes=stageLogs.filter(l=>l.pass_fail==="Pass").reduce((s,l)=>s+(l.qty||0),0);
          const fails=stageLogs.filter(l=>l.pass_fail==="Fail").reduce((s,l)=>s+(l.qty||0),0);

          return(
            <div key={stage.id} style={{...cardS,marginBottom:10,borderLeft:`4px solid ${total>0?stage.color:T.border}`}}>
              {/* Stage header */}
              <div style={{display:"flex",alignItems:"center",gap:10,marginBottom:10}}>
                <div style={{width:38,height:38,borderRadius:10,background:`${stage.color}20`,border:`2px solid ${stage.color}60`,display:"flex",alignItems:"center",justifyContent:"center",fontSize:18,flexShrink:0}}>
                  {stage.icon}
                </div>
                <div style={{flex:1}}>
                  <div style={{fontSize:14,fontWeight:800,color:T.text}}>{stage.label}</div>
                  <div style={{display:"flex",gap:10,marginTop:2}}>
                    <span style={{fontSize:12,color:stage.color,fontWeight:700}}>{total} logged</span>
                    {ordered>0&&<span style={{fontSize:11,color:T.muted}}>of {ordered} ordered ({Math.round((total/ordered)*100)}%)</span>}
                    {stage.fields.some(f=>f.k==="pass_fail")&&total>0&&<>
                      <span style={{fontSize:11,color:T.green,fontWeight:700}}>✓ {passes}</span>
                      {fails>0&&<span style={{fontSize:11,color:T.red,fontWeight:700}}>✗ {fails}</span>}
                    </>}
                  </div>
                </div>
                {canAdmin&&<button onClick={()=>{
                    if(isOpen){setOpenStage(null);setSf({});}
                    else{
                      const defaults={log_date:today()};
                      stage.fields.forEach(f=>{if(f.def)defaults[f.k]=f.def;if(f.t==="select"&&f.opts)defaults[f.k]=f.opts[0];});
                      setSf(defaults);setOpenStage(stage.id);
                    }
                  }}
                  style={{background:isOpen?T.surface:stage.color,color:isOpen?T.muted:"#000",border:`1px solid ${stage.color}60`,borderRadius:10,padding:"6px 12px",fontSize:12,fontWeight:700,cursor:"pointer",fontFamily:"inherit",flexShrink:0}}>
                  {isOpen?"✕ Cancel":"+ Log"}
                </button>}
              </div>

              {/* Log entry form */}
              {isOpen&&<div style={{background:T.surface,borderRadius:10,padding:12,marginBottom:10,border:`1px solid ${stage.color}40`}}>
                <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8,marginBottom:8}}>
                  <div><label style={lbl}>Date</label><input type="date" value={sf.log_date||today()} onChange={e=>setSf(x=>({...x,log_date:e.target.value}))} style={inp}/></div>
                  {stage.fields.filter(f=>f.k!=="notes").map(f=>(
                    <div key={f.k}>
                      <label style={lbl}>{f.l}</label>
                      {f.t==="select"
                        ?<select value={sf[f.k]||f.opts?.[0]||""} onChange={e=>setSf(x=>({...x,[f.k]:e.target.value}))} style={inpSel}>{(f.opts||[]).map(o=><option key={o}>{o}</option>)}</select>
                        :<input type={f.t||"text"} value={sf[f.k]||""} onChange={e=>setSf(x=>({...x,[f.k]:e.target.value}))} placeholder={f.ph} style={inp}/>}
                    </div>
                  ))}
                </div>
                <div style={{marginBottom:10}}><label style={lbl}>Notes</label><input value={sf.notes||""} onChange={e=>setSf(x=>({...x,notes:e.target.value}))} placeholder="Optional notes" style={inp}/></div>
                <button onClick={saveEntry} disabled={!sf.qty||saving}
                  style={{...primBtn,borderRadius:12,background:stage.color,color:"#000",opacity:sf.qty&&!saving?1:0.5,width:"100%"}}>
                  {saving?"Saving…":`✓ Save ${stage.label} Entry`}
                </button>
              </div>}

              {/* Log entries */}
              {stageLogs.length>0&&<div>
                <div style={{fontSize:10,fontWeight:700,color:T.muted,textTransform:"uppercase",letterSpacing:"1px",marginBottom:6}}>Log ({stageLogs.length} entries)</div>
                {stageLogs.map(entry=>(
                  <div key={entry.id} style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",padding:"7px 0",borderBottom:`1px solid ${T.border}`}}>
                    <div style={{flex:1}}>
                      <div style={{display:"flex",gap:8,alignItems:"center",flexWrap:"wrap"}}>
                        <span style={{fontSize:13,fontWeight:800,color:stage.color}}>{entry.qty} {entry.qty===1?"unit":"units"}</span>
                        <span style={{fontSize:11,color:T.muted}}>{entry.log_date}</span>
                        {entry.worker_name&&<span style={{fontSize:11,color:T.sub}}>· {entry.worker_name}</span>}
                        {entry.entered_by&&entry.entered_by!==entry.worker_name&&<span style={{fontSize:11,color:T.muted}}>· {entry.entered_by}</span>}
                        {entry.pass_fail&&<span style={{fontSize:11,fontWeight:700,color:entry.pass_fail==="Pass"?T.green:T.red}}>· {entry.pass_fail}</span>}
                        {entry.jig_number&&<span style={{fontSize:11,color:T.muted}}>· Jig {entry.jig_number}</span>}
                        {entry.heat_number&&<span style={{fontSize:11,color:T.muted}}>· Heat {entry.heat_number}</span>}
                        {entry.pallet_number&&<span style={{fontSize:11,color:T.muted}}>· Pallet {entry.pallet_number}</span>}
                      </div>
                      {entry.notes&&<div style={{fontSize:11,color:T.muted,fontStyle:"italic",marginTop:2}}>{entry.notes}</div>}
                    </div>
                    {canAdmin&&<button onClick={()=>deleteEntry(entry.id)} style={{background:"none",border:"none",color:T.red,cursor:"pointer",fontSize:13,padding:"0 4px",flexShrink:0}}>🗑</button>}
                  </div>
                ))}
              </div>}

              {stageLogs.length===0&&!isOpen&&<div style={{fontSize:12,color:T.muted,fontStyle:"italic",textAlign:"center",padding:"8px 0"}}>
                No entries yet — tap <strong style={{color:stage.color}}>+ Log</strong> to record production
              </div>}
            </div>
          );
        })}
      </div>
    </div>
  );
}

function LaborTab({job,parts,labor,user,canAdmin,onRefresh}){
  const [showForm,setShowForm]=useState(false);
  const [f,setF]=useState({worker_name:user.name,work_date:today(),hours:"",part_id:"",operation:"",notes:""});
  const [saving,setSaving]=useState(false);
  const set=(k,v)=>setF(x=>({...x,[k]:v}));
  const OPERATIONS=["Material Received","Material Inspection","Tacking","Welding","Welder QC","Manager QC","Palletizing/Shipping","General/Other"];

  async function save(){
    if(!f.worker_name.trim()||!f.hours)return;
    setSaving(true);
    try{
      await API.mfg.labor.create({...f,job_id:job.id,hours:parseFloat(f.hours),part_id:f.part_id||null});
      setShowForm(false);setF({worker_name:user.name,work_date:today(),hours:"",part_id:"",operation:"",notes:""});
      await onRefresh();
    }catch(e){alert(e.message);}
    setSaving(false);
  }

  const byWorker={};
  labor.forEach(l=>{
    const n=l.worker_name;
    if(!byWorker[n])byWorker[n]={name:n,total:0,entries:[]};
    byWorker[n].total+=parseFloat(l.hours)||0;
    byWorker[n].entries.push(l);
  });
  const totalHrs=labor.reduce((s,l)=>s+(parseFloat(l.hours)||0),0);

  return(
    <div>
      <button onClick={()=>setShowForm(s=>!s)} style={{...primBtn,borderRadius:14,marginBottom:14,background:T.blue}}>+ Log Hours</button>

      {showForm&&<div style={{...cardS,marginBottom:14,border:`1px solid ${T.blue}40`}}>
        <div style={{fontSize:13,fontWeight:800,color:T.blue,marginBottom:12}}>⏱ Log Labor Hours</div>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:10}}>
          <div><label style={lbl}>Worker *</label><input value={f.worker_name} onChange={e=>set("worker_name",e.target.value)} style={inp}/></div>
          <div><label style={lbl}>Date</label><input type="date" value={f.work_date} onChange={e=>set("work_date",e.target.value)} style={inp}/></div>
        </div>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:10}}>
          <div><label style={lbl}>Hours *</label><input type="number" step="0.5" value={f.hours} onChange={e=>set("hours",e.target.value)} placeholder="8.0" style={inp}/></div>
          <div><label style={lbl}>Part #</label>
            <select value={f.part_id} onChange={e=>set("part_id",e.target.value)} style={inpSel}>
              <option value="">— All / General —</option>
              {parts.map(p=><option key={p.id} value={p.id}>{p.part_number}</option>)}
            </select>
          </div>
        </div>
        <div style={{marginBottom:10}}><label style={lbl}>Operation</label>
          <select value={f.operation} onChange={e=>set("operation",e.target.value)} style={inpSel}>
            <option value="">— Select operation —</option>
            {OPERATIONS.map(o=><option key={o}>{o}</option>)}
          </select>
        </div>
        <div style={{marginBottom:12}}><label style={lbl}>Notes</label><input value={f.notes} onChange={e=>set("notes",e.target.value)} placeholder="Optional notes" style={inp}/></div>
        <div style={{display:"flex",gap:8}}>
          <button onClick={save} disabled={!f.worker_name.trim()||!f.hours||saving} style={{...primBtn,flex:2,borderRadius:12,background:T.blue,opacity:f.worker_name.trim()&&f.hours&&!saving?1:0.5}}>{saving?"Saving…":"Save Hours"}</button>
          <button onClick={()=>setShowForm(false)} style={{...ghostBtn,flex:1,textAlign:"center"}}>Cancel</button>
        </div>
      </div>}

      {/* Summary */}
      {labor.length>0&&<div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8,marginBottom:14}}>
        <div style={{...cardS,textAlign:"center"}}><div style={{fontSize:24,fontWeight:900,color:T.blue}}>{totalHrs.toFixed(1)}</div><div style={{fontSize:10,color:T.muted,textTransform:"uppercase"}}>Total Hours</div></div>
        <div style={{...cardS,textAlign:"center"}}><div style={{fontSize:24,fontWeight:900,color:T.purple}}>{Object.keys(byWorker).length}</div><div style={{fontSize:10,color:T.muted,textTransform:"uppercase"}}>Workers</div></div>
      </div>}

      {/* By worker */}
      {Object.values(byWorker).sort((a,b)=>b.total-a.total).map(w=>(
        <div key={w.name} style={{...cardS,marginBottom:10}}>
          <div style={{display:"flex",justifyContent:"space-between",marginBottom:8}}>
            <div style={{fontSize:14,fontWeight:800,color:T.blue}}>{w.name}</div>
            <div style={{fontSize:16,fontWeight:900,color:T.green}}>{w.total.toFixed(1)}h</div>
          </div>
          {w.entries.map(e=>{
            const p=parts.find(x=>x.id===e.part_id);
            return(
              <div key={e.id} style={{display:"flex",justifyContent:"space-between",padding:"6px 0",borderBottom:`1px solid ${T.border}`,fontSize:12}}>
                <div>
                  <span style={{color:T.sub}}>{e.work_date}</span>
                  {p&&<span style={{color:T.purple,marginLeft:8}}>· {p.part_number}</span>}
                  {e.operation&&<span style={{color:T.muted,marginLeft:8}}>· {e.operation}</span>}
                </div>
                <span style={{color:T.green,fontWeight:700}}>{parseFloat(e.hours).toFixed(1)}h</span>
              </div>
            );
          })}
        </div>
      ))}

      {labor.length===0&&<div style={{textAlign:"center",padding:"40px 16px",color:T.muted}}>
        <div style={{fontSize:44,marginBottom:12}}>⏱</div>
        <div style={{fontSize:14,fontWeight:700,color:T.sub,marginBottom:6}}>No Labor Logged</div>
        <div style={{fontSize:12}}>Tap <strong style={{color:T.blue}}>+ Log Hours</strong> to track worker hours per operation.</div>
      </div>}
    </div>
  );
}

function NcrTab({job,parts,ncrs,user,canAdmin,onRefresh}){
  const [showForm,setShowForm]=useState(false);
  const [resolving,setResolving]=useState(null);
  const [f,setF]=useState({part_id:"",found_by:user.name,found_date:today(),stage:"",issue_desc:"",qty_affected:1,disposition:"Rework"});
  const [rf,setRf]=useState({resolved_by:user.name,resolved_date:today(),rework_hours:"",resolution_notes:""});
  const [saving,setSaving]=useState(false);
  const set=(k,v)=>setF(x=>({...x,[k]:v}));
  const setr=(k,v)=>setRf(x=>({...x,[k]:v}));
  const open=ncrs.filter(n=>n.status==="Open");
  const closed=ncrs.filter(n=>n.status==="Closed");

  async function saveNcr(){
    if(!f.issue_desc.trim()||!f.part_id)return;
    setSaving(true);
    try{
      await API.mfg.ncr.create({...f,job_id:job.id});
      setShowForm(false);setF({part_id:"",found_by:user.name,found_date:today(),stage:"",issue_desc:"",qty_affected:1,disposition:"Rework"});
      await onRefresh();
    }catch(e){alert(e.message);}
    setSaving(false);
  }
  async function resolve(ncr){
    setSaving(true);
    try{
      await API.mfg.ncr.update(ncr.id,{...rf,rework_hours:parseFloat(rf.rework_hours)||0,status:"Closed"});
      setResolving(null);await onRefresh();
    }catch(e){alert(e.message);}
    setSaving(false);
  }

  const STAGES=["Material Received","Material Inspection","Tacking","Welding","Welder QC","Manager QC","Palletizing"];

  return(
    <div>
      {open.length>0&&<div style={{background:T.redLow,border:`1px solid ${T.red}40`,borderRadius:12,padding:"10px 14px",marginBottom:14,fontSize:12,color:T.red,fontWeight:700}}>
        ⚠️ {open.length} Open NCR{open.length!==1?"s":""} — action required
      </div>}
      <button onClick={()=>setShowForm(s=>!s)} style={{...primBtn,borderRadius:14,marginBottom:14,background:T.red}}>+ Log NCR / Reject</button>

      {showForm&&<div style={{...cardS,marginBottom:14,border:`1px solid ${T.red}40`}}>
        <div style={{fontSize:13,fontWeight:800,color:T.red,marginBottom:12}}>❌ New Non-Conformance Report</div>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:10}}>
          <div><label style={lbl}>Part # *</label>
            <select value={f.part_id} onChange={e=>set("part_id",e.target.value)} style={inpSel}>
              <option value="">— Select part —</option>
              {parts.map(p=><option key={p.id} value={p.id}>{p.part_number}</option>)}
            </select>
          </div>
          <div><label style={lbl}>Found By</label><input value={f.found_by} onChange={e=>set("found_by",e.target.value)} style={inp}/></div>
        </div>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:10}}>
          <div><label style={lbl}>Stage Found</label>
            <select value={f.stage} onChange={e=>set("stage",e.target.value)} style={inpSel}>
              <option value="">— Select stage —</option>
              {STAGES.map(s=><option key={s}>{s}</option>)}
            </select>
          </div>
          <div><label style={lbl}>Qty Affected</label><input type="number" value={f.qty_affected} onChange={e=>set("qty_affected",e.target.value)} style={inp}/></div>
        </div>
        <div style={{marginBottom:10}}><label style={lbl}>Issue Description *</label><textarea value={f.issue_desc} onChange={e=>set("issue_desc",e.target.value)} rows={3} placeholder="Describe the non-conformance in detail" style={{...inp,resize:"vertical"}}/></div>
        <div style={{marginBottom:12}}><label style={lbl}>Disposition</label>
          <select value={f.disposition} onChange={e=>set("disposition",e.target.value)} style={inpSel}>
            {["Rework","Scrap","Accept-As-Is"].map(d=><option key={d}>{d}</option>)}
          </select>
        </div>
        <div style={{display:"flex",gap:8}}>
          <button onClick={saveNcr} disabled={!f.issue_desc.trim()||!f.part_id||saving} style={{...primBtn,flex:2,borderRadius:12,background:T.red,opacity:f.issue_desc.trim()&&f.part_id&&!saving?1:0.5}}>{saving?"Saving…":"Log NCR"}</button>
          <button onClick={()=>setShowForm(false)} style={{...ghostBtn,flex:1,textAlign:"center"}}>Cancel</button>
        </div>
      </div>}

      {/* Open NCRs */}
      {open.map(ncr=>{
        const p=parts.find(x=>x.id===ncr.part_id);
        return(
          <div key={ncr.id} style={{...cardS,marginBottom:10,borderLeft:`4px solid ${T.red}`}}>
            <div style={{display:"flex",justifyContent:"space-between",marginBottom:6}}>
              <div>
                <div style={{fontSize:13,fontWeight:800,color:T.red}}>❌ NCR — {p?.part_number||"—"}</div>
                <div style={{fontSize:11,color:T.muted}}>{ncr.found_date} · Found by {ncr.found_by}{ncr.stage?" at "+ncr.stage:""}</div>
              </div>
              <span style={{...pill(T.red),height:"fit-content"}}>OPEN</span>
            </div>
            <div style={{fontSize:13,color:T.sub,marginBottom:8,lineHeight:1.5}}>{ncr.issue_desc}</div>
            <div style={{display:"flex",gap:6,marginBottom:8,flexWrap:"wrap"}}>
              <span style={{...pill(T.yellow)}}>Qty: {ncr.qty_affected}</span>
              <span style={{...pill(ncr.disposition==="Scrap"?T.red:T.orange)}}>{ncr.disposition}</span>
            </div>
            {resolving===ncr.id?<div style={{background:T.surface,borderRadius:10,padding:12,border:`1px solid ${T.green}30`}}>
              <div style={{fontSize:12,fontWeight:800,color:T.green,marginBottom:10}}>✓ Resolve NCR</div>
              <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8,marginBottom:8}}>
                <div><label style={lbl}>Resolved By</label><input value={rf.resolved_by} onChange={e=>setr("resolved_by",e.target.value)} style={inp}/></div>
                <div><label style={lbl}>Rework Hours</label><input type="number" step="0.5" value={rf.rework_hours} onChange={e=>setr("rework_hours",e.target.value)} placeholder="0" style={inp}/></div>
              </div>
              <div style={{marginBottom:10}}><label style={lbl}>Resolution Notes</label><textarea value={rf.resolution_notes} onChange={e=>setr("resolution_notes",e.target.value)} rows={2} style={{...inp,resize:"vertical"}} placeholder="What was done to resolve it"/></div>
              <div style={{display:"flex",gap:8}}>
                <button onClick={()=>resolve(ncr)} style={{...primBtn,flex:2,borderRadius:10,background:T.green,color:"#000"}}>{saving?"Saving…":"Mark Resolved"}</button>
                <button onClick={()=>setResolving(null)} style={{...ghostBtn,flex:1,textAlign:"center"}}>Cancel</button>
              </div>
            </div>:
            <button onClick={()=>setResolving(ncr.id)} style={{...ghostBtn,width:"100%",textAlign:"center",color:T.green,border:`1px solid ${T.green}40`,fontSize:13}}>✓ Resolve NCR</button>}
          </div>
        );
      })}

      {/* Closed NCRs */}
      {closed.length>0&&<>
        <div style={{fontSize:11,fontWeight:700,color:T.muted,textTransform:"uppercase",letterSpacing:"1px",margin:"14px 0 8px"}}>Closed ({closed.length})</div>
        {closed.map(ncr=>{
          const p=parts.find(x=>x.id===ncr.part_id);
          return(
            <div key={ncr.id} style={{...cardS,marginBottom:8,opacity:0.7,borderLeft:`4px solid ${T.green}`}}>
              <div style={{display:"flex",justifyContent:"space-between"}}>
                <div><div style={{fontSize:12,fontWeight:700,color:T.green}}>✓ {p?.part_number||"—"}</div>
                  <div style={{fontSize:11,color:T.muted}}>{ncr.issue_desc?.slice(0,60)}…</div></div>
                <span style={{...pill(T.green),height:"fit-content",fontSize:9}}>CLOSED</span>
              </div>
              {ncr.resolution_notes&&<div style={{fontSize:11,color:T.muted,marginTop:4}}>Resolution: {ncr.resolution_notes}</div>}
            </div>
          );
        })}
      </>}

      {ncrs.length===0&&<div style={{textAlign:"center",padding:"40px 16px",color:T.muted}}>
        <div style={{fontSize:44,marginBottom:12}}>✅</div>
        <div style={{fontSize:14,fontWeight:700,color:T.sub,marginBottom:6}}>No NCRs Logged</div>
        <div style={{fontSize:12}}>Non-conformances, rejects, and rework are logged here. Quality issues from QC stages appear here automatically.</div>
      </div>}
    </div>
  );
}

function ManufacturingDashboard({jobs,user,onSelectJob}){
  const [allParts,setAllParts]=useState([]);
  const [travelers,setTravelers]=useState({});
  const [ncrs,setNcrs]=useState([]);
  const [loading,setLoading]=useState(true);

  useEffect(()=>{load();},[jobs]);
  async function load(){
    setLoading(true);
    try{
      const partArrays=await Promise.all(jobs.map(j=>API.mfg.parts.forJob(j.id).catch(()=>[])));
      const parts=partArrays.flat().map((p,_,arr)=>{
        const job=jobs.find(j=>j.id===p.job_id);
        return{...p,jobName:job?.job_number||"",customer:job?.customer||"",due_date:job?.due_date};
      });
      setAllParts(parts);
      const travMap={};
      const ncrAll=[];
      await Promise.all(parts.map(async p=>{
        const [t,n]=await Promise.all([API.mfg.travelers.forPart(p.id).catch(()=>[]),API.mfg.ncr.forPart(p.id).catch(()=>[])]);
        travMap[p.id]=Array.isArray(t)&&t.length>0?t[0]:null;
        if(Array.isArray(n))ncrAll.push(...n);
      }));
      setTravelers(travMap);setNcrs(ncrAll);
    }catch(e){console.error("MFG error:",e.message||e);}
    setLoading(false);
  }

  if(loading)return<Spinner/>;

  const getStage=p=>travelers[p.id]?.current_stage||0;
  const notStarted=allParts.filter(p=>getStage(p)===0);
  const inProgress=allParts.filter(p=>getStage(p)>0&&getStage(p)<7);
  const complete=allParts.filter(p=>getStage(p)>=7);
  const openNcrs=ncrs.filter(n=>n.status==="Open");
  const totalOrdered=allParts.reduce((s,p)=>s+(p.qty_ordered||0),0);
  const totalShipped=allParts.reduce((s,p)=>s+(p.qty_shipped||0),0);

  return(
    <div style={{padding:"0 0 20px"}}>
      {/* Top stats */}
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8,marginBottom:16}}>
        {[
          [totalOrdered,"Total Ordered",T.blue],
          [totalShipped,"Shipped",T.green],
          [openNcrs.length,"Open NCRs",openNcrs.length>0?T.red:T.muted],
          [notStarted.length,"Awaiting Materials",notStarted.length>0?T.yellow:T.muted],
        ].map(([v,l,c])=>(
          <div key={l} style={{...cardS,textAlign:"center",padding:"12px 8px"}}>
            <div style={{fontSize:26,fontWeight:900,color:c}}>{v}</div>
            <div style={{fontSize:10,color:T.muted,textTransform:"uppercase",marginTop:2}}>{l}</div>
          </div>
        ))}
      </div>
      <div style={{...cardS,marginBottom:16}}>
        <div style={{display:"flex",justifyContent:"space-between",marginBottom:6}}>
          <span style={{fontSize:13,fontWeight:700,color:T.text}}>Production Output</span>
          <span style={{fontSize:12,color:T.green,fontWeight:700}}>{totalShipped}/{totalOrdered} shipped</span>
        </div>
        <div style={{background:T.border,borderRadius:6,height:12,overflow:"hidden"}}>
          <div style={{height:12,background:`linear-gradient(90deg,${T.green},${T.blue})`,width:`${totalOrdered>0?(totalShipped/totalOrdered)*100:0}%`,borderRadius:6,transition:"width 0.4s"}}/>
        </div>
        <div style={{display:"flex",justifyContent:"space-between",marginTop:8,fontSize:10,color:T.muted}}>
          <span>🟢 {complete.length} complete</span>
          <span>🟡 {inProgress.length} in progress</span>
          <span>⚪ {notStarted.length} not started</span>
        </div>
      </div>
      {openNcrs.length>0&&<div style={{...cardS,marginBottom:16,border:`1px solid ${T.red}40`}}>
        <div style={{fontSize:13,fontWeight:800,color:T.red,marginBottom:10}}>⚠️ Open NCRs — {openNcrs.length} need attention</div>
        {openNcrs.slice(0,5).map(n=>{
          const p=allParts.find(x=>x.id===n.part_id);
          return(<div key={n.id} style={{padding:"6px 0",borderBottom:`1px solid ${T.border}`,fontSize:12}}>
            <span style={{color:T.red,fontWeight:700}}>{p?.part_number||"—"}</span>
            <span style={{color:T.muted,marginLeft:8}}>{n.issue_desc?.slice(0,50)}…</span>
            <span style={{color:T.muted,marginLeft:8}}>· {n.disposition}</span>
          </div>);
        })}
      </div>}

      {/* Per-job output */}
      <div style={{fontSize:11,fontWeight:700,color:T.muted,textTransform:"uppercase",letterSpacing:"1px",marginBottom:8}}>Output by Job</div>
      {jobs.filter(j=>j.status==="active").map(job=>{
        const jobParts=allParts.filter(p=>p.job_id===job.id);
        const shipped=jobParts.reduce((s,p)=>s+(p.qty_shipped||0),0);
        const ordered=jobParts.reduce((s,p)=>s+(p.qty_ordered||0),0);
        const done=jobParts.filter(p=>getStage(p)>=7).length;
        const pct=ordered>0?(shipped/ordered)*100:0;
        const daysLeft=job.due_date?Math.ceil((new Date(job.due_date+"T12:00:00")-new Date())/86400000):null;
        return(
          <div key={job.id} style={{...cardS,marginBottom:8,cursor:"pointer"}} onClick={()=>onSelectJob(job)}>
            <div style={{display:"flex",justifyContent:"space-between",marginBottom:6}}>
              <div>
                <div style={{fontSize:14,fontWeight:800,color:T.purple}}>{job.job_number}</div>
                <div style={{fontSize:11,color:T.muted}}>{job.customer} · {jobParts.length} part{jobParts.length!==1?"s":""}</div>
              </div>
              <div style={{textAlign:"right"}}>
                <div style={{fontSize:14,fontWeight:800,color:T.green}}>{shipped}/{ordered}</div>
                <div style={{fontSize:9,color:T.muted}}>shipped</div>
              </div>
            </div>
            <div style={{background:T.border,borderRadius:4,height:6,marginBottom:4}}>
              <div style={{height:6,borderRadius:4,background:pct>=100?T.green:`linear-gradient(90deg,${T.purple},${T.blue})`,width:`${Math.min(100,pct)}%`}}/>
            </div>
            <div style={{display:"flex",justifyContent:"space-between",fontSize:10,color:T.muted}}>
              <span>{done}/{jobParts.length} parts done</span>
              {daysLeft!==null&&<span style={{color:daysLeft<0?T.red:daysLeft<=7?T.yellow:T.muted}}>{daysLeft<0?`${Math.abs(daysLeft)}d overdue`:daysLeft===0?"Due today":`${daysLeft}d left`}</span>}
            </div>
          </div>
        );
      })}
  </div>
  );
}

function PackingSlipScreen({job,parts,user,onBack,onSaved,existingSlip}){
  const [slipId,setSlipId]=useState(existingSlip?.id||null);
  const [slipNo,setSlipNo]=useState(existingSlip?.slip_number||existingSlip?.data?.slipNo||"");
  const [shipDate,setShipDate]=useState(existingSlip?.ship_date||existingSlip?.data?.shipDate||today());
  const [carrier,setCarrier]=useState(existingSlip?.data?.carrier||"Vendor Truck");
  const [truckTrailer,setTruckTrailer]=useState(existingSlip?.data?.truckTrailer||"");
  const [bolTracking,setBolTracking]=useState(existingSlip?.data?.bolTracking||"");
  const [skidCount,setSkidCount]=useState(existingSlip?.data?.skidCount||"");
  const [totalWeight,setTotalWeight]=useState(existingSlip?.data?.totalWeight||"");
  const [sealNo,setSealNo]=useState("");
  const [sqrs,setSqrs]=useState("SQR03, SQR04, SQR06, SQR33, SQR35");
  const [drawingRev,setDrawingRev]=useState("");
  const [deliveryAppt,setDeliveryAppt]=useState("Call 24 hrs. in advance: (814) 539-6922 x299");
  const [recvHours,setRecvHours]=useState("7:00 AM thru 3:00 PM ONLY");
  const [notes,setNotes]=useState(existingSlip?.data?.notes||"");

  const blankLine={poLine:"",partNo:"",description:"",qtyOrdered:"",qtyShipped:"",qtyBackordered:"",pkgSkid:"",notes:""};
  const [lines,setLines]=useState(()=>existingSlip?.data?.lines||parts.map(p=>({...blankLine,partNo:p.part_number,description:p.description||"",qtyOrdered:String(p.qty_ordered||"")})).concat(Array(Math.max(0,6-parts.length)).fill(blankLine)));
  const setLine=(i,k,v)=>setLines(ls=>ls.map((l,j)=>j===i?{...l,[k]:v}:l));

  const [checks,setChecks]=useState(existingSlip?.data?.checks||{
    packingSlip:false,shippingTicket:false,cofc:false,qcSheet:false,
    ctqLog:false,bolTicket:false,yellowRibbon:false,whiteRibbon:false,
    weatherProtect:false,damageProtect:false,noHold:false,qtyVerified:false,
  });
  const setCheck=(k)=>setChecks(c=>({...c,[k]:!c[k]}));
  const [saving,setSaving]=useState(false);
  const [saved,setSaved]=useState(!!existingSlip);
  const [saveMsg,setSaveMsg]=useState("");

  async function saveSlip(){
    setSaving(true);setSaveMsg("");
    const data={slipNo,shipDate,carrier,truckTrailer,bolTracking,skidCount,totalWeight,sealNo,sqrs,drawingRev,deliveryAppt,recvHours,notes,checks,lines};
    try{
      if(slipId){
        await API.mfg.packingSlips.update(slipId,{slip_number:slipNo,ship_date:shipDate||null,data,updated_at:new Date().toISOString()});
      }else{
        const r=await API.mfg.packingSlips.create({job_id:job.id,slip_number:slipNo,ship_date:shipDate||null,data,created_by:user.name});
        const newId=Array.isArray(r)?r[0]?.id:r?.id;
        if(newId)setSlipId(newId);
      }
      setSaved(true);setSaveMsg("✓ Saved");
      setTimeout(()=>setSaveMsg(""),3000);
      onSaved&&onSaved();
    }catch(e){setSaveMsg("Error: "+e.message);}
    setSaving(false);
  }

  function printSlip(){
    const CHECKMARK="✓";
    const CB=(checked)=>checked?`<span style="font-size:14px">☑</span>`:`<span style="font-size:14px">☐</span>`;

    const html=`<!DOCTYPE html><html><head><meta charset="utf-8">
<title>Packing Slip ${slipNo||""} — ${job.job_number}</title>
<style>
@page{size:letter portrait;margin:0.3in;}
*{box-sizing:border-box;margin:0;padding:0;font-family:Arial,sans-serif;font-size:7.5pt;}
body{color:#000;}
.header{display:grid;grid-template-columns:160px 1fr 130px 110px;gap:0;border:1.5px solid #1f3864;margin-bottom:4px;}
.logo-cell{padding:5px 6px;border-right:1px solid #1f3864;display:flex;flex-direction:column;justify-content:center;}
.logo-text{font-size:15pt;font-weight:900;color:#1f3864;letter-spacing:2px;}
.logo-sub{font-size:6pt;color:#444;margin-top:1px;}
.title-cell{background:#1f3864;color:#fff;display:flex;align-items:center;justify-content:center;font-size:17pt;font-weight:300;letter-spacing:3px;padding:6px;}
.slip-no-cell{border-left:1px solid #1f3864;padding:5px 8px;}
.ship-date-cell{border-left:1px solid #1f3864;padding:5px 8px;}
.field-label{font-size:6pt;font-weight:700;color:#555;text-transform:uppercase;margin-bottom:2px;}
.field-value{font-size:8.5pt;border-bottom:1px solid #000;min-height:14px;padding-bottom:1px;}
.section-header{background:#1f3864;color:#fff;font-weight:700;font-size:7pt;padding:2px 6px;text-transform:uppercase;letter-spacing:0.5px;margin:4px 0 2px 0;}
table{width:100%;border-collapse:collapse;}
th{background:#1f3864;color:#fff;padding:3px 4px;text-align:left;font-size:6.5pt;font-weight:700;}
td{padding:3px 4px;border:1px solid #ccc;font-size:7.5pt;min-height:16px;}
.info-row{display:grid;grid-template-columns:110px 1fr;border-bottom:1px solid #ccc;}
.info-row:last-child{border-bottom:none;}
.info-label{background:#f0f4ff;padding:3px 6px;font-weight:700;font-size:6.5pt;border-right:1px solid #ccc;}
.info-value{padding:3px 6px;font-size:7.5pt;}
.ship-row{border-bottom:1px solid #ccc;display:grid;grid-template-columns:100px 1fr;}
.ship-row:last-child{border-bottom:none;}
.ship-label{background:#f0f4ff;padding:2px 6px;font-weight:700;font-size:6.5pt;border-right:1px solid #ccc;}
.ship-value{padding:2px 6px;font-size:7.5pt;}
.check-grid{display:grid;grid-template-columns:1fr 1fr 1fr 1fr;gap:2px;padding:3px;border:1px solid #ccc;}
.check-item{font-size:7pt;display:flex;align-items:center;gap:3px;padding:1px;}
.sig-grid{display:grid;grid-template-columns:1fr 1fr 1fr 1fr;border:1px solid #ccc;}
.sig-cell{padding:5px 6px;border-right:1px solid #ccc;}
.sig-cell:last-child{border-right:none;}
.sig-line{border-bottom:1px solid #000;margin:12px 0 3px;min-height:22px;}
.sig-sub{font-size:6.5pt;color:#555;}
@media print{body{-webkit-print-color-adjust:exact;print-color-adjust:exact;}}
</style></head><body>

<!-- HEADER -->
<div class="header">
  <div class="logo-cell">
    <div class="logo-text">AIME</div>
    <div class="logo-sub">Atlantic Industrial Mechanical &amp; Environmental</div>
  </div>
  <div class="title-cell">PACKING SLIP</div>
  <div class="slip-no-cell">
    <div class="field-label">Packing Slip No.</div>
    <div class="field-value">${slipNo||"&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"}</div>
  </div>
  <div class="ship-date-cell">
    <div class="field-label">Ship Date</div>
    <div class="field-value">${shipDate||"____/____/______"}</div>
  </div>
</div>

<!-- SECTION 1: SHIP TO / ORDER INFO -->
<div class="section-header">1. Ship To / Order Information</div>
<div style="display:grid;grid-template-columns:1fr 1fr;border:1px solid #ccc;margin-bottom:4px;">
  <div style="border-right:1px solid #ccc;">
    ${[["From","AIME / Atlantic Welders Inc.<br/>5730 Pennington Ave, Baltimore, MD 21226"],
       ["Customer / Project",job.customer||"JWF Industries"],
       ["Buyer / Contact","JWFI Purchasing"],
       ["Ship Via",carrier||"Vendor Truck"]].map(([l,v])=>`
    <div class="info-row"><div class="info-label">${l}</div><div class="info-value">${v}</div></div>`).join("")}
  </div>
  <div>
    ${[["Ship To","JWFI<br/>84 Iron Street - Dock 2 Johnstown, PA 15906"],
       ["Customer PO #",job.po_number||"222577-00"],
       ["Delivery Appointment",deliveryAppt],
       ["Receiving Hours",recvHours]].map(([l,v])=>`
    <div class="info-row"><div class="info-label">${l}</div><div class="info-value">${v}</div></div>`).join("")}
  </div>
</div>

<!-- SECTION 2: SHIPMENT DETAILS -->
<div class="section-header">2. Shipment Details</div>
<div style="display:grid;grid-template-columns:1fr 1fr 1fr;border:1px solid #ccc;margin-bottom:4px;">
  ${[["Part Number",parts.map(p=>p.part_number).join(", ")||"1651"],
     ["Description",parts.map(p=>p.description).filter(Boolean).join(", ")||""],
     ["Drawing Rev.",drawingRev],
     ["Operation","OSP Fit / Weld"],
     ["UOM","EA"],
     ["SQRs",sqrs],
     ["Carrier / Driver",carrier],
     ["Truck / Trailer #",truckTrailer],
     ["BOL / Tracking #",bolTracking],
     ["Skid / Package Count",skidCount],
     ["Total Weight",totalWeight?(totalWeight+" lbs"):""],
     ["Seal #",sealNo]].map(([l,v])=>`
  <div class="ship-row"><div class="ship-label">${l}</div><div class="ship-value">${v||"&nbsp;"}</div></div>`).join("")}
</div>

<!-- SECTION 3: PARTS PACKED/SHIPPED -->
<div class="section-header">3. Parts Packed / Shipped</div>
<table style="margin-bottom:4px;">
  <thead><tr>
    <th style="width:30px">Line</th>
    <th>PO Line / WO #</th><th>Part No.</th><th>Description</th>
    <th style="text-align:center">Qty Ordered</th>
    <th style="text-align:center">Qty Shipped</th>
    <th style="text-align:center">Qty Backordered</th>
    <th>Package / Skid #</th><th>Notes</th>
  </tr></thead>
  <tbody>
    ${lines.slice(0,6).map((l,i)=>`
    <tr>
      <td style="text-align:center">${i+1}</td>
      <td>${l.poLine||"&nbsp;"}</td>
      <td>${l.partNo||"&nbsp;"}</td>
      <td>${l.description||"&nbsp;"}</td>
      <td style="text-align:center">${l.qtyOrdered||"&nbsp;"}</td>
      <td style="text-align:center">${l.qtyShipped||"&nbsp;"}</td>
      <td style="text-align:center">${l.qtyBackordered||"&nbsp;"}</td>
      <td>${l.pkgSkid||"&nbsp;"}</td>
      <td>${l.notes||"&nbsp;"}</td>
    </tr>`).join("")}
  </tbody>
</table>

<!-- SECTION 4: DOCUMENTS -->
<div class="section-header">4. Documents / Checks Included</div>
<div class="check-grid" style="margin-bottom:4px;">
  ${[["packingSlip","Packing slip attached"],["shippingTicket","Shipping ticket attached"],["cofc","CofC included"],["qcSheet","QC inspection sheet included"],
     ["ctqLog","CTQ / dimensional log included"],["bolTicket","BOL / delivery ticket included"],["yellowRibbon","Yellow ribbon attached"],["whiteRibbon","White ribbon attached"],
     ["weatherProtect","Parts protected from weather"],["damageProtect","Parts protected from damage"],["noHold","No HOLD material included"],["qtyVerified","Qty verified before loading"]
  ].map(([k,label])=>`<div class="check-item">${CB(checks[k])} ${label}</div>`).join("")}
</div>

<!-- SECTION 5: NOTES -->
<div class="section-header">5. Notes / Exceptions</div>
<div style="border:1px solid #ccc;min-height:50px;padding:8px;margin-bottom:2px;font-size:7.5pt;">${notes||"&nbsp;"}</div>

<!-- SECTION 6: SIGNATURES -->
<div class="section-header">6. Release / Receipt</div>
<div class="sig-grid">
  ${[["Packed / Prepared By",""],["QC Release",""],["Delivered By / Driver",""],["Received By / Customer",""]].map(([title])=>`
  <div class="sig-cell">
    <div style="font-size:7pt;font-weight:700;color:#1f3864;margin-bottom:4px;">${title}</div>
    <div class="sig-sub">Name / Signature:</div>
    <div class="sig-line"></div>
    <div class="sig-sub">Date / Time:</div>
    <div class="sig-line" style="margin-top:8px;"></div>
  </div>`).join("")}
</div>

<div style="text-align:center;margin-top:5px;font-size:6.5pt;color:#666;border-top:1px solid #eee;padding-top:3px;">
  Atlantic Industrial Mechanical &amp; Environmental | 5730 Pennington Ave, Baltimore, Maryland, 21226
</div>

</body></html>`;

    const win=window.open("","_blank","width=950,height=800");
    if(!win){alert("Popup blocked — please allow popups.");return;}
    win.document.write(html);
    win.document.close();
    win.focus();
    setTimeout(()=>{win.focus();win.print();},1200);
  }

  const CB_UI=({checked,onChange,label})=>(
    <label style={{display:"flex",alignItems:"center",gap:6,cursor:"pointer",fontSize:11,color:T.sub,padding:"4px 0"}}>
      <div onClick={onChange} style={{width:16,height:16,border:`2px solid ${checked?T.blue:T.border}`,borderRadius:3,background:checked?T.blue:T.surface,display:"flex",alignItems:"center",justifyContent:"center",cursor:"pointer",flexShrink:0}}>
        {checked&&<span style={{color:"#fff",fontSize:10,fontWeight:900}}>✓</span>}
      </div>
      {label}
    </label>
  );

  return(
    <div style={{background:T.bg,minHeight:"100vh",fontFamily:"inherit"}}>
      <div style={{background:T.surface,borderBottom:`1px solid ${T.border}`,padding:"12px 16px",position:"sticky",top:0,zIndex:50}}>
        <button onClick={onBack} style={{background:"none",border:"none",color:T.sub,fontSize:13,cursor:"pointer",fontFamily:"inherit",display:"block",marginBottom:4}}>← Back</button>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}>
          <div style={{fontSize:15,fontWeight:900,color:T.text}}>📄 Packing Slip</div>
          <div style={{display:"flex",gap:8,alignItems:"center"}}>
            {saveMsg&&<span style={{fontSize:11,color:saveMsg.startsWith("✓")?T.green:T.red,fontWeight:700}}>{saveMsg}</span>}
            <button onClick={saveSlip} disabled={saving}
              style={{background:saved?T.green:T.orange,color:"#000",border:"none",borderRadius:10,padding:"8px 16px",fontSize:12,fontWeight:700,cursor:"pointer",fontFamily:"inherit",opacity:saving?0.6:1}}>
              {saving?"Saving…":saved?"✓ Saved":"💾 Save"}
            </button>
            <button onClick={printSlip} style={{background:"#1f3864",color:"#fff",border:"none",borderRadius:10,padding:"8px 14px",fontSize:12,fontWeight:700,cursor:"pointer",fontFamily:"inherit"}}>
              🖨️ Print
            </button>
          </div>
        </div>
      </div>

      <div style={{padding:"14px 16px 100px"}}>
        {/* Header fields */}
        <div style={{...cardS,marginBottom:12}}>
          <div style={{fontSize:12,fontWeight:800,color:T.blue,marginBottom:12}}>📋 Slip Info</div>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}>
            <div><label style={lbl}>Packing Slip #</label><input value={slipNo} onChange={e=>setSlipNo(e.target.value)} placeholder="Slip #" style={inp}/></div>
            <div><label style={lbl}>Ship Date</label><input type="date" value={shipDate} onChange={e=>setShipDate(e.target.value)} style={inp}/></div>
          </div>
        </div>

        {/* Auto-filled from job */}
        <div style={{...cardS,marginBottom:12,background:T.blueLow,border:`1px solid ${T.blue}30`}}>
          <div style={{fontSize:12,fontWeight:800,color:T.blue,marginBottom:8}}>📦 Auto-filled from Job</div>
          {[["Customer / Project",job.customer||"JWF Industries"],["Customer PO #",job.po_number||""],["Part Number",parts.map(p=>p.part_number).join(", ")],["Description",parts.map(p=>p.description).filter(Boolean).join(", ")]].map(([l,v])=>(
            <div key={l} style={{display:"flex",justifyContent:"space-between",padding:"5px 0",borderBottom:`1px solid ${T.border}`,fontSize:12}}>
              <span style={{color:T.muted,fontWeight:700}}>{l}</span>
              <span style={{color:T.text}}>{v||"—"}</span>
            </div>
          ))}
        </div>

        {/* Shipment details */}
        <div style={{...cardS,marginBottom:12}}>
          <div style={{fontSize:12,fontWeight:800,color:T.text,marginBottom:12}}>🚛 Shipment Details</div>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:10}}>
            <div><label style={lbl}>Carrier / Driver</label><input value={carrier} onChange={e=>setCarrier(e.target.value)} style={inp}/></div>
            <div><label style={lbl}>Truck / Trailer #</label><input value={truckTrailer} onChange={e=>setTruckTrailer(e.target.value)} placeholder="Truck/Trailer #" style={inp}/></div>
          </div>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:10}}>
            <div><label style={lbl}>BOL / Tracking #</label><input value={bolTracking} onChange={e=>setBolTracking(e.target.value)} placeholder="BOL #" style={inp}/></div>
            <div><label style={lbl}>Skid / Package Count</label><input value={skidCount} onChange={e=>setSkidCount(e.target.value)} placeholder="e.g. 2 skids" style={inp}/></div>
          </div>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}>
            <div><label style={lbl}>Total Weight (lbs)</label><input type="number" value={totalWeight} onChange={e=>setTotalWeight(e.target.value)} placeholder="lbs" style={inp}/></div>
            <div><label style={lbl}>Seal #</label><input value={sealNo} onChange={e=>setSealNo(e.target.value)} placeholder="Seal #" style={inp}/></div>
          </div>
        </div>

        {/* Line items */}
        <div style={{...cardS,marginBottom:12}}>
          <div style={{fontSize:12,fontWeight:800,color:T.text,marginBottom:12}}>📦 Parts Packed / Shipped</div>
          {lines.slice(0,6).map((line,i)=>(
            <div key={i} style={{marginBottom:10,paddingBottom:10,borderBottom:`1px solid ${T.border}`}}>
              <div style={{fontSize:10,fontWeight:700,color:T.muted,marginBottom:6}}>Line {i+1}</div>
              <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8,marginBottom:6}}>
                <div><label style={lbl}>Part No.</label><input value={line.partNo} onChange={e=>setLine(i,"partNo",e.target.value)} placeholder="Part #" style={inp}/></div>
                <div><label style={lbl}>Description</label><input value={line.description} onChange={e=>setLine(i,"description",e.target.value)} placeholder="Description" style={inp}/></div>
              </div>
              <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr 1fr",gap:8}}>
                <div><label style={lbl}>PO/WO #</label><input value={line.poLine} onChange={e=>setLine(i,"poLine",e.target.value)} style={inp}/></div>
                <div><label style={lbl}>Qty Ordered</label><input type="number" value={line.qtyOrdered} onChange={e=>setLine(i,"qtyOrdered",e.target.value)} style={inp}/></div>
                <div><label style={lbl}>Qty Shipped</label><input type="number" value={line.qtyShipped} onChange={e=>setLine(i,"qtyShipped",e.target.value)} style={inp}/></div>
                <div><label style={lbl}>Pkg / Skid #</label><input value={line.pkgSkid} onChange={e=>setLine(i,"pkgSkid",e.target.value)} style={inp}/></div>
              </div>
            </div>
          ))}
        </div>

        {/* Checkboxes */}
        <div style={{...cardS,marginBottom:12}}>
          <div style={{fontSize:12,fontWeight:800,color:T.text,marginBottom:10}}>✅ Documents / Checks Included</div>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:4}}>
            {[["packingSlip","Packing slip attached"],["shippingTicket","Shipping ticket attached"],["cofc","CofC included"],["qcSheet","QC inspection sheet included"],
              ["ctqLog","CTQ / dimensional log included"],["bolTicket","BOL / delivery ticket included"],["yellowRibbon","🟡 Yellow ribbon attached"],["whiteRibbon","⚪ White ribbon attached"],
              ["weatherProtect","Parts protected from weather"],["damageProtect","Parts protected from damage"],["noHold","No HOLD material included"],["qtyVerified","Qty verified before loading"]
            ].map(([k,label])=>(
              <CB_UI key={k} checked={checks[k]} onChange={()=>setCheck(k)} label={label}/>
            ))}
          </div>
        </div>

        {/* Notes */}
        <div style={{...cardS,marginBottom:16}}>
          <div style={{fontSize:12,fontWeight:800,color:T.text,marginBottom:8}}>📝 Notes / Exceptions</div>
          <textarea value={notes} onChange={e=>setNotes(e.target.value)} rows={3} placeholder="Any notes or exceptions..." style={{...inp,resize:"vertical"}}/>
        </div>

        {/* Print button */}
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:8}}>
          <button onClick={saveSlip} disabled={saving} style={{...primBtn,borderRadius:14,background:saved?T.green:T.orange,color:"#000",opacity:saving?0.6:1,fontSize:14}}>
            {saving?"Saving…":saved?"✓ Saved":"💾 Save Slip"}
          </button>
          <button onClick={printSlip} style={{...primBtn,borderRadius:14,background:"#1f3864",fontSize:14}}>
            🖨️ Print
          </button>
        </div>
        <div style={{fontSize:11,color:T.muted,textAlign:"center"}}>
          Save to reference later · Print to sign · Upload signed copy to Docs
        </div>
      </div>
    </div>
  );
}

function AppInner(){
  const [publicRfiId]            = useState(()=>new URLSearchParams(window.location.search).get("rfi"));
  const [publicCoId]             = useState(()=>new URLSearchParams(window.location.search).get("co"));
  const [publicInspId]           = useState(()=>new URLSearchParams(window.location.search).get("inspect"));
  const [publicTMId]             = useState(()=>new URLSearchParams(window.location.search).get("tmsign"));
  const [user,setUser]           = useState(null);
  const [restoring,setRestoring] = useState(true);   // checking for a saved session
  const [projects,setProjects]   = useState([]);
  const [screen,setScreen]       = useState("division");
  const [selectedDiv,setSelectedDiv] = useState(null);
  const [selectedProject,setSelectedProject] = useState(null);
  const [selectedMfgJob,setSelectedMfgJob]   = useState(null);
  const [selectedMfgPart,setSelectedMfgPart] = useState(null);
  const [isOnline,setIsOnline]   = useState(navigator.onLine);
  const [pendingCount,setPendingCount] = useState(0);
  const [syncMsg,setSyncMsg]     = useState("");
  const [err,setErr]             = useState("");

  useEffect(()=>{
    let cancelled=false;
    (async()=>{
      const p=await restoreSession();
      if(cancelled)return;
      if(p)setUser(p);
      setRestoring(false);
    })();
    return()=>{cancelled=true;};
  },[]);

  if(publicRfiId) return <PublicRFIForm rfiId={publicRfiId}/>;
  if(publicCoId) return <PublicCOForm coId={publicCoId}/>;
  if(publicInspId) return <PublicInspectorForm reportId={publicInspId}/>;
  if(publicTMId) return <PublicTMSignForm ticketId={publicTMId}/>;

  useEffect(()=>{
    const onOnline=()=>{setIsOnline(true);syncQueue();}
    const onOffline=()=>setIsOnline(false);
    window.addEventListener("online",onOnline);window.addEventListener("offline",onOffline);
    setupPushNotifications();
    return()=>{window.removeEventListener("online",onOnline);window.removeEventListener("offline",onOffline);};
  },[]);

  useEffect(()=>{
    if(user) loadProjects();
  },[user]);

  async function loadProjects(){
    try{
      const r=await API.projects.list();
      const ps=Array.isArray(r)?r:[];
      const enriched=await Promise.all(ps.map(async p=>{
        try{
          const [reps,rfis,cos,photos]=await Promise.all([
            API.reports.forProject(p.id).catch(()=>[]),
            API.rfis.forProject(p.id).catch(()=>[]),
            API.changeOrders.forProject(p.id).catch(()=>[]),
            API.photos.forProject(p.id).catch(()=>[]),
          ]);
          const repList=Array.isArray(reps)?reps:[];
          const rfiList=Array.isArray(rfis)?rfis:[];
          const coList=Array.isArray(cos)?cos:[];
          const photoList=Array.isArray(photos)?photos:[];
          const billed=repList.reduce((s,r)=>{
            return s+(r.labor_total||0)+(r.equipment_total||0);
          },0);
          const lastRep=repList.length>0?repList.sort((a,b)=>b.date?.localeCompare(a.date))[0].date:null;
          const openRfis=rfiList.filter(r=>r.status==="Open"||r.status==="Overdue").length;
          const pendingCOs=coList.filter(c=>c.status==="Pending");
          const pendingCOTotal=pendingCOs.reduce((s,c)=>s+(parseFloat(c.amount)||0),0);
          return{...p,_reports:repList.length,_billed:billed,_lastReport:lastRep,
            _openRfis:openRfis,_pendingCOs:pendingCOs.length,_pendingCOTotal:pendingCOTotal,
            _photos:photoList.length};
        }catch{return{...p,_reports:0,_billed:0,_openRfis:0,_pendingCOs:0,_pendingCOTotal:0,_photos:0};}
      }));
      setProjects(enriched);
    }catch(e){setErr(e.message);}
  }

  async function syncQueue(){
    const queue=getQueue();
    if(!queue.length)return;
    setSyncMsg(`Syncing ${queue.length} queued report${queue.length!==1?"s":""}…`);
    let synced=0;
    for(const item of queue){
      try{
        if(item.type==="report"){
          const {rental_equipment,...dbData}=item.data;
          try{await API.reports.create({...dbData,rental_equipment,project_id:item.data.project_id});}
          catch{await API.reports.create({...dbData,project_id:item.data.project_id});}
          removeFromQueue(item.qid);
          const proj=projects.find(p=>p.id===item.data.project_id);
          if(proj) await autoPopulateTimeCards(item.data,proj).catch(()=>{});
          synced++;
        }
      }catch(e){console.warn("Sync failed for item",item.qid,e);}
    }
    setSyncMsg(synced>0?`✓ Synced ${synced} report${synced!==1?"s":""}!`:"");
    if(synced>0){await loadProjects();}
    setTimeout(()=>setSyncMsg(""),4000);
    setPendingCount(getQueue().length);
  }

  function handleLogin(profile){saveSession(profile);setUser(profile);}
  function handleLogout(){clearSession();setUser(null);setScreen("division");setSelectedDiv(null);setSelectedProject(null);}
  function handleDivisionSelect(div){setSelectedDiv(div);setSelectedMfgJob(null);setSelectedMfgPart(null);setScreen("jobs");}
  function handleSelectProject(p){setSelectedProject(p);setScreen("detail");}

  if(!user) return(
    <div style={{maxWidth:480,margin:"0 auto",fontFamily:"'DM Sans',system-ui,sans-serif",color:T.text,background:T.bg,minHeight:"100vh"}}>
      <LoginScreen onLogin={handleLogin}/>
    </div>
  );

  const canEst=canEstimate(user);

  return(
    <div style={{maxWidth:shellMax(screen),margin:"0 auto",transition:"max-width 0.15s ease",fontFamily:"'DM Sans',system-ui,sans-serif",color:T.text,background:T.bg,minHeight:"100vh"}}>
      {syncMsg&&<div style={{background:T.green,color:"#000",padding:"10px 16px",fontSize:13,fontWeight:700,textAlign:"center"}}>{syncMsg}</div>}
      {err&&<div style={{background:T.red,color:"#fff",padding:"8px 16px",fontSize:12,cursor:"pointer"}} onClick={()=>setErr("")}>{err} ✕</div>}
      {!user&&restoring&&(
        <div style={{minHeight:"100vh",display:"flex",alignItems:"center",justifyContent:"center"}}>
          <Spinner/>
        </div>
      )}
      {!user&&!restoring&&<LoginScreen onLogin={handleLogin}/>}
      {user&&screen==="division"&&(
        <DivisionScreen user={user} projects={projects} onSelect={handleDivisionSelect} onLogout={handleLogout}
          onCrew={()=>setScreen("crewDirectory")} onDash={()=>setScreen("pmDashboard")}
          onTimeCards={()=>setScreen("timeCards")} onEstimating={()=>setScreen("estimating")}
          isOnline={isOnline} pendingCount={pendingCount} onSync={syncQueue}/>
      )}
      {user&&screen==="estimating"&&canEstimate(user)&&(
        <EstimatingScreen user={user} onBack={()=>setScreen("division")}/>
      )}
      {user&&screen==="jobs"&&selectedDiv==="Manufacturing"&&!selectedMfgJob&&(
        <ManufacturingJobBoard user={user} onBack={()=>setScreen("division")} onSelectJob={j=>{setSelectedMfgJob(j);setSelectedMfgPart(null);}}/>
      )}
      {user&&screen==="jobs"&&selectedDiv==="Manufacturing"&&selectedMfgJob&&!selectedMfgPart&&(
        <ManufacturingJobDetail job={selectedMfgJob} user={user} onBack={()=>setSelectedMfgJob(null)} onSelectPart={p=>setSelectedMfgPart(p)}/>
      )}
      {user&&screen==="jobs"&&selectedDiv==="Manufacturing"&&selectedMfgJob&&selectedMfgPart&&(
        <ManufacturingTraveler part={selectedMfgPart} job={selectedMfgJob} user={user} onBack={()=>setSelectedMfgPart(null)}/>
      )}
      {user&&screen==="jobs"&&selectedDiv&&selectedDiv!=="Manufacturing"&&(
        <JobBoard user={user} projects={projects} division={selectedDiv} onSelect={handleSelectProject}
          onBack={()=>setScreen("division")} onNew={()=>setScreen("newProject")} onRefresh={loadProjects}/>
      )}
      {user&&screen==="detail"&&selectedProject&&(
        <ProjectDetail project={projects.find(p=>p.id===selectedProject.id)||selectedProject}
          user={user} onBack={()=>setScreen("jobs")} onRefresh={loadProjects}
          onProjectUpdated={async(updated)=>{await loadProjects();setSelectedProject(updated);}}
          onErr={setErr} isOnline={isOnline}/>
      )}
      {user&&screen==="newProject"&&(
        <ProjectForm user={user} onSave={async(data)=>{try{const{_billed,_reports,_lastReport,...dbData}=data;await API.projects.create(dbData);await loadProjects();setScreen("jobs");}catch(e){setErr(e.message);}}} onCancel={()=>setScreen("jobs")}/>
      )}
      {user&&screen==="pmDashboard"&&(
        <PMDashboard user={user} projects={projects} onBack={()=>setScreen("division")} onRefresh={loadProjects} onErr={setErr}/>
      )}
      {user&&screen==="crewDirectory"&&(
        <CrewDirectoryScreen user={user} onBack={()=>setScreen("division")}/>
      )}
      {user&&screen==="timeCards"&&(
        <TimeCardsScreen user={user} projects={projects} onBack={()=>setScreen("division")}/>
      )}
      {user&&screen==="userManagement"&&(
        <UserManagementScreen user={user} onBack={()=>setScreen("division")}/>
      )}
    </div>
  );
}

export default function App(){
  return(
    <ErrorBoundary>
      <AppInner/>
    </ErrorBoundary>
  );
}

/* ── T&M TICKET LIST (inside ProjectDetail tab) ─────────────── */
function TMTicketList({project,user,onOpen,onNew}){
  const [tickets,setTickets]=useState([]);
  const [loading,setLoading]=useState(true);
  const canCreate=user.role==="admin"||user.role==="pm";

  useEffect(()=>{load();},[project.id]);
  async function load(){
    setLoading(true);
    try{const t=await API.tmTickets.forProject(project.id);setTickets(Array.isArray(t)?t:[]);}
    catch(e){console.error(e);}
    setLoading(false);
  }
  async function removeTicket(id){
    if(!window.confirm("Delete this T&M ticket?"))return;
    try{await API.tmTickets.remove(id);await load();}catch(e){}
  }

  const statusColor={draft:T.muted,submitted:T.yellow,approved:T.green};
  const total=tickets.reduce((s,t)=>s+(t.grand_total||0),0);

  return(
    <div>
      {canCreate&&<button onClick={onNew} style={{...primBtn,borderRadius:14,marginBottom:14,background:T.orange,color:"#000"}}>
        + New T&M Ticket
      </button>}
      {tickets.length>0&&<div style={{...cardS,marginBottom:14,display:"flex",justifyContent:"space-between",alignItems:"center"}}>
        <div style={{fontSize:12,color:T.muted}}>{tickets.length} ticket{tickets.length!==1?"s":""}</div>
        <div style={{fontSize:15,fontWeight:900,color:T.green}}>${total.toLocaleString("en-US",{minimumFractionDigits:2})}</div>
      </div>}
      {loading&&<Spinner/>}
      {!loading&&tickets.length===0&&<div style={{textAlign:"center",padding:"40px 16px",color:T.muted}}>
        <div style={{fontSize:44,marginBottom:12}}>🧾</div>
        <div style={{fontSize:14,fontWeight:700,color:T.sub,marginBottom:6}}>No T&M Tickets</div>
        <div style={{fontSize:12}}>Create your first Time & Materials ticket above.</div>
      </div>}
      {tickets.map(t=>(
        <div key={t.id} onClick={()=>onOpen(t)} style={{...cardS,marginBottom:8,cursor:"pointer",borderLeft:`3px solid ${statusColor[t.status]||T.muted}`}}>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start"}}>
            <div>
              <div style={{fontSize:14,fontWeight:800,color:T.orange}}>T&M #{t.ticket_no||"—"}</div>
              <div style={{fontSize:12,color:T.sub}}>{t.ticket_date} {t.submitted_by?"· "+t.submitted_by:""}</div>
              {t.description&&<div style={{fontSize:11,color:T.muted,marginTop:2}}>{t.description.slice(0,60)}{t.description.length>60?"…":""}</div>}
            </div>
            <div style={{textAlign:"right"}}>
              <div style={{fontSize:16,fontWeight:900,color:T.green}}>${(t.grand_total||0).toLocaleString("en-US",{minimumFractionDigits:2})}</div>
              <span style={{background:statusColor[t.status]+"20",color:statusColor[t.status],borderRadius:6,padding:"2px 8px",fontSize:10,fontWeight:700,textTransform:"uppercase"}}>{t.status}</span>
            </div>
          </div>
        </div>
      ))}
    </div>
  );
}

/* ── T&M TICKET FORM ─────────────────────────────────────────── */
function TMTicketForm({project,user,ticket,onBack,onSaved}){
  const isNew=!ticket?.id;
  const canEdit=user.role==="admin"||user.role==="pm"||user.role==="foreman";
  const division=project.division||"Pipeline";
  const positions=getPositions(division);
  const equipList=getEquipList(division);

  // Header
  const [ticketNo,setTicketNo]=useState(ticket?.ticket_no||"");
  const [ticketDate,setTicketDate]=useState(ticket?.ticket_date||today());
  const [description,setDescription]=useState(ticket?.description||"");
  const [poNumber,setPoNumber]=useState(ticket?.po_number||project.po_number||"");
  const [afeNumber,setAfeNumber]=useState(ticket?.afe_number||project.afe_number||"");
  const [workOrder,setWorkOrder]=useState(ticket?.work_order||"");
  const [location,setLocation]=useState(ticket?.location||"");
  const [markupPct,setMarkupPct]=useState(String(ticket?.markup_pct||"0"));
  const [notes,setNotes]=useState(ticket?.notes||"");
  const [status,setStatus]=useState(ticket?.status||"draft");

  // Line items
  const [labor,setLabor]=useState(ticket?.labor||[]);
  const [equipment,setEquipment]=useState(ticket?.equipment||[]);
  const [materials,setMaterials]=useState(ticket?.materials||[]);
  const [other,setOther]=useState(ticket?.other_charges||[]);

  const [saving,setSaving]=useState(false);
  const [tab,setTab]=useState("labor");
  const [showSigPad,setShowSigPad]=useState(false);
  const [signerName,setSignerName]=useState(ticket?.inspector_name||"");
  const [sigData,setSigData]=useState(ticket?.inspector_signature||null);
  const [sigAt,setSigAt]=useState(ticket?.inspector_signed_at||null);
  const [matUploading,setMatUploading]=useState(false);
  const fileInputRef=useRef(null);
  const [showBoxSignModal,setShowBoxSignModal]=useState(false);
  const [bsEmail,setBsEmail]=useState(ticket?.client_email||project.client_email||"");
  const [bsName,setBsName]=useState(ticket?.client_contact||"");
  const [bsSending,setBsSending]=useState(false);
  const [bsError,setBsError]=useState("");
  const [bsSent,setBsSent]=useState(ticket?.hellosign_status==="pending"||ticket?.hellosign_status==="signed");
  // Client sign-off (mirrors Daily Report: Sign Here / Send Link / eSign)
  const [clientSig,setClientSig]=useState(ticket?.client_signature||null);
  const [clientSigAt,setClientSigAt]=useState(ticket?.client_signed_at||null);
  const [clientName,setClientName]=useState(ticket?.client_contact||"");
  const [showClientSigPad,setShowClientSigPad]=useState(false);
  const [showTMShare,setShowTMShare]=useState(false);
  const [tmLinkCopied,setTmLinkCopied]=useState(false);

  async function sendBoxSign(){
    if(!bsEmail.trim()||!bsName.trim()){setBsError("Please enter client name and email.");return;}
    setBsSending(true);setBsError("");
    try{
      // Save ticket first to get an ID
      let ticketId=ticket?.id;
      if(!ticketId){await save();return;} // prompt save first
      const m=(n)=>"$"+Number(n||0).toLocaleString("en-US",{minimumFractionDigits:2,maximumFractionDigits:2});
      const res=await fetch("/.netlify/functions/box-sign-create",{
        method:"POST",
        headers:{"Content-Type":"application/json"},
        body:JSON.stringify({
          docType:"tm",
          reportId:ticketId,
          inspectorEmail:bsEmail.trim(),
          inspectorName:bsName.trim(),
          projectName:project.name,
          customer:project.client||"",
          poNumber:poNumber||"",
          afeNumber:afeNumber||workOrder||"",
          location:location||"",
          description:description||"",
          reportNo:ticketNo,
          reportDate:ticketDate,
          submittedBy:user.name,
          lineItems:{
            labor:labor.map(r=>({
              name:r.name||"",
              classification:r.classification||"",
              hours:r.hours||0,
              rate:m(r.rate),
              amount:m((parseFloat(r.hours)||0)*(parseFloat(r.rate)||0)),
            })),
            equipment:equipment.map(r=>({
              description:r.description||"",
              unit:r.unit||"",
              qty:r.qty||0,
              rate:m(r.rate),
              amount:m((parseFloat(r.qty)||0)*(parseFloat(r.rate)||0)),
            })),
            rental:[],
            materials:materials.map(r=>({
              description:r.description||"",
              qty:r.qty||0,
              unit_price:m(r.unit_price),
              amount:m((parseFloat(r.qty)||0)*(parseFloat(r.unit_price)||0)),
            })),
            other:other.map(r=>({
              description:r.description||"",
              amount:m(r.amount),
            })),
          },
          laborTotal:m(laborTotal),
          equipmentTotal:m(equipTotal),
          materialsTotal:m(matsTotal),
          otherTotal:m(otherTotal),
          subtotal:m(subtotal),
          markupPct:markupPct||"0",
          markupAmount:m(markupAmt),
          grandTotal:m(grandTotal),
        }),
      });
      const data=await res.json();
      if(!res.ok)throw new Error(data.error||"Failed to send");
      await API.tmTickets.update(ticketId,{hellosign_request_id:data.requestId,hellosign_status:"pending",client_email:bsEmail.trim(),client_contact:bsName.trim()});
      setBsSent(true);setShowBoxSignModal(false);
    }catch(e){setBsError("Error: "+e.message);}
    setBsSending(false);
  }

  const uid=()=>Math.random().toString(36).slice(2,10);

  // Auto-number new tickets
  useEffect(()=>{
    if(isNew&&!ticketNo){
      API.tmTickets.forProject(project.id).then(existing=>{
        const count=(Array.isArray(existing)?existing:[]).length+1;
        setTicketNo(`${project.name||project.job_number||"JOB"}-TM-${count}`);
      }).catch(()=>{});
    }
  },[]);

  // Line item helpers
  const addRow=(setter,template)=>setter(rows=>[...rows,{id:uid(),...template}]);
  const updateRow=(setter,id,key,val)=>setter(rows=>rows.map(r=>r.id===id?{...r,[key]:val}:r));
  const removeRow=(setter,id)=>setter(rows=>rows.filter(r=>r.id!==id));

  // Live totals (computed from state every render)
  const laborTotal=labor.reduce((s,r)=>s+((parseFloat(r.hours)||0)*(parseFloat(r.rate)||0)),0);
  const equipTotal=equipment.reduce((s,r)=>s+((parseFloat(r.qty)||0)*(parseFloat(r.rate)||0)),0);
  const matsTotal=materials.reduce((s,r)=>s+((parseFloat(r.qty)||0)*(parseFloat(r.unit_price)||0)),0);
  const otherTotal=other.reduce((s,r)=>s+(parseFloat(r.amount)||0),0);
  const subtotal=laborTotal+equipTotal+matsTotal+otherTotal;
  const markupAmt=subtotal*(parseFloat(markupPct)||0)/100;
  const grandTotal=subtotal+markupAmt;

  async function uploadMaterialAttachment(file,matId){
    setMatUploading(true);
    try{
      const ext=file.name.split(".").pop();
      const path=`tm-materials/${project.id}/${matId}-${Date.now()}.${ext}`;
      await storageUpload("documents",path,file,file.type||undefined);
      const publicUrl=storagePublicUrl("documents",path);
      updateRow(setMaterials,matId,"attachment_url",publicUrl);
      updateRow(setMaterials,matId,"attachment_name",file.name);
    }catch(e){alert("Upload failed: "+e.message);}
    setMatUploading(false);
  }

  async function save(newStatus){
    setSaving(true);
    const data={
      project_id:project.id,
      ticket_no:ticketNo,ticket_date:ticketDate,description,
      po_number:poNumber,afe_number:afeNumber,work_order:workOrder,
      location,notes,markup_pct:parseFloat(markupPct)||0,
      status:newStatus||status,
      submitted_by:ticket?.submitted_by||user.name,
      labor,equipment,materials,other_charges:other,
      client_email:bsEmail||null,client_contact:(clientName||bsName)||null,
      labor_total:laborTotal,equipment_total:equipTotal,
      materials_total:matsTotal,other_total:otherTotal,
      subtotal,markup_amount:markupAmt,grand_total:grandTotal,
      inspector_name:signerName||null,
      inspector_signature:sigData||null,
      inspector_signed_at:sigAt||null,
      client_signature:clientSig||null,
      client_signed_at:clientSigAt||null,
      customer:project.client||null,
      updated_at:new Date().toISOString(),
    };
    try{
      if(isNew)await API.tmTickets.create(data);
      else await API.tmTickets.update(ticket.id,data);
      onSaved&&onSaved();
    }catch(e){alert("Error saving: "+e.message);}
    setSaving(false);
  }

  function printTicket(){
    const fmt=n=>"$"+(n||0).toLocaleString("en-US",{minimumFractionDigits:2,maximumFractionDigits:2});
    const thStyle=`padding:5px 8px;font-size:8pt;background:#1f3864;color:#fff;text-align:`;
    const tdStyle=`padding:5px 8px;font-size:8.5pt;border-bottom:1px solid #e5e7eb;text-align:`;
    const html=`<!DOCTYPE html><html><head><meta charset="utf-8">
<title>T&M #${ticketNo}</title>
<style>@page{size:letter portrait;margin:0.4in;}*{box-sizing:border-box;margin:0;padding:0;font-family:Arial,sans-serif;}@media print{body{-webkit-print-color-adjust:exact;print-color-adjust:exact;}}</style>
</head><body style="color:#111;font-size:9pt;">
<div style="display:flex;justify-content:space-between;align-items:flex-start;border-bottom:3px solid #1f3864;padding-bottom:10px;margin-bottom:10px">
  <div style="display:flex;align-items:center;gap:12px">
    <div style="background:#1f3864;color:#fff;font-size:26pt;font-weight:900;letter-spacing:2px;padding:6px 14px;border-radius:6px;line-height:1;font-family:Arial Black,sans-serif">AIME</div>
    <div>
      <div style="font-size:8pt;font-weight:700;color:#1f3864;letter-spacing:1px;text-transform:uppercase">Atlantic Industrial Mechanical</div>
      <div style="font-size:7.5pt;color:#555">&amp; Environmental Inc.</div>
      <div style="font-size:7pt;color:#888">5730 Pennington Ave, Baltimore, MD 21226</div>
    </div>
  </div>
  <div style="text-align:right">
    <div style="font-size:16pt;font-weight:700;color:#1f3864">Time &amp; Materials Ticket</div>
    <div style="font-size:9pt;font-weight:700;color:#333">T&amp;M # ${ticketNo}</div>
    <div style="font-size:8.5pt;color:#555">Date: ${ticketDate}</div>
    <div style="background:#1f3864;color:#fff;font-size:7pt;font-weight:700;padding:3px 10px;border-radius:4px;display:inline-block;margin-top:4px;text-transform:uppercase;letter-spacing:1px">${status}</div>
  </div>
</div>
<div style="display:grid;grid-template-columns:1fr 1fr;border:1px solid #e5e7eb;margin-bottom:8px;">
  ${[["Project",project.name||"—"],["Customer",project.client||"—"],["PO #",poNumber||"—"],["AFE/WO #",(afeNumber||workOrder||"—")],["Location",location||"—"],["Submitted By",user.name]].map(([l,v])=>`<div style="display:flex;border-bottom:1px solid #e5e7eb"><div style="background:#f3f4f6;padding:3px 8px;font-weight:700;font-size:7.5pt;width:100px;border-right:1px solid #e5e7eb;flex-shrink:0">${l}</div><div style="padding:3px 8px;font-size:8pt">${v}</div></div>`).join("")}
</div>
${description?`<div style="border:1px solid #e5e7eb;padding:5px 8px;margin-bottom:8px;font-size:8.5pt"><b>Description:</b> ${description}</div>`:""}
${labor.length?`<div style="margin-bottom:8px"><div style="background:#1f3864;color:#fff;font-size:7.5pt;font-weight:700;padding:3px 8px;text-transform:uppercase">Labor</div>
<table style="width:100%;border-collapse:collapse;border:1px solid #e5e7eb"><thead><tr>
  <th style="${thStyle}left">Name</th><th style="${thStyle}left">Classification</th>
  <th style="${thStyle}center">Hours</th><th style="${thStyle}right">Rate/Hr</th><th style="${thStyle}right">Amount</th>
</tr></thead><tbody>
${labor.map(r=>`<tr><td style="${tdStyle}left">${r.name||""}</td><td style="${tdStyle}left">${r.classification||""}</td><td style="${tdStyle}center">${r.hours||0}</td><td style="${tdStyle}right">${fmt(r.rate)}</td><td style="${tdStyle}right">${fmt((r.hours||0)*(r.rate||0))}</td></tr>`).join("")}
<tr style="font-weight:700;background:#f9fafb"><td colspan="4" style="${tdStyle}right">Labor Total</td><td style="${tdStyle}right">${fmt(laborTotal)}</td></tr>
</tbody></table></div>`:""}
${equipment.length?`<div style="margin-bottom:8px"><div style="background:#374151;color:#fff;font-size:7.5pt;font-weight:700;padding:3px 8px;text-transform:uppercase">Equipment</div>
<table style="width:100%;border-collapse:collapse;border:1px solid #e5e7eb"><thead><tr>
  <th style="${thStyle}left">Equipment</th><th style="${thStyle}center">Unit</th><th style="${thStyle}center">Qty</th><th style="${thStyle}right">Rate</th><th style="${thStyle}right">Amount</th>
</tr></thead><tbody>
${equipment.map(r=>`<tr><td style="${tdStyle}left">${r.description||""}</td><td style="${tdStyle}center">${r.unit||""}</td><td style="${tdStyle}center">${r.qty||0}</td><td style="${tdStyle}right">${fmt(r.rate)}</td><td style="${tdStyle}right">${fmt((r.qty||0)*(r.rate||0))}</td></tr>`).join("")}
<tr style="font-weight:700;background:#f9fafb"><td colspan="4" style="${tdStyle}right">Equipment Total</td><td style="${tdStyle}right">${fmt(equipTotal)}</td></tr>
</tbody></table></div>`:""}
${materials.length?`<div style="margin-bottom:8px"><div style="background:#4B5563;color:#fff;font-size:7.5pt;font-weight:700;padding:3px 8px;text-transform:uppercase">Materials</div>
<table style="width:100%;border-collapse:collapse;border:1px solid #e5e7eb"><thead><tr>
  <th style="${thStyle}left">Description</th><th style="${thStyle}center">Qty</th><th style="${thStyle}center">Unit</th><th style="${thStyle}right">Unit Price</th><th style="${thStyle}right">Amount</th>
</tr></thead><tbody>
${materials.map(r=>`<tr><td style="${tdStyle}left">${r.description||""}</td><td style="${tdStyle}center">${r.qty||0}</td><td style="${tdStyle}center">${r.unit||""}</td><td style="${tdStyle}right">${fmt(r.unit_price)}</td><td style="${tdStyle}right">${fmt((r.qty||0)*(r.unit_price||0))}</td></tr>`).join("")}
<tr style="font-weight:700;background:#f9fafb"><td colspan="4" style="${tdStyle}right">Materials Total</td><td style="${tdStyle}right">${fmt(matsTotal)}</td></tr>
</tbody></table></div>`:""}
${other.length?`<div style="margin-bottom:8px"><div style="background:#6B7280;color:#fff;font-size:7.5pt;font-weight:700;padding:3px 8px;text-transform:uppercase">Other</div>
<table style="width:100%;border-collapse:collapse;border:1px solid #e5e7eb"><thead><tr>
  <th style="${thStyle}left">Description</th><th style="${thStyle}right">Amount</th>
</tr></thead><tbody>
${other.map(r=>`<tr><td style="${tdStyle}left">${r.description||""}</td><td style="${tdStyle}right">${fmt(r.amount)}</td></tr>`).join("")}
</tbody></table></div>`:""}
<table style="width:260px;margin-left:auto;border-collapse:collapse;border:1px solid #e5e7eb;margin-bottom:10px">
  ${[["Labor",fmt(laborTotal)],["Equipment",fmt(equipTotal)],["Materials",fmt(matsTotal)],["Other",fmt(otherTotal)],["Subtotal",fmt(subtotal)],...(parseFloat(markupPct)>0?[[`Markup (${markupPct}%)`,fmt(markupAmt)]]:[])].map(([l,v])=>`<tr><td style="padding:4px 8px;font-size:8.5pt;border-bottom:1px solid #e5e7eb">${l}</td><td style="padding:4px 8px;font-size:8.5pt;text-align:right;border-bottom:1px solid #e5e7eb">${v}</td></tr>`).join("")}
  <tr style="background:#1f3864;color:#fff"><td style="padding:6px 8px;font-size:10pt;font-weight:900">GRAND TOTAL</td><td style="padding:6px 8px;font-size:10pt;font-weight:900;text-align:right">${fmt(grandTotal)}</td></tr>
</table>
${notes?`<div style="border:1px solid #e5e7eb;padding:5px 8px;margin-bottom:10px;font-size:8.5pt"><b>Notes:</b> ${notes}</div>`:""}
<div style="display:grid;grid-template-columns:1fr 1fr;gap:16px;margin-top:14px;border-top:2px solid #e5e7eb;padding-top:10px">
  <div><div style="font-size:7.5pt;font-weight:700;color:#555;margin-bottom:4px">PREPARED BY / PM</div>
    ${sigData?`<img src="${sigData}" style="max-height:50px;max-width:180px;display:block;background:#fff;border:1px solid #ccc;border-radius:4px;padding:4px"/>`:
    `<div style="border-bottom:1px solid #000;height:36px;width:200px;margin-bottom:4px"></div>`}
    <div style="font-size:8.5pt;font-weight:700;margin-top:4px">${signerName||user.name}</div>
    <div style="font-size:7.5pt;color:#555">Date: ${ticketDate}</div>
  </div>
  <div><div style="font-size:7.5pt;font-weight:700;color:#555;margin-bottom:4px">CUSTOMER APPROVAL</div>
    ${clientSig?`<img src="${clientSig}" style="max-height:50px;max-width:180px;display:block;background:#fff;border:1px solid #ccc;border-radius:4px;padding:4px"/>
    <div style="font-size:8.5pt;font-weight:700;margin-top:4px">${clientName||bsName||""}</div>
    <div style="font-size:7.5pt;color:#555">Date: ${clientSigAt?new Date(clientSigAt).toLocaleDateString():ticketDate}</div>`:
    `<div style="border-bottom:1px solid #000;height:36px;width:200px;margin-bottom:4px"></div>
    <div style="font-size:7.5pt;color:#555">Name / Signature</div>
    <div style="border-bottom:1px solid #000;width:200px;margin:8px 0 4px"></div>
    <div style="font-size:7.5pt;color:#555">Date / Title</div>`}
  </div>
</div>
<div style="text-align:center;margin-top:8px;font-size:6.5pt;color:#999;border-top:1px solid #eee;padding-top:5px">AIME Field Pro · ${project.name} · T&M #${ticketNo} · Generated ${new Date().toLocaleString()}</div>
</body></html>`;
    const win=window.open("","_blank","width=950,height=800");
    if(!win){alert("Allow popups to print.");return;}
    win.document.write(html);win.document.close();
    setTimeout(()=>{win.focus();win.print();},1200);
  }

  const ri={...inp,fontSize:13,padding:"7px 10px"};
  const fmt=n=>"$"+(n||0).toLocaleString("en-US",{minimumFractionDigits:2,maximumFractionDigits:2});

  if(showTMShare){
    const link=`${window.location.origin}${window.location.pathname}?tmsign=${ticket?.id}`;
    return(
      <div style={{background:T.bg,minHeight:"100vh",fontFamily:"inherit"}}>
        <div style={{background:T.surface,borderBottom:`1px solid ${T.border}`,padding:"12px 16px"}}>
          <button onClick={()=>setShowTMShare(false)} style={{background:"none",border:"none",color:T.sub,fontSize:13,cursor:"pointer",fontFamily:"inherit",display:"block",marginBottom:4}}>← Back</button>
          <div style={{fontSize:15,fontWeight:900,color:T.text}}>📤 Send Client Link</div>
        </div>
        <div style={{padding:"20px 16px"}}>
          <div style={{...cardS,marginBottom:14,background:T.blueLow,border:`1px solid ${T.blue}40`}}>
            <div style={{fontSize:14,fontWeight:800,color:T.blue,marginBottom:4}}>How it works</div>
            <div style={{fontSize:12,color:T.sub,lineHeight:1.7}}>
              1. Copy the link below and text or email it to the client<br/>
              2. They open it on their phone — no login needed<br/>
              3. They review the itemized charges and draw their signature<br/>
              4. Saves directly to this ticket automatically
            </div>
          </div>
          <div style={{...cardS,marginBottom:12,padding:"12px 14px"}}>
            <div style={{fontSize:11,fontWeight:700,color:T.muted,marginBottom:6}}>SIGNING LINK</div>
            <div style={{fontSize:12,color:T.sub,wordBreak:"break-all",lineHeight:1.6,background:T.surface,borderRadius:8,padding:"10px 12px"}}>{link}</div>
          </div>
          <button onClick={()=>{navigator.clipboard.writeText(link).then(()=>{setTmLinkCopied(true);setTimeout(()=>setTmLinkCopied(false),3000);}).catch(()=>{const el=document.createElement("textarea");el.value=link;document.body.appendChild(el);el.select();document.execCommand("copy");document.body.removeChild(el);setTmLinkCopied(true);setTimeout(()=>setTmLinkCopied(false),3000);});}}
            style={{...primBtn,borderRadius:14,marginBottom:10,background:tmLinkCopied?T.green:T.orange,transition:"background 0.3s"}}>
            {tmLinkCopied?"✅ Copied! Paste into a text or email":"📋 Copy Link to Clipboard"}
          </button>
          {tmLinkCopied&&<div style={{background:T.greenLow,border:`1px solid ${T.green}40`,borderRadius:10,padding:"10px 14px",fontSize:13,color:T.green,textAlign:"center",marginBottom:10}}>✓ Link copied — paste it into a text or email to the client</div>}
          <div style={{display:"flex",alignItems:"center",gap:10,marginBottom:10}}>
            <div style={{flex:1,height:1,background:T.border}}/><span style={{fontSize:11,color:T.muted}}>OR</span><div style={{flex:1,height:1,background:T.border}}/>
          </div>
          <button onClick={()=>{
            const subj=`T&M Ticket ${ticketNo} — ${project.name} — Approval Requested`;
            const body=`Please review and sign the Time & Materials ticket for ${project.name} dated ${ticketDate}.%0D%0A%0D%0AOpen this link — no login required:%0D%0A%0D%0A${link}%0D%0A%0D%0AThank you`;
            window.location.href=`mailto:${bsEmail||""}?subject=${encodeURIComponent(subj)}&body=${encodeURIComponent(body).replace(/%250D%250A/g,'%0D%0A')}`;
          }} style={{...ghostBtn,width:"100%",textAlign:"center",fontSize:14}}>
            📧 Open Email Draft
          </button>
        </div>
      </div>
    );
  }

  return(
    <div style={{background:T.bg,minHeight:"100vh",fontFamily:"inherit"}}>
      {/* Sticky header with live grand total */}
      <div style={{background:T.surface,borderBottom:`1px solid ${T.border}`,padding:"12px 16px",position:"sticky",top:0,zIndex:50}}>
        <button onClick={onBack} style={{background:"none",border:"none",color:T.sub,fontSize:13,cursor:"pointer",fontFamily:"inherit",display:"block",marginBottom:4}}>← Back</button>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}>
          <div>
            <div style={{fontSize:13,fontWeight:900,color:T.orange}}>🧾 {isNew?"New T&M":"T&M #"+ticketNo}</div>
            <div style={{fontSize:10,color:T.muted}}>{project.name}</div>
          </div>
          <div style={{textAlign:"right"}}>
            <div style={{fontSize:20,fontWeight:900,color:T.green}}>{fmt(grandTotal)}</div>
            <div style={{fontSize:9,color:T.muted}}>Grand Total</div>
          </div>
        </div>
      </div>

      <div style={{padding:"12px 16px 100px"}}>
        {/* Header info card */}
        <div style={{...cardS,marginBottom:12}}>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8,marginBottom:8}}>
            <div><label style={lbl}>Ticket # (auto)</label><input value={ticketNo} onChange={e=>setTicketNo(e.target.value)} style={ri}/></div>
            <div><label style={lbl}>Date</label><input type="date" value={ticketDate} onChange={e=>setTicketDate(e.target.value)} style={ri}/></div>
          </div>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8,marginBottom:8}}>
            <div><label style={lbl}>PO #</label><input value={poNumber} onChange={e=>setPoNumber(e.target.value)} style={ri}/></div>
            <div><label style={lbl}>AFE / WO #</label><input value={afeNumber} onChange={e=>setAfeNumber(e.target.value)} style={ri}/></div>
          </div>
          <div style={{marginBottom:8}}><label style={lbl}>Location</label><input value={location} onChange={e=>setLocation(e.target.value)} placeholder="Site location" style={ri}/></div>
          <div><label style={lbl}>Description of Work</label><textarea value={description} onChange={e=>setDescription(e.target.value)} rows={2} style={{...ri,resize:"vertical",width:"100%"}}/></div>
        </div>

        {/* Live mini totals bar */}
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr 1fr",gap:6,marginBottom:12}}>
          {[["👷",laborTotal,"Labor"],["🚜",equipTotal,"Equip"],["🔩",matsTotal,"Matls"],["➕",otherTotal,"Other"]].map(([icon,val,lbl2])=>(
            <div key={lbl2} style={{...cardS,textAlign:"center",padding:"8px 4px"}}>
              <div style={{fontSize:10}}>{icon}</div>
              <div style={{fontSize:13,fontWeight:900,color:val>0?T.green:T.muted}}>{fmt(val)}</div>
              <div style={{fontSize:8,color:T.muted,textTransform:"uppercase"}}>{lbl2}</div>
            </div>
          ))}
        </div>

        {/* Tab bar */}
        <div style={{display:"flex",background:T.surface,borderRadius:12,padding:4,marginBottom:12,gap:3,overflowX:"auto"}}>
          {[["labor","👷 Labor"],["equipment","🚜 Equip"],["materials","🔩 Materials"],["other","➕ Other"],["summary","📊 Summary"]].map(([id,label])=>(
            <button key={id} onClick={()=>setTab(id)} style={{flexShrink:0,padding:"8px 10px",background:tab===id?T.orange:"transparent",color:tab===id?"#000":T.muted,border:"none",borderRadius:9,fontSize:11,fontWeight:700,cursor:"pointer",fontFamily:"inherit"}}>
              {label}
            </button>
          ))}
        </div>

        {/* ── LABOR TAB ── */}
        {tab==="labor"&&<div>
          <button onClick={()=>addRow(setLabor,{name:"",classification:"",hours:"",rate:""})}
            style={{...primBtn,borderRadius:12,marginBottom:10,background:T.blue,fontSize:13}}>+ Add Worker</button>
          {labor.map(r=>(
            <div key={r.id} style={{...cardS,marginBottom:10,borderLeft:`3px solid ${T.blue}`}}>
              <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:8}}>
                <div style={{fontSize:12,fontWeight:700,color:T.blue}}>{r.name||"New Worker"}</div>
                <div style={{display:"flex",alignItems:"center",gap:8}}>
                  <span style={{fontSize:15,fontWeight:900,color:T.green}}>{fmt((r.hours||0)*(r.rate||0))}</span>
                  <button onClick={()=>removeRow(setLabor,r.id)} style={{background:"none",border:`1px solid ${T.red}30`,borderRadius:6,padding:"3px 8px",color:T.red,fontSize:11,cursor:"pointer",fontFamily:"inherit"}}>🗑</button>
                </div>
              </div>
              <div style={{marginBottom:8}}>
                <label style={lbl}>Name</label>
                <select value={r.name} onChange={e=>updateRow(setLabor,r.id,"name",e.target.value)} style={{...ri,width:"100%"}}>
                  <option value="">— Select Worker —</option>
                  {NAMES.map(n=><option key={n} value={n}>{n}</option>)}
                  <option value="__other">Other (type below)</option>
                </select>
                {r.name==="__other"&&<input value={r.customName||""} onChange={e=>updateRow(setLabor,r.id,"customName",e.target.value)} placeholder="Enter name" style={{...ri,marginTop:6}}/>}
              </div>
              <div style={{marginBottom:8}}>
                <label style={lbl}>Classification</label>
                <select value={r.classification} onChange={e=>updateRow(setLabor,r.id,"classification",e.target.value)} style={{...ri,width:"100%"}}>
                  <option value="">— Select —</option>
                  {positions.map(p=><option key={p.name} value={p.name}>{p.name}</option>)}
                </select>
              </div>
              <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}>
                <div><label style={lbl}>Hours</label>
                  <input type="number" step="0.5" value={r.hours} onChange={e=>updateRow(setLabor,r.id,"hours",e.target.value)} placeholder="0" style={ri}/>
                </div>
                <div><label style={lbl}>Rate / Hr ($) — manual</label>
                  <input type="number" step="0.01" value={r.rate} onChange={e=>updateRow(setLabor,r.id,"rate",e.target.value)} placeholder="0.00" style={ri}/>
                </div>
              </div>
            </div>
          ))}
          {labor.length===0&&<div style={{textAlign:"center",padding:"24px",color:T.muted,fontSize:12}}>No workers added — tap + Add Worker</div>}
          <div style={{textAlign:"right",fontWeight:900,color:T.green,fontSize:15,marginTop:6}}>Labor Total: {fmt(laborTotal)}</div>
        </div>}

        {/* ── EQUIPMENT TAB ── */}
        {tab==="equipment"&&<div>
          <button onClick={()=>addRow(setEquipment,{description:"",unit:"Hours",qty:"",rate:""})}
            style={{...primBtn,borderRadius:12,marginBottom:10,background:T.blue,fontSize:13}}>+ Add Equipment</button>
          {equipment.map(r=>(
            <div key={r.id} style={{...cardS,marginBottom:10,borderLeft:`3px solid ${T.orange}`}}>
              <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:8}}>
                <div style={{fontSize:12,fontWeight:700,color:T.orange}}>{r.description||"New Equipment"}</div>
                <div style={{display:"flex",alignItems:"center",gap:8}}>
                  <span style={{fontSize:15,fontWeight:900,color:T.green}}>{fmt((r.qty||0)*(r.rate||0))}</span>
                  <button onClick={()=>removeRow(setEquipment,r.id)} style={{background:"none",border:`1px solid ${T.red}30`,borderRadius:6,padding:"3px 8px",color:T.red,fontSize:11,cursor:"pointer",fontFamily:"inherit"}}>🗑</button>
                </div>
              </div>
              <div style={{marginBottom:8}}>
                <label style={lbl}>Equipment</label>
                <select value={r.description} onChange={e=>{
                  const item=equipList.find(i=>i.name===e.target.value);
                  updateRow(setEquipment,r.id,"description",e.target.value);
                  if(item){updateRow(setEquipment,r.id,"unit",item.unit||"Hours");}
                }} style={{...ri,width:"100%"}}>
                  <option value="">— Select Equipment —</option>
                  {equipList.filter(e=>e.name).map(e=><option key={e.name} value={e.name}>{e.name}</option>)}
                  <option value="__other">Other (type below)</option>
                </select>
                {r.description==="__other"&&<input value={r.customDesc||""} onChange={e=>updateRow(setEquipment,r.id,"customDesc",e.target.value)} placeholder="Describe equipment" style={{...ri,marginTop:6}}/>}
              </div>
              <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:8}}>
                <div><label style={lbl}>Unit</label>
                  <select value={r.unit} onChange={e=>updateRow(setEquipment,r.id,"unit",e.target.value)} style={{...ri,width:"100%"}}>
                    {["Hours","Days","Weeks","Miles","EA"].map(u=><option key={u}>{u}</option>)}
                  </select>
                </div>
                <div><label style={lbl}>Qty</label>
                  <input type="number" step="0.5" value={r.qty} onChange={e=>updateRow(setEquipment,r.id,"qty",e.target.value)} placeholder="1" style={ri}/>
                </div>
                <div><label style={lbl}>Rate ($) — manual</label>
                  <input type="number" step="0.01" value={r.rate} onChange={e=>updateRow(setEquipment,r.id,"rate",e.target.value)} placeholder="0.00" style={ri}/>
                </div>
              </div>
            </div>
          ))}
          {equipment.length===0&&<div style={{textAlign:"center",padding:"24px",color:T.muted,fontSize:12}}>No equipment — tap + Add Equipment</div>}
          <div style={{textAlign:"right",fontWeight:900,color:T.green,fontSize:15,marginTop:6}}>Equipment Total: {fmt(equipTotal)}</div>
        </div>}

        {/* ── MATERIALS TAB ── */}
        {tab==="materials"&&<div>
          <button onClick={()=>addRow(setMaterials,{description:"",qty:"",unit:"EA",unit_price:"",attachment_url:null,attachment_name:null})}
            style={{...primBtn,borderRadius:12,marginBottom:10,background:T.blue,fontSize:13}}>+ Add Material</button>
          {materials.map(r=>(
            <div key={r.id} style={{...cardS,marginBottom:10,borderLeft:`3px solid ${T.green}`}}>
              <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:8}}>
                <div style={{fontSize:12,fontWeight:700,color:T.green}}>{r.description||"New Material"}</div>
                <div style={{display:"flex",alignItems:"center",gap:8}}>
                  <span style={{fontSize:15,fontWeight:900,color:T.green}}>{fmt((r.qty||0)*(r.unit_price||0))}</span>
                  <button onClick={()=>removeRow(setMaterials,r.id)} style={{background:"none",border:`1px solid ${T.red}30`,borderRadius:6,padding:"3px 8px",color:T.red,fontSize:11,cursor:"pointer",fontFamily:"inherit"}}>🗑</button>
                </div>
              </div>
              <div style={{marginBottom:8}}><label style={lbl}>Description</label>
                <input value={r.description} onChange={e=>updateRow(setMaterials,r.id,"description",e.target.value)} placeholder="Material description" style={{...ri,width:"100%"}}/>
              </div>
              <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:8,marginBottom:8}}>
                <div><label style={lbl}>Qty</label><input type="number" step="0.01" value={r.qty} onChange={e=>updateRow(setMaterials,r.id,"qty",e.target.value)} placeholder="1" style={ri}/></div>
                <div><label style={lbl}>Unit</label><input value={r.unit} onChange={e=>updateRow(setMaterials,r.id,"unit",e.target.value)} placeholder="EA" style={ri}/></div>
                <div><label style={lbl}>Unit Price ($)</label><input type="number" step="0.01" value={r.unit_price} onChange={e=>updateRow(setMaterials,r.id,"unit_price",e.target.value)} placeholder="0.00" style={ri}/></div>
              </div>
              {/* Attachment */}
              <div style={{borderTop:`1px solid ${T.border}`,paddingTop:8,marginTop:4}}>
                <label style={lbl}>📎 Attachment (receipt, photo, invoice)</label>
                {r.attachment_url?(
                  <div style={{display:"flex",alignItems:"center",gap:8,marginTop:4}}>
                    <a href={r.attachment_url} target="_blank" rel="noreferrer"
                      style={{fontSize:12,color:T.blue,textDecoration:"none",flex:1,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>
                      📎 {r.attachment_name||"Attachment"}
                    </a>
                    <button onClick={()=>{updateRow(setMaterials,r.id,"attachment_url",null);updateRow(setMaterials,r.id,"attachment_name",null);}}
                      style={{background:"none",border:`1px solid ${T.red}30`,borderRadius:6,padding:"2px 8px",color:T.red,fontSize:11,cursor:"pointer",fontFamily:"inherit",flexShrink:0}}>Remove</button>
                  </div>
                ):(
                  <div style={{marginTop:4}}>
                    <input type="file" style={{display:"none"}} ref={el=>{if(el)el._matId=r.id;}}
                      id={`file-${r.id}`}
                      onChange={async e=>{if(e.target.files[0])await uploadMaterialAttachment(e.target.files[0],r.id);}}/>
                    <label htmlFor={`file-${r.id}`}
                      style={{display:"inline-block",background:T.blueLow,border:`1px dashed ${T.blue}`,borderRadius:8,padding:"6px 14px",fontSize:12,color:T.blue,cursor:"pointer",fontWeight:600}}>
                      {matUploading?"Uploading…":"📎 Upload Receipt / Photo"}
                    </label>
                  </div>
                )}
              </div>
            </div>
          ))}
          {materials.length===0&&<div style={{textAlign:"center",padding:"24px",color:T.muted,fontSize:12}}>No materials — tap + Add Material</div>}
          <div style={{textAlign:"right",fontWeight:900,color:T.green,fontSize:15,marginTop:6}}>Materials Total: {fmt(matsTotal)}</div>
        </div>}

        {/* ── OTHER TAB ── */}
        {tab==="other"&&<div>
          <button onClick={()=>addRow(setOther,{description:"",amount:""})}
            style={{...primBtn,borderRadius:12,marginBottom:10,background:T.blue,fontSize:13}}>+ Add Charge</button>
          {other.map(r=>(
            <div key={r.id} style={{...cardS,marginBottom:8,display:"flex",gap:10,alignItems:"flex-end"}}>
              <div style={{flex:2}}><label style={lbl}>Description</label><input value={r.description} onChange={e=>updateRow(setOther,r.id,"description",e.target.value)} placeholder="Misc charge" style={ri}/></div>
              <div style={{flex:1}}><label style={lbl}>Amount ($)</label><input type="number" step="0.01" value={r.amount} onChange={e=>updateRow(setOther,r.id,"amount",e.target.value)} placeholder="0.00" style={ri}/></div>
              <button onClick={()=>removeRow(setOther,r.id)} style={{background:"none",border:`1px solid ${T.red}30`,borderRadius:6,padding:"8px 10px",color:T.red,fontSize:12,cursor:"pointer",fontFamily:"inherit",marginBottom:2}}>🗑</button>
            </div>
          ))}
          {other.length===0&&<div style={{textAlign:"center",padding:"24px",color:T.muted,fontSize:12}}>No other charges</div>}
          <div style={{textAlign:"right",fontWeight:900,color:T.green,fontSize:15,marginTop:6}}>Other Total: {fmt(otherTotal)}</div>
        </div>}

        {/* ── SUMMARY TAB ── */}
        {tab==="summary"&&<div>
          {/* Full totals breakdown - always live */}
          <div style={{...cardS,marginBottom:12,border:`1px solid ${T.green}30`}}>
            <div style={{fontSize:13,fontWeight:800,color:T.text,marginBottom:12}}>💰 Cost Summary</div>
            {[["👷 Labor",laborTotal,labor.length+" worker"+(labor.length!==1?"s":"")],
              ["🚜 Equipment",equipTotal,equipment.length+" item"+(equipment.length!==1?"s":"")],
              ["🔩 Materials",matsTotal,materials.length+" item"+(materials.length!==1?"s":"")],
              ["➕ Other",otherTotal,other.length+" charge"+(other.length!==1?"s":"")]
            ].map(([l,v,sub])=>(
              <div key={l} style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"8px 0",borderBottom:`1px solid ${T.border}`}}>
                <div><div style={{fontSize:13,color:T.sub}}>{l}</div><div style={{fontSize:10,color:T.muted}}>{sub}</div></div>
                <span style={{fontSize:14,fontWeight:700,color:v>0?T.text:T.muted}}>{fmt(v)}</span>
              </div>
            ))}
            <div style={{display:"flex",justifyContent:"space-between",padding:"8px 0",borderBottom:`1px solid ${T.border}`}}>
              <span style={{fontSize:13,color:T.sub,fontWeight:700}}>Subtotal</span>
              <span style={{fontSize:15,fontWeight:800}}>{fmt(subtotal)}</span>
            </div>
            <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",padding:"8px 0",borderBottom:`1px solid ${T.border}`}}>
              <div style={{display:"flex",alignItems:"center",gap:10}}>
                <span style={{fontSize:13,color:T.sub}}>Markup %</span>
                <input type="number" step="0.5" value={markupPct} onChange={e=>setMarkupPct(e.target.value)}
                  style={{...inp,width:65,textAlign:"center",fontSize:14}} placeholder="0"/>
              </div>
              <span style={{fontSize:13,fontWeight:700,color:T.orange}}>{fmt(markupAmt)}</span>
            </div>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"12px 0",borderTop:`2px solid ${T.border}`,marginTop:4}}>
              <span style={{fontSize:17,fontWeight:900}}>GRAND TOTAL</span>
              <span style={{fontSize:26,fontWeight:900,color:T.green}}>{fmt(grandTotal)}</span>
            </div>
          </div>

          <div style={{...cardS,marginBottom:12}}>
            <label style={lbl}>Notes / Special Conditions</label>
            <textarea value={notes} onChange={e=>setNotes(e.target.value)} rows={3} placeholder="Any notes…" style={{...inp,resize:"vertical"}}/>
          </div>

          {canEdit&&<div style={{...cardS,marginBottom:12}}>
            <label style={lbl}>Status</label>
            <select value={status} onChange={e=>setStatus(e.target.value)} style={inpSel}>
              <option value="draft">Draft</option>
              <option value="submitted">Submitted</option>
              <option value="approved">Approved</option>
            </select>
          </div>}

          {/* Client sign-off — same three options as the Daily Report */}
          {clientSig?(
            <div style={{...cardS,marginBottom:12,borderLeft:`3px solid ${T.green}`}}>
              <div style={{fontSize:12,fontWeight:700,color:T.green,textTransform:"uppercase",letterSpacing:"1px"}}>✅ Client Sign-Off</div>
              <div style={{fontSize:14,fontWeight:700,color:T.orange,marginTop:2}}>{clientName||bsName||"Client"}</div>
              {clientSigAt&&<div style={{fontSize:11,color:T.muted,marginTop:2}}>{new Date(clientSigAt).toLocaleString()}</div>}
              <div style={{background:"#fff",borderRadius:10,padding:4,marginTop:8}}>
                <img src={clientSig} alt="Client signature" style={{width:"100%",borderRadius:8,display:"block"}}/>
              </div>
              {canEdit&&<button onClick={()=>{if(window.confirm("Clear the client signature?")){setClientSig(null);setClientSigAt(null);}}}
                style={{...ghostBtn,fontSize:11,color:T.red,border:`1px solid ${T.red}30`,marginTop:8}}>Clear</button>}
            </div>
          ):(
            <div style={{marginBottom:12}}>
              <div style={{fontSize:11,fontWeight:700,color:T.muted,textTransform:"uppercase",letterSpacing:"1px",marginBottom:8}}>🔏 Client Sign-Off</div>
              {bsSent&&<div style={{background:T.blueLow,border:`1px solid ${T.blue}40`,borderRadius:10,padding:"10px 14px",marginBottom:8,fontSize:12,color:T.blue,fontWeight:600}}>
                📦 Box Sign request sent — waiting for client to sign
              </div>}
              <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:8}}>
                <button onClick={()=>setShowClientSigPad(true)}
                  style={{...primBtn,background:T.greenLow,color:T.green,border:`1px solid ${T.green}40`,borderRadius:12,fontSize:12}}>
                  ✍️ Sign Here
                </button>
                <button onClick={()=>{if(isNew){alert("Save the ticket first — the link needs a ticket ID.");return;}setShowTMShare(true);}}
                  style={{...primBtn,background:T.blueLow,color:T.blue,border:`1px solid ${T.blue}40`,borderRadius:12,fontSize:12}}>
                  🔗 Send Link
                </button>
                <button onClick={()=>setShowBoxSignModal(true)}
                  style={{...primBtn,background:"#1e3a5f",color:"#60A5FA",border:"1px solid #2563EB40",borderRadius:12,fontSize:12}}>
                  ✉️ eSign
                </button>
              </div>
              <div style={{fontSize:10,color:T.muted,textAlign:"center",marginTop:5}}>
                Sign Here = on this device · Send Link = client's phone · eSign = Box Sign email
              </div>
            </div>
          )}

          <div style={{...cardS,marginBottom:12}}>
            <div style={{fontSize:12,fontWeight:800,color:T.text,marginBottom:8}}>✍️ PM Sign-Off</div>
            {sigData?(
              <div>
                <div style={{background:"#fff",border:"1px solid #ccc",borderRadius:8,padding:8,display:"inline-block",marginBottom:6}}>
                  <img src={sigData} style={{maxHeight:60,maxWidth:200,display:"block"}}/>
                </div>
                <div style={{fontSize:12,color:T.sub,marginBottom:6}}>{signerName} · {sigAt?new Date(sigAt).toLocaleString():""}</div>
                <button onClick={()=>{setSigData(null);setSigAt(null);}} style={{...ghostBtn,fontSize:11,color:T.red,border:`1px solid ${T.red}30`}}>Clear</button>
              </div>
            ):(
              <button onClick={()=>setShowSigPad(true)} style={{...primBtn,background:T.greenLow,color:T.green,border:`1px solid ${T.green}40`,borderRadius:12}}>✍️ Sign Here</button>
            )}
          </div>
        </div>}

        {/* Action buttons */}
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginTop:16}}>
          <button onClick={()=>save()} disabled={saving}
            style={{...primBtn,borderRadius:14,background:T.orange,color:"#000",opacity:saving?0.5:1,fontSize:15}}>
            {saving?"Saving…":"💾 Save"}
          </button>
          <button onClick={printTicket} style={{...primBtn,borderRadius:14,background:"#1f3864",fontSize:15}}>
            🖨️ Print
          </button>
        </div>
        {!isNew&&canEdit&&<button onClick={async()=>{if(window.confirm("Delete this ticket?"))try{await API.tmTickets.remove(ticket.id);onBack();}catch(e){}}}
          style={{...primBtn,borderRadius:14,background:T.redLow,color:T.red,border:`1px solid ${T.red}30`,marginTop:8,width:"100%",fontSize:13}}>
          🗑 Delete Ticket
        </button>}
      </div>

      {showSigPad&&<SignaturePad reportName={`T&M #${ticketNo} · ${project.name}`}
        onSave={async(name,sig)=>{setSignerName(name);setSigData(sig);setSigAt(new Date().toISOString());setShowSigPad(false);}}
        onCancel={()=>setShowSigPad(false)}/>}

      {showClientSigPad&&<SignaturePad reportName={`Client Sign-Off · T&M #${ticketNo} · ${project.name}`}
        onSave={async(name,sig)=>{setClientName(name);setClientSig(sig);setClientSigAt(new Date().toISOString());setShowClientSigPad(false);}}
        onCancel={()=>setShowClientSigPad(false)}/>}

      {showBoxSignModal&&<div style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.85)",zIndex:200,display:"flex",alignItems:"center",justifyContent:"center",padding:16,fontFamily:"inherit"}}>
        <div style={{background:T.card,borderRadius:16,padding:24,width:"100%",maxWidth:400,border:"1px solid #1f386440"}}>
          <div style={{fontSize:16,fontWeight:900,color:"#60A5FA",marginBottom:4}}>📦 Send T&M for Client Signature</div>
          <div style={{fontSize:12,color:T.muted,marginBottom:16}}>Client receives a Box Sign email with the T&M ticket. You get notified when they sign.</div>
          {bsError&&<div style={{background:T.redLow,border:`1px solid ${T.red}40`,borderRadius:8,padding:"8px 12px",marginBottom:12,fontSize:12,color:T.red}}>{bsError}</div>}
          {!ticket?.id&&<div style={{background:T.yellowLow||T.blueLow,border:`1px solid ${T.yellow}40`,borderRadius:8,padding:"8px 12px",marginBottom:12,fontSize:12,color:T.yellow}}>💾 Save the ticket first before sending for signature.</div>}
          <div style={{marginBottom:10}}><label style={lbl}>Client Contact Name *</label><input value={bsName} onChange={e=>setBsName(e.target.value)} placeholder="John Smith" style={inp} autoFocus/></div>
          <div style={{marginBottom:16}}><label style={lbl}>Client Email *</label><input type="email" value={bsEmail} onChange={e=>setBsEmail(e.target.value)} placeholder="client@company.com" style={inp}/></div>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8}}>
            <button onClick={sendBoxSign} disabled={bsSending||!bsEmail.trim()||!bsName.trim()||!ticket?.id}
              style={{...primBtn,borderRadius:12,background:"#1f3864",opacity:bsSending||!bsEmail.trim()||!bsName.trim()||!ticket?.id?0.5:1}}>
              {bsSending?"Sending…":"📤 Send"}
            </button>
            <button onClick={()=>{setShowBoxSignModal(false);setBsError("");}}
              style={{...ghostBtn,textAlign:"center",borderRadius:12}}>Cancel</button>
          </div>
        </div>
      </div>}
    </div>
  );
}
