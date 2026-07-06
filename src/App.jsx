import React, { useState, useEffect, useRef } from "react";
import * as XLSX from "xlsx";
/* ── ERROR BOUNDARY ─────────────────────────────────────────── */
class ErrorBoundary extends React.Component {
  constructor(props) {
    super(props);
    this.state = { error: null };
  }
  static getDerivedStateFromError(error) {
    return { error };
  }
  componentDidCatch(error, info) {
    console.error("App Error:", error, info);
  }
  render() {
    if (this.state.error) {
      return (
        <div
          style={{
            background: "#EDEBE6",
            minHeight: "100vh",
            display: "flex",
            alignItems: "center",
            justifyContent: "center",
            fontFamily: "'Archivo','Helvetica Neue',system-ui,sans-serif",
            padding: 20,
          }}
        >
          <div
            style={{
              background: "#FFFFFF",
              border: "1px solid #EF4444",
              borderRadius: 16,
              padding: 24,
              maxWidth: 480,
              width: "100%",
            }}
          >
            <div style={{ fontSize: 24, marginBottom: 8 }}>⚠️</div>
            <div style={{ fontSize: 16, fontWeight: 700, color: "#EF4444", marginBottom: 8 }}>App Error</div>
            <div
              style={{
                fontSize: 12,
                color: "#6B7280",
                marginBottom: 16,
                wordBreak: "break-all",
                background: "#FFFFFF",
                padding: 12,
                borderRadius: 8,
                lineHeight: 1.6,
              }}
            >
              {this.state.error?.message || String(this.state.error)}
            </div>
            <div style={{ fontSize: 11, color: "#6B7280", marginBottom: 16 }}>
              Screenshot this message and send it to your developer.
            </div>
            <button
              onClick={() => {
                this.setState({ error: null });
                window.location.reload();
              }}
              style={{
                background: "#F97316",
                color: "#fff",
                border: "none",
                borderRadius: 10,
                padding: "10px 20px",
                fontSize: 14,
                fontWeight: 700,
                cursor: "pointer",
                fontFamily: "inherit",
                width: "100%",
              }}
            >
              🔄 Reload App
            </button>
          </div>
        </div>
      );
    }
    return this.props.children;
  }
}

/* ── OFFLINE / DRAFT UTILITIES ──────────────────────────────── */
const QUEUE_KEY = "aime_pending_queue";
const DRAFT_PREFIX = "aime_draft_";

function getQueue() {
  try {
    return JSON.parse(localStorage.getItem(QUEUE_KEY) || "[]");
  } catch {
    return [];
  }
}
function addToQueue(item) {
  const q = getQueue();
  q.push({ ...item, qid: Math.random().toString(36).slice(2), queued_at: new Date().toISOString() });
  localStorage.setItem(QUEUE_KEY, JSON.stringify(q));
}
function removeFromQueue(qid) {
  localStorage.setItem(QUEUE_KEY, JSON.stringify(getQueue().filter((i) => i.qid !== qid)));
}
function saveDraft(key, data) {
  try {
    localStorage.setItem(DRAFT_PREFIX + key, JSON.stringify({ data, saved_at: new Date().toISOString() }));
  } catch (e) {
    console.warn("Draft save failed:", e);
  }
}
function loadDraft(key) {
  try {
    const s = localStorage.getItem(DRAFT_PREFIX + key);
    return s ? JSON.parse(s) : null;
  } catch {
    return null;
  }
}
function clearDraft(key) {
  try {
    localStorage.removeItem(DRAFT_PREFIX + key);
  } catch {}
}

// Excel template now lives at public/daily-report-template.xlsx (fetched at export time)

/* ── SUPABASE ───────────────────────────────────────────────── */
const SUPA_URL = import.meta.env.VITE_SUPABASE_URL || "https://uicmfyudiullulbbwzmh.supabase.co";
const SUPA_KEY = import.meta.env.VITE_SUPABASE_KEY || "sb_publishable_9h9AyvXpkp9glLxDVWRuGw_1eKVS7sE";
async function supa(path, { method = "GET", body, prefer } = {}) {
  const r = await fetch(`${SUPA_URL}/rest/v1${path}`, {
    method,
    headers: {
      apikey: SUPA_KEY,
      Authorization: `Bearer ${SUPA_KEY}`,
      "Content-Type": "application/json",
      ...(prefer ? { Prefer: prefer } : {}),
      ...(body !== undefined ? { body: JSON.stringify(body) } : {}),
      ...(body !== undefined ? {} : {}),
    },
    body: body !== undefined ? JSON.stringify(body) : undefined,
  });
  if (!r.ok) throw new Error((await r.text()) || `HTTP ${r.status}`);
  const t = await r.text();
  return t ? JSON.parse(t) : null;
}
// fix: proper fetch helper
async function sb(path, opts = {}) {
  const { method = "GET", body, prefer } = opts;
  const headers = { apikey: SUPA_KEY, Authorization: `Bearer ${SUPA_KEY}`, "Content-Type": "application/json" };
  if (prefer) headers["Prefer"] = prefer;
  const res = await fetch(`${SUPA_URL}/rest/v1${path}`, {
    method,
    headers,
    ...(body !== undefined ? { body: JSON.stringify(body) } : {}),
  });
  if (!res.ok) throw new Error((await res.text()) || `HTTP ${res.status}`);
  const t = await res.text();
  return t ? JSON.parse(t) : null;
}

const API = {
  projects: {
    list: () => sb("/projects?select=*&order=created_at.desc"),
    byDivision: (div) => sb(`/projects?division=eq.${encodeURIComponent(div)}&select=*&order=created_at.desc`),
    create: (d) => sb("/projects", { method: "POST", body: d, prefer: "return=representation" }),
    update: (id, d) => sb(`/projects?id=eq.${id}`, { method: "PATCH", body: d, prefer: "return=representation" }),
    remove: (id) => sb(`/projects?id=eq.${id}`, { method: "DELETE" }),
  },
  reports: {
    forProject: (pid) => sb(`/daily_reports?project_id=eq.${pid}&order=date.desc`),
    all: () => sb("/daily_reports?select=*,projects(id,name,division)&order=date.desc&limit=300"),
    pending: () => sb("/daily_reports?status=eq.submitted&select=*,projects(id,name,division)&order=created_at.desc"),
    create: (d) => sb("/daily_reports", { method: "POST", body: d, prefer: "return=representation" }),
    update: (id, d) => sb(`/daily_reports?id=eq.${id}`, { method: "PATCH", body: d, prefer: "return=representation" }),
    count: (id) => sb(`/daily_reports?id=eq.${id}&select=id`),
    remove: (id) => sb(`/daily_reports?id=eq.${id}`, { method: "DELETE" }),
  },
  safety: {
    forProject: (pid) => sb(`/safety_logs?project_id=eq.${pid}&order=created_at.desc`),
    create: (d) => sb("/safety_logs", { method: "POST", body: d, prefer: "return=representation" }),
    remove: (id) => sb(`/safety_logs?id=eq.${id}`, { method: "DELETE" }),
  },
  photos: {
    forProject: (pid) => sb(`/project_photos?project_id=eq.${pid}&order=created_at.desc`),
    create: (d) => sb("/project_photos", { method: "POST", body: d, prefer: "return=representation" }),
    remove: (id) => sb(`/project_photos?id=eq.${id}`, { method: "DELETE" }),
  },
  timeCards: {
    forProject: (pid) => sb(`/time_cards?project_id=eq.${pid}&order=date.desc,created_at.desc`),
    all: () => sb("/time_cards?order=date.desc,created_at.desc&limit=500"),
    byDate: (date) => sb(`/time_cards?date=eq.${date}&order=worker_name.asc`),
    byRange: (from, to) => sb(`/time_cards?date=gte.${from}&date=lte.${to}&order=date.desc,worker_name.asc`),
    find: (name, date, pid) =>
      sb(`/time_cards?worker_name=eq.${encodeURIComponent(name)}&date=eq.${date}&project_id=eq.${pid}&limit=1`),
    create: (d) => sb("/time_cards", { method: "POST", body: d, prefer: "return=representation" }),
    update: (id, d) => sb(`/time_cards?id=eq.${id}`, { method: "PATCH", body: d, prefer: "return=representation" }),
    remove: (id) => sb(`/time_cards?id=eq.${id}`, { method: "DELETE" }),
  },
  weather: {
    forProject: (pid) => sb(`/weather_logs?project_id=eq.${pid}&order=date.desc&limit=14`),
    upsert: (d) =>
      sb("/weather_logs", { method: "POST", body: d, prefer: "return=representation,resolution=merge-duplicates" }),
    remove: (id) => sb(`/weather_logs?id=eq.${id}`, { method: "DELETE" }),
  },
  equipment: {
    forProject: (pid) => sb(`/equipment_on_site?project_id=eq.${pid}&order=date.desc,created_at.desc`),
    create: (d) => sb("/equipment_on_site", { method: "POST", body: d, prefer: "return=representation" }),
    remove: (id) => sb(`/equipment_on_site?id=eq.${id}`, { method: "DELETE" }),
  },
  subs: {
    forProject: (pid) => sb(`/subcontractors?project_id=eq.${pid}&order=date.desc,created_at.desc`),
    create: (d) => sb("/subcontractors", { method: "POST", body: d, prefer: "return=representation" }),
    remove: (id) => sb(`/subcontractors?id=eq.${id}`, { method: "DELETE" }),
  },
  docs: {
    forProject: (pid) => sb(`/documents?project_id=eq.${pid}&order=created_at.desc`),
    create: (d) => sb("/documents", { method: "POST", body: d, prefer: "return=representation" }),
    update: (id, d) => sb(`/documents?id=eq.${id}`, { method: "PATCH", body: d, prefer: "return=representation" }),
    remove: (id) => sb(`/documents?id=eq.${id}`, { method: "DELETE" }),
  },
  milestones: {
    forProject: (pid) => sb(`/milestones?project_id=eq.${pid}&order=sort_order.asc,target_date.asc`),
    create: (d) => sb("/milestones", { method: "POST", body: d, prefer: "return=representation" }),
    update: (id, d) => sb(`/milestones?id=eq.${id}`, { method: "PATCH", body: d, prefer: "return=representation" }),
    remove: (id) => sb(`/milestones?id=eq.${id}`, { method: "DELETE" }),
  },
  crew: {
    list: () => sb("/crew_members?order=name.asc"),
    create: (d) => sb("/crew_members", { method: "POST", body: d, prefer: "return=representation" }),
    update: (id, d) => sb(`/crew_members?id=eq.${id}`, { method: "PATCH", body: d, prefer: "return=representation" }),
    remove: (id) => sb(`/crew_members?id=eq.${id}`, { method: "DELETE" }),
  },
  notifications: {
    list: () => sb("/notifications?order=created_at.desc&limit=50"),
    unread: () => sb("/notifications?read=eq.false&order=created_at.desc"),
    markRead: (id) => sb(`/notifications?id=eq.${id}`, { method: "PATCH", body: { read: true } }),
    markAllRead: () => sb("/notifications?read=eq.false", { method: "PATCH", body: { read: true } }),
    create: (d) => sb("/notifications", { method: "POST", body: d, prefer: "return=representation" }),
  },
  notifSettings: {
    get: (name) => sb(`/notification_settings?pm_name=eq.${encodeURIComponent(name)}&limit=1`),
    upsert: (d) =>
      sb("/notification_settings", {
        method: "POST",
        body: d,
        prefer: "return=representation,resolution=merge-duplicates",
      }),
  },
  userProfiles: {
    list: () => sb("/user_profiles?order=name.asc"),
    getByName: (name) => sb(`/user_profiles?name=eq.${encodeURIComponent(name)}&limit=1`),
    create: (d) => sb("/user_profiles", { method: "POST", body: d, prefer: "return=representation" }),
    update: (id, d) => sb(`/user_profiles?id=eq.${id}`, { method: "PATCH", body: d, prefer: "return=representation" }),
    upsert: (d) =>
      sb("/user_profiles", { method: "POST", body: d, prefer: "return=representation,resolution=merge-duplicates" }),
    remove: (id) => sb(`/user_profiles?id=eq.${id}`, { method: "DELETE" }),
  },
  catalog: {
    list: (div, cat, search) => {
      let url = "/cost_catalog?active=eq.true&order=name.asc&limit=200";
      if (div) url += `&division=eq.${encodeURIComponent(div)}`;
      if (cat) url += `&category=eq.${encodeURIComponent(cat)}`;
      if (search) url += `&name=ilike.${encodeURIComponent("*" + search + "*")}`;
      return sb(url);
    },
    categories: (div) =>
      sb(`/cost_catalog?active=eq.true&division=eq.${encodeURIComponent(div)}&select=category&order=category.asc`),
    update: (id, d) => sb(`/cost_catalog?id=eq.${id}`, { method: "PATCH", body: d, prefer: "return=representation" }),
    create: (d) => sb("/cost_catalog", { method: "POST", body: d, prefer: "return=representation" }),
    remove: (id) => sb(`/cost_catalog?id=eq.${id}`, { method: "PATCH", body: { active: false } }),
  },
  estimates: {
    list: () => sb("/estimates?order=created_at.desc"),
    create: (d) => sb("/estimates", { method: "POST", body: d, prefer: "return=representation" }),
    update: (id, d) => sb(`/estimates?id=eq.${id}`, { method: "PATCH", body: d, prefer: "return=representation" }),
    remove: (id) => sb(`/estimates?id=eq.${id}`, { method: "DELETE" }),
    items: (eid) => sb(`/estimate_items?estimate_id=eq.${eid}&order=sort_order.asc,category.asc`),
    addItem: (d) => sb("/estimate_items", { method: "POST", body: d, prefer: "return=representation" }),
    updateItem: (id, d) => sb(`/estimate_items?id=eq.${id}`, { method: "PATCH", body: d }),
    removeItem: (id) => sb(`/estimate_items?id=eq.${id}`, { method: "DELETE" }),
  },
  changeOrders: {
    forProject: (pid) => sb(`/change_orders?project_id=eq.${pid}&order=date_submitted.asc`),
    create: (d) => sb("/change_orders", { method: "POST", body: d, prefer: "return=representation" }),
    update: (id, d) => sb(`/change_orders?id=eq.${id}`, { method: "PATCH", body: d, prefer: "return=representation" }),
    remove: (id) => sb(`/change_orders?id=eq.${id}`, { method: "DELETE" }),
  },
  rfis: {
    forProject: (pid) => sb(`/rfis?project_id=eq.${pid}&order=date_submitted.asc`),
    create: (d) => sb("/rfis", { method: "POST", body: d, prefer: "return=representation" }),
    update: (id, d) => sb(`/rfis?id=eq.${id}`, { method: "PATCH", body: d, prefer: "return=representation" }),
    remove: (id) => sb(`/rfis?id=eq.${id}`, { method: "DELETE" }),
  },
};

/* ── THEME ──────────────────────────────────────────────────── */
const T = {
  bg: "#EDEBE6",
  surface: "#FFFFFF",
  card: "#FFFFFF",
  border: "#D8D4CC",
  orange: "#F25C05",
  orangeLow: "#F25C0514",
  orangeMid: "#F25C0530",
  green: "#2E7D46",
  greenLow: "#2E7D4614",
  red: "#C0392B",
  redLow: "#C0392B14",
  yellow: "#B45309",
  yellowLow: "#B4530914",
  blue: "#1D4ED8",
  blueLow: "#1D4ED814",
  purple: "#7C3AED",
  teal: "#0F766E",
  text: "#16181D",
  sub: "#4B5563",
  muted: "#6B7280",
};
const inp = {
  width: "100%",
  boxSizing: "border-box",
  background: "#FFFFFF",
  border: `1px solid ${T.border}`,
  borderRadius: 12,
  color: T.text,
  fontSize: 15,
  padding: "13px 14px",
  outline: "none",
  fontFamily: "inherit",
  appearance: "none",
  WebkitAppearance: "none",
};
const inpSel = { ...inp, color: T.orange };
const lbl = {
  display: "block",
  fontSize: 11,
  fontWeight: 700,
  color: "#6B7280",
  letterSpacing: "1px",
  textTransform: "uppercase",
  marginBottom: 6,
};
const cardS = { background: T.card, border: `1px solid ${T.border}`, borderRadius: 16, padding: "16px" };
const pill = (c) => ({
  display: "inline-flex",
  alignItems: "center",
  background: c + "20",
  color: c,
  borderRadius: 20,
  padding: "3px 10px",
  fontSize: 11,
  fontWeight: 700,
});
const primBtn = {
  width: "100%",
  background: T.orange,
  color: "#FFFFFF",
  border: "none",
  borderRadius: 14,
  padding: "16px",
  fontSize: 16,
  fontWeight: 800,
  cursor: "pointer",
  fontFamily: "inherit",
  display: "flex",
  alignItems: "center",
  justifyContent: "center",
  gap: 8,
};
const ghostBtn = {
  background: "transparent",
  border: `1px solid ${T.border}`,
  borderRadius: 12,
  padding: "12px 16px",
  color: T.sub,
  fontSize: 14,
  cursor: "pointer",
  fontFamily: "inherit",
  fontWeight: 600,
};
const dangerBtn = {
  background: T.redLow,
  border: `1px solid ${T.red}30`,
  borderRadius: 12,
  padding: "12px 16px",
  color: T.red,
  fontSize: 14,
  cursor: "pointer",
  fontFamily: "inherit",
  fontWeight: 600,
  width: "100%",
  textAlign: "center",
};

/* ── DATA ───────────────────────────────────────────────────── */
const POSITIONS_PIPELINE = [
  { name: "Project Manager", rate: 64.5 },
  { name: "Foreman", rate: 63.25 },
  { name: "Technician", rate: 60.75 },
  { name: "Inspector", rate: 53.75 },
  { name: "Certified Welder", rate: 60.75 },
  { name: "Fitter", rate: 58.5 },
  { name: "Mechanic", rate: 58.5 },
  { name: "Operator", rate: 58.5 },
  { name: "Truck Driver", rate: 58.5 },
  { name: "Helper (Welder)", rate: 57.25 },
  { name: "Laborer", rate: 51.0 },
  { name: "Foreman (Elect)", rate: 82.25 },
  { name: "Electrician", rate: 82.25 },
  { name: "Helper (Elect)", rate: 45.5 },
  { name: "Per Diem", rate: 190.0, flat: true },
];
const NAMES = [
  "Alan Fairbrother",
  "Alan Robinson",
  "Doug Friedel",
  "Jaden Pugh",
  "Brandon Milano",
  "Charles Acree",
  "Charles Dovel",
  "Chris Utz",
  "Christopher Dean",
  "Chuck Dean",
  "Clay Lau",
  "Connor Kestner",
  "Morgan Schramek",
  "Eric Bowens",
  "Eric Shumate",
  "Jackson Fama",
  "Howard Lau",
  "Jeff White",
  "Jessica Vance",
  "John Baier",
  "John P. Cosner Jr.",
  "Jordan Gorwell",
  "Joseph Lau",
  "Josh Gladhill",
  "Kevin Gabrish",
  "Kurt Batterton",
  "Leo Velez",
  "Edgrado Ruiz",
  "Mark Hamilton",
  "Alejandro Figueroa",
  "Matthew Linton",
  "Mike Gamble",
  "Mike Gamble III",
  "Mike Seiler",
  "Pat Gorman",
  "Paul Howard",
  "Rich Raborg",
  "Robert Neslein",
  "Roland Long",
  "Shane Hower",
  "Steve Kestner",
  "Tom Hatfield",
  "Troy Strother",
  "Tyrone Davis",
  "Walter Chicas-Luna",
  "Will Wychulis",
  "Wyatt Gill",
].sort();
const EQUIP_LIST_PIPELINE = [
  { section: "Trucks & Trailers" },
  { name: "Truck - 1 Ton", rate: 21.5, unit: "Hours" },
  { name: "Truck - 3/4 Ton w/ Snow Plow", rate: 350, unit: "Days" },
  { name: "Truck - 1/2 Ton", rate: 18.5, unit: "Hours" },
  { name: "Truck - Boom (20-29 Ton)", rate: 65, unit: "Hours" },
  { name: "Truck - Bucket", rate: 45, unit: "Hours" },
  { name: "Truck - Dump Truck (3 Axle)", rate: 35, unit: "Hours" },
  { name: "Truck - Haul Truck - No Trailer", rate: 70, unit: "Hours" },
  { name: "Truck - Tru-Vac", rate: 13500, unit: "Month" },
  { name: "Truck - Welding Rig", rate: 35, unit: "Hours" },
  { name: "Trailer - Electrical - Colonial", rate: 147, unit: "Month" },
  { name: "Trailer - Lowboy - 2 Axle", rate: 28, unit: "Hours" },
  { name: "Trailer - Tag Along", rate: 50, unit: "Days" },
  { name: "Trailer - Tool Trailer - 18-25ft", rate: 175, unit: "Days" },
  { name: "Trailer - Tool Trailer - 26-40ft", rate: 200, unit: "Days" },
  { section: "Earthmoving & ROW" },
  { name: "ATV - 4 Wheel", rate: 125, unit: "Days" },
  { name: "Backhoe Loader - 80-105 HP", rate: 62.45, unit: "Hours" },
  { name: "Excavator - Mini - 2-8K LB", rate: 299, unit: "Days" },
  { name: "Excavator - Mini - 9K LB", rate: 335, unit: "Days" },
  { name: "Excavator - Mini - 12-16K LB", rate: 475, unit: "Days" },
  { name: "Excavator - Small - 21-29K LB", rate: 565, unit: "Days" },
  { name: "Excavator - Small - 30-33K LB", rate: 632, unit: "Days" },
  { name: "Excavator - Medium - 48-55K LB", rate: 852, unit: "Days" },
  { name: "Excavator - Large - 80-89K LB", rate: 1050, unit: "Days" },
  { name: "Excavator - Large - 90-119K LB", rate: 1350, unit: "Days" },
  { name: "Skidsteer Loader - 70-80 HP", rate: 440, unit: "Days" },
  { name: "Skidsteer Loader - 81-100 HP", rate: 475, unit: "Days" },
  { name: "Tractor - 50 HP 4x4 w/ Bush Hog", rate: 36.5, unit: "Hours" },
  { name: "Mower - Riding/Zero Turn", rate: 175, unit: "Days" },
  { section: "Air, Compressors & Blast" },
  { name: "Air Compressor - 185 CFM", rate: 195, unit: "Days" },
  { name: "Air Compressor - 375 CFM", rate: 275, unit: "Days" },
  { name: "Air Impact Wrench - 1in", rate: 50, unit: "Days" },
  { name: "Air Spade / Knife", rate: 55, unit: "Days" },
  { name: "Blast Rig - 4 Bag Pot w/ 185 CFM AC", rate: 55.5, unit: "Hours" },
  { name: "Blast Rig - 1 Pot w/ 375 CFM AC", rate: 500, unit: "Days" },
  { section: "Testing & Misc. Tools" },
  { name: "Holiday Detector / Pipe Jeep", rate: 72, unit: "Days" },
  { name: "Hydraulic Torque", rate: 200, unit: "Days" },
  { name: "Hydro Test Pump", rate: 60, unit: "Days" },
  { name: "Hydrotest - High Pressure", rate: 3800, unit: "Days" },
  { name: "Jack Hammer", rate: 72, unit: "Days" },
  { name: "LEL/Gas Monitor - 4 Gas", rate: 50, unit: "Days" },
  { name: "Line Locator", rate: 50, unit: "Days" },
  { name: "HEPA Vacuum", rate: 100, unit: "Days" },
  { name: "Torque Wrench w/Sockets Hyd/Pneu", rate: 195, unit: "Days" },
  { name: "Pipe Beveling Machine 16-22in", rate: 100, unit: "Days" },
];

const POSITIONS_MECHANICAL = [
  { name: "Project Manager", rate: 93.0 },
  { name: "Foreman", rate: 72.5 },
  { name: "Inspector", rate: 140.0 },
  { name: "Certified Welder", rate: 69.5 },
  { name: "Mechanic", rate: 69.5 },
  { name: "Operator", rate: 69.5 },
  { name: "Truck Driver", rate: 69.5 },
  { name: "Dock Watch", rate: 70.0 },
  { name: "Per Diem", rate: 190.0, flat: true },
];

const EQUIP_LIST_MECHANICAL = [
  { section: "Trucks & Trailers" },
  { name: "Truck - Tool Truck", rate: 40, unit: "Hours" },
  { name: "Truck - 3/4 Ton w/ Snow Plow", rate: 100, unit: "Hours" },
  { name: "Truck - Boom (20-29 Ton)", rate: 98, unit: "Hours" },
  { name: "Truck - Dump Truck", rate: 40, unit: "Hours" },
  { name: "Truck - Welding Rig", rate: 40, unit: "Hours" },
  { name: "Trailer - Lowboy - 2 Axle", rate: 29, unit: "Hours" },
  { name: "Trailer - Tag Along", rate: 50, unit: "Days" },
  { name: "Trailer - Tool Trailer - 18-25ft", rate: 175, unit: "Days" },
  { section: "Earthmoving" },
  { name: "Backhoe", rate: 450, unit: "Days" },
  { name: "Bobcat Skidsteer", rate: 25, unit: "Hours" },
  { section: "Air, Compressors & Blast" },
  { name: "Air Compressor - 185 CFM", rate: 195, unit: "Days" },
  { name: "Air Impact Wrench - 1in", rate: 50, unit: "Days" },
  { name: "Air Spade / Knife", rate: 55, unit: "Days" },
  { name: "Blast Rig - 4 Bag Pot w/ 185 CFM AC", rate: 55.5, unit: "Hours" },
  { name: "Blast Rig - 1 Pot w/ 375 CFM AC", rate: 500, unit: "Days" },
  { section: "Tools & Testing" },
  { name: "Bench & Volt Meter", rate: 875.5, unit: "Month" },
  { name: "Beveling Band - 30in", rate: 25, unit: "Days" },
  { name: "Dearman Pipe Clamps", rate: 25, unit: "Days" },
  { name: "Gasoline Emergency Response Equipment", rate: 5000, unit: "Week" },
  { name: "HEPA Vacuum", rate: 100, unit: "Days" },
  { name: "Holiday Detector / Pipe Jeep", rate: 72, unit: "Days" },
  { name: "Hydraulic Torque", rate: 200, unit: "Days" },
  { name: "German Air Saw (plus Blades)", rate: 125, unit: "Days" },
  { name: "Hydrotest / Pressure Test Equipment", rate: 200, unit: "Days" },
  { name: "Jack Hammer", rate: 72, unit: "Days" },
  { name: "Laser Pump Aligner", rate: 275, unit: "Days" },
  { name: "LEL/Gas Monitor - 4 Gas", rate: 50, unit: "Days" },
  { name: "Line Locator", rate: 50, unit: "Days" },
  { name: "Pipe Band Crawler", rate: 25, unit: "Days" },
  { name: "Pipe Beveling Machine 1.5-3in", rate: 25, unit: "Days" },
  { name: "Pipe Beveling Machine 10-14in", rate: 40, unit: "Days" },
  { name: "Pipe Beveling Machine 16-22in", rate: 100, unit: "Days" },
  { name: "Tap Machine - 2in", rate: 190, unit: "Days" },
  { name: "Torque Wrench - Pneumatic J5", rate: 175, unit: "Days" },
  { name: "Torque Wrench w/Multiplier Hand", rate: 25, unit: "Days" },
  { name: "Torque Wrench w/Sockets Hyd/Pneu", rate: 195, unit: "Days" },
  { name: "Wach Pipe Cutting Saw", rate: 30, unit: "Hours" },
  { name: "Cold Cutters 0.5-2in", rate: 20, unit: "Days" },
  { name: "Cold Cutters 2-4in", rate: 30, unit: "Days" },
  { name: "Cold Cutters 4-8in", rate: 40, unit: "Days" },
  { name: "Cold Cutters 8-12in", rate: 60, unit: "Days" },
  { name: "Cold Cutters 12-14in", rate: 75, unit: "Days" },
  { name: "Concrete Saw (plus blades)", rate: 50, unit: "Days" },
  { name: "Confined Space Equipment", rate: 375, unit: "Days" },
  { name: "Hydrotest Pump", rate: 125, unit: "Days" },
  { section: "Weld Rates" },
  { name: "1G Weld", rate: 8, unit: "Ft" },
  { name: "3G Weld", rate: 8.5, unit: "Ft" },
  { name: "4G Weld", rate: 9, unit: "Ft" },
];

// Division-aware helpers
function getPositions(division) {
  return division === "Mechanical" ? POSITIONS_MECHANICAL : POSITIONS_PIPELINE;
}
function getEquipList(division) {
  return division === "Mechanical" ? EQUIP_LIST_MECHANICAL : EQUIP_LIST_PIPELINE;
}
// Keep POSITIONS and EQUIP_LIST as aliases for backward compat
const POSITIONS = POSITIONS_PIPELINE;
const EQUIP_LIST = EQUIP_LIST_PIPELINE;
// All unique position names across all divisions (for Crew Directory)
function getAllPositions() {
  const all = [...POSITIONS_PIPELINE, ...POSITIONS_MECHANICAL];
  const seen = new Set();
  return all.filter((p) => {
    if (seen.has(p.name)) return false;
    seen.add(p.name);
    return true;
  });
}

const WMO = {
  0: ["Clear Sky", "☀️"],
  1: ["Mainly Clear", "🌤️"],
  2: ["Partly Cloudy", "⛅"],
  3: ["Overcast", "☁️"],
  45: ["Foggy", "🌫️"],
  48: ["Icy Fog", "🌫️"],
  51: ["Light Drizzle", "🌦️"],
  53: ["Drizzle", "🌦️"],
  55: ["Heavy Drizzle", "🌦️"],
  61: ["Light Rain", "🌧️"],
  63: ["Rain", "🌧️"],
  65: ["Heavy Rain", "🌧️"],
  71: ["Light Snow", "🌨️"],
  73: ["Snow", "🌨️"],
  75: ["Heavy Snow", "❄️"],
  80: ["Light Showers", "🌦️"],
  81: ["Showers", "🌦️"],
  82: ["Violent Showers", "⛈️"],
  95: ["Thunderstorm", "⛈️"],
  96: ["Thunderstorm + Hail", "⛈️"],
  99: ["Severe Thunderstorm", "⛈️"],
};
const DIVISIONS = ["Mechanical", "Pipeline", "Structural"];
const DIV_META = {
  Mechanical: { icon: "⚙️", color: "#F97316", desc: "Mechanical projects and equipment" },
  Pipeline: { icon: "🔧", color: "#3B82F6", desc: "Pipeline construction and maintenance" },
  Structural: { icon: "🏗️", color: "#22C55E", desc: "Structural steel and civil work" },
};
const ROLES = ["crew", "foreman", "pm", "admin"];
const ROLE_META = {
  crew: { label: "Field Crew", color: T.green, desc: "Submit daily reports and time cards" },
  foreman: { label: "Foreman", color: T.yellow, desc: "Reports, time, safety, equipment, docs, schedule" },
  pm: { label: "Project Manager", color: T.orange, desc: "Approve reports, manage jobs, PM dashboard" },
  estimator: { label: "Estimator", color: T.purple, desc: "Estimating platform access only" },
  admin: { label: "Admin", color: T.red, desc: "Full access, user management" },
};

/* ── PERMISSIONS ────────────────────────────────────────────── */
const PERMS = {
  admin: [
    "manage_users",
    "create_job",
    "edit_job",
    "archive_job",
    "approve_report",
    "flag_report",
    "view_dashboard",
    "submit_report",
    "time_card",
    "safety",
    "photos",
    "docs",
    "schedule",
    "weather",
    "subs",
    "crew_equip",
    "crew_directory",
    "custom_reports",
    "notifications",
    "estimating",
  ],
  pm: [
    "create_job",
    "edit_job",
    "archive_job",
    "approve_report",
    "flag_report",
    "view_dashboard",
    "submit_report",
    "time_card",
    "safety",
    "photos",
    "docs",
    "schedule",
    "weather",
    "subs",
    "crew_equip",
    "crew_directory",
    "custom_reports",
    "notifications",
  ],
  estimator: ["estimating", "view_dashboard", "crew_directory"],
  foreman: [
    "submit_report",
    "time_card",
    "safety",
    "photos",
    "docs",
    "schedule",
    "weather",
    "subs",
    "crew_equip",
    "crew_directory",
  ],
  crew: ["submit_report", "time_card", "photos", "crew_directory"],
};
const can = (user, action) => (PERMS[user?.role] || PERMS.crew).includes(action);

/* ── HELPERS ────────────────────────────────────────────────── */
const uid = () => Math.random().toString(36).slice(2, 9);
const today = () => new Date().toISOString().split("T")[0];
const fmt = (n) => Number(n || 0).toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
const fmtDate = (d) =>
  d ? new Date(d + "T12:00:00").toLocaleDateString("en-US", { month: "short", day: "numeric", year: "numeric" }) : "—";
const fmtShort = (d) =>
  d ? new Date(d + "T12:00:00").toLocaleDateString("en-US", { weekday: "short", month: "short", day: "numeric" }) : "—";
const daysUntil = (d) => {
  if (!d) return null;
  const diff = new Date(d + "T12:00:00") - new Date();
  return Math.ceil(diff / 86400000);
};
function laborAmt(r, division) {
  const positions = getPositions(division);
  const p = positions.find((x) => x.name === r.classification);
  if (!p) return 0;
  if (p.flat) return p.rate;
  return p.rate * ((parseFloat(r.regHrs) || 0) + (parseFloat(r.otHrs) || 0) * 1.5 + (parseFloat(r.travelHrs) || 0));
}
function equipAmt(r, division) {
  let rate = parseFloat(r.rate) || 0;
  if (!rate && r.description) {
    const eq = getEquipList(division).find((e) => !e.section && e.name === r.description);
    if (eq) rate = eq.rate;
  }
  const qty = parseFloat(r.qty) || 0;
  const usage = parseFloat(r.usage) || 0;
  return qty * rate * (usage || 1);
}
function reportTotals(r, division) {
  const labor = (r.labor || []).reduce((s, x) => s + laborAmt(x, division), 0);
  const equip = (r.equipment || []).reduce((s, x) => s + equipAmt(x, division), 0);
  const rental = (r.rental_equipment || []).reduce(
    (s, x) => s + (parseFloat(x.qty) || 0) * (parseFloat(x.rate) || 0) * (parseFloat(x.usage) || 1),
    0,
  );
  const mats = (r.materials || []).reduce((s, x) => s + (parseFloat(x.amount) || 0), 0);
  return { labor, equip, rental, mats, grand: labor + equip + rental + mats };
}
function calcHours(ci, co) {
  if (!ci || !co) return 0;
  const [ih, im] = ci.split(":").map(Number);
  const [oh, om] = co.split(":").map(Number);
  const diff = oh * 60 + om - (ih * 60 + im);
  return diff > 0 ? Math.round((diff / 60) * 100) / 100 : 0;
}
function getWeekStart() {
  const d = new Date();
  const day = d.getDay();
  d.setDate(d.getDate() - (day === 0 ? 6 : day - 1));
  return d.toISOString().split("T")[0];
}
async function compressImg(file, maxW = 900, q = 0.65) {
  return new Promise((res) => {
    const rd = new FileReader();
    rd.onload = (ev) => {
      const img = new Image();
      img.onload = () => {
        const sc = Math.min(1, maxW / img.width);
        const c = document.createElement("canvas");
        c.width = Math.round(img.width * sc);
        c.height = Math.round(img.height * sc);
        c.getContext("2d").drawImage(img, 0, 0, c.width, c.height);
        res(c.toDataURL("image/jpeg", q));
      };
      img.src = ev.target.result;
    };
    rd.readAsDataURL(file);
  });
}
async function fetchWeather(location) {
  const gR = await fetch(
    `https://geocoding-api.open-meteo.com/v1/search?name=${encodeURIComponent(location)}&count=1&language=en&format=json`,
  );
  const gD = await gR.json();
  if (!gD.results?.length) throw new Error(`Cannot find: "${location}"`);
  const { latitude: lat, longitude: lon, name, admin1 } = gD.results[0];
  const wR = await fetch(
    `https://api.open-meteo.com/v1/forecast?latitude=${lat}&longitude=${lon}&current=temperature_2m,apparent_temperature,weathercode,windspeed_10m,precipitation&daily=temperature_2m_max,temperature_2m_min,precipitation_sum,weathercode&timezone=auto&temperature_unit=fahrenheit&windspeed_unit=mph&precipitation_unit=inch&forecast_days=1`,
  );
  const wD = await wR.json();
  return { ...wD, locationName: `${name}, ${admin1}` };
}
async function notify(type, title, body, extra = {}) {
  try {
    await API.notifications.create({ type, title, body, ...extra });
  } catch {}
}

/* ── SHARED UI ──────────────────────────────────────────────── */
function Spinner() {
  return (
    <div style={{ display: "flex", justifyContent: "center", padding: "48px 0" }}>
      <div
        style={{
          width: 32,
          height: 32,
          border: `3px solid ${T.border}`,
          borderTopColor: T.orange,
          borderRadius: "50%",
          animation: "spin 0.7s linear infinite",
        }}
      />
      <style>{`@keyframes spin{to{transform:rotate(360deg)}} select,select option{background:#1a1a1a !important;color:#F97316 !important;} select:focus{outline:none;} select *{color:#F97316 !important;background:#1a1a1a !important;}`}</style>
    </div>
  );
}
function ErrBanner({ msg, onDismiss }) {
  if (!msg) return null;
  return (
    <div
      style={{
        background: T.redLow,
        border: `1px solid ${T.red}40`,
        borderRadius: 12,
        padding: "12px 14px",
        marginBottom: 14,
        display: "flex",
        justifyContent: "space-between",
        alignItems: "center",
      }}
    >
      <span style={{ fontSize: 13, color: T.red }}>⚠️ {msg}</span>
      <button
        onClick={onDismiss}
        style={{
          background: "none",
          border: "none",
          color: T.red,
          cursor: "pointer",
          fontSize: 18,
          padding: "0 0 0 10px",
        }}
      >
        ×
      </button>
    </div>
  );
}
function Lightbox({ src, onClose }) {
  if (!src) return null;
  return (
    <div
      onClick={onClose}
      style={{
        position: "fixed",
        inset: 0,
        zIndex: 999,
        background: "rgba(0,0,0,0.95)",
        display: "flex",
        alignItems: "center",
        justifyContent: "center",
        padding: 16,
      }}
    >
      <img
        src={src}
        alt=""
        style={{ maxWidth: "100%", maxHeight: "90vh", borderRadius: 12 }}
        onClick={(e) => e.stopPropagation()}
      />
      <button
        onClick={onClose}
        style={{
          position: "absolute",
          top: 16,
          right: 16,
          background: "#6B7280",
          border: "none",
          color: "#fff",
          borderRadius: "50%",
          width: 36,
          height: 36,
          fontSize: 18,
          cursor: "pointer",
        }}
      >
        ×
      </button>
    </div>
  );
}
function DashedAdd({ label, onClick, color }) {
  const c = color || T.muted;
  return (
    <button
      onClick={onClick}
      style={{
        width: "100%",
        border: `2px dashed ${c}50`,
        background: c + "08",
        color: c,
        borderRadius: 14,
        padding: "14px",
        fontSize: 15,
        cursor: "pointer",
        fontFamily: "inherit",
        fontWeight: 600,
      }}
    >
      {label}
    </button>
  );
}
function StatBar({ items }) {
  return (
    <div style={{ display: "grid", gridTemplateColumns: `repeat(${items.length},1fr)`, gap: 8 }}>
      {items.map(({ label, val, color }) => (
        <div
          key={label}
          style={{
            background: T.card,
            border: `1px solid ${T.border}`,
            borderRadius: 12,
            padding: "10px 8px",
            textAlign: "center",
          }}
        >
          <div style={{ fontSize: 16, fontWeight: 900, color: color || T.text }}>{val}</div>
          <div
            style={{ fontSize: 9, color: T.muted, textTransform: "uppercase", letterSpacing: "0.7px", marginTop: 2 }}
          >
            {label}
          </div>
        </div>
      ))}
    </div>
  );
}
function TopBar({ title, sub, onBack, right }) {
  return (
    <div
      style={{
        background: T.surface,
        borderBottom: `1px solid ${T.border}`,
        padding: "14px 16px",
        position: "sticky",
        top: 0,
        zIndex: 50,
      }}
    >
      {onBack && (
        <button
          onClick={onBack}
          style={{
            background: "none",
            border: "none",
            color: T.sub,
            fontSize: 13,
            cursor: "pointer",
            marginBottom: 8,
            padding: 0,
            fontFamily: "inherit",
          }}
        >
          ← Back
        </button>
      )}
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start" }}>
        <div style={{ flex: 1, minWidth: 0 }}>
          <div style={{ fontSize: 20, fontWeight: 900, color: T.text, letterSpacing: "-0.5px" }}>{title}</div>
          {sub && <div style={{ fontSize: 12, color: T.muted, marginTop: 2 }}>{sub}</div>}
        </div>
        {right && <div style={{ flexShrink: 0, marginLeft: 12 }}>{right}</div>}
      </div>
    </div>
  );
}

/* ── FORM CARDS ─────────────────────────────────────────────── */
function LaborCard({ row, onChange, onRemove, division }) {
  const positions = getPositions(division);
  const pos = positions.find((p) => p.name === row.classification);
  const amt = laborAmt(row, division);
  const set = (k, v) => {
    const u = { ...row, [k]: v };
    if (k === "classification") {
      const p = getPositions(division).find((x) => x.name === v);
      u.rate = p ? p.rate : "";
    }
    onChange(u);
  };
  return (
    <div style={{ ...cardS, marginBottom: 10, borderLeft: `3px solid ${T.orange}` }}>
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8, marginBottom: 8 }}>
        <div style={{ gridColumn: "1/-1" }}>
          <label style={lbl}>Name</label>
          <select value={row.name || ""} onChange={(e) => set("name", e.target.value)} style={inpSel}>
            <option value="">— Select —</option>
            {NAMES.map((n) => (
              <option key={n}>{n}</option>
            ))}
          </select>
        </div>
        <div style={{ gridColumn: "1/-1" }}>
          <label style={lbl}>Classification</label>
          <select
            value={row.classification || ""}
            onChange={(e) => set("classification", e.target.value)}
            style={inpSel}
          >
            <option value="">— Select —</option>
            {getAllPositions().map((p) => (
              <option key={p.name}>{p.name}</option>
            ))}
          </select>
        </div>
      </div>
      {pos && !pos.flat && (
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 8, marginBottom: 8 }}>
          {[
            ["regHrs", "Reg Hrs"],
            ["otHrs", "OT Hrs"],
            ["travelHrs", "Travel"],
          ].map(([k, l]) => (
            <div key={k}>
              <label style={lbl}>{l}</label>
              <input
                type="number"
                min="0"
                step="0.5"
                placeholder="0"
                value={row[k] || ""}
                onChange={(e) => set(k, e.target.value)}
                style={inp}
              />
            </div>
          ))}
        </div>
      )}
      <div
        style={{
          display: "flex",
          justifyContent: "space-between",
          alignItems: "center",
          paddingTop: 8,
          borderTop: `1px solid ${T.border}`,
        }}
      >
        <span style={{ fontSize: 11, color: T.muted }}>
          {pos ? `$${pos.rate.toFixed(2)}${pos.flat ? " flat" : "/hr"}` : ""}
        </span>
        <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
          {amt > 0 && <span style={{ fontSize: 16, fontWeight: 800, color: T.green }}>${fmt(amt)}</span>}
          <button
            onClick={onRemove}
            style={{ background: "none", border: "none", color: T.red, cursor: "pointer", fontSize: 20, padding: 0 }}
          >
            ×
          </button>
        </div>
      </div>
    </div>
  );
}
function EquipCard({ row, onChange, onRemove, division }) {
  const eqList = getEquipList(division);
  const eq = eqList.find((e) => !e.section && e.name === row.description);
  const amt = equipAmt(row);
  const set = (k, v) => {
    const u = { ...row, [k]: v };
    if (k === "description") {
      const e = eqList.find((x) => !x.section && x.name === v);
      u.rate = e ? e.rate : "";
      u.unit = e ? e.unit : "";
    }
    onChange(u);
  };
  return (
    <div style={{ ...cardS, marginBottom: 10, borderLeft: `3px solid ${T.yellow}` }}>
      <div style={{ marginBottom: 8 }}>
        <label style={lbl}>Equipment</label>
        <select value={row.description || ""} onChange={(e) => set("description", e.target.value)} style={inpSel}>
          <option value="">— Select —</option>
          {eqList.map((e, i) =>
            e.section ? (
              <option key={i} disabled>
                ── {e.section} ──
              </option>
            ) : (
              <option key={i} value={e.name}>
                {e.name}
              </option>
            ),
          )}
        </select>
      </div>
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8, marginBottom: 8 }}>
        <div>
          <label style={lbl}>Qty</label>
          <input
            type="number"
            min="0"
            placeholder="0"
            value={row.qty || ""}
            onChange={(e) => set("qty", e.target.value)}
            style={inp}
          />
        </div>
        <div>
          <label style={lbl}>
            {eq ? (
              <span>
                {eq.unit === "Hours"
                  ? "⏱️ Hours"
                  : eq.unit === "Days"
                    ? "📅 Days"
                    : eq.unit === "Ft"
                      ? "📏 Feet"
                      : eq.unit === "Week"
                        ? "📅 Weeks"
                        : eq.unit === "Month"
                          ? "📅 Months"
                          : "📊 " + eq.unit}
              </span>
            ) : (
              "Hrs / Days"
            )}
          </label>
          <input
            type="number"
            min="0"
            step="0.5"
            placeholder="0"
            value={row.usage || ""}
            onChange={(e) => set("usage", e.target.value)}
            style={inp}
          />
        </div>
      </div>
      <div
        style={{
          display: "flex",
          justifyContent: "space-between",
          alignItems: "center",
          paddingTop: 8,
          borderTop: `1px solid ${T.border}`,
        }}
      >
        <span style={{ fontSize: 11, color: T.muted }}>{eq ? `$${eq.rate.toLocaleString()}/${eq.unit}` : ""}</span>
        <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
          {amt > 0 && <span style={{ fontSize: 16, fontWeight: 800, color: T.green }}>${fmt(amt)}</span>}
          <button
            onClick={onRemove}
            style={{ background: "none", border: "none", color: T.red, cursor: "pointer", fontSize: 20, padding: 0 }}
          >
            ×
          </button>
        </div>
      </div>
    </div>
  );
}
function RentedEquipCard({ row, onChange, onRemove }) {
  const set = (k, v) => onChange({ ...row, [k]: v });
  const amt = (parseFloat(row.qty) || 0) * (parseFloat(row.rate) || 0) * (parseFloat(row.usage) || 1);
  return (
    <div style={{ ...cardS, marginBottom: 10, borderLeft: `3px solid ${T.purple}` }}>
      <div style={{ marginBottom: 8 }}>
        <label style={lbl}>Equipment Name / Description</label>
        <input
          type="text"
          placeholder="e.g. 40-Ton Crane, Scissor Lift, Generator…"
          value={row.description || ""}
          onChange={(e) => set("description", e.target.value)}
          style={inp}
        />
      </div>
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 8, marginBottom: 8 }}>
        <div>
          <label style={lbl}>Qty</label>
          <input
            type="number"
            min="0"
            placeholder="0"
            value={row.qty || ""}
            onChange={(e) => set("qty", e.target.value)}
            style={inp}
          />
        </div>
        <div>
          <label style={lbl}>Days / Hrs</label>
          <input
            type="number"
            min="0"
            step="0.5"
            placeholder="0"
            value={row.usage || ""}
            onChange={(e) => set("usage", e.target.value)}
            style={inp}
          />
        </div>
        <div>
          <label style={lbl}>Rate / Unit</label>
          <input
            type="number"
            min="0"
            placeholder="0.00"
            value={row.rate || ""}
            onChange={(e) => set("rate", e.target.value)}
            style={inp}
          />
        </div>
      </div>
      <div
        style={{
          display: "flex",
          justifyContent: "space-between",
          alignItems: "center",
          paddingTop: 8,
          borderTop: `1px solid ${T.border}`,
        }}
      >
        <span style={{ fontSize: 11, color: T.muted }}>Qty × Rate × Days/Hrs</span>
        <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
          {amt > 0 && <span style={{ fontSize: 16, fontWeight: 800, color: T.green }}>${fmt(amt)}</span>}
          <button
            onClick={onRemove}
            style={{ background: "none", border: "none", color: T.red, cursor: "pointer", fontSize: 20, padding: 0 }}
          >
            ×
          </button>
        </div>
      </div>
    </div>
  );
}

function MatCard({ row, onChange, onRemove }) {
  const fileRef = useRef(null);
  const receipts = row.receipts || [];
  async function handleFiles(files) {
    const n = [];
    for (const f of files) {
      if (!f.type.startsWith("image/")) continue;
      const src = await compressImg(f, 800, 0.6);
      n.push({ id: uid(), src });
    }
    onChange({ ...row, receipts: [...receipts, ...n] });
  }
  return (
    <div style={{ ...cardS, marginBottom: 10, borderLeft: `3px solid ${T.blue}` }}>
      <div style={{ display: "grid", gridTemplateColumns: "56px 1fr 88px", gap: 8, marginBottom: 10 }}>
        <div>
          <label style={lbl}>Qty</label>
          <input
            type="number"
            min="0"
            placeholder="0"
            value={row.qty || ""}
            onChange={(e) => onChange({ ...row, qty: e.target.value })}
            style={inp}
          />
        </div>
        <div>
          <label style={lbl}>Description</label>
          <input
            type="text"
            placeholder="Item / material"
            value={row.description || ""}
            onChange={(e) => onChange({ ...row, description: e.target.value })}
            style={inp}
          />
        </div>
        <div>
          <label style={lbl}>Amount</label>
          <input
            type="number"
            min="0"
            placeholder="0.00"
            value={row.amount || ""}
            onChange={(e) => onChange({ ...row, amount: e.target.value })}
            style={inp}
          />
        </div>
      </div>
      <div style={{ borderTop: `1px solid ${T.border}`, paddingTop: 10 }}>
        <label style={{ ...lbl, marginBottom: 8 }}>📎 Receipts</label>
        <div style={{ display: "flex", gap: 8, flexWrap: "wrap", alignItems: "center" }}>
          {receipts.map((r) => (
            <div key={r.id} style={{ position: "relative" }}>
              <img
                src={r.src}
                alt=""
                style={{
                  width: 60,
                  height: 60,
                  objectFit: "cover",
                  borderRadius: 10,
                  border: `2px solid ${T.blue}40`,
                  display: "block",
                }}
              />
              <button
                onClick={() => onChange({ ...row, receipts: receipts.filter((x) => x.id !== r.id) })}
                style={{
                  position: "absolute",
                  top: -5,
                  right: -5,
                  width: 18,
                  height: 18,
                  borderRadius: "50%",
                  background: T.red,
                  border: "none",
                  color: "#fff",
                  fontSize: 11,
                  cursor: "pointer",
                  display: "flex",
                  alignItems: "center",
                  justifyContent: "center",
                }}
              >
                ×
              </button>
            </div>
          ))}
          <button
            onClick={() => fileRef.current?.click()}
            style={{
              width: 60,
              height: 60,
              borderRadius: 10,
              border: `2px dashed ${T.blue}40`,
              background: T.blueLow,
              color: T.blue,
              cursor: "pointer",
              display: "flex",
              flexDirection: "column",
              alignItems: "center",
              justifyContent: "center",
              fontSize: 18,
              gap: 2,
            }}
          >
            <span>📷</span>
            <span style={{ fontSize: 9, fontWeight: 700 }}>ADD</span>
          </button>
          <input
            ref={fileRef}
            type="file"
            accept="image/*"
            capture="environment"
            multiple
            style={{ display: "none" }}
            onChange={(e) => {
              handleFiles(Array.from(e.target.files));
              e.target.value = "";
            }}
          />
        </div>
      </div>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginTop: 10 }}>
        {row.amount > 0 && (
          <span style={{ fontSize: 14, fontWeight: 700, color: T.green }}>${fmt(parseFloat(row.amount) || 0)}</span>
        )}
        <button
          onClick={onRemove}
          style={{
            background: "none",
            border: "none",
            color: T.red,
            cursor: "pointer",
            fontSize: 12,
            fontWeight: 700,
            fontFamily: "inherit",
            marginLeft: "auto",
          }}
        >
          Remove
        </button>
      </div>
    </div>
  );
}

/* ── LOGIN SCREEN ───────────────────────────────────────────── */
function LoginScreen({ onLogin }) {
  const [name, setName] = useState("");
  const [pin, setPin] = useState("");
  const [profile, setProfile] = useState(null);
  const [loading, setLoading] = useState(false);
  const [err, setErr] = useState("");

  async function handleNameChange(n) {
    setName(n);
    setPin("");
    setErr("");
    setProfile(null);
    if (!n) return;
    setLoading(true);
    try {
      const rows = await API.userProfiles.getByName(n);
      setProfile(rows && rows.length > 0 ? rows[0] : { name: n, role: "crew", division: null, pin: null });
    } catch {
      setProfile({ name: n, role: "crew", division: null, pin: null });
    }
    setLoading(false);
  }

  async function handleLogin() {
    if (!name || !profile) return;
    // If profile has no PIN set yet, allow login but warn admin to set one
    if (profile.pin) {
      if (pin !== profile.pin) {
        setErr("Incorrect PIN");
        return;
      }
    } else if (name === "Admin") {
      // Admin fallback PIN
      if (pin !== "1234") {
        setErr("Incorrect PIN (default: 1234)");
        return;
      }
    }
    // If no PIN set for regular user, allow in but note it
    onLogin(profile);
  }

  const roleM = profile ? ROLE_META[profile.role] : null;
  const pinIsSet = profile && (profile.pin || name === "Admin");

  return (
    <div
      style={{
        minHeight: "100vh",
        background: T.bg,
        display: "flex",
        flexDirection: "column",
        justifyContent: "center",
        padding: 24,
        fontFamily: "inherit",
      }}
    >
      <div style={{ textAlign: "center", marginBottom: 40 }}>
        <div style={{ display: "inline-flex", alignItems: "center", gap: 12, marginBottom: 12 }}>
          <div style={{ textAlign: "center" }}>
            <img
              src="/aime-report-logo.png"
              alt="AIME"
              style={{ height: 72, maxWidth: 240, objectFit: "contain", display: "block", marginBottom: 6 }}
            />
            <div
              style={{
                fontSize: 13,
                color: T.muted,
                letterSpacing: "3px",
                textTransform: "uppercase",
                fontWeight: 600,
              }}
            >
              Field Pro
            </div>
          </div>
        </div>
      </div>

      <div style={{ ...cardS, maxWidth: 400, margin: "0 auto", width: "100%" }}>
        <ErrBanner msg={err} onDismiss={() => setErr("")} />
        <div style={{ marginBottom: 14 }}>
          <label style={lbl}>Select Your Name</label>
          <select value={name} onChange={(e) => handleNameChange(e.target.value)} style={{ ...inp, color: T.orange }}>
            <option value="" style={{ background: "#6B7280", color: T.orange }}>
              — Select your name —
            </option>
            {NAMES.map((n) => (
              <option key={n} style={{ background: "#6B7280", color: T.orange }}>
                {n}
              </option>
            ))}
            <option value="Admin" style={{ background: "#6B7280", color: T.orange }}>
              Admin
            </option>
          </select>
        </div>

        {loading && (
          <div style={{ textAlign: "center", padding: "10px 0", color: T.muted, fontSize: 13 }}>
            Looking up profile…
          </div>
        )}

        {profile && !loading && (
          <div
            style={{
              ...cardS,
              marginBottom: 14,
              background: T.surface,
              borderLeft: `3px solid ${roleM?.color || T.border}`,
            }}
          >
            <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
              <div
                style={{
                  width: 36,
                  height: 36,
                  borderRadius: 10,
                  background: (roleM?.color || T.muted) + "20",
                  display: "flex",
                  alignItems: "center",
                  justifyContent: "center",
                  fontSize: 18,
                }}
              >
                {profile.role === "admin"
                  ? "🔴"
                  : profile.role === "pm"
                    ? "🟠"
                    : profile.role === "foreman"
                      ? "🟡"
                      : "🟢"}
              </div>
              <div>
                <div style={{ fontSize: 14, fontWeight: 700, color: T.text }}>{roleM?.label || "Field Crew"}</div>
                <div style={{ fontSize: 12, color: T.muted }}>{profile.division || "All Divisions"}</div>
              </div>
            </div>
          </div>
        )}

        {profile && !loading && (
          <div style={{ marginBottom: 14 }}>
            <label style={lbl}>PIN</label>
            <input
              type="password"
              maxLength={6}
              placeholder={pinIsSet ? "Enter your PIN" : "No PIN set — contact admin"}
              value={pin}
              onChange={(e) => setPin(e.target.value)}
              onKeyDown={(e) => e.key === "Enter" && handleLogin()}
              style={{ ...inp, opacity: pinIsSet ? 1 : 0.5 }}
              disabled={!pinIsSet}
            />
            {!pinIsSet && profile && (
              <div style={{ fontSize: 11, color: T.yellow, marginTop: 4 }}>
                ⚠️ Ask your admin to set your PIN in PM Dashboard → Users
              </div>
            )}
          </div>
        )}

        <button
          onClick={handleLogin}
          style={{ ...primBtn, opacity: name && !loading && (pinIsSet ? pin.length > 0 : false) ? 1 : 0.45 }}
          disabled={!name || loading || !pinIsSet || pin.length === 0}
        >
          {loading ? "Loading…" : "Sign In →"}
        </button>
      </div>
    </div>
  );
}

/* ── DIVISION SELECTION SCREEN ──────────────────────────────── */
function DivisionScreen({
  user,
  projects,
  onSelect,
  onLogout,
  onCrew,
  onDash,
  onTimeCards,
  onEstimating,
  isOnline,
  pendingCount,
  onSync,
}) {
  const divStats = DIVISIONS.map((div) => {
    const divProjects = projects.filter((p) => p.division === div && p.status === "active");
    const totalBilled = divProjects.reduce((s, p) => s + (p._billed || 0), 0);
    const totalReports = divProjects.reduce((s, p) => s + (p._reports || 0), 0);
    return { div, count: divProjects.length, billed: totalBilled, reports: totalReports };
  });

  return (
    <div style={{ background: T.bg, minHeight: "100vh", fontFamily: "inherit", color: T.text }}>
      {/* Header */}
      <div style={{ background: T.surface, borderBottom: `1px solid ${T.border}`, padding: "16px" }}>
        {/* Offline / pending banner */}
        {!isOnline && (
          <div
            style={{
              background: "#7c2d12",
              borderRadius: 10,
              padding: "8px 12px",
              marginBottom: 10,
              display: "flex",
              alignItems: "center",
              gap: 8,
            }}
          >
            <span style={{ fontSize: 16 }}>📡</span>
            <div>
              <div style={{ fontSize: 13, fontWeight: 700, color: "#fed7aa" }}>No Connection</div>
              <div style={{ fontSize: 11, color: "#fdba74" }}>
                {pendingCount > 0
                  ? `${pendingCount} report${pendingCount !== 1 ? "s" : ""} will sync when back online`
                  : "Reports will save locally until reconnected"}
              </div>
            </div>
          </div>
        )}
        {isOnline && pendingCount > 0 && (
          <div
            style={{
              background: T.greenLow,
              border: `1px solid ${T.green}40`,
              borderRadius: 10,
              padding: "8px 12px",
              marginBottom: 10,
              display: "flex",
              justifyContent: "space-between",
              alignItems: "center",
            }}
          >
            <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
              <span style={{ fontSize: 16 }}>⏳</span>
              <div style={{ fontSize: 13, fontWeight: 700, color: T.green }}>{pendingCount} pending — tap to sync</div>
            </div>
            <button
              onClick={onSync}
              style={{
                background: T.green,
                color: "#FFFFFF",
                border: "none",
                borderRadius: 8,
                padding: "5px 12px",
                fontSize: 12,
                fontWeight: 700,
                cursor: "pointer",
                fontFamily: "inherit",
              }}
            >
              Sync Now
            </button>
          </div>
        )}
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 6 }}>
          <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
            <div style={{ display: "flex", flexDirection: "column", gap: 2 }}>
              <img
                src="/aime-report-logo.png"
                alt="AIME"
                style={{ height: 28, maxWidth: 90, objectFit: "contain", display: "block" }}
              />
              <div
                style={{
                  fontSize: 9,
                  color: T.muted,
                  letterSpacing: "2.5px",
                  textTransform: "uppercase",
                  fontWeight: 700,
                  marginTop: 1,
                }}
              >
                Field Pro
              </div>
            </div>
            <div style={{ fontSize: 11, color: T.muted }}>
              {user.role === "admin" ? "🔴" : user.role === "pm" ? "🟠" : user.role === "foreman" ? "🟡" : "🟢"}{" "}
              {user.name} · {ROLE_META[user.role]?.label}{" "}
              {isOnline ? (
                <span style={{ color: T.green, fontSize: 10 }}>● online</span>
              ) : (
                <span style={{ color: "#f97316", fontSize: 10 }}>● offline</span>
              )}
            </div>
          </div>
          <div style={{ display: "flex", gap: 8 }}>
            {can(user, "view_dashboard") && (
              <button
                onClick={onDash}
                style={{
                  background: T.orangeLow,
                  border: `1px solid ${T.orange}40`,
                  borderRadius: 10,
                  padding: "8px 12px",
                  color: T.orange,
                  fontSize: 12,
                  fontWeight: 700,
                  cursor: "pointer",
                  fontFamily: "inherit",
                }}
              >
                📊
              </button>
            )}
            {(user.role === "admin" || user.role === "pm") && (
              <button
                onClick={onTimeCards}
                style={{
                  background: T.greenLow,
                  border: `1px solid ${T.green}40`,
                  borderRadius: 10,
                  padding: "8px 12px",
                  color: T.green,
                  fontSize: 12,
                  fontWeight: 700,
                  cursor: "pointer",
                  fontFamily: "inherit",
                }}
              >
                ⏱️
              </button>
            )}
            {can(user, "crew_directory") && (
              <button
                onClick={onCrew}
                style={{
                  background: T.blueLow,
                  border: `1px solid ${T.blue}40`,
                  borderRadius: 10,
                  padding: "8px 12px",
                  color: T.blue,
                  fontSize: 12,
                  fontWeight: 700,
                  cursor: "pointer",
                  fontFamily: "inherit",
                }}
              >
                👥
              </button>
            )}
            {can(user, "estimating") && (
              <button
                onClick={onEstimating}
                style={{
                  background: `${T.purple}15`,
                  border: `1px solid ${T.purple}40`,
                  borderRadius: 10,
                  padding: "8px 12px",
                  color: T.purple,
                  fontSize: 12,
                  fontWeight: 700,
                  cursor: "pointer",
                  fontFamily: "inherit",
                }}
              >
                📊
              </button>
            )}
            <button onClick={onLogout} style={{ ...ghostBtn, padding: "8px 12px", fontSize: 12 }}>
              Out
            </button>
          </div>
        </div>
      </div>

      {/* Division picker */}
      <div style={{ padding: "20px 16px 80px" }}>
        <div style={{ marginBottom: 24 }}>
          <div style={{ fontSize: 22, fontWeight: 900, color: T.text, letterSpacing: "-0.5px", marginBottom: 4 }}>
            Select Division
          </div>
          <div style={{ fontSize: 13, color: T.muted }}>Choose the division you are working in today</div>
        </div>

        {DIVISIONS.map((div, i) => {
          const meta = DIV_META[div];
          const stats = divStats.find((s) => s.div === div);
          const divColor = meta.color;
          return (
            <div
              key={div}
              onClick={() => onSelect(div)}
              style={{
                background: T.card,
                border: `1px solid ${T.border}`,
                borderRadius: 20,
                marginBottom: 14,
                cursor: "pointer",
                overflow: "hidden",
                transition: "transform 0.1s",
              }}
            >
              {/* Top gradient bar */}
              <div style={{ height: 4, background: `linear-gradient(90deg,${divColor},${divColor}88)` }} />
              <div style={{ padding: "20px 20px 16px" }}>
                <div style={{ display: "flex", alignItems: "center", gap: 14, marginBottom: 14 }}>
                  <div
                    style={{
                      width: 56,
                      height: 56,
                      borderRadius: 16,
                      background: divColor + "20",
                      border: `2px solid ${divColor}40`,
                      display: "flex",
                      alignItems: "center",
                      justifyContent: "center",
                      fontSize: 28,
                      flexShrink: 0,
                    }}
                  >
                    {meta.icon}
                  </div>
                  <div style={{ flex: 1 }}>
                    <div style={{ fontSize: 22, fontWeight: 900, color: T.text, letterSpacing: "-0.5px" }}>{div}</div>
                    <div style={{ fontSize: 13, color: T.sub, marginTop: 2 }}>{meta.desc}</div>
                  </div>
                  <div style={{ fontSize: 22, color: divColor }}>→</div>
                </div>
                {/* Stats */}
                <div
                  style={{
                    display: "grid",
                    gridTemplateColumns: "1fr 1fr 1fr",
                    gap: 8,
                    paddingTop: 12,
                    borderTop: `1px solid ${T.border}`,
                  }}
                >
                  {[
                    ["Active Jobs", stats?.count || 0, divColor],
                    ["Reports", stats?.reports || 0, T.green],
                    [
                      "Billed",
                      "$" + (stats?.billed >= 1000 ? (stats.billed / 1000).toFixed(1) + "k" : fmt(stats?.billed || 0)),
                      T.blue,
                    ],
                  ].map(([l, v, c]) => (
                    <div key={l} style={{ textAlign: "center" }}>
                      <div style={{ fontSize: 15, fontWeight: 800, color: c }}>{v}</div>
                      <div
                        style={{
                          fontSize: 9,
                          color: T.muted,
                          textTransform: "uppercase",
                          letterSpacing: "0.6px",
                          marginTop: 2,
                        }}
                      >
                        {l}
                      </div>
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
function JobBoard({ user, division, projects, loading, onSelect, onNew, onBack }) {
  const [search, setSearch] = useState("");
  const [filter, setFilter] = useState("active");
  const meta = DIV_META[division] || { icon: "🏗️", color: T.orange };

  const divProjects = projects.filter((p) => p.division === division);
  const filtered = divProjects.filter((p) => {
    const ms = filter === "all" ? true : p.status === filter;
    const q = search.toLowerCase();
    const ms2 =
      !q ||
      p.name?.toLowerCase().includes(q) ||
      p.location?.toLowerCase().includes(q) ||
      p.afe?.toLowerCase().includes(q) ||
      p.client?.toLowerCase().includes(q);
    return ms && ms2;
  });
  const active = divProjects.filter((p) => p.status === "active");
  const canCreate = user.role === "admin" || user.role === "pm" || can(user, "create_job");

  return (
    <div style={{ background: T.bg, minHeight: "100vh", fontFamily: "inherit", color: T.text }}>
      <div style={{ background: T.surface, borderBottom: `1px solid ${T.border}`, padding: "14px 16px 0" }}>
        <button
          onClick={onBack}
          style={{
            background: "none",
            border: "none",
            color: T.sub,
            fontSize: 13,
            cursor: "pointer",
            marginBottom: 10,
            padding: 0,
            fontFamily: "inherit",
          }}
        >
          ← Divisions
        </button>
        <div style={{ display: "flex", alignItems: "center", gap: 12, marginBottom: 14 }}>
          <div
            style={{
              width: 44,
              height: 44,
              borderRadius: 14,
              background: meta.color + "20",
              border: `2px solid ${meta.color}40`,
              display: "flex",
              alignItems: "center",
              justifyContent: "center",
              fontSize: 22,
              flexShrink: 0,
            }}
          >
            {meta.icon}
          </div>
          <div style={{ flex: 1 }}>
            <div style={{ fontSize: 20, fontWeight: 900, color: T.text, letterSpacing: "-0.5px" }}>{division}</div>
            <div style={{ fontSize: 11, color: T.muted }}>
              {active.length} active job{active.length !== 1 ? "s" : ""}
            </div>
          </div>
          {canCreate && (
            <button
              onClick={onNew}
              style={{
                background: T.orange,
                color: "#FFFFFF",
                border: "none",
                borderRadius: 12,
                padding: "10px 16px",
                fontSize: 13,
                fontWeight: 800,
                cursor: "pointer",
                fontFamily: "inherit",
              }}
            >
              + New Job
            </button>
          )}
        </div>
        {/* Search */}
        <div style={{ position: "relative", marginBottom: 10 }}>
          <span
            style={{
              position: "absolute",
              left: 13,
              top: "50%",
              transform: "translateY(-50%)",
              fontSize: 14,
              pointerEvents: "none",
            }}
          >
            🔍
          </span>
          <input
            type="text"
            placeholder="Search jobs…"
            value={search}
            onChange={(e) => setSearch(e.target.value)}
            style={{ ...inp, paddingLeft: 38, borderRadius: 12, fontSize: 14 }}
          />
          {search && (
            <button
              onClick={() => setSearch("")}
              style={{
                position: "absolute",
                right: 12,
                top: "50%",
                transform: "translateY(-50%)",
                background: "none",
                border: "none",
                color: T.muted,
                cursor: "pointer",
                fontSize: 18,
                padding: 0,
              }}
            >
              ×
            </button>
          )}
        </div>
        {/* Tabs */}
        <div style={{ display: "flex", gap: 6 }}>
          {[
            ["active", "Active"],
            ["archived", "Archived"],
            ["all", "All"],
          ].map(([v, l]) => (
            <button
              key={v}
              onClick={() => setFilter(v)}
              style={{
                padding: "8px 14px",
                borderRadius: "10px 10px 0 0",
                background: filter === v ? T.bg : "transparent",
                border: filter === v ? `1px solid ${T.border}` : "1px solid transparent",
                borderBottom: filter === v ? `1px solid ${T.bg}` : "none",
                color: filter === v ? T.text : T.muted,
                fontSize: 13,
                fontWeight: filter === v ? 700 : 500,
                cursor: "pointer",
                fontFamily: "inherit",
                position: "relative",
                zIndex: filter === v ? 1 : 0,
                marginBottom: filter === v ? -1 : 0,
              }}
            >
              {l}
              {v === "active" && active.length > 0 && (
                <span
                  style={{
                    marginLeft: 5,
                    background: meta.color + "25",
                    color: meta.color,
                    borderRadius: 20,
                    padding: "1px 6px",
                    fontSize: 10,
                    fontWeight: 800,
                  }}
                >
                  {active.length}
                </span>
              )}
            </button>
          ))}
        </div>
      </div>

      <div style={{ padding: "12px 16px 80px" }}>
        {loading && <Spinner />}
        {canCreate && (
          <div style={{ position: "fixed", bottom: 20, right: "max(16px,calc(50vw - 224px))", zIndex: 100 }}>
            <button
              onClick={onNew}
              style={{
                background: T.orange,
                color: "#FFFFFF",
                border: "none",
                borderRadius: 50,
                padding: "14px 22px",
                fontSize: 15,
                fontWeight: 900,
                cursor: "pointer",
                fontFamily: "inherit",
                boxShadow: "0 4px 24px rgba(249,115,22,0.5)",
                display: "flex",
                alignItems: "center",
                gap: 8,
              }}
            >
              ＋ New Job
            </button>
          </div>
        )}
        {!loading && filtered.length === 0 && (
          <div style={{ textAlign: "center", padding: "60px 20px", color: T.muted }}>
            <div style={{ fontSize: 48, marginBottom: 12 }}>{meta.icon}</div>
            <div style={{ fontSize: 17, fontWeight: 700, color: T.sub, marginBottom: 6 }}>
              {search
                ? `No jobs matching "${search}"`
                : filter === "archived"
                  ? "No archived jobs"
                  : "No active jobs in " + division}
            </div>
            {!search && filter === "active" && canCreate && (
              <div style={{ fontSize: 13 }}>Tap + New Job to create one.</div>
            )}
          </div>
        )}
        {!loading && filtered.map((p) => <JobCard key={p.id} p={p} onSelect={onSelect} divColor={meta.color} />)}
      </div>
    </div>
  );
}

function JobCard({ p, onSelect, divColor }) {
  const isArchived = p.status !== "active";
  const daysSince = p._lastReport
    ? Math.floor((Date.now() - new Date(p._lastReport + "T12:00:00").getTime()) / 86400000)
    : null;
  const actColor = daysSince === null ? T.muted : daysSince === 0 ? T.green : daysSince <= 2 ? T.orange : T.red;
  const actLabel =
    daysSince === null
      ? "No reports yet"
      : daysSince === 0
        ? "Reported today"
        : daysSince === 1
          ? "Reported yesterday"
          : `${daysSince}d since last report`;
  const c = divColor || T.orange;
  return (
    <div
      onClick={() => onSelect(p)}
      style={{
        background: T.card,
        border: `1px solid ${T.border}`,
        borderRadius: 16,
        marginBottom: 10,
        cursor: "pointer",
        overflow: "hidden",
        opacity: isArchived ? 0.55 : 1,
      }}
    >
      <div style={{ height: 3, background: isArchived ? T.border : `linear-gradient(90deg,${c},${c}88)` }} />
      <div style={{ padding: "14px 16px" }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 8 }}>
          <div style={{ flex: 1, minWidth: 0, paddingRight: 12 }}>
            <div style={{ fontSize: 17, fontWeight: 900, color: T.text, letterSpacing: "-0.3px", lineHeight: 1.2 }}>
              {p.name}
            </div>
            <div style={{ fontSize: 12, color: T.sub, marginTop: 3 }}>
              {[p.client, p.location].filter(Boolean).join(" · ") || "No details"}
            </div>
          </div>
          <div style={{ textAlign: "right", flexShrink: 0 }}>
            <div style={{ fontSize: 20, fontWeight: 900, color: T.green, letterSpacing: "-0.5px" }}>
              ${(p._billed || 0) >= 1000 ? ((p._billed || 0) / 1000).toFixed(1) + "k" : fmt(p._billed || 0)}
            </div>
            <div
              style={{ fontSize: 9, color: T.muted, textTransform: "uppercase", letterSpacing: "0.5px", marginTop: 1 }}
            >
              Total Billed
            </div>
          </div>
        </div>
        <div style={{ display: "flex", gap: 6, flexWrap: "wrap", marginBottom: 10 }}>
          {p.afe && <span style={pill(T.muted)}>AFE: {p.afe}</span>}
          {p.work_order && <span style={pill(T.muted)}>PO: {p.work_order}</span>}
          <span style={pill(p.job_type === "Contract" ? T.blue : T.orange)}>{p.job_type || "T&M"}</span>
          <span style={pill(isArchived ? T.muted : T.green)}>{isArchived ? "Archived" : "Active"}</span>
        </div>
        <div
          style={{
            display: "flex",
            justifyContent: "space-between",
            alignItems: "center",
            paddingTop: 10,
            borderTop: `1px solid ${T.border}`,
          }}
        >
          <div style={{ display: "flex", gap: 16 }}>
            <div>
              <div style={{ fontSize: 16, fontWeight: 800, color: c }}>{p._reports || 0}</div>
              <div style={{ fontSize: 9, color: T.muted, textTransform: "uppercase", letterSpacing: "0.5px" }}>
                Reports
              </div>
            </div>
            <div>
              <div style={{ fontSize: 11, fontWeight: 700, color: actColor, marginTop: 2 }}>{actLabel}</div>
              <div
                style={{
                  fontSize: 9,
                  color: T.muted,
                  textTransform: "uppercase",
                  letterSpacing: "0.5px",
                  marginTop: 1,
                }}
              >
                Last Activity
              </div>
            </div>
          </div>
          <div
            style={{
              display: "flex",
              alignItems: "center",
              gap: 6,
              background: c + "15",
              border: `1px solid ${c}40`,
              borderRadius: 10,
              padding: "8px 14px",
            }}
          >
            <span style={{ fontSize: 13, fontWeight: 700, color: c }}>Enter Job</span>
            <span style={{ fontSize: 16, color: c }}>→</span>
          </div>
        </div>
      </div>
    </div>
  );
}

/* ── PROJECT FORM ───────────────────────────────────────────── */
function ProjectForm({ initial, onSave, onCancel, saving, defaultDivision, externalErr, onClearErr }) {
  const [f, setF] = useState(
    initial || {
      name: "",
      client: "",
      location: "",
      afe: "",
      work_order: "",
      start_date: today(),
      notes: "",
      status: "active",
      division: defaultDivision || "Pipeline",
      job_type: "T&M",
      contract_value: "",
      contract_hours: "",
      estimated_budget: "",
    },
  );
  const set = (k, v) => setF((x) => ({ ...x, [k]: v }));
  return (
    <div style={{ background: T.bg, minHeight: "100vh", fontFamily: "inherit", color: T.text }}>
      <TopBar title={initial ? "Edit Job" : "New Job"} onBack={onCancel} />
      <div style={{ padding: "16px 16px 100px" }}>
        <ErrBanner msg={externalErr} onDismiss={onClearErr} />
        {[
          { k: "name", l: "Job Number *", ph: "e.g. HDD-2026-001" },
          { k: "client", l: "Client", ph: "Colonial Pipeline" },
          { k: "location", l: "Location", ph: "City, State or Milepost" },
          { k: "afe", l: "AFE No.", ph: "AFE #" },
          { k: "work_order", l: "PO #", ph: "PO #" },
        ].map(({ k, l, ph }) => (
          <div key={k} style={{ marginBottom: 12 }}>
            <label style={lbl}>{l}</label>
            <input
              type="text"
              placeholder={ph}
              value={f[k] || ""}
              onChange={(e) => set(k, e.target.value)}
              style={inp}
            />
          </div>
        ))}

        {/* Job Type */}
        <div style={{ marginBottom: 12 }}>
          <label style={lbl}>Job Type</label>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8 }}>
            {[
              ["T&M", "⏱️ Time & Material"],
              ["Contract", "📋 Contract"],
            ].map(([val, label]) => (
              <button
                key={val}
                onClick={() => set("job_type", val)}
                style={{
                  padding: "14px 10px",
                  borderRadius: 12,
                  border: `2px solid ${f.job_type === val ? T.orange : T.border}`,
                  background: f.job_type === val ? T.orangeLow : T.surface,
                  color: f.job_type === val ? T.orange : T.sub,
                  fontWeight: 700,
                  fontSize: 14,
                  cursor: "pointer",
                  fontFamily: "inherit",
                  textAlign: "center",
                }}
              >
                {label}
              </button>
            ))}
          </div>
        </div>

        {/* Contract value — only shown for Contract jobs */}
        {f.job_type === "Contract" && (
          <>
            <div style={{ marginBottom: 12 }}>
              <label style={lbl}>Contract Total Value ($)</label>
              <input
                type="number"
                min="0"
                step="0.01"
                placeholder="e.g. 250000.00"
                value={f.contract_value || ""}
                onChange={(e) => set("contract_value", e.target.value)}
                style={inp}
              />
            </div>
            <div style={{ marginBottom: 12 }}>
              <label style={lbl}>Contract Total Hours</label>
              <input
                type="number"
                min="0"
                step="0.5"
                placeholder="e.g. 2000 — total hours budgeted for this contract"
                value={f.contract_hours || ""}
                onChange={(e) => set("contract_hours", e.target.value)}
                style={inp}
              />
              <div style={{ fontSize: 11, color: T.muted, marginTop: 4 }}>
                Track hours used vs. contract hours — shows a progress bar on the job.
              </div>
            </div>
          </>
        )}
        {/* Estimated budget for T&M jobs */}
        {f.job_type === "T&M" && (
          <div style={{ marginBottom: 12 }}>
            <label style={lbl}>Estimated Budget (optional, $)</label>
            <input
              type="number"
              min="0"
              step="0.01"
              placeholder="e.g. 50000.00 — leave blank if open-ended"
              value={f.estimated_budget || ""}
              onChange={(e) => set("estimated_budget", e.target.value)}
              style={inp}
            />
            <div style={{ fontSize: 11, color: T.muted, marginTop: 4 }}>
              Used to show budget vs. actual billing progress on the job.
            </div>
          </div>
        )}

        <div style={{ marginBottom: 12 }}>
          <label style={lbl}>Division</label>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 8 }}>
            {DIVISIONS.map((div) => {
              const m = DIV_META[div];
              return (
                <button
                  key={div}
                  onClick={() => set("division", div)}
                  style={{
                    padding: "12px 8px",
                    borderRadius: 12,
                    border: `2px solid ${f.division === div ? m.color : T.border}`,
                    background: f.division === div ? m.color + "20" : T.surface,
                    color: f.division === div ? m.color : T.sub,
                    fontWeight: 700,
                    fontSize: 12,
                    cursor: "pointer",
                    fontFamily: "inherit",
                    textAlign: "center",
                  }}
                >
                  <div style={{ fontSize: 20, marginBottom: 4 }}>{m.icon}</div>
                  {div}
                </button>
              );
            })}
          </div>
        </div>
        <div style={{ marginBottom: 12 }}>
          <label style={lbl}>Start Date</label>
          <input
            type="date"
            value={f.start_date || today()}
            onChange={(e) => set("start_date", e.target.value)}
            style={inp}
          />
        </div>
        <div style={{ marginBottom: 20 }}>
          <label style={lbl}>Notes</label>
          <textarea
            placeholder="Project notes, scope, special instructions…"
            value={f.notes || ""}
            onChange={(e) => set("notes", e.target.value)}
            rows={3}
            style={{ ...inp, resize: "vertical", lineHeight: 1.5 }}
          />
        </div>
        <button
          onClick={() =>
            f.name.trim() &&
            !saving &&
            onSave({
              ...f,
              contract_value: f.contract_value && f.contract_value !== "" ? parseFloat(f.contract_value) : null,
              contract_hours: f.contract_hours && f.contract_hours !== "" ? parseFloat(f.contract_hours) : null,
              estimated_budget: f.estimated_budget && f.estimated_budget !== "" ? parseFloat(f.estimated_budget) : null,
            })
          }
          style={{ ...primBtn, opacity: f.name.trim() && !saving ? 1 : 0.5 }}
        >
          {saving ? "Saving…" : initial ? "Save Changes" : "Create Job"}
        </button>
      </div>
    </div>
  );
}

/* ── DAILY REPORT FORM ──────────────────────────────────────── */
const RSTEPS = ["Job Info", "Labor", "Equipment", "Materials", "Site Notes", "Review"];

/* ── AUTO-POPULATE TIME CARDS FROM DAILY REPORT ─────────────── */
async function autoPopulateTimeCards(report, project) {
  const labor = (report.labor || []).filter((l) => l.name && l.name.trim());
  if (!labor.length) return { created: 0, updated: 0 };
  let created = 0,
    updated = 0;
  for (const entry of labor) {
    const reg = parseFloat(entry.regHrs) || 0;
    const ot = parseFloat(entry.otHrs) || 0;
    const travel = parseFloat(entry.travelHrs) || 0;
    if (reg + ot + travel === 0) continue;
    try {
      // Check for existing time card for this worker+date+job
      const existing = await API.timeCards.find(entry.name, report.date, project.id);
      const card = Array.isArray(existing) ? existing[0] : null;
      if (card) {
        // Add hours to existing card (worker may have multiple report entries)
        const newReg = (parseFloat(card.reg_hours) || 0) + reg;
        const newOT = (parseFloat(card.ot_hours) || 0) + ot;
        const newTravel = (parseFloat(card.travel_hours) || 0) + travel;
        await API.timeCards.update(card.id, {
          reg_hours: newReg,
          ot_hours: newOT,
          travel_hours: newTravel,
          total_hours: newReg + newOT + newTravel,
        });
        updated++;
      } else {
        // Create new time card
        await API.timeCards.create({
          worker_name: entry.name,
          date: report.date,
          project_id: project.id,
          division: project.division,
          classification: entry.classification || "",
          reg_hours: reg,
          ot_hours: ot,
          travel_hours: travel,
          total_hours: reg + ot + travel,
          notes: `Auto-filled from daily report${report.report_no ? " #" + report.report_no : ""} · ${project.name}`,
        });
        created++;
      }
    } catch (e) {}
  }
  return { created, updated };
}

/* ── VISITOR ADD ROW ─────────────────────────────────────────── */
function VisitorAddRow({ onAdd }) {
  const [show, setShow] = useState(false);
  const [f, setF] = useState({ name: "", company: "", type: "Inspector", notes: "" });
  if (!show)
    return (
      <button
        onClick={() => setShow(true)}
        style={{ ...ghostBtn, width: "100%", textAlign: "center", fontSize: 13, marginTop: 4 }}
      >
        + Add Visitor
      </button>
    );
  return (
    <div style={{ background: "#FFFFFF", borderRadius: 10, padding: 12, marginTop: 8 }}>
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8, marginBottom: 8 }}>
        <div>
          <label style={lbl}>Name *</label>
          <input
            value={f.name}
            onChange={(e) => setF((x) => ({ ...x, name: e.target.value }))}
            placeholder="Full name"
            style={inp}
          />
        </div>
        <div>
          <label style={lbl}>Company</label>
          <input
            value={f.company}
            onChange={(e) => setF((x) => ({ ...x, company: e.target.value }))}
            placeholder="Company"
            style={inp}
          />
        </div>
      </div>
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8, marginBottom: 8 }}>
        <div>
          <label style={lbl}>Type</label>
          <select value={f.type} onChange={(e) => setF((x) => ({ ...x, type: e.target.value }))} style={inpSel}>
            {["Inspector", "Client", "Engineer", "Vendor", "Safety Officer", "Owner Rep", "Subcontractor", "Other"].map(
              (t) => (
                <option key={t}>{t}</option>
              ),
            )}
          </select>
        </div>
        <div>
          <label style={lbl}>Notes</label>
          <input
            value={f.notes}
            onChange={(e) => setF((x) => ({ ...x, notes: e.target.value }))}
            placeholder="Purpose of visit"
            style={inp}
          />
        </div>
      </div>
      <div style={{ display: "flex", gap: 8 }}>
        <button
          onClick={() => {
            if (!f.name.trim()) return;
            onAdd({ ...f });
            setF({ name: "", company: "", type: "Inspector", notes: "" });
            setShow(false);
          }}
          disabled={!f.name.trim()}
          style={{ ...primBtn, flex: 2, borderRadius: 10, fontSize: 13, opacity: f.name.trim() ? 1 : 0.5 }}
        >
          Add Visitor
        </button>
        <button onClick={() => setShow(false)} style={{ ...ghostBtn, flex: 1, textAlign: "center", fontSize: 13 }}>
          Cancel
        </button>
      </div>
    </div>
  );
}

/* ── DELAY ADD ROW ───────────────────────────────────────────── */
function DelayAddRow({ onAdd }) {
  const [show, setShow] = useState(false);
  const CAUSES = [
    "Weather",
    "Material Delay",
    "Equipment Breakdown",
    "Owner/Client",
    "Subcontractor",
    "Labor",
    "Design/Engineering",
    "Permitting",
    "Safety Stop",
    "Other",
  ];
  const [f, setF] = useState({ cause: "Weather", description: "", hours: 0, impact: "" });
  if (!show)
    return (
      <button
        onClick={() => setShow(true)}
        style={{
          ...ghostBtn,
          width: "100%",
          textAlign: "center",
          fontSize: 13,
          marginTop: 4,
          color: T.red,
          border: `1px solid ${T.red}30`,
        }}
      >
        + Log Delay / Issue
      </button>
    );
  return (
    <div style={{ background: "#FFFFFF", borderRadius: 10, padding: 12, marginTop: 8, border: `1px solid ${T.red}30` }}>
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8, marginBottom: 8 }}>
        <div>
          <label style={lbl}>Cause *</label>
          <select
            value={f.cause}
            onChange={(e) => setF((x) => ({ ...x, cause: e.target.value }))}
            style={{ ...inpSel, color: T.red }}
          >
            {CAUSES.map((c) => (
              <option key={c}>{c}</option>
            ))}
          </select>
        </div>
        <div>
          <label style={lbl}>Hours Lost</label>
          <input
            type="number"
            value={f.hours}
            onChange={(e) => setF((x) => ({ ...x, hours: e.target.value }))}
            min={0}
            step={0.5}
            style={inp}
          />
        </div>
      </div>
      <div style={{ marginBottom: 8 }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 6 }}>
          <label style={lbl}>Description *</label>
        </div>
        <input
          value={f.description}
          onChange={(e) => setF((x) => ({ ...x, description: e.target.value }))}
          placeholder="What happened?"
          style={inp}
        />
      </div>
      <div style={{ marginBottom: 10 }}>
        <label style={lbl}>Impact / Action Taken</label>
        <input
          value={f.impact}
          onChange={(e) => setF((x) => ({ ...x, impact: e.target.value }))}
          placeholder="How did it affect work?"
          style={inp}
        />
      </div>
      <div style={{ display: "flex", gap: 8 }}>
        <button
          onClick={() => {
            if (!f.description.trim()) return;
            onAdd({ ...f, hours: parseFloat(f.hours) || 0 });
            setF({ cause: "Weather", description: "", hours: 0, impact: "" });
            setShow(false);
          }}
          disabled={!f.description.trim()}
          style={{
            ...primBtn,
            flex: 2,
            borderRadius: 10,
            fontSize: 13,
            background: T.red,
            opacity: f.description.trim() ? 1 : 0.5,
          }}
        >
          Log Delay
        </button>
        <button onClick={() => setShow(false)} style={{ ...ghostBtn, flex: 1, textAlign: "center", fontSize: 13 }}>
          Cancel
        </button>
      </div>
    </div>
  );
}

function DailyReportForm({ user, project, onSave, onCancel, isOnline }) {
  const draftKey = `${project.id}_${user.name}`;
  const existingDraft = loadDraft(draftKey);

  const [step, setStep] = useState(1);
  const [saving, setSaving] = useState(false);
  const [draftSaved, setDraftSaved] = useState(false);
  const [showDraftBanner, setShowDraftBanner] = useState(!!existingDraft);
  const [rpt, setRpt] = useState(
    existingDraft?.data || {
      date: today(),
      description: "",
      report_no: "",
      labor: [],
      equipment: [],
      rental_equipment: [],
      materials: [],
      visitor_log: [],
      delays: [],
      site_conditions: "",
    },
  );
  const topRef = useRef(null);
  const setR = (k, v) => setRpt((r) => ({ ...r, [k]: v }));
  function add(key, item) {
    setR(key, [...rpt[key], item]);
  }
  function upd(key, i, row) {
    const a = [...rpt[key]];
    a[i] = row;
    setR(key, a);
  }
  function del(key, i) {
    setR(
      key,
      rpt[key].filter((_, j) => j !== i),
    );
  }

  // Voice-to-text
  const [listening, setListening] = useState(false);
  const [listenTarget, setListenTarget] = useState(null);
  function startVoice(key) {
    const SR = window.SpeechRecognition || window.webkitSpeechRecognition;
    if (!SR) {
      alert("Voice-to-text not supported in this browser. Try Chrome.");
      return;
    }
    const r = new SR();
    r.lang = "en-US";
    r.continuous = false;
    r.interimResults = false;
    setListening(true);
    setListenTarget(key);
    r.onresult = (e) => {
      const txt = e.results[0][0].transcript;
      setR(key, (rpt[key] || "") + (rpt[key] ? " " : "") + txt);
      setListening(false);
      setListenTarget(null);
    };
    r.onerror = () => {
      setListening(false);
      setListenTarget(null);
    };
    r.onend = () => {
      setListening(false);
      setListenTarget(null);
    };
    r.start();
  }

  // Weather auto-fill
  const [weatherFilling, setWeatherFilling] = useState(false);
  async function autoFillWeather() {
    const location = project.location || project.client || project.name;
    if (!location) {
      alert("No location set on this project.");
      return;
    }
    setWeatherFilling(true);
    try {
      const w = await fetchWeather(location);
      setR(
        "site_conditions",
        `${w.conditionText || ""} · ${w.tempF || ""}°F · Wind: ${w.windMph || "0"}mph · Humidity: ${w.humidity || ""}%`,
      );
    } catch (e) {
      alert("Could not fetch weather: " + e.message);
    }
    setWeatherFilling(false);
  }
  const tot = reportTotals(rpt, project.division);

  // Auto-save draft every time form changes
  useEffect(() => {
    const t = setTimeout(() => {
      saveDraft(draftKey, rpt);
      setDraftSaved(true);
      setTimeout(() => setDraftSaved(false), 2000);
    }, 800);
    return () => clearTimeout(t);
  }, [rpt]);

  async function submit() {
    setSaving(true);
    // Strip fields not in DB schema before saving
    const { rental_equipment, ...rptClean } = rpt;
    const reportData = {
      ...rptClean,
      submitted_by: user.name,
      status: "submitted",
      project_id: project.id,
      rental_equipment: rental_equipment || [],
    };

    // ── Duplicate check ──
    if (isOnline) {
      try {
        const existing = await API.reports.forProject(project.id);
        const dupe = (existing || []).find((r) => r.date === rpt.date);
        if (dupe) {
          const proceed = window.confirm(
            `⚠️ A report for ${fmtDate(rpt.date)} already exists on this job (submitted by ${dupe.submitted_by || "someone"}).\n\nDo you still want to submit a second report for this date?`,
          );
          if (!proceed) {
            setSaving(false);
            return;
          }
        }
      } catch {
        /* ignore — if check fails just proceed */
      }
    }

    if (!isOnline) {
      addToQueue({ type: "report", data: reportData });
      clearDraft(draftKey);
      setSaving(false);
      alert("No connection — report saved and will sync automatically when you're back online.");
      onCancel();
      return;
    }
    try {
      await onSave(reportData);
      clearDraft(draftKey);
      // Auto-populate time cards for each worker on this report
      try {
        const tcResult = await autoPopulateTimeCards(reportData, project);
      } catch (e) {}
      await notify("report_submitted", "New Report Submitted", `${user.name} submitted a report for ${project.name}`, {
        project_id: project.id,
      });
    } catch (e) {
      addToQueue({ type: "report", data: reportData });
      clearDraft(draftKey);
      alert("Couldn't reach server — report queued and will sync when reconnected.");
      onCancel();
    }
    setSaving(false);
  }
  const scroll = () => topRef.current?.scrollIntoView({ behavior: "smooth" });
  const divMeta = DIV_META[project.division] || { color: T.orange };
  return (
    <div ref={topRef} style={{ background: T.bg, minHeight: "100vh", fontFamily: "inherit" }}>
      <div
        style={{
          background: T.surface,
          borderBottom: `1px solid ${T.border}`,
          padding: "14px 16px",
          position: "sticky",
          top: 0,
          zIndex: 50,
        }}
      >
        {/* Offline indicator */}
        {!isOnline && (
          <div
            style={{
              background: "#7c2d12",
              borderRadius: 8,
              padding: "6px 10px",
              marginBottom: 8,
              fontSize: 12,
              color: "#fed7aa",
              display: "flex",
              alignItems: "center",
              gap: 6,
            }}
          >
            <span>📡</span>Offline — report will save locally and sync when reconnected
          </div>
        )}
        {/* Draft restore banner */}
        {showDraftBanner && (
          <div
            style={{
              background: T.blueLow,
              border: `1px solid ${T.blue}40`,
              borderRadius: 8,
              padding: "8px 10px",
              marginBottom: 8,
              display: "flex",
              justifyContent: "space-between",
              alignItems: "center",
            }}
          >
            <div style={{ fontSize: 12, color: T.blue }}>
              <strong>Draft restored</strong> · saved{" "}
              {existingDraft?.saved_at ? new Date(existingDraft.saved_at).toLocaleTimeString() : ""}
            </div>
            <button
              onClick={() => setShowDraftBanner(false)}
              style={{ background: "none", border: "none", color: T.blue, cursor: "pointer", fontSize: 16, padding: 0 }}
            >
              ×
            </button>
          </div>
        )}
        <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 12 }}>
          <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
            <div style={{ fontSize: 16, fontWeight: 800 }}>New Daily Report</div>
            {draftSaved && <span style={{ fontSize: 10, color: T.green, fontWeight: 600 }}>✓ Draft saved</span>}
          </div>
          <button
            onClick={onCancel}
            style={{
              background: "none",
              border: "none",
              color: T.sub,
              cursor: "pointer",
              fontSize: 13,
              fontFamily: "inherit",
            }}
          >
            Cancel
          </button>
        </div>
        <div style={{ display: "flex", alignItems: "center" }}>
          {RSTEPS.map((s, i) => (
            <div key={i} style={{ display: "flex", alignItems: "center", flex: i < RSTEPS.length - 1 ? 1 : undefined }}>
              <div style={{ display: "flex", flexDirection: "column", alignItems: "center", gap: 3 }}>
                <div
                  style={{
                    width: 26,
                    height: 26,
                    borderRadius: "50%",
                    background: i + 1 < step ? T.green : i + 1 === step ? divMeta.color : T.border,
                    color: i + 1 <= step ? "#16181D" : T.muted,
                    display: "flex",
                    alignItems: "center",
                    justifyContent: "center",
                    fontSize: 11,
                    fontWeight: 800,
                  }}
                >
                  {i + 1 < step ? "✓" : i + 1}
                </div>
                <div
                  style={{
                    fontSize: 8,
                    color: i + 1 === step ? divMeta.color : T.muted,
                    fontWeight: i + 1 === step ? 700 : 400,
                    whiteSpace: "nowrap",
                  }}
                >
                  {s}
                </div>
              </div>
              {i < RSTEPS.length - 1 && (
                <div
                  style={{
                    flex: 1,
                    height: 2,
                    background: i + 1 < step ? T.green : T.border,
                    margin: "0 3px",
                    marginBottom: 14,
                  }}
                />
              )}
            </div>
          ))}
        </div>
      </div>
      <div style={{ padding: "14px 16px 100px" }}>
        {step === 1 && (
          <div>
            <div style={{ ...cardS, marginBottom: 14, borderLeft: `3px solid ${divMeta.color}` }}>
              <div style={{ fontSize: 11, color: T.muted, marginBottom: 2 }}>Project · {project.division}</div>
              <div style={{ fontSize: 15, fontWeight: 700 }}>{project.name}</div>
              {project.afe && (
                <div style={{ fontSize: 12, color: T.sub }}>
                  AFE: {project.afe}
                  {project.work_order ? " · PO: " + project.work_order : ""}
                </div>
              )}
              <div style={{ marginTop: 6, display: "flex", gap: 6 }}>
                <span style={pill(divMeta.color)}>
                  {divMeta.icon} {project.division} Rates
                </span>
              </div>
            </div>
            <div style={{ marginBottom: 12 }}>
              <label style={lbl}>Date</label>
              <input type="date" value={rpt.date} onChange={(e) => setR("date", e.target.value)} style={inp} />
            </div>
            <div style={{ marginBottom: 12 }}>
              <label style={lbl}>Report No.</label>
              <input
                type="text"
                placeholder="Report #"
                value={rpt.report_no || ""}
                onChange={(e) => setR("report_no", e.target.value)}
                style={inp}
              />
            </div>
            {/* Site Conditions + Weather */}
            <div style={{ ...cardS, marginBottom: 12, border: `1px solid ${T.blue}30`, padding: "12px 14px" }}>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 8 }}>
                <label style={lbl}>🌤️ Site Conditions / Weather</label>
                <button
                  onClick={autoFillWeather}
                  disabled={weatherFilling}
                  style={{
                    background: T.blueLow,
                    border: `1px solid ${T.blue}40`,
                    borderRadius: 8,
                    padding: "5px 10px",
                    color: T.blue,
                    fontSize: 11,
                    fontWeight: 700,
                    cursor: "pointer",
                    fontFamily: "inherit",
                  }}
                >
                  {weatherFilling ? "Fetching…" : "⚡ Auto-Fill"}
                </button>
              </div>
              <input
                value={rpt.site_conditions || ""}
                onChange={(e) => setR("site_conditions", e.target.value)}
                placeholder="e.g. Clear · 72°F · Wind 5mph · Humidity 45%"
                style={inp}
              />
            </div>

            {/* Description with voice-to-text */}
            <div style={{ marginBottom: 12 }}>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 6 }}>
                <label style={lbl}>Description of Work Done</label>
                <button
                  onClick={() => startVoice("description")}
                  title="Voice to text"
                  style={{
                    background: listening && listenTarget === "description" ? T.redLow : T.card,
                    border: `1px solid ${listening && listenTarget === "description" ? T.red : T.border}`,
                    borderRadius: 8,
                    padding: "5px 10px",
                    color: listening && listenTarget === "description" ? T.red : T.muted,
                    fontSize: 13,
                    cursor: "pointer",
                    fontFamily: "inherit",
                    animation: listening && listenTarget === "description" ? "pulse 1s infinite" : undefined,
                  }}
                >
                  {listening && listenTarget === "description" ? "🔴 Listening…" : "🎤"}
                </button>
              </div>
              <textarea
                placeholder="Describe the work performed today… or tap 🎤 to speak"
                value={rpt.description || ""}
                onChange={(e) => setR("description", e.target.value)}
                rows={4}
                style={{ ...inp, resize: "vertical", lineHeight: 1.5 }}
              />
            </div>
          </div>
        )}
        {step === 2 && (
          <div>
            <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 12 }}>
              <div style={{ fontSize: 17, fontWeight: 800 }}>👷 Labor</div>
              {tot.labor > 0 && can(user, "view_dashboard") && (
                <div style={{ fontSize: 16, fontWeight: 800, color: T.green }}>${fmt(tot.labor)}</div>
              )}
            </div>
            {rpt.labor.map((row, i) => (
              <LaborCard
                key={row.id}
                row={row}
                onChange={(r) => upd("labor", i, r)}
                onRemove={() => del("labor", i)}
                division={project.division}
              />
            ))}
            <DashedAdd
              label="+ Add Worker"
              onClick={() =>
                add("labor", { id: uid(), name: "", classification: "", regHrs: "", otHrs: "", travelHrs: "" })
              }
              color={T.orange}
            />
          </div>
        )}
        {step === 3 && (
          <div>
            {/* Company Equipment */}
            <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 10 }}>
              <div style={{ fontSize: 17, fontWeight: 800 }}>🚜 Equipment</div>
              {tot.equip > 0 && can(user, "view_dashboard") && (
                <div style={{ fontSize: 16, fontWeight: 800, color: T.green }}>${fmt(tot.equip)}</div>
              )}
            </div>
            {rpt.equipment.map((row, i) => (
              <EquipCard
                key={row.id}
                row={row}
                onChange={(r) => upd("equipment", i, r)}
                onRemove={() => del("equipment", i)}
                division={project.division}
              />
            ))}
            <DashedAdd
              label="+ Add Company Equipment"
              onClick={() => add("equipment", { id: uid(), description: "", qty: "", usage: "", rate: "", unit: "" })}
              color={T.yellow}
            />

            {/* Rented Equipment */}
            <div style={{ marginTop: 18, marginBottom: 10, paddingTop: 14, borderTop: `2px dashed ${T.border}` }}>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                <div style={{ fontSize: 17, fontWeight: 800 }}>🔑 Rented Equipment</div>
                {(rpt.rental_equipment || []).length > 0 && can(user, "view_dashboard") && (
                  <div style={{ fontSize: 16, fontWeight: 800, color: T.green }}>
                    $
                    {fmt(
                      (rpt.rental_equipment || []).reduce((s, r) => {
                        const a = (parseFloat(r.qty) || 0) * (parseFloat(r.rate) || 0) * (parseFloat(r.usage) || 1);
                        return s + a;
                      }, 0),
                    )}
                  </div>
                )}
              </div>
              <div style={{ fontSize: 12, color: T.muted, marginTop: 2, marginBottom: 10 }}>
                Equipment you don't own — type any description
              </div>
            </div>
            {(rpt.rental_equipment || []).map((row, i) => (
              <RentedEquipCard
                key={row.id}
                row={row}
                onChange={(r) => upd("rental_equipment", i, r)}
                onRemove={() => del("rental_equipment", i)}
              />
            ))}
            <DashedAdd
              label="+ Add Rented Equipment"
              onClick={() => add("rental_equipment", { id: uid(), description: "", qty: "", usage: "", rate: "" })}
              color={T.purple}
            />
          </div>
        )}
        {step === 4 && (
          <div>
            <div style={{ fontSize: 17, fontWeight: 800, marginBottom: 12 }}>📦 Materials & Misc.</div>
            {rpt.materials.map((row, i) => (
              <MatCard
                key={row.id}
                row={row}
                onChange={(r) => upd("materials", i, r)}
                onRemove={() => del("materials", i)}
              />
            ))}
            <DashedAdd
              label="+ Add Material / Item"
              onClick={() => add("materials", { id: uid(), qty: "", description: "", amount: "", receipts: [] })}
              color={T.blue}
            />
          </div>
        )}
        {step === 5 && (
          <div>
            <div style={{ fontSize: 17, fontWeight: 800, marginBottom: 16 }}>📋 Site Notes</div>

            {/* Manpower Summary — auto-count from labor */}
            <div style={{ ...cardS, marginBottom: 12, background: T.blueLow, border: `1px solid ${T.blue}30` }}>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                <div>
                  <div
                    style={{
                      fontSize: 12,
                      fontWeight: 700,
                      color: T.blue,
                      textTransform: "uppercase",
                      letterSpacing: "1px",
                    }}
                  >
                    👷 Manpower On Site
                  </div>
                  <div style={{ fontSize: 28, fontWeight: 900, color: T.text, marginTop: 2 }}>
                    {rpt.labor?.length || 0}
                  </div>
                  <div style={{ fontSize: 11, color: T.muted }}>workers from labor log</div>
                </div>
                <div style={{ textAlign: "right" }}>
                  <div style={{ fontSize: 11, color: T.muted }}>Total Hrs</div>
                  <div style={{ fontSize: 22, fontWeight: 800, color: T.green }}>
                    {(rpt.labor || [])
                      .reduce(
                        (s, l) =>
                          s + (parseFloat(l.regHrs) || 0) + (parseFloat(l.otHrs) || 0) + (parseFloat(l.travelHrs) || 0),
                        0,
                      )
                      .toFixed(1)}
                  </div>
                </div>
              </div>
            </div>

            {/* Visitor Log */}
            <div style={{ ...cardS, marginBottom: 12 }}>
              <div style={{ fontSize: 14, fontWeight: 800, color: T.text, marginBottom: 12 }}>🏗️ Visitor Log</div>
              {(rpt.visitor_log || []).map((v, i) => (
                <div
                  key={i}
                  style={{
                    background: T.surface,
                    borderRadius: 10,
                    padding: "10px 12px",
                    marginBottom: 8,
                    display: "flex",
                    justifyContent: "space-between",
                    alignItems: "flex-start",
                  }}
                >
                  <div>
                    <div style={{ fontSize: 13, fontWeight: 700, color: T.text }}>{v.name}</div>
                    <div style={{ fontSize: 11, color: T.muted }}>
                      {v.company && v.company + " · "}
                      {v.type}
                    </div>
                    {v.notes && (
                      <div style={{ fontSize: 11, color: T.sub, marginTop: 2, fontStyle: "italic" }}>{v.notes}</div>
                    )}
                  </div>
                  <button
                    onClick={() => del("visitor_log", i)}
                    style={{
                      background: "none",
                      border: "none",
                      color: T.red,
                      cursor: "pointer",
                      fontSize: 16,
                      padding: 0,
                    }}
                  >
                    ✕
                  </button>
                </div>
              ))}
              <VisitorAddRow onAdd={(v) => add("visitor_log", v)} />
            </div>

            {/* Delay / Issue Tracking */}
            <div style={{ ...cardS, marginBottom: 12 }}>
              <div style={{ fontSize: 14, fontWeight: 800, color: T.text, marginBottom: 12 }}>⚠️ Delays & Issues</div>
              {(rpt.delays || []).map((d, i) => (
                <div
                  key={i}
                  style={{
                    background: T.surface,
                    borderRadius: 10,
                    padding: "10px 12px",
                    marginBottom: 8,
                    borderLeft: `3px solid ${T.red}`,
                    display: "flex",
                    justifyContent: "space-between",
                    alignItems: "flex-start",
                  }}
                >
                  <div>
                    <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 2 }}>
                      <span style={{ ...pill(T.red), fontSize: 10 }}>{d.cause}</span>
                      {d.hours > 0 && <span style={{ fontSize: 11, color: T.muted }}>{d.hours}h delay</span>}
                    </div>
                    <div style={{ fontSize: 13, color: T.text }}>{d.description}</div>
                    {d.impact && <div style={{ fontSize: 11, color: T.muted, marginTop: 2 }}>Impact: {d.impact}</div>}
                  </div>
                  <button
                    onClick={() => del("delays", i)}
                    style={{
                      background: "none",
                      border: "none",
                      color: T.red,
                      cursor: "pointer",
                      fontSize: 16,
                      padding: 0,
                    }}
                  >
                    ✕
                  </button>
                </div>
              ))}
              <DelayAddRow onAdd={(d) => add("delays", d)} />
            </div>
          </div>
        )}

        {step === 6 && (
          <div>
            <div style={{ fontSize: 17, fontWeight: 800, marginBottom: 12 }}>✅ Review & Submit</div>
            <div style={{ ...cardS, marginBottom: 12 }}>
              <div
                style={{
                  fontSize: 11,
                  color: divMeta.color,
                  fontWeight: 700,
                  textTransform: "uppercase",
                  letterSpacing: "1px",
                  marginBottom: 10,
                }}
              >
                Summary
              </div>
              {[
                ["Project", project.name],
                ["Division", project.division],
                ["Date", fmtDate(rpt.date)],
                ["Report No.", rpt.report_no || "—"],
                ["Workers", rpt.labor.length],
                ["Equipment", rpt.equipment.length + " items"],
                ["Rented Equip", (rpt.rental_equipment || []).length + " items"],
                ["Materials", rpt.materials.length + " items"],
              ].map(([l, v]) => (
                <div
                  key={l}
                  style={{
                    display: "flex",
                    justifyContent: "space-between",
                    padding: "7px 0",
                    borderBottom: `1px solid ${T.border}`,
                  }}
                >
                  <span style={{ fontSize: 13, color: T.sub }}>{l}</span>
                  <span style={{ fontSize: 13, fontWeight: 600 }}>{v}</span>
                </div>
              ))}
              {[
                ["Labor", tot.labor, T.green],
                ["Equipment", tot.equip, T.green],
                ["Materials", tot.mats, T.green],
              ].map(([l, v, c]) => (
                <div
                  key={l}
                  style={{
                    display: "flex",
                    justifyContent: "space-between",
                    padding: "7px 0",
                    borderBottom: `1px solid ${T.border}`,
                  }}
                >
                  <span style={{ fontSize: 13, color: T.sub }}>{l}</span>
                  <span style={{ fontSize: 13, fontWeight: 700, color: c }}>${fmt(v)}</span>
                </div>
              ))}
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", paddingTop: 14 }}>
                <span style={{ fontSize: 16, fontWeight: 800 }}>Grand Total</span>
                <span style={{ fontSize: 26, fontWeight: 900, color: divMeta.color, letterSpacing: "-1px" }}>
                  ${fmt(tot.grand)}
                </span>
              </div>
            </div>
          </div>
        )}
      </div>
      <div
        style={{
          position: "fixed",
          bottom: 0,
          left: "50%",
          transform: "translateX(-50%)",
          width: "100%",
          maxWidth: 480,
          background: T.bg + "EE",
          backdropFilter: "blur(12px)",
          borderTop: `1px solid ${T.border}`,
          padding: "12px 16px",
          display: "flex",
          gap: 10,
        }}
      >
        {step > 1 && (
          <button
            onClick={() => {
              setStep((s) => s - 1);
              scroll();
            }}
            style={{ ...ghostBtn, flex: 1 }}
          >
            ← Back
          </button>
        )}
        {step < 6 ? (
          <button
            onClick={() => {
              setStep((s) => s + 1);
              scroll();
            }}
            style={{ ...primBtn, flex: 2, borderRadius: 12, background: divMeta.color }}
          >
            {step === 4 ? "Review →" : "Next →"}
          </button>
        ) : (
          <button
            onClick={submit}
            style={{ ...primBtn, flex: 2, borderRadius: 12, background: divMeta.color, opacity: saving ? 0.6 : 1 }}
          >
            {saving ? "Saving…" : "💾 Save Report"}
          </button>
        )}
      </div>
    </div>
  );
}

/* ── PDF / PRINT REPORT ─────────────────────────────────────── */
function printReport(report, project) {
  const positions = getPositions(project.division);
  const tot = reportTotals(report, project.division);
  const [yr, mo, dy] = (report.date || "").split("-");
  const dateStr = `${mo}/${dy}/${yr}`;
  const fmt2 = (n) =>
    (parseFloat(n) || 0).toLocaleString("en-US", { style: "currency", currency: "USD", minimumFractionDigits: 2 });

  const laborRows = [...(report.labor || []).filter((l) => l.classification !== "Per Diem")];
  while (laborRows.length < 14) laborRows.push(null);
  const perDiemEntry = (report.labor || []).find((l) => l.classification === "Per Diem");
  const equipRows = [...(report.equipment || [])];
  while (equipRows.length < 15) equipRows.push(null);
  const mats = report.materials || [];
  // Combine rented equipment into rental section for PDF
  const rentals = (report.rental_equipment || []).map((r) => ({
    qty: r.qty || "",
    description: r.description || "",
    amount: (parseFloat(r.qty) || 0) * (parseFloat(r.rate) || 0) * (parseFloat(r.usage) || 1),
  }));

  function laborRow(lr, i) {
    if (!lr) return `<tr><td></td><td colspan="2"></td><td></td><td></td><td></td><td></td><td></td></tr>`;
    const pos = positions.find((p) => p.name === lr.classification);
    return `<tr>
      <td>${lr.name || ""}</td>
      <td colspan="2">${lr.classification || ""}</td>
      <td class="num">${pos && !pos.flat ? lr.regHrs || 0 : ""}</td>
      <td class="num">${pos && !pos.flat ? lr.otHrs || 0 : ""}</td>
      <td class="num">${pos && !pos.flat ? lr.travelHrs || 0 : ""}</td>
      <td class="num">${pos ? pos.rate : ""}</td>
      <td class="num amt">${fmt2(laborAmt(lr, project.division))}</td>
    </tr>`;
  }

  function equipRow(er) {
    if (!er) return `<tr><td colspan="4"></td><td></td><td></td><td></td><td class="num amt">$0.00</td></tr>`;
    return `<tr>
      <td colspan="4">${er.description || ""}</td>
      <td class="num">${er.qty || ""}</td>
      <td class="num">${er.usage || ""}</td>
      <td class="num">${er.rate || ""}</td>
      <td class="num amt">${fmt2(equipAmt(er))}</td>
    </tr>`;
  }

  function matPair(left, right) {
    return `
    <tr>
      <td class="num">${left ? left.qty || "" : ""}</td>
      <td colspan="3">${left ? left.description || "" : ""}</td>
      <td class="num amt">${left ? fmt2(parseFloat(left.amount) || 0) : "$0.00"}</td>
      <td class="num">${right ? right.qty || "" : ""}</td>
      <td colspan="2">${right ? right.description || "" : ""}</td>
      <td class="num amt">${right ? fmt2(parseFloat(right.amount) || 0) : "$0.00"}</td>
    </tr>
    <tr class="sub-row">
      <td></td><td colspan="3" class="sub-label">Tax</td><td class="num">$0.00</td>
      <td></td><td colspan="2" class="sub-label">Tax</td><td class="num">$0.00</td>
    </tr>
    <tr class="sub-row">
      <td></td><td colspan="3" class="sub-label">Total</td>
      <td class="num">${left ? fmt2(parseFloat(left.amount) || 0) : "$0.00"}</td>
      <td></td><td colspan="2" class="sub-label">Total</td>
      <td class="num">${right ? fmt2(parseFloat(right.amount) || 0) : "$0.00"}</td>
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
    <td colspan="2" style="width:22%">${project.client || ""}</td>
    <td style="width:13%;font-weight:bold;text-align:center">Work Order / PO #</td>
    <td colspan="2" style="width:17%">${project.work_order || ""}</td>
    <td style="width:10%;font-weight:bold;text-align:center">Report Date:</td>
    <td colspan="2" style="width:14%">${dateStr}</td>
  </tr>

  <!-- Row 3: Location -->
  <tr>
    <td style="font-weight:bold">Job Location:</td>
    <td colspan="8">${project.location || ""}</td>
  </tr>

  <!-- Row 4: Job Number / Report # -->
  <tr>
    <td style="font-weight:bold">Job Number:</td>
    <td colspan="2">${project.name || ""}</td>
    <td style="font-weight:bold;text-align:center">Report #:</td>
    <td colspan="5">${report.report_no || ""}</td>
  </tr>

  <!-- Description -->
  <tr>
    <td style="font-weight:bold">Description of Work Done:</td>
    <td colspan="8" style="min-height:16px">${report.description || ""}</td>
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
  ${laborRows.map((lr, i) => laborRow(lr, i)).join("")}
  <!-- Per Diem row -->
  <tr>
    <td></td>
    <td colspan="2" style="font-style:italic">Per Diem</td>
    <td></td><td></td><td></td><td></td>
    <td class="num amt">${fmt2(perDiemEntry ? laborAmt(perDiemEntry, project.division) : 0)}</td>
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
  ${equipRows.map((er) => equipRow(er)).join("")}
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
  ${rentals[0] || rentals[1] ? matPair(rentals[0] || null, rentals[1] || null) : ""}
  ${rentals[2] || rentals[3] ? matPair(rentals[2] || null, rentals[3] || null) : ""}
  ${rentals[4] || rentals[5] ? matPair(rentals[4] || null, rentals[5] || null) : ""}
  ${!rentals[0] && !rentals[1] && !rentals[2] && !rentals[3] && !rentals[4] && !rentals[5] ? matPair(null, null) : ""}
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
    <td class="sig-box" style="height:50px">${report.inspector_name || ""}</td>
    <td class="sig-box">${report.inspector_signed_at ? new Date(report.inspector_signed_at).toLocaleDateString() : ""}</td>
    <td class="sig-box"></td>
    <td class="sig-box"></td>
    <td class="sig-box" style="text-align:center">${report.inspector_signature ? `<img src="${report.inspector_signature}" style="max-height:44px;max-width:100%;object-fit:contain">` : ""}</td>
  </tr>
</table>

</body></html>`;

  const win = window.open("", "_blank", "width=900,height=700");
  win.document.write(html);
  win.document.close();
  win.focus();
  setTimeout(() => {
    win.print();
  }, 400);
}

/* ── SIGNATURE PAD ──────────────────────────────────────────── */
function SignaturePad({ onSave, onCancel, reportName }) {
  const canvasRef = useRef(null);
  const [drawing, setDrawing] = useState(false);
  const [hasStrokes, setHasStrokes] = useState(false);
  const [inspectorName, setInspectorName] = useState("");
  const [saving, setSaving] = useState(false);
  const [sigErr, setSigErr] = useState("");
  const lastPos = useRef(null);

  function getPos(e, canvas) {
    const rect = canvas.getBoundingClientRect();
    const scaleX = canvas.width / rect.width;
    const scaleY = canvas.height / rect.height;
    if (e.touches && e.touches.length > 0) {
      return { x: (e.touches[0].clientX - rect.left) * scaleX, y: (e.touches[0].clientY - rect.top) * scaleY };
    }
    return { x: (e.clientX - rect.left) * scaleX, y: (e.clientY - rect.top) * scaleY };
  }

  function startDraw(e) {
    e.preventDefault();
    const canvas = canvasRef.current;
    if (!canvas) return;
    const ctx = canvas.getContext("2d");
    const pos = getPos(e, canvas);
    ctx.beginPath();
    ctx.moveTo(pos.x, pos.y);
    lastPos.current = pos;
    setDrawing(true);
    setHasStrokes(true);
  }

  function draw(e) {
    e.preventDefault();
    if (!drawing) return;
    const canvas = canvasRef.current;
    if (!canvas) return;
    const ctx = canvas.getContext("2d");
    const pos = getPos(e, canvas);
    ctx.lineWidth = 3;
    ctx.lineCap = "round";
    ctx.lineJoin = "round";
    ctx.strokeStyle = "#F97316";
    ctx.lineTo(pos.x, pos.y);
    ctx.stroke();
    ctx.beginPath();
    ctx.moveTo(pos.x, pos.y);
    lastPos.current = pos;
  }

  function endDraw(e) {
    if (e) e.preventDefault();
    setDrawing(false);
  }

  function clearPad() {
    const canvas = canvasRef.current;
    if (!canvas) return;
    const ctx = canvas.getContext("2d");
    ctx.clearRect(0, 0, canvas.width, canvas.height);
    setHasStrokes(false);
    setSigErr("");
  }

  async function save() {
    if (!hasStrokes || !inspectorName.trim() || saving) return;
    const canvas = canvasRef.current;
    if (!canvas) {
      setSigErr("Signature canvas not ready. Try again.");
      return;
    }
    setSaving(true);
    setSigErr("");
    try {
      // Use JPEG at 0.7 quality — much smaller than PNG, still clear
      const sig = canvas.toDataURL("image/jpeg", 0.7);
      await onSave(inspectorName.trim(), sig);
    } catch (err) {
      setSigErr("Failed to save: " + err.message);
      setSaving(false);
    }
  }

  const canSign = hasStrokes && inspectorName.trim().length > 0 && !saving;

  return (
    <div
      style={{
        position: "fixed",
        inset: 0,
        zIndex: 300,
        background: T.bg,
        display: "flex",
        flexDirection: "column",
        fontFamily: "inherit",
      }}
    >
      {/* Header */}
      <div
        style={{
          background: T.surface,
          borderBottom: `1px solid ${T.border}`,
          padding: "16px 16px 12px",
          flexShrink: 0,
        }}
      >
        <div style={{ fontSize: 18, fontWeight: 900, color: T.text, marginBottom: 2 }}>✍️ Inspector Sign-Off</div>
        <div style={{ fontSize: 12, color: T.muted }}>{reportName}</div>
      </div>

      <div style={{ flex: 1, overflow: "auto", padding: "16px" }}>
        {sigErr && (
          <div
            style={{
              background: T.redLow,
              border: `1px solid ${T.red}40`,
              borderRadius: 10,
              padding: "10px 12px",
              marginBottom: 12,
              fontSize: 13,
              color: T.red,
            }}
          >
            ⚠️ {sigErr}
          </div>
        )}

        {/* Inspector name */}
        <div style={{ marginBottom: 16 }}>
          <label style={lbl}>Inspector Name *</label>
          <input
            type="text"
            placeholder="Print inspector's full name"
            value={inspectorName}
            onChange={(e) => {
              setInspectorName(e.target.value);
              setSigErr("");
            }}
            style={inp}
            autoComplete="off"
          />
        </div>

        {/* Signature canvas */}
        <div style={{ marginBottom: 8 }}>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 8 }}>
            <label style={lbl}>Signature *</label>
            {hasStrokes && (
              <button
                onClick={clearPad}
                style={{
                  background: "none",
                  border: "none",
                  color: T.orange,
                  cursor: "pointer",
                  fontSize: 13,
                  fontFamily: "inherit",
                  fontWeight: 700,
                }}
              >
                Clear
              </button>
            )}
          </div>
          <div
            style={{
              background: "#fff",
              borderRadius: 14,
              border: `2px solid ${hasStrokes ? T.orange : T.border}`,
              overflow: "hidden",
              position: "relative",
              userSelect: "none",
            }}
          >
            {!hasStrokes && (
              <div
                style={{
                  position: "absolute",
                  inset: 0,
                  display: "flex",
                  alignItems: "center",
                  justifyContent: "center",
                  pointerEvents: "none",
                }}
              >
                <div style={{ textAlign: "center", color: "#6B7280" }}>
                  <div style={{ fontSize: 36, marginBottom: 6 }}>✍️</div>
                  <div style={{ fontSize: 14, fontWeight: 600, color: T.text }}>Sign here</div>
                  <div style={{ fontSize: 11, marginTop: 2 }}>Use finger or stylus</div>
                </div>
              </div>
            )}
            <canvas
              ref={canvasRef}
              width={600}
              height={200}
              style={{ width: "100%", height: 200, display: "block", touchAction: "none", cursor: "crosshair" }}
              onMouseDown={startDraw}
              onMouseMove={draw}
              onMouseUp={endDraw}
              onMouseLeave={endDraw}
              onTouchStart={startDraw}
              onTouchMove={draw}
              onTouchEnd={endDraw}
            />
          </div>
          {hasStrokes && (
            <div style={{ fontSize: 11, color: T.green, marginTop: 4, textAlign: "center", fontWeight: 600 }}>
              ✓ Signature captured
            </div>
          )}
        </div>
      </div>

      {/* Buttons */}
      <div
        style={{
          padding: "16px",
          display: "flex",
          flexDirection: "column",
          gap: 10,
          borderTop: `1px solid ${T.border}`,
          background: T.surface,
          flexShrink: 0,
        }}
      >
        <button
          onClick={save}
          style={{
            ...primBtn,
            borderRadius: 14,
            opacity: canSign ? 1 : 0.45,
            background: canSign ? T.green : T.muted,
            color: "#FFFFFF",
          }}
        >
          {saving ? "Saving…" : "✅ Confirm & Sign Report"}
        </button>
        <button onClick={onCancel} style={{ ...ghostBtn, width: "100%", textAlign: "center" }} disabled={saving}>
          Cancel
        </button>
        {!canSign && !saving && (
          <div style={{ fontSize: 11, color: T.muted, textAlign: "center" }}>
            {!inspectorName.trim()
              ? "Enter inspector name above"
              : !hasStrokes
                ? "Draw signature above to continue"
                : ""}
          </div>
        )}
      </div>
    </div>
  );
}

/* ── REPORT DETAIL ──────────────────────────────────────────── */
function ReportDetail({ report: initReport, project, user, onBack, onDelete, onApprove, onFlag }) {
  const [report, setReport] = useState(initReport);
  const [lb, setLb] = useState(null);
  const [flagNote, setFlagNote] = useState("");
  const [flagging, setFlagging] = useState(false);
  const [showSigPad, setShowSigPad] = useState(false);
  const [sigSaving, setSigSaving] = useState(false);
  const [editing, setEditing] = useState(false);
  const [editErr, setEditErr] = useState("");
  const [editSaving, setEditSaving] = useState(false);
  const [editData, setEditData] = useState(null);

  async function saveEdit(updated) {
    setEditSaving(true);
    setEditErr("");
    try {
      await API.reports.update(report.id, { ...updated, status: "submitted", updated_at: new Date().toISOString() });
      setReport((r) => ({ ...r, ...updated, status: "submitted" }));
      setEditing(false);
    } catch (e) {
      setEditErr(e.message);
    }
    setEditSaving(false);
  }
  const tot = reportTotals(report, project.division);
  const sc = { submitted: T.yellow, approved: T.green, flagged: T.red, signed: T.green }[report.status] || T.muted;
  const divColor = DIV_META[project.division]?.color || T.orange;

  async function saveSignature(inspectorName, sigData) {
    setSigSaving(true);
    try {
      if (!report || !report.id) {
        setSigSaving(false);
        throw new Error("Report ID missing. Please go back and reopen this report.");
      }
      // Compress signature further to avoid payload issues
      const compressSig = () =>
        new Promise((res) => {
          const img = new Image();
          img.onload = () => {
            const c = document.createElement("canvas");
            c.width = Math.min(img.width, 800);
            c.height = Math.round(img.height * (c.width / img.width));
            c.getContext("2d").drawImage(img, 0, 0, c.width, c.height);
            res(c.toDataURL("image/jpeg", 0.5));
          };
          img.src = sigData;
        });
      const compressedSig = await compressSig();
      await API.reports.update(report.id, {
        inspector_name: inspectorName,
        inspector_signature: compressedSig,
        inspector_signed_at: new Date().toISOString(),
        status: "signed",
      });
      setReport((r) => ({
        ...r,
        inspector_name: inspectorName,
        inspector_signature: compressedSig,
        inspector_signed_at: new Date().toISOString(),
        status: "signed",
      }));
      setShowSigPad(false);
    } catch (e) {
      const msg = e.message || "Could not save signature.";
      setErr("Signature save failed: " + msg);
      throw new Error(msg);
    } finally {
      setSigSaving(false);
    }
  }

  async function exportXLSX() {
    const tplResp = await fetch("/daily-report-template.xlsx");
    if (!tplResp.ok) {
      alert("Could not load the Excel template. Make sure daily-report-template.xlsx is in the public folder.");
      return;
    }
    const tplBuf = await tplResp.arrayBuffer();
    const wb = XLSX.read(tplBuf, { type: "array", cellStyles: true, cellFormula: true });
    const ws = wb.Sheets["3-24-2026"];

    // Clear ALL formulas from the sheet first to prevent #REF! / #VALUE!
    Object.keys(ws).forEach((addr) => {
      if (addr.startsWith("!")) return;
      const cell = ws[addr];
      if (cell && cell.f) {
        delete cell.f;
        cell.v = cell.v || 0;
      }
    });

    // Helper: set value, preserve existing border/style, clear formula
    function sc(addr, val) {
      const existing = ws[addr] || {};
      const s = existing.s || {};
      if (typeof val === "number") {
        ws[addr] = { ...existing, s, t: "n", v: val, w: undefined, f: undefined };
      } else {
        ws[addr] = { ...existing, s, t: "s", v: val == null ? "" : String(val), w: undefined, f: undefined };
      }
    }
    // Helper: set numeric/currency value
    function scn(addr, val) {
      const existing = ws[addr] || {};
      const s = existing.s || {};
      const n = parseFloat(val) || 0;
      ws[addr] = { ...existing, s, t: "n", v: n, z: '"$"#,##0.00', w: undefined, f: undefined };
    }
    // Helper: ensure a cell exists with $0.00 if not already set
    function scn0(addr) {
      const existing = ws[addr] || {};
      if (!existing.v) scn(addr, 0);
    }

    const positions = getPositions(project.division);
    const [yr, mo, dy] = (report.date || "").split("-");
    const dateStr = `${mo}/${dy}/${yr}`;
    const tot = reportTotals(report, project.division);

    // ── HEADER ──
    sc("D3", project.client || "");
    sc("H3", project.work_order || "");
    sc("J3", dateStr);
    sc("D4", project.location || "");
    sc("C5", project.name || "");
    sc("F5", report.report_no || "");
    sc("B7", report.description || "");

    // ── LABOR rows 10-23 ──
    const laborRows = [...(report.labor || []).filter((l) => l.classification !== "Per Diem")];
    while (laborRows.length < 14) laborRows.push(null);
    laborRows.slice(0, 14).forEach((lr, i) => {
      const row = 10 + i;
      if (lr) {
        const pos = positions.find((p) => p.name === lr.classification);
        sc(`B${row}`, lr.name || "");
        sc(`D${row}`, lr.classification || "");
        if (pos && !pos.flat) {
          sc(`F${row}`, parseFloat(lr.regHrs) || 0);
          sc(`G${row}`, parseFloat(lr.otHrs) || 0);
          sc(`H${row}`, parseFloat(lr.travelHrs) || 0);
        } else {
          sc(`F${row}`, 0);
          sc(`G${row}`, 0);
          sc(`H${row}`, 0);
        }
        sc(`I${row}`, pos ? pos.rate : 0);
        scn(`J${row}`, laborAmt(lr, project.division));
      } else {
        sc(`B${row}`, "");
        sc(`D${row}`, "");
        sc(`F${row}`, 0);
        sc(`G${row}`, 0);
        sc(`H${row}`, 0);
        sc(`I${row}`, 0);
        scn(`J${row}`, 0);
      }
    });

    // ── Row 24: Per Diem ──
    const perDiemEntry = (report.labor || []).find((l) => l.classification === "Per Diem");
    scn("J24", perDiemEntry ? laborAmt(perDiemEntry, project.division) : 0);

    // ── Row 25: Total Labor ──
    scn("J25", tot.labor);

    // ── EQUIPMENT rows 29-43 ──
    const equipRows = [...(report.equipment || [])];
    while (equipRows.length < 15) equipRows.push(null);
    equipRows.slice(0, 15).forEach((er, i) => {
      const row = 29 + i;
      if (er) {
        sc(`B${row}`, er.description || "");
        sc(`G${row}`, parseFloat(er.qty) || 0);
        sc(`H${row}`, parseFloat(er.usage) || 0);
        sc(`I${row}`, parseFloat(er.rate) || 0);
        scn(`J${row}`, equipAmt(er));
      } else {
        sc(`B${row}`, "");
        sc(`G${row}`, 0);
        sc(`H${row}`, 0);
        sc(`I${row}`, 0);
        scn(`J${row}`, 0);
      }
    });

    // ── Row 44: Total Equipment ──
    scn("J44", tot.equip);

    // ── MATERIALS rows 49-60 ──
    const mats = report.materials || [];
    // Item 1 (row 49, totals in 52)
    const m0 = mats[0] || null;
    const m1 = mats[1] || null;
    sc("B49", m0 ? m0.qty || "" : "");
    sc("C49", m0 ? m0.description || "" : "");
    scn("F49", m0 ? parseFloat(m0.amount) || 0 : 0);
    scn("F52", m0 ? parseFloat(m0.amount) || 0 : 0);
    sc("G49", m1 ? m1.qty || "" : "");
    sc("H49", m1 ? m1.description || "" : "");
    scn("J49", m1 ? parseFloat(m1.amount) || 0 : 0);
    scn("J52", m1 ? parseFloat(m1.amount) || 0 : 0);
    scn0("F50");
    scn0("J50");
    // Item 2 (row 53, totals in 56)
    const m2 = mats[2] || null;
    const m3 = mats[3] || null;
    sc("B53", m2 ? m2.qty || "" : "");
    sc("C53", m2 ? m2.description || "" : "");
    scn("F54", m2 ? parseFloat(m2.amount) || 0 : 0);
    scn("F56", m2 ? parseFloat(m2.amount) || 0 : 0);
    sc("G53", m3 ? m3.qty || "" : "");
    sc("H53", m3 ? m3.description || "" : "");
    scn("J54", m3 ? parseFloat(m3.amount) || 0 : 0);
    scn("J56", m3 ? parseFloat(m3.amount) || 0 : 0);
    scn0("F53");
    scn0("J53");
    // Item 3 (row 57, totals in 60)
    const m4 = mats[4] || null;
    const m5 = mats[5] || null;
    sc("B57", m4 ? m4.qty || "" : "");
    sc("C57", m4 ? m4.description || "" : "");
    scn("F58", m4 ? parseFloat(m4.amount) || 0 : 0);
    scn("F60", m4 ? parseFloat(m4.amount) || 0 : 0);
    sc("G57", m5 ? m5.qty || "" : "");
    sc("H57", m5 ? m5.description || "" : "");
    scn("J58", m5 ? parseFloat(m5.amount) || 0 : 0);
    scn("J60", m5 ? parseFloat(m5.amount) || 0 : 0);
    scn0("F57");
    scn0("J57");

    // ── Totals ──
    scn("J61", tot.mats);
    scn("J62", tot.grand);

    // ── Signature row ──
    sc("D64", dateStr);
    if (report.inspector_name) {
      sc("B64", report.inspector_name);
      sc("G64", report.inspector_signed_at ? new Date(report.inspector_signed_at).toLocaleDateString() : "");
    }

    // Rename sheet
    const shIdx = wb.SheetNames.indexOf("3-24-2026");
    if (shIdx >= 0) {
      wb.SheetNames[shIdx] = "Daily Report";
      wb.Sheets["Daily Report"] = ws;
      delete wb.Sheets["3-24-2026"];
    }

    // Ensure fit-to-page is written correctly
    if (ws["!pageSetup"]) {
      ws["!pageSetup"].fitToPage = true;
      ws["!pageSetup"].fitToWidth = 1;
      ws["!pageSetup"].fitToHeight = 1;
      ws["!pageSetup"].scale = undefined;
      ws["!pageSetup"].orientation = "portrait";
      ws["!pageSetup"].paperSize = 1;
    } else {
      ws["!pageSetup"] = { fitToPage: true, fitToWidth: 1, fitToHeight: 1, orientation: "portrait", paperSize: 1 };
    }
    ws["!sheetPr"] = { ...(ws["!sheetPr"] || {}), pageSetUpPr: { fitToPage: true } };
    ws["!margins"] = { left: 0.25, right: 0.25, top: 0.5, bottom: 0.5, header: 0.3, footer: 0.3 };
    XLSX.writeFile(
      wb,
      `AIME_${(project.name || "").replace(/\s+/g, "_")}_${(report.date || "").replace(/-/g, "")}.xlsx`,
      { cellStyles: true, bookSST: false },
    );
  }
  // If editing, render a simplified edit form
  if (editing && editData) {
    return (
      <div style={{ background: T.bg, minHeight: "100vh", fontFamily: "inherit", color: T.text }}>
        <TopBar title="Edit Report" onBack={() => setEditing(false)} />
        <div style={{ padding: "14px 16px 100px" }}>
          <ErrBanner msg={editErr} onDismiss={() => setEditErr("")} />
          <div style={{ ...cardS, marginBottom: 12, background: T.yellowLow, border: `1px solid ${T.yellow}40` }}>
            <div style={{ fontSize: 12, color: T.yellow }}>
              ⚠️ Editing this report will reset its status to Submitted and require re-approval.
            </div>
          </div>
          <div style={{ marginBottom: 12 }}>
            <label style={lbl}>Date</label>
            <input
              type="date"
              value={editData.date || ""}
              onChange={(e) => setEditData((d) => ({ ...d, date: e.target.value }))}
              style={inp}
            />
          </div>
          <div style={{ marginBottom: 12 }}>
            <label style={lbl}>Report No.</label>
            <input
              type="text"
              value={editData.report_no || ""}
              onChange={(e) => setEditData((d) => ({ ...d, report_no: e.target.value }))}
              style={inp}
            />
          </div>
          <div style={{ marginBottom: 20 }}>
            <label style={lbl}>Description of Work Done</label>
            <textarea
              rows={4}
              value={editData.description || ""}
              onChange={(e) => setEditData((d) => ({ ...d, description: e.target.value }))}
              style={{ ...inp, resize: "vertical", lineHeight: 1.5 }}
            />
          </div>
          <div style={{ fontSize: 13, color: T.muted, marginBottom: 12 }}>
            To edit labor, equipment, or materials in detail — delete this report and create a new one.
          </div>
          <button
            onClick={() => saveEdit(editData)}
            style={{ ...primBtn, opacity: editSaving ? 0.6 : 1, borderRadius: 14 }}
          >
            {editSaving ? "Saving…" : "Save Changes"}
          </button>
        </div>
      </div>
    );
  }

  return (
    <div style={{ background: T.bg, minHeight: "100vh", padding: 16, fontFamily: "inherit" }}>
      {showSigPad && (
        <SignaturePad
          reportName={`${project.name} · ${fmtDate(report.date)}`}
          onSave={saveSignature}
          onCancel={() => setShowSigPad(false)}
        />
      )}
      <Lightbox src={lb} onClose={() => setLb(null)} />
      <button onClick={onBack} style={{ ...ghostBtn, marginBottom: 14 }}>
        ← Reports
      </button>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 4 }}>
        <div style={{ fontSize: 20, fontWeight: 900, letterSpacing: "-0.5px" }}>{fmtDate(report.date)}</div>
        <span style={pill(sc)}>{(report.status || "submitted").toUpperCase()}</span>
      </div>
      {report.submitted_by && (
        <div style={{ fontSize: 12, color: T.muted, marginBottom: 14 }}>by {report.submitted_by}</div>
      )}
      {report.pm_notes && (
        <div style={{ ...cardS, marginBottom: 14, borderLeft: `3px solid ${T.red}`, background: T.redLow }}>
          <div style={{ fontSize: 11, color: T.red, fontWeight: 700, marginBottom: 4 }}>🚩 PM NOTE</div>
          <div style={{ fontSize: 13, color: T.sub }}>{report.pm_notes}</div>
        </div>
      )}
      {report.description && (
        <div style={{ ...cardS, marginBottom: 12, borderLeft: `3px solid ${T.blue}` }}>
          <div
            style={{ fontSize: 11, color: T.muted, textTransform: "uppercase", letterSpacing: "1px", marginBottom: 6 }}
          >
            Work Done
          </div>
          <div style={{ fontSize: 14, color: T.sub, lineHeight: 1.6 }}>{report.description}</div>
        </div>
      )}
      {(report.labor || []).length > 0 && (
        <div style={{ ...cardS, marginBottom: 12 }}>
          <div
            style={{
              fontSize: 12,
              color: divColor,
              fontWeight: 700,
              textTransform: "uppercase",
              letterSpacing: "1px",
              marginBottom: 10,
            }}
          >
            Labor{can(user, "view_dashboard") && <span style={{ color: T.green }}> · ${fmt(tot.labor)}</span>}
          </div>
          {report.labor.map((r, i) => (
            <div
              key={i}
              style={{
                display: "flex",
                justifyContent: "space-between",
                padding: "8px 0",
                borderBottom: i < report.labor.length - 1 ? `1px solid ${T.border}` : "none",
              }}
            >
              <div>
                <div style={{ fontSize: 14, fontWeight: 600, color: T.text }}>{r.name || "—"}</div>
                <div style={{ fontSize: 11, color: T.muted }}>
                  {r.classification} · {r.regHrs || 0}reg {r.otHrs || 0}OT {r.travelHrs || 0}tr
                </div>
              </div>
              {can(user, "view_dashboard") && (
                <div style={{ fontSize: 14, fontWeight: 800, color: T.green }}>${fmt(laborAmt(r))}</div>
              )}
            </div>
          ))}
        </div>
      )}
      {(report.equipment || []).length > 0 && (
        <div style={{ ...cardS, marginBottom: 12 }}>
          <div
            style={{
              fontSize: 12,
              color: divColor,
              fontWeight: 700,
              textTransform: "uppercase",
              letterSpacing: "1px",
              marginBottom: 10,
            }}
          >
            Equipment{can(user, "view_dashboard") && <span style={{ color: T.green }}> · ${fmt(tot.equip)}</span>}
          </div>
          {report.equipment.map((r, i) => (
            <div
              key={i}
              style={{
                display: "flex",
                justifyContent: "space-between",
                padding: "8px 0",
                borderBottom: i < report.equipment.length - 1 ? `1px solid ${T.border}` : "none",
              }}
            >
              <div style={{ flex: 1, paddingRight: 10 }}>
                <div style={{ fontSize: 13, fontWeight: 600, color: T.text }}>{r.description}</div>
                <div style={{ fontSize: 11, color: T.muted }}>
                  Qty {r.qty} x {r.usage} {r.unit}
                </div>
              </div>
              {can(user, "view_dashboard") && (
                <div style={{ fontSize: 14, fontWeight: 800, color: T.green }}>
                  ${fmt(equipAmt(r, project.division))}
                </div>
              )}
            </div>
          ))}
        </div>
      )}
      {(report.rental_equipment || []).length > 0 && (
        <div style={{ ...cardS, marginBottom: 12, borderLeft: `3px solid ${T.purple}` }}>
          <div
            style={{
              fontSize: 12,
              color: T.purple,
              fontWeight: 700,
              textTransform: "uppercase",
              letterSpacing: "1px",
              marginBottom: 10,
            }}
          >
            🔑 Rented Equipment
            {can(user, "view_dashboard") && (
              <span style={{ color: T.green }}>
                {" "}
                · $
                {fmt(
                  (report.rental_equipment || []).reduce(
                    (s, r) => s + (parseFloat(r.qty) || 0) * (parseFloat(r.rate) || 0) * (parseFloat(r.usage) || 1),
                    0,
                  ),
                )}
              </span>
            )}
          </div>
          {(report.rental_equipment || []).map((r, i) => {
            const amt = (parseFloat(r.qty) || 0) * (parseFloat(r.rate) || 0) * (parseFloat(r.usage) || 1);
            return (
              <div
                key={i}
                style={{
                  display: "flex",
                  justifyContent: "space-between",
                  padding: "8px 0",
                  borderBottom: i < report.rental_equipment.length - 1 ? `1px solid ${T.border}` : "none",
                }}
              >
                <div style={{ flex: 1, paddingRight: 10 }}>
                  <div style={{ fontSize: 13, fontWeight: 600, color: T.text }}>{r.description}</div>
                  <div style={{ fontSize: 11, color: T.muted }}>
                    Qty {r.qty || 0} × {r.usage || 0} days/hrs @ ${r.rate || 0}
                  </div>
                </div>
                {can(user, "view_dashboard") && (
                  <div style={{ fontSize: 14, fontWeight: 800, color: T.green }}>${fmt(amt)}</div>
                )}
              </div>
            );
          })}
        </div>
      )}
      {(report.materials || []).length > 0 && (
        <div style={{ ...cardS, marginBottom: 12 }}>
          <div
            style={{
              fontSize: 12,
              color: divColor,
              fontWeight: 700,
              textTransform: "uppercase",
              letterSpacing: "1px",
              marginBottom: 10,
            }}
          >
            Materials{can(user, "view_dashboard") && <span style={{ color: T.green }}> · ${fmt(tot.mats)}</span>}
          </div>
          {report.materials.map((r, i) => (
            <div
              key={i}
              style={{
                padding: "8px 0",
                borderBottom: i < report.materials.length - 1 ? `1px solid ${T.border}` : "none",
              }}
            >
              <div
                style={{
                  display: "flex",
                  justifyContent: "space-between",
                  marginBottom: r.receipts?.length > 0 ? 8 : 0,
                }}
              >
                <span style={{ fontSize: 13 }}>
                  {r.qty ? `${r.qty}x ` : ""}
                  {r.description}
                </span>
                {can(user, "view_dashboard") && (
                  <span style={{ fontSize: 13, fontWeight: 700, color: T.green }}>
                    ${fmt(parseFloat(r.amount) || 0)}
                  </span>
                )}
              </div>
              {r.receipts?.length > 0 && (
                <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
                  {r.receipts.map((rc) => (
                    <img
                      key={rc.id}
                      src={rc.src}
                      alt=""
                      onClick={() => setLb(rc.src)}
                      style={{ width: 56, height: 56, objectFit: "cover", borderRadius: 8, cursor: "pointer" }}
                    />
                  ))}
                </div>
              )}
            </div>
          ))}
        </div>
      )}
      {can(user, "view_dashboard") && (
        <div
          style={{
            ...cardS,
            background: divColor + "12",
            border: `1px solid ${divColor}40`,
            marginBottom: 12,
            display: "flex",
            justifyContent: "space-between",
            alignItems: "center",
          }}
        >
          <span style={{ fontSize: 15, fontWeight: 800 }}>Grand Total</span>
          <span style={{ fontSize: 26, fontWeight: 900, color: divColor, letterSpacing: "-1px" }}>
            ${fmt(tot.grand)}
          </span>
        </div>
      )}
      {can(user, "approve_report") && report.status === "submitted" && (
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10, marginBottom: 10 }}>
          <button
            onClick={() => onApprove && onApprove(report.id)}
            style={{
              ...primBtn,
              background: T.greenLow,
              color: T.green,
              border: `1px solid ${T.green}40`,
              borderRadius: 12,
            }}
          >
            ✓ Approve
          </button>
          <button
            onClick={() => setFlagging(!flagging)}
            style={{ ...primBtn, background: T.redLow, color: T.red, border: `1px solid ${T.red}40`, borderRadius: 12 }}
          >
            🚩 Flag
          </button>
        </div>
      )}
      {flagging && (
        <div style={{ ...cardS, marginBottom: 10 }}>
          <label style={lbl}>Flag Note for Crew</label>
          <textarea
            value={flagNote}
            onChange={(e) => setFlagNote(e.target.value)}
            rows={3}
            placeholder="What needs to be corrected…"
            style={{ ...inp, resize: "vertical", marginBottom: 10 }}
          />
          <button
            onClick={() => {
              onFlag && onFlag(report.id, flagNote);
              setFlagging(false);
            }}
            style={{ ...primBtn, borderRadius: 12 }}
          >
            Send Flag
          </button>
        </div>
      )}
      {/* Signature section */}
      {report.inspector_signature ? (
        <div style={{ ...cardS, marginBottom: 12, borderLeft: `3px solid ${T.green}` }}>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 8 }}>
            <div>
              <div
                style={{
                  fontSize: 12,
                  fontWeight: 700,
                  color: T.green,
                  textTransform: "uppercase",
                  letterSpacing: "1px",
                }}
              >
                ✅ Inspector Sign-Off
              </div>
              <div style={{ fontSize: 14, fontWeight: 700, color: T.orange, marginTop: 2 }}>
                {report.inspector_name}
              </div>
              {report.inspector_signed_at && (
                <div style={{ fontSize: 11, color: T.muted, marginTop: 2 }}>
                  {new Date(report.inspector_signed_at).toLocaleString()}
                </div>
              )}
            </div>
          </div>
          <div style={{ background: "#fff", borderRadius: 10, padding: 4, marginTop: 4 }}>
            <img
              src={report.inspector_signature}
              alt="Inspector signature"
              style={{ width: "100%", borderRadius: 8, display: "block" }}
            />
          </div>
        </div>
      ) : (
        <button
          onClick={() => setShowSigPad(true)}
          style={{
            ...primBtn,
            background: T.greenLow,
            color: T.green,
            border: `1px solid ${T.green}40`,
            marginBottom: 12,
            borderRadius: 14,
          }}
        >
          ✍️ Get Inspector Signature
        </button>
      )}

      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10, marginBottom: 10 }}>
        <button
          onClick={() => printReport(report, project)}
          style={{ ...primBtn, background: "#1f3864", color: "#fff", borderRadius: 14 }}
        >
          🖨️ Print / Save PDF
        </button>
        <button
          onClick={exportXLSX}
          style={{
            ...primBtn,
            background: divColor + "15",
            color: divColor,
            border: `1px solid ${divColor}40`,
            borderRadius: 14,
          }}
        >
          📥 Excel (.xlsx)
        </button>
      </div>
      {can(user, "approve_report") && !editing && (
        <button
          onClick={() => {
            setEditData({
              date: report.date,
              report_no: report.report_no || "",
              description: report.description || "",
            });
            setEditing(true);
          }}
          style={{ ...ghostBtn, width: "100%", textAlign: "center", marginBottom: 10 }}
        >
          ✏️ Edit Report
        </button>
      )}
      <button onClick={() => window.confirm("Delete this report?") && onDelete(report.id)} style={dangerBtn}>
        🗑 Delete Report
      </button>
    </div>
  );
}

/* ── TIME CARDS TAB ─────────────────────────────────────────── */
function TimeCardsTab({ projectId, user, onErr }) {
  const [cards, setCards] = useState([]);
  const [loading, setLoading] = useState(true);
  const [showForm, setShowForm] = useState(false);
  const [saving, setSaving] = useState(false);
  const [f, setF] = useState({ worker_name: user.name, date: today(), clock_in: "07:00", clock_out: "", notes: "" });
  const weekStart = getWeekStart();
  async function load() {
    setLoading(true);
    try {
      setCards((await API.timeCards.forProject(projectId)) || []);
    } catch (e) {
      onErr(e.message);
    }
    setLoading(false);
  }
  useEffect(() => {
    load();
  }, [projectId]);
  async function save() {
    if (!f.worker_name || !f.date) return;
    setSaving(true);
    const total_hours = calcHours(f.clock_in, f.clock_out);
    const ot_hours = Math.max(0, total_hours - 8);
    try {
      await API.timeCards.create({ ...f, project_id: projectId, total_hours, ot_hours });
      await load();
      setShowForm(false);
      setF({ worker_name: user.name, date: today(), clock_in: "07:00", clock_out: "", notes: "" });
    } catch (e) {
      onErr(e.message);
    }
    setSaving(false);
  }
  async function remove(id) {
    try {
      await API.timeCards.remove(id);
      await load();
    } catch (e) {
      onErr(e.message);
    }
  }
  const weekCards = cards.filter((c) => c.date >= weekStart);
  const byWorker = {};
  weekCards.forEach((c) => {
    if (!byWorker[c.worker_name]) byWorker[c.worker_name] = { name: c.worker_name, reg: 0, ot: 0, total: 0 };
    byWorker[c.worker_name].total += c.total_hours || 0;
    byWorker[c.worker_name].ot += c.ot_hours || 0;
    byWorker[c.worker_name].reg += Math.max(0, (c.total_hours || 0) - (c.ot_hours || 0));
  });
  const workerRows = Object.values(byWorker).sort((a, b) => b.total - a.total);
  const todayCards = cards.filter((c) => c.date === today());
  const recentCards = cards.filter((c) => c.date !== today()).slice(0, 30);
  return (
    <div>
      <div
        style={{
          background: T.greenLow,
          border: `1px solid ${T.green}40`,
          borderRadius: 12,
          padding: "10px 14px",
          marginBottom: 14,
          fontSize: 12,
          color: T.green,
          lineHeight: 1.5,
        }}
      >
        <strong>⚡ Auto-filled from daily reports</strong> — hours are added automatically when a foreman submits a
        daily report. Manual entries below for anything not on a daily.
      </div>
      <button onClick={() => setShowForm(!showForm)} style={{ ...primBtn, marginBottom: 14, borderRadius: 14 }}>
        {showForm ? "✕ Cancel" : "⏱️ Log Time"}
      </button>
      {showForm && (
        <div style={{ ...cardS, marginBottom: 14, borderLeft: `3px solid ${T.green}` }}>
          <div style={{ marginBottom: 10 }}>
            <label style={lbl}>Worker</label>
            <select
              value={f.worker_name}
              onChange={(e) => setF((x) => ({ ...x, worker_name: e.target.value }))}
              style={inpSel}
            >
              {NAMES.map((n) => (
                <option key={n}>{n}</option>
              ))}
            </select>
          </div>
          <div style={{ marginBottom: 10 }}>
            <label style={lbl}>Date</label>
            <input
              type="date"
              value={f.date}
              onChange={(e) => setF((x) => ({ ...x, date: e.target.value }))}
              style={inp}
            />
          </div>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10, marginBottom: 10 }}>
            <div>
              <label style={lbl}>Clock In</label>
              <input
                type="time"
                value={f.clock_in}
                onChange={(e) => setF((x) => ({ ...x, clock_in: e.target.value }))}
                style={inp}
              />
            </div>
            <div>
              <label style={lbl}>Clock Out</label>
              <input
                type="time"
                value={f.clock_out}
                onChange={(e) => setF((x) => ({ ...x, clock_out: e.target.value }))}
                style={inp}
              />
            </div>
          </div>
          {f.clock_in &&
            f.clock_out &&
            (() => {
              const h = calcHours(f.clock_in, f.clock_out);
              const ot = Math.max(0, h - 8);
              return (
                h > 0 && (
                  <div
                    style={{
                      background: T.greenLow,
                      borderRadius: 10,
                      padding: "10px 12px",
                      marginBottom: 10,
                      display: "flex",
                      gap: 16,
                    }}
                  >
                    <div>
                      <div style={{ fontSize: 18, fontWeight: 900, color: T.green }}>{h.toFixed(2)}h</div>
                      <div style={{ fontSize: 10, color: T.muted, textTransform: "uppercase" }}>Total</div>
                    </div>
                    <div>
                      <div style={{ fontSize: 18, fontWeight: 900, color: ot > 0 ? T.yellow : T.muted }}>
                        {Math.min(h, 8).toFixed(2)}h
                      </div>
                      <div style={{ fontSize: 10, color: T.muted, textTransform: "uppercase" }}>Regular</div>
                    </div>
                    {ot > 0 && (
                      <div>
                        <div style={{ fontSize: 18, fontWeight: 900, color: T.yellow }}>{ot.toFixed(2)}h</div>
                        <div style={{ fontSize: 10, color: T.muted, textTransform: "uppercase" }}>OT</div>
                      </div>
                    )}
                  </div>
                )
              );
            })()}
          <div style={{ marginBottom: 10 }}>
            <label style={lbl}>Notes</label>
            <input
              type="text"
              placeholder="Optional…"
              value={f.notes}
              onChange={(e) => setF((x) => ({ ...x, notes: e.target.value }))}
              style={inp}
            />
          </div>
          <button onClick={save} style={{ ...primBtn, background: T.green, color: "#FFFFFF", borderRadius: 12 }}>
            {saving ? "Saving…" : "Save Time Card"}
          </button>
        </div>
      )}
      {loading && <Spinner />}
      {!loading && (
        <>
          {workerRows.length > 0 && (
            <div style={{ ...cardS, marginBottom: 14 }}>
              <div
                style={{
                  fontSize: 12,
                  fontWeight: 700,
                  color: T.green,
                  textTransform: "uppercase",
                  letterSpacing: "1px",
                  marginBottom: 10,
                }}
              >
                This Week
              </div>
              {workerRows.map((w) => (
                <div
                  key={w.name}
                  style={{
                    display: "flex",
                    justifyContent: "space-between",
                    alignItems: "center",
                    padding: "8px 0",
                    borderBottom: `1px solid ${T.border}`,
                  }}
                >
                  <span style={{ fontSize: 14, fontWeight: 600, color: T.text }}>{w.name}</span>
                  <div style={{ display: "flex", gap: 12, alignItems: "center" }}>
                    <span style={{ fontSize: 12, color: T.muted }}>{w.reg.toFixed(1)}reg</span>
                    {w.ot > 0 && <span style={{ fontSize: 12, color: T.yellow }}>{w.ot.toFixed(1)}OT</span>}
                    <span style={{ fontSize: 15, fontWeight: 800, color: T.green }}>{w.total.toFixed(1)}h</span>
                  </div>
                </div>
              ))}
            </div>
          )}
          {todayCards.length > 0 && (
            <div
              style={{
                fontSize: 12,
                fontWeight: 700,
                color: T.muted,
                textTransform: "uppercase",
                letterSpacing: "1px",
                marginBottom: 10,
              }}
            >
              Today
            </div>
          )}
          {todayCards.map((c) => (
            <div
              key={c.id}
              style={{
                ...cardS,
                marginBottom: 8,
                display: "flex",
                justifyContent: "space-between",
                alignItems: "center",
              }}
            >
              <div>
                <div style={{ fontSize: 14, fontWeight: 700, color: T.text }}>{c.worker_name}</div>
                <div style={{ fontSize: 11, color: T.muted, marginTop: 3 }}>
                  {fmtShort(c.date)}
                  {c.clock_in ? " · " + c.clock_in : ""}
                  {c.clock_out ? " → " + c.clock_out : ""}
                </div>
                {c.notes && <div style={{ fontSize: 11, color: T.sub, marginTop: 2 }}>{c.notes}</div>}
              </div>
              <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
                <div style={{ textAlign: "right" }}>
                  <div style={{ fontSize: 16, fontWeight: 800, color: T.green }}>
                    {(c.total_hours || 0).toFixed(1)}h
                  </div>
                  {(c.ot_hours || 0) > 0 && (
                    <div style={{ fontSize: 10, color: T.yellow }}>{c.ot_hours.toFixed(1)} OT</div>
                  )}
                </div>
                <button
                  onClick={() => remove(c.id)}
                  style={{
                    background: "none",
                    border: "none",
                    color: T.muted,
                    cursor: "pointer",
                    fontSize: 16,
                    padding: 0,
                  }}
                >
                  🗑
                </button>
              </div>
            </div>
          ))}
          {recentCards.length > 0 && (
            <div
              style={{
                fontSize: 12,
                fontWeight: 700,
                color: T.muted,
                textTransform: "uppercase",
                letterSpacing: "1px",
                margin: "14px 0 10px",
              }}
            >
              Recent
            </div>
          )}
          {recentCards.map((c) => (
            <div
              key={c.id}
              style={{
                ...cardS,
                marginBottom: 8,
                display: "flex",
                justifyContent: "space-between",
                alignItems: "center",
              }}
            >
              <div>
                <div style={{ fontSize: 14, fontWeight: 700, color: T.text }}>{c.worker_name}</div>
                <div style={{ fontSize: 11, color: T.muted, marginTop: 3 }}>
                  {fmtShort(c.date)}
                  {c.clock_in ? " · " + c.clock_in : ""}
                  {c.clock_out ? " → " + c.clock_out : ""}
                </div>
              </div>
              <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
                <div style={{ fontSize: 16, fontWeight: 800, color: T.green }}>{(c.total_hours || 0).toFixed(1)}h</div>
                <button
                  onClick={() => remove(c.id)}
                  style={{
                    background: "none",
                    border: "none",
                    color: T.muted,
                    cursor: "pointer",
                    fontSize: 16,
                    padding: 0,
                  }}
                >
                  🗑
                </button>
              </div>
            </div>
          ))}
          {cards.length === 0 && !showForm && (
            <div style={{ textAlign: "center", padding: "40px 0", color: T.muted }}>
              <div style={{ fontSize: 32, marginBottom: 8 }}>⏱️</div>
              <div>No time cards yet.</div>
            </div>
          )}
        </>
      )}
    </div>
  );
}

/* ── CREW/EQUIP ON SITE TAB ─────────────────────────────────── */
function CrewEquipTab({ projectId, user, onErr }) {
  const [equip, setEquip] = useState([]);
  const [loading, setLoading] = useState(true);
  const [showForm, setShowForm] = useState(false);
  const [saving, setSaving] = useState(false);
  const [f, setF] = useState({
    equipment_name: "",
    quantity: 1,
    operator_name: "",
    hours_used: "",
    notes: "",
    date: today(),
  });
  async function load() {
    setLoading(true);
    try {
      setEquip((await API.equipment.forProject(projectId)) || []);
    } catch (e) {
      onErr(e.message);
    }
    setLoading(false);
  }
  useEffect(() => {
    load();
  }, [projectId]);
  async function save() {
    if (!f.equipment_name) return;
    setSaving(true);
    try {
      await API.equipment.create({ ...f, project_id: projectId });
      await load();
      setShowForm(false);
      setF({ equipment_name: "", quantity: 1, operator_name: "", hours_used: "", notes: "", date: today() });
    } catch (e) {
      onErr(e.message);
    }
    setSaving(false);
  }
  async function remove(id) {
    try {
      await API.equipment.remove(id);
      await load();
    } catch (e) {
      onErr(e.message);
    }
  }
  const todayEquip = equip.filter((e) => e.date === today());
  const prevEquip = equip.filter((e) => e.date !== today()).slice(0, 20);
  return (
    <div>
      <div
        style={{
          background: T.greenLow,
          border: `1px solid ${T.green}40`,
          borderRadius: 12,
          padding: "10px 14px",
          marginBottom: 14,
          fontSize: 12,
          color: T.green,
          lineHeight: 1.5,
        }}
      >
        <strong>⚡ Auto-filled from daily reports</strong> — hours are added automatically when a foreman submits a
        daily report. Manual entries below for anything not on a daily.
      </div>
      <button onClick={() => setShowForm(!showForm)} style={{ ...primBtn, marginBottom: 14, borderRadius: 14 }}>
        {showForm ? "✕ Cancel" : "🚜 Log Equipment On Site"}
      </button>
      {showForm && (
        <div style={{ ...cardS, marginBottom: 14, borderLeft: `3px solid ${T.yellow}` }}>
          <div style={{ marginBottom: 10 }}>
            <label style={lbl}>Equipment</label>
            <select
              value={f.equipment_name}
              onChange={(e) => setF((x) => ({ ...x, equipment_name: e.target.value }))}
              style={inpSel}
            >
              <option value="">— Select —</option>
              {EQUIP_LIST.filter((e) => !e.section).map((e) => (
                <option key={e.name} value={e.name}>
                  {e.name}
                </option>
              ))}
            </select>
          </div>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10, marginBottom: 10 }}>
            <div>
              <label style={lbl}>Quantity</label>
              <input
                type="number"
                min="1"
                value={f.quantity}
                onChange={(e) => setF((x) => ({ ...x, quantity: e.target.value }))}
                style={inp}
              />
            </div>
            <div>
              <label style={lbl}>Hours Used</label>
              <input
                type="number"
                min="0"
                step="0.5"
                placeholder="0"
                value={f.hours_used}
                onChange={(e) => setF((x) => ({ ...x, hours_used: e.target.value }))}
                style={inp}
              />
            </div>
          </div>
          <div style={{ marginBottom: 10 }}>
            <label style={lbl}>Operator</label>
            <select
              value={f.operator_name}
              onChange={(e) => setF((x) => ({ ...x, operator_name: e.target.value }))}
              style={inpSel}
            >
              <option value="">— Optional —</option>
              {NAMES.map((n) => (
                <option key={n}>{n}</option>
              ))}
            </select>
          </div>
          <div style={{ marginBottom: 10 }}>
            <label style={lbl}>Date</label>
            <input
              type="date"
              value={f.date}
              onChange={(e) => setF((x) => ({ ...x, date: e.target.value }))}
              style={inp}
            />
          </div>
          <div style={{ marginBottom: 10 }}>
            <label style={lbl}>Notes</label>
            <input
              type="text"
              placeholder="Condition, issues…"
              value={f.notes}
              onChange={(e) => setF((x) => ({ ...x, notes: e.target.value }))}
              style={inp}
            />
          </div>
          <button onClick={save} style={{ ...primBtn, background: T.yellow, color: "#FFFFFF", borderRadius: 12 }}>
            {saving ? "Saving…" : "Save Entry"}
          </button>
        </div>
      )}
      {loading && <Spinner />}
      {!loading && (
        <>
          {todayEquip.length > 0 && (
            <div
              style={{
                fontSize: 12,
                fontWeight: 700,
                color: T.muted,
                textTransform: "uppercase",
                letterSpacing: "1px",
                marginBottom: 10,
              }}
            >
              On Site Today
            </div>
          )}
          {todayEquip.map((e) => (
            <div
              key={e.id}
              style={{
                ...cardS,
                marginBottom: 8,
                borderLeft: `3px solid ${T.yellow}`,
                display: "flex",
                justifyContent: "space-between",
                alignItems: "center",
              }}
            >
              <div>
                <div style={{ fontSize: 14, fontWeight: 700, color: T.text }}>{e.equipment_name}</div>
                <div style={{ fontSize: 11, color: T.muted, marginTop: 3 }}>
                  {fmtShort(e.date)} · Qty {e.quantity || 1}
                  {e.operator_name ? " · " + e.operator_name : ""}
                  {e.hours_used ? " · " + e.hours_used + "h" : ""}
                </div>
              </div>
              <button
                onClick={() => remove(e.id)}
                style={{
                  background: "none",
                  border: "none",
                  color: T.muted,
                  cursor: "pointer",
                  fontSize: 16,
                  padding: 0,
                  marginLeft: 12,
                }}
              >
                🗑
              </button>
            </div>
          ))}
          {prevEquip.length > 0 && (
            <div
              style={{
                fontSize: 12,
                fontWeight: 700,
                color: T.muted,
                textTransform: "uppercase",
                letterSpacing: "1px",
                margin: "14px 0 10px",
              }}
            >
              Previous
            </div>
          )}
          {prevEquip.map((e) => (
            <div
              key={e.id}
              style={{
                ...cardS,
                marginBottom: 8,
                borderLeft: `3px solid ${T.yellow}40`,
                display: "flex",
                justifyContent: "space-between",
                alignItems: "center",
              }}
            >
              <div>
                <div style={{ fontSize: 13, fontWeight: 700 }}>{e.equipment_name}</div>
                <div style={{ fontSize: 11, color: T.muted }}>
                  {fmtShort(e.date)} · Qty {e.quantity || 1}
                  {e.operator_name ? " · " + e.operator_name : ""}
                </div>
              </div>
              <button
                onClick={() => remove(e.id)}
                style={{
                  background: "none",
                  border: "none",
                  color: T.muted,
                  cursor: "pointer",
                  fontSize: 16,
                  padding: 0,
                  marginLeft: 12,
                }}
              >
                🗑
              </button>
            </div>
          ))}
          {equip.length === 0 && !showForm && (
            <div style={{ textAlign: "center", padding: "40px 0", color: T.muted }}>
              <div style={{ fontSize: 32, marginBottom: 8 }}>🚜</div>
              <div>No equipment logged.</div>
            </div>
          )}
        </>
      )}
    </div>
  );
}

/* ── SUBS TAB ───────────────────────────────────────────────── */
function SubsTab({ projectId, user, onErr }) {
  const [subs, setSubs] = useState([]);
  const [loading, setLoading] = useState(true);
  const [showForm, setShowForm] = useState(false);
  const [saving, setSaving] = useState(false);
  const [f, setF] = useState({
    date: today(),
    company_name: "",
    trade: "",
    contact_name: "",
    contact_phone: "",
    workers_count: 1,
    hours_worked: "",
    work_description: "",
  });
  const set = (k, v) => setF((x) => ({ ...x, [k]: v }));
  async function load() {
    setLoading(true);
    try {
      setSubs((await API.subs.forProject(projectId)) || []);
    } catch (e) {
      onErr(e.message);
    }
    setLoading(false);
  }
  useEffect(() => {
    load();
  }, [projectId]);
  async function save() {
    if (!f.company_name) return;
    setSaving(true);
    try {
      await API.subs.create({ ...f, project_id: projectId, created_by: user.name });
      await load();
      setShowForm(false);
      setF({
        date: today(),
        company_name: "",
        trade: "",
        contact_name: "",
        contact_phone: "",
        workers_count: 1,
        hours_worked: "",
        work_description: "",
      });
    } catch (e) {
      onErr(e.message);
    }
    setSaving(false);
  }
  async function remove(id) {
    try {
      await API.subs.remove(id);
      await load();
    } catch (e) {
      onErr(e.message);
    }
  }
  const trades = [
    "Electrical",
    "Mechanical",
    "Civil",
    "Welding",
    "Coating",
    "Survey",
    "Inspection",
    "HDD",
    "Boring",
    "Concrete",
    "Other",
  ];
  return (
    <div>
      <div
        style={{
          background: T.greenLow,
          border: `1px solid ${T.green}40`,
          borderRadius: 12,
          padding: "10px 14px",
          marginBottom: 14,
          fontSize: 12,
          color: T.green,
          lineHeight: 1.5,
        }}
      >
        <strong>⚡ Auto-filled from daily reports</strong> — hours are added automatically when a foreman submits a
        daily report. Manual entries below for anything not on a daily.
      </div>
      <button onClick={() => setShowForm(!showForm)} style={{ ...primBtn, marginBottom: 14, borderRadius: 14 }}>
        {showForm ? "✕ Cancel" : "🏢 Log Subcontractor"}
      </button>
      {showForm && (
        <div style={{ ...cardS, marginBottom: 14, borderLeft: `3px solid ${T.purple}` }}>
          <div style={{ marginBottom: 10 }}>
            <label style={lbl}>Company Name *</label>
            <input
              type="text"
              placeholder="Sub company name"
              value={f.company_name}
              onChange={(e) => set("company_name", e.target.value)}
              style={inp}
            />
          </div>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10, marginBottom: 10 }}>
            <div>
              <label style={lbl}>Trade</label>
              <select value={f.trade} onChange={(e) => set("trade", e.target.value)} style={inpSel}>
                <option value="">— Select —</option>
                {trades.map((t) => (
                  <option key={t}>{t}</option>
                ))}
              </select>
            </div>
            <div>
              <label style={lbl}>Date</label>
              <input type="date" value={f.date} onChange={(e) => set("date", e.target.value)} style={inp} />
            </div>
            <div>
              <label style={lbl}>Contact</label>
              <input
                type="text"
                placeholder="Foreman"
                value={f.contact_name}
                onChange={(e) => set("contact_name", e.target.value)}
                style={inp}
              />
            </div>
            <div>
              <label style={lbl}>Phone</label>
              <input
                type="tel"
                placeholder="555-555-5555"
                value={f.contact_phone}
                onChange={(e) => set("contact_phone", e.target.value)}
                style={inp}
              />
            </div>
            <div>
              <label style={lbl}>Workers</label>
              <input
                type="number"
                min="0"
                value={f.workers_count}
                onChange={(e) => set("workers_count", e.target.value)}
                style={inp}
              />
            </div>
            <div>
              <label style={lbl}>Hours</label>
              <input
                type="number"
                min="0"
                step="0.5"
                placeholder="0"
                value={f.hours_worked}
                onChange={(e) => set("hours_worked", e.target.value)}
                style={inp}
              />
            </div>
          </div>
          <div style={{ marginBottom: 10 }}>
            <label style={lbl}>Work Description</label>
            <textarea
              placeholder="What work was performed?"
              value={f.work_description}
              onChange={(e) => set("work_description", e.target.value)}
              rows={3}
              style={{ ...inp, resize: "vertical" }}
            />
          </div>
          <button onClick={save} style={{ ...primBtn, background: T.purple, color: "#fff", borderRadius: 12 }}>
            {saving ? "Saving…" : "Save Sub Entry"}
          </button>
        </div>
      )}
      {loading && <Spinner />}
      {!loading &&
        subs.map((s) => (
          <div key={s.id} style={{ ...cardS, marginBottom: 10, borderLeft: `3px solid ${T.purple}` }}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start" }}>
              <div style={{ flex: 1 }}>
                <div style={{ fontSize: 15, fontWeight: 800 }}>{s.company_name}</div>
                <div style={{ display: "flex", gap: 6, marginTop: 5, flexWrap: "wrap" }}>
                  {s.trade && <span style={pill(T.purple)}>{s.trade}</span>}
                  <span style={pill(T.muted)}>{fmtShort(s.date)}</span>
                  {s.workers_count > 0 && <span style={pill(T.blue)}>👷 {s.workers_count}</span>}
                  {s.hours_worked > 0 && <span style={pill(T.green)}>{s.hours_worked}h</span>}
                </div>
                {s.contact_name && (
                  <div style={{ fontSize: 12, color: T.sub, marginTop: 6 }}>
                    📞 {s.contact_name}
                    {s.contact_phone ? " · " + s.contact_phone : ""}
                  </div>
                )}
                {s.work_description && (
                  <div style={{ fontSize: 12, color: T.sub, marginTop: 4, lineHeight: 1.5 }}>{s.work_description}</div>
                )}
              </div>
              <button
                onClick={() => remove(s.id)}
                style={{
                  background: "none",
                  border: "none",
                  color: T.muted,
                  cursor: "pointer",
                  fontSize: 16,
                  padding: "0 0 0 10px",
                }}
              >
                🗑
              </button>
            </div>
          </div>
        ))}
      {!loading && subs.length === 0 && !showForm && (
        <div style={{ textAlign: "center", padding: "40px 0", color: T.muted }}>
          <div style={{ fontSize: 32, marginBottom: 8 }}>🏢</div>
          <div>No subcontractors logged.</div>
        </div>
      )}
    </div>
  );
}

/* ── SAFETY TAB ─────────────────────────────────────────────── */
function SafetyTab({ projectId, safety, user, onRefresh, onErr }) {
  const [showForm, setShowForm] = useState(false);
  const [saving, setSaving] = useState(false);
  const [type, setType] = useState("toolbox");
  const [f, setF] = useState({ date: today(), topic: "", notes: "", severity: "low" });
  const TC = { toolbox: T.blue, observation: T.yellow, incident: T.red, nearmiss: T.orange, jsa: T.purple };
  const TL = {
    toolbox: "🛠 Toolbox Talk",
    observation: "👁 Observation",
    incident: "🚨 Incident",
    nearmiss: "⚠️ Near Miss",
    jsa: "📋 JSA",
  };
  async function save() {
    if (!f.topic.trim()) return;
    setSaving(true);
    try {
      await API.safety.create({ ...f, type, project_id: projectId, created_by: user.name });
      await onRefresh();
      setShowForm(false);
      setF({ date: today(), topic: "", notes: "", severity: "low" });
    } catch (e) {
      onErr(e.message);
    }
    setSaving(false);
  }
  async function del(id) {
    try {
      await API.safety.remove(id);
      await onRefresh();
    } catch (e) {
      onErr(e.message);
    }
  }
  return (
    <div>
      <div
        style={{
          background: T.greenLow,
          border: `1px solid ${T.green}40`,
          borderRadius: 12,
          padding: "10px 14px",
          marginBottom: 14,
          fontSize: 12,
          color: T.green,
          lineHeight: 1.5,
        }}
      >
        <strong>⚡ Auto-filled from daily reports</strong> — hours are added automatically when a foreman submits a
        daily report. Manual entries below for anything not on a daily.
      </div>
      <button onClick={() => setShowForm(!showForm)} style={{ ...primBtn, marginBottom: 14, borderRadius: 14 }}>
        {showForm ? "✕ Cancel" : "⛑️ Log Safety Entry"}
      </button>
      {showForm && (
        <div style={{ ...cardS, marginBottom: 14, borderLeft: `3px solid ${T.yellow}` }}>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8, marginBottom: 12 }}>
            {Object.entries(TL).map(([k, v]) => (
              <button
                key={k}
                onClick={() => setType(k)}
                style={{
                  padding: "10px",
                  borderRadius: 10,
                  border: `2px solid ${type === k ? TC[k] : T.border}`,
                  background: type === k ? TC[k] + "20" : T.surface,
                  color: type === k ? TC[k] : T.sub,
                  fontSize: 12,
                  fontWeight: 700,
                  cursor: "pointer",
                  fontFamily: "inherit",
                }}
              >
                {v}
              </button>
            ))}
          </div>
          <div style={{ marginBottom: 10 }}>
            <label style={lbl}>Date</label>
            <input
              type="date"
              value={f.date}
              onChange={(e) => setF((x) => ({ ...x, date: e.target.value }))}
              style={inp}
            />
          </div>
          <div style={{ marginBottom: 10 }}>
            <label style={lbl}>
              {type === "toolbox" ? "Topic" : type === "jsa" ? "Job / Task Name" : "Description"}
            </label>
            <input
              type="text"
              placeholder="Describe…"
              value={f.topic}
              onChange={(e) => setF((x) => ({ ...x, topic: e.target.value }))}
              style={inp}
            />
          </div>
          <div style={{ marginBottom: 10 }}>
            <label style={lbl}>Notes / Corrective Action</label>
            <textarea
              rows={3}
              placeholder="Additional details…"
              value={f.notes}
              onChange={(e) => setF((x) => ({ ...x, notes: e.target.value }))}
              style={{ ...inp, resize: "vertical" }}
            />
          </div>
          {(type === "incident" || type === "nearmiss") && (
            <div style={{ marginBottom: 10 }}>
              <label style={lbl}>Severity</label>
              <select
                value={f.severity}
                onChange={(e) => setF((x) => ({ ...x, severity: e.target.value }))}
                style={inpSel}
              >
                <option value="low">Low – First Aid</option>
                <option value="medium">Medium – Recordable</option>
                <option value="high">High – Lost Time</option>
              </select>
            </div>
          )}
          <button onClick={save} style={{ ...primBtn, background: T.yellow, color: "#FFFFFF", borderRadius: 12 }}>
            {saving ? "Saving…" : "Save Entry"}
          </button>
        </div>
      )}
      {safety.length === 0 && !showForm && (
        <div style={{ textAlign: "center", padding: "40px 0", color: T.muted }}>
          <div style={{ fontSize: 32, marginBottom: 8 }}>⛑️</div>
          <div>No safety entries yet.</div>
        </div>
      )}
      {[...safety].map((s) => (
        <div key={s.id} style={{ ...cardS, marginBottom: 9, borderLeft: `3px solid ${TC[s.type] || T.border}` }}>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start" }}>
            <div style={{ flex: 1 }}>
              <span style={{ ...pill(TC[s.type] || T.muted), marginBottom: 6, display: "inline-flex" }}>
                {TL[s.type] || s.type}
              </span>
              <div style={{ fontSize: 14, fontWeight: 700, marginTop: 4 }}>{s.topic}</div>
              {s.notes && <div style={{ fontSize: 12, color: T.sub, marginTop: 4, lineHeight: 1.5 }}>{s.notes}</div>}
              <div style={{ fontSize: 11, color: T.muted, marginTop: 6 }}>
                {fmtDate(s.date)} · {s.created_by}
              </div>
            </div>
            <div style={{ display: "flex", flexDirection: "column", alignItems: "flex-end", gap: 6 }}>
              {(s.type === "incident" || s.type === "nearmiss") && (
                <span style={pill(s.severity === "high" ? T.red : s.severity === "medium" ? T.yellow : T.green)}>
                  {(s.severity || "low").toUpperCase()}
                </span>
              )}
              <button
                onClick={() => del(s.id)}
                style={{
                  background: "none",
                  border: "none",
                  color: T.muted,
                  cursor: "pointer",
                  fontSize: 16,
                  padding: 0,
                }}
              >
                🗑
              </button>
            </div>
          </div>
        </div>
      ))}
    </div>
  );
}

/* ── DOCS TAB ───────────────────────────────────────────────── */
function DocsTab({ projectId, user, onErr }) {
  const [docs, setDocs] = useState([]);
  const [loading, setLoading] = useState(true);
  const [showForm, setShowForm] = useState(false);
  const [saving, setSaving] = useState(false);
  const [dragging, setDragging] = useState(false);
  const [uploading, setUploading] = useState(false);
  const [editPerms, setEditPerms] = useState(null); // doc being permission-edited
  const fileRef = useRef(null);
  const allRoles = ["admin", "pm", "estimator", "foreman", "crew"];
  const roleLabels = { admin: "🔴 Admin", pm: "🟠 PM", foreman: "🟡 Foreman", crew: "🟢 Field Crew" };
  const [f, setF] = useState({
    name: "",
    doc_type: "Drawing",
    notes: "",
    visible_to: ["admin", "pm", "estimator", "foreman", "crew"],
    can_download: ["admin", "pm", "estimator", "foreman", "crew"],
  });
  const docTypes = [
    "Drawing",
    "Specification",
    "Manual",
    "Permit",
    "Contract",
    "As-Built",
    "ITP",
    "Procedure",
    "Safety Plan",
    "Other",
  ];
  const docIcons = {
    Drawing: "📐",
    Specification: "📄",
    Manual: "📗",
    Permit: "🗂️",
    Contract: "📝",
    "As-Built": "🗺️",
    ITP: "✅",
    Procedure: "📋",
    "Safety Plan": "⛑️",
    Other: "📁",
  };
  const mimeIcons = {
    "application/pdf": "📄",
    "image/": "🖼️",
    "application/vnd.openxmlformats-officedocument.spreadsheetml": "📊",
    "application/vnd.openxmlformats-officedocument.wordprocessingml": "📝",
    "application/vnd.openxmlformats-officedocument.presentationml": "📊",
    "video/": "🎬",
    "text/": "📝",
  };

  function getMimeIcon(mime = "") {
    for (const [k, v] of Object.entries(mimeIcons)) {
      if (mime.startsWith(k)) return v;
    }
    return "📁";
  }
  function fmtSize(bytes) {
    if (!bytes) return "";
    if (bytes < 1024) return bytes + "B";
    if (bytes < 1048576) return (bytes / 1024).toFixed(1) + "KB";
    return (bytes / 1048576).toFixed(1) + "MB";
  }
  function guessType(filename = "") {
    const ext = filename.split(".").pop().toLowerCase();
    const map = {
      pdf: "Specification",
      dwg: "Drawing",
      dxf: "Drawing",
      png: "Drawing",
      jpg: "Drawing",
      jpeg: "Drawing",
      xlsx: "Other",
      xls: "Other",
      docx: "Other",
      doc: "Other",
      pptx: "Other",
      ifc: "Drawing",
    };
    return map[ext] || "Other";
  }

  async function load() {
    setLoading(true);
    try {
      setDocs((await API.docs.forProject(projectId)) || []);
    } catch (e) {
      onErr(e.message);
    }
    setLoading(false);
  }
  useEffect(() => {
    load();
  }, [projectId]);

  async function fileToBase64(file) {
    return new Promise((res, rej) => {
      const rd = new FileReader();
      rd.onload = (e) => res(e.target.result);
      rd.onerror = rej;
      rd.readAsDataURL(file);
    });
  }

  async function uploadFiles(files) {
    setUploading(true);
    let uploaded = 0;
    for (const file of files) {
      if (file.size > 8 * 1024 * 1024) {
        onErr(`${file.name} is too large (max 8MB)`);
        continue;
      }
      try {
        const data = await fileToBase64(file);
        await API.docs.create({
          project_id: projectId,
          name: file.name.replace(/\.[^/.]+$/, ""),
          file_name: file.name,
          file_mime: file.type,
          file_size: file.size,
          file_data: data,
          doc_type: guessType(file.name),
          notes: "",
          uploaded_by: user.name,
          visible_to: ["admin", "pm", "estimator", "foreman", "crew"],
          can_download: ["admin", "pm", "estimator", "foreman", "crew"],
        });
        uploaded++;
      } catch (e) {
        onErr(`Failed to upload ${file.name}: ${e.message}`);
      }
    }
    await load();
    setUploading(false);
    if (uploaded > 0) onErr(""); // clear any prior errors
  }

  function handleDrop(e) {
    e.preventDefault();
    setDragging(false);
    uploadFiles(Array.from(e.dataTransfer.files));
  }

  function downloadDoc(doc) {
    if (!doc.file_data) return;
    const a = document.createElement("a");
    a.href = doc.file_data;
    a.download = doc.file_name || doc.name;
    a.click();
  }

  async function savePerms(doc, visible_to, can_download) {
    try {
      (await API.docs.update?.(doc.id, { visible_to, can_download })) ||
        (await sb(`/documents?id=eq.${doc.id}`, { method: "PATCH", body: { visible_to, can_download } }));
      await load();
      setEditPerms(null);
    } catch (e) {
      onErr(e.message);
    }
  }

  async function remove(id) {
    try {
      await API.docs.remove(id);
      await load();
    } catch (e) {
      onErr(e.message);
    }
  }

  // Filter docs user can see
  const visibleDocs = docs.filter((d) => {
    const vt = d.visible_to || ["admin", "pm", "estimator", "foreman", "crew"];
    return vt.includes(user.role) || user.role === "admin";
  });
  const grouped = {};
  visibleDocs.forEach((d) => {
    const t = d.doc_type || "Other";
    if (!grouped[t]) grouped[t] = [];
    grouped[t].push(d);
  });

  return (
    <div>
      {/* Permission editor modal */}
      {editPerms && (
        <div
          onClick={() => setEditPerms(null)}
          style={{
            position: "fixed",
            inset: 0,
            zIndex: 200,
            background: "rgba(0,0,0,0.85)",
            display: "flex",
            alignItems: "flex-end",
            justifyContent: "center",
          }}
        >
          <div
            onClick={(e) => e.stopPropagation()}
            style={{
              background: T.card,
              border: `1px solid ${T.border}`,
              borderRadius: "20px 20px 0 0",
              padding: "20px 20px 40px",
              width: "100%",
              maxWidth: 480,
            }}
          >
            <div style={{ fontSize: 16, fontWeight: 800, marginBottom: 4 }}>{editPerms.name}</div>
            <div style={{ fontSize: 12, color: T.muted, marginBottom: 16 }}>
              Set who can see and download this document
            </div>
            {["visible_to", "can_download"].map((perm) => (
              <div key={perm} style={{ marginBottom: 16 }}>
                <div
                  style={{
                    fontSize: 11,
                    fontWeight: 700,
                    color: T.muted,
                    textTransform: "uppercase",
                    letterSpacing: "1px",
                    marginBottom: 8,
                  }}
                >
                  {perm === "visible_to" ? "👁 Can View" : "📥 Can Download"}
                </div>
                <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 6 }}>
                  {allRoles.map((role) => {
                    const current = editPerms[perm] || allRoles;
                    const on = current.includes(role);
                    const m = ROLE_META[role];
                    return (
                      <button
                        key={role}
                        onClick={() => {
                          const next = on ? current.filter((r) => r !== role) : [...current, role];
                          setEditPerms((p) => ({ ...p, [perm]: next }));
                        }}
                        style={{
                          display: "flex",
                          alignItems: "center",
                          gap: 8,
                          padding: "10px 12px",
                          borderRadius: 10,
                          border: `2px solid ${on ? m.color : T.border}`,
                          background: on ? m.color + "18" : T.surface,
                          cursor: "pointer",
                          fontFamily: "inherit",
                        }}
                      >
                        <div
                          style={{
                            width: 16,
                            height: 16,
                            borderRadius: 4,
                            border: `2px solid ${on ? m.color : T.border}`,
                            background: on ? m.color : "transparent",
                            display: "flex",
                            alignItems: "center",
                            justifyContent: "center",
                            fontSize: 10,
                            color: "#fff",
                            fontWeight: 900,
                            flexShrink: 0,
                          }}
                        >
                          {on ? "✓" : ""}
                        </div>
                        <span style={{ fontSize: 12, fontWeight: 700, color: on ? m.color : T.sub }}>
                          {roleLabels[role]}
                        </span>
                      </button>
                    );
                  })}
                </div>
              </div>
            ))}
            <button
              onClick={() => savePerms(editPerms, editPerms.visible_to || allRoles, editPerms.can_download || allRoles)}
              style={{ ...primBtn, borderRadius: 12 }}
            >
              Save Permissions
            </button>
            <button
              onClick={() => setEditPerms(null)}
              style={{ ...ghostBtn, width: "100%", textAlign: "center", marginTop: 8 }}
            >
              Cancel
            </button>
          </div>
        </div>
      )}

      {/* Drag and drop zone */}
      <div
        onDrop={handleDrop}
        onDragOver={(e) => {
          e.preventDefault();
          setDragging(true);
        }}
        onDragLeave={() => setDragging(false)}
        onClick={() => fileRef.current?.click()}
        style={{
          border: `2px dashed ${dragging ? T.orange : T.teal}`,
          borderRadius: 16,
          background: dragging ? T.orangeLow : T.teal + "0A",
          padding: "28px 16px",
          textAlign: "center",
          cursor: "pointer",
          marginBottom: 14,
          transition: "all 0.15s",
        }}
      >
        <div style={{ fontSize: 36, marginBottom: 8 }}>{uploading ? "⏳" : dragging ? "📂" : "📁"}</div>
        <div style={{ fontSize: 15, fontWeight: 700, color: dragging ? T.orange : T.teal, marginBottom: 4 }}>
          {uploading ? "Uploading…" : dragging ? "Drop to upload" : "Drag & drop files here"}
        </div>
        <div style={{ fontSize: 12, color: T.muted }}>
          or tap to browse · PDF, images, Word, Excel, drawings · max 8MB each
        </div>
        <input
          ref={fileRef}
          type="file"
          multiple
          style={{ display: "none" }}
          onChange={(e) => {
            uploadFiles(Array.from(e.target.files));
            e.target.value = "";
          }}
        />
      </div>

      {loading && <Spinner />}
      {!loading && visibleDocs.length === 0 && (
        <div style={{ textAlign: "center", padding: "32px 0", color: T.muted }}>
          <div style={{ fontSize: 32, marginBottom: 8 }}>📁</div>
          <div>No documents yet. Drag and drop files above to upload.</div>
        </div>
      )}

      {!loading &&
        Object.entries(grouped).map(([type, typeDocs]) => (
          <div key={type} style={{ marginBottom: 16 }}>
            <div
              style={{
                fontSize: 11,
                fontWeight: 700,
                color: T.muted,
                textTransform: "uppercase",
                letterSpacing: "1px",
                marginBottom: 8,
              }}
            >
              {docIcons[type] || "📁"} {type} ({typeDocs.length})
            </div>
            {typeDocs.map((doc) => {
              const icon = doc.file_mime ? getMimeIcon(doc.file_mime) : docIcons[doc.doc_type] || "📁";
              const canDl = (doc.can_download || allRoles).includes(user.role) || user.role === "admin";
              const vt = doc.visible_to || allRoles;
              const restrictedView = !vt.includes("crew") || !vt.includes("foreman");
              return (
                <div key={doc.id} style={{ ...cardS, marginBottom: 8, borderLeft: `3px solid ${T.teal}` }}>
                  <div style={{ display: "flex", alignItems: "flex-start", gap: 12 }}>
                    <div style={{ fontSize: 28, flexShrink: 0, marginTop: 2 }}>{icon}</div>
                    <div style={{ flex: 1, minWidth: 0 }}>
                      <div style={{ fontSize: 14, fontWeight: 700, color: T.text }}>{doc.name}</div>
                      <div style={{ display: "flex", gap: 6, flexWrap: "wrap", marginTop: 4 }}>
                        {doc.file_name && (
                          <span style={pill(T.teal)}>{doc.file_name.split(".").pop().toUpperCase()}</span>
                        )}
                        {doc.file_size && <span style={pill(T.muted)}>{fmtSize(doc.file_size)}</span>}
                        {restrictedView && <span style={pill(T.yellow)}>🔒 Restricted</span>}
                      </div>
                      {doc.notes && <div style={{ fontSize: 12, color: T.sub, marginTop: 4 }}>{doc.notes}</div>}
                      <div style={{ fontSize: 11, color: T.muted, marginTop: 4 }}>Uploaded by {doc.uploaded_by}</div>
                    </div>
                    <div style={{ display: "flex", flexDirection: "column", gap: 6, flexShrink: 0 }}>
                      {doc.file_data && canDl && (
                        <button
                          onClick={() => downloadDoc(doc)}
                          style={{
                            background: T.teal + "20",
                            border: `1px solid ${T.teal}40`,
                            borderRadius: 8,
                            padding: "6px 12px",
                            color: T.teal,
                            fontSize: 12,
                            fontWeight: 700,
                            cursor: "pointer",
                            fontFamily: "inherit",
                            whiteSpace: "nowrap",
                          }}
                        >
                          📥 Download
                        </button>
                      )}
                      {(user.role === "admin" || user.role === "pm") && (
                        <button
                          onClick={() =>
                            setEditPerms({
                              ...doc,
                              visible_to: doc.visible_to || allRoles,
                              can_download: doc.can_download || allRoles,
                            })
                          }
                          style={{
                            background: T.yellowLow,
                            border: `1px solid ${T.yellow}40`,
                            borderRadius: 8,
                            padding: "6px 12px",
                            color: T.yellow,
                            fontSize: 11,
                            fontWeight: 700,
                            cursor: "pointer",
                            fontFamily: "inherit",
                            whiteSpace: "nowrap",
                          }}
                        >
                          🔒 Perms
                        </button>
                      )}
                      <button
                        onClick={() => remove(doc.id)}
                        style={{
                          background: "none",
                          border: "none",
                          color: T.muted,
                          cursor: "pointer",
                          fontSize: 16,
                          padding: "4px 0",
                          textAlign: "center",
                        }}
                      >
                        🗑
                      </button>
                    </div>
                  </div>
                </div>
              );
            })}
          </div>
        ))}
    </div>
  );
}

/* ── SCHEDULE TAB ───────────────────────────────────────────── */
function ScheduleTab({ projectId, user, onErr }) {
  const [milestones, setMilestones] = useState([]);
  const [loading, setLoading] = useState(true);
  const [showForm, setShowForm] = useState(false);
  const [saving, setSaving] = useState(false);
  const [f, setF] = useState({ title: "", description: "", target_date: "", status: "pending" });
  const set = (k, v) => setF((x) => ({ ...x, [k]: v }));
  async function load() {
    setLoading(true);
    try {
      setMilestones((await API.milestones.forProject(projectId)) || []);
    } catch (e) {
      onErr(e.message);
    }
    setLoading(false);
  }
  useEffect(() => {
    load();
  }, [projectId]);
  async function save() {
    if (!f.title.trim()) return;
    setSaving(true);
    try {
      await API.milestones.create({ ...f, project_id: projectId, sort_order: milestones.length });
      await load();
      setShowForm(false);
      setF({ title: "", description: "", target_date: "", status: "pending" });
    } catch (e) {
      onErr(e.message);
    }
    setSaving(false);
  }
  async function toggleStatus(m) {
    const next = { pending: "in_progress", in_progress: "completed", completed: "pending" };
    const completed_date = next[m.status] === "completed" ? today() : null;
    try {
      const [u] = await API.milestones.update(m.id, { status: next[m.status], completed_date });
      setMilestones((ms) => ms.map((x) => (x.id === m.id ? u : x)));
    } catch (e) {
      onErr(e.message);
    }
  }
  async function del(id) {
    try {
      await API.milestones.remove(id);
      await load();
    } catch (e) {
      onErr(e.message);
    }
  }
  const statusColor = { pending: T.muted, in_progress: T.orange, completed: T.green, delayed: T.red };
  const statusIcon = { pending: "○", in_progress: "◐", completed: "●", delayed: "⚠️" };
  const completed = milestones.filter((m) => m.status === "completed").length;
  const total = milestones.length;
  return (
    <div>
      <div
        style={{
          background: T.greenLow,
          border: `1px solid ${T.green}40`,
          borderRadius: 12,
          padding: "10px 14px",
          marginBottom: 14,
          fontSize: 12,
          color: T.green,
          lineHeight: 1.5,
        }}
      >
        <strong>⚡ Auto-filled from daily reports</strong> — hours are added automatically when a foreman submits a
        daily report. Manual entries below for anything not on a daily.
      </div>
      <button onClick={() => setShowForm(!showForm)} style={{ ...primBtn, marginBottom: 14, borderRadius: 14 }}>
        {showForm ? "✕ Cancel" : "📅 + Add Milestone"}
      </button>
      {showForm && (
        <div style={{ ...cardS, marginBottom: 14, borderLeft: `3px solid ${T.blue}` }}>
          <div style={{ marginBottom: 10 }}>
            <label style={lbl}>Milestone Title *</label>
            <input
              type="text"
              placeholder="e.g. HDD Bore Complete"
              value={f.title}
              onChange={(e) => set("title", e.target.value)}
              style={inp}
            />
          </div>
          <div style={{ marginBottom: 10 }}>
            <label style={lbl}>Description</label>
            <input
              type="text"
              placeholder="Optional details…"
              value={f.description}
              onChange={(e) => set("description", e.target.value)}
              style={inp}
            />
          </div>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10, marginBottom: 10 }}>
            <div>
              <label style={lbl}>Target Date</label>
              <input
                type="date"
                value={f.target_date}
                onChange={(e) => set("target_date", e.target.value)}
                style={inp}
              />
            </div>
            <div>
              <label style={lbl}>Status</label>
              <select value={f.status} onChange={(e) => set("status", e.target.value)} style={inpSel}>
                <option value="pending">Pending</option>
                <option value="in_progress">In Progress</option>
                <option value="completed">Completed</option>
                <option value="delayed">Delayed</option>
              </select>
            </div>
          </div>
          <button onClick={save} style={{ ...primBtn, background: T.blue, borderRadius: 12 }}>
            {saving ? "Saving…" : "Save Milestone"}
          </button>
        </div>
      )}
      {loading && <Spinner />}
      {!loading && total > 0 && (
        <div style={{ ...cardS, marginBottom: 14 }}>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 8 }}>
            <div style={{ fontSize: 13, fontWeight: 700 }}>Progress</div>
            <div style={{ fontSize: 13, fontWeight: 700, color: T.green }}>
              {completed}/{total} complete
            </div>
          </div>
          <div style={{ height: 8, background: T.border, borderRadius: 4, overflow: "hidden" }}>
            <div
              style={{
                height: "100%",
                background: `linear-gradient(90deg,${T.orange},${T.green})`,
                borderRadius: 4,
                width: `${(completed / total) * 100}%`,
                transition: "width 0.4s",
              }}
            />
          </div>
        </div>
      )}
      {!loading && milestones.length === 0 && !showForm && (
        <div style={{ textAlign: "center", padding: "40px 0", color: T.muted }}>
          <div style={{ fontSize: 32, marginBottom: 8 }}>📅</div>
          <div>No milestones yet.</div>
        </div>
      )}
      {!loading &&
        milestones.map((m) => {
          const du = daysUntil(m.target_date);
          const overdue = du !== null && du < 0 && m.status !== "completed";
          const dueSoon = du !== null && du >= 0 && du <= 7 && m.status !== "completed";
          return (
            <div
              key={m.id}
              style={{
                ...cardS,
                marginBottom: 10,
                borderLeft: `3px solid ${statusColor[m.status] || T.border}`,
                opacity: m.status === "completed" ? 0.7 : 1,
              }}
            >
              <div style={{ display: "flex", alignItems: "flex-start", gap: 12 }}>
                <button
                  onClick={() => toggleStatus(m)}
                  style={{
                    width: 32,
                    height: 32,
                    borderRadius: "50%",
                    border: `2px solid ${statusColor[m.status] || T.border}`,
                    background: m.status === "completed" ? T.green : T.surface,
                    color: m.status === "completed" ? "#16181D" : statusColor[m.status] || T.muted,
                    cursor: "pointer",
                    display: "flex",
                    alignItems: "center",
                    justifyContent: "center",
                    fontSize: 14,
                    fontWeight: 800,
                    flexShrink: 0,
                    marginTop: 2,
                  }}
                >
                  {statusIcon[m.status] || "○"}
                </button>
                <div style={{ flex: 1, minWidth: 0 }}>
                  <div
                    style={{
                      fontSize: 15,
                      fontWeight: 700,
                      textDecoration: m.status === "completed" ? "line-through" : "none",
                      color: m.status === "completed" ? T.muted : T.text,
                    }}
                  >
                    {m.title}
                  </div>
                  {m.description && <div style={{ fontSize: 12, color: T.sub, marginTop: 2 }}>{m.description}</div>}
                  <div style={{ display: "flex", gap: 8, marginTop: 6, flexWrap: "wrap" }}>
                    <span style={pill(statusColor[m.status] || T.muted)}>
                      {m.status.replace("_", " ").toUpperCase()}
                    </span>
                    {m.target_date && (
                      <span style={pill(overdue ? T.red : dueSoon ? T.yellow : T.muted)}>
                        {overdue
                          ? `${Math.abs(du)}d overdue`
                          : dueSoon
                            ? `Due in ${du}d`
                            : `Target: ${fmtDate(m.target_date)}`}
                      </span>
                    )}
                    {m.completed_date && <span style={pill(T.green)}>Done: {fmtDate(m.completed_date)}</span>}
                  </div>
                </div>
                <button
                  onClick={() => del(m.id)}
                  style={{
                    background: "none",
                    border: "none",
                    color: T.muted,
                    cursor: "pointer",
                    fontSize: 16,
                    padding: 0,
                    flexShrink: 0,
                  }}
                >
                  🗑
                </button>
              </div>
            </div>
          );
        })}
    </div>
  );
}

/* ── PHOTOS TAB ─────────────────────────────────────────────── */
function PhotosTab({ projectId, photos, onRefresh, onErr }) {
  const CATEGORIES = ["Progress", "Safety", "Equipment", "Issue/Deficiency", "Before", "After", "Inspection", "Other"];
  const [caption, setCaption] = useState("");
  const [category, setCategory] = useState("Progress");
  const [saving, setSaving] = useState(false);
  const [lb, setLb] = useState(null);
  const [filterCat, setFilterCat] = useState("All");
  const [annotating, setAnnotating] = useState(null); // photo src being annotated
  const fileRef = useRef(null);
  const annotCanvasRef = useRef(null);
  const [annotColor, setAnnotColor] = useState("#FF0000");
  const [annotMode, setAnnotMode] = useState("draw"); // draw | text
  const [annotText, setAnnotText] = useState("");
  const [annotDrawing, setAnnotDrawing] = useState(false);

  // Geolocation
  async function getGeoLocation() {
    return new Promise((res) => {
      if (!navigator.geolocation) {
        res({ lat: null, lng: null });
        return;
      }
      navigator.geolocation.getCurrentPosition(
        (pos) => res({ lat: pos.coords.latitude, lng: pos.coords.longitude }),
        () => res({ lat: null, lng: null }),
        { timeout: 5000 },
      );
    });
  }

  async function handleFiles(files) {
    setSaving(true);
    try {
      const { lat, lng } = await getGeoLocation();
      for (const f of files) {
        if (!f.type.startsWith("image/")) continue;
        const src = await compressImg(f, 1100, 0.72);
        await API.photos.create({
          project_id: projectId,
          src,
          caption,
          date: today(),
          category,
          lat: lat || null,
          lng: lng || null,
        });
      }
      await onRefresh();
      setCaption("");
    } catch (e) {
      onErr(e.message);
    }
    setSaving(false);
  }

  async function del(id) {
    try {
      await API.photos.remove(id);
      await onRefresh();
    } catch (e) {
      onErr(e.message);
    }
  }

  // Annotation canvas handlers
  function annotStart(e) {
    const c = annotCanvasRef.current;
    if (!c) return;
    const r = c.getBoundingClientRect();
    const x = (e.touches ? e.touches[0].clientX : e.clientX) - r.left;
    const y = (e.touches ? e.touches[0].clientY : e.clientY) - r.top;
    const ctx = c.getContext("2d");
    ctx.beginPath();
    ctx.moveTo(x * (c.width / r.width), y * (c.height / r.height));
    c._drawing = true;
  }
  function annotDraw(e) {
    e.preventDefault();
    const c = annotCanvasRef.current;
    if (!c || !c._drawing) return;
    const r = c.getBoundingClientRect();
    const x = (e.touches ? e.touches[0].clientX : e.clientX) - r.left;
    const y = (e.touches ? e.touches[0].clientY : e.clientY) - r.top;
    const ctx = c.getContext("2d");
    ctx.lineWidth = 3;
    ctx.lineCap = "round";
    ctx.strokeStyle = annotColor;
    ctx.lineTo(x * (c.width / r.width), y * (c.height / r.height));
    ctx.stroke();
    ctx.beginPath();
    ctx.moveTo(x * (c.width / r.width), y * (c.height / r.height));
  }
  function annotEnd(e) {
    const c = annotCanvasRef.current;
    if (c) c._drawing = false;
  }

  async function saveAnnotation() {
    const c = annotCanvasRef.current;
    if (!c) return;
    // Merge original image with canvas overlay
    const img = new Image();
    img.crossOrigin = "anonymous";
    img.onload = async () => {
      const merge = document.createElement("canvas");
      merge.width = img.width;
      merge.height = img.height;
      const mctx = merge.getContext("2d");
      mctx.drawImage(img, 0, 0);
      // Scale canvas overlay to image dimensions
      mctx.drawImage(c, 0, 0, img.width, img.height);
      const annotated = merge.toDataURL("image/jpeg", 0.85);
      const { lat, lng } = await getGeoLocation();
      await API.photos.create({
        project_id: projectId,
        src: annotated,
        caption: `Annotated · ${caption || "Photo"}`,
        date: today(),
        category,
        lat,
        lng,
      });
      await onRefresh();
      setAnnotating(null);
    };
    img.src = annotating;
  }

  // Filter photos by category
  const filtered = filterCat === "All" ? photos : photos.filter((p) => p.category === filterCat);

  // Group by date
  const byDate = {};
  filtered.forEach((p) => {
    const d = p.date || "Unknown";
    if (!byDate[d]) byDate[d] = [];
    byDate[d].push(p);
  });
  const sortedDates = Object.keys(byDate).sort().reverse();

  const catColor = {
    Progress: T.blue,
    Safety: T.red,
    Equipment: T.yellow,
    "Issue/Deficiency": T.red,
    Before: T.purple,
    After: T.green,
    Inspection: T.orange,
    Other: T.muted,
  };

  // Annotation screen
  if (annotating)
    return (
      <div style={{ background: T.bg, minHeight: "100vh", fontFamily: "inherit" }}>
        <div
          style={{
            background: "#1a1a1a",
            padding: "12px 16px",
            display: "flex",
            justifyContent: "space-between",
            alignItems: "center",
            borderBottom: `1px solid ${T.border}`,
          }}
        >
          <div style={{ fontSize: 15, fontWeight: 800, color: T.text }}>✏️ Annotate Photo</div>
          <div style={{ display: "flex", gap: 8 }}>
            {["#FF0000", "#FFFF00", "#00FF00", "#FFFFFF", "#000000"].map((c) => (
              <button
                key={c}
                onClick={() => setAnnotColor(c)}
                style={{
                  width: 24,
                  height: 24,
                  borderRadius: "50%",
                  background: c,
                  border: annotColor === c ? "3px solid #F97316" : "2px solid #555",
                  cursor: "pointer",
                }}
              />
            ))}
          </div>
        </div>
        <div style={{ position: "relative", display: "inline-block", width: "100%" }}>
          <img src={annotating} style={{ width: "100%", display: "block" }} alt="annotate" />
          <canvas
            ref={annotCanvasRef}
            width={800}
            height={600}
            style={{
              position: "absolute",
              top: 0,
              left: 0,
              width: "100%",
              height: "100%",
              touchAction: "none",
              cursor: "crosshair",
            }}
            onMouseDown={annotStart}
            onMouseMove={annotDraw}
            onMouseUp={annotEnd}
            onTouchStart={annotStart}
            onTouchMove={annotDraw}
            onTouchEnd={annotEnd}
          />
        </div>
        <div style={{ display: "flex", gap: 10, padding: "12px 16px" }}>
          <button
            onClick={() => {
              const c = annotCanvasRef.current;
              if (c) c.getContext("2d").clearRect(0, 0, c.width, c.height);
            }}
            style={{ ...ghostBtn, flex: 1, textAlign: "center" }}
          >
            🗑 Clear
          </button>
          <button onClick={() => setAnnotating(null)} style={{ ...ghostBtn, flex: 1, textAlign: "center" }}>
            Cancel
          </button>
          <button onClick={saveAnnotation} style={{ ...primBtn, flex: 2, borderRadius: 12 }}>
            ✅ Save Annotated
          </button>
        </div>
      </div>
    );

  return (
    <div>
      <Lightbox src={lb} onClose={() => setLb(null)} />

      {/* Upload card */}
      <div style={{ ...cardS, marginBottom: 14, borderStyle: "dashed", borderColor: T.orange + "44" }}>
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8, marginBottom: 10 }}>
          <div>
            <label style={lbl}>Category</label>
            <select value={category} onChange={(e) => setCategory(e.target.value)} style={inpSel}>
              {CATEGORIES.map((c) => (
                <option key={c}>{c}</option>
              ))}
            </select>
          </div>
          <div>
            <label style={lbl}>Caption</label>
            <input
              type="text"
              placeholder="Optional caption"
              value={caption}
              onChange={(e) => setCaption(e.target.value)}
              style={inp}
            />
          </div>
        </div>
        <div style={{ fontSize: 10, color: T.muted, marginBottom: 8 }}>
          📍 GPS location will be automatically captured
        </div>
        <button onClick={() => fileRef.current?.click()} style={{ ...primBtn }}>
          {saving ? "Uploading…" : "📷 Add Site Photos"}
        </button>
        <input
          ref={fileRef}
          type="file"
          accept="image/*"
          capture="environment"
          multiple
          style={{ display: "none" }}
          onChange={(e) => {
            handleFiles(Array.from(e.target.files));
            e.target.value = "";
          }}
        />
      </div>

      {/* Category filter */}
      {photos.length > 0 && (
        <div
          style={{
            display: "flex",
            gap: 6,
            overflowX: "auto",
            paddingBottom: 6,
            marginBottom: 12,
            WebkitOverflowScrolling: "touch",
          }}
        >
          {["All", ...CATEGORIES].map((c) => (
            <button
              key={c}
              onClick={() => setFilterCat(c)}
              style={{
                flexShrink: 0,
                padding: "5px 10px",
                borderRadius: 8,
                border: `1px solid ${filterCat === c ? catColor[c] || T.orange : T.border}`,
                background: filterCat === c ? `${catColor[c] || T.orange}15` : T.card,
                color: filterCat === c ? catColor[c] || T.orange : T.muted,
                fontSize: 11,
                fontWeight: 700,
                cursor: "pointer",
                fontFamily: "inherit",
                whiteSpace: "nowrap",
              }}
            >
              {c}
              {c !== "All" ? ` (${photos.filter((p) => p.category === c).length})` : `(${photos.length})`}
            </button>
          ))}
        </div>
      )}

      {filtered.length === 0 && (
        <div style={{ textAlign: "center", padding: "40px 0", color: T.muted }}>
          <div style={{ fontSize: 32, marginBottom: 8 }}>📷</div>
          <div>{filterCat === "All" ? "No photos yet — add some above" : `No ${filterCat} photos yet`}</div>
        </div>
      )}

      {/* Photos grouped by date */}
      {sortedDates.map((date) => (
        <div key={date} style={{ marginBottom: 20 }}>
          <div
            style={{
              fontSize: 11,
              fontWeight: 700,
              color: T.muted,
              textTransform: "uppercase",
              letterSpacing: "1px",
              marginBottom: 8,
              paddingBottom: 4,
              borderBottom: `1px solid ${T.border}`,
            }}
          >
            📅 {date}{" "}
            <span style={{ color: T.border }}>
              ({byDate[date].length} photo{byDate[date].length !== 1 ? "s" : ""})
            </span>
          </div>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8 }}>
            {byDate[date].map((p) => (
              <div
                key={p.id}
                style={{
                  position: "relative",
                  borderRadius: 12,
                  overflow: "hidden",
                  aspectRatio: "4/3",
                  background: T.card,
                }}
              >
                <img
                  src={p.src}
                  alt={p.caption}
                  onClick={() => setLb(p.src)}
                  style={{ width: "100%", height: "100%", objectFit: "cover", cursor: "pointer", display: "block" }}
                />
                {/* Category badge */}
                {p.category && (
                  <div
                    style={{
                      position: "absolute",
                      top: 6,
                      left: 6,
                      background: "rgba(0,0,0,0.75)",
                      borderRadius: 6,
                      padding: "2px 7px",
                      fontSize: 9,
                      fontWeight: 700,
                      color: catColor[p.category] || T.muted,
                    }}
                  >
                    {p.category}
                  </div>
                )}
                {/* GPS badge */}
                {p.lat && (
                  <div
                    style={{
                      position: "absolute",
                      bottom: 24,
                      left: 6,
                      background: "rgba(0,0,0,0.75)",
                      borderRadius: 6,
                      padding: "2px 6px",
                      fontSize: 9,
                      color: "#4ade80",
                    }}
                  >
                    📍 GPS
                  </div>
                )}
                {p.caption && (
                  <div
                    style={{
                      position: "absolute",
                      bottom: 0,
                      left: 0,
                      right: 0,
                      background: "linear-gradient(transparent,rgba(0,0,0,0.8))",
                      padding: "16px 8px 6px",
                      fontSize: 10,
                      color: "#fff",
                    }}
                  >
                    {p.caption}
                  </div>
                )}
                {/* Action buttons */}
                <div style={{ position: "absolute", top: 6, right: 6, display: "flex", gap: 4 }}>
                  <button
                    onClick={() => setAnnotating(p.src)}
                    style={{
                      background: "rgba(0,0,0,0.75)",
                      border: "none",
                      color: "#fff",
                      borderRadius: 6,
                      width: 24,
                      height: 24,
                      fontSize: 11,
                      cursor: "pointer",
                      display: "flex",
                      alignItems: "center",
                      justifyContent: "center",
                    }}
                  >
                    ✏️
                  </button>
                  <button
                    onClick={() => del(p.id)}
                    style={{
                      background: "rgba(0,0,0,0.75)",
                      border: "none",
                      color: "#fff",
                      borderRadius: "50%",
                      width: 24,
                      height: 24,
                      fontSize: 12,
                      cursor: "pointer",
                      display: "flex",
                      alignItems: "center",
                      justifyContent: "center",
                    }}
                  >
                    ✕
                  </button>
                </div>
              </div>
            ))}
          </div>
        </div>
      ))}
    </div>
  );
}

function WeatherTab({ projectId, project, weather, onRefresh, onErr }) {
  const [fetching, setFetching] = useState(false);
  const [liveWeather, setLiveWeather] = useState(null);
  const [manualNote, setManualNote] = useState("");
  const [saving, setSaving] = useState(false);
  async function autoFetch() {
    if (!project.location) {
      onErr("Add a location to this job (Info tab).");
      return;
    }
    setFetching(true);
    setLiveWeather(null);
    try {
      setLiveWeather(await fetchWeather(project.location));
    } catch (e) {
      onErr(e.message);
    }
    setFetching(false);
  }
  async function logWeather() {
    if (!liveWeather) return;
    setSaving(true);
    const c = liveWeather.current;
    const [desc] = WMO[c.weathercode] || ["Unknown"];
    try {
      await API.weather.upsert({
        project_id: projectId,
        date: today(),
        temp_high: liveWeather.daily?.temperature_2m_max?.[0] || c.temperature_2m,
        temp_low: liveWeather.daily?.temperature_2m_min?.[0] || c.temperature_2m,
        conditions: desc,
        wind_speed: c.windspeed_10m,
        precipitation: liveWeather.daily?.precipitation_sum?.[0] || 0,
        notes: manualNote,
      });
      await onRefresh();
      setLiveWeather(null);
      setManualNote("");
    } catch (e) {
      onErr(e.message);
    }
    setSaving(false);
  }
  async function del(id) {
    try {
      await API.weather.remove(id);
      await onRefresh();
    } catch (e) {
      onErr(e.message);
    }
  }
  return (
    <div>
      <button
        onClick={autoFetch}
        style={{ ...primBtn, marginBottom: 14, borderRadius: 14, opacity: fetching ? 0.6 : 1 }}
      >
        {fetching ? "🌐 Fetching…" : "🌤️ Auto-Fetch Today's Weather"}
      </button>
      {!project.location && (
        <div style={{ ...cardS, marginBottom: 14, background: T.yellowLow, border: `1px solid ${T.yellow}40` }}>
          <div style={{ fontSize: 13, color: T.yellow }}>
            ⚠️ Add a location to this job (Info tab) to auto-fetch weather.
          </div>
        </div>
      )}
      {liveWeather &&
        (() => {
          const c = liveWeather.current;
          const [desc, icon] = WMO[c.weathercode] || ["Unknown", "🌡️"];
          return (
            <div style={{ ...cardS, marginBottom: 14, borderLeft: `3px solid ${T.blue}` }}>
              <div
                style={{
                  fontSize: 11,
                  color: T.blue,
                  fontWeight: 700,
                  textTransform: "uppercase",
                  letterSpacing: "1px",
                  marginBottom: 10,
                }}
              >
                Live · {liveWeather.locationName}
              </div>
              <div style={{ display: "flex", alignItems: "center", gap: 12, marginBottom: 12 }}>
                <span style={{ fontSize: 44 }}>{icon}</span>
                <div>
                  <div style={{ fontSize: 28, fontWeight: 900, letterSpacing: "-1px" }}>
                    {Math.round(c.temperature_2m)}°F
                  </div>
                  <div style={{ fontSize: 14, color: T.sub }}>{desc}</div>
                  <div style={{ fontSize: 12, color: T.muted }}>Feels {Math.round(c.apparent_temperature)}°F</div>
                </div>
              </div>
              <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 8, marginBottom: 12 }}>
                {[
                  ["Wind", Math.round(c.windspeed_10m) + " mph"],
                  ["High", Math.round(liveWeather.daily?.temperature_2m_max?.[0] || c.temperature_2m) + "°F"],
                  ["Precip", (liveWeather.daily?.precipitation_sum?.[0] || 0).toFixed(2) + "in"],
                ].map(([l, v]) => (
                  <div key={l} style={{ background: T.surface, borderRadius: 10, padding: "8px", textAlign: "center" }}>
                    <div style={{ fontSize: 13, fontWeight: 700 }}>{v}</div>
                    <div style={{ fontSize: 10, color: T.muted }}>{l}</div>
                  </div>
                ))}
              </div>
              <div style={{ marginBottom: 10 }}>
                <label style={lbl}>Field Notes</label>
                <input
                  type="text"
                  placeholder="Work impacted by weather?"
                  value={manualNote}
                  onChange={(e) => setManualNote(e.target.value)}
                  style={inp}
                />
              </div>
              <button onClick={logWeather} style={{ ...primBtn, background: T.blue, borderRadius: 12 }}>
                {saving ? "Saving…" : "💾 Log This Weather"}
              </button>
            </div>
          );
        })()}
      {weather.length > 0 && (
        <div
          style={{
            fontSize: 12,
            fontWeight: 700,
            color: T.muted,
            textTransform: "uppercase",
            letterSpacing: "1px",
            marginBottom: 10,
          }}
        >
          History
        </div>
      )}
      {weather.map((w) => {
        const entry = Object.entries(WMO).find(([, v]) => v[0] === w.conditions);
        const icon = entry ? entry[1][1] : "🌡️";
        return (
          <div
            key={w.id}
            style={{
              ...cardS,
              marginBottom: 8,
              display: "flex",
              justifyContent: "space-between",
              alignItems: "center",
            }}
          >
            <div>
              <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                <span style={{ fontSize: 20 }}>{icon}</span>
                <div>
                  <div style={{ fontSize: 14, fontWeight: 700 }}>{w.conditions || "Logged"}</div>
                  <div style={{ fontSize: 11, color: T.muted }}>
                    {fmtShort(w.date)}
                    {w.wind_speed ? " · " + Math.round(w.wind_speed) + "mph" : ""}
                    {w.precipitation > 0 ? " · " + w.precipitation + "in" : ""}
                  </div>
                </div>
              </div>
              {w.notes && <div style={{ fontSize: 12, color: T.sub, marginTop: 4 }}>{w.notes}</div>}
            </div>
            <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
              {w.temp_high && (
                <div style={{ textAlign: "right" }}>
                  <div style={{ fontSize: 15, fontWeight: 800, color: T.orange }}>{Math.round(w.temp_high)}°</div>
                  {w.temp_low && <div style={{ fontSize: 10, color: T.muted }}>{Math.round(w.temp_low)}° lo</div>}
                </div>
              )}
              <button
                onClick={() => del(w.id)}
                style={{
                  background: "none",
                  border: "none",
                  color: T.muted,
                  cursor: "pointer",
                  fontSize: 16,
                  padding: 0,
                }}
              >
                🗑
              </button>
            </div>
          </div>
        );
      })}
      {weather.length === 0 && !liveWeather && (
        <div style={{ textAlign: "center", padding: "40px 0", color: T.muted }}>
          <div style={{ fontSize: 32, marginBottom: 8 }}>🌤️</div>
          <div>No weather logs yet.</div>
        </div>
      )}
    </div>
  );
}

/* ── INFO TAB ───────────────────────────────────────────────── */

function InfoTab({ project, user, onEdit, onArchive, onDelete }) {
  return (
    <div>
      <div style={cardS}>
        {[
          ["Division", project.division],
          ["Client", project.client],
          ["Location", project.location],
          ["AFE No.", project.afe],
          ["Work Order", project.work_order],
          ["Start Date", fmtDate(project.start_date)],
          ["Status", project.status],
          ["Created By", project.created_by],
        ].map(([l, v]) =>
          v ? (
            <div
              key={l}
              style={{
                display: "flex",
                justifyContent: "space-between",
                padding: "11px 0",
                borderBottom: `1px solid ${T.border}`,
              }}
            >
              <span style={{ fontSize: 13, color: T.muted }}>{l}</span>
              <span style={{ fontSize: 13, fontWeight: 600 }}>{v}</span>
            </div>
          ) : null,
        )}
      </div>
      {project.notes && (
        <div style={{ ...cardS, marginTop: 12 }}>
          <div
            style={{ fontSize: 11, color: T.muted, textTransform: "uppercase", letterSpacing: "1px", marginBottom: 8 }}
          >
            Notes
          </div>
          <div style={{ fontSize: 14, color: T.sub, lineHeight: 1.6 }}>{project.notes}</div>
        </div>
      )}
      {can(user, "edit_job") && (
        <div style={{ marginTop: 16, display: "flex", flexDirection: "column", gap: 10 }}>
          <button onClick={onEdit} style={{ ...ghostBtn, width: "100%", textAlign: "center" }}>
            ✏️ Edit Job
          </button>
          <button onClick={onArchive} style={{ ...ghostBtn, width: "100%", textAlign: "center", color: T.muted }}>
            {project.status === "active" ? "📦 Archive Job" : "♻️ Restore Job"}
          </button>
          <button onClick={onDelete} style={{ ...dangerBtn }}>
            🗑 Delete Job Permanently
          </button>
        </div>
      )}
    </div>
  );
}

const PTABS = [
  { id: "reports", icon: "📋", label: "Reports", perm: "submit_report" },
  { id: "time", icon: "⏱️", label: "Time", perm: "approve_report" },
  { id: "crew", icon: "🚜", label: "Crew", perm: "crew_equip" },
  { id: "subs", icon: "🏢", label: "Subs", perm: "subs" },
  { id: "safety", icon: "⛑️", label: "Safety", perm: "safety" },
  { id: "docs", icon: "📁", label: "Docs", perm: "docs" },
  { id: "schedule", icon: "📅", label: "Schedule", perm: "schedule" },
  { id: "photos", icon: "📷", label: "Photos", perm: "photos" },
  { id: "weather", icon: "🌤️", label: "Weather", perm: "weather" },
  { id: "co", icon: "📋", label: "CO", perm: "subs" },
  { id: "rfi", icon: "📝", label: "RFI", perm: "subs" },
  { id: "info", icon: "ℹ️", label: "Info", perm: null },
];

function PMDashboard({ onBack, user, projects: initProjects, onRefresh, onErr }) {
  const [projects, setProjects] = useState(initProjects || []);
  const [reports, setReports] = useState([]);
  const [pending, setPending] = useState([]);
  const [loading, setLoading] = useState(true);
  const [err, setErr] = useState("");
  const [pmTab, setPmTab] = useState("overview");
  const [activeReport, setActiveReport] = useState(null);
  const [activeProject, setActiveProject] = useState(null);
  const [unread, setUnread] = useState(0);
  const [showNotifs, setShowNotifs] = useState(false);

  async function load() {
    setLoading(true);
    setErr("");
    try {
      const [projs, reps, pend] = await Promise.all([API.projects.list(), API.reports.all(), API.reports.pending()]);
      setProjects(projs || []);
      setReports(reps || []);
      setPending(pend || []);
    } catch (e) {
      setErr(e.message);
    }
    setLoading(false);
  }
  useEffect(() => {
    load();
  }, []);

  async function approve(id) {
    try {
      await API.reports.update(id, {
        status: "approved",
        approved_by: user.name,
        approved_at: new Date().toISOString(),
      });
      await load();
    } catch (e) {
      setErr(e.message);
    }
  }
  async function flag(id, notes) {
    try {
      await API.reports.update(id, { status: "flagged", pm_notes: notes });
      await load();
    } catch (e) {
      setErr(e.message);
    }
  }

  const fmt = (n) => "$" + Number(n || 0).toLocaleString("en-US", { minimumFractionDigits: 0 });

  if (showNotifs)
    return (
      <div style={{ background: T.bg, minHeight: "100vh", fontFamily: "inherit" }}>
        <TopBar
          title="🔔 Notifications"
          onBack={() => {
            setShowNotifs(false);
            load();
          }}
        />
        <NotificationsPanel
          onClose={() => {
            setShowNotifs(false);
            load();
          }}
        />
      </div>
    );

  if (activeReport && activeProject)
    return (
      <ReportDetail
        report={activeReport}
        project={activeProject}
        user={user}
        onBack={() => {
          setActiveReport(null);
          setActiveProject(null);
          load();
        }}
        onDelete={async (id) => {
          await API.reports.remove(id);
          setActiveReport(null);
          setActiveProject(null);
          load();
        }}
        onApprove={approve}
        onFlag={flag}
      />
    );

  const DMTABS = [
    { id: "overview", l: "📊 Overview" },
    { id: "approvals", l: `✅ Approvals${pending.length > 0 ? " (" + pending.length + ")" : ""}` },
    { id: "workers", l: "👷 Workers" },
    { id: "billing", l: "💰 Billing" },
    { id: "reports", l: "📄 Reports" },
    { id: "users", l: "👤 Users" },
  ];

  // Compute stats
  const allTot = reports.reduce(
    (s, r) => {
      const t = reportTotals(r);
      return { l: s.l + t.labor, e: s.e + t.equip, g: s.g + t.grand };
    },
    { l: 0, e: 0, g: 0 },
  );
  const projMap = {};
  projects.forEach((p) => {
    projMap[p.id] = { ...p, grand: 0, count: 0 };
  });
  reports.forEach((r) => {
    if (!projMap[r.project_id]) return;
    const t = reportTotals(r);
    projMap[r.project_id].grand += t.grand;
    projMap[r.project_id].count++;
  });
  const projRows = Object.values(projMap)
    .filter((p) => p.status === "active")
    .sort((a, b) => b.grand - a.grand);
  const workerHours = {};
  reports.forEach((r) =>
    (r.labor || []).forEach((l) => {
      if (!l.name) return;
      if (!workerHours[l.name]) workerHours[l.name] = { name: l.name, reg: 0, ot: 0, pay: 0 };
      workerHours[l.name].reg += parseFloat(l.regHrs) || 0;
      workerHours[l.name].ot += parseFloat(l.otHrs) || 0;
      workerHours[l.name].pay += laborAmt(l);
    }),
  );
  const workerRows = Object.values(workerHours).sort((a, b) => b.pay - a.pay);

  return (
    <div style={{ background: T.bg, minHeight: "100vh", fontFamily: "inherit" }}>
      <div
        style={{
          background: T.surface,
          borderBottom: `1px solid ${T.border}`,
          padding: "14px 16px",
          position: "sticky",
          top: 0,
          zIndex: 50,
          display: "flex",
          justifyContent: "space-between",
          alignItems: "center",
        }}
      >
        <button
          onClick={onBack}
          style={{
            background: "none",
            border: "none",
            color: T.sub,
            fontSize: 13,
            cursor: "pointer",
            fontFamily: "inherit",
          }}
        >
          ← Back
        </button>
        <div style={{ fontSize: 16, fontWeight: 900, color: T.orange }}>PM Dashboard</div>
        <button
          onClick={() => setShowNotifs(true)}
          style={{ background: "none", border: "none", color: T.muted, fontSize: 20, cursor: "pointer" }}
        >
          {unread > 0 ? "🔔" : "🔕"}
        </button>
      </div>
      <ErrBanner msg={err} onDismiss={() => setErr("")} />

      {/* Tab bar */}
      <div
        style={{
          display: "flex",
          overflowX: "auto",
          borderBottom: `1px solid ${T.border}`,
          background: T.surface,
          WebkitOverflowScrolling: "touch",
        }}
      >
        {DMTABS.map((t) => (
          <button
            key={t.id}
            onClick={() => setPmTab(t.id)}
            style={{
              flexShrink: 0,
              padding: "12px 14px",
              background: "none",
              border: "none",
              borderBottom: `2px solid ${pmTab === t.id ? T.orange : "transparent"}`,
              color: pmTab === t.id ? T.orange : T.muted,
              fontSize: 12,
              fontWeight: 700,
              cursor: "pointer",
              fontFamily: "inherit",
              whiteSpace: "nowrap",
            }}
          >
            {t.l}
          </button>
        ))}
      </div>

      <div style={{ padding: "14px 16px 80px" }}>
        {loading && <Spinner />}

        {/* OVERVIEW */}
        {pmTab === "overview" && !loading && (
          <div>
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10, marginBottom: 14 }}>
              {[
                [`${reports.length}`, "Total Reports", T.blue],
                [`${pending.length}`, "Pending Approval", T.yellow],
                [fmt(allTot.l), "Total Labor", T.green],
                [fmt(allTot.g), "Total Billed", T.orange],
              ].map(([v, l, c]) => (
                <div key={l} style={{ ...cardS, textAlign: "center" }}>
                  <div style={{ fontSize: 22, fontWeight: 900, color: c }}>{v}</div>
                  <div
                    style={{
                      fontSize: 11,
                      color: T.muted,
                      textTransform: "uppercase",
                      letterSpacing: "1px",
                      marginTop: 2,
                    }}
                  >
                    {l}
                  </div>
                </div>
              ))}
            </div>
            <div
              style={{
                fontSize: 12,
                fontWeight: 700,
                color: T.muted,
                textTransform: "uppercase",
                letterSpacing: "1px",
                marginBottom: 8,
              }}
            >
              Active Jobs by Billings
            </div>
            {projRows.slice(0, 5).map((p) => (
              <div
                key={p.id}
                style={{
                  ...cardS,
                  marginBottom: 8,
                  display: "flex",
                  justifyContent: "space-between",
                  alignItems: "center",
                }}
              >
                <div>
                  <div style={{ fontSize: 13, fontWeight: 700, color: T.orange }}>{p.name}</div>
                  <div style={{ fontSize: 11, color: T.muted }}>{p.count} reports</div>
                </div>
                <div style={{ fontSize: 15, fontWeight: 800, color: T.green }}>{fmt(p.grand)}</div>
              </div>
            ))}
          </div>
        )}

        {/* APPROVALS */}
        {pmTab === "approvals" && !loading && (
          <div>
            {pending.length === 0 && (
              <div style={{ textAlign: "center", padding: "40px 0", color: T.muted }}>
                <div style={{ fontSize: 32 }}>✅</div>
                <div style={{ marginTop: 8 }}>All reports approved</div>
              </div>
            )}
            {pending.map((r) => {
              const proj = projects.find((p) => p.id === r.project_id) || { name: "Unknown" };
              const tot = reportTotals(r);
              return (
                <div key={r.id} style={{ ...cardS, marginBottom: 10 }}>
                  <div style={{ fontSize: 14, fontWeight: 700, color: T.orange }}>{proj.name}</div>
                  <div style={{ fontSize: 12, color: T.muted, marginBottom: 8 }}>
                    {r.date} · {r.submitted_by} · {fmt(tot.grand)}
                  </div>
                  <div style={{ display: "flex", gap: 8 }}>
                    <button
                      onClick={() => {
                        setActiveReport(r);
                        setActiveProject(proj);
                      }}
                      style={{ ...ghostBtn, flex: 1, textAlign: "center", fontSize: 13 }}
                    >
                      👁 View
                    </button>
                    <button
                      onClick={() => approve(r.id)}
                      style={{ ...primBtn, flex: 1, fontSize: 13, borderRadius: 10, background: T.green }}
                    >
                      ✓ Approve
                    </button>
                  </div>
                </div>
              );
            })}
          </div>
        )}

        {/* WORKERS */}
        {pmTab === "workers" && !loading && (
          <div>
            {workerRows.length === 0 && (
              <div style={{ textAlign: "center", padding: "40px 0", color: T.muted }}>No labor data yet</div>
            )}
            {workerRows.map((w) => (
              <div
                key={w.name}
                style={{
                  ...cardS,
                  marginBottom: 8,
                  display: "flex",
                  justifyContent: "space-between",
                  alignItems: "center",
                }}
              >
                <div>
                  <div style={{ fontSize: 13, fontWeight: 700, color: T.orange }}>{w.name}</div>
                  <div style={{ fontSize: 11, color: T.muted }}>
                    {w.reg.toFixed(1)}h reg · {w.ot.toFixed(1)}h OT
                  </div>
                </div>
                <div style={{ fontSize: 15, fontWeight: 800, color: T.green }}>{fmt(w.pay)}</div>
              </div>
            ))}
          </div>
        )}

        {/* BILLING */}
        {pmTab === "billing" && !loading && (
          <div>
            {projRows.map((p) => (
              <div key={p.id} style={{ ...cardS, marginBottom: 8 }}>
                <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                  <div>
                    <div style={{ fontSize: 13, fontWeight: 700, color: T.orange }}>{p.name}</div>
                    <div style={{ fontSize: 11, color: T.muted }}>
                      {p.client || "No client"} · {p.count} reports
                    </div>
                  </div>
                  <div style={{ fontSize: 15, fontWeight: 800, color: T.green }}>{fmt(p.grand)}</div>
                </div>
              </div>
            ))}
          </div>
        )}

        {/* REPORTS */}
        {pmTab === "reports" && !loading && (
          <div>
            {reports.slice(0, 50).map((r) => {
              const proj = projects.find((p) => p.id === r.project_id) || { name: "Unknown" };
              const tot = reportTotals(r);
              const statusColor = { approved: T.green, flagged: T.red, submitted: T.yellow }[r.status] || T.muted;
              return (
                <div
                  key={r.id}
                  style={{ ...cardS, marginBottom: 8, borderLeft: `3px solid ${statusColor}`, cursor: "pointer" }}
                  onClick={() => {
                    setActiveReport(r);
                    setActiveProject(proj);
                  }}
                >
                  <div style={{ display: "flex", justifyContent: "space-between" }}>
                    <div>
                      <div style={{ fontSize: 13, fontWeight: 700, color: T.orange }}>{proj.name}</div>
                      <div style={{ fontSize: 11, color: T.muted }}>
                        {r.date} · {r.submitted_by}
                      </div>
                    </div>
                    <div style={{ textAlign: "right" }}>
                      <div style={{ fontSize: 13, fontWeight: 800, color: T.green }}>{fmt(tot.grand)}</div>
                      <span style={pill(statusColor)}>{r.status}</span>
                    </div>
                  </div>
                </div>
              );
            })}
          </div>
        )}

        {/* USERS */}
        {pmTab === "users" && <UserManagementScreen user={user} onBack={() => setPmTab("overview")} />}
      </div>
    </div>
  );
}

function CrewDirectoryScreen({ onBack, user }) {
  const [members, setMembers] = useState([]);
  const [loading, setLoading] = useState(true);
  const [err, setErr] = useState("");
  const [mode, setMode] = useState("list");
  const [active, setActive] = useState(null);
  const [saving, setSaving] = useState(false);
  const blank = {
    name: "",
    classification: "",
    phone: "",
    email: "",
    emergency_contact_name: "",
    emergency_contact_phone: "",
    certifications: [],
    notes: "",
    active: true,
  };
  const [f, setF] = useState({ ...blank });
  const set = (k, v) => setF((x) => ({ ...x, [k]: v }));
  async function load() {
    setLoading(true);
    try {
      setMembers((await API.crew.list()) || []);
    } catch (e) {
      setErr(e.message);
    }
    setLoading(false);
  }
  useEffect(() => {
    load();
  }, []);
  async function save() {
    if (!f.name.trim()) return;
    setSaving(true);
    try {
      if (active) {
        await API.crew.update(active.id, f);
        setActive({ ...active, ...f });
      } else {
        await API.crew.create(f);
      }
      await load();
      setMode("list");
      setActive(null);
      setF({ ...blank });
    } catch (e) {
      setErr(e.message);
    }
    setSaving(false);
  }
  async function remove(id) {
    if (!window.confirm("Remove crew member?")) return;
    try {
      await API.crew.remove(id);
      await load();
      setMode("list");
      setActive(null);
    } catch (e) {
      setErr(e.message);
    }
  }
  function addCert() {
    set("certifications", [...(f.certifications || []), { id: uid(), name: "", expiry: "", cert_number: "" }]);
  }
  function updateCert(i, k, v) {
    const c = [...(f.certifications || [])];
    c[i] = { ...c[i], [k]: v };
    set("certifications", c);
  }
  function removeCert(i) {
    set(
      "certifications",
      (f.certifications || []).filter((_, j) => j !== i),
    );
  }
  const CERT_TYPES = [
    "OSHA 10",
    "OSHA 30",
    "First Aid / CPR",
    "Confined Space Entry",
    "Crane Operator",
    "Welding Certification",
    "Pipeline Operator Qualification",
    "Hydro Test Operator",
    "Excavation Competent Person",
    "H2S Safety",
    "Driver CDL",
    "Other",
  ];
  const active_m = members.filter((m) => m.active);
  const inactive_m = members.filter((m) => !m.active);

  if (mode === "new" || mode === "edit")
    return (
      <div style={{ background: T.bg, minHeight: "100vh", fontFamily: "inherit" }}>
        <TopBar
          title={mode === "edit" ? "Edit Member" : "Add Member"}
          onBack={() => {
            setMode("list");
            setActive(null);
            setF({ ...blank });
          }}
        />
        <div style={{ padding: "16px 16px 100px" }}>
          <ErrBanner msg={err} onDismiss={() => setErr("")} />
          <div style={{ marginBottom: 12 }}>
            <label style={lbl}>Full Name *</label>
            <select value={f.name} onChange={(e) => set("name", e.target.value)} style={inp}>
              <option value="">— Select —</option>
              {NAMES.map((n) => (
                <option key={n}>{n}</option>
              ))}
            </select>
          </div>
          <div style={{ marginBottom: 12 }}>
            <label style={lbl}>Classification</label>
            <select value={f.classification} onChange={(e) => set("classification", e.target.value)} style={inp}>
              <option value="">— Select —</option>
              {POSITIONS.map((p) => (
                <option key={p.name}>{p.name}</option>
              ))}
            </select>
          </div>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10, marginBottom: 12 }}>
            <div>
              <label style={lbl}>Cell Phone</label>
              <input
                type="tel"
                placeholder="555-555-5555"
                value={f.phone}
                onChange={(e) => set("phone", e.target.value)}
                style={inp}
              />
            </div>
            <div>
              <label style={lbl}>Email</label>
              <input
                type="email"
                placeholder="email@example.com"
                value={f.email}
                onChange={(e) => set("email", e.target.value)}
                style={inp}
              />
            </div>
          </div>
          <div style={{ ...cardS, marginBottom: 12 }}>
            <div
              style={{
                fontSize: 12,
                fontWeight: 700,
                color: T.red,
                textTransform: "uppercase",
                letterSpacing: "1px",
                marginBottom: 10,
              }}
            >
              🆘 Emergency Contact
            </div>
            <div style={{ marginBottom: 10 }}>
              <label style={lbl}>Name</label>
              <input
                type="text"
                placeholder="Spouse, parent…"
                value={f.emergency_contact_name}
                onChange={(e) => set("emergency_contact_name", e.target.value)}
                style={inp}
              />
            </div>
            <div>
              <label style={lbl}>Phone</label>
              <input
                type="tel"
                placeholder="555-555-5555"
                value={f.emergency_contact_phone}
                onChange={(e) => set("emergency_contact_phone", e.target.value)}
                style={inp}
              />
            </div>
          </div>
          <div style={{ ...cardS, marginBottom: 12 }}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 10 }}>
              <div
                style={{
                  fontSize: 12,
                  fontWeight: 700,
                  color: T.blue,
                  textTransform: "uppercase",
                  letterSpacing: "1px",
                }}
              >
                🎖️ Certifications
              </div>
              <button
                onClick={addCert}
                style={{
                  background: T.blueLow,
                  border: `1px solid ${T.blue}40`,
                  borderRadius: 8,
                  padding: "6px 12px",
                  color: T.blue,
                  fontSize: 12,
                  fontWeight: 700,
                  cursor: "pointer",
                  fontFamily: "inherit",
                }}
              >
                + Add
              </button>
            </div>
            {(f.certifications || []).length === 0 && (
              <div style={{ fontSize: 13, color: T.muted, textAlign: "center", padding: "10px 0" }}>
                No certifications added.
              </div>
            )}
            {(f.certifications || []).map((cert, i) => (
              <div
                key={cert.id}
                style={{ borderTop: `1px solid ${T.border}`, paddingTop: 10, marginTop: i > 0 ? 10 : 0 }}
              >
                <div style={{ marginBottom: 8 }}>
                  <label style={lbl}>Certification</label>
                  <select value={cert.name} onChange={(e) => updateCert(i, "name", e.target.value)} style={inp}>
                    <option value="">— Select —</option>
                    {CERT_TYPES.map((t) => (
                      <option key={t}>{t}</option>
                    ))}
                  </select>
                </div>
                <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8, marginBottom: 8 }}>
                  <div>
                    <label style={lbl}>Cert Number</label>
                    <input
                      type="text"
                      placeholder="Optional"
                      value={cert.cert_number}
                      onChange={(e) => updateCert(i, "cert_number", e.target.value)}
                      style={inp}
                    />
                  </div>
                  <div>
                    <label style={lbl}>Expiry Date</label>
                    <input
                      type="date"
                      value={cert.expiry}
                      onChange={(e) => updateCert(i, "expiry", e.target.value)}
                      style={inp}
                    />
                  </div>
                </div>
                <button
                  onClick={() => removeCert(i)}
                  style={{
                    background: "none",
                    border: "none",
                    color: T.red,
                    cursor: "pointer",
                    fontSize: 12,
                    fontWeight: 700,
                    fontFamily: "inherit",
                  }}
                >
                  Remove
                </button>
              </div>
            ))}
          </div>
          <div style={{ marginBottom: 20 }}>
            <label style={lbl}>Notes</label>
            <textarea
              placeholder="Skills, notes, restrictions…"
              value={f.notes}
              onChange={(e) => set("notes", e.target.value)}
              rows={3}
              style={{ ...inp, resize: "vertical" }}
            />
          </div>
          <button onClick={save} style={{ ...primBtn, opacity: f.name && !saving ? 1 : 0.5 }}>
            {saving ? "Saving…" : mode === "edit" ? "Save Changes" : "Add Member"}
          </button>
        </div>
      </div>
    );

  if (mode === "view" && active)
    return (
      <div style={{ background: T.bg, minHeight: "100vh", padding: 16, fontFamily: "inherit" }}>
        <button
          onClick={() => {
            setMode("list");
            setActive(null);
          }}
          style={{ ...ghostBtn, marginBottom: 14 }}
        >
          ← Directory
        </button>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 14 }}>
          <div>
            <div style={{ fontSize: 22, fontWeight: 900, letterSpacing: "-0.5px" }}>{active.name}</div>
            {active.classification && (
              <div style={{ fontSize: 14, color: T.sub, marginTop: 2 }}>{active.classification}</div>
            )}
          </div>
          <button
            onClick={() => {
              setF({ ...blank, ...active, certifications: active.certifications || [] });
              setMode("edit");
            }}
            style={{
              background: T.orangeLow,
              border: `1px solid ${T.orange}40`,
              borderRadius: 10,
              padding: "8px 14px",
              color: T.orange,
              fontSize: 13,
              fontWeight: 700,
              cursor: "pointer",
              fontFamily: "inherit",
            }}
          >
            ✏️ Edit
          </button>
        </div>
        {(active.phone || active.email) && (
          <div style={{ ...cardS, marginBottom: 12 }}>
            <div
              style={{
                fontSize: 12,
                fontWeight: 700,
                color: T.blue,
                textTransform: "uppercase",
                letterSpacing: "1px",
                marginBottom: 10,
              }}
            >
              Contact
            </div>
            {active.phone && (
              <div
                style={{
                  display: "flex",
                  justifyContent: "space-between",
                  padding: "8px 0",
                  borderBottom: `1px solid ${T.border}`,
                }}
              >
                <span style={{ fontSize: 13, color: T.muted }}>Cell</span>
                <a
                  href={`tel:${active.phone}`}
                  style={{ fontSize: 13, fontWeight: 600, color: T.blue, textDecoration: "none" }}
                >
                  {active.phone}
                </a>
              </div>
            )}
            {active.email && (
              <div style={{ display: "flex", justifyContent: "space-between", padding: "8px 0" }}>
                <span style={{ fontSize: 13, color: T.muted }}>Email</span>
                <a
                  href={`mailto:${active.email}`}
                  style={{ fontSize: 13, fontWeight: 600, color: T.blue, textDecoration: "none" }}
                >
                  {active.email}
                </a>
              </div>
            )}
          </div>
        )}
        {(active.emergency_contact_name || active.emergency_contact_phone) && (
          <div style={{ ...cardS, marginBottom: 12, borderLeft: `3px solid ${T.red}` }}>
            <div
              style={{
                fontSize: 12,
                fontWeight: 700,
                color: T.red,
                textTransform: "uppercase",
                letterSpacing: "1px",
                marginBottom: 10,
              }}
            >
              🆘 Emergency Contact
            </div>
            {active.emergency_contact_name && (
              <div style={{ fontSize: 14, fontWeight: 700, marginBottom: 4 }}>{active.emergency_contact_name}</div>
            )}
            {active.emergency_contact_phone && (
              <a
                href={`tel:${active.emergency_contact_phone}`}
                style={{ fontSize: 14, color: T.red, textDecoration: "none", fontWeight: 700 }}
              >
                📞 {active.emergency_contact_phone}
              </a>
            )}
          </div>
        )}
        {(active.certifications || []).length > 0 && (
          <div style={{ ...cardS, marginBottom: 12 }}>
            <div
              style={{
                fontSize: 12,
                fontWeight: 700,
                color: T.blue,
                textTransform: "uppercase",
                letterSpacing: "1px",
                marginBottom: 10,
              }}
            >
              🎖️ Certifications
            </div>
            {(active.certifications || []).map((cert, i) => {
              const exp = cert.expiry ? daysUntil(cert.expiry) : null;
              const expired = exp !== null && exp < 0;
              const expiring = exp !== null && exp >= 0 && exp <= 30;
              return (
                <div
                  key={cert.id || i}
                  style={{
                    display: "flex",
                    justifyContent: "space-between",
                    alignItems: "center",
                    padding: "8px 0",
                    borderBottom: i < active.certifications.length - 1 ? `1px solid ${T.border}` : "none",
                  }}
                >
                  <div>
                    <div style={{ fontSize: 13, fontWeight: 600 }}>{cert.name}</div>
                    {cert.cert_number && <div style={{ fontSize: 11, color: T.muted }}>#{cert.cert_number}</div>}
                  </div>
                  {cert.expiry && (
                    <span style={pill(expired ? T.red : expiring ? T.yellow : T.green)}>
                      {expired ? "EXPIRED" : expiring ? `Exp ${exp}d` : fmtDate(cert.expiry)}
                    </span>
                  )}
                </div>
              );
            })}
          </div>
        )}
        {active.notes && (
          <div style={{ ...cardS, marginBottom: 12 }}>
            <div
              style={{
                fontSize: 11,
                color: T.muted,
                textTransform: "uppercase",
                letterSpacing: "1px",
                marginBottom: 6,
              }}
            >
              Notes
            </div>
            <div style={{ fontSize: 14, color: T.sub, lineHeight: 1.6 }}>{active.notes}</div>
          </div>
        )}
        {can(user, "crew_directory") && (
          <button onClick={() => remove(active.id)} style={{ ...dangerBtn, marginTop: 8 }}>
            Remove from Directory
          </button>
        )}
      </div>
    );

  return (
    <div style={{ background: T.bg, minHeight: "100vh", fontFamily: "inherit" }}>
      <div
        style={{
          background: T.surface,
          borderBottom: `1px solid ${T.border}`,
          padding: "14px 16px",
          position: "sticky",
          top: 0,
          zIndex: 50,
        }}
      >
        <button
          onClick={onBack}
          style={{
            background: "none",
            border: "none",
            color: T.sub,
            fontSize: 13,
            cursor: "pointer",
            marginBottom: 8,
            padding: 0,
            fontFamily: "inherit",
          }}
        >
          ← Back
        </button>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
          <div style={{ fontSize: 20, fontWeight: 900, letterSpacing: "-0.5px" }}>👥 Crew Directory</div>
          <button
            onClick={() => {
              setF({ ...blank });
              setMode("new");
            }}
            style={{
              background: T.orange,
              color: "#FFFFFF",
              border: "none",
              borderRadius: 10,
              padding: "8px 14px",
              fontSize: 13,
              fontWeight: 800,
              cursor: "pointer",
              fontFamily: "inherit",
            }}
          >
            + Add
          </button>
        </div>
      </div>
      <div style={{ padding: "14px 16px 80px" }}>
        <ErrBanner msg={err} onDismiss={() => setErr("")} />
        {loading && <Spinner />}
        {!loading && (
          <>
            {active_m.length === 0 && (
              <div style={{ textAlign: "center", padding: "40px 0", color: T.muted }}>
                <div style={{ fontSize: 36, marginBottom: 8 }}>👥</div>
                <div>No crew members yet.</div>
              </div>
            )}
            {active_m.map((m) => {
              const expiredCerts = (m.certifications || []).filter((c) => c.expiry && daysUntil(c.expiry) < 0);
              const expiringSoon = (m.certifications || []).filter(
                (c) => c.expiry && daysUntil(c.expiry) >= 0 && daysUntil(c.expiry) <= 30,
              );
              return (
                <div
                  key={m.id}
                  onClick={() => {
                    setActive(m);
                    setMode("view");
                  }}
                  style={{
                    ...cardS,
                    marginBottom: 10,
                    cursor: "pointer",
                    display: "flex",
                    alignItems: "center",
                    gap: 12,
                  }}
                >
                  <div
                    style={{
                      width: 44,
                      height: 44,
                      borderRadius: 12,
                      background: T.orangeLow,
                      border: `2px solid ${T.orange}40`,
                      display: "flex",
                      alignItems: "center",
                      justifyContent: "center",
                      fontSize: 16,
                      fontWeight: 800,
                      color: T.orange,
                      flexShrink: 0,
                    }}
                  >
                    {m.name
                      .split(" ")
                      .map((w) => w[0])
                      .slice(0, 2)
                      .join("")}
                  </div>
                  <div style={{ flex: 1, minWidth: 0 }}>
                    <div style={{ fontSize: 15, fontWeight: 700 }}>{m.name}</div>
                    <div style={{ fontSize: 12, color: T.sub }}>
                      {m.classification || "No classification"}
                      {m.phone ? " · " + m.phone : ""}
                    </div>
                    <div style={{ display: "flex", gap: 6, marginTop: 4, flexWrap: "wrap" }}>
                      {(m.certifications || []).length > 0 && (
                        <span style={pill(T.blue)}>{m.certifications.length} certs</span>
                      )}
                      {expiredCerts.length > 0 && <span style={pill(T.red)}>{expiredCerts.length} expired</span>}
                      {expiringSoon.length > 0 && <span style={pill(T.yellow)}>{expiringSoon.length} expiring</span>}
                    </div>
                  </div>
                  <span style={{ fontSize: 16, color: T.muted }}>›</span>
                </div>
              );
            })}
            {inactive_m.length > 0 && (
              <>
                <div
                  style={{
                    fontSize: 12,
                    fontWeight: 700,
                    color: T.muted,
                    textTransform: "uppercase",
                    letterSpacing: "1px",
                    margin: "20px 0 10px",
                  }}
                >
                  Inactive
                </div>
                {inactive_m.map((m) => (
                  <div
                    key={m.id}
                    onClick={() => {
                      setActive(m);
                      setMode("view");
                    }}
                    style={{
                      ...cardS,
                      marginBottom: 8,
                      cursor: "pointer",
                      opacity: 0.5,
                      display: "flex",
                      alignItems: "center",
                      gap: 12,
                    }}
                  >
                    <div
                      style={{
                        width: 36,
                        height: 36,
                        borderRadius: 10,
                        background: T.surface,
                        display: "flex",
                        alignItems: "center",
                        justifyContent: "center",
                        fontSize: 13,
                        fontWeight: 700,
                        color: T.muted,
                        flexShrink: 0,
                      }}
                    >
                      {m.name
                        .split(" ")
                        .map((w) => w[0])
                        .slice(0, 2)
                        .join("")}
                    </div>
                    <div>
                      <div style={{ fontSize: 14, fontWeight: 600 }}>{m.name}</div>
                      <div style={{ fontSize: 12, color: T.muted }}>{m.classification}</div>
                    </div>
                  </div>
                ))}
              </>
            )}
          </>
        )}
      </div>
    </div>
  );
}

/* ── PROJECT DETAIL (ORCHESTRATOR) ─────────────────────────── */

function UserManagementScreen({ onBack, currentUser }) {
  const [profiles, setProfiles] = useState([]);
  const [loading, setLoading] = useState(true);
  const [err, setErr] = useState("");
  const [mode, setMode] = useState("list"); // list | edit
  const [active, setActive] = useState(null);
  const [saving, setSaving] = useState(false);
  const blank = { name: "", role: "crew", division: null, pin: "", active: true };
  const [f, setF] = useState({ ...blank });
  const set = (k, v) => setF((x) => ({ ...x, [k]: v }));

  async function load() {
    setLoading(true);
    try {
      setProfiles((await API.userProfiles.list()) || []);
    } catch (e) {
      setErr(e.message);
    }
    setLoading(false);
  }
  useEffect(() => {
    load();
  }, []);

  async function save() {
    if (!f.name.trim()) return;
    setSaving(true);
    try {
      if (active) {
        await API.userProfiles.update(active.id, { role: f.role, division: f.division, pin: f.pin, active: f.active });
      } else {
        await API.userProfiles.upsert({
          name: f.name,
          role: f.role,
          division: f.division || null,
          pin: f.pin || null,
          active: true,
        });
      }
      await load();
      setMode("list");
      setActive(null);
      setF({ ...blank });
    } catch (e) {
      setErr(e.message);
    }
    setSaving(false);
  }

  async function remove(id) {
    if (!window.confirm("Remove this user profile?")) return;
    try {
      await API.userProfiles.remove(id);
      await load();
    } catch (e) {
      setErr(e.message);
    }
  }

  // Build map of who has a profile
  const profileMap = {};
  profiles.forEach((p) => (profileMap[p.name] = p));
  // Everyone in NAMES list + any extra profiles
  const allNames = [...new Set([...NAMES, ...profiles.map((p) => p.name)])].sort();

  if (mode === "edit")
    return (
      <div style={{ background: T.bg, minHeight: "100vh", fontFamily: "inherit" }}>
        <TopBar
          title={active ? "Edit User" : "Add User"}
          onBack={() => {
            setMode("list");
            setActive(null);
            setF({ ...blank });
          }}
        />
        <div style={{ padding: "16px 16px 100px" }}>
          <ErrBanner msg={err} onDismiss={() => setErr("")} />
          {!active && (
            <div style={{ marginBottom: 14 }}>
              <label style={lbl}>Name</label>
              <select value={f.name} onChange={(e) => set("name", e.target.value)} style={inp}>
                <option value="">— Select —</option>
                {NAMES.map((n) => (
                  <option key={n}>{n}</option>
                ))}
              </select>
            </div>
          )}
          {active && (
            <div style={{ ...cardS, marginBottom: 14 }}>
              <div style={{ fontSize: 16, fontWeight: 800 }}>{active.name}</div>
              <div style={{ fontSize: 12, color: T.muted }}>Editing permissions</div>
            </div>
          )}

          <div style={{ marginBottom: 14 }}>
            <label style={lbl}>Permission Level</label>
            <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
              {ROLES.map((role) => {
                const m = ROLE_META[role];
                return (
                  <button
                    key={role}
                    onClick={() => set("role", role)}
                    style={{
                      display: "flex",
                      alignItems: "center",
                      gap: 12,
                      padding: "14px",
                      borderRadius: 12,
                      border: `2px solid ${f.role === role ? m.color : T.border}`,
                      background: f.role === role ? m.color + "18" : T.surface,
                      cursor: "pointer",
                      fontFamily: "inherit",
                      textAlign: "left",
                    }}
                  >
                    <div style={{ width: 12, height: 12, borderRadius: "50%", background: m.color, flexShrink: 0 }} />
                    <div style={{ flex: 1 }}>
                      <div style={{ fontSize: 14, fontWeight: 700, color: f.role === role ? m.color : T.text }}>
                        {m.label}
                      </div>
                      <div style={{ fontSize: 12, color: T.muted }}>{m.desc}</div>
                    </div>
                    {f.role === role && <div style={{ fontSize: 16, color: m.color }}>✓</div>}
                  </button>
                );
              })}
            </div>
          </div>

          <div style={{ marginBottom: 14 }}>
            <label style={lbl}>Assigned Division (optional)</label>
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8 }}>
              <button
                onClick={() => set("division", null)}
                style={{
                  padding: "12px",
                  borderRadius: 12,
                  border: `2px solid ${f.division === null ? T.orange : T.border}`,
                  background: f.division === null ? T.orangeLow : T.surface,
                  color: f.division === null ? T.orange : T.sub,
                  fontWeight: 600,
                  fontSize: 13,
                  cursor: "pointer",
                  fontFamily: "inherit",
                }}
              >
                All Divisions
              </button>
              {DIVISIONS.map((div) => {
                const m = DIV_META[div];
                return (
                  <button
                    key={div}
                    onClick={() => set("division", div)}
                    style={{
                      padding: "12px",
                      borderRadius: 12,
                      border: `2px solid ${f.division === div ? m.color : T.border}`,
                      background: f.division === div ? m.color + "18" : T.surface,
                      color: f.division === div ? m.color : T.sub,
                      fontWeight: 600,
                      fontSize: 13,
                      cursor: "pointer",
                      fontFamily: "inherit",
                    }}
                  >
                    {m.icon} {div}
                  </button>
                );
              })}
            </div>
          </div>

          {(f.role === "pm" || f.role === "admin") && (
            <div style={{ marginBottom: 14 }}>
              <label style={lbl}>{f.role === "admin" ? "Admin PIN" : "PM PIN"} (required)</label>
              <input
                type="text"
                maxLength={6}
                placeholder="Set a PIN (numbers)"
                value={f.pin || ""}
                onChange={(e) => set("pin", e.target.value)}
                style={inp}
              />
              <div style={{ fontSize: 11, color: T.muted, marginTop: 4 }}>
                This person will need to enter this PIN to sign in.
              </div>
            </div>
          )}

          {active && (
            <div style={{ marginBottom: 14 }}>
              <label style={lbl}>Status</label>
              <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8 }}>
                {[true, false].map((v) => (
                  <button
                    key={String(v)}
                    onClick={() => set("active", v)}
                    style={{
                      padding: "12px",
                      borderRadius: 12,
                      border: `2px solid ${f.active === v ? (v ? T.green : T.red) : T.border}`,
                      background: f.active === v ? (v ? T.greenLow : T.redLow) : T.surface,
                      color: f.active === v ? (v ? T.green : T.red) : T.sub,
                      fontWeight: 700,
                      fontSize: 14,
                      cursor: "pointer",
                      fontFamily: "inherit",
                    }}
                  >
                    {v ? "✅ Active" : "❌ Inactive"}
                  </button>
                ))}
              </div>
            </div>
          )}

          <button onClick={save} style={{ ...primBtn, opacity: f.name && !saving ? 1 : 0.5 }}>
            {saving ? "Saving…" : active ? "Save Changes" : "Add User"}
          </button>
        </div>
      </div>
    );

  return (
    <div style={{ background: T.bg, minHeight: "100vh", fontFamily: "inherit" }}>
      <div
        style={{
          background: T.surface,
          borderBottom: `1px solid ${T.border}`,
          padding: "14px 16px",
          position: "sticky",
          top: 0,
          zIndex: 50,
        }}
      >
        <button
          onClick={onBack}
          style={{
            background: "none",
            border: "none",
            color: T.sub,
            fontSize: 13,
            cursor: "pointer",
            marginBottom: 8,
            padding: 0,
            fontFamily: "inherit",
          }}
        >
          ← Back
        </button>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
          <div>
            <div style={{ fontSize: 20, fontWeight: 900, letterSpacing: "-0.5px" }}>👤 User Management</div>
            <div style={{ fontSize: 12, color: T.muted }}>Set permission levels for your crew</div>
          </div>
          <button
            onClick={() => {
              setF({ ...blank });
              setMode("edit");
            }}
            style={{
              background: T.orange,
              color: "#FFFFFF",
              border: "none",
              borderRadius: 10,
              padding: "8px 14px",
              fontSize: 13,
              fontWeight: 800,
              cursor: "pointer",
              fontFamily: "inherit",
            }}
          >
            + Add
          </button>
        </div>
      </div>
      <div style={{ padding: "14px 16px 80px" }}>
        <ErrBanner msg={err} onDismiss={() => setErr("")} />
        {loading && <Spinner />}
        {!loading && (
          <>
            {/* Users with profiles */}
            <div
              style={{
                fontSize: 12,
                fontWeight: 700,
                color: T.muted,
                textTransform: "uppercase",
                letterSpacing: "1px",
                marginBottom: 12,
              }}
            >
              Configured Users ({profiles.length})
            </div>
            {profiles.length === 0 && (
              <div style={{ ...cardS, marginBottom: 14, background: T.yellowLow, border: `1px solid ${T.yellow}40` }}>
                <div style={{ fontSize: 13, color: T.yellow }}>
                  ⚠️ No user profiles set. All users will sign in as Field Crew until you configure them.
                </div>
              </div>
            )}
            {profiles.map((p) => {
              const m = ROLE_META[p.role] || ROLE_META.crew;
              return (
                <div
                  key={p.id}
                  style={{ ...cardS, marginBottom: 8, borderLeft: `3px solid ${m.color}`, opacity: p.active ? 1 : 0.5 }}
                >
                  <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
                    <div style={{ width: 10, height: 10, borderRadius: "50%", background: m.color, flexShrink: 0 }} />
                    <div style={{ flex: 1 }}>
                      <div style={{ fontSize: 14, fontWeight: 700 }}>{p.name}</div>
                      <div style={{ display: "flex", gap: 6, marginTop: 4, flexWrap: "wrap" }}>
                        <span style={pill(m.color)}>{m.label}</span>
                        {p.division && (
                          <span style={pill(DIV_META[p.division]?.color || T.muted)}>
                            {DIV_META[p.division]?.icon} {p.division}
                          </span>
                        )}
                        {!p.active && <span style={pill(T.red)}>INACTIVE</span>}
                      </div>
                    </div>
                    <div style={{ display: "flex", gap: 6 }}>
                      <button
                        onClick={() => {
                          setActive(p);
                          setF({
                            name: p.name,
                            role: p.role,
                            division: p.division || null,
                            pin: p.pin || "",
                            active: p.active,
                          });
                          setMode("edit");
                        }}
                        style={{
                          background: T.orangeLow,
                          border: `1px solid ${T.orange}40`,
                          borderRadius: 8,
                          padding: "6px 12px",
                          color: T.orange,
                          fontSize: 12,
                          fontWeight: 700,
                          cursor: "pointer",
                          fontFamily: "inherit",
                        }}
                      >
                        Edit
                      </button>
                      {p.name !== currentUser.name && (
                        <button
                          onClick={() => remove(p.id)}
                          style={{
                            background: "none",
                            border: "none",
                            color: T.muted,
                            cursor: "pointer",
                            fontSize: 16,
                            padding: 0,
                          }}
                        >
                          🗑
                        </button>
                      )}
                    </div>
                  </div>
                </div>
              );
            })}

            {/* Everyone else defaults to Crew */}
            <div
              style={{
                fontSize: 12,
                fontWeight: 700,
                color: T.muted,
                textTransform: "uppercase",
                letterSpacing: "1px",
                margin: "20px 0 10px",
              }}
            >
              Unconfigured (Default: Field Crew)
            </div>
            {NAMES.filter((n) => !profileMap[n])
              .slice(0, 15)
              .map((n) => (
                <div
                  key={n}
                  style={{
                    ...cardS,
                    marginBottom: 6,
                    display: "flex",
                    alignItems: "center",
                    justifyContent: "space-between",
                    opacity: 0.5,
                  }}
                >
                  <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
                    <div style={{ width: 10, height: 10, borderRadius: "50%", background: T.green }} />
                    <span style={{ fontSize: 13 }}>{n}</span>
                    <span style={pill(T.green)}>Field Crew</span>
                  </div>
                  <button
                    onClick={() => {
                      setF({ ...blank, name: n });
                      setMode("edit");
                    }}
                    style={{
                      background: T.surface,
                      border: `1px solid ${T.border}`,
                      borderRadius: 8,
                      padding: "5px 10px",
                      color: T.sub,
                      fontSize: 12,
                      cursor: "pointer",
                      fontFamily: "inherit",
                    }}
                  >
                    Set Role
                  </button>
                </div>
              ))}
            {NAMES.filter((n) => !profileMap[n]).length > 15 && (
              <div style={{ fontSize: 12, color: T.muted, textAlign: "center", padding: "8px 0" }}>
                + {NAMES.filter((n) => !profileMap[n]).length - 15} more (tap + Add to configure)
              </div>
            )}
          </>
        )}
      </div>
    </div>
  );
}

/* ── CREW DIRECTORY SCREEN ──────────────────────────────────── */

function NotificationsPanel() {
  const [notifs, setNotifs] = useState([]);
  const [nl, setNl] = useState(true);
  async function loadN() {
    setNl(true);
    try {
      setNotifs((await API.notifications.list()) || []);
    } catch {}
    setNl(false);
  }
  useEffect(() => {
    loadN();
  }, []);
  const typeIcon = { report_submitted: "📋", report_flagged: "🚩", report_approved: "✅" };
  return (
    <div style={{ padding: "14px 16px 80px" }}>
      {unread > 0 && (
        <button
          onClick={async () => {
            await API.notifications.markAllRead();
            setUnread(0);
            await loadN();
          }}
          style={{ ...ghostBtn, width: "100%", textAlign: "center", marginBottom: 14 }}
        >
          Mark all read
        </button>
      )}
      {nl && <Spinner />}
      {!nl && notifs.length === 0 && (
        <div style={{ textAlign: "center", padding: "40px 0", color: T.muted }}>
          <div style={{ fontSize: 36, marginBottom: 8 }}>🔔</div>
          <div>No notifications yet.</div>
        </div>
      )}
      {!nl &&
        notifs.map((n) => (
          <div
            key={n.id}
            onClick={async () => {
              if (!n.read) {
                await API.notifications.markRead(n.id);
                setUnread((u) => Math.max(0, u - 1));
                await loadN();
              }
            }}
            style={{
              ...cardS,
              marginBottom: 8,
              borderLeft: `3px solid ${n.read ? T.border : T.orange}`,
              opacity: n.read ? 0.6 : 1,
              cursor: n.read ? "default" : "pointer",
            }}
          >
            <div style={{ display: "flex", gap: 10, alignItems: "flex-start" }}>
              <span style={{ fontSize: 18, flexShrink: 0 }}>{typeIcon[n.type] || "📬"}</span>
              <div style={{ flex: 1 }}>
                <div style={{ fontSize: 14, fontWeight: 700 }}>{n.title}</div>
                {n.body && <div style={{ fontSize: 12, color: T.sub, marginTop: 2 }}>{n.body}</div>}
                <div style={{ fontSize: 11, color: T.muted, marginTop: 4 }}>
                  {n.created_at ? new Date(n.created_at).toLocaleString() : ""}
                </div>
              </div>
              {!n.read && (
                <div
                  style={{
                    width: 8,
                    height: 8,
                    borderRadius: "50%",
                    background: T.orange,
                    flexShrink: 0,
                    marginTop: 4,
                  }}
                />
              )}
            </div>
          </div>
        ))}
    </div>
  );
}

/* ── TIME CARDS SCREEN ───────────────────────────────────────── */
function TimeCardsScreen({ user, projects, onBack }) {
  const [cards, setCards] = useState([]);
  const [loading, setLoading] = useState(true);
  const [err, setErr] = useState("");
  useEffect(() => {
    (async () => {
      setLoading(true);
      try {
        const r = await API.timeCards.all();
        setCards(Array.isArray(r) ? r : []);
      } catch (e) {
        setErr(e.message);
      }
      setLoading(false);
    })();
  }, []);
  async function remove(id) {
    try {
      await API.timeCards.remove(id);
      setCards((c) => c.filter((x) => x.id !== id));
    } catch (e) {
      setErr(e.message);
    }
  }
  const byWorker = {};
  const wkStart = getWeekStart();
  const weekCards = cards.filter((c) => c.date >= wkStart);
  weekCards.forEach((c) => {
    const n = c.worker_name || "?";
    if (!byWorker[n]) byWorker[n] = { name: n, total: 0, ot: 0 };
    const reg = parseFloat(c.reg_hours) || 0;
    const ot = parseFloat(c.ot_hours) || 0;
    const trav = parseFloat(c.travel_hours) || 0;
    const tot = c.total_hours ? parseFloat(c.total_hours) : reg + ot + trav;
    byWorker[n].total += tot;
    byWorker[n].ot += ot;
  });
  const workerRows = Object.values(byWorker).sort((a, b) => b.total - a.total);
  const todayCards = cards.filter((c) => c.date === today());
  const recentCards = cards.filter((c) => c.date !== today()).slice(0, 30);
  const fmt = (n) => Number(n || 0).toFixed(1);
  return (
    <div style={{ background: T.bg, minHeight: "100vh", fontFamily: "inherit" }}>
      <TopBar title="⏱️ Time Cards" onBack={onBack} />
      <div style={{ padding: "12px 16px 80px" }}>
        <ErrBanner msg={err} onDismiss={() => setErr("")} />
        <div
          style={{
            background: T.greenLow,
            border: `1px solid ${T.green}40`,
            borderRadius: 12,
            padding: "10px 14px",
            marginBottom: 14,
            fontSize: 12,
            color: T.green,
            lineHeight: 1.5,
          }}
        >
          <strong>⚡ Auto-filled from daily reports</strong> — hours are added automatically when a foreman submits a
          daily report.
        </div>
        {workerRows.length > 0 && (
          <div style={{ ...cardS, marginBottom: 14 }}>
            <div
              style={{
                fontSize: 11,
                fontWeight: 700,
                color: T.muted,
                textTransform: "uppercase",
                letterSpacing: "1px",
                marginBottom: 10,
              }}
            >
              THIS WEEK
            </div>
            {workerRows.map((w) => (
              <div
                key={w.name}
                style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 6 }}
              >
                <span style={{ fontSize: 13, color: T.text }}>{w.name}</span>
                <div style={{ display: "flex", gap: 10, fontSize: 12 }}>
                  {w.ot > 0 && <span style={{ color: T.yellow }}>{fmt(w.ot)}OT</span>}
                  <span style={{ color: T.green, fontWeight: 800 }}>{fmt(w.total)}h</span>
                </div>
              </div>
            ))}
          </div>
        )}
        {loading && <Spinner />}
        {!loading && todayCards.length === 0 && recentCards.length === 0 && (
          <div style={{ textAlign: "center", padding: "40px 0", color: T.muted }}>
            <div style={{ fontSize: 32 }}>⏱️</div>
            <div style={{ marginTop: 8 }}>No time cards yet</div>
          </div>
        )}
        {todayCards.length > 0 && (
          <div
            style={{
              fontSize: 11,
              fontWeight: 700,
              color: T.muted,
              textTransform: "uppercase",
              letterSpacing: "1px",
              marginBottom: 8,
            }}
          >
            TODAY
          </div>
        )}
        {[...todayCards, ...recentCards].map((c) => {
          const reg = parseFloat(c.reg_hours) || 0;
          const ot = parseFloat(c.ot_hours) || 0;
          const trav = parseFloat(c.travel_hours) || 0;
          const tot = c.total_hours ? parseFloat(c.total_hours) : reg + ot + trav;
          const proj = projects.find((p) => p.id === c.project_id);
          return (
            <div
              key={c.id}
              style={{
                ...cardS,
                marginBottom: 8,
                display: "flex",
                justifyContent: "space-between",
                alignItems: "flex-start",
              }}
            >
              <div style={{ flex: 1 }}>
                <div style={{ fontSize: 13, fontWeight: 700, color: T.orange }}>{c.worker_name}</div>
                <div style={{ fontSize: 11, color: T.muted }}>
                  {c.date}
                  {proj ? ` · ${proj.name}` : ""}
                </div>
                {c.notes && (
                  <div style={{ fontSize: 11, color: T.muted, fontStyle: "italic", marginTop: 2 }}>{c.notes}</div>
                )}
              </div>
              <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                <div style={{ textAlign: "right" }}>
                  <div style={{ fontSize: 16, fontWeight: 800, color: T.green }}>{fmt(tot)}h</div>
                  {ot > 0 && <div style={{ fontSize: 10, color: T.yellow }}>{fmt(ot)} OT</div>}
                  {trav > 0 && <div style={{ fontSize: 10, color: T.blue }}>{fmt(trav)} travel</div>}
                </div>
                <button
                  onClick={() => remove(c.id)}
                  style={{ background: "none", border: "none", color: T.red, cursor: "pointer", fontSize: 14 }}
                >
                  🗑
                </button>
              </div>
            </div>
          );
        })}
      </div>
    </div>
  );
}

/* ── ESTIMATING SCREEN (Maintenance) ────────────────────────── */
function EstimatingScreen({ user, onBack }) {
  return (
    <div style={{ background: T.bg, minHeight: "100vh", fontFamily: "inherit" }}>
      <TopBar title="📊 Estimating" onBack={onBack} />
      <div
        style={{
          display: "flex",
          flexDirection: "column",
          alignItems: "center",
          justifyContent: "center",
          padding: "60px 24px",
          textAlign: "center",
        }}
      >
        <div style={{ fontSize: 64, marginBottom: 20 }}>🚧</div>
        <div style={{ fontSize: 22, fontWeight: 900, color: T.text, marginBottom: 8 }}>Under Maintenance</div>
        <div style={{ fontSize: 14, color: T.muted, lineHeight: 1.7, maxWidth: 320 }}>
          The Estimating platform is currently being updated. Check back soon.
        </div>
      </div>
    </div>
  );
}

/* ── EMPLOYEE HISTORY / FINANCIALS STUBS ─────────────────────── */
function EmployeeHistory({ user, projects, onBack }) {
  return <TimeCardsScreen user={user} projects={projects} onBack={onBack} />;
}
function FinancialsScreen({ user, projects, onBack, onErr }) {
  return (
    <div style={{ background: T.bg, minHeight: "100vh" }}>
      <TopBar title="💵 Financials" onBack={onBack} />
      <div style={{ padding: 20, color: T.muted, textAlign: "center", marginTop: 40 }}>Financials coming soon</div>
    </div>
  );
}
function TimecardReportScreen({ user, projects, onBack }) {
  return <TimeCardsScreen user={user} projects={projects} onBack={onBack} />;
}
function EmailSummaryModal({ user, projects, onClose }) {
  return (
    <div
      style={{
        position: "fixed",
        inset: 0,
        background: "rgba(0,0,0,0.8)",
        display: "flex",
        alignItems: "center",
        justifyContent: "center",
        zIndex: 200,
      }}
    >
      <div style={{ background: T.card, borderRadius: 16, padding: 24, maxWidth: 400, width: "90%" }}>
        <div style={{ fontSize: 16, fontWeight: 800, color: T.text, marginBottom: 12 }}>📧 Email Summary</div>
        <button onClick={onClose} style={{ ...primBtn, borderRadius: 12 }}>
          Close
        </button>
      </div>
    </div>
  );
}

function ProjectDetail({ project: initP, user, onBack, onProjectUpdated }) {
  const [project, setProject] = useState(initP);
  const [reports, setReports] = useState([]);
  const [safety, setSafety] = useState([]);
  const [photos, setPhotos] = useState([]);
  const [weather, setWeather] = useState([]);
  const [tab, setTab] = useState("reports");
  const [loading, setLoading] = useState(true);
  const [err, setErr] = useState("");
  const [screen, setScreen] = useState("detail");
  const [activeReport, setActiveReport] = useState(null);
  const [editProject, setEditProject] = useState(false);
  const visibleTabs = PTABS.filter((t) => !t.perm || can(user, t.perm));
  const divMeta = DIV_META[project.division] || { color: T.orange, icon: "🏗️" };

  async function load(silent = false) {
    if (!silent) setLoading(true);
    try {
      const [reps, saf, phs, wx] = await Promise.all([
        API.reports.forProject(project.id),
        API.safety.forProject(project.id),
        API.photos.forProject(project.id),
        API.weather.forProject(project.id),
      ]);
      setReports(reps || []);
      setSafety(saf || []);
      setPhotos(phs || []);
      setWeather(wx || []);
    } catch (e) {
      setErr(e.message);
    }
    if (!silent) setLoading(false);
  }
  useEffect(() => {
    load();
    const firstTab = visibleTabs[0]?.id || "info";
    setTab(firstTab);
  }, [project.id]);

  async function saveReport(d) {
    try {
      const { rental_equipment, ...dbData } = d;
      let saved = false;
      try {
        await API.reports.create({ ...dbData, rental_equipment, project_id: project.id });
        saved = true;
      } catch (colErr) {
        if (colErr.message && colErr.message.includes("rental_equipment")) {
          await API.reports.create({ ...dbData, project_id: project.id });
          saved = true;
        } else {
          throw colErr;
        }
      }
      if (!saved) throw new Error("Save failed");
      await load(true);
      setScreen("detail");
    } catch (e) {
      setErr(e.message);
    }
  }
  async function deleteReport(id) {
    try {
      await API.reports.remove(id);
      setActiveReport(null);
      await load(true);
      setScreen("detail");
    } catch (e) {
      setErr(e.message);
    }
  }
  async function approveReport(id) {
    try {
      await API.reports.update(id, {
        status: "approved",
        approved_by: user.name,
        approved_at: new Date().toISOString(),
      });
      setActiveReport((r) => ({ ...r, status: "approved" }));
      await load(true);
    } catch (e) {
      setErr(e.message);
    }
  }
  async function flagReport(id, pm_notes) {
    try {
      await API.reports.update(id, { status: "flagged", pm_notes });
      setActiveReport((r) => ({ ...r, status: "flagged", pm_notes }));
      await notify("report_flagged", "Report Flagged", pm_notes, { project_id: project.id, report_id: id });
      await load(true);
    } catch (e) {
      setErr(e.message);
    }
  }
  async function updateProject(data) {
    try {
      const [u] = await API.projects.update(project.id, data);
      setProject(u);
      onProjectUpdated(u);
      setEditProject(false);
    } catch (e) {
      setErr(e.message);
    }
  }
  async function archiveProject() {
    if (!window.confirm(project.status === "active" ? "Archive this job?" : "Restore?")) return;
    await updateProject({ status: project.status === "active" ? "archived" : "active" });
    onBack();
  }
  async function deleteProject() {
    if (!window.confirm("Permanently delete this job and ALL its data? This cannot be undone.")) return;
    if (!window.confirm("Are you sure? All reports, photos, time cards and safety logs will be deleted.")) return;
    try {
      await API.projects.remove(project.id);
      onBack();
    } catch (e) {
      setErr(e.message);
    }
  }

  const tot = reports.reduce(
    (s, r) => {
      const t = reportTotals(r);
      return { l: s.l + t.labor, e: s.e + t.equip, m: s.m + t.mats, g: s.g + t.grand };
    },
    { l: 0, e: 0, m: 0, g: 0 },
  );

  if (screen === "newReport" && can(user, "submit_report"))
    return <DailyReportForm user={user} project={project} onSave={saveReport} onCancel={() => setScreen("detail")} />;
  if (screen === "reportDetail" && activeReport)
    return (
      <ReportDetail
        report={activeReport}
        project={project}
        user={user}
        onBack={() => setScreen("detail")}
        onDelete={deleteReport}
        onApprove={approveReport}
        onFlag={flagReport}
      />
    );
  if (editProject && can(user, "edit_job"))
    return (
      <ProjectForm
        initial={project}
        onSave={updateProject}
        onCancel={() => setEditProject(false)}
        defaultDivision={project.division}
      />
    );

  return (
    <div style={{ background: T.bg, minHeight: "100vh", fontFamily: "inherit" }}>
      <div
        style={{
          background: T.surface,
          borderBottom: `1px solid ${T.border}`,
          padding: "14px 16px",
          position: "sticky",
          top: 0,
          zIndex: 50,
        }}
      >
        <button
          onClick={onBack}
          style={{
            background: "none",
            border: "none",
            color: T.sub,
            fontSize: 13,
            cursor: "pointer",
            marginBottom: 8,
            padding: 0,
            fontFamily: "inherit",
          }}
        >
          ← {project.division}
        </button>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 12 }}>
          <div style={{ flex: 1, minWidth: 0, paddingRight: 10 }}>
            <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 2 }}>
              <span style={{ fontSize: 16 }}>{divMeta.icon}</span>
              <div style={{ fontSize: 19, fontWeight: 900, color: T.text, letterSpacing: "-0.4px", lineHeight: 1.2 }}>
                {project.name}
              </div>
            </div>
            <div style={{ fontSize: 12, color: T.sub }}>
              {[project.client, project.location].filter(Boolean).join(" · ") || "No details"}
            </div>
          </div>
          <span style={pill(project.status === "active" ? T.green : T.muted)}>{project.status}</span>
        </div>
        <StatBar
          items={[
            { label: "Reports", val: reports.length, color: divMeta.color },
            {
              label: "Labor",
              val: "$" + (tot.l >= 1000 ? (tot.l / 1000).toFixed(1) + "k" : fmt(tot.l)),
              color: T.green,
            },
            {
              label: "Equip",
              val: "$" + (tot.e >= 1000 ? (tot.e / 1000).toFixed(1) + "k" : fmt(tot.e)),
              color: T.yellow,
            },
            {
              label: "Total",
              val: "$" + (tot.g >= 1000 ? (tot.g / 1000).toFixed(1) + "k" : fmt(tot.g)),
              color: T.blue,
            },
          ]}
        />
        <div
          style={{
            display: "flex",
            gap: 4,
            marginTop: 12,
            overflowX: "auto",
            paddingBottom: 2,
            WebkitOverflowScrolling: "touch",
          }}
        >
          {visibleTabs.map((t) => (
            <button
              key={t.id}
              onClick={() => setTab(t.id)}
              style={{
                flexShrink: 0,
                background: tab === t.id ? divMeta.color : "transparent",
                border: tab === t.id ? "none" : `1px solid ${T.border}`,
                borderRadius: 10,
                padding: "8px 10px",
                fontSize: 11,
                fontWeight: tab === t.id ? 800 : 500,
                cursor: "pointer",
                color: tab === t.id ? "#16181D" : T.sub,
                fontFamily: "inherit",
                whiteSpace: "nowrap",
              }}
            >
              {t.icon} {t.label}
            </button>
          ))}
        </div>
      </div>
      <div style={{ padding: "14px 16px 80px" }}>
        <ErrBanner msg={err} onDismiss={() => setErr("")} />
        {loading && <Spinner />}
        {!loading && tab === "reports" && (
          <div>
            {can(user, "submit_report") && (
              <button
                onClick={() => setScreen("newReport")}
                style={{
                  ...primBtn,
                  marginBottom: 16,
                  borderRadius: 14,
                  padding: "18px",
                  fontSize: 17,
                  background: divMeta.color,
                }}
              >
                📋 + New Daily Report
              </button>
            )}
            {reports.length === 0 && (
              <div style={{ textAlign: "center", padding: "32px 0", color: T.muted }}>
                <div style={{ fontSize: 36, marginBottom: 8 }}>📋</div>
                <div style={{ fontSize: 15, fontWeight: 600, color: T.sub, marginBottom: 4 }}>No reports yet</div>
              </div>
            )}
            {reports.map((r) => {
              const t = reportTotals(r);
              const sc = { submitted: T.yellow, approved: T.green, flagged: T.red }[r.status || "submitted"] || T.muted;
              return (
                <div
                  key={r.id}
                  onClick={() => {
                    setActiveReport(r);
                    setScreen("reportDetail");
                  }}
                  style={{
                    ...cardS,
                    marginBottom: 9,
                    cursor: "pointer",
                    borderLeft: `3px solid ${sc}`,
                    display: "flex",
                    justifyContent: "space-between",
                    alignItems: "center",
                  }}
                >
                  <div>
                    <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                      <div style={{ fontSize: 15, fontWeight: 700 }}>{fmtShort(r.date)}</div>
                      <span style={pill(sc)}>{(r.status || "submitted").toUpperCase()}</span>
                    </div>
                    <div style={{ fontSize: 11, color: T.muted, marginTop: 4, display: "flex", gap: 8 }}>
                      {(r.labor || []).length > 0 && <span>👷 {r.labor.length}</span>}
                      {(r.equipment || []).length > 0 && <span>🚜 {r.equipment.length}</span>}
                      {r.submitted_by && <span>by {r.submitted_by}</span>}
                    </div>
                  </div>
                  <div style={{ textAlign: "right" }}>
                    <div style={{ fontSize: 17, fontWeight: 900, color: T.green }}>${fmt(t.grand)}</div>
                    <div style={{ fontSize: 9, color: T.muted }}>TOTAL</div>
                  </div>
                </div>
              );
            })}
          </div>
        )}
        {!loading && tab === "time" && can(user, "time_card") && (
          <TimeCardsTab projectId={project.id} user={user} onErr={setErr} />
        )}
        {!loading && tab === "crew" && can(user, "crew_equip") && (
          <CrewEquipTab projectId={project.id} user={user} onErr={setErr} />
        )}
        {!loading && tab === "subs" && can(user, "subs") && (
          <SubsTab projectId={project.id} user={user} onErr={setErr} />
        )}
        {!loading && tab === "safety" && can(user, "safety") && (
          <SafetyTab projectId={project.id} safety={safety} user={user} onRefresh={() => load(true)} onErr={setErr} />
        )}
        {!loading && tab === "co" && <ChangeOrdersTab project={project} user={user} onErr={setErr} />}
        {!loading && tab === "rfi" && <RFIsTab project={project} user={user} onErr={setErr} />}
        {!loading && tab === "docs" && can(user, "docs") && (
          <DocsTab projectId={project.id} user={user} onErr={setErr} />
        )}
        {!loading && tab === "schedule" && can(user, "schedule") && (
          <ScheduleTab projectId={project.id} user={user} onErr={setErr} />
        )}
        {!loading && tab === "photos" && can(user, "photos") && (
          <PhotosTab projectId={project.id} photos={photos} onRefresh={() => load(true)} onErr={setErr} />
        )}
        {!loading && tab === "weather" && can(user, "weather") && (
          <WeatherTab
            projectId={project.id}
            project={project}
            weather={weather}
            onRefresh={() => load(true)}
            onErr={setErr}
          />
        )}
        {!loading && tab === "info" && (
          <InfoTab
            project={project}
            user={user}
            onEdit={() => setEditProject(true)}
            onArchive={archiveProject}
            onDelete={deleteProject}
          />
        )}
      </div>
    </div>
  );
}

/* ── PM DASHBOARD ───────────────────────────────────────────── */

/* ── CHANGE ORDERS TAB ───────────────────────────────────────── */
function ChangeOrdersTab({ project, user, onErr }) {
  const canEdit = user.role === "admin" || user.role === "pm";
  const [cos, setCos] = useState([]);
  const [loading, setLoading] = useState(true);
  const [showForm, setShowForm] = useState(false);
  const [editing, setEditing] = useState(null);
  const [saving, setSaving] = useState(false);
  const [shareCo, setShareCo] = useState(null);
  const [coCopied, setCoCopied] = useState(false);
  const appUrl = window.location.origin + window.location.pathname;
  const blank = {
    co_number: "",
    description: "",
    date_submitted: today(),
    amount: "",
    status: "Pending",
    submitted_by: user.name,
    approved_by: "",
    approved_date: "",
    notes: "",
    client_signature: null,
    client_signed_by: "",
  };
  const [f, setF] = useState(blank);
  const set = (k, v) => setF((x) => ({ ...x, [k]: v }));
  const statusColor = { Pending: T.yellow, Approved: T.green, Rejected: T.red };

  useEffect(() => {
    load();
  }, [project.id]);
  async function load() {
    setLoading(true);
    try {
      const r = await API.changeOrders.forProject(project.id);
      setCos(Array.isArray(r) ? r : []);
    } catch (e) {
      onErr(e.message);
    }
    setLoading(false);
  }

  async function save() {
    setSaving(true);
    try {
      const toDate = (v) => (v && v.trim() && v !== "Invalid Date" ? v : null);
      const payload = {
        ...f,
        project_id: project.id,
        amount: parseFloat(f.amount) || 0,
        date_submitted: toDate(f.date_submitted),
        approved_date: toDate(f.approved_date),
      };
      if (editing) {
        await API.changeOrders.update(editing, payload);
      } else {
        await API.changeOrders.create(payload);
      }
      setShowForm(false);
      setEditing(null);
      setF(blank);
      await load();
    } catch (e) {
      onErr(e.message);
    }
    setSaving(false);
  }

  async function remove(id) {
    if (!window.confirm("Delete this change order?")) return;
    try {
      await API.changeOrders.remove(id);
      await load();
    } catch (e) {
      onErr(e.message);
    }
  }

  function copyCoLink(co) {
    const link = `${appUrl}?co=${co.id}`;
    navigator.clipboard
      .writeText(link)
      .then(() => {
        setCoCopied(true);
        setTimeout(() => setCoCopied(false), 3000);
      })
      .catch(() => {
        const el = document.createElement("textarea");
        el.value = link;
        document.body.appendChild(el);
        el.select();
        document.execCommand("copy");
        document.body.removeChild(el);
        setCoCopied(true);
        setTimeout(() => setCoCopied(false), 3000);
      });
  }

  function emailCO(co) {
    const link = `${appUrl}?co=${co.id}`;
    const subj = `Change Order ${co.co_number} — ${project.name} — Signature Required`;
    const ln = "%0D%0A";
    const body = [
      `Please review and sign the following Change Order:`,
      ln,
      ln,
      `━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━`,
      ln,
      `  📋  CHANGE ORDER — ${co.co_number}`,
      ln,
      `━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━`,
      ln,
      ln,
      `  Project: ${project.name}`,
      ln,
      `  Amount: $${Number(co.amount || 0).toLocaleString("en-US", { minimumFractionDigits: 2 })}`,
      ln,
      ln,
      `  DESCRIPTION:`,
      ln,
      `  ${co.description || ""}`,
      ln,
      ln,
      `  ➤  ${link}`,
      ln,
      ln,
      `(No login required — open in any browser, sign, and click Approve)`,
      ln,
      ln,
      `Thank you,`,
      ln,
      `${co.submitted_by || "AIME Field Operations"}`,
      ln,
      `Atlantic Industrial Mechanical & Environmental Inc.`,
      ln,
    ]
      .filter(Boolean)
      .join("");
    if (navigator.clipboard) navigator.clipboard.writeText(link).catch(() => {});
    window.location.href = `mailto:?subject=${encodeURIComponent(subj)}&body=${body}`;
  }

  function printCO(co) {
    const contractVal = parseFloat(project.contract_value) || 0;
    const approvedCOs = cos.filter((c) => c.status === "Approved").reduce((s, c) => s + (parseFloat(c.amount) || 0), 0);
    const fmt = (n) => "$" + Number(n || 0).toLocaleString("en-US", { minimumFractionDigits: 2 });
    const html = `<!DOCTYPE html><html><head><meta charset="utf-8"><title>CO ${co.co_number}</title>
<style>@page{size:letter portrait;margin:0.6in;}*{box-sizing:border-box;margin:0;padding:0;font-family:Arial,sans-serif;}body{font-size:10pt;color:#000;}
.lh{display:flex;justify-content:space-between;align-items:flex-start;padding-bottom:14px;border-bottom:3px solid #1f3864;margin-bottom:20px;}
.co{font-size:20pt;font-weight:900;color:#1f3864;}.co-sub{font-size:9pt;color:#555;margin-top:4px;}
.doc-title{text-align:right;}.doc-title h1{font-size:20pt;font-weight:900;color:#1f3864;}
.badge{display:inline-block;padding:4px 14px;border-radius:20px;font-weight:700;font-size:10pt;margin-top:6px;background:${co.status === "Approved" ? "#dcfce7" : co.status === "Rejected" ? "#fee2e2" : "#fef9c3"};color:${co.status === "Approved" ? "#166534" : co.status === "Rejected" ? "#991b1b" : "#713f12"};}
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
<div><div class="fl">Project</div><div class="fv">${project.name || "—"}</div></div>
<div><div class="fl">Client</div><div class="fv">${project.client || "—"}</div></div>
<div><div class="fl">Division</div><div class="fv">${project.division || "—"}</div></div>
<div><div class="fl">Date</div><div class="fv">${co.date_submitted || "—"}</div></div>
<div><div class="fl">Submitted By</div><div class="fv">${co.submitted_by || "—"}</div></div>
<div><div class="fl">AFE / PO</div><div class="fv">${project.afe || "—"}</div></div></div>
<div class="amt-box"><div style="font-size:9pt;text-transform:uppercase;letter-spacing:1px;opacity:0.8">Change Order Amount</div><div style="font-size:26pt;font-weight:900;margin-top:4px">${fmt(co.amount)}</div></div>
<div><div class="fl" style="margin-bottom:6px">Description of Change</div><div class="desc-box">${co.description || "—"}</div></div>
${co.notes ? `<div><div class="fl" style="margin-bottom:6px">Notes</div><div class="desc-box">${co.notes}</div></div>` : ""}
${co.client_signature ? `<div style="background:#fff;border:1px solid #86efac;border-radius:8px;padding:12px;margin-bottom:12px"><div class="fl" style="color:#166534;margin-bottom:6px">Client Signature — ${co.client_signed_by || ""}</div><img src="${co.client_signature}" style="max-height:80px;max-width:300px;object-fit:contain"/></div>` : ""}
<div class="sig-grid">
<div class="sig-box"><div style="height:48px"></div><div class="sig-label">Authorized by (AIME)</div><div style="font-size:10pt;font-weight:700;margin-top:4px">${co.submitted_by || ""}</div><div style="font-size:9pt;color:#555;margin-top:4px">Date: ______________</div></div>
<div class="sig-box">${co.client_signature ? `<img src="${co.client_signature}" style="max-height:60px;max-width:200px;object-fit:contain;display:block;margin-bottom:4px"/>` : `<div style="height:60px;border-bottom:1px solid #000;margin-bottom:4px"></div>`}<div class="sig-label">Accepted by (Client)</div><div style="font-size:10pt;font-weight:700;margin-top:4px">${co.client_signed_by || ""}</div><div style="font-size:9pt;color:#555;margin-top:4px">Date: ${co.approved_date || "______________"}</div></div></div>
<div class="footer"><span>AIME Field Pro · CO ${co.co_number} · ${project.name}</span><span>Generated: ${new Date().toLocaleString()}</span></div>
</body></html>`;
    const win = window.open("", "_blank", "width=900,height=750");
    win.document.write(html);
    win.document.close();
    win.focus();
    setTimeout(() => win.print(), 400);
  }

  const contractVal = parseFloat(project.contract_value) || 0;
  const approvedCOs = cos.filter((c) => c.status === "Approved").reduce((s, c) => s + (parseFloat(c.amount) || 0), 0);
  const pendingCOs = cos.filter((c) => c.status === "Pending").reduce((s, c) => s + (parseFloat(c.amount) || 0), 0);
  const revisedContract = contractVal + approvedCOs;
  const fmt = (n) => "$" + Number(n || 0).toLocaleString("en-US", { minimumFractionDigits: 2 });

  if (showForm || editing)
    return (
      <div style={{ padding: "14px 16px 80px" }}>
        <button
          onClick={() => {
            setShowForm(false);
            setEditing(null);
            setF(blank);
          }}
          style={{ ...ghostBtn, marginBottom: 14 }}
        >
          ← Back
        </button>
        <div style={{ fontSize: 17, fontWeight: 800, marginBottom: 16, color: T.text }}>
          {editing ? "Edit" : "New"} Change Order
        </div>
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10, marginBottom: 10 }}>
          <div>
            <label style={lbl}>CO Number *</label>
            <input
              value={f.co_number}
              onChange={(e) => set("co_number", e.target.value)}
              placeholder="CO-001"
              style={inp}
            />
          </div>
          <div>
            <label style={lbl}>Date</label>
            <input
              type="date"
              value={f.date_submitted || ""}
              onChange={(e) => set("date_submitted", e.target.value)}
              style={inp}
            />
          </div>
        </div>
        <div style={{ marginBottom: 10 }}>
          <label style={lbl}>Description *</label>
          <textarea
            rows={3}
            value={f.description}
            onChange={(e) => set("description", e.target.value)}
            placeholder="Describe the change in scope..."
            style={{ ...inp, resize: "vertical" }}
          />
        </div>
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10, marginBottom: 10 }}>
          <div>
            <label style={lbl}>Amount ($)</label>
            <input
              type="number"
              value={f.amount}
              onChange={(e) => set("amount", e.target.value)}
              placeholder="0.00"
              style={inp}
            />
          </div>
          <div>
            <label style={lbl}>Status</label>
            <select value={f.status} onChange={(e) => set("status", e.target.value)} style={inpSel}>
              {["Pending", "Approved", "Rejected"].map((s) => (
                <option key={s}>{s}</option>
              ))}
            </select>
          </div>
        </div>
        {f.status === "Approved" && (
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10, marginBottom: 10 }}>
            <div>
              <label style={lbl}>Approved By</label>
              <input value={f.approved_by} onChange={(e) => set("approved_by", e.target.value)} style={inp} />
            </div>
            <div>
              <label style={lbl}>Approval Date</label>
              <input
                type="date"
                value={f.approved_date || ""}
                onChange={(e) => set("approved_date", e.target.value)}
                style={inp}
              />
            </div>
          </div>
        )}
        <div style={{ marginBottom: 14 }}>
          <label style={lbl}>Notes</label>
          <textarea
            rows={2}
            value={f.notes}
            onChange={(e) => set("notes", e.target.value)}
            style={{ ...inp, resize: "vertical" }}
          />
        </div>
        <button
          onClick={save}
          disabled={!f.co_number.trim() || saving}
          style={{ ...primBtn, opacity: f.co_number.trim() && !saving ? 1 : 0.5, borderRadius: 14 }}
        >
          {saving ? "Saving…" : "Save Change Order"}
        </button>
      </div>
    );

  return (
    <div style={{ padding: "14px 16px 80px" }}>
      {/* Share Modal */}
      {shareCo && (
        <div
          onClick={() => {
            setShareCo(null);
            setCoCopied(false);
          }}
          style={{
            position: "fixed",
            inset: 0,
            zIndex: 300,
            background: "rgba(0,0,0,0.85)",
            display: "flex",
            alignItems: "flex-end",
            justifyContent: "center",
          }}
        >
          <div
            onClick={(e) => e.stopPropagation()}
            style={{
              background: T.card,
              borderRadius: "20px 20px 0 0",
              padding: "20px 20px 40px",
              width: "100%",
              maxWidth: 480,
            }}
          >
            <div style={{ fontSize: 17, fontWeight: 900, marginBottom: 4, color: T.text }}>
              📤 Send CO {shareCo.co_number}
            </div>
            <div style={{ fontSize: 12, color: T.muted, marginBottom: 16 }}>
              Client opens the link, reviews the CO, signs it, and it saves directly to your app.
            </div>
            <div
              style={{
                background: T.greenLow,
                border: `1px solid ${T.green}40`,
                borderRadius: 10,
                padding: "10px 14px",
                marginBottom: 14,
                display: "flex",
                justifyContent: "space-between",
                alignItems: "center",
              }}
            >
              <span style={{ fontSize: 12, color: T.sub }}>{shareCo.description?.substring(0, 50)}</span>
              <span style={{ fontSize: 16, fontWeight: 800, color: T.green }}>
                ${Number(shareCo.amount || 0).toLocaleString("en-US", { minimumFractionDigits: 2 })}
              </span>
            </div>
            <div
              style={{
                background: T.surface,
                border: `1px solid ${T.border}`,
                borderRadius: 12,
                padding: "10px 14px",
                marginBottom: 14,
                wordBreak: "break-all",
                fontSize: 12,
                color: T.sub,
              }}
            >{`${appUrl}?co=${shareCo.id}`}</div>
            <button
              onClick={() => copyCoLink(shareCo)}
              style={{
                ...primBtn,
                borderRadius: 14,
                marginBottom: 10,
                background: coCopied ? T.green : T.orange,
                transition: "background 0.2s",
              }}
            >
              {coCopied ? "✅ Link Copied! Paste into your email" : "📋 Copy Link to Clipboard"}
            </button>
            {coCopied && (
              <div
                style={{
                  background: T.greenLow,
                  border: `1px solid ${T.green}40`,
                  borderRadius: 10,
                  padding: "10px 14px",
                  marginBottom: 10,
                  fontSize: 13,
                  color: T.green,
                  textAlign: "center",
                  fontWeight: 600,
                }}
              >
                ✓ Paste this link into an email or text to your client
              </div>
            )}
            <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 10 }}>
              <div style={{ flex: 1, height: 1, background: T.border }} />
              <span style={{ fontSize: 11, color: T.muted }}>OR</span>
              <div style={{ flex: 1, height: 1, background: T.border }} />
            </div>
            <button
              onClick={() => emailCO(shareCo)}
              style={{ ...ghostBtn, width: "100%", textAlign: "center", marginBottom: 10, fontSize: 14 }}
            >
              📧 Open Email Draft with Link
            </button>
            <div style={{ fontSize: 11, color: T.muted, textAlign: "center", marginBottom: 14 }}>
              Note: Copy Link above gives a clickable link.
            </div>
            <button
              onClick={() => {
                setShareCo(null);
                setCoCopied(false);
              }}
              style={{ ...ghostBtn, width: "100%", textAlign: "center" }}
            >
              Done
            </button>
          </div>
        </div>
      )}

      {contractVal > 0 && (
        <div style={{ ...cardS, marginBottom: 14, background: T.blueLow, border: `1px solid ${T.blue}40` }}>
          <div
            style={{
              fontSize: 12,
              fontWeight: 700,
              color: T.blue,
              textTransform: "uppercase",
              letterSpacing: "1px",
              marginBottom: 10,
            }}
          >
            Contract Summary
          </div>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8, marginBottom: 8 }}>
            <div style={{ background: T.card, borderRadius: 10, padding: "8px 10px" }}>
              <div style={{ fontSize: 11, color: T.muted }}>Original Contract</div>
              <div style={{ fontSize: 15, fontWeight: 800, color: T.text }}>{fmt(contractVal)}</div>
            </div>
            <div style={{ background: T.card, borderRadius: 10, padding: "8px 10px" }}>
              <div style={{ fontSize: 11, color: T.muted }}>Approved COs</div>
              <div style={{ fontSize: 15, fontWeight: 800, color: T.green }}>+{fmt(approvedCOs)}</div>
            </div>
          </div>
          {pendingCOs > 0 && (
            <div style={{ background: T.card, borderRadius: 10, padding: "8px 10px", marginBottom: 8 }}>
              <div style={{ fontSize: 11, color: T.muted }}>Pending COs</div>
              <div style={{ fontSize: 15, fontWeight: 800, color: T.yellow }}>+{fmt(pendingCOs)}</div>
            </div>
          )}
          <div
            style={{
              background: T.card,
              borderRadius: 10,
              padding: "10px",
              textAlign: "center",
              border: `1px solid ${T.blue}40`,
            }}
          >
            <div style={{ fontSize: 11, color: T.muted, textTransform: "uppercase", letterSpacing: "1px" }}>
              Revised Contract Value
            </div>
            <div style={{ fontSize: 22, fontWeight: 900, color: T.blue }}>{fmt(revisedContract)}</div>
          </div>
        </div>
      )}

      {canEdit && (
        <button onClick={() => setShowForm(true)} style={{ ...primBtn, borderRadius: 14, marginBottom: 14 }}>
          + New Change Order
        </button>
      )}
      {loading && <Spinner />}
      {!loading && cos.length === 0 && (
        <div style={{ textAlign: "center", padding: "40px 0", color: T.muted }}>
          <div style={{ fontSize: 32, marginBottom: 8 }}>📋</div>
          <div style={{ fontWeight: 700, color: T.sub }}>No Change Orders Yet</div>
        </div>
      )}
      {cos.map((co) => (
        <div
          key={co.id}
          style={{ ...cardS, marginBottom: 10, borderLeft: `3px solid ${statusColor[co.status] || T.muted}` }}
        >
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 8 }}>
            <div>
              <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 2 }}>
                <span style={{ fontSize: 15, fontWeight: 800, color: T.orange }}>{co.co_number}</span>
                <span style={pill(statusColor[co.status] || T.muted)}>{co.status}</span>
              </div>
              <div style={{ fontSize: 12, color: T.muted }}>
                {co.date_submitted} · {co.submitted_by}
              </div>
            </div>
            <div style={{ fontSize: 18, fontWeight: 900, color: co.status === "Rejected" ? T.red : T.green }}>
              {co.status === "Rejected" ? "—" : "+"}
              {fmt(co.amount)}
            </div>
          </div>
          {co.description && (
            <div style={{ fontSize: 13, color: T.text, marginBottom: 8, lineHeight: 1.5 }}>{co.description}</div>
          )}
          {co.status === "Approved" && co.approved_by && (
            <div style={{ fontSize: 11, color: T.green }}>
              ✓ Approved by {co.approved_by}
              {co.approved_date ? " on " + co.approved_date : ""}
            </div>
          )}
          {co.client_signature && (
            <div style={{ background: "#fff", borderRadius: 8, padding: 4, marginTop: 6 }}>
              <div style={{ fontSize: 9, color: "#999", paddingLeft: 4, marginBottom: 2 }}>
                CLIENT SIGNATURE — {co.client_signed_by || ""}
              </div>
              <img
                src={co.client_signature}
                alt="sig"
                style={{ width: "100%", maxHeight: 60, objectFit: "contain", borderRadius: 6 }}
              />
            </div>
          )}
          {co.notes && (
            <div style={{ fontSize: 11, color: T.muted, marginTop: 4, fontStyle: "italic" }}>{co.notes}</div>
          )}
          <div
            style={{
              display: "flex",
              gap: 6,
              marginTop: 10,
              paddingTop: 10,
              borderTop: `1px solid ${T.border}`,
              flexWrap: "wrap",
            }}
          >
            <button
              onClick={() => {
                setShareCo(co);
                setCoCopied(false);
              }}
              style={{
                ...ghostBtn,
                flex: 1,
                textAlign: "center",
                fontSize: 12,
                color: T.orange,
                border: `1px solid ${T.orange}40`,
                fontWeight: 700,
                minWidth: 90,
              }}
            >
              📤 Send Link
            </button>
            <button
              onClick={() => printCO(co)}
              style={{
                ...ghostBtn,
                flex: 1,
                textAlign: "center",
                fontSize: 12,
                color: "#CBD5E1",
                border: "1px solid #27272A",
                minWidth: 70,
              }}
            >
              🖨️ PDF
            </button>
            {canEdit && (
              <>
                <button
                  onClick={() => {
                    setEditing(co.id);
                    setF({
                      ...co,
                      amount: co.amount || "",
                      client_signature: co.client_signature || null,
                      client_signed_by: co.client_signed_by || "",
                    });
                  }}
                  style={{ ...ghostBtn, flex: 1, textAlign: "center", fontSize: 12, minWidth: 50 }}
                >
                  ✏️
                </button>
                <button
                  onClick={() => remove(co.id)}
                  style={{
                    ...ghostBtn,
                    flex: 1,
                    textAlign: "center",
                    fontSize: 12,
                    color: T.red,
                    border: `1px solid ${T.red}40`,
                    minWidth: 40,
                  }}
                >
                  🗑
                </button>
              </>
            )}
          </div>
        </div>
      ))}
    </div>
  );
}

/* ── RFIs TAB ────────────────────────────────────────────────── */
function RFIsTab({ project, user, onErr }) {
  const canEdit = user.role === "admin" || user.role === "pm";
  const [rfis, setRfis] = useState([]);
  const [loading, setLoading] = useState(true);
  const [showForm, setShowForm] = useState(false);
  const [editing, setEditing] = useState(null);
  const [saving, setSaving] = useState(false);
  const [shareRfi, setShareRfi] = useState(null);
  const [copied, setCopied] = useState(false);
  const appUrl = window.location.origin + window.location.pathname;
  const blank = {
    rfi_number: "",
    date_submitted: today(),
    submitted_by: user.name,
    question: "",
    description: "",
    due_date: "",
    response: "",
    responded_by: "",
    response_date: "",
    status: "Open",
    notes: "",
    ball_in_court: "",
    ball_in_court_email: "",
  };
  const [f, setF] = useState(blank);
  const set = (k, v) => setF((x) => ({ ...x, [k]: v }));
  const statusColor = { Open: T.yellow, Answered: T.blue, Closed: T.green, Overdue: T.red };

  useEffect(() => {
    load();
  }, [project.id]);
  async function load() {
    setLoading(true);
    try {
      const r = await API.rfis.forProject(project.id);
      setRfis(Array.isArray(r) ? r : []);
    } catch (e) {
      onErr(e.message);
    }
    setLoading(false);
  }

  async function save() {
    setSaving(true);
    try {
      const toDate = (v) => (v && v.trim() && v !== "Invalid Date" ? v : null);
      const payload = {
        ...f,
        project_id: project.id,
        date_submitted: toDate(f.date_submitted),
        due_date: toDate(f.due_date),
        response_date: toDate(f.response_date),
      };
      if (editing) {
        await API.rfis.update(editing, payload);
      } else {
        await API.rfis.create(payload);
      }
      setShowForm(false);
      setEditing(null);
      setF(blank);
      await load();
    } catch (e) {
      onErr(e.message);
    }
    setSaving(false);
  }

  async function remove(id) {
    if (!window.confirm("Delete this RFI?")) return;
    try {
      await API.rfis.remove(id);
      await load();
    } catch (e) {
      onErr(e.message);
    }
  }

  function copyLink(rfi) {
    const link = `${appUrl}?rfi=${rfi.id}`;
    navigator.clipboard
      .writeText(link)
      .then(() => {
        setCopied(true);
        setTimeout(() => setCopied(false), 3000);
      })
      .catch(() => {
        const el = document.createElement("textarea");
        el.value = link;
        document.body.appendChild(el);
        el.select();
        document.execCommand("copy");
        document.body.removeChild(el);
        setCopied(true);
        setTimeout(() => setCopied(false), 3000);
      });
  }

  function openEmailDraft(rfi) {
    const link = `${appUrl}?rfi=${rfi.id}`;
    const subj = `RFI #${rfi.rfi_number} — ${project.name} — Response Required`;
    const ln = "%0D%0A";
    const sep = "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━";
    const body = [
      `Hi ${rfi.ball_in_court || ""},`,
      ln,
      ln,
      `Please review and respond to the RFI below.`,
      ln,
      ln,
      sep,
      ln,
      `  📋  REQUEST FOR INFORMATION — RFI #${rfi.rfi_number}`,
      ln,
      sep,
      ln,
      ln,
      `  Project:    ${project.name}`,
      ln,
      rfi.due_date ? `  Due By:    ${rfi.due_date}${ln}` : "",
      ln,
      `  QUESTION:`,
      ln,
      `  ${rfi.question}`,
      ln,
      ln,
      sep,
      ln,
      ln,
      `To respond, click or copy this link into your browser:`,
      ln,
      ln,
      `  ${link}`,
      ln,
      ln,
      `(No login required — fill in the form and click Submit)`,
      ln,
      ln,
      sep,
      ln,
      ln,
      `Thank you,`,
      ln,
      `${rfi.submitted_by || "AIME Field Operations"}`,
      ln,
      `Atlantic Industrial Mechanical & Environmental Inc.`,
      ln,
    ]
      .filter(Boolean)
      .join("");
    window.location.href = `mailto:${rfi.ball_in_court_email || ""}?subject=${encodeURIComponent(subj)}&body=${body}`;
  }

  function printRFI(rfi) {
    const isOverdue = rfi.due_date && new Date(rfi.due_date) < new Date() && rfi.status === "Open";
    const effStatus = isOverdue ? "Overdue" : rfi.status;
    const html = `<!DOCTYPE html><html><head><meta charset="utf-8"><title>RFI ${rfi.rfi_number}</title>
<style>@page{size:letter portrait;margin:0.6in;}*{box-sizing:border-box;margin:0;padding:0;font-family:Arial,sans-serif;}body{font-size:10pt;color:#000;}
.action-bar{background:#1f3864;color:#fff;padding:12px 20px;display:flex;justify-content:space-between;align-items:center;gap:12px;position:sticky;top:0;z-index:100;}
.action-bar h2{font-size:12pt;font-weight:700;margin:0;}.btns{display:flex;gap:10px;}
.btn{padding:8px 16px;border-radius:8px;border:none;font-size:11pt;font-weight:700;cursor:pointer;font-family:inherit;}
.btn-print{background:#F97316;color:#000;}.btn-clear{background:transparent;color:#fff;border:1px solid rgba(255,255,255,0.4);}
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
<div><div class="fl">Project</div><div class="fv">${project.name || "—"}</div></div>
<div><div class="fl">Client</div><div class="fv">${project.client || "—"}</div></div>
<div><div class="fl">Division</div><div class="fv">${project.division || "—"}</div></div>
<div><div class="fl">Date Submitted</div><div class="fv">${rfi.date_submitted || "—"}</div></div>
<div><div class="fl">Submitted By</div><div class="fv">${rfi.submitted_by || "—"}</div></div>
<div><div class="fl">Response Due</div><div class="fv">${rfi.due_date || "—"}</div></div></div>
${rfi.ball_in_court ? `<div style="background:#fff7ed;border:1.5px solid #fed7aa;border-radius:8px;padding:12px 16px;margin-bottom:16px;display:flex;justify-content:space-between;align-items:center"><div><div class="fl" style="color:#9a3412">🏀 Ball in Court</div><div style="font-size:13pt;font-weight:900;color:#c2410c;margin-top:2px">${rfi.ball_in_court}</div>${rfi.ball_in_court_email ? `<div style="font-size:9pt;color:#9a3412">${rfi.ball_in_court_email}</div>` : ""}</div></div>` : ""}
<div><div class="fl" style="margin-bottom:6px">Question / Issue</div><div class="q-box">${rfi.question || "—"}</div>${rfi.description ? `<div style="background:#f9fafb;border:1px solid #e5e7eb;border-radius:6px;padding:12px;font-size:10pt;line-height:1.7">${rfi.description}</div>` : ""}</div><br>
${
  rfi.response
    ? `<div class="resp-section"><h2>Response</h2><div style="background:#fff;border:1px solid #86efac;border-radius:8px;padding:12px"><div style="font-size:8pt;font-weight:700;color:#166534;margin-bottom:4px">From ${rfi.responded_by || "—"} · ${rfi.response_date || ""}</div><div>${rfi.response}</div>${rfi.response_signature ? `<img src="${rfi.response_signature}" style="max-height:60px;width:100%;object-fit:contain;display:block;margin-top:8px"/>` : ""}</div></div>`
    : `<div class="resp-section"><h2>Response <span style="font-weight:400;font-size:9pt;text-transform:none">(Please complete and return)</span></h2>
<div class="fill-hint">✏️ Type your response below, then click Print / Save as PDF</div>
<textarea id="resp" class="fillable" placeholder="Enter your response here..." rows="6"></textarea>
<div class="fill-row"><div class="fill-field"><div class="fill-label">Responded By</div><input id="respBy" type="text" class="fillable" placeholder="Your name"/></div>
<div class="fill-field"><div class="fill-label">Response Date</div><input id="respDate" type="date" class="fillable"/></div></div></div>`
}
<div class="sig-grid">
<div class="sig-box"><div style="height:48px"></div><div class="sig-label">Submitted by (AIME)</div><div style="font-size:10pt;font-weight:700;margin-top:4px">${rfi.submitted_by || ""}</div><div style="font-size:9pt;color:#555;margin-top:4px">Date: ${rfi.date_submitted || "______________"}</div></div>
<div class="sig-box"><div style="height:48px"></div><div class="sig-label">Response by (Client)</div><div style="height:24px"></div><div style="font-size:9pt;color:#555;margin-top:4px">Date: ______________</div></div></div>
<div class="footer"><span>AIME Field Pro · RFI #${rfi.rfi_number} · ${project.name}</span><span>Generated: ${new Date().toLocaleString()}</span></div>
</div></body></html>`;
    const win = window.open("", "_blank", "width=850,height=800");
    win.document.write(html);
    win.document.close();
    win.focus();
  }

  const open = rfis.filter((r) => r.status === "Open" || r.status === "Overdue").length;
  const answered = rfis.filter((r) => r.status === "Answered").length;

  if (showForm || editing)
    return (
      <div style={{ padding: "14px 16px 80px" }}>
        <button
          onClick={() => {
            setShowForm(false);
            setEditing(null);
            setF(blank);
          }}
          style={{ ...ghostBtn, marginBottom: 14 }}
        >
          ← Back
        </button>
        <div style={{ fontSize: 17, fontWeight: 800, marginBottom: 16, color: T.text }}>
          {editing ? "Edit" : "New"} RFI
        </div>
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10, marginBottom: 10 }}>
          <div>
            <label style={lbl}>RFI Number *</label>
            <input
              value={f.rfi_number}
              onChange={(e) => set("rfi_number", e.target.value)}
              placeholder="RFI-001"
              style={inp}
            />
          </div>
          <div>
            <label style={lbl}>Date Submitted</label>
            <input
              type="date"
              value={f.date_submitted || ""}
              onChange={(e) => set("date_submitted", e.target.value)}
              style={inp}
            />
          </div>
        </div>
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10, marginBottom: 10 }}>
          <div>
            <label style={lbl}>Submitted By</label>
            <input value={f.submitted_by} onChange={(e) => set("submitted_by", e.target.value)} style={inp} />
          </div>
          <div>
            <label style={lbl}>Response Due</label>
            <input type="date" value={f.due_date || ""} onChange={(e) => set("due_date", e.target.value)} style={inp} />
          </div>
        </div>
        <div style={{ marginBottom: 10 }}>
          <label style={lbl}>Question / Issue *</label>
          <input
            value={f.question}
            onChange={(e) => set("question", e.target.value)}
            placeholder="Brief summary of the question..."
            style={inp}
          />
        </div>
        <div style={{ marginBottom: 10 }}>
          <label style={lbl}>Full Description</label>
          <textarea
            rows={3}
            value={f.description}
            onChange={(e) => set("description", e.target.value)}
            placeholder="Detailed description..."
            style={{ ...inp, resize: "vertical" }}
          />
        </div>
        <div style={{ marginBottom: 10 }}>
          <label style={lbl}>Status</label>
          <select value={f.status} onChange={(e) => set("status", e.target.value)} style={inpSel}>
            {["Open", "Answered", "Closed", "Overdue"].map((s) => (
              <option key={s}>{s}</option>
            ))}
          </select>
        </div>
        <div style={{ ...cardS, marginBottom: 10, borderLeft: `3px solid ${T.orange}`, background: T.orangeLow }}>
          <div style={{ fontSize: 12, fontWeight: 700, color: T.orange, marginBottom: 8 }}>🏀 BALL IN COURT</div>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10 }}>
            <div>
              <label style={lbl}>Person Responsible</label>
              <input
                value={f.ball_in_court}
                onChange={(e) => set("ball_in_court", e.target.value)}
                placeholder="Name of person who needs to act"
                style={inp}
              />
            </div>
            <div>
              <label style={lbl}>Their Email</label>
              <input
                type="email"
                value={f.ball_in_court_email}
                onChange={(e) => set("ball_in_court_email", e.target.value)}
                placeholder="email@company.com"
                style={inp}
              />
            </div>
          </div>
        </div>
        {(f.status === "Answered" || f.status === "Closed") && (
          <>
            <div style={{ marginBottom: 10 }}>
              <label style={lbl}>Response</label>
              <textarea
                rows={3}
                value={f.response}
                onChange={(e) => set("response", e.target.value)}
                style={{ ...inp, resize: "vertical" }}
              />
            </div>
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10, marginBottom: 10 }}>
              <div>
                <label style={lbl}>Responded By</label>
                <input value={f.responded_by} onChange={(e) => set("responded_by", e.target.value)} style={inp} />
              </div>
              <div>
                <label style={lbl}>Response Date</label>
                <input
                  type="date"
                  value={f.response_date || ""}
                  onChange={(e) => set("response_date", e.target.value)}
                  style={inp}
                />
              </div>
            </div>
          </>
        )}
        <button
          onClick={save}
          disabled={!f.rfi_number.trim() || !f.question.trim() || saving}
          style={{
            ...primBtn,
            opacity: f.rfi_number.trim() && f.question.trim() && !saving ? 1 : 0.5,
            borderRadius: 14,
          }}
        >
          {saving ? "Saving…" : "Save RFI"}
        </button>
      </div>
    );

  return (
    <div style={{ padding: "14px 16px 80px" }}>
      {/* Share Modal */}
      {shareRfi && (
        <div
          onClick={() => setShareRfi(null)}
          style={{
            position: "fixed",
            inset: 0,
            zIndex: 300,
            background: "rgba(0,0,0,0.85)",
            display: "flex",
            alignItems: "flex-end",
            justifyContent: "center",
          }}
        >
          <div
            onClick={(e) => e.stopPropagation()}
            style={{
              background: T.card,
              borderRadius: "20px 20px 0 0",
              padding: "20px 20px 40px",
              width: "100%",
              maxWidth: 480,
            }}
          >
            <div style={{ fontSize: 17, fontWeight: 900, marginBottom: 4, color: T.text }}>
              📤 Share RFI #{shareRfi.rfi_number}
            </div>
            <div style={{ fontSize: 12, color: T.muted, marginBottom: 20 }}>
              Send this link to {shareRfi.ball_in_court || "the recipient"} — they fill in their response and it saves
              directly to your app.
            </div>
            <div
              style={{
                background: T.surface,
                border: `1px solid ${T.border}`,
                borderRadius: 12,
                padding: "12px 14px",
                marginBottom: 14,
                wordBreak: "break-all",
                fontSize: 12,
                color: T.sub,
              }}
            >{`${appUrl}?rfi=${shareRfi.id}`}</div>
            <button
              onClick={() => copyLink(shareRfi)}
              style={{
                ...primBtn,
                borderRadius: 14,
                marginBottom: 10,
                background: copied ? T.green : T.orange,
                transition: "background 0.2s",
              }}
            >
              {copied ? "✅ Link Copied! Paste it into your email" : "📋 Copy Link to Clipboard"}
            </button>
            {copied && (
              <div
                style={{
                  background: T.greenLow,
                  border: `1px solid ${T.green}40`,
                  borderRadius: 10,
                  padding: "10px 14px",
                  marginBottom: 10,
                  fontSize: 13,
                  color: T.green,
                  textAlign: "center",
                  fontWeight: 600,
                }}
              >
                ✓ Link copied! Open your email app, paste it, and send.
              </div>
            )}
            <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 10 }}>
              <div style={{ flex: 1, height: 1, background: T.border }} />
              <span style={{ fontSize: 11, color: T.muted }}>OR</span>
              <div style={{ flex: 1, height: 1, background: T.border }} />
            </div>
            <button
              onClick={() => openEmailDraft(shareRfi)}
              style={{ ...ghostBtn, width: "100%", textAlign: "center", marginBottom: 10, fontSize: 14 }}
            >
              📧 Open Email Draft {shareRfi.ball_in_court_email ? `(to ${shareRfi.ball_in_court_email})` : ""}
            </button>
            <div style={{ fontSize: 11, color: T.muted, textAlign: "center", marginBottom: 14 }}>
              Note: Copy Link gives a clickable link. Email draft is plain text only.
            </div>
            <button onClick={() => setShareRfi(null)} style={{ ...ghostBtn, width: "100%", textAlign: "center" }}>
              Done
            </button>
          </div>
        </div>
      )}

      {rfis.length > 0 && (
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 8, marginBottom: 14 }}>
          {[
            [open, "Open / Overdue", T.yellow],
            [answered, "Answered", T.blue],
            [rfis.filter((r) => r.status === "Closed").length, "Closed", T.green],
          ].map(([v, l, c]) => (
            <div
              key={l}
              style={{
                background: T.card,
                borderRadius: 10,
                padding: "10px",
                textAlign: "center",
                border: `1px solid ${c}30`,
              }}
            >
              <div style={{ fontSize: 20, fontWeight: 900, color: c }}>{v}</div>
              <div style={{ fontSize: 9, color: T.muted, textTransform: "uppercase", marginTop: 2 }}>{l}</div>
            </div>
          ))}
        </div>
      )}
      {canEdit && (
        <button onClick={() => setShowForm(true)} style={{ ...primBtn, borderRadius: 14, marginBottom: 14 }}>
          + New RFI
        </button>
      )}
      {loading && <Spinner />}
      {!loading && rfis.length === 0 && (
        <div style={{ textAlign: "center", padding: "40px 0", color: T.muted }}>
          <div style={{ fontSize: 32, marginBottom: 8 }}>📝</div>
          <div style={{ fontWeight: 700, color: T.sub }}>No RFIs Yet</div>
        </div>
      )}
      {rfis.map((rfi) => {
        const isOverdue = rfi.due_date && new Date(rfi.due_date) < new Date() && rfi.status === "Open";
        const effStatus = isOverdue ? "Overdue" : rfi.status;
        const sc = statusColor[effStatus] || T.muted;
        return (
          <div key={rfi.id} style={{ ...cardS, marginBottom: 10, borderLeft: `3px solid ${sc}` }}>
            <div
              style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 8 }}
            >
              <div>
                <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 2 }}>
                  <span style={{ fontSize: 15, fontWeight: 800, color: T.orange }}>{rfi.rfi_number}</span>
                  <span style={pill(sc)}>{effStatus.toUpperCase()}</span>
                </div>
                <div style={{ fontSize: 12, color: T.muted }}>
                  {rfi.date_submitted} · {rfi.submitted_by}
                </div>
                {rfi.due_date && (
                  <div style={{ fontSize: 11, color: isOverdue ? T.red : T.muted, marginTop: 2 }}>
                    {isOverdue ? "⚠️ Overdue — " : "Due: "}
                    {rfi.due_date}
                  </div>
                )}
                {rfi.ball_in_court && (
                  <div style={{ fontSize: 11, color: T.orange, marginTop: 2, fontWeight: 600 }}>
                    🏀 Ball in Court: {rfi.ball_in_court}
                  </div>
                )}
              </div>
            </div>
            <div style={{ fontSize: 14, fontWeight: 700, color: T.text, marginBottom: 4 }}>{rfi.question}</div>
            {rfi.description && (
              <div style={{ fontSize: 12, color: T.sub, marginBottom: 8, lineHeight: 1.5 }}>{rfi.description}</div>
            )}
            {rfi.response && (
              <div
                style={{
                  background: T.greenLow,
                  border: `1px solid ${T.green}40`,
                  borderRadius: 10,
                  padding: "10px",
                  marginTop: 8,
                }}
              >
                <div style={{ fontSize: 11, fontWeight: 700, color: T.green, marginBottom: 4 }}>
                  ✓ RESPONSE — {rfi.responded_by}
                  {rfi.response_date ? " · " + rfi.response_date : ""}
                </div>
                <div style={{ fontSize: 12, color: T.text, lineHeight: 1.5 }}>{rfi.response}</div>
                {rfi.response_signature && (
                  <div style={{ marginTop: 8, background: "#fff", borderRadius: 8, padding: 4 }}>
                    <div style={{ fontSize: 9, color: "#999", marginBottom: 2, paddingLeft: 4 }}>SIGNATURE</div>
                    <img
                      src={rfi.response_signature}
                      alt="Signature"
                      style={{ width: "100%", maxHeight: 80, objectFit: "contain", display: "block", borderRadius: 6 }}
                    />
                  </div>
                )}
              </div>
            )}
            <div
              style={{
                display: "flex",
                gap: 6,
                marginTop: 10,
                paddingTop: 10,
                borderTop: `1px solid ${T.border}`,
                flexWrap: "wrap",
              }}
            >
              <button
                onClick={() => {
                  setShareRfi(rfi);
                  setCopied(false);
                }}
                style={{
                  ...ghostBtn,
                  flex: 1,
                  textAlign: "center",
                  fontSize: 12,
                  color: T.orange,
                  border: `1px solid ${T.orange}40`,
                  fontWeight: 700,
                  minWidth: 90,
                }}
              >
                📤 Send Link
              </button>
              <button
                onClick={() => printRFI(rfi)}
                style={{
                  ...ghostBtn,
                  flex: 1,
                  textAlign: "center",
                  fontSize: 12,
                  color: "#CBD5E1",
                  border: "1px solid #27272A",
                  minWidth: 70,
                }}
              >
                🖨️ PDF
              </button>
              {canEdit && (
                <>
                  <button
                    onClick={() => {
                      setEditing(rfi.id);
                      setF({
                        ...rfi,
                        notes: rfi.notes || "",
                        ball_in_court: rfi.ball_in_court || "",
                        ball_in_court_email: rfi.ball_in_court_email || "",
                      });
                    }}
                    style={{ ...ghostBtn, flex: 1, textAlign: "center", fontSize: 12, minWidth: 60 }}
                  >
                    ✏️
                  </button>
                  <button
                    onClick={() => remove(rfi.id)}
                    style={{
                      ...ghostBtn,
                      flex: 1,
                      textAlign: "center",
                      fontSize: 12,
                      color: T.red,
                      border: `1px solid ${T.red}40`,
                      minWidth: 50,
                    }}
                  >
                    🗑
                  </button>
                </>
              )}
            </div>
          </div>
        );
      })}
    </div>
  );
}

/* ── PUBLIC RFI FORM (no login required) ────────────────────── */
function PublicRFIForm({ rfiId }) {
  const [rfi, setRfi] = useState(null);
  const [project, setProject] = useState(null);
  const [loading, setLoading] = useState(true);
  const [saving, setSaving] = useState(false);
  const [submitted, setSubmitted] = useState(false);
  const [err, setErr] = useState("");
  const [resp, setResp] = useState("");
  const [respBy, setRespBy] = useState("");
  const [respDate, setRespDate] = useState(today());
  const [respTitle, setRespTitle] = useState("");
  const [sigData, setSigData] = useState(null);
  const [drawing, setDrawing] = useState(false);
  const sigRef = React.useRef(null);
  useEffect(() => {
    (async () => {
      try {
        const d = await sb(`/rfis?id=eq.${rfiId}&limit=1`);
        const r = Array.isArray(d) ? d[0] : d;
        if (!r) {
          setErr("RFI not found.");
          setLoading(false);
          return;
        }
        setRfi(r);
        if (r.response) {
          setResp(r.response);
          setRespBy(r.responded_by || "");
          setSubmitted(true);
        }
        const pd = await sb(`/projects?id=eq.${r.project_id}&limit=1`);
        setProject(Array.isArray(pd) ? pd[0] : pd || { name: "Unknown" });
      } catch (e) {
        setErr(e.message);
      }
      setLoading(false);
    })();
  }, [rfiId]);
  function getPos(e, c) {
    const r = c.getBoundingClientRect();
    const sx = c.width / r.width;
    const sy = c.height / r.height;
    if (e.touches) return { x: (e.touches[0].clientX - r.left) * sx, y: (e.touches[0].clientY - r.top) * sy };
    return { x: (e.clientX - r.left) * sx, y: (e.clientY - r.top) * sy };
  }
  function startSig(e) {
    e.preventDefault();
    const c = sigRef.current;
    if (!c) return;
    const p = getPos(e, c);
    c.getContext("2d").beginPath();
    c.getContext("2d").moveTo(p.x, p.y);
    setDrawing(true);
  }
  function drawSig(e) {
    e.preventDefault();
    if (!drawing) return;
    const c = sigRef.current;
    if (!c) return;
    const ctx = c.getContext("2d");
    const p = getPos(e, c);
    ctx.lineWidth = 2.5;
    ctx.lineCap = "round";
    ctx.strokeStyle = "#1f3864";
    ctx.lineTo(p.x, p.y);
    ctx.stroke();
    ctx.beginPath();
    ctx.moveTo(p.x, p.y);
  }
  function endSig(e) {
    e.preventDefault();
    setDrawing(false);
    const c = sigRef.current;
    if (c) setSigData(c.toDataURL("image/jpeg", 0.7));
  }
  function clearSig() {
    const c = sigRef.current;
    if (c) c.getContext("2d").clearRect(0, 0, c.width, c.height);
    setSigData(null);
  }
  async function submit() {
    if (!resp.trim() || !respBy.trim() || !sigData) {
      setErr("Please complete all required fields and sign.");
      return;
    }
    setSaving(true);
    setErr("");
    try {
      await sb(`/rfis?id=eq.${rfiId}`, {
        method: "PATCH",
        body: {
          response: resp,
          responded_by: respBy + (respTitle ? " (" + respTitle + ")" : ""),
          response_date: respDate || today(),
          status: "Answered",
          response_signature: sigData,
        },
        prefer: "return=representation",
      });
      setSubmitted(true);
    } catch (e) {
      setErr("Failed: " + e.message);
    }
    setSaving(false);
  }
  const s = {
    bg: "#16181D",
    card: "#FFFFFF",
    inp: {
      width: "100%",
      background: "#FFFFFF",
      border: "1px solid #27272A",
      borderRadius: 10,
      color: "#16181D",
      fontSize: 14,
      padding: "12px 14px",
      fontFamily: "inherit",
      outline: "none",
    },
    lbl: {
      display: "block",
      fontSize: 11,
      fontWeight: 700,
      color: "#6B7280",
      textTransform: "uppercase",
      letterSpacing: "1px",
      marginBottom: 6,
    },
  };
  if (loading)
    return (
      <div
        style={{
          background: s.bg,
          minHeight: "100vh",
          display: "flex",
          alignItems: "center",
          justifyContent: "center",
        }}
      >
        <div style={{ color: "#16181D" }}>Loading...</div>
      </div>
    );
  if (err && !rfi)
    return (
      <div
        style={{
          background: s.bg,
          minHeight: "100vh",
          display: "flex",
          alignItems: "center",
          justifyContent: "center",
          padding: 20,
        }}
      >
        <div style={{ background: s.card, borderRadius: 16, padding: 32, maxWidth: 400, textAlign: "center" }}>
          <div style={{ fontSize: 40 }}>⚠️</div>
          <div style={{ color: "#16181D", fontWeight: 700, margin: "8px 0" }}>{err}</div>
        </div>
      </div>
    );
  if (submitted)
    return (
      <div style={{ background: s.bg, minHeight: "100vh", fontFamily: "system-ui", padding: 20 }}>
        <div
          style={{
            maxWidth: 600,
            margin: "0 auto",
            background: s.card,
            borderRadius: 16,
            padding: 32,
            textAlign: "center",
            border: "1px solid #22C55E40",
          }}
        >
          <div style={{ fontSize: 48 }}>✅</div>
          <div style={{ fontSize: 20, fontWeight: 800, color: "#22C55E", margin: "8px 0" }}>Response Submitted!</div>
          <div style={{ color: "#6B7280" }}>RFI #{rfi?.rfi_number} response has been saved.</div>
        </div>
      </div>
    );
  return (
    <div
      style={{
        background: s.bg,
        minHeight: "100vh",
        fontFamily: "system-ui",
        color: "#16181D",
        padding: "16px 16px 60px",
      }}
    >
      <div style={{ maxWidth: 600, margin: "0 auto" }}>
        <div
          style={{
            background: "#1f3864",
            borderRadius: 16,
            padding: "16px 20px",
            marginBottom: 16,
            display: "flex",
            justifyContent: "space-between",
            alignItems: "center",
          }}
        >
          <div>
            <div style={{ fontSize: 20, fontWeight: 900, color: "#F97316" }}>AIME</div>
            <div style={{ fontSize: 11, color: "rgba(255,255,255,0.7)" }}>
              Atlantic Industrial Mechanical & Environmental Inc.
            </div>
          </div>
          <div style={{ textAlign: "right" }}>
            <div style={{ fontSize: 18, fontWeight: 800 }}>RFI #{rfi.rfi_number}</div>
          </div>
        </div>
        <div style={{ background: "#1f3864", borderRadius: 12, padding: 16, marginBottom: 12 }}>
          <div style={{ fontSize: 10, color: "rgba(255,255,255,0.6)", textTransform: "uppercase", marginBottom: 6 }}>
            Question / Issue
          </div>
          <div style={{ fontSize: 16, fontWeight: 700, lineHeight: 1.5 }}>{rfi.question}</div>
          {rfi.description && (
            <div style={{ fontSize: 13, color: "rgba(255,255,255,0.7)", marginTop: 8 }}>{rfi.description}</div>
          )}
        </div>
        {err && (
          <div
            style={{
              background: "#EF444420",
              border: "1px solid #EF4444",
              borderRadius: 10,
              padding: "10px 14px",
              marginBottom: 12,
              fontSize: 13,
              color: "#EF4444",
            }}
          >
            {err}
          </div>
        )}
        <div style={{ background: s.card, borderRadius: 16, padding: 20, border: "2px solid #22C55E40" }}>
          <div style={{ fontSize: 14, fontWeight: 800, color: "#22C55E", marginBottom: 14 }}>✏️ Your Response</div>
          <div style={{ marginBottom: 12 }}>
            <label style={s.lbl}>Response *</label>
            <textarea
              rows={6}
              value={resp}
              onChange={(e) => setResp(e.target.value)}
              placeholder="Enter your response..."
              style={{ ...s.inp, resize: "vertical", lineHeight: 1.6 }}
            />
          </div>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10, marginBottom: 12 }}>
            <div>
              <label style={s.lbl}>Your Name *</label>
              <input value={respBy} onChange={(e) => setRespBy(e.target.value)} placeholder="Full name" style={s.inp} />
            </div>
            <div>
              <label style={s.lbl}>Date</label>
              <input type="date" value={respDate} onChange={(e) => setRespDate(e.target.value)} style={s.inp} />
            </div>
          </div>
          <div style={{ marginBottom: 14 }}>
            <label style={s.lbl}>Company / Title</label>
            <input
              value={respTitle}
              onChange={(e) => setRespTitle(e.target.value)}
              placeholder="e.g. Colonial Pipeline · Engineer"
              style={s.inp}
            />
          </div>
          <div style={{ marginBottom: 16 }}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 6 }}>
              <label style={s.lbl}>Signature *</label>
              {sigData && (
                <button
                  onClick={clearSig}
                  style={{ background: "none", border: "none", color: "#EF4444", fontSize: 12, cursor: "pointer" }}
                >
                  Clear
                </button>
              )}
            </div>
            <div style={{ background: "#fff", borderRadius: 10, overflow: "hidden", border: "2px solid #22C55E" }}>
              <canvas
                ref={sigRef}
                width={600}
                height={160}
                style={{ width: "100%", height: 160, display: "block", touchAction: "none", cursor: "crosshair" }}
                onMouseDown={startSig}
                onMouseMove={drawSig}
                onMouseUp={endSig}
                onMouseLeave={endSig}
                onTouchStart={startSig}
                onTouchMove={drawSig}
                onTouchEnd={endSig}
              />
            </div>
            {!sigData && (
              <div style={{ textAlign: "center", fontSize: 11, color: "#6B7280", marginTop: 4 }}>
                Draw your signature above
              </div>
            )}
            {sigData && (
              <div style={{ textAlign: "center", fontSize: 11, color: "#22C55E", marginTop: 4 }}>
                ✓ Signature captured
              </div>
            )}
          </div>
          <button
            onClick={submit}
            disabled={saving || !resp.trim() || !respBy.trim() || !sigData}
            style={{
              width: "100%",
              background: saving || !resp.trim() || !respBy.trim() || !sigData ? "#D8D4CC" : "#22C55E",
              color: saving || !resp.trim() || !respBy.trim() || !sigData ? "#6B7280" : "#000",
              border: "none",
              borderRadius: 12,
              padding: 14,
              fontSize: 16,
              fontWeight: 800,
              cursor: "pointer",
              fontFamily: "inherit",
            }}
          >
            {saving ? "Submitting…" : "✅ Submit Signed Response"}
          </button>
        </div>
      </div>
    </div>
  );
}

/* ── PUBLIC CO FORM (no login required) ─────────────────────── */
function PublicCOForm({ coId }) {
  const [co, setCo] = useState(null);
  const [project, setProject] = useState(null);
  const [loading, setLoading] = useState(true);
  const [saving, setSaving] = useState(false);
  const [submitted, setSubmitted] = useState(false);
  const [err, setErr] = useState("");
  const [signerName, setSignerName] = useState("");
  const [signerTitle, setSignerTitle] = useState("");
  const [sigData, setSigData] = useState(null);
  const [drawing, setDrawing] = useState(false);
  const sigRef = React.useRef(null);
  useEffect(() => {
    (async () => {
      try {
        const d = await sb(`/change_orders?id=eq.${coId}&limit=1`);
        const c = Array.isArray(d) ? d[0] : d;
        if (!c) {
          setErr("CO not found.");
          setLoading(false);
          return;
        }
        setCo(c);
        if (c.client_signature) setSubmitted(true);
        const pd = await sb(`/projects?id=eq.${c.project_id}&limit=1`);
        setProject(Array.isArray(pd) ? pd[0] : pd || { name: "Unknown" });
      } catch (e) {
        setErr(e.message);
      }
      setLoading(false);
    })();
  }, [coId]);
  function getPos(e, c) {
    const r = c.getBoundingClientRect();
    const sx = c.width / r.width;
    const sy = c.height / r.height;
    if (e.touches) return { x: (e.touches[0].clientX - r.left) * sx, y: (e.touches[0].clientY - r.top) * sy };
    return { x: (e.clientX - r.left) * sx, y: (e.clientY - r.top) * sy };
  }
  function startSig(e) {
    e.preventDefault();
    const c = sigRef.current;
    if (!c) return;
    const p = getPos(e, c);
    c.getContext("2d").beginPath();
    c.getContext("2d").moveTo(p.x, p.y);
    setDrawing(true);
  }
  function drawSig(e) {
    e.preventDefault();
    if (!drawing) return;
    const c = sigRef.current;
    if (!c) return;
    const ctx = c.getContext("2d");
    const p = getPos(e, c);
    ctx.lineWidth = 2.5;
    ctx.lineCap = "round";
    ctx.strokeStyle = "#1f3864";
    ctx.lineTo(p.x, p.y);
    ctx.stroke();
    ctx.beginPath();
    ctx.moveTo(p.x, p.y);
  }
  function endSig(e) {
    e.preventDefault();
    setDrawing(false);
    const c = sigRef.current;
    if (c) setSigData(c.toDataURL("image/jpeg", 0.7));
  }
  function clearSig() {
    const c = sigRef.current;
    if (c) c.getContext("2d").clearRect(0, 0, c.width, c.height);
    setSigData(null);
  }
  async function submit() {
    if (!signerName.trim() || !sigData) {
      setErr("Please enter your name and signature.");
      return;
    }
    setSaving(true);
    setErr("");
    try {
      await sb(`/change_orders?id=eq.${coId}`, {
        method: "PATCH",
        body: {
          client_signature: sigData,
          client_signed_by: signerName + (signerTitle ? " (" + signerTitle + ")" : ""),
          client_signed_date: today(),
          status: "Approved",
        },
        prefer: "return=representation",
      });
      setSubmitted(true);
    } catch (e) {
      setErr("Failed: " + e.message);
    }
    setSaving(false);
  }
  const s = {
    bg: "#16181D",
    card: "#FFFFFF",
    inp: {
      width: "100%",
      background: "#FFFFFF",
      border: "1px solid #27272A",
      borderRadius: 10,
      color: "#16181D",
      fontSize: 14,
      padding: "12px 14px",
      fontFamily: "inherit",
      outline: "none",
    },
    lbl: {
      display: "block",
      fontSize: 11,
      fontWeight: 700,
      color: "#6B7280",
      textTransform: "uppercase",
      letterSpacing: "1px",
      marginBottom: 6,
    },
  };
  const fmt = (n) => "$" + Number(n || 0).toLocaleString("en-US", { minimumFractionDigits: 2 });
  if (loading)
    return (
      <div
        style={{
          background: s.bg,
          minHeight: "100vh",
          display: "flex",
          alignItems: "center",
          justifyContent: "center",
        }}
      >
        <div style={{ color: "#16181D" }}>Loading...</div>
      </div>
    );
  if (err && !co)
    return (
      <div
        style={{
          background: s.bg,
          minHeight: "100vh",
          display: "flex",
          alignItems: "center",
          justifyContent: "center",
          padding: 20,
        }}
      >
        <div style={{ background: s.card, borderRadius: 16, padding: 32, maxWidth: 400, textAlign: "center" }}>
          <div style={{ fontSize: 40 }}>⚠️</div>
          <div style={{ color: "#16181D", fontWeight: 700, margin: "8px 0" }}>{err}</div>
        </div>
      </div>
    );
  if (submitted)
    return (
      <div style={{ background: s.bg, minHeight: "100vh", fontFamily: "system-ui", padding: 20 }}>
        <div
          style={{
            maxWidth: 560,
            margin: "0 auto",
            background: s.card,
            borderRadius: 16,
            padding: 32,
            textAlign: "center",
            border: "1px solid #22C55E40",
          }}
        >
          <div style={{ fontSize: 48 }}>✅</div>
          <div style={{ fontSize: 20, fontWeight: 800, color: "#22C55E", margin: "8px 0" }}>Change Order Signed!</div>
          <div style={{ color: "#6B7280" }}>CO {co?.co_number} has been approved.</div>
        </div>
      </div>
    );
  return (
    <div
      style={{
        background: s.bg,
        minHeight: "100vh",
        fontFamily: "system-ui",
        color: "#16181D",
        padding: "16px 16px 60px",
      }}
    >
      <div style={{ maxWidth: 560, margin: "0 auto" }}>
        <div
          style={{
            background: "#1f3864",
            borderRadius: 16,
            padding: "16px 20px",
            marginBottom: 16,
            display: "flex",
            justifyContent: "space-between",
            alignItems: "center",
          }}
        >
          <div>
            <div style={{ fontSize: 20, fontWeight: 900, color: "#F97316" }}>AIME</div>
            <div style={{ fontSize: 11, color: "rgba(255,255,255,0.7)" }}>
              Atlantic Industrial Mechanical & Environmental Inc.
            </div>
          </div>
          <div>
            <div style={{ fontSize: 18, fontWeight: 800 }}>{co.co_number}</div>
            <div style={{ fontSize: 11, color: "rgba(255,255,255,0.6)" }}>Change Order</div>
          </div>
        </div>
        <div style={{ background: "#1f3864", borderRadius: 12, padding: 16, marginBottom: 12 }}>
          <div style={{ fontSize: 10, color: "rgba(255,255,255,0.6)", textTransform: "uppercase", marginBottom: 6 }}>
            Description of Change
          </div>
          <div style={{ fontSize: 15, fontWeight: 700, lineHeight: 1.5, marginBottom: 12 }}>{co.description}</div>
          <div style={{ background: "rgba(255,255,255,0.1)", borderRadius: 10, padding: 12, textAlign: "center" }}>
            <div
              style={{ fontSize: 10, color: "rgba(255,255,255,0.6)", textTransform: "uppercase", letterSpacing: "1px" }}
            >
              Change Order Amount
            </div>
            <div style={{ fontSize: 28, fontWeight: 900, color: "#22C55E", marginTop: 4 }}>{fmt(co.amount)}</div>
          </div>
        </div>
        {err && (
          <div
            style={{
              background: "#EF444420",
              border: "1px solid #EF4444",
              borderRadius: 10,
              padding: "10px 14px",
              marginBottom: 12,
              fontSize: 13,
              color: "#EF4444",
            }}
          >
            {err}
          </div>
        )}
        <div style={{ background: s.card, borderRadius: 16, padding: 20, border: "2px solid #F9731640" }}>
          <div style={{ fontSize: 14, fontWeight: 800, color: "#F97316", marginBottom: 14 }}>
            ✍️ Sign & Approve Change Order
          </div>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10, marginBottom: 12 }}>
            <div>
              <label style={s.lbl}>Your Name *</label>
              <input
                value={signerName}
                onChange={(e) => setSignerName(e.target.value)}
                placeholder="Full name"
                style={s.inp}
              />
            </div>
            <div>
              <label style={s.lbl}>Title / Company</label>
              <input
                value={signerTitle}
                onChange={(e) => setSignerTitle(e.target.value)}
                placeholder="Your title"
                style={s.inp}
              />
            </div>
          </div>
          <div style={{ marginBottom: 16 }}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 6 }}>
              <label style={s.lbl}>Signature *</label>
              {sigData && (
                <button
                  onClick={clearSig}
                  style={{ background: "none", border: "none", color: "#EF4444", fontSize: 12, cursor: "pointer" }}
                >
                  Clear
                </button>
              )}
            </div>
            <div style={{ background: "#fff", borderRadius: 10, overflow: "hidden", border: "2px solid #F97316" }}>
              <canvas
                ref={sigRef}
                width={600}
                height={160}
                style={{ width: "100%", height: 160, display: "block", touchAction: "none", cursor: "crosshair" }}
                onMouseDown={startSig}
                onMouseMove={drawSig}
                onMouseUp={endSig}
                onMouseLeave={endSig}
                onTouchStart={startSig}
                onTouchMove={drawSig}
                onTouchEnd={endSig}
              />
            </div>
            {!sigData && (
              <div style={{ textAlign: "center", fontSize: 11, color: "#6B7280", marginTop: 4 }}>
                Draw your signature above
              </div>
            )}
            {sigData && (
              <div style={{ textAlign: "center", fontSize: 11, color: "#22C55E", marginTop: 4 }}>
                ✓ Signature captured
              </div>
            )}
          </div>
          <button
            onClick={submit}
            disabled={saving || !signerName.trim() || !sigData}
            style={{
              width: "100%",
              background: saving || !signerName.trim() || !sigData ? "#D8D4CC" : "#F97316",
              color: saving || !signerName.trim() || !sigData ? "#6B7280" : "#000",
              border: "none",
              borderRadius: 12,
              padding: 14,
              fontSize: 16,
              fontWeight: 800,
              cursor: "pointer",
              fontFamily: "inherit",
            }}
          >
            {saving ? "Submitting…" : "✅ Sign & Approve Change Order"}
          </button>
        </div>
      </div>
    </div>
  );
}

/* ── PUSH NOTIFICATIONS SETUP ───────────────────────────────── */
async function setupPushNotifications() {
  // Only runs in native Capacitor app, no-op on web
  const isNative = typeof window !== "undefined" && window.Capacitor?.isNativePlatform?.();
  if (!isNative) return;
  try {
    const cap = window.Capacitor?.Plugins;
    if (!cap) return;
    const { PushNotifications } = cap;
    if (!PushNotifications) return;
    const perm = await PushNotifications.requestPermissions();
    if (perm.receive === "granted") {
      await PushNotifications.register();
    }
  } catch (e) {
    // Silent fail on web
  }
}

/* ── APP INNER ───────────────────────────────────────────────── */
const SITE_CODE = "AIME2026"; // under-construction access code — change and share only with testers

function GateScreen({ onUnlock }) {
  const [code, setCode] = useState("");
  const [bad, setBad] = useState(false);
  const go = () => {
    if (code.trim().toUpperCase() === SITE_CODE.toUpperCase()) onUnlock();
    else setBad(true);
  };
  return (
    <div
      style={{
        background: T.bg,
        minHeight: "100vh",
        display: "flex",
        flexDirection: "column",
        alignItems: "center",
        justifyContent: "center",
        padding: 24,
      }}
    >
      <div
        style={{
          height: 8,
          width: "100%",
          maxWidth: 420,
          background: "repeating-linear-gradient(-45deg,#F25C05 0 14px,#16181D 14px 28px)",
          borderRadius: 4,
        }}
      />
      <div
        style={{
          background: "#fff",
          border: `1px solid ${T.border}`,
          borderRadius: 16,
          padding: 24,
          maxWidth: 420,
          width: "100%",
          margin: "14px 0",
        }}
      >
        <img src="/aime-report-logo.png" alt="AIME" style={{ width: "100%", height: "auto", marginBottom: 14 }} />
        <div style={{ textAlign: "center", marginBottom: 16 }}>
          <span
            style={{
              display: "inline-block",
              padding: "3px 10px",
              fontSize: 11,
              fontWeight: 700,
              letterSpacing: 2,
              textTransform: "uppercase",
              color: "#B45309",
              border: "2px solid #B45309",
              borderRadius: 4,
              transform: "rotate(-1.5deg)",
            }}
          >
            Under construction
          </span>
        </div>
        <label
          style={{
            display: "block",
            fontSize: 11,
            fontWeight: 700,
            color: T.muted,
            letterSpacing: 1,
            textTransform: "uppercase",
            marginBottom: 6,
          }}
        >
          Access code
        </label>
        <input
          type="password"
          value={code}
          autoFocus
          placeholder="••••••••"
          onChange={(e) => {
            setCode(e.target.value);
            setBad(false);
          }}
          onKeyDown={(e) => e.key === "Enter" && go()}
          style={{
            width: "100%",
            boxSizing: "border-box",
            background: "#fff",
            border: `1px solid ${bad ? T.red : T.border}`,
            borderRadius: 12,
            color: T.text,
            fontSize: 16,
            padding: "13px 14px",
            outline: "none",
            letterSpacing: 4,
          }}
        />
        {bad && <div style={{ color: T.red, fontSize: 13, marginTop: 8 }}>That code isn't right. Check with Doug.</div>}
        <button
          onClick={go}
          style={{
            width: "100%",
            background: T.orange,
            color: "#fff",
            border: "none",
            borderRadius: 14,
            padding: "16px",
            fontSize: 16,
            fontWeight: 800,
            cursor: "pointer",
            marginTop: 16,
          }}
        >
          ENTER
        </button>
        <div style={{ fontSize: 11, color: T.muted, textAlign: "center", marginTop: 12 }}>
          AIME Field Pro is in testing. Access is limited to invited users.
        </div>
      </div>
      <div
        style={{
          height: 8,
          width: "100%",
          maxWidth: 420,
          background: "repeating-linear-gradient(-45deg,#F25C05 0 14px,#16181D 14px 28px)",
          borderRadius: 4,
        }}
      />
    </div>
  );
}

function AppInner() {
  const [publicRfiId] = useState(() => new URLSearchParams(window.location.search).get("rfi"));
  const [publicCoId] = useState(() => new URLSearchParams(window.location.search).get("co"));
  const [user, setUser] = useState(null);
  const [projects, setProjects] = useState([]);
  const [screen, setScreen] = useState("division");
  const [selectedDiv, setSelectedDiv] = useState(null);
  const [selectedProject, setSelectedProject] = useState(null);
  const [isOnline, setIsOnline] = useState(navigator.onLine);
  const [pendingCount, setPendingCount] = useState(0);
  const [syncMsg, setSyncMsg] = useState("");
  const [err, setErr] = useState("");

  // Public forms (no login needed)
  if (publicRfiId) return <PublicRFIForm rfiId={publicRfiId} />;
  if (publicCoId) return <PublicCOForm coId={publicCoId} />;

  useEffect(() => {
    const onOnline = () => {
      setIsOnline(true);
      syncQueue();
    };
    const onOffline = () => setIsOnline(false);
    window.addEventListener("online", onOnline);
    window.addEventListener("offline", onOffline);
    setupPushNotifications();
    return () => {
      window.removeEventListener("online", onOnline);
      window.removeEventListener("offline", onOffline);
    };
  }, []);

  useEffect(() => {
    if (user) loadProjects();
  }, [user]);

  async function loadProjects() {
    try {
      const r = await API.projects.list();
      const ps = Array.isArray(r) ? r : [];
      // Load report counts and billing per project
      const enriched = await Promise.all(
        ps.map(async (p) => {
          try {
            const reps = await API.reports.forProject(p.id);
            const repList = Array.isArray(reps) ? reps : [];
            const billed = repList.reduce((s, r) => {
              const tot = r.labor_total || 0;
              const etot = r.equipment_total || 0;
              return s + tot + etot;
            }, 0);
            return { ...p, _reports: repList.length, _billed: billed };
          } catch {
            return { ...p, _reports: 0, _billed: 0 };
          }
        }),
      );
      setProjects(enriched);
    } catch (e) {
      setErr(e.message);
    }
  }

  async function syncQueue() {
    const queue = getQueue();
    if (!queue.length) return;
    setSyncMsg(`Syncing ${queue.length} queued report${queue.length !== 1 ? "s" : ""}…`);
    let synced = 0;
    for (const item of queue) {
      try {
        if (item.type === "report") {
          const { rental_equipment, ...dbData } = item.data;
          try {
            await API.reports.create({ ...dbData, rental_equipment, project_id: item.data.project_id });
          } catch {
            await API.reports.create({ ...dbData, project_id: item.data.project_id });
          }
          removeFromQueue(item.qid);
          const proj = projects.find((p) => p.id === item.data.project_id);
          if (proj) await autoPopulateTimeCards(item.data, proj).catch(() => {});
          synced++;
        }
      } catch (e) {
        console.warn("Sync failed for item", item.qid, e);
      }
    }
    setSyncMsg(synced > 0 ? `✓ Synced ${synced} report${synced !== 1 ? "s" : ""}!` : "");
    if (synced > 0) {
      await loadProjects();
    }
    setTimeout(() => setSyncMsg(""), 4000);
    setPendingCount(getQueue().length);
  }

  function handleLogin(profile) {
    setUser(profile);
  }
  function handleLogout() {
    setUser(null);
    setScreen("division");
    setSelectedDiv(null);
    setSelectedProject(null);
  }
  function handleDivisionSelect(div) {
    setSelectedDiv(div);
    setScreen("jobs");
  }
  function handleSelectProject(p) {
    setSelectedProject(p);
    setScreen("detail");
  }

  if (!user)
    return (
      <div
        style={{
          maxWidth: 480,
          margin: "0 auto",
          fontFamily: "'DM Sans',system-ui,sans-serif",
          color: T.text,
          background: T.bg,
          minHeight: "100vh",
        }}
      >
        <LoginScreen onLogin={handleLogin} />
      </div>
    );

  const canEst = can(user, "estimating");

  return (
    <div
      style={{
        maxWidth: 480,
        margin: "0 auto",
        fontFamily: "'DM Sans',system-ui,sans-serif",
        color: T.text,
        background: T.bg,
        minHeight: "100vh",
      }}
    >
      {syncMsg && (
        <div
          style={{
            background: T.green,
            color: "#fff",
            padding: "10px 16px",
            fontSize: 13,
            fontWeight: 700,
            textAlign: "center",
          }}
        >
          {syncMsg}
        </div>
      )}
      {err && (
        <div
          style={{ background: T.red, color: "#fff", padding: "8px 16px", fontSize: 12, cursor: "pointer" }}
          onClick={() => setErr("")}
        >
          {err} ✕
        </div>
      )}
      {screen === "division" && (
        <DivisionScreen
          user={user}
          projects={projects}
          onSelect={handleDivisionSelect}
          onLogout={handleLogout}
          onCrew={() => setScreen("crewDirectory")}
          onDash={() => setScreen("pmDashboard")}
          onTimeCards={() => setScreen("timeCards")}
          onEstimating={() => setScreen("estimating")}
          isOnline={isOnline}
          pendingCount={pendingCount}
          onSync={syncQueue}
        />
      )}
      {screen === "estimating" && can(user, "estimating") && (
        <EstimatingScreen user={user} onBack={() => setScreen("division")} />
      )}
      {screen === "jobs" && selectedDiv && (
        <JobBoard
          user={user}
          projects={projects}
          division={selectedDiv}
          onSelect={handleSelectProject}
          onBack={() => setScreen("division")}
          onNew={() => setScreen("newProject")}
          onRefresh={loadProjects}
        />
      )}
      {screen === "detail" && selectedProject && (
        <ProjectDetail
          project={projects.find((p) => p.id === selectedProject.id) || selectedProject}
          user={user}
          onBack={() => setScreen("jobs")}
          onRefresh={loadProjects}
          onErr={setErr}
          isOnline={isOnline}
        />
      )}
      {screen === "newProject" && (
        <ProjectForm
          user={user}
          onSave={async (data) => {
            try {
              await API.projects.create(data);
              await loadProjects();
              setScreen("jobs");
            } catch (e) {
              setErr(e.message);
            }
          }}
          onCancel={() => setScreen("jobs")}
        />
      )}
      {screen === "pmDashboard" && (
        <PMDashboard
          user={user}
          projects={projects}
          onBack={() => setScreen("division")}
          onRefresh={loadProjects}
          onErr={setErr}
        />
      )}
      {screen === "crewDirectory" && <CrewDirectoryScreen user={user} onBack={() => setScreen("division")} />}
      {screen === "timeCards" && (
        <TimeCardsScreen user={user} projects={projects} onBack={() => setScreen("division")} />
      )}
      {screen === "userManagement" && <UserManagementScreen user={user} onBack={() => setScreen("division")} />}
    </div>
  );
}

export default function App() {
  // Public client links (?rfi= / ?co=) bypass the construction gate — clients sign without a code
  const isPublicLink = (() => {
    const q = new URLSearchParams(window.location.search);
    return q.has("rfi") || q.has("co");
  })();
  const [unlocked, setUnlocked] = useState(isPublicLink);
  return (
    <ErrorBoundary>
      <style>{`@import url('https://fonts.googleapis.com/css2?family=Barlow+Condensed:wght@600;700&family=Archivo:wght@400;600;700;800&family=IBM+Plex+Mono:wght@400;600&display=swap');
body{font-family:'Archivo','Helvetica Neue',system-ui,sans-serif;background:#EDEBE6;}
button,input,select,textarea{font-family:inherit;}
h1,h2,h3{font-family:'Barlow Condensed','Arial Narrow',sans-serif;}`}</style>
      {unlocked ? <AppInner /> : <GateScreen onUnlock={() => setUnlocked(true)} />}
    </ErrorBoundary>
  );
}
