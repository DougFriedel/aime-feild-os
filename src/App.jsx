// AIME Training — standalone video training app
// Vite + React + Supabase (same project as APP_V3)
// npm i @supabase/supabase-js tus-js-client
// .env: VITE_SUPABASE_URL=...  VITE_SUPABASE_ANON_KEY=...

import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import { createClient } from "@supabase/supabase-js";
import * as tus from "tus-js-client";

const SUPABASE_URL = import.meta.env.VITE_SUPABASE_URL;
const SUPABASE_ANON_KEY = import.meta.env.VITE_SUPABASE_ANON_KEY;
const BUCKET = "training-videos";
const supabase = SUPABASE_URL && SUPABASE_ANON_KEY ? createClient(SUPABASE_URL, SUPABASE_ANON_KEY) : null;

/* ------------------------------------------------------------------ */
/* helpers                                                             */
/* ------------------------------------------------------------------ */

const pad = (n) => String(n).padStart(2, "0");
function fmtDur(s) {
  if (s === null || s === undefined || s === "" || !isFinite(Number(s))) return "—";
  s = Math.round(Number(s));
  const h = Math.floor(s / 3600), m = Math.floor((s % 3600) / 60), sec = s % 60;
  return h ? `${h}:${pad(m)}:${pad(sec)}` : `${m}:${pad(sec)}`;
}
function fmtDate(d) {
  if (!d) return "—";
  const dt = new Date(d.length === 10 ? d + "T00:00:00" : d);
  return dt.toLocaleDateString(undefined, { month: "short", day: "numeric", year: "numeric" });
}
function fmtSize(b) {
  if (!b) return "—";
  if (b > 1e9) return (b / 1e9).toFixed(1) + " GB";
  return Math.round(b / 1e6) + " MB";
}
function todayISO() {
  const d = new Date();
  return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
}
const safeName = (n) => n.replace(/[^a-zA-Z0-9._-]+/g, "_").slice(-120);
const pkey = (u, v) => `${u}:${v}`;

// new | watching | quiz | done
function statusOf(video, prog) {
  if (!prog) return "new";
  if (prog.completed_at) return "done";
  if (prog.watched_complete) return video?.quiz_count > 0 ? "quiz" : "done";
  if (Number(prog.max_position) > 0) return "watching";
  return "new";
}
const STATUS_LABEL = { new: "Not started", watching: "In progress", quiz: "Quiz needed", done: "Complete" };
const isLate = (st, due) => st !== "done" && !!due && due < todayISO();
function watchPct(video, prog) {
  const d = Number(video?.duration_seconds);
  if (!prog) return 0;
  if (prog.watched_complete) return 100;
  if (!d) return 0;
  return Math.min(100, Math.round((100 * Number(prog.max_position)) / d));
}

function groupVideos(videos, categories) {
  const groups = categories.map((c) => ({ id: c.id, name: c.name, videos: [] }));
  const byId = Object.fromEntries(groups.map((g) => [g.id, g]));
  const loose = { id: "none", name: "Uncategorized", videos: [] };
  for (const v of videos) (byId[v.category_id] || loose).videos.push(v);
  return [...groups, loose].filter((g) => g.videos.length);
}

async function fetchAll(table, ...orderCols) {
  const out = [];
  const size = 1000;
  for (let from = 0; ; from += size) {
    let query = supabase.from(table).select("*");
    for (const c of orderCols) query = query.order(c);
    const { data, error } = await query.range(from, from + size - 1);
    if (error) throw error;
    out.push(...data);
    if (data.length < size) return out;
  }
}

function getVideoDuration(file) {
  return new Promise((resolve) => {
    const url = URL.createObjectURL(file);
    const v = document.createElement("video");
    v.preload = "metadata";
    let finished = false;
    const done = (d) => {
      if (finished) return;
      finished = true;
      URL.revokeObjectURL(url);
      resolve(d && isFinite(d) ? d : null);
    };
    v.onloadedmetadata = () => done(v.duration);
    v.onerror = () => done(null);
    setTimeout(() => done(null), 15000);
    v.src = url;
  });
}

// Resumable upload (handles large video files, retries on dropped connections)
async function uploadResumable(path, file, onProgress) {
  const { data } = await supabase.auth.getSession();
  if (!data.session) throw new Error("Your session expired. Sign in again.");
  await new Promise((resolve, reject) => {
    const upload = new tus.Upload(file, {
      endpoint: `${SUPABASE_URL}/storage/v1/upload/resumable`,
      retryDelays: [0, 3000, 5000, 10000, 20000],
      headers: {
        authorization: `Bearer ${data.session.access_token}`,
        apikey: SUPABASE_ANON_KEY,
        "x-upsert": "false",
      },
      uploadDataDuringCreation: true,
      removeFingerprintOnSuccess: true,
      storeFingerprintForResuming: false,
      metadata: {
        bucketName: BUCKET,
        objectName: path,
        contentType: file.type || "video/mp4",
        cacheControl: "3600",
      },
      chunkSize: 6 * 1024 * 1024, // Supabase requires 6 MB chunks
      onBeforeRequest: async (req) => {
        const { data: s } = await supabase.auth.getSession();
        if (s.session) req.setHeader("authorization", `Bearer ${s.session.access_token}`);
      },
      onProgress: (sent, total) => onProgress(total ? sent / total : 0),
      onError: (e) => reject(e),
      onSuccess: () => resolve(),
    });
    upload.start();
  });
}

function downloadCSV(filename, rows) {
  const esc = (v) => `"${String(v ?? "").replace(/"/g, '""')}"`;
  const csv = rows.map((r) => r.map(esc).join(",")).join("\r\n");
  const a = document.createElement("a");
  a.href = URL.createObjectURL(new Blob([csv], { type: "text/csv" }));
  a.download = filename;
  a.click();
  setTimeout(() => URL.revokeObjectURL(a.href), 2000);
}

/* ------------------------------------------------------------------ */
/* styles                                                              */
/* ------------------------------------------------------------------ */

const CSS = `
@import url('https://fonts.googleapis.com/css2?family=Barlow:wght@400;500;600;700&family=Barlow+Condensed:wght@600;700&display=swap');
html, body, #root { margin:0; padding:0; background:#F2F4F5; }
.tr { --steel:#22313A; --steel-2:#2E4250; --ink:#1B2328; --muted:#5B6A73; --panel:#F2F4F5; --card:#FFFFFF;
  --line:#D6DCE0; --hivis:#F5B700; --hivis-ink:#3A2C00; --go:#2F7D4E; --go-bg:#E3F1E8; --stop:#B42318;
  --stop-bg:#FBE9E7; --warn-bg:#FFF4D1;
  font-family: Barlow, "Segoe UI", Roboto, Arial, sans-serif; color: var(--ink); background: var(--panel);
  min-height: 100vh; font-size: 16px; line-height: 1.45; text-align: left; color-scheme: light; }
.tr *, .tr *::before, .tr *::after { box-sizing: border-box; }
.tr h1, .tr h2, .tr h3 { font-family: "Barlow Condensed", Barlow, sans-serif; font-weight: 700; line-height: 1.1; margin: 0; }
.tr h1 { font-size: 2rem; } .tr h2 { font-size: 1.6rem; } .tr h3 { font-size: 1.2rem; }
.tr p { margin: 0; }
.tr :focus-visible { outline: 3px solid var(--hivis); outline-offset: 2px; }

.tr-top { background: var(--steel); color: #fff; }
.tr-top-in { max-width: 1180px; margin: 0 auto; padding: 10px 20px; display: flex; align-items: center; gap: 18px; flex-wrap: wrap; }
.tr-brand { display: flex; align-items: center; gap: 10px; font-family: "Barlow Condensed", sans-serif; font-weight: 700; font-size: 1.4rem; }
.tr-brand img { height: 34px; width: auto; display: block; }
.tr-tabs { display: flex; gap: 2px; flex-wrap: wrap; }
.tr-tab { appearance: none; border: 0; background: transparent; color: #C9D3D9; font: 600 1rem Barlow, sans-serif; padding: 10px 14px; cursor: pointer; border-radius: 6px 6px 0 0; }
.tr-tab:hover { color: #fff; background: var(--steel-2); }
.tr-tab[aria-current="page"] { color: #fff; box-shadow: inset 0 -3px 0 var(--hivis); }
.tr-top-right { margin-left: auto; display: flex; align-items: center; gap: 12px; font-size: .95rem; }
.tr-main { max-width: 1180px; margin: 0 auto; padding: 24px 20px 72px; }

.tr-btn { appearance: none; display: inline-flex; align-items: center; justify-content: center; gap: 6px; font: 600 .95rem Barlow, sans-serif;
  padding: 9px 16px; border-radius: 6px; border: 1px solid var(--steel); background: var(--steel); color: #fff; cursor: pointer; white-space: nowrap; }
.tr-btn:hover { background: var(--steel-2); }
.tr-btn:disabled { opacity: .5; cursor: not-allowed; }
.tr-btn.ghost { background: transparent; color: var(--steel); border-color: var(--line); }
.tr-btn.ghost:hover { background: #fff; border-color: var(--steel); }
.tr-btn.danger { background: transparent; color: var(--stop); border-color: #E6B8B2; }
.tr-btn.danger:hover { background: var(--stop-bg); }
.tr-btn.go { background: var(--hivis); border-color: var(--hivis); color: var(--hivis-ink); }
.tr-btn.go:hover { background: #FFC52E; }
.tr-btn.sm { padding: 5px 10px; font-size: .85rem; }
.tr-top .tr-btn.ghost, .tr-sticky .tr-btn.ghost { color: #fff; border-color: #4A5E6B; }
.tr-top .tr-btn.ghost:hover, .tr-sticky .tr-btn.ghost:hover { background: var(--steel-2); }
.tr-link { appearance: none; background: none; border: 0; padding: 0; color: var(--steel); font: inherit; text-decoration: underline; cursor: pointer; }
.tr-link.danger { color: var(--stop); }

.tr-field { display: flex; flex-direction: column; gap: 4px; }
.tr-field > span { font-size: .85rem; font-weight: 600; color: var(--muted); }
.tr input[type=text], .tr input[type=email], .tr input[type=password], .tr input[type=date], .tr input[type=number],
.tr input[type=search], .tr select, .tr textarea { font: inherit; color: var(--ink); background: #fff; border: 1px solid var(--line);
  border-radius: 6px; padding: 8px 10px; width: 100%; }
.tr textarea { min-height: 72px; resize: vertical; }
.tr input:focus, .tr select:focus, .tr textarea:focus { outline: 2px solid var(--steel); outline-offset: 0; border-color: var(--steel); }
.tr-check { display: inline-flex; align-items: center; gap: 8px; cursor: pointer; }
.tr-check input { width: 18px; height: 18px; accent-color: var(--steel); margin: 0; }
.tr-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(220px, 1fr)); gap: 14px; }

.tr-card { background: var(--card); border: 1px solid var(--line); border-radius: 8px; padding: 18px; }
.tr-stack { display: flex; flex-direction: column; gap: 18px; }
.tr-row { display: flex; align-items: center; gap: 10px; flex-wrap: wrap; }
.tr-spacer { flex: 1; }
.tr-muted { color: var(--muted); }
.tr-small { font-size: .875rem; }
.tr-msg { padding: 10px 14px; border-radius: 6px; font-weight: 500; }
.tr-msg.err { background: var(--stop-bg); color: var(--stop); }
.tr-msg.ok { background: var(--go-bg); color: var(--go); }
.tr-msg.warn { background: var(--warn-bg); color: #6B4E00; }

.tr-chip { display: inline-block; font-size: .8rem; font-weight: 600; padding: 2px 9px; border-radius: 999px; background: #E6EAED; color: var(--muted); white-space: nowrap; }
.tr-chip.done { background: var(--go-bg); color: var(--go); }
.tr-chip.quiz { background: var(--warn-bg); color: #6B4E00; }
.tr-chip.watching { background: #E1ECF4; color: #1F5375; }
.tr-chip.late { background: var(--stop-bg); color: var(--stop); }

.tr-bar { height: 8px; border-radius: 4px; background: #E1E6E9; overflow: hidden; min-width: 70px; }
.tr-bar > i { display: block; height: 100%; background: var(--go); }

.tr-hero { background: var(--steel); color: #fff; }
.tr-hero-in { max-width: 1180px; margin: 0 auto; padding: 26px 20px 30px; display: grid; grid-template-columns: auto 1fr; gap: 14px 28px; align-items: end; }
.tr-count { font-family: "Barlow Condensed", sans-serif; font-weight: 700; font-size: clamp(3.5rem, 10vw, 6rem); line-height: .85; font-variant-numeric: tabular-nums; }
.tr-count small { font-size: .42em; color: #AFC0CA; }
.tr-hero p { color: #C9D3D9; font-size: 1.05rem; }
.tr-hero-bar { grid-column: 1 / -1; height: 14px; border-radius: 3px; overflow: hidden;
  background: repeating-linear-gradient(-45deg, #2E4250 0 10px, #37505F 10px 20px); }
.tr-hero-bar > i { display: block; height: 100%; background: var(--hivis); transition: width .7s ease; }
@media (prefers-reduced-motion: reduce) { .tr-hero-bar > i { transition: none; } }

.tr-next { border-left: 6px solid var(--hivis); }
.tr-list { list-style: none; margin: 0; padding: 0; background: #fff; border: 1px solid var(--line); border-radius: 8px; overflow: hidden; }
.tr-list > li + li { border-top: 1px solid var(--line); }
.tr-item { display: flex; align-items: center; gap: 14px; padding: 13px 16px; width: 100%; }
.tr-item-btn { appearance: none; text-align: left; background: #fff; border: 0; font: inherit; color: inherit; cursor: pointer; }
.tr-item-btn:hover { background: #F7F9FA; }
.tr-item-main { flex: 1; min-width: 0; }
.tr-item-title { font-weight: 600; }

.tr-tablewrap { overflow-x: auto; background: #fff; border: 1px solid var(--line); border-radius: 8px; }
.tr-table { width: 100%; border-collapse: collapse; font-size: .95rem; }
.tr-table th { text-align: left; font-size: .82rem; color: var(--muted); font-weight: 600; padding: 10px 12px; border-bottom: 1px solid var(--line); background: #FAFBFB; white-space: nowrap; }
.tr-table td { padding: 10px 12px; border-bottom: 1px solid #EAEEF0; vertical-align: middle; }
.tr-table tr:last-child td { border-bottom: 0; }
.tr-table tr.click { cursor: pointer; }
.tr-table tr.click:hover td { background: #F7F9FA; }
.tr-num { text-align: right; font-variant-numeric: tabular-nums; }

.tr-player { background: #000; border-radius: 8px; overflow: hidden; aspect-ratio: 16 / 9; display: grid; place-items: center; color: #C9D3D9; }
.tr-player video { width: 100%; height: 100%; display: block; background: #000; }

.tr-q { padding: 18px 0; border-top: 1px solid var(--line); }
.tr-q:first-child { border-top: 0; padding-top: 0; }
.tr-q fieldset { border: 0; margin: 0; padding: 0; min-width: 0; }
.tr-q legend { font-weight: 600; padding: 0; margin-bottom: 8px; font-size: 1.05rem; }
.tr-opt { display: flex; gap: 10px; align-items: flex-start; padding: 9px 12px; border: 1px solid var(--line); border-radius: 6px; margin-top: 6px; cursor: pointer; background: #fff; }
.tr-opt:has(input:checked) { border-color: var(--steel); background: #EEF2F4; }
.tr-opt input { margin-top: 4px; accent-color: var(--steel); }
.tr-mark { font-weight: 700; margin-left: 8px; }
.tr-mark.y { color: var(--go); } .tr-mark.n { color: var(--stop); }

.tr-split { display: grid; grid-template-columns: minmax(260px, 1fr) 2fr; gap: 18px; align-items: start; }
.tr-pick { max-height: 62vh; overflow: auto; padding: 0; }
.tr-pickhead { padding: 12px 14px; border-bottom: 1px solid var(--line); background: #FAFBFB; position: sticky; top: 0; z-index: 1; }
.tr-pickrow { display: flex; align-items: center; gap: 10px; padding: 9px 14px; border-top: 1px solid #EAEEF0; }
.tr-pickrow:first-of-type { border-top: 0; }
.tr-catrow { display: flex; align-items: center; gap: 10px; padding: 10px 14px; background: #F4F6F7; border-top: 1px solid var(--line); font-weight: 700; }
.tr-sticky { position: sticky; bottom: 12px; background: var(--steel); color: #fff; padding: 12px 16px; border-radius: 8px; display: flex; gap: 12px; align-items: center; flex-wrap: wrap; margin-top: 18px; z-index: 5; }
.tr-sticky input[type=date] { width: auto; }

.tr-auth { min-height: 100vh; display: grid; place-items: center; padding: 20px; background: var(--steel); }
.tr-auth .tr-card { width: 100%; max-width: 410px; }
.tr-auth .tr-brand { color: #fff; justify-content: center; margin-bottom: 18px; font-size: 1.8rem; }

.tr-qedit { border: 1px solid var(--line); border-radius: 8px; padding: 14px; background: #FAFBFB; }
.tr-optedit { display: grid; grid-template-columns: auto 1fr auto; gap: 8px; align-items: center; margin-top: 6px; }
.tr-catchips { display: flex; gap: 8px; flex-wrap: wrap; }
.tr-catchip { display: inline-flex; align-items: center; gap: 6px; background: #fff; border: 1px solid var(--line); border-radius: 999px; padding: 3px 6px 3px 12px; font-weight: 600; font-size: .9rem; }
.tr-catchip button { appearance: none; border: 0; background: none; color: var(--muted); cursor: pointer; font-size: 1rem; line-height: 1; padding: 2px 6px; border-radius: 999px; }
.tr-catchip button:hover { background: var(--stop-bg); color: var(--stop); }

@media (max-width: 820px) {
  .tr-split { grid-template-columns: 1fr; }
  .tr-hero-in { grid-template-columns: 1fr; }
  .tr-top-right { margin-left: 0; width: 100%; justify-content: space-between; }
}
`;

/* ------------------------------------------------------------------ */
/* shared bits                                                         */
/* ------------------------------------------------------------------ */

function Logo() {
  const [ok, setOk] = useState(true);
  return ok ? <img src="/aime-logo.png" alt="AIME" onError={() => setOk(false)} /> : <span>AIME</span>;
}

function TopBar({ profile, tabs, tab, setTab }) {
  return (
    <header className="tr-top">
      <div className="tr-top-in">
        <div className="tr-brand"><Logo /><span>Training</span></div>
        {tabs && (
          <nav className="tr-tabs" aria-label="Sections">
            {tabs.map(([k, label]) => (
              <button key={k} className="tr-tab" aria-current={tab === k ? "page" : undefined} onClick={() => setTab(k)}>{label}</button>
            ))}
          </nav>
        )}
        <div className="tr-top-right">
          <span>{profile.full_name || profile.email}</span>
          <button className="tr-btn ghost sm" onClick={() => supabase.auth.signOut()}>Sign out</button>
        </div>
      </div>
    </header>
  );
}

function StatusChip({ st, late }) {
  if (late) return <span className="tr-chip late">Overdue</span>;
  return <span className={`tr-chip ${st}`}>{STATUS_LABEL[st]}</span>;
}

function Msg({ msg }) {
  if (!msg) return null;
  return <div className={`tr-msg ${msg.type}`} role={msg.type === "err" ? "alert" : "status"}>{msg.text}</div>;
}

function Loading({ text = "Loading…" }) {
  return <div className="tr-main tr-muted">{text}</div>;
}

/* ------------------------------------------------------------------ */
/* app root                                                            */
/* ------------------------------------------------------------------ */

export default function App() {
  return (
    <div className="tr">
      <style>{CSS}</style>
      {supabase ? <Root /> : (
        <div className="tr-main"><div className="tr-msg err">Missing VITE_SUPABASE_URL or VITE_SUPABASE_ANON_KEY. Add them to .env (and to Netlify environment variables), then rebuild.</div></div>
      )}
    </div>
  );
}

function Root() {
  const [session, setSession] = useState(undefined);
  const [recovery, setRecovery] = useState(false);
  const [profile, setProfile] = useState(null);
  const [profileError, setProfileError] = useState("");

  useEffect(() => {
    supabase.auth.getSession().then(({ data }) => setSession(data.session));
    const { data: sub } = supabase.auth.onAuthStateChange((event, s) => {
      if (event === "PASSWORD_RECOVERY") setRecovery(true);
      setSession(s);
    });
    return () => sub.subscription.unsubscribe();
  }, []);

  const user = session?.user;
  const uid = user?.id;
  useEffect(() => {
    if (!uid) { setProfile(null); return; }
    let cancelled = false;
    (async () => {
      let { data, error } = await supabase.from("training_profiles").select("*").eq("id", uid).maybeSingle();
      if (!error && !data) {
        const ins = await supabase.from("training_profiles")
          .insert({ id: uid, email: user.email, full_name: user.user_metadata?.full_name || user.email.split("@")[0] })
          .select().single();
        data = ins.data; error = ins.error;
      }
      if (cancelled) return;
      if (error) setProfileError(error.message);
      else { setProfileError(""); setProfile(data); }
    })();
    return () => { cancelled = true; };
  }, [uid]); // eslint-disable-line react-hooks/exhaustive-deps

  if (session === undefined) return <Loading />;
  if (recovery) return <SetPasswordScreen onDone={() => setRecovery(false)} />;
  if (!session) return <AuthScreen />;
  if (profileError) return <div className="tr-main"><div className="tr-msg err">Couldn't load your training profile: {profileError}</div></div>;
  if (!profile) return <Loading />;
  if (!profile.active) {
    return (
      <>
        <TopBar profile={profile} />
        <div className="tr-main"><div className="tr-msg warn">Your training account is turned off. Ask your supervisor to turn it back on.</div></div>
      </>
    );
  }
  return profile.role === "admin" ? <AdminApp profile={profile} /> : <TraineeApp profile={profile} />;
}

/* ------------------------------------------------------------------ */
/* sign in                                                             */
/* ------------------------------------------------------------------ */

function AuthScreen() {
  const [mode, setMode] = useState("signin"); // signin | signup | reset
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [name, setName] = useState("");
  const [busy, setBusy] = useState(false);
  const [msg, setMsg] = useState(null);

  async function submit(e) {
    e.preventDefault();
    setBusy(true); setMsg(null);
    try {
      if (mode === "signin") {
        const { error } = await supabase.auth.signInWithPassword({ email, password });
        if (error) throw error;
      } else if (mode === "signup") {
        if (!name.trim()) throw new Error("Enter your full name.");
        const { data, error } = await supabase.auth.signUp({
          email, password, options: { data: { full_name: name.trim() }, emailRedirectTo: window.location.origin },
        });
        if (error) throw error;
        if (!data.session) setMsg({ type: "ok", text: "Check your email to confirm your account, then sign in." });
      } else {
        const { error } = await supabase.auth.resetPasswordForEmail(email, { redirectTo: window.location.origin });
        if (error) throw error;
        setMsg({ type: "ok", text: "Reset link sent. Check your email." });
      }
    } catch (err) {
      setMsg({ type: "err", text: err.message });
    } finally {
      setBusy(false);
    }
  }

  const titles = { signin: "Sign in", signup: "Create your account", reset: "Reset your password" };
  const actions = { signin: "Sign in", signup: "Create account", reset: "Send reset link" };

  return (
    <div className="tr-auth">
      <div style={{ width: "100%", maxWidth: 410 }}>
        <div className="tr-brand"><Logo /><span>Training</span></div>
        <form className="tr-card tr-stack" onSubmit={submit}>
          <h2>{titles[mode]}</h2>
          {mode === "signup" && (
            <label className="tr-field"><span>Full name</span>
              <input type="text" value={name} onChange={(e) => setName(e.target.value)} autoComplete="name" required />
            </label>
          )}
          <label className="tr-field"><span>Work email</span>
            <input type="email" value={email} onChange={(e) => setEmail(e.target.value)} autoComplete="email" required />
          </label>
          {mode !== "reset" && (
            <label className="tr-field"><span>Password</span>
              <input type="password" value={password} onChange={(e) => setPassword(e.target.value)} minLength={6}
                autoComplete={mode === "signup" ? "new-password" : "current-password"} required />
            </label>
          )}
          <Msg msg={msg} />
          <button className="tr-btn" disabled={busy}>{busy ? "Working…" : actions[mode]}</button>
          <div className="tr-row tr-small">
            {mode !== "signin" && <button type="button" className="tr-link" onClick={() => { setMode("signin"); setMsg(null); }}>Back to sign in</button>}
            {mode === "signin" && <button type="button" className="tr-link" onClick={() => { setMode("signup"); setMsg(null); }}>New hire? Create an account</button>}
            <span className="tr-spacer" />
            {mode === "signin" && <button type="button" className="tr-link" onClick={() => { setMode("reset"); setMsg(null); }}>Forgot password</button>}
          </div>
        </form>
      </div>
    </div>
  );
}

function SetPasswordScreen({ onDone }) {
  const [pw, setPw] = useState("");
  const [busy, setBusy] = useState(false);
  const [msg, setMsg] = useState(null);
  async function submit(e) {
    e.preventDefault();
    setBusy(true);
    const { error } = await supabase.auth.updateUser({ password: pw });
    setBusy(false);
    if (error) setMsg({ type: "err", text: error.message });
    else onDone();
  }
  return (
    <div className="tr-auth">
      <form className="tr-card tr-stack" onSubmit={submit}>
        <h2>Set a new password</h2>
        <label className="tr-field"><span>New password</span>
          <input type="password" value={pw} onChange={(e) => setPw(e.target.value)} minLength={6} autoComplete="new-password" required />
        </label>
        <Msg msg={msg} />
        <button className="tr-btn" disabled={busy}>{busy ? "Saving…" : "Save password"}</button>
      </form>
    </div>
  );
}

/* ------------------------------------------------------------------ */
/* trainee                                                             */
/* ------------------------------------------------------------------ */

function TraineeApp({ profile }) {
  const [items, setItems] = useState(null);
  const [error, setError] = useState("");
  const [open, setOpen] = useState(null); // { videoId, mode: 'watch' | 'quiz' }

  const load = useCallback(async () => {
    const [a, p] = await Promise.all([
      supabase.from("training_assignments")
        .select("id, due_date, assigned_at, video:training_videos(id, title, description, duration_seconds, quiz_count, pass_score, storage_path, require_full_watch, category:training_categories(id, name, sort_order))")
        .eq("user_id", profile.id),
      supabase.from("training_progress").select("*").eq("user_id", profile.id),
    ]);
    if (a.error || p.error) { setError((a.error || p.error).message); return; }
    const pm = Object.fromEntries(p.data.map((r) => [r.video_id, r]));
    setItems(a.data.filter((r) => r.video).map((r) => ({
      due: r.due_date, assignedAt: r.assigned_at, video: r.video, progress: pm[r.video.id] || null,
    })));
  }, [profile.id]);

  useEffect(() => { load(); }, [load]);

  const close = () => { setOpen(null); load(); };
  const current = open && items?.find((i) => i.video.id === open.videoId);

  if (error) return <><TopBar profile={profile} /><div className="tr-main"><div className="tr-msg err">{error}</div></div></>;
  if (!items) return <><TopBar profile={profile} /><Loading /></>;

  if (current) {
    return (
      <>
        <TopBar profile={profile} />
        <main className="tr-main">
          {open.mode === "quiz"
            ? <QuizView item={current} onBack={close} onRewatch={() => setOpen({ videoId: current.video.id, mode: "watch" })} />
            : <WatchView item={current} onBack={close} onQuiz={() => { setOpen({ videoId: current.video.id, mode: "quiz" }); load(); }} />}
        </main>
      </>
    );
  }

  const withSt = items.map((i) => {
    const st = statusOf(i.video, i.progress);
    return { ...i, st, late: isLate(st, i.due) };
  });
  const total = withSt.length;
  const done = withSt.filter((i) => i.st === "done").length;
  const next = withSt.filter((i) => i.st !== "done")
    .sort((a, b) => (a.due || "9999").localeCompare(b.due || "9999") || a.assignedAt.localeCompare(b.assignedAt))[0];

  const cats = {};
  for (const i of withSt) {
    const c = i.video.category;
    const k = c?.id || "none";
    if (!cats[k]) cats[k] = { name: c?.name || "General", order: c ? c.sort_order : 1e9, items: [] };
    cats[k].items.push(i);
  }
  const groups = Object.values(cats).sort((a, b) => a.order - b.order || a.name.localeCompare(b.name));
  const firstName = (profile.full_name || "").split(" ")[0] || "there";
  const openItem = (i) => setOpen({ videoId: i.video.id, mode: i.st === "quiz" ? "quiz" : "watch" });

  return (
    <>
      <TopBar profile={profile} />
      <section className="tr-hero">
        <div className="tr-hero-in">
          <div className="tr-count" aria-label={`${done} of ${total} videos complete`}>{done}<small>/{total}</small></div>
          <div>
            <h1>Hi, {firstName}</h1>
            <p>{total === 0 ? "No training assigned yet." : done === total ? "All of your assigned training is complete." : `${total - done} video${total - done === 1 ? "" : "s"} left to finish.`}</p>
          </div>
          {total > 0 && <div className="tr-hero-bar"><i style={{ width: `${(100 * done) / total}%` }} /></div>}
        </div>
      </section>
      <main className="tr-main tr-stack">
        {total === 0 && <div className="tr-card">Your supervisor hasn't assigned any videos yet. They'll show up here as soon as they do.</div>}
        {next && (
          <div className="tr-card tr-next tr-row">
            <div className="tr-item-main">
              <div className="tr-small tr-muted">Up next</div>
              <h2>{next.video.title}</h2>
              <div className="tr-small tr-muted">
                {fmtDur(next.video.duration_seconds)}
                {next.due && <>, due {fmtDate(next.due)}</>}
              </div>
            </div>
            {next.late && <StatusChip st={next.st} late />}
            <button className="tr-btn go" onClick={() => openItem(next)}>
              {next.st === "quiz" ? "Take the quiz" : next.st === "watching" ? "Keep watching" : "Start video"}
            </button>
          </div>
        )}
        {groups.map((g) => (
          <section key={g.name} className="tr-stack" style={{ gap: 8 }}>
            <h3>{g.name}</h3>
            <ul className="tr-list">
              {g.items.map((i) => (
                <li key={i.video.id}>
                  <button className="tr-item tr-item-btn" onClick={() => openItem(i)}>
                    <div className="tr-item-main">
                      <div className="tr-item-title">{i.video.title}</div>
                      <div className="tr-small tr-muted">
                        {fmtDur(i.video.duration_seconds)}
                        {i.video.quiz_count > 0 && <>, {i.video.quiz_count}-question quiz</>}
                        {i.due && <>, due {fmtDate(i.due)}</>}
                      </div>
                    </div>
                    {i.st === "watching" && <div className="tr-bar" style={{ width: 90 }}><i style={{ width: `${watchPct(i.video, i.progress)}%` }} /></div>}
                    <StatusChip st={i.st} late={i.late} />
                  </button>
                </li>
              ))}
            </ul>
          </section>
        ))}
      </main>
    </>
  );
}

function WatchView({ item, onBack, onQuiz }) {
  const { video } = item;
  const [url, setUrl] = useState(null);
  const [err, setErr] = useState("");
  const [notice, setNotice] = useState("");
  const [prog, setProg] = useState(item.progress);
  const [pct, setPct] = useState(watchPct(video, item.progress));
  const maxPos = useRef(Number(item.progress?.max_position || 0));
  const lastSaved = useRef(maxPos.current);
  const dur = useRef(Number(video.duration_seconds) || 0);
  const saving = useRef(false);

  const complete = !!prog?.watched_complete;
  const locked = video.require_full_watch && !complete;
  const st = statusOf(video, prog);

  useEffect(() => {
    if (!video.storage_path) { setErr("This video hasn't been uploaded yet. Let your supervisor know."); return; }
    supabase.storage.from(BUCKET).createSignedUrl(video.storage_path, 60 * 60 * 6)
      .then(({ data, error }) => (error ? setErr(error.message) : setUrl(data.signedUrl)));
  }, [video.storage_path]);

  const save = useCallback(async () => {
    if (!dur.current || saving.current) return;
    saving.current = true;
    lastSaved.current = maxPos.current;
    const { data, error } = await supabase.rpc("training_record_watch", {
      p_video: video.id, p_position: maxPos.current, p_duration: dur.current,
    });
    saving.current = false;
    if (!error && data) setProg(data);
  }, [video.id]);

  useEffect(() => () => { if (maxPos.current > lastSaved.current) save(); }, [save]);

  function onLoaded(e) {
    const el = e.currentTarget;
    if (isFinite(el.duration) && el.duration > 0) dur.current = el.duration;
    if (!complete && maxPos.current > 5 && maxPos.current < dur.current - 5) el.currentTime = maxPos.current;
  }
  function onTime(e) {
    const el = e.currentTarget;
    const t = el.currentTime;
    const allowJump = !video.require_full_watch;
    if (!el.seeking && t > maxPos.current && (allowJump || t - maxPos.current < 3)) {
      maxPos.current = t;
      if (dur.current) setPct(Math.min(100, Math.round((100 * t) / dur.current)));
    }
    if (maxPos.current - lastSaved.current >= 10) save();
  }
  function onSeeking(e) {
    const el = e.currentTarget;
    if (locked && el.currentTime > maxPos.current + 2) {
      el.currentTime = maxPos.current;
      setNotice("You can skip ahead once you've watched that part.");
      setTimeout(() => setNotice(""), 4000);
    }
  }
  function onEnded() {
    if (dur.current) { maxPos.current = dur.current; setPct(100); }
    save();
  }

  return (
    <div className="tr-stack">
      <div><button className="tr-link" onClick={onBack}>Back to my training</button></div>
      <div>
        <h1>{video.title}</h1>
        {video.category && <div className="tr-muted">{video.category.name}</div>}
      </div>
      <div className="tr-player">
        {err ? <span style={{ padding: 20 }}>{err}</span> : url ? (
          <video src={url} controls playsInline preload="metadata" controlsList="nodownload noplaybackrate"
            disablePictureInPicture onContextMenu={(e) => e.preventDefault()}
            onLoadedMetadata={onLoaded} onTimeUpdate={onTime} onSeeking={onSeeking} onPause={save} onEnded={onEnded} />
        ) : <span>Loading video…</span>}
      </div>
      {notice && <div className="tr-msg warn" role="status">{notice}</div>}
      <div className="tr-row">
        <div className="tr-bar" style={{ flex: 1 }}><i style={{ width: `${complete ? 100 : pct}%` }} /></div>
        <span className="tr-small tr-muted">{complete ? "Watched" : `${pct}% watched`}</span>
      </div>
      {video.description && <p style={{ maxWidth: "70ch", whiteSpace: "pre-wrap" }}>{video.description}</p>}
      {st === "quiz" && (
        <div className="tr-card tr-next tr-row">
          <div className="tr-item-main">
            <h3>Video done. One more step.</h3>
            <div className="tr-muted">{video.quiz_count} question{video.quiz_count === 1 ? "" : "s"}. You need {video.pass_score}% to pass.</div>
          </div>
          <button className="tr-btn go" onClick={onQuiz}>Take the quiz</button>
        </div>
      )}
      {st === "done" && (
        <div className="tr-msg ok tr-row">
          <span className="tr-item-main">This training is complete.</span>
          <button className="tr-btn ghost sm" onClick={onBack}>Back to my training</button>
        </div>
      )}
      {!complete && video.require_full_watch && video.quiz_count > 0 && (
        <p className="tr-small tr-muted">Watch the whole video to unlock the quiz.</p>
      )}
    </div>
  );
}

function QuizView({ item, onBack, onRewatch }) {
  const { video } = item;
  const [qs, setQs] = useState(null);
  const [answers, setAnswers] = useState({});
  const [result, setResult] = useState(null);
  const [busy, setBusy] = useState(false);
  const [err, setErr] = useState("");

  useEffect(() => {
    supabase.rpc("training_get_quiz", { p_video: video.id })
      .then(({ data, error }) => (error ? setErr(error.message) : setQs(data)));
  }, [video.id]);

  async function submit(e) {
    e.preventDefault();
    if (qs.some((q) => answers[q.question_id] === undefined)) { setErr("Answer every question before submitting."); return; }
    setBusy(true); setErr("");
    const { data, error } = await supabase.rpc("training_submit_quiz", { p_video: video.id, p_answers: answers });
    setBusy(false);
    if (error) setErr(error.message);
    else { setResult(data); window.scrollTo({ top: 0, behavior: "smooth" }); }
  }
  function retry() { setAnswers({}); setResult(null); window.scrollTo({ top: 0 }); }

  return (
    <form className="tr-stack" onSubmit={submit} style={{ maxWidth: 760 }}>
      <div><button type="button" className="tr-link" onClick={onBack}>Back to my training</button></div>
      <div>
        <h1>Quiz: {video.title}</h1>
        <div className="tr-muted">You need {video.pass_score}% to pass. You can retake it as many times as you need.</div>
      </div>
      {result && (
        <div className={`tr-msg ${result.passed ? "ok" : "err"}`} role="status">
          {result.passed
            ? `Passed with ${result.score}% (${result.correct} of ${result.total}). This training is complete.`
            : `${result.score}% (${result.correct} of ${result.total}). You need ${result.pass_score}% to pass. Questions you missed are marked below.`}
        </div>
      )}
      {err && <div className="tr-msg err" role="alert">{err}</div>}
      {!qs ? <div className="tr-muted">Loading quiz…</div> : (
        <div className="tr-card">
          {qs.map((q, qi) => {
            const mark = result ? result.results[q.question_id] : null;
            return (
              <div key={q.question_id} className="tr-q">
                <fieldset disabled={!!result}>
                  <legend>
                    {qi + 1}. {q.prompt}
                    {result && <span className={`tr-mark ${mark ? "y" : "n"}`}>{mark ? "Correct" : "Incorrect"}</span>}
                  </legend>
                  {q.options.map((opt, oi) => (
                    <label key={oi} className="tr-opt">
                      <input type="radio" name={q.question_id} checked={answers[q.question_id] === oi}
                        onChange={() => setAnswers((a) => ({ ...a, [q.question_id]: oi }))} />
                      <span>{opt}</span>
                    </label>
                  ))}
                </fieldset>
              </div>
            );
          })}
        </div>
      )}
      <div className="tr-row">
        {!result && qs && <button className="tr-btn go" disabled={busy}>{busy ? "Checking…" : "Submit answers"}</button>}
        {result && !result.passed && <>
          <button type="button" className="tr-btn go" onClick={retry}>Try again</button>
          <button type="button" className="tr-btn ghost" onClick={onRewatch}>Rewatch the video</button>
        </>}
        {result?.passed && <button type="button" className="tr-btn" onClick={onBack}>Back to my training</button>}
      </div>
    </form>
  );
}

/* ------------------------------------------------------------------ */
/* admin                                                               */
/* ------------------------------------------------------------------ */

function useMaps(data) {
  return useMemo(() => {
    if (!data) return null;
    const videoById = Object.fromEntries(data.videos.map((v) => [v.id, v]));
    const catById = Object.fromEntries(data.categories.map((c) => [c.id, c]));
    const prog = Object.fromEntries(data.progress.map((p) => [pkey(p.user_id, p.video_id), p]));
    const asgByUser = {}, asgByVideo = {}, asg = {};
    for (const a of data.assignments) {
      (asgByUser[a.user_id] ||= []).push(a);
      (asgByVideo[a.video_id] ||= []).push(a);
      asg[pkey(a.user_id, a.video_id)] = a;
    }
    return { videoById, catById, prog, asgByUser, asgByVideo, asg };
  }, [data]);
}

function summarize(userId, maps) {
  const list = maps.asgByUser[userId] || [];
  let done = 0, overdue = 0, last = null, total = 0;
  for (const a of list) {
    const v = maps.videoById[a.video_id];
    if (!v) continue;
    total++;
    const p = maps.prog[pkey(userId, a.video_id)];
    const st = statusOf(v, p);
    if (st === "done") done++;
    else if (isLate(st, a.due_date)) overdue++;
    if (p?.last_activity && (!last || p.last_activity > last)) last = p.last_activity;
  }
  return { total, done, overdue, last };
}

function AdminApp({ profile }) {
  const [tab, setTab] = useState("people");
  const [data, setData] = useState(null);
  const [error, setError] = useState("");
  const [assignFor, setAssignFor] = useState(null);

  const load = useCallback(async () => {
    try {
      const [people, videos, categories, assignments, progress] = await Promise.all([
        fetchAll("training_profiles", "id"),
        fetchAll("training_videos", "created_at", "id"),
        fetchAll("training_categories", "sort_order"),
        fetchAll("training_assignments", "id"),
        fetchAll("training_progress", "user_id", "video_id"),
      ]);
      people.sort((a, b) => (a.full_name || a.email || "").localeCompare(b.full_name || b.email || ""));
      categories.sort((a, b) => a.sort_order - b.sort_order || a.name.localeCompare(b.name));
      setData({ people, videos, categories, assignments, progress });
      setError("");
    } catch (e) {
      setError(e.message);
    }
  }, []);
  useEffect(() => { load(); }, [load]);
  const maps = useMaps(data);

  const tabs = [["people", "People"], ["assign", "Assign"], ["library", "Video library"]];
  const goAssign = (personId) => { setAssignFor(personId); setTab("assign"); };

  return (
    <>
      <TopBar profile={profile} tabs={tabs} tab={tab} setTab={(t) => { setTab(t); if (t !== "assign") setAssignFor(null); }} />
      <main className="tr-main">
        {error && <div className="tr-msg err">{error}</div>}
        {!data ? <div className="tr-muted">Loading…</div> : (
          <>
            {tab === "people" && <PeopleTab data={data} maps={maps} reload={load} me={profile} onAssign={goAssign} />}
            {tab === "assign" && <AssignTab key={assignFor || "all"} data={data} maps={maps} reload={load} preselect={assignFor} />}
            {tab === "library" && <LibraryTab data={data} maps={maps} reload={load} />}
          </>
        )}
      </main>
    </>
  );
}

/* ---------- People ---------- */

function PeopleTab({ data, maps, reload, me, onAssign }) {
  const [q, setQ] = useState("");
  const [showInactive, setShowInactive] = useState(false);
  const [onlyOpen, setOnlyOpen] = useState(false);
  const [openId, setOpenId] = useState(null);
  const [copied, setCopied] = useState(false);

  const person = openId && data.people.find((p) => p.id === openId);
  if (person) return <PersonDetail person={person} data={data} maps={maps} reload={reload} me={me} onBack={() => setOpenId(null)} onAssign={onAssign} />;

  const needle = q.trim().toLowerCase();
  const rows = data.people
    .filter((p) => showInactive || p.active)
    .filter((p) => !needle || [p.full_name, p.email, p.division].some((s) => (s || "").toLowerCase().includes(needle)))
    .map((p) => ({ p, s: summarize(p.id, maps) }))
    .filter((r) => !onlyOpen || r.s.done < r.s.total);

  function exportCSV() {
    const out = [["Name", "Email", "Division", "Hire date", "Video", "Category", "Assigned", "Due", "Status", "Watched %", "Best quiz score", "Quiz attempts", "Completed"]];
    for (const p of data.people) {
      for (const a of maps.asgByUser[p.id] || []) {
        const v = maps.videoById[a.video_id];
        if (!v) continue;
        const pr = maps.prog[pkey(p.id, v.id)];
        const st = statusOf(v, pr);
        out.push([p.full_name, p.email, p.division, p.hire_date, v.title, maps.catById[v.category_id]?.name || "",
          a.assigned_at?.slice(0, 10), a.due_date, isLate(st, a.due_date) ? "Overdue" : STATUS_LABEL[st],
          watchPct(v, pr), pr?.quiz_score ?? "", pr?.quiz_attempts ?? 0, pr?.completed_at?.slice(0, 10) || ""]);
      }
    }
    downloadCSV(`aime-training-${todayISO()}.csv`, out);
  }
  async function copyLink() {
    try { await navigator.clipboard.writeText(window.location.origin); setCopied(true); setTimeout(() => setCopied(false), 2000); } catch { /* ignore */ }
  }

  return (
    <div className="tr-stack">
      <div className="tr-row">
        <h1>People</h1>
        <span className="tr-spacer" />
        <button className="tr-btn ghost" onClick={exportCSV}>Export to CSV</button>
      </div>
      <div className="tr-card tr-row">
        <div className="tr-item-main">
          New hires create their own account at <strong>{window.location.origin}</strong> with their work email.
          They show up here as trainees, and then you assign their videos.
        </div>
        <button className="tr-btn ghost sm" onClick={copyLink}>{copied ? "Link copied" : "Copy link"}</button>
      </div>
      <div className="tr-row">
        <input type="search" placeholder="Search by name, email or division" value={q} onChange={(e) => setQ(e.target.value)} style={{ maxWidth: 340 }} />
        <label className="tr-check"><input type="checkbox" checked={onlyOpen} onChange={(e) => setOnlyOpen(e.target.checked)} />Unfinished training only</label>
        <label className="tr-check"><input type="checkbox" checked={showInactive} onChange={(e) => setShowInactive(e.target.checked)} />Show inactive</label>
      </div>
      <div className="tr-tablewrap">
        <table className="tr-table">
          <thead><tr><th>Name</th><th>Division</th><th>Hired</th><th>Progress</th><th className="tr-num">Overdue</th><th>Last activity</th><th>Role</th></tr></thead>
          <tbody>
            {rows.length === 0 && <tr><td colSpan={7} className="tr-muted">No one matches. Clear the search or filters.</td></tr>}
            {rows.map(({ p, s }) => (
              <tr key={p.id} className="click" onClick={() => setOpenId(p.id)} tabIndex={0} onKeyDown={(e) => { if (e.key === "Enter") setOpenId(p.id); }}>
                <td><div style={{ fontWeight: 600 }}>{p.full_name || "—"}{!p.active && <span className="tr-chip" style={{ marginLeft: 8 }}>Inactive</span>}</div><div className="tr-small tr-muted">{p.email}</div></td>
                <td>{p.division || "—"}</td>
                <td>{fmtDate(p.hire_date)}</td>
                <td style={{ minWidth: 170 }}>
                  {s.total === 0 ? <span className="tr-muted tr-small">Nothing assigned</span> : (
                    <div className="tr-row" style={{ flexWrap: "nowrap" }}>
                      <div className="tr-bar" style={{ flex: 1 }}><i style={{ width: `${(100 * s.done) / s.total}%` }} /></div>
                      <span className="tr-small">{s.done} of {s.total}</span>
                    </div>
                  )}
                </td>
                <td className="tr-num">{s.overdue ? <span className="tr-chip late">{s.overdue}</span> : "—"}</td>
                <td>{fmtDate(s.last)}</td>
                <td>{p.role === "admin" ? "Admin" : "Trainee"}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}

function PersonDetail({ person, data, maps, reload, me, onBack, onAssign }) {
  const [form, setForm] = useState({
    full_name: person.full_name || "", division: person.division || "", hire_date: person.hire_date || "",
    role: person.role, active: person.active,
  });
  const [msg, setMsg] = useState(null);
  const [busy, setBusy] = useState(false);
  const set = (k) => (e) => setForm((f) => ({ ...f, [k]: e.target.type === "checkbox" ? e.target.checked : e.target.value }));

  async function save() {
    const self = person.id === me.id;
    if (self && (form.role !== "admin" || !form.active) &&
      !window.confirm("This removes your own admin access. You won't be able to manage training after this. Continue?")) return;
    setBusy(true);
    const { error } = await supabase.from("training_profiles").update({
      full_name: form.full_name.trim() || null, division: form.division.trim() || null,
      hire_date: form.hire_date || null, role: form.role, active: form.active,
    }).eq("id", person.id);
    setBusy(false);
    if (error) { setMsg({ type: "err", text: error.message }); return; }
    if (self && (form.role !== "admin" || !form.active)) { window.location.reload(); return; }
    setMsg({ type: "ok", text: "Saved." });
    reload();
  }

  const rows = (maps.asgByUser[person.id] || [])
    .map((a) => ({ a, v: maps.videoById[a.video_id], p: maps.prog[pkey(person.id, a.video_id)] }))
    .filter((r) => r.v)
    .sort((x, y) => x.v.title.localeCompare(y.v.title));

  async function reset(r) {
    if (!window.confirm(`Reset "${r.v.title}" for ${person.full_name}? They'll need to watch it and pass the quiz again.`)) return;
    const { error } = await supabase.from("training_progress").delete().eq("user_id", person.id).eq("video_id", r.v.id);
    if (error) setMsg({ type: "err", text: error.message });
    reload();
  }
  async function unassign(r) {
    if (!window.confirm(`Remove "${r.v.title}" from ${person.full_name}'s training?`)) return;
    const { error } = await supabase.from("training_assignments").delete().eq("id", r.a.id);
    if (error) setMsg({ type: "err", text: error.message });
    reload();
  }

  const s = summarize(person.id, maps);
  return (
    <div className="tr-stack">
      <div><button className="tr-link" onClick={onBack}>Back to people</button></div>
      <div className="tr-row">
        <div className="tr-item-main">
          <h1>{person.full_name || person.email}</h1>
          <div className="tr-muted">{person.email}. {s.total ? `${s.done} of ${s.total} complete.` : "Nothing assigned yet."}</div>
        </div>
        <button className="tr-btn go" onClick={() => onAssign(person.id)}>Assign videos</button>
      </div>

      <div className="tr-card tr-stack">
        <h3>Details</h3>
        <div className="tr-grid">
          <label className="tr-field"><span>Full name</span><input type="text" value={form.full_name} onChange={set("full_name")} /></label>
          <label className="tr-field"><span>Division</span>
            <input type="text" list="tr-divisions" value={form.division} onChange={set("division")} placeholder="Mechanical, Pipeline, Structural, Manufacturing" />
            <datalist id="tr-divisions">
              {["Mechanical", "Pipeline", "Structural", "Manufacturing", "Office"].map((d) => <option key={d} value={d} />)}
            </datalist>
          </label>
          <label className="tr-field"><span>Hire date</span><input type="date" value={form.hire_date} onChange={set("hire_date")} /></label>
          <label className="tr-field"><span>Role</span>
            <select value={form.role} onChange={set("role")}>
              <option value="trainee">Trainee (watches assigned videos)</option>
              <option value="admin">Admin (manages videos and assignments)</option>
            </select>
          </label>
        </div>
        <div className="tr-row">
          <label className="tr-check"><input type="checkbox" checked={form.active} onChange={set("active")} />Active</label>
          <span className="tr-spacer" />
          <button className="tr-btn" onClick={save} disabled={busy}>{busy ? "Saving…" : "Save details"}</button>
        </div>
        <Msg msg={msg} />
      </div>

      <h3>Assigned videos</h3>
      {rows.length === 0 ? <div className="tr-card tr-muted">No videos assigned. Use “Assign videos” to pick what {person.full_name || "they"} should watch.</div> : (
        <div className="tr-tablewrap">
          <table className="tr-table">
            <thead><tr><th>Video</th><th>Due</th><th>Status</th><th>Watched</th><th className="tr-num">Best quiz</th><th className="tr-num">Attempts</th><th>Completed</th><th /></tr></thead>
            <tbody>
              {rows.map((r) => {
                const st = statusOf(r.v, r.p);
                return (
                  <tr key={r.a.id}>
                    <td><div style={{ fontWeight: 600 }}>{r.v.title}</div><div className="tr-small tr-muted">{maps.catById[r.v.category_id]?.name || "Uncategorized"}</div></td>
                    <td>{fmtDate(r.a.due_date)}</td>
                    <td><StatusChip st={st} late={isLate(st, r.a.due_date)} /></td>
                    <td style={{ minWidth: 110 }}><div className="tr-bar"><i style={{ width: `${watchPct(r.v, r.p)}%` }} /></div></td>
                    <td className="tr-num">{r.v.quiz_count ? (r.p?.quiz_score != null ? `${r.p.quiz_score}%` : "—") : "No quiz"}</td>
                    <td className="tr-num">{r.p?.quiz_attempts || 0}</td>
                    <td>{fmtDate(r.p?.completed_at)}</td>
                    <td><div className="tr-row" style={{ flexWrap: "nowrap" }}>
                      {r.p && <button className="tr-link" onClick={() => reset(r)}>Reset</button>}
                      <button className="tr-link danger" onClick={() => unassign(r)}>Remove</button>
                    </div></td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      )}
    </div>
  );
}

/* ---------- Assign ---------- */

function AssignTab({ data, maps, reload, preselect }) {
  const [selPeople, setSelPeople] = useState(() => new Set(preselect ? [preselect] : []));
  const [selVideos, setSelVideos] = useState(() => new Set());
  const [q, setQ] = useState("");
  const [trainOnly, setTrainOnly] = useState(true);
  const [due, setDue] = useState("");
  const [msg, setMsg] = useState(null);
  const [busy, setBusy] = useState(false);

  const needle = q.trim().toLowerCase();
  const people = data.people.filter((p) => p.active && (!trainOnly || p.role === "trainee" || selPeople.has(p.id)))
    .filter((p) => !needle || [p.full_name, p.email, p.division].some((s) => (s || "").toLowerCase().includes(needle)));
  const groups = groupVideos(data.videos.filter((v) => v.storage_path), data.categories);
  const single = selPeople.size === 1 ? [...selPeople][0] : null;
  const singlePerson = single && data.people.find((p) => p.id === single);

  const toggle = (setter, id) => setter((s) => { const n = new Set(s); if (n.has(id)) n.delete(id); else n.add(id); return n; });
  function toggleGroup(g) {
    const all = g.videos.every((v) => selVideos.has(v.id));
    setSelVideos((s) => { const n = new Set(s); g.videos.forEach((v) => (all ? n.delete(v.id) : n.add(v.id))); return n; });
  }

  async function assign() {
    const rows = [];
    for (const u of selPeople) for (const v of selVideos) rows.push(due ? { user_id: u, video_id: v, due_date: due } : { user_id: u, video_id: v });
    setBusy(true); setMsg(null);
    const { error } = await supabase.from("training_assignments").upsert(rows, { onConflict: "user_id,video_id", ignoreDuplicates: !due });
    setBusy(false);
    if (error) { setMsg({ type: "err", text: error.message }); return; }
    setMsg({ type: "ok", text: `Assigned ${selVideos.size} video${selVideos.size === 1 ? "" : "s"} to ${selPeople.size} ${selPeople.size === 1 ? "person" : "people"}.` });
    setSelVideos(new Set());
    reload();
  }
  async function unassign(videoId) {
    const a = maps.asg[pkey(single, videoId)];
    if (!a || !window.confirm(`Remove this video from ${singlePerson?.full_name || "this person"}'s training?`)) return;
    const { error } = await supabase.from("training_assignments").delete().eq("id", a.id);
    if (error) setMsg({ type: "err", text: error.message });
    reload();
  }

  if (data.videos.length === 0) {
    return <div className="tr-stack"><h1>Assign videos</h1><div className="tr-card">Add videos in the Video library first, then come back here to assign them.</div></div>;
  }

  return (
    <div className="tr-stack">
      <div>
        <h1>Assign videos</h1>
        <p className="tr-muted">Pick one or more people, then the videos they should watch.</p>
      </div>
      <div className="tr-split">
        <div className="tr-card tr-pick">
          <div className="tr-pickhead tr-stack" style={{ gap: 8 }}>
            <div className="tr-row"><strong>People</strong><span className="tr-spacer" /><span className="tr-small tr-muted">{selPeople.size} selected</span></div>
            <input type="search" placeholder="Search" value={q} onChange={(e) => setQ(e.target.value)} />
            <label className="tr-check tr-small"><input type="checkbox" checked={trainOnly} onChange={(e) => setTrainOnly(e.target.checked)} />Trainees only</label>
          </div>
          {people.map((p) => {
            const s = summarize(p.id, maps);
            return (
              <label key={p.id} className="tr-pickrow" style={{ cursor: "pointer" }}>
                <input type="checkbox" checked={selPeople.has(p.id)} onChange={() => toggle(setSelPeople, p.id)} style={{ width: 18, height: 18, accentColor: "var(--steel)" }} />
                <span className="tr-item-main">
                  <span style={{ fontWeight: 600 }}>{p.full_name || p.email}</span>
                  <span className="tr-small tr-muted" style={{ display: "block" }}>{p.division || "No division"}{s.total ? `, ${s.done} of ${s.total} done` : ""}</span>
                </span>
              </label>
            );
          })}
          {people.length === 0 && <div className="tr-pickrow tr-muted">No one matches.</div>}
        </div>

        <div className="tr-card tr-pick">
          <div className="tr-pickhead tr-row">
            <strong>Videos</strong>
            {singlePerson && <span className="tr-small tr-muted">Showing {singlePerson.full_name}'s status</span>}
            <span className="tr-spacer" />
            <span className="tr-small tr-muted">{selVideos.size} selected</span>
          </div>
          {groups.map((g) => {
            const all = g.videos.every((v) => selVideos.has(v.id));
            return (
              <div key={g.id}>
                <div className="tr-catrow">
                  <span className="tr-item-main">{g.name}</span>
                  <button className="tr-link tr-small" onClick={() => toggleGroup(g)}>{all ? "Clear" : "Select all"}</button>
                </div>
                {g.videos.map((v) => {
                  const a = single && maps.asg[pkey(single, v.id)];
                  const st = a && statusOf(v, maps.prog[pkey(single, v.id)]);
                  return (
                    <label key={v.id} className="tr-pickrow" style={{ cursor: "pointer" }}>
                      <input type="checkbox" checked={selVideos.has(v.id)} onChange={() => toggle(setSelVideos, v.id)} style={{ width: 18, height: 18, accentColor: "var(--steel)" }} />
                      <span className="tr-item-main">
                        <span style={{ fontWeight: 600 }}>{v.title}</span>
                        <span className="tr-small tr-muted" style={{ display: "block" }}>
                          {fmtDur(v.duration_seconds)}{v.quiz_count ? `, ${v.quiz_count}-question quiz` : ", no quiz"}{!v.published ? ", hidden" : ""}
                        </span>
                      </span>
                      {a && <StatusChip st={st} late={isLate(st, a.due_date)} />}
                      {a && <button type="button" className="tr-link danger tr-small" onClick={(e) => { e.preventDefault(); unassign(v.id); }}>Remove</button>}
                    </label>
                  );
                })}
              </div>
            );
          })}
        </div>
      </div>

      <div className="tr-sticky">
        <span>{selVideos.size} video{selVideos.size === 1 ? "" : "s"} for {selPeople.size} {selPeople.size === 1 ? "person" : "people"}</span>
        <span className="tr-spacer" />
        <label className="tr-row" style={{ gap: 6 }}><span className="tr-small">Due date (optional)</span>
          <input type="date" value={due} min={todayISO()} onChange={(e) => setDue(e.target.value)} />
        </label>
        <button className="tr-btn go" disabled={busy || !selPeople.size || !selVideos.size} onClick={assign}>{busy ? "Assigning…" : "Assign"}</button>
      </div>
      <Msg msg={msg} />
    </div>
  );
}

/* ---------- Library ---------- */

function LibraryTab({ data, maps, reload }) {
  const [editing, setEditing] = useState(null); // "new" | video
  const [newCat, setNewCat] = useState("");
  const [msg, setMsg] = useState(null);

  if (editing) {
    return <VideoEditor video={editing === "new" ? null : editing} categories={data.categories}
      onClose={(changed) => { setEditing(null); if (changed) reload(); }} />;
  }

  async function addCategory(e) {
    e.preventDefault();
    const name = newCat.trim();
    if (!name) return;
    const { error } = await supabase.from("training_categories").insert({ name, sort_order: data.categories.length });
    if (error) setMsg({ type: "err", text: error.code === "23505" ? `There's already a category called "${name}".` : error.message });
    else { setNewCat(""); setMsg(null); reload(); }
  }
  async function removeCategory(c) {
    if (!window.confirm(`Delete the "${c.name}" category? Its videos move to Uncategorized.`)) return;
    const { error } = await supabase.from("training_categories").delete().eq("id", c.id);
    if (error) setMsg({ type: "err", text: error.message });
    reload();
  }

  const groups = groupVideos(data.videos, data.categories);

  return (
    <div className="tr-stack">
      <div className="tr-row">
        <h1>Video library</h1>
        <span className="tr-spacer" />
        <button className="tr-btn go" onClick={() => setEditing("new")}>Add video</button>
      </div>

      <div className="tr-card tr-stack" style={{ gap: 12 }}>
        <h3>Categories</h3>
        <div className="tr-catchips">
          {data.categories.length === 0 && <span className="tr-muted">No categories yet. Try Safety, Pipeline, Shop, or Company policies.</span>}
          {data.categories.map((c) => (
            <span key={c.id} className="tr-catchip">{c.name}<button aria-label={`Delete ${c.name}`} onClick={() => removeCategory(c)}>×</button></span>
          ))}
        </div>
        <form className="tr-row" onSubmit={addCategory}>
          <input type="text" placeholder="New category name" value={newCat} onChange={(e) => setNewCat(e.target.value)} style={{ maxWidth: 280 }} />
          <button className="tr-btn ghost">Add category</button>
        </form>
        <Msg msg={msg} />
      </div>

      {data.videos.length === 0 && <div className="tr-card">No videos yet. Use “Add video” to upload your first training video and write its quiz.</div>}

      {groups.map((g) => (
        <section key={g.id} className="tr-stack" style={{ gap: 8 }}>
          <h3>{g.name}</h3>
          <div className="tr-tablewrap">
            <table className="tr-table">
              <thead><tr><th>Video</th><th className="tr-num">Length</th><th className="tr-num">Quiz</th><th className="tr-num">Assigned</th><th className="tr-num">Completed</th><th /></tr></thead>
              <tbody>
                {g.videos.map((v) => {
                  const asg = maps.asgByVideo[v.id] || [];
                  const doneCount = asg.filter((a) => statusOf(v, maps.prog[pkey(a.user_id, v.id)]) === "done").length;
                  return (
                    <tr key={v.id} className="click" onClick={() => setEditing(v)}>
                      <td>
                        <div style={{ fontWeight: 600 }}>{v.title}</div>
                        <div className="tr-row" style={{ gap: 6, marginTop: 2 }}>
                          {!v.storage_path && <span className="tr-chip late">No video file</span>}
                          {!v.published && <span className="tr-chip">Hidden</span>}
                          {v.storage_path && <span className="tr-small tr-muted">{fmtSize(v.file_size)}</span>}
                        </div>
                      </td>
                      <td className="tr-num">{fmtDur(v.duration_seconds)}</td>
                      <td className="tr-num">{v.quiz_count ? `${v.quiz_count} q` : "—"}</td>
                      <td className="tr-num">{asg.length}</td>
                      <td className="tr-num">{doneCount}</td>
                      <td><button className="tr-link" onClick={(e) => { e.stopPropagation(); setEditing(v); }}>Edit</button></td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        </section>
      ))}
    </div>
  );
}

let qKey = 0;
const blankQuestion = () => ({ key: ++qKey, prompt: "", options: ["", ""], correct: 0 });

function VideoEditor({ video, categories, onClose }) {
  const [title, setTitle] = useState(video?.title || "");
  const [description, setDescription] = useState(video?.description || "");
  const [categoryId, setCategoryId] = useState(video?.category_id || "");
  const [passScore, setPassScore] = useState(video?.pass_score ?? 80);
  const [requireFull, setRequireFull] = useState(video?.require_full_watch ?? true);
  const [published, setPublished] = useState(video?.published ?? true);
  const [file, setFile] = useState(null);
  const [questions, setQuestions] = useState(video ? null : []);
  const [quizDirty, setQuizDirty] = useState(false);
  const [previewUrl, setPreviewUrl] = useState(null);
  const [busy, setBusy] = useState(false);
  const [stage, setStage] = useState("");
  const [upPct, setUpPct] = useState(0);
  const [err, setErr] = useState("");
  const savedId = useRef(video?.id || null);
  const changed = useRef(false);

  useEffect(() => {
    if (!video) return;
    supabase.from("training_quiz_questions").select("*").eq("video_id", video.id).order("sort_order")
      .then(({ data, error }) => {
        if (error) { setErr(error.message); setQuestions([]); return; }
        setQuestions(data.map((q) => ({ key: ++qKey, prompt: q.prompt, options: q.options, correct: q.correct_index })));
      });
    if (video.storage_path) {
      supabase.storage.from(BUCKET).createSignedUrl(video.storage_path, 3600).then(({ data }) => data && setPreviewUrl(data.signedUrl));
    }
  }, [video]);

  const editQ = (i, patch) => { setQuizDirty(true); setQuestions((qs) => qs.map((q, j) => (j === i ? { ...q, ...patch } : q))); };
  const editOpt = (i, oi, val) => editQ(i, { options: questions[i].options.map((o, k) => (k === oi ? val : o)) });
  const addOpt = (i) => editQ(i, { options: [...questions[i].options, ""] });
  const removeOpt = (i, oi) => {
    const q = questions[i];
    const options = q.options.filter((_, k) => k !== oi);
    const correct = q.correct === oi ? 0 : q.correct > oi ? q.correct - 1 : q.correct;
    editQ(i, { options, correct });
  };
  const addQ = () => { setQuizDirty(true); setQuestions((qs) => [...qs, blankQuestion()]); };
  const removeQ = (i) => { setQuizDirty(true); setQuestions((qs) => qs.filter((_, j) => j !== i)); };
  const moveQ = (i, d) => {
    setQuizDirty(true);
    setQuestions((qs) => { const n = [...qs]; const [q] = n.splice(i, 1); n.splice(i + d, 0, q); return n; });
  };

  function validate() {
    if (!title.trim()) return "Give the video a title.";
    if (!savedId.current && !file && !video?.storage_path) return "Choose a video file to upload.";
    const p = Number(passScore);
    if (!(p >= 0 && p <= 100)) return "Passing score must be between 0 and 100.";
    for (const [i, q] of (questions || []).entries()) {
      if (!q.prompt.trim()) return `Question ${i + 1} needs a question.`;
      const opts = q.options.map((o) => o.trim());
      if (opts.length < 2 || opts.some((o) => !o)) return `Question ${i + 1}: fill in every answer, at least two.`;
    }
    return "";
  }

  async function save() {
    const v = validate();
    if (v) { setErr(v); return; }
    setErr(""); setBusy(true);
    try {
      const fields = {
        title: title.trim(), description: description.trim() || null, category_id: categoryId || null,
        pass_score: Number(passScore), require_full_watch: requireFull, published, updated_at: new Date().toISOString(),
      };
      setStage("Saving details…");
      if (!savedId.current) {
        const { data, error } = await supabase.from("training_videos").insert(fields).select().single();
        if (error) throw error;
        savedId.current = data.id;
      } else {
        const { error } = await supabase.from("training_videos").update(fields).eq("id", savedId.current);
        if (error) throw error;
      }
      changed.current = true;
      const id = savedId.current;

      if (questions && (quizDirty || !video)) {
        setStage("Saving quiz…");
        const { error: dErr } = await supabase.from("training_quiz_questions").delete().eq("video_id", id);
        if (dErr) throw dErr;
        if (questions.length) {
          const rows = questions.map((q, i) => ({
            video_id: id, prompt: q.prompt.trim(), options: q.options.map((o) => o.trim()), correct_index: q.correct, sort_order: i,
          }));
          const { error } = await supabase.from("training_quiz_questions").insert(rows);
          if (error) throw error;
        }
        setQuizDirty(false);
      }

      if (file) {
        const path = `${id}/${Date.now()}-${safeName(file.name)}`;
        setStage("Reading video…");
        const duration = await getVideoDuration(file);
        setStage("Uploading video…");
        setUpPct(0);
        await uploadResumable(path, file, setUpPct);
        const { data: cur } = await supabase.from("training_videos").select("storage_path").eq("id", id).single();
        const { error } = await supabase.from("training_videos")
          .update({ storage_path: path, file_name: file.name, file_size: file.size, duration_seconds: duration }).eq("id", id);
        if (error) throw error;
        if (cur?.storage_path && cur.storage_path !== path) await supabase.storage.from(BUCKET).remove([cur.storage_path]);
        setFile(null);
      }
      onClose(true);
    } catch (e) {
      setErr((e?.message || String(e)) + (savedId.current ? " Your other changes were saved; press Save again to retry." : ""));
    } finally {
      setBusy(false); setStage("");
    }
  }

  async function remove() {
    if (!window.confirm(`Delete "${video.title}"? It's removed from everyone's training, along with its quiz and progress records.`)) return;
    setBusy(true);
    try {
      if (video.storage_path) await supabase.storage.from(BUCKET).remove([video.storage_path]);
      const { error } = await supabase.from("training_videos").delete().eq("id", video.id);
      if (error) throw error;
      onClose(true);
    } catch (e) {
      setErr(e.message); setBusy(false);
    }
  }

  return (
    <div className="tr-stack" style={{ maxWidth: 900 }}>
      <div><button className="tr-link" onClick={() => onClose(changed.current)} disabled={busy}>Back to library</button></div>
      <h1>{video ? "Edit video" : "Add video"}</h1>

      <div className="tr-card tr-stack">
        <label className="tr-field"><span>Title</span><input type="text" value={title} onChange={(e) => setTitle(e.target.value)} placeholder="Confined space entry basics" /></label>
        <label className="tr-field"><span>Description (shown under the video)</span><textarea value={description} onChange={(e) => setDescription(e.target.value)} /></label>
        <div className="tr-grid">
          <label className="tr-field"><span>Category</span>
            <select value={categoryId} onChange={(e) => setCategoryId(e.target.value)}>
              <option value="">Uncategorized</option>
              {categories.map((c) => <option key={c.id} value={c.id}>{c.name}</option>)}
            </select>
          </label>
          <label className="tr-field"><span>Passing score (%)</span>
            <input type="number" min={0} max={100} value={passScore} onChange={(e) => setPassScore(e.target.value)} />
          </label>
        </div>
        <div className="tr-stack" style={{ gap: 8 }}>
          <label className="tr-check"><input type="checkbox" checked={requireFull} onChange={(e) => setRequireFull(e.target.checked)} />Don't let trainees skip ahead the first time they watch</label>
          <label className="tr-check"><input type="checkbox" checked={published} onChange={(e) => setPublished(e.target.checked)} />Visible to assigned trainees</label>
        </div>
      </div>

      <div className="tr-card tr-stack">
        <h3>Video file</h3>
        {previewUrl && !file && (
          <div className="tr-player" style={{ maxWidth: 560 }}><video src={previewUrl} controls preload="metadata" /></div>
        )}
        {video?.file_name && !file && <div className="tr-small tr-muted">{video.file_name}, {fmtSize(video.file_size)}, {fmtDur(video.duration_seconds)}</div>}
        <label className="tr-field"><span>{video?.storage_path ? "Replace with a new file" : "Choose a video file"} (MP4 works best)</span>
          <input type="file" accept="video/mp4,video/quicktime,video/webm,video/x-m4v" onChange={(e) => setFile(e.target.files?.[0] || null)} disabled={busy} />
        </label>
        {file && <div className="tr-small">{file.name}, {fmtSize(file.size)}</div>}
        {stage === "Uploading video…" && (
          <div className="tr-row">
            <div className="tr-bar" style={{ flex: 1, height: 10 }}><i style={{ width: `${Math.round(upPct * 100)}%` }} /></div>
            <span className="tr-small">{Math.round(upPct * 100)}%</span>
          </div>
        )}
      </div>

      <div className="tr-card tr-stack">
        <div className="tr-row">
          <h3>Quiz</h3>
          <span className="tr-small tr-muted">Trainees take it after watching. Leave it empty if this video doesn't need one.</span>
        </div>
        {questions === null ? <div className="tr-muted">Loading quiz…</div> : (
          <>
            {questions.map((q, i) => (
              <div key={q.key} className="tr-qedit tr-stack" style={{ gap: 8 }}>
                <div className="tr-row">
                  <strong>Question {i + 1}</strong>
                  <span className="tr-spacer" />
                  <button className="tr-link tr-small" disabled={i === 0} onClick={() => moveQ(i, -1)}>Move up</button>
                  <button className="tr-link tr-small" disabled={i === questions.length - 1} onClick={() => moveQ(i, 1)}>Move down</button>
                  <button className="tr-link danger tr-small" onClick={() => removeQ(i)}>Delete</button>
                </div>
                <input type="text" value={q.prompt} onChange={(e) => editQ(i, { prompt: e.target.value })} placeholder="What must you check before entering a confined space?" aria-label={`Question ${i + 1}`} />
                <div className="tr-small tr-muted">Answers. Select the correct one.</div>
                {q.options.map((o, oi) => (
                  <div key={oi} className="tr-optedit">
                    <input type="radio" name={`correct-${q.key}`} checked={q.correct === oi} onChange={() => editQ(i, { correct: oi })} aria-label={`Answer ${oi + 1} is correct`} style={{ width: 18, height: 18, accentColor: "var(--go)" }} />
                    <input type="text" value={o} onChange={(e) => editOpt(i, oi, e.target.value)} placeholder={`Answer ${oi + 1}`} />
                    <button className="tr-link danger tr-small" disabled={q.options.length <= 2} onClick={() => removeOpt(i, oi)}>Remove</button>
                  </div>
                ))}
                {q.options.length < 6 && <div><button className="tr-link tr-small" onClick={() => addOpt(i)}>Add an answer</button></div>}
              </div>
            ))}
            <div><button className="tr-btn ghost" onClick={addQ}>Add question</button></div>
          </>
        )}
      </div>

      {err && <div className="tr-msg err" role="alert">{err}</div>}
      <div className="tr-row">
        <button className="tr-btn go" onClick={save} disabled={busy || questions === null}>{busy ? stage || "Saving…" : video ? "Save changes" : "Save video"}</button>
        <button className="tr-btn ghost" onClick={() => onClose(changed.current)} disabled={busy}>Cancel</button>
        <span className="tr-spacer" />
        {video && <button className="tr-btn danger" onClick={remove} disabled={busy}>Delete video</button>}
      </div>
      {busy && stage === "Uploading video…" && <div className="tr-msg warn">Keep this tab open until the upload finishes.</div>}
    </div>
  );
}
