// ============================================================
// SheetMind PRO v15 — PREMIUM UI + TEMPLATE SYSTEM
// Updated: Save/Update logic, Free 4-table+24h cooldown,
//          bKash Premium with Admin Approval flow
// ============================================================
import { initializeApp } from "firebase/app";
import {
  getAuth,
  GoogleAuthProvider,
  signInWithPopup,
  signOut,
  onAuthStateChanged,
} from "firebase/auth";
import {
  getFirestore,
  doc,
  setDoc,
  getDoc,
  collection,
  onSnapshot,
  serverTimestamp,
  getDocs,
  increment,
} from "firebase/firestore";
import { jsPDF } from "jspdf";
import { autoTable } from "jspdf-autotable";
import * as XLSX from "xlsx";
import React, { useState, useEffect, useRef, useCallback, useMemo } from "react";

// ============================================================
// FIREBASE CONFIG
// ============================================================
const firebaseConfig = {
  apiKey: "AIzaSyBFqaj-vl_W0sBAde_-XdbJPtKYI_tA3Wk",
  authDomain: "smart-sheet-pro.firebaseapp.com",
  projectId: "smart-sheet-pro",
  storageBucket: "smart-sheet-pro.firebasestorage.app",
  messagingSenderId: "1005248265803",
  appId: "1:1005248265803:web:5b173115e8198633cdd7ce",
  measurementId: "G-NCY4VFHVV5",
};

// ============================================================
// CONSTANTS
// ============================================================
const ADMIN_EMAIL = "iftia5061@gmail.com";
const FREE_TABLE_LIMIT = 4;
const BKASH_NUMBER = "01825453585";
const LEMON_WEB_LINK = "https://iftiahamed.lemonsqueezy.com/checkout/buy/47f6d218-c4cc-4aa6-a33d-e511ab24f8cf";
const LEMON_APP_LINK = "https://iftiahamed.lemonsqueezy.com/checkout/buy/70cb73d8-2986-4247-8e7a-6a122828592b";
const AUTOSAVE_DELAY = 1500;
const AI_FREE_TABLE_LIMIT = 2;
const AI_TABLE_ENDPOINT = "/api/ai-table";
const ENABLE_AI_TABLE = true;
// ↓ REPLACE with your real AdSense Publisher ID
// Line 51-57 এ এই অংশটা replace করো:
const ADSENSE_PUB_ID = "ca-pub-5628645711343874";
const AD_SLOTS = {
  sidebar: "2562908490",
  banner:  "7361245890",
  inFeed:  "4128956703",
};

// ============================================================
// FIREBASE INIT
// ============================================================
let app, auth, db, provider;
try {
  app = initializeApp(firebaseConfig);
  auth = getAuth(app);
  db = getFirestore(app);
  provider = new GoogleAuthProvider();
} catch (e) {
  console.error("Firebase init error:", e);
}

// ============================================================
// CELL OBJECT HELPERS
// ============================================================
const makeCellObj = (value = "", overrides = {}) => ({
  value: String(value ?? ""),
  color: "",
  bgColor: "",
  fontSize: "",
  fontWeight: "",
  ...overrides,
});
const cellVal = (cell) => {
  if (cell == null) return "";
  if (typeof cell === "object") return String(cell.value ?? "");
  return String(cell);
};
const migrateRow = (row, columns) => {
  const migrated = {};
  for (const col of columns) {
    const existing = row[col];
    if (existing && typeof existing === "object" && "value" in existing) {
      migrated[col] = existing;
    } else {
      migrated[col] = makeCellObj(existing ?? "");
    }
  }
  return migrated;
};
const makeBlankRow = (columns, id) => {
  const row = {};
  for (const col of columns) {
    row[col] = makeCellObj(col === "ID" ? id : "");
  }
  return row;
};

// ============================================================
// TAB FACTORY
// ============================================================
const makeNewTab = (title, columns, rowCount = 8) => {
  const id = Date.now() + Math.random();
  const cols = columns || ["ID", "Name", "Email", "Status"];
  return {
    id,
    title,
    columns: cols,
    rows: Array.from({ length: rowCount }, (_, i) => makeBlankRow(cols, i + 1)),
    columnTypes: {},
    columnFormats: {},
    tabColor: "",
  };
};

// ============================================================
// TEMPLATE DEFINITIONS
// ============================================================
const TEMPLATES = [
  {
    id: "budget",
    icon: "💰",
    name: "Monthly Budget",
    description: "Track income, expenses and savings",
    color: "#10b981",
    columns: ["ID", "Category", "Description", "Budgeted", "Actual", "Variance", "Status"],
    sampleRows: [
      ["1","Income","Salary","","","",""],
      ["2","Income","Freelance","","","",""],
      ["3","Housing","Rent","","","",""],
      ["4","Housing","Utilities","","","",""],
      ["5","Food","Groceries","","","",""],
      ["6","Food","Dining Out","","","",""],
      ["7","Transport","Fuel","","","",""],
      ["8","Entertainment","Streaming","","","",""],
      ["9","Health","Insurance","","","",""],
      ["10","Savings","Emergency Fund","","","",""],
    ],
  },
  {
    id: "inventory",
    icon: "📦",
    name: "Inventory Tracker",
    description: "Manage products, stock and suppliers",
    color: "#6366f1",
    columns: ["ID", "Product Name", "SKU", "Category", "Stock", "Reorder Level", "Unit Price", "Supplier"],
    sampleRows: [
      ["1","Product A","SKU-001","Electronics","","20","",""],
      ["2","Product B","SKU-002","Clothing","","15","",""],
      ["3","Product C","SKU-003","Food","","50","",""],
    ],
  },
  {
    id: "schedule",
    icon: "📅",
    name: "Project Schedule",
    description: "Plan tasks, deadlines and assignees",
    color: "#f59e0b",
    columns: ["ID", "Task", "Assignee", "Start Date", "Due Date", "Priority", "Progress", "Status"],
    sampleRows: [
      ["1","Research","","","","High","0%","Pending"],
      ["2","Design","","","","High","0%","Pending"],
      ["3","Development","","","","Medium","0%","Pending"],
      ["4","Testing","","","","Medium","0%","Pending"],
      ["5","Deployment","","","","Low","0%","Pending"],
    ],
  },
  {
    id: "crm",
    icon: "🤝",
    name: "CRM / Contacts",
    description: "Manage leads, clients and deals",
    color: "#ec4899",
    columns: ["ID", "Name", "Company", "Email", "Phone", "Stage", "Deal Value", "Last Contact"],
    sampleRows: [
      ["1","","","","","Lead","",""],
      ["2","","","","","Prospect","",""],
      ["3","","","","","Customer","",""],
    ],
  },
  {
    id: "payroll",
    icon: "💳",
    name: "Payroll Sheet",
    description: "Salaries, bonuses and deductions",
    color: "#22d3ee",
    columns: ["ID", "Name", "Position", "Basic Salary", "Bonus", "Deductions", "Net Salary", "Status"],
    sampleRows: Array.from({length: 8}, (_, i) => [String(i+1),"","","","","","",""]),
  },
  {
    id: "student",
    icon: "🎓",
    name: "Student Grades",
    description: "Track students, marks and grades",
    color: "#8b5cf6",
    columns: ["ID", "Name", "Roll", "Class", "Math", "Science", "English", "Total", "Grade"],
    sampleRows: Array.from({length: 10}, (_, i) => [String(i+1),"","","","","","","",""]),
  },
  {
    id: "blank",
    icon: "✨",
    name: "Blank Sheet",
    description: "Start fresh with a clean canvas",
    color: "#64748b",
    columns: ["ID", "Name", "Email", "Status"],
    sampleRows: Array.from({length: 8}, (_, i) => [String(i+1),"","",""]),
  },
];

// ============================================================
// REAL GOOGLE ADSENSE COMPONENT
// ============================================================
function AdSense({ slot, format = "auto", style = {}, isDark, label = "Advertisement" }) {
  const adRef = useRef(null);
  useEffect(() => {
    try {
      if (window.adsbygoogle && adRef.current) {
        (window.adsbygoogle = window.adsbygoogle || []).push({});
      }
    } catch (e) {}
  }, [slot]);

  return (
    <div
      className={`relative rounded-2xl overflow-hidden border ${
        isDark ? "border-slate-800 bg-slate-900/40" : "border-slate-200 bg-slate-50"
      }`}
      style={style}
    >
      <span
        className={`absolute top-1.5 left-3 text-[8px] font-black uppercase tracking-widest z-10 ${
          isDark ? "text-slate-600" : "text-slate-400"
        }`}
      >
        {label}
      </span>
      <ins
        ref={adRef}
        className="adsbygoogle"
        style={{ display: "block", ...style }}
        data-ad-client={ADSENSE_PUB_ID}
        data-ad-slot={slot}
        data-ad-format={format}
        data-full-width-responsive="true"
      />
    </div>
  );
}

function SidebarAd({ isDark }) {
  return (
    <AdSense
      slot={AD_SLOTS.sidebar}
      format="vertical"
      isDark={isDark}
      style={{ minHeight: 280, width: "100%", paddingTop: 20 }}
    />
  );
}

function BannerAd({ isDark }) {
  return (
    <AdSense
      slot={AD_SLOTS.banner}
      format="horizontal"
      isDark={isDark}
      style={{ minHeight: 90, width: "100%", paddingTop: 20 }}
    />
  );
}

function InFeedAd({ isDark }) {
  return (
    <AdSense
      slot={AD_SLOTS.inFeed}
      format="fluid"
      isDark={isDark}
      label="Sponsored"
      style={{ minHeight: 120, width: "100%", paddingTop: 18 }}
    />
  );
}

// ============================================================
// TEMPLATE GALLERY MODAL
// ============================================================
function TemplateGallery({ onSelect, onClose, isDark }) {
  const [hovered, setHovered] = useState(null);
  const bg = isDark ? "bg-[#080d1a]" : "bg-slate-50";
  const card = isDark ? "bg-[#0f1929] border-slate-800" : "bg-white border-slate-200";
  const text = isDark ? "text-slate-100" : "text-slate-900";
  const sub = isDark ? "text-slate-500" : "text-slate-400";

  return (
    <div className="fixed inset-0 z-[250] flex items-center justify-center bg-black/80 backdrop-blur-md p-4">
      <div
        className={`${bg} ${text} w-full max-w-3xl rounded-[2rem] border ${
          isDark ? "border-slate-800" : "border-slate-200"
        } shadow-2xl overflow-hidden`}
        style={{ maxHeight: "90vh", display: "flex", flexDirection: "column" }}
      >
        <div className={`px-8 pt-8 pb-6 border-b ${isDark ? "border-slate-800" : "border-slate-200"}`}>
          <div className="flex justify-between items-start">
            <div>
              <h2 className="text-2xl font-black uppercase tracking-tight">
                <span className="bg-gradient-to-r from-indigo-400 to-sky-400 bg-clip-text text-transparent">
                  Template Gallery
                </span>
              </h2>
              <p className={`text-xs mt-1 ${sub}`}>
                Choose a template to start instantly — all pre-configured
              </p>
            </div>
            <button
              onClick={onClose}
              className={`w-9 h-9 rounded-xl flex items-center justify-center transition-all ${
                isDark
                  ? "bg-slate-800 hover:bg-slate-700 text-slate-400"
                  : "bg-slate-100 hover:bg-slate-200 text-slate-600"
              }`}
            >
              ✕
            </button>
          </div>
        </div>

        <div className="overflow-y-auto p-8" style={{ scrollbarWidth: "thin" }}>
          <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 gap-4">
            {TEMPLATES.map((tmpl) => (
              <button
                key={tmpl.id}
                onMouseEnter={() => setHovered(tmpl.id)}
                onMouseLeave={() => setHovered(null)}
                onClick={() => onSelect(tmpl)}
                className={`${card} border rounded-[1.5rem] p-5 text-left transition-all duration-200 group relative overflow-hidden`}
                style={{
                  transform: hovered === tmpl.id ? "translateY(-3px)" : "none",
                  boxShadow: hovered === tmpl.id ? `0 12px 32px ${tmpl.color}22` : "none",
                  borderColor: hovered === tmpl.id ? tmpl.color + "55" : undefined,
                }}
              >
                <div
                  className="absolute top-0 left-0 right-0 h-0.5 opacity-0 group-hover:opacity-100 transition-opacity"
                  style={{ background: `linear-gradient(to right, ${tmpl.color}, ${tmpl.color}44)` }}
                />
                <div
                  className="w-11 h-11 rounded-2xl flex items-center justify-center text-2xl mb-4"
                  style={{ backgroundColor: tmpl.color + "20" }}
                >
                  {tmpl.icon}
                </div>
                <h3 className={`font-black text-sm uppercase tracking-wide ${text} mb-1`}>{tmpl.name}</h3>
                <p className={`text-[11px] ${sub} leading-relaxed`}>{tmpl.description}</p>
                <div className="flex items-center gap-1.5 mt-3">
                  {tmpl.columns.slice(0, 4).map((col, i) => (
                    <span
                      key={i}
                      className="text-[8px] font-black px-2 py-0.5 rounded-full uppercase"
                      style={{ backgroundColor: tmpl.color + "18", color: tmpl.color }}
                    >
                      {col}
                    </span>
                  ))}
                  {tmpl.columns.length > 4 && (
                    <span className={`text-[8px] font-black ${sub}`}>+{tmpl.columns.length - 4}</span>
                  )}
                </div>
              </button>
            ))}
          </div>
        </div>
      </div>
    </div>
  );
}

// ============================================================
// CHART COMPONENTS
// ============================================================
function MiniBarChart({ data, labels, color = "#6366f1" }) {
  const canvasRef = useRef(null);
  useEffect(() => {
    if (!canvasRef.current || !data.length) return;
    const canvas = canvasRef.current;
    const ctx = canvas.getContext("2d");
    const W = canvas.width, H = canvas.height;
    ctx.clearRect(0, 0, W, H);
    const max = Math.max(...data, 1);
    const barW = (W - 20) / data.length - 4;
    data.forEach((val, i) => {
      const h = (val / max) * (H - 30);
      const x = 10 + i * (barW + 4);
      const y = H - 25 - h;
      const grad = ctx.createLinearGradient(x, y, x, H - 25);
      grad.addColorStop(0, color + "ff");
      grad.addColorStop(1, color + "44");
      ctx.fillStyle = grad;
      ctx.beginPath();
      ctx.roundRect(x, y, barW, h, 3);
      ctx.fill();
      if (labels[i]) {
        ctx.fillStyle = "#94a3b8";
        ctx.font = "9px monospace";
        ctx.textAlign = "center";
        ctx.fillText(String(labels[i]).slice(0, 5), x + barW / 2, H - 8);
      }
    });
  }, [data, labels, color]);
  return <canvas ref={canvasRef} width={280} height={120} className="w-full" />;
}

function MiniPieChart({ data, labels }) {
  const canvasRef = useRef(null);
  useEffect(() => {
    const COLORS = ["#6366f1","#22d3ee","#10b981","#f59e0b","#ef4444","#8b5cf6","#ec4899"];
    if (!canvasRef.current || !data.length) return;
    const canvas = canvasRef.current;
    const ctx = canvas.getContext("2d");
    const W = canvas.width, H = canvas.height;
    ctx.clearRect(0, 0, W, H);
    const total = data.reduce((a, b) => a + b, 0);
    if (!total) return;
    let start = -Math.PI / 2;
    const cx = W / 2 - 20, cy = H / 2, r = Math.min(cx, cy) - 10;
    data.forEach((val, i) => {
      const slice = (val / total) * 2 * Math.PI;
      ctx.beginPath();
      ctx.moveTo(cx, cy);
      ctx.arc(cx, cy, r, start, start + slice);
      ctx.closePath();
      ctx.fillStyle = COLORS[i % COLORS.length];
      ctx.fill();
      ctx.strokeStyle = "#0f172a";
      ctx.lineWidth = 2;
      ctx.stroke();
      start += slice;
    });
    labels.slice(0, 5).forEach((lbl, i) => {
      ctx.fillStyle = COLORS[i % COLORS.length];
      ctx.fillRect(W - 85, 10 + i * 18, 10, 10);
      ctx.fillStyle = "#94a3b8";
      ctx.font = "9px monospace";
      ctx.textAlign = "left";
      ctx.fillText(String(lbl).slice(0, 8), W - 70, 19 + i * 18);
    });
  }, [data, labels]);
  return <canvas ref={canvasRef} width={280} height={130} className="w-full" />;
}

// ============================================================
// LANDING PAGE
// ============================================================
function LandingPage({ onLogin, isDark }) {
  const features = [
    { icon: "⚡", title: "AI Generator", desc: "Generate table structures instantly from natural language prompts." },
    { icon: "🧮", title: "Calc Engine", desc: "Real-time formula calculations across columns with filter support." },
    { icon: "📊", title: "Auto Analytics", desc: "Bar & pie charts generated automatically from your numeric data." },
    { icon: "🎨", title: "Smart Formatting", desc: "Per-cell color, font size, bold — isolated per sheet." },
    { icon: "🔐", title: "Secure Auth", desc: "Google sign-in with admin monitoring of all active users." },
    { icon: "📤", title: "Export Anywhere", desc: "Download as PDF or Excel with cell styles preserved." },
  ];
  const base = "bg-[#eef4ff] text-[#172033]";
  const card = "bg-white border-[#c8d7ec] shadow-[0_18px_45px_rgba(43,76,126,0.10)]";
  const sub = "text-[#5f6f89]";

  return (
    <div className={`min-h-screen ${base} font-sans`} style={{ fontFamily: "'DM Mono','Fira Mono',monospace" }}>
      <div className="w-full px-8 pt-4">
        <BannerAd isDark={isDark} />
      </div>

      <nav className="flex justify-between items-center px-8 py-5 border-b border-[#c8d7ec] bg-white/70 backdrop-blur-xl">
        <div className="flex items-center gap-3">
          <div className="w-8 h-8 rounded-xl bg-gradient-to-br from-indigo-500 to-sky-500 flex items-center justify-center">
            <span className="text-white text-xs font-black">SM</span>
          </div>
          <div>
            <span className="text-lg font-black bg-gradient-to-r from-indigo-400 via-sky-400 to-emerald-400 bg-clip-text text-transparent uppercase tracking-tighter">SHEETMIND</span>
            <span className={`text-[9px] font-black ml-2 ${sub} tracking-[0.3em] uppercase`}>PRO v15</span>
          </div>
        </div>
        <button
          onClick={onLogin}
          className="bg-gradient-to-r from-[#2563eb] to-[#0ea5e9] hover:from-[#1d4ed8] hover:to-[#0284c7] text-white px-7 py-3 rounded-xl font-black text-[11px] uppercase tracking-widest transition-all shadow-[0_12px_28px_rgba(37,99,235,0.28)] border border-white/20"
        >
          Sign In
        </button>
      </nav>

      <div className="max-w-5xl mx-auto px-8 pt-14 pb-8 text-center">
        <div className="inline-flex items-center gap-2 px-4 py-1.5 rounded-full border text-[10px] font-black uppercase tracking-widest mb-8 border-[#bfd2ee] bg-white text-[#2457d6] shadow-sm">
          <span className="w-1.5 h-1.5 bg-emerald-400 rounded-full animate-pulse"></span>
          Now with Template Gallery & Real Ads
        </div>
        <p className={`text-xl md:text-2xl max-w-3xl mx-auto mb-10 ${sub} font-bold leading-relaxed`}>Premium AI SHEETMIND and data-entry workspace for teams, freelancers, and business records.</p>
        <button
          onClick={onLogin}
          className="bg-gradient-to-r from-[#2563eb] to-[#0ea5e9] hover:from-[#1d4ed8] hover:to-[#0284c7] text-white px-11 py-4 rounded-xl font-black text-sm uppercase tracking-widest transition-all shadow-[0_18px_40px_rgba(37,99,235,0.28)] border border-white/20"
        >
          Continue with Google
        </button>
      </div>
      {/* Premium Pricing Section */}
      <div className="max-w-6xl mx-auto px-8 pb-12">
        <div className="relative overflow-hidden rounded-[2rem] border border-[#c8d7ec] bg-white shadow-[0_28px_80px_rgba(43,76,126,0.16)]">
          <div className="absolute inset-x-0 top-0 h-1 bg-gradient-to-r from-[#2563eb] via-[#0ea5e9] to-[#14b8a6]" />
          <div className="px-6 md:px-10 pt-9 pb-6 text-center">
            <p className="text-[10px] font-black uppercase tracking-[0.32em] text-[#3978d8] mb-3">
              Pricing Plans
            </p>
            <h2 className="text-3xl md:text-4xl font-black uppercase tracking-tight text-[#172033] mb-3">
              Start Free. Upgrade When Work Grows.
            </h2>
            <p className="text-sm md:text-base text-[#5f6f89] max-w-2xl mx-auto leading-relaxed">
              Built for company data entry, freelancers, client records, inventory, reports, and AI-generated sheets.
            </p>
          </div>

          <div className="grid grid-cols-1 md:grid-cols-3 gap-4 px-6 md:px-10 pb-8">
            <div className="rounded-2xl border border-[#d8e4f3] bg-[#f7fbff] p-6 text-left shadow-sm">
              <p className="text-[10px] font-black uppercase tracking-[0.28em] text-[#5f6f89] mb-4">Free</p>
              <div className="flex items-end gap-1 mb-2">
                <span className="text-4xl font-black text-[#172033]">$0</span>
                <span className="text-xs font-black text-[#5f6f89] mb-1 uppercase">forever</span>
              </div>
              <p className="text-[11px] text-[#5f6f89] mb-5">For testing and small personal work.</p>
              {["4 lifetime tables", "2 AI table generations", "PDF and Excel export", "Google login"].map((f, i) => (
                <p key={i} className="text-[11px] text-[#36516f] font-bold mt-2">{f}</p>
              ))}
              <button onClick={onLogin} className="mt-6 w-full rounded-xl border border-[#3978d8]/30 bg-white px-4 py-3 text-[11px] font-black uppercase tracking-widest text-[#2457d6] transition-all hover:bg-[#eaf2ff]">
                Start Free
              </button>
            </div>

            <div className="relative rounded-2xl border border-[#2563eb]/35 bg-gradient-to-br from-[#0f2f66] via-[#174c94] to-[#0f766e] p-6 text-left text-white shadow-[0_24px_60px_rgba(37,99,235,0.28)]">
              <div className="absolute right-4 top-4 rounded-full bg-white/15 px-3 py-1 text-[8px] font-black uppercase tracking-widest text-white">Best Value</div>
              <p className="text-[10px] font-black uppercase tracking-[0.28em] text-[#bde7ff] mb-4">Pro Monthly</p>
              <div className="flex items-end gap-1 mb-2">
                <span className="text-4xl font-black">$3</span>
                <span className="text-xs font-black text-[#bde7ff] mb-1 uppercase">per month</span>
              </div>
              <p className="text-[11px] text-[#d9f3ff] mb-5">Around 300 BDT/month for serious data work.</p>
              {["Unlimited tables", "Unlimited AI table generation", "No free-plan limit", "Ad-free workspace", "Priority support"].map((f, i) => (
                <p key={i} className="text-[11px] font-bold text-white/90 mt-2">{f}</p>
              ))}
              <a href={LEMON_WEB_LINK} target="_blank" rel="noopener noreferrer" className="mt-6 flex w-full items-center justify-center rounded-xl bg-white px-4 py-3 text-[11px] font-black uppercase tracking-widest text-[#174c94] transition-all hover:bg-[#eaf7ff]">
                Upgrade Monthly
              </a>
            </div>

            <div className="rounded-2xl border border-[#d8e4f3] bg-white p-6 text-left shadow-sm">
              <p className="text-[10px] font-black uppercase tracking-[0.28em] text-[#0f766e] mb-4">App Lifetime</p>
              <div className="flex items-end gap-1 mb-2">
                <span className="text-4xl font-black text-[#172033]">$15</span>
                <span className="text-xs font-black text-[#5f6f89] mb-1 uppercase">one time</span>
              </div>
              <p className="text-[11px] text-[#5f6f89] mb-5">For users who want the mobile app package.</p>
              {["Android APK", "All Pro features", "Offline support", "Free updates", "One-time payment"].map((f, i) => (
                <p key={i} className="text-[11px] text-[#36516f] font-bold mt-2">{f}</p>
              ))}
              <a href={LEMON_APP_LINK} target="_blank" rel="noopener noreferrer" className="mt-6 flex w-full items-center justify-center rounded-xl border border-[#0f766e]/25 bg-[#ecfdf8] px-4 py-3 text-[11px] font-black uppercase tracking-widest text-[#0f766e] transition-all hover:bg-[#d9fff2]">
                Get Lifetime App
              </a>
            </div>
          </div>

          <div className="border-t border-[#d8e4f3] bg-[#f7fbff] px-6 py-4 text-center">
            <p className="text-[10px] font-black uppercase tracking-widest text-[#36516f]">
              Bangladesh users can pay with bKash: 300 BDT/month. Admin approval activates Premium.
            </p>
          </div>
        </div>
      </div>

      <div className="max-w-5xl mx-auto px-8 pb-12 grid grid-cols-1 md:grid-cols-3 gap-5">
        {features.map((f, i) => (
          <div key={i} className={`p-6 rounded-[2rem] border ${card} transition-all hover:-translate-y-1 duration-200`}>
            <div className="text-3xl mb-3">{f.icon}</div>
            <h3 className="font-black text-sm uppercase tracking-wider mb-2">{f.title}</h3>
            <p className={`text-xs leading-relaxed ${sub}`}>{f.desc}</p>
          </div>
        ))}
      </div>
      <div className="max-w-5xl mx-auto px-8 pb-8">
        <BannerAd isDark={isDark} />
      </div>

      <footer className={`border-t ${isDark ? "border-white/5" : "border-slate-200"} py-8 text-center`}>
        <p className={`text-[10px] font-black uppercase tracking-[0.3em] ${sub}`}>SheetMind Pro · v15.0</p>
      </footer>
    </div>
  );
}

// ============================================================
// ADMIN PANEL
// ============================================================

function AdminPanel({ onClose, isDark, currentUser }) {
  const [allUsers,         setAllUsers]         = useState([]);
  const [pendingPayments,  setPendingPayments]  = useState([]);
  const [loading,          setLoading]          = useState(true);
  const [payLoading,       setPayLoading]       = useState(true);
  const [stats,            setStats]            = useState({ totalUsers: 0, premiumUsers: 0, totalTables: 0, totalRevenue: 0 });
  const [activeTab,        setActiveTab]        = useState("overview");
  const [actionMsg,        setActionMsg]        = useState("");
  const [searchUser,       setSearchUser]       = useState("");
  const [filterPlan,       setFilterPlan]       = useState("all");

  // Colors — Deep Blue + Gold theme
  const bg     = isDark ? "bg-[#060e1c]"         : "bg-white";
  const card   = isDark ? "bg-[#08122a] border-yellow-900/20" : "bg-slate-50 border-slate-200";
  const text   = isDark ? "text-slate-100"        : "text-slate-900";
  const sub    = isDark ? "text-blue-400"         : "text-slate-500";
  const border = isDark ? "border-[#d4af3722]"   : "border-blue-200";
  const input  = isDark ? "bg-[#050d1f] border-[#d4af3730] text-slate-200 placeholder:text-blue-900" : "bg-blue-50 border-blue-300 text-slate-800";

  const pendingCount = pendingPayments.filter(p => p.status === "pending").length;

  // ── Fetch Users ──
  const fetchAllUsers = useCallback(async () => {
    if (!db) { setLoading(false); return; }
    try {
      const snap = await getDocs(collection(db, "users"));
      const list = [];
      let premiumCount = 0, tableCount = 0;
      snap.forEach(d => {
        const data = { ...d.data(), uid: d.id };
        list.push(data);
        if (data.isPremium || data.email === ADMIN_EMAIL) premiumCount++;
        tableCount += data.tableCount || 0;
      });
      list.sort((a, b) => (b.lastLogin?.seconds || 0) - (a.lastLogin?.seconds || 0));
      setAllUsers(list);
      setStats(prev => ({ ...prev, totalUsers: snap.size, premiumUsers: premiumCount, totalTables: tableCount }));
    } catch (err) { console.error(err); }
    finally { setLoading(false); }
  }, []);

  // ── Fetch Payments ──
  const fetchPendingPayments = useCallback(async () => {
    if (!db) { setPayLoading(false); return; }
    try {
      const snap = await getDocs(collection(db, "payment_requests"));
      const list = [];
      let revenue = 0;
      snap.forEach(d => {
        const data = { ...d.data(), docId: d.id };
        list.push(data);
        if (data.status === "approved") revenue += data.amount || 0;
      });
      list.sort((a, b) => (b.createdAt?.seconds || 0) - (a.createdAt?.seconds || 0));
      setPendingPayments(list);
      setStats(prev => ({ ...prev, totalRevenue: revenue }));
    } catch (err) { console.error(err); }
    finally { setPayLoading(false); }
  }, []);

  useEffect(() => { fetchAllUsers(); fetchPendingPayments(); }, [fetchAllUsers, fetchPendingPayments]);

  // ── Make Premium ──
  const handleMakePremium = async (u) => {
    if (!db) return;
    try {
      const expiry = new Date(); expiry.setMonth(expiry.getMonth() + 1);
      await setDoc(doc(db, "users", u.uid), {
        isPremium: true, plan: "pro",
        premiumPlan: "monthly", premiumExpiry: expiry,
        premiumGrantedAt: serverTimestamp(), premiumGrantedBy: "admin",
      }, { merge: true });
      await setDoc(doc(db, "spreadsheets", u.uid), { isPremium: true }, { merge: true });
      setActionMsg(`✅ ${u.email} is now Premium!`);
      fetchAllUsers(); setTimeout(() => setActionMsg(""), 3000);
    } catch (err) { setActionMsg("❌ " + err.message); }
  };

  // ── Revoke Premium ──
  const handleRevokePremium = async (u) => {
    if (!db) return;
    try {
      await setDoc(doc(db, "users", u.uid), { isPremium: false, plan: "free" }, { merge: true });
      await setDoc(doc(db, "spreadsheets", u.uid), { isPremium: false }, { merge: true });
      setActionMsg(`⚠️ Premium revoked for ${u.email}`);
      fetchAllUsers(); setTimeout(() => setActionMsg(""), 3000);
    } catch (err) { setActionMsg("❌ " + err.message); }
  };

  // ── Approve Payment ──
  const handleApprovePayment = async (payment) => {
    if (!db) return;
    try {
      const expiry = new Date(); expiry.setMonth(expiry.getMonth() + (payment.plan === "yearly" ? 12 : 1));
      await setDoc(doc(db, "users", payment.uid), {
        isPremium: true, plan: "pro",
        premiumPlan: payment.plan || "monthly", premiumExpiry: expiry,
        premiumGrantedAt: serverTimestamp(), premiumGrantedBy: "admin_bkash",
        paymentStatus: "approved",
      }, { merge: true });
      await setDoc(doc(db, "spreadsheets", payment.uid), { isPremium: true }, { merge: true });
      await setDoc(doc(db, "payment_requests", payment.docId), {
        status: "approved", approvedAt: serverTimestamp(), approvedBy: currentUser?.email,
      }, { merge: true });
      setActionMsg(`✅ ${payment.email} is now Premium!`);
      fetchPendingPayments(); fetchAllUsers(); setTimeout(() => setActionMsg(""), 4000);
    } catch (err) { setActionMsg("❌ " + err.message); }
  };

  // ── Reject Payment ──
  const handleRejectPayment = async (payment) => {
    if (!db) return;
    try {
      await setDoc(doc(db, "payment_requests", payment.docId), {
        status: "rejected", rejectedAt: serverTimestamp(), rejectedBy: currentUser?.email,
      }, { merge: true });
      setActionMsg(`❌ Payment rejected for ${payment.email}`);
      fetchPendingPayments(); setTimeout(() => setActionMsg(""), 3000);
    } catch (err) { setActionMsg("❌ " + err.message); }
  };

  // ── Filtered users ──
  const filteredUsers = allUsers.filter(u => {
    const matchSearch = !searchUser || u.email?.toLowerCase().includes(searchUser.toLowerCase()) || u.name?.toLowerCase().includes(searchUser.toLowerCase());
    const matchPlan = filterPlan === "all" || (filterPlan === "premium" && (u.isPremium || u.email === ADMIN_EMAIL)) || (filterPlan === "free" && !u.isPremium && u.email !== ADMIN_EMAIL);
    return matchSearch && matchPlan;
  });

  const tabs = [
    { id: "overview", label: "📊 Overview" },
    { id: "users",    label: "👥 Users",    badge: stats.totalUsers },
    { id: "payments", label: "💸 Payments", badge: pendingCount > 0 ? pendingCount : null, badgeRed: true },
  ];

  return (
    <div className="fixed inset-0 z-[200] flex items-center justify-center bg-black/80 backdrop-blur-sm p-4">
      <div className={`${bg} ${text} w-full max-w-3xl rounded-[2rem] border ${border} shadow-2xl max-h-[92vh] flex flex-col overflow-hidden`}>

        {/* ── Header ── */}
        <div className={`px-7 pt-7 pb-5 border-b ${border} flex-shrink-0`}>
          <div className="flex justify-between items-start">
            <div>
              <h2 className="text-xl font-black uppercase tracking-tight bg-gradient-to-r from-blue-400 to-yellow-300 bg-clip-text text-transparent flex items-center gap-2">
                 Admin Dashboard
              </h2>
              <p className={`text-[10px] ${sub} mt-1 font-mono`}>{currentUser?.email}</p>
            </div>
            <div className="flex items-center gap-2">
              <button onClick={() => { fetchAllUsers(); fetchPendingPayments(); }}
                className={`px-3 py-1.5 rounded-xl text-[9px] font-black uppercase border ${border} ${sub} hover:text-yellow-400 hover:border-yellow-600/40 transition-all`}>
                Refresh
              </button>
              <button onClick={onClose} className="text-xl hover:opacity-60 transition-opacity">✕</button>
            </div>
          </div>

          {/* Action Message */}
          {actionMsg && (
            <div className={`mt-3 p-3 rounded-xl text-[11px] font-black border ${
              actionMsg.startsWith("✅") ? "bg-emerald-500/10 text-emerald-400 border-emerald-500/20" : "bg-red-500/10 text-red-400 border-red-500/20"
            }`}>{actionMsg}</div>
          )}

          {/* Tabs */}
          <div className="flex gap-2 mt-4">
            {tabs.map(t => (
              <button key={t.id} onClick={() => setActiveTab(t.id)}
                className={`px-4 py-2 rounded-xl font-black text-[10px] uppercase tracking-wider transition-all flex items-center gap-1.5 ${
                  activeTab === t.id
                    ? "bg-gradient-to-r from-yellow-500 to-amber-400 text-slate-900"
                    : `border ${border} ${sub} hover:text-yellow-400`
                }`}>
                {t.label}
                {t.badge != null && (
                  <span className={`text-[8px] font-black px-1.5 py-0.5 rounded-full ${t.badgeRed ? "bg-red-500 text-white" : "bg-blue-500/20 text-blue-400"}`}>
                    {t.badge}
                  </span>
                )}
              </button>
            ))}
          </div>
        </div>

        {/* ── Content ── */}
        <div className="overflow-y-auto flex-1 p-6">

          {/* ══ OVERVIEW TAB ══ */}
          {activeTab === "overview" && (
            <div className="space-y-5">

              {/* Stats Grid */}
              <div className="grid grid-cols-2 md:grid-cols-4 gap-3">
                {[
                  { label: "Total Users",   value: loading ? "…" : stats.totalUsers,   icon: "👥", color: "text-sky-400",     bg: "bg-sky-500/10",     border: "border-sky-500/20" },
                  { label: "Premium",       value: loading ? "…" : stats.premiumUsers, icon: "⭐", color: "text-yellow-400",  bg: "bg-yellow-500/10",  border: "border-yellow-500/20" },
                  { label: "Total Tables",  value: loading ? "…" : stats.totalTables,  icon: "📊", color: "text-emerald-400", bg: "bg-emerald-500/10", border: "border-emerald-500/20" },
                  { label: "Revenue (BDT)", value: loading ? "…" : `${stats.totalRevenue.toLocaleString()}`, icon: "💰", color: "text-amber-400", bg: "bg-amber-500/10", border: "border-amber-500/20" },
                ].map((s, i) => (
                  <div key={i} className={`p-4 rounded-2xl border ${s.border} ${s.bg} text-center`}>
                    <div className="text-2xl mb-1">{s.icon}</div>
                    <div className={`text-2xl font-black ${s.color}`}>{s.value}</div>
                    <div className={`text-[9px] font-black uppercase tracking-widest ${sub} mt-1`}>{s.label}</div>
                  </div>
                ))}
              </div>

              {/* Recent Activity */}
              <div className={`p-5 rounded-2xl border ${border} ${isDark ? "bg-[#08122a]" : "bg-slate-50"}`}>
                <h3 className={`text-[10px] font-black uppercase tracking-widest ${sub} mb-4`}>📋 Recent Users</h3>
                <div className="space-y-2">
                  {loading ? (
                    <div className="flex justify-center py-4"><div className="w-6 h-6 border-2 border-yellow-500/30 border-t-yellow-500 rounded-full animate-spin"></div></div>
                  ) : allUsers.slice(0, 5).map((u, i) => (
                    <div key={i} className="flex items-center gap-3">
                      <img src={u.photoURL || `https://ui-avatars.com/api/?name=${encodeURIComponent(u.name || "U")}&background=1e3a8a&color=d4af37`}
                        alt="" className="w-7 h-7 rounded-full flex-shrink-0"
                        onError={e => { e.target.src = `https://ui-avatars.com/api/?name=U&background=1e3a8a&color=d4af37`; }} />
                      <div className="flex-1 min-w-0">
                        <p className="text-[11px] font-black truncate">{u.name || "Unknown"}</p>
                        <p className={`text-[9px] ${sub} font-mono truncate`}>{u.email}</p>
                      </div>
                      <span className={`text-[8px] font-black px-2 py-0.5 rounded-full uppercase flex-shrink-0 ${
                        u.isPremium || u.email === ADMIN_EMAIL ? "bg-yellow-500/20 text-yellow-400" : "bg-slate-500/20 text-slate-400"
                      }`}>{u.email === ADMIN_EMAIL ? "Admin" : u.isPremium ? "Pro" : "Free"}</span>
                    </div>
                  ))}
                </div>
              </div>

              {/* Pending payments alert */}
              {pendingCount > 0 && (
                <div className={`p-4 rounded-2xl border border-red-500/30 bg-red-500/5 flex items-center justify-between`}>
                  <div className="flex items-center gap-3">
                    <span className="text-2xl">⚠️</span>
                    <div>
                      <p className="text-[11px] font-black text-red-400">{pendingCount} Pending Payment{pendingCount > 1 ? "s" : ""}</p>
                      <p className={`text-[9px] ${sub}`}>Approve to activate Premium for users</p>
                    </div>
                  </div>
                  <button onClick={() => setActiveTab("payments")}
                    className="bg-red-500 hover:bg-red-400 text-white text-[9px] font-black px-3 py-1.5 rounded-xl uppercase transition-all">
                    Review →
                  </button>
                </div>
              )}
            </div>
          )}

          {/* ══ USERS TAB ══ */}
          {activeTab === "users" && (
            <div className="space-y-4">
              {/* Search + Filter */}
              <div className="flex gap-2">
                <input type="text" value={searchUser} onChange={e => setSearchUser(e.target.value)}
                  placeholder="🔍 Search by name or email..."
                  className={`flex-1 px-4 py-2.5 rounded-xl border outline-none text-[11px] transition-all ${input}`} />
                <select value={filterPlan} onChange={e => setFilterPlan(e.target.value)}
                  className={`px-3 py-2.5 rounded-xl border outline-none text-[11px] font-black uppercase ${input}`}>
                  <option value="all">All</option>
                  <option value="premium">Premium</option>
                  <option value="free">Free</option>
                </select>
              </div>

              {loading ? (
                <div className="flex justify-center py-12"><div className="w-8 h-8 border-2 border-yellow-500/30 border-t-yellow-500 rounded-full animate-spin"></div></div>
              ) : filteredUsers.length === 0 ? (
                <p className={`text-center py-8 ${sub} text-sm`}>No users found.</p>
              ) : (
                <div className="space-y-2">
                  {filteredUsers.map((u, i) => {
                    const isAdminUser = u.email === ADMIN_EMAIL;
                    const isPrem = u.isPremium || isAdminUser;
                    const joinDate = u.lastLogin ? new Date(u.lastLogin.seconds * 1000).toLocaleDateString() : "—";
                    const expiry = u.premiumExpiry ? new Date(u.premiumExpiry.seconds * 1000).toLocaleDateString() : null;
                    return (
                      <div key={i} className={`flex items-center gap-3 p-4 rounded-2xl border ${card}`}>
                        <img src={u.photoURL || `https://ui-avatars.com/api/?name=${encodeURIComponent(u.name || "U")}&background=1e3a8a&color=d4af37`}
                          alt="" className="w-9 h-9 rounded-full flex-shrink-0"
                          onError={e => { e.target.src = `https://ui-avatars.com/api/?name=U&background=1e3a8a&color=d4af37`; }} />
                        <div className="flex-1 min-w-0">
                          <p className="font-black text-sm truncate">{u.name || "Unknown"}</p>
                          <p className={`text-[9px] ${sub} font-mono truncate`}>{u.email}</p>
                          <div className="flex items-center gap-2 mt-1 flex-wrap">
                            <span className={`text-[8px] font-black px-2 py-0.5 rounded-full uppercase ${isPrem ? "bg-yellow-500/20 text-yellow-400" : "bg-slate-500/20 text-slate-400"}`}>
                              {isAdminUser ? "Admin" : isPrem ? "Pro" : "Free"}
                            </span>
                            <span className={`text-[8px] ${sub}`}>Login: {joinDate}</span>
                            {expiry && <span className={`text-[8px] ${sub}`}>Expires: {expiry}</span>}
                            <span className="bg-blue-500/20 text-blue-400 text-[8px] font-black px-2 py-0.5 rounded-full uppercase">{u.tableCount || 0} tables</span>
                          </div>
                        </div>
                        {!isAdminUser && (
                          isPrem
                            ? <button onClick={() => handleRevokePremium(u)} className="bg-red-500/15 hover:bg-red-500 text-red-400 hover:text-white text-[9px] font-black px-3 py-1.5 rounded-xl uppercase transition-all flex-shrink-0">Revoke</button>
                            : <button onClick={() => handleMakePremium(u)} className="bg-gradient-to-r from-yellow-500 to-amber-400 hover:from-yellow-400 hover:to-amber-300 text-slate-900 text-[9px] font-black px-3 py-1.5 rounded-xl uppercase transition-all flex-shrink-0">⭐ Make Pro</button>
                        )}
                      </div>
                    );
                  })}
                </div>
              )}
            </div>
          )}

          {/* ══ PAYMENTS TAB ══ */}
          {activeTab === "payments" && (
            <div className="space-y-3">
              {payLoading ? (
                <div className="flex justify-center py-12"><div className="w-8 h-8 border-2 border-yellow-500/30 border-t-yellow-500 rounded-full animate-spin"></div></div>
              ) : pendingPayments.length === 0 ? (
                <div className="text-center py-12">
                  <div className="text-4xl mb-3">💸</div>
                  <p className={`${sub} text-sm font-black uppercase`}>No payment requests yet</p>
                </div>
              ) : (
                pendingPayments.map((p, i) => (
                  <div key={i} className={`p-4 rounded-2xl border ${card}`}>
                    <div className="flex justify-between items-start gap-4">
                      <div className="flex-1 min-w-0">
                        <div className="flex items-center gap-2 mb-1">
                          <p className="font-black text-sm truncate">{p.email || "Unknown"}</p>
                          <span className={`text-[8px] font-black px-2 py-0.5 rounded-full uppercase flex-shrink-0 ${
                            p.status === "approved" ? "bg-emerald-500/20 text-emerald-400" :
                            p.status === "rejected" ? "bg-red-500/20 text-red-400" :
                            "bg-amber-500/20 text-amber-400"
                          }`}>{p.status}</span>
                        </div>
                        <p className={`text-[10px] font-mono ${sub}`}>
                          TxID: <span className="text-yellow-400 font-black">{p.transactionId}</span>
                        </p>
                        <div className="flex items-center gap-3 mt-1">
                          <span className={`text-[9px] ${sub}`}>
                            Amount: <span className="font-black text-emerald-400">{p.amount} BDT</span>
                          </span>
                          <span className={`text-[9px] ${sub}`}>
                            Plan: <span className="font-black uppercase text-blue-400">{p.plan}</span>
                          </span>
                          {p.createdAt && (
                            <span className={`text-[9px] ${sub}`}>
                              {new Date(p.createdAt.seconds * 1000).toLocaleDateString()}
                            </span>
                          )}
                        </div>
                      </div>
                      {p.status === "pending" && (
                        <div className="flex flex-col gap-1.5 flex-shrink-0">
                          <button onClick={() => handleApprovePayment(p)}
                            className="bg-gradient-to-r from-yellow-500 to-amber-400 hover:from-yellow-400 hover:to-amber-300 text-slate-900 text-[9px] font-black px-3 py-1.5 rounded-xl uppercase transition-all">
                            ✓ Approve
                          </button>
                          <button onClick={() => handleRejectPayment(p)}
                            className="bg-red-500/15 hover:bg-red-500 text-red-400 hover:text-white text-[9px] font-black px-3 py-1.5 rounded-xl uppercase transition-all">
                            ✗ Reject
                          </button>
                        </div>
                      )}
                    </div>
                  </div>
                ))
              )}
            </div>
          )}
        </div>
      </div>
    </div>
  );
}

function WelcomeModal({ onClose, isDark, userName }) {
  const [step, setStep] = useState(0);
  const bg = isDark
  ? "bg-[#0f2038]"
  : "bg-[#ffffff]";

const text = isDark
  ? "text-[#f1f6ff]"
  : "text-[#172033]";

const sub = isDark
  ? "text-[#a9bbd6]"
  : "text-[#5f6f89]";
 
  const steps = [
    {
      icon: "👋",
      title: `Welcome, ${userName || "Friend"}!`,
      desc: "NeoSheet Pro is your AI-powered smart SHEETMIND. Let's take a quick tour!",
      color: "#d4af37",
    },
    {
      icon: "⚡",
      title: "AI Generator",
      desc: "Type any command like 'Salary sheet 20 rows' and AI will build the table structure instantly for you.",
      color: "#6366f1",
    },
    {
      icon: "📋",
      title: "Templates",
      desc: "Choose from Student Sheet, Salary, Inventory, CRM and more — all pre-configured and ready to use.",
      color: "#10b981",
    },
    {
      icon: "💾",
      title: "Auto Save",
      desc: "Your data saves automatically to the cloud. Access your sheets from any device, anytime.",
      color: "#f59e0b",
    },
    {
      icon: "📤",
      title: "Export Anywhere",
      desc: "Download your data as PDF or Excel with one click. Share with anyone, anywhere.",
      color: "#22d3ee",
    },
    {
      icon: "🚀",
      title: "You're Ready!",
      desc: "Free plan includes 4 tables. Upgrade to Pro for unlimited tables, AI features and more.",
      color: "#d4af37",
    },
  ];
 
  const current = steps[step];
  const isLast  = step === steps.length - 1;
 
  return (
    <div className="fixed inset-0 z-[400] flex items-center justify-center bg-black/85 backdrop-blur-sm p-4">
      <div className={`${bg} ${text} w-full max-w-sm rounded-[2rem] border overflow-hidden`}
  style={{
    borderColor: isDark ? "#3f5f89" : "#c8d7ec",
    boxShadow: isDark
      ? "0 28px 80px rgba(0,0,0,0.45), inset 0 1px 0 rgba(255,255,255,0.06)"
      : "0 28px 80px rgba(43,76,126,0.22), inset 0 1px 0 rgba(255,255,255,0.9)"
  }}>
 
        {/* Top color bar */}
        <div className="h-1.5 w-full transition-all duration-500"
          style={{ background: `linear-gradient(to right, ${current.color}, ${current.color}88)` }} />
 
        {/* Content */}
        <div className="p-8 text-center">
          <div className="w-20 h-20 rounded-full flex items-center justify-center mx-auto mb-5 text-4xl transition-all duration-300"
            style={{ backgroundColor: current.color + "18", border: `2px solid ${current.color}40` }}>
            {current.icon}
          </div>
 
          <h2 className="text-xl font-black uppercase tracking-tight mb-3"
            style={{ color: current.color }}>
            {current.title}
          </h2>
 
          <p className={`text-[12px] ${sub} leading-relaxed mb-6`}>
            {current.desc}
          </p>
 
          {/* Step dots */}
          <div className="flex justify-center gap-2 mb-6">
            {steps.map((_, i) => (
              <button key={i} onClick={() => setStep(i)}
                className="rounded-full transition-all duration-300"
                style={{
                  width:  i === step ? "20px" : "8px",
                  height: "8px",
                  backgroundColor: i === step ? current.color : (isDark ? "#1e3a8a" : "#e2e8f0"),
                }} />
            ))}
          </div>
 
          {/* Buttons */}
          <div className="flex gap-3">
            {step > 0 && (
              <button onClick={() => setStep(s => s - 1)}
                className={`flex-1 py-3 rounded-2xl font-black text-xs border transition-all ${
                  isDark ? "border-yellow-900/40 text-blue-400 hover:bg-yellow-500/10" : "border-blue-200 text-slate-500 hover:bg-blue-50"
                }`}>
                ← Back
              </button>
            )}
            <button
              onClick={isLast ? onClose : () => setStep(s => s + 1)}
              className="flex-1 py-3 rounded-2xl font-black text-sm text-slate-900 transition-all shadow-lg"
              style={{ background: `linear-gradient(135deg, ${current.color}, ${current.color}cc)` }}>
              {isLast ? "Get Started!" : "Next →"}
            </button>
          </div>
 
          {/* Skip */}
          {!isLast && (
            <button onClick={onClose}
              className={`mt-3 text-[10px] font-black uppercase ${sub} hover:opacity-60 transition-opacity`}>
              Skip Tour
            </button>
          )}
        </div>
      </div>
    </div>
  );
}
// ============================================================
// PRO GATE MODAL
// ============================================================
function ProGateModal({ message, onUpgrade, onClose, isDark, tabCount }) {
  const bg   = isDark ? "bg-[#060e1c]" : "bg-white";
  const text = isDark ? "text-slate-100" : "text-slate-900";
  const sub  = isDark ? "text-blue-400"  : "text-slate-500";
  const card = isDark ? "bg-[#08122a] border-yellow-900/20" : "bg-blue-50 border-blue-200";
  const isCooldown = message.includes("hours") || message.includes("min");

  return (
    <div className="fixed inset-0 z-[300] flex items-center justify-center bg-black/80 backdrop-blur-sm p-4">
      <div className={`${bg} ${text} w-full max-w-sm rounded-[2rem] border border-yellow-500/30 shadow-2xl overflow-hidden`}>

        <div className="bg-gradient-to-br from-blue-900/40 via-yellow-900/10 to-transparent p-7 text-center border-b border-yellow-500/10">
          <div className="w-16 h-16 rounded-full bg-yellow-500/10 border-2 border-yellow-500/30 flex items-center justify-center mx-auto mb-4 text-3xl">
            {isCooldown ? "⏳" : "🔒"}
          </div>
          <h2 className="text-xl font-black uppercase tracking-tight bg-gradient-to-r from-blue-400 to-yellow-300 bg-clip-text text-transparent">
            {isCooldown ? "Cooldown Active" : "Free Plan Limit"}
          </h2>
          <p className={`text-[11px] ${sub} mt-2 leading-relaxed`}>{message}</p>
        </div>

        <div className="p-6">
          <div className={`p-4 rounded-2xl border mb-5 ${card}`}>
            <div className="flex justify-between items-center mb-3">
              <span className={`text-[10px] font-black uppercase tracking-widest ${sub}`}>তোমার Plan</span>
              <span className="text-[10px] font-black px-3 py-1 rounded-full bg-blue-500/20 text-blue-400 uppercase">Free</span>
            </div>
            <div className="flex justify-between items-center mb-2">
              <span className={`text-[11px] ${sub}`}>Tables used</span>
              <span className="text-[11px] font-black text-yellow-400">{tabCount} / {FREE_TABLE_LIMIT}</span>
            </div>
            <div className={`w-full h-2 rounded-full ${isDark ? "bg-slate-800" : "bg-slate-200"} overflow-hidden`}>
              <div
                className="h-full rounded-full bg-gradient-to-r from-yellow-500 to-amber-400 transition-all"
                style={{ width: `${Math.min((tabCount / FREE_TABLE_LIMIT) * 100, 100)}%` }}
              />
            </div>
          </div>

          <div className={`p-4 rounded-2xl border mb-5 ${isDark ? "border-yellow-500/20 bg-yellow-500/5" : "border-yellow-200 bg-yellow-50"}`}>
            <p className="text-[10px] font-black text-yellow-400 uppercase tracking-widest mb-3">⭐ Pro তে যা পাবে</p>
            {[
              "Unlimited tables — কোনো limit নেই",
              "24h cooldown নেই",
              "সব AI features unlimited",
              "Ad-free experience",
            ].map((b, i) => (
              <p key={i} className={`text-[10px] ${sub} flex items-center gap-2 mt-1.5`}>
                <span className="text-yellow-400 font-black">✓</span> {b}
              </p>
            ))}
          </div>

          <button onClick={onUpgrade}
            className="w-full py-3.5 rounded-2xl font-black text-sm bg-gradient-to-r from-yellow-500 to-amber-400 hover:from-yellow-400 hover:to-amber-300 text-slate-900 transition-all shadow-lg shadow-yellow-900/20 mb-3">
            ⭐ Upgrade to Pro — 300 BDT/month
          </button>
          <button onClick={onClose}
            className={`w-full py-3 rounded-2xl font-black text-xs border transition-all ${isDark ? "border-yellow-900/40 hover:bg-yellow-500/10 text-blue-400" : "border-blue-200 hover:bg-blue-50 text-slate-500"}`}>
            পরে করবো
          </button>
        </div>
      </div>
    </div>
  );
}

// ============================================================
// UPGRADE MODAL — Deep Blue + Gold Theme
// ============================================================
function UpgradeModal({ onClose, isDark, currentUser }) {
  const [plan, setPlan] = useState("monthly");
  const [step, setStep] = useState("plans");
  const [txId, setTxId] = useState("");
  const [amount, setAmount] = useState("");
  const [errors, setErrors] = useState({});
  const [submitting, setSubmitting] = useState(false);
  const [copied, setCopied] = useState(false);

  const PLANS = {
    monthly: { label: "Monthly", bdtPrice: 300, usdPrice: 3, period: "month", months: 1 },
    yearly:  { label: "Yearly",  bdtPrice: 2500, usdPrice: 25, period: "year", months: 12, save: "Save 30%" },
  };
  const selectedPlan = PLANS[plan];

  // Deep Blue + Gold colors
  const bg      = isDark ? "bg-[#060e1c]" : "bg-white";
  const cardBg  = isDark ? "bg-[#08122a]/80 border-yellow-900/20" : "bg-blue-50 border-blue-200";
  const text    = isDark ? "text-slate-100" : "text-slate-900";
  const sub     = isDark ? "text-blue-400" : "text-slate-500";
  const inputCls = isDark
    ? "bg-[#050d1f] border-[#d4af3730] text-slate-200 placeholder:text-blue-900 focus:border-yellow-500"
    : "bg-blue-50 border-blue-300 text-slate-800 placeholder:text-slate-400 focus:border-yellow-400";

  const copyNumber = () => {
    navigator.clipboard.writeText(BKASH_NUMBER);
    setCopied(true);
    setTimeout(() => setCopied(false), 2000);
  };

  const validateBkash = () => {
    const newErrors = {};
    if (!txId.trim() || txId.trim().length < 5)
      newErrors.txId = "Valid Transaction ID দাও (কমপক্ষে ৫ অক্ষর)";
    const parsedAmount = parseFloat(amount);
    if (isNaN(parsedAmount) || parsedAmount < selectedPlan.bdtPrice)
      newErrors.amount = `কমপক্ষে ${selectedPlan.bdtPrice} BDT পাঠাতে হবে`;
    setErrors(newErrors);
    return Object.keys(newErrors).length === 0;
  };

  const handleSubmitPayment = async () => {
    if (!validateBkash()) return;
    setSubmitting(true);
    try {
      if (db && currentUser) {
        const requestId = `bkash_${currentUser.uid}_${Date.now()}`;

        // Save payment request
        await setDoc(doc(db, "payment_requests", requestId), {
          uid:           currentUser.uid,
          email:         currentUser.email,
          name:          currentUser.displayName || "",
          transactionId: txId.trim().toUpperCase(),
          amount:        parseFloat(amount),
          plan,
          bdtPrice:      selectedPlan.bdtPrice,
          status:        "pending",
          method:        "bkash",
          createdAt:     serverTimestamp(),
        });

        // Also update user profile payment status
        await setDoc(doc(db, "users", currentUser.uid), {
          paymentStatus:       "pending",
          paymentRef:          txId.trim().toUpperCase(),
          paymentAmount:       parseFloat(amount),
          paymentPlan:         plan,
          paymentSubmittedAt:  serverTimestamp(),
        }, { merge: true });

        // Save admin notification
        await setDoc(doc(db, "admin_notifications", requestId), {
          type:          "payment_request",
          uid:           currentUser.uid,
          email:         currentUser.email,
          transactionId: txId.trim().toUpperCase(),
          amount:        parseFloat(amount),
          plan,
          read:          false,
          createdAt:     serverTimestamp(),
        });
      }
      setStep("submitted");
    } catch (err) {
      setErrors({ general: "Submit করতে সমস্যা হয়েছে। আবার চেষ্টা করো।" });
    } finally {
      setSubmitting(false);
    }
  };

  // ── Success Screen ──
  if (step === "submitted") {
    return (
      <div className="fixed inset-0 z-[300] flex items-center justify-center bg-black/90 backdrop-blur-sm p-4">
        <div className={`${bg} ${text} w-full max-w-sm rounded-[2rem] border border-yellow-500/40 shadow-2xl p-10 text-center`}>
          <div className="w-20 h-20 rounded-full bg-yellow-500/10 border-2 border-yellow-500/40 flex items-center justify-center mx-auto mb-6 text-4xl">✅</div>
          <h2 className="text-2xl font-black uppercase tracking-tight mb-2 bg-gradient-to-r from-yellow-400 to-amber-300 bg-clip-text text-transparent">
            Request Submitted!
          </h2>
          <p className={`text-sm ${sub} mb-3`}>
            তোমার payment request admin এর কাছে পাঠানো হয়েছে। Approve হলে Premium activate হবে।
          </p>
          <div className={`p-3 rounded-xl border mb-2 ${isDark ? "border-yellow-500/20 bg-yellow-500/5" : "border-yellow-200 bg-yellow-50"}`}>
            <p className={`text-[10px] font-mono ${sub}`}>TXN ID: <span className="font-black text-yellow-400">{txId.toUpperCase()}</span></p>
            <p className={`text-[10px] ${sub} mt-1`}>Plan: <span className="font-black uppercase text-yellow-400">{plan}</span> · Amount: <span className="font-black text-yellow-400">{amount} BDT</span></p>
          </div>
          <div className={`p-3 rounded-xl border ${isDark ? "border-blue-500/20 bg-blue-500/5" : "border-blue-200 bg-blue-50"} mb-5`}>
            <p className="text-[10px] font-black text-blue-400 uppercase tracking-wider">⏳ Admin Approval Pending</p>
            <p className={`text-[9px] ${sub} mt-1`}>সাধারণত কয়েক ঘণ্টার মধ্যে approve হয়।</p>
          </div>
          <button onClick={onClose}
            className="w-full py-3 rounded-2xl font-black text-sm bg-gradient-to-r from-yellow-500 to-amber-400 hover:from-yellow-400 hover:to-amber-300 text-slate-900 transition-all">
            বন্ধ করো
          </button>
        </div>
      </div>
    );
  }

  return (
    <div className="fixed inset-0 z-[300] flex items-center justify-center bg-black/85 backdrop-blur-sm p-4">
      <div className={`${bg} ${text} w-full max-w-md rounded-[2rem] border border-yellow-500/30 shadow-2xl overflow-hidden`}
        style={{ maxHeight: "90vh", overflowY: "auto" }}>

        {/* Header */}
        <div className="bg-gradient-to-br from-blue-900/40 via-yellow-900/10 to-transparent p-8 border-b border-yellow-500/10">
          <div className="flex justify-between items-start">
            <div>
              <p className="text-[10px] font-black text-yellow-400 uppercase tracking-[0.3em] mb-2">⭐ Upgrade to Premium</p>
              <h2 className="text-2xl font-black tracking-tight bg-gradient-to-r from-blue-400 to-yellow-300 bg-clip-text text-transparent">
                NeoSheet Pro
              </h2>
              <p className={`text-[10px] ${sub} mt-1`}>bKash দিয়ে payment করো, admin activate করবে</p>
            </div>
            <button onClick={onClose} className="text-xl hover:opacity-60 transition-opacity mt-1">✕</button>
          </div>
        </div>

        <div className="p-6">

          {/* STEP 1: Plan Selection */}
          {step === "plans" && (
            <>
              <div className="grid grid-cols-2 gap-3 mb-5">
                {Object.entries(PLANS).map(([key, p]) => (
                  <button key={key} onClick={() => setPlan(key)}
                    className={`relative p-4 rounded-2xl border text-left transition-all ${
                      plan === key
                        ? "border-yellow-500 bg-yellow-500/10"
                        : cardBg
                    }`}>
                    {p.save && (
                      <span className="absolute -top-2 -right-2 bg-yellow-500 text-slate-900 text-[8px] font-black px-2 py-0.5 rounded-full uppercase">{p.save}</span>
                    )}
                    <p className={`text-[10px] font-black uppercase tracking-widest mb-2 ${plan === key ? "text-yellow-400" : sub}`}>{p.label}</p>
                    <p className="text-xl font-black">{p.bdtPrice} <span className={`text-sm font-normal ${sub}`}>BDT</span></p>
                    <p className={`text-[10px] ${sub}`}>≈ ${p.usdPrice} / {p.period}</p>
                  </button>
                ))}
              </div>

              {/* Benefits */}
              <div className={`p-4 rounded-2xl border mb-5 ${isDark ? "border-blue-500/20 bg-blue-500/5" : "border-blue-200 bg-blue-50"}`}>
                <p className="text-[10px] font-black text-blue-400 uppercase tracking-widest mb-2">⭐ Premium সুবিধা</p>
                {[
                  "Unlimited tables (free তে মাত্র ৪টা)",
                  "24 ঘণ্টা cooldown নেই",
                  "সব AI features unlimited",
                  "Priority support",
                  "Ad-free experience"
                ].map((b, i) => (
                  <p key={i} className={`text-[10px] ${sub} flex items-center gap-2 mt-1.5`}>
                    <span className="text-yellow-400">✓</span> {b}
                  </p>
                ))}
              </div>

              <button onClick={() => setStep("bkash")}
                className="w-full bg-gradient-to-r from-yellow-500 to-amber-400 hover:from-yellow-400 hover:to-amber-300 text-slate-900 py-4 rounded-2xl font-black text-sm uppercase tracking-widest transition-all flex items-center justify-center gap-2 shadow-lg shadow-yellow-900/20">
                <span>💸</span> bKash দিয়ে Pay করো →
              </button>
              {/* OR divider */}
              <div className="flex items-center gap-3 my-3">
                <div className={`flex-1 h-px ${isDark ? "bg-yellow-900/30" : "bg-blue-100"}`} />
                <span className={`text-[9px] font-black uppercase tracking-widest ${sub}`}>or</span>
                <div className={`flex-1 h-px ${isDark ? "bg-yellow-900/30" : "bg-blue-100"}`} />
              </div>
 
              {/* Card Payment Button */}
              <a
                href={LEMON_WEB_LINK}
                target="_blank"
                rel="noopener noreferrer"
                className="w-full py-4 rounded-2xl font-black text-sm uppercase tracking-widest transition-all flex items-center justify-center gap-2 shadow-lg border border-blue-500/30 hover:border-blue-400/60"
                style={{ background: "linear-gradient(135deg, #1e3a8a, #1d4ed8)", color: "white" }}>
                💳 Pay with Card (International)
              </a>
              <p className={`text-[9px] ${sub} text-center mt-1`}>
                Visa, Mastercard, PayPal — Instant activation
              </p>
            </>
          )}

          {/* STEP 2: bKash Payment Form */}
          {step === "bkash" && (
            <>
              {/* Instructions */}
              <div className={`p-4 rounded-2xl border mb-5 ${isDark ? "border-yellow-500/20 bg-yellow-500/5" : "border-yellow-200 bg-yellow-50"}`}>
                <p className="text-[10px] font-black text-yellow-400 uppercase tracking-widest mb-3">📱 bKash দিয়ে কীভাবে Pay করবে</p>
                <ol className={`text-[10px] ${sub} space-y-2`}>
                  <li className="flex items-start gap-2">
                    <span className="text-yellow-400 font-black flex-shrink-0">১.</span>
                    bKash app খোলো → <span className="text-yellow-300 font-black">Send Money</span> চাপো
                  </li>
                  <li className="flex items-start gap-2">
                    <span className="text-yellow-400 font-black flex-shrink-0">২.</span>
                    নিচের নম্বরে <span className="text-yellow-300 font-black">{selectedPlan.bdtPrice} BDT</span> পাঠাও
                  </li>
                  <li className="flex items-start gap-2">
                    <span className="text-yellow-400 font-black flex-shrink-0">৩.</span>
                    Confirmation SMS থেকে <span className="font-black">Transaction ID</span> copy করো
                  </li>
                  <li className="flex items-start gap-2">
                    <span className="text-yellow-400 font-black flex-shrink-0">৪.</span>
                    নিচের form এ TXN ID ও amount দিয়ে submit করো
                  </li>
                </ol>

                {/* bKash number box */}
                <div className={`mt-4 p-3 rounded-xl ${isDark ? "bg-[#050d1f]" : "bg-white"} border ${isDark ? "border-yellow-500/20" : "border-yellow-300"}`}>
                  <div className="flex items-center justify-between mb-2">
                    <span className={`text-[9px] font-black uppercase tracking-widest ${sub}`}>bKash Number</span>
                    <button onClick={copyNumber}
                      className={`text-[9px] font-black px-2 py-1 rounded-lg transition-all ${copied ? "bg-yellow-500 text-slate-900" : "bg-yellow-500/20 text-yellow-400 hover:bg-yellow-500/30"}`}>
                      {copied ? "✓ Copied!" : "Copy"}
                    </button>
                  </div>
                  <p className="text-xl font-black text-yellow-400 tracking-widest text-center">{BKASH_NUMBER}</p>
                  <div className="flex items-center justify-between mt-2">
                    <span className={`text-[9px] ${sub}`}>পাঠানোর পরিমাণ</span>
                    <span className="text-base font-black text-yellow-400">{selectedPlan.bdtPrice} BDT</span>
                  </div>
                </div>
              </div>

              {/* Form */}
              <div className="space-y-4 mb-5">
                <div>
                  <label className={`text-[10px] font-black uppercase tracking-widest ${sub} mb-1.5 block`}>
                    Transaction ID (TrxID) *
                  </label>
                  <input
                    type="text"
                    value={txId}
                    onChange={(e) => setTxId(e.target.value.toUpperCase())}
                    placeholder="যেমন: AB12345678"
                    className={`w-full px-4 py-3 rounded-2xl border outline-none text-sm font-mono tracking-widest transition-all ${inputCls} ${errors.txId ? "border-red-500" : ""}`}
                  />
                  {errors.txId && <p className="text-red-400 text-[10px] mt-1">{errors.txId}</p>}
                </div>
                <div>
                  <label className={`text-[10px] font-black uppercase tracking-widest ${sub} mb-1.5 block`}>
                    Amount Sent (BDT) *
                  </label>
                  <input
                    type="number"
                    value={amount}
                    onChange={(e) => setAmount(e.target.value)}
                    placeholder={`${selectedPlan.bdtPrice}`}
                    min={selectedPlan.bdtPrice}
                    className={`w-full px-4 py-3 rounded-2xl border outline-none text-sm transition-all ${inputCls} ${errors.amount ? "border-red-500" : ""}`}
                  />
                  {errors.amount && <p className="text-red-400 text-[10px] mt-1">{errors.amount}</p>}
                </div>
                {errors.general && (
                  <div className="p-3 rounded-xl bg-red-500/10 border border-red-500/20">
                    <p className="text-red-400 text-[10px]">{errors.general}</p>
                  </div>
                )}
              </div>

              {/* Warning */}
              <div className={`p-3 rounded-xl border mb-5 ${isDark ? "border-blue-500/20 bg-blue-500/5" : "border-blue-200 bg-blue-50"}`}>
                <p className="text-[9px] font-black text-blue-400 uppercase tracking-wider">⚠️ Manual Approval Required</p>
                <p className={`text-[9px] ${sub} mt-0.5`}>Submit করার পর admin verify করবে। Instant activate হবে না।</p>
              </div>

              <div className="flex gap-3">
                <button onClick={() => setStep("plans")}
                  className={`px-5 py-3 rounded-2xl font-black text-xs border transition-all ${isDark ? "border-yellow-900/40 hover:bg-yellow-500/10 text-blue-400" : "border-blue-200 hover:bg-blue-50"}`}>
                  Back
                </button>
                <button
                  onClick={handleSubmitPayment}
                  disabled={submitting}
                  className="flex-1 bg-gradient-to-r from-yellow-500 to-amber-400 hover:from-yellow-400 hover:to-amber-300 disabled:opacity-50 text-slate-900 py-3 rounded-2xl font-black text-xs uppercase tracking-widest transition-all flex items-center justify-center gap-2 shadow-md shadow-yellow-900/20"
                >
                  {submitting
                    ? <><div className="w-3 h-3 border-2 border-slate-900/30 border-t-slate-900 rounded-full animate-spin"></div>Submitting...</>
                    : "✓ Submit for Approval"
                  }
                </button>
              </div>
            </>
          )}
        </div>
      </div>
    </div>
  );
}
//==========================================================
  //AITableModal
//==========================================================
function AITableModal({ onClose, onGenerate, isDark, isPremium, aiTableUsage }) {
  const [prompt, setPrompt] = useState("");

  const remaining = Math.max(AI_FREE_TABLE_LIMIT - aiTableUsage, 0);

  const bg = isDark ? "bg-[#0f2038]" : "bg-white";
  const text = isDark ? "text-[#f1f6ff]" : "text-[#172033]";
  const sub = isDark ? "text-[#9fb3d1]" : "text-[#5f6f89]";
  const input = isDark
    ? "bg-[#0a1729] border-[#315174] text-[#f1f6ff] placeholder:text-[#7186a6]"
    : "bg-white border-[#c8d7ec] text-[#172033] placeholder:text-[#7b8aa3]";

  return (
    <div className="fixed inset-0 z-[320] flex items-center justify-center bg-black/70 backdrop-blur-sm p-4">
      <div
        className={`${bg} ${text} w-full max-w-md rounded-[2rem] border overflow-hidden`}
        style={{
          borderColor: isDark ? "#3f5f89" : "#c8d7ec",
          boxShadow: isDark
            ? "0 28px 80px rgba(0,0,0,0.45)"
            : "0 28px 80px rgba(43,76,126,0.22)",
        }}
      >
        <div className="p-7 border-b" style={{ borderColor: isDark ? "#29456b" : "#d9e4f3" }}>
          <div className="flex justify-between items-start gap-4">
            <div>
              <p className="text-[10px] font-black uppercase tracking-[0.25em] text-[#3978d8] mb-2">
                AI Table Builder
              </p>
              <h2 className="text-xl font-black uppercase tracking-tight">
                Build a table from a command
              </h2>
              <p className={`text-[11px] ${sub} mt-2 leading-relaxed`}>
                Tell AI what table you need. It will create columns and rows automatically.
              </p>
            </div>
            <button onClick={onClose} className={`text-xl ${sub} hover:opacity-60`}>
              ×
            </button>
          </div>
        </div>

        <div className="p-6">
          {!isPremium && (
            <div className="mb-4 rounded-2xl border px-4 py-3"
              style={{
                borderColor: remaining > 0 ? "#3978d833" : "#ef444433",
                background: remaining > 0 ? "#3978d80d" : "#ef44440d",
              }}
            >
              <p className={`text-[10px] font-black uppercase ${remaining > 0 ? "text-[#3978d8]" : "text-red-400"}`}>
                Free AI tables: {aiTableUsage}/{AI_FREE_TABLE_LIMIT}
              </p>
              <p className={`text-[9px] ${sub} mt-1`}>
                {remaining > 0 ? `${remaining} AI table remaining.` : "Upgrade for unlimited AI tables."}
              </p>
            </div>
          )}

          <textarea
            value={prompt}
            onChange={e => setPrompt(e.target.value)}
            placeholder="Example: Create a product inventory table with 15 rows, SKU, stock, supplier, reorder level and price."
            className={`w-full h-36 resize-none rounded-2xl border outline-none px-4 py-3 text-sm leading-relaxed ${input} focus:border-[#3978d8] focus:ring-2 focus:ring-[#3978d8]/15`}
          />

          <div className="flex gap-3 mt-5">
            <button
              onClick={onClose}
              className={`flex-1 py-3 rounded-xl font-black text-[11px] uppercase border ${
                isDark
                  ? "border-[#315174] text-[#c6d7ef] hover:bg-[#173251]"
                  : "border-[#c8d7ec] text-[#36516f] hover:bg-[#eaf2ff]"
              }`}
            >
              Cancel
            </button>

            <button
              onClick={() => onGenerate(prompt)}
              disabled={!prompt.trim()}
              className="flex-1 py-3 rounded-xl font-black text-[11px] uppercase bg-gradient-to-r from-[#2563eb] to-[#0ea5e9] hover:from-[#1d4ed8] hover:to-[#0284c7] disabled:opacity-50 text-white shadow-[0_12px_28px_rgba(37,99,235,0.28)]"
            >
              Generate
            </button>
          </div>
        </div>
      </div>
    </div>
  );
}
// ============================================================
// AI ANALYSIS MODAL
// ============================================================
function AIAnalysisModal({ columns, rows, isDark, onClose }) {
  const bg   = isDark ? "bg-[#0f172a]" : "bg-white";
  const text = isDark ? "text-slate-100" : "text-slate-900";
  const sub  = isDark ? "text-slate-400" : "text-slate-500";
  const card = isDark ? "bg-slate-800/60 border-white/10" : "bg-slate-50 border-slate-200";

  const numericCols = columns.filter(col =>
    col !== "ID" && rows.some(r => !isNaN(parseFloat(cellVal(r[col]))) && cellVal(r[col]) !== "")
  );

  const getStats = (col) => {
    const vals = rows.map(r => parseFloat(cellVal(r[col]))).filter(v => !isNaN(v));
    if (!vals.length) return null;
    const sum = vals.reduce((a, b) => a + b, 0);
    const avg = sum / vals.length;
    const sorted = [...vals].sort((a, b) => a - b);
    const median = sorted.length % 2 === 0
      ? (sorted[sorted.length / 2 - 1] + sorted[sorted.length / 2]) / 2
      : sorted[Math.floor(sorted.length / 2)];
    return { sum, avg, min: Math.min(...vals), max: Math.max(...vals), median, count: vals.length };
  };

  const textStats = columns.filter(c => c !== "ID" && !numericCols.includes(c)).map(col => {
    const vals = rows.map(r => cellVal(r[col])).filter(v => v !== "");
    const freq = {};
    vals.forEach(v => { freq[v] = (freq[v] || 0) + 1; });
    return { col, count: vals.length, unique: Object.keys(freq).length, top: Object.entries(freq).sort((a, b) => b[1] - a[1]).slice(0, 3) };
  });

  return (
    <div className="fixed inset-0 z-[200] flex items-center justify-center bg-black/70 backdrop-blur-sm p-4">
      <div className={`${bg} ${text} w-full max-w-2xl rounded-[2rem] border ${isDark ? "border-white/10" : "border-slate-200"} shadow-2xl p-8 max-h-[90vh] overflow-y-auto`}>
        <div className="flex justify-between items-center mb-6">
          <div>
            <h2 className="text-xl font-black uppercase tracking-tight">🤖 AI Analysis</h2>
            <p className={`text-[10px] ${sub} uppercase tracking-widest mt-1`}>{rows.length} rows · {columns.length} columns</p>
          </div>
          <button onClick={onClose} className="text-2xl hover:opacity-60 transition-opacity">✕</button>
        </div>
        {numericCols.length > 0 && (
          <div className="mb-6">
            <h3 className={`text-[10px] font-black uppercase tracking-widest ${sub} mb-3`}>📊 Numeric Columns</h3>
            <div className="space-y-3">
              {numericCols.map(col => {
                const s = getStats(col);
                if (!s) return null;
                return (
                  <div key={col} className={`p-4 rounded-2xl border ${card}`}>
                    <p className="font-black text-sm text-sky-400 mb-3 uppercase">{col}</p>
                    <div className="grid grid-cols-5 gap-2">
                      {[
                        { label: "Sum",    value: s.sum.toLocaleString(undefined, { maximumFractionDigits: 2 }) },
                        { label: "Avg",    value: s.avg.toLocaleString(undefined, { maximumFractionDigits: 2 }) },
                        { label: "Min",    value: s.min },
                        { label: "Max",    value: s.max },
                        { label: "Median", value: s.median.toLocaleString(undefined, { maximumFractionDigits: 2 }) },
                      ].map(stat => (
                        <div key={stat.label} className={`p-2 rounded-xl text-center ${isDark ? "bg-slate-900/60" : "bg-white"}`}>
                          <p className="text-xs font-black text-emerald-400">{stat.value}</p>
                          <p className={`text-[8px] font-black uppercase ${sub} mt-0.5`}>{stat.label}</p>
                        </div>
                      ))}
                    </div>
                  </div>
                );
              })}
            </div>
          </div>
        )}
        {textStats.length > 0 && (
          <div>
            <h3 className={`text-[10px] font-black uppercase tracking-widest ${sub} mb-3`}>📝 Text Columns</h3>
            <div className="space-y-3">
              {textStats.map(({ col, count, unique, top }) => (
                <div key={col} className={`p-4 rounded-2xl border ${card}`}>
                  <div className="flex justify-between items-start mb-2">
                    <p className="font-black text-sm text-indigo-400 uppercase">{col}</p>
                    <span className={`text-[9px] font-black ${sub}`}>{count} filled · {unique} unique</span>
                  </div>
                  <div className="flex flex-wrap gap-1.5">
                    {top.map(([val, cnt]) => (
                      <span key={val} className={`text-[9px] font-black px-2 py-1 rounded-full ${isDark ? "bg-slate-700 text-slate-300" : "bg-slate-200 text-slate-700"}`}>{val} ({cnt})</span>
                    ))}
                  </div>
                </div>
              ))}
            </div>
          </div>
        )}
        {numericCols.length === 0 && textStats.length === 0 && <p className={`text-center py-8 ${sub}`}>No data to analyze yet.</p>}
      </div>
    </div>
  );
}

// ============================================================
// COLUMN TYPE MODAL
// ============================================================
function ColumnTypeModal({ col, currentType, currentOptions, onSave, onClose, isDark }) {
  const [type, setType]           = useState(currentType || "text");
  const [optionsStr, setOptionsStr] = useState(currentOptions ? currentOptions.join(", ") : "");
  const bg     = isDark ? "bg-slate-900" : "bg-white";
  const border = isDark ? "border-white/10" : "border-slate-200";
  const text   = isDark ? "text-slate-100" : "text-slate-900";
  const input  = isDark ? "bg-slate-800 border-white/10 text-slate-200" : "bg-slate-100 border-slate-300 text-slate-900";
  const types  = [
    { value: "text",     label: "📝 Text",     desc: "Free-form text input" },
    { value: "number",   label: "🔢 Number",   desc: "Numeric values only" },
    { value: "email",    label: "📧 Email",    desc: "Validates email format" },
    { value: "date",     label: "📅 Date",     desc: "Date picker input" },
    { value: "checkbox", label: "☑️ Checkbox", desc: "Boolean true/false" },
    { value: "dropdown", label: "📋 Dropdown", desc: "Select from options" },
  ];
  return (
    <div className="fixed inset-0 z-[200] flex items-center justify-center bg-black/70 backdrop-blur-sm p-4">
      <div className={`${bg} ${text} w-full max-w-md rounded-[2rem] border ${border} shadow-2xl p-8`}>
        <div className="flex justify-between items-center mb-6">
          <h3 className="font-black text-lg uppercase">Column Type: {col}</h3>
          <button onClick={onClose} className="text-xl hover:opacity-60">✕</button>
        </div>
        <div className="grid grid-cols-2 gap-2 mb-6">
          {types.map(t => (
            <button key={t.value} onClick={() => setType(t.value)}
              className={`p-3 rounded-2xl border text-left transition-all ${
                type === t.value
                  ? "border-indigo-500 bg-indigo-500/10"
                  : `border-transparent ${isDark ? "bg-slate-800 hover:bg-slate-700" : "bg-slate-100 hover:bg-slate-200"}`
              }`}>
              <div className="text-sm font-black">{t.label}</div>
              <div className={`text-[10px] mt-1 ${isDark ? "text-slate-500" : "text-slate-400"}`}>{t.desc}</div>
            </button>
          ))}
        </div>
        {type === "dropdown" && (
          <div className="mb-6">
            <label className={`text-[10px] font-black uppercase tracking-widest ${isDark ? "text-slate-400" : "text-slate-500"} mb-2 block`}>Options (comma-separated)</label>
            <input type="text" value={optionsStr} onChange={e => setOptionsStr(e.target.value)} placeholder="Pending, In Progress, Done"
              className={`w-full px-4 py-3 rounded-2xl border outline-none text-sm ${input} focus:border-indigo-500`} />
          </div>
        )}
        <div className="flex gap-3">
          <button onClick={onClose} className={`flex-1 py-3 rounded-2xl font-black text-[11px] uppercase border ${isDark ? "border-white/10 hover:bg-slate-800" : "border-slate-200 hover:bg-slate-100"} transition-all`}>Cancel</button>
          <button onClick={() => onSave(type, optionsStr.split(",").map(s => s.trim()).filter(Boolean))} className="flex-1 py-3 rounded-2xl font-black text-[11px] uppercase bg-indigo-600 hover:bg-indigo-500 text-white transition-all">Save Type</button>
        </div>
      </div>
    </div>
  );
}

// ============================================================
// CELL FORMAT MODAL
// ============================================================
function CellFormatModal({ col, rowIdx, scope, currentCell, currentColFormat, onSave, onClose, isDark, isPremium }) {
  const current       = scope === "cell" ? (currentCell || {}) : (currentColFormat || {});
  const [fontSize,   setFontSize]   = useState(current.fontSize   || 14);
  const [color,      setColor]      = useState(current.color      || (isDark ? "#e2e8f0" : "#0f172a"));
  const [bgColor,    setBgColor]    = useState(current.bgColor    || "");
  const [fontWeight, setFontWeight] = useState(current.fontWeight || "normal");

  const bg     = isDark ? "bg-slate-900" : "bg-white";
  const border = isDark ? "border-white/10" : "border-slate-200";
  const text   = isDark ? "text-slate-100" : "text-slate-900";
  const sub    = isDark ? "text-slate-400" : "text-slate-500";
  const premiumColors = ["#ff6b6b","#ffd93d","#6bcb77","#4d96ff","#c77dff","#ff9a3c","#00f5d4"];
  const title  = scope === "cell" ? `Cell [${col}, row ${rowIdx + 1}]` : scope === "tab" ? "Tab Color" : `Column: ${col}`;
  const bold   = fontWeight === "bold";

  return (
    <div className="fixed inset-0 z-[200] flex items-center justify-center bg-black/70 backdrop-blur-sm p-4">
      <div className={`${bg} ${text} w-full max-w-sm rounded-[2rem] border ${border} shadow-2xl p-8`}>
        <div className="flex justify-between items-center mb-6">
          <h3 className="font-black text-lg uppercase">Format: {title}</h3>
          <button onClick={onClose} className="text-xl hover:opacity-60">✕</button>
        </div>
        <div className="space-y-5">
          {scope !== "tab" && (
            <>
              <div>
                <label className={`text-[10px] font-black uppercase tracking-widest ${sub} mb-2 block`}>Font Size: {fontSize}px</label>
                <input type="range" min="10" max="24" value={fontSize} onChange={e => setFontSize(Number(e.target.value))} className="w-full accent-indigo-500" />
              </div>
              <div className="flex items-center gap-3">
                <label className={`text-[10px] font-black uppercase tracking-widest ${sub}`}>Bold</label>
                <button onClick={() => setFontWeight(b => b === "bold" ? "normal" : "bold")}
                  className={`w-10 h-6 rounded-full transition-all ${bold ? "bg-indigo-500" : isDark ? "bg-slate-700" : "bg-slate-200"}`}>
                  <div className={`w-4 h-4 rounded-full bg-white mx-0.5 shadow transition-transform ${bold ? "translate-x-4" : ""}`}></div>
                </button>
              </div>
              <div>
                <label className={`text-[10px] font-black uppercase tracking-widest ${sub} mb-2 block`}>Font Color</label>
                <div className="flex items-center gap-3 flex-wrap">
                  <input type="color" value={color} onChange={e => setColor(e.target.value)} className="w-10 h-10 rounded-xl border-0 cursor-pointer bg-transparent" />
                  <span className="text-xs font-mono">{color}</span>
                  {isPremium && premiumColors.map(c => (
                    <button key={c} onClick={() => setColor(c)} style={{ background: c }} className="w-5 h-5 rounded-full border-2 border-white/20 hover:scale-125 transition-transform" />
                  ))}
                  {!isPremium && <span className="text-[9px] text-indigo-400 font-black uppercase">⭐ Premium colors</span>}
                </div>
              </div>
            </>
          )}
          <div>
            <label className={`text-[10px] font-black uppercase tracking-widest ${sub} mb-2 block`}>
              {scope === "tab" ? "Tab Background Color" : "Background Color"}
            </label>
            <div className="flex items-center gap-3 flex-wrap">
              <input type="color" value={bgColor || "#1e293b"} onChange={e => setBgColor(e.target.value)} className="w-10 h-10 rounded-xl border-0 cursor-pointer bg-transparent" />
              {premiumColors.map(c => (
                <button key={c} onClick={() => setBgColor(c)} style={{ background: c }} className="w-5 h-5 rounded-full border-2 border-white/20 hover:scale-125 transition-transform" />
              ))}
              <button onClick={() => setBgColor("")} className="text-[10px] text-red-400 hover:text-red-300 font-black uppercase">Clear</button>
            </div>
          </div>
          {scope !== "tab" && (
            <div className={`p-4 rounded-2xl border ${border} text-center`}
              style={{ fontSize: `${fontSize}px`, color, backgroundColor: bgColor || "transparent", fontWeight }}>
              Preview Text Sample
            </div>
          )}
        </div>
        <div className="flex gap-3 mt-6">
          <button onClick={onClose} className={`flex-1 py-3 rounded-2xl font-black text-[11px] uppercase border ${isDark ? "border-white/10 hover:bg-slate-800" : "border-slate-200 hover:bg-slate-100"} transition-all`}>Cancel</button>
          <button onClick={() => onSave({ fontSize, color, bgColor, fontWeight })} className="flex-1 py-3 rounded-2xl font-black text-[11px] uppercase bg-indigo-600 hover:bg-indigo-500 text-white transition-all">Apply</button>
        </div>
      </div>
    </div>
  );
}

function evaluateFormula(formula, rows, columns) {
  if (!formula || !formula.startsWith("=")) return formula;
  
  try {
    const expr = formula.slice(1).toUpperCase().trim();
    
    // Helper: get cell value by reference like A1, B2
    const getCellVal = (ref) => {
      const match = ref.match(/^([A-Z]+)(\d+)$/);
      if (!match) return 0;
      const colIndex = match[1].split("").reduce((acc, c) => acc * 26 + c.charCodeAt(0) - 64, 0) - 1;
      const rowIndex = parseInt(match[2]) - 1;
      if (rowIndex < 0 || rowIndex >= rows.length) return 0;
      const col = columns[colIndex];
      if (!col) return 0;
      const val = parseFloat(cellVal(rows[rowIndex]?.[col]));
      return isNaN(val) ? 0 : val;
    };
 
    // Helper: get range values like A1:A10
    const getRangeVals = (range) => {
      const parts = range.split(":");
      if (parts.length !== 2) return [];
      const startMatch = parts[0].match(/^([A-Z]+)(\d+)$/);
      const endMatch   = parts[1].match(/^([A-Z]+)(\d+)$/);
      if (!startMatch || !endMatch) return [];
      const startCol = startMatch[1].split("").reduce((acc, c) => acc * 26 + c.charCodeAt(0) - 64, 0) - 1;
      const endCol   = endMatch[1].split("").reduce((acc, c) => acc * 26 + c.charCodeAt(0) - 64, 0) - 1;
      const startRow = parseInt(startMatch[2]) - 1;
      const endRow   = parseInt(endMatch[2]) - 1;
      const vals = [];
      for (let r = startRow; r <= endRow; r++) {
        for (let c = startCol; c <= endCol; c++) {
          const col = columns[c];
          if (!col) continue;
          const v = parseFloat(cellVal(rows[r]?.[col]));
          if (!isNaN(v)) vals.push(v);
        }
      }
      return vals;
    };
 
    // ── SUM ──
    if (expr.startsWith("SUM(")) {
      const inner = expr.slice(4, -1);
      const vals = getRangeVals(inner);
      return vals.reduce((a, b) => a + b, 0).toString();
    }
 
    // ── AVERAGE / AVG ──
    if (expr.startsWith("AVERAGE(") || expr.startsWith("AVG(")) {
      const inner = expr.startsWith("AVG(") ? expr.slice(4, -1) : expr.slice(8, -1);
      const vals = getRangeVals(inner);
      if (!vals.length) return "0";
      return (vals.reduce((a, b) => a + b, 0) / vals.length).toFixed(2);
    }
 
    // ── MIN ──
    if (expr.startsWith("MIN(")) {
      const inner = expr.slice(4, -1);
      const vals = getRangeVals(inner);
      return vals.length ? Math.min(...vals).toString() : "0";
    }
 
    // ── MAX ──
    if (expr.startsWith("MAX(")) {
      const inner = expr.slice(4, -1);
      const vals = getRangeVals(inner);
      return vals.length ? Math.max(...vals).toString() : "0";
    }
 
    // ── COUNT ──
    if (expr.startsWith("COUNT(")) {
      const inner = expr.slice(6, -1);
      return getRangeVals(inner).length.toString();
    }
 
    // ── IF ──
    if (expr.startsWith("IF(")) {
      const inner = expr.slice(3, -1);
      const parts = inner.split(",");
      if (parts.length < 3) return "ERROR";
      const condition = parts[0].trim();
      const ifTrue    = parts[1].trim().replace(/"/g, "");
      const ifFalse   = parts[2].trim().replace(/"/g, "");
      // Simple conditions: A1>10, A1=5, A1<100
      const condMatch = condition.match(/^([A-Z]+\d+)\s*([><=!]+)\s*(.+)$/);
      if (condMatch) {
        const cellValue = getCellVal(condMatch[1]);
        const operator  = condMatch[2];
        const compareVal = parseFloat(condMatch[3]);
        let result = false;
        if (operator === ">")  result = cellValue >  compareVal;
        if (operator === "<")  result = cellValue <  compareVal;
        if (operator === ">=") result = cellValue >= compareVal;
        if (operator === "<=") result = cellValue <= compareVal;
        if (operator === "=")  result = cellValue === compareVal;
        if (operator === "!=") result = cellValue !== compareVal;
        return result ? ifTrue : ifFalse;
      }
      return "ERROR";
    }
 
    // ── CONCAT ──
    if (expr.startsWith("CONCAT(")) {
      const inner = expr.slice(7, -1);
      const parts = inner.split(",").map(p => p.trim().replace(/"/g, ""));
      return parts.map(p => {
        if (/^[A-Z]+\d+$/.test(p)) {
          const colIdx = p.match(/^([A-Z]+)/)[1].split("").reduce((acc, c) => acc * 26 + c.charCodeAt(0) - 64, 0) - 1;
          const rowIdx = parseInt(p.match(/(\d+)$/)[1]) - 1;
          const col = columns[colIdx];
          return col ? cellVal(rows[rowIdx]?.[col]) : "";
        }
        return p;
      }).join("");
    }
 
    // ── ROUND ──
    if (expr.startsWith("ROUND(")) {
      const inner = expr.slice(6, -1);
      const parts = inner.split(",");
      const val  = getCellVal(parts[0].trim()) || parseFloat(parts[0]);
      const dec  = parseInt(parts[1]?.trim() || "0");
      return parseFloat(val.toFixed(dec)).toString();
    }
 
    // ── Simple arithmetic: =A1+B1, =A1*2 ──
    const arithmeticExpr = expr.replace(/[A-Z]+\d+/g, (ref) => getCellVal(ref));
    // eslint-disable-next-line no-new-func
    const result = new Function("return " + arithmeticExpr)();
    return isNaN(result) ? "ERROR" : result.toString();
 
  } catch (e) {
    return "ERROR";
  }
}
// ============================================================
// UTILITY
// ============================================================
const hasNonAscii = (text) => {
  return Array.from(String(text ?? "")).some(ch => ch.charCodeAt(0) > 127);
};

const hexToRgb = (hex) => {
  if (!hex || !hex.startsWith("#")) return [15, 23, 42];
  const h = hex.replace("#", "");
  return [parseInt(h.substring(0, 2), 16) || 0, parseInt(h.substring(2, 4), 16) || 0, parseInt(h.substring(4, 6), 16) || 0];
};

function BorderModal({ selectedCell, rows, activeTabId, setTabs, onClose, isDark }) {
  const cell = rows[selectedCell?.rIdx]?.[selectedCell?.col];
  const cellObj = cell && typeof cell === "object" ? cell : {};
 
  const [bTop,    setBTop]    = useState(cellObj.borderTop    || "");
  const [bBottom, setBBottom] = useState(cellObj.borderBottom || "");
  const [bLeft,   setBLeft]   = useState(cellObj.borderLeft   || "");
  const [bRight,  setBRight]  = useState(cellObj.borderRight  || "");
  const [bColor,  setBColor]  = useState("#6366f1");
  const [bWidth,  setBWidth]  = useState("1");
  const [bStyle,  setBStyle]  = useState("solid");
 
  const makeBorder = () => `${bWidth}px ${bStyle} ${bColor}`;
 
  const handleApply = () => {
    const { rIdx, col } = selectedCell;
    setTabs(prev => prev.map(t => {
      if (t.id !== activeTabId) return t;
      const updatedRows = [...t.rows];
      const existing = updatedRows[rIdx]?.[col] || {};
      const existingObj = typeof existing === "object" ? existing : { value: String(existing) };
      updatedRows[rIdx] = {
        ...updatedRows[rIdx],
        [col]: { ...existingObj, borderTop: bTop, borderBottom: bBottom, borderLeft: bLeft, borderRight: bRight },
      };
      return { ...t, rows: updatedRows };
    }));
    onClose();
  };
 
  return (
    <div className="fixed inset-0 z-[300] flex items-center justify-center bg-black/70 backdrop-blur-sm p-4">
      <div className={`w-full max-w-sm rounded-[2rem] border shadow-2xl p-8 ${isDark ? "bg-[#0a1628] border-[#1e3a5f] text-slate-100" : "bg-white border-slate-200 text-slate-900"}`}>
        <div className="flex justify-between items-center mb-6">
          <h3 className="font-black text-lg uppercase">Cell Border</h3>
          <button onClick={onClose} className="text-xl hover:opacity-60">✕</button>
        </div>
 
        <div className="grid grid-cols-3 gap-3 mb-5">
          <div>
            <label className={`text-[9px] font-black uppercase tracking-widest block mb-1 ${isDark ? "text-slate-400" : "text-slate-500"}`}>Color</label>
            <input type="color" value={bColor} onChange={e => setBColor(e.target.value)} className="w-full h-9 rounded-xl border-0 cursor-pointer" />
          </div>
          <div>
            <label className={`text-[9px] font-black uppercase tracking-widest block mb-1 ${isDark ? "text-slate-400" : "text-slate-500"}`}>Width</label>
            <select value={bWidth} onChange={e => setBWidth(e.target.value)}
              className={`w-full h-9 rounded-xl px-2 text-xs font-black border outline-none ${isDark ? "bg-[#050d1f] border-[#1e3a5f] text-slate-200" : "bg-slate-50 border-slate-200"}`}>
              <option value="1">1px</option>
              <option value="2">2px</option>
              <option value="3">3px</option>
            </select>
          </div>
          <div>
            <label className={`text-[9px] font-black uppercase tracking-widest block mb-1 ${isDark ? "text-slate-400" : "text-slate-500"}`}>Style</label>
            <select value={bStyle} onChange={e => setBStyle(e.target.value)}
              className={`w-full h-9 rounded-xl px-2 text-xs font-black border outline-none ${isDark ? "bg-[#050d1f] border-[#1e3a5f] text-slate-200" : "bg-slate-50 border-slate-200"}`}>
              <option value="solid">Solid</option>
              <option value="dashed">Dashed</option>
              <option value="dotted">Dotted</option>
              <option value="double">Double</option>
            </select>
          </div>
        </div>
 
        {/* Preview */}
        <div className="flex items-center justify-center mb-5">
          <div className="w-24 h-24 flex items-center justify-center text-[10px] font-black text-slate-400"
            style={{
              borderTop:    bTop    || "1px dashed #334155",
              borderBottom: bBottom || "1px dashed #334155",
              borderLeft:   bLeft   || "1px dashed #334155",
              borderRight:  bRight  || "1px dashed #334155",
            }}>
            Preview
          </div>
        </div>
 
        {/* Side buttons */}
        <div className="grid grid-cols-2 gap-2 mb-4">
          {[
            { label: "Top",    val: bTop,    set: setBTop },
            { label: "Bottom", val: bBottom, set: setBBottom },
            { label: "Left",   val: bLeft,   set: setBLeft },
            { label: "Right",  val: bRight,  set: setBRight },
          ].map(({ label, val, set }) => (
            <button key={label} onClick={() => set(val ? "" : makeBorder())}
              className={`py-2 rounded-xl font-black text-[10px] uppercase transition-all border ${
                val
                  ? isDark ? "bg-blue-600/20 border-blue-500/40 text-blue-400" : "bg-blue-50 border-blue-300 text-blue-600"
                  : isDark ? "border-[#1e3a5f] text-slate-500 hover:text-slate-300" : "border-slate-200 text-slate-400 hover:text-slate-600"
              }`}>
              {label} {val ? "On" : "Off"}
            </button>
          ))}
        </div>
 
        {/* Presets */}
        <div className="flex gap-2 mb-5">
          <button onClick={() => { const b = makeBorder(); setBTop(b); setBBottom(b); setBLeft(b); setBRight(b); }}
            className={`flex-1 py-2 rounded-xl font-black text-[10px] uppercase transition-all ${isDark ? "bg-blue-600/20 text-blue-400 hover:bg-blue-600 hover:text-white" : "bg-blue-50 text-blue-600 hover:bg-blue-600 hover:text-white"}`}>
            All Sides
          </button>
          <button onClick={() => { setBTop(""); setBBottom(""); setBLeft(""); setBRight(""); }}
            className={`flex-1 py-2 rounded-xl font-black text-[10px] uppercase transition-all ${isDark ? "bg-red-500/10 text-red-400 hover:bg-red-500 hover:text-white" : "bg-red-50 text-red-500 hover:bg-red-500 hover:text-white"}`}>
            Clear
          </button>
        </div>
 
        <div className="flex gap-3">
          <button onClick={onClose}
            className={`flex-1 py-3 rounded-2xl font-black text-xs uppercase border transition-all ${isDark ? "border-[#1e3a5f] hover:bg-white/5 text-slate-400" : "border-slate-200 hover:bg-slate-50 text-slate-500"}`}>
            Cancel
          </button>
          <button onClick={handleApply}
            className="flex-1 py-3 rounded-2xl font-black text-xs uppercase bg-blue-600 hover:bg-blue-500 text-white transition-all">
            Apply
          </button>
        </div>
      </div>
    </div>
  );
}
 
function ConditionalModal({ columns, conditionalRules, onApply, onClose, isDark }) {
  const [rules, setRules] = useState(conditionalRules);
 
  const addRule = () => setRules(prev => [...prev, { col: "all", type: "gt", value: "", color: "#ef4444", bgColor: "" }]);
  const removeRule = (i) => setRules(prev => prev.filter((_, idx) => idx !== i));
  const updateRule = (i, key, val) => setRules(prev => prev.map((r, idx) => idx === i ? { ...r, [key]: val } : r));
 
  return (
    <div className="fixed inset-0 z-[300] flex items-center justify-center bg-black/70 backdrop-blur-sm p-4">
      <div className={`w-full max-w-lg rounded-[2rem] border shadow-2xl overflow-hidden ${isDark ? "bg-[#0a1628] border-[#1e3a5f] text-slate-100" : "bg-white border-slate-200 text-slate-900"}`}
        style={{ maxHeight: "85vh", display: "flex", flexDirection: "column" }}>
 
        <div className={`px-7 pt-7 pb-5 border-b ${isDark ? "border-[#1e3a5f]" : "border-slate-200"} flex-shrink-0`}>
          <div className="flex justify-between items-center">
            <div>
              <h3 className="font-black text-lg uppercase">Conditional Formatting</h3>
              <p className={`text-[10px] mt-1 ${isDark ? "text-slate-400" : "text-slate-500"}`}>Rules apply automatically to matching cells</p>
            </div>
            <button onClick={onClose} className="text-xl hover:opacity-60">✕</button>
          </div>
        </div>
 
        <div className="overflow-y-auto flex-1 p-6 space-y-3">
          {rules.length === 0 && (
            <div className="text-center py-8">
              <p className={`text-sm font-black uppercase ${isDark ? "text-slate-500" : "text-slate-400"}`}>No rules yet</p>
              <p className={`text-[10px] mt-2 ${isDark ? "text-slate-600" : "text-slate-400"}`}>Click Add Rule to get started</p>
            </div>
          )}
          {rules.map((rule, i) => (
            <div key={i} className={`p-4 rounded-2xl border ${isDark ? "bg-[#070f1e] border-[#1e3a5f]" : "bg-slate-50 border-slate-200"}`}>
              <div className="flex items-center gap-2 mb-3 flex-wrap">
                <select value={rule.col} onChange={e => updateRule(i, "col", e.target.value)}
                  className={`px-3 py-2 rounded-xl text-[11px] font-black border outline-none ${isDark ? "bg-[#050d1f] border-[#1e3a5f] text-slate-200" : "bg-white border-slate-200"}`}>
                  <option value="all">All Columns</option>
                  {columns.filter(c => c !== "ID").map(c => <option key={c} value={c}>{c}</option>)}
                </select>
 
                <select value={rule.type} onChange={e => updateRule(i, "type", e.target.value)}
                  className={`px-3 py-2 rounded-xl text-[11px] font-black border outline-none ${isDark ? "bg-[#050d1f] border-[#1e3a5f] text-slate-200" : "bg-white border-slate-200"}`}>
                  <option value="gt">Greater than</option>
                  <option value="lt">Less than</option>
                  <option value="eq">Equal to</option>
                  <option value="contains">Contains</option>
                  <option value="empty">Is Empty</option>
                </select>
 
                {rule.type !== "empty" && (
                  <input type="text" value={rule.value} onChange={e => updateRule(i, "value", e.target.value)}
                    placeholder="Value..."
                    className={`px-3 py-2 rounded-xl text-[11px] border outline-none w-24 ${isDark ? "bg-[#050d1f] border-[#1e3a5f] text-slate-200 placeholder:text-slate-600" : "bg-white border-slate-200"}`} />
                )}
 
                <button onClick={() => removeRule(i)} className="text-red-400 hover:text-red-300 text-lg ml-auto">✕</button>
              </div>
 
              <div className="flex items-center gap-3 flex-wrap">
                <label className={`text-[9px] font-black uppercase tracking-widest ${isDark ? "text-slate-400" : "text-slate-500"}`}>Text</label>
                <input type="color" value={rule.color || "#ef4444"} onChange={e => updateRule(i, "color", e.target.value)} className="w-8 h-8 rounded-lg border-0 cursor-pointer" />
                <label className={`text-[9px] font-black uppercase tracking-widest ${isDark ? "text-slate-400" : "text-slate-500"}`}>Background</label>
                <input type="color" value={rule.bgColor || "#ffffff"} onChange={e => updateRule(i, "bgColor", e.target.value)} className="w-8 h-8 rounded-lg border-0 cursor-pointer" />
                {[
                  { label: "Red",    color: "#ef4444", bg: "#fef2f2" },
                  { label: "Green",  color: "#10b981", bg: "#f0fdf4" },
                  { label: "Blue",   color: "#3b82f6", bg: "#eff6ff" },
                  { label: "Yellow", color: "#f59e0b", bg: "#fffbeb" },
                ].map(preset => (
                  <button key={preset.label}
                    onClick={() => { updateRule(i, "color", preset.color); updateRule(i, "bgColor", preset.bg); }}
                    className="w-5 h-5 rounded-full border-2 border-white/20 hover:scale-125 transition-transform"
                    style={{ backgroundColor: preset.color }}
                    title={preset.label} />
                ))}
              </div>
            </div>
          ))}
        </div>
 
        <div className={`px-6 pb-6 pt-4 border-t ${isDark ? "border-[#1e3a5f]" : "border-slate-200"} flex-shrink-0`}>
          <button onClick={addRule}
            className={`w-full py-3 rounded-2xl font-black text-[11px] uppercase border transition-all mb-3 ${isDark ? "border-[#1e3a5f] text-blue-400 hover:bg-blue-500/10" : "border-blue-200 text-blue-600 hover:bg-blue-50"}`}>
            + Add Rule
          </button>
          <div className="flex gap-3">
            <button onClick={() => { onApply([]); onClose(); }}
              className={`flex-1 py-3 rounded-2xl font-black text-xs uppercase border transition-all ${isDark ? "border-[#1e3a5f] text-slate-400 hover:bg-white/5" : "border-slate-200 text-slate-500 hover:bg-slate-50"}`}>
              Clear All
            </button>
            <button onClick={() => { onApply(rules); onClose(); }}
              className="flex-1 py-3 rounded-2xl font-black text-xs uppercase bg-blue-600 hover:bg-blue-500 text-white transition-all">
              Apply Rules
            </button>
          </div>
        </div>
      </div>
    </div>
  );
}
// ============================================================
// MAIN APP
// ============================================================
function App() {
  const [genText,            setGenText]            = useState("");
  const [inputText,          setInputText]          = useState("");
  const [searchTerm,         setSearchTerm]         = useState("");
  const [user,               setUser]               = useState(null);
  const [page,               setPage]               = useState("landing");
  const [isDark,             setIsDark]             = useState(true);
  const [showAdmin,          setShowAdmin]          = useState(false);
  const [showAnalytics,      setShowAnalytics]      = useState(false);
  const [showUpgrade,        setShowUpgrade]        = useState(false);
  const [showAIAnalysis,     setShowAIAnalysis]     = useState(false);
  const [showTemplates,      setShowTemplates]      = useState(false);
  const [isPremium,          setIsPremium]          = useState(false);
  const [isAdmin,            setIsAdmin]            = useState(false);
  const [syncStatus,         setSyncStatus]         = useState("idle");
  const [lastSaved,          setLastSaved]          = useState(null);
  const [isLoading,          setIsLoading]          = useState(false);
  const [calculationResult,  setCalculationResult]  = useState(null);
  const [selectedRows,       setSelectedRows]       = useState(new Set());
  const [dragRowIdx,         setDragRowIdx]         = useState(null);
  const [dragOverRowIdx,     setDragOverRowIdx]     = useState(null);
  const [dragColIdx,         setDragColIdx]         = useState(null);
  const [dragOverColIdx,     setDragOverColIdx]     = useState(null);
  const [menuConfig,         setMenuConfig]         = useState({ show: false, x: 0, y: 0, col: null });
  const [editingColType,     setEditingColType]     = useState(null);
  const [editingFormat,      setEditingFormat]      = useState(null);
  const [history,            setHistory]            = useState([]);
  const [currentTableId,     setCurrentTableId]     = useState(null);
  const [tabs,               setTabs]               = useState([makeNewTab("Sheet 1", ["ID", "Name", "Email", "Status"], 8)]);
  const [activeTabId,        setActiveTabId]        = useState(null);
  const [showAITable, setShowAITable] = useState(false);
  const [aiTableUsage, setAiTableUsage] = useState(0);
  const [aiBuildLoading, setAiBuildLoading] = useState(false);
  const [lifetimeTablesCreated, setLifetimeTablesCreated] = useState(1);

  // ── NEW: track table creation timestamps for 24h cooldown ──
  // Shape: [{ id, createdAt: timestamp_ms }, ...]
  const [tableCreationLog,   setTableCreationLog]   = useState([]);
  const [showProGate,    setShowProGate]    = useState(false);
  const [showWelcome,    setShowWelcome]    = useState(false);
  const [proGateMessage, setProGateMessage] = useState("");

  const [colWidths,      setColWidths]      = useState({});
  const [rowHeights,     setRowHeights]     = useState({});
  const [copiedCell,     setCopiedCell]     = useState(null);
  const [freezeCol,      setFreezeCol]      = useState(true);
  const [selectedCell, setSelectedCell] = useState(null);
  const [selectionStart, setSelectionStart] = useState(null);
  const [selectionEnd, setSelectionEnd] = useState(null);
  const [isSelecting, setIsSelecting] = useState(false);
  const [showConditional, setShowConditional] = useState(false);
  const [conditionalRules, setConditionalRules] = useState([]);
  const [showBorderModal, setShowBorderModal] = useState(false);
  const [undoStack,      setUndoStack]      = useState([]);
  const [redoStack,      setRedoStack]      = useState([]);
  const [frozenCols,     setFrozenCols]     = useState(1);  // কতটা column freeze
  const [showFindReplace, setShowFindReplace] = useState(false);
  const [hiddenRows,     setHiddenRows]     = useState(new Set());
  const [hiddenCols,     setHiddenCols]     = useState(new Set());
  const [formulaBarVal,  setFormulaBarVal]  = useState("");
  const showFormulaBar = true;
  const [findText,       setFindText]       = useState("");
  const [replaceText,    setReplaceText]    = useState("");
  const [findResults,    setFindResults]    = useState([]);
  const [findIndex,      setFindIndex]      = useState(0);
  const resizingCol      = useRef(null);
  const resizingRow      = useRef(null);
  const startX           = useRef(0);
  const startY           = useRef(0);
  const startWidth       = useRef(0);
  const startHeight      = useRef(0);

  const menuRef        = useRef(null);
  const tableRef       = useRef(null);
  const fileInputRef   = useRef(null);
  const autoSaveTimer  = useRef(null);
  const firestoreUnsubRef = useRef(null);
  const isInitialLoad  = useRef(true);
  const userRef        = useRef(null);

  useEffect(() => {
    if (!activeTabId && tabs.length > 0) setActiveTabId(tabs[0].id);
  }, [activeTabId, tabs]);

  const activeTab = useMemo(
    () => tabs.find(t => t.id === activeTabId) || tabs[0],
    [tabs, activeTabId]
  );
  const columns      = useMemo(() => activeTab?.columns || [], [activeTab]);
  const rows         = useMemo(() => activeTab?.rows || [], [activeTab]);
  const tableTitle   = activeTab?.title        || "";
  const columnTypes  = useMemo(() => activeTab?.columnTypes || {}, [activeTab]);
  const columnFormats = useMemo(() => activeTab?.columnFormats || {}, [activeTab]);

  const updateActiveTab = useCallback((updater) => {
    setTabs(prev => prev.map(t => {
      if (t.id !== activeTabId) return t;
      return typeof updater === "function" ? updater(t) : { ...t, ...updater };
    }));
  }, [activeTabId]);

  const setColumns    = useCallback((cols) => updateActiveTab({ columns: cols }),  [updateActiveTab]);
  const setTableTitle = useCallback((title) => updateActiveTab({ title }),          [updateActiveTab]);

  const handleCellChange = useCallback((rIdx, col, value) => {
    setTabs(prev => prev.map(t => {
      if (t.id !== activeTabId) return t;
      const updatedRows = [...t.rows];
      const existingCell = updatedRows[rIdx]?.[col] || {};
      updatedRows[rIdx] = {
        ...updatedRows[rIdx],
        [col]: { ...(typeof existingCell === "object" ? existingCell : {}), value: String(value ?? "") },
      };
      return { ...t, rows: updatedRows };
    }));
  }, [activeTabId]);

  const applyCellFormat = useCallback((rIdx, col, fmt) => {
    setTabs(prev => prev.map(t => {
      if (t.id !== activeTabId) return t;
      const updatedRows = [...t.rows];
      const existingCell = updatedRows[rIdx]?.[col] || {};
      updatedRows[rIdx] = {
        ...updatedRows[rIdx],
        [col]: {
          value:      typeof existingCell === "object" ? (existingCell.value ?? "") : String(existingCell),
          color:      fmt.color      || "",
          bgColor:    fmt.bgColor    || "",
          fontSize:   fmt.fontSize   || "",
          fontWeight: fmt.fontWeight || "normal",
        },
      };
      return { ...t, rows: updatedRows };
    }));
  }, [activeTabId]);

  const applyColumnFormat = useCallback((col, fmt) => {
    updateActiveTab(t => ({ ...t, columnFormats: { ...t.columnFormats, [col]: fmt } }));
  }, [updateActiveTab]);

  const setColumnTypes = useCallback((updater) => {
    updateActiveTab(t => ({
      ...t,
      columnTypes: typeof updater === "function" ? updater(t.columnTypes || {}) : updater,
    }));
  }, [updateActiveTab]);

  const setRows = useCallback((rws) => updateActiveTab({ rows: rws }), [updateActiveTab]);
  // ============================================================
  // FREE PLAN LIFETIME LIMIT
  // ============================================================
  const canCreateTable = useCallback(() => {
  if (isPremium) return { allowed: true, reason: "" };

  if (lifetimeTablesCreated >= FREE_TABLE_LIMIT) {
    return {
      allowed: false,
      reason: `Free plan allows ${FREE_TABLE_LIMIT} lifetime tables per Gmail. You already used all free tables. Upgrade to Pro for unlimited tables.`,
    };
  }

  if (tabs.length >= FREE_TABLE_LIMIT) {
    return {
      allowed: false,
      reason: `Free plan allows max ${FREE_TABLE_LIMIT} open sheets. Delete one to continue, but lifetime table count will not reset.`,
    };
  }

  return { allowed: true, reason: "" };
}, [isPremium, lifetimeTablesCreated, tabs.length]);

  const getFreshLifetimeTableCount = useCallback(async () => {
  if (!user?.uid || !db) return lifetimeTablesCreated;

  try {
    const snap = await getDoc(doc(db, "users", user.uid));
    const count = snap.exists() ? snap.data().lifetimeTablesCreated : null;
    return typeof count === "number" ? count : lifetimeTablesCreated;
  } catch (err) {
    console.error("Lifetime table count check failed:", err);
    return lifetimeTablesCreated;
  }
}, [user?.uid, lifetimeTablesCreated]);

  // ============================================================
  // AUTO-SAVE
  // ============================================================
  const saveToFirestore = useCallback(async (tabsData, histData, creationLog) => {
    const currentUser = userRef.current;
    try {
      localStorage.setItem("ss_tabs_v14",         JSON.stringify(tabsData));
      localStorage.setItem("ss_history_v14",      JSON.stringify(histData));
      localStorage.setItem("ss_creation_log_v14", JSON.stringify(creationLog));
      localStorage.setItem("ss_ai_usage_v15", JSON.stringify(aiTableUsage));
    } catch (e) {}
    if (currentUser && db) {
      try {
        await setDoc(doc(db, "spreadsheets", currentUser.uid), {
          tabs:            tabsData,
          history:         histData,
          tableCreationLog: creationLog,
          tableCount:      tabsData.length,
          updatedAt:       serverTimestamp(),
          uid:             currentUser.uid,
          aiTableUsage,
        }, { merge: true });
        await setDoc(doc(db, "users", currentUser.uid), { tableCount: tabsData.length }, { merge: true });
        setLastSaved(new Date().toLocaleTimeString());
        setSyncStatus("saved");
      } catch (err) { console.error("Firestore save error:", err); setSyncStatus("local"); }
    } else { setSyncStatus("local"); }
  }, [aiTableUsage]);

  useEffect(() => {
    if (isInitialLoad.current) return;
    setSyncStatus("syncing");
    if (autoSaveTimer.current) clearTimeout(autoSaveTimer.current);
    autoSaveTimer.current = setTimeout(() => {
      saveToFirestore(tabs, history, tableCreationLog);
    }, AUTOSAVE_DELAY);
    
  }, [tabs, history, tableCreationLog, aiTableUsage, saveToFirestore]);

  // ============================================================
  // LOAD DATA
  // ============================================================
  const loadUserDataFromFirestore = useCallback((uid) => {
    if (firestoreUnsubRef.current) firestoreUnsubRef.current();
    if (!db) { loadFromLocalStorage(); return; }
    const userDocRef = doc(db, "spreadsheets", uid);
    const unsub = onSnapshot(userDocRef, (snap) => {
      if (snap.exists()) {
        const data = snap.data();
        if (typeof data.aiTableUsage === "number") setAiTableUsage(data.aiTableUsage);
        isInitialLoad.current = true;
        if (data.tabs && Array.isArray(data.tabs) && data.tabs.length > 0) {
          const migratedTabs = data.tabs.map(tab => ({
            ...tab,
            columnTypes:   tab.columnTypes   || {},
            columnFormats: tab.columnFormats || {},
            tabColor:      tab.tabColor      || "",
            rows: (tab.rows || []).map(row => migrateRow(row, tab.columns || [])),
          }));
          setTabs(migratedTabs);
          setActiveTabId(prevActiveId => {
            const stillExists = migratedTabs.some(tab => tab.id === prevActiveId);
            return stillExists ? prevActiveId : migratedTabs[0].id;
          });
        }
        if (data.history)          setHistory(data.history);
        if (data.tableCreationLog) setTableCreationLog(data.tableCreationLog);
        if (data.isPremium)        setIsPremium(true);
        setLastSaved(data.updatedAt ? new Date(data.updatedAt.seconds * 1000).toLocaleTimeString() : null);
        setSyncStatus("saved");
        setTimeout(() => { isInitialLoad.current = false; }, 600);
      } else {
        isInitialLoad.current = false;
        loadFromLocalStorage();
      }
    }, (err) => {
      console.error("Firestore sync error:", err);
      setSyncStatus("local");
      isInitialLoad.current = false;
      loadFromLocalStorage();
    });
    firestoreUnsubRef.current = unsub;
  }, []);

  const loadFromLocalStorage = () => {
    try {
      const savedTabs        = localStorage.getItem("ss_tabs_v14");
      const savedHistory     = localStorage.getItem("ss_history_v14");
      const savedCreationLog = localStorage.getItem("ss_creation_log_v14");
      if (savedTabs) {
        const parsed = JSON.parse(savedTabs);
        const migratedTabs = parsed.map(tab => ({
          ...tab,
          columnTypes:   tab.columnTypes   || {},
          columnFormats: tab.columnFormats || {},
          tabColor:      tab.tabColor      || "",
          rows: (tab.rows || []).map(row => migrateRow(row, tab.columns || [])),
        }));
        setTabs(migratedTabs);
        if (migratedTabs.length > 0) setActiveTabId(migratedTabs[0].id);
      }
      if (savedHistory)     setHistory(JSON.parse(savedHistory));
      if (savedCreationLog) setTableCreationLog(JSON.parse(savedCreationLog));
      const savedAiUsage = localStorage.getItem("ss_ai_usage_v15");
      if (savedAiUsage) setAiTableUsage(JSON.parse(savedAiUsage));
    } catch (e) { console.error("LocalStorage load error:", e); }
    finally { isInitialLoad.current = false; }
  };

  const checkUserRole = useCallback(async (uid, email) => {
    if (email === ADMIN_EMAIL) { setIsAdmin(true); setIsPremium(true); return; }
    if (db) {
      try {
        const userDoc = await getDoc(doc(db, "users", uid));
        if (!userDoc.exists()) {
          // নতুন user — welcome modal দেখাও
          setShowWelcome(true);
          setLifetimeTablesCreated(1);
        }
        if (userDoc.exists()) {
          const data = userDoc.data();
          if (typeof data.lifetimeTablesCreated === "number") {
            setLifetimeTablesCreated(data.lifetimeTablesCreated);
          } else {
            setLifetimeTablesCreated(1);
          }
          if (data.isPremium) {
            if (data.premiumExpiry) {
              if (new Date(data.premiumExpiry.seconds * 1000) > new Date()) setIsPremium(true);
            } else setIsPremium(true);
          }
        }
      } catch (err) { console.error("Role check error:", err); }
    }
  }, []);

  useEffect(() => {
    if (!auth) { loadFromLocalStorage(); return; }
    const unsub = onAuthStateChanged(auth, async (u) => {
      userRef.current = u;
      if (u) {
        setUser(u); setPage("app");
        await checkUserRole(u.uid, u.email);
        try {
          await setDoc(doc(db, "users", u.uid), 
        {
          name: u.displayName,
          email: u.email,
          photoURL: u.photoURL,
          lastLogin: serverTimestamp(),
          uid: u.uid,
          lifetimeTableLimit: FREE_TABLE_LIMIT,
        }, { merge: true });
        } catch (err) {}
        loadUserDataFromFirestore(u.uid);
      } else {
        loadFromLocalStorage();
      }
    });
    return () => unsub();
  }, [checkUserRole, loadUserDataFromFirestore]);

  useEffect(() => {
    const handleClickOutside = (e) => {
      if (menuRef.current && !menuRef.current.contains(e.target))
        setMenuConfig(p => ({ ...p, show: false }));
    };
    document.addEventListener("mousedown", handleClickOutside);
    return () => {
      document.removeEventListener("mousedown", handleClickOutside);
      if (firestoreUnsubRef.current) firestoreUnsubRef.current();
      if (autoSaveTimer.current)     clearTimeout(autoSaveTimer.current);
    };
  }, []);

  const handleGoogleLogin = async () => {
    if (!auth) { alert("Firebase not initialized."); return; }
    try {
      const result = await signInWithPopup(auth, provider);
      const u = result.user;
      userRef.current = u; setUser(u); setPage("app");
      await checkUserRole(u.uid, u.email);
      try {
        await setDoc(doc(db, "users", u.uid), {
          name: u.displayName, email: u.email, photoURL: u.photoURL,
          lastLogin: serverTimestamp(), uid: u.uid,
        }, { merge: true });
      } catch (err) {}
      loadUserDataFromFirestore(u.uid);
    } catch (err) { console.error("Login error:", err); }
  };

  const handleLogout = () => {
    if (firestoreUnsubRef.current) firestoreUnsubRef.current();
    userRef.current = null;
    if (auth) signOut(auth).then(() => {
      setUser(null); setPage("landing");
      setIsAdmin(false); setIsPremium(false);
    });
  };

  // ============================================================
  // CALCULATION ENGINE
  // ============================================================
  useEffect(() => {
    const performCalculation = () => {
      if (!inputText.trim()) { setCalculationResult(null); return; }
      try {
        let finalValue = 0;
        const inputLower = inputText.toLowerCase();
        let targetRows = rows;
        if (inputLower.includes(" for ") || inputLower.includes(" of ")) {
          const parts = inputLower.split(/for|of/);
          const targets = parts[1].split(",").map(t => t.trim().toLowerCase());
          targetRows = rows.filter(row => {
            const nameMatch = targets.some(t => cellVal(row["Name"]).toLowerCase().includes(t));
            const idMatch   = targets.some(t => cellVal(row["ID"]) === t);
            return nameMatch || idMatch;
          });
        }
        const operators = [
          { symbol: "+", fn: (a, b) => a + b, label: "Sum" },
          { symbol: "-", fn: (a, b) => a - b, label: "Sub" },
          { symbol: "*", fn: (a, b) => a * b, label: "Mul" },
          { symbol: "/", fn: (a, b) => b !== 0 ? a / b : 0, label: "Div" },
          { symbol: "%", fn: (a, b) => (a * b) / 100, label: "Percentage" },
        ];
        let activeOp = operators.find(op => inputText.includes(op.symbol)) || operators[0];
        const calculationPart = inputLower.split(/for|of/)[0];
        const parts = calculationPart.split(activeOp.symbol).map(p => p.trim());
        targetRows.forEach(row => {
          const vals = parts.map(p => {
            const colMatch = columns.find(c => c.toLowerCase() === p);
            if (colMatch) return parseFloat(String(cellVal(row[colMatch])).replace(/[^0-9.]/g, "")) || 0;
            return parseFloat(p) || 0;
          });
          if (vals.length >= 2) finalValue += vals.reduce((acc, curr) => activeOp.fn(acc, curr));
          else if (vals.length === 1) finalValue += vals[0];
        });
        setCalculationResult({
          label:   activeOp.label,
          formula: calculationPart.toUpperCase(),
          value:   finalValue.toLocaleString(undefined, { minimumFractionDigits: 0, maximumFractionDigits: 2 }),
          count:   targetRows.length,
        });
      } catch (err) { console.error("Calc error", err); }
    };
    const t = setTimeout(performCalculation, 300);
    return () => clearTimeout(t);
  }, [inputText, rows, columns]);

  // ============================================================
  // TAB CREATION — with creation log + cooldown enforcement
  // ============================================================
  const addNewTab = useCallback(async (templateData = null) => {
  const check = canCreateTable();

  if (!check.allowed) {
    setProGateMessage(check.reason);
    setShowProGate(true);
    return;
  }

  // Fresh Firestore check: user already used 4 lifetime free tables কিনা
  if (!isPremium) {
    const freshCount = await getFreshLifetimeTableCount();

    if (freshCount >= FREE_TABLE_LIMIT) {
      setLifetimeTablesCreated(freshCount);
      setProGateMessage(
        `Free plan allows ${FREE_TABLE_LIMIT} lifetime tables per Gmail. Upgrade to Pro for unlimited tables.`
      );
      setShowProGate(true);
      return;
    }
  }

  const newTab = templateData
    ? makeNewTab(templateData.name, templateData.columns, Math.max(templateData.sampleRows?.length || 0, 5))
    : makeNewTab(`Sheet ${tabs.length + 1}`, ["ID", "Name", "Email", "Status"], 5);

  if (templateData?.sampleRows?.length) {
    newTab.rows = templateData.sampleRows.map((rowArr) => {
      const obj = {};
      templateData.columns.forEach((col, ci) => {
        obj[col] = makeCellObj(rowArr[ci] !== undefined ? String(rowArr[ci]) : "");
      });
      return obj;
    });
  }

  const creationEntry = { id: newTab.id, createdAt: Date.now() };

  setTabs(prev => [...prev, newTab]);
  setTableCreationLog(prev => [...prev, creationEntry]);
  if (!isPremium) {
    setLifetimeTablesCreated(prev => prev + 1);

    if (user?.uid && db) {
      setDoc(doc(db, "users", user.uid), {
        lifetimeTablesCreated: increment(1),
        lifetimeTableLimit: FREE_TABLE_LIMIT,
        lastTableCreatedAt: serverTimestamp(),
      }, { merge: true }).catch(err => {
        console.error("Lifetime table count update failed:", err);
      });
    }
  }

  queueMicrotask(() => {
    setActiveTabId(newTab.id);
    setCurrentTableId(null);
    setSelectedRows(new Set());
    setSearchTerm("");
  });
}, [
  canCreateTable,
  getFreshLifetimeTableCount,
  isPremium,
  tabs.length,
  user?.uid,
]);
  const applyTemplateToCurrentTab = useCallback((template) => {
    const newRows = template.sampleRows.map((rowArr) => {
      const obj = {};
      template.columns.forEach((col, ci) => {
        obj[col] = makeCellObj(rowArr[ci] !== undefined ? String(rowArr[ci]) : "");
      });
      return obj;
    });
    updateActiveTab({
      columns: template.columns,
      rows: newRows,
      title: template.name,
      columnTypes: {},
      columnFormats: {},
    });
    setShowTemplates(false);
    setCurrentTableId(null);
  }, [updateActiveTab]);

  const closeTab = (e, id) => {
    e.stopPropagation();
    if (tabs.length === 1) return;
    setTabs(prev => {
      const newTabs = prev.filter(t => t.id !== id);
      if (activeTabId === id) setActiveTabId(newTabs[0].id);
      return newTabs;
    });
  };

  const switchTab = useCallback((id) => {
    setActiveTabId(id);
    setCurrentTableId(null);
    setSelectedRows(new Set());
    setSearchTerm("");
  }, []);

  // ── Column Resize ──
  const handleColResizeStart = useCallback((e, col) => {
    e.preventDefault();
    e.stopPropagation();
    resizingCol.current = col;
    startX.current = e.clientX;
    startWidth.current = colWidths[col] || 150;
 
    const onMove = (ev) => {
      const diff = ev.clientX - startX.current;
      const newWidth = Math.max(80, startWidth.current + diff);
      setColWidths(prev => ({ ...prev, [resizingCol.current]: newWidth }));
    };
    const onUp = () => {
      resizingCol.current = null;
      document.removeEventListener("mousemove", onMove);
      document.removeEventListener("mouseup", onUp);
    };
    document.addEventListener("mousemove", onMove);
    document.addEventListener("mouseup", onUp);
  }, [colWidths]);
 
  // ── Row Resize ──
  const handleRowResizeStart = useCallback((e, rIdx) => {
    e.preventDefault();
    e.stopPropagation();
    resizingRow.current = rIdx;
    startY.current = e.clientY;
    startHeight.current = rowHeights[rIdx] || 44;
 
    const onMove = (ev) => {
      const diff = ev.clientY - startY.current;
      const newHeight = Math.max(32, startHeight.current + diff);
      setRowHeights(prev => ({ ...prev, [resizingRow.current]: newHeight }));
    };
    const onUp = () => {
      resizingRow.current = null;
      document.removeEventListener("mousemove", onMove);
      document.removeEventListener("mouseup", onUp);
    };
    document.addEventListener("mousemove", onMove);
    document.addEventListener("mouseup", onUp);
  }, [rowHeights]);
 
  const isCellInSelection = useCallback((rIdx, cIdx) => {
    if (!selectionStart || !selectionEnd) return false;
    const minR = Math.min(selectionStart.rIdx, selectionEnd.rIdx);
    const maxR = Math.max(selectionStart.rIdx, selectionEnd.rIdx);
    const minC = Math.min(selectionStart.cIdx, selectionEnd.cIdx);
    const maxC = Math.max(selectionStart.cIdx, selectionEnd.cIdx);
    return rIdx >= minR && rIdx <= maxR && cIdx >= minC && cIdx <= maxC;
  }, [selectionStart, selectionEnd]);

  // Undo snapshot নাও
  const takeSnapshot = useCallback(() => {
    setUndoStack(prev => {
      const snapshot = JSON.parse(JSON.stringify(rows));
      const newStack = [...prev, snapshot];
      return newStack.slice(-30); // max 30 undo steps
    });
    setRedoStack([]);
  }, [rows]);
 
  // Undo
  const handleUndo = useCallback(() => {
    if (undoStack.length === 0) return;
    const snapshot = undoStack[undoStack.length - 1];
    setRedoStack(prev => [...prev, JSON.parse(JSON.stringify(rows))]);
    setUndoStack(prev => prev.slice(0, -1));
    setRows(snapshot);
  }, [undoStack, rows, setRows]);
 
  // Redo
  const handleRedo = useCallback(() => {
    if (redoStack.length === 0) return;
    const snapshot = redoStack[redoStack.length - 1];
    setUndoStack(prev => [...prev, JSON.parse(JSON.stringify(rows))]);
    setRedoStack(prev => prev.slice(0, -1));
    setRows(snapshot);
  }, [redoStack, rows, setRows]);
 
  // Row Height Auto
  const autoFitRowHeight = useCallback((rIdx) => {
    const row = rows[rIdx];
    if (!row) return;
    let maxLines = 1;
    columns.forEach(col => {
      const val = String(row[col]?.value || row[col] || "");
      const lines = val.split("\n").length;
      const charLines = Math.ceil(val.length / 30);
      maxLines = Math.max(maxLines, lines, charLines);
    });
    const newHeight = Math.max(44, Math.min(maxLines * 28, 200));
    setRowHeights(prev => ({ ...prev, [rIdx]: newHeight }));
  }, [rows, columns]);
 
  const autoFitAllRows = useCallback(() => {
    const newHeights = {};
    rows.forEach((row, rIdx) => {
      let maxLines = 1;
      columns.forEach(col => {
        const val = String(row[col]?.value || row[col] || "");
        const lines = val.split("\n").length;
        const charLines = Math.ceil(val.length / 30);
        maxLines = Math.max(maxLines, lines, charLines);
      });
      newHeights[rIdx] = Math.max(44, Math.min(maxLines * 28, 200));
    });
    setRowHeights(newHeights);
  }, [rows, columns]);
 
  // Find & Replace
  const handleFind = useCallback(() => {
    if (!findText) { setFindResults([]); return; }
    const results = [];
    rows.forEach((row, rIdx) => {
      columns.forEach(col => {
        const val = String(row[col]?.value || row[col] || "");
        if (val.toLowerCase().includes(findText.toLowerCase())) {
          results.push({ rIdx, col });
        }
      });
    });
    setFindResults(results);
    setFindIndex(0);
    if (results.length > 0) setSelectedCell(results[0]);
  }, [findText, rows, columns]);
 
  const handleReplaceOne = useCallback(() => {
    if (findResults.length === 0) return;
    const { rIdx, col } = findResults[findIndex];
    takeSnapshot();
    const val = String(rows[rIdx]?.[col]?.value || rows[rIdx]?.[col] || "");
    const newVal = val.replace(new RegExp(findText, "gi"), replaceText);
    handleCellChange(rIdx, col, newVal);
    const next = (findIndex + 1) % findResults.length;
    setFindIndex(next);
  }, [findResults, findIndex, findText, replaceText, rows, takeSnapshot, handleCellChange]);
 
  const handleReplaceAll = useCallback(() => {
    if (!findText) return;
    takeSnapshot();
    setTabs(prev => prev.map(t => {
      if (t.id !== activeTabId) return t;
      const updatedRows = t.rows.map(row => {
        const newRow = { ...row };
        columns.forEach(col => {
          const val = String(row[col]?.value || row[col] || "");
          if (val.toLowerCase().includes(findText.toLowerCase())) {
            const newVal = val.replace(new RegExp(findText, "gi"), replaceText);
            const existing = row[col];
            const existingObj = typeof existing === "object" ? existing : { value: val };
            newRow[col] = { ...existingObj, value: newVal };
          }
        });
        return newRow;
      });
      return { ...t, rows: updatedRows };
    }));
    handleFind();
  }, [findText, replaceText, columns, activeTabId, takeSnapshot, handleFind]);
 
  // Multi-cell format apply করো
  const applyMultiCellFormat = useCallback((fmt) => {
    if (!selectionStart || !selectionEnd) {
      // single cell
      if (selectedCell) {
        const { rIdx, col } = selectedCell;
        setTabs(prev => prev.map(t => {
          if (t.id !== activeTabId) return t;
          const updatedRows = [...t.rows];
          const existing = updatedRows[rIdx]?.[col] || {};
          const existingObj = typeof existing === "object" ? existing : { value: String(existing) };
          updatedRows[rIdx] = {
            ...updatedRows[rIdx],
            [col]: { ...existingObj, ...fmt },
          };
          return { ...t, rows: updatedRows };
        }));
      }
      return;
    }
    const minR = Math.min(selectionStart.rIdx, selectionEnd.rIdx);
    const maxR = Math.max(selectionStart.rIdx, selectionEnd.rIdx);
    const minC = Math.min(selectionStart.cIdx, selectionEnd.cIdx);
    const maxC = Math.max(selectionStart.cIdx, selectionEnd.cIdx);
 
    setTabs(prev => prev.map(t => {
      if (t.id !== activeTabId) return t;
      const updatedRows = [...t.rows];
      for (let r = minR; r <= maxR; r++) {
        for (let c = minC; c <= maxC; c++) {
          const col = columns[c];
          if (!col || col === "ID") continue;
          const existing = updatedRows[r]?.[col] || {};
          const existingObj = typeof existing === "object" ? existing : { value: String(existing) };
          updatedRows[r] = {
            ...updatedRows[r],
            [col]: { ...existingObj, ...fmt },
          };
        }
      }
      return { ...t, rows: updatedRows };
    }));
  }, [selectionStart, selectionEnd, selectedCell, activeTabId, columns]);
  const getConditionalStyle = useCallback((col, val) => {
    for (const rule of conditionalRules) {
      if (rule.col !== col && rule.col !== "all") continue;
      const num = parseFloat(val);
      let match = false;
      if (rule.type === "gt"       && !isNaN(num) && num >  parseFloat(rule.value)) match = true;
      if (rule.type === "lt"       && !isNaN(num) && num <  parseFloat(rule.value)) match = true;
      if (rule.type === "eq"       && val === rule.value)                            match = true;
      if (rule.type === "contains" && val.toLowerCase().includes(rule.value.toLowerCase())) match = true;
      if (rule.type === "empty"    && val === "")                                   match = true;
      if (match) return { color: rule.color || "", backgroundColor: rule.bgColor || "" };
    }
    return {};
  }, [conditionalRules]);
  // ── Copy Cell ──
  const applyToolbarFormat = useCallback((fmt) => {
  if (!selectedCell) return;
  const { rIdx, col } = selectedCell;
  setTabs(prev => prev.map(t => {
    if (t.id !== activeTabId) return t;
    const updatedRows = [...t.rows];
    const existing = updatedRows[rIdx]?.[col] || {};
    const existingObj = typeof existing === "object" ? existing : { value: String(existing) };
    updatedRows[rIdx] = {
      ...updatedRows[rIdx],
      [col]: {
        ...existingObj,
        ...(fmt.color !== undefined && { color: fmt.color }),
        ...(fmt.bgColor !== undefined && { bgColor: fmt.bgColor }),
        ...(fmt.fontSize !== undefined && { fontSize: fmt.fontSize }),
        ...(fmt.fontWeight !== undefined && { fontWeight: fmt.fontWeight }),
        ...(fmt.fontStyle !== undefined && { fontStyle: fmt.fontStyle }),
        ...(fmt.textAlign !== undefined && { textAlign: fmt.textAlign }),
        ...(fmt.textDecoration !== undefined && { textDecoration: fmt.textDecoration }),
      },
    };
    return { ...t, rows: updatedRows };
  }));
}, [selectedCell, activeTabId]);

const getSelectedCellFormat = useCallback(() => {
  if (!selectedCell) return {};
  const { rIdx, col } = selectedCell;
  const cell = rows[rIdx]?.[col];
  if (!cell || typeof cell !== "object") return {};
  return {
    color:          cell.color          || "",
    bgColor:        cell.bgColor        || "",
    fontSize:       cell.fontSize       || 14,
    fontWeight:     cell.fontWeight     || "normal",
    fontStyle:      cell.fontStyle      || "normal",
    textAlign:      cell.textAlign      || "left",
    textDecoration: cell.textDecoration || "none",
  };
}, [selectedCell, rows]);
  const handleCopyCell = useCallback((rIdx, col) => {
    const val = cellVal(rows[rIdx]?.[col]);
    setCopiedCell({ value: val, fromRow: rIdx, fromCol: col });
    navigator.clipboard.writeText(val).catch(() => {});
  }, [rows]);
 
  // ── Paste Cell ──
  const handlePasteCell = useCallback((rIdx, col) => {
    if (!copiedCell) return;
    handleCellChange(rIdx, col, copiedCell.value);
  }, [copiedCell, handleCellChange]);
 
  // ── Keyboard Copy/Paste Global ──
  useEffect(() => {
    const handleMouseUp = () => setIsSelecting(false);
    document.addEventListener("mouseup", handleMouseUp);
 
    const handleKeyboard = (e) => {
      if ((e.ctrlKey || e.metaKey) && e.key === "z" && !e.shiftKey) {
        e.preventDefault();
        handleUndo();
      }
      if ((e.ctrlKey || e.metaKey) && (e.key === "y" || (e.key === "z" && e.shiftKey))) {
        e.preventDefault();
        handleRedo();
      }
      if ((e.ctrlKey || e.metaKey) && e.key === "h") {
        e.preventDefault();
        setShowFindReplace(true);
      }
    };
    document.addEventListener("keydown", handleKeyboard);
    return () => {
      document.removeEventListener("keydown", handleKeyboard);
      document.removeEventListener("mouseup", handleMouseUp);
    };
  }, [handleUndo, handleRedo]);

  // COLUMN OPERATIONS
  // ============================================================
  const addColumn = () => {
    const newColName = prompt("Enter New Column Name:");
    if (newColName && !columns.includes(newColName)) {
      setColumns([...columns, newColName]);
      setRows(rows.map(row => ({ ...row, [newColName]: makeCellObj("") })));
    }
  };

  const handleColumnHeaderClick = (e, col) => {
    if (col === "ID") return;
    e.preventDefault(); e.stopPropagation();
    const rect = e.currentTarget.getBoundingClientRect();
    setMenuConfig({ show: true, x: rect.left > window.innerWidth - 220 ? rect.left - 200 : rect.left, y: rect.bottom + 4, col });
  };

  const renameColumn = (e) => {
    e.stopPropagation();
    const oldName = menuConfig.col;
    const newName = prompt("Enter new name:", oldName);
    if (newName && newName !== oldName) {
      setColumns(columns.map(c => c === oldName ? newName : c));
      setRows(rows.map(r => { const nr = { ...r, [newName]: r[oldName] || makeCellObj("") }; delete nr[oldName]; return nr; }));
      setColumnTypes(prev => { const ct = { ...prev }; if (ct[oldName]) { ct[newName] = ct[oldName]; delete ct[oldName]; } return ct; });
      updateActiveTab(t => {
        const cf = { ...(t.columnFormats || {}) };
        if (cf[oldName]) { cf[newName] = cf[oldName]; delete cf[oldName]; }
        return { ...t, columnFormats: cf };
      });
    }
    setMenuConfig(p => ({ ...p, show: false }));
  };

  const moveColumn = (e, direction) => {
    e.stopPropagation();
    const idx = columns.indexOf(menuConfig.col);
    const newCols = [...columns];
    const targetIdx = direction === "left" ? idx - 1 : idx + 1;
    if (targetIdx > 0 && targetIdx < columns.length) {
      [newCols[idx], newCols[targetIdx]] = [newCols[targetIdx], newCols[idx]];
      setColumns(newCols);
    }
    setMenuConfig(p => ({ ...p, show: false }));
  };

  const deleteColumn = (e, colName) => {
    e.stopPropagation();
    if (colName === "ID") return;
    if (window.confirm(`Delete column "${colName}"?`)) {
      setColumns(columns.filter(c => c !== colName));
      setRows(rows.map(row => { const nr = { ...row }; delete nr[colName]; return nr; }));
    }
    setMenuConfig(p => ({ ...p, show: false }));
  };

  const clearColumn = (colName) => {
    if (window.confirm(`Clear all values in "${colName}"?`))
      setRows(rows.map(r => ({ ...r, [colName]: { ...(r[colName] || {}), value: "" } })));
    setMenuConfig(p => ({ ...p, show: false }));
  };

  // ============================================================
  // ROW OPERATIONS
  // ============================================================
  const deleteRow = (idx) => {
    if (window.confirm("Delete row?"))
      setRows(rows.filter((_, i) => i !== idx).map((row, i) => ({ ...row, ID: { ...row.ID, value: String(i + 1) } })));
  };

  const deleteSelectedRows = () => {
    if (!selectedRows.size) return;
    if (window.confirm(`Delete ${selectedRows.size} selected rows?`)) {
      setRows(rows.filter((_, i) => !selectedRows.has(i)).map((row, i) => ({ ...row, ID: { ...(row.ID || {}), value: String(i + 1) } })));
      setSelectedRows(new Set());
    }
  };

  const addRow = () => {
    const newRow = makeBlankRow(columns, rows.length + 1);
    setRows([...rows, newRow]);
  };

  const cloneRow = (row) => JSON.parse(JSON.stringify(row || {}));

const cloneCell = (cell) => {
  if (cell && typeof cell === "object") {
    return JSON.parse(JSON.stringify(cell));
  }
  return makeCellObj(cell ?? "");
};

const normalizeRowIds = (list) => {
  return list.map((row, i) => ({
    ...row,
    ID: {
      ...(row.ID && typeof row.ID === "object" ? row.ID : makeCellObj(row.ID)),
      value: String(i + 1),
    },
  }));
};

const autoNumberRows = () => {
  setRows(normalizeRowIds(rows));
};

const cleanData = () => {
  const cleaned = rows.map(row => {
    const next = {};

    columns.forEach(col => {
      const cell = row[col];
      const obj = cell && typeof cell === "object" ? cell : makeCellObj(cell);

      next[col] = {
        ...obj,
        value: col === "ID"
          ? String(obj.value ?? "")
          : String(obj.value ?? "").replace(/\s+/g, " ").trim(),
      };
    });

    return next;
  });

  setRows(normalizeRowIds(cleaned));
};

const removeEmptyRows = () => {
  const nonEmptyRows = rows.filter(row =>
    columns.some(col => col !== "ID" && cellVal(row[col]).trim() !== "")
  );

  setRows(normalizeRowIds(
    nonEmptyRows.length ? nonEmptyRows : [makeBlankRow(columns, 1)]
  ));

  setSelectedRows(new Set());
};

const duplicateSelectedRows = () => {
  if (!selectedRows.size) return;

  const selectedIndexes = [...selectedRows].sort((a, b) => a - b);
  const duplicates = selectedIndexes
    .map(idx => rows[idx])
    .filter(Boolean)
    .map(row => cloneRow(row));

  setRows(normalizeRowIds([...rows, ...duplicates]));
  setSelectedRows(new Set());
};

const fillDownSelectedRows = () => {
  if (!selectedRows.size) return;

  const selectedIndexes = [...selectedRows].sort((a, b) => a - b);
  const nextRows = rows.map(row => cloneRow(row));

  selectedIndexes.forEach(idx => {
    if (idx <= 0 || !nextRows[idx]) return;

    columns.forEach(col => {
      if (col === "ID") return;

      const currentValue = cellVal(nextRows[idx][col]).trim();

      if (currentValue === "") {
        nextRows[idx][col] = cloneCell(nextRows[idx - 1][col]);
      }
    });
  });

  setRows(normalizeRowIds(nextRows));
};
  // ============================================================
  // DRAG & DROP
  // ============================================================
  const handleRowDragStart = (idx) => setDragRowIdx(idx);
  const handleRowDragOver  = (e, idx) => { e.preventDefault(); setDragOverRowIdx(idx); };
  const handleRowDrop = (idx) => {
    if (dragRowIdx === null || dragRowIdx === idx) { setDragRowIdx(null); setDragOverRowIdx(null); return; }
    const newRows = [...rows];
    const [moved] = newRows.splice(dragRowIdx, 1);
    newRows.splice(idx, 0, moved);
    setRows(newRows.map((r, i) => ({ ...r, ID: { ...(r.ID || {}), value: String(i + 1) } })));
    setDragRowIdx(null); setDragOverRowIdx(null);
  };

  const handleColDragStart = (idx) => setDragColIdx(idx);
  const handleColDragOver  = (e, idx) => { e.preventDefault(); setDragOverColIdx(idx); };
  const handleColDrop = (idx) => {
    if (dragColIdx === null || dragColIdx === idx || dragColIdx === 0 || idx === 0) {
      setDragColIdx(null); setDragOverColIdx(null); return;
    }
    const newCols = [...columns];
    const [moved] = newCols.splice(dragColIdx, 1);
    newCols.splice(idx, 0, moved);
    setColumns(newCols);
    setDragColIdx(null); setDragOverColIdx(null);
  };

  // ============================================================
  // FILE IMPORT
  // ============================================================
  const handleFileUpload = async (e) => {
    const file = e.target.files[0];
    if (!file) return;
    setIsLoading(true);
    try {
      const data = await file.arrayBuffer();
      const wb   = XLSX.read(data);
      const ws   = wb.Sheets[wb.SheetNames[0]];
      const json = XLSX.utils.sheet_to_json(ws, { header: 1 });
      if (json.length < 1) return;
      const newCols = json[0].map(String);
      const newRows = json.slice(1).map((r) => {
        const obj = {};
        newCols.forEach((col, ci) => { obj[col] = makeCellObj(r[ci] !== undefined ? String(r[ci]) : ""); });
        return obj;
      });
      updateActiveTab({ columns: newCols, rows: newRows, title: file.name.replace(/\.[^.]+$/, ""), columnTypes: {}, columnFormats: {} });
    } catch (err) { alert("Failed to import file: " + err.message); }
    finally { setIsLoading(false); e.target.value = ""; }
  };

  // ============================================================
  // AI TABLE GENERATOR
  // ============================================================
  const handleGenerateTable = () => {
    if (!genText) return;
    if (!isPremium && history.length >= FREE_TABLE_LIMIT) { setShowUpgrade(true); return; }
    setIsLoading(true);
    setTimeout(() => {
      const input = genText.toLowerCase();
      let newCols, finalTitle, rowCount = 10;
      const rowMatch = input.match(/(\d+)\s*row/);
      if (rowMatch) rowCount = parseInt(rowMatch[1]);
      if (input.includes("salary") || input.includes("payroll")) {
        newCols = ["ID", "Name", "Position", "Basic Salary", "Bonus", "Total"]; finalTitle = "Salary Sheet";
      } else if (input.includes("student") || input.includes("school")) {
        newCols = ["ID", "Name", "Roll", "Class", "Subject", "Mark", "Grade"]; finalTitle = "Student List";
      } else { newCols = ["ID", "Name", "Email", "Status"]; finalTitle = genText.toUpperCase(); }
      const newRows = Array.from({ length: rowCount }, (_, i) => makeBlankRow(newCols, i + 1));
      updateActiveTab({ columns: newCols, rows: newRows, title: finalTitle, columnTypes: {}, columnFormats: {} });
      setIsLoading(false); setGenText(""); setCurrentTableId(null);
    }, 1200);
  };

  // ============================================================
  const buildRowsFromAI = (columns, aiRows) => {
  const safeRows = Array.isArray(aiRows) ? aiRows : [];

  return safeRows.map((row, rowIndex) => {
    const obj = {};

    columns.forEach((col, colIndex) => {
      let value = "";

      if (Array.isArray(row)) {
        value = row[colIndex] ?? "";
      } else if (row && typeof row === "object") {
        value = row[col] ?? "";
      }

      obj[col] = makeCellObj(col === "ID" ? rowIndex + 1 : value);
    });

    return obj;
  });
};

const handleAITableGenerate = async (prompt) => {
  if (!prompt.trim()) return;

  if (!isPremium && aiTableUsage >= AI_FREE_TABLE_LIMIT) {
    setProGateMessage(
      `Free plan allows ${AI_FREE_TABLE_LIMIT} AI-generated tables. Upgrade to Pro for unlimited AI table generation.`
    );
    setShowAITable(false);
    setShowProGate(true);
    return;
  }

  const check = canCreateTable();

  if (!check.allowed) {
    setProGateMessage(check.reason);
    setShowAITable(false);
    setShowProGate(true);
    return;
  }

  setAiBuildLoading(true);

  try {
    const idToken = user ? await user.getIdToken() : null;
    const res = await fetch(AI_TABLE_ENDPOINT, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        Authorization: `Bearer ${idToken}`,
      },
      body: JSON.stringify({
        prompt,
        uid: user?.uid || null,
        email: user?.email || null,
        isPremium,
      }),
    });

    const data = await res.json().catch(() => ({}));

    if (!res.ok) {
      throw new Error(data.error || data.code || "AI table request failed");
    }

    const aiColumns = Array.isArray(data.columns) && data.columns.length
      ? data.columns.map(String)
      : ["ID", "Name", "Email", "Status"];

    const columnsWithId = aiColumns.includes("ID")
      ? aiColumns
      : ["ID", ...aiColumns];

    const aiRows = buildRowsFromAI(columnsWithId, data.rows || []);

    const finalRows = aiRows.length
      ? aiRows
      : Array.from({ length: 8 }, (_, i) => makeBlankRow(columnsWithId, i + 1));

    const newTab = {
      id: Date.now() + Math.random(),
      title: data.title || "AI Generated Table",
      columns: columnsWithId,
      rows: finalRows,
      columnTypes: data.columnTypes || {},
      columnFormats: {},
      tabColor: "#3978d8",
    };

    const creationEntry = { id: newTab.id, createdAt: Date.now() };

    setTabs(prev => [...prev, newTab]);
    setTableCreationLog(prev => [...prev, creationEntry]);
  if (!isPremium) {
    setLifetimeTablesCreated(prev => prev + 1);

    if (user?.uid && db) {
      setDoc(doc(db, "users", user.uid), {
        lifetimeTablesCreated: increment(1),
        lifetimeTableLimit: FREE_TABLE_LIMIT,
        lastTableCreatedAt: serverTimestamp(),
      }, { merge: true }).catch(err => {
        console.error("Lifetime table count update failed:", err);
      });
    }
  }
    setAiTableUsage(prev => prev + 1);
    setActiveTabId(newTab.id);
    setCurrentTableId(null);
    setSelectedRows(new Set());
    setSearchTerm("");
    setShowAITable(false);
  } catch (err) {
    console.error("AI table error:", err);
    alert("AI table বানাতে সমস্যা হয়েছে। API endpoint/backend ঠিক আছে কিনা check করো.");
  } finally {
    setAiBuildLoading(false);
  }
};

// ============================================================

  /**
   * Whether the current active sheet already has a saved history entry.
   * We match by currentTableId.
   */
  const isCurrentTableSaved = currentTableId !== null && history.some(h => h.id === currentTableId);

  const saveToHistory = () => {
  const isNewSave = !isCurrentTableSaved;

  if (!isPremium && isNewSave && history.length >= FREE_TABLE_LIMIT) {
    setShowUpgrade(true);
    return;
  }

  const snapshot = {
    id: isCurrentTableSaved ? currentTableId : Date.now(),
    title: tableTitle || "Untitled",
    date: new Date().toLocaleString(),
    columns: [...columns],
    rows: JSON.parse(JSON.stringify(rows.map(row => migrateRow(row, columns)))),
    columnTypes: { ...columnTypes },
    columnFormats: { ...columnFormats },
    tabColor: activeTab?.tabColor || "",
  };

  if (isCurrentTableSaved) {
    setHistory(prev => prev.map(item =>
      item.id === currentTableId ? snapshot : item
    ));
    alert("Table Updated! ✅");
  } else {
    setHistory(prev => [snapshot, ...prev]);
    setCurrentTableId(snapshot.id);
    alert("Table Saved! ✅");
  }
};
const handleHistoryClick = (item) => {
  if (currentTableId === item.id) {
    setCurrentTableId(null);
    return;
  }

  const migratedRows = (item.rows || []).map(row =>
    migrateRow(row, item.columns || [])
  );

  updateActiveTab({
    columns: item.columns || [],
    rows: migratedRows,
    title: item.title || "Untitled",
    columnTypes: item.columnTypes || {},
    columnFormats: item.columnFormats || {},
    tabColor: item.tabColor || "",
  });

  setCurrentTableId(item.id);
};

const deleteHistoryItem = (e, historyId) => {
  e.stopPropagation();

  if (window.confirm("Delete this table from history?")) {
    setHistory(prev => prev.filter(item => item.id !== historyId));

    if (currentTableId === historyId) {
      setCurrentTableId(null);
    }
  }
};

const clearAllHistory = () => {
  if (window.confirm("Clear all history?")) {
    setHistory([]);
    setCurrentTableId(null);
  }
};
  // ============================================================
  // KEYBOARD NAVIGATION
  // ============================================================
  const handleKeyDown = (e, rIdx, cIdx) => {
    const input          = e.target;
    const cursorPosition = input.selectionStart;
    const textLength     = input.value.length;
    let targetRow = rIdx, targetCol = cIdx;
    if (e.key === "ArrowRight")                       { if (cursorPosition < textLength) return; targetCol++; }
    else if (e.key === "ArrowLeft")                   { if (cursorPosition > 0) return; targetCol--; }
    else if (e.key === "ArrowDown" || e.key === "Enter") { e.preventDefault(); targetRow++; }
    else if (e.key === "ArrowUp")                     { targetRow--; }
    else return;
    if (targetRow >= 0 && targetRow < rows.length && targetCol >= 0 && targetCol < columns.length) {
      const inputs = tableRef.current?.querySelectorAll('input[type="text"], input[type="date"]');
      const next   = targetRow * columns.length + targetCol;
      if (inputs?.[next]) { inputs[next].focus(); if (e.key !== "ArrowLeft") inputs[next].select(); }
    }
  };

  // ============================================================
  // ANALYTICS
  // ============================================================
  const getAnalyticsData = () => {
    const numericCols = columns.filter(col =>
      col !== "ID" && rows.some(r => !isNaN(parseFloat(cellVal(r[col]))) && cellVal(r[col]) !== "")
    );
    if (!numericCols.length) return null;
    const col    = numericCols[0];
    const vals   = rows.map(r => parseFloat(cellVal(r[col])) || 0);
    const labels = rows.map(r => cellVal(r["Name"] || r["ID"] || ""));
    return { col, vals, labels };
  };

  // ============================================================
  // CELL RENDERING
  // ============================================================
  const getColMinWidth = (col) => {
    const colFmt   = columnFormats[col];
    const headerLen = col.length;
    const maxCellLen = rows.reduce((max, row) => {
      const v = cellVal(row[col]);
      const cell = row[col];
      const effectiveFontSize = (cell && typeof cell === "object" ? cell.fontSize : null) || colFmt?.fontSize || 14;
      const len = v.length * (effectiveFontSize / 14);
      return Math.max(max, len);
    }, 0);
    const maxLen = Math.max(headerLen * 1.1, maxCellLen, 8);
    return `${Math.min(Math.max(maxLen * 9 + 60, 110), 600)}px`;
  };

  const getConditionalTextColor = (col, val) => {
    const num = parseFloat(val);
    if (isNaN(num) || val === "" || col === "ID") return null;
    if (/salary|income|wage|pay|earning/i.test(col) && num > 50000) return "#10b981";
    if (num < 33) return "#ef4444";
    return null;
  };

  const buildCellStyle = (col, cell) => {
    const colFmt    = columnFormats[col];
    const cellObj   = cell && typeof cell === "object" ? cell : null;
    const val       = cellObj ? cellObj.value : String(cell ?? "");
    const conditionalColor = getConditionalTextColor(col, val);
    const condFmt   = getConditionalStyle(col, val);
    const s = {};
    const effectiveColor      = (cellObj?.color)      || (colFmt?.color)      || conditionalColor || condFmt.color;
    const effectiveBg         = (cellObj?.bgColor)    || (colFmt?.bgColor)    || condFmt.backgroundColor;
    const effectiveFontSize   = (cellObj?.fontSize)   || (colFmt?.fontSize);
    const effectiveFontWeight = (cellObj?.fontWeight) || (colFmt?.fontWeight);
    const effectiveFontStyle  = (cellObj?.fontStyle);
    const effectiveTextAlign  = (cellObj?.textAlign);
    const effectiveTextDec    = (cellObj?.textDecoration);
    // Border
    const borderTop    = cellObj?.borderTop;
    const borderBottom = cellObj?.borderBottom;
    const borderLeft   = cellObj?.borderLeft;
    const borderRight  = cellObj?.borderRight;
 
    if (effectiveColor)      s.color           = effectiveColor;
    if (effectiveBg)         s.backgroundColor = effectiveBg;
    if (effectiveFontSize)   s.fontSize        = `${effectiveFontSize}px`;
    if (effectiveFontWeight) s.fontWeight      = effectiveFontWeight;
    if (effectiveFontStyle)  s.fontStyle       = effectiveFontStyle;
    if (effectiveTextAlign)  s.textAlign       = effectiveTextAlign;
    if (effectiveTextDec)    s.textDecoration  = effectiveTextDec;
    if (borderTop)           s.borderTop       = borderTop;
    if (borderBottom)        s.borderBottom    = borderBottom;
    if (borderLeft)          s.borderLeft      = borderLeft;
    if (borderRight)         s.borderRight     = borderRight;
 
    const colType = columnTypes[col];
    if (colType === "email" && val && !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(val)) s.color = "#ef4444";
    return s;
  };

  const renderCell = (row, rIdx, col, cIdx) => {
    const colType  = columnTypes[col];
    const cell     = row[col];
    const rawVal   = cellVal(cell);
    const val      = rawVal.startsWith("=") ? evaluateFormula(rawVal, rows, columns) : rawVal;
    const styleObj = buildCellStyle(col, cell);
    const baseCls = `bg-transparent outline-none text-sm transition-all focus:bg-sky-500/[0.08] px-4 py-3.5 ${isDark ? "text-[#f1f6ff]" : "text-[#172033]"}`;
    const inputStyle = { ...styleObj, width: "100%", minWidth: "max-content", boxSizing: "border-box" };

    if (col === "ID") return (
      <input type="text" value={val} readOnly
        className="bg-transparent px-4 py-3.5 outline-none text-sm text-[#3978d8] font-black text-center w-full"
        style={{ minWidth: "50px" }} />
    );

    if (colType === "checkbox") return (
      <div className="flex items-center justify-center py-3">
        <input type="checkbox" checked={val === "true" || val === true}
          onChange={e => handleCellChange(rIdx, col, String(e.target.checked))}
          className="w-4 h-4 accent-indigo-500 cursor-pointer" />
      </div>
    );

    if (colType === "date") return (
      <input type="date" value={val}
        onKeyDown={e => handleKeyDown(e, rIdx, cIdx)}
        onChange={e => handleCellChange(rIdx, col, e.target.value)}
        className={`${baseCls} cursor-pointer w-full`}
        style={{ ...inputStyle, colorScheme: isDark ? "dark" : "light" }} />
    );

    if (colType === "dropdown") {
      const opts = columnTypes[col + "__options"] || [];
      return (
        <select value={val} onChange={e => handleCellChange(rIdx, col, e.target.value)}
          className="bg-transparent px-4 py-3.5 outline-none text-sm cursor-pointer appearance-none w-full"
          style={styleObj}>
          <option value="">— Select —</option>
          {opts.map(o => <option key={o} value={o} style={{ background: isDark ? "#1e293b" : "#fff" }}>{o}</option>)}
        </select>
      );
    }

    return (
      <input type="text" value={val}
        onFocus={e => {
          setSelectedCell({ rIdx, col });
          setFormulaBarVal(rawVal);
          if (rawVal.startsWith("=")) e.target.value = rawVal;
        }}
        onChange={e => {
          handleCellChange(rIdx, col, e.target.value);
          setFormulaBarVal(e.target.value);
        }}
        onClick={e => e.stopPropagation()}
        onKeyDown={e => handleKeyDown(e, rIdx, cIdx)}
        className={baseCls}
        style={inputStyle} />
    );
  };

  const getRowStyle = (row) => {
    const statusCols = ["Status", "status"];
    for (const col of statusCols) {
      if (row[col] && ["checked", "done", "complete", "completed"].includes(cellVal(row[col]).toLowerCase()))
        return isDark ? "bg-emerald-500/10" : "bg-emerald-50";
    }
    return "";
  };

  // ============================================================
  // PDF EXPORT
  // ============================================================
  const downloadPDF = async () => {
    try {
      const docPDF = new jsPDF({ orientation: columns.length > 6 ? "landscape" : "portrait" });
      docPDF.setFillColor(255, 255, 255);
      docPDF.rect(0, 0, docPDF.internal.pageSize.width, docPDF.internal.pageSize.height, "F");
      const hasUnicode = rows.some(row => columns.some(col => hasNonAscii(cellVal(row[col]))));
      const getCellPdfStyle = (row, col) => {
        const cell    = row[col];
        const colFmt  = columnFormats[col];
        const cellObj = cell && typeof cell === "object" ? cell : null;
        return {
          bgColor:    (cellObj?.bgColor)    || (colFmt?.bgColor)    || "",
          color:      (cellObj?.color)      || (colFmt?.color)      || "#0f172a",
          fontSize:   (cellObj?.fontSize)   || (colFmt?.fontSize)   || 11,
          fontWeight: (cellObj?.fontWeight) || (colFmt?.fontWeight) || "normal",
        };
      };
      docPDF.setFontSize(18); docPDF.setTextColor(99, 102, 241);
      docPDF.text(tableTitle || "SheetMind", 14, 18);
      docPDF.setFontSize(9);  docPDF.setTextColor(150, 150, 150);
      docPDF.text(`Exported: ${new Date().toLocaleString()} · ${rows.length} rows · ${columns.length} cols`, 14, 26);
      if (hasUnicode) {
        const renderTextToCanvas = (text, fontSize = 12, fontColor = "#0f172a", bgColor = "#ffffff", isBold = false) => {
          const canvas = document.createElement("canvas");
          canvas.width = 300; canvas.height = 40;
          const ctx = canvas.getContext("2d");
          ctx.fillStyle = bgColor || "#ffffff"; ctx.fillRect(0, 0, 300, 40);
          ctx.fillStyle = fontColor || "#0f172a";
          ctx.font = `${isBold ? "bold " : ""}${fontSize}px 'Noto Sans Bengali', Arial, sans-serif`;
          ctx.fillText(text, 4, 28);
          return canvas.toDataURL("image/png");
        };
        const cellW = Math.min((docPDF.internal.pageSize.width - 28) / columns.length, 55);
        const cellH = 10;
        let startY = 34;
        columns.forEach((col, ci) => {
          docPDF.setFillColor(30, 41, 59); docPDF.rect(14 + ci * cellW, startY, cellW, cellH, "F");
          docPDF.setTextColor(148, 163, 184); docPDF.setFontSize(8);
          docPDF.text(col.slice(0, 10), 16 + ci * cellW, startY + 7);
        });
        startY += cellH;
        rows.forEach((row) => {
          if (startY > docPDF.internal.pageSize.height - 20) {
            docPDF.addPage(); docPDF.setFillColor(255, 255, 255);
            docPDF.rect(0, 0, docPDF.internal.pageSize.width, docPDF.internal.pageSize.height, "F");
            startY = 14;
          }
          columns.forEach((col, ci) => {
            const cellText = cellVal(row[col]);
            const fmt      = getCellPdfStyle(row, col);
            const cellBg   = fmt.bgColor || (rows.indexOf(row) % 2 === 0 ? "#f8fafc" : "#ffffff");
            const [r, g, b] = hexToRgb(cellBg);
            docPDF.setFillColor(r, g, b); docPDF.rect(14 + ci * cellW, startY, cellW, cellH, "F");
            docPDF.setDrawColor(226, 232, 240); docPDF.rect(14 + ci * cellW, startY, cellW, cellH, "S");
            if (hasNonAscii(cellText)) {
              const imgData = renderTextToCanvas(cellText, fmt.fontSize, fmt.color, cellBg, fmt.fontWeight === "bold");
              docPDF.addImage(imgData, "PNG", 14 + ci * cellW, startY, cellW, cellH);
            } else {
              const [cr, cg, cb] = hexToRgb(fmt.color || "#0f172a");
              docPDF.setTextColor(cr, cg, cb); docPDF.setFontSize(Math.min(fmt.fontSize, 10));
              docPDF.setFont("helvetica", fmt.fontWeight === "bold" ? "bold" : "normal");
              docPDF.text(cellText.slice(0, 12), 16 + ci * cellW, startY + 7);
            }
          });
          startY += cellH;
        });
      } else {
        const head = [columns];
        const body = rows.map(row => columns.map(col => cellVal(row[col])));
        const didParseCell = (data) => {
          if (data.section === "body") {
            const { row, column } = data;
            const col      = columns[column.index];
            const tableRow = rows[row.index];
            if (!tableRow) return;
            const fmt = getCellPdfStyle(tableRow, col);
            if (fmt.bgColor)              { const [r, g, b] = hexToRgb(fmt.bgColor); data.cell.styles.fillColor = [r, g, b]; }
            if (fmt.color)                { const [r, g, b] = hexToRgb(fmt.color);   data.cell.styles.textColor = [r, g, b]; }
            if (fmt.fontSize)             data.cell.styles.fontSize  = Math.min(fmt.fontSize, 10);
            if (fmt.fontWeight === "bold") data.cell.styles.fontStyle = "bold";
          }
        };
        autoTable(docPDF, {
          head, body, startY: 32,
          styles:            { fontSize: 9, cellPadding: 5, font: "helvetica", lineColor: [226, 232, 240], lineWidth: 0.3 },
          headStyles:        { fillColor: [30, 41, 59], textColor: [148, 163, 184], fontStyle: "bold", fontSize: 9 },
          alternateRowStyles: { fillColor: [248, 250, 252] },
          bodyStyles:        { fillColor: [255, 255, 255], textColor: [15, 23, 42] },
          theme: "grid", didParseCell,
        });
      }
      docPDF.save(`${(tableTitle || "SHEETMIND").replace(/\s+/g, "_")}.pdf`);
    } catch (err) { console.error("PDF export error:", err); alert("PDF export failed: " + err.message); }
  };

  // ============================================================
  // EXCEL EXPORT
  // ============================================================
  const downloadExcel = () => {
    try {
      const wsData = [columns, ...rows.map(row => columns.map(col => cellVal(row[col])))];
      const ws     = XLSX.utils.aoa_to_sheet(wsData);
      const colWidths = columns.map(col => {
        const maxLen = Math.max(col.length, ...rows.map(r => cellVal(r[col]).length));
        return { wch: Math.min(Math.max(maxLen + 2, 10), 40) };
      });
      ws["!cols"] = colWidths;
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, tableTitle || "Sheet1");
      XLSX.writeFile(wb, `${(tableTitle || "SHEETMIND").replace(/\s+/g, "_")}.xlsx`);
    } catch (err) { alert("Excel export failed: " + err.message); }
  };

  // ============================================================
  // UI STATE
  // ============================================================
  const analyticsData  = getAnalyticsData();
  const filteredRows = useMemo(() => {
  const indexedRows = rows.map((row, originalIndex) => ({ row, originalIndex }));

  if (!searchTerm) return indexedRows;

  return indexedRows.filter(({ row }) =>
    Object.values(row).some(v =>
      cellVal(v).toLowerCase().includes(searchTerm.toLowerCase())
    )
  );
}, [rows, searchTerm]);

  if (page === "landing") return <LandingPage onLogin={handleGoogleLogin} isDark={isDark} />;

  // ============================================================
  // DESIGN TOKENS — DEEP BLUE + GOLD PREMIUM THEME
  // ============================================================
  const bgPage = isDark
  ? "bg-[#07111f]"
  : "bg-[#eef4ff]";


const bgPanel = isDark
  ? "bg-[#0f2038] border-[#29456b] shadow-[0_18px_55px_rgba(0,0,0,0.28)]"
  : "bg-white border-[#c8d7ec] shadow-[0_18px_55px_rgba(43,76,126,0.12)]";

const bgSubtle = isDark
  ? "bg-[#132844]"
  : "bg-[#f5f8fd]";

const textMain = isDark
  ? "text-[#f1f6ff]"
  : "text-[#172033]";

const textSub = isDark
  ? "text-[#9fb3d1]"
  : "text-[#5f6f89]";

const borderColor = isDark
  ? "border-[#29456b]"
  : "border-[#c8d7ec]";


const inputCls = isDark
  ? "bg-[#0a1729] border-[#315174] text-[#f1f6ff] placeholder:text-[#7186a6] focus:border-[#6fb6ff] focus:ring-2 focus:ring-[#6fb6ff]/20"
  : "bg-white border-[#c8d7ec] text-[#172033] placeholder:text-[#7b8aa3] focus:border-[#3978d8] focus:ring-2 focus:ring-[#3978d8]/15 shadow-sm";

const tableHeadCls = isDark
  ? "bg-[#11243d]"
  : "bg-[#e8f0fb]";

const dividerCls = isDark
  ? "divide-[#29456b]"
  : "divide-[#d9e4f3]";

const btnPrimary =
  "bg-gradient-to-r from-[#2563eb] to-[#0ea5e9] hover:from-[#1d4ed8] hover:to-[#0284c7] text-white font-black shadow-[0_10px_24px_rgba(37,99,235,0.28)] transition-all duration-200 rounded-xl border border-white/10";

const btnGhost = isDark
  ? "border border-[#315174] bg-[#10223a] hover:bg-[#173251] text-[#c6d7ef] hover:text-white transition-all duration-200 rounded-xl shadow-sm"
  : "border border-[#c8d7ec] bg-white hover:bg-[#eaf2ff] text-[#36516f] hover:text-[#17417a] transition-all duration-200 rounded-xl shadow-sm";

const btnDanger =
  "bg-gradient-to-r from-[#ef4444] to-[#f97316] hover:from-[#dc2626] hover:to-[#ea580c] border border-white/10 text-white font-black shadow-[0_10px_24px_rgba(239,68,68,0.22)] transition-all duration-200 rounded-xl";

const btnSuccess =
  "bg-gradient-to-r from-[#059669] to-[#14b8a6] hover:from-[#047857] hover:to-[#0d9488] text-white font-black shadow-[0_10px_24px_rgba(20,184,166,0.22)] transition-all duration-200 rounded-xl";
  const syncBadge = () => {
    if (syncStatus === "syncing") return <span className="flex items-center gap-1 text-[8px] text-sky-400 font-black"><div className="w-1.5 h-1.5 border border-sky-400/50 border-t-sky-400 rounded-full animate-spin"></div>Saving...</span>;
    if (syncStatus === "saved" && lastSaved) return <span className="text-[8px] text-emerald-400 font-black">✓ {lastSaved}</span>;
    if (syncStatus === "local") return <span className="text-[8px] text-amber-400 font-black">💾 Local</span>;
    return null;
  };

  // Cooldown status for UI hint
  const freeTablesRemaining = Math.max(FREE_TABLE_LIMIT - lifetimeTablesCreated, 0);
  const showFreeLimitHint = !isPremium;

  return (
    <div
      className={`min-h-screen ${bgPage} ${textMain} p-2 md:p-5 font-sans relative`}
      style={{ fontFamily: "'DM Mono','Fira Mono',monospace", transition: "background 0.3s, color 0.3s" }}
    >
      <link rel="preconnect" href="https://fonts.googleapis.com" />
      <link rel="preconnect" href="https://fonts.gstatic.com" crossOrigin="anonymous" />
      <link href="https://fonts.googleapis.com/css2?family=Noto+Sans+Bengali&display=swap" rel="stylesheet" />

      <input type="file" ref={fileInputRef} onChange={handleFileUpload} className="hidden" accept=".csv,.xlsx,.xls" />

      {/* ======== MODALS ======== */}
      {showAdmin    && <AdminPanel onClose={() => setShowAdmin(false)} isDark={isDark} currentUser={user} />}
      {showUpgrade  && <UpgradeModal onClose={() => setShowUpgrade(false)} isDark={isDark} currentUser={user} />}
      {showAITable && (
        <AITableModal
          isDark={isDark}
          isPremium={isPremium}
          aiTableUsage={aiTableUsage}
          onClose={() => setShowAITable(false)}
          onGenerate={handleAITableGenerate}
        />
      )}
      
      {showWelcome && (
        <WelcomeModal
          isDark={isDark}
          userName={user?.displayName?.split(" ")[0]}
          onClose={() => setShowWelcome(false)}
        />
      )}
      {showProGate && (
        <ProGateModal
          message={proGateMessage}
          tabCount={tabs.length}
          isDark={isDark}
          onClose={() => setShowProGate(false)}
          onUpgrade={() => { setShowProGate(false); setShowUpgrade(true); }}
        />
      )}

      {showAIAnalysis && <AIAnalysisModal columns={columns} rows={rows} isDark={isDark} onClose={() => setShowAIAnalysis(false)} />}

      {showTemplates && (
        <TemplateGallery
          isDark={isDark}
          onClose={() => setShowTemplates(false)}
          onSelect={(template) => {
            if (template.id === "blank") {
              addNewTab(null);
            } else {
              const choice = window.confirm(
                `Apply "${template.name}" template?\n\nOK = Apply to current sheet\nCancel = Open in new sheet`
              );
              if (choice) {
                applyTemplateToCurrentTab(template);
              } else {
                addNewTab(template);
              }
            }
            setShowTemplates(false);
          }}
        />
      )}

      {editingColType && (
        <ColumnTypeModal
          col={editingColType}
          currentType={columnTypes[editingColType]}
          currentOptions={columnTypes[editingColType + "__options"]}
          isDark={isDark}
          onClose={() => setEditingColType(null)}
          onSave={(type, opts) => {
            setColumnTypes(prev => ({ ...prev, [editingColType]: type, [editingColType + "__options"]: opts }));
            setEditingColType(null);
          }}
        />
      )}
      {/* BORDER MODAL */}
      {showBorderModal && selectedCell && (
        <BorderModal
          selectedCell={selectedCell}
          rows={rows}
          activeTabId={activeTabId}
          setTabs={setTabs}
          onClose={() => setShowBorderModal(false)}
          isDark={isDark}
        />
      )}

      {showConditional && (
        <ConditionalModal
          columns={columns}
          conditionalRules={conditionalRules}
          onApply={(rules) => setConditionalRules(rules)}
          onClose={() => setShowConditional(false)}
          isDark={isDark}
        />
      )}
      {/* CONDITIONAL FORMATTING MODAL */}
      {editingFormat && (
        <CellFormatModal
          col={editingFormat.col}
          rowIdx={editingFormat.rowIdx}
          scope={editingFormat.scope}
          currentCell={editingFormat.scope === "cell" ? (rows[editingFormat.rowIdx]?.[editingFormat.col] || {}) : null}
          currentColFormat={editingFormat.scope === "column" ? (columnFormats[editingFormat.col] || {}) : null}
          isDark={isDark}
          isPremium={isPremium}
          onClose={() => setEditingFormat(null)}
          onSave={(fmt) => {
            if (editingFormat.scope === "cell")        applyCellFormat(editingFormat.rowIdx, editingFormat.col, fmt);
            else if (editingFormat.scope === "column") applyColumnFormat(editingFormat.col, fmt);
            else if (editingFormat.scope === "tab") {
          const targetTabId = editingFormat.tabId || activeTabId;
            setTabs(prev => prev.map(t =>
              t.id === targetTabId ? { ...t, tabColor: fmt.bgColor || "" } : t
            ));
          }
            setEditingFormat(null);
          }}
        />
      )}

      {/* COLUMN CONTEXT MENU */}
      {menuConfig.show && (
        <div
          ref={menuRef}
          style={{ top: menuConfig.y, left: menuConfig.x }}
          className={`fixed z-[100] w-56 ${
            isDark ? "bg-[#0d1726] border-[#1e2d42]" : "bg-white border-slate-200"
          } border rounded-2xl shadow-2xl p-1.5 backdrop-blur-xl`}
        >
          {[
            { icon: "✏️", label: "Rename Column",   fn: renameColumn },
            { icon: "🔧", label: "Set Column Type", fn: (e) => { e.stopPropagation(); setEditingColType(menuConfig.col); setMenuConfig(p => ({ ...p, show: false })); } },
            { icon: "🎨", label: "Format Column",   fn: (e) => { e.stopPropagation(); setEditingFormat({ col: menuConfig.col, rowIdx: null, scope: "column" }); setMenuConfig(p => ({ ...p, show: false })); } },
            { icon: "←",  label: "Move Left",       fn: (e) => moveColumn(e, "left") },
            { icon: "→",  label: "Move Right",      fn: (e) => moveColumn(e, "right") },
            { icon: "🧹", label: "Clear Column",    fn: (e) => { e.stopPropagation(); clearColumn(menuConfig.col); } },
            { icon: "👁", label: hiddenCols.has(menuConfig.col) ? "Show Column" : "Hide Column",
              fn: (e) => {
                e.stopPropagation();
                setHiddenCols(prev => {
                  const next = new Set(prev);
                  if (next.has(menuConfig.col)) next.delete(menuConfig.col);
                  else next.add(menuConfig.col);
                  return next;
                });
                setMenuConfig(p => ({ ...p, show: false }));
              }
            },
          ].map((item, i) => (
            <button key={i} onClick={item.fn}
              className={`w-full text-left px-3.5 py-2.5 text-[11px] rounded-xl flex items-center gap-3 font-black uppercase tracking-wide transition-colors ${
                isDark ? "hover:bg-white/[0.06] text-slate-400 hover:text-slate-200" : "hover:bg-slate-100 text-slate-500 hover:text-slate-800"
              }`}>
              {item.icon} {item.label}
            </button>
          ))}
          <div className={`h-px my-1.5 ${isDark ? "bg-[#1e2d42]" : "bg-slate-100"}`} />
          <button onClick={(e) => deleteColumn(e, menuConfig.col)}
            className="w-full text-left px-3.5 py-2.5 text-[11px] rounded-xl flex items-center gap-3 font-black uppercase tracking-wide transition-colors text-red-400 hover:bg-red-500/15 hover:text-red-300">
             Delete Column
          </button>
        </div>
      )}

      {/* ======= MAIN LAYOUT ======= */}
      <div className="max-w-[1440px] mx-auto grid grid-cols-1 lg:grid-cols-4 gap-5">
        <div className="lg:col-span-3 flex flex-col gap-4">

          <BannerAd isDark={isDark} />

          {/* ===== HEADER ===== */}
          <header className={`${bgPanel} border ${borderColor} rounded-[1.75rem] px-6 py-4 shadow-xl flex flex-col md:flex-row justify-between items-center gap-4`}>
            <div className="flex items-center gap-3">
              <div className="w-9 h-9 rounded-xl flex items-center justify-center flex-shrink-0"
                style={{ background: "linear-gradient(135deg, #6366f1, #22d3ee)" }}>
                <span className="text-white text-xs font-black">SM</span>
              </div>
              <div>
                <h1 className="text-xl font-black bg-gradient-to-r from-[#6fb6ff] via-[#22d3ee] to-[#34d399] bg-clip-text text-transparent uppercase tracking-tight leading-none">
                  SheetMind
                </h1>
                <div className="flex items-center gap-2 mt-0.5 flex-wrap">
                  <span className={`text-[8px] font-black tracking-[0.3em] uppercase ${textSub}`}>PRO v15</span>
                  {syncBadge()}
                  {isPremium && (
                    <span className="text-[8px] font-black px-2 py-0.5 rounded-full bg-indigo-500/15 text-indigo-400 uppercase border border-indigo-500/20">
                      ⭐ {isAdmin ? "Admin" : "Premium"}
                    </span>
                  )}
                  {/* 24h Cooldown hint */}
                  {showFreeLimitHint && (
                  <span className="text-[8px] font-black px-2 py-0.5 rounded-full bg-amber-500/15 text-amber-400 uppercase border border-amber-500/20">
                    Free Tables: {freeTablesRemaining}/{FREE_TABLE_LIMIT}
                  </span>
                  )}
                </div>
              </div>
            </div>

            <div className="flex flex-wrap justify-end gap-1 max-w-[620px]">
              <button onClick={() => setIsDark(d => !d)} className={`px-2 py-1.5 rounded-lg font-black text-[9px] uppercase ${btnGhost}`} title="Toggle theme">
                {isDark ? "☀️" : "🌙"}
              </button>

              {user ? (
                <div className={`flex items-center gap-2 px-2 pr-3 py-1.5 rounded-xl border ${borderColor} ${isDark ? "bg-[#0a1018]" : "bg-slate-100"}`}>
                  <img src={user.photoURL} alt="" className="w-6 h-6 rounded-full border border-sky-500/40"
                    onError={e => { e.target.src = "https://ui-avatars.com/api/?name=U&background=6366f1&color=fff"; }} />
                  <span className="text-[10px] font-black uppercase tracking-widest text-sky-400">{user.displayName?.split(" ")[0]}</span>
                  <button onClick={handleLogout} className="text-[9px] text-red-400 hover:text-red-300 font-black uppercase ml-0.5">Exit</button>
                </div>
              ) : (
                <button onClick={handleGoogleLogin} className={`px-2 py-1.5 rounded-lg font-black text-[9px] uppercase ${btnPrimary}`}>Sign In</button>
              )}

              {(isAdmin || user?.email === ADMIN_EMAIL) && (
                <button onClick={() => setShowAdmin(true)}
                  className="px-2 py-1.5 rounded-lg font-black text-[9px] uppercase bg-purple-500/10 hover:bg-purple-500 border border-purple-500/20 hover:border-purple-500 text-purple-400 hover:text-white transition-all duration-150">
                  Admin
                </button>
              )}

              <button onClick={() => setShowTemplates(true)}
                className="px-2 py-1.5 rounded-lg font-black text-[9px] uppercase bg-emerald-500/10 hover:bg-emerald-500 border border-emerald-500/20 hover:border-emerald-500 text-emerald-400 hover:text-white transition-all duration-150"
                title="Template Gallery">
                 Templates
              </button>
              <button
                onClick={() => setFreezeCol(f => !f)}
                className={`px-3 py-2 rounded-xl font-black text-[10px] uppercase transition-all duration-150 ${
                  freezeCol
                    ? "bg-yellow-500/20 border border-yellow-500/40 text-yellow-400"
                    : `${btnGhost}`
                }`}
                title={freezeCol ? "Unfreeze first column" : "Freeze first column"}>
                {freezeCol ? "Frozen" : "Freeze"}
              </button>

              <button onClick={() => fileInputRef.current.click()} className={`px-2 py-1.5 rounded-lg font-black text-[9px] uppercase ${btnGhost}`}>Import File</button>

              {/* ── SMART SAVE / UPDATE BUTTON ── */}
              <button
                onClick={saveToHistory}
                className={`px-3 py-2 rounded-xl font-black text-[10px] uppercase ${
                  isCurrentTableSaved
                    ? "bg-sky-600 hover:bg-sky-500 active:bg-sky-700 text-white shadow-sm shadow-sky-900/30 transition-all duration-150"
                    : btnPrimary
                }`}
                title={isCurrentTableSaved ? "Update existing saved table" : "Save new table to history"}
              >
                {isCurrentTableSaved ? "Update" : "Save"}
              </button>

              <button onClick={downloadExcel} className={`px-2 py-1.5 rounded-lg font-black text-[9px] uppercase ${btnSuccess}`}>Export Excel</button>
              <button onClick={downloadPDF}   className={`px-2 py-1.5 rounded-lg font-black text-[9px] uppercase ${btnDanger}`}>Export PDF</button>
              {ENABLE_AI_TABLE && (
                <button
                  onClick={() => setShowAITable(true)}
                  disabled={aiBuildLoading}
                  className="px-2 py-1.5 rounded-lg font-black text-[9px] uppercase bg-gradient-to-r from-[#2563eb] to-[#0ea5e9] hover:from-[#1d4ed8] hover:to-[#0284c7] text-white transition-all duration-150 shadow-[0_10px_24px_rgba(37,99,235,0.24)] disabled:opacity-50"
                >
                  {aiBuildLoading ? "Building" : "AI Table"}
                </button>
              )}
              <button onClick={() => setShowAIAnalysis(true)}
                className="px-2 py-1.5 rounded-lg font-black text-[9px] uppercase bg-sky-500/10 hover:bg-sky-500 border border-sky-500/20 hover:border-sky-500 text-sky-400 hover:text-white transition-all duration-150">
                 AI
              </button>
              {!isPremium && (
                <button onClick={() => setShowUpgrade(true)}
                  className="px-2 py-1.5 rounded-lg font-black text-[9px] uppercase bg-gradient-to-r from-pink-600 to-red-500 hover:from-pink-500 hover:to-red-400 text-white transition-all duration-150 shadow-sm shadow-pink-900/20">
                  Upgrade
                </button>
              )}
            </div>
          </header>

          {/* ===== SHEET TABS ===== */}
          <div className="flex items-end gap-1 overflow-x-auto" style={{ scrollbarWidth: "none", paddingBottom: 0 }}>
            {tabs.map((tab) => {
              const isActive  = activeTabId === tab.id;
              const tabColor  = tab.tabColor;
              return (
                <div
                  key={tab.id}
                  onClick={() => switchTab(tab.id)}
                  className={`group relative flex items-center gap-2 px-4 py-3 rounded-t-2xl cursor-pointer transition-all duration-150 min-w-[110px] max-w-[180px] border-t border-x select-none ${
                    isActive
                      ? `${bgPanel} ${borderColor} ${textMain} shadow-[0_12px_30px_rgba(37,99,235,0.14)]`
                      : `border-transparent ${textSub} hover:${isDark ? "text-slate-300" : "text-slate-600"}`
                  }`}
                  style={tabColor
                    ? { borderTopColor: tabColor + "80", borderLeftColor: tabColor + "30", borderRightColor: tabColor + "30" }
                    : {}}
                >
                  {tabColor && <div className="w-2 h-2 rounded-full flex-shrink-0" style={{ backgroundColor: tabColor }} />}
                  <span className="text-[10px] font-black uppercase tracking-widest truncate flex-1">{tab.title}</span>

                  <div className="flex items-center gap-0.5 opacity-0 group-hover:opacity-100 transition-opacity">
                    <button
                      onClick={e => { e.stopPropagation(); setEditingFormat({ col: "", rowIdx: null, scope: "tab", tabId: tab.id }); }}
                      className={`w-5 h-5 rounded flex items-center justify-center text-[9px] hover:text-indigo-400 transition-colors ${textSub}`}
                      title="Tab color">
                      🎨
                    </button>
                    {tabs.length > 1 && (
                      <button onClick={e => closeTab(e, tab.id)}
                        className={`w-5 h-5 rounded flex items-center justify-center text-[10px] hover:text-red-400 transition-colors ${textSub}`}>
                        ✕
                      </button>
                    )}
                  </div>

                  {isActive && (
                    <div className="absolute bottom-0 left-3 right-3 h-[2px] rounded-full"
                      style={{ background: tabColor
                        ? `linear-gradient(to right, ${tabColor}, ${tabColor}88)`
                        : "linear-gradient(to right, #6366f1, #38bdf8, #34d399)" }} />
                  )}
                </div>
              );
            })}

            {/* Add tab button */}
            <button
              onClick={() => {
                const check = canCreateTable();
                if (!check.allowed) {
                  setProGateMessage(check.reason);
                  setShowProGate(true);
                } else {
                  setShowTemplates(true);
                }
              }}
              className={`px-4 py-3 rounded-t-2xl font-black text-sm transition-all duration-150 ${
                isDark
                  ? "text-emerald-400 hover:bg-emerald-500/10 hover:text-emerald-300"
                  : "text-emerald-600 hover:bg-emerald-50"
              } ${(!isPremium && lifetimeTablesCreated >= FREE_TABLE_LIMIT) ? "opacity-40 cursor-not-allowed" : ""}`}
              title={!isPremium ? `Free Tables: ${freeTablesRemaining}/${FREE_TABLE_LIMIT}` : "New Sheet"}
            >
              {(!isPremium && lifetimeTablesCreated >= FREE_TABLE_LIMIT) ? `Limit` : `+`}
            </button>
          </div>

          {/* ===== MAIN SHEET PANEL ===== */}
          <div className={`${bgPanel} border ${borderColor} rounded-b-[2rem] rounded-tr-[2rem] shadow-xl`} style={{ borderTopWidth: 0 }}>
            <div className="p-5">

              {/* TOOLS ROW */}
              <div className="grid grid-cols-1 md:grid-cols-2 gap-4 mb-5">
                {/* AI Generator */}
                <div className={`relative ${bgSubtle} border ${borderColor} p-5 rounded-[1.5rem]`}>
                  <div className="flex items-center gap-2 mb-3">
                    <div className="w-2 h-2 bg-indigo-500 rounded-full animate-pulse"></div>
                    <label className="text-[9px] font-black text-indigo-400 uppercase tracking-[0.2em]">AI Generator</label>
                  </div>
                  <input type="text" className={`w-full border rounded-xl px-4 py-3 outline-none text-sm mb-3 transition-all ${inputCls}`}
                    placeholder="e.g. 'Payroll 15 rows' or 'Student List'"
                    value={genText} onChange={e => setGenText(e.target.value)}
                    onKeyDown={e => e.key === "Enter" && handleGenerateTable()} />
                  <button onClick={handleGenerateTable} disabled={isLoading || !genText}
                    className={`w-full py-3 rounded-xl font-black text-[10px] uppercase tracking-widest transition-all flex items-center justify-center gap-2 ${
                      isLoading || !genText
                        ? "opacity-50 cursor-not-allowed bg-indigo-600/50 text-white"
                        : "bg-gradient-to-r from-indigo-600 to-sky-600 hover:from-indigo-500 hover:to-sky-500 text-white shadow-sm shadow-indigo-900/20"
                    }`}>
                    {isLoading ? <><div className="w-3 h-3 border-2 border-white/30 border-t-white rounded-full animate-spin"></div>Building...</> : "Build Structure"}
                  </button>
                </div>

                {/* Calc Engine */}
                <div className={`${bgSubtle} border ${borderColor} p-5 rounded-[1.5rem]`}>
                  <div className="flex items-center gap-2 mb-3">
                    <div className="w-2 h-2 bg-emerald-500 rounded-full"></div>
                    <label className="text-[9px] font-black text-emerald-400 uppercase tracking-[0.2em]">Calc Engine</label>
                  </div>
                  <input type="text" className={`w-full border rounded-xl px-4 py-3 outline-none text-sm mb-3 transition-all ${inputCls}`}
                    placeholder="Formula... (e.g. Salary + Bonus)"
                    value={inputText} onChange={e => setInputText(e.target.value)} />
                  {calculationResult ? (
                    <div className={`p-4 rounded-xl border flex justify-between items-center ${isDark ? "bg-emerald-500/5 border-emerald-500/20" : "bg-emerald-50 border-emerald-200"}`}>
                      <div>
                        <p className="text-[9px] text-emerald-500/60 font-black uppercase">{calculationResult.label}</p>
                        <h3 className="text-2xl font-black text-emerald-400">{calculationResult.value}</h3>
                      </div>
                      <div className="text-right">
                        <p className={`text-[9px] ${textSub} font-mono`}>Formula</p>
                        <p className={`text-[10px] ${textMain} font-black uppercase`}>{calculationResult.formula}</p>
                      </div>
                    </div>
                  ) : (
                    <div className={`py-5 text-center border border-dashed ${borderColor} rounded-xl`}>
                      <p className={`text-[9px] ${textSub} font-black uppercase tracking-widest`}>Enter a formula above</p>
                    </div>
                  )}
                </div>
              </div>

              {/* ANALYTICS TOGGLE */}
              <div className="mb-5">
                <button onClick={() => setShowAnalytics(a => !a)}
                  className={`text-[10px] font-black uppercase tracking-widest px-4 py-2 rounded-xl border transition-all duration-150 ${btnGhost}`}>
                  {showAnalytics ? "▲ Hide" : "▼ Show"} Analytics
                </button>
                {showAnalytics && analyticsData && (
                  <div className="mt-4 grid grid-cols-1 md:grid-cols-2 gap-4">
                    <div className={`${bgSubtle} border ${borderColor} p-5 rounded-[1.5rem]`}>
                      <p className={`text-[9px] font-black uppercase tracking-widest ${textSub} mb-3`}>Bar Chart · {analyticsData.col}</p>
                      <MiniBarChart data={analyticsData.vals} labels={analyticsData.labels} color={isDark ? "#6366f1" : "#4f46e5"} />
                    </div>
                    <div className={`${bgSubtle} border ${borderColor} p-5 rounded-[1.5rem]`}>
                      <p className={`text-[9px] font-black uppercase tracking-widest ${textSub} mb-3`}>Pie Chart · {analyticsData.col}</p>
                      <MiniPieChart data={analyticsData.vals} labels={analyticsData.labels} />
                    </div>
                  </div>
                )}
                {showAnalytics && !analyticsData && (
                  <div className={`mt-4 ${bgSubtle} border ${borderColor} p-5 rounded-[1.5rem] text-center`}>
                    <p className={`text-xs ${textSub}`}>Add numeric columns to see analytics.</p>
                  </div>
                )}
              </div>
              {/* ===== FORMULA BAR ===== */}
              {showFormulaBar && (
                <div className={`mb-3 flex items-center gap-2 px-3 py-2 rounded-xl border ${borderColor} ${isDark ? "bg-[#070f1e]" : "bg-white shadow-sm"}`}>
                  <div className={`px-2 py-1 rounded-lg text-[10px] font-black uppercase tracking-widest flex-shrink-0 ${isDark ? "bg-[#0a1628] text-blue-400 border border-[#1e3a5f]" : "bg-blue-50 text-blue-600 border border-blue-200"}`}
                    style={{ minWidth: 60, textAlign: "center" }}>
                    {selectedCell ? `${selectedCell.col}${selectedCell.rIdx + 1}` : "fx"}
                  </div>
                  <span className={`text-sm font-black ${isDark ? "text-blue-400" : "text-blue-600"} flex-shrink-0`}>=</span>
                  <input
                    type="text"
                    value={formulaBarVal}
                    onChange={e => {
                      setFormulaBarVal(e.target.value);
                      if (selectedCell) handleCellChange(selectedCell.rIdx, selectedCell.col, e.target.value);
                    }}
                    onKeyDown={e => {
                      if (e.key === "Enter" && selectedCell) {
                        handleCellChange(selectedCell.rIdx, selectedCell.col, formulaBarVal);
                        e.target.blur();
                      }
                      if (e.key === "Escape") {
                        setFormulaBarVal(selectedCell ? cellVal(rows[selectedCell.rIdx]?.[selectedCell.col]) : "");
                        e.target.blur();
                      }
                    }}
                    placeholder={selectedCell ? "Type value or =SUM(A1:A5)" : "Click a cell to edit..."}
                    className={`flex-1 bg-transparent outline-none text-sm font-mono ${isDark ? "text-slate-200 placeholder:text-slate-700" : "text-slate-800 placeholder:text-slate-400"}`}
                  />
                  {formulaBarVal.startsWith("=") && (
                    <span className={`text-[9px] font-black px-2 py-1 rounded-lg flex-shrink-0 ${isDark ? "bg-emerald-500/10 text-emerald-400" : "bg-emerald-50 text-emerald-600"}`}>
                      FORMULA
                    </span>
                  )}
                </div>
              )}
              {/* ===== FORMATTING TOOLBAR ===== */}
              {(() => {
                const fmt = getSelectedCellFormat();
                const isBold       = fmt.fontWeight     === "bold";
                const isItalic     = fmt.fontStyle      === "italic";
                const isUnderline  = fmt.textDecoration === "underline";
                const isLeft       = fmt.textAlign      === "left"   || !fmt.textAlign;
                const isCenter     = fmt.textAlign      === "center";
                const isRight      = fmt.textAlign      === "right";
 
                const tbBtn = (active, onClick, children) => (
                  <button
                    onClick={onClick}
                    className={`h-8 px-2.5 rounded-lg text-xs font-black transition-all duration-150 flex items-center justify-center min-w-[32px] ${
                      active
                        ? isDark
                          ? "bg-blue-600 text-white shadow-sm shadow-blue-900/40"
                          : "bg-blue-600 text-white shadow-sm"
                        : isDark
                          ? "text-slate-400 hover:bg-white/8 hover:text-slate-200"
                          : "text-slate-500 hover:bg-slate-100 hover:text-slate-800"
                    }`}
                  >
                    {children}
                  </button>
                );
 
                return (
                  <div className={`mb-4 flex items-center gap-1 px-3 py-2 rounded-2xl border ${borderColor} ${isDark ? "bg-[#0a1628]" : "bg-white shadow-md"} flex-wrap`}>

                    {tbBtn(false, () => handleUndo(),
                      <span style={{ fontSize: 11, fontWeight: 900 }}>↩</span>
                    )}
                    {tbBtn(false, () => handleRedo(),
                      <span style={{ fontSize: 11, fontWeight: 900 }}>↪</span>
                    )}
                    {tbBtn(showFindReplace, () => setShowFindReplace(f => !f),
                      <span style={{ fontSize: 11, fontWeight: 900 }}>🔍</span>
                    )}
                    {tbBtn(false, () => autoFitAllRows(),
                      <span style={{ fontSize: 10, fontWeight: 900 }}>Auto H</span>
                    )}
                    <select value={frozenCols} onChange={e => setFrozenCols(Number(e.target.value))}
                      className={`h-8 px-2 rounded-lg text-[10px] font-black border outline-none ${isDark ? "bg-[#0a1628] border-[#1e3a5f] text-slate-300" : "bg-white border-slate-200 text-slate-700 shadow-sm"}`}>
                      <option value={0}>No Freeze</option>
                      <option value={1}>Freeze 1</option>
                      <option value={2}>Freeze 2</option>
                      <option value={3}>Freeze 3</option>
                    </select>
                    <div className={`w-px h-6 mx-1 ${isDark ? "bg-slate-700" : "bg-slate-200"}`} />
 
                    {/* Cell indicator */}
                    <div className={`px-2 py-1 rounded-lg text-[9px] font-black uppercase tracking-widest mr-2 ${
                      selectedCell
                        ? isDark ? "bg-blue-600/20 text-blue-400" : "bg-blue-50 text-blue-600"
                        : isDark ? "bg-slate-800 text-slate-600" : "bg-slate-100 text-slate-400"
                    }`}>
                      {selectedCell ? `${selectedCell.col} · R${selectedCell.rIdx + 1}` : "No cell"}
                    </div>
 
                    {/* Divider */}
                    <div className={`w-px h-6 mx-1 ${isDark ? "bg-slate-700" : "bg-slate-200"}`} />
 
                    {/* Font Size */}
                    <select
                      value={fmt.fontSize || 14}
                      onChange={e => applyToolbarFormat({ fontSize: Number(e.target.value) })}
                      disabled={!selectedCell}
                      className={`h-8 px-2 rounded-lg text-[11px] font-black border outline-none transition-all cursor-pointer ${
                        isDark
                          ? "bg-[#0a1628] border-[#1e3a5f] text-slate-300 disabled:opacity-30"
                          : "bg-white border-slate-200 text-slate-700 disabled:opacity-30 shadow-sm"
                      }`}
                    >
                      {[10, 11, 12, 13, 14, 15, 16, 18, 20, 22, 24].map(s => (
                        <option key={s} value={s}>{s}px</option>
                      ))}
                    </select>
 
                    {/* Divider */}
                    <div className={`w-px h-6 mx-1 ${isDark ? "bg-slate-700" : "bg-slate-200"}`} />
 
                    {/* Bold */}
                    {tbBtn(isBold, () => applyToolbarFormat({ fontWeight: isBold ? "normal" : "bold" }),
                      <span style={{ fontWeight: 900, fontSize: 13 }}>B</span>
                    )}
 
                    {/* Italic */}
                    {tbBtn(isItalic, () => applyToolbarFormat({ fontStyle: isItalic ? "normal" : "italic" }),
                      <span style={{ fontStyle: "italic", fontWeight: 700, fontSize: 13 }}>I</span>
                    )}
 
                    {/* Underline */}
                    {tbBtn(isUnderline, () => applyToolbarFormat({ textDecoration: isUnderline ? "none" : "underline" }),
                      <span style={{ textDecoration: "underline", fontWeight: 700, fontSize: 13 }}>U</span>
                    )}
 
                    {/* Divider */}
                    <div className={`w-px h-6 mx-1 ${isDark ? "bg-slate-700" : "bg-slate-200"}`} />
 
                    {/* Align Left */}
                    {tbBtn(isLeft, () => applyToolbarFormat({ textAlign: "left" }),
                      <svg width="14" height="14" viewBox="0 0 14 14" fill="currentColor">
                        <rect x="0" y="1" width="14" height="2" rx="1"/>
                        <rect x="0" y="5" width="10" height="2" rx="1"/>
                        <rect x="0" y="9" width="14" height="2" rx="1"/>
                        <rect x="0" y="13" width="8" height="2" rx="1"/>
                      </svg>
                    )}
 
                    {/* Align Center */}
                    {tbBtn(isCenter, () => applyToolbarFormat({ textAlign: "center" }),
                      <svg width="14" height="14" viewBox="0 0 14 14" fill="currentColor">
                        <rect x="0" y="1" width="14" height="2" rx="1"/>
                        <rect x="2" y="5" width="10" height="2" rx="1"/>
                        <rect x="0" y="9" width="14" height="2" rx="1"/>
                        <rect x="3" y="13" width="8" height="2" rx="1"/>
                      </svg>
                    )}
 
                    {/* Align Right */}
                    {tbBtn(isRight, () => applyToolbarFormat({ textAlign: "right" }),
                      <svg width="14" height="14" viewBox="0 0 14 14" fill="currentColor">
                        <rect x="0" y="1" width="14" height="2" rx="1"/>
                        <rect x="4" y="5" width="10" height="2" rx="1"/>
                        <rect x="0" y="9" width="14" height="2" rx="1"/>
                        <rect x="6" y="13" width="8" height="2" rx="1"/>
                      </svg>
                    )}
 
                    {/* Divider */}
                    <div className={`w-px h-6 mx-1 ${isDark ? "bg-slate-700" : "bg-slate-200"}`} />
 
                    {/* Font Color */}
                    <div className="relative flex items-center">
                      <label className={`h-8 px-2 rounded-lg flex items-center gap-1.5 cursor-pointer transition-all text-[10px] font-black ${
                        isDark ? "text-slate-400 hover:bg-white/8 hover:text-slate-200" : "text-slate-500 hover:bg-slate-100"
                      } ${!selectedCell ? "opacity-30 pointer-events-none" : ""}`}>
                        <span>A</span>
                        <div className="w-4 h-1.5 rounded-full" style={{ backgroundColor: fmt.color || (isDark ? "#e2e8f0" : "#0f172a") }} />
                        <input type="color"
                          value={fmt.color || (isDark ? "#e2e8f0" : "#0f172a")}
                          onChange={e => applyToolbarFormat({ color: e.target.value })}
                          className="absolute opacity-0 w-0 h-0" />
                      </label>
                    </div>
 
                    {/* Bg Color */}
                    <div className="relative flex items-center">
                      <label className={`h-8 px-2 rounded-lg flex items-center gap-1.5 cursor-pointer transition-all text-[10px] font-black ${
                        isDark ? "text-slate-400 hover:bg-white/8 hover:text-slate-200" : "text-slate-500 hover:bg-slate-100"
                      } ${!selectedCell ? "opacity-30 pointer-events-none" : ""}`}>
                        <svg width="12" height="12" viewBox="0 0 12 12" fill="currentColor">
                          <rect width="12" height="12" rx="2"/>
                        </svg>
                        <div className="w-4 h-1.5 rounded-full border" style={{
                          backgroundColor: fmt.bgColor || "transparent",
                          borderColor: isDark ? "#334155" : "#e2e8f0"
                        }} />
                        <input type="color"
                          value={fmt.bgColor || "#ffffff"}
                          onChange={e => applyToolbarFormat({ bgColor: e.target.value })}
                          className="absolute opacity-0 w-0 h-0" />
                      </label>
                    </div>
 
                    {/* Clear Format */}
                    <button
                      onClick={() => {
                        if (!selectedCell) return;
                        applyToolbarFormat({ color: "", bgColor: "", fontSize: 14, fontWeight: "normal", fontStyle: "normal", textAlign: "left", textDecoration: "none" });
                      }}
                      disabled={!selectedCell}
                      className={`h-8 px-2.5 rounded-lg text-[10px] font-black transition-all ml-1 ${
                        isDark
                          ? "text-red-400/60 hover:bg-red-500/10 hover:text-red-400 disabled:opacity-20"
                          : "text-red-400 hover:bg-red-50 disabled:opacity-20"
                      }`}
                      title="Clear all formatting"
                    >
                      Clear
                    </button>

                    <div className={`w-px h-6 mx-1 ${isDark ? "bg-slate-700" : "bg-slate-200"}`} />

                    {selectionStart && selectionEnd && !(selectionStart.rIdx === selectionEnd.rIdx && selectionStart.cIdx === selectionEnd.cIdx) && (
                      <div className={`px-2 py-1 rounded-lg text-[9px] font-black uppercase ${
                        isDark ? "bg-blue-600/20 text-blue-400" : "bg-blue-50 text-blue-600"
                      }`}>
                        {Math.abs(selectionEnd.rIdx - selectionStart.rIdx) + 1} × {Math.abs(selectionEnd.cIdx - selectionStart.cIdx) + 1} selected
                      </div>
                    )}

                    {selectionStart && selectionEnd && !(selectionStart.rIdx === selectionEnd.rIdx && selectionStart.cIdx === selectionEnd.cIdx) && (
                      <button
                        onClick={() => applyMultiCellFormat({ fontWeight: fmt.fontWeight, color: fmt.color, bgColor: fmt.bgColor, fontSize: fmt.fontSize })}
                        className={`h-8 px-2.5 rounded-lg text-[10px] font-black transition-all ${
                          isDark ? "bg-blue-600/20 text-blue-400 hover:bg-blue-600 hover:text-white" : "bg-blue-50 text-blue-600 hover:bg-blue-600 hover:text-white"
                        }`}>
                        Apply All
                      </button>
                    )}

                    <button
                      onClick={() => setShowBorderModal(true)}
                      disabled={!selectedCell}
                      className={`h-8 px-2.5 rounded-lg text-[10px] font-black transition-all ${
                        isDark ? "text-slate-400 hover:bg-white/8 hover:text-slate-200 disabled:opacity-30" : "text-slate-500 hover:bg-slate-100 disabled:opacity-30"
                      }`}>
                      Border
                    </button>

                    <button
                      onClick={() => setShowConditional(true)}
                      className={`h-8 px-2.5 rounded-lg text-[10px] font-black transition-all ${
                        conditionalRules.length > 0
                          ? isDark ? "bg-purple-600/20 text-purple-400" : "bg-purple-50 text-purple-600"
                          : isDark ? "text-slate-400 hover:bg-white/8 hover:text-slate-200" : "text-slate-500 hover:bg-slate-100"
                      }`}>
                      {conditionalRules.length > 0 ? `Rules (${conditionalRules.length})` : "Conditions"}
                    </button>

                  </div>
                );
              })()}
              {/* FIND & REPLACE */}
              {showFindReplace && (
                <div className={`mb-4 p-4 rounded-2xl border ${borderColor} ${isDark ? "bg-[#070f1e]" : "bg-slate-50"} flex flex-wrap items-center gap-3`}>
                  <div className="flex items-center gap-2 flex-1 min-w-[200px]">
                    <input
                      type="text"
                      value={findText}
                      onChange={e => setFindText(e.target.value)}
                      onKeyDown={e => e.key === "Enter" && handleFind()}
                      placeholder="Find..."
                      className={`flex-1 px-3 py-2 rounded-xl text-xs border outline-none ${inputCls}`}
                    />
                    <input
                      type="text"
                      value={replaceText}
                      onChange={e => setReplaceText(e.target.value)}
                      placeholder="Replace with..."
                      className={`flex-1 px-3 py-2 rounded-xl text-xs border outline-none ${inputCls}`}
                    />
                  </div>
                  <div className="flex items-center gap-2 flex-wrap">
                    <button onClick={handleFind}
                      className={`px-3 py-2 rounded-xl font-black text-[10px] uppercase ${btnPrimary}`}>
                      Find
                    </button>
                    <button onClick={handleReplaceOne}
                      className={`px-3 py-2 rounded-xl font-black text-[10px] uppercase ${btnGhost}`}
                      disabled={findResults.length === 0}>
                      Replace
                    </button>
                    <button onClick={handleReplaceAll}
                      className={`px-3 py-2 rounded-xl font-black text-[10px] uppercase ${btnGhost}`}
                      disabled={findResults.length === 0}>
                      Replace All
                    </button>
                    {findResults.length > 0 && (
                      <span className={`text-[10px] font-black ${isDark ? "text-blue-400" : "text-blue-600"}`}>
                        {findIndex + 1}/{findResults.length} found
                      </span>
                    )}
                    {findText && findResults.length === 0 && (
                      <span className="text-[10px] font-black text-red-400">Not found</span>
                    )}
                    <button onClick={() => { setShowFindReplace(false); setFindText(""); setReplaceText(""); setFindResults([]); }}
                      className="text-slate-400 hover:text-slate-200 text-lg">✕</button>
                  </div>
                </div>
              )}

              {/* TITLE + SEARCH */}
              <div className={`${bgSubtle} border ${borderColor} px-4 py-3.5 rounded-2xl mb-4 flex flex-col md:flex-row justify-between items-center gap-3`}>
                <input type="text" className="bg-transparent border-none outline-none text-xl font-black placeholder:text-slate-700/30 flex-1 w-full min-w-0"
                  placeholder="Untitled Document" value={tableTitle} onChange={e => setTableTitle(e.target.value)} />
                <input type="text" className={`border rounded-xl px-4 py-2 outline-none text-xs w-full md:w-52 transition-all flex-shrink-0 ${inputCls}`}
                  placeholder="🔍 Search records..." value={searchTerm} onChange={e => setSearchTerm(e.target.value)} />
              </div>

              {/* BULK ACTIONS */}
              {selectedRows.size > 0 && (
  <div className={`mb-4 flex flex-wrap items-center gap-2 px-5 py-3 rounded-2xl border ${isDark ? "bg-[#132844] border-[#29456b]" : "bg-[#f5f8fd] border-[#c8d7ec]"}`}>
    <span className="text-[11px] font-black text-[#3978d8] mr-2">
      {selectedRows.size} row(s) selected
    </span>

    <button onClick={duplicateSelectedRows} className={`px-4 py-2 rounded-xl font-black text-[10px] uppercase ${btnGhost}`}>
      Duplicate
    </button>

    <button onClick={fillDownSelectedRows} className={`px-4 py-2 rounded-xl font-black text-[10px] uppercase ${btnGhost}`}>
      Fill Down
    </button>

    <button onClick={deleteSelectedRows} className={`px-4 py-2 rounded-xl font-black text-[10px] uppercase ${btnDanger}`}>
      Delete
    </button>

    <button onClick={() => setSelectedRows(new Set())} className={`px-4 py-2 rounded-xl font-black text-[10px] uppercase ${btnGhost}`}>
      Clear
    </button>
  </div>
)}

              {/* ===== TABLE ===== */}
              <div className={`${bgSubtle} rounded-2xl border ${borderColor} shadow-inner overflow-hidden`}>
                <div className="ss-scroll-container" ref={tableRef} style={{ overflowX: "auto", overflowY: "auto", maxHeight: "520px" }}>
                  <table className="border-collapse" style={{ tableLayout: "auto", width: "max-content", minWidth: "100%" }}>
                    <thead className={`${tableHeadCls} sticky top-0 z-10 border-b ${borderColor}`}>
                      <tr>
                        <th className={`px-3 py-4 border-r ${borderColor} sticky left-0 z-20`}
                          style={{ background: isDark ? "#070c17" : "#f8fafc", width: "44px" }}>
                          <input type="checkbox"
                            checked={selectedRows.size === rows.length && rows.length > 0}
                            onChange={e => setSelectedRows(e.target.checked ? new Set(rows.map((_, i) => i)) : new Set())}
                            className="accent-indigo-500" />
                        </th>
                        {columns.map((col, i) => {
                          const colFmt = columnFormats[col];
                          if (hiddenCols.has(col)) return null;
                          return (
                            <th key={i} draggable={col !== "ID"}
                              onDragStart={() => handleColDragStart(i)}
                              onDragOver={e => handleColDragOver(e, i)}
                              onDrop={() => handleColDrop(i)}
                              className={`py-4 font-black text-[10px] uppercase border-r ${borderColor} tracking-widest whitespace-nowrap transition-all ${dragOverColIdx === i ? "bg-indigo-500/10" : ""}`}
                              style={{
                                minWidth: colWidths[col] ? `${colWidths[col]}px` : getColMinWidth(col),
                                width: col === "ID" ? "60px" : (colWidths[col] ? `${colWidths[col]}px` : undefined),
                                backgroundColor: colFmt?.bgColor ? colFmt.bgColor + "18" : undefined,
                                padding: "0 20px",
                                cursor: col !== "ID" ? "grab" : "default",
                                color: isDark ? "#d7e6ff" : "#18324f",
                                position: "relative",
                              }}>
                              <span onClick={e => handleColumnHeaderClick(e, col)}
                                className="cursor-pointer transition-colors flex items-center gap-2 hover:text-indigo-400 group select-none"
                                onMouseDown={e => e.stopPropagation()}
                                style={{ userSelect: "none", display: "flex", visibility: "visible" }}>
                                <span
                                      style={{
                                      fontWeight: 900,
                                      letterSpacing: "0.04em",
                                      color: isDark ? "#d7e6ff" : "#18324f",
                                      visibility: "visible",
                                      opacity: 1,
                                      display: "inline-block"
                                    }}
                                >
                                    {col}
                                </span>
                                {columnTypes[col] && (
                                  <span className="text-indigo-400/60 text-[10px]">
                                    {columnTypes[col] === "dropdown" ? "⊞" : columnTypes[col] === "date" ? "◷" : columnTypes[col] === "checkbox" ? "☑" : columnTypes[col] === "email" ? "@" : ""}
                                  </span>
                                )}
                                {col !== "ID" && <span className="opacity-0 group-hover:opacity-50 text-indigo-500 text-[9px]">▾</span>}
                              </span>                              
                              {col !== "ID" && (
                                <div
                                  onMouseDown={e => handleColResizeStart(e, col)}
                                  className="absolute right-0 top-0 bottom-0 w-1.5 cursor-col-resize group/resize flex items-center justify-center"
                                  style={{ zIndex: 20 }}>
                                  <div className="w-0.5 h-4 rounded-full opacity-0 group-hover/resize:opacity-100 transition-opacity"
                                    style={{ backgroundColor: "#d4af37" }} />
                                </div>
                              )}
                            </th>
                          );
                        })}
                        <th className={`px-4 py-4 ${tableHeadCls} border-r ${borderColor}`} style={{ width: "56px" }}>
                          <button onClick={addColumn}
                            className={`w-7 h-7 rounded-xl border text-sm font-black transition-all ${
                              isDark
                                ? "bg-slate-800 border-slate-700 text-slate-400 hover:bg-indigo-500/20 hover:border-indigo-500/40 hover:text-indigo-400"
                                : "bg-slate-100 border-slate-200 text-slate-400 hover:bg-indigo-50 hover:border-indigo-200 hover:text-indigo-600"
                            }`}>+</button>
                        </th>
                        <th className={`${tableHeadCls}`} style={{ width: "40px" }}></th>
                      </tr>
                    </thead>

                    <tbody className={`divide-y ${dividerCls}`}>
                      {filteredRows.map(({ row, originalIndex }) => (
                        <tr key={originalIndex} draggable
                          onDragStart={() => handleRowDragStart(originalIndex)}
                          onDragOver={e => handleRowDragOver(e, originalIndex)}
                          onDrop={() => handleRowDrop(originalIndex)}
                          onDoubleClick={() => autoFitRowHeight(originalIndex)}
                          className={`transition-all group/row ${getRowStyle(row)} ${dragOverRowIdx === originalIndex ? "opacity-40" : ""} ${isDark ? "hover:bg-white/[0.012]" : "hover:bg-slate-50/80"}`}
                          style={{
                            cursor: "grab",
                            height: rowHeights[originalIndex] ? `${rowHeights[originalIndex]}px` : undefined
                          }}>
                          <td className={`px-3 py-0 border-r ${borderColor} sticky left-0`}
                            style={{ background: isDark ? "#0a1018" : "#ffffff", zIndex: 5, width: "44px", position: freezeCol ? "sticky" : "relative", left: freezeCol ? 0 : undefined }}>
                            <input type="checkbox" checked={selectedRows.has(originalIndex)}
                              onChange={e => { const s = new Set(selectedRows); e.target.checked ? s.add(originalIndex) : s.delete(originalIndex); setSelectedRows(s); }}
                              className="accent-indigo-500" />
                          </td>
                          {columns.map((col, cIdx) => {
                            const cell    = row[col];
                            const cellObj = cell && typeof cell === "object" ? cell : null;
                            const colFmt  = columnFormats[col];
                            const tdBg    = cellObj?.bgColor || (colFmt?.bgColor ? colFmt.bgColor + "12" : undefined);
                            if (hiddenCols.has(col)) return null;
                            return (
                              <td key={cIdx} className={`border-r ${borderColor} p-0 relative group/cell transition-colors ${
                                isCellInSelection(originalIndex, cIdx)
                                  ? isDark
                                    ? "bg-blue-600/20 ring-1 ring-inset ring-blue-500/40"
                                    : "bg-blue-100/60 ring-1 ring-inset ring-blue-400/40"
                                  : ""
                              }`}
                                style={{ minWidth: getColMinWidth(col), backgroundColor: tdBg,position: cIdx < frozenCols ? "sticky" : "relative",
                                left: cIdx < frozenCols ? `${44 + cIdx * 150}px` : undefined,
                                zIndex: cIdx < frozenCols ? 4 : undefined,
                                background: cIdx < frozenCols ? (isDark ? "#0a1018" : "#ffffff") : (tdBg || undefined), overflow: "visible", whiteSpace: "nowrap", userSelect: "text", WebkitUserSelect: "text", cursor: "text" }}
                                onMouseDown={e => {
                                  e.stopPropagation();
                                  setSelectionStart({ rIdx: originalIndex, cIdx });
                                  setSelectionEnd({ rIdx: originalIndex, cIdx });
                                  setIsSelecting(true);
                                  setSelectedCell({ rIdx: originalIndex, col });
                                }}
                                onMouseEnter={() => {
                                  if (isSelecting) setSelectionEnd({ rIdx: originalIndex, cIdx });
                                }}
                                onMouseUp={() => setIsSelecting(false)}
                                onContextMenu={e => {
                                  e.preventDefault();
                                  const menu = document.createElement("div");
                                  menu.style.cssText = `position:fixed;top:${e.clientY}px;left:${e.clientX}px;z-index:9999;background:${isDark?"#08122a":"#fff"};border:1px solid ${isDark?"#d4af3722":"#e2e8f0"};border-radius:12px;padding:6px;min-width:140px;box-shadow:0 8px 32px rgba(0,0,0,0.4)`;
                                  const btnStyle = `display:block;width:100%;text-align:left;padding:8px 12px;font-size:10px;font-weight:900;letter-spacing:0.08em;text-transform:uppercase;border-radius:8px;cursor:pointer;color:${isDark?"#94a3b8":"#475569"};background:transparent;border:none;`;
                                  menu.innerHTML = `<button style="${btnStyle}" id="nm-copy">📋 Copy</button><button style="${btnStyle}" id="nm-paste">📌 Paste</button><button style="${btnStyle}" id="nm-clear">🧹 Clear</button>`;
                                  document.body.appendChild(menu);
                                  document.getElementById("nm-copy").onclick = () => { handleCopyCell(originalIndex, col); document.body.removeChild(menu); };
                                  document.getElementById("nm-paste").onclick = () => { handlePasteCell(originalIndex, col); document.body.removeChild(menu); };
                                  document.getElementById("nm-clear").onclick = () => { handleCellChange(originalIndex, col, ""); document.body.removeChild(menu); };
                                  const close = () => { if(document.body.contains(menu)) document.body.removeChild(menu); document.removeEventListener("click", close); };
                                  setTimeout(() => document.addEventListener("click", close), 100);
                                }}>
                                {renderCell(row, originalIndex, col, cIdx)}
                                {col !== "ID" && (
                                  <button title="Format cell"
                                    onClick={e => { e.stopPropagation(); setEditingFormat({ col, rowIdx: originalIndex, scope: "cell" }); }}
                                    className="absolute top-1 right-1 opacity-0 group-hover/cell:opacity-100 transition-opacity bg-indigo-500/15 hover:bg-indigo-500 text-indigo-300 hover:text-white rounded-md px-1.5 text-[8px] font-black z-10 uppercase tracking-wider"
                                    style={{ pointerEvents: "auto" }}>
                                    fx
                                  </button>
                                )}
                              </td>
                            );
                          })}
                          <td className={`border-r ${borderColor} p-2`}></td>
                          {/* Row resize handle */}
                          <td className="p-0 relative" style={{ width: "4px" }}>
                            <div
                              onMouseDown={e => handleRowResizeStart(e, originalIndex)}
                              className="absolute bottom-0 left-0 right-0 h-1 cursor-row-resize group/rowresize flex items-center justify-center"
                              style={{ zIndex: 10 }}>
                              <div className="w-6 h-0.5 rounded-full opacity-0 group-hover/rowresize:opacity-100 transition-opacity"
                                style={{ backgroundColor: "#d4af37" }} />
                            </div>
                          </td>
                          <td className="text-center py-0 opacity-0 group-hover/row:opacity-100 transition-opacity px-2" style={{ width: "40px" }}>
                          <button
                            onClick={() => {
                              setHiddenRows(prev => {
                                const next = new Set(prev);
                                next.add(originalIndex);
                                return next;
                              });
                            }}
                            className={`opacity-0 group-hover/row:opacity-100 transition-opacity text-[9px] font-black px-2 py-1 rounded-lg mr-1 ${isDark ? "text-slate-500 hover:bg-slate-800 hover:text-slate-300" : "text-slate-400 hover:bg-slate-100"}`}
                            title="Hide Row"
                          >
                            Hide
                          </button>
                            <button onClick={() => deleteRow(originalIndex)} className="text-red-500/60 hover:text-red-500 transition-colors text-[9px] font-black uppercase flex-shrink-0">✖</button>
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>

                {/* Table Footer */}
                <div className={`px-5 py-4 border-t ${borderColor} ${bgSubtle} flex flex-col md:flex-row justify-between items-start md:items-center gap-3`}>
                <div className="flex flex-wrap items-center gap-2">
                  <button onClick={addRow}
                    className={`px-4 py-2 rounded-xl font-black text-[10px] uppercase ${btnPrimary}`}>
                    Add Row
                  </button>

                  {(hiddenRows.size > 0 || hiddenCols.size > 0) && (
                    <div className="flex items-center gap-2">
                      {hiddenRows.size > 0 && (
                        <button onClick={() => setHiddenRows(new Set())}
                          className={`px-3 py-2 rounded-xl font-black text-[10px] uppercase ${isDark ? "bg-amber-500/10 text-amber-400 hover:bg-amber-500/20" : "bg-amber-50 text-amber-600"}`}>
                          Show {hiddenRows.size} hidden rows
                        </button>
                      )}
                      {hiddenCols.size > 0 && (
                        <button onClick={() => setHiddenCols(new Set())}
                          className={`px-3 py-2 rounded-xl font-black text-[10px] uppercase ${isDark ? "bg-amber-500/10 text-amber-400 hover:bg-amber-500/20" : "bg-amber-50 text-amber-600"}`}>
                          Show {hiddenCols.size} hidden cols
                        </button>
                      )}
                    </div>
                  )}
                  <button onClick={autoNumberRows}
                    className={`px-4 py-2 rounded-xl font-black text-[10px] uppercase ${btnGhost}`}>
                    Auto Number
                  </button>

                  <button onClick={cleanData}
                    className={`px-4 py-2 rounded-xl font-black text-[10px] uppercase ${btnGhost}`}>
                    Clean Spaces
                  </button>

                  <button onClick={removeEmptyRows}
                    className={`px-4 py-2 rounded-xl font-black text-[10px] uppercase ${btnGhost}`}>
                    Remove Empty
                  </button>
                </div>
                  <div className="flex items-center gap-4">
                    <span className={`text-[9px] ${textSub} font-mono`}>{rows.length} rows · {columns.length} cols</span>
                    {copiedCell && (
                      <span className="text-[9px] text-yellow-400 font-black flex items-center gap-1">
                        📋 Copied: <span className="font-mono">{String(copiedCell.value).slice(0, 12)}</span>
                        <button onClick={() => setCopiedCell(null)} className="text-[8px] hover:opacity-60">✕</button>
                      </span>
                    )}
                    {searchTerm && <span className="text-[9px] text-indigo-400 font-black">{filteredRows.length} match{filteredRows.length !== 1 ? "es" : ""}</span>}
                  </div>
                </div>
              </div>
            </div>
          </div>
        </div>

        {/* ======== SIDEBAR ======== */}
        <div className="lg:col-span-1 flex flex-col gap-4">
          {!isPremium && <SidebarAd isDark={isDark} />}

          {/* History Panel */}
          <div className={`${bgPanel} border ${borderColor} rounded-[2rem] p-5 shadow-xl flex flex-col flex-1`}>
            <div className={`flex justify-between items-center mb-5 pb-4 border-b ${borderColor}`}>
              <div>
                <h2 className="text-[11px] font-black text-sky-400 uppercase tracking-[0.2em]">Vault History</h2>
                {!isPremium && (
                  <p className={`text-[8px] ${textSub} font-black mt-0.5`}>{history.length}/{FREE_TABLE_LIMIT} saved</p>
                )}
              </div>
              {history.length > 0 && (
                <button onClick={clearAllHistory} className="text-[9px] font-black uppercase text-red-400 hover:text-red-300 transition-colors">Clear</button>
              )}
            </div>

            <div className="space-y-2 overflow-y-auto flex-1" style={{ scrollbarWidth: "none" }}>
              {history.length === 0 ? (
                <div className="flex flex-col items-center justify-center py-12 gap-3">
                  <div className={`w-12 h-12 rounded-2xl flex items-center justify-center text-[10px] font-black uppercase tracking-widest ${isDark ? "bg-[#132844] text-[#9fb3d1]" : "bg-[#e8f0fb] text-[#36516f]"}`}>
                    Empty
                  </div>
                  <p className={`text-[10px] ${textSub} font-black uppercase tracking-widest text-center`}>No saved tables</p>
                  <p className={`text-[9px] ${textSub} text-center`}>Press Save to store your current sheet</p>
                </div>
              ) : (
                history.map((item, idx) => (
                  <React.Fragment key={item.id}>
                    {!isPremium && idx === 3 && history.length > 4 && <InFeedAd isDark={isDark} />}
                    <div
                      onClick={() => handleHistoryClick(item)}
                      className={`p-3.5 rounded-2xl border transition-all duration-150 cursor-pointer ${
                        currentTableId === item.id
                          ? `border-sky-500/40 ${isDark ? "bg-sky-500/5" : "bg-sky-50"}`
                          : `${borderColor} ${isDark ? "bg-[#0a1018]/60 hover:border-slate-700" : "bg-slate-50 hover:border-slate-300"}`
                      }`}>
                      <div className="flex justify-between items-start">
                        <div className="flex-1 min-w-0 pr-2">
                          <div className="flex items-center gap-1.5 mb-0.5">
                            <h4 className={`font-black text-xs truncate uppercase ${currentTableId === item.id ? "text-sky-400" : textMain}`}>
                              {item.title}
                            </h4>
                            {/* Badge indicating if this is the currently loaded + saved table */}
                            {currentTableId === item.id && (
                              <span className="text-[7px] font-black px-1.5 py-0.5 rounded-full bg-sky-500/20 text-sky-400 uppercase flex-shrink-0">Active</span>
                            )}
                          </div>
                          <p className={`text-[8px] ${textSub} mt-0.5 font-mono`}>{item.date}</p>
                          <p className={`text-[8px] ${textSub} mt-0.5`}>{item.columns?.length} cols · {item.rows?.length} rows</p>
                        </div>
                        <button onClick={e => deleteHistoryItem(e, item.id)}
                          className="text-red-500/25 hover:text-red-500 transition-colors text-sm flex-shrink-0">Delete</button>
                      </div>
                    </div>
                  </React.Fragment>
                ))
              )}
            </div>

            {!isPremium && history.length >= FREE_TABLE_LIMIT && (
              <button onClick={() => setShowUpgrade(true)}
                className="mt-4 w-full py-3 rounded-2xl font-black text-[11px] uppercase bg-gradient-to-r from-pink-600 to-red-500 hover:from-pink-500 hover:to-red-400 text-white transition-all shadow-sm shadow-pink-900/20">
                ⭐ Upgrade for Unlimited
              </button>
            )}
          </div>

          {/* Template Quick-Access */}
          <div className={`${bgPanel} border ${borderColor} rounded-[2rem] p-5 shadow-xl`}>
            <h2 className={`text-[10px] font-black uppercase tracking-[0.2em] mb-4 ${isDark ? "text-emerald-400" : "text-emerald-600"}`}>
              Quick Templates
            </h2>
            <div className="grid grid-cols-2 gap-2">
              {TEMPLATES.slice(0, 6).map(tmpl => (
                <button key={tmpl.id}
                  onClick={() => { if (tmpl.id === "blank") addNewTab(null); else addNewTab(tmpl); }}
                  className={`p-2.5 rounded-xl border text-left transition-all duration-150 group ${borderColor} ${isDark ? "bg-[#0a1018] hover:bg-[#0f1929]" : "bg-slate-50 hover:bg-white"}`}>
                  <div className="text-[8px] font-black uppercase tracking-widest mb-1">{tmpl.icon}</div>
                  <p className={`text-[9px] font-black uppercase tracking-wide ${textMain} truncate`}>{tmpl.name}</p>
                </button>
              ))}
            </div>
            <button onClick={() => setShowTemplates(true)}
              className={`mt-3 w-full py-2.5 rounded-xl font-black text-[10px] uppercase tracking-widest border transition-all duration-150 ${btnGhost}`}>
              View All Templates
            </button>
          </div>
        </div>
      </div>

      {/* ============================================================
          GLOBAL STYLES
          ============================================================ */}
      <style>{`
        * { -webkit-user-select: none !important; -moz-user-select: none !important; user-select: none !important; }
        td, td input, td select { -webkit-user-select: text !important; -moz-user-select: text !important; user-select: text !important; }
        th { -webkit-user-select: none !important; user-select: none !important; } -moz-user-select: text !important; user-select: text !important; }
        th { -webkit-user-select: none !important; user-select: none !important; }
        th span.select-none { -webkit-user-select: none !important; user-select: none !important; }
        .flex-1.overflow-y-auto::-webkit-scrollbar { display: none; }
        /* Mobile Responsive */
        @media (max-width: 768px) {
          .mobile-hide { display: none !important; }
          .mobile-full { width: 100% !important; }
          .mobile-stack { flex-direction: column !important; }
          .mobile-small { font-size: 9px !important; padding: 4px 8px !important; }
          .mobile-scroll { overflow-x: auto !important; -webkit-overflow-scrolling: touch; }
        }

        .ss-scroll-container {
          scrollbar-width: thin;
          scrollbar-color: rgba(99,102,241,0.3) transparent;
        }
        .ss-scroll-container::-webkit-scrollbar { display: block !important; height: 6px; width: 6px; }
        .ss-scroll-container::-webkit-scrollbar-track {
          background: ${isDark ? "rgba(255,255,255,0.02)" : "rgba(0,0,0,0.02)"};
          border-radius: 99px; margin: 0 12px;
        }
        .ss-scroll-container::-webkit-scrollbar-thumb {
          background: linear-gradient(90deg, #6366f1aa, #22d3eeaa);
          border-radius: 99px;
        }
        .ss-scroll-container::-webkit-scrollbar-thumb:hover {
          background: linear-gradient(90deg, #6366f1, #22d3ee);
        }
        .ss-scroll-container::-webkit-scrollbar-corner { background: transparent; }

        input[type="date"]::-webkit-calendar-picker-indicator {
          filter: ${isDark ? "invert(1) opacity(0.35)" : "opacity(0.4)"};
          cursor: pointer;
        }
        select option {
          background: ${isDark ? "#0f1929" : "#fff"};
          color: ${isDark ? "#e2e8f0" : "#0f172a"};
        }
        table input, table select { line-height: 1.6; letter-spacing: 0.01em; }
        table td input[type="text"], table td input[type="date"] { cursor: text; }
        table td { overflow: visible !important; }
        table td input { white-space: nowrap; }
        thead th span { visibility: visible !important; opacity: 1 !important; display: flex !important; }
        button:active { transform: scale(0.97); }
        button:active.no-scale { transform: none; }
      `}</style>
    </div>
  );
}

export default App;






