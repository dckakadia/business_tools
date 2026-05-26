import { LayoutDashboard, Box, Factory, FileText, ShoppingCart, Printer, Bath, Check, AlertTriangle, User, Trash2, X, Plus, Search, Zap, EyeOff, Eye, Package, IndianRupee, TrendingUp, Download, FolderOpen, RefreshCw, CheckCircle, Loader2, Upload, Type, ArrowLeft, Tag, ClipboardList, Edit2, Calendar, Save, Hash, BarChart2, Sparkles, ChevronUp, ChevronDown, ChevronsUpDown, Minus } from "lucide-react";
import { useState, useEffect, useRef, useCallback } from "react";

const TABS = ["Dashboard", "Items", "Products", "Cost Sheet", "Purchase Order"];
const CATEGORIES_ITEM = ["Raw Material", "Chemical", "Electrical", "Fittings", "Plumbing", "Metal", "Other"];
const CATEGORIES_PROD = ["Spa", "Jacuzzi", "Hot Tub", "Swim Spa", "Other"];

const initialItems = [
  { id: "i1", name: "Acrylic Sheet", unit: "pcs", cost: 4500, category: "Raw Material" },
  { id: "i2", name: "Fiberglass Resin", unit: "kg", cost: 320, category: "Chemical" },
  { id: "i3", name: "Pump Motor 1HP", unit: "pcs", cost: 8500, category: "Electrical" },
  { id: "i4", name: "Jets Nozzle", unit: "pcs", cost: 150, category: "Fittings" },
  { id: "i5", name: "PVC Pipe 2 inch", unit: "mtr", cost: 85, category: "Plumbing" },
  { id: "i6", name: "Heater Element", unit: "pcs", cost: 1200, category: "Electrical" },
  { id: "i7", name: "Control Panel", unit: "pcs", cost: 3200, category: "Electrical" },
  { id: "i8", name: "Insulation Foam", unit: "kg", cost: 210, category: "Raw Material" },
  { id: "i9", name: "LED Light Strip", unit: "mtr", cost: 420, category: "Electrical" },
  { id: "i10", name: "Stainless Steel Frame", unit: "kg", cost: 180, category: "Metal" },
];

const initialProducts = [
  { id: "p1", name: "Classic 2-Person Spa", category: "Spa", factoryCost: 5000, overheadCost: 2000, shippingCost: 0, profitMargin: 25,
    materials: [{ itemId: "i1", qty: 4 },{ itemId: "i2", qty: 8 },{ itemId: "i3", qty: 1 },{ itemId: "i4", qty: 12 },{ itemId: "i5", qty: 10 },{ itemId: "i6", qty: 1 },{ itemId: "i7", qty: 1 },{ itemId: "i8", qty: 5 },{ itemId: "i9", qty: 3 },{ itemId: "i10", qty: 15 }] },
  { id: "p2", name: "Luxury 6-Person Jacuzzi", category: "Jacuzzi", factoryCost: 12000, overheadCost: 5000, shippingCost: 0, profitMargin: 30,
    materials: [{ itemId: "i1", qty: 8 },{ itemId: "i2", qty: 20 },{ itemId: "i3", qty: 2 },{ itemId: "i4", qty: 32 },{ itemId: "i5", qty: 25 },{ itemId: "i6", qty: 2 },{ itemId: "i7", qty: 1 },{ itemId: "i8", qty: 12 },{ itemId: "i9", qty: 8 },{ itemId: "i10", qty: 40 }] },
];

// ── Helpers ──────────────────────────────────────────────────────────────────
function calcMaterialCost(p, items) { return p.materials.reduce((s,m) => { const i=items.find(x=>x.id===m.itemId); return s+(i?i.cost*m.qty:0); },0); }
function calcLandingCost(p, items) { return calcMaterialCost(p,items)+(p.factoryCost||0)+(p.overheadCost||0)+(p.shippingCost||0); }
function calcSellingPrice(p, items) { return calcLandingCost(p,items)*(1+(p.profitMargin||0)/100); }
function uid() { return "id"+Math.random().toString(36).substr(2,9); }
function fmt(n) { return n.toLocaleString("en-IN",{maximumFractionDigits:0}); }

function applySortKey(arr, sortKey, getters) {
  if (!sortKey) return arr;
  return [...arr].sort((a,b)=>{
    for (const [key,getter] of Object.entries(getters)) {
      if (sortKey===key+"-az"||sortKey===key+"-asc") return getter(a)>getter(b)?1:getter(a)<getter(b)?-1:0;
      if (sortKey===key+"-za"||sortKey===key+"-desc") return getter(a)<getter(b)?1:getter(a)>getter(b)?-1:0;
    }
    return 0;
  });
}

// ── SheetJS loader (shared by import & export) ───────────────────────────────
async function loadSheetJS() {
  const XLSX = await import("xlsx");
  return XLSX;
}

// ── Shared print utility ─────────────────────────────────────────────────────
const PRINT_CSS = `
  *{box-sizing:border-box;margin:0;padding:0}
  body{font-family:'Segoe UI',Arial,sans-serif;font-size:12px;color:#111;background:#fff;padding:20px 28px}
  .ph{display:flex;justify-content:space-between;align-items:flex-start;border-bottom:2px solid #111;padding-bottom:10px;margin-bottom:14px}
  .co{font-size:17px;font-weight:700} .cs{font-size:10px;color:#666;margin-top:2px}
  .dt{font-size:14px;font-weight:700;text-align:right} .dd{font-size:10px;color:#666;text-align:right;margin-top:2px}
  table{width:100%;border-collapse:collapse;margin-bottom:14px}
  th{background:#111;color:#fff;padding:6px 9px;font-size:10px;text-transform:uppercase;letter-spacing:.05em;text-align:left;font-weight:600}
  td{padding:6px 9px;border-bottom:1px solid #e5e5e5;font-size:12px}
  tr:nth-child(even) td{background:#f9f9f9}
  .r{text-align:right} .b{font-weight:700}
  .tr td{background:#efefef!important;font-weight:700;border-top:2px solid #111}
  .sb{display:flex;gap:10px;flex-wrap:wrap;margin-bottom:14px}
  .s{flex:1;min-width:110px;border:1px solid #ccc;border-radius:4px;padding:9px 11px}
  .sl{font-size:9px;text-transform:uppercase;color:#888;font-weight:700;letter-spacing:.05em;margin-bottom:3px}
  .sv{font-size:15px;font-weight:700}
  .sd{background:#f0f0f0;border-color:#999}
  .ft{margin-top:18px;border-top:1px solid #ddd;padding-top:7px;font-size:10px;color:#aaa;display:flex;justify-content:space-between}
  .stl{font-size:11px;font-weight:700;margin:10px 0 5px;text-transform:uppercase;letter-spacing:.05em;border-left:3px solid #111;padding-left:7px}
  @media print{body{padding:0}@page{margin:12mm 10mm;size:A4}}
`;

function printPage(title, bodyHtml) {
  const date = new Date().toLocaleDateString("en-IN",{day:"2-digit",month:"short",year:"numeric"});
  const html = `<!DOCTYPE html><html><head><meta charset="UTF-8"/><title>${title}</title><style>${PRINT_CSS}</style></head><body>
  <div class="ph"><div><div class="co">SpaTrack Pro</div><div class="cs">Manufacturing Cost Manager</div></div><div><div class="dt">${title}</div><div class="dd">Printed: ${date}</div></div></div>
  ${bodyHtml}
  <div class="ft"><span>SpaTrack Pro — Manufacturing Cost Manager</span><span>${date}</span></div>
  <script>window.onload=()=>window.print()<\/script></body></html>`;
  const w=window.open("","_blank"); w.document.write(html); w.document.close();
}

// ── PrintBtn ─────────────────────────────────────────────────────────────────
function PrintBtn({ onClick, label = <><Printer size={16} style={{flexShrink:0}} /> Print</> }) {
  return (
    <button onClick={onClick} style={{display:"flex",alignItems:"center",gap:6,background:"#1e2130",border:"1px solid #2e3350",color:"#e8eaf0",padding:"9px 16px",borderRadius:8,fontSize:13,fontWeight:600,fontFamily:"inherit",cursor:"pointer"}}>
      {label}
    </button>
  );
}

// ── SortTh ───────────────────────────────────────────────────────────────────
function SortTh({ label, field, sortKey, onSort }) {
  const isActive = sortKey===field+"-az"||sortKey===field+"-za"||sortKey===field+"-asc"||sortKey===field+"-desc";
  const isAsc = sortKey===field+"-az"||sortKey===field+"-asc";
  const toggle = () => onSort(isAsc ? field+(field==="name"||field==="category"||field==="unit"?"-za":"-desc") : field+(field==="name"||field==="category"||field==="unit"?"-az":"-asc"));
  return <th onClick={toggle} style={{cursor:"pointer",userSelect:"none",color:isActive?"#5d7cff":"#6b7280",whiteSpace:"nowrap"}}>{label} <span style={{fontSize:10}}>{isActive?(isAsc?<ChevronUp size={12} style={{flexShrink:0}} />:<ChevronDown size={12} style={{flexShrink:0}} />):<ChevronsUpDown size={12} style={{flexShrink:0}} />}</span></th>;
}

// ── CategoryFilter ───────────────────────────────────────────────────────────
function CategoryFilter({ label, options, selected, onToggle, onSelectAll, onClear }) {
  const [open,setOpen]=useState(false); const ref=useRef();
  useEffect(()=>{ const h=e=>{if(ref.current&&!ref.current.contains(e.target))setOpen(false)}; document.addEventListener("mousedown",h); return()=>document.removeEventListener("mousedown",h); },[]);
  const isActive=selected.length>0;
  return (
    <div ref={ref} style={{position:"relative",display:"inline-block"}}>
      <button onClick={()=>setOpen(!open)} style={{display:"flex",alignItems:"center",gap:6,padding:"9px 14px",background:isActive?"linear-gradient(135deg,#5d7cff,#8b5cf6)":"#1e2130",border:isActive?"none":"1px solid #2e3350",color:isActive?"#fff":"#9ba3c0",borderRadius:8,fontSize:13,fontFamily:"inherit",cursor:"pointer",fontWeight:isActive?700:400}}>
        <Zap size={16} style={{flexShrink:0}} /> {label}{isActive?` (${selected.length})`:""} <span style={{fontSize:10}}><ChevronDown size={12} style={{flexShrink:0}} /></span>
      </button>
      {open&&(
        <div style={{position:"absolute",top:"calc(100% + 6px)",left:0,zIndex:9999,background:"#1a1d2e",border:"1px solid #2e3350",borderRadius:12,padding:16,minWidth:200,boxShadow:"0 8px 32px rgba(0,0,0,0.6)"}}>
          <div style={{display:"flex",justifyContent:"space-between",marginBottom:10}}>
            <span style={{fontSize:11,color:"#6b7280",fontWeight:600,textTransform:"uppercase"}}>{label}</span>
            <div style={{display:"flex",gap:8}}>
              <button onClick={onSelectAll} style={{fontSize:11,color:"#5d7cff",background:"none",border:"none",cursor:"pointer",padding:0}}>All</button>
              <button onClick={onClear} style={{fontSize:11,color:"#ff4757",background:"none",border:"none",cursor:"pointer",padding:0}}>Clear</button>
            </div>
          </div>
          {options.map(opt=>(
            <label key={opt} style={{display:"flex",alignItems:"center",gap:8,padding:"6px 0",cursor:"pointer",borderBottom:"1px solid #1e2235"}}>
              <input type="checkbox" checked={selected.includes(opt)} onChange={()=>onToggle(opt)} style={{width:14,height:14,accentColor:"#5d7cff",cursor:"pointer"}}/>
              <span style={{fontSize:13,color:"#e8eaf0"}}>{opt}</span>
            </label>
          ))}
        </div>
      )}
    </div>
  );
}

// ── Central API + Auth ───────────────────────────────────────────────────────
function getToken() { return sessionStorage.getItem("spatrack_token"); }
function getUser()  { return sessionStorage.getItem("spatrack_user") || ""; }
function setSession(token, user) { sessionStorage.setItem("spatrack_token", token); sessionStorage.setItem("spatrack_user", user); }
function clearSession() { sessionStorage.removeItem("spatrack_token"); sessionStorage.removeItem("spatrack_user"); }

async function apiLogin(username, password) {
  const r = await fetch("/spa/api/login", { method:"POST", headers:{"Content-Type":"application/json"}, body:JSON.stringify({username,password}) });
  const j = await r.json();
  if(!j.ok) throw new Error(j.error||"Login failed");
  return j;
}
async function apiLoad() {
  const r = await fetch("/spa/api/data", { headers:{"Authorization":"Bearer "+getToken()} });
  if (r.status===401) throw new Error("401");
  if (!r.ok) throw new Error("HTTP "+r.status);
  return (await r.json()).data;
}
async function apiSave(items, products, purchaseOrders=[]) {
  const r = await fetch("/spa/api/data", { method:"POST", headers:{"Content-Type":"application/json","Authorization":"Bearer "+getToken()}, body:JSON.stringify({items,products,purchaseOrders}) });
  const j = await r.json();
  if(!j.ok) throw new Error(j.error||"Save failed");
  return j;
}
async function apiLogout() {
  await fetch("/spa/api/logout",{method:"POST",headers:{"Authorization":"Bearer "+getToken()}}).catch(()=>{});
  clearSession();
}

// ── Login Screen ──────────────────────────────────────────────────────────────
function LoginScreen({ onLogin }) {
  const [username,setUsername]=useState("");
  const [password,setPassword]=useState("");
  const [showPass,setShowPass]=useState(false);
  const [loading,setLoading]=useState(false);
  const [error,setError]=useState("");

  async function handleSubmit(e) {
    e.preventDefault();
    if(!username||!password){setError("Please enter username and password");return;}
    setLoading(true);setError("");
    try {
      const j = await apiLogin(username.trim(), password);
      setSession(j.token, j.username);
      onLogin();
    } catch(err) { setError(err.message); }
    setLoading(false);
  }

  return (
    <div style={{minHeight:"100vh",width:"100%",background:"linear-gradient(45deg, #0d0f1a, #1a1d2e, #0d0f1a)",backgroundSize:"400% 400%",animation:"gradientBG 15s ease infinite",display:"flex",alignItems:"center",justifyContent:"center",fontFamily:"'DM Sans','Segoe UI',sans-serif",padding:24}}>
      <style>{`@keyframes gradientBG { 0% { background-position: 0% 50%; } 50% { background-position: 100% 50%; } 100% { background-position: 0% 50%; } } @import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600;700&family=Space+Grotesk:wght@600;700&display=swap');*{box-sizing:border-box;margin:0;padding:0}html,body,#root{min-width:1280px;width:100%;overflow-x:auto}`}</style>
      <div style={{width:"100%",maxWidth:400}}>
        <div style={{textAlign:"center",marginBottom:36}}>
          <div style={{width:64,height:64,background:"linear-gradient(135deg,#5d7cff,#8b5cf6)",borderRadius:18,display:"flex",alignItems:"center",justifyContent:"center",fontSize:30,margin:"0 auto 14px"}}><Bath size={24} style={{flexShrink:0}} /></div>
          <div style={{fontFamily:"'Space Grotesk'",fontWeight:700,fontSize:22,color:"#fff"}}>SpaTrack Pro</div>
          <div style={{fontSize:13,color:"#4a5070",marginTop:4}}>Manufacturing Cost Manager</div>
        </div>
        <div style={{background:"#161925",border:"1px solid #22273a",borderRadius:20,padding:32}}>
          <div style={{fontFamily:"'Space Grotesk'",fontWeight:700,fontSize:18,color:"#fff",marginBottom:4}}>Sign in</div>
          <div style={{fontSize:13,color:"#4a5070",marginBottom:24}}>Enter your credentials to continue</div>
          <form onSubmit={handleSubmit}>
            <div style={{marginBottom:14}}>
              <label className={error ? "error" : ""} style={{fontSize:11,fontWeight:700,color:error?"#ff4757":"#9ba3c0",textTransform:"uppercase",letterSpacing:".05em",display:"block",marginBottom:6,transition:"color 0.2s"}}>Username</label>
              <input type="text" value={username} onChange={e=>setUsername(e.target.value)} placeholder="Enter username" autoFocus autoComplete="username" disabled={loading}
                className={error ? "error" : ""} style={{width:"100%",background:"#1e2130",border:"1px solid #2e3350",borderLeft:"3px solid #2e3350",color:"#e8eaf0",borderRadius:10,padding:"11px 14px",fontSize:14,outline:"none",fontFamily:"inherit",transition:"all 0.2s ease"}}
                />
            </div>
            <div style={{marginBottom:22}}>
              <label className={error ? "error" : ""} style={{fontSize:11,fontWeight:700,color:error?"#ff4757":"#9ba3c0",textTransform:"uppercase",letterSpacing:".05em",display:"block",marginBottom:6,transition:"color 0.2s"}}>Password</label>
              <div style={{position:"relative"}}>
                <input type={showPass?"text":"password"} value={password} onChange={e=>setPassword(e.target.value)} placeholder="Enter password" autoComplete="current-password" disabled={loading}
                  className={error ? "error" : ""} style={{width:"100%",background:"#1e2130",border:"1px solid #2e3350",borderLeft:"3px solid #2e3350",color:"#e8eaf0",borderRadius:10,padding:"11px 44px 11px 14px",fontSize:14,outline:"none",fontFamily:"inherit",transition:"all 0.2s ease"}}
                  />
                <button type="button" onClick={()=>setShowPass(!showPass)}
                  style={{position:"absolute",right:12,top:"50%",transform:"translateY(-50%)",background:"none",border:"none",color:"#4a5070",cursor:"pointer",fontSize:15}}>
                  {showPass?<EyeOff size={16} style={{flexShrink:0}} />:<Eye size={16} style={{flexShrink:0}} />}
                </button>
              </div>
            </div>
            {error&&<div style={{display:"flex",alignItems:"center",gap:8,background:"#3a1a1a",border:"1px solid #ff4757",borderRadius:8,padding:"9px 14px",marginBottom:16,fontSize:13,color:"#ff4757"}}><AlertTriangle size={16} style={{flexShrink:0}} /> {error}</div>}
            <button type="submit" disabled={loading}
              style={{width:"100%",background:"linear-gradient(135deg,#5d7cff,#8b5cf6)",border:"none",color:"#fff",borderRadius:10,padding:"13px",fontSize:15,fontWeight:700,fontFamily:"inherit",cursor:loading?"not-allowed":"pointer",opacity:loading?0.7:1,display:"flex",alignItems:"center",justifyContent:"center",gap:8}}>
              {loading?<><Loader2 size={16} className="lucide-spin" style={{flexShrink:0}}/> Signing in…</>:"Sign In →"}
            </button>
          </form>
        </div>
        <div style={{textAlign:"center",marginTop:14,fontSize:12,color:"#2e3350"}}>SpaTrack Pro • Secured access</div>
      </div>
    </div>
  );
}


// ── TopNav ───────────────────────────────────────────────────────────────────
function TopNav({ tab, setTab }) {
  const [indicator, setIndicator] = useState({ left: 0, width: 0 });
  const navRef = useRef(null);

  const TAB_ICONS = {
    "Dashboard": LayoutDashboard,
    "Items": Package,
    "Products": Factory,
    "Cost Sheet": FileText,
    "Purchase Order": ShoppingCart
  };

  useEffect(() => {
    if (navRef.current) {
      const activeBtn = navRef.current.querySelector('[data-active="true"]');
      if (activeBtn) {
        setIndicator({
          left: activeBtn.offsetLeft,
          width: activeBtn.offsetWidth
        });
      }
    }
  }, [tab]);

  return (
    <div ref={navRef} style={{ display: "flex", gap: 12, position: "relative" }}>
      {TABS.map(t => {
        const Icon = TAB_ICONS[t];
        const isActive = tab === t;
        return (
          <button
            key={t}
            data-active={isActive}
            onClick={() => setTab(t)}
            style={{
              display: "flex", alignItems: "center", gap: 8,
              padding: "10px 16px", borderRadius: 10, border: "none",
              fontSize: 14, fontWeight: 600, fontFamily: "inherit",
              background: "transparent",
              color: isActive ? "#fff" : "#6b7280",
              cursor: "pointer", whiteSpace: "nowrap",
              transition: "all 0.2s ease"
            }}
            onMouseEnter={e => !isActive && (e.currentTarget.style.background = "rgba(255,255,255,0.06)")}
            onMouseLeave={e => !isActive && (e.currentTarget.style.background = "transparent")}
          >
            {Icon && <Icon size={16} strokeWidth={isActive ? 2.5 : 2} style={{ color: isActive ? "#5d7cff" : "#6b7280", transition: "color 0.2s ease" }} />}
            {t}
          </button>
        );
      })}
      {indicator.width > 0 && (
        <div style={{
          position: "absolute",
          bottom: -16, // aligns with bottom of the header assuming 64px height and flex centering
          left: indicator.left,
          width: indicator.width,
          height: 3,
          background: "linear-gradient(135deg,#5d7cff,#8b5cf6)",
          transition: "all 0.3s cubic-bezier(0.25, 1, 0.5, 1)",
          borderRadius: "3px 3px 0 0",
          boxShadow: "0 -2px 10px rgba(93,124,255,0.4)",
          pointerEvents: "none"
        }} />
      )}
    </div>
  );
}


// ── SaveIndicator ─────────────────────────────────────────────────────────────
function SaveIndicator({ status, lastSaved }) {
  const [visible, setVisible] = useState(true);

  useEffect(() => {
    if (status === "saved" || status === "error") {
      setVisible(true);
      const t = setTimeout(() => setVisible(false), status === "saved" ? 3000 : 8000);
      return () => clearTimeout(t);
    } else {
      setVisible(true);
    }
  }, [status, lastSaved]);

  if (!visible && status === "saved") return null;

  return (
    <div style={{
      position: "fixed", bottom: 32, right: 32, zIndex: 9999,
      background: "#161925", border: "1px solid #2e3350", borderRadius: 12,
      padding: "12px 16px", display: "flex", alignItems: "center", gap: 12,
      boxShadow: "0 8px 32px rgba(0,0,0,0.5)", overflow: "hidden",
      animation: "floatUp 0.3s cubic-bezier(0.2, 0.8, 0.2, 1)",
      opacity: visible ? 1 : 0, transition: "opacity 0.3s ease", pointerEvents: visible ? "auto" : "none"
    }}>
      {status === "loading" && <><Loader2 size={16} className="lucide-spin" style={{color:"#60a5fa"}}/><span style={{fontSize:13,color:"#60a5fa",fontWeight:600}}>Loading data…</span></>}
      {status === "saving" && <><Save size={16} style={{color:"#f59e0b"}}/><span style={{fontSize:13,color:"#f59e0b",fontWeight:600}}>Saving changes…</span>
        <div style={{position:"absolute", bottom:0, left:0, height:2, width:"100%", background:"linear-gradient(90deg, transparent, #f59e0b, transparent)", animation:"slideRight 1s infinite linear"}}/>
      </>}
      {status === "saved" && <><CheckCircle size={16} style={{color:"#4ade80"}}/><span style={{fontSize:13,color:"#4ade80",fontWeight:600}}>{lastSaved ? `Saved ${lastSaved.toLocaleTimeString("en-IN",{hour:"2-digit",minute:"2-digit"})}` : "All data saved"}</span></>}
      {status === "error" && <><AlertTriangle size={16} style={{color:"#ff4757"}}/><span style={{fontSize:13,color:"#ff4757",fontWeight:600}}>Server unreachable</span></>}
    </div>
  );
}

// ── Main App (shown after login) ─────────────────────────────────────────────
function MainApp({ onLogout }) {
  const [tab,setTab]=useState("Dashboard");
  const [items,setItems]=useState(initialItems);
  const [products,setProducts]=useState(initialProducts);
  const [purchaseOrders,setPurchaseOrders]=useState([]);
  const [saveStatus,setSaveStatus]=useState("loading");
  const [lastSaved,setLastSaved]=useState(null);
  const saveTimer=useRef(null);
  const isFirstRender=useRef(true);

  // ── Load data on mount ───────────────────────────────────────────────────
  useEffect(()=>{
    setSaveStatus("loading");
    apiLoad()
      .then(data=>{ setItems(data.items||initialItems); setProducts(data.products||initialProducts); setPurchaseOrders(data.purchaseOrders||[]); setSaveStatus("saved"); })
      .catch(err=>{ if(err.message==="401"){clearSession();onLogout();}else{setSaveStatus("error");} });
  },[]);

  // ── Auto-save on change (debounced 800ms) ────────────────────────────────
  useEffect(()=>{
    if(isFirstRender.current){isFirstRender.current=false;return;}
    if(saveStatus==="loading") return;
    setSaveStatus("saving");
    clearTimeout(saveTimer.current);
    saveTimer.current=setTimeout(()=>{
      apiSave(items,products,purchaseOrders)
        .then(()=>{ setSaveStatus("saved"); setLastSaved(new Date()); })
        .catch(err=>{ if(err.message==="401"){clearSession();onLogout();}else{setSaveStatus("error");} });
    },800);
    return()=>clearTimeout(saveTimer.current);
  },[items,products,purchaseOrders]);

  async function handleLogout(){ await apiLogout(); onLogout(); }

  return (
    <div style={{fontFamily:"'DM Sans','Segoe UI',sans-serif",minHeight:"100vh",background:"#0f1117",color:"#e8eaf0",minWidth:1280}}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@300;400;500;600;700&family=Space+Grotesk:wght@400;500;600;700&display=swap');
        *{box-sizing:border-box;margin:0;padding:0}
        html,body,#root{min-width:1280px;width:100%;overflow-x:auto}
        ::-webkit-scrollbar{width:6px}::-webkit-scrollbar-track{background:#1a1d27}::-webkit-scrollbar-thumb{background:#3a3f5c;border-radius:3px}
        input,select{background:#1e2130;border:1px solid #2e3350;border-left:3px solid #2e3350;color:#e8eaf0;border-radius:8px;padding:10px 14px;font-family:inherit;font-size:14px;outline:none;width:100%;transition:all 0.2s ease}
        input:focus,select:focus{border-color:#5d7cff;border-left:3px solid #5d7cff}
        input.error,select.error{border-color:#ff4757;border-left:3px solid #ff4757}
        label.error{color:#ff4757 !important}
        input::placeholder{color:#4a5070}
        button{cursor:pointer;font-family:inherit;transition:all 0.15s ease}
        button:disabled{opacity:0.5;cursor:not-allowed}
        .modal-close{transition:all 0.15s ease}
        .modal-close:hover{color:#ff4757 !important;background:rgba(255,71,87,0.1) !important}
        .btn-primary{background:linear-gradient(135deg,#5d7cff,#8b5cf6);border:none;color:white;padding:10px 20px;border-radius:10px;font-size:14px;font-weight:600}
        .btn-primary:hover:not(:disabled){transform:translateY(-1px);box-shadow:0 4px 16px rgba(93,124,255,.35)}
        .btn-danger{background:#ff4757;border:none;color:white;padding:7px 14px;border-radius:8px;font-size:12px;font-weight:600}
        .btn-danger:hover:not(:disabled){background:#ff4757;transform:translateY(-1px);box-shadow:0 4px 16px rgba(255,71,87,.3)}
        .btn-ghost{background:transparent;border:1px solid #2e3350;color:#9ba3c0;padding:8px 16px;border-radius:8px;font-size:13px}
        .btn-ghost:hover{border-color:#5d7cff;color:#5d7cff}
        .card{background:#161925;border:1px solid #22273a;border-radius:16px;padding:24px;box-shadow:0 1px 3px rgba(0,0,0,.4), 0 4px 16px rgba(0,0,0,.2)}
        .tag{display:inline-block;padding:3px 10px;border-radius:20px;font-size:11px;font-weight:600;border:1px solid color-mix(in srgb, currentColor 15%, transparent)}
        .tag-elec{background:#1e3a5f;color:#60a5fa}.tag-raw{background:#1a3a2a;color:#4ade80}.tag-chem{background:#3a1a3a;color:#d946ef}
        .tag-fit{background:#3a2a1a;color:#fb923c}.tag-plum{background:#1a2a3a;color:#38bdf8}.tag-met{background:#2a2a1a;color:#fbbf24}.tag-oth{background:#2a1a1a;color:#f87171}
        .table-wrapper{border-radius:12px;overflow:hidden;border:1px solid #1e2235}
        table{width:100%;border-collapse:collapse}
        th{text-align:left;padding:12px 16px;font-size:11px;font-weight:600;color:#6b7280;text-transform:uppercase;letter-spacing:.08em;border-bottom:1px solid #1e2235;white-space:nowrap}
        th:hover{color:#9ba3c0}
        td{padding:12px 16px;font-size:14px;border-bottom:1px solid #181d2e;transition:background 0.15s ease}
        tr:last-child td{border-bottom:none} tr:hover td{background:#1c2135}
        .modal-overlay{position:fixed;inset:0;background:rgba(0,0,0,.7);backdrop-filter:blur(4px);display:flex;align-items:center;justify-content:center;z-index:100;padding:32px}
        @keyframes modalEnter{from{opacity:0;transform:translateY(16px)}to{opacity:1;transform:translateY(0)}}
        .modal{background:#161925;border:1px solid #22273a;border-radius:20px;padding:32px;width:100%;max-width:860px;max-height:92vh;overflow-y:auto;animation:modalEnter 200ms ease-out forwards}
        .form-row{margin-bottom:16px}
        .form-label{font-size:12px;font-weight:600;color:#9ba3c0;margin-bottom:6px;display:block;text-transform:uppercase;letter-spacing:.05em;transition:color 0.2s}
        .grid-2{display:grid;grid-template-columns:1fr 1fr;gap:16px}
        .section-title{font-family:'Space Grotesk',sans-serif;font-size:22px;font-weight:700;color:#fff;margin-bottom:4px}
        .section-sub{font-size:13px;color:#6b7280;margin-bottom:24px}
        .cost-row{display:flex;justify-content:space-between;align-items:center;padding:10px 0;border-bottom:1px solid #1e2235}
        .cost-row:last-child{border-bottom:none}
        .lucide-spin{animation: spin 2s linear infinite}
        @keyframes spin { 100% { transform: rotate(360deg); } }
        @keyframes pulse{0%,100%{opacity:1}50%{opacity:.4}}
        .saving-dot{animation:pulse 1s infinite}
        @keyframes slideRight { from { transform: translateX(-100%); } to { transform: translateX(100%); } }
        @keyframes floatUp { from { opacity: 0, transform: translateY(16px); } to { opacity: 1, transform: translateY(0); } }
      `}</style>

      {/* Header — sticky desktop nav */}
      <div style={{position:"sticky",top:0,zIndex:50,background:"#0d0f1a",borderBottom:"1px solid #1a1d2e",padding:"0 32px",display:"flex",alignItems:"center",justifyContent:"space-between",height:64,boxShadow:"0 2px 12px rgba(0,0,0,.4)"}}>
        <div style={{display:"flex",alignItems:"center",gap:12,flexShrink:0}}>
          <div style={{width:36,height:36,background:"linear-gradient(135deg,#5d7cff,#8b5cf6)",borderRadius:10,display:"flex",alignItems:"center",justifyContent:"center",fontSize:18}}><Bath size={24} style={{flexShrink:0}} /></div>
          <div><div style={{fontFamily:"'Space Grotesk'",fontWeight:700,fontSize:16,color:"#fff",whiteSpace:"nowrap"}}>SpaTrack Pro</div><div style={{fontSize:11,color:"#4a5070",whiteSpace:"nowrap"}}>Manufacturing Cost Manager</div></div>
        </div>
        <TopNav tab={tab} setTab={setTab} />
        <div style={{display:"flex",alignItems:"center",gap:10,flexShrink:0}}>
          {/* User + Logout */}
          <div style={{display:"flex",alignItems:"center",gap:8,marginLeft:8,paddingLeft:12,borderLeft:"1px solid #1e2235"}}>
            <span style={{fontSize:12,color:"#6b7280"}}><User size={16} style={{flexShrink:0}} /> {getUser()}</span>
            <button onClick={handleLogout}
              style={{background:"#1e2130",border:"1px solid #2e3350",color:"#9ba3c0",padding:"6px 12px",borderRadius:8,fontSize:12,fontWeight:600,cursor:"pointer"}}
              onMouseEnter={e=>{e.currentTarget.style.borderColor="#ff4757";e.currentTarget.style.color="#ff4757";}}
              onMouseLeave={e=>{e.currentTarget.style.borderColor="#2e3350";e.currentTarget.style.color="#9ba3c0";}}>
              Sign Out
            </button>
          </div>
        </div>
      </div>

      <div style={{padding:"32px 40px",maxWidth:1500,margin:"0 auto"}}>
        {tab==="Dashboard"&&<Dashboard products={products} items={items} setTab={setTab}/>}
        {tab==="Items"&&<ItemsPage items={items} setItems={setItems}/>}
        {tab==="Products"&&<ProductsPage products={products} setProducts={setProducts} items={items}/>}
        {tab==="Cost Sheet"&&<CostSheet products={products} items={items}/>}
        {tab==="Purchase Order"&&<PurchaseOrder products={products} items={items} purchaseOrders={purchaseOrders} setPurchaseOrders={setPurchaseOrders}/>}
      </div>
      <SaveIndicator status={saveStatus} lastSaved={lastSaved} />
    </div>
  );
}

// ── App root — no login, auth handled by portal ──────────────────────────────
export default function App() {
  const [token, setToken] = useState(getToken());

  if (!token) {
    return <LoginScreen onLogin={() => setToken(getToken())} />;
  }

  return <MainApp onLogout={() => { clearSession(); setToken(null); }} />;
}

// ── Dashboard ────────────────────────────────────────────────────────────────
function Dashboard({ products, items, setTab }) {
  const [sortKey,setSortKey]=useState("");
  const sorted=applySortKey(products,sortKey,{name:p=>p.name,material:p=>calcMaterialCost(p,items),landing:p=>calcLandingCost(p,items),selling:p=>calcSellingPrice(p,items),profit:p=>p.profitMargin});

  function handlePrint() {
    const rows = sorted.map((p,i)=>{
      const mat=calcMaterialCost(p,items),landing=calcLandingCost(p,items),selling=calcSellingPrice(p,items);
      return `<tr><td>${i+1}</td><td class="b">${p.name}</td><td>${p.category}</td><td>₹${fmt(mat)}</td><td>₹${fmt((p.factoryCost||0)+(p.overheadCost||0)+(p.shippingCost||0))}</td><td class="b">₹${fmt(landing)}</td><td class="b">₹${fmt(selling)}</td><td>${p.profitMargin}%</td></tr>`;
    }).join("");
    const tl=products.reduce((s,p)=>s+calcLandingCost(p,items),0);
    const ts=products.reduce((s,p)=>s+calcSellingPrice(p,items),0);
    printPage("Dashboard — All Products Summary",`
      <div class="sb">
        <div class="s"><div class="sl">Total Products</div><div class="sv">${products.length}</div></div>
        <div class="s"><div class="sl">Total Items</div><div class="sv">${items.length}</div></div>
        <div class="s sd"><div class="sl">Total Landing Cost</div><div class="sv">₹${fmt(tl)}</div></div>
        <div class="s sd"><div class="sl">Total Selling Value</div><div class="sv">₹${fmt(ts)}</div></div>
      </div>
      <table>
        <thead><tr><th>#</th><th>Product</th><th>Category</th><th>Material</th><th>Factory+OH</th><th>Landing Cost</th><th>Selling Price</th><th>Profit%</th></tr></thead>
        <tbody>${rows}<tr class="tr"><td colspan="5" class="b">TOTAL</td><td class="b">₹${fmt(tl)}</td><td class="b">₹${fmt(ts)}</td><td></td></tr></tbody>
      </table>`);
  }

  return (
    <div>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:4}}>
        <div className="section-title">Manufacturing Dashboard</div>
        <PrintBtn onClick={handlePrint} label={<><Printer size={16} style={{flexShrink:0}} /> Print Dashboard</>}/>
      </div>
      <div className="section-sub">Overview of your spa & jacuzzi production costs</div>
      <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:16,marginBottom:24}}>
        {[{label:"Total Products",value:products.length,icon:<Factory size={24} style={{flexShrink:0}} />,color:"#5d7cff"},{label:"Total Items",value:items.length,icon:<Package size={24} style={{flexShrink:0}} />,color:"#8b5cf6"},{label:"Avg Landing Cost",value:"₹"+(products.length?(products.reduce((s,p)=>s+calcLandingCost(p,items),0)/products.length).toLocaleString("en-IN",{maximumFractionDigits:0}):0),icon:<IndianRupee size={24} style={{flexShrink:0}} />,color:"#f59e0b"},{label:"Avg Selling Price",value:"₹"+(products.length?(products.reduce((s,p)=>s+calcSellingPrice(p,items),0)/products.length).toLocaleString("en-IN",{maximumFractionDigits:0}):0),icon:<TrendingUp size={24} style={{flexShrink:0}} />,color:"#10b981"}].map(s=>(
          <div key={s.label} className="card" style={{position:"relative", overflow:"hidden", border:"none"}}>
            <div style={{position:"absolute",top:0,left:0,right:0,height:3,background:`linear-gradient(90deg, ${s.color}, transparent)`}} />
            <div style={{fontSize:28,marginBottom:8}}>{s.icon}</div>
            <div style={{fontSize:26,fontFamily:"'Space Grotesk'",fontWeight:700,color:"#fff"}}>{s.value}</div>
            <div style={{fontSize:12,color:"#6b7280",marginTop:4}}>{s.label}</div>
          </div>
        ))}
      </div>
      <div className="card">
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:20}}>
          <div style={{fontFamily:"'Space Grotesk'",fontWeight:700,fontSize:18,color:"#fff"}}>Products Cost Overview</div>
          <button className="btn-ghost" onClick={()=>setTab("Cost Sheet")}>View Full Cost Sheet →</button>
        </div>
        <table>
          <thead><tr>
            <SortTh label="Product" field="name" sortKey={sortKey} onSort={setSortKey}/>
            <SortTh label="Material Cost" field="material" sortKey={sortKey} onSort={setSortKey}/>
            <th>Factory+OH+Ship</th>
            <SortTh label="Landing Cost" field="landing" sortKey={sortKey} onSort={setSortKey}/>
            <SortTh label="Selling Price" field="selling" sortKey={sortKey} onSort={setSortKey}/>
            <SortTh label="Profit %" field="profit" sortKey={sortKey} onSort={setSortKey}/>
          </tr></thead>
          <tbody>
            {sorted.map(p=>{const mat=calcMaterialCost(p,items),landing=calcLandingCost(p,items),selling=calcSellingPrice(p,items);return(
              <tr key={p.id}>
                <td><div style={{fontWeight:600,color:"#fff"}}>{p.name}</div><div style={{fontSize:11,color:"#4a5070"}}>{p.category}</div></td>
                <td>₹{fmt(mat)}</td><td>₹{fmt((p.factoryCost||0)+(p.overheadCost||0)+(p.shippingCost||0))}</td>
                <td style={{fontWeight:700,color:"#f59e0b"}}>₹{fmt(landing)}</td>
                <td style={{fontWeight:700,color:"#10b981"}}>₹{fmt(selling)}</td>
                <td><span style={{background:"#1a3a2a",color:"#4ade80",padding:"3px 10px",borderRadius:20,fontSize:12,fontWeight:700}}>{p.profitMargin}%</span></td>
              </tr>
            )})}
          </tbody>
        </table>
        {products.length===0&&(
          <div style={{textAlign:"center",padding:"40px 20px",color:"#4a5070",display:"flex",flexDirection:"column",alignItems:"center",gap:16}}>
            <div style={{width:64,height:64,borderRadius:"50%",background:"rgba(139,92,246,0.1)",display:"flex",alignItems:"center",justifyContent:"center",color:"#8b5cf6",marginBottom:8}}>
              <Factory size={32} />
            </div>
            <div>
              <div style={{fontSize:16,fontWeight:700,color:"#e8eaf0",marginBottom:4}}>No products yet</div>
              <div style={{fontSize:14,color:"#6b7280",maxWidth:300,margin:"0 auto"}}>You haven't created any products. Add products to start tracking manufacturing costs.</div>
            </div>
          </div>
        )}
      </div>
    </div>
  );
}

// ── Excel Import Modal ────────────────────────────────────────────────────────
function ExcelImportModal({ onClose, onImport, existingItems }) {
  const [stage, setStage] = useState("drop");
  const [rows, setRows] = useState([]);
  const [errors, setErrors] = useState([]);
  const [mode, setMode] = useState("append");
  const [loading, setLoading] = useState(false);
  const fileRef = useRef();

  const VALID_CATS = ["Raw Material","Chemical","Electrical","Fittings","Plumbing","Metal","Other"];

  async function parseFile(file) {
    setLoading(true);
    try {
      const XLSX = await loadSheetJS();
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf, { type: "array" });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const allRows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
      let headerRowIdx = 0;
      for (let i = 0; i < Math.min(allRows.length, 15); i++) {
        const row = allRows[i];
        const nonEmpty = row.filter(c => String(c).trim() !== "");
        const rowStr = row.map(c => String(c).toLowerCase()).join(" ");
        const hasNameCol = rowStr.includes("item name") || (rowStr.includes("item") && rowStr.includes("category"));
        const hasCostCol = rowStr.includes("cost") || rowStr.includes("price") || rowStr.includes("rate");
        if (nonEmpty.length >= 3 && hasNameCol && hasCostCol) { headerRowIdx = i; break; }
      }
      const raw = XLSX.utils.sheet_to_json(ws, { defval: "", range: headerRowIdx });

      function matchCat(raw) {
        if (!raw) return null;
        const s = raw.toString().trim().toLowerCase();
        for (const c of VALID_CATS) { if (s === c.toLowerCase()) return c; }
        const aliases = {"electrial":"Electrical","electric":"Electrical","electronics":"Electrical","fitting":"Fittings","fitings":"Fittings","fiting":"Fittings","plumb":"Plumbing","pipe":"Plumbing","chemical":"Chemical","chemicals":"Chemical","raw":"Raw Material","rawmaterial":"Raw Material","material":"Raw Material","metal":"Metal","metals":"Metal"};
        for (const [alias, cat] of Object.entries(aliases)) { if (s.includes(alias) || alias.includes(s)) return cat; }
        for (const c of VALID_CATS) { if (c.toLowerCase().startsWith(s.split(" ")[0]) || s.startsWith(c.toLowerCase().split(" ")[0])) return c; }
        return null;
      }

      const parsed = []; const errs = [];
      raw.forEach((row, i) => {
        const norm = {};
        Object.keys(row).forEach(k => { const key = k.toLowerCase().replace(/[₹\(\)\s\-\/\\]/g,"").replace(/[^a-z0-9]/g,""); norm[key] = row[k]; });
        const name = String(norm["itemname"]||norm["name"]||norm["item"]||norm["items"]||"").trim();
        const catRaw = String(norm["category"]||norm["cat"]||norm["type"]||"").trim();
        const unit = String(norm["unit"]||norm["uom"]||norm["measure"]||"pcs").trim()||"pcs";
        const costRaw = norm["unitcost"]||norm["cost"]||norm["price"]||norm["rate"]||norm["unitcostrs"]||norm["costrs"]||norm["unitprice"]||0;
        const cost = parseFloat(String(costRaw).replace(/[^0-9.]/g,""))||0;
        if (!name) { errs.push(`Row ${headerRowIdx+i+2}: Item Name is empty — skipped`); return; }
        const matchedCat = matchCat(catRaw);
        if (!matchedCat) errs.push(`Row ${headerRowIdx+i+2}: "${catRaw}" is not a valid category — assigned "Other"`);
        else if (matchedCat !== catRaw) errs.push(`Row ${headerRowIdx+i+2}: "${catRaw}" auto-corrected to "${matchedCat}"`);
        if (cost <= 0) errs.push(`Row ${headerRowIdx+i+2}: Cost is 0 or missing for "${name}"`);
        const dupInApp = existingItems.find(e => e.name.toLowerCase() === name.toLowerCase());
        parsed.push({ name, unit, cost, category: matchedCat||"Other", _dupInApp:!!dupInApp, _selected:true });
      });
      setRows(parsed); setErrors(errs); setStage("preview");
    } catch(e) { alert("Could not read file: "+e.message); }
    setLoading(false);
  }

  function handleDrop(e) { e.preventDefault(); const f=e.dataTransfer.files[0]; if(f)parseFile(f); }
  function toggleRow(i) { setRows(r=>r.map((x,idx)=>idx===i?{...x,_selected:!x._selected}:x)); }
  function toggleAll(v) { setRows(r=>r.map(x=>({...x,_selected:v}))); }
  function doImport() { const selected=rows.filter(r=>r._selected).map(({name,unit,cost,category})=>({id:uid(),name,unit,cost,category})); onImport(selected,mode); setStage("done"); }

  const selectedCount=rows.filter(r=>r._selected).length;
  const dupCount=rows.filter(r=>r._selected&&r._dupInApp).length;

  return (
    <div className="modal-overlay" onClick={onClose}>
      <div className="modal" onClick={e=>e.stopPropagation()} style={{maxWidth:820}}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:20,borderBottom:"1px solid #1e2235",paddingBottom:16}}>
          <div style={{display:"flex",alignItems:"center",gap:10}}>
            <div style={{display:"flex",alignItems:"center",justifyContent:"center",width:32,height:32,borderRadius:8,background:"linear-gradient(135deg,#5d7cff,#8b5cf6)",color:"#fff"}}><Download size={18} /></div>
            <span style={{fontFamily:"'Space Grotesk'",fontWeight:700,fontSize:20,color:"#fff"}}>Import Items from Excel</span>
          </div>
          <button className="modal-close" onClick={onClose} style={{background:"none",border:"none",color:"#6b7280",cursor:"pointer",display:"flex",alignItems:"center",justifyContent:"center",width:32,height:32,borderRadius:"50%"}}><X size={20} /></button>
        </div>

        <div style={{display:"flex",gap:0,marginBottom:24,background:"#0d0f1a",borderRadius:10,overflow:"hidden",border:"1px solid #1e2235"}}>
          {[["1","Upload File","drop"],["2","Review & Select","preview"],["3","Done","done"]].map(([n,label,s],i)=>(
            <div key={s} style={{flex:1,padding:"10px 0",textAlign:"center",background:stage===s?"linear-gradient(135deg,#5d7cff,#8b5cf6)":i<["drop","preview","done"].indexOf(stage)?"#1a3a2a":"transparent",borderRight:i<2?"1px solid #1e2235":"none"}}>
              <div style={{fontSize:11,fontWeight:700,color:stage===s?"#fff":i<["drop","preview","done"].indexOf(stage)?"#4ade80":"#4a5070",textTransform:"uppercase",letterSpacing:".05em"}}>Step {n}</div>
              <div style={{fontSize:12,color:stage===s?"#fff":i<["drop","preview","done"].indexOf(stage)?"#4ade80":"#6b7280",marginTop:2}}>{label}</div>
            </div>
          ))}
        </div>

        {stage==="drop"&&(
          <div>
            <div onDrop={handleDrop} onDragOver={e=>e.preventDefault()} onClick={()=>fileRef.current.click()}
              style={{border:"2px dashed #2e3350",borderRadius:14,padding:"44px 20px",textAlign:"center",cursor:"pointer",background:"#0d0f1a"}}
              onMouseEnter={e=>e.currentTarget.style.borderColor="#5d7cff"}
              onMouseLeave={e=>e.currentTarget.style.borderColor="#2e3350"}
            >
              <div style={{fontSize:40,marginBottom:10}}><FolderOpen size={40} style={{flexShrink:0}} /></div>
              <div style={{fontWeight:700,color:"#fff",fontSize:15,marginBottom:6}}>Drop your Excel file here</div>
              <div style={{color:"#6b7280",fontSize:13,marginBottom:16}}>or click to browse — supports .xlsx and .xls</div>
              <div style={{display:"inline-block",background:"linear-gradient(135deg,#5d7cff,#8b5cf6)",color:"#fff",padding:"8px 20px",borderRadius:8,fontWeight:600,fontSize:13}}>{loading?"Reading…":"Choose File"}</div>
              <input ref={fileRef} type="file" accept=".xlsx,.xls" style={{display:"none"}} onChange={e=>e.target.files[0]&&parseFile(e.target.files[0])}/>
            </div>
            <div style={{marginTop:20,background:"#0d0f1a",borderRadius:12,padding:16,border:"1px solid #1e2235"}}>
              <div style={{fontWeight:700,color:"#9ba3c0",fontSize:12,textTransform:"uppercase",letterSpacing:".05em",marginBottom:10}}>Required Column Headers</div>
              <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8}}>
                {[["Item Name","Name of the item — required"],["Category","Raw Material / Chemical / Electrical / Fittings / Plumbing / Metal / Other"],["Unit","pcs / kg / mtr / litre / set etc."],["Unit Cost (₹)","Number — cost per unit"]].map(([col,desc])=>(
                  <div key={col} style={{display:"flex",gap:8,alignItems:"flex-start"}}>
                    <span style={{background:"#1e2a4a",color:"#60a5fa",padding:"2px 8px",borderRadius:6,fontSize:11,fontWeight:700,flexShrink:0,fontFamily:"monospace"}}>{col}</span>
                    <span style={{fontSize:12,color:"#6b7280"}}>{desc}</span>
                  </div>
                ))}
              </div>
            </div>
          </div>
        )}

        {stage==="preview"&&(
          <div>
            <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:10,marginBottom:16}}>
              {[{label:"Total Rows",val:rows.length,color:"#60a5fa"},{label:"Selected",val:selectedCount,color:"#4ade80"},{label:"Duplicate Names",val:rows.filter(r=>r._dupInApp).length,color:"#f59e0b"},{label:"Warnings",val:errors.length,color:"#f87171"}].map(s=>(
                <div key={s.label} style={{background:"#0d0f1a",border:"1px solid #1e2235",borderRadius:10,padding:"10px 14px",textAlign:"center"}}>
                  <div style={{fontSize:22,fontWeight:800,color:s.color}}>{s.val}</div>
                  <div style={{fontSize:11,color:"#6b7280",marginTop:2}}>{s.label}</div>
                </div>
              ))}
            </div>
            {errors.length>0&&(
              <div style={{background:"#2a1a0a",border:"1px solid #f59e0b",borderRadius:10,padding:"10px 14px",marginBottom:12,maxHeight:80,overflowY:"auto"}}>
                {errors.map((e,i)=><div key={i} style={{fontSize:12,color:"#f59e0b",marginBottom:2}}><AlertTriangle size={16} style={{flexShrink:0}} /> {e}</div>)}
              </div>
            )}
            <div style={{display:"flex",gap:10,marginBottom:12}}>
              {[["append",<><Plus size={16} style={{flexShrink:0}} /> Add to existing items (keep current)</>],["replace",<><RefreshCw size={16} style={{flexShrink:0}} /> Replace ALL items (wipe current)</>]].map(([v,label])=>(
                <button key={v} onClick={()=>setMode(v)} style={{flex:1,padding:"9px 12px",borderRadius:8,fontSize:12,fontWeight:600,cursor:"pointer",background:mode===v?(v==="replace"?"#3a1a1a":"#1a3a2a"):"#0d0f1a",border:`1.5px solid ${mode===v?(v==="replace"?"#ff4757":"#4ade80"):"#1e2235"}`,color:mode===v?(v==="replace"?"#ff4757":"#4ade80"):"#6b7280"}}>{label}</button>
              ))}
            </div>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:6}}>
              <div style={{fontSize:12,color:"#6b7280"}}>{selectedCount} of {rows.length} rows selected</div>
              <div style={{display:"flex",gap:8}}>
                <button onClick={()=>toggleAll(true)} style={{fontSize:11,color:"#5d7cff",background:"none",border:"none",cursor:"pointer"}}>Select All</button>
                <button onClick={()=>toggleAll(false)} style={{fontSize:11,color:"#ff4757",background:"none",border:"none",cursor:"pointer"}}>Deselect All</button>
              </div>
            </div>
            <div style={{maxHeight:280,overflowY:"auto",borderRadius:10,border:"1px solid #1e2235",marginBottom:16}}>
              <table style={{fontSize:12}}>
                <thead>
                  <tr>
                    <th style={{width:36,padding:"8px 10px"}}><Check size={16} style={{flexShrink:0}} /></th>
                    <th style={{padding:"8px 10px"}}>Item Name</th>
                    <th style={{padding:"8px 10px"}}>Category</th>
                    <th style={{padding:"8px 10px"}}>Unit</th>
                    <th style={{padding:"8px 10px",textAlign:"right"}}>Cost (₹)</th>
                    <th style={{padding:"8px 10px"}}>Status</th>
                  </tr>
                </thead>
                <tbody>
                  {rows.map((row,i)=>(
                    <tr key={i} style={{opacity:row._selected?1:0.4}} onClick={()=>toggleRow(i)}>
                      <td style={{padding:"7px 10px",textAlign:"center"}}><input type="checkbox" checked={row._selected} onChange={()=>toggleRow(i)} onClick={e=>e.stopPropagation()} style={{width:14,height:14,accentColor:"#5d7cff",cursor:"pointer"}}/></td>
                      <td style={{padding:"7px 10px",fontWeight:600,color:"#e8eaf0"}}>{row.name}</td>
                      <td style={{padding:"7px 10px"}}><span style={{fontSize:10,padding:"2px 7px",borderRadius:8,background:"#1e2a4a",color:"#60a5fa",fontWeight:600}}>{row.category}</span></td>
                      <td style={{padding:"7px 10px",color:"#9ba3c0"}}>{row.unit}</td>
                      <td style={{padding:"7px 10px",textAlign:"right",fontWeight:700,color:"#f59e0b"}}>₹{row.cost.toLocaleString("en-IN")}</td>
                      <td style={{padding:"7px 10px"}}>
                        {row._dupInApp?<span style={{fontSize:10,padding:"2px 7px",borderRadius:8,background:"#3a2a0a",color:"#f59e0b",fontWeight:600}}><AlertTriangle size={16} style={{flexShrink:0}} /> Duplicate</span>:<span style={{fontSize:10,padding:"2px 7px",borderRadius:8,background:"#1a3a2a",color:"#4ade80",fontWeight:600}}><Check size={16} style={{flexShrink:0}} /> New</span>}
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
            {dupCount>0&&mode==="append"&&(
              <div style={{background:"#2a2a0a",border:"1px solid #f59e0b",borderRadius:8,padding:"8px 14px",marginBottom:12,fontSize:12,color:"#f59e0b"}}>
                <AlertTriangle size={16} style={{flexShrink:0}} /> {dupCount} item(s) have the same name as existing items. They will be added as duplicates.
              </div>
            )}
            <div style={{display:"flex",gap:12}}>
              <button className="btn-primary" onClick={doImport} style={{flex:2,padding:13}}>Import {selectedCount} Item{selectedCount!==1?"s":""} {mode==="replace"?"(Replace All)":"(Add to List)"}</button>
              <button className="btn-ghost" onClick={()=>setStage("drop")} style={{flex:1,padding:13}}>← Back</button>
              <button className="btn-ghost" onClick={onClose} style={{flex:1,padding:13}}>Cancel</button>
            </div>
          </div>
        )}

        {stage==="done"&&(
          <div style={{textAlign:"center",padding:"32px 0"}}>
            <div style={{fontSize:52,marginBottom:16}}><CheckCircle size={52} style={{flexShrink:0}} /></div>
            <div style={{fontWeight:700,fontSize:20,color:"#fff",marginBottom:8}}>Import Successful!</div>
            <div style={{color:"#6b7280",fontSize:14,marginBottom:28}}>{selectedCount} item{selectedCount!==1?"s":""} have been {mode==="replace"?"imported (replaced all)":"added to your items list"}.</div>
            <button className="btn-primary" onClick={onClose} style={{padding:"12px 40px",fontSize:15}}>Done</button>
          </div>
        )}
      </div>
    </div>
  );
}

// ── Items Page ───────────────────────────────────────────────────────────────
function ItemsPage({ items, setItems }) {
  const [showModal,setShowModal]=useState(false);
  const [showImport,setShowImport]=useState(false);
  const [editItem,setEditItem]=useState(null);
  const [search,setSearch]=useState("");
  const [sortKey,setSortKey]=useState("");
  const [selCats,setSelCats]=useState([]);
  const [priceMin,setPriceMin]=useState("");
  const [priceMax,setPriceMax]=useState("");
  const [exportLoading,setExportLoading]=useState(false);
  const [form,setForm]=useState({name:"",unit:"pcs",cost:"",category:"Raw Material"});

  const allCats=[...new Set(items.map(i=>i.category))];
  function toggleCat(cat){setSelCats(p=>p.includes(cat)?p.filter(c=>c!==cat):[...p,cat]);}
  function openAdd(){setForm({name:"",unit:"pcs",cost:"",category:"Raw Material"});setEditItem(null);setShowModal(true);}
  function openEdit(item){setForm({...item});setEditItem(item.id);setShowModal(true);}
  function save(){if(!form.name||!form.cost)return;if(editItem)setItems(items.map(i=>i.id===editItem?{...form,id:editItem,cost:parseFloat(form.cost)}:i));else setItems([...items,{...form,id:uid(),cost:parseFloat(form.cost)}]);setShowModal(false);}
  function del(id){if(window.confirm("Delete this item?"))setItems(items.filter(i=>i.id!==id));}
  function handleImport(newItems, mode) { if(mode==="replace")setItems(newItems); else setItems(prev=>[...prev,...newItems]); }

  const tagClass=cat=>{const m={Electrical:"tag-elec","Raw Material":"tag-raw",Chemical:"tag-chem",Fittings:"tag-fit",Plumbing:"tag-plum",Metal:"tag-met"};return"tag "+(m[cat]||"tag-oth");};

  const filtered=applySortKey(items.filter(i=>{
    const ms=i.name.toLowerCase().includes(search.toLowerCase())||i.category.toLowerCase().includes(search.toLowerCase());
    const mc=selCats.length===0||selCats.includes(i.category);
    const mn=priceMin===""||i.cost>=parseFloat(priceMin);
    const mx=priceMax===""||i.cost<=parseFloat(priceMax);
    return ms&&mc&&mn&&mx;
  }),sortKey,{name:i=>i.name,category:i=>i.category,cost:i=>i.cost});

  // ── Export to Excel ──────────────────────────────────────────────────────
  async function handleExport() {
    setExportLoading(true);
    try {
      const XLSX = await loadSheetJS();

      // Sheet 1: Items data — same columns import expects so export→edit→re-import works perfectly
      const exportData = filtered.map((item, idx) => ({
        "#": idx + 1,
        "Item Name": item.name,
        "Category": item.category,
        "Unit": item.unit,
        "Unit Cost (Rs)": item.cost,
      }));

      const ws = XLSX.utils.json_to_sheet(exportData);

      // Column widths
      ws["!cols"] = [{ wch: 4 }, { wch: 32 }, { wch: 16 }, { wch: 8 }, { wch: 16 }];

      // Freeze top row
      ws["!freeze"] = { xSplit: 0, ySplit: 1 };

      // Sheet 2: Export info
      const date = new Date().toLocaleDateString("en-IN",{day:"2-digit",month:"short",year:"numeric"});
      const infoRows = [
        ["Exported from", "SpaTrack Pro — Manufacturing Cost Manager"],
        ["Export date", date],
        ["Total items exported", filtered.length],
        ["Total items in app", items.length],
      ];
      if (selCats.length > 0) infoRows.push(["Category filter applied", selCats.join(", ")]);
      if (priceMin) infoRows.push(["Min price filter", `Rs ${priceMin}`]);
      if (priceMax) infoRows.push(["Max price filter", `Rs ${priceMax}`]);
      infoRows.push([]);
      infoRows.push(["NOTE: To re-import, keep column headers exactly as-is and save as .xlsx"]);

      const wsInfo = XLSX.utils.aoa_to_sheet(infoRows);
      wsInfo["!cols"] = [{ wch: 28 }, { wch: 44 }];

      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, "Items");
      XLSX.utils.book_append_sheet(wb, wsInfo, "Export Info");

      const isFiltered = filtered.length < items.length;
      const dateStr = new Date().toISOString().slice(0,10);
      XLSX.writeFile(wb, isFiltered ? `SpaTrack_Items_Filtered_${dateStr}.xlsx` : `SpaTrack_Items_${dateStr}.xlsx`);
    } catch(e) {
      alert("Export failed: " + e.message);
    }
    setExportLoading(false);
  }

  function handlePrint() {
    const rows=filtered.map((item,idx)=>`<tr><td>${idx+1}</td><td class="b">${item.name}</td><td>${item.category}</td><td>${item.unit}</td><td class="r b">₹${item.cost.toLocaleString("en-IN")}</td></tr>`).join("");
    const filterNote=selCats.length>0?`Category filter: ${selCats.join(", ")}`:"All categories";
    printPage("Items / Raw Materials List",`
      <div class="sb">
        <div class="s"><div class="sl">Total Items Shown</div><div class="sv">${filtered.length}</div></div>
        <div class="s"><div class="sl">Filter</div><div class="sv" style="font-size:11px;margin-top:2px">${filterNote}</div></div>
        ${priceMin||priceMax?`<div class="s"><div class="sl">Price Range</div><div class="sv" style="font-size:12px">${priceMin?`₹${priceMin}`:""} – ${priceMax?`₹${priceMax}`:""}</div></div>`:""}
      </div>
      <table>
        <thead><tr><th>#</th><th>Item Name</th><th>Category</th><th>Unit</th><th style="text-align:right">Unit Cost (₹)</th></tr></thead>
        <tbody>${rows}</tbody>
      </table>`);
  }

  return (
    <div>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:24}}>
        <div><div className="section-title">Raw Materials & Items</div><div className="section-sub">Manage all your manufacturing components — add once, use everywhere</div></div>
        <div style={{display:"flex",gap:10,flexWrap:"wrap",justifyContent:"flex-end"}}>
          <PrintBtn onClick={handlePrint} label={<><Printer size={16} style={{flexShrink:0}} /> Print Items</>}/>
          {/* ── Export to Excel button ── */}
          <button
            onClick={handleExport}
            disabled={exportLoading || filtered.length === 0}
            style={{display:"flex",alignItems:"center",gap:6,background:"#1a2a3a",border:"1px solid #1e3a5f",color:"#60a5fa",padding:"9px 16px",borderRadius:8,fontSize:13,fontWeight:600,fontFamily:"inherit",cursor:exportLoading||filtered.length===0?"not-allowed":"pointer",opacity:filtered.length===0?0.5:1,transition:"all .15s"}}
            onMouseEnter={e=>{if(!exportLoading&&filtered.length>0)e.currentTarget.style.background="#1e3a5a";}}
            onMouseLeave={e=>e.currentTarget.style.background="#1a2a3a"}
          >
            {exportLoading ? <><Loader2 size={16} className="lucide-spin" style={{flexShrink:0}} /> Exporting…</> : <><Upload size={16} style={{flexShrink:0}} /> Export to Excel{filtered.length < items.length ? ` (${filtered.length})` : ""}</>}
          </button>
          {/* ── Import from Excel button ── */}
          <button onClick={()=>setShowImport(true)} style={{display:"flex",alignItems:"center",gap:6,background:"#1a3a2a",border:"1px solid #1e4a35",color:"#4ade80",padding:"9px 16px",borderRadius:8,fontSize:13,fontWeight:600,fontFamily:"inherit",cursor:"pointer"}}>
            <Download size={18} style={{flexShrink:0}} /> Import from Excel
          </button>
          <button className="btn-primary" onClick={openAdd}>+ Add New Item</button>
        </div>
      </div>

      <div style={{marginBottom:12,display:"flex",gap:10,alignItems:"center",flexWrap:"wrap"}}>
        <input placeholder="Search by name or category..." value={search} onChange={e=>setSearch(e.target.value)} style={{maxWidth:340}}/>
        <CategoryFilter label="Category" options={allCats} selected={selCats} onToggle={toggleCat} onSelectAll={()=>setSelCats([...allCats])} onClear={()=>setSelCats([])}/>
        <div style={{display:"flex",alignItems:"center",gap:6}}>
          <input type="number" placeholder="Min ₹" value={priceMin} onChange={e=>setPriceMin(e.target.value)} style={{width:90,padding:"8px 10px",fontSize:13}}/>
          <span style={{color:"#6b7280",fontSize:12}}>–</span>
          <input type="number" placeholder="Max ₹" value={priceMax} onChange={e=>setPriceMax(e.target.value)} style={{width:90,padding:"8px 10px",fontSize:13}}/>
        </div>
        {(selCats.length>0||priceMin||priceMax)&&<button onClick={()=>{setSelCats([]);setPriceMin("");setPriceMax("");}} style={{background:"#ff4757",border:"none",color:"white",padding:"8px 14px",borderRadius:8,fontSize:12,fontWeight:600}}><X size={16} style={{flexShrink:0}} /> Clear</button>}
      </div>
      <div style={{marginBottom:8,fontSize:12,color:"#6b7280"}}>Showing {filtered.length} of {items.length} items · Click column headers to sort</div>

      <div className="card" style={{padding:0,overflow:"hidden"}}>
        <table>
          <thead><tr>
            <th style={{width:40}}>#</th>
            <SortTh label="Item Name" field="name" sortKey={sortKey} onSort={setSortKey}/>
            <SortTh label="Category" field="category" sortKey={sortKey} onSort={setSortKey}/>
            <th>Unit</th>
            <SortTh label="Unit Cost (₹)" field="cost" sortKey={sortKey} onSort={setSortKey}/>
            <th>Actions</th>
          </tr></thead>
          <tbody>
            {filtered.map((item,idx)=>(
              <tr key={item.id}>
                <td style={{color:"#4a5070",fontSize:12}}>{idx+1}</td>
                <td style={{fontWeight:600,color:"#fff"}}>{item.name}</td>
                <td><span className={tagClass(item.category)}>{item.category}</span></td>
                <td style={{color:"#9ba3c0"}}>{item.unit}</td>
                <td style={{fontWeight:700,color:"#f59e0b"}}>₹{item.cost.toLocaleString("en-IN")}</td>
                <td style={{display:"flex",gap:8}}>
                  <button className="btn-ghost" style={{padding:"5px 12px",fontSize:12}} onClick={()=>openEdit(item)}>Edit</button>
                  <button className="btn-danger" onClick={()=>del(item.id)}>Delete</button>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
        {filtered.length===0&&(
        <div style={{textAlign:"center",padding:"60px 20px",color:"#4a5070",display:"flex",flexDirection:"column",alignItems:"center",gap:16}}>
          <div style={{width:64,height:64,borderRadius:"50%",background:"rgba(93,124,255,0.1)",display:"flex",alignItems:"center",justifyContent:"center",color:"#5d7cff",marginBottom:8}}>
            <Package size={32} />
          </div>
          <div>
            <div style={{fontSize:16,fontWeight:700,color:"#e8eaf0",marginBottom:4}}>No items found</div>
            <div style={{fontSize:14,color:"#6b7280",maxWidth:300,margin:"0 auto"}}>We couldn't find any items matching your criteria. Try adjusting your filters or add a new item.</div>
          </div>
          <button className="btn-primary" onClick={openAdd} style={{marginTop:8}}>+ Add New Item</button>
        </div>
      )}
      </div>

      {showModal&&(
        <div className="modal-overlay" onClick={()=>setShowModal(false)}>
          <div className="modal" onClick={e=>e.stopPropagation()}>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:24,borderBottom:"1px solid #1e2235",paddingBottom:16}}>
            <div style={{display:"flex",alignItems:"center",gap:10}}>
              <div style={{display:"flex",alignItems:"center",justifyContent:"center",width:32,height:32,borderRadius:8,background:"linear-gradient(135deg,#5d7cff,#8b5cf6)",color:"#fff"}}><Package size={18} /></div>
              <span style={{fontFamily:"'Space Grotesk'",fontWeight:700,fontSize:20,color:"#fff"}}>{editItem?"Edit Item":"Add New Item"}</span>
            </div>
            <button className="modal-close" onClick={()=>setShowModal(false)} style={{background:"none",border:"none",color:"#6b7280",cursor:"pointer",display:"flex",alignItems:"center",justifyContent:"center",width:32,height:32,borderRadius:"50%"}}><X size={20} /></button>
          </div>
            <div className="form-row"><label className="form-label">Item Name</label><input placeholder="e.g. Pump Motor 1HP" value={form.name} onChange={e=>setForm({...form,name:e.target.value})}/></div>
            <div className="grid-2">
              <div className="form-row"><label className="form-label">Category</label><select value={form.category} onChange={e=>setForm({...form,category:e.target.value})}>{CATEGORIES_ITEM.map(c=><option key={c}>{c}</option>)}</select></div>
              <div className="form-row"><label className="form-label">Unit</label><input placeholder="pcs / kg / mtr" value={form.unit} onChange={e=>setForm({...form,unit:e.target.value})}/></div>
            </div>
            <div className="form-row"><label className="form-label">Unit Cost (₹)</label><input type="number" placeholder="0.00" value={form.cost} onChange={e=>setForm({...form,cost:e.target.value})}/></div>
            <div style={{display:"flex",gap:12,marginTop:24}}>
              <button className="btn-primary" onClick={save} style={{flex:1,padding:14}}>{editItem?"Update Item":"Add Item"}</button>
              <button className="btn-ghost" onClick={()=>setShowModal(false)} style={{flex:1,padding:14}}>Cancel</button>
            </div>
          </div>
        </div>
      )}
      {showImport&&<ExcelImportModal onClose={()=>setShowImport(false)} onImport={handleImport} existingItems={items}/>}
    </div>
  );
}

// ── BomPicker ─────────────────────────────────────────────────────────────────
function BomPicker({ items, form, setForm, updateMat, removeMat }) {
  const [search, setSearch] = useState("");
  const [activeCat, setActiveCat] = useState("All");
  const [bomSort, setBomSort] = useState("az");
  const searchRef = useRef();

  const cats = ["All", ...new Set(items.map(i => i.category))];
  const catColors = {Electrical:"#60a5fa","Raw Material":"#4ade80",Chemical:"#d946ef",Fittings:"#fb923c",Plumbing:"#38bdf8",Metal:"#fbbf24",Other:"#f87171"};
  const catBg = {Electrical:"#1e3a5f","Raw Material":"#1a3a2a",Chemical:"#3a1a3a",Fittings:"#3a2a1a",Plumbing:"#1a2a3a",Metal:"#2a2a1a",Other:"#2a1a1a"};

  const available = [...items.filter(i => {
    const inBom = form.materials.some(m => m.itemId === i.id);
    const catOk = activeCat === "All" || i.category === activeCat;
    const searchOk = search === "" || i.name.toLowerCase().includes(search.toLowerCase()) || i.category.toLowerCase().includes(search.toLowerCase());
    return catOk && searchOk && !inBom;
  })].sort((a,b) => a.name.localeCompare(b.name));

  function addItem(item) { setForm(f => ({ ...f, materials: [...f.materials, { itemId: item.id, qty: 1 }] })); }

  const rawBomItems = form.materials.map(m => ({ ...m, item: items.find(i => i.id === m.itemId) })).filter(m => m.item);
  const bomItems = bomSort === "none" ? rawBomItems : [...rawBomItems].sort((a, b) => bomSort === "az" ? a.item.name.localeCompare(b.item.name) : b.item.name.localeCompare(a.item.name));
  const totalMat = rawBomItems.reduce((s, m) => s + m.item.cost * m.qty, 0);

  function cycleSort() { setBomSort(s => s === "az" ? "za" : s === "za" ? "none" : "az"); }
  const sortLabel = bomSort === "az" ? "A→Z <ChevronUp size={12} style={{flexShrink:0}} />" : bomSort === "za" ? "Z→A <ChevronDown size={12} style={{flexShrink:0}} />" : "Order <ChevronsUpDown size={12} style={{flexShrink:0}} />";
  const sortColor = bomSort === "none" ? "#4a5070" : "#5d7cff";

  return (
    <div style={{ marginTop: 8 }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 12 }}>
        <div style={{ fontWeight: 700, color: "#fff", fontSize: 15 }}>
          <Package size={24} style={{flexShrink:0}} /> Bill of Materials
          <span style={{ marginLeft: 8, background: "#1e2a4a", color: "#60a5fa", padding: "2px 8px", borderRadius: 10, fontSize: 11, fontWeight: 600 }}>{form.materials.length} items</span>
        </div>
        <div style={{ fontSize: 12, color: "#6b7280" }}>Click any item to add → set quantity on the right</div>
      </div>

      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12, minHeight: 380 }}>
        {/* LEFT: Item Picker */}
        <div style={{ background: "#0d0f1a", borderRadius: 12, border: "1px solid #1e2235", display: "flex", flexDirection: "column", overflow: "hidden" }}>
          <div style={{ padding: "10px 10px 6px" }}>
            <div style={{ position: "relative" }}>
              <span style={{ position: "absolute", left: 10, top: "50%", transform: "translateY(-50%)", color: "#4a5070", fontSize: 14 }}><Search size={16} style={{flexShrink:0}} /></span>
              <input ref={searchRef} placeholder="Search items..." value={search} onChange={e => setSearch(e.target.value)} style={{ paddingLeft: 32, fontSize: 13, background: "#161925", border: "1px solid #2e3350" }} autoFocus={false}/>
              {search && <button onClick={() => setSearch("")} style={{ position: "absolute", right: 8, top: "50%", transform: "translateY(-50%)", background: "none", border: "none", color: "#6b7280", fontSize: 16, cursor: "pointer", padding: "0 4px" }}><X size={16} style={{flexShrink:0}} /></button>}
            </div>
          </div>
          <div style={{ display: "flex", gap: 4, padding: "4px 10px 8px", overflowX: "auto", flexShrink: 0 }}>
            {cats.map(cat => (
              <button key={cat} onClick={() => setActiveCat(cat)} style={{ padding: "4px 10px", borderRadius: 20, border: "none", fontSize: 11, fontWeight: 600, whiteSpace: "nowrap", cursor: "pointer", flexShrink: 0, background: activeCat === cat ? (cat === "All" ? "linear-gradient(135deg,#5d7cff,#8b5cf6)" : catBg[cat]||"#2a1a1a") : "#161925", color: activeCat === cat ? (cat === "All" ? "#fff" : catColors[cat]||"#f87171") : "#6b7280", border: activeCat === cat ? "none" : "1px solid #1e2235" }}>{cat}</button>
            ))}
          </div>
          <div style={{ padding: "0 10px 6px", fontSize: 11, color: "#4a5070" }}>{available.length} available · {form.materials.length} in BOM</div>
          <div style={{ flex: 1, overflowY: "auto", padding: "0 8px 8px" }}>
            {available.length === 0 && <div style={{ textAlign: "center", padding: "30px 16px", color: "#4a5070", fontSize: 13 }}>{search ? `No items match "${search}"` : "All items in this category are already added"}</div>}
            {available.map(item => (
              <div key={item.id} onClick={() => addItem(item)}
                style={{ display: "flex", alignItems: "center", justifyContent: "space-between", padding: "8px 10px", borderRadius: 8, marginBottom: 4, cursor: "pointer", background: "#161925", border: "1px solid #1e2235", transition: "all .12s" }}
                onMouseEnter={e => { e.currentTarget.style.background = "#1c2235"; e.currentTarget.style.borderColor = "#5d7cff"; }}
                onMouseLeave={e => { e.currentTarget.style.background = "#161925"; e.currentTarget.style.borderColor = "#1e2235"; }}
              >
                <div style={{ flex: 1, minWidth: 0 }}>
                  <div style={{ fontWeight: 600, color: "#e8eaf0", fontSize: 13, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{item.name}</div>
                  <div style={{ display: "flex", alignItems: "center", gap: 6, marginTop: 2 }}>
                    <span style={{ fontSize: 10, padding: "1px 6px", borderRadius: 8, background: catBg[item.category]||"#2a1a1a", color: catColors[item.category]||"#f87171", fontWeight: 600 }}>{item.category}</span>
                    <span style={{ fontSize: 11, color: "#6b7280" }}>{item.unit}</span>
                  </div>
                </div>
                <div style={{ display: "flex", alignItems: "center", gap: 8, flexShrink: 0 }}>
                  <span style={{ fontWeight: 700, color: "#f59e0b", fontSize: 13 }}>₹{item.cost.toLocaleString("en-IN")}</span>
                  <span style={{ background: "linear-gradient(135deg,#5d7cff,#8b5cf6)", color: "#fff", borderRadius: 6, width: 24, height: 24, display: "flex", alignItems: "center", justifyContent: "center", fontSize: 16, fontWeight: 700, flexShrink: 0 }}>+</span>
                </div>
              </div>
            ))}
          </div>
        </div>

        {/* RIGHT: BOM List */}
        <div style={{ background: "#0d0f1a", borderRadius: 12, border: "1px solid #1e2235", display: "flex", flexDirection: "column", overflow: "hidden" }}>
          <div style={{ padding: "8px 12px", borderBottom: "1px solid #1e2235", flexShrink: 0, display: "flex", justifyContent: "space-between", alignItems: "center" }}>
            <div style={{ fontSize: 12, fontWeight: 700, color: "#9ba3c0", textTransform: "uppercase", letterSpacing: ".05em" }}>Selected Items</div>
            {rawBomItems.length > 1 && (
              <button onClick={cycleSort} style={{ display: "flex", alignItems: "center", gap: 5, padding: "4px 10px", borderRadius: 6, background: bomSort !== "none" ? "#1e2a4a" : "#161925", border: `1px solid ${bomSort !== "none" ? "#5d7cff" : "#2e3350"}`, color: sortColor, fontSize: 11, fontWeight: 700, cursor: "pointer" }}>
                <Type size={16} style={{flexShrink:0}} /> {sortLabel}
              </button>
            )}
          </div>
          <div style={{ flex: 1, overflowY: "auto", padding: "8px" }}>
            {bomItems.length === 0 && <div style={{ textAlign: "center", padding: "40px 16px", color: "#4a5070" }}><div style={{ fontSize: 28, marginBottom: 8 }}><ArrowLeft size={16} style={{flexShrink:0}} /></div><div style={{ fontSize: 13 }}>Click items on the left to add them here</div></div>}
            {bomItems.map((m) => {
              const origIdx = form.materials.findIndex(x => x.itemId === m.itemId);
              return (
                <div key={m.itemId} style={{ display: "flex", alignItems: "center", gap: 8, padding: "7px 8px", borderRadius: 8, marginBottom: 4, background: "#161925", border: "1px solid #1e2235" }}>
                  <div style={{ flex: 1, minWidth: 0 }}>
                    <div style={{ fontWeight: 600, color: "#e8eaf0", fontSize: 12, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{m.item.name}</div>
                    <div style={{ fontSize: 11, color: "#6b7280", marginTop: 1 }}>₹{m.item.cost.toLocaleString("en-IN")} × {m.qty} = <span style={{ color: "#f59e0b", fontWeight: 700 }}>₹{fmt(m.item.cost * m.qty)}</span></div>
                  </div>
                  <div style={{ display: "flex", alignItems: "center", gap: 0, background: "#1e2235", borderRadius: 8, overflow: "hidden", flexShrink: 0 }}>
                    <button onClick={() => { if(m.qty > 1) updateMat(origIdx, "qty", m.qty - 1); }} style={{ width: 28, height: 32, background: "none", border: "none", color: "#9ba3c0", fontSize: 16, cursor: "pointer", fontWeight: 700 }}>−</button>
                    <input type="number" value={m.qty} onChange={e => updateMat(origIdx, "qty", parseFloat(e.target.value)||1)} style={{ width: 44, height: 32, textAlign: "center", background: "none", border: "none", borderLeft: "1px solid #2e3350", borderRight: "1px solid #2e3350", color: "#fff", fontWeight: 700, fontSize: 13, padding: 0 }}/>
                    <button onClick={() => updateMat(origIdx, "qty", m.qty + 1)} style={{ width: 28, height: 32, background: "none", border: "none", color: "#9ba3c0", fontSize: 16, cursor: "pointer", fontWeight: 700 }}>+</button>
                  </div>
                  <button onClick={() => removeMat(origIdx)} style={{ width: 28, height: 28, background: "#ff4757", border: "none", color: "white", borderRadius: 6, fontSize: 14, cursor: "pointer", display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0 }}><X size={16} style={{flexShrink:0}} /></button>
                </div>
              );
            })}
          </div>
          {rawBomItems.length > 0 && (
            <div style={{ padding: "10px 12px", borderTop: "1px solid #1e2235", display: "flex", justifyContent: "space-between", alignItems: "center", flexShrink: 0 }}>
              <span style={{ fontSize: 12, color: "#6b7280", fontWeight: 600 }}>MATERIAL TOTAL</span>
              <span style={{ fontWeight: 800, color: "#f59e0b", fontSize: 16 }}>₹{fmt(totalMat)}</span>
            </div>
          )}
        </div>
      </div>
    </div>
  );
}

// ── Products Page ─────────────────────────────────────────────────────────────
function ProductsPage({ products, setProducts, items }) {
  const [showModal,setShowModal]=useState(false);
  const [showImport,setShowImport]=useState(false);
  const [editProd,setEditProd]=useState(null);
  const [search,setSearch]=useState("");
  const [sortKey,setSortKey]=useState("");
  const [selCats,setSelCats]=useState([]);
  const [landingMin,setLandingMin]=useState("");
  const [landingMax,setLandingMax]=useState("");
  const [exportLoading,setExportLoading]=useState(false);
  const [form,setForm]=useState({name:"",category:"Spa",factoryCost:"",overheadCost:"",shippingCost:"",profitMargin:20,materials:[]});

  const allCats=[...new Set(products.map(p=>p.category))];
  function toggleCat(cat){setSelCats(p=>p.includes(cat)?p.filter(c=>c!==cat):[...p,cat]);}
  function openAdd(){setForm({name:"",category:"Spa",factoryCost:"",overheadCost:"",shippingCost:"",profitMargin:20,materials:[]});setEditProd(null);setShowModal(true);}
  function openEdit(p){setForm({...p});setEditProd(p.id);setShowModal(true);}
  function del(id){if(window.confirm("Delete this product?"))setProducts(products.filter(p=>p.id!==id));}
  function updateMat(idx,field,val){const m=[...form.materials];m[idx]={...m[idx],[field]:field==="qty"?parseFloat(val)||0:val};setForm(f=>({...f,materials:m}));}
  function removeMat(idx){setForm(f=>({...f,materials:f.materials.filter((_,i)=>i!==idx)}));}
  function save(){if(!form.name)return;const p={...form,id:editProd||uid(),factoryCost:parseFloat(form.factoryCost)||0,overheadCost:parseFloat(form.overheadCost)||0,shippingCost:parseFloat(form.shippingCost)||0,profitMargin:parseFloat(form.profitMargin)||0};if(editProd)setProducts(products.map(x=>x.id===editProd?p:x));else setProducts([...products,p]);setShowModal(false);}

  // ── Export Products to Excel ───────────────────────────────────────────────
  async function handleExport() {
    setExportLoading(true);
    try {
      const XLSX = await loadSheetJS();
      const wb = XLSX.utils.book_new();
      const date = new Date().toLocaleDateString("en-IN",{day:"2-digit",month:"short",year:"numeric"});

      // Sheet 1: Products Summary — one row per product with all cost fields
      const summaryData = products.map((p, idx) => {
        const mat = calcMaterialCost(p, items);
        const landing = calcLandingCost(p, items);
        const selling = calcSellingPrice(p, items);
        return {
          "#": idx + 1,
          "Product Name": p.name,
          "Category": p.category,
          "Factory Cost (Rs)": p.factoryCost || 0,
          "Overhead Cost (Rs)": p.overheadCost || 0,
          "Shipping Cost (Rs)": p.shippingCost || 0,
          "Profit Margin %": p.profitMargin || 0,
          "BOM Items Count": p.materials.length,
          "Material Cost (Rs)": mat,
          "Landing Cost (Rs)": landing,
          "Selling Price (Rs)": selling,
        };
      });
      const wsSummary = XLSX.utils.json_to_sheet(summaryData);
      wsSummary["!cols"] = [{wch:4},{wch:28},{wch:14},{wch:16},{wch:16},{wch:16},{wch:14},{wch:12},{wch:18},{wch:18},{wch:18}];
      XLSX.utils.book_append_sheet(wb, wsSummary, "Products Summary");

      // Sheet 2: BOM Details — all products' BOM in one flat sheet (importable format)
      const bomData = [];
      products.forEach(p => {
        p.materials.forEach(m => {
          const item = items.find(i => i.id === m.itemId);
          bomData.push({
            "Product Name": p.name,
            "Product Category": p.category,
            "Factory Cost (Rs)": p.factoryCost || 0,
            "Overhead Cost (Rs)": p.overheadCost || 0,
            "Shipping Cost (Rs)": p.shippingCost || 0,
            "Profit Margin %": p.profitMargin || 0,
            "Item Name": item ? item.name : m.itemId,
            "Item Category": item ? item.category : "",
            "Unit": item ? item.unit : "",
            "Qty": m.qty,
            "Unit Cost (Rs)": item ? item.cost : 0,
            "Line Total (Rs)": item ? item.cost * m.qty : 0,
          });
        });
        // blank separator between products
        bomData.push({});
      });
      const wsBom = XLSX.utils.json_to_sheet(bomData);
      wsBom["!cols"] = [{wch:28},{wch:14},{wch:16},{wch:16},{wch:16},{wch:14},{wch:28},{wch:14},{wch:8},{wch:6},{wch:14},{wch:14}];
      XLSX.utils.book_append_sheet(wb, wsBom, "BOM Details (Import This)");

      // Sheet 3: Export info
      const wsInfo = XLSX.utils.aoa_to_sheet([
        ["Exported from", "SpaTrack Pro — Manufacturing Cost Manager"],
        ["Export date", date],
        ["Total products", products.length],
        ["Total items in app", items.length],
        [],
        ["HOW TO RESTORE / IMPORT:"],
        ["1. Use the 'BOM Details (Import This)' sheet — do NOT rename it"],
        ["2. Do not delete or rename any columns"],
        ["3. Click 'Import from Excel' on the Products page"],
        ["4. Choose 'Replace All' to fully restore, or 'Add' to merge"],
        [],
        ["IMPORTANT: Item names in BOM must exactly match item names in your Items list"],
        ["If an item is not found, that BOM line will be skipped"],
      ]);
      wsInfo["!cols"] = [{wch:24},{wch:52}];
      XLSX.utils.book_append_sheet(wb, wsInfo, "How To Restore");

      const dateStr = new Date().toISOString().slice(0,10);
      XLSX.writeFile(wb, `SpaTrack_Products_Backup_${dateStr}.xlsx`);
    } catch(e) {
      alert("Export failed: " + e.message);
    }
    setExportLoading(false);
  }

  // ── Import handler (called by modal) ──────────────────────────────────────
  function handleImportProducts(importedProducts, mode) {
    if (mode === "replace") setProducts(importedProducts);
    else setProducts(prev => [...prev, ...importedProducts]);
  }

  const matPreview=form.materials.reduce((s,m)=>{const i=items.find(x=>x.id===m.itemId);return s+(i?i.cost*m.qty:0);},0);
  const landingPreview=matPreview+(parseFloat(form.factoryCost)||0)+(parseFloat(form.overheadCost)||0)+(parseFloat(form.shippingCost)||0);

  const filtered=applySortKey(products.filter(p=>{
    const ms=p.name.toLowerCase().includes(search.toLowerCase());
    const mc=selCats.length===0||selCats.includes(p.category);
    const landing=calcLandingCost(p,items);
    const mn=landingMin===""||landing>=parseFloat(landingMin);
    const mx=landingMax===""||landing<=parseFloat(landingMax);
    return ms&&mc&&mn&&mx;
  }),sortKey,{name:p=>p.name,category:p=>p.category,landing:p=>calcLandingCost(p,items),selling:p=>calcSellingPrice(p,items),profit:p=>p.profitMargin});

  function handlePrint() {
    const rows=filtered.map((p,i)=>{const mat=calcMaterialCost(p,items),landing=calcLandingCost(p,items),selling=calcSellingPrice(p,items);return `<tr><td>${i+1}</td><td class="b">${p.name}</td><td>${p.category}</td><td>${p.materials.length}</td><td>₹${fmt(mat)}</td><td>₹${fmt(p.factoryCost||0)}</td><td>₹${fmt(p.overheadCost||0)}</td><td>₹${fmt(p.shippingCost||0)}</td><td class="b r">₹${fmt(landing)}</td><td class="b r">₹${fmt(selling)}</td><td>${p.profitMargin}%</td></tr>`;}).join("");
    const tl=filtered.reduce((s,p)=>s+calcLandingCost(p,items),0);
    const ts=filtered.reduce((s,p)=>s+calcSellingPrice(p,items),0);
    printPage("Products List — BOM Summary",`
      <div class="sb">
        <div class="s"><div class="sl">Products Shown</div><div class="sv">${filtered.length}</div></div>
        <div class="s sd"><div class="sl">Total Landing Cost</div><div class="sv">₹${fmt(tl)}</div></div>
        <div class="s sd"><div class="sl">Total Selling Value</div><div class="sv">₹${fmt(ts)}</div></div>
      </div>
      <table>
        <thead><tr><th>#</th><th>Product</th><th>Category</th><th>BOM</th><th>Material</th><th>Factory</th><th>Overhead</th><th>Shipping</th><th style="text-align:right">Landing Cost</th><th style="text-align:right">Selling Price</th><th>Margin</th></tr></thead>
        <tbody>${rows}<tr class="tr"><td colspan="8" class="b">TOTAL</td><td class="b r">₹${fmt(tl)}</td><td class="b r">₹${fmt(ts)}</td><td></td></tr></tbody>
      </table>`);
  }

  function handlePrintProduct(p) {
    const mat=calcMaterialCost(p,items),landing=calcLandingCost(p,items),selling=calcSellingPrice(p,items);
    const sortedMats=[...p.materials].map(m=>({...m,item:items.find(x=>x.id===m.itemId)})).filter(m=>m.item).sort((a,b)=>a.item.name.localeCompare(b.item.name));
    const rows=sortedMats.map((m,i)=>`<tr><td>${i+1}</td><td class="b">${m.item.name}</td><td>${m.item.category}</td><td>${m.item.unit}</td><td>${m.qty}</td><td class="r">₹${m.item.cost.toLocaleString("en-IN")}</td><td class="r b">₹${fmt(m.item.cost*m.qty)}</td></tr>`).join("");
    printPage(`Cost Sheet — ${p.name}`,`
      <div class="sb" style="margin-bottom:12px">
        <div class="s"><div class="sl">Product</div><div class="sv" style="font-size:13px">${p.name}</div></div>
        <div class="s"><div class="sl">Category</div><div class="sv" style="font-size:13px">${p.category}</div></div>
        <div class="s"><div class="sl">Profit Margin</div><div class="sv" style="font-size:13px">${p.profitMargin}%</div></div>
        <div class="s"><div class="sl">BOM Items</div><div class="sv" style="font-size:13px">${p.materials.length}</div></div>
      </div>
      <div class="stl">Bill of Materials</div>
      <table>
        <thead><tr><th>#</th><th>Item</th><th>Category</th><th>Unit</th><th>Qty</th><th style="text-align:right">Rate (₹)</th><th style="text-align:right">Amount (₹)</th></tr></thead>
        <tbody>${rows}<tr class="tr"><td colspan="6" class="b">Total Material Cost</td><td class="r b">₹${fmt(mat)}</td></tr></tbody>
      </table>
      <div class="stl">Cost Summary</div>
      <div class="sb">
        <div class="s"><div class="sl">Material Cost</div><div class="sv">₹${fmt(mat)}</div></div>
        <div class="s"><div class="sl">Factory Cost</div><div class="sv">₹${fmt(p.factoryCost||0)}</div></div>
        <div class="s"><div class="sl">Overhead Cost</div><div class="sv">₹${fmt(p.overheadCost||0)}</div></div>
        <div class="s"><div class="sl">Shipping Cost</div><div class="sv">₹${fmt(p.shippingCost||0)}</div></div>
        <div class="s sd"><div class="sl">Landing Cost</div><div class="sv">₹${fmt(landing)}</div></div>
        <div class="s sd"><div class="sl">Selling Price (+${p.profitMargin}%)</div><div class="sv">₹${fmt(selling)}</div></div>
      </div>`);
  }

  return (
    <div>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:24}}>
        <div><div className="section-title">Products (BOM)</div><div className="section-sub">Define products with Bill of Materials — reuse items across multiple products</div></div>
        <div style={{display:"flex",gap:10,flexWrap:"wrap",justifyContent:"flex-end"}}>
          <PrintBtn onClick={handlePrint} label={<><Printer size={16} style={{marginRight: 6}} /> Print All Products</>}/>
          {/* Export Products */}
          <button
            onClick={handleExport}
            disabled={exportLoading || products.length === 0}
            style={{display:"flex",alignItems:"center",gap:6,background:"#1a2a3a",border:"1px solid #1e3a5f",color:"#60a5fa",padding:"9px 16px",borderRadius:8,fontSize:13,fontWeight:600,fontFamily:"inherit",cursor:exportLoading||products.length===0?"not-allowed":"pointer",opacity:products.length===0?0.5:1}}
            onMouseEnter={e=>{if(!exportLoading&&products.length>0)e.currentTarget.style.background="#1e3a5a";}}
            onMouseLeave={e=>e.currentTarget.style.background="#1a2a3a"}
          >
            {exportLoading ? <><Loader2 size={16} className="lucide-spin" style={{flexShrink:0}} /> Exporting…</> : <><Upload size={16} style={{flexShrink:0}} /> Export to Excel</>}
          </button>
          {/* Import Products */}
          <button onClick={()=>setShowImport(true)} style={{display:"flex",alignItems:"center",gap:6,background:"#1a3a2a",border:"1px solid #1e4a35",color:"#4ade80",padding:"9px 16px",borderRadius:8,fontSize:13,fontWeight:600,fontFamily:"inherit",cursor:"pointer"}}>
            <Download size={18} style={{flexShrink:0}} /> Import from Excel
          </button>
          <button className="btn-primary" onClick={openAdd}>+ Add New Product</button>
        </div>
      </div>

      <div style={{marginBottom:12,display:"flex",gap:10,alignItems:"center",flexWrap:"wrap"}}>
        <input placeholder="Search products..." value={search} onChange={e=>setSearch(e.target.value)} style={{maxWidth:300}}/>
        <CategoryFilter label="Category" options={allCats.length>0?allCats:CATEGORIES_PROD} selected={selCats} onToggle={toggleCat} onSelectAll={()=>setSelCats(allCats.length>0?[...allCats]:[...CATEGORIES_PROD])} onClear={()=>setSelCats([])}/>
        <div style={{display:"flex",alignItems:"center",gap:6}}>
          <input type="number" placeholder="Min Landing ₹" value={landingMin} onChange={e=>setLandingMin(e.target.value)} style={{width:120,padding:"8px 10px",fontSize:13}}/>
          <span style={{color:"#6b7280",fontSize:12}}>–</span>
          <input type="number" placeholder="Max Landing ₹" value={landingMax} onChange={e=>setLandingMax(e.target.value)} style={{width:120,padding:"8px 10px",fontSize:13}}/>
        </div>
        {(selCats.length>0||landingMin||landingMax)&&<button onClick={()=>{setSelCats([]);setLandingMin("");setLandingMax("");}} style={{background:"#ff4757",border:"none",color:"white",padding:"8px 14px",borderRadius:8,fontSize:12,fontWeight:600}}><X size={16} style={{flexShrink:0}} /> Clear</button>}
      </div>
      <div style={{marginBottom:12,fontSize:12,color:"#6b7280"}}>Showing {filtered.length} of {products.length} products · Click column headers to sort</div>

      <div className="card" style={{padding:0,overflow:"hidden",marginBottom:20}}>
        <table>
          <thead><tr>
            <SortTh label="Product Name" field="name" sortKey={sortKey} onSort={setSortKey}/>
            <SortTh label="Category" field="category" sortKey={sortKey} onSort={setSortKey}/>
            <th>BOM</th>
            <SortTh label="Landing Cost" field="landing" sortKey={sortKey} onSort={setSortKey}/>
            <SortTh label="Selling Price" field="selling" sortKey={sortKey} onSort={setSortKey}/>
            <SortTh label="Profit %" field="profit" sortKey={sortKey} onSort={setSortKey}/>
            <th>Actions</th>
          </tr></thead>
          <tbody>
            {filtered.map(p=>{const landing=calcLandingCost(p,items),selling=calcSellingPrice(p,items);return(
              <tr key={p.id}>
                <td style={{fontWeight:600,color:"#fff"}}>{p.name}</td>
                <td><span style={{background:"#1e2a4a",color:"#60a5fa",padding:"2px 8px",borderRadius:12,fontSize:11,fontWeight:600}}>{p.category}</span></td>
                <td style={{color:"#6b7280",fontSize:13}}>{p.materials.length} items</td>
                <td style={{fontWeight:700,color:"#f59e0b"}}>₹{fmt(landing)}</td>
                <td style={{fontWeight:700,color:"#10b981"}}>₹{fmt(selling)}</td>
                <td><span style={{background:"#1a3a2a",color:"#4ade80",padding:"3px 10px",borderRadius:20,fontSize:12,fontWeight:700}}>{p.profitMargin}%</span></td>
                <td>
                  <div style={{display:"flex",gap:6}}>
                    <button className="btn-ghost" style={{padding:"5px 10px",fontSize:11}} onClick={()=>openEdit(p)}>Edit</button>
                    <button onClick={()=>handlePrintProduct(p)} style={{background:"#1e2130",border:"1px solid #2e3350",color:"#9ba3c0",padding:"5px 10px",borderRadius:6,fontSize:11,cursor:"pointer"}}><Printer size={16} style={{flexShrink:0}} /></button>
                    <button className="btn-danger" style={{padding:"5px 8px"}} onClick={()=>del(p.id)}><X size={16} style={{flexShrink:0}} /></button>
                  </div>
                </td>
              </tr>
            )})}
          </tbody>
        </table>
        {filtered.length===0&&(
        <div style={{textAlign:"center",padding:"60px 20px",color:"#4a5070",display:"flex",flexDirection:"column",alignItems:"center",gap:16}}>
          <div style={{width:64,height:64,borderRadius:"50%",background:"rgba(139,92,246,0.1)",display:"flex",alignItems:"center",justifyContent:"center",color:"#8b5cf6",marginBottom:8}}>
            <Factory size={32} />
          </div>
          <div>
            <div style={{fontSize:16,fontWeight:700,color:"#e8eaf0",marginBottom:4}}>No products found</div>
            <div style={{fontSize:14,color:"#6b7280",maxWidth:300,margin:"0 auto"}}>We couldn't find any products matching your criteria. Try adjusting your filters or add a new product.</div>
          </div>
          <button className="btn-primary" onClick={openAdd} style={{marginTop:8}}>+ Add New Product</button>
        </div>
      )}
      </div>

      {showModal&&(
        <div className="modal-overlay" onClick={()=>setShowModal(false)}>
          <div className="modal" onClick={e=>e.stopPropagation()} style={{maxWidth:980,padding:"28px 32px"}}>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:20,borderBottom:"1px solid #1e2235",paddingBottom:16}}>
            <div style={{display:"flex",alignItems:"center",gap:10}}>
              <div style={{display:"flex",alignItems:"center",justifyContent:"center",width:32,height:32,borderRadius:8,background:"linear-gradient(135deg,#5d7cff,#8b5cf6)",color:"#fff"}}><Factory size={18} /></div>
              <span style={{fontFamily:"'Space Grotesk'",fontWeight:700,fontSize:20,color:"#fff"}}>{editProd?"Edit Product":"Add New Product"}</span>
            </div>
            <button className="modal-close" onClick={()=>setShowModal(false)} style={{background:"none",border:"none",color:"#6b7280",cursor:"pointer",display:"flex",alignItems:"center",justifyContent:"center",width:32,height:32,borderRadius:"50%"}}><X size={20} /></button>
          </div>
            <div style={{display:"grid",gridTemplateColumns:"2fr 1fr 1fr 1fr 1fr",gap:12,marginBottom:16}}>
              <div className="form-row" style={{marginBottom:0}}><label className="form-label">Product Name</label><input placeholder="e.g. Classic 2-Person Spa" value={form.name} onChange={e=>setForm({...form,name:e.target.value})}/></div>
              <div className="form-row" style={{marginBottom:0}}><label className="form-label">Category</label><select value={form.category} onChange={e=>setForm({...form,category:e.target.value})}>{CATEGORIES_PROD.map(c=><option key={c}>{c}</option>)}</select></div>
              <div className="form-row" style={{marginBottom:0}}><label className="form-label">Factory Cost (₹)</label><input type="number" placeholder="0" value={form.factoryCost} onChange={e=>setForm({...form,factoryCost:e.target.value})}/></div>
              <div className="form-row" style={{marginBottom:0}}><label className="form-label">Overhead Cost (₹)</label><input type="number" placeholder="0" value={form.overheadCost} onChange={e=>setForm({...form,overheadCost:e.target.value})}/></div>
              <div className="form-row" style={{marginBottom:0}}><label className="form-label">Shipping Cost (₹)</label><input type="number" placeholder="0" value={form.shippingCost} onChange={e=>setForm({...form,shippingCost:e.target.value})}/></div>
            </div>
            <div style={{display:"grid",gridTemplateColumns:"1fr 3fr",gap:12,marginBottom:16}}>
              <div className="form-row" style={{marginBottom:0}}><label className="form-label">Profit Margin %</label><input type="number" placeholder="20" value={form.profitMargin} onChange={e=>setForm({...form,profitMargin:e.target.value})}/></div>
              <div style={{background:"#0d0f1a",borderRadius:10,padding:"8px 16px",display:"grid",gridTemplateColumns:"1fr 1fr 1fr",alignItems:"center",border:"1px solid #1e2235"}}>
                <div style={{textAlign:"center"}}><div style={{fontSize:10,color:"#6b7280",fontWeight:600,textTransform:"uppercase",marginBottom:2}}>Material</div><div style={{fontWeight:700,color:"#9ba3c0",fontSize:15}}>₹{fmt(matPreview)}</div></div>
                <div style={{textAlign:"center",borderLeft:"1px solid #1e2235",borderRight:"1px solid #1e2235"}}><div style={{fontSize:10,color:"#6b7280",fontWeight:600,textTransform:"uppercase",marginBottom:2}}>Landing Cost</div><div style={{fontWeight:800,color:"#f59e0b",fontSize:18}}>₹{fmt(landingPreview)}</div></div>
                <div style={{textAlign:"center"}}><div style={{fontSize:10,color:"#6b7280",fontWeight:600,textTransform:"uppercase",marginBottom:2}}>Selling Price</div><div style={{fontWeight:700,color:"#10b981",fontSize:15}}>₹{fmt(landingPreview*(1+(parseFloat(form.profitMargin)||0)/100))}</div></div>
              </div>
            </div>
            <BomPicker items={items} form={form} setForm={setForm} updateMat={updateMat} removeMat={removeMat}/>
            <div style={{display:"flex",gap:12,marginTop:20}}>
              <button className="btn-primary" onClick={save} style={{flex:1,padding:14}}>{editProd?"Update Product":"Add Product"}</button>
              <button className="btn-ghost" onClick={()=>setShowModal(false)} style={{flex:1,padding:14}}>Cancel</button>
            </div>
          </div>
        </div>
      )}
      {showImport&&<ProductsImportModal onClose={()=>setShowImport(false)} onImport={handleImportProducts} existingProducts={products} items={items}/>}
    </div>
  );
}

// ── Products Import Modal ─────────────────────────────────────────────────────
function ProductsImportModal({ onClose, onImport, existingProducts, items }) {
  const [stage, setStage] = useState("drop"); // drop | preview | done
  const [rows, setRows] = useState([]); // parsed products ready to import
  const [errors, setErrors] = useState([]);
  const [mode, setMode] = useState("append"); // append | replace
  const [loading, setLoading] = useState(false);
  const fileRef = useRef();

  async function parseFile(file) {
    setLoading(true);
    try {
      const XLSX = await loadSheetJS();
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf, { type: "array" });

      // Find the BOM Details sheet (preferred) or fall back to first sheet
      const sheetName = wb.SheetNames.find(n => n.toLowerCase().includes("bom") || n.toLowerCase().includes("detail")) || wb.SheetNames[0];
      const ws = wb.Sheets[sheetName];
      const raw = XLSX.utils.sheet_to_json(ws, { defval: "" });

      const errs = [];
      // Group rows by Product Name to rebuild products + their materials
      const productMap = new Map(); // productName -> product object

      raw.forEach((row, i) => {
        // Normalise column keys
        const norm = {};
        Object.keys(row).forEach(k => { norm[k.toLowerCase().replace(/[^a-z0-9]/g,"")] = row[k]; });

        const productName = String(norm["productname"] || norm["product"] || norm["name"] || "").trim();
        if (!productName) return; // blank separator row

        const itemName = String(norm["itemname"] || norm["item"] || "").trim();
        const qty = parseFloat(String(norm["qty"] || norm["quantity"] || 1).replace(/[^0-9.]/g,"")) || 1;

        if (!productMap.has(productName)) {
          // Parse product-level fields from this row
          const cat = String(norm["productcategory"] || norm["category"] || "Spa").trim();
          const factoryCost = parseFloat(String(norm["factorycostrs"] || norm["factorycost"] || 0).replace(/[^0-9.]/g,"")) || 0;
          const overheadCost = parseFloat(String(norm["overheadcostrs"] || norm["overheadcost"] || 0).replace(/[^0-9.]/g,"")) || 0;
          const shippingCost = parseFloat(String(norm["shippingcostrs"] || norm["shippingcost"] || 0).replace(/[^0-9.]/g,"")) || 0;
          const profitMargin = parseFloat(String(norm["profitmargin"] || norm["profit"] || 20).replace(/[^0-9.]/g,"")) || 0;
          const dupInApp = existingProducts.find(p => p.name.toLowerCase() === productName.toLowerCase());
          productMap.set(productName, {
            name: productName, category: cat,
            factoryCost, overheadCost, shippingCost, profitMargin,
            materials: [], _dupInApp: !!dupInApp, _selected: true,
          });
        }

        // Add BOM line if item name given
        if (itemName) {
          const matchedItem = items.find(i => i.name.toLowerCase() === itemName.toLowerCase());
          if (!matchedItem) {
            errs.push(`Row ${i+2}: Item "${itemName}" not found in your Items list — skipped`);
          } else {
            const prod = productMap.get(productName);
            const alreadyAdded = prod.materials.some(m => m.itemId === matchedItem.id);
            if (!alreadyAdded) prod.materials.push({ itemId: matchedItem.id, qty });
          }
        }
      });

      const parsed = [...productMap.values()];
      if (parsed.length === 0) errs.push("No products found. Make sure you're using the 'BOM Details (Import This)' sheet from the export file.");

      setRows(parsed); setErrors(errs); setStage("preview");
    } catch(e) { alert("Could not read file: " + e.message); }
    setLoading(false);
  }

  function handleDrop(e) { e.preventDefault(); const f=e.dataTransfer.files[0]; if(f)parseFile(f); }
  function toggleRow(i) { setRows(r=>r.map((x,idx)=>idx===i?{...x,_selected:!x._selected}:x)); }
  function toggleAll(v) { setRows(r=>r.map(x=>({...x,_selected:v}))); }
  function doImport() {
    const selected = rows.filter(r=>r._selected).map(({name,category,factoryCost,overheadCost,shippingCost,profitMargin,materials}) => ({
      id: uid(), name, category, factoryCost, overheadCost, shippingCost, profitMargin, materials
    }));
    onImport(selected, mode);
    setStage("done");
  }

  const selectedCount = rows.filter(r=>r._selected).length;
  const dupCount = rows.filter(r=>r._selected && r._dupInApp).length;
  const totalBomLines = rows.filter(r=>r._selected).reduce((s,r)=>s+r.materials.length,0);

  return (
    <div className="modal-overlay" onClick={onClose}>
      <div className="modal" onClick={e=>e.stopPropagation()} style={{maxWidth:820}}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:20,borderBottom:"1px solid #1e2235",paddingBottom:16}}>
          <div style={{display:"flex",alignItems:"center",gap:10}}>
            <div style={{display:"flex",alignItems:"center",justifyContent:"center",width:32,height:32,borderRadius:8,background:"linear-gradient(135deg,#5d7cff,#8b5cf6)",color:"#fff"}}><Download size={18} /></div>
            <span style={{fontFamily:"'Space Grotesk'",fontWeight:700,fontSize:20,color:"#fff"}}>Import Products from Excel</span>
          </div>
          <button className="modal-close" onClick={onClose} style={{background:"none",border:"none",color:"#6b7280",cursor:"pointer",display:"flex",alignItems:"center",justifyContent:"center",width:32,height:32,borderRadius:"50%"}}><X size={20} /></button>
        </div>

        {/* Step indicator */}
        <div style={{display:"flex",marginBottom:24,background:"#0d0f1a",borderRadius:10,overflow:"hidden",border:"1px solid #1e2235"}}>
          {[["1","Upload File","drop"],["2","Review & Import","preview"],["3","Done","done"]].map(([n,label,s],i)=>(
            <div key={s} style={{flex:1,padding:"10px 0",textAlign:"center",background:stage===s?"linear-gradient(135deg,#5d7cff,#8b5cf6)":i<["drop","preview","done"].indexOf(stage)?"#1a3a2a":"transparent",borderRight:i<2?"1px solid #1e2235":"none"}}>
              <div style={{fontSize:11,fontWeight:700,color:stage===s?"#fff":i<["drop","preview","done"].indexOf(stage)?"#4ade80":"#4a5070",textTransform:"uppercase",letterSpacing:".05em"}}>Step {n}</div>
              <div style={{fontSize:12,color:stage===s?"#fff":i<["drop","preview","done"].indexOf(stage)?"#4ade80":"#6b7280",marginTop:2}}>{label}</div>
            </div>
          ))}
        </div>

        {/* ── DROP ── */}
        {stage==="drop"&&(
          <div>
            <div onDrop={handleDrop} onDragOver={e=>e.preventDefault()} onClick={()=>fileRef.current.click()}
              style={{border:"2px dashed #2e3350",borderRadius:14,padding:"44px 20px",textAlign:"center",cursor:"pointer",background:"#0d0f1a"}}
              onMouseEnter={e=>e.currentTarget.style.borderColor="#5d7cff"}
              onMouseLeave={e=>e.currentTarget.style.borderColor="#2e3350"}
            >
              <div style={{fontSize:40,marginBottom:10}}><FolderOpen size={40} style={{flexShrink:0}} /></div>
              <div style={{fontWeight:700,color:"#fff",fontSize:15,marginBottom:6}}>Drop your exported Products Excel file here</div>
              <div style={{color:"#6b7280",fontSize:13,marginBottom:16}}>Supports the file exported from this app (.xlsx)</div>
              <div style={{display:"inline-block",background:"linear-gradient(135deg,#5d7cff,#8b5cf6)",color:"#fff",padding:"8px 20px",borderRadius:8,fontWeight:600,fontSize:13}}>{loading?"Reading…":"Choose File"}</div>
              <input ref={fileRef} type="file" accept=".xlsx,.xls" style={{display:"none"}} onChange={e=>e.target.files[0]&&parseFile(e.target.files[0])}/>
            </div>
            {/* Tips */}
            <div style={{marginTop:16,background:"#0d0f1a",borderRadius:12,padding:16,border:"1px solid #1e2235"}}>
              <div style={{fontWeight:700,color:"#9ba3c0",fontSize:12,textTransform:"uppercase",letterSpacing:".05em",marginBottom:10}}>How to restore your data</div>
              {[["1️⃣","Export first — use '<Upload size={16} style={{flexShrink:0}} /> Export to Excel' to create a backup"],["2️⃣","Upload that same file here to restore"],["3️⃣","The app reads the 'BOM Details (Import This)' sheet automatically"],["4️⃣","Items must already exist in your Items list — export & import Items first if needed"]].map(([icon,text])=>(
                <div key={icon} style={{display:"flex",gap:10,marginBottom:8,alignItems:"flex-start"}}>
                  <span style={{fontSize:16,flexShrink:0}}>{icon}</span>
                  <span style={{fontSize:13,color:"#6b7280"}}>{text}</span>
                </div>
              ))}
            </div>
          </div>
        )}

        {/* ── PREVIEW ── */}
        {stage==="preview"&&(
          <div>
            <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:10,marginBottom:16}}>
              {[{label:"Products Found",val:rows.length,color:"#60a5fa"},{label:"Selected",val:selectedCount,color:"#4ade80"},{label:"BOM Lines",val:totalBomLines,color:"#f59e0b"},{label:"Warnings",val:errors.length,color:"#f87171"}].map(s=>(
                <div key={s.label} style={{background:"#0d0f1a",border:"1px solid #1e2235",borderRadius:10,padding:"10px 14px",textAlign:"center"}}>
                  <div style={{fontSize:22,fontWeight:800,color:s.color}}>{s.val}</div>
                  <div style={{fontSize:11,color:"#6b7280",marginTop:2}}>{s.label}</div>
                </div>
              ))}
            </div>

            {errors.length>0&&(
              <div style={{background:"#2a1a0a",border:"1px solid #f59e0b",borderRadius:10,padding:"10px 14px",marginBottom:12,maxHeight:90,overflowY:"auto"}}>
                {errors.map((e,i)=><div key={i} style={{fontSize:12,color:"#f59e0b",marginBottom:2}}><AlertTriangle size={16} style={{flexShrink:0}} /> {e}</div>)}
              </div>
            )}

            {/* Import mode */}
            <div style={{display:"flex",gap:10,marginBottom:12}}>
              {[["append",<><Plus size={16} style={{flexShrink:0}} /> Add to existing products (keep current)</>],["replace",<><RefreshCw size={16} style={{flexShrink:0}} /> Replace ALL products (full restore)</>]].map(([v,label])=>(
                <button key={v} onClick={()=>setMode(v)} style={{flex:1,padding:"9px 12px",borderRadius:8,fontSize:12,fontWeight:600,cursor:"pointer",background:mode===v?(v==="replace"?"#3a1a1a":"#1a3a2a"):"#0d0f1a",border:`1.5px solid ${mode===v?(v==="replace"?"#ff4757":"#4ade80"):"#1e2235"}`,color:mode===v?(v==="replace"?"#ff4757":"#4ade80"):"#6b7280"}}>{label}</button>
              ))}
            </div>

            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:6}}>
              <div style={{fontSize:12,color:"#6b7280"}}>{selectedCount} of {rows.length} products selected</div>
              <div style={{display:"flex",gap:8}}>
                <button onClick={()=>toggleAll(true)} style={{fontSize:11,color:"#5d7cff",background:"none",border:"none",cursor:"pointer"}}>Select All</button>
                <button onClick={()=>toggleAll(false)} style={{fontSize:11,color:"#ff4757",background:"none",border:"none",cursor:"pointer"}}>Deselect All</button>
              </div>
            </div>

            <div style={{maxHeight:300,overflowY:"auto",borderRadius:10,border:"1px solid #1e2235",marginBottom:16}}>
              <table style={{fontSize:12}}>
                <thead>
                  <tr>
                    <th style={{padding:"8px 10px",width:36}}><Check size={16} style={{flexShrink:0}} /></th>
                    <th style={{padding:"8px 10px"}}>Product Name</th>
                    <th style={{padding:"8px 10px"}}>Category</th>
                    <th style={{padding:"8px 10px",textAlign:"center"}}>BOM</th>
                    <th style={{padding:"8px 10px",textAlign:"right"}}>Landing Cost</th>
                    <th style={{padding:"8px 10px"}}>Status</th>
                  </tr>
                </thead>
                <tbody>
                  {rows.map((row,i)=>{
                    const fakeP = { ...row, materials: row.materials };
                    const landing = row.materials.reduce((s,m)=>{const it=items.find(x=>x.id===m.itemId);return s+(it?it.cost*m.qty:0);},0)+(row.factoryCost||0)+(row.overheadCost||0)+(row.shippingCost||0);
                    return (
                    <tr key={i} style={{opacity:row._selected?1:0.4,cursor:"pointer"}} onClick={()=>toggleRow(i)}>
                      <td style={{padding:"7px 10px",textAlign:"center"}}><input type="checkbox" checked={row._selected} onChange={()=>toggleRow(i)} onClick={e=>e.stopPropagation()} style={{width:14,height:14,accentColor:"#5d7cff",cursor:"pointer"}}/></td>
                      <td style={{padding:"7px 10px",fontWeight:600,color:"#e8eaf0"}}>{row.name}</td>
                      <td style={{padding:"7px 10px"}}><span style={{fontSize:10,padding:"2px 7px",borderRadius:8,background:"#1e2a4a",color:"#60a5fa",fontWeight:600}}>{row.category}</span></td>
                      <td style={{padding:"7px 10px",textAlign:"center",color:"#9ba3c0"}}>{row.materials.length} items</td>
                      <td style={{padding:"7px 10px",textAlign:"right",fontWeight:700,color:"#f59e0b"}}>₹{fmt(landing)}</td>
                      <td style={{padding:"7px 10px"}}>
                        {row._dupInApp?<span style={{fontSize:10,padding:"2px 7px",borderRadius:8,background:"#3a2a0a",color:"#f59e0b",fontWeight:600}}><AlertTriangle size={16} style={{flexShrink:0}} /> Duplicate</span>:<span style={{fontSize:10,padding:"2px 7px",borderRadius:8,background:"#1a3a2a",color:"#4ade80",fontWeight:600}}><Check size={16} style={{flexShrink:0}} /> New</span>}
                      </td>
                    </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>

            {dupCount>0&&mode==="append"&&(
              <div style={{background:"#2a2a0a",border:"1px solid #f59e0b",borderRadius:8,padding:"8px 14px",marginBottom:12,fontSize:12,color:"#f59e0b"}}>
                <AlertTriangle size={16} style={{flexShrink:0}} /> {dupCount} product(s) have the same name as existing products. They will be added as duplicates.
              </div>
            )}

            <div style={{display:"flex",gap:12}}>
              <button className="btn-primary" onClick={doImport} style={{flex:2,padding:13}}>
                {mode==="replace"?<><RefreshCw size={16} style={{flexShrink:0}} /> Restore</>:<><Plus size={16} style={{flexShrink:0}} /> Import</>} {selectedCount} Product{selectedCount!==1?"s":""} ({totalBomLines} BOM lines)
              </button>
              <button className="btn-ghost" onClick={()=>setStage("drop")} style={{flex:1,padding:13}}>← Back</button>
              <button className="btn-ghost" onClick={onClose} style={{flex:1,padding:13}}>Cancel</button>
            </div>
          </div>
        )}

        {/* ── DONE ── */}
        {stage==="done"&&(
          <div style={{textAlign:"center",padding:"32px 0"}}>
            <div style={{fontSize:52,marginBottom:16}}><CheckCircle size={52} style={{flexShrink:0}} /></div>
            <div style={{fontWeight:700,fontSize:20,color:"#fff",marginBottom:8}}>Import Successful!</div>
            <div style={{color:"#6b7280",fontSize:14,marginBottom:28}}>
              {selectedCount} product{selectedCount!==1?"s":""} with {totalBomLines} BOM lines have been {mode==="replace"?"restored (replaced all)":"added to your products list"}.
            </div>
            <button className="btn-primary" onClick={onClose} style={{padding:"12px 40px",fontSize:15}}>Done</button>
          </div>
        )}
      </div>
    </div>
  );
}

// ── Cost Sheet ────────────────────────────────────────────────────────────────
function CostSheet({ products, items }) {
  const [selectedProd,setSelectedProd]=useState(products[0]?.id||"");
  const [sortKey,setSortKey]=useState("name-az");
  const prod=products.find(p=>p.id===selectedProd);

  const sortedMaterials=prod?applySortKey(
    prod.materials.map(m=>({...m,item:items.find(i=>i.id===m.itemId)})).filter(m=>m.item),
    sortKey,{name:m=>m.item.name,unit:m=>m.item.unit,qty:m=>m.qty,cost:m=>m.item.cost,amount:m=>m.item.cost*m.qty}
  ):[];

  function handlePrint() {
    if(!prod)return;
    const mat=calcMaterialCost(prod,items),landing=calcLandingCost(prod,items),selling=calcSellingPrice(prod,items);
    const printMats=[...prod.materials].map(m=>({...m,item:items.find(i=>i.id===m.itemId)})).filter(m=>m.item).sort((a,b)=>a.item.name.localeCompare(b.item.name));
    const rows=printMats.map((m,i)=>`<tr><td>${i+1}</td><td class="b">${m.item.name}</td><td>${m.item.category}</td><td>${m.item.unit}</td><td>${m.qty}</td><td class="r">₹${m.item.cost.toLocaleString("en-IN")}</td><td class="r b">₹${fmt(m.item.cost*m.qty)}</td></tr>`).join("");
    printPage(`Cost Sheet — ${prod.name}`,`
      <div class="sb" style="margin-bottom:12px">
        <div class="s"><div class="sl">Product</div><div class="sv" style="font-size:13px">${prod.name}</div></div>
        <div class="s"><div class="sl">Category</div><div class="sv" style="font-size:13px">${prod.category}</div></div>
        <div class="s"><div class="sl">Profit Margin</div><div class="sv" style="font-size:13px">${prod.profitMargin}%</div></div>
        <div class="s"><div class="sl">BOM Items</div><div class="sv" style="font-size:13px">${prod.materials.length}</div></div>
      </div>
      <div class="stl">Bill of Materials (A–Z)</div>
      <table>
        <thead><tr><th>#</th><th>Item</th><th>Category</th><th>Unit</th><th>Qty</th><th style="text-align:right">Rate (₹)</th><th style="text-align:right">Amount (₹)</th></tr></thead>
        <tbody>${rows}<tr class="tr"><td colspan="6" class="b">Total Material Cost</td><td class="r b">₹${fmt(mat)}</td></tr></tbody>
      </table>
      <div class="stl">Cost Summary</div>
      <div class="sb">
        <div class="s"><div class="sl">Material Cost</div><div class="sv">₹${fmt(mat)}</div></div>
        <div class="s"><div class="sl">Factory Cost</div><div class="sv">₹${fmt(prod.factoryCost||0)}</div></div>
        <div class="s"><div class="sl">Overhead Cost</div><div class="sv">₹${fmt(prod.overheadCost||0)}</div></div>
        <div class="s"><div class="sl">Shipping Cost</div><div class="sv">₹${fmt(prod.shippingCost||0)}</div></div>
        <div class="s sd"><div class="sl">Landing Cost</div><div class="sv">₹${fmt(landing)}</div></div>
        <div class="s sd"><div class="sl">Selling Price (+${prod.profitMargin}%)</div><div class="sv">₹${fmt(selling)}</div></div>
      </div>`);
  }

  return (
    <div>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:4}}>
        <div className="section-title">Detailed Cost Sheet</div>
        {prod&&<PrintBtn onClick={handlePrint} label={<><Printer size={16} style={{flexShrink:0}} /> Print Cost Sheet</>}/>}
      </div>
      <div className="section-sub">Full breakdown of material, labor, overhead, and landing cost per product</div>
      <div style={{marginBottom:24,maxWidth:360}}>
        <select value={selectedProd} onChange={e=>setSelectedProd(e.target.value)}>
          {products.map(p=><option key={p.id} value={p.id}>{p.name}</option>)}
        </select>
      </div>
      {prod?(
        <div style={{display:"grid",gridTemplateColumns:"1.5fr 1fr",gap:24}}>
          <div className="card" style={{padding:0,overflow:"hidden"}}>
            <div style={{padding:"20px 20px 8px",fontWeight:700,fontSize:16,color:"#fff"}}>Bill of Materials — {prod.name}</div>
            <div style={{padding:"0 20px 12px",fontSize:11,color:"#6b7280"}}>Click column headers to sort</div>
            <table>
              <thead><tr>
                <SortTh label="Item" field="name" sortKey={sortKey} onSort={setSortKey}/>
                <SortTh label="Unit" field="unit" sortKey={sortKey} onSort={setSortKey}/>
                <SortTh label="Qty" field="qty" sortKey={sortKey} onSort={setSortKey}/>
                <SortTh label="Rate (₹)" field="cost" sortKey={sortKey} onSort={setSortKey}/>
                <SortTh label="Amount (₹)" field="amount" sortKey={sortKey} onSort={setSortKey}/>
              </tr></thead>
              <tbody>
                {sortedMaterials.map((m,idx)=>(
                  <tr key={idx}>
                    <td style={{fontWeight:600,color:"#fff"}}>{m.item.name}</td>
                    <td style={{color:"#9ba3c0"}}>{m.item.unit}</td>
                    <td>{m.qty}</td>
                    <td>₹{m.item.cost.toLocaleString("en-IN")}</td>
                    <td style={{fontWeight:700,color:"#e8eaf0"}}>₹{fmt(m.item.cost*m.qty)}</td>
                  </tr>
                ))}
              </tbody>
            </table>
            <div style={{padding:"12px 16px",borderTop:"2px solid #2e3350",display:"flex",justifyContent:"space-between",background:"#1c2135"}}>
              <span style={{fontWeight:700,color:"#fff"}}>Total Material Cost</span>
              <span style={{fontWeight:800,color:"#f59e0b",fontSize:16}}>₹{fmt(calcMaterialCost(prod,items))}</span>
            </div>
          </div>
          <div>
            <div className="card" style={{marginBottom:20}}>
              <div style={{fontWeight:700,fontSize:16,color:"#fff",marginBottom:16}}>Cost Summary</div>
              {[{label:"Material Cost",val:calcMaterialCost(prod,items)},{label:"Factory Cost",val:prod.factoryCost||0},{label:"Overhead Cost",val:prod.overheadCost||0},{label:"Shipping Cost",val:prod.shippingCost||0}].map(r=>(
                <div key={r.label} className="cost-row">
                  <span style={{color:"#9ba3c0",fontSize:14}}>{r.label}</span>
                  <span style={{fontWeight:600,color:"#e8eaf0"}}>₹{fmt(r.val)}</span>
                </div>
              ))}
              <div style={{background:"#1c2135",borderRadius:12,padding:"16px",marginTop:12}}>
                <div style={{display:"flex",justifyContent:"space-between",marginBottom:8}}>
                  <span style={{color:"#f59e0b",fontWeight:700}}><Tag size={16} style={{flexShrink:0}} /> Landing Cost</span>
                  <span style={{color:"#f59e0b",fontWeight:800,fontSize:22}}>₹{fmt(calcLandingCost(prod,items))}</span>
                </div>
                <div style={{display:"flex",justifyContent:"space-between"}}>
                  <span style={{color:"#10b981",fontWeight:700}}><IndianRupee size={24} style={{flexShrink:0}} /> Selling Price (+{prod.profitMargin}%)</span>
                  <span style={{color:"#10b981",fontWeight:800,fontSize:22}}>₹{fmt(calcSellingPrice(prod,items))}</span>
                </div>
              </div>
            </div>
            <div className="card">
              <div style={{fontWeight:700,fontSize:16,color:"#fff",marginBottom:16}}>Cost Breakdown</div>
              {(()=>{const mat=calcMaterialCost(prod,items),landing=calcLandingCost(prod,items);return[{label:"Material",val:mat,color:"#5d7cff"},{label:"Factory",val:prod.factoryCost||0,color:"#8b5cf6"},{label:"Overhead",val:prod.overheadCost||0,color:"#f59e0b"},{label:"Shipping",val:prod.shippingCost||0,color:"#10b981"}].map(s=>(
                <div key={s.label} style={{marginBottom:12}}>
                  <div style={{display:"flex",justifyContent:"space-between",marginBottom:4,fontSize:13}}>
                    <span style={{color:"#9ba3c0"}}>{s.label}</span>
                    <span style={{color:"#e8eaf0",fontWeight:600}}>{landing>0?((s.val/landing)*100).toFixed(1):0}%</span>
                  </div>
                  <div style={{height:8,background:"#1e2235",borderRadius:4,overflow:"hidden"}}>
                    <div style={{height:"100%",width:`${landing>0?(s.val/landing)*100:0}%`,background:s.color,borderRadius:4,transition:"width .5s"}}></div>
                  </div>
                </div>
              ))})()}
            </div>
          </div>
        </div>
      ):(
        <div className="card" style={{textAlign:"center",padding:"60px 20px",color:"#4a5070",display:"flex",flexDirection:"column",alignItems:"center",gap:16}}>
        <div style={{width:64,height:64,borderRadius:"50%",background:"rgba(139,92,246,0.1)",display:"flex",alignItems:"center",justifyContent:"center",color:"#8b5cf6",marginBottom:8}}>
          <FileText size={32} />
        </div>
        <div>
          <div style={{fontSize:16,fontWeight:700,color:"#e8eaf0",marginBottom:4}}>No products available</div>
          <div style={{fontSize:14,color:"#6b7280",maxWidth:300,margin:"0 auto"}}>You need to add products before you can view cost sheets.</div>
        </div>
      </div>
      )}
    </div>
  );
}

// ── Purchase Order ────────────────────────────────────────────────────────────
function PurchaseOrder({ products, items, purchaseOrders, setPurchaseOrders }) {

  // ── View: "list" shows saved POs, "editor" shows the PO form ──────────────
  const [view, setView] = useState("list"); // "list" | "editor"
  const [editingId, setEditingId] = useState(null); // null = new PO

  // ── Editor state ───────────────────────────────────────────────────────────
  const [orderLines, setOrderLines] = useState([]);
  const [selProducts, setSelProducts] = useState({});
  const [poNumber, setPoNumber] = useState("");
  const [supplier, setSupplier] = useState("");
  const [poDate, setPoDate] = useState(() => new Date().toISOString().slice(0,10));
  const [poStatus, setPoStatus] = useState("Draft"); // Draft | Ordered | Received
  const [notes, setNotes] = useState("");
  const [searchItem, setSearchItem] = useState("");
  const [selCats, setSelCats] = useState([]);

  // ── List search ────────────────────────────────────────────────────────────
  const [listSearch, setListSearch] = useState("");
  const [listStatus, setListStatus] = useState("All");

  // ── Helpers ────────────────────────────────────────────────────────────────
  const allCats = [...new Set(items.map(i => i.category))].sort();
  const catColors = { Electrical:"#60a5fa","Raw Material":"#4ade80",Chemical:"#d946ef",Fittings:"#fb923c",Plumbing:"#38bdf8",Metal:"#fbbf24",Other:"#f87171" };
  const catBg    = { Electrical:"#1e3a5f","Raw Material":"#1a3a2a",Chemical:"#3a1a3a",Fittings:"#3a2a1a",Plumbing:"#1a2a3a",Metal:"#2a2a1a",Other:"#2a1a1a" };
  const tagClass = cat => { const m={Electrical:"tag-elec","Raw Material":"tag-raw",Chemical:"tag-chem",Fittings:"tag-fit",Plumbing:"tag-plum",Metal:"tag-met"}; return "tag "+(m[cat]||"tag-oth"); };
  const statusColor = s => s==="Received"?"#4ade80":s==="Ordered"?"#f59e0b":"#6b7280";
  const statusBg    = s => s==="Received"?"#1a3a2a":s==="Ordered"?"#2a2a0a":"#1e2235";

  const grandTotal = orderLines.reduce((s,l) => s+l.orderQty*l.unitCost, 0);
  const totalUnits = orderLines.reduce((s,l) => s+l.orderQty, 0);

  const sortedLines = [...orderLines].sort((a,b) => {
    const ia=items.find(i=>i.id===a.itemId), ib=items.find(i=>i.id===b.itemId);
    return (ia?.name||"").localeCompare(ib?.name||"");
  });

  // ── Open new PO editor ─────────────────────────────────────────────────────
  function openNew() {
    const newNum = "PO-"+new Date().getFullYear()+"-"+String(Math.floor(Math.random()*9000)+1000);
    setEditingId(null);
    setOrderLines([]); setSelProducts({});
    setPoNumber(newNum); setSupplier(""); setPoDate(new Date().toISOString().slice(0,10));
    setPoStatus("Draft"); setNotes(""); setSearchItem(""); setSelCats([]);
    setView("editor");
  }

  // ── Open existing PO for editing ───────────────────────────────────────────
  function openEdit(po) {
    setEditingId(po.id);
    setOrderLines(po.orderLines||[]);
    setSelProducts(po.selProducts||{});
    setPoNumber(po.poNumber||""); setSupplier(po.supplier||"");
    setPoDate(po.poDate||new Date().toISOString().slice(0,10));
    setPoStatus(po.status||"Draft"); setNotes(po.notes||"");
    setSearchItem(""); setSelCats([]);
    setView("editor");
  }

  // ── Save (create or update) ────────────────────────────────────────────────
  function savePO() {
    if (!poNumber.trim()) { alert("Please enter a PO Number"); return; }
    if (orderLines.length === 0) { alert("Please add at least one item"); return; }
    const po = {
      id: editingId || uid(),
      poNumber, supplier, poDate, status: poStatus, notes,
      orderLines, selProducts,
      grandTotal, totalItems: orderLines.length,
      createdAt: editingId ? (purchaseOrders.find(p=>p.id===editingId)?.createdAt || new Date().toISOString()) : new Date().toISOString(),
      updatedAt: new Date().toISOString(),
    };
    if (editingId) {
      setPurchaseOrders(prev => prev.map(p => p.id===editingId ? po : p));
    } else {
      setPurchaseOrders(prev => [po, ...prev]);
    }
    setView("list");
  }

  // ── Delete PO ──────────────────────────────────────────────────────────────
  function deletePO(id) {
    if (window.confirm("Delete this Purchase Order? This cannot be undone.")) {
      setPurchaseOrders(prev => prev.filter(p => p.id !== id));
    }
  }

  // ── Duplicate PO ──────────────────────────────────────────────────────────
  function duplicatePO(po) {
    const newNum = "PO-"+new Date().getFullYear()+"-"+String(Math.floor(Math.random()*9000)+1000);
    setEditingId(null);
    setOrderLines([...po.orderLines]); setSelProducts({...po.selProducts});
    setPoNumber(newNum); setSupplier(po.supplier||"");
    setPoDate(new Date().toISOString().slice(0,10));
    setPoStatus("Draft"); setNotes(po.notes||"");
    setSearchItem(""); setSelCats([]);
    setView("editor");
  }

  // ── Order line helpers ─────────────────────────────────────────────────────
  function addItem(itemId) {
    if (orderLines.find(l=>l.itemId===itemId)) return;
    const item = items.find(i=>i.id===itemId);
    setOrderLines(prev=>[...prev,{itemId,orderQty:1,unitCost:item?.cost||0}]);
  }
  function removeItem(itemId) { setOrderLines(prev=>prev.filter(l=>l.itemId!==itemId)); }
  function updateLine(itemId,field,val) {
    setOrderLines(prev=>prev.map(l=>l.itemId===itemId?{...l,[field]:parseFloat(val)||0}:l));
  }

  // ── Auto-fill from products ────────────────────────────────────────────────
  function applyProductDemand() {
    const demand={};
    Object.entries(selProducts).forEach(([pid,prodQty])=>{
      const prod=products.find(p=>p.id===pid);
      if(!prod||!prodQty) return;
      prod.materials.forEach(m=>{ demand[m.itemId]=(demand[m.itemId]||0)+m.qty*prodQty; });
    });
    const newLines=Object.entries(demand).map(([itemId,qty])=>{
      const item=items.find(i=>i.id===itemId);
      const existing=orderLines.find(l=>l.itemId===itemId);
      return {itemId,orderQty:qty,unitCost:existing?.unitCost??item?.cost??0};
    });
    const productItemIds=new Set(Object.keys(demand));
    const manualOnly=orderLines.filter(l=>!productItemIds.has(l.itemId));
    setOrderLines([...manualOnly,...newLines].sort((a,b)=>{
      const ia=items.find(i=>i.id===a.itemId),ib=items.find(i=>i.id===b.itemId);
      return (ia?.name||"").localeCompare(ib?.name||"");
    }));
  }

  // ── Available items for manual add ────────────────────────────────────────
  const inOrder=new Set(orderLines.map(l=>l.itemId));
  const availableItems=applySortKey(
    items.filter(i=>!inOrder.has(i.id)
      &&(searchItem===""||i.name.toLowerCase().includes(searchItem.toLowerCase())||i.category.toLowerCase().includes(searchItem.toLowerCase()))
      &&(selCats.length===0||selCats.includes(i.category))
    ),"name-az",{name:i=>i.name,category:i=>i.category,cost:i=>i.cost}
  );

  // ── Filtered saved PO list ─────────────────────────────────────────────────
  const filteredPOs = purchaseOrders.filter(po =>
    (listStatus==="All"||po.status===listStatus) &&
    (listSearch===""||po.poNumber.toLowerCase().includes(listSearch.toLowerCase())||
      (po.supplier||"").toLowerCase().includes(listSearch.toLowerCase()))
  );

  // ── Print ──────────────────────────────────────────────────────────────────
  function handlePrint(po) {
    const lines = po ? [...(po.orderLines||[])].sort((a,b)=>{const ia=items.find(i=>i.id===a.itemId),ib=items.find(i=>i.id===b.itemId);return(ia?.name||"").localeCompare(ib?.name||"");}) : sortedLines;
    const pNum = po?.poNumber||poNumber; const pSup = po?.supplier||supplier;
    const pDate = po?.poDate||poDate; const pNotes = po?.notes||notes;
    const pTotal = po?.grandTotal??grandTotal;

    const rows=lines.map((l,i)=>{
      const item=items.find(x=>x.id===l.itemId); if(!item) return "";
      return `<tr><td>${i+1}</td><td class="b">${item.name}</td><td>${item.category}</td><td>${item.unit}</td><td style="text-align:right">${l.orderQty}</td><td style="text-align:right">₹${l.unitCost.toLocaleString("en-IN",{maximumFractionDigits:2})}</td><td style="text-align:right" class="b">₹${fmt(l.orderQty*l.unitCost)}</td></tr>`;
    }).join("");

    const catSummary={};
    lines.forEach(l=>{ const item=items.find(x=>x.id===l.itemId);if(!item)return; catSummary[item.category]=(catSummary[item.category]||0)+l.orderQty*l.unitCost; });
    const catRows=Object.entries(catSummary).sort((a,b)=>b[1]-a[1])
      .map(([cat,val])=>`<div class="s"><div class="sl">${cat}</div><div class="sv" style="font-size:14px">₹${fmt(val)}</div></div>`).join("");

    printPage(`Purchase Order — ${pNum}`,`
      <div class="sb" style="margin-bottom:14px">
        <div class="s"><div class="sl">PO Number</div><div class="sv" style="font-size:14px">${pNum}</div></div>
        <div class="s"><div class="sl">Date</div><div class="sv" style="font-size:14px">${new Date(pDate).toLocaleDateString("en-IN",{day:"2-digit",month:"short",year:"numeric"})}</div></div>
        <div class="s"><div class="sl">Supplier</div><div class="sv" style="font-size:14px">${pSup||"—"}</div></div>
        <div class="s"><div class="sl">Status</div><div class="sv" style="font-size:14px">${po?.status||poStatus}</div></div>
        <div class="s sd"><div class="sl">Grand Total</div><div class="sv">₹${fmt(pTotal)}</div></div>
      </div>
      <div class="stl">Order Items (A–Z)</div>
      <table>
        <thead><tr><th>#</th><th>Item Name</th><th>Category</th><th>Unit</th><th style="text-align:right">Qty</th><th style="text-align:right">Unit Cost (₹)</th><th style="text-align:right">Amount (₹)</th></tr></thead>
        <tbody>${rows}<tr class="tr"><td colspan="4" class="b">GRAND TOTAL</td><td style="text-align:right" class="b">${lines.reduce((s,l)=>s+l.orderQty,0)}</td><td></td><td style="text-align:right" class="b">₹${fmt(pTotal)}</td></tr></tbody>
      </table>
      <div class="stl">Category Summary</div>
      <div class="sb">${catRows}</div>
      ${pNotes?`<div class="stl">Notes</div><div style="font-size:12px;color:#555;padding:8px;border:1px solid #ddd;border-radius:4px">${pNotes}</div>`:""}
    `);
  }

  // ══════════════════════════════════════════════════════════════════════════
  // RENDER: LIST VIEW
  // ══════════════════════════════════════════════════════════════════════════
  if (view === "list") return (
    <div>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:4}}>
        <div>
          <div className="section-title"><ShoppingCart size={16} style={{flexShrink:0}} /> Purchase Orders</div>
          <div className="section-sub">Save, manage and track all your purchase orders</div>
        </div>
        <button className="btn-primary" onClick={openNew} style={{padding:"10px 22px",fontSize:14}}>+ New Purchase Order</button>
      </div>

      {/* Summary stats */}
      <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:14,marginBottom:20}}>
        {[
          {label:"Total POs",val:purchaseOrders.length,color:"#60a5fa",icon:<ClipboardList size={16} style={{flexShrink:0}} />},
          {label:"Draft",val:purchaseOrders.filter(p=>p.status==="Draft").length,color:"#6b7280",icon:<Edit2 size={14} style={{flexShrink:0}} />},
          {label:"Ordered",val:purchaseOrders.filter(p=>p.status==="Ordered").length,color:"#f59e0b",icon:<Package size={24} style={{flexShrink:0}} />},
          {label:"Received",val:purchaseOrders.filter(p=>p.status==="Received").length,color:"#4ade80",icon:<CheckCircle size={52} style={{flexShrink:0}} />},
        ].map(s=>(
          <div key={s.label} className="card" style={{padding:"16px 20px",textAlign:"center"}}>
            <div style={{fontSize:22,marginBottom:4}}>{s.icon}</div>
            <div style={{fontSize:22,fontWeight:800,color:s.color}}>{s.val}</div>
            <div style={{fontSize:11,color:"#6b7280",marginTop:4,textTransform:"uppercase",letterSpacing:".05em"}}>{s.label}</div>
          </div>
        ))}
      </div>

      {/* Filters */}
      <div style={{display:"flex",gap:10,marginBottom:16,alignItems:"center",flexWrap:"wrap"}}>
        <input placeholder="Search by PO number or supplier…" value={listSearch} onChange={e=>setListSearch(e.target.value)} style={{maxWidth:320}}/>
        <div style={{display:"flex",gap:6}}>
          {["All","Draft","Ordered","Received"].map(s=>(
            <button key={s} onClick={()=>setListStatus(s)} style={{
              padding:"7px 14px",borderRadius:8,border:"none",fontSize:12,fontWeight:700,cursor:"pointer",
              background:listStatus===s?(s==="All"?"linear-gradient(135deg,#5d7cff,#8b5cf6)":statusBg(s)):"#1e2235",
              color:listStatus===s?(s==="All"?"#fff":statusColor(s)):"#6b7280",
              border:listStatus===s&&s!=="All"?`1px solid ${statusColor(s)}`:"1px solid transparent",
            }}>{s}</button>
          ))}
        </div>
      </div>

      {/* PO List */}
      {filteredPOs.length === 0 ? (
        <div className="card" style={{textAlign:"center",padding:60,color:"#4a5070"}}>
          <div style={{fontSize:40,marginBottom:12}}><ShoppingCart size={16} style={{flexShrink:0}} /></div>
          <div style={{fontSize:16,fontWeight:600,marginBottom:8}}>
            {purchaseOrders.length===0?"No purchase orders yet":"No orders match your filter"}
          </div>
          <div style={{fontSize:13,marginBottom:20}}>
            {purchaseOrders.length===0?"Create your first PO to get started":"Try clearing your search or filter"}
          </div>
          {purchaseOrders.length===0&&<button className="btn-primary" onClick={openNew} style={{padding:"10px 24px"}}>+ Create First PO</button>}
        </div>
      ) : (
        <div style={{display:"flex",flexDirection:"column",gap:12}}>
          {filteredPOs.map(po=>(
            <div key={po.id} className="card" style={{padding:"16px 20px"}}>
              <div style={{display:"flex",alignItems:"flex-start",gap:16}}>
                {/* Left info */}
                <div style={{flex:1,minWidth:0}}>
                  <div style={{display:"flex",alignItems:"center",gap:10,marginBottom:6,flexWrap:"wrap"}}>
                    <span style={{fontWeight:800,color:"#fff",fontSize:16}}>{po.poNumber}</span>
                    <span style={{padding:"3px 10px",borderRadius:20,fontSize:11,fontWeight:700,background:statusBg(po.status),color:statusColor(po.status),border:`1px solid ${statusColor(po.status)}44`}}>
                      {po.status==="Received"?<CheckCircle size={16} style={{flexShrink:0}} />:po.status==="Ordered"?<Package size={16} style={{flexShrink:0}} />:<Edit2 size={14} style={{flexShrink:0}} />} {po.status}
                    </span>
                    {po.supplier&&<span style={{fontSize:12,color:"#9ba3c0"}}><Factory size={24} style={{flexShrink:0}} /> {po.supplier}</span>}
                    <span style={{fontSize:12,color:"#6b7280"}}><Calendar size={16} style={{flexShrink:0}} /> {new Date(po.poDate).toLocaleDateString("en-IN",{day:"2-digit",month:"short",year:"numeric"})}</span>
                  </div>

                  {/* Item count + category breakdown */}
                  <div style={{display:"flex",gap:8,flexWrap:"wrap",marginBottom:8}}>
                    <span style={{fontSize:12,color:"#6b7280"}}>{po.totalItems} items · {po.orderLines?.reduce((s,l)=>s+l.orderQty,0)||0} units</span>
                  </div>

                  {/* Category chips */}
                  <div style={{display:"flex",gap:6,flexWrap:"wrap"}}>
                    {(()=>{
                      const cats={};
                      (po.orderLines||[]).forEach(l=>{
                        const item=items.find(i=>i.id===l.itemId); if(!item) return;
                        cats[item.category]=(cats[item.category]||0)+l.orderQty*l.unitCost;
                      });
                      return Object.entries(cats).sort((a,b)=>b[1]-a[1]).slice(0,5).map(([cat,val])=>(
                        <span key={cat} style={{padding:"2px 8px",borderRadius:20,fontSize:10,fontWeight:700,background:catBg[cat]||"#2a1a1a",color:catColors[cat]||"#f87171",border:"1px solid color-mix(in srgb, currentColor 15%, transparent)"}}>
                          {cat} ₹{fmt(val)}
                        </span>
                      ));
                    })()}
                  </div>
                </div>

                {/* Right: total + actions */}
                <div style={{display:"flex",flexDirection:"column",alignItems:"flex-end",gap:10,flexShrink:0}}>
                  <div style={{textAlign:"right"}}>
                    <div style={{fontSize:11,color:"#6b7280",textTransform:"uppercase",letterSpacing:".05em"}}>Grand Total</div>
                    <div style={{fontSize:22,fontWeight:800,color:"#f59e0b"}}>₹{fmt(po.grandTotal||0)}</div>
                  </div>
                  <div style={{display:"flex",gap:8}}>
                    <button onClick={()=>handlePrint(po)}
                      style={{padding:"6px 12px",borderRadius:8,background:"#1e2130",border:"1px solid #2e3350",color:"#9ba3c0",fontSize:12,fontWeight:600,cursor:"pointer"}}>
                      <Printer size={16} style={{flexShrink:0}} /> Print
                    </button>
                    <button onClick={()=>duplicatePO(po)}
                      style={{padding:"6px 12px",borderRadius:8,background:"#1a2a3a",border:"1px solid #1e3a5f",color:"#60a5fa",fontSize:12,fontWeight:600,cursor:"pointer"}}>
                      <ClipboardList size={16} style={{flexShrink:0}} /> Copy
                    </button>
                    <button onClick={()=>openEdit(po)}
                      style={{padding:"6px 12px",borderRadius:8,background:"linear-gradient(135deg,#5d7cff,#8b5cf6)",border:"none",color:"#fff",fontSize:12,fontWeight:600,cursor:"pointer"}}>
                      <Edit2 size={14} style={{flexShrink:0}} /> Edit
                    </button>
                    <button onClick={()=>deletePO(po.id)}
                      style={{padding:"6px 12px",borderRadius:8,background:"#2a1a1a",border:"1px solid #ff4757",color:"#ff4757",fontSize:12,fontWeight:600,cursor:"pointer"}}>
                      <Trash2 size={16} style={{flexShrink:0}} />
                    </button>
                  </div>
                </div>
              </div>

              {/* Notes if any */}
              {po.notes&&<div style={{marginTop:10,padding:"8px 12px",background:"#0d0f1a",borderRadius:8,fontSize:12,color:"#6b7280",borderLeft:"3px solid #2e3350"}}><FileText size={16} style={{flexShrink:0}} /> {po.notes}</div>}
            </div>
          ))}
        </div>
      )}
    </div>
  );

  // ══════════════════════════════════════════════════════════════════════════
  // RENDER: EDITOR VIEW
  // ══════════════════════════════════════════════════════════════════════════
  return (
    <div>
      {/* Header */}
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:16}}>
        <div style={{display:"flex",alignItems:"center",gap:12}}>
          <button onClick={()=>setView("list")} style={{background:"#1e2130",border:"1px solid #2e3350",color:"#9ba3c0",padding:"8px 14px",borderRadius:8,fontSize:13,fontWeight:600,cursor:"pointer"}}>← Back</button>
          <div>
            <div className="section-title" style={{marginBottom:0}}>{editingId?<><Edit2 size={16} style={{flexShrink:0}} /> Edit Purchase Order</>:<><Plus size={16} style={{flexShrink:0}} /> New Purchase Order</>}</div>
            <div style={{fontSize:12,color:"#6b7280",marginTop:2}}>{editingId?"Update your order details below":"Fill in details and add items, then save"}</div>
          </div>
        </div>
        <div style={{display:"flex",gap:10}}>
          {orderLines.length>0&&<PrintBtn onClick={()=>handlePrint(null)} label={<><Printer size={16} style={{flexShrink:0}} /> Print</>}/>}
          <button onClick={savePO} style={{background:"linear-gradient(135deg,#10b981,#059669)",border:"none",color:"#fff",padding:"10px 24px",borderRadius:10,fontSize:14,fontWeight:700,cursor:"pointer"}}>
            <Save size={16} style={{flexShrink:0}} /> {editingId?"Update PO":"Save PO"}
          </button>
        </div>
      </div>

      {/* PO Details */}
      <div className="card" style={{marginBottom:16,padding:"16px 20px"}}>
        <div style={{fontWeight:700,color:"#9ba3c0",fontSize:11,textTransform:"uppercase",letterSpacing:".05em",marginBottom:12}}>Purchase Order Details</div>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr 1fr 1fr",gap:12}}>
          <div className="form-row" style={{marginBottom:0}}>
            <label className="form-label">PO Number *</label>
            <input value={poNumber} onChange={e=>setPoNumber(e.target.value)} placeholder="PO-2025-0001"/>
          </div>
          <div className="form-row" style={{marginBottom:0}}>
            <label className="form-label">Order Date</label>
            <input type="date" value={poDate} onChange={e=>setPoDate(e.target.value)}/>
          </div>
          <div className="form-row" style={{marginBottom:0}}>
            <label className="form-label">Supplier / Vendor</label>
            <input value={supplier} onChange={e=>setSupplier(e.target.value)} placeholder="Supplier name"/>
          </div>
          <div className="form-row" style={{marginBottom:0}}>
            <label className="form-label">Status</label>
            <select value={poStatus} onChange={e=>setPoStatus(e.target.value)}>
              <option>Draft</option><option>Ordered</option><option>Received</option>
            </select>
          </div>
          <div className="form-row" style={{marginBottom:0}}>
            <label className="form-label">Notes</label>
            <input value={notes} onChange={e=>setNotes(e.target.value)} placeholder="Payment terms, delivery…"/>
          </div>
        </div>
      </div>

      {/* Summary */}
      <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:12,marginBottom:16}}>
        {[
          {label:"Items",val:orderLines.length,color:"#60a5fa",icon:<Package size={24} style={{flexShrink:0}} />},
          {label:"Total Units",val:fmt(totalUnits),color:"#8b5cf6",icon:<Hash size={16} style={{flexShrink:0}} />},
          {label:"Grand Total",val:"₹"+fmt(grandTotal),color:"#f59e0b",icon:<IndianRupee size={24} style={{flexShrink:0}} />,big:true},
          {label:"Avg per Item",val:orderLines.length?"₹"+fmt(grandTotal/orderLines.length):"—",color:"#10b981",icon:<BarChart2 size={16} style={{flexShrink:0}} />},
        ].map(s=>(
          <div key={s.label} className="card" style={{padding:"12px 16px",textAlign:"center"}}>
            <div style={{fontSize:18,marginBottom:3}}>{s.icon}</div>
            <div style={{fontSize:s.big?20:16,fontWeight:800,color:s.color}}>{s.val}</div>
            <div style={{fontSize:11,color:"#6b7280",marginTop:3,textTransform:"uppercase",letterSpacing:".05em"}}>{s.label}</div>
          </div>
        ))}
      </div>

      <div style={{display:"grid",gridTemplateColumns:"380px 1fr",gap:16,alignItems:"start"}}>

        {/* LEFT: add items */}
        <div style={{display:"flex",flexDirection:"column",gap:14}}>

          {/* Auto-fill */}
          <div className="card" style={{padding:"16px 18px"}}>
            <div style={{fontWeight:700,color:"#fff",fontSize:13,marginBottom:4}}><Zap size={16} style={{flexShrink:0}} /> Auto-fill from Products</div>
            <div style={{fontSize:12,color:"#6b7280",marginBottom:12}}>Set production quantities to auto-calculate items</div>
            {products.map(p=>(
              <div key={p.id} style={{display:"flex",alignItems:"center",gap:10,background:"#0d0f1a",borderRadius:8,padding:"8px 10px",border:"1px solid #1e2235",marginBottom:6}}>
                <div style={{flex:1}}>
                  <div style={{fontWeight:600,color:"#e8eaf0",fontSize:12}}>{p.name}</div>
                  <div style={{fontSize:11,color:"#6b7280"}}>{p.materials.length} items</div>
                </div>
                <input type="number" min="0" value={selProducts[p.id]||""} onChange={e=>setSelProducts(prev=>({...prev,[p.id]:parseFloat(e.target.value)||0}))}
                  placeholder="Qty" style={{width:56,padding:"4px 6px",fontSize:12,textAlign:"center"}}/>
              </div>
            ))}
            {Object.values(selProducts).some(q=>q>0)&&(
              <button className="btn-primary" onClick={applyProductDemand} style={{width:"100%",padding:10,marginTop:6,fontSize:13}}><Sparkles size={16} style={{flexShrink:0}} /> Fill Order Lines</button>
            )}
          </div>

          {/* Manual add */}
          <div className="card" style={{padding:0,overflow:"hidden"}}>
            <div style={{padding:"12px 14px 8px",borderBottom:"1px solid #1e2235"}}>
              <div style={{fontWeight:700,color:"#fff",fontSize:13,marginBottom:8}}><Plus size={16} style={{flexShrink:0}} /> Add Items Manually</div>
              <input placeholder="Search items…" value={searchItem} onChange={e=>setSearchItem(e.target.value)} style={{fontSize:12,marginBottom:8}}/>
              <div style={{display:"flex",gap:4,flexWrap:"wrap"}}>
                {allCats.map(cat=>(
                  <button key={cat} onClick={()=>setSelCats(prev=>prev.includes(cat)?prev.filter(c=>c!==cat):[...prev,cat])}
                    style={{padding:"2px 8px",borderRadius:20,border:"none",fontSize:10,fontWeight:700,cursor:"pointer",
                      background:selCats.includes(cat)?(catBg[cat]||"#2a1a1a"):"#1e2235",color:selCats.includes(cat)?(catColors[cat]||"#f87171"):"#6b7280",border:selCats.includes(cat)?"1px solid color-mix(in srgb, currentColor 15%, transparent)":"1px solid transparent"}}>
                    {cat}
                  </button>
                ))}
                {selCats.length>0&&<button onClick={()=>setSelCats([])} style={{padding:"2px 8px",borderRadius:20,border:"none",fontSize:10,fontWeight:700,cursor:"pointer",background:"#3a1a1a",color:"#ff4757"}}><X size={16} style={{flexShrink:0}} /></button>}
              </div>
            </div>
            <div style={{maxHeight:300,overflowY:"auto"}}>
              {availableItems.length===0&&<div style={{textAlign:"center",padding:"20px",color:"#4a5070",fontSize:12}}>{searchItem||selCats.length>0?"No items match":"All items already in order"}</div>}
              {availableItems.map(item=>(
                <div key={item.id} onClick={()=>addItem(item.id)}
                  style={{display:"flex",alignItems:"center",gap:8,padding:"8px 14px",borderBottom:"1px solid #181d2e",cursor:"pointer",transition:"background .1s"}}
                  onMouseEnter={e=>e.currentTarget.style.background="#1c2135"}
                  onMouseLeave={e=>e.currentTarget.style.background="transparent"}>
                  <div style={{flex:1,minWidth:0}}>
                    <div style={{fontWeight:600,color:"#e8eaf0",fontSize:12,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{item.name}</div>
                    <div style={{display:"flex",gap:5,marginTop:2}}>
                      <span className={tagClass(item.category)} style={{fontSize:10}}>{item.category}</span>
                      <span style={{fontSize:11,color:"#6b7280"}}>{item.unit}</span>
                    </div>
                  </div>
                  <span style={{fontWeight:700,color:"#f59e0b",fontSize:12,flexShrink:0}}>₹{item.cost.toLocaleString("en-IN")}</span>
                  <span style={{background:"linear-gradient(135deg,#5d7cff,#8b5cf6)",color:"#fff",borderRadius:6,width:20,height:20,display:"flex",alignItems:"center",justifyContent:"center",fontSize:14,fontWeight:700,flexShrink:0}}>+</span>
                </div>
              ))}
            </div>
          </div>
        </div>

        {/* RIGHT: order lines */}
        <div className="card" style={{padding:0,overflow:"hidden"}}>
          <div style={{padding:"12px 18px",borderBottom:"1px solid #1e2235",display:"flex",justifyContent:"space-between",alignItems:"center"}}>
            <div>
              <div style={{fontWeight:700,color:"#fff",fontSize:14}}><ClipboardList size={16} style={{flexShrink:0}} /> Order Lines</div>
              <div style={{fontSize:11,color:"#6b7280",marginTop:1}}>{orderLines.length} items · adjust qty &amp; cost</div>
            </div>
            {orderLines.length>0&&<div style={{fontWeight:800,color:"#f59e0b",fontSize:16}}>₹{fmt(grandTotal)}</div>}
          </div>

          {orderLines.length===0?(
            <div style={{textAlign:"center",padding:"50px 20px",color:"#4a5070"}}>
              <div style={{fontSize:36,marginBottom:10}}><ShoppingCart size={16} style={{flexShrink:0}} /></div>
              <div style={{fontSize:14,fontWeight:600,marginBottom:4}}>No items yet</div>
              <div style={{fontSize:12}}>Auto-fill or add items from the left</div>
            </div>
          ):(
            <>
              <div style={{display:"grid",gridTemplateColumns:"1fr 70px 100px 100px 80px 30px",padding:"7px 14px",background:"#0d0f1a",borderBottom:"1px solid #1e2235"}}>
                {["Item","Unit","Qty","Unit Cost","Total",""].map((h,i)=>(
                  <div key={i} style={{fontSize:10,fontWeight:700,color:"#6b7280",textTransform:"uppercase",textAlign:i>=2?"right":"left"}}>{h}</div>
                ))}
              </div>
              <div style={{maxHeight:"55vh",overflowY:"auto"}}>
                {sortedLines.map((l,idx)=>{
                  const item=items.find(i=>i.id===l.itemId); if(!item) return null;
                  return (
                    <div key={l.itemId} style={{display:"grid",gridTemplateColumns:"1fr 70px 100px 100px 80px 30px",padding:"8px 14px",borderBottom:"1px solid #181d2e",alignItems:"center",background:idx%2?"#0d0f1a":"transparent"}}
                      onMouseEnter={e=>e.currentTarget.style.background="#1c2135"}
                      onMouseLeave={e=>e.currentTarget.style.background=idx%2?"#0d0f1a":"transparent"}>
                      <div style={{minWidth:0}}>
                        <div style={{fontWeight:600,color:"#e8eaf0",fontSize:12,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{item.name}</div>
                        <span className={tagClass(item.category)} style={{fontSize:10}}>{item.category}</span>
                      </div>
                      <div style={{color:"#9ba3c0",fontSize:11,textAlign:"right"}}>{item.unit}</div>
                      <div style={{display:"flex",justifyContent:"flex-end"}}>
                        <div style={{display:"flex",alignItems:"center",background:"#1e2235",borderRadius:6,overflow:"hidden"}}>
                          <button onClick={()=>updateLine(l.itemId,"orderQty",Math.max(1,l.orderQty-1))} style={{width:22,height:26,background:"none",border:"none",color:"#9ba3c0",fontSize:14,cursor:"pointer",fontWeight:700}}>−</button>
                          <input type="number" min="0" value={l.orderQty} onChange={e=>updateLine(l.itemId,"orderQty",e.target.value)}
                            style={{width:40,height:26,textAlign:"center",background:"none",border:"none",borderLeft:"1px solid #2e3350",borderRight:"1px solid #2e3350",color:"#fff",fontWeight:700,fontSize:12,padding:0}}/>
                          <button onClick={()=>updateLine(l.itemId,"orderQty",l.orderQty+1)} style={{width:22,height:26,background:"none",border:"none",color:"#9ba3c0",fontSize:14,cursor:"pointer",fontWeight:700}}>+</button>
                        </div>
                      </div>
                      <div style={{display:"flex",justifyContent:"flex-end"}}>
                        <div style={{position:"relative"}}>
                          <span style={{position:"absolute",left:6,top:"50%",transform:"translateY(-50%)",color:"#6b7280",fontSize:11,pointerEvents:"none"}}>₹</span>
                          <input type="number" min="0" value={l.unitCost} onChange={e=>updateLine(l.itemId,"unitCost",e.target.value)}
                            style={{width:82,paddingLeft:16,textAlign:"right",fontSize:12,color:"#f59e0b",fontWeight:600}}/>
                        </div>
                      </div>
                      <div style={{textAlign:"right",fontWeight:700,color:"#e8eaf0",fontSize:12}}>₹{fmt(l.orderQty*l.unitCost)}</div>
                      <div style={{display:"flex",justifyContent:"flex-end"}}>
                        <button onClick={()=>removeItem(l.itemId)} style={{width:24,height:24,background:"#2a1a1a",border:"1px solid #ff4757",color:"#ff4757",borderRadius:5,fontSize:12,cursor:"pointer",display:"flex",alignItems:"center",justifyContent:"center",fontWeight:700}}><X size={16} style={{flexShrink:0}} /></button>
                      </div>
                    </div>
                  );
                })}
              </div>
              {/* Footer totals */}
              <div style={{padding:"10px 18px 14px",borderTop:"2px solid #2e3350",background:"#0d0f1a"}}>
                <div style={{display:"flex",gap:8,flexWrap:"wrap",marginBottom:10}}>
                  {(()=>{
                    const catTotals={};
                    sortedLines.forEach(l=>{ const item=items.find(i=>i.id===l.itemId);if(!item)return; catTotals[item.category]=(catTotals[item.category]||0)+l.orderQty*l.unitCost; });
                    return Object.entries(catTotals).sort((a,b)=>b[1]-a[1]).map(([cat,val])=>(
                      <div key={cat} style={{background:catBg[cat]||"#2a1a1a",borderRadius:8,padding:"4px 10px",display:"flex",gap:8,alignItems:"center"}}>
                        <span style={{fontSize:10,fontWeight:700,color:catColors[cat]||"#f87171"}}>{cat}</span>
                        <span style={{fontSize:12,fontWeight:800,color:"#e8eaf0"}}>₹{fmt(val)}</span>
                      </div>
                    ));
                  })()}
                </div>
                <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",paddingTop:10,borderTop:"1px solid #1e2235"}}>
                  <button onClick={()=>{if(window.confirm("Clear all lines?"))setOrderLines([]);}} style={{background:"#2a1a1a",border:"1px solid #ff4757",color:"#ff4757",padding:"6px 12px",borderRadius:7,fontSize:12,fontWeight:600,cursor:"pointer"}}><Trash2 size={16} style={{flexShrink:0}} /> Clear Lines</button>
                  <div style={{textAlign:"right"}}>
                    <div style={{fontSize:11,color:"#6b7280",fontWeight:600,textTransform:"uppercase",letterSpacing:".05em",marginBottom:3}}>Grand Total</div>
                    <div style={{fontSize:26,fontWeight:800,color:"#f59e0b"}}>₹{fmt(grandTotal)}</div>
                  </div>
                </div>
              </div>
            </>
          )}
        </div>
      </div>
    </div>
  );
}
