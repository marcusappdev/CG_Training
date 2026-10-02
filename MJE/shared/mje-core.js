/* =====================================================================
   MJE Hubs - shared helpers, used by BOTH apps (Safety Hub + Service Hub)
   Plain script (no build step). Load order in each app:
     shared/mje-config.js -> shared/mje-core.js -> the app's own <script>
   The app defines `const CONFIG = { ...MJE_TENANT, LISTS, SCOPES, ... }`;
   these helpers read CONFIG when they are called, not when loaded.
   No register data in this file - the repo is public.
   ===================================================================== */

/* ---------------- where things live ---------------- */
const MJE_SHARED = new URL(".", document.currentScript.src).href;    // .../MJE/shared/
const MJE_ROOT   = new URL("..", MJE_SHARED).href;                     // .../MJE/
const MJE_APPS = [
  { key:"safety",  name:"Safety Hub",  href: MJE_ROOT },
  { key:"service", name:"Service Hub", href: MJE_ROOT + "service/" },
];
// Header app switcher: Safety Hub | Service Hub
function mjeAppSwitcher(current){
  return `<div class="apps" role="navigation" aria-label="MJE apps">${MJE_APPS.map(a =>
    `<a href="${a.href}" class="${a.key === current ? "on" : ""}"${a.key === current ? ' aria-current="page"' : ""}>${a.name}</a>`).join("")}</div>`;
}

/* ---------------- helpers ---------------- */
const $ = s => document.querySelector(s);
const esc = v => String(v ?? "").replace(/[&<>"']/g, c => ({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c]));
const pad = n => String(n).padStart(2, "0");
// Local-date YYYY-MM-DD (never toISOString — avoids the UTC+10 shift)
const ymd = d => `${d.getFullYear()}-${pad(d.getMonth()+1)}-${pad(d.getDate())}`;
const hhmm = d => `${pad(d.getHours())}:${pad(d.getMinutes())}`;          // local HH:mm
const TODAY = ymd(new Date());
const YEAR = TODAY.slice(0,4);
const addDays = (s, n) => { const [y,m,d] = s.split("-").map(Number); return ymd(new Date(y, m-1, d+n)); };
const addMonths = (s, n) => { const [y,m,d] = s.split("-").map(Number); return ymd(new Date(y, m-1+n, d)); };
const fmtDate = s => { if (!s) return ""; const [y,m,d] = s.split("-"); return `${d}/${m}/${y}`; };
const num = v => (v === "" || v == null || isNaN(+v)) ? 0 : +v;
const kina = v => "K " + num(v).toLocaleString("en-AU", {minimumFractionDigits:0, maximumFractionDigits:0});
const money = v => v === undefined || v === null || v === "" ? "" : kina(v);
function toast(msg){ const t=$("#toast"); t.textContent=msg; t.classList.add("on"); clearTimeout(t._h); t._h=setTimeout(()=>t.classList.remove("on"),2600); }

/* ---------------- auth (MSAL, PKCE) ---------------- */
// Both apps share the sign-in: same Entra app, same browser storage.
let msalApp, account;
async function initAuth(){
  msalApp = new msal.PublicClientApplication({
    auth:{ clientId:CONFIG.CLIENT_ID, authority:`https://login.microsoftonline.com/${CONFIG.TENANT_ID}`, redirectUri: location.origin + location.pathname },
    cache:{ cacheLocation:"localStorage" }
  });
  await msalApp.initialize();
  const res = await msalApp.handleRedirectPromise();
  account = res?.account || msalApp.getAllAccounts()[0];
  if (account) msalApp.setActiveAccount(account);
}
async function token(){
  try { return (await msalApp.acquireTokenSilent({ scopes:CONFIG.SCOPES, account })).accessToken; }
  catch(e){ await msalApp.acquireTokenRedirect({ scopes:CONFIG.SCOPES, account }); }
}
const signIn  = () => msalApp.loginRedirect({ scopes:CONFIG.SCOPES });
const signOut = () => msalApp.logoutRedirect({ account });

/* ---------------- Graph data layer ---------------- */
let siteId;
async function graph(path, opts = {}){
  const r = await fetch(path.startsWith("http") ? path : "https://graph.microsoft.com/v1.0" + path, {
    ...opts, headers:{ Authorization:"Bearer " + await token(), "Content-Type":"application/json", ...(opts.headers||{}) }
  });
  if (!r.ok){ let m = r.status + " " + r.statusText; try{ m = (await r.json()).error.message; }catch{} const e = new Error(m); e.status = r.status; throw e; }
  return r.status === 204 ? null : r.json();
}
async function getSite(){ siteId = (await graph(`/sites/${CONFIG.SP_HOST}:${CONFIG.SITE_PATH}`)).id; }

// Every item of a SharePoint list, following @odata.nextLink page by page
async function listItems(listName){
  let url = `/sites/${siteId}/lists/${listName}/items?expand=fields&$top=999`, out = [];
  while (url){ const j = await graph(url); out.push(...j.value.map(i => ({ id:i.id, ...i.fields }))); url = j["@odata.nextLink"]; }
  return out;
}
// True when Graph says the list doesn't exist (not set up yet) — apps show a "not set up" message instead of failing
const isMissingList = e => e.status === 404 || /not (be )?found|does not exist/i.test(e.message);
// True when SharePoint refused the change for permission reasons
const isForbidden = e => e.status === 403 || /access denied|forbidden|not authorized/i.test(e.message);

// Create/update with field fallback: strip any column SharePoint doesn't recognise and retry
async function saveListItem(listName, id, fields){
  const base = `/sites/${siteId}/lists/${listName}/items`;
  let f = { ...fields }, stripped = [];
  for (let attempt = 0; attempt < 8; attempt++){
    try {
      const res = id ? await graph(`${base}/${id}/fields`, { method:"PATCH", body:JSON.stringify(f) })
                     : await graph(base, { method:"POST", body:JSON.stringify({ fields:f }) });
      if (stripped.length) console.warn("Columns not on list, skipped:", stripped);
      return id ? { id, ...res } : { id:res.id, ...res.fields };
    } catch(e){
      const m = /Field '([^']+)' is not recognized/i.exec(e.message);
      if (m && m[1] in f){ stripped.push(m[1]); delete f[m[1]]; continue; }
      throw e;
    }
  }
}

/* ---------------- photos & files ---------------- */
// Resize to max 1600 px JPEG before upload (falls back to the original file)
async function shrink(file){
  try {
    const img = await createImageBitmap(file), s = Math.min(1, 1600 / Math.max(img.width, img.height));
    const c = document.createElement("canvas"); c.width = Math.round(img.width * s); c.height = Math.round(img.height * s);
    c.getContext("2d").drawImage(img, 0, 0, c.width, c.height);
    return await new Promise(r => c.toBlob(b => r(b || file), "image/jpeg", 0.8));
  } catch(e){ return file; }
}
// Drive id of a document library (null if the library isn't there)
async function libraryDrive(listName){
  try { return (await graph(`/sites/${siteId}/lists/${listName}/drive`)).id; } catch(e){ return null; }
}
// PUT a file into <library>/<folder>/<name>; the folder is created by Graph if needed
async function uploadToDrive(driveId, folder, name, blob, contentType = "image/jpeg"){
  const r = await fetch(`https://graph.microsoft.com/v1.0/drives/${driveId}/root:/${encodeURIComponent(folder)}/${name}:/content`,
    { method:"PUT", headers:{ Authorization:"Bearer " + await token(), "Content-Type":contentType }, body: blob });
  if (!r.ok){ const e = new Error(r.status === 403 ? "you don't have permission to add photos" : `upload failed (${r.status})`); e.status = r.status; throw e; }
  return r.json();
}
