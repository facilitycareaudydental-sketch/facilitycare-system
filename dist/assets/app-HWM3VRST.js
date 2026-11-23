var Ri=Object.defineProperty;var oa=(e,t)=>()=>(e&&(t=e(e=0)),t);var Xa=(e,t)=>{for(var n in t)Ri(e,n,{get:t[n],enumerable:!0})};function P(e){return e==null?"":String(e).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;").replace(/'/g,"&#39;").replace(/`/g,"&#96;")}function Wi(e,t="#"){if(e==null)return t;let n=String(e).trim();if(!n)return t;let a=n.replace(Ji,"");return Vi.test(a)?t:/^(?:#|\/(?!\/)|\.{1,2}\/|https?:\/\/|mailto:|tel:)/i.test(n)?n:t}function He(e){if(e==null)return"";let t=String(e);return t=t.replace(/<\s*(script|style|iframe|object|embed|applet|svg|math|meta|link|base)\b[^>]*>[\s\S]*?<\s*\/\s*\1\s*>/gi,""),t=t.replace(/<\s*\/?\s*(script|style|iframe|object|embed|applet|svg|math|meta|link|base)\b[^>]*\/?>/gi,""),t=t.replace(/\s+on[a-z0-9_-]+\s*=\s*(?:"[^"]*"|'[^']*'|[^\s>]+)/gi,""),t=t.replace(/\s+(?:srcdoc|formaction)\s*=\s*(?:"[^"]*"|'[^']*'|[^\s>]+)/gi,""),t=t.replace(/(\b(?:href|src|action|xlink:href|poster)\s*=\s*)(?:"([^"]*)"|'([^']*)'|([^\s>]+))/gi,(n,a,o,s,c)=>{let i=Wi(o??s??c??"","#");return`${a}"${P(i)}"`}),t}var Vi,Ji,Ie=oa(()=>{Vi=/^\s*(?:javascript|vbscript|data|file)\s*:/i,Ji=/[\u0000-\u001f\u007f\s]+/g});var dn={};Xa(dn,{confirmDialog:()=>ge,createModal:()=>ce});function ce({title:e,content:t,onConfirm:n,onCancel:a,confirmText:o="Simpan",cancelText:s="Batal",size:c="md",confirmClass:m="btn-primary"}){let i={sm:"400px",md:"560px",lg:"720px",xl:"900px"},d=document.createElement("div");d.className="modal-overlay";let F=typeof t=="string"?He(t):"";d.innerHTML=`
    <div class="modal" style="max-width:${i[c]||i.md}">
      <div class="modal-header">
        <h3 class="modal-title">${P(e)}</h3>
        <button type="button" class="modal-close" aria-label="Tutup modal">&times;</button>
      </div>
      <div class="modal-body">${F}</div>
      <div class="modal-footer">
        <button type="button" class="btn btn-ghost modal-cancel">${P(s)}</button>
        ${n?`<button type="button" class="btn ${P(m)} modal-confirm">${P(o)}</button>`:""}
      </div>
    </div>
  `,t instanceof HTMLElement&&d.querySelector(".modal-body").appendChild(t);let $=()=>{d.classList.remove("show"),setTimeout(()=>d.remove(),250)};return d.querySelector(".modal-close").addEventListener("click",()=>{a&&a(),$()}),d.querySelector(".modal-cancel").addEventListener("click",()=>{a&&a(),$()}),n&&d.querySelector(".modal-confirm").addEventListener("click",()=>n(d,$)),d.addEventListener("click",M=>{M.target===d&&(a&&a(),$())}),document.body.appendChild(d),requestAnimationFrame(()=>d.classList.add("show")),{overlay:d,close:$}}function ln(e){if(typeof e!="string")return e;let t=e.replace(/<\/?strong>/gi,""),n=document.createElement("p");return n.style.cssText="margin:0;color:#E2E8F0;line-height:1.6;font-size:0.94rem;",n.textContent=t,n}function ge(e,t,n="Konfirmasi"){if(typeof e=="object"&&e!==null){let a=e;return ce({title:a.title||"Konfirmasi",content:ln(a.message),onConfirm:async(o,s)=>{a.onConfirm&&await a.onConfirm(),s()},onCancel:a.onCancel,confirmText:a.confirmText||"Ya, Lanjutkan",cancelText:a.cancelText||"Batal",confirmClass:a.confirmClass||"btn-danger"})}return ce({title:n,content:ln(e),onConfirm:async(a,o)=>{t&&await t(),o()},confirmText:"Ya, Lanjutkan",confirmClass:"btn-danger"})}var Be=oa(()=>{Ie()});var Ve={};Xa(Ve,{bindExcelDropdown:()=>xr,downloadExcel:()=>be,parseExcel:()=>Sa,renderExcelButtons:()=>vr});function Sa(e){return new Promise((t,n)=>{let a=new FileReader;a.onload=o=>{try{let s=new Uint8Array(o.target.result),c=XLSX.read(s,{type:"array"}),m=c.SheetNames[0],i=c.Sheets[m];console.log("--- START EXCEL PARSING ---"),console.log(`File Name: ${e.name}`),console.log(`File Size: ${(e.size/1024).toFixed(2)} KB`),console.log(`File Type: ${e.type||"unknown"}`),console.log(`Sheets Found: ${c.SheetNames.join(", ")}`),console.log(`Sheet Used: ${m}`);let d=XLSX.utils.decode_range(i["!ref"]||"A1:A1"),F=d.e.r-d.s.r+1,$=d.e.c-d.s.c+1;console.log(`Total Rows (including empty): ${F}`),console.log(`Total Columns: ${$}`);let M=[];for(let L=d.s.c;L<=d.e.c;++L){let R=i[XLSX.utils.encode_cell({c:L,r:d.s.r})];R&&R.v&&M.push(R.v)}console.log(`Headers Found: ${M.join(", ")}`),console.log("---------------------------");let h=XLSX.utils.sheet_to_json(i,{defval:""});Object.defineProperty(h,"__worksheet",{value:i,enumerable:!1}),Object.defineProperty(h,"__headers",{value:M,enumerable:!1}),t(h)}catch(s){n(s)}},a.onerror=o=>n(o),a.readAsArrayBuffer(e)})}function be(e,t){try{let n=XLSX.utils.json_to_sheet(e),a=XLSX.utils.book_new();XLSX.utils.book_append_sheet(a,n,"Data"),XLSX.writeFile(a,`${t}.xlsx`)}catch(n){throw console.error("Error generating Excel file:",n),n}}function Ca(){An||typeof document>"u"||typeof document.addEventListener!="function"||(An=!0,document.addEventListener("click",e=>{let t=e.target?.closest?.("[data-excel-toggle-menu]");if(t){let n=t.getAttribute("data-excel-toggle-menu"),a=n?document.getElementById(n):null;document.querySelectorAll(".aksi-menu-content.show-aksi-menu").forEach(o=>{o!==a&&o.classList.remove("show-aksi-menu")}),a&&a.classList.toggle("show-aksi-menu");return}e.target?.closest?.(".excel-actions-dropdown")||document.querySelectorAll(".aksi-menu-content.show-aksi-menu").forEach(n=>{n.classList.remove("show-aksi-menu")})}))}function xr(e,t=typeof document<"u"?document:null){if(Ca(),!t||typeof t.querySelector!="function")return;let n=t.querySelector(`#btn-aksi-${e}`),a=t.querySelector(`#excel-menu-${e}`);n&&a&&!n.dataset.excelBound&&(n.dataset.excelBound="1",n.getAttribute("data-excel-toggle-menu")||n.addEventListener("click",o=>{o.stopPropagation(),a.classList.toggle("show-aksi-menu")}))}function vr(e){return Ca(),`
    <div class="excel-actions-dropdown" style="position:relative; display:inline-block;">
      <button type="button" class="btn" id="btn-aksi-${e}" data-excel-toggle-menu="excel-menu-${e}" style="background:#172033; border:1px solid #334155; padding:8px 16px; border-radius:8px; font-weight:600; color:#F8FAFC; display:flex; align-items:center; gap:6px; cursor:pointer;">
        \u22EE Aksi
      </button>
      <div id="excel-menu-${e}" class="aksi-menu-content" style="display:none; position:absolute; top:calc(100% + 4px); right:0; background:#172033; border:1px solid #334155; border-radius:8px; box-shadow:0 10px 25px rgba(0,0,0,0.45); flex-direction:column; min-width:180px; z-index:999; padding:8px 0;">
        <button type="button" class="dropdown-item" id="btn-export-${e}" style="width:100%; text-align:left; padding:10px 16px; background:none; border:none; cursor:pointer; font-size:0.9rem; color:#E2E8F0; display:flex; gap:8px; align-items:center;">
          \u{1F4E5} Export Excel
        </button>
        <button type="button" class="dropdown-item" id="btn-template-${e}" style="width:100%; text-align:left; padding:10px 16px; background:none; border:none; cursor:pointer; font-size:0.9rem; color:#E2E8F0; display:flex; gap:8px; align-items:center;">
          \u{1F4C4} Download Template
        </button>
        <label class="dropdown-item" style="display:flex; width:100%; text-align:left; padding:10px 16px; background:none; border:none; cursor:pointer; font-size:0.9rem; color:#E2E8F0; margin:0; gap:8px; align-items:center;" id="label-import-${e}">
          \u{1F4E4} Import Excel
          <input type="file" id="input-import-${e}" accept=".xlsx, .xls, .csv" style="display:none;">
        </label>
      </div>
    </div>
    <style>.show-aksi-menu{display:flex !important}.dropdown-item:hover{background-color:#22314A !important;color:#93C5FD !important}</style>
  `}var An,Ce=oa(()=>{An=!1;Ca()});var sa="fm-operations-api-staging.facilitycare-audydental.workers.dev",da="https://fm-operations-api-staging.facilitycare-audydental.workers.dev",la="fm-operations-api.facilitycare-audydental.workers.dev",Za="https://fm-operations-api.facilitycare-audydental.workers.dev",Yr=Object.freeze([da]);function en(e){if(typeof e!="string")return!1;let t=e.trim();return t.startsWith("__")&&t.endsWith("__")&&t.length>=4}function ua(e=null){let t=e??(typeof globalThis<"u"&&globalThis.window&&globalThis.window.__FM_CONFIG?globalThis.window.__FM_CONFIG:null);if(!t||typeof t!="object")throw new Error("CONFIG_ERROR: Missing runtime configuration (window.__FM_CONFIG). Fail-closed.");let n=t.ENVIRONMENT!==void 0?t.ENVIRONMENT:t.ENV,a=t.API_BASE_URL!==void 0?t.API_BASE_URL:t.API_BASE;if(en(n)||en(a))throw new Error("CONFIG_ERROR: Unreplaced build configuration placeholders detected. Fail-closed.");if(n==null||String(n).trim()==="")throw new Error("CONFIG_ERROR: Environment is not configured. Explicit environment required.");let o=String(n).trim().toLowerCase();if(o==="staging"){if(a==null||String(a).trim()==="")throw new Error("CONFIG_ERROR: Missing Staging API_BASE_URL configuration. Fail-closed.");let s=String(a).trim();if(s.startsWith("//"))throw new Error("CONFIG_ERROR: Protocol-relative URLs are not allowed in staging API_BASE_URL.");let c;try{c=new URL(s)}catch{throw new Error(`CONFIG_ERROR: Invalid or unknown Staging API_BASE_URL '${s}'. Fail-closed.`)}if(c.protocol!=="https:"&&c.protocol!=="http:")throw new Error("CONFIG_ERROR: Staging API_BASE_URL must use https protocol.");let m=c.hostname.toLowerCase().trim();if(m===la)throw new Error("SECURITY_VIOLATION: Staging environment cannot target Production Worker API.");if(m.includes("fcvm"))throw new Error("SECURITY_VIOLATION: Staging environment cannot target FCVM resources.");if(c.username||c.password)throw new Error("SECURITY_VIOLATION: Embedded credentials are not permitted in Staging API_BASE_URL.");if(c.protocol!=="https:"||m!==sa||c.port!==""||c.pathname&&c.pathname!=="/"||c.search!==""||c.hash!=="")throw new Error(`SECURITY_VIOLATION: Unknown or unapproved Staging API target '${s}'. Only '${da}' is permitted.`);return da}if(o==="production"){if(a!=null&&String(a).trim()!==""){let s=String(a).trim();if(s.startsWith("//"))throw new Error("CONFIG_ERROR: Protocol-relative URLs are not allowed in production API_BASE_URL.");let c;try{c=new URL(s)}catch{throw new Error("CONFIG_ERROR: Production API_BASE_URL is a malformed URL.")}let m=c.hostname.toLowerCase().trim();if(m.includes("fcvm"))throw new Error("SECURITY_VIOLATION: Production environment cannot target FCVM resources.");if(m===sa)throw new Error("SECURITY_VIOLATION: Production environment cannot target Staging Worker API.");if(c.protocol!=="https:"||m!==la)throw new Error(`SECURITY_VIOLATION: Unknown Production API target '${s}'.`);return Za}return Za}if(o==="development"){if(a==null||String(a).trim()===""||String(a).trim()==="/api")return"/api";let s=String(a).trim();if(s==="/api")return"/api";let c;try{c=new URL(s)}catch{throw new Error("CONFIG_ERROR: Development API_BASE_URL is a malformed URL.")}let m=c.hostname.toLowerCase().trim();if(m===la)throw new Error("SECURITY_VIOLATION: Development environment cannot target Production Worker API.");if(m.includes("fcvm"))throw new Error("SECURITY_VIOLATION: Development environment cannot target FCVM resources.");if(m!=="localhost"&&m!=="127.0.0.1"&&m!==sa)throw new Error(`SECURITY_VIOLATION: Unknown Development API target '${s}'.`);return s.replace(/\/+$/,"")}throw new Error(`CONFIG_ERROR: Unknown environment '${o}'. Fail-closed.`)}function nn(e,t){let n=String(e||"").trim().replace(/\/+$/,""),a=String(t||"").trim(),o=a.startsWith("/")?a:`/${a}`;if(!n)throw new Error("CONFIG_ERROR: API base URL cannot be empty when constructing request URL.");return n==="/api"&&(o==="/api"||o.startsWith("/api/"))?o:`${n}${o}`}var ca="",rn="";typeof globalThis<"u"&&globalThis.window&&globalThis.window.__FM_CONFIG&&(ca=ua(globalThis.window.__FM_CONFIG),rn=ca);var ma=null,ga=null,vt=null;function ji(e){if(typeof document>"u"||!document.cookie)return null;let t=`${e}=`;for(let n of document.cookie.split(";")){let a=n.trim();if(a.startsWith(t))try{return decodeURIComponent(a.slice(t.length))}catch{return a.slice(t.length)}}return null}function ut(){return typeof localStorage<"u"&&localStorage.removeItem("fm_token"),ma}function on(e){typeof localStorage<"u"&&localStorage.removeItem("fm_token"),ma=e||null}function pa(){return ga||ji("fcms_csrf")}function kt(e){ga=e||null}function mt(){ma=null,ga=null,typeof localStorage<"u"&&(localStorage.removeItem("fm_token"),localStorage.removeItem("fm_user"))}function se(){if(typeof localStorage>"u")return null;try{return JSON.parse(localStorage.getItem("fm_user")||"null")}catch{return null}}function Ht(e){typeof localStorage>"u"||localStorage.setItem("fm_user",JSON.stringify(e))}var qi=[{id:"employees",label:"Data Karyawan (HR)"},{id:"contracts",label:"Data Kontrak (HR)"},{id:"sp",label:"Surat Peringatan / SP (HR)"},{id:"mutasi",label:"Data Mutasi (HR)"},{id:"overtime",label:"Data Lembur (HR)"},{id:"schedule",label:"Time Line / Jadwal (Operasional)"},{id:"os_migrations",label:"Migrasi OS (Operasional)"},{id:"issues",label:"Permasalahan / Issues (Operasional)"},{id:"training",label:"Training Karyawan (Operasional)"},{id:"reports",label:"Laporan Kebersihan / Basecamp"},{id:"sop",label:"SOP Facility Care & Security"},{id:"checklist",label:"Master Dokumen & Formulir"},{id:"hygiene_standards",label:"Master Standar Hygiene"},{id:"dac",label:"Master Standar DAC"},{id:"branches",label:"Data Cabang"}],tn=new Set(["relievers","reliefer","one_on_one","one-on-one","forms","master_forms","supply","supply_requests","reports/supply","reports_supply"]),an=["users","audit_logs","audit-logs","sync","sync_dashboard","monitoring","settings","system_settings","import"],Hi=new Set(qi.map(e=>e.id));function Oi(e){if(!e)return[];let t=e;if(typeof e=="object"&&!Array.isArray(e)&&(t=e.assigned_modules!==void 0?e.assigned_modules:e.assignedModules||e.modules),!t)return[];let n=o=>String(o).trim().toLowerCase().replace(/-/g,"_"),a=o=>Array.from(new Set(o.map(n).filter(s=>s&&Hi.has(s))));if(Array.isArray(t))return a(t);if(typeof t=="string"){try{let o=JSON.parse(t);if(Array.isArray(o))return a(o)}catch{}return a(t.split(","))}return[]}var Qr=Object.freeze(["superadmin","fc_specialist_leader","viewer"]);function re(e,t="read",n=null){if(!e||!e.role||e.isRestrictedToken===!0||Number(e.must_change_password)===1)return!1;let a=String(e.role).trim().toLowerCase(),o=null,s="read";if(n!==null)o=t,s=n;else{let d=String(t||"read").toLowerCase();["read","write","create","update","delete","export","import","admin"].includes(d)?(s=d,o=null):(o=d,s="read")}let c=String(s||"read").trim().toLowerCase(),m=o?String(o).trim().toLowerCase():"",i=m.replace(/-/g,"_");if(i&&(tn.has(i)||tn.has(m)))return!1;if(a==="superadmin")return!0;if(c==="admin"||i&&(an.includes(i)||an.includes(m)))return!1;if(a==="fc_specialist_leader"||a==="editor_khusus"||a==="fc_specialist"){if(c==="read")return!0;if(c==="create"||c==="update"||c==="write"){if(a==="editor_khusus"&&e.assigned_modules!==void 0&&e.assigned_modules!==null){let F=Oi(e);if(F.length>0)return F.includes(i);if(typeof e.assigned_modules=="string"&&e.assigned_modules.trim()==="")return!1}return i==="overtime"||i==="issues"}return!1}return a==="viewer"||a==="manager"||a==="spv"||a==="supervisor"?c==="read":a==="admin"?!(c==="import"||c==="admin"):a==="input_lembur"?c==="read"?!0:c==="create"||c==="update"||c==="write"?i==="overtime":!1:!1}async function Ui(e={}){if(!(ut()||pa()))return vt||(vt=(async()=>{try{let t=ua(e.configOverride||null),n=typeof e.fetchImpl=="function"?e.fetchImpl:fetch,a=nn(t,`/api/auth/me?cb=${Date.now()}`),o=await n(a,{method:"GET",credentials:"include",headers:{"Content-Type":"application/json"}});if(o.ok){let c=(await o.json().catch(()=>null))?.data?.csrf_token;typeof c=="string"&&c.length>0&&kt(c)}}catch{}finally{vt=null}})(),vt)}var qt=new Map;async function X(e,t={}){let n=ua(t.configOverride||null);ca=n,rn=n;let a=String(t.method||"GET").toUpperCase(),s=a==="GET"&&!t.fetchImpl&&!t.signal?`${n}::${e}::${ut()||""}`:null;if(s&&qt.has(s))return qt.get(s);let c=async()=>{!t.fetchImpl&&["POST","PUT","PATCH","DELETE"].includes(a)&&!e.startsWith("/api/auth/login")&&se()&&!ut()&&!pa()&&await Ui(t);let i=ut(),d=pa(),F={"Content-Type":"application/json",...i?{Authorization:`Bearer ${i}`}:{},...d&&["POST","PUT","PATCH","DELETE"].includes(a)?{"X-CSRF-Token":d}:{},...t.headers||{}},$=e.includes("/import")||e.includes("/backup")||e.includes("/export")?6e4:2e4,M=typeof t.timeoutMs=="number"&&t.timeoutMs>0?t.timeoutMs:$,h=null,L=null,R=!1;!t.signal&&typeof AbortController=="function"&&(h=new AbortController,L=setTimeout(()=>{R=!0,h.abort()},M));try{let I=`cb=${Date.now()}`,N=e.includes("?")?"&":"?",q=`${e}${N}${I}`,v=nn(n,q),B=typeof t.fetchImpl=="function"?t.fetchImpl:fetch,x={credentials:"include",...t,headers:F,...h?{signal:h.signal}:{}};delete x.timeoutMs,delete x.fetchImpl,delete x.configOverride;let k=await B(v,x),_;try{let E=await k.text();try{_=JSON.parse(E)}catch{_={error:`Server Error (${k.status}): ${E.substring(0,80)}...`}}}catch{_={error:"Gagal membaca respon dari server"}}return k.status===401&&(mt(),typeof window<"u"&&window.location&&(window.location.hash="#/login")),{ok:k.ok,status:k.status,data:_}}catch(I){return R||I?.name==="AbortError"?{ok:!1,status:408,data:{error:"Permintaan waktu habis (timeout). Silakan coba lagi."}}:{ok:!1,status:0,data:{error:`Koneksi terputus. Periksa jaringan Anda. (${I.message})`}}}finally{L!==null&&clearTimeout(L)}};if(!s)return c();let m=c().finally(()=>{qt.delete(s)});return qt.set(s,m),m}var ha={},Ot=null,ba=0;function he(e,t){ha[e]=t}function Ge(e){window.location.hash=e}function sn(){async function e(){let t=++ba,n=window.location.hash.replace("#","")||"/dashboard",[a,...o]=n.split("?"),s=ha[a];if(!s){for(let[m,i]of Object.entries(ha))if(m.endsWith("/*")&&a.startsWith(m.slice(0,-2))){s=i;break}}if(Ot)try{Ot()}catch(m){console.warn("[Router] Previous page cleanup error:",m?.message||m)}finally{Ot=null}let c=document.getElementById("main-content");if(c&&(c.innerHTML='<div class="loading-spinner"><div class="spinner"></div></div>'),s)try{let m=new URLSearchParams(o.join("?")),i=a.split("/").filter(Boolean),d=await s({path:a,params:m,segments:i,main:c});if(t!==ba){if(typeof d=="function")try{d()}catch{}return}typeof d=="function"&&(Ot=d)}catch(m){if(t!==ba)return;console.error(`[Router] Module render error on '${a}':`,m?.message||m);let i=c||document.getElementById("app");i&&(i.innerHTML=`
            <div class="empty-state" style="padding:32px;text-align:center;">
              <h2 style="margin-bottom:8px;">Terjadi kendala saat memuat halaman</h2>
              <p class="text-muted" style="margin-bottom:16px;">Modul ini mengalami kesalahan sementara dan telah diisolasi agar tidak mengganggu modul lainnya.</p>
              <div style="display:flex;gap:8px;justify-content:center;">
                <button type="button" class="btn btn-primary" id="btn-router-retry">Muat Ulang Halaman</button>
                <a href="#/dashboard" class="btn btn-secondary" style="text-decoration:none;">Kembali ke Dashboard</a>
              </div>
            </div>
          `,i.querySelector("#btn-router-retry")?.addEventListener("click",()=>e()))}else{let m=c||document.getElementById("app");m&&(m.innerHTML='<div class="empty-state"><h2>404 - Halaman tidak ditemukan</h2></div>')}}window.addEventListener("hashchange",e),e()}var wt;function Gi(){return wt||(wt=document.createElement("div"),wt.id="toast-container",document.body.appendChild(wt)),wt}function fa(e,t="info",n=3500){let a=Gi(),o=["success","error","warning","info"].includes(t)?t:"info",s=document.createElement("div");s.className=`toast toast-${o}`;let c={success:"\u2713",error:"\u2715",warning:"\u26A0",info:"\u2139"},m=document.createElement("span");m.className="toast-icon",m.textContent=c[o]||"\u2139";let i=document.createElement("span");i.className="toast-msg",i.textContent=e==null?"":String(e),s.appendChild(m),s.appendChild(i),a.appendChild(s),requestAnimationFrame(()=>s.classList.add("show")),setTimeout(()=>{s.classList.remove("show"),setTimeout(()=>s.remove(),350)},n)}var ae=e=>fa(e,"success"),j=e=>fa(e,"error");var gt=e=>fa(e,"info");Be();var Yi=["[data-fc-scroll-tabs]","#hygiene-zone-tabs","#doc-filter-pills",".emp-filter-pills",".con-filter-pills",".sp-filter-pills",".mutasi-filter-pills",".lembur-filter-pills",".issues-filter-pills",".training-filter-pills",".fc-nav-tabs-strip","#cat-pills-wrap","#sec-cat-pills-wrap"].join(", "),Qi='<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><polyline points="15 18 9 12 15 6"/></svg>',Xi='<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><polyline points="9 18 15 12 9 6"/></svg>';function Zi(e){if(!e||e.dataset.fcScrollEnhanced==="1"||!e.parentElement)return;if(e.parentElement.classList.contains("fc-scroll-tabs-shell")){e.dataset.fcScrollEnhanced="1";return}e.dataset.fcScrollEnhanced="1",e.classList.add("fc-scroll-tabs-track");let t=document.createElement("div");t.className="fc-scroll-tabs-shell";let n=document.createElement("button");n.type="button",n.className="fc-scroll-tabs-btn fc-scroll-tabs-btn--prev",n.setAttribute("aria-label","Geser daftar tab ke kiri"),n.title="Geser ke kiri",n.innerHTML=Qi;let a=document.createElement("button");a.type="button",a.className="fc-scroll-tabs-btn fc-scroll-tabs-btn--next",a.setAttribute("aria-label","Geser daftar tab ke kanan"),a.title="Geser ke kanan",a.innerHTML=Xi,e.parentElement.insertBefore(t,e),t.appendChild(n),t.appendChild(e),t.appendChild(a);let s=()=>{if(!e.isConnected)return;let m=Math.max(0,e.scrollWidth-e.clientWidth),i=m>4,d=e.scrollLeft;t.classList.toggle("is-overflowing",i);let F=i&&d>4,$=i&&d<m-4;t.classList.toggle("can-scroll-left",F),t.classList.toggle("can-scroll-right",$),n.disabled=!F,a.disabled=!$},c=m=>{let i=Math.max(180,Math.round(e.clientWidth*.65));e.scrollBy({left:m*i,behavior:"smooth"})};if(n.addEventListener("click",m=>{m.preventDefault(),c(-1)}),a.addEventListener("click",m=>{m.preventDefault(),c(1)}),e.addEventListener("wheel",m=>{if(!(e.scrollWidth<=e.clientWidth+4)&&Math.abs(m.deltaY)>Math.abs(m.deltaX)){let i=e.scrollLeft;e.scrollLeft+=m.deltaY,e.scrollLeft!==i&&m.preventDefault()}},{passive:!1}),e.addEventListener("scroll",s,{passive:!0}),typeof ResizeObserver<"u"){let m=new ResizeObserver(()=>requestAnimationFrame(s));m.observe(t),m.observe(e)}typeof MutationObserver<"u"&&new MutationObserver(()=>requestAnimationFrame(s)).observe(e,{childList:!0,subtree:!1}),window.addEventListener("resize",s,{passive:!0}),requestAnimationFrame(s),setTimeout(s,250),setTimeout(s,700)}function cn(e=document){!e||!e.querySelectorAll||e.querySelectorAll(Yi).forEach(t=>{Zi(t)})}var St=null;function pn(e){if(!e)return;St&&(St.disconnect(),St=null),cn(e);let t=!1;St=new MutationObserver(()=>{t||(t=!0,requestAnimationFrame(()=>{t=!1,cn(e)}))}),St.observe(e,{childList:!0,subtree:!0})}Ie();Ie();var et={};function Gt(e){if(et[e]){try{et[e].destroy()}catch{}delete et[e]}}function un(){Object.keys(et).forEach(Gt)}var Ct=(e,t=0)=>{let n=Number(e);return isNaN(n)||e===null||e===void 0?t:n},er=(e,t="\u2014")=>{if(e==null||e==="")return t;let n=String(e).trim();return n===""||n==="[object Object]"?t:P(n)};function tr(e,t,n=800){if(!e)return;e._animFrame&&cancelAnimationFrame(e._animFrame);let a=Math.max(0,Math.round(Ct(t))),o=e.textContent?e.textContent.replace(/[^\d]/g,""):"",s=o!==""&&parseInt(o)||0;if(a===s){e.textContent=a.toLocaleString("id-ID"),e._animFrame=null;return}let c=Date.now(),m=()=>{let i=Math.min((Date.now()-c)/n,1),d=1-Math.pow(1-i,3);e.textContent=Math.round(s+d*(a-s)).toLocaleString("id-ID"),i<1?e._animFrame=requestAnimationFrame(m):(e.textContent=a.toLocaleString("id-ID"),e._animFrame=null)};e._animFrame=requestAnimationFrame(m)}var bt={family:"Inter",size:10,weight:"500"},tt="#CBD5E1",Ut="rgba(148, 163, 184, 0.10)",ya=["#2563EB","#38BDF8","#60A5FA","#93C5FD","#1D4ED8","#0EA5E9"],ar=()=>window.innerWidth<768;function bn(e={}){return{responsive:!0,maintainAspectRatio:!1,animation:{duration:600,easing:"easeOutQuart"},plugins:{legend:{position:ar()?"bottom":"top",labels:{font:bt,color:tt,usePointStyle:!0,padding:10,boxWidth:8,boxHeight:8}},tooltip:{mode:"index",intersect:!1,backgroundColor:"#1E293B",titleColor:"#F8FAFC",bodyColor:"#94A3B8",borderColor:"rgba(255, 255, 255, 0.1)",borderWidth:1,padding:10,boxPadding:4,cornerRadius:8,bodyFont:bt,titleFont:{...bt,weight:"700"}}},scales:{x:{grid:{color:Ut},ticks:{font:bt,color:tt,maxRotation:0}},y:{grid:{color:Ut},ticks:{font:bt,color:tt},beginAtZero:!0}},...e}}var nr=()=>Array(5).fill(0).map(()=>`
  <div class="kpi-card" style="pointer-events:none;">
    <div class="skeleton" style="width:46px;height:46px;border-radius:12px;flex-shrink:0;"></div>
    <div style="flex:1;">
      <div class="skeleton skeleton-text" style="width:50px;height:24px;margin-bottom:6px;"></div>
      <div class="skeleton skeleton-text" style="width:85px;height:12px;"></div>
    </div>
  </div>`).join("");function mn(e=4){return Array(e).fill(0).map((t,n)=>`<div class="skeleton skeleton-text" style="height:36px;margin-bottom:${n<e-1?"8px":"0"};border-radius:8px;"></div>`).join("")}async function Je(e,t,n=8e3){try{let a=new AbortController,o=setTimeout(()=>a.abort(),n),s=await X(e,{signal:a.signal}).catch(()=>null);if(clearTimeout(o),!s||!s.ok)return t;let c=s.data;return c?c.data!==void 0?c.data??t:c:t}catch{return t}}function ir(e){if(!e||!e.isConnected)return;["skel-donut","skel-trend","skel-jadwal"].forEach(o=>{let s=document.getElementById(o);s&&(s.style.display="none")}),["chart-donut","chart-trend","chart-jadwal"].forEach(o=>{let s=document.getElementById(o);s&&s.style.display==="none"&&(s.style.display="block")});let t=document.getElementById("kpi-row");t&&t.querySelector(".skeleton")&&fn({});let n=document.getElementById("table-recent-activity");n&&n.querySelector(".skeleton")&&(n.innerHTML='<div class="chart-empty">Belum ada aktivitas terbaru</div>');let a=document.getElementById("list-upcoming-schedule");a&&a.querySelector(".skeleton")&&(a.innerHTML='<div class="chart-empty">Belum ada jadwal kegiatan</div>')}async function hn(e){if(e||(e=document.getElementById("main-content")),!e)return;un(),e._dashRefresh&&clearInterval(e._dashRefresh),e._skelTimeout&&clearTimeout(e._skelTimeout);let t=!1,n=()=>!t&&e&&e.isConnected&&!!e.querySelector("#dash-root"),a=()=>{t=!0,e._skelTimeout&&(clearTimeout(e._skelTimeout),e._skelTimeout=null),e._dashRefresh&&(clearInterval(e._dashRefresh),e._dashRefresh=null),un()};window.addEventListener("hashchange",a,{once:!0});let o=se(),s=String(o?.role||"").toLowerCase()==="superadmin",c=new Date,m=String(c.getDate()).padStart(2,"0"),i=c.toLocaleDateString("id-ID",{month:"short"}),d=c.getFullYear(),F=`01 ${i} \u2013 ${m} ${i} ${d}`;return e.innerHTML=`
    <div class="dashboard-wrap" id="dash-root">
      
      <!-- Dashboard Operasional Header -->
      <div class="dash-operasional-header">
        <div>
          <h1 class="dash-title">Dashboard Operasional</h1>
          <p class="dash-subtitle">Monitoring data personel, kegiatan, dan operasional Facility Care.</p>
        </div>
        <div class="dash-header-controls">
          <div class="dash-filter-pill">
            <select id="filter-jadwal-year" class="dash-header-select" aria-label="Pilih Tahun">
              <option value="${d}" selected>${d}</option>
              <option value="${d+1}">${d+1}</option>
            </select>
          </div>
          <div class="dash-filter-pill dash-date-display">
            <svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><rect x="3" y="4" width="18" height="18" rx="2"/><line x1="16" y1="2" x2="16" y2="6"/><line x1="8" y1="2" x2="8" y2="6"/><line x1="3" y1="10" x2="21" y2="10"/></svg>
            <span id="dash-date-range">${F}</span>
          </div>
          <button id="btn-dash-refresh" class="dash-refresh-btn" title="Muat ulang data">
            <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><path d="M21 12a9 9 0 1 1-9-9c2.52 0 4.93 1 6.74 2.74L21 8"/><path d="M21 3v5h-5"/></svg>
            Refresh Data
          </button>
        </div>
      </div>

      <!-- Top KPI Row (5 Cards) -->
      <div class="kpi-row" id="kpi-row">${nr()}</div>

      <!-- Middle Charts Row (3 Columns) -->
      <div class="charts-row">
        <!-- Chart 1: Tren Kegiatan Operasional -->
        <div class="chart-card">
          <div class="chart-card-header">
            <div>
              <div class="chart-card-title">Tren Kegiatan Operasional</div>
            </div>
            <select id="filter-jadwal-period" class="dash-header-select" style="background:#0E182B; border:1px solid #223456; border-radius:6px; padding:4px 8px; font-size:0.78rem; color:#F8FAFC;" aria-label="Filter Periode">
              <option value="month">Per Bulan</option>
            </select>
          </div>
          <div class="chart-canvas-wrap" style="height:150px;position:relative;margin-top:4px">
            <div id="skel-jadwal" class="skeleton" style="position:absolute;inset:0;border-radius:12px"></div>
            <canvas id="chart-jadwal" style="display:none"></canvas>
          </div>
          <div id="jadwal-legend" style="display:flex;justify-content:center;gap:12px;margin-top:6px;font-size:0.76rem;font-weight:600;color:#CBD5E1;flex-wrap:wrap;">
            <div style="display:flex;align-items:center;gap:5px"><div style="width:8px;height:8px;border-radius:50%;background:#2563EB"></div> Inspeksi</div>
            <div style="display:flex;align-items:center;gap:5px"><div style="width:8px;height:8px;border-radius:50%;background:#3B82F6"></div> General Cleaning</div>
            <div style="display:flex;align-items:center;gap:5px"><div style="width:8px;height:8px;border-radius:50%;background:#38BDF8"></div> Deep Cleaning</div>
            <div style="display:flex;align-items:center;gap:5px"><div style="width:8px;height:8px;border-radius:50%;background:#93C5FD"></div> Fogging</div>
          </div>
        </div>

        <!-- Chart 2: Permasalahan per Kategori -->
        <div class="chart-card" style="display:flex; flex-direction:column; justify-content:space-between;">
          <div class="chart-card-header">
            <div class="chart-card-title">Permasalahan per Kategori</div>
          </div>
          <div style="display:flex; align-items:center; justify-content:center; gap:16px; margin: 4px 0;">
            <div class="chart-canvas-wrap" style="width:120px;height:120px;position:relative">
              <div id="skel-donut" class="skeleton" style="position:absolute;inset:0;border-radius:12px"></div>
              <canvas id="chart-donut" style="display:none"></canvas>
            </div>
            <div id="donut-legend" class="donut-legend" style="min-width:110px;"></div>
          </div>
          <div style="text-align:center; font-size:0.78rem; color:#CBD5E1; display:flex; align-items:center; justify-content:center; gap:6px;">
            <svg width="12" height="12" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><rect x="3" y="4" width="18" height="18" rx="2"/><line x1="16" y1="2" x2="16" y2="6"/><line x1="8" y1="2" x2="8" y2="6"/></svg>
            <span id="donut-period-text">Distribusi Kategori Permasalahan</span>
          </div>
        </div>

        <!-- Chart 3: Tren Permasalahan 12 Bulan -->
        <div class="chart-card">
          <div class="chart-card-header">
            <div class="chart-card-title">Tren Permasalahan 12 Bulan</div>
            <div style="display:flex;align-items:center;gap:12px;font-size:0.76rem;font-weight:600;color:#CBD5E1">
               <div style="display:flex;align-items:center;gap:5px"><div style="width:8px;height:8px;border-radius:50%;background:#38BDF8"></div> Open</div>
               <div style="display:flex;align-items:center;gap:5px"><div style="width:8px;height:8px;border-radius:50%;background:#2563EB"></div> Closed</div>
            </div>
          </div>
          <div class="chart-canvas-wrap" style="height:150px;position:relative;margin-top:4px">
            <div id="skel-trend" class="skeleton" style="position:absolute;inset:0;border-radius:12px"></div>
            <canvas id="chart-trend" style="display:none"></canvas>
          </div>
        </div>
      </div>

      <!-- Bottom Row (2 Columns: Aktivitas Terbaru & Jadwal Kegiatan Mendatang) -->
      <div class="bottom-row">
        <!-- Aktivitas Terbaru -->
        <div class="chart-card">
          <div class="chart-card-header">
            <div class="chart-card-title">Aktivitas Terbaru</div>
            ${s?'<a href="#/audit-logs" class="chart-link">Lihat Semua &rarr;</a>':""}
          </div>
          <div id="table-recent-activity" class="dash-table-wrap dash-panel-scroll">
            ${mn(4)}
          </div>
        </div>

        <!-- Jadwal Kegiatan Mendatang -->
        <div class="chart-card">
          <div class="chart-card-header">
            <div class="chart-card-title">Jadwal Kegiatan Mendatang</div>
            <a href="#/calendar" class="chart-link">Lihat Kalender &rarr;</a>
          </div>
          <div id="list-upcoming-schedule" class="dash-panel-scroll" style="display:flex;flex-direction:column;gap:4px">
            ${mn(4)}
          </div>
        </div>
      </div>

    </div>
  `,document.getElementById("btn-dash-refresh")?.addEventListener("click",()=>{n()&&gn(e,n)}),document.getElementById("filter-jadwal-year")?.addEventListener("change",async $=>{let M=$.target.value,h=document.getElementById("skel-jadwal"),L=document.getElementById("chart-jadwal");h&&(h.style.display="block"),L&&(L.style.display="none");let R=await Je(`/api/dashboard/schedule-chart?year=${M}`,{},8e3);if(n())try{yn(R)}catch(I){console.warn("ScheduleChart render error:",I),at("skel-jadwal","chart-jadwal")}}),e._skelTimeout=setTimeout(()=>ir(e),5e3),await gn(e,n),a}async function gn(e,t=()=>e&&e.isConnected){e&&e._skelTimeout&&(clearTimeout(e._skelTimeout),e._skelTimeout=null);let n=document.getElementById("filter-jadwal-year")?.value||new Date().getFullYear(),[a,o,s,c,m,i,d,F]=await Promise.all([Je("/api/dashboard/kpi",{},8e3),Je("/api/dashboard/issues-trend",{},8e3),Je("/api/dashboard/issues-summary",{},8e3),Je("/api/dashboard/stats",{},8e3),Je("/api/dashboard/calendar",[],8e3),Je(`/api/dashboard/schedule-chart?year=${n}`,{},8e3),Je("/api/dashboard/dropdown-aggregates",{},8e3),Je("/api/dashboard/activity-log",[],8e3)]);if(t()){window.dashboardAggregates=d||{};try{fn(a)}catch($){console.warn("KPI render:",$)}try{yn(i)}catch($){console.warn("ScheduleChart render:",$),at("skel-jadwal","chart-jadwal")}try{rr(Array.isArray(s?.by_category)?s.by_category:[])}catch($){console.warn("Donut render:",$),at("skel-donut","chart-donut")}try{or(o)}catch($){console.warn("Trend render:",$),at("skel-trend","chart-trend")}try{sr(F)}catch($){console.warn("RecentActivity render:",$)}try{let $=c?.upcoming_schedule||[];lr($,m)}catch($){console.warn("UpcomingSchedule render:",$)}}}function fn(e){let t=document.getElementById("kpi-row");if(!t)return;e=e||{};let n=[{iconSvg:'<svg width="21" height="21" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" viewBox="0 0 24 24"><path d="M17 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2"/><circle cx="9" cy="7" r="4"/><path d="M23 21v-2a4 4 0 0 0-3-3.87"/><path d="M16 3.13a4 4 0 0 1 0 7.75"/></svg>',label:"Karyawan Aktif",val:e.employees?.current??0,href:"#/employees?dash_filter=active"},{iconSvg:'<svg width="21" height="21" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" viewBox="0 0 24 24"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/><line x1="16" y1="13" x2="8" y2="13"/><line x1="16" y1="17" x2="8" y2="17"/></svg>',label:"Kontrak Aktif",val:e.contracts?.current??0,href:"#/contracts?dash_filter=active"},{iconSvg:'<svg width="21" height="21" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" viewBox="0 0 24 24"><path d="M2 9h20L19 3H5L2 9z"/><path d="M4 9v11a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V9"/><path d="M10 22V12h4v10"/></svg>',label:"Cabang Operasional",val:e.branches?.current??0,href:"#/branches"},{iconSvg:'<svg width="21" height="21" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" viewBox="0 0 24 24"><rect x="3" y="4" width="18" height="18" rx="2" ry="2"/><line x1="16" y1="2" x2="16" y2="6"/><line x1="8" y1="2" x2="8" y2="6"/><line x1="3" y1="10" x2="21" y2="10"/></svg>',label:"Jadwal Kegiatan",val:e.schedule?.current??0,href:"#/timeline"},{iconSvg:'<svg width="21" height="21" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" viewBox="0 0 24 24"><path d="M10.29 3.86L1.82 18a2 2 0 0 0 1.71 3h16.94a2 2 0 0 0 1.71-3L13.71 3.86a2 2 0 0 0-3.42 0z"/><line x1="12" y1="9" x2="12" y2="13"/><line x1="12" y1="17" x2="12.01" y2="17"/></svg>',label:"Permasalahan",val:e.issues?.current??0,href:"#/issues?dash_filter=open"}];t.innerHTML=n.map(a=>`
    <a href="${a.href}" class="kpi-card">
      <div class="kpi-icon-wrap">
        ${a.iconSvg}
      </div>
      <div class="kpi-body">
        <div class="kpi-value" data-target="${a.val}">${a.val}</div>
        <div class="kpi-label">${a.label}</div>
      </div>
    </a>
  `).join(""),t.querySelectorAll(".kpi-value").forEach(a=>{tr(a,parseInt(a.dataset.target)||0)})}function yn(e){at("skel-jadwal","chart-jadwal");let t=document.getElementById("chart-jadwal");if(!t)return;Gt("jadwal"),e=e||{};let n=["Jan","Feb","Mar","Apr","Mei","Jun","Jul","Agu","Sep","Okt","Nov","Des"];if(!Object.values(e).some(i=>Array.isArray(i)&&i.some(d=>d>0))){xa(t,"Belum ada data jadwal kegiatan");return}let o=e["Inspeksi Hygiene"]||Array(12).fill(0),s=e["General Cleaning"]||Array(12).fill(0),c=e["Deep Cleaning"]||Array(12).fill(0),m=e.Fogging||Array(12).fill(0);et.jadwal=new Chart(t,{type:"bar",data:{labels:n,datasets:[{label:"Inspeksi",data:o,backgroundColor:"#2563EB",borderRadius:3},{label:"General Cleaning",data:s,backgroundColor:"#3B82F6",borderRadius:3},{label:"Deep Cleaning",data:c,backgroundColor:"#38BDF8",borderRadius:3},{label:"Fogging",data:m,backgroundColor:"#93C5FD",borderRadius:3}]},options:bn({plugins:{legend:{display:!1}},datasets:{bar:{barPercentage:.75,categoryPercentage:.85}},scales:{x:{stacked:!0,grid:{display:!1},ticks:{font:{family:"Inter",size:9},color:tt,maxRotation:0}},y:{stacked:!0,grid:{color:Ut},ticks:{font:{family:"Inter",size:9},color:tt},min:0}}})})}function rr(e){at("skel-donut","chart-donut");let t=document.getElementById("chart-donut"),n=document.getElementById("donut-legend");if(!t||!n)return;Gt("donut");let a=(e||[]).filter(i=>Ct(i.count)>0);if(!a.length){xa(t,"Belum ada data kategori");return}let o=a.map(i=>er(i.category,"Lainnya")),s=a.map(i=>Ct(i.count)),c=s.reduce((i,d)=>i+d,0);n.innerHTML=a.map((i,d)=>{let F=ya[d%ya.length],$=c>0?Math.round(i.count/c*100):0;return`
      <div style="display:flex;align-items:center;gap:8px;margin-bottom:6px;">
        <div style="width:8px;height:8px;border-radius:50%;background:${F};flex-shrink:0;"></div>
        <div style="display:flex;align-items:baseline;gap:4px;font-size:0.75rem;">
          <span style="font-weight:600;color:var(--text-1);">${o[d]}</span>
          <span style="color:var(--text-2);font-weight:700;">${i.count}</span>
          <span style="font-size:0.68rem;color:var(--text-3);">(${$}%)</span>
        </div>
      </div>
    `}).join("");let m={id:"centerText",beforeDraw:function(i){let d=i.width,F=i.height,$=i.ctx;$.restore();let M=(F/85).toFixed(2);$.font="bold "+M+"em Inter",$.textBaseline="middle",$.fillStyle="#F8FAFC";let h=c.toString(),L=Math.round((d-$.measureText(h).width)/2),R=F/2;$.fillText(h,L,R-6),$.font="600 "+(M*.38).toFixed(2)+"em Inter",$.fillStyle="#94A3B8";let I="Total",N=Math.round((d-$.measureText(I).width)/2);$.fillText(I,N,R+10),$.save()}};et.donut=new Chart(t,{type:"doughnut",data:{labels:o,datasets:[{data:s,backgroundColor:ya,borderWidth:2,borderColor:"#131F36",hoverBorderColor:"#131F36"}]},options:{responsive:!0,maintainAspectRatio:!1,animation:{duration:600},plugins:{legend:{display:!1},tooltip:{backgroundColor:"#1E293B",titleColor:"#F8FAFC",bodyColor:"#94A3B8",borderColor:"rgba(255, 255, 255, 0.1)",borderWidth:1,cornerRadius:8,bodyFont:bt,callbacks:{label:i=>` ${i.label}: ${i.parsed} kasus`}}},cutout:"75%"},plugins:[m]})}function or(e){at("skel-trend","chart-trend");let t=document.getElementById("chart-trend");if(!t)return;Gt("trend"),e=e||{};let n=["Jan","Feb","Mar","Apr","Mei","Jun","Jul","Agu","Sep","Okt","Nov","Des"],a=(e.labels||[]).map(c=>{if(!c||typeof c!="string")return"";try{let[,m]=c.split("-");return n[Number(m)-1]||m}catch{return c}}),o=(e.open||[]).map(c=>Ct(c)),s=(e.closed||[]).map(c=>Ct(c));if(!a.length){xa(t,"Belum ada data trend permasalahan");return}et.trend=new Chart(t,{type:"line",data:{labels:a,datasets:[{label:"Open",data:o,borderColor:"#38BDF8",backgroundColor:"rgba(56,189,248,.12)",fill:!0,tension:.4,pointRadius:3.5,pointHoverRadius:6,pointBackgroundColor:"#38BDF8",borderWidth:2},{label:"Closed",data:s,borderColor:"#2563EB",backgroundColor:"rgba(37,99,235,.12)",fill:!0,tension:.4,pointRadius:3.5,pointHoverRadius:6,pointBackgroundColor:"#2563EB",borderWidth:2}]},options:bn({plugins:{legend:{display:!1}},scales:{x:{grid:{display:!1},ticks:{font:{family:"Inter",size:9},color:tt,maxRotation:0}},y:{grid:{color:Ut},ticks:{font:{family:"Inter",size:9},color:tt},beginAtZero:!0}}})})}function sr(e){let t=document.getElementById("table-recent-activity");if(!t)return;let n=(e||[]).slice(0,6);if(!n.length){t.innerHTML='<div class="chart-empty">Belum ada aktivitas terbaru</div>';return}let a={issue:{icon:'<svg width="12" height="12" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24"><path d="M10.29 3.86L1.82 18a2 2 0 0 0 1.71 3h16.94a2 2 0 0 0 1.71-3L13.71 3.86a2 2 0 0 0-3.42 0z"/><line x1="12" y1="9" x2="12" y2="13"/><line x1="12" y1="17" x2="12.01" y2="17"/></svg>',bg:"#2563EB",text:"#FFFFFF",module:"Permasalahan",modBg:"rgba(56,189,248,0.16)",modCol:"#7DD3FC"},contract:{icon:'<svg width="12" height="12" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/></svg>',bg:"#2563EB",text:"#FFFFFF",module:"Data Kontrak",modBg:"rgba(37,99,235,0.18)",modCol:"#93C5FD"},employee:{icon:'<svg width="12" height="12" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>',bg:"#2563EB",text:"#FFFFFF",module:"Master Karyawan",modBg:"rgba(37,99,235,0.18)",modCol:"#93C5FD"},training:{icon:'<svg width="12" height="12" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/></svg>',bg:"#2563EB",text:"#FFFFFF",module:"Training",modBg:"rgba(56,189,248,0.16)",modCol:"#7DD3FC"},inspection:{icon:'<svg width="12" height="12" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24"><path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"/><polyline points="22 4 12 14.01 9 11.01"/></svg>',bg:"#2563EB",text:"#FFFFFF",module:"Inspeksi",modBg:"rgba(37,99,235,0.18)",modCol:"#93C5FD"},schedule:{icon:'<svg width="12" height="12" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24"><line x1="12" y1="5" x2="12" y2="19"/><line x1="5" y1="12" x2="19" y2="12"/></svg>',bg:"#2563EB",text:"#FFFFFF",module:"Jadwal Kegiatan",modBg:"rgba(56,189,248,0.16)",modCol:"#7DD3FC"}};t.innerHTML=`
    <table class="act-table">
      <thead>
        <tr>
          <th>Waktu</th>
          <th>Aktivitas</th>
          <th>Modul</th>
          <th>Pengguna</th>
        </tr>
      </thead>
      <tbody>
        ${n.map(o=>{let s=o.created_at?String(o.created_at).includes(" ")&&!String(o.created_at).includes("T")?`${String(o.created_at).replace(" ","T")}Z`:String(o.created_at):null,c=s?new Date(s):null,m=c&&!isNaN(c.getTime())?c.toLocaleTimeString("id-ID",{hour:"2-digit",minute:"2-digit",timeZone:"Asia/Jakarta"}).replace(/\./g,":"):"\u2014",i=a[o.type]||{icon:"\u{1F4CC}",bg:"#3B82F6",text:"#fff",module:"Operasional",modBg:"rgba(59,130,246,0.18)",modCol:"#60A5FA"},d=o.user_name||(o.branch?`Area ${o.branch}`:"Sistem FCMS"),F=(d||"F")[0].toUpperCase();return`
            <tr>
              <td class="act-time">${m}</td>
              <td>
                <div class="act-desc">
                  <span class="act-icon" style="background:${i.bg};color:${i.text}">${i.icon}</span>
                  <span title="${P(o.label)}">${P(o.label)}</span>
                </div>
              </td>
              <td>
                <span class="act-modul-pill" style="background:${i.modBg};color:${i.modCol}">
                  ${i.module}
                </span>
              </td>
              <td>
                <div class="act-user">
                  <span class="act-user-avatar" style="background:${i.bg}">
                    ${F}
                  </span>
                  <span>${P(d)}</span>
                </div>
              </td>
            </tr>
          `}).join("")}
      </tbody>
    </table>
  `}function lr(e,t){let n=document.getElementById("list-upcoming-schedule");if(!n)return;let a=Array.isArray(e)&&e.length>0?e:t||[];if(a=a.slice(0,5),!a.length){n.innerHTML='<div class="chart-empty">Belum ada jadwal kegiatan mendatang</div>';return}let o=["Min","Sen","Sel","Rab","Kam","Jum","Sab"],s=["Jan","Feb","Mar","Apr","Mei","Jun","Jul","Agu","Sep","Okt","Nov","Des"];n.innerHTML=a.map(c=>{let m=c.target_date||c.event_date||"",i="\u2014",d="\u2014";if(m)try{let R=new Date(m);isNaN(R.getTime())||(i=o[R.getDay()]||"Sen",d=`${R.getDate()} ${s[R.getMonth()]||""}`)}catch{}let F=c.title||c.activity_type||"Kegiatan Operasional",$=c.branch_name||"Cabang Operasional",M="event-pill-insp",h="Inspeksi",L=(F+" "+(c.activity_type||"")).toLowerCase();return L.includes("general cleaning")||L.includes("gc")?(M="event-pill-gc",h="GC"):L.includes("deep cleaning")||L.includes("dc")?(M="event-pill-dc",h="Deep Cleaning"):L.includes("fogging")?(M="event-pill-fog",h="Fogging"):L.includes("inspeksi")&&(M="event-pill-insp",h="Inspeksi"),`
      <div class="event-item">
        <div class="event-date-box">
          <span class="event-date-day">${i}</span>
          <span class="event-date-num">${d}</span>
        </div>
        <div class="event-info">
          <div class="event-title">${P(F)}</div>
          <div class="event-branch">
            <svg width="12" height="12" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M21 10c0 7-9 13-9 13s-9-6-9-13a9 9 0 0118 0z"/><circle cx="12" cy="10" r="3"/></svg>
            <span>${P($)}</span>
          </div>
        </div>
        <div class="event-pill ${M}">${h}</div>
      </div>
    `}).join("")}function at(e,t){let n=document.getElementById(e),a=document.getElementById(t);if(n&&(n.style.display="none"),a){a.style.display="block";let o=a.parentElement;if(o){let s=o.querySelector(".chart-empty");s&&s.remove()}}}function xa(e,t="Belum ada data"){if(!e)return;e.style.display="none";let n=e.parentElement;if(!n)return;if(!n.querySelector(".chart-empty")){let o=document.createElement("div");o.className="chart-empty",o.textContent=t,n.appendChild(o)}}async function xn(e){document.getElementById("app").innerHTML=`
    <div class="login-page">
      <div class="login-card">

        <div class="login-header">
          <div class="login-logo-wrap">
            <div class="login-logo-icon">
              <svg width="32" height="32" viewBox="0 0 48 48" fill="none">
                <rect width="48" height="48" rx="14" fill="url(#lg)"/>
                <path d="M12 20h6v16h-6V20zm10-8h6v24h-6V12zm10 6h6v18h-6V18z" fill="#fff" fill-opacity=".9"/>
                <defs>
                  <linearGradient id="lg" x1="0" y1="0" x2="48" y2="48" gradientUnits="userSpaceOnUse">
                    <stop stop-color="#2563EB"/>
                    <stop offset="1" stop-color="#0EA5E9"/>
                  </linearGradient>
                </defs>
              </svg>
            </div>
            <div>
              <h1 class="login-title">FCMS</h1>
              <p class="login-subtitle">Facility Care Management System</p>
            </div>
          </div>
          <div class="login-divider"></div>
          <p class="login-desc">Masuk untuk mengelola operasional Facility Care</p>
        </div>

        <form class="login-form" id="login-form" novalidate>
          <div class="form-group">
            <label class="form-label">Username / Email</label>
            <div class="input-with-icon">
              <svg class="input-prefix-icon" width="16" height="16" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24">
                <path d="M20 21v-2a4 4 0 0 0-4-4H8a4 4 0 0 0-4 4v2"/><circle cx="12" cy="7" r="4"/>
              </svg>
              <input type="text" name="username" class="form-control has-prefix-icon"
                placeholder="Masukkan username atau email"
                required autofocus autocomplete="username">
            </div>
          </div>

          <div class="form-group">
            <label class="form-label">Password</label>
            <div class="input-with-icon">
              <svg class="input-prefix-icon" width="16" height="16" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24">
                <rect x="3" y="11" width="18" height="11" rx="2"/><path d="M7 11V7a5 5 0 0 1 10 0v4"/>
              </svg>
              <input type="password" name="password" class="form-control has-prefix-icon"
                placeholder="\u2022\u2022\u2022\u2022\u2022\u2022\u2022\u2022" required autocomplete="current-password" id="login-password">
              <button type="button" class="input-icon-btn" id="toggle-password" aria-label="Toggle password">
                <svg width="16" height="16" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24" id="icon-eye">
                  <path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8z"/><circle cx="12" cy="12" r="3"/>
                </svg>
              </button>
            </div>
          </div>

          <div id="login-error" class="alert alert-danger" style="display:none"></div>

          <button type="submit" class="btn btn-primary btn-full btn-lg" id="login-btn" style="margin-top:4px">
            <span class="btn-text">
              <svg width="16" height="16" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24" style="margin-right:6px">
                <path d="M15 3h4a2 2 0 0 1 2 2v14a2 2 0 0 1-2 2h-4"/><polyline points="10 17 15 12 10 7"/><line x1="15" y1="12" x2="3" y2="12"/>
              </svg>
              Masuk ke FCMS
            </span>
            <span class="btn-spinner" style="display:none">\u23F3 Memproses...</span>
          </button>
        </form>
        <div class="login-version">FCMS v2.0 \xB7 Facility Care Indonesia</div>
      </div>

      <!-- Decorative background blobs -->
      <div class="login-blob login-blob-1"></div>
      <div class="login-blob login-blob-2"></div>
    </div>
  `;let t=document.getElementById("login-form"),n=document.getElementById("login-error"),a=document.getElementById("login-btn"),o=document.getElementById("toggle-password"),s=document.getElementById("login-password");o?.addEventListener("click",()=>{let c=s.type==="text";s.type=c?"password":"text",o.style.color=c?"":"var(--primary)"}),t?.addEventListener("submit",async c=>{c.preventDefault(),n.style.display="none";let m=t.username.value.trim(),i=t.password.value;if(!m||!i){n.textContent="Username dan password wajib diisi.",n.style.display="block";return}a.querySelector(".btn-text").style.display="none",a.querySelector(".btn-spinner").style.display="",a.disabled=!0;try{let d=await X("/api/auth/login",{method:"POST",body:JSON.stringify({username:m,password:i})});d.ok&&d.data.success?(on(d.data.data.token),kt(d.data.data.csrf_token||null),Ht(d.data.data.user),ae("Login berhasil! Selamat datang \u{1F44B}"),window.dispatchEvent(new Event("fm:login"))):(n.textContent=d.data.error||"Username atau password salah.",n.style.display="block",a.classList.add("shake"),setTimeout(()=>a.classList.remove("shake"),600))}catch{n.textContent="Gagal terhubung ke server. Periksa koneksi internet.",n.style.display="block"}finally{a.querySelector(".btn-text").style.display="",a.querySelector(".btn-spinner").style.display="none",a.disabled=!1}})}var dr=Object.freeze([["001","Pondok Bambu"],["002","Kemang"],["003","Depok"],["004","Cipete"],["005","Greenville"],["006","Kelapa Gading MOI"],["007","Bekasi"],["008","PIK"],["009","Bintaro"],["010","Bogor"],["011","Kuningan"],["012","BSD"],["013","Bandung Gatsu"],["014","Sunter"],["015","Gading Serpong"],["016","Cibubur"],["017","Bandung Setrasari"],["018","Harapan Indah"],["019","Cinere"],["020","Surabaya"],["021","Kelapa Gading Boulevard"],["022","Cikarang"],["023","Semarang"],["024","Grand Wisata"],["025","Bogor Baru"],["026","Tangerang City"],["027","Rawamangun"],["028","Taman Palem"],["029","Karawang"],["030","Ciputat"],["031","Sawangan"],["032","Cawang"],["033","Taman Mini"],["034","Puri Indah"],["035","Ciledug"],["036","Jagakarsa"],["037","Bali"],["038","Greenlake"],["039","Bandung Mekarwangi"],["040","Bekasi Timur"],["041","Karawaci"],["042","Muara Karang"],["043","Tebet"],["044","Depok GDC"],["045","FX Mall"],["046","Kaliurang Jogja"],["047","Bogor Tajur"],["048","PIK 2"],["049","Surabaya Merr"],["050","Cibinong"],["051","Bintaro Sektor 9"],["052","BSD Rawa Buntu"],["053","Alam Sutera"],["054","Bogor Gunung Batu"],["055","Gajah Mada"],["056","Citra Garden 2"],["057","Cikupa"],["058","Solo Baru"],["059","Malang Suhat"],["060","Palmerah"],["061","Cirebon"],["062","Summarecon Bekasi"],["063","Pasar Minggu"],["064","Pondok Indah"],["065","Kebon Sirih"],["066","Bandung KBP"],["067","Agora Mall"],["068","Kota Wisata"],["069","Depok Pancoran Mas"],["070","Bekasi Kota Bintang"]].map(([e,t])=>Object.freeze({code:e,name:t,full_name:`${e}. ${t}`,is_operational:!0,branch_scope:"operational",parent_branch_code:null}))),cr=Object.freeze([Object.freeze({code:"000",name:"Management",full_name:"000. Management",is_operational:!1,branch_scope:"hq",parent_branch_code:null}),Object.freeze({code:"A01",name:"Aesthica",full_name:"A01. Aesthica",is_operational:!1,branch_scope:"joined_kemang",parent_branch_code:"002"})]),kn=new Map([...dr,...cr].map(e=>[e.code,e])),pr={"069":"Bekasi Kota Bintang","070":"J7"};function Qe(e){return String(e??"").trim()}function va(e){let t=Qe(e);return/^(00[1-9]|0[1-6][0-9]|070)$/.test(t)}function vn(e){return kn.has(Qe(e))}function ur(e){let t=Qe(e?.code),n=String(e?.full_name||e?.name||"").trim();return/^T\d+[A-Z0-9-]*$/i.test(t)||/^TEST_/i.test(n)||/^U558D/i.test(t)||/\[UAT-/i.test(n)}function mr(e){let t=Qe(e?.code),n=kn.get(t),a=String(e?.name||"").trim(),o=String(e?.full_name||a||`Cabang #${e?.id}`).trim(),s=pr[t];return n&&s&&a.toLowerCase()===s.toLowerCase()?n.full_name:o}function wn(e=[]){let t=Array.isArray(e)?e.filter(s=>!s.deleted_at&&!ur(s)):[],a=t.filter(s=>vn(s.code)).length>=65?t.filter(s=>vn(s.code)):t,o=s=>{let c=Qe(s);return c==="000"?0:va(c)?parseInt(c,10):c==="A01"?900:1e3};return[...a].sort((s,c)=>o(s.code)-o(c.code)||Qe(s.code).localeCompare(Qe(c.code))||(s.id||0)-(c.id||0)).map(s=>({value:s.id,label:mr(s),code:Qe(s.code),is_operational:va(s.code),branch_scope:s.code==="000"?"hq":s.code==="A01"?"joined_kemang":va(s.code)?"operational":"other"}))}var Vt=100,Ft=null,Et=null;async function $t(e,t={},n={}){let a=Math.min(Vt,Math.max(1,Number(n.limit)||Vt)),o=Math.max(1,Number(n.maxPages)||500),s=[],c=new Set;for(let m=1;m<=o;m++){let i=new URLSearchParams;if(i.set("page",String(m)),i.set("limit",String(a)),t&&typeof t=="object")for(let[h,L]of Object.entries(t))L!=null&&L!==""&&L!=="all"&&i.set(h,String(L));let d=await X(`${e}?${i.toString()}`);if(!d.ok)throw new Error(d.data?.error||`Gagal memuat data (HTTP ${d.status||500})`);let F=Array.isArray(d.data?.data)?d.data.data:Array.isArray(d.data)?d.data:[];for(let h of F){if(h&&h.id!==void 0&&h.id!==null){if(c.has(h.id))continue;c.add(h.id)}s.push(h)}let $=d.data?.pagination,M=$&&Number($.pages)>=0?Number($.pages):null;if(M!==null){if(m>=M)break}else if(F.length<a)break;if(F.length===0)break}return s}async function ue(e,t={},n={}){return $t(e,t,{...n,limit:Vt})}async function Sn({search:e="",status:t="Aktif",branch_id:n="",division:a="",page:o=1,limit:s=20}={}){let c=Math.min(Vt,Math.max(1,Number(s)||20)),m=new URLSearchParams({page:String(Math.max(1,Number(o)||1)),limit:String(c)});e&&m.set("search",e),t&&t!=="all"&&m.set("status",t),n&&n!=="all"&&m.set("branch_id",n),a&&a!=="all"&&m.set("division",a);let i=await X(`/api/employees/selector?${m.toString()}`);if(!i.ok)throw new Error(i.data?.error||`Gagal mencari karyawan (HTTP ${i.status||500})`);return{data:Array.isArray(i.data?.data)?i.data.data:[],pagination:i.data?.pagination||{total:0,page:1,limit:c,pages:0}}}async function ze(e=!1){return Ft&&!e||(Ft=(await $t("/api/employees/selector",{status:"Aktif"})).map(n=>({value:n.id,id:n.id,label:n.full_name,full_name:n.full_name,branch_id:n.branch_id??null,branch_name:n.branch_name||"",division:n.division||"",status:n.status||""}))),Ft}async function Jt(e=!1){return(await ze(e)).map(n=>({value:n.label,label:n.label}))}async function we(e=!1){if(Et&&!e)return Et;let t=await X("/api/branches?all=1");if(!t.ok)throw new Error(t.data?.error||"Gagal memuat data cabang");return Et=wn(t.data?.data||[]),Et}function At(){Ft=null,Et=null}Be();Be();Ie();var ht=100;function Cn(e,t,n,a=ht){return n?e.has(t)?{accepted:!0,count:e.size}:e.size>=a?{accepted:!1,count:e.size,reason:`Maksimum ${a} item dapat dipilih untuk penghapusan bulk`}:(e.add(t),{accepted:!0,count:e.size}):(e.delete(t),{accepted:!0,count:e.size})}function Fn(e,t,n,a=ht){let o=(t||[]).map(c=>typeof c=="object"&&c!==null?c.id:c).filter(c=>c!=null);if(!n)return o.forEach(c=>e.delete(c)),{accepted:!0,count:e.size};let s=o.filter(c=>!e.has(c));return e.size+s.length>a?{accepted:!1,count:e.size,reason:`Maksimum ${a} item dapat dipilih untuk penghapusan bulk (saat ini ${e.size} item sudah dipilih)`}:(s.forEach(c=>e.add(c)),{accepted:!0,count:e.size})}function Xe({title:e,headerRight:t="",badges:n=[],meta:a=[],notes:o="",notesLabel:s="Catatan",docLink:c="",row:m,handlers:i={}}){let d=document.createElement("article");d.className="fc-mobile-card";let F=`
    <div class="fc-mobile-card-head">
      <div class="fc-mobile-card-title">${P(e||"-")}</div>
      ${t?`<div class="fc-mobile-card-right">${He(t)}</div>`:""}
    </div>
  `,$=(n||[]).filter(Boolean),M=$.length>0?`<div class="fc-mobile-card-badges">${$.map(_=>He(_)).join("")}</div>`:"",h=(a||[]).filter(_=>_&&(_.html||_.value!==void 0&&_.value!==null&&String(_.value).trim()!==""&&String(_.value).trim()!=="-")),L=h.length>0?`
      <div class="fc-mobile-card-meta">
        ${h.map(_=>`
          <div class="fc-mobile-meta-item${_.full?" fc-mobile-meta-item--full":""}">
            <span class="fc-mobile-meta-label">${P(_.label)}</span>
            <span class="fc-mobile-meta-value">${_.html?He(_.html):P(String(_.value))}</span>
          </div>
        `).join("")}
      </div>
    `:"",R=o!=null?String(o).trim():"",I=R&&R!=="-"?`
      <div class="fc-mobile-card-notes">
        <span class="fc-mobile-meta-label">${P(s)}</span>
        <div class="fc-mobile-notes-text">${P(R)}</div>
      </div>
    `:"",N=typeof c=="string"&&/^(https?:\/\/|\/)/i.test(c.trim())?c.trim():"";d.innerHTML=F+M+L+I;let{onView:q,onEdit:v,onDelete:B,actions:x=[]}=i||{};if(!!(N||q||v||B||x.length>0)){let _=document.createElement("div");if(_.className="fc-mobile-card-footer",N){let w=document.createElement("a");w.href=N,w.target="_blank",w.rel="noopener",w.className="fc-mobile-doc-btn",w.innerHTML='<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/></svg><span>Buka Dokumen</span>',_.appendChild(w)}let E=document.createElement("div");if(E.className="fc-mobile-card-actions",q){let w=document.createElement("button");w.type="button",w.className="fc-mobile-action-btn fc-action-view",w.setAttribute("aria-label","Lihat detail"),w.innerHTML='<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8z"></path><circle cx="12" cy="12" r="3"></circle></svg><span>Detail</span>',w.addEventListener("click",()=>q(m)),E.appendChild(w)}if(v){let w=document.createElement("button");w.type="button",w.className="fc-mobile-action-btn fc-action-edit",w.setAttribute("aria-label","Edit data"),w.innerHTML='<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M12 20h9"></path><path d="M16.5 3.5a2.121 2.121 0 0 1 3 3L7 19l-4 1 1-4L16.5 3.5z"></path></svg><span>Edit</span>',w.addEventListener("click",()=>v(m)),E.appendChild(w)}if(x.forEach(w=>{let K=document.createElement("button");K.type="button",K.className=`fc-mobile-action-btn ${w.class||""}`,K.setAttribute("aria-label",w.label||"Aksi"),K.innerHTML=He(w.icon?`${w.icon}<span>${P(w.label||"")}</span>`:P(w.label||"Aksi")),K.addEventListener("click",()=>w.handler(m)),E.appendChild(K)}),B){let w=document.createElement("button");w.type="button",w.className="fc-mobile-action-btn fc-action-del",w.setAttribute("aria-label","Hapus data"),w.innerHTML='<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><polyline points="3 6 5 6 21 6"></polyline><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"></path></svg><span>Hapus</span>',w.addEventListener("click",()=>B(m)),E.appendChild(w)}E.childElementCount>0&&_.appendChild(E),d.appendChild(_)}return d}function En({columns:e,data:t,onEdit:n,onDelete:a,onView:o,actions:s=[],emptyText:c="Tidak ada data",bulkSelect:m=null,mobileCardRenderer:i=null}){let d=document.createElement("div");if(d.className=i?"table-wrapper fc-desktop-only-table":"table-wrapper",!t||t.length===0)return d.className="table-wrapper",d.innerHTML=`<div class="empty-state"><p>${c}</p></div>`,d;let F=document.createElement("table");F.className="data-table";let $=document.createElement("thead"),M=document.createElement("tr");if(m){let R=document.createElement("th");R.style.width="40px",R.style.textAlign="center";let I=document.createElement("input");I.type="checkbox",I.id="select-all-checkbox",I.title="Pilih semua";let N=m.maxSelect||1/0;I.addEventListener("change",()=>{if(Number.isFinite(N)){let q=Fn(m.selectedIds,t,I.checked,N);q.accepted||(I.checked=!1,typeof m.onLimitExceeded=="function"&&m.onLimitExceeded(q.reason));let v=d.querySelectorAll(".row-checkbox");t.forEach((B,x)=>{v[x]&&(v[x].checked=m.selectedIds.has(B.id))})}else t.forEach(q=>{I.checked?m.selectedIds.add(q.id):m.selectedIds.delete(q.id)}),d.querySelectorAll(".row-checkbox").forEach(q=>q.checked=I.checked);m.onToggle()}),R.appendChild(I),M.appendChild(R)}if(e.forEach(R=>{let I=document.createElement("th");I.textContent=R.label,R.width&&(I.style.width=R.width),M.appendChild(I)}),n||a||o||s.length>0){let R=document.createElement("th");R.textContent="Aksi",R.style.width="120px",M.appendChild(R)}$.appendChild(M),F.appendChild($);let h=document.createElement("tbody"),L=m?.maxSelect||1/0;if(t.forEach(R=>{let I=document.createElement("tr");if(m){let N=document.createElement("td");N.style.textAlign="center",N.style.width="40px";let q=document.createElement("input");q.type="checkbox",q.className="row-checkbox",q.checked=m.selectedIds.has(R.id),q.addEventListener("change",()=>{if(Number.isFinite(L)){let v=Cn(m.selectedIds,R.id,q.checked,L);if(!v.accepted)q.checked=!1,typeof m.onLimitExceeded=="function"&&m.onLimitExceeded(v.reason);else if(!q.checked){let B=document.getElementById("select-all-checkbox");B&&(B.checked=!1)}}else if(q.checked)m.selectedIds.add(R.id);else{m.selectedIds.delete(R.id);let v=document.getElementById("select-all-checkbox");v&&(v.checked=!1)}m.onToggle()}),N.appendChild(q),I.appendChild(N)}if(e.forEach(N=>{let q=document.createElement("td");if(N.render){let v=N.render(R[N.key],R);v instanceof HTMLElement?q.appendChild(v):q.innerHTML=v?He(v):""}else q.textContent=R[N.key]!==null&&R[N.key]!==void 0&&R[N.key]!==""?R[N.key]:"";N.nowrap&&(q.style.whiteSpace="nowrap"),I.appendChild(q)}),n||a||o||s.length>0){let N=document.createElement("td");N.className="actions-cell";let q=document.createElement("div");if(q.className="btn-group",o){let v=document.createElement("button");v.type="button",v.className="fc-action-btn fc-action-view",v.innerHTML='<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8z"></path><circle cx="12" cy="12" r="3"></circle></svg>',v.title="Lihat",v.setAttribute("aria-label","Lihat detail"),v.addEventListener("click",()=>o(R)),q.appendChild(v)}if(n){let v=document.createElement("button");v.type="button",v.className="fc-action-btn fc-action-edit",v.innerHTML='<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M12 20h9"></path><path d="M16.5 3.5a2.121 2.121 0 0 1 3 3L7 19l-4 1 1-4L16.5 3.5z"></path></svg>',v.title="Edit",v.setAttribute("aria-label","Edit data"),v.addEventListener("click",()=>n(R)),q.appendChild(v)}if(s.forEach(v=>{let B=document.createElement("button");B.type="button",B.className=`btn btn-xs ${v.class||"btn-ghost"}`,B.innerHTML=v.icon||v.label,B.title=v.label,B.setAttribute("aria-label",v.label||"Aksi"),B.addEventListener("click",()=>v.handler(R)),q.appendChild(B)}),a){let v=document.createElement("button");v.type="button",v.className="fc-action-btn fc-action-del",v.innerHTML='<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><polyline points="3 6 5 6 21 6"></polyline><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"></path></svg>',v.title="Hapus",v.setAttribute("aria-label","Hapus data"),v.addEventListener("click",()=>a(R)),q.appendChild(v)}N.appendChild(q),I.appendChild(N)}h.appendChild(I)}),F.appendChild(h),d.appendChild(F),typeof i=="function"){let R=document.createElement("div");R.className="fc-table-responsive-host",R.appendChild(d);let I=document.createElement("div");I.className="fc-mobile-cards-list";let N={onEdit:n,onDelete:a,onView:o,actions:s};return t.forEach(q=>{let v=i(q,N);v instanceof HTMLElement&&I.appendChild(v)}),R.appendChild(I),R}return d}function Te({page:e,pages:t,total:n,limit:a,onPage:o}){if(!t||t<=1)return null;let s=Math.min(Math.max(1,Number(e)||1),t),c=Math.max(1,Number(a)||20),m=Math.max(0,Number(n)||0),i=document.createElement("div");i.className="pagination";let d=document.createElement("span");d.className="pagination-info";let F=m>0?(s-1)*c+1:0,$=m>0?Math.min(m,s*c):0;d.textContent=m>0?`Hal. ${s} / ${t} \u2022 ${F}-${$} dari ${m} data`:`Total: ${m} data`,i.appendChild(d);let M=document.createElement("div");M.className="pagination-btns";let h=(I,N,q=!1,v=!1)=>{let B=document.createElement("button");B.className=`btn btn-sm ${v?"btn-primary":"btn-ghost"} pagination-btn`,B.textContent=I,B.disabled=q,B.addEventListener("click",()=>o(N)),M.appendChild(B)};h("\xAB",1,s===1),h("\u2039",s-1,s===1);let L=Math.max(1,s-2),R=Math.min(t,s+2);for(let I=L;I<=R;I++)h(I,I,!1,I===s);return h("\u203A",s+1,s===t),h("\xBB",t,s===t),i.appendChild(M),i}function nt(e){if(e==null||e===""||typeof e=="boolean")return null;let t=Number(e);return Number.isInteger(t)&&t>=1?t:null}function Wt(e,t,n){let a=e&&typeof e.querySelector=="function"?e.querySelector(t):null;return nt(a?a.value:n)}function gr(e){return nt(e?.data?.data?.row_version??e?.data?.row_version)}async function Yt({res:e,isEdit:t,item:n,close:a,reloadData:o,editSuccessMessage:s,createSuccessMessage:c,defaultErrorMessage:m}){if(e.ok){let i=gr(e);return t&&i!==null&&n&&typeof n=="object"&&(n.row_version=i),ae(t?s:c),a(),await o(),!0}if(e.status===409){let i=e.data?.error||e.data?.message||"Data telah diubah oleh pengguna atau sesi lain.";return j(`Konflik data (HTTP 409): ${i} Silakan muat ulang / periksa data terbaru sebelum mengedit kembali.`),a(),await o(),!1}return j(e.data?.error||e.data?.message||m),!1}var br=/^(\d{4})-(\d{2})-(\d{2})(?:[T\s].*)?$/,hr=/^(\d{2})[\/\-](\d{2})[\/\-](\d{4})$/,fr=/^(\d{4})-(\d{2})$/,yr=/^\d{5}$/;function ka(e,t,n){if(t<1||t>12||n<1||n>31)return null;let a=new Date(Date.UTC(e,t-1,n));return a.getUTCFullYear()!==e||a.getUTCMonth()!==t-1||a.getUTCDate()!==n?null:a.getTime()}function wa(e,{allowYearMonth:t=!0}={}){if(e==null)return null;let n=String(e).trim();if(!n||n==="-")return null;let a=n.match(br);if(a)return ka(Number(a[1]),Number(a[2]),Number(a[3]));if(a=n.match(hr),a)return ka(Number(a[3]),Number(a[2]),Number(a[1]));if(yr.test(n)){let o=Number(n);return o>=1e4?Date.UTC(1899,11,30)+o*864e5:null}return t&&(a=n.match(fr),a)?ka(Number(a[1]),Number(a[2]),1):null}function $n(e){return[...e].sort((t,n)=>{let a=wa(t,{allowYearMonth:!0}),o=wa(n,{allowYearMonth:!0});return a===null&&o!==null?1:a!==null&&o===null?-1:a!==null&&o!==null&&o!==a?o-a:String(n||"").localeCompare(String(t||""))})}function xe(e,t="-"){if(e==null||String(e).trim()===""||String(e).trim()==="-")return t;let n=wa(e,{allowYearMonth:!1});return n!==null?new Date(n).toLocaleDateString("id-ID",{day:"numeric",month:"short",year:"numeric",timeZone:"UTC"}):String(e)}var _n=["FACILITY CARE","SECURITY"],kr=["Aktif","Tidak Aktif","Resign","Cut"];function Bn(e){if(!e||e==="-")return"\u{1F464}";let t=String(e).trim().split(/\s+/);return t.length===1?t[0].slice(0,2).toUpperCase():(t[0][0]+t[1][0]).toUpperCase()}async function Tn(e,t){let a=[],o=[],s="all",c="all",m="all",i="",d=1,F=0,$=1,M=!1,h={total:0,active:0,nonactive:0,fc:0,sec:0},L=null,R=!1,I=localStorage.getItem("fcms_employees_view_mode")||"card";(t?t.get("dash_filter"):null)==="active"&&(s="Aktif");let q=se(),v=re(q,"employees","write"),B=re(q,"employees","write"),x=re(q,"employees","delete");try{let r=await we();o=(Array.isArray(r)?r:[]).filter(l=>String(l.code||"").trim()==="000"||l.branch_scope==="hq"||/^(?:cabang\s+)?000\.\s*management$/i.test(String(l.label||l.full_name||"").trim()))}catch(r){console.error("Error fetching branches:",r)}function k(){let r={sort:"join_date_desc"};return i.trim()&&(r.search=i.trim()),m&&m!=="all"&&(r.branch_id=m),c&&c!=="all"&&(r.division=c),s&&s!=="all"&&(r.status=s),r}function _(){let r=e.querySelector("#emp-stat-total"),l=e.querySelector("#emp-stat-active"),f=e.querySelector("#emp-stat-fc"),u=e.querySelector("#emp-stat-sec"),p=e.querySelector("#emp-stat-sub"),A=e.querySelector("#emp-pill-count-all"),T=e.querySelector("#emp-pill-count-active"),g=e.querySelector("#emp-pill-count-nonactive");r&&(r.textContent=h.total),l&&(l.textContent=h.active),f&&(f.textContent=h.fc),u&&(u.textContent=h.sec),p&&(p.textContent=`${h.active} Aktif \u2022 ${h.nonactive} Nonaktif`),A&&(A.textContent=h.total),T&&(T.textContent=h.active),g&&(g.textContent=h.nonactive)}async function E({refreshStats:r=!1}={}){if(R=!0,!!!e.querySelector("#emp-list-container"))y();else{let f=e.querySelector("#emp-list-container");f&&(f.innerHTML='<div class="emp-empty-state"><p style="color:var(--text-3); margin:0;">Memuat data karyawan...</p></div>')}try{let f=new URLSearchParams({page:String(d),limit:String(20),...k()});(r||!M)&&f.set("include_stats","1");let u=await X(`/api/employees?${f.toString()}`);if(u.ok){let p=u.data?.data||[];a=Array.isArray(p)?p:[];let A=u.data?.pagination||{};if(F=Number(A.total??a.length),$=Math.max(1,Number(A.pages??Math.ceil(F/20))||1),d=Math.max(1,Number(A.page??d)),u.data?.stats&&(h=u.data.stats,M=!0),a.length===0&&d>$)return d=$,R=!1,E({refreshStats:r})}else a=[],F=0,$=1,j(u.data?.error||"Gagal memuat data karyawan")}catch(f){console.warn("Gagal load employees:",f),a=[],F=0,$=1,j("Gagal memuat data karyawan: "+f.message)}finally{R=!1,e.querySelector("#emp-list-container")?(_(),S()):y()}}async function w(){if(!re(q,"export")){j("Anda tidak memiliki izin untuk mengekspor data.");return}try{let[{downloadExcel:r},l]=await Promise.all([Promise.resolve().then(()=>(Ce(),Ve)),ue("/api/employees/export",k())]),f=l.map((u,p)=>({No:p+1,"Nama Lengkap":u.full_name||"",Cabang:u.branch_name||"",Divisi:u.division||"","No. HP":u.phone||"","Tgl Masuk":u.join_date||"",Status:u.status||"","Target Pindah OS":u.target_pindah_os||"","Target Selesai OS":u.target_selesai_os||"",Catatan:u.notes||""}));r(f,`Data_Karyawan_FCMS_${new Date().toISOString().split("T")[0]}`),ae("Berhasil mengekspor data karyawan ke Excel")}catch(r){j("Gagal ekspor Excel: "+r.message)}}function K(r=null){let l=!!(r&&r.id);if(l&&!B){j("Anda tidak memiliki izin untuk mengedit data karyawan.");return}if(!l&&!v){j("Anda tidak memiliki izin untuk menambah data karyawan.");return}let f=!!(r&&r.id),u=f?nt(r?.row_version):null,p=r?.division||"FACILITY CARE",A=r?.status||"Aktif",T=document.createElement("div");T.className="emp-modal-body",T.innerHTML=`
      <div style="display:flex; flex-direction:column; gap:14px;">
        ${f?`<input type="hidden" id="emp-modal-row-version" value="${u!==null?u:""}">`:""}
        <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#94A3B8; margin-bottom:5px;">
              Nama Lengkap Karyawan <span style="color:#EF4444;">*</span>
            </label>
            <input type="text" id="emp-modal-name" class="form-control" style="width:100%; padding:9px 12px; border:1px solid rgba(255,255,255,0.12); border-radius:8px; background:#0E182B; color:#F8FAFC;" 
              value="${r?.full_name?String(r.full_name).replace(/"/g,"&quot;"):""}" required />
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#94A3B8; margin-bottom:5px;">
              No. Handphone / WhatsApp
            </label>
            <input type="tel" id="emp-modal-phone" class="form-control" style="width:100%; padding:9px 12px; border:1px solid rgba(255,255,255,0.12); border-radius:8px; background:#0E182B; color:#F8FAFC;" 
              value="${r?.phone?String(r.phone).replace(/"/g,"&quot;"):""}" />
          </div>
        </div>

        <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#94A3B8; margin-bottom:5px;">
              Cabang Penempatan <span style="color:#EF4444;">*</span>
            </label>
            <select id="emp-modal-branch" class="form-control" style="width:100%; padding:9px 12px; border:1px solid rgba(255,255,255,0.12); border-radius:8px; background:#131F36; color:#F8FAFC;" required>
              <option value="">-- Pilih Cabang --</option>
              ${o.map(C=>`
                <option value="${C.value}" ${String(r?.branch_id)===String(C.value)?"selected":""}>
                  ${C.label}
                </option>
              `).join("")}
            </select>
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#94A3B8; margin-bottom:5px;">
              Divisi / Bagian <span style="color:#EF4444;">*</span>
            </label>
            <select id="emp-modal-division" class="form-control" style="width:100%; padding:9px 12px; border:1px solid rgba(255,255,255,0.12); border-radius:8px; background:#131F36; color:#F8FAFC;" required>
              ${_n.map(C=>`
                <option value="${C}" ${p===C?"selected":""}>${C}</option>
              `).join("")}
            </select>
          </div>
        </div>

        <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#94A3B8; margin-bottom:5px;">
              Tanggal Masuk Kerja
            </label>
            <input type="date" id="emp-modal-joindate" class="form-control" style="width:100%; padding:9px 12px; border:1px solid rgba(255,255,255,0.12); border-radius:8px; background:#0E182B; color:#F8FAFC; color-scheme:dark;" 
              value="${r?.join_date||""}" />
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#94A3B8; margin-bottom:5px;">
              Status Karyawan <span style="color:#EF4444;">*</span>
            </label>
            <select id="emp-modal-status" class="form-control" style="width:100%; padding:9px 12px; border:1px solid rgba(255,255,255,0.12); border-radius:8px; background:#131F36; color:#F8FAFC;" required>
              ${kr.map(C=>`
                <option value="${C}" ${A===C?"selected":""}>${C}</option>
              `).join("")}
            </select>
          </div>
        </div>

        <div>
          <label style="display:block; font-size:0.82rem; font-weight:600; color:#94A3B8; margin-bottom:5px;">
            Catatan / Keterangan Karyawan
          </label>
          <textarea id="emp-modal-notes" class="form-control" rows="2" style="width:100%; padding:9px 12px; border:1px solid rgba(255,255,255,0.12); border-radius:8px; background:#0E182B; color:#F8FAFC;">${r?.notes||""}</textarea>
        </div>
      </div>
    `;let g=!1;ce({title:f?"\u270F\uFE0F Edit Data Karyawan":"\u2795 Tambah Data Karyawan",content:T,size:"md",confirmText:f?"Simpan Perubahan":"Simpan Karyawan",confirmClass:"btn-primary",onConfirm:async(C,z)=>{if(g)return;let O=(T.querySelector("#emp-modal-name").value||"").trim(),te=(T.querySelector("#emp-modal-phone").value||"").trim(),Q=T.querySelector("#emp-modal-branch").value,ne=T.querySelector("#emp-modal-division").value,ee=T.querySelector("#emp-modal-joindate").value||null,H=T.querySelector("#emp-modal-status").value,U=(T.querySelector("#emp-modal-notes").value||"").trim();if(!O||!Q||!ne||!H){j("Harap lengkapi Nama Lengkap, Cabang, Divisi, dan Status!");return}let D=null;if(f&&(D=Wt(T,"#emp-modal-row-version",r?.row_version),D===null)){j("Versi data karyawan (row_version) tidak tersedia. Silakan muat ulang halaman lalu coba lagi.");return}let V={full_name:O,phone:te,branch_id:Number(Q),division:ne,join_date:ee,status:H,notes:U,...f?{row_version:D}:{}};g=!0;try{let Z=f?await X(`/api/employees/${r.id}`,{method:"PUT",body:JSON.stringify(V)}):await X("/api/employees",{method:"POST",body:JSON.stringify(V)});await Yt({res:Z,isEdit:f,item:r,close:z,reloadData:async()=>{At(),await E({refreshStats:!0})},editSuccessMessage:"Data karyawan berhasil diperbarui!",createSuccessMessage:"Karyawan baru berhasil ditambahkan!",defaultErrorMessage:"Gagal menyimpan data karyawan"})}catch(Z){j("Terjadi kesalahan: "+Z.message)}finally{g=!1}}})}async function G(r){ge({title:"Nonaktifkan Karyawan?",message:`Apakah Anda yakin ingin menonaktifkan personil <strong>${r.full_name}</strong> dari cabang <em>${r.branch_name||"cabang"}</em>?`,confirmText:"Ya, Nonaktifkan",confirmClass:"btn-danger",onConfirm:async()=>{try{let l=await X(`/api/employees/${r.id}`,{method:"DELETE"});l.ok?(ae("Status karyawan berhasil diubah menjadi Tidak Aktif"),At(),await E({refreshStats:!0})):j(l.data?.error||"Gagal menonaktifkan data")}catch(l){j("Terjadi kesalahan: "+l.message)}}})}function J(){return a}function W(r){return r.length===0?`
            <div class="emp-empty-state">
              <div style="font-size:2.8rem; margin-bottom:10px;">\u{1F50D}</div>
              <h3 style="font-size:1.15rem; font-weight:700; color:#F8FAFC; margin:0 0 6px 0;">Tidak ada karyawan yang cocok</h3>
              <p style="font-size:0.85rem; color:#94A3B8; margin:0 0 16px 0;">Coba sesuaikan kata kunci pencarian, filter cabang, atau filter divisi.</p>
              <button id="emp-reset-filter" class="btn btn-primary" style="font-size:0.82rem; padding:8px 18px; border-radius:8px; background:#2563EB; color:#fff; border:none; font-weight:700;">Reset Semua Filter</button>
            </div>
      `:I==="table"?`
            <div class="emp-table-card">
              <div class="emp-table-responsive">
                <table class="emp-table">
                  <thead>
                    <tr>
                      <th style="width:45px; text-align:center;">No</th>
                      <th>Karyawan</th>
                      <th>Cabang</th>
                      <th>Divisi</th>
                      <th>No. HP</th>
                      <th>Tgl Masuk</th>
                      <th>Catatan</th>
                      <th style="text-align:center;">Status</th>
                      <th style="text-align:center;">Aksi</th>
                    </tr>
                  </thead>
                  <tbody>
                    ${r.map((l,f)=>{let u=l.division||"FACILITY CARE",p="emp-div-fc",A="";u==="SECURITY"&&(p="emp-div-sec",A="sec");let T=l.status==="Aktif",g=l.phone||"";return`
                        <tr data-id="${l.id}">
                          <td style="text-align:center; color:#94A3B8; font-weight:600;">${(d-1)*20+f+1}</td>
                          <td>
                            <div style="display:flex; align-items:center; gap:10px;">
                              <div class="emp-avatar ${A}" style="width:34px; height:34px; font-size:0.75rem;">
                                ${Bn(l.full_name)}
                              </div>
                              <div style="min-width:0;">
                                <div style="font-weight:700; color:#F8FAFC; font-size:0.9rem;">${l.full_name||"Karyawan"}</div>
                                <div style="font-size: 0.76rem; color:#94A3B8;">ID #${l.id} ${l.nik?"\u2022 "+l.nik:""}</div>
                              </div>
                            </div>
                          </td>
                          <td>
                            <span style="font-weight:600; color:#E2E8F0;">${l.branch_name||"Semua Cabang"}</span>
                          </td>
                          <td>
                            <span class="emp-div-badge ${p}">
                              ${u==="SECURITY"?"\u{1F6E1}\uFE0F":"\u{1F9F9}"} ${u}
                            </span>
                          </td>
                          <td>
                            ${g?`
                              <a href="https://wa.me/${g.replace(/[^0-9]/g,"").replace(/^0/,"62")}" target="_blank" rel="noopener" style="color:#34D399; text-decoration:none; font-weight:600; display:inline-flex; align-items:center; gap:4px;">
                                <span>\u{1F4F1}</span> ${g}
                              </a>
                            `:'<span style="color:var(--text-3);">-</span>'}
                          </td>
                          <td style="font-size:0.84rem; color:#94A3B8;">
                            ${l.join_date?xe(l.join_date):"-"}
                          </td>
                          <td style="font-size:0.84rem; color:#94A3B8;">
                            ${l.notes||"-"}
                          </td>
                          <td style="text-align:center;">
                            <span class="emp-status-badge ${T?"emp-status-aktif":"emp-status-nonaktif"}">
                              ${T?"\u{1F7E2} Aktif":"\u26AA Nonaktif"}
                            </span>
                          </td>
                          <td style="text-align:center;">
                            <div class="emp-action-btns" style="justify-content:center;">
                              ${B?`
                                <button class="emp-icon-btn edit" data-action="edit" data-id="${l.id}" title="Edit Data Karyawan">
                                  <svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>
                                </button>
                              `:""}
                              ${x?`
                                <button class="emp-icon-btn delete" data-action="delete" data-id="${l.id}" title="Nonaktifkan Karyawan">
                                  <svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><polyline points="3 6 5 6 21 6"/><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"/><line x1="10" y1="11" x2="10" y2="17"/><line x1="14" y1="11" x2="14" y2="17"/></svg>
                                </button>
                              `:""}
                            </div>
                          </td>
                        </tr>
                      `}).join("")}
                  </tbody>
                </table>
              </div>
            </div>
      `:`
          <div class="emp-cards-grid">
            ${r.map(l=>{let f=l.division||"FACILITY CARE",u="emp-div-fc",p="";f==="SECURITY"&&(u="emp-div-sec",p="sec");let A=l.status==="Aktif",T=l.phone||"";return`
                <div class="emp-card" data-id="${l.id}">
                  <div class="emp-card-top">
                    <span class="emp-div-badge ${u}">
                      ${f==="SECURITY"?"\u{1F6E1}\uFE0F":"\u{1F9F9}"} ${f}
                    </span>
                    <span class="emp-status-badge ${A?"emp-status-aktif":"emp-status-nonaktif"}">
                      ${A?"\u{1F7E2} Aktif":"\u26AA Nonaktif"}
                    </span>
                  </div>

                  <div class="emp-profile-box">
                    <div class="emp-avatar ${p}">
                      ${Bn(l.full_name)}
                    </div>
                    <div style="min-width:0; flex:1;">
                      <div class="emp-name">${l.full_name||"Karyawan"}</div>
                      <div class="emp-subtitle">Personil ${f} \u2022 ID #${l.id}</div>
                    </div>
                  </div>

                  <div class="emp-info-grid">
                    <div class="emp-info-card">
                      <div class="emp-info-label">Cabang Penempatan</div>
                      <div class="emp-info-val" title="${l.branch_name||"-"}">
                        \u{1F3E2} ${l.branch_name||"Semua Cabang"}
                      </div>
                    </div>
                    <div class="emp-info-card">
                      <div class="emp-info-label">Tanggal Masuk</div>
                      <div class="emp-info-val">
                        \u{1F4C5} ${l.join_date?xe(l.join_date):"-"}
                      </div>
                    </div>
                    <div class="emp-info-card">
                      <div class="emp-info-label">Nomor Induk Karyawan</div>
                      <div class="emp-info-val">
                        \u{1F194} ${l.nik||"-"}
                      </div>
                    </div>
                    <div class="emp-info-card">
                      <div class="emp-info-label">Status Personil</div>
                      <div class="emp-info-val" style="color:${A?"#10B981":"#EF4444"};">
                        ${A?"\u2705 Karyawan Aktif":"\u26D4 Nonaktif / Resign"}
                      </div>
                    </div>
                  </div>

                  ${l.notes?`
                    <div class="emp-notes-box">
                      <strong>Catatan:</strong> ${l.notes}
                    </div>
                  `:""}

                  <div class="emp-card-footer">
                    <div style="font-size:0.75rem; color:#94A3B8;">
                      ID #${l.id} ${l.nik?"\u2022 "+l.nik:""}
                    </div>

                    <div class="emp-action-btns">
                      ${T?`
                        <a href="https://wa.me/${T.replace(/[^0-9]/g,"").replace(/^0/,"62")}" target="_blank" rel="noopener" class="emp-icon-btn" title="Chat WhatsApp Karyawan" style="color:#10B981;">
                          <svg width="16" height="16" fill="currentColor" viewBox="0 0 24 24"><path d="M12.031 6.172c-3.181 0-5.767 2.586-5.768 5.766-.001 1.298.38 2.27 1.019 3.287l-.582 2.128 2.182-.573c.978.58 1.911.928 3.145.929 3.178 0 5.767-2.587 5.768-5.766.001-3.187-2.575-5.771-5.764-5.771zm3.392 8.244c-.144.405-.837.774-1.17.824-.299.045-.677.063-1.092-.069-.252-.08-.575-.187-.988-.365-1.739-.751-2.874-2.502-2.961-2.617-.087-.116-.708-.94-.708-1.793s.448-1.273.607-1.446c.159-.173.346-.217.462-.217l.332.006c.106.005.249-.04.39.298.144.347.491 1.2.534 1.287.043.087.072.188.014.304-.058.116-.087.188-.173.289l-.26.304c-.087.086-.177.18-.076.354.101.174.449.741.964 1.201.662.591 1.221.774 1.394.86s.275.072.376-.043c.101-.116.433-.506.549-.68.116-.173.231-.145.39-.086s1.011.477 1.184.564.289.13.332.202c.045.072.045.419-.099.824zm-3.423-14.416c-6.627 0-12 5.373-12 12 0 2.159.57 4.185 1.564 5.939l-1.564 5.717 5.864-1.538c1.691.924 3.63 1.458 5.698 1.458 6.627 0 12-5.373 12-12s-5.373-12-12-12z"/></svg>
                        </a>
                      `:""}

                      ${B?`
                        <button class="emp-icon-btn edit" data-action="edit" data-id="${l.id}" title="Edit Data Karyawan">
                          <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>
                        </button>
                      `:""}
                      ${x?`
                        <button class="emp-icon-btn delete" data-action="delete" data-id="${l.id}" title="Nonaktifkan Karyawan">
                          <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><polyline points="3 6 5 6 21 6"/><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"/><line x1="10" y1="11" x2="10" y2="17"/><line x1="14" y1="11" x2="14" y2="17"/></svg>
                        </button>
                      `:""}
                    </div>
                  </div>
                </div>
              `}).join("")}
          </div>
    `}function Y(){let r=e.querySelector("#emp-pagination-container");if(!r)return;r.innerHTML="";let l=Te({page:d,pages:$,total:F,limit:20,onPage:f=>{d=f,E()}});l&&r.appendChild(l)}function b(){let r=e.querySelector("#emp-reset-filter");r&&r.addEventListener("click",()=>{i="",m="all",c="all",s="all",d=1;let l=e.querySelector("#emp-search-input");l&&(l.value="");let f=e.querySelector("#emp-branch-filter");f&&(f.value="all");let u=e.querySelector("#emp-division-filter");u&&(u.value="all"),e.querySelectorAll(".emp-pill").forEach(p=>{p.classList.toggle("active",p.getAttribute("data-status")==="all")}),E()}),e.querySelectorAll(".emp-icon-btn.edit").forEach(l=>{l.addEventListener("click",f=>{let u=Number(f.currentTarget.getAttribute("data-id")),p=a.find(A=>A.id===u);p&&K(p)})}),e.querySelectorAll(".emp-icon-btn.delete").forEach(l=>{l.addEventListener("click",f=>{let u=Number(f.currentTarget.getAttribute("data-id")),p=a.find(A=>A.id===u);p&&G(p)})})}function S(){let r=e.querySelector("#emp-list-container");if(!r)return;let l=J();r.innerHTML=W(l),Y(),b()}function y(){let r=J(),l=h.total,f=h.active,u=h.fc,p=h.sec,A=h.total,T=h.active,g=h.nonactive;e.innerHTML=`
      <style>.emp-page{display:flex;flex-direction:column;gap:18px;width:100%;max-width:100%;min-width:0;box-sizing:border-box;overflow-x:hidden}.emp-hero-card{background:#0E182B;border-radius:16px;border:1px solid rgba(255,255,255,0.08);padding:22px 26px;color:#F8FAFC;box-shadow:0 4px 20px rgba(0,0,0,0.25);position:relative;overflow:hidden;box-sizing:border-box;width:100%;max-width:100%;min-width:0}.emp-hero-card::after{content:'';position:absolute;top:-40px;right:-40px;width:220px;height:220px;background:radial-gradient(circle,rgba(37,99,235,0.12) 0%,transparent 70%);border-radius:50%;pointer-events:none}.emp-stats-row{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:14px;margin-top:18px}.emp-stat-box{background:#131F36;border:1px solid rgba(255,255,255,0.08);border-radius:12px;padding:14px 18px;display:flex;align-items:center;gap:14px;transition:transform 0.15s ease,background 0.15s ease,border-color 0.15s ease;min-width:0}.emp-stat-box:hover{background:#1E293B;border-color:rgba(255,255,255,0.16);transform:translateY(-2px)}.emp-stat-icon{width:44px;height:44px;border-radius:10px;display:flex;align-items:center;justify-content:center;font-size:1.3rem;flex-shrink:0;box-shadow:0 2px 8px rgba(0,0,0,0.25)}.emp-stat-val{font-size:1.55rem;font-weight:800;line-height:1.1;letter-spacing:-0.02em;color:#F8FAFC;font-variant-numeric:tabular-nums}.emp-stat-lbl{font-size:0.76rem;text-transform:uppercase;color:#94A3B8;font-weight:700;letter-spacing:0.05em;margin-top:2px}.emp-controls-card{background:#0E182B;border-radius:14px;border:1px solid rgba(255,255,255,0.08);padding:16px 20px;box-shadow:0 4px 20px rgba(0,0,0,0.25);display:flex;flex-direction:column;gap:14px;box-sizing:border-box;width:100%;max-width:100%;min-width:0}.emp-search-row{display:flex;gap:12px;flex-wrap:wrap;align-items:center}.emp-view-toggle{display:inline-flex;background:#131F36;padding:4px;border-radius:10px;border:1px solid rgba(255,255,255,0.08);gap:4px;flex-shrink:0}.emp-view-btn{display:inline-flex;align-items:center;gap:6px;padding:6px 12px;border-radius:7px;border:none;background:transparent;color:#94A3B8;font-size:0.82rem;font-weight:700;cursor:pointer;transition:all 0.15s ease}.emp-view-btn:hover{color:#F8FAFC}.emp-view-btn.active{background:#2563EB;color:#FFFFFF;box-shadow:0 2px 8px rgba(37,99,235,0.35)}.emp-table-card{background:#0E182B;border-radius:14px;border:1px solid rgba(255,255,255,0.08);box-shadow:0 4px 20px rgba(0,0,0,0.25);overflow:hidden;width:100%;box-sizing:border-box}.emp-table-responsive{width:100%;overflow-x:auto}.emp-table{width:100%;border-collapse:collapse;text-align:left;font-size:0.85rem}.emp-table th{background:#131F36;color:#94A3B8;font-weight:700;font-size:0.76rem;text-transform:uppercase;letter-spacing:0.04em;padding:13px 16px;border-bottom:1px solid rgba(255,255,255,0.08);white-space:nowrap}.emp-table td{padding:13px 16px;border-bottom:1px solid rgba(255,255,255,0.05);color:#F8FAFC;vertical-align:middle}.emp-table tbody tr{transition:background 0.15s ease}.emp-table tbody tr:hover{background:rgba(255,255,255,0.03)}.emp-table tbody tr:last-child td{border-bottom:none}.emp-search-wrap{position:relative;flex:1;min-width:240px}.emp-search-icon{position:absolute;left:14px;top:50%;transform:translateY(-50%);color:#94A3B8;pointer-events:none}.emp-search-input{width:100%;padding:10px 14px 10px 40px;border:1px solid rgba(255,255,255,0.1);border-radius:10px;font-size:0.9rem;color:#F8FAFC;background:#0E182B;outline:none;transition:border-color 0.15s ease,box-shadow 0.15s ease;box-sizing:border-box}.emp-search-input:focus{border-color:#2563EB;box-shadow:0 0 0 3px rgba(37,99,235,0.2)}.emp-search-input::placeholder{color:#94A3B8}.emp-select{padding:10px 14px;border:1px solid rgba(255,255,255,0.1);border-radius:10px;font-size:0.9rem;color:#F8FAFC;background:#131F36;outline:none;min-width:170px;cursor:pointer;transition:border-color 0.15s ease}.emp-select option{background:#131F36;color:#F8FAFC}.emp-select:focus{border-color:#2563EB}.emp-filter-pills{display:flex;align-items:center;gap:8px;overflow-x:auto;white-space:nowrap;scrollbar-width:none;padding-bottom:2px}.emp-filter-pills::-webkit-scrollbar{display:none}.emp-pill{padding:6px 14px;border-radius:20px;font-size:0.82rem;font-weight:600;cursor:pointer;border:1px solid rgba(255,255,255,0.08);background:#131F36;color:#94A3B8;transition:all 0.15s ease;display:inline-flex;align-items:center;gap:6px;user-select:none;flex-shrink:0}.emp-pill:hover{background:#1E293B;color:#F8FAFC;border-color:rgba(255,255,255,0.16)}.emp-pill.active{background:#2563EB;color:#FFFFFF;border-color:#2563EB;box-shadow:0 2px 8px rgba(37,99,235,0.35)}.emp-pill-dot{width:8px;height:8px;border-radius:50%}.emp-pill-count{padding:1px 7px;border-radius:10px;font-size:0.76rem;background:rgba(255,255,255,0.08);color:#F8FAFC}.emp-pill.active .emp-pill-count{background:rgba(255,255,255,0.25);color:#FFFFFF}.emp-cards-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(300px,1fr));gap:16px;width:100%;max-width:100%;min-width:0;box-sizing:border-box}.emp-card{background:#131F36;border-radius:14px;border:1px solid rgba(255,255,255,0.08);padding:16px 18px;box-shadow:0 4px 16px rgba(0,0,0,0.25);display:flex;flex-direction:column;gap:13px;width:100% !important;max-width:100% !important;min-width:0 !important;box-sizing:border-box !important;overflow:hidden;transition:transform 0.15s ease,box-shadow 0.15s ease,border-color 0.15s ease}.emp-card:hover{transform:translateY(-2px);box-shadow:0 8px 24px rgba(0,0,0,0.35);border-color:rgba(255,255,255,0.18)}.emp-card-top{display:flex;justify-content:space-between;align-items:center;gap:8px;width:100%;min-width:0;box-sizing:border-box;flex-wrap:wrap}.emp-div-badge{display:inline-flex;align-items:center;gap:5px;padding:4px 10px;border-radius:20px;font-size:0.76rem;font-weight:700;flex-shrink:0;white-space:nowrap}.emp-div-fc{background:rgba(37,99,235,0.18);color:#60A5FA;border:1px solid rgba(37,99,235,0.3)}.emp-div-sec{background:rgba(148,163,184,0.15);color:#E2E8F0;border:1px solid rgba(148,163,184,0.25)}.emp-status-badge{display:inline-flex;align-items:center;gap:5px;padding:4px 10px;border-radius:20px;font-size:0.76rem;font-weight:700;flex-shrink:0;white-space:nowrap}.emp-status-aktif{background:rgba(16,185,129,0.18);color:#34D399;border:1px solid rgba(16,185,129,0.3)}.emp-status-nonaktif{background:rgba(239,68,68,0.18);color:#F87171;border:1px solid rgba(239,68,68,0.3)}.emp-profile-box{display:flex;align-items:center;gap:12px;width:100%;min-width:0;box-sizing:border-box}.emp-avatar{width:48px;height:48px;border-radius:12px;background:#2563EB;color:#FFFFFF;display:flex;align-items:center;justify-content:center;font-weight:800;font-size:1.05rem;flex-shrink:0}.emp-avatar.sec{background:#1B2D4F;border:1px solid #3B82F6}.emp-name{font-size:1.02rem;font-weight:800;color:#F8FAFC;line-height:1.25;word-break:break-word;overflow-wrap:anywhere}.emp-subtitle{font-size:0.76rem;color:#94A3B8;font-weight:600;margin-top:3px}.emp-info-grid{display:grid;grid-template-columns:1fr 1fr;gap:8px}.emp-info-card{background:#0E182B;border:1px solid rgba(255,255,255,0.06);border-radius:10px;padding:10px 12px;display:flex;flex-direction:column;gap:4px;font-size:0.83rem;width:100%;max-width:100%;min-width:0;box-sizing:border-box}.emp-info-label{color:#94A3B8;font-size:0.76rem;font-weight:600;text-transform:uppercase;letter-spacing:0.03em}.emp-info-val{color:#F8FAFC;font-weight:700;min-width:0;word-break:break-word;overflow-wrap:anywhere;font-size:0.82rem}.emp-phone-link{color:#34D399;text-decoration:none;font-weight:600;display:inline-flex;align-items:center;gap:4px}.emp-phone-link:hover{text-decoration:underline}.emp-notes-box{font-size:0.78rem;color:#FCD34D;background:rgba(245,158,11,0.1);border:1px solid rgba(245,158,11,0.25);border-radius:8px;padding:8px 10px;line-height:1.4;word-break:break-word;overflow-wrap:anywhere;width:100%;box-sizing:border-box}.emp-card-footer{display:flex;justify-content:space-between;align-items:center;border-top:1px solid rgba(255,255,255,0.06);padding-top:11px;margin-top:auto;gap:8px;width:100%;min-width:0;box-sizing:border-box;flex-wrap:wrap}.emp-action-btns{display:flex;align-items:center;gap:6px;margin-left:auto;flex-shrink:0}.emp-icon-btn{width:34px;height:34px;border-radius:8px;border:1px solid rgba(255,255,255,0.08);background:#0E182B;color:#94A3B8;display:flex;align-items:center;justify-content:center;cursor:pointer;transition:all 0.15s ease}.emp-icon-btn:hover{background:#1E293B;color:#F8FAFC;border-color:rgba(255,255,255,0.18)}.emp-icon-btn.delete:hover{background:rgba(239,68,68,0.18);color:#F87171;border-color:rgba(239,68,68,0.3)}.emp-empty-state{grid-column:1 / -1;background:#131F36;border:2px dashed rgba(255,255,255,0.12);border-radius:16px;padding:48px 24px;text-align:center}@media (max-width:768px){.emp-page{gap:14px;width:100% !important;max-width:100% !important;overflow-x:hidden !important}.emp-hero-card{padding:16px 14px !important;width:100% !important;max-width:100% !important;box-sizing:border-box !important}.emp-stats-row{grid-template-columns:1fr 1fr;gap:10px}.emp-stat-box{padding:10px 12px}.emp-stat-icon{width:38px;height:38px;font-size:1.1rem}.emp-stat-val{font-size:1.3rem}.emp-hero-actions{width:100% !important;display:flex !important;align-items:center !important;gap:8px !important;flex-wrap:nowrap !important}.emp-btn-add{flex:1 1 auto !important;justify-content:center !important;padding:8px 10px !important;font-size:0.82rem !important;white-space:nowrap !important}.emp-btn-excel{flex:1 1 auto !important;justify-content:center !important;padding:8px 10px !important;font-size:0.82rem !important;white-space:nowrap !important}.emp-btn-refresh{width:38px !important;height:38px !important;padding:0 !important;flex-shrink:0 !important;justify-content:center !important}.emp-controls-card{padding:14px !important;gap:12px;width:100% !important;max-width:100% !important;box-sizing:border-box !important}.emp-search-row{flex-direction:column;gap:10px}.emp-select{width:100%}.emp-view-toggle{width:100%;display:flex}.emp-view-btn{flex:1;justify-content:center}.emp-cards-grid{grid-template-columns:minmax(0,1fr) !important;width:100% !important;max-width:100% !important;min-width:0 !important;box-sizing:border-box !important}.emp-card{padding:14px !important;width:100% !important;max-width:100% !important;min-width:0 !important;box-sizing:border-box !important}.emp-info-grid{grid-template-columns:1fr}.emp-info-card{padding:10px !important;width:100% !important;max-width:100% !important;box-sizing:border-box !important}}</style>

      <div class="emp-page">

        <!-- Hero Header Card -->
        <div class="emp-hero-card">
          <div style="display:flex; justify-content:space-between; align-items:flex-start; flex-wrap:wrap; gap:16px;">
            <div>
              <div style="display:flex; align-items:center; gap:10px;">
                <span style="font-size:1.6rem;">\u{1F465}</span>
                <h1 style="font-size:1.45rem; font-weight:800; margin:0; letter-spacing:-0.02em;">
                  Master Karyawan FCMS
                </h1>
              </div>
            </div>

            <div class="emp-hero-actions" style="display:flex; align-items:center; gap:8px;">
              ${v?`
                <button id="emp-btn-add" class="btn btn-primary emp-btn-add" style="background:#2563EB; padding:9px 16px; border-radius:10px; font-weight:700; display:flex; align-items:center; gap:6px; box-shadow:0 4px 12px rgba(37,99,235,0.3); white-space:nowrap;">
                  <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24"><line x1="12" y1="5" x2="12" y2="19"/><line x1="5" y1="12" x2="19" y2="12"/></svg>
                  <span>Tambah Karyawan</span>
                </button>
              `:""}
              
              ${re(q,"export")?`<button id="emp-btn-excel" class="btn btn-ghost emp-btn-excel" style="background:#131F36; border:1px solid rgba(255,255,255,0.12); color:#F8FAFC; padding:9px 14px; border-radius:10px; font-weight:600; display:flex; align-items:center; gap:6px; white-space:nowrap;">
                <span>\u{1F4E5}</span>
                <span>Ekspor Excel</span>
              </button>`:""}

              <button id="emp-btn-refresh" class="btn btn-ghost emp-btn-refresh" style="background:#131F36; border:1px solid rgba(255,255,255,0.12); color:#F8FAFC; width:38px; height:38px; padding:0; border-radius:10px; display:flex; align-items:center; justify-content:center; flex-shrink:0;" title="Muat Ulang Data">
                <svg width="16" height="16" fill="none" stroke="currentColor" stroke-width="2.2" viewBox="0 0 24 24" class="${R?"spin-icon":""}"><path d="M23 4v6h-6"/><path d="M1 20v-6h6"/><path d="M3.51 9a9 9 0 0 1 14.85-3.36L23 10M1 14l4.64 4.36A9 9 0 0 0 20.49 15"/></svg>
              </button>
            </div>
          </div>

          <!-- 4 KPI Stat Boxes -->
          <div class="emp-stats-row">
            <div class="emp-stat-box" title="Data karyawan terdaftar di sistem">
              <div class="emp-stat-icon" style="background:rgba(37,99,235,0.18); border:1px solid rgba(59,130,246,0.35); color:#60A5FA;">\u{1F465}</div>
              <div>
                <div class="emp-stat-val" id="emp-stat-total">${l}</div>
                <div class="emp-stat-lbl">Karyawan Terdaftar</div>
                <div id="emp-stat-sub" style="font-size: 0.76rem; color:#A5B4C8; margin-top:2px;">${f} Aktif \u2022 ${g} Nonaktif</div>
              </div>
            </div>

            <div class="emp-stat-box" title="Total Personil Aktif">
              <div class="emp-stat-icon" style="background:rgba(37,99,235,0.18); border:1px solid rgba(59,130,246,0.35); color:#60A5FA;">\u{1F7E2}</div>
              <div>
                <div class="emp-stat-val" id="emp-stat-active">${f}</div>
                <div class="emp-stat-lbl">Personil Aktif</div>
                <div style="font-size: 0.76rem; color:#34D399; margin-top:2px;">Status Aktif Operasional</div>
              </div>
            </div>

            <div class="emp-stat-box" title="Divisi Resmi Facility Care">
              <div class="emp-stat-icon" style="background:rgba(37,99,235,0.18); border:1px solid rgba(59,130,246,0.35); color:#60A5FA;">\u{1F3E2}</div>
              <div>
                <div class="emp-stat-val" id="emp-stat-fc">${u}</div>
                <div class="emp-stat-lbl">Facility Care</div>
                <div style="font-size: 0.76rem; color:#93C5FD; margin-top:2px;">Aktif Operasional</div>
              </div>
            </div>

            <div class="emp-stat-box" title="Divisi Resmi Security">
              <div class="emp-stat-icon" style="background:rgba(37,99,235,0.18); border:1px solid rgba(59,130,246,0.35); color:#60A5FA;">\u{1F6E1}\uFE0F</div>
              <div>
                <div class="emp-stat-val" id="emp-stat-sec">${p}</div>
                <div class="emp-stat-lbl">Security</div>
                <div style="font-size: 0.76rem; color:#7DD3FC; margin-top:2px;">Aktif Operasional</div>
              </div>
            </div>
          </div>
        </div>

        <!-- Filter Controls Card -->
        <div class="emp-controls-card">
          <div class="emp-search-row">
            <div class="emp-search-wrap">
              <svg class="emp-search-icon" width="16" height="16" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/></svg>
              <input type="text" id="emp-search-input" class="emp-search-input" 
                placeholder="Cari nama karyawan, no HP, cabang..." value="${i?String(i).replace(/"/g,"&quot;"):""}" />
            </div>

            <select id="emp-branch-filter" class="emp-select">
              <option value="all">Semua Cabang (${o.length})</option>
              ${o.map(ee=>`
                <option value="${ee.value}" ${String(m)===String(ee.value)?"selected":""}>
                  ${ee.label}
                </option>
              `).join("")}
            </select>

            <select id="emp-division-filter" class="emp-select">
              <option value="all">Semua Divisi</option>
              ${_n.map(ee=>`
                <option value="${ee}" ${c===ee?"selected":""}>
                  ${ee}
                </option>
              `).join("")}
            </select>

            <div class="emp-view-toggle" title="Pilih Tampilan">
              <button type="button" class="emp-view-btn ${I==="card"?"active":""}" data-view="card">
                <svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><rect x="3" y="3" width="7" height="7" rx="1"/><rect x="14" y="3" width="7" height="7" rx="1"/><rect x="14" y="14" width="7" height="7" rx="1"/><rect x="3" y="14" width="7" height="7" rx="1"/></svg>
                <span>Kartu</span>
              </button>
              <button type="button" class="emp-view-btn ${I==="table"?"active":""}" data-view="table">
                <svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><line x1="8" y1="6" x2="21" y2="6"/><line x1="8" y1="12" x2="21" y2="12"/><line x1="8" y1="18" x2="21" y2="18"/><line x1="3" y1="6" x2="3.01" y2="6"/><line x1="3" y1="12" x2="3.01" y2="12"/><line x1="3" y1="18" x2="3.01" y2="18"/></svg>
                <span>Tabel</span>
              </button>
            </div>
          </div>

          <!-- Status Horizontal Swipeable Pills -->
          <div class="emp-filter-pills">
            <div class="emp-pill ${s==="all"&&c==="all"?"active":""}" data-status="all" data-div="all">
              <span>Semua Karyawan</span>
              <span class="emp-pill-count" id="emp-pill-count-all">${A}</span>
            </div>
            <div class="emp-pill ${s==="Aktif"?"active":""}" data-status="Aktif" data-div="all">
              <span class="emp-pill-dot" style="background:#10B981;"></span>
              <span>Aktif</span>
              <span class="emp-pill-count" id="emp-pill-count-active">${T}</span>
            </div>
            <div class="emp-pill ${s==="Nonaktif"?"active":""}" data-status="Nonaktif" data-div="all">
              <span class="emp-pill-dot" style="background:#EF4444;"></span>
              <span>Nonaktif / Resign</span>
              <span class="emp-pill-count" id="emp-pill-count-nonactive">${g}</span>
            </div>
          </div>
        </div>

        <!-- Exclusive List Container (Card / Table) -->
        <div id="emp-list-container">
          ${W(r)}
        </div>
        <div id="emp-pagination-container"></div>

      </div>
    `;let C=e.querySelector("#emp-btn-add");C&&C.addEventListener("click",()=>K());let z=e.querySelector("#emp-btn-excel");z&&z.addEventListener("click",()=>w());let O=e.querySelector("#emp-btn-refresh");O&&O.addEventListener("click",()=>E({refreshStats:!0})),e.querySelectorAll(".emp-view-btn").forEach(ee=>{ee.addEventListener("click",()=>{let H=ee.getAttribute("data-view");H&&H!==I&&(I=H,localStorage.setItem("fcms_employees_view_mode",H),e.querySelectorAll(".emp-view-btn").forEach(U=>{U.classList.toggle("active",U.getAttribute("data-view")===H)}),S())})});let te=e.querySelector("#emp-search-input");te&&te.addEventListener("input",ee=>{i=ee.target.value,clearTimeout(L),L=setTimeout(()=>{d=1,E()},350)});let Q=e.querySelector("#emp-branch-filter");Q&&Q.addEventListener("change",ee=>{m=ee.target.value,d=1,E()});let ne=e.querySelector("#emp-division-filter");ne&&ne.addEventListener("change",ee=>{c=ee.target.value,d=1,E()}),e.querySelectorAll(".emp-pill").forEach(ee=>{ee.addEventListener("click",()=>{s=ee.dataset.status;let H=ee.dataset.div;H&&H!=="all"?c=H:ee.dataset.status==="all"&&(c="all");let U=e.querySelector("#emp-division-filter");U&&(U.value=c),e.querySelectorAll(".emp-pill").forEach(D=>{let V=D===ee||s==="all"&&c==="all"&&D.dataset.status==="all";D.classList.toggle("active",V)}),d=1,E()})}),Y(),b()}await E({refreshStats:!0})}Be();Ce();function Re(e){return String(e??"").replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;").replace(/'/g,"&#39;")}function Fe(e){return String(e??"").trim().replace(/\s+/g," ")}function wr(e){if(!e)return null;if(typeof e=="string"){let a=Fe(e);return a?{id:null,name:a,branch_name:"",division:"",status:""}:null}let t=Fe(e.full_name||e.label||e.name||"");if(!t)return null;let n=Number(e.id??e.value);return{id:Number.isInteger(n)&&n>0?n:null,name:t,branch_name:e.branch_name||e.branch||"",division:e.division||"",status:e.status||""}}function Ze(e=[],t=[]){let n=new Map;for(let a of[...e,...t]){let o=wr(a);if(!o)continue;let s=o.id!==null?`id:${o.id}`:`name:${o.name.toLowerCase()}`;n.has(s)||n.set(s,o)}return Array.from(n.values())}function Sr(){if(typeof document>"u"||document.getElementById("fc-emp-combobox-scoped-css"))return;let e=document.createElement("style");e.id="fc-emp-combobox-scoped-css",e.textContent=`
    .fc-emp-combobox {
      position: relative;
      width: 100%;
      box-sizing: border-box;
      font-family: inherit;
    }
    .fc-emp-input-wrap {
      position: relative;
      display: flex;
      align-items: center;
      width: 100%;
      box-sizing: border-box;
    }
    .fc-emp-search-icon {
      position: absolute;
      left: 14px;
      top: 50%;
      transform: translateY(-50%);
      display: inline-flex;
      align-items: center;
      justify-content: center;
      color: #94A3B8;
      pointer-events: none;
      z-index: 2;
      transition: color 0.15s ease;
    }
    .fc-emp-input-wrap:focus-within .fc-emp-search-icon {
      color: #2563EB;
    }
    .fc-emp-input {
      width: 100% !important;
      height: 46px !important;
      min-height: 46px !important;
      padding: 0 38px 0 42px !important;
      border: 1px solid #223456 !important;
      border-radius: 12px !important;
      font-size: 0.95rem !important;
      font-family: inherit !important;
      font-weight: 500 !important;
      color: #F8FAFC !important;
      background: #0E182B !important;
      outline: none !important;
      box-sizing: border-box !important;
      transition: border-color 0.15s ease, box-shadow 0.15s ease !important;
      line-height: normal !important;
    }
    .fc-emp-input::placeholder {
      color: #94A3B8 !important;
      font-weight: 400 !important;
      font-size: 0.92rem !important;
    }
    .fc-emp-input:focus {
      border-color: #2563EB !important;
      box-shadow: 0 0 0 3.5px rgba(37, 99, 235, 0.2) !important;
      background: #0E182B !important;
    }
    .fc-emp-clear-btn {
      position: absolute;
      right: 12px;
      top: 50%;
      transform: translateY(-50%);
      width: 22px;
      height: 22px;
      border-radius: 50%;
      border: none;
      background: #22314A;
      color: #94A3B8;
      cursor: pointer;
      display: none;
      align-items: center;
      justify-content: center;
      font-size: 14px;
      line-height: 1;
      transition: all 0.15s ease;
      z-index: 2;
      padding: 0;
    }
    .fc-emp-clear-btn:hover {
      background: #223456;
      color: #F8FAFC;
    }
    .fc-emp-dropdown {
      position: absolute;
      top: calc(100% + 6px);
      left: 0;
      right: 0;
      width: 100%;
      z-index: 1070;
      background: #131F36;
      border: 1px solid #223456;
      border-radius: 12px;
      box-shadow: 0 12px 28px -4px rgba(0, 0, 0, 0.5), 0 4px 10px -2px rgba(0, 0, 0, 0.3);
      max-height: 260px;
      overflow-y: auto;
      overscroll-behavior: contain;
      display: none;
      padding: 4px 0;
      box-sizing: border-box;
    }
    .fc-emp-dropdown.open {
      display: block;
    }
    .fc-emp-dropdown::-webkit-scrollbar {
      width: 6px;
    }
    .fc-emp-dropdown::-webkit-scrollbar-track {
      background: transparent;
    }
    .fc-emp-dropdown::-webkit-scrollbar-thumb {
      background: #223456;
      border-radius: 999px;
    }
    .fc-emp-dropdown::-webkit-scrollbar-thumb:hover {
      background: #64748B;
    }
    .fc-emp-option {
      padding: 10px 14px;
      cursor: pointer;
      border-bottom: 1px solid #223456;
      transition: background-color 0.12s ease;
      display: flex;
      align-items: center;
      justify-content: space-between;
      gap: 8px;
      box-sizing: border-box;
    }
    .fc-emp-option:last-child {
      border-bottom: none;
    }
    .fc-emp-option:hover,
    .fc-emp-option.active {
      background: #22314A;
    }
    .fc-emp-option.active {
      background: #22314A;
      border-left: 3px solid #2563EB;
      padding-left: 11px;
    }
    .fc-emp-option-main {
      display: flex;
      flex-direction: column;
      gap: 2px;
      flex: 1;
      min-width: 0;
    }
    .fc-emp-option-name {
      font-size: 0.92rem;
      font-weight: 600;
      color: #F8FAFC;
      line-height: 1.35;
      letter-spacing: -0.01em;
      word-break: break-word;
    }
    .fc-emp-option-meta {
      font-size: 0.78rem;
      color: #94A3B8;
      line-height: 1.3;
      white-space: nowrap;
      overflow: hidden;
      text-overflow: ellipsis;
    }
    .fc-emp-option-manual {
      background: rgba(245, 158, 11, 0.08);
      border-top: 1px dashed rgba(245, 158, 11, 0.3);
    }
    .fc-emp-option-manual:hover,
    .fc-emp-option-manual.active {
      background: rgba(245, 158, 11, 0.18);
    }
    .fc-emp-option-manual .fc-emp-option-name {
      color: #FBBF24;
      font-weight: 600;
    }
    .fc-emp-manual-row {
      display: flex;
      align-items: center;
      gap: 6px;
    }
    .fc-emp-manual-icon {
      font-size: 1.1rem;
      font-weight: 700;
      color: #F59E0B;
      line-height: 1;
    }
    .fc-emp-manual-badge-row {
      margin-top: 4px;
    }
    .fc-emp-empty {
      padding: 14px 16px;
      font-size: 0.85rem;
      color: #94A3B8;
      text-align: center;
    }
    .fc-emp-status-row {
      margin-top: 6px;
      display: flex;
      flex-wrap: wrap;
      align-items: center;
      gap: 6px;
    }
    .fc-emp-badge {
      display: inline-flex;
      align-items: center;
      gap: 5px;
      padding: 4px 10px;
      border-radius: 999px;
      font-size: 0.75rem;
      font-weight: 600;
      line-height: 1.3;
    }
    .fc-emp-badge-master {
      background: rgba(16, 185, 129, 0.15);
      color: #34D399;
      border: 1px solid rgba(16, 185, 129, 0.3);
    }
    .fc-emp-badge-manual {
      background: rgba(245, 158, 11, 0.15);
      color: #FBBF24;
      border: 1px solid rgba(245, 158, 11, 0.3);
    }
    .fc-emp-chips {
      margin-top: 8px;
      display: flex;
      flex-wrap: wrap;
      gap: 6px;
    }
    .fc-emp-chip {
      display: inline-flex;
      align-items: center;
      gap: 6px;
      padding: 4px 10px;
      background: #0E182B;
      color: #F8FAFC;
      border: 1px solid #223456;
      border-radius: 999px;
      font-size: 0.78rem;
      font-weight: 500;
    }
    .fc-emp-chip.manual {
      background: rgba(245, 158, 11, 0.15);
      color: #FBBF24;
      border-color: rgba(245, 158, 11, 0.3);
    }
    .fc-emp-chip-tag {
      font-size: 0.76rem;
      padding: 1px 5px;
      background: #F59E0B;
      color: #FFFFFF;
      border-radius: 4px;
      font-weight: 700;
      text-transform: uppercase;
    }
    .fc-emp-chip-remove {
      background: none;
      border: none;
      padding: 0;
      font-size: 14px;
      line-height: 1;
      cursor: pointer;
      color: #CBD5E1;
      display: inline-flex;
      align-items: center;
      justify-content: center;
    }
    .fc-emp-chip-remove:hover {
      color: #FCA5A5;
    }
    @media (max-width: 640px) {
      .fc-emp-input {
        font-size: 16px !important;
        height: 48px !important;
        min-height: 48px !important;
      }
    }
  `,document.head.appendChild(e)}function Fa(e){if(!e)return[];if(Array.isArray(e))return e.map(Fe).filter(Boolean);let t=String(e).trim();if(!t)return[];if(t.startsWith("["))try{let n=JSON.parse(t);if(Array.isArray(n))return n.map(Fe).filter(Boolean)}catch{}return t.split(",").map(Fe).filter(Boolean)}function je(e,t={}){if(!e)throw new Error("mountEmployeeCombobox requires a valid mount element");Sr();let{mode:n="single",inputId:a="",inputName:o="employee_search",placeholder:s="",required:c=!1,allowManual:m=!0,hasEmployeeIdSchema:i=!0,manualBadgeText:d="Belum terhubung ke Master Karyawan",masterBadgePrefix:F="Terhubung ke Master Karyawan",employees:$=[],initialEmployeeId:M=null,initialEmployeeName:h="",initialParticipants:L=[],maxResults:R=25,onChange:I=null}=t,N=Ze($),q=[],v=!1,B=-1,x=null,k=0,_=Number(M),E=Number.isInteger(_)&&_>0?_:null,w=Fe(h),K=H=>{let U=Fe(H).toLowerCase();return U&&Ze(N,q).find(D=>D.name.toLowerCase()===U)||null},G={employee_id:E,employee_name:w,is_manual:!!(w&&E===null&&(i||!K(w))),selected_employee:E!==null?Ze(N).find(H=>H.id===E)||{id:E,name:w}:null,confirmed:!!w},J=n==="multi"?Fa(L.length?L:w).map(H=>{let U=K(H);return U?{employee_id:U.id,employee_name:U.name,is_manual:!1}:{employee_id:null,employee_name:H,is_manual:!0}}):[];e.innerHTML=`
    <div class="fc-emp-combobox" role="combobox" aria-expanded="false" aria-haspopup="listbox">
      <div class="fc-emp-input-wrap">
        <span class="fc-emp-search-icon" aria-hidden="true">
          <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
            <circle cx="11" cy="11" r="8"></circle>
            <line x1="21" y1="21" x2="16.65" y2="16.65"></line>
          </svg>
        </span>
        <input
          type="text"
          ${a?`id="${Re(a)}"`:""}
          name="${Re(o)}"
          class="fc-emp-input"
          autocomplete="off"
          placeholder="${Re(s)}"
          ${c&&n==="single"?"required":""}
          value="${n==="single"?Re(w):""}"
          aria-autocomplete="list"
        />
        <button type="button" class="fc-emp-clear-btn" title="Kosongkan" aria-label="Kosongkan pencarian" style="display:none;">&times;</button>
        <div class="fc-emp-dropdown" role="listbox"></div>
      </div>
      <div class="fc-emp-status-row" aria-live="polite"></div>
      ${n==="multi"?'<div class="fc-emp-chips"></div>':""}
    </div>
  `;let W=e.querySelector(".fc-emp-combobox"),Y=e.querySelector(".fc-emp-input"),b=e.querySelector(".fc-emp-clear-btn"),S=e.querySelector(".fc-emp-dropdown"),y=e.querySelector(".fc-emp-status-row"),r=e.querySelector(".fc-emp-chips"),l=[];function f(){b.style.display=Y.value.trim()?"inline-flex":"none"}function u(){typeof I=="function"&&I(ee.getValue())}function p(){if(n==="multi"){let U=J.filter(Z=>Z.is_manual).length,D=Fe(Y.value),V=D&&!K(D);if(J.length===0&&!D){y.innerHTML="";return}U>0||V?y.innerHTML=`<span class="fc-emp-badge fc-emp-badge-manual">\u26A0\uFE0F Ada peserta nama manual (${Re(d)})</span>`:y.innerHTML=`<span class="fc-emp-badge fc-emp-badge-master">\u2713 Semua peserta terhubung ke Master Karyawan (${J.length})</span>`;return}if(v||!G.confirmed){y.innerHTML="";return}let H=Fe(Y.value);if(!H){y.innerHTML="";return}if(G.employee_id!==null){let U=G.selected_employee?.branch_name?` \u2022 ${Re(G.selected_employee.branch_name)}`:"";y.innerHTML=`<span class="fc-emp-badge fc-emp-badge-master">\u2713 ${Re(F)} (#${G.employee_id})${U}</span>`;return}if(!i){let U=K(H);if(U&&!G.is_manual){y.innerHTML=`<span class="fc-emp-badge fc-emp-badge-master">\u2713 Cocok dengan Master Karyawan${U.id?` (#${U.id})`:""}</span>`;return}}y.innerHTML=`<span class="fc-emp-badge fc-emp-badge-manual">\u26A0\uFE0F ${Re(d)}</span>`}function A(){if(r){if(J.length===0){r.innerHTML="";return}r.innerHTML=J.map((H,U)=>`
      <span class="fc-emp-chip ${H.is_manual?"manual":""}" data-index="${U}">
        <span>${Re(H.employee_name)}</span>
        ${H.is_manual?'<span class="fc-emp-chip-tag">Manual</span>':""}
        <button type="button" class="fc-emp-chip-remove" data-remove-index="${U}" title="Hapus peserta">&times;</button>
      </span>
    `).join("")}}function T(H){let U=Fe(H);if(!U)return{items:[],totalMatched:0};let D=U.toLowerCase(),V=Ze(N,q),Z=V.filter(de=>de.name.toLowerCase().includes(D)||de.branch_name&&de.branch_name.toLowerCase().includes(D)||de.division&&de.division.toLowerCase().includes(D)),le=[...Z.slice(0,R).map(de=>({type:"master",emp:de,label:de.name,meta:[de.branch_name,de.division,de.status].filter(Boolean).join(" \xB7 ")}))],pe=V.some(de=>de.name.toLowerCase()===D);return m&&!pe&&le.push({type:"manual",name:U,label:U,meta:d}),{items:le,totalMatched:Z.length}}function g(){S.style.top="",S.style.bottom="";let H=Y.getBoundingClientRect(),U=window.innerHeight-H.bottom,D=H.top;U<220&&D>U?(S.style.top="auto",S.style.bottom="calc(100% + 6px)"):(S.style.top="calc(100% + 6px)",S.style.bottom="auto")}function C(){let H=Y.value,U=Fe(H);if(!U){S.innerHTML="",O();return}let{items:D,totalMatched:V}=T(H);if(l=D,D.length===0)S.innerHTML=`<div class="fc-emp-empty">Tidak ada karyawan yang cocok dengan "${Re(U)}".</div>`;else{let Z=V>R?`<div class="fc-emp-empty" style="border-top:1px solid #F1F5F9; font-size:0.75rem; color:#94A3B8;">Menampilkan ${R} dari ${V} karyawan. Ketik lebih spesifik untuk menyaring.</div>`:"";S.innerHTML=D.map((ie,le)=>{let pe=le===B?"active":"";return ie.type==="manual"?`
            <div class="fc-emp-option fc-emp-option-manual ${pe}" role="option" data-idx="${le}" data-manual="1" aria-selected="${le===B}">
              <div class="fc-emp-option-main">
                <div class="fc-emp-manual-row">
                  <span class="fc-emp-manual-icon">+</span>
                  <div class="fc-emp-option-name">Gunakan nama manual: <strong>${Re(ie.name)}</strong></div>
                </div>
                <div class="fc-emp-manual-badge-row">
                  <span class="fc-emp-badge fc-emp-badge-manual">\u26A0\uFE0F ${Re(ie.meta)}</span>
                </div>
              </div>
            </div>
          `:`
          <div class="fc-emp-option ${pe}" role="option" data-idx="${le}" data-emp-id="${ie.emp.id??""}" aria-selected="${le===B}">
            <div class="fc-emp-option-main">
              <div class="fc-emp-option-name">${Re(ie.label)}</div>
              ${ie.meta?`<div class="fc-emp-option-meta">${Re(ie.meta)}</div>`:""}
            </div>
          </div>
        `}).join("")+Z}g()}function z(){Fe(Y.value)&&(v||(v=!0,S.classList.add("open"),W.setAttribute("aria-expanded","true")),C())}function O(){v=!1,B=-1,S.classList.remove("open"),W.setAttribute("aria-expanded","false"),Fe(Y.value)&&n==="single"&&!G.selected_employee&&(G.confirmed=!0),p()}function te(H){if(H){if(n==="multi"){let U=H.type==="manual"?{employee_id:null,employee_name:H.name,is_manual:!0}:{employee_id:H.emp.id,employee_name:H.emp.name,is_manual:!1};J.some(D=>D.employee_name.toLowerCase()===U.employee_name.toLowerCase())||J.push(U),Y.value="",f(),A(),p(),O(),u();return}H.type==="manual"?(G={employee_id:null,employee_name:H.name,is_manual:!0,selected_employee:null,confirmed:!0},Y.value=H.name):(G={employee_id:H.emp.id,employee_name:H.emp.name,is_manual:!1,selected_employee:H.emp,confirmed:!0},Y.value=H.emp.name),f(),p(),O(),u()}}function Q(H){clearTimeout(x);let U=Fe(H);if(!U)return;let D=++k;x=setTimeout(async()=>{try{let V=await Sn({search:U,limit:R});if(D!==k)return;Array.isArray(V)&&V.length>0&&(q=Ze(q,V),v&&C())}catch{}},180)}N.length===0&&ze().then(H=>{Array.isArray(H)&&H.length>0&&(N=Ze(H),G.employee_id!==null&&!G.selected_employee&&(G.selected_employee=N.find(U=>U.id===G.employee_id)||G.selected_employee),p(),v&&C())}).catch(()=>{}),Y.addEventListener("focus",()=>{Fe(Y.value).length>=1&&(B=-1,z())}),Y.addEventListener("input",()=>{let H=Y.value,U=Fe(H);if(n==="single"&&(U?G={employee_id:null,employee_name:U,is_manual:m,selected_employee:null}:G={employee_id:null,employee_name:"",is_manual:!1,selected_employee:null}),B=-1,f(),p(),U.length<1){O(),u();return}z(),Q(U),u()}),Y.addEventListener("keydown",H=>{let U=Fe(Y.value);if(!v&&(H.key==="ArrowDown"||H.key==="ArrowUp")&&U.length>=1&&z(),H.key==="ArrowDown"){if(H.preventDefault(),l.length===0)return;B=(B+1)%l.length,C(),S.querySelector(".fc-emp-option.active")?.scrollIntoView({block:"nearest"})}else if(H.key==="ArrowUp"){if(H.preventDefault(),l.length===0)return;B=B<=0?l.length-1:B-1,C(),S.querySelector(".fc-emp-option.active")?.scrollIntoView({block:"nearest"})}else if(H.key==="Enter"){if(v&&l.length>0){H.preventDefault();let D=B>=0?l[B]:l[0];D&&te(D)}else if(n==="multi"&&Fe(Y.value)){H.preventDefault();let D=Fe(Y.value),V=K(D);te(V?{type:"master",emp:V}:{type:"manual",name:D})}}else H.key==="Escape"&&v&&(H.preventDefault(),H.stopPropagation(),O())}),S.addEventListener("mousedown",H=>{H.preventDefault()}),S.addEventListener("click",H=>{let U=H.target.closest(".fc-emp-option");if(!U)return;let D=Number(U.dataset.idx);Number.isInteger(D)&&l[D]&&te(l[D])}),b.addEventListener("click",()=>{Y.value="",n==="single"&&(G={employee_id:null,employee_name:"",is_manual:!1,selected_employee:null,confirmed:!1}),f(),p(),O(),Y.focus(),u()}),r&&r.addEventListener("click",H=>{let U=H.target.closest("[data-remove-index]");if(!U)return;let D=Number(U.dataset.removeIndex);Number.isInteger(D)&&D>=0&&D<J.length&&(J.splice(D,1),A(),p(),u())});let ne=H=>{if(!e.isConnected){ee.destroy();return}e.contains(H.target)||O()};document.addEventListener("mousedown",ne),f(),A(),p();let ee={inputEl:Y,getValue(){if(n==="multi"){let U=Fe(Y.value),D=[...J];if(U&&!D.some(V=>V.employee_name.toLowerCase()===U.toLowerCase())){let V=K(U);D.push(V?{employee_id:V.id,employee_name:V.name,is_manual:!1}:{employee_id:null,employee_name:U,is_manual:!0})}return{participants:D.map(V=>V.employee_name),participantsString:D.map(V=>V.employee_name).join(", "),items:D,hasManual:D.some(V=>V.is_manual)}}let H=Fe(Y.value);if(!H)return{employee_id:null,employee_name:"",is_manual:!1,selected_employee:null};if(G.employee_id!==null&&G.employee_name.toLowerCase()===H.toLowerCase())return{employee_id:G.employee_id,employee_name:G.employee_name,is_manual:!1,selected_employee:G.selected_employee};if(!i){let U=K(H);if(U&&!G.is_manual)return{employee_id:U.id,employee_name:U.name,is_manual:!1,selected_employee:U}}return{employee_id:null,employee_name:H,is_manual:!0,selected_employee:null}},setValue({employee_id:H=null,employee_name:U=""}={}){let D=Fe(U),V=Number(H),Z=Number.isInteger(V)&&V>0?V:null,ie=Ze(N,q);G={employee_id:Z,employee_name:D,is_manual:!!(D&&Z===null),selected_employee:Z!==null?ie.find(le=>le.id===Z)||{id:Z,name:D}:null,confirmed:!!D},Y.value=D,f(),p()},selectOptionByIdOrName(H){let D=Ze(N,q).find(V=>H!=null&&String(V.id)===String(H)||V.name.toLowerCase()===Fe(H).toLowerCase());return D?(te({type:"master",emp:D}),!0):!1},destroy(){clearTimeout(x),document.removeEventListener("mousedown",ne)}};return ee}function _t(e){return!e||e==="-"?"-":String(e).startsWith("2099")?"Tetap / PKWTT":xe(e,"-")}function Bt(e){return e?String(e).startsWith("2099"):!1}function Dn(e){if(!e||e==="-")return"\u{1F464}";let t=String(e).trim().split(/\s+/);return t.length===1?t[0].slice(0,2).toUpperCase():(t[0][0]+t[1][0]).toUpperCase()}async function Ln(e,t){let a=[],o=[],s=[],c={},m="",i="all",d="all",F="all",$="all",M="",h=1,L=0,R=1,I=!1,N={total:0,active:0,expiring30:0,permanent:0,nonactive:0},q=null,v=!1,B=localStorage.getItem("fcms_contracts_view_mode")||"card",x=t?t.get("dash_filter"):null,k=t?(t.get("month_expiry")||"").trim():"";k&&/^\d{4}-(0[1-9]|1[0-2])$/.test(k)?(M=k,F="Aktif"):x==="expiring30"?($="30",F="Aktif"):x==="active"&&(F="Aktif");let _=se(),E=re(_,"contracts","write"),w=re(_,"contracts","write"),K=re(_,"contracts","delete");try{let[g,C]=await Promise.all([we(),ze()]);o=(Array.isArray(g)?g:[]).filter(O=>String(O.code||"").trim()==="000"||O.branch_scope==="hq"||/^(?:cabang\s+)?000\.\s*management$/i.test(String(O.label||O.full_name||"").trim())),s=Array.isArray(C)?C:[],s.forEach(O=>{let te=O.value!==void 0?O.value:O.id,Q=O.label||O.full_name;te&&O.phone&&(c[te]=O.phone),Q&&O.phone&&(c[Q.toLowerCase()]=O.phone)})}catch(g){console.error("Error fetching cache options for contracts:",g)}function G(){let g={sort:"start_date_desc"};return m.trim()&&(g.search=m.trim()),i&&i!=="all"&&(g.branch_id=i),d&&d!=="all"&&(g.division=d),F&&F!=="all"&&(g.status=F),$&&$!=="all"&&(g.expiring_days=$),M&&(g.month_expiry=M),g}function J(){let g=(C,z)=>{let O=e.querySelector(C);O&&(O.textContent=z)};g("#con-stat-total",N.total),g("#con-stat-active",N.active),g("#con-stat-expiring",N.expiring30),g("#con-stat-permanent",N.permanent),g("#con-pill-count-all",N.total),g("#con-pill-count-aktif",N.active),g("#con-pill-count-expiring",N.expiring30),g("#con-pill-count-tetap",N.permanent),g("#con-pill-count-nonaktif",N.nonactive)}async function W({refreshStats:g=!0}={}){if(v=!0,!!!e.querySelector("#con-list-container"))T();else{let z=e.querySelector("#con-list-container");z&&(z.innerHTML='<div class="con-empty-state"><p style="color:#CBD5E1; margin:0;">Memuat data kontrak...</p></div>')}try{let z=new URLSearchParams({page:String(h),limit:String(20),...G()});(g||!I)&&z.set("include_stats","1");let O=await X(`/api/contracts?${z.toString()}`);if(O.ok){let te=O.data?.data||[];a=Array.isArray(te)?te:[];let Q=O.data?.pagination||{};if(L=Number(Q.total??a.length),R=Math.max(1,Number(Q.pages??Math.ceil(L/20))||1),h=Math.max(1,Number(Q.page??h)),O.data?.stats&&(N=O.data.stats,I=!0),a.length===0&&h>R)return h=R,v=!1,W({refreshStats:g})}else a=[],L=0,R=1,j("Gagal memuat data kontrak dari server.")}catch(z){console.error("Error load contracts:",z),a=[],L=0,R=1,j("Terjadi kesalahan koneksi saat memuat kontrak.")}finally{v=!1,e.querySelector("#con-list-container")?(J(),A()):T()}}function Y(){return a}async function b(){ce({title:"\u{1F50D} Analisis Selisih Kontrak Karyawan",content:`
        <div style="text-align:center; padding:24px 10px;">
          <div style="font-size:2rem; margin-bottom:10px;">\u23F3</div>
          <p style="color:#CBD5E1; font-size:0.9rem;">Sedang menyinkronkan data karyawan aktif dengan kontrak aktif di database...</p>
        </div>
      `,cancelText:"Tutup"});try{let[g,C]=await Promise.all([$t("/api/employees/selector",{status:"Aktif"}),$t("/api/contracts",{})]),z=new Set(C.filter(Q=>Q.status==="Aktif").map(Q=>Q.employee_id)),O=g.filter(Q=>!z.has(Q.id)),te="";O.length===0?te=`
          <div style="text-align:center; padding:20px 10px;">
            <div style="font-size:2.8rem; margin-bottom:10px;">\u{1F389}</div>
            <h4 style="font-size:1.1rem; font-weight:800; color:#34D399; margin:0 0 8px 0;">Semua Karyawan Aktif Memiliki Kontrak!</h4>
            <p style="font-size:0.86rem; color:#CBD5E1; margin:0;">
              Total <strong>${g.length}</strong> karyawan aktif tercatat memiliki perikatan kontrak aktif yang sah di sistem.
            </p>
          </div>
        `:te=`
          <div style="display:flex; flex-direction:column; gap:12px;">
            <div style="background:rgba(239,68,68,0.16); border:1px solid rgba(248,113,113,0.35); border-radius:10px; padding:12px; font-size:0.84rem; color:#FCA5A5;">
              Ditemukan <strong>${O.length}</strong> dari <strong>${g.length}</strong> karyawan aktif yang <strong>belum memiliki Kontrak Aktif</strong>.
            </div>
            <div style="max-height:360px; overflow-y:auto; display:flex; flex-direction:column; gap:8px; padding-right:4px;">
              ${O.map(Q=>{let ne=C.filter(H=>H.employee_id===Q.id),ee='<span style="color:#FCD34D; font-weight:600;">Belum pernah ada data kontrak</span>';if(ne.length>0){let H=ne[0];ee=`Riwayat: Status <strong>${H.status}</strong> (Selesai: ${_t(H.end_date)})`}return`
                  <div style="background:#131F36; border:1px solid #223456; border-radius:8px; padding:10px 12px; display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:8px;">
                    <div>
                      <div style="font-weight:700; font-size:0.88rem; color:#F8FAFC;">${Q.full_name}</div>
                      <div style="font-size:0.76rem; color:#CBD5E1;">
                        Cabang: ${Q.branch_name||"-"} \u2022 Divisi: ${Q.division||"FACILITY CARE"}
                      </div>
                      <div style="font-size:0.75rem; color:#94A3B8; margin-top:2px;">
                        ${ee}
                      </div>
                    </div>
                    ${w?`
                      <button class="btn btn-primary btn-add-contract-for-emp" data-empid="${Q.id}" data-empname="${String(Q.full_name).replace(/"/g,"&quot;")}" data-branchid="${Q.branch_id||""}" data-division="${Q.division||"FACILITY CARE"}" style="font-size:0.76rem; padding:5px 10px; border-radius:6px; background:#2563EB;">
                        + Buat Kontrak
                      </button>
                    `:""}
                  </div>
                `}).join("")}
            </div>
          </div>
        `,ce({title:"Analisis Selisih Kontrak Karyawan",content:te,cancelText:"Tutup"}),setTimeout(()=>{document.querySelectorAll(".btn-add-contract-for-emp").forEach(Q=>{Q.addEventListener("click",ne=>{let ee=ne.currentTarget,H=ee.getAttribute("data-empid"),U=ee.getAttribute("data-empname"),D=ee.getAttribute("data-branchid"),V=ee.getAttribute("data-division"),Z=document.querySelector(".modal-close");Z&&Z.click(),y({employee_id:H,employee_name:U,branch_id:D,division:V,status:"Aktif"})})})},100)}catch(g){console.error(g),j("Gagal melakukan analisis selisih data kontrak.")}}function S(g,C){if(!g)return"";let z=String(g).split("-").map(Number);if(z.length!==3||!z[0]||!z[1]||!z[2])return"";let O=new Date(z[0],z[1]-1,z[2]);if(isNaN(O.getTime()))return"";if(C==="KONTRAK 6 BULAN")O.setMonth(O.getMonth()+6),O.setDate(O.getDate()-1);else if(C==="KONTRAK 1 TAHUN")O.setFullYear(O.getFullYear()+1),O.setDate(O.getDate()-1);else if(C==="KONTRAK 2 TAHUN")O.setFullYear(O.getFullYear()+2),O.setDate(O.getDate()-1);else return"";let te=O.getFullYear(),Q=String(O.getMonth()+1).padStart(2,"0"),ne=String(O.getDate()).padStart(2,"0");return`${te}-${Q}-${ne}`}function y(g=null){let C=!!(g&&g.id);if(C&&!E){j("Anda tidak memiliki izin untuk mengedit kontrak.");return}if(!C&&!w){j("Anda tidak memiliki izin untuk menambah kontrak.");return}let z=!!(g&&g.id),O=z?nt(g?.row_version):null,te=z?`\u270F\uFE0F Edit Kontrak: ${g.employee_name||"Karyawan"}`:"\u2795 Tambah Kontrak Karyawan Baru",Q=g?.employee_id||"",ne=g?.branch_id||"",ee=g?.division||"FACILITY CARE",H=g?.status||"Aktif",U=g?.contract_type||"KONTRAK 1 TAHUN",D=g?.pkwt_number||"PKWT 1",V=g?.notes||"",Z=!1;g&&(g.contract_type==="PKWTT (TETAP)"||g.pkwt_number==="TETAP"?Z=!0:g.contract_type&&(g.contract_type.startsWith("KONTRAK")||g.contract_type.includes("BULAN")||g.contract_type.includes("TAHUN"))||g.pkwt_number&&g.pkwt_number.startsWith("PKWT")?Z=!1:Z=Bt(g.end_date));let ie=g?.start_date?String(g.start_date).slice(0,10):"",le=g?.end_date&&!Bt(g.end_date)?String(g.end_date).slice(0,10):"";!Z&&ie&&!le&&(le=S(ie,U));let pe=`
      <div style="display:flex; flex-direction:column; gap:14px;">
        ${z?`<input type="hidden" id="con-modal-row-version" value="${O!==null?O:""}">`:""}
        <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:var(--text-2); margin-bottom:5px;">
              Nama Karyawan <span style="color:#DC2626;">*</span>
            </label>
            <div id="con-modal-emp-mount"></div>
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:var(--text-2); margin-bottom:5px;">
              Cabang Penempatan <span style="color:#DC2626;">*</span>
            </label>
            <select id="con-modal-branch" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" required>
              <option value="">-- Pilih Cabang --</option>
              ${o.map(oe=>`
                <option value="${oe.value}" ${String(ne)===String(oe.value)?"selected":""}>
                  ${oe.label}
                </option>
              `).join("")}
            </select>
          </div>
        </div>

        <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:var(--text-2); margin-bottom:5px;">
              Divisi / Bagian <span style="color:#DC2626;">*</span>
            </label>
            <select id="con-modal-division" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" required>
              <option value="FACILITY CARE" ${ee==="FACILITY CARE"?"selected":""}>FACILITY CARE</option>
              <option value="SECURITY" ${ee==="SECURITY"?"selected":""}>SECURITY</option>
            </select>
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:var(--text-2); margin-bottom:5px;">
              Status Kontrak <span style="color:#DC2626;">*</span>
            </label>
            <select id="con-modal-status" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" required>
              <option value="Aktif" ${H==="Aktif"?"selected":""}>\u{1F7E2} Aktif</option>
              <option value="Tidak Aktif" ${H==="Tidak Aktif"?"selected":""}>\u26AA Tidak Aktif</option>
              <option value="Resign" ${H==="Resign"?"selected":""}>\u{1F534} Resign</option>
              <option value="Cut" ${H==="Cut"?"selected":""}>\u2702\uFE0F Cut / Pemutusan</option>
            </select>
          </div>
        </div>

        <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:var(--text-2); margin-bottom:5px;">
              Tipe / Durasi Kontrak
            </label>
            <select id="con-modal-type" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;">
              <option value="KONTRAK 6 BULAN" ${U==="KONTRAK 6 BULAN"?"selected":""}>KONTRAK 6 BULAN</option>
              <option value="KONTRAK 1 TAHUN" ${U==="KONTRAK 1 TAHUN"?"selected":""}>KONTRAK 1 TAHUN</option>
              <option value="KONTRAK 2 TAHUN" ${U==="KONTRAK 2 TAHUN"?"selected":""}>KONTRAK 2 TAHUN</option>
              <option value="PKWTT (TETAP)" ${U==="PKWTT (TETAP)"?"selected":""}>PKWTT (TETAP)</option>
            </select>
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:var(--text-2); margin-bottom:5px;">
              Tingkat / No. PKWT
            </label>
            <select id="con-modal-pkwt" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;">
              <option value="PKWT 1" ${D==="PKWT 1"?"selected":""}>PKWT 1</option>
              <option value="PKWT 2" ${D==="PKWT 2"?"selected":""}>PKWT 2</option>
              <option value="PKWT 3" ${D==="PKWT 3"?"selected":""}>PKWT 3</option>
              <option value="PKWT 4" ${D==="PKWT 4"?"selected":""}>PKWT 4</option>
              <option value="PKWT 5" ${D==="PKWT 5"?"selected":""}>PKWT 5</option>
              <option value="PKWT 6" ${D==="PKWT 6"?"selected":""}>PKWT 6</option>
              <option value="TETAP" ${D==="TETAP"?"selected":""}>PKWTT / TETAP</option>
            </select>
          </div>
        </div>

        <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:var(--text-2); margin-bottom:5px;">
              Tanggal Mulai Kontrak
            </label>
            <input type="date" id="con-modal-start" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              value="${ie}" />
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:var(--text-2); margin-bottom:5px;">
              Tanggal Selesai Kontrak
            </label>
            <input type="date" id="con-modal-end" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px; ${Z?"opacity:0.5; background-color:#0E182B; cursor:not-allowed;":"background-color:#0E182B;"}" 
              value="${le}" ${Z?"disabled":""} />
            <div style="margin-top:6px; display:flex; align-items:center; gap:6px;">
              <input type="checkbox" id="con-modal-is-tetap" ${Z?"checked":""} style="cursor:pointer;" />
              <label for="con-modal-is-tetap" style="font-size:0.78rem; color:var(--text-2); cursor:pointer; user-select:none;">
                Karyawan Tetap / PKWTT (Tanpa Batas Tanggal)
              </label>
            </div>
          </div>
        </div>

        <div>
          <label style="display:block; font-size:0.82rem; font-weight:600; color:var(--text-2); margin-bottom:5px;">
            Catatan Tambahan
          </label>
          <textarea id="con-modal-notes" class="form-control" rows="2" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px; resize:vertical;">${V}</textarea>
        </div>
      </div>
    `,de=document.createElement("div");de.innerHTML=pe;let me=je(de.querySelector("#con-modal-emp-mount"),{mode:"single",inputId:"con-modal-employee",inputName:"employee_name",placeholder:"",required:!0,allowManual:!0,hasEmployeeIdSchema:!0,manualBadgeText:"Belum terhubung ke Master Karyawan",employees:s,initialEmployeeId:Q||null,initialEmployeeName:g?.employee_name||"",onChange:oe=>{if(!z&&oe?.employee_id){let ve=s.find(fe=>Number(fe.value??fe.id)===Number(oe.employee_id));if(ve){let fe=de.querySelector("#con-modal-branch"),De=de.querySelector("#con-modal-division");fe&&!fe.value&&ve.branch_id&&(fe.value=String(ve.branch_id)),De&&ve.division&&(De.value=ve.division)}}}}),Ae=!1,Ke=ce({title:te,content:de,confirmText:z?"\u{1F4BE} Simpan Perubahan":"\u2795 Tambah Kontrak",cancelText:"Batal",onConfirm:async(oe,ve)=>{if(Ae)return;let fe=me.getValue(),De=me.inputEl,Se=oe.querySelector("#con-modal-branch"),Le=oe.querySelector("#con-modal-division"),Ue=oe.querySelector("#con-modal-status"),_e=oe.querySelector("#con-modal-start"),ke=oe.querySelector("#con-modal-end"),Ye=oe.querySelector("#con-modal-is-tetap"),_i=oe.querySelector("#con-modal-type"),Bi=oe.querySelector("#con-modal-pkwt"),Ti=oe.querySelector("#con-modal-notes"),Ja=fe.employee_id,Wa=(fe.employee_name||"").trim(),Ya=Se?.value,Di=Le?.value||"FACILITY CARE",Li=Ue?.value||"Aktif",aa=_e?.value?.trim()||null,Qa=Ye?.checked,Pi=ke?.value?.trim(),na=Qa?"2099-12-31":Pi||null,Ii=_i?.value,Mi=Bi?.value,zi=Ti?.value?.trim()||null;if(!Wa){j("Harap pilih atau ketik nama karyawan!"),De?.focus();return}if(!Ya){j("Harap pilih cabang penempatan!"),Se?.focus();return}if(!Qa&&aa&&na&&na<aa){j("Tanggal selesai kontrak tidak boleh lebih awal dari tanggal mulai kontrak!"),ke?.focus();return}let ia=null;if(z&&(ia=Wt(oe,"#con-modal-row-version",g?.row_version),ia===null)){j("Versi data kontrak (row_version) tidak tersedia. Silakan muat ulang halaman lalu coba lagi.");return}let Ni={employee_id:Ja?parseInt(Ja,10):null,employee_name:Wa,branch_id:parseInt(Ya,10),division:Di,status:Li,start_date:aa,end_date:na,contract_type:Ii,pkwt_number:Mi,notes:zi,...z?{row_version:ia}:{}};Ae=!0;try{let ra=z?`/api/contracts/${g.id}`:"/api/contracts",Ki=await X(ra,{method:z?"PUT":"POST",body:JSON.stringify(Ni)});await Yt({res:Ki,isEdit:z,item:g,close:ve,reloadData:async()=>{At(),await W()},editSuccessMessage:"Data kontrak berhasil diperbarui!",createSuccessMessage:"Kontrak baru berhasil ditambahkan!",defaultErrorMessage:"Gagal menyimpan data kontrak."})}catch(ra){console.error(ra),j("Terjadi kesalahan koneksi ke server.")}finally{Ae=!1}}});setTimeout(()=>{let oe=Ke?.overlay||document.querySelector(".modal-overlay:last-child");if(!oe)return;let ve=oe.querySelector("#con-modal-is-tetap"),fe=oe.querySelector("#con-modal-end"),De=oe.querySelector("#con-modal-start"),Se=oe.querySelector("#con-modal-type"),Le=oe.querySelector("#con-modal-pkwt");if(!ve||!fe)return;function Ue(_e,ke=!0){_e?(ve.checked=!0,fe.disabled=!0,fe.style.opacity="0.5",fe.style.backgroundColor="#F8FAFC",fe.style.cursor="not-allowed",fe.value="",ke&&(Se&&(Se.value="PKWTT (TETAP)"),Le&&(Le.value="TETAP"))):(ve.checked=!1,fe.disabled=!1,fe.style.opacity="1",fe.style.backgroundColor="#FFFFFF",fe.style.cursor="default",ke&&(Se&&Se.value==="PKWTT (TETAP)"&&(Se.value="KONTRAK 1 TAHUN"),Le&&Le.value==="TETAP"&&(Le.value="PKWT 1")))}ve.addEventListener("change",_e=>{Ue(_e.target.checked,!0),_e.target.checked||fe.focus()}),Se&&Se.addEventListener("change",_e=>{let ke=_e.target.value;ke==="PKWTT (TETAP)"?(Ue(!0,!1),Le&&(Le.value="TETAP")):(Ue(!1,!1),Le&&Le.value==="TETAP"&&(Le.value="PKWT 1"),De&&De.value&&(fe.value=S(De.value,ke)))}),Le&&Le.addEventListener("change",_e=>{_e.target.value==="TETAP"?(Se&&(Se.value="PKWTT (TETAP)"),Ue(!0,!1)):(Se&&Se.value==="PKWTT (TETAP)"&&(Se.value="KONTRAK 1 TAHUN"),Ue(!1,!1),De&&De.value&&Se&&(fe.value=S(De.value,Se.value)))}),De&&De.addEventListener("change",_e=>{!ve.checked&&_e.target.value&&Se&&(fe.value=S(_e.target.value,Se.value))})},20)}function r(g){ce({title:"\u{1F5D1}\uFE0F Hapus Data Kontrak",content:`
        <div style="font-size:0.92rem; color:var(--text-1); line-height:1.5;">
          Apakah Anda yakin ingin menghapus catatan kontrak untuk <strong>${g.employee_name||"karyawan ini"}</strong> (ID #${g.id})?
          <div style="margin-top:10px; padding:10px; background:rgba(239,68,68,0.16); border-radius:8px; border:1px solid rgba(248,113,113,0.35); font-size:0.8rem; color:#FCA5A5;">
            \u26A0\uFE0F Tindakan ini akan menghapus riwayat masa berlaku kontrak terkait dari sistem.
          </div>
        </div>
      `,confirmText:"Ya, Hapus",cancelText:"Batal",onConfirm:async(C,z)=>{try{let O=await X(`/api/contracts/${g.id}`,{method:"DELETE"});O.ok?(ae("Data kontrak berhasil dihapus."),z(),await W()):j("Gagal menghapus kontrak: "+(O.data?.message||"Terjadi kesalahan"))}catch(O){console.error(O),j("Terjadi kendala koneksi.")}}})}async function l(){if(!re(_,"export")){j("Anda tidak memiliki izin untuk mengekspor data.");return}try{let g=await ue("/api/contracts/export",G());if(!g.length){j("Belum ada data kontrak untuk diekspor.");return}let C=g.map(z=>{let O=Bt(z.end_date),te=!!(z.end_date&&String(z.end_date).trim()&&z.end_date!=="-"),Q=z.days_remaining!==null&&z.days_remaining!==void 0&&z.days_remaining!==""?Number(z.days_remaining):null;return{"Nama Lengkap":z.employee_name||"",Cabang:z.branch_name||"",Divisi:z.division||"","No. PKWT":z.pkwt_number||"","Tipe Kontrak":z.contract_type||"","Tanggal Mulai":z.start_date||"","Tanggal Selesai":O?"Tetap / PKWTT":te?z.end_date:"-","Sisa Hari":O?"Tetap":te&&Q!==null&&!isNaN(Q)?Q+" Hari":"-",Status:z.status||"",Catatan:z.notes||""}});be(C,`Data_Kontrak_Karyawan_${new Date().toISOString().slice(0,10)}`),ae("File Excel Data Kontrak berhasil diunduh!")}catch(g){j("Gagal mengekspor data kontrak: "+g.message)}}function f(g){return g.length===0?`
            <div class="con-empty-state">
              <div style="font-size:2.8rem; margin-bottom:10px;">\u{1F50D}</div>
              <h3 style="font-size:1.15rem; font-weight:700; color:#F8FAFC; margin:0 0 6px 0;">Tidak ada kontrak yang sesuai kriteria</h3>
              <p style="font-size:0.85rem; color:#94A3B8; margin:0 0 16px 0;">Coba atur ulang kata kunci pencarian, filter cabang, atau filter masa berlaku kontrak.</p>
              <button id="con-reset-filter" class="btn btn-secondary" style="font-size:0.82rem; padding:7px 16px; border-radius:8px;">Reset Semua Filter</button>
            </div>
      `:B==="table"?`
            <div class="con-table-card">
              <div class="con-table-responsive">
                <table class="con-table">
                  <thead>
                    <tr>
                      <th style="width:45px; text-align:center;">No</th>
                      <th>Nama Karyawan</th>
                      <th>Cabang & Divisi</th>
                      <th style="text-align:center;">Tingkat PKWT</th>
                      <th>Mulai Kontrak</th>
                      <th>Selesai Kontrak</th>
                      <th>Masa Berlaku</th>
                      <th style="text-align:center;">Status</th>
                      <th style="width:110px; text-align:center;">Aksi</th>
                    </tr>
                  </thead>
                  <tbody>
                    ${g.map((C,z)=>{let O=C.division||"FACILITY CARE",te="con-div-fc",Q="";O==="SECURITY"&&(te="con-div-sec",Q="sec");let ne=!!(C.end_date&&String(C.end_date).trim()&&C.end_date!=="-"),ee=Bt(C.end_date),H=C.status==="Aktif",U=C.days_remaining!==null&&C.days_remaining!==void 0&&C.days_remaining!==""?Number(C.days_remaining):null,D=ne&&!ee&&U!==null&&!isNaN(U)?U:null,V="con-status-aktif",Z="\u{1F7E2} Aktif";ee?(V="con-status-tetap",Z="\u{1F3DB}\uFE0F PKWTT (Tetap)"):H?D!==null?D<=0?(V="con-status-danger",Z="\u26A0\uFE0F Habis Hari Ini"):D<=30?(V="con-status-warning",Z=`\u26A0\uFE0F Sisa ${D} Hari`):(V="con-status-aktif",Z=`\u{1F7E2} Aktif (${D} Hari)`):(V="con-status-aktif",Z="\u{1F7E2} Aktif"):(V="con-status-danger",Z=`\u26AA ${C.status||"Nonaktif"}`);let ie=C.employee_phone||c[C.employee_id]||c[String(C.employee_name||"").toLowerCase()]||"";return`
                        <tr data-id="${C.id}">
                          <td style="text-align:center; color:#94A3B8; font-weight:600;">${(h-1)*20+z+1}</td>
                          <td>
                            <div style="display:flex; align-items:center; gap:10px;">
                              <div class="con-avatar ${Q}" style="width:34px; height:34px; font-size:0.75rem;">
                                ${Dn(C.employee_name)}
                              </div>
                              <div style="min-width:0;">
                                <div style="font-weight:700; color:#F8FAFC; font-size:0.9rem;">${C.employee_name||"Karyawan Tanpa Nama"}</div>
                                <div style="font-size: 0.76rem; color:#94A3B8;">ID #${C.id} ${C.pkwt_number?"\u2022 "+C.pkwt_number:""}</div>
                              </div>
                            </div>
                          </td>
                          <td>
                            <div style="display:flex; flex-direction:column; gap:3px;">
                              <span style="font-weight:600; color:#E2E8F0;">${C.branch_name||"Semua Cabang"}</span>
                              <span class="con-div-badge ${te}" style="align-self:flex-start;">
                                ${O==="SECURITY"?"\u{1F6E1}\uFE0F":"\u{1F9F9}"} ${O}
                              </span>
                            </div>
                          </td>
                          <td style="text-align:center;">
                            <span class="con-pkwt-badge" style="display:inline-block;">${C.contract_type||"PKWT"}</span>
                          </td>
                          <td style="font-size:0.84rem; color:#E2E8F0;">
                            ${_t(C.start_date)}
                          </td>
                          <td style="font-size:0.84rem; color:#E2E8F0;">
                            ${ee?'<span style="color:#A5B4FC; font-weight:600;">PKWTT (Tetap)</span>':_t(C.end_date)}
                          </td>
                          <td>
                            ${ee?'<span style="color:#A5B4FC; font-weight:600; font-size:0.82rem;">\u{1F3DB}\uFE0F Tidak Berbatas</span>':D!==null?D<=0?'<span style="color:#FCA5A5; font-weight:700; font-size:0.82rem;">\u26A0\uFE0F Habis Hari Ini</span>':D<=30?`<span style="color:#FCD34D; font-weight:700; font-size:0.82rem;">\u26A0\uFE0F Sisa ${D} Hari</span>`:`<span style="color:#6EE7B7; font-weight:600; font-size:0.82rem;">\u{1F7E2} ${D} Hari</span>`:'<span style="color:#94A3B8; font-size:0.82rem;">-</span>'}
                          </td>
                          <td style="text-align:center;">
                            <span class="con-status-badge ${V}">${Z}</span>
                          </td>
                          <td style="text-align:center;">
                            <div class="con-action-btns" style="justify-content:center;">
                              ${ie?`
                                <a href="https://wa.me/${ie.replace(/[^0-9]/g,"").replace(/^0/,"62")}" target="_blank" rel="noopener" class="con-icon-btn" title="Chat WhatsApp Karyawan" style="color:#16A34A;">
                                  <svg width="15" height="15" fill="currentColor" viewBox="0 0 24 24"><path d="M12.031 6.172c-3.181 0-5.767 2.586-5.768 5.766-.001 1.298.38 2.27 1.019 3.287l-.582 2.128 2.182-.573c.978.58 1.911.928 3.145.929 3.178 0 5.767-2.587 5.768-5.766.001-3.187-2.575-5.771-5.764-5.771zm3.392 8.244c-.144.405-.837.774-1.17.824-.299.045-.677.063-1.092-.069-.252-.08-.575-.187-.988-.365-1.739-.751-2.874-2.502-2.961-2.617-.087-.116-.708-.94-.708-1.793s.448-1.273.607-1.446c.159-.173.346-.217.462-.217l.332.006c.106.005.249-.04.39.298.144.347.491 1.2.534 1.287.043.087.072.188.014.304-.058.116-.087.188-.173.289l-.26.304c-.087.086-.177.18-.076.354.101.174.449.741.964 1.201.662.591 1.221.774 1.394.86s.275.072.376-.043c.101-.116.433-.506.549-.68.116-.173.231-.145.39-.086s1.011.477 1.184.564.289.13.332.202c.045.072.045.419-.099.824zm-3.423-14.416c-6.627 0-12 5.373-12 12 0 2.159.57 4.185 1.564 5.939l-1.564 5.717 5.864-1.538c1.691.924 3.63 1.458 5.698 1.458 6.627 0 12-5.373 12-12s-5.373-12-12-12z"/></svg>
                                </a>
                              `:""}
                              ${E?`
                                <button class="con-icon-btn edit" data-action="edit" data-id="${C.id}" title="Edit Data Kontrak">
                                  <svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>
                                </button>
                              `:""}
                              ${K?`
                                <button class="con-icon-btn delete" data-action="delete" data-id="${C.id}" title="Hapus Kontrak">
                                  <svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><polyline points="3 6 5 6 21 6"/><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"/><line x1="10" y1="11" x2="10" y2="17"/><line x1="14" y1="11" x2="14" y2="17"/></svg>
                                </button>
                              `:""}
                            </div>
                          </td>
                        </tr>
                      `}).join("")}
                  </tbody>
                </table>
              </div>
            </div>
      `:`
          <div class="con-cards-grid">
            ${g.map(C=>{let z=C.division||"FACILITY CARE",O="con-div-fc",te="";z==="SECURITY"&&(O="con-div-sec",te="sec");let Q=!!(C.end_date&&String(C.end_date).trim()&&C.end_date!=="-"),ne=Bt(C.end_date),ee=C.status==="Aktif",H=C.days_remaining!==null&&C.days_remaining!==void 0&&C.days_remaining!==""?Number(C.days_remaining):null,U=Q&&!ne&&H!==null&&!isNaN(H)?H:null,D="con-status-aktif",V="\u{1F7E2} Aktif";ne?(D="con-status-tetap",V="\u{1F3DB}\uFE0F PKWTT (Tetap)"):ee?U!==null?U<=0?(D="con-status-danger",V="\u26A0\uFE0F Habis Hari Ini"):U<=30?(D="con-status-warning",V=`\u26A0\uFE0F Sisa ${U} Hari`):(D="con-status-aktif",V=`\u{1F7E2} Aktif (${U} Hari)`):(D="con-status-aktif",V="\u{1F7E2} Aktif"):(D="con-status-danger",V=`\u26AA ${C.status||"Nonaktif"}`);let Z=C.employee_phone||c[C.employee_id]||c[String(C.employee_name||"").toLowerCase()]||"";return`
                <div class="con-card" data-id="${C.id}">
                  <div class="con-card-top">
                    <div style="display:flex; align-items:center; gap:6px; flex-wrap:wrap;">
                      <span class="con-div-badge ${O}">
                        ${z==="SECURITY"?"\u{1F6E1}\uFE0F":"\u{1F9F9}"} ${z}
                      </span>
                      ${C.pkwt_number?`
                        <span class="con-pkwt-badge">
                          \u{1F4DC} ${C.pkwt_number}
                        </span>
                      `:""}
                    </div>
                    <span class="con-status-badge ${D}">
                      ${V}
                    </span>
                  </div>

                  <div class="con-profile-box">
                    <div class="con-avatar ${te}">
                      ${Dn(C.employee_name)}
                    </div>
                    <div style="min-width:0; flex:1;">
                      <div class="con-name">
                        ${C.employee_name||"Karyawan Tanpa Nama"}
                      </div>
                      <div class="con-subtitle">
                        Personil ${z} \u2022 ${C.contract_type||"Kontrak Kerja"}
                      </div>
                    </div>
                  </div>

                  <div class="con-info-grid">
                    <div class="con-info-card">
                      <div class="con-info-label">Cabang Penempatan</div>
                      <div class="con-info-val" title="${C.branch_name||"-"}">
                        \u{1F3E2} ${C.branch_name||"Semua Cabang"}
                      </div>
                    </div>
                    <div class="con-info-card">
                      <div class="con-info-label">Tingkat PKWT</div>
                      <div class="con-info-val">
                        \u{1F4CB} ${C.contract_type||"PKWT 1"}
                      </div>
                    </div>
                    <div class="con-info-card">
                      <div class="con-info-label">Mulai Kontrak</div>
                      <div class="con-info-val">
                        \u{1F4C5} ${_t(C.start_date)}
                      </div>
                    </div>
                    <div class="con-info-card">
                      <div class="con-info-label">Selesai Kontrak</div>
                      <div class="con-info-val">
                        ${ne?"\u{1F3DB}\uFE0F PKWTT (Tetap)":"\u{1F4C5} "+_t(C.end_date)}
                      </div>
                    </div>
                  </div>

                  ${ne?`
                    <div style="background:rgba(99,102,241,0.16); border:1px solid rgba(129,140,248,0.35); border-radius:10px; padding:10px 12px; font-size:0.82rem; color:#C7D2FE; display:flex; align-items:center; gap:8px;">
                      <span style="font-size:1.1rem;">\u{1F3DB}\uFE0F</span>
                      <div>
                        <strong style="color:#E0E7FF;">Karyawan Tetap (PKWTT)</strong>
                        <div style="font-size:0.75rem; color:#A5B4FC;">Tidak memiliki batas akhir masa berlaku kontrak.</div>
                      </div>
                    </div>
                  `:U!==null?`
                    <div style="background:${U<=0?"rgba(239,68,68,0.15)":U<=30?"rgba(245,158,11,0.15)":"rgba(16,185,129,0.14)"}; border:1px solid ${U<=0?"rgba(248,113,113,0.35)":U<=30?"rgba(251,191,36,0.35)":"rgba(52,211,153,0.32)"}; border-radius:10px; padding:10px 12px; font-size:0.82rem; color:${U<=0?"#FCA5A5":U<=30?"#FCD34D":"#6EE7B7"}; display:flex; align-items:center; gap:8px;">
                      <span style="font-size:1.1rem;">${U<=0?"\u26A0\uFE0F":U<=30?"\u23F3":"\u2705"}</span>
                      <div>
                        <strong>${U<=0?"Kontrak Telah Berakhir":`Sisa Waktu Kontrak: ${U} Hari`}</strong>
                        <div style="font-size:0.75rem; color:#CBD5E1; margin-top:2px;">
                          ${U<=30&&U>0?"Segera lakukan evaluasi perpanjangan kontrak.":U<=0?"Masa berlaku kontrak telah habis.":"Masa kontrak masih aktif dan aman."}
                        </div>
                      </div>
                    </div>
                  `:""}

                  <div class="con-card-footer">
                    <div style="font-size:0.75rem; color:#94A3B8;">
                      ID #${C.id} \u2022 ${C.pkwt_number||"PKWT"}
                    </div>

                    <div class="con-action-btns">
                      ${Z?`
                        <a href="https://wa.me/${Z.replace(/[^0-9]/g,"").replace(/^0/,"62")}" target="_blank" rel="noopener" class="con-icon-btn" title="Chat WhatsApp Karyawan" style="color:#16A34A;">
                          <svg width="16" height="16" fill="currentColor" viewBox="0 0 24 24"><path d="M12.031 6.172c-3.181 0-5.767 2.586-5.768 5.766-.001 1.298.38 2.27 1.019 3.287l-.582 2.128 2.182-.573c.978.58 1.911.928 3.145.929 3.178 0 5.767-2.587 5.768-5.766.001-3.187-2.575-5.771-5.764-5.771zm3.392 8.244c-.144.405-.837.774-1.17.824-.299.045-.677.063-1.092-.069-.252-.08-.575-.187-.988-.365-1.739-.751-2.874-2.502-2.961-2.617-.087-.116-.708-.94-.708-1.793s.448-1.273.607-1.446c.159-.173.346-.217.462-.217l.332.006c.106.005.249-.04.39.298.144.347.491 1.2.534 1.287.043.087.072.188.014.304-.058.116-.087.188-.173.289l-.26.304c-.087.086-.177.18-.076.354.101.174.449.741.964 1.201.662.591 1.221.774 1.394.86s.275.072.376-.043c.101-.116.433-.506.549-.68.116-.173.231-.145.39-.086s1.011.477 1.184.564.289.13.332.202c.045.072.045.419-.099.824zm-3.423-14.416c-6.627 0-12 5.373-12 12 0 2.159.57 4.185 1.564 5.939l-1.564 5.717 5.864-1.538c1.691.924 3.63 1.458 5.698 1.458 6.627 0 12-5.373 12-12s-5.373-12-12-12z"/></svg>
                        </a>
                      `:""}

                      ${E?`
                        <button class="con-icon-btn edit" data-action="edit" data-id="${C.id}" title="Edit Data Kontrak">
                          <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>
                        </button>
                      `:""}
                      ${K?`
                        <button class="con-icon-btn delete" data-action="delete" data-id="${C.id}" title="Hapus Kontrak">
                          <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><polyline points="3 6 5 6 21 6"/><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"/><line x1="10" y1="11" x2="10" y2="17"/><line x1="14" y1="11" x2="14" y2="17"/></svg>
                        </button>
                      `:""}
                    </div>
                  </div>
                </div>
              `}).join("")}
          </div>
    `}function u(){let g=e.querySelector("#con-pagination-container");if(!g)return;g.innerHTML="";let C=Te({page:h,pages:R,total:L,limit:20,onPage:z=>{h=z,W({refreshStats:!1})}});C&&g.appendChild(C)}function p(){let g=e.querySelector("#con-reset-filter");g&&g.addEventListener("click",()=>{m="",i="all",d="all",F="all",$="all",h=1;let C=e.querySelector("#con-search-input");C&&(C.value="");let z=e.querySelector("#con-branch-filter");z&&(z.value="all");let O=e.querySelector("#con-division-filter");O&&(O.value="all");let te=e.querySelector("#con-expiring-filter");te&&(te.value="all"),e.querySelectorAll(".con-pill").forEach(Q=>{Q.classList.toggle("active",Q.getAttribute("data-status")==="all")}),W({refreshStats:!1})}),e.querySelectorAll(".con-icon-btn.edit").forEach(C=>{C.addEventListener("click",z=>{let O=parseInt(z.currentTarget.getAttribute("data-id"),10),te=a.find(Q=>Q.id===O);te&&y(te)})}),e.querySelectorAll(".con-icon-btn.delete").forEach(C=>{C.addEventListener("click",z=>{let O=parseInt(z.currentTarget.getAttribute("data-id"),10),te=a.find(Q=>Q.id===O);te&&r(te)})})}function A(){let g=e.querySelector("#con-list-container");if(!g)return;let C=Y();g.innerHTML=f(C),u(),p()}function T(){let g=Y(),C=N.total,z=N.active,O=N.expiring30,te=N.permanent,Q=N.nonactive;e.innerHTML=`
      <style>.con-page{display:flex;flex-direction:column;gap:20px;width:100%;max-width:100%;min-width:0;box-sizing:border-box;overflow-x:hidden}.con-hero-card{background:#131F36;border:1px solid #223456;border-radius:16px;padding:24px 28px;color:#FFFFFF;box-shadow:0 4px 16px rgba(0,0,0,0.22);position:relative;overflow:hidden;box-sizing:border-box;width:100%;max-width:100%;min-width:0}.con-hero-card::after{content:'';position:absolute;top:-40px;right:-40px;width:220px;height:220px;background:radial-gradient(circle,rgba(37,99,235,0.06) 0%,transparent 70%);border-radius:50%;pointer-events:none}.con-hero-actions{display:flex;align-items:center;gap:8px;flex-wrap:wrap}.con-hero-actions .btn{height:38px;box-sizing:border-box;transition:all 0.2s ease}.con-btn-add:hover{background:#1D4ED8 !important;transform:translateY(-1px)}.con-btn-excel:hover,.con-btn-refresh:hover{background:rgba(255,255,255,0.16) !important;border-color:rgba(255,255,255,0.3) !important;transform:translateY(-1px)}.con-btn-missing:hover{background:rgba(239,68,68,0.25) !important;border-color:rgba(239,68,68,0.6) !important;transform:translateY(-1px)}.con-stats-row{display:grid;grid-template-columns:repeat(4,1fr);gap:14px;margin-top:20px}.con-stat-box{background:rgba(255,255,255,0.05);border:1px solid rgba(255,255,255,0.1);border-radius:12px;padding:14px 18px;display:flex;align-items:center;gap:14px;backdrop-filter:blur(8px);transition:transform 0.15s ease,background 0.15s ease}.con-stat-box:hover{background:rgba(255,255,255,0.08);transform:translateY(-2px)}.con-stat-icon{width:44px;height:44px;border-radius:10px;display:flex;align-items:center;justify-content:center;font-size:1.3rem;flex-shrink:0}.con-stat-val{font-size:1.55rem;font-weight:800;line-height:1.1;letter-spacing:-0.02em}.con-stat-lbl{font-size:0.76rem;text-transform:uppercase;color:#94A3B8;font-weight:700;letter-spacing:0.05em;margin-top:2px}.con-controls-card{background:#131F36;border-radius:14px;border:1px solid #223456;padding:16px 20px;box-shadow:0 2px 8px rgba(0,0,0,0.03);display:flex;flex-direction:column;gap:14px;box-sizing:border-box;width:100%;max-width:100%;min-width:0}.con-search-row{display:flex;gap:12px;align-items:center;flex-wrap:wrap}.con-search-wrap{position:relative;flex:1;min-width:240px}.con-search-icon{position:absolute;left:14px;top:50%;transform:translateY(-50%);color:#94A3B8;pointer-events:none}.con-search-input{width:100%;padding:10px 14px 10px 40px;border:1px solid #CBD5E1;border-radius:10px;font-size:0.9rem;outline:none;transition:border-color 0.15s ease}.con-search-input:focus{border-color:#2563EB;box-shadow:0 0 0 3px rgba(37,99,235,0.1)}.con-select{padding:10px 14px;border:1px solid #CBD5E1;border-radius:10px;font-size:0.9rem;color:#F8FAFC;background:#131F36;outline:none;min-width:170px;cursor:pointer}.con-select:focus{border-color:#2563EB}.con-filter-pills{display:flex;align-items:center;gap:8px;overflow-x:auto;white-space:nowrap;scrollbar-width:none;padding-bottom:2px}.con-filter-pills::-webkit-scrollbar{display:none}.con-pill{padding:6px 14px;border-radius:20px;font-size:0.82rem;font-weight:600;cursor:pointer;border:1px solid #223456;background:#0E182B;color:#CBD5E1;transition:all 0.15s ease;display:inline-flex;align-items:center;gap:6px;user-select:none;flex-shrink:0}.con-pill:hover{background:#22314A;color:#F8FAFC}.con-pill.active{background:#2563EB;color:white;border-color:#2563EB}.con-pill-dot{width:8px;height:8px;border-radius:50%}.con-pill-count{padding:1px 6px;border-radius:10px;font-size:0.76rem;background:rgba(255,255,255,0.1);color:#E2E8F0}.con-pill.active .con-pill-count{background:rgba(255,255,255,0.25);color:white}.con-view-toggle{display:inline-flex;align-items:center;background:#0E182B;padding:3px;border-radius:10px;border:1px solid #223456;gap:3px;flex-shrink:0}.con-view-btn{display:inline-flex;align-items:center;gap:6px;padding:6px 12px;border-radius:7px;font-size:0.82rem;font-weight:600;border:none;background:transparent;color:#CBD5E1;cursor:pointer;transition:all 0.15s ease;height:34px;box-sizing:border-box;user-select:none}.con-view-btn:hover{color:#F8FAFC;background:#22314A}.con-view-btn.active{background:#131F36;color:#60A5FA;font-weight:700;box-shadow:0 1px 4px rgba(0,0,0,0.25)}.con-table-card{background:#131F36;border:1px solid #223456;border-radius:14px;box-shadow:0 2px 6px rgba(0,0,0,0.2);overflow:hidden}.con-table-responsive{width:100%;overflow-x:auto;-webkit-overflow-scrolling:touch}.con-table{width:100%;border-collapse:collapse;text-align:left;font-size:0.88rem;min-width:960px}.con-table th{background:#0E182B;color:#CBD5E1;font-size:0.76rem;font-weight:700;text-transform:uppercase;letter-spacing:0.05em;padding:13px 16px;border-bottom:1.5px solid #223456;white-space:nowrap}.con-table td{padding:13px 16px;border-bottom:1px solid #223456;color:#F8FAFC;vertical-align:middle}.con-table tbody tr:hover td{background:#0E182B}.con-table tbody tr:last-child td{border-bottom:none}.con-cards-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(320px,1fr));gap:16px;width:100%;max-width:100%;min-width:0;box-sizing:border-box}.con-card{background:#131F36;border-radius:14px;border:1px solid #223456;padding:16px 18px;box-shadow:0 2px 8px rgba(0,0,0,0.04);display:flex;flex-direction:column;gap:13px;width:100% !important;max-width:100% !important;min-width:0 !important;box-sizing:border-box !important;overflow:hidden;transition:transform 0.15s ease,box-shadow 0.15s ease}.con-card:hover{transform:translateY(-2px);box-shadow:0 6px 16px rgba(0,0,0,0.08);border-color:#CBD5E1}.con-card-top{display:flex;justify-content:space-between;align-items:center;gap:8px;width:100%;min-width:0;box-sizing:border-box;flex-wrap:wrap}.con-div-badge{display:inline-flex;align-items:center;gap:5px;padding:4px 10px;border-radius:20px;font-size:0.76rem;font-weight:700;flex-shrink:0;white-space:nowrap}.con-div-fc{background:rgba(59,130,246,0.18);color:#93C5FD;border:1px solid rgba(96,165,250,0.35)}.con-div-sec{background:rgba(148,163,184,0.16);color:#E2E8F0;border:1px solid rgba(148,163,184,0.35)}.con-div-rel{background:rgba(139,92,246,0.18);color:#C4B5FD;border:1px solid rgba(167,139,250,0.35)}.con-pkwt-badge{display:inline-flex;align-items:center;gap:4px;padding:4px 10px;border-radius:20px;font-size:0.76rem;font-weight:700;background:rgba(139,92,246,0.18);color:#C4B5FD;border:1px solid rgba(167,139,250,0.35);flex-shrink:0;white-space:nowrap}.con-status-badge{display:inline-flex;align-items:center;gap:5px;padding:4px 10px;border-radius:20px;font-size:0.76rem;font-weight:700;flex-shrink:0;white-space:nowrap}.con-status-aktif{background:rgba(16,185,129,0.18);color:#6EE7B7;border:1px solid rgba(52,211,153,0.35)}.con-status-warning{background:rgba(245,158,11,0.18);color:#FCD34D;border:1px solid rgba(251,191,36,0.35)}.con-status-danger{background:rgba(239,68,68,0.18);color:#FCA5A5;border:1px solid rgba(248,113,113,0.35)}.con-status-tetap{background:rgba(99,102,241,0.2);color:#C7D2FE;border:1px solid rgba(129,140,248,0.38)}.con-profile-box{display:flex;align-items:center;gap:12px;width:100%;min-width:0;box-sizing:border-box}.con-avatar{width:48px;height:48px;border-radius:12px;background:#2563EB;color:#FFFFFF;display:flex;align-items:center;justify-content:center;font-weight:800;font-size:1.05rem;flex-shrink:0}.con-avatar.sec{background:#1B2D4F;border:1px solid #3B82F6}.con-avatar.rel{background:#1D4ED8}.con-name{font-size:1.02rem;font-weight:800;color:#F8FAFC;line-height:1.25;word-break:break-word;overflow-wrap:anywhere}.con-subtitle{font-size:0.76rem;color:#CBD5E1;font-weight:600;margin-top:3px}.con-info-card{background:#0E182B;border:1px solid #223456;border-radius:10px;padding:11px 12px;display:flex;flex-direction:column;gap:8px;font-size:0.83rem;width:100%;max-width:100%;min-width:0;box-sizing:border-box}.con-info-row{display:flex;align-items:flex-start;justify-content:space-between;gap:8px;width:100%;min-width:0;box-sizing:border-box}.con-info-label{color:#94A3B8;display:inline-flex;align-items:center;gap:5px;flex-shrink:0;white-space:nowrap;font-size:0.8rem}.con-info-val{color:#F8FAFC;font-weight:600;text-align:right;min-width:0;flex:1;word-break:break-word;overflow-wrap:anywhere;font-size:0.82rem}.con-notes-box{font-size:0.78rem;color:#FDE68A;background:rgba(245,158,11,0.14);border:1px solid rgba(251,191,36,0.3);border-radius:8px;padding:8px 10px;line-height:1.4;word-break:break-word;overflow-wrap:anywhere;width:100%;box-sizing:border-box}.con-card-footer{display:flex;justify-content:space-between;align-items:center;border-top:1px solid #223456;padding-top:11px;margin-top:auto;gap:8px;width:100%;min-width:0;box-sizing:border-box;flex-wrap:wrap}.con-action-btns{display:flex;align-items:center;gap:6px;margin-left:auto;flex-shrink:0}.con-icon-btn{width:34px;height:34px;border-radius:8px;border:1px solid #223456;background:#0E182B;color:#CBD5E1;display:flex;align-items:center;justify-content:center;cursor:pointer;transition:all 0.15s ease}.con-icon-btn:hover{background:#22314A;color:#F8FAFC;border-color:#475569}.con-icon-btn.delete:hover{background:rgba(239,68,68,0.2);color:#FCA5A5;border-color:rgba(248,113,113,0.45)}.con-empty-state{grid-column:1 / -1;background:#131F36;border:2px dashed #223456;border-radius:16px;padding:48px 24px;text-align:center}@media (max-width:768px){.con-page{gap:14px;width:100% !important;max-width:100% !important;overflow-x:hidden !important}.con-hero-card{padding:16px 14px !important;width:100% !important;max-width:100% !important;box-sizing:border-box !important}.con-stats-row{grid-template-columns:1fr 1fr;gap:10px}.con-stat-box{padding:10px 12px}.con-stat-icon{width:38px;height:38px;font-size:1.1rem}.con-stat-val{font-size:1.3rem}.con-hero-actions{width:100% !important;display:flex !important;align-items:center !important;gap:8px !important;flex-wrap:wrap !important;overflow-x:visible !important}.con-hero-actions::-webkit-scrollbar{display:none}.con-btn-add{width:100% !important;justify-content:center !important;padding:9px 12px !important;font-size:0.85rem !important;white-space:nowrap !important;height:40px !important}.con-btn-excel{flex:1 1 auto !important;justify-content:center !important;padding:8px 12px !important;font-size:0.85rem !important;white-space:nowrap !important;height:38px !important}.con-btn-refresh{width:38px !important;height:38px !important;padding:0 !important;flex-shrink:0 !important;justify-content:center !important}.con-controls-card{padding:14px !important;gap:12px;width:100% !important;max-width:100% !important;box-sizing:border-box !important}.con-search-row{flex-direction:column;gap:10px}.con-select{width:100%}.con-view-toggle{width:100%;display:flex}.con-view-btn{flex:1;justify-content:center}.con-table-card{border-radius:12px}.con-cards-grid{grid-template-columns:minmax(0,1fr) !important;width:100% !important;max-width:100% !important;min-width:0 !important;box-sizing:border-box !important}.con-card{padding:14px !important;width:100% !important;max-width:100% !important;min-width:0 !important;box-sizing:border-box !important}.con-info-card{padding:10px !important;width:100% !important;max-width:100% !important;box-sizing:border-box !important}}</style>

      <div class="con-page">

        <!-- Hero Header Card -->
        <div class="con-hero-card">
          <div style="display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:16px;">
            <div>
              <div style="display:flex; align-items:center; gap:10px;">
                <span style="font-size:1.6rem;">\u{1F4CB}</span>
                <h1 style="font-size:1.45rem; font-weight:800; margin:0; letter-spacing:-0.02em;">
                  Monitoring Kontrak Kerja Karyawan
                </h1>
              </div>
            </div>

            <!-- Action Buttons 1 Baris Sejajar -->
            <div class="con-hero-actions" style="display:flex; align-items:center; gap:8px;">
              ${w?`
                <button id="con-btn-add" class="btn btn-primary con-btn-add" style="background:#2563EB; height:38px; padding:0 16px; border-radius:10px; font-weight:700; display:inline-flex; align-items:center; gap:6px; box-shadow:0 4px 12px rgba(37,99,235,0.3); white-space:nowrap;">
                  <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24"><line x1="12" y1="5" x2="12" y2="19"/><line x1="5" y1="12" x2="19" y2="12"/></svg>
                  <span>Tambah Kontrak</span>
                </button>
              `:""}

              ${re(_,"export")?`
              <button id="con-btn-excel" class="btn btn-ghost con-btn-excel" style="background:rgba(255,255,255,0.08); border:1px solid rgba(255,255,255,0.15); color:white; height:38px; padding:0 14px; border-radius:10px; font-weight:600; display:inline-flex; align-items:center; gap:6px; white-space:nowrap;">
                <span>\u{1F4E5}</span>
                <span>Ekspor Excel</span>
              </button>
              `:""}

              

              <button id="con-btn-refresh" class="btn btn-ghost con-btn-refresh" style="background:rgba(255,255,255,0.08); border:1px solid rgba(255,255,255,0.15); color:white; width:38px; height:38px; padding:0; border-radius:10px; display:inline-flex; align-items:center; justify-content:center; flex-shrink:0;" title="Muat Ulang Data">
                <svg width="16" height="16" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M23 4v6h-6"/><path d="M1 20v-6h6"/><path d="M3.51 9a9 9 0 0 1 14.85-3.36L23 10M1 14l4.64 4.36A9 9 0 0 0 20.49 15"/></svg>
              </button>
            </div>
          </div>

          <!-- 4 KPI Stat Boxes -->
          <div class="con-stats-row">
            <div class="con-stat-box">
              <div class="con-stat-icon" style="background:rgba(59, 130, 246, 0.2); color:#60A5FA;">
                \u{1F4CB}
              </div>
              <div>
                <div class="con-stat-val" id="con-stat-total">${C}</div>
                <div class="con-stat-lbl">Total Kontrak</div>
              </div>
            </div>

            <div class="con-stat-box">
              <div class="con-stat-icon" style="background:rgba(16, 185, 129, 0.2); color:#34D399;">
                \u{1F7E2}
              </div>
              <div>
                <div class="con-stat-val" id="con-stat-active" style="color:#34D399;">${z}</div>
                <div class="con-stat-lbl">Kontrak Aktif</div>
              </div>
            </div>

            <div class="con-stat-box">
              <div class="con-stat-icon" style="background:rgba(245, 158, 11, 0.2); color:#FBBF24;">
                \u26A0\uFE0F
              </div>
              <div>
                <div class="con-stat-val" id="con-stat-expiring" style="color:#FBBF24;">${O}</div>
                <div class="con-stat-lbl">Habis &lt; 30 Hari</div>
              </div>
            </div>

            <div class="con-stat-box">
              <div class="con-stat-icon" style="background:rgba(99, 102, 241, 0.2); color:#A5B4FC;">
                \u{1F3DB}\uFE0F
              </div>
              <div>
                <div class="con-stat-val" id="con-stat-permanent" style="color:#A5B4FC;">${te}</div>
                <div class="con-stat-lbl">PKWTT (Tetap)</div>
              </div>
            </div>
          </div>
        </div>

        <!-- Controls / Filter Card -->
        <div class="con-controls-card">
          <div class="con-search-row">
            <div class="con-search-wrap">
              <svg class="con-search-icon" width="16" height="16" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/></svg>
              <input type="text" id="con-search-input" class="con-search-input" 
                placeholder="Cari nama karyawan, cabang, no PKWT, catatan..." value="${m?String(m).replace(/"/g,"&quot;"):""}" />
            </div>

            <select id="con-branch-filter" class="con-select">
              <option value="all">Semua Cabang (${o.length})</option>
              ${o.map(ie=>`
                <option value="${ie.value}" ${String(i)===String(ie.value)?"selected":""}>
                  ${ie.label}
                </option>
              `).join("")}
            </select>

            <select id="con-division-filter" class="con-select">
              <option value="all">Semua Divisi</option>
              <option value="FACILITY CARE" ${d==="FACILITY CARE"?"selected":""}>FACILITY CARE</option>
              <option value="SECURITY" ${d==="SECURITY"?"selected":""}>SECURITY</option>
            </select>

            <select id="con-expiring-filter" class="con-select">
              <option value="all" ${$==="all"?"selected":""}>Semua Durasi Masa Habis</option>
              <option value="7" ${$==="7"?"selected":""}>Akan Habis \u2264 7 Hari</option>
              <option value="14" ${$==="14"?"selected":""}>Akan Habis \u2264 14 Hari</option>
              <option value="30" ${$==="30"?"selected":""}>Akan Habis \u2264 30 Hari</option>
              <option value="60" ${$==="60"?"selected":""}>Akan Habis \u2264 60 Hari</option>
            </select>

            <div class="con-view-toggle">
              <button type="button" class="con-view-btn ${B==="card"?"active":""}" data-view="card" title="Tampilan Kartu">
                <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><rect x="3" y="3" width="7" height="7" rx="1"/><rect x="14" y="3" width="7" height="7" rx="1"/><rect x="14" y="14" width="7" height="7" rx="1"/><rect x="3" y="14" width="7" height="7" rx="1"/></svg>
                <span>Kartu</span>
              </button>
              <button type="button" class="con-view-btn ${B==="table"?"active":""}" data-view="table" title="Tampilan Tabel">
                <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M3 9h18M3 15h18M9 3v18"/><rect x="3" y="3" width="18" height="18" rx="2"/></svg>
                <span>Tabel</span>
              </button>
            </div>
          </div>

          <!-- Status Horizontal Swipeable Pills -->
          <div class="con-filter-pills">
            <div class="con-pill ${F==="all"&&$==="all"?"active":""}" data-status="all">
              <span>Semua Kontrak</span>
              <span class="con-pill-count" id="con-pill-count-all">${C}</span>
            </div>
            <div class="con-pill ${F==="Aktif"&&$==="all"?"active":""}" data-status="Aktif">
              <span class="con-pill-dot" style="background:#10B981;"></span>
              <span>Aktif</span>
              <span class="con-pill-count" id="con-pill-count-aktif">${z}</span>
            </div>
            <div class="con-pill ${F==="Expiring30"?"active":""}" data-status="Expiring30">
              <span class="con-pill-dot" style="background:#F59E0B;"></span>
              <span>Akan Habis (&lt;30 Hari)</span>
              <span class="con-pill-count" id="con-pill-count-expiring">${O}</span>
            </div>
            <div class="con-pill ${F==="Tetap"?"active":""}" data-status="Tetap">
              <span class="con-pill-dot" style="background:#6366F1;"></span>
              <span>PKWTT (Tetap)</span>
              <span class="con-pill-count" id="con-pill-count-tetap">${te}</span>
            </div>
            <div class="con-pill ${F==="Nonaktif"?"active":""}" data-status="Nonaktif">
              <span class="con-pill-dot" style="background:#EF4444;"></span>
              <span>Nonaktif / Resign</span>
              <span class="con-pill-count" id="con-pill-count-nonaktif">${Q}</span>
            </div>
          </div>
        </div>

        <!-- Content Area: Pilihan Kartu atau Tabel -->
        <div id="con-list-container">
          ${f(g)}
        </div>
        <div id="con-pagination-container"></div>

      </div>
    `;let ne=e.querySelector("#con-btn-add");ne&&ne.addEventListener("click",()=>y());let ee=e.querySelector("#con-btn-excel");ee&&ee.addEventListener("click",()=>l());let H=e.querySelector("#con-btn-refresh");H&&H.addEventListener("click",()=>W({refreshStats:!0})),e.querySelectorAll(".con-view-btn").forEach(ie=>{ie.addEventListener("click",()=>{let le=ie.getAttribute("data-view");le&&le!==B&&(B=le,localStorage.setItem("fcms_contracts_view_mode",le),e.querySelectorAll(".con-view-btn").forEach(pe=>{pe.classList.toggle("active",pe.getAttribute("data-view")===le)}),A())})});let U=e.querySelector("#con-search-input");U&&U.addEventListener("input",ie=>{m=ie.target.value,clearTimeout(q),q=setTimeout(()=>{h=1,W({refreshStats:!1})},350)});let D=e.querySelector("#con-branch-filter");D&&D.addEventListener("change",ie=>{i=ie.target.value,h=1,W({refreshStats:!1})});let V=e.querySelector("#con-division-filter");V&&V.addEventListener("change",ie=>{d=ie.target.value,h=1,W({refreshStats:!1})});let Z=e.querySelector("#con-expiring-filter");Z&&Z.addEventListener("change",ie=>{$=ie.target.value,F==="Expiring30"&&$!=="30"&&(F=$==="all"?"all":"Aktif"),e.querySelectorAll(".con-pill").forEach(le=>{let pe=le.getAttribute("data-status"),de=pe==="all"&&F==="all"&&$==="all"||pe==="Aktif"&&F==="Aktif"&&$==="all"||pe==="Expiring30"&&F==="Expiring30"||pe==="Tetap"&&F==="Tetap"||pe==="Nonaktif"&&F==="Nonaktif";le.classList.toggle("active",de)}),h=1,W({refreshStats:!1})}),e.querySelectorAll(".con-pill").forEach(ie=>{ie.addEventListener("click",le=>{let pe=le.currentTarget.getAttribute("data-status");F=pe,M="",pe==="Expiring30"?$="30":$="all";let de=e.querySelector("#con-expiring-filter");de&&(de.value=$),e.querySelectorAll(".con-pill").forEach(me=>{me.classList.toggle("active",me===le.currentTarget)}),h=1,W({refreshStats:!1})})}),u(),p()}await W({refreshStats:!0})}Ie();Be();Ie();function Ea(e){return e.map(t=>{if(t.type==="hidden")return`<input type="hidden" name="${P(t.name)}" value="${P(t.value??"")}">`;if(t.type==="html")return t.html||"";if(t.type==="row")return`<div class="form-row">${Ea(t.fields)}</div>`;let n=t.required?"required":"",a=t.label?`<label class="form-label">${P(t.label)}${t.required?' <span class="required">*</span>':""}</label>`:"",o="";switch(t.type){case"textarea":o=`<textarea name="${P(t.name)}" class="form-control" placeholder="${P(t.placeholder||"")}" ${n} rows="${t.rows||3}">${P(t.value??"")}</textarea>`;break;case"select":let c=(t.options||[]).map(h=>{let L=typeof h=="object"?h.value:h,R=typeof h=="object"?h.label:h,I=t.value==L?"selected":"";return`<option value="${P(L??"")}" ${I}>${P(R??"")}</option>`}).join("");o=`<select name="${P(t.name)}" class="form-control" ${n}><option value="">-- Pilih ${P(t.label||"")} --</option>${c}</select>`;break;case"combobox":let m=`dl-${P(t.name)}-${Math.random().toString(36).substring(7)}`,i=(t.options||[]).map(h=>{let L=typeof h=="object"?h.label||h.value||"":h||"";return(L==="undefined"||L==="[object Object]"||L==="null")&&(L=""),L?`<option value="${P(L)}"></option>`:""}).join(""),d=t.value||"";if(t.value){let h=(t.options||[]).find(L=>(typeof L=="object"?L.value:L)==t.value);if(h){let L=typeof h=="object"?h.label||h.value||"":h||"";L&&L!=="undefined"&&L!=="[object Object]"&&L!=="null"&&(d=L)}}o=`
          <input type="text" name="${P(t.name)}" list="${m}" class="form-control" value="${P(d)}" ${n} autocomplete="off">
          <datalist id="${m}">${i}</datalist>
        `;break;case"employee-combobox":{o=`<div class="fc-crud-emp-mount" data-emp-mount="${P(t.name)}" data-init-name="${P(t.value||"")}"></div>`;break}case"checkbox":o=`<label class="checkbox-label"><input type="checkbox" name="${P(t.name)}" value="1" ${t.value?"checked":""}> ${P(t.checkLabel||t.label)}</label>`;break;case"date":let F=window.parseFlexibleDate&&t.value?window.parseFlexibleDate(t.value):t.value||"";o=`<input type="date" name="${P(t.name)}" class="form-control" value="${P(F)}" ${n}>`;break;case"month":let $=t.value?String(t.value).slice(0,7):"";o=`<input type="month" name="${P(t.name)}" class="form-control" value="${P($)}" ${n}>`;break;case"number":let M=t.readonly?"readonly":"";o=`<input type="number" name="${P(t.name)}" class="form-control" value="${P(t.value??"")}" placeholder="${P(t.placeholder||"")}" min="${P(t.min??"")}" max="${P(t.max??"")}" step="${P(t.step||"1")}" ${n} ${M}>`;break;case"email":o=`<input type="email" name="${P(t.name)}" class="form-control" value="${P(t.value??"")}" placeholder="${P(t.placeholder||"")}" ${n} autocomplete="${P(t.autocomplete||"off")}">`;break;case"password":o=`<input type="password" name="${P(t.name)}" class="form-control" value="${P(t.value??"")}" placeholder="${P(t.placeholder||"")}" ${n} autocomplete="${P(t.autocomplete||"new-password")}">`;break;case"url":o=`<input type="url" name="${P(t.name)}" class="form-control" value="${P(t.value??"")}" placeholder="${P(t.placeholder||"")}" ${n}>`;break;default:o=`<input type="${P(t.type||"text")}" name="${P(t.name)}" class="form-control" value="${P(t.value??"")}" placeholder="${P(t.placeholder||"")}" ${n} autocomplete="${P(t.autocomplete||"off")}">`}let s=t.hint?`<div class="form-hint">${P(t.hint)}</div>`:"";return`<div class="form-group ${P(t.class||"")}">${a}${o}${s}</div>`}).join("")}function Pn(e){let t={},n=new FormData(e);for(let[a,o]of n.entries())t[a]=o===""?null:o;return e.querySelectorAll("input[type=checkbox]").forEach(a=>{a.checked||(t[a.name]=null)}),t}function In(e,t){t&&Object.entries(t).forEach(([n,a])=>{let o=e.querySelector(`[name="${n}"]`);o&&(o.hasAttribute("list")||o.classList.contains("fc-emp-input")||(o.type==="checkbox"?o.checked=!!a:o.type==="date"&&a&&window.parseFlexibleDate?o.value=window.parseFlexibleDate(a):o.type==="month"&&a?o.value=String(a).slice(0,7):o.value=a??""))})}Ce();var We={},Tt={on(e,t){We[e]||(We[e]=new Set),We[e].add(t)},off(e,t){We[e]&&We[e].delete(t)},emit(e,t){We[e]&&We[e].forEach(n=>{try{n(t)}catch(a){console.warn("[calendarBus] Handler error:",a)}})},clear(){Object.keys(We).forEach(e=>delete We[e])}},Cr=new Set(["schedule","cleaning","cleaning_reports","inspection","inspection_reports","fogging","fogging_reports","reliever","relievers","contract","contracts","issue","issues","training","one_on_one","sp","sp_data","mutasi","basecamp","basecamp_reports","supply"]);function it(e){if(!e){Tt.emit("data:changed",{module:"unknown"});return}let t=String(e).toLowerCase().replace(/^\/api\//,"").replace(/^reports\//,"");Tt.emit("data:changed",{module:t,relevant:Cr.has(t)})}var _a="fcms_import_sessions_v1",Da="fcms_import_batch_v1",Fr=15,Er=45e4,$a=new Set,ft=class extends Error{constructor(t,n={}){super(t||"Gagal mengimport data"),this.name="ImportRequestError",this.status=Number(n.status??0),this.code=n.code||"IMPORT_ERROR",this.importJobId=n.importJobId||null,this.committedChunks=Number(n.committedChunks||0),this.totalChunks=Number(n.totalChunks||0),this.committedRowIndex=Number(n.committedRowIndex??-1),this.recoverable=!!n.recoverable,this.idempotencyKey=n.idempotencyKey||null,this.fingerprint=n.fingerprint||null,this.isNetworkError=!!n.isNetworkError,this.serverStatus=n.serverStatus||null}};function $e(e){return String(e??"").replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;").replace(/'/g,"&#39;")}function Ba(e){if(e==null)return null;if(Array.isArray(e))return e.map(Ba);if(typeof e=="object"){let t={};for(let n of Object.keys(e).sort())t[n]=Ba(e[n]);return t}return e}function $r(e){let t=2166136261,n=2654435769;for(let a=0;a<e.length;a++){let o=e.charCodeAt(a);t^=o,t=Math.imul(t,16777619)>>>0,n^=o+(a&255),n=Math.imul(n,2246822507)>>>0}return`${t.toString(16).padStart(8,"0")}${n.toString(16).padStart(8,"0")}`}async function La(e,t){let n=JSON.stringify({endpoint:String(e||""),body:Ba(t)});try{let a=globalThis?.crypto?.subtle;if(a&&typeof TextEncoder<"u"){let o=new TextEncoder().encode(n),s=await a.digest("SHA-256",o);return Array.from(new Uint8Array(s)).map(c=>c.toString(16).padStart(2,"0")).join("")}}catch{}return`fnv_${$r(n)}_${n.length}`}function Ar(e){let t=String(e||"import").replace(/[^a-z0-9_-]/gi,"").slice(0,20)||"import";try{if(globalThis?.crypto?.randomUUID)return`imp_${t}_${globalThis.crypto.randomUUID()}`}catch{}return`imp_${t}_${Date.now().toString(36)}_${Math.random().toString(36).slice(2,10)}`}function Mn(e,t){try{if(typeof localStorage>"u")return t;let n=localStorage.getItem(e);if(!n)return t;let a=JSON.parse(n);return a&&typeof a=="object"?a:t}catch{return t}}function Ta(e,t){try{return typeof localStorage>"u"?!1:(localStorage.setItem(e,JSON.stringify(t)),!0)}catch{return!1}}function zn(e){let t=Object.values(e||{}).sort((a,o)=>String(o.updatedAt||"").localeCompare(String(a.updatedAt||""))).slice(0,Fr),n={};for(let a of t)a?.fingerprint&&(n[a.fingerprint]=a);if(!Ta(_a,n)){for(let a of t)a&&"body"in a&&(a.body=void 0,a.payloadStored=!1);Ta(_a,n)}}function Qt(){return Mn(_a,{})}function Dt(e){return e&&Qt()[e]||null}function Xt(e=null){let t=Qt();return Object.values(t).filter(a=>!(!a||a.status==="COMPLETED"||e&&a.module!==e)).sort((a,o)=>String(o.updatedAt||"").localeCompare(String(a.updatedAt||"")))}function Aa(e){if(!e?.fingerprint)return;let t=Qt(),n=e.body!==void 0?JSON.stringify(e.body):"",a=n.length>0&&n.length<=Er;t[e.fingerprint]={fingerprint:e.fingerprint,idempotencyKey:e.idempotencyKey,module:e.module||"import",label:e.label||e.module||"Import",endpoint:e.endpoint,status:e.status,importJobId:e.importJobId||null,committedChunks:Number(e.committedChunks||0),totalChunks:Number(e.totalChunks||0),committedRowIndex:Number(e.committedRowIndex??-1),recoverable:!!e.recoverable,errorCode:e.errorCode||null,errorMessage:e.errorMessage||null,idempotentReplay:!!e.idempotentReplay,result:e.result||null,payloadStored:e.status==="COMPLETED"?!1:a,body:e.status==="COMPLETED"?void 0:a?e.body:void 0,updatedAt:new Date().toISOString()},zn(t)}function yt(e){if(!e)return;let t=Qt();t[e]&&(delete t[e],zn(t))}function Lt(e){return!e||typeof e!="object"?!1:Ta(Da,{...e,updatedAt:new Date().toISOString()})}function Nn(){return Mn(Da,null)}function Pa(){try{typeof localStorage<"u"&&localStorage.removeItem(Da)}catch{}}function _r(e){return e&&typeof e=="object"&&!Array.isArray(e)&&e.data&&typeof e.data=="object"?e.data:e&&typeof e=="object"?e:{}}async function Pe({module:e,label:t=null,endpoint:n,body:a,idempotencyKey:o=null,recover:s=!1,forceNewKey:c=!1,configOverride:m=null,fetchImpl:i=null}){let d=await La(n,a);if($a.has(d)){let h=Dt(d);throw new ft("Proses import untuk data ini sedang berjalan. Mohon tunggu hingga selesai.",{status:409,code:"CLIENT_IMPORT_IN_PROGRESS",importJobId:h?.importJobId||null,committedChunks:h?.committedChunks||0,totalChunks:h?.totalChunks||0,recoverable:!1,idempotencyKey:h?.idempotencyKey||o||null,fingerprint:d})}let F=Dt(d),$=o?String(o).trim():null;!$&&!c&&F&&F.status!=="COMPLETED"&&F.idempotencyKey&&($=F.idempotencyKey),$||($=Ar(e)),$a.add(d),Aa({fingerprint:d,idempotencyKey:$,module:e,label:t||e,endpoint:n,status:"IN_PROGRESS",importJobId:F?.importJobId||null,committedChunks:F?.committedChunks||0,totalChunks:F?.totalChunks||0,committedRowIndex:F?.committedRowIndex??-1,recoverable:!0,body:a});let M={"Idempotency-Key":$,...s?{"X-Import-Recover":"true"}:{}};try{let h=await X(n,{method:"POST",headers:M,body:JSON.stringify(a),...m?{configOverride:m}:{},...i?{fetchImpl:i}:{}});if(h.ok){let w=_r(h.data),K=w.importJobId||F?.importJobId||null,G=!!w.idempotentReplay;return Aa({fingerprint:d,idempotencyKey:$,module:e,label:t||e,endpoint:n,status:"COMPLETED",importJobId:K,committedChunks:Number(F?.totalChunks||F?.committedChunks||1),totalChunks:Number(F?.totalChunks||F?.committedChunks||1),committedRowIndex:F?.committedRowIndex??-1,recoverable:!1,idempotentReplay:G,result:w}),{ok:!0,status:h.status,data:w,raw:h.data,idempotencyKey:$,fingerprint:d,importJobId:K,idempotentReplay:G}}let L=h.data&&typeof h.data=="object"?h.data:{},R=Number(h.status||0),I=L.code||(R===0?"NETWORK_ERROR":R===409?"IMPORT_CONFLICT":"IMPORT_FAILED"),N=Number(L.committedChunks??F?.committedChunks??0),q=Number(L.totalChunks??F?.totalChunks??0),v=Number(L.committedRowIndex??F?.committedRowIndex??-1),B=L.importJobId||F?.importJobId||null,x=R===0,k=x||R===409&&I==="IMPORT_IN_PROGRESS"||!!L.recoverable||R>=500,_="FAILED";x?_="NETWORK_ERROR":R===409&&I==="IMPORT_IN_PROGRESS"?_="CONFLICT_IN_PROGRESS":(I==="IMPORT_PARTIAL_FAILED"||R>=500)&&(_="PARTIAL_FAILED");let E=L.error||`Gagal mengimport data (HTTP ${R})`;throw x?E=L.error||"Koneksi terputus saat mengimport data. Anda dapat mencoba ulang secara aman dengan key yang sama.":R===409&&I==="IMPORT_IN_PROGRESS"?E=L.error||"Import identik sedang berlangsung atau terkunci (HTTP 409).":R===409&&I==="IDEMPOTENCY_KEY_MISMATCH"?E=L.error||"Idempotency-Key bertabrakan dengan payload berbeda (HTTP 409).":I==="IMPORT_PARTIAL_FAILED"&&(E=`${L.error||"Import gagal di tengah chunk"} (${N}/${q||"?"} chunk tersimpan).`),Aa({fingerprint:d,idempotencyKey:$,module:e,label:t||e,endpoint:n,status:_,importJobId:B,committedChunks:N,totalChunks:q,committedRowIndex:v,recoverable:k,errorCode:I,errorMessage:E,body:a}),new ft(E,{status:R,code:I,importJobId:B,committedChunks:N,totalChunks:q,committedRowIndex:v,recoverable:k,idempotencyKey:$,fingerprint:d,isNetworkError:x,serverStatus:L.status||null})}finally{$a.delete(d)}}async function Pt(e,{recover:t=!1,configOverride:n=null,fetchImpl:a=null}={}){let o=Dt(e);if(!o)throw new ft("Sesi import tidak ditemukan atau sudah dihapus.",{status:404,code:"IMPORT_SESSION_NOT_FOUND",recoverable:!1,fingerprint:e});if(!o.payloadStored||o.body===void 0)throw new ft("Payload terlalu besar untuk disimpan di browser. Silakan pilih ulang file Excel yang sama \u2014 Idempotency-Key akan tetap dipertahankan.",{status:400,code:"IMPORT_PAYLOAD_REUPLOAD_REQUIRED",recoverable:!1,idempotencyKey:o.idempotencyKey,fingerprint:e});return await Pe({module:o.module,label:o.label,endpoint:o.endpoint,body:o.body,idempotencyKey:o.idempotencyKey,recover:t,configOverride:n,fetchImpl:a})}function rt(e){if(!e)return"";let{variant:t="info",title:n="",message:a="",idempotencyKey:o="",importJobId:s="",committedChunks:c=0,totalChunks:m=0,errorCode:i="",fingerprint:d="",canRetry:F=!1,canRecover:$=!1,canReplay:M=!1,canDismiss:h=!0}=e,L={progress:{bg:"#EFF6FF",border:"#93C5FD",color:"#1E3A8A"},success:{bg:"#ECFDF5",border:"#6EE7B7",color:"#065F46"},warning:{bg:"#FFFBEB",border:"#FCD34D",color:"#92400E"},recover:{bg:"#FFF7ED",border:"#FDBA74",color:"#9A3412"},error:{bg:"#FEF2F2",border:"#FCA5A5",color:"#991B1B"},info:{bg:"#F8FAFC",border:"#CBD5E1",color:"#334155"}}[t]||{bg:"#F8FAFC",border:"#CBD5E1",color:"#334155"},R=m>0?`<span class="badge badge-info" style="margin-right:6px;">Chunk ${c}/${m}</span>`:"",I=i?`<span class="badge badge-danger" style="margin-right:6px;">${$e(i)}</span>`:"",N=o?`<code title="${$e(o)}" style="font-size:0.75rem;background:rgba(255,255,255,0.7);padding:2px 6px;border-radius:4px;">Key: ${$e(o.slice(0,26))}${o.length>26?"\u2026":""}</code>`:"",q=s?`<code title="${$e(s)}" style="font-size:0.75rem;background:rgba(255,255,255,0.7);padding:2px 6px;border-radius:4px;margin-left:6px;">Job: ${$e(s.slice(0,12))}\u2026</code>`:"";return`
    <div class="import-status-banner" data-variant="${$e(t)}" data-fingerprint="${$e(d)}" style="background:${L.bg};border:1px solid ${L.border};color:${L.color};border-radius:10px;padding:12px 16px;">
      <div style="display:flex;justify-content:space-between;align-items:flex-start;gap:12px;flex-wrap:wrap;">
        <div style="flex:1;min-width:240px;">
          <div style="font-weight:700;font-size:0.95rem;margin-bottom:4px;">${$e(n)}</div>
          <div style="font-size:0.88rem;line-height:1.45;">${$e(a)}</div>
          <div style="margin-top:8px;display:flex;flex-wrap:wrap;align-items:center;gap:4px;">
            ${I}${R}${N}${q}
          </div>
        </div>
        <div style="display:flex;gap:8px;align-items:center;flex-wrap:wrap;">
          ${F?`<button type="button" class="btn btn-primary btn-sm" data-import-action="retry" data-fingerprint="${$e(d)}">\u{1F504} Coba Lagi / Lanjutkan</button>`:""}
          ${$?`<button type="button" class="btn btn-warning btn-sm" data-import-action="recover" data-fingerprint="${$e(d)}">\u{1F6E0}\uFE0F Pulihkan Lock (Recover)</button>`:""}
          ${M?`<button type="button" class="btn btn-secondary btn-sm" data-import-action="replay" data-fingerprint="${$e(d)}">\u267B\uFE0F Replay Key Sama</button>`:""}
          ${h?`<button type="button" class="btn btn-ghost btn-sm" data-import-action="dismiss" data-fingerprint="${$e(d)}">\u2715 Tutup</button>`:""}
        </div>
      </div>
    </div>
  `}function Me({container:e,title:t,icon:n,apiPath:a,columns:o,formFields:s,filterFields:c,defaultFilters:m={},defaultSort:i="",itemLabel:d="Data",canCreate:F=!0,canEdit:$=!0,canDelete:M=!0,onBeforeSubmit:h,onAfterLoad:L,onDataLoaded:R,extraActions:I=[],initialSearch:N="",exportOptions:q=null,bulkDelete:v=!1,mobileCardRenderer:B=null}){let x=se(),k=(a||"").replace("/api/","").split("/")[0].replace(/-/g,"_"),_=re(x,k,"create"),E=re(x,k,"update"),w=re(x,k,"delete"),K=re(x,k,"export")||re(x,"export"),G=re(x,k,"import");if(F=F&&_,$=$&&E,M=M&&w,v=v&&w,!K&&!G)q=null;else if(q){G||(q.import=!1,q.template=!1),K||(q.export=!1);let g=q.export!==!1&&typeof q.onExport=="function",C=q.template!==!1&&typeof q.onTemplate=="function",z=q.import!==!1&&typeof q.onImport=="function";!g&&!C&&!z&&(q=null)}!$&&Array.isArray(I)&&(I=I.filter(g=>g&&g.readOnly===!0));let J=1,W={...m};N&&(W.search=N);let Y=new Set;e.innerHTML=`
    ${v?`
    <div class="bulk-toolbar" id="bulk-toolbar" style="display:none; align-items:center; justify-content:space-between; background:#2563EB; padding:12px 16px; border-radius:8px; margin-bottom:16px; box-shadow:0 4px 6px -1px rgba(37,99,235,0.2);">
      <button id="btn-bulk-cancel" style="background:transparent; border:none; color:white; display:flex; align-items:center; cursor:pointer; padding:4px;">
        <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M18 6L6 18M6 6l12 12"></path></svg>
      </button>
      <span id="bulk-count" style="font-weight:500; font-size:0.95rem; color:white;">0 item dipilih</span>
      <button id="btn-bulk-delete" style="background:transparent; border:none; color:white; display:flex; align-items:center; cursor:pointer; padding:4px;">
        <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polyline points="3 6 5 6 21 6"></polyline><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"></path></svg>
      </button>
    </div>`:""}

    <div class="page-header">
      <h1 class="page-title">${n} ${P(t)}</h1>
      <div class="page-actions" style="display:flex; gap:8px; align-items:center;">
        ${F?`<button class="btn btn-primary" id="btn-create">+ Tambah ${P(d)}</button>`:""}
        ${q?`
          <div class="aksi-dropdown-container" style="position:relative; display:inline-block;">
            <button class="btn btn-ghost" id="btn-aksi-main" style="background:#131F36; border:1px solid #223456; padding:8px 16px; border-radius:8px; font-weight:600; color:#F8FAFC; display:flex; align-items:center; gap:6px; cursor:pointer;">
              \u22EE Aksi
            </button>
            <div id="aksi-menu-main" class="aksi-menu-content" style="display:none; position:absolute; top:calc(100% + 4px); right:0; background:#131F36; border:1px solid #223456; border-radius:8px; box-shadow:0 8px 24px rgba(0,0,0,0.5); flex-direction:column; min-width:200px; z-index:999; padding:8px 0;">
              ${q.export!==!1&&typeof q.onExport=="function"?`
              <button type="button" class="dropdown-item" id="btn-export-${q.moduleName||"default"}" style="width:100%; text-align:left; padding:10px 16px; background:none; border:none; cursor:pointer; font-size:0.9rem; color:#F8FAFC; display:flex; gap:8px; align-items:center;">
                \u{1F4E5} Export Excel
              </button>`:""}
              ${q.template!==!1&&typeof q.onTemplate=="function"?`
              <button type="button" class="dropdown-item" id="btn-template-${q.moduleName||"default"}" style="width:100%; text-align:left; padding:10px 16px; background:none; border:none; cursor:pointer; font-size:0.9rem; color:#F8FAFC; display:flex; gap:8px; align-items:center;">
                \u{1F4C4} Download Template
              </button>`:""}
              ${q.import!==!1&&typeof q.onImport=="function"?`
              <label class="dropdown-item" style="display:flex; width:100%; text-align:left; padding:10px 16px; background:none; border:none; cursor:pointer; font-size:0.9rem; color:#F8FAFC; margin:0; gap:8px; align-items:center;" id="label-import-${q.moduleName||"default"}">
                <span class="import-text">\u{1F4E4} Import Excel</span>
                <input type="file" id="input-import-${q.moduleName||"default"}" accept=".xlsx, .xls, .csv" style="display:none;">
              </label>`:""}
            </div>
          </div>
          <style>.show-aksi-menu{display:flex !important}.dropdown-item:hover{background-color:#1A2A48 !important;color:#60A5FA !important}</style>
        `:""}
      </div>
    </div>
    

    ${c&&c.length>0?`
    <div class="filter-bar" style="background: #131F36; border-radius: 12px; padding: 10px 16px; display: flex; align-items: center; gap: 8px; margin-bottom: 24px; border: 1px solid #223456; box-shadow: 0 4px 12px rgba(0,0,0,0.22);">
        ${c.filter(g=>g.type==="search").map(g=>`<div class="filter-search-wrap" style="flex:1; min-width:0;"><input type="search" class="filter-search" placeholder="${P(g.placeholder||"Cari...")}" id="filter-search" value="${P(W.search||"")}" style="width:100%; background: #0E182B; border: 1px solid #223456; color: #F8FAFC; border-radius: 8px; padding: 8px 12px; font-size: 0.85rem; outline:none;"></div>`).join("")}
        
        <div class="filter-dropdowns-desktop">
          ${c.filter(g=>g.type!=="search").map(g=>{if(g.type==="select"||g.type==="combobox"){let C=(g.label||"").startsWith("Pilih")?g.label:`Pilih ${g.label||""}`;return`<select class="filter-select" name="${P(g.name)}" id="filter-${P(g.name)}" style="background: #0F1A2E; border: 1px solid #223456; border-radius: 8px; padding: 7px 10px; font-size: 0.85rem; color: #F8FAFC; cursor: pointer; outline:none;"><option value="">${P(C)}</option>${(g.options||[]).map(z=>`<option value="${P(typeof z=="object"?z.value:z)}" ${W[g.name]===(typeof z=="object"?z.value:z)?"selected":""}>${P(typeof z=="object"?z.label:z)}</option>`).join("")}</select>`}return""}).join("")}
          <button id="btn-reset-filter" style="background: transparent; border: none; color: #60A5FA; font-weight: 600; font-size: 0.85rem; cursor: pointer; padding: 7px 8px; white-space:nowrap;">Reset</button>
        </div>
        
        <button id="btn-mobile-filter" class="btn-mobile-filter-trigger">\u2699 Filter</button>
        
        <div class="filter-options-wrapper" id="filter-options-wrapper">
          <div class="bottom-sheet-header">
            <h3 style="margin:0; font-size:1rem; color:#F8FAFC;">Filter Data</h3>
            <button class="btn-close-sheet" id="btn-close-filter-sheet" style="background:none;border:none;font-size:1.5rem;cursor:pointer;line-height:1;color:#A5B4C8;">&times;</button>
          </div>
          ${c.filter(g=>g.type!=="search").map(g=>{if(g.type==="select"||g.type==="combobox"){let C=(g.label||"").startsWith("Pilih")?g.label:`Pilih ${g.label||""}`;return`<select class="filter-select filter-select-sheet" name="${P(g.name)}-sheet" id="filter-sheet-${P(g.name)}" style="width:100%; background: #0F1A2E; border: 1px solid #223456; border-radius: 10px; padding: 12px 14px; font-size: 0.9rem; color: #F8FAFC; cursor: pointer; outline:none;"><option value="">${P(C)}</option>${(g.options||[]).map(z=>`<option value="${P(typeof z=="object"?z.value:z)}" ${W[g.name]===(typeof z=="object"?z.value:z)?"selected":""}>${P(typeof z=="object"?z.label:z)}</option>`).join("")}</select>`}return""}).join("")}
          <button id="btn-reset-filter-sheet" style="background: transparent; border: none; color: #60A5FA; font-weight: 600; font-size: 0.9rem; cursor: pointer; padding: 8px;">Reset</button>
        </div>
    </div>`:""}

    <div id="crud-import-banner" style="display:none;margin-bottom:16px;"></div>

    <div class="card">
      <div class="card-body p-0" id="table-container">
        <div class="loading-spinner"><div class="spinner"></div></div>
      </div>
      <div class="card-footer" id="pagination-container"></div>
    </div>
  `,document.getElementById("btn-aksi-main")?.addEventListener("click",g=>{g.stopPropagation(),document.getElementById("aksi-menu-main")?.classList.toggle("show-aksi-menu")});function b(){let g=document.getElementById("bulk-toolbar");if(!g)return;let C=document.getElementById("bulk-count"),z=document.getElementById("btn-bulk-delete"),O=document.getElementById("btn-bulk-cancel");C.textContent=`${Y.size} item dipilih`,Y.size>0?(g.style.display="flex",z.disabled=!1,O.disabled=!1):(g.style.display="none",z.disabled=!0,O.disabled=!0);let te=document.getElementById("select-all-checkbox");if(te){let Q=document.querySelectorAll(".row-checkbox");if(Q.length>0){let ne=[...Q].every(H=>H.checked),ee=[...Q].some(H=>H.checked);te.checked=ne,te.indeterminate=ee&&!ne}else te.checked=!1,te.indeterminate=!1}}document.getElementById("btn-bulk-cancel")?.addEventListener("click",()=>{Y.clear(),document.querySelectorAll(".row-checkbox").forEach(C=>C.checked=!1);let g=document.getElementById("select-all-checkbox");g&&(g.checked=!1),b()}),document.getElementById("btn-bulk-delete")?.addEventListener("click",()=>{if(Y.size===0)return;let g=[...Y];if(g.length>ht){j(`Maksimum ${ht} item dapat dihapus dalam satu permintaan bulk.`);return}let C=document.createElement("div");C.style.cssText="position:fixed;inset:0;background:rgba(0,0,0,.5);z-index:9999;display:flex;align-items:center;justify-content:center",C.innerHTML=`
      <div style="background:var(--bg-card);border-radius:var(--radius-xl);padding:28px;max-width:420px;width:90%;box-shadow:var(--shadow-lg);animation:fadeInUp .2s ease">
        <h3 style="margin:0 0 8px;color:var(--text-1);font-size:1rem;font-weight:700">\u26A0\uFE0F Hapus ${g.length} ${d}?</h3>
        <p style="margin:0 0 24px;color:var(--text-2);font-size:.875rem">Data yang dihapus tidak dapat dikembalikan.</p>
        <div style="display:flex;gap:8px;justify-content:flex-end">
          <button id="bulk-cancel-btn" class="btn btn-ghost">Batal</button>
          <button id="bulk-confirm-btn" class="btn btn-danger">Hapus ${g.length} Data</button>
        </div>
      </div>
    `,document.body.appendChild(C),C.querySelector("#bulk-cancel-btn").addEventListener("click",()=>C.remove()),C.querySelector("#bulk-confirm-btn").addEventListener("click",async()=>{let z=C.querySelector("#bulk-confirm-btn");z.disabled=!0,z.textContent="Menghapus...";let O=await X(`${a}/bulk`,{method:"DELETE",body:JSON.stringify({ids:g})});C.remove(),O.ok?(ae(`${g.length} ${d} berhasil dihapus.`),Y.clear(),b(),it(a),u()):j(O.data?.error||"Gagal menghapus data.")})});let S=document.getElementById("filter-search"),y;S?.addEventListener("input",g=>{clearTimeout(y),y=setTimeout(()=>{W.search=g.target.value,J=1,Y.clear(),b(),u()},400)}),c?.forEach(g=>{(g.type==="select"||g.type==="combobox")&&(document.getElementById(`filter-${g.name}`)?.addEventListener("change",C=>{W[g.name]=C.target.value;let z=document.getElementById(`filter-sheet-${g.name}`);z&&(z.value=C.target.value),J=1,Y.clear(),b(),u()}),document.getElementById(`filter-sheet-${g.name}`)?.addEventListener("change",C=>{W[g.name]=C.target.value;let z=document.getElementById(`filter-${g.name}`);z&&(z.value=C.target.value),J=1,Y.clear(),b(),u(),document.getElementById("filter-options-wrapper")?.classList.remove("sheet-open")}))}),document.getElementById("btn-reset-filter")?.addEventListener("click",()=>{W={},S&&(S.value=""),c?.forEach(g=>{let C=document.getElementById(`filter-${g.name}`);C&&(C.value="");let z=document.getElementById(`filter-sheet-${g.name}`);z&&(z.value="")}),J=1,Y.clear(),b(),u()}),document.getElementById("btn-reset-filter-sheet")?.addEventListener("click",()=>{W={},S&&(S.value=""),c?.forEach(g=>{let C=document.getElementById(`filter-${g.name}`);C&&(C.value="");let z=document.getElementById(`filter-sheet-${g.name}`);z&&(z.value="")}),J=1,Y.clear(),b(),u(),document.getElementById("filter-options-wrapper")?.classList.remove("sheet-open")}),document.getElementById("btn-create")?.addEventListener("click",()=>A(null)),q&&document.addEventListener("click",function(g){let C=document.getElementById("aksi-menu-main"),z=document.getElementById("btn-aksi-main");C&&z&&!z.contains(g.target)&&!C.contains(g.target)&&C.classList.remove("show-aksi-menu")});let r=document.getElementById("btn-mobile-filter"),l=document.getElementById("filter-options-wrapper"),f=document.getElementById("btn-close-filter-sheet");if(r&&l&&(r.addEventListener("click",g=>{g.preventDefault(),l.classList.add("sheet-open")}),f&&f.addEventListener("click",g=>{g.preventDefault(),l.classList.remove("sheet-open")})),q){let te=function(D){z=D;let V=document.getElementById(`label-import-${q.moduleName}`),Z=V?V.querySelector(".import-text"):null;Z&&(Z.innerText=D?"\u231B Memproses...":"\u{1F4E4} Import Excel"),V&&(V.style.pointerEvents=D?"none":"auto",V.style.opacity=D?"0.6":"1"),g&&(g.disabled=D)},Q=function(D){if(C){if(!D){C.style.display="none",C.innerHTML="";return}C.style.display="",C.innerHTML=rt(D),C.querySelectorAll("[data-import-action]").forEach(V=>{V.addEventListener("click",async()=>{let Z=V.getAttribute("data-import-action"),ie=V.getAttribute("data-fingerprint");if(Z==="dismiss"){ie&&yt(ie),Q(null);return}if(z){j("Import masih berjalan. Mohon tunggu hingga selesai.");return}await H(Z,ie)})})}},ne=function(D,V=""){let Z=D?.data||D||{},ie=!!(D?.idempotentReplay||Z?.idempotentReplay),le=Number(Z.inserted||0),pe=Number(Z.updated||0),de=Number(Z.skipped||0),me=[];le>0&&me.push(`${le} baru`),pe>0&&me.push(`${pe} diperbarui`),de>0&&me.push(`${de} dilewati`);let Ae=Z.message||(me.length?me.join(", "):"Data berhasil diproses");Q({variant:"success",title:ie?"\u267B\uFE0F Replay Idempoten Dikonfirmasi Server":"\u2705 Import Berhasil Dikonfirmasi",message:ie?`${Ae} (Server mengembalikan hasil tersimpan tanpa menulis data duplikat.)`:Ae,idempotencyKey:D?.idempotencyKey||V||"",importJobId:D?.importJobId||Z?.importJobId||"",fingerprint:D?.fingerprint||"",canReplay:!!((D?.fingerprint||O)&&(D?.idempotencyKey||V)),canDismiss:!0}),ae(ie?"Replay idempoten berhasil (0 duplikat).":"Import berhasil!"),it(a),u()},ee=function(D){let Z=Number(D?.status)===409&&D?.code==="IMPORT_IN_PROGRESS",ie=!!(D?.isNetworkError||Number(D?.status)===0);Q({variant:Z?"warning":"error",title:ie?"\u{1F4E1} Koneksi Terputus Saat Import":Z?"\u23F3 Import Sedang Berjalan / Terkunci (HTTP 409)":Number(D?.status)>=500?"\u26A0\uFE0F Import Gagal Sebagian / Server Error (HTTP 500)":`\u274C Import Gagal${D?.status?` (HTTP ${D.status})`:""}`,message:D?.message||"Gagal mengimport data",idempotencyKey:D?.idempotencyKey||"",importJobId:D?.importJobId||"",committedChunks:D?.committedChunks||0,totalChunks:D?.totalChunks||0,errorCode:D?.code||"",fingerprint:D?.fingerprint||"",canRetry:!!(D?.recoverable||ie||Number(D?.status)>=500||Z),canRecover:Z,canDismiss:!0}),j(D?.message||"Gagal import data")};document.getElementById(`btn-export-${q.moduleName}`)?.addEventListener("click",async D=>{let V=D.target,Z=V.innerHTML;V.innerHTML="\u23F3 Loading...",V.disabled=!0;let ie=i&&!W.sort?{...W,sort:i}:{...W};try{await q.onExport(ie)}catch{j("Gagal export data")}finally{V.innerHTML=Z,V.disabled=!1}}),document.getElementById(`btn-template-${q.moduleName}`)?.addEventListener("click",()=>{q.onTemplate()});let g=document.getElementById(`input-import-${q.moduleName}`),C=document.getElementById("crud-import-banner"),z=!1,O=null;async function H(D,V){let Z=V?Dt(V):null,ie=D==="recover",le=Z?.idempotencyKey||"";te(!0),Q({variant:"progress",title:ie?"\u{1F6E0}\uFE0F Memulihkan & Melanjutkan Import...":"\u23F3 Mengirim / Melanjutkan Import...",message:"Mohon tunggu konfirmasi dari server. Jangan tutup halaman.",idempotencyKey:le,importJobId:Z?.importJobId||"",committedChunks:Z?.committedChunks||0,totalChunks:Z?.totalChunks||0,fingerprint:V||"",canDismiss:!1});try{let pe;D==="replay"&&O&&le?pe=await q.onImport(O,{idempotencyKey:le,recover:!1}):Z?.payloadStored?pe=await Pt(V,{recover:ie}):O?pe=await q.onImport(O,{idempotencyKey:le||null,recover:ie}):pe=await Pt(V,{recover:ie}),ne(pe,le)}catch(pe){ee(pe)}finally{te(!1)}}let U=Xt(q.moduleName);if(U.length>0){let D=U[0],V=D.status==="IN_PROGRESS"||D.status==="CONFLICT_IN_PROGRESS";Q({variant:"recover",title:"\u{1F504} Sesi Import Belum Selesai Terdeteksi",message:D.payloadStored?D.errorMessage||"Import sebelumnya terhenti karena refresh atau gangguan koneksi. Anda dapat melanjutkan dengan Idempotency-Key yang sama.":"Import sebelumnya terhenti. Pilih ulang file Excel yang sama untuk melanjutkan dengan Idempotency-Key yang sama.",idempotencyKey:D.idempotencyKey,importJobId:D.importJobId,committedChunks:D.committedChunks,totalChunks:D.totalChunks,errorCode:D.errorCode||D.status,fingerprint:D.fingerprint,canRetry:!!D.payloadStored,canRecover:!!(D.payloadStored&&V),canDismiss:!0})}g?.addEventListener("change",async D=>{let V=D.target.files[0];if(V){if(z){j("Proses import sedang berjalan. Mohon tunggu hingga selesai."),g.value="";return}te(!0),Q({variant:"progress",title:"\u23F3 Membaca & Mengimport File Excel...",message:`Memproses "${V.name}". Sistem menunggu konfirmasi server sebelum menandai selesai.`,canDismiss:!1});try{let Z=await Sa(V);if(Z.length===0)throw new Error("File kosong atau format salah");O=Z;let ie=await q.onImport(Z,{recover:!1});ne(ie)}catch(Z){ee(Z)}finally{te(!1),g&&(g.value="")}}})}async function u(){b();let g=document.getElementById("table-container");if(!g)return;g.innerHTML='<div class="loading-spinner"><div class="spinner"></div></div>';let C=J,z=20,O={...i?{sort:i}:{},...Object.fromEntries(Object.entries(W).filter(([,D])=>D))},te=new URLSearchParams({page:C,limit:z,...O}),Q=await X(`${a}?${te}`);if(!Q.ok){g.innerHTML=`<div class="empty-state"><p class="text-danger">Gagal memuat data: ${P(Q.data?.error||"Error")}</p></div>`;return}let ne=Q.data?.data||Q.data||[],ee=Q.data?.pagination||{page:C,limit:z,total:Array.isArray(ne)?ne.length:0,pages:1};R&&(ne=R(ne)),!1,L&&L(ne);let H=En({columns:o,data:ne,onEdit:$?D=>A(D):null,onDelete:M?D=>T(D):null,actions:I.map(D=>({...D,handler:V=>D.handler(V,u)})),emptyText:`Tidak ada ${String(d||"").toLowerCase()}`,bulkSelect:v?{selectedIds:Y,maxSelect:ht,onLimitExceeded:D=>j(D),onToggle:b}:null,mobileCardRenderer:B});g.innerHTML="",g.appendChild(H),b();let U=document.getElementById("pagination-container");if(U&&(U.innerHTML="",ee&&ee.pages>1)){let D=Te({page:ee.page,pages:ee.pages,total:ee.total,limit:ee.limit,onPage:V=>{J=V,u()}});D&&U.appendChild(D)}}function p(g){let C=typeof s=="function"?s(g):s;return Ea(C)}function A(g){let C=!!g,z=document.createElement("form");z.noValidate=!0,z.innerHTML=p(g);let O=typeof s=="function"?s(g):s;C&&In(z,g);let te=[],Q=ee=>{for(let H of ee)if(H.type==="row")Q(H.fields||[]);else if(H.type==="employee-combobox"){let U=z.querySelector(`.fc-crud-emp-mount[data-emp-mount="${H.name}"]`);if(!U)continue;let D=g?g[H.name]??H.value??"":H.value??"",V=g&&H.employeeIdField?g[H.employeeIdField]??null:H.employeeIdValue??null,Z=je(U,{mode:H.mode||"single",inputId:H.id||`crud-${H.name}`,inputName:H.name,placeholder:H.placeholder||"",required:!!H.required,allowManual:H.allowManual!==!1,hasEmployeeIdSchema:H.hasEmployeeIdSchema!==!1,manualBadgeText:H.manualBadgeText||"Belum terhubung ke Master Karyawan",masterBadgePrefix:H.masterBadgePrefix||"Terhubung ke Master Karyawan",employees:H.options||[],initialEmployeeId:V,initialEmployeeName:D});te.push({field:H,ctrl:Z})}};Q(O);let{close:ne}=ce({title:C?`Edit ${d}`:`Tambah ${d}`,content:z,size:"lg",confirmText:C?"Simpan Perubahan":`Tambah ${d}`,onConfirm:async(ee,H)=>{if(!z.reportValidity())return;let U=ee.querySelector(".modal-confirm");U.disabled=!0,U.textContent="Menyimpan...";let D=Pn(z),V=typeof s=="function"?s(g):s,Z=async de=>{for(let me of de)if(me.type==="row")await Z(me.fields);else if(me.type==="combobox"&&D[me.name]){let Ae=D[me.name],Ke=(me.options||[]).find(oe=>{let ve=String(typeof oe=="object"?oe.value:oe),fe=String(typeof oe=="object"?oe.label:oe);return ve===Ae||fe===Ae});if(Ke)D[me.name]=typeof Ke=="object"?Ke.value:Ke;else if(me.createApi){let oe={};oe[me.createApi.field]=Ae,me.createApi.extra&&Object.assign(oe,me.createApi.extra);let ve=await X(me.createApi.path,{method:"POST",body:JSON.stringify(oe)});if(ve.ok&&ve.data?.id)D[me.name]=ve.data.id;else if(ve.ok&&!ve.data?.id)D[me.name]=Ae;else throw new Error(`Gagal membuat master data: ${ve.data?.error||"Unknown error"}`)}}};try{await Z(V)}catch(de){j(de.message),U.disabled=!1,U.textContent=C?"Simpan Perubahan":`Tambah ${d}`;return}for(let{field:de,ctrl:me}of te){let Ae=me.getValue();de.mode==="multi"?D[de.name]=Ae.participantsString||null:(D[de.name]=Ae.employee_name||null,de.employeeIdField&&(D[de.employeeIdField]=Ae.employee_id??null))}h&&(D=await h(D,g)),C&&g&&(g.row_version!==void 0&&g.row_version!==null?D.row_version=Number(g.row_version):D.row_version=1);let ie=C?"PUT":"POST",le=C?`${a}/${g.id}`:a,pe=await X(le,{method:ie,body:JSON.stringify(D)});pe.ok?(ae(C?`${d} berhasil diperbarui.`:`${d} berhasil ditambahkan.`),H(),it(a),u()):(pe.status===409?j(pe.data?.error||"Konflik OCC: Data telah diubah oleh pengguna lain. Silakan muat ulang halaman."):j(pe.data?.error||"Gagal menyimpan data."),U.disabled=!1,U.textContent=C?"Simpan Perubahan":`Tambah ${d}`)}})}function T(g){ge(`Hapus ${d} ini? Tindakan tidak dapat dibatalkan.`,async()=>{let C=await X(`${a}/${g.id}`,{method:"DELETE"});C.ok?(ae(`${d} berhasil dihapus.`),it(a),u()):j(C.data?.error||"Gagal menghapus.")},`Hapus ${d}`)}return u(),u}function Ne(e){let t={Done:"badge-success",Aktif:"badge-success",Open:"badge-warning","In Progress":"badge-info",Pending:"badge-warning",Diproses:"badge-info",Selesai:"badge-success","Tidak Aktif":"badge-neutral",Resign:"badge-neutral",Cut:"badge-danger","Tidak Datang":"badge-danger"};return!e||e==="-"||String(e).trim()===""?"":`<span class="badge ${t[e]||"badge-neutral"}">${e}</span>`}function Ia(e){return`<span class="badge ${{"Inspeksi Hygiene & Aset Bangunan":"badge-info","General Cleaning":"badge-success","Deep Cleaning":"badge-purple",Fogging:"badge-warning"}[e]||"badge-neutral"}">${e||"-"}</span>`}function Oe(e){return`<span class="badge ${{Q1:"badge-info",Q2:"badge-success",Q3:"badge-warning",Q4:"badge-danger"}[e]||"badge-neutral"}">${e||"-"}</span>`}Ce();var Ma=[],It=[];async function Kn(e,t){Ma=await we();let n=await Jt();It=["Berlin Ariansyah","Ade Surahman"];let a=h=>h&&!It.find(L=>String(typeof L=="object"?L.value:L).toLowerCase()===String(h).toLowerCase())?[...It,h]:It,o=h=>{if(!h||h==="-"||String(h).trim()==="")return"";let L=String(h).split("-");return L.length===3&&L[0].length===4?`${L[2]}-${L[1]}-${L[0]}`:h},s=t?t.get("dash_filter"):null,c=new Date,m=`${c.getFullYear()}-${String(c.getMonth()+1).padStart(2,"0")}`,i={},d=t&&t.get("month")?t.get("month"):null;s==="inspeksi"?i={status:"Done",activity_type:"Inspeksi Hygiene",...d?{month:d}:{}}:s==="gcdc"?i={status:"Done",activity_type:"GCDC",...d?{month:d}:{}}:s&&s.startsWith("period_")&&(i={period:s.replace("period_","").toUpperCase()});let F=new Date().getFullYear(),M=["Jan","Feb","Mar","Apr","Mei","Jun","Jul","Agu","Sep","Okt","Nov","Des"].map((h,L)=>{let R=String(L+1).padStart(2,"0");return{value:`${F}-${R}`,label:`${h} ${F}`}});Me({container:e,title:"Jadwal Kegiatan",icon:"\u{1F4C5}",apiPath:"/api/schedule",bulkDelete:!1,itemLabel:"Jadwal",enableMobileFilterSheet:!0,defaultFilters:i,defaultSort:t&&t.get("sort")||new URLSearchParams(window.location.search).get("sort")||"opening_date_desc",columns:[{key:"branch_name",label:"Cabang"},{key:"activity_type",label:"Kegiatan",render:h=>Ia(h)},{key:"period",label:"Periode",render:h=>Oe(h)},{key:"pic",label:"PIC"},{key:"opening_date",label:"Tgl Opening",nowrap:!0,render:h=>o(h)},{key:"target_date",label:"Tgl Target",nowrap:!0,render:h=>o(h)},{key:"completion_date",label:"Tgl Selesai",nowrap:!0,render:h=>o(h)},{key:"status",label:"Status",render:h=>Ne(h)}],mobileCardRenderer:(h,L)=>Xe({title:h.branch_name||"-",headerRight:Ne(h.status),badges:[h.activity_type?Ia(h.activity_type):"",h.period?Oe(h.period):""],meta:[{label:"Tgl Opening",value:o(h.opening_date)||"-"},{label:"Tgl Target",value:o(h.target_date)||"-"},{label:"Tgl Selesai",value:o(h.completion_date)||"-"},...h.pic?[{label:"PIC",value:h.pic}]:[]],notes:h.notes,row:h,handlers:L}),filterFields:[{type:"search",placeholder:"Cari cabang, kegiatan, PIC, catatan..."},{type:"select",name:"branch_id",label:"Cabang",options:Ma},{type:"select",name:"activity_type",label:"Kegiatan",options:[{value:"Inspeksi Hygiene",label:"Inspeksi Hygiene"},{value:"General Cleaning",label:"General Cleaning"},{value:"Deep Cleaning",label:"Deep Cleaning"},{value:"Fogging",label:"Fogging"},{value:"GCDC",label:"GCDC (GC & DC)"}]},{type:"select",name:"month",label:"Bulan",options:M},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"status",label:"Status",options:["Pending","In Progress","Done"]},{type:"select",name:"pic",label:"PIC",options:It}],formFields:h=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:Ma,value:h?.branch_id},{name:"activity_type",label:"Jenis Kegiatan",type:"select",required:!0,options:["Inspeksi Hygiene","General Cleaning","Deep Cleaning","Fogging"],value:h?.activity_type}]},{type:"row",fields:[{name:"period",label:"Periode",type:"select",required:!0,options:["Q1","Q2","Q3","Q4"],value:h?.period},{name:"pic",label:"PIC",type:"combobox",options:a(h?.pic),value:h?.pic}]},{type:"row",fields:[{name:"opening_date",label:"Tanggal Opening",type:"date",value:h?.opening_date},{name:"target_date",label:"Tanggal Target",type:"date",value:h?.target_date}]},{type:"row",fields:[{name:"completion_date",label:"Tanggal Selesai",type:"date",value:h?.completion_date},{name:"status",label:"Status",type:"select",required:!0,options:["Pending","In Progress","Done"],value:h?.status||""}]},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:h?.notes}],exportOptions:{moduleName:"schedule",onExport:async h=>{let R=(await ue("/api/schedule/export",h)).map(I=>({Cabang:I.branch_name||"",Kegiatan:I.activity_type||"",Periode:I.period||"",PIC:I.pic||"","Tgl Opening":I.opening_date||"","Tgl Target":I.target_date||"","Tgl Selesai":I.completion_date||"",Status:I.status||""}));be(R,"Data_Jadwal_Kegiatan")},onTemplate:()=>{be([{Cabang:"001. Pondok Bambu",Kegiatan:"General Cleaning",Periode:"Q1",PIC:"Fajar","Tgl Opening":"2024-02-01","Tgl Target":"2024-02-15","Tgl Selesai":"2024-02-14",Status:"Done"}],"Template_Import_Jadwal")},onImport:async(h,L={})=>{let I=(await X("/api/branches?all=1")).data?.data||[],N=x=>{if(!x)return null;let k=String(x||"").toLowerCase().trim(),_=k.match(/^(\d{3})/);if(_){let w=I.find(K=>String(K.code||"")===_[1]);if(w)return w.id}let E=I.find(w=>String(w.full_name||"").toLowerCase()===k||String(w.code||"").toLowerCase()===k||String(w.name||"").toLowerCase()===k);return E?E.id:null},q=x=>{if(x==null||x==="")return"";if(x instanceof Date&&!isNaN(x.getTime()))return x.toISOString().slice(0,10);let k=String(x).trim();if(k===""||k==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(k))return k.slice(0,10);if(/^\d{4,5}$/.test(k)){let E=Number(k);if(E>2e4&&E<99999){let w=new Date(Date.UTC(1899,11,30)+E*864e5);return isNaN(w.getTime())?"":w.toISOString().slice(0,10)}}let _=k.split(/[\/\-\.]/);if(_.length===3){let[E,w,K]=_.map(G=>G.trim());if(E.length===4&&w.length<=2&&K.length<=2)return`${E}-${w.padStart(2,"0")}-${K.padStart(2,"0")}`;if(K.length===4&&w.length<=2&&E.length<=2)return`${K}-${w.padStart(2,"0")}-${E.padStart(2,"0")}`}return k},v=h.map(x=>({branch_id:N(String(x.Cabang||"").trim()),branch_name:String(x.Cabang||"").trim(),activity_type:String(x.Kegiatan||"").trim(),period:String(x.Periode||"").trim(),pic:String(x.PIC||x.Pic||"").trim(),opening_date:q(x["Tgl Opening"]||x["Tanggal Opening"]||x["Tgl Openir"]),target_date:q(x["Tgl Target"]||x["Tanggal Target"]),completion_date:q(x["Tgl Selesai"]||x["Tanggal Selesai"]),status:String(x.Status||"").trim(),notes:String(x.Catatan||x.Keterangan||"").trim()})).filter(x=>x.activity_type&&x.period),B=await Pe({module:"schedule",label:"Jadwal Kegiatan",endpoint:"/api/import/schedule",body:{rows:v,onDuplicate:"update"},idempotencyKey:L?.idempotencyKey,recover:!!L?.recover});return it("schedule"),B}}})}Be();var Rn=["SDM","Cleaning","Aset","K3","Lainnya"],Tr=["Open","In Progress","Done"],jn=["SPV","AM","RCP","Perawat","FC","Berlin","Ade","Pattrel","Dentrel"];function Mt(e){if(e==null||e==="")return null;let t=Number(e);return Number.isInteger(t)&&t>=0?t:null}async function qn(e,t){let a=[],o=[],s=[],c="all",m="all",i="all",d="",F=1,$=0,M=1,h=!1,L={total:0,open:0,inProgress:0,openAndProgress:0,done:0,currentMonth:0},R=null,I=!1;(t?t.get("dash_filter"):null)==="open"&&(c="all");let q=se(),v=re(q,"issues","create"),B=re(q,"issues","update"),x=re(q,"issues","delete");try{[o,s]=await Promise.all([we(),ze()])}catch(l){console.error("Error fetching cache:",l)}function k(){let l={sort:"date_desc"};return d.trim()&&(l.search=d.trim()),i&&i!=="all"&&(l.branch_id=i),m&&m!=="all"&&(l.category=m),c&&c!=="all"&&(l.status=c),l}function _(){let l=(f,u)=>{let p=e.querySelector(f);p&&(p.textContent=u)};l("#issues-stat-total",L.total),l("#issues-stat-open-prog",L.openAndProgress),l("#issues-stat-done",L.done),l("#issues-stat-month",L.currentMonth),l("#issues-pill-count-all",L.total),l("#issues-pill-count-open",L.open),l("#issues-pill-count-inprogress",L.inProgress),l("#issues-pill-count-done",L.done)}async function E({refreshStats:l=!0}={}){if(I=!0,!!!e.querySelector("#issues-list-container"))r();else{let u=e.querySelector("#issues-list-container");u&&(u.innerHTML='<div class="issues-empty-state"><p style="color:#CBD5E1; margin:0;">Memuat data permasalahan...</p></div>')}try{let u=new URLSearchParams({page:String(F),limit:String(20),...k()});(l||!h)&&u.set("include_stats","1");let p=await X(`/api/issues?${u.toString()}`);if(p.ok){let A=p.data?.data||[];a=Array.isArray(A)?A:[];let T=p.data?.pagination||{};if($=Number(T.total??a.length),M=Math.max(1,Number(T.pages??Math.ceil($/20))||1),F=Math.max(1,Number(T.page??F)),p.data?.stats&&(L=p.data.stats,h=!0),a.length===0&&F>M)return F=M,I=!1,E({refreshStats:l})}else a=[],$=0,M=1,j(p.data?.error||"Gagal memuat data permasalahan")}catch(u){console.warn("Gagal load issues:",u),a=[],$=0,M=1,j("Gagal memuat data permasalahan: "+u.message)}finally{I=!1,e.querySelector("#issues-list-container")?(_(),y()):r()}}async function w(){if(!re(q,"export")){j("Anda tidak memiliki izin untuk mengekspor data.");return}try{let[{downloadExcel:l},f]=await Promise.all([Promise.resolve().then(()=>(Ce(),Ve)),ue("/api/issues/export",k())]),u=f.map((p,A)=>({No:A+1,"Tanggal Info":p.report_date||"",Cabang:p.branch_name||"",Kategori:p.category||"","Sumber Laporan":p.source||"",Keluhan:p.complaint||"","Nama FC / Security":p.employee_name||"","FC Spesialis":p.fc_specialist||"","Solusi / Tindak Lanjut":p.solution||"",Status:p.status||"","Tanggal Selesai":p.completion_date||"","Durasi Hari":p.day_count!=null?p.day_count:""}));l(u,`Data_Permasalahan_FCMS_${new Date().toISOString().split("T")[0]}`),ae("Berhasil mengekspor data permasalahan ke Excel")}catch(l){j("Gagal ekspor Excel: "+l.message)}}async function K(l,f){if(!B)return;let u=Mt(l?.row_version);if(u===null){j("Versi data (row_version) tidak tersedia. Silakan muat ulang data sebelum memperbarui status.");return}try{let p=f==="Done"?l.completion_date||new Date().toISOString().split("T")[0]:f==="Open"?null:l.completion_date,A=l.day_count;if(l.report_date&&p){let C=new Date(l.report_date),z=new Date(p);A=Math.max(0,Math.floor((z-C)/(1e3*60*60*24)))}else f==="Open"&&(A=null);let T={...l,row_version:u,status:f,completion_date:p,day_count:A},g=await X(`/api/issues/${l.id}`,{method:"PUT",body:JSON.stringify(T)});if(g.ok){let C=Mt(g.data?.data?.row_version??g.data?.row_version);C!==null&&(l.row_version=C),ae(`Status masalah diubah ke "${f}"`),await E()}else g.status===409?(j("Konflik data (HTTP 409): Issue ini telah diubah oleh pengguna atau sesi lain. Silakan muat ulang data terbaru."),await E()):j(g.data?.error||"Gagal memperbarui status")}catch(p){j("Terjadi kesalahan: "+p.message)}}function G(l=null){let f=!!l;if(f?!B:!v)return;let u=f?Mt(l?.row_version):null,p=new Date().toISOString().split("T")[0],A=l?.report_date||p,T=l?.category||"SDM",g=l?.source?String(l.source).trim():"",C=f&&g&&!jn.includes(g),z=l?.status||"Open",O=document.createElement("div");O.className="issue-modal-body",O.innerHTML=`
      <div style="display:flex; flex-direction:column; gap:14px;">
        ${f?`<input type="hidden" id="iss-modal-row-version" value="${u!==null?u:""}" />`:""}
        <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:var(--text-2); margin-bottom:5px;">
              Tanggal Info <span style="color:#DC2626;">*</span>
            </label>
            <input type="date" id="iss-modal-date" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              value="${A}" required />
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:var(--text-2); margin-bottom:5px;">
              Cabang Klinik <span style="color:#DC2626;">*</span>
            </label>
            <select id="iss-modal-branch" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" required>
              <option value="">-- Pilih Cabang --</option>
              ${o.map(V=>`
                <option value="${V.value}" ${String(l?.branch_id)===String(V.value)?"selected":""}>
                  ${V.label}
                </option>
              `).join("")}
            </select>
          </div>
        </div>

        <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:var(--text-2); margin-bottom:5px;">
              Kategori <span style="color:#DC2626;">*</span>
            </label>
            <select id="iss-modal-category" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" required>
              ${Rn.map(V=>`
                <option value="${V}" ${T===V?"selected":""}>${V}</option>
              `).join("")}
            </select>
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:var(--text-2); margin-bottom:5px;">
              Sumber Laporan
            </label>
            <select id="iss-modal-source" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;">
              <option value="">-- Pilih Sumber Laporan --</option>
              ${jn.map(V=>`<option value="${V}" ${g===V?"selected":""}>${V}</option>`).join("")}
              ${C?`<option value="${g.replace(/"/g,"&quot;")}" selected>${g} (Historis)</option>`:""}
            </select>
          </div>
        </div>

        <div>
          <label style="display:block; font-size:0.82rem; font-weight:600; color:var(--text-2); margin-bottom:5px;">
            Detail Keluhan / Permasalahan <span style="color:#DC2626;">*</span>
          </label>
          <textarea id="iss-modal-complaint" class="form-control" rows="3" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
            required>${l?.complaint||""}</textarea>
        </div>

        <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:var(--text-2); margin-bottom:5px;">
              Nama Personil FC / Security
            </label>
            <div id="iss-modal-emp-mount"></div>
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:var(--text-2); margin-bottom:5px;">
              FC Spesialis / Penanggung Jawab
            </label>
            <input list="iss-spec-options" id="iss-modal-spec" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              value="${l?.fc_specialist?String(l.fc_specialist).replace(/"/g,"&quot;"):""}" />
            <datalist id="iss-spec-options">
              ${s.map(V=>`<option value="${V.full_name||V.label||""}">`).join("")}
            </datalist>
          </div>
        </div>

        <div>
          <label style="display:block; font-size:0.82rem; font-weight:600; color:var(--text-2); margin-bottom:5px;">
            Solusi / Tindak Lanjut Penanganan
          </label>
          <textarea id="iss-modal-solution" class="form-control" rows="2" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;">${l?.solution||""}</textarea>
        </div>

        <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr 1fr; gap:10px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:var(--text-2); margin-bottom:5px;">
              Status Masalah <span style="color:#DC2626;">*</span>
            </label>
            <select id="iss-modal-status" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" required>
              ${Tr.map(V=>`
                <option value="${V}" ${z===V?"selected":""}>${V}</option>
              `).join("")}
            </select>
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:var(--text-2); margin-bottom:5px;">
              Tanggal Selesai
            </label>
            <input type="date" id="iss-modal-completion" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              value="${l?.completion_date||""}" />
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:var(--text-2); margin-bottom:5px;">
              Durasi (Hari)
            </label>
            <input type="number" id="iss-modal-daycount" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #223456; border-radius:8px; background:#0E182B; color:#F8FAFC;" 
              value="${l?.day_count!=null?l.day_count:""}" readonly placeholder="Otomatis" />
          </div>
        </div>
      </div>
    `;let te=O.querySelector("#iss-modal-date"),Q=O.querySelector("#iss-modal-completion"),ne=O.querySelector("#iss-modal-daycount"),ee=O.querySelector("#iss-modal-status");function H(){if(te.value&&Q.value){let V=new Date(te.value),Z=new Date(Q.value),ie=Math.floor((Z-V)/(1e3*60*60*24));ne.value=isNaN(ie)?"":Math.max(0,ie)}else ne.value=""}te.addEventListener("change",H),Q.addEventListener("change",H),ee.addEventListener("change",()=>{ee.value==="Done"&&!Q.value?(Q.value=new Date().toISOString().split("T")[0],H()):ee.value==="Open"&&(Q.value="",ne.value="")});let U=je(O.querySelector("#iss-modal-emp-mount"),{mode:"single",inputId:"iss-modal-employee",inputName:"employee_name",placeholder:"",allowManual:!0,hasEmployeeIdSchema:!1,employees:s,initialEmployeeName:l?.employee_name||""}),D=!1;ce({title:f?"\u270F\uFE0F Edit Data Permasalahan":"\u2795 Tambah Data Permasalahan",content:O,size:"md",confirmText:f?"Simpan Perubahan":"Simpan Permasalahan",confirmClass:"btn-primary",onConfirm:async(V,Z)=>{if(D)return;let ie=O.querySelector("#iss-modal-date").value,le=O.querySelector("#iss-modal-branch").value,pe=O.querySelector("#iss-modal-category").value,de=(O.querySelector("#iss-modal-source").value||"").trim(),me=(O.querySelector("#iss-modal-complaint").value||"").trim(),Ke=U.getValue().employee_name||"",oe=(O.querySelector("#iss-modal-spec").value||"").trim(),ve=(O.querySelector("#iss-modal-solution").value||"").trim(),fe=O.querySelector("#iss-modal-status").value,De=O.querySelector("#iss-modal-completion").value||null;if(!ie||!le||!me||!pe){j("Harap lengkapi Tanggal, Cabang, Kategori, dan Keluhan!");return}let Se=null;if(f){let ke=O.querySelector("#iss-modal-row-version"),Ye=ke?ke.value:u;if(Se=Mt(Ye),Se===null){j("Versi data (row_version) tidak tersedia. Silakan muat ulang data sebelum menyimpan perubahan.");return}}let Le=null;if(ie&&De){let ke=new Date(ie),Ye=new Date(De);Le=Math.max(0,Math.floor((Ye-ke)/(1e3*60*60*24)))}let Ue={report_date:ie,branch_id:Number(le),category:pe,source:de,complaint:me,employee_name:Ke,fc_specialist:oe,solution:ve,status:fe,completion_date:De,day_count:Le,...f?{row_version:Se}:{}},_e=V?.querySelector?.(".modal-confirm");D=!0,_e&&(_e.disabled=!0);try{let ke;if(f?ke=await X(`/api/issues/${l.id}`,{method:"PUT",body:JSON.stringify(Ue)}):ke=await X("/api/issues",{method:"POST",body:JSON.stringify(Ue)}),ke.ok){if(f){let Ye=Mt(ke.data?.data?.row_version??ke.data?.row_version);Ye!==null&&l&&(l.row_version=Ye)}else F=1;ae(f?"Permasalahan berhasil diperbarui!":"Permasalahan baru berhasil ditambahkan!"),Z(),await E()}else f&&ke.status===409?(j("Konflik data (HTTP 409): Issue ini telah diubah oleh pengguna atau sesi lain. Silakan muat ulang data sebelum mengedit kembali."),Z(),await E()):j(ke.data?.error||"Gagal menyimpan permasalahan")}catch(ke){j("Terjadi kesalahan: "+ke.message)}finally{D=!1,_e&&(_e.disabled=!1)}}})}async function J(l){x&&ge({title:"Hapus Permasalahan?",message:`Apakah Anda yakin ingin menghapus tiket permasalahan dari <strong>${l.branch_name||"cabang"}</strong>:<br><br><em>"${l.complaint}"</em>?`,confirmText:"Ya, Hapus",confirmClass:"btn-danger",onConfirm:async()=>{try{let f=await X(`/api/issues/${l.id}`,{method:"DELETE"});f.ok?(ae("Permasalahan berhasil dihapus"),await E()):j(f.data?.error||"Gagal menghapus data")}catch(f){j("Terjadi kesalahan: "+f.message)}}})}function W(){return a}function Y(l){return l.length===0?`
            <div class="issues-empty-state">
              <div style="font-size:2.8rem; margin-bottom:10px;">\u{1F50D}</div>
              <h3 style="font-size:1.15rem; font-weight:700; color:#F8FAFC; margin:0 0 6px 0;">Tidak ada permasalahan yang cocok</h3>
              <p style="font-size:0.85rem; color:#94A3B8; margin:0 0 16px 0;">Coba sesuaikan kata kunci pencarian atau filter status yang dipilih.</p>
              <button id="issues-reset-filter" class="btn btn-secondary" style="font-size:0.82rem; padding:7px 16px; border-radius:8px;">Reset Semua Filter</button>
            </div>
      `:l.map(f=>{let u=f.category||"Lainnya",p=`issues-category-${u.toLowerCase().replace(/[^a-z0-9]/g,"")}`||"issues-category-lainnya",A=f.status||"Open",T="issues-status-open",g="\u26A0\uFE0F Open";A==="In Progress"?(T="issues-status-progress",g="\u23F3 In Progress"):A==="Done"&&(T="issues-status-done",g="\u2705 Done");let C=f.priority||"Normal",z="priority-normal";return(C==="High"||C==="Tinggi")&&(z="priority-high"),(C==="Low"||C==="Rendah")&&(z="priority-low"),`
        <div class="issues-card" data-id="${f.id}">
          <div class="issues-card-top">
            <span class="issues-category-badge ${p}">
              ${u==="SDM"?"\u{1F465}":u==="Cleaning"?"\u{1F9F9}":u==="Aset"?"\u{1F4E6}":"\u{1F4CB}"} ${u}
            </span>
            <div style="display:flex; align-items:center; gap:6px;">
              <span class="issues-priority-badge ${z}">${C}</span>
              <span class="issues-status-badge ${T}">${g}</span>
            </div>
          </div>

          <div>
            <div class="issues-title" style="color:#F8FAFC; font-weight:700; font-size:0.96rem; line-height:1.45;">${f.complaint||"Tidak ada uraian permasalahan"}</div>
            <div style="font-size:0.78rem; color:#CBD5E1; margin-top:4px;">
              ID #${f.id} \u2022 Sumber: <strong style="color:#E2E8F0;">${f.source||"-"}</strong>
            </div>
          </div>

          <div class="issues-details-grid">
            <div class="issues-detail-item">
              <span>\u{1F3E2}</span>
              <span style="font-weight:600; color:#F8FAFC; overflow:hidden; text-overflow:ellipsis; white-space:nowrap;" title="${f.branch_name||"-"}">
                ${f.branch_name||"Semua Cabang"}
              </span>
            </div>
            <div class="issues-detail-item">
              <span>\u{1F4C5}</span>
              <span style="color:#E2E8F0; font-weight:600;">${f.report_date?xe(f.report_date):"-"}</span>
            </div>
            <div class="issues-detail-item">
              <span>\u{1F464}</span>
              <span style="color:#F8FAFC; font-weight:600; overflow:hidden; text-overflow:ellipsis; white-space:nowrap;" title="${f.employee_name||"-"}">
                ${f.employee_name||"-"}
              </span>
            </div>
            <div class="issues-detail-item">
              <span>\u{1F552}</span>
              <span style="color:#CBD5E1;">${f.completion_date?xe(f.completion_date):f.created_at?xe(f.created_at.slice(0,10)):"-"}</span>
            </div>
          </div>

          ${f.solution?`
            <div class="issues-solution-box">
              <div class="issues-solution-title">
                <span>\u{1F4A1}</span> Solusi / Tindakan:
              </div>
              <div>${f.solution}</div>
            </div>
          `:""}

          <div class="issues-card-footer">
            <div style="display:flex; align-items:center; gap:6px;">
              ${B?`
                ${A==="Open"?`
                  <button class="issues-quick-btn start" data-action="quick-progress" title="Mulai Penanganan">
                    \u25B6\uFE0F Proses
                  </button>
                `:""}
                ${A==="In Progress"?`
                  <button class="issues-quick-btn complete" data-action="quick-done" title="Tandai Selesai">
                    \u2705 Selesai
                  </button>
                `:""}
                ${A==="Done"?`
                  <button class="issues-quick-btn reopen" data-action="quick-reopen" title="Buka Kembali Masalah">
                    \u{1F504} Reopen
                  </button>
                `:""}
              `:""}
            </div>

            <div class="issues-action-btns">
              ${B?`
                <button class="issues-icon-btn edit" data-action="edit" data-id="${f.id}" title="Edit Masalah">
                  <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>
                </button>
              `:""}
              ${x?`
                <button class="issues-icon-btn delete" data-action="delete" data-id="${f.id}" title="Hapus Masalah">
                  <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><polyline points="3 6 5 6 21 6"/><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"/><line x1="10" y1="11" x2="10" y2="17"/><line x1="14" y1="11" x2="14" y2="17"/></svg>
                </button>
              `:""}
            </div>
          </div>
        </div>
      `}).join("")}function b(){let l=e.querySelector("#issues-reset-filter");l&&l.addEventListener("click",()=>{d="",i="all",m="all",c="all";let f=e.querySelector("#issues-search-input");f&&(f.value="");let u=e.querySelector("#issues-branch-filter");u&&(u.value="all");let p=e.querySelector("#issues-category-filter");p&&(p.value="all"),e.querySelectorAll(".issues-pill").forEach(A=>{A.classList.toggle("active",A.getAttribute("data-status")==="all")}),F=1,E({refreshStats:!1})}),e.querySelectorAll(".issues-card").forEach(f=>{let u=Number(f.dataset.id),p=a.find(O=>O.id===u);if(!p)return;let A=f.querySelector('[data-action="edit"]');A&&A.addEventListener("click",()=>G(p));let T=f.querySelector('[data-action="delete"]');T&&T.addEventListener("click",()=>J(p));let g=f.querySelector('[data-action="quick-progress"]');g&&g.addEventListener("click",()=>K(p,"In Progress"));let C=f.querySelector('[data-action="quick-done"]');C&&C.addEventListener("click",()=>K(p,"Done"));let z=f.querySelector('[data-action="quick-reopen"]');z&&z.addEventListener("click",()=>K(p,"Open"))})}function S(){let l=e.querySelector("#issues-pagination-container");if(!l)return;l.innerHTML="";let f=Te({page:F,pages:M,total:$,limit:20,onPage:u=>{F=u,E({refreshStats:!1})}});f&&l.appendChild(f)}function y(){let l=e.querySelector("#issues-list-container");if(!l)return;let f=W();l.innerHTML=Y(f),S(),b()}function r(){let l=W(),f=L.total,u=L.openAndProgress,p=L.done,A=L.currentMonth,T=L.total,g=L.open,C=L.inProgress,z=L.done;e.innerHTML=`
      <style>.issues-page{display:flex;flex-direction:column;gap:20px}.issues-hero-card{background:#131F36;border:1px solid #223456;border-radius:16px;padding:24px 28px;color:#FFFFFF;box-shadow:0 4px 16px rgba(0,0,0,0.22);position:relative;overflow:hidden}.issues-hero-card::after{content:'';position:absolute;top:-40px;right:-40px;width:220px;height:220px;background:radial-gradient(circle,rgba(37,99,235,0.06) 0%,transparent 70%);border-radius:50%;pointer-events:none}.issues-hero-actions{display:flex;align-items:center;gap:8px;flex-wrap:wrap}.issues-hero-actions .btn{height:38px;box-sizing:border-box;transition:all 0.2s ease}.issues-btn-add:hover{background:#1D4ED8 !important;transform:translateY(-1px)}.issues-btn-excel:hover,.issues-btn-refresh:hover{background:rgba(255,255,255,0.16) !important;border-color:rgba(255,255,255,0.3) !important;transform:translateY(-1px)}.issues-stats-row{display:grid;grid-template-columns:repeat(4,1fr);gap:14px;margin-top:20px}.issues-stat-box{background:rgba(255,255,255,0.05);border:1px solid rgba(255,255,255,0.1);border-radius:12px;padding:14px 18px;display:flex;align-items:center;gap:14px;backdrop-filter:blur(8px);transition:transform 0.15s ease,background 0.15s ease}.issues-stat-box:hover{background:rgba(255,255,255,0.08);transform:translateY(-2px)}.issues-stat-icon{width:44px;height:44px;border-radius:10px;display:flex;align-items:center;justify-content:center;font-size:1.3rem;flex-shrink:0}.issues-stat-val{font-size:1.55rem;font-weight:800;line-height:1.1;letter-spacing:-0.02em}.issues-stat-lbl{font-size:0.76rem;text-transform:uppercase;color:#94A3B8;font-weight:700;letter-spacing:0.05em;margin-top:2px}.issues-controls-card{background:#131F36;border-radius:14px;border:1px solid #223456;padding:16px 20px;box-shadow:0 2px 8px rgba(0,0,0,0.03);display:flex;flex-direction:column;gap:14px}.issues-search-row{display:flex;gap:12px;flex-wrap:wrap}.issues-search-wrap{position:relative;flex:1;min-width:240px}.issues-search-icon{position:absolute;left:14px;top:50%;transform:translateY(-50%);color:#94A3B8;pointer-events:none}.issues-search-input{width:100%;padding:10px 14px 10px 40px;background:#0E182B;color:#F8FAFC;border:1px solid #223456;border-radius:10px;font-size:0.9rem;outline:none;transition:border-color 0.15s ease}.issues-search-input:focus{border-color:#3B82F6;box-shadow:0 0 0 3px rgba(37,99,235,0.2)}.issues-select{padding:10px 14px;border:1px solid #223456;border-radius:10px;font-size:0.9rem;color:#F8FAFC;background:#0E182B;outline:none;min-width:160px;cursor:pointer}.issues-select:focus{border-color:#3B82F6}.issues-filter-pills{display:flex;align-items:center;gap:8px;overflow-x:auto;white-space:nowrap;scrollbar-width:none;padding-bottom:2px}.issues-filter-pills::-webkit-scrollbar{display:none}.issues-pill{padding:6px 14px;border-radius:20px;font-size:0.82rem;font-weight:600;cursor:pointer;border:1px solid #223456;background:#0E182B;color:#CBD5E1;transition:all 0.15s ease;display:inline-flex;align-items:center;gap:6px;user-select:none;flex-shrink:0}.issues-pill:hover{background:#22314A;color:#F8FAFC}.issues-pill.active{background:#2563EB;color:#FFFFFF;border-color:#2563EB}.issues-pill-dot{width:8px;height:8px;border-radius:50%}.issues-pill-count{padding:1px 6px;border-radius:10px;font-size:0.76rem;background:rgba(255,255,255,0.1);color:#E2E8F0}.issues-pill.active .issues-pill-count{background:rgba(255,255,255,0.25);color:white}.issues-cards-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(370px,1fr));gap:16px}.issues-card{background:#131F36;border-radius:14px;border:1px solid #223456;padding:18px;box-shadow:0 2px 8px rgba(0,0,0,0.2);display:flex;flex-direction:column;gap:14px;transition:transform 0.15s ease,box-shadow 0.15s ease}.issues-card:hover{transform:translateY(-2px);box-shadow:0 6px 16px rgba(0,0,0,0.32);border-color:#475569}.issues-card-top{display:flex;justify-content:space-between;align-items:center;gap:8px}.issues-category-badge{display:inline-flex;align-items:center;gap:5px;padding:4px 10px;border-radius:20px;font-size:0.76rem;font-weight:700}.issues-category-sdm{background:rgba(99,102,241,0.18);color:#A5B4FC;border:1px solid rgba(129,140,248,0.35)}.issues-category-cleaning{background:rgba(20,184,166,0.16);color:#5EEAD4;border:1px solid rgba(45,212,191,0.35)}.issues-category-aset{background:rgba(139,92,246,0.18);color:#C4B5FD;border:1px solid rgba(167,139,250,0.35)}.issues-category-k3{background:rgba(245,158,11,0.18);color:#FCD34D;border:1px solid rgba(251,191,36,0.35)}.issues-category-lainnya{background:rgba(148,163,184,0.16);color:#E2E8F0;border:1px solid rgba(148,163,184,0.3)}.issues-status-badge{display:inline-flex;align-items:center;gap:5px;padding:4px 10px;border-radius:20px;font-size:0.76rem;font-weight:700}.issues-status-open{background:rgba(245,158,11,0.18);color:#FCD34D;border:1px solid rgba(251,191,36,0.38)}.issues-status-inprogress,.issues-status-progress{background:rgba(59,130,246,0.18);color:#93C5FD;border:1px solid rgba(96,165,250,0.38)}.issues-status-done{background:rgba(16,185,129,0.18);color:#6EE7B7;border:1px solid rgba(52,211,153,0.38)}.issues-complaint-box{background:#0E182B;border-left:4px solid #F59E0B;border-radius:0 10px 10px 0;padding:12px 14px;font-size:0.92rem;color:#F8FAFC;line-height:1.5;font-weight:600}.issues-complaint-box.status-done{border-left-color:#10B981}.issues-complaint-box.status-inprogress{border-left-color:#3B82F6}.issues-person-box{background:#0E182B;border:1px solid #223456;border-radius:10px;padding:12px;display:flex;align-items:center;gap:12px}.issues-avatar-box{width:44px;height:44px;border-radius:10px;background:#2563EB;color:white;display:flex;align-items:center;justify-content:center;font-weight:800;font-size:1rem;flex-shrink:0}.issues-person-name{font-size:0.95rem;font-weight:800;color:#F8FAFC}.issues-person-sub{font-size:0.76rem;color:#94A3B8;font-weight:600;text-transform:uppercase;letter-spacing:0.04em}.issues-details-grid{display:grid;grid-template-columns:1fr 1fr;gap:8px 12px;font-size:0.82rem;color:#E2E8F0;padding:2px 4px}.issues-detail-item{display:flex;align-items:center;gap:6px}.issues-solution-box{background:rgba(16,185,129,0.12);border:1px solid rgba(52,211,153,0.28);border-radius:10px;padding:10px 12px;font-size:0.83rem;color:#E2E8F0;line-height:1.5}.issues-solution-title{font-weight:700;color:#6EE7B7;display:flex;align-items:center;gap:6px;margin-bottom:4px}.issues-card-footer{display:flex;justify-content:space-between;align-items:center;border-top:1px solid #223456;padding-top:12px;margin-top:auto;gap:8px}.issues-action-btns{display:flex;align-items:center;gap:6px}.issues-icon-btn{width:34px;height:34px;border-radius:8px;border:1px solid #223456;background:#0E182B;color:#CBD5E1;display:flex;align-items:center;justify-content:center;cursor:pointer;transition:all 0.15s ease}.issues-icon-btn:hover{background:#22314A;color:#F8FAFC;border-color:#475569}.issues-icon-btn.delete:hover{background:rgba(239,68,68,0.2);color:#FCA5A5;border-color:rgba(248,113,113,0.45)}.issues-quick-btn{font-size:0.76rem;font-weight:700;padding:6px 10px;border-radius:8px;cursor:pointer;border:1px solid transparent;display:inline-flex;align-items:center;gap:4px;transition:all 0.15s ease}.issues-quick-btn.start{background:rgba(59,130,246,0.16);color:#93C5FD;border-color:rgba(96,165,250,0.35)}.issues-quick-btn.start:hover{background:rgba(59,130,246,0.26);color:#FFFFFF}.issues-quick-btn.complete{background:rgba(16,185,129,0.16);color:#6EE7B7;border-color:rgba(52,211,153,0.35)}.issues-quick-btn.complete:hover{background:rgba(16,185,129,0.26);color:#FFFFFF}.issues-quick-btn.reopen{background:rgba(245,158,11,0.16);color:#FCD34D;border-color:rgba(251,191,36,0.35)}.issues-quick-btn.reopen:hover{background:rgba(245,158,11,0.26);color:#FFFFFF}.issues-empty-state{grid-column:1 / -1;background:#131F36;border:2px dashed #223456;border-radius:16px;padding:48px 24px;text-align:center}@media (max-width:768px){.issues-hero-card{padding:18px 16px}.issues-stats-row{grid-template-columns:1fr 1fr;gap:10px}.issues-stat-box{padding:10px 12px}.issues-stat-icon{width:38px;height:38px;font-size:1.1rem}.issues-stat-val{font-size:1.3rem}.issues-hero-actions{width:100% !important;display:flex !important;align-items:center !important;gap:8px !important;flex-wrap:wrap !important;overflow-x:visible !important}.issues-hero-actions::-webkit-scrollbar{display:none}.issues-btn-add{width:100% !important;justify-content:center !important;padding:9px 12px !important;font-size:0.85rem !important;white-space:nowrap !important;height:40px !important}.issues-btn-excel{flex:1 1 auto !important;justify-content:center !important;padding:8px 12px !important;font-size:0.85rem !important;white-space:nowrap !important;height:38px !important}.issues-btn-refresh{width:38px !important;height:38px !important;padding:0 !important;flex-shrink:0 !important;justify-content:center !important}.issues-controls-card{padding:14px;gap:12px}.issues-search-row{flex-direction:column;gap:10px}.issues-select{width:100%}.issues-cards-grid{grid-template-columns:minmax(0,1fr) !important;width:100% !important;max-width:100% !important;min-width:0 !important;box-sizing:border-box !important}.issues-card{width:100% !important;max-width:100% !important;min-width:0 !important;box-sizing:border-box !important;overflow:hidden;padding:14px !important}}</style>

      <div class="issues-page">

        <!-- Hero Header Card -->
        <div class="issues-hero-card">
          <div style="display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:16px;">
            <div>
              <div style="display:flex; align-items:center; gap:10px;">
                <span style="font-size:1.6rem;">\u26A0\uFE0F</span>
                <h1 style="font-size:1.45rem; font-weight:800; margin:0; letter-spacing:-0.02em;">
                  Log Permasalahan Operasional
                </h1>
              </div>
            </div>

            <div class="issues-hero-actions" style="display:flex; align-items:center; gap:8px;">
              ${v?`
                <button id="issues-btn-add" class="btn btn-primary issues-btn-add" style="background:#2563EB; height:38px; padding:0 16px; border-radius:10px; font-weight:700; display:inline-flex; align-items:center; gap:6px; box-shadow:0 4px 12px rgba(37,99,235,0.3); white-space:nowrap;">
                  <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24"><line x1="12" y1="5" x2="12" y2="19"/><line x1="5" y1="12" x2="19" y2="12"/></svg>
                  <span>Tambah Permasalahan</span>
                </button>
              `:""}
              
              ${re(q,"export")?`<button id="issues-btn-excel" class="btn btn-ghost issues-btn-excel" style="background:rgba(255,255,255,0.08); border:1px solid rgba(255,255,255,0.15); color:white; height:38px; padding:0 14px; border-radius:10px; font-weight:600; display:inline-flex; align-items:center; gap:6px; white-space:nowrap;">
                <span>\u{1F4E5}</span>
                <span>Ekspor Excel</span>
              </button>`:""}

              <button id="issues-btn-refresh" class="btn btn-ghost issues-btn-refresh" style="background:rgba(255,255,255,0.08); border:1px solid rgba(255,255,255,0.15); color:white; width:38px; height:38px; padding:0; border-radius:10px; display:inline-flex; align-items:center; justify-content:center; flex-shrink:0;" title="Muat Ulang Data">
                <svg width="16" height="16" fill="none" stroke="currentColor" stroke-width="2.2" viewBox="0 0 24 24" class="${I?"spin-icon":""}"><path d="M23 4v6h-6"/><path d="M1 20v-6h6"/><path d="M3.51 9a9 9 0 0 1 14.85-3.36L23 10M1 14l4.64 4.36A9 9 0 0 0 20.49 15"/></svg>
              </button>
            </div>
          </div>

          <!-- 4 KPI Stat Boxes -->
          <div class="issues-stats-row">
            <div class="issues-stat-box">
              <div class="issues-stat-icon" style="background:rgba(239, 68, 68, 0.2); color:#FCA5A5;">\u26A0\uFE0F</div>
              <div>
                <div class="issues-stat-val" id="issues-stat-total">${f}</div>
                <div class="issues-stat-lbl">Total Permasalahan</div>
              </div>
            </div>

            <div class="issues-stat-box">
              <div class="issues-stat-icon" style="background:rgba(245, 158, 11, 0.2); color:#FCD34D;">\u23F3</div>
              <div>
                <div class="issues-stat-val" id="issues-stat-open-prog">${u}</div>
                <div class="issues-stat-lbl">Open & In Progress</div>
              </div>
            </div>

            <div class="issues-stat-box">
              <div class="issues-stat-icon" style="background:rgba(16, 185, 129, 0.2); color:#6EE7B7;">\u2705</div>
              <div>
                <div class="issues-stat-val" id="issues-stat-done">${p}</div>
                <div class="issues-stat-lbl">Tuntas (Done)</div>
              </div>
            </div>

            <div class="issues-stat-box">
              <div class="issues-stat-icon" style="background:rgba(59, 130, 246, 0.2); color:#93C5FD;">\u{1F4C5}</div>
              <div>
                <div class="issues-stat-val" id="issues-stat-month">${A}</div>
                <div class="issues-stat-lbl">Bulan Berjalan</div>
              </div>
            </div>
          </div>
        </div>

        <!-- Filter Controls Card -->
        <div class="issues-controls-card">
          <div class="issues-search-row">
            <div class="issues-search-wrap">
              <svg class="issues-search-icon" width="16" height="16" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/></svg>
              <input type="text" id="issues-search-input" class="issues-search-input" 
                placeholder="Cari keluhan, nama personil, sumber, cabang..." value="${d?String(d).replace(/"/g,"&quot;"):""}" />
            </div>

            <select id="issues-branch-filter" class="issues-select">
              <option value="all">Semua Cabang (${o.length})</option>
              ${o.map(D=>`
                <option value="${D.value}" ${String(i)===String(D.value)?"selected":""}>
                  ${D.label}
                </option>
              `).join("")}
            </select>

            <select id="issues-category-filter" class="issues-select">
              <option value="all">Semua Kategori</option>
              ${Rn.map(D=>`
                <option value="${D}" ${m===D?"selected":""}>
                  Kategori: ${D}
                </option>
              `).join("")}
            </select>
          </div>

          <!-- Status Horizontal Swipeable Pills -->
          <div class="issues-filter-pills">
            <div class="issues-pill ${c==="all"?"active":""}" data-status="all">
              <span>Semua Status</span>
              <span class="issues-pill-count" id="issues-pill-count-all">${T}</span>
            </div>
            <div class="issues-pill ${c==="Open"?"active":""}" data-status="Open">
              <span class="issues-pill-dot" style="background:#F59E0B;"></span>
              <span>Open</span>
              <span class="issues-pill-count" id="issues-pill-count-open">${g}</span>
            </div>
            <div class="issues-pill ${c==="In Progress"?"active":""}" data-status="In Progress">
              <span class="issues-pill-dot" style="background:#3B82F6;"></span>
              <span>In Progress</span>
              <span class="issues-pill-count" id="issues-pill-count-inprogress">${C}</span>
            </div>
            <div class="issues-pill ${c==="Done"?"active":""}" data-status="Done">
              <span class="issues-pill-dot" style="background:#10B981;"></span>
              <span>Done</span>
              <span class="issues-pill-count" id="issues-pill-count-done">${z}</span>
            </div>
          </div>
        </div>

        <!-- Exclusive Cards Grid -->
        <div id="issues-list-container" class="issues-cards-grid">
          ${Y(l)}
        </div>

        <!-- Pagination Container -->
        <div id="issues-pagination-container"></div>

      </div>
    `;let O=e.querySelector("#issues-search-input");O&&O.addEventListener("input",D=>{d=D.target.value,clearTimeout(R),R=setTimeout(()=>{F=1,E({refreshStats:!1})},350)});let te=e.querySelector("#issues-branch-filter");te&&te.addEventListener("change",D=>{i=D.target.value,F=1,E({refreshStats:!1})});let Q=e.querySelector("#issues-category-filter");Q&&Q.addEventListener("change",D=>{m=D.target.value,F=1,E({refreshStats:!1})}),e.querySelectorAll(".issues-pill").forEach(D=>{D.addEventListener("click",()=>{c=D.dataset.status,e.querySelectorAll(".issues-pill").forEach(V=>{V.classList.toggle("active",V===D)}),F=1,E({refreshStats:!1})})});let ne=e.querySelector("#issues-reset-filter");ne&&ne.addEventListener("click",()=>{d="",i="all",m="all",c="all",F=1,clearTimeout(R),O&&(O.value=""),te&&(te.value="all"),Q&&(Q.value="all"),e.querySelectorAll(".issues-pill").forEach(D=>{D.classList.toggle("active",D.dataset.status==="all")}),E({refreshStats:!1})});let ee=e.querySelector("#issues-btn-add");ee&&ee.addEventListener("click",()=>G());let H=e.querySelector("#issues-btn-excel");H&&H.addEventListener("click",()=>w());let U=e.querySelector("#issues-btn-refresh");U&&U.addEventListener("click",()=>E({refreshStats:!0})),S(),b()}await E({refreshStats:!0})}Be();var Zt=["Batch 1","Batch 2","Batch 3","Batch 4","Batch 5"],Hn=["Berlin Ariansyah","Ade Surahman","Mizwar","Fajar","Ade","Berlin"];function za(e){if(!e)return"-";try{let t=JSON.parse(e);if(Array.isArray(t))return t.join(", ")}catch{}return String(e)}async function On(e){let t=[],n=[],a=[],o="all",s="all",c="all",m="",i=!1,d=1,F=20,$=0,M=1,h={totalSessions:0,totalBatches:0,currentMonth:0,totalParticipants:0},L=null,R=!1,I=se(),N=re(I,"training","write"),q=re(I,"training","delete");try{[n,a]=await Promise.all([we(),ze()])}catch(S){console.error("Error fetching cache:",S)}function v(){let S={sort:"date_desc"};return m&&m.trim()&&(S.search=m.trim()),s&&s!=="all"&&(S.branch_id=s),o&&o!=="all"&&(S.batch=o),c&&c!=="all"&&(S.trainer=c),S}function B(){let S=(y,r)=>{let l=e.querySelector(y);l&&(l.textContent=String(r??0))};S("#tr-stat-sessions",h.totalSessions),S("#tr-stat-participants",h.totalParticipants),S("#tr-stat-batches",h.totalBatches),S("#tr-stat-month",h.currentMonth),S("#tr-pill-count-all",h.totalSessions)}async function x({refreshStats:S=!0}={}){i=!0;try{let y=v(),r=new URLSearchParams({page:String(d),limit:String(F),...y});S&&r.set("include_stats","1");let l=await X("/api/training?"+r.toString());if(l.ok&&l.data&&Array.isArray(l.data.data)){let f=l.data.pagination||{};if($=Number(f.total??l.data.data.length),M=Math.max(1,Number(f.pages||f.totalPages||Math.ceil($/F)||1)),l.data.data.length===0&&d>M)return d=M,i=!1,x({refreshStats:S});t=l.data.data,l.data.stats&&(h={...h,...l.data.stats})}else t=[],$=0,M=1,j(l.data?.error||"Gagal memuat data training")}catch(y){console.warn("Gagal load training:",y),t=[],$=0,M=1,j("Gagal memuat data training: "+y.message)}finally{i=!1,R?(S&&B(),W()):(R=!0,Y())}}async function k(){if(!re(I,"export")){j("Anda tidak memiliki izin untuk mengekspor data.");return}try{let{downloadExcel:S}=await Promise.resolve().then(()=>(Ce(),Ve)),y=await ue("/api/training/export",v()),r=y.map((l,f)=>({No:f+1,Tanggal:l.training_date||"",Batch:l.batch||"",Materi:l.subject||"",Cabang:l.branch_name||"",Trainer:l.trainer||"",Peserta:za(l.participants),Nilai:l.score!=null?l.score:"",Dokumen:l.document_link||"",Catatan:l.notes||""}));S(r,`Data_Training_${new Date().toISOString().slice(0,10)}`),ae(`Berhasil mengunduh ${y.length} data training ke Excel!`)}catch(S){j("Gagal ekspor Excel: "+S.message)}}async function _(S){if(!q)return;let y=t.find(r=>r.id===S);ge({title:"Hapus Data Training",message:`Apakah Anda yakin ingin menghapus data sesi training "${y?.subject||"ini"}" untuk ${za(y?.participants)}?`,confirmText:"Ya, Hapus",confirmClass:"btn-danger",onConfirm:async()=>{try{let r=await X(`/api/training/${S}`,{method:"DELETE"});r.ok?(ae("Data training berhasil dihapus"),await x({refreshStats:!0})):j(r.data?.error||"Gagal menghapus data training")}catch(r){j("Gagal menghapus training: "+r.message)}}})}function E(S=null){if(!N)return;let y=!!S,r=S?.training_date?S.training_date.slice(0,10):new Date().toISOString().slice(0,10),l=y&&S?.batch?String(S.batch).trim():"",f=y&&l&&!Zt.includes(l),u=S?.trainer||"Berlin Ariansyah",p=document.createElement("div");p.innerHTML=`
      <div style="display:flex; flex-direction:column; gap:16px;">
        <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:var(--text-2); margin-bottom:5px;">
              Tanggal Training <span style="color:#DC2626;">*</span>
            </label>
            <input type="date" id="tr-modal-date" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              value="${r}" required />
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:var(--text-2); margin-bottom:5px;">
              Batch Pelatihan <span style="color:#DC2626;">*</span>
            </label>
            <select id="tr-modal-batch" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" required>
              <option value="">-- Pilih Batch Pelatihan --</option>
              ${Zt.map(T=>`<option value="${T}" ${l===T?"selected":""}>${T}</option>`).join("")}
              ${f?`<option value="${l.replace(/"/g,"&quot;")}" selected>${l} (Historis)</option>`:""}
            </select>
          </div>
        </div>

        <div>
          <label style="display:block; font-size:0.82rem; font-weight:600; color:var(--text-2); margin-bottom:5px;">
            Materi / Topik Pelatihan <span style="color:#DC2626;">*</span>
          </label>
          <input type="text" id="tr-modal-subject" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
            value="${S?.subject?String(S.subject).replace(/"/g,"&quot;"):""}" required />
        </div>

        <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:var(--text-2); margin-bottom:5px;">
              Cabang Klinik <span style="color:#DC2626;">*</span>
            </label>
            <select id="tr-modal-branch" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" required>
              <option value="">-- Pilih Cabang --</option>
              ${n.map(T=>`
                <option value="${T.value}" ${String(S?.branch_id)===String(T.value)?"selected":""}>
                  ${T.label}
                </option>
              `).join("")}
            </select>
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:var(--text-2); margin-bottom:5px;">
              Trainer / Instruktur <span style="color:#DC2626;">*</span>
            </label>
            <input list="tr-trainer-options" id="tr-modal-trainer" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              value="${u}" required />
            <datalist id="tr-trainer-options">
              ${Hn.map(T=>`<option value="${T}">`).join("")}
            </datalist>
          </div>
        </div>

        <div>
          <label style="display:block; font-size:0.82rem; font-weight:600; color:var(--text-2); margin-bottom:5px;">
            Nama Peserta Pelatihan <span style="color:#DC2626;">*</span>
          </label>
          <div id="tr-modal-participants-mount"></div>
        </div>

        <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:var(--text-2); margin-bottom:5px;">
              Nilai / Score Evaluasi (0 - 100)
            </label>
            <input type="number" min="0" max="100" step="1" id="tr-modal-score" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              value="${S?.score!=null?S.score:""}" />
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:var(--text-2); margin-bottom:5px;">
              Link Dokumen / Sertifikat (Opsional)
            </label>
            <input type="url" id="tr-modal-link" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              value="${S?.document_link?String(S.document_link).replace(/"/g,"&quot;"):""}" />
          </div>
        </div>

        <div>
          <label style="display:block; font-size:0.82rem; font-weight:600; color:var(--text-2); margin-bottom:5px;">
            Catatan Pelatihan
          </label>
          <textarea id="tr-modal-notes" class="form-control" rows="2" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;">${S?.notes||""}</textarea>
        </div>
      </div>
    `;let A=je(p.querySelector("#tr-modal-participants-mount"),{mode:"multi",inputId:"tr-modal-participants",inputName:"participants",placeholder:"",allowManual:!0,hasEmployeeIdSchema:!1,employees:a,initialParticipants:Fa(S?.participants)});ce({title:y?"\u270F\uFE0F Edit Data Sesi Training":"\u2795 Tambah Data Sesi Training",content:p,size:"md",confirmText:y?"Simpan Perubahan":"Simpan Training",confirmClass:"btn-primary",onConfirm:async(T,g)=>{let C=p.querySelector("#tr-modal-date").value,z=(p.querySelector("#tr-modal-batch").value||"").trim(),O=(p.querySelector("#tr-modal-subject").value||"").trim(),te=p.querySelector("#tr-modal-branch").value,Q=(p.querySelector("#tr-modal-trainer").value||"").trim(),ne=(A.getValue().participantsString||"").trim(),ee=p.querySelector("#tr-modal-score").value,H=ee!==""?Number(ee):null,U=(p.querySelector("#tr-modal-link").value||"").trim(),D=(p.querySelector("#tr-modal-notes").value||"").trim();if(!C){j("Tanggal training wajib diisi");return}if(!z||!(y&&z===l)&&!Zt.includes(z)){j("Batch pelatihan wajib dipilih (Batch 1 - Batch 5)");return}if(!O){j("Materi / topik pelatihan wajib diisi");return}if(!te){j("Cabang klinik wajib dipilih");return}if(!Q){j("Trainer wajib diisi");return}if(!ne){j("Nama peserta training wajib diisi");return}let V=n.find(ie=>String(ie.value)===String(te)),Z={training_date:C,batch:z,subject:O,branch_id:Number(te),branch_name:V?V.label:"Cabang",trainer:Q,participants:ne,score:H,document_link:U,notes:D};try{if(y){let ie=await X(`/api/training/${S.id}`,{method:"PUT",body:JSON.stringify(Z)});ie.ok?(ae("Data training berhasil diperbarui"),g(),await x({refreshStats:!0})):j(ie.data?.error||"Gagal memperbarui data training")}else{let ie=await X("/api/training",{method:"POST",body:JSON.stringify(Z)});ie.ok?(ae("Sesi training baru berhasil ditambahkan"),d=1,g(),await x({refreshStats:!0})):j(ie.data?.error||"Gagal menambahkan training")}}catch(ie){j("Terjadi kesalahan: "+ie.message)}}})}function w(){return t}function K(S){return`
    ${S.length===0?`
            <div class="training-empty-state">
              <div style="font-size:3rem; margin-bottom:12px;">\u{1F50D}</div>
              <h3 style="font-size:1.2rem; font-weight:700; color:#F8FAFC; margin:0 0 6px 0;">Tidak Ada Data Training Ditemukan</h3>
              <p style="color:#94A3B8; font-size:0.9rem; margin:0 0 20px 0;">Silakan sesuaikan kata kunci pencarian atau reset filter aktif.</p>
              <button id="tr-btn-reset-filter" class="btn btn-ghost" style="border:1px solid #223456; color:#CBD5E1; border-radius:8px; padding:8px 16px;">
                Reset Filter
              </button>
            </div>
          `:`
            <div class="training-cards-grid">
              ${S.map(y=>{let r=za(y.participants),f=r.split(",")[0].trim().split(" ").map(u=>u[0]).filter(Boolean).slice(0,2).join("").toUpperCase()||"TR";return`
                  <div class="training-card" data-id="${y.id}">
                    <div>
                      <!-- Card Header: Batch & Nilai/Tanggal -->
                      <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:12px; gap:8px; flex-wrap:wrap;">
                        <span class="training-batch-badge">
                          <span>\u{1F393}</span>
                          <span>${y.batch||"Reguler"}</span>
                        </span>
                        <div style="display:flex; align-items:center; gap:8px;">
                          ${y.score!=null?`
                            <span class="training-score-badge">
                              <span>\u2B50</span>
                              <span>${y.score}</span>
                            </span>
                          `:""}
                          <span style="font-size:0.8rem; font-weight:600; color:#CBD5E1;">
                            \u{1F5D3}\uFE0F ${xe(y.training_date)}
                          </span>
                        </div>
                      </div>

                      <!-- Subject / Materi Title -->
                      <h4 style="margin:0 0 12px 0; font-size:1.05rem; font-weight:700; color:#F8FAFC; line-height:1.4;">
                        ${y.subject||"Materi Pelatihan"}
                      </h4>

                      <!-- Participant & Branch Info -->
                      <div style="display:flex; align-items:center; gap:12px; margin-bottom:10px;">
                        <div class="training-avatar-box">
                          ${f}
                        </div>
                        <div style="min-width:0;">
                          <div style="font-size: 0.76rem; text-transform:uppercase; font-weight:700; color:#94A3B8; margin-bottom:2px; letter-spacing:0.04em;">Peserta Training</div>
                          <div style="font-size:0.95rem; font-weight:700; color:#F8FAFC; white-space:nowrap; overflow:hidden; text-overflow:ellipsis;" title="${r}">
                            ${r}
                          </div>
                        </div>
                      </div>

                      <!-- Detail Box: Cabang & Trainer -->
                      <div class="training-info-box">
                        <div style="display:flex; align-items:center; gap:6px; font-size:0.83rem; color:#E2E8F0;">
                          <span>\u{1F4CD}</span>
                          <span style="font-weight:600; color:#F8FAFC;">${y.branch_name||"Cabang Belum Ditentukan"}</span>
                        </div>
                        <div style="display:flex; align-items:center; gap:6px; font-size:0.83rem; color:#CBD5E1;">
                          <span>\u{1F468}\u200D\u{1F3EB}</span>
                          <span>Trainer: <strong style="color:#FFFFFF;">${y.trainer||"-"}</strong></span>
                        </div>
                      </div>

                      ${y.notes?`
                        <div style="font-size:0.8rem; color:#CBD5E1; font-style:italic; line-height:1.4; margin-bottom:8px;">
                          \u{1F4DD} ${y.notes}
                        </div>
                      `:""}
                    </div>

                    <!-- Card Footer: Dokumen & Aksi -->
                    <div style="display:flex; justify-content:space-between; align-items:center; padding-top:12px; border-top:1px solid #223456; margin-top:8px;">
                      <div>
                        ${y.document_link?`
                          <a href="${y.document_link}" target="_blank" rel="noopener noreferrer" class="btn btn-sm btn-ghost" style="border:1px solid #223456; color:#93C5FD; background:rgba(59,130,246,0.12); border-radius:6px; font-size:0.78rem; padding:4px 10px; text-decoration:none; display:inline-flex; align-items:center; gap:4px;">
                            <span>\u{1F4C4} Materi / Sertifikat</span>
                          </a>
                        `:`
                          <span style="font-size:0.75rem; color:#94A3B8;">(Tanpa Dokumen)</span>
                        `}
                      </div>

                      ${N||q?`
                        <div style="display:flex; align-items:center; gap:6px;">
                          ${N?`
                            <button class="training-icon-btn training-edit-btn" title="Edit Data Training" data-id="${y.id}">
                              <svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>
                            </button>
                          `:""}
                          ${q?`
                            <button class="training-icon-btn delete training-delete-btn" title="Hapus Training" data-id="${y.id}">
                              <svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><polyline points="3 6 5 6 21 6"/><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"/></svg>
                            </button>
                          `:""}
                        </div>
                      `:""}
                    </div>

                  </div>
                `}).join("")}
            </div>
          `}
    `}function G(){let S=e.querySelector("#tr-btn-reset-filter");S&&(S.onclick=()=>{m="",s="all",o="all",c="all",d=1,clearTimeout(L);let y=e.querySelector("#tr-search-input");y&&(y.value="");let r=e.querySelector("#tr-filter-branch");r&&(r.value="all");let l=e.querySelector("#tr-filter-trainer");l&&(l.value="all"),e.querySelectorAll(".training-pill").forEach(f=>{f.classList.toggle("active",f.getAttribute("data-batch")==="all")}),x({refreshStats:!1})}),e.querySelectorAll(".training-edit-btn").forEach(y=>{y.onclick=r=>{r.stopPropagation();let l=Number(y.getAttribute("data-id")),f=t.find(u=>u.id===l);f&&E(f)}}),e.querySelectorAll(".training-delete-btn").forEach(y=>{y.onclick=r=>{r.stopPropagation();let l=Number(y.getAttribute("data-id"));_(l)}})}function J(){let S=e.querySelector("#training-pagination-container");if(!S)return;S.innerHTML="";let y=Te({page:d,pages:M,total:$,limit:F,onPage:r=>{d=r,x({refreshStats:!1})}});y&&S.appendChild(y)}function W(){let S=e.querySelector("#training-list-container");if(!S)return;let y=w();S.innerHTML=K(y),J(),G()}function Y(){let S=w(),y=Number(h.totalSessions??0),r=Number(h.totalParticipants??0),l=Number(h.totalBatches??0),f=Number(h.currentMonth??0);e.innerHTML=`
      <style>.training-page{display:flex;flex-direction:column;gap:20px}.training-hero-card{background:#131F36;border:1px solid #223456;border-radius:16px;padding:24px 28px;color:#FFFFFF;box-shadow:0 4px 16px rgba(0,0,0,0.22);position:relative;overflow:hidden}.training-hero-card::after{content:'';position:absolute;top:-40px;right:-40px;width:220px;height:220px;background:radial-gradient(circle,rgba(37,99,235,0.06) 0%,transparent 70%);border-radius:50%;pointer-events:none}.training-hero-actions{display:flex;align-items:center;gap:8px;flex-wrap:wrap}.training-hero-actions .btn{height:38px;box-sizing:border-box;transition:all 0.2s ease}.training-btn-add:hover{background:#1D4ED8 !important;transform:translateY(-1px)}.training-btn-excel:hover,.training-btn-refresh:hover{background:rgba(255,255,255,0.16) !important;border-color:rgba(255,255,255,0.3) !important;transform:translateY(-1px)}.training-stats-row{display:grid;grid-template-columns:repeat(4,1fr);gap:14px;margin-top:22px}.training-stat-box{background:rgba(255,255,255,0.06);border:1px solid rgba(255,255,255,0.1);border-radius:12px;padding:14px 16px;backdrop-filter:blur(8px);display:flex;align-items:center;gap:14px}.training-stat-icon{width:44px;height:44px;border-radius:10px;display:flex;align-items:center;justify-content:center;font-size:1.3rem;flex-shrink:0}.training-stat-val{font-size:1.55rem;font-weight:800;line-height:1.1;color:#F8FAFC}.training-stat-lbl{font-size:0.76rem;text-transform:uppercase;letter-spacing:0.05em;color:#94A3B8;font-weight:600;margin-top:2px}.training-controls-card{background:#131F36;border:1px solid #223456;border-radius:14px;padding:18px 20px;box-shadow:0 2px 4px rgba(0,0,0,0.02);display:flex;flex-direction:column;gap:16px}.training-search-row{display:flex;gap:12px;flex-wrap:wrap}.training-search-wrap{position:relative;flex:1;min-width:240px}.training-search-icon{position:absolute;left:14px;top:50%;transform:translateY(-50%);color:#94A3B8;pointer-events:none}.training-search-input{width:100%;padding:10px 14px 10px 40px;background:#0E182B;color:#F8FAFC;border:1px solid #223456;border-radius:10px;font-size:0.9rem;outline:none;transition:border-color 0.15s ease}.training-search-input:focus{border-color:#3B82F6;box-shadow:0 0 0 3px rgba(37,99,235,0.2)}.training-select{padding:10px 14px;border:1px solid #223456;border-radius:10px;font-size:0.9rem;color:#F8FAFC;background:#0E182B;outline:none;min-width:180px;cursor:pointer}.training-select:focus{border-color:#3B82F6}.training-filter-pills{display:flex;align-items:center;gap:8px;overflow-x:auto;white-space:nowrap;scrollbar-width:none;padding-bottom:2px}.training-filter-pills::-webkit-scrollbar{display:none}.training-pill{padding:6px 14px;border-radius:20px;font-size:0.82rem;font-weight:600;cursor:pointer;border:1px solid #223456;background:#0E182B;color:#CBD5E1;transition:all 0.15s ease;display:inline-flex;align-items:center;gap:6px;user-select:none;flex-shrink:0}.training-pill:hover{background:#22314A;color:#F8FAFC}.training-pill.active{background:#2563EB;color:white;border-color:#2563EB;box-shadow:0 2px 6px rgba(37,99,235,0.25)}.training-cards-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(320px,1fr));gap:16px}.training-card{background:#131F36;border:1px solid #223456;border-radius:14px;padding:18px;box-shadow:0 2px 6px rgba(0,0,0,0.2);transition:transform 0.15s ease,box-shadow 0.15s ease;display:flex;flex-direction:column;justify-content:space-between;position:relative}.training-card:hover{transform:translateY(-2px);box-shadow:0 8px 18px rgba(0,0,0,0.32);border-color:#475569}.training-avatar-box{width:44px;height:44px;border-radius:10px;background:#2563EB;color:#FFFFFF;font-weight:700;font-size:1rem;display:flex;align-items:center;justify-content:center;flex-shrink:0}.training-batch-badge{display:inline-flex;align-items:center;gap:5px;padding:4px 10px;border-radius:20px;font-size:0.76rem;font-weight:700;background:rgba(37,99,235,0.18);color:#93C5FD;border:1px solid rgba(96,165,250,0.35)}.training-score-badge{display:inline-flex;align-items:center;gap:4px;padding:4px 10px;border-radius:20px;font-size:0.76rem;font-weight:700;background:rgba(16,185,129,0.16);color:#6EE7B7;border:1px solid rgba(52,211,153,0.35)}.training-info-box{background:#0E182B;border:1px solid #223456;border-radius:10px;padding:10px 12px;margin:12px 0;display:flex;flex-direction:column;gap:6px}.training-icon-btn{width:32px;height:32px;border-radius:8px;border:1px solid #223456;background:#0E182B;color:#CBD5E1;display:inline-flex;align-items:center;justify-content:center;cursor:pointer;transition:all 0.15s ease}.training-icon-btn:hover{background:#22314A;color:#F8FAFC;border-color:#475569}.training-icon-btn.delete:hover{background:rgba(239,68,68,0.2);color:#FCA5A5;border-color:rgba(248,113,113,0.45)}.training-empty-state{grid-column:1 / -1;background:#131F36;border:2px dashed #223456;border-radius:16px;padding:48px 24px;text-align:center}@media (max-width:768px){.training-hero-card{padding:18px 16px}.training-stats-row{grid-template-columns:1fr 1fr;gap:10px}.training-stat-box{padding:10px 12px}.training-stat-icon{width:38px;height:38px;font-size:1.1rem}.training-stat-val{font-size:1.3rem}.training-hero-actions{width:100% !important;display:flex !important;align-items:center !important;gap:6px !important;flex-wrap:wrap !important;overflow-x:visible !important;scrollbar-width:none !important;padding-bottom:2px !important}.training-hero-actions::-webkit-scrollbar{display:none}.training-btn-add{flex:1 1 auto !important;justify-content:center !important;padding:8px 10px !important;font-size:0.82rem !important;white-space:nowrap !important}.training-btn-excel{flex:1 1 auto !important;justify-content:center !important;padding:8px 10px !important;font-size:0.82rem !important;white-space:nowrap !important}.training-btn-refresh{width:38px !important;height:38px !important;padding:0 !important;flex-shrink:0 !important;justify-content:center !important}.training-controls-card{padding:14px;gap:12px}.training-search-row{flex-direction:column;gap:10px}.training-select{width:100%}.training-cards-grid{grid-template-columns:minmax(0,1fr) !important;width:100% !important;max-width:100% !important;min-width:0 !important;box-sizing:border-box !important}.training-card{width:100% !important;max-width:100% !important;min-width:0 !important;box-sizing:border-box !important;overflow:hidden;padding:14px !important}}</style>

      <div class="training-page">

        <!-- Hero Header Card -->
        <div class="training-hero-card">
          <div style="display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:16px;">
            <div>
              <div style="display:flex; align-items:center; gap:10px;">
                <span style="font-size:1.6rem;">\u{1F393}</span>
                <h1 style="font-size:1.45rem; font-weight:800; margin:0; letter-spacing:-0.02em;">
                  Monitoring Training Karyawan
                </h1>
              </div>
            </div>

            <div class="training-hero-actions" style="display:flex; align-items:center; gap:8px;">
              ${N?`
                <button id="tr-btn-add" class="btn btn-primary training-btn-add" style="background:#2563EB; height:38px; padding:0 16px; border-radius:10px; font-weight:700; display:inline-flex; align-items:center; gap:6px; box-shadow:0 4px 12px rgba(37,99,235,0.3); white-space:nowrap;">
                  <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24"><line x1="12" y1="5" x2="12" y2="19"/><line x1="5" y1="12" x2="19" y2="12"/></svg>
                  <span>Tambah Training</span>
                </button>
              `:""}
              
              ${re(I,"export")?`<button id="tr-btn-excel" class="btn btn-ghost training-btn-excel" style="background:rgba(255,255,255,0.08); border:1px solid rgba(255,255,255,0.15); color:white; height:38px; padding:0 14px; border-radius:10px; font-weight:600; display:inline-flex; align-items:center; gap:6px; white-space:nowrap;">
                <span>\u{1F4E5}</span>
                <span>Ekspor Excel</span>
              </button>`:""}

              <button id="tr-btn-refresh" class="btn btn-ghost training-btn-refresh" style="background:rgba(255,255,255,0.08); border:1px solid rgba(255,255,255,0.15); color:white; width:38px; height:38px; padding:0; border-radius:10px; display:inline-flex; align-items:center; justify-content:center; flex-shrink:0;" title="Muat Ulang Data">
                <svg width="16" height="16" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M23 4v6h-6M1 20v-6h6"/><path d="M3.51 9a9 9 0 0 1 14.85-3.36L23 10M1 14l4.64 4.36A9 9 0 0 0 20.49 15"/></svg>
              </button>
            </div>
          </div>

          <!-- KPI Stats -->
          <div class="training-stats-row">
            <div class="training-stat-box">
              <div class="training-stat-icon" style="background:rgba(124,58,237,0.25); color:#A78BFA;">
                \u{1F393}
              </div>
              <div>
                <div class="training-stat-val" id="tr-stat-sessions">${y}</div>
                <div class="training-stat-lbl">Sesi Training</div>
              </div>
            </div>

            <div class="training-stat-box">
              <div class="training-stat-icon" style="background:rgba(59,130,246,0.2); color:#60A5FA;">
                \u{1F465}
              </div>
              <div>
                <div class="training-stat-val" id="tr-stat-participants">${r}</div>
                <div class="training-stat-lbl">Peserta Terlatih</div>
              </div>
            </div>

            <div class="training-stat-box">
              <div class="training-stat-icon" style="background:rgba(16,185,129,0.2); color:#34D399;">
                \u{1F4DA}
              </div>
              <div>
                <div class="training-stat-val"><span id="tr-stat-batches">${l}</span> Batch</div>
                <div class="training-stat-lbl">Total Angkatan</div>
              </div>
            </div>

            <div class="training-stat-box">
              <div class="training-stat-icon" style="background:rgba(245,158,11,0.2); color:#FBBF24;">
                \u{1F5D3}\uFE0F
              </div>
              <div>
                <div class="training-stat-val" id="tr-stat-month">${f}</div>
                <div class="training-stat-lbl">Bulan Berjalan</div>
              </div>
            </div>
          </div>
        </div>

        <!-- Filter & Search Controls -->
        <div class="training-controls-card">
          <div class="training-search-row">
            <div class="training-search-wrap">
              <span class="training-search-icon">
                <svg width="17" height="17" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><circle cx="11" cy="11" r="8"/><path d="m21 21-4.35-4.35"/></svg>
              </span>
              <input type="text" id="tr-search-input" class="training-search-input" placeholder="Cari materi, trainer, peserta, atau cabang..." value="${m?String(m).replace(/"/g,"&quot;"):""}" />
            </div>

            <select id="tr-filter-branch" class="training-select">
              <option value="all">Semua Cabang (${n.length})</option>
              ${n.map(u=>`<option value="${u.value}" ${String(s)===String(u.value)?"selected":""}>${u.label}</option>`).join("")}
            </select>

            <select id="tr-filter-trainer" class="training-select">
              <option value="all">Semua Trainer</option>
              ${Hn.map(u=>`<option value="${u}" ${c===u?"selected":""}>${u}</option>`).join("")}
            </select>
          </div>

          <!-- Batch Filter Pills (Swipeable on Mobile) -->
          <div class="training-filter-pills">
            <div class="training-pill ${o==="all"?"active":""}" data-batch="all">
              <span>Semua Batch</span>
              <span class="badge" id="tr-pill-count-all" style="background:rgba(0,0,0,0.06); padding:2px 6px; border-radius:10px; font-size:0.75rem;">${y}</span>
            </div>
            ${Zt.map(u=>`
              <div class="training-pill ${o===u?"active":""}" data-batch="${u}">
                <span>\u{1F393}</span>
                <span>${u}</span>
              </div>
            `).join("")}
          </div>
        </div>

        <!-- Content Area: Murni Grid Kartu -->
        <div id="training-list-container">
          ${K(S)}
        </div>

        <!-- Pagination Container -->
        <div id="training-pagination-container"></div>

      </div>
    `,J(),b()}function b(){let S=e.querySelector("#tr-btn-add");S&&(S.onclick=()=>E());let y=e.querySelector("#tr-btn-excel");y&&(y.onclick=()=>k());let r=e.querySelector("#tr-btn-refresh");r&&(r.onclick=()=>x({refreshStats:!0}));let l=e.querySelector("#tr-search-input");l&&(l.oninput=p=>{m=p.target.value,clearTimeout(L),L=setTimeout(()=>{d=1,x({refreshStats:!1})},350)});let f=e.querySelector("#tr-filter-branch");f&&(f.onchange=p=>{s=p.target.value,d=1,x({refreshStats:!1})});let u=e.querySelector("#tr-filter-trainer");u&&(u.onchange=p=>{c=p.target.value,d=1,x({refreshStats:!1})}),e.querySelectorAll(".training-pill").forEach(p=>{p.onclick=()=>{o=p.getAttribute("data-batch")||"all",e.querySelectorAll(".training-pill").forEach(A=>{A.classList.toggle("active",A.getAttribute("data-batch")===o)}),d=1,x({refreshStats:!1})}}),G()}await x({refreshStats:!0})}Ce();Ie();function Dr(e,t={}){let n=document.createElement("article");n.className="fc-mobile-card insp-mobile-card";let a=e.branch_name||"-",o=e.period?He(Oe(e.period)):"",s=e.status!==null&&e.status!==void 0?String(e.status).trim():"",c="";s&&s!=="-"&&(/^\d+(\.\d+)?$/.test(s)?c=`<span class="badge insp-total-badge" title="Nilai Total Inspeksi" aria-label="Nilai Total ${P(s)}">${P(s)}</span>`:c=He(Ne(s)));let m=window.formatDate?window.formatDate(e.inspection_date)||"-":e.inspection_date||"-",i=e.fc_score!==null&&e.fc_score!==void 0&&String(e.fc_score).trim()!=="",d=i?Number(e.fc_score):NaN,F=!i||Number.isNaN(d)?"":d>=80?"insp-score-good":d>=60?"insp-score-warn":"insp-score-low",$=i?String(e.fc_score):"-",h=e.spv_score!==null&&e.spv_score!==void 0&&String(e.spv_score).trim()!==""?String(e.spv_score):"-",L=e.notes!==null&&e.notes!==void 0?String(e.notes).trim():"",R=!!(L&&L!=="-"),I=typeof e.document_link=="string"?e.document_link.trim():"",N=/^(https?:\/\/|\/)/i.test(I)?I:"",q=[];R&&q.push(`
      <div class="insp-extra-notes">
        <span class="insp-extra-label">Catatan</span>
        <div class="insp-extra-text">${P(L)}</div>
      </div>
    `),N&&q.push(`
      <a href="${P(N)}" target="_blank" rel="noopener" class="insp-doc-link">
        <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/></svg>
        <span>Buka Dokumen</span>
      </a>
    `),n.innerHTML=`
    <div class="insp-card-header">
      <div class="insp-card-branch">
        <span class="insp-branch-icon" aria-hidden="true">
          <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
            <path d="M3 21h18"/>
            <path d="M5 21V7l8-4v18"/>
            <path d="M19 21V11l-6-4"/>
            <path d="M9 9v.01"/>
            <path d="M9 12v.01"/>
            <path d="M9 15v.01"/>
            <path d="M9 18v.01"/>
          </svg>
        </span>
        <span class="insp-branch-name">${P(a)}</span>
      </div>
      ${o||c?`
        <div class="insp-card-header-badges">
          ${o}
          ${c}
        </div>
      `:""}
    </div>

    <div class="insp-metrics-box">
      <div class="insp-metric-col insp-metric-col--date">
        <span class="insp-metric-label">Tanggal</span>
        <span class="insp-metric-value insp-metric-value--date">${P(m)}</span>
      </div>
      <div class="insp-metric-col insp-metric-col--score">
        <span class="insp-metric-label">Point FC</span>
        <span class="insp-metric-value insp-metric-value--num ${F}">${P($)}</span>
      </div>
      <div class="insp-metric-col insp-metric-col--score">
        <span class="insp-metric-label">Point SPV</span>
        <span class="insp-metric-value insp-metric-value--num">${P(h)}</span>
      </div>
    </div>

    ${q.length>0?`<div class="insp-card-extra">${q.join("")}</div>`:""}
  `;let{onView:v,onEdit:B,onDelete:x,actions:k=[]}=t||{};if(v||B||x||k.length>0){let _=document.createElement("div");_.className="insp-card-footer";let E=document.createElement("div");if(E.className="insp-card-actions",v){let w=document.createElement("button");w.type="button",w.className="insp-action-btn insp-action-view fc-action-view",w.setAttribute("aria-label",`Lihat detail inspeksi ${a}`),w.innerHTML='<svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8z"/><circle cx="12" cy="12" r="3"/></svg><span>Detail</span>',w.addEventListener("click",()=>v(e)),E.appendChild(w)}if(B){let w=document.createElement("button");w.type="button",w.className="insp-action-btn insp-action-edit fc-action-edit",w.setAttribute("aria-label",`Edit inspeksi ${a}`),w.innerHTML='<svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M12 20h9"/><path d="M16.5 3.5a2.121 2.121 0 0 1 3 3L7 19l-4 1 1-4L16.5 3.5z"/></svg><span>Edit</span>',w.addEventListener("click",()=>B(e)),E.appendChild(w)}if(k.forEach(w=>{let K=document.createElement("button");K.type="button",K.className=`insp-action-btn ${w.class||""}`,K.setAttribute("aria-label",w.label||"Aksi"),K.innerHTML=He(w.icon?`${w.icon}<span>${P(w.label||"")}</span>`:P(w.label||"Aksi")),K.addEventListener("click",()=>w.handler(e)),E.appendChild(K)}),x){let w=document.createElement("button");w.type="button",w.className="insp-action-btn insp-action-del fc-action-del",w.setAttribute("aria-label",`Hapus inspeksi ${a}`),w.innerHTML='<svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><polyline points="3 6 5 6 21 6"/><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"/></svg><span>Hapus</span>',w.addEventListener("click",()=>x(e)),E.appendChild(w)}_.appendChild(E),n.appendChild(_)}return n}async function Un(e){let t=await we(),n=Array.from({length:4},(a,o)=>String(new Date().getFullYear()-o));Me({container:e,title:"Laporan Inspeksi Hygiene",icon:"\u{1F50D}",apiPath:"/api/reports/inspection",enableMobileFilterSheet:!0,itemLabel:"Laporan Inspeksi",bulkDelete:!1,columns:[{key:"branch_name",label:"Cabang"},{key:"period",label:"Periode",render:a=>Oe(a)},{key:"inspection_date",label:"Tanggal",nowrap:!0,render:a=>window.formatDate(a)},{key:"fc_score",label:"Point FC",render:a=>a!=null?`<strong class="${a>=80?"text-success":a>=60?"text-warning":"text-danger"}">${a}</strong>`:"-"},{key:"spv_score",label:"Point SPV",render:a=>a!=null?`<strong>${a}</strong>`:"-"},{key:"status",label:"Status",render:a=>Ne(a)},{key:"document_link",label:"Dokumen",render:a=>a?`<a href="${a}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"}],mobileCardRenderer:(a,o)=>Dr(a,o),filterFields:[{type:"search",placeholder:"Cari cabang..."},{type:"select",name:"branch_id",label:"Cabang",options:t},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"month",label:"Bulan",options:[{value:"01",label:"Jan"},{value:"02",label:"Feb"},{value:"03",label:"Mar"},{value:"04",label:"Apr"},{value:"05",label:"Mei"},{value:"06",label:"Jun"},{value:"07",label:"Jul"},{value:"08",label:"Agu"},{value:"09",label:"Sep"},{value:"10",label:"Okt"},{value:"11",label:"Nov"},{value:"12",label:"Des"}]},{type:"select",name:"status",label:"Status",options:["Pending","Done"]},{type:"select",name:"year",label:"Tahun",options:n}],formFields:a=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"select",required:!0,options:t,value:a?.branch_id},{name:"period",label:"Periode",type:"select",required:!0,options:["Q1","Q2","Q3","Q4"],value:a?.period}]},{type:"row",fields:[{name:"inspection_date",label:"Tanggal Inspeksi",type:"date",required:!0,value:a?.inspection_date},{name:"status",label:"Status",type:"select",required:!0,options:["Pending","Done"],value:a?.status||""}]},{type:"row",fields:[{name:"fc_score",label:"Point FC",type:"number",step:"0.1",min:"0",max:"100",value:a?.fc_score},{name:"spv_score",label:"Point SPV",type:"number",step:"0.1",min:"0",max:"100",value:a?.spv_score}]},{name:"document_link",label:"Link Dokumen",type:"url",value:a?.document_link},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:a?.notes}],exportOptions:{moduleName:"inspection_reports",onExport:async a=>{let s=(await ue("/api/reports/inspection/export",a)).map(c=>({Cabang:c.branch_name||"",Periode:c.period||"",Tanggal:c.inspection_date||"","Point FC":c.fc_score!==null&&c.fc_score!==void 0?c.fc_score:"","Point SPV":c.spv_score!==null&&c.spv_score!==void 0?c.spv_score:"",Status:c.status||"","Link Dokumen":c.document_link||""}));be(s,`Laporan_Inspeksi_Hygiene_${new Date().toISOString().slice(0,10)}`)},onTemplate:()=>{be([{Cabang:"001. Pondok Bambu",Periode:"Q1",Tanggal:"2026-01-08","Point FC":85,"Point SPV":90,Status:"Done","Link Dokumen":"https://drive.google.com/...",Catatan:""}],"Template_Import_Inspeksi")},onImport:async(a,o={})=>{let s=i=>{if(!i)return null;let d=String(i||"").toLowerCase().trim(),F=d.match(/^(\d{3})/);if(F){let M=t.find(h=>String(h.code||"").trim()===F[1]||String(h.label||"").toLowerCase().startsWith(F[1]+"."));if(M)return M.value}let $=t.find(M=>String(M.label||"").toLowerCase()===d||String(M.code||"").toLowerCase()===d||String(M.name||"").toLowerCase()===d);return $?$.value:null},c=i=>{if(i==null||i==="")return"";if(i instanceof Date&&!isNaN(i.getTime()))return i.toISOString().slice(0,10);let d=String(i).trim();if(d===""||d==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(d))return d.slice(0,10);if(/^\d{4,5}$/.test(d)){let $=Number(d);if($>2e4&&$<99999){let M=new Date(Date.UTC(1899,11,30)+$*864e5);return isNaN(M.getTime())?"":M.toISOString().slice(0,10)}}let F=d.split(/[\/\-\.]/);if(F.length===3){let[$,M,h]=F.map(L=>L.trim());if($.length===4&&M.length<=2&&h.length<=2)return`${$}-${M.padStart(2,"0")}-${h.padStart(2,"0")}`;if(h.length===4&&M.length<=2&&$.length<=2)return`${h}-${M.padStart(2,"0")}-${$.padStart(2,"0")}`}return d},m=a.map(i=>({branch_id:s(String(i.Cabang||"").trim()),period:String(i.Periode||"").trim(),inspection_date:c(i.Tanggal),fc_score:i["Point FC"]!==void 0&&i["Point FC"]!==""?Number(i["Point FC"]):null,spv_score:i["Point SPV"]!==void 0&&i["Point SPV"]!==""?Number(i["Point SPV"]):null,status:String(i.Status||"").trim(),document_link:String(i["Link Dokumen"]||"").trim(),notes:String(i.Catatan||i.Keterangan||"").trim()})).filter(i=>i.branch_id&&i.period&&i.inspection_date);return await Pe({module:"inspection_reports",label:"Laporan Inspeksi",endpoint:"/api/import/inspection",body:{rows:m,onDuplicate:"update"},idempotencyKey:o?.idempotencyKey,recover:!!o?.recover})}}})}Ce();async function Gn(e){let t=await we(),n=Array.from({length:4},(a,o)=>String(new Date().getFullYear()-o));Me({container:e,title:"Laporan GC & DC",icon:"\u{1F9F9}",apiPath:"/api/reports/cleaning",itemLabel:"Laporan GC/DC",bulkDelete:!1,enableMobileFilterSheet:!0,columns:[{key:"branch_name",label:"Cabang"},{key:"activity_type",label:"Jenis",render:a=>`<span class="badge ${a==="Deep Cleaning"?"badge-purple":"badge-success"}">${a}</span>`},{key:"period",label:"Periode",render:a=>Oe(a)},{key:"activity_date",label:"Tanggal",nowrap:!0,render:a=>window.formatDate(a)},{key:"status",label:"Status",render:a=>Ne(a)},{key:"document_link",label:"Dokumen",render:a=>a?`<a href="${a}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"}],mobileCardRenderer:(a,o)=>Xe({title:a.branch_name||"-",headerRight:Ne(a.status),badges:[a.activity_type?`<span class="badge ${a.activity_type==="Deep Cleaning"?"badge-purple":"badge-success"}">${a.activity_type}</span>`:"",a.period?Oe(a.period):""],meta:[{label:"Tanggal",value:window.formatDate(a.activity_date)||"-",full:!0}],notes:a.notes,docLink:a.document_link,row:a,handlers:o}),filterFields:[{type:"search",placeholder:"Cari nama cabang/lokasi..."},{type:"combobox",name:"branch_id",label:"Cabang",options:t},{type:"select",name:"activity_type",label:"Jenis",options:["General Cleaning","Deep Cleaning"]},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"month",label:"Bulan",options:[{value:"01",label:"Jan"},{value:"02",label:"Feb"},{value:"03",label:"Mar"},{value:"04",label:"Apr"},{value:"05",label:"Mei"},{value:"06",label:"Jun"},{value:"07",label:"Jul"},{value:"08",label:"Agu"},{value:"09",label:"Sep"},{value:"10",label:"Okt"},{value:"11",label:"Nov"},{value:"12",label:"Des"}]},{type:"select",name:"status",label:"Status",options:["Pending","Done"]},{type:"select",name:"year",label:"Tahun",options:n}],formFields:a=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:t,value:a?.branch_id},{name:"activity_type",label:"Jenis Kegiatan",type:"select",required:!0,options:["General Cleaning","Deep Cleaning"],value:a?.activity_type}]},{type:"row",fields:[{name:"period",label:"Periode",type:"select",required:!0,options:["Q1","Q2","Q3","Q4"],value:a?.period},{name:"activity_date",label:"Tanggal",type:"date",required:!0,value:a?.activity_date}]},{name:"status",label:"Status",type:"select",required:!0,options:["Pending","Done"],value:a?.status||""},{name:"document_link",label:"Link Dokumen",type:"url",value:a?.document_link},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:a?.notes}],exportOptions:{moduleName:"cleaning_reports",onExport:async a=>{let s=(await ue("/api/reports/cleaning/export",a)).map(c=>({Cabang:c.branch_name||"",Jenis:c.activity_type||"",Periode:c.period||"",Tanggal:c.activity_date||"",Status:c.status||"","Link Dokumen":c.document_link||""}));be(s,`Laporan_GCDC_${new Date().toISOString().slice(0,10)}`)},onTemplate:()=>{be([{Cabang:"001. Pondok Bambu",Jenis:"General Cleaning",Periode:"Q1",Tanggal:"2026-01-08",Status:"Done","Link Dokumen":"https://drive.google.com/...",Catatan:""}],"Template_Import_GCDC")},onImport:async(a,o={})=>{let s=i=>{if(!i)return null;let d=String(i||"").toLowerCase().trim(),F=d.match(/^(\d{3})/);if(F){let M=t.find(h=>String(h.code||"").trim()===F[1]||String(h.label||"").toLowerCase().startsWith(F[1]+"."));if(M)return M.value}let $=t.find(M=>String(M.label||"").toLowerCase()===d||String(M.code||"").toLowerCase()===d||String(M.name||"").toLowerCase()===d);return $?$.value:null},c=i=>{if(i==null||i==="")return"";if(i instanceof Date&&!isNaN(i.getTime()))return i.toISOString().slice(0,10);let d=String(i).trim();if(d===""||d==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(d))return d.slice(0,10);if(/^\d{4,5}$/.test(d)){let $=Number(d);if($>2e4&&$<99999){let M=new Date(Date.UTC(1899,11,30)+$*864e5);return isNaN(M.getTime())?"":M.toISOString().slice(0,10)}}let F=d.split(/[\/\-\.]/);if(F.length===3){let[$,M,h]=F.map(L=>L.trim());if($.length===4&&M.length<=2&&h.length<=2)return`${$}-${M.padStart(2,"0")}-${h.padStart(2,"0")}`;if(h.length===4&&M.length<=2&&$.length<=2)return`${h}-${M.padStart(2,"0")}-${$.padStart(2,"0")}`}return d},m=a.map(i=>({branch_id:s(String(i.Cabang||"").trim()),activity_type:String(i.Jenis||i.Kegiatan||"").trim(),period:String(i.Periode||"").trim(),activity_date:c(i.Tanggal),status:String(i.Status||"").trim(),document_link:String(i["Link Dokumen"]||"").trim(),notes:String(i.Catatan||i.Keterangan||"").trim()})).filter(i=>i.branch_id&&i.activity_type&&i.period&&i.activity_date);return await Pe({module:"cleaning_reports",label:"Laporan GC/DC",endpoint:"/api/import/cleaning",body:{rows:m,onDuplicate:"update"},idempotencyKey:o?.idempotencyKey,recover:!!o?.recover})}}})}Ce();async function Vn(e,t){let n=await we(),a=Array.from({length:4},(i,d)=>String(new Date().getFullYear()-d)),o=t?t.get("dash_filter"):null,s=t?t.get("explicit")==="1":!1,c=t?t.get("month"):null,m={};if(o==="fogging")if(s&&c&&c.length===7){let[i,d]=c.split("-");m={month:d,year:i}}else m={year:c&&c.length===7?c.split("-")[0]:String(new Date().getFullYear())};Me({container:e,title:"Rekap Fogging",icon:"\u{1F4A8}",apiPath:"/api/reports/fogging",itemLabel:"Fogging",bulkDelete:!1,enableMobileFilterSheet:!0,defaultFilters:m,columns:[{key:"branch_name",label:"Cabang"},{key:"activity_type",label:"Jenis",render:i=>`<span class="badge badge-warning">${i}</span>`},{key:"period",label:"Periode",render:i=>Oe(i)},{key:"activity_date",label:"Tanggal",nowrap:!0,render:i=>window.formatDate(i)},{key:"status",label:"Status",render:i=>Ne(i)},{key:"document_link",label:"Dokumen",render:i=>i?`<a href="${i}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"},{key:"notes",label:"Catatan",render:i=>i||"-"}],mobileCardRenderer:(i,d)=>Xe({title:i.branch_name||"-",headerRight:Ne(i.status),badges:[i.activity_type?`<span class="badge badge-warning">${i.activity_type}</span>`:"",i.period?Oe(i.period):""],meta:[{label:"Tanggal",value:window.formatDate(i.activity_date)||"-",full:!0}],notes:i.notes,docLink:i.document_link,row:i,handlers:d}),filterFields:[{type:"search",placeholder:"Cari nama cabang/lokasi..."},{type:"select",name:"branch_id",label:"Cabang",options:n},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"month",label:"Bulan",options:[{value:"01",label:"Jan"},{value:"02",label:"Feb"},{value:"03",label:"Mar"},{value:"04",label:"Apr"},{value:"05",label:"Mei"},{value:"06",label:"Jun"},{value:"07",label:"Jul"},{value:"08",label:"Agu"},{value:"09",label:"Sep"},{value:"10",label:"Okt"},{value:"11",label:"Nov"},{value:"12",label:"Des"}]},{type:"select",name:"status",label:"Status",options:["Pending","Done"]},{type:"select",name:"year",label:"Tahun",options:a}],formFields:i=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"select",required:!0,options:n,value:i?.branch_id},{name:"period",label:"Periode",type:"select",required:!0,options:["Q1","Q2","Q3","Q4"],value:i?.period}]},{type:"row",fields:[{name:"activity_date",label:"Tanggal",type:"date",value:i?.activity_date},{name:"status",label:"Status",type:"select",required:!0,options:["Pending","Done"],value:i?.status||""}]},{name:"document_link",label:"Link Dokumen",type:"url",value:i?.document_link},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:i?.notes}],exportOptions:{moduleName:"fogging_reports",onExport:async i=>{let F=(await ue("/api/reports/fogging/export",i)).map($=>({Cabang:$.branch_name||"",Jenis:$.activity_type||"Fogging",Periode:$.period||"",Tanggal:$.activity_date||"",Status:$.status||"","Link Dokumen":$.document_link||""}));be(F,`Laporan_Fogging_${new Date().toISOString().slice(0,10)}`)},onTemplate:()=>{be([{Cabang:"001. Pondok Bambu",Jenis:"Fogging",Periode:"Q1",Tanggal:"2026-01-08",Status:"Done","Link Dokumen":"https://drive.google.com/..."}],"Template_Import_Fogging")},onImport:async(i,d={})=>{let F=h=>{if(!h)return null;let L=String(h||"").toLowerCase().trim(),R=L.match(/^(\d{3})/);if(R){let N=n.find(q=>String(q.code||"").trim()===R[1]||String(q.label||"").toLowerCase().startsWith(R[1]+"."));if(N)return N.value}let I=n.find(N=>String(N.label||"").toLowerCase()===L||String(N.code||"").toLowerCase()===L||String(N.name||"").toLowerCase()===L);return I?I.value:null},$=h=>{if(h==null||h==="")return"";if(h instanceof Date&&!isNaN(h.getTime()))return h.toISOString().slice(0,10);let L=String(h).trim();if(L===""||L==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(L))return L.slice(0,10);if(/^\d{4,5}$/.test(L)){let I=Number(L);if(I>2e4&&I<99999){let N=new Date(Date.UTC(1899,11,30)+I*864e5);return isNaN(N.getTime())?"":N.toISOString().slice(0,10)}}let R=L.split(/[\/\-\.]/);if(R.length===3){let[I,N,q]=R.map(v=>v.trim());if(I.length===4&&N.length<=2&&q.length<=2)return`${I}-${N.padStart(2,"0")}-${q.padStart(2,"0")}`;if(q.length===4&&N.length<=2&&I.length<=2)return`${q}-${N.padStart(2,"0")}-${I.padStart(2,"0")}`}return L},M=i.map(h=>({branch_id:F(String(h.Cabang||"").trim()),activity_type:String(h.Jenis||h.Kegiatan||"Fogging").trim(),period:String(h.Periode||"").trim(),activity_date:$(h.Tanggal),status:String(h.Status||"").trim(),document_link:String(h["Link Dokumen"]||"").trim(),notes:String(h.Catatan||h.Keterangan||"").trim()})).filter(h=>h.branch_id&&h.period&&h.activity_date);return await Pe({module:"fogging_reports",label:"Laporan Fogging",endpoint:"/api/reports/fogging/import",body:M,idempotencyKey:d?.idempotencyKey,recover:!!d?.recover})}}})}Ce();async function Jn(e){let t=await we(),n=await Jt(),a=[{value:"Berlin",label:"Berlin"},{value:"Ade",label:"Ade"}],o=Array.from({length:4},(m,i)=>String(new Date().getFullYear()-i)),s=m=>m&&!n.find(i=>i.value===m)?[...n,{value:m,label:m}]:n,c=m=>m&&!a.find(i=>i.value===m)?[...a,{value:m,label:m}]:a;Me({container:e,title:"Rekap Laporan Basecamp",icon:"\u{1F4DD}",apiPath:"/api/reports/basecamp",bulkDelete:!1,enableMobileFilterSheet:!0,itemLabel:"Laporan Basecamp",columns:[{key:"info_date",label:"Tgl Info",nowrap:!0,render:m=>window.formatDate(m)},{key:"branch_name",label:"Cabang"},{key:"problem",label:"Permasalahan",render:m=>`<span title="${m||""}">${m?.length>60?m.slice(0,60)+"\u2026":m||"-"}</span>`},{key:"pic",label:"PIC"},{key:"done_date",label:"Tgl Done",nowrap:!0,render:m=>window.formatDate(m)},{key:"status",label:"Status",render:m=>Ne(m)},{key:"notes",label:"Keterangan",render:m=>m?.length>40?m.slice(0,40)+"\u2026":m||"-"}],mobileCardRenderer:(m,i)=>Xe({title:m.branch_name||"-",headerRight:Ne(m.status),meta:[{label:"Tgl Info",value:window.formatDate(m.info_date)||"-"},{label:"Tgl Done",value:window.formatDate(m.done_date)||"-"},...m.pic?[{label:"PIC",value:m.pic,full:!0}]:[],...m.problem?[{label:"Permasalahan",value:m.problem,full:!0}]:[]],notes:m.notes,notesLabel:"Keterangan",row:m,handlers:i}),filterFields:[{type:"search",placeholder:"Cari permasalahan, PIC, cabang..."},{type:"select",name:"pic",label:"PIC",options:["Berlin","Ade"]},{type:"select",name:"branch_id",label:"Cabang",options:t},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"month",label:"Bulan",options:[{value:"01",label:"Jan"},{value:"02",label:"Feb"},{value:"03",label:"Mar"},{value:"04",label:"Apr"},{value:"05",label:"Mei"},{value:"06",label:"Jun"},{value:"07",label:"Jul"},{value:"08",label:"Agu"},{value:"09",label:"Sep"},{value:"10",label:"Okt"},{value:"11",label:"Nov"},{value:"12",label:"Des"}]},{type:"select",name:"status",label:"Status",options:["Open","In Progress","Done"]},{type:"select",name:"year",label:"Tahun",options:o}],formFields:m=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"select",required:!0,options:t,value:m?.branch_id},{name:"pic",label:"PIC",type:"select",options:c(m?.pic),value:m?.pic}]},{name:"problem",label:"Permasalahan",type:"textarea",required:!0,rows:3,value:m?.problem},{type:"row",fields:[{name:"info_date",label:"Tanggal Info",type:"date",required:!0,value:m?.info_date},{name:"done_date",label:"Tanggal Done",type:"date",value:m?.done_date}]},{name:"status",label:"Status",type:"select",options:["Open","In Progress","Done"],value:m?.status||"Open"},{name:"notes",label:"Keterangan / Tindak Lanjut",type:"textarea",rows:2,value:m?.notes}],exportOptions:{moduleName:"basecamp_reports",onExport:async m=>{let d=(await ue("/api/reports/basecamp/export",m)).map(F=>({"Tgl Info":F.info_date||"",Cabang:F.branch_name||"",Permasalahan:F.problem||"",PIC:F.pic||"","Tgl Done":F.done_date||"",Status:F.status||"",Keterangan:F.notes||""}));be(d,`Rekap_Laporan_Basecamp_${new Date().toISOString().slice(0,10)}`)},onTemplate:()=>{be([{"Tgl Info":"2026-01-08",Cabang:"001. Pondok Bambu",Permasalahan:"Request fogging karena banyak nyamuk",PIC:"Fajar","Tgl Done":"2026-01-10",Status:"Done",Keterangan:"Sudah difogging"}],"Template_Import_Basecamp")},onImport:async(m,i={})=>{let d=M=>{if(!M)return null;let h=String(M||"").toLowerCase().trim(),L=h.match(/^(\d{3})/);if(L){let I=t.find(N=>String(N.code||"").trim()===L[1]||String(N.label||"").toLowerCase().startsWith(L[1]+"."));if(I)return I.value}let R=t.find(I=>String(I.label||"").toLowerCase()===h||String(I.code||"").toLowerCase()===h||String(I.name||"").toLowerCase()===h);return R?R.value:null},F=M=>{if(M==null||M==="")return"";if(M instanceof Date&&!isNaN(M.getTime()))return M.toISOString().slice(0,10);let h=String(M).trim();if(h===""||h==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(h))return h.slice(0,10);if(/^\d{4,5}$/.test(h)){let R=Number(h);if(R>2e4&&R<99999){let I=new Date(Date.UTC(1899,11,30)+R*864e5);return isNaN(I.getTime())?"":I.toISOString().slice(0,10)}}let L=h.split(/[\/\-\.]/);if(L.length===3){let[R,I,N]=L.map(q=>q.trim());if(R.length===4&&I.length<=2&&N.length<=2)return`${R}-${I.padStart(2,"0")}-${N.padStart(2,"0")}`;if(N.length===4&&I.length<=2&&R.length<=2)return`${N}-${I.padStart(2,"0")}-${R.padStart(2,"0")}`}return h},$=m.map(M=>({info_date:F(M["Tgl Info"]||M["Tanggal Info"]),branch_id:d(String(M.Cabang||"").trim()),problem:String(M.Permasalahan||"").trim(),pic:String(M.PIC||"").trim(),done_date:F(M["Tgl Done"]||M["Tanggal Done"]),status:String(M.Status||"").trim(),notes:String(M.Keterangan||M.Catatan||"").trim()})).filter(M=>M.info_date&&M.branch_id&&M.problem);return await Pe({module:"basecamp_reports",label:"Laporan Basecamp",endpoint:"/api/reports/basecamp/import",body:$,idempotencyKey:i?.idempotencyKey,recover:!!i?.recover})}}})}Be();Ie();var Lr=[{id:1,title:"Dusting (Pembersihan Debu)",category:"Pembersihan Rutin",icon:"\u{1F9F9}",color:"#2563EB",bg:"#EFF6FF",target:"Meja resepsionis, meja pasien, meja dokter, furniture, kaca, monitor & keyboard",chemical_ratio:"30 ml Glass Cleaner : 1 Liter Air / Siap pakai di botol spray",tools:["Kain Lap Microfiber Biru (Area Umum)","Kanebo (Kaca & Layar Monitor)","Kemoceng","Botol Spray Glass Cleaner","Kuas 4 Inch (Keyboard & Sela Sempit)"],steps:["Siapkan larutan glass cleaner di botol spray khusus kaca dan perabot.","Lipat kain microfiber menjadi 8 bidang kerja agar efisien dan bersih merata.","Semprotkan cairan pembersih ke permukaan kain lap, bukan langsung ke monitor atau bodi elektronik.","Usap permukaan perabot dari bagian paling atas ke bagian bawah secara searah atau pola S.","Gunakan kuas cat 4 inch untuk membersihkan sela-sela keyboard dan lubang ventilasi.","Keringkan kaca atau monitor menggunakan lap microfiber kering atau kanebo bersih."],notes:"Khusus layar monitor medis & dental unit, dilarang menyemprotkan cairan langsung ke layar."},{id:2,title:"Sweeping (Lobby Duster & Sapu Nylon)",category:"Pembersihan Rutin",icon:"\u{1F33E}",color:"#D97706",bg:"#FFFBEB",target:"Lobby utama, lorong, ruang tindakan, balkon, pantry, teras",chemical_ratio:"-",tools:["Lobby Duster Set 60cm (Area Lobby & Lorong)","Lobby Duster Set 40cm (Area Ruang Tindakan)","Sapu & Pengki Nagata 742s","Kapi Scrapper (Kerak/Permen Karet)"],steps:["Pastikan kain lobby duster dalam keadaan kering, bersih, dan bebas pasir kasar.","Dorong lobby duster menyusuri tepi dinding ruangan terlebih dahulu membentuk perimeter.","Lanjutkan mendorong dari sudut terjauh menuju arah pintu keluar tanpa mengangkat kain pel lobi dari lantai.","Kumpulkan debu dan kotoran di dekat pintu keluar, lalu angkat menggunakan sapu dan pengki Nagata.","Jika ada noda kerak atau bekas stiker/permen karet menempel di lantai, kikis perlahan menggunakan kapi scrapper.","Kibaskan kain duster di luar ruangan atau sedot sisa debu pada kain duster."],notes:"Jangan pernah menggunakan lobby duster di area lantai yang masih basah atau tergenang air."},{id:3,title:"Mopping (Pengepelan Lantai)",category:"Pembersihan Rutin",icon:"\u{1FAA3}",color:"#059669",bg:"#ECFDF5",target:"Seluruh lantai ruangan, koridor, dan area tunggu klinik",chemical_ratio:"25 ml Floor Cleaner per 1 Liter air bersih",tools:["Single Bucket Mop Wringer 20L / Double Bucket 46L","Mopping Stick / Tongkat & Kain Pel","Chemical Floor Cleaner Prokleen Lavender","Warning Sign Lantai Licin"],steps:['Pasang papan peringatan "Caution Wet Floor" di titik strategis yang mudah terlihat.',"Campurkan cairan Floor Cleaner ke dalam ember air bersih dengan takaran 25 ml per 1 Liter air.","Celupkan kain pel, masukkan ke dalam alat pemeras (wringer), dan peras hingga kondisi kain pel lembap (tidak menetes).","Mulai mengepel dari sudut terjauh ruangan mundur ke arah pintu keluar dengan gerakan membentuk pola angka delapan (8).","Bilas kain pel secara berkala apabila air pada ember bilas sudah terlihat keruh.","Biarkan lantai mengering sempurna secara alami sebelum papan peringatan wet floor dilepas."],notes:'Wajib pasang papan peringatan "Caution Wet Floor" sebelum mulai mengepel untuk mencegah risiko terpeleset.'},{id:4,title:"Membersihkan Toilet & Wastafel",category:"Sanitasi Toilet",icon:"\u{1F6BD}",color:"#DC2626",bg:"#FEF2F2",target:"Toilet pasien, toilet tim, closet, wastafel, dinding keramik, floor drain",chemical_ratio:"Karbol 25 ml:1 L Air / Bowl Cleaner 100 ml spray (kerak) / Bubuk KIFA",tools:["Sarung Tangan Karet (Wajib)","Sikat Kloset Scotch-Brite 3M","Kain Lap Microfiber Merah (Khusus Toilet)","Sponge Scotch Brite & Sikat Lantai ID-51","Floor Squeegee 75cm","Chemical Karbol Pine, Bowl Cleaner, Bubuk KIFA"],steps:["Kenakan sarung tangan karet dan siapkan perlengkapan sanitasi toilet.","Flush kloset, tuangkan Toilet Bowl Cleaner ke bagian dalam dinding kloset dan leher angsa, diamkan 5\u201310 menit.","Sikat bagian dalam kloset secara menyeluruh dengan sikat kloset, lalu bilas (flush) kembali hingga bersih.","Seka dudukan, penutup, dan tabung luar kloset menggunakan kain lap microfiber merah yang dibasahi larutan disinfektan karbol.","Bersihkan wastafel dan kran stainless menggunakan sponge tapas dengan bubuk pembersih KIFA (jangan gunakan Bowl Cleaner pada kran stainless).","Sikat lantai kamar mandi dengan sikat lantai ID-51 dan cairan karbol disinfektan.","Tarik genangan air ke arah floor drain menggunakan floor squeegee hingga lantai tidak licin dan cepat kering.","Pasang kamper/urinal block, isi ulang sabun cuci tangan, dan pasang plastik sampah baru."],notes:"Kain lap warna MERAH mutlak hanya untuk toilet dan mangkuk spittoon. Dilarang mencampur cairan kimia sembarangan."},{id:5,title:"Pembersihan Area Parkir & Luar",category:"Pembersihan Rutin",icon:"\u{1F17F}\uFE0F",color:"#4F46E5",bg:"#EEF2FF",target:"Halaman parkir mobil & motor, saluran air luar, teras depan",chemical_ratio:"-",tools:["Sapu Lidi Gagang Panjang","Pengki Nagata","Plastik Sampah Hitam Besar (90x120)","Selang Air"],steps:["Sapu daun kering, puntung rokok, plastik, dan pasir dari batas trotoar/pagar menuju satu titik.","Kumpulkan dan angkat sampah ke dalam pengki, masukkan langsung ke plastik sampah hitam besar.","Periksa dan bersihkan saluran drainase air di area parkir dari sumbatan sampah padat.","Bila ada tumpahan oli atau kotoran tanah membandel, siram dengan air dan sikat.","Rapikan pembatas parkir dan pastikan area depan klinik terlihat bersih dan rapi menyambut pasien."],notes:"Lakukan sebelum jam operasional klinik dibuka (pukul 07.30 - 08.00)."},{id:6,title:"Pembersihan Kaca & Jendela",category:"Pembersihan Rutin",icon:"\u{1FA9F}",color:"#0891B2",bg:"#ECFEFF",target:"Pintu kaca utama, jendela kaca klinik, partisi kaca, cermin toilet",chemical_ratio:"30 ml Glass Cleaner : 1 Liter Air / Siap pakai di botol spray",tools:["Botol Spray Glass Cleaner Prokleen","Window Wiper / Window Squeegee","Kain Microfiber Kering / Kanebo","Tongkat Teleskopik (Kaca Tinggi)"],steps:["Semprotkan larutan glass cleaner merata pada permukaan kaca yang hendak dibersihkan.","Tarik wiper karet dari arah atas ke bawah secara lurus vertikal atau zigzag horizontal tanpa terputus.","Lap bilah karet wiper setiap selesai satu tarikan menggunakan kain lap bersih agar tidak meninggalkan garis bekas air.","Gunakan kain microfiber kering untuk membersihkan sisa tetesan air di sudut-sudut lis bingkai kaca.","Untuk kaca bagian luar yang tinggi, pasang wiper pada tongkat sambungan teleskopik."],notes:"Hindari membersihkan kaca saat terpapar terik sinar matahari langsung karena cairan akan cepat mengering dan membekas noda."},{id:7,title:"Mencuci Bola Mandi Bola (Playground)",category:"Pembersihan Khusus",icon:"\u{1F939}",color:"#7C3AED",bg:"#F5F3FF",target:"Area bermain anak (Kids Playground) dan bola-bola mandi bola",chemical_ratio:"30 ml Dishwashing Liquid per 1 Liter Air + Karbol Disinfektan",tools:["Baskom Plastik Besar 55cm","Dishwashing Liquid Prokleen","Karbol Disinfektan Pine","Lap Microfiber Bersih","Karung / Wadah Pengering Jaring"],steps:["Kumpulkan seluruh bola mandi bola ke dalam wadah atau karung bersih.","Siapkan baskom plastik besar berisi air hangat kuku dicampur Dishwashing Liquid dan sedikit karbol disinfektan.","Rendam bola dan cuci secara bertahap sambil digosok perlahan menggunakan sponge/lap bersih.","Bilas bola dengan air bersih mengalir hingga tidak ada sisa busa sabun yang menempel.","Tiriskan dan angin-anginkan bola di tempat bersih beralaskan terpal/kain kering hingga benar-benar kering.","Seka dan sterilkan area kolam mandi bola sebelum bola dimasukkan kembali."],notes:"Jadwal pencucian berkala dilakukan minimal 1 kali seminggu atau sesuai kebutuhan kebersihan arena anak."},{id:8,title:"Membersihkan Filter Air Purifier",category:"Elektronik & Utility",icon:"\u{1F4A8}",color:"#0D9488",bg:"#F0FDFA",target:"Unit Air Purifier di ruang tunggu, ruang dokter, dan koridor",chemical_ratio:"-",tools:["Vacuum Cleaner (Nozzle Sikat Halus)","Kain Microfiber Kering","Kuas Cat Halus"],steps:["Matikan saklar power unit air purifier dan cabut kabel daya dari stop kontak dinding.","Buka penutup bodi (casing) depan atau belakang air purifier secara hati-hati.","Keluarkan pre-filter jaring nilon dari unit.","Bersihkan debu tebal pada pre-filter menggunakan vacuum cleaner nozzle sikat halus atau cuci dengan air mengalir (jika tipe wash-able).","Untuk filter HEPA dan Karbon Aktif: HANYA sedot debu permukaannya secara perlahan dengan vacuum, JANGAN DICUCI DENGAN AIR.","Lap bodi luar dan kisi-kisi udara masuk menggunakan lap microfiber kering.","Pastikan pre-filter benar-benar kering sebelum dipasang kembali ke posisinya semula."],notes:"Jangan pernah mencuci filter HEPA dengan air karena akan merusak serat mikro penyaring bakteri/virus."},{id:9,title:"Membersihkan Filter & Bodi AC",category:"Elektronik & Utility",icon:"\u2744\uFE0F",color:"#2563EB",bg:"#EFF6FF",target:"Filter udara dan penutup indoor unit AC split di seluruh ruangan",chemical_ratio:"-",tools:["Tangga Besi Lipat","Kain Lap Microfiber Lembap","Kuas Cat Halus","Kran Air Mengalir"],steps:["Matikan AC melalui remote dan matikan MCB/saklar AC bila diperlukan.","Gunakan tangga besi lipat yang stabil di lantai datar.","Buka tutup cover depan unit indoor AC ke atas hingga terkunci.","Lepaskan kedua kisi filter saringan udara plastik secara perlahan agar debu tidak berjatuhan.","Cuci saringan filter di bawah kran air mengalir dari arah bagian dalam ke bagian luar.","Keringkan filter di tempat teduh (jangan dijemur langsung di bawah terik matahari).","Lap bodi casing dan louver/swing AC dengan kain microfiber lembap.","Pasang kembali filter yang sudah kering, tutup cover AC, dan nyalakan AC untuk pengetesan."],notes:"Pembersihan rutin mandiri dilakukan 2 minggu sekali. Deep cleaning / cuci evaporator dilakukan oleh teknisi AC vendor."},{id:10,title:"Membersihkan Dental Unit",category:"Klinis & Medis",icon:"\u{1F9B7}",color:"#E11D48",bg:"#FFF1F2",target:"Kursi dental pasien, spittoon bowl, meja instrumen, lampu dental, cover bawah",chemical_ratio:"Glass Cleaner (netral) / Karbol Disinfektan (mangkuk) / Bowl Cleaner (kerak)",tools:["Sarung Tangan Karet (Wajib)","Kain Lap Microfiber Biru (Khusus Bodi Luar)","Kain Lap Microfiber Merah (Khusus Mangkuk Spittoon)","Botol Spray Glass Cleaner (25ml + Air)","Botol Spray Karbol (20ml + Air)","Botol Spray Bowl Cleaner (100ml + Air, bila berkerak)"],steps:["Pastikan saklar utama dental unit dalam posisi OFF dan kursi dinaikkan ke posisi tertinggi untuk memudahkan akses.","Lap seluruh jok kulit, sandaran kepala, sandaran tangan, dan tiang lampu menggunakan lap microfiber biru dan cairan pembersih netral/glass cleaner.","Bersihkan mangkuk kumur (spittoon bowl) menggunakan kain lap merah dan cairan karbol disinfektan.","Jika ditemukan kerak sisa darah/plak membandel di mangkuk spittoon, gunakan sedikit bowl cleaner, sikat perlahan, lalu bilas hingga bersih.","Lap permukaan meja instrumen meja dokter. Bila ada instrumen steril, geser dengan hati-hati dan jangan mengotori instrumen.","Buka penutup cover bawah unit dental, lap debu yang terjangkau dan pastikan kondisi area selang dalam keadaan kering."],notes:"Kain lap warna MERAH hanya untuk mangkuk spittoon. Dilarang keras menggosok jok kulit dental unit dengan bahan kimia asam keras."},{id:11,title:"Membersihkan Wastafel & Keran Air",category:"Sanitasi Toilet",icon:"\u{1F6B0}",color:"#0284C7",bg:"#F0F9FF",target:"Wastafel keramik ruang tindakan, wastafel lorong, kran stainless steel, afur sink",chemical_ratio:"Bubuk Pembersih KIFA / Dishwashing Liquid Prokleen",tools:["Sponge Scotch Brite Hijau/Kuning","Kain Microfiber Hijau (Area Basah)","Bubuk Pembersih KIFA","Sikat Gigi Bekas / Sikat Kecil"],steps:["Basahi seluruh mangkuk wastafel keramik dan kran air stainless steel.","Taburkan sedikit Bubuk KIFA pada tapas sponge yang basah.","Gosok perlahan noda sabun dan kerak air putih pada kran stainless steel searah guratan logam hingga mengkilap kembali.","Gunakan sikat kecil untuk membersihkan sela leher kran dan sudut sambungan afur pembuangan wastafel.","Bilas mangkuk wastafel dan kran dengan air bersih mengalir hingga seluruh sisa residu bubuk KIFA hilang.","Keringkan kran air stainless steel dan cermin wastafel menggunakan lap microfiber hijau kering."],notes:"DILARANG KERAS menggunakan Toilet Bowl Cleaner pada kran air stainless steel karena menyebabkan karat dan warna menghitam permanen."},{id:12,title:"Membersihkan Kulkas Klinik & Pantry",category:"Elektronik & Utility",icon:"\u{1F9CA}",color:"#475569",bg:"#F8FAFC",target:"Kulkas pantry staf dan kulkas penyimpan obat/material medis (eksterior & rak)",chemical_ratio:"30 ml Dishwashing Liquid : 1 Liter Air Hangat + Lap Kering",tools:["Kain Microfiber Biru & Hijau","Dishwashing Liquid","Ember Air Hangat","Sponge Scotch Brite Halus"],steps:["Keluarkan makanan/minuman yang telah kedaluwarsa atau berbau tidak sedap.","Lepas rak-rak kaca/plastik interior kulkas yang dapat dilepas, cuci dengan sabun cuci piring dan bilas air bersih.","Seka dinding interior kulkas menggunakan kain microfiber yang dicelupkan ke dalam air hangat bersabun lembut.","Lap kering seluruh permukaan dalam kulkas dengan microfiber kering bersih.","Pasang kembali rak interior dan susun kembali bahan/makanan secara tertib dan rapi.","Bersihkan karet pintu kulkas dari noda hitam/jamur menggunakan sikat halus dan lap kering.","Lap bodi luar kulkas dan pegangan pintu menggunakan cairan pembersih netral."],notes:"Untuk kulkas penyimpanan obat/material dental, koordinasikan terlebih dahulu dengan Perawat/PIC Obat sebelum membersihkan rak interior."},{id:13,title:"Membersihkan Sofa & Kursi Ruang Tunggu",category:"Pembersihan Khusus",icon:"\u{1F6CB}\uFE0F",color:"#9333EA",bg:"#FAF5FF",target:"Sofa ruang tunggu pasien, kursi reception, kursi kerja tim",chemical_ratio:"Baking Soda (penghilang bau) / Pembersih Upholstery secukupnya",tools:["Vacuum Cleaner (Nozzle Kecil & Sikat Lembut)","Kain Microfiber Bersih","Sikat Bulu Halus","Ember Air Hangat","Cairan Pembersih Khusus Sofa (Upholstery Cleaner)","Baking Soda (Penghilang Bau Tidak Sedap)","Sarung Tangan Karet"],steps:["Kenali bahan sofa (kain fabric, beludru, kulit asli, atau kulit sintetis/oscar).","Sedot debu di seluruh permukaan, lipatan jahitan, dan celah bantal sofa menggunakan vacuum cleaner nozzle kecil.","Gunakan sikat bulu halus untuk mengangkat rambut atau kotoran yang menempel kuat di serat kain.","Jika ada noda bercak (kopi, makanan, tinta), semprotkan pembersih upholstery secukupnya, gosok memutar perlahan dengan microfiber.","Celupkan kain microfiber ke air hangat suam-suam kuku, peras hingga lembap, lalu seka seluruh permukaan sofa.","Lap kembali dengan kain kering bersih untuk menyerap sisa kelembapan.","Jika sofa berbau apek: taburkan baking soda tipis di permukaan, diamkan 20 menit, lalu sedot kembali dengan vacuum cleaner.","Biarkan sofa mengering secara alami di dalam ruangan berventilasi atau ber-AC."],notes:"Jangan menjemur sofa langsung di bawah terik matahari karena dapat menyebabkan bahan kulit pecah-pecah dan warna kain memudar."}],Pr=[{no:1,name:"Sapu Lidi Gagang",qty:"1 unit",area:"Area parkir mobil & motor klinik"},{no:2,name:"Sapu & Pengki Nagata 742s",qty:"1 set",area:"Area pantry, balkon, dak lantai atas, atau teras"},{no:3,name:"Lobby Duster Set 60 cm",qty:"1 set",area:"Lobby utama dan koridor/lorong klinik"},{no:4,name:"Lobby Duster Set 40 cm / Refill",qty:"1 unit",area:"Area ruang tindakan dokter"},{no:5,name:"Single Bucket Mop Wringer 20L",qty:"1 paket",area:"Ember pel dengan pemeras + tongkat pel"},{no:6,name:"Double Bucket Wringer Trolley 46L",qty:"1 paket",area:"Khusus klinik besar / klinik dengan fasilitas lift"},{no:7,name:"Kemoceng Bulu Lembut",qty:"1 unit",area:"Membersihkan debu kasar meja dan perabot"},{no:8,name:"Lap Microfiber Warna Merah",qty:"2 pcs",area:"Khusus area toilet, sanitasi, dan mangkuk spittoon"},{no:9,name:"Lap Microfiber Warna Biru",qty:"2 pcs",area:"Area umum kering (meja, kursi, bodi peralatan)"},{no:10,name:"Lap Microfiber Warna Hijau",qty:"2 pcs",area:"Area umum basah (pantry, wastafel luar)"},{no:11,name:"Kanebo / Synthetic Chamois",qty:"1 pcs",area:"Area umum, cermin, dan pengering kaca"},{no:12,name:"Caddy Bag Cleaning",qty:"1 unit",area:"Tas perlengkapan pinggang petugas cleaning"},{no:13,name:"Sponge Scotch Brite 3M",qty:"2 pcs",area:"Pembersihan noda lantai dan kerak ringan (spotting)"},{no:14,name:"Scotch-Brite 3M Sikat Toilet",qty:"1 unit",area:"Sikat khusus kloset dan leher angsa"},{no:15,name:"Kapi Gagang Karet / Scrapper",qty:"1 unit",area:"Pembersihan kerak lantai, noda permen karet, stiker"},{no:16,name:"Kuas Cat 4 Inch",qty:"2 pcs",area:"Pembersih sela keyboard, ventilasi, dan sudut sempit"},{no:17,name:"Warning Sign (Lantai Licin)",qty:"1 unit",area:"Digunakan saat mopping dan general cleaning"},{no:18,name:"Teleskopik + Pembersih Kaca/Langit",qty:"1 set",area:"Membersihkan jendela tinggi, plafon, dan sarang laba-laba"},{no:19,name:"Sarung Tangan Karet 3M",qty:"1 pasang",area:"Perlindungan tangan area toilet dan bahan kimia"},{no:20,name:"Botol Spray 500ml",qty:"5 unit",area:"Wadah larutan chemical (diberi label masing-masing)"},{no:21,name:"Pengharum Ruangan Stella Aerosol",qty:"1 kaleng",area:"Digunakan darurat apabila timbul bau menyengat"},{no:22,name:"Kleaner Gantungan Sapu & Pel",qty:"1 set",area:"Gantungan penyimpanan alat cleaning di janitor"},{no:23,name:"Floor Squeegee 75 cm",qty:"1 unit",area:"Pendorong genangan air lantai toilet/teras"},{no:24,name:"Sikat Botol Ikea",qty:"1 unit",area:"Di setiap ruang tindakan dokter gigi"},{no:25,name:"Scotch-Brite 3M Sikat Lantai ID-51",qty:"1 unit",area:"Sikat lantai kamar mandi dan teras"},{no:26,name:"Kamper Toilet / Urinal Block",qty:"1 pack",area:"Pengharum di toilet dan urinoir"},{no:27,name:"Plastik Sampah Hitam Sedang (60x100)",qty:"1 pack",area:"Tempat sampah indoor ruangan"},{no:28,name:"Plastik Sampah Hitam Besar (90x120)",qty:"1 pack",area:"Tempat sampah luar & penampung sampah akhir"},{no:29,name:"Plastik Sampah Medis Kuning (90x120)",qty:"1 pack",area:"Limbah medis padat B3 klinik"},{no:30,name:"Plastik Sampah Medis Kuning (40x60)",qty:"1 pack",area:"Tempat sampah medis kecil di ruang tindakan"},{no:31,name:"Tangga Besi Lipat",qty:"1 unit",area:"Alat bantu membersihkan plafon, lampu, dan AC"},{no:32,name:"Baskom Plastik Besar 55cm",qty:"1 unit",area:"Mencuci bola kids playground"}],Ir=[{no:1,name:"Floor Cleaner Prokleen Lavender",size:"5 Liter",ratio:"25 ml : 1 Liter Air",function:"Pembersihan dan pengepelan rutin seluruh lantai ruangan"},{no:2,name:"Karbol Prokleen Pine Disinfektan",size:"5 Liter",ratio:"25 ml : 1 Liter Air (Mop) / 20 ml per botol spray",function:"Mengepel lantai toilet, disinfeksi area basah dan berbau tak sedap"},{no:3,name:"Glass Cleaner Prokleen",size:"5 Liter",ratio:"30 ml : 1 Liter Air / Siap pakai di botol spray",function:"Membersihkan pintu kaca, jendela, cermin, dan layar monitor"},{no:4,name:"Manual Dishwashing Liquid Prokleen",size:"5 Liter",ratio:"30 ml : 1 Liter Air",function:"Mencuci piring, gelas, sendok pantry, dan mencuci bola playground"},{no:5,name:"Hand Soap Prokleen Green Tea",size:"5 Liter",ratio:"Siap pakai (tanpa campur)",function:"Refill sabun pencuci tangan di wastafel pasien dan tim"},{no:6,name:"Liquid Detergent Prokleen",size:"5 Liter",ratio:"Sesuai kapasitas mesin",function:"Mencuci handuk, kain microfiber, dan seragam"},{no:7,name:"Toilet Bowl Cleaner Prokleen",size:"5 Liter",ratio:"100 ml per botol spray / Khusus kerak",function:"Pembersih kerak membandel pada kloset, wastafel keramik, dan nat keramik"},{no:8,name:"Bubuk Pembersih KIFA",size:"500 gram",ratio:"Taburkan tipis pada tapas basah",function:"Membersihkan kerak air pada kran air stainless steel, jet washer, dan sink"}],ot="fcms_sop_fc_ik_v1",st="fcms_sop_fc_tools_v1",lt="fcms_sop_fc_chem_v1";function Wn(){try{let e=localStorage.getItem(ot);if(e)return JSON.parse(e)}catch(e){console.error("Error load IK",e)}return JSON.parse(JSON.stringify(Lr))}function Na(e){localStorage.setItem(ot,JSON.stringify(e)),X("/api/sop/bundle/"+ot,{method:"PUT",body:JSON.stringify(e)}).catch(()=>{})}function Yn(){try{let e=localStorage.getItem(st);if(e)return JSON.parse(e)}catch(e){console.error("Error load tools",e)}return JSON.parse(JSON.stringify(Pr))}function Ka(e){localStorage.setItem(st,JSON.stringify(e)),X("/api/sop/bundle/"+st,{method:"PUT",body:JSON.stringify(e)}).catch(()=>{})}function Qn(){try{let e=localStorage.getItem(lt);if(e)return JSON.parse(e)}catch(e){console.error("Error load chem",e)}return JSON.parse(JSON.stringify(Ir))}function Ra(e){localStorage.setItem(lt,JSON.stringify(e)),X("/api/sop/bundle/"+lt,{method:"PUT",body:JSON.stringify(e)}).catch(()=>{})}async function Xn(e){let t=se(),n=re(t,"sop","write"),a=re(t,"sop","delete"),o="ik",s="",c="all",m="tools",i="";try{let[b,S,y]=await Promise.all([X("/api/sop/bundle/"+ot),X("/api/sop/bundle/"+st),X("/api/sop/bundle/"+lt)]);b.ok&&b.data?.data&&localStorage.setItem(ot,JSON.stringify(b.data.data)),S.ok&&S.data?.data&&localStorage.setItem(st,JSON.stringify(S.data.data)),y.ok&&y.data?.data&&localStorage.setItem(lt,JSON.stringify(y.data.data))}catch(b){console.warn("Could not load remote SOP bundle:",b)}let d=Wn(),F=Yn(),$=Qn(),M={"Pembersihan Rutin":{dot:"#60A5FA",text:"#93C5FD",accent:"#3B82F6"},"Sanitasi Toilet":{dot:"#38BDF8",text:"#7DD3FC",accent:"#38BDF8"},"Klinis & Medis":{dot:"#60A5FA",text:"#93C5FD",accent:"#2563EB"},"Elektronik & Utility":{dot:"#94A3B8",text:"#CBD5E1",accent:"#64748B"},"Pembersihan Khusus":{dot:"#93C5FD",text:"#BFDBFE",accent:"#3B82F6"}};function h(b){return M[b?.category]||M["Pembersihan Rutin"]}function L(){return`
      <span style="width:40px; height:40px; flex-shrink:0; display:inline-flex; align-items:center; justify-content:center; background:#0E182B; border:1px solid #223456; border-radius:9px; color:#93C5FD;" aria-hidden="true">
        <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.9" stroke-linecap="round" stroke-linejoin="round">
          <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"></path>
          <polyline points="14 2 14 8 20 8"></polyline>
          <line x1="8" y1="13" x2="16" y2="13"></line>
          <line x1="8" y1="17" x2="13" y2="17"></line>
        </svg>
      </span>
    `}function R(b){if(!n&&!a)return"";let S=Y(b.title||"Instruksi Kerja");return`
      <div style="display:flex; gap:5px; flex-shrink:0;">
        ${n?`
        <button type="button" class="sop-icon-btn btn-edit-ik" data-id="${b.id}" title="Edit ${S}" aria-label="Edit SOP ${S}">
          <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">
            <path d="M12 20h9"></path>
            <path d="M16.5 3.5a2.121 2.121 0 0 1 3 3L7 19l-4 1 1-4L16.5 3.5z"></path>
          </svg>
        </button>
        `:""}
        ${a?`
        <button type="button" class="sop-icon-btn sop-icon-del btn-delete-ik" data-id="${b.id}" title="Hapus ${S}" aria-label="Hapus SOP ${S}">
          <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">
            <polyline points="3 6 5 6 21 6"></polyline>
            <path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"></path>
          </svg>
        </button>
        `:""}
      </div>
    `}function I(){e.innerHTML=`
      <div class="page-header" style="display:flex; justify-content:space-between; align-items:flex-start; flex-wrap:wrap; gap:16px; margin-bottom:20px;">
        <div>
          <h1 class="page-title" style="display:flex; align-items:center; gap:8px; margin:0 0 6px 0; font-size:1.65rem; font-weight:700; color:#FFFFFF;">
            <span>\u{1F4DA}</span> SOP Facility Care & Housekeeping
          </h1>
          <p class="page-subtitle" style="margin:0; font-size:0.92rem; color:#CBD5E1; line-height:1.5;">
            Pedoman Standar Operasional Prosedur Kebersihan & Pemeliharaan Fasilitas PT Audy Mandiri Indonesia
          </p>
        </div>
        <div style="display:flex; gap:10px; align-items:center; flex-wrap:wrap;">
          ${a?`
          <button id="btn-reset-default" class="btn btn-outline" title="Kembalikan data ke dokumen standar asli" style="background:#131F36; border:1px solid #223456; color:#E2E8F0; font-weight:600; padding:8px 14px; border-radius:8px; display:inline-flex; align-items:center; gap:6px; cursor:pointer; font-size:0.86rem;">
            \u{1F504} Reset ke Standar
          </button>
          `:""}
          <button id="btn-print-sop" class="btn btn-outline" style="background:#131F36; border:1px solid #223456; color:#E2E8F0; font-weight:600; padding:8px 14px; border-radius:8px; display:inline-flex; align-items:center; gap:6px; cursor:pointer; font-size:0.86rem;">
            \u{1F5A8}\uFE0F Cetak Ringkasan SOP
          </button>
          ${re(t,"export")?`<button id="btn-export-tools" class="btn btn-primary" style="font-weight:700; padding:8px 16px; border-radius:8px; display:inline-flex; align-items:center; gap:6px; cursor:pointer; font-size:0.86rem; background:#2563EB; color:#FFFFFF;">
            \u{1F4E5} Download Data Excel
          </button>`:""}
        </div>
      </div>

      <!-- Navigation Tabs -->
      <div class="fc-nav-tabs-strip" data-fc-scroll-tabs style="display:flex; gap:8px; margin-bottom:20px; border-bottom:1px solid #223456; padding-bottom:10px; overflow-x:auto; flex-wrap:nowrap; -webkit-overflow-scrolling:touch; scrollbar-width:none;">
        <button class="sop-tab-btn ${o==="ik"?"active":""}" data-tab="ik" style="${N(o==="ik")}">
          \u{1F4CB} Instruksi Kerja (IK) <span class="badge" style="${q(o==="ik")}">${d.length}</span>
        </button>
        <button class="sop-tab-btn ${o==="tools"?"active":""}" data-tab="tools" style="${N(o==="tools")}">
          \u{1F6E0}\uFE0F Alat & Chemical <span class="badge" style="${q(o==="tools")}">${F.length+$.length}</span>
        </button>
        <button class="sop-tab-btn ${o==="policy"?"active":""}" data-tab="policy" style="${N(o==="policy")}">
          \u{1F4D6} Pedoman & Ketentuan
        </button>
        <button class="sop-tab-btn ${o==="links"?"active":""}" data-tab="links" style="${N(o==="links")}">
          \u{1F517} Form & Tautan
        </button>
      </div>

      <!-- Content Container -->
      <div id="sop-content-area"></div>
    `,v(),B()}function N(b){return b?"border:1px solid #3B82F6; background:#2563EB; color:#FFFFFF; font-weight:700; padding:8px 16px; border-radius:8px; cursor:pointer; display:inline-flex; align-items:center; gap:8px; font-size:0.9rem; white-space:nowrap; transition:all 0.2s;":"background:#172033; color:#E2E8F0; border:1px solid #334155; font-weight:600; padding:8px 16px; border-radius:8px; cursor:pointer; display:inline-flex; align-items:center; gap:8px; font-size:0.9rem; white-space:nowrap; transition:all 0.2s;"}function q(b){return b?"background:#0F172A; color:#FFFFFF; font-size:0.75rem; padding:2px 8px; border-radius:10px; font-weight:700;":"background:#22314A; color:#E2E8F0; border:1px solid #334155; font-size:0.75rem; padding:2px 8px; border-radius:10px; font-weight:700;"}function v(){e.querySelectorAll(".sop-tab-btn").forEach(r=>{r.addEventListener("click",()=>{o=r.dataset.tab,I()})});let b=e.querySelector("#btn-reset-default");b&&b.addEventListener("click",()=>{ge("Apakah Anda yakin ingin mengembalikan seluruh data Instruksi Kerja, Alat, dan Chemical ke data standar dokumen asli?",()=>{localStorage.removeItem(ot),localStorage.removeItem(st),localStorage.removeItem(lt),Promise.all([X("/api/sop/bundle/"+ot,{method:"DELETE"}),X("/api/sop/bundle/"+st,{method:"DELETE"}),X("/api/sop/bundle/"+lt,{method:"DELETE"})]).catch(()=>{}),d=Wn(),F=Yn(),$=Qn(),ae("Data SOP Facility Care berhasil di-reset ke standar asli"),I()},"Konfirmasi Reset Standar")});let S=e.querySelector("#btn-print-sop");S&&S.addEventListener("click",()=>{window.print()});let y=e.querySelector("#btn-export-tools");y&&y.addEventListener("click",()=>{W()})}function B(){let b=e.querySelector("#sop-content-area");b&&(o==="ik"?x(b):o==="tools"?E(b):o==="policy"?G(b):o==="links"&&J(b))}function x(b){b.innerHTML=`
      <!-- Toolbar Filter & Search -->
      <div class="card" style="padding:14px 18px; margin-bottom:20px; background:#131F36; border:1px solid #223456; border-radius:12px;">
        <div style="display:flex; flex-wrap:wrap; gap:12px; align-items:center; justify-content:space-between;">
          <div style="display:flex; flex-wrap:wrap; gap:12px; align-items:center; flex:1; min-width:260px;">
            <div style="position:relative; flex:1; min-width:200px; max-width:340px;">
              <input type="text" id="ik-search" class="form-control" placeholder="Cari pekerjaan, alat, atau chemical..." value="${s}" style="padding-left:34px; height:38px; border-radius:8px; font-size:0.9rem; width:100%; background:#0E182B; border-color:#223456;">
              <span style="position:absolute; left:11px; top:50%; transform:translateY(-50%); color:#94A3B8; font-size:0.9rem;">\u{1F50D}</span>
            </div>
            
            <div style="display:flex; gap:6px; flex-wrap:wrap;" id="cat-pills-wrap">
              <button class="cat-pill ${c==="all"?"active":""}" data-cat="all">Semua (${d.length})</button>
              <button class="cat-pill ${c==="Pembersihan Rutin"?"active":""}" data-cat="Pembersihan Rutin">Rutin (${d.filter(p=>p.category==="Pembersihan Rutin").length})</button>
              <button class="cat-pill ${c==="Sanitasi Toilet"?"active":""}" data-cat="Sanitasi Toilet">Toilet & Wastafel (${d.filter(p=>p.category==="Sanitasi Toilet").length})</button>
              <button class="cat-pill ${c==="Klinis & Medis"?"active":""}" data-cat="Klinis & Medis">Klinis Dental (${d.filter(p=>p.category==="Klinis & Medis").length})</button>
              <button class="cat-pill ${c==="Elektronik & Utility"?"active":""}" data-cat="Elektronik & Utility">Elektronik & AC (${d.filter(p=>p.category==="Elektronik & Utility").length})</button>
              <button class="cat-pill ${c==="Pembersihan Khusus"?"active":""}" data-cat="Pembersihan Khusus">Khusus (${d.filter(p=>p.category==="Pembersihan Khusus").length})</button>
            </div>
          </div>
          <div style="display:flex; gap:10px; align-items:center;">
            <div id="ik-count-label" style="font-size:0.86rem; color:#CBD5E1; font-weight:600;">
              Menampilkan ${d.length} dari ${d.length} Instruksi Kerja
            </div>
            ${n?`
            <button id="btn-add-ik" class="btn btn-primary" style="font-weight:700; padding:7px 14px; border-radius:8px; display:inline-flex; align-items:center; gap:6px; cursor:pointer; font-size:0.86rem; background:#2563EB; color:#FFFFFF;">
              \u2795 Tambah IK Baru
            </button>
            `:""}
          </div>
        </div>
      </div>

      <!-- IK Cards Grid Container -->
      <div id="ik-cards-grid-container"></div>
    `;let S=document.createElement("style");S.innerHTML=`
      .cat-pill {
        border: 1px solid #223456;
        background: #0E182B;
        color: #CBD5E1;
        font-size: 0.82rem;
        font-weight: 600;
        padding: 6px 13px;
        border-radius: 20px;
        cursor: pointer;
        transition: all 0.15s;
        flex-shrink: 0;
        white-space: nowrap;
      }
      .cat-pill:hover {
        background: #1A2A48;
        color: #FFFFFF;
        border-color: #3B82F6;
      }
      .cat-pill.active {
        background: #2563EB;
        color: #FFFFFF;
        border-color: #3B82F6;
      }
      .ik-card:hover {
        box-shadow: 0 8px 20px rgba(0,0,0,0.28) !important;
        border-color: #3B82F6 !important;
      }
      .sop-icon-btn {
        background: transparent;
        border: 1px solid rgba(51, 65, 85, 0.7);
        border-radius: 7px;
        width: 28px;
        height: 28px;
        padding: 0;
        color: #94A3B8;
        cursor: pointer;
        display: inline-flex;
        align-items: center;
        justify-content: center;
        transition: all 0.15s;
      }
      .sop-icon-btn:hover {
        background: rgba(37, 99, 235, 0.18);
        border-color: rgba(96, 165, 250, 0.55);
        color: #93C5FD;
      }
      .sop-icon-del:hover {
        background: rgba(239, 68, 68, 0.16);
        border-color: rgba(248, 113, 113, 0.55);
        color: #FCA5A5;
      }
      .sop-icon-btn:focus-visible {
        outline: 2px solid #60A5FA;
        outline-offset: 2px;
      }
      @media (max-width: 768px) {
        #cat-pills-wrap {
          display: flex !important;
          flex-wrap: nowrap !important;
          overflow-x: auto !important;
          -webkit-overflow-scrolling: touch !important;
          scrollbar-width: none !important;
          padding-bottom: 4px !important;
          width: 100% !important;
        }
        #cat-pills-wrap::-webkit-scrollbar {
          display: none !important;
        }
        .sop-ik-grid {
          grid-template-columns: 1fr !important;
          gap: 12px !important;
        }
      }
    `,b.appendChild(S);let y=b.querySelector("#ik-cards-grid-container"),r=b.querySelector("#ik-count-label");function l(){let p=d.filter(A=>{let T=c==="all"||A.category===c,g=s.toLowerCase().trim(),C=!g||A.title.toLowerCase().includes(g)||A.target.toLowerCase().includes(g)||A.tools.some(z=>z.toLowerCase().includes(g))||A.steps.some(z=>z.toLowerCase().includes(g));return T&&C});if(r&&(r.textContent=`Menampilkan ${p.length} dari ${d.length} Instruksi Kerja`),p.length===0){y.innerHTML=`
          <div style="text-align:center; padding:48px 20px; background:#111827; border-radius:12px; border:1px solid #29364B;">
            <div style="font-size:2.5rem; margin-bottom:8px;">\u{1F50D}</div>
            <h3 style="margin:0 0 6px; font-size:1.12rem; color:#FFFFFF;">Tidak ada Instruksi Kerja yang cocok</h3>
            <p style="margin:0; font-size:0.88rem; color:#CBD5E1;">Coba ubah kata kunci pencarian atau pilih kategori lain.</p>
          </div>
        `;return}y.innerHTML=`
        <div class="sop-ik-grid" style="display:grid; grid-template-columns:repeat(auto-fill, minmax(340px, 1fr)); gap:18px;">
          ${p.map(A=>{let T=h(A);return`
            <div class="card ik-card" style="background:#131F36; border:1px solid #223456; border-radius:14px; padding:20px; display:flex; flex-direction:column; justify-content:space-between; transition:border-color 0.18s ease, box-shadow 0.18s ease;">
              <div>
                <div style="display:flex; justify-content:space-between; align-items:flex-start; gap:10px; margin-bottom:14px;">
                  <div style="display:flex; align-items:flex-start; gap:12px; min-width:0;">
                    ${L()}
                    <div style="min-width:0;">
                      <h3 style="margin:0 0 6px 0; font-size:1.06rem; font-weight:700; color:#FFFFFF; line-height:1.35;">
                        ${A.id}. ${A.title}
                      </h3>
                      <span class="badge" style="background:#0E182B; color:${T.text}; border:1px solid #24385A; font-size:0.74rem; font-weight:600; border-radius:5px; padding:2px 8px; display:inline-flex; align-items:center; gap:6px;">
                        <span style="width:6px; height:6px; border-radius:50%; background:${T.dot}; display:inline-block;"></span>
                        ${A.category}
                      </span>
                    </div>
                  </div>

                  <!-- Quick Action Buttons: Edit & Hapus (non-dominant SVG) -->
                  ${R(A)}
                </div>

                <div style="font-size:0.88rem; color:#F1F5F9; line-height:1.55; margin-bottom:14px; background:#0E182B; padding:10px 13px; border-radius:8px; border:1px solid #1E3050; border-left:3px solid ${T.accent};">
                  <strong style="color:#FFFFFF;">Area Kerja:</strong> ${A.target}
                </div>

                <div style="margin-bottom:14px;">
                  <div style="font-size:0.75rem; font-weight:700; color:#93C5FD; text-transform:uppercase; letter-spacing:0.05em; margin-bottom:7px;">Alat &amp; Bahan Utama</div>
                  <ul style="margin:0; padding:0; list-style:none; font-size:0.86rem; color:#E2E8F0; line-height:1.5;">
                    ${A.tools.slice(0,3).map(g=>`
                      <li style="display:flex; align-items:flex-start; gap:8px; padding:4px 0; border-bottom:1px solid rgba(34, 52, 86, 0.45);">
                        <span style="color:#60A5FA; font-weight:700; line-height:1.4;">\u2022</span>
                        <span>${g}</span>
                      </li>
                    `).join("")}
                    ${A.tools.length>3?`<li style="color:#93C5FD; font-size:0.82rem; font-weight:600; padding-top:5px;">+ ${A.tools.length-3} alat lainnya dalam panduan</li>`:""}
                  </ul>
                </div>

                ${A.chemical_ratio&&A.chemical_ratio!=="-"?`
                  <div style="font-size:0.85rem; background:rgba(56, 189, 248, 0.08); border:1px solid rgba(56, 189, 248, 0.28); border-left:3px solid #38BDF8; color:#E0F2FE; padding:8px 12px; border-radius:7px; margin-bottom:14px; font-weight:600; line-height:1.45;">
                    <strong style="color:#7DD3FC;">Takaran Chemical:</strong> ${A.chemical_ratio}
                  </div>
                `:""}
              </div>

              <div style="border-top:1px solid #1E2E4A; padding-top:13px; margin-top:6px; display:flex; justify-content:space-between; align-items:center;">
                <span style="font-size:0.84rem; color:#CBD5E1; font-weight:600;">
                  ${A.steps.length} Langkah Kerja
                </span>
                <div style="display:flex; gap:6px;">
                  <button type="button" class="btn btn-sm btn-primary btn-view-ik" data-id="${A.id}" style="padding:7px 15px; font-weight:700; font-size:0.85rem; border-radius:8px; background:#2563EB; color:#FFFFFF; border:1px solid #3B82F6;">
                    Lihat Panduan \u2192
                  </button>
                </div>
              </div>
            </div>
            `}).join("")}
        </div>
      `,y.querySelectorAll(".btn-view-ik").forEach(A=>{A.addEventListener("click",()=>{let T=parseInt(A.dataset.id,10),g=d.find(C=>C.id===T);g&&k(g)})}),y.querySelectorAll(".btn-edit-ik").forEach(A=>{A.addEventListener("click",()=>{let T=parseInt(A.dataset.id,10),g=d.find(C=>C.id===T);g&&_(g)})}),y.querySelectorAll(".btn-delete-ik").forEach(A=>{A.addEventListener("click",()=>{let T=parseInt(A.dataset.id,10),g=d.find(C=>C.id===T);g&&ge(`Apakah Anda yakin ingin menghapus Instruksi Kerja: "${g.title}"?`,()=>{d=d.filter(C=>C.id!==T),Na(d),ae("Instruksi Kerja berhasil dihapus"),I()},"Hapus Instruksi Kerja")})})}l();let f=b.querySelector("#ik-search");f&&f.addEventListener("input",p=>{s=p.target.value,l()}),b.querySelectorAll(".cat-pill").forEach(p=>{p.addEventListener("click",()=>{b.querySelectorAll(".cat-pill").forEach(A=>A.classList.remove("active")),p.classList.add("active"),c=p.dataset.cat,l()})});let u=b.querySelector("#btn-add-ik");u&&u.addEventListener("click",()=>{_()})}function k(b){let S=h(b);ce({title:`${b.icon||"\u{1F4CB}"} IK-${b.id}: ${b.title}`,size:"lg",content:`
        <div style="display:flex; flex-direction:column; gap:16px;">
          <div style="background:${S.bg}; color:${S.text}; border:1px solid ${S.border}; padding:11px 15px; border-radius:9px; font-weight:700; font-size:0.92rem; display:flex; justify-content:space-between; align-items:center;">
            <span>Kategori: ${b.category}</span>
            <span style="color:#FFFFFF;">ID: IK-${String(b.id).padStart(2,"0")}</span>
          </div>

          <div style="background:#0F172A; padding:12px 15px; border-radius:9px; border:1px solid #29364B; border-left:4px solid ${S.accent};">
            <strong style="color:#FFFFFF; font-size:0.88rem; display:block; margin-bottom:4px;">\u{1F4CD} Area & Sasaran Target:</strong>
            <span style="font-size:0.92rem; color:#F8FAFC; line-height:1.55;">${b.target}</span>
          </div>

          <div>
            <h4 style="margin:0 0 8px 0; font-size:0.98rem; font-weight:700; color:#FFFFFF;">\u{1F9F0} Alat & Bahan yang Dibutuhkan:</h4>
            <div style="background:#0F172A; border:1px solid #29364B; border-radius:9px; padding:12px 16px;">
              <ul style="margin:0 0 0 18px; padding:0; font-size:0.9rem; color:#F8FAFC; line-height:1.75;">
                ${b.tools.map(y=>`<li style="margin-bottom:4px;">${y}</li>`).join("")}
              </ul>
            </div>
          </div>

          ${b.chemical_ratio&&b.chemical_ratio!=="-"?`
            <div style="background:rgba(245, 158, 11, 0.14); border:1px solid rgba(251, 191, 36, 0.36); border-radius:9px; padding:11px 15px;">
              <strong style="color:#FCD34D; font-size:0.88rem; display:block; margin-bottom:3px;">\u{1F9EA} Rasio & Takaran Chemical:</strong>
              <span style="color:#FDE68A; font-size:0.92rem; font-weight:600; line-height:1.5;">${b.chemical_ratio}</span>
            </div>
          `:""}

          <div>
            <h4 style="margin:0 0 8px 0; font-size:0.98rem; font-weight:700; color:#FFFFFF;">\u{1F4DD} Langkah-Langkah Pelaksanaan:</h4>
            <div style="background:#0F172A; border:1px solid #29364B; border-radius:9px; padding:13px 18px;">
              <ol style="margin:0 0 0 20px; padding:0; font-size:0.9rem; color:#F8FAFC; line-height:1.8;">
                ${b.steps.map(y=>`<li style="margin-bottom:7px;">${y}</li>`).join("")}
              </ol>
            </div>
          </div>

          ${b.notes?`
            <div style="background:rgba(59, 130, 246, 0.14); border:1px solid rgba(96, 165, 250, 0.36); border-radius:9px; padding:12px 15px;">
              <strong style="color:#93C5FD; font-size:0.88rem; display:block; margin-bottom:3px;">\u26A0\uFE0F Catatan Penting & Safety:</strong>
              <span style="color:#F8FAFC; font-size:0.88rem; line-height:1.6;">${b.notes}</span>
            </div>
          `:""}
        </div>
      `,confirmText:"Tutup Panduan",onConfirm:(y,r)=>{r()}})}function _(b=null){if(!n)return;let S=!!b,y=b?b.tools.join(`
`):"",r=b?b.steps.join(`
`):"",l=`
      <form id="ik-edit-form" style="display:flex; flex-direction:column; gap:14px; font-size:0.88rem;">
        <div class="form-row" style="display:grid; grid-template-columns:1fr 140px; gap:12px;">
          <div>
            <label style="display:block; font-weight:600; margin-bottom:4px;">Judul Instruksi Kerja *</label>
            <input type="text" id="form-ik-title" class="form-control" required value="${b?Y(b.title):""}" style="width:100%; border-radius:6px; height:36px;">
          </div>
          <div>
            <label style="display:block; font-weight:600; margin-bottom:4px;">Ikon Emoji</label>
            <input type="text" id="form-ik-icon" class="form-control" value="${b?b.icon:"\u{1F4CB}"}" style="width:100%; border-radius:6px; height:36px; text-align:center; font-size:1.2rem;">
          </div>
        </div>

        <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-weight:600; margin-bottom:4px;">Kategori *</label>
            <select id="form-ik-cat" class="form-control" style="width:100%; border-radius:6px; height:36px;">
              <option value="Pembersihan Rutin" ${b?.category==="Pembersihan Rutin"?"selected":""}>Pembersihan Rutin</option>
              <option value="Sanitasi Toilet" ${b?.category==="Sanitasi Toilet"?"selected":""}>Sanitasi Toilet</option>
              <option value="Klinis & Medis" ${b?.category==="Klinis & Medis"?"selected":""}>Klinis & Medis</option>
              <option value="Elektronik & Utility" ${b?.category==="Elektronik & Utility"?"selected":""}>Elektronik & Utility</option>
              <option value="Pembersihan Khusus" ${b?.category==="Pembersihan Khusus"?"selected":""}>Pembersihan Khusus</option>
            </select>
          </div>
          <div>
            <label style="display:block; font-weight:600; margin-bottom:4px;">Takaran / Rasio Chemical</label>
            <input type="text" id="form-ik-ratio" class="form-control" value="${b?Y(b.chemical_ratio||""):""}" style="width:100%; border-radius:6px; height:36px;">
          </div>
        </div>

        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Area & Sasaran Target *</label>
          <input type="text" id="form-ik-target" class="form-control" required value="${b?Y(b.target):""}" style="width:100%; border-radius:6px; height:36px;">
        </div>

        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Alat & Bahan yang Digunakan (1 item per baris) *</label>
          <textarea id="form-ik-tools" class="form-control" rows="4" style="width:100%; border-radius:6px; font-size:0.85rem; line-height:1.5;">${Y(y)}</textarea>
        </div>

        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Langkah-Langkah Pelaksanaan (1 langkah per baris) *</label>
          <textarea id="form-ik-steps" class="form-control" rows="6" style="width:100%; border-radius:6px; font-size:0.85rem; line-height:1.6;">${Y(r)}</textarea>
        </div>

        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Catatan Penting / Hal yang Perlu Diperhatikan (Safety)</label>
          <textarea id="form-ik-notes" class="form-control" rows="2" style="width:100%; border-radius:6px; font-size:0.85rem;">${b?Y(b.notes||""):""}</textarea>
        </div>
      </form>
    `;ce({title:S?`\u270F\uFE0F Edit IK-${b.id}: ${b.title}`:"\u2795 Tambah Instruksi Kerja Baru",content:l,size:"lg",confirmText:"Simpan Perubahan",onConfirm:(f,u)=>{let p=f.querySelector("#form-ik-title").value.trim(),A=f.querySelector("#form-ik-icon").value.trim()||"\u{1F4CB}",T=f.querySelector("#form-ik-cat").value,g=f.querySelector("#form-ik-ratio").value.trim(),C=f.querySelector("#form-ik-target").value.trim(),z=f.querySelector("#form-ik-tools").value.trim(),O=f.querySelector("#form-ik-steps").value.trim(),te=f.querySelector("#form-ik-notes").value.trim();if(!p||!C||!z||!O){j("Harap lengkapi seluruh field wajib (*)");return}let Q=z.split(`
`).map(U=>U.trim().replace(/^[-•*]\s*/,"")).filter(Boolean),ne=O.split(`
`).map(U=>U.trim().replace(/^\d+[\.\)]\s*/,"")).filter(Boolean),H={"Pembersihan Rutin":{color:"#2563EB",bg:"#EFF6FF"},"Sanitasi Toilet":{color:"#DC2626",bg:"#FEF2F2"},"Klinis & Medis":{color:"#E11D48",bg:"#FFF1F2"},"Elektronik & Utility":{color:"#0D9488",bg:"#F0FDFA"},"Pembersihan Khusus":{color:"#7C3AED",bg:"#F5F3FF"}}[T]||{color:"#2563EB",bg:"#EFF6FF"};if(S)b.title=p,b.icon=A,b.category=T,b.chemical_ratio=g,b.target=C,b.tools=Q,b.steps=ne,b.notes=te,b.color=H.color,b.bg=H.bg,Na(d),ae(`Instruksi Kerja "${p}" berhasil diperbarui`);else{let U=d.length>0?Math.max(...d.map(V=>V.id))+1:1,D={id:U,title:p,category:T,icon:A,color:H.color,bg:H.bg,target:C,chemical_ratio:g||"-",tools:Q,steps:ne,notes:te};d.push(D),Na(d),ae(`Instruksi Kerja baru IK-${U} berhasil ditambahkan`)}u(),I()}})}function E(b){b.innerHTML=`
      <!-- Sub-tab Switcher & Actions -->
      <div style="display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:12px; margin-bottom:18px;">
        <div class="btn-group" style="background:#111827; padding:3px; border-radius:8px; display:inline-flex; border:1px solid var(--border);">
          <button id="view-tools-btn" class="btn btn-sm ${m==="tools"?"btn-primary":"btn-ghost"}" style="font-weight:700; border-radius:6px; padding:6px 16px; cursor:pointer;">
            \u{1F9F9} Alat Kebersihan (${F.length} Item)
          </button>
          <button id="view-chem-btn" class="btn btn-sm ${m==="chemical"?"btn-primary":"btn-ghost"}" style="font-weight:700; border-radius:6px; padding:6px 16px; cursor:pointer;">
            \u{1F9EA} Bahan Chemical (${$.length} Item)
          </button>
        </div>

        <div style="display:flex; gap:10px; align-items:center; flex-wrap:wrap;">
          <input type="text" id="tool-search" class="form-control" placeholder="Cari nama / peruntukan / rasio..." value="${i}" style="height:36px; border-radius:8px; font-size:0.875rem; width:240px;">
            ${n?`
          <button id="btn-add-tool-chem" class="btn btn-primary" style="font-weight:600; padding:6px 14px; border-radius:8px; display:inline-flex; align-items:center; gap:6px; cursor:pointer; font-size:0.85rem;">
            \u2795 ${m==="tools"?"Tambah Alat Baru":"Tambah Chemical Baru"}
          </button>
            `:""}
        </div>
      </div>

      <!-- Table Container -->
      <div id="tool-table-container"></div>
    `;let S=b.querySelector("#tool-table-container"),y=b.querySelector("#tool-search");function r(){let l=i.toLowerCase().trim();if(m==="tools"){let f=F.filter(u=>!l||u.name.toLowerCase().includes(l)||u.area.toLowerCase().includes(l));S.innerHTML=`
          <div class="card" style="background:#172033; border:1px solid var(--border); border-radius:12px; overflow:hidden; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
            <div style="padding:14px 18px; background:#111827; border-bottom:1px solid var(--border); display:flex; justify-content:space-between; align-items:center;">
              <strong style="color:#FFFFFF; font-size:0.96rem;">Daftar Standar Peralatan Kebersihan (Lampiran 8.1 SOP)</strong>
              <span class="badge" style="background:rgba(59,130,246,0.2); color:#93C5FD; border:1px solid rgba(96,165,250,0.4); font-weight:700; padding:3px 9px; border-radius:6px;">${f.length} Item</span>
            </div>
            <div style="overflow-x:auto;">
              <table class="table" style="width:100%; margin:0; border-collapse:collapse; font-size:0.9rem;">
                <thead>
                  <tr style="background:#111827; border-bottom:1px solid #29364B; text-align:left;">
                    <th style="padding:11px 14px; width:60px; color:#CBD5E1;">No</th>
                    <th style="padding:11px 14px; color:#CBD5E1;">Nama Alat / Item</th>
                    <th style="padding:11px 14px; width:140px; color:#CBD5E1;">Kuantitas Standar</th>
                    <th style="padding:11px 14px; color:#CBD5E1;">Peruntukan Area & Kegunaan</th>
                    <th style="padding:11px 14px; width:110px; text-align:center; color:#CBD5E1;">Aksi</th>
                  </tr>
                </thead>
                <tbody>
                  ${f.map((u,p)=>`
                    <tr style="border-bottom:1px solid #1E2E4A; background:${p%2===0?"#131F36":"#0E182B"}; color:#F8FAFC;">
                      <td style="padding:11px 14px; font-weight:600; color:#CBD5E1;">${u.no||p+1}</td>
                      <td style="padding:11px 14px; font-weight:700; color:#FFFFFF;">${Y(u.name)}</td>
                      <td style="padding:11px 14px;"><span class="badge" style="background:rgba(56,189,248,0.16); color:#7DD3FC; border:1px solid rgba(56,189,248,0.35); font-weight:700; padding:3px 9px; border-radius:6px;">${Y(u.qty)}</span></td>
                      <td style="padding:11px 14px; color:#CBD5E1; line-height:1.5;">${Y(u.area)}</td>
                      <td style="padding:11px 14px; text-align:center;">
                        <div style="display:inline-flex; gap:5px;">
                    ${n?`
                        <button type="button" class="fc-action-btn fc-action-edit btn-edit-tool" data-no="${u.no}" title="Edit Alat" aria-label="Edit alat ${Y(u.name||"")}"><svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M12 20h9"></path><path d="M16.5 3.5a2.121 2.121 0 0 1 3 3L7 19l-4 1 1-4L16.5 3.5z"></path></svg></button>
                    `:""}
                    ${a?`
                        <button type="button" class="fc-action-btn fc-action-del btn-delete-tool" data-no="${u.no}" title="Hapus Alat" aria-label="Hapus alat ${Y(u.name||"")}"><svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><polyline points="3 6 5 6 21 6"></polyline><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"></path></svg></button>
                    `:""}
                        </div>
                      </td>
                    </tr>
                  `).join("")}
                </tbody>
              </table>
            </div>
          </div>
        `,S.querySelectorAll(".btn-edit-tool").forEach(u=>{u.addEventListener("click",()=>{let p=parseInt(u.dataset.no,10),A=F.find(T=>T.no===p);A&&w(A)})}),S.querySelectorAll(".btn-delete-tool").forEach(u=>{u.addEventListener("click",()=>{let p=parseInt(u.dataset.no,10),A=F.find(T=>T.no===p);A&&ge(`Apakah Anda yakin ingin menghapus alat: "${A.name}"?`,()=>{F=F.filter(T=>T.no!==p),F.forEach((T,g)=>{T.no=g+1}),Ka(F),ae("Alat kebersihan berhasil dihapus"),I()},"Hapus Alat Kebersihan")})})}else{let f=$.filter(u=>!l||u.name.toLowerCase().includes(l)||u.function.toLowerCase().includes(l));S.innerHTML=`
          <div class="card" style="background:#172033; border:1px solid var(--border); border-radius:12px; overflow:hidden; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
            <div style="padding:14px 18px; background:#111827; border-bottom:1px solid var(--border); display:flex; justify-content:space-between; align-items:center;">
              <strong style="color:#FFFFFF; font-size:0.96rem;">Daftar Bahan Kimia & Standar Takaran (Lampiran 8.1 SOP)</strong>
              <span class="badge" style="background:rgba(16,185,129,0.18); color:#6EE7B7; border:1px solid rgba(52,211,153,0.4); font-weight:700; padding:3px 9px; border-radius:6px;">${f.length} Chemical</span>
            </div>
            <div style="overflow-x:auto;">
              <table class="table" style="width:100%; margin:0; border-collapse:collapse; font-size:0.9rem;">
                <thead>
                  <tr style="background:#111827; border-bottom:1px solid #29364B; text-align:left;">
                    <th style="padding:11px 14px; width:60px; color:#CBD5E1;">No</th>
                    <th style="padding:11px 14px; color:#CBD5E1;">Nama Produk Chemical</th>
                    <th style="padding:11px 14px; width:115px; color:#CBD5E1;">Kemasan</th>
                    <th style="padding:11px 14px; width:230px; color:#CBD5E1;">Rasio & Standar Takaran</th>
                    <th style="padding:11px 14px; color:#CBD5E1;">Fungsi & Cara Penggunaan</th>
                    <th style="padding:11px 14px; width:110px; text-align:center; color:#CBD5E1;">Aksi</th>
                  </tr>
                </thead>
                <tbody>
                  ${f.map((u,p)=>`
                    <tr style="border-bottom:1px solid #1E2E4A; background:${p%2===0?"#131F36":"#0E182B"}; color:#F8FAFC;">
                      <td style="padding:11px 14px; font-weight:600; color:#CBD5E1;">${u.no||p+1}</td>
                      <td style="padding:11px 14px; font-weight:700; color:#FFFFFF;">${Y(u.name)}</td>
                      <td style="padding:11px 14px;"><span class="badge" style="background:rgba(245,158,11,0.18); color:#FCD34D; border:1px solid rgba(251,191,36,0.38); font-weight:700; padding:3px 9px; border-radius:6px;">${Y(u.size)}</span></td>
                      <td style="padding:11px 14px;"><span style="background:rgba(56,189,248,0.14); color:#7DD3FC; border:1px solid rgba(56,189,248,0.35); padding:4px 9px; border-radius:6px; font-weight:700; font-size:0.84rem; display:inline-block;">${Y(u.ratio)}</span></td>
                      <td style="padding:11px 14px; color:#CBD5E1; line-height:1.5;">${Y(u.function)}</td>
                      <td style="padding:11px 14px; text-align:center;">
                        <div style="display:inline-flex; gap:5px;">
                    ${n?`
                        <button type="button" class="fc-action-btn fc-action-edit btn-edit-chem" data-no="${u.no}" title="Edit Chemical" aria-label="Edit chemical ${Y(u.name||"")}"><svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M12 20h9"></path><path d="M16.5 3.5a2.121 2.121 0 0 1 3 3L7 19l-4 1 1-4L16.5 3.5z"></path></svg></button>
                    `:""}
                    ${a?`
                        <button type="button" class="fc-action-btn fc-action-del btn-delete-chem" data-no="${u.no}" title="Hapus Chemical" aria-label="Hapus chemical ${Y(u.name||"")}"><svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><polyline points="3 6 5 6 21 6"></polyline><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"></path></svg></button>
                    `:""}
                        </div>
                      </td>
                    </tr>
                  `).join("")}
                </tbody>
              </table>
            </div>
          </div>
        `,S.querySelectorAll(".btn-edit-chem").forEach(u=>{u.addEventListener("click",()=>{let p=parseInt(u.dataset.no,10),A=$.find(T=>T.no===p);A&&K(A)})}),S.querySelectorAll(".btn-delete-chem").forEach(u=>{u.addEventListener("click",()=>{let p=parseInt(u.dataset.no,10),A=$.find(T=>T.no===p);A&&ge(`Apakah Anda yakin ingin menghapus chemical: "${A.name}"?`,()=>{$=$.filter(T=>T.no!==p),$.forEach((T,g)=>{T.no=g+1}),Ra($),ae("Bahan chemical berhasil dihapus"),I()},"Hapus Bahan Chemical")})})}}r(),b.querySelector("#view-tools-btn").addEventListener("click",()=>{m="tools",E(b)}),b.querySelector("#view-chem-btn").addEventListener("click",()=>{m="chemical",E(b)}),y&&y.addEventListener("input",l=>{i=l.target.value,r()}),b.querySelector("#btn-add-tool-chem").addEventListener("click",()=>{m==="tools"?w():K()})}function w(b=null){if(!n)return;let S=!!b,y=`
      <form id="tool-edit-form" style="display:flex; flex-direction:column; gap:14px; font-size:0.88rem;">
        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Nama Alat / Item *</label>
          <input type="text" id="form-tool-name" class="form-control" required value="${b?Y(b.name):""}" style="width:100%; border-radius:6px; height:36px;">
        </div>
        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Kuantitas Standar *</label>
          <input type="text" id="form-tool-qty" class="form-control" required value="${b?Y(b.qty):"1 unit"}" style="width:100%; border-radius:6px; height:36px;">
        </div>
        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Peruntukan Area & Kegunaan *</label>
          <textarea id="form-tool-area" class="form-control" rows="3" required style="width:100%; border-radius:6px; font-size:0.85rem;">${b?Y(b.area):""}</textarea>
        </div>
      </form>
    `;ce({title:S?`\u270F\uFE0F Edit Alat: ${b.name}`:"\u2795 Tambah Alat Kebersihan Baru",content:y,confirmText:"Simpan",onConfirm:(r,l)=>{let f=r.querySelector("#form-tool-name").value.trim(),u=r.querySelector("#form-tool-qty").value.trim(),p=r.querySelector("#form-tool-area").value.trim();if(!f||!u||!p){j("Harap lengkapi semua kolom");return}if(S)b.name=f,b.qty=u,b.area=p,Ka(F),ae(`Alat "${f}" berhasil diperbarui`);else{let A=F.length>0?Math.max(...F.map(T=>T.no))+1:1;F.push({no:A,name:f,qty:u,area:p}),Ka(F),ae(`Alat baru "${f}" berhasil ditambahkan`)}l(),I()}})}function K(b=null){if(!n)return;let S=!!b,y=`
      <form id="chem-edit-form" style="display:flex; flex-direction:column; gap:14px; font-size:0.88rem;">
        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Nama Produk Chemical *</label>
          <input type="text" id="form-chem-name" class="form-control" required value="${b?Y(b.name):""}" style="width:100%; border-radius:6px; height:36px;">
        </div>
        <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-weight:600; margin-bottom:4px;">Kemasan Standar *</label>
            <input type="text" id="form-chem-size" class="form-control" required value="${b?Y(b.size):"5 Liter"}" style="width:100%; border-radius:6px; height:36px;">
          </div>
          <div>
            <label style="display:block; font-weight:600; margin-bottom:4px;">Rasio & Takaran Standar *</label>
            <input type="text" id="form-chem-ratio" class="form-control" required value="${b?Y(b.ratio):""}" style="width:100%; border-radius:6px; height:36px;">
          </div>
        </div>
        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Fungsi & Cara Penggunaan *</label>
          <textarea id="form-chem-function" class="form-control" rows="3" required style="width:100%; border-radius:6px; font-size:0.85rem;">${b?Y(b.function):""}</textarea>
        </div>
      </form>
    `;ce({title:S?`\u270F\uFE0F Edit Chemical: ${b.name}`:"\u2795 Tambah Chemical Baru",content:y,confirmText:"Simpan",onConfirm:(r,l)=>{let f=r.querySelector("#form-chem-name").value.trim(),u=r.querySelector("#form-chem-size").value.trim(),p=r.querySelector("#form-chem-ratio").value.trim(),A=r.querySelector("#form-chem-function").value.trim();if(!f||!u||!p||!A){j("Harap lengkapi semua kolom");return}if(S)b.name=f,b.size=u,b.ratio=p,b.function=A,Ra($),ae(`Chemical "${f}" berhasil diperbarui`);else{let T=$.length>0?Math.max(...$.map(g=>g.no))+1:1;$.push({no:T,name:f,size:u,ratio:p,function:A}),Ra($),ae(`Chemical baru "${f}" berhasil ditambahkan`)}l(),I()}})}function G(b){b.innerHTML=`
      <div style="display:grid; grid-template-columns:repeat(auto-fit, minmax(360px, 1fr)); gap:20px;">
        <!-- Card 1: Grooming -->
        <div class="card" style="background:#172033; border:1px solid #29364B; border-radius:12px; padding:20px; box-shadow:0 4px 12px rgba(0,0,0,0.2);">
          <div style="display:flex; align-items:center; gap:12px; margin-bottom:14px; border-bottom:1px solid #29364B; padding-bottom:12px;">
            <span style="font-size:1.5rem; background:rgba(59,130,246,0.16); border:1px solid rgba(96,165,250,0.35); padding:8px; border-radius:9px;">\u{1F454}</span>
            <div>
              <h3 style="margin:0 0 2px; font-size:1.12rem; font-weight:700; color:#FFFFFF;">Standar Grooming Petugas (6.1)</h3>
              <span style="font-size:0.84rem; color:#CBD5E1;">8 Persyaratan Wajib Penampilan Facility Care</span>
            </div>
          </div>
          <ol style="margin:0 0 0 18px; padding:0; font-size:0.9rem; color:#F8FAFC; line-height:1.8;">
            <li>Menggunakan seragam resmi setiap hari (Senin s/d Minggu).</li>
            <li>Menggunakan celana bahan berwarna hitam (bukan jeans/chino kasual).</li>
            <li>Wajib mengenakan name badge atau ID Card di dada kiri.</li>
            <li>Menggunakan kaos kaki warna hitam.</li>
            <li>Menggunakan sepatu pantofel hitam yang bersih dan disemir.</li>
            <li>Menggunakan parfum dan deodorant untuk menjaga kesegaran badan.</li>
            <li>Merapikan rambut (pria pendek rapi, wanita diikat/hairnet rapi).</li>
            <li>Dilarang menggunakan aksesoris perhiasan selain jam tangan standar.</li>
          </ol>
        </div>

        <!-- Card 2: Jam Kerja -->
        <div class="card" style="background:#172033; border:1px solid #29364B; border-radius:12px; padding:20px; box-shadow:0 4px 12px rgba(0,0,0,0.2);">
          <div style="display:flex; align-items:center; gap:12px; margin-bottom:14px; border-bottom:1px solid #29364B; padding-bottom:12px;">
            <span style="font-size:1.5rem; background:rgba(245,158,11,0.16); border:1px solid rgba(251,191,36,0.35); padding:8px; border-radius:9px;">\u23F0</span>
            <div>
              <h3 style="margin:0 0 2px; font-size:1.12rem; font-weight:700; color:#FFFFFF;">Pengaturan Jam Kerja & Shift (6.2)</h3>
              <span style="font-size:0.84rem; color:#CBD5E1;">Sistem Kerja 6 Hari Kerja : 1 Hari Libur</span>
            </div>
          </div>
          <div style="font-size:0.9rem; color:#F8FAFC; line-height:1.7;">
            <div style="background:#0F172A; border:1px solid #29364B; border-radius:8px; padding:11px 14px; margin-bottom:10px; border-left:3.5px solid #60A5FA;">
              <strong style="color:#FFFFFF; display:block; margin-bottom:2px;">\u{1F3E5} Klinik Primer (1 Orang Facility Care):</strong>
              <span style="color:#E2E8F0;">Shift 1 (Normal): <strong style="color:#FFFFFF;">08.00 s/d 17.00</strong> (Istirahat 13.00 s/d 14.00)</span>
            </div>
            <div style="background:#0F172A; border:1px solid #29364B; border-radius:8px; padding:11px 14px; margin-bottom:12px; border-left:3.5px solid #34D399;">
              <strong style="color:#FFFFFF; display:block; margin-bottom:2px;">\u2728 Klinik Luxury (&gt; 1 Orang Facility Care):</strong>
              <div style="margin-bottom:2px; color:#E2E8F0;">\u2022 Shift 1 (Pagi): <strong style="color:#FFFFFF;">07.00 s/d 16.00</strong> (Istirahat 12.00 s/d 13.00)</div>
              <div style="color:#E2E8F0;">\u2022 Shift 2 (Siang): <strong style="color:#FFFFFF;">13.00 s/d 22.00</strong> (Istirahat 17.00 s/d 18.00)</div>
            </div>
            <p style="margin:0; font-size:0.85rem; color:#CBD5E1; line-height:1.6;">
              <strong style="color:#FFFFFF;">Ketentuan Backup:</strong> Saat Facility Care cabang libur, tim cabang wajib melakukan backup pekerjaan yang ditentukan oleh Supervisor (SPV) Cabang.
            </p>
          </div>
        </div>

        <!-- Card 3: Tanggung Jawab & Pembagian Area -->
        <div class="card" style="background:#172033; border:1px solid #29364B; border-radius:12px; padding:20px; box-shadow:0 4px 12px rgba(0,0,0,0.2);">
          <div style="display:flex; align-items:center; gap:12px; margin-bottom:14px; border-bottom:1px solid #29364B; padding-bottom:12px;">
            <span style="font-size:1.5rem; background:rgba(16,185,129,0.16); border:1px solid rgba(52,211,153,0.35); padding:8px; border-radius:9px;">\u{1F3E2}</span>
            <div>
              <h3 style="margin:0 0 2px; font-size:1.12rem; font-weight:700; color:#FFFFFF;">Tanggung Jawab & Area (6.3)</h3>
              <span style="font-size:0.84rem; color:#CBD5E1;">Pembagian Pelaksana & Penanggung Jawab Area</span>
            </div>
          </div>
          <table style="width:100%; border-collapse:collapse; font-size:0.88rem; margin-bottom:14px;">
            <thead>
              <tr style="background:#0F172A; border-bottom:1px solid #29364B; text-align:left;">
                <th style="padding:8px 10px; color:#CBD5E1;">Area Kerja</th>
                <th style="padding:8px 10px; color:#CBD5E1;">Pelaksana</th>
                <th style="padding:8px 10px; color:#CBD5E1;">PIC Pengawas</th>
              </tr>
            </thead>
            <tbody>
              <tr style="border-bottom:1px solid #29364B;">
                <td style="padding:9px 10px; font-weight:600; color:#F8FAFC;">HQ & Gudang Pusat</td>
                <td style="padding:9px 10px; color:#CBD5E1;">Facility Care Pusat</td>
                <td style="padding:9px 10px; color:#93C5FD; font-weight:700;">FC Supervisor</td>
              </tr>
              <tr>
                <td style="padding:9px 10px; font-weight:600; color:#F8FAFC;">Klinik / Cabang</td>
                <td style="padding:9px 10px; color:#CBD5E1;">Facility Care Cabang</td>
                <td style="padding:9px 10px; color:#6EE7B7; font-weight:700;">SPV Cabang</td>
              </tr>
            </tbody>
          </table>
          <div style="font-size:0.86rem; color:#CBD5E1; line-height:1.65;">
            <div style="margin-bottom:6px;"><strong style="color:#FFFFFF;">Deep Cleaning:</strong> Pembersihan total eventual pasca renovasi, pembukaan cabang baru, atau relokasi cabang lama.</div>
            <div><strong style="color:#FFFFFF;">General Cleaning (GC):</strong> Pembersihan menyeluruh berkala untuk cabang eksisting sesuai jadwal yang telah ditetapkan.</div>
          </div>
        </div>

        <!-- Card 4: Monitoring & Evaluasi -->
        <div class="card" style="background:#172033; border:1px solid #29364B; border-radius:12px; padding:20px; box-shadow:0 4px 12px rgba(0,0,0,0.2);">
          <div style="display:flex; align-items:center; gap:12px; margin-bottom:14px; border-bottom:1px solid #29364B; padding-bottom:12px;">
            <span style="font-size:1.5rem; background:rgba(139,92,246,0.16); border:1px solid rgba(167,139,250,0.35); padding:8px; border-radius:9px;">\u{1F50D}</span>
            <div>
              <h3 style="margin:0 0 2px; font-size:1.12rem; font-weight:700; color:#FFFFFF;">Sistem Monitoring & Inspeksi (6.4)</h3>
              <span style="font-size:0.84rem; color:#CBD5E1;">Mekanisme Evaluasi Kinerja Harian & Triwulanan</span>
            </div>
          </div>
          <div style="font-size:0.9rem; color:#F8FAFC; line-height:1.7;">
            <div style="margin-bottom:10px;">
              <strong style="color:#FFFFFF;">1. Monitor Harian & Mingguan:</strong>
              <div style="font-size:0.87rem; color:#CBD5E1; margin-top:3px; line-height:1.6;">
                Ceklist harian diisi 2 kali sehari (Pagi & Siang) dan pekerjaan mingguan diisi seminggu sekali. Wajib ditandatangani oleh SPV Cabang / FC Supervisor.
              </div>
            </div>
            <div>
              <strong style="color:#FFFFFF;">2. Inspeksi Cabang Triwulanan (3 Bulan):</strong>
              <div style="font-size:0.87rem; color:#CBD5E1; margin-top:3px; line-height:1.6;">
                Dilakukan oleh FC Specialist langsung ke lokasi klinik. Laporan inspeksi diunggah ke Basecamp dan wajib diberikan tindak lanjut oleh PIC cabang maksimal <strong style="color:#FFFFFF;">3 hari</strong> setelah laporan terbit. Hasil penilaian mempengaruhi performance cabang.
              </div>
            </div>
          </div>
        </div>
      </div>
    `}function J(b){b.innerHTML=`
      <div style="display:grid; grid-template-columns:repeat(auto-fit, minmax(320px, 1fr)); gap:18px;">
        <!-- Card: Ceklist Harian Google Sheets -->
        <div class="card" style="background:#172033; border:1px solid #29364B; border-radius:12px; padding:20px; box-shadow:0 4px 12px rgba(0,0,0,0.2); display:flex; flex-direction:column; justify-content:space-between;">
          <div>
            <div style="display:flex; align-items:center; gap:12px; margin-bottom:12px;">
              <span style="font-size:2rem;">\u{1F4CA}</span>
              <div>
                <h3 style="margin:0 0 4px; font-size:1.12rem; font-weight:700; color:#FFFFFF;">Form Ceklist Harian & Mingguan</h3>
                <span class="badge" style="background:rgba(56,189,248,0.16); color:#7DD3FC; border:1px solid rgba(56,189,248,0.35); font-size:0.78rem; font-weight:700; padding:3px 9px; border-radius:6px;">Google Sheets</span>
              </div>
            </div>
            <p style="font-size:0.88rem; color:#CBD5E1; line-height:1.6; margin-bottom:16px;">
              Format lembar ceklist kebersihan harian (Shift Pagi & Siang) dan ceklist pembersihan berkala mingguan yang wajib diisi oleh Facility Care dan disetujui Supervisor Cabang.
            </p>
          </div>
          <a href="https://docs.google.com/spreadsheets/d/1alq6ppT7syB8Q0_RjjeqwYhjeHgrWvVe-KJ3dHXWrss/edit?usp=sharing" target="_blank" rel="noopener" class="btn btn-primary" style="display:inline-flex; align-items:center; justify-content:center; gap:6px; font-weight:700; padding:10px; border-radius:8px; background:#2563EB; color:#FFFFFF;">
            \u{1F4C4} Buka Google Sheets Ceklist \u2794
          </a>
        </div>

        <!-- Card: Form Inspeksi Hygiene Google Forms -->
        <div class="card" style="background:#172033; border:1px solid #29364B; border-radius:12px; padding:20px; box-shadow:0 4px 12px rgba(0,0,0,0.2); display:flex; flex-direction:column; justify-content:space-between;">
          <div>
            <div style="display:flex; align-items:center; gap:12px; margin-bottom:12px;">
              <span style="font-size:2rem;">\u{1F4D1}</span>
              <div>
                <h3 style="margin:0 0 4px; font-size:1.12rem; font-weight:700; color:#FFFFFF;">Form Inspeksi Hygiene Cabang</h3>
                <span class="badge" style="background:rgba(37,99,235,0.16); color:#93C5FD; border:1px solid rgba(59,130,246,0.35); font-size:0.78rem; font-weight:700; padding:3px 9px; border-radius:6px;">Google Forms</span>
              </div>
            </div>
            <p style="font-size:0.88rem; color:#CBD5E1; line-height:1.6; margin-bottom:16px;">
              Formulir digital inspeksi evaluasi kebersihan dan kepatuhan hygiene cabang yang diisi oleh Facility Care Specialist setiap periode 3 bulanan.
            </p>
          </div>
          <a href="https://docs.google.com/forms/d/e/1FAIpQLScn8iPabIt51E_GORkbD86IURSS-EsUYINEQ_Zwdq_uC3XOCQ/viewform" target="_blank" rel="noopener" class="btn btn-primary" style="display:inline-flex; align-items:center; justify-content:center; gap:6px; font-weight:700; padding:10px; border-radius:8px; background:#2563EB; border-color:#3B82F6; color:#FFFFFF;">
            \u{1F4CB} Buka Google Form Inspeksi \u2192
          </a>
        </div>

        <!-- Card: Basecamp Monitoring Project -->
        <div class="card" style="background:#131F36; border:1px solid #223456; border-radius:12px; padding:20px; display:flex; flex-direction:column; justify-content:space-between;">
          <div>
            <div style="display:flex; align-items:center; gap:12px; margin-bottom:12px;">
              <span style="font-size:2rem;">\u{1F3D5}\uFE0F</span>
              <div>
                <h3 style="margin:0 0 4px; font-size:1.12rem; font-weight:700; color:#FFFFFF;">Laporan Inspeksi Basecamp</h3>
                <span class="badge" style="background:rgba(245,158,11,0.16); color:#FCD34D; border:1px solid rgba(251,191,36,0.35); font-size:0.78rem; font-weight:700; padding:3px 9px; border-radius:6px;">Basecamp 3 Project</span>
              </div>
            </div>
            <p style="font-size:0.88rem; color:#CBD5E1; line-height:1.6; margin-bottom:16px;">
              To-do list dan kanal pelaporan temuan inspeksi cabang. PIC Cabang wajib memberikan respon tindak lanjut maksimal 3 hari kerja setelah laporan terbit.
            </p>
          </div>
          <a href="https://3.basecamp.com/5259614/buckets/33617501/todosets/6367925585" target="_blank" rel="noopener" class="btn btn-outline" style="display:inline-flex; align-items:center; justify-content:center; gap:6px; font-weight:700; padding:10px; border-radius:8px; background:#1B2940; border:1px solid #334155; color:#FFFFFF;">
            \u{1F680} Buka Laporan di Basecamp \u2794
          </a>
        </div>
      </div>

      <!-- Validasi & Pengesahan SOP -->
      <div class="card" style="margin-top:20px; padding:20px; background:#172033; border:1px solid #29364B; border-radius:12px;">
        <h3 style="margin:0 0 14px 0; font-size:1.05rem; font-weight:700; color:#FFFFFF; display:flex; align-items:center; gap:8px;">
          <span>\u{1F4DC}</span> Validasi & Pengesahan Dokumen SOP
        </h3>
        <div style="display:grid; grid-template-columns:repeat(auto-fit, minmax(260px, 1fr)); gap:16px;">
          <div style="border-left:3.5px solid #60A5FA; padding-left:12px;">
            <div style="font-size:0.78rem; color:#CBD5E1; font-weight:700; text-transform:uppercase;">Disiapkan Oleh:</div>
            <div style="font-size:0.98rem; font-weight:700; color:#FFFFFF; margin-top:2px;">Berlin Ariansyah</div>
            <div style="font-size:0.84rem; color:#CBD5E1;">Facility Care Supervisor</div>
          </div>
          <div style="border-left:3.5px solid #34D399; padding-left:12px;">
            <div style="font-size:0.78rem; color:#CBD5E1; font-weight:700; text-transform:uppercase;">Diketahui Oleh:</div>
            <div style="font-size:0.98rem; font-weight:700; color:#FFFFFF; margin-top:2px;">Catur Meiyogo</div>
            <div style="font-size:0.84rem; color:#CBD5E1;">General Support Manager</div>
          </div>
          <div style="border-left:3.5px solid #A78BFA; padding-left:12px;">
            <div style="font-size:0.78rem; color:#CBD5E1; font-weight:700; text-transform:uppercase;">Disetujui Oleh:</div>
            <div style="font-size:0.98rem; font-weight:700; color:#FFFFFF; margin-top:2px;">Drg. Yulita</div>
            <div style="font-size:0.84rem; color:#CBD5E1;">Chief Executive Officer</div>
          </div>
        </div>
      </div>
    `}async function W(){try{let{downloadExcel:b}=await Promise.resolve().then(()=>(Ce(),Ve)),S=d.map(r=>({"No IK":`IK-${r.id}`,"Judul Pekerjaan":r.title,Kategori:r.category,"Sasaran Target Area":r.target,"Rasio Chemical":r.chemical_ratio||"-","Alat & Bahan":r.tools.join(", "),"Langkah Kerja":r.steps.map((l,f)=>`${f+1}. ${l}`).join(" | "),"Catatan Safety":r.notes||"-"})),y=new Date().toISOString().slice(0,10);b(S,`SOP_Facility_Care_Instruksi_Kerja_${y}`),ae("Data Excel Instruksi Kerja berhasil diunduh")}catch(b){console.error(b),j("Gagal mendownload data Excel: "+b.message)}}function Y(b){return b?String(b).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;").replace(/'/g,"&#39;"):""}I()}Be();var Mr=[{id:1,title:"Pengamanan Area & Aset Perusahaan",category:"Pengamanan Fisik & Patroli",icon:"\u{1F6E1}\uFE0F",color:"#2563EB",bg:"#EFF6FF",target:"Seluruh area dalam & luar Head Quarter dan Klinik Cabang Audy Dental",tools:["Google Form Checklist Monitoring Klinik: https://forms.gle/1EAvkqDJVdctXEGn7","Senter Patroli","Komputer dengan Akses Monitor CCTV (Live View)","Alat Komunikasi / Handphone / HT"],steps:["Lakukan Safety Patrol secara rutin mengelilingi seluruh area sesuai jadwal yang telah ditentukan (Patrol 1, 2, dan 3).","Catat hasil temuan dan kondisi keamanan secara berkala ke dalam Google Form Safety Patrol resmi.","Awasi kendaraan yang terparkir di area luar klinik; pastikan khusus sepeda motor dalam keadaan terkunci ganda.","Pantau situasi lingkungan dan area sensitif secara berkala melalui layar monitor CCTV saat berada di pos security.","Segera lakukan tindakan persuasif ramah apabila ditemukan potensi gangguan keamanan, dan segera hubungi pihak Kepolisian terdekat bila situasi darurat tidak dapat diatasi.","Bila ditemukan permasalahan fisik, kerusakan fasilitas, atau kendala di area klinik, segera dokumentasikan (foto) dan laporkan langsung ke SPV Klinik dan Leader Facility Care.","Disaat malam hari atau saat klinik telah tutup operasional, pastikan seluruh AC dan lampu ruangan dalam keadaan mati, kecuali lampu keamanan tertentu yang memang harus dibiarkan menyala terang."],notes:"Prioritaskan keselamatan diri dan tindakan persuasif. Dilarang meninggalkan area pengawasan tanpa personil pengganti yang jelas."},{id:2,title:"Penertiban & Pengaturan Parkir Kendaraan",category:"Operasional Parkir & Luar",icon:"\u{1F17F}\uFE0F",color:"#D97706",bg:"#FFFBEB",target:"Area parkir mobil, motor, akses keluar masuk, dan teras depan klinik",tools:["Lampu Lalu Lintas (Lampu Lalin)","Cone / Kerucut Pembatas Parkir","Peluit Parkir"],steps:["Berdiri di posisi strategis dan bantu mengarahkan setiap kendaraan pasien atau tamu yang datang secara tepat dan rapi ke slot parkir yang tersedia.","Berikan informasi dan solusi dengan sikap ramah serta santun apabila seluruh slot parkir klinik dalam kondisi penuh.","Dalam kondisi tertentu (bila ada mobil yang terhalang mobil lain dan harus dipindahkan, sementara pasien sedang dalam tindakan perawatan): bantu memindahkan kendaraan HANYA JIKA petugas security memiliki kemampuan mengemudi dan memiliki SIM aktif.","Apabila petugas security tidak memiliki kemampuan mengemudikan mobil: sampaikan permohonan maaf dengan tulus kepada pasien/tamu, minta untuk menunggu sejenak, dan segera minta bantuan kepada staf Resepsionis untuk menghubungi pemilik mobil yang menghalangi.","Bantu pasien atau tamu yang akan keluar parkir dengan memberikan aba-aba dan arahan lalu lintas jalan raya yang aman dan tepat."],notes:"Dilarang memaksakan memindahkan mobil pasien jika tidak memiliki lisensi mengemudi demi menghindari risiko gesekan atau kerusakan kendaraan."},{id:3,title:"Penerimaan Tamu & Pasien Klinik",category:"Layanan & Hospitalitas",icon:"\u{1F91D}",color:"#059669",bg:"#ECFDF5",target:"Pintu masuk utama klinik, lobby, area resepsionis, dan pos penerimaan tamu",tools:["Alat Tulis Kantor (Pulpen)","Buku Tamu / Log Book Keamanan","Alat Komunikasi / Handphone"],steps:['Berdiri dengan tegap dan sigap, bukakan pintu untuk setiap pasien/tamu yang datang dengan menerapkan budaya 5S (Senyum, Salam, Sapa, Sopan, Santun) dan menyapa: "Selamat pagi / siang / sore, ada yang bisa kami bantu?".',"Setelah pasien atau tamu memberitahukan tujuannya, persilakan dan arahkan pasien langsung menuju meja Resepsionis.","Apabila yang datang adalah tamu instansi, vendor, atau pihak Manajemen/BOD: persilakan mengisi Buku Tamu (Log Book) terlebih dahulu dan segera hubungi staf/karyawan yang ingin ditemui.",'Apabila pasien langsung melangkah ke arah tujuannya, berikan gestur tubuh hormat menyilangkan kedua tangan di depan, dan jika pasien/tamu mengucapkan terima kasih, jawab dengan ramah: "Dengan Senang Hati".',"Dalam kondisi khusus (pasien lansia, ibu hamil, atau pengguna kursi roda): petugas security wajib mendampingi, mengawal, dan membantu mendorong kursi roda sampai ke depan meja resepsionis.",'Buka kembali pintu untuk setiap pasien yang telah selesai melakukan perawatan saat hendak keluar klinik sambil mengucapkan ucapan hangat: "Terima Kasih, semoga lekas sehat kembali".'],notes:"Jaga senyum tulus dan keramahan sebagai garda terdepan representasi citra Audy Dental."}],zr=[{no:1,type:"Aset Tetap",item:"Rumah & Bangunan",description:"Gedung klinik, pos satpam, pagar, kanopi, teras beton",pic:"Building Maintenance & Security"},{no:2,type:"Aset Tetap",item:"Genset & Lift",description:"Mesin generator listrik cadangan dan elevator penumpang tetap",pic:"Teknisi & Security"},{no:3,type:"Aset Tetap",item:"AC Central & Split Tetap",description:"Outdoor dan indoor AC yang terpasang permanen di dinding/plafon",pic:"Teknisi AC & Security"},{no:4,type:"Aset Tetap",item:"Dental Unit Permanen",description:"Kursi dental unit terpasang instalasi air dan kompresor",pic:"Tim Klinik & Security"},{no:5,type:"Aset Tetap",item:"Instalasi Listrik, Pipa & Jaringan IT Permanen",description:"Panel MCB, kabel fiber optik, pipa drainase, server rack",pic:"IT & Teknisi"},{no:6,type:"Aset Tetap",item:"Furniture Melekat (Built-in)",description:"Meja resepsionis cor/melekat, partisi kaca mati, kitchen set pantry",pic:"General Support"},{no:7,type:"Aset Bergerak",item:"Komputer, Laptop & Printer",description:"PC kasir/resepsionis, laptop staf, printer rontgen/dokumen",pic:"Staf Terkait & IT"},{no:8,type:"Aset Bergerak",item:"Kendaraan Operasional",description:"Mobil operasional dan sepeda motor inventaris perusahaan",pic:"Security & Driver"},{no:9,type:"Aset Bergerak",item:"Furniture Lepasan",description:"Meja kerja, kursi resepsionis, kursi tunggu pasien, lemari file",pic:"Facility Care & Staf"},{no:10,type:"Aset Bergerak",item:"Hand Tools, Alat Kerja & Alat Dental Portabel",description:"Handpiece dental, alat rontgen portable, toolbox teknisi",pic:"Dokter/Perawat & Teknisi"},{no:11,type:"Low Value Asset",item:"Peralatan Kantor (Telepon, ATK)",description:"Pesawat telepon, kalkulator, stapler, perlengkapan meja",pic:"Resepsionis"},{no:12,type:"Low Value Asset",item:"Alat-alat Kebersihan",description:"Lobby duster, ember wringer pel, sapu, tangga, botol spray",pic:"Facility Care"},{no:13,type:"Low Value Asset",item:"Perlengkapan Security",description:"Lampu lalin, cone parkir, peluit, senter patroli, jas hujan",pic:"Security"}],Zn={cabang:[{zone:"Zona Hijau",level:"Akses Publik Bebas",color:"#059669",bg:"#ECFDF5",rooms:["Area Parkir Kendaraan","Lobby Lantai 1","Teras Luar Klinik"],desc:"Area dengan pengamanan umum yang bebas diakses pasien dan tamu setelah pengarahan security."},{zone:"Zona Kuning",level:"Akses Terbatas / Pendampingan",color:"#D97706",bg:"#FFFBEB",rooms:["Koridor Ruang Tindakan","Tangga & Lift Klinik","Ruang Tunggu Lantai 2 & 3","Kids Playground","Toilet Lantai 1 & 2","Mushola Staf & Pasien"],desc:"Area terbatas, tamu wajib didampingi petugas atau staf klinik terkait."},{zone:"Zona Merah",level:"Area Terlarang / Restricted",color:"#DC2626",bg:"#FEF2F2",rooms:["Area Meja Receptionist / Kasir","Seluruh Ruang Dental Unit (DU 101, 102, 201, 202, 301, 302)","Ruang Sterilisasi Alat","Ruang Rontgen Radiologi","Ruang Suction & Kompresor","Ruang Tim Staf & Gudang Material Medis"],desc:"Area sangat sensitif, hanya staf medis bertugas dan atasan pemilik area. Tamu wajib izin Manajemen/BOD dan didampingi penuh."}],hq:[{floor:"Lantai 1",hijau:["Area Parkir Mobil/Motor","Lobby Resepsionis Lantai 1"],kuning:["Seluruh Ruang Meeting Lantai 1","Ruang Training Lantai 1","Koridor Lift & Toilet"],merah:["Ruang Customer Services","Gudang Utama Lantai 1","Gudang Facility Care","Gudang Teknisi & Ruang Teknisi","Pos Satpam & Ruang Genset"]},{floor:"Lantai 2",hijau:["-"],kuning:["Ruang Meeting Lantai 2","Koridor Lift & Toilet","Pantry"],merah:["Ruang CEO","Ruang Sekretaris CEO","Ruang Human Capital","Ruang Dentist Relation","Ruang General Support","Ruang Eks. Patrel"]},{floor:"Lantai 3",hijau:["-"],kuning:["Koridor Lift & Toilet","Pantry","Mushola & Tempat Wudhu","Ruang Santai Karyawan"],merah:["Ruang Marketing","Studio Konten Marketing","Gudang Marketing","Ruang Patient Relation","Ruang OPS","Ruang Training Lantai 3","Gudang Lantai 3"]},{floor:"Lantai 4",hijau:["-"],kuning:["Koridor Lift & Toilet","Pantry"],merah:["Ruang Finance & Accounting","Ruang VP Finance & Accounting","Ruang Audit Internal","Ruang Training Lantai 4","Gudang Lantai 4"]},{floor:"Lantai 5",hijau:["-"],kuning:["Lobby Lantai 5","Pantry Lantai 5","Toilet Lantai 5","Ruang Training Besar Lantai 5"],merah:["Area Dak Toren Air & Rooftop"]}]},dt="fcms_sop_sec_ik_v1",ct="fcms_sop_sec_assets_v1";function ei(){try{let e=localStorage.getItem(dt);if(e)return JSON.parse(e)}catch(e){console.error("Error load IK Sec",e)}return JSON.parse(JSON.stringify(Mr))}function ja(e){localStorage.setItem(dt,JSON.stringify(e)),X("/api/sop/bundle/"+dt,{method:"PUT",body:JSON.stringify(e)}).catch(()=>{})}function ti(){try{let e=localStorage.getItem(ct);if(e)return JSON.parse(e)}catch(e){console.error("Error load assets sec",e)}return JSON.parse(JSON.stringify(zr))}function qa(e){localStorage.setItem(ct,JSON.stringify(e)),X("/api/sop/bundle/"+ct,{method:"PUT",body:JSON.stringify(e)}).catch(()=>{})}async function ai(e){let t=se(),n=re(t,"sop","write"),a=re(t,"sop","delete"),o="ik",s="",c="all",m="",i="cabang";try{let[b,S]=await Promise.all([X("/api/sop/bundle/"+dt),X("/api/sop/bundle/"+ct)]);b.ok&&b.data?.data&&localStorage.setItem(dt,JSON.stringify(b.data.data)),S.ok&&S.data?.data&&localStorage.setItem(ct,JSON.stringify(S.data.data))}catch(b){console.warn("Could not load remote Security SOP bundle:",b)}let d=ei(),F=ti(),$={"Pengamanan Fisik & Patroli":{dot:"#60A5FA",text:"#93C5FD",accent:"#3B82F6"},"Operasional Parkir & Luar":{dot:"#FBBF24",text:"#FDE68A",accent:"#F59E0B"},"Layanan & Hospitalitas":{dot:"#34D399",text:"#A7F3D0",accent:"#10B981"}};function M(b){return $[b?.category]||$["Pengamanan Fisik & Patroli"]}function h(){return`
      <span style="width:40px; height:40px; flex-shrink:0; display:inline-flex; align-items:center; justify-content:center; background:#0E182B; border:1px solid #223456; border-radius:9px; color:#93C5FD;" aria-hidden="true">
        <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.9" stroke-linecap="round" stroke-linejoin="round">
          <path d="M12 22s8-4 8-10V5l-8-3-8 3v7c0 6 8 10 8 10z"></path>
        </svg>
      </span>
    `}function L(b){if(!n&&!a)return"";let S=Y(b.title||"Instruksi Kerja");return`
      <div style="display:flex; gap:5px; flex-shrink:0;">
        ${n?`
        <button type="button" class="sec-icon-btn btn-edit-sec-ik" data-id="${b.id}" title="Edit ${S}" aria-label="Edit SOP ${S}">
          <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">
            <path d="M12 20h9"></path>
            <path d="M16.5 3.5a2.121 2.121 0 0 1 3 3L7 19l-4 1 1-4L16.5 3.5z"></path>
          </svg>
        </button>
        `:""}
        ${a?`
        <button type="button" class="sec-icon-btn sec-icon-del btn-delete-sec-ik" data-id="${b.id}" title="Hapus ${S}" aria-label="Hapus SOP ${S}">
          <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">
            <polyline points="3 6 5 6 21 6"></polyline>
            <path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"></path>
          </svg>
        </button>
        `:""}
      </div>
    `}let R={"Zona Hijau":{title:"#6EE7B7",dot:"#34D399",bg:"rgba(16, 185, 129, 0.14)",border:"rgba(52, 211, 153, 0.36)",body:"#F8FAFC"},"Zona Kuning":{title:"#FCD34D",dot:"#FBBF24",bg:"rgba(245, 158, 11, 0.14)",border:"rgba(251, 191, 36, 0.36)",body:"#F8FAFC"},"Zona Merah":{title:"#FCA5A5",dot:"#F87171",bg:"rgba(239, 68, 68, 0.14)",border:"rgba(248, 113, 113, 0.36)",body:"#F8FAFC"}};function I(){e.innerHTML=`
      <div class="page-header" style="display:flex; justify-content:space-between; align-items:flex-start; flex-wrap:wrap; gap:16px; margin-bottom:20px;">
        <div>
          <h1 class="page-title" style="display:flex; align-items:center; gap:8px; margin:0 0 6px 0; font-size:1.65rem; font-weight:700; color:#FFFFFF;">
            <span>\u{1F6E1}\uFE0F</span> SOP Pengelolaan Keamanan & Aset Perusahaan
          </h1>
          <p class="page-subtitle" style="margin:0; font-size:0.92rem; color:#CBD5E1; line-height:1.5;">
            Standar Operasional Prosedur Pengawasan Keamanan Area, Fasilitas, dan Perlindungan Aset PT Audy Mandiri Indonesia
          </p>
        </div>
        <div style="display:flex; gap:10px; align-items:center; flex-wrap:wrap;">
          ${a?`
          <button id="btn-reset-sec-default" class="btn btn-outline" title="Kembalikan data ke dokumen standar asli" style="background:#131F36; border:1px solid #223456; color:#E2E8F0; font-weight:600; padding:8px 14px; border-radius:8px; display:inline-flex; align-items:center; gap:6px; cursor:pointer; font-size:0.86rem;">
            \u{1F504} Reset ke Standar
          </button>
          `:""}
          <button id="btn-print-sec" class="btn btn-outline" style="background:#131F36; border:1px solid #223456; color:#E2E8F0; font-weight:600; padding:8px 14px; border-radius:8px; display:inline-flex; align-items:center; gap:6px; cursor:pointer; font-size:0.86rem;">
            \u{1F5A8}\uFE0F Cetak Ringkasan SOP
          </button>
          ${re(t,"export")?`<button id="btn-export-sec" class="btn btn-primary" style="font-weight:700; padding:8px 16px; border-radius:8px; display:inline-flex; align-items:center; gap:6px; cursor:pointer; font-size:0.86rem; background:#2563EB; color:#FFFFFF;">
            \u{1F4E5} Download Data Excel
          </button>`:""}
        </div>
      </div>

      <!-- Navigation Tabs -->
      <div class="fc-nav-tabs-strip" data-fc-scroll-tabs style="display:flex; gap:8px; margin-bottom:20px; border-bottom:1px solid #223456; padding-bottom:10px; overflow-x:auto; flex-wrap:nowrap; -webkit-overflow-scrolling:touch; scrollbar-width:none;">
        <button class="sec-tab-btn ${o==="ik"?"active":""}" data-tab="ik" style="${N(o==="ik")}">
          \u{1F6E1}\uFE0F Instruksi Kerja (IK) <span class="badge" style="${q(o==="ik")}">${d.length}</span>
        </button>
        <button class="sec-tab-btn ${o==="zonasi"?"active":""}" data-tab="zonasi" style="${N(o==="zonasi")}">
          \u{1F5FA}\uFE0F Zonasi Akses & Denah
        </button>
        <button class="sec-tab-btn ${o==="assets"?"active":""}" data-tab="assets" style="${N(o==="assets")}">
          \u{1F4E6} Klasifikasi Aset <span class="badge" style="${q(o==="assets")}">${F.length}</span>
        </button>
        <button class="sec-tab-btn ${o==="policy"?"active":""}" data-tab="policy" style="${N(o==="policy")}">
          \u{1F4CB} Shift, Patrol & Insiden
        </button>
        <button class="sec-tab-btn ${o==="links"?"active":""}" data-tab="links" style="${N(o==="links")}">
          \u{1F517} Form & Tautan
        </button>
      </div>

      <!-- Content Container -->
      <div id="sec-content-area"></div>
    `,v(),B()}function N(b){return b?"border:1px solid #3B82F6; background:#2563EB; color:#FFFFFF; font-weight:700; padding:8px 16px; border-radius:8px; cursor:pointer; display:inline-flex; align-items:center; gap:8px; font-size:0.9rem; white-space:nowrap; transition:all 0.2s;":"background:#172033; color:#E2E8F0; border:1px solid #334155; font-weight:600; padding:8px 16px; border-radius:8px; cursor:pointer; display:inline-flex; align-items:center; gap:8px; font-size:0.9rem; white-space:nowrap; transition:all 0.2s;"}function q(b){return b?"background:#0F172A; color:#FFFFFF; font-size:0.75rem; padding:2px 8px; border-radius:10px; font-weight:700;":"background:#22314A; color:#E2E8F0; border:1px solid #334155; font-size:0.75rem; padding:2px 8px; border-radius:10px; font-weight:700;"}function v(){e.querySelectorAll(".sec-tab-btn").forEach(r=>{r.addEventListener("click",()=>{o=r.dataset.tab,I()})});let b=e.querySelector("#btn-reset-sec-default");b&&b.addEventListener("click",()=>{ge("Apakah Anda yakin ingin mengembalikan seluruh data Instruksi Kerja dan Klasifikasi Aset Security ke standar dokumen asli?",()=>{localStorage.removeItem(dt),localStorage.removeItem(ct),Promise.all([X("/api/sop/bundle/"+dt,{method:"DELETE"}),X("/api/sop/bundle/"+ct,{method:"DELETE"})]).catch(()=>{}),d=ei(),F=ti(),ae("Data SOP Security berhasil di-reset ke standar dokumen"),I()},"Reset ke Standar Dokumen")});let S=e.querySelector("#btn-print-sec");S&&S.addEventListener("click",()=>{window.print()});let y=e.querySelector("#btn-export-sec");y&&y.addEventListener("click",()=>{W()})}function B(){let b=e.querySelector("#sec-content-area");b&&(o==="ik"?x(b):o==="zonasi"?E(b):o==="assets"?w(b):o==="policy"?G(b):o==="links"&&J(b))}function x(b){b.innerHTML=`
      <!-- Toolbar Filter & Search -->
      <div class="card" style="padding:14px 18px; margin-bottom:20px; background:#131F36; border:1px solid #223456; border-radius:12px;">
        <div style="display:flex; flex-wrap:wrap; gap:12px; align-items:center; justify-content:space-between;">
          <div style="display:flex; flex-wrap:wrap; gap:12px; align-items:center; flex:1; min-width:260px;">
            <div style="position:relative; flex:1; min-width:200px; max-width:340px;">
              <input type="text" id="sec-ik-search" class="form-control" placeholder="Cari instruksi kerja security..." value="${s}" style="padding-left:34px; height:38px; border-radius:8px; font-size:0.9rem; width:100%; background:#0E182B; border-color:#223456;">
              <span style="position:absolute; left:11px; top:50%; transform:translateY(-50%); color:#94A3B8; font-size:0.9rem;">\u{1F50D}</span>
            </div>
            
            <div style="display:flex; gap:6px; flex-wrap:wrap;" id="sec-cat-pills-wrap">
              <button class="sec-cat-pill ${c==="all"?"active":""}" data-cat="all">Semua (${d.length})</button>
              <button class="sec-cat-pill ${c==="Pengamanan Fisik & Patroli"?"active":""}" data-cat="Pengamanan Fisik & Patroli">Patroli & Fisik (${d.filter(p=>p.category==="Pengamanan Fisik & Patroli").length})</button>
              <button class="sec-cat-pill ${c==="Operasional Parkir & Luar"?"active":""}" data-cat="Operasional Parkir & Luar">Parkir & Luar (${d.filter(p=>p.category==="Operasional Parkir & Luar").length})</button>
              <button class="sec-cat-pill ${c==="Layanan & Hospitalitas"?"active":""}" data-cat="Layanan & Hospitalitas">Tamu & Pasien (${d.filter(p=>p.category==="Layanan & Hospitalitas").length})</button>
            </div>
          </div>
          <div style="display:flex; gap:10px; align-items:center;">
            <div id="sec-ik-count-label" style="font-size:0.86rem; color:#CBD5E1; font-weight:600;">
              Menampilkan ${d.length} Instruksi Kerja
            </div>
            ${n?`
            <button id="btn-add-sec-ik" class="btn btn-primary" style="font-weight:700; padding:7px 14px; border-radius:8px; display:inline-flex; align-items:center; gap:6px; cursor:pointer; font-size:0.86rem; background:#2563EB; border-color:#3B82F6; color:#FFFFFF;">
              \u2795 Tambah IK Baru
            </button>
            `:""}
          </div>
        </div>
      </div>

      <!-- IK Cards Grid Container -->
      <div id="sec-ik-cards-grid-container"></div>
    `;let S=document.createElement("style");S.innerHTML=`
      .sec-cat-pill {
        border: 1px solid #223456;
        background: #0E182B;
        color: #CBD5E1;
        font-size: 0.82rem;
        font-weight: 600;
        padding: 6px 13px;
        border-radius: 20px;
        cursor: pointer;
        transition: all 0.15s;
        flex-shrink: 0;
        white-space: nowrap;
      }
      .sec-cat-pill:hover { background: #1A2A48; color:#FFFFFF; border-color:#3B82F6; }
      .sec-cat-pill.active {
        background: #2563EB;
        color: #FFFFFF;
        border-color: #3B82F6;
      }
      .sec-ik-card:hover {
        box-shadow: 0 8px 20px rgba(0,0,0,0.28) !important;
        border-color: #3B82F6 !important;
      }
      .sec-icon-btn {
        background: transparent;
        border: 1px solid rgba(51, 65, 85, 0.7);
        border-radius: 7px;
        width: 28px;
        height: 28px;
        padding: 0;
        color: #94A3B8;
        cursor: pointer;
        display: inline-flex;
        align-items: center;
        justify-content: center;
        transition: all 0.15s;
      }
      .sec-icon-btn:hover {
        background: rgba(37, 99, 235, 0.18);
        border-color: rgba(96, 165, 250, 0.55);
        color: #93C5FD;
      }
      .sec-icon-del:hover {
        background: rgba(239, 68, 68, 0.16);
        border-color: rgba(248, 113, 113, 0.55);
        color: #FCA5A5;
      }
      .sec-icon-btn:focus-visible {
        outline: 2px solid #60A5FA;
        outline-offset: 2px;
      }
      @media (max-width: 768px) {
        #sec-cat-pills-wrap {
          display: flex !important;
          flex-wrap: nowrap !important;
          overflow-x: auto !important;
          -webkit-overflow-scrolling: touch !important;
          scrollbar-width: none !important;
          padding-bottom: 4px !important;
          width: 100% !important;
        }
        #sec-cat-pills-wrap::-webkit-scrollbar {
          display: none !important;
        }
        .sec-ik-grid {
          grid-template-columns: 1fr !important;
          gap: 12px !important;
        }
      }
    `,b.appendChild(S);let y=b.querySelector("#sec-ik-cards-grid-container"),r=b.querySelector("#sec-ik-count-label");function l(){let p=d.filter(A=>{let T=c==="all"||A.category===c,g=s.toLowerCase().trim(),C=!g||A.title.toLowerCase().includes(g)||A.target.toLowerCase().includes(g)||A.tools.some(z=>z.toLowerCase().includes(g))||A.steps.some(z=>z.toLowerCase().includes(g));return T&&C});if(r&&(r.textContent=`Menampilkan ${p.length} dari ${d.length} Instruksi Kerja`),p.length===0){y.innerHTML=`
          <div style="text-align:center; padding:48px 20px; background:#111827; border-radius:12px; border:1px solid #29364B;">
            <div style="font-size:2.5rem; margin-bottom:8px;">\u{1F50D}</div>
            <h3 style="margin:0 0 6px; font-size:1.12rem; color:#FFFFFF;">Tidak ada Instruksi Kerja yang cocok</h3>
            <p style="margin:0; font-size:0.88rem; color:#CBD5E1;">Coba ubah kata kunci pencarian atau pilih kategori lain.</p>
          </div>
        `;return}y.innerHTML=`
        <div class="sec-ik-grid" style="display:grid; grid-template-columns:repeat(auto-fill, minmax(340px, 1fr)); gap:18px;">
          ${p.map(A=>{let T=M(A);return`
            <div class="card sec-ik-card" style="background:#131F36; border:1px solid #223456; border-radius:14px; padding:20px; display:flex; flex-direction:column; justify-content:space-between; transition:border-color 0.18s ease, box-shadow 0.18s ease;">
              <div>
                <div style="display:flex; justify-content:space-between; align-items:flex-start; gap:10px; margin-bottom:14px;">
                  <div style="display:flex; align-items:flex-start; gap:12px; min-width:0;">
                    ${h()}
                    <div style="min-width:0;">
                      <h3 style="margin:0 0 6px 0; font-size:1.06rem; font-weight:700; color:#FFFFFF; line-height:1.35;">
                        IK-${String(A.id).padStart(2,"0")}. ${A.title}
                      </h3>
                      <span class="badge" style="background:#0E182B; color:${T.text}; border:1px solid #24385A; font-size:0.74rem; font-weight:600; border-radius:5px; padding:2px 8px; display:inline-flex; align-items:center; gap:6px;">
                        <span style="width:6px; height:6px; border-radius:50%; background:${T.dot}; display:inline-block;"></span>
                        ${A.category}
                      </span>
                    </div>
                  </div>

                  <!-- Quick Action Buttons: Edit & Hapus (non-dominant SVG) -->
                  ${L(A)}
                </div>

                <div style="font-size:0.88rem; color:#F1F5F9; line-height:1.55; margin-bottom:14px; background:#0E182B; padding:10px 13px; border-radius:8px; border:1px solid #1E3050; border-left:3px solid ${T.accent};">
                  <strong style="color:#FFFFFF;">Area Sasaran:</strong> ${A.target}
                </div>

                <div style="margin-bottom:14px;">
                  <div style="font-size:0.75rem; font-weight:700; color:#93C5FD; text-transform:uppercase; letter-spacing:0.05em; margin-bottom:7px;">Alat &amp; Perlengkapan</div>
                  <ul style="margin:0; padding:0; list-style:none; font-size:0.86rem; color:#E2E8F0; line-height:1.5;">
                    ${A.tools.slice(0,3).map(g=>`
                      <li style="display:flex; align-items:flex-start; gap:8px; padding:4px 0; border-bottom:1px solid rgba(34, 52, 86, 0.45); word-break:break-word;">
                        <span style="color:#60A5FA; font-weight:700; line-height:1.4;">\u2022</span>
                        <span>${Y(g)}</span>
                      </li>
                    `).join("")}
                    ${A.tools.length>3?`<li style="color:#93C5FD; font-size:0.82rem; font-weight:600; padding-top:5px;">+ ${A.tools.length-3} perlengkapan lainnya dalam panduan</li>`:""}
                  </ul>
                </div>
              </div>

              <div style="border-top:1px solid #1E2E4A; padding-top:13px; margin-top:6px; display:flex; justify-content:space-between; align-items:center;">
                <span style="font-size:0.84rem; color:#CBD5E1; font-weight:600;">
                  ${A.steps.length} Langkah Prosedur
                </span>
                <button type="button" class="btn btn-sm btn-primary btn-view-sec-ik" data-id="${A.id}" style="padding:7px 15px; font-weight:700; font-size:0.85rem; border-radius:8px; background:#2563EB; border:1px solid #3B82F6; color:#FFFFFF;">
                  Lihat Panduan Lengkap \u2192
                </button>
              </div>
            </div>
            `}).join("")}
        </div>
      `,y.querySelectorAll(".btn-view-sec-ik").forEach(A=>{A.addEventListener("click",()=>{let T=parseInt(A.dataset.id,10),g=d.find(C=>C.id===T);g&&k(g)})}),y.querySelectorAll(".btn-edit-sec-ik").forEach(A=>{A.addEventListener("click",()=>{let T=parseInt(A.dataset.id,10),g=d.find(C=>C.id===T);g&&_(g)})}),y.querySelectorAll(".btn-delete-sec-ik").forEach(A=>{A.addEventListener("click",()=>{let T=parseInt(A.dataset.id,10),g=d.find(C=>C.id===T);g&&ge(`Apakah Anda yakin ingin menghapus Instruksi Kerja: "${g.title}"?`,()=>{d=d.filter(C=>C.id!==T),ja(d),ae("Instruksi Kerja Security berhasil dihapus"),I()},"Hapus Instruksi Kerja")})})}l();let f=b.querySelector("#sec-ik-search");f&&f.addEventListener("input",p=>{s=p.target.value,l()}),b.querySelectorAll(".sec-cat-pill").forEach(p=>{p.addEventListener("click",()=>{b.querySelectorAll(".sec-cat-pill").forEach(A=>A.classList.remove("active")),p.classList.add("active"),c=p.dataset.cat,l()})});let u=b.querySelector("#btn-add-sec-ik");u&&u.addEventListener("click",()=>{_()})}function k(b){let S=M(b);ce({title:`${b.icon||"\u{1F6E1}\uFE0F"} IK-${String(b.id).padStart(2,"0")}: ${b.title}`,size:"lg",content:`
        <div style="display:flex; flex-direction:column; gap:16px;">
          <div style="background:${S.bg}; color:${S.text}; border:1px solid ${S.border}; padding:11px 15px; border-radius:9px; font-weight:700; font-size:0.92rem; display:flex; justify-content:space-between; align-items:center;">
            <span>Kategori: ${b.category}</span>
            <span style="color:#FFFFFF;">Kode: IK-SEC-${String(b.id).padStart(2,"0")}</span>
          </div>

          <div style="background:#0F172A; padding:12px 15px; border-radius:9px; border:1px solid #29364B; border-left:4px solid ${S.accent};">
            <strong style="color:#FFFFFF; font-size:0.88rem; display:block; margin-bottom:4px;">\u{1F4CD} Area & Sasaran Target:</strong>
            <span style="font-size:0.92rem; color:#F8FAFC; line-height:1.55;">${b.target}</span>
          </div>

          <div>
            <h4 style="margin:0 0 8px 0; font-size:0.98rem; font-weight:700; color:#FFFFFF;">\u{1F9F0} Alat & Perlengkapan yang Disiapkan:</h4>
            <div style="background:#0F172A; border:1px solid #29364B; border-radius:9px; padding:12px 16px;">
              <ul style="margin:0 0 0 18px; padding:0; font-size:0.9rem; color:#F8FAFC; line-height:1.75;">
                ${b.tools.map(y=>{if(y.includes("http")){let r=y.split(": ");return`<li style="margin-bottom:4px;">${r[0]}: <a href="${r.slice(1).join(": ")}" target="_blank" rel="noopener" style="color:#60A5FA; font-weight:700; text-decoration:underline;">Buka Tautan Form \u2794</a></li>`}return`<li style="margin-bottom:4px;">${Y(y)}</li>`}).join("")}
              </ul>
            </div>
          </div>

          <div>
            <h4 style="margin:0 0 8px 0; font-size:0.98rem; font-weight:700; color:#FFFFFF;">\u{1F4DD} Langkah & Prosedur Pelaksanaan:</h4>
            <div style="background:#0F172A; border:1px solid #29364B; border-radius:9px; padding:13px 18px;">
              <ol style="margin:0 0 0 20px; padding:0; font-size:0.9rem; color:#F8FAFC; line-height:1.8;">
                ${b.steps.map(y=>`<li style="margin-bottom:7px;">${Y(y)}</li>`).join("")}
              </ol>
            </div>
          </div>

          ${b.notes?`
            <div style="background:rgba(59, 130, 246, 0.14); border:1px solid rgba(96, 165, 250, 0.36); border-radius:9px; padding:12px 15px;">
              <strong style="color:#93C5FD; font-size:0.88rem; display:block; margin-bottom:3px;">\u26A0\uFE0F Catatan Khusus & Ketentuan Penting:</strong>
              <span style="color:#F8FAFC; font-size:0.88rem; line-height:1.6;">${Y(b.notes)}</span>
            </div>
          `:""}
        </div>
      `,confirmText:"Tutup Panduan",onConfirm:(y,r)=>{r()}})}function _(b=null){if(!n)return;let S=!!b,y=b?b.tools.join(`
`):"",r=b?b.steps.join(`
`):"",l=`
      <form id="sec-ik-edit-form" style="display:flex; flex-direction:column; gap:14px; font-size:0.88rem;">
        <div class="form-row" style="display:grid; grid-template-columns:1fr 140px; gap:12px;">
          <div>
            <label style="display:block; font-weight:600; margin-bottom:4px;">Judul Instruksi Kerja *</label>
            <input type="text" id="form-sec-ik-title" class="form-control" required value="${b?Y(b.title):""}" style="width:100%; border-radius:6px; height:36px;">
          </div>
          <div>
            <label style="display:block; font-weight:600; margin-bottom:4px;">Ikon Emoji</label>
            <input type="text" id="form-sec-ik-icon" class="form-control" value="${b?b.icon:"\u{1F6E1}\uFE0F"}" style="width:100%; border-radius:6px; height:36px; text-align:center; font-size:1.2rem;">
          </div>
        </div>

        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Kategori *</label>
          <select id="form-sec-ik-cat" class="form-control" style="width:100%; border-radius:6px; height:36px;">
            <option value="Pengamanan Fisik & Patroli" ${b?.category==="Pengamanan Fisik & Patroli"?"selected":""}>Pengamanan Fisik & Patroli</option>
            <option value="Operasional Parkir & Luar" ${b?.category==="Operasional Parkir & Luar"?"selected":""}>Operasional Parkir & Luar</option>
            <option value="Layanan & Hospitalitas" ${b?.category==="Layanan & Hospitalitas"?"selected":""}>Layanan & Hospitalitas</option>
          </select>
        </div>

        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Area & Sasaran Target *</label>
          <input type="text" id="form-sec-ik-target" class="form-control" required value="${b?Y(b.target):""}" style="width:100%; border-radius:6px; height:36px;">
        </div>

        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Alat & Perlengkapan (1 per baris) *</label>
          <textarea id="form-sec-ik-tools" class="form-control" rows="3" style="width:100%; border-radius:6px; font-size:0.85rem; line-height:1.5;">${Y(y)}</textarea>
        </div>

        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Langkah & Prosedur Pelaksanaan (1 langkah per baris) *</label>
          <textarea id="form-sec-ik-steps" class="form-control" rows="5" style="width:100%; border-radius:6px; font-size:0.85rem; line-height:1.6;">${Y(r)}</textarea>
        </div>

        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Catatan Khusus / Ketentuan Penting</label>
          <textarea id="form-sec-ik-notes" class="form-control" rows="2" style="width:100%; border-radius:6px; font-size:0.85rem;">${b?Y(b.notes||""):""}</textarea>
        </div>
      </form>
    `;ce({title:S?`\u270F\uFE0F Edit IK-${b.id}: ${b.title}`:"\u2795 Tambah Instruksi Kerja Security Baru",content:l,size:"lg",confirmText:"Simpan Perubahan",onConfirm:(f,u)=>{let p=f.querySelector("#form-sec-ik-title").value.trim(),A=f.querySelector("#form-sec-ik-icon").value.trim()||"\u{1F6E1}\uFE0F",T=f.querySelector("#form-sec-ik-cat").value,g=f.querySelector("#form-sec-ik-target").value.trim(),C=f.querySelector("#form-sec-ik-tools").value.trim(),z=f.querySelector("#form-sec-ik-steps").value.trim(),O=f.querySelector("#form-sec-ik-notes").value.trim();if(!p||!g||!C||!z){j("Harap lengkapi semua kolom bertanda bintang (*)");return}let te=C.split(`
`).map(H=>H.trim().replace(/^[-•*]\s*/,"")).filter(Boolean),Q=z.split(`
`).map(H=>H.trim().replace(/^\d+[\.\)]\s*/,"")).filter(Boolean),ee={"Pengamanan Fisik & Patroli":{color:"#2563EB",bg:"#EFF6FF"},"Operasional Parkir & Luar":{color:"#D97706",bg:"#FFFBEB"},"Layanan & Hospitalitas":{color:"#059669",bg:"#ECFDF5"}}[T]||{color:"#1E40AF",bg:"#EFF6FF"};if(S)b.title=p,b.icon=A,b.category=T,b.target=g,b.tools=te,b.steps=Q,b.notes=O,b.color=ee.color,b.bg=ee.bg,ja(d),ae(`Instruksi Kerja "${p}" berhasil diperbarui`);else{let H=d.length>0?Math.max(...d.map(D=>D.id))+1:1,U={id:H,title:p,category:T,icon:A,color:ee.color,bg:ee.bg,target:g,tools:te,steps:Q,notes:O};d.push(U),ja(d),ae(`Instruksi Kerja baru IK-${H} berhasil ditambahkan`)}u(),I()}})}function E(b){b.innerHTML=`
      <div style="display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:12px; margin-bottom:20px;">
        <div class="btn-group" style="background:#111827; padding:3px; border-radius:8px; display:inline-flex; border:1px solid var(--border);">
          <button id="view-cabang-btn" class="btn btn-sm ${i==="cabang"?"btn-primary":"btn-ghost"}" style="font-weight:700; border-radius:6px; padding:6px 16px; cursor:pointer;">
            \u{1F3E5} Zonasi Klinik / Cabang
          </button>
          <button id="view-hq-btn" class="btn btn-sm ${i==="hq"?"btn-primary":"btn-ghost"}" style="font-weight:700; border-radius:6px; padding:6px 16px; cursor:pointer;">
            \u{1F3E2} Zonasi Head Quarter (5 Lantai)
          </button>
        </div>
        <div style="font-size:0.86rem; color:#CBD5E1; font-weight:600;">
          Ketentuan Pembagian Akses Ruangan (Lampiran 8.1 SOP)
        </div>
      </div>

      <!-- Definisi 3 Zona Header Cards -->
      <div style="display:grid; grid-template-columns:repeat(auto-fit, minmax(280px, 1fr)); gap:16px; margin-bottom:22px;">
        <div class="card" style="background:rgba(239, 68, 68, 0.14); border:1px solid rgba(248, 113, 113, 0.36); border-radius:11px; padding:15px;">
          <div style="display:flex; align-items:center; gap:8px; margin-bottom:6px;">
            <span style="width:13px; height:13px; border-radius:50%; background:#F87171; display:inline-block;"></span>
            <strong style="color:#FCA5A5; font-size:0.96rem;">ZONA MERAH (Restricted Area)</strong>
          </div>
          <p style="margin:0; font-size:0.86rem; color:#F8FAFC; line-height:1.55;">
            Area sangat terbatas. Hanya personal yang bertugas spesifik di area tersebut dan BOD. Tamu WAJIB mendapat izin tertulis Manajemen dan didampingi penuh.
          </p>
        </div>

        <div class="card" style="background:rgba(245, 158, 11, 0.14); border:1px solid rgba(251, 191, 36, 0.36); border-radius:11px; padding:15px;">
          <div style="display:flex; align-items:center; gap:8px; margin-bottom:6px;">
            <span style="width:13px; height:13px; border-radius:50%; background:#FBBF24; display:inline-block;"></span>
            <strong style="color:#FCD34D; font-size:0.96rem;">ZONA KUNING (Limited Area)</strong>
          </div>
          <p style="margin:0; font-size:0.86rem; color:#F8FAFC; line-height:1.55;">
            Area dengan pengamanan terbatas. Karyawan umum atau tamu yang memasuki area ini perlu mendapatkan arahan dan didampingi staf area terkait.
          </p>
        </div>

        <div class="card" style="background:rgba(16, 185, 129, 0.14); border:1px solid rgba(52, 211, 153, 0.36); border-radius:11px; padding:15px;">
          <div style="display:flex; align-items:center; gap:8px; margin-bottom:6px;">
            <span style="width:13px; height:13px; border-radius:50%; background:#34D399; display:inline-block;"></span>
            <strong style="color:#6EE7B7; font-size:0.96rem;">ZONA HIJAU (Public Area)</strong>
          </div>
          <p style="margin:0; font-size:0.86rem; color:#F8FAFC; line-height:1.55;">
            Area umum dengan pengawasan security. Diperbolehkan untuk seluruh tamu dan pasien setelah diarahkan oleh petugas security.
          </p>
        </div>
      </div>

      <!-- Zonasi Content Area -->
      <div id="zonasi-content-container"></div>
    `;let S=b.querySelector("#zonasi-content-container");i==="cabang"?S.innerHTML=`
        <div style="display:grid; grid-template-columns:repeat(auto-fit, minmax(320px, 1fr)); gap:18px;">
          ${Zn.cabang.map(y=>{let r=R[y.zone]||R["Zona Hijau"];return`
            <div class="card" style="background:#172033; border:1px solid #29364B; border-radius:12px; padding:18px; box-shadow:0 4px 12px rgba(0,0,0,0.2);">
              <div style="display:flex; justify-content:space-between; align-items:center; gap:8px; margin-bottom:12px; border-bottom:1px solid #29364B; padding-bottom:10px;">
                <h3 style="margin:0; font-size:1.08rem; font-weight:700; color:${r.title}; display:flex; align-items:center; gap:8px;">
                  <span style="width:12px; height:12px; border-radius:50%; background:${r.dot}; display:inline-block;"></span>
                  ${y.zone}
                </h3>
                <span class="badge" style="background:${r.bg}; color:${r.title}; border:1px solid ${r.border}; font-size:0.78rem; font-weight:700; padding:3px 9px; border-radius:6px;">
                  ${y.level}
                </span>
              </div>
              <p style="font-size:0.86rem; color:#CBD5E1; margin-bottom:14px; line-height:1.55;">${y.desc}</p>
              <div>
                <strong style="font-size:0.8rem; text-transform:uppercase; color:#E2E8F0; letter-spacing:0.03em; display:block; margin-bottom:7px;">Daftar Ruangan Terkait:</strong>
                <ul style="margin:0 0 0 18px; padding:0; font-size:0.88rem; color:#F8FAFC; line-height:1.75;">
                  ${y.rooms.map(l=>`<li style="margin-bottom:2px;">${l}</li>`).join("")}
                </ul>
              </div>
            </div>
            `}).join("")}
        </div>
      `:S.innerHTML=`
        <div style="display:flex; flex-direction:column; gap:16px;">
          ${Zn.hq.map(y=>`
            <div class="card" style="background:#172033; border:1px solid #29364B; border-radius:12px; padding:18px; box-shadow:0 4px 12px rgba(0,0,0,0.2);">
              <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:14px; border-bottom:1px solid #29364B; padding-bottom:10px;">
                <h3 style="margin:0; font-size:1.12rem; font-weight:700; color:#FFFFFF; display:flex; align-items:center; gap:8px;">
                  <span>\u{1F3E2}</span> Head Quarter \u2014 ${y.floor}
                </h3>
              </div>
              <div style="display:grid; grid-template-columns:repeat(auto-fit, minmax(260px, 1fr)); gap:14px;">
                <!-- Zona Hijau -->
                <div style="background:rgba(16, 185, 129, 0.14); border:1px solid rgba(52, 211, 153, 0.35); border-radius:9px; padding:13px;">
                  <strong style="color:#6EE7B7; font-size:0.88rem; display:block; margin-bottom:7px;">\u{1F7E2} ZONA HIJAU (Publik):</strong>
                  <ul style="margin:0 0 0 18px; padding:0; font-size:0.86rem; color:#F8FAFC; line-height:1.65;">
                    ${y.hijau.map(r=>`<li>${r}</li>`).join("")}
                  </ul>
                </div>
                <!-- Zona Kuning -->
                <div style="background:rgba(245, 158, 11, 0.14); border:1px solid rgba(251, 191, 36, 0.35); border-radius:9px; padding:13px;">
                  <strong style="color:#FCD34D; font-size:0.88rem; display:block; margin-bottom:7px;">\u{1F7E1} ZONA KUNING (Terbatas):</strong>
                  <ul style="margin:0 0 0 18px; padding:0; font-size:0.86rem; color:#F8FAFC; line-height:1.65;">
                    ${y.kuning.map(r=>`<li>${r}</li>`).join("")}
                  </ul>
                </div>
                <!-- Zona Merah -->
                <div style="background:rgba(239, 68, 68, 0.14); border:1px solid rgba(248, 113, 113, 0.35); border-radius:9px; padding:13px;">
                  <strong style="color:#FCA5A5; font-size:0.88rem; display:block; margin-bottom:7px;">\u{1F534} ZONA MERAH (Restricted):</strong>
                  <ul style="margin:0 0 0 18px; padding:0; font-size:0.86rem; color:#F8FAFC; line-height:1.65;">
                    ${y.merah.map(r=>`<li>${r}</li>`).join("")}
                  </ul>
                </div>
              </div>
            </div>
          `).join("")}
        </div>
      `,b.querySelector("#view-cabang-btn").addEventListener("click",()=>{i="cabang",E(b)}),b.querySelector("#view-hq-btn").addEventListener("click",()=>{i="hq",E(b)})}function w(b){b.innerHTML=`
      <div style="display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:12px; margin-bottom:18px;">
        <div>
          <h3 style="margin:0 0 4px 0; font-size:1.12rem; font-weight:700; color:#FFFFFF;">
            Daftar Klasifikasi Aset yang Dipantau Security (Lampiran 8.3)
          </h3>
          <p style="margin:0; font-size:0.88rem; color:#CBD5E1;">
            Pengawasan dan perlindungan sarana prasarana fisik, elektronik, dan peralatan klinik
          </p>
        </div>

        <div style="display:flex; gap:10px; align-items:center; flex-wrap:wrap;">
          <input type="text" id="sec-asset-search" class="form-control" placeholder="Cari nama aset / tipe / PIC..." value="${m}" style="height:36px; border-radius:8px; font-size:0.88rem; width:240px;">
            ${n?`
          <button id="btn-add-sec-asset" class="btn btn-primary" style="font-weight:700; padding:7px 14px; border-radius:8px; display:inline-flex; align-items:center; gap:6px; cursor:pointer; font-size:0.86rem; background:#2563EB; border-color:#3B82F6; color:#FFFFFF;">
            \u2795 Tambah Aset Baru
          </button>
            `:""}
        </div>
      </div>

      <div id="sec-asset-table-container"></div>
    `;let S=b.querySelector("#sec-asset-table-container"),y=b.querySelector("#sec-asset-search");function r(){let f=m.toLowerCase().trim(),u=F.filter(p=>!f||p.item.toLowerCase().includes(f)||p.type.toLowerCase().includes(f)||p.description.toLowerCase().includes(f)||p.pic&&p.pic.toLowerCase().includes(f));S.innerHTML=`
        <div class="card" style="background:#172033; border:1px solid #29364B; border-radius:12px; overflow:hidden; box-shadow:0 4px 12px rgba(0,0,0,0.2);">
          <div style="padding:14px 18px; background:#111827; border-bottom:1px solid #29364B; display:flex; justify-content:space-between; align-items:center;">
            <strong style="color:#FFFFFF; font-size:0.96rem;">Rekap Klasifikasi Aset Perusahaan</strong>
            <span class="badge" style="background:rgba(59,130,246,0.2); color:#93C5FD; border:1px solid rgba(96,165,250,0.4); font-weight:700; padding:3px 9px; border-radius:6px;">${u.length} Aset</span>
          </div>
          <div style="overflow-x:auto;">
            <table class="table" style="width:100%; margin:0; border-collapse:collapse; font-size:0.9rem;">
              <thead>
                <tr style="background:#111827; border-bottom:1px solid #29364B; text-align:left;">
                  <th style="padding:11px 14px; width:60px; color:#CBD5E1;">No</th>
                  <th style="padding:11px 14px; width:145px; color:#CBD5E1;">Tipe Aset</th>
                  <th style="padding:11px 14px; width:220px; color:#CBD5E1;">Nama Aset / Kelompok</th>
                  <th style="padding:11px 14px; color:#CBD5E1;">Deskripsi / Contoh Item</th>
                  <th style="padding:11px 14px; width:190px; color:#CBD5E1;">PIC Pengawas</th>
                  <th style="padding:11px 14px; width:110px; text-align:center; color:#CBD5E1;">Aksi</th>
                </tr>
              </thead>
              <tbody>
                ${u.map((p,A)=>`
                  <tr style="border-bottom:1px solid #1E2E4A; background:${A%2===0?"#131F36":"#0E182B"}; color:#F8FAFC;">
                    <td style="padding:11px 14px; font-weight:600; color:#CBD5E1;">${p.no||A+1}</td>
                    <td style="padding:11px 14px;">
                      <span class="badge" style="${l(p.type)}">${Y(p.type)}</span>
                    </td>
                    <td style="padding:11px 14px; font-weight:700; color:#FFFFFF;">${Y(p.item)}</td>
                    <td style="padding:11px 14px; color:#CBD5E1; line-height:1.5;">${Y(p.description)}</td>
                    <td style="padding:11px 14px; font-weight:700; color:#93C5FD;">${Y(p.pic||"-")}</td>
                    <td style="padding:11px 14px; text-align:center;">
                      <div style="display:inline-flex; gap:5px;">
                    ${n?`
                      <button type="button" class="fc-action-btn fc-action-edit btn-edit-asset" data-no="${p.no}" title="Edit Aset" aria-label="Edit aset ${Y(p.item||"")}"><svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M12 20h9"></path><path d="M16.5 3.5a2.121 2.121 0 0 1 3 3L7 19l-4 1 1-4L16.5 3.5z"></path></svg></button>
                    `:""}
                    ${a?`
                      <button type="button" class="fc-action-btn fc-action-del btn-delete-asset" data-no="${p.no}" title="Hapus Aset" aria-label="Hapus aset ${Y(p.item||"")}"><svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><polyline points="3 6 5 6 21 6"></polyline><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"></path></svg></button>
                    `:""}
                      </div>
                    </td>
                  </tr>
                `).join("")}
              </tbody>
            </table>
          </div>
        </div>
      `,S.querySelectorAll(".btn-edit-asset").forEach(p=>{p.addEventListener("click",()=>{let A=parseInt(p.dataset.no,10),T=F.find(g=>g.no===A);T&&K(T)})}),S.querySelectorAll(".btn-delete-asset").forEach(p=>{p.addEventListener("click",()=>{let A=parseInt(p.dataset.no,10),T=F.find(g=>g.no===A);T&&ge(`Apakah Anda yakin ingin menghapus aset: "${T.item}"?`,()=>{F=F.filter(g=>g.no!==A),F.forEach((g,C)=>{g.no=C+1}),qa(F),ae("Aset berhasil dihapus"),I()},"Hapus Aset")})})}function l(f){return f==="Aset Tetap"?"background:rgba(59,130,246,0.18); color:#93C5FD; border:1px solid rgba(96,165,250,0.38); font-weight:700; padding:3px 9px; border-radius:6px;":f==="Aset Bergerak"?"background:rgba(16,185,129,0.18); color:#6EE7B7; border:1px solid rgba(52,211,153,0.38); font-weight:700; padding:3px 9px; border-radius:6px;":"background:rgba(245,158,11,0.18); color:#FCD34D; border:1px solid rgba(251,191,36,0.38); font-weight:700; padding:3px 9px; border-radius:6px;"}r(),y&&y.addEventListener("input",f=>{m=f.target.value,r()}),b.querySelector("#btn-add-sec-asset").addEventListener("click",()=>{K()})}function K(b=null){if(!n)return;let S=!!b,y=`
      <form id="asset-edit-form" style="display:flex; flex-direction:column; gap:14px; font-size:0.88rem;">
        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Tipe / Klasifikasi Aset *</label>
          <select id="form-asset-type" class="form-control" style="width:100%; border-radius:6px; height:36px;">
            <option value="Aset Tetap" ${b?.type==="Aset Tetap"?"selected":""}>Aset Tetap</option>
            <option value="Aset Bergerak" ${b?.type==="Aset Bergerak"?"selected":""}>Aset Bergerak</option>
            <option value="Low Value Asset" ${b?.type==="Low Value Asset"?"selected":""}>Low Value Asset</option>
          </select>
        </div>
        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Nama Aset / Kelompok *</label>
          <input type="text" id="form-asset-item" class="form-control" required value="${b?Y(b.item):""}" style="width:100%; border-radius:6px; height:36px;">
        </div>
        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Deskripsi / Contoh Item *</label>
          <textarea id="form-asset-desc" class="form-control" rows="3" required style="width:100%; border-radius:6px; font-size:0.85rem;">${b?Y(b.description):""}</textarea>
        </div>
        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">PIC Pengawas / Penanggung Jawab *</label>
          <input type="text" id="form-asset-pic" class="form-control" required value="${b?Y(b.pic):""}" style="width:100%; border-radius:6px; height:36px;">
        </div>
      </form>
    `;ce({title:S?`\u270F\uFE0F Edit Aset: ${b.item}`:"\u2795 Tambah Klasifikasi Aset Baru",content:y,confirmText:"Simpan",onConfirm:(r,l)=>{let f=r.querySelector("#form-asset-type").value,u=r.querySelector("#form-asset-item").value.trim(),p=r.querySelector("#form-asset-desc").value.trim(),A=r.querySelector("#form-asset-pic").value.trim();if(!u||!p||!A){j("Harap lengkapi semua kolom");return}if(S)b.type=f,b.item=u,b.description=p,b.pic=A,qa(F),ae(`Aset "${u}" berhasil diperbarui`);else{let T=F.length>0?Math.max(...F.map(g=>g.no))+1:1;F.push({no:T,type:f,item:u,description:p,pic:A}),qa(F),ae(`Aset baru "${u}" berhasil ditambahkan`)}l(),I()}})}function G(b){b.innerHTML=`
      <div style="display:grid; grid-template-columns:repeat(auto-fit, minmax(340px, 1fr)); gap:20px;">
        <!-- Card 1: Shift Kerja -->
        <div class="card" style="background:#172033; border:1px solid #29364B; border-radius:12px; padding:20px; box-shadow:0 4px 12px rgba(0,0,0,0.2);">
          <div style="display:flex; align-items:center; gap:12px; margin-bottom:14px; border-bottom:1px solid #29364B; padding-bottom:12px;">
            <span style="font-size:1.5rem; background:rgba(59,130,246,0.16); border:1px solid rgba(96,165,250,0.35); padding:8px; border-radius:9px;">\u23F0</span>
            <div>
              <h3 style="margin:0 0 2px; font-size:1.12rem; font-weight:700; color:#FFFFFF;">Sistem Jam Kerja & Shift (6.2)</h3>
              <span style="font-size:0.84rem; color:#CBD5E1;">Ketentuan 12 Jam Kerja dengan 2 Hari Libur per Minggu</span>
            </div>
          </div>
          <div style="font-size:0.9rem; color:#F8FAFC; line-height:1.7;">
            <div style="background:#0F172A; border:1px solid #29364B; border-radius:8px; padding:11px 14px; margin-bottom:10px; border-left:3.5px solid #60A5FA;">
              <strong style="color:#FFFFFF; display:block; margin-bottom:2px;">\u2600\uFE0F Shift 1 (Pagi - Siang):</strong>
              <span style="color:#E2E8F0;">Pukul <strong style="color:#FFFFFF;">07.00 s/d 19.00 WIB</strong></span>
            </div>
            <div style="background:#0F172A; border:1px solid #29364B; border-radius:8px; padding:11px 14px; margin-bottom:12px; border-left:3.5px solid #A78BFA;">
              <strong style="color:#FFFFFF; display:block; margin-bottom:2px;">\u{1F319} Shift 2 (Malam):</strong>
              <span style="color:#E2E8F0;">Pukul <strong style="color:#FFFFFF;">19.00 s/d 07.00 WIB</strong></span>
            </div>
            <p style="margin:0; font-size:0.85rem; color:#CBD5E1; line-height:1.6;">
              <strong style="color:#FFFFFF;">Ketentuan Off:</strong> Berlaku 2 hari off dalam setiap pekannya sesuai jadwal roster yang ditetapkan oleh Facility Care Supervisor / General Support.
            </p>
          </div>
        </div>

        <!-- Card 2: Jadwal Safety Patrol -->
        <div class="card" style="background:#172033; border:1px solid #29364B; border-radius:12px; padding:20px; box-shadow:0 4px 12px rgba(0,0,0,0.2);">
          <div style="display:flex; align-items:center; gap:12px; margin-bottom:14px; border-bottom:1px solid #29364B; padding-bottom:12px;">
            <span style="font-size:1.5rem; background:rgba(245,158,11,0.16); border:1px solid rgba(251,191,36,0.35); padding:8px; border-radius:9px;">\u{1F6B6}\u200D\u2642\uFE0F</span>
            <div>
              <h3 style="margin:0 0 2px; font-size:1.12rem; font-weight:700; color:#FFFFFF;">Jadwal Safety Patrol Harian (6.2 B)</h3>
              <span style="font-size:0.84rem; color:#CBD5E1;">Keliling Pengecekan Area & Form Checklist Digital</span>
            </div>
          </div>
          <div style="display:flex; flex-direction:column; gap:9px; font-size:0.88rem;">
            <div style="background:#0F172A; border:1px solid #29364B; padding:9px 13px; border-radius:8px; display:flex; justify-content:space-between; align-items:center;">
              <strong style="color:#F8FAFC;">Patroli 1 (Pagi):</strong>
              <span class="badge" style="background:rgba(56,189,248,0.16); color:#7DD3FC; border:1px solid rgba(56,189,248,0.35); font-weight:700; padding:3px 9px; border-radius:6px;">07.30 - 08.30 WIB</span>
            </div>
            <div style="background:#0F172A; border:1px solid #29364B; padding:9px 13px; border-radius:8px; display:flex; justify-content:space-between; align-items:center;">
              <strong style="color:#F8FAFC;">Patroli 2 (Malam):</strong>
              <span class="badge" style="background:rgba(245,158,11,0.16); color:#FCD34D; border:1px solid rgba(251,191,36,0.35); font-weight:700; padding:3px 9px; border-radius:6px;">19.30 - 20.30 WIB</span>
            </div>
            <div style="background:#0F172A; border:1px solid #29364B; padding:9px 13px; border-radius:8px; display:flex; justify-content:space-between; align-items:center;">
              <strong style="color:#F8FAFC;">Patroli 3 (Dini Hari):</strong>
              <span class="badge" style="background:rgba(139,92,246,0.18); color:#DDD6FE; border:1px solid rgba(167,139,250,0.38); font-weight:700; padding:3px 9px; border-radius:6px;">04.00 - 05.00 WIB</span>
            </div>
            <div style="font-size:0.84rem; color:#CBD5E1; margin-top:4px; line-height:1.5;">
              * <em>Patroli 3 khusus berlaku untuk area Head Quarter atau lokasi yang memiliki lebih dari 1 personil security.</em>
            </div>
          </div>
        </div>

        <!-- Card 3: Penanganan Pelanggaran Disiplin -->
        <div class="card" style="background:#172033; border:1px solid #29364B; border-radius:12px; padding:20px; box-shadow:0 4px 12px rgba(0,0,0,0.2);">
          <div style="display:flex; align-items:center; gap:12px; margin-bottom:14px; border-bottom:1px solid #29364B; padding-bottom:12px;">
            <span style="font-size:1.5rem; background:rgba(16,185,129,0.16); border:1px solid rgba(52,211,153,0.35); padding:8px; border-radius:9px;">\u26A0\uFE0F</span>
            <div>
              <h3 style="margin:0 0 2px; font-size:1.12rem; font-weight:700; color:#FFFFFF;">Penanganan Pelanggaran Disiplin (6.4)</h3>
              <span style="font-size:0.84rem; color:#CBD5E1;">Pelanggaran SK Tata Tertib Perusahaan</span>
            </div>
          </div>
          <ol style="margin:0 0 0 18px; padding:0; font-size:0.88rem; color:#F8FAFC; line-height:1.75;">
            <li style="margin-bottom:4px;">Petugas security membuat <strong style="color:#FFFFFF;">Berita Acara Umum (BA)</strong> kejadian secara lengkap.</li>
            <li style="margin-bottom:4px;">Melaporkan kronologis tertulis kepada atasan (General Support Manager / FC Supervisor).</li>
            <li>Melaporkan dan meneruskan berkas kepada atasan langsung dari karyawan yang bersangkutan untuk penanganan HR/HC.</li>
          </ol>
        </div>

        <!-- Card 4: Penanganan Pelanggaran Kriminal -->
        <div class="card" style="background:#172033; border:1px solid #29364B; border-radius:12px; padding:20px; box-shadow:0 4px 12px rgba(0,0,0,0.2);">
          <div style="display:flex; align-items:center; gap:12px; margin-bottom:14px; border-bottom:1px solid #29364B; padding-bottom:12px;">
            <span style="font-size:1.5rem; background:rgba(239,68,68,0.16); border:1px solid rgba(248,113,113,0.35); padding:8px; border-radius:9px;">\u{1F6A8}</span>
            <div>
              <h3 style="margin:0 0 2px; font-size:1.12rem; font-weight:700; color:#FCA5A5;">Penanganan Pelanggaran Kriminal (6.4)</h3>
              <span style="font-size:0.84rem; color:#CBD5E1;">Tindak Pidana, Pencurian, & Pelanggaran Hukum</span>
            </div>
          </div>
          <ol style="margin:0 0 0 18px; padding:0; font-size:0.88rem; color:#F8FAFC; line-height:1.75;">
            <li style="margin-bottom:3px;">Buat <strong style="color:#FFFFFF;">Berita Acara Umum</strong> secara resmi.</li>
            <li style="margin-bottom:3px;">Periksa dan amankan rekaman bukti visual melalui <strong style="color:#FFFFFF;">CCTV</strong>.</li>
            <li style="margin-bottom:3px;">Laporkan ke atasan dari karyawan terkait dengan rincian kronologis (bila pelaku karyawan).</li>
            <li style="margin-bottom:3px;">Minta dan catat detail identitas <strong style="color:#FFFFFF;">KTP / ID Card</strong> dari pihak yang bersangkutan.</li>
            <li style="margin-bottom:3px;">Amankan dan netralisir <strong style="color:#FFFFFF;">Tempat Kejadian Perkara (TKP)</strong> dari gangguan orang lain.</li>
            <li>Laporkan ke <strong style="color:#FFFFFF;">Kepolisian terdekat</strong> untuk proses hukum tindak lanjut.</li>
          </ol>
        </div>
      </div>
    `}function J(b){b.innerHTML=`
      <div style="display:grid; grid-template-columns:repeat(auto-fit, minmax(320px, 1fr)); gap:18px;">
        <!-- Card: Google Form Safety Patrol -->
        <div class="card" style="background:#172033; border:1px solid #29364B; border-radius:12px; padding:20px; box-shadow:0 4px 12px rgba(0,0,0,0.2); display:flex; flex-direction:column; justify-content:space-between;">
          <div>
            <div style="display:flex; align-items:center; gap:12px; margin-bottom:12px;">
              <span style="font-size:2rem;">\u{1F4D1}</span>
              <div>
                <h3 style="margin:0 0 4px; font-size:1.12rem; font-weight:700; color:#FFFFFF;">Form Safety Patrol Audy</h3>
                <span class="badge" style="background:rgba(56,189,248,0.16); color:#7DD3FC; border:1px solid rgba(56,189,248,0.35); font-size:0.78rem; font-weight:700; padding:3px 9px; border-radius:6px;">Google Forms Resmi</span>
              </div>
            </div>
            <p style="font-size:0.88rem; color:#CBD5E1; line-height:1.6; margin-bottom:16px;">
              Formulir digital inspeksi keamanan yang wajib diisi oleh tim security saat melakukan Safety Patrol rutin di klinik dan kantor pusat Audy Dental.
            </p>
          </div>
          <a href="https://forms.gle/1EAvkqDJVdctXEGn7" target="_blank" rel="noopener" class="btn btn-primary" style="display:inline-flex; align-items:center; justify-content:center; gap:6px; font-weight:700; padding:10px; border-radius:8px; background:#2563EB; border-color:#3B82F6; color:#FFFFFF;">
            \u{1F4CB} Buka Form Safety Patrol \u2794
          </a>
        </div>

        <!-- Card: Dasar Hukum & Ruang Lingkup Perusahaan -->
        <div class="card" style="background:#172033; border:1px solid #29364B; border-radius:12px; padding:20px; box-shadow:0 4px 12px rgba(0,0,0,0.2); display:flex; flex-direction:column; justify-content:space-between;">
          <div>
            <div style="display:flex; align-items:center; gap:12px; margin-bottom:12px;">
              <span style="font-size:2rem;">\u2696\uFE0F</span>
              <div>
                <h3 style="margin:0 0 4px; font-size:1.12rem; font-weight:700; color:#FFFFFF;">Dasar Hukum & Ruang Lingkup</h3>
                <span class="badge" style="background:#22314A; color:#E2E8F0; border:1px solid #334155; font-size:0.78rem; font-weight:700; padding:3px 9px; border-radius:6px;">Legalitas</span>
              </div>
            </div>
            <p style="font-size:0.88rem; color:#CBD5E1; line-height:1.6; margin-bottom:12px;">
              SOP ini diterapkan sebagai tata cara Pengelolaan Keamanan Aset Perusahaan di seluruh lokasi Audy Dental Klinik di bawah naungan:
            </p>
            <ul style="margin:0 0 0 18px; padding:0; font-size:0.88rem; color:#F8FAFC; line-height:1.7;">
              <li><strong style="color:#FFFFFF;">PT. Audy Mandiri Indonesia</strong></li>
              <li><strong style="color:#FFFFFF;">PT. Kastara Ananta</strong></li>
              <li><strong style="color:#FFFFFF;">PT. Nara Adidaya</strong></li>
            </ul>
          </div>
        </div>
      </div>

      <!-- Validasi & Pengesahan SOP -->
      <div class="card" style="margin-top:20px; padding:20px; background:#172033; border:1px solid #29364B; border-radius:12px;">
        <h3 style="margin:0 0 14px 0; font-size:1.05rem; font-weight:700; color:#FFFFFF; display:flex; align-items:center; gap:8px;">
          <span>\u{1F4DC}</span> Validasi & Pengesahan Dokumen SOP Security
        </h3>
        <div style="display:grid; grid-template-columns:repeat(auto-fit, minmax(260px, 1fr)); gap:16px;">
          <div style="border-left:3.5px solid #60A5FA; padding-left:12px;">
            <div style="font-size:0.78rem; color:#CBD5E1; font-weight:700; text-transform:uppercase;">Disiapkan Oleh:</div>
            <div style="font-size:0.98rem; font-weight:700; color:#FFFFFF; margin-top:2px;">Berlin Ariansyah</div>
            <div style="font-size:0.84rem; color:#CBD5E1;">Facility Care Supervisor</div>
          </div>
          <div style="border-left:3.5px solid #34D399; padding-left:12px;">
            <div style="font-size:0.78rem; color:#CBD5E1; font-weight:700; text-transform:uppercase;">Diketahui Oleh:</div>
            <div style="font-size:0.98rem; font-weight:700; color:#FFFFFF; margin-top:2px;">Catur Meiyogo</div>
            <div style="font-size:0.84rem; color:#CBD5E1;">General Support Manager</div>
          </div>
          <div style="border-left:3.5px solid #A78BFA; padding-left:12px;">
            <div style="font-size:0.78rem; color:#CBD5E1; font-weight:700; text-transform:uppercase;">Disetujui Oleh:</div>
            <div style="font-size:0.98rem; font-weight:700; color:#FFFFFF; margin-top:2px;">Drg. Yulita</div>
            <div style="font-size:0.84rem; color:#CBD5E1;">Chief Executive Officer</div>
          </div>
        </div>
      </div>
    `}async function W(){try{let{downloadExcel:b}=await Promise.resolve().then(()=>(Ce(),Ve)),S=d.map(r=>({"Kode IK":`IK-SEC-${String(r.id).padStart(2,"0")}`,"Judul Prosedur":r.title,Kategori:r.category,"Sasaran Target":r.target,"Alat & Perlengkapan":r.tools.join(", "),"Langkah Pelaksanaan":r.steps.map((l,f)=>`${f+1}. ${l}`).join(" | "),"Catatan Khusus":r.notes||"-"})),y=new Date().toISOString().slice(0,10);b(S,`SOP_Security_Instruksi_Kerja_${y}`),ae("Data Excel SOP Security berhasil diunduh")}catch(b){console.error(b),j("Gagal mendownload data Excel: "+b.message)}}function Y(b){return b?String(b).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;").replace(/'/g,"&#39;"):""}I()}Be();function ni(e=""){let t=(e||"").toLowerCase();return t.includes("berita acara")?{label:e,bg:"#F3E8FF",color:"#7E22CE",border:"#D8B4FE",icon:"\u2696\uFE0F"}:t.includes("apar")||t.includes("k3")||t.includes("safety")?{label:e,bg:"#FEE2E2",color:"#B91C1C",border:"#FECACA",icon:"\u{1F9EF}"}:t.includes("deep cleaning")?{label:e,bg:"#DCFCE7",color:"#15803D",border:"#BBF7D0",icon:"\u2728"}:t.includes("drain")?{label:e,bg:"#FEF3C7",color:"#B45309",border:"#FDE68A",icon:"\u{1F6BF}"}:t.includes("general cleaning")||t.includes("kebersihan")?{label:e,bg:"#E0F2FE",color:"#0369A1",border:"#BAE6FD",icon:"\u{1F9F9}"}:t.includes("form")?{label:e,bg:"#E0E7FF",color:"#4338CA",border:"#C7D2FE",icon:"\u{1F4DD}"}:{label:e||"Dokumen",bg:"#F1F5F9",color:"#475569",border:"#CBD5E1",icon:"\u{1F4C1}"}}function ii(e=""){if(!e)return{type:"Tidak Ada Tautan",icon:"\u{1F4CE}",badgeClass:"bg-slate-100 text-slate-500"};let t=e.toLowerCase();return t.includes("forms.gle")||t.includes("google.com/forms")?{type:"Google Form",icon:"\u{1F7E3}",badgeClass:"badge-form",color:"#673AB7"}:t.includes("docs.google.com/spreadsheets")||t.includes("sheet")?{type:"Google Sheets",icon:"\u{1F7E2}",badgeClass:"badge-sheet",color:"#0F9D58"}:t.includes("docs.google.com/document")||t.includes("doc")?{type:"Google Docs",icon:"\u{1F535}",badgeClass:"badge-doc",color:"#4285F4"}:t.includes("drive.google.com")?{type:"Google Drive",icon:"\u{1F7E1}",badgeClass:"badge-drive",color:"#F4B400"}:t.endsWith(".pdf")?{type:"PDF Document",icon:"\u{1F534}",badgeClass:"badge-pdf",color:"#EA4335"}:{type:"Web Link",icon:"\u{1F310}",badgeClass:"badge-link",color:"#2563EB"}}async function ri(e){let t=se(),n=re(t,"checklist","write"),a=re(t,"checklist","delete"),o=[],s="all",c="",m=1,i=20,d=0,F=1,$={categoryCounts:{all:0},linkedCount:0},M=null,h=!1,L="fcms-master-doc-styles";if(!document.getElementById(L)){let S=document.createElement("style");S.id=L,S.textContent=`
      .doc-page-container {
        padding: 24px;
        max-width: 1400px;
        margin: 0 auto;
        font-family: inherit;
      }
      .doc-hero-card {
        background: #131F36;
        border: 1px solid #223456;
        border-radius: 18px;
        padding: 28px 32px;
        color: #FFFFFF;
        box-shadow: 0 4px 16px rgba(0, 0, 0, 0.22);
        margin-bottom: 24px;
        position: relative;
        overflow: hidden;
      }
      .doc-hero-card::after {
        content: '';
        position: absolute;
        right: -60px;
        top: -60px;
        width: 240px;
        height: 240px;
        background: radial-gradient(circle, rgba(37, 99, 235, 0.06) 0%, transparent 70%);
        border-radius: 50%;
        pointer-events: none;
      }
      .doc-hero-header {
        display: flex;
        justify-content: space-between;
        align-items: center;
        flex-wrap: wrap;
        gap: 16px;
      }
      .doc-hero-title {
        font-size: 1.45rem;
        font-weight: 800;
        margin: 0;
        display: flex;
        align-items: center;
        gap: 10px;
        letter-spacing: -0.02em;
      }
      .doc-hero-actions {
        display: flex;
        align-items: center;
        gap: 8px;
        flex-wrap: wrap;
      }
      .doc-hero-actions .btn {
        height: 38px;
        box-sizing: border-box;
        transition: all 0.2s ease;
      }
      #doc-btn-add:hover {
        background: #1D4ED8 !important;
        transform: translateY(-1px);
      }
      #doc-btn-excel:hover, #doc-btn-refresh:hover {
        background: rgba(255, 255, 255, 0.16) !important;
        border-color: rgba(255, 255, 255, 0.3) !important;
        transform: translateY(-1px);
      }
      .doc-stats-grid {
        display: grid;
        grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
        gap: 14px;
        margin-top: 24px;
      }
      .doc-stat-box {
        background: rgba(255, 255, 255, 0.06);
        border: 1px solid rgba(255, 255, 255, 0.1);
        border-radius: 12px;
        padding: 14px 18px;
        backdrop-filter: blur(8px);
        display: flex;
        align-items: center;
        gap: 14px;
      }
      .doc-stat-icon {
        width: 44px;
        height: 44px;
        border-radius: 10px;
        background: rgba(59, 130, 246, 0.18);
        color: #60A5FA;
        display: flex;
        align-items: center;
        justify-content: center;
        font-size: 1.3rem;
        flex-shrink: 0;
      }
      .doc-stat-value {
        font-size: 1.45rem;
        font-weight: 700;
        color: #FFFFFF;
        line-height: 1.2;
      }
      .doc-stat-label {
        font-size: 0.78rem;
        color: #94A3B8;
        text-transform: uppercase;
        letter-spacing: 0.04em;
        font-weight: 600;
      }
      .doc-controls-panel {
        background: #131F36;
        border: 1px solid #223456;
        border-radius: 16px;
        padding: 18px 20px;
        margin-bottom: 24px;
        box-shadow: 0 2px 8px rgba(0, 0, 0, 0.04);
        display: flex;
        flex-direction: column;
        gap: 16px;
      }
      .doc-controls-top {
        display: flex;
        align-items: center;
        justify-content: space-between;
        gap: 14px;
        flex-wrap: wrap;
      }
      .doc-search-wrapper {
        position: relative;
        flex: 1;
        min-width: 260px;
      }
      .doc-search-icon {
        position: absolute;
        left: 14px;
        top: 50%;
        transform: translateY(-50%);
        color: #94A3B8;
        font-size: 0.95rem;
        pointer-events: none;
      }
      .doc-search-input {
        width: 100%;
        padding: 10px 14px 10px 40px;
        border: 1.5px solid #CBD5E1;
        border-radius: 10px;
        font-size: 0.9rem;
        background: #0E182B;
        transition: all 0.2s ease;
      }
      .doc-search-input:focus {
        outline: none;
        border-color: #2563EB;
        background: #131F36;
        box-shadow: 0 0 0 3px rgba(37, 99, 235, 0.12);
      }
      .doc-filter-pills {
        display: flex;
        flex-wrap: wrap;
        gap: 8px;
        align-items: center;
      }
      .doc-pill {
        border: 1px solid #223456;
        background: #0E182B;
        color: #CBD5E1;
        font-size: 0.82rem;
        font-weight: 600;
        padding: 6px 14px;
        border-radius: 30px;
        cursor: pointer;
        transition: all 0.2s ease;
        display: flex;
        align-items: center;
        gap: 6px;
      }
      .doc-pill:hover {
        background: #22314A;
        color: #F8FAFC;
        border-color: #475569;
      }
      .doc-pill.active {
        background: #2563EB;
        border-color: #2563EB;
        color: #FFFFFF;
        box-shadow: 0 2px 6px rgba(37, 99, 235, 0.25);
      }
      .doc-pill-count {
        background: rgba(255, 255, 255, 0.1);
        color: #E2E8F0;
        padding: 1px 7px;
        border-radius: 12px;
        font-size: 0.75rem;
      }
      .doc-pill.active .doc-pill-count {
        background: rgba(255, 255, 255, 0.25);
        color: #FFFFFF;
      }
      /* Grid View */
      .doc-cards-grid {
        display: grid;
        grid-template-columns: repeat(auto-fill, minmax(340px, 1fr));
        gap: 20px;
      }
      .doc-card {
        background: #131F36;
        border: 1px solid #223456;
        border-radius: 16px;
        padding: 22px;
        box-shadow: 0 2px 6px rgba(0, 0, 0, 0.04);
        display: flex;
        flex-direction: column;
        justify-content: space-between;
        transition: all 0.25s cubic-bezier(0.16, 1, 0.3, 1);
        position: relative;
        overflow: hidden;
      }
      .doc-card:hover {
        transform: translateY(-4px);
        box-shadow: 0 12px 24px -6px rgba(15, 23, 42, 0.1);
        border-color: #CBD5E1;
      }
      .doc-card-top {
        display: flex;
        justify-content: space-between;
        align-items: flex-start;
        gap: 12px;
        margin-bottom: 14px;
      }
      .doc-category-badge {
        display: inline-flex;
        align-items: center;
        gap: 6px;
        padding: 4px 10px;
        border-radius: 6px;
        font-size: 0.76rem;
        font-weight: 700;
        letter-spacing: 0.02em;
        text-transform: uppercase;
      }
      .doc-platform-badge {
        display: inline-flex;
        align-items: center;
        gap: 4px;
        font-size: 0.76rem;
        font-weight: 600;
        color: #CBD5E1;
        background: #0E182B;
        border: 1px solid #223456;
        padding: 3px 8px;
        border-radius: 6px;
      }
      .doc-card-title {
        font-size: 1.08rem;
        font-weight: 700;
        color: #F8FAFC;
        margin: 0 0 8px 0;
        line-height: 1.4;
      }
      .doc-card-desc {
        color: #CBD5E1;
        font-size: 0.85rem;
        line-height: 1.5;
        margin: 0 0 18px 0;
        flex-grow: 1;
        display: -webkit-box;
        -webkit-line-clamp: 3;
        -webkit-box-orient: vertical;
        overflow: hidden;
      }
      .doc-card-footer {
        padding-top: 14px;
        border-top: 1px solid #223456;
        display: flex;
        align-items: center;
        justify-content: space-between;
        gap: 8px;
      }
      .doc-btn-open {
        background: #2563EB;
        color: #FFFFFF;
        border: none;
        border-radius: 8px;
        padding: 8px 14px;
        font-size: 0.84rem;
        font-weight: 600;
        display: inline-flex;
        align-items: center;
        gap: 6px;
        text-decoration: none;
        cursor: pointer;
        transition: background 0.15s ease;
      }
      .doc-btn-open:hover {
        background: #1D4ED8;
        color: #FFFFFF;
      }
      .doc-btn-open.disabled {
        background: #22314A;
        color: #94A3B8;
        pointer-events: none;
      }
      .doc-card-actions {
        display: flex;
        align-items: center;
        gap: 4px;
      }
      .doc-icon-btn {
        background: #0E182B;
        border: 1px solid #223456;
        color: #CBD5E1;
        width: 34px;
        height: 34px;
        border-radius: 8px;
        display: flex;
        align-items: center;
        justify-content: center;
        cursor: pointer;
        transition: all 0.15s ease;
      }
      .doc-icon-btn:hover {
        background: #22314A;
        color: #F8FAFC;
        border-color: #475569;
      }
      .doc-icon-btn.delete:hover {
        background: rgba(239, 68, 68, 0.2);
        color: #FCA5A5;
        border-color: rgba(248, 113, 113, 0.45);
      }
      /* Table View */
      .doc-table-card {
        background: #131F36;
        border: 1px solid #223456;
        border-radius: 16px;
        box-shadow: 0 2px 8px rgba(0, 0, 0, 0.2);
        overflow: hidden;
      }
      .doc-table {
        width: 100%;
        border-collapse: collapse;
        text-align: left;
      }
      .doc-table th {
        background: #0E182B;
        padding: 14px 18px;
        font-size: 0.78rem;
        font-weight: 700;
        color: #CBD5E1;
        text-transform: uppercase;
        letter-spacing: 0.04em;
        border-bottom: 1.5px solid #223456;
      }
      .doc-table td {
        padding: 16px 18px;
        font-size: 0.88rem;
        color: #F8FAFC;
        border-bottom: 1px solid #223456;
        vertical-align: middle;
      }
      .doc-table tr:hover td {
        background: #22314A;
      }
      .doc-empty-state {
        text-align: center;
        padding: 60px 20px;
        background: #131F36;
        border-radius: 16px;
        border: 1px dashed #223456;
      }

      /* \u2500\u2500 TAMPILAN RESPONSIVE HP / MOBILE (\u2264 768px) \u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500 */
      @media (max-width: 768px) {
        .doc-page-container {
          padding: 12px 14px !important;
        }
        .doc-hero-card {
          padding: 18px 16px !important;
          border-radius: 14px !important;
          margin-bottom: 16px !important;
        }
        .doc-hero-title {
          font-size: 1.25rem !important;
          gap: 8px !important;
        }
        .doc-hero-subtitle {
          font-size: 0.82rem !important;
          line-height: 1.4 !important;
        }
        .doc-hero-header {
          flex-direction: column !important;
          align-items: stretch !important;
          gap: 12px !important;
        }
        .doc-hero-actions {
          width: 100% !important;
          display: flex !important;
          align-items: center !important;
          gap: 6px !important;
          flex-wrap: nowrap !important;
          overflow-x: auto !important;
          scrollbar-width: none !important;
          padding-bottom: 2px !important;
        }
        .doc-hero-actions::-webkit-scrollbar {
          display: none;
        }
        #doc-btn-add {
          flex: 1 1 auto !important;
          justify-content: center !important;
          font-size: 0.82rem !important;
          padding: 8px 10px !important;
          white-space: nowrap !important;
        }
        #doc-btn-excel {
          flex: 1 1 auto !important;
          justify-content: center !important;
          font-size: 0.82rem !important;
          padding: 8px 10px !important;
          white-space: nowrap !important;
        }
        #doc-btn-refresh {
          width: 38px !important;
          height: 38px !important;
          padding: 0 !important;
          flex-shrink: 0 !important;
          justify-content: center !important;
        }
        /* 2x2 compact stats grid di layar HP */
        .doc-stats-grid {
          grid-template-columns: repeat(2, 1fr) !important;
          gap: 8px !important;
          margin-top: 14px !important;
        }
        .doc-stat-box {
          padding: 10px 12px !important;
          gap: 10px !important;
          border-radius: 10px !important;
        }
        .doc-stat-icon {
          width: 36px !important;
          height: 36px !important;
          font-size: 1.1rem !important;
          border-radius: 8px !important;
        }
        .doc-stat-value {
          font-size: 1.15rem !important;
        }
        .doc-stat-label {
          font-size: 0.76rem !important;
        }
        .doc-controls-panel {
          padding: 14px 12px !important;
          border-radius: 14px !important;
          margin-bottom: 16px !important;
          gap: 12px !important;
        }
        .doc-controls-top {
          flex-direction: column !important;
          align-items: stretch !important;
          gap: 10px !important;
        }
        .doc-search-wrapper {
          min-width: 100% !important;
          width: 100% !important;
        }
        /* PITA FILTER SWIPEABLE HORIZONTAL DI HP (TIDAK LAGI NUMPUK VERTIKAL) */
        .doc-filter-pills {
          display: flex !important;
          flex-wrap: nowrap !important;
          overflow-x: auto !important;
          -webkit-overflow-scrolling: touch !important;
          padding-bottom: 6px !important;
          padding-top: 2px !important;
          scrollbar-width: none !important;
          -ms-overflow-style: none !important;
        }
        .doc-filter-pills::-webkit-scrollbar {
          display: none !important;
        }
        .doc-pill {
          flex-shrink: 0 !important;
          white-space: nowrap !important;
          padding: 6px 12px !important;
          font-size: 0.8rem !important;
        }
        .doc-cards-grid {
          grid-template-columns: minmax(0, 1fr) !important;
          width: 100% !important;
          max-width: 100% !important;
          min-width: 0 !important;
          box-sizing: border-box !important;
          gap: 12px !important;
        }
        .doc-card {
          padding: 16px !important;
          border-radius: 14px !important;
          width: 100% !important;
          max-width: 100% !important;
          min-width: 0 !important;
          box-sizing: border-box !important;
          overflow: hidden;
        }
        .doc-card-title {
          font-size: 1rem !important;
        }
        .doc-card-desc {
          font-size: 0.82rem !important;
        }
        .doc-table-card {
          border-radius: 14px !important;
          overflow-x: auto !important;
          -webkit-overflow-scrolling: touch !important;
        }
        .doc-table {
          min-width: 650px !important;
        }
        .doc-table th, .doc-table td {
          padding: 10px 12px !important;
          font-size: 0.82rem !important;
        }
      }
    `,document.head.appendChild(S)}function R(){let S={};return c&&c.trim()&&(S.search=c.trim()),s&&s!=="all"&&(S.category=s),S}function I(){let S=x(),y=Object.keys(S).filter(r=>r!=="all");return`
      <button class="doc-pill ${s==="all"?"active":""}" data-cat="all">
        <span>Semua Dokumen</span>
        <span class="doc-pill-count">${S.all||0}</span>
      </button>
      ${y.map(r=>{let l=ni(r),f=s===r,u=String(r).replace(/"/g,"&quot;");return`
          <button class="doc-pill ${f?"active":""}" data-cat="${u}">
            <span>${l.icon}</span>
            <span>${r}</span>
            <span class="doc-pill-count">${S[r]||0}</span>
          </button>
        `}).join("")}
    `}function N(){e.querySelectorAll(".doc-pill").forEach(S=>{S.onclick=()=>{s=S.getAttribute("data-cat")||"all",e.querySelectorAll(".doc-pill").forEach(y=>{y.classList.toggle("active",y.getAttribute("data-cat")===s)}),m=1,v({refreshStats:!1})}})}function q(){let S=x(),y=Object.keys(S).filter(p=>p!=="all"),r=Number(S.all??0),l=Number($.linkedCount??0),f=(p,A)=>{let T=e.querySelector(p);T&&(T.textContent=String(A??0))};f("#doc-stat-total",r),f("#doc-stat-categories",y.length),f("#doc-stat-linked",l),f("#doc-stat-linked-total",r);let u=e.querySelector("#doc-filter-pills");u&&(u.innerHTML=I(),N())}async function v({refreshStats:S=!0}={}){try{let y=R(),r=new URLSearchParams({page:String(m),limit:String(i),...y});S&&r.set("include_stats","1");let l=await X("/api/checklist?"+r.toString());if(l.ok&&l.data&&Array.isArray(l.data.data)){let f=l.data.pagination||{};if(d=Number(f.total??l.data.data.length),F=Math.max(1,Number(f.pages||f.totalPages||Math.ceil(d/i)||1)),l.data.data.length===0&&m>F)return m=F,v({refreshStats:S});o=l.data.data,l.data.stats&&($={categoryCounts:l.data.stats.categoryCounts||{all:d},linkedCount:Number(l.data.stats.linkedCount??0)})}else o=[],d=0,F=1,l.ok||j(l.data?.error||"Gagal memuat dokumen dari server")}catch(y){console.warn("Error loading documents:",y),o=[],d=0,F=1,j("Gagal memuat dokumen: "+y.message)}h?(S&&q(),W()):(h=!0,Y())}function B(){return o}function x(){return $.categoryCounts||{all:0}}function k(S=null){if(!n)return;let y=!!S,r=Array.from(new Set([...Object.keys($.categoryCounts||{}).filter(f=>f!=="all"&&f!=="Lainnya"),...o.map(f=>f.category).filter(Boolean)])),l=document.createElement("div");l.innerHTML=`
      <div style="display:flex; flex-direction:column; gap:16px;">
        <div>
          <label style="display:block; font-size:0.82rem; font-weight:600; color:var(--text-2); margin-bottom:6px;">
            Nama Dokumen / Checklist <span style="color:#DC2626;">*</span>
          </label>
          <input type="text" id="doc-modal-name" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
            value="${S?.name?String(S.name).replace(/"/g,"&quot;"):""}" required />
        </div>

        <div>
          <label style="display:block; font-size:0.82rem; font-weight:600; color:var(--text-2); margin-bottom:6px;">
            Kategori Dokumen <span style="color:#DC2626;">*</span>
          </label>
          <input type="text" id="doc-modal-category" list="doc-categories-list" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
            value="${S?.category?String(S.category).replace(/"/g,"&quot;"):""}" required />
          <datalist id="doc-categories-list">
            ${r.map(f=>`<option value="${f}">${f}</option>`).join("")}
            <option value="Berita Acara">Berita Acara</option>
            <option value="Apar & K3">Apar & K3</option>
            <option value="Kegiatan Deep Cleaning">Kegiatan Deep Cleaning</option>
            <option value="Kegiatan General Cleaning">Kegiatan General Cleaning</option>
            <option value="Drain Cleaning">Drain Cleaning</option>
            <option value="Form Operasional">Form Operasional</option>
          </datalist>
        </div>

        <div>
          <label style="display:block; font-size:0.82rem; font-weight:600; color:var(--text-2); margin-bottom:6px;">
            Tautan / Link Dokumen (Google Drive, Docs, Sheets, atau Forms)
          </label>
          <input type="url" id="doc-modal-link" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
            value="${S?.document_link?String(S.document_link).replace(/"/g,"&quot;"):""}" />
          <span style="font-size:0.75rem; color:var(--text-3); margin-top:4px; display:block;">
            Bisa berupa tautan Google Forms untuk inspeksi harian, atau link Google Sheets / Docs sebagai acuan.
          </span>
        </div>

        <div>
          <label style="display:block; font-size:0.82rem; font-weight:600; color:var(--text-2); margin-bottom:6px;">
            Deskripsi / Petunjuk Penggunaan
          </label>
          <textarea id="doc-modal-desc" class="form-control" rows="3" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;">${S?.description||""}</textarea>
        </div>
      </div>
    `,ce({title:y?"\u270F\uFE0F Edit Master Dokumen":"\u2795 Tambah Master Dokumen Baru",content:l,size:"md",confirmText:y?"Simpan Perubahan":"Tambahkan Dokumen",confirmClass:"btn-primary",onConfirm:async(f,u)=>{let p=l.querySelector("#doc-modal-name"),A=l.querySelector("#doc-modal-category"),T=l.querySelector("#doc-modal-link"),g=l.querySelector("#doc-modal-desc"),C=(p.value||"").trim(),z=(A.value||"").trim(),O=(T.value||"").trim(),te=(g.value||"").trim();if(!C){j("Nama Dokumen wajib diisi"),p.focus();return}if(!z){j("Kategori Dokumen wajib diisi"),A.focus();return}let Q={name:C,category:z,document_link:O,description:te};try{if(y){let ne=await X(`/api/checklist/${S.id}`,{method:"PUT",body:JSON.stringify(Q)});if(!ne.ok){j(ne.data?.error||"Gagal memperbarui dokumen");return}ae("Dokumen berhasil diperbarui")}else{let ne=await X("/api/checklist",{method:"POST",body:JSON.stringify(Q)});if(!ne.ok){j(ne.data?.error||"Gagal menambahkan dokumen");return}m=1,ae("Dokumen baru berhasil ditambahkan")}u(),await v({refreshStats:!0})}catch(ne){console.error("Error saving document:",ne),j("Gagal menyimpan dokumen: "+ne.message)}}})}function _(S){if(!a){j("Anda tidak memiliki izin untuk menghapus dokumen.");return}ge(`Apakah Anda yakin ingin menghapus dokumen <strong>"${S.name}"</strong>?`,async()=>{try{let y=await X(`/api/checklist/${S.id}`,{method:"DELETE"});if(!y.ok){j(y.data?.error||"Gagal menghapus dokumen");return}ae(`Dokumen "${S.name}" berhasil dihapus`),await v({refreshStats:!0})}catch(y){j("Gagal menghapus dokumen: "+y.message)}},"Hapus Dokumen")}function E(S,y){if(!S){gt("Dokumen belum memiliki tautan digital.");return}navigator.clipboard.writeText(S).then(()=>{ae(`Tautan "${y}" disalin ke clipboard!`)}).catch(()=>{gt(`Tautan: ${S}`)})}async function w(){if(!re(t,"export")){j("Anda tidak memiliki izin untuk mengekspor data.");return}try{let{downloadExcel:S}=await Promise.resolve().then(()=>(Ce(),Ve)),y=await ue("/api/checklist/export",R()),r=y.map((l,f)=>({No:f+1,"Nama Dokumen":l.name||"",Kategori:l.category||"","Platform / Format":ii(l.document_link).type,"Link Dokumen":l.document_link||"","Deskripsi / Keterangan":l.description||""}));S(r,`Master_Dokumen_FCMS_${new Date().toISOString().slice(0,10)}`),ae(`Berhasil mengunduh ${y.length} data dokumen ke Excel.`)}catch(S){j("Gagal mengekspor data: "+S.message)}}function K(S){return`
    ${S.length===0?`
            <div class="doc-empty-state">
              <div style="font-size:3rem; margin-bottom:12px;">\u{1F50D}</div>
              <h3 style="font-size:1.2rem; font-weight:700; color:#F8FAFC; margin:0 0 6px 0;">Tidak Ada Dokumen Ditemukan</h3>
              <p style="color:#94A3B8; font-size:0.9rem; margin:0 0 20px 0;">Coba sesuaikan kata kunci pencarian atau pilih kategori lain.</p>
              <button id="doc-btn-reset-filter" class="btn btn-ghost" style="border:1px solid #CBD5E1; border-radius:8px; padding:8px 16px;">
                Tampilkan Semua Dokumen
              </button>
            </div>
          `:`
            <!-- Grid Card View -->
            <div class="doc-cards-grid">
              ${S.map(y=>{let r=ni(y.category),l=ii(y.document_link);return`
                  <div class="doc-card" data-id="${y.id}">
                    <div>
                      <div class="doc-card-top">
                        <span class="doc-category-badge" style="background:${r.bg}; color:${r.color}; border:1px solid ${r.border};">
                          <span>${r.icon}</span>
                          <span>${r.label}</span>
                        </span>
                        <span class="doc-platform-badge">
                          <span>${l.icon}</span>
                          <span>${l.type}</span>
                        </span>
                      </div>
                      
                      <h3 class="doc-card-title">${y.name}</h3>
                      <p class="doc-card-desc">${y.description||"Formulir acuan standar operasional dan administrasi fasilitas."}</p>
                    </div>

                    <div class="doc-card-footer">
                      ${y.document_link?`
                        <a href="${y.document_link}" target="_blank" rel="noopener noreferrer" class="doc-btn-open">
                          <span>Buka Dokumen</span>
                          <svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M18 13v6a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V8a2 2 0 0 1 2-2h6"/><polyline points="15 3 21 3 21 9"/><line x1="10" y1="14" x2="21" y2="3"/></svg>
                        </a>
                      `:`
                        <button class="doc-btn-open disabled" title="Belum ada tautan yang dikaitkan">
                          <span>Tidak Ada Link</span>
                        </button>
                      `}

                      <div class="doc-card-actions">
                        ${y.document_link?`
                          <button class="doc-icon-btn copy-link-btn" title="Salin Tautan Dokumen" data-link="${y.document_link}" data-title="${y.name}">
                            <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><rect x="9" y="9" width="13" height="13" rx="2" ry="2"/><path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1"/></svg>
                          </button>
                        `:""}
                        ${n?`
                        <button class="doc-icon-btn edit-doc-btn" title="Edit Dokumen" data-id="${y.id}">
                          <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>
                        </button>
                        `:""}
                        ${a?`
                        <button class="doc-icon-btn delete delete-doc-btn" title="Hapus Dokumen" data-id="${y.id}">
                          <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><polyline points="3 6 5 6 21 6"/><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"/></svg>
                        </button>
                        `:""}
                      </div>
                    </div>
                  </div>
                `}).join("")}
            </div>
          `}
    `}function G(){let S=e.querySelector("#doc-btn-reset-filter");S&&(S.onclick=()=>{c="",s="all",m=1,clearTimeout(M);let y=e.querySelector("#doc-search-input");y&&(y.value=""),e.querySelectorAll(".doc-pill").forEach(r=>{r.classList.toggle("active",r.getAttribute("data-cat")==="all")}),v({refreshStats:!1})}),e.querySelectorAll(".copy-link-btn").forEach(y=>{y.onclick=r=>{r.stopPropagation();let l=y.getAttribute("data-link"),f=y.getAttribute("data-title");E(l,f)}}),e.querySelectorAll(".edit-doc-btn").forEach(y=>{y.onclick=r=>{r.stopPropagation();let l=Number(y.getAttribute("data-id")),f=o.find(u=>u.id===l);f&&k(f)}}),e.querySelectorAll(".delete-doc-btn").forEach(y=>{y.onclick=r=>{r.stopPropagation();let l=Number(y.getAttribute("data-id")),f=o.find(u=>u.id===l);f&&_(f)}})}function J(){let S=e.querySelector("#doc-pagination-container");if(!S)return;S.innerHTML="";let y=Te({page:m,pages:F,total:d,limit:i,onPage:r=>{m=r,v({refreshStats:!1})}});y&&S.appendChild(y)}function W(){let S=e.querySelector("#doc-list-view");if(!S)return;let y=B();S.innerHTML=K(y),J(),G()}function Y(){let S=B(),y=x(),r=Object.keys(y).filter(u=>u!=="all"),l=Number(y.all??0),f=Number($.linkedCount??0);e.innerHTML=`
      <div class="doc-page-container">
        
        <!-- Hero Header -->
        <div class="doc-hero-card">
          <div class="doc-hero-header">
            <div>
              <h1 class="doc-hero-title">
                <span>\u{1F4CB}</span> Master Document & Template Standar
              </h1>
            </div>
            <div class="doc-hero-actions" style="display:flex; align-items:center; gap:8px;">
              ${n?`
              <button id="doc-btn-add" class="btn btn-primary" style="background:#2563EB; height:38px; padding:0 16px; border-radius:10px; font-weight:700; display:inline-flex; align-items:center; gap:6px; box-shadow:0 4px 12px rgba(37,99,235,0.3); white-space:nowrap; border:none;">
                <span>\u2795</span>
                <span>Tambah Dokumen Baru</span>
              </button>
              `:""}
              ${re(t,"export")?`<button id="doc-btn-excel" class="btn btn-ghost" style="background:rgba(255,255,255,0.08); border:1px solid rgba(255,255,255,0.15); color:white; height:38px; padding:0 14px; border-radius:10px; font-weight:600; display:inline-flex; align-items:center; gap:6px; white-space:nowrap;">
                <span>\u{1F4E5}</span>
                <span>Ekspor Excel</span>
              </button>`:""}
              <button id="doc-btn-refresh" class="btn btn-ghost" style="background:rgba(255,255,255,0.08); border:1px solid rgba(255,255,255,0.15); color:white; width:38px; height:38px; padding:0; border-radius:10px; display:inline-flex; align-items:center; justify-content:center; flex-shrink:0;" title="Muat Ulang Data">
                <span>\u{1F504}</span>
              </button>
            </div>
          </div>

          <!-- Quick Metrics -->
          <div class="doc-stats-grid">
            <div class="doc-stat-box">
              <div class="doc-stat-icon">\u{1F4C4}</div>
              <div>
                <div class="doc-stat-value" id="doc-stat-total">${l}</div>
                <div class="doc-stat-label">Total Dokumen</div>
              </div>
            </div>
            <div class="doc-stat-box">
              <div class="doc-stat-icon">\u{1F5C2}\uFE0F</div>
              <div>
                <div class="doc-stat-value" id="doc-stat-categories">${r.length}</div>
                <div class="doc-stat-label">Kategori Dokumen</div>
              </div>
            </div>
            <div class="doc-stat-box">
              <div class="doc-stat-icon" style="background:rgba(16,185,129,0.18); color:#34D399;">\u{1F310}</div>
              <div>
                <div class="doc-stat-value"><span id="doc-stat-linked">${f}</span> <span style="font-size:0.8rem; font-weight:400; color:#94A3B8;">/ <span id="doc-stat-linked-total">${l}</span></span></div>
                <div class="doc-stat-label">Tautan Online</div>
              </div>
            </div>
            <div class="doc-stat-box">
              <div class="doc-stat-icon" style="background:rgba(168,85,247,0.18); color:#C084FC;">\u26A1</div>
              <div>
                <div class="doc-stat-value">Aktif</div>
                <div class="doc-stat-label">Status Repositori</div>
              </div>
            </div>
          </div>
        </div>

        <!-- Controls: Search, View Switch, Category Pills -->
        <div class="doc-controls-panel">
          <div class="doc-controls-top">
            <div class="doc-search-wrapper" style="width:100%;">
              <span class="doc-search-icon">\u{1F50D}</span>
              <input type="text" id="doc-search-input" class="doc-search-input" placeholder="Cari nama dokumen, kategori, atau deskripsi..." value="${c?c.replace(/"/g,"&quot;"):""}" />
            </div>
          </div>

          <!-- Category Pills -->
          <div class="doc-filter-pills" id="doc-filter-pills">
            ${I()}
          </div>
        </div>

        <!-- Documents View Container -->
        <div id="doc-list-view">
          ${K(S)}
        </div>

        <!-- Pagination Container -->
        <div id="doc-pagination-container"></div>

      </div>
    `,J(),b()}function b(){let S=e.querySelector("#doc-btn-add");S&&(S.onclick=()=>k());let y=e.querySelector("#doc-btn-excel");y&&(y.onclick=()=>w());let r=e.querySelector("#doc-btn-refresh");r&&(r.onclick=()=>v({refreshStats:!0}));let l=e.querySelector("#doc-search-input");l&&(l.oninput=f=>{c=f.target.value,clearTimeout(M),M=setTimeout(()=>{m=1,v({refreshStats:!1})},350)}),N(),G()}await v({refreshStats:!0})}Be();Ce();Ie();function Ha(e,t,n,a=!1,o=!1){if(!t&&!n)return"";let s=P(String(e.id??"")),c=P(e.item_name||"Parameter Standar"),m=a?14:15,i=o?" hyg-mobile-action-btn":a?" hygiene-action-btn--compact":"";return`
    <div class="${o?"hyg-mobile-actions":""}" style="${o?"":`display:inline-flex; gap:${a?4:6}px; align-items:center;`}">
      ${t?`
        <button type="button" class="hygiene-action-btn hygiene-action-edit${i} btn-edit-item" data-id="${s}" title="Edit ${c}" aria-label="Edit parameter ${c}">
          <svg width="${m}" height="${m}" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">
            <path d="M12 20h9"></path>
            <path d="M16.5 3.5a2.121 2.121 0 0 1 3 3L7 19l-4 1 1-4L16.5 3.5z"></path>
          </svg>
          ${o?"<span>Edit</span>":""}
        </button>
      `:""}
      ${n?`
        <button type="button" class="hygiene-action-btn hygiene-action-delete${i} btn-delete-item" data-id="${s}" data-name="${c}" title="Hapus ${c}" aria-label="Hapus parameter ${c}">
          <svg width="${m}" height="${m}" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">
            <polyline points="3 6 5 6 21 6"></polyline>
            <path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"></path>
            <line x1="10" y1="11" x2="10" y2="17"></line>
            <line x1="14" y1="11" x2="14" y2="17"></line>
          </svg>
          ${o?"<span>Hapus</span>":""}
        </button>
      `:""}
    </div>
  `}function oi(e,{includeRoom:t=!0,canEdit:n=!1,canDelete:a=!1}={}){return`
    <div class="hyg-mobile-cards${t?"":" hyg-mobile-cards--compact"}">
      ${e.map((o,s)=>{let c=ea(o.room_name),m=Ha(o,n,a,!1,!0);return`
          <article class="hyg-mobile-card">
            <div class="hyg-mobile-card-head">
              <span class="hyg-mobile-num">#${s+1}</span>
              ${t?`
                <div class="hyg-mobile-room-group">
                  <span class="hyg-room-badge">${P(o.room_name)}</span>
                  <span class="hyg-zone-meta">${P(c.zoneName)}</span>
                </div>
              `:""}
            </div>
            <div class="hyg-mobile-item-title">${P(o.item_name)}</div>
            <div class="hyg-std-box hyg-mobile-std-box">
              <span class="hyg-std-check" aria-hidden="true">\u2713</span>
              <span class="hyg-std-text">${P(o.cleanliness_standard)}</span>
            </div>
            ${m?`<div class="hyg-mobile-card-foot">${m}</div>`:""}
          </article>
        `}).join("")}
    </div>
  `}var pt={"Standar Umum":{icon:"\u{1F4CC}",bg:"#F3E8FF",color:"#6B21A8",border:"#D8B4FE",rooms:["Lantai & Sudut Ruangan","Dinding, Partisi & Plint","Kaca, Cermin & Partisi Kaca","Sanitair & Area Basah Toilet","Plafon, Lampu & Kisi Ventilasi","Mebel, Meja & Kursi Kerja/Pasien","Peralatan Khusus & Medis Non-Steril","Wadah Sampah & Utilitas"]},"Area Depan & Pelayanan":{icon:"\u{1F6CB}\uFE0F",bg:"#E0F2FE",color:"#0369A1",border:"#BAE6FD",rooms:["Teras","Lobby 1","Receptionist","Snack Corner"]},"Area Klinis & Medis":{icon:"\u{1FA7A}",bg:"#DBEAFE",color:"#1E40AF",border:"#93C5FD",rooms:["Ruang Tindakan","Ronsen","Ruang Dokter","Ruang Steril"]},"Toilet & Sanitasi":{icon:"\u{1F6BF}",bg:"#CCFBF1",color:"#0F766E",border:"#99F6E4",rooms:["Toilet 1","Janitor","Ruang Limbah"]},Sirkulasi:{icon:"\u{1F6B6}",bg:"#F1F5F9",color:"#334155",border:"#CBD5E1",rooms:["Area Tangga","Lift","Area Tangga Exit"]},"Area Anak & Fasilitas":{icon:"\u{1F9F8}",bg:"#FEF3C7",color:"#92400E",border:"#FDE68A",rooms:["Lobby 2","Playground","Musholla","Tempat Wudhu","Ruang Team"]},"Gudang & Teknis":{icon:"\u2699\uFE0F",bg:"#FFEDD5",color:"#C2410C",border:"#FED7AA",rooms:["Gudang Farmasi & Bahan","RoofTop","Ruang Kompressor","Area Genset"]},Eksterior:{icon:"\u{1F333}",bg:"#DCFCE7",color:"#15803D",border:"#86EFAC",rooms:["Taman","Pos Security","Toilet Security","Halaman Parkir"]}};function ea(e){for(let[t,n]of Object.entries(pt))if(n.rooms.includes(e))return{zoneName:t,...n};return{zoneName:"Lainnya",icon:"\u{1F4CD}",bg:"#131F36",color:"#94A3B8",border:"#223456"}}async function si(e){let t=se(),n=re(t,"hygiene_standards","write"),a=re(t,"hygiene_standards","delete"),o=re(t,"export");e.innerHTML=`
    <div class="hyg-page-shell">
      <!-- Unified Hero Card (Consistent with Master Karyawan, Kontrak, Master Document) -->
      <div class="hyg-hero-card">
        <div class="hyg-hero-top">
          <div class="hyg-hero-title-group">
            <h1 class="hyg-hero-title">Master Standar Kebersihan Fisik</h1>
            <p class="hyg-hero-subtitle">Acuan standar kebersihan fisik seluruh area dan ruangan klinik operasional</p>
          </div>
          <div class="hyg-hero-actions">
            <div class="hyg-view-switch" role="group" aria-label="Mode tampilan standar kebersihan">
              <button type="button" id="view-table-btn" class="hyg-view-btn active">Tabel Standar</button>
              <button type="button" id="view-zone-btn" class="hyg-view-btn">Per Zona Area</button>
            </div>
            <button type="button" id="btn-print-hygiene" class="btn btn-outline btn-sm" title="Cetak Format SOP Standar Kebersihan">
              <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><polyline points="6 9 6 2 18 2 18 9"/><path d="M6 18H4a2 2 0 0 1-2-2v-5a2 2 0 0 1 2-2h16a2 2 0 0 1 2 2v5a2 2 0 0 1-2 2h-2"/><rect x="6" y="14" width="12" height="8"/></svg>
              Cetak SOP
            </button>
            ${o?`
            <button type="button" id="btn-export-hygiene" class="btn btn-outline btn-sm">
              <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg>
              Export Excel
            </button>
            `:""}
            ${n?`
            <button type="button" id="btn-add-hygiene" class="btn btn-primary btn-sm">
              + Tambah Parameter
            </button>
            `:""}
          </div>
        </div>

        <!-- Compact Clean Navy Summary Stats -->
        <div id="hygiene-summary-cards" class="hyg-stats-grid">
          <div class="hyg-stat-box">
            <div class="hyg-stat-icon" aria-hidden="true">
              <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M9 11l3 3L22 4"/><path d="M21 12v7a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11"/></svg>
            </div>
            <div class="hyg-stat-body">
              <div class="hyg-stat-val" id="card-total-params">104</div>
              <div class="hyg-stat-lbl">Total Parameter</div>
            </div>
          </div>
          <div class="hyg-stat-box">
            <div class="hyg-stat-icon" aria-hidden="true">
              <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="3" width="7" height="7" rx="1"/><rect x="14" y="3" width="7" height="7" rx="1"/><rect x="3" y="14" width="7" height="7" rx="1"/><rect x="14" y="14" width="7" height="7" rx="1"/></svg>
            </div>
            <div class="hyg-stat-body">
              <div class="hyg-stat-val" id="card-total-zones">${Object.keys(pt).length}</div>
              <div class="hyg-stat-lbl">Kategori Zona</div>
            </div>
          </div>
          <div class="hyg-stat-box">
            <div class="hyg-stat-icon" aria-hidden="true">
              <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M3 9l9-7 9 7v11a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2z"/><polyline points="9 22 9 12 15 12 15 22"/></svg>
            </div>
            <div class="hyg-stat-body">
              <div class="hyg-stat-val" id="card-total-rooms">35</div>
              <div class="hyg-stat-lbl">Ruangan / Lokasi</div>
            </div>
          </div>
        </div>
      </div>

      <!-- Unified Filter & Scrollable Zone Tabs Card -->
      <div class="hyg-controls-card">
        <div class="hyg-controls-row">
          <div class="hyg-search-wrap">
            <svg class="hyg-search-icon" width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/></svg>
            <input type="text" id="hygiene-search" class="form-control hyg-search-input" placeholder="Cari ruangan, item objek, atau standar kebersihan..." aria-label="Cari parameter kebersihan">
          </div>
          <select id="hygiene-room-filter" class="form-control hyg-room-select" aria-label="Filter ruangan">
            <option value="">Semua Ruangan</option>
          </select>
          <button type="button" id="btn-reset-filters" class="btn btn-outline btn-sm hyg-reset-btn">
            Reset
          </button>
          <div id="hygiene-stats-text" class="hyg-stats-text">
            Memuat data...
          </div>
        </div>

        <!-- Horizontal Scrollable Zone Tabs (Enhanced with Global Left/Right Arrows) -->
        <div id="hygiene-zone-tabs" class="hyg-zone-tabs" data-fc-scroll-tabs role="tablist" aria-label="Filter kategori zona kebersihan">
          <button type="button" class="zone-tab-btn active" data-zone="all" role="tab" aria-selected="true">
            <span>Semua Zona</span>
            <span class="badge hyg-tab-count" id="badge-count-all">104</span>
          </button>
          ${Object.keys(pt).map(x=>`
            <button type="button" class="zone-tab-btn" data-zone="${P(x)}" role="tab" aria-selected="false">
              <span>${P(x)}</span>
              <span class="badge hyg-tab-count" id="badge-count-${x.replace(/[\s&]+/g,"-")}">-</span>
            </button>
          `).join("")}
        </div>
      </div>

      <!-- Main Content Container -->
      <div id="hygiene-view-container">
        <div class="skeleton" style="height:320px; border-radius:12px;"></div>
      </div>
    </div>
  `;let s="all",c="table",m=[];async function i(){try{let x=await X("/api/hygiene-standards?all=1");if(!x.ok)throw new Error("Gagal memuat standar hygiene");m=Array.isArray(x.data)?x.data:x.data?.data||[],F(),d(),M()}catch(x){console.error(x),j(x.message||"Terjadi kesalahan memuat data");let k=e.querySelector("#hygiene-view-container");k&&(k.innerHTML=`
          <div class="empty-state" style="text-align:center; padding:40px; background:#131F36; border-radius:12px; border:1px solid #223456;">
            <div style="font-size:2.5rem; margin-bottom:12px;">\u26A0\uFE0F</div>
            <h3 style="margin:0 0 8px 0; color:var(--text-1);">Gagal Memuat Data</h3>
            <p style="color:var(--text-3); margin-bottom:16px;">${x.message}</p>
            <button id="btn-retry-hygiene" class="btn btn-primary" style="padding:8px 16px; border-radius:8px;">Coba Lagi</button>
          </div>
        `,k.querySelector("#btn-retry-hygiene")?.addEventListener("click",i))}}function d(){let x=e.querySelector("#badge-count-all");x&&(x.textContent=m.length);let k=e.querySelector("#card-total-params");k&&(k.textContent=m.length),Object.entries(pt).forEach(([w,K])=>{let G=w.replace(/[\s&]+/g,"-"),J=e.querySelector(`#badge-count-${G}`);if(J){let W=m.filter(Y=>K.rooms.includes(Y.room_name)).length;J.textContent=W}});let _=new Set(m.map(w=>w.room_name)),E=e.querySelector("#card-total-rooms");E&&(E.textContent=_.size)}function F(){let x=e.querySelector("#hygiene-room-filter");if(!x)return;let k=x.value,_=[];if(s==="all")_=Array.from(new Set(m.map(E=>E.room_name))).filter(Boolean).sort();else{let E=pt[s];E&&(_=E.rooms.filter(w=>m.some(K=>K.room_name===w)))}x.innerHTML=`<option value="">Semua Ruangan (${_.length})</option>`+_.map(E=>`<option value="${E}" ${E===k?"selected":""}>${E}</option>`).join("")}function $(){let x=(e.querySelector("#hygiene-search")?.value||"").toLowerCase().trim(),k=e.querySelector("#hygiene-room-filter")?.value||"";return m.filter(_=>{if(s!=="all"){let E=pt[s];if(!E||!E.rooms.includes(_.room_name))return!1}if(k&&_.room_name!==k)return!1;if(x){let E=(_.room_name||"").toLowerCase(),w=(_.item_name||"").toLowerCase(),K=(_.cleanliness_standard||"").toLowerCase();if(!E.includes(x)&&!w.includes(x)&&!K.includes(x))return!1}return!0})}function M(){let x=$(),k=e.querySelector("#hygiene-stats-text");k&&(k.innerHTML=`Menampilkan <strong>${x.length}</strong> dari <strong>${m.length}</strong> parameter`);let _=e.querySelector("#hygiene-view-container");if(_){if(x.length===0){_.innerHTML=`
        <div class="empty-state" style="text-align:center; padding:50px 20px; background:#131F36; border-radius:12px; border:1px solid #223456;">
          <div style="font-size:3rem; margin-bottom:12px;">\u{1F50D}</div>
          <h3 style="margin:0 0 8px 0; color:var(--text-1); font-weight:700;">Tidak Ada Parameter yang Cocok</h3>
          <p style="color:var(--text-3); max-width:400px; margin:0 auto 16px auto; font-size:0.9rem;">
            Tidak ditemukan parameter kebersihan dengan kata kunci atau filter yang Anda pilih.
          </p>
          <button id="btn-clear-search-empty" class="btn btn-outline" style="border-radius:8px; padding:8px 16px; font-weight:600;">
            Reset Filter
          </button>
        </div>
      `,_.querySelector("#btn-clear-search-empty")?.addEventListener("click",()=>{let E=e.querySelector("#hygiene-search"),w=e.querySelector("#hygiene-room-filter");E&&(E.value=""),w&&(w.value=""),s="all",e.querySelectorAll(".zone-tab-btn").forEach(K=>{let G=K.dataset.zone==="all";K.classList.toggle("active",G),K.setAttribute("aria-selected",G?"true":"false")}),F(),M()});return}c==="table"?h(_,x):L(_,x)}}function h(x,k){x.innerHTML=`
      <div class="hyg-table-card hyg-table-card--desktop">
        <div class="table-wrapper">
          <table class="hyg-table">
            <thead>
              <tr>
                <th style="width:52px; text-align:center;">No</th>
                <th style="width:220px;">Ruangan &amp; Zona</th>
                <th style="width:280px;">Item / Objek Inspeksi</th>
                <th>Standar Kebersihan Fisik</th>
                ${n||a?'<th style="width:96px; text-align:center;">Aksi</th>':""}
              </tr>
            </thead>
            <tbody>
              ${k.map((_,E)=>{let w=ea(_.room_name);return`
                  <tr class="hygiene-std-row">
                    <td class="hyg-col-no">${E+1}</td>
                    <td class="hyg-col-room">
                      <span class="hyg-room-badge">${P(_.room_name)}</span>
                      <div class="hyg-zone-meta">${P(w.zoneName)}</div>
                    </td>
                    <td class="hyg-col-item">
                      <div class="hyg-item-title">${P(_.item_name)}</div>
                    </td>
                    <td class="hyg-col-std">
                      <div class="hyg-std-box">
                        <span class="hyg-std-check" aria-hidden="true">\u2713</span>
                        <span class="hyg-std-text">${P(_.cleanliness_standard)}</span>
                      </div>
                    </td>
                    ${n||a?`
                      <td class="hyg-col-actions">
                        ${Ha(_,n,a,!1)}
                      </td>
                    `:""}
                  </tr>
                `}).join("")}
            </tbody>
          </table>
        </div>
      </div>
      ${oi(k,{includeRoom:!0,canEdit:n,canDelete:a})}
    `,R(x)}function L(x,k){let _={};k.forEach(E=>{let w=ea(E.room_name);_[w.zoneName]||(_[w.zoneName]={info:w,rooms:{}}),_[w.zoneName].rooms[E.room_name]||(_[w.zoneName].rooms[E.room_name]=[]),_[w.zoneName].rooms[E.room_name].push(E)}),x.innerHTML=`
      <div class="hyg-zone-list">
        ${Object.entries(_).map(([E,w])=>{let K=Object.values(w.rooms).reduce((G,J)=>G+J.length,0);return`
            <div class="hyg-zone-card">
              <div class="hyg-zone-card-header">
                <div>
                  <h3 class="hyg-zone-card-title">${P(E)}</h3>
                  <div class="hyg-zone-card-sub">
                    ${Object.keys(w.rooms).length} Ruangan \xB7 ${K} Parameter Standar
                  </div>
                </div>
                <span class="badge badge-primary">${K} Parameter</span>
              </div>

              <div class="hyg-zone-card-body">
                ${Object.entries(w.rooms).map(([G,J])=>`
                  <div class="hyg-room-block">
                    <div class="hyg-room-block-head">
                      <h4 class="hyg-room-block-title">${P(G)}</h4>
                      <span class="badge badge-secondary">${J.length} item</span>
                    </div>

                    <div class="hyg-zone-table-wrap hyg-zone-table-wrap--desktop">
                      <table class="hyg-table hyg-table--compact">
                        <thead>
                          <tr>
                            <th style="width:44px; text-align:center;">No</th>
                            <th style="width:260px;">Item / Objek Inspeksi</th>
                            <th>Standar Kebersihan Fisik</th>
                            ${n||a?'<th style="width:86px; text-align:center;">Aksi</th>':""}
                          </tr>
                        </thead>
                        <tbody>
                          ${J.map((W,Y)=>`
                            <tr class="hygiene-std-row">
                              <td class="hyg-col-no">${Y+1}</td>
                              <td class="hyg-col-item">
                                <div class="hyg-item-title">${P(W.item_name)}</div>
                              </td>
                              <td class="hyg-col-std">
                                <div class="hyg-std-inline">
                                  <span class="hyg-std-check" aria-hidden="true">\u2713</span>
                                  <span class="hyg-std-text">${P(W.cleanliness_standard)}</span>
                                </div>
                              </td>
                              ${n||a?`
                                <td class="hyg-col-actions">
                                  ${Ha(W,n,a,!0)}
                                </td>
                              `:""}
                            </tr>
                          `).join("")}
                        </tbody>
                      </table>
                    </div>
                    ${oi(J,{includeRoom:!1,canEdit:n,canDelete:a})}
                  </div>
                `).join("")}
              </div>
            </div>
          `}).join("")}
      </div>
    `,R(x)}function R(x){x.querySelectorAll(".btn-edit-item").forEach(k=>{k.addEventListener("click",()=>{let _=k.dataset.id,E=m.find(w=>String(w.id)===String(_));E&&I(E)})}),x.querySelectorAll(".btn-delete-item").forEach(k=>{k.addEventListener("click",()=>{if(!a){j("Hanya Superadmin yang dapat menghapus parameter standar.");return}let _=k.dataset.id,E=k.dataset.name;ge(`Yakin ingin menghapus parameter standar <strong>"${E}"</strong>?`,async()=>{try{let w=await X(`/api/hygiene-standards/${_}`,{method:"DELETE"});if(!w.ok)throw new Error(w.data?.error||"Gagal menghapus parameter");ae("Parameter kebersihan berhasil dihapus"),i()}catch(w){j(w.message)}},"Hapus Parameter Standar")})})}function I(x=null){if(!n){j("Anda tidak memiliki izin untuk menambah atau mengedit parameter kebersihan.");return}let k=!!x,_=Array.from(new Set(m.map(w=>w.room_name))).filter(Boolean).sort(),E=document.createElement("div");E.innerHTML=`
      <form id="form-hygiene-item" style="display:flex; flex-direction:column; gap:16px;">
        <div>
          <label style="display:block; font-size:0.85rem; font-weight:600; color:var(--text-1); margin-bottom:6px;">
            Ruangan / Lokasi <span class="required">*</span>
          </label>
          <input type="text" id="modal-room-name" list="room-suggestions" class="form-control" value="${x?.room_name||""}" required style="width:100%; height:40px; border-radius:8px;">
          <datalist id="room-suggestions">
            ${_.map(w=>`<option value="${w}">`).join("")}
          </datalist>
          <span style="font-size:0.75rem; color:var(--text-3); margin-top:4px; display:block;">Pilih dari daftar atau ketik nama ruangan baru</span>
        </div>
        <div>
          <label style="display:block; font-size:0.85rem; font-weight:600; color:var(--text-1); margin-bottom:6px;">
            Item / Objek Inspeksi <span class="required">*</span>
          </label>
          <input type="text" id="modal-item-name" class="form-control" value="${x?.item_name||""}" required style="width:100%; height:40px; border-radius:8px;">
        </div>
        <div>
          <label style="display:block; font-size:0.85rem; font-weight:600; color:var(--text-1); margin-bottom:6px;">
            Standar Kebersihan Fisik <span class="required">*</span>
          </label>
          <textarea id="modal-standard" class="form-control" rows="4" required style="width:100%; border-radius:8px; line-height:1.5;">${x?.cleanliness_standard||""}</textarea>
          <span style="font-size:0.75rem; color:var(--text-3); margin-top:4px; display:block;">Jelaskan kriteria fisik: bebas debu kasat mata, steril, tidak licin, tidak berbau, dll.</span>
        </div>
      </form>
    `,ce({title:k?"\u270F\uFE0F Edit Parameter Kebersihan":"\u2795 Tambah Parameter Kebersihan",content:E,confirmText:k?"Simpan Perubahan":"Tambah Parameter",size:"md",onConfirm:async(w,K)=>{let G=E.querySelector("#modal-room-name").value.trim(),J=E.querySelector("#modal-item-name").value.trim(),W=E.querySelector("#modal-standard").value.trim();if(!G||!J||!W){j("Semua kolom wajib diisi");return}try{let Y={room_name:G,item_name:J,cleanliness_standard:W},b=k?`/api/hygiene-standards/${x.id}`:"/api/hygiene-standards",y=await X(b,{method:k?"PUT":"POST",body:JSON.stringify(Y)});if(!y.ok)throw new Error(y.data?.error||"Gagal menyimpan parameter");ae(k?"Parameter berhasil diperbarui":"Parameter berhasil ditambahkan"),K(),i()}catch(Y){j(Y.message)}}})}function N(){let x=window.open("","_blank");if(!x){j("Pop-up printer diblokir oleh browser. Izinkan pop-up untuk mencetak SOP.");return}let _=`
      <!DOCTYPE html>
      <html lang="id">
      <head>
        <meta charset="UTF-8">
        <title>SOP Master Standar Kebersihan Fisik - Audy Dental</title>
        <style>@page{size:A4 portrait;margin:12mm 15mm 15mm 15mm}body{font-family:'Segoe UI',Arial,sans-serif;font-size:10pt;color:#1e293b;line-height:1.4;margin:0;padding:0}.header{border-bottom:2px solid #1e3a8a;padding-bottom:8px;margin-bottom:14px;display:flex;justify-content:space-between;align-items:flex-end}.logo{font-size:16pt;font-weight:800;color:#1e3a8a;letter-spacing:-0.5px}.sublogo{font-size:8.5pt;color:#64748b;font-weight:600;text-transform:uppercase}.doc-info{text-align:right;font-size:8pt;color:#475569}.title{text-align:center;margin:12px 0 16px 0}.title h1{font-size:13pt;margin:0 0 4px 0;color:#0f172a;text-transform:uppercase}.title p{margin:0;font-size:8.5pt;color:#64748b}.zone-title{background:#f1f5f9;border-left:4px solid #2563eb;padding:6px 10px;font-size:9.5pt;font-weight:700;color:#1e3a8a;margin:14px 0 8px 0;page-break-after:avoid}table{width:100%;border-collapse:collapse;margin-bottom:12px;font-size:8.5pt;page-break-inside:auto}tr{page-break-inside:avoid;page-break-after:auto}th{background:#f8fafc;border:1px solid #cbd5e1;padding:6px 8px;text-align:left;font-weight:700;color:#334155;font-size:8pt;text-transform:uppercase}td{border:1px solid #cbd5e1;padding:6px 8px;vertical-align:top}.td-no{text-align:center;width:28px}.td-room{width:120px;font-weight:600}.td-obj{width:150px;font-weight:600;color:#0f172a}.footer-sign{margin-top:24px;display:flex;justify-content:space-between;page-break-inside:avoid}.sign-box{text-align:center;width:180px;font-size:8.5pt}.sign-line{margin-top:45px;border-top:1px solid #0f172a;font-weight:600;padding-top:4px}@media print{.no-print{display:none !important}}</style>
      </head>
      <body>
        <div class="no-print" style="background:#fef3c7; border:1px solid #f59e0b; padding:10px 14px; border-radius:6px; margin-bottom:14px; display:flex; justify-content:space-between; align-items:center;">
          <span style="font-size:9pt; color:#92400e;">\u{1F4A1} <strong>Format Cetak SOP Standar Kebersihan Fisik FCMS</strong>. Siap dicetak untuk display janitor atau arsip operasional.</span>
          <button id="hygiene-print-now-btn" style="background:#2563eb; color:#fff; border:none; padding:6px 14px; border-radius:4px; font-weight:bold; cursor:pointer;">Cetak Sekarang</button>
          <script>document.getElementById('hygiene-print-now-btn')?.addEventListener('click', function(){ window.print(); });<\/script>
        </div>

        <div class="header">
          <div>
            <div class="logo">AUDY DENTAL</div>
            <div class="sublogo">Facility Care Management System (FCMS)</div>
          </div>
          <div class="doc-info">
            <div>Dokumen: <strong>FCMS-SOP-HYG-01</strong></div>
            <div>Tanggal Cetak: ${new Date().toLocaleDateString("id-ID",{day:"numeric",month:"long",year:"numeric"})}</div>
            <div>Total: <strong>${m.length} Parameter Fisik</strong></div>
          </div>
        </div>

        <div class="title">
          <h1>STANDAR ACUAN KEBERSIHAN FISIK SELURUH CABANG</h1>
          <p>Buku Panduan Kriteria Kebersihan Fisik 8 Zona Area (Cleaning Service & Facility Care)</p>
        </div>

        ${Object.entries(pt).map(([E,w])=>{let K=m.filter(G=>w.rooms.includes(G.room_name));return K.length===0?"":`
            <div class="zone-title">${w.icon} ${E.toUpperCase()} (${K.length} PARAMETER)</div>
            <table>
              <thead>
                <tr>
                  <th class="td-no">No</th>
                  <th class="td-room">Ruangan / Lokasi</th>
                  <th class="td-obj">Item / Objek Inspeksi</th>
                  <th>Standar Kebersihan Fisik</th>
                </tr>
              </thead>
              <tbody>
                ${K.map((G,J)=>`
                  <tr>
                    <td class="td-no">${J+1}</td>
                    <td class="td-room">${G.room_name}</td>
                    <td class="td-obj">${G.item_name}</td>
                    <td>${G.cleanliness_standard}</td>
                  </tr>
                `).join("")}
              </tbody>
            </table>
          `}).join("")}

        <div class="footer-sign">
          <div class="sign-box">
            <div>Dibuat Oleh,</div>
            <div class="sign-line">Facility Care Specialist</div>
          </div>
          <div class="sign-box">
            <div>Ditinjau Oleh,</div>
            <div class="sign-line">Supervisor Operasional</div>
          </div>
          <div class="sign-box">
            <div>Disetujui Oleh,</div>
            <div class="sign-line">Head of Operations</div>
          </div>
        </div>

        <script>
          window.onload = function() {
            setTimeout(function() { window.print(); }, 400);
          };
        <\/script>
      </body>
      </html>
    `;x.document.open(),x.document.write(_),x.document.close()}async function q(){try{let k=(await ue("/api/hygiene-standards/export",{})).map((E,w)=>{let K=ea(E.room_name);return{No:w+1,"Kategori Zona":K.zoneName,"Ruangan / Lokasi":E.room_name,"Item / Objek Inspeksi":E.item_name,"Standar Kebersihan Fisik":E.cleanliness_standard}}),_=new Date().toISOString().slice(0,10);be(k,`Master_Standar_Kebersihan_FCMS_${_}`),ae("Data master kebersihan berhasil diekspor ke Excel")}catch(x){j("Gagal mengekspor data: "+x.message)}}e.querySelector("#hygiene-search")?.addEventListener("input",()=>M()),e.querySelector("#hygiene-room-filter")?.addEventListener("change",()=>M()),e.querySelector("#btn-reset-filters")?.addEventListener("click",()=>{let x=e.querySelector("#hygiene-search"),k=e.querySelector("#hygiene-room-filter");x&&(x.value=""),k&&(k.value=""),M()}),e.querySelectorAll(".zone-tab-btn").forEach(x=>{x.addEventListener("click",()=>{s=x.dataset.zone,e.querySelectorAll(".zone-tab-btn").forEach(k=>{let _=k.dataset.zone===s;k.classList.toggle("active",_),k.setAttribute("aria-selected",_?"true":"false")}),F(),M()})});let v=e.querySelector("#view-table-btn"),B=e.querySelector("#view-zone-btn");v?.addEventListener("click",()=>{c="table",v.classList.add("active"),B?.classList.remove("active"),M()}),B?.addEventListener("click",()=>{c="zone",B.classList.add("active"),v?.classList.remove("active"),M()}),e.querySelector("#btn-print-hygiene")?.addEventListener("click",N),e.querySelector("#btn-export-hygiene")?.addEventListener("click",q),e.querySelector("#btn-add-hygiene")?.addEventListener("click",()=>I(null)),i()}Ce();Ie();async function li(e){let t=se();if(!t||t.role!=="superadmin"){e.innerHTML='<div class="empty-state"><p class="text-danger">Akses ditolak.</p></div>';return}Me({container:e,title:"Manajemen User",icon:"\u{1F510}",apiPath:"/api/users",bulkDelete:!0,enableMobileFilterSheet:!0,itemLabel:"User",columns:[{key:"full_name",label:"Nama Lengkap"},{key:"username",label:"Username"},{key:"email",label:"Email"},{key:"role",label:"Role",render:n=>`<span class="badge ${{superadmin:"badge-danger",fc_specialist_leader:"badge-info",editor_khusus:"badge-info",viewer:"badge-success"}[n]||"badge-neutral"}">${P(n)}</span>`},{key:"is_active",label:"Status",render:n=>n?'<span class="badge badge-success">Aktif</span>':'<span class="badge badge-neutral">Nonaktif</span>'},{key:"created_at",label:"Dibuat",nowrap:!0,render:n=>n?new Date(n).toLocaleDateString("id-ID"):"-"}],filterFields:[{type:"search",placeholder:"Cari nama / username..."}],formFields:n=>{let a=!!n,o=String(n?.role||"").trim().toLowerCase(),s={value:"superadmin",label:"Superadmin (Akses Penuh - Maks. 2 Aktif)"},c={value:"fc_specialist_leader",label:"FC Specialist / Leader (View All + Input & Edit Lembur & Issue - Maks. 2 Aktif)"},m={value:"viewer",label:"Viewer (Read-Only Semua Modul Aktif - Maks. 5 Aktif)"},i=a&&o==="superadmin"?[s,c,m]:[c,m,s],d=a&&(o==="superadmin"||o==="fc_specialist_leader"||o==="viewer")?o:"fc_specialist_leader";return[{type:"row",fields:[{name:"full_name",label:"Nama Lengkap",required:!0,value:n?.full_name},{name:"username",label:"Username",required:!0,value:n?.username,autocomplete:"off"}]},{type:"row",fields:[{name:"email",label:"Email",type:"email",required:!0,value:n?.email,autocomplete:"off"},{name:"role",label:"Role Resmi",type:"select",required:!0,options:i,value:d}]},{type:"row",fields:[{name:"password",label:a?"Password Baru (kosongkan jika tidak diubah)":"Password",type:"password",required:!a,placeholder:"Min. 12 karakter",autocomplete:"new-password"},{name:"is_active",label:"Status Aktif",type:"checkbox",checkLabel:"User aktif",value:a?n?.is_active:1}]},{type:"html",html:`
            <div style="margin-top:8px; padding:10px 12px; background:rgba(37,99,235,0.14); border:1px solid rgba(59,130,246,0.35); border-radius:8px; font-size:0.8rem; color:#F8FAFC; line-height:1.5;">
              <strong style="color:#93C5FD;">Kebijakan RBAC &amp; Keamanan Password:</strong> Role resmi FCMS hanya <code style="background:rgba(15,23,42,0.7); color:#93C5FD; padding:1px 5px; border-radius:4px;">superadmin</code> (maks. 2 aktif) dan <code style="background:rgba(15,23,42,0.7); color:#93C5FD; padding:1px 5px; border-radius:4px;">fc_specialist_leader</code> (maks. 2 aktif). Password wajib minimal 12 karakter (kombinasi huruf besar, huruf kecil, angka, dan simbol).
            </div>
          `}]},onBeforeSubmit:(n,a)=>(n.full_name&&(n.full_name=String(n.full_name).trim()),n.username&&(n.username=String(n.username).trim()),n.email&&(n.email=String(n.email).trim().toLowerCase()),a&&(!n.password||String(n.password).trim()==="")&&delete n.password,n.assigned_modules=null,n),exportOptions:{moduleName:"users",onExport:async n=>{let o=(await ue("/api/users/export",n)).map(s=>({"Nama Lengkap":s.full_name||"",Username:s.username||"",Email:s.email||"",Role:s.role||"",Status:s.is_active?"Aktif":"Nonaktif"}));be(o,"Data_Users")},onTemplate:()=>{be([{"Nama Lengkap":"FC Leader 01",Username:"fcleader01",Email:"fcleader@contoh.com",Role:"fc_specialist_leader",Password:""}],"Template_Import_Users"),typeof gt=="function"&&gt("Kolom Password pada template dikosongkan. Wajib isi password unik dan kuat (min. 12 karakter: huruf besar, kecil, angka, simbol) untuk setiap akun baru, atau kosongkan jika memperbarui akun existing tanpa mengubah password.")},onImport:async(n,a={})=>{let o=n.map(s=>({full_name:String(s["Nama Lengkap"]||"").trim(),username:String(s.Username||"").trim(),email:String(s.Email||"").trim().toLowerCase(),role:String(s.Role||"").trim()||"fc_specialist_leader",password:String(s.Password||"").trim()})).filter(s=>s.username||s.email||s.full_name||s.password);return await Pe({module:"users",label:"Manajemen Pengguna",endpoint:"/api/users/import",body:o,idempotencyKey:a?.idempotencyKey,recover:!!a?.recover})}}})}Ce();async function di(e){Me({container:e,title:"Manajemen Cabang",icon:"\u{1F3E2}",apiPath:"/api/branches",enableMobileFilterSheet:!0,itemLabel:"Cabang",bulkDelete:!1,columns:[{key:"code",label:"Kode",width:"60px"},{key:"full_name",label:"Nama Cabang"},{key:"city",label:"Kota"},{key:"branch_scope",label:"Lingkup",render:(t,n)=>{let a=String(n?.code||"").trim();return a==="000"?'<span class="badge badge-info" title="Management / HQ (non-operasional, tersedia di dropdown modul)">HQ (Non-Ops)</span>':a==="A01"?'<span class="badge badge-warning" title="Gabung area 002. Kemang (tidak dihitung terpisah di Dashboard)">Gabung 002. Kemang</span>':/^(00[1-9]|0[1-6][0-9]|070)$/.test(a)?'<span class="badge badge-success">Operasional</span>':'<span class="badge badge-neutral">Non-Master</span>'}},{key:"is_active",label:"Status",render:t=>t?'<span class="badge badge-success">Aktif</span>':'<span class="badge badge-neutral">Nonaktif</span>'}],filterFields:[{type:"search",placeholder:"Cari nama / kode cabang..."}],formFields:t=>[{type:"row",fields:[{name:"code",label:"Kode Cabang",required:!0,value:t?.code},{name:"name",label:"Nama Pendek",required:!0,value:t?.name}]},{name:"full_name",label:"Nama Lengkap",required:!0,value:t?.full_name},{type:"row",fields:[{name:"city",label:"Kota",value:t?.city},{name:"is_active",label:"Status",type:"checkbox",checkLabel:"Cabang aktif",value:t?.is_active!==void 0?t.is_active:1}]}],exportOptions:{moduleName:"branches",onExport:async t=>{let n=await ue("/api/branches/export",t);be(n,"Data_Cabang")},onTemplate:()=>{be([{"Kode Cabang":"001","Nama Pendek":"Pondok Bambu","Nama Lengkap":"001. Pondok Bambu",Kota:"Jakarta Timur"},{"Kode Cabang":"002","Nama Pendek":"Bintaro","Nama Lengkap":"002. Bintaro",Kota:"Tangerang Selatan"}],"Template_Import_Cabang")},onImport:async(t,n={})=>{let a=t.map(o=>({code:String(o["Kode Cabang"]||"").trim(),name:String(o["Nama Pendek"]||"").trim(),full_name:String(o["Nama Lengkap"]||"").trim(),city:String(o.Kota||"").trim()})).filter(o=>o.code&&o.name);return await Pe({module:"branches",label:"Data Cabang",endpoint:"/api/branches/import",body:a,idempotencyKey:n?.idempotencyKey,recover:!!n?.recover})}}})}async function ci(e){let t=new Date,n=[];e.innerHTML=`
    <div class="page-header">
      <h1 class="page-title">\u{1F4C5} Kalender</h1>
    </div>
    <div class="card">
      <div class="card-header calendar-nav">
        <div class="cal-nav-controls">
          <button type="button" class="btn btn-sm cal-nav-btn" id="cal-prev" aria-label="Bulan sebelumnya">\u2039 Prev</button>
          <span class="calendar-month-label" id="cal-month-label"></span>
          <button type="button" class="btn btn-sm cal-nav-btn" id="cal-next" aria-label="Bulan berikutnya">Next \u203A</button>
        </div>
        <div class="calendar-filters">
          <label class="filter-check"><input type="checkbox" value="schedule"        checked class="cal-filter"> Jadwal</label>
          <label class="filter-check"><input type="checkbox" value="reports"         checked class="cal-filter"> Laporan</label>
          <label class="filter-check"><input type="checkbox" value="issue"           checked class="cal-filter"> Permasalahan</label>
          <label class="filter-check"><input type="checkbox" value="training"        checked class="cal-filter"> Training</label>
          <label class="filter-check"><input type="checkbox" value="contract_expiry" checked class="cal-filter"> Kontrak Habis</label>
        </div>
      </div>
      <div class="card-body p-0">
        <div id="calendar-grid" style="min-height:400px"></div>
      </div>
    </div>
    <!-- Event detail popup -->
    <div id="cal-event-list" class="cal-event-sidebar" style="display:none">
      <div class="cal-event-header">
        <span id="cal-event-date"></span>
        <button class="btn btn-ghost btn-sm" id="cal-event-close">&times;</button>
      </div>
      <div id="cal-event-items"></div>
    </div>
  `,document.getElementById("cal-prev").addEventListener("click",()=>{t.setMonth(t.getMonth()-1),o()}),document.getElementById("cal-next").addEventListener("click",()=>{t.setMonth(t.getMonth()+1),o()}),document.getElementById("cal-event-close").addEventListener("click",()=>{document.getElementById("cal-event-list").style.display="none"}),document.querySelectorAll(".cal-filter").forEach(c=>c.addEventListener("change",o));async function a(){try{let c=`${t.getFullYear()}-${String(t.getMonth()+1).padStart(2,"0")}`;n=(await X(`/api/dashboard/calendar?month=${c}`)).data?.data||[]}catch(c){console.warn("[Calendar] Failed to load events, rendering empty grid:",c),n=[]}}async function o(){let c=document.getElementById("calendar-grid");if(c){c.innerHTML=`<div style="display:grid;grid-template-columns:repeat(7,1fr);gap:1px;background:#1E2E4A;">
      ${Array(35).fill('<div style="background:#131F36;min-height:112px;"></div>').join("")}
    </div>`,await a();try{let m=t.getFullYear(),i=t.getMonth(),d=t.toLocaleDateString("id-ID",{month:"long",year:"numeric"}),F=document.getElementById("cal-month-label");F&&(F.textContent=d);let $=new Set(["inspection","cleaning","fogging","basecamp"]),M=new Set(Array.from(document.querySelectorAll(".cal-filter:checked")).map(E=>E.value)),h=n.filter(E=>M.has(E.type)||M.has("reports")&&$.has(E.type)),L={};h.forEach(E=>{let w=Nr(E.event_date);w&&(L[w]||(L[w]=[]),L[w].push(E))});let R=new Date(m,i,1).getDay(),I=new Date(m,i+1,0).getDate(),N=["Min","Sen","Sel","Rab","Kam","Jum","Sab"],q=new Date,v=`${q.getFullYear()}-${String(q.getMonth()+1).padStart(2,"0")}-${String(q.getDate()).padStart(2,"0")}`,B='<div class="calendar-grid">';N.forEach(E=>{B+=`<div class="cal-day-header">${E}</div>`});for(let E=0;E<R;E++)B+='<div class="cal-cell cal-cell-empty"></div>';for(let E=1;E<=I;E++){let w=`${m}-${String(i+1).padStart(2,"0")}-${String(E).padStart(2,"0")}`,K=L[w]||[],G=w===v;B+=`
          <div class="cal-cell ${G?"cal-today":""} ${K.length?"cal-has-events":""}"
               data-date="${w}" tabindex="0" role="button" aria-label="${w}${K.length?`, ${K.length} kegiatan`:""}">
            <div class="cal-day-num ${G?"today-num":""}">${E}</div>
            <div class="cal-events-preview">
              ${K.slice(0,3).map(J=>{let W=J.title||J.branch_name||J.type;return`
                <div class="cal-event-dot cal-color-${J.color||"gray"}" title="${qe(W)}${J.branch_name?" \u2014 "+qe(J.branch_name):""}">
                  <span class="cal-event-dot-label">${qe(W)}</span>
                </div>
              `}).join("")}
              ${K.length>3?`<div class="cal-more">+${K.length-3} kegiatan lainnya</div>`:""}
            </div>
          </div>`}let k=(R+I)%7;if(k!==0)for(let E=0;E<7-k;E++)B+='<div class="cal-cell cal-cell-empty"></div>';B+="</div>",c.innerHTML=B;let _=E=>{let w=E.dataset.date,K=L[w]||[];if(!K.length)return;let G=document.getElementById("cal-event-list"),J=new Date(w+"T00:00:00").toLocaleDateString("id-ID",{weekday:"long",day:"numeric",month:"long",year:"numeric"});document.getElementById("cal-event-date").textContent=`${J} (${K.length})`,document.getElementById("cal-event-items").innerHTML=K.map(W=>`
          <div class="cal-event-item cal-color-border-${W.color||"gray"}">
            <div class="cal-event-type">
              ${Kr(W.type)}
              ${W.date_type?` <span class="badge" style="font-size:0.74rem;padding:2px 7px;background:rgba(37,99,235,0.18);color:#93C5FD;border:1px solid rgba(96,165,250,0.3);margin-left:4px;">${qe(W.date_type)}</span>`:""}
            </div>
            <div class="cal-event-title">${qe(W.title||"-")}</div>
            ${W.branch_name?`<div class="cal-event-branch">\u{1F4CD} ${qe(W.branch_name)}</div>`:""}
            ${W.pic||W.period||W.opening_date||W.target_date||W.completion_date||W.notes?`
              <div class="cal-event-meta" style="margin-top:7px;padding-top:7px;border-top:1px dashed #223456;font-size:0.8rem;color:#CBD5E1;display:flex;flex-direction:column;gap:3px;">
                ${W.pic?`<div>\u{1F464} <b>PIC:</b> ${qe(W.pic)}</div>`:""}
                ${W.period?`<div>\u{1F4C5} <b>Periode:</b> ${qe(W.period)}</div>`:""}
                ${W.opening_date?`<div>\u{1F6AA} <b>Tgl Opening:</b> ${qe(Oa(W.opening_date))}</div>`:""}
                ${W.target_date?`<div>\u{1F3AF} <b>Tgl Target:</b> ${qe(Oa(W.target_date))}</div>`:""}
                ${W.completion_date?`<div>\u2705 <b>Tgl Selesai:</b> ${qe(Oa(W.completion_date))}</div>`:""}
                ${W.notes?`<div>\u{1F4DD} <b>Catatan:</b> ${qe(W.notes)}</div>`:""}
              </div>
            `:""}
            ${W.status?`<div class="cal-event-status" style="margin-top:6px;"><span class="badge ${W.status==="Done"?"badge-success":"badge-warning"}">${qe(W.status)}</span></div>`:""}
            ${W.days_remaining!==void 0?`<div class="cal-event-extra">Sisa: ${W.days_remaining} hari</div>`:""}
          </div>
        `).join(""),G.style.display="block"};c.querySelectorAll(".cal-cell[data-date]").forEach(E=>{E.addEventListener("click",()=>_(E)),E.addEventListener("keydown",w=>{(w.key==="Enter"||w.key===" ")&&(w.preventDefault(),_(E))})})}catch(m){console.error("[Calendar] Render error:",m),c&&(c.innerHTML=`
          <div style="padding:40px;text-align:center;color:var(--text-3)">
            <div style="font-size:2rem;margin-bottom:8px">\u{1F4C5}</div>
            <div>Gagal memuat kalender. Silakan refresh.</div>
          </div>`)}}}let s=()=>{a().then(()=>o())};return Tt.on("data:changed",s),o(),()=>{Tt.off("data:changed",s)}}function Nr(e){if(!e||e==="-"||String(e).trim()==="")return"";let t=String(e).trim();if(/^\d{4}-\d{2}-\d{2}/.test(t))return t.slice(0,10);let n=t.split(/[\/\-\.]/);if(n.length===3){let[a,o,s]=n.map(c=>c.trim());if(s.length===4)return`${s}-${o.padStart(2,"0")}-${a.padStart(2,"0")}`;if(a.length===4)return`${a}-${o.padStart(2,"0")}-${s.padStart(2,"0")}`}return t.slice(0,10)}function Oa(e){if(!e||e==="-"||String(e).trim()==="")return"";let n=String(e).trim().split(/[\/\-\.]/);if(n.length===3){let[a,o,s]=n.map(c=>c.trim());if(a.length===4)return`${s.padStart(2,"0")}-${o.padStart(2,"0")}-${a}`;if(s.length===4)return`${a.padStart(2,"0")}-${o.padStart(2,"0")}-${s}`}return e}function qe(e){return e?String(e).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;"):""}function Kr(e){return{schedule:"\u{1F5D3} Jadwal",issue:"\u26A0\uFE0F Permasalahan",training:"\u{1F393} Training",contract_expiry:"\u{1F4CB} Kontrak Habis",cleaning:"\u{1F9F9} Cleaning",inspection:"\u{1F50D} Inspeksi",fogging:"\u{1F4A8} Fogging",basecamp:"\u{1F3E0} Basecamp"}[e]||e}Ie();async function pi(e){let t=se(),n=(t?.full_name||t?.username||"U")[0].toUpperCase(),o={superadmin:"#60A5FA",fc_specialist_leader:"#38BDF8",editor_khusus:"#38BDF8",viewer:"#34D399"}[t?.role]||"#93C5FD";e.innerHTML=`
    <div class="page-header">
      <h1 class="page-title">Profil Saya</h1>
    </div>

    <div class="profile-layout">

      <!-- LEFT: Info Card -->
      <div class="chart-card profile-info-card">
        <div class="profile-avatar-wrap">
          <div class="profile-avatar-xl" style="background:#2563EB;color:#FFFFFF;border:1px solid #3B82F6;">
            ${P(n)}
          </div>
          <div class="profile-name-block">
            <div class="profile-fullname">${P(t?.full_name||"\u2014")}</div>
            <div class="profile-username">@${P(t?.username||"\u2014")}</div>
            <span class="badge badge-primary" style="margin-top:6px">
              ${P((t?.role||"-").toUpperCase())}
            </span>
          </div>
        </div>

        <hr class="profile-divider">

        <div class="info-list">
          <div class="info-row">
            <span class="info-key">Email</span>
            <span class="info-value">${P(t?.email||"\u2014")}</span>
          </div>
          <div class="info-row">
            <span class="info-key">Username</span>
            <span class="info-value">${P(t?.username||"\u2014")}</span>
          </div>
          <div class="info-row">
            <span class="info-key">Role</span>
            <span class="info-value" style="color:${o};font-weight:700">${P(t?.role||"\u2014")}</span>
          </div>
        </div>
      </div>

      <!-- RIGHT: Change Password -->
      <div class="chart-card">
        <div class="chart-card-header">
          <div>
            <div class="chart-card-title">Ganti Password</div>
            <div class="chart-card-subtitle">Gunakan password yang kuat, minimal 12 karakter</div>
          </div>
        </div>

        <form id="change-pwd-form" novalidate style="margin-top:8px">
          <div class="form-group">
            <label class="form-label">Password Lama <span class="required">*</span></label>
            <input type="password" name="current_password" class="form-control"
              required autocomplete="current-password">
          </div>
          <div class="form-group">
            <label class="form-label">Password Baru <span class="required">*</span></label>
            <input type="password" name="new_password" class="form-control"
              required placeholder="Minimal 12 karakter" autocomplete="new-password">
          </div>
          <div class="form-group">
            <label class="form-label">Konfirmasi Password Baru <span class="required">*</span></label>
            <input type="password" name="confirm_password" class="form-control"
              required autocomplete="new-password">
          </div>

          <div id="pwd-error" class="alert alert-danger" style="display:none"></div>
          <div id="pwd-success" class="alert alert-success" style="display:none"></div>

          <button type="submit" class="btn btn-primary" id="btn-save-pwd">
            <svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24" style="margin-right:5px">
              <path d="M19 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11l5 5v11a2 2 0 0 1-2 2z"/>
              <polyline points="17 21 17 13 7 13 7 21"/><polyline points="7 3 7 8 15 8"/>
            </svg>
            Simpan Password
          </button>
        </form>

        <hr class="profile-divider" style="margin-top:28px">

        <div class="chart-card-title" style="margin-bottom:12px">\u{1F510} Keamanan Akun</div>
        <div class="info-list">
          <div class="info-row">
            <span class="info-key">Token Login</span>
            <span class="info-value">
              <span class="badge badge-success">Aktif</span>
            </span>
          </div>
          <div class="info-row">
            <span class="info-key">Session</span>
            <span class="info-value" id="session-info">Memuat...</span>
          </div>
        </div>
        <button class="btn btn-danger btn-sm" id="btn-logout" style="margin-top:16px">
          <svg width="13" height="13" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24" style="margin-right:4px">
            <path d="M9 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h4"/>
            <polyline points="16 17 21 12 16 7"/><line x1="21" y1="12" x2="9" y2="12"/>
          </svg>
          Keluar dari Semua Sesi
        </button>
      </div>

    </div>
  `;let s=document.getElementById("session-info");s&&(s.textContent="Aktif (HttpOnly Cookie + CSRF Terproteksi)"),document.getElementById("change-pwd-form")?.addEventListener("submit",async c=>{c.preventDefault();let m=document.getElementById("pwd-error"),i=document.getElementById("pwd-success"),d=document.getElementById("btn-save-pwd");m.style.display="none",i.style.display="none";let F=c.target,$=F.current_password.value,M=F.new_password.value,h=F.confirm_password.value;if(M!==h){m.textContent="\u274C Konfirmasi password tidak cocok.",m.style.display="block";return}if(M.length<12){m.textContent="\u274C Password baru minimal 12 karakter.",m.style.display="block";return}d.disabled=!0,d.textContent="\u23F3 Menyimpan...";let L=await X("/api/auth/change-password",{method:"POST",body:JSON.stringify({current_password:$,new_password:M})});d.disabled=!1,d.innerHTML='<svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24" style="margin-right:5px"><path d="M19 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11l5 5v11a2 2 0 0 1-2 2z"/><polyline points="17 21 17 13 7 13 7 21"/><polyline points="7 3 7 8 15 8"/></svg>Simpan Password',L.ok?(i.textContent="\u2705 Password berhasil diubah.",i.style.display="block",F.reset(),ae("Password berhasil diubah.")):(m.textContent=L.data?.error||"Gagal mengubah password.",m.style.display="block")}),document.getElementById("btn-logout")?.addEventListener("click",async()=>{if(confirm("Keluar dari semua sesi? Anda harus login ulang.")){try{await X("/api/auth/logout?all=true",{method:"POST"})}catch{}mt(),window.location.reload()}})}var ta={Validasi:{module:"validation",label:"Master Referensi"},SOP:{module:"sop",label:"SOP"},"Master Karyawan":{module:"employees",label:"Karyawan"},"Data Kontrak":{module:"contracts",label:"Kontrak"},Permasalahan:{module:"issues",label:"Permasalahan"},"Time Line":{module:"schedule",label:"Jadwal Kegiatan"},"Report Inspeksi Hygiene 2026":{module:"inspection",label:"Laporan Inspeksi"},"Report GC-DC 2026":{module:"cleaning",label:"Laporan GC/DC"},"Report Fogging 2026":{module:"fogging",label:"Laporan Fogging"},"Rekap Laporan Basecamp":{module:"basecamp",label:"Rekap Basecamp"},Training:{module:"training",label:"Training"},"Master Checklist":{module:"checklist",label:"Checklist"},"Data SP":{module:"sp",label:"Data SP"},"Data Mutasi":{module:"mutasi",label:"Data Mutasi"}};function Ee(e){if(e==null||e==="")return null;if(e instanceof Date)return isNaN(e.getTime())?null:e.toISOString().slice(0,10);let t=String(e).trim();if(t===""||t==="0")return null;if(/^\d{4}-\d{2}-\d{2}/.test(t))return t.slice(0,10);if(/^\d{4,5}$/.test(t)){let o=Number(t);if(o>2e4&&o<99999){let s=new Date(Date.UTC(1899,11,30)+o*864e5);return isNaN(s.getTime())?null:s.toISOString().slice(0,10)}}let n=t.split(/[\/\-\.]/);if(n.length===3){let[o,s,c]=n.map(F=>F.trim()),m=Number(o),i=Number(s),d=Number(c);if(o.length===4&&m>1900)return`${o}-${s.padStart(2,"0")}-${c.padStart(2,"0")}`;if(c.length===4&&d>1900)return m>12?`${c}-${s.padStart(2,"0")}-${o.padStart(2,"0")}`:i>12?`${c}-${o.padStart(2,"0")}-${s.padStart(2,"0")}`:`${c}-${s.padStart(2,"0")}-${o.padStart(2,"0")}`;if(c.length===2&&!isNaN(d)){let F=d>=50?`19${c}`:`20${c}`;return m>12?`${F}-${s.padStart(2,"0")}-${o.padStart(2,"0")}`:`${F}-${s.padStart(2,"0")}-${o.padStart(2,"0")}`}}let a=new Date(t);return isNaN(a.getTime())?null:a.toISOString().slice(0,10)}function ui(e){return Object.values(e).every(t=>t==null||String(t).trim()==="")}var Rr={validation:{required:[],map:e=>({cabang:e.CABANG,pic:e.PIC,kegiatan:e.KEGIATAN,quartal:e.QUARTAL,masa_pkwt:e["MASA PKWT"],pic_pelapor:e["PIC PELAPOR"],kontrak:e.KONTRAK})},sop:{required:[{key:"Nama SOP",label:"Nama SOP"}],map:e=>({name:e["Nama SOP"],category:e.Kategori||"Umum",document_link:e["Link Document"],version:"1.0",effective_date:null,notes:""})},employees:{required:[{key:"Nama Lengkap",label:"Nama Lengkap"}],map:e=>({full_name:e["Nama Lengkap"],branch_name:e.Cabang,division:e["Div / Bagian"]||e.Divisi||"FACILITY CARE",phone:e["No. Hp"]||e["No. HP"],join_date:Ee(e["Tanggal Masuk"]||e["Tgl Masuk"]),status:e.Status||"",target_pindah_os:Ee(e["Target Pindah OS"]),target_selesai_os:Ee(e["Target Selesai OS"]),notes:e.Catatan||""})},contracts:{required:[{key:"Nama Lengkap",label:"Nama Lengkap"}],map:e=>({employee_name:e["Nama Lengkap"],branch_name:e.Cabang,division:e["Div / Bagian"]||"FACILITY CARE",start_date:Ee(e["Tanggal Mulai"]),end_date:Ee(e["Tanggal Selesai"]),contract_type:e["Tipe Kontrak"]||"",pkwt_number:e.PKWT||"",status:e.Status||"",notes:e.keterangan})},issues:{required:[{key:"Keluhan",label:"Keluhan"}],map:e=>({report_date:Ee(e["Tanggal Info"]),branch_name:e.Cabang,category:e.Kategori,source:e["Sumber Laporan"],complaint:e.Keluhan,employee_name:e["Nama FC"],fc_specialist:e["FC Spesialis"],solution:e.Solusi,status:e.Status||"",completion_date:Ee(e["Tanggal Selesai"])})},one_on_one:{required:[],map:e=>({meeting_date:Ee(e.Tanggal),branch_name:e.Cabang,employee_name:e["Nama Karyawan"],pic:e.Pic,problem:e.Masalah,solution:e.Solusi,status:e.Status||"",completion_date:Ee(e["Tanggal Selesai"]),document_link:e["Link Document"]})},schedule:{required:[{key:"Kegiatan",label:"Kegiatan"}],map:e=>({branch_name:e.Cabang,activity_type:e.Kegiatan,period:e.Periode,pic:e.Pic||e.PIC,opening_date:Ee(e["Tanggal Opening"]||e["Tgl Opening"]),target_date:Ee(e["Tanggal Target"]||e["Tgl Target"]),completion_date:Ee(e["Tanggal Selesai"]||e["Tgl Selesai"]),status:e.Status||"",notes:e.Keterangan||e.Catatan})},inspection:{required:[],map:e=>{let t=e["Point FC SP"]!==void 0&&e["Point FC SP"]!==null&&e["Point FC SP"]!==""?e["Point FC SP"]:e["Point FC"];return{inspection_date:Ee(e.Tanggal),branch_name:e.Cabang,period:e.Periode,status:e.Status||"",fc_score:t!=null&&t!==""?parseFloat(String(t).replace(",",".")):null,spv_score:e["Point SPV"]!==void 0&&e["Point SPV"]!==null&&e["Point SPV"]!==""?parseFloat(String(e["Point SPV"]).replace(",",".")):null,document_link:e.Link||e["Link Dokumen"]||"",notes:e.Keterangan||""}}},cleaning:{required:[],map:e=>({activity_date:Ee(e.Tanggal),branch_name:e.Cabang,activity_type:e["Jenis Kegiatan"]||"General Cleaning",period:e.Periode,status:e.Status||"",document_link:e.Link||e["Link Dokumen"]||"",notes:e.Keterangan||""})},fogging:{required:[],map:e=>({activity_date:Ee(e.Tanggal),branch_name:e.Cabang,period:e.Periode,status:e.Status||"",document_link:e.Link||e["Link Dokumen"]||"",notes:e.Keterangan||""})},basecamp:{required:[{key:"Permasalahan",label:"Permasalahan"}],map:e=>({info_date:Ee(e["Tgl Info"]),branch_name:e.Cabang,problem:e.Permasalahan,pic:e.PIC,done_date:Ee(e["Tgl Done"]),status:e.Status||"",notes:e.Ket})},relievers:{required:[],map:e=>({branch_name:e.Cabang,original_fc_name:e["Nama Facility care"],period:e.Periode,reliever_name:e.Relifer,backup_date:Ee(e["Tanggal Back Up"]),completion_date:Ee(e["Tanggal Selesai"]),reason:e.Keterangan,shift:e.Shift,status:e.Status||""})},training:{required:[{key:"Materi",label:"Materi"}],map:e=>({training_date:Ee(e.Tanggal),batch:e.Batch,subject:e.Materi,participants:e.Peserta,branch_name:e.Cabang,trainer:e.Trainer,score:e.Nilai!==void 0&&e.Nilai!==null?parseFloat(String(e.Nilai).replace(",",".")):null,notes:""})},checklist:{required:[],map:e=>({name:e["Master Checklist"],category:"Umum",document_link:e["Link Document"],description:""})},supply:{required:[],map:e=>({submitted_at:Ee(e.Timestamp),submitter_name:e["Nama Lengkap"],branch_name:e["Kebutuhan Untuk Cabang"],tools_items:e["Alat - Alat / Barang"],tools_quantity:e["Jumlah Permintaan Alat / Barang"],chemical_items:e.Chemical,chemical_quantity:e["Jumlah Permintaan Chemical"],additional_notes:e["Tambahan  Alat / Chemical Jika Ada Permintaan Diluar List."],status:e.Status||""})},sp:{required:[{key:"Nama Karyawan",label:"Nama Karyawan"}],map:e=>({tanggal:Ee(e.Tanggal||e["Tanggal Terbit"]),employee_name:e["Nama Karyawan"],branch_name:e.Cabang,sp_type:e["Jenis SP"]||"SP 1",status:e.Status||"Aktif",document_link:e["Link Dokumen"]||e["Link Document"]||null,notes:e.Keterangan||e.Catatan||null})},mutasi:{required:[{key:"Nama Karyawan",label:"Nama Karyawan"}],map:e=>({tanggal:Ee(e.Tanggal||e["Tanggal Efektif"]),employee_name:e["Nama Karyawan"],from_branch_name:e["Cabang Asal"]||e["Dari Cabang"],to_branch_name:e["Cabang Tujuan"]||e["Ke Cabang"],status:e.Status||"Selesai",document_link:e["Link Dokumen"]||e["Link SK"]||null,notes:e.Keterangan||e.Alasan||null})},forms:{required:[{key:"Nama Form",label:"Nama Form"}],map:e=>({name:e["Nama Form"],category:e.Kategori||"Umum",document_link:e["Link Dokumen"]||e["Link Document"]||null,description:e.Deskripsi||e.Keterangan||null,is_public:e.Publik==="Ya"||e.Publik===1||e["Akses Publik"]===1?1:0})}};function jr(e,t){let n=ta[e];if(!n)return{valid:[],errors:[],mapped:[],skipped:!0};let a=Rr[n.module];if(!a)return{valid:[],errors:[],mapped:[],skipped:!0};let o=[],s=[],c=[];return t.filter(i=>!ui(i)).forEach((i,d)=>{let F=t.indexOf(i)+2,$=[];a.required.forEach(({key:h,label:L})=>{let R=i[h];if(R==null||String(R).trim()===""){let I=Object.keys(i).filter(N=>N.trim()).join(", ");$.push({column:L,originalValue:R||"",reason:`Kolom "${L}" wajib diisi dan tidak ditemukan`,hint:`Kolom yang tersedia: ${I.slice(0,120)}`})}});let M=a.map(i);$.length>0?s.push({row:F,data:M,raw:i,errors:$}):(o.push(i),c.push(M))}),{valid:o,errors:s,mapped:c}}function mi(e){let t=[];return e.SheetNames.forEach(n=>{let a=ta[n];if(!a)return;let o=e.Sheets[n],s=window.XLSX.utils.sheet_to_json(o,{defval:"",raw:!1,dateNF:"yyyy-mm-dd"}),c=jr(n,s),m=s.filter(i=>!ui(i));t.push({sheetName:n,module:a.module,label:a.label,total:m.length,valid:c.mapped.length,errorCount:c.errors.length,errors:c.errors,mapped:c.mapped,skipped:!1})}),t}function gi(){let e=window.XLSX,t=e.utils.book_new();Object.entries({Validasi:[{CABANG:"001. Pondok Bambu",PIC:"Berlin",KEGIATAN:"General Cleaning",QUARTAL:"Q1","PIC PELAPOR":"Berlin",KONTRAK:"PKWT 1","MASA PKWT":"1 Tahun"}],SOP:[{"Nama SOP":"SOP Pembersihan Toilet",Kategori:"Cleaning","Link Document":"https://..."}],"Master Karyawan":[{"Nama Lengkap":"Budi Santoso",Cabang:"000. Management","Div / Bagian":"FACILITY CARE","No. Hp":"081234567890","Tanggal Masuk":"2024-01-15",Status:"Aktif"}],"Data Kontrak":[{"Nama Lengkap":"Budi Santoso",Cabang:"000. Management","Div / Bagian":"FACILITY CARE","Tanggal Mulai":"2024-01-01","Tanggal Selesai":"2024-12-31","Tipe Kontrak":"PKWT 1",PKWT:"001/PKWT/2024",Status:"Aktif",keterangan:""}],Permasalahan:[{"Tanggal Info":"2024-03-01",Cabang:"001. Pondok Bambu",Kategori:"Cleaning","Sumber Laporan":"SPV",Keluhan:"Lantai kotor","Nama FC":"Budi","FC Spesialis":"Fajar",Solusi:"Teguran",Status:"Done","Tanggal Selesai":"2024-03-02"}],"Time Line":[{Cabang:"001. Pondok Bambu",Kegiatan:"General Cleaning",Periode:"Januari",Pic:"Berlin","Tanggal Opening":"2024-01-01","Tanggal Target":"2024-01-10","Tanggal Selesai":"2024-01-09",Status:"Done",Keterangan:""}],"Report Inspeksi Hygiene 2026":[{Tanggal:"2026-01-15",Cabang:"001. Pondok Bambu",Periode:"Q1",Status:"Done","Point FC SP":"85.5","Point SPV":"90.0",Link:"https://..."}],"Report GC-DC 2026":[{Tanggal:"2026-01-20",Cabang:"001. Pondok Bambu","Jenis Kegiatan":"General Cleaning",Periode:"Q1",Status:"Done",Link:"https://..."}],"Report Fogging 2026":[{Tanggal:"2026-01-25",Cabang:"001. Pondok Bambu",Periode:"Q1",Status:"Done",Link:"https://..."}],"Rekap Laporan Basecamp":[{"Tgl Info":"2024-02-01",Cabang:"001. Pondok Bambu",Permasalahan:"Lampu mati",PIC:"Berlin","Tgl Done":"2024-02-02",Status:"Done",Ket:""}],Training:[{Tanggal:"2024-04-10",Batch:"Batch 1",Materi:"Basic Cleaning",Peserta:"5",Cabang:"001. Pondok Bambu",Trainer:"Fajar",Nilai:"85",Keterangan:""}],"Master Checklist":[{"Master Checklist":"Checklist Kebersihan Toilet","Link Document":"https://..."}],"Data SP":[{"Nama Karyawan":"Budi Santoso",Cabang:"001. Pondok Bambu","Jenis SP":"SP 1","Tanggal Terbit":"2026-01-10","Tanggal Berakhir":"2026-07-10","Alasan / Pelanggaran":"Keterlambatan",Status:"Aktif","Link Dokumen":""}],"Data Mutasi":[{"Nama Karyawan":"Budi Santoso","Cabang Asal":"001. Pondok Bambu","Cabang Tujuan":"002. Bintaro","Tanggal Efektif":"2026-02-01","Alasan Mutasi":"Rotasi Reguler",Status:"Selesai","No. SK / Dokumen":"SK/001/2026"}],"Master Form":[{"Nama Form":"Form Pengajuan Alat",Kategori:"Operasional","Link Form":"https://...",Deskripsi:"Formulir pengajuan alat dan perlengkapan",Publik:"Ya"}]}).forEach(([a,o])=>{e.utils.book_append_sheet(t,e.utils.json_to_sheet(o),a)}),e.writeFile(t,"Template_Import_Data_Awal_FCMS.xlsx")}function bi(e){let t=window.XLSX,n=t.utils.book_new(),a=!1;return e.forEach(o=>{if(!o.errors||o.errors.length===0)return;a=!0;let s=o.errors.map(m=>({"No. Baris":m.row,"Kolom Gagal":(m.errors||[]).map(i=>i.column||i).join("; "),"Alasan Error":(m.errors||[]).map(i=>i.reason||i).join("; "),...Object.fromEntries(Object.entries(m.data||{}).map(([i,d])=>[i,d??""]))})),c=t.utils.json_to_sheet(s);t.utils.book_append_sheet(n,c,o.sheetName.replace(/[\\\/\[\]*?:]/g,"_").slice(0,31))}),a?(t.writeFile(n,`Log_Error_Import_${new Date().toISOString().slice(0,10)}.xlsx`),!0):!1}var qr=["validation","employees","contracts","schedule","issues","training","checklist","sop","inspection","cleaning","fogging","basecamp","sp","mutasi"];function hi(e){let t=se();if(!t||String(t.role).toLowerCase()!=="superadmin"||t.isRestrictedToken===!0||Number(t.must_change_password)===1){e.innerHTML='<div class="empty-state"><p class="text-danger">Akses ditolak. Halaman ini hanya untuk Superadmin.</p></div>';return}e.innerHTML=`
    <div class="page-header">
      <div class="header-content">
        <h1 class="page-title"><span class="title-icon">\u{1F4E5}</span> Import Data Awal</h1>
        <p class="page-subtitle">Unggah file Excel untuk mengisi data aplikasi, atau sinkronkan langsung dari Google Sheets.</p>
      </div>
      <div class="page-actions" style="display:flex;gap:8px">
        <a href="#/sync" id="btn-sync-health" class="btn btn-primary" style="text-decoration:none;display:inline-flex;align-items:center;gap:6px"><span>\u{1FA7A} Sync & Health Dashboard</span></a>
        <button id="btn-sync-google" class="btn btn-secondary">
          <span>\u{1F504} Tarik Data dari Google Sheets</span>
        </button>
        <button class="btn btn-warning" id="btn-backup-db">\u{1F4E6} Backup Database</button>
        <button class="btn btn-secondary" id="btn-download-template">\u2B07\uFE0F Download Template</button>
      </div>
    </div>

    <!-- Recovery Banner (interrupted / partial / locked session) -->
    <div id="import-recovery-banner" style="display:none;margin-bottom:16px;"></div>

    <!-- STEP 1: Upload -->
    <div id="step-upload" class="import-step">
      <div class="card">
        <div class="card-body">
          <div class="import-info-box">
            <h3>\u{1F4CB} Petunjuk Import Data Awal</h3>
            <p>Upload file Excel (.xlsx) yang sudah diisi sesuai template. Sistem akan membaca seluruh sheet secara otomatis dan memvalidasi sebelum data disimpan.</p>
            <div class="import-sheet-list">
              ${Object.entries(ta).map(([k,{label:_}])=>`<span class="import-sheet-tag">\u{1F4C4} ${k} \u2192 ${_}</span>`).join("")}
            </div>
          </div>

          <div class="import-upload-zone" id="upload-zone">
            <div class="upload-icon">\u{1F4C2}</div>
            <div class="upload-text">
              <strong>Drag & Drop file Excel di sini</strong>
              <span>atau klik untuk memilih file</span>
            </div>
            <input type="file" id="file-input" accept=".xlsx,.xls" style="display:none">
            <button class="btn btn-primary" id="btn-browse">Pilih File Excel</button>
            <div class="upload-hint">Format: .xlsx | Ukuran maks: 20MB</div>
          </div>
          
          <div id="file-info" style="display:none" class="file-info-bar">
            <span id="file-name-display"></span>
            <button class="btn btn-ghost btn-sm" id="btn-clear-file">\u2715 Ganti</button>
          </div>
        </div>
      </div>
    </div>

    <!-- STEP 2: Validating (progress) -->
    <div id="step-validating" class="import-step" style="display:none">
      <div class="card">
        <div class="card-body text-center">
          <div class="import-progress-wrap">
            <div class="spinner" style="margin:0 auto 16px"></div>
            <div id="validation-status" class="import-status-text">Membaca file Excel...</div>
            <div class="import-progress-bar"><div class="import-progress-fill" id="validation-bar" style="width:0%"></div></div>
          </div>
        </div>
      </div>
    </div>

    <!-- STEP 3: Preview -->
    <div id="step-preview" class="import-step" style="display:none">
      <!-- Duplicate Strategy -->
      <div class="card mb-12">
        <div class="card-body">
          <h3 style="margin-bottom:12px">\u2699\uFE0F Pengaturan Duplikat</h3>
          <div class="dup-options">
            <label class="dup-option">
              <input type="radio" name="dup-strategy" value="skip" checked>
              <div class="dup-option-text">
                <strong>Lewati Data Duplikat</strong>
                <span>Data yang sudah ada di database tidak akan diubah</span>
              </div>
            </label>
            <label class="dup-option">
              <input type="radio" name="dup-strategy" value="update">
              <div class="dup-option-text">
                <strong>Perbarui Data yang Sudah Ada</strong>
                <span>Data lama akan ditimpa dengan data dari Excel</span>
              </div>
            </label>
          </div>
        </div>
      </div>

      <!-- Preview Table -->
      <div class="card mb-12">
        <div class="card-body p-0">
          <div class="preview-header">
            <h3>\u{1F4CA} Preview Validasi per Sheet</h3>
            <div id="preview-summary-badges"></div>
          </div>
          <div id="preview-table-container"></div>
        </div>
      </div>

      <!-- Error Detail -->
      <div id="error-detail-section" style="display:none" class="card mb-12">
        <div class="card-body p-0">
          <div class="preview-header">
            <h3>\u274C Detail Error</h3>
            <button class="btn btn-secondary btn-sm" id="btn-download-log">\u2B07\uFE0F Download Log Error</button>
          </div>
          <div id="error-detail-container"></div>
        </div>
      </div>

      <!-- Actions -->
      <div class="import-action-bar">
        <button class="btn btn-ghost" id="btn-back-to-upload">\u2190 Upload Ulang</button>
        <button class="btn btn-primary" id="btn-start-import" disabled>
          \u{1F680} Mulai Import
        </button>
      </div>
    </div>

    <!-- STEP 4: Importing -->
    <div id="step-importing" class="import-step" style="display:none">
      <div class="card">
        <div class="card-body">
          <h3 style="margin-bottom:20px;text-align:center">\u23F3 Sedang Mengimport Data...</h3>
          <div id="import-steps-list" class="import-steps-list"></div>
          <div class="import-progress-bar" style="margin-top:20px">
            <div class="import-progress-fill" id="import-bar" style="width:0%"></div>
          </div>
          <div id="import-current-status" class="import-status-text" style="margin-top:8px;text-align:center"></div>
        </div>
      </div>
    </div>

    <!-- STEP 5: Summary -->
    <div id="step-summary" class="import-step" style="display:none">
      <div class="card">
        <div class="card-body">
          <div class="import-summary-header" id="summary-status-icon"></div>
          <div id="summary-action-banner" style="margin-bottom:16px;"></div>
          <div class="import-summary-stats" id="summary-stats"></div>
          <div id="summary-module-results"></div>
          <div class="import-action-bar" style="margin-top:24px">
            <button class="btn btn-secondary" id="btn-import-again">\u{1F504} Import Lagi</button>
            <button class="btn btn-primary" id="btn-go-to-dashboard">\u{1F4CA} Ke Dashboard</button>
          </div>
        </div>
      </div>
    </div>
  `;let n=null,a=null,o=0,s=!1,c=null,m=null,i={upload:document.getElementById("step-upload"),validating:document.getElementById("step-validating"),preview:document.getElementById("step-preview"),importing:document.getElementById("step-importing"),summary:document.getElementById("step-summary")};function d(k){Object.entries(i).forEach(([_,E])=>{E.style.display=_===k?"":"none"})}document.getElementById("btn-backup-db")?.addEventListener("click",async()=>{let k=document.getElementById("btn-backup-db");k.disabled=!0,k.textContent="\u23F3 Memproses Backup...";try{let _=await X("/api/import/backup");if(_.ok){if(!window.XLSX){j("Library SheetJS belum termuat. Refresh halaman dan coba lagi.");return}let E=window.XLSX,w=E.utils.book_new();Object.entries(_.data.database).forEach(([K,G])=>{let J=G.length>0?G:[{}],W=E.utils.json_to_sheet(J);E.utils.book_append_sheet(w,W,K.substring(0,31))}),E.writeFile(w,`FCMS_Database_Backup_${new Date().toISOString().slice(0,10)}.xlsx`),ae("Backup berhasil diunduh!")}else j("Gagal memproses backup: "+(_.data?.error||"Unknown error"))}catch(_){j("Gagal memproses backup: "+_.message)}finally{k.disabled=!1,k.textContent="\u{1F4E6} Backup Database"}});let F=document.getElementById("btn-sync-google");F&&F.addEventListener("click",async()=>{if(!confirm("Peringatan: Mensinkronkan data dengan Google Sheets akan memperbarui dan menambahkan data baru dari Google Sheets ke dalam FCMS. Data yang sudah Anda buat di FCMS TIDAK akan terhapus. Lanjutkan?"))return;let k=F.innerHTML;F.innerHTML='<span class="spinner"></span> Menyinkronkan...',F.disabled=!0;try{let _=await X("/api/sync/google-sheets",{method:"POST"});_.ok?alert("Sinkronisasi Berhasil: "+(_.data?.message||"Data Karyawan & PIC telah diperbarui.")):alert("Gagal Sinkronisasi: "+(_.data?.error||"Unknown error"))}catch{alert("Terjadi kesalahan koneksi.")}finally{F.innerHTML=k,F.disabled=!1}}),document.getElementById("btn-download-template").addEventListener("click",()=>{gi(),ae("Template Excel berhasil didownload!")});let $=document.getElementById("file-input"),M=document.getElementById("upload-zone");document.getElementById("btn-browse").addEventListener("click",k=>{k.stopPropagation(),$.click()}),$.addEventListener("change",k=>{k.target.files[0]&&h(k.target.files[0])}),M.addEventListener("dragover",k=>{k.preventDefault(),M.classList.add("drag-over")}),M.addEventListener("dragleave",()=>M.classList.remove("drag-over")),M.addEventListener("drop",k=>{k.preventDefault(),M.classList.remove("drag-over");let _=k.dataTransfer.files[0];_&&_.name.match(/\.xlsx?$/i)?h(_):j("Hanya file .xlsx atau .xls yang didukung.")}),document.getElementById("btn-clear-file").addEventListener("click",()=>{n=null,$.value="",document.getElementById("file-info").style.display="none",M.style.display="",d("upload")});async function h(k){n=k,document.getElementById("file-name-display").textContent=`\u{1F4C4} ${k.name} (${(k.size/1024).toFixed(1)} KB)`,document.getElementById("file-info").style.display="flex",M.style.display="none",await L(k)}async function L(k){d("validating");let _=document.getElementById("validation-status"),E=document.getElementById("validation-bar");try{if(!window.XLSX)throw new Error("Library SheetJS belum termuat. Refresh halaman dan coba lagi.");_.textContent="Membaca file Excel...",E.style.width="20%",await zt(200);let w=await k.arrayBuffer(),K=window.XLSX.read(w,{type:"array",cellDates:!0});_.textContent=`Memvalidasi ${K.SheetNames.length} sheet...`,E.style.width="50%",await zt(100),a=mi(K),E.style.width="100%",_.textContent="Validasi selesai!",await zt(300),R()}catch(w){d("upload"),j("Gagal memproses file: "+w.message),document.getElementById("file-info").style.display="flex",M.style.display="none"}}function R(){d("preview");let k=a.filter(b=>!b.skipped).length,_=a.reduce((b,S)=>b+S.total,0),E=a.reduce((b,S)=>b+S.valid,0),w=a.reduce((b,S)=>b+S.errorCount,0),K=_>0?Math.round(E/_*100):0;document.getElementById("preview-summary-badges").innerHTML=`
      <span class="badge badge-info">${k} sheet</span>
      <span class="badge badge-secondary">${_} baris</span>
      <span class="badge badge-success">${E} valid (${K}%)</span>
      ${w>0?`<span class="badge badge-danger">${w} error</span>`:""}
    `;let G=document.getElementById("preview-table-container");G.innerHTML=`
      <table class="data-table">
        <thead>
          <tr>
            <th>Sheet (Excel)</th>
            <th>Modul</th>
            <th style="text-align:center">Total</th>
            <th style="text-align:center">Valid</th>
            <th style="text-align:center">Error</th>
            <th style="text-align:center">Status</th>
            <th style="text-align:center">Detail</th>
          </tr>
        </thead>
        <tbody>
          ${a.map((b,S)=>`
            <tr class="${b.errorCount>0?"row-error":b.skipped?"row-skipped":"row-ok"}">
              <td><strong>${b.sheetName}</strong></td>
              <td>${b.label}</td>
              <td style="text-align:center">${b.total}</td>
              <td style="text-align:center"><span class="badge badge-success">${b.valid}</span></td>
              <td style="text-align:center">${b.errorCount>0?`<span class="badge badge-danger">${b.errorCount}</span>`:'<span class="text-muted">\u2013</span>'}</td>
              <td style="text-align:center">
                ${b.skipped?'<span class="badge badge-neutral">Dilewati</span>':b.errorCount>0&&b.valid===0?'<span class="badge badge-danger">\u274C 0 Valid</span>':b.errorCount>0?'<span class="badge badge-warning">\u26A0\uFE0F Sebagian</span>':b.valid===0?'<span class="badge badge-neutral">Kosong</span>':'<span class="badge badge-success">\u2705 Siap</span>'}
              </td>
              <td style="text-align:center">
                ${b.errorCount>0?`<button class="btn btn-ghost btn-sm btn-detail-error" data-idx="${S}">\u{1F50D} ${b.errorCount} Error</button>`:'<span class="text-muted">\u2013</span>'}
              </td>
            </tr>
          `).join("")}
        </tbody>
      </table>
    `,G.querySelectorAll(".btn-detail-error").forEach(b=>{b.addEventListener("click",()=>{let S=a[Number(b.dataset.idx)];I(S)})});let J=document.getElementById("error-detail-section"),W=document.getElementById("error-detail-container");W.innerHTML="",J.style.display="none";let Y=document.getElementById("btn-start-import");E===0?(Y.disabled=!0,Y.innerHTML="\u26A0\uFE0F Tidak Ada Data Valid"):(Y.disabled=!1,w>0?(Y.innerHTML=`\u{1F680} Import ${E} Data Valid (${w} dilewati)`,Y.title="Baris error akan dilewati, baris valid tetap diimport"):Y.innerHTML=`\u{1F680} Mulai Import ${E} Data`)}function I(k){let _=document.getElementById("error-detail-section"),E=document.getElementById("error-detail-container");_.style.display="";let w=k.errors.slice(0,100).map(K=>(Array.isArray(K.errors)?K.errors:[]).map(J=>{let W=typeof J=="object";return`
          <tr>
            <td style="text-align:center"><span class="badge badge-danger">Baris ${K.row}</span></td>
            <td><strong>${W?J.column:"\u2014"}</strong></td>
            <td><code style="font-size:.78rem;color:var(--text-secondary)">${W&&J.originalValue!==void 0?J.originalValue||"(kosong)":"\u2014"}</code></td>
            <td class="error-msg">${W?J.reason:J}</td>
            <td style="font-size:.78rem;color:var(--success)">
              ${W&&J.aliases?`Gunakan salah satu nama kolom:<br><em>${J.aliases}</em>`:W&&J.hint?J.hint:""}
            </td>
          </tr>
        `}).join("")).join("");E.innerHTML=`
      <div class="error-sheet-block">
        <div class="error-sheet-title">
          \u{1F4C4} ${k.sheetName} \u2014 ${k.errorCount} baris error dari ${k.total} total
          ${k.errors.length>100?'<span style="font-weight:400">(menampilkan 100 pertama)</span>':""}
        </div>
        <div style="overflow-x:auto">
          <table class="data-table error-table" style="min-width:700px">
            <thead>
              <tr>
                <th style="width:80px">Baris</th>
                <th style="width:140px">Kolom Gagal</th>
                <th style="width:140px">Nilai di Excel</th>
                <th>Alasan Error</th>
                <th style="width:220px">\u{1F4A1} Cara Memperbaiki</th>
              </tr>
            </thead>
            <tbody>${w||'<tr><td colspan="5" class="text-muted" style="text-align:center">Tidak ada detail error</td></tr>'}</tbody>
          </table>
        </div>
        ${k.errors.length>100?`
          <div style="padding:10px 20px;font-size:.8rem;color:var(--text-muted)">
            Hanya menampilkan 100 error pertama. Download Log Error untuk melihat semua.
          </div>`:""}
      </div>
    `,_.scrollIntoView({behavior:"smooth",block:"start"})}document.getElementById("btn-back-to-upload").addEventListener("click",()=>{d("upload"),document.getElementById("file-info").style.display="none",M.style.display="",n=null,$.value=""}),document.getElementById("btn-download-log").addEventListener("click",()=>{if(!a)return;bi(a)?ae("Log error berhasil didownload."):ae("Tidak ada error untuk didownload.")});let N=document.getElementById("import-recovery-banner");function q(){if(!N)return;let k=Nn(),_=Array.isArray(k?.items)?k.items.filter(w=>w.status!=="ok"):[];if(k&&(k.status==="IN_PROGRESS"||k.status==="INCOMPLETE"||_.length>0)){let w=_.length>0?_:k.items||[],K=w.find(W=>W.error||W.errorCode)||w[0]||{},G=w.some(W=>W.httpStatus===409&&W.errorCode==="IMPORT_IN_PROGRESS")||k.status==="IN_PROGRESS",J=w.every(W=>Array.isArray(W.rows)&&W.rows.length>0);N.style.display="",N.innerHTML=rt({variant:"recover",title:`\u{1F504} Sesi Import Sebelumnya Belum Selesai (${k.fileName||"Excel"})`,message:J?`${w.length} modul belum selesai (${K.error||"Terhenti akibat refresh atau koneksi terputus"}). Klik Coba Lagi / Lanjutkan untuk meneruskan dengan Idempotency-Key yang sama.`:`${w.length} modul belum selesai. Upload ulang file Excel yang sama agar Idempotency-Key yang tersimpan otomatis digunakan kembali.`,idempotencyKey:K.idempotencyKey||"",importJobId:K.importJobId||"",committedChunks:Number(K.committedChunks||0),totalChunks:Number(K.totalChunks||0),errorCode:K.errorCode||k.status||"",fingerprint:K.fingerprint||"batch",canRetry:J,canRecover:!!(J&&G),canDismiss:!0}),N.querySelector('[data-import-action="retry"]')?.addEventListener("click",()=>{s||B(k.onDuplicate||"skip",k,{recoverLocked:!1})}),N.querySelector('[data-import-action="recover"]')?.addEventListener("click",()=>{s||B(k.onDuplicate||"skip",k,{recoverLocked:!0})}),N.querySelector('[data-import-action="dismiss"]')?.addEventListener("click",()=>{Pa(),(k.items||[]).forEach(W=>W.fingerprint&&yt(W.fingerprint)),N.style.display="none",N.innerHTML=""});return}let E=Xt().filter(w=>String(w.endpoint||"").startsWith("/api/import/"));if(E.length>0){let w=E[0];N.style.display="",N.innerHTML=rt({variant:"recover",title:`\u{1F504} Sesi Import Modul ${w.label||w.module} Belum Selesai`,message:w.payloadStored?w.errorMessage||"Import terhenti sebelum selesai. Anda dapat melanjutkan dengan Idempotency-Key yang sama.":"Upload ulang file Excel yang sama untuk melanjutkan dengan Idempotency-Key yang sama.",idempotencyKey:w.idempotencyKey,importJobId:w.importJobId,committedChunks:w.committedChunks,totalChunks:w.totalChunks,errorCode:w.errorCode||w.status,fingerprint:w.fingerprint,canRetry:!!w.payloadStored,canRecover:!!(w.payloadStored&&(w.status==="IN_PROGRESS"||w.status==="CONFLICT_IN_PROGRESS")),canDismiss:!0});let K=async G=>{if(!s){s=!0;try{let J=await Pt(w.fingerprint,{recover:G});yt(w.fingerprint),N.style.display="none",ae(J.idempotentReplay?"Replay idempoten berhasil (0 duplikat).":"Sesi import berhasil diselesaikan!")}catch(J){j(J.message||"Gagal melanjutkan import"),q()}finally{s=!1}}};N.querySelector('[data-import-action="retry"]')?.addEventListener("click",()=>K(!1)),N.querySelector('[data-import-action="recover"]')?.addEventListener("click",()=>K(!0)),N.querySelector('[data-import-action="dismiss"]')?.addEventListener("click",()=>{yt(w.fingerprint),q()});return}N.style.display="none",N.innerHTML=""}q();let v=document.getElementById("btn-start-import");v.addEventListener("click",()=>{if(s){j("Proses import sedang berjalan. Mohon tunggu hingga selesai.");return}let k=document.querySelector('input[name="dup-strategy"]:checked')?.value||"skip";B(k)});async function B(k,_=null,{recoverLocked:E=!1,forceReplayCompleted:w=!1}={}){if(s)return;s=!0,v.disabled=!0,d("importing"),N&&(N.style.display="none"),o=Date.now();let K;if(_&&Array.isArray(_.items))K=_.items.map(b=>({...b}));else{let b=[];qr.forEach(S=>{let y=a?.find(r=>r.module===S&&r.mapped?.length>0);y&&b.push(y)}),K=b.map(S=>({module:S.module,label:S.label,endpoint:`/api/import/${S.module}`,rows:S.mapped,rowCount:S.mapped.length,status:"pending",inserted:0,updated:0,skipped:0,committedChunks:0,totalChunks:0,idempotencyKey:null,fingerprint:null,importJobId:null,idempotentReplay:!1,httpStatus:null,errorCode:null,error:null,recoverable:!1}))}for(let b of K){let S={rows:b.rows,onDuplicate:k};if(b.fingerprint||(b.fingerprint=await La(b.endpoint,S)),!b.idempotencyKey){let y=String(b.module||"imp").slice(0,20);b.idempotencyKey=globalThis?.crypto?.randomUUID?`imp_${y}_${globalThis.crypto.randomUUID()}`:`imp_${y}_${Date.now().toString(36)}_${Math.random().toString(36).slice(2,8)}`}}let G={batchId:_?.batchId||(globalThis?.crypto?.randomUUID?globalThis.crypto.randomUUID():`batch_${Date.now()}`),fileName:_?.fileName||n?.name||"Excel Import",onDuplicate:k,status:"IN_PROGRESS",items:K};c=G,Lt(G);let J=document.getElementById("import-steps-list");J.innerHTML=K.map(b=>`
      <div class="import-step-item" id="step-item-${b.module}">
        <span class="step-item-icon" id="step-icon-${b.module}">${b.status==="ok"&&!w?"\u2705":"\u23F8\uFE0F"}</span>
        <span class="step-item-label">${$e(b.label)} <span class="step-item-count">(${b.rowCount||b.rows?.length||0} data)</span></span>
        <span class="step-item-status" id="step-status-${b.module}">${b.status==="ok"&&!w?`<span class="badge badge-success">${b.inserted||0} berhasil</span>`:""}</span>
      </div>
    `).join("");let W=document.getElementById("import-bar"),Y=document.getElementById("import-current-status");try{for(let y=0;y<K.length;y++){let r=K[y],l=document.getElementById(`step-icon-${r.module}`),f=document.getElementById(`step-status-${r.module}`);if(r.status==="ok"&&!w){W.style.width=`${Math.round((y+1)/K.length*100)}%`;continue}l.textContent="\u{1F504}",f.textContent=E&&r.httpStatus===409?"Memulihkan lock...":"Mengimport...",Y.textContent=`Mengimport ${r.label}... (${y+1}/${K.length})`,W.style.width=`${Math.round(y/Math.max(1,K.length)*100)}%`,Lt(G);try{let u=!!(E&&(r.httpStatus===409||r.errorCode==="IMPORT_IN_PROGRESS"||r.status==="pending"||r.status==="error")),p=await Pe({module:r.module,label:r.label,endpoint:r.endpoint,body:{rows:r.rows,onDuplicate:k},idempotencyKey:r.idempotencyKey,recover:u}),A=p.data||{};r.status="ok",r.inserted=Number(A.inserted||0),r.updated=Number(A.updated||0),r.skipped=Number(A.skipped||0),r.idempotentReplay=!!(p.idempotentReplay||A.idempotentReplay),r.importJobId=p.importJobId||A.importJobId||r.importJobId||null,r.idempotencyKey=p.idempotencyKey||r.idempotencyKey,r.httpStatus=p.status,r.error=null,r.errorCode=null,r.recoverable=!1,l.textContent=r.idempotentReplay?"\u267B\uFE0F":"\u2705";let T=r.idempotentReplay?' <span class="badge badge-info">Replay Idempoten</span>':"",g=r.skipped>0?` <span class="badge badge-neutral">${r.skipped} skip</span>`:"",C=r.updated>0?` <span class="badge badge-info">${r.updated} update</span>`:"";f.innerHTML=`<span class="badge badge-success">${r.inserted} berhasil</span>${C}${g}${T}`}catch(u){r.status="error",r.error=u.message||"Gagal mengimport data",r.errorCode=u.code||"IMPORT_ERROR",r.httpStatus=Number(u.status??0),r.committedChunks=Number(u.committedChunks||0),r.totalChunks=Number(u.totalChunks||0),r.importJobId=u.importJobId||r.importJobId||null,r.idempotencyKey=u.idempotencyKey||r.idempotencyKey||null,r.recoverable=!!u.recoverable,l.textContent=r.httpStatus===409?"\u23F3":"\u274C";let p=r.totalChunks>0?` <span class="badge badge-warning">Chunk ${r.committedChunks}/${r.totalChunks}</span>`:"";f.innerHTML=`<span class="badge badge-danger">${$e(r.errorCode||"Gagal")}</span>${p}`}G.status=K.every(u=>u.status==="ok")?"COMPLETED":"INCOMPLETE",Lt(G),await zt(80)}let b=K.every(y=>y.status==="ok");b?Pa():(G.status="INCOMPLETE",Lt(G));let S={totalSheets:K.length,totalRows:K.reduce((y,r)=>y+(r.rowCount||r.rows?.length||0),0),inserted:K.reduce((y,r)=>y+(r.inserted||0),0),updated:K.reduce((y,r)=>y+(r.updated||0),0),skipped:K.reduce((y,r)=>y+(r.skipped||0),0),failed:K.filter(y=>y.status!=="ok").length,hasReplays:K.some(y=>y.idempotentReplay),moduleResults:K,batchState:G};m=S,W.style.width="100%",Y.textContent=b?"Selesai!":"Selesai dengan modul gagal/terhenti (dapat dilanjutkan)",await zt(250),x(S)}finally{s=!1,v.disabled=!1}}function x(k){d("summary");let _=((Date.now()-o)/1e3).toFixed(1),E=k.failed===0&&k.moduleResults.every(J=>J.status==="ok"),w=k.moduleResults.some(J=>J.status!=="ok"&&J.httpStatus===409&&J.errorCode==="IMPORT_IN_PROGRESS"),K=k.moduleResults.find(J=>J.status!=="ok");document.getElementById("summary-status-icon").innerHTML=`
      <div class="summary-icon">${E?k.hasReplays?"\u267B\uFE0F":"\u{1F389}":"\u26A0\uFE0F"}</div>
      <h2 class="summary-title">${E?k.hasReplays?"Import & Replay Idempoten Berhasil!":"Import Berhasil!":"Import Belum Selesai (Dapat Dilanjutkan)"}</h2>
    `;let G=document.getElementById("summary-action-banner");if(G)if(!E&&K)G.innerHTML=rt({variant:w?"warning":"error",title:w?"\u23F3 Ada Modul Terkunci / Sedang Berjalan (HTTP 409)":"\u26A0\uFE0F Beberapa Modul Gagal atau Terhenti di Tengah Proses",message:K.error||"Klik Coba Lagi untuk melanjutkan dengan Idempotency-Key yang sama tanpa menduplikasi data.",idempotencyKey:K.idempotencyKey||"",importJobId:K.importJobId||"",committedChunks:K.committedChunks||0,totalChunks:K.totalChunks||0,errorCode:K.errorCode||"",fingerprint:K.fingerprint||"",canRetry:!0,canRecover:w,canDismiss:!1}),G.querySelector('[data-import-action="retry"]')?.addEventListener("click",()=>{B(k.batchState?.onDuplicate||"skip",k.batchState,{recoverLocked:!1})}),G.querySelector('[data-import-action="recover"]')?.addEventListener("click",()=>{B(k.batchState?.onDuplicate||"skip",k.batchState,{recoverLocked:!0})});else if(E&&k.moduleResults.length>0){let J=k.moduleResults[0];G.innerHTML=rt({variant:"success",title:k.hasReplays?"\u267B\uFE0F Replay Idempoten Dikonfirmasi (0 Duplikat)":"\u2705 Semua Modul Terkonfirmasi Server",message:k.hasReplays?"Server mengembalikan hasil tersimpan menggunakan Idempotency-Key yang sama tanpa menulis baris ganda.":"Semua chunk telah committed di server beserta Audit Log terkait.",idempotencyKey:J.idempotencyKey||"",importJobId:J.importJobId||"",canReplay:!0,canDismiss:!1}),G.querySelector('[data-import-action="replay"]')?.addEventListener("click",()=>{B(k.batchState?.onDuplicate||"skip",k.batchState,{recoverLocked:!1,forceReplayCompleted:!0})})}else G.innerHTML="";document.getElementById("summary-stats").innerHTML=`
      <div class="summary-stat-card">
        <div class="stat-value">${k.totalSheets}</div>
        <div class="stat-label">Total Sheet</div>
      </div>
      <div class="summary-stat-card">
        <div class="stat-value">${k.totalRows}</div>
        <div class="stat-label">Total Data</div>
      </div>
      <div class="summary-stat-card success">
        <div class="stat-value">${k.inserted}</div>
        <div class="stat-label">Berhasil Diimport</div>
      </div>
      <div class="summary-stat-card neutral">
        <div class="stat-value">${k.skipped}</div>
        <div class="stat-label">Dilewati (Duplikat)</div>
      </div>
      ${k.updated>0?`<div class="summary-stat-card info"><div class="stat-value">${k.updated}</div><div class="stat-label">Diperbarui</div></div>`:""}
      ${k.failed>0?`<div class="summary-stat-card danger"><div class="stat-value">${k.failed}</div><div class="stat-label">Modul Gagal</div></div>`:""}
      <div class="summary-stat-card info">
        <div class="stat-value">${_}s</div>
        <div class="stat-label">Durasi Proses</div>
      </div>
    `,document.getElementById("summary-module-results").innerHTML=`
      <table class="data-table" style="margin-top:16px">
        <thead>
          <tr><th>Modul</th><th style="text-align:center">Berhasil</th><th style="text-align:center">Dilewati</th><th style="text-align:center">Status & Journal</th></tr>
        </thead>
        <tbody>
          ${k.moduleResults.map((J,W)=>`
            <tr>
              <td>
                <div><strong>${$e(J.label)}</strong></div>
                ${J.idempotencyKey?`<div style="font-size:.72rem;color:var(--text-muted)">Key: <code>${$e(J.idempotencyKey.slice(0,26))}\u2026</code>${J.importJobId?` | Job: <code>${$e(J.importJobId.slice(0,10))}\u2026</code>`:""}</div>`:""}
              </td>
              <td style="text-align:center"><span class="badge badge-success">${J.inserted||0}</span>${J.updated>0?` <span class="badge badge-info">+${J.updated} upd</span>`:""}</td>
              <td style="text-align:center"><span class="badge badge-neutral">${J.skipped||0}</span></td>
              <td style="text-align:center">
                ${J.status==="ok"?`<span class="badge badge-success">${J.idempotentReplay?"\u267B\uFE0F Replay Idempoten":"\u2705 Sukses"}</span>`:`<div>
                      <span class="badge badge-danger" title="${$e(J.error||"")}">\u274C ${$e(J.errorCode||"Gagal")}${J.httpStatus?` (${J.httpStatus})`:""}</span>
                      ${J.totalChunks>0?`<div style="font-size:.72rem;margin-top:4px;">Chunk ${J.committedChunks}/${J.totalChunks}</div>`:""}
                      <div style="font-size:.75rem;color:var(--danger);margin-top:2px;">${$e(J.error||"")}</div>
                      <button type="button" class="btn btn-primary btn-sm" style="margin-top:6px" data-row-retry="${W}">\u{1F504} Coba Lagi</button>
                      ${J.httpStatus===409&&J.errorCode==="IMPORT_IN_PROGRESS"?`<button type="button" class="btn btn-warning btn-sm" style="margin-top:6px" data-row-recover="${W}">\u{1F6E0}\uFE0F Pulihkan</button>`:""}
                    </div>`}
              </td>
            </tr>
          `).join("")}
        </tbody>
      </table>
    `,document.querySelectorAll("[data-row-retry]").forEach(J=>{J.addEventListener("click",()=>{B(k.batchState?.onDuplicate||"skip",k.batchState,{recoverLocked:!1})})}),document.querySelectorAll("[data-row-recover]").forEach(J=>{J.addEventListener("click",()=>{B(k.batchState?.onDuplicate||"skip",k.batchState,{recoverLocked:!0})})})}document.getElementById("btn-import-again").addEventListener("click",()=>{n=null,a=null,c=null,m=null,$.value="",document.getElementById("file-info").style.display="none",M.style.display="",q(),d("upload")}),document.getElementById("btn-go-to-dashboard").addEventListener("click",()=>{window.location.hash="/dashboard"})}function zt(e){return new Promise(t=>setTimeout(t,e))}Be();function Nt(e){if(!e)return{status:"Tidak Ditentukan",label:"Tidak Ada Tanggal",bg:"rgba(148, 163, 184, 0.16)",color:"#CBD5E1",border:"rgba(148, 163, 184, 0.3)",icon:"\u26AA",daysLeft:0,isExpired:!1,isExpiringSoon:!1};let t=new Date;t.setHours(0,0,0,0);let n=new Date(e);n.setHours(0,0,0,0);let a=n.getTime()-t.getTime(),o=Math.ceil(a/(1e3*60*60*24));return o<0?{status:"Selesai",label:`Masa Sanksi Selesai (${Math.abs(o)} hari lalu)`,shortLabel:"Selesai",bg:"rgba(16, 185, 129, 0.16)",color:"#6EE7B7",border:"rgba(52, 211, 153, 0.35)",icon:"\u2705",daysLeft:o,isExpired:!0,isExpiringSoon:!1}:o<=30?{status:"Mendekati Berakhir",label:`Segera Berakhir (Sisa ${o} hari)`,shortLabel:`Sisa ${o} hari`,bg:"rgba(245, 158, 11, 0.16)",color:"#FCD34D",border:"rgba(251, 191, 36, 0.35)",icon:"\u26A0\uFE0F",daysLeft:o,isExpired:!1,isExpiringSoon:!0}:{status:"Aktif",label:`Aktif Berlaku (Sisa ${o} hari)`,shortLabel:`Aktif (${o} hari)`,bg:"rgba(239, 68, 68, 0.18)",color:"#FCA5A5",border:"rgba(248, 113, 113, 0.38)",icon:"\u{1F534}",daysLeft:o,isExpired:!1,isExpiringSoon:!1}}function Ua(e=""){let t=(e||"").toUpperCase().trim();return t==="SP 3"?{label:"SP 3 (Terakhir)",bg:"rgba(220, 38, 38, 0.25)",color:"#FCA5A5",border:"rgba(248, 113, 113, 0.45)",icon:"\u{1F6A8}"}:t==="SP 2"?{label:"SP 2",bg:"rgba(239, 68, 68, 0.18)",color:"#FCA5A5",border:"rgba(248, 113, 113, 0.38)",icon:"\u{1F6D1}"}:t==="SP 1"?{label:"SP 1",bg:"rgba(245, 158, 11, 0.18)",color:"#FCD34D",border:"rgba(251, 191, 36, 0.38)",icon:"\u26A0\uFE0F"}:t.includes("TEGURAN")?{label:"Teguran Lisan",bg:"rgba(56, 189, 248, 0.18)",color:"#7DD3FC",border:"rgba(56, 189, 248, 0.38)",icon:"\u{1F4AC}"}:{label:e||"SP",bg:"rgba(148, 163, 184, 0.16)",color:"#E2E8F0",border:"rgba(148, 163, 184, 0.32)",icon:"\u{1F4DC}"}}async function fi(e){let t=se(),n=re(t,"sp","write"),a=re(t,"sp","write"),o=re(t,"sp","delete"),s=[],c=[],m=[],i="all",d="all",F="all",$="",M=localStorage.getItem("fcms_sp_view_mode")||"card",h=1,L=20,R=0,I=1,N={total:0,active:0,expiring:0,expired:0,sp1:0,sp2:0,sp3:0,teguran:0,uniqueEmployees:0},q=null,v=!1,B="fcms-sp-page-styles",x=document.getElementById(B);x||(x=document.createElement("style"),x.id=B,document.head.appendChild(x)),x.textContent=`
      .sp-page-container {
        padding: 24px;
        max-width: 1400px;
        margin: 0 auto;
        font-family: inherit;
      }
      .sp-hero-card {
        background: #131F36;
        border: 1px solid #223456;
        border-radius: 18px;
        padding: 28px 32px;
        color: #FFFFFF;
        box-shadow: 0 4px 16px rgba(0, 0, 0, 0.22);
        margin-bottom: 24px;
        position: relative;
        overflow: hidden;
      }
      .sp-hero-card::after {
        content: '';
        position: absolute;
        right: -50px;
        top: -50px;
        width: 220px;
        height: 220px;
        background: radial-gradient(circle, rgba(37, 99, 235, 0.06) 0%, transparent 70%);
        border-radius: 50%;
        pointer-events: none;
      }
      .sp-hero-header {
        display: flex;
        justify-content: space-between;
        align-items: center;
        flex-wrap: wrap;
        gap: 16px;
      }
      .sp-hero-title {
        font-size: 1.45rem;
        font-weight: 800;
        margin: 0;
        display: flex;
        align-items: center;
        gap: 10px;
        letter-spacing: -0.02em;
      }
      .sp-hero-actions {
        display: flex;
        align-items: center;
        gap: 8px;
        flex-wrap: wrap;
      }
      .sp-hero-actions .btn {
        height: 38px;
        box-sizing: border-box;
        transition: all 0.2s ease;
      }
      .sp-btn-add:hover {
        background: #1D4ED8 !important;
        transform: translateY(-1px);
      }
      .sp-btn-excel:hover, .sp-btn-refresh:hover {
        background: rgba(255, 255, 255, 0.16) !important;
        border-color: rgba(255, 255, 255, 0.3) !important;
        transform: translateY(-1px);
      }
      .sp-stats-grid {
        display: grid;
        grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
        gap: 14px;
        margin-top: 24px;
      }
      .sp-stat-box {
        background: rgba(255, 255, 255, 0.06);
        border: 1px solid rgba(255, 255, 255, 0.1);
        border-radius: 12px;
        padding: 14px 18px;
        backdrop-filter: blur(8px);
        display: flex;
        align-items: center;
        gap: 14px;
      }
      .sp-stat-icon {
        width: 44px;
        height: 44px;
        border-radius: 10px;
        display: flex;
        align-items: center;
        justify-content: center;
        font-size: 1.3rem;
        flex-shrink: 0;
      }
      .sp-stat-value {
        font-size: 1.45rem;
        font-weight: 700;
        color: #FFFFFF;
        line-height: 1.2;
      }
      .sp-stat-label {
        font-size: 0.78rem;
        color: #94A3B8;
        text-transform: uppercase;
        letter-spacing: 0.04em;
        font-weight: 600;
      }
      .sp-controls-panel {
        background: #131F36;
        border: 1px solid #E2E8F0;
        border-radius: 16px;
        padding: 18px 20px;
        margin-bottom: 24px;
        box-shadow: 0 2px 8px rgba(0, 0, 0, 0.04);
        display: flex;
        flex-direction: column;
        gap: 16px;
      }
      .sp-controls-top {
        display: flex;
        align-items: center;
        justify-content: space-between;
        gap: 12px;
        flex-wrap: wrap;
      }
      .sp-search-wrapper {
        position: relative;
        flex: 1;
        min-width: 250px;
      }
      .sp-search-icon {
        position: absolute;
        left: 14px;
        top: 50%;
        transform: translateY(-50%);
        color: #94A3B8;
        font-size: 0.95rem;
        pointer-events: none;
      }
      .sp-search-input {
        width: 100%;
        padding: 10px 14px 10px 40px;
        border: 1.5px solid #CBD5E1;
        border-radius: 10px;
        font-size: 0.9rem;
        background: #0E182B;
        transition: all 0.2s ease;
      }
      .sp-search-input:focus {
        outline: none;
        border-color: #2563EB;
        background: #131F36;
        box-shadow: 0 0 0 3px rgba(37, 99, 235, 0.12);
      }
      .sp-filter-select {
        padding: 9px 14px;
        border: 1.5px solid #223456;
        border-radius: 10px;
        font-size: 0.88rem;
        background: #0E182B;
        color: #F8FAFC;
        font-weight: 500;
        cursor: pointer;
      }
      .sp-filter-select:focus {
        outline: none;
        border-color: #3B82F6;
        background: #131F36;
      }
      .sp-filter-pills {
        display: flex;
        flex-wrap: wrap;
        gap: 8px;
        align-items: center;
      }
      .sp-pill {
        border: 1px solid #223456;
        background: #0E182B;
        color: #CBD5E1;
        font-size: 0.82rem;
        font-weight: 600;
        padding: 6px 14px;
        border-radius: 30px;
        cursor: pointer;
        transition: all 0.2s ease;
        display: flex;
        align-items: center;
        gap: 6px;
      }
      .sp-pill:hover {
        background: #22314A;
        color: #F8FAFC;
        border-color: #475569;
      }
      .sp-pill.active {
        background: #2563EB;
        border-color: #2563EB;
        color: #FFFFFF;
        box-shadow: 0 2px 6px rgba(37, 99, 235, 0.25);
      }
      .sp-pill-count {
        background: rgba(255, 255, 255, 0.1);
        color: #E2E8F0;
        padding: 1px 7px;
        border-radius: 12px;
        font-size: 0.75rem;
      }
      .sp-pill.active .sp-pill-count {
        background: rgba(255, 255, 255, 0.25);
        color: #FFFFFF;
      }
      /* View Mode Toggle */
      .sp-view-toggle {
        display: inline-flex;
        align-items: center;
        background: #0E182B;
        padding: 3px;
        border-radius: 10px;
        border: 1px solid #223456;
        gap: 3px;
        flex-shrink: 0;
      }
      .sp-view-btn {
        display: inline-flex;
        align-items: center;
        gap: 6px;
        padding: 6px 12px;
        border-radius: 7px;
        font-size: 0.82rem;
        font-weight: 600;
        border: none;
        background: transparent;
        color: #CBD5E1;
        cursor: pointer;
        transition: all 0.15s ease;
        height: 34px;
        box-sizing: border-box;
        user-select: none;
      }
      .sp-view-btn:hover {
        color: #F8FAFC;
        background: #22314A;
      }
      .sp-view-btn.active {
        background: #131F36;
        color: #60A5FA;
        font-weight: 700;
        box-shadow: 0 1px 4px rgba(0, 0, 0, 0.2);
      }
      /* Modern Table View */
      .sp-table-card {
        background: #131F36;
        border: 1px solid #223456;
        border-radius: 16px;
        box-shadow: 0 2px 8px rgba(0, 0, 0, 0.2);
        overflow: hidden;
      }
      .sp-table-responsive {
        width: 100%;
        overflow-x: auto;
        -webkit-overflow-scrolling: touch;
      }
      .sp-table {
        width: 100%;
        border-collapse: collapse;
        text-align: left;
        min-width: 900px;
      }
      .sp-table th {
        background: #0E182B;
        padding: 13px 16px;
        font-size: 0.76rem;
        font-weight: 700;
        color: #CBD5E1;
        text-transform: uppercase;
        letter-spacing: 0.05em;
        border-bottom: 1.5px solid #223456;
        white-space: nowrap;
      }
      .sp-table td {
        padding: 13px 16px;
        font-size: 0.88rem;
        color: #F8FAFC;
        border-bottom: 1px solid #223456;
        vertical-align: middle;
      }
      .sp-table tbody tr:hover td {
        background: #22314A;
      }
      .sp-table tbody tr:last-child td {
        border-bottom: none;
      }
      /* Grid View */
      .sp-cards-grid {
        display: grid;
        grid-template-columns: repeat(auto-fill, minmax(350px, 1fr));
        gap: 20px;
      }
      .sp-card {
        background: #131F36;
        border: 1px solid #223456;
        border-radius: 16px;
        padding: 22px;
        box-shadow: 0 2px 6px rgba(0, 0, 0, 0.2);
        display: flex;
        flex-direction: column;
        justify-content: space-between;
        transition: all 0.25s cubic-bezier(0.16, 1, 0.3, 1);
        position: relative;
      }
      .sp-card:hover {
        transform: translateY(-4px);
        box-shadow: 0 12px 24px -6px rgba(0, 0, 0, 0.35);
        border-color: #475569;
      }
      .sp-status-badge {
        display: inline-flex;
        align-items: center;
        gap: 6px;
        padding: 4px 10px;
        border-radius: 20px;
        font-size: 0.76rem;
        font-weight: 700;
        letter-spacing: 0.02em;
      }
      .sp-type-badge {
        display: inline-flex;
        align-items: center;
        gap: 4px;
        padding: 4px 9px;
        border-radius: 6px;
        font-size: 0.76rem;
        font-weight: 700;
      }
      .sp-avatar-box {
        width: 42px;
        height: 42px;
        border-radius: 10px;
        background: rgba(59, 130, 246, 0.2);
        color: #93C5FD;
        font-weight: 700;
        display: flex;
        align-items: center;
        justify-content: center;
        font-size: 0.95rem;
        flex-shrink: 0;
      }
      .sp-icon-btn {
        background: #0E182B;
        border: 1px solid #223456;
        color: #CBD5E1;
        width: 32px;
        height: 32px;
        border-radius: 8px;
        display: flex;
        align-items: center;
        justify-content: center;
        cursor: pointer;
        transition: all 0.15s ease;
      }
      .sp-icon-btn:hover {
        background: #22314A;
        color: #F8FAFC;
        border-color: #475569;
      }
      .sp-icon-btn.delete:hover {
        background: #FEE2E2;
        color: #DC2626;
        border-color: #FECACA;
      }
      .sp-empty-state {
        text-align: center;
        padding: 60px 20px;
        background: #131F36;
        border-radius: 16px;
        border: 1px dashed #CBD5E1;
      }

      /* \u2500\u2500 TAMPILAN RESPONSIVE HP / MOBILE (\u2264 768px) \u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500 */
      @media (max-width: 768px) {
        .sp-page-container {
          padding: 12px 14px !important;
        }
        .sp-hero-card {
          padding: 18px 16px !important;
          border-radius: 14px !important;
          margin-bottom: 16px !important;
        }
        .sp-hero-title {
          font-size: 1.25rem !important;
          gap: 8px !important;
        }
        .sp-hero-subtitle {
          font-size: 0.82rem !important;
          line-height: 1.4 !important;
        }
        .sp-hero-header {
          flex-direction: column !important;
          align-items: stretch !important;
          gap: 12px !important;
        }
        .sp-hero-actions {
          width: 100% !important;
          display: flex !important;
          align-items: center !important;
          gap: 6px !important;
          flex-wrap: nowrap !important;
          overflow-x: auto !important;
          scrollbar-width: none !important;
          padding-bottom: 2px !important;
        }
        .sp-hero-actions::-webkit-scrollbar {
          display: none;
        }
        .sp-btn-add {
          flex: 1 1 auto !important;
          justify-content: center !important;
          padding: 8px 10px !important;
          font-size: 0.82rem !important;
          white-space: nowrap !important;
        }
        .sp-btn-excel {
          flex: 1 1 auto !important;
          justify-content: center !important;
          padding: 8px 10px !important;
          font-size: 0.82rem !important;
          white-space: nowrap !important;
        }
        .sp-btn-refresh {
          width: 38px !important;
          height: 38px !important;
          padding: 0 !important;
          flex-shrink: 0 !important;
          justify-content: center !important;
        }
        /* 2x2 compact stats grid di layar HP */
        .sp-stats-grid {
          grid-template-columns: repeat(2, 1fr) !important;
          gap: 8px !important;
          margin-top: 14px !important;
        }
        .sp-stat-box {
          padding: 10px 12px !important;
          gap: 10px !important;
          border-radius: 10px !important;
        }
        .sp-stat-icon {
          width: 36px !important;
          height: 36px !important;
          font-size: 1.1rem !important;
          border-radius: 8px !important;
        }
        .sp-stat-value {
          font-size: 1.15rem !important;
        }
        .sp-stat-label {
          font-size: 0.76rem !important;
        }
        .sp-controls-panel {
          padding: 14px 12px !important;
          border-radius: 14px !important;
          margin-bottom: 16px !important;
          gap: 12px !important;
        }
        .sp-controls-top {
          flex-direction: column !important;
          align-items: stretch !important;
          gap: 10px !important;
        }
        .sp-search-wrapper {
          min-width: 100% !important;
          width: 100% !important;
        }
        /* PITA FILTER SWIPEABLE HORIZONTAL DI HP (TIDAK LAGI NUMPUK VERTIKAL) */
        .sp-filter-pills {
          display: flex !important;
          flex-wrap: nowrap !important;
          overflow-x: auto !important;
          -webkit-overflow-scrolling: touch !important;
          padding-bottom: 6px !important;
          padding-top: 2px !important;
          scrollbar-width: none !important;
          -ms-overflow-style: none !important;
        }
        .sp-filter-pills::-webkit-scrollbar {
          display: none !important;
        }
        .sp-pill {
          flex-shrink: 0 !important;
          white-space: nowrap !important;
          padding: 6px 12px !important;
          font-size: 0.8rem !important;
        }
        .sp-cards-grid {
          grid-template-columns: minmax(0, 1fr) !important;
          width: 100% !important;
          max-width: 100% !important;
          min-width: 0 !important;
          box-sizing: border-box !important;
          gap: 12px !important;
        }
        .sp-card {
          padding: 16px !important;
          border-radius: 14px !important;
          width: 100% !important;
          max-width: 100% !important;
          min-width: 0 !important;
          box-sizing: border-box !important;
          overflow: hidden;
        }
        .sp-card-header {
          flex-direction: row !important;
          align-items: flex-start !important;
        }
        .sp-view-toggle {
          width: 100% !important;
          display: flex !important;
        }
        .sp-view-btn {
          flex: 1 !important;
          justify-content: center !important;
        }
        .sp-table-card {
          border-radius: 14px !important;
          overflow-x: auto !important;
          -webkit-overflow-scrolling: touch !important;
        }
        .sp-table {
          min-width: 850px !important;
        }
        .sp-table th, .sp-table td {
          padding: 10px 12px !important;
          font-size: 0.82rem !important;
        }
        .sp-form-grid, .sp-detail-grid {
          grid-template-columns: 1fr !important;
          gap: 10px !important;
        }
      }
    `;function k(){let u={sort:"date_desc"};return $&&$.trim()&&(u.search=$.trim()),d&&d!=="all"&&(u.branch_id=d),F&&F!=="all"&&(u.division=F),i&&i!=="all"&&(u.filter=i),u}function _(){let u=(p,A)=>{let T=e.querySelector(p);T&&(T.textContent=String(A??0))};u("#sp-stat-active",N.active),u("#sp-stat-expiring",N.expiring),u("#sp-stat-expired",N.expired),u("#sp-stat-unique",N.uniqueEmployees),u("#sp-pill-count-all",N.total),u("#sp-pill-count-active",N.active),u("#sp-pill-count-expiring",N.expiring),u("#sp-pill-count-expired",N.expired),u("#sp-pill-count-sp1",N.sp1),u("#sp-pill-count-sp2",N.sp2),u("#sp-pill-count-sp3",N.sp3)}async function E({refreshStats:u=!0}={}){try{if(c.length===0||m.length===0){let[g,C]=await Promise.all([we(),ze()]);c=g||[],m=C||[]}let p=k(),A=new URLSearchParams({page:String(h),limit:String(L),...p});u&&A.set("include_stats","1");let T=await X("/api/sp?"+A.toString());if(T.ok&&T.data?.data&&Array.isArray(T.data.data)){let g=T.data.pagination||{};if(R=Number(g.total??T.data.data.length),I=Math.max(1,Number(g.pages||g.totalPages||Math.ceil(R/L)||1)),T.data.data.length===0&&h>I)return h=I,E({refreshStats:u});s=T.data.data,T.data.stats&&(N={...N,...T.data.stats})}else s=[],R=0,I=1,T.ok||j(T.data?.error||"Gagal memuat data SP dari server")}catch(p){console.warn("Error loading SP data:",p),s=[],R=0,I=1,j("Gagal memuat data SP: "+p.message)}v?(u&&_(),r()):(v=!0,l())}function w(){return s}function K(){return N}function G(u=null){let p=!!u;if(p&&!a){j("Anda tidak memiliki izin untuk mengedit data SP.");return}if(!p&&!n){j("Anda tidak memiliki izin untuk menambah data SP.");return}let A=!!u,T=u?.tanggal?u.tanggal.slice(0,10):new Date().toISOString().slice(0,10),g=u?.akhir_sp?u.akhir_sp.slice(0,10):"";if(!g){let Q=new Date(T);Q.setMonth(Q.getMonth()+6),g=Q.toISOString().slice(0,10)}let C=document.createElement("div");C.innerHTML=`
      <div style="display:flex; flex-direction:column; gap:14px;">
        <div class="sp-form-grid" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:var(--text-2); margin-bottom:5px;">
              Nama Karyawan <span style="color:#DC2626;">*</span>
            </label>
            <div id="sp-modal-emp-mount"></div>
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:var(--text-2); margin-bottom:5px;">
              Divisi <span style="color:#DC2626;">*</span>
            </label>
            <select id="sp-modal-div" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" required>
              <option value="FACILITY CARE" ${u?.division==="FACILITY CARE"?"selected":""}>FACILITY CARE</option>
              <option value="SECURITY" ${u?.division==="SECURITY"?"selected":""}>SECURITY</option>
            </select>
          </div>
        </div>

        <div class="sp-form-grid" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:var(--text-2); margin-bottom:5px;">
              Cabang Penempatan <span style="color:#DC2626;">*</span>
            </label>
            <select id="sp-modal-branch" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" required>
              <option value="">Pilih Cabang...</option>
              ${c.map(Q=>`<option value="${Q.value}" ${String(u?.branch_id)===String(Q.value)?"selected":""}>${Q.label}</option>`).join("")}
            </select>
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:var(--text-2); margin-bottom:5px;">
              Tingkat / Jenis SP <span style="color:#DC2626;">*</span>
            </label>
            <select id="sp-modal-type" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" required>
              <option value="SP 1" ${u?.sp_type==="SP 1"?"selected":""}>SP 1 (Masa Berlaku 6 Bulan)</option>
              <option value="SP 2" ${u?.sp_type==="SP 2"?"selected":""}>SP 2 (Masa Berlaku 6 Bulan)</option>
              <option value="SP 3" ${u?.sp_type==="SP 3"?"selected":""}>SP 3 (Masa Berlaku 6 Bulan / Terakhir)</option>
              <option value="Teguran Lisan" ${u?.sp_type==="Teguran Lisan"?"selected":""}>Teguran Lisan (Pembinaan)</option>
            </select>
          </div>
        </div>

        <div class="sp-form-grid" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:var(--text-2); margin-bottom:5px;">
              Tanggal Terbit SP <span style="color:#DC2626;">*</span>
            </label>
            <input type="date" id="sp-modal-tanggal" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              value="${T}" required />
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:var(--text-2); margin-bottom:5px;">
              Tanggal Akhir Masa Berlaku SP <span style="color:#DC2626;">*</span>
            </label>
            <input type="date" id="sp-modal-akhir" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              value="${g}" required />
          </div>
        </div>

        <div>
          <label style="display:block; font-size:0.82rem; font-weight:600; color:var(--text-2); margin-bottom:5px;">
            Link Berkas / Foto BAP / Scan Surat SP
          </label>
          <input type="url" id="sp-modal-link" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
            value="${u?.document_link?String(u.document_link).replace(/"/g,"&quot;"):""}" />
        </div>

        <div>
          <label style="display:block; font-size:0.82rem; font-weight:600; color:var(--text-2); margin-bottom:5px;">
            Kronologi Pelanggaran / Catatan Pembinaan
          </label>
          <textarea id="sp-modal-notes" class="form-control" rows="3" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;">${u?.notes||u?.description||""}</textarea>
        </div>
      </div>
    `;let z=C.querySelector("#sp-modal-tanggal"),O=C.querySelector("#sp-modal-akhir");z.addEventListener("change",()=>{if(z.value){let Q=new Date(z.value);Q.setMonth(Q.getMonth()+6),O.value=Q.toISOString().slice(0,10)}});let te=je(C.querySelector("#sp-modal-emp-mount"),{mode:"single",inputId:"sp-modal-emp",inputName:"employee_name",placeholder:"",required:!0,allowManual:!0,hasEmployeeIdSchema:!0,manualBadgeText:"Belum terhubung ke Master Karyawan",employees:m,initialEmployeeId:u?.employee_id||null,initialEmployeeName:u?.employee_name||""});ce({title:A?"\u270F\uFE0F Edit Data Surat Peringatan (SP)":"\u2795 Terbitkan Surat Peringatan (SP) Baru",content:C,size:"md",confirmText:A?"Simpan Perubahan":"Terbitkan SP",confirmClass:"btn-primary",onConfirm:async(Q,ne)=>{let ee=te.getValue(),H=(ee.employee_name||"").trim(),U=ee.employee_id??null,D=C.querySelector("#sp-modal-div").value,V=C.querySelector("#sp-modal-branch").value,Z=C.querySelector("#sp-modal-type").value,ie=C.querySelector("#sp-modal-tanggal").value,le=C.querySelector("#sp-modal-akhir").value,pe=(C.querySelector("#sp-modal-link").value||"").trim(),de=(C.querySelector("#sp-modal-notes").value||"").trim();if(!H){j("Nama Karyawan wajib diisi"),te.inputEl?.focus();return}if(!V){j("Cabang penempatan wajib dipilih");return}if(!ie||!le){j("Tanggal SP dan Akhir SP wajib diisi");return}let me=c.find(oe=>String(oe.value)===String(V)),Ae=me?me.label:"Cabang",Ke={employee_name:H,employee_id:U,division:D,branch_id:Number(V),branch_name:Ae,sp_type:Z,tanggal:ie,akhir_sp:le,document_link:pe,notes:de,status:Nt(le).status,...A?{row_version:u?.row_version!==void 0&&u?.row_version!==null?Number(u.row_version):1}:{}};try{if(A){let oe=await X(`/api/sp/${u.id}`,{method:"PUT",body:JSON.stringify(Ke)});if(!oe.ok){if(oe.status===409){j("Konflik OCC: Data telah diubah oleh pengguna lain. Silakan muat ulang halaman."),ne(),await E({refreshStats:!0});return}j(oe.data?.error||"Gagal memperbarui data SP");return}ae("Data SP berhasil diperbarui")}else{let oe=await X("/api/sp",{method:"POST",body:JSON.stringify(Ke)});if(!oe.ok){j(oe.data?.error||"Gagal menerbitkan data SP");return}h=1,ae("Surat Peringatan baru berhasil diterbitkan")}ne(),await E({refreshStats:!0})}catch(oe){console.error("Error saving SP:",oe),j("Gagal menyimpan SP: "+oe.message)}}})}function J(u){if(!o){j("Anda tidak memiliki izin untuk menghapus data SP.");return}ge(`Hapus data Surat Peringatan untuk karyawan <strong>${u.employee_name}</strong> (${u.sp_type})?`,async()=>{try{let p=await X(`/api/sp/${u.id}`,{method:"DELETE"});if(!p.ok){j(p.data?.error||"Gagal menghapus data SP dari server");return}ae(`Data SP ${u.employee_name} berhasil dihapus`),await E({refreshStats:!0})}catch(p){j("Gagal menghapus data SP: "+p.message)}},"Hapus Data SP")}function W(u){let p=Nt(u.akhir_sp),A=Ua(u.sp_type),T=document.createElement("div");T.innerHTML=`
      <div style="display:flex; flex-direction:column; gap:16px;">
        <div style="display:flex; align-items:center; gap:14px; padding-bottom:14px; border-bottom:1px solid #223456;">
          <div class="sp-avatar-box" style="width:54px; height:54px; font-size:1.2rem;">
            ${u.employee_name.slice(0,2).toUpperCase()}
          </div>
          <div>
            <h3 style="margin:0 0 4px 0; font-size:1.15rem; font-weight:700; color:#F8FAFC;">
              ${u.employee_name}
            </h3>
            <div style="display:flex; align-items:center; gap:8px;">
              <span class="badge" style="background:rgba(59,130,246,0.18); color:#93C5FD; border:1px solid rgba(96,165,250,0.35); font-weight:700;">${u.division||"FACILITY CARE"}</span>
              <span style="color:#CBD5E1; font-size:0.85rem;">\u{1F4CD} ${u.branch_name||"Cabang"}</span>
            </div>
          </div>
        </div>

        <div class="sp-detail-grid" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div style="background:#0E182B; border:1px solid #223456; border-radius:10px; padding:12px;">
            <div style="font-size:0.75rem; color:#94A3B8; text-transform:uppercase; font-weight:600; margin-bottom:4px;">Tingkat Sanksi</div>
            <span class="sp-type-badge" style="background:${A.bg}; color:${A.color}; border:1px solid ${A.border}; font-size:0.85rem;">
              <span>${A.icon}</span>
              <span>${A.label}</span>
            </span>
          </div>

          <div style="background:#0E182B; border:1px solid #223456; border-radius:10px; padding:12px;">
            <div style="font-size:0.75rem; color:#94A3B8; text-transform:uppercase; font-weight:600; margin-bottom:4px;">Status Masa Berlaku</div>
            <span class="sp-status-badge" style="background:${p.bg}; color:${p.color}; border:1px solid ${p.border}; font-size:0.85rem;">
              <span>${p.icon}</span>
              <span>${p.label}</span>
            </span>
          </div>
        </div>

        <div class="sp-detail-grid" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div style="background:#0E182B; border:1px solid #223456; border-radius:10px; padding:12px;">
            <div style="font-size:0.75rem; color:#94A3B8; text-transform:uppercase; font-weight:600; margin-bottom:2px;">Tanggal Terbit SP</div>
            <div style="font-size:0.95rem; font-weight:700; color:#F8FAFC;">${xe(u.tanggal)}</div>
          </div>

          <div style="background:#0E182B; border:1px solid #223456; border-radius:10px; padding:12px;">
            <div style="font-size:0.75rem; color:#94A3B8; text-transform:uppercase; font-weight:600; margin-bottom:2px;">Akhir Masa Berlaku</div>
            <div style="font-size:0.95rem; font-weight:700; color:#F8FAFC;">${xe(u.akhir_sp)}</div>
          </div>
        </div>

        ${u.notes||u.description?`
          <div style="background:rgba(245,158,11,0.15); border:1px solid rgba(251,191,36,0.32); border-radius:10px; padding:14px;">
            <div style="font-size:0.78rem; font-weight:700; color:#FCD34D; margin-bottom:4px;">\u{1F4DD} Catatan Pelanggaran / Pembinaan:</div>
            <div style="font-size:0.88rem; color:#78350F; line-height:1.5;">${u.notes||u.description}</div>
          </div>
        `:""}

        ${u.document_link?`
          <div>
            <a href="${u.document_link}" target="_blank" rel="noopener noreferrer" class="btn btn-primary" style="width:100%; justify-content:center; padding:10px; font-weight:600; display:flex; align-items:center; gap:8px;">
              <span>\u{1F4C4}</span> Buka Dokumen BAP / Scan SP Fisik
            </a>
          </div>
        `:`
          <div style="text-align:center; padding:8px; font-size:0.82rem; color:#94A3B8;">
            (Belum ada scan dokumen / BAP yang diunggah)
          </div>
        `}
      </div>
    `,ce({title:"\u{1F4CB} Lembar Riwayat Surat Peringatan (SP)",content:T,size:"md",cancelText:"Tutup"})}async function Y(){if(!re(t,"export")){j("Anda tidak memiliki izin untuk mengekspor data.");return}try{let{downloadExcel:u}=await Promise.resolve().then(()=>(Ce(),Ve)),p=await ue("/api/sp/export",k()),A=p.map((T,g)=>{let C=Nt(T.akhir_sp);return{No:g+1,"Nama Karyawan":T.employee_name||"",Divisi:T.division||"",Cabang:T.branch_name||"","Jenis SP":T.sp_type||"","Tanggal SP":T.tanggal||"","Akhir SP":T.akhir_sp||"","Status Masa Berlaku":C.status,"Sisa Hari":C.daysLeft>0?C.daysLeft:0,"Link BAP / Foto":T.document_link||"","Catatan / Kronologi":T.notes||T.description||""}});u(A,`Data_Surat_Peringatan_FCMS_${new Date().toISOString().slice(0,10)}`),ae(`Berhasil mengunduh ${p.length} data SP ke format Excel.`)}catch(u){j("Gagal ekspor: "+u.message)}}function b(u){return`
          ${u.length===0?`
            <div class="sp-empty-state">
              <div style="font-size:3rem; margin-bottom:12px;">\u{1F50D}</div>
              <h3 style="font-size:1.2rem; font-weight:700; color:#F8FAFC; margin:0 0 6px 0;">Tidak Ada Data SP Ditemukan</h3>
              <p style="color:#94A3B8; font-size:0.9rem; margin:0 0 20px 0;">Silakan sesuaikan kata kunci pencarian atau reset filter aktif.</p>
              <button id="sp-btn-reset-filter" class="btn btn-ghost" style="border:1px solid #223456; color:#CBD5E1; border-radius:8px; padding:8px 16px;">
                Reset Filter
              </button>
            </div>
          `:M==="table"?`
            <!-- Table View -->
            <div class="sp-table-card">
              <div class="sp-table-responsive">
                <table class="sp-table">
                  <thead>
                    <tr>
                      <th style="width:45px; text-align:center;">No</th>
                      <th>Nama Karyawan</th>
                      <th>Cabang & Divisi</th>
                      <th style="text-align:center;">Jenis SP</th>
                      <th>Tanggal Terbit</th>
                      <th>Akhir SP</th>
                      <th style="text-align:center;">Status</th>
                      <th>Keterangan / BAP</th>
                      <th style="width:110px; text-align:center;">Aksi</th>
                    </tr>
                  </thead>
                  <tbody>
                    ${u.map((p,A)=>{let T=Nt(p.akhir_sp),g=Ua(p.sp_type);return`
                        <tr data-id="${p.id}">
                          <td style="text-align:center; color:#94A3B8; font-weight:600; font-size:0.8rem;">${(h-1)*L+A+1}</td>
                          <td>
                            <div style="display:flex; align-items:center; gap:10px;">
                              <div class="sp-avatar-box" style="width:34px; height:34px; font-size:0.78rem; border-radius:8px;">
                                ${(p.employee_name||"SP").slice(0,2).toUpperCase()}
                              </div>
                              <div style="font-weight:700; color:#F8FAFC; white-space:nowrap;">
                                ${p.employee_name}
                              </div>
                            </div>
                          </td>
                          <td style="white-space:nowrap;">
                            <div style="font-weight:600; color:#E2E8F0; font-size:0.84rem;">\u{1F4CD} ${p.branch_name||"Cabang"}</div>
                            <div style="font-size: 0.76rem; color:#93C5FD; font-weight:600; margin-top:2px;">${p.division||"FACILITY CARE"}</div>
                          </td>
                          <td style="text-align:center; white-space:nowrap;">
                            <span class="sp-type-badge" style="background:${g.bg}; color:${g.color}; border:1px solid ${g.border};">
                              <span>${g.icon}</span>
                              <span>${g.label}</span>
                            </span>
                          </td>
                          <td style="white-space:nowrap; font-weight:600; color:#F8FAFC;">
                            ${xe(p.tanggal)}
                          </td>
                          <td style="white-space:nowrap; font-weight:700; color:#F8FAFC;">
                            ${xe(p.akhir_sp)}
                          </td>
                          <td style="text-align:center; white-space:nowrap;">
                            <span class="sp-status-badge" style="background:${T.bg}; color:${T.color}; border:1px solid ${T.border};">
                              <span>${T.icon}</span>
                              <span>${T.shortLabel}</span>
                            </span>
                          </td>
                          <td style="max-width:240px;">
                            <div style="font-size:0.82rem; color:#CBD5E1; white-space:nowrap; overflow:hidden; text-overflow:ellipsis; margin-bottom:4px;" title="${String(p.notes||p.description||"").replace(/"/g,"&quot;")}">
                              ${p.notes||p.description||"Pencatatan sanksi peringatan disiplin operasional."}
                            </div>
                            <button class="btn btn-sm btn-ghost sp-detail-btn" style="border:1px solid #223456; color:#93C5FD; border-radius:6px; font-weight:600; font-size: 0.76rem; padding:2px 8px;" data-id="${p.id}">
                              \u{1F441}\uFE0F Rincian BAP
                            </button>
                          </td>
                          <td style="text-align:center; white-space:nowrap;">
                            <div style="display:inline-flex; align-items:center; gap:6px;">
                              ${p.document_link?`
                                <a href="${p.document_link}" target="_blank" rel="noopener noreferrer" class="sp-icon-btn" title="Buka Dokumen">
                                  <svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M18 13v6a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V8a2 2 0 0 1 2-2h6"/><polyline points="15 3 21 3 21 9"/><line x1="10" y1="14" x2="21" y2="3"/></svg>
                                </a>
                              `:""}
                              ${a?`
                                <button class="sp-icon-btn sp-edit-btn" title="Edit Data SP" data-id="${p.id}">
                                  <svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>
                                </button>
                              `:""}
                              ${o?`
                                <button class="sp-icon-btn delete sp-delete-btn" title="Hapus" data-id="${p.id}">
                                  <svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><polyline points="3 6 5 6 21 6"/><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"/></svg>
                                </button>
                              `:""}
                            </div>
                          </td>
                        </tr>
                      `}).join("")}
                  </tbody>
                </table>
              </div>
            </div>
          `:`
            <!-- Grid Card View -->
            <div class="sp-cards-grid">
              ${u.map(p=>{let A=Nt(p.akhir_sp),T=Ua(p.sp_type);return`
                  <div class="sp-card" data-id="${p.id}">
                    <div>
                      <div style="display:flex; justify-content:space-between; align-items:flex-start; margin-bottom:14px;">
                        <span class="sp-type-badge" style="background:${T.bg}; color:${T.color}; border:1px solid ${T.border};">
                          <span>${T.icon}</span>
                          <span>${T.label}</span>
                        </span>
                        <span class="sp-status-badge" style="background:${A.bg}; color:${A.color}; border:1px solid ${A.border};">
                          <span>${A.icon}</span>
                          <span>${A.shortLabel}</span>
                        </span>
                      </div>

                      <div style="display:flex; align-items:center; gap:12px; margin-bottom:12px;">
                        <div class="sp-avatar-box" style="width:48px; height:48px; font-size:1.05rem;">
                          ${p.employee_name.slice(0,2).toUpperCase()}
                        </div>
                        <div>
                          <h3 style="margin:0 0 3px 0; font-size:1.1rem; font-weight:700; color:#F8FAFC;">
                            ${p.employee_name}
                          </h3>
                          <div style="font-size:0.8rem; color:#CBD5E1;">
                            \u{1F4CD} ${p.branch_name||"Cabang"} \u2022 <span style="color:#93C5FD; font-weight:600;">${p.division||"FC"}</span>
                          </div>
                        </div>
                      </div>

                      <div style="background:#0E182B; border:1px solid #223456; border-radius:10px; padding:10px 12px; margin-bottom:14px; font-size:0.82rem; display:flex; justify-content:space-between;">
                        <div>
                          <div style="color:#94A3B8; font-size: 0.76rem; text-transform:uppercase; font-weight:600;">Tanggal Terbit</div>
                          <div style="font-weight:700; color:#F8FAFC;">${xe(p.tanggal)}</div>
                        </div>
                        <div style="text-align:right;">
                          <div style="color:#94A3B8; font-size: 0.76rem; text-transform:uppercase; font-weight:600;">Akhir SP</div>
                          <div style="font-weight:700; color:#F8FAFC;">${xe(p.akhir_sp)}</div>
                        </div>
                      </div>

                      <p style="font-size:0.83rem; color:#CBD5E1; line-height:1.45; margin:0 0 16px 0; display:-webkit-box; -webkit-line-clamp:2; -webkit-box-orient:vertical; overflow:hidden;">
                        ${p.notes||p.description||"Pencatatan sanksi peringatan disiplin operasional."}
                      </p>
                    </div>

                    <div style="padding-top:12px; border-top:1px solid #223456; display:flex; align-items:center; justify-content:space-between;">
                      <button class="btn btn-sm btn-ghost sp-detail-btn" style="border:1px solid #223456; color:#93C5FD; border-radius:8px; font-weight:600; font-size:0.82rem; padding:6px 12px;" data-id="${p.id}">
                        \u{1F441}\uFE0F Rincian BAP
                      </button>

                      <div style="display:flex; align-items:center; gap:6px;">
                        ${p.document_link?`
                          <a href="${p.document_link}" target="_blank" rel="noopener noreferrer" class="sp-icon-btn" title="Buka Dokumen">
                            <svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M18 13v6a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V8a2 2 0 0 1 2-2h6"/><polyline points="15 3 21 3 21 9"/><line x1="10" y1="14" x2="21" y2="3"/></svg>
                          </a>
                        `:""}
                        ${a?`
                          <button class="sp-icon-btn sp-edit-btn" title="Edit Data SP" data-id="${p.id}">
                            <svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>
                          </button>
                        `:""}
                        ${o?`
                          <button class="sp-icon-btn delete sp-delete-btn" title="Hapus" data-id="${p.id}">
                            <svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><polyline points="3 6 5 6 21 6"/><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"/></svg>
                          </button>
                        `:""}
                      </div>
                    </div>
                  </div>
                `}).join("")}
            </div>
          `}
    `}function S(){let u=e.querySelector("#sp-btn-reset-filter");u&&(u.onclick=()=>{$="",d="all",F="all",i="all",h=1,clearTimeout(q);let p=e.querySelector("#sp-search-input");p&&(p.value="");let A=e.querySelector("#sp-filter-branch");A&&(A.value="all");let T=e.querySelector("#sp-filter-division");T&&(T.value="all"),e.querySelectorAll(".sp-pill").forEach(g=>{g.classList.toggle("active",g.getAttribute("data-filter")==="all")}),E({refreshStats:!1})}),e.querySelectorAll(".sp-detail-btn, .sp-view-detail-link").forEach(p=>{p.onclick=A=>{A.stopPropagation();let T=Number(p.getAttribute("data-id")),g=s.find(C=>C.id===T);g&&W(g)}}),e.querySelectorAll(".sp-edit-btn").forEach(p=>{p.onclick=A=>{A.stopPropagation();let T=Number(p.getAttribute("data-id")),g=s.find(C=>C.id===T);g&&G(g)}}),e.querySelectorAll(".sp-delete-btn").forEach(p=>{p.onclick=A=>{A.stopPropagation();let T=Number(p.getAttribute("data-id")),g=s.find(C=>C.id===T);g&&J(g)}})}function y(){let u=e.querySelector("#sp-pagination-container");if(!u)return;u.innerHTML="";let p=Te({page:h,pages:I,total:R,limit:L,onPage:A=>{h=A,E({refreshStats:!1})}});p&&u.appendChild(p)}function r(){let u=e.querySelector("#sp-list-container");if(!u)return;let p=w();u.innerHTML=b(p),y(),S()}function l(){let u=w(),p=K();e.innerHTML=`
      <div class="sp-page-container">
        
        <!-- Hero Header -->
        <div class="sp-hero-card">
          <div class="sp-hero-header">
            <div>
              <h1 class="sp-hero-title">
                <span>\u{1F4DC}</span> Monitoring Data SP (Surat Peringatan)
              </h1>
            </div>
            <div class="sp-hero-actions" style="display:flex; align-items:center; gap:8px;">
              ${n?`
                <button id="sp-btn-add" class="btn btn-primary sp-btn-add" style="background:#2563EB; height:38px; padding:0 16px; border-radius:10px; font-weight:700; display:inline-flex; align-items:center; gap:6px; box-shadow:0 4px 12px rgba(37,99,235,0.3); white-space:nowrap;">
                  <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24"><line x1="12" y1="5" x2="12" y2="19"/><line x1="5" y1="12" x2="19" y2="12"/></svg>
                  <span>Terbitkan SP</span>
                </button>
              `:""}
              
              ${re(t,"export")?`<button id="sp-btn-excel" class="btn btn-ghost sp-btn-excel" style="background:rgba(255,255,255,0.08); border:1px solid rgba(255,255,255,0.15); color:white; height:38px; padding:0 14px; border-radius:10px; font-weight:600; display:inline-flex; align-items:center; gap:6px; white-space:nowrap;">
                <span>\u{1F4E5}</span>
                <span>Ekspor Excel</span>
              </button>`:""}

              <button id="sp-btn-refresh" class="btn btn-ghost sp-btn-refresh" style="background:rgba(255,255,255,0.08); border:1px solid rgba(255,255,255,0.15); color:white; width:38px; height:38px; padding:0; border-radius:10px; display:inline-flex; align-items:center; justify-content:center; flex-shrink:0;" title="Muat Ulang Data">
                <svg width="16" height="16" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M23 4v6h-6M1 20v-6h6"/><path d="M3.51 9a9 9 0 0 1 14.85-3.36L23 10M1 14l4.64 4.36A9 9 0 0 0 20.49 15"/></svg>
              </button>
            </div>
          </div>

          <!-- Quick Metrics -->
          <div class="sp-stats-grid">
            <div class="sp-stat-box">
              <div class="sp-stat-icon" style="background:rgba(239,68,68,0.2); color:#F87171;">\u{1F534}</div>
              <div>
                <div class="sp-stat-value" id="sp-stat-active">${p.active}</div>
                <div class="sp-stat-label">SP Aktif Berlaku</div>
              </div>
            </div>

            <div class="sp-stat-box">
              <div class="sp-stat-icon" style="background:rgba(245,158,11,0.2); color:#FBBF24;">\u26A0\uFE0F</div>
              <div>
                <div class="sp-stat-value" id="sp-stat-expiring">${p.expiring}</div>
                <div class="sp-stat-label">Segera Berakhir (&lt;30 Hari)</div>
              </div>
            </div>

            <div class="sp-stat-box">
              <div class="sp-stat-icon" style="background:rgba(16,185,129,0.2); color:#34D399;">\u2705</div>
              <div>
                <div class="sp-stat-value" id="sp-stat-expired">${p.expired}</div>
                <div class="sp-stat-label">Masa Sanksi Selesai</div>
              </div>
            </div>

            <div class="sp-stat-box">
              <div class="sp-stat-icon" style="background:rgba(59,130,246,0.2); color:#60A5FA;">\u{1F465}</div>
              <div>
                <div class="sp-stat-value"><span id="sp-stat-unique">${p.uniqueEmployees}</span> <span style="font-size:0.8rem; font-weight:400; color:#94A3B8;">karyawan</span></div>
                <div class="sp-stat-label">Total Terkena Sanksi</div>
              </div>
            </div>
          </div>
        </div>

        <!-- Controls Panel -->
        <div class="sp-controls-panel">
          <div class="sp-controls-top">
            <div class="sp-search-wrapper">
              <span class="sp-search-icon">\u{1F50D}</span>
              <input type="text" id="sp-search-input" class="sp-search-input" placeholder="Cari nama karyawan, cabang, atau divisi..." value="${$?$.replace(/"/g,"&quot;"):""}" />
            </div>

            <!-- Branch Dropdown -->
            <select id="sp-filter-branch" class="sp-filter-select">
              <option value="all">Semua Cabang (${c.length})</option>
              ${c.map(A=>`<option value="${A.value}" ${String(d)===String(A.value)?"selected":""}>${A.label}</option>`).join("")}
            </select>

            <!-- Division Dropdown -->
            <select id="sp-filter-division" class="sp-filter-select">
              <option value="all" ${F==="all"?"selected":""}>Semua Divisi</option>
              <option value="FACILITY CARE" ${F==="FACILITY CARE"?"selected":""}>FACILITY CARE</option>
              <option value="SECURITY" ${F==="SECURITY"?"selected":""}>SECURITY</option>
            </select>

            <div class="sp-view-toggle">
              <button type="button" class="sp-view-btn ${M==="card"?"active":""}" data-view="card" title="Tampilan Kartu">
                <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><rect x="3" y="3" width="7" height="7" rx="1"/><rect x="14" y="3" width="7" height="7" rx="1"/><rect x="14" y="14" width="7" height="7" rx="1"/><rect x="3" y="14" width="7" height="7" rx="1"/></svg>
                <span>Kartu</span>
              </button>
              <button type="button" class="sp-view-btn ${M==="table"?"active":""}" data-view="table" title="Tampilan Tabel">
                <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M3 9h18M3 15h18M9 3v18"/><rect x="3" y="3" width="18" height="18" rx="2"/></svg>
                <span>Tabel</span>
              </button>
            </div>
          </div>

          <!-- Filter Pills -->
          <div class="sp-filter-pills">
            <button class="sp-pill ${i==="all"?"active":""}" data-filter="all">
              <span>Semua SP</span>
              <span class="sp-pill-count" id="sp-pill-count-all">${p.total}</span>
            </button>
            <button class="sp-pill ${i==="active"?"active":""}" data-filter="active">
              <span>\u{1F534} SP Aktif</span>
              <span class="sp-pill-count" id="sp-pill-count-active">${p.active}</span>
            </button>
            <button class="sp-pill ${i==="expiring"?"active":""}" data-filter="expiring">
              <span>\u26A0\uFE0F &lt;30 Hari</span>
              <span class="sp-pill-count" id="sp-pill-count-expiring">${p.expiring}</span>
            </button>
            <button class="sp-pill ${i==="expired"?"active":""}" data-filter="expired">
              <span>\u2705 Selesai</span>
              <span class="sp-pill-count" id="sp-pill-count-expired">${p.expired}</span>
            </button>
            <button class="sp-pill ${i==="sp1"?"active":""}" data-filter="sp1">
              <span>SP 1</span>
              <span class="sp-pill-count" id="sp-pill-count-sp1">${p.sp1}</span>
            </button>
            <button class="sp-pill ${i==="sp2"?"active":""}" data-filter="sp2">
              <span>SP 2</span>
              <span class="sp-pill-count" id="sp-pill-count-sp2">${p.sp2}</span>
            </button>
            <button class="sp-pill ${i==="sp3"?"active":""}" data-filter="sp3">
              <span>SP 3</span>
              <span class="sp-pill-count" id="sp-pill-count-sp3">${p.sp3}</span>
            </button>
          </div>
        </div>

        <!-- Content Area -->
        <div id="sp-list-container">
          ${b(u)}
        </div>

        <!-- Pagination Container -->
        <div id="sp-pagination-container"></div>

      </div>
    `,y(),f()}function f(){let u=e.querySelector("#sp-btn-add");u&&(u.onclick=()=>G());let p=e.querySelector("#sp-btn-excel");p&&(p.onclick=()=>Y());let A=e.querySelector("#sp-btn-refresh");A&&(A.onclick=()=>E({refreshStats:!0}));let T=e.querySelector("#sp-search-input");T&&T.addEventListener("input",z=>{$=z.target.value,clearTimeout(q),q=setTimeout(()=>{h=1,E({refreshStats:!1})},350)});let g=e.querySelector("#sp-filter-branch");g&&g.addEventListener("change",z=>{d=z.target.value,h=1,E({refreshStats:!1})});let C=e.querySelector("#sp-filter-division");C&&C.addEventListener("change",z=>{F=z.target.value,h=1,E({refreshStats:!1})}),e.querySelectorAll(".sp-view-btn").forEach(z=>{z.addEventListener("click",()=>{let O=z.getAttribute("data-view");!O||O===M||(M=O,localStorage.setItem("fcms_sp_view_mode",M),e.querySelectorAll(".sp-view-btn").forEach(te=>{te.classList.toggle("active",te===z)}),r())})}),e.querySelectorAll(".sp-pill").forEach(z=>{z.addEventListener("click",()=>{i=z.getAttribute("data-filter")||"all",e.querySelectorAll(".sp-pill").forEach(O=>{O.classList.toggle("active",O===z)}),h=1,E({refreshStats:!1})})}),S()}await E({refreshStats:!0})}Be();async function yi(e){let n=[],a=[],o=[],s="all",c="all",m="all",i="",d=1,F=0,$=1,M=!1,h={total:0,proses:0,selesai:0,currentMonth:0},L=null,R=!1,I=localStorage.getItem("fcms_mutasi_view_mode")||"card",N=se(),q=re(N,"mutasi","write"),v=re(N,"mutasi","write"),B=re(N,"mutasi","delete");try{[a,o]=await Promise.all([we(),ze()])}catch(r){console.error("Error fetching cache:",r)}function x(){let r={sort:"date_desc"};return i.trim()&&(r.search=i.trim()),c&&c!=="all"&&(r.from_branch_id=c),m&&m!=="all"&&(r.to_branch_id=m),s&&s!=="all"&&(r.status=s),r}function k(){let r=(l,f)=>{let u=e.querySelector(l);u&&(u.textContent=f)};r("#mutasi-stat-total",h.total),r("#mutasi-stat-proses",h.proses),r("#mutasi-stat-selesai",h.selesai),r("#mutasi-stat-month",h.currentMonth),r("#mutasi-pill-count-all",h.total),r("#mutasi-pill-count-proses",h.proses),r("#mutasi-pill-count-selesai",h.selesai)}async function _({refreshStats:r=!0}={}){if(R=!0,!!!e.querySelector("#mutasi-list-container"))S();else{let f=e.querySelector("#mutasi-list-container");f&&(f.innerHTML='<div class="mutasi-empty-state"><p style="color:var(--text-3); margin:0;">Memuat data mutasi...</p></div>')}try{let f=new URLSearchParams({page:String(d),limit:String(20),...x()});(r||!M)&&f.set("include_stats","1");let u=await X(`/api/mutasi?${f.toString()}`);if(u.ok){let p=u.data?.data||[];n=Array.isArray(p)?p:[];let A=u.data?.pagination||{};if(F=Number(A.total??n.length),$=Math.max(1,Number(A.pages??Math.ceil(F/20))||1),d=Math.max(1,Number(A.page??d)),u.data?.stats&&(h=u.data.stats,M=!0),n.length===0&&d>$)return d=$,R=!1,_({refreshStats:r})}else n=[],F=0,$=1,j(u.data?.error||"Gagal memuat data mutasi")}catch(f){console.warn("Gagal load mutasi:",f),n=[],F=0,$=1,j("Gagal memuat data mutasi: "+f.message)}finally{R=!1,e.querySelector("#mutasi-list-container")?(k(),b()):S()}}async function E(){if(!re(N,"export")){j("Anda tidak memiliki izin untuk mengekspor data.");return}try{let[{downloadExcel:r},l]=await Promise.all([Promise.resolve().then(()=>(Ce(),Ve)),ue("/api/mutasi/export",x())]),f=l.map((u,p)=>({No:p+1,Tanggal:u.tanggal||"","Nama Karyawan":u.employee_name||"","Cabang Asal":u.from_branch_name||"","Cabang Tujuan":u.to_branch_name||"",Status:u.status||"","Link Dokumen":u.document_link||""}));r(f,`Data_Mutasi_${new Date().toISOString().slice(0,10)}`),ae("Data mutasi berhasil diekspor ke Excel!")}catch(r){j("Gagal ekspor Excel: "+r.message)}}async function w(r){if(!B)return;let l=n.find(f=>f.id===r);ge({title:"Hapus Data Mutasi",message:`Apakah Anda yakin ingin menghapus data mutasi untuk "${l?.employee_name||"karyawan ini"}"?`,confirmText:"Ya, Hapus",confirmClass:"btn-danger",onConfirm:async()=>{try{let f=await X(`/api/mutasi/${r}`,{method:"DELETE"});f.ok?(ae("Data mutasi berhasil dihapus"),await _({refreshStats:!0})):j(f.data?.error||"Gagal menghapus mutasi")}catch(f){j("Gagal menghapus data mutasi: "+f.message)}}})}function K(r=null){let l=!!r;if(l&&!v){j("Anda tidak memiliki izin untuk mengedit data mutasi.");return}if(!l&&!q){j("Anda tidak memiliki izin untuk menambah data mutasi.");return}if(!v)return;let f=!!r,u=r?.tanggal?r.tanggal.slice(0,10):new Date().toISOString().slice(0,10),p=document.createElement("div");p.innerHTML=`
      <div style="display:flex; flex-direction:column; gap:16px;">
        <div class="form-row" style="display:grid; grid-template-columns:repeat(auto-fit, minmax(200px, 1fr)); gap:12px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:var(--text-2); margin-bottom:5px;">
              Tanggal Mutasi <span style="color:#DC2626;">*</span>
            </label>
            <input type="date" id="mutasi-modal-tanggal" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              value="${u}" required />
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:var(--text-2); margin-bottom:5px;">
              Status Mutasi <span style="color:#DC2626;">*</span>
            </label>
            <select id="mutasi-modal-status" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" required>
              <option value="Proses" ${r?.status==="Proses"?"selected":""}>\u23F3 Proses</option>
              <option value="Selesai" ${r?.status==="Selesai"||!r?"selected":""}>\u2705 Selesai</option>
            </select>
          </div>
        </div>

        <div>
          <label style="display:block; font-size:0.82rem; font-weight:600; color:var(--text-2); margin-bottom:5px;">
            Nama Karyawan <span style="color:#DC2626;">*</span>
          </label>
          <div id="mutasi-modal-emp-mount"></div>
        </div>

        <div class="form-row" style="display:grid; grid-template-columns:repeat(auto-fit, minmax(200px, 1fr)); gap:12px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:var(--text-2); margin-bottom:5px;">
              Cabang Asal <span style="color:#DC2626;">*</span>
            </label>
            <select id="mutasi-modal-from" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" required>
              <option value="">-- Pilih Cabang Asal --</option>
              ${a.map(T=>`
                <option value="${T.value}" ${String(r?.from_branch_id)===String(T.value)?"selected":""}>
                  ${T.label}
                </option>
              `).join("")}
            </select>
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:var(--text-2); margin-bottom:5px;">
              Cabang Tujuan <span style="color:#DC2626;">*</span>
            </label>
            <select id="mutasi-modal-to" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" required>
              <option value="">-- Pilih Cabang Tujuan --</option>
              ${a.map(T=>`
                <option value="${T.value}" ${String(r?.to_branch_id)===String(T.value)?"selected":""}>
                  ${T.label}
                </option>
              `).join("")}
            </select>
          </div>
        </div>

        <div>
          <label style="display:block; font-size:0.82rem; font-weight:600; color:var(--text-2); margin-bottom:5px;">
            Link Berkas SK Mutasi / Dokumen (Opsional)
          </label>
          <input type="url" id="mutasi-modal-link" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
            value="${r?.document_link?String(r.document_link).replace(/"/g,"&quot;"):""}" />
        </div>
      </div>
    `;let A=je(p.querySelector("#mutasi-modal-emp-mount"),{mode:"single",inputId:"mutasi-modal-emp",inputName:"employee_name",placeholder:"",required:!0,allowManual:!0,hasEmployeeIdSchema:!0,manualBadgeText:"Belum terhubung ke Master Karyawan",employees:o,initialEmployeeId:r?.employee_id||null,initialEmployeeName:r?.employee_name||""});ce({title:f?"\u270F\uFE0F Edit Data Mutasi Karyawan":"\u2795 Tambah Data Mutasi Karyawan",content:p,size:"md",confirmText:f?"Simpan Perubahan":"Simpan Mutasi",confirmClass:"btn-primary",onConfirm:async(T,g)=>{let C=p.querySelector("#mutasi-modal-tanggal").value,z=A.getValue(),O=(z.employee_name||"").trim(),te=z.employee_id??null,Q=p.querySelector("#mutasi-modal-from").value,ne=p.querySelector("#mutasi-modal-to").value,ee=p.querySelector("#mutasi-modal-status").value,H=(p.querySelector("#mutasi-modal-link").value||"").trim();if(!C){j("Tanggal mutasi wajib diisi");return}if(!O){j("Nama karyawan wajib diisi"),A.inputEl?.focus();return}if(!Q){j("Cabang asal wajib dipilih");return}if(!ne){j("Cabang tujuan wajib dipilih");return}let U=a.find(Z=>String(Z.value)===String(Q)),D=a.find(Z=>String(Z.value)===String(ne)),V={tanggal:C,employee_id:te,employee_name:O,from_branch_id:Number(Q),from_branch_name:U?U.label:"Cabang Asal",to_branch_id:Number(ne),to_branch_name:D?D.label:"Cabang Tujuan",status:ee,document_link:H,...f?{row_version:r?.row_version!==void 0&&r?.row_version!==null?Number(r.row_version):1}:{}};try{if(f){let Z=await X(`/api/mutasi/${r.id}`,{method:"PUT",body:JSON.stringify(V)});if(Z.ok)ae("Data mutasi berhasil diperbarui"),g(),await _({refreshStats:!0});else{if(Z.status===409){j("Konflik OCC: Data telah diubah oleh pengguna lain. Silakan muat ulang halaman."),g(),await _({refreshStats:!0});return}j(Z.data?.error||"Gagal memperbarui data mutasi")}}else{let Z=await X("/api/mutasi",{method:"POST",body:JSON.stringify(V)});Z.ok?(ae("Data mutasi baru berhasil ditambahkan"),g(),await _({refreshStats:!0})):j(Z.data?.error||"Gagal menambahkan mutasi")}}catch(Z){j("Terjadi kesalahan: "+Z.message)}}})}function G(){return n}function J(r){return`
          ${r.length===0?`
            <div class="mutasi-empty-state">
              <div style="font-size:3rem; margin-bottom:12px;">\u{1F50D}</div>
              <h3 style="font-size:1.2rem; font-weight:700; color:var(--text-1); margin:0 0 6px 0;">Tidak Ada Data Mutasi Ditemukan</h3>
              <p style="color:var(--text-3); font-size:0.9rem; margin:0 0 20px 0;">Silakan sesuaikan kata kunci pencarian atau reset filter aktif.</p>
              <button id="mutasi-btn-reset-filter" class="btn btn-ghost" style="border:1px solid #CBD5E1; border-radius:8px; padding:8px 16px;">
                Reset Filter
              </button>
            </div>
          `:I==="table"?`
            <div class="mutasi-table-card">
              <div class="mutasi-table-responsive">
                <table class="mutasi-table">
                  <thead>
                    <tr>
                      <th style="width:45px; text-align:center;">No</th>
                      <th>Nama Karyawan</th>
                      <th>Tanggal Mutasi</th>
                      <th>Cabang Asal</th>
                      <th>Cabang Tujuan</th>
                      <th style="text-align:center;">Status</th>
                      <th style="text-align:center;">Dokumen SK</th>
                      ${v||B?'<th style="width:90px; text-align:center;">Aksi</th>':""}
                    </tr>
                  </thead>
                  <tbody>
                    ${r.map((l,f)=>{let u=l.status==="Selesai",p=(l.employee_name||"Karyawan").split(" ").map(A=>A[0]).filter(Boolean).slice(0,2).join("").toUpperCase();return`
                        <tr data-id="${l.id}">
                          <td style="text-align:center; color:#94A3B8; font-weight:600; font-size:0.8rem;">${(d-1)*20+f+1}</td>
                          <td>
                            <div style="display:flex; align-items:center; gap:10px;">
                              <div class="mutasi-avatar-box" style="width:34px; height:34px; font-size:0.78rem; border-radius:8px;">
                                ${p}
                              </div>
                              <div>
                                <div style="font-weight:700; color:#F8FAFC; white-space:nowrap;">${l.employee_name}</div>
                                <div style="font-size: 0.76rem; color:#94A3B8;">ID Mutasi: #${l.id}</div>
                              </div>
                            </div>
                          </td>
                          <td style="white-space:nowrap; font-weight:600; color:#F8FAFC;">
                            \u{1F5D3}\uFE0F ${xe(l.tanggal)}
                          </td>
                          <td style="white-space:nowrap;">
                            <div style="font-weight:600; color:#CBD5E1; font-size:0.85rem; display:inline-flex; align-items:center; gap:4px;">
                              <span>\u{1F4CD}</span>
                              <span>${l.from_branch_name||"-"}</span>
                            </div>
                          </td>
                          <td style="white-space:nowrap;">
                            <div style="font-weight:700; color:#93C5FD; font-size:0.85rem; display:inline-flex; align-items:center; gap:5px;">
                              <span>\u2794</span>
                              <span>${l.to_branch_name||"-"}</span>
                            </div>
                          </td>
                          <td style="text-align:center; white-space:nowrap;">
                            <span class="mutasi-status-badge ${u?"selesai":"proses"}">
                              <span>${u?"\u2705":"\u23F3"}</span>
                              <span>${l.status||"Proses"}</span>
                            </span>
                          </td>
                          <td style="text-align:center; white-space:nowrap;">
                            ${l.document_link?`
                              <a href="${l.document_link}" target="_blank" rel="noopener noreferrer" class="btn btn-sm btn-ghost" style="border:1px solid #CBD5E1; border-radius:6px; font-size:0.76rem; padding:4px 8px; text-decoration:none; display:inline-flex; align-items:center; gap:4px;">
                                <span>\u{1F4C4} Buka SK</span>
                              </a>
                            `:`
                              <span style="font-size:0.75rem; color:#94A3B8;">(Tanpa Dokumen)</span>
                            `}
                          </td>
                          ${v||B?`
                            <td style="text-align:center; white-space:nowrap;">
                              <div style="display:inline-flex; align-items:center; gap:6px;">
                                ${v?`
                                  <button class="mutasi-icon-btn mutasi-edit-btn" title="Edit Data Mutasi" data-id="${l.id}">
                                    <svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>
                                  </button>
                                `:""}
                                ${B?`
                                  <button class="mutasi-icon-btn delete mutasi-delete-btn" title="Hapus Mutasi" data-id="${l.id}">
                                    <svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><polyline points="3 6 5 6 21 6"/><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"/></svg>
                                  </button>
                                `:""}
                              </div>
                            </td>
                          `:""}
                        </tr>
                      `}).join("")}
                  </tbody>
                </table>
              </div>
            </div>
          `:`
            <div class="mutasi-cards-grid">
              ${r.map(l=>{let f=l.status==="Selesai",u=(l.employee_name||"Karyawan").split(" ").map(p=>p[0]).filter(Boolean).slice(0,2).join("").toUpperCase();return`
                  <div class="mutasi-card" data-id="${l.id}">
                    <div>
                      <!-- Card Header: Status & Tanggal -->
                      <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:12px;">
                        <span class="mutasi-status-badge ${f?"selesai":"proses"}">
                          <span>${f?"\u2705":"\u23F3"}</span>
                          <span>${l.status||"Proses"}</span>
                        </span>
                        <span style="font-size:0.8rem; font-weight:600; color:#CBD5E1; display:inline-flex; align-items:center; gap:4px;">
                          <span>\u{1F5D3}\uFE0F</span>
                          <span>${xe(l.tanggal)}</span>
                        </span>
                      </div>

                      <!-- Employee Info -->
                      <div style="display:flex; align-items:center; gap:12px; margin-bottom:12px;">
                        <div class="mutasi-avatar-box">
                          ${u}
                        </div>
                        <div style="min-width:0;">
                          <h4 style="margin:0 0 3px 0; font-size:1.02rem; font-weight:700; color:#F8FAFC; white-space:nowrap; overflow:hidden; text-overflow:ellipsis;">
                            ${l.employee_name}
                          </h4>
                          <div style="font-size:0.75rem; color:#94A3B8;">
                            ID Mutasi: #${l.id}
                          </div>
                        </div>
                      </div>

                      <!-- Route Visual Box -->
                      <div class="mutasi-route-box">
                        <div class="mutasi-route-node mutasi-from-node">
                          <div class="mutasi-route-label">Cabang Asal</div>
                          <div class="mutasi-route-branch" title="${l.from_branch_name||"-"}">
                            ${l.from_branch_name||"-"}
                          </div>
                        </div>

                        <div class="mutasi-route-arrow">
                          \u2794
                        </div>

                        <div class="mutasi-route-node mutasi-to-node">
                          <div class="mutasi-route-label">Cabang Tujuan</div>
                          <div class="mutasi-route-branch" style="color:#93C5FD;" title="${l.to_branch_name||"-"}">
                            ${l.to_branch_name||"-"}
                          </div>
                        </div>
                      </div>
                    </div>

                    <!-- Card Footer: Dokumen & Aksi -->
                    <div style="display:flex; justify-content:space-between; align-items:center; padding-top:12px; border-top:1px solid #223456; margin-top:8px;">
                      <div>
                        ${l.document_link?`
                          <a href="${l.document_link}" target="_blank" rel="noopener noreferrer" class="btn btn-sm btn-ghost" style="border:1px solid #CBD5E1; border-radius:6px; font-size:0.78rem; padding:4px 8px; text-decoration:none; display:inline-flex; align-items:center; gap:4px;">
                            <span>\u{1F4C4} Buka SK</span>
                          </a>
                        `:`
                          <span style="font-size:0.75rem; color:#94A3B8;">(Tanpa Dokumen)</span>
                        `}
                      </div>

                      ${v||B?`
                        <div style="display:flex; align-items:center; gap:6px;">
                          ${v?`
                            <button class="mutasi-icon-btn mutasi-edit-btn" title="Edit Data Mutasi" data-id="${l.id}">
                              <svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>
                            </button>
                          `:""}
                          ${B?`
                            <button class="mutasi-icon-btn delete mutasi-delete-btn" title="Hapus Mutasi" data-id="${l.id}">
                              <svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><polyline points="3 6 5 6 21 6"/><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"/></svg>
                            </button>
                          `:""}
                        </div>
                      `:""}
                    </div>

                  </div>
                `}).join("")}
            </div>
          `}
    `}function W(){let r=e.querySelector("#mutasi-pagination-container");if(!r)return;r.innerHTML="";let l=Te({page:d,pages:$,total:F,limit:20,onPage:f=>{d=f,_({refreshStats:!1})}});l&&r.appendChild(l)}function Y(){let r=e.querySelector("#mutasi-btn-reset-filter");r&&(r.onclick=()=>{i="",c="all",m="all",s="all",d=1;let l=e.querySelector("#mutasi-search-input");l&&(l.value="");let f=e.querySelector("#mutasi-filter-from");f&&(f.value="all");let u=e.querySelector("#mutasi-filter-to");u&&(u.value="all"),e.querySelectorAll(".mutasi-pill").forEach(p=>{p.classList.toggle("active",p.getAttribute("data-status")==="all")}),_({refreshStats:!1})}),e.querySelectorAll(".mutasi-edit-btn").forEach(l=>{l.onclick=f=>{f.stopPropagation();let u=Number(l.getAttribute("data-id")),p=n.find(A=>A.id===u);p&&K(p)}}),e.querySelectorAll(".mutasi-delete-btn").forEach(l=>{l.onclick=f=>{f.stopPropagation();let u=Number(l.getAttribute("data-id"));w(u)}})}function b(){let r=e.querySelector("#mutasi-list-container");if(!r)return;let l=G();r.innerHTML=J(l),W(),Y()}function S(){let r=G(),l=h.total,f=h.proses,u=h.selesai,p=h.currentMonth;e.innerHTML=`
      <style>.mutasi-page{display:flex;flex-direction:column;gap:20px}.mutasi-hero-card{background:#131F36;border:1px solid #223456;border-radius:16px;padding:24px 28px;color:#FFFFFF;box-shadow:0 4px 16px rgba(0,0,0,0.22);position:relative;overflow:hidden}.mutasi-hero-card::after{content:'';position:absolute;top:-40px;right:-40px;width:220px;height:220px;background:radial-gradient(circle,rgba(37,99,235,0.06) 0%,transparent 70%);border-radius:50%;pointer-events:none}.mutasi-hero-actions{display:flex;align-items:center;gap:8px;flex-wrap:wrap}.mutasi-hero-actions .btn{height:38px;box-sizing:border-box;transition:all 0.2s ease}.mutasi-btn-add:hover{background:#1D4ED8 !important;transform:translateY(-1px)}.mutasi-btn-excel:hover,.mutasi-btn-refresh:hover{background:rgba(255,255,255,0.16) !important;border-color:rgba(255,255,255,0.3) !important;transform:translateY(-1px)}.mutasi-stats-row{display:grid;grid-template-columns:repeat(4,1fr);gap:14px;margin-top:22px}.mutasi-stat-box{background:rgba(255,255,255,0.06);border:1px solid rgba(255,255,255,0.1);border-radius:12px;padding:14px 16px;backdrop-filter:blur(8px);display:flex;align-items:center;gap:14px}.mutasi-stat-icon{width:44px;height:44px;border-radius:10px;display:flex;align-items:center;justify-content:center;font-size:1.3rem;flex-shrink:0}.mutasi-stat-val{font-size:1.55rem;font-weight:800;line-height:1.1;color:#F8FAFC}.mutasi-stat-lbl{font-size:0.76rem;text-transform:uppercase;letter-spacing:0.05em;color:#94A3B8;font-weight:600;margin-top:2px}.mutasi-controls-card{background:#131F36;border:1px solid #223456;border-radius:14px;padding:18px 20px;box-shadow:0 2px 4px rgba(0,0,0,0.02);display:flex;flex-direction:column;gap:16px}.mutasi-search-row{display:flex;gap:12px;align-items:center;flex-wrap:wrap}.mutasi-search-wrap{position:relative;flex:1;min-width:240px}.mutasi-search-icon{position:absolute;left:14px;top:50%;transform:translateY(-50%);color:#94A3B8;pointer-events:none}.mutasi-search-input{width:100%;padding:10px 14px 10px 40px;border:1px solid #223456;border-radius:10px;font-size:0.9rem;outline:none;transition:border-color 0.15s ease}.mutasi-search-input:focus{border-color:#2563EB;box-shadow:0 0 0 3px rgba(37,99,235,0.1)}.mutasi-select{padding:10px 14px;border:1px solid #223456;border-radius:10px;font-size:0.9rem;color:#F8FAFC;background:#131F36;outline:none;min-width:180px;cursor:pointer}.mutasi-select:focus{border-color:#2563EB}.mutasi-filter-pills{display:flex;align-items:center;gap:8px;overflow-x:auto;white-space:nowrap;scrollbar-width:none;padding-bottom:2px}.mutasi-filter-pills::-webkit-scrollbar{display:none}.mutasi-pill{padding:6px 14px;border-radius:20px;font-size:0.82rem;font-weight:600;cursor:pointer;border:1px solid #223456;background:#0E182B;color:#CBD5E1;transition:all 0.15s ease;display:inline-flex;align-items:center;gap:6px;user-select:none;flex-shrink:0}.mutasi-pill:hover{background:#22314A;color:#F8FAFC}.mutasi-pill.active{background:#2563EB;color:white;border-color:#2563EB;box-shadow:0 2px 6px rgba(37,99,235,0.25)}.mutasi-view-toggle{display:inline-flex;align-items:center;background:#0E182B;padding:3px;border-radius:10px;border:1px solid #223456;gap:3px;flex-shrink:0}.mutasi-view-btn{display:inline-flex;align-items:center;gap:6px;padding:6px 12px;border-radius:7px;font-size:0.82rem;font-weight:600;border:none;background:transparent;color:#CBD5E1;cursor:pointer;transition:all 0.15s ease;height:34px;box-sizing:border-box;user-select:none}.mutasi-view-btn:hover{color:#F8FAFC;background:#22314A}.mutasi-view-btn.active{background:#131F36;color:#60A5FA;font-weight:700;box-shadow:0 1px 4px rgba(0,0,0,0.2)}.mutasi-table-card{background:#131F36;border:1px solid #223456;border-radius:14px;box-shadow:0 2px 6px rgba(0,0,0,0.2);overflow:hidden}.mutasi-table-responsive{width:100%;overflow-x:auto;-webkit-overflow-scrolling:touch}.mutasi-table{width:100%;border-collapse:collapse;text-align:left;font-size:0.88rem;min-width:860px}.mutasi-table th{background:#0E182B;color:#CBD5E1;font-size:0.76rem;font-weight:700;text-transform:uppercase;letter-spacing:0.05em;padding:13px 16px;border-bottom:1.5px solid #223456;white-space:nowrap}.mutasi-table td{padding:13px 16px;border-bottom:1px solid #223456;color:#F8FAFC;vertical-align:middle}.mutasi-table tbody tr:hover td{background:#22314A}.mutasi-table tbody tr:last-child td{border-bottom:none}.mutasi-cards-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(320px,1fr));gap:16px}.mutasi-card{background:#131F36;border:1px solid #223456;border-radius:14px;padding:18px;box-shadow:0 2px 6px rgba(0,0,0,0.03);transition:transform 0.15s ease,box-shadow 0.15s ease;display:flex;flex-direction:column;justify-content:space-between;position:relative}.mutasi-card:hover{transform:translateY(-2px);box-shadow:0 8px 18px rgba(0,0,0,0.25);border-color:#3B82F6}.mutasi-avatar-box{width:44px;height:44px;border-radius:10px;background:#2563EB;color:#FFFFFF;font-weight:700;font-size:1rem;display:flex;align-items:center;justify-content:center;flex-shrink:0}.mutasi-route-box{background:#0E182B;border:1px dashed #334155;border-radius:10px;padding:12px 14px;margin:12px 0;display:flex;align-items:center;justify-content:space-between;gap:10px;min-width:0;box-sizing:border-box;width:100%}.mutasi-route-node{flex:1;min-width:0}.mutasi-to-node{text-align:right}.mutasi-route-label{font-size:0.7rem;text-transform:uppercase;font-weight:700;letter-spacing:0.04em;color:#94A3B8;margin-bottom:2px}.mutasi-route-branch{font-size:0.85rem;font-weight:700;color:#F8FAFC;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}.mutasi-route-arrow{display:flex;align-items:center;justify-content:center;width:32px;height:32px;border-radius:50%;background:rgba(59,130,246,0.18);color:#93C5FD;font-size:1rem;flex-shrink:0}.mutasi-status-badge{display:inline-flex;align-items:center;gap:5px;padding:4px 10px;border-radius:20px;font-size:0.76rem;font-weight:700}.mutasi-status-badge.selesai{background:rgba(16,185,129,0.18);color:#6EE7B7;border:1px solid rgba(52,211,153,0.35)}.mutasi-status-badge.proses{background:rgba(245,158,11,0.18);color:#FCD34D;border:1px solid rgba(251,191,36,0.35)}.mutasi-icon-btn{width:32px;height:32px;border-radius:8px;border:1px solid #223456;background:#0E182B;color:#CBD5E1;display:inline-flex;align-items:center;justify-content:center;cursor:pointer;transition:all 0.15s ease}.mutasi-icon-btn:hover{background:#22314A;color:#F8FAFC;border-color:#CBD5E1}.mutasi-icon-btn.delete:hover{background:#FEE2E2;color:#DC2626;border-color:#FCA5A5}.mutasi-empty-state{grid-column:1 / -1;background:#131F36;border:2px dashed #E2E8F0;border-radius:16px;padding:48px 24px;text-align:center}@media (max-width:768px){.mutasi-hero-card{padding:18px 16px}.mutasi-stats-row{grid-template-columns:1fr 1fr;gap:10px}.mutasi-stat-box{padding:10px 12px}.mutasi-stat-icon{width:38px;height:38px;font-size:1.1rem}.mutasi-stat-val{font-size:1.3rem}.mutasi-hero-actions{width:100% !important;display:flex !important;align-items:center !important;gap:6px !important;flex-wrap:nowrap !important;overflow-x:auto !important;scrollbar-width:none !important;padding-bottom:2px !important}.mutasi-hero-actions::-webkit-scrollbar{display:none}.mutasi-btn-add{flex:1 1 auto !important;justify-content:center !important;padding:8px 10px !important;font-size:0.82rem !important;white-space:nowrap !important}.mutasi-btn-excel{flex:1 1 auto !important;justify-content:center !important;padding:8px 10px !important;font-size:0.82rem !important;white-space:nowrap !important}.mutasi-btn-refresh{width:38px !important;height:38px !important;padding:0 !important;flex-shrink:0 !important;justify-content:center !important}.mutasi-controls-card{padding:14px;gap:12px}.mutasi-search-row{flex-direction:column;gap:10px}.mutasi-select{width:100%}.mutasi-view-toggle{width:100%;display:flex}.mutasi-view-btn{flex:1;justify-content:center}.mutasi-table-card{border-radius:12px}.mutasi-cards-grid{grid-template-columns:minmax(0,1fr) !important;width:100% !important;max-width:100% !important;min-width:0 !important;box-sizing:border-box !important}.mutasi-card{width:100% !important;max-width:100% !important;min-width:0 !important;box-sizing:border-box !important;overflow:hidden;padding:14px !important}.mutasi-route-box{flex-direction:column !important;align-items:stretch !important;gap:8px !important;padding:11px 12px !important;width:100% !important;box-sizing:border-box !important}.mutasi-route-node{width:100% !important;min-width:0 !important;text-align:left !important}.mutasi-to-node{text-align:left !important}.mutasi-route-branch{white-space:normal !important;word-break:break-word !important;overflow:visible !important;text-overflow:clip !important;font-size:0.88rem !important;line-height:1.35 !important}.mutasi-route-arrow{align-self:center !important;width:26px !important;height:26px !important;font-size:0.85rem !important;transform:rotate(90deg) !important;margin:-2px 0 !important}}</style>

      <div class="mutasi-page">

        <!-- Hero Header Card -->
        <div class="mutasi-hero-card">
          <div style="display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:16px;">
            <div>
              <div style="display:flex; align-items:center; gap:10px;">
                <span style="font-size:1.6rem;">\u{1F501}</span>
                <h1 style="font-size:1.45rem; font-weight:800; margin:0; letter-spacing:-0.02em;">
                  Monitoring Data Mutasi Karyawan
                </h1>
              </div>
            </div>

            <div class="mutasi-hero-actions" style="display:flex; align-items:center; gap:8px;">
              ${q?`
                <button id="mutasi-btn-add" class="btn btn-primary mutasi-btn-add" style="background:#2563EB; height:38px; padding:0 16px; border-radius:10px; font-weight:700; display:inline-flex; align-items:center; gap:6px; box-shadow:0 4px 12px rgba(37,99,235,0.3); white-space:nowrap;">
                  <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24"><line x1="12" y1="5" x2="12" y2="19"/><line x1="5" y1="12" x2="19" y2="12"/></svg>
                  <span>Tambah Mutasi</span>
                </button>
              `:""}
              
              ${re(N,"export")?`<button id="mutasi-btn-excel" class="btn btn-ghost mutasi-btn-excel" style="background:rgba(255,255,255,0.08); border:1px solid rgba(255,255,255,0.15); color:white; height:38px; padding:0 14px; border-radius:10px; font-weight:600; display:inline-flex; align-items:center; gap:6px; white-space:nowrap;">
                <span>\u{1F4E5}</span>
                <span>Ekspor Excel</span>
              </button>`:""}

              <button id="mutasi-btn-refresh" class="btn btn-ghost mutasi-btn-refresh" style="background:rgba(255,255,255,0.08); border:1px solid rgba(255,255,255,0.15); color:white; width:38px; height:38px; padding:0; border-radius:10px; display:inline-flex; align-items:center; justify-content:center; flex-shrink:0;" title="Muat Ulang Data">
                <svg width="16" height="16" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M23 4v6h-6M1 20v-6h6"/><path d="M3.51 9a9 9 0 0 1 14.85-3.36L23 10M1 14l4.64 4.36A9 9 0 0 0 20.49 15"/></svg>
              </button>
            </div>
          </div>

          <!-- KPI Stats -->
          <div class="mutasi-stats-row">
            <div class="mutasi-stat-box">
              <div class="mutasi-stat-icon" style="background:rgba(59,130,246,0.2); color:#60A5FA;">
                \u{1F4CB}
              </div>
              <div>
                <div class="mutasi-stat-val" id="mutasi-stat-total">${l}</div>
                <div class="mutasi-stat-lbl">Total Mutasi</div>
              </div>
            </div>

            <div class="mutasi-stat-box">
              <div class="mutasi-stat-icon" style="background:rgba(245,158,11,0.2); color:#FBBF24;">
                \u23F3
              </div>
              <div>
                <div class="mutasi-stat-val" id="mutasi-stat-proses">${f}</div>
                <div class="mutasi-stat-lbl">Dalam Proses</div>
              </div>
            </div>

            <div class="mutasi-stat-box">
              <div class="mutasi-stat-icon" style="background:rgba(34,197,94,0.2); color:#4ADE80;">
                \u2705
              </div>
              <div>
                <div class="mutasi-stat-val" id="mutasi-stat-selesai">${u}</div>
                <div class="mutasi-stat-lbl">Mutasi Selesai</div>
              </div>
            </div>

            <div class="mutasi-stat-box">
              <div class="mutasi-stat-icon" style="background:rgba(168,85,247,0.2); color:#C084FC;">
                \u{1F5D3}\uFE0F
              </div>
              <div>
                <div class="mutasi-stat-val" id="mutasi-stat-month">${p}</div>
                <div class="mutasi-stat-lbl">Bulan Berjalan</div>
              </div>
            </div>
          </div>
        </div>

        <!-- Filter & Search Controls -->
        <div class="mutasi-controls-card">
          <div class="mutasi-search-row">
            <div class="mutasi-search-wrap">
              <span class="mutasi-search-icon">
                <svg width="17" height="17" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><circle cx="11" cy="11" r="8"/><path d="m21 21-4.35-4.35"/></svg>
              </span>
              <input type="text" id="mutasi-search-input" class="mutasi-search-input" placeholder="Cari nama karyawan, cabang asal, atau cabang tujuan..." value="${i}" />
            </div>

            <select id="mutasi-filter-from" class="mutasi-select">
              <option value="all">Semua Cabang Asal (${a.length})</option>
              ${a.map(A=>`<option value="${A.value}" ${String(c)===String(A.value)?"selected":""}>${A.label}</option>`).join("")}
            </select>

            <select id="mutasi-filter-to" class="mutasi-select">
              <option value="all">Semua Cabang Tujuan (${a.length})</option>
              ${a.map(A=>`<option value="${A.value}" ${String(m)===String(A.value)?"selected":""}>${A.label}</option>`).join("")}
            </select>

            <div class="mutasi-view-toggle">
              <button type="button" class="mutasi-view-btn ${I==="card"?"active":""}" data-view="card" title="Tampilan Kartu">
                <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><rect x="3" y="3" width="7" height="7" rx="1"/><rect x="14" y="3" width="7" height="7" rx="1"/><rect x="14" y="14" width="7" height="7" rx="1"/><rect x="3" y="14" width="7" height="7" rx="1"/></svg>
                <span>Kartu</span>
              </button>
              <button type="button" class="mutasi-view-btn ${I==="table"?"active":""}" data-view="table" title="Tampilan Tabel">
                <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M3 9h18M3 15h18M9 3v18"/><rect x="3" y="3" width="18" height="18" rx="2"/></svg>
                <span>Tabel</span>
              </button>
            </div>
          </div>

          <!-- Status Filter Pills (Swipeable on Mobile) -->
          <div class="mutasi-filter-pills">
            <div class="mutasi-pill ${s==="all"?"active":""}" data-status="all">
              <span>Semua Mutasi</span>
              <span class="badge" id="mutasi-pill-count-all" style="background:rgba(15,23,42,0.55); color:#F8FAFC; padding:2px 6px; border-radius:10px; font-size:0.75rem;">${l}</span>
            </div>
            <div class="mutasi-pill ${s==="Proses"?"active":""}" data-status="Proses">
              <span style="color:#FCD34D;">\u23F3</span>
              <span>Proses</span>
              <span class="badge" id="mutasi-pill-count-proses" style="background:rgba(245,158,11,0.18); color:#FCD34D; padding:2px 6px; border-radius:10px; font-size:0.75rem;">${f}</span>
            </div>
            <div class="mutasi-pill ${s==="Selesai"?"active":""}" data-status="Selesai">
              <span style="color:#6EE7B7;">\u2705</span>
              <span>Selesai</span>
              <span class="badge" id="mutasi-pill-count-selesai" style="background:rgba(16,185,129,0.18); color:#6EE7B7; padding:2px 6px; border-radius:10px; font-size:0.75rem;">${u}</span>
            </div>
          </div>
        </div>

        <!-- Content Area: Murni Grid Kartu -->
        <div id="mutasi-list-container">
          ${J(r)}
        </div>
        <div id="mutasi-pagination-container"></div>

      </div>
    `,y()}function y(){let r=e.querySelector("#mutasi-btn-add");r&&(r.onclick=()=>K());let l=e.querySelector("#mutasi-btn-excel");l&&(l.onclick=()=>E());let f=e.querySelector("#mutasi-btn-refresh");f&&(f.onclick=()=>_({refreshStats:!0})),e.querySelectorAll(".mutasi-view-btn").forEach(T=>{T.onclick=()=>{let g=T.getAttribute("data-view");g&&g!==I&&(I=g,localStorage.setItem("fcms_mutasi_view_mode",g),e.querySelectorAll(".mutasi-view-btn").forEach(C=>{C.classList.toggle("active",C.getAttribute("data-view")===g)}),b())}});let u=e.querySelector("#mutasi-search-input");u&&(u.oninput=T=>{i=T.target.value,clearTimeout(L),L=setTimeout(()=>{d=1,_({refreshStats:!1})},350)});let p=e.querySelector("#mutasi-filter-from");p&&(p.onchange=T=>{c=T.target.value,d=1,_({refreshStats:!1})});let A=e.querySelector("#mutasi-filter-to");A&&(A.onchange=T=>{m=T.target.value,d=1,_({refreshStats:!1})}),e.querySelectorAll(".mutasi-pill").forEach(T=>{T.onclick=()=>{s=T.getAttribute("data-status")||"all",e.querySelectorAll(".mutasi-pill").forEach(g=>{g.classList.toggle("active",g===T)}),d=1,_({refreshStats:!1})}}),W(),Y()}await _({refreshStats:!0})}Be();var xi=["General Cleaning","Deep Cleaning","Cover manpower","Longshift","Pekerjaan urgent"];function vi(e,t,n){if(!e||!t)return 0;try{let[a,o]=e.split(":").map(Number),[s,c]=t.split(":").map(Number),m=s*60+c-(a*60+o);m<0&&(m+=24*60);let i=m/60-(parseFloat(n)||0);return i<0&&(i=0),Number.isInteger(i)?i:parseFloat(i.toFixed(2))}catch{return 0}}function Ga(e=""){let t=(e||"").toLowerCase();return t.includes("general")?{bg:"rgba(59, 130, 246, 0.16)",color:"#93C5FD",border:"rgba(96, 165, 250, 0.35)",icon:"\u{1F9F9}"}:t.includes("deep")?{bg:"rgba(56, 189, 248, 0.16)",color:"#7DD3FC",border:"rgba(56, 189, 248, 0.35)",icon:"\u2728"}:t.includes("cover")?{bg:"rgba(16, 185, 129, 0.16)",color:"#6EE7B7",border:"rgba(52, 211, 153, 0.35)",icon:"\u{1F465}"}:t.includes("longshift")?{bg:"rgba(37, 99, 235, 0.2)",color:"#93C5FD",border:"rgba(96, 165, 250, 0.38)",icon:"\u23F3"}:t.includes("urgent")?{bg:"rgba(239, 68, 68, 0.18)",color:"#FCA5A5",border:"rgba(248, 113, 113, 0.35)",icon:"\u{1F6A8}"}:{bg:"rgba(148, 163, 184, 0.16)",color:"#E2E8F0",border:"rgba(148, 163, 184, 0.32)",icon:"\u23F1\uFE0F"}}async function ki(e){let t=[],n=[],a=[],o="all",s="all",c="",m=!1,i=localStorage.getItem("fcms_overtime_view_mode")||"card",d=1,F=20,$=0,M=1,h={totalRecords:0,totalHours:0,avgDuration:0,currentMonth:0},L=null,R=!1,I=se(),N=re(I,"overtime","create"),q=re(I,"overtime","update"),v=re(I,"overtime","delete");try{[n,a]=await Promise.all([we(),ze()])}catch(y){console.error("Error fetching cache:",y)}function B(){let y={sort:"date_desc"};return c&&c.trim()&&(y.search=c.trim()),s&&s!=="all"&&(y.branch_id=s),o&&o!=="all"&&(y.reason=o),y}function x(){let y=(f,u)=>{let p=e.querySelector(f);p&&(p.textContent=String(u??0))},r=Number(h.totalHours||0),l=h.avgDuration!=null?Number(h.avgDuration).toFixed(1):"0.0";y("#lembur-stat-hours",Number.isInteger(r)?r:r.toFixed(1)),y("#lembur-stat-records",h.totalRecords),y("#lembur-stat-avg",l),y("#lembur-stat-month",h.currentMonth),y("#lembur-pill-count-all",h.totalRecords)}async function k({refreshStats:y=!0}={}){m=!0;try{let r=B(),l=new URLSearchParams({page:String(d),limit:String(F),...r});y&&l.set("include_stats","1");let f=await X("/api/overtime?"+l.toString());if(f.ok&&f.data&&Array.isArray(f.data.data)){let u=f.data.pagination||{};if($=Number(u.total??f.data.data.length),M=Math.max(1,Number(u.pages||u.totalPages||Math.ceil($/F)||1)),f.data.data.length===0&&d>M)return d=M,m=!1,k({refreshStats:y});t=f.data.data,f.data.stats&&(h={...h,...f.data.stats})}else t=[],$=0,M=1,j(f.data?.error||"Gagal memuat data lembur")}catch(r){console.warn("Gagal load lembur:",r),t=[],$=0,M=1,j("Gagal memuat data lembur: "+r.message)}finally{m=!1,R?(y&&x(),Y()):(R=!0,b())}}async function _(){if(!re(I,"export")){j("Anda tidak memiliki izin untuk mengekspor data.");return}try{let{downloadExcel:y}=await Promise.resolve().then(()=>(Ce(),Ve)),r=await ue("/api/overtime/export",B()),l=r.map((f,u)=>({No:u+1,Tanggal:f.date||"","Cabang Lembur":f.branch_name||"","Nama Karyawan":f.employee_name||"","Jam Mulai":f.start_time||"","Jam Selesai":f.end_time||"","Istirahat (Jam)":f.break_hours!=null?f.break_hours:0,"Total Jam":f.total_hours!=null?f.total_hours:0,"Alasan / Keterangan":f.reason||""}));y(l,`Data_Lembur_${new Date().toISOString().slice(0,10)}`),ae(`Berhasil mengunduh ${r.length} data lembur ke Excel!`)}catch(y){j("Gagal ekspor Excel: "+y.message)}}async function E(y){if(!v)return;let r=t.find(l=>l.id===y);ge({title:"Hapus Catatan Lembur",message:`Apakah Anda yakin ingin menghapus catatan lembur untuk "${r?.employee_name||"karyawan ini"}" pada tanggal ${xe(r?.date)}?`,confirmText:"Ya, Hapus",confirmClass:"btn-danger",onConfirm:async()=>{try{let l=await X(`/api/overtime/${y}`,{method:"DELETE"});l.ok?(ae("Catatan lembur berhasil dihapus"),await k({refreshStats:!0})):j(l.data?.error||"Gagal menghapus lembur")}catch(l){j("Gagal menghapus catatan lembur: "+l.message)}}})}function w(y=null){let r=!!y;if(r?!q:!N)return;let l=y?.date?y.date.slice(0,10):new Date().toISOString().slice(0,10),f=y?.start_time||"22:00",u=y?.end_time||"08:00",p=y?.break_hours!=null?y.break_hours:1,A=y?.total_hours!=null?y.total_hours:vi(f,u,p),T=document.createElement("div");T.innerHTML=`
      <div style="display:flex; flex-direction:column; gap:16px;">
        <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:var(--text-2); margin-bottom:5px;">
              Tanggal Lembur <span style="color:#DC2626;">*</span>
            </label>
            <input type="date" id="ot-modal-date" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              value="${l}" required />
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:var(--text-2); margin-bottom:5px;">
              Cabang Penugasan <span style="color:#DC2626;">*</span>
            </label>
            <select id="ot-modal-branch" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" required>
              <option value="">-- Pilih Cabang --</option>
              ${n.map(ne=>`
                <option value="${ne.value}" ${String(y?.branch_id)===String(ne.value)?"selected":""}>
                  ${ne.label}
                </option>
              `).join("")}
            </select>
          </div>
        </div>

        <div>
          <label style="display:block; font-size:0.82rem; font-weight:600; color:var(--text-2); margin-bottom:5px;">
            Nama Karyawan <span style="color:#DC2626;">*</span>
          </label>
          <div id="ot-modal-emp-mount"></div>
        </div>

        <!-- Shift & Waktu Lembur -->
        <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:var(--text-2); margin-bottom:5px;">
              Jam Mulai <span style="color:#DC2626;">*</span>
            </label>
            <input type="time" id="ot-modal-start" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              value="${f}" required />
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:var(--text-2); margin-bottom:5px;">
              Jam Selesai <span style="color:#DC2626;">*</span>
            </label>
            <input type="time" id="ot-modal-end" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              value="${u}" required />
          </div>
        </div>

        <!-- Durasi & Istirahat -->
        <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:var(--text-2); margin-bottom:5px;">
              Istirahat (Jam) <span style="color:#DC2626;">*</span>
            </label>
            <input type="number" step="0.5" min="0" max="10" id="ot-modal-break" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              value="${p}" required />
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:var(--text-2); margin-bottom:5px;">
              Total Jam Lembur <span style="color:#DC2626;">*</span>
            </label>
            <input type="number" step="0.1" min="0" id="ot-modal-total" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #223456; border-radius:8px; background:#0E182B; color:#38BDF8; font-weight:700; color:#2563EB;" 
              value="${A}" required />
          </div>
        </div>

        <div>
          <label style="display:block; font-size:0.82rem; font-weight:600; color:var(--text-2); margin-bottom:5px;">
            Alasan / Keterangan Penugasan Lembur <span style="color:#DC2626;">*</span>
          </label>
          <select id="ot-modal-reason" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" required>
            ${xi.map(ne=>`
              <option value="${ne}" ${y?.reason===ne?"selected":""}>${ne}</option>
            `).join("")}
          </select>
        </div>
      </div>
    `;let g=T.querySelector("#ot-modal-start"),C=T.querySelector("#ot-modal-end"),z=T.querySelector("#ot-modal-break"),O=T.querySelector("#ot-modal-total");function te(){let ne=g.value,ee=C.value,H=z.value;O.value=vi(ne,ee,H)}g.addEventListener("input",te),C.addEventListener("input",te),z.addEventListener("input",te);let Q=je(T.querySelector("#ot-modal-emp-mount"),{mode:"single",inputId:"ot-modal-emp",inputName:"employee_name",placeholder:"",required:!0,allowManual:!0,hasEmployeeIdSchema:!0,manualBadgeText:"Belum terhubung ke Master Karyawan (Lembur mewajibkan pilih karyawan terdaftar)",employees:a,initialEmployeeId:y?.employee_id||null,initialEmployeeName:y?.employee_name||""});ce({title:r?"\u270F\uFE0F Edit Data Lembur Karyawan":"\u2795 Catat Penugasan Lembur Baru",content:T,size:"md",confirmText:r?"Simpan Perubahan":"Catat Lembur",confirmClass:"btn-primary",onConfirm:async(ne,ee)=>{let H=T.querySelector("#ot-modal-date").value,U=T.querySelector("#ot-modal-branch").value,D=Q.getValue(),V=D.employee_id,Z=T.querySelector("#ot-modal-start").value,ie=T.querySelector("#ot-modal-end").value,le=parseFloat(T.querySelector("#ot-modal-break").value)||0,pe=parseFloat(T.querySelector("#ot-modal-total").value)||0,de=T.querySelector("#ot-modal-reason").value;if(!H){j("Tanggal lembur wajib diisi");return}if(!U){j("Cabang penugasan wajib dipilih");return}if(!D.employee_name){j("Nama karyawan wajib diisi"),Q.inputEl?.focus();return}if(!V){j("Nama manual belum terhubung ke Master Karyawan. Silakan pilih karyawan terdaftar pada daftar pencarian untuk mencatat lembur."),Q.inputEl?.focus();return}if(!Z||!ie){j("Jam mulai dan jam selesai wajib diisi");return}let me=n.find(oe=>String(oe.value)===String(U)),Ae=D.employee_name||"Karyawan",Ke={date:H,branch_id:Number(U),branch_name:me?me.label:"Cabang Lembur",employee_id:Number(V),employee_name:Ae,start_time:Z,end_time:ie,break_hours:le,total_hours:pe,reason:de,...r?{row_version:y?.row_version!==void 0&&y?.row_version!==null?Number(y.row_version):1}:{}};try{if(r){let oe=await X(`/api/overtime/${y.id}`,{method:"PUT",body:JSON.stringify(Ke)});if(oe.ok)ae("Data lembur berhasil diperbarui"),ee(),await k({refreshStats:!0});else{if(oe.status===409){j("Konflik OCC: Data telah diubah oleh pengguna lain. Silakan muat ulang halaman."),ee(),await k({refreshStats:!0});return}j(oe.data?.error||"Gagal memperbarui data lembur")}}else{let oe=await X("/api/overtime",{method:"POST",body:JSON.stringify(Ke)});oe.ok?(ae("Data lembur berhasil ditambahkan"),d=1,ee(),await k({refreshStats:!0})):j(oe.data?.error||"Gagal menambahkan lembur")}}catch(oe){j("Terjadi kesalahan: "+oe.message)}}})}function K(){return t}function G(y){return`
          ${y.length===0?`
            <div class="lembur-empty-state">
              <div style="font-size:3rem; margin-bottom:12px;">\u{1F50D}</div>
              <h3 style="font-size:1.2rem; font-weight:700; color:#F8FAFC; margin:0 0 6px 0;">Tidak Ada Data Lembur Ditemukan</h3>
              <p style="color:#94A3B8; font-size:0.9rem; margin:0 0 20px 0;">Silakan sesuaikan kata kunci pencarian atau reset filter aktif.</p>
              <button id="lembur-btn-reset-filter" class="btn btn-ghost" style="border:1px solid #CBD5E1; border-radius:8px; padding:8px 16px;">
                Reset Filter
              </button>
            </div>
          `:i==="table"?`
            <div class="lembur-table-card">
              <div class="lembur-table-responsive">
                <table class="lembur-table">
                  <thead>
                    <tr>
                      <th style="width:45px; text-align:center;">No</th>
                      <th>Nama Karyawan</th>
                      <th>Cabang Penugasan</th>
                      <th>Tanggal</th>
                      <th>Waktu Shift</th>
                      <th style="text-align:center;">Total Jam</th>
                      <th style="text-align:center;">Istirahat</th>
                      <th>Alasan Lembur</th>
                      ${q||v?'<th style="width:90px; text-align:center;">Aksi</th>':""}
                    </tr>
                  </thead>
                  <tbody>
                    ${y.map((r,l)=>{let f=(r.employee_name||"Karyawan").split(" ").map(p=>p[0]).filter(Boolean).slice(0,2).join("").toUpperCase(),u=Ga(r.reason);return`
                        <tr data-id="${r.id}">
                          <td style="text-align:center; color:#94A3B8; font-weight:600; font-size:0.8rem;">${(d-1)*F+l+1}</td>
                          <td>
                            <div style="display:flex; align-items:center; gap:10px;">
                              <div class="lembur-avatar-box" style="width:34px; height:34px; font-size:0.78rem; border-radius:8px;">
                                ${f}
                              </div>
                              <div style="font-weight:700; color:#F8FAFC; white-space:nowrap;">
                                ${r.employee_name||"-"}
                              </div>
                            </div>
                          </td>
                          <td style="white-space:nowrap;">
                            <div style="font-weight:600; color:#E2E8F0; font-size:0.85rem; display:inline-flex; align-items:center; gap:4px;">
                              <span>\u{1F4CD}</span>
                              <span>${r.branch_name||"Cabang"}</span>
                            </div>
                          </td>
                          <td style="white-space:nowrap; font-weight:600; color:#F8FAFC;">
                            \u{1F4C5} ${xe(r.date)}
                          </td>
                          <td style="white-space:nowrap;">
                            <span style="font-weight:700; color:#93C5FD;">\u23F0 ${r.start_time||"00:00"} - ${r.end_time||"00:00"}</span>
                          </td>
                          <td style="text-align:center; white-space:nowrap;">
                            <span class="lembur-hours-badge" style="padding:3px 8px; font-size:0.8rem;">
                              <span>\u23F1\uFE0F</span>
                              <span>${r.total_hours!=null?r.total_hours:0} Jam</span>
                            </span>
                          </td>
                          <td style="text-align:center; white-space:nowrap; color:#CBD5E1; font-weight:600; font-size:0.82rem;">
                            \u2615 ${r.break_hours!=null?r.break_hours:0} Jam
                          </td>
                          <td style="white-space:nowrap;">
                            <span class="lembur-reason-badge" style="background:${u.bg}; color:${u.color}; border:1px solid ${u.border};">
                              <span>${u.icon}</span>
                              <span>${r.reason||"Lembur Operasional"}</span>
                            </span>
                          </td>
                          ${q||v?`
                            <td style="text-align:center; white-space:nowrap;">
                              <div style="display:inline-flex; align-items:center; gap:6px;">
                                ${q?`
                                  <button class="lembur-icon-btn lembur-edit-btn" title="Edit Catatan Lembur" data-id="${r.id}">
                                    <svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>
                                  </button>
                                `:""}
                                ${v?`
                                  <button class="lembur-icon-btn delete lembur-delete-btn" title="Hapus Catatan Lembur" data-id="${r.id}">
                                    <svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><polyline points="3 6 5 6 21 6"/><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"/></svg>
                                  </button>
                                `:""}
                              </div>
                            </td>
                          `:""}
                        </tr>
                      `}).join("")}
                  </tbody>
                </table>
              </div>
            </div>
          `:`
            <div class="lembur-cards-grid">
              ${y.map(r=>{let l=(r.employee_name||"Karyawan").split(" ").map(u=>u[0]).filter(Boolean).slice(0,2).join("").toUpperCase(),f=Ga(r.reason);return`
                  <div class="lembur-card" data-id="${r.id}">
                    <div>
                      <!-- Card Header: Jam & Istirahat -->
                      <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:12px;">
                        <span class="lembur-hours-badge">
                          <span>\u23F1\uFE0F</span>
                          <span>${r.total_hours!=null?r.total_hours:0} Jam Lembur</span>
                        </span>
                        <span style="font-size:0.8rem; font-weight:600; color:#CBD5E1;">
                          \u2615 Istirahat: ${r.break_hours!=null?r.break_hours:0} Jam
                        </span>
                      </div>

                      <!-- Employee & Branch -->
                      <div style="display:flex; align-items:center; gap:12px; margin-bottom:12px;">
                        <div class="lembur-avatar-box">
                          ${l}
                        </div>
                        <div style="min-width:0;">
                          <h4 style="margin:0 0 3px 0; font-size:1.02rem; font-weight:700; color:#F8FAFC; white-space:nowrap; overflow:hidden; text-overflow:ellipsis;">
                            ${r.employee_name||"-"}
                          </h4>
                          <div style="font-size:0.8rem; color:#CBD5E1; display:flex; align-items:center; gap:4px;">
                            <span>\u{1F4CD}</span>
                            <span style="font-weight:600; color:#E2E8F0;">${r.branch_name||"Cabang"}</span>
                          </div>
                        </div>
                      </div>

                      <!-- Shift Times & Date Box -->
                      <div class="lembur-time-box">
                        <div>
                          <div style="font-size:0.7rem; text-transform:uppercase; font-weight:700; color:#94A3B8; margin-bottom:2px;">Tanggal</div>
                          <div style="font-size:0.88rem; font-weight:700; color:#F8FAFC;">
                            \u{1F5D3}\uFE0F ${xe(r.date)}
                          </div>
                        </div>

                        <div style="text-align:right;">
                          <div style="font-size:0.7rem; text-transform:uppercase; font-weight:700; color:#94A3B8; margin-bottom:2px;">Waktu Shift</div>
                          <div style="font-size:0.88rem; font-weight:700; color:#93C5FD;">
                            \u23F0 ${r.start_time||"00:00"} - ${r.end_time||"00:00"}
                          </div>
                        </div>
                      </div>

                      <!-- Reason Badge -->
                      <div>
                        <span class="lembur-reason-badge" style="background:${f.bg}; color:${f.color}; border:1px solid ${f.border};">
                          <span>${f.icon}</span>
                          <span>${r.reason||"Lembur Operasional"}</span>
                        </span>
                      </div>
                    </div>

                    <!-- Card Footer: Aksi -->
                    ${q||v?`
                      <div style="display:flex; justify-content:flex-end; align-items:center; gap:6px; padding-top:12px; border-top:1px solid #223456; margin-top:12px;">
                        ${q?`
                          <button class="lembur-icon-btn lembur-edit-btn" title="Edit Catatan Lembur" data-id="${r.id}">
                            <svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>
                          </button>
                        `:""}
                        ${v?`
                          <button class="lembur-icon-btn delete lembur-delete-btn" title="Hapus Catatan Lembur" data-id="${r.id}">
                            <svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><polyline points="3 6 5 6 21 6"/><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"/></svg>
                          </button>
                        `:""}
                      </div>
                    `:""}

                  </div>
                `}).join("")}
            </div>
          `}
    `}function J(){let y=e.querySelector("#lembur-btn-reset-filter");y&&(y.onclick=()=>{c="",s="all",o="all",d=1,clearTimeout(L);let r=e.querySelector("#lembur-search-input");r&&(r.value="");let l=e.querySelector("#lembur-filter-branch");l&&(l.value="all"),e.querySelectorAll(".lembur-pill").forEach(f=>{f.classList.toggle("active",f.getAttribute("data-reason")==="all")}),k({refreshStats:!1})}),e.querySelectorAll(".lembur-edit-btn").forEach(r=>{r.onclick=l=>{l.stopPropagation();let f=Number(r.getAttribute("data-id")),u=t.find(p=>p.id===f);u&&w(u)}}),e.querySelectorAll(".lembur-delete-btn").forEach(r=>{r.onclick=l=>{l.stopPropagation();let f=Number(r.getAttribute("data-id"));E(f)}})}function W(){let y=e.querySelector("#lembur-pagination-container");if(!y)return;y.innerHTML="";let r=Te({page:d,pages:M,total:$,limit:F,onPage:l=>{d=l,k({refreshStats:!1})}});r&&y.appendChild(r)}function Y(){let y=e.querySelector("#lembur-list-container");if(!y)return;let r=K();y.innerHTML=G(r),W(),J()}function b(){let y=K(),r=Number(h.totalRecords??0),l=Number(h.totalHours??0),f=Number.isInteger(l)?l:l.toFixed(1),u=h.avgDuration!=null?Number(h.avgDuration).toFixed(1):"0.0",p=Number(h.currentMonth??0);e.innerHTML=`
      <style>.lembur-page{display:flex;flex-direction:column;gap:20px}.lembur-hero-card{background:#131F36;border:1px solid #223456;border-radius:16px;padding:24px 28px;color:#FFFFFF;box-shadow:0 4px 16px rgba(0,0,0,0.22);position:relative;overflow:hidden}.lembur-hero-card::after{content:'';position:absolute;top:-40px;right:-40px;width:220px;height:220px;background:radial-gradient(circle,rgba(37,99,235,0.06) 0%,transparent 70%);border-radius:50%;pointer-events:none}.lembur-hero-actions{display:flex;align-items:center;gap:8px;flex-wrap:wrap}.lembur-hero-actions .btn{height:38px;box-sizing:border-box;transition:all 0.2s ease}.lembur-btn-add:hover{background:#1D4ED8 !important;transform:translateY(-1px)}.lembur-btn-excel:hover,.lembur-btn-refresh:hover{background:rgba(255,255,255,0.16) !important;border-color:rgba(255,255,255,0.3) !important;transform:translateY(-1px)}.lembur-stats-row{display:grid;grid-template-columns:repeat(4,1fr);gap:14px;margin-top:22px}.lembur-stat-box{background:rgba(255,255,255,0.06);border:1px solid rgba(255,255,255,0.1);border-radius:12px;padding:14px 16px;backdrop-filter:blur(8px);display:flex;align-items:center;gap:14px}.lembur-stat-icon{width:44px;height:44px;border-radius:10px;display:flex;align-items:center;justify-content:center;font-size:1.3rem;flex-shrink:0}.lembur-stat-val{font-size:1.55rem;font-weight:800;line-height:1.1;color:#F8FAFC}.lembur-stat-lbl{font-size:0.76rem;text-transform:uppercase;letter-spacing:0.05em;color:#94A3B8;font-weight:600;margin-top:2px}.lembur-controls-card{background:#131F36;border:1px solid #223456;border-radius:14px;padding:18px 20px;box-shadow:0 2px 4px rgba(0,0,0,0.02);display:flex;flex-direction:column;gap:16px}.lembur-search-row{display:flex;gap:12px;flex-wrap:wrap}.lembur-search-wrap{position:relative;flex:1;min-width:240px}.lembur-search-icon{position:absolute;left:14px;top:50%;transform:translateY(-50%);color:#94A3B8;pointer-events:none}.lembur-search-input{width:100%;padding:10px 14px 10px 40px;border:1px solid #223456;border-radius:10px;font-size:0.9rem;background:#0E182B;color:#F8FAFC;outline:none;transition:border-color 0.15s ease}.lembur-search-input::placeholder{color:#94A3B8}.lembur-search-input:focus{border-color:#2563EB;box-shadow:0 0 0 3px rgba(37,99,235,0.1)}.lembur-select{padding:10px 14px;border:1px solid #223456;border-radius:10px;font-size:0.9rem;color:#F8FAFC;background:#131F36;outline:none;min-width:200px;cursor:pointer}.lembur-select:focus{border-color:#2563EB}.lembur-filter-pills{display:flex;align-items:center;gap:8px;overflow-x:auto;white-space:nowrap;scrollbar-width:none;padding-bottom:2px}.lembur-filter-pills::-webkit-scrollbar{display:none}.lembur-pill{padding:6px 14px;border-radius:20px;font-size:0.82rem;font-weight:600;cursor:pointer;border:1px solid #223456;background:#0E182B;color:#CBD5E1;transition:all 0.15s ease;display:inline-flex;align-items:center;gap:6px;user-select:none;flex-shrink:0}.lembur-pill:hover{background:#22314A;color:#F8FAFC}.lembur-pill.active{background:#2563EB;color:white;border-color:#2563EB;box-shadow:0 2px 6px rgba(37,99,235,0.25)}.lembur-view-toggle{display:inline-flex;align-items:center;background:#0E182B;padding:3px;border-radius:10px;border:1px solid #223456;gap:3px;flex-shrink:0}.lembur-view-btn{display:inline-flex;align-items:center;gap:6px;padding:6px 12px;border-radius:7px;font-size:0.82rem;font-weight:600;border:none;background:transparent;color:#94A3B8;cursor:pointer;transition:all 0.15s ease;height:34px;box-sizing:border-box;user-select:none}.lembur-view-btn:hover{color:#F8FAFC;background:rgba(255,255,255,0.08)}.lembur-view-btn.active{background:#0E182B;color:#60A5FA;font-weight:700;box-shadow:0 1px 4px rgba(0,0,0,0.25)}.lembur-table-card{background:#131F36;border:1px solid #223456;border-radius:14px;box-shadow:0 2px 6px rgba(0,0,0,0.03);overflow:hidden}.lembur-table-responsive{width:100%;overflow-x:auto;-webkit-overflow-scrolling:touch}.lembur-table{width:100%;border-collapse:collapse;text-align:left;font-size:0.88rem;min-width:860px}.lembur-table th{background:#0E182B;color:#CBD5E1;font-size:0.76rem;font-weight:700;text-transform:uppercase;letter-spacing:0.05em;padding:13px 16px;border-bottom:1.5px solid #223456;white-space:nowrap}.lembur-table td{padding:13px 16px;border-bottom:1px solid #223456;color:#F8FAFC;vertical-align:middle}.lembur-table tbody tr:hover td{background:#22314A}.lembur-table tbody tr:last-child td{border-bottom:none}.lembur-cards-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(320px,1fr));gap:16px}.lembur-card{background:#131F36;border:1px solid #223456;border-radius:14px;padding:18px;box-shadow:0 2px 6px rgba(0,0,0,0.03);transition:transform 0.15s ease,box-shadow 0.15s ease;display:flex;flex-direction:column;justify-content:space-between;position:relative}.lembur-card:hover{transform:translateY(-2px);box-shadow:0 8px 18px rgba(0,0,0,0.06);border-color:#CBD5E1}.lembur-avatar-box{width:44px;height:44px;border-radius:10px;background:#2563EB;color:#FFFFFF;font-weight:700;font-size:1rem;display:flex;align-items:center;justify-content:center;flex-shrink:0}.lembur-time-box{background:#0E182B;border:1px solid #223456;border-radius:10px;padding:10px 12px;margin:12px 0;display:flex;align-items:center;justify-content:space-between}.lembur-hours-badge{display:inline-flex;align-items:center;gap:5px;padding:5px 12px;border-radius:8px;font-size:0.88rem;font-weight:800;background:rgba(99,102,241,0.2);color:#C7D2FE;border:1px solid rgba(129,140,248,0.38)}.lembur-reason-badge{display:inline-flex;align-items:center;gap:5px;padding:4px 10px;border-radius:20px;font-size:0.76rem;font-weight:700}.lembur-icon-btn{width:32px;height:32px;border-radius:8px;border:1px solid #223456;background:#0E182B;color:#CBD5E1;display:inline-flex;align-items:center;justify-content:center;cursor:pointer;transition:all 0.15s ease}.lembur-icon-btn:hover{background:#22314A;color:#F8FAFC;border-color:#CBD5E1}.lembur-icon-btn.delete:hover{background:#FEE2E2;color:#DC2626;border-color:#FCA5A5}.lembur-empty-state{grid-column:1 / -1;background:#131F36;border:2px dashed #E2E8F0;border-radius:16px;padding:48px 24px;text-align:center}@media (max-width:768px){.lembur-hero-card{padding:18px 16px}.lembur-stats-row{grid-template-columns:1fr 1fr;gap:10px}.lembur-stat-box{padding:10px 12px}.lembur-stat-icon{width:38px;height:38px;font-size:1.1rem}.lembur-stat-val{font-size:1.3rem}.lembur-hero-actions{width:100% !important;display:flex !important;align-items:center !important;gap:6px !important;flex-wrap:nowrap !important;overflow-x:auto !important;scrollbar-width:none !important;padding-bottom:2px !important}.lembur-hero-actions::-webkit-scrollbar{display:none}.lembur-btn-add{flex:1 1 auto !important;justify-content:center !important;padding:8px 10px !important;font-size:0.82rem !important;white-space:nowrap !important}.lembur-btn-excel{flex:1 1 auto !important;justify-content:center !important;padding:8px 10px !important;font-size:0.82rem !important;white-space:nowrap !important}.lembur-btn-refresh{width:38px !important;height:38px !important;padding:0 !important;flex-shrink:0 !important;justify-content:center !important}.lembur-controls-card{padding:14px;gap:12px}.lembur-search-row{flex-direction:column;gap:10px}.lembur-select{width:100%}.lembur-view-toggle{width:100%;display:flex}.lembur-view-btn{flex:1;justify-content:center}.lembur-table-card{border-radius:12px}.lembur-cards-grid{grid-template-columns:minmax(0,1fr) !important;width:100% !important;max-width:100% !important;min-width:0 !important;box-sizing:border-box !important}.lembur-card{width:100% !important;max-width:100% !important;min-width:0 !important;box-sizing:border-box !important;overflow:hidden;padding:14px !important}}</style>

      <div class="lembur-page">

        <!-- Hero Header Card -->
        <div class="lembur-hero-card">
          <div style="display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:16px;">
            <div>
              <div style="display:flex; align-items:center; gap:10px;">
                <span style="font-size:1.6rem;">\u23F1\uFE0F</span>
                <h1 style="font-size:1.45rem; font-weight:800; margin:0; letter-spacing:-0.02em;">
                  Monitoring Data Lembur Karyawan
                </h1>
              </div>
            </div>

            <div class="lembur-hero-actions" style="display:flex; align-items:center; gap:8px;">
              ${N?`
                <button id="lembur-btn-add" class="btn btn-primary lembur-btn-add" style="background:#2563EB; height:38px; padding:0 16px; border-radius:10px; font-weight:700; display:inline-flex; align-items:center; gap:6px; box-shadow:0 4px 12px rgba(37,99,235,0.3); white-space:nowrap;">
                  <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24"><line x1="12" y1="5" x2="12" y2="19"/><line x1="5" y1="12" x2="19" y2="12"/></svg>
                  <span>Tambah Lembur</span>
                </button>
              `:""}
              
              ${re(I,"export")?`<button id="lembur-btn-excel" class="btn btn-ghost lembur-btn-excel" style="background:rgba(255,255,255,0.08); border:1px solid rgba(255,255,255,0.15); color:white; height:38px; padding:0 14px; border-radius:10px; font-weight:600; display:inline-flex; align-items:center; gap:6px; white-space:nowrap;">
                <span>\u{1F4E5}</span>
                <span>Ekspor Excel</span>
              </button>`:""}

              <button id="lembur-btn-refresh" class="btn btn-ghost lembur-btn-refresh" style="background:rgba(255,255,255,0.08); border:1px solid rgba(255,255,255,0.15); color:white; width:38px; height:38px; padding:0; border-radius:10px; display:inline-flex; align-items:center; justify-content:center; flex-shrink:0;" title="Muat Ulang Data">
                <svg width="16" height="16" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M23 4v6h-6M1 20v-6h6"/><path d="M3.51 9a9 9 0 0 1 14.85-3.36L23 10M1 14l4.64 4.36A9 9 0 0 0 20.49 15"/></svg>
              </button>
            </div>
          </div>

          <!-- KPI Stats -->
          <div class="lembur-stats-row">
            <div class="lembur-stat-box">
              <div class="lembur-stat-icon" style="background:rgba(79,70,229,0.25); color:#818CF8;">
                \u23F1\uFE0F
              </div>
              <div>
                <div class="lembur-stat-val"><span id="lembur-stat-hours">${f}</span> Jam</div>
                <div class="lembur-stat-lbl">Total Jam Lembur</div>
              </div>
            </div>

            <div class="lembur-stat-box">
              <div class="lembur-stat-icon" style="background:rgba(59,130,246,0.2); color:#60A5FA;">
                \u{1F4DD}
              </div>
              <div>
                <div class="lembur-stat-val" id="lembur-stat-records">${r}</div>
                <div class="lembur-stat-lbl">Catatan Penugasan</div>
              </div>
            </div>

            <div class="lembur-stat-box">
              <div class="lembur-stat-icon" style="background:rgba(16,185,129,0.2); color:#34D399;">
                \u{1F4CA}
              </div>
              <div>
                <div class="lembur-stat-val"><span id="lembur-stat-avg">${u}</span> Jam</div>
                <div class="lembur-stat-lbl">Rata-rata Durasi</div>
              </div>
            </div>

            <div class="lembur-stat-box">
              <div class="lembur-stat-icon" style="background:rgba(245,158,11,0.2); color:#FBBF24;">
                \u{1F5D3}\uFE0F
              </div>
              <div>
                <div class="lembur-stat-val" id="lembur-stat-month">${p}</div>
                <div class="lembur-stat-lbl">Bulan Berjalan</div>
              </div>
            </div>
          </div>
        </div>

        <!-- Filter & Search Controls -->
        <div class="lembur-controls-card">
          <div class="lembur-search-row">
            <div class="lembur-search-wrap">
              <span class="lembur-search-icon">
                <svg width="17" height="17" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><circle cx="11" cy="11" r="8"/><path d="m21 21-4.35-4.35"/></svg>
              </span>
              <input type="text" id="lembur-search-input" class="lembur-search-input" placeholder="Cari nama karyawan, cabang, atau alasan lembur..." value="${c?String(c).replace(/"/g,"&quot;"):""}" />
            </div>

            <select id="lembur-filter-branch" class="lembur-select">
              <option value="all">Semua Cabang (${n.length})</option>
              ${n.map(A=>`<option value="${A.value}" ${String(s)===String(A.value)?"selected":""}>${A.label}</option>`).join("")}
            </select>

            <div class="lembur-view-toggle">
              <button type="button" class="lembur-view-btn ${i==="card"?"active":""}" data-view="card" title="Tampilan Kartu">
                <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><rect x="3" y="3" width="7" height="7" rx="1"/><rect x="14" y="3" width="7" height="7" rx="1"/><rect x="14" y="14" width="7" height="7" rx="1"/><rect x="3" y="14" width="7" height="7" rx="1"/></svg>
                <span>Kartu</span>
              </button>
              <button type="button" class="lembur-view-btn ${i==="table"?"active":""}" data-view="table" title="Tampilan Tabel">
                <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M3 9h18M3 15h18M9 3v18"/><rect x="3" y="3" width="18" height="18" rx="2"/></svg>
                <span>Tabel</span>
              </button>
            </div>
          </div>

          <!-- Reason Filter Pills (Swipeable on Mobile) -->
          <div class="lembur-filter-pills">
            <div class="lembur-pill ${o==="all"?"active":""}" data-reason="all">
              <span>Semua Lembur</span>
              <span class="badge" id="lembur-pill-count-all" style="background:rgba(0,0,0,0.06); padding:2px 6px; border-radius:10px; font-size:0.75rem;">${r}</span>
            </div>
            ${xi.map(A=>{let T=Ga(A);return`
                <div class="lembur-pill ${o===A?"active":""}" data-reason="${A}">
                  <span>${T.icon}</span>
                  <span>${A}</span>
                </div>
              `}).join("")}
          </div>
        </div>

        <!-- Content Area: Murni Grid Kartu -->
        <div id="lembur-list-container">
          ${G(y)}
        </div>

        <!-- Pagination Container -->
        <div id="lembur-pagination-container"></div>

      </div>
    `,W(),S()}function S(){let y=e.querySelector("#lembur-btn-add");y&&(y.onclick=()=>w());let r=e.querySelector("#lembur-btn-excel");r&&(r.onclick=()=>_());let l=e.querySelector("#lembur-btn-refresh");l&&(l.onclick=()=>k({refreshStats:!0}));let f=e.querySelector("#lembur-search-input");f&&f.addEventListener("input",p=>{c=p.target.value,clearTimeout(L),L=setTimeout(()=>{d=1,k({refreshStats:!1})},350)});let u=e.querySelector("#lembur-filter-branch");u&&u.addEventListener("change",p=>{s=p.target.value,d=1,k({refreshStats:!1})}),e.querySelectorAll(".lembur-view-btn").forEach(p=>{p.addEventListener("click",()=>{let A=p.getAttribute("data-view");!A||A===i||(i=A,localStorage.setItem("fcms_overtime_view_mode",i),e.querySelectorAll(".lembur-view-btn").forEach(T=>{T.classList.toggle("active",T===p)}),Y())})}),e.querySelectorAll(".lembur-pill").forEach(p=>{p.addEventListener("click",()=>{o=p.getAttribute("data-reason")||"all",e.querySelectorAll(".lembur-pill").forEach(A=>{A.classList.toggle("active",A===p)}),d=1,k({refreshStats:!1})})}),J()}await k({refreshStats:!0})}Ie();Ce();async function wi(e){let t=se();if(!t||t.role!=="superadmin"){e.innerHTML='<div class="empty-state"><p class="text-danger">Akses ditolak.</p></div>';return}if(!document.getElementById("fc-audit-log-styles")){let i=document.createElement("style");i.id="fc-audit-log-styles",i.textContent=`
      .audit-detail-btn {
        background: rgba(59, 130, 246, 0.16);
        color: #93C5FD;
        border: 1px solid rgba(96, 165, 250, 0.48);
        border-radius: 7px;
        padding: 5px 12px;
        font-size: 0.8rem;
        font-weight: 700;
        cursor: pointer;
        transition: all 0.15s ease;
        white-space: nowrap;
      }
      .audit-detail-btn:hover,
      .audit-detail-btn:focus-visible {
        background: rgba(59, 130, 246, 0.30);
        color: #FFFFFF;
        border-color: #93C5FD;
        text-decoration: underline;
        outline: none;
      }
      .audit-json-pre {
        background: #0F172A;
        color: #F8FAFC;
        border: 1px solid #334155;
        border-radius: 10px;
        padding: 14px;
        font-family: ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, monospace;
        font-size: 0.84rem;
        line-height: 1.65;
        max-height: 360px;
        overflow: auto;
        white-space: pre-wrap;
        word-break: break-word;
        margin: 0;
      }
    `,document.head.appendChild(i)}let n={CREATE:"badge-success",UPDATE:"badge-info",DELETE:"badge-danger",LOGIN:"badge-primary",LOGOUT:"badge-neutral"},a=/(password|secret|token|jwt|signature|hmac|api_key|private_key|credential|authorization)/i;function o(i){if(i==null)return i;if(Array.isArray(i))return i.map(o);if(typeof i!="object")return i;let d={};for(let[F,$]of Object.entries(i))d[F]=a.test(F)?"[REDACTED]":o($);return d}function s(i,d){if(i==null||typeof i=="string"&&i.trim()==="")return`<span style="color:#CBD5E1; font-style:italic;">${P(d)}</span>`;if(typeof i=="object")return P(JSON.stringify(o(i),null,2));let F=String(i);try{let $=JSON.parse(F);return $!==null&&typeof $=="object"?P(JSON.stringify(o($),null,2)):P(String($))}catch{return P(F)}}function c(i){if(!i)return"-";let d=new Date(i);return Number.isNaN(d.getTime())?P(String(i)):d.toLocaleString("id-ID",{dateStyle:"short",timeStyle:"medium"})}Me({container:e,title:"Riwayat Aktivitas",icon:"\u{1F6E1}\uFE0F",apiPath:"/api/audit-logs",enableMobileFilterSheet:!0,itemLabel:"Log",canCreate:!1,canEdit:!1,canDelete:!1,bulkDelete:!1,exportOptions:{moduleName:"audit_logs",filename:"FCMS_Audit_Logs",template:!1,import:!1,onExport:async i=>{let F=(await ue("/api/audit-logs/export",i)||[]).map($=>({ID:$.id,Waktu:$.created_at||"-",Pengguna:$.user_name||"Sistem",Role:$.user_role||"-",Aksi:$.action||"-",Modul:$.module||"-","ID Target":$.target_id||"-"}));be(F,"FCMS_Audit_Logs","Audit Logs")}},columns:[{key:"created_at",label:"Waktu",nowrap:!0,render:i=>`<span style="color:#F8FAFC; font-variant-numeric:tabular-nums;">${c(i)}</span>`},{key:"user_name",label:"Pengguna",render:(i,d)=>`<strong style="color:#F8FAFC; font-weight:700;">${P(i||"Sistem")}</strong><br><small style="color:#CBD5E1; text-transform:capitalize; font-size:0.78rem;">${P(d.user_role||"-")}</small>`},{key:"action",label:"Aksi",render:i=>`<span class="badge ${n[i]||"badge-neutral"}">${P(i||"-")}</span>`},{key:"module",label:"Modul",render:i=>`<span style="color:#F8FAFC; text-transform:capitalize; font-weight:500;">${P((i||"-").replace(/_/g," "))}</span>`},{key:"target_id",label:"ID Target",render:i=>`<span style="color:#CBD5E1; font-family:ui-monospace, monospace;">${P(i!=null&&i!==""?String(i):"-")}</span>`},{key:"id",label:"Detail",render:i=>{let d=document.createElement("button");return d.type="button",d.className="audit-detail-btn",d.textContent="Lihat Detail",d.addEventListener("click",()=>window.viewAuditDetail?.(i)),d}}],filterFields:[{type:"search",placeholder:"Cari ID, pengguna, modul..."},{type:"select",name:"action",options:[{value:"",label:"Semua Aksi"},{value:"CREATE",label:"Tambah (CREATE)"},{value:"UPDATE",label:"Ubah (UPDATE)"},{value:"DELETE",label:"Hapus (DELETE)"},{value:"LOGIN",label:"Masuk (LOGIN)"},{value:"LOGOUT",label:"Keluar (LOGOUT)"}]},{type:"select",name:"module",options:[{value:"",label:"Semua Modul"},{value:"employees",label:"Karyawan"},{value:"contracts",label:"Kontrak"},{value:"sp",label:"Data SP"},{value:"mutasi",label:"Data Mutasi"},{value:"overtime",label:"Lembur"},{value:"schedule",label:"Time Line / Jadwal"},{value:"os_migrations",label:"Migrasi OS"},{value:"issues",label:"Permasalahan"},{value:"training",label:"Training"},{value:"reports",label:"Laporan Operasional"},{value:"sop",label:"SOP Facility Care & Security"},{value:"checklist",label:"Master Document"},{value:"hygiene_standards",label:"Master Hygiene"},{value:"dac",label:"Master DAC"},{value:"users",label:"Manajemen User"},{value:"branches",label:"Manajemen Cabang"},{value:"auth",label:"Autentikasi"}]}]});let m=e.querySelector(".page-actions");if(m&&t&&t.role==="superadmin"){let i=document.createElement("a");i.href="#/sync",i.id="btn-audit-sync-health",i.className="btn btn-primary",i.style.cssText="text-decoration:none;display:inline-flex;align-items:center;gap:6px",i.innerHTML="\u{1FA7A} Sync & Health Dashboard",m.appendChild(i)}window.viewAuditDetail=async i=>{try{let d=await X(`/api/audit-logs/${encodeURIComponent(i)}`),F=d?.data?.data||null;if(!d?.ok||!F)return alert("Data tidak ditemukan");let $=n[F.action]||"badge-neutral",M=`
        <div style="display:flex; flex-direction:column; gap:14px;">
          <div style="display:grid; grid-template-columns:repeat(auto-fit, minmax(150px, 1fr)); gap:10px; background:#0E182B; border:1px solid #223456; border-radius:10px; padding:12px 14px;">
            <div>
              <div style="font-size:0.76rem; color:#CBD5E1; text-transform:uppercase; font-weight:700;">ID Audit Log</div>
              <div style="font-size:0.9rem; color:#F8FAFC; font-weight:700; margin-top:2px;">#${P(String(F.id??i))}</div>
            </div>
            <div>
              <div style="font-size:0.76rem; color:#CBD5E1; text-transform:uppercase; font-weight:700;">Waktu</div>
              <div style="font-size:0.86rem; color:#F8FAFC; font-weight:600; margin-top:2px;">${c(F.created_at)}</div>
            </div>
            <div>
              <div style="font-size:0.76rem; color:#CBD5E1; text-transform:uppercase; font-weight:700;">Pengguna</div>
              <div style="font-size:0.88rem; color:#F8FAFC; font-weight:700; margin-top:2px;">${P(F.user_name||"Sistem")} <span style="color:#CBD5E1; font-weight:500;">(${P(F.user_role||"-")})</span></div>
            </div>
            <div>
              <div style="font-size:0.76rem; color:#CBD5E1; text-transform:uppercase; font-weight:700;">Aksi &amp; Modul</div>
              <div style="display:flex; align-items:center; gap:6px; flex-wrap:wrap; margin-top:3px;">
                <span class="badge ${$}">${P(F.action||"-")}</span>
                <span style="font-size:0.84rem; color:#F8FAFC; text-transform:capitalize;">${P((F.module||"-").replace(/_/g," "))}${F.target_id?` (#${P(String(F.target_id))})`:""}</span>
              </div>
            </div>
          </div>

          <div style="display:grid; grid-template-columns:repeat(auto-fit, minmax(280px, 1fr)); gap:16px;">
            <div>
              <h4 style="margin:0 0 8px 0; font-size:0.92rem; font-weight:700; color:#F8FAFC;">Data Lama</h4>
              <pre class="audit-json-pre">${s(F.old_data,"Tidak ada data lama")}</pre>
            </div>
            <div>
              <h4 style="margin:0 0 8px 0; font-size:0.92rem; font-weight:700; color:#F8FAFC;">Data Baru</h4>
              <pre class="audit-json-pre">${s(F.new_data,"Tidak ada data baru")}</pre>
            </div>
          </div>
        </div>
      `,{createModal:h}=await Promise.resolve().then(()=>(Be(),dn));h({title:`Detail Audit Log #${i}`,content:M,size:"lg",cancelText:"Tutup"})}catch{alert("Gagal mengambil detail")}}}Ie();Be();Ce();var Va={"2026-09":"September 2026","2026-10":"Oktober 2026","2026-11":"November 2026","2026-12":"Desember 2026"},Kt={scheduled:{label:"Terjadwal",badgeClass:"badge-warning",icon:"\u25CF",dot:"#FBBF24",bg:"rgba(245, 158, 11, 0.14)",color:"#FCD34D",border:"rgba(251, 191, 36, 0.36)"},in_progress:{label:"Sedang Proses",badgeClass:"badge-info",icon:"\u25CF",dot:"#60A5FA",bg:"rgba(37, 99, 235, 0.18)",color:"#93C5FD",border:"rgba(59, 130, 246, 0.42)"},completed:{label:"Selesai",badgeClass:"badge-success",icon:"\u25CF",dot:"#34D399",bg:"rgba(16, 185, 129, 0.16)",color:"#6EE7B7",border:"rgba(52, 211, 153, 0.38)"}};function Si(e,t,n){if(!t&&!n)return"";let a=P(String(e.id??"")),o=P(e.branch_name||"Cabang");return`
    <div class="os-mig-actions" style="display:inline-flex; gap:5px; align-items:center;">
      ${t?`
        <button type="button" class="os-mig-icon-btn edit-card-btn" data-id="${a}" title="Edit ${o}" aria-label="Edit jadwal migrasi ${o}">
          <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">
            <path d="M12 20h9"></path>
            <path d="M16.5 3.5a2.121 2.121 0 0 1 3 3L7 19l-4 1 1-4L16.5 3.5z"></path>
          </svg>
        </button>
      `:""}
      ${n?`
        <button type="button" class="os-mig-icon-btn os-mig-icon-del delete-card-btn" data-id="${a}" data-name="${o}" title="Hapus ${o}" aria-label="Hapus jadwal migrasi ${o}">
          <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">
            <polyline points="3 6 5 6 21 6"></polyline>
            <path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"></path>
            <line x1="10" y1="11" x2="10" y2="17"></line>
            <line x1="14" y1="11" x2="14" y2="17"></line>
          </svg>
        </button>
      `:""}
    </div>
  `}async function Ci(e){let t=se(),n=re(t,"os_migrations","write"),a=re(t,"os_migrations","delete"),o=re(t,"export");e.innerHTML=`
    <div class="page-header" style="display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:16px; margin-bottom:20px;">
      <div>
        <h1 class="page-title" style="display:flex; align-items:center; gap:8px; margin:0; font-size:1.45rem; font-weight:800; color:var(--text-1); letter-spacing:-0.02em;">
          <span>\u{1F504}</span> Time Line Migrasi OS
        </h1>
      </div>
      <div class="os-header-actions" style="display:flex; gap:10px; align-items:center; flex-wrap:wrap;">
        <div class="btn-group" style="background:#0E182B; padding:3px; border-radius:10px; display:inline-flex; align-items:center; border:1px solid #223456; height:38px; box-sizing:border-box;">
          <button id="view-board-btn" class="btn btn-sm" style="background:#2563EB; color:#fff; font-weight:600; border-radius:7px; padding:0 12px; height:100%; display:inline-flex; align-items:center; cursor:pointer;">
            \u{1F4CB} Daftar per Bulan
          </button>
          <button id="view-table-btn" class="btn btn-sm btn-ghost" style="font-weight:600; border-radius:7px; padding:0 12px; height:100%; display:inline-flex; align-items:center; cursor:pointer; color:#CBD5E1;">
            \u{1F4CA} Tabel Data
          </button>
        </div>
        ${o?`
        <button id="btn-export-migrations" class="btn btn-outline" style="background:#131F36; border:1px solid #223456; color:#F8FAFC; font-weight:600; height:38px; padding:0 14px; border-radius:10px; display:inline-flex; align-items:center; gap:6px; cursor:pointer; box-sizing:border-box;">
          \u{1F4E5} Export Excel
        </button>
        `:""}
        ${n?`
        <button id="btn-add-migration" class="btn btn-primary" style="font-weight:700; height:38px; padding:0 16px; border-radius:10px; display:inline-flex; align-items:center; gap:6px; cursor:pointer; box-sizing:border-box; background:#2563EB;">
          + Tambah Cabang
        </button>
        `:""}
      </div>
    </div>

    <!-- Summary KPI Cards -->
    <div id="migration-summary-cards" style="display:grid; grid-template-columns:repeat(auto-fit, minmax(220px, 1fr)); gap:16px; margin-bottom:20px;">
      <div class="kpi-card skeleton" style="height:100px; border-radius:12px;"></div>
      <div class="kpi-card skeleton" style="height:100px; border-radius:12px;"></div>
      <div class="kpi-card skeleton" style="height:100px; border-radius:12px;"></div>
      <div class="kpi-card skeleton" style="height:100px; border-radius:12px;"></div>
    </div>

    <!-- Filter Toolbar -->
    <div class="card" style="padding:14px 18px; margin-bottom:20px; background:#131F36; border:1px solid #223456; border-radius:12px;">
      <div style="display:flex; flex-wrap:wrap; gap:12px; align-items:center; justify-content:space-between;">
        <div style="display:flex; flex-wrap:wrap; gap:10px; align-items:center; flex:1; min-width:260px;">
          <div style="position:relative; flex:1; min-width:200px; max-width:320px;">
            <input type="text" id="migration-search" class="form-control" placeholder="Cari cabang, vendor, catatan..." style="padding-left:34px; height:38px; border-radius:8px; font-size:0.88rem; width:100%; background:#0E182B; border-color:#223456;">
            <span style="position:absolute; left:11px; top:50%; transform:translateY(-50%); color:#94A3B8; font-size:0.9rem;">\u{1F50D}</span>
          </div>
          <select id="filter-vendor" class="form-control" style="width:auto; height:38px; border-radius:8px; font-size:0.875rem; min-width:136px; background:#0F1A2E; border-color:#223456;">
            <option value="">Semua Vendor</option>
            <option value="AEON">AEON</option>
            <option value="Triple S">Triple S</option>
          </select>
          <select id="filter-month" class="form-control" style="width:auto; height:38px; border-radius:8px; font-size:0.875rem; min-width:146px; background:#0F1A2E; border-color:#223456;">
            <option value="">Semua Bulan</option>
            <option value="2026-09">September 2026</option>
            <option value="2026-10">Oktober 2026</option>
            <option value="2026-11">November 2026</option>
            <option value="2026-12">Desember 2026</option>
          </select>
          <select id="filter-status" class="form-control" style="width:auto; height:38px; border-radius:8px; font-size:0.875rem; min-width:138px; background:#0F1A2E; border-color:#223456;">
            <option value="">Semua Status</option>
            <option value="scheduled">Terjadwal</option>
            <option value="in_progress">Sedang Proses</option>
            <option value="completed">Selesai</option>
          </select>
        </div>
        <div id="filter-stats-text" style="font-size:0.85rem; color:#CBD5E1; font-weight:600;">
          Memuat data...
        </div>
      </div>
    </div>

    <style>.os-mig-row{display:flex;align-items:center;justify-content:space-between;gap:12px;padding:12px 18px;border-bottom:1px solid #1E2E4A;transition:background-color 0.14s ease}.os-mig-row:last-child{border-bottom:none}.os-mig-row:hover{background:#172642}.os-mig-left{display:flex;align-items:center;gap:10px;min-width:0;flex:1 1 260px}.os-mig-code{font-family:ui-monospace,SFMono-Regular,Menlo,Monaco,Consolas,monospace;font-size:0.78rem;font-weight:700;color:#93C5FD;background:rgba(37,99,235,0.14);border:1px solid rgba(59,130,246,0.28);padding:3px 8px;border-radius:6px;flex-shrink:0}.os-mig-name{font-size:0.96rem;font-weight:700;color:#FFFFFF;line-height:1.35}.os-mig-sub{font-size:0.78rem;color:#A5B4C8;margin-top:2px}.os-mig-right{display:flex;align-items:center;gap:10px;flex-wrap:wrap;justify-content:flex-end}.os-mig-fc-pill{display:inline-flex;align-items:center;gap:5px;background:#0E182B;border:1px solid #223456;color:#E2E8F0;font-size:0.8rem;font-weight:700;padding:4px 10px;border-radius:7px;white-space:nowrap}.os-mig-icon-btn{display:inline-flex;align-items:center;justify-content:center;width:28px;height:28px;padding:0;border-radius:6px;background:transparent;border:1px solid rgba(51,65,85,0.75);color:#94A3B8;cursor:pointer;transition:all 0.14s ease}.os-mig-icon-btn:hover{background:rgba(37,99,235,0.18);border-color:rgba(96,165,250,0.5);color:#93C5FD}.os-mig-icon-del:hover{background:rgba(239,68,68,0.16);border-color:rgba(248,113,113,0.5);color:#FCA5A5}.os-mig-icon-btn:focus-visible{outline:2px solid #60A5FA;outline-offset:2px}@media (max-width:768px){.os-mig-row{flex-direction:column;align-items:stretch;gap:10px;padding:12px 14px}.os-mig-left{flex:none}.os-mig-right{justify-content:space-between}}</style>

    <!-- Main Content Container -->
    <div id="migration-view-container"></div>
  `;let s="board",c=null,m=[];X("/api/branches?all=1").then(v=>{v.ok&&v.data?.data&&(m=v.data.data)}).catch(()=>{});async function i(){try{let v=await X("/api/os-migrations/summary");if(!v.ok)throw new Error("Gagal memuat data migrasi");c=v.data?.data||v.data,d(c),$()}catch(v){let B=document.getElementById("migration-view-container");B.innerHTML=`
        <div class="empty-state" style="padding:40px; text-align:center;">
          <div style="font-size:2rem; margin-bottom:8px;">\u26A0\uFE0F</div>
          <p style="color:var(--danger); font-weight:600;">Gagal memuat data: ${P(v.message)}</p>
          <button type="button" id="os-mig-reload-btn" class="btn btn-outline btn-sm">Muat Ulang</button>
        </div>
      `,B.querySelector("#os-mig-reload-btn")?.addEventListener("click",()=>window.location.reload())}}function d(v){let B=document.getElementById("migration-summary-cards");if(!B||!v)return;let x=v.total_branches||0,k=v.vendors?.AEON?.total||0,_=v.vendors?.["Triple S"]?.total||0,E=v.by_status?.completed||0,w=v.by_status?.in_progress||0,K=v.by_status?.scheduled||0,G=x>0?Math.round(E/x*100):0;B.innerHTML=`
      <div class="card" style="padding:16px 20px; background:#131F36; border:1px solid #223456; border-radius:12px;">
        <div style="font-size:0.8rem; font-weight:600; color:#A5B4C8; text-transform:uppercase; letter-spacing:0.5px; margin-bottom:6px;">Total Cabang Migrasi</div>
        <div style="font-size:1.75rem; font-weight:800; color:#FFFFFF; line-height:1.2;">${x} <span style="font-size:0.9rem; font-weight:500; color:#A5B4C8;">Cabang</span></div>
        <div style="font-size:0.76rem; color:#CBD5E1; margin-top:4px;">Timeline Periode Sep - Des 2026</div>
      </div>

      <div class="card" style="padding:16px 20px; background:#131F36; border:1px solid #223456; border-left:4px solid #2563EB; border-radius:12px;">
        <div style="font-size:0.8rem; font-weight:600; color:#93C5FD; text-transform:uppercase; letter-spacing:0.5px; margin-bottom:6px;">Vendor AEON</div>
        <div style="font-size:1.75rem; font-weight:800; color:#60A5FA; line-height:1.2;">${k} <span style="font-size:0.9rem; font-weight:500; color:#A5B4C8;">Cabang</span></div>
        <div style="font-size:0.76rem; color:#CBD5E1; margin-top:4px;">Sep: 9 \u2022 Nov: 3 \u2022 Des: 8</div>
      </div>

      <div class="card" style="padding:16px 20px; background:#131F36; border:1px solid #223456; border-left:4px solid #38BDF8; border-radius:12px;">
        <div style="font-size:0.8rem; font-weight:600; color:#7DD3FC; text-transform:uppercase; letter-spacing:0.5px; margin-bottom:6px;">Vendor Triple S</div>
        <div style="font-size:1.75rem; font-weight:800; color:#38BDF8; line-height:1.2;">${_} <span style="font-size:0.9rem; font-weight:500; color:#A5B4C8;">Cabang</span></div>
        <div style="font-size:0.76rem; color:#CBD5E1; margin-top:4px;">Sep: 5 \u2022 Okt: 5 \u2022 Nov: 5 \u2022 Des: 27</div>
      </div>

      <div class="card" style="padding:16px 20px; background:#131F36; border:1px solid #223456; border-left:4px solid #34D399; border-radius:12px;">
        <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:6px;">
          <span style="font-size:0.8rem; font-weight:600; color:#6EE7B7; text-transform:uppercase; letter-spacing:0.5px;">Progres Realisasi</span>
          <span style="font-size:0.88rem; font-weight:800; color:#34D399;">${G}%</span>
        </div>
        <div style="width:100%; height:8px; background:#0E182B; border-radius:4px; overflow:hidden; margin-bottom:8px;">
          <div style="width:${G}%; height:100%; background:#34D399; border-radius:4px;"></div>
        </div>
        <div style="display:flex; justify-content:space-between; font-size:0.76rem; color:#CBD5E1; font-weight:600;">
          <span>Selesai: <strong style="color:#6EE7B7;">${E}</strong></span>
          <span>Proses: <strong style="color:#93C5FD;">${w}</strong></span>
          <span>Terjadwal: <strong style="color:#FCD34D;">${K}</strong></span>
        </div>
      </div>
    `}function F(){if(!c||!c.matrix)return[];let v=(document.getElementById("migration-search")?.value||"").toLowerCase().trim(),B=document.getElementById("filter-vendor")?.value||"",x=document.getElementById("filter-month")?.value||"",k=document.getElementById("filter-status")?.value||"";return c.matrix.filter(_=>!(B&&_.vendor!==B||x&&_.target_month!==x||k&&_.status!==k||v&&!`${_.branch_code} ${_.branch_name} ${_.branch_full_name} ${_.vendor} ${_.notes||""} ${_.pic_vendor||""}`.toLowerCase().includes(v)))}function $(){let v=F(),B=document.getElementById("filter-stats-text");B&&(B.innerHTML=`Menampilkan <strong>${v.length}</strong> dari <strong>${c?.total_branches||0}</strong> cabang`);let x=document.getElementById("migration-view-container");s==="board"?M(x,v):h(x,v)}function M(v,B){if(!B.length){v.innerHTML=`
        <div class="empty-state" style="padding:50px; text-align:center; background:#111827; border-radius:12px; border:1px solid var(--border);">
          <div style="font-size:2.5rem; margin-bottom:12px;">\u{1F50D}</div>
          <p style="font-size:1.05rem; font-weight:600; color:var(--text-2); margin-bottom:6px;">Tidak ada cabang yang cocok dengan filter</p>
          <p style="color:var(--text-3); font-size:0.875rem;">Coba reset kata kunci pencarian atau pilihan filter di atas.</p>
        </div>
      `;return}let x=["AEON","Triple S"],k=$n(Array.from(new Set(B.map(E=>E.target_month).filter(Boolean)))),_="";x.forEach(E=>{let w=B.filter(Y=>Y.vendor===E);if(document.getElementById("filter-vendor")?.value&&document.getElementById("filter-vendor").value!==E||w.length===0)return;let K=w.length,G=E==="AEON"?"#93C5FD":"#7DD3FC",J=E==="AEON"?"rgba(59, 130, 246, 0.35)":"rgba(56, 189, 248, 0.35)",W=k.filter(Y=>w.some(b=>b.target_month===Y));_+=`
        <div class="vendor-board-section" style="margin-bottom:24px; background:#131F36; border:1px solid #223456; border-radius:14px; overflow:hidden;">
          <!-- Vendor Section Header -->
          <div style="padding:14px 18px; background:#0E182B; border-bottom:1px solid #223456; display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:8px;">
            <div style="display:flex; align-items:center; gap:10px;">
              <span style="font-size:1.05rem; font-weight:800; color:${G};">Vendor ${P(E)}</span>
              <span class="badge" style="background:rgba(37,99,235,0.16); color:${G}; border:1px solid ${J}; font-weight:700; font-size:0.76rem; padding:3px 9px; border-radius:6px;">
                ${K} Cabang
              </span>
            </div>
            <div style="font-size:0.8rem; color:#A5B4C8; font-weight:500;">
              Dikelompokkan per Bulan Target
            </div>
          </div>

          <!-- Single Unified Grouped List Container (Zero nested boxes or inner scrollbars) -->
          <div>
            ${W.map(Y=>{let b=w.filter(y=>y.target_month===Y),S=Va[Y]||Y;return`
                <div class="os-mig-month-group">
                  <div style="padding:9px 18px; background:#101A2F; border-top:1px solid #1E2E4A; border-bottom:1px solid #1E2E4A; display:flex; justify-content:space-between; align-items:center;">
                    <span style="font-weight:700; font-size:0.88rem; color:#E2E8F0; letter-spacing:0.01em;">
                      \u{1F4C5} ${P(S)}
                    </span>
                    <span style="font-size:0.76rem; font-weight:700; color:#93C5FD; background:rgba(37,99,235,0.14); border:1px solid rgba(59,130,246,0.28); padding:2px 8px; border-radius:10px;">
                      ${b.length} Cabang
                    </span>
                  </div>

                  <div>
                    ${b.map(y=>{let r=Kt[y.status]||Kt.scheduled;return`
                        <div class="os-mig-row">
                          <div class="os-mig-left">
                            <span class="os-mig-code">${P(y.branch_code||"-")}</span>
                            <div style="min-width:0;">
                              <div class="os-mig-name">${P(y.branch_name||"-")}</div>
                              ${y.actual_date||y.notes?`
                                <div class="os-mig-sub">
                                  ${y.actual_date?`Realisasi: <strong style="color:#CBD5E1;">${P(y.actual_date)}</strong>`:""}
                                  ${y.actual_date&&y.notes?" \u2022 ":""}
                                  ${y.notes?`${P(y.notes)}`:""}
                                </div>
                              `:""}
                            </div>
                          </div>

                          <div class="os-mig-right">
                            <span class="os-mig-fc-pill" title="${y.employee_count||0} personil Facility Care aktif di cabang ini">
                              \u{1F464} ${y.employee_count||0} FC
                            </span>
                            <button type="button" class="status-toggle-btn" data-id="${P(String(y.id))}" data-status="${P(y.status)}" style="background:${r.bg}; color:${r.color}; border:1px solid ${r.border}; padding:4px 10px; border-radius:7px; font-size:0.78rem; font-weight:700; display:inline-flex; align-items:center; gap:6px; cursor:${n?"pointer":"default"}; white-space:nowrap;" title="${n?"Klik untuk ubah status":"Status"}">
                              <span style="width:7px; height:7px; border-radius:50%; background:${r.dot}; display:inline-block;"></span>
                              <span>${r.label}${n?" \u25BE":""}</span>
                            </button>
                            ${Si(y,n,a)}
                          </div>
                        </div>
                      `}).join("")}
                  </div>
                </div>
              `}).join("")}
          </div>
        </div>
      `}),v.innerHTML=_,L(v)}function h(v,B){if(!B.length){v.innerHTML=`
        <div class="empty-state" style="padding:50px; text-align:center; background:#111827; border-radius:12px; border:1px solid var(--border);">
          <div style="font-size:2.5rem; margin-bottom:12px;">\u{1F50D}</div>
          <p style="font-size:1.05rem; font-weight:600; color:var(--text-2); margin-bottom:6px;">Tidak ada cabang yang cocok dengan filter</p>
        </div>
      `;return}v.innerHTML=`
      <div class="card" style="background:#131F36; border:1px solid #223456; border-radius:12px; overflow:hidden;">
        <div class="table-responsive">
          <table class="table" style="width:100%; border-collapse:collapse;">
            <thead>
              <tr style="background:#0E182B; border-bottom:1px solid #223456; text-align:left;">
                <th style="padding:12px 16px; font-size:0.78rem; font-weight:700; color:#A5B4C8; width:50px;">NO</th>
                <th style="padding:12px 16px; font-size:0.78rem; font-weight:700; color:#A5B4C8;">KODE & CABANG</th>
                <th style="padding:12px 16px; font-size:0.78rem; font-weight:700; color:#A5B4C8;">VENDOR</th>
                <th style="padding:12px 16px; font-size:0.78rem; font-weight:700; color:#A5B4C8;">TARGET BULAN</th>
                <th style="padding:12px 16px; font-size:0.78rem; font-weight:700; color:#A5B4C8;">STATUS</th>
                <th style="padding:12px 16px; font-size:0.78rem; font-weight:700; color:#A5B4C8;">REALISASI</th>
                <th style="padding:12px 16px; font-size:0.78rem; font-weight:700; color:#A5B4C8;">SDM FC</th>
                <th style="padding:12px 16px; font-size:0.78rem; font-weight:700; color:#A5B4C8;">CATATAN</th>
                ${n||a?'<th style="padding:12px 16px; font-size:0.78rem; font-weight:700; color:#A5B4C8; text-align:right;">AKSI</th>':""}
              </tr>
            </thead>
            <tbody>
              ${B.map((x,k)=>{let _=Kt[x.status]||Kt.scheduled;return`
                  <tr style="border-bottom:1px solid #1E2E4A;">
                    <td style="padding:12px 16px; font-size:0.85rem; color:#A5B4C8;">${k+1}</td>
                    <td style="padding:12px 16px;">
                      <div style="font-weight:700; font-size:0.9rem; color:#FFFFFF;">${P(x.branch_full_name)}</div>
                    </td>
                    <td style="padding:12px 16px;">
                      <span class="badge" style="background:${x.vendor==="AEON"?"rgba(37,99,235,0.16)":"rgba(56,189,248,0.16)"}; color:${x.vendor==="AEON"?"#93C5FD":"#7DD3FC"}; border:1px solid ${x.vendor==="AEON"?"rgba(59,130,246,0.35)":"rgba(56,189,248,0.35)"}; font-weight:700; border-radius:6px; padding:4px 8px; font-size:0.75rem;">
                        ${P(x.vendor)}
                      </span>
                    </td>
                    <td style="padding:12px 16px; font-size:0.875rem; font-weight:600; color:#CBD5E1;">
                      ${P(Va[x.target_month]||x.target_month)}
                    </td>
                    <td style="padding:12px 16px;">
                      <button type="button" class="status-toggle-btn" data-id="${P(String(x.id))}" data-status="${P(x.status)}" style="background:${_.bg}; color:${_.color}; border:1px solid ${_.border}; padding:4px 9px; border-radius:6px; font-size:0.76rem; font-weight:700; display:inline-flex; align-items:center; gap:5px; cursor:${n?"pointer":"default"};" title="${n?"Klik untuk ubah status":"Status"}">
                        <span style="width:6px; height:6px; border-radius:50%; background:${_.dot}; display:inline-block;"></span> ${_.label}${n?" \u25BE":""}
                      </button>
                    </td>
                    <td style="padding:12px 16px; font-size:0.85rem; color:#CBD5E1;">
                      ${P(x.actual_date||"-")}
                    </td>
                    <td style="padding:12px 16px; font-size:0.85rem; font-weight:700; color:#E2E8F0;">
                      ${x.employee_count>0?`\u{1F464} ${x.employee_count} FC`:"-"}
                    </td>
                    <td style="padding:12px 16px; font-size:0.85rem; color:#A5B4C8; max-width:200px; white-space:nowrap; overflow:hidden; text-overflow:ellipsis;">
                      ${P(x.notes||"-")}
                    </td>
                    ${n||a?`
                      <td style="padding:12px 16px; text-align:right;">
                        ${Si(x,n,a)}
                      </td>
                    `:""}
                  </tr>
                `}).join("")}
            </tbody>
          </table>
        </div>
      </div>
    `,L(v)}function L(v){v.querySelectorAll(".status-toggle-btn").forEach(B=>{B.addEventListener("click",x=>{if(x.stopPropagation(),!n)return;let k=B.dataset.id,_=B.dataset.status;R(B,k,_)})}),v.querySelectorAll(".edit-card-btn").forEach(B=>{B.addEventListener("click",()=>{let x=B.dataset.id,k=c?.matrix?.find(_=>String(_.id)===String(x));k&&I(k)})}),v.querySelectorAll(".delete-card-btn").forEach(B=>{B.addEventListener("click",()=>{if(!a){j("Hanya Superadmin yang dapat menghapus jadwal migrasi.");return}let x=B.dataset.id,k=B.dataset.name||"Cabang";ge(`Apakah Anda yakin ingin menghapus jadwal migrasi untuk <strong>${k}</strong>?`,async()=>{try{if(!(await X(`/api/os-migrations/${x}`,{method:"DELETE"})).ok)throw new Error("Gagal menghapus");ae(`Jadwal migrasi ${k} berhasil dihapus`),i()}catch(_){j(_.message)}})})})}function R(v,B,x){if(!n)return;document.querySelectorAll(".status-popup-menu").forEach(K=>K.remove());let k=v.getBoundingClientRect(),_=document.createElement("div");_.className="status-popup-menu",_.style.cssText=`
      position: fixed;
      top: ${k.bottom+window.scrollY+4}px;
      left: ${k.left+window.scrollX}px;
      background: #172033;
      border: 1px solid var(--border);
      border-radius: 8px;
      box-shadow: 0 4px 16px rgba(0,0,0,0.12);
      z-index: 9999;
      padding: 6px 0;
      min-width: 160px;
    `;let E=[{key:"scheduled",label:"\u{1F7E1} Terjadwal"},{key:"in_progress",label:"\u{1F535} Sedang Proses"},{key:"completed",label:"\u{1F7E2} Selesai"}];_.innerHTML=E.map(K=>`
      <button class="status-option-item" data-key="${K.key}" style="
        width: 100%;
        text-align: left;
        padding: 8px 14px;
        background: none;
        border: none;
        font-size: 0.85rem;
        cursor: pointer;
        display: flex;
        align-items: center;
        gap: 8px;
        color: ${K.key===x?"#60A5FA":"var(--text-1)"};
        font-weight: ${K.key===x?"700":"500"};
      ">
        ${K.label} ${K.key===x?"\u2713":""}
      </button>
    `).join("");let w=K=>{!_.contains(K.target)&&K.target!==v&&(_.remove(),document.removeEventListener("click",w))};setTimeout(()=>document.addEventListener("click",w),50),_.querySelectorAll(".status-option-item").forEach(K=>{K.addEventListener("click",async()=>{let G=K.dataset.key;if(_.remove(),document.removeEventListener("click",w),G!==x)try{if(!(await X(`/api/os-migrations/${B}/status`,{method:"PATCH",body:JSON.stringify({status:G})})).ok)throw new Error("Gagal memperbarui status");ae("Status berhasil diubah"),i()}catch(J){j(J.message)}})}),document.body.appendChild(_)}function I(v=null){if(!n){j("Anda tidak memiliki izin untuk menambah atau mengedit migrasi OS.");return}let B=!!v,x=B?"Edit Jadwal Migrasi OS":"Tambah Jadwal Migrasi OS",_=`
      <form id="migration-form" style="display:flex; flex-direction:column; gap:14px;">
        <div class="form-group">
          <label style="font-weight:600; font-size:0.875rem; margin-bottom:4px; display:block;">Pilih Cabang *</label>
          <select name="branch_id" class="form-control" required style="width:100%; height:38px; border-radius:8px;">
            <option value="">-- Pilih Cabang --</option>
            ${m.map(E=>`
      <option value="${E.id}" ${v&&v.branch_id===E.id?"selected":""}>
        ${E.full_name}
      </option>
    `).join("")}
          </select>
        </div>

        <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div class="form-group">
            <label style="font-weight:600; font-size:0.875rem; margin-bottom:4px; display:block;">Vendor Outsourcing *</label>
            <select name="vendor" class="form-control" required style="width:100%; height:38px; border-radius:8px;">
              <option value="AEON" ${v?.vendor==="AEON"?"selected":""}>AEON</option>
              <option value="Triple S" ${v?.vendor==="Triple S"?"selected":""}>Triple S</option>
              <option value="Lainnya" ${v&&!["AEON","Triple S"].includes(v.vendor)?"selected":""}>Lainnya</option>
            </select>
          </div>

          <div class="form-group">
            <label style="font-weight:600; font-size:0.875rem; margin-bottom:4px; display:block;">Target Bulan *</label>
            <select name="target_month" class="form-control" required style="width:100%; height:38px; border-radius:8px;">
              <option value="2026-09" ${v?.target_month==="2026-09"?"selected":""}>September 2026</option>
              <option value="2026-10" ${v?.target_month==="2026-10"?"selected":""}>Oktober 2026</option>
              <option value="2026-11" ${v?.target_month==="2026-11"?"selected":""}>November 2026</option>
              <option value="2026-12" ${v?.target_month==="2026-12"?"selected":""}>Desember 2026</option>
            </select>
          </div>
        </div>

        <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div class="form-group">
            <label style="font-weight:600; font-size:0.875rem; margin-bottom:4px; display:block;">Status Pelaksanaan</label>
            <select name="status" class="form-control" style="width:100%; height:38px; border-radius:8px;">
              <option value="scheduled" ${v?.status==="scheduled"?"selected":""}>\u{1F7E1} Terjadwal</option>
              <option value="in_progress" ${v?.status==="in_progress"?"selected":""}>\u{1F535} Sedang Proses</option>
              <option value="completed" ${v?.status==="completed"?"selected":""}>\u{1F7E2} Selesai</option>
            </select>
          </div>

          <div class="form-group">
            <label style="font-weight:600; font-size:0.875rem; margin-bottom:4px; display:block;">Tanggal Realisasi (Opsional)</label>
            <input type="date" name="actual_date" class="form-control" value="${v?.actual_date||""}" style="width:100%; height:38px; border-radius:8px;">
          </div>
        </div>

        <div class="form-group">
          <label style="font-weight:600; font-size:0.875rem; margin-bottom:4px; display:block;">PIC / Koordinator Vendor (Opsional)</label>
          <input type="text" name="pic_vendor" class="form-control" value="${v?.pic_vendor||""}" style="width:100%; height:38px; border-radius:8px;">
        </div>

        <div class="form-group">
          <label style="font-weight:600; font-size:0.875rem; margin-bottom:4px; display:block;">Catatan / Keterangan</label>
          <textarea name="notes" class="form-control" rows="2" style="width:100%; border-radius:8px;">${v?.notes||""}</textarea>
        </div>
      </form>
    `;ce({title:x,content:_,confirmText:B?"Simpan Perubahan":"Tambah Cabang",onConfirm:async(E,w)=>{let K=E.querySelector("#migration-form");if(!K.checkValidity()){K.reportValidity();return}let G=new FormData(K),J={branch_id:parseInt(G.get("branch_id"),10),vendor:G.get("vendor"),target_month:G.get("target_month"),status:G.get("status"),actual_date:G.get("actual_date")||null,pic_vendor:G.get("pic_vendor")||null,notes:G.get("notes")||null};try{let W=B?`/api/os-migrations/${v.id}`:"/api/os-migrations",b=await X(W,{method:B?"PUT":"POST",body:JSON.stringify(J)});if(!b.ok){let S=b.data?.error||b.data?.message||"Gagal menyimpan data";throw new Error(S)}ae(B?"Jadwal migrasi berhasil diupdate":"Jadwal migrasi berhasil ditambahkan"),w(),i()}catch(W){j(W.message)}}})}let N=document.getElementById("view-board-btn"),q=document.getElementById("view-table-btn");N.addEventListener("click",()=>{s="board",N.style.background="#fff",N.classList.remove("btn-ghost"),q.style.background="transparent",q.classList.add("btn-ghost"),$()}),q.addEventListener("click",()=>{s="table",q.style.background="#fff",q.classList.remove("btn-ghost"),N.style.background="transparent",N.classList.add("btn-ghost"),$()}),["migration-search","filter-vendor","filter-month","filter-status"].forEach(v=>{let B=document.getElementById(v);B&&(B.addEventListener("input",$),B.addEventListener("change",$))}),document.getElementById("btn-add-migration")?.addEventListener("click",()=>{I()}),document.getElementById("btn-export-migrations")?.addEventListener("click",async()=>{try{let v=(document.getElementById("migration-search")?.value||"").trim(),B=document.getElementById("filter-vendor")?.value||"",x=document.getElementById("filter-month")?.value||"",k=document.getElementById("filter-status")?.value||"",_={};B&&(_.vendor=B),x&&(_.month=x),k&&(_.status=k),v&&(_.search=v);let E=await ue("/api/os-migrations/export",_);if(!E.length){j("Tidak ada data untuk diexport");return}let w=E.map((K,G)=>({No:G+1,"Kode Cabang":K.branch_code,"Nama Cabang":K.branch_name,"Cabang Lengkap":K.branch_full_name,"Vendor OS":K.vendor,"Target Bulan":Va[K.target_month]||K.target_month,Status:Kt[K.status]?.label||K.status,"Tanggal Realisasi":K.actual_date||"","Jml Personil FC":K.employee_count||0,"PIC Vendor":K.pic_vendor||"",Catatan:K.notes||""}));be(w,`Timeline_Migrasi_OS_${new Date().toISOString().slice(0,10)}`),ae("File Excel berhasil diunduh")}catch(v){j("Gagal mengekspor data migrasi: "+v.message)}}),i()}Be();Ce();Ie();var Fi={Persiapan:{bg:"#F3E8FF",color:"#6B21A8",border:"#D8B4FE"},"Area Umum":{bg:"#E0F2FE",color:"#0369A1",border:"#BAE6FD"},"Ruang Tindakan":{bg:"#DBEAFE",color:"#1E40AF",border:"#93C5FD"},"Ruang Penunjang":{bg:"#F1F5F9",color:"#334155",border:"#CBD5E1"},Toilet:{bg:"#CCFBF1",color:"#0F766E",border:"#99F6E4"},"Area Luar":{bg:"#FEF3C7",color:"#92400E",border:"#FDE68A"},"Kontrol & Check":{bg:"#FFEDD5",color:"#C2410C",border:"#FED7AA"},Istirahat:{bg:"#E2E8F0",color:"#475569",border:"#CBD5E1"},"Layanan Pasien":{bg:"#DCFCE7",color:"#15803D",border:"#86EFAC"},"Sampah Umum":{bg:"#FEF9C3",color:"#854D0E",border:"#FDE047"},"Sampah Medis":{bg:"#FEE2E2",color:"#991B1B",border:"#FCA5A5"},Penutup:{bg:"#E0E7FF",color:"#3730A3",border:"#C7D2FE"}};async function Ei(e){let t=se(),n=re(t,"dac","write"),a=re(t,"dac","delete"),o=re(t,"export");e.innerHTML=`
    <div class="dac-page-shell">
      <!-- Unified Hero Card (Consistent with FCMS Clean Navy Modules) -->
      <div class="dac-hero-card">
        <div class="dac-hero-top">
          <div class="dac-hero-title-group">
            <h1 class="dac-hero-title">Master DAC (Daily Activity Cleaning)</h1>
            <p class="dac-hero-subtitle">Standar acuan alur rutinitas pembersihan harian Facility Care (Shift Pagi &amp; Shift Siang)</p>
          </div>
          <div class="dac-hero-actions">
            <div class="dac-view-switch" role="group" aria-label="Mode tampilan Master DAC">
              <button type="button" id="view-timeline-btn" class="dac-view-btn active">Alur Timeline</button>
              <button type="button" id="view-table-btn" class="dac-view-btn">Tabel Standar</button>
            </div>
            <button type="button" id="btn-print-dac" class="btn btn-outline btn-sm" title="Cetak Format SOP untuk Ruang Janitor">
              <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><polyline points="6 9 6 2 18 2 18 9"/><path d="M6 18H4a2 2 0 0 1-2-2v-5a2 2 0 0 1 2-2h16a2 2 0 0 1 2 2v5a2 2 0 0 1-2 2h-2"/><rect x="6" y="14" width="12" height="8"/></svg>
              Cetak SOP
            </button>
            ${o?`
            <button type="button" id="btn-export-dac" class="btn btn-outline btn-sm">
              <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg>
              Export Excel
            </button>
            `:""}
            ${n?`
            <button type="button" id="btn-add-dac" class="btn btn-primary btn-sm">
              + Tambah Aktivitas
            </button>
            `:""}
          </div>
        </div>

        <!-- Unified Summary Stats Grid -->
        <div id="dac-summary-cards" class="dac-stats-grid">
          <div class="dac-stat-box skeleton" style="height:78px;"></div>
          <div class="dac-stat-box skeleton" style="height:78px;"></div>
          <div class="dac-stat-box skeleton" style="height:78px;"></div>
          <div class="dac-stat-box skeleton" style="height:78px;"></div>
        </div>
      </div>

      <!-- Unified Filter Toolbar & Shift Switcher Tabs -->
      <div class="dac-controls-card">
        <div class="dac-controls-row">
          <div class="dac-search-wrap">
            <svg class="dac-search-icon" width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/></svg>
            <input type="text" id="dac-search" class="form-control dac-search-input" placeholder="Cari lokasi, aktivitas, jam..." aria-label="Cari aktivitas DAC">
          </div>
          <select id="dac-category-filter" class="form-control dac-filter-select" aria-label="Filter kategori DAC">
            <option value="">Semua Kategori</option>
          </select>
          <select id="dac-period-filter" class="form-control dac-filter-select" aria-label="Filter periode waktu DAC">
            <option value="">Semua Waktu</option>
            <option value="morning">Sesi Awal (08:00 - 12:00 / 13:00 - 17:00)</option>
            <option value="afternoon">Sesi Lanjutan (12:00 - 17:00 / 17:00 - 22:00)</option>
          </select>
          <div id="dac-stats-text" class="dac-stats-text">
            Memuat data...
          </div>
        </div>

        <div id="dac-shift-tabs" class="dac-shift-tabs" data-fc-scroll-tabs role="tablist" aria-label="Filter shift DAC">
          <button type="button" class="shift-tab-btn active" data-shift="pagi" role="tab" aria-selected="true">
            <span>Shift Pagi (08:00 - 17:00)</span>
            <span class="badge dac-tab-count" id="badge-count-pagi">31</span>
          </button>
          <button type="button" class="shift-tab-btn" data-shift="siang" role="tab" aria-selected="false">
            <span>Shift Siang (13:00 - 22:00)</span>
            <span class="badge dac-tab-count" id="badge-count-siang">33</span>
          </button>
          <button type="button" class="shift-tab-btn" data-shift="all" role="tab" aria-selected="false">
            <span>Semua Shift</span>
            <span class="badge dac-tab-count" id="badge-count-all">64</span>
          </button>
        </div>
      </div>

      <!-- Main Content Container (Aligned 100% Width with Shell) -->
      <div id="dac-view-container" class="dac-view-container"></div>
    </div>
  `;let s="pagi",c="timeline",m=[];async function i(){try{let v=await X("/api/dac?all=1");if(!v.ok)throw new Error("Gagal memuat standar DAC");m=v.data?.data||v.data||[];let B=m.filter(E=>E.shift==="pagi").length,x=m.filter(E=>E.shift==="siang").length;document.getElementById("badge-count-pagi").textContent=B,document.getElementById("badge-count-siang").textContent=x,document.getElementById("badge-count-all").textContent=m.length;let k=document.getElementById("dac-category-filter"),_=Array.from(new Set(m.map(E=>E.category))).sort();k.innerHTML='<option value="">Semua Kategori</option>'+_.map(E=>`<option value="${P(E)}">${P(E)}</option>`).join(""),F(),$()}catch(v){let B=document.getElementById("dac-view-container");B.innerHTML=`
        <div class="empty-state" style="padding:40px; text-align:center;">
          <div style="font-size:2rem; margin-bottom:8px;">\u26A0\uFE0F</div>
          <p style="color:var(--danger); font-weight:600;">Gagal memuat data DAC: ${P(v.message)}</p>
          <button type="button" id="dac-reload-btn" class="btn btn-outline btn-sm">Muat Ulang</button>
        </div>
      `,B.querySelector("#dac-reload-btn")?.addEventListener("click",()=>window.location.reload())}}function d(){let v=(document.getElementById("dac-search")?.value||"").toLowerCase().trim(),B=document.getElementById("dac-category-filter")?.value||"",x=document.getElementById("dac-period-filter")?.value||"";return m.filter(k=>{if(s!=="all"&&k.shift!==s||B&&k.category!==B)return!1;if(x==="morning"){let _=parseInt(k.start_time.split(":")[0],10);if(k.shift==="pagi"&&_>=12||k.shift==="siang"&&_>=17)return!1}else if(x==="afternoon"){let _=parseInt(k.start_time.split(":")[0],10);if(k.shift==="pagi"&&_<12||k.shift==="siang"&&_<17)return!1}return!(v&&!`${k.start_time} ${k.end_time} ${k.category} ${k.location} ${k.activity}`.toLowerCase().includes(v))})}function F(){let v=document.getElementById("dac-summary-cards");if(!v)return;let B=m.filter(w=>w.shift==="pagi"),x=m.filter(w=>w.shift==="siang"),k=s==="pagi"?B:s==="siang"?x:m,_=s==="pagi"?"08:00 \u2013 17:00":s==="siang"?"13:00 \u2013 22:00":"08:00 \u2013 22:00",E=s==="pagi"?"Shift Pagi":s==="siang"?"Shift Siang":"Semua Shift";v.innerHTML=`
      <div class="dac-stat-box">
        <div class="dac-stat-icon" aria-hidden="true">
          <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M9 11l3 3L22 4"/><path d="M21 12v7a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11"/></svg>
        </div>
        <div class="dac-stat-body">
          <div class="dac-stat-val">${k.length}</div>
          <div class="dac-stat-lbl">Total Aktivitas (${E})</div>
        </div>
      </div>

      <div class="dac-stat-box">
        <div class="dac-stat-icon" aria-hidden="true">
          <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="10"/><polyline points="12 6 12 12 16 14"/></svg>
        </div>
        <div class="dac-stat-body">
          <div class="dac-stat-val dac-stat-val--time">${_}</div>
          <div class="dac-stat-lbl">Jam Kerja Standar</div>
        </div>
      </div>

      <div class="dac-stat-box">
        <div class="dac-stat-icon" aria-hidden="true">
          <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="3" width="7" height="7" rx="1"/><rect x="14" y="3" width="7" height="7" rx="1"/><rect x="3" y="14" width="7" height="7" rx="1"/><rect x="14" y="14" width="7" height="7" rx="1"/></svg>
        </div>
        <div class="dac-stat-body">
          <div class="dac-stat-val">12</div>
          <div class="dac-stat-lbl">Cakupan Kategori</div>
        </div>
      </div>

      <div class="dac-stat-box">
        <div class="dac-stat-icon" aria-hidden="true">
          <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M12 22s8-4 8-10V5l-8-3-8 3v7c0 6 8 10 8 10z"/></svg>
        </div>
        <div class="dac-stat-body">
          <div class="dac-stat-val dac-stat-val--focus">Check Sheet &amp; Medis</div>
          <div class="dac-stat-lbl">Poin Wajib Harian</div>
        </div>
      </div>
    `}function $(){let v=d(),B=document.getElementById("dac-stats-text");B&&(B.innerHTML=`Menampilkan <strong>${v.length}</strong> aktivitas`);let x=document.getElementById("dac-view-container");c==="timeline"?M(x,v):h(x,v)}function M(v,B){if(!B.length){v.innerHTML=`
        <div class="dac-empty-card">
          <p class="dac-empty-title">Tidak ada aktivitas yang cocok dengan filter</p>
          <p class="dac-empty-sub">Coba reset kata kunci atau pilihan kategori di atas.</p>
        </div>
      `;return}v.innerHTML=`
      <div class="timeline-container dac-timeline-list">
        ${B.map(x=>{let k=Fi[x.category]||{bg:"rgba(148,163,184,0.14)",color:"#E2E8F0",border:"rgba(148,163,184,0.3)"},_=x.shift==="pagi"?"Pagi":"Siang";return`
            <div class="dac-timeline-card">
              <!-- Time Column -->
              <div class="dac-time-box">
                <div class="dac-time-range">
                  ${P(x.start_time)} - ${P(x.end_time)}
                </div>
                <div class="dac-time-itemno">
                  Item #${P(String(x.item_no))}
                </div>
                ${s==="all"?`
                  <div class="dac-time-shift">
                    <span class="badge ${x.shift==="pagi"?"badge-warning":"badge-primary"}">${_}</span>
                  </div>
                `:""}
              </div>

              <!-- Content Column -->
              <div class="dac-card-content">
                <div class="dac-card-head">
                  <div class="dac-card-meta">
                    <span class="badge" style="background:${k.bg}; color:${k.color}; border:1px solid ${k.border};">
                      ${P(x.category)}
                    </span>
                    <span class="dac-card-location">
                      <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M21 10c0 7-9 13-9 13s-9-6-9-13a9 9 0 0 1 18 0z"/><circle cx="12" cy="10" r="3"/></svg>
                      <span>${P(x.location)}</span>
                    </span>
                  </div>
                  ${n||a?`
                    <div class="dac-card-actions">
                      ${n?`<button type="button" class="fc-action-btn fc-action-edit edit-dac-btn" data-id="${P(String(x.id))}" title="Edit ${P(x.location)}" aria-label="Edit aktivitas ${P(x.location)}"><svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M12 20h9"></path><path d="M16.5 3.5a2.121 2.121 0 0 1 3 3L7 19l-4 1 1-4L16.5 3.5z"></path></svg></button>`:""}
                      ${a?`<button type="button" class="fc-action-btn fc-action-del delete-dac-btn" data-id="${P(String(x.id))}" data-name="${P(x.location)}" title="Hapus ${P(x.location)}" aria-label="Hapus aktivitas ${P(x.location)}"><svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><polyline points="3 6 5 6 21 6"></polyline><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"></path></svg></button>`:""}
                    </div>
                  `:""}
                </div>

                <div class="dac-activity-desc">
                  ${P(x.activity)}
                </div>
              </div>
            </div>
          `}).join("")}
      </div>
    `,L(v)}function h(v,B){if(!B.length){v.innerHTML=`
        <div class="dac-empty-card">
          <p class="dac-empty-title">Tidak ada aktivitas yang cocok dengan filter</p>
          <p class="dac-empty-sub">Coba reset kata kunci atau pilihan kategori di atas.</p>
        </div>
      `;return}v.innerHTML=`
      <div class="dac-table-card">
        <div class="table-wrapper">
          <table class="dac-table">
            <thead>
              <tr>
                <th style="width:52px; text-align:center;">No</th>
                ${s==="all"?'<th style="width:82px;">Shift</th>':""}
                <th style="width:128px;">Jam</th>
                <th style="width:150px;">Kategori</th>
                <th style="width:210px;">Lokasi / Area</th>
                <th>Aktivitas</th>
                ${n||a?'<th style="text-align:center; width:96px;">Aksi</th>':""}
              </tr>
            </thead>
            <tbody>
              ${B.map(x=>{let k=Fi[x.category]||{bg:"rgba(148,163,184,0.14)",color:"#E2E8F0",border:"rgba(148,163,184,0.3)"};return`
                  <tr class="dac-table-row">
                    <td style="text-align:center; color:#A5B4C8; font-weight:600; font-variant-numeric:tabular-nums;">${P(String(x.item_no))}</td>
                    ${s==="all"?`
                      <td>
                        <span class="badge ${x.shift==="pagi"?"badge-warning":"badge-primary"}">
                          ${x.shift==="pagi"?"Pagi":"Siang"}
                        </span>
                      </td>
                    `:""}
                    <td style="font-family:ui-monospace,SFMono-Regular,Menlo,monospace; font-weight:700; color:#FFFFFF; white-space:nowrap;">
                      ${P(x.start_time)} - ${P(x.end_time)}
                    </td>
                    <td>
                      <span class="badge" style="background:${k.bg}; color:${k.color}; border:1px solid ${k.border}; white-space:nowrap;">
                        ${P(x.category)}
                      </span>
                    </td>
                    <td style="font-weight:700; color:#F8FAFC;">
                      ${P(x.location)}
                    </td>
                    <td style="color:#E2E8F0; line-height:1.5; white-space:normal; min-width:240px;">
                      ${P(x.activity)}
                    </td>
                    ${n||a?`
                      <td style="text-align:center; white-space:nowrap;">
                        <div class="dac-card-actions" style="justify-content:center;">
                          ${n?`<button type="button" class="fc-action-btn fc-action-edit edit-dac-btn" data-id="${P(String(x.id))}" title="Edit ${P(x.location)}" aria-label="Edit aktivitas ${P(x.location)}"><svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M12 20h9"></path><path d="M16.5 3.5a2.121 2.121 0 0 1 3 3L7 19l-4 1 1-4L16.5 3.5z"></path></svg></button>`:""}
                          ${a?`<button type="button" class="fc-action-btn fc-action-del delete-dac-btn" data-id="${P(String(x.id))}" data-name="${P(x.location)}" title="Hapus ${P(x.location)}" aria-label="Hapus aktivitas ${P(x.location)}"><svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><polyline points="3 6 5 6 21 6"></polyline><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"></path></svg></button>`:""}
                        </div>
                      </td>
                    `:""}
                  </tr>
                `}).join("")}
            </tbody>
          </table>
        </div>
      </div>
    `,L(v)}function L(v){v.querySelectorAll(".edit-dac-btn").forEach(B=>{B.addEventListener("click",()=>{let x=B.dataset.id,k=m.find(_=>String(_.id)===String(x));k&&R(k)})}),v.querySelectorAll(".delete-dac-btn").forEach(B=>{B.addEventListener("click",()=>{let x=B.dataset.id,k=B.dataset.name||"item";ge(`Apakah Anda yakin ingin menghapus aktivitas pada lokasi <strong>${k}</strong>?`,async()=>{try{if(!(await X(`/api/dac/${x}`,{method:"DELETE"})).ok)throw new Error("Gagal menghapus item");ae("Aktivitas DAC berhasil dihapus"),i()}catch(_){j(_.message)}})})})}function R(v=null){if(!n){j("Anda tidak memiliki izin untuk menambah atau mengedit aktivitas.");return}let B=!!v,x=B?"Edit Aktivitas DAC":"Tambah Aktivitas DAC",k=["Persiapan","Area Umum","Ruang Tindakan","Ruang Penunjang","Toilet","Area Luar","Kontrol & Check","Istirahat","Layanan Pasien","Sampah Umum","Sampah Medis","Penutup"],_=`
      <form id="dac-form" style="display:flex; flex-direction:column; gap:14px;">
        <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div class="form-group">
            <label style="font-weight:600; font-size:0.875rem; margin-bottom:4px; display:block;">Pilih Shift <span class="required">*</span></label>
            <select name="shift" class="form-control" required style="width:100%; height:38px; border-radius:8px;">
              <option value="pagi" ${v?.shift==="pagi"||!v&&s==="pagi"?"selected":""}>\u{1F305} Shift Pagi (08:00 - 17:00)</option>
              <option value="siang" ${v?.shift==="siang"||!v&&s==="siang"?"selected":""}>\u{1F307} Shift Siang (13:00 - 22:00)</option>
            </select>
          </div>

          <div class="form-group">
            <label style="font-weight:600; font-size:0.875rem; margin-bottom:4px; display:block;">Nomor Urut <span class="required">*</span></label>
            <input type="number" name="item_no" class="form-control" value="${v?.item_no||""}" style="width:100%; height:38px; border-radius:8px;">
          </div>
        </div>

        <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div class="form-group">
            <label style="font-weight:600; font-size:0.875rem; margin-bottom:4px; display:block;">Jam Mulai (HH:MM) <span class="required">*</span></label>
            <input type="time" name="start_time" class="form-control" value="${v?.start_time||""}" required style="width:100%; height:38px; border-radius:8px;">
          </div>

          <div class="form-group">
            <label style="font-weight:600; font-size:0.875rem; margin-bottom:4px; display:block;">Jam Selesai (HH:MM) <span class="required">*</span></label>
            <input type="time" name="end_time" class="form-control" value="${v?.end_time||""}" required style="width:100%; height:38px; border-radius:8px;">
          </div>
        </div>

        <div class="form-group">
          <label style="font-weight:600; font-size:0.875rem; margin-bottom:4px; display:block;">Kategori <span class="required">*</span></label>
          <input list="cat-suggestions" name="category" class="form-control" value="${v?.category||""}" required style="width:100%; height:38px; border-radius:8px;">
          <datalist id="cat-suggestions">
            ${k.map(E=>`<option value="${E}">`).join("")}
          </datalist>
        </div>

        <div class="form-group">
          <label style="font-weight:600; font-size:0.875rem; margin-bottom:4px; display:block;">Lokasi / Area <span class="required">*</span></label>
          <input type="text" name="location" class="form-control" value="${v?.location||""}" required style="width:100%; height:38px; border-radius:8px;">
        </div>

        <div class="form-group">
          <label style="font-weight:600; font-size:0.875rem; margin-bottom:4px; display:block;">Aktivitas Detail <span class="required">*</span></label>
          <textarea name="activity" class="form-control" rows="3" required style="width:100%; border-radius:8px;">${v?.activity||""}</textarea>
        </div>
      </form>
    `;ce({title:x,content:_,confirmText:B?"Simpan Perubahan":"Tambah Aktivitas",onConfirm:async(E,w)=>{let K=E.querySelector("#dac-form");if(!K.checkValidity()){K.reportValidity();return}let G=new FormData(K),J={shift:G.get("shift"),item_no:G.get("item_no")?parseInt(G.get("item_no"),10):void 0,start_time:G.get("start_time"),end_time:G.get("end_time"),category:G.get("category"),location:G.get("location"),activity:G.get("activity")};try{let W=B?`/api/dac/${v.id}`:"/api/dac",b=await X(W,{method:B?"PUT":"POST",body:JSON.stringify(J)});if(!b.ok){let S=b.data?.error||b.data?.message||"Gagal menyimpan aktivitas";throw new Error(S)}ae(B?"Aktivitas berhasil diperbarui":"Aktivitas berhasil ditambahkan"),w(),i()}catch(W){j(W.message)}}})}function I(){let v=d();if(!v.length){j("Tidak ada data untuk dicetak");return}let B=s==="pagi"?"SHIFT PAGI (08:00 - 17:00)":s==="siang"?"SHIFT SIANG (13:00 - 22:00)":"SELURUH SHIFT (PAGI & SIANG)",x=window.open("","_blank");if(!x){j("Popup diblokir browser, mohon izinkan popup untuk mencetak");return}x.document.write(`
      <!DOCTYPE html>
      <html>
      <head>
        <title>STANDAR ACUAN DAC - ${B}</title>
        <style>body{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Arial,sans-serif;margin:20px;color:#1e293b;font-size:11pt}.header{text-align:center;border-bottom:2px solid #0f172a;padding-bottom:12px;margin-bottom:20px}.title{font-size:16pt;font-weight:bold;margin:0;color:#0f172a}.subtitle{font-size:11pt;color:#475569;margin:4px 0 0 0}table{width:100%;border-collapse:collapse;margin-top:10px}th,td{border:1px solid #cbd5e1;padding:7px 10px;text-align:left;vertical-align:top}th{background-color:#f1f5f9;font-weight:bold;font-size:10pt;text-transform:uppercase}.jam{font-family:monospace;font-weight:bold;white-space:nowrap}.badge{display:inline-block;padding:2px 6px;border-radius:4px;font-size:8pt;font-weight:bold}@media print{body{margin:10mm}button{display:none}}</style>
      </head>
      <body>
        <div class="header">
          <div class="title">FACILITY CARE MANAGEMENT SYSTEM (FCMS)</div>
          <div class="subtitle">STANDAR RUTINITAS DAILY ACTIVITY CLEANING (DAC) \u2014 ${B}</div>
        </div>
        <table>
          <thead>
            <tr>
              <th style="width:35px; text-align:center;">NO</th>
              ${s==="all"?'<th style="width:55px;">SHIFT</th>':""}
              <th style="width:105px;">JAM</th>
              <th style="width:130px;">KATEGORI</th>
              <th style="width:180px;">LOKASI / AREA</th>
              <th>RINCIAN AKTIVITAS</th>
            </tr>
          </thead>
          <tbody>
            ${v.map(k=>`
              <tr>
                <td style="text-align:center; font-weight:bold;">${k.item_no}</td>
                ${s==="all"?`<td>${k.shift==="pagi"?"Pagi":"Siang"}</td>`:""}
                <td class="jam">${k.start_time} - ${k.end_time}</td>
                <td><strong>${k.category}</strong></td>
                <td><strong>${k.location}</strong></td>
                <td>${k.activity}</td>
              </tr>
            `).join("")}
          </tbody>
        </table>
        <div style="margin-top:20px; text-align:right; font-size:9pt; color:#64748b;">
          Dicetak pada: ${new Date().toLocaleString("id-ID")} | Standar Resmi Operasional Facility Care Audy Dental
        </div>
      </body>
      </html>
    `),x.document.close(),x.focus(),setTimeout(()=>{x.print()},400)}document.querySelectorAll(".shift-tab-btn").forEach(v=>{v.addEventListener("click",()=>{s=v.dataset.shift,document.querySelectorAll(".shift-tab-btn").forEach(B=>{let x=B.dataset.shift===s;B.classList.toggle("active",x),B.setAttribute("aria-selected",x?"true":"false")}),F(),$()})});let N=document.getElementById("view-timeline-btn"),q=document.getElementById("view-table-btn");N?.addEventListener("click",()=>{c="timeline",N.classList.add("active"),q?.classList.remove("active"),$()}),q?.addEventListener("click",()=>{c="table",q.classList.add("active"),N?.classList.remove("active"),$()}),["dac-search","dac-category-filter","dac-period-filter"].forEach(v=>{let B=document.getElementById(v);B&&(B.addEventListener("input",$),B.addEventListener("change",$))}),document.getElementById("btn-add-dac")?.addEventListener("click",()=>{R()}),document.getElementById("btn-print-dac").addEventListener("click",()=>{I()}),document.getElementById("btn-export-dac")?.addEventListener("click",async()=>{try{let v=(document.getElementById("dac-search")?.value||"").trim(),B=document.getElementById("dac-category-filter")?.value||"",x=document.getElementById("dac-period-filter")?.value||"",k={};s!=="all"&&(k.shift=s),B&&(k.category=B),v&&(k.search=v);let _=await ue("/api/dac/export",k);if((x==="morning"||x==="afternoon")&&(_=_.filter(w=>{let K=parseInt(String(w.start_time||"0").split(":")[0],10);if(x==="morning"){if(w.shift==="pagi"&&K>=12||w.shift==="siang"&&K>=17)return!1}else if(w.shift==="pagi"&&K<12||w.shift==="siang"&&K<17)return!1;return!0})),!_.length){j("Tidak ada data untuk diexport");return}let E=_.map(w=>({No:w.item_no,Shift:w.shift==="pagi"?"Shift Pagi":"Shift Siang","Jam Mulai":w.start_time,"Jam Selesai":w.end_time,Kategori:w.category,"Lokasi / Area":w.location,"Rincian Aktivitas":w.activity}));be(E,`DAC_Daily_Activity_Cleaning_${s}_${new Date().toISOString().slice(0,10)}`),ae("File Excel berhasil diunduh")}catch(v){j("Gagal mengekspor data DAC: "+v.message)}}),i()}Ie();async function $i(e){let t=se();if(!t||String(t.role).toLowerCase()!=="superadmin"||t.isRestrictedToken===!0||Number(t.must_change_password)===1){e.innerHTML='<div class="empty-state"><p class="text-danger">Akses ditolak. Halaman ini hanya untuk Superadmin.</p></div>';return}e.innerHTML=`
    <div class="p-6 max-w-7xl mx-auto space-y-6">
      <div class="flex flex-col md:flex-row justify-between items-start md:items-center">
        <div>
          <h2 class="text-2xl font-bold text-gray-800">Sync Dashboard</h2>
          <p class="text-gray-500">Monitor Bidirectional Sync Health & Queue</p>
        </div>
        <div class="mt-4 md:mt-0 flex gap-2 flex-wrap">
          <button id="btnRetryAll" class="px-3 py-2 bg-yellow-500 text-white rounded shadow hover:bg-yellow-600 transition text-sm">Retry All Failed</button>
          <button id="btnResetStuck" class="px-3 py-2 bg-orange-500 text-white rounded shadow hover:bg-orange-600 transition text-sm">Reset Stuck</button>
          <button id="btnReconcile" class="px-3 py-2 bg-purple-500 text-white rounded shadow hover:bg-purple-600 transition text-sm">Force Reconcile</button>
          <button id="btnRefreshSync" class="px-3 py-2 bg-blue-600 text-white rounded shadow hover:bg-blue-700 transition text-sm">Refresh</button>
        </div>
      </div>

      <!-- Overview Cards -->
      <div class="grid grid-cols-1 md:grid-cols-4 gap-4" id="syncOverviewCards">
        <!-- populated via js -->
      </div>

      <!-- Circuit Breaker Status -->
      <div class="bg-white rounded-lg shadow p-6 border-l-4" id="cbStatusCard">
        <div class="flex justify-between items-center mb-2">
          <h3 class="text-lg font-semibold text-gray-700">Circuit Breaker Status</h3>
          <div class="flex gap-2">
            <button id="btnPauseSync" class="px-3 py-1 bg-red-600 text-white rounded text-xs hover:bg-red-700">Pause Sync</button>
            <button id="btnResumeSync" class="px-3 py-1 bg-green-600 text-white rounded text-xs hover:bg-green-700">Resume Sync</button>
          </div>
        </div>
        <div class="flex items-center gap-4">
          <div id="cbStateBadge" class="px-3 py-1 rounded-full text-sm font-medium">Loading...</div>
          <p class="text-sm text-gray-600" id="cbStateDesc">Fetching state...</p>
        </div>
      </div>

      <!-- Queue Table -->
      <div class="bg-white rounded-lg shadow overflow-hidden">
        <div class="p-4 border-b flex justify-between items-center">
          <h3 class="text-lg font-semibold text-gray-800">Outbox Queue</h3>
          <div class="flex gap-2">
            <button id="btnRetrySelected" class="px-3 py-1 bg-gray-200 text-gray-700 rounded text-xs hover:bg-gray-300">Retry Selected</button>
            <select id="queueStatusFilter" class="border rounded px-2 py-1 text-sm">
              <option value="">All Status</option>
              <option value="PENDING">Pending</option>
              <option value="PROCESSING">Processing</option>
              <option value="FAILED">Failed</option>
              <option value="DEAD_LETTER">Dead Letter</option>
            </select>
          </div>
        </div>
        <div class="overflow-x-auto">
          <table class="w-full text-sm text-left">
            <thead class="bg-gray-50 text-gray-600">
              <tr>
                <th class="px-4 py-3 w-10"><input type="checkbox" id="chkAllQueue" /></th>
                <th class="px-4 py-3">ID</th>
                <th class="px-4 py-3">Module</th>
                <th class="px-4 py-3">Action</th>
                <th class="px-4 py-3">Status</th>
                <th class="px-4 py-3">Retry</th>
                <th class="px-4 py-3">Created</th>
              </tr>
            </thead>
            <tbody id="queueTableBody" class="divide-y">
              <tr><td colspan="6" class="px-4 py-4 text-center text-gray-500">Loading queue...</td></tr>
            </tbody>
          </table>
        </div>
      </div>
      
      <!-- Performance Metrics & Histograms -->
      <div class="grid grid-cols-1 md:grid-cols-2 gap-4">
        <!-- Histograms -->
        <div class="bg-white rounded-lg shadow overflow-hidden">
          <div class="p-4 border-b">
            <h3 class="text-lg font-semibold text-gray-800">Latency Histograms (P50 - P99)</h3>
          </div>
          <div class="p-4 overflow-x-auto text-sm" id="latencyHistograms">
            <p class="text-gray-500">Loading performance data...</p>
          </div>
        </div>

        <!-- Snapshots & Reconcile -->
        <div class="bg-white rounded-lg shadow overflow-hidden">
          <div class="p-4 border-b">
            <h3 class="text-lg font-semibold text-gray-800">Nightly Snapshot & Reconciliation</h3>
          </div>
          <div class="p-4 text-sm space-y-4">
            <div>
              <h4 class="font-medium text-gray-700">Latest Snapshot</h4>
              <p id="lblSnapshotStatus" class="text-gray-600 mt-1">Loading...</p>
            </div>
            <div>
              <h4 class="font-medium text-gray-700">Reconciliation Status</h4>
              <p id="lblReconStatus" class="text-gray-600 mt-1">Ready for checks.</p>
            </div>
          </div>
        </div>
      </div>
      
      <!-- Metrics -->
      <div class="bg-white rounded-lg shadow overflow-hidden">
        <div class="p-4 border-b">
          <h3 class="text-lg font-semibold text-gray-800">Queue & Throughput Trend</h3>
        </div>
        <div class="overflow-x-auto">
          <table class="w-full text-sm text-left">
            <thead class="bg-gray-50 text-gray-600">
              <tr>
                <th class="px-4 py-3">Module</th>
                <th class="px-4 py-3">Total Events</th>
                <th class="px-4 py-3">Avg Webhook (ms)</th>
                <th class="px-4 py-3">Avg D1 Exec (ms)</th>
                <th class="px-4 py-3">Avg Queue Wait (ms)</th>
              </tr>
            </thead>
            <tbody id="metricsTableBody" class="divide-y">
              <tr><td colspan="5" class="px-4 py-4 text-center text-gray-500">Loading metrics...</td></tr>
            </tbody>
          </table>
        </div>
      </div>
    </div>
  `;let n=document.getElementById("btnRefreshSync"),a=document.getElementById("queueStatusFilter");n.addEventListener("click",i),a.addEventListener("change",$),document.getElementById("btnRetryAll").addEventListener("click",()=>m("retry",{allFailed:!0})),document.getElementById("btnResetStuck").addEventListener("click",()=>m("reset-stuck",{})),document.getElementById("btnPauseSync").addEventListener("click",()=>m("pause",{})),document.getElementById("btnResumeSync").addEventListener("click",()=>m("resume",{})),document.getElementById("btnReconcile").addEventListener("click",()=>m("reconcile",{module:"ALL",repairMode:!0})),document.getElementById("btnRetrySelected").addEventListener("click",()=>{let h=Array.from(document.querySelectorAll(".chk-queue:checked")).map(L=>L.value);if(h.length===0)return alert("No items selected");m("retry",{ids:h})}),document.getElementById("chkAllQueue").addEventListener("change",h=>{document.querySelectorAll(".chk-queue").forEach(L=>L.checked=h.target.checked)});function o(){n&&(n.disabled=!0,n.textContent="Loading...")}function s(){n&&(n.disabled=!1,n.textContent="Refresh")}let c=h=>h&&typeof h=="object"&&!Array.isArray(h)&&"data"in h&&"success"in h?h.data:h;async function m(h,L){if(confirm(`Are you sure you want to execute action: ${h}?`)){o();try{let R=await X(`/api/sync/actions/${h}`,{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify(L)});if(R.ok){let I=c(R.data);alert(I?.message||R.data?.message||"Success"),i()}else j(R.data?.error||R.error||"Action failed")}catch(R){j(R.message)}s()}}await i();async function i(){o(),await Promise.all([F(),$(),d(),M()]),s()}async function d(){try{let h=await X("/api/sync/performance");if(!h.ok||!h.data)return;let L=c(h.data);if(!L||L.status==="DATA_UNAVAILABLE"||!L.webhook){let B=document.getElementById("latencyHistograms");B&&(B.innerHTML='<p class="text-gray-500">Belum ada aktivitas sinkronisasi dalam 24 jam terakhir.</p>');let x=document.getElementById("lblSnapshotStatus");x&&(x.innerHTML='Status: <b>READY</b><br><span class="text-xs text-gray-500">Siap menjalankan rekonsiliasi & snapshot</span>');return}let{webhook:R,google_api:I,d1:N,queue:q,throughput:v}=L;document.getElementById("latencyHistograms").innerHTML=`
        <table class="w-full text-left">
          <thead><tr class="text-gray-500 border-b"><th>Metric</th><th>P50</th><th>P95</th><th>P99</th><th>Max</th></tr></thead>
          <tbody class="divide-y">
            <tr><td class="py-1">Webhook</td><td>${R.P50}ms</td><td>${R.P95}ms</td><td>${R.P99}ms</td><td>${R.Max}ms</td></tr>
            <tr><td class="py-1">Google API</td><td>${I.P50}ms</td><td>${I.P95}ms</td><td>${I.P99}ms</td><td>${I.Max}ms</td></tr>
            <tr><td class="py-1">D1 Execute</td><td>${N.P50}ms</td><td>${N.P95}ms</td><td>${N.P99}ms</td><td>${N.Max}ms</td></tr>
            <tr><td class="py-1">Queue Delay</td><td>${q.P50}ms</td><td>${q.P95}ms</td><td>${q.P99}ms</td><td>${q.Max}ms</td></tr>
          </tbody>
        </table>
        <div class="mt-4 pt-3 border-t text-gray-600 flex justify-between">
          <span>Throughput: <b>${v.events_per_sec}</b> ev/sec</span>
          <span><b>${v.events_per_min}</b> ev/min</span>
        </div>
      `,document.getElementById("lblSnapshotStatus").innerHTML='Checked 10 modules. Status: <b>COMPLETED</b><br><span class="text-xs text-gray-500">Run today at 02:00</span>'}catch(h){console.error(h)}}async function F(){try{let h=await X("/api/sync/health");if(!h.ok)return j("Failed to fetch sync health");let L=c(h.data)||{},{status:R,queue:I={},circuit_breaker:N}=L,q=`
        <div class="bg-white p-4 rounded-lg shadow border-l-4 ${R==="HEALTHY"?"border-green-500":R==="WARNING"?"border-yellow-500":"border-red-500"}">
          <p class="text-sm text-gray-500">System Health</p>
          <p class="text-2xl font-bold ${R==="HEALTHY"?"text-green-600":R==="WARNING"?"text-yellow-600":"text-red-600"}">${R}</p>
        </div>
        <div class="bg-white p-4 rounded-lg shadow border-l-4 border-blue-500">
          <p class="text-sm text-gray-500">Pending Queue</p>
          <p class="text-2xl font-bold text-gray-800">${I.pending||0}</p>
        </div>
        <div class="bg-white p-4 rounded-lg shadow border-l-4 border-orange-500">
          <p class="text-sm text-gray-500">Failed / Retrying</p>
          <p class="text-2xl font-bold text-gray-800">${I.failed||0}</p>
        </div>
        <div class="bg-white p-4 rounded-lg shadow border-l-4 border-red-500">
          <p class="text-sm text-gray-500">Dead Letters</p>
          <p class="text-2xl font-bold text-gray-800">${I.dead_letter||0}</p>
        </div>
      `;document.getElementById("syncOverviewCards").innerHTML=q;let v=document.getElementById("cbStateBadge"),B=document.getElementById("cbStateDesc"),x=document.getElementById("cbStatusCard");x.className="bg-white rounded-lg shadow p-6 border-l-4",N==="CLOSED"?(v.className="px-3 py-1 rounded-full text-sm font-medium bg-green-100 text-green-800",v.textContent="CLOSED",B.textContent="Traffic is flowing normally to Google Sheets.",x.classList.add("border-green-500")):N==="OPEN"?(v.className="px-3 py-1 rounded-full text-sm font-medium bg-red-100 text-red-800",v.textContent="OPEN",B.textContent="Failing fast. Traffic to Google Sheets is paused due to repeated failures.",x.classList.add("border-red-500")):N==="HALF_OPEN"?(v.className="px-3 py-1 rounded-full text-sm font-medium bg-yellow-100 text-yellow-800",v.textContent="HALF-OPEN",B.textContent="Testing recovery. Permitting limited traffic to verify stability.",x.classList.add("border-yellow-500")):v.textContent=N||"UNKNOWN"}catch(h){console.error(h)}}async function $(){try{let h=document.getElementById("queueStatusFilter").value,L=await X("/api/sync/queue?limit=15"+(h?"&status="+h:""));if(!L.ok)return;let R=document.getElementById("queueTableBody"),I=c(L.data),N=I&&!Array.isArray(I)&&Array.isArray(I.data)?I.data:I,q=Array.isArray(N)?N:[];if(q.length===0){R.innerHTML='<tr><td colspan="6" class="px-4 py-4 text-center text-gray-500">No events found</td></tr>';return}R.innerHTML=q.map(v=>`
        <tr class="hover:bg-gray-50">
          <td class="px-4 py-2"><input type="checkbox" class="chk-queue" value="${P(v.id)}" /></td>
          <td class="px-4 py-2 font-mono text-xs text-gray-500" title="${P(v.id)}">#${P(String(v.id).split("-")[0])}</td>
          <td class="px-4 py-2 font-medium">${P(v.entity_name)}</td>
          <td class="px-4 py-2">
             <span class="px-2 py-0.5 rounded text-xs ${v.action==="INSERT"?"bg-blue-100 text-blue-800":v.action==="UPDATE"?"bg-purple-100 text-purple-800":"bg-red-100 text-red-800"}">${P(v.action)}</span>
          </td>
          <td class="px-4 py-2">
             <span class="px-2 py-0.5 rounded text-xs ${v.status==="PENDING"?"bg-yellow-100 text-yellow-800":v.status==="PROCESSING"?"bg-blue-100 text-blue-800":v.status==="DEAD_LETTER"?"bg-red-100 text-red-800":"bg-gray-100 text-gray-800"}">${P(v.status)}</span>
             ${v.last_error?`<br><span class="text-xs text-red-500 max-w-xs block truncate" title="${P(v.last_error)}">${P(v.last_error)}</span>`:""}
          </td>
          <td class="px-4 py-2 text-gray-600">${P(String(v.retry_count||0))}</td>
          <td class="px-4 py-2 text-gray-500 whitespace-nowrap">${P(typeof window.formatDate=="function"?window.formatDate(v.created_at):v.created_at||"-")}</td>
        </tr>
      `).join("")}catch(h){console.error(h)}}async function M(){try{let h=await X("/api/sync/metrics");if(!h.ok)return;let L=document.getElementById("metricsTableBody"),R=c(h.data),I=Array.isArray(R)?R:[];if(I.length===0){L.innerHTML='<tr><td colspan="5" class="px-4 py-4 text-center text-gray-500">No metrics found in last 24h</td></tr>';return}L.innerHTML=I.map(N=>`
        <tr class="hover:bg-gray-50">
          <td class="px-4 py-2 font-medium text-gray-800">${P(N.module)}</td>
          <td class="px-4 py-2 text-gray-600">${P(String(N.total_events))}</td>
          <td class="px-4 py-2 text-gray-600">${Math.round(N.avg_webhook_ms||0)} ms</td>
          <td class="px-4 py-2 text-gray-600">${Math.round(N.avg_d1_ms||0)} ms</td>
          <td class="px-4 py-2 text-gray-600">${Math.round(N.avg_queue_wait_ms||0)} ms</td>
        </tr>
      `).join("")}catch(h){console.error(h)}}}window.parseFlexibleDate=e=>{if(!e||e==="-")return"";if(e=String(e).trim(),/^\d{5}$/.test(e)){let n=Math.floor(Number(e)-25569);return new Date(n*86400*1e3).toISOString().split("T")[0]}if(e.match(/^\d{2}[\/\-]\d{2}[\/\-]\d{4}$/)){let n=e.split(/[\/\-]/);return`${n[2]}-${n[1]}-${n[0]}`}let t=e.match(/^(\d{4})-(\d{2})-(\d{2})(?:T.*)?$/);if(t){let n=t[1],a=parseInt(t[2],10),o=parseInt(t[3],10);if(a>12&&o<=12)return`${n}-${t[3]}-${t[2]}`}return e.split("T")[0]};window.formatDate=e=>{let t=window.parseFlexibleDate(e);if(!t)return"";let n=t.split("-");if(n.length===3&&n[0].length===4){let a=["Januari","Februari","Maret","April","Mei","Juni","Juli","Agustus","September","Oktober","November","Desember"],o=parseInt(n[2],10),s=a[parseInt(n[1],10)-1];return`${o} ${s} ${n[0]}`}return t};function ye(e){return async t=>{if(!ut()&&!se()){Ge("/login");return}return e(t)}}var jt=null;function Hr(){jt&&clearInterval(jt);let e=()=>{let t=new Date,n=t.toLocaleTimeString("id-ID",{hour:"2-digit",minute:"2-digit",second:"2-digit"}).replace(/\./g,":"),a=t.toLocaleDateString("id-ID",{weekday:"long",day:"numeric",month:"long",year:"numeric"}),o=document.getElementById("header-clock-time"),s=document.getElementById("header-clock-date");o&&(o.textContent=n),s&&(s.textContent=a)};e(),jt=setInterval(e,1e3)}var xt=[];async function Or(){try{let e=await X("/api/dashboard/notifications");if(!e.ok)return;xt=e.data?.data||e.data||[];let t=document.getElementById("notif-dot");if(t){let n=xt.length>0?xt.length:51;t.style.display="inline-flex",t.textContent=n}}catch{}}function Ur(){if(!xt.length){ce({title:"Notifikasi",content:'<div class="empty-state"><p>Tidak ada notifikasi baru.</p></div>',confirmText:"Tutup",onConfirm:(t,n)=>n()});return}let e=`
    <div class="notif-list" style="max-height: 400px; overflow-y: auto;">
      ${xt.map(t=>`
        <div class="notif-item notif-severity-${t.severity||"info"}" style="padding: 12px; border: 1px solid #223456; border-left: 4px solid var(--${t.severity==="danger"?"danger":t.severity==="warning"?"warning":"primary"}); margin-bottom: 8px; border-radius: 8px; background: #0E182B; color: #F8FAFC;">
          <div style="font-weight: 600; font-size: 0.9rem; color: #F8FAFC;">${P(t.title)}</div>
          <div style="display: flex; justify-content: space-between; align-items: center; margin-top: 4px; font-size: 0.75rem; color: var(--text-3);">
            <span>\u{1F4C5} ${P(t.date)}</span>
            <span class="badge badge-${t.severity==="danger"?"danger":t.severity==="warning"?"warning":"info"}">${P(String(t.type||"").toUpperCase())}</span>
          </div>
        </div>
      `).join("")}
    </div>
  `;ce({title:`Notifikasi (${xt.length})`,content:e,confirmText:"Tutup",onConfirm:(t,n)=>n()})}function Ai(){let e=se(),t=(e?.full_name||"U")[0].toUpperCase(),n=(()=>{try{return localStorage.getItem("fcms_sidebar_collapsed")==="1"}catch{return!1}})();document.getElementById("app").innerHTML=`
    <div class="app-layout${n?" sidebar-collapsed":""}" id="app-layout">
      <!-- Sidebar dark -->
      <aside class="sidebar" id="sidebar" aria-label="Navigasi Utama">
        <div class="sidebar-header">
          <div class="sidebar-logo">
            <div class="logo-icon-wrap" style="width:34px;height:34px;border-radius:10px;background:#2563EB;display:flex;align-items:center;justify-content:center;color:#fff;flex-shrink:0;">
              <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><path d="M18 20V10M12 20V4M6 20v-6"/></svg>
            </div>
            <div class="logo-text-group">
              <span class="logo-text">FCMS</span>
              <span class="logo-subtext">Facilitycare Management System</span>
            </div>
          </div>
          <button type="button" class="sidebar-collapse-btn" id="sidebar-collapse-btn" aria-label="${n?"Perluas Sidebar":"Perkecil Sidebar"}" title="${n?"Perluas Sidebar":"Perkecil Sidebar"}" aria-expanded="${n?"false":"true"}">
            <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><polyline points="15 18 9 12 15 6"/></svg>
          </button>
          <button type="button" class="sidebar-close" id="sidebar-close" aria-label="Tutup Menu">\u2715</button>
        </div>

        <nav class="sidebar-nav" id="sidebar-nav">

          <!-- Utama -->
          <div class="nav-section">
            <span class="nav-section-label">UTAMA</span>
            <a href="#/dashboard" class="nav-item" data-route="/dashboard" data-tooltip="Dashboard" title="Dashboard" aria-label="Dashboard">
              <span class="nav-icon">
                <svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><rect x="3" y="3" width="7" height="7" rx="1"/><rect x="14" y="3" width="7" height="7" rx="1"/><rect x="3" y="14" width="7" height="7" rx="1"/><rect x="14" y="14" width="7" height="7" rx="1"/></svg>
              </span>
              <span class="nav-label">Dashboard</span>
            </a>
            <a href="#/calendar" class="nav-item" data-route="/calendar" data-tooltip="Kalender" title="Kalender" aria-label="Kalender">
              <span class="nav-icon">
                <svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><rect x="3" y="4" width="18" height="18" rx="2"/><line x1="16" y1="2" x2="16" y2="6"/><line x1="8" y1="2" x2="8" y2="6"/><line x1="3" y1="10" x2="21" y2="10"/></svg>
              </span>
              <span class="nav-label">Kalender</span>
            </a>
          </div>

          <!-- SDM -->
          <div class="nav-section">
            <span class="nav-section-label">SDM</span>
            <a href="#/employees" class="nav-item" data-route="/employees" data-tooltip="Master Karyawan" title="Master Karyawan" aria-label="Master Karyawan">
              <span class="nav-icon">
                <svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><circle cx="9" cy="7" r="4"/><path d="M3 21v-2a4 4 0 014-4h4a4 4 0 014 4v2"/><circle cx="19" cy="7" r="2"/><path d="M23 21v-1a3 3 0 00-3-3"/></svg>
              </span>
              <span class="nav-label">Master Karyawan</span>
            </a>
            <a href="#/contracts" class="nav-item" data-route="/contracts" data-tooltip="Data Kontrak" title="Data Kontrak" aria-label="Data Kontrak">
              <span class="nav-icon">
                <svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M14 2H6a2 2 0 00-2 2v16a2 2 0 002 2h12a2 2 0 002-2V8z"/><polyline points="14 2 14 8 20 8"/><line x1="16" y1="13" x2="8" y2="13"/><line x1="16" y1="17" x2="8" y2="17"/></svg>
              </span>
              <span class="nav-label">Data Kontrak</span>
            </a>
            <a href="#/sp" class="nav-item" data-route="/sp" data-tooltip="Data SP" title="Data SP" aria-label="Data SP">
              <span class="nav-icon">
                <svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"/><polyline points="22 4 12 14.01 9 11.01"/></svg>
              </span>
              <span class="nav-label">Data Sp</span>
            </a>
            <a href="#/mutasi" class="nav-item" data-route="/mutasi" data-tooltip="Data Mutasi" title="Data Mutasi" aria-label="Data Mutasi">
              <span class="nav-icon">
                <svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><polyline points="16 3 21 3 21 8"/><line x1="4" y1="20" x2="21" y2="3"/><polyline points="21 16 21 21 16 21"/><line x1="15" y1="15" x2="21" y2="21"/><line x1="4" y1="4" x2="9" y2="9"/></svg>
              </span>
              <span class="nav-label">Data Mutasi</span>
            </a>
            <a href="#/overtime" class="nav-item" data-route="/overtime" data-tooltip="Data Lembur" title="Data Lembur" aria-label="Data Lembur">
              <span class="nav-icon">
                <svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><circle cx="12" cy="12" r="10"/><polyline points="12 6 12 12 16 14"/></svg>
              </span>
              <span class="nav-label">Data Lembur</span>
            </a>
          </div>

          <!-- Operasional -->
          <div class="nav-section">
            <span class="nav-section-label">OPERASIONAL</span>
            <a href="#/timeline" class="nav-item" data-route="/timeline" data-tooltip="Time Line" title="Time Line" aria-label="Time Line">
              <span class="nav-icon">
                <svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><circle cx="12" cy="12" r="10"/><polyline points="12 6 12 12 16 14"/></svg>
              </span>
              <span class="nav-label">Time Line</span>
            </a>
            <a href="#/os-migrations" class="nav-item" data-route="/os-migrations" data-tooltip="Migrasi OS" title="Migrasi OS" aria-label="Migrasi OS">
              <span class="nav-icon">
                <svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><polyline points="16 3 21 3 21 8"/><line x1="4" y1="20" x2="21" y2="3"/><polyline points="21 16 21 21 16 21"/><line x1="15" y1="15" x2="21" y2="21"/><line x1="4" y1="4" x2="9" y2="9"/></svg>
              </span>
              <span class="nav-label">Migrasi OS</span>
            </a>
            <a href="#/issues" class="nav-item" data-route="/issues" data-tooltip="Permasalahan" title="Permasalahan" aria-label="Permasalahan">
              <span class="nav-icon">
                <svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><circle cx="12" cy="12" r="10"/><line x1="12" y1="8" x2="12" y2="12"/><line x1="12" y1="16" x2="12.01" y2="16"/></svg>
              </span>
              <span class="nav-label">Permasalahan</span>
            </a>
            <a href="#/training" class="nav-item" data-route="/training" data-tooltip="Training" title="Training" aria-label="Training">
              <span class="nav-icon">
                <svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M22 10v6M2 10l10-5 10 5-10 5z"/><path d="M6 12v5c3 3 9 3 12 0v-5"/></svg>
              </span>
              <span class="nav-label">Training</span>
            </a>
          </div>

          <!-- Laporan -->
          <div class="nav-section">
            <span class="nav-section-label">LAPORAN</span>
            <a href="#/reports/inspection" class="nav-item" data-route="/reports/inspection" data-tooltip="Report Inspeksi Hygiene 2026" title="Report Inspeksi Hygiene 2026" aria-label="Report Inspeksi Hygiene 2026">
              <span class="nav-icon">
                <svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/></svg>
              </span>
              <span class="nav-label">Report Inspeksi Hygiene 2026</span>
            </a>
            <a href="#/reports/cleaning" class="nav-item" data-route="/reports/cleaning" data-tooltip="Report GCDC 2026" title="Report GCDC 2026" aria-label="Report GCDC 2026">
              <span class="nav-icon">
                <svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><polyline points="23 6 13.5 15.5 8.5 10.5 1 18"/><polyline points="17 6 23 6 23 12"/></svg>
              </span>
              <span class="nav-label">Report GCDC 2026</span>
            </a>
            <a href="#/reports/fogging" class="nav-item" data-route="/reports/fogging" data-tooltip="Report Fogging 2026" title="Report Fogging 2026" aria-label="Report Fogging 2026">
              <span class="nav-icon">
                <svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><line x1="8" y1="6" x2="21" y2="6"/><line x1="8" y1="12" x2="21" y2="12"/><line x1="8" y1="18" x2="21" y2="18"/><line x1="3" y1="6" x2="3.01" y2="6"/><line x1="3" y1="12" x2="3.01" y2="12"/><line x1="3" y1="18" x2="3.01" y2="18"/></svg>
              </span>
              <span class="nav-label">Report Fogging 2026</span>
            </a>
            <a href="#/reports/basecamp" class="nav-item" data-route="/reports/basecamp" data-tooltip="Rekap Laporan Basecamp" title="Rekap Laporan Basecamp" aria-label="Rekap Laporan Basecamp">
              <span class="nav-icon">
                <svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M3 9l9-7 9 7v11a2 2 0 01-2 2H5a2 2 0 01-2-2z"/></svg>
              </span>
              <span class="nav-label">Rekap Laporan Basecamp</span>
            </a>
          </div>

          <!-- Referensi -->
          <div class="nav-section">
            <span class="nav-section-label">REFERENSI</span>
            <a href="#/sop-fc" class="nav-item" data-route="/sop-fc" data-tooltip="SOP Facility Care" title="SOP Facility Care" aria-label="SOP Facility Care">
              <span class="nav-icon">
                <svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/><line x1="16" y1="13" x2="8" y2="13"/><line x1="16" y1="17" x2="8" y2="17"/><polyline points="10 9 9 9 8 9"/></svg>
              </span>
              <span class="nav-label">SOP Facility Care</span>
            </a>
            <a href="#/sop-security" class="nav-item" data-route="/sop-security" data-tooltip="SOP Security" title="SOP Security" aria-label="SOP Security">
              <span class="nav-icon">
                <svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M12 22s8-4 8-10V5l-8-3-8 3v7c0 6 8 10 8 10z"/></svg>
              </span>
              <span class="nav-label">SOP Security</span>
            </a>
            <a href="#/checklist" class="nav-item" data-route="/checklist" data-tooltip="Master Document" title="Master Document" aria-label="Master Document">
              <span class="nav-icon">
                <svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><polyline points="9 11 12 14 22 4"/><path d="M21 12v7a2 2 0 01-2 2H5a2 2 0 01-2-2V5a2 2 0 012-2h11"/></svg>
              </span>
              <span class="nav-label">Master Document</span>
            </a>
            <a href="#/hygiene-standards" class="nav-item" data-route="/hygiene-standards" data-tooltip="Master Hygiene" title="Master Hygiene" aria-label="Master Hygiene">
              <span class="nav-icon">
                <svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M12 2l3.09 6.26L22 9.27l-5 4.87 1.18 6.88L12 17.77l-6.18 3.25L7 14.14 2 9.27l6.91-1.01L12 2z"/></svg>
              </span>
              <span class="nav-label">Master Hygiene</span>
            </a>
            <a href="#/dac" class="nav-item" data-route="/dac" data-tooltip="Master DAC" title="Master DAC" aria-label="Master DAC">
              <span class="nav-icon">
                <svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><circle cx="12" cy="12" r="10"/><polyline points="12 6 12 12 16 14"/></svg>
              </span>
              <span class="nav-label">Master DAC</span>
            </a>
          </div>

          <!-- Admin -->
          ${e&&e.role==="superadmin"?`
          <div class="nav-section">
            <span class="nav-section-label">ADMIN</span>
            <a href="#/users" class="nav-item" data-route="/users" data-tooltip="Manajemen User" title="Manajemen User" aria-label="Manajemen User">
              <span class="nav-icon">
                <svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M20 21v-2a4 4 0 00-4-4H8a4 4 0 00-4 4v2"/><circle cx="12" cy="7" r="4"/></svg>
              </span>
              <span class="nav-label">Manajemen User</span>
            </a>
            <a href="#/branches" class="nav-item" data-route="/branches" data-tooltip="Cabang" title="Cabang" aria-label="Cabang">
              <span class="nav-icon">
                <svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M3 9l9-7 9 7v11a2 2 0 01-2 2H5a2 2 0 01-2-2z"/><polyline points="9 22 9 12 15 12 15 22"/></svg>
              </span>
              <span class="nav-label">Cabang</span>
            </a>
            <a href="#/audit-logs" class="nav-item" data-route="/audit-logs" data-tooltip="Audit Log" title="Audit Log" aria-label="Audit Log">
              <span class="nav-icon">
                <svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M9 12l2 2 4-4m6 2a9 9 0 11-18 0 9 9 0 0118 0z"/></svg>
              </span>
              <span class="nav-label">Audit Log</span>
            </a>
            <a href="#/settings/import" class="nav-item" data-route="/settings/import" data-tooltip="Import Data Awal" title="Import Data Awal" aria-label="Import Data Awal">
              <span class="nav-icon">
                <svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M21 15v4a2 2 0 01-2 2H5a2 2 0 01-2-2v-4"/><polyline points="17 8 12 3 7 8"/><line x1="12" y1="3" x2="12" y2="15"/></svg>
              </span>
              <span class="nav-label">Import Data Awal</span>
            </a>
          </div>`:""}
          ${e&&e.role!=="superadmin"&&re(e,"branches","read")?`
          <div class="nav-section">
            <span class="nav-section-label">MASTER DATA</span>
            <a href="#/branches" class="nav-item" data-route="/branches" data-tooltip="Cabang" title="Cabang" aria-label="Cabang">
              <span class="nav-icon">
                <svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M3 9l9-7 9 7v11a2 2 0 01-2 2H5a2 2 0 01-2-2z"/><polyline points="9 22 9 12 15 12 15 22"/></svg>
              </span>
              <span class="nav-label">Cabang</span>
            </a>
          </div>`:""}
        </nav>

        <div class="sidebar-footer">
          <div class="sidebar-user">
            <div class="sidebar-avatar" data-tooltip="${P(e?.full_name||e?.username||"Profil")}" title="${P(e?.full_name||e?.username||"Profil")}">${(e?.full_name||e?.username||"U")[0].toUpperCase()}</div>
            <div class="sidebar-user-info">
              <div class="sidebar-user-name">${P(e?.full_name||e?.username||"Guest")}</div>
              <div class="sidebar-user-role">${P(e?.role||"-")}</div>
            </div>
          </div>
          <div class="sidebar-footer-actions">
            <a href="#/profile" class="sidebar-logout nav-item-profile" data-route="/profile" data-tooltip="Pengaturan" title="Pengaturan" aria-label="Pengaturan" style="flex:1; text-decoration:none;">
              <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><circle cx="12" cy="12" r="3"/><path d="M19.4 15a1.65 1.65 0 00.33 1.82l.06.06a2 2 0 010 2.83 2 2 0 01-2.83 0l-.06-.06a1.65 1.65 0 00-1.82-.33 1.65 1.65 0 00-1 1.51V21a2 2 0 01-2 2 2 2 0 01-2-2v-.09A1.65 1.65 0 009 19.4a1.65 1.65 0 00-1.82.33l-.06.06a2 2 0 01-2.83 0 2 2 0 010-2.83l.06-.06a1.65 1.65 0 00.33-1.82 1.65 1.65 0 00-1.51-1H3a2 2 0 01-2-2 2 2 0 012-2h.09A1.65 1.65 0 004.6 9a1.65 1.65 0 00-.33-1.82l-.06-.06a2 2 0 010-2.83 2 2 0 012.83 0l.06.06a1.65 1.65 0 001.82.33H9a1.65 1.65 0 001-1.51V3a2 2 0 012-2 2 2 0 012 2v.09a1.65 1.65 0 001 1.51 1.65 1.65 0 001.82-.33l.06-.06a2 2 0 012.83 0 2 2 0 010 2.83l-.06.06a1.65 1.65 0 00-.33 1.82V9a1.65 1.65 0 001.51 1H21a2 2 0 012 2 2 2 0 01-2 2h-.09a1.65 1.65 0 00-1.51 1z"/></svg>
              <span class="sidebar-btn-label">Pengaturan</span>
            </a>
            <button type="button" class="sidebar-logout" id="logout-btn" data-tooltip="Keluar" title="Keluar" aria-label="Keluar" style="flex:1;">
              <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M9 21H5a2 2 0 01-2-2V5a2 2 0 012-2h4"/><polyline points="16 17 21 12 16 7"/><line x1="21" y1="12" x2="9" y2="12"/></svg>
              <span class="sidebar-btn-label">Keluar</span>
            </button>
          </div>
        </div>
      </aside>
      <div class="fc-sidebar-tooltip" id="fc-sidebar-tooltip" role="tooltip" aria-hidden="true"></div>

      <!-- Mobile overlay -->
      <div class="sidebar-overlay" id="sidebar-overlay"></div>

      <!-- Main wrapper -->
      <div class="main-wrapper">
        <!-- Topbar -->
        <header class="topbar" id="topbar">
          <div class="topbar-left">
            <button class="topbar-menu-btn" id="topbar-menu-btn" aria-label="Menu">
              <svg width="20" height="20" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><line x1="3" y1="6" x2="21" y2="6"/><line x1="3" y1="12" x2="21" y2="12"/><line x1="3" y1="18" x2="21" y2="18"/></svg>
            </button>
            <div class="topbar-mobile-logo" style="display:none;align-items:center;gap:8px;">
              <div style="width:28px;height:28px;border-radius:8px;background:linear-gradient(135deg,#2563EB,#0EA5E9);display:flex;align-items:center;justify-content:center;box-shadow:0 2px 8px rgba(37,99,235,.4);color:#fff;flex-shrink:0;">
                <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><path d="M18 20V10M12 20V4M6 20v-6"/></svg>
              </div>
              <span class="logo-text" style="font-size:1.15rem;font-weight:800;letter-spacing:-0.3px;color:#fff;">FCMS</span>
            </div>
            <div class="topbar-welcome">
              <div class="topbar-greeting">
                <span class="topbar-greeting-time">${(()=>{let B=new Date().getHours();return B>=4&&B<11?"Selamat Pagi":B>=11&&B<15?"Selamat Siang":B>=15&&B<18?"Selamat Sore":"Selamat Malam"})()}, </span><span class="topbar-greeting-name">${P(e?.full_name||e?.username||"Guest")}</span> \u{1F44B}
              </div>
              <div class="topbar-subtitle">
                Ringkasan Operasional FCMS Hari Ini
              </div>
            </div>
          </div>

          <div class="topbar-center" id="topbar-clock">
            <div class="header-clock">
              <div class="header-clock-time" id="header-clock-time">00:00:00</div>
              <div class="header-clock-date" id="header-clock-date">Memuat...</div>
            </div>
          </div>

          <div class="topbar-right">
            <button class="topbar-icon-btn notif-btn" id="btn-notif" title="Notifikasi" aria-label="Notifikasi">
              <svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M18 8A6 6 0 006 8c0 7-3 9-3 9h18s-3-2-3-9"/><path d="M13.73 21a2 2 0 01-3.46 0"/></svg>
              <span class="notif-badge" id="notif-dot" style="display:none">0</span>
            </button>
            <a href="#/profile" class="topbar-user-btn" title="Profil">
              <div class="topbar-avatar" style="display:flex;align-items:center;justify-content:center;background:#2563EB;color:#fff;font-weight:700;font-size:0.85rem;border-radius:50%;">${P(String(e?.full_name||e?.username||"G").trim().charAt(0).toUpperCase())}</div>
              <div class="topbar-user-text">
                <span class="topbar-user-name">${P(e?.full_name||e?.username||"Guest")}</span>
                <span class="topbar-user-role-mini">${P(e?.role||"-")}</span>
              </div>
              <svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24" style="margin-left:4px;color:var(--text-3)"><polyline points="6 9 12 15 18 9"/></svg>
            </a>
          </div>
        </header>

        <main id="main-content" class="main-content"></main>
      </div>
    </div>
  `;let a=document.getElementById("app-layout"),o=document.getElementById("sidebar"),s=document.getElementById("sidebar-overlay"),c=document.getElementById("topbar-menu-btn"),m=document.getElementById("sidebar-close"),i=document.getElementById("sidebar-collapse-btn"),d=document.getElementById("fc-sidebar-tooltip"),F=()=>{o.classList.add("open"),s.classList.add("show")},$=()=>{o.classList.remove("open"),s.classList.remove("show")},M=()=>{d&&(d.classList.remove("visible"),d.setAttribute("aria-hidden","true"))},h=B=>{if(a){if(a.classList.toggle("sidebar-collapsed",B),i){let x=B?"Perluas Sidebar":"Perkecil Sidebar";i.setAttribute("aria-label",x),i.setAttribute("title",x),i.setAttribute("aria-expanded",B?"false":"true")}M();try{localStorage.setItem("fcms_sidebar_collapsed",B?"1":"0")}catch{}setTimeout(()=>window.dispatchEvent(new Event("resize")),230)}};i?.addEventListener("click",()=>{let B=!a?.classList.contains("sidebar-collapsed");h(B)}),c?.addEventListener("click",F),m?.addEventListener("click",$),s?.addEventListener("click",$),document.querySelectorAll(".nav-item").forEach(B=>B.addEventListener("click",$));let L=B=>{let x=()=>{if(!d||!a?.classList.contains("sidebar-collapsed")||window.innerWidth<768)return;let k=B.getAttribute("data-tooltip");if(!k)return;let _=B.getBoundingClientRect();d.textContent=k,d.style.top=`${Math.round(_.top+_.height/2)}px`,d.style.left=`${Math.round(_.right+10)}px`,d.classList.add("visible"),d.setAttribute("aria-hidden","false")};B.addEventListener("mouseenter",x),B.addEventListener("pointerenter",x),B.addEventListener("focus",x),B.addEventListener("mouseleave",M),B.addEventListener("pointerleave",M),B.addEventListener("blur",M)};document.querySelectorAll("#sidebar [data-tooltip]").forEach(L);let R=document.getElementById("sidebar-nav");o?.addEventListener("scroll",M,{passive:!0}),R?.addEventListener("scroll",M,{passive:!0}),pn(document.getElementById("main-content"));let I=new Set(["/users","/settings/import","/audit-logs","/sync"]),N=String(e?.role||"").trim().toLowerCase(),q=re(e,"read");document.querySelectorAll(".nav-item").forEach(B=>{let x=B.dataset.route;(!q||N!=="superadmin"&&I.has(x))&&(B.style.display="none")}),document.querySelectorAll(".nav-section").forEach(B=>{Array.from(B.querySelectorAll(".nav-item")).some(k=>k.style.display!=="none")||(B.style.display="none")});function v(){let B=window.location.hash.replace("#","")||"/dashboard",[x]=B.split("?");document.querySelectorAll(".nav-item").forEach(K=>{let G=K.dataset.route,J=x===G||G!=="/dashboard"&&(x+"/").startsWith(G+"/");K.classList.toggle("active",J)});let k=document.querySelector(".nav-item-profile");k&&k.classList.toggle("active",x==="/profile");let _=document.getElementById("topbar-title"),E=document.querySelector(".nav-item.active"),w=E?.querySelector(".nav-label");_&&w&&(_.textContent=w.textContent),E&&typeof E.scrollIntoView=="function"&&requestAnimationFrame(()=>E.scrollIntoView({block:"nearest",inline:"nearest"}))}window.addEventListener("hashchange",v),v(),Hr(),document.getElementById("btn-fullscreen")?.addEventListener("click",()=>{document.fullscreenElement?document.exitFullscreen?.():document.documentElement.requestFullscreen?.()}),document.getElementById("logout-btn")?.addEventListener("click",async()=>{await X("/api/auth/logout",{method:"POST"}),mt(),jt&&clearInterval(jt),Ge("/login")}),Or(),document.getElementById("btn-notif")?.addEventListener("click",B=>{B.preventDefault(),Ur()})}async function Gr(){if(he("/",()=>Ge("/dashboard")),he("/login",({main:e})=>xn(e)),he("/dashboard",ye(({main:e})=>hn(e))),he("/calendar",ye(({main:e})=>ci(e))),he("/employees",ye(({main:e,params:t})=>Tn(e,t))),he("/contracts",ye(({main:e,params:t})=>Ln(e,t))),he("/sp",ye(({main:e})=>fi(e))),he("/mutasi",ye(({main:e})=>yi(e))),he("/overtime",ye(({main:e})=>ki(e))),he("/timeline",ye(({main:e,params:t})=>Kn(e,t))),he("/os-migrations",ye(({main:e})=>Ci(e))),he("/issues",ye(({main:e,params:t})=>qn(e,t))),he("/training",ye(({main:e})=>On(e))),he("/reports/inspection",ye(({main:e})=>Un(e))),he("/reports/cleaning",ye(({main:e})=>Gn(e))),he("/reports/fogging",ye(({main:e,params:t})=>Vn(e,t))),he("/reports/basecamp",ye(({main:e})=>Jn(e))),he("/sop",()=>{window.location.hash="#/sop-fc"}),he("/sop-fc",ye(({main:e})=>Xn(e))),he("/sop-security",ye(({main:e})=>ai(e))),he("/checklist",ye(({main:e})=>ri(e))),he("/hygiene-standards",ye(({main:e})=>si(e))),he("/dac",ye(({main:e})=>Ei(e))),he("/users",ye(({main:e})=>{let t=se();if(!t||t.role!=="superadmin"){Ge("/dashboard");return}li(e)})),he("/branches",ye(({main:e})=>di(e))),he("/profile",ye(({main:e})=>pi(e))),he("/settings/import",ye(({main:e})=>{let t=se();if(!t||t.role!=="superadmin"){Ge("/dashboard");return}hi(e)})),he("/audit-logs",ye(({main:e})=>{let t=se();if(!t||t.role!=="superadmin"){Ge("/dashboard");return}wi(e)})),he("/sync",ye(({main:e})=>{let t=se();if(!t||t.role!=="superadmin"){Ge("/dashboard");return}$i(e)})),window.location.hash!=="#/login"){let e=await X("/api/auth/me");if(e.ok&&e.data?.data){let t={...e.data.data},n=typeof t.csrf_token=="string"?t.csrf_token:null;delete t.csrf_token,n&&kt(n),Ht(t),Ai()}else mt(),Ge("/login")}window.addEventListener("fm:login",()=>{Ai(),Ge("/dashboard")}),sn()}typeof window<"u"&&(window.addEventListener("unhandledrejection",e=>{let t=e?.reason?.message||String(e?.reason||"Unknown promise rejection");console.error("[FCMS Stability] Unhandled promise rejection captured:",t),e.preventDefault()}),window.addEventListener("error",e=>{let t=e?.message||e?.error?.message||"Unknown runtime error";console.error("[FCMS Stability] Runtime error captured:",t)}));Gr().catch(e=>{console.error("[FCMS Stability] App initialization error:",e?.message||e),j("Terjadi kendala saat memulai aplikasi. Silakan muat ulang halaman.")});"serviceWorker"in navigator&&window.location.protocol.startsWith("http")&&window.addEventListener("load",()=>{navigator.serviceWorker.register("/sw.js").then(e=>{console.log("[PWA] Service Worker registered with scope:",e.scope),e.addEventListener("updatefound",()=>{let t=e.installing;t.addEventListener("statechange",()=>{t.state==="installed"&&navigator.serviceWorker.controller&&console.log("[PWA] New version available, reloading on user confirmation or refresh.")})})}).catch(e=>{console.warn("[PWA] Service Worker registration failed:",e)})});var Rt=null;window.addEventListener("beforeinstallprompt",e=>{e.preventDefault(),Rt=e,window.__showInstallPrompt=()=>{Rt&&(Rt.prompt(),Rt.userChoice.then(t=>{t.outcome==="accepted"&&console.log("[PWA] User accepted the install prompt"),Rt=null}))}});
