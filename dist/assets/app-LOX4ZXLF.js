var Ii=Object.defineProperty;var na=(e,t)=>()=>(e&&(t=e(e=0)),t);var Va=(e,t)=>{for(var n in t)Ii(e,n,{get:t[n],enumerable:!0})};function Y(e){return e==null?"":String(e).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;").replace(/'/g,"&#39;").replace(/`/g,"&#96;")}function Hi(e,t="#"){if(e==null)return t;let n=String(e).trim();if(!n)return t;let a=n.replace(Oi,"");return qi.test(a)?t:/^(?:#|\/(?!\/)|\.{1,2}\/|https?:\/\/|mailto:|tel:)/i.test(n)?n:t}function qt(e){if(e==null)return"";let t=String(e);return t=t.replace(/<\s*(script|style|iframe|object|embed|applet|svg|math|meta|link|base)\b[^>]*>[\s\S]*?<\s*\/\s*\1\s*>/gi,""),t=t.replace(/<\s*\/?\s*(script|style|iframe|object|embed|applet|svg|math|meta|link|base)\b[^>]*\/?>/gi,""),t=t.replace(/\s+on[a-z0-9_-]+\s*=\s*(?:"[^"]*"|'[^']*'|[^\s>]+)/gi,""),t=t.replace(/\s+(?:srcdoc|formaction)\s*=\s*(?:"[^"]*"|'[^']*'|[^\s>]+)/gi,""),t=t.replace(/(\b(?:href|src|action|xlink:href|poster)\s*=\s*)(?:"([^"]*)"|'([^']*)'|([^\s>]+))/gi,(n,a,o,s,l)=>{let p=Hi(o??s??l??"","#");return`${a}"${Y(p)}"`}),t}var qi,Oi,Ne=na(()=>{qi=/^\s*(?:javascript|vbscript|data|file)\s*:/i,Oi=/[\u0000-\u001f\u007f\s]+/g});var nn={};Va(nn,{confirmDialog:()=>me,createModal:()=>pe});function pe({title:e,content:t,onConfirm:n,onCancel:a,confirmText:o="Simpan",cancelText:s="Batal",size:l="md",confirmClass:h="btn-primary"}){let p={sm:"400px",md:"560px",lg:"720px",xl:"900px"},c=document.createElement("div");c.className="modal-overlay";let E=typeof t=="string"?qt(t):"";c.innerHTML=`
    <div class="modal" style="max-width:${p[l]||p.md}">
      <div class="modal-header">
        <h3 class="modal-title">${Y(e)}</h3>
        <button class="modal-close" aria-label="Close">&times;</button>
      </div>
      <div class="modal-body">${E}</div>
      <div class="modal-footer">
        <button class="btn btn-ghost modal-cancel">${Y(s)}</button>
        ${n?`<button class="btn ${Y(h)} modal-confirm">${Y(o)}</button>`:""}
      </div>
    </div>
  `,t instanceof HTMLElement&&c.querySelector(".modal-body").appendChild(t);let C=()=>{c.classList.remove("show"),setTimeout(()=>c.remove(),250)};return c.querySelector(".modal-close").addEventListener("click",()=>{a&&a(),C()}),c.querySelector(".modal-cancel").addEventListener("click",()=>{a&&a(),C()}),n&&c.querySelector(".modal-confirm").addEventListener("click",()=>n(c,C)),c.addEventListener("click",T=>{T.target===c&&(a&&a(),C())}),document.body.appendChild(c),requestAnimationFrame(()=>c.classList.add("show")),{overlay:c,close:C}}function me(e,t,n="Konfirmasi"){if(typeof e=="object"&&e!==null){let a=e;return pe({title:a.title||"Konfirmasi",content:typeof a.message=="string"?`<p>${Y(a.message)}</p>`:a.message,onConfirm:async(o,s)=>{a.onConfirm&&await a.onConfirm(),s()},onCancel:a.onCancel,confirmText:a.confirmText||"Ya, Lanjutkan",cancelText:a.cancelText||"Batal",confirmClass:a.confirmClass||"btn-danger"})}return pe({title:n,content:typeof e=="string"?`<p>${Y(e)}</p>`:e,onConfirm:async(a,o)=>{t&&await t(),o()},confirmText:"Ya, Lanjutkan",confirmClass:"btn-danger"})}var Fe=na(()=>{Ne()});var Oe={};Va(Oe,{bindExcelDropdown:()=>or,downloadExcel:()=>ge,parseExcel:()=>xa,renderExcelButtons:()=>sr});function xa(e){return new Promise((t,n)=>{let a=new FileReader;a.onload=o=>{try{let s=new Uint8Array(o.target.result),l=XLSX.read(s,{type:"array"}),h=l.SheetNames[0],p=l.Sheets[h];console.log("--- START EXCEL PARSING ---"),console.log(`File Name: ${e.name}`),console.log(`File Size: ${(e.size/1024).toFixed(2)} KB`),console.log(`File Type: ${e.type||"unknown"}`),console.log(`Sheets Found: ${l.SheetNames.join(", ")}`),console.log(`Sheet Used: ${h}`);let c=XLSX.utils.decode_range(p["!ref"]||"A1:A1"),E=c.e.r-c.s.r+1,C=c.e.c-c.s.c+1;console.log(`Total Rows (including empty): ${E}`),console.log(`Total Columns: ${C}`);let T=[];for(let F=c.s.c;F<=c.e.c;++F){let R=p[XLSX.utils.encode_cell({c:F,r:c.s.r})];R&&R.v&&T.push(R.v)}console.log(`Headers Found: ${T.join(", ")}`),console.log("---------------------------");let u=XLSX.utils.sheet_to_json(p,{defval:""});Object.defineProperty(u,"__worksheet",{value:p,enumerable:!1}),Object.defineProperty(u,"__headers",{value:T,enumerable:!1}),t(u)}catch(s){n(s)}},a.onerror=o=>n(o),a.readAsArrayBuffer(e)})}function ge(e,t){try{let n=XLSX.utils.json_to_sheet(e),a=XLSX.utils.book_new();XLSX.utils.book_append_sheet(a,n,"Data"),XLSX.writeFile(a,`${t}.xlsx`)}catch(n){throw console.error("Error generating Excel file:",n),n}}function ya(){vn||typeof document>"u"||typeof document.addEventListener!="function"||(vn=!0,document.addEventListener("click",e=>{let t=e.target?.closest?.("[data-excel-toggle-menu]");if(t){let n=t.getAttribute("data-excel-toggle-menu"),a=n?document.getElementById(n):null;document.querySelectorAll(".aksi-menu-content.show-aksi-menu").forEach(o=>{o!==a&&o.classList.remove("show-aksi-menu")}),a&&a.classList.toggle("show-aksi-menu");return}e.target?.closest?.(".excel-actions-dropdown")||document.querySelectorAll(".aksi-menu-content.show-aksi-menu").forEach(n=>{n.classList.remove("show-aksi-menu")})}))}function or(e,t=typeof document<"u"?document:null){if(ya(),!t||typeof t.querySelector!="function")return;let n=t.querySelector(`#btn-aksi-${e}`),a=t.querySelector(`#excel-menu-${e}`);n&&a&&!n.dataset.excelBound&&(n.dataset.excelBound="1",n.getAttribute("data-excel-toggle-menu")||n.addEventListener("click",o=>{o.stopPropagation(),a.classList.toggle("show-aksi-menu")}))}function sr(e){return ya(),`
    <div class="excel-actions-dropdown" style="position:relative; display:inline-block;">
      <button type="button" class="btn" id="btn-aksi-${e}" data-excel-toggle-menu="excel-menu-${e}" style="background:#fff; border:1px solid #E2E8F0; padding:8px 16px; border-radius:8px; font-weight:600; color:#334155; display:flex; align-items:center; gap:6px; cursor:pointer;">
        \u22EE Aksi
      </button>
      <div id="excel-menu-${e}" class="aksi-menu-content" style="display:none; position:absolute; top:calc(100% + 4px); right:0; background:#fff; border:1px solid #E2E8F0; border-radius:8px; box-shadow:0 4px 12px rgba(0,0,0,0.1); flex-direction:column; min-width:180px; z-index:999; padding:8px 0;">
        <button type="button" class="dropdown-item" id="btn-export-${e}" style="width:100%; text-align:left; padding:10px 16px; background:none; border:none; cursor:pointer; font-size:0.9rem; color:#334155; display:flex; gap:8px; align-items:center;">
          \u{1F4E5} Export Excel
        </button>
        <button type="button" class="dropdown-item" id="btn-template-${e}" style="width:100%; text-align:left; padding:10px 16px; background:none; border:none; cursor:pointer; font-size:0.9rem; color:#334155; display:flex; gap:8px; align-items:center;">
          \u{1F4C4} Download Template
        </button>
        <label class="dropdown-item" style="display:flex; width:100%; text-align:left; padding:10px 16px; background:none; border:none; cursor:pointer; font-size:0.9rem; color:#334155; margin:0; gap:8px; align-items:center;" id="label-import-${e}">
          \u{1F4E4} Import Excel
          <input type="file" id="input-import-${e}" accept=".xlsx, .xls, .csv" style="display:none;">
        </label>
      </div>
    </div>
    <style>.show-aksi-menu{display:flex !important}.dropdown-item:hover{background-color:#F8FAFC !important;color:#2563EB !important}</style>
  `}var vn,we=na(()=>{vn=!1;ya()});var ia="fm-operations-api-staging.facilitycare-audydental.workers.dev",oa="https://fm-operations-api-staging.facilitycare-audydental.workers.dev",ra="fm-operations-api.facilitycare-audydental.workers.dev",Ya="https://fm-operations-api.facilitycare-audydental.workers.dev",Hr=Object.freeze([oa]);function Wa(e){if(typeof e!="string")return!1;let t=e.trim();return t.startsWith("__")&&t.endsWith("__")&&t.length>=4}function da(e=null){let t=e??(typeof globalThis<"u"&&globalThis.window&&globalThis.window.__FM_CONFIG?globalThis.window.__FM_CONFIG:null);if(!t||typeof t!="object")throw new Error("CONFIG_ERROR: Missing runtime configuration (window.__FM_CONFIG). Fail-closed.");let n=t.ENVIRONMENT!==void 0?t.ENVIRONMENT:t.ENV,a=t.API_BASE_URL!==void 0?t.API_BASE_URL:t.API_BASE;if(Wa(n)||Wa(a))throw new Error("CONFIG_ERROR: Unreplaced build configuration placeholders detected. Fail-closed.");if(n==null||String(n).trim()==="")throw new Error("CONFIG_ERROR: Environment is not configured. Explicit environment required.");let o=String(n).trim().toLowerCase();if(o==="staging"){if(a==null||String(a).trim()==="")throw new Error("CONFIG_ERROR: Missing Staging API_BASE_URL configuration. Fail-closed.");let s=String(a).trim();if(s.startsWith("//"))throw new Error("CONFIG_ERROR: Protocol-relative URLs are not allowed in staging API_BASE_URL.");let l;try{l=new URL(s)}catch{throw new Error(`CONFIG_ERROR: Invalid or unknown Staging API_BASE_URL '${s}'. Fail-closed.`)}if(l.protocol!=="https:"&&l.protocol!=="http:")throw new Error("CONFIG_ERROR: Staging API_BASE_URL must use https protocol.");let h=l.hostname.toLowerCase().trim();if(h===ra)throw new Error("SECURITY_VIOLATION: Staging environment cannot target Production Worker API.");if(h.includes("fcvm"))throw new Error("SECURITY_VIOLATION: Staging environment cannot target FCVM resources.");if(l.username||l.password)throw new Error("SECURITY_VIOLATION: Embedded credentials are not permitted in Staging API_BASE_URL.");if(l.protocol!=="https:"||h!==ia||l.port!==""||l.pathname&&l.pathname!=="/"||l.search!==""||l.hash!=="")throw new Error(`SECURITY_VIOLATION: Unknown or unapproved Staging API target '${s}'. Only '${oa}' is permitted.`);return oa}if(o==="production"){if(a!=null&&String(a).trim()!==""){let s=String(a).trim();if(s.startsWith("//"))throw new Error("CONFIG_ERROR: Protocol-relative URLs are not allowed in production API_BASE_URL.");let l;try{l=new URL(s)}catch{throw new Error("CONFIG_ERROR: Production API_BASE_URL is a malformed URL.")}let h=l.hostname.toLowerCase().trim();if(h.includes("fcvm"))throw new Error("SECURITY_VIOLATION: Production environment cannot target FCVM resources.");if(h===ia)throw new Error("SECURITY_VIOLATION: Production environment cannot target Staging Worker API.");if(l.protocol!=="https:"||h!==ra)throw new Error(`SECURITY_VIOLATION: Unknown Production API target '${s}'.`);return Ya}return Ya}if(o==="development"){if(a==null||String(a).trim()===""||String(a).trim()==="/api")return"/api";let s=String(a).trim();if(s==="/api")return"/api";let l;try{l=new URL(s)}catch{throw new Error("CONFIG_ERROR: Development API_BASE_URL is a malformed URL.")}let h=l.hostname.toLowerCase().trim();if(h===ra)throw new Error("SECURITY_VIOLATION: Development environment cannot target Production Worker API.");if(h.includes("fcvm"))throw new Error("SECURITY_VIOLATION: Development environment cannot target FCVM resources.");if(h!=="localhost"&&h!=="127.0.0.1"&&h!==ia)throw new Error(`SECURITY_VIOLATION: Unknown Development API target '${s}'.`);return s.replace(/\/+$/,"")}throw new Error(`CONFIG_ERROR: Unknown environment '${o}'. Fail-closed.`)}function Za(e,t){let n=String(e||"").trim().replace(/\/+$/,""),a=String(t||"").trim(),o=a.startsWith("/")?a:`/${a}`;if(!n)throw new Error("CONFIG_ERROR: API base URL cannot be empty when constructing request URL.");return n==="/api"&&(o==="/api"||o.startsWith("/api/"))?o:`${n}${o}`}var sa="",en="";typeof globalThis<"u"&&globalThis.window&&globalThis.window.__FM_CONFIG&&(sa=da(globalThis.window.__FM_CONFIG),en=sa);var pa=null,ca=null,yt=null;function Mi(e){if(typeof document>"u"||!document.cookie)return null;let t=`${e}=`;for(let n of document.cookie.split(";")){let a=n.trim();if(a.startsWith(t))try{return decodeURIComponent(a.slice(t.length))}catch{return a.slice(t.length)}}return null}function dt(){return typeof localStorage<"u"&&localStorage.removeItem("fm_token"),pa}function tn(e){typeof localStorage<"u"&&localStorage.removeItem("fm_token"),pa=e||null}function la(){return ca||Mi("fcms_csrf")}function vt(e){ca=e||null}function pt(){pa=null,ca=null,typeof localStorage<"u"&&(localStorage.removeItem("fm_token"),localStorage.removeItem("fm_user"))}function de(){if(typeof localStorage>"u")return null;try{return JSON.parse(localStorage.getItem("fm_user")||"null")}catch{return null}}function Rt(e){typeof localStorage>"u"||localStorage.setItem("fm_user",JSON.stringify(e))}var zi=[{id:"employees",label:"Data Karyawan (HR)"},{id:"contracts",label:"Data Kontrak (HR)"},{id:"sp",label:"Surat Peringatan / SP (HR)"},{id:"mutasi",label:"Data Mutasi (HR)"},{id:"overtime",label:"Data Lembur (HR)"},{id:"schedule",label:"Time Line / Jadwal (Operasional)"},{id:"os_migrations",label:"Migrasi OS (Operasional)"},{id:"issues",label:"Permasalahan / Issues (Operasional)"},{id:"training",label:"Training Karyawan (Operasional)"},{id:"reports",label:"Laporan Kebersihan / Basecamp"},{id:"sop",label:"SOP Facility Care & Security"},{id:"checklist",label:"Master Dokumen & Formulir"},{id:"hygiene_standards",label:"Master Standar Hygiene"},{id:"dac",label:"Master Standar DAC"},{id:"branches",label:"Data Cabang"}],Qa=new Set(["relievers","reliefer","one_on_one","one-on-one","forms","master_forms","supply","supply_requests","reports/supply","reports_supply"]),Xa=["users","audit_logs","audit-logs","sync","sync_dashboard","monitoring","settings","system_settings","import"],Ki=new Set(zi.map(e=>e.id));function Ni(e){if(!e)return[];let t=e;if(typeof e=="object"&&!Array.isArray(e)&&(t=e.assigned_modules!==void 0?e.assigned_modules:e.assignedModules||e.modules),!t)return[];let n=o=>String(o).trim().toLowerCase().replace(/-/g,"_"),a=o=>Array.from(new Set(o.map(n).filter(s=>s&&Ki.has(s))));if(Array.isArray(t))return a(t);if(typeof t=="string"){try{let o=JSON.parse(t);if(Array.isArray(o))return a(o)}catch{}return a(t.split(","))}return[]}var Ur=Object.freeze(["superadmin","fc_specialist_leader","viewer"]);function ie(e,t="read",n=null){if(!e||!e.role||e.isRestrictedToken===!0||Number(e.must_change_password)===1)return!1;let a=String(e.role).trim().toLowerCase(),o=null,s="read";if(n!==null)o=t,s=n;else{let c=String(t||"read").toLowerCase();["read","write","create","update","delete","export","import","admin"].includes(c)?(s=c,o=null):(o=c,s="read")}let l=String(s||"read").trim().toLowerCase(),h=o?String(o).trim().toLowerCase():"",p=h.replace(/-/g,"_");if(p&&(Qa.has(p)||Qa.has(h)))return!1;if(a==="superadmin")return!0;if(l==="admin"||p&&(Xa.includes(p)||Xa.includes(h)))return!1;if(a==="fc_specialist_leader"||a==="editor_khusus"||a==="fc_specialist"){if(l==="read")return!0;if(l==="create"||l==="update"||l==="write"){if(a==="editor_khusus"&&e.assigned_modules!==void 0&&e.assigned_modules!==null){let E=Ni(e);if(E.length>0)return E.includes(p);if(typeof e.assigned_modules=="string"&&e.assigned_modules.trim()==="")return!1}return p==="overtime"||p==="issues"}return!1}return a==="viewer"||a==="manager"||a==="spv"||a==="supervisor"?l==="read":a==="admin"?!(l==="import"||l==="admin"):a==="input_lembur"?l==="read"?!0:l==="create"||l==="update"||l==="write"?p==="overtime":!1:!1}async function Ri(e={}){if(!(dt()||la()))return yt||(yt=(async()=>{try{let t=da(e.configOverride||null),n=typeof e.fetchImpl=="function"?e.fetchImpl:fetch,a=Za(t,`/api/auth/me?cb=${Date.now()}`),o=await n(a,{method:"GET",credentials:"include",headers:{"Content-Type":"application/json"}});if(o.ok){let l=(await o.json().catch(()=>null))?.data?.csrf_token;typeof l=="string"&&l.length>0&&vt(l)}}catch{}finally{yt=null}})(),yt)}var Nt=new Map;async function X(e,t={}){let n=da(t.configOverride||null);sa=n,en=n;let a=String(t.method||"GET").toUpperCase(),s=a==="GET"&&!t.fetchImpl&&!t.signal?`${n}::${e}::${dt()||""}`:null;if(s&&Nt.has(s))return Nt.get(s);let l=async()=>{!t.fetchImpl&&["POST","PUT","PATCH","DELETE"].includes(a)&&!e.startsWith("/api/auth/login")&&de()&&!dt()&&!la()&&await Ri(t);let p=dt(),c=la(),E={"Content-Type":"application/json",...p?{Authorization:`Bearer ${p}`}:{},...c&&["POST","PUT","PATCH","DELETE"].includes(a)?{"X-CSRF-Token":c}:{},...t.headers||{}},C=e.includes("/import")||e.includes("/backup")||e.includes("/export")?6e4:2e4,T=typeof t.timeoutMs=="number"&&t.timeoutMs>0?t.timeoutMs:C,u=null,F=null,R=!1;!t.signal&&typeof AbortController=="function"&&(u=new AbortController,F=setTimeout(()=>{R=!0,u.abort()},T));try{let K=`cb=${Date.now()}`,I=e.includes("?")?"&":"?",N=`${e}${I}${K}`,w=Za(n,N),M=typeof t.fetchImpl=="function"?t.fetchImpl:fetch,k={credentials:"include",...t,headers:E,...u?{signal:u.signal}:{}};delete k.timeoutMs,delete k.fetchImpl,delete k.configOverride;let x=await M(w,k),_;try{let B=await x.text();try{_=JSON.parse(B)}catch{_={error:`Server Error (${x.status}): ${B.substring(0,80)}...`}}}catch{_={error:"Gagal membaca respon dari server"}}return x.status===401&&(pt(),typeof window<"u"&&window.location&&(window.location.hash="#/login")),{ok:x.ok,status:x.status,data:_}}catch(K){return R||K?.name==="AbortError"?{ok:!1,status:408,data:{error:"Permintaan waktu habis (timeout). Silakan coba lagi."}}:{ok:!1,status:0,data:{error:`Koneksi terputus. Periksa jaringan Anda. (${K.message})`}}}finally{F!==null&&clearTimeout(F)}};if(!s)return l();let h=l().finally(()=>{Nt.delete(s)});return Nt.set(s,h),h}var ma={},jt=null,ua=0;function be(e,t){ma[e]=t}function qe(e){window.location.hash=e}function an(){async function e(){let t=++ua,n=window.location.hash.replace("#","")||"/dashboard",[a,...o]=n.split("?"),s=ma[a];if(!s){for(let[h,p]of Object.entries(ma))if(h.endsWith("/*")&&a.startsWith(h.slice(0,-2))){s=p;break}}if(jt)try{jt()}catch(h){console.warn("[Router] Previous page cleanup error:",h?.message||h)}finally{jt=null}let l=document.getElementById("main-content");if(l&&(l.innerHTML='<div class="loading-spinner"><div class="spinner"></div></div>'),s)try{let h=new URLSearchParams(o.join("?")),p=a.split("/").filter(Boolean),c=await s({path:a,params:h,segments:p,main:l});if(t!==ua){if(typeof c=="function")try{c()}catch{}return}typeof c=="function"&&(jt=c)}catch(h){if(t!==ua)return;console.error(`[Router] Module render error on '${a}':`,h?.message||h);let p=l||document.getElementById("app");p&&(p.innerHTML=`
            <div class="empty-state" style="padding:32px;text-align:center;">
              <h2 style="margin-bottom:8px;">Terjadi kendala saat memuat halaman</h2>
              <p class="text-muted" style="margin-bottom:16px;">Modul ini mengalami kesalahan sementara dan telah diisolasi agar tidak mengganggu modul lainnya.</p>
              <div style="display:flex;gap:8px;justify-content:center;">
                <button type="button" class="btn btn-primary" id="btn-router-retry">Muat Ulang Halaman</button>
                <a href="#/dashboard" class="btn btn-secondary" style="text-decoration:none;">Kembali ke Dashboard</a>
              </div>
            </div>
          `,p.querySelector("#btn-router-retry")?.addEventListener("click",()=>e()))}else{let h=l||document.getElementById("app");h&&(h.innerHTML='<div class="empty-state"><h2>404 - Halaman tidak ditemukan</h2></div>')}}window.addEventListener("hashchange",e),e()}var kt;function ji(){return kt||(kt=document.createElement("div"),kt.id="toast-container",document.body.appendChild(kt)),kt}function ga(e,t="info",n=3500){let a=ji(),o=["success","error","warning","info"].includes(t)?t:"info",s=document.createElement("div");s.className=`toast toast-${o}`;let l={success:"\u2713",error:"\u2715",warning:"\u26A0",info:"\u2139"},h=document.createElement("span");h.className="toast-icon",h.textContent=l[o]||"\u2139";let p=document.createElement("span");p.className="toast-msg",p.textContent=e==null?"":String(e),s.appendChild(h),s.appendChild(p),a.appendChild(s),requestAnimationFrame(()=>s.classList.add("show")),setTimeout(()=>{s.classList.remove("show"),setTimeout(()=>s.remove(),350)},n)}var ae=e=>ga(e,"success"),z=e=>ga(e,"error");var ct=e=>ga(e,"info");Fe();Ne();Ne();var Qe={};function Ht(e){if(Qe[e]){try{Qe[e].destroy()}catch{}delete Qe[e]}}function rn(){Object.keys(Qe).forEach(Ht)}var et=(e,t=0)=>{let n=Number(e);return isNaN(n)||e===null||e===void 0?t:n},Ui=(e,t="\u2014")=>{if(e==null||e==="")return t;let n=String(e).trim();return n===""||n==="[object Object]"?t:Y(n)};function Gi(e,t,n=800){if(!e)return;e._animFrame&&cancelAnimationFrame(e._animFrame);let a=Math.max(0,Math.round(et(t))),o=e.textContent?e.textContent.replace(/[^\d]/g,""):"",s=o!==""&&parseInt(o)||0;if(a===s){e.textContent=a.toLocaleString("id-ID"),e._animFrame=null;return}let l=Date.now(),h=()=>{let p=Math.min((Date.now()-l)/n,1),c=1-Math.pow(1-p,3);e.textContent=Math.round(s+c*(a-s)).toLocaleString("id-ID"),p<1?e._animFrame=requestAnimationFrame(h):(e.textContent=a.toLocaleString("id-ID"),e._animFrame=null)};e._animFrame=requestAnimationFrame(h)}var ut={family:"Inter",size:10,weight:"500"},Xe="#94A3B8",Ot="rgba(255, 255, 255, 0.06)",ba=["#3B82F6","#10B981","#F59E0B","#EF4444","#8B5CF6","#06B6D4"],Ji=()=>window.innerWidth<768;function ln(e={}){return{responsive:!0,maintainAspectRatio:!1,animation:{duration:600,easing:"easeOutQuart"},plugins:{legend:{position:Ji()?"bottom":"top",labels:{font:ut,color:Xe,usePointStyle:!0,padding:10,boxWidth:8,boxHeight:8}},tooltip:{mode:"index",intersect:!1,backgroundColor:"#1E293B",titleColor:"#F8FAFC",bodyColor:"#94A3B8",borderColor:"rgba(255, 255, 255, 0.1)",borderWidth:1,padding:10,boxPadding:4,cornerRadius:8,bodyFont:ut,titleFont:{...ut,weight:"700"}}},scales:{x:{grid:{color:Ot},ticks:{font:ut,color:Xe,maxRotation:0}},y:{grid:{color:Ot},ticks:{font:ut,color:Xe},beginAtZero:!0}},...e}}var Vi=()=>Array(5).fill(0).map(()=>`
  <div class="kpi-card" style="pointer-events:none;">
    <div class="skeleton" style="width:46px;height:46px;border-radius:12px;flex-shrink:0;"></div>
    <div style="flex:1;">
      <div class="skeleton skeleton-text" style="width:50px;height:24px;margin-bottom:6px;"></div>
      <div class="skeleton skeleton-text" style="width:85px;height:12px;"></div>
    </div>
  </div>`).join("");function on(e=4){return Array(e).fill(0).map((t,n)=>`<div class="skeleton skeleton-text" style="height:36px;margin-bottom:${n<e-1?"8px":"0"};border-radius:8px;"></div>`).join("")}async function Ue(e,t,n=8e3){try{let a=new AbortController,o=setTimeout(()=>a.abort(),n),s=await X(e,{signal:a.signal}).catch(()=>null);if(clearTimeout(o),!s||!s.ok)return t;let l=s.data;return l?l.data!==void 0?l.data??t:l:t}catch{return t}}function Yi(e){if(!e||!e.isConnected)return;["skel-donut","skel-trend","skel-jadwal"].forEach(o=>{let s=document.getElementById(o);s&&(s.style.display="none")}),["chart-donut","chart-trend","chart-jadwal"].forEach(o=>{let s=document.getElementById(o);s&&s.style.display==="none"&&(s.style.display="block")});let t=document.getElementById("kpi-row");t&&t.querySelector(".skeleton")&&pn({});let n=document.getElementById("table-recent-activity");n&&n.querySelector(".skeleton")&&(n.innerHTML='<div class="chart-empty">Belum ada aktivitas terbaru</div>');let a=document.getElementById("list-upcoming-schedule");a&&a.querySelector(".skeleton")&&(a.innerHTML='<div class="chart-empty">Belum ada jadwal kegiatan</div>')}async function dn(e){if(e||(e=document.getElementById("main-content")),!e)return;rn(),e._dashRefresh&&clearInterval(e._dashRefresh),e._skelTimeout&&clearTimeout(e._skelTimeout);let t=!1,n=()=>!t&&e&&e.isConnected&&!!e.querySelector("#dash-root"),a=()=>{t=!0,e._skelTimeout&&(clearTimeout(e._skelTimeout),e._skelTimeout=null),e._dashRefresh&&(clearInterval(e._dashRefresh),e._dashRefresh=null),rn()};window.addEventListener("hashchange",a,{once:!0});let o=de(),s=String(o?.role||"").toLowerCase()==="superadmin",l=new Date,h=String(l.getDate()).padStart(2,"0"),p=l.toLocaleDateString("id-ID",{month:"short"}),c=l.getFullYear(),E=`01 ${p} \u2013 ${h} ${p} ${c}`;return e.innerHTML=`
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
              <option value="${c}" selected>${c}</option>
              <option value="${c+1}">${c+1}</option>
            </select>
          </div>
          <div class="dash-filter-pill dash-date-display">
            <svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><rect x="3" y="4" width="18" height="18" rx="2"/><line x1="16" y1="2" x2="16" y2="6"/><line x1="8" y1="2" x2="8" y2="6"/><line x1="3" y1="10" x2="21" y2="10"/></svg>
            <span id="dash-date-range">${E}</span>
          </div>
          <button id="btn-dash-refresh" class="dash-refresh-btn" title="Muat ulang data">
            <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><path d="M21 12a9 9 0 1 1-9-9c2.52 0 4.93 1 6.74 2.74L21 8"/><path d="M21 3v5h-5"/></svg>
            Refresh Data
          </button>
        </div>
      </div>

      <!-- Top KPI Row (5 Cards) -->
      <div class="kpi-row" id="kpi-row">${Vi()}</div>

      <!-- Middle Charts Row (3 Columns) -->
      <div class="charts-row">
        <!-- Chart 1: Tren Kegiatan Operasional -->
        <div class="chart-card">
          <div class="chart-card-header">
            <div>
              <div class="chart-card-title">Tren Kegiatan Operasional</div>
            </div>
            <select id="filter-jadwal-period" class="dash-header-select" style="background:#162032; border:1px solid rgba(255,255,255,0.08); border-radius:6px; padding:4px 8px; font-size:0.75rem;" aria-label="Filter Periode">
              <option value="month">Per Bulan</option>
            </select>
          </div>
          <div class="chart-canvas-wrap" style="height:150px;position:relative;margin-top:4px">
            <div id="skel-jadwal" class="skeleton" style="position:absolute;inset:0;border-radius:12px"></div>
            <canvas id="chart-jadwal" style="display:none"></canvas>
          </div>
          <div id="jadwal-legend" style="display:flex;justify-content:center;gap:12px;margin-top:6px;font-size:0.7rem;font-weight:600;color:var(--text-2);flex-wrap:wrap;">
            <div style="display:flex;align-items:center;gap:4px"><div style="width:8px;height:8px;border-radius:50%;background:#3B82F6"></div> Inspeksi</div>
            <div style="display:flex;align-items:center;gap:4px"><div style="width:8px;height:8px;border-radius:50%;background:#06B6D4"></div> General Cleaning</div>
            <div style="display:flex;align-items:center;gap:4px"><div style="width:8px;height:8px;border-radius:50%;background:#F59E0B"></div> Deep Cleaning</div>
            <div style="display:flex;align-items:center;gap:4px"><div style="width:8px;height:8px;border-radius:50%;background:#A855F7"></div> Fogging</div>
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
          <div style="text-align:center; font-size:0.72rem; color:var(--text-3); display:flex; align-items:center; justify-content:center; gap:6px;">
            <svg width="12" height="12" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><rect x="3" y="4" width="18" height="18" rx="2"/><line x1="16" y1="2" x2="16" y2="6"/><line x1="8" y1="2" x2="8" y2="6"/></svg>
            <span id="donut-period-text">Distribusi Kategori Permasalahan</span>
          </div>
        </div>

        <!-- Chart 3: Tren Permasalahan 12 Bulan -->
        <div class="chart-card">
          <div class="chart-card-header">
            <div class="chart-card-title">Tren Permasalahan 12 Bulan</div>
            <div style="display:flex;align-items:center;gap:12px;font-size:0.7rem;font-weight:600;color:var(--text-2)">
               <div style="display:flex;align-items:center;gap:5px"><div style="width:8px;height:8px;border-radius:50%;background:#EF4444"></div> Open</div>
               <div style="display:flex;align-items:center;gap:5px"><div style="width:8px;height:8px;border-radius:50%;background:#06B6D4"></div> Closed</div>
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
          <div id="table-recent-activity" class="dash-table-wrap" style="height:230px;overflow-y:auto">
            ${on(4)}
          </div>
        </div>

        <!-- Jadwal Kegiatan Mendatang -->
        <div class="chart-card">
          <div class="chart-card-header">
            <div class="chart-card-title">Jadwal Kegiatan Mendatang</div>
            <a href="#/calendar" class="chart-link">Lihat Kalender &rarr;</a>
          </div>
          <div id="list-upcoming-schedule" style="height:230px;overflow-y:auto;display:flex;flex-direction:column;gap:4px">
            ${on(4)}
          </div>
        </div>
      </div>

    </div>
  `,document.getElementById("btn-dash-refresh")?.addEventListener("click",()=>{n()&&sn(e,n)}),document.getElementById("filter-jadwal-year")?.addEventListener("change",async C=>{let T=C.target.value,u=document.getElementById("skel-jadwal"),F=document.getElementById("chart-jadwal");u&&(u.style.display="block"),F&&(F.style.display="none");let R=await Ue(`/api/dashboard/schedule-chart?year=${T}`,{},8e3);if(n())try{cn(R)}catch(K){console.warn("ScheduleChart render error:",K),Ze("skel-jadwal","chart-jadwal")}}),e._skelTimeout=setTimeout(()=>Yi(e),5e3),await sn(e,n),a}async function sn(e,t=()=>e&&e.isConnected){e&&e._skelTimeout&&(clearTimeout(e._skelTimeout),e._skelTimeout=null);let n=document.getElementById("filter-jadwal-year")?.value||new Date().getFullYear(),[a,o,s,l,h,p,c,E]=await Promise.all([Ue("/api/dashboard/kpi",{},8e3),Ue("/api/dashboard/issues-trend",{},8e3),Ue("/api/dashboard/issues-summary",{},8e3),Ue("/api/dashboard/stats",{},8e3),Ue("/api/dashboard/calendar",[],8e3),Ue(`/api/dashboard/schedule-chart?year=${n}`,{},8e3),Ue("/api/dashboard/dropdown-aggregates",{},8e3),Ue("/api/dashboard/activity-log",[],8e3)]);if(t()){window.dashboardAggregates=c||{};try{pn(a)}catch(C){console.warn("KPI render:",C)}try{cn(p)}catch(C){console.warn("ScheduleChart render:",C),Ze("skel-jadwal","chart-jadwal")}try{Wi(Array.isArray(s?.by_category)?s.by_category:[])}catch(C){console.warn("Donut render:",C),Ze("skel-donut","chart-donut")}try{Qi(o)}catch(C){console.warn("Trend render:",C),Ze("skel-trend","chart-trend")}try{Xi(E)}catch(C){console.warn("RecentActivity render:",C)}try{let C=l?.upcoming_schedule||[];Zi(C,h)}catch(C){console.warn("UpcomingSchedule render:",C)}}}function pn(e){let t=document.getElementById("kpi-row");if(!t)return;e=e||{};let n=(o,s,l="&ndash; 0 (0%)")=>{if(o=et(o),s=et(s),!s||s<=0)return`<span class="kpi-trend neutral">${l}</span>`;let h=o-s,p=Math.abs(Math.round(h/s*1e3)/10).toLocaleString("id-ID");return h>0?`<span class="kpi-trend up">&uarr; +${h} (${p}%)</span>`:h<0?`<span class="kpi-trend down">&darr; ${h} (-${p}%)</span>`:'<span class="kpi-trend neutral">&ndash; 0 (0%)</span>'},a=[{iconSvg:'<svg width="22" height="22" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" viewBox="0 0 24 24"><path d="M17 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2"/><circle cx="9" cy="7" r="4"/><path d="M23 21v-2a4 4 0 0 0-3-3.87"/><path d="M16 3.13a4 4 0 0 1 0 7.75"/></svg>',label:"Karyawan Aktif",val:e.employees?.current??0,delta:n(e.employees?.current,e.employees?.prev),color:"kpi-purple",href:"#/employees?dash_filter=active"},{iconSvg:'<svg width="22" height="22" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" viewBox="0 0 24 24"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/><line x1="16" y1="13" x2="8" y2="13"/><line x1="16" y1="17" x2="8" y2="17"/></svg>',label:"Kontrak Aktif",val:e.contracts?.current??0,delta:n(e.contracts?.current,e.contracts?.prev),color:"kpi-teal",href:"#/contracts?dash_filter=active"},{iconSvg:'<svg width="22" height="22" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" viewBox="0 0 24 24"><path d="M2 9h20L19 3H5L2 9z"/><path d="M4 9v11a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V9"/><path d="M10 22V12h4v10"/></svg>',label:"Cabang Operasional",val:e.branches?.current??0,delta:'<span class="kpi-trend neutral">&ndash; Area aktif</span>',color:"kpi-amber",href:"#/branches"},{iconSvg:'<svg width="22" height="22" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" viewBox="0 0 24 24"><rect x="3" y="4" width="18" height="18" rx="2" ry="2"/><line x1="16" y1="2" x2="16" y2="6"/><line x1="8" y1="2" x2="8" y2="6"/><line x1="3" y1="10" x2="21" y2="10"/></svg>',label:"Jadwal Kegiatan",val:e.schedule?.current??0,delta:'<span class="kpi-trend neutral">&ndash; Total jadwal</span>',color:"kpi-blue",href:"#/timeline"},{iconSvg:'<svg width="22" height="22" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" viewBox="0 0 24 24"><path d="M10.29 3.86L1.82 18a2 2 0 0 0 1.71 3h16.94a2 2 0 0 0 1.71-3L13.71 3.86a2 2 0 0 0-3.42 0z"/><line x1="12" y1="9" x2="12" y2="13"/><line x1="12" y1="17" x2="12.01" y2="17"/></svg>',label:"Permasalahan Open",val:e.issues?.current??0,delta:n(e.issues?.current,e.issues?.prev),color:"kpi-red",href:"#/issues?dash_filter=open"}];t.innerHTML=a.map(o=>`
    <a href="${o.href}" class="kpi-card ${o.color}">
      <div class="kpi-icon-wrap">
        ${o.iconSvg}
      </div>
      <div class="kpi-body">
        <div class="kpi-val-row">
          <div class="kpi-value" data-target="${o.val}">${o.val}</div>
          ${o.delta}
        </div>
        <div class="kpi-label">${o.label}</div>
      </div>
    </a>
  `).join(""),t.querySelectorAll(".kpi-value").forEach(o=>{Gi(o,parseInt(o.dataset.target)||0)})}function cn(e){Ze("skel-jadwal","chart-jadwal");let t=document.getElementById("chart-jadwal");if(!t)return;Ht("jadwal"),e=e||{};let n=["Jan","Feb","Mar","Apr","Mei","Jun","Jul","Agu","Sep","Okt","Nov","Des"];if(!Object.values(e).some(p=>Array.isArray(p)&&p.some(c=>c>0))){ha(t,"Belum ada data jadwal kegiatan");return}let o=e["Inspeksi Hygiene"]||Array(12).fill(0),s=e["General Cleaning"]||Array(12).fill(0),l=e["Deep Cleaning"]||Array(12).fill(0),h=e.Fogging||Array(12).fill(0);Qe.jadwal=new Chart(t,{type:"bar",data:{labels:n,datasets:[{label:"Inspeksi",data:o,backgroundColor:"#3B82F6",borderRadius:3},{label:"General Cleaning",data:s,backgroundColor:"#06B6D4",borderRadius:3},{label:"Deep Cleaning",data:l,backgroundColor:"#F59E0B",borderRadius:3},{label:"Fogging",data:h,backgroundColor:"#A855F7",borderRadius:3}]},options:ln({plugins:{legend:{display:!1}},datasets:{bar:{barPercentage:.75,categoryPercentage:.85}},scales:{x:{stacked:!0,grid:{display:!1},ticks:{font:{family:"Inter",size:9},color:Xe,maxRotation:0}},y:{stacked:!0,grid:{color:Ot},ticks:{font:{family:"Inter",size:9},color:Xe},min:0}}})})}function Wi(e){Ze("skel-donut","chart-donut");let t=document.getElementById("chart-donut"),n=document.getElementById("donut-legend");if(!t||!n)return;Ht("donut");let a=(e||[]).filter(p=>et(p.count)>0);if(!a.length){ha(t,"Belum ada data kategori");return}let o=a.map(p=>Ui(p.category,"Lainnya")),s=a.map(p=>et(p.count)),l=s.reduce((p,c)=>p+c,0);n.innerHTML=a.map((p,c)=>{let E=ba[c%ba.length],C=l>0?Math.round(p.count/l*100):0;return`
      <div style="display:flex;align-items:center;gap:8px;margin-bottom:6px;">
        <div style="width:8px;height:8px;border-radius:50%;background:${E};flex-shrink:0;"></div>
        <div style="display:flex;align-items:baseline;gap:4px;font-size:0.75rem;">
          <span style="font-weight:600;color:var(--text-1);">${o[c]}</span>
          <span style="color:var(--text-2);font-weight:700;">${p.count}</span>
          <span style="font-size:0.68rem;color:var(--text-3);">(${C}%)</span>
        </div>
      </div>
    `}).join("");let h={id:"centerText",beforeDraw:function(p){let c=p.width,E=p.height,C=p.ctx;C.restore();let T=(E/85).toFixed(2);C.font="bold "+T+"em Inter",C.textBaseline="middle",C.fillStyle="#F8FAFC";let u=l.toString(),F=Math.round((c-C.measureText(u).width)/2),R=E/2;C.fillText(u,F,R-6),C.font="600 "+(T*.38).toFixed(2)+"em Inter",C.fillStyle="#94A3B8";let K="Total",I=Math.round((c-C.measureText(K).width)/2);C.fillText(K,I,R+10),C.save()}};Qe.donut=new Chart(t,{type:"doughnut",data:{labels:o,datasets:[{data:s,backgroundColor:ba,borderWidth:2,borderColor:"#111827",hoverBorderColor:"#111827"}]},options:{responsive:!0,maintainAspectRatio:!1,animation:{duration:600},plugins:{legend:{display:!1},tooltip:{backgroundColor:"#1E293B",titleColor:"#F8FAFC",bodyColor:"#94A3B8",borderColor:"rgba(255, 255, 255, 0.1)",borderWidth:1,cornerRadius:8,bodyFont:ut,callbacks:{label:p=>` ${p.label}: ${p.parsed} kasus`}}},cutout:"75%"},plugins:[h]})}function Qi(e){Ze("skel-trend","chart-trend");let t=document.getElementById("chart-trend");if(!t)return;Ht("trend"),e=e||{};let n=["Jan","Feb","Mar","Apr","Mei","Jun","Jul","Agu","Sep","Okt","Nov","Des"],a=(e.labels||[]).map(l=>{if(!l||typeof l!="string")return"";try{let[,h]=l.split("-");return n[Number(h)-1]||h}catch{return l}}),o=(e.open||[]).map(l=>et(l)),s=(e.closed||[]).map(l=>et(l));if(!a.length){ha(t,"Belum ada data trend permasalahan");return}Qe.trend=new Chart(t,{type:"line",data:{labels:a,datasets:[{label:"Open",data:o,borderColor:"#EF4444",backgroundColor:"rgba(239,68,68,.12)",fill:!0,tension:.4,pointRadius:3.5,pointHoverRadius:6,pointBackgroundColor:"#EF4444",borderWidth:2},{label:"Closed",data:s,borderColor:"#06B6D4",backgroundColor:"rgba(6,182,212,.12)",fill:!0,tension:.4,pointRadius:3.5,pointHoverRadius:6,pointBackgroundColor:"#06B6D4",borderWidth:2}]},options:ln({plugins:{legend:{display:!1}},scales:{x:{grid:{display:!1},ticks:{font:{family:"Inter",size:9},color:Xe,maxRotation:0}},y:{grid:{color:Ot},ticks:{font:{family:"Inter",size:9},color:Xe},beginAtZero:!0}}})})}function Xi(e){let t=document.getElementById("table-recent-activity");if(!t)return;let n=(e||[]).slice(0,6);if(!n.length){t.innerHTML='<div class="chart-empty">Belum ada aktivitas terbaru</div>';return}let a={issue:{icon:'<svg width="12" height="12" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24"><path d="M10.29 3.86L1.82 18a2 2 0 0 0 1.71 3h16.94a2 2 0 0 0 1.71-3L13.71 3.86a2 2 0 0 0-3.42 0z"/><line x1="12" y1="9" x2="12" y2="13"/><line x1="12" y1="17" x2="12.01" y2="17"/></svg>',bg:"#EF4444",text:"#FEF2F2",module:"Permasalahan",modBg:"rgba(239,68,68,0.18)",modCol:"#F87171"},contract:{icon:'<svg width="12" height="12" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/></svg>',bg:"#06B6D4",text:"#ECFEFF",module:"Data Kontrak",modBg:"rgba(6,182,212,0.18)",modCol:"#22D3EE"},employee:{icon:'<svg width="12" height="12" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>',bg:"#F59E0B",text:"#FFFBEB",module:"Master Karyawan",modBg:"rgba(37,99,235,0.18)",modCol:"#60A5FA"},training:{icon:'<svg width="12" height="12" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/></svg>',bg:"#3B82F6",text:"#EFF6FF",module:"Training",modBg:"rgba(16,185,129,0.18)",modCol:"#34D399"},inspection:{icon:'<svg width="12" height="12" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24"><path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"/><polyline points="22 4 12 14.01 9 11.01"/></svg>',bg:"#10B981",text:"#ECFDF5",module:"Inspeksi",modBg:"rgba(245,158,11,0.18)",modCol:"#FBBF24"},schedule:{icon:'<svg width="12" height="12" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24"><line x1="12" y1="5" x2="12" y2="19"/><line x1="5" y1="12" x2="19" y2="12"/></svg>',bg:"#10B981",text:"#ECFDF5",module:"Jadwal Kegiatan",modBg:"rgba(99,102,241,0.18)",modCol:"#A5B4FC"}};t.innerHTML=`
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
        ${n.map(o=>{let s=o.created_at?String(o.created_at).includes(" ")&&!String(o.created_at).includes("T")?`${String(o.created_at).replace(" ","T")}Z`:String(o.created_at):null,l=s?new Date(s):null,h=l&&!isNaN(l.getTime())?l.toLocaleTimeString("id-ID",{hour:"2-digit",minute:"2-digit",timeZone:"Asia/Jakarta"}).replace(/\./g,":"):"\u2014",p=a[o.type]||{icon:"\u{1F4CC}",bg:"#3B82F6",text:"#fff",module:"Operasional",modBg:"rgba(59,130,246,0.18)",modCol:"#60A5FA"},c=o.user_name||(o.branch?`Area ${o.branch}`:"Sistem FCMS"),E=(c||"F")[0].toUpperCase();return`
            <tr>
              <td class="act-time">${h}</td>
              <td>
                <div class="act-desc">
                  <span class="act-icon" style="background:${p.bg};color:${p.text}">${p.icon}</span>
                  <span title="${Y(o.label)}">${Y(o.label)}</span>
                </div>
              </td>
              <td>
                <span class="act-modul-pill" style="background:${p.modBg};color:${p.modCol}">
                  ${p.module}
                </span>
              </td>
              <td>
                <div class="act-user">
                  <span class="act-user-avatar" style="background:${p.bg}">
                    ${E}
                  </span>
                  <span>${Y(c)}</span>
                </div>
              </td>
            </tr>
          `}).join("")}
      </tbody>
    </table>
  `}function Zi(e,t){let n=document.getElementById("list-upcoming-schedule");if(!n)return;let a=Array.isArray(e)&&e.length>0?e:t||[];if(a=a.slice(0,5),!a.length){n.innerHTML='<div class="chart-empty">Belum ada jadwal kegiatan mendatang</div>';return}let o=["Min","Sen","Sel","Rab","Kam","Jum","Sab"],s=["Jan","Feb","Mar","Apr","Mei","Jun","Jul","Agu","Sep","Okt","Nov","Des"];n.innerHTML=a.map(l=>{let h=l.target_date||l.event_date||"",p="\u2014",c="\u2014";if(h)try{let R=new Date(h);isNaN(R.getTime())||(p=o[R.getDay()]||"Sen",c=`${R.getDate()} ${s[R.getMonth()]||""}`)}catch{}let E=l.title||l.activity_type||"Kegiatan Operasional",C=l.branch_name||"Cabang Operasional",T="event-pill-insp",u="Inspeksi",F=(E+" "+(l.activity_type||"")).toLowerCase();return F.includes("general cleaning")||F.includes("gc")?(T="event-pill-gc",u="GC"):F.includes("deep cleaning")||F.includes("dc")?(T="event-pill-dc",u="Deep Cleaning"):F.includes("fogging")?(T="event-pill-fog",u="Fogging"):F.includes("inspeksi")&&(T="event-pill-insp",u="Inspeksi"),`
      <div class="event-item">
        <div class="event-date-box">
          <span class="event-date-day">${p}</span>
          <span class="event-date-num">${c}</span>
        </div>
        <div class="event-info">
          <div class="event-title">${Y(E)}</div>
          <div class="event-branch">
            <svg width="12" height="12" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M21 10c0 7-9 13-9 13s-9-6-9-13a9 9 0 0118 0z"/><circle cx="12" cy="10" r="3"/></svg>
            <span>${Y(C)}</span>
          </div>
        </div>
        <div class="event-pill ${T}">${u}</div>
      </div>
    `}).join("")}function Ze(e,t){let n=document.getElementById(e),a=document.getElementById(t);if(n&&(n.style.display="none"),a){a.style.display="block";let o=a.parentElement;if(o){let s=o.querySelector(".chart-empty");s&&s.remove()}}}function ha(e,t="Belum ada data"){if(!e)return;e.style.display="none";let n=e.parentElement;if(!n)return;if(!n.querySelector(".chart-empty")){let o=document.createElement("div");o.className="chart-empty",o.textContent=t,n.appendChild(o)}}async function un(e){document.getElementById("app").innerHTML=`
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
  `;let t=document.getElementById("login-form"),n=document.getElementById("login-error"),a=document.getElementById("login-btn"),o=document.getElementById("toggle-password"),s=document.getElementById("login-password");o?.addEventListener("click",()=>{let l=s.type==="text";s.type=l?"password":"text",o.style.color=l?"":"var(--primary)"}),t?.addEventListener("submit",async l=>{l.preventDefault(),n.style.display="none";let h=t.username.value.trim(),p=t.password.value;if(!h||!p){n.textContent="Username dan password wajib diisi.",n.style.display="block";return}a.querySelector(".btn-text").style.display="none",a.querySelector(".btn-spinner").style.display="",a.disabled=!0;try{let c=await X("/api/auth/login",{method:"POST",body:JSON.stringify({username:h,password:p})});c.ok&&c.data.success?(tn(c.data.data.token),vt(c.data.data.csrf_token||null),Rt(c.data.data.user),ae("Login berhasil! Selamat datang \u{1F44B}"),window.dispatchEvent(new Event("fm:login"))):(n.textContent=c.data.error||"Username atau password salah.",n.style.display="block",a.classList.add("shake"),setTimeout(()=>a.classList.remove("shake"),600))}catch{n.textContent="Gagal terhubung ke server. Periksa koneksi internet.",n.style.display="block"}finally{a.querySelector(".btn-text").style.display="",a.querySelector(".btn-spinner").style.display="none",a.disabled=!1}})}var er=Object.freeze([["001","Pondok Bambu"],["002","Kemang"],["003","Depok"],["004","Cipete"],["005","Greenville"],["006","Kelapa Gading MOI"],["007","Bekasi"],["008","PIK"],["009","Bintaro"],["010","Bogor"],["011","Kuningan"],["012","BSD"],["013","Bandung Gatsu"],["014","Sunter"],["015","Gading Serpong"],["016","Cibubur"],["017","Bandung Setrasari"],["018","Harapan Indah"],["019","Cinere"],["020","Surabaya"],["021","Kelapa Gading Boulevard"],["022","Cikarang"],["023","Semarang"],["024","Grand Wisata"],["025","Bogor Baru"],["026","Tangerang City"],["027","Rawamangun"],["028","Taman Palem"],["029","Karawang"],["030","Ciputat"],["031","Sawangan"],["032","Cawang"],["033","Taman Mini"],["034","Puri Indah"],["035","Ciledug"],["036","Jagakarsa"],["037","Bali"],["038","Greenlake"],["039","Bandung Mekarwangi"],["040","Bekasi Timur"],["041","Karawaci"],["042","Muara Karang"],["043","Tebet"],["044","Depok GDC"],["045","FX Mall"],["046","Kaliurang Jogja"],["047","Bogor Tajur"],["048","PIK 2"],["049","Surabaya Merr"],["050","Cibinong"],["051","Bintaro Sektor 9"],["052","BSD Rawa Buntu"],["053","Alam Sutera"],["054","Bogor Gunung Batu"],["055","Gajah Mada"],["056","Citra Garden 2"],["057","Cikupa"],["058","Solo Baru"],["059","Malang Suhat"],["060","Palmerah"],["061","Cirebon"],["062","Summarecon Bekasi"],["063","Pasar Minggu"],["064","Pondok Indah"],["065","Kebon Sirih"],["066","Bandung KBP"],["067","Agora Mall"],["068","Kota Wisata"],["069","Depok Pancoran Mas"],["070","Bekasi Kota Bintang"]].map(([e,t])=>Object.freeze({code:e,name:t,full_name:`${e}. ${t}`,is_operational:!0,branch_scope:"operational",parent_branch_code:null}))),tr=Object.freeze([Object.freeze({code:"000",name:"Management",full_name:"000. Management",is_operational:!1,branch_scope:"hq",parent_branch_code:null}),Object.freeze({code:"A01",name:"Aesthica",full_name:"A01. Aesthica",is_operational:!1,branch_scope:"joined_kemang",parent_branch_code:"002"})]),gn=new Map([...er,...tr].map(e=>[e.code,e])),ar={"069":"Bekasi Kota Bintang","070":"J7"};function Ve(e){return String(e??"").trim()}function fa(e){let t=Ve(e);return/^(00[1-9]|0[1-6][0-9]|070)$/.test(t)}function mn(e){return gn.has(Ve(e))}function nr(e){let t=Ve(e?.code),n=String(e?.full_name||e?.name||"").trim();return/^T\d+[A-Z0-9-]*$/i.test(t)||/^TEST_/i.test(n)||/^U558D/i.test(t)||/\[UAT-/i.test(n)}function ir(e){let t=Ve(e?.code),n=gn.get(t),a=String(e?.name||"").trim(),o=String(e?.full_name||a||`Cabang #${e?.id}`).trim(),s=ar[t];return n&&s&&a.toLowerCase()===s.toLowerCase()?n.full_name:o}function bn(e=[]){let t=Array.isArray(e)?e.filter(s=>!s.deleted_at&&!nr(s)):[],a=t.filter(s=>mn(s.code)).length>=65?t.filter(s=>mn(s.code)):t,o=s=>{let l=Ve(s);return l==="000"?0:fa(l)?parseInt(l,10):l==="A01"?900:1e3};return[...a].sort((s,l)=>o(s.code)-o(l.code)||Ve(s.code).localeCompare(Ve(l.code))||(s.id||0)-(l.id||0)).map(s=>({value:s.id,label:ir(s),code:Ve(s.code),is_operational:fa(s.code),branch_scope:s.code==="000"?"hq":s.code==="A01"?"joined_kemang":fa(s.code)?"operational":"other"}))}var Ut=100,wt=null,St=null;async function Et(e,t={},n={}){let a=Math.min(Ut,Math.max(1,Number(n.limit)||Ut)),o=Math.max(1,Number(n.maxPages)||500),s=[],l=new Set;for(let h=1;h<=o;h++){let p=new URLSearchParams;if(p.set("page",String(h)),p.set("limit",String(a)),t&&typeof t=="object")for(let[u,F]of Object.entries(t))F!=null&&F!==""&&F!=="all"&&p.set(u,String(F));let c=await X(`${e}?${p.toString()}`);if(!c.ok)throw new Error(c.data?.error||`Gagal memuat data (HTTP ${c.status||500})`);let E=Array.isArray(c.data?.data)?c.data.data:Array.isArray(c.data)?c.data:[];for(let u of E){if(u&&u.id!==void 0&&u.id!==null){if(l.has(u.id))continue;l.add(u.id)}s.push(u)}let C=c.data?.pagination,T=C&&Number(C.pages)>=0?Number(C.pages):null;if(T!==null){if(h>=T)break}else if(E.length<a)break;if(E.length===0)break}return s}async function ue(e,t={},n={}){return Et(e,t,{...n,limit:Ut})}async function hn({search:e="",status:t="Aktif",branch_id:n="",division:a="",page:o=1,limit:s=20}={}){let l=Math.min(Ut,Math.max(1,Number(s)||20)),h=new URLSearchParams({page:String(Math.max(1,Number(o)||1)),limit:String(l)});e&&h.set("search",e),t&&t!=="all"&&h.set("status",t),n&&n!=="all"&&h.set("branch_id",n),a&&a!=="all"&&h.set("division",a);let p=await X(`/api/employees/selector?${h.toString()}`);if(!p.ok)throw new Error(p.data?.error||`Gagal mencari karyawan (HTTP ${p.status||500})`);return{data:Array.isArray(p.data?.data)?p.data.data:[],pagination:p.data?.pagination||{total:0,page:1,limit:l,pages:0}}}async function Ie(e=!1){return wt&&!e||(wt=(await Et("/api/employees/selector",{status:"Aktif"})).map(n=>({value:n.id,id:n.id,label:n.full_name,full_name:n.full_name,branch_id:n.branch_id??null,branch_name:n.branch_name||"",division:n.division||"",status:n.status||""}))),wt}async function Gt(e=!1){return(await Ie(e)).map(n=>({value:n.label,label:n.label}))}async function ve(e=!1){if(St&&!e)return St;let t=await X("/api/branches?all=1");if(!t.ok)throw new Error(t.data?.error||"Gagal memuat data cabang");return St=bn(t.data?.data||[]),St}function Ct(){wt=null,St=null}Fe();Fe();Ne();var mt=100;function fn(e,t,n,a=mt){return n?e.has(t)?{accepted:!0,count:e.size}:e.size>=a?{accepted:!1,count:e.size,reason:`Maksimum ${a} item dapat dipilih untuk penghapusan bulk`}:(e.add(t),{accepted:!0,count:e.size}):(e.delete(t),{accepted:!0,count:e.size})}function xn(e,t,n,a=mt){let o=(t||[]).map(l=>typeof l=="object"&&l!==null?l.id:l).filter(l=>l!=null);if(!n)return o.forEach(l=>e.delete(l)),{accepted:!0,count:e.size};let s=o.filter(l=>!e.has(l));return e.size+s.length>a?{accepted:!1,count:e.size,reason:`Maksimum ${a} item dapat dipilih untuk penghapusan bulk (saat ini ${e.size} item sudah dipilih)`}:(s.forEach(l=>e.add(l)),{accepted:!0,count:e.size})}function yn({columns:e,data:t,onEdit:n,onDelete:a,onView:o,actions:s=[],emptyText:l="Tidak ada data",bulkSelect:h=null}){let p=document.createElement("div");if(p.className="table-wrapper",!t||t.length===0)return p.innerHTML=`<div class="empty-state"><p>${l}</p></div>`,p;let c=document.createElement("table");c.className="data-table";let E=document.createElement("thead"),C=document.createElement("tr");if(h){let F=document.createElement("th");F.style.width="40px",F.style.textAlign="center";let R=document.createElement("input");R.type="checkbox",R.id="select-all-checkbox",R.title="Pilih semua";let K=h.maxSelect||1/0;R.addEventListener("change",()=>{if(Number.isFinite(K)){let I=xn(h.selectedIds,t,R.checked,K);I.accepted||(R.checked=!1,typeof h.onLimitExceeded=="function"&&h.onLimitExceeded(I.reason));let N=p.querySelectorAll(".row-checkbox");t.forEach((w,M)=>{N[M]&&(N[M].checked=h.selectedIds.has(w.id))})}else t.forEach(I=>{R.checked?h.selectedIds.add(I.id):h.selectedIds.delete(I.id)}),p.querySelectorAll(".row-checkbox").forEach(I=>I.checked=R.checked);h.onToggle()}),F.appendChild(R),C.appendChild(F)}if(e.forEach(F=>{let R=document.createElement("th");R.textContent=F.label,F.width&&(R.style.width=F.width),C.appendChild(R)}),n||a||o||s.length>0){let F=document.createElement("th");F.textContent="Aksi",F.style.width="120px",C.appendChild(F)}E.appendChild(C),c.appendChild(E);let T=document.createElement("tbody"),u=h?.maxSelect||1/0;return t.forEach(F=>{let R=document.createElement("tr");if(h){let K=document.createElement("td");K.style.textAlign="center",K.style.width="40px";let I=document.createElement("input");I.type="checkbox",I.className="row-checkbox",I.checked=h.selectedIds.has(F.id),I.addEventListener("change",()=>{if(Number.isFinite(u)){let N=fn(h.selectedIds,F.id,I.checked,u);if(!N.accepted)I.checked=!1,typeof h.onLimitExceeded=="function"&&h.onLimitExceeded(N.reason);else if(!I.checked){let w=document.getElementById("select-all-checkbox");w&&(w.checked=!1)}}else if(I.checked)h.selectedIds.add(F.id);else{h.selectedIds.delete(F.id);let N=document.getElementById("select-all-checkbox");N&&(N.checked=!1)}h.onToggle()}),K.appendChild(I),R.appendChild(K)}if(e.forEach(K=>{let I=document.createElement("td");if(K.render){let N=K.render(F[K.key],F);N instanceof HTMLElement?I.appendChild(N):I.innerHTML=N?qt(N):""}else I.textContent=F[K.key]!==null&&F[K.key]!==void 0&&F[K.key]!==""?F[K.key]:"";K.nowrap&&(I.style.whiteSpace="nowrap"),R.appendChild(I)}),n||a||o||s.length>0){let K=document.createElement("td");K.className="actions-cell";let I=document.createElement("div");if(I.className="btn-group",o){let N=document.createElement("button");N.className="btn btn-xs btn-ghost",N.innerHTML="\u{1F441}",N.title="Lihat",N.addEventListener("click",()=>o(F)),I.appendChild(N)}if(n){let N=document.createElement("button");N.className="btn btn-xs btn-secondary",N.innerHTML="\u270F\uFE0F",N.title="Edit",N.addEventListener("click",()=>n(F)),I.appendChild(N)}if(s.forEach(N=>{let w=document.createElement("button");w.className=`btn btn-xs ${N.class||"btn-ghost"}`,w.innerHTML=N.icon||N.label,w.title=N.label,w.addEventListener("click",()=>N.handler(F)),I.appendChild(w)}),a){let N=document.createElement("button");N.className="btn btn-xs btn-ghost text-danger",N.innerHTML="\u{1F5D1}\uFE0F",N.title="Hapus",N.addEventListener("click",()=>a(F)),I.appendChild(N)}K.appendChild(I),R.appendChild(K)}T.appendChild(R)}),c.appendChild(T),p.appendChild(c),p}function _e({page:e,pages:t,total:n,limit:a,onPage:o}){if(!t||t<=1)return null;let s=Math.min(Math.max(1,Number(e)||1),t),l=Math.max(1,Number(a)||20),h=Math.max(0,Number(n)||0),p=document.createElement("div");p.className="pagination";let c=document.createElement("span");c.className="pagination-info";let E=h>0?(s-1)*l+1:0,C=h>0?Math.min(h,s*l):0;c.textContent=h>0?`Hal. ${s} / ${t} \u2022 ${E}-${C} dari ${h} data`:`Total: ${h} data`,p.appendChild(c);let T=document.createElement("div");T.className="pagination-btns";let u=(K,I,N=!1,w=!1)=>{let M=document.createElement("button");M.className=`btn btn-sm ${w?"btn-primary":"btn-ghost"} pagination-btn`,M.textContent=K,M.disabled=N,M.addEventListener("click",()=>o(I)),T.appendChild(M)};u("\xAB",1,s===1),u("\u2039",s-1,s===1);let F=Math.max(1,s-2),R=Math.min(t,s+2);for(let K=F;K<=R;K++)u(K,K,!1,K===s);return u("\u203A",s+1,s===t),u("\xBB",t,s===t),p.appendChild(T),p}function tt(e){if(e==null||e===""||typeof e=="boolean")return null;let t=Number(e);return Number.isInteger(t)&&t>=1?t:null}function Jt(e,t,n){let a=e&&typeof e.querySelector=="function"?e.querySelector(t):null;return tt(a?a.value:n)}function rr(e){return tt(e?.data?.data?.row_version??e?.data?.row_version)}async function Vt({res:e,isEdit:t,item:n,close:a,reloadData:o,editSuccessMessage:s,createSuccessMessage:l,defaultErrorMessage:h}){if(e.ok){let p=rr(e);return t&&p!==null&&n&&typeof n=="object"&&(n.row_version=p),ae(t?s:l),a(),await o(),!0}if(e.status===409){let p=e.data?.error||e.data?.message||"Data telah diubah oleh pengguna atau sesi lain.";return z(`Konflik data (HTTP 409): ${p} Silakan muat ulang / periksa data terbaru sebelum mengedit kembali.`),a(),await o(),!1}return z(e.data?.error||e.data?.message||h),!1}var kn=["FACILITY CARE","SECURITY"],lr=["Aktif","Tidak Aktif","Resign","Cut"];function wn(e){if(!e)return"-";try{let t=String(e).trim();if(/^\d{5}(\.\d+)?$/.test(t)){let a=new Date(Math.round((Number(t)-25569)*86400*1e3));if(!isNaN(a.getTime()))return a.toLocaleDateString("id-ID",{day:"numeric",month:"short",year:"numeric"})}let n=new Date(t);return isNaN(n.getTime())?t:n.toLocaleDateString("id-ID",{day:"numeric",month:"short",year:"numeric"})}catch{return e}}function Sn(e){if(!e||e==="-")return"\u{1F464}";let t=String(e).trim().split(/\s+/);return t.length===1?t[0].slice(0,2).toUpperCase():(t[0][0]+t[1][0]).toUpperCase()}async function En(e,t){let a=[],o=[],s="all",l="all",h="all",p="",c=1,E=0,C=1,T=!1,u={total:0,active:0,nonactive:0,fc:0,sec:0},F=null,R=!1,K=localStorage.getItem("fcms_employees_view_mode")||"card";(t?t.get("dash_filter"):null)==="active"&&(s="Aktif");let N=de(),w=ie(N,"employees","write"),M=ie(N,"employees","write"),k=ie(N,"employees","delete");try{let i=await ve();o=(Array.isArray(i)?i:[]).filter(r=>String(r.code||"").trim()==="000"||r.branch_scope==="hq"||/^(?:cabang\s+)?000\.\s*management$/i.test(String(r.label||r.full_name||"").trim()))}catch(i){console.error("Error fetching branches:",i)}function x(){let i={sort:"join_date_desc"};return p.trim()&&(i.search=p.trim()),h&&h!=="all"&&(i.branch_id=h),l&&l!=="all"&&(i.division=l),s&&s!=="all"&&(i.status=s),i}function _(){let i=e.querySelector("#emp-stat-total"),r=e.querySelector("#emp-stat-active"),d=e.querySelector("#emp-stat-fc"),b=e.querySelector("#emp-stat-sec"),y=e.querySelector("#emp-stat-sub"),q=e.querySelector("#emp-pill-count-all"),v=e.querySelector("#emp-pill-count-active"),S=e.querySelector("#emp-pill-count-nonactive");i&&(i.textContent=u.total),r&&(r.textContent=u.active),d&&(d.textContent=u.fc),b&&(b.textContent=u.sec),y&&(y.textContent=`${u.active} Aktif \u2022 ${u.nonactive} Nonaktif`),q&&(q.textContent=u.total),v&&(v.textContent=u.active),S&&(S.textContent=u.nonactive)}async function B({refreshStats:i=!1}={}){if(R=!0,!!!e.querySelector("#emp-list-container"))m();else{let d=e.querySelector("#emp-list-container");d&&(d.innerHTML='<div class="emp-empty-state"><p style="color:#64748B; margin:0;">Memuat data karyawan...</p></div>')}try{let d=new URLSearchParams({page:String(c),limit:String(20),...x()});(i||!T)&&d.set("include_stats","1");let b=await X(`/api/employees?${d.toString()}`);if(b.ok){let y=b.data?.data||[];a=Array.isArray(y)?y:[];let q=b.data?.pagination||{};if(E=Number(q.total??a.length),C=Math.max(1,Number(q.pages??Math.ceil(E/20))||1),c=Math.max(1,Number(q.page??c)),b.data?.stats&&(u=b.data.stats,T=!0),a.length===0&&c>C)return c=C,R=!1,B({refreshStats:i})}else a=[],E=0,C=1,z(b.data?.error||"Gagal memuat data karyawan")}catch(d){console.warn("Gagal load employees:",d),a=[],E=0,C=1,z("Gagal memuat data karyawan: "+d.message)}finally{R=!1,e.querySelector("#emp-list-container")?(_(),$()):m()}}async function P(){if(!ie(N,"export")){z("Anda tidak memiliki izin untuk mengekspor data.");return}try{let[{downloadExcel:i},r]=await Promise.all([Promise.resolve().then(()=>(we(),Oe)),ue("/api/employees/export",x())]),d=r.map((b,y)=>({No:y+1,"Nama Lengkap":b.full_name||"",Cabang:b.branch_name||"",Divisi:b.division||"","No. HP":b.phone||"","Tgl Masuk":b.join_date||"",Status:b.status||"","Target Pindah OS":b.target_pindah_os||"","Target Selesai OS":b.target_selesai_os||"",Catatan:b.notes||""}));i(d,`Data_Karyawan_FCMS_${new Date().toISOString().split("T")[0]}`),ae("Berhasil mengekspor data karyawan ke Excel")}catch(i){z("Gagal ekspor Excel: "+i.message)}}function f(i=null){let r=!!(i&&i.id);if(r&&!M){z("Anda tidak memiliki izin untuk mengedit data karyawan.");return}if(!r&&!w){z("Anda tidak memiliki izin untuk menambah data karyawan.");return}let d=!!(i&&i.id),b=d?tt(i?.row_version):null,y=i?.division||"FACILITY CARE",q=i?.status||"Aktif",v=document.createElement("div");v.className="emp-modal-body",v.innerHTML=`
      <div style="display:flex; flex-direction:column; gap:14px;">
        ${d?`<input type="hidden" id="emp-modal-row-version" value="${b!==null?b:""}">`:""}
        <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#94A3B8; margin-bottom:5px;">
              Nama Lengkap Karyawan <span style="color:#EF4444;">*</span>
            </label>
            <input type="text" id="emp-modal-name" class="form-control" style="width:100%; padding:9px 12px; border:1px solid rgba(255,255,255,0.12); border-radius:8px; background:#111827; color:#F8FAFC;" 
              placeholder="Contoh: Achmad Nuryaini" value="${i?.full_name?String(i.full_name).replace(/"/g,"&quot;"):""}" required />
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#94A3B8; margin-bottom:5px;">
              No. Handphone / WhatsApp
            </label>
            <input type="tel" id="emp-modal-phone" class="form-control" style="width:100%; padding:9px 12px; border:1px solid rgba(255,255,255,0.12); border-radius:8px; background:#111827; color:#F8FAFC;" 
              placeholder="08xx-xxxx-xxxx" value="${i?.phone?String(i.phone).replace(/"/g,"&quot;"):""}" />
          </div>
        </div>

        <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#94A3B8; margin-bottom:5px;">
              Cabang Penempatan <span style="color:#EF4444;">*</span>
            </label>
            <select id="emp-modal-branch" class="form-control" style="width:100%; padding:9px 12px; border:1px solid rgba(255,255,255,0.12); border-radius:8px; background:#172033; color:#F8FAFC;" required>
              <option value="">-- Pilih Cabang --</option>
              ${o.map(A=>`
                <option value="${A.value}" ${String(i?.branch_id)===String(A.value)?"selected":""}>
                  ${A.label}
                </option>
              `).join("")}
            </select>
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#94A3B8; margin-bottom:5px;">
              Divisi / Bagian <span style="color:#EF4444;">*</span>
            </label>
            <select id="emp-modal-division" class="form-control" style="width:100%; padding:9px 12px; border:1px solid rgba(255,255,255,0.12); border-radius:8px; background:#172033; color:#F8FAFC;" required>
              ${kn.map(A=>`
                <option value="${A}" ${y===A?"selected":""}>${A}</option>
              `).join("")}
            </select>
          </div>
        </div>

        <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#94A3B8; margin-bottom:5px;">
              Tanggal Masuk Kerja
            </label>
            <input type="date" id="emp-modal-joindate" class="form-control" style="width:100%; padding:9px 12px; border:1px solid rgba(255,255,255,0.12); border-radius:8px; background:#111827; color:#F8FAFC; color-scheme:dark;" 
              value="${i?.join_date||""}" />
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#94A3B8; margin-bottom:5px;">
              Status Karyawan <span style="color:#EF4444;">*</span>
            </label>
            <select id="emp-modal-status" class="form-control" style="width:100%; padding:9px 12px; border:1px solid rgba(255,255,255,0.12); border-radius:8px; background:#172033; color:#F8FAFC;" required>
              ${lr.map(A=>`
                <option value="${A}" ${q===A?"selected":""}>${A}</option>
              `).join("")}
            </select>
          </div>
        </div>

        <div>
          <label style="display:block; font-size:0.82rem; font-weight:600; color:#94A3B8; margin-bottom:5px;">
            Catatan / Keterangan Karyawan
          </label>
          <textarea id="emp-modal-notes" class="form-control" rows="2" style="width:100%; padding:9px 12px; border:1px solid rgba(255,255,255,0.12); border-radius:8px; background:#111827; color:#F8FAFC;" 
            placeholder="Catatan keahlian, tugas khusus, riwayat cover...">${i?.notes||""}</textarea>
        </div>
      </div>
    `;let S=!1;pe({title:d?"\u270F\uFE0F Edit Data Karyawan":"\u2795 Tambah Data Karyawan",content:v,size:"md",confirmText:d?"Simpan Perubahan":"Simpan Karyawan",confirmClass:"btn-primary",onConfirm:async(A,U)=>{if(S)return;let H=(v.querySelector("#emp-modal-name").value||"").trim(),ee=(v.querySelector("#emp-modal-phone").value||"").trim(),Q=v.querySelector("#emp-modal-branch").value,ne=v.querySelector("#emp-modal-division").value,Z=v.querySelector("#emp-modal-joindate").value||null,G=v.querySelector("#emp-modal-status").value,L=(v.querySelector("#emp-modal-notes").value||"").trim();if(!H||!Q||!ne||!G){z("Harap lengkapi Nama Lengkap, Cabang, Divisi, dan Status!");return}let j=null;if(d&&(j=Jt(v,"#emp-modal-row-version",i?.row_version),j===null)){z("Versi data karyawan (row_version) tidak tersedia. Silakan muat ulang halaman lalu coba lagi.");return}let V={full_name:H,phone:ee,branch_id:Number(Q),division:ne,join_date:Z,status:G,notes:L,...d?{row_version:j}:{}};S=!0;try{let te=d?await X(`/api/employees/${i.id}`,{method:"PUT",body:JSON.stringify(V)}):await X("/api/employees",{method:"POST",body:JSON.stringify(V)});await Vt({res:te,isEdit:d,item:i,close:U,reloadData:async()=>{Ct(),await B({refreshStats:!0})},editSuccessMessage:"Data karyawan berhasil diperbarui!",createSuccessMessage:"Karyawan baru berhasil ditambahkan!",defaultErrorMessage:"Gagal menyimpan data karyawan"})}catch(te){z("Terjadi kesalahan: "+te.message)}finally{S=!1}}})}async function g(i){me({title:"Nonaktifkan Karyawan?",message:`Apakah Anda yakin ingin menonaktifkan personil <strong>${i.full_name}</strong> dari cabang <em>${i.branch_name||"cabang"}</em>?`,confirmText:"Ya, Nonaktifkan",confirmClass:"btn-danger",onConfirm:async()=>{try{let r=await X(`/api/employees/${i.id}`,{method:"DELETE"});r.ok?(ae("Status karyawan berhasil diubah menjadi Tidak Aktif"),Ct(),await B({refreshStats:!0})):z(r.data?.error||"Gagal menonaktifkan data")}catch(r){z("Terjadi kesalahan: "+r.message)}}})}function D(){return a}function O(i){return i.length===0?`
            <div class="emp-empty-state">
              <div style="font-size:2.8rem; margin-bottom:10px;">\u{1F50D}</div>
              <h3 style="font-size:1.15rem; font-weight:700; color:#F8FAFC; margin:0 0 6px 0;">Tidak ada karyawan yang cocok</h3>
              <p style="font-size:0.85rem; color:#94A3B8; margin:0 0 16px 0;">Coba sesuaikan kata kunci pencarian, filter cabang, atau filter divisi.</p>
              <button id="emp-reset-filter" class="btn btn-primary" style="font-size:0.82rem; padding:8px 18px; border-radius:8px; background:#2563EB; color:#fff; border:none; font-weight:700;">Reset Semua Filter</button>
            </div>
      `:K==="table"?`
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
                    ${i.map((r,d)=>{let b=r.division||"FACILITY CARE",y="emp-div-fc",q="";b==="SECURITY"&&(y="emp-div-sec",q="sec");let v=r.status==="Aktif",S=r.phone||"";return`
                        <tr data-id="${r.id}">
                          <td style="text-align:center; color:#94A3B8; font-weight:600;">${(c-1)*20+d+1}</td>
                          <td>
                            <div style="display:flex; align-items:center; gap:10px;">
                              <div class="emp-avatar ${q}" style="width:34px; height:34px; font-size:0.75rem;">
                                ${Sn(r.full_name)}
                              </div>
                              <div style="min-width:0;">
                                <div style="font-weight:700; color:#F8FAFC; font-size:0.9rem;">${r.full_name||"Karyawan"}</div>
                                <div style="font-size:0.74rem; color:#94A3B8;">ID #${r.id} ${r.nik?"\u2022 "+r.nik:""}</div>
                              </div>
                            </div>
                          </td>
                          <td>
                            <span style="font-weight:600; color:#E2E8F0;">${r.branch_name||"Semua Cabang"}</span>
                          </td>
                          <td>
                            <span class="emp-div-badge ${y}">
                              ${b==="SECURITY"?"\u{1F6E1}\uFE0F":"\u{1F9F9}"} ${b}
                            </span>
                          </td>
                          <td>
                            ${S?`
                              <a href="https://wa.me/${S.replace(/[^0-9]/g,"").replace(/^0/,"62")}" target="_blank" rel="noopener" style="color:#34D399; text-decoration:none; font-weight:600; display:inline-flex; align-items:center; gap:4px;">
                                <span>\u{1F4F1}</span> ${S}
                              </a>
                            `:'<span style="color:#64748B;">-</span>'}
                          </td>
                          <td style="font-size:0.84rem; color:#94A3B8;">
                            ${r.join_date?wn(r.join_date):"-"}
                          </td>
                          <td style="font-size:0.84rem; color:#94A3B8;">
                            ${r.notes||"-"}
                          </td>
                          <td style="text-align:center;">
                            <span class="emp-status-badge ${v?"emp-status-aktif":"emp-status-nonaktif"}">
                              ${v?"\u{1F7E2} Aktif":"\u26AA Nonaktif"}
                            </span>
                          </td>
                          <td style="text-align:center;">
                            <div class="emp-action-btns" style="justify-content:center;">
                              ${M?`
                                <button class="emp-icon-btn edit" data-action="edit" data-id="${r.id}" title="Edit Data Karyawan">
                                  <svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>
                                </button>
                              `:""}
                              ${k?`
                                <button class="emp-icon-btn delete" data-action="delete" data-id="${r.id}" title="Nonaktifkan Karyawan">
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
            ${i.map(r=>{let d=r.division||"FACILITY CARE",b="emp-div-fc",y="";d==="SECURITY"&&(b="emp-div-sec",y="sec");let q=r.status==="Aktif",v=r.phone||"";return`
                <div class="emp-card" data-id="${r.id}">
                  <div class="emp-card-top">
                    <span class="emp-div-badge ${b}">
                      ${d==="SECURITY"?"\u{1F6E1}\uFE0F":"\u{1F9F9}"} ${d}
                    </span>
                    <span class="emp-status-badge ${q?"emp-status-aktif":"emp-status-nonaktif"}">
                      ${q?"\u{1F7E2} Aktif":"\u26AA Nonaktif"}
                    </span>
                  </div>

                  <div class="emp-profile-box">
                    <div class="emp-avatar ${y}">
                      ${Sn(r.full_name)}
                    </div>
                    <div style="min-width:0; flex:1;">
                      <div class="emp-name">${r.full_name||"Karyawan"}</div>
                      <div class="emp-subtitle">Personil ${d} \u2022 ID #${r.id}</div>
                    </div>
                  </div>

                  <div class="emp-info-grid">
                    <div class="emp-info-card">
                      <div class="emp-info-label">Cabang Penempatan</div>
                      <div class="emp-info-val" title="${r.branch_name||"-"}">
                        \u{1F3E2} ${r.branch_name||"Semua Cabang"}
                      </div>
                    </div>
                    <div class="emp-info-card">
                      <div class="emp-info-label">Tanggal Masuk</div>
                      <div class="emp-info-val">
                        \u{1F4C5} ${r.join_date?wn(r.join_date):"-"}
                      </div>
                    </div>
                    <div class="emp-info-card">
                      <div class="emp-info-label">Nomor Induk Karyawan</div>
                      <div class="emp-info-val">
                        \u{1F194} ${r.nik||"-"}
                      </div>
                    </div>
                    <div class="emp-info-card">
                      <div class="emp-info-label">Status Personil</div>
                      <div class="emp-info-val" style="color:${q?"#10B981":"#EF4444"};">
                        ${q?"\u2705 Karyawan Aktif":"\u26D4 Nonaktif / Resign"}
                      </div>
                    </div>
                  </div>

                  ${r.notes?`
                    <div class="emp-notes-box">
                      <strong>Catatan:</strong> ${r.notes}
                    </div>
                  `:""}

                  <div class="emp-card-footer">
                    <div style="font-size:0.75rem; color:#94A3B8;">
                      ID #${r.id} ${r.nik?"\u2022 "+r.nik:""}
                    </div>

                    <div class="emp-action-btns">
                      ${v?`
                        <a href="https://wa.me/${v.replace(/[^0-9]/g,"").replace(/^0/,"62")}" target="_blank" rel="noopener" class="emp-icon-btn" title="Chat WhatsApp Karyawan" style="color:#10B981;">
                          <svg width="16" height="16" fill="currentColor" viewBox="0 0 24 24"><path d="M12.031 6.172c-3.181 0-5.767 2.586-5.768 5.766-.001 1.298.38 2.27 1.019 3.287l-.582 2.128 2.182-.573c.978.58 1.911.928 3.145.929 3.178 0 5.767-2.587 5.768-5.766.001-3.187-2.575-5.771-5.764-5.771zm3.392 8.244c-.144.405-.837.774-1.17.824-.299.045-.677.063-1.092-.069-.252-.08-.575-.187-.988-.365-1.739-.751-2.874-2.502-2.961-2.617-.087-.116-.708-.94-.708-1.793s.448-1.273.607-1.446c.159-.173.346-.217.462-.217l.332.006c.106.005.249-.04.39.298.144.347.491 1.2.534 1.287.043.087.072.188.014.304-.058.116-.087.188-.173.289l-.26.304c-.087.086-.177.18-.076.354.101.174.449.741.964 1.201.662.591 1.221.774 1.394.86s.275.072.376-.043c.101-.116.433-.506.549-.68.116-.173.231-.145.39-.086s1.011.477 1.184.564.289.13.332.202c.045.072.045.419-.099.824zm-3.423-14.416c-6.627 0-12 5.373-12 12 0 2.159.57 4.185 1.564 5.939l-1.564 5.717 5.864-1.538c1.691.924 3.63 1.458 5.698 1.458 6.627 0 12-5.373 12-12s-5.373-12-12-12z"/></svg>
                        </a>
                      `:""}

                      ${M?`
                        <button class="emp-icon-btn edit" data-action="edit" data-id="${r.id}" title="Edit Data Karyawan">
                          <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>
                        </button>
                      `:""}
                      ${k?`
                        <button class="emp-icon-btn delete" data-action="delete" data-id="${r.id}" title="Nonaktifkan Karyawan">
                          <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><polyline points="3 6 5 6 21 6"/><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"/><line x1="10" y1="11" x2="10" y2="17"/><line x1="14" y1="11" x2="14" y2="17"/></svg>
                        </button>
                      `:""}
                    </div>
                  </div>
                </div>
              `}).join("")}
          </div>
    `}function W(){let i=e.querySelector("#emp-pagination-container");if(!i)return;i.innerHTML="";let r=_e({page:c,pages:C,total:E,limit:20,onPage:d=>{c=d,B()}});r&&i.appendChild(r)}function J(){let i=e.querySelector("#emp-reset-filter");i&&i.addEventListener("click",()=>{p="",h="all",l="all",s="all",c=1;let r=e.querySelector("#emp-search-input");r&&(r.value="");let d=e.querySelector("#emp-branch-filter");d&&(d.value="all");let b=e.querySelector("#emp-division-filter");b&&(b.value="all"),e.querySelectorAll(".emp-pill").forEach(y=>{y.classList.toggle("active",y.getAttribute("data-status")==="all")}),B()}),e.querySelectorAll(".emp-icon-btn.edit").forEach(r=>{r.addEventListener("click",d=>{let b=Number(d.currentTarget.getAttribute("data-id")),y=a.find(q=>q.id===b);y&&f(y)})}),e.querySelectorAll(".emp-icon-btn.delete").forEach(r=>{r.addEventListener("click",d=>{let b=Number(d.currentTarget.getAttribute("data-id")),y=a.find(q=>q.id===b);y&&g(y)})})}function $(){let i=e.querySelector("#emp-list-container");if(!i)return;let r=D();i.innerHTML=O(r),W(),J()}function m(){let i=D(),r=u.total,d=u.active,b=u.fc,y=u.sec,q=u.total,v=u.active,S=u.nonactive;e.innerHTML=`
      <style>.emp-page{display:flex;flex-direction:column;gap:18px;width:100%;max-width:100%;min-width:0;box-sizing:border-box;overflow-x:hidden}.emp-hero-card{background:#111827;border-radius:16px;border:1px solid rgba(255,255,255,0.08);padding:22px 26px;color:#F8FAFC;box-shadow:0 4px 20px rgba(0,0,0,0.25);position:relative;overflow:hidden;box-sizing:border-box;width:100%;max-width:100%;min-width:0}.emp-hero-card::after{content:'';position:absolute;top:-40px;right:-40px;width:220px;height:220px;background:radial-gradient(circle,rgba(37,99,235,0.12) 0%,transparent 70%);border-radius:50%;pointer-events:none}.emp-stats-row{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:14px;margin-top:18px}.emp-stat-box{background:#172033;border:1px solid rgba(255,255,255,0.08);border-radius:12px;padding:14px 18px;display:flex;align-items:center;gap:14px;transition:transform 0.15s ease,background 0.15s ease,border-color 0.15s ease;min-width:0}.emp-stat-box:hover{background:#1E293B;border-color:rgba(255,255,255,0.16);transform:translateY(-2px)}.emp-stat-icon{width:44px;height:44px;border-radius:10px;display:flex;align-items:center;justify-content:center;font-size:1.3rem;flex-shrink:0;box-shadow:0 2px 8px rgba(0,0,0,0.25)}.emp-stat-val{font-size:1.55rem;font-weight:800;line-height:1.1;letter-spacing:-0.02em;color:#F8FAFC;font-variant-numeric:tabular-nums}.emp-stat-lbl{font-size:0.72rem;text-transform:uppercase;color:#94A3B8;font-weight:700;letter-spacing:0.05em;margin-top:2px}.emp-controls-card{background:#111827;border-radius:14px;border:1px solid rgba(255,255,255,0.08);padding:16px 20px;box-shadow:0 4px 20px rgba(0,0,0,0.25);display:flex;flex-direction:column;gap:14px;box-sizing:border-box;width:100%;max-width:100%;min-width:0}.emp-search-row{display:flex;gap:12px;flex-wrap:wrap;align-items:center}.emp-view-toggle{display:inline-flex;background:#172033;padding:4px;border-radius:10px;border:1px solid rgba(255,255,255,0.08);gap:4px;flex-shrink:0}.emp-view-btn{display:inline-flex;align-items:center;gap:6px;padding:6px 12px;border-radius:7px;border:none;background:transparent;color:#94A3B8;font-size:0.82rem;font-weight:700;cursor:pointer;transition:all 0.15s ease}.emp-view-btn:hover{color:#F8FAFC}.emp-view-btn.active{background:#2563EB;color:#FFFFFF;box-shadow:0 2px 8px rgba(37,99,235,0.35)}.emp-table-card{background:#111827;border-radius:14px;border:1px solid rgba(255,255,255,0.08);box-shadow:0 4px 20px rgba(0,0,0,0.25);overflow:hidden;width:100%;box-sizing:border-box}.emp-table-responsive{width:100%;overflow-x:auto}.emp-table{width:100%;border-collapse:collapse;text-align:left;font-size:0.85rem}.emp-table th{background:#172033;color:#94A3B8;font-weight:700;font-size:0.74rem;text-transform:uppercase;letter-spacing:0.04em;padding:13px 16px;border-bottom:1px solid rgba(255,255,255,0.08);white-space:nowrap}.emp-table td{padding:13px 16px;border-bottom:1px solid rgba(255,255,255,0.05);color:#F8FAFC;vertical-align:middle}.emp-table tbody tr{transition:background 0.15s ease}.emp-table tbody tr:hover{background:rgba(255,255,255,0.03)}.emp-table tbody tr:last-child td{border-bottom:none}.emp-search-wrap{position:relative;flex:1;min-width:240px}.emp-search-icon{position:absolute;left:14px;top:50%;transform:translateY(-50%);color:#64748B;pointer-events:none}.emp-search-input{width:100%;padding:10px 14px 10px 40px;border:1px solid rgba(255,255,255,0.1);border-radius:10px;font-size:0.9rem;color:#F8FAFC;background:#111827;outline:none;transition:border-color 0.15s ease,box-shadow 0.15s ease;box-sizing:border-box}.emp-search-input:focus{border-color:#2563EB;box-shadow:0 0 0 3px rgba(37,99,235,0.2)}.emp-search-input::placeholder{color:#64748B}.emp-select{padding:10px 14px;border:1px solid rgba(255,255,255,0.1);border-radius:10px;font-size:0.9rem;color:#F8FAFC;background:#172033;outline:none;min-width:170px;cursor:pointer;transition:border-color 0.15s ease}.emp-select option{background:#172033;color:#F8FAFC}.emp-select:focus{border-color:#2563EB}.emp-filter-pills{display:flex;align-items:center;gap:8px;overflow-x:auto;white-space:nowrap;scrollbar-width:none;padding-bottom:2px}.emp-filter-pills::-webkit-scrollbar{display:none}.emp-pill{padding:6px 14px;border-radius:20px;font-size:0.82rem;font-weight:600;cursor:pointer;border:1px solid rgba(255,255,255,0.08);background:#172033;color:#94A3B8;transition:all 0.15s ease;display:inline-flex;align-items:center;gap:6px;user-select:none;flex-shrink:0}.emp-pill:hover{background:#1E293B;color:#F8FAFC;border-color:rgba(255,255,255,0.16)}.emp-pill.active{background:#2563EB;color:#FFFFFF;border-color:#2563EB;box-shadow:0 2px 8px rgba(37,99,235,0.35)}.emp-pill-dot{width:8px;height:8px;border-radius:50%}.emp-pill-count{padding:1px 7px;border-radius:10px;font-size:0.72rem;background:rgba(255,255,255,0.08);color:#F8FAFC}.emp-pill.active .emp-pill-count{background:rgba(255,255,255,0.25);color:#FFFFFF}.emp-cards-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(300px,1fr));gap:16px;width:100%;max-width:100%;min-width:0;box-sizing:border-box}.emp-card{background:#172033;border-radius:14px;border:1px solid rgba(255,255,255,0.08);padding:16px 18px;box-shadow:0 4px 16px rgba(0,0,0,0.25);display:flex;flex-direction:column;gap:13px;width:100% !important;max-width:100% !important;min-width:0 !important;box-sizing:border-box !important;overflow:hidden;transition:transform 0.15s ease,box-shadow 0.15s ease,border-color 0.15s ease}.emp-card:hover{transform:translateY(-2px);box-shadow:0 8px 24px rgba(0,0,0,0.35);border-color:rgba(255,255,255,0.18)}.emp-card-top{display:flex;justify-content:space-between;align-items:center;gap:8px;width:100%;min-width:0;box-sizing:border-box;flex-wrap:wrap}.emp-div-badge{display:inline-flex;align-items:center;gap:5px;padding:4px 10px;border-radius:20px;font-size:0.74rem;font-weight:700;flex-shrink:0;white-space:nowrap}.emp-div-fc{background:rgba(37,99,235,0.18);color:#60A5FA;border:1px solid rgba(37,99,235,0.3)}.emp-div-sec{background:rgba(148,163,184,0.15);color:#E2E8F0;border:1px solid rgba(148,163,184,0.25)}.emp-status-badge{display:inline-flex;align-items:center;gap:5px;padding:4px 10px;border-radius:20px;font-size:0.74rem;font-weight:700;flex-shrink:0;white-space:nowrap}.emp-status-aktif{background:rgba(16,185,129,0.18);color:#34D399;border:1px solid rgba(16,185,129,0.3)}.emp-status-nonaktif{background:rgba(239,68,68,0.18);color:#F87171;border:1px solid rgba(239,68,68,0.3)}.emp-profile-box{display:flex;align-items:center;gap:12px;width:100%;min-width:0;box-sizing:border-box}.emp-avatar{width:48px;height:48px;border-radius:12px;background:linear-gradient(135deg,#2563EB 0%,#1D4ED8 100%);color:white;display:flex;align-items:center;justify-content:center;font-weight:800;font-size:1.05rem;flex-shrink:0;box-shadow:0 4px 12px rgba(37,99,235,0.3)}.emp-avatar.sec{background:linear-gradient(135deg,#475569 0%,#1E293B 100%);box-shadow:0 4px 12px rgba(15,23,42,0.3)}.emp-name{font-size:1.02rem;font-weight:800;color:#F8FAFC;line-height:1.25;word-break:break-word;overflow-wrap:anywhere}.emp-subtitle{font-size:0.76rem;color:#94A3B8;font-weight:600;margin-top:3px}.emp-info-grid{display:grid;grid-template-columns:1fr 1fr;gap:8px}.emp-info-card{background:#111827;border:1px solid rgba(255,255,255,0.06);border-radius:10px;padding:10px 12px;display:flex;flex-direction:column;gap:4px;font-size:0.83rem;width:100%;max-width:100%;min-width:0;box-sizing:border-box}.emp-info-label{color:#94A3B8;font-size:0.72rem;font-weight:600;text-transform:uppercase;letter-spacing:0.03em}.emp-info-val{color:#F8FAFC;font-weight:700;min-width:0;word-break:break-word;overflow-wrap:anywhere;font-size:0.82rem}.emp-phone-link{color:#34D399;text-decoration:none;font-weight:600;display:inline-flex;align-items:center;gap:4px}.emp-phone-link:hover{text-decoration:underline}.emp-notes-box{font-size:0.78rem;color:#FCD34D;background:rgba(245,158,11,0.1);border:1px solid rgba(245,158,11,0.25);border-radius:8px;padding:8px 10px;line-height:1.4;word-break:break-word;overflow-wrap:anywhere;width:100%;box-sizing:border-box}.emp-card-footer{display:flex;justify-content:space-between;align-items:center;border-top:1px solid rgba(255,255,255,0.06);padding-top:11px;margin-top:auto;gap:8px;width:100%;min-width:0;box-sizing:border-box;flex-wrap:wrap}.emp-action-btns{display:flex;align-items:center;gap:6px;margin-left:auto;flex-shrink:0}.emp-icon-btn{width:34px;height:34px;border-radius:8px;border:1px solid rgba(255,255,255,0.08);background:#111827;color:#94A3B8;display:flex;align-items:center;justify-content:center;cursor:pointer;transition:all 0.15s ease}.emp-icon-btn:hover{background:#1E293B;color:#F8FAFC;border-color:rgba(255,255,255,0.18)}.emp-icon-btn.delete:hover{background:rgba(239,68,68,0.18);color:#F87171;border-color:rgba(239,68,68,0.3)}.emp-empty-state{grid-column:1 / -1;background:#172033;border:2px dashed rgba(255,255,255,0.12);border-radius:16px;padding:48px 24px;text-align:center}@media (max-width:768px){.emp-page{gap:14px;width:100% !important;max-width:100% !important;overflow-x:hidden !important}.emp-hero-card{padding:16px 14px !important;width:100% !important;max-width:100% !important;box-sizing:border-box !important}.emp-stats-row{grid-template-columns:1fr 1fr;gap:10px}.emp-stat-box{padding:10px 12px}.emp-stat-icon{width:38px;height:38px;font-size:1.1rem}.emp-stat-val{font-size:1.3rem}.emp-hero-actions{width:100% !important;display:flex !important;align-items:center !important;gap:8px !important;flex-wrap:nowrap !important}.emp-btn-add{flex:1 1 auto !important;justify-content:center !important;padding:8px 10px !important;font-size:0.82rem !important;white-space:nowrap !important}.emp-btn-excel{flex:1 1 auto !important;justify-content:center !important;padding:8px 10px !important;font-size:0.82rem !important;white-space:nowrap !important}.emp-btn-refresh{width:38px !important;height:38px !important;padding:0 !important;flex-shrink:0 !important;justify-content:center !important}.emp-controls-card{padding:14px !important;gap:12px;width:100% !important;max-width:100% !important;box-sizing:border-box !important}.emp-search-row{flex-direction:column;gap:10px}.emp-select{width:100%}.emp-view-toggle{width:100%;display:flex}.emp-view-btn{flex:1;justify-content:center}.emp-cards-grid{grid-template-columns:minmax(0,1fr) !important;width:100% !important;max-width:100% !important;min-width:0 !important;box-sizing:border-box !important}.emp-card{padding:14px !important;width:100% !important;max-width:100% !important;min-width:0 !important;box-sizing:border-box !important}.emp-info-grid{grid-template-columns:1fr}.emp-info-card{padding:10px !important;width:100% !important;max-width:100% !important;box-sizing:border-box !important}}</style>

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
              ${w?`
                <button id="emp-btn-add" class="btn btn-primary emp-btn-add" style="background:#2563EB; padding:9px 16px; border-radius:10px; font-weight:700; display:flex; align-items:center; gap:6px; box-shadow:0 4px 12px rgba(37,99,235,0.3); white-space:nowrap;">
                  <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24"><line x1="12" y1="5" x2="12" y2="19"/><line x1="5" y1="12" x2="19" y2="12"/></svg>
                  <span>Tambah Karyawan</span>
                </button>
              `:""}
              
              ${ie(N,"export")?`<button id="emp-btn-excel" class="btn btn-ghost emp-btn-excel" style="background:#172033; border:1px solid rgba(255,255,255,0.12); color:#F8FAFC; padding:9px 14px; border-radius:10px; font-weight:600; display:flex; align-items:center; gap:6px; white-space:nowrap;">
                <span>\u{1F4E5}</span>
                <span>Ekspor Excel</span>
              </button>`:""}

              <button id="emp-btn-refresh" class="btn btn-ghost emp-btn-refresh" style="background:#172033; border:1px solid rgba(255,255,255,0.12); color:#F8FAFC; width:38px; height:38px; padding:0; border-radius:10px; display:flex; align-items:center; justify-content:center; flex-shrink:0;" title="Muat Ulang Data">
                <svg width="16" height="16" fill="none" stroke="currentColor" stroke-width="2.2" viewBox="0 0 24 24" class="${R?"spin-icon":""}"><path d="M23 4v6h-6"/><path d="M1 20v-6h6"/><path d="M3.51 9a9 9 0 0 1 14.85-3.36L23 10M1 14l4.64 4.36A9 9 0 0 0 20.49 15"/></svg>
              </button>
            </div>
          </div>

          <!-- 4 KPI Stat Boxes -->
          <div class="emp-stats-row">
            <div class="emp-stat-box" title="Data karyawan terdaftar di sistem">
              <div class="emp-stat-icon" style="background:linear-gradient(135deg, #3B82F6 0%, #2563EB 100%); color:#FFFFFF;">\u{1F465}</div>
              <div>
                <div class="emp-stat-val" id="emp-stat-total">${r}</div>
                <div class="emp-stat-lbl">Karyawan Terdaftar</div>
                <div id="emp-stat-sub" style="font-size:0.68rem; color:#94A3B8; margin-top:2px;">${d} Aktif \u2022 ${S} Nonaktif</div>
              </div>
            </div>

            <div class="emp-stat-box" title="Total Personil Aktif">
              <div class="emp-stat-icon" style="background:linear-gradient(135deg, #10B981 0%, #059669 100%); color:#FFFFFF;">\u{1F7E2}</div>
              <div>
                <div class="emp-stat-val" id="emp-stat-active">${d}</div>
                <div class="emp-stat-lbl">Personil Aktif</div>
                <div style="font-size:0.68rem; color:#34D399; margin-top:2px;">Status Aktif Operasional</div>
              </div>
            </div>

            <div class="emp-stat-box" title="Divisi Resmi Facility Care">
              <div class="emp-stat-icon" style="background:linear-gradient(135deg, #06B6D4 0%, #0891B2 100%); color:#FFFFFF;">\u{1F3E2}</div>
              <div>
                <div class="emp-stat-val" id="emp-stat-fc">${b}</div>
                <div class="emp-stat-lbl">Facility Care</div>
                <div style="font-size:0.68rem; color:#67E8F9; margin-top:2px;">Aktif Operasional</div>
              </div>
            </div>

            <div class="emp-stat-box" title="Divisi Resmi Security">
              <div class="emp-stat-icon" style="background:linear-gradient(135deg, #8B5CF6 0%, #7C3AED 100%); color:#FFFFFF;">\u{1F6E1}\uFE0F</div>
              <div>
                <div class="emp-stat-val" id="emp-stat-sec">${y}</div>
                <div class="emp-stat-lbl">Security</div>
                <div style="font-size:0.68rem; color:#C4B5FD; margin-top:2px;">Aktif Operasional</div>
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
                placeholder="Cari nama karyawan, no HP, cabang..." value="${p?String(p).replace(/"/g,"&quot;"):""}" />
            </div>

            <select id="emp-branch-filter" class="emp-select">
              <option value="all">Semua Cabang (${o.length})</option>
              ${o.map(Z=>`
                <option value="${Z.value}" ${String(h)===String(Z.value)?"selected":""}>
                  ${Z.label}
                </option>
              `).join("")}
            </select>

            <select id="emp-division-filter" class="emp-select">
              <option value="all">Semua Divisi</option>
              ${kn.map(Z=>`
                <option value="${Z}" ${l===Z?"selected":""}>
                  ${Z}
                </option>
              `).join("")}
            </select>

            <div class="emp-view-toggle" title="Pilih Tampilan">
              <button type="button" class="emp-view-btn ${K==="card"?"active":""}" data-view="card">
                <svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><rect x="3" y="3" width="7" height="7" rx="1"/><rect x="14" y="3" width="7" height="7" rx="1"/><rect x="14" y="14" width="7" height="7" rx="1"/><rect x="3" y="14" width="7" height="7" rx="1"/></svg>
                <span>Kartu</span>
              </button>
              <button type="button" class="emp-view-btn ${K==="table"?"active":""}" data-view="table">
                <svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><line x1="8" y1="6" x2="21" y2="6"/><line x1="8" y1="12" x2="21" y2="12"/><line x1="8" y1="18" x2="21" y2="18"/><line x1="3" y1="6" x2="3.01" y2="6"/><line x1="3" y1="12" x2="3.01" y2="12"/><line x1="3" y1="18" x2="3.01" y2="18"/></svg>
                <span>Tabel</span>
              </button>
            </div>
          </div>

          <!-- Status Horizontal Swipeable Pills -->
          <div class="emp-filter-pills">
            <div class="emp-pill ${s==="all"&&l==="all"?"active":""}" data-status="all" data-div="all">
              <span>Semua Karyawan</span>
              <span class="emp-pill-count" id="emp-pill-count-all">${q}</span>
            </div>
            <div class="emp-pill ${s==="Aktif"?"active":""}" data-status="Aktif" data-div="all">
              <span class="emp-pill-dot" style="background:#10B981;"></span>
              <span>Aktif</span>
              <span class="emp-pill-count" id="emp-pill-count-active">${v}</span>
            </div>
            <div class="emp-pill ${s==="Nonaktif"?"active":""}" data-status="Nonaktif" data-div="all">
              <span class="emp-pill-dot" style="background:#EF4444;"></span>
              <span>Nonaktif / Resign</span>
              <span class="emp-pill-count" id="emp-pill-count-nonactive">${S}</span>
            </div>
          </div>
        </div>

        <!-- Exclusive List Container (Card / Table) -->
        <div id="emp-list-container">
          ${O(i)}
        </div>
        <div id="emp-pagination-container"></div>

      </div>
    `;let A=e.querySelector("#emp-btn-add");A&&A.addEventListener("click",()=>f());let U=e.querySelector("#emp-btn-excel");U&&U.addEventListener("click",()=>P());let H=e.querySelector("#emp-btn-refresh");H&&H.addEventListener("click",()=>B({refreshStats:!0})),e.querySelectorAll(".emp-view-btn").forEach(Z=>{Z.addEventListener("click",()=>{let G=Z.getAttribute("data-view");G&&G!==K&&(K=G,localStorage.setItem("fcms_employees_view_mode",G),e.querySelectorAll(".emp-view-btn").forEach(L=>{L.classList.toggle("active",L.getAttribute("data-view")===G)}),$())})});let ee=e.querySelector("#emp-search-input");ee&&ee.addEventListener("input",Z=>{p=Z.target.value,clearTimeout(F),F=setTimeout(()=>{c=1,B()},350)});let Q=e.querySelector("#emp-branch-filter");Q&&Q.addEventListener("change",Z=>{h=Z.target.value,c=1,B()});let ne=e.querySelector("#emp-division-filter");ne&&ne.addEventListener("change",Z=>{l=Z.target.value,c=1,B()}),e.querySelectorAll(".emp-pill").forEach(Z=>{Z.addEventListener("click",()=>{s=Z.dataset.status;let G=Z.dataset.div;G&&G!=="all"?l=G:Z.dataset.status==="all"&&(l="all");let L=e.querySelector("#emp-division-filter");L&&(L.value=l),e.querySelectorAll(".emp-pill").forEach(j=>{let V=j===Z||s==="all"&&l==="all"&&j.dataset.status==="all";j.classList.toggle("active",V)}),c=1,B()})}),W(),J()}await B({refreshStats:!0})}Fe();we();function Me(e){return String(e??"").replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;").replace(/'/g,"&#39;")}function Se(e){return String(e??"").trim().replace(/\s+/g," ")}function dr(e){if(!e)return null;if(typeof e=="string"){let a=Se(e);return a?{id:null,name:a,branch_name:"",division:"",status:""}:null}let t=Se(e.full_name||e.label||e.name||"");if(!t)return null;let n=Number(e.id??e.value);return{id:Number.isInteger(n)&&n>0?n:null,name:t,branch_name:e.branch_name||e.branch||"",division:e.division||"",status:e.status||""}}function Ye(e=[],t=[]){let n=new Map;for(let a of[...e,...t]){let o=dr(a);if(!o)continue;let s=o.id!==null?`id:${o.id}`:`name:${o.name.toLowerCase()}`;n.has(s)||n.set(s,o)}return Array.from(n.values())}function pr(){if(typeof document>"u"||document.getElementById("fc-emp-combobox-scoped-css"))return;let e=document.createElement("style");e.id="fc-emp-combobox-scoped-css",e.textContent=`
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
      border: 1px solid #29364B !important;
      border-radius: 12px !important;
      font-size: 0.95rem !important;
      font-family: inherit !important;
      font-weight: 500 !important;
      color: #F8FAFC !important;
      background: #1B2940 !important;
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
      background: #1B2940 !important;
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
      background: #29364B;
      color: #F8FAFC;
    }
    .fc-emp-dropdown {
      position: absolute;
      top: calc(100% + 6px);
      left: 0;
      right: 0;
      width: 100%;
      z-index: 1070;
      background: #172033;
      border: 1px solid #29364B;
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
      background: #29364B;
      border-radius: 999px;
    }
    .fc-emp-dropdown::-webkit-scrollbar-thumb:hover {
      background: #64748B;
    }
    .fc-emp-option {
      padding: 10px 14px;
      cursor: pointer;
      border-bottom: 1px solid #29364B;
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
      background: #1B2940;
      color: #F8FAFC;
      border: 1px solid #29364B;
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
      font-size: 0.68rem;
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
      color: #64748B;
      display: inline-flex;
      align-items: center;
      justify-content: center;
    }
    .fc-emp-chip-remove:hover {
      color: #EF4444;
    }
    @media (max-width: 640px) {
      .fc-emp-input {
        font-size: 16px !important;
        height: 48px !important;
        min-height: 48px !important;
      }
    }
  `,document.head.appendChild(e)}function va(e){if(!e)return[];if(Array.isArray(e))return e.map(Se).filter(Boolean);let t=String(e).trim();if(!t)return[];if(t.startsWith("["))try{let n=JSON.parse(t);if(Array.isArray(n))return n.map(Se).filter(Boolean)}catch{}return t.split(",").map(Se).filter(Boolean)}function ze(e,t={}){if(!e)throw new Error("mountEmployeeCombobox requires a valid mount element");pr();let{mode:n="single",inputId:a="",inputName:o="employee_search",placeholder:s="Cari nama karyawan...",required:l=!1,allowManual:h=!0,hasEmployeeIdSchema:p=!0,manualBadgeText:c="Belum terhubung ke Master Karyawan",masterBadgePrefix:E="Terhubung ke Master Karyawan",employees:C=[],initialEmployeeId:T=null,initialEmployeeName:u="",initialParticipants:F=[],maxResults:R=25,onChange:K=null}=t,I=Ye(C),N=[],w=!1,M=-1,k=null,x=0,_=Number(T),B=Number.isInteger(_)&&_>0?_:null,P=Se(u),f=G=>{let L=Se(G).toLowerCase();return L&&Ye(I,N).find(j=>j.name.toLowerCase()===L)||null},g={employee_id:B,employee_name:P,is_manual:!!(P&&B===null&&(p||!f(P))),selected_employee:B!==null?Ye(I).find(G=>G.id===B)||{id:B,name:P}:null,confirmed:!!P},D=n==="multi"?va(F.length?F:P).map(G=>{let L=f(G);return L?{employee_id:L.id,employee_name:L.name,is_manual:!1}:{employee_id:null,employee_name:G,is_manual:!0}}):[];e.innerHTML=`
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
          ${a?`id="${Me(a)}"`:""}
          name="${Me(o)}"
          class="fc-emp-input"
          autocomplete="off"
          placeholder="${Me(s)}"
          ${l&&n==="single"?"required":""}
          value="${n==="single"?Me(P):""}"
          aria-autocomplete="list"
        />
        <button type="button" class="fc-emp-clear-btn" title="Kosongkan" aria-label="Kosongkan pencarian" style="display:none;">&times;</button>
        <div class="fc-emp-dropdown" role="listbox"></div>
      </div>
      <div class="fc-emp-status-row" aria-live="polite"></div>
      ${n==="multi"?'<div class="fc-emp-chips"></div>':""}
    </div>
  `;let O=e.querySelector(".fc-emp-combobox"),W=e.querySelector(".fc-emp-input"),J=e.querySelector(".fc-emp-clear-btn"),$=e.querySelector(".fc-emp-dropdown"),m=e.querySelector(".fc-emp-status-row"),i=e.querySelector(".fc-emp-chips"),r=[];function d(){J.style.display=W.value.trim()?"inline-flex":"none"}function b(){typeof K=="function"&&K(Z.getValue())}function y(){if(n==="multi"){let L=D.filter(te=>te.is_manual).length,j=Se(W.value),V=j&&!f(j);if(D.length===0&&!j){m.innerHTML="";return}L>0||V?m.innerHTML=`<span class="fc-emp-badge fc-emp-badge-manual">\u26A0\uFE0F Ada peserta nama manual (${Me(c)})</span>`:m.innerHTML=`<span class="fc-emp-badge fc-emp-badge-master">\u2713 Semua peserta terhubung ke Master Karyawan (${D.length})</span>`;return}if(w||!g.confirmed){m.innerHTML="";return}let G=Se(W.value);if(!G){m.innerHTML="";return}if(g.employee_id!==null){let L=g.selected_employee?.branch_name?` \u2022 ${Me(g.selected_employee.branch_name)}`:"";m.innerHTML=`<span class="fc-emp-badge fc-emp-badge-master">\u2713 ${Me(E)} (#${g.employee_id})${L}</span>`;return}if(!p){let L=f(G);if(L&&!g.is_manual){m.innerHTML=`<span class="fc-emp-badge fc-emp-badge-master">\u2713 Cocok dengan Master Karyawan${L.id?` (#${L.id})`:""}</span>`;return}}m.innerHTML=`<span class="fc-emp-badge fc-emp-badge-manual">\u26A0\uFE0F ${Me(c)}</span>`}function q(){if(i){if(D.length===0){i.innerHTML="";return}i.innerHTML=D.map((G,L)=>`
      <span class="fc-emp-chip ${G.is_manual?"manual":""}" data-index="${L}">
        <span>${Me(G.employee_name)}</span>
        ${G.is_manual?'<span class="fc-emp-chip-tag">Manual</span>':""}
        <button type="button" class="fc-emp-chip-remove" data-remove-index="${L}" title="Hapus peserta">&times;</button>
      </span>
    `).join("")}}function v(G){let L=Se(G);if(!L)return{items:[],totalMatched:0};let j=L.toLowerCase(),V=Ye(I,N),te=V.filter(se=>se.name.toLowerCase().includes(j)||se.branch_name&&se.branch_name.toLowerCase().includes(j)||se.division&&se.division.toLowerCase().includes(j)),le=[...te.slice(0,R).map(se=>({type:"master",emp:se,label:se.name,meta:[se.branch_name,se.division,se.status].filter(Boolean).join(" \xB7 ")}))],ce=V.some(se=>se.name.toLowerCase()===j);return h&&!ce&&le.push({type:"manual",name:L,label:L,meta:c}),{items:le,totalMatched:te.length}}function S(){$.style.top="",$.style.bottom="";let G=W.getBoundingClientRect(),L=window.innerHeight-G.bottom,j=G.top;L<220&&j>L?($.style.top="auto",$.style.bottom="calc(100% + 6px)"):($.style.top="calc(100% + 6px)",$.style.bottom="auto")}function A(){let G=W.value,L=Se(G);if(!L){$.innerHTML="",H();return}let{items:j,totalMatched:V}=v(G);if(r=j,j.length===0)$.innerHTML=`<div class="fc-emp-empty">Tidak ada karyawan yang cocok dengan "${Me(L)}".</div>`;else{let te=V>R?`<div class="fc-emp-empty" style="border-top:1px solid #F1F5F9; font-size:0.75rem; color:#94A3B8;">Menampilkan ${R} dari ${V} karyawan. Ketik lebih spesifik untuk menyaring.</div>`:"";$.innerHTML=j.map((re,le)=>{let ce=le===M?"active":"";return re.type==="manual"?`
            <div class="fc-emp-option fc-emp-option-manual ${ce}" role="option" data-idx="${le}" data-manual="1" aria-selected="${le===M}">
              <div class="fc-emp-option-main">
                <div class="fc-emp-manual-row">
                  <span class="fc-emp-manual-icon">+</span>
                  <div class="fc-emp-option-name">Gunakan nama manual: <strong>${Me(re.name)}</strong></div>
                </div>
                <div class="fc-emp-manual-badge-row">
                  <span class="fc-emp-badge fc-emp-badge-manual">\u26A0\uFE0F ${Me(re.meta)}</span>
                </div>
              </div>
            </div>
          `:`
          <div class="fc-emp-option ${ce}" role="option" data-idx="${le}" data-emp-id="${re.emp.id??""}" aria-selected="${le===M}">
            <div class="fc-emp-option-main">
              <div class="fc-emp-option-name">${Me(re.label)}</div>
              ${re.meta?`<div class="fc-emp-option-meta">${Me(re.meta)}</div>`:""}
            </div>
          </div>
        `}).join("")+te}S()}function U(){Se(W.value)&&(w||(w=!0,$.classList.add("open"),O.setAttribute("aria-expanded","true")),A())}function H(){w=!1,M=-1,$.classList.remove("open"),O.setAttribute("aria-expanded","false"),Se(W.value)&&n==="single"&&!g.selected_employee&&(g.confirmed=!0),y()}function ee(G){if(G){if(n==="multi"){let L=G.type==="manual"?{employee_id:null,employee_name:G.name,is_manual:!0}:{employee_id:G.emp.id,employee_name:G.emp.name,is_manual:!1};D.some(j=>j.employee_name.toLowerCase()===L.employee_name.toLowerCase())||D.push(L),W.value="",d(),q(),y(),H(),b();return}G.type==="manual"?(g={employee_id:null,employee_name:G.name,is_manual:!0,selected_employee:null,confirmed:!0},W.value=G.name):(g={employee_id:G.emp.id,employee_name:G.emp.name,is_manual:!1,selected_employee:G.emp,confirmed:!0},W.value=G.emp.name),d(),y(),H(),b()}}function Q(G){clearTimeout(k);let L=Se(G);if(!L)return;let j=++x;k=setTimeout(async()=>{try{let V=await hn({search:L,limit:R});if(j!==x)return;Array.isArray(V)&&V.length>0&&(N=Ye(N,V),w&&A())}catch{}},180)}I.length===0&&Ie().then(G=>{Array.isArray(G)&&G.length>0&&(I=Ye(G),g.employee_id!==null&&!g.selected_employee&&(g.selected_employee=I.find(L=>L.id===g.employee_id)||g.selected_employee),y(),w&&A())}).catch(()=>{}),W.addEventListener("focus",()=>{Se(W.value).length>=1&&(M=-1,U())}),W.addEventListener("input",()=>{let G=W.value,L=Se(G);if(n==="single"&&(L?g={employee_id:null,employee_name:L,is_manual:h,selected_employee:null}:g={employee_id:null,employee_name:"",is_manual:!1,selected_employee:null}),M=-1,d(),y(),L.length<1){H(),b();return}U(),Q(L),b()}),W.addEventListener("keydown",G=>{let L=Se(W.value);if(!w&&(G.key==="ArrowDown"||G.key==="ArrowUp")&&L.length>=1&&U(),G.key==="ArrowDown"){if(G.preventDefault(),r.length===0)return;M=(M+1)%r.length,A(),$.querySelector(".fc-emp-option.active")?.scrollIntoView({block:"nearest"})}else if(G.key==="ArrowUp"){if(G.preventDefault(),r.length===0)return;M=M<=0?r.length-1:M-1,A(),$.querySelector(".fc-emp-option.active")?.scrollIntoView({block:"nearest"})}else if(G.key==="Enter"){if(w&&r.length>0){G.preventDefault();let j=M>=0?r[M]:r[0];j&&ee(j)}else if(n==="multi"&&Se(W.value)){G.preventDefault();let j=Se(W.value),V=f(j);ee(V?{type:"master",emp:V}:{type:"manual",name:j})}}else G.key==="Escape"&&w&&(G.preventDefault(),G.stopPropagation(),H())}),$.addEventListener("mousedown",G=>{G.preventDefault()}),$.addEventListener("click",G=>{let L=G.target.closest(".fc-emp-option");if(!L)return;let j=Number(L.dataset.idx);Number.isInteger(j)&&r[j]&&ee(r[j])}),J.addEventListener("click",()=>{W.value="",n==="single"&&(g={employee_id:null,employee_name:"",is_manual:!1,selected_employee:null,confirmed:!1}),d(),y(),H(),W.focus(),b()}),i&&i.addEventListener("click",G=>{let L=G.target.closest("[data-remove-index]");if(!L)return;let j=Number(L.dataset.removeIndex);Number.isInteger(j)&&j>=0&&j<D.length&&(D.splice(j,1),q(),y(),b())});let ne=G=>{if(!e.isConnected){Z.destroy();return}e.contains(G.target)||H()};document.addEventListener("mousedown",ne),d(),q(),y();let Z={inputEl:W,getValue(){if(n==="multi"){let L=Se(W.value),j=[...D];if(L&&!j.some(V=>V.employee_name.toLowerCase()===L.toLowerCase())){let V=f(L);j.push(V?{employee_id:V.id,employee_name:V.name,is_manual:!1}:{employee_id:null,employee_name:L,is_manual:!0})}return{participants:j.map(V=>V.employee_name),participantsString:j.map(V=>V.employee_name).join(", "),items:j,hasManual:j.some(V=>V.is_manual)}}let G=Se(W.value);if(!G)return{employee_id:null,employee_name:"",is_manual:!1,selected_employee:null};if(g.employee_id!==null&&g.employee_name.toLowerCase()===G.toLowerCase())return{employee_id:g.employee_id,employee_name:g.employee_name,is_manual:!1,selected_employee:g.selected_employee};if(!p){let L=f(G);if(L&&!g.is_manual)return{employee_id:L.id,employee_name:L.name,is_manual:!1,selected_employee:L}}return{employee_id:null,employee_name:G,is_manual:!0,selected_employee:null}},setValue({employee_id:G=null,employee_name:L=""}={}){let j=Se(L),V=Number(G),te=Number.isInteger(V)&&V>0?V:null,re=Ye(I,N);g={employee_id:te,employee_name:j,is_manual:!!(j&&te===null),selected_employee:te!==null?re.find(le=>le.id===te)||{id:te,name:j}:null,confirmed:!!j},W.value=j,d(),y()},selectOptionByIdOrName(G){let j=Ye(I,N).find(V=>G!=null&&String(V.id)===String(G)||V.name.toLowerCase()===Se(G).toLowerCase());return j?(ee({type:"master",emp:j}),!0):!1},destroy(){clearTimeout(k),document.removeEventListener("mousedown",ne)}};return Z}function $t(e){if(!e||e==="-")return"-";if(String(e).startsWith("2099"))return"Tetap / PKWTT";try{let t=new Date(e);return isNaN(t.getTime())?e:t.toLocaleDateString("id-ID",{day:"numeric",month:"short",year:"numeric"})}catch{return e}}function At(e){return e?String(e).startsWith("2099"):!1}function Cn(e){if(!e||e==="-")return"\u{1F464}";let t=String(e).trim().split(/\s+/);return t.length===1?t[0].slice(0,2).toUpperCase():(t[0][0]+t[1][0]).toUpperCase()}async function $n(e,t){let a=[],o=[],s=[],l={},h="",p="all",c="all",E="all",C="all",T="",u=1,F=0,R=1,K=!1,I={total:0,active:0,expiring30:0,permanent:0,nonactive:0},N=null,w=!1,M=localStorage.getItem("fcms_contracts_view_mode")||"card",k=t?t.get("dash_filter"):null,x=t?(t.get("month_expiry")||"").trim():"";x&&/^\d{4}-(0[1-9]|1[0-2])$/.test(x)?(T=x,E="Aktif"):k==="expiring30"?(C="30",E="Aktif"):k==="active"&&(E="Aktif");let _=de(),B=ie(_,"contracts","write"),P=ie(_,"contracts","write"),f=ie(_,"contracts","delete");try{let[S,A]=await Promise.all([ve(),Ie()]);o=(Array.isArray(S)?S:[]).filter(H=>String(H.code||"").trim()==="000"||H.branch_scope==="hq"||/^(?:cabang\s+)?000\.\s*management$/i.test(String(H.label||H.full_name||"").trim())),s=Array.isArray(A)?A:[],s.forEach(H=>{let ee=H.value!==void 0?H.value:H.id,Q=H.label||H.full_name;ee&&H.phone&&(l[ee]=H.phone),Q&&H.phone&&(l[Q.toLowerCase()]=H.phone)})}catch(S){console.error("Error fetching cache options for contracts:",S)}function g(){let S={sort:"start_date_desc"};return h.trim()&&(S.search=h.trim()),p&&p!=="all"&&(S.branch_id=p),c&&c!=="all"&&(S.division=c),E&&E!=="all"&&(S.status=E),C&&C!=="all"&&(S.expiring_days=C),T&&(S.month_expiry=T),S}function D(){let S=(A,U)=>{let H=e.querySelector(A);H&&(H.textContent=U)};S("#con-stat-total",I.total),S("#con-stat-active",I.active),S("#con-stat-expiring",I.expiring30),S("#con-stat-permanent",I.permanent),S("#con-pill-count-all",I.total),S("#con-pill-count-aktif",I.active),S("#con-pill-count-expiring",I.expiring30),S("#con-pill-count-tetap",I.permanent),S("#con-pill-count-nonaktif",I.nonactive)}async function O({refreshStats:S=!0}={}){if(w=!0,!!!e.querySelector("#con-list-container"))v();else{let U=e.querySelector("#con-list-container");U&&(U.innerHTML='<div class="con-empty-state"><p style="color:#64748B; margin:0;">Memuat data kontrak...</p></div>')}try{let U=new URLSearchParams({page:String(u),limit:String(20),...g()});(S||!K)&&U.set("include_stats","1");let H=await X(`/api/contracts?${U.toString()}`);if(H.ok){let ee=H.data?.data||[];a=Array.isArray(ee)?ee:[];let Q=H.data?.pagination||{};if(F=Number(Q.total??a.length),R=Math.max(1,Number(Q.pages??Math.ceil(F/20))||1),u=Math.max(1,Number(Q.page??u)),H.data?.stats&&(I=H.data.stats,K=!0),a.length===0&&u>R)return u=R,w=!1,O({refreshStats:S})}else a=[],F=0,R=1,z("Gagal memuat data kontrak dari server.")}catch(U){console.error("Error load contracts:",U),a=[],F=0,R=1,z("Terjadi kesalahan koneksi saat memuat kontrak.")}finally{w=!1,e.querySelector("#con-list-container")?(D(),q()):v()}}function W(){return a}async function J(){pe({title:"\u{1F50D} Analisis Selisih Kontrak Karyawan",content:`
        <div style="text-align:center; padding:24px 10px;">
          <div style="font-size:2rem; margin-bottom:10px;">\u23F3</div>
          <p style="color:#64748B; font-size:0.9rem;">Sedang menyinkronkan data karyawan aktif dengan kontrak aktif di database...</p>
        </div>
      `,cancelText:"Tutup"});try{let[S,A]=await Promise.all([Et("/api/employees/selector",{status:"Aktif"}),Et("/api/contracts",{})]),U=new Set(A.filter(Q=>Q.status==="Aktif").map(Q=>Q.employee_id)),H=S.filter(Q=>!U.has(Q.id)),ee="";H.length===0?ee=`
          <div style="text-align:center; padding:20px 10px;">
            <div style="font-size:2.8rem; margin-bottom:10px;">\u{1F389}</div>
            <h4 style="font-size:1.1rem; font-weight:800; color:#10B981; margin:0 0 8px 0;">Semua Karyawan Aktif Memiliki Kontrak!</h4>
            <p style="font-size:0.86rem; color:#64748B; margin:0;">
              Total <strong>${S.length}</strong> karyawan aktif tercatat memiliki perikatan kontrak aktif yang sah di sistem.
            </p>
          </div>
        `:ee=`
          <div style="display:flex; flex-direction:column; gap:12px;">
            <div style="background:#FEF2F2; border:1px solid #FECACA; border-radius:10px; padding:12px; font-size:0.84rem; color:#991B1B;">
              Ditemukan <strong>${H.length}</strong> dari <strong>${S.length}</strong> karyawan aktif yang <strong>belum memiliki Kontrak Aktif</strong>.
            </div>
            <div style="max-height:360px; overflow-y:auto; display:flex; flex-direction:column; gap:8px; padding-right:4px;">
              ${H.map(Q=>{let ne=A.filter(G=>G.employee_id===Q.id),Z='<span style="color:#D97706; font-weight:600;">Belum pernah ada data kontrak</span>';if(ne.length>0){let G=ne[0];Z=`Riwayat: Status <strong>${G.status}</strong> (Selesai: ${$t(G.end_date)})`}return`
                  <div style="background:#111827; border:1px solid #29364B; border-radius:8px; padding:10px 12px; display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:8px;">
                    <div>
                      <div style="font-weight:700; font-size:0.88rem; color:#0F172A;">${Q.full_name}</div>
                      <div style="font-size:0.76rem; color:#64748B;">
                        Cabang: ${Q.branch_name||"-"} \u2022 Divisi: ${Q.division||"FACILITY CARE"}
                      </div>
                      <div style="font-size:0.75rem; color:#475569; margin-top:2px;">
                        ${Z}
                      </div>
                    </div>
                    ${P?`
                      <button class="btn btn-primary btn-add-contract-for-emp" data-empid="${Q.id}" data-empname="${String(Q.full_name).replace(/"/g,"&quot;")}" data-branchid="${Q.branch_id||""}" data-division="${Q.division||"FACILITY CARE"}" style="font-size:0.76rem; padding:5px 10px; border-radius:6px; background:#2563EB;">
                        + Buat Kontrak
                      </button>
                    `:""}
                  </div>
                `}).join("")}
            </div>
          </div>
        `,pe({title:"Analisis Selisih Kontrak Karyawan",content:ee,cancelText:"Tutup"}),setTimeout(()=>{document.querySelectorAll(".btn-add-contract-for-emp").forEach(Q=>{Q.addEventListener("click",ne=>{let Z=ne.currentTarget,G=Z.getAttribute("data-empid"),L=Z.getAttribute("data-empname"),j=Z.getAttribute("data-branchid"),V=Z.getAttribute("data-division"),te=document.querySelector(".modal-close");te&&te.click(),m({employee_id:G,employee_name:L,branch_id:j,division:V,status:"Aktif"})})})},100)}catch(S){console.error(S),z("Gagal melakukan analisis selisih data kontrak.")}}function $(S,A){if(!S)return"";let U=String(S).split("-").map(Number);if(U.length!==3||!U[0]||!U[1]||!U[2])return"";let H=new Date(U[0],U[1]-1,U[2]);if(isNaN(H.getTime()))return"";if(A==="KONTRAK 6 BULAN")H.setMonth(H.getMonth()+6),H.setDate(H.getDate()-1);else if(A==="KONTRAK 1 TAHUN")H.setFullYear(H.getFullYear()+1),H.setDate(H.getDate()-1);else if(A==="KONTRAK 2 TAHUN")H.setFullYear(H.getFullYear()+2),H.setDate(H.getDate()-1);else return"";let ee=H.getFullYear(),Q=String(H.getMonth()+1).padStart(2,"0"),ne=String(H.getDate()).padStart(2,"0");return`${ee}-${Q}-${ne}`}function m(S=null){let A=!!(S&&S.id);if(A&&!B){z("Anda tidak memiliki izin untuk mengedit kontrak.");return}if(!A&&!P){z("Anda tidak memiliki izin untuk menambah kontrak.");return}let U=!!(S&&S.id),H=U?tt(S?.row_version):null,ee=U?`\u270F\uFE0F Edit Kontrak: ${S.employee_name||"Karyawan"}`:"\u2795 Tambah Kontrak Karyawan Baru",Q=S?.employee_id||"",ne=S?.branch_id||"",Z=S?.division||"FACILITY CARE",G=S?.status||"Aktif",L=S?.contract_type||"KONTRAK 1 TAHUN",j=S?.pkwt_number||"PKWT 1",V=S?.notes||"",te=!1;S&&(S.contract_type==="PKWTT (TETAP)"||S.pkwt_number==="TETAP"?te=!0:S.contract_type&&(S.contract_type.startsWith("KONTRAK")||S.contract_type.includes("BULAN")||S.contract_type.includes("TAHUN"))||S.pkwt_number&&S.pkwt_number.startsWith("PKWT")?te=!1:te=At(S.end_date));let re=S?.start_date?String(S.start_date).slice(0,10):"",le=S?.end_date&&!At(S.end_date)?String(S.end_date).slice(0,10):"";!te&&re&&!le&&(le=$(re,L));let ce=`
      <div style="display:flex; flex-direction:column; gap:14px;">
        ${U?`<input type="hidden" id="con-modal-row-version" value="${H!==null?H:""}">`:""}
        <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Nama Karyawan <span style="color:#DC2626;">*</span>
            </label>
            <div id="con-modal-emp-mount"></div>
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
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
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Divisi / Bagian <span style="color:#DC2626;">*</span>
            </label>
            <select id="con-modal-division" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" required>
              <option value="FACILITY CARE" ${Z==="FACILITY CARE"?"selected":""}>FACILITY CARE</option>
              <option value="SECURITY" ${Z==="SECURITY"?"selected":""}>SECURITY</option>
            </select>
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Status Kontrak <span style="color:#DC2626;">*</span>
            </label>
            <select id="con-modal-status" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" required>
              <option value="Aktif" ${G==="Aktif"?"selected":""}>\u{1F7E2} Aktif</option>
              <option value="Tidak Aktif" ${G==="Tidak Aktif"?"selected":""}>\u26AA Tidak Aktif</option>
              <option value="Resign" ${G==="Resign"?"selected":""}>\u{1F534} Resign</option>
              <option value="Cut" ${G==="Cut"?"selected":""}>\u2702\uFE0F Cut / Pemutusan</option>
            </select>
          </div>
        </div>

        <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Tipe / Durasi Kontrak
            </label>
            <select id="con-modal-type" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;">
              <option value="KONTRAK 6 BULAN" ${L==="KONTRAK 6 BULAN"?"selected":""}>KONTRAK 6 BULAN</option>
              <option value="KONTRAK 1 TAHUN" ${L==="KONTRAK 1 TAHUN"?"selected":""}>KONTRAK 1 TAHUN</option>
              <option value="KONTRAK 2 TAHUN" ${L==="KONTRAK 2 TAHUN"?"selected":""}>KONTRAK 2 TAHUN</option>
              <option value="PKWTT (TETAP)" ${L==="PKWTT (TETAP)"?"selected":""}>PKWTT (TETAP)</option>
            </select>
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Tingkat / No. PKWT
            </label>
            <select id="con-modal-pkwt" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;">
              <option value="PKWT 1" ${j==="PKWT 1"?"selected":""}>PKWT 1</option>
              <option value="PKWT 2" ${j==="PKWT 2"?"selected":""}>PKWT 2</option>
              <option value="PKWT 3" ${j==="PKWT 3"?"selected":""}>PKWT 3</option>
              <option value="PKWT 4" ${j==="PKWT 4"?"selected":""}>PKWT 4</option>
              <option value="PKWT 5" ${j==="PKWT 5"?"selected":""}>PKWT 5</option>
              <option value="PKWT 6" ${j==="PKWT 6"?"selected":""}>PKWT 6</option>
              <option value="TETAP" ${j==="TETAP"?"selected":""}>PKWTT / TETAP</option>
            </select>
          </div>
        </div>

        <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Tanggal Mulai Kontrak
            </label>
            <input type="date" id="con-modal-start" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              value="${re}" />
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Tanggal Selesai Kontrak
            </label>
            <input type="date" id="con-modal-end" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px; ${te?"opacity:0.5; background-color:#1B2940; cursor:not-allowed;":"background-color:#1B2940;"}" 
              value="${le}" ${te?"disabled":""} />
            <div style="margin-top:6px; display:flex; align-items:center; gap:6px;">
              <input type="checkbox" id="con-modal-is-tetap" ${te?"checked":""} style="cursor:pointer;" />
              <label for="con-modal-is-tetap" style="font-size:0.78rem; color:#475569; cursor:pointer; user-select:none;">
                Karyawan Tetap / PKWTT (Tanpa Batas Tanggal)
              </label>
            </div>
          </div>
        </div>

        <div>
          <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
            Catatan Tambahan
          </label>
          <textarea id="con-modal-notes" class="form-control" rows="2" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px; resize:vertical;" 
            placeholder="Keterangan perpanjangan, klausul khusus, atau nomor SK...">${V}</textarea>
        </div>
      </div>
    `,se=document.createElement("div");se.innerHTML=ce;let he=ze(se.querySelector("#con-modal-emp-mount"),{mode:"single",inputId:"con-modal-employee",inputName:"employee_name",placeholder:"Cari nama karyawan...",required:!0,allowManual:!0,hasEmployeeIdSchema:!0,manualBadgeText:"Belum terhubung ke Master Karyawan",employees:s,initialEmployeeId:Q||null,initialEmployeeName:S?.employee_name||"",onChange:oe=>{if(!U&&oe?.employee_id){let Be=s.find(fe=>Number(fe.value??fe.id)===Number(oe.employee_id));if(Be){let fe=se.querySelector("#con-modal-branch"),Te=se.querySelector("#con-modal-division");fe&&!fe.value&&Be.branch_id&&(fe.value=String(Be.branch_id)),Te&&Be.division&&(Te.value=Be.division)}}}}),Ke=!1,$e=pe({title:ee,content:se,confirmText:U?"\u{1F4BE} Simpan Perubahan":"\u2795 Tambah Kontrak",cancelText:"Batal",onConfirm:async(oe,Be)=>{if(Ke)return;let fe=he.getValue(),Te=he.inputEl,ke=oe.querySelector("#con-modal-branch"),Pe=oe.querySelector("#con-modal-division"),je=oe.querySelector("#con-modal-status"),Ae=oe.querySelector("#con-modal-start"),ye=oe.querySelector("#con-modal-end"),Je=oe.querySelector("#con-modal-is-tetap"),Ei=oe.querySelector("#con-modal-type"),Ci=oe.querySelector("#con-modal-pkwt"),$i=oe.querySelector("#con-modal-notes"),Ha=fe.employee_id,Ua=(fe.employee_name||"").trim(),Ga=ke?.value,Ai=Pe?.value||"FACILITY CARE",Fi=je?.value||"Aktif",Zt=Ae?.value?.trim()||null,Ja=Je?.checked,_i=ye?.value?.trim(),ea=Ja?"2099-12-31":_i||null,Bi=Ei?.value,Ti=Ci?.value,Pi=$i?.value?.trim()||null;if(!Ua){z("Harap pilih atau ketik nama karyawan!"),Te?.focus();return}if(!Ga){z("Harap pilih cabang penempatan!"),ke?.focus();return}if(!Ja&&Zt&&ea&&ea<Zt){z("Tanggal selesai kontrak tidak boleh lebih awal dari tanggal mulai kontrak!"),ye?.focus();return}let ta=null;if(U&&(ta=Jt(oe,"#con-modal-row-version",S?.row_version),ta===null)){z("Versi data kontrak (row_version) tidak tersedia. Silakan muat ulang halaman lalu coba lagi.");return}let Li={employee_id:Ha?parseInt(Ha,10):null,employee_name:Ua,branch_id:parseInt(Ga,10),division:Ai,status:Fi,start_date:Zt,end_date:ea,contract_type:Bi,pkwt_number:Ti,notes:Pi,...U?{row_version:ta}:{}};Ke=!0;try{let aa=U?`/api/contracts/${S.id}`:"/api/contracts",Di=await X(aa,{method:U?"PUT":"POST",body:JSON.stringify(Li)});await Vt({res:Di,isEdit:U,item:S,close:Be,reloadData:async()=>{Ct(),await O()},editSuccessMessage:"Data kontrak berhasil diperbarui!",createSuccessMessage:"Kontrak baru berhasil ditambahkan!",defaultErrorMessage:"Gagal menyimpan data kontrak."})}catch(aa){console.error(aa),z("Terjadi kesalahan koneksi ke server.")}finally{Ke=!1}}});setTimeout(()=>{let oe=$e?.overlay||document.querySelector(".modal-overlay:last-child");if(!oe)return;let Be=oe.querySelector("#con-modal-is-tetap"),fe=oe.querySelector("#con-modal-end"),Te=oe.querySelector("#con-modal-start"),ke=oe.querySelector("#con-modal-type"),Pe=oe.querySelector("#con-modal-pkwt");if(!Be||!fe)return;function je(Ae,ye=!0){Ae?(Be.checked=!0,fe.disabled=!0,fe.style.opacity="0.5",fe.style.backgroundColor="#F8FAFC",fe.style.cursor="not-allowed",fe.value="",ye&&(ke&&(ke.value="PKWTT (TETAP)"),Pe&&(Pe.value="TETAP"))):(Be.checked=!1,fe.disabled=!1,fe.style.opacity="1",fe.style.backgroundColor="#FFFFFF",fe.style.cursor="default",ye&&(ke&&ke.value==="PKWTT (TETAP)"&&(ke.value="KONTRAK 1 TAHUN"),Pe&&Pe.value==="TETAP"&&(Pe.value="PKWT 1")))}Be.addEventListener("change",Ae=>{je(Ae.target.checked,!0),Ae.target.checked||fe.focus()}),ke&&ke.addEventListener("change",Ae=>{let ye=Ae.target.value;ye==="PKWTT (TETAP)"?(je(!0,!1),Pe&&(Pe.value="TETAP")):(je(!1,!1),Pe&&Pe.value==="TETAP"&&(Pe.value="PKWT 1"),Te&&Te.value&&(fe.value=$(Te.value,ye)))}),Pe&&Pe.addEventListener("change",Ae=>{Ae.target.value==="TETAP"?(ke&&(ke.value="PKWTT (TETAP)"),je(!0,!1)):(ke&&ke.value==="PKWTT (TETAP)"&&(ke.value="KONTRAK 1 TAHUN"),je(!1,!1),Te&&Te.value&&ke&&(fe.value=$(Te.value,ke.value)))}),Te&&Te.addEventListener("change",Ae=>{!Be.checked&&Ae.target.value&&ke&&(fe.value=$(Ae.target.value,ke.value))})},20)}function i(S){pe({title:"\u{1F5D1}\uFE0F Hapus Data Kontrak",content:`
        <div style="font-size:0.92rem; color:#334155; line-height:1.5;">
          Apakah Anda yakin ingin menghapus catatan kontrak untuk <strong>${S.employee_name||"karyawan ini"}</strong> (ID #${S.id})?
          <div style="margin-top:10px; padding:10px; background:#FEF2F2; border-radius:8px; border:1px solid #FECACA; font-size:0.8rem; color:#991B1B;">
            \u26A0\uFE0F Tindakan ini akan menghapus riwayat masa berlaku kontrak terkait dari sistem.
          </div>
        </div>
      `,confirmText:"Ya, Hapus",cancelText:"Batal",onConfirm:async(A,U)=>{try{let H=await X(`/api/contracts/${S.id}`,{method:"DELETE"});H.ok?(ae("Data kontrak berhasil dihapus."),U(),await O()):z("Gagal menghapus kontrak: "+(H.data?.message||"Terjadi kesalahan"))}catch(H){console.error(H),z("Terjadi kendala koneksi.")}}})}async function r(){if(!ie(_,"export")){z("Anda tidak memiliki izin untuk mengekspor data.");return}try{let S=await ue("/api/contracts/export",g());if(!S.length){z("Belum ada data kontrak untuk diekspor.");return}let A=S.map(U=>{let H=At(U.end_date),ee=!!(U.end_date&&String(U.end_date).trim()&&U.end_date!=="-"),Q=U.days_remaining!==null&&U.days_remaining!==void 0&&U.days_remaining!==""?Number(U.days_remaining):null;return{"Nama Lengkap":U.employee_name||"",Cabang:U.branch_name||"",Divisi:U.division||"","No. PKWT":U.pkwt_number||"","Tipe Kontrak":U.contract_type||"","Tanggal Mulai":U.start_date||"","Tanggal Selesai":H?"Tetap / PKWTT":ee?U.end_date:"-","Sisa Hari":H?"Tetap":ee&&Q!==null&&!isNaN(Q)?Q+" Hari":"-",Status:U.status||"",Catatan:U.notes||""}});ge(A,`Data_Kontrak_Karyawan_${new Date().toISOString().slice(0,10)}`),ae("File Excel Data Kontrak berhasil diunduh!")}catch(S){z("Gagal mengekspor data kontrak: "+S.message)}}function d(S){return S.length===0?`
            <div class="con-empty-state">
              <div style="font-size:2.8rem; margin-bottom:10px;">\u{1F50D}</div>
              <h3 style="font-size:1.15rem; font-weight:700; color:#1E293B; margin:0 0 6px 0;">Tidak ada kontrak yang sesuai kriteria</h3>
              <p style="font-size:0.85rem; color:#64748B; margin:0 0 16px 0;">Coba atur ulang kata kunci pencarian, filter cabang, atau filter masa berlaku kontrak.</p>
              <button id="con-reset-filter" class="btn btn-secondary" style="font-size:0.82rem; padding:7px 16px; border-radius:8px;">Reset Semua Filter</button>
            </div>
      `:M==="table"?`
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
                    ${S.map((A,U)=>{let H=A.division||"FACILITY CARE",ee="con-div-fc",Q="";H==="SECURITY"&&(ee="con-div-sec",Q="sec");let ne=!!(A.end_date&&String(A.end_date).trim()&&A.end_date!=="-"),Z=At(A.end_date),G=A.status==="Aktif",L=A.days_remaining!==null&&A.days_remaining!==void 0&&A.days_remaining!==""?Number(A.days_remaining):null,j=ne&&!Z&&L!==null&&!isNaN(L)?L:null,V="con-status-aktif",te="\u{1F7E2} Aktif";Z?(V="con-status-tetap",te="\u{1F3DB}\uFE0F PKWTT (Tetap)"):G?j!==null?j<=0?(V="con-status-danger",te="\u26A0\uFE0F Habis Hari Ini"):j<=30?(V="con-status-warning",te=`\u26A0\uFE0F Sisa ${j} Hari`):(V="con-status-aktif",te=`\u{1F7E2} Aktif (${j} Hari)`):(V="con-status-aktif",te="\u{1F7E2} Aktif"):(V="con-status-danger",te=`\u26AA ${A.status||"Nonaktif"}`);let re=A.employee_phone||l[A.employee_id]||l[String(A.employee_name||"").toLowerCase()]||"";return`
                        <tr data-id="${A.id}">
                          <td style="text-align:center; color:#94A3B8; font-weight:600;">${(u-1)*20+U+1}</td>
                          <td>
                            <div style="display:flex; align-items:center; gap:10px;">
                              <div class="con-avatar ${Q}" style="width:34px; height:34px; font-size:0.75rem;">
                                ${Cn(A.employee_name)}
                              </div>
                              <div style="min-width:0;">
                                <div style="font-weight:700; color:#1E293B; font-size:0.9rem;">${A.employee_name||"Karyawan Tanpa Nama"}</div>
                                <div style="font-size:0.74rem; color:#64748B;">ID #${A.id} ${A.pkwt_number?"\u2022 "+A.pkwt_number:""}</div>
                              </div>
                            </div>
                          </td>
                          <td>
                            <div style="display:flex; flex-direction:column; gap:3px;">
                              <span style="font-weight:600; color:#334155;">${A.branch_name||"Semua Cabang"}</span>
                              <span class="con-div-badge ${ee}" style="align-self:flex-start;">
                                ${H==="SECURITY"?"\u{1F6E1}\uFE0F":"\u{1F9F9}"} ${H}
                              </span>
                            </div>
                          </td>
                          <td style="text-align:center;">
                            <span class="con-pkwt-badge" style="display:inline-block;">${A.contract_type||"PKWT"}</span>
                          </td>
                          <td style="font-size:0.84rem; color:#475569;">
                            ${$t(A.start_date)}
                          </td>
                          <td style="font-size:0.84rem; color:#475569;">
                            ${Z?'<span style="color:#6366F1; font-weight:600;">PKWTT (Tetap)</span>':$t(A.end_date)}
                          </td>
                          <td>
                            ${Z?'<span style="color:#6366F1; font-weight:600; font-size:0.82rem;">\u{1F3DB}\uFE0F Tidak Berbatas</span>':j!==null?j<=0?'<span style="color:#DC2626; font-weight:700; font-size:0.82rem;">\u26A0\uFE0F Habis Hari Ini</span>':j<=30?`<span style="color:#D97706; font-weight:700; font-size:0.82rem;">\u26A0\uFE0F Sisa ${j} Hari</span>`:`<span style="color:#16A34A; font-weight:600; font-size:0.82rem;">\u{1F7E2} ${j} Hari</span>`:'<span style="color:#94A3B8; font-size:0.82rem;">-</span>'}
                          </td>
                          <td style="text-align:center;">
                            <span class="con-status-badge ${V}">${te}</span>
                          </td>
                          <td style="text-align:center;">
                            <div class="con-action-btns" style="justify-content:center;">
                              ${re?`
                                <a href="https://wa.me/${re.replace(/[^0-9]/g,"").replace(/^0/,"62")}" target="_blank" rel="noopener" class="con-icon-btn" title="Chat WhatsApp Karyawan" style="color:#16A34A;">
                                  <svg width="15" height="15" fill="currentColor" viewBox="0 0 24 24"><path d="M12.031 6.172c-3.181 0-5.767 2.586-5.768 5.766-.001 1.298.38 2.27 1.019 3.287l-.582 2.128 2.182-.573c.978.58 1.911.928 3.145.929 3.178 0 5.767-2.587 5.768-5.766.001-3.187-2.575-5.771-5.764-5.771zm3.392 8.244c-.144.405-.837.774-1.17.824-.299.045-.677.063-1.092-.069-.252-.08-.575-.187-.988-.365-1.739-.751-2.874-2.502-2.961-2.617-.087-.116-.708-.94-.708-1.793s.448-1.273.607-1.446c.159-.173.346-.217.462-.217l.332.006c.106.005.249-.04.39.298.144.347.491 1.2.534 1.287.043.087.072.188.014.304-.058.116-.087.188-.173.289l-.26.304c-.087.086-.177.18-.076.354.101.174.449.741.964 1.201.662.591 1.221.774 1.394.86s.275.072.376-.043c.101-.116.433-.506.549-.68.116-.173.231-.145.39-.086s1.011.477 1.184.564.289.13.332.202c.045.072.045.419-.099.824zm-3.423-14.416c-6.627 0-12 5.373-12 12 0 2.159.57 4.185 1.564 5.939l-1.564 5.717 5.864-1.538c1.691.924 3.63 1.458 5.698 1.458 6.627 0 12-5.373 12-12s-5.373-12-12-12z"/></svg>
                                </a>
                              `:""}
                              ${B?`
                                <button class="con-icon-btn edit" data-action="edit" data-id="${A.id}" title="Edit Data Kontrak">
                                  <svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>
                                </button>
                              `:""}
                              ${f?`
                                <button class="con-icon-btn delete" data-action="delete" data-id="${A.id}" title="Hapus Kontrak">
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
            ${S.map(A=>{let U=A.division||"FACILITY CARE",H="con-div-fc",ee="";U==="SECURITY"&&(H="con-div-sec",ee="sec");let Q=!!(A.end_date&&String(A.end_date).trim()&&A.end_date!=="-"),ne=At(A.end_date),Z=A.status==="Aktif",G=A.days_remaining!==null&&A.days_remaining!==void 0&&A.days_remaining!==""?Number(A.days_remaining):null,L=Q&&!ne&&G!==null&&!isNaN(G)?G:null,j="con-status-aktif",V="\u{1F7E2} Aktif";ne?(j="con-status-tetap",V="\u{1F3DB}\uFE0F PKWTT (Tetap)"):Z?L!==null?L<=0?(j="con-status-danger",V="\u26A0\uFE0F Habis Hari Ini"):L<=30?(j="con-status-warning",V=`\u26A0\uFE0F Sisa ${L} Hari`):(j="con-status-aktif",V=`\u{1F7E2} Aktif (${L} Hari)`):(j="con-status-aktif",V="\u{1F7E2} Aktif"):(j="con-status-danger",V=`\u26AA ${A.status||"Nonaktif"}`);let te=A.employee_phone||l[A.employee_id]||l[String(A.employee_name||"").toLowerCase()]||"";return`
                <div class="con-card" data-id="${A.id}">
                  <div class="con-card-top">
                    <div style="display:flex; align-items:center; gap:6px; flex-wrap:wrap;">
                      <span class="con-div-badge ${H}">
                        ${U==="SECURITY"?"\u{1F6E1}\uFE0F":"\u{1F9F9}"} ${U}
                      </span>
                      ${A.pkwt_number?`
                        <span class="con-pkwt-badge">
                          \u{1F4DC} ${A.pkwt_number}
                        </span>
                      `:""}
                    </div>
                    <span class="con-status-badge ${j}">
                      ${V}
                    </span>
                  </div>

                  <div class="con-profile-box">
                    <div class="con-avatar ${ee}">
                      ${Cn(A.employee_name)}
                    </div>
                    <div style="min-width:0; flex:1;">
                      <div class="con-name">
                        ${A.employee_name||"Karyawan Tanpa Nama"}
                      </div>
                      <div class="con-subtitle">
                        Personil ${U} \u2022 ${A.contract_type||"Kontrak Kerja"}
                      </div>
                    </div>
                  </div>

                  <div class="con-info-grid">
                    <div class="con-info-card">
                      <div class="con-info-label">Cabang Penempatan</div>
                      <div class="con-info-val" title="${A.branch_name||"-"}">
                        \u{1F3E2} ${A.branch_name||"Semua Cabang"}
                      </div>
                    </div>
                    <div class="con-info-card">
                      <div class="con-info-label">Tingkat PKWT</div>
                      <div class="con-info-val">
                        \u{1F4CB} ${A.contract_type||"PKWT 1"}
                      </div>
                    </div>
                    <div class="con-info-card">
                      <div class="con-info-label">Mulai Kontrak</div>
                      <div class="con-info-val">
                        \u{1F4C5} ${$t(A.start_date)}
                      </div>
                    </div>
                    <div class="con-info-card">
                      <div class="con-info-label">Selesai Kontrak</div>
                      <div class="con-info-val">
                        ${ne?"\u{1F3DB}\uFE0F PKWTT (Tetap)":"\u{1F4C5} "+$t(A.end_date)}
                      </div>
                    </div>
                  </div>

                  ${ne?`
                    <div style="background:#EEF2FF; border:1px solid #E0E7FF; border-radius:10px; padding:10px 12px; font-size:0.82rem; color:#4338CA; display:flex; align-items:center; gap:8px;">
                      <span style="font-size:1.1rem;">\u{1F3DB}\uFE0F</span>
                      <div>
                        <strong>Karyawan Tetap (PKWTT)</strong>
                        <div style="font-size:0.75rem; color:#6366F1;">Tidak memiliki batas akhir masa berlaku kontrak.</div>
                      </div>
                    </div>
                  `:L!==null?`
                    <div style="background:${L<=0?"#FEF2F2":L<=30?"#FFFBEB":"#F0FDF4"}; border:1px solid ${L<=0?"#FEE2E2":L<=30?"#FEF3C7":"#DCFCE7"}; border-radius:10px; padding:10px 12px; font-size:0.82rem; color:${L<=0?"#991B1B":L<=30?"#92400E":"#166534"}; display:flex; align-items:center; gap:8px;">
                      <span style="font-size:1.1rem;">${L<=0?"\u26A0\uFE0F":L<=30?"\u23F3":"\u2705"}</span>
                      <div>
                        <strong>${L<=0?"Kontrak Telah Berakhir":`Sisa Waktu Kontrak: ${L} Hari`}</strong>
                        <div style="font-size:0.75rem; opacity:0.85;">
                          ${L<=30&&L>0?"Segera lakukan evaluasi perpanjangan kontrak.":L<=0?"Masa berlaku kontrak telah habis.":"Masa kontrak masih aktif dan aman."}
                        </div>
                      </div>
                    </div>
                  `:""}

                  <div class="con-card-footer">
                    <div style="font-size:0.75rem; color:#94A3B8;">
                      ID #${A.id} \u2022 ${A.pkwt_number||"PKWT"}
                    </div>

                    <div class="con-action-btns">
                      ${te?`
                        <a href="https://wa.me/${te.replace(/[^0-9]/g,"").replace(/^0/,"62")}" target="_blank" rel="noopener" class="con-icon-btn" title="Chat WhatsApp Karyawan" style="color:#16A34A;">
                          <svg width="16" height="16" fill="currentColor" viewBox="0 0 24 24"><path d="M12.031 6.172c-3.181 0-5.767 2.586-5.768 5.766-.001 1.298.38 2.27 1.019 3.287l-.582 2.128 2.182-.573c.978.58 1.911.928 3.145.929 3.178 0 5.767-2.587 5.768-5.766.001-3.187-2.575-5.771-5.764-5.771zm3.392 8.244c-.144.405-.837.774-1.17.824-.299.045-.677.063-1.092-.069-.252-.08-.575-.187-.988-.365-1.739-.751-2.874-2.502-2.961-2.617-.087-.116-.708-.94-.708-1.793s.448-1.273.607-1.446c.159-.173.346-.217.462-.217l.332.006c.106.005.249-.04.39.298.144.347.491 1.2.534 1.287.043.087.072.188.014.304-.058.116-.087.188-.173.289l-.26.304c-.087.086-.177.18-.076.354.101.174.449.741.964 1.201.662.591 1.221.774 1.394.86s.275.072.376-.043c.101-.116.433-.506.549-.68.116-.173.231-.145.39-.086s1.011.477 1.184.564.289.13.332.202c.045.072.045.419-.099.824zm-3.423-14.416c-6.627 0-12 5.373-12 12 0 2.159.57 4.185 1.564 5.939l-1.564 5.717 5.864-1.538c1.691.924 3.63 1.458 5.698 1.458 6.627 0 12-5.373 12-12s-5.373-12-12-12z"/></svg>
                        </a>
                      `:""}

                      ${B?`
                        <button class="con-icon-btn edit" data-action="edit" data-id="${A.id}" title="Edit Data Kontrak">
                          <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>
                        </button>
                      `:""}
                      ${f?`
                        <button class="con-icon-btn delete" data-action="delete" data-id="${A.id}" title="Hapus Kontrak">
                          <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><polyline points="3 6 5 6 21 6"/><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"/><line x1="10" y1="11" x2="10" y2="17"/><line x1="14" y1="11" x2="14" y2="17"/></svg>
                        </button>
                      `:""}
                    </div>
                  </div>
                </div>
              `}).join("")}
          </div>
    `}function b(){let S=e.querySelector("#con-pagination-container");if(!S)return;S.innerHTML="";let A=_e({page:u,pages:R,total:F,limit:20,onPage:U=>{u=U,O({refreshStats:!1})}});A&&S.appendChild(A)}function y(){let S=e.querySelector("#con-reset-filter");S&&S.addEventListener("click",()=>{h="",p="all",c="all",E="all",C="all",u=1;let A=e.querySelector("#con-search-input");A&&(A.value="");let U=e.querySelector("#con-branch-filter");U&&(U.value="all");let H=e.querySelector("#con-division-filter");H&&(H.value="all");let ee=e.querySelector("#con-expiring-filter");ee&&(ee.value="all"),e.querySelectorAll(".con-pill").forEach(Q=>{Q.classList.toggle("active",Q.getAttribute("data-status")==="all")}),O({refreshStats:!1})}),e.querySelectorAll(".con-icon-btn.edit").forEach(A=>{A.addEventListener("click",U=>{let H=parseInt(U.currentTarget.getAttribute("data-id"),10),ee=a.find(Q=>Q.id===H);ee&&m(ee)})}),e.querySelectorAll(".con-icon-btn.delete").forEach(A=>{A.addEventListener("click",U=>{let H=parseInt(U.currentTarget.getAttribute("data-id"),10),ee=a.find(Q=>Q.id===H);ee&&i(ee)})})}function q(){let S=e.querySelector("#con-list-container");if(!S)return;let A=W();S.innerHTML=d(A),b(),y()}function v(){let S=W(),A=I.total,U=I.active,H=I.expiring30,ee=I.permanent,Q=I.nonactive;e.innerHTML=`
      <style>.con-page{display:flex;flex-direction:column;gap:20px;width:100%;max-width:100%;min-width:0;box-sizing:border-box;overflow-x:hidden}.con-hero-card{background:linear-gradient(135deg,#0F172A 0%,#1E293B 100%);border-radius:16px;padding:24px 28px;color:white;box-shadow:0 10px 25px -5px rgba(15,23,42,0.25);position:relative;overflow:hidden;box-sizing:border-box;width:100%;max-width:100%;min-width:0}.con-hero-card::after{content:'';position:absolute;top:-40px;right:-40px;width:220px;height:220px;background:radial-gradient(circle,rgba(59,130,246,0.2) 0%,transparent 70%);border-radius:50%;pointer-events:none}.con-hero-actions{display:flex;align-items:center;gap:8px;flex-wrap:wrap}.con-hero-actions .btn{height:38px;box-sizing:border-box;transition:all 0.2s ease}.con-btn-add:hover{background:#1D4ED8 !important;transform:translateY(-1px)}.con-btn-excel:hover,.con-btn-refresh:hover{background:rgba(255,255,255,0.16) !important;border-color:rgba(255,255,255,0.3) !important;transform:translateY(-1px)}.con-btn-missing:hover{background:rgba(239,68,68,0.25) !important;border-color:rgba(239,68,68,0.6) !important;transform:translateY(-1px)}.con-stats-row{display:grid;grid-template-columns:repeat(4,1fr);gap:14px;margin-top:20px}.con-stat-box{background:rgba(255,255,255,0.05);border:1px solid rgba(255,255,255,0.1);border-radius:12px;padding:14px 18px;display:flex;align-items:center;gap:14px;backdrop-filter:blur(8px);transition:transform 0.15s ease,background 0.15s ease}.con-stat-box:hover{background:rgba(255,255,255,0.08);transform:translateY(-2px)}.con-stat-icon{width:44px;height:44px;border-radius:10px;display:flex;align-items:center;justify-content:center;font-size:1.3rem;flex-shrink:0}.con-stat-val{font-size:1.55rem;font-weight:800;line-height:1.1;letter-spacing:-0.02em}.con-stat-lbl{font-size:0.72rem;text-transform:uppercase;color:#94A3B8;font-weight:700;letter-spacing:0.05em;margin-top:2px}.con-controls-card{background:#172033;border-radius:14px;border:1px solid #29364B;padding:16px 20px;box-shadow:0 2px 8px rgba(0,0,0,0.03);display:flex;flex-direction:column;gap:14px;box-sizing:border-box;width:100%;max-width:100%;min-width:0}.con-search-row{display:flex;gap:12px;align-items:center;flex-wrap:wrap}.con-search-wrap{position:relative;flex:1;min-width:240px}.con-search-icon{position:absolute;left:14px;top:50%;transform:translateY(-50%);color:#94A3B8;pointer-events:none}.con-search-input{width:100%;padding:10px 14px 10px 40px;border:1px solid #CBD5E1;border-radius:10px;font-size:0.9rem;outline:none;transition:border-color 0.15s ease}.con-search-input:focus{border-color:#2563EB;box-shadow:0 0 0 3px rgba(37,99,235,0.1)}.con-select{padding:10px 14px;border:1px solid #CBD5E1;border-radius:10px;font-size:0.9rem;color:#F8FAFC;background:#172033;outline:none;min-width:170px;cursor:pointer}.con-select:focus{border-color:#2563EB}.con-filter-pills{display:flex;align-items:center;gap:8px;overflow-x:auto;white-space:nowrap;scrollbar-width:none;padding-bottom:2px}.con-filter-pills::-webkit-scrollbar{display:none}.con-pill{padding:6px 14px;border-radius:20px;font-size:0.82rem;font-weight:600;cursor:pointer;border:1px solid #29364B;background:#1B2940;color:#64748B;transition:all 0.15s ease;display:inline-flex;align-items:center;gap:6px;user-select:none;flex-shrink:0}.con-pill:hover{background:#111827;color:#F8FAFC}.con-pill.active{background:#0F172A;color:white;border-color:#0F172A}.con-pill-dot{width:8px;height:8px;border-radius:50%}.con-pill-count{padding:1px 6px;border-radius:10px;font-size:0.72rem;background:rgba(0,0,0,0.06)}.con-pill.active .con-pill-count{background:rgba(255,255,255,0.2);color:white}.con-view-toggle{display:inline-flex;align-items:center;background:#111827;padding:3px;border-radius:10px;border:1px solid #CBD5E1;gap:3px;flex-shrink:0}.con-view-btn{display:inline-flex;align-items:center;gap:6px;padding:6px 12px;border-radius:7px;font-size:0.82rem;font-weight:600;border:none;background:transparent;color:#64748B;cursor:pointer;transition:all 0.15s ease;height:34px;box-sizing:border-box;user-select:none}.con-view-btn:hover{color:#0F172A;background:rgba(255,255,255,0.5)}.con-view-btn.active{background:#172033;color:#2563EB;font-weight:700;box-shadow:0 1px 4px rgba(0,0,0,0.08)}.con-table-card{background:#172033;border:1px solid #29364B;border-radius:14px;box-shadow:0 2px 6px rgba(0,0,0,0.03);overflow:hidden}.con-table-responsive{width:100%;overflow-x:auto;-webkit-overflow-scrolling:touch}.con-table{width:100%;border-collapse:collapse;text-align:left;font-size:0.88rem;min-width:960px}.con-table th{background:#1B2940;color:#475569;font-size:0.76rem;font-weight:700;text-transform:uppercase;letter-spacing:0.05em;padding:13px 16px;border-bottom:1.5px solid #E2E8F0;white-space:nowrap}.con-table td{padding:13px 16px;border-bottom:1px solid #F1F5F9;color:#F8FAFC;vertical-align:middle}.con-table tbody tr:hover td{background:#1B2940}.con-table tbody tr:last-child td{border-bottom:none}.con-cards-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(320px,1fr));gap:16px;width:100%;max-width:100%;min-width:0;box-sizing:border-box}.con-card{background:#172033;border-radius:14px;border:1px solid #29364B;padding:16px 18px;box-shadow:0 2px 8px rgba(0,0,0,0.04);display:flex;flex-direction:column;gap:13px;width:100% !important;max-width:100% !important;min-width:0 !important;box-sizing:border-box !important;overflow:hidden;transition:transform 0.15s ease,box-shadow 0.15s ease}.con-card:hover{transform:translateY(-2px);box-shadow:0 6px 16px rgba(0,0,0,0.08);border-color:#CBD5E1}.con-card-top{display:flex;justify-content:space-between;align-items:center;gap:8px;width:100%;min-width:0;box-sizing:border-box;flex-wrap:wrap}.con-div-badge{display:inline-flex;align-items:center;gap:5px;padding:4px 10px;border-radius:20px;font-size:0.74rem;font-weight:700;flex-shrink:0;white-space:nowrap}.con-div-fc{background:#EFF6FF;color:#2563EB;border:1px solid #DBEAFE}.con-div-sec{background:#111827;color:#334155;border:1px solid #29364B}.con-div-rel{background:#F5F3FF;color:#7C3AED;border:1px solid #EDE9FE}.con-pkwt-badge{display:inline-flex;align-items:center;gap:4px;padding:4px 10px;border-radius:20px;font-size:0.74rem;font-weight:700;background:#F5F3FF;color:#6D28D9;border:1px solid #DDD6FE;flex-shrink:0;white-space:nowrap}.con-status-badge{display:inline-flex;align-items:center;gap:5px;padding:4px 10px;border-radius:20px;font-size:0.74rem;font-weight:700;flex-shrink:0;white-space:nowrap}.con-status-aktif{background:#DCFCE7;color:#16A34A;border:1px solid #BBF7D0}.con-status-warning{background:#FEF3C7;color:#B45309;border:1px solid #FDE68A}.con-status-danger{background:#FEE2E2;color:#DC2626;border:1px solid #FCA5A5}.con-status-tetap{background:#E0E7FF;color:#4338CA;border:1px solid #C7D2FE}.con-profile-box{display:flex;align-items:center;gap:12px;width:100%;min-width:0;box-sizing:border-box}.con-avatar{width:48px;height:48px;border-radius:12px;background:linear-gradient(135deg,#2563EB 0%,#1D4ED8 100%);color:white;display:flex;align-items:center;justify-content:center;font-weight:800;font-size:1.05rem;flex-shrink:0;box-shadow:0 4px 10px rgba(37,99,235,0.2)}.con-avatar.sec{background:linear-gradient(135deg,#334155 0%,#0F172A 100%);box-shadow:0 4px 10px rgba(15,23,42,0.2)}.con-avatar.rel{background:linear-gradient(135deg,#7C3AED 0%,#5B21B6 100%);box-shadow:0 4px 10px rgba(124,58,237,0.2)}.con-name{font-size:1.02rem;font-weight:800;color:#0F172A;line-height:1.25;word-break:break-word;overflow-wrap:anywhere}.con-subtitle{font-size:0.76rem;color:#64748B;font-weight:600;margin-top:3px}.con-info-card{background:#1B2940;border:1px solid #F1F5F9;border-radius:10px;padding:11px 12px;display:flex;flex-direction:column;gap:8px;font-size:0.83rem;width:100%;max-width:100%;min-width:0;box-sizing:border-box}.con-info-row{display:flex;align-items:flex-start;justify-content:space-between;gap:8px;width:100%;min-width:0;box-sizing:border-box}.con-info-label{color:#64748B;display:inline-flex;align-items:center;gap:5px;flex-shrink:0;white-space:nowrap;font-size:0.8rem}.con-info-val{color:#0F172A;font-weight:600;text-align:right;min-width:0;flex:1;word-break:break-word;overflow-wrap:anywhere;font-size:0.82rem}.con-notes-box{font-size:0.78rem;color:#475569;background:#FFFBEB;border:1px solid #FEF3C7;border-radius:8px;padding:8px 10px;line-height:1.4;word-break:break-word;overflow-wrap:anywhere;width:100%;box-sizing:border-box}.con-card-footer{display:flex;justify-content:space-between;align-items:center;border-top:1px solid #F1F5F9;padding-top:11px;margin-top:auto;gap:8px;width:100%;min-width:0;box-sizing:border-box;flex-wrap:wrap}.con-action-btns{display:flex;align-items:center;gap:6px;margin-left:auto;flex-shrink:0}.con-icon-btn{width:34px;height:34px;border-radius:8px;border:1px solid #29364B;background:#172033;color:#64748B;display:flex;align-items:center;justify-content:center;cursor:pointer;transition:all 0.15s ease}.con-icon-btn:hover{background:#111827;color:#0F172A;border-color:#CBD5E1}.con-icon-btn.delete:hover{background:#FEE2E2;color:#DC2626;border-color:#FCA5A5}.con-empty-state{grid-column:1 / -1;background:#172033;border:2px dashed #29364B;border-radius:16px;padding:48px 24px;text-align:center}@media (max-width:768px){.con-page{gap:14px;width:100% !important;max-width:100% !important;overflow-x:hidden !important}.con-hero-card{padding:16px 14px !important;width:100% !important;max-width:100% !important;box-sizing:border-box !important}.con-stats-row{grid-template-columns:1fr 1fr;gap:10px}.con-stat-box{padding:10px 12px}.con-stat-icon{width:38px;height:38px;font-size:1.1rem}.con-stat-val{font-size:1.3rem}.con-hero-actions{width:100% !important;display:flex !important;align-items:center !important;gap:8px !important;flex-wrap:wrap !important;overflow-x:visible !important}.con-hero-actions::-webkit-scrollbar{display:none}.con-btn-add{width:100% !important;justify-content:center !important;padding:9px 12px !important;font-size:0.85rem !important;white-space:nowrap !important;height:40px !important}.con-btn-excel{flex:1 1 auto !important;justify-content:center !important;padding:8px 12px !important;font-size:0.85rem !important;white-space:nowrap !important;height:38px !important}.con-btn-refresh{width:38px !important;height:38px !important;padding:0 !important;flex-shrink:0 !important;justify-content:center !important}.con-controls-card{padding:14px !important;gap:12px;width:100% !important;max-width:100% !important;box-sizing:border-box !important}.con-search-row{flex-direction:column;gap:10px}.con-select{width:100%}.con-view-toggle{width:100%;display:flex}.con-view-btn{flex:1;justify-content:center}.con-table-card{border-radius:12px}.con-cards-grid{grid-template-columns:minmax(0,1fr) !important;width:100% !important;max-width:100% !important;min-width:0 !important;box-sizing:border-box !important}.con-card{padding:14px !important;width:100% !important;max-width:100% !important;min-width:0 !important;box-sizing:border-box !important}.con-info-card{padding:10px !important;width:100% !important;max-width:100% !important;box-sizing:border-box !important}}</style>

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
              ${P?`
                <button id="con-btn-add" class="btn btn-primary con-btn-add" style="background:#2563EB; height:38px; padding:0 16px; border-radius:10px; font-weight:700; display:inline-flex; align-items:center; gap:6px; box-shadow:0 4px 12px rgba(37,99,235,0.3); white-space:nowrap;">
                  <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24"><line x1="12" y1="5" x2="12" y2="19"/><line x1="5" y1="12" x2="19" y2="12"/></svg>
                  <span>Tambah Kontrak</span>
                </button>
              `:""}

              ${ie(_,"export")?`
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
                <div class="con-stat-val" id="con-stat-total">${A}</div>
                <div class="con-stat-lbl">Total Kontrak</div>
              </div>
            </div>

            <div class="con-stat-box">
              <div class="con-stat-icon" style="background:rgba(16, 185, 129, 0.2); color:#34D399;">
                \u{1F7E2}
              </div>
              <div>
                <div class="con-stat-val" id="con-stat-active" style="color:#34D399;">${U}</div>
                <div class="con-stat-lbl">Kontrak Aktif</div>
              </div>
            </div>

            <div class="con-stat-box">
              <div class="con-stat-icon" style="background:rgba(245, 158, 11, 0.2); color:#FBBF24;">
                \u26A0\uFE0F
              </div>
              <div>
                <div class="con-stat-val" id="con-stat-expiring" style="color:#FBBF24;">${H}</div>
                <div class="con-stat-lbl">Habis &lt; 30 Hari</div>
              </div>
            </div>

            <div class="con-stat-box">
              <div class="con-stat-icon" style="background:rgba(99, 102, 241, 0.2); color:#A5B4FC;">
                \u{1F3DB}\uFE0F
              </div>
              <div>
                <div class="con-stat-val" id="con-stat-permanent" style="color:#A5B4FC;">${ee}</div>
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
                placeholder="Cari nama karyawan, cabang, no PKWT, catatan..." value="${h?String(h).replace(/"/g,"&quot;"):""}" />
            </div>

            <select id="con-branch-filter" class="con-select">
              <option value="all">Semua Cabang (${o.length})</option>
              ${o.map(re=>`
                <option value="${re.value}" ${String(p)===String(re.value)?"selected":""}>
                  ${re.label}
                </option>
              `).join("")}
            </select>

            <select id="con-division-filter" class="con-select">
              <option value="all">Semua Divisi</option>
              <option value="FACILITY CARE" ${c==="FACILITY CARE"?"selected":""}>FACILITY CARE</option>
              <option value="SECURITY" ${c==="SECURITY"?"selected":""}>SECURITY</option>
            </select>

            <select id="con-expiring-filter" class="con-select">
              <option value="all" ${C==="all"?"selected":""}>Semua Durasi Masa Habis</option>
              <option value="7" ${C==="7"?"selected":""}>Akan Habis \u2264 7 Hari</option>
              <option value="14" ${C==="14"?"selected":""}>Akan Habis \u2264 14 Hari</option>
              <option value="30" ${C==="30"?"selected":""}>Akan Habis \u2264 30 Hari</option>
              <option value="60" ${C==="60"?"selected":""}>Akan Habis \u2264 60 Hari</option>
            </select>

            <div class="con-view-toggle">
              <button type="button" class="con-view-btn ${M==="card"?"active":""}" data-view="card" title="Tampilan Kartu">
                <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><rect x="3" y="3" width="7" height="7" rx="1"/><rect x="14" y="3" width="7" height="7" rx="1"/><rect x="14" y="14" width="7" height="7" rx="1"/><rect x="3" y="14" width="7" height="7" rx="1"/></svg>
                <span>Kartu</span>
              </button>
              <button type="button" class="con-view-btn ${M==="table"?"active":""}" data-view="table" title="Tampilan Tabel">
                <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M3 9h18M3 15h18M9 3v18"/><rect x="3" y="3" width="18" height="18" rx="2"/></svg>
                <span>Tabel</span>
              </button>
            </div>
          </div>

          <!-- Status Horizontal Swipeable Pills -->
          <div class="con-filter-pills">
            <div class="con-pill ${E==="all"&&C==="all"?"active":""}" data-status="all">
              <span>Semua Kontrak</span>
              <span class="con-pill-count" id="con-pill-count-all">${A}</span>
            </div>
            <div class="con-pill ${E==="Aktif"&&C==="all"?"active":""}" data-status="Aktif">
              <span class="con-pill-dot" style="background:#10B981;"></span>
              <span>Aktif</span>
              <span class="con-pill-count" id="con-pill-count-aktif">${U}</span>
            </div>
            <div class="con-pill ${E==="Expiring30"?"active":""}" data-status="Expiring30">
              <span class="con-pill-dot" style="background:#F59E0B;"></span>
              <span>Akan Habis (&lt;30 Hari)</span>
              <span class="con-pill-count" id="con-pill-count-expiring">${H}</span>
            </div>
            <div class="con-pill ${E==="Tetap"?"active":""}" data-status="Tetap">
              <span class="con-pill-dot" style="background:#6366F1;"></span>
              <span>PKWTT (Tetap)</span>
              <span class="con-pill-count" id="con-pill-count-tetap">${ee}</span>
            </div>
            <div class="con-pill ${E==="Nonaktif"?"active":""}" data-status="Nonaktif">
              <span class="con-pill-dot" style="background:#EF4444;"></span>
              <span>Nonaktif / Resign</span>
              <span class="con-pill-count" id="con-pill-count-nonaktif">${Q}</span>
            </div>
          </div>
        </div>

        <!-- Content Area: Pilihan Kartu atau Tabel -->
        <div id="con-list-container">
          ${d(S)}
        </div>
        <div id="con-pagination-container"></div>

      </div>
    `;let ne=e.querySelector("#con-btn-add");ne&&ne.addEventListener("click",()=>m());let Z=e.querySelector("#con-btn-excel");Z&&Z.addEventListener("click",()=>r());let G=e.querySelector("#con-btn-refresh");G&&G.addEventListener("click",()=>O({refreshStats:!0})),e.querySelectorAll(".con-view-btn").forEach(re=>{re.addEventListener("click",()=>{let le=re.getAttribute("data-view");le&&le!==M&&(M=le,localStorage.setItem("fcms_contracts_view_mode",le),e.querySelectorAll(".con-view-btn").forEach(ce=>{ce.classList.toggle("active",ce.getAttribute("data-view")===le)}),q())})});let L=e.querySelector("#con-search-input");L&&L.addEventListener("input",re=>{h=re.target.value,clearTimeout(N),N=setTimeout(()=>{u=1,O({refreshStats:!1})},350)});let j=e.querySelector("#con-branch-filter");j&&j.addEventListener("change",re=>{p=re.target.value,u=1,O({refreshStats:!1})});let V=e.querySelector("#con-division-filter");V&&V.addEventListener("change",re=>{c=re.target.value,u=1,O({refreshStats:!1})});let te=e.querySelector("#con-expiring-filter");te&&te.addEventListener("change",re=>{C=re.target.value,E==="Expiring30"&&C!=="30"&&(E=C==="all"?"all":"Aktif"),e.querySelectorAll(".con-pill").forEach(le=>{let ce=le.getAttribute("data-status"),se=ce==="all"&&E==="all"&&C==="all"||ce==="Aktif"&&E==="Aktif"&&C==="all"||ce==="Expiring30"&&E==="Expiring30"||ce==="Tetap"&&E==="Tetap"||ce==="Nonaktif"&&E==="Nonaktif";le.classList.toggle("active",se)}),u=1,O({refreshStats:!1})}),e.querySelectorAll(".con-pill").forEach(re=>{re.addEventListener("click",le=>{let ce=le.currentTarget.getAttribute("data-status");E=ce,T="",ce==="Expiring30"?C="30":C="all";let se=e.querySelector("#con-expiring-filter");se&&(se.value=C),e.querySelectorAll(".con-pill").forEach(he=>{he.classList.toggle("active",he===le.currentTarget)}),u=1,O({refreshStats:!1})})}),b(),y()}await O({refreshStats:!0})}Ne();Fe();Ne();function ka(e){return e.map(t=>{if(t.type==="hidden")return`<input type="hidden" name="${Y(t.name)}" value="${Y(t.value??"")}">`;if(t.type==="html")return t.html||"";if(t.type==="row")return`<div class="form-row">${ka(t.fields)}</div>`;let n=t.required?"required":"",a=t.label?`<label class="form-label">${Y(t.label)}${t.required?' <span class="required">*</span>':""}</label>`:"",o="";switch(t.type){case"textarea":o=`<textarea name="${Y(t.name)}" class="form-control" placeholder="${Y(t.placeholder||"")}" ${n} rows="${t.rows||3}">${Y(t.value??"")}</textarea>`;break;case"select":let l=(t.options||[]).map(u=>{let F=typeof u=="object"?u.value:u,R=typeof u=="object"?u.label:u,K=t.value==F?"selected":"";return`<option value="${Y(F??"")}" ${K}>${Y(R??"")}</option>`}).join("");o=`<select name="${Y(t.name)}" class="form-control" ${n}><option value="">-- Pilih ${Y(t.label||"")} --</option>${l}</select>`;break;case"combobox":let h=`dl-${Y(t.name)}-${Math.random().toString(36).substring(7)}`,p=(t.options||[]).map(u=>{let F=typeof u=="object"?u.label||u.value||"":u||"";return(F==="undefined"||F==="[object Object]"||F==="null")&&(F=""),F?`<option value="${Y(F)}"></option>`:""}).join(""),c=t.value||"";if(t.value){let u=(t.options||[]).find(F=>(typeof F=="object"?F.value:F)==t.value);if(u){let F=typeof u=="object"?u.label||u.value||"":u||"";F&&F!=="undefined"&&F!=="[object Object]"&&F!=="null"&&(c=F)}}o=`
          <input type="text" name="${Y(t.name)}" list="${h}" class="form-control" value="${Y(c)}" placeholder="Pilih atau ketik baru..." ${n} autocomplete="off">
          <datalist id="${h}">${p}</datalist>
        `;break;case"employee-combobox":{o=`<div class="fc-crud-emp-mount" data-emp-mount="${Y(t.name)}" data-init-name="${Y(t.value||"")}"></div>`;break}case"checkbox":o=`<label class="checkbox-label"><input type="checkbox" name="${Y(t.name)}" value="1" ${t.value?"checked":""}> ${Y(t.checkLabel||t.label)}</label>`;break;case"date":let E=window.parseFlexibleDate&&t.value?window.parseFlexibleDate(t.value):t.value||"";o=`<input type="date" name="${Y(t.name)}" class="form-control" value="${Y(E)}" ${n}>`;break;case"month":let C=t.value?String(t.value).slice(0,7):"";o=`<input type="month" name="${Y(t.name)}" class="form-control" value="${Y(C)}" ${n}>`;break;case"number":let T=t.readonly?"readonly":"";o=`<input type="number" name="${Y(t.name)}" class="form-control" value="${Y(t.value??"")}" placeholder="${Y(t.placeholder||"")}" min="${Y(t.min??"")}" max="${Y(t.max??"")}" step="${Y(t.step||"1")}" ${n} ${T}>`;break;case"email":o=`<input type="email" name="${Y(t.name)}" class="form-control" value="${Y(t.value??"")}" placeholder="${Y(t.placeholder||"")}" ${n} autocomplete="${Y(t.autocomplete||"off")}">`;break;case"password":o=`<input type="password" name="${Y(t.name)}" class="form-control" value="${Y(t.value??"")}" placeholder="${Y(t.placeholder||"")}" ${n} autocomplete="${Y(t.autocomplete||"new-password")}">`;break;case"url":o=`<input type="url" name="${Y(t.name)}" class="form-control" value="${Y(t.value??"")}" placeholder="${Y(t.placeholder||"https://...")}" ${n}>`;break;default:o=`<input type="${Y(t.type||"text")}" name="${Y(t.name)}" class="form-control" value="${Y(t.value??"")}" placeholder="${Y(t.placeholder||"")}" ${n} autocomplete="${Y(t.autocomplete||"off")}">`}let s=t.hint?`<div class="form-hint">${Y(t.hint)}</div>`:"";return`<div class="form-group ${Y(t.class||"")}">${a}${o}${s}</div>`}).join("")}function An(e){let t={},n=new FormData(e);for(let[a,o]of n.entries())t[a]=o===""?null:o;return e.querySelectorAll("input[type=checkbox]").forEach(a=>{a.checked||(t[a.name]=null)}),t}function Fn(e,t){t&&Object.entries(t).forEach(([n,a])=>{let o=e.querySelector(`[name="${n}"]`);o&&(o.hasAttribute("list")||o.classList.contains("fc-emp-input")||(o.type==="checkbox"?o.checked=!!a:o.type==="date"&&a&&window.parseFlexibleDate?o.value=window.parseFlexibleDate(a):o.type==="month"&&a?o.value=String(a).slice(0,7):o.value=a??""))})}we();var Ge={},Ft={on(e,t){Ge[e]||(Ge[e]=new Set),Ge[e].add(t)},off(e,t){Ge[e]&&Ge[e].delete(t)},emit(e,t){Ge[e]&&Ge[e].forEach(n=>{try{n(t)}catch(a){console.warn("[calendarBus] Handler error:",a)}})},clear(){Object.keys(Ge).forEach(e=>delete Ge[e])}},cr=new Set(["schedule","cleaning","cleaning_reports","inspection","inspection_reports","fogging","fogging_reports","reliever","relievers","contract","contracts","issue","issues","training","one_on_one","sp","sp_data","mutasi","basecamp","basecamp_reports","supply"]);function at(e){if(!e){Ft.emit("data:changed",{module:"unknown"});return}let t=String(e).toLowerCase().replace(/^\/api\//,"").replace(/^reports\//,"");Ft.emit("data:changed",{module:t,relevant:cr.has(t)})}var Ea="fcms_import_sessions_v1",Aa="fcms_import_batch_v1",ur=15,mr=45e4,wa=new Set,gt=class extends Error{constructor(t,n={}){super(t||"Gagal mengimport data"),this.name="ImportRequestError",this.status=Number(n.status??0),this.code=n.code||"IMPORT_ERROR",this.importJobId=n.importJobId||null,this.committedChunks=Number(n.committedChunks||0),this.totalChunks=Number(n.totalChunks||0),this.committedRowIndex=Number(n.committedRowIndex??-1),this.recoverable=!!n.recoverable,this.idempotencyKey=n.idempotencyKey||null,this.fingerprint=n.fingerprint||null,this.isNetworkError=!!n.isNetworkError,this.serverStatus=n.serverStatus||null}};function Ce(e){return String(e??"").replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;").replace(/'/g,"&#39;")}function Ca(e){if(e==null)return null;if(Array.isArray(e))return e.map(Ca);if(typeof e=="object"){let t={};for(let n of Object.keys(e).sort())t[n]=Ca(e[n]);return t}return e}function gr(e){let t=2166136261,n=2654435769;for(let a=0;a<e.length;a++){let o=e.charCodeAt(a);t^=o,t=Math.imul(t,16777619)>>>0,n^=o+(a&255),n=Math.imul(n,2246822507)>>>0}return`${t.toString(16).padStart(8,"0")}${n.toString(16).padStart(8,"0")}`}async function Fa(e,t){let n=JSON.stringify({endpoint:String(e||""),body:Ca(t)});try{let a=globalThis?.crypto?.subtle;if(a&&typeof TextEncoder<"u"){let o=new TextEncoder().encode(n),s=await a.digest("SHA-256",o);return Array.from(new Uint8Array(s)).map(l=>l.toString(16).padStart(2,"0")).join("")}}catch{}return`fnv_${gr(n)}_${n.length}`}function br(e){let t=String(e||"import").replace(/[^a-z0-9_-]/gi,"").slice(0,20)||"import";try{if(globalThis?.crypto?.randomUUID)return`imp_${t}_${globalThis.crypto.randomUUID()}`}catch{}return`imp_${t}_${Date.now().toString(36)}_${Math.random().toString(36).slice(2,10)}`}function _n(e,t){try{if(typeof localStorage>"u")return t;let n=localStorage.getItem(e);if(!n)return t;let a=JSON.parse(n);return a&&typeof a=="object"?a:t}catch{return t}}function $a(e,t){try{return typeof localStorage>"u"?!1:(localStorage.setItem(e,JSON.stringify(t)),!0)}catch{return!1}}function Bn(e){let t=Object.values(e||{}).sort((a,o)=>String(o.updatedAt||"").localeCompare(String(a.updatedAt||""))).slice(0,ur),n={};for(let a of t)a?.fingerprint&&(n[a.fingerprint]=a);if(!$a(Ea,n)){for(let a of t)a&&"body"in a&&(a.body=void 0,a.payloadStored=!1);$a(Ea,n)}}function Yt(){return _n(Ea,{})}function _t(e){return e&&Yt()[e]||null}function Wt(e=null){let t=Yt();return Object.values(t).filter(a=>!(!a||a.status==="COMPLETED"||e&&a.module!==e)).sort((a,o)=>String(o.updatedAt||"").localeCompare(String(a.updatedAt||"")))}function Sa(e){if(!e?.fingerprint)return;let t=Yt(),n=e.body!==void 0?JSON.stringify(e.body):"",a=n.length>0&&n.length<=mr;t[e.fingerprint]={fingerprint:e.fingerprint,idempotencyKey:e.idempotencyKey,module:e.module||"import",label:e.label||e.module||"Import",endpoint:e.endpoint,status:e.status,importJobId:e.importJobId||null,committedChunks:Number(e.committedChunks||0),totalChunks:Number(e.totalChunks||0),committedRowIndex:Number(e.committedRowIndex??-1),recoverable:!!e.recoverable,errorCode:e.errorCode||null,errorMessage:e.errorMessage||null,idempotentReplay:!!e.idempotentReplay,result:e.result||null,payloadStored:e.status==="COMPLETED"?!1:a,body:e.status==="COMPLETED"?void 0:a?e.body:void 0,updatedAt:new Date().toISOString()},Bn(t)}function bt(e){if(!e)return;let t=Yt();t[e]&&(delete t[e],Bn(t))}function Bt(e){return!e||typeof e!="object"?!1:$a(Aa,{...e,updatedAt:new Date().toISOString()})}function Tn(){return _n(Aa,null)}function _a(){try{typeof localStorage<"u"&&localStorage.removeItem(Aa)}catch{}}function hr(e){return e&&typeof e=="object"&&!Array.isArray(e)&&e.data&&typeof e.data=="object"?e.data:e&&typeof e=="object"?e:{}}async function Le({module:e,label:t=null,endpoint:n,body:a,idempotencyKey:o=null,recover:s=!1,forceNewKey:l=!1,configOverride:h=null,fetchImpl:p=null}){let c=await Fa(n,a);if(wa.has(c)){let u=_t(c);throw new gt("Proses import untuk data ini sedang berjalan. Mohon tunggu hingga selesai.",{status:409,code:"CLIENT_IMPORT_IN_PROGRESS",importJobId:u?.importJobId||null,committedChunks:u?.committedChunks||0,totalChunks:u?.totalChunks||0,recoverable:!1,idempotencyKey:u?.idempotencyKey||o||null,fingerprint:c})}let E=_t(c),C=o?String(o).trim():null;!C&&!l&&E&&E.status!=="COMPLETED"&&E.idempotencyKey&&(C=E.idempotencyKey),C||(C=br(e)),wa.add(c),Sa({fingerprint:c,idempotencyKey:C,module:e,label:t||e,endpoint:n,status:"IN_PROGRESS",importJobId:E?.importJobId||null,committedChunks:E?.committedChunks||0,totalChunks:E?.totalChunks||0,committedRowIndex:E?.committedRowIndex??-1,recoverable:!0,body:a});let T={"Idempotency-Key":C,...s?{"X-Import-Recover":"true"}:{}};try{let u=await X(n,{method:"POST",headers:T,body:JSON.stringify(a),...h?{configOverride:h}:{},...p?{fetchImpl:p}:{}});if(u.ok){let P=hr(u.data),f=P.importJobId||E?.importJobId||null,g=!!P.idempotentReplay;return Sa({fingerprint:c,idempotencyKey:C,module:e,label:t||e,endpoint:n,status:"COMPLETED",importJobId:f,committedChunks:Number(E?.totalChunks||E?.committedChunks||1),totalChunks:Number(E?.totalChunks||E?.committedChunks||1),committedRowIndex:E?.committedRowIndex??-1,recoverable:!1,idempotentReplay:g,result:P}),{ok:!0,status:u.status,data:P,raw:u.data,idempotencyKey:C,fingerprint:c,importJobId:f,idempotentReplay:g}}let F=u.data&&typeof u.data=="object"?u.data:{},R=Number(u.status||0),K=F.code||(R===0?"NETWORK_ERROR":R===409?"IMPORT_CONFLICT":"IMPORT_FAILED"),I=Number(F.committedChunks??E?.committedChunks??0),N=Number(F.totalChunks??E?.totalChunks??0),w=Number(F.committedRowIndex??E?.committedRowIndex??-1),M=F.importJobId||E?.importJobId||null,k=R===0,x=k||R===409&&K==="IMPORT_IN_PROGRESS"||!!F.recoverable||R>=500,_="FAILED";k?_="NETWORK_ERROR":R===409&&K==="IMPORT_IN_PROGRESS"?_="CONFLICT_IN_PROGRESS":(K==="IMPORT_PARTIAL_FAILED"||R>=500)&&(_="PARTIAL_FAILED");let B=F.error||`Gagal mengimport data (HTTP ${R})`;throw k?B=F.error||"Koneksi terputus saat mengimport data. Anda dapat mencoba ulang secara aman dengan key yang sama.":R===409&&K==="IMPORT_IN_PROGRESS"?B=F.error||"Import identik sedang berlangsung atau terkunci (HTTP 409).":R===409&&K==="IDEMPOTENCY_KEY_MISMATCH"?B=F.error||"Idempotency-Key bertabrakan dengan payload berbeda (HTTP 409).":K==="IMPORT_PARTIAL_FAILED"&&(B=`${F.error||"Import gagal di tengah chunk"} (${I}/${N||"?"} chunk tersimpan).`),Sa({fingerprint:c,idempotencyKey:C,module:e,label:t||e,endpoint:n,status:_,importJobId:M,committedChunks:I,totalChunks:N,committedRowIndex:w,recoverable:x,errorCode:K,errorMessage:B,body:a}),new gt(B,{status:R,code:K,importJobId:M,committedChunks:I,totalChunks:N,committedRowIndex:w,recoverable:x,idempotencyKey:C,fingerprint:c,isNetworkError:k,serverStatus:F.status||null})}finally{wa.delete(c)}}async function Tt(e,{recover:t=!1,configOverride:n=null,fetchImpl:a=null}={}){let o=_t(e);if(!o)throw new gt("Sesi import tidak ditemukan atau sudah dihapus.",{status:404,code:"IMPORT_SESSION_NOT_FOUND",recoverable:!1,fingerprint:e});if(!o.payloadStored||o.body===void 0)throw new gt("Payload terlalu besar untuk disimpan di browser. Silakan pilih ulang file Excel yang sama \u2014 Idempotency-Key akan tetap dipertahankan.",{status:400,code:"IMPORT_PAYLOAD_REUPLOAD_REQUIRED",recoverable:!1,idempotencyKey:o.idempotencyKey,fingerprint:e});return await Le({module:o.module,label:o.label,endpoint:o.endpoint,body:o.body,idempotencyKey:o.idempotencyKey,recover:t,configOverride:n,fetchImpl:a})}function nt(e){if(!e)return"";let{variant:t="info",title:n="",message:a="",idempotencyKey:o="",importJobId:s="",committedChunks:l=0,totalChunks:h=0,errorCode:p="",fingerprint:c="",canRetry:E=!1,canRecover:C=!1,canReplay:T=!1,canDismiss:u=!0}=e,F={progress:{bg:"#EFF6FF",border:"#93C5FD",color:"#1E3A8A"},success:{bg:"#ECFDF5",border:"#6EE7B7",color:"#065F46"},warning:{bg:"#FFFBEB",border:"#FCD34D",color:"#92400E"},recover:{bg:"#FFF7ED",border:"#FDBA74",color:"#9A3412"},error:{bg:"#FEF2F2",border:"#FCA5A5",color:"#991B1B"},info:{bg:"#F8FAFC",border:"#CBD5E1",color:"#334155"}}[t]||{bg:"#F8FAFC",border:"#CBD5E1",color:"#334155"},R=h>0?`<span class="badge badge-info" style="margin-right:6px;">Chunk ${l}/${h}</span>`:"",K=p?`<span class="badge badge-danger" style="margin-right:6px;">${Ce(p)}</span>`:"",I=o?`<code title="${Ce(o)}" style="font-size:0.75rem;background:rgba(255,255,255,0.7);padding:2px 6px;border-radius:4px;">Key: ${Ce(o.slice(0,26))}${o.length>26?"\u2026":""}</code>`:"",N=s?`<code title="${Ce(s)}" style="font-size:0.75rem;background:rgba(255,255,255,0.7);padding:2px 6px;border-radius:4px;margin-left:6px;">Job: ${Ce(s.slice(0,12))}\u2026</code>`:"";return`
    <div class="import-status-banner" data-variant="${Ce(t)}" data-fingerprint="${Ce(c)}" style="background:${F.bg};border:1px solid ${F.border};color:${F.color};border-radius:10px;padding:12px 16px;">
      <div style="display:flex;justify-content:space-between;align-items:flex-start;gap:12px;flex-wrap:wrap;">
        <div style="flex:1;min-width:240px;">
          <div style="font-weight:700;font-size:0.95rem;margin-bottom:4px;">${Ce(n)}</div>
          <div style="font-size:0.88rem;line-height:1.45;">${Ce(a)}</div>
          <div style="margin-top:8px;display:flex;flex-wrap:wrap;align-items:center;gap:4px;">
            ${K}${R}${I}${N}
          </div>
        </div>
        <div style="display:flex;gap:8px;align-items:center;flex-wrap:wrap;">
          ${E?`<button type="button" class="btn btn-primary btn-sm" data-import-action="retry" data-fingerprint="${Ce(c)}">\u{1F504} Coba Lagi / Lanjutkan</button>`:""}
          ${C?`<button type="button" class="btn btn-warning btn-sm" data-import-action="recover" data-fingerprint="${Ce(c)}">\u{1F6E0}\uFE0F Pulihkan Lock (Recover)</button>`:""}
          ${T?`<button type="button" class="btn btn-secondary btn-sm" data-import-action="replay" data-fingerprint="${Ce(c)}">\u267B\uFE0F Replay Key Sama</button>`:""}
          ${u?`<button type="button" class="btn btn-ghost btn-sm" data-import-action="dismiss" data-fingerprint="${Ce(c)}">\u2715 Tutup</button>`:""}
        </div>
      </div>
    </div>
  `}function De({container:e,title:t,icon:n,apiPath:a,columns:o,formFields:s,filterFields:l,defaultFilters:h={},defaultSort:p="",itemLabel:c="Data",canCreate:E=!0,canEdit:C=!0,canDelete:T=!0,onBeforeSubmit:u,onAfterLoad:F,onDataLoaded:R,extraActions:K=[],initialSearch:I="",exportOptions:N=null,bulkDelete:w=!1}){let M=de(),k=(a||"").replace("/api/","").split("/")[0].replace(/-/g,"_"),x=ie(M,k,"create"),_=ie(M,k,"update"),B=ie(M,k,"delete"),P=ie(M,k,"export")||ie(M,"export"),f=ie(M,k,"import");if(E=E&&x,C=C&&_,T=T&&B,w=w&&B,!P&&!f)N=null;else if(N){f||(N.import=!1,N.template=!1),P||(N.export=!1);let v=N.export!==!1&&typeof N.onExport=="function",S=N.template!==!1&&typeof N.onTemplate=="function",A=N.import!==!1&&typeof N.onImport=="function";!v&&!S&&!A&&(N=null)}!C&&Array.isArray(K)&&(K=K.filter(v=>v&&v.readOnly===!0));let g=1,D={...h};I&&(D.search=I);let O=new Set;e.innerHTML=`
    ${w?`
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
      <h1 class="page-title">${n} ${Y(t)}</h1>
      <div class="page-actions" style="display:flex; gap:8px; align-items:center;">
        ${E?`<button class="btn btn-primary" id="btn-create">+ Tambah ${Y(c)}</button>`:""}
        ${N?`
          <div class="aksi-dropdown-container" style="position:relative; display:inline-block;">
            <button class="btn btn-ghost" id="btn-aksi-main" style="background:#172033; border:1px solid #29364B; padding:8px 16px; border-radius:8px; font-weight:600; color:#F8FAFC; display:flex; align-items:center; gap:6px; cursor:pointer;">
              \u22EE Aksi
            </button>
            <div id="aksi-menu-main" class="aksi-menu-content" style="display:none; position:absolute; top:calc(100% + 4px); right:0; background:#172033; border:1px solid #29364B; border-radius:8px; box-shadow:0 8px 24px rgba(0,0,0,0.5); flex-direction:column; min-width:200px; z-index:999; padding:8px 0;">
              ${N.export!==!1&&typeof N.onExport=="function"?`
              <button type="button" class="dropdown-item" id="btn-export-${N.moduleName||"default"}" style="width:100%; text-align:left; padding:10px 16px; background:none; border:none; cursor:pointer; font-size:0.9rem; color:#F8FAFC; display:flex; gap:8px; align-items:center;">
                \u{1F4E5} Export Excel
              </button>`:""}
              ${N.template!==!1&&typeof N.onTemplate=="function"?`
              <button type="button" class="dropdown-item" id="btn-template-${N.moduleName||"default"}" style="width:100%; text-align:left; padding:10px 16px; background:none; border:none; cursor:pointer; font-size:0.9rem; color:#F8FAFC; display:flex; gap:8px; align-items:center;">
                \u{1F4C4} Download Template
              </button>`:""}
              ${N.import!==!1&&typeof N.onImport=="function"?`
              <label class="dropdown-item" style="display:flex; width:100%; text-align:left; padding:10px 16px; background:none; border:none; cursor:pointer; font-size:0.9rem; color:#F8FAFC; margin:0; gap:8px; align-items:center;" id="label-import-${N.moduleName||"default"}">
                <span class="import-text">\u{1F4E4} Import Excel</span>
                <input type="file" id="input-import-${N.moduleName||"default"}" accept=".xlsx, .xls, .csv" style="display:none;">
              </label>`:""}
            </div>
          </div>
          <style>.show-aksi-menu{display:flex !important}.dropdown-item:hover{background-color:#22314A !important;color:#38BDF8 !important}</style>
        `:""}
      </div>
    </div>
    

    ${l&&l.length>0?`
    <div class="filter-bar" style="background: #111827; border-radius: 12px; padding: 10px 16px; display: flex; align-items: center; gap: 8px; margin-bottom: 24px; border: 1px solid #29364B; box-shadow: 0 4px 12px rgba(0,0,0,0.25);">
        ${l.filter(v=>v.type==="search").map(v=>`<div class="filter-search-wrap" style="flex:1; min-width:0;"><input type="search" class="filter-search" placeholder="${Y(v.placeholder||"Cari...")}" id="filter-search" value="${Y(D.search||"")}" style="width:100%; background: #1B2940; border: 1px solid #29364B; color: #F8FAFC; border-radius: 8px; padding: 8px 12px; font-size: 0.85rem; outline:none;"></div>`).join("")}
        
        <div class="filter-dropdowns-desktop">
          ${l.filter(v=>v.type!=="search").map(v=>{if(v.type==="select"||v.type==="combobox"){let S=(v.label||"").startsWith("Pilih")?v.label:`Pilih ${v.label||""}`;return`<select class="filter-select" name="${Y(v.name)}" id="filter-${Y(v.name)}" style="background: #172033; border: 1px solid #29364B; border-radius: 8px; padding: 7px 10px; font-size: 0.85rem; color: #F8FAFC; cursor: pointer; outline:none;"><option value="">${Y(S)}</option>${(v.options||[]).map(A=>`<option value="${Y(typeof A=="object"?A.value:A)}" ${D[v.name]===(typeof A=="object"?A.value:A)?"selected":""}>${Y(typeof A=="object"?A.label:A)}</option>`).join("")}</select>`}return""}).join("")}
          <button id="btn-reset-filter" style="background: transparent; border: none; color: #38BDF8; font-weight: 600; font-size: 0.85rem; cursor: pointer; padding: 7px 8px; white-space:nowrap;">Reset</button>
        </div>
        
        <button id="btn-mobile-filter" class="btn-mobile-filter-trigger">\u2699 Filter</button>
        
        <div class="filter-options-wrapper" id="filter-options-wrapper">
          <div class="bottom-sheet-header">
            <h3 style="margin:0; font-size:1rem; color:#F8FAFC;">Filter Data</h3>
            <button class="btn-close-sheet" id="btn-close-filter-sheet" style="background:none;border:none;font-size:1.5rem;cursor:pointer;line-height:1;color:#94A3B8;">&times;</button>
          </div>
          ${l.filter(v=>v.type!=="search").map(v=>{if(v.type==="select"||v.type==="combobox"){let S=(v.label||"").startsWith("Pilih")?v.label:`Pilih ${v.label||""}`;return`<select class="filter-select filter-select-sheet" name="${Y(v.name)}-sheet" id="filter-sheet-${Y(v.name)}" style="width:100%; background: #172033; border: 1px solid #29364B; border-radius: 10px; padding: 12px 14px; font-size: 0.9rem; color: #F8FAFC; cursor: pointer; outline:none;"><option value="">${Y(S)}</option>${(v.options||[]).map(A=>`<option value="${Y(typeof A=="object"?A.value:A)}" ${D[v.name]===(typeof A=="object"?A.value:A)?"selected":""}>${Y(typeof A=="object"?A.label:A)}</option>`).join("")}</select>`}return""}).join("")}
          <button id="btn-reset-filter-sheet" style="background: transparent; border: none; color: #38BDF8; font-weight: 600; font-size: 0.9rem; cursor: pointer; padding: 8px;">Reset</button>
        </div>
    </div>`:""}

    <div id="crud-import-banner" style="display:none;margin-bottom:16px;"></div>

    <div class="card">
      <div class="card-body p-0" id="table-container">
        <div class="loading-spinner"><div class="spinner"></div></div>
      </div>
      <div class="card-footer" id="pagination-container"></div>
    </div>
  `,document.getElementById("btn-aksi-main")?.addEventListener("click",v=>{v.stopPropagation(),document.getElementById("aksi-menu-main")?.classList.toggle("show-aksi-menu")});function W(){let v=document.getElementById("bulk-toolbar");if(!v)return;let S=document.getElementById("bulk-count"),A=document.getElementById("btn-bulk-delete"),U=document.getElementById("btn-bulk-cancel");S.textContent=`${O.size} item dipilih`,O.size>0?(v.style.display="flex",A.disabled=!1,U.disabled=!1):(v.style.display="none",A.disabled=!0,U.disabled=!0);let H=document.getElementById("select-all-checkbox");if(H){let ee=document.querySelectorAll(".row-checkbox");if(ee.length>0){let Q=[...ee].every(Z=>Z.checked),ne=[...ee].some(Z=>Z.checked);H.checked=Q,H.indeterminate=ne&&!Q}else H.checked=!1,H.indeterminate=!1}}document.getElementById("btn-bulk-cancel")?.addEventListener("click",()=>{O.clear(),document.querySelectorAll(".row-checkbox").forEach(S=>S.checked=!1);let v=document.getElementById("select-all-checkbox");v&&(v.checked=!1),W()}),document.getElementById("btn-bulk-delete")?.addEventListener("click",()=>{if(O.size===0)return;let v=[...O];if(v.length>mt){z(`Maksimum ${mt} item dapat dihapus dalam satu permintaan bulk.`);return}let S=document.createElement("div");S.style.cssText="position:fixed;inset:0;background:rgba(0,0,0,.5);z-index:9999;display:flex;align-items:center;justify-content:center",S.innerHTML=`
      <div style="background:var(--bg-card);border-radius:var(--radius-xl);padding:28px;max-width:420px;width:90%;box-shadow:var(--shadow-lg);animation:fadeInUp .2s ease">
        <h3 style="margin:0 0 8px;color:var(--text-1);font-size:1rem;font-weight:700">\u26A0\uFE0F Hapus ${v.length} ${c}?</h3>
        <p style="margin:0 0 24px;color:var(--text-2);font-size:.875rem">Data yang dihapus tidak dapat dikembalikan.</p>
        <div style="display:flex;gap:8px;justify-content:flex-end">
          <button id="bulk-cancel-btn" class="btn btn-ghost">Batal</button>
          <button id="bulk-confirm-btn" class="btn btn-danger">Hapus ${v.length} Data</button>
        </div>
      </div>
    `,document.body.appendChild(S),S.querySelector("#bulk-cancel-btn").addEventListener("click",()=>S.remove()),S.querySelector("#bulk-confirm-btn").addEventListener("click",async()=>{let A=S.querySelector("#bulk-confirm-btn");A.disabled=!0,A.textContent="Menghapus...";let U=await X(`${a}/bulk`,{method:"DELETE",body:JSON.stringify({ids:v})});S.remove(),U.ok?(ae(`${v.length} ${c} berhasil dihapus.`),O.clear(),W(),at(a),d()):z(U.data?.error||"Gagal menghapus data.")})});let J=document.getElementById("filter-search"),$;J?.addEventListener("input",v=>{clearTimeout($),$=setTimeout(()=>{D.search=v.target.value,g=1,O.clear(),W(),d()},400)}),l?.forEach(v=>{(v.type==="select"||v.type==="combobox")&&(document.getElementById(`filter-${v.name}`)?.addEventListener("change",S=>{D[v.name]=S.target.value;let A=document.getElementById(`filter-sheet-${v.name}`);A&&(A.value=S.target.value),g=1,O.clear(),W(),d()}),document.getElementById(`filter-sheet-${v.name}`)?.addEventListener("change",S=>{D[v.name]=S.target.value;let A=document.getElementById(`filter-${v.name}`);A&&(A.value=S.target.value),g=1,O.clear(),W(),d(),document.getElementById("filter-options-wrapper")?.classList.remove("sheet-open")}))}),document.getElementById("btn-reset-filter")?.addEventListener("click",()=>{D={},J&&(J.value=""),l?.forEach(v=>{let S=document.getElementById(`filter-${v.name}`);S&&(S.value="");let A=document.getElementById(`filter-sheet-${v.name}`);A&&(A.value="")}),g=1,O.clear(),W(),d()}),document.getElementById("btn-reset-filter-sheet")?.addEventListener("click",()=>{D={},J&&(J.value=""),l?.forEach(v=>{let S=document.getElementById(`filter-${v.name}`);S&&(S.value="");let A=document.getElementById(`filter-sheet-${v.name}`);A&&(A.value="")}),g=1,O.clear(),W(),d(),document.getElementById("filter-options-wrapper")?.classList.remove("sheet-open")}),document.getElementById("btn-create")?.addEventListener("click",()=>y(null)),N&&document.addEventListener("click",function(v){let S=document.getElementById("aksi-menu-main"),A=document.getElementById("btn-aksi-main");S&&A&&!A.contains(v.target)&&!S.contains(v.target)&&S.classList.remove("show-aksi-menu")});let m=document.getElementById("btn-mobile-filter"),i=document.getElementById("filter-options-wrapper"),r=document.getElementById("btn-close-filter-sheet");if(m&&i&&(m.addEventListener("click",v=>{v.preventDefault(),i.classList.add("sheet-open")}),r&&r.addEventListener("click",v=>{v.preventDefault(),i.classList.remove("sheet-open")})),N){let H=function(L){A=L;let j=document.getElementById(`label-import-${N.moduleName}`),V=j?j.querySelector(".import-text"):null;V&&(V.innerText=L?"\u231B Memproses...":"\u{1F4E4} Import Excel"),j&&(j.style.pointerEvents=L?"none":"auto",j.style.opacity=L?"0.6":"1"),v&&(v.disabled=L)},ee=function(L){if(S){if(!L){S.style.display="none",S.innerHTML="";return}S.style.display="",S.innerHTML=nt(L),S.querySelectorAll("[data-import-action]").forEach(j=>{j.addEventListener("click",async()=>{let V=j.getAttribute("data-import-action"),te=j.getAttribute("data-fingerprint");if(V==="dismiss"){te&&bt(te),ee(null);return}if(A){z("Import masih berjalan. Mohon tunggu hingga selesai.");return}await Z(V,te)})})}},Q=function(L,j=""){let V=L?.data||L||{},te=!!(L?.idempotentReplay||V?.idempotentReplay),re=Number(V.inserted||0),le=Number(V.updated||0),ce=Number(V.skipped||0),se=[];re>0&&se.push(`${re} baru`),le>0&&se.push(`${le} diperbarui`),ce>0&&se.push(`${ce} dilewati`);let he=V.message||(se.length?se.join(", "):"Data berhasil diproses");ee({variant:"success",title:te?"\u267B\uFE0F Replay Idempoten Dikonfirmasi Server":"\u2705 Import Berhasil Dikonfirmasi",message:te?`${he} (Server mengembalikan hasil tersimpan tanpa menulis data duplikat.)`:he,idempotencyKey:L?.idempotencyKey||j||"",importJobId:L?.importJobId||V?.importJobId||"",fingerprint:L?.fingerprint||"",canReplay:!!((L?.fingerprint||U)&&(L?.idempotencyKey||j)),canDismiss:!0}),ae(te?"Replay idempoten berhasil (0 duplikat).":"Import berhasil!"),at(a),d()},ne=function(L){let V=Number(L?.status)===409&&L?.code==="IMPORT_IN_PROGRESS",te=!!(L?.isNetworkError||Number(L?.status)===0);ee({variant:V?"warning":"error",title:te?"\u{1F4E1} Koneksi Terputus Saat Import":V?"\u23F3 Import Sedang Berjalan / Terkunci (HTTP 409)":Number(L?.status)>=500?"\u26A0\uFE0F Import Gagal Sebagian / Server Error (HTTP 500)":`\u274C Import Gagal${L?.status?` (HTTP ${L.status})`:""}`,message:L?.message||"Gagal mengimport data",idempotencyKey:L?.idempotencyKey||"",importJobId:L?.importJobId||"",committedChunks:L?.committedChunks||0,totalChunks:L?.totalChunks||0,errorCode:L?.code||"",fingerprint:L?.fingerprint||"",canRetry:!!(L?.recoverable||te||Number(L?.status)>=500||V),canRecover:V,canDismiss:!0}),z(L?.message||"Gagal import data")};document.getElementById(`btn-export-${N.moduleName}`)?.addEventListener("click",async L=>{let j=L.target,V=j.innerHTML;j.innerHTML="\u23F3 Loading...",j.disabled=!0;let te=p&&!D.sort?{...D,sort:p}:{...D};try{await N.onExport(te)}catch{z("Gagal export data")}finally{j.innerHTML=V,j.disabled=!1}}),document.getElementById(`btn-template-${N.moduleName}`)?.addEventListener("click",()=>{N.onTemplate()});let v=document.getElementById(`input-import-${N.moduleName}`),S=document.getElementById("crud-import-banner"),A=!1,U=null;async function Z(L,j){let V=j?_t(j):null,te=L==="recover",re=V?.idempotencyKey||"";H(!0),ee({variant:"progress",title:te?"\u{1F6E0}\uFE0F Memulihkan & Melanjutkan Import...":"\u23F3 Mengirim / Melanjutkan Import...",message:"Mohon tunggu konfirmasi dari server. Jangan tutup halaman.",idempotencyKey:re,importJobId:V?.importJobId||"",committedChunks:V?.committedChunks||0,totalChunks:V?.totalChunks||0,fingerprint:j||"",canDismiss:!1});try{let le;L==="replay"&&U&&re?le=await N.onImport(U,{idempotencyKey:re,recover:!1}):V?.payloadStored?le=await Tt(j,{recover:te}):U?le=await N.onImport(U,{idempotencyKey:re||null,recover:te}):le=await Tt(j,{recover:te}),Q(le,re)}catch(le){ne(le)}finally{H(!1)}}let G=Wt(N.moduleName);if(G.length>0){let L=G[0],j=L.status==="IN_PROGRESS"||L.status==="CONFLICT_IN_PROGRESS";ee({variant:"recover",title:"\u{1F504} Sesi Import Belum Selesai Terdeteksi",message:L.payloadStored?L.errorMessage||"Import sebelumnya terhenti karena refresh atau gangguan koneksi. Anda dapat melanjutkan dengan Idempotency-Key yang sama.":"Import sebelumnya terhenti. Pilih ulang file Excel yang sama untuk melanjutkan dengan Idempotency-Key yang sama.",idempotencyKey:L.idempotencyKey,importJobId:L.importJobId,committedChunks:L.committedChunks,totalChunks:L.totalChunks,errorCode:L.errorCode||L.status,fingerprint:L.fingerprint,canRetry:!!L.payloadStored,canRecover:!!(L.payloadStored&&j),canDismiss:!0})}v?.addEventListener("change",async L=>{let j=L.target.files[0];if(j){if(A){z("Proses import sedang berjalan. Mohon tunggu hingga selesai."),v.value="";return}H(!0),ee({variant:"progress",title:"\u23F3 Membaca & Mengimport File Excel...",message:`Memproses "${j.name}". Sistem menunggu konfirmasi server sebelum menandai selesai.`,canDismiss:!1});try{let V=await xa(j);if(V.length===0)throw new Error("File kosong atau format salah");U=V;let te=await N.onImport(V,{recover:!1});Q(te)}catch(V){ne(V)}finally{H(!1),v&&(v.value="")}}})}async function d(){W();let v=document.getElementById("table-container");if(!v)return;v.innerHTML='<div class="loading-spinner"><div class="spinner"></div></div>';let S=g,A=20,U={...p?{sort:p}:{},...Object.fromEntries(Object.entries(D).filter(([,L])=>L))},H=new URLSearchParams({page:S,limit:A,...U}),ee=await X(`${a}?${H}`);if(!ee.ok){v.innerHTML=`<div class="empty-state"><p class="text-danger">Gagal memuat data: ${Y(ee.data?.error||"Error")}</p></div>`;return}let Q=ee.data?.data||ee.data||[],ne=ee.data?.pagination||{page:S,limit:A,total:Array.isArray(Q)?Q.length:0,pages:1};R&&(Q=R(Q)),!1,F&&F(Q);let Z=yn({columns:o,data:Q,onEdit:C?L=>y(L):null,onDelete:T?L=>q(L):null,actions:K.map(L=>({...L,handler:j=>L.handler(j,d)})),emptyText:`Tidak ada ${String(c||"").toLowerCase()}`,bulkSelect:w?{selectedIds:O,maxSelect:mt,onLimitExceeded:L=>z(L),onToggle:W}:null});v.innerHTML="",v.appendChild(Z),W();let G=document.getElementById("pagination-container");if(G&&(G.innerHTML="",ne&&ne.pages>1)){let L=_e({page:ne.page,pages:ne.pages,total:ne.total,limit:ne.limit,onPage:j=>{g=j,d()}});L&&G.appendChild(L)}}function b(v){let S=typeof s=="function"?s(v):s;return ka(S)}function y(v){let S=!!v,A=document.createElement("form");A.noValidate=!0,A.innerHTML=b(v);let U=typeof s=="function"?s(v):s;S&&Fn(A,v);let H=[],ee=ne=>{for(let Z of ne)if(Z.type==="row")ee(Z.fields||[]);else if(Z.type==="employee-combobox"){let G=A.querySelector(`.fc-crud-emp-mount[data-emp-mount="${Z.name}"]`);if(!G)continue;let L=v?v[Z.name]??Z.value??"":Z.value??"",j=v&&Z.employeeIdField?v[Z.employeeIdField]??null:Z.employeeIdValue??null,V=ze(G,{mode:Z.mode||"single",inputId:Z.id||`crud-${Z.name}`,inputName:Z.name,placeholder:Z.placeholder||"Cari nama karyawan...",required:!!Z.required,allowManual:Z.allowManual!==!1,hasEmployeeIdSchema:Z.hasEmployeeIdSchema!==!1,manualBadgeText:Z.manualBadgeText||"Belum terhubung ke Master Karyawan",masterBadgePrefix:Z.masterBadgePrefix||"Terhubung ke Master Karyawan",employees:Z.options||[],initialEmployeeId:j,initialEmployeeName:L});H.push({field:Z,ctrl:V})}};ee(U);let{close:Q}=pe({title:S?`Edit ${c}`:`Tambah ${c}`,content:A,size:"lg",confirmText:S?"Simpan Perubahan":`Tambah ${c}`,onConfirm:async(ne,Z)=>{if(!A.reportValidity())return;let G=ne.querySelector(".modal-confirm");G.disabled=!0,G.textContent="Menyimpan...";let L=An(A),j=typeof s=="function"?s(v):s,V=async ce=>{for(let se of ce)if(se.type==="row")await V(se.fields);else if(se.type==="combobox"&&L[se.name]){let he=L[se.name],Ke=(se.options||[]).find($e=>{let oe=String(typeof $e=="object"?$e.value:$e),Be=String(typeof $e=="object"?$e.label:$e);return oe===he||Be===he});if(Ke)L[se.name]=typeof Ke=="object"?Ke.value:Ke;else if(se.createApi){let $e={};$e[se.createApi.field]=he,se.createApi.extra&&Object.assign($e,se.createApi.extra);let oe=await X(se.createApi.path,{method:"POST",body:JSON.stringify($e)});if(oe.ok&&oe.data?.id)L[se.name]=oe.data.id;else if(oe.ok&&!oe.data?.id)L[se.name]=he;else throw new Error(`Gagal membuat master data: ${oe.data?.error||"Unknown error"}`)}}};try{await V(j)}catch(ce){z(ce.message),G.disabled=!1,G.textContent=S?"Simpan Perubahan":`Tambah ${c}`;return}for(let{field:ce,ctrl:se}of H){let he=se.getValue();ce.mode==="multi"?L[ce.name]=he.participantsString||null:(L[ce.name]=he.employee_name||null,ce.employeeIdField&&(L[ce.employeeIdField]=he.employee_id??null))}u&&(L=await u(L,v)),S&&v&&(v.row_version!==void 0&&v.row_version!==null?L.row_version=Number(v.row_version):L.row_version=1);let te=S?"PUT":"POST",re=S?`${a}/${v.id}`:a,le=await X(re,{method:te,body:JSON.stringify(L)});le.ok?(ae(S?`${c} berhasil diperbarui.`:`${c} berhasil ditambahkan.`),Z(),at(a),d()):(le.status===409?z(le.data?.error||"Konflik OCC: Data telah diubah oleh pengguna lain. Silakan muat ulang halaman."):z(le.data?.error||"Gagal menyimpan data."),G.disabled=!1,G.textContent=S?"Simpan Perubahan":`Tambah ${c}`)}})}function q(v){me(`Hapus ${c} ini? Tindakan tidak dapat dibatalkan.`,async()=>{let S=await X(`${a}/${v.id}`,{method:"DELETE"});S.ok?(ae(`${c} berhasil dihapus.`),at(a),d()):z(S.data?.error||"Gagal menghapus.")},`Hapus ${c}`)}return d(),d}function He(e){let t={Done:"badge-success",Aktif:"badge-success",Open:"badge-warning","In Progress":"badge-info",Pending:"badge-warning",Diproses:"badge-info",Selesai:"badge-success","Tidak Aktif":"badge-neutral",Resign:"badge-neutral",Cut:"badge-danger","Tidak Datang":"badge-danger"};return!e||e==="-"||String(e).trim()===""?"":`<span class="badge ${t[e]||"badge-neutral"}">${e}</span>`}function Pn(e){return`<span class="badge ${{"Inspeksi Hygiene & Aset Bangunan":"badge-info","General Cleaning":"badge-success","Deep Cleaning":"badge-purple",Fogging:"badge-warning"}[e]||"badge-neutral"}">${e||"-"}</span>`}function We(e){return`<span class="badge ${{Q1:"badge-info",Q2:"badge-success",Q3:"badge-warning",Q4:"badge-danger"}[e]||"badge-neutral"}">${e||"-"}</span>`}we();var Ba=[],Pt=[];async function Ln(e,t){Ba=await ve();let n=await Gt();Pt=["Berlin Ariansyah","Ade Surahman"];let a=u=>u&&!Pt.find(F=>String(typeof F=="object"?F.value:F).toLowerCase()===String(u).toLowerCase())?[...Pt,u]:Pt,o=u=>{if(!u||u==="-"||String(u).trim()==="")return"";let F=String(u).split("-");return F.length===3&&F[0].length===4?`${F[2]}-${F[1]}-${F[0]}`:u},s=t?t.get("dash_filter"):null,l=new Date,h=`${l.getFullYear()}-${String(l.getMonth()+1).padStart(2,"0")}`,p={},c=t&&t.get("month")?t.get("month"):null;s==="inspeksi"?p={status:"Done",activity_type:"Inspeksi Hygiene",...c?{month:c}:{}}:s==="gcdc"?p={status:"Done",activity_type:"GCDC",...c?{month:c}:{}}:s&&s.startsWith("period_")&&(p={period:s.replace("period_","").toUpperCase()});let E=new Date().getFullYear(),T=["Jan","Feb","Mar","Apr","Mei","Jun","Jul","Agu","Sep","Okt","Nov","Des"].map((u,F)=>{let R=String(F+1).padStart(2,"0");return{value:`${E}-${R}`,label:`${u} ${E}`}});De({container:e,title:"Jadwal Kegiatan",icon:"\u{1F4C5}",apiPath:"/api/schedule",bulkDelete:!1,itemLabel:"Jadwal",enableMobileFilterSheet:!0,defaultFilters:p,defaultSort:t&&t.get("sort")||new URLSearchParams(window.location.search).get("sort")||"opening_date_desc",columns:[{key:"branch_name",label:"Cabang"},{key:"activity_type",label:"Kegiatan",render:u=>Pn(u)},{key:"period",label:"Periode",render:u=>We(u)},{key:"pic",label:"PIC"},{key:"opening_date",label:"Tgl Opening",nowrap:!0,render:u=>o(u)},{key:"target_date",label:"Tgl Target",nowrap:!0,render:u=>o(u)},{key:"completion_date",label:"Tgl Selesai",nowrap:!0,render:u=>o(u)},{key:"status",label:"Status",render:u=>He(u)}],filterFields:[{type:"search",placeholder:"Cari cabang, kegiatan, PIC, catatan..."},{type:"select",name:"branch_id",label:"Cabang",options:Ba},{type:"select",name:"activity_type",label:"Kegiatan",options:[{value:"Inspeksi Hygiene",label:"Inspeksi Hygiene"},{value:"General Cleaning",label:"General Cleaning"},{value:"Deep Cleaning",label:"Deep Cleaning"},{value:"Fogging",label:"Fogging"},{value:"GCDC",label:"GCDC (GC & DC)"}]},{type:"select",name:"month",label:"Bulan",options:T},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"status",label:"Status",options:["Pending","In Progress","Done"]},{type:"select",name:"pic",label:"PIC",options:Pt}],formFields:u=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:Ba,value:u?.branch_id},{name:"activity_type",label:"Jenis Kegiatan",type:"select",required:!0,options:["Inspeksi Hygiene","General Cleaning","Deep Cleaning","Fogging"],value:u?.activity_type}]},{type:"row",fields:[{name:"period",label:"Periode",type:"select",required:!0,options:["Q1","Q2","Q3","Q4"],value:u?.period},{name:"pic",label:"PIC",type:"combobox",options:a(u?.pic),value:u?.pic}]},{type:"row",fields:[{name:"opening_date",label:"Tanggal Opening",type:"date",value:u?.opening_date},{name:"target_date",label:"Tanggal Target",type:"date",value:u?.target_date}]},{type:"row",fields:[{name:"completion_date",label:"Tanggal Selesai",type:"date",value:u?.completion_date},{name:"status",label:"Status",type:"select",required:!0,options:["Pending","In Progress","Done"],value:u?.status||""}]},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:u?.notes}],exportOptions:{moduleName:"schedule",onExport:async u=>{let R=(await ue("/api/schedule/export",u)).map(K=>({Cabang:K.branch_name||"",Kegiatan:K.activity_type||"",Periode:K.period||"",PIC:K.pic||"","Tgl Opening":K.opening_date||"","Tgl Target":K.target_date||"","Tgl Selesai":K.completion_date||"",Status:K.status||""}));ge(R,"Data_Jadwal_Kegiatan")},onTemplate:()=>{ge([{Cabang:"001. Pondok Bambu",Kegiatan:"General Cleaning",Periode:"Q1",PIC:"Fajar","Tgl Opening":"2024-02-01","Tgl Target":"2024-02-15","Tgl Selesai":"2024-02-14",Status:"Done"}],"Template_Import_Jadwal")},onImport:async(u,F={})=>{let K=(await X("/api/branches?all=1")).data?.data||[],I=k=>{if(!k)return null;let x=String(k||"").toLowerCase().trim(),_=x.match(/^(\d{3})/);if(_){let P=K.find(f=>String(f.code||"")===_[1]);if(P)return P.id}let B=K.find(P=>String(P.full_name||"").toLowerCase()===x||String(P.code||"").toLowerCase()===x||String(P.name||"").toLowerCase()===x);return B?B.id:null},N=k=>{if(k==null||k==="")return"";if(k instanceof Date&&!isNaN(k.getTime()))return k.toISOString().slice(0,10);let x=String(k).trim();if(x===""||x==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(x))return x.slice(0,10);if(/^\d{4,5}$/.test(x)){let B=Number(x);if(B>2e4&&B<99999){let P=new Date(Date.UTC(1899,11,30)+B*864e5);return isNaN(P.getTime())?"":P.toISOString().slice(0,10)}}let _=x.split(/[\/\-\.]/);if(_.length===3){let[B,P,f]=_.map(g=>g.trim());if(B.length===4&&P.length<=2&&f.length<=2)return`${B}-${P.padStart(2,"0")}-${f.padStart(2,"0")}`;if(f.length===4&&P.length<=2&&B.length<=2)return`${f}-${P.padStart(2,"0")}-${B.padStart(2,"0")}`}return x},w=u.map(k=>({branch_id:I(String(k.Cabang||"").trim()),branch_name:String(k.Cabang||"").trim(),activity_type:String(k.Kegiatan||"").trim(),period:String(k.Periode||"").trim(),pic:String(k.PIC||k.Pic||"").trim(),opening_date:N(k["Tgl Opening"]||k["Tanggal Opening"]||k["Tgl Openir"]),target_date:N(k["Tgl Target"]||k["Tanggal Target"]),completion_date:N(k["Tgl Selesai"]||k["Tanggal Selesai"]),status:String(k.Status||"").trim(),notes:String(k.Catatan||k.Keterangan||"").trim()})).filter(k=>k.activity_type&&k.period),M=await Le({module:"schedule",label:"Jadwal Kegiatan",endpoint:"/api/import/schedule",body:{rows:w,onDuplicate:"update"},idempotencyKey:F?.idempotencyKey,recover:!!F?.recover});return at("schedule"),M}}})}Fe();var Dn=["SDM","Cleaning","Aset","K3","Lainnya"],xr=["Open","In Progress","Done"],In=["SPV","AM","RCP","Perawat","FC","Berlin","Ade","Pattrel","Dentrel"];function yr(e){if(!e)return"-";try{let t=new Date(e);return isNaN(t.getTime())?e:t.toLocaleDateString("id-ID",{day:"numeric",month:"short",year:"numeric"})}catch{return e}}function Lt(e){if(e==null||e==="")return null;let t=Number(e);return Number.isInteger(t)&&t>=0?t:null}async function Mn(e,t){let a=[],o=[],s=[],l="all",h="all",p="all",c="",E=1,C=0,T=1,u=!1,F={total:0,open:0,inProgress:0,openAndProgress:0,done:0,currentMonth:0},R=null,K=!1;(t?t.get("dash_filter"):null)==="open"&&(l="all");let N=de(),w=ie(N,"issues","create"),M=ie(N,"issues","update"),k=ie(N,"issues","delete");try{[o,s]=await Promise.all([ve(),Ie()])}catch(r){console.error("Error fetching cache:",r)}function x(){let r={sort:"date_desc"};return c.trim()&&(r.search=c.trim()),p&&p!=="all"&&(r.branch_id=p),h&&h!=="all"&&(r.category=h),l&&l!=="all"&&(r.status=l),r}function _(){let r=(d,b)=>{let y=e.querySelector(d);y&&(y.textContent=b)};r("#issues-stat-total",F.total),r("#issues-stat-open-prog",F.openAndProgress),r("#issues-stat-done",F.done),r("#issues-stat-month",F.currentMonth),r("#issues-pill-count-all",F.total),r("#issues-pill-count-open",F.open),r("#issues-pill-count-inprogress",F.inProgress),r("#issues-pill-count-done",F.done)}async function B({refreshStats:r=!0}={}){if(K=!0,!!!e.querySelector("#issues-list-container"))i();else{let b=e.querySelector("#issues-list-container");b&&(b.innerHTML='<div class="issues-empty-state"><p style="color:#64748B; margin:0;">Memuat data permasalahan...</p></div>')}try{let b=new URLSearchParams({page:String(E),limit:String(20),...x()});(r||!u)&&b.set("include_stats","1");let y=await X(`/api/issues?${b.toString()}`);if(y.ok){let q=y.data?.data||[];a=Array.isArray(q)?q:[];let v=y.data?.pagination||{};if(C=Number(v.total??a.length),T=Math.max(1,Number(v.pages??Math.ceil(C/20))||1),E=Math.max(1,Number(v.page??E)),y.data?.stats&&(F=y.data.stats,u=!0),a.length===0&&E>T)return E=T,K=!1,B({refreshStats:r})}else a=[],C=0,T=1,z(y.data?.error||"Gagal memuat data permasalahan")}catch(b){console.warn("Gagal load issues:",b),a=[],C=0,T=1,z("Gagal memuat data permasalahan: "+b.message)}finally{K=!1,e.querySelector("#issues-list-container")?(_(),m()):i()}}async function P(){if(!ie(N,"export")){z("Anda tidak memiliki izin untuk mengekspor data.");return}try{let[{downloadExcel:r},d]=await Promise.all([Promise.resolve().then(()=>(we(),Oe)),ue("/api/issues/export",x())]),b=d.map((y,q)=>({No:q+1,"Tanggal Info":y.report_date||"",Cabang:y.branch_name||"",Kategori:y.category||"","Sumber Laporan":y.source||"",Keluhan:y.complaint||"","Nama FC / Security":y.employee_name||"","FC Spesialis":y.fc_specialist||"","Solusi / Tindak Lanjut":y.solution||"",Status:y.status||"","Tanggal Selesai":y.completion_date||"","Durasi Hari":y.day_count!=null?y.day_count:""}));r(b,`Data_Permasalahan_FCMS_${new Date().toISOString().split("T")[0]}`),ae("Berhasil mengekspor data permasalahan ke Excel")}catch(r){z("Gagal ekspor Excel: "+r.message)}}async function f(r,d){if(!M)return;let b=Lt(r?.row_version);if(b===null){z("Versi data (row_version) tidak tersedia. Silakan muat ulang data sebelum memperbarui status.");return}try{let y=d==="Done"?r.completion_date||new Date().toISOString().split("T")[0]:d==="Open"?null:r.completion_date,q=r.day_count;if(r.report_date&&y){let A=new Date(r.report_date),U=new Date(y);q=Math.max(0,Math.floor((U-A)/(1e3*60*60*24)))}else d==="Open"&&(q=null);let v={...r,row_version:b,status:d,completion_date:y,day_count:q},S=await X(`/api/issues/${r.id}`,{method:"PUT",body:JSON.stringify(v)});if(S.ok){let A=Lt(S.data?.data?.row_version??S.data?.row_version);A!==null&&(r.row_version=A),ae(`Status masalah diubah ke "${d}"`),await B()}else S.status===409?(z("Konflik data (HTTP 409): Issue ini telah diubah oleh pengguna atau sesi lain. Silakan muat ulang data terbaru."),await B()):z(S.data?.error||"Gagal memperbarui status")}catch(y){z("Terjadi kesalahan: "+y.message)}}function g(r=null){let d=!!r;if(d?!M:!w)return;let b=d?Lt(r?.row_version):null,y=new Date().toISOString().split("T")[0],q=r?.report_date||y,v=r?.category||"SDM",S=r?.source?String(r.source).trim():"",A=d&&S&&!In.includes(S),U=r?.status||"Open",H=document.createElement("div");H.className="issue-modal-body",H.innerHTML=`
      <div style="display:flex; flex-direction:column; gap:14px;">
        ${d?`<input type="hidden" id="iss-modal-row-version" value="${b!==null?b:""}" />`:""}
        <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Tanggal Info <span style="color:#DC2626;">*</span>
            </label>
            <input type="date" id="iss-modal-date" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              value="${q}" required />
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Cabang Klinik <span style="color:#DC2626;">*</span>
            </label>
            <select id="iss-modal-branch" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" required>
              <option value="">-- Pilih Cabang --</option>
              ${o.map(V=>`
                <option value="${V.value}" ${String(r?.branch_id)===String(V.value)?"selected":""}>
                  ${V.label}
                </option>
              `).join("")}
            </select>
          </div>
        </div>

        <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Kategori <span style="color:#DC2626;">*</span>
            </label>
            <select id="iss-modal-category" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" required>
              ${Dn.map(V=>`
                <option value="${V}" ${v===V?"selected":""}>${V}</option>
              `).join("")}
            </select>
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Sumber Laporan
            </label>
            <select id="iss-modal-source" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;">
              <option value="">-- Pilih Sumber Laporan --</option>
              ${In.map(V=>`<option value="${V}" ${S===V?"selected":""}>${V}</option>`).join("")}
              ${A?`<option value="${S.replace(/"/g,"&quot;")}" selected>${S} (Historis)</option>`:""}
            </select>
          </div>
        </div>

        <div>
          <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
            Detail Keluhan / Permasalahan <span style="color:#DC2626;">*</span>
          </label>
          <textarea id="iss-modal-complaint" class="form-control" rows="3" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
            placeholder="Jelaskan kendala, kronologi, atau komplain..." required>${r?.complaint||""}</textarea>
        </div>

        <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Nama Personil FC / Security
            </label>
            <div id="iss-modal-emp-mount"></div>
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              FC Spesialis / Penanggung Jawab
            </label>
            <input list="iss-spec-options" id="iss-modal-spec" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              placeholder="Contoh: Mizwar, Ade, Berlin..." value="${r?.fc_specialist?String(r.fc_specialist).replace(/"/g,"&quot;"):""}" />
            <datalist id="iss-spec-options">
              ${s.map(V=>`<option value="${V.full_name||V.label||""}">`).join("")}
            </datalist>
          </div>
        </div>

        <div>
          <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
            Solusi / Tindak Lanjut Penanganan
          </label>
          <textarea id="iss-modal-solution" class="form-control" rows="2" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
            placeholder="Tindakan yang telah atau akan diambil untuk menyelesaikan masalah...">${r?.solution||""}</textarea>
        </div>

        <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr 1fr; gap:10px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Status Masalah <span style="color:#DC2626;">*</span>
            </label>
            <select id="iss-modal-status" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" required>
              ${xr.map(V=>`
                <option value="${V}" ${U===V?"selected":""}>${V}</option>
              `).join("")}
            </select>
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Tanggal Selesai
            </label>
            <input type="date" id="iss-modal-completion" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              value="${r?.completion_date||""}" />
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Durasi (Hari)
            </label>
            <input type="number" id="iss-modal-daycount" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #29364B; border-radius:8px; background:#1B2940; color:#F8FAFC;" 
              value="${r?.day_count!=null?r.day_count:""}" readonly placeholder="Otomatis" />
          </div>
        </div>
      </div>
    `;let ee=H.querySelector("#iss-modal-date"),Q=H.querySelector("#iss-modal-completion"),ne=H.querySelector("#iss-modal-daycount"),Z=H.querySelector("#iss-modal-status");function G(){if(ee.value&&Q.value){let V=new Date(ee.value),te=new Date(Q.value),re=Math.floor((te-V)/(1e3*60*60*24));ne.value=isNaN(re)?"":Math.max(0,re)}else ne.value=""}ee.addEventListener("change",G),Q.addEventListener("change",G),Z.addEventListener("change",()=>{Z.value==="Done"&&!Q.value?(Q.value=new Date().toISOString().split("T")[0],G()):Z.value==="Open"&&(Q.value="",ne.value="")});let L=ze(H.querySelector("#iss-modal-emp-mount"),{mode:"single",inputId:"iss-modal-employee",inputName:"employee_name",placeholder:"Cari nama karyawan...",allowManual:!0,hasEmployeeIdSchema:!1,employees:s,initialEmployeeName:r?.employee_name||""}),j=!1;pe({title:d?"\u270F\uFE0F Edit Data Permasalahan":"\u2795 Tambah Data Permasalahan",content:H,size:"md",confirmText:d?"Simpan Perubahan":"Simpan Permasalahan",confirmClass:"btn-primary",onConfirm:async(V,te)=>{if(j)return;let re=H.querySelector("#iss-modal-date").value,le=H.querySelector("#iss-modal-branch").value,ce=H.querySelector("#iss-modal-category").value,se=(H.querySelector("#iss-modal-source").value||"").trim(),he=(H.querySelector("#iss-modal-complaint").value||"").trim(),$e=L.getValue().employee_name||"",oe=(H.querySelector("#iss-modal-spec").value||"").trim(),Be=(H.querySelector("#iss-modal-solution").value||"").trim(),fe=H.querySelector("#iss-modal-status").value,Te=H.querySelector("#iss-modal-completion").value||null;if(!re||!le||!he||!ce){z("Harap lengkapi Tanggal, Cabang, Kategori, dan Keluhan!");return}let ke=null;if(d){let ye=H.querySelector("#iss-modal-row-version"),Je=ye?ye.value:b;if(ke=Lt(Je),ke===null){z("Versi data (row_version) tidak tersedia. Silakan muat ulang data sebelum menyimpan perubahan.");return}}let Pe=null;if(re&&Te){let ye=new Date(re),Je=new Date(Te);Pe=Math.max(0,Math.floor((Je-ye)/(1e3*60*60*24)))}let je={report_date:re,branch_id:Number(le),category:ce,source:se,complaint:he,employee_name:$e,fc_specialist:oe,solution:Be,status:fe,completion_date:Te,day_count:Pe,...d?{row_version:ke}:{}},Ae=V?.querySelector?.(".modal-confirm");j=!0,Ae&&(Ae.disabled=!0);try{let ye;if(d?ye=await X(`/api/issues/${r.id}`,{method:"PUT",body:JSON.stringify(je)}):ye=await X("/api/issues",{method:"POST",body:JSON.stringify(je)}),ye.ok){if(d){let Je=Lt(ye.data?.data?.row_version??ye.data?.row_version);Je!==null&&r&&(r.row_version=Je)}else E=1;ae(d?"Permasalahan berhasil diperbarui!":"Permasalahan baru berhasil ditambahkan!"),te(),await B()}else d&&ye.status===409?(z("Konflik data (HTTP 409): Issue ini telah diubah oleh pengguna atau sesi lain. Silakan muat ulang data sebelum mengedit kembali."),te(),await B()):z(ye.data?.error||"Gagal menyimpan permasalahan")}catch(ye){z("Terjadi kesalahan: "+ye.message)}finally{j=!1,Ae&&(Ae.disabled=!1)}}})}async function D(r){k&&me({title:"Hapus Permasalahan?",message:`Apakah Anda yakin ingin menghapus tiket permasalahan dari <strong>${r.branch_name||"cabang"}</strong>:<br><br><em>"${r.complaint}"</em>?`,confirmText:"Ya, Hapus",confirmClass:"btn-danger",onConfirm:async()=>{try{let d=await X(`/api/issues/${r.id}`,{method:"DELETE"});d.ok?(ae("Permasalahan berhasil dihapus"),await B()):z(d.data?.error||"Gagal menghapus data")}catch(d){z("Terjadi kesalahan: "+d.message)}}})}function O(){return a}function W(r){return r.length===0?`
            <div class="issues-empty-state">
              <div style="font-size:2.8rem; margin-bottom:10px;">\u{1F50D}</div>
              <h3 style="font-size:1.15rem; font-weight:700; color:#1E293B; margin:0 0 6px 0;">Tidak ada permasalahan yang cocok</h3>
              <p style="font-size:0.85rem; color:#64748B; margin:0 0 16px 0;">Coba sesuaikan kata kunci pencarian atau filter status yang dipilih.</p>
              <button id="issues-reset-filter" class="btn btn-secondary" style="font-size:0.82rem; padding:7px 16px; border-radius:8px;">Reset Semua Filter</button>
            </div>
      `:r.map(d=>{let b=d.category||"Lainnya",y=`issues-category-${b.toLowerCase().replace(/[^a-z0-9]/g,"")}`||"issues-category-lainnya",q=d.status||"Open",v="issues-status-open",S="\u26A0\uFE0F Open";q==="In Progress"?(v="issues-status-progress",S="\u23F3 In Progress"):q==="Done"&&(v="issues-status-done",S="\u2705 Done");let A=d.priority||"Normal",U="priority-normal";return(A==="High"||A==="Tinggi")&&(U="priority-high"),(A==="Low"||A==="Rendah")&&(U="priority-low"),`
        <div class="issues-card" data-id="${d.id}">
          <div class="issues-card-top">
            <span class="issues-category-badge ${y}">
              ${b==="SDM"?"\u{1F465}":b==="Cleaning"?"\u{1F9F9}":b==="Aset"?"\u{1F4E6}":"\u{1F4CB}"} ${b}
            </span>
            <div style="display:flex; align-items:center; gap:6px;">
              <span class="issues-priority-badge ${U}">${A}</span>
              <span class="issues-status-badge ${v}">${S}</span>
            </div>
          </div>

          <div>
            <div class="issues-title">${d.complaint||"Tidak ada uraian permasalahan"}</div>
            <div style="font-size:0.75rem; color:#94A3B8; margin-top:2px;">
              ID #${d.id} \u2022 Sumber: ${d.source||"-"}
            </div>
          </div>

          <div class="issues-details-grid">
            <div class="issues-detail-item">
              <span>\u{1F3E2}</span>
              <span style="font-weight:600; overflow:hidden; text-overflow:ellipsis; white-space:nowrap;" title="${d.branch_name||"-"}">
                ${d.branch_name||"Semua Cabang"}
              </span>
            </div>
            <div class="issues-detail-item">
              <span>\u{1F4C5}</span>
              <span>${d.report_date?yr(d.report_date):"-"}</span>
            </div>
            <div class="issues-detail-item">
              <span>\u{1F464}</span>
              <span style="overflow:hidden; text-overflow:ellipsis; white-space:nowrap;" title="${d.employee_name||"-"}">
                ${d.employee_name||"-"}
              </span>
            </div>
            <div class="issues-detail-item">
              <span>\u{1F552}</span>
              <span>${d.created_at?d.created_at.slice(0,10):"-"}</span>
            </div>
          </div>

          ${d.solution?`
            <div class="issues-solution-box">
              <div class="issues-solution-title">
                <span>\u{1F4A1}</span> Solusi / Tindakan:
              </div>
              <div>${d.solution}</div>
            </div>
          `:""}

          <div class="issues-card-footer">
            <div style="display:flex; align-items:center; gap:6px;">
              ${M?`
                ${q==="Open"?`
                  <button class="issues-quick-btn start" data-action="quick-progress" title="Mulai Penanganan">
                    \u25B6\uFE0F Proses
                  </button>
                `:""}
                ${q==="In Progress"?`
                  <button class="issues-quick-btn complete" data-action="quick-done" title="Tandai Selesai">
                    \u2705 Selesai
                  </button>
                `:""}
                ${q==="Done"?`
                  <button class="issues-quick-btn reopen" data-action="quick-reopen" title="Buka Kembali Masalah">
                    \u{1F504} Reopen
                  </button>
                `:""}
              `:""}
            </div>

            <div class="issues-action-btns">
              ${M?`
                <button class="issues-icon-btn edit" data-action="edit" data-id="${d.id}" title="Edit Masalah">
                  <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>
                </button>
              `:""}
              ${k?`
                <button class="issues-icon-btn delete" data-action="delete" data-id="${d.id}" title="Hapus Masalah">
                  <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><polyline points="3 6 5 6 21 6"/><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"/><line x1="10" y1="11" x2="10" y2="17"/><line x1="14" y1="11" x2="14" y2="17"/></svg>
                </button>
              `:""}
            </div>
          </div>
        </div>
      `}).join("")}function J(){let r=e.querySelector("#issues-reset-filter");r&&r.addEventListener("click",()=>{c="",p="all",h="all",l="all";let d=e.querySelector("#issues-search-input");d&&(d.value="");let b=e.querySelector("#issues-branch-filter");b&&(b.value="all");let y=e.querySelector("#issues-category-filter");y&&(y.value="all"),e.querySelectorAll(".issues-pill").forEach(q=>{q.classList.toggle("active",q.getAttribute("data-status")==="all")}),E=1,B({refreshStats:!1})}),e.querySelectorAll(".issues-card").forEach(d=>{let b=Number(d.dataset.id),y=a.find(H=>H.id===b);if(!y)return;let q=d.querySelector('[data-action="edit"]');q&&q.addEventListener("click",()=>g(y));let v=d.querySelector('[data-action="delete"]');v&&v.addEventListener("click",()=>D(y));let S=d.querySelector('[data-action="quick-progress"]');S&&S.addEventListener("click",()=>f(y,"In Progress"));let A=d.querySelector('[data-action="quick-done"]');A&&A.addEventListener("click",()=>f(y,"Done"));let U=d.querySelector('[data-action="quick-reopen"]');U&&U.addEventListener("click",()=>f(y,"Open"))})}function $(){let r=e.querySelector("#issues-pagination-container");if(!r)return;r.innerHTML="";let d=_e({page:E,pages:T,total:C,limit:20,onPage:b=>{E=b,B({refreshStats:!1})}});d&&r.appendChild(d)}function m(){let r=e.querySelector("#issues-list-container");if(!r)return;let d=O();r.innerHTML=W(d),$(),J()}function i(){let r=O(),d=F.total,b=F.openAndProgress,y=F.done,q=F.currentMonth,v=F.total,S=F.open,A=F.inProgress,U=F.done;e.innerHTML=`
      <style>.issues-page{display:flex;flex-direction:column;gap:20px}.issues-hero-card{background:linear-gradient(135deg,#0F172A 0%,#1E293B 100%);border-radius:16px;padding:24px 28px;color:white;box-shadow:0 10px 25px -5px rgba(15,23,42,0.25);position:relative;overflow:hidden}.issues-hero-card::after{content:'';position:absolute;top:-40px;right:-40px;width:220px;height:220px;background:radial-gradient(circle,rgba(245,158,11,0.18) 0%,transparent 70%);border-radius:50%;pointer-events:none}.issues-hero-actions{display:flex;align-items:center;gap:8px;flex-wrap:wrap}.issues-hero-actions .btn{height:38px;box-sizing:border-box;transition:all 0.2s ease}.issues-btn-add:hover{background:#1D4ED8 !important;transform:translateY(-1px)}.issues-btn-excel:hover,.issues-btn-refresh:hover{background:rgba(255,255,255,0.16) !important;border-color:rgba(255,255,255,0.3) !important;transform:translateY(-1px)}.issues-stats-row{display:grid;grid-template-columns:repeat(4,1fr);gap:14px;margin-top:20px}.issues-stat-box{background:rgba(255,255,255,0.05);border:1px solid rgba(255,255,255,0.1);border-radius:12px;padding:14px 18px;display:flex;align-items:center;gap:14px;backdrop-filter:blur(8px);transition:transform 0.15s ease,background 0.15s ease}.issues-stat-box:hover{background:rgba(255,255,255,0.08);transform:translateY(-2px)}.issues-stat-icon{width:44px;height:44px;border-radius:10px;display:flex;align-items:center;justify-content:center;font-size:1.3rem;flex-shrink:0}.issues-stat-val{font-size:1.55rem;font-weight:800;line-height:1.1;letter-spacing:-0.02em}.issues-stat-lbl{font-size:0.72rem;text-transform:uppercase;color:#94A3B8;font-weight:700;letter-spacing:0.05em;margin-top:2px}.issues-controls-card{background:#172033;border-radius:14px;border:1px solid #29364B;padding:16px 20px;box-shadow:0 2px 8px rgba(0,0,0,0.03);display:flex;flex-direction:column;gap:14px}.issues-search-row{display:flex;gap:12px;flex-wrap:wrap}.issues-search-wrap{position:relative;flex:1;min-width:240px}.issues-search-icon{position:absolute;left:14px;top:50%;transform:translateY(-50%);color:#94A3B8;pointer-events:none}.issues-search-input{width:100%;padding:10px 14px 10px 40px;border:1px solid #CBD5E1;border-radius:10px;font-size:0.9rem;outline:none;transition:border-color 0.15s ease}.issues-search-input:focus{border-color:#2563EB;box-shadow:0 0 0 3px rgba(37,99,235,0.1)}.issues-select{padding:10px 14px;border:1px solid #CBD5E1;border-radius:10px;font-size:0.9rem;color:#F8FAFC;background:#172033;outline:none;min-width:160px;cursor:pointer}.issues-select:focus{border-color:#2563EB}.issues-filter-pills{display:flex;align-items:center;gap:8px;overflow-x:auto;white-space:nowrap;scrollbar-width:none;padding-bottom:2px}.issues-filter-pills::-webkit-scrollbar{display:none}.issues-pill{padding:6px 14px;border-radius:20px;font-size:0.82rem;font-weight:600;cursor:pointer;border:1px solid #29364B;background:#1B2940;color:#64748B;transition:all 0.15s ease;display:inline-flex;align-items:center;gap:6px;user-select:none;flex-shrink:0}.issues-pill:hover{background:#111827;color:#F8FAFC}.issues-pill.active{background:#0F172A;color:white;border-color:#0F172A}.issues-pill-dot{width:8px;height:8px;border-radius:50%}.issues-pill-count{padding:1px 6px;border-radius:10px;font-size:0.72rem;background:rgba(0,0,0,0.06)}.issues-pill.active .issues-pill-count{background:rgba(255,255,255,0.2);color:white}.issues-cards-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(370px,1fr));gap:16px}.issues-card{background:#172033;border-radius:14px;border:1px solid #29364B;padding:18px;box-shadow:0 2px 8px rgba(0,0,0,0.04);display:flex;flex-direction:column;gap:14px;transition:transform 0.15s ease,box-shadow 0.15s ease}.issues-card:hover{transform:translateY(-2px);box-shadow:0 6px 16px rgba(0,0,0,0.08);border-color:#CBD5E1}.issues-card-top{display:flex;justify-content:space-between;align-items:center;gap:8px}.issues-category-badge{display:inline-flex;align-items:center;gap:5px;padding:4px 10px;border-radius:20px;font-size:0.76rem;font-weight:700}.issues-category-sdm{background:#EEF2FF;color:#4F46E5;border:1px solid #E0E7FF}.issues-category-cleaning{background:#ECFEFF;color:#0891B2;border:1px solid #CFFAFE}.issues-category-aset{background:#F5F3FF;color:#7C3AED;border:1px solid #EDE9FE}.issues-category-k3{background:#FEF3C7;color:#D97706;border:1px solid #FDE68A}.issues-category-lainnya{background:#111827;color:#475569;border:1px solid #29364B}.issues-status-badge{display:inline-flex;align-items:center;gap:5px;padding:4px 10px;border-radius:20px;font-size:0.76rem;font-weight:700}.issues-status-open{background:#FEF3C7;color:#B45309;border:1px solid #FDE68A}.issues-status-inprogress{background:#DBEAFE;color:#1D4ED8;border:1px solid #BFDBFE}.issues-status-done{background:#DCFCE7;color:#15803D;border:1px solid #BBF7D0}.issues-complaint-box{background:#1B2940;border-left:4px solid #F59E0B;border-radius:0 10px 10px 0;padding:12px 14px;font-size:0.92rem;color:#0F172A;line-height:1.5;font-weight:600}.issues-complaint-box.status-done{border-left-color:#10B981}.issues-complaint-box.status-inprogress{border-left-color:#3B82F6}.issues-person-box{background:#1B2940;border:1px solid #F1F5F9;border-radius:10px;padding:12px;display:flex;align-items:center;gap:12px}.issues-avatar-box{width:44px;height:44px;border-radius:10px;background:#4F46E5;color:white;display:flex;align-items:center;justify-content:center;font-weight:800;font-size:1rem;flex-shrink:0}.issues-person-name{font-size:0.95rem;font-weight:800;color:#0F172A}.issues-person-sub{font-size:0.74rem;color:#64748B;font-weight:600;text-transform:uppercase;letter-spacing:0.04em}.issues-details-grid{display:grid;grid-template-columns:1fr 1fr;gap:8px 12px;font-size:0.82rem;color:#334155;padding:2px 4px}.issues-detail-item{display:flex;align-items:center;gap:6px}.issues-solution-box{background:#F0FDF4;border:1px solid #DCFCE7;border-radius:10px;padding:10px 12px;font-size:0.82rem;color:#166534;line-height:1.45}.issues-solution-title{font-weight:700;display:flex;align-items:center;gap:6px;margin-bottom:3px}.issues-card-footer{display:flex;justify-content:space-between;align-items:center;border-top:1px solid #F1F5F9;padding-top:12px;margin-top:auto;gap:8px}.issues-action-btns{display:flex;align-items:center;gap:6px}.issues-icon-btn{width:34px;height:34px;border-radius:8px;border:1px solid #29364B;background:#172033;color:#64748B;display:flex;align-items:center;justify-content:center;cursor:pointer;transition:all 0.15s ease}.issues-icon-btn:hover{background:#111827;color:#0F172A;border-color:#CBD5E1}.issues-icon-btn.delete:hover{background:#FEE2E2;color:#DC2626;border-color:#FCA5A5}.issues-quick-btn{font-size:0.76rem;font-weight:700;padding:6px 10px;border-radius:8px;cursor:pointer;border:1px solid transparent;display:inline-flex;align-items:center;gap:4px;transition:all 0.15s ease}.issues-quick-btn.start{background:#EFF6FF;color:#2563EB;border-color:#BFDBFE}.issues-quick-btn.start:hover{background:#DBEAFE}.issues-quick-btn.complete{background:#F0FDF4;color:#16A34A;border-color:#BBF7D0}.issues-quick-btn.complete:hover{background:#DCFCE7}.issues-quick-btn.reopen{background:#FFFBEB;color:#D97706;border-color:#FDE68A}.issues-quick-btn.reopen:hover{background:#FEF3C7}.issues-empty-state{grid-column:1 / -1;background:#172033;border:2px dashed #29364B;border-radius:16px;padding:48px 24px;text-align:center}@media (max-width:768px){.issues-hero-card{padding:18px 16px}.issues-stats-row{grid-template-columns:1fr 1fr;gap:10px}.issues-stat-box{padding:10px 12px}.issues-stat-icon{width:38px;height:38px;font-size:1.1rem}.issues-stat-val{font-size:1.3rem}.issues-hero-actions{width:100% !important;display:flex !important;align-items:center !important;gap:8px !important;flex-wrap:wrap !important;overflow-x:visible !important}.issues-hero-actions::-webkit-scrollbar{display:none}.issues-btn-add{width:100% !important;justify-content:center !important;padding:9px 12px !important;font-size:0.85rem !important;white-space:nowrap !important;height:40px !important}.issues-btn-excel{flex:1 1 auto !important;justify-content:center !important;padding:8px 12px !important;font-size:0.85rem !important;white-space:nowrap !important;height:38px !important}.issues-btn-refresh{width:38px !important;height:38px !important;padding:0 !important;flex-shrink:0 !important;justify-content:center !important}.issues-controls-card{padding:14px;gap:12px}.issues-search-row{flex-direction:column;gap:10px}.issues-select{width:100%}.issues-cards-grid{grid-template-columns:minmax(0,1fr) !important;width:100% !important;max-width:100% !important;min-width:0 !important;box-sizing:border-box !important}.issues-card{width:100% !important;max-width:100% !important;min-width:0 !important;box-sizing:border-box !important;overflow:hidden;padding:14px !important}}</style>

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
              ${w?`
                <button id="issues-btn-add" class="btn btn-primary issues-btn-add" style="background:#2563EB; height:38px; padding:0 16px; border-radius:10px; font-weight:700; display:inline-flex; align-items:center; gap:6px; box-shadow:0 4px 12px rgba(37,99,235,0.3); white-space:nowrap;">
                  <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24"><line x1="12" y1="5" x2="12" y2="19"/><line x1="5" y1="12" x2="19" y2="12"/></svg>
                  <span>Tambah Permasalahan</span>
                </button>
              `:""}
              
              ${ie(N,"export")?`<button id="issues-btn-excel" class="btn btn-ghost issues-btn-excel" style="background:rgba(255,255,255,0.08); border:1px solid rgba(255,255,255,0.15); color:white; height:38px; padding:0 14px; border-radius:10px; font-weight:600; display:inline-flex; align-items:center; gap:6px; white-space:nowrap;">
                <span>\u{1F4E5}</span>
                <span>Ekspor Excel</span>
              </button>`:""}

              <button id="issues-btn-refresh" class="btn btn-ghost issues-btn-refresh" style="background:rgba(255,255,255,0.08); border:1px solid rgba(255,255,255,0.15); color:white; width:38px; height:38px; padding:0; border-radius:10px; display:inline-flex; align-items:center; justify-content:center; flex-shrink:0;" title="Muat Ulang Data">
                <svg width="16" height="16" fill="none" stroke="currentColor" stroke-width="2.2" viewBox="0 0 24 24" class="${K?"spin-icon":""}"><path d="M23 4v6h-6"/><path d="M1 20v-6h6"/><path d="M3.51 9a9 9 0 0 1 14.85-3.36L23 10M1 14l4.64 4.36A9 9 0 0 0 20.49 15"/></svg>
              </button>
            </div>
          </div>

          <!-- 4 KPI Stat Boxes -->
          <div class="issues-stats-row">
            <div class="issues-stat-box">
              <div class="issues-stat-icon" style="background:rgba(239, 68, 68, 0.2); color:#FCA5A5;">\u26A0\uFE0F</div>
              <div>
                <div class="issues-stat-val" id="issues-stat-total">${d}</div>
                <div class="issues-stat-lbl">Total Permasalahan</div>
              </div>
            </div>

            <div class="issues-stat-box">
              <div class="issues-stat-icon" style="background:rgba(245, 158, 11, 0.2); color:#FCD34D;">\u23F3</div>
              <div>
                <div class="issues-stat-val" id="issues-stat-open-prog">${b}</div>
                <div class="issues-stat-lbl">Open & In Progress</div>
              </div>
            </div>

            <div class="issues-stat-box">
              <div class="issues-stat-icon" style="background:rgba(16, 185, 129, 0.2); color:#6EE7B7;">\u2705</div>
              <div>
                <div class="issues-stat-val" id="issues-stat-done">${y}</div>
                <div class="issues-stat-lbl">Tuntas (Done)</div>
              </div>
            </div>

            <div class="issues-stat-box">
              <div class="issues-stat-icon" style="background:rgba(59, 130, 246, 0.2); color:#93C5FD;">\u{1F4C5}</div>
              <div>
                <div class="issues-stat-val" id="issues-stat-month">${q}</div>
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
                placeholder="Cari keluhan, nama personil, sumber, cabang..." value="${c?String(c).replace(/"/g,"&quot;"):""}" />
            </div>

            <select id="issues-branch-filter" class="issues-select">
              <option value="all">Semua Cabang (${o.length})</option>
              ${o.map(j=>`
                <option value="${j.value}" ${String(p)===String(j.value)?"selected":""}>
                  ${j.label}
                </option>
              `).join("")}
            </select>

            <select id="issues-category-filter" class="issues-select">
              <option value="all">Semua Kategori</option>
              ${Dn.map(j=>`
                <option value="${j}" ${h===j?"selected":""}>
                  Kategori: ${j}
                </option>
              `).join("")}
            </select>
          </div>

          <!-- Status Horizontal Swipeable Pills -->
          <div class="issues-filter-pills">
            <div class="issues-pill ${l==="all"?"active":""}" data-status="all">
              <span>Semua Status</span>
              <span class="issues-pill-count" id="issues-pill-count-all">${v}</span>
            </div>
            <div class="issues-pill ${l==="Open"?"active":""}" data-status="Open">
              <span class="issues-pill-dot" style="background:#F59E0B;"></span>
              <span>Open</span>
              <span class="issues-pill-count" id="issues-pill-count-open">${S}</span>
            </div>
            <div class="issues-pill ${l==="In Progress"?"active":""}" data-status="In Progress">
              <span class="issues-pill-dot" style="background:#3B82F6;"></span>
              <span>In Progress</span>
              <span class="issues-pill-count" id="issues-pill-count-inprogress">${A}</span>
            </div>
            <div class="issues-pill ${l==="Done"?"active":""}" data-status="Done">
              <span class="issues-pill-dot" style="background:#10B981;"></span>
              <span>Done</span>
              <span class="issues-pill-count" id="issues-pill-count-done">${U}</span>
            </div>
          </div>
        </div>

        <!-- Exclusive Cards Grid -->
        <div id="issues-list-container" class="issues-cards-grid">
          ${W(r)}
        </div>

        <!-- Pagination Container -->
        <div id="issues-pagination-container"></div>

      </div>
    `;let H=e.querySelector("#issues-search-input");H&&H.addEventListener("input",j=>{c=j.target.value,clearTimeout(R),R=setTimeout(()=>{E=1,B({refreshStats:!1})},350)});let ee=e.querySelector("#issues-branch-filter");ee&&ee.addEventListener("change",j=>{p=j.target.value,E=1,B({refreshStats:!1})});let Q=e.querySelector("#issues-category-filter");Q&&Q.addEventListener("change",j=>{h=j.target.value,E=1,B({refreshStats:!1})}),e.querySelectorAll(".issues-pill").forEach(j=>{j.addEventListener("click",()=>{l=j.dataset.status,e.querySelectorAll(".issues-pill").forEach(V=>{V.classList.toggle("active",V===j)}),E=1,B({refreshStats:!1})})});let ne=e.querySelector("#issues-reset-filter");ne&&ne.addEventListener("click",()=>{c="",p="all",h="all",l="all",E=1,clearTimeout(R),H&&(H.value=""),ee&&(ee.value="all"),Q&&(Q.value="all"),e.querySelectorAll(".issues-pill").forEach(j=>{j.classList.toggle("active",j.dataset.status==="all")}),B({refreshStats:!1})});let Z=e.querySelector("#issues-btn-add");Z&&Z.addEventListener("click",()=>g());let G=e.querySelector("#issues-btn-excel");G&&G.addEventListener("click",()=>P());let L=e.querySelector("#issues-btn-refresh");L&&L.addEventListener("click",()=>B({refreshStats:!0})),$(),J()}await B({refreshStats:!0})}Fe();var Qt=["Batch 1","Batch 2","Batch 3","Batch 4","Batch 5"],zn=["Berlin Ariansyah","Ade Surahman","Mizwar","Fajar","Ade","Berlin"];function vr(e){if(!e)return"-";try{let t=new Date(e);return isNaN(t.getTime())?e:t.toLocaleDateString("id-ID",{day:"numeric",month:"short",year:"numeric"})}catch{return e}}function Ta(e){if(!e)return"-";try{let t=JSON.parse(e);if(Array.isArray(t))return t.join(", ")}catch{}return String(e)}async function Kn(e){let t=[],n=[],a=[],o="all",s="all",l="all",h="",p=!1,c=1,E=20,C=0,T=1,u={totalSessions:0,totalBatches:0,currentMonth:0,totalParticipants:0},F=null,R=!1,K=de(),I=ie(K,"training","write"),N=ie(K,"training","delete");try{[n,a]=await Promise.all([ve(),Ie()])}catch($){console.error("Error fetching cache:",$)}function w(){let $={sort:"date_desc"};return h&&h.trim()&&($.search=h.trim()),s&&s!=="all"&&($.branch_id=s),o&&o!=="all"&&($.batch=o),l&&l!=="all"&&($.trainer=l),$}function M(){let $=(m,i)=>{let r=e.querySelector(m);r&&(r.textContent=String(i??0))};$("#tr-stat-sessions",u.totalSessions),$("#tr-stat-participants",u.totalParticipants),$("#tr-stat-batches",u.totalBatches),$("#tr-stat-month",u.currentMonth),$("#tr-pill-count-all",u.totalSessions)}async function k({refreshStats:$=!0}={}){p=!0;try{let m=w(),i=new URLSearchParams({page:String(c),limit:String(E),...m});$&&i.set("include_stats","1");let r=await X("/api/training?"+i.toString());if(r.ok&&r.data&&Array.isArray(r.data.data)){let d=r.data.pagination||{};if(C=Number(d.total??r.data.data.length),T=Math.max(1,Number(d.pages||d.totalPages||Math.ceil(C/E)||1)),r.data.data.length===0&&c>T)return c=T,p=!1,k({refreshStats:$});t=r.data.data,r.data.stats&&(u={...u,...r.data.stats})}else t=[],C=0,T=1,z(r.data?.error||"Gagal memuat data training")}catch(m){console.warn("Gagal load training:",m),t=[],C=0,T=1,z("Gagal memuat data training: "+m.message)}finally{p=!1,R?($&&M(),O()):(R=!0,W())}}async function x(){if(!ie(K,"export")){z("Anda tidak memiliki izin untuk mengekspor data.");return}try{let{downloadExcel:$}=await Promise.resolve().then(()=>(we(),Oe)),m=await ue("/api/training/export",w()),i=m.map((r,d)=>({No:d+1,Tanggal:r.training_date||"",Batch:r.batch||"",Materi:r.subject||"",Cabang:r.branch_name||"",Trainer:r.trainer||"",Peserta:Ta(r.participants),Nilai:r.score!=null?r.score:"",Dokumen:r.document_link||"",Catatan:r.notes||""}));$(i,`Data_Training_${new Date().toISOString().slice(0,10)}`),ae(`Berhasil mengunduh ${m.length} data training ke Excel!`)}catch($){z("Gagal ekspor Excel: "+$.message)}}async function _($){if(!N)return;let m=t.find(i=>i.id===$);me({title:"Hapus Data Training",message:`Apakah Anda yakin ingin menghapus data sesi training "${m?.subject||"ini"}" untuk ${Ta(m?.participants)}?`,confirmText:"Ya, Hapus",confirmClass:"btn-danger",onConfirm:async()=>{try{let i=await X(`/api/training/${$}`,{method:"DELETE"});i.ok?(ae("Data training berhasil dihapus"),await k({refreshStats:!0})):z(i.data?.error||"Gagal menghapus data training")}catch(i){z("Gagal menghapus training: "+i.message)}}})}function B($=null){if(!I)return;let m=!!$,i=$?.training_date?$.training_date.slice(0,10):new Date().toISOString().slice(0,10),r=m&&$?.batch?String($.batch).trim():"",d=m&&r&&!Qt.includes(r),b=$?.trainer||"Berlin Ariansyah",y=document.createElement("div");y.innerHTML=`
      <div style="display:flex; flex-direction:column; gap:16px;">
        <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Tanggal Training <span style="color:#DC2626;">*</span>
            </label>
            <input type="date" id="tr-modal-date" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              value="${i}" required />
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Batch Pelatihan <span style="color:#DC2626;">*</span>
            </label>
            <select id="tr-modal-batch" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" required>
              <option value="">-- Pilih Batch Pelatihan --</option>
              ${Qt.map(v=>`<option value="${v}" ${r===v?"selected":""}>${v}</option>`).join("")}
              ${d?`<option value="${r.replace(/"/g,"&quot;")}" selected>${r} (Historis)</option>`:""}
            </select>
          </div>
        </div>

        <div>
          <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
            Materi / Topik Pelatihan <span style="color:#DC2626;">*</span>
          </label>
          <input type="text" id="tr-modal-subject" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
            placeholder="Contoh: Basic Knowledge & Practical Skills" value="${$?.subject?String($.subject).replace(/"/g,"&quot;"):""}" required />
        </div>

        <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Cabang Klinik <span style="color:#DC2626;">*</span>
            </label>
            <select id="tr-modal-branch" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" required>
              <option value="">-- Pilih Cabang --</option>
              ${n.map(v=>`
                <option value="${v.value}" ${String($?.branch_id)===String(v.value)?"selected":""}>
                  ${v.label}
                </option>
              `).join("")}
            </select>
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Trainer / Instruktur <span style="color:#DC2626;">*</span>
            </label>
            <input list="tr-trainer-options" id="tr-modal-trainer" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              placeholder="Pilih trainer..." value="${b}" required />
            <datalist id="tr-trainer-options">
              ${zn.map(v=>`<option value="${v}">`).join("")}
            </datalist>
          </div>
        </div>

        <div>
          <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
            Nama Peserta Pelatihan <span style="color:#DC2626;">*</span>
          </label>
          <div id="tr-modal-participants-mount"></div>
        </div>

        <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Nilai / Score Evaluasi (0 - 100)
            </label>
            <input type="number" min="0" max="100" step="1" id="tr-modal-score" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              placeholder="Contoh: 90" value="${$?.score!=null?$.score:""}" />
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Link Dokumen / Sertifikat (Opsional)
            </label>
            <input type="url" id="tr-modal-link" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              placeholder="https://drive.google.com/..." value="${$?.document_link?String($.document_link).replace(/"/g,"&quot;"):""}" />
          </div>
        </div>

        <div>
          <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
            Catatan Pelatihan
          </label>
          <textarea id="tr-modal-notes" class="form-control" rows="2" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
            placeholder="Catatan hasil pelatihan, rekomendasi materi lanjutan...">${$?.notes||""}</textarea>
        </div>
      </div>
    `;let q=ze(y.querySelector("#tr-modal-participants-mount"),{mode:"multi",inputId:"tr-modal-participants",inputName:"participants",placeholder:"Cari nama karyawan peserta (klik / Enter untuk tambah)...",allowManual:!0,hasEmployeeIdSchema:!1,employees:a,initialParticipants:va($?.participants)});pe({title:m?"\u270F\uFE0F Edit Data Sesi Training":"\u2795 Tambah Data Sesi Training",content:y,size:"md",confirmText:m?"Simpan Perubahan":"Simpan Training",confirmClass:"btn-primary",onConfirm:async(v,S)=>{let A=y.querySelector("#tr-modal-date").value,U=(y.querySelector("#tr-modal-batch").value||"").trim(),H=(y.querySelector("#tr-modal-subject").value||"").trim(),ee=y.querySelector("#tr-modal-branch").value,Q=(y.querySelector("#tr-modal-trainer").value||"").trim(),ne=(q.getValue().participantsString||"").trim(),Z=y.querySelector("#tr-modal-score").value,G=Z!==""?Number(Z):null,L=(y.querySelector("#tr-modal-link").value||"").trim(),j=(y.querySelector("#tr-modal-notes").value||"").trim();if(!A){z("Tanggal training wajib diisi");return}if(!U||!(m&&U===r)&&!Qt.includes(U)){z("Batch pelatihan wajib dipilih (Batch 1 - Batch 5)");return}if(!H){z("Materi / topik pelatihan wajib diisi");return}if(!ee){z("Cabang klinik wajib dipilih");return}if(!Q){z("Trainer wajib diisi");return}if(!ne){z("Nama peserta training wajib diisi");return}let V=n.find(re=>String(re.value)===String(ee)),te={training_date:A,batch:U,subject:H,branch_id:Number(ee),branch_name:V?V.label:"Cabang",trainer:Q,participants:ne,score:G,document_link:L,notes:j};try{if(m){let re=await X(`/api/training/${$.id}`,{method:"PUT",body:JSON.stringify(te)});re.ok?(ae("Data training berhasil diperbarui"),S(),await k({refreshStats:!0})):z(re.data?.error||"Gagal memperbarui data training")}else{let re=await X("/api/training",{method:"POST",body:JSON.stringify(te)});re.ok?(ae("Sesi training baru berhasil ditambahkan"),c=1,S(),await k({refreshStats:!0})):z(re.data?.error||"Gagal menambahkan training")}}catch(re){z("Terjadi kesalahan: "+re.message)}}})}function P(){return t}function f($){return`
    ${$.length===0?`
            <div class="training-empty-state">
              <div style="font-size:3rem; margin-bottom:12px;">\u{1F50D}</div>
              <h3 style="font-size:1.2rem; font-weight:700; color:#1E293B; margin:0 0 6px 0;">Tidak Ada Data Training Ditemukan</h3>
              <p style="color:#64748B; font-size:0.9rem; margin:0 0 20px 0;">Silakan sesuaikan kata kunci pencarian atau reset filter aktif.</p>
              <button id="tr-btn-reset-filter" class="btn btn-ghost" style="border:1px solid #CBD5E1; border-radius:8px; padding:8px 16px;">
                Reset Filter
              </button>
            </div>
          `:`
            <div class="training-cards-grid">
              ${$.map(m=>{let i=Ta(m.participants),d=i.split(",")[0].trim().split(" ").map(b=>b[0]).filter(Boolean).slice(0,2).join("").toUpperCase()||"TR";return`
                  <div class="training-card" data-id="${m.id}">
                    <div>
                      <!-- Card Header: Batch & Nilai/Tanggal -->
                      <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:12px;">
                        <span class="training-batch-badge">
                          <span>\u{1F393}</span>
                          <span>${m.batch||"Reguler"}</span>
                        </span>
                        <div style="display:flex; align-items:center; gap:6px;">
                          ${m.score!=null?`
                            <span class="training-score-badge">
                              <span>\u2B50</span>
                              <span>${m.score}</span>
                            </span>
                          `:""}
                          <span style="font-size:0.8rem; font-weight:600; color:#64748B;">
                            \u{1F5D3}\uFE0F ${vr(m.training_date)}
                          </span>
                        </div>
                      </div>

                      <!-- Subject / Materi Title -->
                      <h4 style="margin:0 0 10px 0; font-size:1.05rem; font-weight:700; color:#0F172A; line-height:1.4;">
                        ${m.subject||"Materi Pelatihan"}
                      </h4>

                      <!-- Participant & Branch Info -->
                      <div style="display:flex; align-items:center; gap:12px; margin-bottom:10px;">
                        <div class="training-avatar-box">
                          ${d}
                        </div>
                        <div style="min-width:0;">
                          <div style="font-size:0.72rem; text-transform:uppercase; font-weight:700; color:#94A3B8; margin-bottom:1px;">Peserta Training</div>
                          <div style="font-size:0.95rem; font-weight:700; color:#1E293B; white-space:nowrap; overflow:hidden; text-overflow:ellipsis;" title="${i}">
                            ${i}
                          </div>
                        </div>
                      </div>

                      <!-- Detail Box: Cabang & Trainer -->
                      <div class="training-info-box">
                        <div style="display:flex; align-items:center; gap:6px; font-size:0.82rem; color:#334155;">
                          <span>\u{1F4CD}</span>
                          <span style="font-weight:600;">${m.branch_name||"Cabang Belum Ditentukan"}</span>
                        </div>
                        <div style="display:flex; align-items:center; gap:6px; font-size:0.82rem; color:#64748B;">
                          <span>\u{1F468}\u200D\u{1F3EB}</span>
                          <span>Trainer: <strong style="color:#0F172A;">${m.trainer||"-"}</strong></span>
                        </div>
                      </div>

                      ${m.notes?`
                        <div style="font-size:0.78rem; color:#64748B; font-style:italic; line-height:1.4; margin-bottom:8px;">
                          \u{1F4DD} ${m.notes}
                        </div>
                      `:""}
                    </div>

                    <!-- Card Footer: Dokumen & Aksi -->
                    <div style="display:flex; justify-content:space-between; align-items:center; padding-top:12px; border-top:1px solid #F1F5F9; margin-top:8px;">
                      <div>
                        ${m.document_link?`
                          <a href="${m.document_link}" target="_blank" rel="noopener noreferrer" class="btn btn-sm btn-ghost" style="border:1px solid #CBD5E1; border-radius:6px; font-size:0.78rem; padding:4px 8px; text-decoration:none; display:inline-flex; align-items:center; gap:4px;">
                            <span>\u{1F4C4} Materi / Sertifikat</span>
                          </a>
                        `:`
                          <span style="font-size:0.75rem; color:#94A3B8;">(Tanpa Dokumen)</span>
                        `}
                      </div>

                      ${I||N?`
                        <div style="display:flex; align-items:center; gap:6px;">
                          ${I?`
                            <button class="training-icon-btn training-edit-btn" title="Edit Data Training" data-id="${m.id}">
                              <svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>
                            </button>
                          `:""}
                          ${N?`
                            <button class="training-icon-btn delete training-delete-btn" title="Hapus Training" data-id="${m.id}">
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
    `}function g(){let $=e.querySelector("#tr-btn-reset-filter");$&&($.onclick=()=>{h="",s="all",o="all",l="all",c=1,clearTimeout(F);let m=e.querySelector("#tr-search-input");m&&(m.value="");let i=e.querySelector("#tr-filter-branch");i&&(i.value="all");let r=e.querySelector("#tr-filter-trainer");r&&(r.value="all"),e.querySelectorAll(".training-pill").forEach(d=>{d.classList.toggle("active",d.getAttribute("data-batch")==="all")}),k({refreshStats:!1})}),e.querySelectorAll(".training-edit-btn").forEach(m=>{m.onclick=i=>{i.stopPropagation();let r=Number(m.getAttribute("data-id")),d=t.find(b=>b.id===r);d&&B(d)}}),e.querySelectorAll(".training-delete-btn").forEach(m=>{m.onclick=i=>{i.stopPropagation();let r=Number(m.getAttribute("data-id"));_(r)}})}function D(){let $=e.querySelector("#training-pagination-container");if(!$)return;$.innerHTML="";let m=_e({page:c,pages:T,total:C,limit:E,onPage:i=>{c=i,k({refreshStats:!1})}});m&&$.appendChild(m)}function O(){let $=e.querySelector("#training-list-container");if(!$)return;let m=P();$.innerHTML=f(m),D(),g()}function W(){let $=P(),m=Number(u.totalSessions??0),i=Number(u.totalParticipants??0),r=Number(u.totalBatches??0),d=Number(u.currentMonth??0);e.innerHTML=`
      <style>.training-page{display:flex;flex-direction:column;gap:20px}.training-hero-card{background:linear-gradient(135deg,#0F172A 0%,#1E293B 100%);border-radius:16px;padding:24px 28px;color:white;box-shadow:0 10px 25px -5px rgba(15,23,42,0.25);position:relative;overflow:hidden}.training-hero-card::after{content:'';position:absolute;top:-40px;right:-40px;width:220px;height:220px;background:radial-gradient(circle,rgba(147,51,234,0.18) 0%,transparent 70%);border-radius:50%;pointer-events:none}.training-hero-actions{display:flex;align-items:center;gap:8px;flex-wrap:wrap}.training-hero-actions .btn{height:38px;box-sizing:border-box;transition:all 0.2s ease}.training-btn-add:hover{background:#1D4ED8 !important;transform:translateY(-1px)}.training-btn-excel:hover,.training-btn-refresh:hover{background:rgba(255,255,255,0.16) !important;border-color:rgba(255,255,255,0.3) !important;transform:translateY(-1px)}.training-stats-row{display:grid;grid-template-columns:repeat(4,1fr);gap:14px;margin-top:22px}.training-stat-box{background:rgba(255,255,255,0.06);border:1px solid rgba(255,255,255,0.1);border-radius:12px;padding:14px 16px;backdrop-filter:blur(8px);display:flex;align-items:center;gap:14px}.training-stat-icon{width:44px;height:44px;border-radius:10px;display:flex;align-items:center;justify-content:center;font-size:1.3rem;flex-shrink:0}.training-stat-val{font-size:1.55rem;font-weight:800;line-height:1.1;color:#F8FAFC}.training-stat-lbl{font-size:0.76rem;text-transform:uppercase;letter-spacing:0.05em;color:#94A3B8;font-weight:600;margin-top:2px}.training-controls-card{background:#172033;border:1px solid #29364B;border-radius:14px;padding:18px 20px;box-shadow:0 2px 4px rgba(0,0,0,0.02);display:flex;flex-direction:column;gap:16px}.training-search-row{display:flex;gap:12px;flex-wrap:wrap}.training-search-wrap{position:relative;flex:1;min-width:240px}.training-search-icon{position:absolute;left:14px;top:50%;transform:translateY(-50%);color:#94A3B8;pointer-events:none}.training-search-input{width:100%;padding:10px 14px 10px 40px;border:1px solid #CBD5E1;border-radius:10px;font-size:0.9rem;outline:none;transition:border-color 0.15s ease}.training-search-input:focus{border-color:#2563EB;box-shadow:0 0 0 3px rgba(37,99,235,0.1)}.training-select{padding:10px 14px;border:1px solid #CBD5E1;border-radius:10px;font-size:0.9rem;color:#F8FAFC;background:#172033;outline:none;min-width:180px;cursor:pointer}.training-select:focus{border-color:#2563EB}.training-filter-pills{display:flex;align-items:center;gap:8px;overflow-x:auto;white-space:nowrap;scrollbar-width:none;padding-bottom:2px}.training-filter-pills::-webkit-scrollbar{display:none}.training-pill{padding:6px 14px;border-radius:20px;font-size:0.82rem;font-weight:600;cursor:pointer;border:1px solid #29364B;background:#1B2940;color:#64748B;transition:all 0.15s ease;display:inline-flex;align-items:center;gap:6px;user-select:none;flex-shrink:0}.training-pill:hover{background:#111827;color:#F8FAFC}.training-pill.active{background:#2563EB;color:white;border-color:#2563EB;box-shadow:0 2px 6px rgba(37,99,235,0.25)}.training-cards-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(320px,1fr));gap:16px}.training-card{background:#172033;border:1px solid #29364B;border-radius:14px;padding:18px;box-shadow:0 2px 6px rgba(0,0,0,0.03);transition:transform 0.15s ease,box-shadow 0.15s ease;display:flex;flex-direction:column;justify-content:space-between;position:relative}.training-card:hover{transform:translateY(-2px);box-shadow:0 8px 18px rgba(0,0,0,0.06);border-color:#CBD5E1}.training-avatar-box{width:44px;height:44px;border-radius:10px;background:linear-gradient(135deg,#7C3AED 0%,#6D28D9 100%);color:white;font-weight:700;font-size:1rem;display:flex;align-items:center;justify-content:center;flex-shrink:0;box-shadow:0 3px 6px rgba(124,58,237,0.25)}.training-batch-badge{display:inline-flex;align-items:center;gap:5px;padding:4px 10px;border-radius:20px;font-size:0.76rem;font-weight:700;background:#F3E8FF;color:#7E22CE;border:1px solid #E9D5FF}.training-score-badge{display:inline-flex;align-items:center;gap:4px;padding:4px 10px;border-radius:20px;font-size:0.76rem;font-weight:700;background:#DCFCE7;color:#15803D;border:1px solid #BBF7D0}.training-info-box{background:#1B2940;border:1px solid #29364B;border-radius:10px;padding:10px 12px;margin:12px 0;display:flex;flex-direction:column;gap:6px}.training-icon-btn{width:32px;height:32px;border-radius:8px;border:1px solid #29364B;background:#172033;color:#64748B;display:inline-flex;align-items:center;justify-content:center;cursor:pointer;transition:all 0.15s ease}.training-icon-btn:hover{background:#111827;color:#0F172A;border-color:#CBD5E1}.training-icon-btn.delete:hover{background:#FEE2E2;color:#DC2626;border-color:#FCA5A5}.training-empty-state{grid-column:1 / -1;background:#172033;border:2px dashed #29364B;border-radius:16px;padding:48px 24px;text-align:center}@media (max-width:768px){.training-hero-card{padding:18px 16px}.training-stats-row{grid-template-columns:1fr 1fr;gap:10px}.training-stat-box{padding:10px 12px}.training-stat-icon{width:38px;height:38px;font-size:1.1rem}.training-stat-val{font-size:1.3rem}.training-hero-actions{width:100% !important;display:flex !important;align-items:center !important;gap:6px !important;flex-wrap:wrap !important;overflow-x:visible !important;scrollbar-width:none !important;padding-bottom:2px !important}.training-hero-actions::-webkit-scrollbar{display:none}.training-btn-add{flex:1 1 auto !important;justify-content:center !important;padding:8px 10px !important;font-size:0.82rem !important;white-space:nowrap !important}.training-btn-excel{flex:1 1 auto !important;justify-content:center !important;padding:8px 10px !important;font-size:0.82rem !important;white-space:nowrap !important}.training-btn-refresh{width:38px !important;height:38px !important;padding:0 !important;flex-shrink:0 !important;justify-content:center !important}.training-controls-card{padding:14px;gap:12px}.training-search-row{flex-direction:column;gap:10px}.training-select{width:100%}.training-cards-grid{grid-template-columns:minmax(0,1fr) !important;width:100% !important;max-width:100% !important;min-width:0 !important;box-sizing:border-box !important}.training-card{width:100% !important;max-width:100% !important;min-width:0 !important;box-sizing:border-box !important;overflow:hidden;padding:14px !important}}</style>

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
              ${I?`
                <button id="tr-btn-add" class="btn btn-primary training-btn-add" style="background:#2563EB; height:38px; padding:0 16px; border-radius:10px; font-weight:700; display:inline-flex; align-items:center; gap:6px; box-shadow:0 4px 12px rgba(37,99,235,0.3); white-space:nowrap;">
                  <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24"><line x1="12" y1="5" x2="12" y2="19"/><line x1="5" y1="12" x2="19" y2="12"/></svg>
                  <span>Tambah Training</span>
                </button>
              `:""}
              
              ${ie(K,"export")?`<button id="tr-btn-excel" class="btn btn-ghost training-btn-excel" style="background:rgba(255,255,255,0.08); border:1px solid rgba(255,255,255,0.15); color:white; height:38px; padding:0 14px; border-radius:10px; font-weight:600; display:inline-flex; align-items:center; gap:6px; white-space:nowrap;">
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
                <div class="training-stat-val" id="tr-stat-sessions">${m}</div>
                <div class="training-stat-lbl">Sesi Training</div>
              </div>
            </div>

            <div class="training-stat-box">
              <div class="training-stat-icon" style="background:rgba(59,130,246,0.2); color:#60A5FA;">
                \u{1F465}
              </div>
              <div>
                <div class="training-stat-val" id="tr-stat-participants">${i}</div>
                <div class="training-stat-lbl">Peserta Terlatih</div>
              </div>
            </div>

            <div class="training-stat-box">
              <div class="training-stat-icon" style="background:rgba(16,185,129,0.2); color:#34D399;">
                \u{1F4DA}
              </div>
              <div>
                <div class="training-stat-val"><span id="tr-stat-batches">${r}</span> Batch</div>
                <div class="training-stat-lbl">Total Angkatan</div>
              </div>
            </div>

            <div class="training-stat-box">
              <div class="training-stat-icon" style="background:rgba(245,158,11,0.2); color:#FBBF24;">
                \u{1F5D3}\uFE0F
              </div>
              <div>
                <div class="training-stat-val" id="tr-stat-month">${d}</div>
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
              <input type="text" id="tr-search-input" class="training-search-input" placeholder="Cari materi, trainer, peserta, atau cabang..." value="${h?String(h).replace(/"/g,"&quot;"):""}" />
            </div>

            <select id="tr-filter-branch" class="training-select">
              <option value="all">Semua Cabang (${n.length})</option>
              ${n.map(b=>`<option value="${b.value}" ${String(s)===String(b.value)?"selected":""}>${b.label}</option>`).join("")}
            </select>

            <select id="tr-filter-trainer" class="training-select">
              <option value="all">Semua Trainer</option>
              ${zn.map(b=>`<option value="${b}" ${l===b?"selected":""}>${b}</option>`).join("")}
            </select>
          </div>

          <!-- Batch Filter Pills (Swipeable on Mobile) -->
          <div class="training-filter-pills">
            <div class="training-pill ${o==="all"?"active":""}" data-batch="all">
              <span>Semua Batch</span>
              <span class="badge" id="tr-pill-count-all" style="background:rgba(0,0,0,0.06); padding:2px 6px; border-radius:10px; font-size:0.75rem;">${m}</span>
            </div>
            ${Qt.map(b=>`
              <div class="training-pill ${o===b?"active":""}" data-batch="${b}">
                <span>\u{1F393}</span>
                <span>${b}</span>
              </div>
            `).join("")}
          </div>
        </div>

        <!-- Content Area: Murni Grid Kartu -->
        <div id="training-list-container">
          ${f($)}
        </div>

        <!-- Pagination Container -->
        <div id="training-pagination-container"></div>

      </div>
    `,D(),J()}function J(){let $=e.querySelector("#tr-btn-add");$&&($.onclick=()=>B());let m=e.querySelector("#tr-btn-excel");m&&(m.onclick=()=>x());let i=e.querySelector("#tr-btn-refresh");i&&(i.onclick=()=>k({refreshStats:!0}));let r=e.querySelector("#tr-search-input");r&&(r.oninput=y=>{h=y.target.value,clearTimeout(F),F=setTimeout(()=>{c=1,k({refreshStats:!1})},350)});let d=e.querySelector("#tr-filter-branch");d&&(d.onchange=y=>{s=y.target.value,c=1,k({refreshStats:!1})});let b=e.querySelector("#tr-filter-trainer");b&&(b.onchange=y=>{l=y.target.value,c=1,k({refreshStats:!1})}),e.querySelectorAll(".training-pill").forEach(y=>{y.onclick=()=>{o=y.getAttribute("data-batch")||"all",e.querySelectorAll(".training-pill").forEach(q=>{q.classList.toggle("active",q.getAttribute("data-batch")===o)}),c=1,k({refreshStats:!1})}}),g()}await k({refreshStats:!0})}we();async function Nn(e){let t=await ve(),n=Array.from({length:4},(a,o)=>String(new Date().getFullYear()-o));De({container:e,title:"Laporan Inspeksi Hygiene",icon:"\u{1F50D}",apiPath:"/api/reports/inspection",enableMobileFilterSheet:!0,itemLabel:"Laporan Inspeksi",bulkDelete:!1,columns:[{key:"branch_name",label:"Cabang"},{key:"period",label:"Periode",render:a=>We(a)},{key:"inspection_date",label:"Tanggal",nowrap:!0,render:a=>window.formatDate(a)},{key:"fc_score",label:"Point FC",render:a=>a!=null?`<strong class="${a>=80?"text-success":a>=60?"text-warning":"text-danger"}">${a}</strong>`:"-"},{key:"spv_score",label:"Point SPV",render:a=>a!=null?`<strong>${a}</strong>`:"-"},{key:"status",label:"Status",render:a=>He(a)},{key:"document_link",label:"Dokumen",render:a=>a?`<a href="${a}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari cabang..."},{type:"select",name:"branch_id",label:"Cabang",options:t},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"month",label:"Bulan",options:[{value:"01",label:"Jan"},{value:"02",label:"Feb"},{value:"03",label:"Mar"},{value:"04",label:"Apr"},{value:"05",label:"Mei"},{value:"06",label:"Jun"},{value:"07",label:"Jul"},{value:"08",label:"Agu"},{value:"09",label:"Sep"},{value:"10",label:"Okt"},{value:"11",label:"Nov"},{value:"12",label:"Des"}]},{type:"select",name:"status",label:"Status",options:["Pending","Done"]},{type:"select",name:"year",label:"Tahun",options:n}],formFields:a=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"select",required:!0,options:t,value:a?.branch_id},{name:"period",label:"Periode",type:"select",required:!0,options:["Q1","Q2","Q3","Q4"],value:a?.period}]},{type:"row",fields:[{name:"inspection_date",label:"Tanggal Inspeksi",type:"date",required:!0,value:a?.inspection_date},{name:"status",label:"Status",type:"select",required:!0,options:["Pending","Done"],value:a?.status||""}]},{type:"row",fields:[{name:"fc_score",label:"Point FC",type:"number",step:"0.1",min:"0",max:"100",value:a?.fc_score},{name:"spv_score",label:"Point SPV",type:"number",step:"0.1",min:"0",max:"100",value:a?.spv_score}]},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://drive.google.com/...",value:a?.document_link},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:a?.notes}],exportOptions:{moduleName:"inspection_reports",onExport:async a=>{let s=(await ue("/api/reports/inspection/export",a)).map(l=>({Cabang:l.branch_name||"",Periode:l.period||"",Tanggal:l.inspection_date||"","Point FC":l.fc_score!==null&&l.fc_score!==void 0?l.fc_score:"","Point SPV":l.spv_score!==null&&l.spv_score!==void 0?l.spv_score:"",Status:l.status||"","Link Dokumen":l.document_link||""}));ge(s,`Laporan_Inspeksi_Hygiene_${new Date().toISOString().slice(0,10)}`)},onTemplate:()=>{ge([{Cabang:"001. Pondok Bambu",Periode:"Q1",Tanggal:"2026-01-08","Point FC":85,"Point SPV":90,Status:"Done","Link Dokumen":"https://drive.google.com/...",Catatan:""}],"Template_Import_Inspeksi")},onImport:async(a,o={})=>{let s=p=>{if(!p)return null;let c=String(p||"").toLowerCase().trim(),E=c.match(/^(\d{3})/);if(E){let T=t.find(u=>String(u.code||"").trim()===E[1]||String(u.label||"").toLowerCase().startsWith(E[1]+"."));if(T)return T.value}let C=t.find(T=>String(T.label||"").toLowerCase()===c||String(T.code||"").toLowerCase()===c||String(T.name||"").toLowerCase()===c);return C?C.value:null},l=p=>{if(p==null||p==="")return"";if(p instanceof Date&&!isNaN(p.getTime()))return p.toISOString().slice(0,10);let c=String(p).trim();if(c===""||c==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(c))return c.slice(0,10);if(/^\d{4,5}$/.test(c)){let C=Number(c);if(C>2e4&&C<99999){let T=new Date(Date.UTC(1899,11,30)+C*864e5);return isNaN(T.getTime())?"":T.toISOString().slice(0,10)}}let E=c.split(/[\/\-\.]/);if(E.length===3){let[C,T,u]=E.map(F=>F.trim());if(C.length===4&&T.length<=2&&u.length<=2)return`${C}-${T.padStart(2,"0")}-${u.padStart(2,"0")}`;if(u.length===4&&T.length<=2&&C.length<=2)return`${u}-${T.padStart(2,"0")}-${C.padStart(2,"0")}`}return c},h=a.map(p=>({branch_id:s(String(p.Cabang||"").trim()),period:String(p.Periode||"").trim(),inspection_date:l(p.Tanggal),fc_score:p["Point FC"]!==void 0&&p["Point FC"]!==""?Number(p["Point FC"]):null,spv_score:p["Point SPV"]!==void 0&&p["Point SPV"]!==""?Number(p["Point SPV"]):null,status:String(p.Status||"").trim(),document_link:String(p["Link Dokumen"]||"").trim(),notes:String(p.Catatan||p.Keterangan||"").trim()})).filter(p=>p.branch_id&&p.period&&p.inspection_date);return await Le({module:"inspection_reports",label:"Laporan Inspeksi",endpoint:"/api/import/inspection",body:{rows:h,onDuplicate:"update"},idempotencyKey:o?.idempotencyKey,recover:!!o?.recover})}}})}we();async function Rn(e){let t=await ve(),n=Array.from({length:4},(a,o)=>String(new Date().getFullYear()-o));De({container:e,title:"Laporan GC & DC",icon:"\u{1F9F9}",apiPath:"/api/reports/cleaning",itemLabel:"Laporan GC/DC",bulkDelete:!1,enableMobileFilterSheet:!0,columns:[{key:"branch_name",label:"Cabang"},{key:"activity_type",label:"Jenis",render:a=>`<span class="badge ${a==="Deep Cleaning"?"badge-purple":"badge-success"}">${a}</span>`},{key:"period",label:"Periode",render:a=>We(a)},{key:"activity_date",label:"Tanggal",nowrap:!0,render:a=>window.formatDate(a)},{key:"status",label:"Status",render:a=>He(a)},{key:"document_link",label:"Dokumen",render:a=>a?`<a href="${a}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari nama cabang/lokasi..."},{type:"combobox",name:"branch_id",label:"Cabang",options:t},{type:"select",name:"activity_type",label:"Jenis",options:["General Cleaning","Deep Cleaning"]},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"month",label:"Bulan",options:[{value:"01",label:"Jan"},{value:"02",label:"Feb"},{value:"03",label:"Mar"},{value:"04",label:"Apr"},{value:"05",label:"Mei"},{value:"06",label:"Jun"},{value:"07",label:"Jul"},{value:"08",label:"Agu"},{value:"09",label:"Sep"},{value:"10",label:"Okt"},{value:"11",label:"Nov"},{value:"12",label:"Des"}]},{type:"select",name:"status",label:"Status",options:["Pending","Done"]},{type:"select",name:"year",label:"Tahun",options:n}],formFields:a=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:t,value:a?.branch_id},{name:"activity_type",label:"Jenis Kegiatan",type:"select",required:!0,options:["General Cleaning","Deep Cleaning"],value:a?.activity_type}]},{type:"row",fields:[{name:"period",label:"Periode",type:"select",required:!0,options:["Q1","Q2","Q3","Q4"],value:a?.period},{name:"activity_date",label:"Tanggal",type:"date",required:!0,value:a?.activity_date}]},{name:"status",label:"Status",type:"select",required:!0,options:["Pending","Done"],value:a?.status||""},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://drive.google.com/...",value:a?.document_link},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:a?.notes}],exportOptions:{moduleName:"cleaning_reports",onExport:async a=>{let s=(await ue("/api/reports/cleaning/export",a)).map(l=>({Cabang:l.branch_name||"",Jenis:l.activity_type||"",Periode:l.period||"",Tanggal:l.activity_date||"",Status:l.status||"","Link Dokumen":l.document_link||""}));ge(s,`Laporan_GCDC_${new Date().toISOString().slice(0,10)}`)},onTemplate:()=>{ge([{Cabang:"001. Pondok Bambu",Jenis:"General Cleaning",Periode:"Q1",Tanggal:"2026-01-08",Status:"Done","Link Dokumen":"https://drive.google.com/...",Catatan:""}],"Template_Import_GCDC")},onImport:async(a,o={})=>{let s=p=>{if(!p)return null;let c=String(p||"").toLowerCase().trim(),E=c.match(/^(\d{3})/);if(E){let T=t.find(u=>String(u.code||"").trim()===E[1]||String(u.label||"").toLowerCase().startsWith(E[1]+"."));if(T)return T.value}let C=t.find(T=>String(T.label||"").toLowerCase()===c||String(T.code||"").toLowerCase()===c||String(T.name||"").toLowerCase()===c);return C?C.value:null},l=p=>{if(p==null||p==="")return"";if(p instanceof Date&&!isNaN(p.getTime()))return p.toISOString().slice(0,10);let c=String(p).trim();if(c===""||c==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(c))return c.slice(0,10);if(/^\d{4,5}$/.test(c)){let C=Number(c);if(C>2e4&&C<99999){let T=new Date(Date.UTC(1899,11,30)+C*864e5);return isNaN(T.getTime())?"":T.toISOString().slice(0,10)}}let E=c.split(/[\/\-\.]/);if(E.length===3){let[C,T,u]=E.map(F=>F.trim());if(C.length===4&&T.length<=2&&u.length<=2)return`${C}-${T.padStart(2,"0")}-${u.padStart(2,"0")}`;if(u.length===4&&T.length<=2&&C.length<=2)return`${u}-${T.padStart(2,"0")}-${C.padStart(2,"0")}`}return c},h=a.map(p=>({branch_id:s(String(p.Cabang||"").trim()),activity_type:String(p.Jenis||p.Kegiatan||"").trim(),period:String(p.Periode||"").trim(),activity_date:l(p.Tanggal),status:String(p.Status||"").trim(),document_link:String(p["Link Dokumen"]||"").trim(),notes:String(p.Catatan||p.Keterangan||"").trim()})).filter(p=>p.branch_id&&p.activity_type&&p.period&&p.activity_date);return await Le({module:"cleaning_reports",label:"Laporan GC/DC",endpoint:"/api/import/cleaning",body:{rows:h,onDuplicate:"update"},idempotencyKey:o?.idempotencyKey,recover:!!o?.recover})}}})}we();async function jn(e,t){let n=await ve(),a=Array.from({length:4},(p,c)=>String(new Date().getFullYear()-c)),o=t?t.get("dash_filter"):null,s=t?t.get("explicit")==="1":!1,l=t?t.get("month"):null,h={};if(o==="fogging")if(s&&l&&l.length===7){let[p,c]=l.split("-");h={month:c,year:p}}else h={year:l&&l.length===7?l.split("-")[0]:String(new Date().getFullYear())};De({container:e,title:"Rekap Fogging",icon:"\u{1F4A8}",apiPath:"/api/reports/fogging",itemLabel:"Fogging",bulkDelete:!1,enableMobileFilterSheet:!0,defaultFilters:h,columns:[{key:"branch_name",label:"Cabang"},{key:"activity_type",label:"Jenis",render:p=>`<span class="badge badge-warning">${p}</span>`},{key:"period",label:"Periode",render:p=>We(p)},{key:"activity_date",label:"Tanggal",nowrap:!0,render:p=>window.formatDate(p)},{key:"status",label:"Status",render:p=>He(p)},{key:"document_link",label:"Dokumen",render:p=>p?`<a href="${p}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"},{key:"notes",label:"Catatan",render:p=>p||"-"}],filterFields:[{type:"search",placeholder:"Cari nama cabang/lokasi..."},{type:"select",name:"branch_id",label:"Cabang",options:n},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"month",label:"Bulan",options:[{value:"01",label:"Jan"},{value:"02",label:"Feb"},{value:"03",label:"Mar"},{value:"04",label:"Apr"},{value:"05",label:"Mei"},{value:"06",label:"Jun"},{value:"07",label:"Jul"},{value:"08",label:"Agu"},{value:"09",label:"Sep"},{value:"10",label:"Okt"},{value:"11",label:"Nov"},{value:"12",label:"Des"}]},{type:"select",name:"status",label:"Status",options:["Pending","Done"]},{type:"select",name:"year",label:"Tahun",options:a}],formFields:p=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"select",required:!0,options:n,value:p?.branch_id},{name:"period",label:"Periode",type:"select",required:!0,options:["Q1","Q2","Q3","Q4"],value:p?.period}]},{type:"row",fields:[{name:"activity_date",label:"Tanggal",type:"date",value:p?.activity_date},{name:"status",label:"Status",type:"select",required:!0,options:["Pending","Done"],value:p?.status||""}]},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://...",value:p?.document_link},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:p?.notes}],exportOptions:{moduleName:"fogging_reports",onExport:async p=>{let E=(await ue("/api/reports/fogging/export",p)).map(C=>({Cabang:C.branch_name||"",Jenis:C.activity_type||"Fogging",Periode:C.period||"",Tanggal:C.activity_date||"",Status:C.status||"","Link Dokumen":C.document_link||""}));ge(E,`Laporan_Fogging_${new Date().toISOString().slice(0,10)}`)},onTemplate:()=>{ge([{Cabang:"001. Pondok Bambu",Jenis:"Fogging",Periode:"Q1",Tanggal:"2026-01-08",Status:"Done","Link Dokumen":"https://drive.google.com/..."}],"Template_Import_Fogging")},onImport:async(p,c={})=>{let E=u=>{if(!u)return null;let F=String(u||"").toLowerCase().trim(),R=F.match(/^(\d{3})/);if(R){let I=n.find(N=>String(N.code||"").trim()===R[1]||String(N.label||"").toLowerCase().startsWith(R[1]+"."));if(I)return I.value}let K=n.find(I=>String(I.label||"").toLowerCase()===F||String(I.code||"").toLowerCase()===F||String(I.name||"").toLowerCase()===F);return K?K.value:null},C=u=>{if(u==null||u==="")return"";if(u instanceof Date&&!isNaN(u.getTime()))return u.toISOString().slice(0,10);let F=String(u).trim();if(F===""||F==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(F))return F.slice(0,10);if(/^\d{4,5}$/.test(F)){let K=Number(F);if(K>2e4&&K<99999){let I=new Date(Date.UTC(1899,11,30)+K*864e5);return isNaN(I.getTime())?"":I.toISOString().slice(0,10)}}let R=F.split(/[\/\-\.]/);if(R.length===3){let[K,I,N]=R.map(w=>w.trim());if(K.length===4&&I.length<=2&&N.length<=2)return`${K}-${I.padStart(2,"0")}-${N.padStart(2,"0")}`;if(N.length===4&&I.length<=2&&K.length<=2)return`${N}-${I.padStart(2,"0")}-${K.padStart(2,"0")}`}return F},T=p.map(u=>({branch_id:E(String(u.Cabang||"").trim()),activity_type:String(u.Jenis||u.Kegiatan||"Fogging").trim(),period:String(u.Periode||"").trim(),activity_date:C(u.Tanggal),status:String(u.Status||"").trim(),document_link:String(u["Link Dokumen"]||"").trim(),notes:String(u.Catatan||u.Keterangan||"").trim()})).filter(u=>u.branch_id&&u.period&&u.activity_date);return await Le({module:"fogging_reports",label:"Laporan Fogging",endpoint:"/api/reports/fogging/import",body:T,idempotencyKey:c?.idempotencyKey,recover:!!c?.recover})}}})}we();async function qn(e){let t=await ve(),n=await Gt(),a=[{value:"Berlin",label:"Berlin"},{value:"Ade",label:"Ade"}],o=Array.from({length:4},(h,p)=>String(new Date().getFullYear()-p)),s=h=>h&&!n.find(p=>p.value===h)?[...n,{value:h,label:h}]:n,l=h=>h&&!a.find(p=>p.value===h)?[...a,{value:h,label:h}]:a;De({container:e,title:"Rekap Laporan Basecamp",icon:"\u{1F4DD}",apiPath:"/api/reports/basecamp",bulkDelete:!1,enableMobileFilterSheet:!0,itemLabel:"Laporan Basecamp",columns:[{key:"info_date",label:"Tgl Info",nowrap:!0,render:h=>window.formatDate(h)},{key:"branch_name",label:"Cabang"},{key:"problem",label:"Permasalahan",render:h=>`<span title="${h||""}">${h?.length>60?h.slice(0,60)+"\u2026":h||"-"}</span>`},{key:"pic",label:"PIC"},{key:"done_date",label:"Tgl Done",nowrap:!0,render:h=>window.formatDate(h)},{key:"status",label:"Status",render:h=>He(h)},{key:"notes",label:"Keterangan",render:h=>h?.length>40?h.slice(0,40)+"\u2026":h||"-"}],filterFields:[{type:"search",placeholder:"Cari permasalahan, PIC, cabang..."},{type:"select",name:"pic",label:"PIC",options:["Berlin","Ade"]},{type:"select",name:"branch_id",label:"Cabang",options:t},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"month",label:"Bulan",options:[{value:"01",label:"Jan"},{value:"02",label:"Feb"},{value:"03",label:"Mar"},{value:"04",label:"Apr"},{value:"05",label:"Mei"},{value:"06",label:"Jun"},{value:"07",label:"Jul"},{value:"08",label:"Agu"},{value:"09",label:"Sep"},{value:"10",label:"Okt"},{value:"11",label:"Nov"},{value:"12",label:"Des"}]},{type:"select",name:"status",label:"Status",options:["Open","In Progress","Done"]},{type:"select",name:"year",label:"Tahun",options:o}],formFields:h=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"select",required:!0,options:t,value:h?.branch_id},{name:"pic",label:"PIC",type:"select",options:l(h?.pic),value:h?.pic}]},{name:"problem",label:"Permasalahan",type:"textarea",required:!0,rows:3,value:h?.problem},{type:"row",fields:[{name:"info_date",label:"Tanggal Info",type:"date",required:!0,value:h?.info_date},{name:"done_date",label:"Tanggal Done",type:"date",value:h?.done_date}]},{name:"status",label:"Status",type:"select",options:["Open","In Progress","Done"],value:h?.status||"Open"},{name:"notes",label:"Keterangan / Tindak Lanjut",type:"textarea",rows:2,value:h?.notes}],exportOptions:{moduleName:"basecamp_reports",onExport:async h=>{let c=(await ue("/api/reports/basecamp/export",h)).map(E=>({"Tgl Info":E.info_date||"",Cabang:E.branch_name||"",Permasalahan:E.problem||"",PIC:E.pic||"","Tgl Done":E.done_date||"",Status:E.status||"",Keterangan:E.notes||""}));ge(c,`Rekap_Laporan_Basecamp_${new Date().toISOString().slice(0,10)}`)},onTemplate:()=>{ge([{"Tgl Info":"2026-01-08",Cabang:"001. Pondok Bambu",Permasalahan:"Request fogging karena banyak nyamuk",PIC:"Fajar","Tgl Done":"2026-01-10",Status:"Done",Keterangan:"Sudah difogging"}],"Template_Import_Basecamp")},onImport:async(h,p={})=>{let c=T=>{if(!T)return null;let u=String(T||"").toLowerCase().trim(),F=u.match(/^(\d{3})/);if(F){let K=t.find(I=>String(I.code||"").trim()===F[1]||String(I.label||"").toLowerCase().startsWith(F[1]+"."));if(K)return K.value}let R=t.find(K=>String(K.label||"").toLowerCase()===u||String(K.code||"").toLowerCase()===u||String(K.name||"").toLowerCase()===u);return R?R.value:null},E=T=>{if(T==null||T==="")return"";if(T instanceof Date&&!isNaN(T.getTime()))return T.toISOString().slice(0,10);let u=String(T).trim();if(u===""||u==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(u))return u.slice(0,10);if(/^\d{4,5}$/.test(u)){let R=Number(u);if(R>2e4&&R<99999){let K=new Date(Date.UTC(1899,11,30)+R*864e5);return isNaN(K.getTime())?"":K.toISOString().slice(0,10)}}let F=u.split(/[\/\-\.]/);if(F.length===3){let[R,K,I]=F.map(N=>N.trim());if(R.length===4&&K.length<=2&&I.length<=2)return`${R}-${K.padStart(2,"0")}-${I.padStart(2,"0")}`;if(I.length===4&&K.length<=2&&R.length<=2)return`${I}-${K.padStart(2,"0")}-${R.padStart(2,"0")}`}return u},C=h.map(T=>({info_date:E(T["Tgl Info"]||T["Tanggal Info"]),branch_id:c(String(T.Cabang||"").trim()),problem:String(T.Permasalahan||"").trim(),pic:String(T.PIC||"").trim(),done_date:E(T["Tgl Done"]||T["Tanggal Done"]),status:String(T.Status||"").trim(),notes:String(T.Keterangan||T.Catatan||"").trim()})).filter(T=>T.info_date&&T.branch_id&&T.problem);return await Le({module:"basecamp_reports",label:"Laporan Basecamp",endpoint:"/api/reports/basecamp/import",body:C,idempotencyKey:p?.idempotencyKey,recover:!!p?.recover})}}})}Fe();var kr=[{id:1,title:"Dusting (Pembersihan Debu)",category:"Pembersihan Rutin",icon:"\u{1F9F9}",color:"#2563EB",bg:"#EFF6FF",target:"Meja resepsionis, meja pasien, meja dokter, furniture, kaca, monitor & keyboard",chemical_ratio:"30 ml Glass Cleaner : 1 Liter Air / Siap pakai di botol spray",tools:["Kain Lap Microfiber Biru (Area Umum)","Kanebo (Kaca & Layar Monitor)","Kemoceng","Botol Spray Glass Cleaner","Kuas 4 Inch (Keyboard & Sela Sempit)"],steps:["Siapkan larutan glass cleaner di botol spray khusus kaca dan perabot.","Lipat kain microfiber menjadi 8 bidang kerja agar efisien dan bersih merata.","Semprotkan cairan pembersih ke permukaan kain lap, bukan langsung ke monitor atau bodi elektronik.","Usap permukaan perabot dari bagian paling atas ke bagian bawah secara searah atau pola S.","Gunakan kuas cat 4 inch untuk membersihkan sela-sela keyboard dan lubang ventilasi.","Keringkan kaca atau monitor menggunakan lap microfiber kering atau kanebo bersih."],notes:"Khusus layar monitor medis & dental unit, dilarang menyemprotkan cairan langsung ke layar."},{id:2,title:"Sweeping (Lobby Duster & Sapu Nylon)",category:"Pembersihan Rutin",icon:"\u{1F33E}",color:"#D97706",bg:"#FFFBEB",target:"Lobby utama, lorong, ruang tindakan, balkon, pantry, teras",chemical_ratio:"-",tools:["Lobby Duster Set 60cm (Area Lobby & Lorong)","Lobby Duster Set 40cm (Area Ruang Tindakan)","Sapu & Pengki Nagata 742s","Kapi Scrapper (Kerak/Permen Karet)"],steps:["Pastikan kain lobby duster dalam keadaan kering, bersih, dan bebas pasir kasar.","Dorong lobby duster menyusuri tepi dinding ruangan terlebih dahulu membentuk perimeter.","Lanjutkan mendorong dari sudut terjauh menuju arah pintu keluar tanpa mengangkat kain pel lobi dari lantai.","Kumpulkan debu dan kotoran di dekat pintu keluar, lalu angkat menggunakan sapu dan pengki Nagata.","Jika ada noda kerak atau bekas stiker/permen karet menempel di lantai, kikis perlahan menggunakan kapi scrapper.","Kibaskan kain duster di luar ruangan atau sedot sisa debu pada kain duster."],notes:"Jangan pernah menggunakan lobby duster di area lantai yang masih basah atau tergenang air."},{id:3,title:"Mopping (Pengepelan Lantai)",category:"Pembersihan Rutin",icon:"\u{1FAA3}",color:"#059669",bg:"#ECFDF5",target:"Seluruh lantai ruangan, koridor, dan area tunggu klinik",chemical_ratio:"25 ml Floor Cleaner per 1 Liter air bersih",tools:["Single Bucket Mop Wringer 20L / Double Bucket 46L","Mopping Stick / Tongkat & Kain Pel","Chemical Floor Cleaner Prokleen Lavender","Warning Sign Lantai Licin"],steps:['Pasang papan peringatan "Caution Wet Floor" di titik strategis yang mudah terlihat.',"Campurkan cairan Floor Cleaner ke dalam ember air bersih dengan takaran 25 ml per 1 Liter air.","Celupkan kain pel, masukkan ke dalam alat pemeras (wringer), dan peras hingga kondisi kain pel lembap (tidak menetes).","Mulai mengepel dari sudut terjauh ruangan mundur ke arah pintu keluar dengan gerakan membentuk pola angka delapan (8).","Bilas kain pel secara berkala apabila air pada ember bilas sudah terlihat keruh.","Biarkan lantai mengering sempurna secara alami sebelum papan peringatan wet floor dilepas."],notes:'Wajib pasang papan peringatan "Caution Wet Floor" sebelum mulai mengepel untuk mencegah risiko terpeleset.'},{id:4,title:"Membersihkan Toilet & Wastafel",category:"Sanitasi Toilet",icon:"\u{1F6BD}",color:"#DC2626",bg:"#FEF2F2",target:"Toilet pasien, toilet tim, closet, wastafel, dinding keramik, floor drain",chemical_ratio:"Karbol 25 ml:1 L Air / Bowl Cleaner 100 ml spray (kerak) / Bubuk KIFA",tools:["Sarung Tangan Karet (Wajib)","Sikat Kloset Scotch-Brite 3M","Kain Lap Microfiber Merah (Khusus Toilet)","Sponge Scotch Brite & Sikat Lantai ID-51","Floor Squeegee 75cm","Chemical Karbol Pine, Bowl Cleaner, Bubuk KIFA"],steps:["Kenakan sarung tangan karet dan siapkan perlengkapan sanitasi toilet.","Flush kloset, tuangkan Toilet Bowl Cleaner ke bagian dalam dinding kloset dan leher angsa, diamkan 5\u201310 menit.","Sikat bagian dalam kloset secara menyeluruh dengan sikat kloset, lalu bilas (flush) kembali hingga bersih.","Seka dudukan, penutup, dan tabung luar kloset menggunakan kain lap microfiber merah yang dibasahi larutan disinfektan karbol.","Bersihkan wastafel dan kran stainless menggunakan sponge tapas dengan bubuk pembersih KIFA (jangan gunakan Bowl Cleaner pada kran stainless).","Sikat lantai kamar mandi dengan sikat lantai ID-51 dan cairan karbol disinfektan.","Tarik genangan air ke arah floor drain menggunakan floor squeegee hingga lantai tidak licin dan cepat kering.","Pasang kamper/urinal block, isi ulang sabun cuci tangan, dan pasang plastik sampah baru."],notes:"Kain lap warna MERAH mutlak hanya untuk toilet dan mangkuk spittoon. Dilarang mencampur cairan kimia sembarangan."},{id:5,title:"Pembersihan Area Parkir & Luar",category:"Pembersihan Rutin",icon:"\u{1F17F}\uFE0F",color:"#4F46E5",bg:"#EEF2FF",target:"Halaman parkir mobil & motor, saluran air luar, teras depan",chemical_ratio:"-",tools:["Sapu Lidi Gagang Panjang","Pengki Nagata","Plastik Sampah Hitam Besar (90x120)","Selang Air"],steps:["Sapu daun kering, puntung rokok, plastik, dan pasir dari batas trotoar/pagar menuju satu titik.","Kumpulkan dan angkat sampah ke dalam pengki, masukkan langsung ke plastik sampah hitam besar.","Periksa dan bersihkan saluran drainase air di area parkir dari sumbatan sampah padat.","Bila ada tumpahan oli atau kotoran tanah membandel, siram dengan air dan sikat.","Rapikan pembatas parkir dan pastikan area depan klinik terlihat bersih dan rapi menyambut pasien."],notes:"Lakukan sebelum jam operasional klinik dibuka (pukul 07.30 - 08.00)."},{id:6,title:"Pembersihan Kaca & Jendela",category:"Pembersihan Rutin",icon:"\u{1FA9F}",color:"#0891B2",bg:"#ECFEFF",target:"Pintu kaca utama, jendela kaca klinik, partisi kaca, cermin toilet",chemical_ratio:"30 ml Glass Cleaner : 1 Liter Air / Siap pakai di botol spray",tools:["Botol Spray Glass Cleaner Prokleen","Window Wiper / Window Squeegee","Kain Microfiber Kering / Kanebo","Tongkat Teleskopik (Kaca Tinggi)"],steps:["Semprotkan larutan glass cleaner merata pada permukaan kaca yang hendak dibersihkan.","Tarik wiper karet dari arah atas ke bawah secara lurus vertikal atau zigzag horizontal tanpa terputus.","Lap bilah karet wiper setiap selesai satu tarikan menggunakan kain lap bersih agar tidak meninggalkan garis bekas air.","Gunakan kain microfiber kering untuk membersihkan sisa tetesan air di sudut-sudut lis bingkai kaca.","Untuk kaca bagian luar yang tinggi, pasang wiper pada tongkat sambungan teleskopik."],notes:"Hindari membersihkan kaca saat terpapar terik sinar matahari langsung karena cairan akan cepat mengering dan membekas noda."},{id:7,title:"Mencuci Bola Mandi Bola (Playground)",category:"Pembersihan Khusus",icon:"\u{1F939}",color:"#7C3AED",bg:"#F5F3FF",target:"Area bermain anak (Kids Playground) dan bola-bola mandi bola",chemical_ratio:"30 ml Dishwashing Liquid per 1 Liter Air + Karbol Disinfektan",tools:["Baskom Plastik Besar 55cm","Dishwashing Liquid Prokleen","Karbol Disinfektan Pine","Lap Microfiber Bersih","Karung / Wadah Pengering Jaring"],steps:["Kumpulkan seluruh bola mandi bola ke dalam wadah atau karung bersih.","Siapkan baskom plastik besar berisi air hangat kuku dicampur Dishwashing Liquid dan sedikit karbol disinfektan.","Rendam bola dan cuci secara bertahap sambil digosok perlahan menggunakan sponge/lap bersih.","Bilas bola dengan air bersih mengalir hingga tidak ada sisa busa sabun yang menempel.","Tiriskan dan angin-anginkan bola di tempat bersih beralaskan terpal/kain kering hingga benar-benar kering.","Seka dan sterilkan area kolam mandi bola sebelum bola dimasukkan kembali."],notes:"Jadwal pencucian berkala dilakukan minimal 1 kali seminggu atau sesuai kebutuhan kebersihan arena anak."},{id:8,title:"Membersihkan Filter Air Purifier",category:"Elektronik & Utility",icon:"\u{1F4A8}",color:"#0D9488",bg:"#F0FDFA",target:"Unit Air Purifier di ruang tunggu, ruang dokter, dan koridor",chemical_ratio:"-",tools:["Vacuum Cleaner (Nozzle Sikat Halus)","Kain Microfiber Kering","Kuas Cat Halus"],steps:["Matikan saklar power unit air purifier dan cabut kabel daya dari stop kontak dinding.","Buka penutup bodi (casing) depan atau belakang air purifier secara hati-hati.","Keluarkan pre-filter jaring nilon dari unit.","Bersihkan debu tebal pada pre-filter menggunakan vacuum cleaner nozzle sikat halus atau cuci dengan air mengalir (jika tipe wash-able).","Untuk filter HEPA dan Karbon Aktif: HANYA sedot debu permukaannya secara perlahan dengan vacuum, JANGAN DICUCI DENGAN AIR.","Lap bodi luar dan kisi-kisi udara masuk menggunakan lap microfiber kering.","Pastikan pre-filter benar-benar kering sebelum dipasang kembali ke posisinya semula."],notes:"Jangan pernah mencuci filter HEPA dengan air karena akan merusak serat mikro penyaring bakteri/virus."},{id:9,title:"Membersihkan Filter & Bodi AC",category:"Elektronik & Utility",icon:"\u2744\uFE0F",color:"#2563EB",bg:"#EFF6FF",target:"Filter udara dan penutup indoor unit AC split di seluruh ruangan",chemical_ratio:"-",tools:["Tangga Besi Lipat","Kain Lap Microfiber Lembap","Kuas Cat Halus","Kran Air Mengalir"],steps:["Matikan AC melalui remote dan matikan MCB/saklar AC bila diperlukan.","Gunakan tangga besi lipat yang stabil di lantai datar.","Buka tutup cover depan unit indoor AC ke atas hingga terkunci.","Lepaskan kedua kisi filter saringan udara plastik secara perlahan agar debu tidak berjatuhan.","Cuci saringan filter di bawah kran air mengalir dari arah bagian dalam ke bagian luar.","Keringkan filter di tempat teduh (jangan dijemur langsung di bawah terik matahari).","Lap bodi casing dan louver/swing AC dengan kain microfiber lembap.","Pasang kembali filter yang sudah kering, tutup cover AC, dan nyalakan AC untuk pengetesan."],notes:"Pembersihan rutin mandiri dilakukan 2 minggu sekali. Deep cleaning / cuci evaporator dilakukan oleh teknisi AC vendor."},{id:10,title:"Membersihkan Dental Unit",category:"Klinis & Medis",icon:"\u{1F9B7}",color:"#E11D48",bg:"#FFF1F2",target:"Kursi dental pasien, spittoon bowl, meja instrumen, lampu dental, cover bawah",chemical_ratio:"Glass Cleaner (netral) / Karbol Disinfektan (mangkuk) / Bowl Cleaner (kerak)",tools:["Sarung Tangan Karet (Wajib)","Kain Lap Microfiber Biru (Khusus Bodi Luar)","Kain Lap Microfiber Merah (Khusus Mangkuk Spittoon)","Botol Spray Glass Cleaner (25ml + Air)","Botol Spray Karbol (20ml + Air)","Botol Spray Bowl Cleaner (100ml + Air, bila berkerak)"],steps:["Pastikan saklar utama dental unit dalam posisi OFF dan kursi dinaikkan ke posisi tertinggi untuk memudahkan akses.","Lap seluruh jok kulit, sandaran kepala, sandaran tangan, dan tiang lampu menggunakan lap microfiber biru dan cairan pembersih netral/glass cleaner.","Bersihkan mangkuk kumur (spittoon bowl) menggunakan kain lap merah dan cairan karbol disinfektan.","Jika ditemukan kerak sisa darah/plak membandel di mangkuk spittoon, gunakan sedikit bowl cleaner, sikat perlahan, lalu bilas hingga bersih.","Lap permukaan meja instrumen meja dokter. Bila ada instrumen steril, geser dengan hati-hati dan jangan mengotori instrumen.","Buka penutup cover bawah unit dental, lap debu yang terjangkau dan pastikan kondisi area selang dalam keadaan kering."],notes:"Kain lap warna MERAH hanya untuk mangkuk spittoon. Dilarang keras menggosok jok kulit dental unit dengan bahan kimia asam keras."},{id:11,title:"Membersihkan Wastafel & Keran Air",category:"Sanitasi Toilet",icon:"\u{1F6B0}",color:"#0284C7",bg:"#F0F9FF",target:"Wastafel keramik ruang tindakan, wastafel lorong, kran stainless steel, afur sink",chemical_ratio:"Bubuk Pembersih KIFA / Dishwashing Liquid Prokleen",tools:["Sponge Scotch Brite Hijau/Kuning","Kain Microfiber Hijau (Area Basah)","Bubuk Pembersih KIFA","Sikat Gigi Bekas / Sikat Kecil"],steps:["Basahi seluruh mangkuk wastafel keramik dan kran air stainless steel.","Taburkan sedikit Bubuk KIFA pada tapas sponge yang basah.","Gosok perlahan noda sabun dan kerak air putih pada kran stainless steel searah guratan logam hingga mengkilap kembali.","Gunakan sikat kecil untuk membersihkan sela leher kran dan sudut sambungan afur pembuangan wastafel.","Bilas mangkuk wastafel dan kran dengan air bersih mengalir hingga seluruh sisa residu bubuk KIFA hilang.","Keringkan kran air stainless steel dan cermin wastafel menggunakan lap microfiber hijau kering."],notes:"DILARANG KERAS menggunakan Toilet Bowl Cleaner pada kran air stainless steel karena menyebabkan karat dan warna menghitam permanen."},{id:12,title:"Membersihkan Kulkas Klinik & Pantry",category:"Elektronik & Utility",icon:"\u{1F9CA}",color:"#475569",bg:"#F8FAFC",target:"Kulkas pantry staf dan kulkas penyimpan obat/material medis (eksterior & rak)",chemical_ratio:"30 ml Dishwashing Liquid : 1 Liter Air Hangat + Lap Kering",tools:["Kain Microfiber Biru & Hijau","Dishwashing Liquid","Ember Air Hangat","Sponge Scotch Brite Halus"],steps:["Keluarkan makanan/minuman yang telah kedaluwarsa atau berbau tidak sedap.","Lepas rak-rak kaca/plastik interior kulkas yang dapat dilepas, cuci dengan sabun cuci piring dan bilas air bersih.","Seka dinding interior kulkas menggunakan kain microfiber yang dicelupkan ke dalam air hangat bersabun lembut.","Lap kering seluruh permukaan dalam kulkas dengan microfiber kering bersih.","Pasang kembali rak interior dan susun kembali bahan/makanan secara tertib dan rapi.","Bersihkan karet pintu kulkas dari noda hitam/jamur menggunakan sikat halus dan lap kering.","Lap bodi luar kulkas dan pegangan pintu menggunakan cairan pembersih netral."],notes:"Untuk kulkas penyimpanan obat/material dental, koordinasikan terlebih dahulu dengan Perawat/PIC Obat sebelum membersihkan rak interior."},{id:13,title:"Membersihkan Sofa & Kursi Ruang Tunggu",category:"Pembersihan Khusus",icon:"\u{1F6CB}\uFE0F",color:"#9333EA",bg:"#FAF5FF",target:"Sofa ruang tunggu pasien, kursi reception, kursi kerja tim",chemical_ratio:"Baking Soda (penghilang bau) / Pembersih Upholstery secukupnya",tools:["Vacuum Cleaner (Nozzle Kecil & Sikat Lembut)","Kain Microfiber Bersih","Sikat Bulu Halus","Ember Air Hangat","Cairan Pembersih Khusus Sofa (Upholstery Cleaner)","Baking Soda (Penghilang Bau Tidak Sedap)","Sarung Tangan Karet"],steps:["Kenali bahan sofa (kain fabric, beludru, kulit asli, atau kulit sintetis/oscar).","Sedot debu di seluruh permukaan, lipatan jahitan, dan celah bantal sofa menggunakan vacuum cleaner nozzle kecil.","Gunakan sikat bulu halus untuk mengangkat rambut atau kotoran yang menempel kuat di serat kain.","Jika ada noda bercak (kopi, makanan, tinta), semprotkan pembersih upholstery secukupnya, gosok memutar perlahan dengan microfiber.","Celupkan kain microfiber ke air hangat suam-suam kuku, peras hingga lembap, lalu seka seluruh permukaan sofa.","Lap kembali dengan kain kering bersih untuk menyerap sisa kelembapan.","Jika sofa berbau apek: taburkan baking soda tipis di permukaan, diamkan 20 menit, lalu sedot kembali dengan vacuum cleaner.","Biarkan sofa mengering secara alami di dalam ruangan berventilasi atau ber-AC."],notes:"Jangan menjemur sofa langsung di bawah terik matahari karena dapat menyebabkan bahan kulit pecah-pecah dan warna kain memudar."}],wr=[{no:1,name:"Sapu Lidi Gagang",qty:"1 unit",area:"Area parkir mobil & motor klinik"},{no:2,name:"Sapu & Pengki Nagata 742s",qty:"1 set",area:"Area pantry, balkon, dak lantai atas, atau teras"},{no:3,name:"Lobby Duster Set 60 cm",qty:"1 set",area:"Lobby utama dan koridor/lorong klinik"},{no:4,name:"Lobby Duster Set 40 cm / Refill",qty:"1 unit",area:"Area ruang tindakan dokter"},{no:5,name:"Single Bucket Mop Wringer 20L",qty:"1 paket",area:"Ember pel dengan pemeras + tongkat pel"},{no:6,name:"Double Bucket Wringer Trolley 46L",qty:"1 paket",area:"Khusus klinik besar / klinik dengan fasilitas lift"},{no:7,name:"Kemoceng Bulu Lembut",qty:"1 unit",area:"Membersihkan debu kasar meja dan perabot"},{no:8,name:"Lap Microfiber Warna Merah",qty:"2 pcs",area:"Khusus area toilet, sanitasi, dan mangkuk spittoon"},{no:9,name:"Lap Microfiber Warna Biru",qty:"2 pcs",area:"Area umum kering (meja, kursi, bodi peralatan)"},{no:10,name:"Lap Microfiber Warna Hijau",qty:"2 pcs",area:"Area umum basah (pantry, wastafel luar)"},{no:11,name:"Kanebo / Synthetic Chamois",qty:"1 pcs",area:"Area umum, cermin, dan pengering kaca"},{no:12,name:"Caddy Bag Cleaning",qty:"1 unit",area:"Tas perlengkapan pinggang petugas cleaning"},{no:13,name:"Sponge Scotch Brite 3M",qty:"2 pcs",area:"Pembersihan noda lantai dan kerak ringan (spotting)"},{no:14,name:"Scotch-Brite 3M Sikat Toilet",qty:"1 unit",area:"Sikat khusus kloset dan leher angsa"},{no:15,name:"Kapi Gagang Karet / Scrapper",qty:"1 unit",area:"Pembersihan kerak lantai, noda permen karet, stiker"},{no:16,name:"Kuas Cat 4 Inch",qty:"2 pcs",area:"Pembersih sela keyboard, ventilasi, dan sudut sempit"},{no:17,name:"Warning Sign (Lantai Licin)",qty:"1 unit",area:"Digunakan saat mopping dan general cleaning"},{no:18,name:"Teleskopik + Pembersih Kaca/Langit",qty:"1 set",area:"Membersihkan jendela tinggi, plafon, dan sarang laba-laba"},{no:19,name:"Sarung Tangan Karet 3M",qty:"1 pasang",area:"Perlindungan tangan area toilet dan bahan kimia"},{no:20,name:"Botol Spray 500ml",qty:"5 unit",area:"Wadah larutan chemical (diberi label masing-masing)"},{no:21,name:"Pengharum Ruangan Stella Aerosol",qty:"1 kaleng",area:"Digunakan darurat apabila timbul bau menyengat"},{no:22,name:"Kleaner Gantungan Sapu & Pel",qty:"1 set",area:"Gantungan penyimpanan alat cleaning di janitor"},{no:23,name:"Floor Squeegee 75 cm",qty:"1 unit",area:"Pendorong genangan air lantai toilet/teras"},{no:24,name:"Sikat Botol Ikea",qty:"1 unit",area:"Di setiap ruang tindakan dokter gigi"},{no:25,name:"Scotch-Brite 3M Sikat Lantai ID-51",qty:"1 unit",area:"Sikat lantai kamar mandi dan teras"},{no:26,name:"Kamper Toilet / Urinal Block",qty:"1 pack",area:"Pengharum di toilet dan urinoir"},{no:27,name:"Plastik Sampah Hitam Sedang (60x100)",qty:"1 pack",area:"Tempat sampah indoor ruangan"},{no:28,name:"Plastik Sampah Hitam Besar (90x120)",qty:"1 pack",area:"Tempat sampah luar & penampung sampah akhir"},{no:29,name:"Plastik Sampah Medis Kuning (90x120)",qty:"1 pack",area:"Limbah medis padat B3 klinik"},{no:30,name:"Plastik Sampah Medis Kuning (40x60)",qty:"1 pack",area:"Tempat sampah medis kecil di ruang tindakan"},{no:31,name:"Tangga Besi Lipat",qty:"1 unit",area:"Alat bantu membersihkan plafon, lampu, dan AC"},{no:32,name:"Baskom Plastik Besar 55cm",qty:"1 unit",area:"Mencuci bola kids playground"}],Sr=[{no:1,name:"Floor Cleaner Prokleen Lavender",size:"5 Liter",ratio:"25 ml : 1 Liter Air",function:"Pembersihan dan pengepelan rutin seluruh lantai ruangan"},{no:2,name:"Karbol Prokleen Pine Disinfektan",size:"5 Liter",ratio:"25 ml : 1 Liter Air (Mop) / 20 ml per botol spray",function:"Mengepel lantai toilet, disinfeksi area basah dan berbau tak sedap"},{no:3,name:"Glass Cleaner Prokleen",size:"5 Liter",ratio:"30 ml : 1 Liter Air / Siap pakai di botol spray",function:"Membersihkan pintu kaca, jendela, cermin, dan layar monitor"},{no:4,name:"Manual Dishwashing Liquid Prokleen",size:"5 Liter",ratio:"30 ml : 1 Liter Air",function:"Mencuci piring, gelas, sendok pantry, dan mencuci bola playground"},{no:5,name:"Hand Soap Prokleen Green Tea",size:"5 Liter",ratio:"Siap pakai (tanpa campur)",function:"Refill sabun pencuci tangan di wastafel pasien dan tim"},{no:6,name:"Liquid Detergent Prokleen",size:"5 Liter",ratio:"Sesuai kapasitas mesin",function:"Mencuci handuk, kain microfiber, dan seragam"},{no:7,name:"Toilet Bowl Cleaner Prokleen",size:"5 Liter",ratio:"100 ml per botol spray / Khusus kerak",function:"Pembersih kerak membandel pada kloset, wastafel keramik, dan nat keramik"},{no:8,name:"Bubuk Pembersih KIFA",size:"500 gram",ratio:"Taburkan tipis pada tapas basah",function:"Membersihkan kerak air pada kran air stainless steel, jet washer, dan sink"}],it="fcms_sop_fc_ik_v1",rt="fcms_sop_fc_tools_v1",ot="fcms_sop_fc_chem_v1";function On(){try{let e=localStorage.getItem(it);if(e)return JSON.parse(e)}catch(e){console.error("Error load IK",e)}return JSON.parse(JSON.stringify(kr))}function Pa(e){localStorage.setItem(it,JSON.stringify(e)),X("/api/sop/bundle/"+it,{method:"PUT",body:JSON.stringify(e)}).catch(()=>{})}function Hn(){try{let e=localStorage.getItem(rt);if(e)return JSON.parse(e)}catch(e){console.error("Error load tools",e)}return JSON.parse(JSON.stringify(wr))}function La(e){localStorage.setItem(rt,JSON.stringify(e)),X("/api/sop/bundle/"+rt,{method:"PUT",body:JSON.stringify(e)}).catch(()=>{})}function Un(){try{let e=localStorage.getItem(ot);if(e)return JSON.parse(e)}catch(e){console.error("Error load chem",e)}return JSON.parse(JSON.stringify(Sr))}function Da(e){localStorage.setItem(ot,JSON.stringify(e)),X("/api/sop/bundle/"+ot,{method:"PUT",body:JSON.stringify(e)}).catch(()=>{})}async function Gn(e){let t=de(),n=ie(t,"sop","write"),a=ie(t,"sop","delete"),o="ik",s="",l="all",h="tools",p="";try{let[g,D,O]=await Promise.all([X("/api/sop/bundle/"+it),X("/api/sop/bundle/"+rt),X("/api/sop/bundle/"+ot)]);g.ok&&g.data?.data&&localStorage.setItem(it,JSON.stringify(g.data.data)),D.ok&&D.data?.data&&localStorage.setItem(rt,JSON.stringify(D.data.data)),O.ok&&O.data?.data&&localStorage.setItem(ot,JSON.stringify(O.data.data))}catch(g){console.warn("Could not load remote SOP bundle:",g)}let c=On(),E=Hn(),C=Un();function T(){e.innerHTML=`
      <div class="page-header" style="display:flex; justify-content:space-between; align-items:flex-start; flex-wrap:wrap; gap:16px; margin-bottom:20px;">
        <div>
          <h1 class="page-title" style="display:flex; align-items:center; gap:8px; margin:0 0 6px 0; font-size:1.6rem; font-weight:700; color:var(--text-1);">
            <span>\u{1F4DA}</span> SOP Facility Care & Housekeeping
          </h1>
          <p class="page-subtitle" style="margin:0; font-size:0.9rem; color:var(--text-3);">
            Pedoman Standar Operasional Prosedur Kebersihan & Pemeliharaan Fasilitas PT Audy Mandiri Indonesia
          </p>
        </div>
        <div style="display:flex; gap:10px; align-items:center; flex-wrap:wrap;">
          ${a?`
          <button id="btn-reset-default" class="btn btn-outline" title="Kembalikan data ke dokumen standar asli" style="background:#172033; border:1px solid var(--border); color:#F8FAFC; font-weight:600; padding:8px 14px; border-radius:8px; display:inline-flex; align-items:center; gap:6px; cursor:pointer; font-size:0.85rem;">
            \u{1F504} Reset ke Standar
          </button>
          `:""}
          <button id="btn-print-sop" class="btn btn-outline" style="background:#172033; border:1px solid var(--border); color:#F8FAFC; font-weight:600; padding:8px 14px; border-radius:8px; display:inline-flex; align-items:center; gap:6px; cursor:pointer; font-size:0.85rem;">
            \u{1F5A8}\uFE0F Cetak Ringkasan SOP
          </button>
          ${ie(t,"export")?`<button id="btn-export-tools" class="btn btn-primary" style="font-weight:600; padding:8px 16px; border-radius:8px; display:inline-flex; align-items:center; gap:6px; cursor:pointer; font-size:0.85rem;">
            \u{1F4E5} Download Data Excel
          </button>`:""}
        </div>
      </div>

      <!-- Navigation Tabs -->
      <div style="display:flex; gap:8px; margin-bottom:22px; border-bottom:2px solid #E2E8F0; padding-bottom:8px; overflow-x:auto; flex-wrap:nowrap; -webkit-overflow-scrolling:touch; scrollbar-width:none;">
        <button class="sop-tab-btn ${o==="ik"?"active":""}" data-tab="ik" style="${u(o==="ik")}">
          \u{1F4CB} Instruksi Kerja (IK) <span class="badge" style="${F(o==="ik")}">${c.length}</span>
        </button>
        <button class="sop-tab-btn ${o==="tools"?"active":""}" data-tab="tools" style="${u(o==="tools")}">
          \u{1F6E0}\uFE0F Alat & Chemical <span class="badge" style="${F(o==="tools")}">${E.length+C.length}</span>
        </button>
        <button class="sop-tab-btn ${o==="policy"?"active":""}" data-tab="policy" style="${u(o==="policy")}">
          \u{1F4D6} Pedoman & Ketentuan
        </button>
        <button class="sop-tab-btn ${o==="links"?"active":""}" data-tab="links" style="${u(o==="links")}">
          \u{1F517} Form & Tautan
        </button>
      </div>

      <!-- Content Container -->
      <div id="sop-content-area"></div>
    `,R(),K()}function u(g){return g?"border:none; background:#2563EB; color:#fff; font-weight:700; padding:8px 16px; border-radius:8px; cursor:pointer; display:inline-flex; align-items:center; gap:6px; font-size:0.9rem; white-space:nowrap; transition:all 0.2s;":"border:none; background:#172033; color:var(--text-2); border:1px solid var(--border); font-weight:600; padding:8px 16px; border-radius:8px; cursor:pointer; display:inline-flex; align-items:center; gap:6px; font-size:0.9rem; white-space:nowrap; transition:all 0.2s;"}function F(g){return g?"background:rgba(255,255,255,0.25); color:#fff; font-size:0.75rem; padding:2px 7px; border-radius:10px; font-weight:700;":"background:#CBD5E1; color:#334155; font-size:0.75rem; padding:2px 7px; border-radius:10px; font-weight:700;"}function R(){e.querySelectorAll(".sop-tab-btn").forEach(W=>{W.addEventListener("click",()=>{o=W.dataset.tab,T()})});let g=e.querySelector("#btn-reset-default");g&&g.addEventListener("click",()=>{me("Apakah Anda yakin ingin mengembalikan seluruh data Instruksi Kerja, Alat, dan Chemical ke data standar dokumen asli?",()=>{localStorage.removeItem(it),localStorage.removeItem(rt),localStorage.removeItem(ot),Promise.all([X("/api/sop/bundle/"+it,{method:"DELETE"}),X("/api/sop/bundle/"+rt,{method:"DELETE"}),X("/api/sop/bundle/"+ot,{method:"DELETE"})]).catch(()=>{}),c=On(),E=Hn(),C=Un(),ae("Data SOP Facility Care berhasil di-reset ke standar asli"),T()},"Konfirmasi Reset Standar")});let D=e.querySelector("#btn-print-sop");D&&D.addEventListener("click",()=>{window.print()});let O=e.querySelector("#btn-export-tools");O&&O.addEventListener("click",()=>{P()})}function K(){let g=e.querySelector("#sop-content-area");g&&(o==="ik"?I(g):o==="tools"?M(g):o==="policy"?_(g):o==="links"&&B(g))}function I(g){g.innerHTML=`
      <!-- Toolbar Filter & Search -->
      <div class="card" style="padding:14px 18px; margin-bottom:20px; background:#172033; border:1px solid var(--border); border-radius:12px; box-shadow:0 1px 3px rgba(0,0,0,0.2);">
        <div style="display:flex; flex-wrap:wrap; gap:12px; align-items:center; justify-content:space-between;">
          <div style="display:flex; flex-wrap:wrap; gap:12px; align-items:center; flex:1; min-width:280px;">
            <div style="position:relative; flex:1; min-width:220px; max-width:340px;">
              <input type="text" id="ik-search" class="form-control" placeholder="Cari pekerjaan, alat, atau chemical..." value="${s}" style="padding-left:34px; height:38px; border-radius:8px; font-size:0.9rem; width:100%;">
              <span style="position:absolute; left:11px; top:50%; transform:translateY(-50%); color:var(--text-3); font-size:0.9rem;">\u{1F50D}</span>
            </div>
            
            <div style="display:flex; gap:6px; flex-wrap:wrap;" id="cat-pills-wrap">
              <button class="cat-pill ${l==="all"?"active":""}" data-cat="all">Semua (${c.length})</button>
              <button class="cat-pill ${l==="Pembersihan Rutin"?"active":""}" data-cat="Pembersihan Rutin">Rutin (${c.filter(i=>i.category==="Pembersihan Rutin").length})</button>
              <button class="cat-pill ${l==="Sanitasi Toilet"?"active":""}" data-cat="Sanitasi Toilet">Toilet & Wastafel (${c.filter(i=>i.category==="Sanitasi Toilet").length})</button>
              <button class="cat-pill ${l==="Klinis & Medis"?"active":""}" data-cat="Klinis & Medis">Klinis Dental (${c.filter(i=>i.category==="Klinis & Medis").length})</button>
              <button class="cat-pill ${l==="Elektronik & Utility"?"active":""}" data-cat="Elektronik & Utility">Elektronik & AC (${c.filter(i=>i.category==="Elektronik & Utility").length})</button>
              <button class="cat-pill ${l==="Pembersihan Khusus"?"active":""}" data-cat="Pembersihan Khusus">Khusus (${c.filter(i=>i.category==="Pembersihan Khusus").length})</button>
            </div>
          </div>
          <div style="display:flex; gap:10px; align-items:center;">
            <div id="ik-count-label" style="font-size:0.85rem; color:var(--text-3); font-weight:600;">
              Menampilkan ${c.length} dari ${c.length} Instruksi Kerja
            </div>
            ${n?`
            <button id="btn-add-ik" class="btn btn-primary" style="font-weight:600; padding:7px 14px; border-radius:8px; display:inline-flex; align-items:center; gap:6px; cursor:pointer; font-size:0.85rem;">
              \u2795 Tambah IK Baru
            </button>
            `:""}
          </div>
        </div>
      </div>

      <!-- IK Cards Grid Container -->
      <div id="ik-cards-grid-container"></div>
    `;let D=document.createElement("style");D.innerHTML=`
      .cat-pill {
        border: 1px solid var(--border);
        background: #172033;
        color: var(--text-2);
        font-size: 0.8rem;
        font-weight: 600;
        padding: 5px 12px;
        border-radius: 20px;
        cursor: pointer;
        transition: all 0.15s;
        flex-shrink: 0;
        white-space: nowrap;
      }
      .cat-pill:hover {
        background: #22314A;
      }
      .cat-pill.active {
        background: #2563EB;
        color: #fff;
        border-color: #2563EB;
      }
      .ik-card:hover {
        box-shadow: 0 6px 16px rgba(0,0,0,0.08) !important;
        transform: translateY(-2px);
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
    `,g.appendChild(D);let O=g.querySelector("#ik-cards-grid-container"),W=g.querySelector("#ik-count-label");function J(){let i=c.filter(r=>{let d=l==="all"||r.category===l,b=s.toLowerCase().trim(),y=!b||r.title.toLowerCase().includes(b)||r.target.toLowerCase().includes(b)||r.tools.some(q=>q.toLowerCase().includes(b))||r.steps.some(q=>q.toLowerCase().includes(b));return d&&y});if(W&&(W.textContent=`Menampilkan ${i.length} dari ${c.length} Instruksi Kerja`),i.length===0){O.innerHTML=`
          <div style="text-align:center; padding:48px 20px; background:#111827; border-radius:12px; border:1px solid var(--border);">
            <div style="font-size:2.5rem; margin-bottom:8px;">\u{1F50D}</div>
            <h3 style="margin:0 0 4px; font-size:1.1rem; color:var(--text-1);">Tidak ada Instruksi Kerja yang cocok</h3>
            <p style="margin:0; font-size:0.85rem; color:var(--text-3);">Coba ubah kata kunci pencarian atau pilih kategori lain.</p>
          </div>
        `;return}O.innerHTML=`
        <div class="sop-ik-grid" style="display:grid; grid-template-columns:repeat(auto-fill, minmax(340px, 1fr)); gap:18px;">
          ${i.map(r=>`
            <div class="card ik-card" style="background:#172033; border:1px solid var(--border); border-radius:12px; padding:18px; display:flex; flex-direction:column; justify-content:space-between; box-shadow:0 1px 4px rgba(0,0,0,0.04); transition:all 0.2s;">
              <div>
                <div style="display:flex; justify-content:space-between; align-items:flex-start; margin-bottom:12px;">
                  <div style="display:flex; align-items:center; gap:10px;">
                    <span style="font-size:1.6rem; width:44px; height:44px; display:inline-flex; align-items:center; justify-content:center; background:${r.bg||"#EFF6FF"}; border-radius:10px;">${r.icon||"\u{1F4CB}"}</span>
                    <div>
                      <h3 style="margin:0 0 2px 0; font-size:1.05rem; font-weight:700; color:var(--text-1);">
                        ${r.id}. ${r.title}
                      </h3>
                      <span class="badge" style="background:${r.bg||"#EFF6FF"}; color:${r.color||"#2563EB"}; font-size:0.75rem; font-weight:700; border-radius:6px; padding:2px 8px;">
                        ${r.category}
                      </span>
                    </div>
                  </div>

                  <!-- Quick Action Buttons: Edit & Hapus -->
                  <div style="display:flex; gap:4px;">
                    ${n?`
                    <button class="btn btn-sm btn-ghost btn-edit-ik" data-id="${r.id}" title="Edit Instruksi Kerja" style="padding:4px 8px; border-radius:6px; font-size:0.9rem; cursor:pointer;">
                      \u270F\uFE0F
                    </button>
                    `:""}
                    ${a?`
                    <button class="btn btn-sm btn-ghost btn-delete-ik" data-id="${r.id}" title="Hapus Instruksi Kerja" style="padding:4px 8px; border-radius:6px; font-size:0.9rem; color:#DC2626; cursor:pointer;">
                      \u{1F5D1}\uFE0F
                    </button>
                    `:""}
                  </div>
                </div>

                <div style="font-size:0.85rem; color:var(--text-2); margin-bottom:14px; background:#111827; padding:8px 12px; border-radius:8px; border-left:3px solid ${r.color||"#2563EB"};">
                  <strong>Area:</strong> ${r.target}
                </div>

                <div style="margin-bottom:14px;">
                  <div style="font-size:0.8rem; font-weight:700; color:var(--text-3); text-transform:uppercase; margin-bottom:6px;">\u{1F9F0} Alat & Bahan Utama:</div>
                  <ul style="margin:0 0 0 16px; padding:0; font-size:0.83rem; color:var(--text-2);">
                    ${r.tools.slice(0,3).map(d=>`<li style="margin-bottom:2px;">${d}</li>`).join("")}
                    ${r.tools.length>3?`<li style="color:var(--text-3); font-style:italic;">+ ${r.tools.length-3} alat lainnya...</li>`:""}
                  </ul>
                </div>

                ${r.chemical_ratio&&r.chemical_ratio!=="-"?`
                  <div style="font-size:0.8rem; background:#FEF3C7; color:#92400E; padding:6px 10px; border-radius:6px; margin-bottom:14px; font-weight:600;">
                    \u{1F9EA} Takaran: ${r.chemical_ratio}
                  </div>
                `:""}
              </div>

              <div style="border-top:1px solid #F1F5F9; padding-top:12px; margin-top:8px; display:flex; justify-content:space-between; align-items:center;">
                <span style="font-size:0.8rem; color:var(--text-3); font-weight:500;">
                  ${r.steps.length} Langkah Kerja
                </span>
                <div style="display:flex; gap:6px;">
                  <button class="btn btn-sm btn-primary btn-view-ik" data-id="${r.id}" style="padding:6px 14px; font-weight:600; border-radius:6px;">
                    Lihat Panduan \u2794
                  </button>
                </div>
              </div>
            </div>
          `).join("")}
        </div>
      `,O.querySelectorAll(".btn-view-ik").forEach(r=>{r.addEventListener("click",()=>{let d=parseInt(r.dataset.id,10),b=c.find(y=>y.id===d);b&&N(b)})}),O.querySelectorAll(".btn-edit-ik").forEach(r=>{r.addEventListener("click",()=>{let d=parseInt(r.dataset.id,10),b=c.find(y=>y.id===d);b&&w(b)})}),O.querySelectorAll(".btn-delete-ik").forEach(r=>{r.addEventListener("click",()=>{let d=parseInt(r.dataset.id,10),b=c.find(y=>y.id===d);b&&me(`Apakah Anda yakin ingin menghapus Instruksi Kerja: "${b.title}"?`,()=>{c=c.filter(y=>y.id!==d),Pa(c),ae("Instruksi Kerja berhasil dihapus"),T()},"Hapus Instruksi Kerja")})})}J();let $=g.querySelector("#ik-search");$&&$.addEventListener("input",i=>{s=i.target.value,J()}),g.querySelectorAll(".cat-pill").forEach(i=>{i.addEventListener("click",()=>{g.querySelectorAll(".cat-pill").forEach(r=>r.classList.remove("active")),i.classList.add("active"),l=i.dataset.cat,J()})});let m=g.querySelector("#btn-add-ik");m&&m.addEventListener("click",()=>{w()})}function N(g){pe({title:`${g.icon||"\u{1F4CB}"} IK-${g.id}: ${g.title}`,size:"lg",content:`
        <div style="display:flex; flex-direction:column; gap:16px;">
          <div style="background:${g.bg||"#EFF6FF"}; color:${g.color||"#2563EB"}; padding:10px 14px; border-radius:8px; font-weight:700; font-size:0.9rem; display:flex; justify-content:space-between; align-items:center;">
            <span>Kategori: ${g.category}</span>
            <span>ID: IK-${String(g.id).padStart(2,"0")}</span>
          </div>

          <div style="background:#111827; padding:12px 14px; border-radius:8px; border-left:4px solid ${g.color||"#2563EB"};">
            <strong style="color:var(--text-1); font-size:0.85rem; display:block; margin-bottom:4px;">\u{1F4CD} Area & Sasaran Target:</strong>
            <span style="font-size:0.88rem; color:var(--text-2);">${g.target}</span>
          </div>

          <div>
            <h4 style="margin:0 0 8px 0; font-size:0.95rem; font-weight:700; color:var(--text-1);">\u{1F9F0} Alat & Bahan yang Dibutuhkan:</h4>
            <div style="background:#111827; border:1px solid var(--border); border-radius:8px; padding:10px 14px;">
              <ul style="margin:0 0 0 16px; padding:0; font-size:0.875rem; color:var(--text-2); line-height:1.7;">
                ${g.tools.map(D=>`<li>${D}</li>`).join("")}
              </ul>
            </div>
          </div>

          ${g.chemical_ratio&&g.chemical_ratio!=="-"?`
            <div style="background:#FEF3C7; border:1px solid #FCD34D; border-radius:8px; padding:10px 14px;">
              <strong style="color:#92400E; font-size:0.85rem; display:block; margin-bottom:2px;">\u{1F9EA} Rasio & Takaran Chemical:</strong>
              <span style="color:#78350F; font-size:0.9rem; font-weight:600;">${g.chemical_ratio}</span>
            </div>
          `:""}

          <div>
            <h4 style="margin:0 0 8px 0; font-size:0.95rem; font-weight:700; color:var(--text-1);">\u{1F4DD} Langkah-Langkah Pelaksanaan:</h4>
            <div style="background:#111827; border:1px solid var(--border); border-radius:8px; padding:12px 16px;">
              <ol style="margin:0 0 0 18px; padding:0; font-size:0.875rem; color:var(--text-2); line-height:1.8;">
                ${g.steps.map(D=>`<li style="margin-bottom:6px;">${D}</li>`).join("")}
              </ol>
            </div>
          </div>

          ${g.notes?`
            <div style="background:#EFF6FF; border:1px solid #BFDBFE; border-radius:8px; padding:10px 14px;">
              <strong style="color:#1E40AF; font-size:0.85rem; display:block; margin-bottom:2px;">\u26A0\uFE0F Catatan Penting & Safety:</strong>
              <span style="color:#1E3A8A; font-size:0.85rem; line-height:1.6;">${g.notes}</span>
            </div>
          `:""}
        </div>
      `,confirmText:"Tutup Panduan",onConfirm:(D,O)=>{O()}})}function w(g=null){if(!n)return;let D=!!g,O=g?g.tools.join(`
`):"",W=g?g.steps.join(`
`):"",J=`
      <form id="ik-edit-form" style="display:flex; flex-direction:column; gap:14px; font-size:0.88rem;">
        <div class="form-row" style="display:grid; grid-template-columns:1fr 140px; gap:12px;">
          <div>
            <label style="display:block; font-weight:600; margin-bottom:4px;">Judul Instruksi Kerja *</label>
            <input type="text" id="form-ik-title" class="form-control" required value="${g?f(g.title):""}" placeholder="Contoh: Membersihkan Dental Unit" style="width:100%; border-radius:6px; height:36px;">
          </div>
          <div>
            <label style="display:block; font-weight:600; margin-bottom:4px;">Ikon Emoji</label>
            <input type="text" id="form-ik-icon" class="form-control" value="${g?g.icon:"\u{1F4CB}"}" placeholder="\u{1F9F9}" style="width:100%; border-radius:6px; height:36px; text-align:center; font-size:1.2rem;">
          </div>
        </div>

        <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-weight:600; margin-bottom:4px;">Kategori *</label>
            <select id="form-ik-cat" class="form-control" style="width:100%; border-radius:6px; height:36px;">
              <option value="Pembersihan Rutin" ${g?.category==="Pembersihan Rutin"?"selected":""}>Pembersihan Rutin</option>
              <option value="Sanitasi Toilet" ${g?.category==="Sanitasi Toilet"?"selected":""}>Sanitasi Toilet</option>
              <option value="Klinis & Medis" ${g?.category==="Klinis & Medis"?"selected":""}>Klinis & Medis</option>
              <option value="Elektronik & Utility" ${g?.category==="Elektronik & Utility"?"selected":""}>Elektronik & Utility</option>
              <option value="Pembersihan Khusus" ${g?.category==="Pembersihan Khusus"?"selected":""}>Pembersihan Khusus</option>
            </select>
          </div>
          <div>
            <label style="display:block; font-weight:600; margin-bottom:4px;">Takaran / Rasio Chemical</label>
            <input type="text" id="form-ik-ratio" class="form-control" value="${g?f(g.chemical_ratio||""):""}" placeholder="Contoh: 25 ml : 1 Liter Air" style="width:100%; border-radius:6px; height:36px;">
          </div>
        </div>

        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Area & Sasaran Target *</label>
          <input type="text" id="form-ik-target" class="form-control" required value="${g?f(g.target):""}" placeholder="Contoh: Kursi pasien, meja instrumen, mangkuk spittoon" style="width:100%; border-radius:6px; height:36px;">
        </div>

        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Alat & Bahan yang Digunakan (1 item per baris) *</label>
          <textarea id="form-ik-tools" class="form-control" rows="4" placeholder="Kain Microfiber Biru&#10;Sarung Tangan Karet&#10;Botol Spray Glass Cleaner" style="width:100%; border-radius:6px; font-size:0.85rem; line-height:1.5;">${f(O)}</textarea>
        </div>

        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Langkah-Langkah Pelaksanaan (1 langkah per baris) *</label>
          <textarea id="form-ik-steps" class="form-control" rows="6" placeholder="1. Pastikan posisi saklar OFF.&#10;2. Lap jok dental unit dengan kain microfiber biru.&#10;3. Sikat mangkuk spittoon..." style="width:100%; border-radius:6px; font-size:0.85rem; line-height:1.6;">${f(W)}</textarea>
        </div>

        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Catatan Penting / Hal yang Perlu Diperhatikan (Safety)</label>
          <textarea id="form-ik-notes" class="form-control" rows="2" placeholder="Catatan larangan atau peringatan keselamatan..." style="width:100%; border-radius:6px; font-size:0.85rem;">${g?f(g.notes||""):""}</textarea>
        </div>
      </form>
    `;pe({title:D?`\u270F\uFE0F Edit IK-${g.id}: ${g.title}`:"\u2795 Tambah Instruksi Kerja Baru",content:J,size:"lg",confirmText:"Simpan Perubahan",onConfirm:($,m)=>{let i=$.querySelector("#form-ik-title").value.trim(),r=$.querySelector("#form-ik-icon").value.trim()||"\u{1F4CB}",d=$.querySelector("#form-ik-cat").value,b=$.querySelector("#form-ik-ratio").value.trim(),y=$.querySelector("#form-ik-target").value.trim(),q=$.querySelector("#form-ik-tools").value.trim(),v=$.querySelector("#form-ik-steps").value.trim(),S=$.querySelector("#form-ik-notes").value.trim();if(!i||!y||!q||!v){z("Harap lengkapi seluruh field wajib (*)");return}let A=q.split(`
`).map(Q=>Q.trim().replace(/^[-•*]\s*/,"")).filter(Boolean),U=v.split(`
`).map(Q=>Q.trim().replace(/^\d+[\.\)]\s*/,"")).filter(Boolean),ee={"Pembersihan Rutin":{color:"#2563EB",bg:"#EFF6FF"},"Sanitasi Toilet":{color:"#DC2626",bg:"#FEF2F2"},"Klinis & Medis":{color:"#E11D48",bg:"#FFF1F2"},"Elektronik & Utility":{color:"#0D9488",bg:"#F0FDFA"},"Pembersihan Khusus":{color:"#7C3AED",bg:"#F5F3FF"}}[d]||{color:"#2563EB",bg:"#EFF6FF"};if(D)g.title=i,g.icon=r,g.category=d,g.chemical_ratio=b,g.target=y,g.tools=A,g.steps=U,g.notes=S,g.color=ee.color,g.bg=ee.bg,Pa(c),ae(`Instruksi Kerja "${i}" berhasil diperbarui`);else{let Q=c.length>0?Math.max(...c.map(Z=>Z.id))+1:1,ne={id:Q,title:i,category:d,icon:r,color:ee.color,bg:ee.bg,target:y,chemical_ratio:b||"-",tools:A,steps:U,notes:S};c.push(ne),Pa(c),ae(`Instruksi Kerja baru IK-${Q} berhasil ditambahkan`)}m(),T()}})}function M(g){g.innerHTML=`
      <!-- Sub-tab Switcher & Actions -->
      <div style="display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:12px; margin-bottom:18px;">
        <div class="btn-group" style="background:#111827; padding:3px; border-radius:8px; display:inline-flex; border:1px solid var(--border);">
          <button id="view-tools-btn" class="btn btn-sm ${h==="tools"?"btn-primary":"btn-ghost"}" style="font-weight:700; border-radius:6px; padding:6px 16px; cursor:pointer;">
            \u{1F9F9} Alat Kebersihan (${E.length} Item)
          </button>
          <button id="view-chem-btn" class="btn btn-sm ${h==="chemical"?"btn-primary":"btn-ghost"}" style="font-weight:700; border-radius:6px; padding:6px 16px; cursor:pointer;">
            \u{1F9EA} Bahan Chemical (${C.length} Item)
          </button>
        </div>

        <div style="display:flex; gap:10px; align-items:center; flex-wrap:wrap;">
          <input type="text" id="tool-search" class="form-control" placeholder="Cari nama / peruntukan / rasio..." value="${p}" style="height:36px; border-radius:8px; font-size:0.875rem; width:240px;">
            ${n?`
          <button id="btn-add-tool-chem" class="btn btn-primary" style="font-weight:600; padding:6px 14px; border-radius:8px; display:inline-flex; align-items:center; gap:6px; cursor:pointer; font-size:0.85rem;">
            \u2795 ${h==="tools"?"Tambah Alat Baru":"Tambah Chemical Baru"}
          </button>
            `:""}
        </div>
      </div>

      <!-- Table Container -->
      <div id="tool-table-container"></div>
    `;let D=g.querySelector("#tool-table-container"),O=g.querySelector("#tool-search");function W(){let J=p.toLowerCase().trim();if(h==="tools"){let $=E.filter(m=>!J||m.name.toLowerCase().includes(J)||m.area.toLowerCase().includes(J));D.innerHTML=`
          <div class="card" style="background:#172033; border:1px solid var(--border); border-radius:12px; overflow:hidden; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
            <div style="padding:14px 18px; background:#111827; border-bottom:1px solid var(--border); display:flex; justify-content:space-between; align-items:center;">
              <strong style="color:var(--text-1); font-size:0.95rem;">Daftar Standar Peralatan Kebersihan (Lampiran 8.1 SOP)</strong>
              <span class="badge badge-primary">${$.length} Item</span>
            </div>
            <div style="overflow-x:auto;">
              <table class="table" style="width:100%; margin:0; border-collapse:collapse; font-size:0.875rem;">
                <thead>
                  <tr style="background:#111827; border-bottom:1px solid var(--border); text-align:left;">
                    <th style="padding:10px 14px; width:60px;">No</th>
                    <th style="padding:10px 14px;">Nama Alat / Item</th>
                    <th style="padding:10px 14px; width:130px;">Kuantitas Standar</th>
                    <th style="padding:10px 14px;">Peruntukan Area & Kegunaan</th>
                    <th style="padding:10px 14px; width:110px; text-align:center;">Aksi</th>
                  </tr>
                </thead>
                <tbody>
                  ${$.map((m,i)=>`
                    <tr style="border-bottom:1px solid var(--border); background:${i%2===0?"#172033":"#111827"}; color:var(--text-1);">
                      <td style="padding:10px 14px; font-weight:600; color:var(--text-3);">${m.no||i+1}</td>
                      <td style="padding:10px 14px; font-weight:700; color:var(--text-1);">${f(m.name)}</td>
                      <td style="padding:10px 14px;"><span class="badge" style="background:#E0F2FE; color:#0369A1; font-weight:600;">${f(m.qty)}</span></td>
                      <td style="padding:10px 14px; color:var(--text-2);">${f(m.area)}</td>
                      <td style="padding:10px 14px; text-align:center;">
                    ${n?`
                        <button class="btn btn-sm btn-ghost btn-edit-tool" data-no="${m.no}" title="Edit Alat" style="padding:4px 6px; cursor:pointer;">\u270F\uFE0F</button>
                    `:""}
                    ${a?`
                        <button class="btn btn-sm btn-ghost btn-delete-tool" data-no="${m.no}" title="Hapus Alat" style="padding:4px 6px; color:#DC2626; cursor:pointer;">\u{1F5D1}\uFE0F</button>
                    `:""}
                      </td>
                    </tr>
                  `).join("")}
                </tbody>
              </table>
            </div>
          </div>
        `,D.querySelectorAll(".btn-edit-tool").forEach(m=>{m.addEventListener("click",()=>{let i=parseInt(m.dataset.no,10),r=E.find(d=>d.no===i);r&&k(r)})}),D.querySelectorAll(".btn-delete-tool").forEach(m=>{m.addEventListener("click",()=>{let i=parseInt(m.dataset.no,10),r=E.find(d=>d.no===i);r&&me(`Apakah Anda yakin ingin menghapus alat: "${r.name}"?`,()=>{E=E.filter(d=>d.no!==i),E.forEach((d,b)=>{d.no=b+1}),La(E),ae("Alat kebersihan berhasil dihapus"),T()},"Hapus Alat Kebersihan")})})}else{let $=C.filter(m=>!J||m.name.toLowerCase().includes(J)||m.function.toLowerCase().includes(J));D.innerHTML=`
          <div class="card" style="background:#172033; border:1px solid var(--border); border-radius:12px; overflow:hidden; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
            <div style="padding:14px 18px; background:#111827; border-bottom:1px solid var(--border); display:flex; justify-content:space-between; align-items:center;">
              <strong style="color:var(--text-1); font-size:0.95rem;">Daftar Bahan Kimia & Standar Takaran (Lampiran 8.1 SOP)</strong>
              <span class="badge badge-success">${$.length} Chemical</span>
            </div>
            <div style="overflow-x:auto;">
              <table class="table" style="width:100%; margin:0; border-collapse:collapse; font-size:0.875rem;">
                <thead>
                  <tr style="background:#111827; border-bottom:1px solid var(--border); text-align:left;">
                    <th style="padding:10px 14px; width:60px;">No</th>
                    <th style="padding:10px 14px;">Nama Produk Chemical</th>
                    <th style="padding:10px 14px; width:110px;">Kemasan</th>
                    <th style="padding:10px 14px; width:220px;">Rasio & Standar Takaran</th>
                    <th style="padding:10px 14px;">Fungsi & Cara Penggunaan</th>
                    <th style="padding:10px 14px; width:110px; text-align:center;">Aksi</th>
                  </tr>
                </thead>
                <tbody>
                  ${$.map((m,i)=>`
                    <tr style="border-bottom:1px solid var(--border); background:${i%2===0?"#172033":"#111827"}; color:var(--text-1);">
                      <td style="padding:10px 14px; font-weight:600; color:var(--text-3);">${m.no||i+1}</td>
                      <td style="padding:10px 14px; font-weight:700; color:var(--text-1);">${f(m.name)}</td>
                      <td style="padding:10px 14px;"><span class="badge" style="background:#FEF3C7; color:#92400E; font-weight:600;">${f(m.size)}</span></td>
                      <td style="padding:10px 14px;"><span style="background:#EFF6FF; color:#1E40AF; padding:3px 8px; border-radius:6px; font-weight:600; font-size:0.82rem;">${f(m.ratio)}</span></td>
                      <td style="padding:10px 14px; color:var(--text-2);">${f(m.function)}</td>
                      <td style="padding:10px 14px; text-align:center;">
                    ${n?`
                        <button class="btn btn-sm btn-ghost btn-edit-chem" data-no="${m.no}" title="Edit Chemical" style="padding:4px 6px; cursor:pointer;">\u270F\uFE0F</button>
                    `:""}
                    ${a?`
                        <button class="btn btn-sm btn-ghost btn-delete-chem" data-no="${m.no}" title="Hapus Chemical" style="padding:4px 6px; color:#DC2626; cursor:pointer;">\u{1F5D1}\uFE0F</button>
                    `:""}
                      </td>
                    </tr>
                  `).join("")}
                </tbody>
              </table>
            </div>
          </div>
        `,D.querySelectorAll(".btn-edit-chem").forEach(m=>{m.addEventListener("click",()=>{let i=parseInt(m.dataset.no,10),r=C.find(d=>d.no===i);r&&x(r)})}),D.querySelectorAll(".btn-delete-chem").forEach(m=>{m.addEventListener("click",()=>{let i=parseInt(m.dataset.no,10),r=C.find(d=>d.no===i);r&&me(`Apakah Anda yakin ingin menghapus chemical: "${r.name}"?`,()=>{C=C.filter(d=>d.no!==i),C.forEach((d,b)=>{d.no=b+1}),Da(C),ae("Bahan chemical berhasil dihapus"),T()},"Hapus Bahan Chemical")})})}}W(),g.querySelector("#view-tools-btn").addEventListener("click",()=>{h="tools",M(g)}),g.querySelector("#view-chem-btn").addEventListener("click",()=>{h="chemical",M(g)}),O&&O.addEventListener("input",J=>{p=J.target.value,W()}),g.querySelector("#btn-add-tool-chem").addEventListener("click",()=>{h==="tools"?k():x()})}function k(g=null){if(!n)return;let D=!!g,O=`
      <form id="tool-edit-form" style="display:flex; flex-direction:column; gap:14px; font-size:0.88rem;">
        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Nama Alat / Item *</label>
          <input type="text" id="form-tool-name" class="form-control" required value="${g?f(g.name):""}" placeholder="Contoh: Lobby Duster Set 60 cm" style="width:100%; border-radius:6px; height:36px;">
        </div>
        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Kuantitas Standar *</label>
          <input type="text" id="form-tool-qty" class="form-control" required value="${g?f(g.qty):"1 unit"}" placeholder="Contoh: 1 set, 2 pcs, 1 paket" style="width:100%; border-radius:6px; height:36px;">
        </div>
        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Peruntukan Area & Kegunaan *</label>
          <textarea id="form-tool-area" class="form-control" rows="3" required placeholder="Contoh: Lobby utama dan koridor/lorong klinik" style="width:100%; border-radius:6px; font-size:0.85rem;">${g?f(g.area):""}</textarea>
        </div>
      </form>
    `;pe({title:D?`\u270F\uFE0F Edit Alat: ${g.name}`:"\u2795 Tambah Alat Kebersihan Baru",content:O,confirmText:"Simpan",onConfirm:(W,J)=>{let $=W.querySelector("#form-tool-name").value.trim(),m=W.querySelector("#form-tool-qty").value.trim(),i=W.querySelector("#form-tool-area").value.trim();if(!$||!m||!i){z("Harap lengkapi semua kolom");return}if(D)g.name=$,g.qty=m,g.area=i,La(E),ae(`Alat "${$}" berhasil diperbarui`);else{let r=E.length>0?Math.max(...E.map(d=>d.no))+1:1;E.push({no:r,name:$,qty:m,area:i}),La(E),ae(`Alat baru "${$}" berhasil ditambahkan`)}J(),T()}})}function x(g=null){if(!n)return;let D=!!g,O=`
      <form id="chem-edit-form" style="display:flex; flex-direction:column; gap:14px; font-size:0.88rem;">
        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Nama Produk Chemical *</label>
          <input type="text" id="form-chem-name" class="form-control" required value="${g?f(g.name):""}" placeholder="Contoh: Floor Cleaner Prokleen Lavender" style="width:100%; border-radius:6px; height:36px;">
        </div>
        <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-weight:600; margin-bottom:4px;">Kemasan Standar *</label>
            <input type="text" id="form-chem-size" class="form-control" required value="${g?f(g.size):"5 Liter"}" placeholder="Contoh: 5 Liter, 500 gram" style="width:100%; border-radius:6px; height:36px;">
          </div>
          <div>
            <label style="display:block; font-weight:600; margin-bottom:4px;">Rasio & Takaran Standar *</label>
            <input type="text" id="form-chem-ratio" class="form-control" required value="${g?f(g.ratio):""}" placeholder="Contoh: 25 ml : 1 Liter Air" style="width:100%; border-radius:6px; height:36px;">
          </div>
        </div>
        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Fungsi & Cara Penggunaan *</label>
          <textarea id="form-chem-function" class="form-control" rows="3" required placeholder="Contoh: Pembersihan dan pengepelan rutin seluruh lantai ruangan" style="width:100%; border-radius:6px; font-size:0.85rem;">${g?f(g.function):""}</textarea>
        </div>
      </form>
    `;pe({title:D?`\u270F\uFE0F Edit Chemical: ${g.name}`:"\u2795 Tambah Chemical Baru",content:O,confirmText:"Simpan",onConfirm:(W,J)=>{let $=W.querySelector("#form-chem-name").value.trim(),m=W.querySelector("#form-chem-size").value.trim(),i=W.querySelector("#form-chem-ratio").value.trim(),r=W.querySelector("#form-chem-function").value.trim();if(!$||!m||!i||!r){z("Harap lengkapi semua kolom");return}if(D)g.name=$,g.size=m,g.ratio=i,g.function=r,Da(C),ae(`Chemical "${$}" berhasil diperbarui`);else{let d=C.length>0?Math.max(...C.map(b=>b.no))+1:1;C.push({no:d,name:$,size:m,ratio:i,function:r}),Da(C),ae(`Chemical baru "${$}" berhasil ditambahkan`)}J(),T()}})}function _(g){g.innerHTML=`
      <div style="display:grid; grid-template-columns:repeat(auto-fit, minmax(360px, 1fr)); gap:20px;">
        <!-- Card 1: Grooming -->
        <div class="card" style="background:#172033; border:1px solid var(--border); border-radius:12px; padding:20px; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
          <div style="display:flex; align-items:center; gap:10px; margin-bottom:14px; border-bottom:1px solid #F1F5F9; padding-bottom:10px;">
            <span style="font-size:1.5rem; background:#EFF6FF; padding:8px; border-radius:8px;">\u{1F454}</span>
            <div>
              <h3 style="margin:0; font-size:1.1rem; font-weight:700; color:var(--text-1);">Standar Grooming Petugas (6.1)</h3>
              <span style="font-size:0.8rem; color:var(--text-3);">8 Persyaratan Wajib Penampilan Facility Care</span>
            </div>
          </div>
          <ol style="margin:0 0 0 18px; padding:0; font-size:0.88rem; color:var(--text-2); line-height:1.8;">
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
        <div class="card" style="background:#172033; border:1px solid var(--border); border-radius:12px; padding:20px; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
          <div style="display:flex; align-items:center; gap:10px; margin-bottom:14px; border-bottom:1px solid #F1F5F9; padding-bottom:10px;">
            <span style="font-size:1.5rem; background:#FEF3C7; padding:8px; border-radius:8px;">\u23F0</span>
            <div>
              <h3 style="margin:0; font-size:1.1rem; font-weight:700; color:var(--text-1);">Pengaturan Jam Kerja & Shift (6.2)</h3>
              <span style="font-size:0.8rem; color:var(--text-3);">Sistem Kerja 6 Hari Kerja : 1 Hari Libur</span>
            </div>
          </div>
          <div style="font-size:0.88rem; color:var(--text-2); line-height:1.7;">
            <div style="background:#111827; border-radius:8px; padding:10px 14px; margin-bottom:10px; border-left:3px solid #2563EB;">
              <strong style="color:var(--text-1); display:block; margin-bottom:2px;">\u{1F3E5} Klinik Primer (1 Orang Facility Care):</strong>
              <span>Shift 1 (Normal): <strong>08.00 s/d 17.00</strong> (Istirahat 13.00 s/d 14.00)</span>
            </div>
            <div style="background:#111827; border-radius:8px; padding:10px 14px; margin-bottom:12px; border-left:3px solid #059669;">
              <strong style="color:var(--text-1); display:block; margin-bottom:2px;">\u2728 Klinik Luxury (> 1 Orang Facility Care):</strong>
              <div style="margin-bottom:2px;">\u2022 Shift 1 (Pagi): <strong>07.00 s/d 16.00</strong> (Istirahat 12.00 s/d 13.00)</div>
              <div>\u2022 Shift 2 (Siang): <strong>13.00 s/d 22.00</strong> (Istirahat 17.00 s/d 18.00)</div>
            </div>
            <p style="margin:0; font-size:0.82rem; color:var(--text-3);">
              <strong>Ketentuan Backup:</strong> Saat Facility Care cabang libur, tim cabang wajib melakukan backup pekerjaan yang ditentukan oleh Supervisor (SPV) Cabang.
            </p>
          </div>
        </div>

        <!-- Card 3: Tanggung Jawab & Pembagian Area -->
        <div class="card" style="background:#172033; border:1px solid var(--border); border-radius:12px; padding:20px; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
          <div style="display:flex; align-items:center; gap:10px; margin-bottom:14px; border-bottom:1px solid #F1F5F9; padding-bottom:10px;">
            <span style="font-size:1.5rem; background:#ECFDF5; padding:8px; border-radius:8px;">\u{1F3E2}</span>
            <div>
              <h3 style="margin:0; font-size:1.1rem; font-weight:700; color:var(--text-1);">Tanggung Jawab & Area (6.3)</h3>
              <span style="font-size:0.8rem; color:var(--text-3);">Pembagian Pelaksana & Penanggung Jawab Area</span>
            </div>
          </div>
          <table style="width:100%; border-collapse:collapse; font-size:0.85rem; margin-bottom:14px;">
            <thead>
              <tr style="background:#111827; border-bottom:1px solid var(--border); text-align:left;">
                <th style="padding:6px 10px;">Area Kerja</th>
                <th style="padding:6px 10px;">Pelaksana</th>
                <th style="padding:6px 10px;">PIC Pengawas</th>
              </tr>
            </thead>
            <tbody>
              <tr style="border-bottom:1px solid #F1F5F9;">
                <td style="padding:8px 10px; font-weight:600;">HQ & Gudang Pusat</td>
                <td style="padding:8px 10px;">Facility Care Pusat</td>
                <td style="padding:8px 10px; color:#2563EB; font-weight:600;">FC Supervisor</td>
              </tr>
              <tr>
                <td style="padding:8px 10px; font-weight:600;">Klinik / Cabang</td>
                <td style="padding:8px 10px;">Facility Care Cabang</td>
                <td style="padding:8px 10px; color:#059669; font-weight:600;">SPV Cabang</td>
              </tr>
            </tbody>
          </table>
          <div style="font-size:0.83rem; color:var(--text-2); line-height:1.6;">
            <div style="margin-bottom:6px;"><strong>Deep Cleaning:</strong> Pembersihan total eventual pasca renovasi, pembukaan cabang baru, atau relokasi cabang lama.</div>
            <div><strong>General Cleaning (GC):</strong> Pembersihan menyeluruh berkala untuk cabang eksisting sesuai jadwal yang telah ditetapkan.</div>
          </div>
        </div>

        <!-- Card 4: Monitoring & Evaluasi -->
        <div class="card" style="background:#172033; border:1px solid var(--border); border-radius:12px; padding:20px; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
          <div style="display:flex; align-items:center; gap:10px; margin-bottom:14px; border-bottom:1px solid #F1F5F9; padding-bottom:10px;">
            <span style="font-size:1.5rem; background:#F5F3FF; padding:8px; border-radius:8px;">\u{1F50D}</span>
            <div>
              <h3 style="margin:0; font-size:1.1rem; font-weight:700; color:var(--text-1);">Sistem Monitoring & Inspeksi (6.4)</h3>
              <span style="font-size:0.8rem; color:var(--text-3);">Mekanisme Evaluasi Kinerja Harian & Triwulanan</span>
            </div>
          </div>
          <div style="font-size:0.88rem; color:var(--text-2); line-height:1.7;">
            <div style="margin-bottom:10px;">
              <strong style="color:var(--text-1);">1. Monitor Harian & Mingguan:</strong>
              <div style="font-size:0.84rem; color:var(--text-2); margin-top:2px;">
                Ceklist harian diisi 2 kali sehari (Pagi & Siang) dan pekerjaan mingguan diisi seminggu sekali. Wajib ditandatangani oleh SPV Cabang / FC Supervisor.
              </div>
            </div>
            <div>
              <strong style="color:var(--text-1);">2. Inspeksi Cabang Triwulanan (3 Bulan):</strong>
              <div style="font-size:0.84rem; color:var(--text-2); margin-top:2px;">
                Dilakukan oleh FC Specialist langsung ke lokasi klinik. Laporan inspeksi diunggah ke Basecamp dan wajib diberikan tindak lanjut oleh PIC cabang maksimal <strong>3 hari</strong> setelah laporan terbit. Hasil penilaian mempengaruhi performance cabang.
              </div>
            </div>
          </div>
        </div>
      </div>
    `}function B(g){g.innerHTML=`
      <div style="display:grid; grid-template-columns:repeat(auto-fit, minmax(320px, 1fr)); gap:18px;">
        <!-- Card: Ceklist Harian Google Sheets -->
        <div class="card" style="background:#172033; border:1px solid var(--border); border-radius:12px; padding:20px; box-shadow:0 1px 3px rgba(0,0,0,0.03); display:flex; flex-direction:column; justify-content:space-between;">
          <div>
            <div style="display:flex; align-items:center; gap:10px; margin-bottom:12px;">
              <span style="font-size:2rem;">\u{1F4CA}</span>
              <div>
                <h3 style="margin:0; font-size:1.1rem; font-weight:700; color:var(--text-1);">Form Ceklist Harian & Mingguan</h3>
                <span class="badge" style="background:#E0F2FE; color:#0369A1; font-size:0.75rem; font-weight:700;">Google Sheets</span>
              </div>
            </div>
            <p style="font-size:0.85rem; color:var(--text-2); line-height:1.6; margin-bottom:16px;">
              Format lembar ceklist kebersihan harian (Shift Pagi & Siang) dan ceklist pembersihan berkala mingguan yang wajib diisi oleh Facility Care dan disetujui Supervisor Cabang.
            </p>
          </div>
          <a href="https://docs.google.com/spreadsheets/d/1alq6ppT7syB8Q0_RjjeqwYhjeHgrWvVe-KJ3dHXWrss/edit?usp=sharing" target="_blank" rel="noopener" class="btn btn-primary" style="display:inline-flex; align-items:center; justify-content:center; gap:6px; font-weight:600; padding:10px; border-radius:8px;">
            \u{1F4C4} Buka Google Sheets Ceklist \u2794
          </a>
        </div>

        <!-- Card: Form Inspeksi Hygiene Google Forms -->
        <div class="card" style="background:#172033; border:1px solid var(--border); border-radius:12px; padding:20px; box-shadow:0 1px 3px rgba(0,0,0,0.03); display:flex; flex-direction:column; justify-content:space-between;">
          <div>
            <div style="display:flex; align-items:center; gap:10px; margin-bottom:12px;">
              <span style="font-size:2rem;">\u{1F4D1}</span>
              <div>
                <h3 style="margin:0; font-size:1.1rem; font-weight:700; color:var(--text-1);">Form Inspeksi Hygiene Cabang</h3>
                <span class="badge" style="background:#F3E8FF; color:#7E22CE; font-size:0.75rem; font-weight:700;">Google Forms</span>
              </div>
            </div>
            <p style="font-size:0.85rem; color:var(--text-2); line-height:1.6; margin-bottom:16px;">
              Formulir digital inspeksi evaluasi kebersihan dan kepatuhan hygiene cabang yang diisi oleh Facility Care Specialist setiap periode 3 bulanan.
            </p>
          </div>
          <a href="https://docs.google.com/forms/d/e/1FAIpQLScn8iPabIt51E_GORkbD86IURSS-EsUYINEQ_Zwdq_uC3XOCQ/viewform" target="_blank" rel="noopener" class="btn btn-primary" style="display:inline-flex; align-items:center; justify-content:center; gap:6px; font-weight:600; padding:10px; border-radius:8px; background:#7E22CE; border-color:#7E22CE;">
            \u{1F4CB} Buka Google Form Inspeksi \u2794
          </a>
        </div>

        <!-- Card: Basecamp Monitoring Project -->
        <div class="card" style="background:#172033; border:1px solid var(--border); border-radius:12px; padding:20px; box-shadow:0 1px 3px rgba(0,0,0,0.03); display:flex; flex-direction:column; justify-content:space-between;">
          <div>
            <div style="display:flex; align-items:center; gap:10px; margin-bottom:12px;">
              <span style="font-size:2rem;">\u{1F3D5}\uFE0F</span>
              <div>
                <h3 style="margin:0; font-size:1.1rem; font-weight:700; color:var(--text-1);">Laporan Inspeksi Basecamp</h3>
                <span class="badge" style="background:#FEF3C7; color:#92400E; font-size:0.75rem; font-weight:700;">Basecamp 3 Project</span>
              </div>
            </div>
            <p style="font-size:0.85rem; color:var(--text-2); line-height:1.6; margin-bottom:16px;">
              To-do list dan kanal pelaporan temuan inspeksi cabang. PIC Cabang wajib memberikan respon tindak lanjut maksimal 3 hari kerja setelah laporan terbit.
            </p>
          </div>
          <a href="https://3.basecamp.com/5259614/buckets/33617501/todosets/6367925585" target="_blank" rel="noopener" class="btn btn-outline" style="display:inline-flex; align-items:center; justify-content:center; gap:6px; font-weight:600; padding:10px; border-radius:8px; background:#1B2940; border:1px solid var(--border); color:#F8FAFC;">
            \u{1F680} Buka Laporan di Basecamp \u2794
          </a>
        </div>
      </div>

      <!-- Validasi & Pengesahan SOP -->
      <div class="card" style="margin-top:20px; padding:20px; background:#172033; border:1px solid var(--border); border-radius:12px;">
        <h3 style="margin:0 0 14px 0; font-size:1rem; font-weight:700; color:var(--text-1); display:flex; align-items:center; gap:8px;">
          <span>\u{1F4DC}</span> Validasi & Pengesahan Dokumen SOP
        </h3>
        <div style="display:grid; grid-template-columns:repeat(auto-fit, minmax(260px, 1fr)); gap:16px;">
          <div style="border-left:3px solid #2563EB; padding-left:12px;">
            <div style="font-size:0.75rem; color:var(--text-3); font-weight:700; text-transform:uppercase;">Disiapkan Oleh:</div>
            <div style="font-size:0.95rem; font-weight:700; color:var(--text-1); margin-top:2px;">Berlin Ariansyah</div>
            <div style="font-size:0.8rem; color:var(--text-2);">Facility Care Supervisor</div>
          </div>
          <div style="border-left:3px solid #059669; padding-left:12px;">
            <div style="font-size:0.75rem; color:var(--text-3); font-weight:700; text-transform:uppercase;">Diketahui Oleh:</div>
            <div style="font-size:0.95rem; font-weight:700; color:var(--text-1); margin-top:2px;">Catur Meiyogo</div>
            <div style="font-size:0.8rem; color:var(--text-2);">General Support Manager</div>
          </div>
          <div style="border-left:3px solid #7C3AED; padding-left:12px;">
            <div style="font-size:0.75rem; color:var(--text-3); font-weight:700; text-transform:uppercase;">Disetujui Oleh:</div>
            <div style="font-size:0.95rem; font-weight:700; color:var(--text-1); margin-top:2px;">Drg. Yulita</div>
            <div style="font-size:0.8rem; color:var(--text-2);">Chief Executive Officer</div>
          </div>
        </div>
      </div>
    `}async function P(){try{let{downloadExcel:g}=await Promise.resolve().then(()=>(we(),Oe)),D=c.map(W=>({"No IK":`IK-${W.id}`,"Judul Pekerjaan":W.title,Kategori:W.category,"Sasaran Target Area":W.target,"Rasio Chemical":W.chemical_ratio||"-","Alat & Bahan":W.tools.join(", "),"Langkah Kerja":W.steps.map((J,$)=>`${$+1}. ${J}`).join(" | "),"Catatan Safety":W.notes||"-"})),O=new Date().toISOString().slice(0,10);g(D,`SOP_Facility_Care_Instruksi_Kerja_${O}`),ae("Data Excel Instruksi Kerja berhasil diunduh")}catch(g){console.error(g),z("Gagal mendownload data Excel: "+g.message)}}function f(g){return g?String(g).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;").replace(/'/g,"&#39;"):""}T()}Fe();var Er=[{id:1,title:"Pengamanan Area & Aset Perusahaan",category:"Pengamanan Fisik & Patroli",icon:"\u{1F6E1}\uFE0F",color:"#2563EB",bg:"#EFF6FF",target:"Seluruh area dalam & luar Head Quarter dan Klinik Cabang Audy Dental",tools:["Google Form Checklist Monitoring Klinik: https://forms.gle/1EAvkqDJVdctXEGn7","Senter Patroli","Komputer dengan Akses Monitor CCTV (Live View)","Alat Komunikasi / Handphone / HT"],steps:["Lakukan Safety Patrol secara rutin mengelilingi seluruh area sesuai jadwal yang telah ditentukan (Patrol 1, 2, dan 3).","Catat hasil temuan dan kondisi keamanan secara berkala ke dalam Google Form Safety Patrol resmi.","Awasi kendaraan yang terparkir di area luar klinik; pastikan khusus sepeda motor dalam keadaan terkunci ganda.","Pantau situasi lingkungan dan area sensitif secara berkala melalui layar monitor CCTV saat berada di pos security.","Segera lakukan tindakan persuasif ramah apabila ditemukan potensi gangguan keamanan, dan segera hubungi pihak Kepolisian terdekat bila situasi darurat tidak dapat diatasi.","Bila ditemukan permasalahan fisik, kerusakan fasilitas, atau kendala di area klinik, segera dokumentasikan (foto) dan laporkan langsung ke SPV Klinik dan Leader Facility Care.","Disaat malam hari atau saat klinik telah tutup operasional, pastikan seluruh AC dan lampu ruangan dalam keadaan mati, kecuali lampu keamanan tertentu yang memang harus dibiarkan menyala terang."],notes:"Prioritaskan keselamatan diri dan tindakan persuasif. Dilarang meninggalkan area pengawasan tanpa personil pengganti yang jelas."},{id:2,title:"Penertiban & Pengaturan Parkir Kendaraan",category:"Operasional Parkir & Luar",icon:"\u{1F17F}\uFE0F",color:"#D97706",bg:"#FFFBEB",target:"Area parkir mobil, motor, akses keluar masuk, dan teras depan klinik",tools:["Lampu Lalu Lintas (Lampu Lalin)","Cone / Kerucut Pembatas Parkir","Peluit Parkir"],steps:["Berdiri di posisi strategis dan bantu mengarahkan setiap kendaraan pasien atau tamu yang datang secara tepat dan rapi ke slot parkir yang tersedia.","Berikan informasi dan solusi dengan sikap ramah serta santun apabila seluruh slot parkir klinik dalam kondisi penuh.","Dalam kondisi tertentu (bila ada mobil yang terhalang mobil lain dan harus dipindahkan, sementara pasien sedang dalam tindakan perawatan): bantu memindahkan kendaraan HANYA JIKA petugas security memiliki kemampuan mengemudi dan memiliki SIM aktif.","Apabila petugas security tidak memiliki kemampuan mengemudikan mobil: sampaikan permohonan maaf dengan tulus kepada pasien/tamu, minta untuk menunggu sejenak, dan segera minta bantuan kepada staf Resepsionis untuk menghubungi pemilik mobil yang menghalangi.","Bantu pasien atau tamu yang akan keluar parkir dengan memberikan aba-aba dan arahan lalu lintas jalan raya yang aman dan tepat."],notes:"Dilarang memaksakan memindahkan mobil pasien jika tidak memiliki lisensi mengemudi demi menghindari risiko gesekan atau kerusakan kendaraan."},{id:3,title:"Penerimaan Tamu & Pasien Klinik",category:"Layanan & Hospitalitas",icon:"\u{1F91D}",color:"#059669",bg:"#ECFDF5",target:"Pintu masuk utama klinik, lobby, area resepsionis, dan pos penerimaan tamu",tools:["Alat Tulis Kantor (Pulpen)","Buku Tamu / Log Book Keamanan","Alat Komunikasi / Handphone"],steps:['Berdiri dengan tegap dan sigap, bukakan pintu untuk setiap pasien/tamu yang datang dengan menerapkan budaya 5S (Senyum, Salam, Sapa, Sopan, Santun) dan menyapa: "Selamat pagi / siang / sore, ada yang bisa kami bantu?".',"Setelah pasien atau tamu memberitahukan tujuannya, persilakan dan arahkan pasien langsung menuju meja Resepsionis.","Apabila yang datang adalah tamu instansi, vendor, atau pihak Manajemen/BOD: persilakan mengisi Buku Tamu (Log Book) terlebih dahulu dan segera hubungi staf/karyawan yang ingin ditemui.",'Apabila pasien langsung melangkah ke arah tujuannya, berikan gestur tubuh hormat menyilangkan kedua tangan di depan, dan jika pasien/tamu mengucapkan terima kasih, jawab dengan ramah: "Dengan Senang Hati".',"Dalam kondisi khusus (pasien lansia, ibu hamil, atau pengguna kursi roda): petugas security wajib mendampingi, mengawal, dan membantu mendorong kursi roda sampai ke depan meja resepsionis.",'Buka kembali pintu untuk setiap pasien yang telah selesai melakukan perawatan saat hendak keluar klinik sambil mengucapkan ucapan hangat: "Terima Kasih, semoga lekas sehat kembali".'],notes:"Jaga senyum tulus dan keramahan sebagai garda terdepan representasi citra Audy Dental."}],Cr=[{no:1,type:"Aset Tetap",item:"Rumah & Bangunan",description:"Gedung klinik, pos satpam, pagar, kanopi, teras beton",pic:"Building Maintenance & Security"},{no:2,type:"Aset Tetap",item:"Genset & Lift",description:"Mesin generator listrik cadangan dan elevator penumpang tetap",pic:"Teknisi & Security"},{no:3,type:"Aset Tetap",item:"AC Central & Split Tetap",description:"Outdoor dan indoor AC yang terpasang permanen di dinding/plafon",pic:"Teknisi AC & Security"},{no:4,type:"Aset Tetap",item:"Dental Unit Permanen",description:"Kursi dental unit terpasang instalasi air dan kompresor",pic:"Tim Klinik & Security"},{no:5,type:"Aset Tetap",item:"Instalasi Listrik, Pipa & Jaringan IT Permanen",description:"Panel MCB, kabel fiber optik, pipa drainase, server rack",pic:"IT & Teknisi"},{no:6,type:"Aset Tetap",item:"Furniture Melekat (Built-in)",description:"Meja resepsionis cor/melekat, partisi kaca mati, kitchen set pantry",pic:"General Support"},{no:7,type:"Aset Bergerak",item:"Komputer, Laptop & Printer",description:"PC kasir/resepsionis, laptop staf, printer rontgen/dokumen",pic:"Staf Terkait & IT"},{no:8,type:"Aset Bergerak",item:"Kendaraan Operasional",description:"Mobil operasional dan sepeda motor inventaris perusahaan",pic:"Security & Driver"},{no:9,type:"Aset Bergerak",item:"Furniture Lepasan",description:"Meja kerja, kursi resepsionis, kursi tunggu pasien, lemari file",pic:"Facility Care & Staf"},{no:10,type:"Aset Bergerak",item:"Hand Tools, Alat Kerja & Alat Dental Portabel",description:"Handpiece dental, alat rontgen portable, toolbox teknisi",pic:"Dokter/Perawat & Teknisi"},{no:11,type:"Low Value Asset",item:"Peralatan Kantor (Telepon, ATK)",description:"Pesawat telepon, kalkulator, stapler, perlengkapan meja",pic:"Resepsionis"},{no:12,type:"Low Value Asset",item:"Alat-alat Kebersihan",description:"Lobby duster, ember wringer pel, sapu, tangga, botol spray",pic:"Facility Care"},{no:13,type:"Low Value Asset",item:"Perlengkapan Security",description:"Lampu lalin, cone parkir, peluit, senter patroli, jas hujan",pic:"Security"}],Jn={cabang:[{zone:"Zona Hijau",level:"Akses Publik Bebas",color:"#059669",bg:"#ECFDF5",rooms:["Area Parkir Kendaraan","Lobby Lantai 1","Teras Luar Klinik"],desc:"Area dengan pengamanan umum yang bebas diakses pasien dan tamu setelah pengarahan security."},{zone:"Zona Kuning",level:"Akses Terbatas / Pendampingan",color:"#D97706",bg:"#FFFBEB",rooms:["Koridor Ruang Tindakan","Tangga & Lift Klinik","Ruang Tunggu Lantai 2 & 3","Kids Playground","Toilet Lantai 1 & 2","Mushola Staf & Pasien"],desc:"Area terbatas, tamu wajib didampingi petugas atau staf klinik terkait."},{zone:"Zona Merah",level:"Area Terlarang / Restricted",color:"#DC2626",bg:"#FEF2F2",rooms:["Area Meja Receptionist / Kasir","Seluruh Ruang Dental Unit (DU 101, 102, 201, 202, 301, 302)","Ruang Sterilisasi Alat","Ruang Rontgen Radiologi","Ruang Suction & Kompresor","Ruang Tim Staf & Gudang Material Medis"],desc:"Area sangat sensitif, hanya staf medis bertugas dan atasan pemilik area. Tamu wajib izin Manajemen/BOD dan didampingi penuh."}],hq:[{floor:"Lantai 1",hijau:["Area Parkir Mobil/Motor","Lobby Resepsionis Lantai 1"],kuning:["Seluruh Ruang Meeting Lantai 1","Ruang Training Lantai 1","Koridor Lift & Toilet"],merah:["Ruang Customer Services","Gudang Utama Lantai 1","Gudang Facility Care","Gudang Teknisi & Ruang Teknisi","Pos Satpam & Ruang Genset"]},{floor:"Lantai 2",hijau:["-"],kuning:["Ruang Meeting Lantai 2","Koridor Lift & Toilet","Pantry"],merah:["Ruang CEO","Ruang Sekretaris CEO","Ruang Human Capital","Ruang Dentist Relation","Ruang General Support","Ruang Eks. Patrel"]},{floor:"Lantai 3",hijau:["-"],kuning:["Koridor Lift & Toilet","Pantry","Mushola & Tempat Wudhu","Ruang Santai Karyawan"],merah:["Ruang Marketing","Studio Konten Marketing","Gudang Marketing","Ruang Patient Relation","Ruang OPS","Ruang Training Lantai 3","Gudang Lantai 3"]},{floor:"Lantai 4",hijau:["-"],kuning:["Koridor Lift & Toilet","Pantry"],merah:["Ruang Finance & Accounting","Ruang VP Finance & Accounting","Ruang Audit Internal","Ruang Training Lantai 4","Gudang Lantai 4"]},{floor:"Lantai 5",hijau:["-"],kuning:["Lobby Lantai 5","Pantry Lantai 5","Toilet Lantai 5","Ruang Training Besar Lantai 5"],merah:["Area Dak Toren Air & Rooftop"]}]},st="fcms_sop_sec_ik_v1",lt="fcms_sop_sec_assets_v1";function Vn(){try{let e=localStorage.getItem(st);if(e)return JSON.parse(e)}catch(e){console.error("Error load IK Sec",e)}return JSON.parse(JSON.stringify(Er))}function Ia(e){localStorage.setItem(st,JSON.stringify(e)),X("/api/sop/bundle/"+st,{method:"PUT",body:JSON.stringify(e)}).catch(()=>{})}function Yn(){try{let e=localStorage.getItem(lt);if(e)return JSON.parse(e)}catch(e){console.error("Error load assets sec",e)}return JSON.parse(JSON.stringify(Cr))}function Ma(e){localStorage.setItem(lt,JSON.stringify(e)),X("/api/sop/bundle/"+lt,{method:"PUT",body:JSON.stringify(e)}).catch(()=>{})}async function Wn(e){let t=de(),n=ie(t,"sop","write"),a=ie(t,"sop","delete"),o="ik",s="",l="all",h="",p="cabang";try{let[f,g]=await Promise.all([X("/api/sop/bundle/"+st),X("/api/sop/bundle/"+lt)]);f.ok&&f.data?.data&&localStorage.setItem(st,JSON.stringify(f.data.data)),g.ok&&g.data?.data&&localStorage.setItem(lt,JSON.stringify(g.data.data))}catch(f){console.warn("Could not load remote Security SOP bundle:",f)}let c=Vn(),E=Yn();function C(){e.innerHTML=`
      <div class="page-header" style="display:flex; justify-content:space-between; align-items:flex-start; flex-wrap:wrap; gap:16px; margin-bottom:20px;">
        <div>
          <h1 class="page-title" style="display:flex; align-items:center; gap:8px; margin:0 0 6px 0; font-size:1.6rem; font-weight:700; color:var(--text-1);">
            <span>\u{1F6E1}\uFE0F</span> SOP Pengelolaan Keamanan & Aset Perusahaan
          </h1>
          <p class="page-subtitle" style="margin:0; font-size:0.9rem; color:var(--text-3);">
            Standar Operasional Prosedur Pengawasan Keamanan Area, Fasilitas, dan Perlindungan Aset PT Audy Mandiri Indonesia
          </p>
        </div>
        <div style="display:flex; gap:10px; align-items:center; flex-wrap:wrap;">
          ${a?`
          <button id="btn-reset-sec-default" class="btn btn-outline" title="Kembalikan data ke dokumen standar asli" style="background:#172033; border:1px solid var(--border); color:#F8FAFC; font-weight:600; padding:8px 14px; border-radius:8px; display:inline-flex; align-items:center; gap:6px; cursor:pointer; font-size:0.85rem;">
            \u{1F504} Reset ke Standar
          </button>
          `:""}
          <button id="btn-print-sec" class="btn btn-outline" style="background:#172033; border:1px solid var(--border); color:#F8FAFC; font-weight:600; padding:8px 14px; border-radius:8px; display:inline-flex; align-items:center; gap:6px; cursor:pointer; font-size:0.85rem;">
            \u{1F5A8}\uFE0F Cetak Ringkasan SOP
          </button>
          ${ie(t,"export")?`<button id="btn-export-sec" class="btn btn-primary" style="font-weight:600; padding:8px 16px; border-radius:8px; display:inline-flex; align-items:center; gap:6px; cursor:pointer; font-size:0.85rem;">
            \u{1F4E5} Download Data Excel
          </button>`:""}
        </div>
      </div>

      <!-- Navigation Tabs -->
      <div style="display:flex; gap:8px; margin-bottom:22px; border-bottom:2px solid #E2E8F0; padding-bottom:8px; overflow-x:auto; flex-wrap:nowrap; -webkit-overflow-scrolling:touch; scrollbar-width:none;">
        <button class="sec-tab-btn ${o==="ik"?"active":""}" data-tab="ik" style="${T(o==="ik")}">
          \u{1F6E1}\uFE0F Instruksi Kerja (IK) <span class="badge" style="${u(o==="ik")}">${c.length}</span>
        </button>
        <button class="sec-tab-btn ${o==="zonasi"?"active":""}" data-tab="zonasi" style="${T(o==="zonasi")}">
          \u{1F5FA}\uFE0F Zonasi Akses & Denah
        </button>
        <button class="sec-tab-btn ${o==="assets"?"active":""}" data-tab="assets" style="${T(o==="assets")}">
          \u{1F4E6} Klasifikasi Aset <span class="badge" style="${u(o==="assets")}">${E.length}</span>
        </button>
        <button class="sec-tab-btn ${o==="policy"?"active":""}" data-tab="policy" style="${T(o==="policy")}">
          \u{1F4CB} Shift, Patrol & Insiden
        </button>
        <button class="sec-tab-btn ${o==="links"?"active":""}" data-tab="links" style="${T(o==="links")}">
          \u{1F517} Form & Tautan
        </button>
      </div>

      <!-- Content Container -->
      <div id="sec-content-area"></div>
    `,F(),R()}function T(f){return f?"border:none; background:#1E40AF; color:#fff; font-weight:700; padding:8px 16px; border-radius:8px; cursor:pointer; display:inline-flex; align-items:center; gap:6px; font-size:0.9rem; white-space:nowrap; transition:all 0.2s;":"border:none; background:#172033; color:var(--text-2); border:1px solid var(--border); font-weight:600; padding:8px 16px; border-radius:8px; cursor:pointer; display:inline-flex; align-items:center; gap:6px; font-size:0.9rem; white-space:nowrap; transition:all 0.2s;"}function u(f){return f?"background:rgba(255,255,255,0.25); color:#fff; font-size:0.75rem; padding:2px 7px; border-radius:10px; font-weight:700;":"background:#CBD5E1; color:#334155; font-size:0.75rem; padding:2px 7px; border-radius:10px; font-weight:700;"}function F(){e.querySelectorAll(".sec-tab-btn").forEach(O=>{O.addEventListener("click",()=>{o=O.dataset.tab,C()})});let f=e.querySelector("#btn-reset-sec-default");f&&f.addEventListener("click",()=>{me("Apakah Anda yakin ingin mengembalikan seluruh data Instruksi Kerja dan Klasifikasi Aset Security ke standar dokumen asli?",()=>{localStorage.removeItem(st),localStorage.removeItem(lt),Promise.all([X("/api/sop/bundle/"+st,{method:"DELETE"}),X("/api/sop/bundle/"+lt,{method:"DELETE"})]).catch(()=>{}),c=Vn(),E=Yn(),ae("Data SOP Security berhasil di-reset ke standar dokumen"),C()},"Reset ke Standar Dokumen")});let g=e.querySelector("#btn-print-sec");g&&g.addEventListener("click",()=>{window.print()});let D=e.querySelector("#btn-export-sec");D&&D.addEventListener("click",()=>{B()})}function R(){let f=e.querySelector("#sec-content-area");f&&(o==="ik"?K(f):o==="zonasi"?w(f):o==="assets"?M(f):o==="policy"?x(f):o==="links"&&_(f))}function K(f){f.innerHTML=`
      <!-- Toolbar Filter & Search -->
      <div class="card" style="padding:14px 18px; margin-bottom:20px; background:#172033; border:1px solid var(--border); border-radius:12px; box-shadow:0 1px 3px rgba(0,0,0,0.2);">
        <div style="display:flex; flex-wrap:wrap; gap:12px; align-items:center; justify-content:space-between;">
          <div style="display:flex; flex-wrap:wrap; gap:12px; align-items:center; flex:1; min-width:280px;">
            <div style="position:relative; flex:1; min-width:220px; max-width:340px;">
              <input type="text" id="sec-ik-search" class="form-control" placeholder="Cari instruksi kerja security..." value="${s}" style="padding-left:34px; height:38px; border-radius:8px; font-size:0.9rem; width:100%;">
              <span style="position:absolute; left:11px; top:50%; transform:translateY(-50%); color:var(--text-3); font-size:0.9rem;">\u{1F50D}</span>
            </div>
            
            <div style="display:flex; gap:6px; flex-wrap:wrap;" id="sec-cat-pills-wrap">
              <button class="sec-cat-pill ${l==="all"?"active":""}" data-cat="all">Semua (${c.length})</button>
              <button class="sec-cat-pill ${l==="Pengamanan Fisik & Patroli"?"active":""}" data-cat="Pengamanan Fisik & Patroli">Patroli & Fisik (${c.filter(m=>m.category==="Pengamanan Fisik & Patroli").length})</button>
              <button class="sec-cat-pill ${l==="Operasional Parkir & Luar"?"active":""}" data-cat="Operasional Parkir & Luar">Parkir & Luar (${c.filter(m=>m.category==="Operasional Parkir & Luar").length})</button>
              <button class="sec-cat-pill ${l==="Layanan & Hospitalitas"?"active":""}" data-cat="Layanan & Hospitalitas">Tamu & Pasien (${c.filter(m=>m.category==="Layanan & Hospitalitas").length})</button>
            </div>
          </div>
          <div style="display:flex; gap:10px; align-items:center;">
            <div id="sec-ik-count-label" style="font-size:0.85rem; color:var(--text-3); font-weight:600;">
              Menampilkan ${c.length} Instruksi Kerja
            </div>
            ${n?`
            <button id="btn-add-sec-ik" class="btn btn-primary" style="font-weight:600; padding:7px 14px; border-radius:8px; display:inline-flex; align-items:center; gap:6px; cursor:pointer; font-size:0.85rem; background:#1E40AF; border-color:#1E40AF;">
              \u2795 Tambah IK Baru
            </button>
            `:""}
          </div>
        </div>
      </div>

      <!-- IK Cards Grid Container -->
      <div id="sec-ik-cards-grid-container"></div>
    `;let g=document.createElement("style");g.innerHTML=`
      .sec-cat-pill {
        border: 1px solid var(--border);
        background: #172033;
        color: var(--text-2);
        font-size: 0.8rem;
        font-weight: 600;
        padding: 5px 12px;
        border-radius: 20px;
        cursor: pointer;
        transition: all 0.15s;
        flex-shrink: 0;
        white-space: nowrap;
      }
      .sec-cat-pill:hover { background: #22314A; color:#F8FAFC; }
      .sec-cat-pill.active {
        background: #1E40AF;
        color: #fff;
        border-color: #1E40AF;
      }
      .sec-ik-card:hover {
        box-shadow: 0 6px 16px rgba(0,0,0,0.08) !important;
        transform: translateY(-2px);
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
    `,f.appendChild(g);let D=f.querySelector("#sec-ik-cards-grid-container"),O=f.querySelector("#sec-ik-count-label");function W(){let m=c.filter(i=>{let r=l==="all"||i.category===l,d=s.toLowerCase().trim(),b=!d||i.title.toLowerCase().includes(d)||i.target.toLowerCase().includes(d)||i.tools.some(y=>y.toLowerCase().includes(d))||i.steps.some(y=>y.toLowerCase().includes(d));return r&&b});if(O&&(O.textContent=`Menampilkan ${m.length} dari ${c.length} Instruksi Kerja`),m.length===0){D.innerHTML=`
          <div style="text-align:center; padding:48px 20px; background:#111827; border-radius:12px; border:1px solid var(--border);">
            <div style="font-size:2.5rem; margin-bottom:8px;">\u{1F50D}</div>
            <h3 style="margin:0 0 4px; font-size:1.1rem; color:var(--text-1);">Tidak ada Instruksi Kerja yang cocok</h3>
            <p style="margin:0; font-size:0.85rem; color:var(--text-3);">Coba ubah kata kunci pencarian atau pilih kategori lain.</p>
          </div>
        `;return}D.innerHTML=`
        <div class="sec-ik-grid" style="display:grid; grid-template-columns:repeat(auto-fill, minmax(340px, 1fr)); gap:18px;">
          ${m.map(i=>`
            <div class="card sec-ik-card" style="background:#172033; border:1px solid var(--border); border-radius:12px; padding:18px; display:flex; flex-direction:column; justify-content:space-between; box-shadow:0 1px 4px rgba(0,0,0,0.04); transition:all 0.2s;">
              <div>
                <div style="display:flex; justify-content:space-between; align-items:flex-start; margin-bottom:12px;">
                  <div style="display:flex; align-items:center; gap:10px;">
                    <span style="font-size:1.6rem; width:44px; height:44px; display:inline-flex; align-items:center; justify-content:center; background:${i.bg||"#EFF6FF"}; border-radius:10px;">${i.icon||"\u{1F6E1}\uFE0F"}</span>
                    <div>
                      <h3 style="margin:0 0 2px 0; font-size:1.05rem; font-weight:700; color:var(--text-1);">
                        IK-${String(i.id).padStart(2,"0")}. ${i.title}
                      </h3>
                      <span class="badge" style="background:${i.bg||"#EFF6FF"}; color:${i.color||"#1E40AF"}; font-size:0.75rem; font-weight:700; border-radius:6px; padding:2px 8px;">
                        ${i.category}
                      </span>
                    </div>
                  </div>

                  <!-- Quick Action Buttons: Edit & Hapus -->
                  <div style="display:flex; gap:4px;">
                    ${n?`
                    <button class="btn btn-sm btn-ghost btn-edit-sec-ik" data-id="${i.id}" title="Edit Instruksi Kerja" style="padding:4px 8px; border-radius:6px; font-size:0.9rem; cursor:pointer;">
                      \u270F\uFE0F
                    </button>
                    `:""}
                    ${a?`
                    <button class="btn btn-sm btn-ghost btn-delete-sec-ik" data-id="${i.id}" title="Hapus Instruksi Kerja" style="padding:4px 8px; border-radius:6px; font-size:0.9rem; color:#DC2626; cursor:pointer;">
                      \u{1F5D1}\uFE0F
                    </button>
                    `:""}
                  </div>
                </div>

                <div style="font-size:0.85rem; color:var(--text-2); margin-bottom:14px; background:#111827; padding:8px 12px; border-radius:8px; border-left:3px solid ${i.color||"#1E40AF"};">
                  <strong>Area Sasaran:</strong> ${i.target}
                </div>

                <div style="margin-bottom:14px;">
                  <div style="font-size:0.8rem; font-weight:700; color:var(--text-3); text-transform:uppercase; margin-bottom:6px;">\u{1F9F0} Alat & Perlengkapan:</div>
                  <ul style="margin:0 0 0 16px; padding:0; font-size:0.83rem; color:var(--text-2);">
                    ${i.tools.slice(0,3).map(r=>`<li style="margin-bottom:2px;">${P(r)}</li>`).join("")}
                    ${i.tools.length>3?`<li style="color:var(--text-3); font-style:italic;">+ ${i.tools.length-3} perlengkapan lainnya...</li>`:""}
                  </ul>
                </div>
              </div>

              <div style="border-top:1px solid #F1F5F9; padding-top:12px; margin-top:8px; display:flex; justify-content:space-between; align-items:center;">
                <span style="font-size:0.8rem; color:var(--text-3); font-weight:500;">
                  ${i.steps.length} Langkah Prosedur
                </span>
                <button class="btn btn-sm btn-primary btn-view-sec-ik" data-id="${i.id}" style="padding:6px 14px; font-weight:600; border-radius:6px; background:#1E40AF; border-color:#1E40AF;">
                  Lihat Panduan Lengkap \u2794
                </button>
              </div>
            </div>
          `).join("")}
        </div>
      `,D.querySelectorAll(".btn-view-sec-ik").forEach(i=>{i.addEventListener("click",()=>{let r=parseInt(i.dataset.id,10),d=c.find(b=>b.id===r);d&&I(d)})}),D.querySelectorAll(".btn-edit-sec-ik").forEach(i=>{i.addEventListener("click",()=>{let r=parseInt(i.dataset.id,10),d=c.find(b=>b.id===r);d&&N(d)})}),D.querySelectorAll(".btn-delete-sec-ik").forEach(i=>{i.addEventListener("click",()=>{let r=parseInt(i.dataset.id,10),d=c.find(b=>b.id===r);d&&me(`Apakah Anda yakin ingin menghapus Instruksi Kerja: "${d.title}"?`,()=>{c=c.filter(b=>b.id!==r),Ia(c),ae("Instruksi Kerja Security berhasil dihapus"),C()},"Hapus Instruksi Kerja")})})}W();let J=f.querySelector("#sec-ik-search");J&&J.addEventListener("input",m=>{s=m.target.value,W()}),f.querySelectorAll(".sec-cat-pill").forEach(m=>{m.addEventListener("click",()=>{f.querySelectorAll(".sec-cat-pill").forEach(i=>i.classList.remove("active")),m.classList.add("active"),l=m.dataset.cat,W()})});let $=f.querySelector("#btn-add-sec-ik");$&&$.addEventListener("click",()=>{N()})}function I(f){pe({title:`${f.icon||"\u{1F6E1}\uFE0F"} IK-${String(f.id).padStart(2,"0")}: ${f.title}`,size:"lg",content:`
        <div style="display:flex; flex-direction:column; gap:16px;">
          <div style="background:${f.bg||"#EFF6FF"}; color:${f.color||"#1E40AF"}; padding:10px 14px; border-radius:8px; font-weight:700; font-size:0.9rem; display:flex; justify-content:space-between; align-items:center;">
            <span>Kategori: ${f.category}</span>
            <span>Kode: IK-SEC-${String(f.id).padStart(2,"0")}</span>
          </div>

          <div style="background:#111827; padding:12px 14px; border-radius:8px; border-left:4px solid ${f.color||"#1E40AF"};">
            <strong style="color:var(--text-1); font-size:0.85rem; display:block; margin-bottom:4px;">\u{1F4CD} Area & Sasaran Target:</strong>
            <span style="font-size:0.88rem; color:var(--text-2);">${f.target}</span>
          </div>

          <div>
            <h4 style="margin:0 0 8px 0; font-size:0.95rem; font-weight:700; color:var(--text-1);">\u{1F9F0} Alat & Perlengkapan yang Disiapkan:</h4>
            <div style="background:#111827; border:1px solid var(--border); border-radius:8px; padding:10px 14px;">
              <ul style="margin:0 0 0 16px; padding:0; font-size:0.875rem; color:var(--text-2); line-height:1.7;">
                ${f.tools.map(g=>{if(g.includes("http")){let D=g.split(": ");return`<li>${D[0]}: <a href="${D.slice(1).join(": ")}" target="_blank" rel="noopener" style="color:#2563EB; font-weight:600; text-decoration:underline;">Buka Tautan Form \u2794</a></li>`}return`<li>${P(g)}</li>`}).join("")}
              </ul>
            </div>
          </div>

          <div>
            <h4 style="margin:0 0 8px 0; font-size:0.95rem; font-weight:700; color:var(--text-1);">\u{1F4DD} Langkah & Prosedur Pelaksanaan:</h4>
            <div style="background:#111827; border:1px solid var(--border); border-radius:8px; padding:12px 16px;">
              <ol style="margin:0 0 0 18px; padding:0; font-size:0.875rem; color:var(--text-2); line-height:1.8;">
                ${f.steps.map(g=>`<li style="margin-bottom:6px;">${P(g)}</li>`).join("")}
              </ol>
            </div>
          </div>

          ${f.notes?`
            <div style="background:#EFF6FF; border:1px solid #BFDBFE; border-radius:8px; padding:10px 14px;">
              <strong style="color:#1E40AF; font-size:0.85rem; display:block; margin-bottom:2px;">\u26A0\uFE0F Catatan Khusus & Ketentuan Penting:</strong>
              <span style="color:#1E3A8A; font-size:0.85rem; line-height:1.6;">${P(f.notes)}</span>
            </div>
          `:""}
        </div>
      `,confirmText:"Tutup Panduan",onConfirm:(g,D)=>{D()}})}function N(f=null){if(!n)return;let g=!!f,D=f?f.tools.join(`
`):"",O=f?f.steps.join(`
`):"",W=`
      <form id="sec-ik-edit-form" style="display:flex; flex-direction:column; gap:14px; font-size:0.88rem;">
        <div class="form-row" style="display:grid; grid-template-columns:1fr 140px; gap:12px;">
          <div>
            <label style="display:block; font-weight:600; margin-bottom:4px;">Judul Instruksi Kerja *</label>
            <input type="text" id="form-sec-ik-title" class="form-control" required value="${f?P(f.title):""}" placeholder="Contoh: Pengamanan Area & Aset Perusahaan" style="width:100%; border-radius:6px; height:36px;">
          </div>
          <div>
            <label style="display:block; font-weight:600; margin-bottom:4px;">Ikon Emoji</label>
            <input type="text" id="form-sec-ik-icon" class="form-control" value="${f?f.icon:"\u{1F6E1}\uFE0F"}" placeholder="\u{1F6E1}\uFE0F" style="width:100%; border-radius:6px; height:36px; text-align:center; font-size:1.2rem;">
          </div>
        </div>

        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Kategori *</label>
          <select id="form-sec-ik-cat" class="form-control" style="width:100%; border-radius:6px; height:36px;">
            <option value="Pengamanan Fisik & Patroli" ${f?.category==="Pengamanan Fisik & Patroli"?"selected":""}>Pengamanan Fisik & Patroli</option>
            <option value="Operasional Parkir & Luar" ${f?.category==="Operasional Parkir & Luar"?"selected":""}>Operasional Parkir & Luar</option>
            <option value="Layanan & Hospitalitas" ${f?.category==="Layanan & Hospitalitas"?"selected":""}>Layanan & Hospitalitas</option>
          </select>
        </div>

        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Area & Sasaran Target *</label>
          <input type="text" id="form-sec-ik-target" class="form-control" required value="${f?P(f.target):""}" placeholder="Contoh: Seluruh area luar klinik dan akses parkir" style="width:100%; border-radius:6px; height:36px;">
        </div>

        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Alat & Perlengkapan (1 per baris) *</label>
          <textarea id="form-sec-ik-tools" class="form-control" rows="3" placeholder="Senter Patroli&#10;Lampu Lalin&#10;Peluit" style="width:100%; border-radius:6px; font-size:0.85rem; line-height:1.5;">${P(D)}</textarea>
        </div>

        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Langkah & Prosedur Pelaksanaan (1 langkah per baris) *</label>
          <textarea id="form-sec-ik-steps" class="form-control" rows="5" placeholder="1. Lakukan Safety Patrol sesuai jadwal.&#10;2. Arahkan kendaraan pasien dengan ramah..." style="width:100%; border-radius:6px; font-size:0.85rem; line-height:1.6;">${P(O)}</textarea>
        </div>

        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Catatan Khusus / Ketentuan Penting</label>
          <textarea id="form-sec-ik-notes" class="form-control" rows="2" placeholder="Peringatan keselamatan atau pedoman khusus..." style="width:100%; border-radius:6px; font-size:0.85rem;">${f?P(f.notes||""):""}</textarea>
        </div>
      </form>
    `;pe({title:g?`\u270F\uFE0F Edit IK-${f.id}: ${f.title}`:"\u2795 Tambah Instruksi Kerja Security Baru",content:W,size:"lg",confirmText:"Simpan Perubahan",onConfirm:(J,$)=>{let m=J.querySelector("#form-sec-ik-title").value.trim(),i=J.querySelector("#form-sec-ik-icon").value.trim()||"\u{1F6E1}\uFE0F",r=J.querySelector("#form-sec-ik-cat").value,d=J.querySelector("#form-sec-ik-target").value.trim(),b=J.querySelector("#form-sec-ik-tools").value.trim(),y=J.querySelector("#form-sec-ik-steps").value.trim(),q=J.querySelector("#form-sec-ik-notes").value.trim();if(!m||!d||!b||!y){z("Harap lengkapi semua kolom bertanda bintang (*)");return}let v=b.split(`
`).map(H=>H.trim().replace(/^[-•*]\s*/,"")).filter(Boolean),S=y.split(`
`).map(H=>H.trim().replace(/^\d+[\.\)]\s*/,"")).filter(Boolean),U={"Pengamanan Fisik & Patroli":{color:"#2563EB",bg:"#EFF6FF"},"Operasional Parkir & Luar":{color:"#D97706",bg:"#FFFBEB"},"Layanan & Hospitalitas":{color:"#059669",bg:"#ECFDF5"}}[r]||{color:"#1E40AF",bg:"#EFF6FF"};if(g)f.title=m,f.icon=i,f.category=r,f.target=d,f.tools=v,f.steps=S,f.notes=q,f.color=U.color,f.bg=U.bg,Ia(c),ae(`Instruksi Kerja "${m}" berhasil diperbarui`);else{let H=c.length>0?Math.max(...c.map(Q=>Q.id))+1:1,ee={id:H,title:m,category:r,icon:i,color:U.color,bg:U.bg,target:d,tools:v,steps:S,notes:q};c.push(ee),Ia(c),ae(`Instruksi Kerja baru IK-${H} berhasil ditambahkan`)}$(),C()}})}function w(f){f.innerHTML=`
      <div style="display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:12px; margin-bottom:20px;">
        <div class="btn-group" style="background:#111827; padding:3px; border-radius:8px; display:inline-flex; border:1px solid var(--border);">
          <button id="view-cabang-btn" class="btn btn-sm ${p==="cabang"?"btn-primary":"btn-ghost"}" style="font-weight:700; border-radius:6px; padding:6px 16px; cursor:pointer;">
            \u{1F3E5} Zonasi Klinik / Cabang
          </button>
          <button id="view-hq-btn" class="btn btn-sm ${p==="hq"?"btn-primary":"btn-ghost"}" style="font-weight:700; border-radius:6px; padding:6px 16px; cursor:pointer;">
            \u{1F3E2} Zonasi Head Quarter (5 Lantai)
          </button>
        </div>
        <div style="font-size:0.85rem; color:var(--text-3); font-weight:500;">
          Ketentuan Pembagian Akses Ruangan (Lampiran 8.1 SOP)
        </div>
      </div>

      <!-- Definisi 3 Zona Header Cards -->
      <div style="display:grid; grid-template-columns:repeat(auto-fit, minmax(280px, 1fr)); gap:16px; margin-bottom:22px;">
        <div class="card" style="background:#FEF2F2; border:1px solid #FECACA; border-radius:10px; padding:14px;">
          <div style="display:flex; align-items:center; gap:8px; margin-bottom:6px;">
            <span style="width:14px; height:14px; border-radius:50%; background:#DC2626; display:inline-block;"></span>
            <strong style="color:#991B1B; font-size:0.95rem;">ZONA MERAH (Restricted Area)</strong>
          </div>
          <p style="margin:0; font-size:0.82rem; color:#7F1D1D; line-height:1.5;">
            Area sangat terbatas. Hanya personal yang bertugas spesifik di area tersebut dan BOD. Tamu WAJIB mendapat izin tertulis Manajemen dan didampingi penuh.
          </p>
        </div>

        <div class="card" style="background:#FFFBEB; border:1px solid #FDE68A; border-radius:10px; padding:14px;">
          <div style="display:flex; align-items:center; gap:8px; margin-bottom:6px;">
            <span style="width:14px; height:14px; border-radius:50%; background:#D97706; display:inline-block;"></span>
            <strong style="color:#92400E; font-size:0.95rem;">ZONA KUNING (Limited Area)</strong>
          </div>
          <p style="margin:0; font-size:0.82rem; color:#78350F; line-height:1.5;">
            Area dengan pengamanan terbatas. Karyawan umum atau tamu yang memasuki area ini perlu mendapatkan arahan dan didampingi staf area terkait.
          </p>
        </div>

        <div class="card" style="background:#ECFDF5; border:1px solid #A7F3D0; border-radius:10px; padding:14px;">
          <div style="display:flex; align-items:center; gap:8px; margin-bottom:6px;">
            <span style="width:14px; height:14px; border-radius:50%; background:#059669; display:inline-block;"></span>
            <strong style="color:#065F46; font-size:0.95rem;">ZONA HIJAU (Public Area)</strong>
          </div>
          <p style="margin:0; font-size:0.82rem; color:#064E3B; line-height:1.5;">
            Area umum dengan pengawasan security. Diperbolehkan untuk seluruh tamu dan pasien setelah diarahkan oleh petugas security.
          </p>
        </div>
      </div>

      <!-- Zonasi Content Area -->
      <div id="zonasi-content-container"></div>
    `;let g=f.querySelector("#zonasi-content-container");p==="cabang"?g.innerHTML=`
        <div style="display:grid; grid-template-columns:repeat(auto-fit, minmax(320px, 1fr)); gap:18px;">
          ${Jn.cabang.map(D=>`
            <div class="card" style="background:#172033; border:1px solid var(--border); border-radius:12px; padding:18px; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
              <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:12px; border-bottom:1px solid #F1F5F9; padding-bottom:8px;">
                <h3 style="margin:0; font-size:1.05rem; font-weight:700; color:${D.color}; display:flex; align-items:center; gap:8px;">
                  <span style="width:12px; height:12px; border-radius:50%; background:${D.color}; display:inline-block;"></span>
                  ${D.zone}
                </h3>
                <span class="badge" style="background:${D.bg}; color:${D.color}; font-size:0.75rem; font-weight:700; padding:3px 8px; border-radius:6px;">
                  ${D.level}
                </span>
              </div>
              <p style="font-size:0.82rem; color:var(--text-3); margin-bottom:14px; line-height:1.5;">${D.desc}</p>
              <div>
                <strong style="font-size:0.8rem; text-transform:uppercase; color:var(--text-2); display:block; margin-bottom:6px;">Daftar Ruangan Terkait:</strong>
                <ul style="margin:0 0 0 16px; padding:0; font-size:0.85rem; color:var(--text-1); line-height:1.7;">
                  ${D.rooms.map(O=>`<li>${O}</li>`).join("")}
                </ul>
              </div>
            </div>
          `).join("")}
        </div>
      `:g.innerHTML=`
        <div style="display:flex; flex-direction:column; gap:16px;">
          ${Jn.hq.map(D=>`
            <div class="card" style="background:#172033; border:1px solid var(--border); border-radius:12px; padding:18px; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
              <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:14px; border-bottom:2px solid #F1F5F9; padding-bottom:8px;">
                <h3 style="margin:0; font-size:1.1rem; font-weight:700; color:var(--text-1); display:flex; align-items:center; gap:8px;">
                  <span>\u{1F3E2}</span> Head Quarter \u2014 ${D.floor}
                </h3>
              </div>
              <div style="display:grid; grid-template-columns:repeat(auto-fit, minmax(260px, 1fr)); gap:14px;">
                <!-- Zona Hijau -->
                <div style="background:#ECFDF5; border:1px solid #A7F3D0; border-radius:8px; padding:12px;">
                  <strong style="color:#065F46; font-size:0.85rem; display:block; margin-bottom:6px;">\u{1F7E2} ZONA HIJAU (Publik):</strong>
                  <ul style="margin:0 0 0 16px; padding:0; font-size:0.82rem; color:#064E3B; line-height:1.6;">
                    ${D.hijau.map(O=>`<li>${O}</li>`).join("")}
                  </ul>
                </div>
                <!-- Zona Kuning -->
                <div style="background:#FFFBEB; border:1px solid #FDE68A; border-radius:8px; padding:12px;">
                  <strong style="color:#92400E; font-size:0.85rem; display:block; margin-bottom:6px;">\u{1F7E1} ZONA KUNING (Terbatas):</strong>
                  <ul style="margin:0 0 0 16px; padding:0; font-size:0.82rem; color:#78350F; line-height:1.6;">
                    ${D.kuning.map(O=>`<li>${O}</li>`).join("")}
                  </ul>
                </div>
                <!-- Zona Merah -->
                <div style="background:#FEF2F2; border:1px solid #FECACA; border-radius:8px; padding:12px;">
                  <strong style="color:#991B1B; font-size:0.85rem; display:block; margin-bottom:6px;">\u{1F534} ZONA MERAH (Restricted):</strong>
                  <ul style="margin:0 0 0 16px; padding:0; font-size:0.82rem; color:#7F1D1D; line-height:1.6;">
                    ${D.merah.map(O=>`<li>${O}</li>`).join("")}
                  </ul>
                </div>
              </div>
            </div>
          `).join("")}
        </div>
      `,f.querySelector("#view-cabang-btn").addEventListener("click",()=>{p="cabang",w(f)}),f.querySelector("#view-hq-btn").addEventListener("click",()=>{p="hq",w(f)})}function M(f){f.innerHTML=`
      <div style="display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:12px; margin-bottom:18px;">
        <div>
          <h3 style="margin:0 0 4px 0; font-size:1.1rem; font-weight:700; color:var(--text-1);">
            Daftar Klasifikasi Aset yang Dipantau Security (Lampiran 8.3)
          </h3>
          <p style="margin:0; font-size:0.85rem; color:var(--text-3);">
            Pengawasan dan perlindungan sarana prasarana fisik, elektronik, dan peralatan klinik
          </p>
        </div>

        <div style="display:flex; gap:10px; align-items:center; flex-wrap:wrap;">
          <input type="text" id="sec-asset-search" class="form-control" placeholder="Cari nama aset / tipe / PIC..." value="${h}" style="height:36px; border-radius:8px; font-size:0.875rem; width:240px;">
            ${n?`
          <button id="btn-add-sec-asset" class="btn btn-primary" style="font-weight:600; padding:6px 14px; border-radius:8px; display:inline-flex; align-items:center; gap:6px; cursor:pointer; font-size:0.85rem; background:#1E40AF; border-color:#1E40AF;">
            \u2795 Tambah Aset Baru
          </button>
            `:""}
        </div>
      </div>

      <div id="sec-asset-table-container"></div>
    `;let g=f.querySelector("#sec-asset-table-container"),D=f.querySelector("#sec-asset-search");function O(){let J=h.toLowerCase().trim(),$=E.filter(m=>!J||m.item.toLowerCase().includes(J)||m.type.toLowerCase().includes(J)||m.description.toLowerCase().includes(J)||m.pic&&m.pic.toLowerCase().includes(J));g.innerHTML=`
        <div class="card" style="background:#172033; border:1px solid var(--border); border-radius:12px; overflow:hidden; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
          <div style="padding:14px 18px; background:#111827; border-bottom:1px solid var(--border); display:flex; justify-content:space-between; align-items:center;">
            <strong style="color:var(--text-1); font-size:0.95rem;">Rekap Klasifikasi Aset Perusahaan</strong>
            <span class="badge badge-primary">${$.length} Aset</span>
          </div>
          <div style="overflow-x:auto;">
            <table class="table" style="width:100%; margin:0; border-collapse:collapse; font-size:0.875rem;">
              <thead>
                <tr style="background:#111827; border-bottom:1px solid var(--border); text-align:left;">
                  <th style="padding:10px 14px; width:60px;">No</th>
                  <th style="padding:10px 14px; width:140px;">Tipe Aset</th>
                  <th style="padding:10px 14px; width:220px;">Nama Aset / Kelompok</th>
                  <th style="padding:10px 14px;">Deskripsi / Contoh Item</th>
                  <th style="padding:10px 14px; width:180px;">PIC Pengawas</th>
                  <th style="padding:10px 14px; width:110px; text-align:center;">Aksi</th>
                </tr>
              </thead>
              <tbody>
                ${$.map((m,i)=>`
                  <tr style="border-bottom:1px solid var(--border); background:${i%2===0?"#172033":"#111827"}; color:var(--text-1);">
                    <td style="padding:10px 14px; font-weight:600; color:var(--text-3);">${m.no||i+1}</td>
                    <td style="padding:10px 14px;">
                      <span class="badge" style="${W(m.type)}">${P(m.type)}</span>
                    </td>
                    <td style="padding:10px 14px; font-weight:700; color:var(--text-1);">${P(m.item)}</td>
                    <td style="padding:10px 14px; color:var(--text-2);">${P(m.description)}</td>
                    <td style="padding:10px 14px; font-weight:600; color:#1E40AF;">${P(m.pic||"-")}</td>
                    <td style="padding:10px 14px; text-align:center;">
                    ${n?`
                      <button class="btn btn-sm btn-ghost btn-edit-asset" data-no="${m.no}" title="Edit Aset" style="padding:4px 6px; cursor:pointer;">\u270F\uFE0F</button>
                    `:""}
                    ${a?`
                      <button class="btn btn-sm btn-ghost btn-delete-asset" data-no="${m.no}" title="Hapus Aset" style="padding:4px 6px; color:#DC2626; cursor:pointer;">\u{1F5D1}\uFE0F</button>
                    `:""}
                    </td>
                  </tr>
                `).join("")}
              </tbody>
            </table>
          </div>
        </div>
      `,g.querySelectorAll(".btn-edit-asset").forEach(m=>{m.addEventListener("click",()=>{let i=parseInt(m.dataset.no,10),r=E.find(d=>d.no===i);r&&k(r)})}),g.querySelectorAll(".btn-delete-asset").forEach(m=>{m.addEventListener("click",()=>{let i=parseInt(m.dataset.no,10),r=E.find(d=>d.no===i);r&&me(`Apakah Anda yakin ingin menghapus aset: "${r.item}"?`,()=>{E=E.filter(d=>d.no!==i),E.forEach((d,b)=>{d.no=b+1}),Ma(E),ae("Aset berhasil dihapus"),C()},"Hapus Aset")})})}function W(J){return J==="Aset Tetap"?"background:#EFF6FF; color:#1E40AF; font-weight:700; padding:3px 8px; border-radius:6px;":J==="Aset Bergerak"?"background:#ECFDF5; color:#047857; font-weight:700; padding:3px 8px; border-radius:6px;":"background:#FEF3C7; color:#92400E; font-weight:700; padding:3px 8px; border-radius:6px;"}O(),D&&D.addEventListener("input",J=>{h=J.target.value,O()}),f.querySelector("#btn-add-sec-asset").addEventListener("click",()=>{k()})}function k(f=null){if(!n)return;let g=!!f,D=`
      <form id="asset-edit-form" style="display:flex; flex-direction:column; gap:14px; font-size:0.88rem;">
        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Tipe / Klasifikasi Aset *</label>
          <select id="form-asset-type" class="form-control" style="width:100%; border-radius:6px; height:36px;">
            <option value="Aset Tetap" ${f?.type==="Aset Tetap"?"selected":""}>Aset Tetap</option>
            <option value="Aset Bergerak" ${f?.type==="Aset Bergerak"?"selected":""}>Aset Bergerak</option>
            <option value="Low Value Asset" ${f?.type==="Low Value Asset"?"selected":""}>Low Value Asset</option>
          </select>
        </div>
        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Nama Aset / Kelompok *</label>
          <input type="text" id="form-asset-item" class="form-control" required value="${f?P(f.item):""}" placeholder="Contoh: Genset & Elevator" style="width:100%; border-radius:6px; height:36px;">
        </div>
        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Deskripsi / Contoh Item *</label>
          <textarea id="form-asset-desc" class="form-control" rows="3" required placeholder="Contoh: Generator daya darurat dan panel listrik utama" style="width:100%; border-radius:6px; font-size:0.85rem;">${f?P(f.description):""}</textarea>
        </div>
        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">PIC Pengawas / Penanggung Jawab *</label>
          <input type="text" id="form-asset-pic" class="form-control" required value="${f?P(f.pic):""}" placeholder="Contoh: Teknisi & Security" style="width:100%; border-radius:6px; height:36px;">
        </div>
      </form>
    `;pe({title:g?`\u270F\uFE0F Edit Aset: ${f.item}`:"\u2795 Tambah Klasifikasi Aset Baru",content:D,confirmText:"Simpan",onConfirm:(O,W)=>{let J=O.querySelector("#form-asset-type").value,$=O.querySelector("#form-asset-item").value.trim(),m=O.querySelector("#form-asset-desc").value.trim(),i=O.querySelector("#form-asset-pic").value.trim();if(!$||!m||!i){z("Harap lengkapi semua kolom");return}if(g)f.type=J,f.item=$,f.description=m,f.pic=i,Ma(E),ae(`Aset "${$}" berhasil diperbarui`);else{let r=E.length>0?Math.max(...E.map(d=>d.no))+1:1;E.push({no:r,type:J,item:$,description:m,pic:i}),Ma(E),ae(`Aset baru "${$}" berhasil ditambahkan`)}W(),C()}})}function x(f){f.innerHTML=`
      <div style="display:grid; grid-template-columns:repeat(auto-fit, minmax(340px, 1fr)); gap:20px;">
        <!-- Card 1: Shift Kerja -->
        <div class="card" style="background:#172033; border:1px solid var(--border); border-radius:12px; padding:20px; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
          <div style="display:flex; align-items:center; gap:10px; margin-bottom:14px; border-bottom:1px solid #F1F5F9; padding-bottom:10px;">
            <span style="font-size:1.5rem; background:#EFF6FF; padding:8px; border-radius:8px;">\u23F0</span>
            <div>
              <h3 style="margin:0; font-size:1.1rem; font-weight:700; color:var(--text-1);">Sistem Jam Kerja & Shift (6.2)</h3>
              <span style="font-size:0.8rem; color:var(--text-3);">Ketentuan 12 Jam Kerja dengan 2 Hari Libur per Minggu</span>
            </div>
          </div>
          <div style="font-size:0.88rem; color:var(--text-2); line-height:1.7;">
            <div style="background:#111827; border-radius:8px; padding:10px 14px; margin-bottom:10px; border-left:3px solid #2563EB;">
              <strong style="color:var(--text-1); display:block; margin-bottom:2px;">\u2600\uFE0F Shift 1 (Pagi - Siang):</strong>
              <span>Pukul <strong>07.00 s/d 19.00 WIB</strong></span>
            </div>
            <div style="background:#111827; border-radius:8px; padding:10px 14px; margin-bottom:12px; border-left:3px solid #1E293B;">
              <strong style="color:var(--text-1); display:block; margin-bottom:2px;">\u{1F319} Shift 2 (Malam):</strong>
              <span>Pukul <strong>19.00 s/d 07.00 WIB</strong></span>
            </div>
            <p style="margin:0; font-size:0.82rem; color:var(--text-3);">
              <strong>Ketentuan Off:</strong> Berlaku 2 hari off dalam setiap pekannya sesuai jadwal roster yang ditetapkan oleh Facility Care Supervisor / General Support.
            </p>
          </div>
        </div>

        <!-- Card 2: Jadwal Safety Patrol -->
        <div class="card" style="background:#172033; border:1px solid var(--border); border-radius:12px; padding:20px; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
          <div style="display:flex; align-items:center; gap:10px; margin-bottom:14px; border-bottom:1px solid #F1F5F9; padding-bottom:10px;">
            <span style="font-size:1.5rem; background:#FEF3C7; padding:8px; border-radius:8px;">\u{1F6B6}\u200D\u2642\uFE0F</span>
            <div>
              <h3 style="margin:0; font-size:1.1rem; font-weight:700; color:var(--text-1);">Jadwal Safety Patrol Harian (6.2 B)</h3>
              <span style="font-size:0.8rem; color:var(--text-3);">Keliling Pengecekan Area & Form Checklist Digital</span>
            </div>
          </div>
          <div style="display:flex; flex-direction:column; gap:8px; font-size:0.86rem;">
            <div style="background:#111827; padding:8px 12px; border-radius:6px; display:flex; justify-content:space-between; align-items:center;">
              <strong>Patroli 1 (Pagi):</strong>
              <span class="badge" style="background:#E0F2FE; color:#0369A1; font-weight:700;">07.30 - 08.30 WIB</span>
            </div>
            <div style="background:#111827; padding:8px 12px; border-radius:6px; display:flex; justify-content:space-between; align-items:center;">
              <strong>Patroli 2 (Malam):</strong>
              <span class="badge" style="background:#FEF3C7; color:#92400E; font-weight:700;">19.30 - 20.30 WIB</span>
            </div>
            <div style="background:#111827; padding:8px 12px; border-radius:6px; display:flex; justify-content:space-between; align-items:center;">
              <strong>Patroli 3 (Dini Hari):</strong>
              <span class="badge" style="background:#F3E8FF; color:#7E22CE; font-weight:700;">04.00 - 05.00 WIB</span>
            </div>
            <div style="font-size:0.8rem; color:var(--text-3); margin-top:4px;">
              * <em>Patroli 3 khusus berlaku untuk area Head Quarter atau lokasi yang memiliki lebih dari 1 personil security.</em>
            </div>
          </div>
        </div>

        <!-- Card 3: Penanganan Pelanggaran Disiplin -->
        <div class="card" style="background:#172033; border:1px solid var(--border); border-radius:12px; padding:20px; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
          <div style="display:flex; align-items:center; gap:10px; margin-bottom:14px; border-bottom:1px solid #F1F5F9; padding-bottom:10px;">
            <span style="font-size:1.5rem; background:#ECFDF5; padding:8px; border-radius:8px;">\u26A0\uFE0F</span>
            <div>
              <h3 style="margin:0; font-size:1.1rem; font-weight:700; color:var(--text-1);">Penanganan Pelanggaran Disiplin (6.4)</h3>
              <span style="font-size:0.8rem; color:var(--text-3);">Pelanggaran SK Tata Tertib Perusahaan</span>
            </div>
          </div>
          <ol style="margin:0 0 0 18px; padding:0; font-size:0.85rem; color:var(--text-2); line-height:1.7;">
            <li>Petugas security membuat <strong>Berita Acara Umum (BA)</strong> kejadian secara lengkap.</li>
            <li>Melaporkan kronologis tertulis kepada atasan (General Support Manager / FC Supervisor).</li>
            <li>Melaporkan dan meneruskan berkas kepada atasan langsung dari karyawan yang bersangkutan untuk penanganan HR/HC.</li>
          </ol>
        </div>

        <!-- Card 4: Penanganan Pelanggaran Kriminal -->
        <div class="card" style="background:#172033; border:1px solid var(--border); border-radius:12px; padding:20px; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
          <div style="display:flex; align-items:center; gap:10px; margin-bottom:14px; border-bottom:1px solid #F1F5F9; padding-bottom:10px;">
            <span style="font-size:1.5rem; background:#FEF2F2; padding:8px; border-radius:8px;">\u{1F6A8}</span>
            <div>
              <h3 style="margin:0; font-size:1.1rem; font-weight:700; color:#DC2626;">Penanganan Pelanggaran Kriminal (6.4)</h3>
              <span style="font-size:0.8rem; color:var(--text-3);">Tindak Pidana, Pencurian, & Pelanggaran Hukum</span>
            </div>
          </div>
          <ol style="margin:0 0 0 18px; padding:0; font-size:0.84rem; color:var(--text-2); line-height:1.7;">
            <li>Buat <strong>Berita Acara Umum</strong> secara resmi.</li>
            <li>Periksa dan amankan rekaman bukti visual melalui <strong>CCTV</strong>.</li>
            <li>Laporkan ke atasan dari karyawan terkait dengan rincian kronologis (bila pelaku karyawan).</li>
            <li>Minta dan catat detail identitas <strong>KTP / ID Card</strong> dari pihak yang bersangkutan.</li>
            <li>Amankan dan netralisir <strong>Tempat Kejadian Perkara (TKP)</strong> dari gangguan orang lain.</li>
            <li>Laporkan ke <strong>Kepolisian terdekat</strong> untuk proses hukum tindak lanjut.</li>
          </ol>
        </div>
      </div>
    `}function _(f){f.innerHTML=`
      <div style="display:grid; grid-template-columns:repeat(auto-fit, minmax(320px, 1fr)); gap:18px;">
        <!-- Card: Google Form Safety Patrol -->
        <div class="card" style="background:#172033; border:1px solid var(--border); border-radius:12px; padding:20px; box-shadow:0 1px 3px rgba(0,0,0,0.03); display:flex; flex-direction:column; justify-content:space-between;">
          <div>
            <div style="display:flex; align-items:center; gap:10px; margin-bottom:12px;">
              <span style="font-size:2rem;">\u{1F4D1}</span>
              <div>
                <h3 style="margin:0; font-size:1.1rem; font-weight:700; color:var(--text-1);">Form Safety Patrol Audy</h3>
                <span class="badge" style="background:#E0F2FE; color:#0369A1; font-size:0.75rem; font-weight:700;">Google Forms Resmi</span>
              </div>
            </div>
            <p style="font-size:0.85rem; color:var(--text-2); line-height:1.6; margin-bottom:16px;">
              Formulir digital inspeksi keamanan yang wajib diisi oleh tim security saat melakukan Safety Patrol rutin di klinik dan kantor pusat Audy Dental.
            </p>
          </div>
          <a href="https://forms.gle/1EAvkqDJVdctXEGn7" target="_blank" rel="noopener" class="btn btn-primary" style="display:inline-flex; align-items:center; justify-content:center; gap:6px; font-weight:600; padding:10px; border-radius:8px; background:#1E40AF; border-color:#1E40AF;">
            \u{1F4CB} Buka Form Safety Patrol \u2794
          </a>
        </div>

        <!-- Card: Dasar Hukum & Ruang Lingkup Perusahaan -->
        <div class="card" style="background:#172033; border:1px solid var(--border); border-radius:12px; padding:20px; box-shadow:0 1px 3px rgba(0,0,0,0.03); display:flex; flex-direction:column; justify-content:space-between;">
          <div>
            <div style="display:flex; align-items:center; gap:10px; margin-bottom:12px;">
              <span style="font-size:2rem;">\u2696\uFE0F</span>
              <div>
                <h3 style="margin:0; font-size:1.1rem; font-weight:700; color:var(--text-1);">Dasar Hukum & Ruang Lingkup</h3>
                <span class="badge" style="background:#1B2940; color:var(--text-2); font-size:0.75rem; font-weight:700;">Legalitas</span>
              </div>
            </div>
            <p style="font-size:0.85rem; color:var(--text-2); line-height:1.6; margin-bottom:12px;">
              SOP ini diterapkan sebagai tata cara Pengelolaan Keamanan Aset Perusahaan di seluruh lokasi Audy Dental Klinik di bawah naungan:
            </p>
            <ul style="margin:0 0 0 16px; padding:0; font-size:0.84rem; color:var(--text-2); line-height:1.6;">
              <li><strong>PT. Audy Mandiri Indonesia</strong></li>
              <li><strong>PT. Kastara Ananta</strong></li>
              <li><strong>PT. Nara Adidaya</strong></li>
            </ul>
          </div>
        </div>
      </div>

      <!-- Validasi & Pengesahan SOP -->
      <div class="card" style="margin-top:20px; padding:20px; background:#172033; border:1px solid var(--border); border-radius:12px;">
        <h3 style="margin:0 0 14px 0; font-size:1rem; font-weight:700; color:var(--text-1); display:flex; align-items:center; gap:8px;">
          <span>\u{1F4DC}</span> Validasi & Pengesahan Dokumen SOP Security
        </h3>
        <div style="display:grid; grid-template-columns:repeat(auto-fit, minmax(260px, 1fr)); gap:16px;">
          <div style="border-left:3px solid #2563EB; padding-left:12px;">
            <div style="font-size:0.75rem; color:var(--text-3); font-weight:700; text-transform:uppercase;">Disiapkan Oleh:</div>
            <div style="font-size:0.95rem; font-weight:700; color:var(--text-1); margin-top:2px;">Berlin Ariansyah</div>
            <div style="font-size:0.8rem; color:var(--text-2);">Facility Care Supervisor</div>
          </div>
          <div style="border-left:3px solid #059669; padding-left:12px;">
            <div style="font-size:0.75rem; color:var(--text-3); font-weight:700; text-transform:uppercase;">Diketahui Oleh:</div>
            <div style="font-size:0.95rem; font-weight:700; color:var(--text-1); margin-top:2px;">Catur Meiyogo</div>
            <div style="font-size:0.8rem; color:var(--text-2);">General Support Manager</div>
          </div>
          <div style="border-left:3px solid #7C3AED; padding-left:12px;">
            <div style="font-size:0.75rem; color:var(--text-3); font-weight:700; text-transform:uppercase;">Disetujui Oleh:</div>
            <div style="font-size:0.95rem; font-weight:700; color:var(--text-1); margin-top:2px;">Drg. Yulita</div>
            <div style="font-size:0.8rem; color:var(--text-2);">Chief Executive Officer</div>
          </div>
        </div>
      </div>
    `}async function B(){try{let{downloadExcel:f}=await Promise.resolve().then(()=>(we(),Oe)),g=c.map(O=>({"Kode IK":`IK-SEC-${String(O.id).padStart(2,"0")}`,"Judul Prosedur":O.title,Kategori:O.category,"Sasaran Target":O.target,"Alat & Perlengkapan":O.tools.join(", "),"Langkah Pelaksanaan":O.steps.map((W,J)=>`${J+1}. ${W}`).join(" | "),"Catatan Khusus":O.notes||"-"})),D=new Date().toISOString().slice(0,10);f(g,`SOP_Security_Instruksi_Kerja_${D}`),ae("Data Excel SOP Security berhasil diunduh")}catch(f){console.error(f),z("Gagal mendownload data Excel: "+f.message)}}function P(f){return f?String(f).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;").replace(/'/g,"&#39;"):""}C()}Fe();function Qn(e=""){let t=(e||"").toLowerCase();return t.includes("berita acara")?{label:e,bg:"#F3E8FF",color:"#7E22CE",border:"#D8B4FE",icon:"\u2696\uFE0F"}:t.includes("apar")||t.includes("k3")||t.includes("safety")?{label:e,bg:"#FEE2E2",color:"#B91C1C",border:"#FECACA",icon:"\u{1F9EF}"}:t.includes("deep cleaning")?{label:e,bg:"#DCFCE7",color:"#15803D",border:"#BBF7D0",icon:"\u2728"}:t.includes("drain")?{label:e,bg:"#FEF3C7",color:"#B45309",border:"#FDE68A",icon:"\u{1F6BF}"}:t.includes("general cleaning")||t.includes("kebersihan")?{label:e,bg:"#E0F2FE",color:"#0369A1",border:"#BAE6FD",icon:"\u{1F9F9}"}:t.includes("form")?{label:e,bg:"#E0E7FF",color:"#4338CA",border:"#C7D2FE",icon:"\u{1F4DD}"}:{label:e||"Dokumen",bg:"#F1F5F9",color:"#475569",border:"#CBD5E1",icon:"\u{1F4C1}"}}function Xn(e=""){if(!e)return{type:"Tidak Ada Tautan",icon:"\u{1F4CE}",badgeClass:"bg-slate-100 text-slate-500"};let t=e.toLowerCase();return t.includes("forms.gle")||t.includes("google.com/forms")?{type:"Google Form",icon:"\u{1F7E3}",badgeClass:"badge-form",color:"#673AB7"}:t.includes("docs.google.com/spreadsheets")||t.includes("sheet")?{type:"Google Sheets",icon:"\u{1F7E2}",badgeClass:"badge-sheet",color:"#0F9D58"}:t.includes("docs.google.com/document")||t.includes("doc")?{type:"Google Docs",icon:"\u{1F535}",badgeClass:"badge-doc",color:"#4285F4"}:t.includes("drive.google.com")?{type:"Google Drive",icon:"\u{1F7E1}",badgeClass:"badge-drive",color:"#F4B400"}:t.endsWith(".pdf")?{type:"PDF Document",icon:"\u{1F534}",badgeClass:"badge-pdf",color:"#EA4335"}:{type:"Web Link",icon:"\u{1F310}",badgeClass:"badge-link",color:"#2563EB"}}async function Zn(e){let t=de(),n=ie(t,"checklist","write"),a=ie(t,"checklist","delete"),o=[],s="all",l="",h=1,p=20,c=0,E=1,C={categoryCounts:{all:0},linkedCount:0},T=null,u=!1,F="fcms-master-doc-styles";if(!document.getElementById(F)){let $=document.createElement("style");$.id=F,$.textContent=`
      .doc-page-container {
        padding: 24px;
        max-width: 1400px;
        margin: 0 auto;
        font-family: inherit;
      }
      .doc-hero-card {
        background: linear-gradient(135deg, #1E293B 0%, #0F172A 100%);
        border-radius: 18px;
        padding: 28px 32px;
        color: #ffffff;
        box-shadow: 0 10px 25px -5px rgba(15, 23, 42, 0.2);
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
        background: radial-gradient(circle, rgba(59, 130, 246, 0.18) 0%, transparent 70%);
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
        background: #172033;
        border: 1px solid #29364B;
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
        background: #1B2940;
        transition: all 0.2s ease;
      }
      .doc-search-input:focus {
        outline: none;
        border-color: #2563EB;
        background: #172033;
        box-shadow: 0 0 0 3px rgba(37, 99, 235, 0.12);
      }
      .doc-filter-pills {
        display: flex;
        flex-wrap: wrap;
        gap: 8px;
        align-items: center;
      }
      .doc-pill {
        border: 1px solid #29364B;
        background: #1B2940;
        color: #475569;
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
        background: #EEF2F6;
        border-color: #CBD5E1;
      }
      .doc-pill.active {
        background: #2563EB;
        border-color: #2563EB;
        color: #FFFFFF;
        box-shadow: 0 2px 6px rgba(37, 99, 235, 0.25);
      }
      .doc-pill-count {
        background: rgba(0, 0, 0, 0.08);
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
        background: #172033;
        border: 1px solid #29364B;
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
        font-size: 0.74rem;
        font-weight: 700;
        letter-spacing: 0.02em;
        text-transform: uppercase;
      }
      .doc-platform-badge {
        display: inline-flex;
        align-items: center;
        gap: 4px;
        font-size: 0.74rem;
        font-weight: 600;
        color: #64748B;
        background: #111827;
        padding: 3px 8px;
        border-radius: 6px;
      }
      .doc-card-title {
        font-size: 1.08rem;
        font-weight: 700;
        color: #0F172A;
        margin: 0 0 8px 0;
        line-height: 1.4;
      }
      .doc-card-desc {
        color: #64748B;
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
        border-top: 1px solid #F1F5F9;
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
        background: #1B2940;
        border: 1px solid #29364B;
        color: #64748B;
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
        background: #EDF2F7;
        color: #1E293B;
        border-color: #CBD5E1;
      }
      .doc-icon-btn.delete:hover {
        background: #FEE2E2;
        color: #DC2626;
        border-color: #FECACA;
      }
      /* Table View */
      .doc-table-card {
        background: #172033;
        border: 1px solid #29364B;
        border-radius: 16px;
        box-shadow: 0 2px 8px rgba(0, 0, 0, 0.04);
        overflow: hidden;
      }
      .doc-table {
        width: 100%;
        border-collapse: collapse;
        text-align: left;
      }
      .doc-table th {
        background: #1B2940;
        padding: 14px 18px;
        font-size: 0.78rem;
        font-weight: 700;
        color: #475569;
        text-transform: uppercase;
        letter-spacing: 0.04em;
        border-bottom: 1.5px solid #E2E8F0;
      }
      .doc-table td {
        padding: 16px 18px;
        font-size: 0.88rem;
        color: #1E293B;
        border-bottom: 1px solid #F1F5F9;
        vertical-align: middle;
      }
      .doc-table tr:hover td {
        background: #1B2940;
      }
      .doc-empty-state {
        text-align: center;
        padding: 60px 20px;
        background: #172033;
        border-radius: 16px;
        border: 1px dashed #29364B;
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
          font-size: 0.68rem !important;
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
    `,document.head.appendChild($)}function R(){let $={};return l&&l.trim()&&($.search=l.trim()),s&&s!=="all"&&($.category=s),$}function K(){let $=k(),m=Object.keys($).filter(i=>i!=="all");return`
      <button class="doc-pill ${s==="all"?"active":""}" data-cat="all">
        <span>Semua Dokumen</span>
        <span class="doc-pill-count">${$.all||0}</span>
      </button>
      ${m.map(i=>{let r=Qn(i),d=s===i,b=String(i).replace(/"/g,"&quot;");return`
          <button class="doc-pill ${d?"active":""}" data-cat="${b}">
            <span>${r.icon}</span>
            <span>${i}</span>
            <span class="doc-pill-count">${$[i]||0}</span>
          </button>
        `}).join("")}
    `}function I(){e.querySelectorAll(".doc-pill").forEach($=>{$.onclick=()=>{s=$.getAttribute("data-cat")||"all",e.querySelectorAll(".doc-pill").forEach(m=>{m.classList.toggle("active",m.getAttribute("data-cat")===s)}),h=1,w({refreshStats:!1})}})}function N(){let $=k(),m=Object.keys($).filter(y=>y!=="all"),i=Number($.all??0),r=Number(C.linkedCount??0),d=(y,q)=>{let v=e.querySelector(y);v&&(v.textContent=String(q??0))};d("#doc-stat-total",i),d("#doc-stat-categories",m.length),d("#doc-stat-linked",r),d("#doc-stat-linked-total",i);let b=e.querySelector("#doc-filter-pills");b&&(b.innerHTML=K(),I())}async function w({refreshStats:$=!0}={}){try{let m=R(),i=new URLSearchParams({page:String(h),limit:String(p),...m});$&&i.set("include_stats","1");let r=await X("/api/checklist?"+i.toString());if(r.ok&&r.data&&Array.isArray(r.data.data)){let d=r.data.pagination||{};if(c=Number(d.total??r.data.data.length),E=Math.max(1,Number(d.pages||d.totalPages||Math.ceil(c/p)||1)),r.data.data.length===0&&h>E)return h=E,w({refreshStats:$});o=r.data.data,r.data.stats&&(C={categoryCounts:r.data.stats.categoryCounts||{all:c},linkedCount:Number(r.data.stats.linkedCount??0)})}else o=[],c=0,E=1,r.ok||z(r.data?.error||"Gagal memuat dokumen dari server")}catch(m){console.warn("Error loading documents:",m),o=[],c=0,E=1,z("Gagal memuat dokumen: "+m.message)}u?($&&N(),O()):(u=!0,W())}function M(){return o}function k(){return C.categoryCounts||{all:0}}function x($=null){if(!n)return;let m=!!$,i=Array.from(new Set([...Object.keys(C.categoryCounts||{}).filter(d=>d!=="all"&&d!=="Lainnya"),...o.map(d=>d.category).filter(Boolean)])),r=document.createElement("div");r.innerHTML=`
      <div style="display:flex; flex-direction:column; gap:16px;">
        <div>
          <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:6px;">
            Nama Dokumen / Checklist <span style="color:#DC2626;">*</span>
          </label>
          <input type="text" id="doc-modal-name" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
            placeholder="Contoh: Checklist Pembersihan Wastafel & Toilet" value="${$?.name?String($.name).replace(/"/g,"&quot;"):""}" required />
        </div>

        <div>
          <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:6px;">
            Kategori Dokumen <span style="color:#DC2626;">*</span>
          </label>
          <input type="text" id="doc-modal-category" list="doc-categories-list" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
            placeholder="Pilih atau ketik kategori baru..." value="${$?.category?String($.category).replace(/"/g,"&quot;"):""}" required />
          <datalist id="doc-categories-list">
            ${i.map(d=>`<option value="${d}">${d}</option>`).join("")}
            <option value="Berita Acara">Berita Acara</option>
            <option value="Apar & K3">Apar & K3</option>
            <option value="Kegiatan Deep Cleaning">Kegiatan Deep Cleaning</option>
            <option value="Kegiatan General Cleaning">Kegiatan General Cleaning</option>
            <option value="Drain Cleaning">Drain Cleaning</option>
            <option value="Form Operasional">Form Operasional</option>
          </datalist>
        </div>

        <div>
          <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:6px;">
            Tautan / Link Dokumen (Google Drive, Docs, Sheets, atau Forms)
          </label>
          <input type="url" id="doc-modal-link" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
            placeholder="https://docs.google.com/..." value="${$?.document_link?String($.document_link).replace(/"/g,"&quot;"):""}" />
          <span style="font-size:0.75rem; color:#64748B; margin-top:4px; display:block;">
            Bisa berupa tautan Google Forms untuk inspeksi harian, atau link Google Sheets / Docs sebagai acuan.
          </span>
        </div>

        <div>
          <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:6px;">
            Deskripsi / Petunjuk Penggunaan
          </label>
          <textarea id="doc-modal-desc" class="form-control" rows="3" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
            placeholder="Tuliskan keterangan singkat mengenai tujuan dan instruksi pengisian dokumen ini...">${$?.description||""}</textarea>
        </div>
      </div>
    `,pe({title:m?"\u270F\uFE0F Edit Master Dokumen":"\u2795 Tambah Master Dokumen Baru",content:r,size:"md",confirmText:m?"Simpan Perubahan":"Tambahkan Dokumen",confirmClass:"btn-primary",onConfirm:async(d,b)=>{let y=r.querySelector("#doc-modal-name"),q=r.querySelector("#doc-modal-category"),v=r.querySelector("#doc-modal-link"),S=r.querySelector("#doc-modal-desc"),A=(y.value||"").trim(),U=(q.value||"").trim(),H=(v.value||"").trim(),ee=(S.value||"").trim();if(!A){z("Nama Dokumen wajib diisi"),y.focus();return}if(!U){z("Kategori Dokumen wajib diisi"),q.focus();return}let Q={name:A,category:U,document_link:H,description:ee};try{if(m){let ne=await X(`/api/checklist/${$.id}`,{method:"PUT",body:JSON.stringify(Q)});if(!ne.ok){z(ne.data?.error||"Gagal memperbarui dokumen");return}ae("Dokumen berhasil diperbarui")}else{let ne=await X("/api/checklist",{method:"POST",body:JSON.stringify(Q)});if(!ne.ok){z(ne.data?.error||"Gagal menambahkan dokumen");return}h=1,ae("Dokumen baru berhasil ditambahkan")}b(),await w({refreshStats:!0})}catch(ne){console.error("Error saving document:",ne),z("Gagal menyimpan dokumen: "+ne.message)}}})}function _($){if(!a){z("Anda tidak memiliki izin untuk menghapus dokumen.");return}me(`Apakah Anda yakin ingin menghapus dokumen <strong>"${$.name}"</strong>?`,async()=>{try{let m=await X(`/api/checklist/${$.id}`,{method:"DELETE"});if(!m.ok){z(m.data?.error||"Gagal menghapus dokumen");return}ae(`Dokumen "${$.name}" berhasil dihapus`),await w({refreshStats:!0})}catch(m){z("Gagal menghapus dokumen: "+m.message)}},"Hapus Dokumen")}function B($,m){if(!$){ct("Dokumen belum memiliki tautan digital.");return}navigator.clipboard.writeText($).then(()=>{ae(`Tautan "${m}" disalin ke clipboard!`)}).catch(()=>{ct(`Tautan: ${$}`)})}async function P(){if(!ie(t,"export")){z("Anda tidak memiliki izin untuk mengekspor data.");return}try{let{downloadExcel:$}=await Promise.resolve().then(()=>(we(),Oe)),m=await ue("/api/checklist/export",R()),i=m.map((r,d)=>({No:d+1,"Nama Dokumen":r.name||"",Kategori:r.category||"","Platform / Format":Xn(r.document_link).type,"Link Dokumen":r.document_link||"","Deskripsi / Keterangan":r.description||""}));$(i,`Master_Dokumen_FCMS_${new Date().toISOString().slice(0,10)}`),ae(`Berhasil mengunduh ${m.length} data dokumen ke Excel.`)}catch($){z("Gagal mengekspor data: "+$.message)}}function f($){return`
    ${$.length===0?`
            <div class="doc-empty-state">
              <div style="font-size:3rem; margin-bottom:12px;">\u{1F50D}</div>
              <h3 style="font-size:1.2rem; font-weight:700; color:#1E293B; margin:0 0 6px 0;">Tidak Ada Dokumen Ditemukan</h3>
              <p style="color:#64748B; font-size:0.9rem; margin:0 0 20px 0;">Coba sesuaikan kata kunci pencarian atau pilih kategori lain.</p>
              <button id="doc-btn-reset-filter" class="btn btn-ghost" style="border:1px solid #CBD5E1; border-radius:8px; padding:8px 16px;">
                Tampilkan Semua Dokumen
              </button>
            </div>
          `:`
            <!-- Grid Card View -->
            <div class="doc-cards-grid">
              ${$.map(m=>{let i=Qn(m.category),r=Xn(m.document_link);return`
                  <div class="doc-card" data-id="${m.id}">
                    <div>
                      <div class="doc-card-top">
                        <span class="doc-category-badge" style="background:${i.bg}; color:${i.color}; border:1px solid ${i.border};">
                          <span>${i.icon}</span>
                          <span>${i.label}</span>
                        </span>
                        <span class="doc-platform-badge">
                          <span>${r.icon}</span>
                          <span>${r.type}</span>
                        </span>
                      </div>
                      
                      <h3 class="doc-card-title">${m.name}</h3>
                      <p class="doc-card-desc">${m.description||"Formulir acuan standar operasional dan administrasi fasilitas."}</p>
                    </div>

                    <div class="doc-card-footer">
                      ${m.document_link?`
                        <a href="${m.document_link}" target="_blank" rel="noopener noreferrer" class="doc-btn-open">
                          <span>Buka Dokumen</span>
                          <svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M18 13v6a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V8a2 2 0 0 1 2-2h6"/><polyline points="15 3 21 3 21 9"/><line x1="10" y1="14" x2="21" y2="3"/></svg>
                        </a>
                      `:`
                        <button class="doc-btn-open disabled" title="Belum ada tautan yang dikaitkan">
                          <span>Tidak Ada Link</span>
                        </button>
                      `}

                      <div class="doc-card-actions">
                        ${m.document_link?`
                          <button class="doc-icon-btn copy-link-btn" title="Salin Tautan Dokumen" data-link="${m.document_link}" data-title="${m.name}">
                            <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><rect x="9" y="9" width="13" height="13" rx="2" ry="2"/><path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1"/></svg>
                          </button>
                        `:""}
                        ${n?`
                        <button class="doc-icon-btn edit-doc-btn" title="Edit Dokumen" data-id="${m.id}">
                          <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>
                        </button>
                        `:""}
                        ${a?`
                        <button class="doc-icon-btn delete delete-doc-btn" title="Hapus Dokumen" data-id="${m.id}">
                          <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><polyline points="3 6 5 6 21 6"/><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"/></svg>
                        </button>
                        `:""}
                      </div>
                    </div>
                  </div>
                `}).join("")}
            </div>
          `}
    `}function g(){let $=e.querySelector("#doc-btn-reset-filter");$&&($.onclick=()=>{l="",s="all",h=1,clearTimeout(T);let m=e.querySelector("#doc-search-input");m&&(m.value=""),e.querySelectorAll(".doc-pill").forEach(i=>{i.classList.toggle("active",i.getAttribute("data-cat")==="all")}),w({refreshStats:!1})}),e.querySelectorAll(".copy-link-btn").forEach(m=>{m.onclick=i=>{i.stopPropagation();let r=m.getAttribute("data-link"),d=m.getAttribute("data-title");B(r,d)}}),e.querySelectorAll(".edit-doc-btn").forEach(m=>{m.onclick=i=>{i.stopPropagation();let r=Number(m.getAttribute("data-id")),d=o.find(b=>b.id===r);d&&x(d)}}),e.querySelectorAll(".delete-doc-btn").forEach(m=>{m.onclick=i=>{i.stopPropagation();let r=Number(m.getAttribute("data-id")),d=o.find(b=>b.id===r);d&&_(d)}})}function D(){let $=e.querySelector("#doc-pagination-container");if(!$)return;$.innerHTML="";let m=_e({page:h,pages:E,total:c,limit:p,onPage:i=>{h=i,w({refreshStats:!1})}});m&&$.appendChild(m)}function O(){let $=e.querySelector("#doc-list-view");if(!$)return;let m=M();$.innerHTML=f(m),D(),g()}function W(){let $=M(),m=k(),i=Object.keys(m).filter(b=>b!=="all"),r=Number(m.all??0),d=Number(C.linkedCount??0);e.innerHTML=`
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
              ${ie(t,"export")?`<button id="doc-btn-excel" class="btn btn-ghost" style="background:rgba(255,255,255,0.08); border:1px solid rgba(255,255,255,0.15); color:white; height:38px; padding:0 14px; border-radius:10px; font-weight:600; display:inline-flex; align-items:center; gap:6px; white-space:nowrap;">
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
                <div class="doc-stat-value" id="doc-stat-total">${r}</div>
                <div class="doc-stat-label">Total Dokumen</div>
              </div>
            </div>
            <div class="doc-stat-box">
              <div class="doc-stat-icon">\u{1F5C2}\uFE0F</div>
              <div>
                <div class="doc-stat-value" id="doc-stat-categories">${i.length}</div>
                <div class="doc-stat-label">Kategori Dokumen</div>
              </div>
            </div>
            <div class="doc-stat-box">
              <div class="doc-stat-icon" style="background:rgba(16,185,129,0.18); color:#34D399;">\u{1F310}</div>
              <div>
                <div class="doc-stat-value"><span id="doc-stat-linked">${d}</span> <span style="font-size:0.8rem; font-weight:400; color:#94A3B8;">/ <span id="doc-stat-linked-total">${r}</span></span></div>
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
              <input type="text" id="doc-search-input" class="doc-search-input" placeholder="Cari nama dokumen, kategori, atau deskripsi..." value="${l?l.replace(/"/g,"&quot;"):""}" />
            </div>
          </div>

          <!-- Category Pills -->
          <div class="doc-filter-pills" id="doc-filter-pills">
            ${K()}
          </div>
        </div>

        <!-- Documents View Container -->
        <div id="doc-list-view">
          ${f($)}
        </div>

        <!-- Pagination Container -->
        <div id="doc-pagination-container"></div>

      </div>
    `,D(),J()}function J(){let $=e.querySelector("#doc-btn-add");$&&($.onclick=()=>x());let m=e.querySelector("#doc-btn-excel");m&&(m.onclick=()=>P());let i=e.querySelector("#doc-btn-refresh");i&&(i.onclick=()=>w({refreshStats:!0}));let r=e.querySelector("#doc-search-input");r&&(r.oninput=d=>{l=d.target.value,clearTimeout(T),T=setTimeout(()=>{h=1,w({refreshStats:!1})},350)}),I(),g()}await w({refreshStats:!0})}Fe();we();var ht={"Standar Umum":{icon:"\u{1F4CC}",bg:"#F3E8FF",color:"#6B21A8",border:"#D8B4FE",rooms:["Lantai & Sudut Ruangan","Dinding, Partisi & Plint","Kaca, Cermin & Partisi Kaca","Sanitair & Area Basah Toilet","Plafon, Lampu & Kisi Ventilasi","Mebel, Meja & Kursi Kerja/Pasien","Peralatan Khusus & Medis Non-Steril","Wadah Sampah & Utilitas"]},"Area Depan & Pelayanan":{icon:"\u{1F6CB}\uFE0F",bg:"#E0F2FE",color:"#0369A1",border:"#BAE6FD",rooms:["Teras","Lobby 1","Receptionist","Snack Corner"]},"Area Klinis & Medis":{icon:"\u{1FA7A}",bg:"#DBEAFE",color:"#1E40AF",border:"#93C5FD",rooms:["Ruang Tindakan","Ronsen","Ruang Dokter","Ruang Steril"]},"Toilet & Sanitasi":{icon:"\u{1F6BF}",bg:"#CCFBF1",color:"#0F766E",border:"#99F6E4",rooms:["Toilet 1","Janitor","Ruang Limbah"]},Sirkulasi:{icon:"\u{1F6B6}",bg:"#F1F5F9",color:"#334155",border:"#CBD5E1",rooms:["Area Tangga","Lift","Area Tangga Exit"]},"Area Anak & Fasilitas":{icon:"\u{1F9F8}",bg:"#FEF3C7",color:"#92400E",border:"#FDE68A",rooms:["Lobby 2","Playground","Musholla","Tempat Wudhu","Ruang Team"]},"Gudang & Teknis":{icon:"\u2699\uFE0F",bg:"#FFEDD5",color:"#C2410C",border:"#FED7AA",rooms:["Gudang Farmasi & Bahan","RoofTop","Ruang Kompressor","Area Genset"]},Eksterior:{icon:"\u{1F333}",bg:"#DCFCE7",color:"#15803D",border:"#86EFAC",rooms:["Taman","Pos Security","Toilet Security","Halaman Parkir"]}};function za(e){for(let[t,n]of Object.entries(ht))if(n.rooms.includes(e))return{zoneName:t,...n};return{zoneName:"Lainnya",icon:"\u{1F4CD}",bg:"#172033",color:"#94A3B8",border:"#29364B"}}async function ei(e){let t=de(),n=ie(t,"hygiene_standards","write"),a=ie(t,"hygiene_standards","delete"),o=ie(t,"export");e.innerHTML=`
    <!-- Header Section -->
    <div class="page-header" style="display:flex; justify-content:space-between; align-items:flex-start; flex-wrap:wrap; gap:16px; margin-bottom:20px;">
      <div>
        <h1 class="page-title" style="display:flex; align-items:center; gap:8px; margin:0 0 6px 0; font-size:1.6rem; font-weight:700; color:var(--text-1);">
          <span>\u2728</span> Master Standar Kebersihan Fisik
        </h1>
        <p class="page-subtitle" style="margin:0; font-size:0.9rem; color:var(--text-3);">
          Acuan Standar Kebersihan Fisik Seluruh Area & Ruangan Klinik (Audy Dental \xB7 Facility Care)
        </p>
      </div>
      <div style="display:flex; gap:10px; align-items:center; flex-wrap:wrap;">
        <div class="btn-group" style="background:#172033; padding:3px; border-radius:8px; display:inline-flex; border:1px solid #29364B;">
          <button id="view-table-btn" class="btn btn-sm" style="background:#2563EB; color:#fff; font-weight:600; border-radius:6px; box-shadow:0 1px 2px rgba(0,0,0,0.2); padding:6px 12px; cursor:pointer;">
            \u{1F4CB} Tabel Standar
          </button>
          <button id="view-zone-btn" class="btn btn-sm btn-ghost" style="color:#94A3B8; font-weight:600; border-radius:6px; padding:6px 12px; cursor:pointer;">
            \u{1F5C2}\uFE0F Per Zona Area
          </button>
        </div>
        <button id="btn-print-hygiene" class="btn btn-outline" style="background:#172033; border:1px solid #29364B; color:#F8FAFC; font-weight:600; padding:8px 14px; border-radius:8px; display:inline-flex; align-items:center; gap:6px; cursor:pointer;" title="Cetak Format SOP Standar Kebersihan">
          \u{1F5A8}\uFE0F Cetak SOP
        </button>
        ${o?`
        <button id="btn-export-hygiene" class="btn btn-outline" style="background:#172033; border:1px solid #29364B; color:#F8FAFC; font-weight:600; padding:8px 14px; border-radius:8px; display:inline-flex; align-items:center; gap:6px; cursor:pointer;">
          \u{1F4E5} Export Excel
        </button>
        `:""}
        ${n?`
        <button id="btn-add-hygiene" class="btn btn-primary" style="font-weight:600; padding:8px 16px; border-radius:8px; display:inline-flex; align-items:center; gap:6px; cursor:pointer;">
          + Tambah Parameter
        </button>
        `:""}
      </div>
    </div>

    <!-- KPI Summary Cards -->
    <div id="hygiene-summary-cards" style="display:grid; grid-template-columns:repeat(auto-fit, minmax(220px, 1fr)); gap:16px; margin-bottom:24px;">
      <div class="card" style="padding:16px 20px; border-radius:12px; border:1px solid #29364B; background:#172033; display:flex; align-items:center; gap:16px; box-shadow:0 4px 12px rgba(0,0,0,0.25);">
        <div style="width:48px; height:48px; border-radius:10px; background:rgba(37,99,235,0.2); color:#60A5FA; display:flex; align-items:center; justify-content:center; font-size:1.5rem; flex-shrink:0;">
          \u{1F4CB}
        </div>
        <div>
          <div style="font-size:0.8rem; color:var(--text-3); font-weight:600; text-transform:uppercase; letter-spacing:0.5px;">Total Parameter</div>
          <div style="font-size:1.6rem; font-weight:800; color:var(--text-1);" id="card-total-params">104</div>
          <div style="font-size:0.75rem; color:#16A34A; font-weight:500;">\u2713 Terverifikasi Standar SOP</div>
        </div>
      </div>
      <div class="card" style="padding:16px 20px; border-radius:12px; border:1px solid #29364B; background:#172033; display:flex; align-items:center; gap:16px; box-shadow:0 4px 12px rgba(0,0,0,0.25);">
        <div style="width:48px; height:48px; border-radius:10px; background:rgba(139,92,246,0.2); color:#C084FC; display:flex; align-items:center; justify-content:center; font-size:1.5rem; flex-shrink:0;">
          \u{1F3E2}
        </div>
        <div>
          <div style="font-size:0.8rem; color:var(--text-3); font-weight:600; text-transform:uppercase; letter-spacing:0.5px;">Kategori Zona</div>
          <div style="font-size:1.6rem; font-weight:800; color:var(--text-1);" id="card-total-zones">8</div>
          <div style="font-size:0.75rem; color:var(--text-3); font-weight:500;">Klinis, Publik, Sanitair, dll.</div>
        </div>
      </div>
      <div class="card" style="padding:16px 20px; border-radius:12px; border:1px solid #29364B; background:#172033; display:flex; align-items:center; gap:16px; box-shadow:0 4px 12px rgba(0,0,0,0.25);">
        <div style="width:48px; height:48px; border-radius:10px; background:rgba(16,185,129,0.2); color:#34D399; display:flex; align-items:center; justify-content:center; font-size:1.5rem; flex-shrink:0;">
          \u{1F6AA}
        </div>
        <div>
          <div style="font-size:0.8rem; color:var(--text-3); font-weight:600; text-transform:uppercase; letter-spacing:0.5px;">Ruangan / Lokasi</div>
          <div style="font-size:1.6rem; font-weight:800; color:var(--text-1);" id="card-total-rooms">35</div>
          <div style="font-size:0.75rem; color:#059669; font-weight:500;">Termasuk Ruang Tindakan</div>
        </div>
      </div>
      <div class="card" style="padding:16px 20px; border-radius:12px; border:1px solid #29364B; background:#172033; display:flex; align-items:center; gap:16px; box-shadow:0 4px 12px rgba(0,0,0,0.25);">
        <div style="width:48px; height:48px; border-radius:10px; background:rgba(245,158,11,0.2); color:#FBBF24; display:flex; align-items:center; justify-content:center; font-size:1.5rem; flex-shrink:0;">
          \u{1F6E1}\uFE0F
        </div>
        <div>
          <div style="font-size:0.8rem; color:var(--text-3); font-weight:600; text-transform:uppercase; letter-spacing:0.5px;">Kondisi Fisik</div>
          <div style="font-size:1.6rem; font-weight:800; color:#16A34A;">100%</div>
          <div style="font-size:0.75rem; color:var(--text-3); font-weight:500;">Bebas Debu, Kerak & Bau</div>
        </div>
      </div>
    </div>

    <!-- Zona Switcher Tabs -->
    <div style="overflow-x:auto; padding-bottom:8px; margin-bottom:18px;">
      <div id="hygiene-zone-tabs" style="display:flex; gap:8px; min-width:max-content; border-bottom:2px solid #29364B; padding-bottom:8px;">
        <button class="zone-tab-btn active" data-zone="all" style="border:none; background:#2563EB; color:#fff; font-weight:700; padding:8px 16px; border-radius:8px; cursor:pointer; display:inline-flex; align-items:center; gap:6px; font-size:0.88rem;">
          \u{1F31F} Semua Zona <span class="badge" style="background:rgba(255,255,255,0.25); color:#fff; font-size:0.75rem; padding:2px 6px; border-radius:10px;" id="badge-count-all">104</span>
        </button>
        ${Object.entries(ht).map(([k,x])=>`
          <button class="zone-tab-btn" data-zone="${k}" style="border:1px solid #29364B; background:#172033; color:#94A3B8; font-weight:600; padding:8px 16px; border-radius:8px; cursor:pointer; display:inline-flex; align-items:center; gap:6px; font-size:0.88rem;">
            <span>${x.icon}</span> ${k}
            <span class="badge" style="background:rgba(255,255,255,0.08); color:#F8FAFC; font-size:0.75rem; padding:2px 6px; border-radius:10px;" id="badge-count-${k.replace(/[\s&]+/g,"-")}">-</span>
          </button>
        `).join("")}
      </div>
    </div>

    <!-- Filter Toolbar -->
    <div class="card" style="padding:14px 18px; margin-bottom:20px; background:#111827; border:1px solid #29364B; border-radius:12px; box-shadow:0 4px 12px rgba(0,0,0,0.25);">
      <div style="display:flex; flex-wrap:wrap; gap:12px; align-items:center; justify-content:space-between;">
        <div style="display:flex; flex-wrap:wrap; gap:12px; align-items:center; flex:1; min-width:280px;">
          <div style="position:relative; flex:1; min-width:220px; max-width:360px;">
            <input type="text" id="hygiene-search" class="form-control" placeholder="Cari ruangan, item objek, atau standar kebersihan..." style="padding-left:36px; height:40px; border-radius:8px; font-size:0.9rem; width:100%; background:#1B2940; border:1px solid #29364B; color:#F8FAFC;">
            <span style="position:absolute; left:12px; top:50%; transform:translateY(-50%); color:var(--text-3); font-size:0.95rem;">\u{1F50D}</span>
          </div>
          <select id="hygiene-room-filter" class="form-control" style="width:auto; height:40px; border-radius:8px; font-size:0.875rem; min-width:180px; background:#172033; border:1px solid #29364B; color:#F8FAFC;">
            <option value="">Semua Ruangan</option>
          </select>
          <button id="btn-reset-filters" class="btn btn-sm btn-ghost" style="color:var(--text-3); font-weight:600; cursor:pointer; padding:6px 10px;">
            \u2715 Reset
          </button>
        </div>
        <div id="hygiene-stats-text" style="font-size:0.85rem; color:var(--text-3); font-weight:500;">
          Memuat data...
        </div>
      </div>
    </div>

    <!-- Main Content Container -->
    <div id="hygiene-view-container">
      <div class="skeleton" style="height:320px; border-radius:12px;"></div>
    </div>
  `;let s="all",l="table",h=[];async function p(){try{let k=await X("/api/hygiene-standards?all=1");if(!k.ok)throw new Error("Gagal memuat standar hygiene");h=Array.isArray(k.data)?k.data:k.data?.data||[],E(),c(),T()}catch(k){console.error(k),z(k.message||"Terjadi kesalahan memuat data");let x=e.querySelector("#hygiene-view-container");x&&(x.innerHTML=`
          <div class="empty-state" style="text-align:center; padding:40px; background:#172033; border-radius:12px; border:1px solid #29364B;">
            <div style="font-size:2.5rem; margin-bottom:12px;">\u26A0\uFE0F</div>
            <h3 style="margin:0 0 8px 0; color:var(--text-1);">Gagal Memuat Data</h3>
            <p style="color:var(--text-3); margin-bottom:16px;">${k.message}</p>
            <button id="btn-retry-hygiene" class="btn btn-primary" style="padding:8px 16px; border-radius:8px;">Coba Lagi</button>
          </div>
        `,x.querySelector("#btn-retry-hygiene")?.addEventListener("click",p))}}function c(){let k=e.querySelector("#badge-count-all");k&&(k.textContent=h.length),Object.entries(ht).forEach(([B,P])=>{let f=B.replace(/[\s&]+/g,"-"),g=e.querySelector(`#badge-count-${f}`);if(g){let D=h.filter(O=>P.rooms.includes(O.room_name)).length;g.textContent=D}});let x=new Set(h.map(B=>B.room_name)),_=e.querySelector("#card-total-rooms");_&&(_.textContent=x.size)}function E(){let k=e.querySelector("#hygiene-room-filter");if(!k)return;let x=k.value,_=[];if(s==="all")_=Array.from(new Set(h.map(B=>B.room_name))).filter(Boolean).sort();else{let B=ht[s];B&&(_=B.rooms.filter(P=>h.some(f=>f.room_name===P)))}k.innerHTML=`<option value="">Semua Ruangan (${_.length})</option>`+_.map(B=>`<option value="${B}" ${B===x?"selected":""}>${B}</option>`).join("")}function C(){let k=(e.querySelector("#hygiene-search")?.value||"").toLowerCase().trim(),x=e.querySelector("#hygiene-room-filter")?.value||"";return h.filter(_=>{if(s!=="all"){let B=ht[s];if(!B||!B.rooms.includes(_.room_name))return!1}if(x&&_.room_name!==x)return!1;if(k){let B=(_.room_name||"").toLowerCase(),P=(_.item_name||"").toLowerCase(),f=(_.cleanliness_standard||"").toLowerCase();if(!B.includes(k)&&!P.includes(k)&&!f.includes(k))return!1}return!0})}function T(){let k=C(),x=e.querySelector("#hygiene-stats-text");x&&(x.innerHTML=`Menampilkan <strong>${k.length}</strong> dari <strong>${h.length}</strong> parameter`);let _=e.querySelector("#hygiene-view-container");if(_){if(k.length===0){_.innerHTML=`
        <div class="empty-state" style="text-align:center; padding:50px 20px; background:#172033; border-radius:12px; border:1px solid #29364B;">
          <div style="font-size:3rem; margin-bottom:12px;">\u{1F50D}</div>
          <h3 style="margin:0 0 8px 0; color:var(--text-1); font-weight:700;">Tidak Ada Parameter yang Cocok</h3>
          <p style="color:var(--text-3); max-width:400px; margin:0 auto 16px auto; font-size:0.9rem;">
            Tidak ditemukan parameter kebersihan dengan kata kunci atau filter yang Anda pilih.
          </p>
          <button id="btn-clear-search-empty" class="btn btn-outline" style="border-radius:8px; padding:8px 16px; font-weight:600;">
            Reset Filter
          </button>
        </div>
      `,_.querySelector("#btn-clear-search-empty")?.addEventListener("click",()=>{let B=e.querySelector("#hygiene-search"),P=e.querySelector("#hygiene-room-filter");B&&(B.value=""),P&&(P.value=""),s="all",e.querySelectorAll(".zone-tab-btn").forEach(f=>f.classList.toggle("active",f.dataset.zone==="all")),e.querySelectorAll(".zone-tab-btn").forEach(f=>{f.dataset.zone==="all"?(f.style.background="#2563EB",f.style.color="#fff"):(f.style.background="#172033",f.style.color="#94A3B8")}),E(),T()});return}l==="table"?u(_,k):F(_,k)}}function u(k,x){k.innerHTML=`
      <style>.hygiene-std-row:hover{background-color:#22314A}</style>
      <div class="card" style="background:#172033; border:1px solid #29364B; border-radius:12px; overflow:hidden; box-shadow:0 4px 12px rgba(0,0,0,0.25);">
        <div style="overflow-x:auto;">
          <table style="width:100%; border-collapse:collapse; text-align:left; font-size:0.9rem;">
            <thead>
              <tr style="background:#1B2940; border-bottom:1px solid #29364B; color:#94A3B8; font-size:0.78rem; text-transform:uppercase; letter-spacing:0.6px;">
                <th style="padding:14px 16px; width:50px; text-align:center;">No</th>
                <th style="padding:14px 16px; width:220px;">Ruangan & Zona</th>
                <th style="padding:14px 16px; width:280px;">Item / Objek Inspeksi</th>
                <th style="padding:14px 16px;">Standar Kebersihan Fisik</th>
                ${n||a?'<th style="padding:14px 16px; width:100px; text-align:center;">Aksi</th>':""}
              </tr>
            </thead>
            <tbody>
              ${x.map((_,B)=>{let P=za(_.room_name);return`
                  <tr class="hygiene-std-row" style="border-bottom:1px solid #29364B; transition:background-color 0.15s ease;">
                    <td style="padding:14px 16px; text-align:center; color:var(--text-3); font-weight:600; font-size:0.85rem;">
                      ${B+1}
                    </td>
                    <td style="padding:14px 16px; vertical-align:top;">
                      <div style="margin-bottom:4px;">
                        <span style="display:inline-flex; align-items:center; gap:6px; background:${P.bg}; color:${P.color}; border:1px solid ${P.border}; padding:4px 10px; border-radius:6px; font-weight:700; font-size:0.83rem;">
                          <span>${P.icon}</span> ${_.room_name}
                        </span>
                      </div>
                      <div style="font-size:0.75rem; color:var(--text-3); font-weight:500;">
                        ${P.zoneName}
                      </div>
                    </td>
                    <td style="padding:14px 16px; vertical-align:top;">
                      <div style="font-weight:700; color:var(--text-1); line-height:1.45; font-size:0.9rem;">
                        ${_.item_name}
                      </div>
                    </td>
                    <td style="padding:14px 16px; vertical-align:top;">
                      <div style="background:#111827; border:1px solid #29364B; border-left:3px solid #10B981; border-radius:6px; padding:10px 12px; line-height:1.55; color:#F8FAFC; font-size:0.875rem;">
                        <span style="color:#059669; font-weight:700; margin-right:4px;">\u2713</span> ${_.cleanliness_standard}
                      </div>
                    </td>
                    ${n||a?`
                      <td style="padding:14px 16px; text-align:center; vertical-align:middle;">
                        <div style="display:inline-flex; gap:6px;">
                          ${n?`
                            <button class="btn btn-sm btn-ghost btn-edit-item" data-id="${_.id}" title="Edit Parameter" style="padding:6px 8px; border-radius:6px; color:#2563EB; cursor:pointer;">
                              \u270F\uFE0F
                            </button>
                          `:""}
                          ${a?`
                            <button class="btn btn-sm btn-ghost btn-delete-item" data-id="${_.id}" data-name="${_.item_name}" title="Hapus Parameter" style="padding:6px 8px; border-radius:6px; color:#DC2626; cursor:pointer;">
                              \u{1F5D1}\uFE0F
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
    `,R(k)}function F(k,x){let _={};x.forEach(B=>{let P=za(B.room_name);_[P.zoneName]||(_[P.zoneName]={info:P,rooms:{}}),_[P.zoneName].rooms[B.room_name]||(_[P.zoneName].rooms[B.room_name]=[]),_[P.zoneName].rooms[B.room_name].push(B)}),k.innerHTML=`
      <div style="display:flex; flex-direction:column; gap:24px;">
        ${Object.entries(_).map(([B,P])=>{let f=Object.values(P.rooms).reduce((g,D)=>g+D.length,0);return`
            <div class="card" style="background:#172033; border:1px solid #29364B; border-radius:14px; overflow:hidden; box-shadow:0 4px 12px rgba(0,0,0,0.25);">
              <!-- Zone Card Header -->
              <div style="background:${P.info.bg}; border-bottom:1px solid ${P.info.border}; padding:14px 20px; display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:10px;">
                <div style="display:flex; align-items:center; gap:10px;">
                  <span style="font-size:1.4rem;">${P.info.icon}</span>
                  <div>
                    <h3 style="margin:0; font-size:1.1rem; font-weight:700; color:${P.info.color};">
                      ${B}
                    </h3>
                    <div style="font-size:0.78rem; color:${P.info.color}; opacity:0.85;">
                      ${Object.keys(P.rooms).length} Ruangan \xB7 ${f} Parameter Standar
                    </div>
                  </div>
                </div>
                <span class="badge" style="background:rgba(255,255,255,0.08); color:#F8FAFC; border:1px solid #29364B; font-weight:700; padding:4px 10px; border-radius:12px; font-size:0.8rem;">
                  ${f} Parameter
                </span>
              </div>

              <!-- Zone Rooms List -->
              <div style="padding:16px 20px;">
                ${Object.entries(P.rooms).map(([g,D])=>`
                  <div style="margin-bottom:20px; last-child:margin-bottom:0;">
                    <div style="display:flex; align-items:center; gap:8px; margin-bottom:10px;">
                      <h4 style="margin:0; font-size:0.95rem; font-weight:700; color:var(--text-1);">
                        \u{1F6AA} ${g}
                      </h4>
                      <span class="badge badge-sm" style="background:#172033; color:#94A3B8; border:1px solid #29364B; font-size:0.75rem;">
                        ${D.length} item
                      </span>
                    </div>

                    <div style="border:1px solid #29364B; border-radius:8px; overflow:hidden;">
                      <table style="width:100%; border-collapse:collapse; font-size:0.875rem;">
                        <thead>
                          <tr style="background:#1B2940; border-bottom:1px solid #29364B; color:#94A3B8; font-size:0.75rem; text-transform:uppercase;">
                            <th style="padding:8px 12px; width:40px; text-align:center;">No</th>
                            <th style="padding:8px 12px; width:260px;">Item / Objek Inspeksi</th>
                            <th style="padding:8px 12px;">Standar Kebersihan Fisik</th>
                            ${n||a?'<th style="padding:8px 12px; width:80px; text-align:center;">Aksi</th>':""}
                          </tr>
                        </thead>
                        <tbody>
                          ${D.map((O,W)=>`
                            <tr style="border-bottom:1px solid #29364B;">
                              <td style="padding:10px 12px; text-align:center; color:var(--text-3); font-size:0.8rem;">
                                ${W+1}
                              </td>
                              <td style="padding:10px 12px; font-weight:600; color:var(--text-1); vertical-align:top;">
                                ${O.item_name}
                              </td>
                              <td style="padding:10px 12px; vertical-align:top; color:#334155;">
                                <div style="display:flex; align-items:flex-start; gap:6px;">
                                  <span style="color:#10B981; font-weight:700;">\u2713</span>
                                  <span>${O.cleanliness_standard}</span>
                                </div>
                              </td>
                              ${n||a?`
                                <td style="padding:10px 12px; text-align:center; vertical-align:middle;">
                                  <div style="display:inline-flex; gap:4px;">
                                    ${n?`<button class="btn btn-sm btn-ghost btn-edit-item" data-id="${O.id}" title="Edit" style="padding:4px 6px; cursor:pointer;">\u270F\uFE0F</button>`:""}
                                    ${a?`<button class="btn btn-sm btn-ghost btn-delete-item" data-id="${O.id}" data-name="${O.item_name}" title="Hapus" style="padding:4px 6px; color:#DC2626; cursor:pointer;">\u{1F5D1}\uFE0F</button>`:""}
                                  </div>
                                </td>
                              `:""}
                            </tr>
                          `).join("")}
                        </tbody>
                      </table>
                    </div>
                  </div>
                `).join("")}
              </div>
            </div>
          `}).join("")}
      </div>
    `,R(k)}function R(k){k.querySelectorAll(".btn-edit-item").forEach(x=>{x.addEventListener("click",()=>{let _=x.dataset.id,B=h.find(P=>String(P.id)===String(_));B&&K(B)})}),k.querySelectorAll(".btn-delete-item").forEach(x=>{x.addEventListener("click",()=>{if(!a){z("Hanya Superadmin yang dapat menghapus parameter standar.");return}let _=x.dataset.id,B=x.dataset.name;me(`Yakin ingin menghapus parameter standar <strong>"${B}"</strong>?`,async()=>{try{let P=await X(`/api/hygiene-standards/${_}`,{method:"DELETE"});if(!P.ok)throw new Error(P.data?.error||"Gagal menghapus parameter");ae("Parameter kebersihan berhasil dihapus"),p()}catch(P){z(P.message)}},"Hapus Parameter Standar")})})}function K(k=null){if(!n){z("Anda tidak memiliki izin untuk menambah atau mengedit parameter kebersihan.");return}let x=!!k,_=Array.from(new Set(h.map(P=>P.room_name))).filter(Boolean).sort(),B=document.createElement("div");B.innerHTML=`
      <form id="form-hygiene-item" style="display:flex; flex-direction:column; gap:16px;">
        <div>
          <label style="display:block; font-size:0.85rem; font-weight:600; color:var(--text-1); margin-bottom:6px;">
            Ruangan / Lokasi <span style="color:#DC2626;">*</span>
          </label>
          <input type="text" id="modal-room-name" list="room-suggestions" class="form-control" placeholder="Contoh: Ruang Tindakan, Lobby 1, Toilet 1..." value="${k?.room_name||""}" required style="width:100%; height:40px; border-radius:8px;">
          <datalist id="room-suggestions">
            ${_.map(P=>`<option value="${P}">`).join("")}
          </datalist>
          <span style="font-size:0.75rem; color:var(--text-3); margin-top:4px; display:block;">Pilih dari daftar atau ketik nama ruangan baru</span>
        </div>
        <div>
          <label style="display:block; font-size:0.85rem; font-weight:600; color:var(--text-1); margin-bottom:6px;">
            Item / Objek Inspeksi <span style="color:#DC2626;">*</span>
          </label>
          <input type="text" id="modal-item-name" class="form-control" placeholder="Contoh: Eksterior Dental Unit & Kursi, Lantai & Sudut..." value="${k?.item_name||""}" required style="width:100%; height:40px; border-radius:8px;">
        </div>
        <div>
          <label style="display:block; font-size:0.85rem; font-weight:600; color:var(--text-1); margin-bottom:6px;">
            Standar Kebersihan Fisik <span style="color:#DC2626;">*</span>
          </label>
          <textarea id="modal-standard" class="form-control" rows="4" placeholder="Deskripsi kondisi fisik kebersihan yang dipersyaratkan..." required style="width:100%; border-radius:8px; line-height:1.5;">${k?.cleanliness_standard||""}</textarea>
          <span style="font-size:0.75rem; color:var(--text-3); margin-top:4px; display:block;">Jelaskan kriteria fisik: bebas debu kasat mata, steril, tidak licin, tidak berbau, dll.</span>
        </div>
      </form>
    `,pe({title:x?"\u270F\uFE0F Edit Parameter Kebersihan":"\u2795 Tambah Parameter Kebersihan",content:B,confirmText:x?"Simpan Perubahan":"Tambah Parameter",size:"md",onConfirm:async(P,f)=>{let g=B.querySelector("#modal-room-name").value.trim(),D=B.querySelector("#modal-item-name").value.trim(),O=B.querySelector("#modal-standard").value.trim();if(!g||!D||!O){z("Semua kolom wajib diisi");return}try{let W={room_name:g,item_name:D,cleanliness_standard:O},J=x?`/api/hygiene-standards/${k.id}`:"/api/hygiene-standards",m=await X(J,{method:x?"PUT":"POST",body:JSON.stringify(W)});if(!m.ok)throw new Error(m.data?.error||"Gagal menyimpan parameter");ae(x?"Parameter berhasil diperbarui":"Parameter berhasil ditambahkan"),f(),p()}catch(W){z(W.message)}}})}function I(){let k=window.open("","_blank");if(!k){z("Pop-up printer diblokir oleh browser. Izinkan pop-up untuk mencetak SOP.");return}let _=`
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
            <div>Total: <strong>${h.length} Parameter Fisik</strong></div>
          </div>
        </div>

        <div class="title">
          <h1>STANDAR ACUAN KEBERSIHAN FISIK SELURUH CABANG</h1>
          <p>Buku Panduan Kriteria Kebersihan Fisik 8 Zona Area (Cleaning Service & Facility Care)</p>
        </div>

        ${Object.entries(ht).map(([B,P])=>{let f=h.filter(g=>P.rooms.includes(g.room_name));return f.length===0?"":`
            <div class="zone-title">${P.icon} ${B.toUpperCase()} (${f.length} PARAMETER)</div>
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
                ${f.map((g,D)=>`
                  <tr>
                    <td class="td-no">${D+1}</td>
                    <td class="td-room">${g.room_name}</td>
                    <td class="td-obj">${g.item_name}</td>
                    <td>${g.cleanliness_standard}</td>
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
    `;k.document.open(),k.document.write(_),k.document.close()}async function N(){try{let x=(await ue("/api/hygiene-standards/export",{})).map((B,P)=>{let f=za(B.room_name);return{No:P+1,"Kategori Zona":f.zoneName,"Ruangan / Lokasi":B.room_name,"Item / Objek Inspeksi":B.item_name,"Standar Kebersihan Fisik":B.cleanliness_standard}}),_=new Date().toISOString().slice(0,10);ge(x,`Master_Standar_Kebersihan_FCMS_${_}`),ae("Data master kebersihan berhasil diekspor ke Excel")}catch(k){z("Gagal mengekspor data: "+k.message)}}e.querySelector("#hygiene-search")?.addEventListener("input",()=>T()),e.querySelector("#hygiene-room-filter")?.addEventListener("change",()=>T()),e.querySelector("#btn-reset-filters")?.addEventListener("click",()=>{let k=e.querySelector("#hygiene-search"),x=e.querySelector("#hygiene-room-filter");k&&(k.value=""),x&&(x.value=""),T()}),e.querySelectorAll(".zone-tab-btn").forEach(k=>{k.addEventListener("click",()=>{s=k.dataset.zone,e.querySelectorAll(".zone-tab-btn").forEach(x=>{let _=x.dataset.zone===s;if(x.classList.toggle("active",_),_){x.style.background="#2563EB",x.style.color="#fff";let B=x.querySelector(".badge");B&&(B.style.background="rgba(255,255,255,0.25)",B.style.color="#fff")}else{x.style.background="#172033",x.style.color="#94A3B8";let B=x.querySelector(".badge");B&&(B.style.background="rgba(255,255,255,0.08)",B.style.color="#F8FAFC")}}),E(),T()})});let w=e.querySelector("#view-table-btn"),M=e.querySelector("#view-zone-btn");w?.addEventListener("click",()=>{l="table",w.style.background="#2563EB",w.style.color="#fff",M.style.background="transparent",M.style.color="#94A3B8",T()}),M?.addEventListener("click",()=>{l="zone",M.style.background="#2563EB",M.style.color="#fff",w.style.background="transparent",w.style.color="#94A3B8",T()}),e.querySelector("#btn-print-hygiene")?.addEventListener("click",I),e.querySelector("#btn-export-hygiene")?.addEventListener("click",N),e.querySelector("#btn-add-hygiene")?.addEventListener("click",()=>K(null)),p()}we();Ne();async function ti(e){let t=de();if(!t||t.role!=="superadmin"){e.innerHTML='<div class="empty-state"><p class="text-danger">Akses ditolak.</p></div>';return}De({container:e,title:"Manajemen User",icon:"\u{1F510}",apiPath:"/api/users",bulkDelete:!0,enableMobileFilterSheet:!0,itemLabel:"User",columns:[{key:"full_name",label:"Nama Lengkap"},{key:"username",label:"Username"},{key:"email",label:"Email"},{key:"role",label:"Role",render:n=>`<span class="badge ${{superadmin:"badge-danger",fc_specialist_leader:"badge-info",editor_khusus:"badge-info",viewer:"badge-success"}[n]||"badge-neutral"}">${Y(n)}</span>`},{key:"is_active",label:"Status",render:n=>n?'<span class="badge badge-success">Aktif</span>':'<span class="badge badge-neutral">Nonaktif</span>'},{key:"created_at",label:"Dibuat",nowrap:!0,render:n=>n?new Date(n).toLocaleDateString("id-ID"):"-"}],filterFields:[{type:"search",placeholder:"Cari nama / username..."}],formFields:n=>{let a=!!n,o=String(n?.role||"").trim().toLowerCase(),s={value:"superadmin",label:"Superadmin (Akses Penuh - Maks. 2 Aktif)"},l={value:"fc_specialist_leader",label:"FC Specialist / Leader (View All + Input & Edit Lembur & Issue - Maks. 2 Aktif)"},h={value:"viewer",label:"Viewer (Read-Only Semua Modul Aktif - Maks. 5 Aktif)"},p=a&&o==="superadmin"?[s,l,h]:[l,h,s],c=a&&(o==="superadmin"||o==="fc_specialist_leader"||o==="viewer")?o:"fc_specialist_leader";return[{type:"row",fields:[{name:"full_name",label:"Nama Lengkap",required:!0,placeholder:"Nama lengkap",value:n?.full_name},{name:"username",label:"Username",required:!0,placeholder:"username",value:n?.username,autocomplete:"off"}]},{type:"row",fields:[{name:"email",label:"Email",type:"email",required:!0,placeholder:"email@contoh.com",value:n?.email,autocomplete:"off"},{name:"role",label:"Role Resmi",type:"select",required:!0,options:p,value:c}]},{type:"row",fields:[{name:"password",label:a?"Password Baru (kosongkan jika tidak diubah)":"Password",type:"password",required:!a,placeholder:"Min. 12 karakter",autocomplete:"new-password"},{name:"is_active",label:"Status Aktif",type:"checkbox",checkLabel:"User aktif",value:a?n?.is_active:1}]},{type:"html",html:`
            <div style="margin-top:8px; padding:10px 12px; background:#EFF6FF; border:1px solid #BFDBFE; border-radius:8px; font-size:0.8rem; color:#1E3A8A;">
              <strong>Kebijakan RBAC &amp; Keamanan Password:</strong> Role resmi FCMS hanya <code>superadmin</code> (maks. 2 aktif) dan <code>fc_specialist_leader</code> (maks. 2 aktif). Password wajib minimal 12 karakter (kombinasi huruf besar, huruf kecil, angka, dan simbol).
            </div>
          `}]},onBeforeSubmit:(n,a)=>(n.full_name&&(n.full_name=String(n.full_name).trim()),n.username&&(n.username=String(n.username).trim()),n.email&&(n.email=String(n.email).trim().toLowerCase()),a&&(!n.password||String(n.password).trim()==="")&&delete n.password,n.assigned_modules=null,n),exportOptions:{moduleName:"users",onExport:async n=>{let o=(await ue("/api/users/export",n)).map(s=>({"Nama Lengkap":s.full_name||"",Username:s.username||"",Email:s.email||"",Role:s.role||"",Status:s.is_active?"Aktif":"Nonaktif"}));ge(o,"Data_Users")},onTemplate:()=>{ge([{"Nama Lengkap":"FC Leader 01",Username:"fcleader01",Email:"fcleader@contoh.com",Role:"fc_specialist_leader",Password:""}],"Template_Import_Users"),typeof ct=="function"&&ct("Kolom Password pada template dikosongkan. Wajib isi password unik dan kuat (min. 12 karakter: huruf besar, kecil, angka, simbol) untuk setiap akun baru, atau kosongkan jika memperbarui akun existing tanpa mengubah password.")},onImport:async(n,a={})=>{let o=n.map(s=>({full_name:String(s["Nama Lengkap"]||"").trim(),username:String(s.Username||"").trim(),email:String(s.Email||"").trim().toLowerCase(),role:String(s.Role||"").trim()||"fc_specialist_leader",password:String(s.Password||"").trim()})).filter(s=>s.username||s.email||s.full_name||s.password);return await Le({module:"users",label:"Manajemen Pengguna",endpoint:"/api/users/import",body:o,idempotencyKey:a?.idempotencyKey,recover:!!a?.recover})}}})}we();async function ai(e){De({container:e,title:"Manajemen Cabang",icon:"\u{1F3E2}",apiPath:"/api/branches",enableMobileFilterSheet:!0,itemLabel:"Cabang",bulkDelete:!1,columns:[{key:"code",label:"Kode",width:"60px"},{key:"full_name",label:"Nama Cabang"},{key:"city",label:"Kota"},{key:"branch_scope",label:"Lingkup",render:(t,n)=>{let a=String(n?.code||"").trim();return a==="000"?'<span class="badge badge-info" title="Management / HQ (non-operasional, tersedia di dropdown modul)">HQ (Non-Ops)</span>':a==="A01"?'<span class="badge badge-warning" title="Gabung area 002. Kemang (tidak dihitung terpisah di Dashboard)">Gabung 002. Kemang</span>':/^(00[1-9]|0[1-6][0-9]|070)$/.test(a)?'<span class="badge badge-success">Operasional</span>':'<span class="badge badge-neutral">Non-Master</span>'}},{key:"is_active",label:"Status",render:t=>t?'<span class="badge badge-success">Aktif</span>':'<span class="badge badge-neutral">Nonaktif</span>'}],filterFields:[{type:"search",placeholder:"Cari nama / kode cabang..."}],formFields:t=>[{type:"row",fields:[{name:"code",label:"Kode Cabang",required:!0,placeholder:"001, A01, ...",value:t?.code},{name:"name",label:"Nama Pendek",required:!0,placeholder:"Pondok Bambu",value:t?.name}]},{name:"full_name",label:"Nama Lengkap",required:!0,placeholder:"001. Pondok Bambu",value:t?.full_name},{type:"row",fields:[{name:"city",label:"Kota",placeholder:"Jakarta",value:t?.city},{name:"is_active",label:"Status",type:"checkbox",checkLabel:"Cabang aktif",value:t?.is_active!==void 0?t.is_active:1}]}],exportOptions:{moduleName:"branches",onExport:async t=>{let n=await ue("/api/branches/export",t);ge(n,"Data_Cabang")},onTemplate:()=>{ge([{"Kode Cabang":"001","Nama Pendek":"Pondok Bambu","Nama Lengkap":"001. Pondok Bambu",Kota:"Jakarta Timur"},{"Kode Cabang":"002","Nama Pendek":"Bintaro","Nama Lengkap":"002. Bintaro",Kota:"Tangerang Selatan"}],"Template_Import_Cabang")},onImport:async(t,n={})=>{let a=t.map(o=>({code:String(o["Kode Cabang"]||"").trim(),name:String(o["Nama Pendek"]||"").trim(),full_name:String(o["Nama Lengkap"]||"").trim(),city:String(o.Kota||"").trim()})).filter(o=>o.code&&o.name);return await Le({module:"branches",label:"Data Cabang",endpoint:"/api/branches/import",body:a,idempotencyKey:n?.idempotencyKey,recover:!!n?.recover})}}})}async function ni(e){let t=new Date,n=[];e.innerHTML=`
    <div class="page-header">
      <h1 class="page-title">\u{1F4C5} Kalender</h1>
    </div>
    <div class="card">
      <div class="card-header calendar-nav">
        <button class="btn btn-ghost btn-sm" id="cal-prev">\u2039 Prev</button>
        <span class="calendar-month-label" id="cal-month-label"></span>
        <button class="btn btn-ghost btn-sm" id="cal-next">Next \u203A</button>
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
  `,document.getElementById("cal-prev").addEventListener("click",()=>{t.setMonth(t.getMonth()-1),o()}),document.getElementById("cal-next").addEventListener("click",()=>{t.setMonth(t.getMonth()+1),o()}),document.getElementById("cal-event-close").addEventListener("click",()=>{document.getElementById("cal-event-list").style.display="none"}),document.querySelectorAll(".cal-filter").forEach(l=>l.addEventListener("change",o));async function a(){try{let l=`${t.getFullYear()}-${String(t.getMonth()+1).padStart(2,"0")}`;n=(await X(`/api/dashboard/calendar?month=${l}`)).data?.data||[]}catch(l){console.warn("[Calendar] Failed to load events, rendering empty grid:",l),n=[]}}async function o(){let l=document.getElementById("calendar-grid");if(l){l.innerHTML=`<div style="display:grid;grid-template-columns:repeat(7,1fr);gap:1px;background:var(--border);">
      ${Array(35).fill('<div style="background:#111827;min-height:70px;"></div>').join("")}
    </div>`,await a();try{let h=t.getFullYear(),p=t.getMonth(),c=t.toLocaleDateString("id-ID",{month:"long",year:"numeric"}),E=document.getElementById("cal-month-label");E&&(E.textContent=c);let C=new Set(["inspection","cleaning","fogging","basecamp"]),T=new Set(Array.from(document.querySelectorAll(".cal-filter:checked")).map(x=>x.value)),u=n.filter(x=>T.has(x.type)||T.has("reports")&&C.has(x.type)),F={};u.forEach(x=>{let _=$r(x.event_date);_&&(F[_]||(F[_]=[]),F[_].push(x))});let R=new Date(h,p,1).getDay(),K=new Date(h,p+1,0).getDate(),I=["Min","Sen","Sel","Rab","Kam","Jum","Sab"],N=new Date().toISOString().slice(0,10),w='<div class="calendar-grid">';I.forEach(x=>{w+=`<div class="cal-day-header">${x}</div>`});for(let x=0;x<R;x++)w+='<div class="cal-cell cal-cell-empty"></div>';for(let x=1;x<=K;x++){let _=`${h}-${String(p+1).padStart(2,"0")}-${String(x).padStart(2,"0")}`,B=F[_]||[],P=_===N;w+=`
          <div class="cal-cell ${P?"cal-today":""} ${B.length?"cal-has-events":""}"
               data-date="${_}" tabindex="0" role="button" aria-label="${_}">
            <div class="cal-day-num ${P?"today-num":""}">${x}</div>
            <div class="cal-events-preview">
              ${B.slice(0,3).map(f=>`
                <div class="cal-event-dot cal-color-${f.color||"gray"}" title="${Re(f.title||f.type)}">
                  <span class="cal-event-dot-label">${Ar(f.title||f.branch_name||f.type,18)}</span>
                </div>
              `).join("")}
              ${B.length>3?`<div class="cal-more">+${B.length-3} lagi</div>`:""}
            </div>
          </div>`}let k=(R+K)%7;if(k!==0)for(let x=0;x<7-k;x++)w+='<div class="cal-cell cal-cell-empty"></div>';w+="</div>",l.innerHTML=w,l.querySelectorAll(".cal-cell[data-date]").forEach(x=>{x.addEventListener("click",()=>{let _=x.dataset.date,B=F[_]||[];if(!B.length)return;let P=document.getElementById("cal-event-list"),f=new Date(_+"T00:00:00").toLocaleDateString("id-ID",{weekday:"long",day:"numeric",month:"long",year:"numeric"});document.getElementById("cal-event-date").textContent=f,document.getElementById("cal-event-items").innerHTML=B.map(g=>`
            <div class="cal-event-item cal-color-border-${g.color||"gray"}">
              <div class="cal-event-type">
                ${Fr(g.type)}
                ${g.date_type?` <span class="badge" style="font-size:0.75rem;padding:2px 6px;background:rgba(0,0,0,0.06);margin-left:4px;">${Re(g.date_type)}</span>`:""}
              </div>
              <div class="cal-event-title" style="font-weight:600;font-size:0.95rem;margin-top:2px;">${Re(g.title||"-")}</div>
              ${g.branch_name?`<div class="cal-event-branch" style="color:var(--text-2);font-size:0.85rem;margin-top:2px;">\u{1F4CD} ${Re(g.branch_name)}</div>`:""}
              ${g.pic||g.period||g.opening_date||g.target_date||g.completion_date||g.notes?`
                <div class="cal-event-meta" style="margin-top:6px;padding-top:6px;border-top:1px dashed var(--border);font-size:0.8rem;color:var(--text-2);display:flex;flex-direction:column;gap:3px;">
                  ${g.pic?`<div>\u{1F464} <b>PIC:</b> ${Re(g.pic)}</div>`:""}
                  ${g.period?`<div>\u{1F4C5} <b>Periode:</b> ${Re(g.period)}</div>`:""}
                  ${g.opening_date?`<div>\u{1F6AA} <b>Tgl Opening:</b> ${Re(Ka(g.opening_date))}</div>`:""}
                  ${g.target_date?`<div>\u{1F3AF} <b>Tgl Target:</b> ${Re(Ka(g.target_date))}</div>`:""}
                  ${g.completion_date?`<div>\u2705 <b>Tgl Selesai:</b> ${Re(Ka(g.completion_date))}</div>`:""}
                  ${g.notes?`<div>\u{1F4DD} <b>Catatan:</b> ${Re(g.notes)}</div>`:""}
                </div>
              `:""}
              ${g.status?`<div class="cal-event-status" style="margin-top:6px;"><span class="badge ${g.status==="Done"?"badge-success":"badge-warning"}">${Re(g.status)}</span></div>`:""}
              ${g.days_remaining!==void 0?`<div class="cal-event-extra">Sisa: ${g.days_remaining} hari</div>`:""}
            </div>
          `).join(""),P.style.display="block"})})}catch(h){console.error("[Calendar] Render error:",h),l&&(l.innerHTML=`
          <div style="padding:40px;text-align:center;color:var(--text-3)">
            <div style="font-size:2rem;margin-bottom:8px">\u{1F4C5}</div>
            <div>Gagal memuat kalender. Silakan refresh.</div>
          </div>`)}}}let s=()=>{a().then(()=>o())};return Ft.on("data:changed",s),o(),()=>{Ft.off("data:changed",s)}}function $r(e){if(!e||e==="-"||String(e).trim()==="")return"";let t=String(e).trim();if(/^\d{4}-\d{2}-\d{2}/.test(t))return t.slice(0,10);let n=t.split(/[\/\-\.]/);if(n.length===3){let[a,o,s]=n.map(l=>l.trim());if(s.length===4)return`${s}-${o.padStart(2,"0")}-${a.padStart(2,"0")}`;if(a.length===4)return`${a}-${o.padStart(2,"0")}-${s.padStart(2,"0")}`}return t.slice(0,10)}function Ka(e){if(!e||e==="-"||String(e).trim()==="")return"";let n=String(e).trim().split(/[\/\-\.]/);if(n.length===3){let[a,o,s]=n.map(l=>l.trim());if(a.length===4)return`${s.padStart(2,"0")}-${o.padStart(2,"0")}-${a}`;if(s.length===4)return`${a.padStart(2,"0")}-${o.padStart(2,"0")}-${s}`}return e}function Ar(e,t){return e?e.length>t?e.slice(0,t)+"\u2026":e:""}function Re(e){return e?String(e).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;"):""}function Fr(e){return{schedule:"\u{1F5D3} Jadwal",issue:"\u26A0\uFE0F Permasalahan",training:"\u{1F393} Training",contract_expiry:"\u{1F4CB} Kontrak Habis",cleaning:"\u{1F9F9} Cleaning",inspection:"\u{1F50D} Inspeksi",fogging:"\u{1F4A8} Fogging",basecamp:"\u{1F3E0} Basecamp"}[e]||e}Ne();async function ii(e){let t=de(),n=(t?.full_name||t?.username||"U")[0].toUpperCase(),o={superadmin:"#7C3AED",fc_specialist_leader:"#0284C7",editor_khusus:"#0284C7",viewer:"#10B981"}[t?.role]||"#64748B";e.innerHTML=`
    <div class="page-header">
      <h1 class="page-title">\u{1F464} Profil Saya</h1>
    </div>

    <div class="profile-layout">

      <!-- LEFT: Info Card -->
      <div class="chart-card profile-info-card">
        <div class="profile-avatar-wrap">
          <div class="profile-avatar-xl" style="background:linear-gradient(135deg,${o},${o}99)">
            ${Y(n)}
          </div>
          <div class="profile-name-block">
            <div class="profile-fullname">${Y(t?.full_name||"\u2014")}</div>
            <div class="profile-username">@${Y(t?.username||"\u2014")}</div>
            <span class="badge badge-info" style="background:${o}18;color:${o};margin-top:6px">
              ${Y((t?.role||"-").toUpperCase())}
            </span>
          </div>
        </div>

        <hr class="profile-divider">

        <div class="info-list">
          <div class="info-row">
            <span class="info-key">\u{1F4E7} Email</span>
            <span class="info-value">${Y(t?.email||"\u2014")}</span>
          </div>
          <div class="info-row">
            <span class="info-key">\u{1F464} Username</span>
            <span class="info-value">${Y(t?.username||"\u2014")}</span>
          </div>
          <div class="info-row">
            <span class="info-key">\u{1F3AF} Role</span>
            <span class="info-value" style="color:${o};font-weight:700">${Y(t?.role||"\u2014")}</span>
          </div>
        </div>
      </div>

      <!-- RIGHT: Change Password -->
      <div class="chart-card">
        <div class="chart-card-header">
          <div>
            <div class="chart-card-title">\u{1F511} Ganti Password</div>
            <div class="chart-card-subtitle">Gunakan password yang kuat, minimal 12 karakter</div>
          </div>
        </div>

        <form id="change-pwd-form" novalidate style="margin-top:8px">
          <div class="form-group">
            <label class="form-label">Password Lama <span class="required">*</span></label>
            <input type="password" name="current_password" class="form-control"
              required placeholder="Masukkan password saat ini" autocomplete="current-password">
          </div>
          <div class="form-group">
            <label class="form-label">Password Baru <span class="required">*</span></label>
            <input type="password" name="new_password" class="form-control"
              required placeholder="Minimal 12 karakter" autocomplete="new-password">
          </div>
          <div class="form-group">
            <label class="form-label">Konfirmasi Password Baru <span class="required">*</span></label>
            <input type="password" name="confirm_password" class="form-control"
              required placeholder="Ulangi password baru" autocomplete="new-password">
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
  `;let s=document.getElementById("session-info");s&&(s.textContent="Aktif (HttpOnly Cookie + CSRF Terproteksi)"),document.getElementById("change-pwd-form")?.addEventListener("submit",async l=>{l.preventDefault();let h=document.getElementById("pwd-error"),p=document.getElementById("pwd-success"),c=document.getElementById("btn-save-pwd");h.style.display="none",p.style.display="none";let E=l.target,C=E.current_password.value,T=E.new_password.value,u=E.confirm_password.value;if(T!==u){h.textContent="\u274C Konfirmasi password tidak cocok.",h.style.display="block";return}if(T.length<12){h.textContent="\u274C Password baru minimal 12 karakter.",h.style.display="block";return}c.disabled=!0,c.textContent="\u23F3 Menyimpan...";let F=await X("/api/auth/change-password",{method:"POST",body:JSON.stringify({current_password:C,new_password:T})});c.disabled=!1,c.innerHTML='<svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24" style="margin-right:5px"><path d="M19 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11l5 5v11a2 2 0 0 1-2 2z"/><polyline points="17 21 17 13 7 13 7 21"/><polyline points="7 3 7 8 15 8"/></svg>Simpan Password',F.ok?(p.textContent="\u2705 Password berhasil diubah.",p.style.display="block",E.reset(),ae("Password berhasil diubah.")):(h.textContent=F.data?.error||"Gagal mengubah password.",h.style.display="block")}),document.getElementById("btn-logout")?.addEventListener("click",async()=>{if(confirm("Keluar dari semua sesi? Anda harus login ulang.")){try{await X("/api/auth/logout?all=true",{method:"POST"})}catch{}pt(),window.location.reload()}})}var Xt={Validasi:{module:"validation",label:"Master Referensi"},SOP:{module:"sop",label:"SOP"},"Master Karyawan":{module:"employees",label:"Karyawan"},"Data Kontrak":{module:"contracts",label:"Kontrak"},Permasalahan:{module:"issues",label:"Permasalahan"},"Time Line":{module:"schedule",label:"Jadwal Kegiatan"},"Report Inspeksi Hygiene 2026":{module:"inspection",label:"Laporan Inspeksi"},"Report GC-DC 2026":{module:"cleaning",label:"Laporan GC/DC"},"Report Fogging 2026":{module:"fogging",label:"Laporan Fogging"},"Rekap Laporan Basecamp":{module:"basecamp",label:"Rekap Basecamp"},Training:{module:"training",label:"Training"},"Master Checklist":{module:"checklist",label:"Checklist"},"Data SP":{module:"sp",label:"Data SP"},"Data Mutasi":{module:"mutasi",label:"Data Mutasi"}};function Ee(e){if(e==null||e==="")return null;if(e instanceof Date)return isNaN(e.getTime())?null:e.toISOString().slice(0,10);let t=String(e).trim();if(t===""||t==="0")return null;if(/^\d{4}-\d{2}-\d{2}/.test(t))return t.slice(0,10);if(/^\d{4,5}$/.test(t)){let o=Number(t);if(o>2e4&&o<99999){let s=new Date(Date.UTC(1899,11,30)+o*864e5);return isNaN(s.getTime())?null:s.toISOString().slice(0,10)}}let n=t.split(/[\/\-\.]/);if(n.length===3){let[o,s,l]=n.map(E=>E.trim()),h=Number(o),p=Number(s),c=Number(l);if(o.length===4&&h>1900)return`${o}-${s.padStart(2,"0")}-${l.padStart(2,"0")}`;if(l.length===4&&c>1900)return h>12?`${l}-${s.padStart(2,"0")}-${o.padStart(2,"0")}`:p>12?`${l}-${o.padStart(2,"0")}-${s.padStart(2,"0")}`:`${l}-${s.padStart(2,"0")}-${o.padStart(2,"0")}`;if(l.length===2&&!isNaN(c)){let E=c>=50?`19${l}`:`20${l}`;return h>12?`${E}-${s.padStart(2,"0")}-${o.padStart(2,"0")}`:`${E}-${s.padStart(2,"0")}-${o.padStart(2,"0")}`}}let a=new Date(t);return isNaN(a.getTime())?null:a.toISOString().slice(0,10)}function ri(e){return Object.values(e).every(t=>t==null||String(t).trim()==="")}var _r={validation:{required:[],map:e=>({cabang:e.CABANG,pic:e.PIC,kegiatan:e.KEGIATAN,quartal:e.QUARTAL,masa_pkwt:e["MASA PKWT"],pic_pelapor:e["PIC PELAPOR"],kontrak:e.KONTRAK})},sop:{required:[{key:"Nama SOP",label:"Nama SOP"}],map:e=>({name:e["Nama SOP"],category:e.Kategori||"Umum",document_link:e["Link Document"],version:"1.0",effective_date:null,notes:""})},employees:{required:[{key:"Nama Lengkap",label:"Nama Lengkap"}],map:e=>({full_name:e["Nama Lengkap"],branch_name:e.Cabang,division:e["Div / Bagian"]||e.Divisi||"FACILITY CARE",phone:e["No. Hp"]||e["No. HP"],join_date:Ee(e["Tanggal Masuk"]||e["Tgl Masuk"]),status:e.Status||"",target_pindah_os:Ee(e["Target Pindah OS"]),target_selesai_os:Ee(e["Target Selesai OS"]),notes:e.Catatan||""})},contracts:{required:[{key:"Nama Lengkap",label:"Nama Lengkap"}],map:e=>({employee_name:e["Nama Lengkap"],branch_name:e.Cabang,division:e["Div / Bagian"]||"FACILITY CARE",start_date:Ee(e["Tanggal Mulai"]),end_date:Ee(e["Tanggal Selesai"]),contract_type:e["Tipe Kontrak"]||"",pkwt_number:e.PKWT||"",status:e.Status||"",notes:e.keterangan})},issues:{required:[{key:"Keluhan",label:"Keluhan"}],map:e=>({report_date:Ee(e["Tanggal Info"]),branch_name:e.Cabang,category:e.Kategori,source:e["Sumber Laporan"],complaint:e.Keluhan,employee_name:e["Nama FC"],fc_specialist:e["FC Spesialis"],solution:e.Solusi,status:e.Status||"",completion_date:Ee(e["Tanggal Selesai"])})},one_on_one:{required:[],map:e=>({meeting_date:Ee(e.Tanggal),branch_name:e.Cabang,employee_name:e["Nama Karyawan"],pic:e.Pic,problem:e.Masalah,solution:e.Solusi,status:e.Status||"",completion_date:Ee(e["Tanggal Selesai"]),document_link:e["Link Document"]})},schedule:{required:[{key:"Kegiatan",label:"Kegiatan"}],map:e=>({branch_name:e.Cabang,activity_type:e.Kegiatan,period:e.Periode,pic:e.Pic||e.PIC,opening_date:Ee(e["Tanggal Opening"]||e["Tgl Opening"]),target_date:Ee(e["Tanggal Target"]||e["Tgl Target"]),completion_date:Ee(e["Tanggal Selesai"]||e["Tgl Selesai"]),status:e.Status||"",notes:e.Keterangan||e.Catatan})},inspection:{required:[],map:e=>{let t=e["Point FC SP"]!==void 0&&e["Point FC SP"]!==null&&e["Point FC SP"]!==""?e["Point FC SP"]:e["Point FC"];return{inspection_date:Ee(e.Tanggal),branch_name:e.Cabang,period:e.Periode,status:e.Status||"",fc_score:t!=null&&t!==""?parseFloat(String(t).replace(",",".")):null,spv_score:e["Point SPV"]!==void 0&&e["Point SPV"]!==null&&e["Point SPV"]!==""?parseFloat(String(e["Point SPV"]).replace(",",".")):null,document_link:e.Link||e["Link Dokumen"]||"",notes:e.Keterangan||""}}},cleaning:{required:[],map:e=>({activity_date:Ee(e.Tanggal),branch_name:e.Cabang,activity_type:e["Jenis Kegiatan"]||"General Cleaning",period:e.Periode,status:e.Status||"",document_link:e.Link||e["Link Dokumen"]||"",notes:e.Keterangan||""})},fogging:{required:[],map:e=>({activity_date:Ee(e.Tanggal),branch_name:e.Cabang,period:e.Periode,status:e.Status||"",document_link:e.Link||e["Link Dokumen"]||"",notes:e.Keterangan||""})},basecamp:{required:[{key:"Permasalahan",label:"Permasalahan"}],map:e=>({info_date:Ee(e["Tgl Info"]),branch_name:e.Cabang,problem:e.Permasalahan,pic:e.PIC,done_date:Ee(e["Tgl Done"]),status:e.Status||"",notes:e.Ket})},relievers:{required:[],map:e=>({branch_name:e.Cabang,original_fc_name:e["Nama Facility care"],period:e.Periode,reliever_name:e.Relifer,backup_date:Ee(e["Tanggal Back Up"]),completion_date:Ee(e["Tanggal Selesai"]),reason:e.Keterangan,shift:e.Shift,status:e.Status||""})},training:{required:[{key:"Materi",label:"Materi"}],map:e=>({training_date:Ee(e.Tanggal),batch:e.Batch,subject:e.Materi,participants:e.Peserta,branch_name:e.Cabang,trainer:e.Trainer,score:e.Nilai!==void 0&&e.Nilai!==null?parseFloat(String(e.Nilai).replace(",",".")):null,notes:""})},checklist:{required:[],map:e=>({name:e["Master Checklist"],category:"Umum",document_link:e["Link Document"],description:""})},supply:{required:[],map:e=>({submitted_at:Ee(e.Timestamp),submitter_name:e["Nama Lengkap"],branch_name:e["Kebutuhan Untuk Cabang"],tools_items:e["Alat - Alat / Barang"],tools_quantity:e["Jumlah Permintaan Alat / Barang"],chemical_items:e.Chemical,chemical_quantity:e["Jumlah Permintaan Chemical"],additional_notes:e["Tambahan  Alat / Chemical Jika Ada Permintaan Diluar List."],status:e.Status||""})},sp:{required:[{key:"Nama Karyawan",label:"Nama Karyawan"}],map:e=>({tanggal:Ee(e.Tanggal||e["Tanggal Terbit"]),employee_name:e["Nama Karyawan"],branch_name:e.Cabang,sp_type:e["Jenis SP"]||"SP 1",status:e.Status||"Aktif",document_link:e["Link Dokumen"]||e["Link Document"]||null,notes:e.Keterangan||e.Catatan||null})},mutasi:{required:[{key:"Nama Karyawan",label:"Nama Karyawan"}],map:e=>({tanggal:Ee(e.Tanggal||e["Tanggal Efektif"]),employee_name:e["Nama Karyawan"],from_branch_name:e["Cabang Asal"]||e["Dari Cabang"],to_branch_name:e["Cabang Tujuan"]||e["Ke Cabang"],status:e.Status||"Selesai",document_link:e["Link Dokumen"]||e["Link SK"]||null,notes:e.Keterangan||e.Alasan||null})},forms:{required:[{key:"Nama Form",label:"Nama Form"}],map:e=>({name:e["Nama Form"],category:e.Kategori||"Umum",document_link:e["Link Dokumen"]||e["Link Document"]||null,description:e.Deskripsi||e.Keterangan||null,is_public:e.Publik==="Ya"||e.Publik===1||e["Akses Publik"]===1?1:0})}};function Br(e,t){let n=Xt[e];if(!n)return{valid:[],errors:[],mapped:[],skipped:!0};let a=_r[n.module];if(!a)return{valid:[],errors:[],mapped:[],skipped:!0};let o=[],s=[],l=[];return t.filter(p=>!ri(p)).forEach((p,c)=>{let E=t.indexOf(p)+2,C=[];a.required.forEach(({key:u,label:F})=>{let R=p[u];if(R==null||String(R).trim()===""){let K=Object.keys(p).filter(I=>I.trim()).join(", ");C.push({column:F,originalValue:R||"",reason:`Kolom "${F}" wajib diisi dan tidak ditemukan`,hint:`Kolom yang tersedia: ${K.slice(0,120)}`})}});let T=a.map(p);C.length>0?s.push({row:E,data:T,raw:p,errors:C}):(o.push(p),l.push(T))}),{valid:o,errors:s,mapped:l}}function oi(e){let t=[];return e.SheetNames.forEach(n=>{let a=Xt[n];if(!a)return;let o=e.Sheets[n],s=window.XLSX.utils.sheet_to_json(o,{defval:"",raw:!1,dateNF:"yyyy-mm-dd"}),l=Br(n,s),h=s.filter(p=>!ri(p));t.push({sheetName:n,module:a.module,label:a.label,total:h.length,valid:l.mapped.length,errorCount:l.errors.length,errors:l.errors,mapped:l.mapped,skipped:!1})}),t}function si(){let e=window.XLSX,t=e.utils.book_new();Object.entries({Validasi:[{CABANG:"001. Pondok Bambu",PIC:"Berlin",KEGIATAN:"General Cleaning",QUARTAL:"Q1","PIC PELAPOR":"Berlin",KONTRAK:"PKWT 1","MASA PKWT":"1 Tahun"}],SOP:[{"Nama SOP":"SOP Pembersihan Toilet",Kategori:"Cleaning","Link Document":"https://..."}],"Master Karyawan":[{"Nama Lengkap":"Budi Santoso",Cabang:"000. Management","Div / Bagian":"FACILITY CARE","No. Hp":"081234567890","Tanggal Masuk":"2024-01-15",Status:"Aktif"}],"Data Kontrak":[{"Nama Lengkap":"Budi Santoso",Cabang:"000. Management","Div / Bagian":"FACILITY CARE","Tanggal Mulai":"2024-01-01","Tanggal Selesai":"2024-12-31","Tipe Kontrak":"PKWT 1",PKWT:"001/PKWT/2024",Status:"Aktif",keterangan:""}],Permasalahan:[{"Tanggal Info":"2024-03-01",Cabang:"001. Pondok Bambu",Kategori:"Cleaning","Sumber Laporan":"SPV",Keluhan:"Lantai kotor","Nama FC":"Budi","FC Spesialis":"Fajar",Solusi:"Teguran",Status:"Done","Tanggal Selesai":"2024-03-02"}],"Time Line":[{Cabang:"001. Pondok Bambu",Kegiatan:"General Cleaning",Periode:"Januari",Pic:"Berlin","Tanggal Opening":"2024-01-01","Tanggal Target":"2024-01-10","Tanggal Selesai":"2024-01-09",Status:"Done",Keterangan:""}],"Report Inspeksi Hygiene 2026":[{Tanggal:"2026-01-15",Cabang:"001. Pondok Bambu",Periode:"Q1",Status:"Done","Point FC SP":"85.5","Point SPV":"90.0",Link:"https://..."}],"Report GC-DC 2026":[{Tanggal:"2026-01-20",Cabang:"001. Pondok Bambu","Jenis Kegiatan":"General Cleaning",Periode:"Q1",Status:"Done",Link:"https://..."}],"Report Fogging 2026":[{Tanggal:"2026-01-25",Cabang:"001. Pondok Bambu",Periode:"Q1",Status:"Done",Link:"https://..."}],"Rekap Laporan Basecamp":[{"Tgl Info":"2024-02-01",Cabang:"001. Pondok Bambu",Permasalahan:"Lampu mati",PIC:"Berlin","Tgl Done":"2024-02-02",Status:"Done",Ket:""}],Training:[{Tanggal:"2024-04-10",Batch:"Batch 1",Materi:"Basic Cleaning",Peserta:"5",Cabang:"001. Pondok Bambu",Trainer:"Fajar",Nilai:"85",Keterangan:""}],"Master Checklist":[{"Master Checklist":"Checklist Kebersihan Toilet","Link Document":"https://..."}],"Data SP":[{"Nama Karyawan":"Budi Santoso",Cabang:"001. Pondok Bambu","Jenis SP":"SP 1","Tanggal Terbit":"2026-01-10","Tanggal Berakhir":"2026-07-10","Alasan / Pelanggaran":"Keterlambatan",Status:"Aktif","Link Dokumen":""}],"Data Mutasi":[{"Nama Karyawan":"Budi Santoso","Cabang Asal":"001. Pondok Bambu","Cabang Tujuan":"002. Bintaro","Tanggal Efektif":"2026-02-01","Alasan Mutasi":"Rotasi Reguler",Status:"Selesai","No. SK / Dokumen":"SK/001/2026"}],"Master Form":[{"Nama Form":"Form Pengajuan Alat",Kategori:"Operasional","Link Form":"https://...",Deskripsi:"Formulir pengajuan alat dan perlengkapan",Publik:"Ya"}]}).forEach(([a,o])=>{e.utils.book_append_sheet(t,e.utils.json_to_sheet(o),a)}),e.writeFile(t,"Template_Import_Data_Awal_FCMS.xlsx")}function li(e){let t=window.XLSX,n=t.utils.book_new(),a=!1;return e.forEach(o=>{if(!o.errors||o.errors.length===0)return;a=!0;let s=o.errors.map(h=>({"No. Baris":h.row,"Kolom Gagal":(h.errors||[]).map(p=>p.column||p).join("; "),"Alasan Error":(h.errors||[]).map(p=>p.reason||p).join("; "),...Object.fromEntries(Object.entries(h.data||{}).map(([p,c])=>[p,c??""]))})),l=t.utils.json_to_sheet(s);t.utils.book_append_sheet(n,l,o.sheetName.replace(/[\\\/\[\]*?:]/g,"_").slice(0,31))}),a?(t.writeFile(n,`Log_Error_Import_${new Date().toISOString().slice(0,10)}.xlsx`),!0):!1}var Tr=["validation","employees","contracts","schedule","issues","training","checklist","sop","inspection","cleaning","fogging","basecamp","sp","mutasi"];function di(e){let t=de();if(!t||String(t.role).toLowerCase()!=="superadmin"||t.isRestrictedToken===!0||Number(t.must_change_password)===1){e.innerHTML='<div class="empty-state"><p class="text-danger">Akses ditolak. Halaman ini hanya untuk Superadmin.</p></div>';return}e.innerHTML=`
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
              ${Object.entries(Xt).map(([x,{label:_}])=>`<span class="import-sheet-tag">\u{1F4C4} ${x} \u2192 ${_}</span>`).join("")}
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
  `;let n=null,a=null,o=0,s=!1,l=null,h=null,p={upload:document.getElementById("step-upload"),validating:document.getElementById("step-validating"),preview:document.getElementById("step-preview"),importing:document.getElementById("step-importing"),summary:document.getElementById("step-summary")};function c(x){Object.entries(p).forEach(([_,B])=>{B.style.display=_===x?"":"none"})}document.getElementById("btn-backup-db")?.addEventListener("click",async()=>{let x=document.getElementById("btn-backup-db");x.disabled=!0,x.textContent="\u23F3 Memproses Backup...";try{let _=await X("/api/import/backup");if(_.ok){if(!window.XLSX){z("Library SheetJS belum termuat. Refresh halaman dan coba lagi.");return}let B=window.XLSX,P=B.utils.book_new();Object.entries(_.data.database).forEach(([f,g])=>{let D=g.length>0?g:[{}],O=B.utils.json_to_sheet(D);B.utils.book_append_sheet(P,O,f.substring(0,31))}),B.writeFile(P,`FCMS_Database_Backup_${new Date().toISOString().slice(0,10)}.xlsx`),ae("Backup berhasil diunduh!")}else z("Gagal memproses backup: "+(_.data?.error||"Unknown error"))}catch(_){z("Gagal memproses backup: "+_.message)}finally{x.disabled=!1,x.textContent="\u{1F4E6} Backup Database"}});let E=document.getElementById("btn-sync-google");E&&E.addEventListener("click",async()=>{if(!confirm("Peringatan: Mensinkronkan data dengan Google Sheets akan memperbarui dan menambahkan data baru dari Google Sheets ke dalam FCMS. Data yang sudah Anda buat di FCMS TIDAK akan terhapus. Lanjutkan?"))return;let x=E.innerHTML;E.innerHTML='<span class="spinner"></span> Menyinkronkan...',E.disabled=!0;try{let _=await X("/api/sync/google-sheets",{method:"POST"});_.ok?alert("Sinkronisasi Berhasil: "+(_.data?.message||"Data Karyawan & PIC telah diperbarui.")):alert("Gagal Sinkronisasi: "+(_.data?.error||"Unknown error"))}catch{alert("Terjadi kesalahan koneksi.")}finally{E.innerHTML=x,E.disabled=!1}}),document.getElementById("btn-download-template").addEventListener("click",()=>{si(),ae("Template Excel berhasil didownload!")});let C=document.getElementById("file-input"),T=document.getElementById("upload-zone");document.getElementById("btn-browse").addEventListener("click",x=>{x.stopPropagation(),C.click()}),C.addEventListener("change",x=>{x.target.files[0]&&u(x.target.files[0])}),T.addEventListener("dragover",x=>{x.preventDefault(),T.classList.add("drag-over")}),T.addEventListener("dragleave",()=>T.classList.remove("drag-over")),T.addEventListener("drop",x=>{x.preventDefault(),T.classList.remove("drag-over");let _=x.dataTransfer.files[0];_&&_.name.match(/\.xlsx?$/i)?u(_):z("Hanya file .xlsx atau .xls yang didukung.")}),document.getElementById("btn-clear-file").addEventListener("click",()=>{n=null,C.value="",document.getElementById("file-info").style.display="none",T.style.display="",c("upload")});async function u(x){n=x,document.getElementById("file-name-display").textContent=`\u{1F4C4} ${x.name} (${(x.size/1024).toFixed(1)} KB)`,document.getElementById("file-info").style.display="flex",T.style.display="none",await F(x)}async function F(x){c("validating");let _=document.getElementById("validation-status"),B=document.getElementById("validation-bar");try{if(!window.XLSX)throw new Error("Library SheetJS belum termuat. Refresh halaman dan coba lagi.");_.textContent="Membaca file Excel...",B.style.width="20%",await Dt(200);let P=await x.arrayBuffer(),f=window.XLSX.read(P,{type:"array",cellDates:!0});_.textContent=`Memvalidasi ${f.SheetNames.length} sheet...`,B.style.width="50%",await Dt(100),a=oi(f),B.style.width="100%",_.textContent="Validasi selesai!",await Dt(300),R()}catch(P){c("upload"),z("Gagal memproses file: "+P.message),document.getElementById("file-info").style.display="flex",T.style.display="none"}}function R(){c("preview");let x=a.filter(J=>!J.skipped).length,_=a.reduce((J,$)=>J+$.total,0),B=a.reduce((J,$)=>J+$.valid,0),P=a.reduce((J,$)=>J+$.errorCount,0),f=_>0?Math.round(B/_*100):0;document.getElementById("preview-summary-badges").innerHTML=`
      <span class="badge badge-info">${x} sheet</span>
      <span class="badge badge-secondary">${_} baris</span>
      <span class="badge badge-success">${B} valid (${f}%)</span>
      ${P>0?`<span class="badge badge-danger">${P} error</span>`:""}
    `;let g=document.getElementById("preview-table-container");g.innerHTML=`
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
          ${a.map((J,$)=>`
            <tr class="${J.errorCount>0?"row-error":J.skipped?"row-skipped":"row-ok"}">
              <td><strong>${J.sheetName}</strong></td>
              <td>${J.label}</td>
              <td style="text-align:center">${J.total}</td>
              <td style="text-align:center"><span class="badge badge-success">${J.valid}</span></td>
              <td style="text-align:center">${J.errorCount>0?`<span class="badge badge-danger">${J.errorCount}</span>`:'<span class="text-muted">\u2013</span>'}</td>
              <td style="text-align:center">
                ${J.skipped?'<span class="badge badge-neutral">Dilewati</span>':J.errorCount>0&&J.valid===0?'<span class="badge badge-danger">\u274C 0 Valid</span>':J.errorCount>0?'<span class="badge badge-warning">\u26A0\uFE0F Sebagian</span>':J.valid===0?'<span class="badge badge-neutral">Kosong</span>':'<span class="badge badge-success">\u2705 Siap</span>'}
              </td>
              <td style="text-align:center">
                ${J.errorCount>0?`<button class="btn btn-ghost btn-sm btn-detail-error" data-idx="${$}">\u{1F50D} ${J.errorCount} Error</button>`:'<span class="text-muted">\u2013</span>'}
              </td>
            </tr>
          `).join("")}
        </tbody>
      </table>
    `,g.querySelectorAll(".btn-detail-error").forEach(J=>{J.addEventListener("click",()=>{let $=a[Number(J.dataset.idx)];K($)})});let D=document.getElementById("error-detail-section"),O=document.getElementById("error-detail-container");O.innerHTML="",D.style.display="none";let W=document.getElementById("btn-start-import");B===0?(W.disabled=!0,W.innerHTML="\u26A0\uFE0F Tidak Ada Data Valid"):(W.disabled=!1,P>0?(W.innerHTML=`\u{1F680} Import ${B} Data Valid (${P} dilewati)`,W.title="Baris error akan dilewati, baris valid tetap diimport"):W.innerHTML=`\u{1F680} Mulai Import ${B} Data`)}function K(x){let _=document.getElementById("error-detail-section"),B=document.getElementById("error-detail-container");_.style.display="";let P=x.errors.slice(0,100).map(f=>(Array.isArray(f.errors)?f.errors:[]).map(D=>{let O=typeof D=="object";return`
          <tr>
            <td style="text-align:center"><span class="badge badge-danger">Baris ${f.row}</span></td>
            <td><strong>${O?D.column:"\u2014"}</strong></td>
            <td><code style="font-size:.78rem;color:var(--text-secondary)">${O&&D.originalValue!==void 0?D.originalValue||"(kosong)":"\u2014"}</code></td>
            <td class="error-msg">${O?D.reason:D}</td>
            <td style="font-size:.78rem;color:var(--success)">
              ${O&&D.aliases?`Gunakan salah satu nama kolom:<br><em>${D.aliases}</em>`:O&&D.hint?D.hint:""}
            </td>
          </tr>
        `}).join("")).join("");B.innerHTML=`
      <div class="error-sheet-block">
        <div class="error-sheet-title">
          \u{1F4C4} ${x.sheetName} \u2014 ${x.errorCount} baris error dari ${x.total} total
          ${x.errors.length>100?'<span style="font-weight:400">(menampilkan 100 pertama)</span>':""}
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
            <tbody>${P||'<tr><td colspan="5" class="text-muted" style="text-align:center">Tidak ada detail error</td></tr>'}</tbody>
          </table>
        </div>
        ${x.errors.length>100?`
          <div style="padding:10px 20px;font-size:.8rem;color:var(--text-muted)">
            Hanya menampilkan 100 error pertama. Download Log Error untuk melihat semua.
          </div>`:""}
      </div>
    `,_.scrollIntoView({behavior:"smooth",block:"start"})}document.getElementById("btn-back-to-upload").addEventListener("click",()=>{c("upload"),document.getElementById("file-info").style.display="none",T.style.display="",n=null,C.value=""}),document.getElementById("btn-download-log").addEventListener("click",()=>{if(!a)return;li(a)?ae("Log error berhasil didownload."):ae("Tidak ada error untuk didownload.")});let I=document.getElementById("import-recovery-banner");function N(){if(!I)return;let x=Tn(),_=Array.isArray(x?.items)?x.items.filter(P=>P.status!=="ok"):[];if(x&&(x.status==="IN_PROGRESS"||x.status==="INCOMPLETE"||_.length>0)){let P=_.length>0?_:x.items||[],f=P.find(O=>O.error||O.errorCode)||P[0]||{},g=P.some(O=>O.httpStatus===409&&O.errorCode==="IMPORT_IN_PROGRESS")||x.status==="IN_PROGRESS",D=P.every(O=>Array.isArray(O.rows)&&O.rows.length>0);I.style.display="",I.innerHTML=nt({variant:"recover",title:`\u{1F504} Sesi Import Sebelumnya Belum Selesai (${x.fileName||"Excel"})`,message:D?`${P.length} modul belum selesai (${f.error||"Terhenti akibat refresh atau koneksi terputus"}). Klik Coba Lagi / Lanjutkan untuk meneruskan dengan Idempotency-Key yang sama.`:`${P.length} modul belum selesai. Upload ulang file Excel yang sama agar Idempotency-Key yang tersimpan otomatis digunakan kembali.`,idempotencyKey:f.idempotencyKey||"",importJobId:f.importJobId||"",committedChunks:Number(f.committedChunks||0),totalChunks:Number(f.totalChunks||0),errorCode:f.errorCode||x.status||"",fingerprint:f.fingerprint||"batch",canRetry:D,canRecover:!!(D&&g),canDismiss:!0}),I.querySelector('[data-import-action="retry"]')?.addEventListener("click",()=>{s||M(x.onDuplicate||"skip",x,{recoverLocked:!1})}),I.querySelector('[data-import-action="recover"]')?.addEventListener("click",()=>{s||M(x.onDuplicate||"skip",x,{recoverLocked:!0})}),I.querySelector('[data-import-action="dismiss"]')?.addEventListener("click",()=>{_a(),(x.items||[]).forEach(O=>O.fingerprint&&bt(O.fingerprint)),I.style.display="none",I.innerHTML=""});return}let B=Wt().filter(P=>String(P.endpoint||"").startsWith("/api/import/"));if(B.length>0){let P=B[0];I.style.display="",I.innerHTML=nt({variant:"recover",title:`\u{1F504} Sesi Import Modul ${P.label||P.module} Belum Selesai`,message:P.payloadStored?P.errorMessage||"Import terhenti sebelum selesai. Anda dapat melanjutkan dengan Idempotency-Key yang sama.":"Upload ulang file Excel yang sama untuk melanjutkan dengan Idempotency-Key yang sama.",idempotencyKey:P.idempotencyKey,importJobId:P.importJobId,committedChunks:P.committedChunks,totalChunks:P.totalChunks,errorCode:P.errorCode||P.status,fingerprint:P.fingerprint,canRetry:!!P.payloadStored,canRecover:!!(P.payloadStored&&(P.status==="IN_PROGRESS"||P.status==="CONFLICT_IN_PROGRESS")),canDismiss:!0});let f=async g=>{if(!s){s=!0;try{let D=await Tt(P.fingerprint,{recover:g});bt(P.fingerprint),I.style.display="none",ae(D.idempotentReplay?"Replay idempoten berhasil (0 duplikat).":"Sesi import berhasil diselesaikan!")}catch(D){z(D.message||"Gagal melanjutkan import"),N()}finally{s=!1}}};I.querySelector('[data-import-action="retry"]')?.addEventListener("click",()=>f(!1)),I.querySelector('[data-import-action="recover"]')?.addEventListener("click",()=>f(!0)),I.querySelector('[data-import-action="dismiss"]')?.addEventListener("click",()=>{bt(P.fingerprint),N()});return}I.style.display="none",I.innerHTML=""}N();let w=document.getElementById("btn-start-import");w.addEventListener("click",()=>{if(s){z("Proses import sedang berjalan. Mohon tunggu hingga selesai.");return}let x=document.querySelector('input[name="dup-strategy"]:checked')?.value||"skip";M(x)});async function M(x,_=null,{recoverLocked:B=!1,forceReplayCompleted:P=!1}={}){if(s)return;s=!0,w.disabled=!0,c("importing"),I&&(I.style.display="none"),o=Date.now();let f;if(_&&Array.isArray(_.items))f=_.items.map(J=>({...J}));else{let J=[];Tr.forEach($=>{let m=a?.find(i=>i.module===$&&i.mapped?.length>0);m&&J.push(m)}),f=J.map($=>({module:$.module,label:$.label,endpoint:`/api/import/${$.module}`,rows:$.mapped,rowCount:$.mapped.length,status:"pending",inserted:0,updated:0,skipped:0,committedChunks:0,totalChunks:0,idempotencyKey:null,fingerprint:null,importJobId:null,idempotentReplay:!1,httpStatus:null,errorCode:null,error:null,recoverable:!1}))}for(let J of f){let $={rows:J.rows,onDuplicate:x};if(J.fingerprint||(J.fingerprint=await Fa(J.endpoint,$)),!J.idempotencyKey){let m=String(J.module||"imp").slice(0,20);J.idempotencyKey=globalThis?.crypto?.randomUUID?`imp_${m}_${globalThis.crypto.randomUUID()}`:`imp_${m}_${Date.now().toString(36)}_${Math.random().toString(36).slice(2,8)}`}}let g={batchId:_?.batchId||(globalThis?.crypto?.randomUUID?globalThis.crypto.randomUUID():`batch_${Date.now()}`),fileName:_?.fileName||n?.name||"Excel Import",onDuplicate:x,status:"IN_PROGRESS",items:f};l=g,Bt(g);let D=document.getElementById("import-steps-list");D.innerHTML=f.map(J=>`
      <div class="import-step-item" id="step-item-${J.module}">
        <span class="step-item-icon" id="step-icon-${J.module}">${J.status==="ok"&&!P?"\u2705":"\u23F8\uFE0F"}</span>
        <span class="step-item-label">${Ce(J.label)} <span class="step-item-count">(${J.rowCount||J.rows?.length||0} data)</span></span>
        <span class="step-item-status" id="step-status-${J.module}">${J.status==="ok"&&!P?`<span class="badge badge-success">${J.inserted||0} berhasil</span>`:""}</span>
      </div>
    `).join("");let O=document.getElementById("import-bar"),W=document.getElementById("import-current-status");try{for(let m=0;m<f.length;m++){let i=f[m],r=document.getElementById(`step-icon-${i.module}`),d=document.getElementById(`step-status-${i.module}`);if(i.status==="ok"&&!P){O.style.width=`${Math.round((m+1)/f.length*100)}%`;continue}r.textContent="\u{1F504}",d.textContent=B&&i.httpStatus===409?"Memulihkan lock...":"Mengimport...",W.textContent=`Mengimport ${i.label}... (${m+1}/${f.length})`,O.style.width=`${Math.round(m/Math.max(1,f.length)*100)}%`,Bt(g);try{let b=!!(B&&(i.httpStatus===409||i.errorCode==="IMPORT_IN_PROGRESS"||i.status==="pending"||i.status==="error")),y=await Le({module:i.module,label:i.label,endpoint:i.endpoint,body:{rows:i.rows,onDuplicate:x},idempotencyKey:i.idempotencyKey,recover:b}),q=y.data||{};i.status="ok",i.inserted=Number(q.inserted||0),i.updated=Number(q.updated||0),i.skipped=Number(q.skipped||0),i.idempotentReplay=!!(y.idempotentReplay||q.idempotentReplay),i.importJobId=y.importJobId||q.importJobId||i.importJobId||null,i.idempotencyKey=y.idempotencyKey||i.idempotencyKey,i.httpStatus=y.status,i.error=null,i.errorCode=null,i.recoverable=!1,r.textContent=i.idempotentReplay?"\u267B\uFE0F":"\u2705";let v=i.idempotentReplay?' <span class="badge badge-info">Replay Idempoten</span>':"",S=i.skipped>0?` <span class="badge badge-neutral">${i.skipped} skip</span>`:"",A=i.updated>0?` <span class="badge badge-info">${i.updated} update</span>`:"";d.innerHTML=`<span class="badge badge-success">${i.inserted} berhasil</span>${A}${S}${v}`}catch(b){i.status="error",i.error=b.message||"Gagal mengimport data",i.errorCode=b.code||"IMPORT_ERROR",i.httpStatus=Number(b.status??0),i.committedChunks=Number(b.committedChunks||0),i.totalChunks=Number(b.totalChunks||0),i.importJobId=b.importJobId||i.importJobId||null,i.idempotencyKey=b.idempotencyKey||i.idempotencyKey||null,i.recoverable=!!b.recoverable,r.textContent=i.httpStatus===409?"\u23F3":"\u274C";let y=i.totalChunks>0?` <span class="badge badge-warning">Chunk ${i.committedChunks}/${i.totalChunks}</span>`:"";d.innerHTML=`<span class="badge badge-danger">${Ce(i.errorCode||"Gagal")}</span>${y}`}g.status=f.every(b=>b.status==="ok")?"COMPLETED":"INCOMPLETE",Bt(g),await Dt(80)}let J=f.every(m=>m.status==="ok");J?_a():(g.status="INCOMPLETE",Bt(g));let $={totalSheets:f.length,totalRows:f.reduce((m,i)=>m+(i.rowCount||i.rows?.length||0),0),inserted:f.reduce((m,i)=>m+(i.inserted||0),0),updated:f.reduce((m,i)=>m+(i.updated||0),0),skipped:f.reduce((m,i)=>m+(i.skipped||0),0),failed:f.filter(m=>m.status!=="ok").length,hasReplays:f.some(m=>m.idempotentReplay),moduleResults:f,batchState:g};h=$,O.style.width="100%",W.textContent=J?"Selesai!":"Selesai dengan modul gagal/terhenti (dapat dilanjutkan)",await Dt(250),k($)}finally{s=!1,w.disabled=!1}}function k(x){c("summary");let _=((Date.now()-o)/1e3).toFixed(1),B=x.failed===0&&x.moduleResults.every(D=>D.status==="ok"),P=x.moduleResults.some(D=>D.status!=="ok"&&D.httpStatus===409&&D.errorCode==="IMPORT_IN_PROGRESS"),f=x.moduleResults.find(D=>D.status!=="ok");document.getElementById("summary-status-icon").innerHTML=`
      <div class="summary-icon">${B?x.hasReplays?"\u267B\uFE0F":"\u{1F389}":"\u26A0\uFE0F"}</div>
      <h2 class="summary-title">${B?x.hasReplays?"Import & Replay Idempoten Berhasil!":"Import Berhasil!":"Import Belum Selesai (Dapat Dilanjutkan)"}</h2>
    `;let g=document.getElementById("summary-action-banner");if(g)if(!B&&f)g.innerHTML=nt({variant:P?"warning":"error",title:P?"\u23F3 Ada Modul Terkunci / Sedang Berjalan (HTTP 409)":"\u26A0\uFE0F Beberapa Modul Gagal atau Terhenti di Tengah Proses",message:f.error||"Klik Coba Lagi untuk melanjutkan dengan Idempotency-Key yang sama tanpa menduplikasi data.",idempotencyKey:f.idempotencyKey||"",importJobId:f.importJobId||"",committedChunks:f.committedChunks||0,totalChunks:f.totalChunks||0,errorCode:f.errorCode||"",fingerprint:f.fingerprint||"",canRetry:!0,canRecover:P,canDismiss:!1}),g.querySelector('[data-import-action="retry"]')?.addEventListener("click",()=>{M(x.batchState?.onDuplicate||"skip",x.batchState,{recoverLocked:!1})}),g.querySelector('[data-import-action="recover"]')?.addEventListener("click",()=>{M(x.batchState?.onDuplicate||"skip",x.batchState,{recoverLocked:!0})});else if(B&&x.moduleResults.length>0){let D=x.moduleResults[0];g.innerHTML=nt({variant:"success",title:x.hasReplays?"\u267B\uFE0F Replay Idempoten Dikonfirmasi (0 Duplikat)":"\u2705 Semua Modul Terkonfirmasi Server",message:x.hasReplays?"Server mengembalikan hasil tersimpan menggunakan Idempotency-Key yang sama tanpa menulis baris ganda.":"Semua chunk telah committed di server beserta Audit Log terkait.",idempotencyKey:D.idempotencyKey||"",importJobId:D.importJobId||"",canReplay:!0,canDismiss:!1}),g.querySelector('[data-import-action="replay"]')?.addEventListener("click",()=>{M(x.batchState?.onDuplicate||"skip",x.batchState,{recoverLocked:!1,forceReplayCompleted:!0})})}else g.innerHTML="";document.getElementById("summary-stats").innerHTML=`
      <div class="summary-stat-card">
        <div class="stat-value">${x.totalSheets}</div>
        <div class="stat-label">Total Sheet</div>
      </div>
      <div class="summary-stat-card">
        <div class="stat-value">${x.totalRows}</div>
        <div class="stat-label">Total Data</div>
      </div>
      <div class="summary-stat-card success">
        <div class="stat-value">${x.inserted}</div>
        <div class="stat-label">Berhasil Diimport</div>
      </div>
      <div class="summary-stat-card neutral">
        <div class="stat-value">${x.skipped}</div>
        <div class="stat-label">Dilewati (Duplikat)</div>
      </div>
      ${x.updated>0?`<div class="summary-stat-card info"><div class="stat-value">${x.updated}</div><div class="stat-label">Diperbarui</div></div>`:""}
      ${x.failed>0?`<div class="summary-stat-card danger"><div class="stat-value">${x.failed}</div><div class="stat-label">Modul Gagal</div></div>`:""}
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
          ${x.moduleResults.map((D,O)=>`
            <tr>
              <td>
                <div><strong>${Ce(D.label)}</strong></div>
                ${D.idempotencyKey?`<div style="font-size:.72rem;color:var(--text-muted)">Key: <code>${Ce(D.idempotencyKey.slice(0,26))}\u2026</code>${D.importJobId?` | Job: <code>${Ce(D.importJobId.slice(0,10))}\u2026</code>`:""}</div>`:""}
              </td>
              <td style="text-align:center"><span class="badge badge-success">${D.inserted||0}</span>${D.updated>0?` <span class="badge badge-info">+${D.updated} upd</span>`:""}</td>
              <td style="text-align:center"><span class="badge badge-neutral">${D.skipped||0}</span></td>
              <td style="text-align:center">
                ${D.status==="ok"?`<span class="badge badge-success">${D.idempotentReplay?"\u267B\uFE0F Replay Idempoten":"\u2705 Sukses"}</span>`:`<div>
                      <span class="badge badge-danger" title="${Ce(D.error||"")}">\u274C ${Ce(D.errorCode||"Gagal")}${D.httpStatus?` (${D.httpStatus})`:""}</span>
                      ${D.totalChunks>0?`<div style="font-size:.72rem;margin-top:4px;">Chunk ${D.committedChunks}/${D.totalChunks}</div>`:""}
                      <div style="font-size:.75rem;color:var(--danger);margin-top:2px;">${Ce(D.error||"")}</div>
                      <button type="button" class="btn btn-primary btn-sm" style="margin-top:6px" data-row-retry="${O}">\u{1F504} Coba Lagi</button>
                      ${D.httpStatus===409&&D.errorCode==="IMPORT_IN_PROGRESS"?`<button type="button" class="btn btn-warning btn-sm" style="margin-top:6px" data-row-recover="${O}">\u{1F6E0}\uFE0F Pulihkan</button>`:""}
                    </div>`}
              </td>
            </tr>
          `).join("")}
        </tbody>
      </table>
    `,document.querySelectorAll("[data-row-retry]").forEach(D=>{D.addEventListener("click",()=>{M(x.batchState?.onDuplicate||"skip",x.batchState,{recoverLocked:!1})})}),document.querySelectorAll("[data-row-recover]").forEach(D=>{D.addEventListener("click",()=>{M(x.batchState?.onDuplicate||"skip",x.batchState,{recoverLocked:!0})})})}document.getElementById("btn-import-again").addEventListener("click",()=>{n=null,a=null,l=null,h=null,C.value="",document.getElementById("file-info").style.display="none",T.style.display="",N(),c("upload")}),document.getElementById("btn-go-to-dashboard").addEventListener("click",()=>{window.location.hash="/dashboard"})}function Dt(e){return new Promise(t=>setTimeout(t,e))}Fe();function It(e){if(!e)return{status:"Tidak Ditentukan",label:"Tidak Ada Tanggal",bg:"#F1F5F9",color:"#64748B",border:"#CBD5E1",icon:"\u26AA",daysLeft:0,isExpired:!1,isExpiringSoon:!1};let t=new Date;t.setHours(0,0,0,0);let n=new Date(e);n.setHours(0,0,0,0);let a=n.getTime()-t.getTime(),o=Math.ceil(a/(1e3*60*60*24));return o<0?{status:"Selesai",label:`Masa Sanksi Selesai (${Math.abs(o)} hari lalu)`,shortLabel:"Selesai",bg:"#DCFCE7",color:"#15803D",border:"#BBF7D0",icon:"\u2705",daysLeft:o,isExpired:!0,isExpiringSoon:!1}:o<=30?{status:"Mendekati Berakhir",label:`Segera Berakhir (Sisa ${o} hari)`,shortLabel:`Sisa ${o} hari`,bg:"#FEF3C7",color:"#B45309",border:"#FDE68A",icon:"\u26A0\uFE0F",daysLeft:o,isExpired:!1,isExpiringSoon:!0}:{status:"Aktif",label:`Aktif Berlaku (Sisa ${o} hari)`,shortLabel:`Aktif (${o} hari)`,bg:"#FEE2E2",color:"#B91C1C",border:"#FECACA",icon:"\u{1F534}",daysLeft:o,isExpired:!1,isExpiringSoon:!1}}function Na(e=""){let t=(e||"").toUpperCase().trim();return t==="SP 3"?{label:"SP 3 (Terakhir)",bg:"#7F1D1D",color:"#FFFFFF",border:"#991B1B",icon:"\u{1F6A8}"}:t==="SP 2"?{label:"SP 2",bg:"#FEE2E2",color:"#B91C1C",border:"#FECACA",icon:"\u{1F6D1}"}:t==="SP 1"?{label:"SP 1",bg:"#FEF3C7",color:"#B45309",border:"#FDE68A",icon:"\u26A0\uFE0F"}:t.includes("TEGURAN")?{label:"Teguran Lisan",bg:"#E0F2FE",color:"#0369A1",border:"#BAE6FD",icon:"\u{1F4AC}"}:{label:e||"SP",bg:"#F1F5F9",color:"#475569",border:"#CBD5E1",icon:"\u{1F4DC}"}}function ft(e){if(!e)return"-";try{let t=new Date(e);return isNaN(t.getTime())?e:t.toLocaleDateString("id-ID",{day:"numeric",month:"short",year:"numeric"})}catch{return e}}async function pi(e){let t=de(),n=ie(t,"sp","write"),a=ie(t,"sp","write"),o=ie(t,"sp","delete"),s=[],l=[],h=[],p="all",c="all",E="all",C="",T=localStorage.getItem("fcms_sp_view_mode")||"card",u=1,F=20,R=0,K=1,I={total:0,active:0,expiring:0,expired:0,sp1:0,sp2:0,sp3:0,teguran:0,uniqueEmployees:0},N=null,w=!1,M="fcms-sp-page-styles",k=document.getElementById(M);k||(k=document.createElement("style"),k.id=M,document.head.appendChild(k)),k.textContent=`
      .sp-page-container {
        padding: 24px;
        max-width: 1400px;
        margin: 0 auto;
        font-family: inherit;
      }
      .sp-hero-card {
        background: linear-gradient(135deg, #1E293B 0%, #0F172A 100%);
        border-radius: 18px;
        padding: 28px 32px;
        color: #FFFFFF;
        box-shadow: 0 10px 25px -5px rgba(15, 23, 42, 0.2);
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
        background: radial-gradient(circle, rgba(239, 68, 68, 0.15) 0%, transparent 70%);
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
        background: #172033;
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
        background: #1B2940;
        transition: all 0.2s ease;
      }
      .sp-search-input:focus {
        outline: none;
        border-color: #2563EB;
        background: #172033;
        box-shadow: 0 0 0 3px rgba(37, 99, 235, 0.12);
      }
      .sp-filter-select {
        padding: 9px 14px;
        border: 1.5px solid #CBD5E1;
        border-radius: 10px;
        font-size: 0.88rem;
        background: #1B2940;
        color: #334155;
        font-weight: 500;
        cursor: pointer;
      }
      .sp-filter-select:focus {
        outline: none;
        border-color: #2563EB;
        background: #172033;
      }
      .sp-filter-pills {
        display: flex;
        flex-wrap: wrap;
        gap: 8px;
        align-items: center;
      }
      .sp-pill {
        border: 1px solid #29364B;
        background: #172033;
        color: #94A3B8;
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
        background: #EEF2F6;
        border-color: #CBD5E1;
      }
      .sp-pill.active {
        background: #2563EB;
        border-color: #2563EB;
        color: #FFFFFF;
        box-shadow: 0 2px 6px rgba(37, 99, 235, 0.25);
      }
      .sp-pill-count {
        background: rgba(0, 0, 0, 0.08);
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
        background: #111827;
        padding: 3px;
        border-radius: 10px;
        border: 1px solid #29364B;
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
        color: #64748B;
        cursor: pointer;
        transition: all 0.15s ease;
        height: 34px;
        box-sizing: border-box;
        user-select: none;
      }
      .sp-view-btn:hover {
        color: #F8FAFC;
        background: rgba(255, 255, 255, 0.5);
      }
      .sp-view-btn.active {
        background: #172033;
        color: #2563EB;
        font-weight: 700;
        box-shadow: 0 1px 4px rgba(0, 0, 0, 0.08);
      }
      /* Modern Table View */
      .sp-table-card {
        background: #172033;
        border: 1px solid #E2E8F0;
        border-radius: 16px;
        box-shadow: 0 2px 8px rgba(0, 0, 0, 0.04);
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
        background: #111827;
        padding: 13px 16px;
        font-size: 0.76rem;
        font-weight: 700;
        color: #475569;
        text-transform: uppercase;
        letter-spacing: 0.05em;
        border-bottom: 1.5px solid #E2E8F0;
        white-space: nowrap;
      }
      .sp-table td {
        padding: 13px 16px;
        font-size: 0.88rem;
        color: #F8FAFC;
        border-bottom: 1px solid #F1F5F9;
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
        background: #172033;
        border: 1px solid #E2E8F0;
        border-radius: 16px;
        padding: 22px;
        box-shadow: 0 2px 6px rgba(0, 0, 0, 0.04);
        display: flex;
        flex-direction: column;
        justify-content: space-between;
        transition: all 0.25s cubic-bezier(0.16, 1, 0.3, 1);
        position: relative;
      }
      .sp-card:hover {
        transform: translateY(-4px);
        box-shadow: 0 12px 24px -6px rgba(15, 23, 42, 0.1);
        border-color: #CBD5E1;
      }
      .sp-status-badge {
        display: inline-flex;
        align-items: center;
        gap: 6px;
        padding: 4px 10px;
        border-radius: 20px;
        font-size: 0.74rem;
        font-weight: 700;
        letter-spacing: 0.02em;
      }
      .sp-type-badge {
        display: inline-flex;
        align-items: center;
        gap: 4px;
        padding: 4px 9px;
        border-radius: 6px;
        font-size: 0.74rem;
        font-weight: 700;
      }
      .sp-avatar-box {
        width: 42px;
        height: 42px;
        border-radius: 10px;
        background: #EFF6FF;
        color: #2563EB;
        font-weight: 700;
        display: flex;
        align-items: center;
        justify-content: center;
        font-size: 0.95rem;
        flex-shrink: 0;
      }
      .sp-icon-btn {
        background: #1B2940;
        border: 1px solid #29364B;
        color: #64748B;
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
        background: #EDF2F7;
        color: #F8FAFC;
        border-color: #CBD5E1;
      }
      .sp-icon-btn.delete:hover {
        background: #FEE2E2;
        color: #DC2626;
        border-color: #FECACA;
      }
      .sp-empty-state {
        text-align: center;
        padding: 60px 20px;
        background: #172033;
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
          font-size: 0.68rem !important;
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
    `;function x(){let b={sort:"date_desc"};return C&&C.trim()&&(b.search=C.trim()),c&&c!=="all"&&(b.branch_id=c),E&&E!=="all"&&(b.division=E),p&&p!=="all"&&(b.filter=p),b}function _(){let b=(y,q)=>{let v=e.querySelector(y);v&&(v.textContent=String(q??0))};b("#sp-stat-active",I.active),b("#sp-stat-expiring",I.expiring),b("#sp-stat-expired",I.expired),b("#sp-stat-unique",I.uniqueEmployees),b("#sp-pill-count-all",I.total),b("#sp-pill-count-active",I.active),b("#sp-pill-count-expiring",I.expiring),b("#sp-pill-count-expired",I.expired),b("#sp-pill-count-sp1",I.sp1),b("#sp-pill-count-sp2",I.sp2),b("#sp-pill-count-sp3",I.sp3)}async function B({refreshStats:b=!0}={}){try{if(l.length===0||h.length===0){let[S,A]=await Promise.all([ve(),Ie()]);l=S||[],h=A||[]}let y=x(),q=new URLSearchParams({page:String(u),limit:String(F),...y});b&&q.set("include_stats","1");let v=await X("/api/sp?"+q.toString());if(v.ok&&v.data?.data&&Array.isArray(v.data.data)){let S=v.data.pagination||{};if(R=Number(S.total??v.data.data.length),K=Math.max(1,Number(S.pages||S.totalPages||Math.ceil(R/F)||1)),v.data.data.length===0&&u>K)return u=K,B({refreshStats:b});s=v.data.data,v.data.stats&&(I={...I,...v.data.stats})}else s=[],R=0,K=1,v.ok||z(v.data?.error||"Gagal memuat data SP dari server")}catch(y){console.warn("Error loading SP data:",y),s=[],R=0,K=1,z("Gagal memuat data SP: "+y.message)}w?(b&&_(),i()):(w=!0,r())}function P(){return s}function f(){return I}function g(b=null){let y=!!b;if(y&&!a){z("Anda tidak memiliki izin untuk mengedit data SP.");return}if(!y&&!n){z("Anda tidak memiliki izin untuk menambah data SP.");return}let q=!!b,v=b?.tanggal?b.tanggal.slice(0,10):new Date().toISOString().slice(0,10),S=b?.akhir_sp?b.akhir_sp.slice(0,10):"";if(!S){let Q=new Date(v);Q.setMonth(Q.getMonth()+6),S=Q.toISOString().slice(0,10)}let A=document.createElement("div");A.innerHTML=`
      <div style="display:flex; flex-direction:column; gap:14px;">
        <div class="sp-form-grid" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Nama Karyawan <span style="color:#DC2626;">*</span>
            </label>
            <div id="sp-modal-emp-mount"></div>
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Divisi <span style="color:#DC2626;">*</span>
            </label>
            <select id="sp-modal-div" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" required>
              <option value="FACILITY CARE" ${b?.division==="FACILITY CARE"?"selected":""}>FACILITY CARE</option>
              <option value="SECURITY" ${b?.division==="SECURITY"?"selected":""}>SECURITY</option>
            </select>
          </div>
        </div>

        <div class="sp-form-grid" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Cabang Penempatan <span style="color:#DC2626;">*</span>
            </label>
            <select id="sp-modal-branch" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" required>
              <option value="">Pilih Cabang...</option>
              ${l.map(Q=>`<option value="${Q.value}" ${String(b?.branch_id)===String(Q.value)?"selected":""}>${Q.label}</option>`).join("")}
            </select>
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Tingkat / Jenis SP <span style="color:#DC2626;">*</span>
            </label>
            <select id="sp-modal-type" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" required>
              <option value="SP 1" ${b?.sp_type==="SP 1"?"selected":""}>SP 1 (Masa Berlaku 6 Bulan)</option>
              <option value="SP 2" ${b?.sp_type==="SP 2"?"selected":""}>SP 2 (Masa Berlaku 6 Bulan)</option>
              <option value="SP 3" ${b?.sp_type==="SP 3"?"selected":""}>SP 3 (Masa Berlaku 6 Bulan / Terakhir)</option>
              <option value="Teguran Lisan" ${b?.sp_type==="Teguran Lisan"?"selected":""}>Teguran Lisan (Pembinaan)</option>
            </select>
          </div>
        </div>

        <div class="sp-form-grid" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Tanggal Terbit SP <span style="color:#DC2626;">*</span>
            </label>
            <input type="date" id="sp-modal-tanggal" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              value="${v}" required />
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Tanggal Akhir Masa Berlaku SP <span style="color:#DC2626;">*</span>
            </label>
            <input type="date" id="sp-modal-akhir" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              value="${S}" required />
          </div>
        </div>

        <div>
          <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
            Link Berkas / Foto BAP / Scan Surat SP
          </label>
          <input type="url" id="sp-modal-link" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
            placeholder="https://drive.google.com/..." value="${b?.document_link?String(b.document_link).replace(/"/g,"&quot;"):""}" />
        </div>

        <div>
          <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
            Kronologi Pelanggaran / Catatan Pembinaan
          </label>
          <textarea id="sp-modal-notes" class="form-control" rows="3" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
            placeholder="Tuliskan ringkasan pelanggaran disiplin kerja yang terjadi...">${b?.notes||b?.description||""}</textarea>
        </div>
      </div>
    `;let U=A.querySelector("#sp-modal-tanggal"),H=A.querySelector("#sp-modal-akhir");U.addEventListener("change",()=>{if(U.value){let Q=new Date(U.value);Q.setMonth(Q.getMonth()+6),H.value=Q.toISOString().slice(0,10)}});let ee=ze(A.querySelector("#sp-modal-emp-mount"),{mode:"single",inputId:"sp-modal-emp",inputName:"employee_name",placeholder:"Cari nama karyawan...",required:!0,allowManual:!0,hasEmployeeIdSchema:!0,manualBadgeText:"Belum terhubung ke Master Karyawan",employees:h,initialEmployeeId:b?.employee_id||null,initialEmployeeName:b?.employee_name||""});pe({title:q?"\u270F\uFE0F Edit Data Surat Peringatan (SP)":"\u2795 Terbitkan Surat Peringatan (SP) Baru",content:A,size:"md",confirmText:q?"Simpan Perubahan":"Terbitkan SP",confirmClass:"btn-primary",onConfirm:async(Q,ne)=>{let Z=ee.getValue(),G=(Z.employee_name||"").trim(),L=Z.employee_id??null,j=A.querySelector("#sp-modal-div").value,V=A.querySelector("#sp-modal-branch").value,te=A.querySelector("#sp-modal-type").value,re=A.querySelector("#sp-modal-tanggal").value,le=A.querySelector("#sp-modal-akhir").value,ce=(A.querySelector("#sp-modal-link").value||"").trim(),se=(A.querySelector("#sp-modal-notes").value||"").trim();if(!G){z("Nama Karyawan wajib diisi"),ee.inputEl?.focus();return}if(!V){z("Cabang penempatan wajib dipilih");return}if(!re||!le){z("Tanggal SP dan Akhir SP wajib diisi");return}let he=l.find(oe=>String(oe.value)===String(V)),Ke=he?he.label:"Cabang",$e={employee_name:G,employee_id:L,division:j,branch_id:Number(V),branch_name:Ke,sp_type:te,tanggal:re,akhir_sp:le,document_link:ce,notes:se,status:It(le).status,...q?{row_version:b?.row_version!==void 0&&b?.row_version!==null?Number(b.row_version):1}:{}};try{if(q){let oe=await X(`/api/sp/${b.id}`,{method:"PUT",body:JSON.stringify($e)});if(!oe.ok){if(oe.status===409){z("Konflik OCC: Data telah diubah oleh pengguna lain. Silakan muat ulang halaman."),ne(),await B({refreshStats:!0});return}z(oe.data?.error||"Gagal memperbarui data SP");return}ae("Data SP berhasil diperbarui")}else{let oe=await X("/api/sp",{method:"POST",body:JSON.stringify($e)});if(!oe.ok){z(oe.data?.error||"Gagal menerbitkan data SP");return}u=1,ae("Surat Peringatan baru berhasil diterbitkan")}ne(),await B({refreshStats:!0})}catch(oe){console.error("Error saving SP:",oe),z("Gagal menyimpan SP: "+oe.message)}}})}function D(b){if(!o){z("Anda tidak memiliki izin untuk menghapus data SP.");return}me(`Hapus data Surat Peringatan untuk karyawan <strong>${b.employee_name}</strong> (${b.sp_type})?`,async()=>{try{let y=await X(`/api/sp/${b.id}`,{method:"DELETE"});if(!y.ok){z(y.data?.error||"Gagal menghapus data SP dari server");return}ae(`Data SP ${b.employee_name} berhasil dihapus`),await B({refreshStats:!0})}catch(y){z("Gagal menghapus data SP: "+y.message)}},"Hapus Data SP")}function O(b){let y=It(b.akhir_sp),q=Na(b.sp_type),v=document.createElement("div");v.innerHTML=`
      <div style="display:flex; flex-direction:column; gap:16px;">
        <div style="display:flex; align-items:center; gap:14px; padding-bottom:14px; border-bottom:1px solid #E2E8F0;">
          <div class="sp-avatar-box" style="width:54px; height:54px; font-size:1.2rem;">
            ${b.employee_name.slice(0,2).toUpperCase()}
          </div>
          <div>
            <h3 style="margin:0 0 4px 0; font-size:1.15rem; font-weight:700; color:#0F172A;">
              ${b.employee_name}
            </h3>
            <div style="display:flex; align-items:center; gap:8px;">
              <span class="badge" style="background:#EFF6FF; color:#1D4ED8; font-weight:700;">${b.division||"FACILITY CARE"}</span>
              <span style="color:#64748B; font-size:0.85rem;">\u{1F4CD} ${b.branch_name||"Cabang"}</span>
            </div>
          </div>
        </div>

        <div class="sp-detail-grid" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div style="background:#111827; border:1px solid #29364B; border-radius:10px; padding:12px;">
            <div style="font-size:0.75rem; color:#64748B; text-transform:uppercase; font-weight:600; margin-bottom:4px;">Tingkat Sanksi</div>
            <span class="sp-type-badge" style="background:${q.bg}; color:${q.color}; border:1px solid ${q.border}; font-size:0.85rem;">
              <span>${q.icon}</span>
              <span>${q.label}</span>
            </span>
          </div>

          <div style="background:#111827; border:1px solid #29364B; border-radius:10px; padding:12px;">
            <div style="font-size:0.75rem; color:#64748B; text-transform:uppercase; font-weight:600; margin-bottom:4px;">Status Masa Berlaku</div>
            <span class="sp-status-badge" style="background:${y.bg}; color:${y.color}; border:1px solid ${y.border}; font-size:0.85rem;">
              <span>${y.icon}</span>
              <span>${y.label}</span>
            </span>
          </div>
        </div>

        <div class="sp-detail-grid" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div style="background:#111827; border:1px solid #29364B; border-radius:10px; padding:12px;">
            <div style="font-size:0.75rem; color:#64748B; text-transform:uppercase; font-weight:600; margin-bottom:2px;">Tanggal Terbit SP</div>
            <div style="font-size:0.95rem; font-weight:700; color:#1E293B;">${ft(b.tanggal)}</div>
          </div>

          <div style="background:#111827; border:1px solid #29364B; border-radius:10px; padding:12px;">
            <div style="font-size:0.75rem; color:#64748B; text-transform:uppercase; font-weight:600; margin-bottom:2px;">Akhir Masa Berlaku</div>
            <div style="font-size:0.95rem; font-weight:700; color:#1E293B;">${ft(b.akhir_sp)}</div>
          </div>
        </div>

        ${b.notes||b.description?`
          <div style="background:rgba(245,158,11,0.15); border:1px solid rgba(245,158,11,0.3); border-radius:10px; padding:14px;">
            <div style="font-size:0.78rem; font-weight:700; color:#92400E; margin-bottom:4px;">\u{1F4DD} Catatan Pelanggaran / Pembinaan:</div>
            <div style="font-size:0.88rem; color:#78350F; line-height:1.5;">${b.notes||b.description}</div>
          </div>
        `:""}

        ${b.document_link?`
          <div>
            <a href="${b.document_link}" target="_blank" rel="noopener noreferrer" class="btn btn-primary" style="width:100%; justify-content:center; padding:10px; font-weight:600; display:flex; align-items:center; gap:8px;">
              <span>\u{1F4C4}</span> Buka Dokumen BAP / Scan SP Fisik
            </a>
          </div>
        `:`
          <div style="text-align:center; padding:8px; font-size:0.82rem; color:#94A3B8;">
            (Belum ada scan dokumen / BAP yang diunggah)
          </div>
        `}
      </div>
    `,pe({title:"\u{1F4CB} Lembar Riwayat Surat Peringatan (SP)",content:v,size:"md",cancelText:"Tutup"})}async function W(){if(!ie(t,"export")){z("Anda tidak memiliki izin untuk mengekspor data.");return}try{let{downloadExcel:b}=await Promise.resolve().then(()=>(we(),Oe)),y=await ue("/api/sp/export",x()),q=y.map((v,S)=>{let A=It(v.akhir_sp);return{No:S+1,"Nama Karyawan":v.employee_name||"",Divisi:v.division||"",Cabang:v.branch_name||"","Jenis SP":v.sp_type||"","Tanggal SP":v.tanggal||"","Akhir SP":v.akhir_sp||"","Status Masa Berlaku":A.status,"Sisa Hari":A.daysLeft>0?A.daysLeft:0,"Link BAP / Foto":v.document_link||"","Catatan / Kronologi":v.notes||v.description||""}});b(q,`Data_Surat_Peringatan_FCMS_${new Date().toISOString().slice(0,10)}`),ae(`Berhasil mengunduh ${y.length} data SP ke format Excel.`)}catch(b){z("Gagal ekspor: "+b.message)}}function J(b){return`
          ${b.length===0?`
            <div class="sp-empty-state">
              <div style="font-size:3rem; margin-bottom:12px;">\u{1F50D}</div>
              <h3 style="font-size:1.2rem; font-weight:700; color:#1E293B; margin:0 0 6px 0;">Tidak Ada Data SP Ditemukan</h3>
              <p style="color:#64748B; font-size:0.9rem; margin:0 0 20px 0;">Silakan sesuaikan kata kunci pencarian atau reset filter aktif.</p>
              <button id="sp-btn-reset-filter" class="btn btn-ghost" style="border:1px solid #CBD5E1; border-radius:8px; padding:8px 16px;">
                Reset Filter
              </button>
            </div>
          `:T==="table"?`
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
                    ${b.map((y,q)=>{let v=It(y.akhir_sp),S=Na(y.sp_type);return`
                        <tr data-id="${y.id}">
                          <td style="text-align:center; color:#94A3B8; font-weight:600; font-size:0.8rem;">${(u-1)*F+q+1}</td>
                          <td>
                            <div style="display:flex; align-items:center; gap:10px;">
                              <div class="sp-avatar-box" style="width:34px; height:34px; font-size:0.78rem; border-radius:8px;">
                                ${(y.employee_name||"SP").slice(0,2).toUpperCase()}
                              </div>
                              <div style="font-weight:700; color:#0F172A; white-space:nowrap;">
                                ${y.employee_name}
                              </div>
                            </div>
                          </td>
                          <td style="white-space:nowrap;">
                            <div style="font-weight:600; color:#334155; font-size:0.84rem;">\u{1F4CD} ${y.branch_name||"Cabang"}</div>
                            <div style="font-size:0.74rem; color:#2563EB; font-weight:600; margin-top:2px;">${y.division||"FACILITY CARE"}</div>
                          </td>
                          <td style="text-align:center; white-space:nowrap;">
                            <span class="sp-type-badge" style="background:${S.bg}; color:${S.color}; border:1px solid ${S.border};">
                              <span>${S.icon}</span>
                              <span>${S.label}</span>
                            </span>
                          </td>
                          <td style="white-space:nowrap; font-weight:600; color:#1E293B;">
                            ${ft(y.tanggal)}
                          </td>
                          <td style="white-space:nowrap; font-weight:700; color:#0F172A;">
                            ${ft(y.akhir_sp)}
                          </td>
                          <td style="text-align:center; white-space:nowrap;">
                            <span class="sp-status-badge" style="background:${v.bg}; color:${v.color}; border:1px solid ${v.border};">
                              <span>${v.icon}</span>
                              <span>${v.shortLabel}</span>
                            </span>
                          </td>
                          <td style="max-width:240px;">
                            <div style="font-size:0.82rem; color:#64748B; white-space:nowrap; overflow:hidden; text-overflow:ellipsis; margin-bottom:4px;" title="${String(y.notes||y.description||"").replace(/"/g,"&quot;")}">
                              ${y.notes||y.description||"Pencatatan sanksi peringatan disiplin operasional."}
                            </div>
                            <button class="btn btn-sm btn-ghost sp-detail-btn" style="border:1px solid #CBD5E1; border-radius:6px; font-weight:600; font-size:0.74rem; padding:2px 8px;" data-id="${y.id}">
                              \u{1F441}\uFE0F Rincian BAP
                            </button>
                          </td>
                          <td style="text-align:center; white-space:nowrap;">
                            <div style="display:inline-flex; align-items:center; gap:6px;">
                              ${y.document_link?`
                                <a href="${y.document_link}" target="_blank" rel="noopener noreferrer" class="sp-icon-btn" title="Buka Dokumen">
                                  <svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M18 13v6a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V8a2 2 0 0 1 2-2h6"/><polyline points="15 3 21 3 21 9"/><line x1="10" y1="14" x2="21" y2="3"/></svg>
                                </a>
                              `:""}
                              ${a?`
                                <button class="sp-icon-btn sp-edit-btn" title="Edit Data SP" data-id="${y.id}">
                                  <svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>
                                </button>
                              `:""}
                              ${o?`
                                <button class="sp-icon-btn delete sp-delete-btn" title="Hapus" data-id="${y.id}">
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
              ${b.map(y=>{let q=It(y.akhir_sp),v=Na(y.sp_type);return`
                  <div class="sp-card" data-id="${y.id}">
                    <div>
                      <div style="display:flex; justify-content:space-between; align-items:flex-start; margin-bottom:14px;">
                        <span class="sp-type-badge" style="background:${v.bg}; color:${v.color}; border:1px solid ${v.border};">
                          <span>${v.icon}</span>
                          <span>${v.label}</span>
                        </span>
                        <span class="sp-status-badge" style="background:${q.bg}; color:${q.color}; border:1px solid ${q.border};">
                          <span>${q.icon}</span>
                          <span>${q.shortLabel}</span>
                        </span>
                      </div>

                      <div style="display:flex; align-items:center; gap:12px; margin-bottom:12px;">
                        <div class="sp-avatar-box" style="width:48px; height:48px; font-size:1.05rem;">
                          ${y.employee_name.slice(0,2).toUpperCase()}
                        </div>
                        <div>
                          <h3 style="margin:0 0 3px 0; font-size:1.1rem; font-weight:700; color:#0F172A;">
                            ${y.employee_name}
                          </h3>
                          <div style="font-size:0.8rem; color:#64748B;">
                            \u{1F4CD} ${y.branch_name||"Cabang"} \u2022 <span style="color:#2563EB; font-weight:600;">${y.division||"FC"}</span>
                          </div>
                        </div>
                      </div>

                      <div style="background:#111827; border-radius:10px; padding:10px 12px; margin-bottom:14px; font-size:0.82rem; display:flex; justify-content:space-between;">
                        <div>
                          <div style="color:#64748B; font-size:0.72rem; text-transform:uppercase; font-weight:600;">Tanggal Terbit</div>
                          <div style="font-weight:700; color:#1E293B;">${ft(y.tanggal)}</div>
                        </div>
                        <div style="text-align:right;">
                          <div style="color:#64748B; font-size:0.72rem; text-transform:uppercase; font-weight:600;">Akhir SP</div>
                          <div style="font-weight:700; color:#0F172A;">${ft(y.akhir_sp)}</div>
                        </div>
                      </div>

                      <p style="font-size:0.83rem; color:#64748B; line-height:1.45; margin:0 0 16px 0; display:-webkit-box; -webkit-line-clamp:2; -webkit-box-orient:vertical; overflow:hidden;">
                        ${y.notes||y.description||"Pencatatan sanksi peringatan disiplin operasional."}
                      </p>
                    </div>

                    <div style="padding-top:12px; border-top:1px solid #29364B; display:flex; align-items:center; justify-content:space-between;">
                      <button class="btn btn-sm btn-ghost sp-detail-btn" style="border:1px solid #CBD5E1; border-radius:8px; font-weight:600; font-size:0.82rem; padding:6px 12px;" data-id="${y.id}">
                        \u{1F441}\uFE0F Rincian BAP
                      </button>

                      <div style="display:flex; align-items:center; gap:6px;">
                        ${y.document_link?`
                          <a href="${y.document_link}" target="_blank" rel="noopener noreferrer" class="sp-icon-btn" title="Buka Dokumen">
                            <svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M18 13v6a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V8a2 2 0 0 1 2-2h6"/><polyline points="15 3 21 3 21 9"/><line x1="10" y1="14" x2="21" y2="3"/></svg>
                          </a>
                        `:""}
                        ${a?`
                          <button class="sp-icon-btn sp-edit-btn" title="Edit Data SP" data-id="${y.id}">
                            <svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>
                          </button>
                        `:""}
                        ${o?`
                          <button class="sp-icon-btn delete sp-delete-btn" title="Hapus" data-id="${y.id}">
                            <svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><polyline points="3 6 5 6 21 6"/><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"/></svg>
                          </button>
                        `:""}
                      </div>
                    </div>
                  </div>
                `}).join("")}
            </div>
          `}
    `}function $(){let b=e.querySelector("#sp-btn-reset-filter");b&&(b.onclick=()=>{C="",c="all",E="all",p="all",u=1,clearTimeout(N);let y=e.querySelector("#sp-search-input");y&&(y.value="");let q=e.querySelector("#sp-filter-branch");q&&(q.value="all");let v=e.querySelector("#sp-filter-division");v&&(v.value="all"),e.querySelectorAll(".sp-pill").forEach(S=>{S.classList.toggle("active",S.getAttribute("data-filter")==="all")}),B({refreshStats:!1})}),e.querySelectorAll(".sp-detail-btn, .sp-view-detail-link").forEach(y=>{y.onclick=q=>{q.stopPropagation();let v=Number(y.getAttribute("data-id")),S=s.find(A=>A.id===v);S&&O(S)}}),e.querySelectorAll(".sp-edit-btn").forEach(y=>{y.onclick=q=>{q.stopPropagation();let v=Number(y.getAttribute("data-id")),S=s.find(A=>A.id===v);S&&g(S)}}),e.querySelectorAll(".sp-delete-btn").forEach(y=>{y.onclick=q=>{q.stopPropagation();let v=Number(y.getAttribute("data-id")),S=s.find(A=>A.id===v);S&&D(S)}})}function m(){let b=e.querySelector("#sp-pagination-container");if(!b)return;b.innerHTML="";let y=_e({page:u,pages:K,total:R,limit:F,onPage:q=>{u=q,B({refreshStats:!1})}});y&&b.appendChild(y)}function i(){let b=e.querySelector("#sp-list-container");if(!b)return;let y=P();b.innerHTML=J(y),m(),$()}function r(){let b=P(),y=f();e.innerHTML=`
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
              
              ${ie(t,"export")?`<button id="sp-btn-excel" class="btn btn-ghost sp-btn-excel" style="background:rgba(255,255,255,0.08); border:1px solid rgba(255,255,255,0.15); color:white; height:38px; padding:0 14px; border-radius:10px; font-weight:600; display:inline-flex; align-items:center; gap:6px; white-space:nowrap;">
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
                <div class="sp-stat-value" id="sp-stat-active">${y.active}</div>
                <div class="sp-stat-label">SP Aktif Berlaku</div>
              </div>
            </div>

            <div class="sp-stat-box">
              <div class="sp-stat-icon" style="background:rgba(245,158,11,0.2); color:#FBBF24;">\u26A0\uFE0F</div>
              <div>
                <div class="sp-stat-value" id="sp-stat-expiring">${y.expiring}</div>
                <div class="sp-stat-label">Segera Berakhir (&lt;30 Hari)</div>
              </div>
            </div>

            <div class="sp-stat-box">
              <div class="sp-stat-icon" style="background:rgba(16,185,129,0.2); color:#34D399;">\u2705</div>
              <div>
                <div class="sp-stat-value" id="sp-stat-expired">${y.expired}</div>
                <div class="sp-stat-label">Masa Sanksi Selesai</div>
              </div>
            </div>

            <div class="sp-stat-box">
              <div class="sp-stat-icon" style="background:rgba(59,130,246,0.2); color:#60A5FA;">\u{1F465}</div>
              <div>
                <div class="sp-stat-value"><span id="sp-stat-unique">${y.uniqueEmployees}</span> <span style="font-size:0.8rem; font-weight:400; color:#94A3B8;">karyawan</span></div>
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
              <input type="text" id="sp-search-input" class="sp-search-input" placeholder="Cari nama karyawan, cabang, atau divisi..." value="${C?C.replace(/"/g,"&quot;"):""}" />
            </div>

            <!-- Branch Dropdown -->
            <select id="sp-filter-branch" class="sp-filter-select">
              <option value="all">Semua Cabang (${l.length})</option>
              ${l.map(q=>`<option value="${q.value}" ${String(c)===String(q.value)?"selected":""}>${q.label}</option>`).join("")}
            </select>

            <!-- Division Dropdown -->
            <select id="sp-filter-division" class="sp-filter-select">
              <option value="all" ${E==="all"?"selected":""}>Semua Divisi</option>
              <option value="FACILITY CARE" ${E==="FACILITY CARE"?"selected":""}>FACILITY CARE</option>
              <option value="SECURITY" ${E==="SECURITY"?"selected":""}>SECURITY</option>
            </select>

            <div class="sp-view-toggle">
              <button type="button" class="sp-view-btn ${T==="card"?"active":""}" data-view="card" title="Tampilan Kartu">
                <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><rect x="3" y="3" width="7" height="7" rx="1"/><rect x="14" y="3" width="7" height="7" rx="1"/><rect x="14" y="14" width="7" height="7" rx="1"/><rect x="3" y="14" width="7" height="7" rx="1"/></svg>
                <span>Kartu</span>
              </button>
              <button type="button" class="sp-view-btn ${T==="table"?"active":""}" data-view="table" title="Tampilan Tabel">
                <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M3 9h18M3 15h18M9 3v18"/><rect x="3" y="3" width="18" height="18" rx="2"/></svg>
                <span>Tabel</span>
              </button>
            </div>
          </div>

          <!-- Filter Pills -->
          <div class="sp-filter-pills">
            <button class="sp-pill ${p==="all"?"active":""}" data-filter="all">
              <span>Semua SP</span>
              <span class="sp-pill-count" id="sp-pill-count-all">${y.total}</span>
            </button>
            <button class="sp-pill ${p==="active"?"active":""}" data-filter="active">
              <span>\u{1F534} SP Aktif</span>
              <span class="sp-pill-count" id="sp-pill-count-active">${y.active}</span>
            </button>
            <button class="sp-pill ${p==="expiring"?"active":""}" data-filter="expiring">
              <span>\u26A0\uFE0F &lt;30 Hari</span>
              <span class="sp-pill-count" id="sp-pill-count-expiring">${y.expiring}</span>
            </button>
            <button class="sp-pill ${p==="expired"?"active":""}" data-filter="expired">
              <span>\u2705 Selesai</span>
              <span class="sp-pill-count" id="sp-pill-count-expired">${y.expired}</span>
            </button>
            <button class="sp-pill ${p==="sp1"?"active":""}" data-filter="sp1">
              <span>SP 1</span>
              <span class="sp-pill-count" id="sp-pill-count-sp1">${y.sp1}</span>
            </button>
            <button class="sp-pill ${p==="sp2"?"active":""}" data-filter="sp2">
              <span>SP 2</span>
              <span class="sp-pill-count" id="sp-pill-count-sp2">${y.sp2}</span>
            </button>
            <button class="sp-pill ${p==="sp3"?"active":""}" data-filter="sp3">
              <span>SP 3</span>
              <span class="sp-pill-count" id="sp-pill-count-sp3">${y.sp3}</span>
            </button>
          </div>
        </div>

        <!-- Content Area -->
        <div id="sp-list-container">
          ${J(b)}
        </div>

        <!-- Pagination Container -->
        <div id="sp-pagination-container"></div>

      </div>
    `,m(),d()}function d(){let b=e.querySelector("#sp-btn-add");b&&(b.onclick=()=>g());let y=e.querySelector("#sp-btn-excel");y&&(y.onclick=()=>W());let q=e.querySelector("#sp-btn-refresh");q&&(q.onclick=()=>B({refreshStats:!0}));let v=e.querySelector("#sp-search-input");v&&v.addEventListener("input",U=>{C=U.target.value,clearTimeout(N),N=setTimeout(()=>{u=1,B({refreshStats:!1})},350)});let S=e.querySelector("#sp-filter-branch");S&&S.addEventListener("change",U=>{c=U.target.value,u=1,B({refreshStats:!1})});let A=e.querySelector("#sp-filter-division");A&&A.addEventListener("change",U=>{E=U.target.value,u=1,B({refreshStats:!1})}),e.querySelectorAll(".sp-view-btn").forEach(U=>{U.addEventListener("click",()=>{let H=U.getAttribute("data-view");!H||H===T||(T=H,localStorage.setItem("fcms_sp_view_mode",T),e.querySelectorAll(".sp-view-btn").forEach(ee=>{ee.classList.toggle("active",ee===U)}),i())})}),e.querySelectorAll(".sp-pill").forEach(U=>{U.addEventListener("click",()=>{p=U.getAttribute("data-filter")||"all",e.querySelectorAll(".sp-pill").forEach(H=>{H.classList.toggle("active",H===U)}),u=1,B({refreshStats:!1})})}),$()}await B({refreshStats:!0})}Fe();function ci(e){if(!e)return"-";try{let t=new Date(e);return isNaN(t.getTime())?e:t.toLocaleDateString("id-ID",{day:"numeric",month:"short",year:"numeric"})}catch{return e}}async function ui(e){let n=[],a=[],o=[],s="all",l="all",h="all",p="",c=1,E=0,C=1,T=!1,u={total:0,proses:0,selesai:0,currentMonth:0},F=null,R=!1,K=localStorage.getItem("fcms_mutasi_view_mode")||"card",I=de(),N=ie(I,"mutasi","write"),w=ie(I,"mutasi","write"),M=ie(I,"mutasi","delete");try{[a,o]=await Promise.all([ve(),Ie()])}catch(i){console.error("Error fetching cache:",i)}function k(){let i={sort:"date_desc"};return p.trim()&&(i.search=p.trim()),l&&l!=="all"&&(i.from_branch_id=l),h&&h!=="all"&&(i.to_branch_id=h),s&&s!=="all"&&(i.status=s),i}function x(){let i=(r,d)=>{let b=e.querySelector(r);b&&(b.textContent=d)};i("#mutasi-stat-total",u.total),i("#mutasi-stat-proses",u.proses),i("#mutasi-stat-selesai",u.selesai),i("#mutasi-stat-month",u.currentMonth),i("#mutasi-pill-count-all",u.total),i("#mutasi-pill-count-proses",u.proses),i("#mutasi-pill-count-selesai",u.selesai)}async function _({refreshStats:i=!0}={}){if(R=!0,!!!e.querySelector("#mutasi-list-container"))$();else{let d=e.querySelector("#mutasi-list-container");d&&(d.innerHTML='<div class="mutasi-empty-state"><p style="color:#64748B; margin:0;">Memuat data mutasi...</p></div>')}try{let d=new URLSearchParams({page:String(c),limit:String(20),...k()});(i||!T)&&d.set("include_stats","1");let b=await X(`/api/mutasi?${d.toString()}`);if(b.ok){let y=b.data?.data||[];n=Array.isArray(y)?y:[];let q=b.data?.pagination||{};if(E=Number(q.total??n.length),C=Math.max(1,Number(q.pages??Math.ceil(E/20))||1),c=Math.max(1,Number(q.page??c)),b.data?.stats&&(u=b.data.stats,T=!0),n.length===0&&c>C)return c=C,R=!1,_({refreshStats:i})}else n=[],E=0,C=1,z(b.data?.error||"Gagal memuat data mutasi")}catch(d){console.warn("Gagal load mutasi:",d),n=[],E=0,C=1,z("Gagal memuat data mutasi: "+d.message)}finally{R=!1,e.querySelector("#mutasi-list-container")?(x(),J()):$()}}async function B(){if(!ie(I,"export")){z("Anda tidak memiliki izin untuk mengekspor data.");return}try{let[{downloadExcel:i},r]=await Promise.all([Promise.resolve().then(()=>(we(),Oe)),ue("/api/mutasi/export",k())]),d=r.map((b,y)=>({No:y+1,Tanggal:b.tanggal||"","Nama Karyawan":b.employee_name||"","Cabang Asal":b.from_branch_name||"","Cabang Tujuan":b.to_branch_name||"",Status:b.status||"","Link Dokumen":b.document_link||""}));i(d,`Data_Mutasi_${new Date().toISOString().slice(0,10)}`),ae("Data mutasi berhasil diekspor ke Excel!")}catch(i){z("Gagal ekspor Excel: "+i.message)}}async function P(i){if(!M)return;let r=n.find(d=>d.id===i);me({title:"Hapus Data Mutasi",message:`Apakah Anda yakin ingin menghapus data mutasi untuk "${r?.employee_name||"karyawan ini"}"?`,confirmText:"Ya, Hapus",confirmClass:"btn-danger",onConfirm:async()=>{try{let d=await X(`/api/mutasi/${i}`,{method:"DELETE"});d.ok?(ae("Data mutasi berhasil dihapus"),await _({refreshStats:!0})):z(d.data?.error||"Gagal menghapus mutasi")}catch(d){z("Gagal menghapus data mutasi: "+d.message)}}})}function f(i=null){let r=!!i;if(r&&!w){z("Anda tidak memiliki izin untuk mengedit data mutasi.");return}if(!r&&!N){z("Anda tidak memiliki izin untuk menambah data mutasi.");return}if(!w)return;let d=!!i,b=i?.tanggal?i.tanggal.slice(0,10):new Date().toISOString().slice(0,10),y=document.createElement("div");y.innerHTML=`
      <div style="display:flex; flex-direction:column; gap:16px;">
        <div class="form-row" style="display:grid; grid-template-columns:repeat(auto-fit, minmax(200px, 1fr)); gap:12px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Tanggal Mutasi <span style="color:#DC2626;">*</span>
            </label>
            <input type="date" id="mutasi-modal-tanggal" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              value="${b}" required />
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Status Mutasi <span style="color:#DC2626;">*</span>
            </label>
            <select id="mutasi-modal-status" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" required>
              <option value="Proses" ${i?.status==="Proses"?"selected":""}>\u23F3 Proses</option>
              <option value="Selesai" ${i?.status==="Selesai"||!i?"selected":""}>\u2705 Selesai</option>
            </select>
          </div>
        </div>

        <div>
          <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
            Nama Karyawan <span style="color:#DC2626;">*</span>
          </label>
          <div id="mutasi-modal-emp-mount"></div>
        </div>

        <div class="form-row" style="display:grid; grid-template-columns:repeat(auto-fit, minmax(200px, 1fr)); gap:12px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Cabang Asal <span style="color:#DC2626;">*</span>
            </label>
            <select id="mutasi-modal-from" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" required>
              <option value="">-- Pilih Cabang Asal --</option>
              ${a.map(v=>`
                <option value="${v.value}" ${String(i?.from_branch_id)===String(v.value)?"selected":""}>
                  ${v.label}
                </option>
              `).join("")}
            </select>
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Cabang Tujuan <span style="color:#DC2626;">*</span>
            </label>
            <select id="mutasi-modal-to" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" required>
              <option value="">-- Pilih Cabang Tujuan --</option>
              ${a.map(v=>`
                <option value="${v.value}" ${String(i?.to_branch_id)===String(v.value)?"selected":""}>
                  ${v.label}
                </option>
              `).join("")}
            </select>
          </div>
        </div>

        <div>
          <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
            Link Berkas SK Mutasi / Dokumen (Opsional)
          </label>
          <input type="url" id="mutasi-modal-link" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
            placeholder="https://drive.google.com/file/d/..." value="${i?.document_link?String(i.document_link).replace(/"/g,"&quot;"):""}" />
        </div>
      </div>
    `;let q=ze(y.querySelector("#mutasi-modal-emp-mount"),{mode:"single",inputId:"mutasi-modal-emp",inputName:"employee_name",placeholder:"Cari nama karyawan...",required:!0,allowManual:!0,hasEmployeeIdSchema:!0,manualBadgeText:"Belum terhubung ke Master Karyawan",employees:o,initialEmployeeId:i?.employee_id||null,initialEmployeeName:i?.employee_name||""});pe({title:d?"\u270F\uFE0F Edit Data Mutasi Karyawan":"\u2795 Tambah Data Mutasi Karyawan",content:y,size:"md",confirmText:d?"Simpan Perubahan":"Simpan Mutasi",confirmClass:"btn-primary",onConfirm:async(v,S)=>{let A=y.querySelector("#mutasi-modal-tanggal").value,U=q.getValue(),H=(U.employee_name||"").trim(),ee=U.employee_id??null,Q=y.querySelector("#mutasi-modal-from").value,ne=y.querySelector("#mutasi-modal-to").value,Z=y.querySelector("#mutasi-modal-status").value,G=(y.querySelector("#mutasi-modal-link").value||"").trim();if(!A){z("Tanggal mutasi wajib diisi");return}if(!H){z("Nama karyawan wajib diisi"),q.inputEl?.focus();return}if(!Q){z("Cabang asal wajib dipilih");return}if(!ne){z("Cabang tujuan wajib dipilih");return}let L=a.find(te=>String(te.value)===String(Q)),j=a.find(te=>String(te.value)===String(ne)),V={tanggal:A,employee_id:ee,employee_name:H,from_branch_id:Number(Q),from_branch_name:L?L.label:"Cabang Asal",to_branch_id:Number(ne),to_branch_name:j?j.label:"Cabang Tujuan",status:Z,document_link:G,...d?{row_version:i?.row_version!==void 0&&i?.row_version!==null?Number(i.row_version):1}:{}};try{if(d){let te=await X(`/api/mutasi/${i.id}`,{method:"PUT",body:JSON.stringify(V)});if(te.ok)ae("Data mutasi berhasil diperbarui"),S(),await _({refreshStats:!0});else{if(te.status===409){z("Konflik OCC: Data telah diubah oleh pengguna lain. Silakan muat ulang halaman."),S(),await _({refreshStats:!0});return}z(te.data?.error||"Gagal memperbarui data mutasi")}}else{let te=await X("/api/mutasi",{method:"POST",body:JSON.stringify(V)});te.ok?(ae("Data mutasi baru berhasil ditambahkan"),S(),await _({refreshStats:!0})):z(te.data?.error||"Gagal menambahkan mutasi")}}catch(te){z("Terjadi kesalahan: "+te.message)}}})}function g(){return n}function D(i){return`
          ${i.length===0?`
            <div class="mutasi-empty-state">
              <div style="font-size:3rem; margin-bottom:12px;">\u{1F50D}</div>
              <h3 style="font-size:1.2rem; font-weight:700; color:#1E293B; margin:0 0 6px 0;">Tidak Ada Data Mutasi Ditemukan</h3>
              <p style="color:#64748B; font-size:0.9rem; margin:0 0 20px 0;">Silakan sesuaikan kata kunci pencarian atau reset filter aktif.</p>
              <button id="mutasi-btn-reset-filter" class="btn btn-ghost" style="border:1px solid #CBD5E1; border-radius:8px; padding:8px 16px;">
                Reset Filter
              </button>
            </div>
          `:K==="table"?`
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
                      ${w||M?'<th style="width:90px; text-align:center;">Aksi</th>':""}
                    </tr>
                  </thead>
                  <tbody>
                    ${i.map((r,d)=>{let b=r.status==="Selesai",y=(r.employee_name||"Karyawan").split(" ").map(q=>q[0]).filter(Boolean).slice(0,2).join("").toUpperCase();return`
                        <tr data-id="${r.id}">
                          <td style="text-align:center; color:#94A3B8; font-weight:600; font-size:0.8rem;">${(c-1)*20+d+1}</td>
                          <td>
                            <div style="display:flex; align-items:center; gap:10px;">
                              <div class="mutasi-avatar-box" style="width:34px; height:34px; font-size:0.78rem; border-radius:8px;">
                                ${y}
                              </div>
                              <div>
                                <div style="font-weight:700; color:#0F172A; white-space:nowrap;">${r.employee_name}</div>
                                <div style="font-size:0.72rem; color:#64748B;">ID Mutasi: #${r.id}</div>
                              </div>
                            </div>
                          </td>
                          <td style="white-space:nowrap; font-weight:600; color:#1E293B;">
                            \u{1F5D3}\uFE0F ${ci(r.tanggal)}
                          </td>
                          <td style="white-space:nowrap;">
                            <div style="font-weight:600; color:#334155; font-size:0.85rem; display:inline-flex; align-items:center; gap:4px;">
                              <span>\u{1F4CD}</span>
                              <span>${r.from_branch_name||"-"}</span>
                            </div>
                          </td>
                          <td style="white-space:nowrap;">
                            <div style="font-weight:700; color:#2563EB; font-size:0.85rem; display:inline-flex; align-items:center; gap:5px;">
                              <span>\u2794</span>
                              <span>${r.to_branch_name||"-"}</span>
                            </div>
                          </td>
                          <td style="text-align:center; white-space:nowrap;">
                            <span class="mutasi-status-badge ${b?"selesai":"proses"}">
                              <span>${b?"\u2705":"\u23F3"}</span>
                              <span>${r.status||"Proses"}</span>
                            </span>
                          </td>
                          <td style="text-align:center; white-space:nowrap;">
                            ${r.document_link?`
                              <a href="${r.document_link}" target="_blank" rel="noopener noreferrer" class="btn btn-sm btn-ghost" style="border:1px solid #CBD5E1; border-radius:6px; font-size:0.76rem; padding:4px 8px; text-decoration:none; display:inline-flex; align-items:center; gap:4px;">
                                <span>\u{1F4C4} Buka SK</span>
                              </a>
                            `:`
                              <span style="font-size:0.75rem; color:#94A3B8;">(Tanpa Dokumen)</span>
                            `}
                          </td>
                          ${w||M?`
                            <td style="text-align:center; white-space:nowrap;">
                              <div style="display:inline-flex; align-items:center; gap:6px;">
                                ${w?`
                                  <button class="mutasi-icon-btn mutasi-edit-btn" title="Edit Data Mutasi" data-id="${r.id}">
                                    <svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>
                                  </button>
                                `:""}
                                ${M?`
                                  <button class="mutasi-icon-btn delete mutasi-delete-btn" title="Hapus Mutasi" data-id="${r.id}">
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
              ${i.map(r=>{let d=r.status==="Selesai",b=(r.employee_name||"Karyawan").split(" ").map(y=>y[0]).filter(Boolean).slice(0,2).join("").toUpperCase();return`
                  <div class="mutasi-card" data-id="${r.id}">
                    <div>
                      <!-- Card Header: Status & Tanggal -->
                      <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:12px;">
                        <span class="mutasi-status-badge ${d?"selesai":"proses"}">
                          <span>${d?"\u2705":"\u23F3"}</span>
                          <span>${r.status||"Proses"}</span>
                        </span>
                        <span style="font-size:0.8rem; font-weight:600; color:#64748B; display:inline-flex; align-items:center; gap:4px;">
                          <span>\u{1F5D3}\uFE0F</span>
                          <span>${ci(r.tanggal)}</span>
                        </span>
                      </div>

                      <!-- Employee Info -->
                      <div style="display:flex; align-items:center; gap:12px; margin-bottom:12px;">
                        <div class="mutasi-avatar-box">
                          ${b}
                        </div>
                        <div style="min-width:0;">
                          <h4 style="margin:0 0 3px 0; font-size:1.02rem; font-weight:700; color:#0F172A; white-space:nowrap; overflow:hidden; text-overflow:ellipsis;">
                            ${r.employee_name}
                          </h4>
                          <div style="font-size:0.75rem; color:#64748B;">
                            ID Mutasi: #${r.id}
                          </div>
                        </div>
                      </div>

                      <!-- Route Visual Box -->
                      <div class="mutasi-route-box">
                        <div class="mutasi-route-node mutasi-from-node">
                          <div class="mutasi-route-label">Cabang Asal</div>
                          <div class="mutasi-route-branch" title="${r.from_branch_name||"-"}">
                            ${r.from_branch_name||"-"}
                          </div>
                        </div>

                        <div class="mutasi-route-arrow">
                          \u2794
                        </div>

                        <div class="mutasi-route-node mutasi-to-node">
                          <div class="mutasi-route-label">Cabang Tujuan</div>
                          <div class="mutasi-route-branch" style="color:#2563EB;" title="${r.to_branch_name||"-"}">
                            ${r.to_branch_name||"-"}
                          </div>
                        </div>
                      </div>
                    </div>

                    <!-- Card Footer: Dokumen & Aksi -->
                    <div style="display:flex; justify-content:space-between; align-items:center; padding-top:12px; border-top:1px solid #29364B; margin-top:8px;">
                      <div>
                        ${r.document_link?`
                          <a href="${r.document_link}" target="_blank" rel="noopener noreferrer" class="btn btn-sm btn-ghost" style="border:1px solid #CBD5E1; border-radius:6px; font-size:0.78rem; padding:4px 8px; text-decoration:none; display:inline-flex; align-items:center; gap:4px;">
                            <span>\u{1F4C4} Buka SK</span>
                          </a>
                        `:`
                          <span style="font-size:0.75rem; color:#94A3B8;">(Tanpa Dokumen)</span>
                        `}
                      </div>

                      ${w||M?`
                        <div style="display:flex; align-items:center; gap:6px;">
                          ${w?`
                            <button class="mutasi-icon-btn mutasi-edit-btn" title="Edit Data Mutasi" data-id="${r.id}">
                              <svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>
                            </button>
                          `:""}
                          ${M?`
                            <button class="mutasi-icon-btn delete mutasi-delete-btn" title="Hapus Mutasi" data-id="${r.id}">
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
    `}function O(){let i=e.querySelector("#mutasi-pagination-container");if(!i)return;i.innerHTML="";let r=_e({page:c,pages:C,total:E,limit:20,onPage:d=>{c=d,_({refreshStats:!1})}});r&&i.appendChild(r)}function W(){let i=e.querySelector("#mutasi-btn-reset-filter");i&&(i.onclick=()=>{p="",l="all",h="all",s="all",c=1;let r=e.querySelector("#mutasi-search-input");r&&(r.value="");let d=e.querySelector("#mutasi-filter-from");d&&(d.value="all");let b=e.querySelector("#mutasi-filter-to");b&&(b.value="all"),e.querySelectorAll(".mutasi-pill").forEach(y=>{y.classList.toggle("active",y.getAttribute("data-status")==="all")}),_({refreshStats:!1})}),e.querySelectorAll(".mutasi-edit-btn").forEach(r=>{r.onclick=d=>{d.stopPropagation();let b=Number(r.getAttribute("data-id")),y=n.find(q=>q.id===b);y&&f(y)}}),e.querySelectorAll(".mutasi-delete-btn").forEach(r=>{r.onclick=d=>{d.stopPropagation();let b=Number(r.getAttribute("data-id"));P(b)}})}function J(){let i=e.querySelector("#mutasi-list-container");if(!i)return;let r=g();i.innerHTML=D(r),O(),W()}function $(){let i=g(),r=u.total,d=u.proses,b=u.selesai,y=u.currentMonth;e.innerHTML=`
      <style>.mutasi-page{display:flex;flex-direction:column;gap:20px}.mutasi-hero-card{background:linear-gradient(135deg,#0F172A 0%,#1E293B 100%);border-radius:16px;padding:24px 28px;color:white;box-shadow:0 10px 25px -5px rgba(15,23,42,0.25);position:relative;overflow:hidden}.mutasi-hero-card::after{content:'';position:absolute;top:-40px;right:-40px;width:220px;height:220px;background:radial-gradient(circle,rgba(59,130,246,0.15) 0%,transparent 70%);border-radius:50%;pointer-events:none}.mutasi-hero-actions{display:flex;align-items:center;gap:8px;flex-wrap:wrap}.mutasi-hero-actions .btn{height:38px;box-sizing:border-box;transition:all 0.2s ease}.mutasi-btn-add:hover{background:#1D4ED8 !important;transform:translateY(-1px)}.mutasi-btn-excel:hover,.mutasi-btn-refresh:hover{background:rgba(255,255,255,0.16) !important;border-color:rgba(255,255,255,0.3) !important;transform:translateY(-1px)}.mutasi-stats-row{display:grid;grid-template-columns:repeat(4,1fr);gap:14px;margin-top:22px}.mutasi-stat-box{background:rgba(255,255,255,0.06);border:1px solid rgba(255,255,255,0.1);border-radius:12px;padding:14px 16px;backdrop-filter:blur(8px);display:flex;align-items:center;gap:14px}.mutasi-stat-icon{width:44px;height:44px;border-radius:10px;display:flex;align-items:center;justify-content:center;font-size:1.3rem;flex-shrink:0}.mutasi-stat-val{font-size:1.55rem;font-weight:800;line-height:1.1;color:#F8FAFC}.mutasi-stat-lbl{font-size:0.76rem;text-transform:uppercase;letter-spacing:0.05em;color:#94A3B8;font-weight:600;margin-top:2px}.mutasi-controls-card{background:#172033;border:1px solid #29364B;border-radius:14px;padding:18px 20px;box-shadow:0 2px 4px rgba(0,0,0,0.02);display:flex;flex-direction:column;gap:16px}.mutasi-search-row{display:flex;gap:12px;align-items:center;flex-wrap:wrap}.mutasi-search-wrap{position:relative;flex:1;min-width:240px}.mutasi-search-icon{position:absolute;left:14px;top:50%;transform:translateY(-50%);color:#94A3B8;pointer-events:none}.mutasi-search-input{width:100%;padding:10px 14px 10px 40px;border:1px solid #29364B;border-radius:10px;font-size:0.9rem;outline:none;transition:border-color 0.15s ease}.mutasi-search-input:focus{border-color:#2563EB;box-shadow:0 0 0 3px rgba(37,99,235,0.1)}.mutasi-select{padding:10px 14px;border:1px solid #29364B;border-radius:10px;font-size:0.9rem;color:#F8FAFC;background:#172033;outline:none;min-width:180px;cursor:pointer}.mutasi-select:focus{border-color:#2563EB}.mutasi-filter-pills{display:flex;align-items:center;gap:8px;overflow-x:auto;white-space:nowrap;scrollbar-width:none;padding-bottom:2px}.mutasi-filter-pills::-webkit-scrollbar{display:none}.mutasi-pill{padding:6px 14px;border-radius:20px;font-size:0.82rem;font-weight:600;cursor:pointer;border:1px solid #29364B;background:#111827;color:#64748B;transition:all 0.15s ease;display:inline-flex;align-items:center;gap:6px;user-select:none;flex-shrink:0}.mutasi-pill:hover{background:#22314A;color:#F8FAFC}.mutasi-pill.active{background:#2563EB;color:white;border-color:#2563EB;box-shadow:0 2px 6px rgba(37,99,235,0.25)}.mutasi-view-toggle{display:inline-flex;align-items:center;background:#111827;padding:3px;border-radius:10px;border:1px solid #29364B;gap:3px;flex-shrink:0}.mutasi-view-btn{display:inline-flex;align-items:center;gap:6px;padding:6px 12px;border-radius:7px;font-size:0.82rem;font-weight:600;border:none;background:transparent;color:#64748B;cursor:pointer;transition:all 0.15s ease;height:34px;box-sizing:border-box;user-select:none}.mutasi-view-btn:hover{color:#F8FAFC;background:rgba(255,255,255,0.5)}.mutasi-view-btn.active{background:#172033;color:#2563EB;font-weight:700;box-shadow:0 1px 4px rgba(0,0,0,0.08)}.mutasi-table-card{background:#172033;border:1px solid #29364B;border-radius:14px;box-shadow:0 2px 6px rgba(0,0,0,0.03);overflow:hidden}.mutasi-table-responsive{width:100%;overflow-x:auto;-webkit-overflow-scrolling:touch}.mutasi-table{width:100%;border-collapse:collapse;text-align:left;font-size:0.88rem;min-width:860px}.mutasi-table th{background:#111827;color:#475569;font-size:0.76rem;font-weight:700;text-transform:uppercase;letter-spacing:0.05em;padding:13px 16px;border-bottom:1.5px solid #E2E8F0;white-space:nowrap}.mutasi-table td{padding:13px 16px;border-bottom:1px solid #29364B;color:#F8FAFC;vertical-align:middle}.mutasi-table tbody tr:hover td{background:#111827}.mutasi-table tbody tr:last-child td{border-bottom:none}.mutasi-cards-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(320px,1fr));gap:16px}.mutasi-card{background:#172033;border:1px solid #29364B;border-radius:14px;padding:18px;box-shadow:0 2px 6px rgba(0,0,0,0.03);transition:transform 0.15s ease,box-shadow 0.15s ease;display:flex;flex-direction:column;justify-content:space-between;position:relative}.mutasi-card:hover{transform:translateY(-2px);box-shadow:0 8px 18px rgba(0,0,0,0.06);border-color:#CBD5E1}.mutasi-avatar-box{width:44px;height:44px;border-radius:10px;background:linear-gradient(135deg,#2563EB 0%,#1D4ED8 100%);color:white;font-weight:700;font-size:1rem;display:flex;align-items:center;justify-content:center;flex-shrink:0;box-shadow:0 3px 6px rgba(37,99,235,0.2)}.mutasi-route-box{background:#111827;border:1px dashed #CBD5E1;border-radius:10px;padding:12px 14px;margin:12px 0;display:flex;align-items:center;justify-content:space-between;gap:10px;min-width:0;box-sizing:border-box;width:100%}.mutasi-route-node{flex:1;min-width:0}.mutasi-to-node{text-align:right}.mutasi-route-label{font-size:0.7rem;text-transform:uppercase;font-weight:700;letter-spacing:0.04em;color:#94A3B8;margin-bottom:2px}.mutasi-route-branch{font-size:0.85rem;font-weight:700;color:#F8FAFC;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}.mutasi-route-arrow{display:flex;align-items:center;justify-content:center;width:32px;height:32px;border-radius:50%;background:#EFF6FF;color:#2563EB;font-size:1rem;flex-shrink:0}.mutasi-status-badge{display:inline-flex;align-items:center;gap:5px;padding:4px 10px;border-radius:20px;font-size:0.74rem;font-weight:700}.mutasi-status-badge.selesai{background:#DCFCE7;color:#15803D;border:1px solid #BBF7D0}.mutasi-status-badge.proses{background:#FEF3C7;color:#B45309;border:1px solid #FDE68A}.mutasi-icon-btn{width:32px;height:32px;border-radius:8px;border:1px solid #29364B;background:#172033;color:#64748B;display:inline-flex;align-items:center;justify-content:center;cursor:pointer;transition:all 0.15s ease}.mutasi-icon-btn:hover{background:#22314A;color:#F8FAFC;border-color:#CBD5E1}.mutasi-icon-btn.delete:hover{background:#FEE2E2;color:#DC2626;border-color:#FCA5A5}.mutasi-empty-state{grid-column:1 / -1;background:#172033;border:2px dashed #E2E8F0;border-radius:16px;padding:48px 24px;text-align:center}@media (max-width:768px){.mutasi-hero-card{padding:18px 16px}.mutasi-stats-row{grid-template-columns:1fr 1fr;gap:10px}.mutasi-stat-box{padding:10px 12px}.mutasi-stat-icon{width:38px;height:38px;font-size:1.1rem}.mutasi-stat-val{font-size:1.3rem}.mutasi-hero-actions{width:100% !important;display:flex !important;align-items:center !important;gap:6px !important;flex-wrap:nowrap !important;overflow-x:auto !important;scrollbar-width:none !important;padding-bottom:2px !important}.mutasi-hero-actions::-webkit-scrollbar{display:none}.mutasi-btn-add{flex:1 1 auto !important;justify-content:center !important;padding:8px 10px !important;font-size:0.82rem !important;white-space:nowrap !important}.mutasi-btn-excel{flex:1 1 auto !important;justify-content:center !important;padding:8px 10px !important;font-size:0.82rem !important;white-space:nowrap !important}.mutasi-btn-refresh{width:38px !important;height:38px !important;padding:0 !important;flex-shrink:0 !important;justify-content:center !important}.mutasi-controls-card{padding:14px;gap:12px}.mutasi-search-row{flex-direction:column;gap:10px}.mutasi-select{width:100%}.mutasi-view-toggle{width:100%;display:flex}.mutasi-view-btn{flex:1;justify-content:center}.mutasi-table-card{border-radius:12px}.mutasi-cards-grid{grid-template-columns:minmax(0,1fr) !important;width:100% !important;max-width:100% !important;min-width:0 !important;box-sizing:border-box !important}.mutasi-card{width:100% !important;max-width:100% !important;min-width:0 !important;box-sizing:border-box !important;overflow:hidden;padding:14px !important}.mutasi-route-box{flex-direction:column !important;align-items:stretch !important;gap:8px !important;padding:11px 12px !important;width:100% !important;box-sizing:border-box !important}.mutasi-route-node{width:100% !important;min-width:0 !important;text-align:left !important}.mutasi-to-node{text-align:left !important}.mutasi-route-branch{white-space:normal !important;word-break:break-word !important;overflow:visible !important;text-overflow:clip !important;font-size:0.88rem !important;line-height:1.35 !important}.mutasi-route-arrow{align-self:center !important;width:26px !important;height:26px !important;font-size:0.85rem !important;transform:rotate(90deg) !important;margin:-2px 0 !important}}</style>

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
              ${N?`
                <button id="mutasi-btn-add" class="btn btn-primary mutasi-btn-add" style="background:#2563EB; height:38px; padding:0 16px; border-radius:10px; font-weight:700; display:inline-flex; align-items:center; gap:6px; box-shadow:0 4px 12px rgba(37,99,235,0.3); white-space:nowrap;">
                  <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24"><line x1="12" y1="5" x2="12" y2="19"/><line x1="5" y1="12" x2="19" y2="12"/></svg>
                  <span>Tambah Mutasi</span>
                </button>
              `:""}
              
              ${ie(I,"export")?`<button id="mutasi-btn-excel" class="btn btn-ghost mutasi-btn-excel" style="background:rgba(255,255,255,0.08); border:1px solid rgba(255,255,255,0.15); color:white; height:38px; padding:0 14px; border-radius:10px; font-weight:600; display:inline-flex; align-items:center; gap:6px; white-space:nowrap;">
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
                <div class="mutasi-stat-val" id="mutasi-stat-total">${r}</div>
                <div class="mutasi-stat-lbl">Total Mutasi</div>
              </div>
            </div>

            <div class="mutasi-stat-box">
              <div class="mutasi-stat-icon" style="background:rgba(245,158,11,0.2); color:#FBBF24;">
                \u23F3
              </div>
              <div>
                <div class="mutasi-stat-val" id="mutasi-stat-proses">${d}</div>
                <div class="mutasi-stat-lbl">Dalam Proses</div>
              </div>
            </div>

            <div class="mutasi-stat-box">
              <div class="mutasi-stat-icon" style="background:rgba(34,197,94,0.2); color:#4ADE80;">
                \u2705
              </div>
              <div>
                <div class="mutasi-stat-val" id="mutasi-stat-selesai">${b}</div>
                <div class="mutasi-stat-lbl">Mutasi Selesai</div>
              </div>
            </div>

            <div class="mutasi-stat-box">
              <div class="mutasi-stat-icon" style="background:rgba(168,85,247,0.2); color:#C084FC;">
                \u{1F5D3}\uFE0F
              </div>
              <div>
                <div class="mutasi-stat-val" id="mutasi-stat-month">${y}</div>
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
              <input type="text" id="mutasi-search-input" class="mutasi-search-input" placeholder="Cari nama karyawan, cabang asal, atau cabang tujuan..." value="${p}" />
            </div>

            <select id="mutasi-filter-from" class="mutasi-select">
              <option value="all">Semua Cabang Asal (${a.length})</option>
              ${a.map(q=>`<option value="${q.value}" ${String(l)===String(q.value)?"selected":""}>${q.label}</option>`).join("")}
            </select>

            <select id="mutasi-filter-to" class="mutasi-select">
              <option value="all">Semua Cabang Tujuan (${a.length})</option>
              ${a.map(q=>`<option value="${q.value}" ${String(h)===String(q.value)?"selected":""}>${q.label}</option>`).join("")}
            </select>

            <div class="mutasi-view-toggle">
              <button type="button" class="mutasi-view-btn ${K==="card"?"active":""}" data-view="card" title="Tampilan Kartu">
                <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><rect x="3" y="3" width="7" height="7" rx="1"/><rect x="14" y="3" width="7" height="7" rx="1"/><rect x="14" y="14" width="7" height="7" rx="1"/><rect x="3" y="14" width="7" height="7" rx="1"/></svg>
                <span>Kartu</span>
              </button>
              <button type="button" class="mutasi-view-btn ${K==="table"?"active":""}" data-view="table" title="Tampilan Tabel">
                <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M3 9h18M3 15h18M9 3v18"/><rect x="3" y="3" width="18" height="18" rx="2"/></svg>
                <span>Tabel</span>
              </button>
            </div>
          </div>

          <!-- Status Filter Pills (Swipeable on Mobile) -->
          <div class="mutasi-filter-pills">
            <div class="mutasi-pill ${s==="all"?"active":""}" data-status="all">
              <span>Semua Mutasi</span>
              <span class="badge" id="mutasi-pill-count-all" style="background:rgba(0,0,0,0.06); padding:2px 6px; border-radius:10px; font-size:0.75rem;">${r}</span>
            </div>
            <div class="mutasi-pill ${s==="Proses"?"active":""}" data-status="Proses">
              <span style="color:#D97706;">\u23F3</span>
              <span>Proses</span>
              <span class="badge" id="mutasi-pill-count-proses" style="background:rgba(217,119,6,0.15); color:#B45309; padding:2px 6px; border-radius:10px; font-size:0.75rem;">${d}</span>
            </div>
            <div class="mutasi-pill ${s==="Selesai"?"active":""}" data-status="Selesai">
              <span style="color:#16A34A;">\u2705</span>
              <span>Selesai</span>
              <span class="badge" id="mutasi-pill-count-selesai" style="background:rgba(22,163,74,0.15); color:#15803D; padding:2px 6px; border-radius:10px; font-size:0.75rem;">${b}</span>
            </div>
          </div>
        </div>

        <!-- Content Area: Murni Grid Kartu -->
        <div id="mutasi-list-container">
          ${D(i)}
        </div>
        <div id="mutasi-pagination-container"></div>

      </div>
    `,m()}function m(){let i=e.querySelector("#mutasi-btn-add");i&&(i.onclick=()=>f());let r=e.querySelector("#mutasi-btn-excel");r&&(r.onclick=()=>B());let d=e.querySelector("#mutasi-btn-refresh");d&&(d.onclick=()=>_({refreshStats:!0})),e.querySelectorAll(".mutasi-view-btn").forEach(v=>{v.onclick=()=>{let S=v.getAttribute("data-view");S&&S!==K&&(K=S,localStorage.setItem("fcms_mutasi_view_mode",S),e.querySelectorAll(".mutasi-view-btn").forEach(A=>{A.classList.toggle("active",A.getAttribute("data-view")===S)}),J())}});let b=e.querySelector("#mutasi-search-input");b&&(b.oninput=v=>{p=v.target.value,clearTimeout(F),F=setTimeout(()=>{c=1,_({refreshStats:!1})},350)});let y=e.querySelector("#mutasi-filter-from");y&&(y.onchange=v=>{l=v.target.value,c=1,_({refreshStats:!1})});let q=e.querySelector("#mutasi-filter-to");q&&(q.onchange=v=>{h=v.target.value,c=1,_({refreshStats:!1})}),e.querySelectorAll(".mutasi-pill").forEach(v=>{v.onclick=()=>{s=v.getAttribute("data-status")||"all",e.querySelectorAll(".mutasi-pill").forEach(S=>{S.classList.toggle("active",S===v)}),c=1,_({refreshStats:!1})}}),O(),W()}await _({refreshStats:!0})}Fe();var mi=["General Cleaning","Deep Cleaning","Cover manpower","Longshift","Pekerjaan urgent"];function gi(e,t,n){if(!e||!t)return 0;try{let[a,o]=e.split(":").map(Number),[s,l]=t.split(":").map(Number),h=s*60+l-(a*60+o);h<0&&(h+=24*60);let p=h/60-(parseFloat(n)||0);return p<0&&(p=0),Number.isInteger(p)?p:parseFloat(p.toFixed(2))}catch{return 0}}function Ra(e=""){let t=(e||"").toLowerCase();return t.includes("general")?{bg:"#EFF6FF",color:"#1D4ED8",border:"#BFDBFE",icon:"\u{1F9F9}"}:t.includes("deep")?{bg:"#F5F3FF",color:"#6D28D9",border:"#DDD6FE",icon:"\u2728"}:t.includes("cover")?{bg:"#ECFDF5",color:"#047857",border:"#A7F3D0",icon:"\u{1F465}"}:t.includes("longshift")?{bg:"#FFFBEB",color:"#B45309",border:"#FDE68A",icon:"\u23F3"}:t.includes("urgent")?{bg:"#FFF1F2",color:"#BE123C",border:"#FECDD3",icon:"\u{1F6A8}"}:{bg:"#F8FAFC",color:"#475569",border:"#CBD5E1",icon:"\u23F1\uFE0F"}}function ja(e){if(!e)return"-";try{let t=new Date(e);return isNaN(t.getTime())?e:t.toLocaleDateString("id-ID",{day:"numeric",month:"short",year:"numeric"})}catch{return e}}async function bi(e){let t=[],n=[],a=[],o="all",s="all",l="",h=!1,p=localStorage.getItem("fcms_overtime_view_mode")||"card",c=1,E=20,C=0,T=1,u={totalRecords:0,totalHours:0,avgDuration:0,currentMonth:0},F=null,R=!1,K=de(),I=ie(K,"overtime","create"),N=ie(K,"overtime","update"),w=ie(K,"overtime","delete");try{[n,a]=await Promise.all([ve(),Ie()])}catch(m){console.error("Error fetching cache:",m)}function M(){let m={sort:"date_desc"};return l&&l.trim()&&(m.search=l.trim()),s&&s!=="all"&&(m.branch_id=s),o&&o!=="all"&&(m.reason=o),m}function k(){let m=(d,b)=>{let y=e.querySelector(d);y&&(y.textContent=String(b??0))},i=Number(u.totalHours||0),r=u.avgDuration!=null?Number(u.avgDuration).toFixed(1):"0.0";m("#lembur-stat-hours",Number.isInteger(i)?i:i.toFixed(1)),m("#lembur-stat-records",u.totalRecords),m("#lembur-stat-avg",r),m("#lembur-stat-month",u.currentMonth),m("#lembur-pill-count-all",u.totalRecords)}async function x({refreshStats:m=!0}={}){h=!0;try{let i=M(),r=new URLSearchParams({page:String(c),limit:String(E),...i});m&&r.set("include_stats","1");let d=await X("/api/overtime?"+r.toString());if(d.ok&&d.data&&Array.isArray(d.data.data)){let b=d.data.pagination||{};if(C=Number(b.total??d.data.data.length),T=Math.max(1,Number(b.pages||b.totalPages||Math.ceil(C/E)||1)),d.data.data.length===0&&c>T)return c=T,h=!1,x({refreshStats:m});t=d.data.data,d.data.stats&&(u={...u,...d.data.stats})}else t=[],C=0,T=1,z(d.data?.error||"Gagal memuat data lembur")}catch(i){console.warn("Gagal load lembur:",i),t=[],C=0,T=1,z("Gagal memuat data lembur: "+i.message)}finally{h=!1,R?(m&&k(),W()):(R=!0,J())}}async function _(){if(!ie(K,"export")){z("Anda tidak memiliki izin untuk mengekspor data.");return}try{let{downloadExcel:m}=await Promise.resolve().then(()=>(we(),Oe)),i=await ue("/api/overtime/export",M()),r=i.map((d,b)=>({No:b+1,Tanggal:d.date||"","Cabang Lembur":d.branch_name||"","Nama Karyawan":d.employee_name||"","Jam Mulai":d.start_time||"","Jam Selesai":d.end_time||"","Istirahat (Jam)":d.break_hours!=null?d.break_hours:0,"Total Jam":d.total_hours!=null?d.total_hours:0,"Alasan / Keterangan":d.reason||""}));m(r,`Data_Lembur_${new Date().toISOString().slice(0,10)}`),ae(`Berhasil mengunduh ${i.length} data lembur ke Excel!`)}catch(m){z("Gagal ekspor Excel: "+m.message)}}async function B(m){if(!w)return;let i=t.find(r=>r.id===m);me({title:"Hapus Catatan Lembur",message:`Apakah Anda yakin ingin menghapus catatan lembur untuk "${i?.employee_name||"karyawan ini"}" pada tanggal ${ja(i?.date)}?`,confirmText:"Ya, Hapus",confirmClass:"btn-danger",onConfirm:async()=>{try{let r=await X(`/api/overtime/${m}`,{method:"DELETE"});r.ok?(ae("Catatan lembur berhasil dihapus"),await x({refreshStats:!0})):z(r.data?.error||"Gagal menghapus lembur")}catch(r){z("Gagal menghapus catatan lembur: "+r.message)}}})}function P(m=null){let i=!!m;if(i?!N:!I)return;let r=m?.date?m.date.slice(0,10):new Date().toISOString().slice(0,10),d=m?.start_time||"22:00",b=m?.end_time||"08:00",y=m?.break_hours!=null?m.break_hours:1,q=m?.total_hours!=null?m.total_hours:gi(d,b,y),v=document.createElement("div");v.innerHTML=`
      <div style="display:flex; flex-direction:column; gap:16px;">
        <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Tanggal Lembur <span style="color:#DC2626;">*</span>
            </label>
            <input type="date" id="ot-modal-date" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              value="${r}" required />
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Cabang Penugasan <span style="color:#DC2626;">*</span>
            </label>
            <select id="ot-modal-branch" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" required>
              <option value="">-- Pilih Cabang --</option>
              ${n.map(ne=>`
                <option value="${ne.value}" ${String(m?.branch_id)===String(ne.value)?"selected":""}>
                  ${ne.label}
                </option>
              `).join("")}
            </select>
          </div>
        </div>

        <div>
          <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
            Nama Karyawan <span style="color:#DC2626;">*</span>
          </label>
          <div id="ot-modal-emp-mount"></div>
        </div>

        <!-- Shift & Waktu Lembur -->
        <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Jam Mulai <span style="color:#DC2626;">*</span>
            </label>
            <input type="time" id="ot-modal-start" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              value="${d}" required />
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Jam Selesai <span style="color:#DC2626;">*</span>
            </label>
            <input type="time" id="ot-modal-end" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              value="${b}" required />
          </div>
        </div>

        <!-- Durasi & Istirahat -->
        <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Istirahat (Jam) <span style="color:#DC2626;">*</span>
            </label>
            <input type="number" step="0.5" min="0" max="10" id="ot-modal-break" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              value="${y}" required />
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Total Jam Lembur <span style="color:#DC2626;">*</span>
            </label>
            <input type="number" step="0.1" min="0" id="ot-modal-total" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #29364B; border-radius:8px; background:#1B2940; color:#38BDF8; font-weight:700; color:#2563EB;" 
              value="${q}" required />
          </div>
        </div>

        <div>
          <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
            Alasan / Keterangan Penugasan Lembur <span style="color:#DC2626;">*</span>
          </label>
          <select id="ot-modal-reason" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" required>
            ${mi.map(ne=>`
              <option value="${ne}" ${m?.reason===ne?"selected":""}>${ne}</option>
            `).join("")}
          </select>
        </div>
      </div>
    `;let S=v.querySelector("#ot-modal-start"),A=v.querySelector("#ot-modal-end"),U=v.querySelector("#ot-modal-break"),H=v.querySelector("#ot-modal-total");function ee(){let ne=S.value,Z=A.value,G=U.value;H.value=gi(ne,Z,G)}S.addEventListener("input",ee),A.addEventListener("input",ee),U.addEventListener("input",ee);let Q=ze(v.querySelector("#ot-modal-emp-mount"),{mode:"single",inputId:"ot-modal-emp",inputName:"employee_name",placeholder:"Cari nama karyawan...",required:!0,allowManual:!0,hasEmployeeIdSchema:!0,manualBadgeText:"Belum terhubung ke Master Karyawan (Lembur mewajibkan pilih karyawan terdaftar)",employees:a,initialEmployeeId:m?.employee_id||null,initialEmployeeName:m?.employee_name||""});pe({title:i?"\u270F\uFE0F Edit Data Lembur Karyawan":"\u2795 Catat Penugasan Lembur Baru",content:v,size:"md",confirmText:i?"Simpan Perubahan":"Catat Lembur",confirmClass:"btn-primary",onConfirm:async(ne,Z)=>{let G=v.querySelector("#ot-modal-date").value,L=v.querySelector("#ot-modal-branch").value,j=Q.getValue(),V=j.employee_id,te=v.querySelector("#ot-modal-start").value,re=v.querySelector("#ot-modal-end").value,le=parseFloat(v.querySelector("#ot-modal-break").value)||0,ce=parseFloat(v.querySelector("#ot-modal-total").value)||0,se=v.querySelector("#ot-modal-reason").value;if(!G){z("Tanggal lembur wajib diisi");return}if(!L){z("Cabang penugasan wajib dipilih");return}if(!j.employee_name){z("Nama karyawan wajib diisi"),Q.inputEl?.focus();return}if(!V){z("Nama manual belum terhubung ke Master Karyawan. Silakan pilih karyawan terdaftar pada daftar pencarian untuk mencatat lembur."),Q.inputEl?.focus();return}if(!te||!re){z("Jam mulai dan jam selesai wajib diisi");return}let he=n.find(oe=>String(oe.value)===String(L)),Ke=j.employee_name||"Karyawan",$e={date:G,branch_id:Number(L),branch_name:he?he.label:"Cabang Lembur",employee_id:Number(V),employee_name:Ke,start_time:te,end_time:re,break_hours:le,total_hours:ce,reason:se,...i?{row_version:m?.row_version!==void 0&&m?.row_version!==null?Number(m.row_version):1}:{}};try{if(i){let oe=await X(`/api/overtime/${m.id}`,{method:"PUT",body:JSON.stringify($e)});if(oe.ok)ae("Data lembur berhasil diperbarui"),Z(),await x({refreshStats:!0});else{if(oe.status===409){z("Konflik OCC: Data telah diubah oleh pengguna lain. Silakan muat ulang halaman."),Z(),await x({refreshStats:!0});return}z(oe.data?.error||"Gagal memperbarui data lembur")}}else{let oe=await X("/api/overtime",{method:"POST",body:JSON.stringify($e)});oe.ok?(ae("Data lembur berhasil ditambahkan"),c=1,Z(),await x({refreshStats:!0})):z(oe.data?.error||"Gagal menambahkan lembur")}}catch(oe){z("Terjadi kesalahan: "+oe.message)}}})}function f(){return t}function g(m){return`
          ${m.length===0?`
            <div class="lembur-empty-state">
              <div style="font-size:3rem; margin-bottom:12px;">\u{1F50D}</div>
              <h3 style="font-size:1.2rem; font-weight:700; color:#1E293B; margin:0 0 6px 0;">Tidak Ada Data Lembur Ditemukan</h3>
              <p style="color:#64748B; font-size:0.9rem; margin:0 0 20px 0;">Silakan sesuaikan kata kunci pencarian atau reset filter aktif.</p>
              <button id="lembur-btn-reset-filter" class="btn btn-ghost" style="border:1px solid #CBD5E1; border-radius:8px; padding:8px 16px;">
                Reset Filter
              </button>
            </div>
          `:p==="table"?`
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
                      ${N||w?'<th style="width:90px; text-align:center;">Aksi</th>':""}
                    </tr>
                  </thead>
                  <tbody>
                    ${m.map((i,r)=>{let d=(i.employee_name||"Karyawan").split(" ").map(y=>y[0]).filter(Boolean).slice(0,2).join("").toUpperCase(),b=Ra(i.reason);return`
                        <tr data-id="${i.id}">
                          <td style="text-align:center; color:#94A3B8; font-weight:600; font-size:0.8rem;">${(c-1)*E+r+1}</td>
                          <td>
                            <div style="display:flex; align-items:center; gap:10px;">
                              <div class="lembur-avatar-box" style="width:34px; height:34px; font-size:0.78rem; border-radius:8px;">
                                ${d}
                              </div>
                              <div style="font-weight:700; color:#0F172A; white-space:nowrap;">
                                ${i.employee_name}
                              </div>
                            </div>
                          </td>
                          <td style="white-space:nowrap;">
                            <div style="font-weight:600; color:#334155; font-size:0.85rem; display:inline-flex; align-items:center; gap:4px;">
                              <span>\u{1F4CD}</span>
                              <span>${i.branch_name||"Cabang"}</span>
                            </div>
                          </td>
                          <td style="white-space:nowrap; font-weight:600; color:#1E293B;">
                            \u{1F4C5} ${ja(i.date)}
                          </td>
                          <td style="white-space:nowrap;">
                            <span style="font-weight:700; color:#2563EB;">\u23F0 ${i.start_time||"00:00"} - ${i.end_time||"00:00"}</span>
                          </td>
                          <td style="text-align:center; white-space:nowrap;">
                            <span class="lembur-hours-badge" style="padding:3px 8px; font-size:0.8rem;">
                              <span>\u23F1\uFE0F</span>
                              <span>${i.total_hours!=null?i.total_hours:0} Jam</span>
                            </span>
                          </td>
                          <td style="text-align:center; white-space:nowrap; color:#64748B; font-weight:600; font-size:0.82rem;">
                            \u2615 ${i.break_hours!=null?i.break_hours:0} Jam
                          </td>
                          <td style="white-space:nowrap;">
                            <span class="lembur-reason-badge" style="background:${b.bg}; color:${b.color}; border:1px solid ${b.border};">
                              <span>${b.icon}</span>
                              <span>${i.reason||"Lembur Operasional"}</span>
                            </span>
                          </td>
                          ${N||w?`
                            <td style="text-align:center; white-space:nowrap;">
                              <div style="display:inline-flex; align-items:center; gap:6px;">
                                ${N?`
                                  <button class="lembur-icon-btn lembur-edit-btn" title="Edit Catatan Lembur" data-id="${i.id}">
                                    <svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>
                                  </button>
                                `:""}
                                ${w?`
                                  <button class="lembur-icon-btn delete lembur-delete-btn" title="Hapus Catatan Lembur" data-id="${i.id}">
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
              ${m.map(i=>{let r=(i.employee_name||"Karyawan").split(" ").map(b=>b[0]).filter(Boolean).slice(0,2).join("").toUpperCase(),d=Ra(i.reason);return`
                  <div class="lembur-card" data-id="${i.id}">
                    <div>
                      <!-- Card Header: Jam & Istirahat -->
                      <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:12px;">
                        <span class="lembur-hours-badge">
                          <span>\u23F1\uFE0F</span>
                          <span>${i.total_hours!=null?i.total_hours:0} Jam Lembur</span>
                        </span>
                        <span style="font-size:0.8rem; font-weight:600; color:#64748B;">
                          \u2615 Istirahat: ${i.break_hours!=null?i.break_hours:0} Jam
                        </span>
                      </div>

                      <!-- Employee & Branch -->
                      <div style="display:flex; align-items:center; gap:12px; margin-bottom:12px;">
                        <div class="lembur-avatar-box">
                          ${r}
                        </div>
                        <div style="min-width:0;">
                          <h4 style="margin:0 0 3px 0; font-size:1.02rem; font-weight:700; color:#0F172A; white-space:nowrap; overflow:hidden; text-overflow:ellipsis;">
                            ${i.employee_name}
                          </h4>
                          <div style="font-size:0.8rem; color:#64748B; display:flex; align-items:center; gap:4px;">
                            <span>\u{1F4CD}</span>
                            <span style="font-weight:600; color:#334155;">${i.branch_name||"Cabang"}</span>
                          </div>
                        </div>
                      </div>

                      <!-- Shift Times & Date Box -->
                      <div class="lembur-time-box">
                        <div>
                          <div style="font-size:0.7rem; text-transform:uppercase; font-weight:700; color:#94A3B8; margin-bottom:2px;">Tanggal</div>
                          <div style="font-size:0.88rem; font-weight:700; color:#1E293B;">
                            \u{1F5D3}\uFE0F ${ja(i.date)}
                          </div>
                        </div>

                        <div style="text-align:right;">
                          <div style="font-size:0.7rem; text-transform:uppercase; font-weight:700; color:#94A3B8; margin-bottom:2px;">Waktu Shift</div>
                          <div style="font-size:0.88rem; font-weight:700; color:#2563EB;">
                            \u23F0 ${i.start_time||"00:00"} - ${i.end_time||"00:00"}
                          </div>
                        </div>
                      </div>

                      <!-- Reason Badge -->
                      <div>
                        <span class="lembur-reason-badge" style="background:${d.bg}; color:${d.color}; border:1px solid ${d.border};">
                          <span>${d.icon}</span>
                          <span>${i.reason||"Lembur Operasional"}</span>
                        </span>
                      </div>
                    </div>

                    <!-- Card Footer: Aksi -->
                    ${N||w?`
                      <div style="display:flex; justify-content:flex-end; align-items:center; gap:6px; padding-top:12px; border-top:1px solid #29364B; margin-top:12px;">
                        ${N?`
                          <button class="lembur-icon-btn lembur-edit-btn" title="Edit Catatan Lembur" data-id="${i.id}">
                            <svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>
                          </button>
                        `:""}
                        ${w?`
                          <button class="lembur-icon-btn delete lembur-delete-btn" title="Hapus Catatan Lembur" data-id="${i.id}">
                            <svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><polyline points="3 6 5 6 21 6"/><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"/></svg>
                          </button>
                        `:""}
                      </div>
                    `:""}

                  </div>
                `}).join("")}
            </div>
          `}
    `}function D(){let m=e.querySelector("#lembur-btn-reset-filter");m&&(m.onclick=()=>{l="",s="all",o="all",c=1,clearTimeout(F);let i=e.querySelector("#lembur-search-input");i&&(i.value="");let r=e.querySelector("#lembur-filter-branch");r&&(r.value="all"),e.querySelectorAll(".lembur-pill").forEach(d=>{d.classList.toggle("active",d.getAttribute("data-reason")==="all")}),x({refreshStats:!1})}),e.querySelectorAll(".lembur-edit-btn").forEach(i=>{i.onclick=r=>{r.stopPropagation();let d=Number(i.getAttribute("data-id")),b=t.find(y=>y.id===d);b&&P(b)}}),e.querySelectorAll(".lembur-delete-btn").forEach(i=>{i.onclick=r=>{r.stopPropagation();let d=Number(i.getAttribute("data-id"));B(d)}})}function O(){let m=e.querySelector("#lembur-pagination-container");if(!m)return;m.innerHTML="";let i=_e({page:c,pages:T,total:C,limit:E,onPage:r=>{c=r,x({refreshStats:!1})}});i&&m.appendChild(i)}function W(){let m=e.querySelector("#lembur-list-container");if(!m)return;let i=f();m.innerHTML=g(i),O(),D()}function J(){let m=f(),i=Number(u.totalRecords??0),r=Number(u.totalHours??0),d=Number.isInteger(r)?r:r.toFixed(1),b=u.avgDuration!=null?Number(u.avgDuration).toFixed(1):"0.0",y=Number(u.currentMonth??0);e.innerHTML=`
      <style>.lembur-page{display:flex;flex-direction:column;gap:20px}.lembur-hero-card{background:linear-gradient(135deg,#0F172A 0%,#1E293B 100%);border-radius:16px;padding:24px 28px;color:white;box-shadow:0 10px 25px -5px rgba(15,23,42,0.25);position:relative;overflow:hidden}.lembur-hero-card::after{content:'';position:absolute;top:-40px;right:-40px;width:220px;height:220px;background:radial-gradient(circle,rgba(99,102,241,0.18) 0%,transparent 70%);border-radius:50%;pointer-events:none}.lembur-hero-actions{display:flex;align-items:center;gap:8px;flex-wrap:wrap}.lembur-hero-actions .btn{height:38px;box-sizing:border-box;transition:all 0.2s ease}.lembur-btn-add:hover{background:#1D4ED8 !important;transform:translateY(-1px)}.lembur-btn-excel:hover,.lembur-btn-refresh:hover{background:rgba(255,255,255,0.16) !important;border-color:rgba(255,255,255,0.3) !important;transform:translateY(-1px)}.lembur-stats-row{display:grid;grid-template-columns:repeat(4,1fr);gap:14px;margin-top:22px}.lembur-stat-box{background:rgba(255,255,255,0.06);border:1px solid rgba(255,255,255,0.1);border-radius:12px;padding:14px 16px;backdrop-filter:blur(8px);display:flex;align-items:center;gap:14px}.lembur-stat-icon{width:44px;height:44px;border-radius:10px;display:flex;align-items:center;justify-content:center;font-size:1.3rem;flex-shrink:0}.lembur-stat-val{font-size:1.55rem;font-weight:800;line-height:1.1;color:#F8FAFC}.lembur-stat-lbl{font-size:0.76rem;text-transform:uppercase;letter-spacing:0.05em;color:#94A3B8;font-weight:600;margin-top:2px}.lembur-controls-card{background:#172033;border:1px solid #29364B;border-radius:14px;padding:18px 20px;box-shadow:0 2px 4px rgba(0,0,0,0.02);display:flex;flex-direction:column;gap:16px}.lembur-search-row{display:flex;gap:12px;flex-wrap:wrap}.lembur-search-wrap{position:relative;flex:1;min-width:240px}.lembur-search-icon{position:absolute;left:14px;top:50%;transform:translateY(-50%);color:#94A3B8;pointer-events:none}.lembur-search-input{width:100%;padding:10px 14px 10px 40px;border:1px solid #29364B;border-radius:10px;font-size:0.9rem;outline:none;transition:border-color 0.15s ease}.lembur-search-input:focus{border-color:#2563EB;box-shadow:0 0 0 3px rgba(37,99,235,0.1)}.lembur-select{padding:10px 14px;border:1px solid #29364B;border-radius:10px;font-size:0.9rem;color:#F8FAFC;background:#172033;outline:none;min-width:200px;cursor:pointer}.lembur-select:focus{border-color:#2563EB}.lembur-filter-pills{display:flex;align-items:center;gap:8px;overflow-x:auto;white-space:nowrap;scrollbar-width:none;padding-bottom:2px}.lembur-filter-pills::-webkit-scrollbar{display:none}.lembur-pill{padding:6px 14px;border-radius:20px;font-size:0.82rem;font-weight:600;cursor:pointer;border:1px solid #29364B;background:#111827;color:#64748B;transition:all 0.15s ease;display:inline-flex;align-items:center;gap:6px;user-select:none;flex-shrink:0}.lembur-pill:hover{background:#22314A;color:#F8FAFC}.lembur-pill.active{background:#2563EB;color:white;border-color:#2563EB;box-shadow:0 2px 6px rgba(37,99,235,0.25)}.lembur-view-toggle{display:inline-flex;align-items:center;background:#111827;padding:3px;border-radius:10px;border:1px solid #29364B;gap:3px;flex-shrink:0}.lembur-view-btn{display:inline-flex;align-items:center;gap:6px;padding:6px 12px;border-radius:7px;font-size:0.82rem;font-weight:600;border:none;background:transparent;color:#64748B;cursor:pointer;transition:all 0.15s ease;height:34px;box-sizing:border-box;user-select:none}.lembur-view-btn:hover{color:#F8FAFC;background:rgba(255,255,255,0.5)}.lembur-view-btn.active{background:#172033;color:#2563EB;font-weight:700;box-shadow:0 1px 4px rgba(0,0,0,0.08)}.lembur-table-card{background:#172033;border:1px solid #29364B;border-radius:14px;box-shadow:0 2px 6px rgba(0,0,0,0.03);overflow:hidden}.lembur-table-responsive{width:100%;overflow-x:auto;-webkit-overflow-scrolling:touch}.lembur-table{width:100%;border-collapse:collapse;text-align:left;font-size:0.88rem;min-width:860px}.lembur-table th{background:#111827;color:#475569;font-size:0.76rem;font-weight:700;text-transform:uppercase;letter-spacing:0.05em;padding:13px 16px;border-bottom:1.5px solid #E2E8F0;white-space:nowrap}.lembur-table td{padding:13px 16px;border-bottom:1px solid #29364B;color:#F8FAFC;vertical-align:middle}.lembur-table tbody tr:hover td{background:#111827}.lembur-table tbody tr:last-child td{border-bottom:none}.lembur-cards-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(320px,1fr));gap:16px}.lembur-card{background:#172033;border:1px solid #29364B;border-radius:14px;padding:18px;box-shadow:0 2px 6px rgba(0,0,0,0.03);transition:transform 0.15s ease,box-shadow 0.15s ease;display:flex;flex-direction:column;justify-content:space-between;position:relative}.lembur-card:hover{transform:translateY(-2px);box-shadow:0 8px 18px rgba(0,0,0,0.06);border-color:#CBD5E1}.lembur-avatar-box{width:44px;height:44px;border-radius:10px;background:linear-gradient(135deg,#4F46E5 0%,#3730A3 100%);color:white;font-weight:700;font-size:1rem;display:flex;align-items:center;justify-content:center;flex-shrink:0;box-shadow:0 3px 6px rgba(79,70,229,0.25)}.lembur-time-box{background:#111827;border:1px solid #29364B;border-radius:10px;padding:10px 12px;margin:12px 0;display:flex;align-items:center;justify-content:space-between}.lembur-hours-badge{display:inline-flex;align-items:center;gap:5px;padding:5px 12px;border-radius:8px;font-size:0.88rem;font-weight:800;background:#EEF2FF;color:#4F46E5;border:1px solid #C7D2FE}.lembur-reason-badge{display:inline-flex;align-items:center;gap:5px;padding:4px 10px;border-radius:20px;font-size:0.76rem;font-weight:700}.lembur-icon-btn{width:32px;height:32px;border-radius:8px;border:1px solid #29364B;background:#172033;color:#64748B;display:inline-flex;align-items:center;justify-content:center;cursor:pointer;transition:all 0.15s ease}.lembur-icon-btn:hover{background:#22314A;color:#F8FAFC;border-color:#CBD5E1}.lembur-icon-btn.delete:hover{background:#FEE2E2;color:#DC2626;border-color:#FCA5A5}.lembur-empty-state{grid-column:1 / -1;background:#172033;border:2px dashed #E2E8F0;border-radius:16px;padding:48px 24px;text-align:center}@media (max-width:768px){.lembur-hero-card{padding:18px 16px}.lembur-stats-row{grid-template-columns:1fr 1fr;gap:10px}.lembur-stat-box{padding:10px 12px}.lembur-stat-icon{width:38px;height:38px;font-size:1.1rem}.lembur-stat-val{font-size:1.3rem}.lembur-hero-actions{width:100% !important;display:flex !important;align-items:center !important;gap:6px !important;flex-wrap:nowrap !important;overflow-x:auto !important;scrollbar-width:none !important;padding-bottom:2px !important}.lembur-hero-actions::-webkit-scrollbar{display:none}.lembur-btn-add{flex:1 1 auto !important;justify-content:center !important;padding:8px 10px !important;font-size:0.82rem !important;white-space:nowrap !important}.lembur-btn-excel{flex:1 1 auto !important;justify-content:center !important;padding:8px 10px !important;font-size:0.82rem !important;white-space:nowrap !important}.lembur-btn-refresh{width:38px !important;height:38px !important;padding:0 !important;flex-shrink:0 !important;justify-content:center !important}.lembur-controls-card{padding:14px;gap:12px}.lembur-search-row{flex-direction:column;gap:10px}.lembur-select{width:100%}.lembur-view-toggle{width:100%;display:flex}.lembur-view-btn{flex:1;justify-content:center}.lembur-table-card{border-radius:12px}.lembur-cards-grid{grid-template-columns:minmax(0,1fr) !important;width:100% !important;max-width:100% !important;min-width:0 !important;box-sizing:border-box !important}.lembur-card{width:100% !important;max-width:100% !important;min-width:0 !important;box-sizing:border-box !important;overflow:hidden;padding:14px !important}}</style>

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
              ${I?`
                <button id="lembur-btn-add" class="btn btn-primary lembur-btn-add" style="background:#2563EB; height:38px; padding:0 16px; border-radius:10px; font-weight:700; display:inline-flex; align-items:center; gap:6px; box-shadow:0 4px 12px rgba(37,99,235,0.3); white-space:nowrap;">
                  <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24"><line x1="12" y1="5" x2="12" y2="19"/><line x1="5" y1="12" x2="19" y2="12"/></svg>
                  <span>Tambah Lembur</span>
                </button>
              `:""}
              
              ${ie(K,"export")?`<button id="lembur-btn-excel" class="btn btn-ghost lembur-btn-excel" style="background:rgba(255,255,255,0.08); border:1px solid rgba(255,255,255,0.15); color:white; height:38px; padding:0 14px; border-radius:10px; font-weight:600; display:inline-flex; align-items:center; gap:6px; white-space:nowrap;">
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
                <div class="lembur-stat-val"><span id="lembur-stat-hours">${d}</span> Jam</div>
                <div class="lembur-stat-lbl">Total Jam Lembur</div>
              </div>
            </div>

            <div class="lembur-stat-box">
              <div class="lembur-stat-icon" style="background:rgba(59,130,246,0.2); color:#60A5FA;">
                \u{1F4DD}
              </div>
              <div>
                <div class="lembur-stat-val" id="lembur-stat-records">${i}</div>
                <div class="lembur-stat-lbl">Catatan Penugasan</div>
              </div>
            </div>

            <div class="lembur-stat-box">
              <div class="lembur-stat-icon" style="background:rgba(16,185,129,0.2); color:#34D399;">
                \u{1F4CA}
              </div>
              <div>
                <div class="lembur-stat-val"><span id="lembur-stat-avg">${b}</span> Jam</div>
                <div class="lembur-stat-lbl">Rata-rata Durasi</div>
              </div>
            </div>

            <div class="lembur-stat-box">
              <div class="lembur-stat-icon" style="background:rgba(245,158,11,0.2); color:#FBBF24;">
                \u{1F5D3}\uFE0F
              </div>
              <div>
                <div class="lembur-stat-val" id="lembur-stat-month">${y}</div>
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
              <input type="text" id="lembur-search-input" class="lembur-search-input" placeholder="Cari nama karyawan, cabang, atau alasan lembur..." value="${l?String(l).replace(/"/g,"&quot;"):""}" />
            </div>

            <select id="lembur-filter-branch" class="lembur-select">
              <option value="all">Semua Cabang (${n.length})</option>
              ${n.map(q=>`<option value="${q.value}" ${String(s)===String(q.value)?"selected":""}>${q.label}</option>`).join("")}
            </select>

            <div class="lembur-view-toggle">
              <button type="button" class="lembur-view-btn ${p==="card"?"active":""}" data-view="card" title="Tampilan Kartu">
                <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><rect x="3" y="3" width="7" height="7" rx="1"/><rect x="14" y="3" width="7" height="7" rx="1"/><rect x="14" y="14" width="7" height="7" rx="1"/><rect x="3" y="14" width="7" height="7" rx="1"/></svg>
                <span>Kartu</span>
              </button>
              <button type="button" class="lembur-view-btn ${p==="table"?"active":""}" data-view="table" title="Tampilan Tabel">
                <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M3 9h18M3 15h18M9 3v18"/><rect x="3" y="3" width="18" height="18" rx="2"/></svg>
                <span>Tabel</span>
              </button>
            </div>
          </div>

          <!-- Reason Filter Pills (Swipeable on Mobile) -->
          <div class="lembur-filter-pills">
            <div class="lembur-pill ${o==="all"?"active":""}" data-reason="all">
              <span>Semua Lembur</span>
              <span class="badge" id="lembur-pill-count-all" style="background:rgba(0,0,0,0.06); padding:2px 6px; border-radius:10px; font-size:0.75rem;">${i}</span>
            </div>
            ${mi.map(q=>{let v=Ra(q);return`
                <div class="lembur-pill ${o===q?"active":""}" data-reason="${q}">
                  <span>${v.icon}</span>
                  <span>${q}</span>
                </div>
              `}).join("")}
          </div>
        </div>

        <!-- Content Area: Murni Grid Kartu -->
        <div id="lembur-list-container">
          ${g(m)}
        </div>

        <!-- Pagination Container -->
        <div id="lembur-pagination-container"></div>

      </div>
    `,O(),$()}function $(){let m=e.querySelector("#lembur-btn-add");m&&(m.onclick=()=>P());let i=e.querySelector("#lembur-btn-excel");i&&(i.onclick=()=>_());let r=e.querySelector("#lembur-btn-refresh");r&&(r.onclick=()=>x({refreshStats:!0}));let d=e.querySelector("#lembur-search-input");d&&d.addEventListener("input",y=>{l=y.target.value,clearTimeout(F),F=setTimeout(()=>{c=1,x({refreshStats:!1})},350)});let b=e.querySelector("#lembur-filter-branch");b&&b.addEventListener("change",y=>{s=y.target.value,c=1,x({refreshStats:!1})}),e.querySelectorAll(".lembur-view-btn").forEach(y=>{y.addEventListener("click",()=>{let q=y.getAttribute("data-view");!q||q===p||(p=q,localStorage.setItem("fcms_overtime_view_mode",p),e.querySelectorAll(".lembur-view-btn").forEach(v=>{v.classList.toggle("active",v===y)}),W())})}),e.querySelectorAll(".lembur-pill").forEach(y=>{y.addEventListener("click",()=>{o=y.getAttribute("data-reason")||"all",e.querySelectorAll(".lembur-pill").forEach(q=>{q.classList.toggle("active",q===y)}),c=1,x({refreshStats:!1})})}),D()}await x({refreshStats:!0})}Ne();we();async function hi(e){let t=de();if(!t||t.role!=="superadmin"){e.innerHTML='<div class="empty-state"><p class="text-danger">Akses ditolak.</p></div>';return}De({container:e,title:"Riwayat Aktivitas",icon:"\u{1F6E1}\uFE0F",apiPath:"/api/audit-logs",enableMobileFilterSheet:!0,itemLabel:"Log",canCreate:!1,canEdit:!1,canDelete:!1,bulkDelete:!1,exportOptions:{moduleName:"audit_logs",filename:"FCMS_Audit_Logs",template:!1,import:!1,onExport:async a=>{let s=(await ue("/api/audit-logs/export",a)||[]).map(l=>({ID:l.id,Waktu:l.created_at||"-",Pengguna:l.user_name||"Sistem",Role:l.user_role||"-",Aksi:l.action||"-",Modul:l.module||"-","ID Target":l.target_id||"-"}));ge(s,"FCMS_Audit_Logs","Audit Logs")}},columns:[{key:"created_at",label:"Waktu",nowrap:!0,render:a=>{let o=new Date(a);return Number.isNaN(o.getTime())?Y(a||"-"):o.toLocaleString("id-ID",{dateStyle:"short",timeStyle:"medium"})}},{key:"user_name",label:"Pengguna",render:(a,o)=>`<strong>${Y(a||"Sistem")}</strong><br><small class="text-muted" style="text-transform:capitalize">${Y(o.user_role||"")}</small>`},{key:"action",label:"Aksi",render:a=>`<span class="badge ${{CREATE:"badge-success",UPDATE:"badge-info",DELETE:"badge-danger",LOGIN:"badge-primary",LOGOUT:"badge-neutral"}[a]||"badge-neutral"}">${Y(a)}</span>`},{key:"module",label:"Modul",render:a=>`<span style="text-transform:capitalize">${Y((a||"").replace(/_/g," "))}</span>`},{key:"target_id",label:"ID Target"},{key:"id",label:"Detail",render:a=>{let o=document.createElement("button");return o.type="button",o.className="btn btn-xs btn-outline",o.textContent="Lihat Detail",o.addEventListener("click",()=>window.viewAuditDetail?.(a)),o}}],filterFields:[{type:"search",placeholder:"Cari ID, pengguna, modul..."},{type:"select",name:"action",options:[{value:"",label:"Semua Aksi"},{value:"CREATE",label:"Tambah (CREATE)"},{value:"UPDATE",label:"Ubah (UPDATE)"},{value:"DELETE",label:"Hapus (DELETE)"},{value:"LOGIN",label:"Masuk (LOGIN)"},{value:"LOGOUT",label:"Keluar (LOGOUT)"}]},{type:"select",name:"module",options:[{value:"",label:"Semua Modul"},{value:"employees",label:"Karyawan"},{value:"contracts",label:"Kontrak"},{value:"sp",label:"Data SP"},{value:"mutasi",label:"Data Mutasi"},{value:"overtime",label:"Lembur"},{value:"schedule",label:"Time Line / Jadwal"},{value:"os_migrations",label:"Migrasi OS"},{value:"issues",label:"Permasalahan"},{value:"training",label:"Training"},{value:"reports",label:"Laporan Operasional"},{value:"sop",label:"SOP Facility Care & Security"},{value:"checklist",label:"Master Document"},{value:"hygiene_standards",label:"Master Hygiene"},{value:"dac",label:"Master DAC"},{value:"users",label:"Manajemen User"},{value:"branches",label:"Manajemen Cabang"},{value:"auth",label:"Autentikasi"}]}]});let n=e.querySelector(".page-actions");if(n&&t&&t.role==="superadmin"){let a=document.createElement("a");a.href="#/sync",a.id="btn-audit-sync-health",a.className="btn btn-primary",a.style.cssText="text-decoration:none;display:inline-flex;align-items:center;gap:6px",a.innerHTML="\u{1FA7A} Sync & Health Dashboard",n.appendChild(a)}window.viewAuditDetail=async a=>{try{let o=await X(`/api/audit-logs/${encodeURIComponent(a)}`),s=o?.data?.data||null;if(!o?.ok||!s)return alert("Data tidak ditemukan");let l=E=>String(E).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;"),h=E=>{if(!E)return"Tidak ada data";try{let C=typeof E=="object"?E:JSON.parse(E);return l(JSON.stringify(C,null,2))}catch{return l(String(E))}},p=`
         <div style="display:flex; gap:1rem; flex-wrap:wrap">
           <div style="flex:1; min-width:300px">
              <h4>Data Lama</h4>
              <pre style="background:#f8f9fa; padding:10px; border-radius:5px; font-size:12px; overflow-x:auto; border:1px solid #ddd; max-height:400px; overflow-y:auto;">${h(s.old_data)}</pre>
           </div>
           <div style="flex:1; min-width:300px">
              <h4>Data Baru</h4>
              <pre style="background:#f8f9fa; padding:10px; border-radius:5px; font-size:12px; overflow-x:auto; border:1px solid #ddd; max-height:400px; overflow-y:auto;">${h(s.new_data)}</pre>
           </div>
         </div>
       `,{createModal:c}=await Promise.resolve().then(()=>(Fe(),nn));c({title:`Detail Audit Log #${a}`,content:p,width:"800px",hideFooter:!0})}catch{alert("Gagal mengambil detail")}}}Ne();Fe();we();var Pr=/^(\d{4})-(\d{2})-(\d{2})(?:[T\s].*)?$/,Lr=/^(\d{2})[\/\-](\d{2})[\/\-](\d{4})$/,Dr=/^(\d{4})-(\d{2})$/,Ir=/^\d{5}$/;function qa(e,t,n){if(t<1||t>12||n<1||n>31)return null;let a=new Date(Date.UTC(e,t-1,n));return a.getUTCFullYear()!==e||a.getUTCMonth()!==t-1||a.getUTCDate()!==n?null:a.getTime()}function fi(e,{allowYearMonth:t=!0}={}){if(e==null)return null;let n=String(e).trim();if(!n||n==="-")return null;let a=n.match(Pr);if(a)return qa(Number(a[1]),Number(a[2]),Number(a[3]));if(a=n.match(Lr),a)return qa(Number(a[3]),Number(a[2]),Number(a[1]));if(Ir.test(n)){let o=Number(n);return o>=1e4?Date.UTC(1899,11,30)+o*864e5:null}return t&&(a=n.match(Dr),a)?qa(Number(a[1]),Number(a[2]),1):null}function xi(e){return[...e].sort((t,n)=>{let a=fi(t,{allowYearMonth:!0}),o=fi(n,{allowYearMonth:!0});return a===null&&o!==null?1:a!==null&&o===null?-1:a!==null&&o!==null&&o!==a?o-a:String(n||"").localeCompare(String(t||""))})}var Oa={"2026-09":"September 2026","2026-10":"Oktober 2026","2026-11":"November 2026","2026-12":"Desember 2026"},Mt={scheduled:{label:"Terjadwal",badgeClass:"badge-warning",icon:"\u{1F7E1}",bg:"#FEF9C3",color:"#854D0E",border:"#FDE047"},in_progress:{label:"Sedang Proses",badgeClass:"badge-info",icon:"\u{1F535}",bg:"#DBEAFE",color:"#1E40AF",border:"#93C5FD"},completed:{label:"Selesai",badgeClass:"badge-success",icon:"\u{1F7E2}",bg:"#DCFCE7",color:"#166534",border:"#86EFAC"}};async function yi(e){let t=de(),n=ie(t,"os_migrations","write"),a=ie(t,"os_migrations","delete"),o=ie(t,"export");e.innerHTML=`
    <div class="page-header" style="display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:16px; margin-bottom:20px;">
      <div>
        <h1 class="page-title" style="display:flex; align-items:center; gap:8px; margin:0; font-size:1.45rem; font-weight:800; color:var(--text-1); letter-spacing:-0.02em;">
          <span>\u{1F504}</span> Time Line Migrasi OS
        </h1>
      </div>
      <div class="os-header-actions" style="display:flex; gap:10px; align-items:center; flex-wrap:wrap;">
        <div class="btn-group" style="background:#111827; padding:3px; border-radius:10px; display:inline-flex; align-items:center; border:1px solid var(--border); height:38px; box-sizing:border-box;">
          <button id="view-board-btn" class="btn btn-sm" style="background:#2563EB; color:#fff; font-weight:600; border-radius:7px; box-shadow:0 1px 2px rgba(0,0,0,0.05); padding:0 12px; height:100%; display:inline-flex; align-items:center; cursor:pointer;">
            \u{1F5C2}\uFE0F Matrix Board
          </button>
          <button id="view-table-btn" class="btn btn-sm btn-ghost" style="font-weight:600; border-radius:7px; padding:0 12px; height:100%; display:inline-flex; align-items:center; cursor:pointer;">
            \u{1F4CB} Tabel Data
          </button>
        </div>
        ${o?`
        <button id="btn-export-migrations" class="btn btn-outline" style="background:#172033; border:1px solid var(--border); color:#F8FAFC; font-weight:600; height:38px; padding:0 14px; border-radius:10px; display:inline-flex; align-items:center; gap:6px; cursor:pointer; box-sizing:border-box;">
          \u{1F4E5} Export Excel
        </button>
        `:""}
        ${n?`
        <button id="btn-add-migration" class="btn btn-primary" style="font-weight:700; height:38px; padding:0 16px; border-radius:10px; display:inline-flex; align-items:center; gap:6px; cursor:pointer; box-sizing:border-box; background:#2563EB; box-shadow:0 4px 12px rgba(37,99,235,0.3);">
          + Tambah Cabang
        </button>
        `:""}
      </div>
    </div>

    <!-- Summary KPI Cards -->
    <div id="migration-summary-cards" style="display:grid; grid-template-columns:repeat(auto-fit, minmax(220px, 1fr)); gap:16px; margin-bottom:24px;">
      <div class="kpi-card skeleton" style="height:100px; border-radius:12px;"></div>
      <div class="kpi-card skeleton" style="height:100px; border-radius:12px;"></div>
      <div class="kpi-card skeleton" style="height:100px; border-radius:12px;"></div>
      <div class="kpi-card skeleton" style="height:100px; border-radius:12px;"></div>
    </div>

    <!-- Filter Toolbar -->
    <div class="card" style="padding:16px 20px; margin-bottom:24px; background:#172033; border:1px solid var(--border); border-radius:12px; box-shadow:0 1px 3px rgba(0,0,0,0.2);">
      <div style="display:flex; flex-wrap:wrap; gap:14px; align-items:center; justify-content:space-between;">
        <div style="display:flex; flex-wrap:wrap; gap:12px; align-items:center; flex:1; min-width:280px;">
          <div style="position:relative; flex:1; min-width:200px; max-width:320px;">
            <input type="text" id="migration-search" class="form-control" placeholder="Cari cabang, vendor, catatan..." style="padding-left:34px; height:38px; border-radius:8px; font-size:0.9rem; width:100%;">
            <span style="position:absolute; left:11px; top:50%; transform:translateY(-50%); color:var(--text-3); font-size:0.9rem;">\u{1F50D}</span>
          </div>
          <select id="filter-vendor" class="form-control" style="width:auto; height:38px; border-radius:8px; font-size:0.875rem; min-width:140px;">
            <option value="">Semua Vendor</option>
            <option value="AEON">AEON</option>
            <option value="Triple S">Triple S</option>
          </select>
          <select id="filter-month" class="form-control" style="width:auto; height:38px; border-radius:8px; font-size:0.875rem; min-width:150px;">
            <option value="">Semua Bulan</option>
            <option value="2026-09">September 2026</option>
            <option value="2026-10">Oktober 2026</option>
            <option value="2026-11">November 2026</option>
            <option value="2026-12">Desember 2026</option>
          </select>
          <select id="filter-status" class="form-control" style="width:auto; height:38px; border-radius:8px; font-size:0.875rem; min-width:140px;">
            <option value="">Semua Status</option>
            <option value="scheduled">\u{1F7E1} Terjadwal</option>
            <option value="in_progress">\u{1F535} Sedang Proses</option>
            <option value="completed">\u{1F7E2} Selesai</option>
          </select>
        </div>
        <div id="filter-stats-text" style="font-size:0.85rem; color:var(--text-3); font-weight:500;">
          Memuat data...
        </div>
      </div>
    </div>

    <!-- Main Content Container -->
    <div id="migration-view-container"></div>
  `;let s="board",l=null,h=[];X("/api/branches?all=1").then(w=>{w.ok&&w.data?.data&&(h=w.data.data)}).catch(()=>{});async function p(){try{let w=await X("/api/os-migrations/summary");if(!w.ok)throw new Error("Gagal memuat data migrasi");l=w.data?.data||w.data,c(l),C()}catch(w){let M=document.getElementById("migration-view-container");M.innerHTML=`
        <div class="empty-state" style="padding:40px; text-align:center;">
          <div style="font-size:2rem; margin-bottom:8px;">\u26A0\uFE0F</div>
          <p style="color:var(--danger); font-weight:600;">Gagal memuat data: ${Y(w.message)}</p>
          <button type="button" id="os-mig-reload-btn" class="btn btn-outline btn-sm">Muat Ulang</button>
        </div>
      `,M.querySelector("#os-mig-reload-btn")?.addEventListener("click",()=>window.location.reload())}}function c(w){let M=document.getElementById("migration-summary-cards");if(!M||!w)return;let k=w.total_branches||0,x=w.vendors?.AEON?.total||0,_=w.vendors?.["Triple S"]?.total||0,B=w.by_status?.completed||0,P=w.by_status?.in_progress||0,f=w.by_status?.scheduled||0,g=k>0?Math.round(B/k*100):0;M.innerHTML=`
      <div class="card" style="padding:16px 20px; background:#172033; border:1px solid var(--border); border-radius:12px; box-shadow:0 1px 3px rgba(0,0,0,0.2);">
        <div style="font-size:0.8rem; font-weight:600; color:var(--text-3); text-transform:uppercase; letter-spacing:0.5px; margin-bottom:6px;">Total Cabang Migrasi</div>
        <div style="font-size:1.8rem; font-weight:800; color:var(--text-1); line-height:1.2;">${k} <span style="font-size:0.9rem; font-weight:500; color:var(--text-3);">Cabang</span></div>
        <div style="font-size:0.75rem; color:var(--text-3); margin-top:4px;">Timeline Periode Sep - Des 2026</div>
      </div>

      <div class="card" style="padding:16px 20px; background:#172033; border:1px solid var(--border); border-radius:12px; border-left:4px solid #2563EB; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
        <div style="font-size:0.8rem; font-weight:600; color:#1E40AF; text-transform:uppercase; letter-spacing:0.5px; margin-bottom:6px;">Vendor AEON</div>
        <div style="font-size:1.8rem; font-weight:800; color:#1E40AF; line-height:1.2;">${x} <span style="font-size:0.9rem; font-weight:500; color:var(--text-3);">Cabang</span></div>
        <div style="font-size:0.75rem; color:var(--text-3); margin-top:4px;">Sep: 9 \u2022 Nov: 3 \u2022 Des: 8</div>
      </div>

      <div class="card" style="padding:16px 20px; background:#172033; border:1px solid var(--border); border-radius:12px; border-left:4px solid #0D9488; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
        <div style="font-size:0.8rem; font-weight:600; color:#0F766E; text-transform:uppercase; letter-spacing:0.5px; margin-bottom:6px;">Vendor Triple S</div>
        <div style="font-size:1.8rem; font-weight:800; color:#0F766E; line-height:1.2;">${_} <span style="font-size:0.9rem; font-weight:500; color:var(--text-3);">Cabang</span></div>
        <div style="font-size:0.75rem; color:var(--text-3); margin-top:4px;">Sep: 5 \u2022 Okt: 5 \u2022 Nov: 5 \u2022 Des: 27</div>
      </div>

      <div class="card" style="padding:16px 20px; background:#172033; border:1px solid var(--border); border-radius:12px; border-left:4px solid #16A34A; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
        <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:6px;">
          <span style="font-size:0.8rem; font-weight:600; color:#15803D; text-transform:uppercase; letter-spacing:0.5px;">Progres Realisasi</span>
          <span style="font-size:0.85rem; font-weight:700; color:#15803D;">${g}%</span>
        </div>
        <div style="width:100%; height:8px; background:#1B2940; border-radius:4px; overflow:hidden; margin-bottom:8px;">
          <div style="width:${g}%; height:100%; background:#16A34A; border-radius:4px;"></div>
        </div>
        <div style="display:flex; justify-content:space-between; font-size:0.75rem; color:var(--text-2); font-weight:500;">
          <span>\u{1F7E2} Selesai: <strong>${B}</strong></span>
          <span>\u{1F535} Proses: <strong>${P}</strong></span>
          <span>\u{1F7E1} Terjadwal: <strong>${f}</strong></span>
        </div>
      </div>
    `}function E(){if(!l||!l.matrix)return[];let w=(document.getElementById("migration-search")?.value||"").toLowerCase().trim(),M=document.getElementById("filter-vendor")?.value||"",k=document.getElementById("filter-month")?.value||"",x=document.getElementById("filter-status")?.value||"";return l.matrix.filter(_=>!(M&&_.vendor!==M||k&&_.target_month!==k||x&&_.status!==x||w&&!`${_.branch_code} ${_.branch_name} ${_.branch_full_name} ${_.vendor} ${_.notes||""} ${_.pic_vendor||""}`.toLowerCase().includes(w)))}function C(){let w=E(),M=document.getElementById("filter-stats-text");M&&(M.innerHTML=`Menampilkan <strong>${w.length}</strong> dari <strong>${l?.total_branches||0}</strong> cabang`);let k=document.getElementById("migration-view-container");s==="board"?T(k,w):u(k,w)}function T(w,M){if(!M.length){w.innerHTML=`
        <div class="empty-state" style="padding:50px; text-align:center; background:#111827; border-radius:12px; border:1px solid var(--border);">
          <div style="font-size:2.5rem; margin-bottom:12px;">\u{1F50D}</div>
          <p style="font-size:1.05rem; font-weight:600; color:var(--text-2); margin-bottom:6px;">Tidak ada cabang yang cocok dengan filter</p>
          <p style="color:var(--text-3); font-size:0.875rem;">Coba reset kata kunci pencarian atau pilihan filter di atas.</p>
        </div>
      `;return}let k=["AEON","Triple S"],x=xi(Array.from(new Set(M.map(B=>B.target_month).filter(Boolean)))),_="";k.forEach(B=>{let P=M.filter(J=>J.vendor===B);if(document.getElementById("filter-vendor")?.value&&document.getElementById("filter-vendor").value!==B||P.length===0)return;let f=P.length,g=B==="AEON"?"#93C5FD":"#5EEAD4",D=B==="AEON"?"rgba(59, 130, 246, 0.14)":"rgba(20, 184, 166, 0.14)",O=B==="AEON"?"rgba(147, 197, 253, 0.3)":"rgba(94, 234, 212, 0.3)",W=x.filter(J=>P.some($=>$.target_month===J));_+=`
        <div class="vendor-board-section" style="margin-bottom:32px; background:#111827; border:1px solid var(--border); border-radius:14px; overflow:hidden; box-shadow:0 1px 4px rgba(0,0,0,0.04);">
          <!-- Vendor Section Header -->
          <div style="padding:14px 20px; background:${D}; border-bottom:1px solid ${O}; display:flex; justify-content:space-between; align-items:center;">
            <div style="display:flex; align-items:center; gap:10px;">
              <span style="font-size:1.1rem; font-weight:800; color:${g};">${B==="AEON"?"\u{1F3E2} AEON":"\u{1F6E1}\uFE0F Triple S"}</span>
              <span class="badge" style="background:#172033; color:${g}; border:1px solid ${O}; font-weight:700; font-size:0.75rem; padding:3px 8px; border-radius:6px;">
                ${f} Cabang
              </span>
            </div>
            <div style="font-size:0.8rem; color:var(--text-3); font-weight:500;">
              Format Timeline Berdasarkan Dokumen Migrasi
            </div>
          </div>

          <!-- Months Grid (Hanya menampilkan bulan yang memiliki jadwal) -->
          <div style="display:grid; grid-template-columns:repeat(auto-fit, minmax(260px, 1fr)); gap:16px; padding:18px;">
            ${W.map(J=>{let $=P.filter(r=>r.target_month===J);return`
                <div class="month-column" style="background:#172033; border:1px solid var(--border); border-radius:10px; display:flex; flex-direction:column; min-height:220px;">
                  <!-- Month Header -->
                  <div style="padding:12px 14px; background:#111827; border-bottom:1px solid var(--border); border-radius:9px 9px 0 0; display:flex; justify-content:space-between; align-items:center;">
                    <div style="font-weight:700; font-size:0.9rem; color:var(--text-1);">
                      \u{1F4C5} ${(Oa[J]||J).split(" ")[0]}
                    </div>
                    <span class="badge" style="background:#1B2940; color:var(--text-2); font-weight:700; font-size:0.75rem; padding:2px 7px; border-radius:4px;">
                      ${$.length}
                    </span>
                  </div>

                  <!-- Branch Cards Container -->
                  <div style="padding:10px; display:flex; flex-direction:column; gap:8px; flex:1; max-height:480px; overflow-y:auto;">
                    ${$.map(r=>{let d=Mt[r.status]||Mt.scheduled;return`
                        <div class="branch-card" style="background:#172033; border:1px solid var(--border); border-radius:8px; padding:10px 12px; box-shadow:0 1px 2px rgba(0,0,0,0.03); transition:transform 0.15s, box-shadow 0.15s;">
                          <div style="display:flex; justify-content:space-between; align-items:flex-start; gap:6px; margin-bottom:6px;">
                            <div style="font-weight:700; font-size:0.9rem; color:var(--text-1); line-height:1.3;">
                              <span style="color:#2563EB; font-family:monospace; margin-right:4px;">${r.branch_code}.</span>${r.branch_name}
                            </div>
                            ${n||a?`
                              <div style="display:flex; gap:2px;">
                                ${n?`<button class="btn-icon-mini edit-card-btn" data-id="${r.id}" title="Edit" style="background:none; border:none; cursor:pointer; padding:2px 4px; font-size:0.8rem; color:var(--text-3);">\u270F\uFE0F</button>`:""}
                                ${a?`<button class="btn-icon-mini delete-card-btn" data-id="${r.id}" data-name="${r.branch_name}" title="Hapus" style="background:none; border:none; cursor:pointer; padding:2px 4px; font-size:0.8rem; color:var(--danger);">\u{1F5D1}\uFE0F</button>`:""}
                              </div>
                            `:""}
                          </div>

                          <div style="display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:6px; margin-top:8px;">
                            <!-- Status Selector Dropdown -->
                            <div style="position:relative;">
                              <button class="status-toggle-btn" data-id="${r.id}" data-status="${r.status}" style="background:${d.bg}; color:${d.color}; border:1px solid ${d.border}; padding:3px 8px; border-radius:6px; font-size:0.75rem; font-weight:700; display:inline-flex; align-items:center; gap:4px; cursor:${n?"pointer":"default"};" title="${n?"Klik untuk ubah status":"Status"}">
                                ${d.icon} ${d.label}${n?" \u25BE":""}
                              </button>
                            </div>

                            ${r.employee_count>0?`
                              <span style="font-size:0.75rem; color:var(--text-3); font-weight:600; display:inline-flex; align-items:center; gap:3px;" title="${r.employee_count} personil Facility Care aktif di cabang ini">
                                \u{1F464} ${r.employee_count} FC
                              </span>
                            `:""}
                          </div>

                          ${r.actual_date?`
                            <div style="font-size:0.7rem; color:var(--text-3); margin-top:6px;">
                              Realisasi: <strong>${r.actual_date}</strong>
                            </div>
                          `:""}
                        </div>
                      `}).join("")}
                  </div>
                </div>
              `}).join("")}
          </div>
        </div>
      `}),w.innerHTML=_,F(w)}function u(w,M){if(!M.length){w.innerHTML=`
        <div class="empty-state" style="padding:50px; text-align:center; background:#111827; border-radius:12px; border:1px solid var(--border);">
          <div style="font-size:2.5rem; margin-bottom:12px;">\u{1F50D}</div>
          <p style="font-size:1.05rem; font-weight:600; color:var(--text-2); margin-bottom:6px;">Tidak ada cabang yang cocok dengan filter</p>
        </div>
      `;return}w.innerHTML=`
      <div class="card" style="background:#172033; border:1px solid var(--border); border-radius:12px; overflow:hidden; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
        <div class="table-responsive">
          <table class="table" style="width:100%; border-collapse:collapse;">
            <thead>
              <tr style="background:#111827; border-bottom:1px solid var(--border); text-align:left;">
                <th style="padding:12px 16px; font-size:0.8rem; font-weight:700; color:var(--text-3); width:50px;">NO</th>
                <th style="padding:12px 16px; font-size:0.8rem; font-weight:700; color:var(--text-3);">KODE & CABANG</th>
                <th style="padding:12px 16px; font-size:0.8rem; font-weight:700; color:var(--text-3);">VENDOR</th>
                <th style="padding:12px 16px; font-size:0.8rem; font-weight:700; color:var(--text-3);">TARGET BULAN</th>
                <th style="padding:12px 16px; font-size:0.8rem; font-weight:700; color:var(--text-3);">STATUS</th>
                <th style="padding:12px 16px; font-size:0.8rem; font-weight:700; color:var(--text-3);">REALISASI</th>
                <th style="padding:12px 16px; font-size:0.8rem; font-weight:700; color:var(--text-3);">SDM FC</th>
                <th style="padding:12px 16px; font-size:0.8rem; font-weight:700; color:var(--text-3);">CATATAN</th>
                ${n||a?'<th style="padding:12px 16px; font-size:0.8rem; font-weight:700; color:var(--text-3); text-align:right;">AKSI</th>':""}
              </tr>
            </thead>
            <tbody>
              ${M.map((k,x)=>{let _=Mt[k.status]||Mt.scheduled;return`
                  <tr style="border-bottom:1px solid var(--border);">
                    <td style="padding:12px 16px; font-size:0.85rem; color:var(--text-3);">${x+1}</td>
                    <td style="padding:12px 16px;">
                      <div style="font-weight:700; font-size:0.9rem; color:var(--text-1);">${k.branch_full_name}</div>
                    </td>
                    <td style="padding:12px 16px;">
                      <span class="badge" style="background:${k.vendor==="AEON"?"#DBEAFE":"#CCFBF1"}; color:${k.vendor==="AEON"?"#1E40AF":"#0F766E"}; font-weight:700; border-radius:6px; padding:4px 8px; font-size:0.75rem;">
                        ${k.vendor}
                      </span>
                    </td>
                    <td style="padding:12px 16px; font-size:0.875rem; font-weight:600; color:var(--text-2);">
                      ${Oa[k.target_month]||k.target_month}
                    </td>
                    <td style="padding:12px 16px;">
                      <button class="status-toggle-btn" data-id="${k.id}" data-status="${k.status}" style="background:${_.bg}; color:${_.color}; border:1px solid ${_.border}; padding:4px 8px; border-radius:6px; font-size:0.75rem; font-weight:700; cursor:${n?"pointer":"default"};" title="${n?"Klik untuk ubah status":"Status"}">
                        ${_.icon} ${_.label}${n?" \u25BE":""}
                      </button>
                    </td>
                    <td style="padding:12px 16px; font-size:0.85rem; color:var(--text-2);">
                      ${k.actual_date||"-"}
                    </td>
                    <td style="padding:12px 16px; font-size:0.85rem; font-weight:600; color:var(--text-2);">
                      ${k.employee_count>0?`\u{1F464} ${k.employee_count}`:"-"}
                    </td>
                    <td style="padding:12px 16px; font-size:0.85rem; color:var(--text-3); max-width:200px; white-space:nowrap; overflow:hidden; text-overflow:ellipsis;">
                      ${k.notes||"-"}
                    </td>
                    ${n||a?`
                      <td style="padding:12px 16px; text-align:right;">
                        ${n?`<button class="btn btn-outline btn-sm edit-card-btn" data-id="${k.id}" style="padding:3px 8px; font-size:0.75rem; border-radius:6px; margin-right:4px;">Edit</button>`:""}
                        ${a?`<button class="btn btn-danger btn-sm delete-card-btn" data-id="${k.id}" data-name="${k.branch_name}" style="padding:3px 8px; font-size:0.75rem; border-radius:6px;">Hapus</button>`:""}
                      </td>
                    `:""}
                  </tr>
                `}).join("")}
            </tbody>
          </table>
        </div>
      </div>
    `,F(w)}function F(w){w.querySelectorAll(".status-toggle-btn").forEach(M=>{M.addEventListener("click",k=>{if(k.stopPropagation(),!n)return;let x=M.dataset.id,_=M.dataset.status;R(M,x,_)})}),w.querySelectorAll(".edit-card-btn").forEach(M=>{M.addEventListener("click",()=>{let k=M.dataset.id,x=l?.matrix?.find(_=>String(_.id)===String(k));x&&K(x)})}),w.querySelectorAll(".delete-card-btn").forEach(M=>{M.addEventListener("click",()=>{if(!a){z("Hanya Superadmin yang dapat menghapus jadwal migrasi.");return}let k=M.dataset.id,x=M.dataset.name||"Cabang";me(`Apakah Anda yakin ingin menghapus jadwal migrasi untuk <strong>${x}</strong>?`,async()=>{try{if(!(await X(`/api/os-migrations/${k}`,{method:"DELETE"})).ok)throw new Error("Gagal menghapus");ae(`Jadwal migrasi ${x} berhasil dihapus`),p()}catch(_){z(_.message)}})})})}function R(w,M,k){if(!n)return;document.querySelectorAll(".status-popup-menu").forEach(f=>f.remove());let x=w.getBoundingClientRect(),_=document.createElement("div");_.className="status-popup-menu",_.style.cssText=`
      position: fixed;
      top: ${x.bottom+window.scrollY+4}px;
      left: ${x.left+window.scrollX}px;
      background: #172033;
      border: 1px solid var(--border);
      border-radius: 8px;
      box-shadow: 0 4px 16px rgba(0,0,0,0.12);
      z-index: 9999;
      padding: 6px 0;
      min-width: 160px;
    `;let B=[{key:"scheduled",label:"\u{1F7E1} Terjadwal"},{key:"in_progress",label:"\u{1F535} Sedang Proses"},{key:"completed",label:"\u{1F7E2} Selesai"}];_.innerHTML=B.map(f=>`
      <button class="status-option-item" data-key="${f.key}" style="
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
        color: ${f.key===k?"#2563EB":"var(--text-1)"};
        font-weight: ${f.key===k?"700":"500"};
      ">
        ${f.label} ${f.key===k?"\u2713":""}
      </button>
    `).join("");let P=f=>{!_.contains(f.target)&&f.target!==w&&(_.remove(),document.removeEventListener("click",P))};setTimeout(()=>document.addEventListener("click",P),50),_.querySelectorAll(".status-option-item").forEach(f=>{f.addEventListener("click",async()=>{let g=f.dataset.key;if(_.remove(),document.removeEventListener("click",P),g!==k)try{if(!(await X(`/api/os-migrations/${M}/status`,{method:"PATCH",body:JSON.stringify({status:g})})).ok)throw new Error("Gagal memperbarui status");ae("Status berhasil diubah"),p()}catch(D){z(D.message)}})}),document.body.appendChild(_)}function K(w=null){if(!n){z("Anda tidak memiliki izin untuk menambah atau mengedit migrasi OS.");return}let M=!!w,k=M?"Edit Jadwal Migrasi OS":"Tambah Jadwal Migrasi OS",_=`
      <form id="migration-form" style="display:flex; flex-direction:column; gap:14px;">
        <div class="form-group">
          <label style="font-weight:600; font-size:0.875rem; margin-bottom:4px; display:block;">Pilih Cabang *</label>
          <select name="branch_id" class="form-control" required style="width:100%; height:38px; border-radius:8px;">
            <option value="">-- Pilih Cabang --</option>
            ${h.map(B=>`
      <option value="${B.id}" ${w&&w.branch_id===B.id?"selected":""}>
        ${B.full_name}
      </option>
    `).join("")}
          </select>
        </div>

        <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div class="form-group">
            <label style="font-weight:600; font-size:0.875rem; margin-bottom:4px; display:block;">Vendor Outsourcing *</label>
            <select name="vendor" class="form-control" required style="width:100%; height:38px; border-radius:8px;">
              <option value="AEON" ${w?.vendor==="AEON"?"selected":""}>AEON</option>
              <option value="Triple S" ${w?.vendor==="Triple S"?"selected":""}>Triple S</option>
              <option value="Lainnya" ${w&&!["AEON","Triple S"].includes(w.vendor)?"selected":""}>Lainnya</option>
            </select>
          </div>

          <div class="form-group">
            <label style="font-weight:600; font-size:0.875rem; margin-bottom:4px; display:block;">Target Bulan *</label>
            <select name="target_month" class="form-control" required style="width:100%; height:38px; border-radius:8px;">
              <option value="2026-09" ${w?.target_month==="2026-09"?"selected":""}>September 2026</option>
              <option value="2026-10" ${w?.target_month==="2026-10"?"selected":""}>Oktober 2026</option>
              <option value="2026-11" ${w?.target_month==="2026-11"?"selected":""}>November 2026</option>
              <option value="2026-12" ${w?.target_month==="2026-12"?"selected":""}>Desember 2026</option>
            </select>
          </div>
        </div>

        <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div class="form-group">
            <label style="font-weight:600; font-size:0.875rem; margin-bottom:4px; display:block;">Status Pelaksanaan</label>
            <select name="status" class="form-control" style="width:100%; height:38px; border-radius:8px;">
              <option value="scheduled" ${w?.status==="scheduled"?"selected":""}>\u{1F7E1} Terjadwal</option>
              <option value="in_progress" ${w?.status==="in_progress"?"selected":""}>\u{1F535} Sedang Proses</option>
              <option value="completed" ${w?.status==="completed"?"selected":""}>\u{1F7E2} Selesai</option>
            </select>
          </div>

          <div class="form-group">
            <label style="font-weight:600; font-size:0.875rem; margin-bottom:4px; display:block;">Tanggal Realisasi (Opsional)</label>
            <input type="date" name="actual_date" class="form-control" value="${w?.actual_date||""}" style="width:100%; height:38px; border-radius:8px;">
          </div>
        </div>

        <div class="form-group">
          <label style="font-weight:600; font-size:0.875rem; margin-bottom:4px; display:block;">PIC / Koordinator Vendor (Opsional)</label>
          <input type="text" name="pic_vendor" class="form-control" placeholder="Nama PIC vendor / no HP" value="${w?.pic_vendor||""}" style="width:100%; height:38px; border-radius:8px;">
        </div>

        <div class="form-group">
          <label style="font-weight:600; font-size:0.875rem; margin-bottom:4px; display:block;">Catatan / Keterangan</label>
          <textarea name="notes" class="form-control" rows="2" placeholder="Catatan serah terima / info tambahan..." style="width:100%; border-radius:8px;">${w?.notes||""}</textarea>
        </div>
      </form>
    `;pe({title:k,content:_,confirmText:M?"Simpan Perubahan":"Tambah Cabang",onConfirm:async(B,P)=>{let f=B.querySelector("#migration-form");if(!f.checkValidity()){f.reportValidity();return}let g=new FormData(f),D={branch_id:parseInt(g.get("branch_id"),10),vendor:g.get("vendor"),target_month:g.get("target_month"),status:g.get("status"),actual_date:g.get("actual_date")||null,pic_vendor:g.get("pic_vendor")||null,notes:g.get("notes")||null};try{let O=M?`/api/os-migrations/${w.id}`:"/api/os-migrations",J=await X(O,{method:M?"PUT":"POST",body:JSON.stringify(D)});if(!J.ok){let $=J.data?.error||J.data?.message||"Gagal menyimpan data";throw new Error($)}ae(M?"Jadwal migrasi berhasil diupdate":"Jadwal migrasi berhasil ditambahkan"),P(),p()}catch(O){z(O.message)}}})}let I=document.getElementById("view-board-btn"),N=document.getElementById("view-table-btn");I.addEventListener("click",()=>{s="board",I.style.background="#fff",I.classList.remove("btn-ghost"),N.style.background="transparent",N.classList.add("btn-ghost"),C()}),N.addEventListener("click",()=>{s="table",N.style.background="#fff",N.classList.remove("btn-ghost"),I.style.background="transparent",I.classList.add("btn-ghost"),C()}),["migration-search","filter-vendor","filter-month","filter-status"].forEach(w=>{let M=document.getElementById(w);M&&(M.addEventListener("input",C),M.addEventListener("change",C))}),document.getElementById("btn-add-migration")?.addEventListener("click",()=>{K()}),document.getElementById("btn-export-migrations")?.addEventListener("click",async()=>{try{let w=(document.getElementById("migration-search")?.value||"").trim(),M=document.getElementById("filter-vendor")?.value||"",k=document.getElementById("filter-month")?.value||"",x=document.getElementById("filter-status")?.value||"",_={};M&&(_.vendor=M),k&&(_.month=k),x&&(_.status=x),w&&(_.search=w);let B=await ue("/api/os-migrations/export",_);if(!B.length){z("Tidak ada data untuk diexport");return}let P=B.map((f,g)=>({No:g+1,"Kode Cabang":f.branch_code,"Nama Cabang":f.branch_name,"Cabang Lengkap":f.branch_full_name,"Vendor OS":f.vendor,"Target Bulan":Oa[f.target_month]||f.target_month,Status:Mt[f.status]?.label||f.status,"Tanggal Realisasi":f.actual_date||"","Jml Personil FC":f.employee_count||0,"PIC Vendor":f.pic_vendor||"",Catatan:f.notes||""}));ge(P,`Timeline_Migrasi_OS_${new Date().toISOString().slice(0,10)}`),ae("File Excel berhasil diunduh")}catch(w){z("Gagal mengekspor data migrasi: "+w.message)}}),p()}Fe();we();Ne();var vi={Persiapan:{bg:"#F3E8FF",color:"#6B21A8",border:"#D8B4FE"},"Area Umum":{bg:"#E0F2FE",color:"#0369A1",border:"#BAE6FD"},"Ruang Tindakan":{bg:"#DBEAFE",color:"#1E40AF",border:"#93C5FD"},"Ruang Penunjang":{bg:"#F1F5F9",color:"#334155",border:"#CBD5E1"},Toilet:{bg:"#CCFBF1",color:"#0F766E",border:"#99F6E4"},"Area Luar":{bg:"#FEF3C7",color:"#92400E",border:"#FDE68A"},"Kontrol & Check":{bg:"#FFEDD5",color:"#C2410C",border:"#FED7AA"},Istirahat:{bg:"#E2E8F0",color:"#475569",border:"#CBD5E1"},"Layanan Pasien":{bg:"#DCFCE7",color:"#15803D",border:"#86EFAC"},"Sampah Umum":{bg:"#FEF9C3",color:"#854D0E",border:"#FDE047"},"Sampah Medis":{bg:"#FEE2E2",color:"#991B1B",border:"#FCA5A5"},Penutup:{bg:"#E0E7FF",color:"#3730A3",border:"#C7D2FE"}};async function ki(e){let t=de(),n=ie(t,"dac","write"),a=ie(t,"dac","delete"),o=ie(t,"export");e.innerHTML=`
    <div class="page-header" style="display:flex; justify-content:space-between; align-items:flex-start; flex-wrap:wrap; gap:16px; margin-bottom:20px;">
      <div>
        <h1 class="page-title" style="display:flex; align-items:center; gap:8px; margin:0 0 6px 0; font-size:1.6rem; font-weight:700; color:var(--text-1);">
          <span>\u23F1\uFE0F</span> Master DAC (Daily Activity Cleaning)
        </h1>
        <p class="page-subtitle" style="margin:0; font-size:0.9rem; color:var(--text-3);">
          Standar Acuan Alur Rutinitas Pembersihan Harian Facility Care (Shift Pagi & Shift Siang)
        </p>
      </div>
      <div style="display:flex; gap:10px; align-items:center; flex-wrap:wrap;">
        <div class="btn-group" style="background:#172033; padding:3px; border-radius:8px; display:inline-flex; border:1px solid #29364B;">
          <button id="view-timeline-btn" class="btn btn-sm" style="background:#2563EB; color:#fff; font-weight:600; border-radius:6px; box-shadow:0 1px 2px rgba(0,0,0,0.2); padding:6px 12px; cursor:pointer;">
            \u23F1\uFE0F Alur Timeline
          </button>
          <button id="view-table-btn" class="btn btn-sm btn-ghost" style="color:#94A3B8; font-weight:600; border-radius:6px; padding:6px 12px; cursor:pointer;">
            \u{1F4CB} Tabel Standar
          </button>
        </div>
        <button id="btn-print-dac" class="btn btn-outline" style="background:#172033; border:1px solid #29364B; color:#F8FAFC; font-weight:600; padding:8px 14px; border-radius:8px; display:inline-flex; align-items:center; gap:6px; cursor:pointer;" title="Cetak Format SOP untuk Ruang Janitor">
          \u{1F5A8}\uFE0F Cetak SOP
        </button>
        ${o?`
        <button id="btn-export-dac" class="btn btn-outline" style="background:#172033; border:1px solid #29364B; color:#F8FAFC; font-weight:600; padding:8px 14px; border-radius:8px; display:inline-flex; align-items:center; gap:6px; cursor:pointer;">
          \u{1F4E5} Export Excel
        </button>
        `:""}
        ${n?`
        <button id="btn-add-dac" class="btn btn-primary" style="font-weight:600; padding:8px 16px; border-radius:8px; display:inline-flex; align-items:center; gap:6px; cursor:pointer;">
          + Tambah Aktivitas
        </button>
        `:""}
      </div>
    </div>

    <!-- Shift Switcher Tabs -->
    <div style="display:flex; gap:8px; margin-bottom:20px; border-bottom:2px solid #29364B; padding-bottom:8px;">
      <button class="shift-tab-btn active" data-shift="pagi" style="border:none; background:#2563EB; color:#fff; font-weight:700; padding:8px 18px; border-radius:8px; cursor:pointer; display:inline-flex; align-items:center; gap:6px; font-size:0.9rem;">
        \u{1F305} Shift Pagi (08:00 - 17:00) <span class="badge" style="background:rgba(255,255,255,0.25); color:#fff; font-size:0.75rem; padding:2px 6px; border-radius:10px;" id="badge-count-pagi">31</span>
      </button>
      <button class="shift-tab-btn" data-shift="siang" style="border:1px solid #29364B; background:#172033; color:#94A3B8; font-weight:700; padding:8px 18px; border-radius:8px; cursor:pointer; display:inline-flex; align-items:center; gap:6px; font-size:0.9rem;">
        \u{1F307} Shift Siang (13:00 - 22:00) <span class="badge" style="background:rgba(255,255,255,0.08); color:#F8FAFC; font-size:0.75rem; padding:2px 6px; border-radius:10px;" id="badge-count-siang">33</span>
      </button>
      <button class="shift-tab-btn" data-shift="all" style="border:1px solid #29364B; background:#172033; color:#94A3B8; font-weight:700; padding:8px 18px; border-radius:8px; cursor:pointer; display:inline-flex; align-items:center; gap:6px; font-size:0.9rem;">
        \u{1F4D1} Semua Shift <span class="badge" style="background:rgba(255,255,255,0.08); color:#F8FAFC; font-size:0.75rem; padding:2px 6px; border-radius:10px;" id="badge-count-all">64</span>
      </button>
    </div>

    <!-- Summary KPI Cards -->
    <div id="dac-summary-cards" style="display:grid; grid-template-columns:repeat(auto-fit, minmax(220px, 1fr)); gap:16px; margin-bottom:24px;">
      <div class="kpi-card skeleton" style="height:90px; border-radius:12px;"></div>
      <div class="kpi-card skeleton" style="height:90px; border-radius:12px;"></div>
      <div class="kpi-card skeleton" style="height:90px; border-radius:12px;"></div>
      <div class="kpi-card skeleton" style="height:90px; border-radius:12px;"></div>
    </div>

    <!-- Filter Toolbar -->
    <div class="card" style="padding:14px 18px; margin-bottom:24px; background:#111827; border:1px solid #29364B; border-radius:12px; box-shadow:0 4px 12px rgba(0,0,0,0.25);">
      <div style="display:flex; flex-wrap:wrap; gap:12px; align-items:center; justify-content:space-between;">
        <div style="display:flex; flex-wrap:wrap; gap:12px; align-items:center; flex:1; min-width:280px;">
          <div style="position:relative; flex:1; min-width:200px; max-width:320px;">
            <input type="text" id="dac-search" class="form-control" placeholder="Cari lokasi, aktivitas, jam..." style="padding-left:34px; height:38px; border-radius:8px; font-size:0.9rem; width:100%; background:#1B2940; border:1px solid #29364B; color:#F8FAFC;">
            <span style="position:absolute; left:11px; top:50%; transform:translateY(-50%); color:var(--text-3); font-size:0.9rem;">\u{1F50D}</span>
          </div>
          <select id="dac-category-filter" class="form-control" style="width:auto; height:38px; border-radius:8px; font-size:0.875rem; min-width:160px; background:#172033; border:1px solid #29364B; color:#F8FAFC;">
            <option value="">Semua Kategori</option>
          </select>
          <select id="dac-period-filter" class="form-control" style="width:auto; height:38px; border-radius:8px; font-size:0.875rem; min-width:150px; background:#172033; border:1px solid #29364B; color:#F8FAFC;">
            <option value="">Semua Waktu</option>
            <option value="morning">Sesi Awal (08:00 - 12:00 / 13:00 - 17:00)</option>
            <option value="afternoon">Sesi Lanjutan (12:00 - 17:00 / 17:00 - 22:00)</option>
          </select>
        </div>
        <div id="dac-stats-text" style="font-size:0.85rem; color:var(--text-3); font-weight:500;">
          Memuat data...
        </div>
      </div>
    </div>

    <!-- Main Content Container -->
    <div id="dac-view-container"></div>
  `;let s="pagi",l="timeline",h=[];async function p(){try{let w=await X("/api/dac?all=1");if(!w.ok)throw new Error("Gagal memuat standar DAC");h=w.data?.data||w.data||[];let M=h.filter(B=>B.shift==="pagi").length,k=h.filter(B=>B.shift==="siang").length;document.getElementById("badge-count-pagi").textContent=M,document.getElementById("badge-count-siang").textContent=k,document.getElementById("badge-count-all").textContent=h.length;let x=document.getElementById("dac-category-filter"),_=Array.from(new Set(h.map(B=>B.category))).sort();x.innerHTML='<option value="">Semua Kategori</option>'+_.map(B=>`<option value="${Y(B)}">${Y(B)}</option>`).join(""),E(),C()}catch(w){let M=document.getElementById("dac-view-container");M.innerHTML=`
        <div class="empty-state" style="padding:40px; text-align:center;">
          <div style="font-size:2rem; margin-bottom:8px;">\u26A0\uFE0F</div>
          <p style="color:var(--danger); font-weight:600;">Gagal memuat data DAC: ${Y(w.message)}</p>
          <button type="button" id="dac-reload-btn" class="btn btn-outline btn-sm">Muat Ulang</button>
        </div>
      `,M.querySelector("#dac-reload-btn")?.addEventListener("click",()=>window.location.reload())}}function c(){let w=(document.getElementById("dac-search")?.value||"").toLowerCase().trim(),M=document.getElementById("dac-category-filter")?.value||"",k=document.getElementById("dac-period-filter")?.value||"";return h.filter(x=>{if(s!=="all"&&x.shift!==s||M&&x.category!==M)return!1;if(k==="morning"){let _=parseInt(x.start_time.split(":")[0],10);if(x.shift==="pagi"&&_>=12||x.shift==="siang"&&_>=17)return!1}else if(k==="afternoon"){let _=parseInt(x.start_time.split(":")[0],10);if(x.shift==="pagi"&&_<12||x.shift==="siang"&&_<17)return!1}return!(w&&!`${x.start_time} ${x.end_time} ${x.category} ${x.location} ${x.activity}`.toLowerCase().includes(w))})}function E(){let w=document.getElementById("dac-summary-cards");if(!w)return;let M=h.filter(P=>P.shift==="pagi"),k=h.filter(P=>P.shift==="siang"),x=s==="pagi"?M:s==="siang"?k:h,_=s==="pagi"?"08:00 \u2013 17:00":s==="siang"?"13:00 \u2013 22:00":"08:00 \u2013 22:00",B=s==="pagi"?"Shift Pagi":s==="siang"?"Shift Siang":"Semua Shift";w.innerHTML=`
      <div class="card" style="padding:16px 20px; background:#172033; border:1px solid #29364B; border-radius:12px; box-shadow:0 4px 12px rgba(0,0,0,0.25);">
        <div style="font-size:0.8rem; font-weight:600; color:var(--text-3); text-transform:uppercase; letter-spacing:0.5px; margin-bottom:6px;">Total Aktivitas (${B})</div>
        <div style="font-size:1.8rem; font-weight:800; color:var(--text-1); line-height:1.2;">${x.length} <span style="font-size:0.9rem; font-weight:500; color:var(--text-3);">Aktivitas</span></div>
        <div style="font-size:0.75rem; color:var(--text-3); margin-top:4px;">Urutan Terjadwal dari Awal s/d Penutup</div>
      </div>

      <div class="card" style="padding:16px 20px; background:#172033; border:1px solid #29364B; border-left:4px solid #2563EB; border-radius:12px; box-shadow:0 4px 12px rgba(0,0,0,0.25);">
        <div style="font-size:0.8rem; font-weight:600; color:#1E40AF; text-transform:uppercase; letter-spacing:0.5px; margin-bottom:6px;">Jam Kerja Standar</div>
        <div style="font-size:1.5rem; font-weight:800; color:#1E40AF; line-height:1.2;">${_}</div>
        <div style="font-size:0.75rem; color:var(--text-3); margin-top:4px;">9 Jam Operasional (Break 1 Jam)</div>
      </div>

      <div class="card" style="padding:16px 20px; background:#172033; border:1px solid #29364B; border-left:4px solid #14B8A6; border-radius:12px; box-shadow:0 4px 12px rgba(0,0,0,0.25);">
        <div style="font-size:0.8rem; font-weight:600; color:#0F766E; text-transform:uppercase; letter-spacing:0.5px; margin-bottom:6px;">Cakupan Operasional</div>
        <div style="font-size:1.8rem; font-weight:800; color:#0F766E; line-height:1.2;">12 <span style="font-size:0.9rem; font-weight:500; color:var(--text-3);">Kategori</span></div>
        <div style="font-size:0.75rem; color:var(--text-3); margin-top:4px;">Tindakan, Toilet, Area Umum, Medis, dll.</div>
      </div>

      <div class="card" style="padding:16px 20px; background:#172033; border:1px solid #29364B; border-left:4px solid #F59E0B; border-radius:12px; box-shadow:0 4px 12px rgba(0,0,0,0.25);">
        <div style="font-size:0.8rem; font-weight:600; color:#B45309; text-transform:uppercase; letter-spacing:0.5px; margin-bottom:6px;">Poin Wajib Harian</div>
        <div style="font-size:1.1rem; font-weight:700; color:#B45309; line-height:1.4;">Check Sheet & Sampah Medis</div>
        <div style="font-size:0.75rem; color:var(--text-3); margin-top:4px;">Kontrol mobilitas bertahap & sterilisasi</div>
      </div>
    `}function C(){let w=c(),M=document.getElementById("dac-stats-text");M&&(M.innerHTML=`Menampilkan <strong>${w.length}</strong> aktivitas`);let k=document.getElementById("dac-view-container");l==="timeline"?T(k,w):u(k,w)}function T(w,M){if(!M.length){w.innerHTML=`
        <div class="empty-state" style="padding:50px; text-align:center; background:#111827; border-radius:12px; border:1px solid var(--border);">
          <div style="font-size:2.5rem; margin-bottom:12px;">\u{1F50D}</div>
          <p style="font-size:1.05rem; font-weight:600; color:var(--text-2); margin-bottom:6px;">Tidak ada aktivitas yang cocok dengan filter</p>
          <p style="color:var(--text-3); font-size:0.875rem;">Coba reset kata kunci atau pilihan kategori di atas.</p>
        </div>
      `;return}w.innerHTML=`
      <div class="timeline-container" style="display:flex; flex-direction:column; gap:12px; max-width:960px; margin:0 auto;">
        ${M.map(k=>{let x=vi[k.category]||{bg:"#F1F5F9",color:"#334155",border:"#CBD5E1"},_=k.shift==="pagi"?"\u{1F305} Pagi":"\u{1F307} Siang",B=k.shift==="pagi"?"#FEF3C7":"#E0E7FF",P=k.shift==="pagi"?"#92400E":"#3730A3";return`
            <div class="dac-timeline-card" style="background:#172033; border:1px solid var(--border); border-radius:12px; padding:16px 20px; box-shadow:0 1px 3px rgba(0,0,0,0.2); display:flex; gap:16px; align-items:flex-start; position:relative;">
              <!-- Time Column -->
              <div style="min-width:130px; text-align:center; background:#111827; border:1px solid var(--border); border-radius:8px; padding:10px 8px;">
                <div style="font-size:1.1rem; font-weight:800; color:var(--text-1); font-family:monospace;">
                  ${k.start_time} - ${k.end_time}
                </div>
                <div style="font-size:0.75rem; color:var(--text-3); font-weight:600; margin-top:4px;">
                  Item #${k.item_no}
                </div>
                ${s==="all"?`
                  <div style="margin-top:6px;">
                    <span class="badge" style="background:${B}; color:${P}; font-size:0.7rem; font-weight:700; padding:2px 6px; border-radius:4px;">
                      ${_}
                    </span>
                  </div>
                `:""}
              </div>

              <!-- Content Column -->
              <div style="flex:1;">
                <div style="display:flex; justify-content:space-between; align-items:flex-start; gap:8px; margin-bottom:6px;">
                  <div style="display:flex; align-items:center; gap:8px; flex-wrap:wrap;">
                    <span class="badge" style="background:${x.bg}; color:${x.color}; border:1px solid ${x.border}; font-weight:700; font-size:0.75rem; padding:3px 10px; border-radius:6px;">
                      ${k.category}
                    </span>
                    <span style="font-weight:700; font-size:0.95rem; color:var(--text-1);">
                      \u{1F4CD} ${k.location}
                    </span>
                  </div>
                  ${n||a?`
                    <div style="display:flex; gap:4px;">
                      ${n?`<button class="btn btn-ghost btn-sm edit-dac-btn" data-id="${k.id}" title="Edit" style="padding:4px 8px; font-size:0.8rem; cursor:pointer;">\u270F\uFE0F</button>`:""}
                      ${a?`<button class="btn btn-ghost btn-sm delete-dac-btn" data-id="${k.id}" data-name="${k.location}" title="Hapus" style="padding:4px 8px; font-size:0.8rem; color:var(--danger); cursor:pointer;">\u{1F5D1}\uFE0F</button>`:""}
                    </div>
                  `:""}
                </div>

                <div style="font-size:0.9rem; color:var(--text-2); line-height:1.5; background:#1B2940; padding:10px 14px; border-radius:8px; border-left:3px solid ${x.color}; margin-top:8px;">
                  ${k.activity}
                </div>
              </div>
            </div>
          `}).join("")}
      </div>
    `,F(w)}function u(w,M){if(!M.length){w.innerHTML=`
        <div class="empty-state" style="padding:50px; text-align:center; background:#111827; border-radius:12px; border:1px solid var(--border);">
          <div style="font-size:2.5rem; margin-bottom:12px;">\u{1F50D}</div>
          <p style="font-size:1.05rem; font-weight:600; color:var(--text-2); margin-bottom:6px;">Tidak ada aktivitas yang cocok dengan filter</p>
        </div>
      `;return}w.innerHTML=`
      <div class="card" style="background:#172033; border:1px solid #29364B; border-radius:12px; overflow:hidden; box-shadow:0 4px 12px rgba(0,0,0,0.25);">
        <div class="table-responsive">
          <table class="table" style="width:100%; border-collapse:collapse;">
            <thead>
              <tr style="background:#1B2940; border-bottom:1px solid #29364B; text-align:left; color:#94A3B8;">
                <th style="padding:12px 14px; font-size:0.8rem; font-weight:700; color:var(--text-3); width:50px;">NO</th>
                ${s==="all"?'<th style="padding:12px 14px; font-size:0.8rem; font-weight:700; color:var(--text-3);">SHIFT</th>':""}
                <th style="padding:12px 14px; font-size:0.8rem; font-weight:700; color:var(--text-3); width:110px;">JAM</th>
                <th style="padding:12px 14px; font-size:0.8rem; font-weight:700; color:var(--text-3); width:140px;">KATEGORI</th>
                <th style="padding:12px 14px; font-size:0.8rem; font-weight:700; color:var(--text-3); width:200px;">LOKASI / AREA</th>
                <th style="padding:12px 14px; font-size:0.8rem; font-weight:700; color:var(--text-3);">AKTIVITAS</th>
                ${n||a?'<th style="padding:12px 14px; font-size:0.8rem; font-weight:700; color:var(--text-3); text-align:right; width:100px;">AKSI</th>':""}
              </tr>
            </thead>
            <tbody>
              ${M.map(k=>{let x=vi[k.category]||{bg:"#F1F5F9",color:"#334155",border:"#CBD5E1"};return`
                  <tr style="border-bottom:1px solid var(--border);">
                    <td style="padding:12px 14px; font-size:0.85rem; color:var(--text-3); font-weight:600;">${k.item_no}</td>
                    ${s==="all"?`
                      <td style="padding:12px 14px;">
                        <span class="badge" style="background:${k.shift==="pagi"?"#FEF3C7":"#E0E7FF"}; color:${k.shift==="pagi"?"#92400E":"#3730A3"}; font-size:0.75rem; font-weight:700; padding:3px 8px; border-radius:4px;">
                          ${k.shift==="pagi"?"Pagi":"Siang"}
                        </span>
                      </td>
                    `:""}
                    <td style="padding:12px 14px; font-family:monospace; font-weight:700; font-size:0.85rem; color:var(--text-1); white-space:nowrap;">
                      ${k.start_time} - ${k.end_time}
                    </td>
                    <td style="padding:12px 14px;">
                      <span class="badge" style="background:${x.bg}; color:${x.color}; border:1px solid ${x.border}; font-weight:700; font-size:0.75rem; padding:3px 8px; border-radius:6px; white-space:nowrap;">
                        ${k.category}
                      </span>
                    </td>
                    <td style="padding:12px 14px; font-size:0.875rem; font-weight:600; color:var(--text-1);">
                      ${k.location}
                    </td>
                    <td style="padding:12px 14px; font-size:0.875rem; color:var(--text-2); line-height:1.4;">
                      ${k.activity}
                    </td>
                    ${n||a?`
                      <td style="padding:12px 14px; text-align:right; white-space:nowrap;">
                        ${n?`<button class="btn btn-outline btn-sm edit-dac-btn" data-id="${k.id}" style="padding:3px 8px; font-size:0.75rem; border-radius:6px; margin-right:4px;">Edit</button>`:""}
                        ${a?`<button class="btn btn-danger btn-sm delete-dac-btn" data-id="${k.id}" data-name="${k.location}" style="padding:3px 8px; font-size:0.75rem; border-radius:6px;">Hapus</button>`:""}
                      </td>
                    `:""}
                  </tr>
                `}).join("")}
            </tbody>
          </table>
        </div>
      </div>
    `,F(w)}function F(w){w.querySelectorAll(".edit-dac-btn").forEach(M=>{M.addEventListener("click",()=>{let k=M.dataset.id,x=h.find(_=>String(_.id)===String(k));x&&R(x)})}),w.querySelectorAll(".delete-dac-btn").forEach(M=>{M.addEventListener("click",()=>{let k=M.dataset.id,x=M.dataset.name||"item";me(`Apakah Anda yakin ingin menghapus aktivitas pada lokasi <strong>${x}</strong>?`,async()=>{try{if(!(await X(`/api/dac/${k}`,{method:"DELETE"})).ok)throw new Error("Gagal menghapus item");ae("Aktivitas DAC berhasil dihapus"),p()}catch(_){z(_.message)}})})})}function R(w=null){if(!n){z("Anda tidak memiliki izin untuk menambah atau mengedit aktivitas.");return}let M=!!w,k=M?"Edit Aktivitas DAC":"Tambah Aktivitas DAC",x=["Persiapan","Area Umum","Ruang Tindakan","Ruang Penunjang","Toilet","Area Luar","Kontrol & Check","Istirahat","Layanan Pasien","Sampah Umum","Sampah Medis","Penutup"],_=`
      <form id="dac-form" style="display:flex; flex-direction:column; gap:14px;">
        <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div class="form-group">
            <label style="font-weight:600; font-size:0.875rem; margin-bottom:4px; display:block;">Pilih Shift *</label>
            <select name="shift" class="form-control" required style="width:100%; height:38px; border-radius:8px;">
              <option value="pagi" ${w?.shift==="pagi"||!w&&s==="pagi"?"selected":""}>\u{1F305} Shift Pagi (08:00 - 17:00)</option>
              <option value="siang" ${w?.shift==="siang"||!w&&s==="siang"?"selected":""}>\u{1F307} Shift Siang (13:00 - 22:00)</option>
            </select>
          </div>

          <div class="form-group">
            <label style="font-weight:600; font-size:0.875rem; margin-bottom:4px; display:block;">Nomor Urut *</label>
            <input type="number" name="item_no" class="form-control" value="${w?.item_no||""}" placeholder="Otomatis jika kosong" style="width:100%; height:38px; border-radius:8px;">
          </div>
        </div>

        <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div class="form-group">
            <label style="font-weight:600; font-size:0.875rem; margin-bottom:4px; display:block;">Jam Mulai (HH:MM) *</label>
            <input type="time" name="start_time" class="form-control" value="${w?.start_time||""}" required style="width:100%; height:38px; border-radius:8px;">
          </div>

          <div class="form-group">
            <label style="font-weight:600; font-size:0.875rem; margin-bottom:4px; display:block;">Jam Selesai (HH:MM) *</label>
            <input type="time" name="end_time" class="form-control" value="${w?.end_time||""}" required style="width:100%; height:38px; border-radius:8px;">
          </div>
        </div>

        <div class="form-group">
          <label style="font-weight:600; font-size:0.875rem; margin-bottom:4px; display:block;">Kategori *</label>
          <input list="cat-suggestions" name="category" class="form-control" value="${w?.category||""}" required placeholder="Pilih atau ketik kategori..." style="width:100%; height:38px; border-radius:8px;">
          <datalist id="cat-suggestions">
            ${x.map(B=>`<option value="${B}">`).join("")}
          </datalist>
        </div>

        <div class="form-group">
          <label style="font-weight:600; font-size:0.875rem; margin-bottom:4px; display:block;">Lokasi / Area *</label>
          <input type="text" name="location" class="form-control" value="${w?.location||""}" required placeholder="Contoh: Ruang tindakan Du 101, Lobby, Toilet GF" style="width:100%; height:38px; border-radius:8px;">
        </div>

        <div class="form-group">
          <label style="font-weight:600; font-size:0.875rem; margin-bottom:4px; display:block;">Aktivitas Detail *</label>
          <textarea name="activity" class="form-control" rows="3" required placeholder="Deskripsi aktivitas pembersihan..." style="width:100%; border-radius:8px;">${w?.activity||""}</textarea>
        </div>
      </form>
    `;pe({title:k,content:_,confirmText:M?"Simpan Perubahan":"Tambah Aktivitas",onConfirm:async(B,P)=>{let f=B.querySelector("#dac-form");if(!f.checkValidity()){f.reportValidity();return}let g=new FormData(f),D={shift:g.get("shift"),item_no:g.get("item_no")?parseInt(g.get("item_no"),10):void 0,start_time:g.get("start_time"),end_time:g.get("end_time"),category:g.get("category"),location:g.get("location"),activity:g.get("activity")};try{let O=M?`/api/dac/${w.id}`:"/api/dac",J=await X(O,{method:M?"PUT":"POST",body:JSON.stringify(D)});if(!J.ok){let $=J.data?.error||J.data?.message||"Gagal menyimpan aktivitas";throw new Error($)}ae(M?"Aktivitas berhasil diperbarui":"Aktivitas berhasil ditambahkan"),P(),p()}catch(O){z(O.message)}}})}function K(){let w=c();if(!w.length){z("Tidak ada data untuk dicetak");return}let M=s==="pagi"?"SHIFT PAGI (08:00 - 17:00)":s==="siang"?"SHIFT SIANG (13:00 - 22:00)":"SELURUH SHIFT (PAGI & SIANG)",k=window.open("","_blank");if(!k){z("Popup diblokir browser, mohon izinkan popup untuk mencetak");return}k.document.write(`
      <!DOCTYPE html>
      <html>
      <head>
        <title>STANDAR ACUAN DAC - ${M}</title>
        <style>body{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Arial,sans-serif;margin:20px;color:#1e293b;font-size:11pt}.header{text-align:center;border-bottom:2px solid #0f172a;padding-bottom:12px;margin-bottom:20px}.title{font-size:16pt;font-weight:bold;margin:0;color:#0f172a}.subtitle{font-size:11pt;color:#475569;margin:4px 0 0 0}table{width:100%;border-collapse:collapse;margin-top:10px}th,td{border:1px solid #cbd5e1;padding:7px 10px;text-align:left;vertical-align:top}th{background-color:#f1f5f9;font-weight:bold;font-size:10pt;text-transform:uppercase}.jam{font-family:monospace;font-weight:bold;white-space:nowrap}.badge{display:inline-block;padding:2px 6px;border-radius:4px;font-size:8pt;font-weight:bold}@media print{body{margin:10mm}button{display:none}}</style>
      </head>
      <body>
        <div class="header">
          <div class="title">FACILITY CARE MANAGEMENT SYSTEM (FCMS)</div>
          <div class="subtitle">STANDAR RUTINITAS DAILY ACTIVITY CLEANING (DAC) \u2014 ${M}</div>
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
            ${w.map(x=>`
              <tr>
                <td style="text-align:center; font-weight:bold;">${x.item_no}</td>
                ${s==="all"?`<td>${x.shift==="pagi"?"Pagi":"Siang"}</td>`:""}
                <td class="jam">${x.start_time} - ${x.end_time}</td>
                <td><strong>${x.category}</strong></td>
                <td><strong>${x.location}</strong></td>
                <td>${x.activity}</td>
              </tr>
            `).join("")}
          </tbody>
        </table>
        <div style="margin-top:20px; text-align:right; font-size:9pt; color:#64748b;">
          Dicetak pada: ${new Date().toLocaleString("id-ID")} | Standar Resmi Operasional Facility Care Audy Dental
        </div>
      </body>
      </html>
    `),k.document.close(),k.focus(),setTimeout(()=>{k.print()},400)}document.querySelectorAll(".shift-tab-btn").forEach(w=>{w.addEventListener("click",()=>{document.querySelectorAll(".shift-tab-btn").forEach(M=>{M.style.background="#F1F5F9",M.style.color="var(--text-2)"}),w.style.background="#2563EB",w.style.color="#fff",s=w.dataset.shift,E(),C()})});let I=document.getElementById("view-timeline-btn"),N=document.getElementById("view-table-btn");I.addEventListener("click",()=>{l="timeline",I.style.background="#fff",I.classList.remove("btn-ghost"),N.style.background="transparent",N.classList.add("btn-ghost"),C()}),N.addEventListener("click",()=>{l="table",N.style.background="#fff",N.classList.remove("btn-ghost"),I.style.background="transparent",I.classList.add("btn-ghost"),C()}),["dac-search","dac-category-filter","dac-period-filter"].forEach(w=>{let M=document.getElementById(w);M&&(M.addEventListener("input",C),M.addEventListener("change",C))}),document.getElementById("btn-add-dac")?.addEventListener("click",()=>{R()}),document.getElementById("btn-print-dac").addEventListener("click",()=>{K()}),document.getElementById("btn-export-dac")?.addEventListener("click",async()=>{try{let w=(document.getElementById("dac-search")?.value||"").trim(),M=document.getElementById("dac-category-filter")?.value||"",k=document.getElementById("dac-period-filter")?.value||"",x={};s!=="all"&&(x.shift=s),M&&(x.category=M),w&&(x.search=w);let _=await ue("/api/dac/export",x);if((k==="morning"||k==="afternoon")&&(_=_.filter(P=>{let f=parseInt(String(P.start_time||"0").split(":")[0],10);if(k==="morning"){if(P.shift==="pagi"&&f>=12||P.shift==="siang"&&f>=17)return!1}else if(P.shift==="pagi"&&f<12||P.shift==="siang"&&f<17)return!1;return!0})),!_.length){z("Tidak ada data untuk diexport");return}let B=_.map(P=>({No:P.item_no,Shift:P.shift==="pagi"?"Shift Pagi":"Shift Siang","Jam Mulai":P.start_time,"Jam Selesai":P.end_time,Kategori:P.category,"Lokasi / Area":P.location,"Rincian Aktivitas":P.activity}));ge(B,`DAC_Daily_Activity_Cleaning_${s}_${new Date().toISOString().slice(0,10)}`),ae("File Excel berhasil diunduh")}catch(w){z("Gagal mengekspor data DAC: "+w.message)}}),p()}Ne();async function wi(e){let t=de();if(!t||String(t.role).toLowerCase()!=="superadmin"||t.isRestrictedToken===!0||Number(t.must_change_password)===1){e.innerHTML='<div class="empty-state"><p class="text-danger">Akses ditolak. Halaman ini hanya untuk Superadmin.</p></div>';return}e.innerHTML=`
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
  `;let n=document.getElementById("btnRefreshSync"),a=document.getElementById("queueStatusFilter");n.addEventListener("click",p),a.addEventListener("change",C),document.getElementById("btnRetryAll").addEventListener("click",()=>h("retry",{allFailed:!0})),document.getElementById("btnResetStuck").addEventListener("click",()=>h("reset-stuck",{})),document.getElementById("btnPauseSync").addEventListener("click",()=>h("pause",{})),document.getElementById("btnResumeSync").addEventListener("click",()=>h("resume",{})),document.getElementById("btnReconcile").addEventListener("click",()=>h("reconcile",{module:"ALL",repairMode:!0})),document.getElementById("btnRetrySelected").addEventListener("click",()=>{let u=Array.from(document.querySelectorAll(".chk-queue:checked")).map(F=>F.value);if(u.length===0)return alert("No items selected");h("retry",{ids:u})}),document.getElementById("chkAllQueue").addEventListener("change",u=>{document.querySelectorAll(".chk-queue").forEach(F=>F.checked=u.target.checked)});function o(){n&&(n.disabled=!0,n.textContent="Loading...")}function s(){n&&(n.disabled=!1,n.textContent="Refresh")}let l=u=>u&&typeof u=="object"&&!Array.isArray(u)&&"data"in u&&"success"in u?u.data:u;async function h(u,F){if(confirm(`Are you sure you want to execute action: ${u}?`)){o();try{let R=await X(`/api/sync/actions/${u}`,{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify(F)});if(R.ok){let K=l(R.data);alert(K?.message||R.data?.message||"Success"),p()}else z(R.data?.error||R.error||"Action failed")}catch(R){z(R.message)}s()}}await p();async function p(){o(),await Promise.all([E(),C(),c(),T()]),s()}async function c(){try{let u=await X("/api/sync/performance");if(!u.ok||!u.data)return;let F=l(u.data);if(!F||F.status==="DATA_UNAVAILABLE"||!F.webhook){let M=document.getElementById("latencyHistograms");M&&(M.innerHTML='<p class="text-gray-500">Belum ada aktivitas sinkronisasi dalam 24 jam terakhir.</p>');let k=document.getElementById("lblSnapshotStatus");k&&(k.innerHTML='Status: <b>READY</b><br><span class="text-xs text-gray-500">Siap menjalankan rekonsiliasi & snapshot</span>');return}let{webhook:R,google_api:K,d1:I,queue:N,throughput:w}=F;document.getElementById("latencyHistograms").innerHTML=`
        <table class="w-full text-left">
          <thead><tr class="text-gray-500 border-b"><th>Metric</th><th>P50</th><th>P95</th><th>P99</th><th>Max</th></tr></thead>
          <tbody class="divide-y">
            <tr><td class="py-1">Webhook</td><td>${R.P50}ms</td><td>${R.P95}ms</td><td>${R.P99}ms</td><td>${R.Max}ms</td></tr>
            <tr><td class="py-1">Google API</td><td>${K.P50}ms</td><td>${K.P95}ms</td><td>${K.P99}ms</td><td>${K.Max}ms</td></tr>
            <tr><td class="py-1">D1 Execute</td><td>${I.P50}ms</td><td>${I.P95}ms</td><td>${I.P99}ms</td><td>${I.Max}ms</td></tr>
            <tr><td class="py-1">Queue Delay</td><td>${N.P50}ms</td><td>${N.P95}ms</td><td>${N.P99}ms</td><td>${N.Max}ms</td></tr>
          </tbody>
        </table>
        <div class="mt-4 pt-3 border-t text-gray-600 flex justify-between">
          <span>Throughput: <b>${w.events_per_sec}</b> ev/sec</span>
          <span><b>${w.events_per_min}</b> ev/min</span>
        </div>
      `,document.getElementById("lblSnapshotStatus").innerHTML='Checked 10 modules. Status: <b>COMPLETED</b><br><span class="text-xs text-gray-500">Run today at 02:00</span>'}catch(u){console.error(u)}}async function E(){try{let u=await X("/api/sync/health");if(!u.ok)return z("Failed to fetch sync health");let F=l(u.data)||{},{status:R,queue:K={},circuit_breaker:I}=F,N=`
        <div class="bg-white p-4 rounded-lg shadow border-l-4 ${R==="HEALTHY"?"border-green-500":R==="WARNING"?"border-yellow-500":"border-red-500"}">
          <p class="text-sm text-gray-500">System Health</p>
          <p class="text-2xl font-bold ${R==="HEALTHY"?"text-green-600":R==="WARNING"?"text-yellow-600":"text-red-600"}">${R}</p>
        </div>
        <div class="bg-white p-4 rounded-lg shadow border-l-4 border-blue-500">
          <p class="text-sm text-gray-500">Pending Queue</p>
          <p class="text-2xl font-bold text-gray-800">${K.pending||0}</p>
        </div>
        <div class="bg-white p-4 rounded-lg shadow border-l-4 border-orange-500">
          <p class="text-sm text-gray-500">Failed / Retrying</p>
          <p class="text-2xl font-bold text-gray-800">${K.failed||0}</p>
        </div>
        <div class="bg-white p-4 rounded-lg shadow border-l-4 border-red-500">
          <p class="text-sm text-gray-500">Dead Letters</p>
          <p class="text-2xl font-bold text-gray-800">${K.dead_letter||0}</p>
        </div>
      `;document.getElementById("syncOverviewCards").innerHTML=N;let w=document.getElementById("cbStateBadge"),M=document.getElementById("cbStateDesc"),k=document.getElementById("cbStatusCard");k.className="bg-white rounded-lg shadow p-6 border-l-4",I==="CLOSED"?(w.className="px-3 py-1 rounded-full text-sm font-medium bg-green-100 text-green-800",w.textContent="CLOSED",M.textContent="Traffic is flowing normally to Google Sheets.",k.classList.add("border-green-500")):I==="OPEN"?(w.className="px-3 py-1 rounded-full text-sm font-medium bg-red-100 text-red-800",w.textContent="OPEN",M.textContent="Failing fast. Traffic to Google Sheets is paused due to repeated failures.",k.classList.add("border-red-500")):I==="HALF_OPEN"?(w.className="px-3 py-1 rounded-full text-sm font-medium bg-yellow-100 text-yellow-800",w.textContent="HALF-OPEN",M.textContent="Testing recovery. Permitting limited traffic to verify stability.",k.classList.add("border-yellow-500")):w.textContent=I||"UNKNOWN"}catch(u){console.error(u)}}async function C(){try{let u=document.getElementById("queueStatusFilter").value,F=await X("/api/sync/queue?limit=15"+(u?"&status="+u:""));if(!F.ok)return;let R=document.getElementById("queueTableBody"),K=l(F.data),I=K&&!Array.isArray(K)&&Array.isArray(K.data)?K.data:K,N=Array.isArray(I)?I:[];if(N.length===0){R.innerHTML='<tr><td colspan="6" class="px-4 py-4 text-center text-gray-500">No events found</td></tr>';return}R.innerHTML=N.map(w=>`
        <tr class="hover:bg-gray-50">
          <td class="px-4 py-2"><input type="checkbox" class="chk-queue" value="${Y(w.id)}" /></td>
          <td class="px-4 py-2 font-mono text-xs text-gray-500" title="${Y(w.id)}">#${Y(String(w.id).split("-")[0])}</td>
          <td class="px-4 py-2 font-medium">${Y(w.entity_name)}</td>
          <td class="px-4 py-2">
             <span class="px-2 py-0.5 rounded text-xs ${w.action==="INSERT"?"bg-blue-100 text-blue-800":w.action==="UPDATE"?"bg-purple-100 text-purple-800":"bg-red-100 text-red-800"}">${Y(w.action)}</span>
          </td>
          <td class="px-4 py-2">
             <span class="px-2 py-0.5 rounded text-xs ${w.status==="PENDING"?"bg-yellow-100 text-yellow-800":w.status==="PROCESSING"?"bg-blue-100 text-blue-800":w.status==="DEAD_LETTER"?"bg-red-100 text-red-800":"bg-gray-100 text-gray-800"}">${Y(w.status)}</span>
             ${w.last_error?`<br><span class="text-xs text-red-500 max-w-xs block truncate" title="${Y(w.last_error)}">${Y(w.last_error)}</span>`:""}
          </td>
          <td class="px-4 py-2 text-gray-600">${Y(String(w.retry_count||0))}</td>
          <td class="px-4 py-2 text-gray-500 whitespace-nowrap">${Y(typeof window.formatDate=="function"?window.formatDate(w.created_at):w.created_at||"-")}</td>
        </tr>
      `).join("")}catch(u){console.error(u)}}async function T(){try{let u=await X("/api/sync/metrics");if(!u.ok)return;let F=document.getElementById("metricsTableBody"),R=l(u.data),K=Array.isArray(R)?R:[];if(K.length===0){F.innerHTML='<tr><td colspan="5" class="px-4 py-4 text-center text-gray-500">No metrics found in last 24h</td></tr>';return}F.innerHTML=K.map(I=>`
        <tr class="hover:bg-gray-50">
          <td class="px-4 py-2 font-medium text-gray-800">${Y(I.module)}</td>
          <td class="px-4 py-2 text-gray-600">${Y(String(I.total_events))}</td>
          <td class="px-4 py-2 text-gray-600">${Math.round(I.avg_webhook_ms||0)} ms</td>
          <td class="px-4 py-2 text-gray-600">${Math.round(I.avg_d1_ms||0)} ms</td>
          <td class="px-4 py-2 text-gray-600">${Math.round(I.avg_queue_wait_ms||0)} ms</td>
        </tr>
      `).join("")}catch(u){console.error(u)}}}window.parseFlexibleDate=e=>{if(!e||e==="-")return"";if(e=String(e).trim(),/^\d{5}$/.test(e)){let n=Math.floor(Number(e)-25569);return new Date(n*86400*1e3).toISOString().split("T")[0]}if(e.match(/^\d{2}[\/\-]\d{2}[\/\-]\d{4}$/)){let n=e.split(/[\/\-]/);return`${n[2]}-${n[1]}-${n[0]}`}let t=e.match(/^(\d{4})-(\d{2})-(\d{2})(?:T.*)?$/);if(t){let n=t[1],a=parseInt(t[2],10),o=parseInt(t[3],10);if(a>12&&o<=12)return`${n}-${t[3]}-${t[2]}`}return e.split("T")[0]};window.formatDate=e=>{let t=window.parseFlexibleDate(e);if(!t)return"";let n=t.split("-");if(n.length===3&&n[0].length===4){let a=["Januari","Februari","Maret","April","Mei","Juni","Juli","Agustus","September","Oktober","November","Desember"],o=parseInt(n[2],10),s=a[parseInt(n[1],10)-1];return`${o} ${s} ${n[0]}`}return t};function xe(e){return async t=>{if(!dt()&&!de()){qe("/login");return}return e(t)}}var Kt=null;function Mr(){Kt&&clearInterval(Kt);let e=()=>{let t=new Date,n=t.toLocaleTimeString("id-ID",{hour:"2-digit",minute:"2-digit",second:"2-digit"}).replace(/\./g,":"),a=t.toLocaleDateString("id-ID",{weekday:"long",day:"numeric",month:"long",year:"numeric"}),o=document.getElementById("header-clock-time"),s=document.getElementById("header-clock-date");o&&(o.textContent=n),s&&(s.textContent=a)};e(),Kt=setInterval(e,1e3)}async function zr(){try{let e=await X("/api/dashboard/kpi");if(!e.ok)return;let t=e.data?.data||e.data||{},n=(a,o)=>{let s=document.getElementById(a);s&&(s.textContent=o>0?o:"",s.style.display=o>0?"inline-flex":"none")};n("badge-issues",t.issues?.current||0),n("badge-contracts",t.expiring30?.current||0),n("badge-schedule",t.schedule?.current||0)}catch{}}var xt=[];async function Kr(){try{let e=await X("/api/dashboard/notifications");if(!e.ok)return;xt=e.data?.data||e.data||[];let t=document.getElementById("notif-dot");if(t){let n=xt.length>0?xt.length:51;t.style.display="inline-flex",t.textContent=n}}catch{}}function Nr(){if(!xt.length){pe({title:"Notifikasi",content:'<div class="empty-state"><p>Tidak ada notifikasi baru.</p></div>',confirmText:"Tutup",onConfirm:(t,n)=>n()});return}let e=`
    <div class="notif-list" style="max-height: 400px; overflow-y: auto;">
      ${xt.map(t=>`
        <div class="notif-item notif-severity-${t.severity||"info"}" style="padding: 12px; border-bottom: 1px solid var(--border); border-left: 4px solid var(--${t.severity==="danger"?"danger":t.severity==="warning"?"warning":"primary"}); margin-bottom: 8px; border-radius: 8px; background: #172033; color: #F8FAFC;">
          <div style="font-weight: 600; font-size: 0.9rem; color: #F8FAFC;">${Y(t.title)}</div>
          <div style="display: flex; justify-content: space-between; align-items: center; margin-top: 4px; font-size: 0.75rem; color: var(--text-3);">
            <span>\u{1F4C5} ${Y(t.date)}</span>
            <span class="badge badge-${t.severity==="danger"?"danger":t.severity==="warning"?"warning":"info"}">${Y(String(t.type||"").toUpperCase())}</span>
          </div>
        </div>
      `).join("")}
    </div>
  `;pe({title:`Notifikasi (${xt.length})`,content:e,confirmText:"Tutup",onConfirm:(t,n)=>n()})}function Si(){let e=de(),t=(e?.full_name||"U")[0].toUpperCase();document.getElementById("app").innerHTML=`
    <div class="app-layout">
      <!-- Sidebar dark -->
      <aside class="sidebar" id="sidebar">
        <div class="sidebar-header">
          <div class="sidebar-logo">
            <div class="logo-icon-wrap" style="width:34px;height:34px;border-radius:10px;background:linear-gradient(135deg,#2563EB,#0EA5E9);display:flex;align-items:center;justify-content:center;box-shadow:0 4px 12px rgba(37,99,235,0.4);color:#fff;flex-shrink:0;">
              <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><path d="M18 20V10M12 20V4M6 20v-6"/></svg>
            </div>
            <div class="logo-text-group">
              <span class="logo-text">FCMS</span>
              <span class="logo-subtext">Facilitycare Management System</span>
            </div>
          </div>
          <button class="sidebar-close" id="sidebar-close" aria-label="Close">\u2715</button>
        </div>

        <nav class="sidebar-nav" id="sidebar-nav">

          <!-- Utama -->
          <div class="nav-section">
            <span class="nav-section-label">UTAMA</span>
            <a href="#/dashboard" class="nav-item" data-route="/dashboard">
              <span class="nav-icon">
                <svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><rect x="3" y="3" width="7" height="7" rx="1"/><rect x="14" y="3" width="7" height="7" rx="1"/><rect x="3" y="14" width="7" height="7" rx="1"/><rect x="14" y="14" width="7" height="7" rx="1"/></svg>
              </span>
              <span class="nav-label">Dashboard</span>
            </a>
            <a href="#/calendar" class="nav-item" data-route="/calendar">
              <span class="nav-icon">
                <svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><rect x="3" y="4" width="18" height="18" rx="2"/><line x1="16" y1="2" x2="16" y2="6"/><line x1="8" y1="2" x2="8" y2="6"/><line x1="3" y1="10" x2="21" y2="10"/></svg>
              </span>
              <span class="nav-label">Kalender</span>
            </a>
          </div>

          <!-- SDM -->
          <div class="nav-section">
            <span class="nav-section-label">SDM</span>
            <a href="#/employees" class="nav-item" data-route="/employees">
              <span class="nav-icon">
                <svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><circle cx="9" cy="7" r="4"/><path d="M3 21v-2a4 4 0 014-4h4a4 4 0 014 4v2"/><circle cx="19" cy="7" r="2"/><path d="M23 21v-1a3 3 0 00-3-3"/></svg>
              </span>
              <span class="nav-label">Master Karyawan</span>
            </a>
            <a href="#/contracts" class="nav-item" data-route="/contracts">
              <span class="nav-icon">
                <svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M14 2H6a2 2 0 00-2 2v16a2 2 0 002 2h12a2 2 0 002-2V8z"/><polyline points="14 2 14 8 20 8"/><line x1="16" y1="13" x2="8" y2="13"/><line x1="16" y1="17" x2="8" y2="17"/></svg>
              </span>
              <span class="nav-label">Data Kontrak</span>
              <span class="nav-badge" id="badge-contracts"></span>
            </a>
            <a href="#/sp" class="nav-item" data-route="/sp">
              <span class="nav-icon">
                <svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"/><polyline points="22 4 12 14.01 9 11.01"/></svg>
              </span>
              <span class="nav-label">Data Sp</span>
            </a>
            <a href="#/mutasi" class="nav-item" data-route="/mutasi">
              <span class="nav-icon">
                <svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><polyline points="16 3 21 3 21 8"/><line x1="4" y1="20" x2="21" y2="3"/><polyline points="21 16 21 21 16 21"/><line x1="15" y1="15" x2="21" y2="21"/><line x1="4" y1="4" x2="9" y2="9"/></svg>
              </span>
              <span class="nav-label">Data Mutasi</span>
            </a>
            <a href="#/overtime" class="nav-item" data-route="/overtime">
              <span class="nav-icon">
                <svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><circle cx="12" cy="12" r="10"/><polyline points="12 6 12 12 16 14"/></svg>
              </span>
              <span class="nav-label">Data Lembur</span>
            </a>
          </div>

          <!-- Operasional -->
          <div class="nav-section">
            <span class="nav-section-label">OPERASIONAL</span>
            <a href="#/timeline" class="nav-item" data-route="/timeline">
              <span class="nav-icon">
                <svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><circle cx="12" cy="12" r="10"/><polyline points="12 6 12 12 16 14"/></svg>
              </span>
              <span class="nav-label">Time Line</span>
              <span class="nav-badge" id="badge-schedule"></span>
            </a>
            <a href="#/os-migrations" class="nav-item" data-route="/os-migrations">
              <span class="nav-icon">
                <svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><polyline points="16 3 21 3 21 8"/><line x1="4" y1="20" x2="21" y2="3"/><polyline points="21 16 21 21 16 21"/><line x1="15" y1="15" x2="21" y2="21"/><line x1="4" y1="4" x2="9" y2="9"/></svg>
              </span>
              <span class="nav-label">Migrasi OS</span>
            </a>
            <a href="#/issues" class="nav-item" data-route="/issues">
              <span class="nav-icon">
                <svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><circle cx="12" cy="12" r="10"/><line x1="12" y1="8" x2="12" y2="12"/><line x1="12" y1="16" x2="12.01" y2="16"/></svg>
              </span>
              <span class="nav-label">Permasalahan</span>
              <span class="nav-badge badge-danger" id="badge-issues"></span>
            </a>
            <a href="#/training" class="nav-item" data-route="/training">
              <span class="nav-icon">
                <svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M22 10v6M2 10l10-5 10 5-10 5z"/><path d="M6 12v5c3 3 9 3 12 0v-5"/></svg>
              </span>
              <span class="nav-label">Training</span>
            </a>
          </div>

          <!-- Laporan -->
          <div class="nav-section">
            <span class="nav-section-label">LAPORAN</span>
            <a href="#/reports/inspection" class="nav-item" data-route="/reports/inspection">
              <span class="nav-icon">
                <svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/></svg>
              </span>
              <span class="nav-label">Report Inspeksi Hygiene 2026</span>
            </a>
            <a href="#/reports/cleaning" class="nav-item" data-route="/reports/cleaning">
              <span class="nav-icon">
                <svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><polyline points="23 6 13.5 15.5 8.5 10.5 1 18"/><polyline points="17 6 23 6 23 12"/></svg>
              </span>
              <span class="nav-label">Report GCDC 2026</span>
            </a>
            <a href="#/reports/fogging" class="nav-item" data-route="/reports/fogging">
              <span class="nav-icon">
                <svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><line x1="8" y1="6" x2="21" y2="6"/><line x1="8" y1="12" x2="21" y2="12"/><line x1="8" y1="18" x2="21" y2="18"/><line x1="3" y1="6" x2="3.01" y2="6"/><line x1="3" y1="12" x2="3.01" y2="12"/><line x1="3" y1="18" x2="3.01" y2="18"/></svg>
              </span>
              <span class="nav-label">Report Fogging 2026</span>
            </a>
            <a href="#/reports/basecamp" class="nav-item" data-route="/reports/basecamp">
              <span class="nav-icon">
                <svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M3 9l9-7 9 7v11a2 2 0 01-2 2H5a2 2 0 01-2-2z"/></svg>
              </span>
              <span class="nav-label">Rekap Laporan Basecamp</span>
            </a>
          </div>

          <!-- Referensi -->
          <div class="nav-section">
            <span class="nav-section-label">REFERENSI</span>
            <a href="#/sop-fc" class="nav-item" data-route="/sop-fc">
              <span class="nav-icon">
                <svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/><line x1="16" y1="13" x2="8" y2="13"/><line x1="16" y1="17" x2="8" y2="17"/><polyline points="10 9 9 9 8 9"/></svg>
              </span>
              <span class="nav-label">SOP Facility Care</span>
            </a>
            <a href="#/sop-security" class="nav-item" data-route="/sop-security">
              <span class="nav-icon">
                <svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M12 22s8-4 8-10V5l-8-3-8 3v7c0 6 8 10 8 10z"/></svg>
              </span>
              <span class="nav-label">SOP Security</span>
            </a>
            <a href="#/checklist" class="nav-item" data-route="/checklist">
              <span class="nav-icon">
                <svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><polyline points="9 11 12 14 22 4"/><path d="M21 12v7a2 2 0 01-2 2H5a2 2 0 01-2-2V5a2 2 0 012-2h11"/></svg>
              </span>
              <span class="nav-label">Master Document</span>
            </a>
            <a href="#/hygiene-standards" class="nav-item" data-route="/hygiene-standards">
              <span class="nav-icon">
                <svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M12 2l3.09 6.26L22 9.27l-5 4.87 1.18 6.88L12 17.77l-6.18 3.25L7 14.14 2 9.27l6.91-1.01L12 2z"/></svg>
              </span>
              <span class="nav-label">Master Hygiene</span>
            </a>
            <a href="#/dac" class="nav-item" data-route="/dac">
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
            <a href="#/users" class="nav-item" data-route="/users">
              <span class="nav-icon">
                <svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M20 21v-2a4 4 0 00-4-4H8a4 4 0 00-4 4v2"/><circle cx="12" cy="7" r="4"/></svg>
              </span>
              <span class="nav-label">Manajemen User</span>
            </a>
            <a href="#/branches" class="nav-item" data-route="/branches">
              <span class="nav-icon">
                <svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M3 9l9-7 9 7v11a2 2 0 01-2 2H5a2 2 0 01-2-2z"/><polyline points="9 22 9 12 15 12 15 22"/></svg>
              </span>
              <span class="nav-label">Cabang</span>
            </a>
            <a href="#/audit-logs" class="nav-item" data-route="/audit-logs">
              <span class="nav-icon">
                <svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M9 12l2 2 4-4m6 2a9 9 0 11-18 0 9 9 0 0118 0z"/></svg>
              </span>
              <span class="nav-label">Audit Log</span>
            </a>
            <a href="#/settings/import" class="nav-item" data-route="/settings/import">
              <span class="nav-icon">
                <svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M21 15v4a2 2 0 01-2 2H5a2 2 0 01-2-2v-4"/><polyline points="17 8 12 3 7 8"/><line x1="12" y1="3" x2="12" y2="15"/></svg>
              </span>
              <span class="nav-label">Import Data Awal</span>
            </a>
          </div>`:""}
          ${e&&e.role!=="superadmin"&&ie(e,"branches","read")?`
          <div class="nav-section">
            <span class="nav-section-label">MASTER DATA</span>
            <a href="#/branches" class="nav-item" data-route="/branches">
              <span class="nav-icon">
                <svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M3 9l9-7 9 7v11a2 2 0 01-2 2H5a2 2 0 01-2-2z"/><polyline points="9 22 9 12 15 12 15 22"/></svg>
              </span>
              <span class="nav-label">Cabang</span>
            </a>
          </div>`:""}
        </nav>

        <div class="sidebar-footer">
          <div class="sidebar-user">
            <div class="sidebar-avatar">${(e?.full_name||e?.username||"U")[0].toUpperCase()}</div>
            <div class="sidebar-user-info">
              <div class="sidebar-user-name">${e?.full_name||e?.username||"Guest"}</div>
              <div class="sidebar-user-role">${Y(e?.role||"-")}</div>
            </div>
          </div>
          <div style="display:flex; gap:6px; margin-top:8px;">
            <a href="#/profile" class="sidebar-logout nav-item-profile" data-route="/profile" style="flex:1; text-decoration:none;" title="Pengaturan">
              <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><circle cx="12" cy="12" r="3"/><path d="M19.4 15a1.65 1.65 0 00.33 1.82l.06.06a2 2 0 010 2.83 2 2 0 01-2.83 0l-.06-.06a1.65 1.65 0 00-1.82-.33 1.65 1.65 0 00-1 1.51V21a2 2 0 01-2 2 2 2 0 01-2-2v-.09A1.65 1.65 0 009 19.4a1.65 1.65 0 00-1.82.33l-.06.06a2 2 0 01-2.83 0 2 2 0 010-2.83l.06-.06a1.65 1.65 0 00.33-1.82 1.65 1.65 0 00-1.51-1H3a2 2 0 01-2-2 2 2 0 012-2h.09A1.65 1.65 0 004.6 9a1.65 1.65 0 00-.33-1.82l-.06-.06a2 2 0 010-2.83 2 2 0 012.83 0l.06.06a1.65 1.65 0 001.82.33H9a1.65 1.65 0 001-1.51V3a2 2 0 012-2 2 2 0 012 2v.09a1.65 1.65 0 001 1.51 1.65 1.65 0 001.82-.33l.06-.06a2 2 0 012.83 0 2 2 0 010 2.83l-.06.06a1.65 1.65 0 00-.33 1.82V9a1.65 1.65 0 001.51 1H21a2 2 0 012 2 2 2 0 01-2 2h-.09a1.65 1.65 0 00-1.51 1z"/></svg>
              Pengaturan
            </a>
            <button class="sidebar-logout" id="logout-btn" style="flex:1;" title="Keluar">
              <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M9 21H5a2 2 0 01-2-2V5a2 2 0 012-2h4"/><polyline points="16 17 21 12 16 7"/><line x1="21" y1="12" x2="9" y2="12"/></svg>
              Keluar
            </button>
          </div>
        </div>
      </aside>

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
                <span class="topbar-greeting-time">${(()=>{let T=new Date().getHours();return T>=4&&T<11?"Selamat Pagi":T>=11&&T<15?"Selamat Siang":T>=15&&T<18?"Selamat Sore":"Selamat Malam"})()}, </span><span class="topbar-greeting-name">${Y(e?.full_name||e?.username||"Guest")}</span> \u{1F44B}
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
              <div class="topbar-avatar" style="display:flex;align-items:center;justify-content:center;background:#2563EB;color:#fff;font-weight:700;font-size:0.85rem;border-radius:50%;">${Y(String(e?.full_name||e?.username||"G").trim().charAt(0).toUpperCase())}</div>
              <div class="topbar-user-text">
                <span class="topbar-user-name">${Y(e?.full_name||e?.username||"Guest")}</span>
                <span class="topbar-user-role-mini">${Y(e?.role||"-")}</span>
              </div>
              <svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24" style="margin-left:4px;color:var(--text-3)"><polyline points="6 9 12 15 18 9"/></svg>
            </a>
          </div>
        </header>

        <main id="main-content" class="main-content"></main>
      </div>
    </div>
  `;let n=document.getElementById("sidebar"),a=document.getElementById("sidebar-overlay"),o=document.getElementById("topbar-menu-btn"),s=document.getElementById("sidebar-close"),l=()=>{n.classList.add("open"),a.classList.add("show")},h=()=>{n.classList.remove("open"),a.classList.remove("show")};o?.addEventListener("click",l),s?.addEventListener("click",h),a?.addEventListener("click",h),document.querySelectorAll(".nav-item").forEach(T=>T.addEventListener("click",h));let p=new Set(["/users","/settings/import","/audit-logs","/sync"]),c=String(e?.role||"").trim().toLowerCase(),E=ie(e,"read");document.querySelectorAll(".nav-item").forEach(T=>{let u=T.dataset.route;(!E||c!=="superadmin"&&p.has(u))&&(T.style.display="none")}),document.querySelectorAll(".nav-section").forEach(T=>{Array.from(T.querySelectorAll(".nav-item")).some(F=>F.style.display!=="none")||(T.style.display="none")});function C(){let T=window.location.hash.replace("#","")||"/dashboard",[u]=T.split("?");document.querySelectorAll(".nav-item").forEach(I=>{let N=I.dataset.route,w=u===N||N!=="/dashboard"&&(u+"/").startsWith(N+"/");I.classList.toggle("active",w)});let F=document.querySelector(".nav-item-profile");F&&F.classList.toggle("active",u==="/profile");let R=document.getElementById("topbar-title"),K=document.querySelector(".nav-item.active .nav-label");R&&K&&(R.textContent=K.textContent)}window.addEventListener("hashchange",C),C(),Mr(),document.getElementById("btn-fullscreen")?.addEventListener("click",()=>{document.fullscreenElement?document.exitFullscreen?.():document.documentElement.requestFullscreen?.()}),document.getElementById("logout-btn")?.addEventListener("click",async()=>{await X("/api/auth/logout",{method:"POST"}),pt(),Kt&&clearInterval(Kt),qe("/login")}),zr(),Kr(),document.getElementById("btn-notif")?.addEventListener("click",T=>{T.preventDefault(),Nr()})}async function Rr(){if(be("/",()=>qe("/dashboard")),be("/login",({main:e})=>un(e)),be("/dashboard",xe(({main:e})=>dn(e))),be("/calendar",xe(({main:e})=>ni(e))),be("/employees",xe(({main:e,params:t})=>En(e,t))),be("/contracts",xe(({main:e,params:t})=>$n(e,t))),be("/sp",xe(({main:e})=>pi(e))),be("/mutasi",xe(({main:e})=>ui(e))),be("/overtime",xe(({main:e})=>bi(e))),be("/timeline",xe(({main:e,params:t})=>Ln(e,t))),be("/os-migrations",xe(({main:e})=>yi(e))),be("/issues",xe(({main:e,params:t})=>Mn(e,t))),be("/training",xe(({main:e})=>Kn(e))),be("/reports/inspection",xe(({main:e})=>Nn(e))),be("/reports/cleaning",xe(({main:e})=>Rn(e))),be("/reports/fogging",xe(({main:e,params:t})=>jn(e,t))),be("/reports/basecamp",xe(({main:e})=>qn(e))),be("/sop",()=>{window.location.hash="#/sop-fc"}),be("/sop-fc",xe(({main:e})=>Gn(e))),be("/sop-security",xe(({main:e})=>Wn(e))),be("/checklist",xe(({main:e})=>Zn(e))),be("/hygiene-standards",xe(({main:e})=>ei(e))),be("/dac",xe(({main:e})=>ki(e))),be("/users",xe(({main:e})=>{let t=de();if(!t||t.role!=="superadmin"){qe("/dashboard");return}ti(e)})),be("/branches",xe(({main:e})=>ai(e))),be("/profile",xe(({main:e})=>ii(e))),be("/settings/import",xe(({main:e})=>{let t=de();if(!t||t.role!=="superadmin"){qe("/dashboard");return}di(e)})),be("/audit-logs",xe(({main:e})=>{let t=de();if(!t||t.role!=="superadmin"){qe("/dashboard");return}hi(e)})),be("/sync",xe(({main:e})=>{let t=de();if(!t||t.role!=="superadmin"){qe("/dashboard");return}wi(e)})),window.location.hash!=="#/login"){let e=await X("/api/auth/me");if(e.ok&&e.data?.data){let t={...e.data.data},n=typeof t.csrf_token=="string"?t.csrf_token:null;delete t.csrf_token,n&&vt(n),Rt(t),Si()}else pt(),qe("/login")}window.addEventListener("fm:login",()=>{Si(),qe("/dashboard")}),an()}typeof window<"u"&&(window.addEventListener("unhandledrejection",e=>{let t=e?.reason?.message||String(e?.reason||"Unknown promise rejection");console.error("[FCMS Stability] Unhandled promise rejection captured:",t),e.preventDefault()}),window.addEventListener("error",e=>{let t=e?.message||e?.error?.message||"Unknown runtime error";console.error("[FCMS Stability] Runtime error captured:",t)}));Rr().catch(e=>{console.error("[FCMS Stability] App initialization error:",e?.message||e),z("Terjadi kendala saat memulai aplikasi. Silakan muat ulang halaman.")});"serviceWorker"in navigator&&window.location.protocol.startsWith("http")&&window.addEventListener("load",()=>{navigator.serviceWorker.register("/sw.js").then(e=>{console.log("[PWA] Service Worker registered with scope:",e.scope),e.addEventListener("updatefound",()=>{let t=e.installing;t.addEventListener("statechange",()=>{t.state==="installed"&&navigator.serviceWorker.controller&&console.log("[PWA] New version available, reloading on user confirmation or refresh.")})})}).catch(e=>{console.warn("[PWA] Service Worker registration failed:",e)})});var zt=null;window.addEventListener("beforeinstallprompt",e=>{e.preventDefault(),zt=e,window.__showInstallPrompt=()=>{zt&&(zt.prompt(),zt.userChoice.then(t=>{t.outcome==="accepted"&&console.log("[PWA] User accepted the install prompt"),zt=null}))}});
