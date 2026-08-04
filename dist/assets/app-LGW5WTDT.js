var ja=Object.defineProperty;var Fe=(a,e)=>()=>(a&&(e=a(a=0)),e);var Oe=(a,e)=>{for(var s in e)ja(a,s,{get:e[s],enumerable:!0})};var ne={};Oe(ne,{API:()=>aa,apiFetch:()=>f,clearToken:()=>pe,getToken:()=>de,getUser:()=>ae,setToken:()=>Me,setUser:()=>me});function de(){return localStorage.getItem("fm_token")}function Me(a){localStorage.setItem("fm_token",a)}function pe(){localStorage.removeItem("fm_token"),localStorage.removeItem("fm_user")}function ae(){try{return JSON.parse(localStorage.getItem("fm_user")||"null")}catch{return null}}function me(a){localStorage.setItem("fm_user",JSON.stringify(a))}async function f(a,e={}){let s=de(),n={"Content-Type":"application/json",...s?{Authorization:`Bearer ${s}`}:{},...e.headers||{}};try{let t=`cb=${Date.now()}`,c=a.includes("?")?"&":"?",g=`${aa}${a}${c}${t}`,u=await fetch(g,{...e,headers:n}),r;try{let l=await u.text();try{r=JSON.parse(l)}catch{r={error:`Server Error (${u.status}): ${l.substring(0,80)}...`}}}catch{r={error:"Gagal membaca respon dari server"}}return u.status===401&&(pe(),window.location.hash="#/login"),{ok:u.ok,status:u.status,data:r}}catch(t){return{ok:!1,status:0,data:{error:`Koneksi terputus. Periksa jaringan Anda. (${t.message})`}}}}var Ua,aa,I=Fe(()=>{Ua="",aa=Ua});var ia={};Oe(ia,{confirmDialog:()=>Ke,createModal:()=>Y});function Y({title:a,content:e,onConfirm:s,onCancel:n,confirmText:t="Simpan",cancelText:c="Batal",size:g="md",confirmClass:u="btn-primary"}){let r={sm:"400px",md:"560px",lg:"720px",xl:"900px"},l=document.createElement("div");l.className="modal-overlay",l.innerHTML=`
    <div class="modal" style="max-width:${r[g]||r.md}">
      <div class="modal-header">
        <h3 class="modal-title">${a}</h3>
        <button class="modal-close" aria-label="Close">&times;</button>
      </div>
      <div class="modal-body">${typeof e=="string"?e:""}</div>
      <div class="modal-footer">
        <button class="btn btn-ghost modal-cancel">${c}</button>
        ${s?`<button class="btn ${u} modal-confirm">${t}</button>`:""}
      </div>
    </div>
  `,e instanceof HTMLElement&&l.querySelector(".modal-body").appendChild(e);let o=()=>{l.classList.remove("show"),setTimeout(()=>l.remove(),250)};return l.querySelector(".modal-close").addEventListener("click",()=>{n&&n(),o()}),l.querySelector(".modal-cancel").addEventListener("click",()=>{n&&n(),o()}),s&&l.querySelector(".modal-confirm").addEventListener("click",()=>s(l,o)),l.addEventListener("click",i=>{i.target===l&&(n&&n(),o())}),document.body.appendChild(l),requestAnimationFrame(()=>l.classList.add("show")),{overlay:l,close:o}}function Ke(a,e,s="Konfirmasi"){return Y({title:s,content:`<p>${a}</p>`,onConfirm:(n,t)=>{e(),t()},confirmText:"Ya, Lanjutkan",confirmClass:"btn-danger"})}var ge=Fe(()=>{});var G={};Oe(G,{downloadExcel:()=>$,parseExcel:()=>Qe,renderExcelButtons:()=>ze});function Qe(a){return new Promise((e,s)=>{let n=new FileReader;n.onload=t=>{try{let c=new Uint8Array(t.target.result),g=XLSX.read(c,{type:"array"}),u=g.SheetNames[0],r=g.Sheets[u];console.log("--- START EXCEL PARSING ---"),console.log(`File Name: ${a.name}`),console.log(`File Size: ${(a.size/1024).toFixed(2)} KB`),console.log(`File Type: ${a.type||"unknown"}`),console.log(`Sheets Found: ${g.SheetNames.join(", ")}`),console.log(`Sheet Used: ${u}`);let l=XLSX.utils.decode_range(r["!ref"]||"A1:A1"),o=l.e.r-l.s.r+1,i=l.e.c-l.s.c+1;console.log(`Total Rows (including empty): ${o}`),console.log(`Total Columns: ${i}`);let y=[];for(let p=l.s.c;p<=l.e.c;++p){let d=r[XLSX.utils.encode_cell({c:p,r:l.s.r})];d&&d.v&&y.push(d.v)}console.log(`Headers Found: ${y.join(", ")}`),console.log("---------------------------");let b=XLSX.utils.sheet_to_json(r,{defval:""});Object.defineProperty(b,"__worksheet",{value:r,enumerable:!1}),Object.defineProperty(b,"__headers",{value:y,enumerable:!1}),e(b)}catch(c){s(c)}},n.onerror=t=>s(t),n.readAsArrayBuffer(a)})}function $(a,e){try{let s=XLSX.utils.json_to_sheet(a),n=XLSX.utils.book_new();XLSX.utils.book_append_sheet(n,s,"Data"),XLSX.writeFile(n,`${e}.xlsx`)}catch(s){throw console.error("Error generating Excel file:",s),s}}function ze(a){return`
    <div class="excel-actions" style="display:flex;gap:0.5rem;margin-bottom:1rem;">
      <button class="btn btn-secondary btn-sm" id="btn-export-${a}">
        \u{1F4E5} Export Excel
      </button>
      <button class="btn btn-secondary btn-sm" id="btn-template-${a}">
        \u{1F4C4} Download Template
      </button>
      <label class="btn btn-primary btn-sm" style="cursor:pointer;margin:0;" id="label-import-${a}">
        <span class="import-text">\u{1F4E4} Import Excel</span>
        <input type="file" id="input-import-${a}" accept=".xlsx, .xls, .csv" style="display:none;">
      </label>
    </div>
  `}var B=Fe(()=>{});I();var Re={},Te=null;function A(a,e){Re[a]=e}function ie(a){window.location.hash=a}function ta(){async function a(){let e=window.location.hash.replace("#","")||"/dashboard",[s,...n]=e.split("?"),t=Re[s];if(!t){for(let[g,u]of Object.entries(Re))if(g.endsWith("/*")&&s.startsWith(g.slice(0,-2))){t=u;break}}Te&&(Te(),Te=null);let c=document.getElementById("main-content");if(c&&(c.innerHTML='<div class="loading-spinner"><div class="spinner"></div></div>'),t){let g=new URLSearchParams(n.join("?")),u=s.split("/").filter(Boolean),r=await t({path:s,params:g,segments:u,main:c});r&&(Te=r)}else{let g=c||document.getElementById("app");g&&(g.innerHTML='<div class="empty-state"><h2>404 - Halaman tidak ditemukan</h2></div>')}}window.addEventListener("hashchange",a),a()}var ue;function Ja(){return ue||(ue=document.createElement("div"),ue.id="toast-container",document.body.appendChild(ue)),ue}function na(a,e="info",s=3500){let n=Ja(),t=document.createElement("div");t.className=`toast toast-${e}`;let c={success:"\u2713",error:"\u2715",warning:"\u26A0",info:"\u2139"};t.innerHTML=`<span class="toast-icon">${c[e]||"\u2139"}</span><span class="toast-msg">${a}</span>`,n.appendChild(t),requestAnimationFrame(()=>t.classList.add("show")),setTimeout(()=>{t.classList.remove("show"),setTimeout(()=>t.remove(),350)},s)}var q=a=>na(a,"success"),U=a=>na(a,"error");ge();I();var ee={};function be(a){if(ee[a]){try{ee[a].destroy()}catch{}delete ee[a]}}function Ga(){Object.keys(ee).forEach(be)}var X=(a,e=0)=>{let s=Number(a);return isNaN(s)||a===null||a===void 0?e:s},se=(a,e="\u2014")=>{if(a==null||a==="")return e;let s=String(a).trim();return s===""||s==="[object Object]"?e:s};var Qa=a=>{if(!a||typeof a!="string")return"";try{let[e,s]=a.split("-");return new Date(Number(e),Number(s)-1).toLocaleDateString("id-ID",{month:"short",year:"2-digit"})}catch{return a}};function ra(a,e,s=900){if(!a)return;let n=Math.max(0,Math.round(X(e)));if(n===0){a.textContent="0";return}let t=Date.now(),c=()=>{let g=Math.min((Date.now()-t)/s,1),u=1-Math.pow(1-g,3);a.textContent=Math.round(u*n).toLocaleString("id-ID"),g<1?requestAnimationFrame(c):a.textContent=n.toLocaleString("id-ID")};requestAnimationFrame(c)}var za={Done:"pill-success",Aktif:"pill-success",Selesai:"pill-success",Open:"pill-danger",Pending:"pill-warning","In Progress":"pill-info","Tidak Aktif":"pill-neutral",Resign:"pill-neutral",Cut:"pill-neutral"},Va=a=>{let e=se(a,"\u2014");return`<span class="status-pill ${za[e]||"pill-neutral"}">${e}</span>`};var Q={family:"Inter",size:11},le="#94A3B8",$e="#F1F5F9",qe=["#2563EB","#10B981","#F59E0B","#EF4444","#8B5CF6","#0EA5E9","#F97316","#14B8A6","#6366F1","#EC4899"],Wa=()=>window.innerWidth<768;function Ue(a={}){return{responsive:!0,maintainAspectRatio:!1,animation:{duration:700,easing:"easeOutQuart"},plugins:{legend:{position:Wa()?"bottom":"top",labels:{font:Q,color:"#64748B",usePointStyle:!0,padding:10,boxWidth:8,boxHeight:8}},tooltip:{mode:"index",intersect:!1,bodyFont:Q,titleFont:{...Q,weight:"700"}}},scales:{x:{grid:{color:$e},ticks:{font:Q,color:le,maxRotation:0}},y:{grid:{color:$e},ticks:{font:Q,color:le},beginAtZero:!0}},...a}}var Ya=()=>Array(5).fill(0).map(()=>`
  <div class="kpi-card" style="pointer-events:none;padding:16px">
    <div style="display:flex; gap:16px; align-items:flex-start">
      <div class="skeleton" style="width:48px;height:48px;border-radius:12px;flex-shrink:0"></div>
      <div style="flex:1">
        <div class="skeleton skeleton-text" style="width:40px;height:24px;margin-bottom:6px"></div>
        <div class="skeleton skeleton-text" style="width:80px;height:12px;margin-bottom:4px"></div>
        <div class="skeleton skeleton-text" style="width:100px;height:10px"></div>
      </div>
    </div>
    <div style="display:flex; align-items:flex-end; gap:8px; margin-top:16px">
      <div class="skeleton" style="flex:1;height:24px;border-radius:4px"></div>
      <div class="skeleton skeleton-text" style="width:30px;height:12px"></div>
    </div>
  </div>`).join(""),Xa=()=>Array(7).fill(0).map(()=>`
  <div class="mini-stat" style="pointer-events:none">
    <div class="skeleton" style="width:40px;height:40px;border-radius:10px;flex-shrink:0"></div>
    <div style="flex:1">
      <div class="skeleton skeleton-text" style="width:45%;height:22px;margin-bottom:5px"></div>
      <div class="skeleton skeleton-text" style="width:80%;height:11px"></div>
    </div>
  </div>`).join("");function He(a=3){return Array(a).fill(0).map((e,s)=>`<div class="skeleton skeleton-text" style="height:38px;margin-bottom:${s<a-1?"6px":"0"};border-radius:6px"></div>`).join("")}async function re(a,e,s=8e3){try{let n=new AbortController,t=setTimeout(()=>n.abort(),s),c=await f(a,{signal:n.signal}).catch(()=>null);if(clearTimeout(t),!c||!c.ok)return e;let g=c.data;return g?g.data!==void 0?g.data??e:g:e}catch{return e}}function Za(){["skel-donut","skel-trend","skel-insp","skel-contract"].forEach(n=>{let t=document.getElementById(n);t&&(t.style.display="none")}),["chart-donut","chart-trend","chart-insp","chart-contract"].forEach(n=>{let t=document.getElementById(n);if(t&&t.style.display==="none"){t.style.display="block";let c=t.parentElement;if(c&&!c.querySelector(".chart-empty")){let g=document.createElement("div");g.className="chart-empty",g.textContent="Belum ada data",t.style.display="none",c.appendChild(g)}}});let a=document.getElementById("kpi-row");a&&a.querySelector(".skeleton")&&la({});let e=document.getElementById("mini-stats-row");e&&e.querySelector(".skeleton")&&oa({}),["table-contracts","table-issues"].forEach(n=>{let t=document.getElementById(n);t&&t.querySelector(".skeleton")&&(t.innerHTML='<div class="chart-empty">Belum ada data</div>')});let s=document.getElementById("activity-log");s&&s.querySelector(".skeleton")&&(s.innerHTML='<div class="chart-empty">Belum ada aktivitas</div>')}async function sa(a){Ga(),a._dashRefresh&&clearInterval(a._dashRefresh),a._skelTimeout&&clearTimeout(a._skelTimeout),a.innerHTML=`
    <div class="dashboard-wrap" id="dash-root">


      <!-- KPI -->
      <div class="kpi-row" id="kpi-row">${Ya()}</div>

      <!-- Mini Stats -->
      <div class="mini-stats-row" id="mini-stats-row">${Xa()}</div>

      <!-- Charts Row -->
      <div class="charts-row">
        <div class="chart-card">
          <div class="chart-card-header">
            <div class="chart-card-title">Permasalahan per Kategori</div>
          </div>
          <div style="display:flex; gap:20px; align-items:center; height:140px">
            <div class="chart-canvas-wrap" style="flex:1;height:100%;position:relative">
              <div id="skel-donut" class="skeleton" style="position:absolute;inset:0;border-radius:12px"></div>
              <canvas id="chart-donut" style="display:none"></canvas>
            </div>
            <div id="donut-legend" class="donut-legend" style="width:110px"></div>
          </div>
          <div style="text-align:center; font-size:0.75rem; color:var(--text-3); margin-top:16px">
            Periode: 22 Juni - 22 Juli 2026
          </div>
        </div>
        <div class="chart-card">
          <div class="chart-card-header">
            <div class="chart-card-title">Trend Permasalahan 12 Bulan</div>
            <div style="display:flex;align-items:center;gap:16px;font-size:0.75rem;font-weight:600;color:var(--text-2)">
               <div style="display:flex;align-items:center;gap:6px"><div style="width:16px;height:8px;border:2px solid #EF4444;border-radius:2px"></div> Open</div>
               <div style="display:flex;align-items:center;gap:6px"><div style="width:16px;height:8px;border:2px solid #10B981;border-radius:2px"></div> Closed</div>
            </div>
          </div>
          <div class="chart-canvas-wrap" style="height:140px;position:relative">
            <div id="skel-trend" class="skeleton" style="position:absolute;inset:0;border-radius:12px"></div>
            <canvas id="chart-trend" style="display:none"></canvas>
          </div>
        </div>
        <div class="chart-card">
          <div class="chart-card-header" style="align-items:flex-start">
            <div>
              <a href="#/reports/inspection" class="chart-card-title" style="text-decoration:none; display:inline-block">Rata-rata Skor Inspeksi per Cabang <span style="font-size:0.8rem; color:var(--primary); font-weight:600; margin-left:8px">Lihat Laporan &rarr;</span></a>
              <div class="chart-card-subtitle" style="font-size:0.65rem">Skor rata-rata SCM & Cleaning</div>
            </div>
            <select class="btn-ghost" style="padding:4px;font-size:0.7rem;border:1px solid var(--border);border-radius:4px;cursor:pointer"><option>Bulan Ini</option></select>
          </div>
          <div class="chart-canvas-wrap" style="height:140px;position:relative;margin-top:10px">
            <div id="skel-insp" class="skeleton" style="position:absolute;inset:0;border-radius:12px"></div>
            <canvas id="chart-insp" style="display:none"></canvas>
          </div>
        </div>
      </div>

      <!-- Bottom Row -->
      <div class="bottom-row" style="margin-top:24px;">
        <!-- Jadwal Hari Ini -->
        <div class="chart-card">
          <div class="chart-card-header">
            <div class="chart-card-title">Jadwal Hari Ini</div>
            <a href="#/calendar" class="chart-link">Lihat Kalender</a>
          </div>
          <div id="widget-agenda" class="dash-table-wrap" style="height:160px;overflow-y:auto">${He(3)}</div>
        </div>
        <!-- KPI Kebersihan -->
        <div class="chart-card">
          <div class="chart-card-header">
            <div class="chart-card-title">KPI Kebersihan</div>
          </div>
          <div id="widget-kpi-kebersihan" style="margin-top:0px">${He(4)}</div>
        </div>
        <!-- Permasalahan Terbaru -->
        <div class="chart-card">
          <div class="chart-card-header">
            <div class="chart-card-title">Permasalahan Terbaru</div>
            <a href="#/issues" class="chart-link">Lihat Semua</a>
          </div>
          <div id="table-issues" class="dash-table-wrap" style="height:160px;overflow-y:auto">${He(3)}</div>
        </div>
        <!-- Kontrak Akan Habis -->
        <div class="chart-card">
          <div class="chart-card-header">
            <div class="chart-card-title">Kontrak Akan Habis</div>
            <a href="#/contracts" class="chart-link">Lihat Data</a>
          </div>
          <div class="chart-canvas-wrap" style="height:160px;position:relative;margin-top:10px">
            <div id="skel-contract-mini" class="skeleton" style="position:absolute;inset:0;border-radius:12px"></div>
            <canvas id="chart-contract-mini" style="display:none"></canvas>
          </div>
        </div>
      </div>

      <!-- Quick Actions Row -->
      <div class="actions-wrap">
        <div class="actions-title">Aksi Cepat</div>
        <div class="actions-row" id="quick-actions">
          <!-- Rendered in JS -->
        </div>
      </div>



    </div>
  `,document.getElementById("btn-dash-refresh")?.addEventListener("click",()=>je(a)),a._skelTimeout=setTimeout(()=>Za(),5e3),await je(a),a._dashRefresh=setInterval(()=>{document.getElementById("dash-root")?je(a):clearInterval(a._dashRefresh)},6e4)}async function je(a){a._skelTimeout&&(clearTimeout(a._skelTimeout),a._skelTimeout=null);let[e,s,n,t,c,g]=await Promise.all([re("/api/dashboard/kpi",{},8e3),re("/api/dashboard/issues-trend",{},8e3),re("/api/dashboard/issues-summary",{},8e3),re("/api/dashboard/inspection-bar",{},8e3),re("/api/dashboard/stats",{},8e3),re("/api/dashboard/calendar",[],8e3)]);try{la(e)}catch(u){console.warn("KPI render:",u)}try{oa(e)}catch(u){console.warn("MiniStats render:",u)}try{et(Array.isArray(n?.by_category)?n.by_category:[])}catch(u){console.warn("Donut render:",u),te("skel-donut","chart-donut")}try{at(s)}catch(u){console.warn("Trend render:",u),te("skel-trend","chart-trend")}try{tt(t)}catch(u){console.warn("InspBar render:",u),te("skel-insp","chart-insp")}try{let u=Array.isArray(c)?c:Array.isArray(c?.recent_issues)?c.recent_issues:[];it(u)}catch(u){console.warn("IssuesTable render:",u)}try{let u=Array.isArray(c?.expiring_contracts)?c.expiring_contracts:[];nt()}catch(u){console.warn("ContractsTable render:",u)}try{rt(Array.isArray(g)?g:[])}catch(u){console.warn("Agenda render:",u)}try{st(e)}catch(u){console.warn("KPI Kebersihan render:",u)}try{lt()}catch(u){console.warn("Quick Actions render:",u)}}function la(a){let e=document.getElementById("kpi-row");if(!e)return;a=a||{};let s=[{icon:"\u{1F465}",label:"Karyawan Aktif",sub:"Total karyawan aktif",href:"#/employees",color:"kpi-blue",key:"employees",trendPct:"+2%",trendColor:"#10B981",points:"0,20 10,18 20,22 30,12 40,15 50,8 60,10 70,5 80,6 90,2 100,0"},{icon:"\u{1F4C4}",label:"Kontrak Aktif",sub:"Kontrak yang masih berjalan",href:"#/contracts",color:"kpi-green",key:"contracts",trendPct:"+1%",trendColor:"#10B981",points:"0,15 20,18 40,10 60,12 80,5 100,2"},{icon:"\u23F3",label:"Kontrak Habis 30 Hari",sub:"Akan segera berakhir",href:"#/contracts",color:"kpi-warn",key:"expiring30",trendPct:"+25%",trendColor:"#F59E0B",points:"0,25 20,22 40,24 60,15 80,18 100,5"},{icon:"\u26A0\uFE0F",label:"Permasalahan Open",sub:"Belum diselesaikan",href:"#/issues",color:"kpi-red",key:"issues",trendPct:"0%",trendColor:"#EF4444",points:"0,20 20,18 40,22 60,19 80,21 100,20"},{icon:"\u{1F4AC}",label:"One on One Pending",sub:"Menunggu tindak lanjut",href:"#/one-on-one",color:"kpi-purple",key:"one_on_one",trendPct:"+8%",trendColor:"#10B981",points:"0,25 20,15 40,18 60,8 80,10 100,2"}];e.innerHTML=s.map(n=>{let t=X(a[n.key]?.current,0);return`
      <a href="${n.href}" class="kpi-card ${n.color}" style="text-decoration:none;padding:12px 16px">
        <div style="display:flex; gap:16px; align-items:center;">
          <div class="kpi-icon-wrap" style="width:48px;height:48px;border-radius:12px;display:flex;align-items:center;justify-content:center;flex-shrink:0"><span class="kpi-icon-emoji">${n.icon}</span></div>
          <div style="flex:1;min-width:0;">
            <div class="kpi-value" data-target="${t}" style="font-size:1.8rem; font-weight:800; line-height:1; color:var(--text-1)">${t}</div>
            <div class="kpi-label" style="font-size:0.85rem; font-weight:700; color:var(--text-2); margin-top:6px">${n.label}</div>
            <div class="kpi-subtitle" style="font-size:0.7rem; color:var(--text-3); margin-top:2px; white-space:nowrap; overflow:hidden; text-overflow:ellipsis">${n.sub}</div>
          </div>
        </div>
      </a>`}).join(""),e.querySelectorAll(".kpi-value").forEach(n=>{ra(n,parseInt(n.dataset.target)||0)})}function oa(a){let e=document.getElementById("mini-stats-row");if(!e)return;a=a||{};let s=[{icon:"\u{1F4C5}",label:"Jadwal",val:a.schedule?.current,href:"#/timeline",color:"mini-blue"},{icon:"\u{1F393}",label:"Training",val:a.training_month?.current,href:"#/training",color:"mini-gray"},{icon:"\u{1F4E6}",label:"Permintaan",val:a.supply?.current,href:"#/reports/supply",color:"mini-orange"},{icon:"\u{1F50D}",label:"Inspeksi",val:a.inspection_month?.current,href:"#/reports/inspection",color:"mini-blue"},{icon:"\u{1F9F9}",label:"GCDC",val:a.cleaning_month?.current,href:"#/reports/cleaning",color:"mini-green"},{icon:"\u{1F4A8}",label:"Fogging",val:a.fogging_month?.current,href:"#/reports/fogging",color:"mini-purple"},{icon:"\u{1F3E2}",label:"Cabang",val:a.branches?.current,href:"#/branches",color:"mini-teal"}];e.innerHTML=s.map(n=>`
    <a href="${n.href}" class="mini-stat ${n.color}" style="text-decoration:none">
      <div class="mini-stat-icon">${n.icon}</div>
      <div class="mini-stat-body">
        <div class="mini-stat-value" data-target="${X(n.val)}">0</div>
        <div class="mini-stat-text">${n.label}</div>
      </div>
    </a>`).join(""),e.querySelectorAll(".mini-stat-value").forEach(n=>ra(n,parseInt(n.dataset.target)||0,700))}function et(a){te("skel-donut","chart-donut");let e=document.getElementById("chart-donut"),s=document.getElementById("donut-legend");if(!e||!s)return;be("donut");let n=(a||[]).filter(r=>X(r.count)>0);if(!n.length){Je(e,"Belum ada data permasalahan");return}let t=n.map(r=>`${se(r.category,"Lainnya")}`),c=n.map(r=>X(r.count)),g=c.reduce((r,l)=>r+l,0);s.innerHTML=n.map((r,l)=>{let o=qe[l%qe.length],i=g>0?Math.round(r.count/g*100):0;return`
      <div class="donut-legend-item">
        <div class="donut-legend-color" style="background:${o}"></div>
        <div>
          <div class="donut-legend-val"><span style="color:var(--text-1)">${r.count}</span> <span style="font-size:0.7rem;font-weight:600;color:var(--text-3)">(${i}%)</span></div>
          <div class="donut-legend-label">${t[l]}</div>
        </div>
      </div>
    `}).join("");let u={id:"centerText",beforeDraw:function(r){let l=r.width,o=r.height,i=r.ctx;i.restore();let y=(o/80).toFixed(2);i.font="bold "+y+"em Inter",i.textBaseline="middle",i.fillStyle="#1E293B";let b=g.toString(),p=Math.round((l-i.measureText(b).width)/2),d=o/2;i.fillText(b,p,d-10),i.font="600 "+(y*.35).toFixed(2)+"em Inter",i.fillStyle="#64748B";let m="Total",h=Math.round((l-i.measureText(m).width)/2);i.fillText(m,h,d+15),i.save()}};ee.donut=new Chart(e,{type:"doughnut",data:{labels:t,datasets:[{data:c,backgroundColor:qe,borderWidth:2,borderColor:"#fff",hoverBorderColor:"#fff"}]},options:{responsive:!0,maintainAspectRatio:!1,animation:{duration:700},plugins:{legend:{display:!1},tooltip:{bodyFont:Q,titleFont:{...Q,weight:"700"},callbacks:{label:r=>` ${r.label}: ${r.parsed} kasus`}}},cutout:"75%"},plugins:[u]})}function at(a){te("skel-trend","chart-trend");let e=document.getElementById("chart-trend");if(!e)return;be("trend"),a=a||{};let s=(a.labels||[]).map(Qa),n=(a.open||[]).map(c=>X(c)),t=(a.closed||[]).map(c=>X(c));if(!s.length){Je(e,"Belum ada data trend");return}ee.trend=new Chart(e,{type:"line",data:{labels:s,datasets:[{label:"Open",data:n,borderColor:"#EF4444",backgroundColor:"rgba(239,68,68,.08)",fill:!0,tension:.4,pointRadius:2,pointHoverRadius:4,pointBackgroundColor:"#EF4444",borderWidth:2},{label:"Closed",data:t,borderColor:"#10B981",backgroundColor:"rgba(16,185,129,.08)",fill:!0,tension:.4,pointRadius:2,pointHoverRadius:4,pointBackgroundColor:"#10B981",borderWidth:2}]},options:Ue({plugins:{legend:{display:!1}}})})}function tt(a){te("skel-insp","chart-insp");let e=document.getElementById("chart-insp");if(!e)return;be("inspBar"),a=a||{};let s=a.labels||[],n=(a.fc||[]).map(c=>X(c)),t=(a.spv||[]).map(c=>X(c));if(!s.length){Je(e,"Belum ada data inspeksi");return}ee.inspBar=new Chart(e,{type:"bar",data:{labels:s,datasets:[{label:"Skor FC",data:n,backgroundColor:"rgba(37,99,235,.75)",borderRadius:4,borderSkipped:!1},{label:"Skor SPV",data:t,backgroundColor:"rgba(16,185,129,.75)",borderRadius:4,borderSkipped:!1}]},options:Ue({plugins:{legend:{position:"top"}},scales:{x:{grid:{display:!1},ticks:{font:Q,color:le,maxRotation:45,minRotation:30}},y:{grid:{color:$e},ticks:{font:Q,color:le},min:0,max:100}}})})}function nt(){te("skel-contract-mini","chart-contract-mini");let a=document.getElementById("chart-contract-mini");if(!a)return;be("contractMiniBar");let e=["Jun","Jul","Agu","Sep","Okt","Nov","Des"],s=[12,18,9,24,15,30,42],t=a.getContext("2d").createLinearGradient(0,0,0,200);t.addColorStop(0,"#60A5FA"),t.addColorStop(1,"#2563EB"),ee.contractMiniBar=new Chart(a,{type:"bar",data:{labels:e,datasets:[{label:"Kontrak Habis",data:s,backgroundColor:t,borderRadius:4,borderSkipped:!1,barPercentage:.6,categoryPercentage:.7}]},options:Ue({plugins:{legend:{display:!1}},scales:{x:{grid:{display:!1},ticks:{font:Q,color:le,maxRotation:0}},y:{grid:{color:$e,borderDash:[4,4],drawBorder:!1},ticks:{font:Q,color:le,precision:0,maxTicksLimit:5},min:0}},animation:{y:{duration:1e3,easing:"easeOutQuart"}}})})}function it(a){let e=document.getElementById("table-issues");if(!e)return;let s=(a||[]).slice(0,8);if(!s.length){e.innerHTML='<div class="chart-empty">\u2705 Tidak ada permasalahan terbuka</div>';return}e.innerHTML=`
    <div class="dash-list">
      ${s.map(n=>`
        <div class="dash-list-item">
          <div style="flex-shrink:0">${Va(n.status)}</div>
          <div style="flex:1;min-width:0;margin-left:4px">
            <div style="font-size:0.85rem;font-weight:700;color:var(--text-1);display:-webkit-box;-webkit-line-clamp:2;-webkit-box-orient:vertical;overflow:hidden;line-height:1.3">${se(n.complaint)}</div>
            <div style="font-size:0.75rem;color:var(--text-3);margin-top:2px">${se(n.branch_name)}</div>
          </div>
        </div>
      `).join("")}
    </div>`}function rt(a){let e=document.getElementById("widget-agenda");if(!e)return;let s=(a||[]).slice(0,10);if(!s.length){e.innerHTML='<div class="chart-empty">\u2705 Tidak ada agenda hari ini</div>';return}e.innerHTML=`
    <div style="display:flex;flex-direction:column;gap:12px;padding-right:8px">
      ${s.map(n=>{let t="#3B82F6",c="#EFF6FF",g="Agenda",u=(n.title||"").toLowerCase();return u.includes("inspeksi")?(t="#10B981",c="#ECFDF5",g="Inspeksi"):u.includes("cleaning")||u.includes("gcdc")?(t="#3B82F6",c="#EFF6FF",g="Cleaning"):u.includes("reliefer")?(t="#F59E0B",c="#FFFBEB",g="Reliefer"):u.includes("fogging")&&(t="#8B5CF6",c="#F5F3FF",g="Fogging"),`
        <div style="display:flex;gap:16px;align-items:flex-start;padding-bottom:12px;border-bottom:1px solid var(--border)">
          <div style="font-size:0.85rem;font-weight:700;color:var(--text-1);margin-top:2px">${new Date(n.event_date).toLocaleTimeString("id-ID",{hour:"2-digit",minute:"2-digit"})}</div>
          <div style="width:8px;height:8px;border-radius:50%;background:${t};margin-top:6px;flex-shrink:0"></div>
          <div style="flex:1">
            <div style="font-weight:700;font-size:0.85rem;color:var(--text-1);line-height:1.2;margin:0 0 4px 0">${se(n.title)}</div>
            <div style="font-size:0.75rem;color:var(--text-3)">${se(n.branch_name)}</div>
          </div>
          <div style="font-size:0.7rem;font-weight:600;padding:2px 8px;border-radius:6px;background:${c};color:${t}">${g}</div>
        </div>
      `}).join("")}
    </div>
  `}function st(a){let e=document.getElementById("widget-kpi-kebersihan");if(!e)return;let s=[{label:"Kebersihan Area",val:"97%",target:"Target 95%",icon:"\u{1F9F9}",bg:"#ECFDF5",color:"#10B981"},{label:"Penyelesaian Complaint",val:"100%",target:"Target 100%",icon:"\u23F1\uFE0F",bg:"#ECFDF5",color:"#10B981"},{label:"Kepatuhan Jadwal Cleaning",val:"99%",target:"Target 100%",icon:"\u23F1\uFE0F",bg:"#EFF6FF",color:"#3B82F6"},{label:"Kepatuhan GCDC",val:"100%",target:"Target 100%",icon:"\u{1F9F9}",bg:"#EFF6FF",color:"#3B82F6"},{label:"Complaint Cleaning (\u226410)",val:"4",target:"Target \u226410",icon:"\u{1F4DD}",bg:"#F5F3FF",color:"#8B5CF6"},{label:"Pelaksanaan Fogging",val:"100%",target:"Target 100%",icon:"\u{1F4A8}",bg:"#F5F3FF",color:"#8B5CF6"}];e.innerHTML=`
    <div style="display:grid;grid-template-columns:1fr;gap:12px;padding-bottom:12px;height:160px;overflow-y:auto;padding-right:8px;">
      ${s.map(n=>{let t=n.val.includes("%")?parseInt(n.val):Math.min(100,parseInt(n.val)*10);return`
        <div class="prog-item">
          <div class="prog-header">
            <div class="prog-title">
              <div class="prog-title-icon" style="background:${n.bg};color:${n.color}">${n.icon}</div>
              ${n.label}
            </div>
            <div class="prog-val">${n.val}</div>
          </div>
          <span class="prog-target">${n.target}</span>
          <div class="prog-bar-bg">
            <div class="prog-bar-fill" style="width:${t}%;background:${n.color}"></div>
          </div>
        </div>
      `}).join("")}
    </div>
  `}function lt(){let a=document.getElementById("quick-actions");if(!a)return;let e=[{label:"Buat Permasalahan",icon:'<svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24"><path d="M12 5v14m-7-7h14"/></svg>',bg:"#3B82F6",href:"#/issues"},{label:"Permintaan Barang",icon:"\u{1F4E6}",bg:"#10B981",href:"#/reports/supply"},{label:"One on One Baru",icon:"\u{1F465}",bg:"#6366F1",href:"#/one-on-one"},{label:"Input Kegiatan",icon:"\u{1F4CB}",bg:"#8B5CF6",href:"#/timeline"},{label:"Buat Checklist",icon:"\u{1F4DD}",bg:"#0EA5E9",href:"#/checklist"},{label:"Laporan Basecamp",icon:"\u{1F4CA}",bg:"#14B8A6",href:"#/reports/basecamp"},{label:"Kalender",icon:"\u{1F4C5}",bg:"#8B5CF6",href:"#/calendar"}];a.innerHTML=e.map(s=>`
    <a href="${s.href}" class="action-btn">
      <div class="action-icon" style="background:${s.bg}">${s.icon}</div>
      ${s.label}
    </a>
  `).join("")}function te(a,e){let s=document.getElementById(a),n=document.getElementById(e);s&&(s.style.display="none",s.style.position=""),n&&(n.style.display="block")}function Je(a,e="Belum ada data"){if(!a)return;a.style.display="none";let s=a.parentElement;if(!s)return;if(!s.querySelector(".chart-empty")){let t=document.createElement("div");t.className="chart-empty",t.textContent=e,s.appendChild(t)}}I();async function ca(a){document.getElementById("app").innerHTML=`
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
  `;let e=document.getElementById("login-form"),s=document.getElementById("login-error"),n=document.getElementById("login-btn"),t=document.getElementById("toggle-password"),c=document.getElementById("login-password");t?.addEventListener("click",()=>{let g=c.type==="text";c.type=g?"password":"text",t.style.color=g?"":"var(--primary)"}),e?.addEventListener("submit",async g=>{g.preventDefault(),s.style.display="none";let u=e.username.value.trim(),r=e.password.value;if(!u||!r){s.textContent="Username dan password wajib diisi.",s.style.display="block";return}n.querySelector(".btn-text").style.display="none",n.querySelector(".btn-spinner").style.display="",n.disabled=!0;try{let l=await f("/api/auth/login",{method:"POST",body:JSON.stringify({username:u,password:r})});l.ok&&l.data.success?(Me(l.data.data.token),me(l.data.data.user),q("Login berhasil! Selamat datang \u{1F44B}"),window.dispatchEvent(new Event("fm:login"))):(s.textContent=l.data.error||"Username atau password salah.",s.style.display="block",n.classList.add("shake"),setTimeout(()=>n.classList.remove("shake"),600))}catch{s.textContent="Gagal terhubung ke server. Periksa koneksi internet.",s.style.display="block"}finally{n.querySelector(".btn-text").style.display="",n.querySelector(".btn-spinner").style.display="none",n.disabled=!1}})}I();function da({columns:a,data:e,onEdit:s,onDelete:n,onView:t,actions:c=[],emptyText:g="Tidak ada data",bulkSelect:u=null}){let r=document.createElement("div");if(r.className="table-wrapper",!e||e.length===0)return r.innerHTML=`<div class="empty-state"><p>${g}</p></div>`,r;let l=document.createElement("table");l.className="data-table";let o=document.createElement("thead"),i=document.createElement("tr");if(u){let b=document.createElement("th");b.style.width="40px",b.style.textAlign="center";let p=document.createElement("input");p.type="checkbox",p.id="select-all-checkbox",p.title="Pilih semua",p.addEventListener("change",()=>{e.forEach(d=>{p.checked?u.selectedIds.add(d.id):u.selectedIds.delete(d.id)}),r.querySelectorAll(".row-checkbox").forEach(d=>d.checked=p.checked),u.onToggle()}),b.appendChild(p),i.appendChild(b)}if(a.forEach(b=>{let p=document.createElement("th");p.textContent=b.label,b.width&&(p.style.width=b.width),i.appendChild(p)}),s||n||t||c.length>0){let b=document.createElement("th");b.textContent="Aksi",b.style.width="120px",i.appendChild(b)}o.appendChild(i),l.appendChild(o);let y=document.createElement("tbody");return e.forEach(b=>{let p=document.createElement("tr");if(u){let d=document.createElement("td");d.style.textAlign="center",d.style.width="40px";let m=document.createElement("input");m.type="checkbox",m.className="row-checkbox",m.checked=u.selectedIds.has(b.id),m.addEventListener("change",()=>{if(m.checked)u.selectedIds.add(b.id);else{u.selectedIds.delete(b.id);let h=document.getElementById("select-all-checkbox");h&&(h.checked=!1)}u.onToggle()}),d.appendChild(m),p.appendChild(d)}if(a.forEach(d=>{let m=document.createElement("td");if(d.render){let h=d.render(b[d.key],b);h instanceof HTMLElement?m.appendChild(h):m.innerHTML=h||""}else m.textContent=b[d.key]!==null&&b[d.key]!==void 0&&b[d.key]!==""?b[d.key]:"";d.nowrap&&(m.style.whiteSpace="nowrap"),p.appendChild(m)}),s||n||t||c.length>0){let d=document.createElement("td");d.className="actions-cell";let m=document.createElement("div");if(m.className="btn-group",t){let h=document.createElement("button");h.className="btn btn-xs btn-ghost",h.innerHTML="\u{1F441}",h.title="Lihat",h.addEventListener("click",()=>t(b)),m.appendChild(h)}if(s){let h=document.createElement("button");h.className="btn btn-xs btn-secondary",h.innerHTML="\u270F\uFE0F",h.title="Edit",h.addEventListener("click",()=>s(b)),m.appendChild(h)}c.forEach(h=>{let v=document.createElement("button");v.className=`btn btn-xs ${h.class||"btn-ghost"}`,v.innerHTML=h.icon||h.label,v.title=h.label,v.addEventListener("click",()=>h.handler(b)),m.appendChild(v)}),d.appendChild(m),p.appendChild(d)}y.appendChild(p)}),l.appendChild(y),r.appendChild(l),r}function pa({page:a,pages:e,total:s,limit:n,onPage:t}){if(e<=1)return null;let c=document.createElement("div");c.className="pagination";let g=document.createElement("span");g.className="pagination-info",g.textContent=`Total: ${s} data`,c.appendChild(g);let u=document.createElement("div");u.className="pagination-btns";let r=(i,y,b=!1,p=!1)=>{let d=document.createElement("button");d.className=`btn btn-sm ${p?"btn-primary":"btn-ghost"} pagination-btn`,d.textContent=i,d.disabled=b,d.addEventListener("click",()=>t(y)),u.appendChild(d)};r("\xAB",1,a===1),r("\u2039",a-1,a===1);let l=Math.max(1,a-2),o=Math.min(e,a+2);for(let i=l;i<=o;i++)r(i,i,!1,i===a);return r("\u203A",a+1,a===e),r("\xBB",e,a===e),c.appendChild(u),c}ge();function Ge(a){return a.map(e=>{if(e.type==="hidden")return`<input type="hidden" name="${e.name}" value="${e.value||""}">`;if(e.type==="row")return`<div class="form-row">${Ge(e.fields)}</div>`;let s=e.required?"required":"",n=e.label?`<label class="form-label">${e.label}${e.required?' <span class="required">*</span>':""}</label>`:"",t="";switch(e.type){case"textarea":t=`<textarea name="${e.name}" class="form-control" placeholder="${e.placeholder||""}" ${s} rows="${e.rows||3}">${e.value||""}</textarea>`;break;case"select":let g=(e.options||[]).map(i=>{let y=typeof i=="object"?i.value:i,b=typeof i=="object"?i.label:i,p=e.value==y?"selected":"";return`<option value="${y}" ${p}>${b}</option>`}).join("");t=`<select name="${e.name}" class="form-control" ${s}><option value="">-- Pilih ${e.label||""} --</option>${g}</select>`;break;case"combobox":let u=`dl-${e.name}-${Math.random().toString(36).substring(7)}`,r=(e.options||[]).map(i=>{let y=typeof i=="object"?i.value:i;return`<option value="${typeof i=="object"?i.label:i}"></option>`}).join(""),l=e.value||"";if(e.value){let i=(e.options||[]).find(y=>(typeof y=="object"?y.value:y)==e.value);i&&(l=typeof i=="object"?i.label:i)}t=`
          <input type="text" name="${e.name}" list="${u}" class="form-control" value="${l}" placeholder="Pilih atau ketik baru..." ${s} autocomplete="off">
          <datalist id="${u}">${r}</datalist>
        `;break;case"checkbox":t=`<label class="checkbox-label"><input type="checkbox" name="${e.name}" value="1" ${e.value?"checked":""}> ${e.checkLabel||e.label}</label>`;break;case"date":let o=window.parseFlexibleDate&&e.value?window.parseFlexibleDate(e.value):e.value||"";t=`<input type="date" name="${e.name}" class="form-control" value="${o}" ${s}>`;break;case"number":t=`<input type="number" name="${e.name}" class="form-control" value="${e.value||""}" placeholder="${e.placeholder||""}" min="${e.min||""}" max="${e.max||""}" step="${e.step||"1"}" ${s}>`;break;case"email":t=`<input type="email" name="${e.name}" class="form-control" value="${e.value||""}" placeholder="${e.placeholder||""}" ${s}>`;break;case"url":t=`<input type="url" name="${e.name}" class="form-control" value="${e.value||""}" placeholder="${e.placeholder||"https://..."}" ${s}>`;break;default:t=`<input type="${e.type||"text"}" name="${e.name}" class="form-control" value="${e.value||""}" placeholder="${e.placeholder||""}" ${s} autocomplete="off">`}let c=e.hint?`<div class="form-hint">${e.hint}</div>`:"";return`<div class="form-group ${e.class||""}">${n}${t}${c}</div>`}).join("")}function ma(a){let e={},s=new FormData(a);for(let[n,t]of s.entries())e[n]=t===""?null:t;return a.querySelectorAll("input[type=checkbox]").forEach(n=>{n.checked||(e[n.name]=null)}),e}function ua(a,e){e&&Object.entries(e).forEach(([s,n])=>{let t=a.querySelector(`[name="${s}"]`);t&&(t.hasAttribute("list")||(t.type==="checkbox"?t.checked=!!n:t.type==="date"&&n&&window.parseFlexibleDate?t.value=window.parseFlexibleDate(n):t.value=n??""))})}B();function E({container:a,title:e,icon:s,apiPath:n,columns:t,formFields:c,filterFields:g,defaultFilters:u={},itemLabel:r="Data",canCreate:l=!0,canEdit:o=!0,canDelete:i=!0,onBeforeSubmit:y,onAfterLoad:b,extraActions:p=[],initialSearch:d="",exportOptions:m=null,bulkDelete:h=!1}){let v=1,k={...u};d&&(k.search=d);let S=new Set;a.innerHTML=`
    <div class="page-header">
      <h1 class="page-title">${s} ${e}</h1>
      <div class="page-actions">
        ${l?`<button class="btn btn-primary" id="btn-create">+ Tambah ${r}</button>`:""}
      </div>
    </div>

    ${h?`
    <div class="bulk-toolbar" id="bulk-toolbar" style="display:flex; align-items:center; gap:1rem; background:var(--bg-card); padding:0.75rem 1.25rem; border-radius:var(--radius-lg); border:1px solid var(--border-color); margin-bottom:1rem;">
      <span id="bulk-count" style="font-weight:600; font-size:0.9rem;">0 item dipilih</span>
      <button class="btn btn-danger btn-sm" id="btn-bulk-delete" disabled>\u{1F5D1}\uFE0F Hapus Terpilih</button>
      <button class="btn btn-ghost btn-sm" id="btn-bulk-cancel" disabled>Batalkan</button>
    </div>`:""}
    
    ${m?ze(m.moduleName):""}

    ${g&&g.length>0?`
    <div class="filter-bar card">
      <div class="filter-bar-inner">
        ${g.map(x=>x.type==="search"?`<div class="filter-search"><input type="search" class="form-control" placeholder="${x.placeholder||"Cari..."}" id="filter-search" value="${k.search||""}"></div>`:x.type==="select"?`<select class="form-control filter-select" name="${x.name}" id="filter-${x.name}"><option value="">-- ${x.label} --</option>${(x.options||[]).map(C=>`<option value="${typeof C=="object"?C.value:C}" ${k[x.name]===(typeof C=="object"?C.value:C)?"selected":""}>${typeof C=="object"?C.label:C}</option>`).join("")}</select>`:"").join("")}
        <button class="btn btn-ghost btn-sm" id="btn-reset-filter">Reset</button>
      </div>
    </div>`:""}

    <div class="card">
      <div class="card-body p-0" id="table-container">
        <div class="loading-spinner"><div class="spinner"></div></div>
      </div>
      <div class="card-footer" id="pagination-container"></div>
    </div>
  `;function w(){if(!document.getElementById("bulk-toolbar"))return;let C=document.getElementById("bulk-count"),P=document.getElementById("btn-bulk-delete"),F=document.getElementById("btn-bulk-cancel");C.textContent=`${S.size} item dipilih`,S.size>0?(P.disabled=!1,F.disabled=!1):(P.disabled=!0,F.disabled=!0)}document.getElementById("btn-bulk-cancel")?.addEventListener("click",()=>{S.clear(),document.querySelectorAll(".row-checkbox").forEach(C=>C.checked=!1);let x=document.getElementById("select-all-checkbox");x&&(x.checked=!1),w()}),document.getElementById("btn-bulk-delete")?.addEventListener("click",()=>{if(S.size===0)return;let x=[...S],C=document.createElement("div");C.style.cssText="position:fixed;inset:0;background:rgba(0,0,0,.5);z-index:9999;display:flex;align-items:center;justify-content:center",C.innerHTML=`
      <div style="background:var(--bg-card);border-radius:var(--radius-xl);padding:28px;max-width:420px;width:90%;box-shadow:var(--shadow-lg);animation:fadeInUp .2s ease">
        <h3 style="margin:0 0 8px;color:var(--text-1);font-size:1rem;font-weight:700">\u26A0\uFE0F Hapus ${x.length} ${r}?</h3>
        <p style="margin:0 0 24px;color:var(--text-2);font-size:.875rem">Data yang dihapus tidak dapat dikembalikan.</p>
        <div style="display:flex;gap:8px;justify-content:flex-end">
          <button id="bulk-cancel-btn" class="btn btn-ghost">Batal</button>
          <button id="bulk-confirm-btn" class="btn btn-danger">Hapus ${x.length} Data</button>
        </div>
      </div>
    `,document.body.appendChild(C),C.querySelector("#bulk-cancel-btn").addEventListener("click",()=>C.remove()),C.querySelector("#bulk-confirm-btn").addEventListener("click",async()=>{let P=C.querySelector("#bulk-confirm-btn");P.disabled=!0,P.textContent="Menghapus...";let F=await f(`${n}/bulk`,{method:"DELETE",body:JSON.stringify({ids:x})});C.remove(),F.ok?(q(`${x.length} ${r} berhasil dihapus.`),S.clear(),w(),_()):U(F.data?.error||"Gagal menghapus data.")})});let T=document.getElementById("filter-search"),L;if(T?.addEventListener("input",x=>{clearTimeout(L),L=setTimeout(()=>{k.search=x.target.value,v=1,_()},400)}),g?.forEach(x=>{x.type==="select"&&document.getElementById(`filter-${x.name}`)?.addEventListener("change",C=>{k[x.name]=C.target.value,v=1,_()})}),document.getElementById("btn-reset-filter")?.addEventListener("click",()=>{k={...u},T&&(T.value=""),g?.forEach(x=>{let C=document.getElementById(`filter-${x.name}`);C&&(C.value="")}),v=1,_()}),document.getElementById("btn-create")?.addEventListener("click",()=>J(null)),m){document.getElementById(`btn-export-${m.moduleName}`)?.addEventListener("click",async C=>{let P=C.target,F=P.innerHTML;P.innerHTML="\u23F3 Loading...",P.disabled=!0;try{await m.onExport()}catch{U("Gagal export data")}finally{P.innerHTML=F,P.disabled=!1}}),document.getElementById(`btn-template-${m.moduleName}`)?.addEventListener("click",()=>{m.onTemplate()});let x=document.getElementById(`input-import-${m.moduleName}`);x?.addEventListener("change",async C=>{let P=C.target.files[0];if(!P)return;let F=document.getElementById(`label-import-${m.moduleName}`),R=F?F.querySelector(".import-text"):null,ce=R?R.innerText:"";R&&(R.innerText="\u231B Memproses..."),F&&(F.style.pointerEvents="none"),x.disabled=!0;try{let H=await Qe(P);if(H.length===0)throw new Error("File kosong atau format salah");await m.onImport(H),q("Import berhasil!"),_()}catch(H){U(H.message||"Gagal import data")}finally{R&&(R.innerText=ce),F&&(F.style.pointerEvents="auto"),x.disabled=!1,x.value=""}})}async function _(){let x=document.getElementById("table-container");if(!x)return;x.innerHTML='<div class="loading-spinner"><div class="spinner"></div></div>';let C=new URLSearchParams({page:v,limit:1e4,...Object.fromEntries(Object.entries(k).filter(([,M])=>M))}),P=await f(`${n}?${C}`);if(!P.ok){x.innerHTML=`<div class="empty-state"><p class="text-danger">Gagal memuat data: ${P.data?.error||"Error"}</p></div>`;return}let F=P.data?.data||P.data||[],R=P.data?.pagination;if(!R&&Array.isArray(F)){let M=F.length,z=20,_e=Math.ceil(M/z);F=F.slice((v-1)*z,v*z),R={page:v,pages:_e,total:M,limit:z}}b&&b(F);let ce=da({columns:t,data:F,onEdit:o?M=>J(M):null,actions:p.map(M=>({...M,handler:z=>M.handler(z,_)})),emptyText:`Tidak ada ${String(r||"").toLowerCase()}`,bulkSelect:h?{selectedIds:S,onToggle:w}:null});x.innerHTML="",x.appendChild(ce);let H=document.getElementById("pagination-container");if(H&&(H.innerHTML="",R&&R.pages>1)){let M=pa({page:R.page,pages:R.pages,total:R.total,limit:R.limit,onPage:z=>{v=z,_()}});M&&H.appendChild(M)}}function D(x){let C=typeof c=="function"?c(x):c;return Ge(C)}function J(x){let C=!!x,P=document.createElement("form");if(P.noValidate=!0,P.innerHTML=D(x),C){let R=typeof c=="function"?c(x):c;ua(P,x)}let{close:F}=Y({title:C?`Edit ${r}`:`Tambah ${r}`,content:P,size:"lg",confirmText:C?"Simpan Perubahan":`Tambah ${r}`,onConfirm:async(R,ce)=>{if(!P.reportValidity())return;let H=R.querySelector(".modal-confirm");H.disabled=!0,H.textContent="Menyimpan...";let M=ma(P),z=typeof c=="function"?c(x):c,_e=async Ne=>{for(let j of Ne)if(j.type==="row")await _e(j.fields);else if(j.type==="combobox"&&M[j.name]){let xe=M[j.name],Ce=(j.options||[]).find(V=>{let Z=String(typeof V=="object"?V.value:V),Ha=String(typeof V=="object"?V.label:V);return Z===xe||Ha===xe});if(Ce)M[j.name]=typeof Ce=="object"?Ce.value:Ce;else if(j.createApi){let V={};V[j.createApi.field]=xe,j.createApi.extra&&Object.assign(V,j.createApi.extra);let Z=await f(j.createApi.path,{method:"POST",body:JSON.stringify(V)});if(Z.ok&&Z.data?.id)M[j.name]=Z.data.id;else if(Z.ok&&!Z.data?.id)M[j.name]=xe;else throw new Error(`Gagal membuat master data: ${Z.data?.error||"Unknown error"}`)}}};try{await _e(z)}catch(Ne){U(Ne.message),H.disabled=!1,H.textContent=C?"Simpan Perubahan":`Tambah ${r}`;return}y&&(M=await y(M,x));let Ka=C?"PUT":"POST",qa=C?`${n}/${x.id}`:n,ea=await f(qa,{method:Ka,body:JSON.stringify(M)});ea.ok?(q(C?`${r} berhasil diperbarui.`:`${r} berhasil ditambahkan.`),ce(),_()):(U(ea.data?.error||"Gagal menyimpan data."),H.disabled=!1,H.textContent=C?"Simpan Perubahan":`Tambah ${r}`)}})}function Ae(x){Ke(`Hapus ${r} ini? Tindakan tidak dapat dibatalkan.`,async()=>{let C=await f(`${n}/${x.id}`,{method:"DELETE"});C.ok?(q(`${r} berhasil dihapus.`),_()):U(C.data?.error||"Gagal menghapus.")},`Hapus ${r}`)}return _(),_}I();function O(a){let e={Done:"badge-success",Aktif:"badge-success",Open:"badge-warning","In Progress":"badge-info",Pending:"badge-warning",Diproses:"badge-info",Selesai:"badge-success","Tidak Aktif":"badge-neutral",Resign:"badge-neutral",Cut:"badge-danger","Tidak Datang":"badge-danger"};return!a||a==="-"||String(a).trim()===""?"":`<span class="badge ${e[a]||"badge-neutral"}">${a}</span>`}function ga(a){return a==null?'<span class="badge badge-neutral">-</span>':a<0?`<span class="badge badge-danger">Expired (${Math.abs(a)}h)</span>`:a<=14?`<span class="badge badge-danger">${a} hari</span>`:a<=30?`<span class="badge badge-warning">${a} hari</span>`:`<span class="badge badge-success">${a} hari</span>`}function Ee(a){return`<span class="badge ${{"FACILITY CARE":"badge-info",SECURITY:"badge-secondary"}[a]||"badge-neutral"}">${a||"-"}</span>`}function ba(a){return`<span class="badge ${{"Inspeksi Hygiene & Aset Bangunan":"badge-info","General Cleaning":"badge-success","Deep Cleaning":"badge-purple",Fogging:"badge-warning"}[a]||"badge-neutral"}">${a||"-"}</span>`}function W(a){return`<span class="badge ${{Q1:"badge-info",Q2:"badge-success",Q3:"badge-warning",Q4:"badge-danger"}[a]||"badge-neutral"}">${a||"-"}</span>`}B();var he=[],Ve=[];async function ot(){Ve=(await f("/api/branches?all=1")).data?.data||[],he=Ve.map(e=>({value:e.id,label:e.full_name}))}async function ha(a){await ot(),E({container:a,title:"Karyawan",icon:"\u{1F465}",apiPath:"/api/employees",itemLabel:"Karyawan",bulkDelete:!0,columns:[{key:"full_name",label:"Nama Lengkap"},{key:"branch_name",label:"Cabang"},{key:"division",label:"Divisi",render:e=>Ee(e)},{key:"phone",label:"No. HP",render:e=>e?`<a href="tel:${e}">${e}</a>`:"-"},{key:"join_date",label:"Tgl Masuk",render:e=>window.formatDate(e)},{key:"status",label:"Status",render:e=>O(e)}],filterFields:[{type:"search",placeholder:"Cari nama karyawan..."},{type:"select",name:"branch_id",label:"Cabang",options:he},{type:"select",name:"division",label:"Divisi",options:["FACILITY CARE","SECURITY"]},{type:"select",name:"status",label:"Status",options:["Aktif","Tidak Aktif","Resign","Cut"]}],formFields:e=>[{type:"row",fields:[{name:"full_name",label:"Nama Lengkap",required:!0,placeholder:"Nama lengkap karyawan",value:e?.full_name},{name:"phone",label:"No. HP",placeholder:"08xx-xxxx-xxxx",value:e?.phone}]},{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",options:e?.branch_id&&!he.find(s=>s.value==e.branch_id)?[...he,{value:e.branch_id,label:e.branch_name||e.branch_id}]:he,createApi:{path:"/api/branches",field:"full_name"},value:e?.branch_id},{name:"division",label:"Divisi",type:"select",required:!0,options:["FACILITY CARE","SECURITY"],value:e?.division||"FACILITY CARE"}]},{type:"row",fields:[{name:"join_date",label:"Tanggal Masuk",type:"date",value:e?.join_date},{name:"status",label:"Status",type:"select",required:!0,options:["Aktif","Tidak Aktif","Resign","Cut"],value:e?.status||""}]},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:e?.notes}],exportOptions:{moduleName:"employees",onExport:async()=>{let e=await f(`/api/employees${window.location.search?window.location.search+"&":"?"}limit=10000`);if(e.ok){let s=e.data.data.map(n=>({"Nama Lengkap":n.full_name,Cabang:n.branch_name||"",Divisi:n.division||"","No. HP":n.phone||"","Tgl Masuk":n.join_date||"",Status:n.status||""}));$(s,"Data_Karyawan")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{$([{"Nama Lengkap":"Budi Santoso",Cabang:"001. Pondok Bambu",Divisi:"FACILITY CARE","No. HP":"08123456789","Tgl Masuk":"2024-01-15",Status:"Aktif"},{"Nama Lengkap":"Andi Saputra",Cabang:"002. Bintaro",Divisi:"SECURITY","No. HP":"08987654321","Tgl Masuk":"2023-11-01",Status:"Aktif"}],"Template_Import_Karyawan")},onImport:async e=>{let s=c=>{if(!c)return null;let g=String(c||"").toLowerCase(),u=Ve.find(r=>String(r.full_name||"").toLowerCase()===g||String(r.code||"").toLowerCase()===g||String(r.name||"").toLowerCase()===g);return u?u.id:null},n=e.map(c=>({full_name:String(c["Nama Lengkap"]||"").trim(),branch_id:s(String(c.Cabang||"").trim()),division:String(c.Divisi||"").trim()||"FACILITY CARE",phone:String(c["No. HP"]||"").trim(),join_date:String(c["Tgl Masuk"]||"").trim(),status:String(c.Status||"").trim(),notes:String(c.Catatan||"").trim()})).filter(c=>c.full_name),t=await f("/api/employees/import",{method:"POST",body:JSON.stringify(n)});if(!t.ok)throw new Error(t.data?.error||"Import gagal")}}})}I();B();var ye=[],Pe=[];async function ct(){let[a,e]=await Promise.all([f("/api/branches?all=1"),f("/api/employees?limit=10000&status=Aktif")]);ye=(a.data?.data||[]).map(s=>({value:s.id,label:s.full_name})),Pe=(e.data?.data||[]).map(s=>({value:s.id,label:s.full_name}))}var We=async a=>{let e=[],s=1;for(;;){let t=await(await Promise.resolve().then(()=>(I(),ne))).apiFetch(`${a}${a.includes("?")?"&":"?"}limit=100&page=${s}`);if(!t.ok)break;let c=t.data?.data||t.data||[],g=Array.isArray(c)?c:[];if(e=e.concat(g),g.length<100||t.data?.pagination&&s>=t.data.pagination.pages)break;s++}return e};async function De(a){await ct(),E({container:a,title:"Data Kontrak",icon:"\u{1F4CB}",apiPath:"/api/contracts",bulkDelete:!0,itemLabel:"Kontrak",columns:[{key:"employee_name",label:"Nama Lengkap"},{key:"branch_name",label:"Cabang"},{key:"division",label:"Div / Bagian",render:e=>Ee(e)},{key:"start_date",label:"Tanggal Mulai",nowrap:!0,render:e=>window.formatDate(e)},{key:"end_date",label:"Tanggal Selesai",nowrap:!0,render:e=>!e||String(e).startsWith("2099")?"Tetap / PKWTT":window.formatDate(e)},{key:"days_remaining",label:"Sisa Kontrak",render:(e,s)=>s.end_date&&String(s.end_date).startsWith("2099")?'<span class="badge badge-success" style="background:#10B981;color:white;padding:4px 8px;border-radius:6px;font-size:0.75rem;font-weight:600">Tetap</span>':ga(e)},{key:"status",label:"Status",render:e=>O(e)}],filterFields:[{type:"search",placeholder:"Cari nama karyawan..."},{type:"select",name:"branch_id",label:"Cabang",options:ye},{type:"select",name:"status",label:"Status",options:["Aktif","Tidak Aktif","Resign","Cut"]},{type:"select",name:"expiring_days",label:"Akan Habis",options:[{value:"7",label:"7 Hari"},{value:"14",label:"14 Hari"},{value:"30",label:"30 Hari"},{value:"60",label:"60 Hari"}]}],onBeforeSubmit:e=>(e.end_date||(e.end_date="2099-12-31"),e),onAfterLoad:()=>{if(!document.getElementById("btn-find-missing")){let e=document.createElement("button");e.id="btn-find-missing",e.className="btn btn-ghost",e.innerHTML="\u{1F50D} Cek Selisih Karyawan",e.style.marginLeft="8px",e.style.color="#EF4444",e.style.border="1px solid currentColor",e.onclick=async()=>{e.innerHTML="\u231B Mencari...",e.disabled=!0;try{let[n,t]=await Promise.all([We("/api/employees?status=Aktif"),We("/api/contracts")]);if(n.length>0){let c=t.filter(l=>l.status==="Aktif"),g=new Set(c.map(l=>l.employee_id)),u=n.filter(l=>!g.has(l.id)),r=`<p style="margin-bottom:12px">Data yang terbaca: <b>${n.length}</b> Karyawan Aktif, dan <b>${c.length}</b> Kontrak Aktif.</p>
              <p style="margin-bottom:12px">Terdapat <b>${u.length}</b> karyawan aktif yang tidak memiliki "Kontrak Aktif". Berikut daftarnya:</p><ul style="padding-left:20px; max-height:400px; overflow-y:auto">`;u.forEach(l=>{let o=t.filter(y=>y.employee_id===l.id),i='<span style="color:#F59E0B">Belum pernah di-input kontrak</span>';if(o.length>0){let y=o[0];i=`Pernah ada kontrak (Status: <b style="color:#EF4444">${y.status}</b>, Selesai: ${window.formatDate(y.end_date)})`}r+=`<li style="margin-bottom:8px"><b>${l.full_name}</b> <br><span style="font-size:0.85em;color:var(--text-2)">Cabang: ${l.branch_name||"-"} | ${i}</span></li>`}),r+="</ul>",Promise.resolve().then(()=>(ge(),ia)).then(l=>l.createModal({title:"Karyawan Tanpa Kontrak Aktif",content:r,cancelText:"Tutup"}))}}catch(n){console.error(n)}e.innerHTML="\u{1F50D} Cek Selisih Karyawan",e.disabled=!1};let s=document.querySelector(".page-actions");s&&s.appendChild(e)}},formFields:e=>[{type:"row",fields:[{name:"employee_id",label:"Nama Lengkap",type:"combobox",required:!0,options:e?.employee_id&&!Pe.find(s=>s.value==e.employee_id)?[...Pe,{value:e.employee_id,label:e.employee_name||e.employee_id}]:Pe,createApi:{path:"/api/employees",field:"full_name"},value:e?.employee_id},{name:"branch_id",label:"Cabang",type:"combobox",options:e?.branch_id&&!ye.find(s=>s.value==e.branch_id)?[...ye,{value:e.branch_id,label:e.branch_name||e.branch_id}]:ye,createApi:{path:"/api/branches",field:"full_name"},value:e?.branch_id}]},{type:"row",fields:[{name:"division",label:"Div / Bagian",type:"select",required:!0,options:["FACILITY CARE","SECURITY"],value:e?.division||"FACILITY CARE"},{name:"status",label:"Status",type:"select",required:!0,options:["Aktif","Tidak Aktif","Resign","Cut"],value:e?.status||""}]},{type:"row",fields:[{name:"start_date",label:"Tanggal Mulai",type:"date",value:e?.start_date},{name:"end_date",label:"Tanggal Selesai",type:"date",value:e?.end_date&&!String(e.end_date).startsWith("2099")?e.end_date:""}]},{type:"row",fields:[{name:"contract_type",label:"Tipe Kontrak",type:"select",options:["KONTRAK 6 BULAN","KONTRAK 1 TAHUN","KONTRAK 2 TAHUN"],value:e?.contract_type},{name:"pkwt_number",label:"No. PKWT",type:"select",options:["PKWT 1","PKWT 2","PKWT 3","PKWT 4","PKWT 5","PKWT 6"],value:e?.pkwt_number}]},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:e?.notes}],exportOptions:{moduleName:"contracts",onExport:async()=>{let e=await f(`/api/contracts${window.location.search?window.location.search+"&":"?"}limit=10000`);if(e.ok){let s=e.data.data.map(n=>({"Nama Lengkap":n.employee_name,Cabang:n.branch_name||"","Div / Bagian":n.division||"","Tanggal Mulai":n.start_date||"","Tanggal Selesai":n.end_date&&String(n.end_date).startsWith("2099")?"":n.end_date||"","Sisa Kontrak":n.end_date&&String(n.end_date).startsWith("2099")?"Tetap":n.days_remaining!==null&&n.days_remaining!==void 0?`${n.days_remaining} Hari`:"",Status:n.status||""}));$(s,"Data_Kontrak")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{$([{"Nama Lengkap":"Budi Santoso",Cabang:"001. Pondok Bambu","Div / Bagian":"FACILITY CARE","Tanggal Mulai":"2024-01-01","Tanggal Selesai":"2024-12-31","Sisa Kontrak":"365 Hari",Status:"Aktif"}],"Template_Import_Kontrak")},onImport:async e=>{let[s,n]=await Promise.all([f("/api/branches?limit=10000"),We("/api/employees")]),t=s.data?.data||[],c=n||[];console.log(`Total employee yang berhasil dimuat dari database : ${c.length}`),c.length>0&&(console.log("Contoh 5 employee pertama:"),c.slice(0,5).forEach((p,d)=>{console.log(`${d+1}. ID: ${p.id}, Name: ${p.full_name}, Status: ${p.status}`)}));let g=p=>{if(!p)return null;let d=String(p||"").replace(/\s+/g," ").toLowerCase().trim(),m=t.find(h=>String(h.full_name||"").replace(/\s+/g," ").toLowerCase().trim()===d||String(h.code||"").replace(/\s+/g," ").toLowerCase().trim()===d||String(h.name||"").replace(/\s+/g," ").toLowerCase().trim()===d);return m?m.id:null},u=(p,d)=>{if(console.log("------------------------------------------------"),console.log(`Row Excel : ${d}`),console.log(`Nama dari Excel : "${p}"`),!p)return console.log("Alasan gagal mapping : Nama kosong"),null;let m=String(p||"").replace(/\s+/g," ").toLowerCase().trim();console.log(`Nama setelah normalisasi : "${m}"`),console.log(`Jumlah employee di database : ${c.length}`);let h=c.find(v=>String(v.full_name||"").replace(/\s+/g," ").toLowerCase().trim()===m);return h?(console.log("Employee ditemukan atau tidak : Ditemukan"),console.log(`Employee ID jika ditemukan : ${h.id}`),h.id):(console.log("Employee ditemukan atau tidak : TIDAK Ditemukan"),console.log("Alasan gagal mapping : Tidak ada kecocokan full_name setelah normalisasi"),null)},r=p=>{if(!p)return"";if(p instanceof Date&&!isNaN(p.getTime()))return p.toISOString().slice(0,10);let d=String(p).trim();if(/^\d{4,5}(\.\d+)?$/.test(d)){let h=Math.floor(Number(d));if(h>2e4&&h<99999){let v=new Date(Date.UTC(1899,11,30)+h*864e5);return isNaN(v.getTime())?"":v.toISOString().slice(0,10)}}if(/^\d{4}-\d{2}-\d{2}/.test(d))return d.slice(0,10);let m=d.split(/[\/\-\.]/);if(m.length===3){let[h,v,k]=m.map(S=>S.trim());if(h.length===4&&v.length<=2&&k.length<=2)return`${h}-${v.padStart(2,"0")}-${k.padStart(2,"0")}`;if(k.length===4&&v.length<=2&&h.length<=2)return`${k}-${v.padStart(2,"0")}-${h.padStart(2,"0")}`}return d},l=e.map((p,d)=>{let m=d+2,h=String(p["Nama Lengkap"]||"").trim(),v=p["Tanggal Mulai"],k=r(v);if(!k){let T=e.__worksheet,L=e.__headers||[],_=L.indexOf("Tanggal Mulai"),D="N/A",J="N/A",Ae="N/A";if(_!==-1&&T&&window.XLSX){let C=window.XLSX.utils.encode_cell({c:_,r:m-1});Ae=C;let P=T[C];P?(D=P.t||"undefined",J=P.w||"undefined"):D="CELL KOSONG/TIDAK ADA DI WORKSHEET"}let x="Unknown";v==null||v===""?x="Kondisi IF: Nilai murni undefined, null, atau string kosong dari parsed JSON.":v instanceof Date&&isNaN(v.getTime())?x="Kondisi IF: Nilai adalah object Date namun invalid (isNaN).":x="Kondisi IF: Tidak lolos Regex YYYY-MM-DD maupun konversi serial number Excel.",console.log("=========================="),console.log("[DEBUG] DATE PARSING FAILED"),console.log("=========================="),console.log(`Excel Row Number : ${m}`),console.log(`Employee Name : ${h}`),console.log(`Column Header Used : "Tanggal Mulai" (Index: ${_})`),console.log(`Raw Cell Value : "${v}"`),console.log(`JavaScript Type : ${typeof v}`),console.log(`SheetJS Cell Type : ${D}`),console.log(`SheetJS Formatted Value : "${J}"`),console.log(`Value After Trim : "${String(v||"").trim()}"`),console.log(`Value After Date Parser : "${k}"`),console.log(`Is Empty : ${!v}`),console.log(`Is Invalid Date : ${v instanceof Date?isNaN(v.getTime()):"Bukan JS Date Object"}`),console.log(`Reason : ${x}`),console.log(`Workbook Sheet : ${T?"Ada":"Tidak Ditemukan"}`),console.log(`Excel Cell Address : ${Ae}`),console.log(`
--- Seluruh Kolom Pada Baris Ini (Mencegah Column Shift) ---`),console.log(JSON.stringify(p,null,2)),console.log(`
--- Daftar Seluruh Header Yang Terbaca ---`),console.log(JSON.stringify(L)),console.log(`==========================
`)}let S=u(h,m),w=null;return S?k||(w="Tanggal Mulai kosong atau tidak berformat tanggal"):w="Karyawan tidak ditemukan di Database",{isValid:!!(S&&k),invalidReason:w,rowNum:m,data:{employee_id:S,branch_id:g(String(p.Cabang||"").trim()),division:String(p["Div / Bagian"]||"").trim()||"FACILITY CARE",start_date:k,end_date:r(p["Tanggal Selesai"])||"2099-12-31",status:String(p.Status||"").trim(),_rawName:h}}}),o=[],i=[];if(l.forEach(p=>{p.isValid?o.push(p.data):i.push({rowNum:p.rowNum,name:p.data._rawName,reason:p.invalidReason})}),console.log(`Split Validation - Valid: ${o.length}, Invalid: ${i.length}`),o.length===0){let p=`SEMUA BARIS GAGAL IMPORT!

Total Excel: ${e.length}
Valid: 0
Invalid: ${i.length}

Daftar Kegagalan (Contoh):
`;i.slice(0,10).forEach(d=>{p+=`- Row ${d.rowNum} | Nama: ${d.name} | Alasan: ${d.reason}
`}),i.length>10&&(p+=`- ... dan ${i.length-10} lainnya.
`),alert(p);return}let y=await f("/api/contracts/import",{method:"POST",body:JSON.stringify(o)}),b=`IMPORT SUMMARY
======================
`;b+=`Total Baris Excel : ${e.length}
`,b+=`Baris Valid       : ${o.length}
`,b+=`Baris Invalid     : ${i.length}

`,y&&y.data&&y.data.metrics?(b+=`Berhasil INSERT   : ${y.data.metrics.inserted}
`,b+=`Berhasil UPDATE   : ${y.data.metrics.updated}
`):b+=`Berhasil diproses : ${o.length}
`,i.length>0&&(b+=`
DAFTAR DATA DILEWATI:
`,i.forEach(p=>{b+=`- Row ${p.rowNum} | ${p.name} | ${p.reason}
`})),alert(b),typeof De=="function"&&De()}}})}I();B();var fe=[],Ye=[];async function ya(a){let[e,s,n]=await Promise.all([f("/api/branches?all=1"),f(`/api/schedule${window.location.search?window.location.search+"&":"?"}limit=10000`),f(`/api/schedule${window.location.search?window.location.search+"&":"?"}limit=10000`)]);fe=(e.data?.data||[]).map(r=>({value:r.id,label:r.full_name}));let t=(s.data?.data||[]).map(r=>({value:r.full_name,label:r.full_name}));Ye=[...(n.data?.data||[]).filter(r=>r.role==="FC Spesialis").map(r=>({value:r.name,label:r.name}))];let g=r=>r&&!t.find(l=>l.value===r)?[...t,{value:r,label:r}]:t,u=r=>{if(!r||r==="-"||String(r).trim()==="")return"";let l=String(r).split("-");return l.length===3&&l[0].length===4?`${l[2]}-${l[1]}-${l[0]}`:r};E({container:a,title:"Jadwal Kegiatan",icon:"\u{1F5D3}\uFE0F",apiPath:"/api/schedule",bulkDelete:!0,itemLabel:"Jadwal",columns:[{key:"branch_name",label:"Cabang"},{key:"activity_type",label:"Kegiatan",render:r=>ba(r)},{key:"period",label:"Periode",render:r=>W(r)},{key:"pic",label:"PIC"},{key:"opening_date",label:"Tgl Opening",nowrap:!0,render:r=>u(r)},{key:"target_date",label:"Tgl Target",nowrap:!0,render:r=>u(r)},{key:"completion_date",label:"Tgl Selesai",nowrap:!0,render:r=>u(r)},{key:"status",label:"Status",render:r=>O(r)}],filterFields:[{type:"select",name:"branch_id",label:"Cabang",options:fe},{type:"select",name:"activity_type",label:"Kegiatan",options:["Inspeksi Hygiene & Aset Bangunan","General Cleaning","Deep Cleaning","Fogging"]},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"status",label:"Status",options:["Pending","In Progress","Done"]},{type:"select",name:"pic",label:"PIC",options:Ye}],formFields:r=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:r?.branch_id&&!fe.find(l=>l.value==r.branch_id)?[...fe,{value:r.branch_id,label:r.branch_name||r.branch_id}]:fe,createApi:{path:"/api/branches",field:"full_name"},value:r?.branch_id},{name:"activity_type",label:"Jenis Kegiatan",type:"select",required:!0,options:["Inspeksi Hygiene & Aset Bangunan","General Cleaning","Deep Cleaning","Fogging"],value:r?.activity_type}]},{type:"row",fields:[{name:"period",label:"Periode",type:"select",required:!0,options:["Q1","Q2","Q3","Q4"],value:r?.period},{name:"pic",label:"PIC",type:"combobox",options:Ye,createApi:{path:"/api/pic",field:"name"},value:r?.pic}]},{type:"row",fields:[{name:"opening_date",label:"Tanggal Opening",type:"date",value:r?.opening_date},{name:"target_date",label:"Tanggal Target",type:"date",value:r?.target_date}]},{type:"row",fields:[{name:"completion_date",label:"Tanggal Selesai",type:"date",value:r?.completion_date},{name:"status",label:"Status",type:"select",required:!0,options:["Pending","In Progress","Done"],value:r?.status||""}]},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:r?.notes}],exportOptions:{moduleName:"schedule",onExport:async()=>{let r=await f(`/api/schedule${window.location.search?window.location.search+"&":"?"}limit=10000`);if(r.ok){let l=r.data.data.map(o=>({Cabang:o.branch_name||"",Kegiatan:o.activity_type||"",Periode:o.period||"",PIC:o.pic||"","Tgl Opening":o.opening_date||"","Tgl Target":o.target_date||"","Tgl Selesai":o.completion_date||"",Status:o.status||""}));$(l,"Data_Jadwal_Kegiatan")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{$([{Cabang:"001. Pondok Bambu",Kegiatan:"General Cleaning",Periode:"Q1",PIC:"Fajar","Tgl Opening":"2024-02-01","Tgl Target":"2024-02-15","Tgl Selesai":"2024-02-14",Status:"Done"}],"Template_Import_Jadwal")},onImport:async r=>{let o=(await f("/api/branches?all=1")).data?.data||[],i=d=>{if(!d)return null;let m=String(d||"").toLowerCase(),h=o.find(v=>String(v.full_name||"").toLowerCase()===m||String(v.code||"").toLowerCase()===m||String(v.name||"").toLowerCase()===m);return h?h.id:null},y=d=>{if(d==null||d==="")return"";if(d instanceof Date&&!isNaN(d.getTime()))return d.toISOString().slice(0,10);let m=String(d).trim();if(m===""||m==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(m))return m.slice(0,10);if(/^\d{4,5}$/.test(m)){let v=Number(m);if(v>2e4&&v<99999){let k=new Date(Date.UTC(1899,11,30)+v*864e5);return isNaN(k.getTime())?"":k.toISOString().slice(0,10)}}let h=m.split(/[\/\-\.]/);if(h.length===3){let[v,k,S]=h.map(w=>w.trim());if(v.length===4&&k.length<=2&&S.length<=2)return`${v}-${k.padStart(2,"0")}-${S.padStart(2,"0")}`;if(S.length===4&&k.length<=2&&v.length<=2)return`${S}-${k.padStart(2,"0")}-${v.padStart(2,"0")}`}return m},b=r.map(d=>({branch_id:i(String(d.Cabang||"").trim()),activity_type:String(d.Kegiatan||"").trim(),period:String(d.Periode||"").trim(),pic:String(d.PIC||d.Pic||"").trim(),opening_date:y(d["Tgl Opening"]||d["Tanggal Opening"]||d["Tgl Openir"]),target_date:y(d["Tgl Target"]||d["Tanggal Target"]),completion_date:y(d["Tgl Selesai"]||d["Tanggal Selesai"]),status:String(d.Status||"").trim(),notes:String(d.Catatan||d.Keterangan||"").trim()})).filter(d=>d.activity_type&&d.period),p=await f("/api/schedule/import",{method:"POST",body:JSON.stringify(b)});if(!p.ok)throw new Error(p.data?.error||"Import gagal")}}})}I();B();var ve=[],Ie=[];async function fa(a){let[e,s,n]=await Promise.all([f("/api/branches?all=1"),f(`/api/issues${window.location.search?window.location.search+"&":"?"}limit=10000`),f(`/api/issues${window.location.search?window.location.search+"&":"?"}limit=10000`)]);ve=(e.data?.data||[]).map(i=>({value:i.id,label:i.full_name})),Ie=(s.data?.data||[]).map(i=>({value:i.full_name,label:i.full_name}));let t=(n.data?.data||[]).filter(i=>i.role==="FC Spesialis").map(i=>({value:i.name,label:i.name})),c=(n.data?.data||[]).filter(i=>i.role==="Pelapor").map(i=>({value:i.name,label:i.name})),g=i=>i&&!Ie.find(y=>y.value===i)?[...Ie,{value:i,label:i}]:Ie,u=i=>i&&!t.find(y=>y.value===i)?[...t,{value:i,label:i}]:t,r=i=>i&&!c.find(y=>y.value===i)?[...c,{value:i,label:i}]:c,l=new Date().getFullYear(),o=Array.from({length:5},(i,y)=>String(l-y));E({container:a,title:"Permasalahan",icon:"\u26A0\uFE0F",apiPath:"/api/issues",bulkDelete:!0,itemLabel:"Permasalahan",columns:[{key:"report_date",label:"Tanggal",nowrap:!0,render:i=>window.formatDate(i)},{key:"branch_name",label:"Cabang"},{key:"category",label:"Kategori",render:i=>`<span class="badge badge-secondary">${i}</span>`},{key:"source",label:"Sumber"},{key:"complaint",label:"Keluhan",render:i=>`<span title="${i}">${i?.length>50?i.slice(0,50)+"\u2026":i}</span>`},{key:"employee_name",label:"Nama FC"},{key:"fc_specialist",label:"FC Spesialis"},{key:"solution",label:"Solusi",render:i=>`<span title="${i||""}">${i?.length>40?i.slice(0,40)+"\u2026":i||"-"}</span>`},{key:"status",label:"Status",render:i=>O(i)},{key:"completion_date",label:"Tgl Selesai",nowrap:!0,render:i=>window.formatDate(i)},{key:"day_count",label:"Hari",render:i=>i??"-"}],filterFields:[{type:"search",placeholder:"Cari keluhan / nama FC..."},{type:"select",name:"branch_id",label:"Cabang",options:ve},{type:"select",name:"category",label:"Kategori",options:["SDM","Cleaning","Aset","K3","Lainnya"]},{type:"select",name:"status",label:"Status",options:["Open","In Progress","Done"]},{type:"select",name:"year",label:"Tahun",options:o}],formFields:i=>[{type:"row",fields:[{name:"report_date",label:"Tanggal Info",type:"date",required:!0,value:i?.report_date},{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:i?.branch_id&&!ve.find(y=>y.value==i.branch_id)?[...ve,{value:i.branch_id,label:i.branch_name||i.branch_id}]:ve,createApi:{path:"/api/branches",field:"full_name"},value:i?.branch_id}]},{type:"row",fields:[{name:"category",label:"Kategori",type:"select",required:!0,options:["SDM","Cleaning","Aset","K3","Lainnya"],value:i?.category},{name:"source",label:"Sumber Laporan",type:"select",options:[...r(i?.source),{value:"Lainnya",label:"Lainnya"}],value:i?.source}]},{name:"complaint",label:"Keluhan",type:"textarea",required:!0,rows:3,value:i?.complaint},{type:"row",fields:[{name:"employee_name",label:"Nama FC / Security",type:"select",options:g(i?.employee_name),value:i?.employee_name},{name:"fc_specialist",label:"FC Spesialis",type:"select",options:u(i?.fc_specialist),value:i?.fc_specialist}]},{name:"solution",label:"Solusi / Tindakan",type:"textarea",rows:3,value:i?.solution},{type:"row",fields:[{name:"status",label:"Status",type:"select",required:!0,options:["Open","In Progress","Done"],value:i?.status||""},{name:"completion_date",label:"Tanggal Selesai",type:"date",value:i?.completion_date}]}],exportOptions:{moduleName:"issues",onExport:async()=>{let i=await f(`/api/issues${window.location.search?window.location.search+"&":"?"}limit=10000`);if(i.ok){let y=i.data.data.map(b=>({Tanggal:b.report_date||"",Cabang:b.branch_name||"",Kategori:b.category||"",Sumber:b.source||"",Keluhan:b.complaint||"","Nama FC":b.employee_name||"","FC Spesialis":b.fc_specialist||"",Solusi:b.solution||"","Tgl Selesai":b.completion_date||"",Status:b.status||""}));$(y,"Data_Permasalahan")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{$([{Tanggal:"2024-03-01",Cabang:"001. Pondok Bambu",Kategori:"Cleaning",Sumber:"SPV",Keluhan:"Lantai kotor","Nama FC":"Budi Santoso","FC Spesialis":"Fajar",Solusi:"Teguran lisan","Tgl Selesai":"2024-03-02",Status:"Done"}],"Template_Import_Permasalahan")},onImport:async i=>{let b=(await f("/api/branches?all=1")).data?.data||[],p=h=>{if(!h)return null;let v=String(h||"").toLowerCase(),k=b.find(S=>String(S.full_name||"").toLowerCase()===v||String(S.code||"").toLowerCase()===v||String(S.name||"").toLowerCase()===v);return k?k.id:null},d=i.map(h=>({branch_id:p(String(h.Cabang||"").trim()),report_date:String(h.Tanggal||"").trim(),category:String(h.Kategori||"").trim(),source:String(h.Sumber||"").trim(),complaint:String(h.Keluhan||"").trim(),employee_name:String(h["Nama FC"]||"").trim(),fc_specialist:String(h["FC Spesialis"]||"").trim(),solution:String(h.Solusi||"").trim(),completion_date:String(h["Tgl Selesai"]||"").trim(),status:String(h.Status||"").trim()})).filter(h=>h.report_date&&h.complaint&&h.category),m=await f("/api/issues/import",{method:"POST",body:JSON.stringify(d)});if(!m.ok)throw new Error(m.data?.error||"Import gagal")}}})}I();async function va(a){let[e,s,n]=await Promise.all([f("/api/branches?all=1"),f(`/api/one_on_one${window.location.search?window.location.search+"&":"?"}limit=10000`),f(`/api/one_on_one${window.location.search?window.location.search+"&":"?"}limit=10000`)]),t=(e.data?.data||[]).map(l=>({value:l.id,label:l.full_name})),c=(s.data?.data||[]).map(l=>({value:l.full_name,label:l.full_name})),g=(n.data?.data||[]).filter(l=>l.role==="FC Spesialis").map(l=>({value:l.name,label:l.name})),u=l=>l&&!c.find(o=>o.value===l)?[...c,{value:l,label:l}]:c,r=l=>l&&!g.find(o=>o.value===l)?[...g,{value:l,label:l}]:g;E({container:a,title:"One on One",icon:"\u{1F91D}",apiPath:"/api/one-on-one",bulkDelete:!0,itemLabel:"One on One",columns:[{key:"meeting_date",label:"Tanggal",nowrap:!0,render:l=>window.formatDate(l)},{key:"branch_name",label:"Cabang"},{key:"employee_name",label:"Nama Karyawan"},{key:"pic",label:"PIC"},{key:"problem",label:"Masalah",render:l=>`<span title="${l||""}">${l?.length>50?l.slice(0,50)+"\u2026":l||"-"}</span>`},{key:"solution",label:"Solusi",render:l=>`<span title="${l||""}">${l?.length>40?l.slice(0,40)+"\u2026":l||"-"}</span>`},{key:"status",label:"Status",render:l=>O(l)},{key:"completion_date",label:"Tgl Selesai",nowrap:!0,render:l=>window.formatDate(l)},{key:"day_count",label:"Hari"},{key:"document_link",label:"Dokumen",render:l=>l?`<a href="${l}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari nama / masalah..."},{type:"select",name:"branch_id",label:"Cabang",options:t},{type:"select",name:"status",label:"Status",options:["Open","Done"]}],exportOptions:{moduleName:"one_on_one",onExport:async l=>{let o=new URLSearchParams(l||{}).toString(),i=await f(`/api/one-on-one?limit=10000&${o}`);if(i.ok){let y=i.data.data.map(p=>({Tanggal:p.meeting_date||"",Cabang:p.branch_name||"","Nama Karyawan":p.employee_name||"",PIC:p.pic||"",Masalah:p.problem||"",Solusi:p.solution||"",Status:p.status||"","Tgl Selesai":p.completion_date||"",Dokumen:p.document_link||""})),{downloadExcel:b}=await Promise.resolve().then(()=>(B(),G));b(y,`Data_One_on_One_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:async()=>{let l=[{Tanggal:"2026-01-08",Cabang:"001. Pondok Bambu","Nama Karyawan":"Widya Astuti",PIC:"Rina",Masalah:"Terlambat terus",Solusi:"Teguran",Status:"Open","Tgl Selesai":"",Dokumen:"https://link.doc"}],{downloadExcel:o}=await Promise.resolve().then(()=>(B(),G));o(l,"Template_Import_OneOnOne")},onImport:async l=>{let o=p=>{if(!p)return null;let d=String(p||"").toLowerCase(),m=e.data?.data.find(h=>String(h.full_name||"").toLowerCase()===d||String(h.code||"").toLowerCase()===d||String(h.name||"").toLowerCase()===d);return m?m.id:null},i=p=>{if(!p)return"";if(p instanceof Date&&!isNaN(p.getTime()))return p.toISOString().slice(0,10);let d=String(p).trim();if(/^\d{4,5}$/.test(d)){let h=Number(d);if(h>2e4&&h<99999){let v=new Date(Date.UTC(1899,11,30)+h*864e5);return isNaN(v.getTime())?"":v.toISOString().slice(0,10)}}if(/^\d{4}-\d{2}-\d{2}/.test(d))return d.slice(0,10);let m=d.split(/[\/\-\.]/);if(m.length===3){let[h,v,k]=m.map(S=>S.trim());if(h.length===4&&v.length<=2&&k.length<=2)return`${h}-${v.padStart(2,"0")}-${k.padStart(2,"0")}`;if(k.length===4&&v.length<=2&&h.length<=2)return`${k}-${v.padStart(2,"0")}-${h.padStart(2,"0")}`}return d},y=l.map(p=>({meeting_date:i(p.Tanggal),employee_name:String(p["Nama Karyawan"]||"").trim(),branch_id:o(String(p.Cabang||"").trim()),pic:String(p.PIC||"").trim(),problem:String(p.Masalah||"").trim(),solution:String(p.Solusi||"").trim(),status:String(p.Status||"").trim(),completion_date:i(p["Tgl Selesai"]),document_link:String(p.Dokumen||"").trim()})).filter(p=>p.meeting_date&&p.employee_name&&p.branch_id),b=await f("/api/one-on-one/import",{method:"POST",body:JSON.stringify(y)});if(!b.ok)throw new Error(b.data?.error||"Import gagal")}},formFields:l=>[{type:"row",fields:[{name:"meeting_date",label:"Tanggal",type:"date",required:!0,value:l?.meeting_date},{name:"branch_id",label:"Cabang",type:"combobox",options:l?.branch_id&&!t.find(o=>o.value==l.branch_id)?[...t,{value:l.branch_id,label:l.branch_name||l.branch_id}]:t,createApi:{path:"/api/branches",field:"full_name"},value:l?.branch_id}]},{type:"row",fields:[{name:"employee_name",label:"Nama Karyawan",type:"select",required:!0,options:u(l?.employee_name),value:l?.employee_name},{name:"pic",label:"PIC",type:"combobox",options:r(l?.pic),createApi:{path:"/api/pic",field:"name"},value:l?.pic}]},{name:"problem",label:"Masalah",type:"textarea",required:!0,rows:3,value:l?.problem},{name:"solution",label:"Solusi",type:"textarea",rows:3,value:l?.solution},{type:"row",fields:[{name:"status",label:"Status",type:"select",required:!0,options:["Open","Done"],value:l?.status||""},{name:"completion_date",label:"Tanggal Selesai",type:"date",value:l?.completion_date}]},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://drive.google.com/...",value:l?.document_link}]})}I();async function ka(a){let[e,s,n]=await Promise.all([f("/api/branches?all=1"),f(`/api/training${window.location.search?window.location.search+"&":"?"}limit=10000`),f(`/api/training${window.location.search?window.location.search+"&":"?"}limit=10000`)]),t=(e.data?.data||[]).map(o=>({value:o.id,label:o.full_name})),c=(s.data?.data||[]).map(o=>({value:o.full_name,label:o.full_name})),g=(n.data?.data||[]).filter(o=>o.role==="FC Spesialis").map(o=>({value:o.name,label:o.name})),u=o=>o&&!c.find(i=>i.value===o)?[...c,{value:o,label:o}]:c,r=o=>o&&!g.find(i=>i.value===o)?[...g,{value:o,label:o}]:g,l=Array.from({length:5},(o,i)=>String(new Date().getFullYear()-i));E({container:a,title:"Training",icon:"\u{1F393}",apiPath:"/api/training",bulkDelete:!0,itemLabel:"Training",columns:[{key:"training_date",label:"Tanggal",nowrap:!0,render:o=>window.formatDate(o)},{key:"batch",label:"Batch"},{key:"subject",label:"Materi"},{key:"branch_name",label:"Cabang"},{key:"trainer",label:"Trainer"},{key:"participants",label:"Peserta",render:o=>{try{let i=JSON.parse(o);return Array.isArray(i)?i.join(", "):o||"-"}catch{return o||"-"}}},{key:"score",label:"Nilai",render:o=>o!=null?`<strong>${o}</strong>`:"-"},{key:"document_link",label:"Dokumen",render:o=>o?`<a href="${o}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari materi / trainer..."},{type:"select",name:"year",label:"Tahun",options:l}],exportOptions:{moduleName:"training",onExport:async o=>{let i=new URLSearchParams(o||{}).toString(),y=await f(`/api/training?limit=10000&${i}`);if(y.ok){let b=y.data.data.map(d=>{let m=d.participants||"";try{let h=JSON.parse(m);m=Array.isArray(h)?h.join(", "):m}catch{}return{Tanggal:d.training_date||"",Batch:d.batch||"",Materi:d.subject||"",Cabang:d.branch_name||"",Trainer:d.trainer||"",Peserta:m,Nilai:d.score!==null&&d.score!==void 0?d.score:"",Dokumen:d.document_link||""}}),{downloadExcel:p}=await Promise.resolve().then(()=>(B(),G));p(b,`Data_Training_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:async()=>{let o=[{Tanggal:"2026-01-08",Batch:"Batch 1",Materi:"Standar Kebersihan",Cabang:"001. Pondok Bambu",Trainer:"Budi",Peserta:"Rina, Agus",Nilai:"85",Dokumen:"https://link.doc"}],{downloadExcel:i}=await Promise.resolve().then(()=>(B(),G));i(o,"Template_Import_Training")},onImport:async o=>{let i=d=>{if(!d)return null;let m=String(d||"").toLowerCase(),h=e.data?.data.find(v=>String(v.full_name||"").toLowerCase()===m||String(v.code||"").toLowerCase()===m||String(v.name||"").toLowerCase()===m);return h?h.id:null},y=d=>{if(!d)return"";if(d instanceof Date&&!isNaN(d.getTime()))return d.toISOString().slice(0,10);let m=String(d).trim();if(/^\d{4,5}$/.test(m)){let v=Number(m);if(v>2e4&&v<99999){let k=new Date(Date.UTC(1899,11,30)+v*864e5);return isNaN(k.getTime())?"":k.toISOString().slice(0,10)}}if(/^\d{4}-\d{2}-\d{2}/.test(m))return m.slice(0,10);let h=m.split(/[\/\-\.]/);if(h.length===3){let[v,k,S]=h.map(w=>w.trim());if(v.length===4&&k.length<=2&&S.length<=2)return`${v}-${k.padStart(2,"0")}-${S.padStart(2,"0")}`;if(S.length===4&&k.length<=2&&v.length<=2)return`${S}-${k.padStart(2,"0")}-${v.padStart(2,"0")}`}return m},b=o.map(d=>({training_date:y(d.Tanggal),batch:String(d.Batch||"").trim(),subject:String(d.Materi||"").trim(),branch_id:i(String(d.Cabang||"").trim()),trainer:String(d.Trainer||"").trim(),participants:String(d.Peserta||"").trim(),score:d.Nilai?Number(d.Nilai):null,document_link:String(d.Dokumen||"").trim()})).filter(d=>d.training_date&&d.subject&&d.branch_id),p=await f("/api/training/import",{method:"POST",body:JSON.stringify(b)});if(!p.ok)throw new Error(p.data?.error||"Import gagal")}},formFields:o=>[{type:"row",fields:[{name:"training_date",label:"Tanggal Training",type:"date",required:!0,value:o?.training_date},{name:"batch",label:"Batch",placeholder:"Batch 1, Batch 2, ...",value:o?.batch}]},{name:"subject",label:"Materi / Topik Training",required:!0,placeholder:"Judul materi training",value:o?.subject},{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",options:o?.branch_id&&!t.find(i=>i.value==o.branch_id)?[...t,{value:o.branch_id,label:o.branch_name||o.branch_id}]:t,createApi:{path:"/api/branches",field:"full_name"},value:o?.branch_id},{name:"trainer",label:"Trainer",type:"combobox",options:r(o?.trainer),createApi:{path:"/api/pic",field:"name"},value:o?.trainer}]},{name:"participants",label:"Peserta (pisahkan dengan koma)",type:"textarea",rows:3,placeholder:"Nama Peserta 1, Nama Peserta 2, ...",value:(()=>{try{let i=JSON.parse(o?.participants);return Array.isArray(i)?i.join(", "):o?.participants||""}catch{return o?.participants||""}})()},{type:"row",fields:[{name:"score",label:"Nilai / Score",type:"number",step:"0.1",min:"0",max:"100",value:o?.score},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://...",value:o?.document_link}]},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:o?.notes}],onBeforeSubmit:async o=>(o.participants&&(o.participants=JSON.stringify(o.participants.split(",").map(i=>i.trim()).filter(Boolean))),o)})}I();B();async function Sa(a){let[e,s]=await Promise.all([f("/api/branches?all=1"),f(`/api/relievers${window.location.search?window.location.search+"&":"?"}limit=10000`)]),n=(e.data?.data||[]).map(r=>({value:r.id,label:r.full_name})),t=(s.data?.data||[]).map(r=>({value:r.full_name,label:r.full_name})),c=r=>r&&!t.find(l=>l.value===r)?[...t,{value:r,label:r}]:t,g=["Krishna Aryaan Permana","Agung Septiadi","Indra Saputro","Wariskin","Iqbal"],u=r=>{let l=g.map(o=>({value:o,label:o}));return r&&!l.find(o=>o.value===r)?[...l,{value:r,label:r}]:l};E({container:a,title:"Jadwal Reliefer",icon:"\u{1F504}",apiPath:"/api/relievers",bulkDelete:!0,itemLabel:"Reliefer",columns:[{key:"backup_date",label:"Tanggal Backup",nowrap:!0,render:r=>window.formatDate(r)},{key:"branch_name",label:"Cabang"},{key:"original_fc_name",label:"FC Digantikan"},{key:"period",label:"Periode",render:r=>W(r)},{key:"reliever_name",label:"Reliefer"},{key:"reason",label:"Keterangan"},{key:"shift",label:"Shift",render:r=>r?`<span class="badge badge-info">${r}</span>`:"-"},{key:"status",label:"Status",render:r=>O(r)}],filterFields:[{type:"search",placeholder:"Cari reliefer / FC..."},{type:"select",name:"branch_id",label:"Cabang",options:n},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"status",label:"Status",options:["Pending","Done","Tidak Datang"]}],formFields:r=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:r?.branch_id&&!n.find(l=>l.value==r.branch_id)?[...n,{value:r.branch_id,label:r.branch_name||r.branch_id}]:n,createApi:{path:"/api/branches",field:"full_name"},value:r?.branch_id},{name:"period",label:"Periode",type:"select",options:["Q1","Q2","Q3","Q4"],value:r?.period}]},{type:"row",fields:[{name:"original_fc_name",label:"FC yang Digantikan",type:"select",options:[{value:"",label:"BELUM ADA FC"},...c(r?.original_fc_name)],value:r?.original_fc_name},{name:"reliever_name",label:"Nama Reliefer",type:"select",required:!0,options:u(r?.reliever_name),value:r?.reliever_name}]},{type:"row",fields:[{name:"backup_date",label:"Tanggal Backup",type:"date",required:!0,value:r?.backup_date},{name:"completion_date",label:"Tanggal Selesai",type:"date",value:r?.completion_date}]},{type:"row",fields:[{name:"reason",label:"Keterangan",type:"select",options:["Cuti","Mengisi Kekosongan","Back Up Training","Deep Cleaning","Training Praktek Skill","Sakit","Lainnya"],value:r?.reason},{name:"shift",label:"Shift",type:"select",options:["Pagi","Siang","Full Shift","Middle"],value:r?.shift}]},{name:"status",label:"Status",type:"select",required:!0,options:["Pending","Done","Tidak Datang"],value:r?.status||""}],exportOptions:{moduleName:"relievers",onExport:async()=>{let r=await f(`/api/relievers${window.location.search?window.location.search+"&":"?"}limit=10000`);if(r.ok){let l=r.data.data.map(o=>({"Tanggal Backup":o.backup_date||"",Cabang:o.branch_name||"","FC Digantikan":o.original_fc_name||"",Periode:o.period||"",Reliefer:o.reliever_name||"",Keterangan:o.reason||"",Shift:o.shift||"","Tanggal Selesai":o.completion_date||"",Status:o.status||""}));$(l,"Data_Reliefer")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{$([{"Tanggal Backup":"2024-03-10",Cabang:"001. Pondok Bambu","FC Digantikan":"Budi Santoso",Periode:"Q1",Reliefer:"Andi",Keterangan:"Sakit",Shift:"Pagi","Tanggal Selesai":"2024-03-10",Status:"Done"}],"Template_Import_Reliefer")},onImport:async r=>{let o=(await f("/api/branches?all=1")).data?.data||[],i=p=>{if(!p)return null;let d=String(p||"").toLowerCase(),m=o.find(h=>String(h.full_name||"").toLowerCase()===d||String(h.code||"").toLowerCase()===d||String(h.name||"").toLowerCase()===d);return m?m.id:null},y=r.map(p=>({branch_id:i(String(p.Cabang||"").trim()),backup_date:String(p["Tanggal Backup"]||"").trim(),original_fc_name:String(p["FC Digantikan"]||"").trim(),reliever_name:String(p.Reliefer||"").trim(),period:String(p.Periode||"").trim(),reason:String(p.Keterangan||"").trim(),shift:String(p.Shift||"").trim(),completion_date:String(p["Tanggal Selesai"]||"").trim(),status:String(p.Status||"").trim()})).filter(p=>p.reliever_name&&p.backup_date),b=await f("/api/relievers/import",{method:"POST",body:JSON.stringify(y)});if(!b.ok)throw new Error(b.data?.error||"Import gagal")}}})}I();B();async function wa(a){let s=((await f("/api/branches?all=1")).data?.data||[]).map(t=>({value:t.id,label:t.full_name})),n=Array.from({length:4},(t,c)=>String(new Date().getFullYear()-c));E({container:a,title:"Laporan Inspeksi Hygiene",icon:"\u{1F50D}",apiPath:"/api/reports/inspection",itemLabel:"Laporan Inspeksi",bulkDelete:!0,columns:[{key:"branch_name",label:"Cabang"},{key:"period",label:"Periode",render:t=>W(t)},{key:"inspection_date",label:"Tanggal",nowrap:!0,render:t=>window.formatDate(t)},{key:"fc_score",label:"Point FC",render:t=>t!=null?`<strong class="${t>=80?"text-success":t>=60?"text-warning":"text-danger"}">${t}</strong>`:"-"},{key:"spv_score",label:"Point SPV",render:t=>t!=null?`<strong>${t}</strong>`:"-"},{key:"status",label:"Status",render:t=>O(t)},{key:"document_link",label:"Dokumen",render:t=>t?`<a href="${t}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari cabang / PIC..."},{type:"select",name:"branch_id",label:"Cabang",options:s},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"status",label:"Status",options:["Pending","Done"]},{type:"select",name:"year",label:"Tahun",options:n}],formFields:t=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:t?.branch_id&&!s.find(c=>c.value==t.branch_id)?[...s,{value:t.branch_id,label:t.branch_name||t.branch_id}]:s,createApi:{path:"/api/branches",field:"full_name"},value:t?.branch_id},{name:"period",label:"Periode",type:"select",required:!0,options:["Q1","Q2","Q3","Q4"],value:t?.period}]},{type:"row",fields:[{name:"inspection_date",label:"Tanggal Inspeksi",type:"date",required:!0,value:t?.inspection_date},{name:"status",label:"Status",type:"select",required:!0,options:["Pending","Done"],value:t?.status||""}]},{type:"row",fields:[{name:"fc_score",label:"Point FC",type:"number",step:"0.1",min:"0",max:"100",value:t?.fc_score},{name:"spv_score",label:"Point SPV",type:"number",step:"0.1",min:"0",max:"100",value:t?.spv_score}]},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://drive.google.com/...",value:t?.document_link},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:t?.notes}],exportOptions:{moduleName:"inspection_reports",onExport:async t=>{let c=new URLSearchParams(t||{}).toString(),g=await f(`/api/reports/inspection?limit=10000&${c}`);if(g.ok){let u=g.data.data.map(r=>({Cabang:r.branch_name||"",Periode:r.period||"",Tanggal:r.inspection_date||"","Point FC":r.fc_score!==null&&r.fc_score!==void 0?r.fc_score:"","Point SPV":r.spv_score!==null&&r.spv_score!==void 0?r.spv_score:"",Status:r.status||"","Link Dokumen":r.document_link||""}));$(u,`Laporan_Inspeksi_Hygiene_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{$([{Cabang:"001. Pondok Bambu",Periode:"Q1",Tanggal:"2026-01-08","Point FC":85,"Point SPV":90,Status:"Done","Link Dokumen":"https://drive.google.com/..."}],"Template_Import_Inspeksi")},onImport:async t=>{let g=(await f("/api/branches?all=1")).data?.data||[],u=i=>{if(!i)return null;let y=String(i||"").toLowerCase(),b=g.find(p=>String(p.full_name||"").toLowerCase()===y||String(p.code||"").toLowerCase()===y||String(p.name||"").toLowerCase()===y);return b?b.id:null},r=i=>{if(i==null||i==="")return"";if(i instanceof Date&&!isNaN(i.getTime()))return i.toISOString().slice(0,10);let y=String(i).trim();if(y===""||y==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(y))return y.slice(0,10);if(/^\d{4,5}$/.test(y)){let p=Number(y);if(p>2e4&&p<99999){let d=new Date(Date.UTC(1899,11,30)+p*864e5);return isNaN(d.getTime())?"":d.toISOString().slice(0,10)}}let b=y.split(/[\/\-\.]/);if(b.length===3){let[p,d,m]=b.map(h=>h.trim());if(p.length===4&&d.length<=2&&m.length<=2)return`${p}-${d.padStart(2,"0")}-${m.padStart(2,"0")}`;if(m.length===4&&d.length<=2&&p.length<=2)return`${m}-${d.padStart(2,"0")}-${p.padStart(2,"0")}`}return y},l=t.map(i=>({branch_id:u(String(i.Cabang||"").trim()),period:String(i.Periode||"").trim(),inspection_date:r(i.Tanggal),fc_score:i["Point FC"]!==void 0&&i["Point FC"]!==""?Number(i["Point FC"]):null,spv_score:i["Point SPV"]!==void 0&&i["Point SPV"]!==""?Number(i["Point SPV"]):null,status:String(i.Status||"").trim(),document_link:String(i["Link Dokumen"]||"").trim(),notes:String(i.Catatan||i.Keterangan||"").trim()})).filter(i=>i.branch_id&&i.period&&i.inspection_date),o=await f("/api/reports/inspection/import",{method:"POST",body:JSON.stringify(l)});if(!o.ok)throw new Error(o.data?.error||"Import gagal")}}})}I();B();async function _a(a){let s=((await f("/api/branches?all=1")).data?.data||[]).map(t=>({value:t.id,label:t.full_name})),n=Array.from({length:4},(t,c)=>String(new Date().getFullYear()-c));E({container:a,title:"Laporan General Cleaning & Deep Cleaning",icon:"\u{1F9F9}",apiPath:"/api/reports/cleaning",itemLabel:"Laporan GC/DC",bulkDelete:!0,columns:[{key:"branch_name",label:"Cabang"},{key:"activity_type",label:"Jenis",render:t=>`<span class="badge ${t==="Deep Cleaning"?"badge-purple":"badge-success"}">${t}</span>`},{key:"period",label:"Periode",render:t=>W(t)},{key:"activity_date",label:"Tanggal",nowrap:!0,render:t=>window.formatDate(t)},{key:"status",label:"Status",render:t=>O(t)},{key:"document_link",label:"Dokumen",render:t=>t?`<a href="${t}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari nama cabang/lokasi..."},{type:"select",name:"branch_id",label:"Cabang",options:s},{type:"select",name:"activity_type",label:"Jenis",options:["General Cleaning","Deep Cleaning"]},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"status",label:"Status",options:["Pending","Done"]},{type:"select",name:"year",label:"Tahun",options:n}],formFields:t=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:t?.branch_id&&!s.find(c=>c.value==t.branch_id)?[...s,{value:t.branch_id,label:t.branch_name||t.branch_id}]:s,createApi:{path:"/api/branches",field:"full_name"},value:t?.branch_id},{name:"activity_type",label:"Jenis Kegiatan",type:"select",required:!0,options:["General Cleaning","Deep Cleaning"],value:t?.activity_type}]},{type:"row",fields:[{name:"period",label:"Periode",type:"select",required:!0,options:["Q1","Q2","Q3","Q4"],value:t?.period},{name:"activity_date",label:"Tanggal",type:"date",required:!0,value:t?.activity_date}]},{name:"status",label:"Status",type:"select",required:!0,options:["Pending","Done"],value:t?.status||""},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://drive.google.com/...",value:t?.document_link},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:t?.notes}],exportOptions:{moduleName:"cleaning_reports",onExport:async t=>{let c=new URLSearchParams(t||{}).toString(),g=await f(`/api/reports/cleaning?limit=10000&${c}`);if(g.ok){let u=g.data.data.map(r=>({Cabang:r.branch_name||"",Jenis:r.activity_type||"",Periode:r.period||"",Tanggal:r.activity_date||"",Status:r.status||"","Link Dokumen":r.document_link||""}));$(u,`Laporan_GCDC_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{$([{Cabang:"001. Pondok Bambu",Jenis:"General Cleaning",Periode:"Q1",Tanggal:"2026-01-08",Status:"Done","Link Dokumen":"https://drive.google.com/..."}],"Template_Import_GCDC")},onImport:async t=>{let g=(await f("/api/branches?all=1")).data?.data||[],u=i=>{if(!i)return null;let y=String(i||"").toLowerCase(),b=g.find(p=>String(p.full_name||"").toLowerCase()===y||String(p.code||"").toLowerCase()===y||String(p.name||"").toLowerCase()===y);return b?b.id:null},r=i=>{if(i==null||i==="")return"";if(i instanceof Date&&!isNaN(i.getTime()))return i.toISOString().slice(0,10);let y=String(i).trim();if(y===""||y==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(y))return y.slice(0,10);if(/^\d{4,5}$/.test(y)){let p=Number(y);if(p>2e4&&p<99999){let d=new Date(Date.UTC(1899,11,30)+p*864e5);return isNaN(d.getTime())?"":d.toISOString().slice(0,10)}}let b=y.split(/[\/\-\.]/);if(b.length===3){let[p,d,m]=b.map(h=>h.trim());if(p.length===4&&d.length<=2&&m.length<=2)return`${p}-${d.padStart(2,"0")}-${m.padStart(2,"0")}`;if(m.length===4&&d.length<=2&&p.length<=2)return`${m}-${d.padStart(2,"0")}-${p.padStart(2,"0")}`}return y},l=t.map(i=>({branch_id:u(String(i.Cabang||"").trim()),activity_type:String(i.Jenis||i.Kegiatan||"").trim(),period:String(i.Periode||"").trim(),activity_date:r(i.Tanggal),status:String(i.Status||"").trim(),document_link:String(i["Link Dokumen"]||"").trim(),notes:String(i.Catatan||i.Keterangan||"").trim()})).filter(i=>i.branch_id&&i.activity_type&&i.period&&i.activity_date),o=await f("/api/reports/cleaning/import",{method:"POST",body:JSON.stringify(l)});if(!o.ok)throw new Error(o.data?.error||"Import gagal")}}})}I();B();async function xa(a){let s=((await f("/api/branches?all=1")).data?.data||[]).map(t=>({value:t.id,label:t.full_name})),n=Array.from({length:4},(t,c)=>String(new Date().getFullYear()-c));E({container:a,title:"Rekap Fogging",icon:"\u{1F4A8}",apiPath:"/api/reports/fogging",itemLabel:"Fogging",bulkDelete:!0,columns:[{key:"branch_name",label:"Cabang"},{key:"activity_type",label:"Jenis",render:t=>`<span class="badge badge-warning">${t}</span>`},{key:"period",label:"Periode",render:t=>W(t)},{key:"activity_date",label:"Tanggal",nowrap:!0,render:t=>window.formatDate(t)},{key:"status",label:"Status",render:t=>O(t)},{key:"document_link",label:"Dokumen",render:t=>t?`<a href="${t}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"},{key:"notes",label:"Catatan",render:t=>t||"-"}],filterFields:[{type:"search",placeholder:"Cari nama cabang/lokasi..."},{type:"select",name:"branch_id",label:"Cabang",options:s},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"status",label:"Status",options:["Pending","Done"]},{type:"select",name:"year",label:"Tahun",options:n}],formFields:t=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:t?.branch_id&&!s.find(c=>c.value==t.branch_id)?[...s,{value:t.branch_id,label:t.branch_name||t.branch_id}]:s,createApi:{path:"/api/branches",field:"full_name"},value:t?.branch_id},{name:"period",label:"Periode",type:"select",required:!0,options:["Q1","Q2","Q3","Q4"],value:t?.period}]},{type:"row",fields:[{name:"activity_date",label:"Tanggal",type:"date",value:t?.activity_date},{name:"status",label:"Status",type:"select",required:!0,options:["Pending","Done"],value:t?.status||""}]},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://...",value:t?.document_link},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:t?.notes}],exportOptions:{moduleName:"fogging_reports",onExport:async t=>{let c=new URLSearchParams(t||{}).toString(),g=await f(`/api/reports/fogging?limit=10000&${c}`);if(g.ok){let u=g.data.data.map(r=>({Cabang:r.branch_name||"",Jenis:r.activity_type||"Fogging",Periode:r.period||"",Tanggal:r.activity_date||"",Status:r.status||"","Link Dokumen":r.document_link||""}));$(u,`Laporan_Fogging_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{$([{Cabang:"001. Pondok Bambu",Jenis:"Fogging",Periode:"Q1",Tanggal:"2026-01-08",Status:"Done","Link Dokumen":"https://drive.google.com/..."}],"Template_Import_Fogging")},onImport:async t=>{let g=(await f("/api/branches?all=1")).data?.data||[],u=i=>{if(!i)return null;let y=String(i||"").toLowerCase(),b=g.find(p=>String(p.full_name||"").toLowerCase()===y||String(p.code||"").toLowerCase()===y||String(p.name||"").toLowerCase()===y);return b?b.id:null},r=i=>{if(i==null||i==="")return"";if(i instanceof Date&&!isNaN(i.getTime()))return i.toISOString().slice(0,10);let y=String(i).trim();if(y===""||y==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(y))return y.slice(0,10);if(/^\d{4,5}$/.test(y)){let p=Number(y);if(p>2e4&&p<99999){let d=new Date(Date.UTC(1899,11,30)+p*864e5);return isNaN(d.getTime())?"":d.toISOString().slice(0,10)}}let b=y.split(/[\/\-\.]/);if(b.length===3){let[p,d,m]=b.map(h=>h.trim());if(p.length===4&&d.length<=2&&m.length<=2)return`${p}-${d.padStart(2,"0")}-${m.padStart(2,"0")}`;if(m.length===4&&d.length<=2&&p.length<=2)return`${m}-${d.padStart(2,"0")}-${p.padStart(2,"0")}`}return y},l=t.map(i=>({branch_id:u(String(i.Cabang||"").trim()),activity_type:String(i.Jenis||i.Kegiatan||"Fogging").trim(),period:String(i.Periode||"").trim(),activity_date:r(i.Tanggal),status:String(i.Status||"").trim(),document_link:String(i["Link Dokumen"]||"").trim(),notes:String(i.Catatan||i.Keterangan||"").trim()})).filter(i=>i.branch_id&&i.period&&i.activity_date),o=await f("/api/reports/fogging/import",{method:"POST",body:JSON.stringify(l)});if(!o.ok)throw new Error(o.data?.error||"Import gagal")}}})}I();B();async function Ca(a){let[e,s,n]=await Promise.all([f("/api/branches?all=1"),f(`/api/basecamp_reports${window.location.search?window.location.search+"&":"?"}limit=10000`),f(`/api/basecamp_reports${window.location.search?window.location.search+"&":"?"}limit=10000`)]),t=(e.data?.data||[]).map(l=>({value:l.id,label:l.full_name})),c=(s.data?.data||[]).map(l=>({value:l.full_name,label:l.full_name})),g=(n.data?.data||[]).filter(l=>l.role==="FC Spesialis").map(l=>({value:l.name,label:l.name})),u=l=>l&&!c.find(o=>o.value===l)?[...c,{value:l,label:l}]:c,r=l=>l&&!g.find(o=>o.value===l)?[...g,{value:l,label:l}]:g;E({container:a,title:"Rekap Laporan Basecamp",icon:"\u{1F4DD}",apiPath:"/api/reports/basecamp",bulkDelete:!0,itemLabel:"Laporan Basecamp",columns:[{key:"info_date",label:"Tgl Info",nowrap:!0,render:l=>window.formatDate(l)},{key:"branch_name",label:"Cabang"},{key:"problem",label:"Permasalahan",render:l=>`<span title="${l||""}">${l?.length>60?l.slice(0,60)+"\u2026":l||"-"}</span>`},{key:"pic",label:"PIC"},{key:"done_date",label:"Tgl Done",nowrap:!0,render:l=>window.formatDate(l)},{key:"status",label:"Status",render:l=>O(l)},{key:"notes",label:"Keterangan",render:l=>l?.length>40?l.slice(0,40)+"\u2026":l||"-"}],filterFields:[{type:"search",placeholder:"Cari permasalahan / PIC..."},{type:"select",name:"branch_id",label:"Cabang",options:t},{type:"select",name:"status",label:"Status",options:["Open","In Progress","Done"]}],formFields:l=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:l?.branch_id&&!t.find(o=>o.value==l.branch_id)?[...t,{value:l.branch_id,label:l.branch_name||l.branch_id}]:t,createApi:{path:"/api/branches",field:"full_name"},value:l?.branch_id},{name:"pic",label:"PIC",type:"combobox",options:r(l?.pic),createApi:{path:"/api/pic",field:"name"},value:l?.pic}]},{name:"problem",label:"Permasalahan",type:"textarea",required:!0,rows:3,value:l?.problem},{type:"row",fields:[{name:"info_date",label:"Tanggal Info",type:"date",required:!0,value:l?.info_date},{name:"done_date",label:"Tanggal Done",type:"date",value:l?.done_date}]},{name:"status",label:"Status",type:"select",options:["Open","In Progress","Done"],value:l?.status||"Open"},{name:"notes",label:"Keterangan / Tindak Lanjut",type:"textarea",rows:2,value:l?.notes}],exportOptions:{moduleName:"basecamp_reports",onExport:async l=>{let o=new URLSearchParams(l||{}).toString(),i=await f(`/api/reports/basecamp?limit=10000&${o}`);if(i.ok){let y=i.data.data.map(b=>({"Tgl Info":b.info_date||"",Cabang:b.branch_name||"",Permasalahan:b.problem||"",PIC:b.pic||"","Tgl Done":b.done_date||"",Status:b.status||"",Keterangan:b.notes||""}));$(y,`Rekap_Laporan_Basecamp_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{$([{"Tgl Info":"2026-01-08",Cabang:"001. Pondok Bambu",Permasalahan:"Request fogging karena banyak nyamuk",PIC:"Fajar","Tgl Done":"2026-01-10",Status:"Done",Keterangan:"Sudah difogging"}],"Template_Import_Basecamp")},onImport:async l=>{let i=(await f("/api/branches?all=1")).data?.data||[],y=m=>{if(!m)return null;let h=String(m||"").toLowerCase(),v=i.find(k=>String(k.full_name||"").toLowerCase()===h||String(k.code||"").toLowerCase()===h||String(k.name||"").toLowerCase()===h);return v?v.id:null},b=m=>{if(m==null||m==="")return"";if(m instanceof Date&&!isNaN(m.getTime()))return m.toISOString().slice(0,10);let h=String(m).trim();if(h===""||h==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(h))return h.slice(0,10);if(/^\d{4,5}$/.test(h)){let k=Number(h);if(k>2e4&&k<99999){let S=new Date(Date.UTC(1899,11,30)+k*864e5);return isNaN(S.getTime())?"":S.toISOString().slice(0,10)}}let v=h.split(/[\/\-\.]/);if(v.length===3){let[k,S,w]=v.map(T=>T.trim());if(k.length===4&&S.length<=2&&w.length<=2)return`${k}-${S.padStart(2,"0")}-${w.padStart(2,"0")}`;if(w.length===4&&S.length<=2&&k.length<=2)return`${w}-${S.padStart(2,"0")}-${k.padStart(2,"0")}`}return h},p=l.map(m=>({info_date:b(m["Tgl Info"]||m["Tanggal Info"]),branch_id:y(String(m.Cabang||"").trim()),problem:String(m.Permasalahan||"").trim(),pic:String(m.PIC||"").trim(),done_date:b(m["Tgl Done"]||m["Tanggal Done"]),status:String(m.Status||"").trim(),notes:String(m.Keterangan||m.Catatan||"").trim()})).filter(m=>m.info_date&&m.branch_id&&m.problem),d=await f("/api/reports/basecamp/import",{method:"POST",body:JSON.stringify(p)});if(!d.ok)throw new Error(d.data?.error||"Import gagal")}}})}async function Ta(a){E({container:a,title:"SOP",icon:"\u{1F4D6}",apiPath:"/api/sop",bulkDelete:!0,itemLabel:"SOP",columns:[{key:"name",label:"Nama SOP"},{key:"category",label:"Kategori"},{key:"document_link",label:"Dokumen",render:e=>e?`<a href="${e}" target="_blank" rel="noopener" class="btn btn-sm btn-primary">\u{1F4C4} Buka Dokumen</a>`:"-"},{key:"notes",label:"Catatan"}],filterFields:[{type:"search",placeholder:"Cari nama SOP..."}],exportOptions:{moduleName:"sop",onExport:async e=>{let s=new URLSearchParams(e||{}).toString(),{apiFetch:n}=await Promise.resolve().then(()=>(I(),ne)),t=await n(`/api/sop?limit=10000&${s}`);if(t.ok){let c=t.data.data.map(u=>({"Nama SOP":u.name||"",Kategori:u.category||"",Dokumen:u.document_link||"",Catatan:u.notes||u.description||""})),{downloadExcel:g}=await Promise.resolve().then(()=>(B(),G));g(c,`Master_SOP_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:async()=>{let e=[{"Nama SOP":"SOP Cuci Tangan",Kategori:"Ketentuan & Basic",Dokumen:"https://link.com",Catatan:"Catatan singkat"}],{downloadExcel:s}=await Promise.resolve().then(()=>(B(),G));s(e,"Template_Import_SOP")},onImport:async e=>{let s=e.map(c=>({name:String(c["Nama SOP"]||"").trim(),category:String(c.Kategori||"").trim(),document_link:String(c.Dokumen||"").trim(),description:String(c.Catatan||"").trim()})).filter(c=>c.name),{apiFetch:n}=await Promise.resolve().then(()=>(I(),ne)),t=await n("/api/sop/import",{method:"POST",body:JSON.stringify(s)});if(!t.ok)throw new Error(t.data?.error||"Import gagal")}},formFields:e=>[{name:"name",label:"Nama SOP",required:!0,placeholder:"Nama SOP",value:e?.name},{name:"category",label:"Kategori",placeholder:"Ketentuan & Basic, Kualitas & Grooming, dst.",value:e?.category},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://docs.google.com/...",value:e?.document_link},{name:"description",label:"Deskripsi / Catatan",type:"textarea",rows:3,value:e?.description}]})}async function $a(a){E({container:a,title:"Master Checklist",icon:"\u2705",apiPath:"/api/checklist",bulkDelete:!0,itemLabel:"Checklist",columns:[{key:"name",label:"Nama Checklist"},{key:"category",label:"Kategori"},{key:"document_link",label:"Dokumen",render:e=>e?`<a href="${e}" target="_blank" rel="noopener" class="btn btn-sm btn-primary">\u{1F4C4} Buka Dokumen</a>`:"-"},{key:"description",label:"Deskripsi"}],filterFields:[{type:"search",placeholder:"Cari checklist..."}],exportOptions:{moduleName:"checklist",onExport:async e=>{let s=new URLSearchParams(e||{}).toString(),{apiFetch:n}=await Promise.resolve().then(()=>(I(),ne)),t=await n(`/api/checklist?limit=10000&${s}`);if(t.ok){let c=t.data.data.map(u=>({"Nama Checklist":u.name||"",Kategori:u.category||"",Dokumen:u.document_link||"",Deskripsi:u.description||""})),{downloadExcel:g}=await Promise.resolve().then(()=>(B(),G));g(c,`Master_Checklist_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:async()=>{let e=[{"Nama Checklist":"Checklist Kebersihan Mingguan",Kategori:"Master Cleaning Program",Dokumen:"https://link.com",Deskripsi:"Deskripsi singkat"}],{downloadExcel:s}=await Promise.resolve().then(()=>(B(),G));s(e,"Template_Import_Checklist")},onImport:async e=>{let s=e.map(c=>({name:String(c["Nama Checklist"]||"").trim(),category:String(c.Kategori||"").trim(),document_link:String(c.Dokumen||"").trim(),description:String(c.Deskripsi||"").trim()})).filter(c=>c.name),{apiFetch:n}=await Promise.resolve().then(()=>(I(),ne)),t=await n("/api/checklist/import",{method:"POST",body:JSON.stringify(s)});if(!t.ok)throw new Error(t.data?.error||"Import gagal")}},formFields:e=>[{name:"name",label:"Nama Checklist",required:!0,placeholder:"Nama checklist",value:e?.name},{name:"category",label:"Kategori",placeholder:"Master Cleaning Program, dll.",value:e?.category},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://docs.google.com/...",value:e?.document_link},{name:"description",label:"Deskripsi",type:"textarea",rows:3,value:e?.description}]})}I();ge();B();async function Xe(a,e="forms"){if(e==="supply")return pt(a);dt(a)}function dt(a){E({container:a,title:"Master Form",icon:"\u{1F4C4}",apiPath:"/api/forms",bulkDelete:!0,itemLabel:"Form",columns:[{key:"name",label:"Nama Form"},{key:"category",label:"Kategori"},{key:"document_link",label:"Dokumen",render:e=>e?`<a href="${e}" target="_blank" rel="noopener" class="btn btn-sm btn-primary">\u{1F4C4} Buka</a>`:"-"},{key:"is_public",label:"Publik",render:e=>e?'<span class="badge badge-success">Ya</span>':'<span class="badge badge-neutral">Tidak</span>'},{key:"description",label:"Deskripsi"}],filterFields:[{type:"search",placeholder:"Cari form..."}],formFields:e=>[{name:"name",label:"Nama Form",required:!0,placeholder:"Nama form",value:e?.name},{name:"category",label:"Kategori",placeholder:"Permintaan Barang, Penilaian, dll.",value:e?.category},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://...",value:e?.document_link},{name:"description",label:"Deskripsi",type:"textarea",rows:2,value:e?.description},{name:"is_public",label:"Akses Publik",type:"checkbox",checkLabel:"Form dapat diakses tanpa login",value:e?.is_public}]})}async function pt(a){let s=((await f("/api/branches?all=1")).data?.data||[]).map(n=>({value:n.id,label:n.full_name}));E({container:a,title:"Permintaan Barang & Chemical",icon:"\u{1F4E6}",apiPath:"/api/reports/supply",bulkDelete:!0,itemLabel:"Permintaan",canCreate:!0,columns:[{key:"submitted_at",label:"Waktu",nowrap:!0,render:n=>n?new Date(n).toLocaleString("id-ID"):"-"},{key:"submitter_name",label:"Pengirim"},{key:"branch_name",label:"Cabang",render:(n,t)=>t.branch_name_ref||t.branch_name||"-"},{key:"tools_items",label:"Alat/Barang",render:n=>{try{let t=JSON.parse(n);return Array.isArray(t)?t.join(", "):n}catch{return n||"-"}}},{key:"chemical_items",label:"Chemical",render:n=>{try{let t=JSON.parse(n);return Array.isArray(t)?t.join(", "):n}catch{return n||"-"}}},{key:"additional_notes",label:"Catatan",render:n=>n?.length>40?n.slice(0,40)+"\u2026":n||"-"},{key:"status",label:"Status",render:n=>O(n)},{key:"processed_by",label:"Diproses Oleh"}],filterFields:[{type:"select",name:"status",label:"Status",options:["Pending","Diproses","Selesai"]}],formFields:n=>{let t=n?.tools_items;try{t=Array.isArray(JSON.parse(t))?JSON.parse(t).join(", "):t}catch{}let c=n?.chemical_items;try{c=Array.isArray(JSON.parse(c))?JSON.parse(c).join(", "):c}catch{}return[{type:"row",fields:[{name:"submitter_name",label:"Nama Pengirim",required:!0,value:n?.submitter_name},{name:"branch_id",label:"Cabang",type:"combobox",options:n?.branch_id&&!s.find(g=>g.value==n.branch_id)?[...s,{value:n.branch_id,label:n.branch_name||n.branch_id}]:s,createApi:{path:"/api/branches",field:"full_name"},value:n?.branch_id}]},{type:"row",fields:[{name:"tools_items",label:"Alat / Barang",placeholder:"Pisahkan dengan koma (Sapu, Mop)",value:t},{name:"tools_quantity",label:"Jumlah Alat",type:"number",value:n?.tools_quantity}]},{type:"row",fields:[{name:"chemical_items",label:"Chemical",placeholder:"Pisahkan dengan koma",value:c},{name:"chemical_quantity",label:"Jumlah Chemical",type:"number",value:n?.chemical_quantity}]},{name:"additional_notes",label:"Catatan",type:"textarea",rows:2,value:n?.additional_notes},{name:"status",label:"Status",type:"select",options:["Pending","Diproses","Selesai"],value:n?.status||""},{name:"processed_by",label:"Diproses Oleh",value:n?.processed_by}]},exportOptions:{moduleName:"supply_requests",onExport:async n=>{let t=new URLSearchParams(n||{}).toString(),c=await f(`/api/reports/supply?limit=10000&${t}`);if(c.ok){let g=c.data.data.map(u=>{let r=u.tools_items;try{r=Array.isArray(JSON.parse(r))?JSON.parse(r).join(", "):r}catch{}let l=u.chemical_items;try{l=Array.isArray(JSON.parse(l))?JSON.parse(l).join(", "):l}catch{}return{Waktu:u.submitted_at||"",Pengirim:u.submitter_name||"",Cabang:u.branch_name_ref||u.branch_name||"","Alat/Barang":r||"",Chemical:l||"",Catatan:u.additional_notes||"",Status:u.status||"","Diproses Oleh":u.processed_by||""}});$(g,`Permintaan_Barang_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{$([{Waktu:"2026-01-08",Pengirim:"Fajar",Cabang:"001. Pondok Bambu","Alat/Barang":"Sapu, Mop",Chemical:"Karbol",Catatan:"Mendesak",Status:"Pending","Diproses Oleh":""}],"Template_Import_Permintaan")},onImport:async n=>{let c=(await f("/api/branches?all=1")).data?.data||[],g=o=>{if(!o)return null;let i=String(o||"").toLowerCase(),y=c.find(b=>String(b.full_name||"").toLowerCase()===i||String(b.code||"").toLowerCase()===i||String(b.name||"").toLowerCase()===i);return y?y.id:null},u=o=>{if(o==null||o==="")return"";if(o instanceof Date&&!isNaN(o.getTime()))return o.toISOString().slice(0,10);let i=String(o).trim();if(i===""||i==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(i))return i.slice(0,10);if(/^\d{4,5}$/.test(i)){let b=Number(i);if(b>2e4&&b<99999){let p=new Date(Date.UTC(1899,11,30)+b*864e5);return isNaN(p.getTime())?"":p.toISOString().slice(0,10)}}let y=i.split(/[\/\-\.]/);if(y.length===3){let[b,p,d]=y.map(m=>m.trim());if(b.length===4&&p.length<=2&&d.length<=2)return`${b}-${p.padStart(2,"0")}-${d.padStart(2,"0")}`;if(d.length===4&&p.length<=2&&b.length<=2)return`${d}-${p.padStart(2,"0")}-${b.padStart(2,"0")}`}return i},r=n.map(o=>({submitted_at:u(o.Waktu||o.Tanggal),submitter_name:String(o.Pengirim||"").trim(),branch_id:g(String(o.Cabang||"").trim()),tools_items:String(o["Alat/Barang"]||o.Alat||"").trim(),chemical_items:String(o.Chemical||"").trim(),additional_notes:String(o.Catatan||o.Keterangan||"").trim(),status:String(o.Status||"").trim(),processed_by:String(o["Diproses Oleh"]||o.PIC||"").trim()})).filter(o=>o.submitted_at&&o.submitter_name&&o.branch_id),l=await f("/api/reports/supply/import",{method:"POST",body:JSON.stringify(r)});if(!l.ok)throw new Error(l.data?.error||"Import gagal")}},extraActions:[{label:"Update Status",icon:"\u{1F504}",class:"btn-secondary",handler:(n,t)=>{let c=Y({title:"Update Status Permintaan",content:`
              <div class="form-group">
                <label class="form-label">Status</label>
                <select class="form-control" id="supply-status">
                  <option value="Pending" ${n.status==="Pending"?"selected":""}>Pending</option>
                  <option value="Diproses" ${n.status==="Diproses"?"selected":""}>Diproses</option>
                  <option value="Selesai" ${n.status==="Selesai"?"selected":""}>Selesai</option>
                </select>
              </div>
              <div class="form-group">
                <label class="form-label">Diproses Oleh</label>
                <input type="text" class="form-control" id="supply-processed-by" value="${n.processed_by||""}" placeholder="Nama">
              </div>
            `,onConfirm:async(g,u)=>{let r=g.querySelector("#supply-status").value,l=g.querySelector("#supply-processed-by").value;(await f(`/api/reports/supply/${n.id}`,{method:"PUT",body:JSON.stringify({status:r,processed_by:l})})).ok?(q("Status diperbarui."),u(),t()):U("Gagal update status.")}})}}]})}I();B();async function Ea(a){let e=ae();if(!e||!["superadmin","admin"].includes(e.role)){a.innerHTML='<div class="empty-state"><p class="text-danger">Akses ditolak.</p></div>';return}E({container:a,title:"Manajemen User",icon:"\u{1F510}",apiPath:"/api/users",bulkDelete:!0,itemLabel:"User",columns:[{key:"full_name",label:"Nama Lengkap"},{key:"username",label:"Username"},{key:"email",label:"Email"},{key:"role",label:"Role",render:s=>`<span class="badge ${{superadmin:"badge-danger",admin:"badge-purple",manager:"badge-info",spv:"badge-secondary",viewer:"badge-neutral"}[s]||"badge-neutral"}">${s}</span>`},{key:"is_active",label:"Status",render:s=>s?'<span class="badge badge-success">Aktif</span>':'<span class="badge badge-neutral">Nonaktif</span>'},{key:"created_at",label:"Dibuat",nowrap:!0,render:s=>s?new Date(s).toLocaleDateString("id-ID"):"-"}],filterFields:[{type:"search",placeholder:"Cari nama / username..."}],formFields:s=>{let n=!!s;return[{type:"row",fields:[{name:"full_name",label:"Nama Lengkap",required:!0,placeholder:"Nama lengkap",value:s?.full_name},{name:"username",label:"Username",required:!n,placeholder:"username",value:s?.username}]},{type:"row",fields:[{name:"email",label:"Email",type:"email",required:!n,placeholder:"email@contoh.com",value:s?.email},{name:"role",label:"Role",type:"select",required:!0,options:[{value:"superadmin",label:"Super Admin"},{value:"admin",label:"Admin"},{value:"manager",label:"Manager"},{value:"spv",label:"Supervisor"},{value:"viewer",label:"Viewer"}],value:s?.role||"viewer"}]},{type:"row",fields:[{name:"password",label:n?"Password Baru (kosongkan jika tidak diubah)":"Password",type:"password",required:!n,placeholder:"Min. 6 karakter"},{name:"is_active",label:"Status Aktif",type:"checkbox",checkLabel:"User aktif",value:n?s?.is_active:1}]}]},exportOptions:{moduleName:"users",onExport:async()=>{let s=await f(`/api/users${window.location.search?window.location.search+"&":"?"}limit=10000`);if(s.ok){let n=s.data.data.map(t=>({"Nama Lengkap":t.full_name||"",Username:t.username||"",Email:t.email||"",Role:t.role||"",Status:t.is_active?"Aktif":"Nonaktif"}));$(n,"Data_Users")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{$([{"Nama Lengkap":"Admin Cabang",Username:"admin01",Email:"admin@contoh.com",Role:"admin",Password:"password123"}],"Template_Import_Users")},onImport:async s=>{let n=s.map(c=>({full_name:String(c["Nama Lengkap"]||"").trim(),username:String(c.Username||"").trim(),email:String(c.Email||"").trim(),role:String(c.Role||"").trim()||"viewer",password:String(c.Password||"").trim()})).filter(c=>c.username&&c.password&&c.email&&c.full_name),t=await f("/api/users/import",{method:"POST",body:JSON.stringify(n)});if(!t.ok)throw new Error(t.data?.error||"Import gagal")}}})}I();B();async function Pa(a){E({container:a,title:"Manajemen Cabang",icon:"\u{1F3E2}",apiPath:"/api/branches",itemLabel:"Cabang",bulkDelete:!0,columns:[{key:"code",label:"Kode",width:"60px"},{key:"full_name",label:"Nama Cabang"},{key:"city",label:"Kota"},{key:"is_active",label:"Status",render:e=>e?'<span class="badge badge-success">Aktif</span>':'<span class="badge badge-neutral">Nonaktif</span>'}],filterFields:[{type:"search",placeholder:"Cari nama / kode cabang..."}],formFields:e=>[{type:"row",fields:[{name:"code",label:"Kode Cabang",required:!0,placeholder:"001, A01, ...",value:e?.code},{name:"name",label:"Nama Pendek",required:!0,placeholder:"Pondok Bambu",value:e?.name}]},{name:"full_name",label:"Nama Lengkap",required:!0,placeholder:"001. Pondok Bambu",value:e?.full_name},{type:"row",fields:[{name:"city",label:"Kota",placeholder:"Jakarta",value:e?.city},{name:"is_active",label:"Status",type:"checkbox",checkLabel:"Cabang aktif",value:e?.is_active!==void 0?e.is_active:1}]}],exportOptions:{moduleName:"branches",onExport:async()=>{let e=await f(`/api/branches${window.location.search?window.location.search+"&":"?"}limit=10000`);if(e.ok)$(e.data.data,"Data_Cabang");else throw new Error("Gagal mengambil data")},onTemplate:()=>{$([{"Kode Cabang":"001","Nama Pendek":"Pondok Bambu","Nama Lengkap":"001. Pondok Bambu",Kota:"Jakarta Timur"},{"Kode Cabang":"002","Nama Pendek":"Bintaro","Nama Lengkap":"002. Bintaro",Kota:"Tangerang Selatan"}],"Template_Import_Cabang")},onImport:async e=>{let s=e.map(t=>({code:String(t["Kode Cabang"]||"").trim(),name:String(t["Nama Pendek"]||"").trim(),full_name:String(t["Nama Lengkap"]||"").trim(),city:String(t.Kota||"").trim()})).filter(t=>t.code&&t.name),n=await f("/api/branches/import",{method:"POST",body:JSON.stringify(s)});if(!n.ok)throw new Error(n.data?.error||"Import gagal")}}})}I();async function Da(a){let e=new Date,s=[];a.innerHTML=`
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
          <label class="filter-check"><input type="checkbox" value="issue"           checked class="cal-filter"> Permasalahan</label>
          <label class="filter-check"><input type="checkbox" value="reliever"        checked class="cal-filter"> Reliefer</label>
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
  `,document.getElementById("cal-prev").addEventListener("click",()=>{e.setMonth(e.getMonth()-1),t()}),document.getElementById("cal-next").addEventListener("click",()=>{e.setMonth(e.getMonth()+1),t()}),document.getElementById("cal-event-close").addEventListener("click",()=>{document.getElementById("cal-event-list").style.display="none"}),document.querySelectorAll(".cal-filter").forEach(c=>c.addEventListener("change",t));async function n(){try{let c=`${e.getFullYear()}-${String(e.getMonth()+1).padStart(2,"0")}`;s=(await f(`/api/dashboard/calendar?month=${c}`)).data?.data||[]}catch(c){console.warn("[Calendar] Failed to load events, rendering empty grid:",c),s=[]}}async function t(){let c=document.getElementById("calendar-grid");if(c){c.innerHTML=`<div style="display:grid;grid-template-columns:repeat(7,1fr);gap:1px;background:var(--border);">
      ${Array(35).fill('<div style="background:#f8fafc;min-height:70px;"></div>').join("")}
    </div>`,await n();try{let g=e.getFullYear(),u=e.getMonth(),r=e.toLocaleDateString("id-ID",{month:"long",year:"numeric"}),l=document.getElementById("cal-month-label");l&&(l.textContent=r);let o=new Set(Array.from(document.querySelectorAll(".cal-filter:checked")).map(S=>S.value)),i=s.filter(S=>o.has(S.type)),y={};i.forEach(S=>{let w=(S.event_date||"").slice(0,10);y[w]||(y[w]=[]),y[w].push(S)});let b=new Date(g,u,1).getDay(),p=new Date(g,u+1,0).getDate(),d=["Min","Sen","Sel","Rab","Kam","Jum","Sab"],m=new Date().toISOString().slice(0,10),h='<div class="calendar-grid">';d.forEach(S=>{h+=`<div class="cal-day-header">${S}</div>`});for(let S=0;S<b;S++)h+='<div class="cal-cell cal-cell-empty"></div>';for(let S=1;S<=p;S++){let w=`${g}-${String(u+1).padStart(2,"0")}-${String(S).padStart(2,"0")}`,T=y[w]||[],L=w===m;h+=`
          <div class="cal-cell ${L?"cal-today":""} ${T.length?"cal-has-events":""}"
               data-date="${w}" tabindex="0" role="button" aria-label="${w}">
            <div class="cal-day-num ${L?"today-num":""}">${S}</div>
            <div class="cal-events-preview">
              ${T.slice(0,3).map(_=>`
                <div class="cal-event-dot cal-color-${_.color||"gray"}" title="${Be(_.title||_.type)}">
                  <span class="cal-event-dot-label">${mt(_.title||_.branch_name||_.type,18)}</span>
                </div>
              `).join("")}
              ${T.length>3?`<div class="cal-more">+${T.length-3} lagi</div>`:""}
            </div>
          </div>`}let k=(b+p)%7;if(k!==0)for(let S=0;S<7-k;S++)h+='<div class="cal-cell cal-cell-empty"></div>';h+="</div>",c.innerHTML=h,c.querySelectorAll(".cal-cell[data-date]").forEach(S=>{S.addEventListener("click",()=>{let w=S.dataset.date,T=y[w]||[];if(!T.length)return;let L=document.getElementById("cal-event-list"),_=new Date(w+"T00:00:00").toLocaleDateString("id-ID",{weekday:"long",day:"numeric",month:"long",year:"numeric"});document.getElementById("cal-event-date").textContent=_,document.getElementById("cal-event-items").innerHTML=T.map(D=>`
            <div class="cal-event-item cal-color-border-${D.color||"gray"}">
              <div class="cal-event-type">${ut(D.type)}</div>
              <div class="cal-event-title">${Be(D.title||"-")}</div>
              <div class="cal-event-branch">${Be(D.branch_name||"")}</div>
              ${D.status?`<div class="cal-event-status">${Be(D.status)}</div>`:""}
              ${D.days_remaining!==void 0?`<div class="cal-event-extra">Sisa: ${D.days_remaining} hari</div>`:""}
            </div>
          `).join(""),L.style.display="block"})})}catch(g){console.error("[Calendar] Render error:",g),c&&(c.innerHTML=`
          <div style="padding:40px;text-align:center;color:var(--text-3)">
            <div style="font-size:2rem;margin-bottom:8px">\u{1F4C5}</div>
            <div>Gagal memuat kalender. Silakan refresh.</div>
          </div>`)}}}t()}function mt(a,e){return a?a.length>e?a.slice(0,e)+"\u2026":a:""}function Be(a){return a?String(a).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;"):""}function ut(a){return{schedule:"\u{1F5D3} Jadwal",issue:"\u26A0\uFE0F Permasalahan",reliever:"\u{1F504} Reliefer",training:"\u{1F393} Training",contract_expiry:"\u{1F4CB} Kontrak Habis"}[a]||a}I();async function Ia(a){let e=ae(),s=(e?.full_name||e?.username||"U")[0].toUpperCase(),t={superadmin:"#7C3AED",admin:"#2563EB",manager:"#0891B2",spv:"#059669",viewer:"#64748B"}[e?.role]||"#64748B";a.innerHTML=`
    <div class="page-header">
      <h1 class="page-title">\u{1F464} Profil Saya</h1>
    </div>

    <div class="profile-layout">

      <!-- LEFT: Info Card -->
      <div class="chart-card profile-info-card">
        <div class="profile-avatar-wrap">
          <div class="profile-avatar-xl" style="background:linear-gradient(135deg,${t},${t}99)">
            ${s}
          </div>
          <div class="profile-name-block">
            <div class="profile-fullname">${e?.full_name||"\u2014"}</div>
            <div class="profile-username">@${e?.username||"\u2014"}</div>
            <span class="badge badge-info" style="background:${t}18;color:${t};margin-top:6px">
              ${(e?.role||"viewer").toUpperCase()}
            </span>
          </div>
        </div>

        <hr class="profile-divider">

        <div class="info-list">
          <div class="info-row">
            <span class="info-key">\u{1F4E7} Email</span>
            <span class="info-value">${e?.email||"\u2014"}</span>
          </div>
          <div class="info-row">
            <span class="info-key">\u{1F464} Username</span>
            <span class="info-value">${e?.username||"\u2014"}</span>
          </div>
          <div class="info-row">
            <span class="info-key">\u{1F3AF} Role</span>
            <span class="info-value" style="color:${t};font-weight:700">${e?.role||"\u2014"}</span>
          </div>
        </div>
      </div>

      <!-- RIGHT: Change Password -->
      <div class="chart-card">
        <div class="chart-card-header">
          <div>
            <div class="chart-card-title">\u{1F511} Ganti Password</div>
            <div class="chart-card-subtitle">Gunakan password yang kuat, minimal 6 karakter</div>
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
              required placeholder="Minimal 6 karakter" autocomplete="new-password">
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
  `;let c=localStorage.getItem("fm_token"),g=document.getElementById("session-info");if(c&&g)try{let u=JSON.parse(atob(c.split(".")[1])),r=new Date(u.exp*1e3);g.textContent=`Berakhir: ${r.toLocaleString("id-ID")}`}catch{g.textContent="Tidak tersedia"}document.getElementById("change-pwd-form")?.addEventListener("submit",async u=>{u.preventDefault();let r=document.getElementById("pwd-error"),l=document.getElementById("pwd-success"),o=document.getElementById("btn-save-pwd");r.style.display="none",l.style.display="none";let i=u.target,y=i.current_password.value,b=i.new_password.value,p=i.confirm_password.value;if(b!==p){r.textContent="\u274C Konfirmasi password tidak cocok.",r.style.display="block";return}if(b.length<6){r.textContent="\u274C Password baru minimal 6 karakter.",r.style.display="block";return}o.disabled=!0,o.textContent="\u23F3 Menyimpan...";let d=await f("/api/auth/change-password",{method:"POST",body:JSON.stringify({current_password:y,new_password:b})});o.disabled=!1,o.innerHTML='<svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24" style="margin-right:5px"><path d="M19 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11l5 5v11a2 2 0 0 1-2 2z"/><polyline points="17 21 17 13 7 13 7 21"/><polyline points="7 3 7 8 15 8"/></svg>Simpan Password',d.ok?(l.textContent="\u2705 Password berhasil diubah.",l.style.display="block",i.reset(),q("Password berhasil diubah.")):(r.textContent=d.data?.error||"Gagal mengubah password.",r.style.display="block")}),document.getElementById("btn-logout")?.addEventListener("click",()=>{confirm("Keluar dari semua sesi? Anda harus login ulang.")&&(localStorage.clear(),window.location.reload())})}I();var Le={Validasi:{module:"validation",label:"Master Referensi"},SOP:{module:"sop",label:"SOP"},"Master Karyawan":{module:"employees",label:"Karyawan"},"Data Kontrak":{module:"contracts",label:"Kontrak"},Permasalahan:{module:"issues",label:"Permasalahan"},"One on One":{module:"one_on_one",label:"One on One"},"Time Line":{module:"schedule",label:"Jadwal Kegiatan"},"Report Inspeksi Hygiene 2026":{module:"inspection",label:"Laporan Inspeksi"},"Report GC-DC 2026":{module:"cleaning",label:"Laporan GC/DC"},"Report Fogging 2026":{module:"fogging",label:"Laporan Fogging"},"Rekap Laporan Basecamp":{module:"basecamp",label:"Rekap Basecamp"},"Jadwal Reliefer":{module:"relievers",label:"Reliefer"},Training:{module:"training",label:"Training"},"Master Checklist":{module:"checklist",label:"Checklist"},"Master Form":{module:"forms",label:"Master Form"},"Permintaan Chemical":{module:"supply",label:"Inventory Chemical"}};function K(a){if(a==null||a==="")return null;if(a instanceof Date)return isNaN(a.getTime())?null:a.toISOString().slice(0,10);let e=String(a).trim();if(e===""||e==="0")return null;if(/^\d{4}-\d{2}-\d{2}/.test(e))return e.slice(0,10);if(/^\d{4,5}$/.test(e)){let t=Number(e);if(t>2e4&&t<99999){let c=new Date(Date.UTC(1899,11,30)+t*864e5);return isNaN(c.getTime())?null:c.toISOString().slice(0,10)}}let s=e.split(/[\/\-\.]/);if(s.length===3){let[t,c,g]=s.map(o=>o.trim()),u=Number(t),r=Number(c),l=Number(g);if(t.length===4&&u>1900)return`${t}-${c.padStart(2,"0")}-${g.padStart(2,"0")}`;if(g.length===4&&l>1900)return u>12?`${g}-${c.padStart(2,"0")}-${t.padStart(2,"0")}`:r>12?`${g}-${t.padStart(2,"0")}-${c.padStart(2,"0")}`:`${g}-${c.padStart(2,"0")}-${t.padStart(2,"0")}`;if(g.length===2&&!isNaN(l)){let o=l>=50?`19${g}`:`20${g}`;return u>12?`${o}-${c.padStart(2,"0")}-${t.padStart(2,"0")}`:`${o}-${c.padStart(2,"0")}-${t.padStart(2,"0")}`}}let n=new Date(e);return isNaN(n.getTime())?null:n.toISOString().slice(0,10)}function Ba(a){return Object.values(a).every(e=>e==null||String(e).trim()==="")}var gt={validation:{required:[],map:a=>({cabang:a.CABANG,pic:a.PIC,kegiatan:a.KEGIATAN,quartal:a.QUARTAL,masa_pkwt:a["MASA PKWT"],pic_pelapor:a["PIC PELAPOR"],kontrak:a.KONTRAK})},sop:{required:[{key:"Nama SOP",label:"Nama SOP"}],map:a=>({name:a["Nama SOP"],category:a.Kategori||"Umum",document_link:a["Link Document"],version:"1.0",effective_date:null,notes:""})},employees:{required:[{key:"Nama Lengkap",label:"Nama Lengkap"}],map:a=>({full_name:a["Nama Lengkap"],branch_name:a.Cabang,division:a["Div / Bagian"]||"FACILITY CARE",phone:a["No. Hp"],join_date:K(a["Tanggal Masuk"]),status:a.Status||"",notes:""})},contracts:{required:[{key:"Nama Lengkap",label:"Nama Lengkap"}],map:a=>({employee_name:a["Nama Lengkap"],branch_name:a.Cabang,division:a["Div / Bagian"]||"FACILITY CARE",start_date:K(a["Tanggal Mulai"]),end_date:K(a["Tanggal Selesai"]),contract_type:a["Tipe Kontrak"]||"",pkwt_number:a.PKWT||"",status:a.Status||"",notes:a.keterangan})},issues:{required:[{key:"Keluhan",label:"Keluhan"}],map:a=>({report_date:K(a["Tanggal Info"]),branch_name:a.Cabang,category:a.Kategori,source:a["Sumber Laporan"],complaint:a.Keluhan,employee_name:a["Nama FC"],fc_specialist:a["FC Spesialis"],solution:a.Solusi,status:a.Status||"",completion_date:K(a["Tanggal Selesai"])})},one_on_one:{required:[],map:a=>({meeting_date:K(a.Tanggal),branch_name:a.Cabang,employee_name:a["Nama Karyawan"],pic:a.Pic,problem:a.Masalah,solution:a.Solusi,status:a.Status||"",completion_date:K(a["Tanggal Selesai"]),document_link:a["Link Document"]})},schedule:{required:[{key:"Kegiatan",label:"Kegiatan"}],map:a=>({branch_name:a.Cabang,activity_type:a.Kegiatan,period:a.Periode,pic:a.Pic||a.PIC,opening_date:K(a["Tanggal Opening"]||a["Tgl Opening"]),target_date:K(a["Tanggal Target"]||a["Tgl Target"]),completion_date:K(a["Tanggal Selesai"]||a["Tgl Selesai"]),status:a.Status||"",notes:a.Keterangan||a.Catatan})},inspection:{required:[],map:a=>({inspection_date:K(a.Tanggal),branch_name:a.Cabang,period:a.Periode,status:a.Status||"",fc_score:a["Point FC SP"]!==void 0&&a["Point FC SP"]!==null?parseFloat(String(a["Point FC SP"]).replace(",",".")):null,spv_score:a["Point SPV"]!==void 0&&a["Point SPV"]!==null?parseFloat(String(a["Point SPV"]).replace(",",".")):null,document_link:a.Link,notes:""})},cleaning:{required:[],map:a=>({activity_date:K(a.Tanggal),branch_name:a.Cabang,activity_type:a["Jenis Kegiatan"]||"General Cleaning",period:a.Periode,status:a.Status||"",document_link:a.Link,notes:""})},fogging:{required:[],map:a=>({activity_date:K(a.Tanggal),branch_name:a.Cabang,period:a.Periode,status:a.Status||"",document_link:a.Link,notes:""})},basecamp:{required:[{key:"Permasalahan",label:"Permasalahan"}],map:a=>({info_date:K(a["Tgl Info"]),branch_name:a.Cabang,problem:a.Permasalahan,pic:a.PIC,done_date:K(a["Tgl Done"]),status:a.Status||"",notes:a.Ket})},relievers:{required:[],map:a=>({branch_name:a.Cabang,original_fc_name:a["Nama Facility care"],period:a.Periode,reliever_name:a.Relifer,backup_date:K(a["Tanggal Back Up"]),completion_date:K(a["Tanggal Selesai"]),reason:a.Keterangan,shift:a.Shift,status:a.Status||""})},training:{required:[{key:"Materi",label:"Materi"}],map:a=>({training_date:K(a.Tanggal),batch:a.Batch,subject:a.Materi,participants:a.Peserta,branch_name:a.Cabang,trainer:a.Trainer,score:a.Nilai!==void 0&&a.Nilai!==null?parseFloat(String(a.Nilai).replace(",",".")):null,notes:""})},checklist:{required:[],map:a=>({name:a["Master Checklist"],category:"Umum",document_link:a["Link Document"],description:""})},forms:{required:[{key:"Master Form",label:"Master Form"}],map:a=>({name:a["Master Form"],category:"Umum",document_link:a["Link Document"],description:""})},supply:{required:[],map:a=>({submitted_at:K(a.Timestamp),submitter_name:a["Nama Lengkap"],branch_name:a["Kebutuhan Untuk Cabang"],tools_items:a["Alat - Alat / Barang"],tools_quantity:a["Jumlah Permintaan Alat / Barang"],chemical_items:a.Chemical,chemical_quantity:a["Jumlah Permintaan Chemical"],additional_notes:a["Tambahan  Alat / Chemical Jika Ada Permintaan Diluar List."],status:a.Status||""})}};function bt(a,e){let s=Le[a];if(!s)return{valid:[],errors:[],mapped:[],skipped:!0};let n=gt[s.module];if(!n)return{valid:[],errors:[],mapped:[],skipped:!0};let t=[],c=[],g=[];return e.filter(r=>!Ba(r)).forEach((r,l)=>{let o=e.indexOf(r)+2,i=[];n.required.forEach(({key:b,label:p})=>{let d=r[b];if(d==null||String(d).trim()===""){let m=Object.keys(r).filter(h=>h.trim()).join(", ");i.push({column:p,originalValue:d||"",reason:`Kolom "${p}" wajib diisi dan tidak ditemukan`,hint:`Kolom yang tersedia: ${m.slice(0,120)}`})}});let y=n.map(r);i.length>0?c.push({row:o,data:y,raw:r,errors:i}):(t.push(r),g.push(y))}),{valid:t,errors:c,mapped:g}}function La(a){let e=[];return a.SheetNames.forEach(s=>{let n=Le[s];if(!n)return;let t=a.Sheets[s],c=window.XLSX.utils.sheet_to_json(t,{defval:"",raw:!1,dateNF:"yyyy-mm-dd"}),g=bt(s,c),u=c.filter(r=>!Ba(r));e.push({sheetName:s,module:n.module,label:n.label,total:u.length,valid:g.mapped.length,errorCount:g.errors.length,errors:g.errors,mapped:g.mapped,skipped:!1})}),e}function Aa(){let a=window.XLSX,e=a.utils.book_new();Object.entries({Validasi:[{CABANG:"001. Pondok Bambu","NAMA KARYAWAN":"Budi Santoso",PIC:"Berlin",KEGIATAN:"General Cleaning",QUARTAL:"Q1","PIC PELAPOR":"Berlin",KONTRAK:"PKWT 1","MASA PKWT":"1 Tahun"}],SOP:[{"Nama SOP":"SOP Pembersihan Toilet",Kategori:"Cleaning","Link Document":"https://..."}],"Master Karyawan":[{"Nama Lengkap":"Budi Santoso",Cabang:"001. Pondok Bambu","Div / Bagian":"FACILITY CARE","No. Hp":"081234567890","Tanggal Masuk":"2024-01-15",Status:"Aktif"}],"Data Kontrak":[{"Nama Lengkap":"Budi Santoso",Cabang:"001. Pondok Bambu","Div / Bagian":"FACILITY CARE","Tanggal Mulai":"2024-01-01","Tanggal Selesai":"2024-12-31","Sisa Kontrak":"365",Status:"Aktif",keterangan:""}],Permasalahan:[{"Tanggal Info":"2024-03-01",Cabang:"001. Pondok Bambu",Kategori:"Cleaning","Sumber Laporan":"SPV",Keluhan:"Lantai kotor","Nama FC":"Budi","FC Spesialis":"Fajar",Solusi:"Teguran",Status:"Done","Tanggal Selesai":"2024-03-02",Day:"1"}],"One on One":[{Tanggal:"2024-03-05",Cabang:"001. Pondok Bambu","Nama Karyawan":"Budi Santoso",Pic:"Berlin",Masalah:"Keterlambatan",Solusi:"Coaching",Status:"Done","Tanggal Selesai":"2024-03-06",Day:"1","Link Document":""}]}).forEach(([n,t])=>{a.utils.book_append_sheet(e,a.utils.json_to_sheet(t),n)}),a.writeFile(e,"Template_Import_Data_Awal_FCMS.xlsx")}function Na(a){let e=window.XLSX,s=e.utils.book_new(),n=!1;return a.forEach(t=>{if(!t.errors||t.errors.length===0)return;n=!0;let c=t.errors.map(u=>({"No. Baris":u.row,"Kolom Gagal":(u.errors||[]).map(r=>r.column||r).join("; "),"Alasan Error":(u.errors||[]).map(r=>r.reason||r).join("; "),...Object.fromEntries(Object.entries(u.data||{}).map(([r,l])=>[r,l??""]))})),g=e.utils.json_to_sheet(c);e.utils.book_append_sheet(s,g,t.sheetName.replace(/[\\\/\[\]*?:]/g,"_").slice(0,31))}),n?(e.writeFile(s,`Log_Error_Import_${new Date().toISOString().slice(0,10)}.xlsx`),!0):!1}var ht=["validation","employees","contracts","relievers","schedule","issues","one_on_one","training","checklist","forms","sop","inspection","cleaning","fogging","basecamp","supply"];function Fa(a){a.innerHTML=`
    <div class="page-header">
      <div class="header-content">
        <h1 class="page-title"><span class="title-icon">\u{1F4E5}</span> Import Data Awal</h1>
        <p class="page-subtitle">Unggah file Excel untuk mengisi data aplikasi, atau sinkronkan langsung dari Google Sheets.</p>
      </div>
      <div class="page-actions" style="display:flex;gap:8px">
        <button id="btn-sync-google" class="btn btn-secondary">
          <span>\u{1F504} Tarik Data dari Google Sheets</span>
        </button>
        <button class="btn btn-warning" id="btn-backup-db">\u{1F4E6} Backup Database</button>
        <button class="btn btn-secondary" id="btn-download-template">\u2B07\uFE0F Download Template</button>
      </div>
    </div>

    <!-- STEP 1: Upload -->
    <div id="step-upload" class="import-step">
      <div class="card">
        <div class="card-body">
          <div class="import-info-box">
            <h3>\u{1F4CB} Petunjuk Import Data Awal</h3>
            <p>Upload file Excel (.xlsx) yang sudah diisi sesuai template. Sistem akan membaca seluruh sheet secara otomatis dan memvalidasi sebelum data disimpan.</p>
            <div class="import-sheet-list">
              ${Object.entries(Le).map(([d,{label:m}])=>`<span class="import-sheet-tag">\u{1F4C4} ${d} \u2192 ${m}</span>`).join("")}
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
          <div class="import-summary-stats" id="summary-stats"></div>
          <div id="summary-module-results"></div>
          <div class="import-action-bar" style="margin-top:24px">
            <button class="btn btn-secondary" id="btn-import-again">\u{1F504} Import Lagi</button>
            <button class="btn btn-primary" id="btn-go-to-dashboard">\u{1F4CA} Ke Dashboard</button>
          </div>
        </div>
      </div>
    </div>
  `;let e=null,s=null,n=0,t={upload:document.getElementById("step-upload"),validating:document.getElementById("step-validating"),preview:document.getElementById("step-preview"),importing:document.getElementById("step-importing"),summary:document.getElementById("step-summary")};function c(d){Object.entries(t).forEach(([m,h])=>{h.style.display=m===d?"":"none"})}document.getElementById("btn-backup-db")?.addEventListener("click",async()=>{let d=document.getElementById("btn-backup-db");d.disabled=!0,d.textContent="\u23F3 Memproses Backup...";try{let m=await f("/api/import/backup");if(m.ok){let h=new Blob([JSON.stringify(m.data,null,2)],{type:"application/json"}),v=URL.createObjectURL(h),k=document.createElement("a");k.href=v,k.download=`FCMS_Database_Backup_${new Date().toISOString().slice(0,10)}.json`,document.body.appendChild(k),k.click(),document.body.removeChild(k),URL.revokeObjectURL(v),q("Backup berhasil diunduh!")}else U("Gagal memproses backup: "+(m.data?.error||"Unknown error"))}catch(m){U("Gagal memproses backup: "+m.message)}finally{d.disabled=!1,d.textContent="\u{1F4E6} Backup Database"}});let g=document.getElementById("btn-sync-google");g&&g.addEventListener("click",async()=>{if(!confirm("Peringatan: Mensinkronkan data dengan Google Sheets akan memperbarui dan menambahkan data baru dari Google Sheets ke dalam FCMS. Data yang sudah Anda buat di FCMS TIDAK akan terhapus. Lanjutkan?"))return;let d=g.innerHTML;g.innerHTML='<span class="spinner"></span> Menyinkronkan...',g.disabled=!0;try{let m=await f("/api/sync/google-sheets",{method:"POST"});m.ok?alert("Sinkronisasi Berhasil: "+(m.data?.message||"Data Karyawan & PIC telah diperbarui.")):alert("Gagal Sinkronisasi: "+(m.data?.error||"Unknown error"))}catch{alert("Terjadi kesalahan koneksi.")}finally{g.innerHTML=d,g.disabled=!1}}),document.getElementById("btn-download-template").addEventListener("click",()=>{Aa(),q("Template Excel berhasil didownload!")});let u=document.getElementById("file-input"),r=document.getElementById("upload-zone");document.getElementById("btn-browse").addEventListener("click",d=>{d.stopPropagation(),u.click()}),u.addEventListener("change",d=>{d.target.files[0]&&l(d.target.files[0])}),r.addEventListener("dragover",d=>{d.preventDefault(),r.classList.add("drag-over")}),r.addEventListener("dragleave",()=>r.classList.remove("drag-over")),r.addEventListener("drop",d=>{d.preventDefault(),r.classList.remove("drag-over");let m=d.dataTransfer.files[0];m&&m.name.match(/\.xlsx?$/i)?l(m):U("Hanya file .xlsx atau .xls yang didukung.")}),document.getElementById("btn-clear-file").addEventListener("click",()=>{e=null,u.value="",document.getElementById("file-info").style.display="none",r.style.display="",c("upload")});async function l(d){e=d,document.getElementById("file-name-display").textContent=`\u{1F4C4} ${d.name} (${(d.size/1024).toFixed(1)} KB)`,document.getElementById("file-info").style.display="flex",r.style.display="none",await o(d)}async function o(d){c("validating");let m=document.getElementById("validation-status"),h=document.getElementById("validation-bar");try{if(!window.XLSX)throw new Error("Library SheetJS belum termuat. Refresh halaman dan coba lagi.");m.textContent="Membaca file Excel...",h.style.width="20%",await ke(200);let v=await d.arrayBuffer(),k=window.XLSX.read(v,{type:"array",cellDates:!0});m.textContent=`Memvalidasi ${k.SheetNames.length} sheet...`,h.style.width="50%",await ke(100),s=La(k),h.style.width="100%",m.textContent="Validasi selesai!",await ke(300),i()}catch(v){c("upload"),U("Gagal memproses file: "+v.message),document.getElementById("file-info").style.display="flex",r.style.display="none"}}function i(){c("preview");let d=s.filter(_=>!_.skipped).length,m=s.reduce((_,D)=>_+D.total,0),h=s.reduce((_,D)=>_+D.valid,0),v=s.reduce((_,D)=>_+D.errorCount,0),k=m>0?Math.round(h/m*100):0;document.getElementById("preview-summary-badges").innerHTML=`
      <span class="badge badge-info">${d} sheet</span>
      <span class="badge badge-secondary">${m} baris</span>
      <span class="badge badge-success">${h} valid (${k}%)</span>
      ${v>0?`<span class="badge badge-danger">${v} error</span>`:""}
    `;let S=document.getElementById("preview-table-container");S.innerHTML=`
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
          ${s.map((_,D)=>`
            <tr class="${_.errorCount>0?"row-error":_.skipped?"row-skipped":"row-ok"}">
              <td><strong>${_.sheetName}</strong></td>
              <td>${_.label}</td>
              <td style="text-align:center">${_.total}</td>
              <td style="text-align:center"><span class="badge badge-success">${_.valid}</span></td>
              <td style="text-align:center">${_.errorCount>0?`<span class="badge badge-danger">${_.errorCount}</span>`:'<span class="text-muted">\u2013</span>'}</td>
              <td style="text-align:center">
                ${_.skipped?'<span class="badge badge-neutral">Dilewati</span>':_.errorCount>0&&_.valid===0?'<span class="badge badge-danger">\u274C 0 Valid</span>':_.errorCount>0?'<span class="badge badge-warning">\u26A0\uFE0F Sebagian</span>':_.valid===0?'<span class="badge badge-neutral">Kosong</span>':'<span class="badge badge-success">\u2705 Siap</span>'}
              </td>
              <td style="text-align:center">
                ${_.errorCount>0?`<button class="btn btn-ghost btn-sm btn-detail-error" data-idx="${D}">\u{1F50D} ${_.errorCount} Error</button>`:'<span class="text-muted">\u2013</span>'}
              </td>
            </tr>
          `).join("")}
        </tbody>
      </table>
    `,S.querySelectorAll(".btn-detail-error").forEach(_=>{_.addEventListener("click",()=>{let D=s[Number(_.dataset.idx)];y(D)})});let w=document.getElementById("error-detail-section"),T=document.getElementById("error-detail-container");T.innerHTML="",w.style.display="none";let L=document.getElementById("btn-start-import");h===0?(L.disabled=!0,L.innerHTML="\u26A0\uFE0F Tidak Ada Data Valid"):(L.disabled=!1,v>0?(L.innerHTML=`\u{1F680} Import ${h} Data Valid (${v} dilewati)`,L.title="Baris error akan dilewati, baris valid tetap diimport"):L.innerHTML=`\u{1F680} Mulai Import ${h} Data`)}function y(d){let m=document.getElementById("error-detail-section"),h=document.getElementById("error-detail-container");m.style.display="";let v=d.errors.slice(0,100).map(k=>(Array.isArray(k.errors)?k.errors:[]).map(w=>{let T=typeof w=="object";return`
          <tr>
            <td style="text-align:center"><span class="badge badge-danger">Baris ${k.row}</span></td>
            <td><strong>${T?w.column:"\u2014"}</strong></td>
            <td><code style="font-size:.78rem;color:var(--text-secondary)">${T&&w.originalValue!==void 0?w.originalValue||"(kosong)":"\u2014"}</code></td>
            <td class="error-msg">${T?w.reason:w}</td>
            <td style="font-size:.78rem;color:var(--success)">
              ${T&&w.aliases?`Gunakan salah satu nama kolom:<br><em>${w.aliases}</em>`:T&&w.hint?w.hint:""}
            </td>
          </tr>
        `}).join("")).join("");h.innerHTML=`
      <div class="error-sheet-block">
        <div class="error-sheet-title">
          \u{1F4C4} ${d.sheetName} \u2014 ${d.errorCount} baris error dari ${d.total} total
          ${d.errors.length>100?'<span style="font-weight:400">(menampilkan 100 pertama)</span>':""}
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
            <tbody>${v||'<tr><td colspan="5" class="text-muted" style="text-align:center">Tidak ada detail error</td></tr>'}</tbody>
          </table>
        </div>
        ${d.errors.length>100?`
          <div style="padding:10px 20px;font-size:.8rem;color:var(--text-muted)">
            Hanya menampilkan 100 error pertama. Download Log Error untuk melihat semua.
          </div>`:""}
      </div>
    `,m.scrollIntoView({behavior:"smooth",block:"start"})}document.getElementById("btn-back-to-upload").addEventListener("click",()=>{c("upload"),document.getElementById("file-info").style.display="none",r.style.display="",e=null,u.value=""}),document.getElementById("btn-download-log").addEventListener("click",()=>{if(!s)return;Na(s)?q("Log error berhasil didownload."):q("Tidak ada error untuk didownload.")}),document.getElementById("btn-start-import").addEventListener("click",()=>{let d=document.querySelector('input[name="dup-strategy"]:checked')?.value||"skip";b(d)});async function b(d){c("importing"),n=Date.now();let m=[];ht.forEach(w=>{let T=s?.find(L=>L.module===w&&L.mapped?.length>0);T&&m.push(T)});let h=document.getElementById("import-steps-list");h.innerHTML=m.map(w=>`
      <div class="import-step-item" id="step-item-${w.module}">
        <span class="step-item-icon" id="step-icon-${w.module}">\u23F8\uFE0F</span>
        <span class="step-item-label">${w.label} <span class="step-item-count">(${w.mapped.length} data)</span></span>
        <span class="step-item-status" id="step-status-${w.module}"></span>
      </div>
    `).join("");let v=document.getElementById("import-bar"),k=document.getElementById("import-current-status"),S={totalSheets:m.length,totalRows:m.reduce((w,T)=>w+T.mapped.length,0),inserted:0,skipped:0,failed:0,moduleResults:[]};for(let w=0;w<m.length;w++){let T=m[w],L=document.getElementById(`step-icon-${T.module}`),_=document.getElementById(`step-status-${T.module}`);L.textContent="\u{1F504}",_.textContent="Mengimport...",k.textContent=`Mengimport ${T.label}...`,v.style.width=`${Math.round(w/m.length*100)}%`;try{let D=await f(`/api/import/${T.module}`,{method:"POST",body:JSON.stringify({rows:T.mapped,onDuplicate:d})});if(D.ok){let J=D.data;S.inserted+=J.inserted||0,S.skipped+=J.skipped||0,S.moduleResults.push({label:T.label,inserted:J.inserted||0,skipped:J.skipped||0,status:"ok"}),L.textContent="\u2705",_.innerHTML=`<span class="badge badge-success">${J.inserted||0} berhasil</span>${J.skipped>0?` <span class="badge badge-neutral">${J.skipped} skip</span>`:""}`}else S.failed++,S.moduleResults.push({label:T.label,inserted:0,skipped:0,status:"error",error:D.data?.error}),L.textContent="\u274C",_.innerHTML='<span class="badge badge-danger">Gagal</span>'}catch(D){S.failed++,S.moduleResults.push({label:T.label,inserted:0,skipped:0,status:"error",error:D.message}),L.textContent="\u274C",_.innerHTML='<span class="badge badge-danger">Gagal</span>'}await ke(150)}v.style.width="100%",k.textContent="Selesai!",await ke(400),p(S)}function p(d){c("summary");let m=((Date.now()-n)/1e3).toFixed(1),h=d.failed===0;document.getElementById("summary-status-icon").innerHTML=`
      <div class="summary-icon">${h?"\u{1F389}":"\u26A0\uFE0F"}</div>
      <h2 class="summary-title">${h?"Import Berhasil!":"Import Selesai dengan Beberapa Error"}</h2>
    `,document.getElementById("summary-stats").innerHTML=`
      <div class="summary-stat-card">
        <div class="stat-value">${d.totalSheets}</div>
        <div class="stat-label">Total Sheet</div>
      </div>
      <div class="summary-stat-card">
        <div class="stat-value">${d.totalRows}</div>
        <div class="stat-label">Total Data</div>
      </div>
      <div class="summary-stat-card success">
        <div class="stat-value">${d.inserted}</div>
        <div class="stat-label">Berhasil Diimport</div>
      </div>
      <div class="summary-stat-card neutral">
        <div class="stat-value">${d.skipped}</div>
        <div class="stat-label">Dilewati (Duplikat)</div>
      </div>
      ${d.failed>0?`<div class="summary-stat-card danger"><div class="stat-value">${d.failed}</div><div class="stat-label">Modul Gagal</div></div>`:""}
      <div class="summary-stat-card info">
        <div class="stat-value">${m}s</div>
        <div class="stat-label">Durasi Proses</div>
      </div>
    `,document.getElementById("summary-module-results").innerHTML=`
      <table class="data-table" style="margin-top:16px">
        <thead>
          <tr><th>Modul</th><th style="text-align:center">Berhasil</th><th style="text-align:center">Dilewati</th><th style="text-align:center">Status</th></tr>
        </thead>
        <tbody>
          ${d.moduleResults.map(v=>`
            <tr>
              <td>${v.label}</td>
              <td style="text-align:center"><span class="badge badge-success">${v.inserted}</span></td>
              <td style="text-align:center"><span class="badge badge-neutral">${v.skipped}</span></td>
              <td style="text-align:center">
                ${v.status==="ok"?'<span class="badge badge-success">\u2705 Sukses</span>':`<span class="badge badge-danger" title="${v.error||""}">\u274C Gagal</span>`}
              </td>
            </tr>
          `).join("")}
        </tbody>
      </table>
    `}document.getElementById("btn-import-again").addEventListener("click",()=>{e=null,s=null,u.value="",document.getElementById("file-info").style.display="none",r.style.display="",c("upload")}),document.getElementById("btn-go-to-dashboard").addEventListener("click",()=>{window.location.hash="/dashboard"})}function ke(a){return new Promise(e=>setTimeout(e,a))}I();var Ze=[],yt=[];async function Oa(a){let[e,s]=await Promise.all([f("/api/branches?all=1"),f(`/api/sp_data${window.location.search?window.location.search+"&":"?"}limit=10000`)]);Ze=(e.data?.data||[]).map(n=>({value:n.id,label:n.full_name})),yt=(s.data?.data||[]).map(n=>({value:n.full_name,label:n.full_name})),E({container:a,title:"Data SP (Surat Peringatan)",icon:"\u2709\uFE0F",apiPath:"/api/sp",itemLabel:"SP",bulkDelete:!0,columns:[{key:"employee_name",label:"Nama Karyawan"},{key:"division",label:"Divisi",render:n=>n?`<span class="badge badge-info">${n}</span>`:"-"},{key:"branch_name",label:"Cabang"},{key:"tanggal",label:"Tanggal Sp",render:n=>n?new Date(n).toLocaleDateString("id-ID",{year:"numeric",month:"short",day:"numeric"}):"-"},{key:"akhir_sp",label:"Akhir Sp",render:n=>n?new Date(n).toLocaleDateString("id-ID",{year:"numeric",month:"short",day:"numeric"}):"-"},{key:"sp_type",label:"Jenis Sp",render:n=>`<span class="badge badge-warning">${n||"-"}</span>`},{key:"document_link",label:"Link Document / Foto",render:n=>n?`<a href="${n}" target="_blank" class="text-primary hover-underline">Lihat</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari nama karyawan..."},{type:"select",name:"branch_id",label:"Cabang",options:Ze}],exportOptions:{moduleName:"sp_data",onExport:async n=>{let t=new URLSearchParams(n||{}).toString(),c=await f(`/api/sp?limit=10000&${t}`);if(c.ok){let g=c.data.data.map(r=>({"Nama Karyawan":r.employee_name||"",Divisi:r.division||"",Cabang:r.branch_name||"","Tanggal Sp":r.tanggal||"","Akhir Sp":r.akhir_sp||"","Jenis Sp":r.sp_type||"","Link Document / Foto":r.document_link||""})),{downloadExcel:u}=await Promise.resolve().then(()=>(B(),G));u(g,`Data_SP_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:async()=>{let n=[{"Nama Karyawan":"Budi Santoso",Divisi:"FACILITY CARE",Cabang:"001. Pondok Bambu","Tanggal Sp":"2026-01-08","Akhir Sp":"2026-07-08","Jenis Sp":"SP 1","Link Document / Foto":"https://link.doc"}],{downloadExcel:t}=await Promise.resolve().then(()=>(B(),G));t(n,"Template_Import_SP")},onImport:async n=>{let c=(await f("/api/branches?all=1")).data?.data||[],g=o=>{if(!o)return null;let i=String(o||"").toLowerCase(),y=c.find(b=>String(b.full_name||"").toLowerCase()===i||String(b.code||"").toLowerCase()===i||String(b.name||"").toLowerCase()===i);return y?y.id:null},u=o=>{if(!o)return"";if(o instanceof Date&&!isNaN(o.getTime()))return o.toISOString().slice(0,10);let i=String(o).trim();if(/^\d{4,5}$/.test(i)){let b=Number(i);if(b>2e4&&b<99999){let p=new Date(Date.UTC(1899,11,30)+b*864e5);return isNaN(p.getTime())?"":p.toISOString().slice(0,10)}}if(/^\d{4}-\d{2}-\d{2}/.test(i))return i.slice(0,10);let y=i.split(/[\/\-\.]/);if(y.length===3){let[b,p,d]=y.map(m=>m.trim());if(b.length===4&&p.length<=2&&d.length<=2)return`${b}-${p.padStart(2,"0")}-${d.padStart(2,"0")}`;if(d.length===4&&p.length<=2&&b.length<=2)return`${d}-${p.padStart(2,"0")}-${b.padStart(2,"0")}`}return i},r=n.map(o=>({employee_name:String(o["Nama Karyawan"]||"").trim(),division:String(o.Divisi||"").trim(),branch_id:g(String(o.Cabang||"").trim()),tanggal:u(o["Tanggal Sp"]),akhir_sp:u(o["Akhir Sp"]),sp_type:String(o["Jenis Sp"]||"").trim(),document_link:String(o["Link Document / Foto"]||"").trim()})).filter(o=>o.employee_name&&o.branch_id),l=await f("/api/sp/import",{method:"POST",body:JSON.stringify(r)});if(!l.ok)throw new Error(l.data?.error||"Import gagal")}},formFields:[{type:"text",name:"employee_name",label:"Nama Karyawan",required:!0},{type:"select",name:"division",label:"Divisi",options:["FACILITY CARE","SECURITY"],required:!0},{type:"select",name:"branch_id",label:"Cabang",required:!0,options:Ze},{type:"date",name:"tanggal",label:"Tanggal Sp",required:!0},{type:"date",name:"akhir_sp",label:"Akhir Sp",required:!0},{type:"select",name:"sp_type",label:"Jenis Sp",required:!0,options:["SP 1","SP 2","SP 3","Teguran Lisan"]},{type:"url",name:"document_link",label:"Link Document / Foto"}]})}I();var Se=[],ft=[];async function Ma(a){let[e,s]=await Promise.all([f("/api/branches?all=1"),f(`/api/mutasi_data${window.location.search?window.location.search+"&":"?"}limit=10000`)]);Se=(e.data?.data||[]).map(n=>({value:n.id,label:n.full_name})),ft=(s.data?.data||[]).map(n=>({value:n.full_name,label:n.full_name})),E({container:a,title:"Data Mutasi",icon:"\u{1F504}",apiPath:"/api/mutasi",itemLabel:"Mutasi",bulkDelete:!0,columns:[{key:"tanggal",label:"Tanggal",render:n=>n?new Date(n).toLocaleDateString("id-ID",{year:"numeric",month:"short",day:"numeric"}):"-"},{key:"employee_name",label:"Nama Karyawan"},{key:"from_branch_name",label:"Cabang Asal"},{key:"to_branch_name",label:"Cabang Tujuan"},{key:"status",label:"Status",render:n=>`<span class="badge ${n==="Selesai"?"badge-success":"badge-warning"}">${n||"-"}</span>`},{key:"document_link",label:"Dokumen",render:n=>n?`<a href="${n}" target="_blank" class="text-primary hover-underline">Lihat</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari nama karyawan..."},{type:"select",name:"from_branch_id",label:"Cabang Asal",options:Se},{type:"select",name:"to_branch_id",label:"Cabang Tujuan",options:Se}],exportOptions:{moduleName:"mutasi_data",onExport:async n=>{let t=new URLSearchParams(n||{}).toString(),c=await f(`/api/mutasi?limit=10000&${t}`);if(c.ok){let g=c.data.data.map(r=>({Tanggal:r.tanggal||"","Nama Karyawan":r.employee_name||"","Cabang Asal":r.from_branch_name||"","Cabang Tujuan":r.to_branch_name||"",Status:r.status||"",Dokumen:r.document_link||""})),{downloadExcel:u}=await Promise.resolve().then(()=>(B(),G));u(g,`Data_Mutasi_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:async()=>{let n=[{Tanggal:"2026-01-08","Nama Karyawan":"Widya Astuti","Cabang Asal":"001. Pondok Bambu","Cabang Tujuan":"007. Bekasi",Status:"Selesai",Dokumen:"https://link.doc"}],{downloadExcel:t}=await Promise.resolve().then(()=>(B(),G));t(n,"Template_Import_Mutasi")},onImport:async n=>{let c=(await f("/api/branches?all=1")).data?.data||[],g=o=>{if(!o)return null;let i=String(o||"").toLowerCase(),y=c.find(b=>String(b.full_name||"").toLowerCase()===i||String(b.code||"").toLowerCase()===i||String(b.name||"").toLowerCase()===i);return y?y.id:null},u=o=>{if(!o)return"";if(o instanceof Date&&!isNaN(o.getTime()))return o.toISOString().slice(0,10);let i=String(o).trim();if(/^\d{4,5}$/.test(i)){let b=Number(i);if(b>2e4&&b<99999){let p=new Date(Date.UTC(1899,11,30)+b*864e5);return isNaN(p.getTime())?"":p.toISOString().slice(0,10)}}if(/^\d{4}-\d{2}-\d{2}/.test(i))return i.slice(0,10);let y=i.split(/[\/\-\.]/);if(y.length===3){let[b,p,d]=y.map(m=>m.trim());if(b.length===4&&p.length<=2&&d.length<=2)return`${b}-${p.padStart(2,"0")}-${d.padStart(2,"0")}`;if(d.length===4&&p.length<=2&&b.length<=2)return`${d}-${p.padStart(2,"0")}-${b.padStart(2,"0")}`}return i},r=n.map(o=>({tanggal:u(o.Tanggal),employee_name:String(o["Nama Karyawan"]||"").trim(),from_branch_id:g(String(o["Cabang Asal"]||"").trim()),to_branch_id:g(String(o["Cabang Tujuan"]||"").trim()),status:String(o.Status||"").trim(),document_link:String(o.Dokumen||"").trim()})).filter(o=>o.tanggal&&o.employee_name&&o.from_branch_id&&o.to_branch_id),l=await f("/api/mutasi/import",{method:"POST",body:JSON.stringify(r)});if(!l.ok)throw new Error(l.data?.error||"Import gagal")}},formFields:[{type:"date",name:"tanggal",label:"Tanggal",required:!0},{type:"text",name:"employee_name",label:"Nama Karyawan",required:!0},{type:"combobox",name:"from_branch_id",label:"Cabang Asal",required:!0,options:Se,createApi:{path:"/api/branches",field:"full_name"}},{type:"combobox",name:"to_branch_id",label:"Cabang Tujuan",required:!0,options:Se,createApi:{path:"/api/branches",field:"full_name"}},{type:"select",name:"status",label:"Status",required:!0,options:["Proses","Selesai"]},{type:"url",name:"document_link",label:"Link Dokumen (Opsional)"}]})}window.parseFlexibleDate=a=>{if(!a||a==="-")return"";if(a=String(a).trim(),/^\d{5}$/.test(a)){let e=Math.floor(Number(a)-25569);return new Date(e*86400*1e3).toISOString().split("T")[0]}if(a.match(/^\d{2}[\/\-]\d{2}[\/\-]\d{4}$/)){let e=a.split(/[\/\-]/);return`${e[2]}-${e[1]}-${e[0]}`}return a.split("T")[0]};window.formatDate=a=>{let e=window.parseFlexibleDate(a);if(!e)return"";let s=e.split("-");if(s.length===3&&s[0].length===4){let n=["Januari","Februari","Maret","April","Mei","Juni","Juli","Agustus","September","Oktober","November","Desember"],t=parseInt(s[2],10),c=n[parseInt(s[1],10)-1];return`${t} ${c} ${s[0]}`}return e};function N(a){return async e=>{if(!de()){ie("/login");return}return a(e)}}var we=null;function vt(){we&&clearInterval(we);let a=()=>{let e=new Date,s=e.toLocaleTimeString("id-ID",{hour:"2-digit",minute:"2-digit",second:"2-digit"}),n=e.toLocaleDateString("id-ID",{weekday:"long",day:"numeric",month:"long",year:"numeric"}),t=document.getElementById("header-clock-time"),c=document.getElementById("header-clock-date");t&&(t.textContent=s),c&&(c.textContent=n)};a(),we=setInterval(a,1e3)}async function kt(){try{let a=await f("/api/dashboard/kpi");if(!a.ok)return;let e=a.data?.data||a.data||{},s=(n,t)=>{let c=document.getElementById(n);c&&(c.textContent=t>0?t:"",c.style.display=t>0?"inline-flex":"none")};s("badge-issues",e.issues?.current||0),s("badge-contracts",e.expiring30?.current||0),s("badge-oo1",e.one_on_one?.current||0),s("badge-schedule",e.schedule?.current||0),s("badge-supply",e.supply?.current||0)}catch{}}var oe=[];async function St(){try{let a=await f("/api/dashboard/notifications");if(!a.ok)return;oe=a.data?.data||a.data||[];let e=document.getElementById("notif-dot");e&&(e.style.display=oe.length>0?"block":"none",e.textContent=oe.length)}catch{}}function wt(){if(!oe.length){Y({title:"Notifikasi",content:'<div class="empty-state"><p>Tidak ada notifikasi baru.</p></div>',confirmText:"Tutup",onConfirm:(e,s)=>s()});return}let a=`
    <div class="notif-list" style="max-height: 400px; overflow-y: auto;">
      ${oe.map(e=>`
        <div class="notif-item notif-severity-${e.severity||"info"}" style="padding: 12px; border-bottom: 1px solid var(--border); border-left: 4px solid var(--${e.severity==="danger"?"danger":e.severity==="warning"?"warning":"primary"}); margin-bottom: 8px; border-radius: 4px; background: #fff;">
          <div style="font-weight: 600; font-size: 0.9rem; color: var(--text-1);">${e.title}</div>
          <div style="display: flex; justify-content: space-between; align-items: center; margin-top: 4px; font-size: 0.75rem; color: var(--text-3);">
            <span>\u{1F4C5} ${e.date}</span>
            <span class="badge badge-${e.severity==="danger"?"danger":e.severity==="warning"?"warning":"info"}">${e.type.toUpperCase()}</span>
          </div>
        </div>
      `).join("")}
    </div>
  `;Y({title:`Notifikasi (${oe.length})`,content:a,confirmText:"Tutup",onConfirm:(e,s)=>s()})}function Ra(){let a=ae(),e=(a?.full_name||"U")[0].toUpperCase();document.getElementById("app").innerHTML=`
    <div class="app-layout">
      <!-- Sidebar dark -->
      <aside class="sidebar" id="sidebar">
        <div class="sidebar-header">
          <div class="sidebar-logo">
            <span class="logo-icon-wrap">\u{1F3E5}</span>
            <span class="logo-text">FC<strong>MS</strong></span>
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
            <a href="#/relievers" class="nav-item" data-route="/relievers">
              <span class="nav-icon">
                <svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><polyline points="17 1 21 5 17 9"/><path d="M3 11V9a4 4 0 014-4h14"/><polyline points="7 23 3 19 7 15"/><path d="M21 13v2a4 4 0 01-4 4H3"/></svg>
              </span>
              <span class="nav-label">Jadwal Reliefer</span>
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
            <a href="#/issues" class="nav-item" data-route="/issues">
              <span class="nav-icon">
                <svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><circle cx="12" cy="12" r="10"/><line x1="12" y1="8" x2="12" y2="12"/><line x1="12" y1="16" x2="12.01" y2="16"/></svg>
              </span>
              <span class="nav-label">Permasalahan</span>
              <span class="nav-badge badge-danger" id="badge-issues"></span>
            </a>
            <a href="#/one-on-one" class="nav-item" data-route="/one-on-one">
              <span class="nav-icon">
                <svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M21 15a2 2 0 01-2 2H7l-4 4V5a2 2 0 012-2h14a2 2 0 012 2z"/></svg>
              </span>
              <span class="nav-label">One on One</span>
              <span class="nav-badge badge-warning" id="badge-oo1"></span>
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
            <a href="#/reports/supply" class="nav-item" data-route="/reports/supply">
              <span class="nav-icon">
                <svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M6 2L3 6v14a2 2 0 002 2h14a2 2 0 002-2V6l-3-4z"/><line x1="3" y1="6" x2="21" y2="6"/><path d="M16 10a4 4 0 01-8 0"/></svg>
              </span>
              <span class="nav-label">Permintaan Chemical</span>
              <span class="nav-badge" id="badge-supply"></span>
            </a>
          </div>

          <!-- Referensi -->
          <div class="nav-section">
            <span class="nav-section-label">REFERENSI</span>
            <a href="#/sop" class="nav-item" data-route="/sop">
              <span class="nav-icon">
                <svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M4 19.5A2.5 2.5 0 016.5 17H20"/><path d="M6.5 2H20v20H6.5A2.5 2.5 0 014 19.5v-15A2.5 2.5 0 016.5 2z"/></svg>
              </span>
              <span class="nav-label">SOP</span>
            </a>
            <a href="#/checklist" class="nav-item" data-route="/checklist">
              <span class="nav-icon">
                <svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><polyline points="9 11 12 14 22 4"/><path d="M21 12v7a2 2 0 01-2 2H5a2 2 0 01-2-2V5a2 2 0 012-2h11"/></svg>
              </span>
              <span class="nav-label">Master Checklist</span>
            </a>
            <a href="#/forms" class="nav-item" data-route="/forms">
              <span class="nav-icon">
                <svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M14 2H6a2 2 0 00-2 2v16a2 2 0 002 2h12a2 2 0 002-2V8z"/><polyline points="14 2 14 8 20 8"/></svg>
              </span>
              <span class="nav-label">Master Form</span>
            </a>
          </div>

          <!-- Admin -->
          ${a&&(a.role==="superadmin"||a.role==="admin")?`
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
            <a href="#/settings/import" class="nav-item" data-route="/settings/import">
              <span class="nav-icon">
                <svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M21 15v4a2 2 0 01-2 2H5a2 2 0 01-2-2v-4"/><polyline points="17 8 12 3 7 8"/><line x1="12" y1="3" x2="12" y2="15"/></svg>
              </span>
              <span class="nav-label">Import Data Awal</span>
            </a>
          </div>`:""}
        </nav>

        <div class="sidebar-footer">
          <div class="sidebar-user">
            <div class="sidebar-avatar">BA</div>
            <div class="sidebar-user-info">
              <div class="sidebar-user-name">Berlin Ariansyah</div>
              <div class="sidebar-user-role">Administrator</div>
            </div>
          </div>
          <button class="sidebar-logout" id="logout-btn">
            <svg width="16" height="16" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M9 21H5a2 2 0 01-2-2V5a2 2 0 012-2h4"/><polyline points="16 17 21 12 16 7"/><line x1="21" y1="12" x2="9" y2="12"/></svg>
            Keluar
          </button>
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
            <div class="topbar-welcome">
              <div class="topbar-greeting">
                <span class="topbar-greeting-time">${(()=>{let l=new Date().getHours();return l>=4&&l<11?"Selamat Pagi":l>=11&&l<15?"Selamat Siang":l>=15&&l<18?"Selamat Sore":"Selamat Malam"})()}, </span><span class="topbar-greeting-name">Berlin Ariansyah</span> \u{1F44B}
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
            <button class="topbar-icon-btn" id="btn-fullscreen" title="Fullscreen" aria-label="Fullscreen">
              <svg width="17" height="17" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><polyline points="15 3 21 3 21 9"/><polyline points="9 21 3 21 3 15"/><line x1="21" y1="3" x2="14" y2="10"/><line x1="3" y1="21" x2="10" y2="14"/></svg>
            </button>
            <button class="topbar-icon-btn notif-btn" id="btn-notif" title="Notifikasi" aria-label="Notifikasi">
              <svg width="17" height="17" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M18 8A6 6 0 006 8c0 7-3 9-3 9h18s-3-2-3-9"/><path d="M13.73 21a2 2 0 01-3.46 0"/></svg>
              <span class="notif-dot" id="notif-dot" style="display:none"></span>
            </button>
            <a href="#/profile" class="topbar-user-btn" title="Profil">
              <img src="https://ui-avatars.com/api/?name=Berlin+Ariansyah&background=2563EB&color=fff&bold=true" class="topbar-avatar" alt="Avatar" />
              <div class="topbar-user-text">
                <span class="topbar-user-name">Berlin Ariansyah</span>
                <span class="topbar-user-role-mini">Administrator</span>
              </div>
              <svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24" style="margin-left:4px;color:var(--gray-400)"><polyline points="6 9 12 15 18 9"/></svg>
            </a>
          </div>
        </header>

        <main id="main-content" class="main-content"></main>
      </div>
    </div>
  `;let s=document.getElementById("sidebar"),n=document.getElementById("sidebar-overlay"),t=document.getElementById("topbar-menu-btn"),c=document.getElementById("sidebar-close"),g=()=>{s.classList.add("open"),n.classList.add("show")},u=()=>{s.classList.remove("open"),n.classList.remove("show")};t?.addEventListener("click",g),c?.addEventListener("click",u),n?.addEventListener("click",u),document.querySelectorAll(".nav-item").forEach(l=>l.addEventListener("click",u));function r(){let l=window.location.hash.replace("#","")||"/dashboard";document.querySelectorAll(".nav-item").forEach(y=>{let b=y.dataset.route;y.classList.toggle("active",l===b||b!=="/dashboard"&&l.startsWith(b))});let o=document.getElementById("topbar-title"),i=document.querySelector(".nav-item.active .nav-label");o&&i&&(o.textContent=i.textContent)}window.addEventListener("hashchange",r),r(),vt(),document.getElementById("btn-fullscreen")?.addEventListener("click",()=>{document.fullscreenElement?document.exitFullscreen?.():document.documentElement.requestFullscreen?.()}),document.getElementById("logout-btn")?.addEventListener("click",async()=>{await f("/api/auth/logout",{method:"POST"}),pe(),we&&clearInterval(we),ie("/login")}),kt(),St(),document.getElementById("btn-notif")?.addEventListener("click",l=>{l.preventDefault(),wt()})}async function _t(){A("/login",({main:e})=>ca(e)),A("/dashboard",N(({main:e})=>sa(e))),A("/calendar",N(({main:e})=>Da(e))),A("/employees",N(({main:e})=>ha(e))),A("/contracts",N(({main:e})=>De(e))),A("/sp",N(({main:e})=>Oa(e))),A("/mutasi",N(({main:e})=>Ma(e))),A("/timeline",N(({main:e})=>ya(e))),A("/issues",N(({main:e})=>fa(e))),A("/one-on-one",N(({main:e})=>va(e))),A("/training",N(({main:e})=>ka(e))),A("/relievers",N(({main:e})=>Sa(e))),A("/reports/inspection",N(({main:e})=>wa(e))),A("/reports/cleaning",N(({main:e})=>_a(e))),A("/reports/fogging",N(({main:e})=>xa(e))),A("/reports/basecamp",N(({main:e})=>Ca(e))),A("/reports/supply",N(({main:e})=>Xe(e,"supply"))),A("/sop",N(({main:e})=>Ta(e))),A("/checklist",N(({main:e})=>$a(e))),A("/forms",N(({main:e})=>Xe(e))),A("/users",N(({main:e})=>Ea(e))),A("/branches",N(({main:e})=>Pa(e))),A("/profile",N(({main:e})=>Ia(e))),A("/settings/import",N(({main:e})=>Fa(e)));let a=de();if(!a&&window.location.hash!=="#/login"&&ie("/login"),a){let e=await f("/api/auth/me");e.ok?(me(e.data.data),Ra()):(pe(),ie("/login"))}window.addEventListener("fm:login",()=>{Ra(),ie("/dashboard")}),ta()}_t();
