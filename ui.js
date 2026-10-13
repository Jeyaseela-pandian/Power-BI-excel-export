"use strict";

/*
 * Interface: signed-in shell, toasts, export options drawer, branding settings,
 * export progress and the diagnostics panel. Loaded last; starts the app.
 * Uses globals from app.js and export.js.
 */

/* ── Saved preferences (per browser) ──────────────────────────────────── */
const OPTIONS_KEY = "pbiExport.options";
const BRANDING_KEY = "pbiExport.branding";
const MAX_LOGO_BYTES = 500 * 1024;

const DEFAULT_OPTIONS = {
  pages: "current",        // "current" | "all" | "choose"
  layout: "report",        // "report" | "stacked"
  collapse: true,
  visibleRows: 8,
  includeFilters: true,
  includeNotes: true,
};
const DEFAULT_BRANDING = {
  companyName: "",
  logo: null,              // data URL
  logoWidth: 0,
  logoHeight: 0,
  accent: "#1A73E8",
  tableStyle: "TableStyleMedium2",
  footer: "",
};

// Storage can be unavailable (private windows, blocked site data): fall back to defaults
function loadJson(key, defaults) {
  try {
    const raw = localStorage.getItem(key);
    return raw ? { ...defaults, ...JSON.parse(raw) } : { ...defaults };
  } catch (e) {
    return { ...defaults };
  }
}

function saveJson(key, value) {
  try {
    localStorage.setItem(key, JSON.stringify(value));
    return true;
  } catch (e) {
    return false;
  }
}

let exportOptions = loadJson(OPTIONS_KEY, DEFAULT_OPTIONS);
let branding = loadJson(BRANDING_KEY, DEFAULT_BRANDING);

// Per-session choices that depend on the open report
const chosenPages = new Map();      // reportId → Set of page names
const excludedVisuals = new Map();  // "reportId:pageName" → Set of visual names

/* ── Toasts ───────────────────────────────────────────────────────────── */
function toast(title, text, kind = "success", details = []) {
  const el = document.createElement("div");
  el.className = `toast ${kind}`;
  el.setAttribute("role", kind === "error" ? "alert" : "status");
  el.innerHTML = `<div class="toast-body"><div class="toast-title"></div><div class="toast-text"></div></div>
    <button class="icon-btn" aria-label="Dismiss"><svg viewBox="0 0 24 24" aria-hidden="true"><path d="M6 6l12 12M18 6 6 18" /></svg></button>`;
  el.querySelector(".toast-title").textContent = title;
  el.querySelector(".toast-text").textContent = text || "";
  if (details.length) {
    const d = document.createElement("details");
    d.innerHTML = `<summary>${details.length} note(s)</summary><ul></ul>`;
    details.forEach((n) => {
      const li = document.createElement("li");
      li.textContent = n;
      d.querySelector("ul").appendChild(li);
    });
    el.querySelector(".toast-body").appendChild(d);
  }
  const close = () => el.remove();
  el.querySelector(".icon-btn").addEventListener("click", close);
  $("toasts").appendChild(el);
  if (kind !== "error" && !details.length) setTimeout(close, 8000);
  return el;
}

/* ── Signed-in shell ──────────────────────────────────────────────────── */
function showSignedIn(account) {
  $("landing").hidden = true;
  $("appShell").hidden = false;
  $("userArea").hidden = false;
  $("userName").textContent = account.username;
  const name = account.name || account.username || "";
  $("avatar").textContent = name.split(/[\s@.]+/).filter(Boolean).slice(0, 2).map((p) => p[0].toUpperCase()).join("");
  $("avatar").title = name;
}

/* ── Export options drawer ────────────────────────────────────────────── */
function pageKey(page) {
  return `${currentReport.id}:${page.name}`;
}

async function openOptionsDrawer() {
  $("scrim").hidden = false;
  $("optionsDrawer").hidden = false;

  document.querySelectorAll('input[name="optPages"]').forEach((r) => { r.checked = r.value === exportOptions.pages; });
  document.querySelectorAll('input[name="optLayout"]').forEach((r) => { r.checked = r.value === exportOptions.layout; });
  $("optCollapse").checked = exportOptions.collapse;
  $("optVisibleRows").value = exportOptions.visibleRows;
  $("optVisibleRows").disabled = !exportOptions.collapse;
  $("optFilters").checked = exportOptions.includeFilters;
  $("optNotes").checked = exportOptions.includeNotes;

  await renderPageChecklist();
  await renderVisualChecklist();
  updateDrawerVisibility();
  $("drawerExport").focus();
}

function closeOptionsDrawer() {
  $("scrim").hidden = true;
  $("optionsDrawer").hidden = true;
}

async function renderPageChecklist() {
  const list = $("pageChecklist");
  list.innerHTML = "";
  const [pages, active] = await Promise.all([embeddedReport.getPages(), embeddedReport.getActivePage()]);
  if (!chosenPages.has(currentReport.id)) chosenPages.set(currentReport.id, new Set([active.name]));
  const chosen = chosenPages.get(currentReport.id);
  for (const p of pages) {
    const hidden = p.visibility === models.SectionVisibility.HiddenInViewMode;
    const label = document.createElement("label");
    label.className = "choice";
    label.innerHTML = `<input type="checkbox"> <span></span>`;
    const box = label.querySelector("input");
    box.checked = chosen.has(p.name);
    box.addEventListener("change", () => (box.checked ? chosen.add(p.name) : chosen.delete(p.name)));
    label.querySelector("span").textContent = p.displayName;
    if (hidden) label.querySelector("span").insertAdjacentHTML("beforeend", "<small>hidden page</small>");
    list.appendChild(label);
  }
}

async function renderVisualChecklist() {
  const list = $("visualChecklist");
  list.innerHTML = "";
  const page = await embeddedReport.getActivePage();
  const key = pageKey(page);
  if (!excludedVisuals.has(key)) excludedVisuals.set(key, new Set());
  const excluded = excludedVisuals.get(key);
  const visuals = await listDataVisuals(page);
  if (!visuals.length) {
    list.innerHTML = `<p class="muted">No data visuals on this page.</p>`;
    return;
  }
  for (const v of visuals) {
    const label = document.createElement("label");
    label.className = "choice";
    label.innerHTML = `<input type="checkbox"> <span></span>`;
    const box = label.querySelector("input");
    box.dataset.visual = v.name;
    box.checked = !excluded.has(v.name);
    box.addEventListener("change", () => (box.checked ? excluded.delete(v.name) : excluded.add(v.name)));
    label.querySelector("span").textContent = v.title;
    list.appendChild(label);
  }
}

function setAllVisuals(checked) {
  document.querySelectorAll("#visualChecklist input[type=checkbox]").forEach((box) => {
    box.checked = checked;
    box.dispatchEvent(new Event("change"));
  });
}

function updateDrawerVisibility() {
  const pages = document.querySelector('input[name="optPages"]:checked')?.value || "current";
  $("pageChecklist").hidden = pages !== "choose";
  $("visualsFieldset").hidden = pages !== "current";
}

/* Reads the drawer controls into exportOptions and saves them */
function readDrawerOptions() {
  const visibleRows = parseInt($("optVisibleRows").value, 10);
  exportOptions = {
    pages: document.querySelector('input[name="optPages"]:checked')?.value || "current",
    layout: document.querySelector('input[name="optLayout"]:checked')?.value || "report",
    collapse: $("optCollapse").checked,
    visibleRows: visibleRows > 0 ? Math.min(visibleRows, 1000) : DEFAULT_OPTIONS.visibleRows,
    includeFilters: $("optFilters").checked,
    includeNotes: $("optNotes").checked,
  };
  saveJson(OPTIONS_KEY, exportOptions);
}

/* ── Branding settings ────────────────────────────────────────────────── */
let pendingLogo = null; // { logo, logoWidth, logoHeight } while the dialog is open

function openSettings() {
  $("brandCompany").value = branding.companyName;
  $("brandAccent").value = branding.accent;
  $("brandTableStyle").value = branding.tableStyle;
  $("brandFooter").value = branding.footer;
  $("brandLogo").value = "";
  pendingLogo = { logo: branding.logo, logoWidth: branding.logoWidth, logoHeight: branding.logoHeight };
  showLogoPreview();
  $("settingsDialog").showModal();
}

function showLogoPreview() {
  const img = $("brandLogoPreview");
  img.hidden = !pendingLogo.logo;
  $("brandLogoRemove").hidden = !pendingLogo.logo;
  if (pendingLogo.logo) img.src = pendingLogo.logo;
}

function readLogoFile(file) {
  if (!file) return;
  if (!/^image\/(png|jpeg)$/.test(file.type)) { toast("Logo not added", "Use a PNG or JPG image.", "error"); return; }
  if (file.size > MAX_LOGO_BYTES) { toast("Logo not added", "The image is larger than 500 KB.", "error"); return; }
  const reader = new FileReader();
  reader.onload = () => {
    const img = new Image();
    img.onload = () => {
      pendingLogo = { logo: reader.result, logoWidth: img.naturalWidth, logoHeight: img.naturalHeight };
      showLogoPreview();
    };
    img.src = reader.result;
  };
  reader.readAsDataURL(file);
}

function saveSettings() {
  branding = {
    companyName: $("brandCompany").value.trim(),
    accent: $("brandAccent").value,
    tableStyle: $("brandTableStyle").value,
    footer: $("brandFooter").value.trim(),
    ...pendingLogo,
  };
  if (!saveJson(BRANDING_KEY, branding)) {
    toast("Branding applied for this session", "This browser blocks saved settings, so they reset when you close the tab.", "info");
  } else {
    toast("Branding saved", "Your next export will use it.");
  }
}

/* ── Export ───────────────────────────────────────────────────────────── */
let cancelRequested = false;

async function pagesToExport() {
  const active = await embeddedReport.getActivePage();
  if (exportOptions.pages === "current") return [active];
  const pages = await embeddedReport.getPages();
  if (exportOptions.pages === "all") {
    return pages.filter((p) => p.visibility !== models.SectionVisibility.HiddenInViewMode);
  }
  const chosen = chosenPages.get(currentReport.id) || new Set([active.name]);
  const picked = pages.filter((p) => chosen.has(p.name));
  return picked.length ? picked : [active];
}

function showProgress({ pageIndex, pageCount, pageName, visualIndex, visualCount, visualTitle }) {
  const fraction = ((pageIndex - 1) + (visualCount ? visualIndex / visualCount : 0)) / pageCount;
  $("progressBar").style.width = `${Math.round(fraction * 100)}%`;
  $("progressLabel").textContent = pageCount > 1 ? `Page ${pageIndex} of ${pageCount} · ${pageName}` : pageName;
  $("progressDetail").textContent = visualCount ? `Visual ${visualIndex} of ${visualCount} – ${visualTitle}` : visualTitle;
}

async function runExport() {
  if (!embeddedReport) return;
  cancelRequested = false;
  setExportEnabled(false);
  $("progressCancel").disabled = false;
  $("progressBar").style.width = "0%";
  $("progressLabel").textContent = "Preparing…";
  $("progressDetail").textContent = "";
  $("progressDialog").showModal();
  try {
    const pages = await pagesToExport();
    const options = { ...exportOptions };
    if (options.pages === "current") options.excludedVisuals = excludedVisuals.get(pageKey(pages[0])) || new Set();
    const result = await exportToExcel(pages, options, branding, showProgress, () => cancelRequested);
    const parts = [
      result.sheets > 1 ? `${result.sheets} pages` : null,
      `${result.tables} table${result.tables === 1 ? "" : "s"}`,
      result.kpis ? `${result.kpis} KPI${result.kpis === 1 ? "" : "s"}` : null,
    ].filter(Boolean);
    toast(`Saved ${result.fileName}`, `${parts.join(", ")} · check your Downloads folder.`, "success", result.notes);
  } catch (e) {
    if (e && e.cancelled) {
      toast("Export cancelled", "No file was created.", "info");
    } else {
      log("Export failed: " + errorText(e));
      toast("Export failed", friendlyError(e), "error");
    }
  } finally {
    $("progressDialog").close();
    setExportEnabled(true);
  }
}

/* ── Diagnostics (support) ────────────────────────────────────────────── */
async function runDiagnostics() {
  if (!embeddedReport) { $("diagSummary").textContent = "Open a report first."; return; }
  $("diagRun").disabled = true;
  const tbody = document.querySelector("#diagResults tbody");
  tbody.innerHTML = "";
  $("diagResults").hidden = false;

  const started = performance.now();
  const page = await embeddedReport.getActivePage();
  const visuals = await page.getVisuals();
  log(`Page "${page.displayName}" (${page.name}): ${visuals.length} visual(s)`);
  log(`Report filters: ${describeFilters(await embeddedReport.getFilters())}`);
  log(`Page filters:   ${describeFilters(await page.getFilters())}`);

  const counts = { ok: 0, fail: 0, skip: 0 };
  // Sequential on purpose: Power BI does not support concurrent exportData calls
  for (let i = 0; i < visuals.length; i++) {
    const v = visuals[i];
    const l = v.layout || {};
    let status = "", cls = "", detail = "";
    const t = performance.now();
    try {
      const result = await v.exportData(models.ExportDataType.Summarized, 30000);
      const lines = result.data.trim().split(/\r?\n/);
      status = "OK"; cls = "ok";
      detail = `${Math.max(lines.length - 1, 0)} row(s), ${Math.round(performance.now() - t)} ms · header: ${lines[0].slice(0, 150)}`;
      counts.ok++;
      console.log(`CSV for "${v.title || v.name}":\n`, result.data);
    } catch (e) {
      if (NON_DATA_TYPES.includes(v.type)) {
        status = "no data"; cls = "skip"; detail = `Decorative visual – Power BI returned: ${errorText(e)}`;
        counts.skip++;
      } else {
        status = "FAILED"; cls = "fail"; detail = errorText(e);
        counts.fail++;
      }
    }
    if (v.type === "slicer") {
      try {
        detail += ` · Selection: ${describeFilters((await v.getSlicerState()).filters)}`;
      } catch (e) {
        detail += ` · Selection could not be read: ${errorText(e)}`;
      }
    }
    const tr = document.createElement("tr");
    tr.innerHTML = `<td>${i + 1}</td><td></td><td></td>` +
      `<td>${Math.round(l.x || 0)}, ${Math.round(l.y || 0)}, ${Math.round(l.width || 0)}, ${Math.round(l.height || 0)}</td>` +
      `<td>${isHiddenVisual(v) ? "Hidden" : "Visible"}</td><td class="${cls}">${status}</td><td></td>`;
    tr.children[1].textContent = `${v.title || "(untitled)"} · ${v.name}`;
    tr.children[2].textContent = v.type;
    tr.lastElementChild.textContent = detail;
    tbody.appendChild(tr);
  }

  const seconds = ((performance.now() - started) / 1000).toFixed(1);
  $("diagSummary").textContent = `Exported OK: ${counts.ok} · Failed: ${counts.fail} · No data: ${counts.skip} · ${seconds}s`;
  log(`Diagnostics done in ${seconds}s (full CSVs are in the browser console, F12)`);
  $("diagRun").disabled = false;
}

/* ── Wiring ───────────────────────────────────────────────────────────── */
$("productName").textContent = cfg.productName || "NeenOpal Export for Power BI";
$("signIn").addEventListener("click", () => signIn());
$("signOut").addEventListener("click", () => signOut());

$("workspaceSelect").addEventListener("change", () => loadReports());
$("reportSearch").addEventListener("input", () => renderReportList());
$("toggleSidebar").addEventListener("click", () => $("appShell").classList.toggle("sidebar-collapsed"));
$("linkForm").addEventListener("submit", (e) => {
  e.preventDefault();
  const link = parseReportLink($("reportLink").value);
  if (!link) {
    toast("That doesn't look like a report link", "Copy the address of a report from app.powerbi.com, e.g. https://app.powerbi.com/groups/…/reports/…", "error");
    return;
  }
  openReportById(link.groupId, link.reportId, link.pageName)
    .then(() => { $("reportLink").value = ""; })
    .catch((err) => toast("Couldn't open report", friendlyError(err), "error"));
});

$("exportBtn").addEventListener("click", () => runExport());
$("exportOptionsBtn").addEventListener("click", () => openOptionsDrawer().catch((e) => toast("Couldn't load options", friendlyError(e), "error")));
$("scrim").addEventListener("click", closeOptionsDrawer);
document.querySelectorAll("[data-close-drawer]").forEach((b) => b.addEventListener("click", () => { readDrawerOptions(); closeOptionsDrawer(); }));
document.querySelectorAll('input[name="optPages"]').forEach((r) => r.addEventListener("change", updateDrawerVisibility));
$("optCollapse").addEventListener("change", () => { $("optVisibleRows").disabled = !$("optCollapse").checked; });
$("visualsAll").addEventListener("click", () => setAllVisuals(true));
$("visualsNone").addEventListener("click", () => setAllVisuals(false));
$("drawerOpenSettings").addEventListener("click", () => { readDrawerOptions(); closeOptionsDrawer(); openSettings(); });
$("drawerExport").addEventListener("click", () => { readDrawerOptions(); closeOptionsDrawer(); runExport(); });
document.addEventListener("keydown", (e) => {
  if (e.key === "Escape" && !$("optionsDrawer").hidden) { readDrawerOptions(); closeOptionsDrawer(); }
});

$("openSettings").addEventListener("click", openSettings);
$("brandLogo").addEventListener("change", (e) => readLogoFile(e.target.files[0]));
$("brandLogoRemove").addEventListener("click", () => {
  pendingLogo = { logo: null, logoWidth: 0, logoHeight: 0 };
  $("brandLogo").value = "";
  showLogoPreview();
});
$("settingsSave").addEventListener("click", (e) => {
  e.preventDefault();
  saveSettings();
  $("settingsDialog").close();
});
$("openDiagnostics").addEventListener("click", () => { $("settingsDialog").close(); $("diagDialog").showModal(); });

$("progressCancel").addEventListener("click", () => {
  cancelRequested = true;
  $("progressCancel").disabled = true;
  $("progressDetail").textContent = "Cancelling after the current visual…";
});
// Esc on the progress dialog asks to cancel instead of hiding a running export
$("progressDialog").addEventListener("cancel", (e) => { e.preventDefault(); $("progressCancel").click(); });

$("diagRun").addEventListener("click", () => runDiagnostics().catch((e) => {
  log("Diagnostics error: " + errorText(e));
  $("diagSummary").textContent = friendlyError(e);
  $("diagRun").disabled = false;
}));
$("diagClose").addEventListener("click", () => $("diagDialog").close());

// ?debug=1 opens the diagnostics panel once signed in (for support sessions)
startApp().then(() => {
  if (new URLSearchParams(window.location.search).has("debug") && !$("appShell").hidden) $("diagDialog").showModal();
});
