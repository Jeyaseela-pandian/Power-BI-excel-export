"use strict";

/*
 * Sign-in (MSAL), Power BI REST calls, report picker, report links and embedding.
 * Also holds helpers shared by export.js and ui.js: $, log, errorText, friendlyError,
 * describeFilters, NON_DATA_TYPES, models, embeddedReport, currentReportName.
 */

/* ── Setup ─────────────────────────────────────────────────────────────── */
const PBI_SCOPES = [
  "https://analysis.windows.net/powerbi/api/Report.Read.All",
  "https://analysis.windows.net/powerbi/api/Workspace.Read.All",
  "https://analysis.windows.net/powerbi/api/Dataset.Read.All",
];
const PBI_API = "https://api.powerbi.com/v1.0/myorg";
const models = window["powerbi-client"].models;
const cfg = window.APP_CONFIG;

// Decorative visual types: export is still attempted, but a failure is expected, not an error
const NON_DATA_TYPES = ["textbox", "image", "shape", "basicShape", "actionButton", "pageNavigator", "bookmarkNavigator"];

const msalApp = new msal.PublicClientApplication({
  auth: {
    clientId: cfg.clientId,
    authority: `https://login.microsoftonline.com/${cfg.tenantId}`,
    redirectUri: cfg.redirectUri,
  },
  cache: { cacheLocation: "sessionStorage" },
});

let embeddedReport = null;
let currentReport = null;
let currentReportName = "Report";
let reports = [];

const $ = (id) => document.getElementById(id);

/* Activity log – shown in the Diagnostics dialog for support */
function log(msg) {
  const el = $("log");
  if (el) el.textContent += `[${new Date().toLocaleTimeString()}] ${msg}\n`;
  console.log(`[export] ${msg}`);
}

// powerbi-client rejects with plain objects, not Error instances
function errorText(e) {
  if (!e) return "Unknown error";
  return e.detailedMessage || e.message || (typeof e === "string" ? e : JSON.stringify(e));
}

/* Messages a business user can act on, instead of raw API errors */
function friendlyError(e) {
  const t = errorText(e);
  if (/AADSTS65001|AADSTS90094|consent/i.test(t)) return "Your organisation's IT admin needs to approve this app before it can read Power BI reports.";
  if (/\b401\b/.test(t)) return "Your session has expired, or your account has no Power BI licence. Sign out and sign in again.";
  if (/\b403\b/.test(t)) return "You don't have permission to open this content in Power BI.";
  if (/\b404\b/.test(t)) return "That report wasn't found, or you don't have access to it.";
  if (/export/i.test(t) && /(disabled|not allowed|not permitted|permission)/i.test(t)) return "Data export is turned off for this report, or by your Power BI admin.";
  return t;
}

/* ── Sign-in and tokens (MSAL) ─────────────────────────────────────────── */
function currentAccount() {
  return msalApp.getActiveAccount() || msalApp.getAllAccounts()[0] || null;
}

async function getToken() {
  const account = currentAccount();
  if (!account) throw new Error("Not signed in");
  try {
    const result = await msalApp.acquireTokenSilent({ scopes: PBI_SCOPES, account });
    return result.accessToken;
  } catch (e) {
    if (e instanceof msal.InteractionRequiredAuthError) {
      await msalApp.acquireTokenRedirect({ scopes: PBI_SCOPES, account });
    }
    throw e;
  }
}

function signIn() {
  return msalApp.loginRedirect({ scopes: PBI_SCOPES });
}

function signOut() {
  return msalApp.logoutRedirect({ postLogoutRedirectUri: cfg.redirectUri });
}

/* ── Power BI REST API ─────────────────────────────────────────────────── */
async function pbiGet(path) {
  const res = await fetch(PBI_API + path, {
    headers: { Authorization: `Bearer ${await getToken()}` },
  });
  if (!res.ok) {
    throw new Error(`${res.status} ${res.statusText} on ${path}: ${await res.text()}`);
  }
  const body = await res.json();
  return body.value !== undefined ? body.value : body;
}

async function loadWorkspaces() {
  const groups = await pbiGet("/groups");
  const sel = $("workspaceSelect");
  sel.innerHTML = "";
  sel.append(new Option("My workspace", ""));
  groups.sort((a, b) => a.name.localeCompare(b.name)).forEach((g) => sel.append(new Option(g.name, g.id)));
  log(`Found ${groups.length} workspace(s) + My workspace`);
  await loadReports();
}

async function loadReports() {
  const groupId = $("workspaceSelect").value;
  renderReportList("Loading reports…");
  try {
    reports = (await pbiGet(groupId ? `/groups/${groupId}/reports` : "/reports"))
      .sort((a, b) => a.name.localeCompare(b.name));
    log(`Found ${reports.length} report(s)`);
    renderReportList();
  } catch (e) {
    reports = [];
    renderReportList(friendlyError(e));
  }
}

/* Report list in the sidebar, filtered by the search box */
function renderReportList(message) {
  const list = $("reportList");
  list.innerHTML = "";
  if (message) {
    list.innerHTML = `<li class="list-message"></li>`;
    list.firstChild.textContent = message;
    return;
  }
  const query = $("reportSearch").value.trim().toLowerCase();
  const shown = reports.filter((r) => r.name.toLowerCase().includes(query));
  if (shown.length === 0) {
    list.innerHTML = `<li class="list-message"></li>`;
    list.firstChild.textContent = reports.length ? "No reports match your search." : "No reports in this workspace.";
    return;
  }
  for (const r of shown) {
    const li = document.createElement("li");
    const btn = document.createElement("button");
    btn.type = "button";
    btn.className = currentReport && currentReport.id === r.id ? "active" : "";
    btn.innerHTML = `<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M4 20V10M10 20V4M16 20v-7M22 20H2" /></svg><span></span>`;
    btn.lastChild.textContent = r.name;
    btn.title = r.name;
    btn.addEventListener("click", () => openReport(r).catch((e) => toast("Couldn't open report", friendlyError(e), "error")));
    li.appendChild(btn);
    list.appendChild(li);
  }
}

/* ── Report links ──────────────────────────────────────────────────────── */
/*
 * Accepts links copied from app.powerbi.com, e.g.
 *   https://app.powerbi.com/groups/<workspaceId|me>/reports/<reportId>/<pageName>?experience=power-bi
 */
function parseReportLink(text) {
  try {
    const url = new URL(text.trim());
    const m = url.pathname.match(/\/groups\/([^/]+)\/reports\/([0-9a-f-]{36})(?:\/([^/?#]+))?/i);
    if (!m) return null;
    return { groupId: m[1].toLowerCase() === "me" ? null : m[1], reportId: m[2], pageName: m[3] || null };
  } catch (e) {
    return null;
  }
}

async function openReportById(groupId, reportId, pageName) {
  const report = await pbiGet(groupId ? `/groups/${groupId}/reports/${reportId}` : `/reports/${reportId}`);
  await openReport(report, pageName);
}

/*
 * Deep link (used by the "Export to Excel" button inside a Power BI report):
 *   http://localhost:5500/?groupId=<workspaceId>&reportId=<reportId>&pageName=<pageName>
 * (omit groupId for My workspace)
 */
async function openFromUrl() {
  const params = new URLSearchParams(window.location.search);
  const reportId = params.get("reportId");
  if (!reportId) return;
  log(`Opening from link: report ${reportId}, page ${params.get("pageName") || "(default)"}`);
  await openReportById(params.get("groupId"), reportId, params.get("pageName"));
}

/* ── Embedding ─────────────────────────────────────────────────────────── */
async function openReport(report, pageName) {
  const container = $("report");
  powerbi.reset(container);
  currentReport = report;
  currentReportName = report.name;
  setExportEnabled(false);
  $("emptyState").hidden = true;
  $("reportTitle").textContent = report.name;
  $("pageTitle").textContent = "Loading…";
  renderReportList();

  const config = {
    type: "report",
    id: report.id,
    embedUrl: report.embedUrl,
    accessToken: await getToken(),
    tokenType: models.TokenType.Aad,
    settings: { panes: { filters: { visible: true } } },
  };
  if (pageName) config.pageName = pageName;

  embeddedReport = powerbi.embed(container, config);
  embeddedReport.on("loaded", async () => {
    log(`Report "${report.name}" loaded`);
    setExportEnabled(true);
    await showActivePage();
  });
  embeddedReport.on("pageChanged", (e) => {
    $("pageTitle").textContent = e.detail.newPage.displayName;
  });
  embeddedReport.on("error", (e) => {
    log("Embed error: " + errorText(e.detail));
    toast("Power BI reported a problem", friendlyError(e.detail), "error");
  });
}

async function showActivePage() {
  try {
    const page = await embeddedReport.getActivePage();
    $("pageTitle").textContent = page.displayName;
  } catch (e) {
    $("pageTitle").textContent = "";
  }
}

function setExportEnabled(enabled) {
  $("exportBtn").disabled = !enabled;
  $("exportOptionsBtn").disabled = !enabled;
}

/* ── Filter descriptions (shared with export.js) ──────────────────────── */
const OPERATOR_TEXT = {
  GreaterThan: ">", GreaterThanOrEqual: "≥", LessThan: "<", LessThanOrEqual: "≤",
  Is: "=", IsNot: "≠", Contains: "contains", DoesNotContain: "does not contain",
  StartsWith: "starts with", DoesNotStartWith: "does not start with",
  IsBlank: "is blank", IsNotBlank: "is not blank",
};
const ISO_DATE = /^\d{4}-\d{2}-\d{2}T/;

/* Readable filter values: ISO dates → local dates, field-parameter refs ('Table'[Column]) → Column */
function formatFilterValue(v) {
  if (typeof v !== "string") return String(v);
  if (ISO_DATE.test(v) && !isNaN(new Date(v))) return new Date(v).toLocaleDateString();
  const fieldRef = v.match(/^'?[^'[]*'?\[(.+)\]$/);
  return fieldRef ? fieldRef[1] : v;
}

function describeConditions(f) {
  const c = f.conditions;
  // Date range slicer: "≥ start AND < end" → "start – (end - 1 day)"
  if (c.length === 2 && c[0].operator === "GreaterThanOrEqual" && c[1].operator === "LessThan" &&
      ISO_DATE.test(c[0].value) && ISO_DATE.test(c[1].value)) {
    const end = new Date(new Date(c[1].value).getTime() - 24 * 3600 * 1000);
    return `${formatFilterValue(c[0].value)} – ${end.toLocaleDateString()}`;
  }
  return c.map((x) => `${OPERATOR_TEXT[x.operator] || x.operator}${x.value !== undefined ? " " + formatFilterValue(x.value) : ""}`)
    .join(f.logicalOperator === "Or" ? " or " : " and ");
}

function describeFilters(filters) {
  if (!filters || filters.length === 0) return "none";
  return filters.map((f) => {
    const field = f.target ? (f.target.column || f.target.measure || f.target.hierarchyLevel || "?") : "?";
    const values = f.values ? f.values.map(formatFilterValue).join(", ")
      : f.conditions ? describeConditions(f) : "";
    return `${field}${values ? ": " + values : ""}`;
  }).join(" | ");
}

/* ── Start-up (called from ui.js once every script is loaded) ─────────── */
async function startApp() {
  if (!cfg.clientId || cfg.clientId.startsWith("PASTE")) {
    showStartupMessage("This app isn't configured yet: add the clientId and tenantId in config.js.");
    $("signIn").disabled = true;
    return;
  }
  try {
    await msalApp.initialize();
    const result = await msalApp.handleRedirectPromise();
    if (result && result.account) msalApp.setActiveAccount(result.account);
  } catch (e) {
    log("Sign-in error: " + errorText(e));
    showStartupMessage(friendlyError(e));
  }
  const account = currentAccount();
  if (!account) return;

  showSignedIn(account);
  try {
    await loadWorkspaces();
    await openFromUrl();
  } catch (e) {
    log("Start-up error: " + errorText(e));
    toast("Couldn't load your Power BI content", friendlyError(e), "error");
  }
}

function showStartupMessage(text) {
  const el = $("startupMessage");
  el.textContent = text;
  el.hidden = false;
}
