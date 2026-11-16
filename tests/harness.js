/*
 * Loads the app's scripts (app.js, colors.js, export.js) into a sandbox, as the browser would,
 * with stand-ins for the browser, sign-in and Power BI libraries, so the export's logic can
 * be checked without a browser. Also reads report files (.pbix) for the checks on real reports.
 */
const fs = require("fs");
const os = require("os");
const path = require("path");
const vm = require("vm");
const zlib = require("zlib");

const ROOT = path.join(__dirname, "..");
// Real report files for the checks that need them (skipped when there are none)
const REPORTS_DIR = process.env.REPORTS_DIR || path.join(os.homedir(), "Downloads");

/* A sandbox with the app's scripts loaded: { ctx, run(code), json(code) } */
function loadApp() {
  const ctx = {
    console, URL, TextDecoder, structuredClone, atob, btoa, encodeURIComponent,
    performance: { now: () => 0 },
    APP_CONFIG: {},
    msal: {
      PublicClientApplication: function () {
        return { getActiveAccount: () => null, getAllAccounts: () => [], handleRedirectPromise: async () => null, initialize: async () => {} };
      },
      InteractionRequiredAuthError: Error,
    },
    window: { "powerbi-client": { models: { ExportDataType: { Summarized: 0 }, TokenType: { Aad: 0 } } }, APP_CONFIG: {} },
    document: { getElementById: () => null, querySelector: () => null, querySelectorAll: () => [], addEventListener() {} },
    navigator: {},
    localStorage: { getItem: () => null, setItem() {} },
  };
  vm.createContext(ctx);
  for (const file of ["app.js", "colors.js", "export.js"]) {
    try {
      vm.runInContext(fs.readFileSync(path.join(ROOT, file), "utf8"), ctx, { filename: file });
    } catch (e) {
      // app.js starts the page when loaded; the functions the checks use are defined before that
    }
  }
  ctx.log = () => {};
  const run = (code) => vm.runInContext(code, ctx);
  // A result as plain values (the sandbox's arrays and objects aren't this file's, for comparing)
  const json = (code) => JSON.parse(run(`JSON.stringify(${code})`));
  return { ctx, run, json };
}

/* A .zip file's entries (a .pbix is one): { names, read(name) → Buffer } */
function readZip(file) {
  const buf = fs.readFileSync(file);
  let end = buf.length - 22;
  while (end >= 0 && buf.readUInt32LE(end) !== 0x06054b50) end--;
  if (end < 0) throw new Error(`${file} is not a zip file`);
  const count = buf.readUInt16LE(end + 10);
  let at = buf.readUInt32LE(end + 16);
  const entries = {};
  for (let i = 0; i < count; i++) {
    const method = buf.readUInt16LE(at + 10), size = buf.readUInt32LE(at + 20);
    const nameLen = buf.readUInt16LE(at + 28), extraLen = buf.readUInt16LE(at + 30), commentLen = buf.readUInt16LE(at + 32);
    const local = buf.readUInt32LE(at + 42);
    const name = buf.toString("utf8", at + 46, at + 46 + nameLen);
    entries[name] = { method, size, local };
    at += 46 + nameLen + extraLen + commentLen;
  }
  const read = (name) => {
    const e = entries[name];
    if (!e) return null;
    const start = e.local + 30 + buf.readUInt16LE(e.local + 26) + buf.readUInt16LE(e.local + 28);
    const data = buf.subarray(start, start + e.size);
    return e.method === 8 ? zlib.inflateRawSync(data) : Buffer.from(data);
  };
  return { names: Object.keys(entries), read };
}

/* A report file as the app's design reader expects a zip (JSZip's shape) */
function designZip(file) {
  const z = readZip(file);
  return {
    files: Object.fromEntries(z.names.map((n) => [n, {}])),
    file: (n) => (z.names.includes(n) ? { async: async (type) => (type === "uint8array" ? new Uint8Array(z.read(n)) : z.read(n).toString("utf8")) } : null),
  };
}

/* The report files in REPORTS_DIR, one per name (copies such as "Report (2).pbix" left out) */
function reportFiles() {
  if (!fs.existsSync(REPORTS_DIR)) return [];
  const seen = new Set();
  return fs.readdirSync(REPORTS_DIR)
    .filter((f) => /\.pbix$/i.test(f))
    .sort()
    .filter((f) => { const key = f.replace(/ \(\d+\)(?=\.pbix$)/i, "").toLowerCase(); if (seen.has(key)) return false; seen.add(key); return true; })
    .map((f) => path.join(REPORTS_DIR, f));
}

/* One report file by name, or null when it isn't there */
function reportFile(name) {
  const file = path.join(REPORTS_DIR, name);
  return fs.existsSync(file) ? file : null;
}

module.exports = { loadApp, readZip, designZip, reportFiles, reportFile, REPORTS_DIR };
