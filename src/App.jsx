import React, { useState, useEffect, useMemo, useRef } from "react";
import {
  GraduationCap, ShieldCheck, Upload, CheckCircle2, XCircle,
  Download, Printer, LogOut, Home, AlertCircle, Loader2, ChevronRight,
  Building2, KeyRound, FileSpreadsheet,
} from "lucide-react";
import { supabase } from "./supabaseClient.js";
import Papa from "papaparse";
import * as XLSX from "xlsx";

/* ============================================================================
   ID generators & small utilities
   ========================================================================== */
let __seq = 0;
function genId(prefix) {
  __seq += 1;
  return prefix + "_" + Date.now().toString(36) + "_" + __seq;
}
function refNo(seqNum) {
  return "AEP-" + new Date().getFullYear() + "-" + String(seqNum).padStart(5, "0");
}
function todayStr() {
  return new Date().toISOString().slice(0, 10);
}
function formatDateTime(iso) {
  if (!iso) return "";
  try {
    return new Date(iso).toLocaleString("en-IN", { day: "2-digit", month: "short", year: "numeric", hour: "numeric", minute: "2-digit" });
  } catch (e) { return iso; }
}
function randomPassword(len = 8) {
  const chars = "ABCDEFGHJKMNPQRSTUVWXYZabcdefghjkmnpqrstuvwxyz23456789";
  let out = "";
  for (let i = 0; i < len; i++) out += chars[Math.floor(Math.random() * chars.length)];
  return out;
}
function toCSV(rows, columns) {
  const esc = (v) => `"${String(v ?? "").replace(/"/g, '""')}"`;
  const header = columns.map((c) => esc(c.label)).join(",");
  const body = rows.map((r) => columns.map((c) => esc(c.get(r))).join(",")).join("\n");
  return header + "\n" + body;
}
function download(filename, content, mime) {
  const blob = new Blob([content], { type: mime });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url; a.download = filename;
  document.body.appendChild(a); a.click(); document.body.removeChild(a);
  URL.revokeObjectURL(url);
}
// Builds a header row + one row per record from a { label, get(r) } column list —
// shared by the CSV export above and the multi-sheet workbook export below.
function toAOA(rows, columns) {
  return [columns.map((c) => c.label), ...rows.map((r) => columns.map((c) => c.get(r)))];
}
// Multi-sheet .xlsx export (one sheet per document type) using the "xlsx" (SheetJS
// Community Edition) package. This only ever WRITES a workbook built from the app's own
// already-trusted in-memory submission data — it never parses an uploaded/untrusted file
// (bulk examiner CSV upload still goes through papaparse) — so the package's known
// vulnerabilities (prototype pollution / ReDoS in its *parsing* code paths) don't apply here.
function downloadWorkbook(filename, sheets) {
  const wb = XLSX.utils.book_new();
  sheets.forEach(({ name, rows, cols }) => {
    const ws = XLSX.utils.aoa_to_sheet(toAOA(rows, cols));
    XLSX.utils.book_append_sheet(wb, ws, name.slice(0, 31));
  });
  XLSX.writeFile(wb, filename);
}
function money(n) {
  return "Rs. " + Math.round(n || 0).toLocaleString("en-IN");
}
// Fills {token} placeholders in a declaration/certification sentence from the
// derived flat field values (e.g. "{name}" -> derived.flat.fullName), falling
// back to a blank underscore run (matching the original paper form's blank
// lines) when a value isn't available yet.
function interpolate(str, vars) {
  return String(str || "").replace(/\{(\w+)\}/g, (_, key) => {
    const v = vars[key];
    return v !== undefined && v !== null && v !== "" ? v : "__________";
  });
}
function numToWords(num) {
  num = Math.round(num || 0);
  if (num === 0) return "Rupees Zero Only";
  const a = ["", "One", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight", "Nine", "Ten",
    "Eleven", "Twelve", "Thirteen", "Fourteen", "Fifteen", "Sixteen", "Seventeen", "Eighteen", "Nineteen"];
  const b = ["", "", "Twenty", "Thirty", "Forty", "Fifty", "Sixty", "Seventy", "Eighty", "Ninety"];
  function two(n) { if (n < 20) return a[n]; return b[Math.floor(n / 10)] + (n % 10 ? " " + a[n % 10] : ""); }
  function three(n) { if (n > 99) return a[Math.floor(n / 100)] + " Hundred" + (n % 100 ? " " + two(n % 100) : ""); return two(n); }
  let str = "", n2 = num;
  const crore = Math.floor(n2 / 10000000); n2 %= 10000000;
  const lakh = Math.floor(n2 / 100000); n2 %= 100000;
  const thousand = Math.floor(n2 / 1000); n2 %= 1000;
  const rest = n2;
  if (crore) str += three(crore) + " Crore ";
  if (lakh) str += three(lakh) + " Lakh ";
  if (thousand) str += three(thousand) + " Thousand ";
  if (rest) str += three(rest);
  return "Rupees " + str.trim() + " Only";
}
function sortByOrder(list) {
  return (list || []).slice().sort((a, b) => (a.order || 0) - (b.order || 0));
}

/* ============================================================================
   Schema-driven form engine (pure functions) — the same computation model
   used across all three documents: sections contain either a grid of fields
   or a dynamic table; any field/column can declare `computed` (a formula
   referencing sibling field ids or literal numbers, with an optional
   min/max clamp) so the admin can configure fee rates without code; any
   section/field can declare `showIf` for conditional visibility (e.g. the
   MBBS/M.Sc./PG remuneration blocks).
   ========================================================================== */
function resolveToken(token, ctx) {
  if (typeof token === "number") return token;
  const n = parseFloat(ctx[token]);
  return isNaN(n) ? 0 : n;
}
function computeFromDef(def, ctx) {
  const c = def.computed;
  if (!c) return undefined;
  let raw = 0;
  if (c.op === "multiply") {
    raw = (c.inputs || []).reduce((acc, t) => acc * resolveToken(t, ctx), 1);
    if (c.rate !== undefined) raw *= resolveToken(c.rate, ctx);
  } else if (c.op === "sumThenMultiply") {
    const s = (c.inputs || []).reduce((acc, t) => acc + resolveToken(t, ctx), 0);
    raw = s * resolveToken(c.rate !== undefined ? c.rate : 1, ctx);
  } else if (c.op === "sum") {
    raw = (c.inputs || []).reduce((acc, t) => acc + resolveToken(t, ctx), 0);
  } else if (c.op === "dateDiffDays") {
    // Inclusive day count between two date-field ids, e.g. onward date -> return date.
    // Used for "No. of days vehicle required" style fields — not a numeric multiply/sum,
    // so it reads the two input tokens as raw date strings from ctx rather than resolveToken.
    const [fromTok, toTok] = c.inputs || [];
    const fromStr = ctx[fromTok];
    const toStr = ctx[toTok];
    const from = fromStr ? new Date(fromStr) : null;
    const to = toStr ? new Date(toStr) : null;
    if (!from || !to || isNaN(from.getTime()) || isNaN(to.getTime())) {
      raw = 0;
    } else {
      const fromTs = new Date(from.getFullYear(), from.getMonth(), from.getDate()).getTime();
      const toTs = new Date(to.getFullYear(), to.getMonth(), to.getDate()).getTime();
      raw = Math.max(1, Math.round((toTs - fromTs) / 86400000) + 1);
    }
  }
  if (c.clamp) {
    const gateVal = resolveToken(c.clamp.gate, ctx);
    if (gateVal > 0) { raw = Math.max(raw, c.clamp.min); raw = Math.min(raw, c.clamp.max); }
    else raw = 0;
  }
  return Math.round(raw * 100) / 100;
}
function evalShowIf(showIf, ctx) {
  if (!showIf || !showIf.field) return true;
  return String(ctx[showIf.field]) === String(showIf.equals);
}
function deriveForm(schema, data) {
  data = data || {};
  const flat = {};
  (schema.sections || []).forEach((s) => {
    if (s.kind !== "fields") return;
    (s.fields || []).forEach((f) => {
      const raw = data?.[s.id]?.[f.id];
      flat[f.id] = f.type === "number" ? (parseFloat(raw) || 0) : (raw !== undefined && raw !== "" ? raw : (f.default ?? ""));
    });
  });
  for (let pass = 0; pass < 2; pass++) {
    (schema.sections || []).forEach((s) => {
      if (s.kind !== "fields") return;
      (s.fields || []).forEach((f) => { if (f.computed) flat[f.id] = computeFromDef(f, flat); });
    });
  }
  const resolvedRowsBySection = {};
  const tableTotals = {};
  (schema.sections || []).forEach((s) => {
    if (s.kind !== "table") return;
    const rows = (data?.[s.id] && data[s.id].length ? data[s.id] : [{}]);
    const resolved = rows.map((row) => {
      const rowFlat = {};
      (s.columns || []).forEach((c) => { rowFlat[c.id] = c.type === "number" ? (parseFloat(row[c.id]) || 0) : (row[c.id] ?? ""); });
      (s.columns || []).forEach((c) => { if (c.computed) rowFlat[c.id] = computeFromDef(c, rowFlat); });
      return rowFlat;
    });
    resolvedRowsBySection[s.id] = resolved;
    const total = s.totalColumnId ? resolved.reduce((sum, r) => sum + (parseFloat(r[s.totalColumnId]) || 0), 0) : 0;
    tableTotals[s.id] = total;
    flat[s.id + "__total"] = total;
  });
  let grandTotal = 0;
  (schema.sections || []).forEach((s) => {
    if (!evalShowIf(s.showIf, flat)) return;
    if (s.kind === "fields") {
      (s.fields || []).forEach((f) => { if (f.includeInGrandTotal) grandTotal += parseFloat(flat[f.id]) || 0; });
    } else if (s.kind === "table" && s.includeInGrandTotal) {
      grandTotal += tableTotals[s.id] || 0;
    }
  });
  return { flat, grandTotal, resolvedRowsBySection, tableTotals };
}
function autofillFromProfile(schema, data, profile) {
  const next = { ...data };
  (schema.sections || []).forEach((s) => {
    if (s.kind !== "fields") return;
    (s.fields || []).forEach((f) => {
      if (!f.autofillFrom) return;
      const existing = next?.[s.id]?.[f.id];
      if (existing) return;
      const parts = f.autofillFrom.split("."); // profile.fieldId
      const val = parts[0] === "profile" ? profile?.[parts[1]] : undefined;
      if (val) next[s.id] = { ...(next[s.id] || {}), [f.id]: val };
    });
  });
  return next;
}

/* ============================================================================
   Category rate-section template — used both to pre-seed the default categories
   beyond MBBS/MSC/PG (B.Sc. Nursing, B.Sc. Allied & Health Care) and to
   auto-build a matching section whenever the admin types a brand-new category
   name in the Examiners tab (see addNewCategory()). Mirrors the MBBS shape:
   question papers + evaluation (books/students), with a Rs.3000–15000
   claimable range on the evaluation amount, admin-editable afterwards under
   Exam Rates / Form Builder.
   ========================================================================== */
function slugifyCategory(name) {
  return (name || "").trim().toUpperCase().replace(/[^A-Z0-9]+/g, "_").replace(/^_+|_+$/g, "");
}
function buildCategorySection(categoryName) {
  const slug = slugifyCategory(categoryName);
  const p = slug.toLowerCase() + "_";
  return {
    id: "remun_" + slug, title: `I. Remuneration — ${categoryName}`, printTitle: "I. REMUNERATION:-", order: 3, kind: "fields", hiddenFromExaminer: true, printAsTable: true,
    showIf: { field: "examCategory", equals: categoryName },
    fields: [
      { id: p + "qpCount", label: "No. of Question Papers Set", type: "number", default: 0, order: 1 },
      // Fixed, admin-configured rate (Admin > Exam Rates) — same shape as MBBS/PG —
      // rather than a manual per-submission entry.
      { id: p + "qpAmount", label: "Amount Payable (Question Paper)", type: "number", readOnly: true, order: 3, computed: { op: "multiply", inputs: [p + "qpCount", 500] } },
      { id: p + "evalBooks", label: "No. of Answer Books Evaluated", type: "number", default: 0, order: 4 },
      { id: p + "evalStudents", label: "No. of Students Examined (Practical/Clinical/Viva)", type: "number", default: 0, order: 5 },
      { id: p + "evalAmount", label: "Amount Payable @ Rs.100/paper or candidate", type: "number", readOnly: true, order: 6, computed: { op: "sumThenMultiply", inputs: [p + "evalBooks", p + "evalStudents"], rate: 100, clamp: { min: 3000, max: 15000, gate: p + "evalBooks" } } },
      { id: p + "net", label: "Net Amount Payable (Remuneration)", type: "number", readOnly: true, order: 7, computed: { op: "sum", inputs: [p + "qpAmount", p + "evalAmount"] }, includeInGrandTotal: true },
    ],
  };
}

/* ============================================================================
   Default (seed) form templates — admin-editable afterwards via the Form
   Builder tab. Same shape as examiner_form_configs.schema in schema.sql.
   ========================================================================== */
// UG and PG examiners get two genuinely separate Remuneration Bill documents (their own
// title, their own fee-table section, editable independently under Exam Rates/Form
// Builder) — matching how AIIMS issues separate physical forms per level — but they
// share an identical surrounding structure (examiner details, travelling/accommodation/
// DA tables, declarations, bank details, the Travel Origin routing question). This
// builds that shared block fresh each call so the two schemas are independent objects,
// not accidentally aliased to the same section arrays.
function buildSharedRemunSections() {
  return [
    // The reference documents have no heading at all above this block — the personal-
    // details fields start right under the "BILL/REMUNERATION FORM" title. "Examiner
    // Details" is an on-screen-only label to orient the examiner; noTitleInPrint drops it.
    { id: "profile", title: "Examiner Details", order: 2, kind: "fields", noTitleInPrint: true, fields: [
      { id: "fullName", label: "Full Name (BLOCK letters)", type: "text", required: true, order: 1 },
      { id: "designation", label: "Designation", type: "text", required: true, order: 2 },
      { id: "basicPay", label: "Basic Pay & Scale of Pay", type: "text", order: 3 },
      { id: "mobile", label: "Mobile Number", type: "tel", required: true, order: 4 },
      { id: "email", label: "Email ID", type: "email", required: true, order: 5 },
      { id: "address", label: "Address", type: "textarea", order: 6 },
      { id: "examSubject", label: "Name of Examination & Subject", type: "text", required: true, order: 7 },
      { id: "purposeJourney", label: "Purpose of Journey", type: "text", order: 8 },
      { id: "examDates", label: "Dates of Examination / Meeting", type: "text", order: 9 },
    ]},
    { id: "travel", title: "II. Travelling", printTitle: "II.TRAVELLING:-", order: 4, kind: "table", includeInGrandTotal: true, totalColumnId: "amt", addRowLabel: "Add journey", columns: [
      // Date & time entered as two separate mandatory fields (rather than one free-text
      // "dd-mm-yyyy hh:mm" box) so a journey can't be saved without a time — matches the
      // original "Departure, date & time" / "Arrival, date & time" columns; printDisplay
      // recombines the pair back into a single column on the generated document.
      { id: "depDate", label: "Departure Date", type: "date", required: true, order: 1, colWidth: 105, printDisplay: { label: "Departure, date & time", pairWith: "depTime" } },
      { id: "depTime", label: "Departure Time", type: "time", required: true, order: 2, colWidth: 95, excludeFromPrint: true },
      { id: "arrDate", label: "Arrival Date", type: "date", required: true, order: 3, colWidth: 105, printDisplay: { label: "Arrival, date & time", pairWith: "arrTime" } },
      { id: "arrTime", label: "Arrival Time", type: "time", required: true, order: 4, colWidth: 95, excludeFromPrint: true },
      { id: "dist", label: "Distance (km)", type: "number", order: 5, colWidth: 90 },
      { id: "mode", label: "Mode of Journey", type: "text", order: 6, placeholder: "bus/car/train/air", colWidth: 110 },
      // Wider than a plain number column so the full claim figure is always visible.
      { id: "amt", label: "Claim Amount (Rs.)", type: "number", order: 7, colWidth: 130 },
      // Widened per feedback — the previous width was too cramped to type a useful note.
      { id: "rem", label: "Remarks", type: "text", order: 8, colWidth: 220 },
    ]},
    // Accommodation is provided directly by the AIIMS guest house, so this isn't
    // collected from the examiner today — kept in the schema (hiddenFromExaminer)
    // in case it's needed again later.
    // printTitle intentionally reproduces the reference document's own spelling
    // ("ACCOMODATION", missing the second M) verbatim rather than "correcting" it —
    // an exact text match matters more here than fixing a typo in the source form.
    { id: "accom", title: "III. Accommodation", printTitle: "III. ACCOMODATION:-", order: 5, kind: "table", hiddenFromExaminer: true, includeInGrandTotal: true, totalColumnId: "amt", addRowLabel: "Add stay", columns: [
      // Wide enough that the "dd-mm-yyyy hh:mm" placeholder isn't cut off (same fix
      // applied to the Travelling table's Claim Amount/Remarks columns).
      { id: "arr", label: "Arrival Date & Time", type: "text", order: 1, placeholder: "dd-mm-yyyy hh:mm", colWidth: 175 },
      { id: "dep", label: "Departure Date & Time", type: "text", order: 2, placeholder: "dd-mm-yyyy hh:mm", colWidth: 175 },
      { id: "days", label: "No. of Days Stayed", type: "number", order: 3, colWidth: 100 },
      { id: "rate", label: "Rate / Day (Rs.)", type: "number", order: 4, colWidth: 110 },
      { id: "amt", label: "Claim Amount (Rs.)", type: "number", readOnly: true, order: 5, colWidth: 120, computed: { op: "multiply", inputs: ["days", "rate"] } },
      { id: "rem", label: "Remarks", type: "text", order: 6, colWidth: 160 },
    ]},
    { id: "da", title: "IV. Daily Allowance", printTitle: "IV. DAILY ALLOWANCE:-", order: 6, kind: "table", hiddenFromExaminer: true, includeInGrandTotal: true, totalColumnId: "amt", addRowLabel: "Add period", columns: [
      { id: "from", label: "Date From", type: "date", order: 1 },
      { id: "to", label: "Date To", type: "date", order: 2 },
      { id: "days", label: "No. of Days", type: "number", order: 3 },
      { id: "rate", label: "D.A. Rate (Rs.)", type: "number", order: 4 },
      { id: "amt", label: "Total D.A. (Rs.)", type: "number", readOnly: true, order: 5, computed: { op: "multiply", inputs: ["days", "rate"] } },
    ]},
    { id: "declaration", title: "Declaration", order: 7, kind: "fields", fields: [
      // Both of these are already covered in prose by the printed declaration notes
      // above (payment mode is listed inline in declaration #1, TA/DA status in #2),
      // so they're kept out of the generated document to avoid a duplicate/extra line.
      { id: "payMode", label: "Payment Mode", type: "select", options: ["NEFT"], default: "NEFT", order: 1, excludeFromPrint: true },
      { id: "taClaim", label: "TA/DA Declaration", type: "select", options: ["have not claimed / will not claim from parent institute"], default: "have not claimed / will not claim from parent institute", order: 2, excludeFromPrint: true },
      // Office-use only, matches the reference form's separate "Total Payable Amount"
      // line in the Office Use box — normally equal to the computed Grand Total, but
      // left admin-editable in case some claimed items are disallowed on verification.
      // Blank (unset) falls back to the Grand Total automatically wherever it's shown.
      { id: "totalPayableAmount", label: "Total Payable Amount (office use — leave blank to match Grand Total)", type: "number", hiddenFromExaminer: true, order: 3, excludeFromPrint: true },
    ]},
    { id: "bank", title: "Bank Details", printTitle: "Bank Details:", order: 8, kind: "fields", fields: [
      { id: "bankName", label: "Bank Name", type: "text", required: true, order: 1 },
      { id: "branchName", label: "Branch Name", type: "text", required: true, order: 2 },
      { id: "acctType", label: "Account Type", type: "select", options: ["Savings", "Current"], default: "Savings", required: true, order: 3 },
      { id: "acctNumber", label: "Account Number", type: "text", required: true, order: 4 },
      { id: "bankAcName", label: "Name (as per bank record)", type: "text", required: true, order: 5 },
      { id: "ifsc", label: "IFSC Code", type: "text", required: true, order: 6 },
      { id: "pan", label: "PAN No. (linked to Aadhar)", type: "text", required: true, order: 7 },
    ]},
    // App-only routing question (decides whether Vehicle/Flight sections apply) — not
    // part of the physical Remuneration Bill, so left out of the generated document.
    { id: "travelOrigin", title: "Travel Origin", order: 9, kind: "fields", excludeFromPrint: true, fields: [
      { id: "originType", label: "Are you travelling from within Telangana or outside Telangana?", type: "select", options: ["Within Telangana", "Outside Telangana"], required: true, order: 1 },
      { id: "vehicleChoice", label: "Vehicle for local travel", type: "select", options: ["Institute vehicle required", "Own vehicle — no institute vehicle needed"], required: true, order: 2, showIf: { field: "originType", equals: "Within Telangana" } },
    ]},
  ];
}
// Declaration text, signature line, and Office Use block are identical wording across
// all three reference docs (MBBS/MSC/PG) — shared here so UG and PG stay in sync if
// that wording ever needs updating, rather than drifting apart in two copies.
function remunMeta(title) {
  return {
    title,
    org1: "ALL INDIA INSTITUTE OF MEDICAL SCIENCES, BIBINAGAR, TELANGANA",
    subtitle: title + " — Examination Cell",
    notes: [
      "I {name} claiming an amount of Rs. {amount} in lieu of Remuneration/Transport allowance/Accommodation/Daily allowance in Cash/NEFT/RTGS/Cheque.",
      "I hereby certify that I will claim/Have not claimed TA/DA or any other remuneration paid by AIIMS/to be paid by AIIMS Bibinagar from parent Institute.",
      "In reference to receipt of remuneration mentioned at Sl No. 1. I hereby declare that I am non official in terms of MoHFW Letter NO. Z-28016/60/2017-SSH(EO-3108590) dated 25 Jun 2017.",
    ],
    signatureBlock: {
      examinerLine: "Signature of faculty/Officials",
      officeUse: {
        title: "Office Use only:",
        certifiedText: "Certified that {name} was present in AIIMS, Bibinagar on {examDates} for the purpose of {purposeJourney} and his/her original bills were checked and verified.",
        authoritySignature: "Officer In charge",
      },
    },
  };
}
// Which examCategory values belong to which of the two separate Remuneration Bill
// documents — drives both schema selection (remunerationSchemaFor) and the category
// dropdown options on each form. These two lists seed the *default* schemas (used
// before formConfigs finishes loading, and as a last-resort fallback below); once
// formConfigs is available, remunerationFormIdFor reads the LIVE option lists off
// each schema's billMeta.examCategory field instead, so a category the admin adds
// later via AdminExaminers.addNewCategory() (which isn't in these hardcoded arrays)
// still resolves to whichever form it was actually added to.
const UG_CATEGORIES = ["MBBS", "BSC NURSING", "BSC ALLIED & HEALTH CARE"];
const PG_CATEGORIES = ["PG", "MSC"];
function isUGCategory(cat) { return UG_CATEGORIES.includes((cat || "").toUpperCase()); }
function remunerationFormIdFor(cat, formConfigs) {
  const val = (cat || "").trim().toUpperCase();
  if (formConfigs) {
    const ugField = formConfigs.ugRemuneration && getCategoryFieldFromSchema(formConfigs.ugRemuneration);
    const pgField = formConfigs.pgRemuneration && getCategoryFieldFromSchema(formConfigs.pgRemuneration);
    const ugOptions = (ugField?.options || []).map((o) => o.trim().toUpperCase());
    const pgOptions = (pgField?.options || []).map((o) => o.trim().toUpperCase());
    if (ugOptions.includes(val)) return "ugRemuneration";
    if (pgOptions.includes(val)) return "pgRemuneration";
  }
  return isUGCategory(cat) ? "ugRemuneration" : "pgRemuneration";
}
function remunerationSchemaFor(cat, formConfigs) { return formConfigs[remunerationFormIdFor(cat, formConfigs)]; }

function defaultFormConfigs() {
  return {
    ugRemuneration: {
      formId: "ugRemuneration",
      meta: remunMeta("UG Bill / Remuneration Form"),
      sections: [
        // The reference doc has "Bill.No.:……." at the very top right of the page —
        // printAlign right-aligns it on the generated document (still after the letterhead
        // rather than literally above it, given the linear top-to-bottom layout here).
        { id: "billMeta", title: "Bill Details", order: 1, kind: "fields", noTitleInPrint: true, printAlign: "right", fields: [
          { id: "billNo", label: "Bill No. (office use)", type: "text", order: 1 },
          // examCategory drives which of the remun_* blocks below is shown — needed by
          // the app to pick the right fee structure, but it isn't its own line on the
          // physical form (each category has its own separate printed form), so it's
          // left out of the generated document to match the reference format exactly.
          { id: "examCategory", label: "Examiner Category", type: "select", required: true, order: 2, options: [...UG_CATEGORIES], default: "MBBS", excludeFromPrint: true },
        ]},
        // Workload counts & rates below are entered/adjusted by the admin (Submissions
        // tab) after reviewing each examiner's category — never shown to the examiner
        // (hiddenFromExaminer), but still part of the final generated bill.
        { id: "remun_MBBS", title: "I. Remuneration — MBBS", printTitle: "I. REMUNERATION:-", order: 3, kind: "fields", hiddenFromExaminer: true, printAsTable: true, showIf: { field: "examCategory", equals: "MBBS" }, fields: [
          { id: "mb_qpCount", label: "No. of Question Papers Set", type: "number", default: 0, order: 1 },
          // Rate is a fixed, admin-configured amount (Admin > Exam Rates) applied
          // automatically — same shape as PG's Rate per Question Paper — rather than a
          // manual per-submission entry, so the amount calculates by itself.
          { id: "mb_qpAmount", label: "Amount Payable (Question Paper)", type: "number", readOnly: true, order: 3, computed: { op: "multiply", inputs: ["mb_qpCount", 500] } },
          { id: "mb_evalBooks", label: "No. of Answer Books Evaluated", type: "number", default: 0, order: 4 },
          { id: "mb_evalStudents", label: "No. of Students Examined (Practical/Clinical/Viva)", type: "number", default: 0, order: 5 },
          { id: "mb_evalAmount", label: "Amount Payable @ Rs.100/paper or candidate", type: "number", readOnly: true, order: 6, computed: { op: "sumThenMultiply", inputs: ["mb_evalBooks", "mb_evalStudents"], rate: 100, clamp: { min: 3000, max: 15000, gate: "mb_evalBooks" } } },
          { id: "mb_net", label: "Net Amount Payable (Remuneration)", type: "number", readOnly: true, order: 7, computed: { op: "sum", inputs: ["mb_qpAmount", "mb_evalAmount"] }, includeInGrandTotal: true },
        ]},
        // B.Sc. Nursing / B.Sc. Allied & Health Care: same shape as MBBS (question
        // papers + evaluation, Rs.3000-15000 claimable range on evaluation) — built
        // via buildCategorySection() so any admin-added UG category later gets the
        // identical structure automatically (see AdminExaminers.addNewCategory()).
        { ...buildCategorySection("BSC NURSING"), showIf: { field: "examCategory", equals: "BSC NURSING" } },
        { ...buildCategorySection("BSC ALLIED & HEALTH CARE"), showIf: { field: "examCategory", equals: "BSC ALLIED & HEALTH CARE" } },
        ...buildSharedRemunSections(),
      ],
    },
    pgRemuneration: {
      formId: "pgRemuneration",
      meta: remunMeta("PG Bill / Remuneration Form"),
      sections: [
        // The reference doc has "Bill.No.:……." at the very top right of the page —
        // printAlign right-aligns it on the generated document (still after the letterhead
        // rather than literally above it, given the linear top-to-bottom layout here).
        { id: "billMeta", title: "Bill Details", order: 1, kind: "fields", noTitleInPrint: true, printAlign: "right", fields: [
          { id: "billNo", label: "Bill No. (office use)", type: "text", order: 1 },
          { id: "examCategory", label: "Examiner Category", type: "select", required: true, order: 2, options: [...PG_CATEGORIES], default: "PG", excludeFromPrint: true },
        ]},
        { id: "remun_MSC", title: "I. Remuneration — M.Sc. / Allied", printTitle: "I. REMUNERATION:-", order: 3, kind: "fields", hiddenFromExaminer: true, printAsTable: true, showIf: { field: "examCategory", equals: "MSC" }, fields: [
          { id: "ms_qpCount", label: "No. of Question Papers Set (@ Rs.500/paper)", type: "number", default: 0, order: 1 },
          { id: "ms_qpAmount", label: "Amount Payable (Question Paper)", type: "number", readOnly: true, order: 2, computed: { op: "multiply", inputs: ["ms_qpCount", 500] } },
          { id: "ms_theoryCount", label: "Answer Books Evaluated (Theory, Rs.1000/candidate, Min 5000/Max 15000)", type: "number", default: 0, order: 3 },
          { id: "ms_theoryAmount", label: "Amount Payable (Theory Evaluation)", type: "number", readOnly: true, order: 4, computed: { op: "multiply", inputs: ["ms_theoryCount", 1000], clamp: { min: 5000, max: 15000, gate: "ms_theoryCount" } } },
          { id: "ms_pracCount", label: "Students Examined (Practical/Clinical/Viva, Rs.2000/candidate, Min 5000/Max 15000)", type: "number", default: 0, order: 5 },
          { id: "ms_pracAmount", label: "Amount Payable (Practical/Clinical/Viva)", type: "number", readOnly: true, order: 6, computed: { op: "multiply", inputs: ["ms_pracCount", 2000], clamp: { min: 5000, max: 15000, gate: "ms_pracCount" } } },
          { id: "ms_thesisCount", label: "No. of Thesis Evaluated (@ Rs.4000/candidate)", type: "number", default: 0, order: 7 },
          { id: "ms_thesisAmount", label: "Amount Payable (Thesis Evaluation)", type: "number", readOnly: true, order: 8, computed: { op: "multiply", inputs: ["ms_thesisCount", 4000] } },
          { id: "ms_net", label: "Net Amount Payable (Remuneration)", type: "number", readOnly: true, order: 9, computed: { op: "sum", inputs: ["ms_qpAmount", "ms_theoryAmount", "ms_pracAmount", "ms_thesisAmount"] }, includeInGrandTotal: true },
        ]},
        { id: "remun_PG", title: "I. Remuneration — PG (MD/MS/DM/MCh)", printTitle: "I. REMUNERATION:-", order: 3, kind: "fields", hiddenFromExaminer: true, printAsTable: true, showIf: { field: "examCategory", equals: "PG" }, fields: [
          { id: "pg_qpCount", label: "No. of Question Papers Set (@ Rs.500/paper)", type: "number", default: 0, order: 1 },
          { id: "pg_qpAmount", label: "Amount Payable (Question Paper)", type: "number", readOnly: true, order: 2, computed: { op: "multiply", inputs: ["pg_qpCount", 500] } },
          { id: "pg_thesisCount", label: "No. of Thesis Evaluated (@ Rs.4000/candidate)", type: "number", default: 0, order: 3 },
          { id: "pg_thesisAmount", label: "Amount Payable (Thesis Evaluation)", type: "number", readOnly: true, order: 4, computed: { op: "multiply", inputs: ["pg_thesisCount", 4000] } },
          // No. of Papers is its own count alongside Students Examined — both charged at
          // the same Rs.2000/each rate and clamped together (Min 5000/Max 15000), rather
          // than one combined field, mirroring the MBBS/B.Sc. evaluation shape (sumThenMultiply).
          { id: "pg_papersCount", label: "No. of Papers Evaluated", type: "number", default: 0, order: 5 },
          { id: "pg_pracCount", label: "Students Examined (Practical/Clinical/Viva)", type: "number", default: 0, order: 6 },
          // Gates the clamp on EITHER count being filled — not just pg_papersCount — so an
          // admin who only fills one of the two (e.g. candidates but no separate papers this
          // time) still gets the Min 5000/Max 15000 floor/ceiling applied, instead of the
          // whole amount silently zeroing out.
          { id: "pg_evalGate", label: "Papers/Candidates Evaluated (either)", type: "number", readOnly: true, order: 6.5, excludeFromPrint: true, hiddenFromExaminer: true, computed: { op: "sum", inputs: ["pg_papersCount", "pg_pracCount"] } },
          { id: "pg_pracAmount", label: "Amount Payable (Papers + Practical/Clinical/Viva, Rs.2000/each, Min 5000/Max 15000)", type: "number", readOnly: true, order: 7, computed: { op: "sumThenMultiply", inputs: ["pg_papersCount", "pg_pracCount"], rate: 2000, clamp: { min: 5000, max: 15000, gate: "pg_evalGate" } } },
          { id: "pg_net", label: "Net Amount Payable (Remuneration)", type: "number", readOnly: true, order: 8, computed: { op: "sum", inputs: ["pg_qpAmount", "pg_thesisAmount", "pg_pracAmount"] }, includeInGrandTotal: true },
        ]},
        ...buildSharedRemunSections(),
      ],
    },
    vehicle: {
      formId: "vehicle",
      meta: {
        title: "Vehicle Requisition",
        org1: "ALL INDIA INSTITUTE OF MEDICAL SCIENCES, BIBINAGAR, HYDERABAD",
        subtitle: "To be submitted well in advance, two to three days before",
        // Verbatim bilingual header block from vehicle format.docx (paragraphs 0-4) —
        // shown above the org1 line the same way the original leads with the Hindi title.
        preHeaderLines: [
          "अखिल भारतीय आयुर्विज्ञान संस्थान, बीबीनगर, हैदराबाद",
          "“An INI under the aegis of Ministry of Health & Family Welfare, Govt. of India”",
          "हैदराबाद मेट्रोपोलिटन क्षेत्र (एच एम आर), तेलंगाना – 508126, भारत",
          "Hyderabad Metropolitan Region (HMR), Telangana - 508126, India",
        ],
        // "Important Instructions" footer, verbatim from the reference doc (paragraphs 41-44).
        notesTitle: "IMPORTANT INSTRUCTIONS",
        // The reference doc has these AFTER the signature/office-use/approval chain, not before.
        notesPosition: "end",
        notes: [
          "To be submitted within well in advance, two to three days before",
          "Signature and seal required at the designated place",
          "Serial number 4 and 5 mandatory to be filled",
          "Email ID and Mobile No to be filled properly",
        ],
        // Matches the reference doc's bottom section exactly (paragraphs 19-35): the
        // Indenting Officer's own signature line and routing, then the office's vehicle
        // availability / taxi-hire approval chain ending in Deputy Director (Admin.).
        vehicleSignoff: {
          date: true,
          signatureLine: "Signature of the Indenting Officer",
          routing: ["Through Controlling Officer/HoD", "To Assistant Administrative Officer"],
          deptVehicleLabel: "Dept. Vehicle",
          taxiLine1: "Submitted for approval for hiring of a taxi as proposed above",
          approverLeft: "Faculty I/C Vehicle",
          approverRight: "Administrative Officer",
          taxiLine2: "Kind concurrence for hiring of Taxi as proposed above",
          finalApprover: "Deputy Director (Admin.)",
        },
      },
      sections: [
        { id: "main", title: "Vehicle Requisition", order: 1, kind: "fields", noTitleInPrint: true, fields: [
          // Office contact details for the requisition — a fixed, admin-configured
          // default (edit under Form Builder), not something the examiner enters.
          { id: "v_officer", label: "Name(s) & Designation of Indenting Officer", type: "text", default: "Dean (Examination), AIIMS Bibinagar", hiddenFromExaminer: true, order: 1 },
          { id: "v_mobile", label: "Mobile / Tel No.", type: "tel", default: "", hiddenFromExaminer: true, order: 2 },
          { id: "v_email", label: "Email ID", type: "email", default: "", hiddenFromExaminer: true, order: 3 },
          // Auto-synced from the Remuneration form's "Travel Origin" answer (not entered here
          // directly) — purely drives which of the sections below apply to this examiner.
          { id: "v_originType", label: "Origin Type", type: "text", hiddenFromExaminer: true, order: 4, excludeFromPrint: true },
          { id: "v_place", label: "Place of visit in brief", type: "text", default: "AIIMS Bibinagar", order: 5 },
          { id: "v_purpose", label: "Purpose of Journey", type: "text", order: 6, autofillFrom: "profile.purposeJourney" },
          // Unconditional on the reference form (item 6) — not restricted to local travel only.
          // Office-filled rather than examiner-entered (admin sets/edits it in the
          // Submissions review, same pattern as v_officer/v_mobile/v_email above).
          { id: "v_distance", label: "Approximate Distance in KM", type: "number", default: "", hiddenFromExaminer: true, order: 7 },
          // Office-use only — not its own line on the reference form's front page, but drives
          // the "Dept. Vehicle: Available / Not Available" line in the Office Use block below.
          { id: "v_deptVehicle", label: "Dept. Vehicle Available?", type: "select", options: ["Available", "Not Available"], hiddenFromExaminer: true, order: 8, excludeFromPrint: true },
        ]},
        // Out-of-station examiners (e.g. from NIMHANS Bangalore) travel from their own
        // home/institute to THEIR local departure airport under their own arrangement —
        // AIIMS Bibinagar can't send a vehicle that far, it only reimburses that leg as a
        // travelling allowance, which the examiner already claims in the "II. Travelling"
        // table above (in the Remuneration form). AIIMS only actually dispatches a vehicle
        // for the leg it's physically able to cover: Hyderabad Airport <-> AIIMS Bibinagar.
        { id: "legAirportAiims", title: "Vehicle — Hyderabad Airport to AIIMS Bibinagar", order: 2, kind: "fields", noTitleInPrint: true, showIf: { field: "v_originType", equals: "Outside Telangana" }, fields: [
          // App-only routing question — decides whether this leg's date/time fields apply
          // (and drives the "claim via Travelling table" message when declined) but isn't
          // itself a line on the reference vehicle format, so it's excluded from print.
          { id: "v_leg2Needed", label: "Do you need an institute vehicle for this leg?", type: "select", options: ["Yes, I need institute vehicle", "No, I will arrange this myself"], default: "Yes, I need institute vehicle", required: true, order: 1, excludeFromPrint: true },
          { id: "v_leg2DateOnward", label: "Onward — Hyderabad Airport to AIIMS Bibinagar — Date", type: "date", required: true, order: 2, showIf: { field: "v_leg2Needed", equals: "Yes, I need institute vehicle" } },
          { id: "v_leg2TimeOnward", label: "Onward — Hyderabad Airport to AIIMS Bibinagar — Time", type: "time", required: true, order: 3, showIf: { field: "v_leg2Needed", equals: "Yes, I need institute vehicle" } },
          { id: "v_leg2DateReturn", label: "Return — AIIMS Bibinagar to Hyderabad Airport — Date", type: "date", required: true, order: 4, showIf: { field: "v_leg2Needed", equals: "Yes, I need institute vehicle" } },
          { id: "v_leg2TimeReturn", label: "Return — AIIMS Bibinagar to Hyderabad Airport — Time", type: "time", required: true, order: 5, showIf: { field: "v_leg2Needed", equals: "Yes, I need institute vehicle" } },
        ]},
        // Within-Telangana examiners who asked for an institute vehicle: a single local leg,
        // with the number of days the vehicle is needed for computed automatically from the
        // onward/return dates rather than asked separately.
        // Mirrors the reference form's item 3(a)/3(b) — "Vehicle required on Date: ___
        // Time hrs From: ___ To: ___" repeated twice — as an onward + return pair, each
        // with its own From/To time window rather than a single point-in-time.
        { id: "local", title: "Vehicle Requisition — Local Travel", order: 4, kind: "fields", noTitleInPrint: true, showIf: { field: "v_originType", equals: "Within Telangana" }, fields: [
          { id: "v_dateOnward", label: "3(a) Vehicle required on Date — Onward", type: "date", required: true, order: 1 },
          { id: "v_timeOnward", label: "Time hrs — From", type: "time", required: true, order: 2, printCombineRange: "v_timeOnwardTo" },
          { id: "v_timeOnwardTo", label: "Time hrs — To", type: "time", required: true, order: 3, excludeFromPrint: true },
          { id: "v_dateReturn", label: "3(b) Vehicle required on Date — Return", type: "date", required: true, order: 4 },
          { id: "v_timeReturn", label: "Time hrs — From", type: "time", required: true, order: 5, printCombineRange: "v_timeReturnTo" },
          { id: "v_timeReturnTo", label: "Time hrs — To", type: "time", required: true, order: 6, excludeFromPrint: true },
          { id: "v_daysRequired", label: "No. of Days Vehicle Required", type: "number", readOnly: true, order: 7, computed: { op: "dateDiffDays", inputs: ["v_dateOnward", "v_dateReturn"] } },
        ]},
      ],
    },
    flight: {
      formId: "flight",
      meta: {
        title: "Air Tickets Booking Details",
        org1: "AIIMS BIBINAGAR",
        subtitle: "Booking / Cancellation Requisition (Hyderabad)",
        notes: ["* — Mandatory fields"],
        // Matches the reference xlsx's fixed "Request by & Contact no" line under the title.
        extraHeaderLine: "Request by & Contact no: AIIMS, Hyderabad & 7013039551",
        // Matches the reference xlsx's signing block exactly (Hindi + English), bottom-right
        // of the sheet — see rows 20-23 of Flight Booking _Format.xlsx.
        footerAuthority: [
          "संकाय प्रभारी प्रशासन / Faculty Incharge Administration",
          "एम्स बीबीनगर / AIIMS Bibinagar-508126",
          "हैदराबाद / Hyderabad",
          "संपर्क / Contact : 08685279306",
        ],
      },
      sections: [
        // Field order below follows the reference xlsx's own top-to-bottom row order
        // (Client Code/Name+Remarks row 4, Booking/Cancellation row 5, Airline Deal
        // code row 6, Name of Passenger row 7, Gender/DOB row 18, Mobile no row 19,
        // FFP No row 20, Meal Request row 21, Seat Request row 22, Email Id row 23,
        // PNR/Ticket no row 24, Employee ID no row 25, GSTR No row 26) so the generated
        // document reads top-to-bottom the same way the original proforma does.
        { id: "main", title: "Air Tickets Booking Details", order: 1, kind: "fields", noTitleInPrint: true, fields: [
          { id: "f_client", label: "Client Code / Name", type: "text", default: "AG000201101", readOnly: true, order: 1 },
          { id: "f_remarks", label: "Remarks", type: "textarea", order: 2 },
          { id: "f_type", label: "Booking / Cancellation", type: "select", options: ["Booking", "Cancellation"], default: "Booking", order: 3 },
          { id: "f_deal", label: "Airline Deal Code", type: "text", hiddenFromExaminer: true, order: 4 },
          { id: "f_name", label: "Name of Passenger", type: "text", required: true, order: 5, autofillFrom: "profile.fullName" },
          { id: "f_gender", label: "Gender / DOB", type: "text", required: true, order: 6, placeholder: "e.g. M / 01-Jan-1980" },
          { id: "f_mobile", label: "Mobile No.", type: "tel", required: true, order: 7, autofillFrom: "profile.mobile" },
          { id: "f_ffp", label: "FFP No.", type: "text", hiddenFromExaminer: true, order: 8 },
          // Split into a Yes/No gate + conditional Veg/Non-Veg for cleaner input than the
          // original's single free-text "Meal Request:" line — printCombine recombines it
          // back into that one line ("Meal Request: Veg" / "Meal Request: No") on the
          // generated document, so the printed form still has just the single field.
          { id: "f_mealRequest", label: "Meal Request", type: "select", options: ["No", "Yes"], default: "No", order: 9, printCombine: "f_mealType" },
          { id: "f_mealType", label: "Meal Type", type: "select", options: ["Veg", "Non-Veg"], required: true, order: 10, showIf: { field: "f_mealRequest", equals: "Yes" }, excludeFromPrint: true },
          { id: "f_seat", label: "Seat Request", type: "text", order: 11 },
          { id: "f_email", label: "Email ID", type: "email", required: true, order: 12, autofillFrom: "profile.email" },
          // PNR only needs asking when cancelling, but the reference document always has
          // this line printed (blank if not applicable) — printAlways keeps it on the
          // generated document regardless of the Booking/Cancellation selection.
          { id: "f_pnr", label: "PNR / Ticket No. (if cancellation)", type: "text", order: 13, showIf: { field: "f_type", equals: "Cancellation" }, printAlways: true },
          { id: "f_empid", label: "Employee ID No.", type: "text", order: 14 },
          { id: "f_gstr", label: "GSTR No.", type: "text", order: 15 },
        ]},
        // Date and Time were originally one free-text "dd-mm-yyyy hh:mm" column — split into
        // separate, individually-mandatory Date and Time columns so an examiner can't enter
        // only a date and skip the time. printDisplay recombines them back into the single
        // "Date & Time" column the reference xlsx actually has (see row 29 of the format)
        // on the generated document, matching the same pattern used for the Remuneration
        // Bill's Travelling itinerary table.
        { id: "travel", title: "Travel Programme", order: 2, kind: "table", addRowLabel: "Add sector", columns: [
          { id: "date", label: "Date", type: "date", required: true, order: 1, printDisplay: { label: "Date & Time", pairWith: "time" } },
          { id: "time", label: "Time", type: "time", required: true, order: 2, excludeFromPrint: true },
          { id: "from", label: "From", type: "text", order: 3 },
          { id: "to", label: "To", type: "text", order: 4 },
          { id: "cls", label: "Class", type: "text", order: 5, placeholder: "Economy/Business" },
          { id: "no", label: "Flight No.", type: "text", order: 6 },
        ]},
      ],
    },
  };
}
function defaultSettings() {
  return {
    instituteName: "All India Institute of Medical Sciences, Bibinagar",
    portalSubtitle: "Examination Cell — Examiner Bill / Requisition Portal",
    publicBaseUrl: "",
    logoUrl: "",
    // Floor applied to the Remuneration Bill's "Total Payable Amount" (not the raw
    // "Total Amount claimed" figure) when the computed Grand Total falls short —
    // admin-configurable from the Exam Rates tab, defaults to Rs.3000.
    minRemuneration: 3000,
  };
}

/* ============================================================================
   Shared style tokens & small components — same visual language as the
   Hall Ticket app (inline styles, no CSS framework).
   ========================================================================== */
const inputStyle = {
  width: "100%", boxSizing: "border-box", padding: "8px 10px", fontSize: 14,
  border: "1px solid #cfd8e3", borderRadius: 6, background: "#fff", color: "#1c2b3a",
};
function Field({ label, children, required, hint }) {
  return (
    <div style={{ display: "block", marginBottom: 14 }}>
      <span style={{ fontSize: 13, fontWeight: 600, color: "#274566" }}>
        {label} {required && <span style={{ color: "#a13a2f" }}>*</span>}
      </span>
      <div style={{ marginTop: 4 }}>{children}</div>
      {hint && <div style={{ fontSize: 11.5, color: "#7a8794", marginTop: 3 }}>{hint}</div>}
    </div>
  );
}
function Btn({ children, onClick, variant = "primary", disabled, style, type = "button" }) {
  const base = {
    padding: "9px 16px", fontSize: 13.5, fontWeight: 600, borderRadius: 6, cursor: disabled ? "not-allowed" : "pointer",
    border: "1px solid transparent", display: "inline-flex", alignItems: "center", gap: 6, opacity: disabled ? 0.55 : 1,
  };
  const variants = {
    primary: { background: "#1a3a5c", color: "#fff" },
    gold: { background: "#a9762f", color: "#fff" },
    outline: { background: "#fff", color: "#1a3a5c", border: "1px solid #1a3a5c" },
    danger: { background: "#fff", color: "#a13a2f", border: "1px solid #a13a2f" },
    success: { background: "#2f6b45", color: "#fff" },
    ghost: { background: "transparent", color: "#1a3a5c" },
  };
  return (
    <button type={type} disabled={disabled} onClick={onClick} style={{ ...base, ...variants[variant], ...style }}>
      {children}
    </button>
  );
}
function Header({ title, subtitle, onBack, right }) {
  return (
    <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", padding: "14px 22px", background: "#1a3a5c", color: "#fff", flexWrap: "wrap", gap: 8 }}>
      <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
        {onBack && (
          <button onClick={onBack} aria-label="Home" style={{ display: "flex", alignItems: "center", gap: 5, background: "rgba(255,255,255,0.08)", border: "none", borderRadius: 6, color: "#cfe0f2", cursor: "pointer", padding: "5px 10px", fontSize: 12, fontWeight: 600 }}>
            <Home size={16} /> Home
          </button>
        )}
        <Building2 size={22} color="#e0b86a" />
        <div>
          <div style={{ fontWeight: 700, fontSize: 15.5, letterSpacing: 0.2 }}>{title}</div>
          {subtitle && <div style={{ fontSize: 11.5, color: "#b9cbe0" }}>{subtitle}</div>}
        </div>
      </div>
      <div>{right}</div>
    </div>
  );
}
function Grid2({ children }) { return <div className="grid-2col">{children}</div>; }
// The circled step number (n) is an on-screen navigation aid for the examiner filling a
// long multi-section form — it has no equivalent in any of the reference documents, which
// use plain section labels (or no label at all) rather than a numbered sequence. printMode
// drops the circle and just renders the title text, so the generated document doesn't
// invent a numbering scheme the physical form never had.
function SectionTitle({ n, title, printMode }) {
  if (printMode) {
    return <div style={{ fontSize: 12, fontWeight: 700, color: "#1c2b3a", margin: "16px 0 8px" }}>{title}</div>;
  }
  return (
    <div style={{ display: "flex", alignItems: "center", gap: 8, margin: "22px 0 12px" }}>
      <div style={{ width: 22, height: 22, borderRadius: "50%", background: "#1a3a5c", color: "#fff", fontSize: 11, fontWeight: 700, display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0 }}>{n}</div>
      <div style={{ fontSize: 13.5, fontWeight: 700, color: "#1c2b3a" }}>{title}</div>
    </div>
  );
}
function TabBtn({ active, children, onClick }) {
  return (
    <button onClick={onClick} style={{
      background: active ? "#fff" : "transparent", color: active ? "#1a3a5c" : "#cfe0f2",
      border: "none", borderRadius: 6, padding: "6px 12px", fontSize: 12.5, fontWeight: 600, cursor: "pointer",
    }}>{children}</button>
  );
}
function StatusPill({ status }) {
  const map = {
    submitted: { bg: "#fdf3e3", color: "#8a6116", label: "Submitted" },
    generated: { bg: "#e7f3ea", color: "#2f6b45", label: "Generated" },
    reopened: { bg: "#fdeee3", color: "#a15a1a", label: "Reopened for editing" },
    invited: { bg: "#eaf0f7", color: "#1a3a5c", label: "Invited" },
  };
  const s = map[status] || map.invited;
  return <span style={{ background: s.bg, color: s.color, fontSize: 11.5, fontWeight: 700, padding: "3px 10px", borderRadius: 12 }}>{s.label}</span>;
}
function DetailRow({ label, value }) {
  return (
    <div style={{ display: "flex", fontSize: 12.5, marginBottom: 6, gap: 8 }}>
      <span style={{ color: "#7a8794", minWidth: 160 }}>{label}</span>
      <span style={{ color: "#1c2b3a", fontWeight: 500 }}>{value}</span>
    </div>
  );
}
function StatCard({ label, value, color }) {
  return (
    <div style={{ background: "#fff", border: "1px solid #dde3ea", borderRadius: 8, padding: "14px 16px", minWidth: 130 }}>
      <div style={{ fontSize: 11.5, color: "#7a8794", marginBottom: 4 }}>{label}</div>
      <div style={{ fontSize: 22, fontWeight: 700, color: color || "#1c2b3a" }}>{value}</div>
    </div>
  );
}

/* ============================================================================
   Generic schema-driven field/table/form renderers — used by the examiner
   form, the admin submission editor, and the printable documents alike.
   ========================================================================== */
function SchemaFieldInput({ def, value, onChange, readOnly }) {
  const disabled = readOnly || def.readOnly;
  const style = disabled ? { ...inputStyle, background: "#f7f9fb", color: "#33455c", fontWeight: 600 } : inputStyle;
  if (def.type === "select") {
    return (
      <select style={style} value={value ?? def.default ?? ""} disabled={disabled} onChange={(e) => onChange(e.target.value)}>
        <option value="">Select…</option>
        {(def.options || []).map((o) => <option key={o} value={o}>{o}</option>)}
      </select>
    );
  }
  if (def.type === "textarea") {
    return <textarea style={{ ...style, minHeight: 60 }} value={value ?? ""} disabled={disabled} onChange={(e) => onChange(e.target.value)} />;
  }
  if (def.type === "time") {
    // A native <input type="time"> picker's 12-hour/24-hour display follows the
    // browser/OS locale, not the page — the `lang="en-GB"` attribute is meant to force
    // 24-hour but doesn't reliably do so on every Chrome/Windows combination (confirmed
    // by the user: it kept showing an AM/PM scroller regardless). A plain text field
    // with a strict HH:MM pattern sidesteps that entirely — there's no picker UI whose
    // format could vary, so it's unambiguously 24-hour on every browser/OS.
    return (
      <input
        type="text"
        inputMode="numeric"
        style={style}
        value={value ?? ""}
        readOnly={disabled}
        placeholder="HH:MM"
        maxLength={5}
        pattern="([01][0-9]|2[0-3]):[0-5][0-9]"
        title="24-hour format, e.g. 14:30"
        onChange={(e) => onChange(e.target.value.replace(/[^0-9:]/g, "").slice(0, 5))}
        onBlur={(e) => {
          // Best-effort normalize on leaving the field — "9:5" / "930" -> "09:05" — but
          // leave anything it can't confidently parse as typed, so a partial/invalid
          // entry doesn't get silently mangled.
          const raw = e.target.value.trim();
          if (!raw) return;
          const m = raw.match(/^(\d{1,2}):?(\d{0,2})$/);
          if (!m) return;
          const hh = m[1].padStart(2, "0");
          const mm = (m[2] || "00").padEnd(2, "0");
          if (parseInt(hh, 10) <= 23 && parseInt(mm, 10) <= 59) onChange(`${hh}:${mm}`);
        }}
      />
    );
  }
  return (
    <input
      type={def.type === "number" ? "number" : (def.type || "text")}
      style={style}
      value={value ?? ""}
      readOnly={disabled}
      placeholder={def.placeholder}
      onChange={(e) => onChange(e.target.value)}
    />
  );
}
function SchemaFieldsSection({ section, values, flat, onFieldChange, readOnly, n, hideExaminerOnly, printMode }) {
  if (!evalShowIf(section.showIf, flat)) return null;
  if (hideExaminerOnly && section.hiddenFromExaminer) return null;
  // printAlways: a field whose visibility is conditional on-screen (e.g. "PNR / Ticket
  // No." only needs asking when Booking/Cancellation = Cancellation) but which is a
  // fixed, always-present line on the reference document — it must still print (blank
  // if not filled) rather than disappear, or the generated document would have a field
  // missing compared to the original template.
  let fields = sortByOrder(section.fields).filter((f) => (printMode && f.printAlways) || evalShowIf(f.showIf, flat));
  if (hideExaminerOnly) fields = fields.filter((f) => !f.hiddenFromExaminer);
  // printMode = the generated document handed to the examination cell / physically
  // signed — must match the uploaded reference formats exactly, so app-only helper
  // fields (a category selector, a per-submission rate input that isn't its own
  // column on the original paper form, etc.) are left out here even though they're
  // still shown and editable in the admin's on-screen submission editor.
  if (printMode) fields = fields.filter((f) => !f.excludeFromPrint);
  if (fields.length === 0) return null;
  // printCombine/printCombineRange: fold a conditional or paired follow-up field back
  // into its parent's single line on the generated document — e.g. Meal Request "Yes" +
  // Meal Type "Veg" -> "Veg", or a From-time + To-time pair -> "HH:MM to HH:MM" — since
  // the reference forms have just the one line each, not two.
  function printValue(f) {
    if (f.printCombine) return flat[f.id] === "Yes" ? (flat[f.printCombine] || flat[f.id]) : flat[f.id];
    if (f.printCombineRange) return [flat[f.id], flat[f.printCombineRange]].filter(Boolean).join(" to ");
    return flat[f.id];
  }
  // Some sections (Bill Details, Examiner Details) exist purely to organize the on-screen
  // form and have no heading at all on the physical document — noTitleInPrint drops them.
  // Others (the fee tables, Bank Details) DO have a heading on the reference document, but
  // worded exactly as printed there (e.g. "I. REMUNERATION:-") rather than the friendlier
  // on-screen title — printTitle supplies that literal wording for print only.
  const showTitle = section.title && !(printMode && section.noTitleInPrint);
  const titleText = printMode && section.printTitle ? section.printTitle : section.title;
  // A handful of sections (the per-category remuneration fee blocks) are laid out as an
  // actual bordered grid table on the reference documents rather than a label/value list
  // — printAsTable reproduces that on the generated document specifically (the on-screen
  // admin/examiner form still uses the normal Grid2 list below, unaffected).
  if (printMode && section.printAsTable) {
    return (
      <>
        {showTitle && <SectionTitle n={n} title={titleText} printMode={printMode} />}
        <div style={{ overflowX: "auto" }}>
          <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 11 }}>
            <thead>
              <tr style={{ background: "#f7f9fb" }}>
                {fields.map((f) => (
                  <th key={f.id} style={{ padding: "6px 8px", textAlign: "left", fontSize: 10, color: "#7a8794", border: "1px solid #e2e8ef" }}>{f.label}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              <tr>
                {fields.map((f) => (
                  <td key={f.id} style={{ padding: "6px 8px", border: "1px solid #e2e8ef", fontWeight: f.includeInGrandTotal ? 700 : 400 }}>
                    {f.type === "number" ? money(printValue(f)).replace("Rs. ", "") : (printValue(f) ?? "")}
                  </td>
                ))}
              </tr>
            </tbody>
          </table>
        </div>
      </>
    );
  }
  // printAlign: a section that should sit at one side of the page on the generated
  // document rather than the normal two-column grid (e.g. Bill No., top-right on the
  // reference) — only meaningful with one or two fields, so it skips Grid2's layout.
  if (printMode && section.printAlign) {
    return (
      <>
        {showTitle && <SectionTitle n={n} title={titleText} printMode={printMode} />}
        <div style={{ display: "flex", justifyContent: section.printAlign === "right" ? "flex-end" : "flex-start", marginBottom: 14 }}>
          {fields.map((f) => (
            <div key={f.id} style={{ minWidth: 180 }}>
              <Field label={f.label} required={f.required}>
                <SchemaFieldInput def={f} value={printValue(f)} readOnly onChange={() => {}} />
              </Field>
            </div>
          ))}
        </div>
      </>
    );
  }
  return (
    <>
      {showTitle && <SectionTitle n={n} title={titleText} printMode={printMode} />}
      <Grid2>
        {fields.map((f) => {
          const val = printMode ? printValue(f) : (f.readOnly ? flat[f.id] : (values?.[f.id] ?? f.default ?? ""));
          // 24-hour format hint on time fields so entries like "2:30" can't be misread
          // as AM/PM — the input itself is forced to a 24-hour clock (see SchemaFieldInput).
          const hint = f.type === "time" && !printMode ? "24-hour format, e.g. 14:30" : undefined;
          return (
            <Field key={f.id} label={f.label} required={f.required} hint={hint}>
              <SchemaFieldInput def={f} value={val} readOnly={readOnly} onChange={(v) => onFieldChange(f.id, v)} />
            </Field>
          );
        })}
      </Grid2>
    </>
  );
}
// dd-mm-yyyy formatting for a <input type="date"> value (yyyy-mm-dd) — matches the
// reference documents' "date & time" columns, which are written dd-mm-yyyy on the
// paper forms rather than ISO order.
function formatDateDDMMYYYY(iso) {
  if (!iso) return "";
  const parts = String(iso).split("-");
  if (parts.length !== 3) return iso;
  const [y, m, d] = parts;
  return `${d}-${m}-${y}`;
}
function SchemaTableSection({ section, rows, resolvedRows, onRowsChange, readOnly, n, hideExaminerOnly, printMode }) {
  if (hideExaminerOnly && section.hiddenFromExaminer) return null;
  const list = rows && rows.length ? rows : [{}];
  const allCols = sortByOrder(section.columns);
  // In print mode, a pair of columns like "Departure Date" + "Departure Time" (kept
  // separate on-screen so both are individually mandatory) recombines into the single
  // "Departure, date & time" column the reference document actually has — see the
  // `printDisplay` declarations on the travel table's date columns.
  let cols = printMode
    ? allCols.filter((c) => !c.excludeFromPrint).map((c) => (c.printDisplay ? { ...c, label: c.printDisplay.label, _pairWith: c.printDisplay.pairWith } : c))
    : allCols;
  function updateCell(i, colId, val) { onRowsChange(list.map((r, idx) => (idx === i ? { ...r, [colId]: val } : r))); }
  function addRow() { onRowsChange([...list, {}]); }
  function removeRow(i) { onRowsChange(list.filter((_, idx) => idx !== i)); }
  const total = section.totalColumnId ? (resolvedRows || []).reduce((s, r) => s + (parseFloat(r[section.totalColumnId]) || 0), 0) : 0;
  // See the matching comment in SchemaFieldsSection — noTitleInPrint/printTitle let a
  // section's on-screen heading differ from (or be entirely absent from) the generated
  // document, to match how the reference documents actually label — or don't label — this
  // table (e.g. "II.TRAVELLING:-" verbatim, no numbered circle).
  const showTitle = section.title && !(printMode && section.noTitleInPrint);
  const titleText = printMode && section.printTitle ? section.printTitle : section.title;
  // colWidth: an optional per-column width (px) hint — used on-screen only (print
  // layout is unaffected) so a table with several columns doesn't squeeze a numeric
  // column like "Claim Amount" too narrow to read, or a text column like "Remarks"
  // too narrow to type in. Only switches the table into fixed-width mode when at
  // least one column actually declares a width, so tables that don't need it (most
  // of them) keep their existing auto-sized behavior untouched.
  const hasColWidths = !printMode && cols.some((c) => c.colWidth);
  const totalWidth = hasColWidths ? cols.reduce((s, c) => s + (c.colWidth || 90), 0) + (!readOnly ? 30 : 0) : null;
  return (
    <>
      {showTitle && <SectionTitle n={n} title={titleText} printMode={printMode} />}
      <div style={{ overflowX: "auto" }}>
        <table style={{ width: totalWidth ? totalWidth : "100%", tableLayout: hasColWidths ? "fixed" : "auto", borderCollapse: "collapse", fontSize: 12.5 }}>
          {hasColWidths && (
            <colgroup>
              {cols.map((c) => <col key={c.id} style={{ width: (c.colWidth || 90) + "px" }} />)}
              {!readOnly && <col style={{ width: 30 }} />}
            </colgroup>
          )}
          <thead>
            <tr style={{ background: "#f7f9fb" }}>
              {cols.map((c) => (
                <th key={c.id} style={{ padding: "6px 8px", textAlign: "left", fontSize: 10.5, color: "#7a8794", borderBottom: "1px solid #e2e8ef" }}>
                  {c.label} {c.type === "time" && !printMode && <span style={{ fontWeight: 400 }}>(24-hr)</span>} {c.required && <span style={{ color: "#a13a2f" }}>*</span>}
                </th>
              ))}
              {!readOnly && <th style={{ width: 30, borderBottom: "1px solid #e2e8ef" }} />}
            </tr>
          </thead>
          <tbody>
            {list.map((row, i) => (
              <tr key={i} style={{ borderBottom: "1px solid #eef1f5" }}>
                {cols.map((c) => (
                  <td key={c.id} style={{ padding: 4 }}>
                    {c._pairWith ? (
                      <span>{[formatDateDDMMYYYY(row[c.id]), row[c._pairWith]].filter(Boolean).join(" ")}</span>
                    ) : (
                      <SchemaFieldInput
                        def={c}
                        value={c.computed ? (resolvedRows?.[i]?.[c.id] ?? 0) : row[c.id]}
                        readOnly={readOnly || !!c.computed}
                        onChange={(v) => updateCell(i, c.id, v)}
                      />
                    )}
                  </td>
                ))}
                {!readOnly && (
                  <td>
                    <button onClick={() => removeRow(i)} style={{ background: "transparent", border: "none", color: "#a13a2f", cursor: "pointer" }}>
                      <XCircle size={15} />
                    </button>
                  </td>
                )}
              </tr>
            ))}
            {section.totalColumnId && (
              <tr>
                {cols.map((c, i) => (
                  <td key={c.id} style={{ padding: "6px 8px", fontWeight: 700, fontSize: 12 }}>
                    {c.id === section.totalColumnId ? money(total) : (i === 0 ? "Total" : "")}
                  </td>
                ))}
                {!readOnly && <td />}
              </tr>
            )}
          </tbody>
        </table>
      </div>
      {!readOnly && (
        <Btn variant="outline" onClick={addRow} style={{ marginTop: 6, padding: "5px 10px", fontSize: 12 }}>
          + {section.addRowLabel || "Add row"}
        </Btn>
      )}
    </>
  );
}
function SchemaForm({ schema, data, onChange, readOnly, startN = 1, hideExaminerOnly, printMode }) {
  const derived = useMemo(() => deriveForm(schema, data), [schema, data]);
  let counter = startN;
  function sectionFieldChange(sectionId, fieldId, val) {
    onChange({ ...data, [sectionId]: { ...(data[sectionId] || {}), [fieldId]: val } });
  }
  function sectionRowsChange(sectionId, rows) {
    onChange({ ...data, [sectionId]: rows });
  }
  return (
    <div>
      {sortByOrder(schema.sections).map((s) => {
        if (!evalShowIf(s.showIf, derived.flat)) return null;
        if (hideExaminerOnly && s.hiddenFromExaminer) return null;
        if (printMode && s.excludeFromPrint) return null;
        const nn = counter++;
        if (s.kind === "table") {
          return (
            <SchemaTableSection key={s.id} section={s} rows={data[s.id]} resolvedRows={derived.resolvedRowsBySection[s.id]}
              onRowsChange={(rows) => sectionRowsChange(s.id, rows)} readOnly={readOnly} n={nn} hideExaminerOnly={hideExaminerOnly} printMode={printMode} />
          );
        }
        return (
          <SchemaFieldsSection key={s.id} section={s} values={data[s.id]} flat={derived.flat}
            onFieldChange={(fid, v) => sectionFieldChange(s.id, fid, v)} readOnly={readOnly} n={nn} hideExaminerOnly={hideExaminerOnly} printMode={printMode} />
        );
      })}
    </div>
  );
}
function GrandTotalBanner({ grandTotal }) {
  return (
    <div style={{ background: "#1a3a5c", color: "#fff", borderRadius: 10, padding: "14px 18px", display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 8, margin: "16px 0" }}>
      <div>
        <div style={{ fontSize: 11.5, color: "#b9cbe0" }}>GRAND TOTAL (I + II + III + IV)</div>
        <div style={{ fontSize: 22, fontWeight: 800, color: "#e0b86a" }}>{money(grandTotal)}</div>
      </div>
      <div style={{ fontSize: 11.5, color: "#dbe4f2", maxWidth: 420, textAlign: "right" }}>{numToWords(grandTotal)}</div>
    </div>
  );
}
function DocumentLetterhead({ schema, logoUrl }) {
  return (
    <div style={{ textAlign: "center", marginBottom: 6 }}>
      {logoUrl && <img src={logoUrl} alt="Logo" style={{ maxHeight: 60, maxWidth: 160, marginBottom: 6 }} />}
      {schema.meta.preHeaderLines && schema.meta.preHeaderLines.map((line, i) => (
        <div key={i} style={{ fontSize: 10.5, color: "#33455c", lineHeight: 1.4 }}>{line}</div>
      ))}
      <div style={{ fontWeight: 800, fontSize: 15, color: "#1a3a5c", marginTop: schema.meta.preHeaderLines ? 4 : 0 }}>{schema.meta.org1}</div>
      <div style={{ fontWeight: 700, fontSize: 13, margin: "4px 0 2px" }}>{schema.meta.title}</div>
      {schema.meta.subtitle && <div style={{ fontSize: 11.5, color: "#7a8794" }}>{schema.meta.subtitle}</div>}
    </div>
  );
}
function DocumentPrintView({ schema, data, logoUrl, minRemuneration }) {
  const derived = useMemo(() => deriveForm(schema, data), [schema, data]);
  const hasGrandTotal = (schema.sections || []).some((s) => (s.kind === "table" && s.includeInGrandTotal) || (s.kind === "fields" && (s.fields || []).some((f) => f.includeInGrandTotal)));
  const printVars = { ...derived.flat, name: derived.flat.fullName, amount: Math.round(derived.grandTotal || 0).toLocaleString("en-IN") };
  const sb = schema.meta.signatureBlock;
  const notesBlock = schema.meta.notes && schema.meta.notes.length > 0 && (
    <div style={{ marginTop: 14, paddingTop: 12, borderTop: "1px dashed #dde3ea" }}>
      <div style={{ fontSize: 11, fontWeight: 700, color: "#1a3a5c", marginBottom: 6 }}>{schema.meta.notesTitle || "DECLARATION / NOTES"}</div>
      {schema.meta.notes.map((note, i) => <div key={i} style={{ fontSize: 10.5, color: "#1c2b3a", marginBottom: 4 }}>{i + 1}. {interpolate(note, printVars)}</div>)}
    </div>
  );
  const authorityBlock = (
    <>
      {sb && (
        <div style={{ marginTop: 18, fontSize: 10.5, color: "#1c2b3a" }}>
          <div style={{ textAlign: "right", marginBottom: 14 }}>
            {sb.examinerLine}&nbsp;&nbsp;&nbsp;&nbsp;Date: ________________
          </div>
          {sb.officeUse && (
            <div style={{ borderTop: "1px dashed #dde3ea", paddingTop: 12 }}>
              <div style={{ fontWeight: 700, marginBottom: 8 }}>{sb.officeUse.title}</div>
              <div style={{ marginBottom: 8 }}>{interpolate(sb.officeUse.certifiedText, printVars)}</div>
              {/* Left blank deliberately (not auto-filled with the computed Grand Total) —
                  the accounts section scrutinizes the claim and may adjust these figures
                  during verification, so pre-printing a number here would misrepresent
                  that step as already decided. The Grand Total banner above still shows
                  the examiner's own claimed total; this is the separate, later, office
                  determination. */}
              <div>Total Amount claimed: Rs. ________________</div>
              <div style={{ marginBottom: 14 }}>Total Payable Amount: Rs. ________________</div>
              <div style={{ textAlign: "right", fontWeight: 700 }}>{sb.officeUse.authoritySignature}</div>
            </div>
          )}
        </div>
      )}
      {schema.meta.footerAuthority && (
        <div style={{ marginTop: 20, textAlign: "right", fontSize: 10.5, color: "#1c2b3a" }}>
          {schema.meta.footerAuthority.map((line, i) => <div key={i}>{line}</div>)}
        </div>
      )}
      {schema.meta.vehicleSignoff && (() => {
        const vs = schema.meta.vehicleSignoff;
        return (
          <div style={{ marginTop: 18, fontSize: 10.5, color: "#1c2b3a" }}>
            {vs.date && <div style={{ marginBottom: 10 }}>Date: ________________</div>}
            <div style={{ textAlign: "left", marginBottom: 4, fontWeight: 700 }}>{vs.signatureLine}</div>
            {vs.routing && vs.routing.map((line, i) => <div key={i} style={{ textAlign: "right" }}>{line}</div>)}
            <div style={{ borderTop: "1px dashed #dde3ea", marginTop: 16, paddingTop: 12 }}>
              <div style={{ fontWeight: 700, marginBottom: 8 }}>Office Use Only</div>
              <div style={{ marginBottom: 10 }}>
                {vs.deptVehicleLabel}: {derived.flat.v_deptVehicle || ""}&nbsp;&nbsp;________________
              </div>
              <div style={{ marginBottom: 10 }}>{vs.taxiLine1}</div>
              <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 10, fontWeight: 700 }}>
                <span>{vs.approverLeft}</span><span>{vs.approverRight}</span>
              </div>
              <div style={{ textAlign: "center", marginBottom: 14 }}>Approved: ________________&nbsp;&nbsp;&nbsp;Not Approved: ________________</div>
              <div style={{ marginBottom: 10 }}>{vs.taxiLine2}</div>
              <div style={{ textAlign: "right", fontWeight: 700 }}>{vs.finalApprover}</div>
            </div>
          </div>
        );
      })()}
    </>
  );
  return (
    <div style={{ maxWidth: 640, margin: "0 auto", background: "#fff", border: "1px solid #dde3ea", borderRadius: 10, padding: 20 }}>
      <DocumentLetterhead schema={schema} logoUrl={logoUrl} />
      {schema.meta.extraHeaderLine && (
        <div style={{ textAlign: "right", fontSize: 10.5, color: "#33455c", marginBottom: 10 }}>{schema.meta.extraHeaderLine}</div>
      )}
      <SchemaForm schema={schema} data={data} onChange={() => {}} readOnly startN={1} printMode />
      {hasGrandTotal && <GrandTotalBanner grandTotal={derived.grandTotal} />}
      {/* notesPosition: "end" — the Vehicle Requisition reference document has its
          "Important Instructions" AFTER the whole signature/office-use/approval chain,
          unlike Remuneration (declaration before Bank Details/signature) and Flight
          ("* Mandatory fields" before the footer signature block), where notes come
          first. Default (unset) keeps notes before the authority block, matching those. */}
      {schema.meta.notesPosition === "end" ? (
        <>{authorityBlock}{notesBlock}</>
      ) : (
        <>{notesBlock}{authorityBlock}</>
      )}
    </div>
  );
}

/* ============================================================================
   Root App — loads everything from Supabase, holds it in state, and diffs
   writes back the same way the Hall Ticket app does (only changed rows are
   sent, not a blind bulk overwrite).
   ========================================================================== */
export default function App() {
  const [loaded, setLoaded] = useState(false);
  const [settings, setSettings] = useState(defaultSettings());
  const [examiners, setExaminers] = useState([]);
  const [formConfigs, setFormConfigs] = useState(defaultFormConfigs());
  const [submissions, setSubmissions] = useState([]);
  const examinersRef = useRef([]);
  const [view, setView] = useState("landing"); // landing | examiner | admin

  async function refreshAll() {
    try {
      const { data: settingsRow } = await supabase.from("examiner_settings").select("*").eq("id", 1).maybeSingle();
      if (settingsRow) {
        setSettings({
          instituteName: settingsRow.institute_name || defaultSettings().instituteName,
          portalSubtitle: settingsRow.portal_subtitle || defaultSettings().portalSubtitle,
          publicBaseUrl: settingsRow.public_base_url || "",
          logoUrl: settingsRow.logo_url || "",
          minRemuneration: settingsRow.min_remuneration ?? defaultSettings().minRemuneration,
        });
      }
    } catch (e) { console.error("Could not load settings:", e.message); }

    try {
      const { data: configRows } = await supabase.from("examiner_form_configs").select("*");
      if (!configRows || configRows.length === 0) {
        const seed = defaultFormConfigs();
        for (const [formId, schema] of Object.entries(seed)) {
          await supabase.from("examiner_form_configs").upsert({ form_id: formId, schema });
        }
        setFormConfigs(seed);
      } else {
        const built = {};
        configRows.forEach((r) => { built[r.form_id] = r.schema; });
        setFormConfigs({ ...defaultFormConfigs(), ...built });
      }
    } catch (e) { console.error("Could not load form configs:", e.message); }

    try {
      const { data: examinerRows } = await supabase.from("examiners").select("*").order("created_at", { ascending: true });
      const built = (examinerRows || []).map((r) => ({
        id: r.id, email: r.email, password: r.password, name: r.name || "", designation: r.designation || "",
        mobile: r.mobile || "", categoryHint: r.category_hint || "", status: r.status || "invited", createdAt: r.created_at,
      }));
      setExaminers(built);
      examinersRef.current = built;
    } catch (e) { console.error("Could not load examiners:", e.message); }

    try {
      const { data: subRows } = await supabase.from("examiner_submissions").select("*").order("created_at", { ascending: true });
      const built = (subRows || []).map((r) => ({
        id: r.id, examinerId: r.examiner_id, refNo: r.ref_no, status: r.status, data: r.data || {},
        generated: !!r.generated, generatedAt: r.generated_at, createdAt: r.created_at, updatedAt: r.updated_at,
      }));
      setSubmissions(built);
    } catch (e) { console.error("Could not load submissions:", e.message); }
  }

  useEffect(() => {
    (async () => { await refreshAll(); setLoaded(true); })();
    // Refresh on tab focus/regained visibility only — no background timer.
    // A poll every 2 minutes re-downloads every submission's full data even
    // when nothing changed, which adds up in Supabase egress for tabs left
    // open a long time. The admin console's own Refresh button still forces
    // a full reload on demand; examiners typically visit once and leave.
    function onVisible() { if (document.visibilityState === "visible") refreshAll(); }
    window.addEventListener("focus", refreshAll);
    document.addEventListener("visibilitychange", onVisible);
    return () => {
      window.removeEventListener("focus", refreshAll);
      document.removeEventListener("visibilitychange", onVisible);
    };
  }, []);

  async function persistSettings(next) {
    setSettings(next);
    try {
      await supabase.from("examiner_settings").upsert({
        id: 1, institute_name: next.instituteName, portal_subtitle: next.portalSubtitle,
        public_base_url: next.publicBaseUrl, logo_url: next.logoUrl || "",
        min_remuneration: next.minRemuneration ?? defaultSettings().minRemuneration,
      });
    } catch (e) { console.error("Supabase persist(settings) error:", e.message); }
  }

  async function persistFormConfig(formId, schema) {
    setFormConfigs((prev) => ({ ...prev, [formId]: schema }));
    try { await supabase.from("examiner_form_configs").upsert({ form_id: formId, schema }); }
    catch (e) { console.error("Supabase persist(form config) error:", e.message); }
  }
  async function resetFormConfig(formId) {
    const seed = defaultFormConfigs()[formId];
    await persistFormConfig(formId, seed);
  }

  // Examiners: diff-based like the Hall Ticket app's courses/regs persist().
  async function persistExaminers(next) {
    const previous = examinersRef.current;
    setExaminers(next);
    examinersRef.current = next;
    try {
      const prevById = new Map(previous.map((r) => [r.id, r]));
      const nextIds = new Set(next.map((r) => r.id));
      for (const r of previous.filter((r) => !nextIds.has(r.id))) {
        await supabase.from("examiners").delete().eq("id", r.id);
      }
      for (const r of next) {
        if (prevById.get(r.id) === r) continue;
        await supabase.from("examiners").upsert({
          id: r.id, email: r.email.trim().toLowerCase(), password: r.password, name: r.name || "",
          designation: r.designation || "", mobile: r.mobile || "", category_hint: r.categoryHint || "",
          status: r.status || "invited", updated_at: new Date().toISOString(),
        });
      }
    } catch (e) { console.error("Supabase persist(examiners) error:", e.message); }
  }

  async function createSubmission(examinerId, data) {
    const seqNum = submissions.length + 1;
    const record = {
      id: genId("sub"), examinerId, refNo: refNo(seqNum), status: "submitted", data,
      generated: false, generatedAt: null, createdAt: new Date().toISOString(), updatedAt: new Date().toISOString(),
    };
    setSubmissions((prev) => [...prev, record]);
    try {
      await supabase.from("examiner_submissions").insert({
        id: record.id, examiner_id: examinerId, ref_no: record.refNo, status: "submitted", data,
      });
      // Narrow RPC instead of a raw table update — examiners aren't signed in via Supabase
      // Auth, so the examiners table's own UPDATE is admin-only; this function only ever
      // flips one row's status to "submitted", nothing else.
      await supabase.rpc("mark_examiner_submitted", { p_examiner_id: examinerId });
      setExaminers((prev) => prev.map((e) => (e.id === examinerId ? { ...e, status: "submitted" } : e)));
    } catch (e) { console.error("Supabase createSubmission error:", e.message); }
    return record;
  }
  async function updateSubmission(id, patch) {
    setSubmissions((prev) => prev.map((s) => (s.id === id ? { ...s, ...patch, updatedAt: new Date().toISOString() } : s)));
    try {
      const row = {};
      if (patch.data !== undefined) row.data = patch.data;
      if (patch.status !== undefined) row.status = patch.status;
      if (patch.generated !== undefined) row.generated = patch.generated;
      if (patch.generatedAt !== undefined) row.generated_at = patch.generatedAt;
      row.updated_at = new Date().toISOString();
      await supabase.from("examiner_submissions").update(row).eq("id", id);
    } catch (e) { console.error("Supabase updateSubmission error:", e.message); }
  }
  async function deleteSubmission(id) {
    setSubmissions((prev) => prev.filter((s) => s.id !== id));
    try { await supabase.from("examiner_submissions").delete().eq("id", id); }
    catch (e) { console.error("Supabase deleteSubmission error:", e.message); }
  }

  if (!loaded) {
    return (
      <div style={{ display: "flex", alignItems: "center", justifyContent: "center", padding: 60, color: "#7a8794", fontFamily: "system-ui, sans-serif" }}>
        <Loader2 size={18} style={{ marginRight: 8, animation: "spin 1s linear infinite" }} />
        Loading examiner portal...
        <style>{`@keyframes spin { from { transform: rotate(0deg); } to { transform: rotate(360deg); } }`}</style>
      </div>
    );
  }

  return (
    <div className="aiims-app-shell" style={{ fontFamily: "system-ui, -apple-system, Segoe UI, Roboto, sans-serif", background: "#f2f4f7", borderRadius: 12, overflow: "hidden", border: "1px solid #dde3ea", maxWidth: 920, margin: "0 auto" }}>
      <style>{`
        @keyframes spin { from { transform: rotate(0deg); } to { transform: rotate(360deg); } }
        input:focus, select:focus, textarea:focus { outline: 2px solid #1a3a5c33; }
        .grid-2col { display: grid; grid-template-columns: 1fr 1fr; align-items: start; gap: 0 20px; }
        @media (max-width: 640px) { .grid-2col { grid-template-columns: 1fr; } }
        @media print {
          body * { visibility: hidden; }
          .print-area, .print-area * { visibility: visible; }
          .print-area { position: absolute; left: 0; top: 0; width: 100%; }
          .print-page-break { break-after: page; page-break-after: always; }
          * { -webkit-print-color-adjust: exact !important; print-color-adjust: exact !important; }
          @page { size: A4; margin: 8mm; }
        }
      `}</style>
      {view === "landing" && <Landing settings={settings} onPick={setView} />}
      {view === "examiner" && (
        <ExaminerPortal
          examiners={examiners} submissions={submissions} formConfigs={formConfigs} settings={settings}
          createSubmission={createSubmission} updateSubmission={updateSubmission} onExit={() => setView("landing")}
        />
      )}
      {view === "admin" && (
        <AdminPortal
          examiners={examiners} persistExaminers={persistExaminers}
          formConfigs={formConfigs} persistFormConfig={persistFormConfig} resetFormConfig={resetFormConfig}
          submissions={submissions} updateSubmission={updateSubmission} deleteSubmission={deleteSubmission}
          settings={settings} persistSettings={persistSettings}
          onRefresh={refreshAll} onExit={() => setView("landing")}
        />
      )}
    </div>
  );
}

function Landing({ settings, onPick }) {
  return (
    <div>
      <Header title={settings.instituteName} subtitle={settings.portalSubtitle} />
      <div style={{ padding: 40, textAlign: "center" }}>
        <GraduationCap size={40} color="#1a3a5c" />
        <h2 style={{ fontSize: 19, color: "#1c2b3a", margin: "12px 0 4px" }}>Examiner Bill / Requisition Portal</h2>
        <p style={{ fontSize: 13.5, color: "#5f6d7a", maxWidth: 480, margin: "0 auto 28px" }}>
          Examiners submit their remuneration, vehicle and flight requisition details here. The Examination Cell
          reviews each submission and generates the final documents.
        </p>
        <div style={{ display: "flex", gap: 16, justifyContent: "center", flexWrap: "wrap" }}>
          <div onClick={() => onPick("examiner")} role="button" tabIndex={0}
            style={{ cursor: "pointer", width: 220, background: "#fff", border: "1px solid #dde3ea", borderRadius: 10, padding: "22px 18px" }}>
            <GraduationCap size={26} color="#1a3a5c" />
            <div style={{ fontWeight: 700, marginTop: 10, color: "#1c2b3a" }}>Examiner login</div>
            <div style={{ fontSize: 12, color: "#7a8794", marginTop: 4 }}>Fill in your bill & requisition details</div>
          </div>
          <div onClick={() => onPick("admin")} role="button" tabIndex={0}
            style={{ cursor: "pointer", width: 220, background: "#fff", border: "1px solid #dde3ea", borderRadius: 10, padding: "22px 18px" }}>
            <ShieldCheck size={26} color="#a9762f" />
            <div style={{ fontWeight: 700, marginTop: 10, color: "#1c2b3a" }}>Administrator login</div>
            <div style={{ fontSize: 12, color: "#7a8794", marginTop: 4 }}>Configure forms, review & generate documents</div>
          </div>
        </div>
      </div>
    </div>
  );
}

/* ============================================================================
   Examiner-facing portal — email + password issued by the admin (no
   self-registration). Examiners fill in and submit their data; they never
   see a "generate" or print button — only the Examination Cell produces the
   final documents (Admin > Submissions).
   ========================================================================== */
function visibleSectionCount(schema, data, hideExaminerOnly) {
  const { flat } = deriveForm(schema, data);
  return (schema.sections || []).filter((s) => evalShowIf(s.showIf, flat) && !(hideExaminerOnly && s.hiddenFromExaminer)).length;
}
function validateSchemaData(schema, data, hideExaminerOnly) {
  const { flat } = deriveForm(schema, data);
  const missing = [];
  (schema.sections || []).forEach((s) => {
    if (!evalShowIf(s.showIf, flat)) return;
    if (hideExaminerOnly && s.hiddenFromExaminer) return;
    if (s.kind === "table") {
      const requiredCols = (s.columns || []).filter((c) => c.required);
      if (requiredCols.length === 0) return;
      const rows = data?.[s.id] && data[s.id].length ? data[s.id] : [];
      // Only enforce required columns on a row the examiner has actually started filling
      // in (any cell non-empty) — a table with no journeys entered at all is allowed, but
      // a half-filled journey (e.g. a date with no time) is not.
      rows.forEach((row, i) => {
        const touched = (s.columns || []).some((c) => String(row[c.id] || "").trim());
        if (!touched) return;
        requiredCols.forEach((c) => {
          if (!String(row[c.id] || "").trim()) missing.push(`${s.title || s.id} row ${i + 1} — ${c.label}`);
        });
      });
      return;
    }
    if (s.kind !== "fields") return;
    (s.fields || []).forEach((f) => {
      if (hideExaminerOnly && f.hiddenFromExaminer) return;
      if (!f.required || !evalShowIf(f.showIf, flat)) return;
      // Fields with a `default` (e.g. Account Type, Examiner Category) already show that
      // default selected in the UI even before the raw data is touched — validate against
      // the same effective value the field displays, not just the untouched raw value.
      const v = data?.[s.id]?.[f.id] ?? f.default;
      if (!String(v || "").trim()) missing.push(f.label);
    });
  });
  return missing;
}

function ExaminerPortal({ examiners, submissions, formConfigs, settings, createSubmission, updateSubmission, onExit }) {
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [loginErr, setLoginErr] = useState("");
  const [examiner, setExaminer] = useState(null);
  const [formData, setFormData] = useState({ remuneration: {}, vehicle: {}, flight: {} });
  const [existingSub, setExistingSub] = useState(null);
  const [submittedRef, setSubmittedRef] = useState("");
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState("");

  function login() {
    const match = examiners.find((e) => e.email.trim().toLowerCase() === email.trim().toLowerCase() && e.password === password);
    if (!match) { setLoginErr("Incorrect email or password. Contact the Examination Cell if you've lost your credentials."); return; }
    setLoginErr("");
    setExaminer(match);
    const mine = submissions.find((s) => s.examinerId === match.id);
    if (mine) {
      setExistingSub(mine);
      setFormData({
        remuneration: mine.data.remuneration || {},
        vehicle: mine.data.vehicle || {},
        flight: mine.data.flight || {},
      });
    } else {
      setFormData({
        remuneration: { profile: { fullName: match.name || "", mobile: match.mobile || "", email: match.email || "" }, billMeta: { examCategory: match.categoryHint || "MBBS" } },
        vehicle: {}, flight: {},
      });
    }
  }

  useEffect(() => {
    if (!examiner) return;
    const profile = formData.remuneration?.profile || {};
    setFormData((prev) => ({
      ...prev,
      vehicle: autofillFromProfile(formConfigs.vehicle, prev.vehicle, profile),
      flight: autofillFromProfile(formConfigs.flight, prev.flight, profile),
    }));
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [examiner, formData.remuneration?.profile?.fullName, formData.remuneration?.profile?.mobile, formData.remuneration?.profile?.email, formData.remuneration?.profile?.purposeJourney]);

  // Which of the two separate Remuneration Bill documents (UG or PG) applies to this
  // examiner — driven by the category on their submission-in-progress if they've already
  // picked one, falling back to the category their account was created with.
  const currentCategory = formData.remuneration?.billMeta?.examCategory || examiner?.categoryHint || "MBBS";
  const remunFormId = remunerationFormIdFor(currentCategory, formConfigs);
  // Widen the dropdown to every UG+PG category, not just whichever half this document
  // happens to belong to — otherwise an examiner who needs to switch (e.g. from PG/MSC
  // over to BSC Allied & Health Care) has no option to do so from this screen at all.
  const remunSchema = useMemo(() => withAllCategoryOptions(formConfigs[remunFormId], formConfigs), [remunFormId, formConfigs]);
  const remunDerived = useMemo(() => deriveForm(remunSchema, formData.remuneration), [remunSchema, formData.remuneration]);
  const originType = remunDerived.flat.originType;
  const vehicleChoice = remunDerived.flat.vehicleChoice;
  const showVehicle = originType === "Outside Telangana" || (originType === "Within Telangana" && vehicleChoice === "Institute vehicle required");
  const showFlight = originType === "Outside Telangana";

  // The Vehicle form's own schema branches (home<->airport leg / airport<->AIIMS leg / local
  // leg) on a copy of the Travel Origin answer living inside the vehicle data itself, since
  // showIf only evaluates fields within the same schema. Keep it in sync (not just filled once)
  // so switching the answer in Remuneration always updates which vehicle sections apply,
  // including for the admin's edit view and the generated document (which read this the same
  // way, straight from the saved data — no special-casing needed there).
  useEffect(() => {
    if (!examiner) return;
    setFormData((prev) => {
      if (prev.vehicle?.main?.v_originType === originType) return prev;
      return { ...prev, vehicle: { ...prev.vehicle, main: { ...(prev.vehicle?.main || {}), v_originType: originType } } };
    });
  }, [examiner, originType]);

  async function handleSubmit() {
    const missing = [
      ...validateSchemaData(remunSchema, formData.remuneration, true),
      ...(showVehicle ? validateSchemaData(formConfigs.vehicle, formData.vehicle, true) : []),
      ...(showFlight ? validateSchemaData(formConfigs.flight, formData.flight, true) : []),
    ];
    if (missing.length > 0) { setError("Please fill in: " + missing.slice(0, 5).join(", ") + (missing.length > 5 ? "…" : "")); return; }
    setError("");
    setBusy(true);
    const payload = { remuneration: formData.remuneration, vehicle: formData.vehicle, flight: formData.flight };
    if (existingSub) {
      await updateSubmission(existingSub.id, { data: payload });
      setSubmittedRef(existingSub.refNo);
    } else {
      const rec = await createSubmission(examiner.id, payload);
      setSubmittedRef(rec.refNo);
    }
    setBusy(false);
  }

  if (!examiner) {
    return (
      <div>
        <Header title="Examiner login" subtitle={settings.instituteName} onBack={onExit} />
        <div style={{ padding: 40, display: "flex", justifyContent: "center" }}>
          <div style={{ background: "#fff", border: "1px solid #dde3ea", borderRadius: 10, padding: 28, width: 340 }}>
            <KeyRound size={26} color="#1a3a5c" />
            <p style={{ fontSize: 13, color: "#5f6d7a", margin: "10px 0 16px" }}>
              Sign in with the official email ID and password shared with you by the Examination Cell.
            </p>
            <Field label="Email ID" required>
              <input style={inputStyle} value={email} onChange={(e) => setEmail(e.target.value)} onKeyDown={(e) => e.key === "Enter" && login()} />
            </Field>
            <Field label="Password" required>
              <input type="password" style={inputStyle} value={password} onChange={(e) => setPassword(e.target.value)} onKeyDown={(e) => e.key === "Enter" && login()} />
            </Field>
            {loginErr && <div style={{ color: "#a13a2f", fontSize: 12.5, marginBottom: 10 }}>{loginErr}</div>}
            <Btn onClick={login} style={{ width: "100%", justifyContent: "center" }}>Sign in</Btn>
          </div>
        </div>
      </div>
    );
  }

  if (existingSub && existingSub.status === "generated") {
    return (
      <div>
        <Header title="Examiner portal" subtitle={settings.instituteName} onBack={onExit} />
        <div style={{ padding: 40, display: "flex", justifyContent: "center" }}>
          <div style={{ background: "#fff", border: "1px solid #dde3ea", borderRadius: 10, padding: 28, maxWidth: 420, textAlign: "center" }}>
            <CheckCircle2 size={32} color="#2f6b45" />
            <h3 style={{ margin: "12px 0 6px", color: "#1c2b3a" }}>Already processed</h3>
            <p style={{ fontSize: 13, color: "#5f6d7a" }}>
              Your submission (Ref {existingSub.refNo}) has already been reviewed and the final documents have been
              generated by the Examination Cell. Contact them directly for any corrections.
            </p>
          </div>
        </div>
      </div>
    );
  }

  if (submittedRef) {
    return (
      <div>
        <Header title="Examiner portal" subtitle={settings.instituteName} onBack={onExit} />
        <div style={{ padding: 40, display: "flex", justifyContent: "center" }}>
          <div style={{ background: "#fff", border: "1px solid #dde3ea", borderRadius: 10, padding: 28, maxWidth: 420, textAlign: "center" }}>
            <CheckCircle2 size={36} color="#2f6b45" />
            <h3 style={{ margin: "12px 0 4px", color: "#1c2b3a" }}>Submitted</h3>
            <p style={{ fontSize: 13.5, color: "#5f6d7a" }}>Your reference number is</p>
            <div style={{ fontFamily: "monospace", fontSize: 18, fontWeight: 700, color: "#1a3a5c", margin: "6px 0 18px" }}>{submittedRef}</div>
            <p style={{ fontSize: 12.5, color: "#7a8794" }}>
              The Examination Cell will review your details and generate your Remuneration Bill, Vehicle Requisition
              and Flight Booking documents. You don't need to do anything further.
            </p>
          </div>
        </div>
      </div>
    );
  }

  const startN2 = 1 + visibleSectionCount(remunSchema, formData.remuneration, true);
  // For Outside Telangana, Flight (Travel Programme) is asked before Vehicle Requisition —
  // the examiner picks their flight first, then requisitions the airport<->AIIMS vehicle leg
  // around it. Within Telangana never shows Flight, so this only changes the Outside order.
  const flightBeforeVehicle = originType === "Outside Telangana";
  const startNFlight = flightBeforeVehicle ? startN2 : (showVehicle ? startN2 + visibleSectionCount(formConfigs.vehicle, formData.vehicle, true) : startN2);
  const startNVehicle = flightBeforeVehicle ? (showFlight ? startN2 + visibleSectionCount(formConfigs.flight, formData.flight, true) : startN2) : startN2;
  const leg2Needed = formData.vehicle?.legAirportAiims?.v_leg2Needed;
  const vehicleForm = (
    <React.Fragment key="vehicle">
      <SchemaForm schema={formConfigs.vehicle} data={formData.vehicle} onChange={(d) => setFormData((p) => ({ ...p, vehicle: d }))} startN={startNVehicle} hideExaminerOnly />
      {leg2Needed === "No, I will arrange this myself" && (
        <p style={{ fontSize: 12, color: "#7a8794", margin: "-8px 0 14px" }}>
          No institute vehicle will be arranged for the Hyderabad Airport ↔ AIIMS Bibinagar leg — please claim this
          journey as a travelling allowance in the <strong>"II. Travelling"</strong> table above instead.
        </p>
      )}
    </React.Fragment>
  );
  const flightForm = (
    <SchemaForm key="flight" schema={formConfigs.flight} data={formData.flight} onChange={(d) => setFormData((p) => ({ ...p, flight: d }))} startN={startNFlight} hideExaminerOnly />
  );

  return (
    <div>
      <Header title="Examiner portal" subtitle={settings.instituteName} onBack={onExit}
        right={<div style={{ fontSize: 12, color: "#cfe0f2" }}>{examiner.email}</div>} />
      <div style={{ padding: 22 }}>
        <div style={{ background: "#fff", border: "1px solid #dde3ea", borderRadius: 10, padding: 24 }}>
          <h3 style={{ margin: "0 0 4px", color: "#1c2b3a" }}>Bill, vehicle & flight requisition details</h3>
          <p style={{ fontSize: 12.5, color: "#7a8794", margin: "0 0 12px" }}>
            Fill this in once — the Examination Cell will generate your Remuneration Bill, Vehicle Requisition and
            Flight Booking Requisition from what you enter here. Fields marked * are required.
            {existingSub && " You previously submitted this — editing and submitting again will update your entry."}
          </p>
          {error && (
            <div style={{ display: "flex", gap: 8, background: "#fbeceb", border: "1px solid #e3b3ae", color: "#a13a2f", padding: "10px 12px", borderRadius: 6, marginBottom: 16, fontSize: 13 }}>
              <AlertCircle size={16} style={{ flexShrink: 0, marginTop: 1 }} /> <span>{error}</span>
            </div>
          )}

          <SchemaForm schema={remunSchema} data={formData.remuneration} onChange={(d) => setFormData((p) => ({ ...p, remuneration: d }))} startN={1} hideExaminerOnly />
          {originType === "Outside Telangana" && (
            <p style={{ fontSize: 12, color: "#7a8794", margin: "-8px 0 14px" }}>
              Note: AIIMS Bibinagar can only send a vehicle for the Hyderabad Airport ↔ AIIMS Bibinagar leg below.
              For travel between your own home/institute and your departure airport, please claim it as a
              travelling allowance in the <strong>"II. Travelling"</strong> table above instead of here.
            </p>
          )}
          {flightBeforeVehicle ? (
            <>
              {showFlight && flightForm}
              {showVehicle && vehicleForm}
            </>
          ) : (
            <>
              {showVehicle && vehicleForm}
              {showFlight && flightForm}
            </>
          )}
          {originType && !showVehicle && !showFlight && (
            <p style={{ fontSize: 12, color: "#7a8794", marginTop: 6 }}>
              You've indicated you'll use your own vehicle for local travel, so no Vehicle Requisition is needed —
              please claim this journey as a travelling allowance in the <strong>"II. Travelling"</strong> table
              above instead.
            </p>
          )}

          <div style={{ marginTop: 20, display: "flex", justifyContent: "flex-end" }}>
            <Btn onClick={handleSubmit} disabled={busy}>
              {busy ? <Loader2 size={15} style={{ animation: "spin 1s linear infinite" }} /> : <CheckCircle2 size={15} />}
              {existingSub ? "Update submission" : "Submit"}
            </Btn>
          </div>
        </div>
      </div>
    </div>
  );
}

/* ============================================================================
   Admin portal — shared password gate (matches the Hall Ticket app's own
   admin login: a single password stored in settings, checked client-side).
   ========================================================================== */
function AdminPortal({
  examiners, persistExaminers, formConfigs, persistFormConfig, resetFormConfig,
  submissions, updateSubmission, deleteSubmission, settings, persistSettings, onRefresh, onExit,
}) {
  const [session, setSession] = useState(null);
  const [authChecked, setAuthChecked] = useState(false);
  const [email, setEmail] = useState("");
  const [pw, setPw] = useState("");
  const [pwErr, setPwErr] = useState("");
  const [signingIn, setSigningIn] = useState(false);
  const [tab, setTab] = useState("dashboard");
  const [refreshing, setRefreshing] = useState(false);

  useEffect(() => {
    let active = true;
    supabase.auth.getSession().then(({ data }) => {
      if (!active) return;
      setSession(data.session || null);
      setAuthChecked(true);
    });
    const { data: sub } = supabase.auth.onAuthStateChange((_event, next) => setSession(next));
    return () => { active = false; sub.subscription.unsubscribe(); };
  }, []);

  const authed = !!session;

  async function handleRefresh() { setRefreshing(true); await onRefresh(); setRefreshing(false); }
  async function login() {
    if (!email.trim() || !pw) { setPwErr("Enter your email and password."); return; }
    setSigningIn(true); setPwErr("");
    const { error } = await supabase.auth.signInWithPassword({ email: email.trim(), password: pw });
    setSigningIn(false);
    if (error) setPwErr(error.message || "Sign-in failed.");
  }
  async function logout() { await supabase.auth.signOut(); onExit(); }

  if (!authChecked) {
    return (
      <div style={{ display: "flex", alignItems: "center", justifyContent: "center", padding: 60, color: "#7a8794", fontFamily: "system-ui, sans-serif" }}>
        <Loader2 size={18} style={{ marginRight: 8, animation: "spin 1s linear infinite" }} /> Checking sign-in...
      </div>
    );
  }

  if (!authed) {
    return (
      <div>
        <Header title="Administrator login" subtitle={settings.instituteName} onBack={onExit} />
        <div style={{ padding: 40, display: "flex", justifyContent: "center" }}>
          <div style={{ background: "#fff", border: "1px solid #dde3ea", borderRadius: 10, padding: 26, width: 300 }}>
            <ShieldCheck size={26} color="#a9762f" />
            <p style={{ fontSize: 13, color: "#5f6d7a", margin: "10px 0 16px" }}>Sign in with your administrator account to continue.</p>
            <input type="email" autoComplete="username" style={{ ...inputStyle, marginBottom: 8 }} value={email} onChange={(e) => setEmail(e.target.value)} onKeyDown={(e) => e.key === "Enter" && login()} placeholder="Email" />
            <input type="password" autoComplete="current-password" style={inputStyle} value={pw} onChange={(e) => setPw(e.target.value)} onKeyDown={(e) => e.key === "Enter" && login()} placeholder="Password" />
            {pwErr && <div style={{ color: "#a13a2f", fontSize: 12, marginTop: 6 }}>{pwErr}</div>}
            <Btn onClick={login} disabled={signingIn} style={{ marginTop: 14, width: "100%", justifyContent: "center" }}>{signingIn ? "Signing in..." : "Sign in"}</Btn>
            <p style={{ fontSize: 11, color: "#a2adb8", marginTop: 12 }}>Administrator accounts are created in Supabase (Authentication → Users) — ask whoever set up the project if you don't have one yet.</p>
          </div>
        </div>
      </div>
    );
  }

  return (
    <div>
      <Header title="Administrator portal" subtitle={settings.instituteName} onBack={onExit}
        right={<div style={{ display: "flex", gap: 4, flexWrap: "wrap", alignItems: "center" }}>
          <TabBtn active={tab === "dashboard"} onClick={() => setTab("dashboard")}>Dashboard</TabBtn>
          <TabBtn active={tab === "examiners"} onClick={() => setTab("examiners")}>Examiners</TabBtn>
          <TabBtn active={tab === "rates"} onClick={() => setTab("rates")}>Exam Rates</TabBtn>
          <TabBtn active={tab === "builder"} onClick={() => setTab("builder")}>Form Builder</TabBtn>
          <TabBtn active={tab === "submissions"} onClick={() => setTab("submissions")}>Submissions</TabBtn>
          <TabBtn active={tab === "reports"} onClick={() => setTab("reports")}>Reports</TabBtn>
          <TabBtn active={tab === "settings"} onClick={() => setTab("settings")}>Settings</TabBtn>
          <TabBtn active={tab === "backup"} onClick={() => setTab("backup")}>Backup</TabBtn>
          <button onClick={handleRefresh} title="Refresh data from Supabase" style={{ display: "flex", alignItems: "center", gap: 4, background: "rgba(255,255,255,0.08)", border: "none", borderRadius: 6, color: "#cfe0f2", cursor: "pointer", padding: "5px 8px", fontSize: 12 }}>
            <Loader2 size={14} style={refreshing ? { animation: "spin 1s linear infinite" } : {}} /> Refresh
          </button>
          <button onClick={logout} title="Log out" style={{ background: "transparent", border: "none", color: "#cfe0f2", cursor: "pointer", padding: 4 }}><LogOut size={17} /></button>
        </div>} />
      <div style={{ padding: 22 }}>
        {tab === "dashboard" && <AdminDashboard examiners={examiners} submissions={submissions} />}
        {tab === "examiners" && <AdminExaminers examiners={examiners} persistExaminers={persistExaminers} submissions={submissions} formConfigs={formConfigs} persistFormConfig={persistFormConfig} />}
        {tab === "rates" && <AdminExamRates formConfigs={formConfigs} persistFormConfig={persistFormConfig} settings={settings} persistSettings={persistSettings} />}
        {tab === "builder" && <AdminFormBuilder formConfigs={formConfigs} persistFormConfig={persistFormConfig} resetFormConfig={resetFormConfig} />}
        {tab === "submissions" && <AdminSubmissions examiners={examiners} submissions={submissions} formConfigs={formConfigs} updateSubmission={updateSubmission} deleteSubmission={deleteSubmission} logoUrl={settings.logoUrl} minRemuneration={settings.minRemuneration} />}
        {tab === "reports" && <AdminReports examiners={examiners} submissions={submissions} formConfigs={formConfigs} />}
        {tab === "settings" && <AdminSettings settings={settings} persistSettings={persistSettings} />}
        {tab === "backup" && <AdminBackup formConfigs={formConfigs} persistFormConfig={persistFormConfig} settings={settings} persistSettings={persistSettings} />}
      </div>
    </div>
  );
}

function AdminDashboard({ examiners, submissions }) {
  const stats = useMemo(() => {
    const total = examiners.length;
    const submitted = submissions.length;
    const generated = submissions.filter((s) => s.status === "generated").length;
    const pending = submissions.filter((s) => s.status !== "generated").length;
    const byCategory = {};
    submissions.forEach((s) => {
      const cat = s.data?.remuneration?.billMeta?.examCategory || "—";
      byCategory[cat] = (byCategory[cat] || 0) + 1;
    });
    return { total, submitted, generated, pending, byCategory };
  }, [examiners, submissions]);
  const maxCat = Math.max(1, ...Object.values(stats.byCategory));
  return (
    <div>
      <div style={{ display: "flex", gap: 12, flexWrap: "wrap", marginBottom: 20 }}>
        <StatCard label="Examiner accounts" value={stats.total} />
        <StatCard label="Submissions received" value={stats.submitted} color="#8a6116" />
        <StatCard label="Generated" value={stats.generated} color="#2f6b45" />
        <StatCard label="Awaiting generation" value={stats.pending} color="#a13a2f" />
      </div>
      <div style={{ background: "#fff", border: "1px solid #dde3ea", borderRadius: 10, padding: 18 }}>
        <div style={{ fontWeight: 700, fontSize: 13.5, color: "#1c2b3a", marginBottom: 12 }}>Submissions by category</div>
        {Object.keys(stats.byCategory).length === 0 && <p style={{ fontSize: 13, color: "#7a8794" }}>No submissions yet.</p>}
        {Object.entries(stats.byCategory).map(([cat, count]) => (
          <div key={cat} style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 8 }}>
            <div style={{ width: 90, fontSize: 12.5, color: "#5f6d7a" }}>{cat}</div>
            <div style={{ flex: 1, background: "#eef2f6", borderRadius: 4, height: 10 }}>
              <div style={{ width: `${(count / maxCat) * 100}%`, background: "#1a3a5c", height: 10, borderRadius: 4 }} />
            </div>
            <div style={{ width: 24, textAlign: "right", fontSize: 12.5 }}>{count}</div>
          </div>
        ))}
      </div>
    </div>
  );
}

function emptyExaminerDraft(defaultCategory) { return { email: "", name: "", designation: "", mobile: "", categoryHint: defaultCategory || "" }; }

function getCategoryFieldFromSchema(schema) {
  const section = schema.sections.find((s) => s.id === "billMeta");
  return section.fields.find((f) => f.id === "examCategory");
}
// Combined list across both Remuneration Bill documents, for places that need every
// category regardless of UG/PG (the examiner-account dropdown, CSV category matching,
// the Submissions tab's category filter) — read live so admin-added categories show up
// immediately, never hardcoded.
function allCategoryOptions(formConfigs) {
  return [
    ...(getCategoryFieldFromSchema(formConfigs.ugRemuneration).options || []),
    ...(getCategoryFieldFromSchema(formConfigs.pgRemuneration).options || []),
  ];
}
// The UG and PG Remuneration Bills are two separate documents/schemas (item 10's split),
// so each one's own "Examiner Category" dropdown only ever lists its OWN half (UG:
// MBBS/BSC Nursing/BSC Allied & Health Care, or PG: PG/MSC) — there was no way for an
// examiner or admin to cross from one half to the other once landed on a given form,
// since the visible options came straight from whichever schema happened to be showing.
// This patches a (deep-cloned, not shared-state-mutating) copy of the given schema so its
// examCategory field offers the FULL combined list instead — picking a category from the
// other half still works correctly afterward since remunerationFormIdFor/remunFormId are
// recomputed fresh from the picked value on every render; only the dropdown's option list
// needed widening, not the underlying routing logic.
function withAllCategoryOptions(schema, formConfigs) {
  const patched = deepClone(schema);
  getCategoryFieldFromSchema(patched).options = allCategoryOptions(formConfigs);
  return patched;
}

function matchCategory(raw, options) {
  const val = (raw || "").trim().toUpperCase();
  const found = (options || []).find((o) => o.toUpperCase() === val);
  return found || options[0] || "";
}

function AdminExaminers({ examiners, persistExaminers, submissions, formConfigs, persistFormConfig }) {
  const fileRef = useRef(null);
  const categoryOptions = allCategoryOptions(formConfigs);
  const [draft, setDraft] = useState(() => emptyExaminerDraft(categoryOptions[0]));
  const [err, setErr] = useState("");
  const [info, setInfo] = useState("");
  const [lastAdded, setLastAdded] = useState([]);

  function submittedCountFor(id) { return submissions.filter((s) => s.examinerId === id).length; }

  async function addNewCategory() {
    const name = window.prompt("New examiner category (e.g. \"BSc Nursing\"):");
    if (!name || !name.trim()) return;
    const trimmed = name.trim();
    if (categoryOptions.includes(trimmed)) { setDraft((d) => ({ ...d, categoryHint: trimmed })); return; }
    const isUG = window.confirm(
      `Is "${trimmed}" an Undergraduate category (goes on the UG Remuneration Bill, alongside MBBS/Nursing/Allied)?\n\n` +
      `OK = UG Remuneration Bill\nCancel = PG Remuneration Bill (alongside PG/MSC)`
    );
    const formId = isUG ? "ugRemuneration" : "pgRemuneration";
    const field = getCategoryFieldFromSchema(formConfigs[formId]);
    const nextSchema = deepClone(formConfigs[formId]);
    getCategoryFieldFromSchema(nextSchema).options = [...(field.options || []), trimmed];
    const sectionId = "remun_" + slugifyCategory(trimmed);
    if (!nextSchema.sections.some((s) => s.id === sectionId)) {
      nextSchema.sections.push(buildCategorySection(trimmed));
    }
    await persistFormConfig(formId, nextSchema);
    setDraft((d) => ({ ...d, categoryHint: trimmed }));
    setInfo(`Added "${trimmed}" as a category with a default remuneration structure (question papers + evaluation, like MBBS, Rs.3000-15000 claimable range). Fine-tune its rates under Exam Rates, or its fields under Form Builder.`);
  }

  async function addOne() {
    if (!draft.email.trim()) { setErr("Email is required."); return; }
    if (examiners.some((e) => e.email.toLowerCase() === draft.email.trim().toLowerCase())) { setErr("An examiner with this email already exists."); return; }
    const rec = { id: genId("exm"), email: draft.email.trim().toLowerCase(), password: randomPassword(), name: draft.name, designation: draft.designation, mobile: draft.mobile, categoryHint: draft.categoryHint, status: "invited" };
    await persistExaminers([...examiners, rec]);
    setLastAdded([rec]);
    setDraft(emptyExaminerDraft(categoryOptions[0]));
    setErr(""); setInfo("Examiner added — copy their password below, or export credentials as CSV.");
  }

  function resetPassword(id) {
    const next = examiners.map((e) => (e.id === id ? { ...e, password: randomPassword() } : e));
    persistExaminers(next);
    setLastAdded(next.filter((e) => e.id === id));
    setInfo("Password reset — share the new password with the examiner.");
  }

  function removeExaminer(id) {
    const e = examiners.find((x) => x.id === id);
    if (!window.confirm(`Remove ${e.email} from the examiner list? This does not delete any submission they've already made.`)) return;
    persistExaminers(examiners.filter((x) => x.id !== id));
  }

  function handleCsvUpload(e) {
    const file = e.target.files[0];
    if (!file) return;
    setErr(""); setInfo("");
    Papa.parse(file, {
      header: true, skipEmptyLines: true,
      complete: (results) => {
        const rows = results.data;
        const existingEmails = new Set(examiners.map((x) => x.email.toLowerCase()));
        const added = [];
        rows.forEach((r) => {
          const rec = {};
          Object.keys(r).forEach((k) => { rec[k.trim().toLowerCase().replace(/\s+/g, "_")] = (r[k] || "").toString().trim(); });
          const emailVal = (rec.email || rec.email_id || "").toLowerCase();
          if (!emailVal || existingEmails.has(emailVal)) return;
          existingEmails.add(emailVal);
          added.push({
            id: genId("exm"), email: emailVal, password: randomPassword(),
            name: rec.name || "", designation: rec.designation || "", mobile: rec.mobile || "",
            categoryHint: matchCategory(rec.category, categoryOptions), status: "invited",
          });
        });
        if (added.length === 0) { setErr('No new rows found. Make sure the CSV has an "email" column, and rows aren\'t already on the list.'); return; }
        persistExaminers([...examiners, ...added]);
        setLastAdded(added);
        setInfo(`Added ${added.length} examiner(s). Export their credentials below to share with them.`);
      },
      error: (parseErr) => setErr("Could not read that file: " + parseErr.message),
    });
    e.target.value = "";
  }

  function downloadTemplate() {
    download("examiners_template.csv", "email,name,designation,mobile,category\nexample@aiimsbibinagar.edu.in,Dr Example Name,Professor,9876543210,MBBS\n", "text/csv");
  }
  function downloadCredentials(list) {
    const csv = toCSV(list, [
      { label: "email", get: (r) => r.email }, { label: "password", get: (r) => r.password },
      { label: "name", get: (r) => r.name }, { label: "designation", get: (r) => r.designation },
    ]);
    download("examiner_credentials.csv", csv, "text/csv");
  }
  function downloadAllCredentials() { downloadCredentials(examiners); }

  return (
    <div style={{ background: "#fff", border: "1px solid #dde3ea", borderRadius: 10, padding: 20 }}>
      <div style={{ fontWeight: 700, fontSize: 13.5, color: "#1c2b3a", marginBottom: 4 }}>Examiner accounts</div>
      <p style={{ fontSize: 12.5, color: "#7a8794", marginBottom: 16 }}>
        Create an account per examiner (one at a time, or bulk via CSV) — each gets a random password. Share the
        email + password with that examiner yourself (email/print); they use it to log in and submit their details.
      </p>

      <Grid2>
        <Field label="Email"><input style={inputStyle} value={draft.email} onChange={(e) => setDraft((d) => ({ ...d, email: e.target.value }))} /></Field>
        <Field label="Name (optional, prefills the form)"><input style={inputStyle} value={draft.name} onChange={(e) => setDraft((d) => ({ ...d, name: e.target.value }))} /></Field>
        <Field label="Designation"><input style={inputStyle} value={draft.designation} onChange={(e) => setDraft((d) => ({ ...d, designation: e.target.value }))} /></Field>
        <Field label="Mobile"><input style={inputStyle} value={draft.mobile} onChange={(e) => setDraft((d) => ({ ...d, mobile: e.target.value }))} /></Field>
        <Field label="Examiner Category">
          <select
            style={inputStyle}
            value={draft.categoryHint}
            onChange={(e) => {
              if (e.target.value === "__add_new__") { addNewCategory(); return; }
              setDraft((d) => ({ ...d, categoryHint: e.target.value }));
            }}
          >
            {categoryOptions.length === 0 && <option value="">— none configured —</option>}
            {categoryOptions.map((c) => <option key={c} value={c}>{c}</option>)}
            <option value="__add_new__">+ Add new category…</option>
          </select>
        </Field>
      </Grid2>
      <Btn onClick={addOne}>+ Add examiner</Btn>

      <div style={{ display: "flex", alignItems: "center", gap: 10, flexWrap: "wrap", margin: "18px 0 10px" }}>
        <Btn variant="outline" onClick={() => fileRef.current.click()}><Upload size={13} /> Bulk add via CSV</Btn>
        <input ref={fileRef} type="file" accept=".csv,text/csv" onChange={handleCsvUpload} style={{ display: "none" }} />
        <Btn variant="outline" onClick={downloadTemplate}><Download size={13} /> Download CSV template</Btn>
        {examiners.length > 0 && <Btn variant="outline" onClick={downloadAllCredentials}><Download size={13} /> Export all credentials</Btn>}
      </div>
      {err && <div style={{ color: "#a13a2f", fontSize: 12.5, marginBottom: 10 }}>{err}</div>}
      {info && <div style={{ color: "#2f6b45", fontSize: 12.5, marginBottom: 10 }}>{info}</div>}

      {lastAdded.length > 0 && (
        <div style={{ background: "#fdf7ec", border: "1px solid #ecd9ae", borderRadius: 8, padding: 12, marginBottom: 14 }}>
          <div style={{ fontSize: 12, fontWeight: 700, color: "#8a6116", marginBottom: 6 }}>Newly generated credentials — copy these now</div>
          {lastAdded.map((r) => (
            <div key={r.id} style={{ fontSize: 12.5, fontFamily: "monospace", color: "#4a3c1f" }}>{r.email}  /  {r.password}</div>
          ))}
          <Btn variant="outline" onClick={() => downloadCredentials(lastAdded)} style={{ marginTop: 8, padding: "5px 10px", fontSize: 12 }}><Download size={12} /> Download as CSV</Btn>
        </div>
      )}

      <div style={{ maxHeight: 340, overflowY: "auto", border: "1px solid #eef1f5", borderRadius: 6 }}>
        <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
          <thead>
            <tr style={{ background: "#f7f9fb", textAlign: "left" }}>
              <th style={{ padding: "6px 10px" }}>Email</th><th style={{ padding: "6px 10px" }}>Name</th>
              <th style={{ padding: "6px 10px" }}>Password</th>
              <th style={{ padding: "6px 10px" }}>Status</th><th style={{ padding: "6px 10px" }}>Submissions</th>
              <th style={{ padding: "6px 10px" }}>Actions</th>
            </tr>
          </thead>
          <tbody>
            {examiners.map((e) => (
              <tr key={e.id} style={{ borderTop: "1px solid #eef1f5" }}>
                <td style={{ padding: "6px 10px" }}>{e.email}</td>
                <td style={{ padding: "6px 10px" }}>{e.name || "—"}</td>
                <td style={{ padding: "6px 10px", fontFamily: "monospace" }}>{e.password}</td>
                <td style={{ padding: "6px 10px" }}><StatusPill status={e.status} /></td>
                <td style={{ padding: "6px 10px" }}>{submittedCountFor(e.id)}</td>
                <td style={{ padding: "6px 10px", display: "flex", gap: 6 }}>
                  <button onClick={() => resetPassword(e.id)} style={{ fontSize: 11, background: "#eef2f8", border: "none", borderRadius: 4, padding: "3px 8px", cursor: "pointer", color: "#1a3a5c" }}>Reset password</button>
                  <button onClick={() => removeExaminer(e.id)} style={{ fontSize: 11, background: "#fbeceb", border: "none", borderRadius: 4, padding: "3px 8px", cursor: "pointer", color: "#a13a2f" }}>Remove</button>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
        {examiners.length === 0 && <div style={{ padding: 14, fontSize: 12.5, color: "#7a8794" }}>No examiner accounts yet.</div>}
      </div>
    </div>
  );
}

/* ============================================================================
   Admin: Exam Rates — a friendlier, purpose-built control panel over just the
   rate-per-item / minimum / maximum amounts baked into the remuneration
   formulas (Rate per Question Paper, Thesis Evaluation rate, Theory/Practical
   evaluation rate + claimable min/max), instead of hunting for them inside
   Form Builder's generic per-field formula editor. Each row is a dropdown of
   common preset amounts (plus a manual "Custom…" override) that writes
   straight into the same computed.rate / computed.clamp.min / .max the
   remuneration schema already uses — no new field type, no data migration.
   Add more entries to RATE_ITEMS below as new admin-configurable amounts
   come up later.
   ========================================================================== */
const RATE_ITEMS = [
  { formId: "ugRemuneration", section: "remun_MBBS", field: "mb_qpAmount", param: "rate", label: "MBBS — Rate per Question Paper Set", presets: [500, 1200] },
  { formId: "ugRemuneration", section: "remun_MBBS", field: "mb_evalAmount", param: "rate", label: "MBBS — Answer Books / Practical / Viva Evaluation (Rs. per paper or candidate)", presets: [100, 150, 200] },
  { formId: "ugRemuneration", section: "remun_MBBS", field: "mb_evalAmount", param: "clampMin", label: "MBBS — Evaluation Minimum Claimable Amount", presets: [3000, 5000] },
  { formId: "ugRemuneration", section: "remun_MBBS", field: "mb_evalAmount", param: "clampMax", label: "MBBS — Evaluation Maximum Claimable Amount", presets: [15000, 20000] },
  { formId: "ugRemuneration", section: "remun_BSC_NURSING", field: "bsc_nursing_qpAmount", param: "rate", label: "B.Sc. Nursing — Rate per Question Paper Set", presets: [500, 1200] },
  { formId: "ugRemuneration", section: "remun_BSC_NURSING", field: "bsc_nursing_evalAmount", param: "rate", label: "B.Sc. Nursing — Answer Books / Practical / Viva Evaluation (Rs. per paper or candidate)", presets: [100, 150, 200] },
  { formId: "ugRemuneration", section: "remun_BSC_NURSING", field: "bsc_nursing_evalAmount", param: "clampMin", label: "B.Sc. Nursing — Evaluation Minimum Claimable Amount", presets: [3000, 5000] },
  { formId: "ugRemuneration", section: "remun_BSC_NURSING", field: "bsc_nursing_evalAmount", param: "clampMax", label: "B.Sc. Nursing — Evaluation Maximum Claimable Amount", presets: [15000, 20000] },
  { formId: "ugRemuneration", section: "remun_BSC_ALLIED_HEALTH_CARE", field: "bsc_allied_health_care_qpAmount", param: "rate", label: "B.Sc. Allied & Health Care — Rate per Question Paper Set", presets: [500, 1200] },
  { formId: "ugRemuneration", section: "remun_BSC_ALLIED_HEALTH_CARE", field: "bsc_allied_health_care_evalAmount", param: "rate", label: "B.Sc. Allied & Health Care — Answer Books / Practical / Viva Evaluation (Rs. per paper or candidate)", presets: [100, 150, 200] },
  { formId: "ugRemuneration", section: "remun_BSC_ALLIED_HEALTH_CARE", field: "bsc_allied_health_care_evalAmount", param: "clampMin", label: "B.Sc. Allied & Health Care — Evaluation Minimum Claimable Amount", presets: [3000, 5000] },
  { formId: "ugRemuneration", section: "remun_BSC_ALLIED_HEALTH_CARE", field: "bsc_allied_health_care_evalAmount", param: "clampMax", label: "B.Sc. Allied & Health Care — Evaluation Maximum Claimable Amount", presets: [15000, 20000] },
  { formId: "pgRemuneration", section: "remun_MSC", field: "ms_qpAmount", param: "rate", label: "MSC — Rate per Question Paper Set", presets: [500, 1200] },
  { formId: "pgRemuneration", section: "remun_MSC", field: "ms_theoryAmount", param: "rate", label: "MSC — Theory Evaluation Rate (per candidate)", presets: [500, 1000, 1200] },
  { formId: "pgRemuneration", section: "remun_MSC", field: "ms_theoryAmount", param: "clampMin", label: "MSC — Theory Evaluation Minimum Claimable Amount", presets: [3000, 5000] },
  { formId: "pgRemuneration", section: "remun_MSC", field: "ms_theoryAmount", param: "clampMax", label: "MSC — Theory Evaluation Maximum Claimable Amount", presets: [15000, 20000] },
  { formId: "pgRemuneration", section: "remun_MSC", field: "ms_pracAmount", param: "rate", label: "MSC — Practical / Clinical / Viva Rate (per candidate)", presets: [1200, 2000] },
  { formId: "pgRemuneration", section: "remun_MSC", field: "ms_pracAmount", param: "clampMin", label: "MSC — Practical / Clinical / Viva Minimum Claimable Amount", presets: [3000, 5000] },
  { formId: "pgRemuneration", section: "remun_MSC", field: "ms_pracAmount", param: "clampMax", label: "MSC — Practical / Clinical / Viva Maximum Claimable Amount", presets: [15000, 20000] },
  { formId: "pgRemuneration", section: "remun_MSC", field: "ms_thesisAmount", param: "rate", label: "MSC — Thesis Evaluation Rate (per candidate)", presets: [4000, 5000] },
  { formId: "pgRemuneration", section: "remun_PG", field: "pg_qpAmount", param: "rate", label: "PG — Rate per Question Paper Set", presets: [500, 1200] },
  { formId: "pgRemuneration", section: "remun_PG", field: "pg_thesisAmount", param: "rate", label: "PG — Thesis Evaluation Rate (per candidate)", presets: [4000, 5000] },
  { formId: "pgRemuneration", section: "remun_PG", field: "pg_pracAmount", param: "rate", label: "PG — Papers Evaluated + Practical / Clinical / Viva Rate (per paper or candidate)", presets: [1200, 2000] },
  { formId: "pgRemuneration", section: "remun_PG", field: "pg_pracAmount", param: "clampMin", label: "PG — Papers + Practical / Clinical / Viva Minimum Claimable Amount", presets: [3000, 5000] },
  { formId: "pgRemuneration", section: "remun_PG", field: "pg_pracAmount", param: "clampMax", label: "PG — Papers + Practical / Clinical / Viva Maximum Claimable Amount", presets: [15000, 20000] },
];
const RATE_ITEM_CATEGORY_LABELS = {
  MBBS: "MBBS", MSC: "M.Sc. / Allied", PG: "PG (MD/MS/DM/MCh)",
  BSC_NURSING: "B.Sc. Nursing", BSC_ALLIED_HEALTH_CARE: "B.Sc. Allied & Health Care",
};
// NOTE: "rate" is stored in two different places depending on the formula's op —
// "multiply" fields (the common case: qpAmount, thesisAmount, pracAmount, etc.) bake
// the rate as the 2nd element of `computed.inputs` (e.g. inputs: ["ms_thesisCount", 4000]);
// only "sumThenMultiply" fields (e.g. mb_evalAmount) use a separate `computed.rate`.
// Writing to computed.rate on a "multiply" field would apply on TOP of the inputs[1]
// literal (raw *= rate after the inputs already multiplied), silently double-counting —
// so these helpers must branch on op rather than assume one universal shape.
function getRateValue(remunerationSchema, item) {
  const section = remunerationSchema.sections.find((s) => s.id === item.section);
  const field = section?.fields.find((f) => f.id === item.field);
  if (!field?.computed) return undefined;
  if (item.param === "rate") {
    return field.computed.op === "sumThenMultiply" ? field.computed.rate : field.computed.inputs?.[1];
  }
  if (item.param === "clampMin") return field.computed.clamp?.min;
  if (item.param === "clampMax") return field.computed.clamp?.max;
}
// Categories whose rate items are hand-curated above in RATE_ITEMS (with nicer,
// specific labels). Anything else — i.e. a category the admin typed in later via
// addNewCategory(), which only gets an auto-generated section — is discovered
// here instead, so it still shows up in this tab without needing a code change.
const KNOWN_RATE_SECTIONS = new Set(RATE_ITEMS.map((i) => i.section));
function categoryLabelFromSection(section) {
  return (section.title || "").replace(/^I\.\s*Remuneration\s*—\s*/, "") || section.id.replace("remun_", "");
}
function deriveAutoRateItems(remunerationSchema, formId) {
  const items = [];
  (remunerationSchema?.sections || []).forEach((section) => {
    if (!section.id.startsWith("remun_") || KNOWN_RATE_SECTIONS.has(section.id)) return;
    const catLabel = categoryLabelFromSection(section);
    (section.fields || []).forEach((field) => {
      const c = field.computed;
      if (!c) return;
      if (c.op === "multiply" || c.op === "sumThenMultiply") {
        items.push({ formId, section: section.id, field: field.id, param: "rate", label: `${catLabel} — ${field.label} Rate`, presets: [100, 500, 1000, 2000, 4000] });
      }
      if (c.clamp) {
        items.push({ formId, section: section.id, field: field.id, param: "clampMin", label: `${catLabel} — ${field.label} Minimum Claimable Amount`, presets: [3000, 5000, 10000] });
        items.push({ formId, section: section.id, field: field.id, param: "clampMax", label: `${catLabel} — ${field.label} Maximum Claimable Amount`, presets: [15000, 20000, 25000] });
      }
    });
  });
  return items;
}
function setRateValue(remunerationSchema, item, value) {
  const section = remunerationSchema.sections.find((s) => s.id === item.section);
  const field = section.fields.find((f) => f.id === item.field);
  if (item.param === "rate") {
    if (field.computed.op === "sumThenMultiply") { field.computed.rate = value; return; }
    const inputs = (field.computed.inputs || []).slice();
    inputs[1] = value;
    field.computed.inputs = inputs;
    return;
  }
  field.computed.clamp = field.computed.clamp || { min: 0, max: 0, gate: item.field.replace("Amount", "Count") };
  if (item.param === "clampMin") field.computed.clamp.min = value; else field.computed.clamp.max = value;
}
function AdminExamRates({ formConfigs, persistFormConfig, settings, persistSettings }) {
  const [status, setStatus] = useState("");
  const [customFor, setCustomFor] = useState(null);
  const [minDraft, setMinDraft] = useState(settings.minRemuneration ?? 3000);

  async function updateItem(item, value) {
    const next = deepClone(formConfigs[item.formId]);
    setRateValue(next, item, value);
    await persistFormConfig(item.formId, next);
    setCustomFor(null);
    setStatus("Saved — applies to every submission in this category immediately.");
    setTimeout(() => setStatus(""), 2500);
  }
  async function saveMinRemuneration() {
    const v = parseFloat(minDraft);
    if (isNaN(v) || v < 0) return;
    await persistSettings({ ...settings, minRemuneration: v });
    setStatus("Saved.");
    setTimeout(() => setStatus(""), 2500);
  }

  const grouped = {};
  [
    ...RATE_ITEMS,
    ...deriveAutoRateItems(formConfigs.ugRemuneration, "ugRemuneration"),
    ...deriveAutoRateItems(formConfigs.pgRemuneration, "pgRemuneration"),
  ].forEach((item) => {
    const cat = item.section.replace("remun_", "");
    (grouped[cat] = grouped[cat] || []).push(item);
  });

  return (
    <div>
      <p style={{ fontSize: 12.5, color: "#7a8794", marginBottom: 16 }}>
        Standard rates, per-candidate fees, and minimum/maximum claimable amounts used to calculate remuneration —
        organised by examiner category. Changing one here updates the formula for every examiner in that category
        immediately; you don't need to edit each submission individually. Anything not listed here (labels, new
        fields, categories, conditional visibility) is still configured under Form Builder.
      </p>
      <div style={{ background: "#fff", border: "1px solid #dde3ea", borderRadius: 10, padding: 16, marginBottom: 14 }}>
        <div style={{ fontWeight: 700, fontSize: 13, color: "#1c2b3a", marginBottom: 4 }}>Minimum Total Remuneration (all categories)</div>
        <p style={{ fontSize: 11.5, color: "#7a8794", marginBottom: 10 }}>
          If an examiner's computed Grand Total works out to less than this, the Remuneration Bill's "Total Payable
          Amount" (in the Office Use section) is raised to this amount instead — the "Total Amount claimed" figure
          above it always shows the actual computed total, unchanged.
        </p>
        <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
          <span style={{ fontSize: 13 }}>Rs.</span>
          <input type="number" style={{ ...inputStyle, maxWidth: 140 }} value={minDraft} onChange={(e) => setMinDraft(e.target.value)} />
          <Btn onClick={saveMinRemuneration} style={{ padding: "6px 14px", fontSize: 12.5 }}>Save</Btn>
        </div>
      </div>
      {Object.entries(grouped).map(([cat, items]) => (
        <div key={cat} style={{ background: "#fff", border: "1px solid #dde3ea", borderRadius: 10, padding: 16, marginBottom: 14 }}>
          <div style={{ fontWeight: 700, fontSize: 13, color: "#1c2b3a", marginBottom: 12 }}>{RATE_ITEM_CATEGORY_LABELS[cat] || cat}</div>
          {items.map((item) => {
            const key = item.section + "." + item.field + "." + item.param;
            const current = getRateValue(formConfigs[item.formId], item);
            const presetSet = Array.from(new Set([...(item.presets || []), current].filter((v) => v !== undefined))).sort((a, b) => a - b);
            const isCustom = customFor === key || (current !== undefined && !presetSet.includes(current));
            return (
              <div key={key} style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 10, flexWrap: "wrap" }}>
                <div style={{ flex: 1, minWidth: 260, fontSize: 12.5, color: "#33455c" }}>{item.label}</div>
                <select
                  style={{ ...inputStyle, maxWidth: 160 }}
                  value={isCustom ? "__custom__" : String(current)}
                  onChange={(e) => {
                    if (e.target.value === "__custom__") { setCustomFor(key); return; }
                    updateItem(item, parseFloat(e.target.value));
                  }}
                >
                  {presetSet.map((p) => <option key={p} value={p}>Rs. {p}</option>)}
                  <option value="__custom__">Custom…</option>
                </select>
                {isCustom && (
                  <input
                    type="number" style={{ ...inputStyle, maxWidth: 120 }} defaultValue={current} placeholder="Rs."
                    onBlur={(e) => { const v = parseFloat(e.target.value); if (!isNaN(v)) updateItem(item, v); }}
                  />
                )}
              </div>
            );
          })}
        </div>
      ))}
      {status && <div style={{ fontSize: 12.5, color: "#2f6b45" }}>{status}</div>}
    </div>
  );
}

/* ============================================================================
   Admin: Form Builder — edit fields, table columns, and fee formulas/rates
   for each of the three documents. No code required; everything is driven by
   the same schema the SchemaForm renderer above understands.
   ========================================================================== */
const FORM_LABELS = { ugRemuneration: "UG Remuneration Bill", pgRemuneration: "PG Remuneration Bill", vehicle: "Vehicle Requisition", flight: "Flight Booking" };
const TYPE_OPTIONS = ["text", "number", "email", "tel", "date", "time", "select", "textarea"];
const OP_OPTIONS = [
  ["", "No formula (editable field)"], ["multiply", "Multiply inputs"],
  ["sumThenMultiply", "Sum inputs, then multiply by rate"], ["sum", "Sum inputs"],
  ["dateDiffDays", "Days between two dates, inclusive (inputs: from-date field, to-date field)"],
];
function csv(arr) { return (arr || []).join(", "); }
function parseCsvTokens(str) {
  return String(str || "").split(",").map((s) => s.trim()).filter(Boolean).map((tok) => (isNaN(Number(tok)) ? tok : Number(tok)));
}
function parseRateToken(str) {
  str = String(str == null ? "" : str).trim();
  if (str === "") return undefined;
  return isNaN(Number(str)) ? str : Number(str);
}
function deepClone(obj) { return typeof structuredClone === "function" ? structuredClone(obj) : JSON.parse(JSON.stringify(obj)); }

function FieldEditorRow({ field, onMutate, onDelete, isColumn }) {
  const [open, setOpen] = useState(false);
  return (
    <div style={{ border: "1px solid #eef1f5", borderRadius: 6, padding: 8, marginBottom: 6 }}>
      <div style={{ display: "grid", gridTemplateColumns: "1.4fr 1fr 0.9fr 50px 50px 50px 60px 50px", gap: 6, alignItems: "center", fontSize: 12 }}>
        <input style={{ ...inputStyle, padding: "5px 7px", fontSize: 12 }} value={field.label} onChange={(e) => onMutate((f) => { f.label = e.target.value; })} placeholder="Label" />
        <input style={{ ...inputStyle, padding: "5px 7px", fontSize: 12, fontFamily: "monospace" }} value={field.id} onChange={(e) => onMutate((f) => { f.id = e.target.value; })} placeholder="field_id" />
        <select style={{ ...inputStyle, padding: "5px 7px", fontSize: 12 }} value={field.type || "text"} onChange={(e) => onMutate((f) => { f.type = e.target.value; })}>
          {TYPE_OPTIONS.map((t) => <option key={t} value={t}>{t}</option>)}
        </select>
        <label style={{ display: "flex", alignItems: "center", gap: 3, justifyContent: "center" }}>
          <input type="checkbox" checked={!!field.required} onChange={(e) => onMutate((f) => { f.required = e.target.checked; })} /> req
        </label>
        <label style={{ display: "flex", alignItems: "center", gap: 3, justifyContent: "center" }}>
          <input type="checkbox" checked={!!field.readOnly} onChange={(e) => onMutate((f) => { f.readOnly = e.target.checked; })} /> calc
        </label>
        <label style={{ display: "flex", alignItems: "center", gap: 3, justifyContent: "center" }} title="Hidden from examiner — admin fills this in">
          <input type="checkbox" checked={!!field.hiddenFromExaminer} onChange={(e) => onMutate((f) => { f.hiddenFromExaminer = e.target.checked; })} /> hide
        </label>
        <input type="number" style={{ ...inputStyle, padding: "5px 7px", fontSize: 12 }} value={field.order || 0} onChange={(e) => onMutate((f) => { f.order = parseFloat(e.target.value) || 0; })} />
        <button onClick={() => setOpen((v) => !v)} style={{ fontSize: 11, background: "#eef2f8", border: "none", borderRadius: 4, padding: "4px 6px", cursor: "pointer" }}>{open ? "Less" : "More"}</button>
      </div>
      {open && (
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 8, marginTop: 8, background: "#f7f9fb", borderRadius: 6, padding: 8, fontSize: 11.5 }}>
          <div><label style={{ color: "#7a8794" }}>Placeholder</label><input style={{ ...inputStyle, padding: "5px 7px", fontSize: 12 }} value={field.placeholder || ""} onChange={(e) => onMutate((f) => { f.placeholder = e.target.value; })} /></div>
          <div><label style={{ color: "#7a8794" }}>Default value</label><input style={{ ...inputStyle, padding: "5px 7px", fontSize: 12 }} value={field.default ?? ""} onChange={(e) => onMutate((f) => { f.default = e.target.value; })} /></div>
          <div><label style={{ color: "#7a8794" }}>Options (comma-separated)</label><input style={{ ...inputStyle, padding: "5px 7px", fontSize: 12 }} value={csv(field.options)} onChange={(e) => onMutate((f) => { f.options = e.target.value.split(",").map((s) => s.trim()).filter(Boolean); })} /></div>
          {!isColumn && (
            <>
              <div><label style={{ color: "#7a8794" }}>Include in Grand Total</label><div><input type="checkbox" checked={!!field.includeInGrandTotal} onChange={(e) => onMutate((f) => { f.includeInGrandTotal = e.target.checked; })} /></div></div>
              <div><label style={{ color: "#7a8794" }}>Autofill from (profile.fieldId)</label><input style={{ ...inputStyle, padding: "5px 7px", fontSize: 12 }} value={field.autofillFrom || ""} onChange={(e) => onMutate((f) => { f.autofillFrom = e.target.value || undefined; })} /></div>
              <div><label style={{ color: "#7a8794" }}>Show only if field…</label><input style={{ ...inputStyle, padding: "5px 7px", fontSize: 12 }} value={field.showIf?.field || ""} onChange={(e) => onMutate((f) => { f.showIf = { field: e.target.value, equals: f.showIf?.equals || "" }; if (!e.target.value) delete f.showIf; })} /></div>
              <div><label style={{ color: "#7a8794" }}>…equals</label><input style={{ ...inputStyle, padding: "5px 7px", fontSize: 12 }} value={field.showIf?.equals || ""} onChange={(e) => onMutate((f) => { f.showIf = { field: f.showIf?.field || "", equals: e.target.value }; })} /></div>
            </>
          )}
          <div>
            <label style={{ color: "#7a8794" }}>Formula</label>
            <select style={{ ...inputStyle, padding: "5px 7px", fontSize: 12 }} value={field.computed?.op || ""} onChange={(e) => onMutate((f) => { if (!e.target.value) delete f.computed; else f.computed = { op: e.target.value, inputs: f.computed?.inputs || [] }; })}>
              {OP_OPTIONS.map(([v, l]) => <option key={v} value={v}>{l}</option>)}
            </select>
          </div>
          <div><label style={{ color: "#7a8794" }}>Formula inputs (ids/numbers, comma-sep)</label><input style={{ ...inputStyle, padding: "5px 7px", fontSize: 12 }} value={field.computed ? csv(field.computed.inputs) : ""} onChange={(e) => onMutate((f) => { if (f.computed) f.computed.inputs = parseCsvTokens(e.target.value); })} /></div>
          <div><label style={{ color: "#7a8794" }}>Rate (id or number)</label><input style={{ ...inputStyle, padding: "5px 7px", fontSize: 12 }} value={field.computed ? (field.computed.rate ?? "") : ""} onChange={(e) => onMutate((f) => { if (f.computed) f.computed.rate = parseRateToken(e.target.value); })} /></div>
          <div><label style={{ color: "#7a8794" }}>Clamp min</label><input type="number" style={{ ...inputStyle, padding: "5px 7px", fontSize: 12 }} value={field.computed?.clamp?.min ?? ""} onChange={(e) => onMutate((f) => { if (!f.computed) return; f.computed.clamp = f.computed.clamp || { min: 0, max: 0, gate: "" }; f.computed.clamp.min = parseFloat(e.target.value) || 0; })} /></div>
          <div><label style={{ color: "#7a8794" }}>Clamp max</label><input type="number" style={{ ...inputStyle, padding: "5px 7px", fontSize: 12 }} value={field.computed?.clamp?.max ?? ""} onChange={(e) => onMutate((f) => { if (!f.computed) return; f.computed.clamp = f.computed.clamp || { min: 0, max: 0, gate: "" }; f.computed.clamp.max = parseFloat(e.target.value) || 0; })} /></div>
          <div><label style={{ color: "#7a8794" }}>Clamp gate (qty field id)</label><input style={{ ...inputStyle, padding: "5px 7px", fontSize: 12 }} value={field.computed?.clamp?.gate ?? ""} onChange={(e) => onMutate((f) => { if (!f.computed) return; if (!e.target.value) { delete f.computed.clamp; return; } f.computed.clamp = f.computed.clamp || { min: 0, max: 0 }; f.computed.clamp.gate = e.target.value; })} /></div>
          <div><button onClick={onDelete} style={{ fontSize: 11, background: "#fbeceb", color: "#a13a2f", border: "none", borderRadius: 4, padding: "5px 8px", cursor: "pointer" }}>Delete this field</button></div>
        </div>
      )}
    </div>
  );
}

function SectionEditor({ section, onMutateSection, onDeleteSection }) {
  const isTable = section.kind === "table";
  const list = isTable ? (section.columns || []) : (section.fields || []);
  function mutateFieldAt(i, fn) {
    onMutateSection((s) => { fn(isTable ? s.columns[i] : s.fields[i]); });
  }
  function deleteFieldAt(i) {
    onMutateSection((s) => { if (isTable) s.columns.splice(i, 1); else s.fields.splice(i, 1); });
  }
  function addField() {
    const n = list.length + 1;
    onMutateSection((s) => {
      const item = { id: (isTable ? "col_" : "field_") + Date.now().toString(36).slice(-5), label: "New " + (isTable ? "Column" : "Field"), type: "text", order: n * 10 };
      if (isTable) s.columns.push(item); else s.fields.push(item);
    });
  }
  return (
    <div style={{ background: "#fff", border: "1px solid #dde3ea", borderRadius: 10, padding: 16, marginBottom: 14 }}>
      <div style={{ display: "flex", gap: 10, flexWrap: "wrap", alignItems: "flex-end", marginBottom: 10, paddingBottom: 10, borderBottom: "1px dashed #dde3ea" }}>
        <div style={{ flex: 1, minWidth: 200 }}>
          <label style={{ fontSize: 11, color: "#7a8794" }}>Section title</label>
          <input style={inputStyle} value={section.title || ""} onChange={(e) => onMutateSection((s) => { s.title = e.target.value; })} />
        </div>
        <div style={{ width: 90 }}>
          <label style={{ fontSize: 11, color: "#7a8794" }}>Order</label>
          <input type="number" style={inputStyle} value={section.order || 0} onChange={(e) => onMutateSection((s) => { s.order = parseFloat(e.target.value) || 0; })} />
        </div>
        <div style={{ width: 150 }}>
          <label style={{ fontSize: 11, color: "#7a8794" }}>Show only if field…</label>
          <input style={inputStyle} value={section.showIf?.field || ""} onChange={(e) => onMutateSection((s) => { s.showIf = { field: e.target.value, equals: s.showIf?.equals || "" }; if (!e.target.value) delete s.showIf; })} />
        </div>
        <div style={{ width: 120 }}>
          <label style={{ fontSize: 11, color: "#7a8794" }}>…equals</label>
          <input style={inputStyle} value={section.showIf?.equals || ""} onChange={(e) => onMutateSection((s) => { s.showIf = { field: s.showIf?.field || "", equals: e.target.value }; })} />
        </div>
        <label style={{ display: "flex", flexDirection: "column", alignItems: "center", gap: 3, fontSize: 11, color: "#7a8794" }}>
          Hide from examiner
          <input type="checkbox" checked={!!section.hiddenFromExaminer} onChange={(e) => onMutateSection((s) => { s.hiddenFromExaminer = e.target.checked; })} />
        </label>
        <Btn variant="danger" onClick={onDeleteSection}>Delete section</Btn>
      </div>
      {section.hiddenFromExaminer && (
        <div style={{ fontSize: 11.5, color: "#8a6116", background: "#fdf7ec", border: "1px solid #ecd9ae", borderRadius: 6, padding: "6px 10px", marginBottom: 10 }}>
          Hidden from the examiner form — this section (and its fields) will only ever appear in the admin's Submissions view and on the generated documents. Admin fills it in.
        </div>
      )}
      {isTable && (
        <Grid2>
          <Field label="Total column (for footer + grand total)">
            <select style={inputStyle} value={section.totalColumnId || ""} onChange={(e) => onMutateSection((s) => { s.totalColumnId = e.target.value || undefined; })}>
              <option value="">(none)</option>
              {list.map((c) => <option key={c.id} value={c.id}>{c.id}</option>)}
            </select>
          </Field>
          <Field label="Include table total in Grand Total">
            <input type="checkbox" checked={!!section.includeInGrandTotal} onChange={(e) => onMutateSection((s) => { s.includeInGrandTotal = e.target.checked; })} />
          </Field>
        </Grid2>
      )}
      {list.map((f, i) => (
        <FieldEditorRow key={f.id + i} field={f} isColumn={isTable} onMutate={(fn) => mutateFieldAt(i, fn)} onDelete={() => deleteFieldAt(i)} />
      ))}
      <Btn variant="outline" onClick={addField} style={{ marginTop: 6, padding: "6px 12px", fontSize: 12 }}>+ Add {isTable ? "column" : "field"}</Btn>
    </div>
  );
}

// Small reusable "list of text lines" editor — used for meta.notes, meta.footerAuthority,
// and meta.vehicleSignoff.routing, all of which are just arrays of plain strings printed
// one per line on the generated document.
function LineListEditor({ lines, onChange, addLabel }) {
  const list = lines || [];
  function update(i, v) { onChange(list.map((l, idx) => (idx === i ? v : l))); }
  function remove(i) { onChange(list.filter((_, idx) => idx !== i)); }
  function add() { onChange([...list, ""]); }
  return (
    <div>
      {list.map((line, i) => (
        <div key={i} style={{ display: "flex", gap: 6, marginBottom: 6 }}>
          <input style={inputStyle} value={line} onChange={(e) => update(i, e.target.value)} />
          <button onClick={() => remove(i)} style={{ background: "transparent", border: "none", color: "#a13a2f", cursor: "pointer" }}>
            <XCircle size={15} />
          </button>
        </div>
      ))}
      <Btn variant="outline" onClick={add} style={{ padding: "5px 10px", fontSize: 12 }}>+ {addLabel || "Add line"}</Btn>
    </div>
  );
}
// Editor for the printed document's non-field text — declaration/notes wording, the
// examiner signature line and Office Use certification text, the footer signing-authority
// block (Flight), and the Vehicle sign-off block (routing lines, Dept. Vehicle / taxi
// approval labels). All of this used to only be changeable by editing code; exposing it
// here lets an admin update signing-authority names or footer wording without a developer.
function MetaTextEditor({ draft, mutate }) {
  const m = draft.meta;
  return (
    <div style={{ background: "#fff", border: "1px solid #dde3ea", borderRadius: 10, padding: 16, marginBottom: 14 }}>
      <div style={{ fontWeight: 700, fontSize: 13, color: "#1c2b3a", marginBottom: 10 }}>Notes, signature & footer text</div>
      {m.notes && (
        <div style={{ marginBottom: 16 }}>
          <Field label="Notes section heading"><input style={inputStyle} value={m.notesTitle || ""} onChange={(e) => mutate((d) => { d.meta.notesTitle = e.target.value; })} /></Field>
          <div style={{ fontSize: 11, color: "#7a8794", margin: "8px 0 4px" }}>Note lines (numbered automatically) — {"{name}"}, {"{amount}"}, {"{examDates}"}, {"{purposeJourney}"} fill in from the submission.</div>
          <LineListEditor lines={m.notes} onChange={(v) => mutate((d) => { d.meta.notes = v; })} addLabel="Add note" />
        </div>
      )}
      {m.signatureBlock && (
        <div style={{ marginBottom: 16, paddingTop: m.notes ? 12 : 0, borderTop: m.notes ? "1px dashed #eef1f5" : "none" }}>
          <Grid2>
            <Field label="Examiner signature line"><input style={inputStyle} value={m.signatureBlock.examinerLine || ""} onChange={(e) => mutate((d) => { d.meta.signatureBlock.examinerLine = e.target.value; })} /></Field>
            <Field label="Office Use heading"><input style={inputStyle} value={m.signatureBlock.officeUse?.title || ""} onChange={(e) => mutate((d) => { d.meta.signatureBlock.officeUse.title = e.target.value; })} /></Field>
            <Field label="Certifying authority signature"><input style={inputStyle} value={m.signatureBlock.officeUse?.authoritySignature || ""} onChange={(e) => mutate((d) => { d.meta.signatureBlock.officeUse.authoritySignature = e.target.value; })} /></Field>
          </Grid2>
          <Field label="Certification text" hint="{name}, {examDates}, {purposeJourney} fill in from the submission.">
            <textarea style={{ ...inputStyle, minHeight: 50 }} value={m.signatureBlock.officeUse?.certifiedText || ""} onChange={(e) => mutate((d) => { d.meta.signatureBlock.officeUse.certifiedText = e.target.value; })} />
          </Field>
        </div>
      )}
      {m.footerAuthority && (
        <div style={{ marginBottom: 16, paddingTop: 12, borderTop: "1px dashed #eef1f5" }}>
          <div style={{ fontSize: 11, color: "#7a8794", marginBottom: 4 }}>Footer signing-authority lines (right-aligned)</div>
          <LineListEditor lines={m.footerAuthority} onChange={(v) => mutate((d) => { d.meta.footerAuthority = v; })} addLabel="Add line" />
        </div>
      )}
      {typeof m.extraHeaderLine === "string" && (
        <div style={{ marginBottom: 16, paddingTop: 12, borderTop: "1px dashed #eef1f5" }}>
          <Field label={'"Request by & Contact no" line (top-right of the form)'}>
            <input style={inputStyle} value={m.extraHeaderLine || ""} onChange={(e) => mutate((d) => { d.meta.extraHeaderLine = e.target.value; })} />
          </Field>
        </div>
      )}
      {m.vehicleSignoff && (() => {
        // Direct, plainly-labeled editors for the two "who is requesting this" fields
        // (Indenting Officer name/designation + mobile no.) — these live as ordinary
        // schema fields (v_officer/v_mobile on the "main" section) rather than under
        // meta, so they're surfaced here explicitly instead of only via the generic
        // per-field "Default value" box in the section editor further down, which was
        // easy to miss.
        const mainSection = draft.sections.find((s) => s.id === "main");
        const officerField = mainSection?.fields?.find((f) => f.id === "v_officer");
        const mobileField = mainSection?.fields?.find((f) => f.id === "v_mobile");
        return (
          <div style={{ paddingTop: 12, borderTop: "1px dashed #eef1f5" }}>
            {(officerField || mobileField) && (
              <div style={{ marginBottom: 16 }}>
                <div style={{ fontSize: 11, color: "#7a8794", marginBottom: 4 }}>Indenting Officer details (shown pre-filled on every new submission)</div>
                <Grid2>
                  {officerField && (
                    <Field label="Name(s) & Designation of Indenting Officer">
                      <input style={inputStyle} value={officerField.default || ""} onChange={(e) => mutate((d) => { d.sections.find((s) => s.id === "main").fields.find((f) => f.id === "v_officer").default = e.target.value; })} />
                    </Field>
                  )}
                  {mobileField && (
                    <Field label="Mobile / Tel No.">
                      <input style={inputStyle} value={mobileField.default || ""} onChange={(e) => mutate((d) => { d.sections.find((s) => s.id === "main").fields.find((f) => f.id === "v_mobile").default = e.target.value; })} />
                    </Field>
                  )}
                </Grid2>
              </div>
            )}
            <Grid2>
              <Field label="Indenting officer's signature line"><input style={inputStyle} value={m.vehicleSignoff.signatureLine || ""} onChange={(e) => mutate((d) => { d.meta.vehicleSignoff.signatureLine = e.target.value; })} /></Field>
              <Field label="Dept. Vehicle label"><input style={inputStyle} value={m.vehicleSignoff.deptVehicleLabel || ""} onChange={(e) => mutate((d) => { d.meta.vehicleSignoff.deptVehicleLabel = e.target.value; })} /></Field>
            </Grid2>
            <div style={{ fontSize: 11, color: "#7a8794", margin: "8px 0 4px" }}>Routing lines (below the signature, e.g. "Through Controlling Officer/HoD")</div>
            <LineListEditor lines={m.vehicleSignoff.routing} onChange={(v) => mutate((d) => { d.meta.vehicleSignoff.routing = v; })} addLabel="Add routing line" />
            <div style={{ marginTop: 10 }} />
            <Grid2>
              <Field label="Taxi approval line 1"><input style={inputStyle} value={m.vehicleSignoff.taxiLine1 || ""} onChange={(e) => mutate((d) => { d.meta.vehicleSignoff.taxiLine1 = e.target.value; })} /></Field>
              <Field label="Approver (left)"><input style={inputStyle} value={m.vehicleSignoff.approverLeft || ""} onChange={(e) => mutate((d) => { d.meta.vehicleSignoff.approverLeft = e.target.value; })} /></Field>
              <Field label="Approver (right)"><input style={inputStyle} value={m.vehicleSignoff.approverRight || ""} onChange={(e) => mutate((d) => { d.meta.vehicleSignoff.approverRight = e.target.value; })} /></Field>
              <Field label="Taxi approval line 2"><input style={inputStyle} value={m.vehicleSignoff.taxiLine2 || ""} onChange={(e) => mutate((d) => { d.meta.vehicleSignoff.taxiLine2 = e.target.value; })} /></Field>
              <Field label="Final approver"><input style={inputStyle} value={m.vehicleSignoff.finalApprover || ""} onChange={(e) => mutate((d) => { d.meta.vehicleSignoff.finalApprover = e.target.value; })} /></Field>
            </Grid2>
          </div>
        );
      })()}
    </div>
  );
}
function AdminFormBuilder({ formConfigs, persistFormConfig, resetFormConfig }) {
  const [formId, setFormId] = useState("ugRemuneration");
  const [draft, setDraft] = useState(() => deepClone(formConfigs[formId]));
  const [status, setStatus] = useState("");

  function switchForm(id) { setFormId(id); setDraft(deepClone(formConfigs[id])); setStatus(""); }
  function mutate(fn) { setDraft((prev) => { const next = deepClone(prev); fn(next); return next; }); }
  function mutateSection(sectionId, fn) { mutate((d) => { fn(d.sections.find((s) => s.id === sectionId)); }); }
  function deleteSection(sectionId) { mutate((d) => { d.sections = d.sections.filter((s) => s.id !== sectionId); }); }
  function addSection(kind) {
    mutate((d) => {
      d.sections.push({ id: "section_" + Date.now().toString(36).slice(-5), title: "New Section", order: (d.sections.length + 1) * 10, kind, fields: kind === "fields" ? [] : undefined, columns: kind === "table" ? [] : undefined });
    });
  }
  async function save() {
    await persistFormConfig(formId, draft);
    setStatus("Saved — live on the examiner form immediately.");
    setTimeout(() => setStatus(""), 3000);
  }
  async function reset() {
    // Reset wipes the WHOLE form back to the hardcoded factory schema — that includes
    // any category an admin added via "+ Add new category" (Admin > Examiners), since
    // those live as extra sections/options on this same schema, not somewhere separate.
    // Call that out by name here rather than only the generic "customisations will be
    // lost" line, which is easy to click past without realizing it means "your added
    // exam categories disappear too." (This is exactly what happened once already —
    // see the project notes.)
    const extraCategories = (draft.sections || [])
      .filter((s) => s.id.startsWith("remun_") && !KNOWN_RATE_SECTIONS.has(s.id))
      .map((s) => categoryLabelFromSection(s));
    const warning = extraCategories.length
      ? `\n\nThis will PERMANENTLY DELETE these admin-added categories and their rates: ${extraCategories.join(", ")}. They are NOT recoverable afterward — you'd have to re-add them from scratch via Admin > Examiners.`
      : "";
    if (!window.confirm(`Reset "${FORM_LABELS[formId]}" to its default field layout? Your customisations will be lost.${warning}`)) return;
    await resetFormConfig(formId);
    setDraft(deepClone(defaultFormConfigs()[formId]));
    setStatus("Reset to default.");
  }

  return (
    <div>
      <div style={{ display: "flex", gap: 8, marginBottom: 14 }}>
        {Object.keys(FORM_LABELS).map((id) => (
          <TabBtn key={id} active={formId === id} onClick={() => switchForm(id)}>{FORM_LABELS[id]}</TabBtn>
        ))}
      </div>
      <div style={{ background: "#fff", border: "1px solid #dde3ea", borderRadius: 10, padding: 16, marginBottom: 14 }}>
        <div style={{ fontWeight: 700, fontSize: 13, color: "#1c2b3a", marginBottom: 10 }}>Form header / letterhead</div>
        <Grid2>
          <Field label="Organisation name"><input style={inputStyle} value={draft.meta.org1} onChange={(e) => mutate((d) => { d.meta.org1 = e.target.value; })} /></Field>
          <Field label="Form title"><input style={inputStyle} value={draft.meta.title} onChange={(e) => mutate((d) => { d.meta.title = e.target.value; })} /></Field>
          <Field label="Subtitle"><input style={inputStyle} value={draft.meta.subtitle || ""} onChange={(e) => mutate((d) => { d.meta.subtitle = e.target.value; })} /></Field>
        </Grid2>
      </div>

      {/* Notes/signature/footer wording were previously only changeable by editing code —
          this exposes them here so an admin can update signing-authority names, footer
          text, etc. without a developer, the same way field labels/defaults already are. */}
      {(draft.meta.notes || draft.meta.signatureBlock || draft.meta.footerAuthority || draft.meta.vehicleSignoff) && (
        <MetaTextEditor draft={draft} mutate={mutate} />
      )}

      {sortByOrder(draft.sections).map((s) => (
        <SectionEditor key={s.id} section={s} onMutateSection={(fn) => mutateSection(s.id, fn)} onDeleteSection={() => deleteSection(s.id)} />
      ))}

      <div style={{ display: "flex", gap: 8, marginBottom: 20 }}>
        <Btn variant="outline" onClick={() => addSection("fields")}>+ Add fields section</Btn>
        <Btn variant="outline" onClick={() => addSection("table")}>+ Add table section</Btn>
      </div>

      <div style={{ display: "flex", gap: 10, alignItems: "center" }}>
        <Btn onClick={save}><CheckCircle2 size={14} /> Save changes</Btn>
        <Btn variant="danger" onClick={reset}>Reset this form to default</Btn>
        {status && <span style={{ fontSize: 12.5, color: "#2f6b45" }}>{status}</span>}
      </div>
    </div>
  );
}

/* ============================================================================
   Admin: Submissions — review what each examiner entered, adjust anything
   (including the workload counts that feed the rate formulas) before
   generating. Examiners never see this "Generate" step or its output.
   ========================================================================== */
function SubmissionDocuments({ submission, formConfigs, logoUrl, minRemuneration }) {
  const remunCategory = submission.data?.remuneration?.billMeta?.examCategory || "";
  const remunSchema = remunerationSchemaFor(remunCategory, formConfigs);
  return (
    <div className="print-area" style={{ display: "flex", flexDirection: "column", gap: 16 }}>
      <div className="print-page-break"><DocumentPrintView schema={remunSchema} data={submission.data.remuneration || {}} logoUrl={logoUrl} minRemuneration={minRemuneration} /></div>
      <div className="print-page-break"><DocumentPrintView schema={formConfigs.vehicle} data={submission.data.vehicle || {}} logoUrl={logoUrl} /></div>
      <div><DocumentPrintView schema={formConfigs.flight} data={submission.data.flight || {}} logoUrl={logoUrl} /></div>
    </div>
  );
}

function SubmissionRow({ submission, examiner, formConfigs, updateSubmission, deleteSubmission, open, onToggle, logoUrl, minRemuneration }) {
  const [editData, setEditData] = useState(submission.data);
  const [saved, setSaved] = useState(false);
  useEffect(() => { setEditData(submission.data); }, [submission.id, submission.updatedAt]);

  const category = editData?.remuneration?.billMeta?.examCategory || "—";
  // Same widened dropdown as the examiner portal — lets the admin move a submission
  // across the UG/PG boundary from this editor too, not just within one half.
  const remunSchema = useMemo(() => withAllCategoryOptions(remunerationSchemaFor(category, formConfigs), formConfigs), [category, formConfigs]);
  const remunTotal = useMemo(() => deriveForm(remunSchema, editData.remuneration || {}).grandTotal, [editData, remunSchema]);

  async function saveEdits() {
    await updateSubmission(submission.id, { data: editData });
    setSaved(true); setTimeout(() => setSaved(false), 2000);
  }
  async function generate() {
    await updateSubmission(submission.id, { data: editData, status: "generated", generated: true, generatedAt: new Date().toISOString() });
  }
  async function reopenForEditing() {
    if (!window.confirm(`Allow ${examiner?.email || submission.examinerId} to edit and resubmit their details? Their already-generated documents stay as they are until you generate again.`)) return;
    await updateSubmission(submission.id, { status: "reopened" });
  }
  async function removeSubmission() {
    if (!window.confirm(`Delete the submission from ${examiner?.email || submission.examinerId}? This cannot be undone.`)) return;
    await deleteSubmission(submission.id);
  }

  return (
    <div style={{ background: "#fff", border: "1px solid #dde3ea", borderRadius: 10, marginBottom: 10, overflow: "hidden" }}>
      <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", padding: "12px 16px", cursor: "pointer" }} onClick={onToggle}>
        <div>
          <div style={{ fontWeight: 700, fontSize: 13.5, color: "#1c2b3a" }}>{editData?.remuneration?.profile?.fullName || examiner?.email || submission.refNo}</div>
          <div style={{ fontSize: 11.5, color: "#7a8794" }}>{examiner?.email} · {category} · Ref {submission.refNo} · {money(remunTotal)}</div>
        </div>
        <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
          <StatusPill status={submission.status} />
          <ChevronRight size={16} style={{ transform: open ? "rotate(90deg)" : "none", color: "#7a8794" }} />
        </div>
      </div>
      {open && (
        <div style={{ borderTop: "1px solid #eef1f5", padding: 16 }}>
          <div style={{ display: "flex", gap: 8, marginBottom: 14 }}>
            <Btn onClick={saveEdits}>{saved ? "Saved!" : "Save changes"}</Btn>
            <Btn variant="success" onClick={generate}><CheckCircle2 size={14} /> {submission.generated ? "Re-generate" : "Generate documents"}</Btn>
            {submission.status === "generated" && (
              <Btn variant="outline" onClick={reopenForEditing}>Allow examiner to edit</Btn>
            )}
            <Btn variant="danger" onClick={removeSubmission}>Delete</Btn>
          </div>
          {submission.status === "reopened" && (
            <div style={{ fontSize: 11.5, color: "#a15a1a", background: "#fdeee3", border: "1px solid #f0cba0", borderRadius: 6, padding: "6px 10px", marginBottom: 14 }}>
              Reopened — the examiner can now edit and resubmit. Their previously generated documents are unchanged until you click Generate/Re-generate again.
            </div>
          )}

          <div style={{ fontWeight: 700, fontSize: 12.5, color: "#274566", margin: "6px 0" }}>Remuneration Bill</div>
          <SchemaForm schema={remunSchema} data={editData.remuneration || {}} onChange={(d) => setEditData((p) => ({ ...p, remuneration: d }))} startN={1} />
          <div style={{ fontWeight: 700, fontSize: 12.5, color: "#274566", margin: "16px 0 6px" }}>Vehicle Requisition</div>
          <SchemaForm schema={formConfigs.vehicle} data={editData.vehicle || {}} onChange={(d) => setEditData((p) => ({ ...p, vehicle: d }))} startN={1} />
          <div style={{ fontWeight: 700, fontSize: 12.5, color: "#274566", margin: "16px 0 6px" }}>Flight Booking</div>
          <SchemaForm schema={formConfigs.flight} data={editData.flight || {}} onChange={(d) => setEditData((p) => ({ ...p, flight: d }))} startN={1} />

          {submission.status === "generated" && (
            <div style={{ marginTop: 18, paddingTop: 14, borderTop: "1px dashed #dde3ea" }}>
              <Btn variant="outline" onClick={() => window.print()}><Printer size={14} /> Print / Save PDF (all 3 documents)</Btn>
              <div style={{ marginTop: 14 }}>
                <SubmissionDocuments submission={{ ...submission, data: editData }} formConfigs={formConfigs} logoUrl={logoUrl} minRemuneration={minRemuneration} />
              </div>
            </div>
          )}
        </div>
      )}
    </div>
  );
}

function AdminSubmissions({ examiners, submissions, formConfigs, updateSubmission, deleteSubmission, logoUrl, minRemuneration }) {
  const [q, setQ] = useState("");
  const [statusFilter, setStatusFilter] = useState("");
  const [categoryFilter, setCategoryFilter] = useState("");
  const [openId, setOpenId] = useState(null);
  const [bulkCategory, setBulkCategory] = useState("");
  const [bulkPreview, setBulkPreview] = useState(false);

  const examinerById = useMemo(() => new Map(examiners.map((e) => [e.id, e])), [examiners]);
  const categoryOptions = allCategoryOptions(formConfigs);

  const filtered = submissions.filter((s) => {
    const examiner = examinerById.get(s.examinerId);
    const category = s.data?.remuneration?.billMeta?.examCategory || "";
    if (statusFilter && s.status !== statusFilter) return false;
    if (categoryFilter && category !== categoryFilter) return false;
    if (q) {
      const needle = q.toLowerCase();
      const name = (s.data?.remuneration?.profile?.fullName || "").toLowerCase();
      if (!name.includes(needle) && !(examiner?.email || "").toLowerCase().includes(needle) && !s.refNo.toLowerCase().includes(needle)) return false;
    }
    return true;
  }).slice().reverse();

  async function bulkGenerate() {
    const targets = submissions.filter((s) => (s.data?.remuneration?.billMeta?.examCategory || "") === bulkCategory && s.status !== "generated");
    if (targets.length === 0) { window.alert("No ungenerated submissions in this category."); return; }
    if (!window.confirm(`Generate documents for ${targets.length} examiner(s) in ${bulkCategory}?`)) return;
    for (const s of targets) {
      await updateSubmission(s.id, { status: "generated", generated: true, generatedAt: new Date().toISOString() });
    }
    setBulkPreview(true);
  }
  const bulkList = bulkCategory ? submissions.filter((s) => (s.data?.remuneration?.billMeta?.examCategory || "") === bulkCategory && s.status === "generated") : [];

  return (
    <div>
      <div style={{ display: "flex", gap: 10, marginBottom: 16, flexWrap: "wrap" }}>
        <input style={{ ...inputStyle, maxWidth: 220 }} placeholder="Search name, email, ref no." value={q} onChange={(e) => setQ(e.target.value)} />
        <select style={{ ...inputStyle, maxWidth: 180 }} value={statusFilter} onChange={(e) => setStatusFilter(e.target.value)}>
          <option value="">All statuses</option><option value="submitted">Submitted</option><option value="reopened">Reopened for editing</option><option value="generated">Generated</option>
        </select>
        <select style={{ ...inputStyle, maxWidth: 160 }} value={categoryFilter} onChange={(e) => setCategoryFilter(e.target.value)}>
          <option value="">All categories</option>
          {categoryOptions.map((c) => <option key={c} value={c}>{c}</option>)}
        </select>
      </div>

      {filtered.length === 0 && <p style={{ fontSize: 13, color: "#7a8794" }}>No submissions match these filters.</p>}
      {filtered.map((s) => (
        <SubmissionRow key={s.id} submission={s} examiner={examinerById.get(s.examinerId)} formConfigs={formConfigs}
          updateSubmission={updateSubmission} deleteSubmission={deleteSubmission} logoUrl={logoUrl} minRemuneration={minRemuneration}
          open={openId === s.id} onToggle={() => setOpenId(openId === s.id ? null : s.id)} />
      ))}

      <div style={{ background: "#fff", border: "1px solid #dde3ea", borderRadius: 10, padding: 20, marginTop: 20 }}>
        <div style={{ fontWeight: 700, fontSize: 13.5, color: "#1c2b3a", marginBottom: 10 }}>Bulk generate by category</div>
        <div style={{ display: "flex", gap: 10, alignItems: "center", flexWrap: "wrap" }}>
          <select style={{ ...inputStyle, maxWidth: 200 }} value={bulkCategory} onChange={(e) => { setBulkCategory(e.target.value); setBulkPreview(false); }}>
            <option value="">Select category</option>
            {categoryOptions.map((c) => <option key={c} value={c}>{c}</option>)}
          </select>
          <Btn onClick={bulkGenerate} disabled={!bulkCategory}>Generate all ungenerated in this category</Btn>
          {bulkList.length > 0 && <Btn variant="outline" onClick={() => window.print()}><Printer size={14} /> Print all generated ({bulkList.length})</Btn>}
        </div>
        {bulkPreview && bulkList.length > 0 && (
          <div className="print-area" style={{ marginTop: 16, display: "flex", flexDirection: "column", gap: 20 }}>
            {bulkList.map((s) => <SubmissionDocuments key={s.id} submission={s} formConfigs={formConfigs} logoUrl={logoUrl} minRemuneration={minRemuneration} />)}
          </div>
        )}
      </div>
    </div>
  );
}

/* ============================================================================
   Admin: Reports, Settings, Backup
   ========================================================================== */
// Turns every field in a schema (fields sections) and every row of every table section
// into its own CSV column, reading resolved/computed values (not just raw saved data) via
// deriveForm — so a computed figure like Vehicle's "No. of Days Vehicle Required" exports
// correctly, and an admin-added field (a new category, a Form-Builder-added column) shows
// up automatically without this export needing a matching code change. A table section
// (e.g. Flight's Travel Programme, which can have several rows) becomes one column with
// its rows joined by " | " rather than a variable number of columns per submission.
function schemaFieldColumns(schema, dataKey, prefix) {
  if (!schema) return [];
  const cols = [];
  const withPrefix = (label) => (prefix ? `${prefix} - ${label}` : label);
  (schema.sections || []).forEach((s) => {
    if (s.kind === "fields") {
      sortByOrder(s.fields).forEach((f) => {
        cols.push({
          label: withPrefix(f.label),
          get: (r) => {
            const derived = deriveForm(schema, r.data?.[dataKey] || {});
            return derived.flat[f.id] ?? "";
          },
        });
      });
    } else if (s.kind === "table") {
      const orderedCols = sortByOrder(s.columns || []);
      cols.push({
        label: withPrefix(s.title || s.id),
        get: (r) => {
          const rows = r.data?.[dataKey]?.[s.id] || [];
          return rows
            .filter((row) => orderedCols.some((c) => String(row[c.id] || "").trim()))
            .map((row) => orderedCols.map((c) => row[c.id] ?? "").join(" "))
            .join(" | ");
        },
      });
    }
  });
  return cols;
}
function AdminReports({ examiners, submissions, formConfigs }) {
  const examinerById = useMemo(() => new Map(examiners.map((e) => [e.id, e])), [examiners]);
  const rows = submissions.map((s) => {
    const remunCategory = s.data.remuneration?.billMeta?.examCategory || "";
    const derived = deriveForm(remunerationSchemaFor(remunCategory, formConfigs), s.data.remuneration || {});
    const bank = s.data.remuneration?.bank || {};
    return { ...s, examiner: examinerById.get(s.examinerId), grandTotal: derived.grandTotal, bank };
  });
  // Shared identifying columns, then one column set per document type — each type
  // becomes its own sheet in the exported workbook rather than one wide flat table.
  const idCols = [
    { label: "Ref No.", get: (r) => r.refNo },
    { label: "Examiner email", get: (r) => r.examiner?.email || "" },
    { label: "Name", get: (r) => r.data.remuneration?.profile?.fullName || "" },
    { label: "Category", get: (r) => r.data.remuneration?.billMeta?.examCategory || "" },
    { label: "Status", get: (r) => r.status },
    { label: "Submitted at", get: (r) => r.createdAt },
  ];
  const remunerationCols = [
    ...idCols,
    { label: "Mobile", get: (r) => r.data.remuneration?.profile?.mobile || "" },
    { label: "Grand Total (Remuneration)", get: (r) => r.grandTotal },
    { label: "Bank name", get: (r) => r.bank.bankName || "" },
    { label: "Account no.", get: (r) => r.bank.acctNumber || "" },
    { label: "IFSC", get: (r) => r.bank.ifsc || "" },
  ];
  // Vehicle/Flight field lists exported dynamically off the current schema so they stay
  // in sync with whatever Form Builder currently has, without a matching code change.
  const vehicleCols = [...idCols, ...schemaFieldColumns(formConfigs.vehicle, "vehicle")];
  const flightCols = [...idCols, ...schemaFieldColumns(formConfigs.flight, "flight")];
  const reportsList = [
    { name: "All submissions", rows },
    { name: "Generated", rows: rows.filter((r) => r.status === "generated") },
    { name: "Awaiting generation", rows: rows.filter((r) => r.status !== "generated") },
  ];
  function exportWorkbook(rep) {
    downloadWorkbook(rep.name.replace(/\s+/g, "_").toLowerCase() + ".xlsx", [
      { name: "Remuneration", rows: rep.rows, cols: remunerationCols },
      { name: "Vehicle", rows: rep.rows, cols: vehicleCols },
      { name: "Flight Booking", rows: rep.rows, cols: flightCols },
    ]);
  }
  return (
    <div style={{ background: "#fff", border: "1px solid #dde3ea", borderRadius: 10, padding: 20 }}>
      <div style={{ fontWeight: 700, fontSize: 13.5, color: "#1c2b3a", marginBottom: 4 }}>Reports</div>
      <p style={{ fontSize: 12.5, color: "#7a8794", marginBottom: 16 }}>Export an Excel workbook of examiner submissions — separate sheets for Remuneration, Vehicle, and Flight Booking — alongside the individual printable documents from the Submissions tab.</p>
      {reportsList.map((rep) => (
        <div key={rep.name} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "10px 0", borderBottom: "1px solid #eef1f5" }}>
          <div>
            <div style={{ fontSize: 13, fontWeight: 600, color: "#1c2b3a" }}>{rep.name}</div>
            <div style={{ fontSize: 11.5, color: "#7a8794" }}>{rep.rows.length} record(s)</div>
          </div>
          <Btn variant="outline" disabled={rep.rows.length === 0} onClick={() => exportWorkbook(rep)}>
            <FileSpreadsheet size={14} /> Export Excel (.xlsx)
          </Btn>
        </div>
      ))}
    </div>
  );
}

// Public storage bucket for branding assets (the logo). Kept out of the
// database itself — see uploadLogo() below for why.
const LOGO_BUCKET = "branding";
function logoStoragePath(publicUrl) {
  // Recognizes a URL this app itself generated (…/storage/v1/object/public/branding/<path>)
  // so a replaced/removed logo can clean up its old file instead of leaving it orphaned.
  if (!publicUrl) return null;
  const marker = `/object/public/${LOGO_BUCKET}/`;
  const i = publicUrl.indexOf(marker);
  return i === -1 ? null : publicUrl.slice(i + marker.length);
}

function AdminSettings({ settings, persistSettings }) {
  const [draft, setDraft] = useState(settings);
  const [saved, setSaved] = useState(false);
  const [err, setErr] = useState("");
  const [logoUploading, setLogoUploading] = useState(false);
  function set(field, val) { setDraft((d) => ({ ...d, [field]: val })); }
  function save() {
    setErr("");
    persistSettings(draft);
    setSaved(true); setTimeout(() => setSaved(false), 2000);
  }
  async function uploadLogo(file) {
    setErr(""); setLogoUploading(true);
    try {
      // Uploaded as a real file in Supabase Storage, not a base64 blob in the
      // settings row — that row is re-downloaded by every open tab every time
      // it refreshes, so a multi-hundred-KB image embedded in it was being
      // re-sent over and over. A Storage file is served with normal HTTP
      // caching, so browsers only fetch it once. A unique filename per
      // upload (rather than always overwriting the same name) sidesteps any
      // stale-cache issue when the logo is later replaced.
      const ext = (file.name.split(".").pop() || "png").toLowerCase().replace(/[^a-z0-9]/g, "") || "png";
      const path = `logo-${Date.now()}.${ext}`;
      const { error: upErr } = await supabase.storage.from(LOGO_BUCKET).upload(path, file, { upsert: true, contentType: file.type || undefined });
      if (upErr) throw upErr;
      const { data } = supabase.storage.from(LOGO_BUCKET).getPublicUrl(path);
      const oldPath = logoStoragePath(draft.logoUrl);
      set("logoUrl", data.publicUrl);
      if (oldPath) supabase.storage.from(LOGO_BUCKET).remove([oldPath]).catch(() => {});
    } catch (e) {
      setErr("Logo upload failed: " + (e.message || "unknown error") + ". Make sure the \"branding\" storage bucket exists (see egress-reduction.sql).");
    } finally {
      setLogoUploading(false);
    }
  }
  function removeLogo() {
    const oldPath = logoStoragePath(draft.logoUrl);
    if (oldPath) supabase.storage.from(LOGO_BUCKET).remove([oldPath]).catch(() => {});
    set("logoUrl", "");
  }
  return (
    <div style={{ background: "#fff", border: "1px solid #dde3ea", borderRadius: 10, padding: 20 }}>
      <div style={{ fontWeight: 700, fontSize: 13.5, color: "#1c2b3a", marginBottom: 16 }}>Portal settings</div>
      <Field label="Institute name"><input style={inputStyle} value={draft.instituteName} onChange={(e) => set("instituteName", e.target.value)} /></Field>
      <Field label="Portal subtitle"><input style={inputStyle} value={draft.portalSubtitle} onChange={(e) => set("portalSubtitle", e.target.value)} /></Field>
      <Field label="Public site URL" hint="Once hosted on a real domain, enter it here (used for any shareable links in future).">
        <input style={inputStyle} placeholder="https://examiners.aiimsbibinagar.edu.in/" value={draft.publicBaseUrl} onChange={(e) => set("publicBaseUrl", e.target.value)} />
      </Field>
      <Field label="Administrator accounts" hint="Admin sign-in is now a real Supabase account, not a portal setting. Add or remove administrators, and reset passwords, from your Supabase dashboard under Authentication → Users.">
        <div style={{ fontSize: 12.5, color: "#7a8794" }}>Managed in Supabase, not here.</div>
      </Field>
      <Field label="Logo" hint="Shown at the top of the Remuneration, Vehicle Requisition, and Air Tickets reports. Either paste an image URL, or upload a file — uploading stores it in Supabase Storage and replaces whatever's in the URL box.">
        <input style={inputStyle} placeholder="https://.../logo.png" value={draft.logoUrl || ""} onChange={(e) => set("logoUrl", e.target.value)} />
        <div style={{ display: "flex", alignItems: "center", gap: 10, marginTop: 8 }}>
          <input type="file" accept="image/*" disabled={logoUploading} onChange={(e) => {
            const file = e.target.files[0];
            if (file) uploadLogo(file);
            e.target.value = "";
          }} />
          {logoUploading && <span style={{ fontSize: 12, color: "#7a8794" }}>Uploading…</span>}
          {!logoUploading && draft.logoUrl && (
            <>
              <img src={draft.logoUrl} alt="Logo preview" style={{ maxHeight: 40, maxWidth: 120 }} />
              <button type="button" onClick={removeLogo} style={{ fontSize: 11.5, color: "#a13a2f", background: "none", border: "none", cursor: "pointer" }}>Remove</button>
            </>
          )}
        </div>
      </Field>
      {err && <div style={{ color: "#a13a2f", fontSize: 12.5, marginBottom: 8 }}>{err}</div>}
      <Btn onClick={save}><CheckCircle2 size={14} /> {saved ? "Saved!" : "Save settings"}</Btn>
    </div>
  );
}

function AdminBackup({ formConfigs, persistFormConfig, settings, persistSettings }) {
  const fileRef = useRef(null);
  const [err, setErr] = useState("");
  const [info, setInfo] = useState("");
  function exportConfig() {
    const payload = { exportedAt: new Date().toISOString(), formConfigs, settings };
    download("aiims_examiner_portal_backup.json", JSON.stringify(payload, null, 2), "application/json");
  }
  function handleImportFile(e) {
    const file = e.target.files[0];
    if (!file) return;
    setErr(""); setInfo("");
    const reader = new FileReader();
    reader.onload = async () => {
      try {
        const parsed = JSON.parse(reader.result);
        if (!parsed.formConfigs || !parsed.settings) throw new Error("This doesn't look like a portal backup file.");
        if (!window.confirm("Import this configuration? This replaces your current form fields, rates, and settings.")) return;
        for (const [formId, schema] of Object.entries(parsed.formConfigs)) await persistFormConfig(formId, schema);
        await persistSettings(parsed.settings);
        setInfo("Configuration imported.");
      } catch (err2) { setErr("Could not read that file: " + err2.message); }
    };
    reader.readAsText(file);
    e.target.value = "";
  }
  return (
    <div style={{ background: "#fff", border: "1px solid #dde3ea", borderRadius: 10, padding: 20 }}>
      <div style={{ fontWeight: 700, fontSize: 13.5, color: "#1c2b3a", marginBottom: 4 }}>Configuration backup</div>
      <p style={{ fontSize: 12.5, color: "#7a8794", marginBottom: 16 }}>
        Save your form fields, rates, and settings as one file, and load them into a new Supabase project instead of
        rebuilding the Form Builder by hand. This does not include examiner accounts or submissions — export those
        separately (Examiners tab, Reports tab).
      </p>
      <div style={{ display: "flex", gap: 10, flexWrap: "wrap" }}>
        <Btn onClick={exportConfig}><Download size={14} /> Export configuration</Btn>
        <Btn variant="outline" onClick={() => fileRef.current.click()}><Upload size={14} /> Import configuration</Btn>
        <input ref={fileRef} type="file" accept=".json,application/json" onChange={handleImportFile} style={{ display: "none" }} />
      </div>
      {err && <div style={{ color: "#a13a2f", fontSize: 12.5, marginTop: 10 }}>{err}</div>}
      {info && <div style={{ color: "#2f6b45", fontSize: 12.5, marginTop: 10 }}>{info}</div>}
    </div>
  );
}
