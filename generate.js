#!/usr/bin/env node
"use strict";

/*
 * Manual fixed-width file generator.
 * Reads the operations workbook in key/value format:
 *   Column A = field name, Column D = operations input
 * Sheets:
 *   Pledge Invocation, BO-to-BO, IDT (NSDL to CDSL)
 *
 * Usage:
 *   npm install
 *   node generate.js --input "NSDL_Operations_Trade_Input_Template(1).xlsx"
 *   node generate.js --input "completed-operations-template.xlsx" --outdir output
 *   node generate.js --self-test
 */

const fs = require("fs");
const path = require("path");
const XLSX = require("xlsx");

const EXPECTED = {
  pledge: [2, 50, 300, 150, 2],
  bobo:   [2, 50, 600, 150, 2],
  idt:    [2, 50, 600, 150, 2],
};

// Fixed pledgee identifiers used for every Pledge Invocation file.
const PLEDGEE_DP_ID = "IN987654";
const PLEDGEE_CLIENT_ID = "88888888";

function fail(message) {
  throw new Error(message);
}
function str(v) {
  if (v === null || v === undefined) return "";
  if (v instanceof Date) {
    const y = v.getFullYear();
    const m = String(v.getMonth() + 1).padStart(2, "0");
    const d = String(v.getDate()).padStart(2, "0");
    return `${y}${m}${d}`;
  }
  return String(v).trim();
}
function cleanNumeric(v, label) {
  const s = str(v).replace(/,/g, "");
  if (!s) return "";
  if (!/^\d+(\.\d{1,3})?$/.test(s)) {
    fail(`${label}: expected a non-negative number, got "${s}".`);
  }
  return s;
}
function padNum(v, width, label) {
  const s = cleanNumeric(v, label);
  if (!s) fail(`${label} is required.`);
  if (s.includes(".")) fail(`${label} must be a whole-number amount (no decimal point).`);
  if (s.length > width) fail(`${label} exceeds ${width} characters.`);
  return s.padStart(width, "0");
}
function padChar(v, width, label, required = true) {
  const s = str(v);
  if (!s && required) fail(`${label} is required.`);
  if (s.length > width) fail(`${label} exceeds ${width} characters (got ${s.length}, max ${width}).`);
  // Shorter character fields are left-padded with spaces.
  return s.padStart(width, " ");
}
function spaces(n) { return " ".repeat(n); }
function zeros(n) { return "0".repeat(n); }
function date8(v, label, required = true) {
  const s = str(v);
  if (!s && !required) return spaces(8);
  if (!/^\d{8}$/.test(s)) fail(`${label} must be YYYYMMDD (8 digits), got "${s}".`);
  const yyyy = Number(s.slice(0,4)), mm = Number(s.slice(4,6)), dd = Number(s.slice(6,8));
  const dt = new Date(Date.UTC(yyyy, mm - 1, dd));
  if (dt.getUTCFullYear() !== yyyy || dt.getUTCMonth() !== mm - 1 || dt.getUTCDate() !== dd) {
    fail(`${label} is not a valid calendar date: ${s}.`);
  }
  return s;
}
function time4(v) {
  let s = str(v);
  if (/^\d{1,2}:\d{2}$/.test(s)) s = s.replace(":", "");
  if (!/^\d{4}$/.test(s)) fail(`File creation time must be HHMM, e.g. 1430. Got "${s}".`);
  const hh = Number(s.slice(0,2)), mm = Number(s.slice(2,4));
  if (hh > 23 || mm > 59) fail(`Invalid file creation time: ${s}.`);
  return s;
}
function quantity15(v, label) {
  const s = cleanNumeric(v, label);
  if (!s) fail(`${label} is required.`);
  const [whole, frac = ""] = s.split(".");
  if (whole.length > 12) fail(`${label}: whole-number portion exceeds 12 digits.`);
  return whole.padStart(12, "0") + frac.padEnd(3, "0");
}
function amount15(v, label, required = true) {
  const s = str(v);
  if (!s && !required) return zeros(15);
  return padNum(s, 15, label);
}
function requireLength(record, width, label) {
  if (record.length !== width) fail(`${label} has length ${record.length}; expected ${width}.`);
  return record;
}
function readSheetColumns(wb, sheetName) {
  const ws = wb.Sheets[sheetName];
  if (!ws) fail(`Workbook is missing sheet "${sheetName}".`);
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
  const fieldsByColumn = [];
  for (let c = 3; c < Math.max(4, ws["!ref"] ? XLSX.utils.decode_range(ws["!ref"]).e.c + 1 : 4); c++) {
    const values = {};
    for (const row of rows) {
      const label = str(row[0]);
      if (!label || label === "Field to provide" || label.startsWith("Operations Input")) continue;
      values[label.toLowerCase()] = row[c];
    }
    fieldsByColumn.push({ column: c + 1, values });
  }
  return fieldsByColumn.filter(x => Object.values(x.values).some(v => str(v) !== ""));
}
function get(v, label, aliases = []) {
  for (const name of [label, ...aliases]) {
    const key = name.toLowerCase();
    const actual = Object.keys(v).find(k => k.toLowerCase() === key);
    if (actual !== undefined && str(v[actual]) !== "") return v[actual];
  }
  return "";
}
function getRequired(v, label, aliases = []) {
  const value = get(v, label, aliases);
  if (str(value) === "") fail(`${label} is required in the Excel sheet.`);
  return value;
}
function creationDateTime(v) {
  const raw = get(v, "File creation Date & time", ["File creation Date and time", "File Creation Date & time"]);
  if (raw) {
    // Accept YYYYMMDDHHMM or YYYYMMDD HHMM or YYYYMMDD,HHMM.
    const compact = str(raw).replace(/[ ,T:-]/g, "");
    if (/^\d{12}$/.test(compact)) {
      return { date: date8(compact.slice(0,8), "File creation date"), time: time4(compact.slice(8,12)) };
    }
    fail('File creation Date & time must be entered as YYYYMMDDHHMM (example 202610221130).');
  }
  const date = get(v, "File creation Date", ["File Creation Date"]);
  const time = get(v, "File creation Time", ["File Creation Time"]);
  if (date && time) return { date: date8(date, "File creation date"), time: time4(time) };
  // If not provided, use current local date/time as a generated default.
  const now = new Date();
  const yyyy = now.getFullYear();
  const mm = String(now.getMonth()+1).padStart(2,"0");
  const dd = String(now.getDate()).padStart(2,"0");
  const hh = String(now.getHours()).padStart(2,"0");
  const mi = String(now.getMinutes()).padStart(2,"0");
  return { date: `${yyyy}${mm}${dd}`, time: `${hh}${mi}` };
}
function commonHeader(batch, recordType, dpId, clientId, count, dt) {
  return requireLength(
    padNum(batch,7,"Batch Number") +
    padNum(recordType,2,"Header Record Type") +
    spaces(6) +
    padChar(dpId,8,"DP ID") +
    padNum(clientId,8,"Client ID") +
    padNum(count,5,"Total no. of records") +
    date8(dt.date,"File creation date") +
    time4(dt.time) +
    spaces(2),
    50, "Header"
  );
}
function commonTrailer(batch, recordType, kind) {
  // BO-to-BO / IDT: 6 spaces then 9 fields of 15 zeroes.
  if (kind === "pledge") {
    return requireLength(padNum(batch,7,"Batch Number") + padNum(recordType,2,"Trailer Record Type") +
      spaces(6) + zeros(45) + zeros(15) + zeros(15) + zeros(15) + zeros(15) + zeros(15) + zeros(15),
      150, "Pledge trailer");
  }
  return requireLength(padNum(batch,7,"Batch Number") + padNum(recordType,2,"Trailer Record Type") +
    spaces(6) + zeros(135), 150, "Transfer trailer");
}
function makePledge(valuesList) {
  const first = valuesList[0];
  const batch = getRequired(first,"Batch Number");
  const dp = PLEDGEE_DP_ID;
  const client = PLEDGEE_CLIENT_ID;
  const dt = creationDateTime(first);
  const details = valuesList.map((v, idx) => {
    for (const fixed of ["Batch Number"]) {
      if (str(get(v,fixed)) !== str(get(first,fixed))) fail(`${fixed} must be identical across all Pledge Invocation columns.`);
    }
    return requireLength(
      padNum(batch,7,"Batch Number") + "02" + padNum(idx+1,5,"Line Number") +
      "910" + "A" +
      padNum(getRequired(v,"Original SPEED-e Instruction ID"),9,"Original SPEED-e Instruction ID") +
      padChar(getRequired(v,"ISIN"),12,"ISIN") +
      quantity15(getRequired(v,"Quantity"),"Quantity") +
      padChar(getRequired(v,"Pledgor DP ID"),8,"Pledgor DP ID") +
      padNum(getRequired(v,"Pledgor Client ID"),8,"Pledgor Client ID") +
      date8(getRequired(v,"Execution Date"),"Execution Date") +
      amount15(getRequired(v,"Consideration"),"Consideration") +
      "1" + spaces(1+8+2+7+11+16+2+2+1+16+13+127),
      300, `Pledge detail ${idx+1}`
    );
  });
  const header = commonHeader(batch, "01", dp, client, details.length, dt);
  return { kind:"pledge", batch, lines:["{}",header,...details,commonTrailer(batch,"99","pledge"),"{}"] };
}
function makeTransfer(valuesList, kind) {
  const isIdt = kind === "idt";
  const first = valuesList[0];
  const batch = getRequired(first,"Batch Number");
  const dp = getRequired(first,"DP ID", isIdt ? [] : ["Pledgee DP ID"]);
  const client = getRequired(first,"Client ID", isIdt ? [] : ["Pledgee Client ID"]);
  const dt = creationDateTime(first);
  const details = valuesList.map((v, idx) => {
    for (const [fixed, aliases] of [["Batch Number",[]],["DP ID",isIdt ? [] : ["Pledgee DP ID"]],["Client ID",isIdt ? [] : ["Pledgee Client ID"]]]) {
      if (str(get(v,fixed,aliases)) !== str(get(first,fixed,aliases))) fail(`${fixed} must be identical across all ${kind} columns.`);
    }
    const prefix = padNum(batch,7,"Batch Number") + "12" + padNum(idx+1,5,"Line Number") +
      (isIdt ? "925" : "904") + "A" + spaces(7) +
      padChar(getRequired(v,"ISIN"),12,"ISIN") +
      quantity15(getRequired(v,"Quantity"),"Quantity") + spaces(9) +
      date8(getRequired(v,"Execution Date"),"Execution Date");
    let destination;
    if (isIdt) {
      const code = padChar(getRequired(v,"Client Code"),16,"Client Code");
      const other = padChar(getRequired(v,"Other Depository ID"),8,"Other Depository ID");
      const codeValue = str(getRequired(v,"Client Code"));
      const otherValue = str(getRequired(v,"Other Depository ID"));
      if (codeValue.length >= 8 && /^\d{8}/.test(codeValue) && codeValue.slice(0,8) !== otherValue) {
        console.warn(`WARNING: IDT column ${idx+4}: Other Depository ID does not equal the first 8 characters of Client Code.`);
      }
      destination = code + other;
    } else {
      destination = padChar(getRequired(v,"Destination DP ID"),8,"Destination DP ID") +
        padNum(getRequired(v,"Destination Client ID"),8,"Destination Client ID") + spaces(8);
    }
    const internalRef = padChar(get(v,"Internal Reference / Transaction ID",["Internal Ref / Remarks"]),20,"Internal Reference / Transaction ID",false);
    const consideration = amount15(get(v,"Consideration"),"Consideration",false);
    const reasonCode = "01";
    const mid = spaces(1+22+13+7);
    const paymentMode = "03";
    const paymentDate = date8(getRequired(v,"Payment Date"),"Payment Date");
    const name = padChar(getRequired(v,"Transferee Name"),135,"Transferee Name");
    const bankAccount = padChar(getRequired(v,"Transferee Bank Account"),30,"Transferee Bank Account");
    const bankName = padChar(getRequired(v,"Transferee Bank Name"),35,"Transferee Bank Name");
    const paymentRef = padChar(getRequired(v,"Payment Reference / UTR"),25,"Payment Reference / UTR");
    const suffix = spaces(1) + "1" + "N" + spaces(10+1+11+16+2+2+147);
    return requireLength(prefix + destination + internalRef + consideration + reasonCode + mid +
      paymentMode + paymentDate + name + bankAccount + bankName + paymentRef + suffix, 600, `${kind} detail ${idx+1}`);
  });
  const header = commonHeader(batch,"11",dp,client,details.length,dt);
  return { kind, batch, lines:["{}",header,...details,commonTrailer(batch,"20",kind),"{}"] };
}
function validateLines(result) {
  const width = result.kind === "pledge" ? 300 : 600;
  const expected = [2,50,...Array(result.detailCount || result.lines.length-4).fill(width),150,2];
  const lengths = result.lines.map(x=>x.length);
  if (lengths.length !== expected.length || lengths.some((n,i)=>n!==expected[i])) {
    fail(`${result.kind}: record lengths ${JSON.stringify(lengths)} do not match expected structure ${JSON.stringify(expected)}.`);
  }
  if (result.lines[0] !== "{}" || result.lines[result.lines.length-1] !== "{}") fail("Missing opening/closing {} markers.");
  const batch = result.lines[1].slice(0,7);
  for (let i=2;i<result.lines.length-2;i++) {
    if (result.lines[i].slice(0,7) !== batch) fail(`Batch number mismatch in detail record ${i-1}.`);
    const expectedLine = String(i-1).padStart(5,"0");
    if (result.lines[i].slice(9,14) !== expectedLine) fail(`Detail line number is not sequential at record ${i-1}.`);
  }
  if (result.lines[result.lines.length-2].slice(0,7) !== batch) fail("Batch number mismatch between header and trailer.");
  const count = Number(result.lines[1].slice(31,36));
  const actualDetails = result.lines.length - 4;
  if (count !== actualDetails) fail(`Header count is ${count}, but there are ${actualDetails} detail records.`);
  return { lengths, batch, detailCount: count };
}
function parseArgs() {
  const args = process.argv.slice(2);
  const opts = { input: "", outdir: "output", selfTest: false };
  for (let i=0;i<args.length;i++) {
    if (args[i] === "--input" || args[i] === "-i") opts.input = args[++i];
    else if (args[i] === "--outdir" || args[i] === "-o") opts.outdir = args[++i];
    else if (args[i] === "--self-test") opts.selfTest = true;
    else if (args[i] === "--help" || args[i] === "-h") opts.help = true;
    else fail(`Unknown argument: ${args[i]}`);
  }
  return opts;
}
function selfTest() {
  const base = {
    "Batch Number":"0000002", "DP ID":"IN987654", "Client ID":"88888888", "Total no. of records":"1",
    "File creation Date & time":"202610221130", "ISIN":"INE075T01014", "Quantity":"1000",
    "Execution Date":"20261019", "Client Code":"1208870416038700", "Other Depository ID":"12088704",
    "Internal Reference / Transaction ID":"O26101900009162738", "Consideration":"500", "Payment Date":"20261021",
    "Transferee Name":"Chaitanya jalan", "Transferee Bank Account":"50100126857449",
    "Transferee Bank Name":"HDFC BANK", "Payment Reference / UTR":"O26101900009162738",
    "Destination DP ID":"IN300126", "Destination Client ID":"11340629",
    "Original SPEED-e Instruction ID":"123456789", "Pledgee DP ID":"IN987654", "Pledgee Client ID":"88888888",
    "Pledgor DP ID":"IN526518", "Pledgor Client ID":"87329372"
  };
  const results = [
    makePledge([{...base, "ISIN":"INE808K01017", "Quantity":"1000", "Execution Date":"20261021", "Consideration":"555000"}]),
    makeTransfer([base],"bobo"),
    makeTransfer([base],"idt")
  ];
  for (const r of results) {
    r.detailCount = r.lines.length - 4;
    const v=validateLines(r);
    console.log(`PASS ${r.kind}: lengths ${v.lengths.join(", ")}; batch ${v.batch}`);
  }
}
function main() {
  const opts = parseArgs();
  if (opts.help) {
    console.log("Usage: node generate.js --input <workbook.xlsx> [--outdir output] | node generate.js --self-test");
    return;
  }
  if (opts.selfTest) return selfTest();
  if (!opts.input) fail("Please provide --input <completed workbook.xlsx>, or run --self-test.");
  const inputPath = path.resolve(opts.input);
  if (!fs.existsSync(inputPath)) fail(`Input workbook not found: ${inputPath}`);
  const outdir = path.resolve(opts.outdir);
  fs.mkdirSync(outdir,{recursive:true});
  const wb = XLSX.readFile(inputPath, {cellDates:true});
  const configs = [
    {sheet:"Pledge Invocation", kind:"pledge", file:"pledge-invocation"},
    {sheet:"BO-to-BO", kind:"bobo", file:"bo-to-bo"},
    {sheet:"IDT (NSDL to CDSL)", kind:"idt", file:"idt"},
  ];
  let ok=0;
  for (const cfg of configs) {
    const columns = readSheetColumns(wb,cfg.sheet);
    if (!columns.length) {
      console.log(`SKIP ${cfg.sheet}: no trade values entered.`);
      continue;
    }
    const valuesList = columns.map(x=>x.values);
    // Each populated input column represents one trade/detail record.
    const result = cfg.kind==="pledge" ? makePledge(valuesList) : makeTransfer(valuesList,cfg.kind);
    result.detailCount = valuesList.length;
    const check=validateLines(result);
    const outPath=path.join(outdir,`${cfg.file}-${result.batch}.txt`);
    fs.writeFileSync(outPath,result.lines.join("\r\n")+"\r\n","ascii");
    console.log(`PASS ${cfg.sheet}: wrote ${outPath}`);
    console.log(`     record lengths: ${check.lengths.join(", ")}; batch ${check.batch}; details ${check.detailCount}`);
    ok++;
  }
  if (!ok) fail("No files generated. Fill at least one trade column on a transaction sheet.");
}
try { main(); }
catch (e) { console.error(`ERROR: ${e.message}`); process.exit(1); }
