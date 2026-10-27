#!/usr/bin/env node
"use strict";

// Compare two fixed-width text files without trimming meaningful spaces.
// Usage:
//   node compare.js --expected "reference.txt" --actual "generated.txt"
//   node compare.js --expected "reference.txt" --actual "generated.txt" --outdir reports

const fs = require("fs");
const path = require("path");

function arg(name, fallback = null) {
  const i = process.argv.indexOf(`--${name}`);
  return i >= 0 ? (process.argv[i + 1] || fallback) : fallback;
}
function usage() {
  console.log(`
Fixed-Width File Comparator
--------------------------------
Required:
  --expected <file>   Reference/expected file
  --actual   <file>   Generated/system file

Optional:
  --outdir <folder>   Report folder (default: comparison_reports)

Example:
  node compare.js --expected "manual.txt" --actual "system.txt" --outdir reports

Notes:
  - Compares records and every character position, including spaces.
  - CRLF and LF line endings are treated as equivalent.
  - A final newline is not treated as an extra record.
  - Does not alter either input file.
`);
}
function readRecords(file) {
  if (!fs.existsSync(file)) throw new Error(`File not found: ${file}`);
  const raw = fs.readFileSync(file, "utf8").replace(/^\uFEFF/, "");
  // Normalize line endings only. Do NOT trim records; spaces are significant.
  const normalized = raw.replace(/\r\n/g, "\n").replace(/\r/g, "\n");
  const records = normalized.split("\n");
  if (records.length && records[records.length - 1] === "") records.pop();
  return records;
}
function visible(s) {
  return JSON.stringify(s)
    .replace(/\\u0000/g, "\\0");
}
function csvCell(v) {
  const s = String(v ?? "");
  return `"${s.replace(/"/g, '""')}"`;
}
function compare(expected, actual) {
  const rows = [];
  const maxRecords = Math.max(expected.length, actual.length);
  let differentRecords = 0;
  let totalCharDiffs = 0;
  let lengthMismatchCount = 0;

  for (let r = 0; r < maxRecords; r++) {
    const e = expected[r];
    const a = actual[r];
    if (e === undefined || a === undefined) {
      differentRecords++;
      rows.push({
        record: r + 1,
        field: "Record",
        position: "",
        expectedLength: e === undefined ? "" : e.length,
        actualLength: a === undefined ? "" : a.length,
        expected: e === undefined ? "<MISSING RECORD>" : e,
        actual: a === undefined ? "<MISSING RECORD>" : a,
        issue: e === undefined ? "Record missing from expected file" : "Record missing from actual file"
      });
      continue;
    }

    let recordHasDiff = false;
    if (e.length !== a.length) {
      recordHasDiff = true;
      lengthMismatchCount++;
      rows.push({
        record: r + 1,
        field: "Record length",
        position: "",
        expectedLength: e.length,
        actualLength: a.length,
        expected: String(e.length),
        actual: String(a.length),
        issue: "Record length mismatch"
      });
    }

    const maxLen = Math.max(e.length, a.length);
    for (let c = 0; c < maxLen; c++) {
      const ec = c < e.length ? e[c] : "";
      const ac = c < a.length ? a[c] : "";
      if (ec !== ac) {
        recordHasDiff = true;
        totalCharDiffs++;
        rows.push({
          record: r + 1,
          field: "Unmapped field",
          position: c + 1,
          expectedLength: e.length,
          actualLength: a.length,
          expected: ec === "" ? "<END OF RECORD>" : ec === " " ? "<SPACE>" : ec,
          actual: ac === "" ? "<END OF RECORD>" : ac === " " ? "<SPACE>" : ac,
          issue: "Character mismatch"
        });
      }
    }
    if (recordHasDiff) differentRecords++;
  }
  return { rows, differentRecords, totalCharDiffs, lengthMismatchCount, expectedCount: expected.length, actualCount: actual.length };
}
function main() {
  if (process.argv.includes("--help") || process.argv.includes("-h")) {
    usage(); return;
  }
  const expectedPath = arg("expected");
  const actualPath = arg("actual");
  if (!expectedPath || !actualPath) {
    usage();
    throw new Error("Both --expected and --actual are required.");
  }

  const expected = readRecords(expectedPath);
  const actual = readRecords(actualPath);
  const result = compare(expected, actual);
  const outDir = arg("outdir", "comparison_reports");
  fs.mkdirSync(outDir, { recursive: true });

  const stamp = new Date().toISOString().replace(/[:.]/g, "-");
  const baseName = `comparison_report_${stamp}`;
  const csvPath = path.join(outDir, `${baseName}.csv`);
  const txtPath = path.join(outDir, `${baseName}.txt`);

  const headers = ["Record", "Field", "Position (1-based)", "Expected Record Length", "Actual Record Length", "Expected Value", "Actual Value", "Issue"];
  const csvLines = [headers.map(csvCell).join(",")];
  for (const row of result.rows) {
    csvLines.push([
      row.record, row.field, row.position, row.expectedLength, row.actualLength,
      row.expected, row.actual, row.issue
    ].map(csvCell).join(","));
  }
  fs.writeFileSync(csvPath, csvLines.join("\r\n") + "\r\n", "utf8");

  const summary = [
    "FIXED-WIDTH FILE COMPARISON REPORT",
    "=======================================",
    `Expected file: ${path.resolve(expectedPath)}`,
    `Actual file:   ${path.resolve(actualPath)}`,
    "",
    `Expected record count: ${result.expectedCount}`,
    `Actual record count:   ${result.actualCount}`,
    `Records with differences: ${result.differentRecords}`,
    `Record length mismatches: ${result.lengthMismatchCount}`,
    `Character position mismatches: ${result.totalCharDiffs}`,
    `Overall result: ${result.rows.length === 0 ? "PASS — files match exactly" : "FAIL — discrepancies found"}`,
    "",
    "Detailed differences (positions are 1-based; spaces are shown as <SPACE>):"
  ];
  if (!result.rows.length) summary.push("No discrepancies found.");
  else {
    for (const row of result.rows) {
      summary.push(
        `Record ${row.record} | ${row.issue} | Field: ${row.field} | Position: ${row.position || "-"} | ` +
        `Expected: ${visible(row.expected)} | Actual: ${visible(row.actual)} | ` +
        `Lengths: ${row.expectedLength || "-"} vs ${row.actualLength || "-"}`
      );
    }
    summary.push("", `Full row-by-row discrepancy details are also available in: ${path.resolve(csvPath)}`);
  }
  fs.writeFileSync(txtPath, summary.join("\n") + "\n", "utf8");

  console.log(summary.slice(0, 16).join("\n"));
  console.log(`\nText report: ${path.resolve(txtPath)}`);
  console.log(`CSV report:  ${path.resolve(csvPath)}`);
  if (result.rows.length) process.exitCode = 1;
}
try { main(); }
catch (err) {
  console.error(`ERROR: ${err.message}`);
  process.exitCode = 2;
}
