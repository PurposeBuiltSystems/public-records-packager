/* Offline unit tests for bundle planning. Run: npm test */
"use strict";
var P = require("../src/packager.js");

var failures = 0;
function check(label, actual, expected) {
  if (actual !== expected) {
    failures++;
    console.error("FAIL  " + label + "\n  expected: " + JSON.stringify(expected) + "\n  actual:   " + JSON.stringify(actual));
  }
}

function msg(id, subject, date, from, to) {
  return {
    id: id, subject: subject, receivedDateTime: date,
    from: { emailAddress: { address: from, name: from } },
    toRecipients: (to || []).map(function (a) { return { emailAddress: { address: a, name: a } }; }),
    ccRecipients: [],
  };
}

// plan: chronological, numbered, sanitized
var plan = P.plan([
  msg("b", 'Bridge "phase 2": costs?', "2026-03-10T10:00:00Z", "ann@x.gov", ["me@dot.gov"]),
  msg("a", "RE: Bridge inspection / schedule", "2026-03-02T10:00:00Z", "bob@x.gov", ["me@dot.gov"]),
]);
check("oldest first", plan[0].message.id, "a");
check("numbering", plan[0].num, "0001");
check("slash stripped from name", plan[0].base.indexOf("/") === -1, true);
check("quote+colon stripped", plan[1].base.indexOf('"') === -1 && plan[1].base.indexOf(":") === -1, true);
check("eml name", plan[0].emlName, plan[0].base + ".eml");
check("att dir", plan[1].attDir, "0002_attachments");

// sanitize edge cases
check("reserved chars", P.sanitize('a<b>c:d"e/f\\g|h?i*j'), "a b c d e f g h i j");
check("trailing dot stripped", P.sanitize("report."), "report");
check("empty becomes untitled", P.sanitize("///"), "untitled");

// attachment de-dup
var names = P.attachmentNames(["scan.pdf", "Scan.PDF", "scan.pdf", "notes.txt"]);
check("dup 1 kept", names[0], "scan.pdf");
check("dup 2 renamed", names[1], "Scan (2).PDF");
check("dup 3 renamed", names[2], "scan (3).pdf");

// manifest csv: quoting + blank exemption columns
plan[0].attachmentNames = ["plan.xlsx"];
var csv = P.manifestCsv(plan);
var lines = csv.trim().split("\r\n");
check("header has exemption cols", lines[0].indexOf("exemption_notes,statute_cited") !== -1, true);
check("row count", lines.length, 3);
check("comma-subject quoted", lines[2].indexOf('"Bridge ""phase 2"": costs?"') !== -1, true);
check("trailing blanks for officer", lines[1].slice(-3), ",,,".slice(0, 3));

// message html: escaping + headers
var html = P.messageHtml(plan[1].message, "<p>body</p>");
check("from present", html.indexOf("ann@x.gov") !== -1, true);
check("body embedded", html.indexOf("<p>body</p>") !== -1, true);
var evil = P.messageHtml(msg("x", "<script>alert(1)</script>", "2026-01-01T00:00:00Z", "a@b.c", []), "");
check("html-unsafe subject escaped", evil.indexOf("<script>") === -1, true);

// filters: date range + participant
var msgs = [
  msg("1", "s", "2026-03-01T00:00:00Z", "ann@x.gov", ["me@dot.gov"]),
  msg("2", "s", "2026-03-15T00:00:00Z", "bob@x.gov", ["me@dot.gov"]),
  msg("3", "s", "2026-04-01T00:00:00Z", "ann@x.gov", ["carol@y.gov"]),
];
check("date filter", P.applyFilters(msgs, { fromDate: "2026-03-10", toDate: "2026-03-31" }).length, 1);
check("participant filter", P.applyFilters(msgs, { participant: "ann@x.gov" }).length, 2);
check("participant+date", P.applyFilters(msgs, { participant: "ann@x.gov", fromDate: "2026-03-20" }).length, 1);

if (failures) {
  console.error("\n" + failures + " packager test(s) FAILED");
  process.exit(1);
}
console.log("All packager tests passed.");
