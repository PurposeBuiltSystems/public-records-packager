/*
 * Public Records Packager — bundle planning (pure logic, no Office/Graph).
 *
 * Turns a set of selected messages into a records-bundle plan: numbered,
 * chronological, filesystem-safe names, plus the manifest.csv content the
 * records officer fills exemptions into. Deterministic; no AI.
 *
 * Works in the browser (global `Packager`) and in Node (module.exports).
 */
(function (root) {
  "use strict";

  /** Filesystem/OneDrive-safe name: strip reserved chars, collapse space. */
  function sanitize(name, maxLen) {
    var s = String(name || "")
      .replace(/[\\/:*?"<>|#%{}~&]+/g, " ")
      .replace(/\s+/g, " ")
      .trim()
      .replace(/[. ]+$/, ""); // OneDrive rejects trailing dots/spaces
    if (!s) { s = "untitled"; }
    var cap = maxLen || 60;
    return s.length > cap ? s.slice(0, cap).trim() : s;
  }

  function pad(n, width) {
    var s = String(n);
    while (s.length < width) { s = "0" + s; }
    return s;
  }

  function dateStamp(iso) {
    var d = new Date(iso);
    if (isNaN(d)) { return "0000-00-00"; }
    return d.toISOString().slice(0, 10);
  }

  function addrList(recips) {
    return (recips || [])
      .map(function (r) { return ((r || {}).emailAddress || {}).address || ""; })
      .filter(Boolean)
      .join("; ");
  }

  /**
   * @param selected messages [{id, subject, from, toRecipients, ccRecipients,
   *                 receivedDateTime, hasAttachments, attachments?: [{name}]}]
   * @returns plan: chronological (oldest first, like a records production),
   *   [{num, message, base ("0001_2026-03-02_Bridge inspection"),
   *     emlName, htmlName, attDir}]
   */
  function plan(selected) {
    var ordered = (selected || []).slice().sort(function (a, b) {
      return (Date.parse(a.receivedDateTime || 0) || 0) - (Date.parse(b.receivedDateTime || 0) || 0);
    });
    return ordered.map(function (m, i) {
      var num = pad(i + 1, 4);
      var base = num + "_" + dateStamp(m.receivedDateTime) + "_" + sanitize(m.subject, 48);
      return {
        num: num,
        message: m,
        base: base,
        emlName: base + ".eml",
        htmlName: base + ".html",
        attDir: num + "_attachments",
      };
    });
  }

  /** De-duplicate attachment filenames within one message's folder. */
  function attachmentNames(names) {
    var seen = {};
    return (names || []).map(function (raw) {
      var name = sanitize(raw, 80);
      var dot = name.lastIndexOf(".");
      var stem = dot > 0 ? name.slice(0, dot) : name;
      var ext = dot > 0 ? name.slice(dot) : "";
      var candidate = name;
      var k = 2;
      while (seen[candidate.toLowerCase()]) {
        candidate = stem + " (" + k + ")" + ext;
        k++;
      }
      seen[candidate.toLowerCase()] = true;
      return candidate;
    });
  }

  function csvField(v) {
    var s = String(v == null ? "" : v);
    if (/[",\n\r]/.test(s)) { return '"' + s.replace(/"/g, '""') + '"'; }
    return s;
  }

  /**
   * manifest.csv — one row per message, with blank exemption columns for the
   * records officer to complete (statute citation is required when withholding).
   */
  function manifestCsv(planItems) {
    var header = [
      "number", "date", "from", "to", "cc", "subject",
      "attachments", "eml_file", "html_file",
      "released_in_full", "exemption_notes", "statute_cited",
    ];
    var lines = [header.join(",")];
    planItems.forEach(function (p) {
      var m = p.message;
      var from = ((m.from || {}).emailAddress || {}).address || "";
      lines.push([
        p.num,
        m.receivedDateTime || "",
        from,
        addrList(m.toRecipients),
        addrList(m.ccRecipients),
        m.subject || "",
        (p.attachmentNames || []).join("; "),
        p.emlName,
        p.htmlName,
        "", "", "",
      ].map(csvField).join(","));
    });
    return lines.join("\r\n") + "\r\n";
  }

  function esc(s) {
    return String(s == null ? "" : s)
      .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
  }

  /** Self-contained printable HTML for one message (headers + body). */
  function messageHtml(m, bodyHtml) {
    var from = (m.from || {}).emailAddress || {};
    return "<!DOCTYPE html><html><head><meta charset='utf-8'><title>" +
      esc(m.subject) + "</title><style>" +
      "body{font-family:-apple-system,Segoe UI,sans-serif;margin:2rem;color:#222}" +
      "table.h{border-collapse:collapse;margin-bottom:1.5rem;font-size:14px}" +
      "table.h td{padding:2px 10px 2px 0;vertical-align:top}" +
      "table.h td:first-child{font-weight:600;color:#555;white-space:nowrap}" +
      ".body{border-top:1px solid #ccc;padding-top:1rem}" +
      "@media print{body{margin:0.5in}}" +
      "</style></head><body><table class='h'>" +
      "<tr><td>From:</td><td>" + esc(from.name || "") + " &lt;" + esc(from.address || "") + "&gt;</td></tr>" +
      "<tr><td>Sent:</td><td>" + esc(m.receivedDateTime || "") + "</td></tr>" +
      "<tr><td>To:</td><td>" + esc(addrList(m.toRecipients)) + "</td></tr>" +
      (m.ccRecipients && m.ccRecipients.length ? "<tr><td>Cc:</td><td>" + esc(addrList(m.ccRecipients)) + "</td></tr>" : "") +
      "<tr><td>Subject:</td><td>" + esc(m.subject || "") + "</td></tr>" +
      (m.hasAttachments ? "<tr><td>Attachments:</td><td>see companion folder</td></tr>" : "") +
      "</table><div class='body'>" + (bodyHtml || "") + "</div></body></html>";
  }

  /** Bundle README so the folder is self-explanatory to counsel/requesters. */
  function readmeText(bundleName, count) {
    return bundleName + "\r\n" +
      "Generated " + new Date().toISOString().slice(0, 10) + " with Public Records Packager.\r\n\r\n" +
      "Contents: " + count + " messages, numbered chronologically (oldest first).\r\n" +
      "- NNNN_date_subject.eml  - exact archival copy of the message (opens in Outlook)\r\n" +
      "- NNNN_date_subject.html - printable copy (open in a browser; print to PDF for release)\r\n" +
      "- NNNN_attachments/      - the message's file attachments in original format\r\n" +
      "- manifest.csv           - index of all messages with blank exemption columns\r\n\r\n" +
      "Redaction and exemption review are performed downstream by the records officer.\r\n";
  }

  /** Client-side filters applied after mailbox search. */
  function applyFilters(messages, opts) {
    opts = opts || {};
    var fromT = opts.fromDate ? Date.parse(opts.fromDate) : null;
    var toT = opts.toDate ? Date.parse(opts.toDate) + 864e5 : null; // inclusive day
    var person = (opts.participant || "").toLowerCase().trim();
    return (messages || []).filter(function (m) {
      var t = Date.parse(m.receivedDateTime || 0) || 0;
      if (fromT && t < fromT) { return false; }
      if (toT && t >= toT) { return false; }
      if (person) {
        var haystack = [((m.from || {}).emailAddress || {}).address, ((m.from || {}).emailAddress || {}).name]
          .concat((m.toRecipients || []).concat(m.ccRecipients || []).map(function (r) {
            var e = (r || {}).emailAddress || {};
            return (e.address || "") + " " + (e.name || "");
          }))
          .join(" ").toLowerCase();
        if (haystack.indexOf(person) === -1) { return false; }
      }
      return true;
    });
  }

  var api = {
    plan: plan,
    manifestCsv: manifestCsv,
    messageHtml: messageHtml,
    readmeText: readmeText,
    attachmentNames: attachmentNames,
    applyFilters: applyFilters,
    sanitize: sanitize,
  };
  if (typeof module !== "undefined" && module.exports) { module.exports = api; }
  else { root.Packager = api; }
})(typeof self !== "undefined" ? self : this);
