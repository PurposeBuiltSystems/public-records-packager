"use strict";
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");
const raw = fs.readFileSync(path.join(__dirname, "../src/taskpane/taskpane.js"), "utf8");
const end = raw.lastIndexOf("})();");
const source = raw.slice(0, end) + `self.workflow={build,search,seed(items,capped){
  found=items;searchCapped=!!capped;checked={};items.forEach(m=>checked[m.id]=true);
}};` + raw.slice(end);
function load(options = {}) {
  const elements = {}, uploads = [], calls = [];
  const el = id => elements[id] || (elements[id] = { value: "", disabled: false, hidden: false, textContent: "" });
  const context = { self: {}, document: { getElementById: el }, Office: { onReady() {} },
    Packager: require("../src/packager.js"), GraphData: {
      getToken: async () => "token", ensureDrive: async () => {},
      createBundleFolder: async () => { calls.push("folder"); if (options.conflict) throw Error("already exists"); return "Records Requests/New"; },
      getMime: async () => "mime", getBody: async () => "<p>body</p>",
      getAttachmentExport: async () => ({ files: [{ name: "empty.txt", contentBytes: "" }], notExported: options.omitted ? [{ name: "linked file", type: "referenceAttachment" }] : [] }),
      ensureFolder: async () => {}, b64ToBytes: () => new Uint8Array(), folderLink: async () => "",
      uploadFile: async (token, dest, content) => {
        if (options.hold && !uploads.length) { await options.hold; }
        if (options.fail && dest.endsWith(".html")) { throw Error("upload lost"); }
        uploads.push({ dest, content });
      },
    } };
  vm.runInNewContext(source, context);
  el("bundleName").value = "New";
  context.self.workflow.seed([{ id: "m", receivedDateTime: "2026-10-01", subject: "Record", hasAttachments: true }], options.capped);
  return { api: context.self.workflow, elements, uploads, calls, graph: context.GraphData };
}
(async () => {
  const clean = load(); await clean.api.build();
  const statuses = clean.uploads.filter(u => u.dest.endsWith("EXPORT-STATUS.txt"));
  assert.match(statuses[0].content, /^INCOMPLETE/);
  assert.match(statuses.at(-1).content, /^EXPORT FINISHED/);
  assert.ok(clean.uploads.some(u => u.dest.endsWith("manifest.csv")));
  for (const options of [{ omitted: true }, { capped: true }]) {
    const limited = load(options); await limited.api.build();
    assert.match(limited.uploads.at(-1).content, /^REVIEW REQUIRED/);
    assert.match(limited.elements.status.textContent, /review required/);
  }
  const failed = load({ fail: true }); await failed.api.build();
  assert.match(failed.elements.status.textContent, /folder is incomplete/);
  assert.equal(failed.uploads.filter(u => u.dest.endsWith("EXPORT-STATUS.txt")).length, 1);
  const conflict = load({ conflict: true }); await conflict.api.build();
  assert.equal(conflict.uploads.length, 0, "never overwrite an existing bundle");
  const stale = load();
  // Simulate new task-pane JS with a cached older Graph data layer.
  // The export guard must stop before a folder or file is written.
  stale.graph.getAttachmentExport = undefined;
  await stale.api.build();
  assert.equal(stale.calls.length, 0);
  assert.match(stale.elements.status.textContent, /cached an older exporter/);
  let release; const hold = new Promise(resolve => { release = resolve; });
  const busy = load({ hold }); const first = busy.api.build();
  await new Promise(setImmediate); await busy.api.build(); await busy.api.search();
  assert.equal(busy.calls.length, 1, "ignore overlapping export/search actions");
  release(); await first;
  assert.equal(busy.elements.search.disabled, false);
  console.log("Records bundle workflow regressions passed (all Graph writes mocked).");
})().catch(error => { console.error(error); process.exitCode = 1; });
