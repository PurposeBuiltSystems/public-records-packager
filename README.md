# Public Records Packager

Outlook add-in for records custodians: search your own mailbox for records
responsive to a public-records request, review the chronological results, and
build a numbered records bundle in your own OneDrive — exact `.eml` per
message, printable `.html` copy, attachments in original format, and a
`manifest.csv` with blank exemption columns for the records officer.

The request-management suites track the request; the *collection* step is
still manual Outlook labor everywhere. This tools that step — inside the
tenant, with chain-of-custody sanity: records never touch a third-party
server. Redaction deliberately stays downstream in the agency's tools.

- `manifest.xml` — add-in manifest
- `src/graph.js` — MSAL NAA + Graph (Mail.Read, Files.ReadWrite; MIME export,
  chunked OneDrive uploads)
- `src/packager.js` — pure bundle planner (numbering, sanitization, manifest
  CSV, printable HTML; offline tests in `test/`)
- `src/taskpane/` — search → review → build pane

`npm run validate` checks the manifest; `npm test` runs the planner tests.
