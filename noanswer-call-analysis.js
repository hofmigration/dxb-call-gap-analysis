// noanswer-call-analysis.js — analyses LOGGED CALLS on "No Answer" leads.
//
// For each No Answer contact created in the window it counts logged calls and
// measures the gap (days) between the 1st and the 2nd ("next") logged call.
// Writes an Excel file (per-contact rows + a summary sheet) AND prints a summary.
// Read-only — it changes nothing in HubSpot.
//
// Needs the HUBSPOT_TOKEN secret. The workflow installs the "xlsx" package for the Excel output.

const XLSX = require("xlsx");
const HUBSPOT_TOKEN = process.env.HUBSPOT_TOKEN;

const LEAD_STAGE_VALUE = "No Answer";
const FROM = Date.UTC(2026, 3, 1);   // 1 Apr 2026
const TO   = Date.UTC(2026, 6, 1);   // 1 Jul 2026 (captures all of June)

// 0 = analyse ALL (~4,240, roughly 4-6 min). Set a number (e.g. 500) for a quick test.
// NOTE: a sample takes the NEWEST contacts first, which under-counts calls.
const SAMPLE_SIZE = 0;

const PORTAL_ID = "23735726";
const OUT_FILE = "noanswer-analysis.xlsx";
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

async function hub(method, path, body) {
  const url = `https://api.hubapi.com${path}`;
  for (let a = 0; a < 6; a++) {
    const res = await fetch(url, {
      method, headers: { Authorization: `Bearer ${HUBSPOT_TOKEN}`, "Content-Type": "application/json" },
      body: body ? JSON.stringify(body) : undefined,
    });
    if (res.status === 429) { await sleep(2000 * (a + 1)); continue; }
    if (!res.ok) { const t = await res.text(); throw new Error(`${method} ${path} -> ${res.status}: ${t.slice(0, 200)}`); }
    return res.status === 204 ? null : res.json();
  }
  throw new Error(`rate-limited: ${method} ${path}`);
}
async function mapPool(items, n, fn) {
  const out = []; let i = 0;
  await Promise.all(Array(Math.min(n, items.length)).fill(0).map(async () => { while (i < items.length) { const idx = i++; out[idx] = await fn(items[idx]); } }));
  return out;
}
const median = (arr) => { if (!arr.length) return 0; const s = [...arr].sort((a, b) => a - b); const m = s.length >> 1; return s.length % 2 ? s[m] : (s[m - 1] + s[m]) / 2; };
const sum = (a) => a.reduce((x, y) => x + y, 0);

async function ownerNameMap() {
  const map = {}; let after;
  for (let i = 0; i < 20; i++) {
    const d = await hub("GET", `/crm/v3/owners/?limit=100${after ? `&after=${after}` : ""}`);
    for (const o of d.results || []) map[String(o.id)] = [o.firstName, o.lastName].filter(Boolean).join(" ") || o.email || String(o.id);
    after = d.paging?.next?.after; if (!after) break;
  }
  return map;
}
async function fetchContacts() {
  const out = []; let after;
  while (true) {
    const d = await hub("POST", "/crm/v3/objects/contacts/search", {
      filterGroups: [{ filters: [
        { propertyName: "lead_stage", operator: "EQ", value: LEAD_STAGE_VALUE },
        { propertyName: "createdate", operator: "BETWEEN", value: String(FROM), highValue: String(TO) },
      ] }],
      sorts: [{ propertyName: "createdate", direction: "DESCENDING" }],
      properties: ["firstname", "lastname", "hubspot_owner_id", "createdate"],
      limit: 100, after,
    });
    out.push(...(d.results || []));
    if (SAMPLE_SIZE && out.length >= SAMPLE_SIZE) return out.slice(0, SAMPLE_SIZE);
    after = d.paging?.next?.after; if (!after) return out;
  }
}
async function callIdsFor(contactId) {
  try { const d = await hub("GET", `/crm/v3/objects/contacts/${contactId}/associations/calls?limit=200`); return (d.results || []).map((r) => r.toObjectId || r.id).filter(Boolean); }
  catch (e) { return []; }
}
async function callTimestamps(ids) {
  const map = {};
  for (let i = 0; i < ids.length; i += 100) {
    const d = await hub("POST", "/crm/v3/objects/calls/batch/read", { properties: ["hs_timestamp"], inputs: ids.slice(i, i + 100).map((id) => ({ id: String(id) })) });
    for (const r of d.results || []) map[r.id] = Date.parse(r.properties.hs_timestamp || 0);
  }
  return map;
}

(async () => {
  if (!HUBSPOT_TOKEN) throw new Error("Missing HUBSPOT_TOKEN");
  console.log(`=== No Answer call analysis — ${new Date().toISOString()} ===`);

  const owners = await ownerNameMap();
  const contacts = await fetchContacts();
  console.log(`Contacts to analyse: ${contacts.length}${SAMPLE_SIZE ? " (SAMPLE)" : " (ALL)"}\n`);

  const idLists = await mapPool(contacts, 6, (c) => callIdsFor(c.id));
  const allIds = [...new Set(idLists.flat())];
  console.log(`Logged calls found: ${allIds.length}`);
  const tsMap = await callTimestamps(allIds);

  const rows = [], callCounts = [], gaps = [];
  let zeroCalls = 0, oneCall = 0;
  contacts.forEach((c, idx) => {
    const stamps = idLists[idx].map((id) => tsMap[id]).filter(Boolean).sort((a, b) => a - b);
    callCounts.push(stamps.length);
    let gap = "";
    if (stamps.length === 0) zeroCalls++;
    else if (stamps.length === 1) oneCall++;
    else { gap = Number(((stamps[1] - stamps[0]) / 86400000).toFixed(1)); gaps.push(gap); }
    rows.push({
      "Contact ID": String(c.id),
      "Name": [c.properties.firstname, c.properties.lastname].filter(Boolean).join(" ") || String(c.id),
      "Owner": owners[c.properties.hubspot_owner_id] || c.properties.hubspot_owner_id || "",
      "Logged Calls": stamps.length,
      "First Call": stamps[0] ? new Date(stamps[0]).toISOString().slice(0, 10) : "",
      "Next Call": stamps[1] ? new Date(stamps[1]).toISOString().slice(0, 10) : "",
      "Gap (days)": gap,
      "Link": `https://app.hubspot.com/contacts/${PORTAL_ID}/record/0-1/${c.id}`,
    });
  });
  rows.sort((a, b) => (b["Gap (days)"] === "" ? -1 : b["Gap (days)"]) - (a["Gap (days)"] === "" ? -1 : a["Gap (days)"]));

  const avgCalls = (sum(callCounts) / (contacts.length || 1)).toFixed(2);
  const avgGap = gaps.length ? (sum(gaps) / gaps.length).toFixed(2) : "n/a";
  const bucket = (arr, ranges) => ranges.map(([l, lo, hi]) => [l, arr.filter((v) => v >= lo && v <= hi).length]);
  const callDist = bucket(callCounts, [["0 calls", 0, 0], ["1", 1, 1], ["2-3", 2, 3], ["4-6", 4, 6], ["7-10", 7, 10], ["11-20", 11, 20], ["21+", 21, 1e9]]);
  const gapDist = bucket(gaps, [["same day", 0, 0.999], ["1 day", 1, 1.999], ["2-3", 2, 3.999], ["4-7", 4, 7.999], ["8-14", 8, 14.999], ["15-30", 15, 30.999], ["30+", 31, 1e9]]);

  // ---- log summary ----
  console.log(`\n--- CALL ATTEMPTS (logged calls per contact) ---`);
  console.log(`Average: ${avgCalls}   Median: ${median(callCounts)}`);
  for (const [l, n] of callDist) console.log(`  ${l.padEnd(8)} ${String(n).padStart(5)}  (${(100 * n / contacts.length).toFixed(1)}%)`);
  console.log(`\n--- GAP: 1st logged call -> next call (days) ---`);
  console.log(`0 calls: ${zeroCalls} | only 1 call: ${oneCall} | 2+ calls (gap measured): ${gaps.length}`);
  if (gaps.length) {
    console.log(`Average gap: ${avgGap} days   Median gap: ${median(gaps).toFixed(1)} days`);
    for (const [l, n] of gapDist) console.log(`  ${l.padEnd(9)} ${String(n).padStart(5)}  (${(100 * n / gaps.length).toFixed(1)}%)`);
  }

  // ---- Excel file ----
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rows), "Per Contact");
  const summary = [
    ["No Answer call analysis", new Date().toISOString().slice(0, 10)],
    ["Window", "2026-04-01 to 2026-06-30"],
    [],
    ["Contacts analysed", contacts.length],
    ["Total logged calls", allIds.length],
    ["Average calls / contact", avgCalls],
    ["Median calls / contact", median(callCounts)],
    ["Contacts with 0 calls", zeroCalls],
    ["Contacts with only 1 call (no follow-up call)", oneCall],
    ["Contacts with 2+ calls (gap measured)", gaps.length],
    ["Average 1st->next gap (days)", avgGap],
    ["Median 1st->next gap (days)", gaps.length ? median(gaps).toFixed(1) : "n/a"],
    [],
    ["Call-count distribution", ""],
    ...callDist,
    [],
    ["Gap distribution (days)", ""],
    ...gapDist,
  ];
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(summary), "Summary");
  XLSX.writeFile(wb, OUT_FILE);
  console.log(`\nWrote ${OUT_FILE} — download it from this run's "Artifacts" section.`);
})().catch((e) => { console.error("FATAL:", e); process.exit(1); });
