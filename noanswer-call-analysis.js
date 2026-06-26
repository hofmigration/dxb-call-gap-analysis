// noanswer-call-analysis.js — analyses LOGGED CALLS on "No Answer" leads.
//
// For each No Answer contact created in the window, it counts logged calls and
// measures the gap (in days) between the 1st logged call and the 2nd ("next") call.
// Prints a summary to the log. Read-only — it changes nothing in HubSpot.
//
// Run it with the HUBSPOT_TOKEN secret (same token as the compliance agent).

const HUBSPOT_TOKEN = process.env.HUBSPOT_TOKEN;

const LEAD_STAGE_VALUE = "No Answer";          // internal value of lead_stage
const FROM = Date.UTC(2026, 3, 1);             // 1 Apr 2026 (month is 0-based)
const TO   = Date.UTC(2026, 6, 1);             // 1 Jul 2026 (captures all of June)

// Set to 0 to analyse ALL ~4,240 contacts (slower — roughly 10-15 min).
// A few hundred gives a fast, representative read first.
const SAMPLE_SIZE = 500;

const PORTAL_ID = "23735726";
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

async function mapPool(items, concurrency, fn) {
  const out = []; let i = 0;
  await Promise.all(Array(Math.min(concurrency, items.length)).fill(0).map(async () => {
    while (i < items.length) { const idx = i++; out[idx] = await fn(items[idx]); }
  }));
  return out;
}

const median = (arr) => { if (!arr.length) return 0; const s = [...arr].sort((a, b) => a - b); const m = s.length >> 1; return s.length % 2 ? s[m] : (s[m - 1] + s[m]) / 2; };

async function fetchContacts() {
  const out = []; let after;
  while (true) {
    const d = await hub("POST", "/crm/v3/objects/contacts/search", {
      filterGroups: [{ filters: [
        { propertyName: "lead_stage", operator: "EQ", value: LEAD_STAGE_VALUE },
        { propertyName: "createdate", operator: "BETWEEN", value: String(FROM), highValue: String(TO) },
      ] }],
      sorts: [{ propertyName: "createdate", direction: "DESCENDING" }],
      properties: ["firstname", "lastname", "hubspot_owner_id"],
      limit: 100, after,
    });
    out.push(...(d.results || []));
    if (SAMPLE_SIZE && out.length >= SAMPLE_SIZE) return out.slice(0, SAMPLE_SIZE);
    after = d.paging?.next?.after;
    if (!after) return out;
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

  const contacts = await fetchContacts();
  console.log(`Contacts to analyse: ${contacts.length}${SAMPLE_SIZE ? ` (SAMPLE — set SAMPLE_SIZE=0 for all)` : " (ALL)"}\n`);

  // 1) call IDs per contact
  const idLists = await mapPool(contacts, 6, (c) => callIdsFor(c.id));
  // 2) one batch read for all unique call IDs
  const allIds = [...new Set(idLists.flat())];
  console.log(`Logged calls found: ${allIds.length}`);
  const tsMap = await callTimestamps(allIds);

  // 3) per-contact metrics
  const callCounts = [], gaps = [], slow = [];
  let zeroCalls = 0, oneCall = 0;
  contacts.forEach((c, idx) => {
    const stamps = idLists[idx].map((id) => tsMap[id]).filter(Boolean).sort((a, b) => a - b);
    callCounts.push(stamps.length);
    if (stamps.length === 0) zeroCalls++;
    else if (stamps.length === 1) oneCall++;
    else {
      const gapDays = (stamps[1] - stamps[0]) / 86400000;
      gaps.push(gapDays);
      slow.push({ id: c.id, name: [c.properties.firstname, c.properties.lastname].filter(Boolean).join(" ") || c.id, gap: gapDays, calls: stamps.length });
    }
  });

  const sum = (a) => a.reduce((x, y) => x + y, 0);
  const bucket = (arr, ranges) => ranges.map(([label, lo, hi]) => [label, arr.filter((v) => v >= lo && v <= hi).length]);

  console.log(`\n--- CALL ATTEMPTS (logged calls per contact) ---`);
  console.log(`Average: ${(sum(callCounts) / contacts.length).toFixed(1)}   Median: ${median(callCounts)}`);
  for (const [l, n] of bucket(callCounts, [["0 calls", 0, 0], ["1", 1, 1], ["2-3", 2, 3], ["4-6", 4, 6], ["7-10", 7, 10], ["11-20", 11, 20], ["21+", 21, 1e9]]))
    console.log(`  ${l.padEnd(8)} ${String(n).padStart(5)}  (${(100 * n / contacts.length).toFixed(1)}%)`);

  console.log(`\n--- GAP: 1st logged call -> next call (days) ---`);
  console.log(`Contacts with 0 calls: ${zeroCalls}  |  with only 1 call (no follow-up call): ${oneCall}  |  with 2+ calls (gap measured): ${gaps.length}`);
  if (gaps.length) {
    console.log(`Average gap: ${(sum(gaps) / gaps.length).toFixed(1)} days   Median gap: ${median(gaps).toFixed(1)} days`);
    for (const [l, n] of bucket(gaps, [["same day", 0, 0.999], ["1 day", 1, 1.999], ["2-3", 2, 3.999], ["4-7", 4, 7.999], ["8-14", 8, 14.999], ["15-30", 15, 30.999], ["30+", 31, 1e9]]))
      console.log(`  ${l.padEnd(9)} ${String(n).padStart(5)}  (${(100 * n / gaps.length).toFixed(1)}%)`);
    console.log(`\nSlowest 10 first->next-call gaps:`);
    slow.sort((a, b) => b.gap - a.gap).slice(0, 10).forEach((s) =>
      console.log(`  ${s.gap.toFixed(1)}d  ${s.name}  (${s.calls} calls)  https://app.hubspot.com/contacts/${PORTAL_ID}/record/0-1/${s.id}`));
  }
  console.log(`\nDone.`);
})().catch((e) => { console.error("FATAL:", e); process.exit(1); });
