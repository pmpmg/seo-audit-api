const express  = require("express");
const path     = require("path");
const pptxgen  = require("pptxgenjs");
// SVG/sharp rendering removed — using pure PPTX shapes instead
const multer   = require("multer");
const csv      = require("csv-parse/sync");
const { parseScreamingFrogCsv } = require("./parse_screaming_frog");
const { parseMajesticCsv }      = require("./parse_majestic");

const app    = express();
const upload = multer({ storage: multer.memoryStorage() });

app.use(express.json({ limit: "10mb" }));
app.use(express.static(path.join(__dirname, "public")));

// ── ENV ──────────────────────────────────────────────────────
const SEMRUSH_KEY     = process.env.SEMRUSH_API_KEY     || "";
const SEMRUSH_PROJECT = process.env.SEMRUSH_PROJECT_ID  || ""; // Set via form input or Railway env var
const BRIGHTLOCAL_KEY = process.env.BRIGHTLOCAL_API_KEY || "";
const ANTHROPIC_KEY   = process.env.ANTHROPIC_API_KEY   || "";
const PAGESPEED_KEY   = process.env.PAGESPEED_API_KEY   || ""; // Optional — avoids rate limits

// ── JOB QUEUE ────────────────────────────────────────────────
const jobs = new Map(); // jobId -> { status, result, error, createdAt }

function createJob() {
  const id = Math.random().toString(36).slice(2) + Date.now().toString(36);
  jobs.set(id, { status: 'pending', result: null, error: null, createdAt: Date.now() });
  // Clean up jobs older than 30 minutes
  setTimeout(() => jobs.delete(id), 30 * 60 * 1000);
  return id;
}

function setJobDone(id, result)  { const j=jobs.get(id); if(j){ j.status='done';  j.result=result; } }
function setJobError(id, error)  { const j=jobs.get(id); if(j){ j.status='error'; j.error=error;   } }
function getJob(id)              { return jobs.get(id) || null; }

// Store raw PPTX buffers separately for direct binary streaming
const jobBuffers = new Map();

// ── COLORS ───────────────────────────────────────────────────
const C = {
  darkBlue:"12284C", lightBlue:"009ABF", emerald:"00684F",
  frost:"C2F3FF", mint:"CAE7D9", banana:"FFF281",
  white:"FFFFFF", offWhite:"F4F7FA", midGray:"6B7A8D",
  dark:"0D1B2A", red:"C0392B",
};

// ── BENCHMARKS (top-ranking law firms) ───────────────────────
const BENCHMARKS = {
  trustFlow:        { good: 35,  label: "35+",   note: "Top law firm avg" },
  citationFlow:     { good: 40,  label: "40+",   note: "Top law firm avg" },
  referringDomains: { good: 200, label: "200+",  note: "Top law firm avg" },
  totalBacklinks:   { good: 2000,label: "2,000+",note: "Top law firm avg" },
  napConsistency:   { good: 85,  label: "85%+",  note: "Industry standard" },
  citationsFound:   { good: 100, label: "100+",  note: "Industry standard" },
  localRankAvg:     { good: 3,   label: "Top 3", note: "Local pack target", lowerIsBetter: true },
  psPerformance:    { good: 90,  label: "90+",   note: "Google threshold"  },
  siteHealth:       { good: 90,  label: "90+",   note: "SEMrush target"    },
};

function scoreColor(val, benchmark) {
  if (!val || !benchmark) return C.midGray;
  const good = benchmark.lowerIsBetter ? val <= benchmark.good : val >= benchmark.good;
  const close = benchmark.lowerIsBetter ? val <= benchmark.good * 1.5 : val >= benchmark.good * 0.7;
  return good ? C.emerald : close ? C.lightBlue : C.red;
}

// ── API CALLS ─────────────────────────────────────────────────

// SEMrush domain overview
async function semrushDomainOverview(domain) {
  try {
    const url = `https://api.semrush.com/?type=domain_ranks&key=${SEMRUSH_KEY}&export_columns=Dn,Rk,Or,Ot,Oc,Ad,At,Ac&domain=${domain}&database=us`;
    const res = await fetch(url);
    const text = await res.text();
    const lines = text.trim().split("\n");
    if (lines.length < 2) return {};
    const headers = lines[0].split(";");
    const values  = lines[1].split(";");
    const row = {};
    headers.forEach((h,i) => row[h.trim()] = values[i]?.trim());
    return {
      organicKeywords: parseInt(row["Organic Keywords"]) || 0,
      organicTraffic:  parseInt(row["Organic Traffic"])  || 0,
      authorityScore:  parseInt(row["Authority Score"])  || 0,
    };
  } catch(e) { return {}; }
}

// SEMrush organic competitors
async function semrushCompetitors(domain) {
  try {
    const url = `https://api.semrush.com/?type=domain_organic_organic&key=${SEMRUSH_KEY}&export_columns=Dn,Co,Np,Or,Ot&domain=${domain}&database=us&display_limit=3`;
    const res  = await fetch(url);
    const text = await res.text();
    const lines = text.trim().split("\n");
    if (lines.length < 2) return [];
    const headers = lines[0].split(";");
    return lines.slice(1).map(line => {
      const vals = line.split(";");
      const row  = {};
      headers.forEach((h,i) => row[h.trim()] = vals[i]?.trim());
      return {
        domain:          row["Domain"]           || "",
        commonKeywords:  parseInt(row["Common Keywords"]) || 0,
        organicKeywords: parseInt(row["Organic Keywords"]) || 0,
        organicTraffic:  parseInt(row["Organic Traffic"])  || 0,
      };
    }).filter(c => c.domain);
  } catch(e) { return []; }
}

// SEMrush site audit (if configured)
async function semrushSiteAudit(domain) {
  // Returns placeholder — full site audit requires a campaign ID setup
  // Staff still enters SEMrush audit data manually for now
  return {};
}

// BrightLocal Citation Tracker — uses existing report in account
async function brightlocalCitationAudit(domain, businessName, reportId, locationId) {
  try {
    const BASE    = "https://tools.brightlocal.com/seo-tools/api";
    const key     = BRIGHTLOCAL_KEY;
    let REPORT_ID = reportId || "";

    // If no report ID but location ID given, auto-lookup latest report
    if (!REPORT_ID && locationId) {
      try {
        const lookupRes = await fetch(`${BASE}/v2/ct/get-all?api-key=${key}&location-id=${locationId}`);
        const lookupData = await lookupRes.json();
        const reports = lookupData.response?.results || [];
        if (reports.length) {
          const latest = reports.sort((a,b) => new Date(b.last_run||0) - new Date(a.last_run||0))[0];
          REPORT_ID = String(latest.report_id || latest["report-id"] || "");
          console.log("BrightLocal: auto-resolved report ID from location:", REPORT_ID);
        }
      } catch(e2) { console.error("BrightLocal auto-lookup error:", e2.message); }
    }

    if (!REPORT_ID) {
      console.log("BrightLocal: no report ID or location ID provided — skipping.");
      return {};
    }
    console.log("BrightLocal: triggering report run for report", REPORT_ID);
    const runRes  = await fetch(`${BASE}/v2/ct/run`, {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: new URLSearchParams({ "api-key": key, "report-id": REPORT_ID })
    });
    const runData = await runRes.json();
    console.log("BrightLocal run response:", JSON.stringify(runData).slice(0, 200));

    // Step 2: Poll for completion (max 90s, every 10s)
    for (let i = 0; i < 9; i++) {
      await new Promise(r => setTimeout(r, 10000));
      const pollRes  = await fetch(`${BASE}/v2/ct/get?api-key=${key}&report-id=${REPORT_ID}`);
      const pollData = await pollRes.json();
      console.log(`BrightLocal poll ${i+1}:`, JSON.stringify(pollData).slice(0, 200));

      const report = pollData?.response || pollData?.report || pollData;
      const status = report?.status || report?.report_status || "";

      if (status === "complete" || status === "Complete") {
        // Step 3: Get results
        const resultsRes  = await fetch(`${BASE}/v2/ct/get-results?api-key=${key}&report-id=${REPORT_ID}`);
        const resultsData = await resultsRes.json();
        console.log("BrightLocal results:", JSON.stringify(resultsData).slice(0, 300));

        // Parse results — active=found, pending+possible=not found
        const results  = resultsData?.response?.results || {};
        const active   = Array.isArray(results?.active)   ? results.active.length   : 0;
        const pending  = Array.isArray(results?.pending)  ? results.pending.length  : 0;
        const possible = Array.isArray(results?.possible) ? results.possible.length : 0;

        // Count NAP errors from active listings
        const napErrorCount = Array.isArray(results?.active)
          ? results.active.filter(c => c["nap-status"] === "error" || c["business-name-status"] === "error" || c["address-status"] === "error" || c["phone-status"] === "error").length
          : 0;

        const found    = active;
        const notFound = pending + possible;
        const total    = found + notFound;
        const correct  = found - napErrorCount;

        // Also fetch the Key Citation Score from v2/ct/get
        let keyCitationScore = null;
        try {
          const scoreRes  = await fetch(`${BASE}/v2/ct/get?api-key=${key}&report-id=${REPORT_ID}`);
          const scoreData = await scoreRes.json();
          console.log("BrightLocal ct/get response:", JSON.stringify(scoreData).slice(0, 400));
          const rep = scoreData?.response || scoreData?.report || scoreData;
          keyCitationScore = rep?.["key-citation-score"] || rep?.key_citation_score || rep?.citation_score || null;
        } catch(e2) { console.error("BrightLocal score fetch error:", e2.message); }

        console.log(`BrightLocal: found=${found}, notFound=${notFound}, total=${total}, napErrors=${napErrorCount}, keyCitationScore=${keyCitationScore}`);

        return {
          citationsFound:    parseInt(found)||0,
          citationsMissing:  parseInt(notFound)||0,
          citationsTotal:    parseInt(total)||0,
          napErrors:         parseInt(napErrorCount)||0,
          napCorrect:        parseInt(correct)||0,
          keyCitationScore:  keyCitationScore ? parseInt(keyCitationScore)||null : null,
          activeListings:    parseInt(found)||0,
          missingListings:   parseInt(notFound)||0,
        };
      }
    }

    // If not complete after 90s, return last known data from the report
    console.log("BrightLocal: timed out waiting for report — fetching last results");
    const lastRes  = await fetch(`${BASE}/v2/ct/get-results?api-key=${key}&report-id=${REPORT_ID}`);
    const lastData = await lastRes.json();
    const results2  = lastData?.response?.results || {};
    const found2    = Array.isArray(results2?.active)   ? results2.active.length   : 0;
    const notFound2 = (Array.isArray(results2?.pending)  ? results2.pending.length  : 0) +
                      (Array.isArray(results2?.possible) ? results2.possible.length : 0);
    return {
      citationsFound:   found2,
      citationsMissing: notFound2,
      citationsTotal:   found2 + notFound2,
      napErrors:        0,
      napCorrect:       found2,
      keyCitationScore: null,
    };

  } catch(e) {
    console.error("BrightLocal error:", e.message);
    return {};
  }
}

// SEMrush Site Audit — pulls full crawl data from existing campaign
async function semrushSiteAuditData(projectId) {
  try {
    const key = SEMRUSH_KEY;

    // The /info endpoint returns the audit summary directly — no snapshot needed.
    // Actual response shape: { id, name, url, status, errors, warnings, notices,
    //   broken, redirected, healthy, crawled, health_score, ... }
    const infoUrl = `https://api.semrush.com/reports/v1/projects/${projectId}/siteaudit/info?key=${key}`;
    console.log("SEMrush Site Audit info URL:", infoUrl);

    const infoRes  = await fetchWithTimeout(infoUrl, 15000);
    const infoText = await infoRes.text();
    console.log("SEMrush info raw response:", infoText.slice(0, 500));

    let info = {};
    try { info = JSON.parse(infoText); } catch(e) {
      console.error("SEMrush info parse error:", e.message, "| raw:", infoText.slice(0, 100));
      return {};
    }

    // If we got an error object or "campaign not found" text
    if (!info?.id && !info?.status) {
      console.log("SEMrush: unexpected response shape:", JSON.stringify(info).slice(0, 200));
      return {};
    }

    // The /info response uses: healthy, errors, warnings, haveIssues, broken, redirected
    // pagesCrawled = healthy + broken + redirected + blocked (all checked pages)
    const pagesCrawled = (info.healthy||0) + (info.broken||0) + (info.redirected||0) + (info.blocked||0) + (info.errors||0);

    // Health score = healthy pages / total crawled * 100
    const total     = pagesCrawled || 1;
    const siteHealth = Math.round(((info.healthy||0) / total) * 100);

    console.log("SEMrush: project found:", info.name,
      "| healthy:", info.healthy, "| errors:", info.errors,
      "| warnings:", info.warnings, "| pagesCrawled:", pagesCrawled,
      "| siteHealth:", siteHealth);

    // Defect codes — map known SEMrush issue codes to our fields
    // From response: defects: { "112": 44, "102": 7, "135": 2, "216": 2, "217": 68, ... }
    const defects = info.defects || {};
    const issueMap = {
      // Common SEMrush issue code mappings
      missingDesc:    parseInt(defects["217"] || defects["17"]  || 0), // missing meta description
      titlesTooLong:  parseInt(defects["216"] || defects["16"]  || 0), // title too long
      missingH1:      parseInt(defects["15"]  || defects["115"] || 0), // missing H1
      missingAlt:     parseInt(defects["112"] || defects["12"]  || 0), // missing alt text
      schemaErrors:   parseInt(defects["223"] || defects["23"]  || 0), // schema issues
      thinPages:      parseInt(defects["105"] || defects["5"]   || 0), // thin content
      brokenExternal: parseInt(defects["135"] || defects["35"]  || 0), // broken external links
      noAnchors:      parseInt(defects["45"]  || 0),                    // no anchor text
    };
    console.log("SEMrush defect map:", issueMap);

    return {
      pagesCrawled,
      siteHealth,
      pages200:     parseInt(info.healthy    || 0),
      redirects:    parseInt(info.redirected || 0),
      errors:       parseInt(info.broken     || 0),
      aiReadiness:  parseInt(info.ai_readiness_score || 0),
      ...issueMap,
    };
  } catch(e) {
    console.error("SEMrush Site Audit error:", e.message);
    return {};
  }
}

// Google PageSpeed — with per-request timeout
async function fetchWithTimeout(url, ms=12000) {
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), ms);
  try {
    const res = await fetch(url, { signal: controller.signal });
    clearTimeout(timer);
    return res;
  } catch(e) { clearTimeout(timer); throw e; }
}

async function fetchPageSpeed(domain) {
  try {
    const base   = "https://www.googleapis.com/pagespeedonline/v5/runPagespeed";
    const url    = `https://${domain}`;
    const keyStr = PAGESPEED_KEY ? `&key=${PAGESPEED_KEY}` : "";
    let mob = {}, dsk = {};
    try {
      const r = await fetchWithTimeout(`${base}?url=${encodeURIComponent(url)}&strategy=mobile${keyStr}`, 30000);
      mob = await r.json();
      if (mob.error) { console.log("PageSpeed mobile API error:", mob.error.message); mob = {}; }
    } catch(e) { console.log("PageSpeed mobile failed:", e.message); }
    try {
      const r = await fetchWithTimeout(`${base}?url=${encodeURIComponent(url)}&strategy=desktop${keyStr}`, 30000);
      dsk = await r.json();
      if (dsk.error) { console.log("PageSpeed desktop API error:", dsk.error.message); dsk = {}; }
    } catch(e) { console.log("PageSpeed desktop failed:", e.message); }

    const getMetric = (data, id) => data?.lighthouseResult?.audits?.[id]?.displayValue || "—";
    const getScore  = (data) => {
      const s = data?.lighthouseResult?.categories?.performance?.score;
      return s ? Math.round(s * 100) : 0;
    };
    return {
      psMobile:      getScore(mob),
      psDesktop:     getScore(dsk),
      psPerformance: getScore(mob),
      psFCP:         getMetric(mob, "first-contentful-paint"),
      psLCP:         getMetric(mob, "largest-contentful-paint"),
      psTBT:         getMetric(mob, "total-blocking-time"),
      psCLS:         getMetric(mob, "cumulative-layout-shift"),
    };
  } catch(e) { console.error("PageSpeed error:", e.message); return {}; }
}

// Claude narrative
async function getNarrative(data) {
  const prompt = `You are an SEO analyst preparing a sales audit for a law firm prospect.
CRITICAL RULES:
- NEVER invent, guess, or name specific competitors. Only reference competitors if their domain appears in the data.
- Use plain language a law firm partner can understand — no jargon.
- Keep all text concise — titles under 8 words, body text under 40 words.
- Base ALL findings strictly on the numbers provided. Do not invent statistics.
- NEVER invent or calculate percentages (e.g. "100% NAP consistency") — only use exact numbers from the data provided.
- For NAP/citations: use citationsFound, citationsTotal, napErrors exactly as given. Never derive percentages from them.
- If a data field is 0 or missing, do not mention it as a positive. Find a different genuine positive to highlight.
- whatIsWorking items must reference a specific number from the audit data, not a derived or assumed statistic.

Return ONLY valid JSON, no markdown:
{
  "executiveSummary": "One sentence — the single biggest opportunity on this site.",
  "whatIsWorking": [
    {"title":"Short title under 6 words — no punctuation","sub":"One SHORT sentence, max 18 words. State ONE fact. Never list multiple numbers or metrics. No em-dashes."},
    {"title":"...","sub":"..."},
    {"title":"...","sub":"..."},
    {"title":"...","sub":"..."},
    {"title":"...","sub":"..."},
    {"title":"...","sub":"..."}
  ],
  "problems": [],
  "actions": [
    {"n":"1","title":"Short action title","body":"One sentence on what to do and why. No timelines.","impact":"High","effort":"Low"},
    {"n":"2","title":"...","body":"...","impact":"High","effort":"Med"},
    {"n":"3","title":"...","body":"...","impact":"High","effort":"Med"},
    {"n":"4","title":"...","body":"...","impact":"High","effort":"Low"},
    {"n":"5","title":"...","body":"...","impact":"Med","effort":"Low"},
    {"n":"6","title":"...","body":"...","impact":"High","effort":"High"}
  ],
  "sequence": [
    {"priority":"1","action":"Most impactful first step","why":"One sentence on why this goes first."},
    {"priority":"2","action":"Second step","why":"One sentence."},
    {"priority":"3","action":"Third step","why":"One sentence."},
    {"priority":"4","action":"Ongoing","why":"One sentence on what requires continuous attention."}
  ]
}
AUDIT DATA: ${JSON.stringify(data)}`;

  const res = await fetch("https://api.anthropic.com/v1/messages", {
    method: "POST",
    headers: {
      "Content-Type":"application/json",
      "x-api-key": ANTHROPIC_KEY,
      "anthropic-version":"2023-06-01"
    },
    body: JSON.stringify({
      model:"claude-haiku-4-5-20251001",
      max_tokens:2000,
      messages:[{role:"user",content:prompt}]
    })
  });
  const json = await res.json();
  if (json.error) throw new Error(json.error.message);
  return JSON.parse(json.content[0].text.replace(/```json|```/g,"").trim());
}

// parseMajesticCsv — imported from ./parse_majestic


// SEMrush organic keywords — top 100 ranking URLs + keywords for content audit
async function semrushOrganicKeywords(domain) {
  try {
    const url = `https://api.semrush.com/?type=domain_organic&key=${SEMRUSH_KEY}&export_columns=Ph,Po,Nq,Ur&domain=${domain}&database=us&display_limit=100&display_sort=nq_desc`;
    const res  = await fetchWithTimeout(url);
    const text = await res.text();
    const rows = text.trim().split("\n").slice(1).map(r => {
      const [keyword, position, volume, url] = r.split(";");
      return {
        keyword: (keyword||"").trim(),
        position: parseInt(position)||99,
        volume:   parseInt(volume)||0,
        url:      (url||"").trim().toLowerCase()
      };
    }).filter(r => r.keyword);

    // Practice area taxonomy for law firms
    const TOPICS = [
      { label:"Personal Injury",   terms:["personal injury","car accident","auto accident","slip and fall","wrongful death","motorcycle accident","truck accident"] },
      { label:"Criminal Defense",  terms:["criminal defense","dui","dwi","felony","misdemeanor","drug charge","assault","criminal lawyer"] },
      { label:"Family Law",        terms:["divorce","family law","child custody","alimony","child support","adoption","domestic violence"] },
      { label:"Estate Planning",   terms:["estate planning","will","trust","probate","power of attorney","elder law"] },
      { label:"Business Law",      terms:["business law","contract","llc","corporate","employment","commercial litigation"] },
      { label:"Real Estate",       terms:["real estate","property","landlord","tenant","foreclosure","title"] },
      { label:"Immigration",       terms:["immigration","visa","green card","citizenship","deportation","asylum"] },
      { label:"Workers Comp",      terms:["workers compensation","workers comp","workplace injury","work accident"] },
      { label:"Social Security",   terms:["social security","disability","ssdi","ssi"] },
      { label:"Bankruptcy",        terms:["bankruptcy","chapter 7","chapter 13","debt relief"] },
    ];

    // Score each topic: 2=strong (top10), 1=weak (11-20), 0=missing
    const topicCoverage = TOPICS.map(t => {
      const matches = rows.filter(r =>
        t.terms.some(term => r.keyword.includes(term) || r.url.includes(term.replace(/ /g,"-")))
      );
      const top10  = matches.filter(r => r.position <= 10).length;
      const top20  = matches.filter(r => r.position <= 20).length;
      const score  = top10 >= 2 ? 2 : top10 >= 1 ? 2 : top20 >= 1 ? 1 : 0;
      return { label: t.label, score, keywords: matches.length };
    });

    // Page type analysis from URLs
    const US_STATES = ["alabama","alaska","arizona","arkansas","california","colorado","connecticut","delaware","florida","georgia","hawaii","idaho","illinois","indiana","iowa","kansas","kentucky","louisiana","maine","maryland","massachusetts","michigan","minnesota","mississippi","missouri","montana","nebraska","nevada","new-hampshire","new-jersey","new-mexico","new-york","north-carolina","north-dakota","ohio","oklahoma","oregon","pennsylvania","rhode-island","south-carolina","south-dakota","tennessee","texas","utah","vermont","virginia","washington","west-virginia","wisconsin","wyoming"];
    const LOCATION_TERMS = ["location","locations","office","offices","city","cities","near","serving",...US_STATES];
    const SERVICE_TERMS  = ["service","services","practice","practice-area","attorney","lawyer","law","legal"];

    const uniqueUrls = [...new Set(rows.map(r => r.url))];
    const urlKeywordCount = {};
    rows.forEach(r => { urlKeywordCount[r.url] = (urlKeywordCount[r.url]||0) + 1; });

    const locationPages = uniqueUrls.filter(u => LOCATION_TERMS.some(t => u.includes(t))).length;
    const servicePages  = uniqueUrls.filter(u => SERVICE_TERMS.some(t => u.includes(t)) && !LOCATION_TERMS.some(t => u.includes(t))).length;
    const thinPages     = uniqueUrls.filter(u => urlKeywordCount[u] <= 2).length;
    const strongPages   = uniqueUrls.filter(u => urlKeywordCount[u] >= 5).length;

    // Position buckets
    const top3  = rows.filter(r => r.position <= 3).length;
    const top10 = rows.filter(r => r.position <= 10).length;
    const top20 = rows.filter(r => r.position <= 20).length;

    return { topicCoverage, locationPages, servicePages, thinPages, strongPages, top3, top10, top20, totalKeywords: rows.length };
  } catch(e) {
    console.error("semrushOrganicKeywords error:", e.message);
    return null;
  }
}


// SEMrush competitor content gap — finds keyword topics competitors rank for that client doesn't
async function semrushCompetitorGap(domain, competitorDomains) {
  try {
    if (!competitorDomains || !competitorDomains.length) return null;

    // Fetch client keywords
    const clientUrl = `https://api.semrush.com/?type=domain_organic&key=${SEMRUSH_KEY}&export_columns=Ph,Po,Nq&domain=${domain}&database=us&display_limit=200&display_sort=nq_desc`;
    const clientRes  = await fetchWithTimeout(clientUrl);
    const clientText = await clientRes.text();
    const clientKws  = new Set(
      clientText.trim().split("\n").slice(1)
        .map(r => r.split(";")[0]?.trim().toLowerCase())
        .filter(Boolean)
    );

    // Fetch top 2 competitors' keywords
    const compData = [];
    for (const compDomain of competitorDomains.slice(0,2)) {
      try {
        const url = `https://api.semrush.com/?type=domain_organic&key=${SEMRUSH_KEY}&export_columns=Ph,Po,Nq&domain=${compDomain}&database=us&display_limit=200&display_sort=nq_desc`;
        const res  = await fetchWithTimeout(url);
        const text = await res.text();
        const rows = text.trim().split("\n").slice(1).map(r => {
          const [kw, pos, vol] = r.split(";");
          return { kw:(kw||"").trim().toLowerCase(), pos:parseInt(pos)||99, vol:parseInt(vol)||0 };
        }).filter(r => r.kw);
        compData.push({ domain: compDomain, keywords: rows });
      } catch(e) { console.error(`Gap fetch error for ${compDomain}:`, e.message); }
    }

    if (!compData.length) return null;

    // Find gap keywords: competitor ranks top 20, client doesn't rank at all
    const gapKeywords = [];
    for (const comp of compData) {
      for (const row of comp.keywords) {
        if (row.pos <= 20 && !clientKws.has(row.kw) && row.vol > 50) {
          gapKeywords.push({ ...row, compDomain: comp.domain });
        }
      }
    }

    // Cluster gaps into topic groups using keyword patterns
    const CLUSTERS = [
      { label:"Car Accident",        terms:["car accident","auto accident","vehicle accident","collision","crash"] },
      { label:"Truck Accident",      terms:["truck accident","semi truck","18 wheeler","commercial vehicle","tractor trailer"] },
      { label:"Motorcycle Accident", terms:["motorcycle","motorbike","bike accident"] },
      { label:"Slip & Fall",         terms:["slip and fall","premises liability","trip and fall","unsafe property"] },
      { label:"Wrongful Death",      terms:["wrongful death","fatal accident","death claim"] },
      { label:"Medical Malpractice", terms:["medical malpractice","doctor error","surgical error","misdiagnosis","hospital negligence"] },
      { label:"Workers Comp",        terms:["workers comp","workers compensation","workplace injury","work accident","on the job"] },
      { label:"Dog Bite",            terms:["dog bite","dog attack","animal bite"] },
      { label:"Uber/Rideshare",      terms:["uber","lyft","rideshare","ride share"] },
      { label:"DUI/Drunk Driver",    terms:["drunk driver","dui accident","drunk driving accident"] },
      { label:"Brain Injury",        terms:["brain injury","traumatic brain","tbi","head injury"] },
      { label:"Spinal Injury",       terms:["spinal","spine injury","back injury","paralysis"] },
      { label:"Product Liability",   terms:["product liability","defective product","product recall"] },
      { label:"Insurance Claims",    terms:["insurance claim","insurance company","bad faith","insurance dispute"] },
      { label:"Geo-targeted",        terms:["near me","in colorado","in grand junction","lawyer near","attorney near","local"] },
    ];

    const clustered = CLUSTERS.map(c => {
      const matches = gapKeywords.filter(g => c.terms.some(t => g.kw.includes(t)));
      const totalVol = matches.reduce((s,g) => s+g.vol, 0);
      const competitors = [...new Set(matches.map(g => g.compDomain))];
      return { label: c.label, count: matches.length, volume: totalVol, competitors };
    }).filter(c => c.count > 0)
      .sort((a,b) => b.volume - a.volume)
      .slice(0,8);

    const totalGapVolume = gapKeywords.reduce((s,g) => s+g.vol, 0);
    const totalGapKeywords = gapKeywords.length;

    return { clustered, totalGapVolume, totalGapKeywords, competitorDomains: compData.map(c=>c.domain) };
  } catch(e) {
    console.error("semrushCompetitorGap error:", e.message);
    return null;
  }
}

// ── PPTX HELPERS ─────────────────────────────────────────────
// Pure PPTX gauge — draws donut using arc shapes
async function drawGauge(pres, slide, x, y, w, h, score, label) {
  const color = score >= 90 ? C.emerald : score >= 50 ? C.lightBlue : C.red;
  const badge = score >= 90 ? "EXCELLENT" : score >= 50 ? "GOOD" : "NEEDS WORK";
  const cx = x + w / 2;

  try {
    const pngBuf = await renderGaugePng(score);
    const b64    = pngBuf.toString("base64");
    // Arc image — full width, top portion of the gauge area
    const imgH = h * 0.72;
    slide.addImage({ data: `image/png;base64,${b64}`, x, y, w, h: imgH });

    // Score number centered inside arc
    slide.addText(`${score}`, {
      x, y: y + imgH * 0.28, w, h: imgH * 0.28,
      fontSize: 28, bold: true, color, fontFace: "Calibri",
      align: "center", valign: "middle", margin: 0
    });
    // /100
    slide.addText("/100", {
      x, y: y + imgH * 0.56, w, h: imgH * 0.18,
      fontSize: 10, color: C.midGray, fontFace: "Calibri",
      align: "center", valign: "middle", margin: 0
    });
    // Label
    slide.addText(label, {
      x, y: y + imgH + 0.04, w, h: 0.26,
      fontSize: 12, bold: true, color: C.darkBlue, fontFace: "Calibri",
      align: "center", margin: 0
    });
    // Badge pill
    const bw = 0.9, bh = 0.22;
    slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x: cx - bw/2, y: y + imgH + 0.34, w: bw, h: bh,
      fill: { color }, line: { color, width: 0 }, rectRadius: 0.05
    });
    slide.addText(badge, {
      x: cx - bw/2, y: y + imgH + 0.34, w: bw, h: bh,
      fontSize: 8, bold: true, color: C.white, fontFace: "Calibri",
      align: "center", valign: "middle", margin: 0
    });
  } catch(e) {
    console.error("Gauge render failed, using fallback:", e.message);
    slide.addShape(pres.shapes.OVAL, { x: cx-0.9, y: y+0.1, w:1.8, h:1.8, line:{color:"E2EAF0",width:8}, fill:{color:C.white} });
    slide.addShape(pres.shapes.OVAL, { x: cx-0.9, y: y+0.1, w:1.8, h:1.8, line:{color,width:8}, fill:{color:C.white} });
    slide.addText(`${score}`, { x, y: y+0.6, w, h:0.5, fontSize:28, bold:true, color, fontFace:"Calibri", align:"center", margin:0 });
    slide.addText("/100", { x, y: y+1.1, w, h:0.24, fontSize:10, color:C.midGray, fontFace:"Calibri", align:"center", margin:0 });
    slide.addText(label, { x, y: y+h-0.3, w, h:0.28, fontSize:11, bold:true, color:C.darkBlue, fontFace:"Calibri", align:"center" });
  }
}

// Simple icon substitutes using colored text/shapes
function drawCheckIcon(pres, slide, x, y) {
  slide.addShape(pres.shapes.OVAL, { x, y, w:0.46, h:0.46, fill:{color:C.emerald}, line:{color:C.emerald,width:0} });
  slide.addText("✓", { x:x+0.01, y:y+0.05, w:0.44, h:0.38, fontSize:13, bold:true, color:C.white, fontFace:"Calibri", align:"center", valign:"middle" });
}

const ms=()=>({type:"outer",blur:10,offset:3,angle:135,color:"000000",opacity:0.07});
const footer=(s,d)=>{
  s.addShape("line",{x:0.4,y:5.28,w:9.2,h:0,line:{color:"D8E4EE",width:0.5}});
  s.addText(`Prepared by ${d.preparedBy} for ${d.clientName}`,{x:0.4,y:5.32,w:7,h:0.22,fontSize:8,color:C.midGray,fontFace:"Calibri"});
  s.addText(d.domain,{x:7.4,y:5.32,w:2.2,h:0.22,fontSize:8,color:C.midGray,fontFace:"Calibri",align:"right"});
};
const slbl=(s,t)=>s.addText(t,{x:0.5,y:0.26,w:9,h:0.2,fontSize:9,color:C.lightBlue,bold:true,charSpacing:4,fontFace:"Calibri"});
const stit=(s,t)=>s.addText(t,{x:0.5,y:0.52,w:9,h:0.72,fontSize:32,bold:true,color:C.darkBlue,fontFace:"Calibri"});

function kpi(pres,s,x,y,w,h,val,lbl,sub,color){
  s.addShape(pres.shapes.RECTANGLE,{x,y,w,h,fill:{color:C.white},shadow:ms(),line:{color:"E2EAF0",width:0.5}});
  s.addShape(pres.shapes.RECTANGLE,{x,y,w,h:0.06,fill:{color},line:{color,width:0}});
  s.addText(val,{x:x+0.18,y:y+0.18,w:w-0.3,h:0.72,fontSize:36,bold:true,color,fontFace:"Calibri",margin:0});
  s.addText(lbl,{x:x+0.18,y:y+0.94,w:w-0.3,h:0.28,fontSize:10,bold:true,color:C.darkBlue,fontFace:"Calibri",margin:0});
  if(sub)s.addText(sub,{x:x+0.18,y:y+1.24,w:w-0.3,h:0.5,fontSize:9,color:C.midGray,fontFace:"Calibri",margin:0});
}

// ── GAUGE IMAGE GENERATOR ─────────────────────────────────────
// Renders ONLY the arc ring as PNG via sharp (no text — text overlaid by pptxgenjs)
async function renderGaugePng(score) {
  const sharp = require("sharp");
  const pct   = Math.max(0, Math.min(100, score));
  const color = pct >= 90 ? "#00684F" : pct >= 50 ? "#009ABF" : "#C0392B";

  const SIZE = 300, cx = 150, cy = 150, R = 108, SW = 24;
  const startDeg = 135, sweepDeg = 270;
  const scoreDeg = (sweepDeg * pct) / 100;

  function polarToXY(deg) {
    const rad = (deg * Math.PI) / 180;
    return { x: +(cx + R * Math.cos(rad)).toFixed(3), y: +(cy + R * Math.sin(rad)).toFixed(3) };
  }

  const s      = polarToXY(startDeg);
  const trackE = polarToXY(startDeg + sweepDeg);
  const e      = polarToXY(startDeg + scoreDeg);
  const trackLarge = 1; // 270° is always > 180
  const largeArc   = scoreDeg > 180 ? 1 : 0;

  const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="${SIZE}" height="${SIZE}">
  <rect width="${SIZE}" height="${SIZE}" fill="white"/>
  <path d="M ${s.x} ${s.y} A ${R} ${R} 0 ${trackLarge} 1 ${trackE.x} ${trackE.y}"
    fill="none" stroke="#E2EAF0" stroke-width="${SW}" stroke-linecap="round"/>
  ${pct > 0 ? `<path d="M ${s.x} ${s.y} A ${R} ${R} 0 ${largeArc} 1 ${e.x} ${e.y}"
    fill="none" stroke="${color}" stroke-width="${SW}" stroke-linecap="round"/>` : ""}
</svg>`;

  return await sharp(Buffer.from(svg)).png().toBuffer();
}

// ── BUILD PPTX ────────────────────────────────────────────────
async function buildPptx(data, narrative) {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";
  const D = data;

  // S1 COVER
  const s1=pres.addSlide(); s1.background={color:C.darkBlue};
  s1.addShape(pres.shapes.RECTANGLE,{x:0,y:0,w:0.22,h:5.625,fill:{color:C.lightBlue},line:{color:C.lightBlue,width:0}});
  s1.addShape(pres.shapes.RECTANGLE,{x:0,y:5.1,w:10,h:0.525,fill:{color:"0C1E3A"},line:{color:"0C1E3A",width:0}});
  s1.addText("SEO OPPORTUNITY AUDIT",{x:0.5,y:0.9,w:9,h:0.28,fontSize:11,color:C.lightBlue,bold:true,charSpacing:5,fontFace:"Calibri"});
  s1.addText("What's working.\nWhat's costing you clients.\nWhat to fix first.",{x:0.5,y:1.3,w:6.5,h:1.8,fontSize:28,bold:true,color:C.white,fontFace:"Calibri",valign:"top"});
  s1.addText(D.domain,{x:0.5,y:3.1,w:9,h:0.6,fontSize:22,color:C.lightBlue,fontFace:"Calibri",bold:true});
  s1.addText(D.clientName,{x:0.5,y:3.68,w:9,h:0.3,fontSize:14,color:"A8C4D8",fontFace:"Calibri"});
  s1.addShape(pres.shapes.RECTANGLE,{x:0.5,y:4.1,w:9,h:0.015,fill:{color:"1E3A5A"},line:{color:"1E3A5A",width:0}});
  s1.addText(`Prepared by ${D.preparedBy}   ·   ${D.date||""}   ·   ${D.pagesCrawled||0} pages crawled`,{x:0.5,y:4.22,w:9,h:0.25,fontSize:9,color:"6A8FA8",fontFace:"Calibri"});

  // S2 BOTTOM LINE
  const s2=pres.addSlide(); s2.background={color:C.white};
  slbl(s2,"THE BOTTOM LINE"); stit(s2,"Four numbers that tell the story.");
  s2.addText(narrative.executiveSummary||"",{x:0.5,y:1.38,w:9,h:0.4,fontSize:11,color:C.dark,fontFace:"Calibri"});
  kpi(pres,s2,0.4, 1.95,2.15,2.1,`${D.siteHealth||0}/100`,"SITE HEALTH",    "Overall technical score (SEMrush)", scoreColor(D.siteHealth,BENCHMARKS.siteHealth));
  kpi(pres,s2,2.72,1.95,2.15,2.1,`${D.psPerformance||0}/100`,"PAGE SPEED", "Google PageSpeed mobile score",     scoreColor(D.psPerformance,BENCHMARKS.psPerformance));
  kpi(pres,s2,5.04,1.95,2.15,2.1,`${D.schemaErrors||0}`,"SCHEMA ERRORS",   "One broken template, sitewide",    D.schemaErrors>0?C.red:C.emerald);
  (() => {
    const cf = parseInt(D.citationsFound)||0;
    const ct = parseInt(D.citationsTotal)||0;
    const ks = parseInt(D.keyCitationScore)||null;
    const val = ks ? `${ks}/100` : cf ? `${cf} / ${ct||"?"}` : "—";
    kpi(pres,s2,7.36,1.95,2.15,2.1,val,"CITATION SCORE","BrightLocal key citation score", scoreColor(ks,{good:70,label:"70+",note:"Competitive for law firms"}));
  })();
  s2.addShape(pres.shapes.RECTANGLE,{x:0.4,y:4.22,w:9.2,h:0.65,fill:{color:C.darkBlue},line:{color:C.darkBlue,width:0}});
  s2.addText(`💡  ${narrative.executiveSummary||"Key opportunities identified across technical SEO, page speed, and local search."}`,{x:0.6,y:4.27,w:8.8,h:0.55,fontSize:10,color:C.white,fontFace:"Calibri"});
  footer(s2,D);

  // S3 WHAT'S WORKING
  const s3=pres.addSlide(); s3.background={color:C.white};
  slbl(s3,"WHAT'S WORKING"); stit(s3,"The expensive stuff is already right.");
  s3.addText("These are the issues that cost the most to fix after the fact — none of them are problems here.",{x:0.5,y:1.38,w:9,h:0.35,fontSize:11,color:C.dark,fontFace:"Calibri"});
  const wins=narrative.whatIsWorking||[
    {title:"Zero server errors",       sub:`No 5xx failures across ${D.pagesCrawled||0} pages.`},
    {title:"No broken internal links", sub:"Internal link graph is fully intact."},
    {title:"Every page has a title",   sub:"No missing or empty title tags found."},
    {title:"Site is fully crawlable",  sub:"No robots.txt or noindex blocking core pages."},
    {title:"HTTPS secure",             sub:"SSL certificate valid across all pages."},
    {title:"Mobile responsive",        sub:"Site renders correctly on mobile devices."},
  ];
  for(let i=0;i<6;i++){
    const col=i%3,row=Math.floor(i/3),x=0.4+col*3.1,y=1.82+row*1.58;
    s3.addShape(pres.shapes.RECTANGLE,{x,y,w:2.9,h:1.42,fill:{color:C.offWhite},shadow:ms(),line:{color:"E2EAF0",width:0.3}});
    s3.addShape(pres.shapes.RECTANGLE,{x,y,w:2.9,h:0.05,fill:{color:C.emerald},line:{color:C.emerald,width:0}});
    drawCheckIcon(pres,s3, x+0.18, y+0.18);
    s3.addText(wins[i]?.title||"",{x:x+0.82,y:y+0.18,w:1.95,h:0.32,fontSize:11,bold:true,color:C.darkBlue,fontFace:"Calibri",margin:0});
    s3.addText(wins[i]?.sub||"",  {x:x+0.82,y:y+0.52,w:1.95,h:0.72,fontSize:9,color:C.midGray,fontFace:"Calibri",margin:0});
  }
  footer(s3,D);

  // S4 TOP ISSUES — data-driven, replaces "Three Patterns" narrative slide
  const s4=pres.addSlide(); s4.background={color:C.white};
  slbl(s4,"WHAT'S COSTING YOU CLIENTS");
  stit(s4,"The biggest opportunities, by the numbers.");
  s4.addText("These are the highest-impact issues found in the crawl — each one affects rankings, traffic, or client trust.",{x:0.5,y:1.38,w:9,h:0.3,fontSize:11,color:C.dark,fontFace:"Calibri"});

  // Build top issues from real data — ranked by severity/count
  const SF4 = D.sf || {};
  const totalPages4 = SF4.htmlPages || D.pagesCrawled || 0;

  // Score each issue by: count * impact weight
  const allIssues4 = [
    { title:"Title tags too long",        stat:`${SF4.titlesOver60||D.titlesTooLong||0} of ${totalPages4} pages`,   val:SF4.titlesOver60||D.titlesTooLong||0, weight:3, why:"Titles truncate in search results, hurting click-through rates on every affected page.",    fix:"Rewrite titles under 60 characters with primary keyword near the front.", effort:"Low"  },
    { title:"Meta descriptions too long", stat:`${SF4.metaOver155||0} of ${totalPages4} pages`,                     val:SF4.metaOver155||0,                   weight:2, why:"Oversized meta descriptions are rewritten by Google, removing your carefully crafted messaging.", fix:"Trim meta descriptions to 150 characters max with a clear call-to-action.", effort:"Low"  },
    { title:"Duplicate meta descriptions",stat:`${SF4.metaDuplicate||D.missingDesc||0} of ${totalPages4} pages`,    val:SF4.metaDuplicate||D.missingDesc||0,  weight:3, why:"Duplicate descriptions confuse search engines on which page to rank for a given query.",     fix:"Write unique descriptions for each page targeting its specific keyword.",   effort:"Med"  },
    { title:"Thin or low content pages",  stat:`${SF4.lowContentPages||D.thinPages||0} of ${totalPages4} pages`,    val:SF4.lowContentPages||D.thinPages||0,  weight:4, why:"Pages with minimal content signal low quality to Google and rarely earn rankings.",            fix:"Expand thin pages with 500+ words of relevant, specific content.",         effort:"High" },
    { title:"Missing canonical tags",     stat:`${SF4.missingCanonical||0} of ${totalPages4} pages`,                val:SF4.missingCanonical||0,              weight:3, why:"Without canonicals, link equity splits across duplicate URLs, diluting ranking power.",        fix:"Add self-referencing canonical tags to all indexable pages via template.",  effort:"Low"  },
    { title:"Images missing alt text",    stat:`${SF4.missingAltText||D.missingAlt||0} images`,                     val:SF4.missingAltText||D.missingAlt||0,  weight:2, why:"Missing alt text loses image search traffic and creates accessibility compliance risk.",         fix:"Add descriptive alt text to all images including target keyword where natural.", effort:"Low" },
    { title:"Schema markup errors",       stat:`${D.schemaErrors||0} errors`,                                        val:D.schemaErrors||0,                    weight:4, why:"Broken schema blocks rich results in search, reducing click-through rates significantly.",     fix:"Fix LocalBusiness and Attorney schema using Google's Rich Results Test.",   effort:"Low"  },
    { title:"Redirect chains",            stat:`${SF4.redirects3xx||0} redirects`,                                   val:SF4.redirects3xx||0,                  weight:2, why:"Redirect chains slow page load and bleed link equity through each extra hop.",                fix:"Update internal links to point directly to final destination URLs.",        effort:"Low"  },
  ].filter(iss => iss.val > 0)
   .sort((a,b) => (b.val * b.weight) - (a.val * a.weight))
   .slice(0,3);

  // Fallback if no SF data
  const top3Issues = allIssues4.length >= 3 ? allIssues4 : [
    {title:"Run a site crawl",stat:"Upload Screaming Frog CSV",why:"A crawl reveals the exact issues holding back rankings.",fix:"Run Screaming Frog on the domain and upload the Crawl Overview CSV.",effort:"Low"},
    {title:"Check schema markup",stat:`${D.schemaErrors||0} known errors`,why:"Schema errors block rich results in search.",fix:"Validate schema using Google's Rich Results Test.",effort:"Low"},
    {title:"Review metadata",stat:`${totalPages4} pages to audit`,why:"Metadata quality directly impacts click-through rates.",fix:"Audit title tags and meta descriptions for length and uniqueness.",effort:"Med"},
  ];

  const pColors4 = [C.red, C.lightBlue, C.emerald];
  top3Issues.forEach((p,i) => {
    const x = 0.4 + i * 3.1;
    // Card
    s4.addShape(pres.shapes.RECTANGLE,{x,y:1.82,w:2.9,h:3.32,fill:{color:C.white},shadow:ms(),line:{color:"E2EAF0",width:0.5}});
    // Colored header bar
    s4.addShape(pres.shapes.RECTANGLE,{x,y:1.82,w:2.9,h:0.7,fill:{color:pColors4[i]},line:{color:pColors4[i],width:0}});
    // Number
    s4.addText(`0${i+1}`,{x:x+0.1,y:1.84,w:2.7,h:0.66,fontSize:26,bold:true,color:C.white,fontFace:"Calibri",align:"right",margin:0});
    // Issue title
    s4.addText(p.title,{x:x+0.18,y:2.6,w:2.55,h:0.38,fontSize:11,bold:true,color:C.darkBlue,fontFace:"Calibri",margin:0});
    // Stat in brand color
    s4.addText(p.stat,{x:x+0.18,y:3.0,w:2.55,h:0.26,fontSize:10,bold:true,color:pColors4[i],fontFace:"Calibri",margin:0});
    // Why it matters
    s4.addText(p.why,{x:x+0.18,y:3.28,w:2.55,h:0.7,fontSize:9,color:C.dark,fontFace:"Calibri",margin:0});
    // Fix
    s4.addShape(pres.shapes.RECTANGLE,{x:x+0.18,y:4.0,w:2.55,h:0.7,fill:{color:C.offWhite},line:{color:"E2EAF0",width:0.3}});
    s4.addText("FIX:",{x:x+0.26,y:4.04,w:0.4,h:0.62,fontSize:8,bold:true,color:pColors4[i],fontFace:"Calibri",valign:"top",margin:0});
    s4.addText(p.fix,{x:x+0.62,y:4.04,w:2.0,h:0.62,fontSize:8,color:C.dark,fontFace:"Calibri",valign:"top",margin:0});
    // Effort tag
    s4.addShape(pres.shapes.RECTANGLE,{x:x+0.18,y:4.78,w:2.55,h:0.24,fill:{color:C.offWhite},line:{color:"E2EAF0",width:0}});
    s4.addText(`Effort: ${p.effort||"Med"}  ·  High impact`,{x:x+0.18,y:4.78,w:2.55,h:0.24,fontSize:8,bold:true,color:C.midGray,fontFace:"Calibri",align:"center",margin:0});
  });
  footer(s4,D);

  // S5 PAGE SPEED
  const s5=pres.addSlide(); s5.background={color:C.white};
  slbl(s5,"SPEED & CORE WEB VITALS · Google PageSpeed Insights");
  stit(s5,"Slow pages lose clients before they read a word.");
  s5.addText("Google uses page speed as a direct ranking factor. For law firms, every second of delay costs consultations.",{x:0.5,y:1.38,w:9,h:0.35,fontSize:11,color:C.dark,fontFace:"Calibri"});
  await drawGauge(pres,s5, 0.4,1.8,2.3,2.5, D.psMobile||D.psPerformance||0, "Mobile");
  await drawGauge(pres,s5, 2.9,1.8,2.3,2.5, D.psDesktop||0, "Desktop");
  s5.addText("Most law firm searches happen on mobile — this score must be above 90.",{x:0.5,y:4.45,w:4.5,h:0.4,fontSize:9,color:C.midGray,fontFace:"Calibri",italic:true,align:"center"});
  const vitals=[
    {label:"First Contentful Paint",   value:D.psFCP||"—", good:"< 1.8s"},
    {label:"Largest Contentful Paint", value:D.psLCP||"—", good:"< 2.5s"},
    {label:"Total Blocking Time",      value:D.psTBT||"—", good:"< 200ms"},
    {label:"Cumulative Layout Shift",  value:D.psCLS||"—", good:"< 0.1"},
  ];
  s5.addText("CORE WEB VITALS",{x:5.3,y:1.78,w:4.3,h:0.22,fontSize:9,bold:true,color:C.lightBlue,charSpacing:3,fontFace:"Calibri"});
  vitals.forEach((v,i)=>{
    const y=2.1+i*0.82;
    // Determine pass/fail color based on value vs goal threshold
    function vitalColor(val, good) {
      if (!val || val === "—") return C.midGray;
      const num = parseFloat(String(val).replace(/[^0-9.]/g,""));
      if (isNaN(num)) return C.midGray;
      const threshold = parseFloat(String(good).replace(/[^0-9.]/g,""));
      if (isNaN(threshold)) return C.midGray;
      return num <= threshold ? C.emerald : C.red;
    }
    const vc = vitalColor(v.value, v.good);
    s5.addShape(pres.shapes.RECTANGLE,{x:5.3,y,w:4.3,h:0.7,fill:{color:C.offWhite},shadow:ms(),line:{color:"E2EAF0",width:0.3}});
    s5.addShape(pres.shapes.RECTANGLE,{x:5.3,y,w:0.06,h:0.7,fill:{color:vc},line:{color:vc,width:0}});
    s5.addText(v.label,{x:5.48,y:y+0.08,w:2.4,h:0.25,fontSize:10,bold:true,color:C.darkBlue,fontFace:"Calibri",margin:0});
    s5.addText(`Goal: ${v.good}`,{x:5.48,y:y+0.35,w:2.4,h:0.22,fontSize:9,color:C.midGray,fontFace:"Calibri",margin:0});
    s5.addText(String(v.value),{x:8.5,y:y+0.15,w:1.0,h:0.35,fontSize:18,bold:true,color:vc,fontFace:"Calibri",align:"right",margin:0});
  });
  footer(s5,D);

  // S6 LOCAL SEO WITH BENCHMARKS + COMPETITORS
  const s6=pres.addSlide(); s6.background={color:C.white};
  slbl(s6,"LOCAL SEO SNAPSHOT · BrightLocal + Majestic + SEMrush");
  stit(s6,"Local search is where law firm clients start.");

  // Comparison table
  // Merge SEMrush competitors with Majestic competitor backlink data
  const semrushComps   = D.competitors || [];
  const majesticComps  = D.majesticCompetitors || [];
  // Merge: try to match by domain, fallback to index order
  const competitors = semrushComps.map((c,i) => {
    const domainMatch = majesticComps.find(m => m.domain && c.domain && m.domain.replace(/^www\./,"") === c.domain.replace(/^www\./,""));
    return { ...c, ...(domainMatch || majesticComps[i] || {}) };
  });
  // If no SEMrush comps but have Majestic comps, use those directly
  if (!competitors.length && majesticComps.length) {
    competitors.push(...majesticComps);
  }
  const hasComps = competitors.length > 0;

  // Header row
  const colW   = hasComps ? 2.1 : 3.0;
  const cols   = hasComps
    ? ["METRIC","YOUR SITE","COMPETITOR 1","COMPETITOR 2","BENCHMARK"]
    : ["METRIC","YOUR SITE","BENCHMARK","TARGET"];
  const colXs  = hasComps
    ? [0.4, 2.55, 4.7, 6.85, 8.5]
    : [0.4, 3.5,  6.0, 8.0];

  // Draw header
  s6.addShape(pres.shapes.RECTANGLE,{x:0.4,y:1.5,w:9.2,h:0.35,fill:{color:C.darkBlue},line:{color:C.darkBlue,width:0}});
  cols.forEach((c,i)=>{
    s6.addText(c,{x:colXs[i],y:1.52,w:colW,h:0.3,fontSize:8,bold:true,color:C.white,fontFace:"Calibri",margin:0});
  });

  const rows = [
    { label:"Trust Flow",         client:D.trustFlow,        comp1:competitors[0]?.trustFlow,    comp2:competitors[1]?.trustFlow,    bench:BENCHMARKS.trustFlow },
    { label:"Citation Flow",      client:D.citationFlow,     comp1:competitors[0]?.citationFlow, comp2:competitors[1]?.citationFlow, bench:BENCHMARKS.citationFlow },
    { label:"Referring Domains",  client:D.referringDomains, comp1:competitors[0]?.referringDomains,comp2:competitors[1]?.referringDomains,bench:BENCHMARKS.referringDomains },
    { label:"Citation Score",      client:D.keyCitationScore!=null?`${D.keyCitationScore}/100`:"—", comp1:null, comp2:null, bench:{label:"70+", note:"Competitive"}, suffix:"" },
    { label:"Citations Found",     client:(parseInt(D.citationsFound)||0) > 0 ? `${parseInt(D.citationsFound)} of ${parseInt(D.citationsTotal)||"?"}` : "—", comp1:null, comp2:null, bench:BENCHMARKS.citationsFound },
    { label:"NAP Errors",          client:D.napErrors!=null?String(D.napErrors):"—", comp1:null, comp2:null, bench:{label:"0",note:"All listings accurate"}, suffix:"" },
    
    { label:"Local Rank Avg",     client:`#${D.localRankAvg||"—"}`, comp1:competitors[0]?.localRank?`#${competitors[0].localRank}`:null,comp2:competitors[1]?.localRank?`#${competitors[1].localRank}`:null,bench:BENCHMARKS.localRankAvg },
  ];

  rows.forEach((row,i)=>{
    const y=1.88+i*0.55;
    const bg=i%2===0?C.white:C.offWhite;
    s6.addShape(pres.shapes.RECTANGLE,{x:0.4,y,w:9.2,h:0.52,fill:{color:bg},line:{color:"E2EAF0",width:0.3}});
    s6.addText(row.label,{x:colXs[0],y:y+0.1,w:2.0,h:0.32,fontSize:10,bold:true,color:C.darkBlue,fontFace:"Calibri",margin:0});
    const clientColor=scoreColor(parseFloat(String(row.client)),row.bench);
    s6.addText(String(row.client||"—"),{x:colXs[1],y:y+0.1,w:colW,h:0.32,fontSize:12,bold:true,color:clientColor,fontFace:"Calibri",margin:0});
    if(hasComps){
      s6.addText(String(row.comp1||"—"),{x:colXs[2],y:y+0.1,w:colW,h:0.32,fontSize:11,color:C.midGray,fontFace:"Calibri",margin:0});
      s6.addText(String(row.comp2||"—"),{x:colXs[3],y:y+0.1,w:colW,h:0.32,fontSize:11,color:C.midGray,fontFace:"Calibri",margin:0});
      s6.addText(row.bench?.label||"—",{x:colXs[4],y:y+0.1,w:1.5,h:0.32,fontSize:10,bold:true,color:C.emerald,fontFace:"Calibri",margin:0});
    } else {
      s6.addText(row.bench?.label||"—",{x:colXs[2],y:y+0.1,w:colW,h:0.32,fontSize:11,bold:true,color:C.emerald,fontFace:"Calibri",margin:0});
      s6.addText(row.bench?.note||"",  {x:colXs[3],y:y+0.1,w:colW,h:0.32,fontSize:9,color:C.midGray,fontFace:"Calibri",italic:true,margin:0});
    }
  });

  // Legend
  s6.addShape(pres.shapes.RECTANGLE,{x:0.4,y:5.2,w:9.2,h:0.15,fill:{color:C.offWhite},line:{color:"E2EAF0",width:0}});
  s6.addShape(pres.shapes.RECTANGLE,{x:0.5,y:5.23,w:0.18,h:0.08,fill:{color:C.emerald},line:{color:C.emerald,width:0}});
  s6.addText("At/above benchmark",{x:0.72,y:5.21,w:2.2,h:0.16,fontSize:7,color:C.midGray,fontFace:"Calibri"});
  s6.addShape(pres.shapes.RECTANGLE,{x:3.0,y:5.23,w:0.18,h:0.08,fill:{color:C.lightBlue},line:{color:C.lightBlue,width:0}});
  s6.addText("Close to benchmark",{x:3.22,y:5.21,w:2.0,h:0.16,fontSize:7,color:C.midGray,fontFace:"Calibri"});
  s6.addShape(pres.shapes.RECTANGLE,{x:5.4,y:5.23,w:0.18,h:0.08,fill:{color:C.red},line:{color:C.red,width:0}});
  s6.addText("Below benchmark",{x:5.62,y:5.21,w:2.0,h:0.16,fontSize:7,color:C.midGray,fontFace:"Calibri"});
  footer(s6,D);

  // S7 CONTENT AUTHORITY + COMPETITOR GAP
  const s7=pres.addSlide(); s7.background={color:C.white};
  slbl(s7,"CONTENT AUTHORITY · SEMrush Competitor Gap Analysis");
  const OA  = D.organicAudit || {};
  const SF  = D.sf || {};
  const CG  = D.contentGap || null;
  const hasSF = Object.keys(SF).length > 0;
  const hasCG = CG && CG.clustered && CG.clustered.length > 0;

  // Dynamic headline based on data
  const s7Pages = SF.htmlPages || D.pagesCrawled || 0;
  const s7Gaps  = hasCG ? CG.clustered.length : 0;
  const s7Title = hasCG && s7Pages
    ? `${s7Pages} pages. Competitors own ${s7Gaps} topic${s7Gaps!==1?"s":""} you don't.`
    : s7Pages
    ? `${s7Pages} pages crawled. Here's what the content tells us.`
    : "Content is there. But is it working for you?";
  stit(s7, s7Title);

  // ── LEFT PANEL: Site architecture + keyword rankings ──────────
  s7.addShape(pres.shapes.RECTANGLE,{x:0.4,y:1.72,w:2.9,h:3.62,fill:{color:C.offWhite},shadow:ms(),line:{color:"E2EAF0",width:0.3}});

  // Keyword ranking chips
  s7.addText("KEYWORD RANKINGS",{x:0.5,y:1.82,w:2.7,h:0.2,fontSize:8,bold:true,color:C.lightBlue,charSpacing:2,fontFace:"Calibri"});
  [{label:"Top 3",val:OA.top3||0,color:C.emerald},{label:"Top 10",val:OA.top10||0,color:C.lightBlue},{label:"Top 20",val:OA.top20||0,color:C.midGray}]
  .forEach((r,i)=>{
    const x=0.5+i*0.92;
    s7.addShape(pres.shapes.RECTANGLE,{x,y:2.06,w:0.82,h:0.64,fill:{color:C.white},line:{color:"E2EAF0",width:0.3}});
    s7.addText(String(r.val),{x,y:2.1,w:0.82,h:0.36,fontSize:18,bold:true,color:r.color,fontFace:"Calibri",align:"center",margin:0});
    s7.addText(r.label,{x,y:2.46,w:0.82,h:0.2,fontSize:8,color:C.midGray,fontFace:"Calibri",align:"center",margin:0});
  });

  // Page architecture from Screaming Frog
  s7.addText("PAGE ARCHITECTURE",{x:0.5,y:2.82,w:2.7,h:0.2,fontSize:8,bold:true,color:C.lightBlue,charSpacing:2,fontFace:"Calibri"});
  const dd = SF.depthDist || {};
  const totalHtml7 = SF.htmlPages || 1;
  [
    {label:"Hub pages (depth 1)",  val:dd.d1||0,     note:"Top-level service/practice pages"},
    {label:"Sub-hub (depth 2)",     val:dd.d2||0,     note:"Topic clusters & subtopics"},
    {label:"Spoke pages (depth 3)", val:dd.d3||0,     note:"Supporting & geo-targeted content"},
    {label:"Deep / blog (4+)",      val:dd.d4plus||0, note:"Long-tail, Q&A, blog posts"},
  ].forEach((r,i)=>{
    const y=3.06+i*0.52;
    const barW = Math.max(0.04, Math.min(1.9, 1.9*(r.val/totalHtml7)));
    const pct  = Math.round((r.val/totalHtml7)*100);
    s7.addShape(pres.shapes.RECTANGLE,{x:0.5,y,w:2.7,h:0.44,fill:{color:C.white},line:{color:"E2EAF0",width:0.3}});
    s7.addText(String(r.val),{x:0.56,y:y+0.04,w:0.42,h:0.36,fontSize:14,bold:true,color:C.darkBlue,fontFace:"Calibri",valign:"middle",margin:0});
    s7.addText(r.label,{x:1.02,y:y+0.04,w:1.7,h:0.2,fontSize:8,bold:true,color:C.darkBlue,fontFace:"Calibri",margin:0});
    s7.addShape(pres.shapes.RECTANGLE,{x:1.02,y:y+0.28,w:1.9,h:0.1,fill:{color:"E2EAF0"},line:{color:"E2EAF0",width:0}});
    if(barW>0.04) s7.addShape(pres.shapes.RECTANGLE,{x:1.02,y:y+0.28,w:barW,h:0.1,fill:{color:C.lightBlue},line:{color:C.lightBlue,width:0}});
    s7.addText(`${pct}%`,{x:2.94,y:y+0.24,w:0.2,h:0.18,fontSize:7,color:C.midGray,fontFace:"Calibri",align:"right",margin:0});
  });

  // ── RIGHT PANEL: Competitor content gap ───────────────────────
  s7.addShape(pres.shapes.RECTANGLE,{x:3.5,y:1.72,w:6.1,h:3.62,fill:{color:C.offWhite},shadow:ms(),line:{color:"E2EAF0",width:0.3}});

  if (hasCG) {
    // Competitor domains header
    const compNames = (CG.competitorDomains||[]).join("  ·  ");
    s7.addText("COMPETITOR CONTENT GAP",{x:3.6,y:1.82,w:4.5,h:0.2,fontSize:8,bold:true,color:C.lightBlue,charSpacing:2,fontFace:"Calibri"});
    s7.addText(`vs. ${compNames}`,{x:3.6,y:2.04,w:5.9,h:0.18,fontSize:8,color:C.midGray,fontFace:"Calibri",margin:0});

    // Summary chips
    s7.addShape(pres.shapes.RECTANGLE,{x:7.6,y:1.8,w:1.0,h:0.42,fill:{color:C.red},line:{color:C.red,width:0}});
    s7.addText(String(CG.totalGapKeywords||0),{x:7.6,y:1.82,w:1.0,h:0.22,fontSize:14,bold:true,color:C.white,fontFace:"Calibri",align:"center",margin:0});
    s7.addText("gap kws",{x:7.6,y:2.02,w:1.0,h:0.18,fontSize:7,color:C.white,fontFace:"Calibri",align:"center",margin:0});

    // Gap cluster rows
    const displayGaps = CG.clustered.slice(0,7);
    displayGaps.forEach((g,i)=>{
      const y=2.3+i*0.48;
      const maxVol = CG.clustered[0].volume || 1;
      const barW   = Math.max(0.1, Math.min(3.6, 3.6*(g.volume/maxVol)));
      const isTop  = i === 0;

      s7.addShape(pres.shapes.RECTANGLE,{x:3.6,y,w:5.9,h:0.42,fill:{color:isTop?C.darkBlue:C.white},line:{color:isTop?C.darkBlue:"E2EAF0",width:0.3}});

      // Topic label
      s7.addText(g.label,{x:3.7,y:y+0.06,w:1.8,h:0.3,fontSize:10,bold:isTop,color:isTop?C.white:C.darkBlue,fontFace:"Calibri",valign:"middle",margin:0});

      // Volume bar
      s7.addShape(pres.shapes.RECTANGLE,{x:5.6,y:y+0.16,w:3.6,h:0.12,fill:{color:isTop?"4A90D9":"E2EAF0"},line:{color:"none",width:0}});
      s7.addShape(pres.shapes.RECTANGLE,{x:5.6,y:y+0.16,w:barW,h:0.12,fill:{color:isTop?C.white:C.red},line:{color:"none",width:0}});

      // Volume label
      const volLabel = g.volume >= 1000 ? `${Math.round(g.volume/1000)}k` : String(g.volume);
      s7.addText(`${volLabel}/mo`,{x:9.2,y:y+0.08,w:0.28,h:0.26,fontSize:8,bold:true,color:isTop?C.white:C.red,fontFace:"Calibri",align:"right",valign:"middle",margin:0});
    });

    // Legend
    s7.addShape(pres.shapes.RECTANGLE,{x:3.6,y:5.66,w:0.14,h:0.14,fill:{color:C.red},line:{color:C.red,width:0}});
    s7.addText("Monthly search volume competitors capture — you don't",{x:3.78,y:5.64,w:5.7,h:0.18,fontSize:8,color:C.midGray,fontFace:"Calibri",margin:0});

  } else {
    // No gap data — show content quality metrics from SF instead
    s7.addText("CONTENT QUALITY AUDIT",{x:3.6,y:1.82,w:5.9,h:0.2,fontSize:8,bold:true,color:C.lightBlue,charSpacing:2,fontFace:"Calibri"});
    const issues7 = [
      {label:"Titles too long",     val:SF.titlesOver60||0,    total:totalHtml7},
      {label:"Duplicate meta desc", val:SF.metaDuplicate||0,   total:totalHtml7},
      {label:"Meta desc too long",  val:SF.metaOver155||0,     total:totalHtml7},
      {label:"Low content pages",   val:SF.lowContentPages||0, total:totalHtml7},
      {label:"Hard to read",        val:SF.readabilityHard||0, total:totalHtml7},
      {label:"Missing alt text",    val:SF.missingAltText||0,  total:totalHtml7},
      {label:"Missing canonicals",  val:SF.missingCanonical||0,total:totalHtml7},
      {label:"Redirects",           val:SF.redirects3xx||0,    total:SF.totalUrlsCrawled||0},
      {label:"4xx errors",          val:SF.errors4xx||0,       total:SF.totalUrlsCrawled||0},
    ];
    issues7.forEach((iss,i)=>{
      const col=i%3,row7=Math.floor(i/3);
      const ix=3.6+col*1.98,iy=2.06+row7*0.96;
      const tc=iss.val===0?C.emerald:iss.val>(iss.total*0.3)?C.red:"F5A623";
      s7.addShape(pres.shapes.RECTANGLE,{x:ix,y:iy,w:1.88,h:0.86,fill:{color:C.white},shadow:ms(),line:{color:"E2EAF0",width:0.3}});
      s7.addShape(pres.shapes.RECTANGLE,{x:ix,y:iy,w:1.88,h:0.04,fill:{color:tc},line:{color:tc,width:0}});
      s7.addText(String(iss.val),{x:ix+0.08,y:iy+0.08,w:0.6,h:0.44,fontSize:22,bold:true,color:tc,fontFace:"Calibri",valign:"middle",margin:0});
      s7.addText(iss.label,{x:ix+0.08,y:iy+0.54,w:1.7,h:0.28,fontSize:8,color:C.midGray,fontFace:"Calibri",margin:0});
    });
  }

  // Insight bar
  const s7Insight = hasCG
    ? `⚠️  Competitors rank for ${CG.totalGapKeywords} keywords you don't — top gap is "${(CG.clustered[0]||{}).label||""}" with ~${((CG.clustered[0]||{}).volume||0).toLocaleString()} monthly searches at stake.`
    : !hasSF
    ? "Upload a Screaming Frog Crawl Overview CSV for detailed content analysis. Add competitor domains to enable gap analysis."
    : `📊  ${SF.htmlPages||0} HTML pages crawled — ${(SF.titlesOver60||0)+(SF.metaDuplicate||0)+(SF.lowContentPages||0)} content issues identified. Add SEMrush competitors to unlock gap analysis.`;
  s7.addShape(pres.shapes.RECTANGLE,{x:0.4,y:5.42,w:9.2,h:0.42,fill:{color:C.darkBlue},line:{color:C.darkBlue,width:0}});
  s7.addText(s7Insight,{x:0.55,y:5.44,w:8.9,h:0.38,fontSize:9,color:C.white,fontFace:"Calibri",valign:"middle"});
  footer(s7,D);

  // S8 PRIORITY RECOMMENDATIONS
  const s8=pres.addSlide(); s8.background={color:C.white};
  slbl(s8,"PRIORITY RECOMMENDATIONS"); stit(s8,"Ordered by client impact, not effort.");
  const acts=(narrative.actions||[]).slice(0,6);
  const aColors=[C.red,C.lightBlue,C.lightBlue,C.emerald,C.emerald,C.midGray];
  acts.forEach((a,i)=>{
    const col=i%2,row=Math.floor(i/2),x=0.4+col*4.85,y=1.82+row*1.08;
    s8.addShape(pres.shapes.RECTANGLE,{x,y,w:4.65,h:0.96,fill:{color:C.white},shadow:ms(),line:{color:"E2EAF0",width:0.4}});
    s8.addShape(pres.shapes.RECTANGLE,{x,y,w:0.5,h:0.96,fill:{color:aColors[i]||C.midGray},line:{color:aColors[i]||C.midGray,width:0}});
    s8.addText(a.n,    {x,y,w:0.5,h:0.96,fontSize:20,bold:true,color:C.white,fontFace:"Calibri",align:"center",valign:"middle"});
    s8.addText(a.title,{x:x+0.6,y:y+0.07,w:3.9,h:0.26,fontSize:11,bold:true,color:C.darkBlue,fontFace:"Calibri",margin:0});
    s8.addText(a.body, {x:x+0.6,y:y+0.35,w:3.9,h:0.36,fontSize:9,color:C.dark,fontFace:"Calibri",margin:0});
    s8.addText(`Impact: ${a.impact||""}   Effort: ${a.effort||""}`,{x:x+0.6,y:y+0.74,w:3.9,h:0.18,fontSize:8,bold:true,color:C.midGray,fontFace:"Calibri",margin:0});
  });
  footer(s8,D);

  // S9 SEQUENCE — priority cards, no timelines
  const s9=pres.addSlide(); s9.background={color:C.darkBlue};
  s9.addShape(pres.shapes.RECTANGLE,{x:0,y:0,w:0.22,h:5.625,fill:{color:C.lightBlue},line:{color:C.lightBlue,width:0}});
  s9.addShape(pres.shapes.RECTANGLE,{x:0,y:5.1,w:10,h:0.525,fill:{color:"0C1E3A"},line:{color:"0C1E3A",width:0}});
  s9.addShape(pres.shapes.OVAL,{x:0.5,y:0.55,w:0.5,h:0.5,fill:{color:C.lightBlue},line:{color:C.lightBlue,width:0}});
  s9.addText("▶",{x:0.5,y:0.6,w:0.5,h:0.4,fontSize:16,color:C.darkBlue,fontFace:"Calibri",align:"center"});
  s9.addText("Recommended Next Steps",{x:0.5,y:1.15,w:9,h:0.55,fontSize:34,bold:true,color:C.white,fontFace:"Calibri"});
  s9.addText("Ordered by impact. Quick wins first, sustainable growth last.",{x:0.5,y:1.76,w:9,h:0.3,fontSize:12,color:"7ABCD4",fontFace:"Calibri",italic:true});
  const seq=narrative.sequence||[];
  const seqColors=[C.lightBlue,C.lightBlue,C.emerald,C.banana];
  const seqIcons=["1","2","3","4"];
  seq.forEach((t,i)=>{
    const col=i%2, row=Math.floor(i/2);
    const x=0.4+col*4.85, y=2.2+row*1.35;
    s9.addShape(pres.shapes.RECTANGLE,{x,y,w:4.65,h:1.18,fill:{color:"0F2040"},shadow:{type:"outer",blur:8,offset:2,angle:135,color:"000000",opacity:0.2},line:{color:"1A3A60",width:0.5}});
    s9.addShape(pres.shapes.RECTANGLE,{x,y,w:0.55,h:1.18,fill:{color:seqColors[i]||C.lightBlue},line:{color:seqColors[i]||C.lightBlue,width:0}});
    s9.addText(seqIcons[i],{x,y,w:0.55,h:1.18,fontSize:22,bold:true,color:C.darkBlue,fontFace:"Calibri",align:"center",valign:"middle"});
    s9.addText(t.action||t.week||"",{x:x+0.65,y:y+0.1,w:3.85,h:0.32,fontSize:12,bold:true,color:C.white,fontFace:"Calibri",margin:0});
    s9.addText(t.why||t.body||"",  {x:x+0.65,y:y+0.46,w:3.85,h:0.58,fontSize:10,color:"A8C4D8",fontFace:"Calibri",margin:0});
  });
  s9.addText(`${D.preparedBy}   ·   ${D.domain}   ·   ${D.date||""}`,{x:0.5,y:5.15,w:9,h:0.25,fontSize:8,color:"3A6080",fontFace:"Calibri"});

  return await pres.write({outputType:"nodebuffer"});
}

// ── ROUTES ────────────────────────────────────────────────────

// GET /debug — shows env config status (no secret values)
app.get("/debug", (req, res) => res.redirect("/version"));

app.get("/v", (req, res) => res.redirect("/version"));
app.get("/status", (req, res) => res.redirect("/version"));

app.get("/version", (req, res) => {
  const fs = require("fs");
  const path = require("path");

  // Read own source to extract function names and slide list
  let src = "";
  try { src = fs.readFileSync(__filename, "utf8"); } catch(e) {}

  // Extract top-level async/regular functions
  const fnMatches = [...src.matchAll(/^(?:async )?function (\w+)/gm)].map(m => m[1]);

  // Extract slide comments (// S1 COVER etc.)
  const slides = [...src.matchAll(/\/\/ (S\d+ [A-Z \+&·']+)/g)].map(m => m[1].trim());

  // Extract app routes
  const routes = [...src.matchAll(/app\.(get|post|put|delete)\("([^"]+)"/g)].map(m => `${m[1].toUpperCase()} ${m[2]}`);

  // Check which optional modules are present
  const modules = ["parse_screaming_frog","parse_majestic","renderGaugePng"].reduce((acc, mod) => {
    acc[mod] = src.includes(mod) ? "✅ present" : "❌ missing";
    return acc;
  }, {});

  // Check which data sources are wired up
  const dataSources = {
    semrush_site_audit:      src.includes("semrushSiteAuditData")  ? "✅" : "❌",
    semrush_organic:         src.includes("semrushOrganicKeywords") ? "✅" : "❌",
    semrush_competitors:     src.includes("semrushCompetitors")     ? "✅" : "❌",
    semrush_competitor_gap:  src.includes("semrushCompetitorGap")   ? "✅" : "❌",
    brightlocal_citations:   src.includes("brightlocalCitationAudit") ? "✅" : "❌",
    brightlocal_auto_lookup: src.includes("v2/ct/get-all")         ? "✅" : "❌",
    pagespeed:               src.includes("fetchPageSpeed")         ? "✅" : "❌",
    screaming_frog_parser:   src.includes("parseScreamingFrogCsv") ? "✅" : "❌",
    majestic_parser:         src.includes("parseMajesticCsv")      ? "✅" : "❌",
    gauge_png_renderer:      src.includes("renderGaugePng")        ? "✅" : "❌",
    binary_pptx_download:    src.includes("/download/:id")         ? "✅" : "❌",
  };

  // Key fixes — check for known patterns
  const fixes = {
    brightlocal_scope_bug_fixed:  src.includes("async function brightlocalCitationAudit(domain, businessName, reportId, locationId)") ? "✅" : "❌",
    content_disposition_ascii:    src.includes("asciiFallback")    ? "✅" : "❌",
    revoke_object_url_delayed:    src.includes("setTimeout")       ? "✅" : "❌",
    actions_capped_at_6:          src.includes("slice(0,6)")       ? "✅" : "❌",
    semrush_merge_empty_string:   src.includes('data[k] === ""')   ? "✅" : "❌",
  };

  res.json({
    app:           "SEO Audit Generator — PMP Marketing Group",
    deployed_at:   new Date().toISOString(),
    node_version:  process.version,
    server_file:   __filename,

    api_keys: {
      semrush:      SEMRUSH_KEY     ? `✅ set (${SEMRUSH_KEY.slice(0,4)}...)` : "❌ NOT SET",
      brightlocal:  BRIGHTLOCAL_KEY ? `✅ set (${BRIGHTLOCAL_KEY.slice(0,4)}...)` : "❌ NOT SET",
      anthropic:    ANTHROPIC_KEY   ? `✅ set` : "❌ NOT SET",
      pagespeed:    PAGESPEED_KEY   ? `✅ set` : "❌ NOT SET",
    },

    config: {
      semrush_project_env: SEMRUSH_PROJECT || "NOT SET — must come from form",
      brightlocal_id:      "form only — no hardcoded fallback",
      pptx_model:          src.match(/model:"([^"]+)"/)?.[1] || "unknown",
    },

    slides:      slides,
    routes:      routes,
    functions:   fnMatches,
    modules,
    data_sources: dataSources,
    known_fixes:  fixes,
  });
});

// POST /generate — accepts multipart, starts background job, returns jobId immediately
app.post("/generate", upload.fields([
  { name: "majesticCsv",      maxCount: 1 },
  { name: "screamingFrogCsv", maxCount: 1 },
]), (req, res) => {
  const jobId = createJob();

  // Parse files into buffers now (before async work)
  const majesticBuf      = req.files?.majesticCsv?.[0]?.buffer      || null;
  const screamingFrogBuf = req.files?.screamingFrogCsv?.[0]?.buffer || null;
  const data = JSON.parse(req.body.data || "{}");

  // Respond immediately with job ID
  res.json({ jobId });

  // Run everything in background — no timeout risk
  (async () => {
    try {
      // Parse CSV uploads
      if (majesticBuf) {
        const majesticData = parseMajesticCsv(majesticBuf);
        // Pull out competitor data before merging
        const majesticComps = majesticData.majesticCompetitors || [];
        delete majesticData.majesticCompetitors;
        Object.assign(data, majesticData);
        // Merge Majestic competitor backlink data into existing competitors
        if (majesticComps.length) {
          data.majesticCompetitors = majesticComps;
        }
      }
      if (screamingFrogBuf) {
        const sf = parseScreamingFrogCsv(screamingFrogBuf);
        Object.assign(data, sf);
        if (!data.pagesCrawled  && sf.sfPagesCrawled)  data.pagesCrawled  = sf.sfPagesCrawled;
        if (!data.missingDesc   && sf.sfMissingDesc)   data.missingDesc   = sf.sfMissingDesc;
        if (!data.titlesTooLong && sf.sfTitleTooLong)  data.titlesTooLong = sf.sfTitleTooLong;
        if (!data.missingH1     && sf.sfMissingH1)     data.missingH1     = sf.sfMissingH1;
      }

      // SEMrush Site Audit — use project ID from form or env variable
      const projectId = data.semrushProjectId || SEMRUSH_PROJECT;
      if (SEMRUSH_KEY && projectId) {
        console.log(`[${jobId}] Fetching SEMrush site audit data for project ${projectId}...`);
        const auditData = await semrushSiteAuditData(projectId);
        // Only fill in fields not already provided by staff
        Object.keys(auditData).forEach(k => {
          if (data[k] === undefined || data[k] === null || data[k] === 0 || data[k] === "" || data[k] === "0") {
            data[k] = auditData[k];
          }
        });
      }

      // PageSpeed
      if (data.domain && !data.psPerformance) {
        console.log(`[${jobId}] Fetching PageSpeed...`);
        Object.assign(data, await fetchPageSpeed(data.domain));
      }

      // SEMrush competitors
      if (data.domain && SEMRUSH_KEY) {
        console.log(`[${jobId}] Fetching SEMrush competitors...`);
        data.competitors = await semrushCompetitors(data.domain);
      }

      // SEMrush organic keyword + content audit
      if (data.domain && SEMRUSH_KEY) {
        console.log(`[${jobId}] Fetching SEMrush organic keywords for content audit...`);
        data.organicAudit = await semrushOrganicKeywords(data.domain);
        // Competitor content gap — uses competitor domains from semrushCompetitors
        if (data.competitors && data.competitors.length) {
          console.log(`[${jobId}] Fetching competitor content gap analysis...`);
          const compDomains = data.competitors.map(c => c.domain).filter(Boolean);
          data.contentGap = await semrushCompetitorGap(data.domain, compDomains);
        }
      }

      // BrightLocal citation audit
      if (data.domain && BRIGHTLOCAL_KEY && (data.citationsFound === undefined || data.citationsFound === null || data.citationsFound === "")) {
        console.log(`[${jobId}] Running BrightLocal audit...`);
        Object.assign(data, await brightlocalCitationAudit(data.domain, data.clientName, data.brightlocalReportId, data.brightlocalLocationId));
      }

      // Claude narrative
      console.log(`[${jobId}] Generating narrative...`);
      const narrative = await getNarrative(data);

      // Build PPTX
      console.log(`[${jobId}] Building PPTX...`);
      const pptxBuffer = await buildPptx(data, narrative);
      const date       = data.date || new Date().toLocaleDateString("en-US",{month:"short",year:"numeric"});
      const fileName   = `SEO Audit — ${data.clientName} — ${date}.pptx`;

      jobBuffers.set(jobId, pptxBuffer);
      setTimeout(() => jobBuffers.delete(jobId), 30 * 60 * 1000); // clean up after 30 min
      setJobDone(jobId, { fileName });
      console.log(`[${jobId}] Done.`);

    } catch(err) {
      console.error(`[${jobId}] Error:`, err.message);
      setJobError(jobId, err.message);
    }
  })();
});


// GET /brightlocal-reports?locationId=XXXX — find latest citation tracker report for a location
app.get("/brightlocal-reports", async (req, res) => {
  const locationId = req.query.locationId;
  if (!locationId) return res.status(400).json({ error: "locationId required" });
  if (!BRIGHTLOCAL_KEY) return res.status(500).json({ error: "BrightLocal API key not configured" });
  try {
    const BASE = "https://tools.brightlocal.com/seo-tools/api";
    // Correct endpoint: GET /v2/ct/get-all with optional location-id filter
    // Try filtered by location-id first, fall back to all reports and match manually
    const url = `${BASE}/v2/ct/get-all?api-key=${BRIGHTLOCAL_KEY}&location-id=${locationId}`;
    console.log("BrightLocal lookup URL:", url.replace(BRIGHTLOCAL_KEY, "***"));
    const r = await fetch(url);
    const d = await r.json();
    console.log("BrightLocal ct/get-all response:", JSON.stringify(d).slice(0, 500));
    let reports = d.response?.results || d.results || [];

    // If filtered call returned nothing, fetch ALL reports and match by location_id
    if (!reports.length) {
      console.log("BrightLocal: filtered call empty, fetching all reports...");
      const r2 = await fetch(`${BASE}/v2/ct/get-all?api-key=${BRIGHTLOCAL_KEY}`);
      const d2 = await r2.json();
      console.log("BrightLocal all reports:", JSON.stringify(d2).slice(0, 800));
      const all = d2.response?.results || d2.results || [];
      reports = all.filter(rep => String(rep.location_id) === String(locationId));
      console.log(`BrightLocal: found ${reports.length} report(s) matching location_id ${locationId} from ${all.length} total`);
      if (!reports.length) {
        // Show available location IDs to help diagnose typos
        const available = [...new Set(all.map(r => r.location_id).filter(Boolean))];
        return res.status(404).json({ 
          error: `No reports found for location ID ${locationId}. Available location IDs in your account: ${available.join(", ")}` 
        });
      }
    }

    const latest = reports.sort((a, b) => new Date(b.last_run || 0) - new Date(a.last_run || 0))[0];
    const reportId = latest.report_id || latest["report-id"];
    const reportName = latest.report_name || latest["report-name"] || "";
    const lastRun = latest.last_run || "";
    res.json({ reportId, reportName, lastRun });
  } catch(e) {
    console.error("BrightLocal reports lookup error:", e.message);
    res.status(500).json({ error: e.message });
  }
});

// GET /job/:id — poll for job status
app.get("/job/:id", (req, res) => {
  const job = getJob(req.params.id);
  if (!job) return res.status(404).json({ status: "not_found" });
  if (job.status === "done")  return res.json({ status: "done",  fileName: job.result.fileName });
  if (job.status === "error") return res.json({ status: "error", error: job.error });
  res.json({ status: "pending" });
});

// GET /download/:id — stream the PPTX binary directly (no base64)
app.get("/download/:id", (req, res) => {
  const job = getJob(req.params.id);
  const buf = jobBuffers.get(req.params.id);
  if (!job || job.status !== "done" || !buf) return res.status(404).json({ error: "Not found" });
  const fileName = job.result.fileName || "SEO-Audit.pptx";
  // ASCII fallback (strips special chars) + UTF-8 encoded full name
  const asciiFallback = fileName.replace(/[^ -~]/g, "-").replace(/"/g, "");
  const utf8Encoded   = encodeURIComponent(fileName);
  res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.presentationml.presentation");
  res.setHeader("Content-Disposition", `attachment; filename="${asciiFallback}"; filename*=UTF-8''${utf8Encoded}`);
  res.setHeader("Content-Length", buf.length);
  res.end(buf);
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`SEO Audit API v3 running on port ${PORT}`));
