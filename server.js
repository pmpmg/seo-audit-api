const express  = require("express");
const path     = require("path");
const pptxgen  = require("pptxgenjs");
const React    = require("react");
const ReactDOM = require("react-dom/server");
const sharp    = require("sharp");
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
const SEMRUSH_PROJECT = process.env.SEMRUSH_PROJECT_ID  || "6093881";
const BRIGHTLOCAL_KEY = process.env.BRIGHTLOCAL_API_KEY || "";
const ANTHROPIC_KEY   = process.env.ANTHROPIC_API_KEY   || "";

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

// BrightLocal citation audit
async function brightlocalCitationAudit(domain, businessName, location) {
  try {
    // Create audit
    const createRes = await fetch("https://tools.brightlocal.com/seo-tools/api/v4/lsrc/add", {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: new URLSearchParams({
        "api-key":       BRIGHTLOCAL_KEY,
        "business-name": businessName || domain,
        "website-address": `https://${domain}`,
        "country":       "USA",
        "city":          location || "",
        "postcode":      "",
      })
    });
    const createData = await createRes.json();
    if (!createData?.response?.["report-id"]) return {};

    const reportId = createData.response["report-id"];

    // Poll for results (max 90s)
    for (let i = 0; i < 18; i++) {
      await new Promise(r => setTimeout(r, 5000));
      const pollRes  = await fetch(`https://tools.brightlocal.com/seo-tools/api/v4/lsrc/get?api-key=${BRIGHTLOCAL_KEY}&report-id=${reportId}`);
      const pollData = await pollRes.json();
      if (pollData?.response?.["report-status"] === "Complete") {
        const r = pollData.response;
        return {
          citationsFound:  r["citations-found"]    || 0,
          napConsistency:  r["nap-consistency"]     || 0,
          activeListings:  r["active-listings"]     || 0,
          missingListings: r["missing-listings"]    || 0,
        };
      }
    }
    return {};
  } catch(e) { return {}; }
}

// SEMrush Site Audit — pulls full crawl data from existing campaign
async function semrushSiteAuditData(projectId) {
  try {
    // SEMrush Site Audit API v1 — correct endpoint format
    const key = SEMRUSH_KEY;

    // Step 1: Get campaign info + latest snapshot ID
    const infoUrl = `https://api.semrush.com/reports/v1/projects/${projectId}/siteaudit/info?key=${key}`;
    console.log("SEMrush info URL:", infoUrl);
    const infoRes  = await fetchWithTimeout(infoUrl, 15000);
    const infoText = await infoRes.text();
    console.log("SEMrush info response:", infoText.slice(0, 200));

    let infoJson = {};
    try { infoJson = JSON.parse(infoText); } catch(e) {
      console.error("SEMrush info parse error:", e.message);
      return {};
    }

    const snapshotId = infoJson?.data?.last_snapshot_id
                    || infoJson?.last_snapshot_id
                    || infoJson?.snapshot_id;

    if (!snapshotId) {
      console.log("No snapshot ID found in SEMrush response");
      return {};
    }
    console.log("Using snapshot ID:", snapshotId);

    // Step 2: Get audit summary for this snapshot
    const summaryUrl = `https://api.semrush.com/reports/v1/projects/${projectId}/siteaudit/snapshots/${snapshotId}/summary?key=${key}`;
    const summaryRes  = await fetchWithTimeout(summaryUrl, 15000);
    const summaryText = await summaryRes.text();
    console.log("SEMrush summary:", summaryText.slice(0, 300));

    let s = {};
    try { s = JSON.parse(summaryText)?.data || JSON.parse(summaryText) || {}; } catch(e) {}

    // Step 3: Get issues list
    const issuesUrl = `https://api.semrush.com/reports/v1/projects/${projectId}/siteaudit/snapshots/${snapshotId}/issues?key=${key}&limit=100`;
    const issuesRes  = await fetchWithTimeout(issuesUrl, 15000);
    const issuesText = await issuesRes.text();

    let issues = [];
    try {
      const ij = JSON.parse(issuesText);
      issues = ij?.data || ij?.issues || [];
    } catch(e) {}

    // Map issues to our fields
    const issueMap = {};
    issues.forEach(issue => {
      const name  = (issue.name || issue.id || issue.check_id || "").toLowerCase();
      const count = parseInt(issue.count || issue.pages_count || issue.errors || 0);
      if (name.includes("schema") || name.includes("structured"))      issueMap.schemaErrors   = (issueMap.schemaErrors   || 0) + count;
      if (name.includes("meta description") && name.includes("miss"))  issueMap.missingDesc    = (issueMap.missingDesc    || 0) + count;
      if (name.includes("title") && name.includes("long"))             issueMap.titlesTooLong  = (issueMap.titlesTooLong  || 0) + count;
      if (name.includes("alt") && name.includes("miss"))               issueMap.missingAlt     = (issueMap.missingAlt     || 0) + count;
      if (name.includes("thin") || name.includes("low word"))          issueMap.thinPages      = (issueMap.thinPages      || 0) + count;
      if (name.includes("broken") && name.includes("external"))        issueMap.brokenExternal = (issueMap.brokenExternal || 0) + count;
      if (name.includes("h1") && name.includes("miss"))                issueMap.missingH1      = (issueMap.missingH1      || 0) + count;
      if (name.includes("anchor") || name.includes("no anchor"))       issueMap.noAnchors      = (issueMap.noAnchors      || 0) + count;
    });

    return {
      pagesCrawled: parseInt(s.pages_crawled  || s.total_pages    || 0),
      siteHealth:   parseInt(s.health_score   || s.site_health    || 0),
      pages200:     parseInt(s.pages_with_200 || 0),
      redirects:    parseInt(s.pages_with_301 || s.pages_with_3xx || 0),
      errors:       parseInt(s.pages_with_4xx || 0),
      aiReadiness:  parseInt(s.ai_readiness_score || 0),
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
    const base = "https://www.googleapis.com/pagespeedonline/v5/runPagespeed";
    const url  = `https://${domain}`;
    let mob = {}, dsk = {};
    try {
      const r = await fetchWithTimeout(`${base}?url=${encodeURIComponent(url)}&strategy=mobile`, 12000);
      mob = await r.json();
    } catch(e) { console.log("PageSpeed mobile failed:", e.message); }
    try {
      const r = await fetchWithTimeout(`${base}?url=${encodeURIComponent(url)}&strategy=desktop`, 12000);
      dsk = await r.json();
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

Return ONLY valid JSON, no markdown:
{
  "executiveSummary": "One sentence — the single biggest opportunity on this site.",
  "whatIsWorking": [
    {"title":"Short title under 6 words","sub":"One sentence explanation based on the data."},
    {"title":"...","sub":"..."},
    {"title":"...","sub":"..."},
    {"title":"...","sub":"..."},
    {"title":"...","sub":"..."},
    {"title":"...","sub":"..."}
  ],
  "problems": [
    {"num":"01","title":"Short title under 6 words","stat":"Key metric from data","body":"2-3 sentences explaining the problem and its impact on client inquiries. No competitor names.","tag":"High impact · Low effort"},
    {"num":"02","title":"...","stat":"...","body":"...","tag":"High impact · Medium effort"},
    {"num":"03","title":"...","stat":"...","body":"...","tag":"High impact · High effort"}
  ],
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

// ── PPTX HELPERS ─────────────────────────────────────────────
const { FaServer,FaLink,FaHeading,FaRobot,FaCode,FaTags,
        FaBolt,FaFlag,FaMobile,FaDesktop,FaArrowRight,FaTrophy } = require("react-icons/fa");

async function iconPng(Icon,color,size=256){
  const svg=ReactDOM.renderToStaticMarkup(React.createElement(Icon,{color,size:String(size)}));
  return "image/png;base64,"+(await sharp(Buffer.from(svg)).png().toBuffer()).toString("base64");
}
async function gaugePng(score,size=280){
  const color=score>=90?C.emerald:score>=50?C.lightBlue:C.red;
  const r=100,cx=150,cy=150,sw=22,circ=2*Math.PI*r,dash=(score/100)*circ;
  const svg=`<svg xmlns="http://www.w3.org/2000/svg" width="${size}" height="${size}" viewBox="0 0 300 300">
    <circle cx="${cx}" cy="${cy}" r="${r}" fill="none" stroke="#E2EAF0" stroke-width="${sw}"/>
    <circle cx="${cx}" cy="${cy}" r="${r}" fill="none" stroke="#${color}" stroke-width="${sw}"
      stroke-dasharray="${dash} ${circ}" stroke-dashoffset="${circ/4}" stroke-linecap="round"/>
    <text x="150" y="158" text-anchor="middle" font-family="Arial" font-size="52" font-weight="bold" fill="#${color}">${score}</text>
    <text x="150" y="185" text-anchor="middle" font-family="Arial" font-size="16" fill="#6B7A8D">/100</text>
  </svg>`;
  return "image/png;base64,"+(await sharp(Buffer.from(svg)).resize(size,size).png().toBuffer()).toString("base64");
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
  kpi(pres,s2,7.36,1.95,2.15,2.1,`${D.napConsistency||0}%`,"NAP CONSISTENCY","Local citation accuracy",        scoreColor(D.napConsistency,BENCHMARKS.napConsistency));
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
  const wIcons=[FaServer,FaLink,FaHeading,FaRobot,FaCode,FaTrophy];
  for(let i=0;i<6;i++){
    const col=i%3,row=Math.floor(i/3),x=0.4+col*3.1,y=1.9+row*1.55;
    const ic=await iconPng(wIcons[i],"#"+C.emerald,256);
    s3.addShape(pres.shapes.RECTANGLE,{x,y,w:2.9,h:1.35,fill:{color:C.offWhite},shadow:ms(),line:{color:"E2EAF0",width:0.3}});
    s3.addShape(pres.shapes.RECTANGLE,{x,y,w:2.9,h:0.05,fill:{color:C.emerald},line:{color:C.emerald,width:0}});
    s3.addShape(pres.shapes.OVAL,{x:x+0.18,y:y+0.22,w:0.5,h:0.5,fill:{color:"CAE7D9"},line:{color:"CAE7D9",width:0}});
    s3.addImage({data:ic,x:x+0.22,y:y+0.27,w:0.42,h:0.42});
    s3.addText(wins[i]?.title||"",{x:x+0.82,y:y+0.18,w:1.95,h:0.32,fontSize:11,bold:true,color:C.darkBlue,fontFace:"Calibri",margin:0});
    s3.addText(wins[i]?.sub||"",  {x:x+0.82,y:y+0.52,w:1.95,h:0.55,fontSize:9,color:C.midGray,fontFace:"Calibri",margin:0});
  }
  footer(s3,D);

  // S4 PROBLEMS
  const s4=pres.addSlide(); s4.background={color:C.white};
  slbl(s4,"WHAT'S COSTING YOU CLIENTS"); stit(s4,"Three patterns, not a hundred problems.");
  s4.addText("Each of these is a template-level fix — meaning one change clears hundreds of issues at once.",{x:0.5,y:1.38,w:9,h:0.35,fontSize:11,color:C.dark,fontFace:"Calibri"});
  const probs=narrative.problems||[
    {num:"01",title:"Broken structured data",stat:`${D.schemaErrors||0} errors`,body:`Every page carries the same broken LocalBusiness schema. One template fix clears all ${D.schemaErrors||0}.`,tag:"High impact · Low effort",color:C.red},
    {num:"02",title:"Metadata gaps",stat:`${(D.missingDesc||0)+(D.titlesTooLong||0)} pages affected`,body:`${D.missingDesc||0} pages have no meta description. ${D.titlesTooLong||0} titles are cut off in search results.`,tag:"High impact · Medium effort",color:C.lightBlue},
    {num:"03",title:"Thin content & weak links",stat:`${D.thinPages||0} light pages`,body:`${D.thinPages||0} pages appear thin to search engines. ${D.noAnchors||0} internal links carry no anchor text.`,tag:"High impact · High effort",color:C.emerald},
  ];
  const pIcons=[FaCode,FaTags,FaBolt];
  const pColors=[C.red,C.lightBlue,C.emerald];
  for(let i=0;i<3;i++){
    const p=probs[i]||{}, x=0.4+i*3.1;
    const ic=await iconPng(pIcons[i],"#FFFFFF",256);
    s4.addShape(pres.shapes.RECTANGLE,{x,y:1.9,w:2.9,h:3.15,fill:{color:C.white},shadow:ms(),line:{color:"E2EAF0",width:0.5}});
    s4.addShape(pres.shapes.RECTANGLE,{x,y:1.9,w:2.9,h:0.68,fill:{color:pColors[i]},line:{color:pColors[i],width:0}});
    s4.addImage({data:ic,x:x+0.18,y:2.02,w:0.38,h:0.38});
    s4.addText(p.num||`0${i+1}`,{x:x+0.1,y:1.92,w:2.7,h:0.64,fontSize:26,bold:true,color:C.white,fontFace:"Calibri",align:"right",margin:0});
    s4.addText(p.title||"",{x:x+0.18,y:2.65,w:2.55,h:0.42,fontSize:11,bold:true,color:C.darkBlue,fontFace:"Calibri"});
    s4.addText(p.stat||"", {x:x+0.18,y:3.1, w:2.55,h:0.24,fontSize:10,bold:true,color:pColors[i],fontFace:"Calibri"});
    s4.addText(p.body||"", {x:x+0.18,y:3.36,w:2.55,h:1.18,fontSize:9,color:C.dark,fontFace:"Calibri"});
    s4.addShape(pres.shapes.RECTANGLE,{x:x+0.18,y:4.75,w:2.55,h:0.22,fill:{color:C.offWhite},line:{color:"E2EAF0",width:0}});
    s4.addText(p.tag||"",{x:x+0.18,y:4.72,w:2.55,h:0.22,fontSize:8,bold:true,color:C.midGray,fontFace:"Calibri",align:"center"});
  }
  footer(s4,D);

  // S5 PAGE SPEED
  const s5=pres.addSlide(); s5.background={color:C.white};
  slbl(s5,"SPEED & CORE WEB VITALS · Google PageSpeed Insights");
  stit(s5,"Slow pages lose clients before they read a word.");
  s5.addText("Google uses page speed as a direct ranking factor. For law firms, every second of delay costs consultations.",{x:0.5,y:1.38,w:9,h:0.35,fontSize:11,color:C.dark,fontFace:"Calibri"});
  const mg=await gaugePng(D.psMobile||D.psPerformance||0);
  const dg=await gaugePng(D.psDesktop||0);
  s5.addImage({data:mg,x:0.5,y:1.85,w:2.1,h:2.1});
  s5.addImage({data:dg,x:2.9,y:1.85,w:2.1,h:2.1});
  const mbi=await iconPng(FaMobile,"#"+C.darkBlue,128);
  const dki=await iconPng(FaDesktop,"#"+C.darkBlue,128);
  s5.addImage({data:mbi,x:0.88,y:4.05,w:0.32,h:0.32});
  s5.addImage({data:dki,x:3.28,y:4.05,w:0.32,h:0.32});
  s5.addText("Mobile", {x:0.5,y:4.08,w:2.1,h:0.28,fontSize:11,bold:true,color:C.darkBlue,fontFace:"Calibri",align:"center"});
  s5.addText("Desktop",{x:2.9,y:4.08,w:2.1,h:0.28,fontSize:11,bold:true,color:C.darkBlue,fontFace:"Calibri",align:"center"});
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
    s5.addShape(pres.shapes.RECTANGLE,{x:5.3,y,w:4.3,h:0.7,fill:{color:C.offWhite},shadow:ms(),line:{color:"E2EAF0",width:0.3}});
    s5.addShape(pres.shapes.RECTANGLE,{x:5.3,y,w:0.06,h:0.7,fill:{color:C.red},line:{color:C.red,width:0}});
    s5.addText(v.label,{x:5.48,y:y+0.08,w:2.4,h:0.25,fontSize:10,bold:true,color:C.darkBlue,fontFace:"Calibri",margin:0});
    s5.addText(`Goal: ${v.good}`,{x:5.48,y:y+0.35,w:2.4,h:0.22,fontSize:9,color:C.midGray,fontFace:"Calibri",margin:0});
    s5.addText(String(v.value),{x:8.5,y:y+0.15,w:1.0,h:0.35,fontSize:18,bold:true,color:C.red,fontFace:"Calibri",align:"right",margin:0});
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
  const competitors = semrushComps.map((c,i) => ({
    ...c,
    ...(majesticComps[i] || {}),
  }));
  // If no SEMrush comps but have Majestic comps, use those
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
    { label:"NAP Consistency",    client:`${D.napConsistency||0}%`, comp1:competitors[0]?.napConsistency?`${competitors[0].napConsistency}%`:null, comp2:null, bench:BENCHMARKS.napConsistency, suffix:"%" },
    { label:"Citations Found",    client:D.citationsFound,   comp1:null,                         comp2:null,                         bench:BENCHMARKS.citationsFound },
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

  // S7 RECOMMENDATIONS
  const s7=pres.addSlide(); s7.background={color:C.white};
  slbl(s7,"PRIORITY RECOMMENDATIONS"); stit(s7,"Ordered by client impact, not effort.");
  const acts=narrative.actions||[];
  const aColors=[C.red,C.lightBlue,C.lightBlue,C.emerald,C.emerald,C.midGray];
  acts.forEach((a,i)=>{
    const col=i%2,row=Math.floor(i/2),x=0.4+col*4.85,y=1.82+row*1.12;
    s7.addShape(pres.shapes.RECTANGLE,{x,y,w:4.65,h:1.0,fill:{color:C.white},shadow:ms(),line:{color:"E2EAF0",width:0.4}});
    s7.addShape(pres.shapes.RECTANGLE,{x,y,w:0.5,h:1.0,fill:{color:aColors[i]||C.midGray},line:{color:aColors[i]||C.midGray,width:0}});
    s7.addText(a.n,    {x,y,w:0.5,h:1.0,fontSize:20,bold:true,color:C.white,fontFace:"Calibri",align:"center",valign:"middle"});
    s7.addText(a.title,{x:x+0.6,y:y+0.08,w:3.9,h:0.28,fontSize:11,bold:true,color:C.darkBlue,fontFace:"Calibri",margin:0});
    s7.addText(a.body, {x:x+0.6,y:y+0.38,w:3.9,h:0.38,fontSize:9,color:C.dark,fontFace:"Calibri",margin:0});
    s7.addText(`Impact: ${a.impact||""}   Effort: ${a.effort||""}`,{x:x+0.6,y:y+0.78,w:3.9,h:0.18,fontSize:8,bold:true,color:C.midGray,fontFace:"Calibri",margin:0});
  });
  footer(s7,D);

  // S8 SEQUENCE — priority cards, no timelines
  const s8=pres.addSlide(); s8.background={color:C.darkBlue};
  s8.addShape(pres.shapes.RECTANGLE,{x:0,y:0,w:0.22,h:5.625,fill:{color:C.lightBlue},line:{color:C.lightBlue,width:0}});
  s8.addShape(pres.shapes.RECTANGLE,{x:0,y:5.1,w:10,h:0.525,fill:{color:"0C1E3A"},line:{color:"0C1E3A",width:0}});
  const fi=await iconPng(FaFlag,"#"+C.lightBlue,256);
  s8.addImage({data:fi,x:0.5,y:0.55,w:0.5,h:0.5});
  s8.addText("Recommended Next Steps",{x:0.5,y:1.15,w:9,h:0.55,fontSize:34,bold:true,color:C.white,fontFace:"Calibri"});
  s8.addText("Ordered by impact. Quick wins first, sustainable growth last.",{x:0.5,y:1.76,w:9,h:0.3,fontSize:12,color:"7ABCD4",fontFace:"Calibri",italic:true});
  const seq=narrative.sequence||[];
  const seqColors=[C.lightBlue,C.lightBlue,C.emerald,C.banana];
  const seqIcons=["1","2","3","4"];
  seq.forEach((t,i)=>{
    const col=i%2, row=Math.floor(i/2);
    const x=0.4+col*4.85, y=2.2+row*1.35;
    s8.addShape(pres.shapes.RECTANGLE,{x,y,w:4.65,h:1.18,fill:{color:"0F2040"},shadow:{type:"outer",blur:8,offset:2,angle:135,color:"000000",opacity:0.2},line:{color:"1A3A60",width:0.5}});
    s8.addShape(pres.shapes.RECTANGLE,{x,y,w:0.55,h:1.18,fill:{color:seqColors[i]||C.lightBlue},line:{color:seqColors[i]||C.lightBlue,width:0}});
    s8.addText(seqIcons[i],{x,y,w:0.55,h:1.18,fontSize:22,bold:true,color:C.darkBlue,fontFace:"Calibri",align:"center",valign:"middle"});
    s8.addText(t.action||t.week||"",{x:x+0.65,y:y+0.1,w:3.85,h:0.32,fontSize:12,bold:true,color:C.white,fontFace:"Calibri",margin:0});
    s8.addText(t.why||t.body||"",  {x:x+0.65,y:y+0.46,w:3.85,h:0.58,fontSize:10,color:"A8C4D8",fontFace:"Calibri",margin:0});
  });
  s8.addText(`${D.preparedBy}   ·   ${D.domain}   ·   ${D.date||""}`,{x:0.5,y:5.15,w:9,h:0.25,fontSize:8,color:"3A6080",fontFace:"Calibri"});

  return await pres.write({outputType:"base64"});
}

// ── ROUTES ────────────────────────────────────────────────────

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
          if (!data[k] || data[k] === 0) data[k] = auditData[k];
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

      // BrightLocal citation audit
      if (data.domain && BRIGHTLOCAL_KEY && !data.citationsFound) {
        console.log(`[${jobId}] Running BrightLocal audit...`);
        Object.assign(data, await brightlocalCitationAudit(data.domain, data.clientName, data.location));
      }

      // Claude narrative
      console.log(`[${jobId}] Generating narrative...`);
      const narrative = await getNarrative(data);

      // Build PPTX
      console.log(`[${jobId}] Building PPTX...`);
      const pptxBase64 = await buildPptx(data, narrative);
      const date       = data.date || new Date().toLocaleDateString("en-US",{month:"short",year:"numeric"});
      const fileName   = `SEO Audit — ${data.clientName} — ${date}.pptx`;

      setJobDone(jobId, { pptxBase64, fileName });
      console.log(`[${jobId}] Done.`);

    } catch(err) {
      console.error(`[${jobId}] Error:`, err.message);
      setJobError(jobId, err.message);
    }
  })();
});

// GET /job/:id — poll for job status
app.get("/job/:id", (req, res) => {
  const job = getJob(req.params.id);
  if (!job) return res.status(404).json({ status: "not_found" });
  if (job.status === "done")  return res.json({ status: "done",  ...job.result });
  if (job.status === "error") return res.json({ status: "error", error: job.error });
  res.json({ status: "pending" });
});

app.get("/", (req,res) => res.sendFile(path.join(__dirname,"public","index.html")));

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`SEO Audit API v3 running on port ${PORT}`));
