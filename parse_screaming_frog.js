// parse_screaming_frog.js
// Parses a Screaming Frog Crawl Overview CSV export into structured data
// Works with the "Crawl Overview" export format (File → Export → Crawl Overview)

"use strict";

function parseScreamingFrogCsv(buffer) {
  try {
    const text = buffer.toString("utf-8").replace(/^\uFEFF/, ""); // strip BOM
    const rows = text.split(/\r?\n/).map(line =>
      line.split(",").map(c => c.trim().replace(/^"|"$/g, ""))
    ).filter(r => r.some(c => c));

    // Build section-aware lookup map
    const sections = {};
    let currentSection = "General";
    for (const r of rows) {
      // Section header = single non-empty cell that isn't purely numeric
      if (r.filter(c => c).length === 1 && r[0] && !/^\d/.test(r[0]) && r[0] !== "Summary") {
        currentSection = r[0];
        if (!sections[currentSection]) sections[currentSection] = {};
        continue;
      }
      if (r.length >= 2 && r[0] && r[1]) {
        const num = parseInt(r[1].replace(/,/g, ""));
        if (!isNaN(num)) {
          if (!sections[currentSection]) sections[currentSection] = {};
          sections[currentSection][r[0]] = num;
        }
      }
    }

    // Helper: get value from a section
    const s = (section, key, def = 0) =>
      (sections[section] || {})[key] ?? def;

    // ── Extract all useful metrics ────────────────────────────────

    // Site overview
    const totalUrlsCrawled   = s("General", "Total URLs Crawled") || s("Internal", "All");
    const htmlPages          = s("Internal", "HTML");
    const indexablePages     = s("General", "Total Internal Indexable URLs");
    const nonIndexablePages  = s("General", "Total Internal Non-Indexable URLs");

    // Response codes
    const errors4xx          = s("Response Codes", "Internal Client Error (4xx)");
    const errors5xx          = s("Response Codes", "Internal Server Error (5xx)");
    const redirects3xx       = s("Response Codes", "Internal Redirection (3xx)");

    // Page titles
    const titlesMissing      = s("Page Titles", "Missing");
    const titlesOver60       = s("Page Titles", "Over 60 Characters");
    const titlesBelow30      = s("Page Titles", "Below 30 Characters");
    const titlesSameAsH1     = s("Page Titles", "Same as H1");
    const titlesDuplicate    = s("Page Titles", "Duplicate");

    // Meta descriptions
    const metaMissing        = s("Meta Description", "Missing");
    const metaDuplicate      = s("Meta Description", "Duplicate");
    const metaOver155        = s("Meta Description", "Over 155 Characters");
    const metaBelow70        = s("Meta Description", "Below 70 Characters");

    // H1/H2
    const h1Missing          = s("H1", "Missing");
    const h1Duplicate        = s("H1", "Duplicate");
    const h1Over70           = s("H1", "Over 70 Characters");
    const h2Missing          = s("H2", "Missing");
    const h2Duplicate        = s("H2", "Duplicate");

    // Content quality
    const lowContentPages    = s("Content", "Low Content Pages");
    const readabilityHard    = s("Content", "Readability Difficult");
    const readabilityVHard   = s("Content", "Readability Very Difficult");
    const nearDuplicates     = s("Content", "Near Duplicates");
    const exactDuplicates    = s("Content", "Exact Duplicates");

    // Images
    const missingAltText     = s("Images", "Missing Alt Text");
    const imagesOver100kb    = s("Images", "Over 100 KB");

    // Canonicals
    const missingCanonical   = s("Canonicals", "Missing");

    // Links
    const noAnchorText       = s("Links", "Internal Outlinks With No Anchor Text");
    const highCrawlDepth     = s("Links", "Pages With High Crawl Depth");

    // Security
    const httpUrls           = s("Security", "HTTP URLs");
    const mixedContent       = s("Security", "Mixed Content");

    // Structured data
    const structuredDataMissing = s("Structured Data", "Missing");
    const structuredDataErrors  = s("Structured Data", "Validation Errors");

    // Depth distribution
    const depth = sections["Depth (Clicks from Start URL)"] || {};
    const depthDist = {
      d0: depth["0"] || 0,
      d1: depth["1"] || 0,
      d2: depth["2"] || 0,
      d3: depth["3"] || 0,
      d4plus: (depth["4"] || 0) + (depth["5"] || 0) +
               (depth["6"] || 0) + (depth["7"] || 0) +
               (depth["8"] || 0) + (depth["9"] || 0) +
               (depth["10+"] || 0),
    };

    // Top inlinked pages (hub detection)
    const inlinkPages = [];
    let inInlinks = false;
    for (const r of rows) {
      if (r[0] && r[0].includes("Inlinks (Top")) { inInlinks = true; continue; }
      if (inInlinks && r[0] && r[0].startsWith("http")) {
        inlinkPages.push(r[0]);
      } else if (inInlinks && !r[0]) {
        inInlinks = false;
      }
    }

    // ── Content architecture signals ─────────────────────────────
    // Infer hub/sub-hub/spoke from depth distribution
    // depth 1 = likely hub pages, depth 2-3 = sub-hub/spoke, 4+ = deep spoke/blog
    const hubPages      = depthDist.d1;
    const subHubPages   = depthDist.d2;
    const spokePages    = depthDist.d3;
    const deepPages     = depthDist.d4plus;
    const blogPages     = inlinkPages.filter(u => u.includes("/blog")).length > 0
      ? depthDist.d3 + depthDist.d4plus  // rough estimate if blog exists
      : 0;

    // Content health score (0-100) for use in slides
    const totalHtml = htmlPages || 1;
    const issueRatio = (
      titlesMissing + metaMissing + h1Missing +
      lowContentPages + errors4xx + missingAltText
    ) / totalHtml;
    const contentHealthScore = Math.max(0, Math.min(100,
      Math.round(100 - (issueRatio * 100) - (metaDuplicate / totalHtml * 30) - (metaOver155 / totalHtml * 20))
    ));

    // ── Return flat data object ───────────────────────────────────
    return {
      // Existing fields the rest of the app already uses
      sfPagesCrawled:   totalUrlsCrawled,
      sfMissingDesc:    metaMissing,
      sfTitleTooLong:   titlesOver60,
      sfMissingH1:      h1Missing,

      // New rich crawl data
      sf: {
        // Overview
        totalUrlsCrawled,
        htmlPages,
        indexablePages,
        nonIndexablePages,

        // Response codes
        errors4xx,
        errors5xx,
        redirects3xx,

        // Titles
        titlesMissing,
        titlesOver60,
        titlesBelow30,
        titlesSameAsH1,
        titlesDuplicate,

        // Meta
        metaMissing,
        metaDuplicate,
        metaOver155,
        metaBelow70,

        // Headings
        h1Missing,
        h1Duplicate,
        h1Over70,
        h2Missing,
        h2Duplicate,

        // Content
        lowContentPages,
        readabilityHard,
        readabilityVHard,
        nearDuplicates,
        exactDuplicates,

        // Images
        missingAltText,
        imagesOver100kb,

        // Technical
        missingCanonical,
        noAnchorText,
        highCrawlDepth,
        httpUrls,
        mixedContent,
        structuredDataMissing,
        structuredDataErrors,

        // Architecture
        depthDist,
        hubPages,
        subHubPages,
        spokePages,
        deepPages,
        topInlinkedPages: inlinkPages.slice(0, 10),

        // Scores
        contentHealthScore,
      }
    };
  } catch (e) {
    console.error("parseScreamingFrogCsv error:", e.message);
    return { sf: null };
  }
}

module.exports = { parseScreamingFrogCsv };
