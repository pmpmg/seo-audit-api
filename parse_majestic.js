// Majestic CSV parser — handles both Summary and Comparison exports

function parseMajesticCsv(buffer) {
  try {
    const text    = buffer.toString("utf8");
    const lines   = text.trim().split("\n").filter(l => l.trim());
    if (lines.length < 2) return {};

    const headers = lines[0].split(",").map(h => h.trim().replace(/"/g,""));
    const rows    = lines.slice(1).map(line => {
      const vals = line.split(",").map(v => v.trim().replace(/"/g,""));
      const row  = {};
      headers.forEach((h,i) => row[h] = vals[i]||"");
      return row;
    }).filter(r => Object.values(r).some(v => v));

    if (!rows.length) return {};

    // ── Detect format ─────────────────────────────────────────
    // Comparison export has a "Domain" or "Item" column
    // Summary export has metrics directly in first row
    const isComparison = headers.some(h =>
      /^(domain|item|url)$/i.test(h.trim())
    );

    if (isComparison) {
      return parseComparison(headers, rows);
    } else {
      return parseSummary(headers, rows);
    }

  } catch(e) {
    console.error("Majestic parse error:", e.message);
    return {};
  }
}

// ── SINGLE DOMAIN SUMMARY ─────────────────────────────────────
function parseSummary(headers, rows) {
  const r = rows[0];
  return {
    trustFlow:        safeFloat(r, ["TrustFlow","Trust Flow","TF"]),
    citationFlow:     safeFloat(r, ["CitationFlow","Citation Flow","CF"]),
    referringDomains: safeInt(  r, ["RefDomains","Referring Domains","RD"]),
    totalBacklinks:   safeInt(  r, ["ExtBackLinks","Total Backlinks","Backlinks","BL"]),
    topicalTrustFlow: safeFloat(r, ["TopicalTrustFlow","Topical Trust Flow"]),
  };
}

// ── COMPARISON EXPORT ─────────────────────────────────────────
// Returns client metrics + competitor array
function parseComparison(headers, rows) {
  const domainCol = headers.find(h => /^(domain|item|url)$/i.test(h)) || headers[0];

  const parsed = rows.map(r => ({
    domain:          r[domainCol] || "",
    trustFlow:        safeFloat(r, ["TrustFlow","Trust Flow","TF"]),
    citationFlow:     safeFloat(r, ["CitationFlow","Citation Flow","CF"]),
    referringDomains: safeInt(  r, ["RefDomains","Referring Domains","RD"]),
    totalBacklinks:   safeInt(  r, ["ExtBackLinks","Total Backlinks","Backlinks","BL"]),
    topicalTrustFlow: safeFloat(r, ["TopicalTrustFlow","Topical Trust Flow"]),
  }));

  if (!parsed.length) return {};

  // First row = client domain
  const client = parsed[0];
  const comps  = parsed.slice(1).map(c => ({
    domain:          c.domain,
    trustFlow:        c.trustFlow,
    citationFlow:     c.citationFlow,
    referringDomains: c.referringDomains,
    totalBacklinks:   c.totalBacklinks,
  }));

  return {
    // Client metrics
    trustFlow:        client.trustFlow,
    citationFlow:     client.citationFlow,
    referringDomains: client.referringDomains,
    totalBacklinks:   client.totalBacklinks,
    topicalTrustFlow: client.topicalTrustFlow,
    // Competitor data merged into existing competitors array
    majesticCompetitors: comps,
  };
}

// ── HELPERS ───────────────────────────────────────────────────
function safeFloat(row, keys) {
  for (const k of keys) {
    if (row[k] !== undefined && row[k] !== "") {
      const v = parseFloat(row[k]);
      if (!isNaN(v)) return v;
    }
  }
  return 0;
}

function safeInt(row, keys) {
  for (const k of keys) {
    if (row[k] !== undefined && row[k] !== "") {
      const v = parseInt(String(row[k]).replace(/,/g,""));
      if (!isNaN(v)) return v;
    }
  }
  return 0;
}

module.exports = { parseMajesticCsv };
