// Screaming Frog Crawl Overview CSV parser
// Extracts key metrics from the exported crawl overview report

function parseScreamingFrogCsv(buffer) {
  try {
    const text = buffer.toString("utf8");
    const lines = text.trim().split("\n");
    if (lines.length < 2) return {};

    // Normalize headers
    const headers = lines[0].split(",").map(h => h.trim().replace(/"/g,"").toLowerCase());
    const result = {};

    // Parse each row looking for key metrics
    const data = [];
    lines.slice(1).forEach(line => {
      const vals = line.split(",").map(v => v.trim().replace(/"/g,""));
      const row = {};
      headers.forEach((h,i) => row[h] = vals[i]||"");
      data.push(row);
    });

    // Count response codes
    const codes = { "200":0, "301":0, "302":0, "404":0, "500":0 };
    let totalPages = 0, missingTitle = 0, missingDesc = 0,
        titleTooLong = 0, descTooLong = 0, missingH1 = 0,
        duplicateTitle = 0, duplicateDesc = 0, missingAlt = 0,
        deepPages = 0, slowPages = 0, redirectChains = 0,
        canonicalIssues = 0, totalSize = 0;

    data.forEach(row => {
      totalPages++;
      const status = row["status code"] || row["status"] || row["statuscode"] || "";
      if (codes[status] !== undefined) codes[status]++;

      // Title issues
      if (!row["title 1"] && !row["title"]) missingTitle++;
      const titleLen = parseInt(row["title 1 length"] || row["title length"] || 0);
      if (titleLen > 60) titleTooLong++;

      // Meta desc issues
      if (!row["meta description 1"] && !row["meta description"]) missingDesc++;
      const descLen = parseInt(row["meta description 1 length"] || row["meta description length"] || 0);
      if (descLen > 160) descTooLong++;

      // H1
      if (!row["h1-1"] && !row["h1"]) missingH1++;

      // Crawl depth
      const depth = parseInt(row["crawl depth"] || row["depth"] || 0);
      if (depth > 4) deepPages++;

      // Response time
      const responseTime = parseFloat(row["response time"] || 0);
      if (responseTime > 3) slowPages++;

      // Redirect chains
      const redirects = parseInt(row["redirect chain"] || row["redirects"] || 0);
      if (redirects > 1) redirectChains++;

      // Canonical
      const canonical = row["canonical link element 1"] || row["canonical"] || "";
      const url = row["address"] || row["url"] || "";
      if (canonical && canonical !== url) canonicalIssues++;
    });

    return {
      sfPagesCrawled:    totalPages,
      sfPages200:        codes["200"],
      sfRedirects301:    codes["301"],
      sfRedirects302:    codes["302"],
      sfErrors404:       codes["404"],
      sfErrors500:       codes["500"],
      sfMissingTitle:    missingTitle,
      sfTitleTooLong:    titleTooLong,
      sfMissingDesc:     missingDesc,
      sfMissingH1:       missingH1,
      sfDeepPages:       deepPages,
      sfSlowPages:       slowPages,
      sfRedirectChains:  redirectChains,
      sfCanonicalIssues: canonicalIssues,
    };
  } catch(e) {
    console.error("Screaming Frog parse error:", e.message);
    return {};
  }
}

module.exports = { parseScreamingFrogCsv };
