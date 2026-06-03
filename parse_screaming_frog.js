// parse_screaming_frog.js
// Parses Screaming Frog Internal HTML export (Internal tab → Export)
// Extracts per-URL data, aggregate stats, and auto-detected practice area pyramids

"use strict";

function parseScreamingFrogCsv(buffer) {
  try {
    const text = buffer.toString("utf-8").replace(/^\uFEFF/, "");

    // Parse CSV properly handling quoted fields
    function parseCSVLine(line) {
      const result = [];
      let current = '';
      let inQuotes = false;
      for (let i = 0; i < line.length; i++) {
        if (line[i] === '"') {
          inQuotes = !inQuotes;
        } else if (line[i] === ',' && !inQuotes) {
          result.push(current.trim());
          current = '';
        } else {
          current += line[i];
        }
      }
      result.push(current.trim());
      return result;
    }

    const lines = text.split(/\r?\n/).filter(l => l.trim());
    if (lines.length < 2) return { sf: null };

    const headers = parseCSVLine(lines[0]).map(h => h.replace(/^"|"$/g, '').trim());
    const rows = lines.slice(1).map(line => {
      const vals = parseCSVLine(line).map(v => v.replace(/^"|"$/g, '').trim());
      const row = {};
      headers.forEach((h, i) => row[h] = vals[i] || '');
      return row;
    }).filter(r => r['Address'] && r['Address'].startsWith('http'));

    // Filter to indexable HTML pages only
    const html = rows.filter(r =>
      (r['Content Type'] || '').includes('text/html') &&
      r['Status Code'] === '200' &&
      r['Indexability'] === 'Indexable'
    );

    // ── AGGREGATE STATS ──────────────────────────────────────────
    const totalUrlsCrawled  = rows.length;
    const htmlPages         = html.length;
    const titlesOver60      = html.filter(r => parseInt(r['Title 1 Length']||0) > 60).length;
    const titlesBelow30     = html.filter(r => { const l=parseInt(r['Title 1 Length']||0); return l>0&&l<30; }).length;
    const titlesMissing     = html.filter(r => parseInt(r['Title 1 Length']||0) === 0).length;
    const metaOver155       = html.filter(r => parseInt(r['Meta Description 1 Length']||0) > 155).length;
    const metaMissing       = html.filter(r => parseInt(r['Meta Description 1 Length']||0) === 0).length;
    const h1Missing         = html.filter(r => !r['H1-1'] || r['H1-1'].trim() === '').length;
    const lowContentPages   = html.filter(r => parseInt(r['Word Count']||0) < 300).length;
    const readabilityHard   = html.filter(r => ['Hard','Very Hard'].includes(r['Readability'])).length;
    const errors4xx         = rows.filter(r => (r['Status Code']||'').startsWith('4')).length;
    const redirects3xx      = rows.filter(r => (r['Status Code']||'').startsWith('3')).length;

    // Depth distribution
    const depthDist = { d0:0, d1:0, d2:0, d3:0, d4plus:0 };
    html.forEach(r => {
      const d = parseInt(r['Crawl Depth']||0);
      if (d===0) depthDist.d0++;
      else if (d===1) depthDist.d1++;
      else if (d===2) depthDist.d2++;
      else if (d===3) depthDist.d3++;
      else depthDist.d4plus++;
    });

    // ── AUTO-DETECT PRACTICE AREA CLUSTERS ───────────────────────
    const STOP = new Set([
      'the','a','an','in','of','for','to','and','or','with','by','at','on','is','are',
      'how','can','what','why','when','your','our','you','we','i','my','do','does','did',
      'lawyer','attorney','law','firm','pc','llc','lp','legal','rights','case','cases',
      'claim','claims','help','need','find','get','free','best','top','local','near',
      'me','vs','it','this','that','its','was','has','had','have','been','being',
      'accident','accidents','injury','injuries','personal','general','common','different','types','type','involving','following','after','result','results','role','victim','victims',
      'page','pages','site','web','blog','post','category','tag','archive','date',
      'after','before','about','from','will','should','would','could','may','might',
      '2020','2021','2022','2023','2024','2025','2026','jan','feb','mar','apr','may',
      'jun','jul','aug','sep','oct','nov','dec','january','february','march','april',
      'june','july','august','september','october','november','december',
    ]);

    // Words that indicate geographic pages (not practice areas)
    const GEO_WORDS = new Set([
      'grand','montrose','delta','mesa','county','junction','colorado','denver','boulder',
      'pueblo','aurora','fort','collins','springs','greeley','thornton','arvada',
      'westminster','lakewood','highlands','ranch','lone','tree','castle','rock',
    ]);

    // Words that indicate firm/utility pages
    const UTILITY = new Set([
      'contact','about','team','staff','results','testimonials','sitemap','disclaimer',
      'privacy','terms','killian','davis','richter','cares','luke','smith','gabriel',
      'maldonado','christopher','damon','keith','jerome',
    ]);

    function slugWords(url) {
      const path = url.replace(/https?:\/\/[^/]+/, '').replace(/\.(html|php|aspx)$/, '');
      return path.split(/[-/]/)
        .map(w => w.toLowerCase().trim())
        .filter(w => w.length > 2 && !STOP.has(w) && !GEO_WORDS.has(w) && !UTILITY.has(w) && !/^\d+$/.test(w));
    }

    // Synonyms/merges — combine related terms into one cluster name
    const MERGE = {
      'compensation': 'workers',
      'comp':         'workers',
      'wrongful':     'wrongful-death',
      'death':        'wrongful-death',
      'motorist':     'uninsured',
      'underinsured': 'uninsured',
      'premises':     'slip-fall',
      'slip':         'slip-fall',
      'fall':         'slip-fall',
      'product':      'product-liability',
      'liability':    'product-liability',
      'brain':        'brain-injury',
      'spinal':       'spinal-injury',
      'spine':        'spinal-injury',
      'construction': 'construction-accident',
      'oilfield':     'oilfield-accident',
      'oil':          'oilfield-accident',
      'dog':          'dog-bite',
      'bite':         'dog-bite',
      'truck':        'truck-accident',
      'car':          'car-accident',
      'motorcycle':   'motorcycle-accident',
      'bicycle':      'bicycle-accident',
      'bike':         'bicycle-accident',
      'pedestrian':   'pedestrian-accident',
      'bus':          'bus-accident',
      'rollover':     'rollover-accident',
      'distracted':   'distracted-driving',
      'fatigued':     'truck-accident',
      'jackknife':    'truck-accident',
      'workers':      'workers-comp',
      'compensation': 'workers-comp',
      'comp':         'workers-comp',
      'radiation':    'radiation-exposure',
      'nursing':      'nursing-home',
    };

    // Build word → URLs map
    const wordToUrls = {};
    html.forEach(r => {
      const words = slugWords(r['Address']);
      words.forEach(w => {
        const key = MERGE[w] || w;
        if (!wordToUrls[key]) wordToUrls[key] = new Set();
        wordToUrls[key].add(r['Address']);
      });
    });

    // Only keep clusters with 2+ pages, not geo/utility noise
    const clusterCandidates = Object.entries(wordToUrls)
      .filter(([w, urls]) => urls.size >= 2)
      .sort((a, b) => b[1].size - a[1].size);

    // Build practice area clusters
    const practiceAreas = [];
    const assignedUrls = new Set();

    for (const [clusterKey, urlSet] of clusterCandidates) {
      const clusterUrls = [...urlSet];

      // Get full row data for each URL
      const pages = clusterUrls.map(url => {
        const r = html.find(h => h['Address'] === url);
        if (!r) return null;
        return {
          url,
          depth:     parseInt(r['Crawl Depth'] || 99),
          inlinks:   parseInt(r['Inlinks'] || 0),
          wordCount: parseInt(r['Word Count'] || 0),
          h1:        r['H1-1'] || '',
          title:     r['Title 1'] || '',
        };
      }).filter(Boolean);

      if (!pages.length) continue;

      // Identify hub: depth 1 page with highest inlinks
      const depth1 = pages.filter(p => p.depth === 1).sort((a,b) => b.inlinks - a.inlinks);
      // Identify sub-hubs: depth 2, 5+ inlinks
      const subHubs = pages.filter(p => p.depth === 2 && p.inlinks >= 5);
      // Spokes: everything else (depth 2 low inlinks, depth 3)
      const spokes  = pages.filter(p => !depth1.includes(p) && !subHubs.includes(p));

      // Score pyramid health
      let status, statusColor;
      if (depth1.length > 0 && subHubs.length >= 3 && spokes.length >= 5) {
        status = "Strong";      statusColor = "00684F";
      } else if (depth1.length > 0 && (subHubs.length >= 1 || spokes.length >= 3)) {
        status = "Needs Clusters"; statusColor = "F5A623";
      } else if (depth1.length > 0) {
        status = "Thin";        statusColor = "C0392B";
      } else {
        status = "No Hub";      statusColor = "C0392B";
      }

      // Format cluster name nicely
      const label = clusterKey
        .split('-')
        .map(w => w.charAt(0).toUpperCase() + w.slice(1))
        .join(' ');

      practiceAreas.push({
        label,
        hubCount:    depth1.length,
        subHubCount: subHubs.length,
        spokeCount:  spokes.length,
        totalPages:  pages.length,
        status,
        statusColor,
        hubPage:     depth1[0]?.url || null,
      });

      clusterUrls.forEach(u => assignedUrls.add(u));
    }

    // Sort by total pages descending, cap at 8
    const topPracticeAreas = practiceAreas
      .sort((a, b) => b.totalPages - a.totalPages)
      .slice(0, 8);

    // Top inlinked pages (hub candidates)
    const topInlinkedPages = [...html]
      .sort((a, b) => parseInt(b['Inlinks']||0) - parseInt(a['Inlinks']||0))
      .slice(0, 10)
      .map(r => r['Address']);

    return {
      // Legacy fields for backward compat
      sfPagesCrawled:  totalUrlsCrawled,
      sfMissingDesc:   metaMissing,
      sfTitleTooLong:  titlesOver60,
      sfMissingH1:     h1Missing,

      // Rich data object
      sf: {
        totalUrlsCrawled,
        htmlPages,
        titlesOver60,
        titlesBelow30,
        titlesMissing,
        metaOver155,
        metaMissing,
        h1Missing,
        lowContentPages,
        readabilityHard,
        errors4xx,
        redirects3xx,
        depthDist,
        topInlinkedPages,
        practiceAreas: topPracticeAreas,
      }
    };

  } catch(e) {
    console.error("parseScreamingFrogCsv error:", e.message);
    return { sf: null };
  }
}

module.exports = { parseScreamingFrogCsv };
