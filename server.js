const express = require("express");
const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");

const app  = express();
app.use(express.json({ limit: "10mb" }));

// ── COLORS ──────────────────────────────────────────────────
const C = {
  navy:      "1B2A4A",
  gold:      "C8972A",
  red:       "C0442A",
  steel:     "4A6FA5",
  green:     "2E7D52",
  lightGray: "F2F4F6",
  midGray:   "8A9BB0",
  white:     "FFFFFF",
  dark:      "111827",
};

// ── ICON HELPER ──────────────────────────────────────────────
const { FaCheckCircle, FaServer, FaLink, FaHeading, FaRobot,
        FaCode, FaTags, FaLayerGroup, FaChartBar, FaArrowRight,
        FaBolt, FaMousePointer, FaClipboardList, FaFlag } = require("react-icons/fa");

function renderIconSvg(IconComponent, color="#FFFFFF", size=256) {
  return ReactDOMServer.renderToStaticMarkup(
    React.createElement(IconComponent, { color, size: String(size) })
  );
}
async function iconPng(IconComponent, color, size=256) {
  const svg = renderIconSvg(IconComponent, color, size);
  const buf = await sharp(Buffer.from(svg)).png().toBuffer();
  return "image/png;base64," + buf.toString("base64");
}
async function donutPng(score, color, size=300) {
  const r=100,cx=150,cy=150,strokeW=22;
  const circ=2*Math.PI*r, dash=(score/100)*circ;
  const svg=`<svg xmlns="http://www.w3.org/2000/svg" width="${size}" height="${size}" viewBox="0 0 300 300">
    <circle cx="${cx}" cy="${cy}" r="${r}" fill="none" stroke="#E5E7EB" stroke-width="${strokeW}"/>
    <circle cx="${cx}" cy="${cy}" r="${r}" fill="none" stroke="#${color}" stroke-width="${strokeW}"
      stroke-dasharray="${dash} ${circ}" stroke-dashoffset="${circ/4}" stroke-linecap="round"/>
  </svg>`;
  const buf=await sharp(Buffer.from(svg)).resize(size,size).png().toBuffer();
  return "image/png;base64,"+buf.toString("base64");
}
async function barChartPng(data) {
  const items=[
    {label:"Missing ALT text",      value:data.missingAlt||0,    color:C.red},
    {label:"Schema errors",         value:data.schemaErrors||0,  color:C.red},
    {label:"Low text-to-HTML",      value:data.thinPages||0,     color:C.gold},
    {label:"Titles too long",       value:data.titlesTooLong||0, color:C.gold},
    {label:"Missing meta desc",     value:data.missingDesc||0,   color:C.gold},
    {label:"Broken external links", value:data.brokenExternal||0,color:C.steel},
    {label:"Missing H1",            value:data.missingH1||0,     color:C.steel},
  ];
  const max=Math.max(...items.map(i=>i.value),1);
  const W=800,barH=36,gap=14,labelW=200,numW=60,margin=20;
  const H=items.length*(barH+gap)+margin*2;
  const barMaxW=W-labelW-numW-margin*2;
  let bars=items.map((item,i)=>{
    const y=margin+i*(barH+gap);
    const bw=Math.max(4,(item.value/max)*barMaxW);
    return `<text x="${labelW-8}" y="${y+barH/2+5}" text-anchor="end" font-family="Arial" font-size="13" fill="#${C.dark}">${item.label}</text>
      <rect x="${labelW}" y="${y}" width="${bw}" height="${barH}" fill="#${item.color}" rx="3"/>
      <text x="${labelW+bw+8}" y="${y+barH/2+5}" font-family="Arial" font-size="13" font-weight="bold" fill="#${C.dark}">${item.value}</text>`;
  }).join("");
  const svg=`<svg xmlns="http://www.w3.org/2000/svg" width="${W}" height="${H}">
    <rect width="${W}" height="${H}" fill="#${C.lightGray}" rx="8"/>${bars}</svg>`;
  const buf=await sharp(Buffer.from(svg)).resize(W,H).png().toBuffer();
  return "image/png;base64,"+buf.toString("base64");
}

const makeShadow=()=>({type:"outer",blur:8,offset:2,angle:135,color:"000000",opacity:0.08});

function footerLine(slide, preparedBy, domain) {
  slide.addText(preparedBy||"", {x:0.4,y:5.35,w:3,h:0.2,fontSize:9,color:C.midGray,fontFace:"Calibri"});
  slide.addText(domain||"",     {x:6.6,y:5.35,w:3,h:0.2,fontSize:9,color:C.midGray,fontFace:"Calibri",align:"right"});
  slide.addShape("line",        {x:0.4,y:5.32,w:9.2,h:0,line:{color:"E5E7EB",width:0.5}});
}
function sectionLabel(slide,text){
  slide.addText(text,{x:0.5,y:0.28,w:9,h:0.22,fontSize:9,color:C.gold,bold:true,charSpacing:4,fontFace:"Calibri"});
}
function slideTitle(slide,text){
  slide.addText(text,{x:0.5,y:0.55,w:9,h:0.75,fontSize:34,bold:true,color:C.navy,fontFace:"Georgia"});
}
function statCard(pres,slide,x,y,w,h,number,label,sub,accentColor){
  slide.addShape(pres.shapes.RECTANGLE,{x,y,w,h,fill:{color:C.white},shadow:makeShadow(),line:{color:"E5E7EB",width:0.5}});
  slide.addShape(pres.shapes.RECTANGLE,{x,y,w,h:0.07,fill:{color:accentColor},line:{color:accentColor,width:0}});
  slide.addText(number,{x:x+0.18,y:y+0.18,w:w-0.3,h:0.78,fontSize:40,bold:true,color:accentColor,fontFace:"Georgia",margin:0});
  slide.addText(label, {x:x+0.18,y:y+1.0, w:w-0.3,h:0.28,fontSize:11,bold:true,color:C.navy,fontFace:"Calibri",margin:0});
  if(sub) slide.addText(sub,{x:x+0.18,y:y+1.3,w:w-0.3,h:0.55,fontSize:10,color:C.midGray,fontFace:"Calibri",margin:0});
}

// ── GENERATE PPTX ────────────────────────────────────────────
async function buildPptx(data, narrative) {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";

  const pb  = data.preparedBy || "PMP Marketing Group";
  const dom = data.domain     || "";

  // SLIDE 1: COVER
  const s1 = pres.addSlide();
  s1.background = {color:C.navy};
  s1.addShape(pres.shapes.RECTANGLE,{x:0,y:0,w:0.18,h:5.625,fill:{color:C.gold},line:{color:C.gold,width:0}});
  s1.addText("TECHNICAL SEO AUDIT",{x:0.45,y:1.1,w:9.1,h:0.3,fontSize:10,color:C.gold,bold:true,charSpacing:5,fontFace:"Calibri"});
  s1.addText(narrative.executiveSummary||"What is working, what is costing you, and what to fix first.",
    {x:0.45,y:1.55,w:8.5,h:0.45,fontSize:18,color:"AABBD0",fontFace:"Calibri",italic:true});
  s1.addText(dom,{x:0.45,y:2.2,w:9,h:0.7,fontSize:42,bold:true,color:C.white,fontFace:"Georgia"});
  s1.addText(data.clientName||"",{x:0.45,y:2.95,w:9,h:0.4,fontSize:20,color:C.gold,fontFace:"Calibri"});
  s1.addShape(pres.shapes.RECTANGLE,{x:0.45,y:3.52,w:9.1,h:0.015,fill:{color:"3A5070"},line:{color:"3A5070",width:0}});
  s1.addText([
    {text:`Prepared by ${pb}`,options:{color:"7A9BB5"}},{text:"   ·   ",options:{color:"4A6A85"}},
    {text:data.date||"",options:{color:"7A9BB5"}},{text:"   ·   ",options:{color:"4A6A85"}},
    {text:`Source: ${data.source||"SEMrush Site Audit"}`,options:{color:"7A9BB5"}},{text:"   ·   ",options:{color:"4A6A85"}},
    {text:`${data.pagesCrawled||0} pages crawled`,options:{color:"7A9BB5"}},
  ],{x:0.45,y:3.68,w:9.1,h:0.28,fontSize:10,fontFace:"Calibri"});

  // SLIDE 2: BOTTOM LINE
  const s2=pres.addSlide();
  s2.background={color:C.white};
  sectionLabel(s2,"THE BOTTOM LINE");
  slideTitle(s2,"The foundation is solid. Three patterns hold it back.");
  s2.addText(narrative.executiveSummary||"",
    {x:0.5,y:1.42,w:9,h:0.55,fontSize:12,color:C.dark,fontFace:"Calibri"});
  statCard(pres,s2,0.4, 2.1,2.9,2.15,`${data.siteHealth||0}/100`,"SITE HEALTH","Solid base, clear ceiling to climb.",C.navy);
  statCard(pres,s2,3.55,2.1,2.9,2.15,`${data.schemaErrors||0}`,"SCHEMA ERRORS","One broken template, repeated sitewide.",C.red);
  statCard(pres,s2,6.7, 2.1,2.9,2.15,`${(data.missingDesc||0)+(data.titlesTooLong||0)}`,"METADATA GAPS",
    `${data.missingDesc||0} missing descriptions, ${data.titlesTooLong||0} titles cut off.`,C.gold);
  footerLine(s2,pb,dom);

  // SLIDE 3: SCORECARD
  const s3=pres.addSlide();
  s3.background={color:C.white};
  sectionLabel(s3,"SITE HEALTH SCORECARD");
  slideTitle(s3,`Two scores, ${data.pagesCrawled||0} pages of evidence.`);
  const d1=await donutPng(data.siteHealth||0,C.navy);
  const d2=await donutPng(data.aiReadiness||0,C.gold);
  s3.addImage({data:d1,x:0.4,y:1.35,w:2.2,h:2.2});
  s3.addText(`${data.siteHealth||0}`,{x:0.4,y:1.9,w:2.2,h:1.1,fontSize:44,bold:true,color:C.navy,fontFace:"Georgia",align:"center"});
  s3.addText("/100",{x:0.4,y:2.85,w:2.2,h:0.3,fontSize:14,color:C.midGray,fontFace:"Calibri",align:"center"});
  s3.addText("Site Health",{x:0.4,y:3.6,w:2.2,h:0.3,fontSize:12,bold:true,color:C.navy,fontFace:"Calibri",align:"center"});
  s3.addImage({data:d2,x:3.1,y:1.35,w:2.2,h:2.2});
  s3.addText(`${data.aiReadiness||0}`,{x:3.1,y:1.9,w:2.2,h:1.1,fontSize:44,bold:true,color:C.gold,fontFace:"Georgia",align:"center"});
  s3.addText("/100",{x:3.1,y:2.85,w:2.2,h:0.3,fontSize:14,color:C.midGray,fontFace:"Calibri",align:"center"});
  s3.addText("AI Search Readiness",{x:3.1,y:3.6,w:2.2,h:0.3,fontSize:12,bold:true,color:C.navy,fontFace:"Calibri",align:"center"});
  s3.addText("AI Search has the most room to move. Clean schema and descriptive links are exactly what it measures.",
    {x:0.4,y:4.0,w:5.0,h:0.6,fontSize:10,color:C.midGray,fontFace:"Calibri",italic:true,align:"center"});
  const gridX=5.7,gridY=1.3,cardW=1.95,cardH=1.0;
  const urlCards=[
    {val:data.pagesCrawled||0,label:"Pages crawled",   color:C.navy},
    {val:data.pages200||0,    label:"Returning 200 OK",color:C.green},
    {val:data.redirects||0,   label:"Redirects (3xx)", color:C.gold},
    {val:data.errors||0,      label:"Error pages (4xx)",color:C.red},
  ];
  s3.addText("Status of every crawled URL",{x:gridX,y:1.1,w:4.1,h:0.25,fontSize:10,bold:true,color:C.gold,fontFace:"Calibri",charSpacing:2});
  urlCards.forEach((c,i)=>{
    const x=gridX+(i%2)*(cardW+0.12), y=gridY+Math.floor(i/2)*(cardH+0.12);
    s3.addShape(pres.shapes.RECTANGLE,{x,y,w:cardW,h:cardH,fill:{color:C.white},shadow:makeShadow(),line:{color:"E5E7EB",width:0.5}});
    s3.addShape(pres.shapes.RECTANGLE,{x,y,w:0.07,h:cardH,fill:{color:c.color},line:{color:c.color,width:0}});
    s3.addText(`${c.val}`,{x:x+0.18,y:y+0.1,w:cardW-0.25,h:0.52,fontSize:30,bold:true,color:c.color,fontFace:"Georgia",margin:0});
    s3.addText(c.label,   {x:x+0.18,y:y+0.62,w:cardW-0.25,h:0.3,fontSize:10,color:C.midGray,fontFace:"Calibri",margin:0});
  });
  footerLine(s3,pb,dom);

  // SLIDE 4: WHAT IS WORKING
  const s4=pres.addSlide();
  s4.background={color:C.white};
  sectionLabel(s4,"WHAT IS ALREADY WORKING");
  slideTitle(s4,"The expensive stuff is right.");
  s4.addText("These are the issues that cost the most to fix after the fact. None of them are here.",
    {x:0.5,y:1.42,w:9,h:0.35,fontSize:12,color:C.dark,fontFace:"Calibri"});
  const wins=narrative.whatIsWorking||[
    {title:"Zero server errors",       sub:`No 5xx failures across ${data.pagesCrawled||0} pages.`},
    {title:"No broken internal links", sub:"Internal link graph is intact."},
    {title:"Every page has a title",   sub:"No missing or empty title tags."},
    {title:"Nothing blocks the crawl", sub:"Core pages are fully accessible."},
  ];
  const winIcons=[FaServer,FaLink,FaHeading,FaRobot];
  for(let i=0;i<4;i++){
    const x=0.4+i*2.35;
    const iconData=await iconPng(winIcons[i],"#"+C.green,256);
    s4.addShape(pres.shapes.RECTANGLE,{x,y:1.9,w:2.2,h:2.7,fill:{color:C.lightGray},shadow:makeShadow(),line:{color:"E5E7EB",width:0.3}});
    s4.addShape(pres.shapes.OVAL,{x:x+0.75,y:2.05,w:0.7,h:0.7,fill:{color:"D1FAE5"},line:{color:"D1FAE5",width:0}});
    s4.addImage({data:iconData,x:x+0.83,y:2.13,w:0.54,h:0.54});
    s4.addText(wins[i]?.title||"",{x:x+0.14,y:2.85,w:1.92,h:0.55,fontSize:12,bold:true,color:C.navy,fontFace:"Calibri",align:"center"});
    s4.addText(wins[i]?.sub||"",  {x:x+0.14,y:3.45,w:1.92,h:0.7, fontSize:10,color:C.midGray,fontFace:"Calibri",align:"center"});
  }
  footerLine(s4,pb,dom);

  // SLIDE 5: PATTERNS OVERVIEW
  const s5=pres.addSlide();
  s5.background={color:C.white};
  sectionLabel(s5,"WHAT IS HOLDING IT BACK");
  slideTitle(s5,"Three patterns, not a hundred problems.");
  const patterns=narrative.patterns||[
    {num:"01",title:"Broken structured data",body:`${data.schemaErrors||0} schema errors from one broken template.`},
    {num:"02",title:"Metadata gaps",          body:`${data.missingDesc||0} pages with no description and ${data.titlesTooLong||0} titles too long.`},
    {num:"03",title:"Thin signals",           body:`${data.thinPages||0} light pages and ${(data.noAnchors||0)+(data.weakAnchorLinks||0)} links with weak anchor text.`},
  ];
  const patColors=[C.red,C.gold,C.steel];
  const patIcons=[FaCode,FaTags,FaLayerGroup];
  for(let i=0;i<3;i++){
    const x=0.4+i*3.1, p=patterns[i];
    const iconData=await iconPng(patIcons[i],"#FFFFFF",256);
    s5.addShape(pres.shapes.RECTANGLE,{x,y:1.85,w:2.9,h:2.9,fill:{color:C.white},shadow:makeShadow(),line:{color:"E5E7EB",width:0.5}});
    s5.addShape(pres.shapes.RECTANGLE,{x,y:1.85,w:2.9,h:0.72,fill:{color:patColors[i]},line:{color:patColors[i],width:0}});
    s5.addImage({data:iconData,x:x+0.18,y:1.98,w:0.38,h:0.38});
    s5.addText(p.num,  {x:x+0.1,y:1.87,w:2.7,h:0.68,fontSize:28,bold:true,color:C.white,fontFace:"Georgia",align:"right",margin:0});
    s5.addText(p.title,{x:x+0.18,y:2.65,w:2.55,h:0.5,fontSize:13,bold:true,color:C.navy,fontFace:"Calibri"});
    s5.addText(p.body, {x:x+0.18,y:3.2, w:2.55,h:1.2,fontSize:10.5,color:C.dark,fontFace:"Calibri"});
  }
  footerLine(s5,pb,dom);

  // SLIDE 6: PATTERN 1
  const s6=pres.addSlide();
  s6.background={color:C.white};
  sectionLabel(s6,"PATTERN 1 · ERRORS");
  slideTitle(s6,"Structured data: one fix, sitewide.");
  s6.addShape(pres.shapes.RECTANGLE,{x:0.4,y:1.35,w:3.5,h:3.5,fill:{color:C.navy},line:{color:C.navy,width:0}});
  const codeIcon=await iconPng(FaCode,"#"+C.gold,256);
  s6.addImage({data:codeIcon,x:0.75,y:1.6,w:0.65,h:0.65});
  s6.addText(`${data.schemaErrors||0}`,{x:0.4,y:2.2,w:3.5,h:1.4,fontSize:80,bold:true,color:C.gold,fontFace:"Georgia",align:"center"});
  s6.addText("schema instances failing the same check",{x:0.55,y:3.65,w:3.2,h:0.5,fontSize:10,color:"7A9BB5",fontFace:"Calibri",align:"center"});
  s6.addText("The root cause",{x:4.2,y:1.4,w:5.4,h:0.28,fontSize:10,bold:true,color:C.gold,charSpacing:2,fontFace:"Calibri"});
  s6.addText(patterns[0]?.body||"",{x:4.2,y:1.73,w:5.4,h:0.7,fontSize:11,color:C.dark,fontFace:"Calibri"});
  s6.addText("Why it matters",{x:4.2,y:2.55,w:5.4,h:0.28,fontSize:10,bold:true,color:C.gold,charSpacing:2,fontFace:"Calibri"});
  s6.addText("Schema is how Google and AI answer engines read who you are and where you practice. Invalid schema forfeits rich results.",
    {x:4.2,y:2.88,w:5.4,h:0.75,fontSize:11,color:C.dark,fontFace:"Calibri"});
  s6.addShape(pres.shapes.RECTANGLE,{x:4.2,y:3.85,w:1.9,h:0.52,fill:{color:C.white},line:{color:C.navy,width:1.5}});
  s6.addText("Fix 1 template",{x:4.2,y:3.85,w:1.9,h:0.52,fontSize:11,bold:true,color:C.navy,fontFace:"Calibri",align:"center"});
  const arrIcon=await iconPng(FaArrowRight,"#"+C.navy,128);
  s6.addImage({data:arrIcon,x:6.2,y:4.0,w:0.3,h:0.22});
  s6.addShape(pres.shapes.RECTANGLE,{x:6.65,y:3.85,w:1.9,h:0.52,fill:{color:C.green},line:{color:C.green,width:0}});
  s6.addText(`Clear all ${data.schemaErrors||0}`,{x:6.65,y:3.85,w:1.9,h:0.52,fontSize:11,bold:true,color:C.white,fontFace:"Calibri",align:"center"});
  footerLine(s6,pb,dom);

  // SLIDE 7: PATTERN 2
  const s7=pres.addSlide();
  s7.background={color:C.white};
  sectionLabel(s7,"PATTERN 2 · WARNINGS");
  slideTitle(s7,"Metadata: stop letting Google guess.");
  const tagIcon=await iconPng(FaTags,"#"+C.white,256);
  const docIcon=await iconPng(FaClipboardList,"#"+C.white,256);
  s7.addShape(pres.shapes.RECTANGLE,{x:0.4,y:1.5,w:4.45,h:2.1,fill:{color:C.lightGray},shadow:makeShadow(),line:{color:"E5E7EB",width:0.5}});
  s7.addShape(pres.shapes.OVAL,{x:0.65,y:1.65,w:0.7,h:0.7,fill:{color:C.gold},line:{color:C.gold,width:0}});
  s7.addImage({data:tagIcon,x:0.75,y:1.73,w:0.5,h:0.5});
  s7.addText(`${data.missingDesc||0}`,{x:1.5,y:1.6,w:2.8,h:0.85,fontSize:48,bold:true,color:C.navy,fontFace:"Georgia"});
  s7.addText("pages with no meta description",{x:0.6,y:2.5,w:4.1,h:0.3,fontSize:11,bold:true,color:C.navy,fontFace:"Calibri"});
  s7.addText("Google writes the search snippet for you, so the firm loses control of the listing.",
    {x:0.6,y:2.85,w:4.1,h:0.55,fontSize:10,color:C.midGray,fontFace:"Calibri"});
  s7.addShape(pres.shapes.RECTANGLE,{x:5.15,y:1.5,w:4.45,h:2.1,fill:{color:C.lightGray},shadow:makeShadow(),line:{color:"E5E7EB",width:0.5}});
  s7.addShape(pres.shapes.OVAL,{x:5.4,y:1.65,w:0.7,h:0.7,fill:{color:C.gold},line:{color:C.gold,width:0}});
  s7.addImage({data:docIcon,x:5.5,y:1.73,w:0.5,h:0.5});
  s7.addText(`${data.titlesTooLong||0}`,{x:6.25,y:1.6,w:2.8,h:0.85,fontSize:48,bold:true,color:C.navy,fontFace:"Georgia"});
  s7.addText("titles too long",{x:5.35,y:2.5,w:4.1,h:0.3,fontSize:11,bold:true,color:C.navy,fontFace:"Calibri"});
  s7.addText("Titles get cut off mid-message in search results, weakening the click.",
    {x:5.35,y:2.85,w:4.1,h:0.55,fontSize:10,color:C.midGray,fontFace:"Calibri"});
  s7.addShape(pres.shapes.RECTANGLE,{x:0.4,y:3.82,w:9.2,h:0.9,fill:{color:C.navy},line:{color:C.navy,width:0}});
  const mouseIcon=await iconPng(FaMousePointer,"#"+C.gold,128);
  s7.addImage({data:mouseIcon,x:0.55,y:4.08,w:0.32,h:0.32});
  s7.addText([
    {text:"What it costs: ",options:{bold:true,color:C.gold}},
    {text:"click-through on pages that already rank. This is the lever that moves traffic without writing a single new ranking page.",options:{color:C.white}},
  ],{x:1.0,y:3.87,w:8.4,h:0.8,fontSize:10.5,fontFace:"Calibri"});
  footerLine(s7,pb,dom);

  // SLIDE 8: PATTERN 3
  const s8=pres.addSlide();
  s8.background={color:C.white};
  sectionLabel(s8,"PATTERN 3 · CONTENT AND LINKING");
  slideTitle(s8,"Thin pages and silent links.");
  const thinCards=[
    {val:data.thinPages||0,      label:"pages low on visible text", sub:"Light content relative to page code. Reads as thin to search engines."},
    {val:data.noAnchors||0,      label:"links with no anchor text", sub:"Links pass no context. Mostly repeating nav, button, and icon patterns."},
    {val:data.weakAnchorLinks||0,label:"links with weak anchors",   sub:'Anchors like "click here" that tell engines nothing about the target.'},
  ];
  thinCards.forEach((c,i)=>{
    const x=0.4+i*3.1;
    s8.addShape(pres.shapes.RECTANGLE,{x,y:1.85,w:2.9,h:2.75,fill:{color:C.lightGray},shadow:makeShadow(),line:{color:"E5E7EB",width:0.4}});
    s8.addShape(pres.shapes.RECTANGLE,{x,y:1.85,w:2.9,h:0.06,fill:{color:C.steel},line:{color:C.steel,width:0}});
    s8.addText(`${c.val.toLocaleString()}`,{x:x+0.15,y:2.0,w:2.6,h:1.0,fontSize:48,bold:true,color:C.navy,fontFace:"Georgia",margin:0});
    s8.addText(c.label,{x:x+0.15,y:3.05,w:2.6,h:0.35,fontSize:11,bold:true,color:C.navy,fontFace:"Calibri",margin:0});
    s8.addText(c.sub,  {x:x+0.15,y:3.42,w:2.6,h:0.85,fontSize:10,color:C.midGray,fontFace:"Calibri",margin:0});
  });
  const boltIcon=await iconPng(FaBolt,"#"+C.gold,128);
  s8.addShape(pres.shapes.RECTANGLE,{x:0.4,y:4.73,w:9.2,h:0.52,fill:{color:C.white},line:{color:C.gold,width:1.5}});
  s8.addImage({data:boltIcon,x:0.58,y:4.87,w:0.25,h:0.25});
  s8.addText([
    {text:"The win: ",options:{bold:true,color:C.navy}},
    {text:"these are template patterns, so fixing the templates clears most of them in bulk. Strengthen the money pages — practice-area and location pages — before the blog.",options:{color:C.dark}},
  ],{x:1.0,y:4.77,w:8.4,h:0.44,fontSize:10.5,fontFace:"Calibri"});
  footerLine(s8,pb,dom);

  // SLIDE 9: VOLUME CHART
  const s9=pres.addSlide();
  s9.background={color:C.white};
  sectionLabel(s9,"WHERE THE VOLUME IS");
  slideTitle(s9,"Fixable issues by pages affected.");
  s9.addText("Bar length shows scale. Color shows severity. The biggest counts are the easiest wins.",
    {x:0.5,y:1.42,w:9,h:0.3,fontSize:11,color:C.midGray,fontFace:"Calibri"});
  const chartBar=await barChartPng(data);
  s9.addImage({data:chartBar,x:0.5,y:1.8,w:7.0,h:3.2});
  [{color:C.red,label:"Error"},{color:C.gold,label:"Warning"},{color:C.steel,label:"Notice"}].forEach((l,i)=>{
    s9.addShape(pres.shapes.RECTANGLE,{x:7.8,y:2.05+i*0.38,w:0.2,h:0.2,fill:{color:l.color},line:{color:l.color,width:0}});
    s9.addText(l.label,{x:8.08,y:2.0+i*0.38,w:1.5,h:0.3,fontSize:10,color:C.dark,fontFace:"Calibri"});
  });
  footerLine(s9,pb,dom);

  // SLIDE 10: ACTION PLAN
  const s10=pres.addSlide();
  s10.background={color:C.white};
  sectionLabel(s10,"PRIORITY ACTION PLAN");
  slideTitle(s10,"Ordered by impact against effort.");
  const actions=narrative.actions||[];
  const aColors=[C.red,C.gold,C.gold,C.steel,C.steel,C.midGray];
  actions.forEach((a,i)=>{
    const col=i%2, row=Math.floor(i/2);
    const x=0.4+col*4.85, y=1.75+row*1.22;
    s10.addShape(pres.shapes.RECTANGLE,{x,y,w:4.65,h:1.08,fill:{color:C.white},shadow:makeShadow(),line:{color:"E5E7EB",width:0.5}});
    s10.addShape(pres.shapes.RECTANGLE,{x,y,w:0.55,h:1.08,fill:{color:aColors[i]||C.steel},line:{color:aColors[i]||C.steel,width:0}});
    s10.addText(a.n,     {x,y,w:0.55,h:1.08,fontSize:22,bold:true,color:C.white,fontFace:"Georgia",align:"center",valign:"middle"});
    s10.addText(a.title, {x:x+0.65,y:y+0.07,w:3.85,h:0.34,fontSize:12,bold:true,color:C.navy,fontFace:"Calibri",margin:0});
    s10.addText(a.body,  {x:x+0.65,y:y+0.4, w:3.85,h:0.3, fontSize:10,color:C.dark,fontFace:"Calibri",margin:0});
    s10.addText(`${a.impact||""}   ${a.effort||""}`,{x:x+0.65,y:y+0.72,w:3.85,h:0.25,fontSize:9,color:C.midGray,fontFace:"Calibri",bold:true,margin:0});
  });
  footerLine(s10,pb,dom);

  // SLIDE 11: TRAJECTORY
  const s11=pres.addSlide();
  s11.background={color:C.white};
  sectionLabel(s11,"WHAT MOVES WHEN WE FIX IT");
  slideTitle(s11,"Headroom, then momentum.");
  s11.addText("SITE HEALTH TRAJECTORY",{x:0.5,y:1.48,w:5,h:0.22,fontSize:9,bold:true,color:C.gold,charSpacing:3,fontFace:"Calibri"});
  s11.addText(`${data.siteHealth||0}`,{x:0.5,y:1.75,w:1.1,h:0.65,fontSize:40,bold:true,color:C.navy,fontFace:"Georgia"});
  const arr2=await iconPng(FaArrowRight,"#"+C.navy,128);
  s11.addImage({data:arr2,x:1.78,y:1.95,w:0.38,h:0.28});
  s11.addText("90+",{x:2.35,y:1.7,w:1.5,h:0.65,fontSize:40,bold:true,color:C.green,fontFace:"Georgia"});
  s11.addText("today",{x:0.5,y:2.45,w:1.1,h:0.22,fontSize:9,color:C.midGray,fontFace:"Calibri"});
  s11.addText("target as warnings clear",{x:2.35,y:2.45,w:2.5,h:0.22,fontSize:9,color:C.midGray,fontFace:"Calibri"});
  const barW=4.5,pct=(data.siteHealth||0)/100;
  s11.addShape(pres.shapes.RECTANGLE,{x:0.5,y:2.78,w:barW,h:0.18,fill:{color:"E5E7EB"},line:{color:"E5E7EB",width:0}});
  s11.addShape(pres.shapes.RECTANGLE,{x:0.5,y:2.78,w:barW*pct,h:0.18,fill:{color:C.navy},line:{color:C.navy,width:0}});
  s11.addShape(pres.shapes.RECTANGLE,{x:0.5+barW*pct,y:2.78,w:barW*0.1,h:0.18,fill:{color:C.green},line:{color:C.green,width:0}});
  const benefits=[
    {title:"AI Search lifts first",        body:`The schema fix targets exactly what the ${data.aiReadiness||0} score measures.`},
    {title:"CTR climbs on current ranks",  body:"Metadata work earns clicks without new ranking pages."},
    {title:"Hundreds cleared per change",  body:"Template fixes resolve issues in bulk, not one by one."},
    {title:"Health trends to 90+",         body:"As warnings close, the headline score follows."},
  ];
  benefits.forEach((b,i)=>{
    const col=i%2,row=Math.floor(i/2);
    const x=0.4+col*4.85, y=3.15+row*1.0;
    s11.addShape(pres.shapes.RECTANGLE,{x,y,w:4.65,h:0.85,fill:{color:C.lightGray},shadow:makeShadow(),line:{color:"E5E7EB",width:0.3}});
    s11.addShape(pres.shapes.OVAL,{x:x+0.15,y:y+0.27,w:0.3,h:0.3,fill:{color:C.gold},line:{color:C.gold,width:0}});
    s11.addText(b.title,{x:x+0.58,y:y+0.08,w:3.9,h:0.3, fontSize:11,bold:true,color:C.navy,fontFace:"Calibri",margin:0});
    s11.addText(b.body, {x:x+0.58,y:y+0.42,w:3.9,h:0.35,fontSize:10,color:C.dark,fontFace:"Calibri",margin:0});
  });
  footerLine(s11,pb,dom);

  // SLIDE 12: SEQUENCE
  const s12=pres.addSlide();
  s12.background={color:C.navy};
  s12.addShape(pres.shapes.RECTANGLE,{x:0,y:0,w:0.18,h:5.625,fill:{color:C.gold},line:{color:C.gold,width:0}});
  const flagIcon=await iconPng(FaFlag,"#"+C.gold,256);
  s12.addImage({data:flagIcon,x:0.5,y:0.55,w:0.55,h:0.55});
  s12.addText("Recommended sequence",{x:0.5,y:1.22,w:9,h:0.65,fontSize:36,bold:true,color:C.white,fontFace:"Georgia"});
  const seq=narrative.sequence||[];
  seq.forEach((t,i)=>{
    const y=2.05+i*0.72;
    s12.addShape(pres.shapes.RECTANGLE,{x:0.45,y:y+0.28,w:8.8,h:0.01,fill:{color:"2A3F5F"},line:{color:"2A3F5F",width:0}});
    s12.addText(t.week,{x:0.5,y,w:1.8,h:0.5,fontSize:13,bold:true,color:C.gold,fontFace:"Calibri"});
    s12.addText(t.body,{x:2.5,y,w:6.8,h:0.5,fontSize:13,color:C.white,fontFace:"Calibri"});
  });
  s12.addText("Full affected-URL lists are delivered in the written audit report.",
    {x:0.5,y:5.1,w:9,h:0.3,fontSize:10,color:"4A6A85",fontFace:"Calibri",italic:true});

  // RETURN AS BASE64
  const base64 = await pres.write({ outputType: "base64" });
  return base64;
}

// ── ROUTE ────────────────────────────────────────────────────
app.post("/generate", async (req, res) => {
  try {
    const { data, narrative } = req.body;
    if (!data) return res.status(400).json({ error: "Missing data" });
    const pptxBase64 = await buildPptx(data, narrative || {});
    res.json({ pptxBase64 });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message });
  }
});

app.get("/", (req, res) => res.send("SEO Audit API running."));

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Listening on port ${PORT}`));
