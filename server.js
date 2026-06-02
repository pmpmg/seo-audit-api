const express  = require("express");
const path     = require("path");
const pptxgen  = require("pptxgenjs");
const React    = require("react");
const ReactDOM = require("react-dom/server");
const sharp    = require("sharp");

const app = express();
app.use(express.json({ limit: "10mb" }));
app.use(express.static(path.join(__dirname, "public")));

// ── COLORS ───────────────────────────────────────────────────
const C = {
  navy:"1B2A4A", gold:"C8972A", red:"C0442A", steel:"4A6FA5",
  green:"2E7D52", lightGray:"F2F4F6", midGray:"8A9BB0", white:"FFFFFF", dark:"111827",
};

// ── ICONS ────────────────────────────────────────────────────
const { FaServer,FaLink,FaHeading,FaRobot,FaCode,FaTags,
        FaLayerGroup,FaArrowRight,FaBolt,FaMousePointer,
        FaClipboardList,FaFlag } = require("react-icons/fa");

async function iconPng(Icon, color, size=256) {
  const svg = ReactDOM.renderToStaticMarkup(React.createElement(Icon,{color,size:String(size)}));
  return "image/png;base64," + (await sharp(Buffer.from(svg)).png().toBuffer()).toString("base64");
}
async function donutPng(score, color) {
  const r=100,cx=150,cy=150,sw=22,c=2*Math.PI*r,d=(score/100)*c;
  const svg=`<svg xmlns="http://www.w3.org/2000/svg" width="300" height="300" viewBox="0 0 300 300">
    <circle cx="${cx}" cy="${cy}" r="${r}" fill="none" stroke="#E5E7EB" stroke-width="${sw}"/>
    <circle cx="${cx}" cy="${cy}" r="${r}" fill="none" stroke="#${color}" stroke-width="${sw}"
      stroke-dasharray="${d} ${c}" stroke-dashoffset="${c/4}" stroke-linecap="round"/></svg>`;
  return "image/png;base64,"+(await sharp(Buffer.from(svg)).resize(300,300).png().toBuffer()).toString("base64");
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
  const W=800,barH=36,gap=14,lW=200,margin=20,barMaxW=W-lW-60-margin*2;
  const H=items.length*(barH+gap)+margin*2;
  const bars=items.map((it,i)=>{
    const y=margin+i*(barH+gap), bw=Math.max(4,(it.value/max)*barMaxW);
    return `<text x="${lW-8}" y="${y+barH/2+5}" text-anchor="end" font-family="Arial" font-size="13" fill="#${C.dark}">${it.label}</text>
      <rect x="${lW}" y="${y}" width="${bw}" height="${barH}" fill="#${it.color}" rx="3"/>
      <text x="${lW+bw+8}" y="${y+barH/2+5}" font-family="Arial" font-size="13" font-weight="bold" fill="#${C.dark}">${it.value}</text>`;
  }).join("");
  const svg=`<svg xmlns="http://www.w3.org/2000/svg" width="${W}" height="${H}"><rect width="${W}" height="${H}" fill="#${C.lightGray}" rx="8"/>${bars}</svg>`;
  return "image/png;base64,"+(await sharp(Buffer.from(svg)).resize(W,H).png().toBuffer()).toString("base64");
}

const ms=()=>({type:"outer",blur:8,offset:2,angle:135,color:"000000",opacity:0.08});
const footer=(s,pb,d)=>{
  s.addText(pb||"",{x:0.4,y:5.35,w:3,h:0.2,fontSize:9,color:C.midGray,fontFace:"Calibri"});
  s.addText(d||"", {x:6.6,y:5.35,w:3,h:0.2,fontSize:9,color:C.midGray,fontFace:"Calibri",align:"right"});
  s.addShape("line",{x:0.4,y:5.32,w:9.2,h:0,line:{color:"E5E7EB",width:0.5}});
};
const slabel=(s,t)=>s.addText(t,{x:0.5,y:0.28,w:9,h:0.22,fontSize:9,color:C.gold,bold:true,charSpacing:4,fontFace:"Calibri"});
const stitle=(s,t)=>s.addText(t,{x:0.5,y:0.55,w:9,h:0.75,fontSize:34,bold:true,color:C.navy,fontFace:"Georgia"});
const scard=(p,s,x,y,w,h,num,lbl,sub,ac)=>{
  s.addShape(p.shapes.RECTANGLE,{x,y,w,h,fill:{color:C.white},shadow:ms(),line:{color:"E5E7EB",width:0.5}});
  s.addShape(p.shapes.RECTANGLE,{x,y,w,h:0.07,fill:{color:ac},line:{color:ac,width:0}});
  s.addText(num,{x:x+0.18,y:y+0.18,w:w-0.3,h:0.78,fontSize:40,bold:true,color:ac,fontFace:"Georgia",margin:0});
  s.addText(lbl,{x:x+0.18,y:y+1.0, w:w-0.3,h:0.28,fontSize:11,bold:true,color:C.navy,fontFace:"Calibri",margin:0});
  if(sub)s.addText(sub,{x:x+0.18,y:y+1.3,w:w-0.3,h:0.55,fontSize:10,color:C.midGray,fontFace:"Calibri",margin:0});
};

// ── GET NARRATIVE FROM CLAUDE ────────────────────────────────
async function getNarrative(data) {
  const prompt=`You are an SEO analyst. Return ONLY valid JSON (no markdown) with this structure:
{"executiveSummary":"one sentence","whatIsWorking":[{"title":"...","sub":"..."},{"title":"...","sub":"..."},{"title":"...","sub":"..."},{"title":"...","sub":"..."}],"patterns":[{"num":"01","title":"...","body":"..."},{"num":"02","title":"...","body":"..."},{"num":"03","title":"...","body":"..."}],"actions":[{"n":"1","title":"...","body":"...","impact":"High impact","effort":"Low effort"},{"n":"2","title":"...","body":"...","impact":"...","effort":"..."},{"n":"3","title":"...","body":"...","impact":"...","effort":"..."},{"n":"4","title":"...","body":"...","impact":"...","effort":"..."},{"n":"5","title":"...","body":"...","impact":"...","effort":"..."},{"n":"6","title":"...","body":"...","impact":"...","effort":"..."}],"sequence":[{"week":"Week 1","body":"..."},{"week":"Weeks 1–3","body":"..."},{"week":"Weeks 2–4","body":"..."},{"week":"Ongoing","body":"..."}]}
AUDIT DATA: ${JSON.stringify(data)}`;

  const res=await fetch("https://api.anthropic.com/v1/messages",{
    method:"POST",
    headers:{"Content-Type":"application/json","x-api-key":process.env.ANTHROPIC_API_KEY,"anthropic-version":"2023-06-01"},
    body:JSON.stringify({model:"claude-haiku-4-5-20251001",max_tokens:1500,messages:[{role:"user",content:prompt}]})
  });
  const json=await res.json();
  if(json.error) throw new Error(json.error.message);
  return JSON.parse(json.content[0].text.replace(/```json|```/g,"").trim());
}

// ── BUILD PPTX ───────────────────────────────────────────────
async function buildPptx(data, n) {
  const pres=new pptxgen();
  pres.layout="LAYOUT_16x9";
  const pb=data.preparedBy||"PMP Marketing Group", dom=data.domain||"";

  // S1 COVER
  const s1=pres.addSlide(); s1.background={color:C.navy};
  s1.addShape(pres.shapes.RECTANGLE,{x:0,y:0,w:0.18,h:5.625,fill:{color:C.gold},line:{color:C.gold,width:0}});
  s1.addText("TECHNICAL SEO AUDIT",{x:0.45,y:1.1,w:9.1,h:0.3,fontSize:10,color:C.gold,bold:true,charSpacing:5,fontFace:"Calibri"});
  s1.addText(n.executiveSummary||"What is working, what is costing you, and what to fix first.",{x:0.45,y:1.55,w:8.5,h:0.45,fontSize:18,color:"AABBD0",fontFace:"Calibri",italic:true});
  s1.addText(dom,{x:0.45,y:2.2,w:9,h:0.7,fontSize:42,bold:true,color:C.white,fontFace:"Georgia"});
  s1.addText(data.clientName||"",{x:0.45,y:2.95,w:9,h:0.4,fontSize:20,color:C.gold,fontFace:"Calibri"});
  s1.addShape(pres.shapes.RECTANGLE,{x:0.45,y:3.52,w:9.1,h:0.015,fill:{color:"3A5070"},line:{color:"3A5070",width:0}});
  s1.addText([{text:`Prepared by ${pb}`,options:{color:"7A9BB5"}},{text:"   ·   ",options:{color:"4A6A85"}},{text:data.date||"",options:{color:"7A9BB5"}},{text:"   ·   ",options:{color:"4A6A85"}},{text:`Source: ${data.source||"SEMrush Site Audit"}`,options:{color:"7A9BB5"}},{text:"   ·   ",options:{color:"4A6A85"}},{text:`${data.pagesCrawled||0} pages crawled`,options:{color:"7A9BB5"}}],{x:0.45,y:3.68,w:9.1,h:0.28,fontSize:10,fontFace:"Calibri"});

  // S2 BOTTOM LINE
  const s2=pres.addSlide(); s2.background={color:C.white};
  slabel(s2,"THE BOTTOM LINE"); stitle(s2,"The foundation is solid. Three patterns hold it back.");
  s2.addText(n.executiveSummary||"",{x:0.5,y:1.42,w:9,h:0.55,fontSize:12,color:C.dark,fontFace:"Calibri"});
  scard(pres,s2,0.4,2.1,2.9,2.15,`${data.siteHealth||0}/100`,"SITE HEALTH","Solid base, clear ceiling to climb.",C.navy);
  scard(pres,s2,3.55,2.1,2.9,2.15,`${data.schemaErrors||0}`,"SCHEMA ERRORS","One broken template, repeated sitewide.",C.red);
  scard(pres,s2,6.7,2.1,2.9,2.15,`${(data.missingDesc||0)+(data.titlesTooLong||0)}`,"METADATA GAPS",`${data.missingDesc||0} missing descriptions, ${data.titlesTooLong||0} titles cut off.`,C.gold);
  footer(s2,pb,dom);

  // S3 SCORECARD
  const s3=pres.addSlide(); s3.background={color:C.white};
  slabel(s3,"SITE HEALTH SCORECARD"); stitle(s3,`Two scores, ${data.pagesCrawled||0} pages of evidence.`);
  s3.addImage({data:await donutPng(data.siteHealth||0,C.navy),x:0.4,y:1.35,w:2.2,h:2.2});
  s3.addText(`${data.siteHealth||0}`,{x:0.4,y:1.9,w:2.2,h:1.1,fontSize:44,bold:true,color:C.navy,fontFace:"Georgia",align:"center"});
  s3.addText("/100",{x:0.4,y:2.85,w:2.2,h:0.3,fontSize:14,color:C.midGray,fontFace:"Calibri",align:"center"});
  s3.addText("Site Health",{x:0.4,y:3.6,w:2.2,h:0.3,fontSize:12,bold:true,color:C.navy,fontFace:"Calibri",align:"center"});
  s3.addImage({data:await donutPng(data.aiReadiness||0,C.gold),x:3.1,y:1.35,w:2.2,h:2.2});
  s3.addText(`${data.aiReadiness||0}`,{x:3.1,y:1.9,w:2.2,h:1.1,fontSize:44,bold:true,color:C.gold,fontFace:"Georgia",align:"center"});
  s3.addText("/100",{x:3.1,y:2.85,w:2.2,h:0.3,fontSize:14,color:C.midGray,fontFace:"Calibri",align:"center"});
  s3.addText("AI Search Readiness",{x:3.1,y:3.6,w:2.2,h:0.3,fontSize:12,bold:true,color:C.navy,fontFace:"Calibri",align:"center"});
  s3.addText("AI Search has the most room to move. Clean schema and descriptive links are exactly what it measures.",{x:0.4,y:4.0,w:5.0,h:0.6,fontSize:10,color:C.midGray,fontFace:"Calibri",italic:true,align:"center"});
  s3.addText("Status of every crawled URL",{x:5.7,y:1.1,w:4.1,h:0.25,fontSize:10,bold:true,color:C.gold,fontFace:"Calibri",charSpacing:2});
  [{val:data.pagesCrawled||0,label:"Pages crawled",color:C.navy},{val:data.pages200||0,label:"Returning 200 OK",color:C.green},{val:data.redirects||0,label:"Redirects (3xx)",color:C.gold},{val:data.errors||0,label:"Error pages (4xx)",color:C.red}].forEach((c,i)=>{
    const x=5.7+(i%2)*2.07, y=1.3+Math.floor(i/2)*1.12;
    s3.addShape(pres.shapes.RECTANGLE,{x,y,w:1.95,h:1.0,fill:{color:C.white},shadow:ms(),line:{color:"E5E7EB",width:0.5}});
    s3.addShape(pres.shapes.RECTANGLE,{x,y,w:0.07,h:1.0,fill:{color:c.color},line:{color:c.color,width:0}});
    s3.addText(`${c.val}`,{x:x+0.18,y:y+0.1,w:1.7,h:0.52,fontSize:30,bold:true,color:c.color,fontFace:"Georgia",margin:0});
    s3.addText(c.label,{x:x+0.18,y:y+0.62,w:1.7,h:0.3,fontSize:10,color:C.midGray,fontFace:"Calibri",margin:0});
  });
  footer(s3,pb,dom);

  // S4 WORKING
  const s4=pres.addSlide(); s4.background={color:C.white};
  slabel(s4,"WHAT IS ALREADY WORKING"); stitle(s4,"The expensive stuff is right.");
  s4.addText("These are the issues that cost the most to fix after the fact. None of them are here.",{x:0.5,y:1.42,w:9,h:0.35,fontSize:12,color:C.dark,fontFace:"Calibri"});
  const wins=n.whatIsWorking||[{title:"Zero server errors",sub:`No 5xx failures across ${data.pagesCrawled||0} pages.`},{title:"No broken internal links",sub:"Internal link graph is intact."},{title:"Every page has a title",sub:"No missing or empty title tags."},{title:"Nothing blocks the crawl",sub:"Core pages are fully accessible."}];
  for(let i=0;i<4;i++){
    const x=0.4+i*2.35;
    const icons=[FaServer,FaLink,FaHeading,FaRobot];
    s4.addShape(pres.shapes.RECTANGLE,{x,y:1.9,w:2.2,h:2.7,fill:{color:C.lightGray},shadow:ms(),line:{color:"E5E7EB",width:0.3}});
    s4.addShape(pres.shapes.OVAL,{x:x+0.75,y:2.05,w:0.7,h:0.7,fill:{color:"D1FAE5"},line:{color:"D1FAE5",width:0}});
    s4.addImage({data:await iconPng(icons[i],"#"+C.green),x:x+0.83,y:2.13,w:0.54,h:0.54});
    s4.addText(wins[i]?.title||"",{x:x+0.14,y:2.85,w:1.92,h:0.55,fontSize:12,bold:true,color:C.navy,fontFace:"Calibri",align:"center"});
    s4.addText(wins[i]?.sub||"",{x:x+0.14,y:3.45,w:1.92,h:0.7,fontSize:10,color:C.midGray,fontFace:"Calibri",align:"center"});
  }
  footer(s4,pb,dom);

  // S5 PATTERNS
  const s5=pres.addSlide(); s5.background={color:C.white};
  slabel(s5,"WHAT IS HOLDING IT BACK"); stitle(s5,"Three patterns, not a hundred problems.");
  const pats=n.patterns||[{num:"01",title:"Broken structured data",body:`${data.schemaErrors||0} schema errors from one broken template.`},{num:"02",title:"Metadata gaps",body:`${data.missingDesc||0} pages with no description and ${data.titlesTooLong||0} titles too long.`},{num:"03",title:"Thin signals",body:`${data.thinPages||0} light pages with weak anchor text.`}];
  const pcols=[C.red,C.gold,C.steel], picons=[FaCode,FaTags,FaLayerGroup];
  for(let i=0;i<3;i++){
    const x=0.4+i*3.1, p=pats[i];
    s5.addShape(pres.shapes.RECTANGLE,{x,y:1.85,w:2.9,h:2.9,fill:{color:C.white},shadow:ms(),line:{color:"E5E7EB",width:0.5}});
    s5.addShape(pres.shapes.RECTANGLE,{x,y:1.85,w:2.9,h:0.72,fill:{color:pcols[i]},line:{color:pcols[i],width:0}});
    s5.addImage({data:await iconPng(picons[i],"#FFFFFF"),x:x+0.18,y:1.98,w:0.38,h:0.38});
    s5.addText(p.num,{x:x+0.1,y:1.87,w:2.7,h:0.68,fontSize:28,bold:true,color:C.white,fontFace:"Georgia",align:"right",margin:0});
    s5.addText(p.title,{x:x+0.18,y:2.65,w:2.55,h:0.5,fontSize:13,bold:true,color:C.navy,fontFace:"Calibri"});
    s5.addText(p.body,{x:x+0.18,y:3.2,w:2.55,h:1.2,fontSize:10.5,color:C.dark,fontFace:"Calibri"});
  }
  footer(s5,pb,dom);

  // S6 PATTERN 1
  const s6=pres.addSlide(); s6.background={color:C.white};
  slabel(s6,"PATTERN 1 · ERRORS"); stitle(s6,"Structured data: one fix, sitewide.");
  s6.addShape(pres.shapes.RECTANGLE,{x:0.4,y:1.35,w:3.5,h:3.5,fill:{color:C.navy},line:{color:C.navy,width:0}});
  s6.addImage({data:await iconPng(FaCode,"#"+C.gold),x:0.75,y:1.6,w:0.65,h:0.65});
  s6.addText(`${data.schemaErrors||0}`,{x:0.4,y:2.2,w:3.5,h:1.4,fontSize:80,bold:true,color:C.gold,fontFace:"Georgia",align:"center"});
  s6.addText("schema instances failing the same check",{x:0.55,y:3.65,w:3.2,h:0.5,fontSize:10,color:"7A9BB5",fontFace:"Calibri",align:"center"});
  s6.addText("The root cause",{x:4.2,y:1.4,w:5.4,h:0.28,fontSize:10,bold:true,color:C.gold,charSpacing:2,fontFace:"Calibri"});
  s6.addText(pats[0]?.body||"",{x:4.2,y:1.73,w:5.4,h:0.7,fontSize:11,color:C.dark,fontFace:"Calibri"});
  s6.addText("Why it matters",{x:4.2,y:2.55,w:5.4,h:0.28,fontSize:10,bold:true,color:C.gold,charSpacing:2,fontFace:"Calibri"});
  s6.addText("Schema is how Google and AI answer engines read who you are and where you practice. Invalid schema forfeits rich results.",{x:4.2,y:2.88,w:5.4,h:0.75,fontSize:11,color:C.dark,fontFace:"Calibri"});
  s6.addShape(pres.shapes.RECTANGLE,{x:4.2,y:3.85,w:1.9,h:0.52,fill:{color:C.white},line:{color:C.navy,width:1.5}});
  s6.addText("Fix 1 template",{x:4.2,y:3.85,w:1.9,h:0.52,fontSize:11,bold:true,color:C.navy,fontFace:"Calibri",align:"center"});
  s6.addImage({data:await iconPng(FaArrowRight,"#"+C.navy,128),x:6.2,y:4.0,w:0.3,h:0.22});
  s6.addShape(pres.shapes.RECTANGLE,{x:6.65,y:3.85,w:1.9,h:0.52,fill:{color:C.green},line:{color:C.green,width:0}});
  s6.addText(`Clear all ${data.schemaErrors||0}`,{x:6.65,y:3.85,w:1.9,h:0.52,fontSize:11,bold:true,color:C.white,fontFace:"Calibri",align:"center"});
  footer(s6,pb,dom);

  // S7 PATTERN 2
  const s7=pres.addSlide(); s7.background={color:C.white};
  slabel(s7,"PATTERN 2 · WARNINGS"); stitle(s7,"Metadata: stop letting Google guess.");
  s7.addShape(pres.shapes.RECTANGLE,{x:0.4,y:1.5,w:4.45,h:2.1,fill:{color:C.lightGray},shadow:ms(),line:{color:"E5E7EB",width:0.5}});
  s7.addShape(pres.shapes.OVAL,{x:0.65,y:1.65,w:0.7,h:0.7,fill:{color:C.gold},line:{color:C.gold,width:0}});
  s7.addImage({data:await iconPng(FaTags,"#FFFFFF"),x:0.75,y:1.73,w:0.5,h:0.5});
  s7.addText(`${data.missingDesc||0}`,{x:1.5,y:1.6,w:2.8,h:0.85,fontSize:48,bold:true,color:C.navy,fontFace:"Georgia"});
  s7.addText("pages with no meta description",{x:0.6,y:2.5,w:4.1,h:0.3,fontSize:11,bold:true,color:C.navy,fontFace:"Calibri"});
  s7.addText("Google writes the search snippet for you, so the firm loses control of the listing.",{x:0.6,y:2.85,w:4.1,h:0.55,fontSize:10,color:C.midGray,fontFace:"Calibri"});
  s7.addShape(pres.shapes.RECTANGLE,{x:5.15,y:1.5,w:4.45,h:2.1,fill:{color:C.lightGray},shadow:ms(),line:{color:"E5E7EB",width:0.5}});
  s7.addShape(pres.shapes.OVAL,{x:5.4,y:1.65,w:0.7,h:0.7,fill:{color:C.gold},line:{color:C.gold,width:0}});
  s7.addImage({data:await iconPng(FaClipboardList,"#FFFFFF"),x:5.5,y:1.73,w:0.5,h:0.5});
  s7.addText(`${data.titlesTooLong||0}`,{x:6.25,y:1.6,w:2.8,h:0.85,fontSize:48,bold:true,color:C.navy,fontFace:"Georgia"});
  s7.addText("titles too long",{x:5.35,y:2.5,w:4.1,h:0.3,fontSize:11,bold:true,color:C.navy,fontFace:"Calibri"});
  s7.addText("Titles get cut off mid-message in search results, weakening the click.",{x:5.35,y:2.85,w:4.1,h:0.55,fontSize:10,color:C.midGray,fontFace:"Calibri"});
  s7.addShape(pres.shapes.RECTANGLE,{x:0.4,y:3.82,w:9.2,h:0.9,fill:{color:C.navy},line:{color:C.navy,width:0}});
  s7.addImage({data:await iconPng(FaMousePointer,"#"+C.gold,128),x:0.55,y:4.08,w:0.32,h:0.32});
  s7.addText([{text:"What it costs: ",options:{bold:true,color:C.gold}},{text:"click-through on pages that already rank. Descriptions are written, titles get trimmed.",options:{color:C.white}}],{x:1.0,y:3.87,w:8.4,h:0.8,fontSize:10.5,fontFace:"Calibri"});
  footer(s7,pb,dom);

  // S8 PATTERN 3
  const s8=pres.addSlide(); s8.background={color:C.white};
  slabel(s8,"PATTERN 3 · CONTENT AND LINKING"); stitle(s8,"Thin pages and silent links.");
  [{val:data.thinPages||0,label:"pages low on visible text",sub:"Light content relative to page code. Reads as thin to search engines."},{val:data.noAnchors||0,label:"links with no anchor text",sub:"Links pass no context. Mostly repeating nav, button, and icon patterns."},{val:data.weakAnchorLinks||0,label:"links with weak anchors",sub:'Anchors like "click here" that tell engines nothing about the target.'}].forEach((c,i)=>{
    const x=0.4+i*3.1;
    s8.addShape(pres.shapes.RECTANGLE,{x,y:1.85,w:2.9,h:2.75,fill:{color:C.lightGray},shadow:ms(),line:{color:"E5E7EB",width:0.4}});
    s8.addShape(pres.shapes.RECTANGLE,{x,y:1.85,w:2.9,h:0.06,fill:{color:C.steel},line:{color:C.steel,width:0}});
    s8.addText(`${c.val.toLocaleString()}`,{x:x+0.15,y:2.0,w:2.6,h:1.0,fontSize:48,bold:true,color:C.navy,fontFace:"Georgia",margin:0});
    s8.addText(c.label,{x:x+0.15,y:3.05,w:2.6,h:0.35,fontSize:11,bold:true,color:C.navy,fontFace:"Calibri",margin:0});
    s8.addText(c.sub,{x:x+0.15,y:3.42,w:2.6,h:0.85,fontSize:10,color:C.midGray,fontFace:"Calibri",margin:0});
  });
  s8.addShape(pres.shapes.RECTANGLE,{x:0.4,y:4.73,w:9.2,h:0.52,fill:{color:C.white},line:{color:C.gold,width:1.5}});
  s8.addImage({data:await iconPng(FaBolt,"#"+C.gold,128),x:0.58,y:4.87,w:0.25,h:0.25});
  s8.addText([{text:"The win: ",options:{bold:true,color:C.navy}},{text:"these are template patterns, so fixing the templates clears most of them in bulk.",options:{color:C.dark}}],{x:1.0,y:4.77,w:8.4,h:0.44,fontSize:10.5,fontFace:"Calibri"});
  footer(s8,pb,dom);

  // S9 CHART
  const s9=pres.addSlide(); s9.background={color:C.white};
  slabel(s9,"WHERE THE VOLUME IS"); stitle(s9,"Fixable issues by pages affected.");
  s9.addText("Bar length shows scale. Color shows severity. The biggest counts are the easiest wins.",{x:0.5,y:1.42,w:9,h:0.3,fontSize:11,color:C.midGray,fontFace:"Calibri"});
  s9.addImage({data:await barChartPng(data),x:0.5,y:1.8,w:7.0,h:3.2});
  [{color:C.red,label:"Error"},{color:C.gold,label:"Warning"},{color:C.steel,label:"Notice"}].forEach((l,i)=>{
    s9.addShape(pres.shapes.RECTANGLE,{x:7.8,y:2.05+i*0.38,w:0.2,h:0.2,fill:{color:l.color},line:{color:l.color,width:0}});
    s9.addText(l.label,{x:8.08,y:2.0+i*0.38,w:1.5,h:0.3,fontSize:10,color:C.dark,fontFace:"Calibri"});
  });
  footer(s9,pb,dom);

  // S10 ACTIONS
  const s10=pres.addSlide(); s10.background={color:C.white};
  slabel(s10,"PRIORITY ACTION PLAN"); stitle(s10,"Ordered by impact against effort.");
  const acts=n.actions||[], acols=[C.red,C.gold,C.gold,C.steel,C.steel,C.midGray];
  acts.forEach((a,i)=>{
    const x=0.4+(i%2)*4.85, y=1.75+Math.floor(i/2)*1.22;
    s10.addShape(pres.shapes.RECTANGLE,{x,y,w:4.65,h:1.08,fill:{color:C.white},shadow:ms(),line:{color:"E5E7EB",width:0.5}});
    s10.addShape(pres.shapes.RECTANGLE,{x,y,w:0.55,h:1.08,fill:{color:acols[i]||C.steel},line:{color:acols[i]||C.steel,width:0}});
    s10.addText(a.n,{x,y,w:0.55,h:1.08,fontSize:22,bold:true,color:C.white,fontFace:"Georgia",align:"center",valign:"middle"});
    s10.addText(a.title,{x:x+0.65,y:y+0.07,w:3.85,h:0.34,fontSize:12,bold:true,color:C.navy,fontFace:"Calibri",margin:0});
    s10.addText(a.body, {x:x+0.65,y:y+0.4, w:3.85,h:0.3, fontSize:10,color:C.dark,fontFace:"Calibri",margin:0});
    s10.addText(`${a.impact||""}   ${a.effort||""}`,{x:x+0.65,y:y+0.72,w:3.85,h:0.25,fontSize:9,color:C.midGray,fontFace:"Calibri",bold:true,margin:0});
  });
  footer(s10,pb,dom);

  // S11 TRAJECTORY
  const s11=pres.addSlide(); s11.background={color:C.white};
  slabel(s11,"WHAT MOVES WHEN WE FIX IT"); stitle(s11,"Headroom, then momentum.");
  s11.addText("SITE HEALTH TRAJECTORY",{x:0.5,y:1.48,w:5,h:0.22,fontSize:9,bold:true,color:C.gold,charSpacing:3,fontFace:"Calibri"});
  s11.addText(`${data.siteHealth||0}`,{x:0.5,y:1.75,w:1.1,h:0.65,fontSize:40,bold:true,color:C.navy,fontFace:"Georgia"});
  s11.addImage({data:await iconPng(FaArrowRight,"#"+C.navy,128),x:1.78,y:1.95,w:0.38,h:0.28});
  s11.addText("90+",{x:2.35,y:1.7,w:1.5,h:0.65,fontSize:40,bold:true,color:C.green,fontFace:"Georgia"});
  s11.addText("today",{x:0.5,y:2.45,w:1.1,h:0.22,fontSize:9,color:C.midGray,fontFace:"Calibri"});
  s11.addText("target as warnings clear",{x:2.35,y:2.45,w:2.5,h:0.22,fontSize:9,color:C.midGray,fontFace:"Calibri"});
  const bw=4.5, pct=(data.siteHealth||0)/100;
  s11.addShape(pres.shapes.RECTANGLE,{x:0.5,y:2.78,w:bw,h:0.18,fill:{color:"E5E7EB"},line:{color:"E5E7EB",width:0}});
  s11.addShape(pres.shapes.RECTANGLE,{x:0.5,y:2.78,w:bw*pct,h:0.18,fill:{color:C.navy},line:{color:C.navy,width:0}});
  s11.addShape(pres.shapes.RECTANGLE,{x:0.5+bw*pct,y:2.78,w:bw*0.1,h:0.18,fill:{color:C.green},line:{color:C.green,width:0}});
  [{title:"AI Search lifts first",body:`The schema fix targets exactly what the ${data.aiReadiness||0} score measures.`},{title:"CTR climbs on current ranks",body:"Metadata work earns clicks without new ranking pages."},{title:"Hundreds cleared per change",body:"Template fixes resolve issues in bulk, not one by one."},{title:"Health trends to 90+",body:"As warnings close, the headline score follows."}].forEach((b,i)=>{
    const x=0.4+(i%2)*4.85, y=3.15+Math.floor(i/2)*1.0;
    s11.addShape(pres.shapes.RECTANGLE,{x,y,w:4.65,h:0.85,fill:{color:C.lightGray},shadow:ms(),line:{color:"E5E7EB",width:0.3}});
    s11.addShape(pres.shapes.OVAL,{x:x+0.15,y:y+0.27,w:0.3,h:0.3,fill:{color:C.gold},line:{color:C.gold,width:0}});
    s11.addText(b.title,{x:x+0.58,y:y+0.08,w:3.9,h:0.3,fontSize:11,bold:true,color:C.navy,fontFace:"Calibri",margin:0});
    s11.addText(b.body, {x:x+0.58,y:y+0.42,w:3.9,h:0.35,fontSize:10,color:C.dark,fontFace:"Calibri",margin:0});
  });
  footer(s11,pb,dom);

  // S12 SEQUENCE
  const s12=pres.addSlide(); s12.background={color:C.navy};
  s12.addShape(pres.shapes.RECTANGLE,{x:0,y:0,w:0.18,h:5.625,fill:{color:C.gold},line:{color:C.gold,width:0}});
  s12.addImage({data:await iconPng(FaFlag,"#"+C.gold),x:0.5,y:0.55,w:0.55,h:0.55});
  s12.addText("Recommended sequence",{x:0.5,y:1.22,w:9,h:0.65,fontSize:36,bold:true,color:C.white,fontFace:"Georgia"});
  (n.sequence||[]).forEach((t,i)=>{
    const y=2.05+i*0.72;
    s12.addShape(pres.shapes.RECTANGLE,{x:0.45,y:y+0.28,w:8.8,h:0.01,fill:{color:"2A3F5F"},line:{color:"2A3F5F",width:0}});
    s12.addText(t.week,{x:0.5,y,w:1.8,h:0.5,fontSize:13,bold:true,color:C.gold,fontFace:"Calibri"});
    s12.addText(t.body,{x:2.5,y,w:6.8,h:0.5,fontSize:13,color:C.white,fontFace:"Calibri"});
  });
  s12.addText("Full affected-URL lists are delivered in the written audit report.",{x:0.5,y:5.1,w:9,h:0.3,fontSize:10,color:"4A6A85",fontFace:"Calibri",italic:true});

  return await pres.write({outputType:"base64"});
}

// ── ROUTES ───────────────────────────────────────────────────
app.get("/", (req,res) => res.sendFile(path.join(__dirname,"public","index.html")));

app.post("/generate", async (req,res) => {
  try {
    const { data } = req.body;
    if (!data) return res.status(400).json({error:"Missing data"});
    const narrative = await getNarrative(data);
    const pptxBase64 = await buildPptx(data, narrative);
    const date = data.date || new Date().toLocaleDateString("en-US",{month:"short",year:"numeric"});
    const fileName = `SEO Audit — ${data.clientName} — ${date}.pptx`;
    res.json({ pptxBase64, fileName });
  } catch(err) {
    console.error(err);
    res.status(500).json({error: err.message});
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Listening on port ${PORT}`));
