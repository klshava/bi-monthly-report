const pptxgen = require("pptxgenjs");

const PURPLE = "5620AD";
const PURPLE_DARK = "3F1580";
const PURPLE_MID = "7B4CD6";
const YELLOW = "FFE800";
const WHITE = "FFFFFF";
const CARD_BG = "6A35BE"; // approximation of rgba(255,255,255,0.08) over purple
const CARD_BORDER = "8257C9"; // approximation of rgba(255,255,255,0.15) over purple

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.333 x 7.5 in, 16:9

const SLIDE_W = 13.333;
const SLIDE_H = 7.5;

function addBase(slide, pageNum) {
  slide.background = { color: PURPLE };
  slide.addImage({ path: "GL.png", x: 0.5, y: 0.35, w: 0.3, h: 0.3 });
  slide.addText("DABBLE", {
    x: 0.86, y: 0.33, w: 1.5, h: 0.34,
    fontFace: "Arial", fontSize: 12, bold: true, color: WHITE,
    charSpacing: 2, margin: 0, valign: "middle",
  });
  slide.addText(pageNum, {
    x: SLIDE_W - 1.0, y: SLIDE_H - 0.45, w: 0.6, h: 0.3,
    fontFace: "Arial", fontSize: 10, color: WHITE, transparency: 40,
    align: "right", margin: 0,
  });
}

function title(slide, text, opts = {}) {
  slide.addText(text.toUpperCase(), {
    x: 0.5, y: opts.y || 0.95, w: opts.w || 12.3, h: opts.h || 0.7,
    fontFace: "Arial", fontSize: opts.fontSize || 30, bold: true, color: YELLOW,
    charSpacing: 0.5, margin: 0,
  });
}

function subtitle(slide, text, opts = {}) {
  slide.addText(text, {
    x: 0.5, y: opts.y || 1.55, w: opts.w || 12.3, h: opts.h || 0.4,
    fontFace: "Arial", fontSize: 14, color: WHITE, transparency: 10, margin: 0,
  });
}

// ---------- SLIDE 1: COVER ----------
{
  const slide = pres.addSlide();
  addBase(slide, "01");

  slide.addText("JULY '26", {
    x: 0.5, y: 2.1, w: 4, h: 0.35,
    fontFace: "Arial", fontSize: 13, color: WHITE, transparency: 15, charSpacing: 2, margin: 0,
  });
  slide.addText("BI TEAM\nMONTHLY UPDATE", {
    x: 0.5, y: 2.45, w: 6.2, h: 1.7,
    fontFace: "Arial", fontSize: 40, bold: true, color: YELLOW, margin: 0, lineSpacing: 42,
  });
  slide.addText("A look back at what the core BI team shipped, solved and supported across July.", {
    x: 0.5, y: 4.25, w: 5.6, h: 0.8,
    fontFace: "Arial", fontSize: 14, color: WHITE, transparency: 10, margin: 0,
  });

  slide.addImage({
    path: "team photo.png",
    x: 7.3, y: 2.15, w: 5.5, h: 2.87,
    sizing: { type: "cover", w: 5.5, h: 2.87 },
    rounding: false,
  });
}

// ---------- SLIDE 2: WHO WE ARE ----------
{
  const slide = pres.addSlide();
  addBase(slide, "02");
  title(slide, "Who We Are", { y: 1.7 });
  subtitle(slide, "The core BI team behind Dabble's reporting and analytics", { y: 2.35 });

  slide.addText(
    [
      { text: "We own Board and exec reporting, regulatory & compliance dashboards, and marketing performance views.", options: { bullet: true, breakLine: true, paraSpaceAfter: 12 } },
      { text: "We're building out a growing self-service reporting layer for the wider business.", options: { bullet: true } },
    ],
    { x: 0.5, y: 2.95, w: 11.5, h: 1.3, fontFace: "Arial", fontSize: 15, color: WHITE, margin: 0, lineSpacing: 22 }
  );

  const names = [
    ["M", "Madi"], ["B", "Brett"], ["T", "Toby"], ["G", "Gaz"], ["K", "Kudzai"],
  ];
  let cx = 0.5;
  const chipY = 4.7;
  names.forEach(([initial, name]) => {
    const chipW = 1.5;
    slide.addShape("roundRect", {
      x: cx, y: chipY, w: chipW, h: 0.6, rectRadius: 0.3,
      fill: { color: CARD_BG }, line: { color: CARD_BORDER, width: 1 },
    });
    slide.addShape("ellipse", {
      x: cx + 0.08, y: chipY + 0.08, w: 0.44, h: 0.44,
      fill: { color: YELLOW }, line: { type: "none" },
    });
    slide.addText(initial, {
      x: cx + 0.08, y: chipY + 0.08, w: 0.44, h: 0.44,
      fontFace: "Arial", fontSize: 13, bold: true, color: PURPLE_DARK,
      align: "center", valign: "middle", margin: 0,
    });
    slide.addText(name, {
      x: cx + 0.58, y: chipY, w: chipW - 0.6, h: 0.6,
      fontFace: "Arial", fontSize: 12, bold: true, color: WHITE,
      valign: "middle", margin: 0,
    });
    cx += chipW + 0.25;
  });
}

// ---------- SLIDE 3: TEAM STATS ----------
{
  const slide = pres.addSlide();
  addBase(slide, "03");
  title(slide, "July 2026: Team Stats", { y: 1.7 });
  subtitle(slide, "Volume across Insights Requests, Helpdesks and Initiatives", { y: 2.35 });

  const stats = [
    { icon: "\u{1F3C6}", value: "36", caption: "IRs Completed", trend: "▲ 20% vs June", color: "7CE38B" },
    { icon: "\u{1F680}", value: "45", caption: "Helpdesks Resolved", trend: "▼ 18% vs June", color: "FF8A8A" },
    { icon: "\u{1F501}", value: "8", caption: "Initiatives / Recurring", trend: "◆ Flat vs June", color: "DDDDDD" },
    { icon: "⚡", value: "92", caption: "Total Tasks Shipped", trend: "◆ Flat vs June", color: "DDDDDD" },
  ];

  const cardW = 2.7, gap = 0.35;
  const totalW = cardW * 4 + gap * 3;
  let cx = (SLIDE_W - totalW) / 2;
  const cardY = 3.0, cardH = 3.1;

  stats.forEach((s) => {
    slide.addShape("roundRect", {
      x: cx, y: cardY, w: cardW, h: cardH, rectRadius: 0.12,
      fill: { color: CARD_BG }, line: { color: CARD_BORDER, width: 1 },
    });
    const badgeSize = 0.95;
    slide.addShape("ellipse", {
      x: cx + (cardW - badgeSize) / 2, y: cardY + 0.35, w: badgeSize, h: badgeSize,
      fill: { color: PURPLE_MID }, line: { color: YELLOW, width: 2 },
    });
    slide.addText(s.icon, {
      x: cx + (cardW - badgeSize) / 2, y: cardY + 0.35, w: badgeSize, h: badgeSize,
      fontSize: 30, align: "center", valign: "middle", margin: 0,
    });
    slide.addText(s.value, {
      x: cx, y: cardY + 1.45, w: cardW, h: 0.6,
      fontFace: "Arial", fontSize: 34, bold: true, color: YELLOW,
      align: "center", margin: 0,
    });
    slide.addText(s.caption, {
      x: cx + 0.15, y: cardY + 2.05, w: cardW - 0.3, h: 0.4,
      fontFace: "Arial", fontSize: 11.5, color: WHITE, transparency: 10,
      align: "center", margin: 0,
    });
    slide.addShape("roundRect", {
      x: cx + 0.35, y: cardY + 2.5, w: cardW - 0.7, h: 0.35, rectRadius: 0.17,
      fill: { color: "FFFFFF", transparency: 88 }, line: { type: "none" },
    });
    slide.addText(s.trend, {
      x: cx + 0.35, y: cardY + 2.5, w: cardW - 0.7, h: 0.35,
      fontFace: "Arial", fontSize: 10.5, bold: true, color: s.color,
      align: "center", valign: "middle", margin: 0,
    });
    cx += cardW + gap;
  });
}

// ---------- SLIDE 4: TEAM HIGHLIGHTS ----------
{
  const slide = pres.addSlide();
  addBase(slide, "04");
  title(slide, "Team Highlights: July", { y: 1.7 });
  subtitle(slide, "What each of us shipped, solved and supported", { y: 2.35 });

  const people = [
    { icon: "\u{1F9FE}", name: "Madi", role: "Board & US CPA", tag: "P1 / P2", desc: "Ran the ongoing Board Reporting series across AUS, USA and UK, and rebuilt the US CPA warehouse logic for Marketing." },
    { icon: "\u{1F3C7}", name: "Brett", role: "Racing & World Cup", tag: "P1", desc: "Closed out the AU Racing Body Fees rebuild and wrapped World Cup Wrapped reporting end-to-end for CRM." },
    { icon: "\u{1F50D}", name: "Toby", role: "FY26 Audit Season", tag: "P1", desc: "Owned the Rocket Boost & Pending Bets FY26 Audits, plus a tight-turnaround Bonus Payout Extract." },
    { icon: "\u{1F3C8}", name: "Gaz", role: "NFL Season 1 Review", tag: "Delivered", desc: "Shipped the NFL Season 1 YoY review (deck, Excel model and QuickSight backbone), then kicked off the Dabble Day Dashboard." },
    { icon: "\u{1F6E1}", name: "Kudzai", role: "Escheatment & Compliance", tag: "Proactive", desc: "Built a dashboard tracking unclaimed-fund balances against state dormancy thresholds across all US states." },
  ];

  const cardW = 3.85, cardH = 1.95, gapX = 0.25, gapY = 0.25;
  const row1 = people.slice(0, 3);
  const row2 = people.slice(3);

  function drawCard(p, x, y) {
    slide.addShape("roundRect", {
      x, y, w: cardW, h: cardH, rectRadius: 0.1,
      fill: { color: CARD_BG }, line: { color: CARD_BORDER, width: 1 },
    });
    const iconSize = 0.5;
    slide.addShape("ellipse", {
      x: x + 0.25, y: y + 0.22, w: iconSize, h: iconSize,
      fill: { color: PURPLE_MID }, line: { color: YELLOW, width: 1.5 },
    });
    slide.addText(p.icon, {
      x: x + 0.25, y: y + 0.22, w: iconSize, h: iconSize,
      fontSize: 18, align: "center", valign: "middle", margin: 0,
    });
    slide.addText(p.name, {
      x: x + 0.88, y: y + 0.2, w: cardW - 1.1, h: 0.3,
      fontFace: "Arial", fontSize: 14, bold: true, color: WHITE, margin: 0,
    });
    slide.addText(p.role, {
      x: x + 0.88, y: y + 0.47, w: cardW - 1.1, h: 0.25,
      fontFace: "Arial", fontSize: 10, color: YELLOW, margin: 0,
    });
    slide.addShape("roundRect", {
      x: x + 0.25, y: y + 0.85, w: 1.1, h: 0.28, rectRadius: 0.14,
      fill: { color: YELLOW }, line: { type: "none" },
    });
    slide.addText(p.tag.toUpperCase(), {
      x: x + 0.25, y: y + 0.85, w: 1.1, h: 0.28,
      fontFace: "Arial", fontSize: 8.5, bold: true, color: PURPLE_DARK,
      align: "center", valign: "middle", margin: 0,
    });
    slide.addText(p.desc, {
      x: x + 0.25, y: y + 1.2, w: cardW - 0.5, h: cardH - 1.3,
      fontFace: "Arial", fontSize: 10, color: WHITE, transparency: 5, margin: 0, lineSpacing: 13,
    });
  }

  const totalRowW = cardW * 3 + gapX * 2;
  let startX = (SLIDE_W - totalRowW) / 2;
  const row1Y = 3.05;
  row1.forEach((p, i) => drawCard(p, startX + i * (cardW + gapX), row1Y));

  const totalRow2W = cardW * 2 + gapX;
  let startX2 = (SLIDE_W - totalRow2W) / 2;
  const row2Y = row1Y + cardH + gapY;
  row2.forEach((p, i) => drawCard(p, startX2 + i * (cardW + gapX), row2Y));
}

// ---------- SLIDE 5: STRATEGY ----------
{
  const slide = pres.addSlide();
  addBase(slide, "05");
  title(slide, "Strategy Update", { y: 1.9 });
  subtitle(slide, "Where BI is investing beyond the monthly grind", { y: 2.55 });

  const items = [
    "Continuing data dictionary uplift and reporting standards work.",
    "Evaluating tooling as self-service demand grows.",
    "Regular catch-ups with teams across Dabble to surface needs early. If you don't have one on the calendar and would like one, reach out.",
    "Partnering with Data Engineering and Data Modelling on the warehouse uplift and global reporting scale.",
  ];

  const cardW = 5.85, cardH = 1.55, gapX = 0.3, gapY = 0.3;
  const totalW = cardW * 2 + gapX;
  let startX = (SLIDE_W - totalW) / 2;
  const startY = 3.25;

  items.forEach((text, i) => {
    const col = i % 2, row = Math.floor(i / 2);
    const x = startX + col * (cardW + gapX);
    const y = startY + row * (cardH + gapY);
    slide.addShape("roundRect", {
      x, y, w: cardW, h: cardH, rectRadius: 0.1,
      fill: { color: CARD_BG }, line: { color: CARD_BORDER, width: 1 },
    });
    const numSize = 0.45;
    slide.addShape("ellipse", {
      x: x + 0.3, y: y + 0.3, w: numSize, h: numSize,
      fill: { color: YELLOW }, line: { type: "none" },
    });
    slide.addText(String(i + 1), {
      x: x + 0.3, y: y + 0.3, w: numSize, h: numSize,
      fontFace: "Arial", fontSize: 14, bold: true, color: PURPLE_DARK,
      align: "center", valign: "middle", margin: 0,
    });
    slide.addText(text, {
      x: x + 0.9, y: y + 0.2, w: cardW - 1.2, h: cardH - 0.4,
      fontFace: "Arial", fontSize: 12, color: WHITE, margin: 0, valign: "middle", lineSpacing: 15,
    });
  });
}

// ---------- SLIDE 6: KUDZAI SPOTLIGHT ----------
{
  const slide = pres.addSlide();
  addBase(slide, "06");

  slide.addShape("roundRect", {
    x: 0.6, y: 0.95, w: 2.1, h: 0.4, rectRadius: 0.2,
    fill: { color: YELLOW }, line: { type: "none" },
  });
  slide.addText("BI TEAM MEMBER SPOTLIGHT", {
    x: 0.6, y: 0.95, w: 2.1, h: 0.4,
    fontFace: "Arial", fontSize: 9, bold: true, color: PURPLE_DARK,
    align: "center", valign: "middle", margin: 0,
  });

  slide.addText('KUDZAI "THE LEDGER" SHAVA', {
    x: 0.6, y: 1.42, w: 5.6, h: 0.9,
    fontFace: "Arial", fontSize: 24, bold: true, color: YELLOW, margin: 0, lineSpacing: 27,
  });

  // circular photo with glow ring, left side
  const photoSize = 3.3;
  const photoX = 0.7, photoY = 2.55;
  slide.addShape("ellipse", {
    x: photoX - 0.25, y: photoY - 0.25, w: photoSize + 0.5, h: photoSize + 0.5,
    fill: { color: PURPLE_MID, transparency: 55 }, line: { type: "none" },
  });
  slide.addShape("ellipse", {
    x: photoX - 0.1, y: photoY - 0.1, w: photoSize + 0.2, h: photoSize + 0.2,
    fill: { color: YELLOW, transparency: 82 }, line: { type: "none" },
  });
  slide.addImage({
    path: "98e0b91b-4b28-43c5-bf79-4cea15136156.png",
    x: photoX, y: photoY, w: photoSize, h: photoSize,
    sizing: { type: "cover", w: photoSize, h: photoSize },
    rounding: true,
  });
  slide.addShape("ellipse", {
    x: photoX, y: photoY, w: photoSize, h: photoSize,
    fill: { type: "none" }, line: { color: YELLOW, width: 3 },
  });

  const details = [
    { h: "Origin", p: "Legend has it Kudzai was forged, not born: assembled sometime in Q2 from spare Asana GIDs and unresolved helpdesk threads. Nobody has seen the moment of creation. There are theories." },
    { h: "BI Superpower", p: "Sheer throughput, touching Compliance/RSG, Marketing and Operations across AU, US and UK. If it touched three regions and needed doing fast, it probably has Kudzai's name on it." },
    { h: "Origin Story, Corrected", p: "Actually just a very committed BI analyst from Harare with a genuine gift for making compliance data behave. The PayID name-matching fix from June still pays dividends: the false positive rate stayed down all July." },
    { h: "Interested In", p: "Turning “please just fix the false positives” into an actual root-cause fix instead of a patch. Ask about the SMSF/corporate trustee account handling if you want to see the enthusiasm in person." },
  ];

  const colW = 3.15, colGap = 0.3;
  const detailStartX = 4.55;
  const rowH = 2.1;
  details.forEach((d, i) => {
    const col = i % 2, row = Math.floor(i / 2);
    const x = detailStartX + col * (colW + colGap);
    const y = 2.55 + row * rowH;
    slide.addText(d.h.toUpperCase(), {
      x, y, w: colW, h: 0.3,
      fontFace: "Arial", fontSize: 11, bold: true, color: YELLOW, margin: 0,
    });
    slide.addText(d.p, {
      x, y: y + 0.32, w: colW, h: rowH - 0.4,
      fontFace: "Arial", fontSize: 9.5, color: WHITE, transparency: 5, margin: 0, lineSpacing: 12,
    });
  });
}

// ---------- SLIDE 7: HOW TO WORK WITH BI ----------
{
  const slide = pres.addSlide();
  addBase(slide, "07");
  title(slide, "How To Work With BI", { y: 1.7 });

  const items = [
    { icon: "\u{1F41B}", text: "For bugs or dashboard access, drop us a line in #helpdesk-data-insights." },
    { icon: "\u{1F4CB}", text: "For a data pull, new dashboard, or updates to an existing one, fill out the Data Insights Request form." },
    { icon: "\u{1F4AC}", text: "For ideas, roadmapping or just a chat, DM any of the team on Slack." },
    { icon: "⏰", text: "Early conversations beat late surprises: talk to us before things are on fire." },
  ];

  let y = 2.7;
  items.forEach((item) => {
    const iconSize = 0.5;
    slide.addShape("ellipse", {
      x: 0.5, y, w: iconSize, h: iconSize,
      fill: { color: YELLOW }, line: { type: "none" },
    });
    slide.addText(item.icon, {
      x: 0.5, y, w: iconSize, h: iconSize,
      fontSize: 20, align: "center", valign: "middle", margin: 0,
    });
    slide.addText(item.text, {
      x: 1.2, y: y - 0.05, w: 11, h: 0.6,
      fontFace: "Arial", fontSize: 14, color: WHITE, margin: 0, valign: "middle",
    });
    y += 0.95;
  });
}

// ---------- SLIDE 8: CLOSING ----------
{
  const slide = pres.addSlide();
  addBase(slide, "08");

  slide.addText("THANK YOU", {
    x: 0, y: 1.7, w: SLIDE_W, h: 0.9,
    fontFace: "Arial", fontSize: 44, bold: true, color: YELLOW,
    align: "center", margin: 0,
  });
  slide.addText("Thanks for reading the July BI Team Monthly Update. Reach out if you'd like a deeper dive into anything we covered.", {
    x: (SLIDE_W - 6.5) / 2, y: 2.7, w: 6.5, h: 0.7,
    fontFace: "Arial", fontSize: 13, color: WHITE, align: "center", margin: 0, lineSpacing: 17,
  });

  const boxW = 6.2, boxH = 1.15;
  const boxX = (SLIDE_W - boxW) / 2, boxY = 3.65;
  slide.addShape("roundRect", {
    x: boxX, y: boxY, w: boxW, h: boxH, rectRadius: 0.12,
    fill: { color: CARD_BG }, line: { color: YELLOW, width: 1, transparency: 40 },
  });
  slide.addText("NEXT UP: AUGUST", {
    x: boxX, y: boxY + 0.14, w: boxW, h: 0.25,
    fontFace: "Arial", fontSize: 10, bold: true, color: YELLOW, align: "center", margin: 0, charSpacing: 1,
  });
  slide.addText(
    "Progress on the warehouse uplift, the next round of self-service tooling decisions, and another BI Team Member Spotlight. See you then.",
    {
      x: boxX + 0.3, y: boxY + 0.4, w: boxW - 0.6, h: 0.65,
      fontFace: "Arial", fontSize: 11, color: WHITE, align: "center", margin: 0, lineSpacing: 14,
    }
  );

  slide.addText("BI  •  Data Insights  •  July 2026", {
    x: 0, y: 5.1, w: SLIDE_W, h: 0.35,
    fontFace: "Arial", fontSize: 11, color: WHITE, transparency: 30, align: "center", margin: 0,
  });
}

pres.writeFile({ fileName: "BI Monthly Reporting.pptx" }).then(() => {
  console.log("PPTX written");
});
