const pptxgen = require("pptxgenjs");

const pres = new pptxgen();
pres.layout = "LAYOUT_16x9";
pres.title = "Impact of Crude Oil on Foreign Exchange Markets";

const C = {
  navy:    "0D1F3C",
  navyMid: "132952",
  gold:    "C9943A",
  goldLt:  "E8B455",
  teal:    "1B7FA6",
  white:   "FFFFFF",
  offWht:  "F0F4F8",
  gray:    "8FA3B1",
  darkGray:"2C3E50",
  red:     "C0392B",
  green:   "1A8754",
};

const makeShadow = () => ({ type: "outer", blur: 8, offset: 3, angle: 135, color: "000000", opacity: 0.18 });

function addFooter(slide) {
  slide.addText("Financial Institutions, Markets & Services  |  Chitkara Business School  |  Jan–Jun 2026", {
    x: 0, y: 5.3, w: 10, h: 0.25,
    fontSize: 8, color: "6E8CA0", align: "center", fontFace: "Calibri"
  });
}

function darkBg(slide) { slide.background = { color: C.navy }; }
function lightBg(slide) { slide.background = { color: C.offWht }; }

function sectionDivider(slide, num, title, subtitle) {
  darkBg(slide);
  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 0.2, h: 5.625, fill: { color: C.gold } });
  slide.addText(`SECTION 0${num}`, {
    x: 0.5, y: 1.1, w: 9, h: 0.5,
    fontSize: 13, bold: true, color: C.goldLt, fontFace: "Calibri", charSpacing: 4
  });
  slide.addText(title, {
    x: 0.5, y: 1.65, w: 9, h: 1.8,
    fontSize: 44, bold: true, color: C.white, fontFace: "Georgia"
  });
  slide.addText(subtitle, {
    x: 0.5, y: 3.55, w: 9, h: 0.6,
    fontSize: 16, color: C.gray, fontFace: "Calibri", italic: true
  });
}

function addSlideTitle(slide, title, isDark) {
  const col = isDark ? C.white : C.navy;
  slide.addText(title, {
    x: 0.5, y: 0.22, w: 9, h: 0.65,
    fontSize: 26, bold: true, color: col, fontFace: "Georgia", margin: 0
  });
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.5, y: 0.88, w: 1.4, h: 0.05, fill: { color: C.gold }
  });
}

function statCard(slide, x, y, w, h, value, label, bg) {
  slide.addShape(pres.shapes.RECTANGLE, { x, y, w, h, fill: { color: bg || C.navy }, shadow: makeShadow() });
  slide.addText(value, {
    x: x + 0.1, y: y + 0.1, w: w - 0.2, h: h * 0.55,
    fontSize: 28, bold: true, color: C.goldLt, fontFace: "Georgia", align: "center", valign: "middle"
  });
  slide.addText(label, {
    x: x + 0.1, y: y + h * 0.55, w: w - 0.2, h: h * 0.42,
    fontSize: 10, color: C.white, fontFace: "Calibri", align: "center", valign: "middle", wrap: true
  });
}

// ─── SLIDE 1: TITLE ───────────────────────────────────────────────────
{
  const sl = pres.addSlide();
  darkBg(sl);
  sl.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 0.18, h: 5.625, fill: { color: C.gold } });
  sl.addShape(pres.shapes.RECTANGLE, { x: 6.6, y: 0, w: 3.4, h: 5.625, fill: { color: C.navyMid } });
  sl.addShape(pres.shapes.RECTANGLE, { x: 6.6, y: 2.6, w: 3.4, h: 0.07, fill: { color: C.gold } });
  sl.addShape(pres.shapes.OVAL, { x: 7.3, y: 0.5, w: 1.6, h: 1.6, fill: { color: C.gold, transparency: 87 } });
  sl.addShape(pres.shapes.OVAL, { x: 8.1, y: 3.4, w: 1.1, h: 1.1, fill: { color: C.teal, transparency: 82 } });

  sl.addText("FINANCIAL INSTITUTIONS, MARKETS & SERVICES", {
    x: 0.4, y: 0.4, w: 6, h: 0.35,
    fontSize: 9, bold: true, color: C.goldLt, fontFace: "Calibri", charSpacing: 2
  });
  sl.addText("Impact of", {
    x: 0.4, y: 1.0, w: 6, h: 0.65,
    fontSize: 36, color: C.white, fontFace: "Georgia", italic: true
  });
  sl.addText("Crude Oil", {
    x: 0.4, y: 1.6, w: 6, h: 1.0,
    fontSize: 56, bold: true, color: C.goldLt, fontFace: "Georgia"
  });
  sl.addText("on Foreign Exchange Markets", {
    x: 0.4, y: 2.55, w: 6, h: 0.75,
    fontSize: 24, color: C.white, fontFace: "Georgia"
  });
  sl.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 3.4, w: 2.4, h: 0.05, fill: { color: C.gold } });
  sl.addText("Petrodollar System • FX Mechanisms • INR Impact • Hedging", {
    x: 0.4, y: 3.55, w: 6, h: 0.4,
    fontSize: 11.5, color: C.gray, fontFace: "Calibri", italic: true
  });
  sl.addText("Unit 3  |  Lectures 31–33  |  24COM4373\nChitkara Business School  |  B.Com & B.Com ACCA  |  4th Semester", {
    x: 0.4, y: 4.3, w: 6, h: 0.8,
    fontSize: 10, color: C.gray, fontFace: "Calibri"
  });
  sl.addText("$7.5T", { x: 6.7, y: 0.7, w: 3.1, h: 1.0, fontSize: 50, bold: true, color: C.goldLt, fontFace: "Georgia", align: "center" });
  sl.addText("Daily FX Market Volume", { x: 6.7, y: 1.65, w: 3.1, h: 0.35, fontSize: 10, color: C.gray, fontFace: "Calibri", align: "center" });
  sl.addText("$3T", { x: 6.7, y: 2.75, w: 3.1, h: 0.9, fontSize: 46, bold: true, color: C.white, fontFace: "Georgia", align: "center" });
  sl.addText("Annual Crude Oil Market", { x: 6.7, y: 3.6, w: 3.1, h: 0.35, fontSize: 10, color: C.gray, fontFace: "Calibri", align: "center" });
  sl.addText("India imports 85%+ of crude", { x: 6.7, y: 4.2, w: 3.1, h: 0.35, fontSize: 10, color: C.goldLt, fontFace: "Calibri", align: "center", italic: true });
}

// ─── SLIDE 2: AGENDA ──────────────────────────────────────────────────
{
  const sl = pres.addSlide();
  lightBg(sl);
  addSlideTitle(sl, "Course Agenda & Learning Objectives");

  const sections = [
    { num: "01", title: "Crude Oil Markets", sub: "Introduction, Types, Structure & Pricing", color: C.gold },
    { num: "02", title: "Foreign Exchange Markets", sub: "Structure, Instruments & Rate Mechanisms", color: C.teal },
    { num: "03", title: "Crude Oil – Forex Nexus", sub: "Petrodollar, INR Impact & Hedging Strategies", color: C.navy },
    { num: "04", title: "Summary & Takeaways", sub: "Key Insights & Policy Implications", color: C.green },
  ];
  sections.forEach((s, i) => {
    const col = i % 2 === 0 ? 0.4 : 5.3;
    const row = i < 2 ? 1.1 : 3.25;
    sl.addShape(pres.shapes.RECTANGLE, { x: col, y: row, w: 4.5, h: 1.75, fill: { color: C.white }, shadow: makeShadow() });
    sl.addShape(pres.shapes.RECTANGLE, { x: col, y: row, w: 4.5, h: 0.08, fill: { color: s.color } });
    sl.addText(s.num, { x: col + 0.15, y: row + 0.15, w: 0.7, h: 0.7, fontSize: 32, bold: true, color: s.color, fontFace: "Georgia", margin: 0 });
    sl.addText(s.title, { x: col + 0.9, y: row + 0.18, w: 3.4, h: 0.42, fontSize: 15, bold: true, color: C.navy, fontFace: "Georgia", margin: 0 });
    sl.addText(s.sub, { x: col + 0.9, y: row + 0.62, w: 3.4, h: 0.8, fontSize: 11, color: C.darkGray, fontFace: "Calibri", margin: 0, wrap: true });
  });
  sl.addText("CLO-03: Analyze reliability & validity of money market instruments  |  CLO-04: Assess financing options for wealth maximization", {
    x: 0.4, y: 5.1, w: 9.2, h: 0.28,
    fontSize: 8.5, color: C.gray, fontFace: "Calibri", italic: true, align: "center"
  });
}

// ─── SLIDE 3: SECTION 1 DIVIDER ───────────────────────────────────────
{
  const sl = pres.addSlide();
  sectionDivider(sl, 1, "Crude Oil Markets", "Introduction • Types • Structure • Pricing");
}

// ─── SLIDE 4: WHAT IS CRUDE OIL ───────────────────────────────────────
{
  const sl = pres.addSlide();
  lightBg(sl);
  addSlideTitle(sl, "What is Crude Oil?");

  sl.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 1.1, w: 4.6, h: 3.9, fill: { color: C.navy }, shadow: makeShadow() });
  sl.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 1.1, w: 4.6, h: 0.07, fill: { color: C.gold } });
  sl.addText("Definition", { x: 0.6, y: 1.2, w: 4.2, h: 0.42, fontSize: 16, bold: true, color: C.goldLt, fontFace: "Georgia", margin: 0 });
  sl.addText("Crude oil is a naturally occurring, unrefined petroleum product composed of hydrocarbon deposits found underground.\n\nIt is the world's most actively traded commodity — central to global economic activity, influencing prices, exchange rates, and national GDPs.", {
    x: 0.6, y: 1.68, w: 4.2, h: 1.55,
    fontSize: 12, color: C.white, fontFace: "Calibri", wrap: true, margin: 0
  });
  sl.addShape(pres.shapes.RECTANGLE, { x: 0.65, y: 3.3, w: 4.1, h: 0.04, fill: { color: C.gold, transparency: 60 } });

  const facts = [
    ["Market Size", "~$3 Trillion / year globally"],
    ["Unit of Trade", "Barrel (1 bbl = 159 litres)"],
    ["Currency", "Priced in USD (Petrodollar)"],
    ["Top Producers", "USA, Saudi Arabia, Russia"],
  ];
  facts.forEach(([lbl, val], i) => {
    sl.addText(`${lbl}:`, { x: 0.6, y: 3.38 + i * 0.38, w: 1.4, h: 0.32, fontSize: 10.5, bold: true, color: C.goldLt, fontFace: "Calibri", margin: 0 });
    sl.addText(val, { x: 2.05, y: 3.38 + i * 0.38, w: 2.8, h: 0.32, fontSize: 10.5, color: C.white, fontFace: "Calibri", margin: 0 });
  });

  [
    { v: "$3T", l: "Annual Global\nMarket Value", bg: C.navyMid },
    { v: "~100M", l: "Barrels Produced\nPer Day", bg: C.teal },
    { v: "#1", l: "Most Traded\nCommodity", bg: C.navyMid },
    { v: "USD", l: "Sole Pricing\nCurrency", bg: C.teal },
  ].forEach((s, i) => {
    statCard(sl, i % 2 === 0 ? 5.35 : 7.65, i < 2 ? 1.1 : 3.1, 2.1, 1.7, s.v, s.l, s.bg);
  });

  addFooter(sl);
}

// ─── SLIDE 5: OIL BENCHMARKS ──────────────────────────────────────────
{
  const sl = pres.addSlide();
  lightBg(sl);
  addSlideTitle(sl, "Major Crude Oil Benchmarks");

  const benchmarks = [
    { name: "Brent Crude", api: "~38° API", sulphur: "Low (Sweet)", origin: "North Sea, UK/Norway", desc: "European & African pricing benchmark — ~67% of global oil priced off Brent", color: C.navy },
    { name: "WTI (West Texas Intermediate)", api: "~39.6° API", sulphur: "Very Low (Sweet)", origin: "Midland, Texas, USA", desc: "US benchmark traded on NYMEX — often ~$2–4 cheaper than Brent crude", color: C.teal },
    { name: "Dubai / Oman Crude", api: "~31° API", sulphur: "Higher (Sour)", origin: "Middle East", desc: "Asian pricing benchmark — India's imports majorly priced off Dubai/Oman", color: C.gold },
    { name: "OPEC Basket", api: "Varies", sulphur: "Varies", origin: "Multi-country blend", desc: "Average of OPEC member crude prices — reference for OPEC production decisions", color: C.green },
  ];

  benchmarks.forEach((b, i) => {
    const y = 1.1 + i * 1.1;
    sl.addShape(pres.shapes.RECTANGLE, { x: 0.4, y, w: 9.2, h: 0.95, fill: { color: C.white }, shadow: makeShadow() });
    sl.addShape(pres.shapes.RECTANGLE, { x: 0.4, y, w: 0.12, h: 0.95, fill: { color: b.color } });
    sl.addShape(pres.shapes.OVAL, { x: 0.65, y: y + 0.22, w: 0.5, h: 0.5, fill: { color: b.color, transparency: 20 } });
    sl.addText(`${i + 1}`, { x: 0.65, y: y + 0.22, w: 0.5, h: 0.5, fontSize: 14, bold: true, color: C.white, fontFace: "Georgia", align: "center", valign: "middle" });
    sl.addText(b.name, { x: 1.3, y: y + 0.08, w: 3.2, h: 0.38, fontSize: 14, bold: true, color: C.navy, fontFace: "Georgia", margin: 0 });
    sl.addText(b.desc, { x: 1.3, y: y + 0.5, w: 3.9, h: 0.38, fontSize: 10.5, color: C.darkGray, fontFace: "Calibri", wrap: true, margin: 0 });
    sl.addShape(pres.shapes.RECTANGLE, { x: 5.5, y: y + 0.15, w: 1.4, h: 0.28, fill: { color: b.color, transparency: 85 } });
    sl.addText(b.api, { x: 5.5, y: y + 0.15, w: 1.4, h: 0.28, fontSize: 9, bold: true, color: b.color, fontFace: "Calibri", align: "center", valign: "middle" });
    sl.addShape(pres.shapes.RECTANGLE, { x: 7.1, y: y + 0.15, w: 1.5, h: 0.28, fill: { color: b.color, transparency: 85 } });
    sl.addText(b.sulphur, { x: 7.1, y: y + 0.15, w: 1.5, h: 0.28, fontSize: 9, bold: true, color: b.color, fontFace: "Calibri", align: "center", valign: "middle" });
    sl.addText(`📍 ${b.origin}`, { x: 5.5, y: y + 0.55, w: 3.6, h: 0.3, fontSize: 10, color: C.gray, fontFace: "Calibri", margin: 0 });
  });

  addFooter(sl);
}

// ─── SLIDE 6: OIL MARKET STRUCTURE ────────────────────────────────────
{
  const sl = pres.addSlide();
  lightBg(sl);
  addSlideTitle(sl, "Oil Market Structure & Key Pricing Factors");

  sl.addText("Market Participants", { x: 0.4, y: 1.1, w: 4.3, h: 0.4, fontSize: 15, bold: true, color: C.navy, fontFace: "Georgia" });
  const participants = [
    "OPEC+ — Cartel controlling ~40% of global oil supply",
    "National Oil Cos. (Saudi Aramco, ADNOC, ONGC)",
    "Intl. Oil Cos. (Shell, BP, ExxonMobil)",
    "Commodity Exchanges (NYMEX, ICE, MCX India)",
    "Hedge Funds & Speculators",
    "Refineries & End-Users (Airlines, Manufacturers)",
  ];
  participants.forEach((p, i) => {
    sl.addShape(pres.shapes.OVAL, { x: 0.4, y: 1.55 + i * 0.54, w: 0.28, h: 0.28, fill: { color: C.gold } });
    sl.addText(p, { x: 0.8, y: 1.52 + i * 0.54, w: 3.9, h: 0.35, fontSize: 11, color: C.darkGray, fontFace: "Calibri", margin: 0 });
  });

  sl.addText("Key Pricing Factors", { x: 5.2, y: 1.1, w: 4.4, h: 0.4, fontSize: 15, bold: true, color: C.navy, fontFace: "Georgia" });
  const factors = [
    { icon: "⚖️", label: "Supply & Demand", desc: "OPEC cuts/surpluses drive global prices", color: C.teal },
    { icon: "🌍", label: "Geopolitics", desc: "Conflicts, sanctions (Iran, Russia)", color: C.red },
    { icon: "💵", label: "USD Strength", desc: "Inverse: ↑ USD → ↓ Oil price", color: C.gold },
    { icon: "📈", label: "Speculative Trading", desc: "Futures & options on NYMEX / ICE", color: C.green },
    { icon: "🏭", label: "Economic Growth", desc: "Recession lowers demand; booms raise it", color: C.navyMid },
    { icon: "🛢️", label: "EIA Inventory Data", desc: "US weekly inventory reports (Wednesdays)", color: C.gray },
  ];
  factors.forEach((f, i) => {
    sl.addShape(pres.shapes.RECTANGLE, { x: 5.2, y: 1.55 + i * 0.54, w: 4.4, h: 0.44, fill: { color: C.white }, shadow: makeShadow() });
    sl.addShape(pres.shapes.RECTANGLE, { x: 5.2, y: 1.55 + i * 0.54, w: 0.07, h: 0.44, fill: { color: f.color } });
    sl.addText(f.icon, { x: 5.35, y: 1.55 + i * 0.54, w: 0.4, h: 0.44, fontSize: 14, fontFace: "Calibri", valign: "middle", align: "center" });
    sl.addText(f.label, { x: 5.8, y: 1.55 + i * 0.54, w: 1.5, h: 0.44, fontSize: 11, bold: true, color: C.navy, fontFace: "Georgia", margin: 0, valign: "middle" });
    sl.addText(f.desc, { x: 7.35, y: 1.55 + i * 0.54, w: 2.2, h: 0.44, fontSize: 9.5, color: C.darkGray, fontFace: "Calibri", margin: 0, valign: "middle", wrap: true });
  });

  addFooter(sl);
}

// ─── SLIDE 7: SECTION 2 DIVIDER ───────────────────────────────────────
{
  const sl = pres.addSlide();
  sectionDivider(sl, 2, "Foreign Exchange\nMarkets", "Structure • Instruments • Rate Mechanisms");
}

// ─── SLIDE 8: FX MARKET OVERVIEW ──────────────────────────────────────
{
  const sl = pres.addSlide();
  lightBg(sl);
  addSlideTitle(sl, "Introduction to Foreign Exchange Markets");

  const cards = [
    { title: "World's Largest Market", stat: "$7.5T", body: "Daily trading volume surpasses all global stock markets combined. Operates 24 hours, 5 days a week.", color: C.navy },
    { title: "OTC Market", stat: "Decentralized", body: "No single exchange. Trading occurs between banks, institutions & brokers via electronic networks (EBS, Reuters).", color: C.teal },
    { title: "Continuous Market", stat: "24 / 5", body: "Sydney → Tokyo → London → New York — a continuous global trading cycle, closing only on weekends.", color: C.gold },
  ];
  cards.forEach((c, i) => {
    const x = 0.4 + i * 3.15;
    sl.addShape(pres.shapes.RECTANGLE, { x, y: 1.1, w: 2.9, h: 3.65, fill: { color: c.color }, shadow: makeShadow() });
    sl.addText(c.title, { x: x + 0.15, y: 1.25, w: 2.6, h: 0.55, fontSize: 14, bold: true, color: C.white, fontFace: "Georgia", margin: 0, wrap: true });
    sl.addShape(pres.shapes.RECTANGLE, { x: x + 0.15, y: 1.85, w: 2.6, h: 0.04, fill: { color: C.white, transparency: 60 } });
    sl.addText(c.stat, { x: x + 0.1, y: 1.95, w: 2.7, h: 1.0, fontSize: 34, bold: true, color: C.goldLt, fontFace: "Georgia", align: "center", valign: "middle" });
    sl.addText(c.body, { x: x + 0.15, y: 3.05, w: 2.6, h: 1.6, fontSize: 11, color: C.white, fontFace: "Calibri", wrap: true, margin: 0 });
  });

  sl.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 4.88, w: 9.2, h: 0.3, fill: { color: C.navy, transparency: 85 } });
  sl.addText("India: RBI regulates the forex market under FEMA, 1999. Authorised Dealers (Banks) are key intermediaries.", {
    x: 0.5, y: 4.9, w: 9.0, h: 0.28,
    fontSize: 9.5, color: C.navy, fontFace: "Calibri", italic: true, align: "center"
  });
  addFooter(sl);
}

// ─── SLIDE 9: FX INSTRUMENTS ──────────────────────────────────────────
{
  const sl = pres.addSlide();
  lightBg(sl);
  addSlideTitle(sl, "Foreign Exchange Instruments");

  const instruments = [
    { name: "Spot Transactions", settlement: "T+2 days", who: "Importers / Exporters", how: "Exchange of currencies at current market (spot) rate for immediate settlement.", color: C.teal },
    { name: "Forward Contracts", settlement: "Weeks / Months", who: "Corporations, Oil PSUs", how: "Agreement to exchange currency at a future date at a rate locked today — eliminates FX uncertainty.", color: C.gold },
    { name: "Currency Futures", settlement: "Last Thursday of expiry", who: "Speculators & Hedgers", how: "Standardised forward contracts traded on exchange (NSE/BSE) for speculation & hedging.", color: C.navy },
    { name: "Currency Options", settlement: "Exercise date", who: "Exporters, Treasuries", how: "Right (not obligation) to buy/sell currency at a predetermined rate. Combines protection with flexibility.", color: C.green },
    { name: "Currency Swaps", settlement: "As per agreement", who: "MNCs & Banks", how: "Simultaneous buy & sell of same currency for different maturities — for long-term FX exposure management.", color: C.red },
  ];
  instruments.forEach((ins, i) => {
    const y = 1.1 + i * 0.88;
    sl.addShape(pres.shapes.RECTANGLE, { x: 0.4, y, w: 9.2, h: 0.78, fill: { color: C.white }, shadow: makeShadow() });
    sl.addShape(pres.shapes.RECTANGLE, { x: 0.4, y, w: 0.1, h: 0.78, fill: { color: ins.color } });
    sl.addText(`${i + 1}. ${ins.name}`, { x: 0.65, y: y + 0.05, w: 2.6, h: 0.38, fontSize: 13, bold: true, color: C.navy, fontFace: "Georgia", margin: 0 });
    sl.addText(ins.how, { x: 0.65, y: y + 0.43, w: 4.2, h: 0.3, fontSize: 10.5, color: C.darkGray, fontFace: "Calibri", wrap: true, margin: 0 });
    sl.addShape(pres.shapes.RECTANGLE, { x: 5.1, y: y + 0.12, w: 2.0, h: 0.26, fill: { color: ins.color, transparency: 85 } });
    sl.addText(`Settlement: ${ins.settlement}`, { x: 5.1, y: y + 0.12, w: 2.0, h: 0.26, fontSize: 9, bold: true, color: ins.color, fontFace: "Calibri", align: "center", valign: "middle" });
    sl.addShape(pres.shapes.RECTANGLE, { x: 7.25, y: y + 0.12, w: 2.3, h: 0.26, fill: { color: ins.color, transparency: 85 } });
    sl.addText(`Users: ${ins.who}`, { x: 7.25, y: y + 0.12, w: 2.3, h: 0.26, fontSize: 9, bold: true, color: ins.color, fontFace: "Calibri", align: "center", valign: "middle" });
  });

  addFooter(sl);
}

// ─── SLIDE 10: SECTION 3 DIVIDER ──────────────────────────────────────
{
  const sl = pres.addSlide();
  sectionDivider(sl, 3, "Crude Oil–Forex\nNexus", "Petrodollar • INR Impact • Hedging Strategies");
}

// ─── SLIDE 11: PETRODOLLAR SYSTEM ─────────────────────────────────────
{
  const sl = pres.addSlide();
  darkBg(sl);
  addSlideTitle(sl, "The Petrodollar System", true);

  sl.addText("Since 1973, global crude oil is priced and traded exclusively in US Dollars — creating a permanent structural link between oil and forex markets.", {
    x: 0.5, y: 0.95, w: 9, h: 0.42,
    fontSize: 11.5, color: C.gray, fontFace: "Calibri", italic: true
  });

  const steps = [
    { label: "Country\nNeeds Oil", detail: "Demands crude\nfrom oil exporters" },
    { label: "Buys USD in\nForex Market", detail: "Sells domestic\ncurrency for USD" },
    { label: "Pays Oil Exporter\nin USD", detail: "e.g. Saudi Aramco\nreceives USD" },
    { label: "Exporter Invests\nUSD Surplus", detail: "In US Treasuries\n& USD assets" },
  ];
  steps.forEach((s, i) => {
    const x = 0.4 + i * 2.4;
    sl.addShape(pres.shapes.RECTANGLE, { x, y: 1.5, w: 2.15, h: 1.4, fill: { color: C.navyMid }, shadow: makeShadow() });
    sl.addShape(pres.shapes.RECTANGLE, { x, y: 1.5, w: 2.15, h: 0.07, fill: { color: C.gold } });
    sl.addText(s.label, { x: x + 0.1, y: 1.6, w: 1.95, h: 0.65, fontSize: 13, bold: true, color: C.goldLt, fontFace: "Georgia", align: "center", valign: "middle", margin: 0 });
    sl.addText(s.detail, { x: x + 0.1, y: 2.28, w: 1.95, h: 0.57, fontSize: 10, color: C.gray, fontFace: "Calibri", align: "center", wrap: true });
    if (i < 3) {
      sl.addShape(pres.shapes.LINE, { x: x + 2.2, y: 2.2, w: 0.2, h: 0, line: { color: C.gold, width: 2 } });
    }
  });

  sl.addText("Implications for India", { x: 0.5, y: 3.05, w: 9, h: 0.42, fontSize: 16, bold: true, color: C.goldLt, fontFace: "Georgia" });

  const impl = [
    { icon: "💸", text: "India must earn or buy USD to pay for crude oil — creating chronic structural demand for USD" },
    { icon: "📉", text: "Rising oil prices → higher USD demand → INR weakens (depreciates against USD)" },
    { icon: "🏦", text: "India's forex reserves act as a buffer to manage sudden oil payment shocks" },
    { icon: "🔄", text: "India–Russia trade in INR/Ruble: an attempt to reduce petrodollar dependency" },
  ];
  impl.forEach((im, i) => {
    const col = i % 2 === 0 ? 0.4 : 5.1;
    const row = 3.6 + Math.floor(i / 2) * 0.65;
    sl.addShape(pres.shapes.RECTANGLE, { x: col, y: row, w: 4.5, h: 0.55, fill: { color: C.navyMid } });
    sl.addText(im.icon, { x: col + 0.1, y: row + 0.03, w: 0.45, h: 0.45, fontSize: 16, fontFace: "Calibri", valign: "middle", align: "center" });
    sl.addText(im.text, { x: col + 0.6, y: row + 0.03, w: 3.8, h: 0.45, fontSize: 10.5, color: C.white, fontFace: "Calibri", valign: "middle", wrap: true, margin: 0 });
  });

  addFooter(sl);
}

// ─── SLIDE 12: HOW OIL MOVES FX RATES ────────────────────────────────
{
  const sl = pres.addSlide();
  lightBg(sl);
  addSlideTitle(sl, "How Oil Prices Move Exchange Rates");

  // Rising oil panel
  sl.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 1.1, w: 4.4, h: 4.05, fill: { color: C.navy }, shadow: makeShadow() });
  sl.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 1.1, w: 4.4, h: 0.07, fill: { color: C.red } });
  sl.addText("⬆ Rising Oil Prices", { x: 0.55, y: 1.2, w: 4.1, h: 0.48, fontSize: 17, bold: true, color: C.white, fontFace: "Georgia", margin: 0 });
  sl.addText("Impact on OIL-IMPORTING Nations (India, Japan, EU)", { x: 0.55, y: 1.7, w: 4.1, h: 0.32, fontSize: 10, color: C.gray, fontFace: "Calibri", italic: true, margin: 0 });
  ["🛢️  Oil prices surge → import bill rises", "💵  Importer buys more USD to pay oil bill", "📉  Domestic currency (INR) weakens vs USD", "🏷️  Imported goods costlier → imported inflation", "📊  Current Account Deficit (CAD) widens", "🏦  RBI may hike rates or sell USD to defend INR"].forEach((s, i) => {
    sl.addText(s, { x: 0.55, y: 2.1 + i * 0.48, w: 4.1, h: 0.4, fontSize: 11, color: C.white, fontFace: "Calibri", margin: 0 });
  });

  // Falling oil panel
  sl.addShape(pres.shapes.RECTANGLE, { x: 5.2, y: 1.1, w: 4.4, h: 4.05, fill: { color: C.navyMid }, shadow: makeShadow() });
  sl.addShape(pres.shapes.RECTANGLE, { x: 5.2, y: 1.1, w: 4.4, h: 0.07, fill: { color: C.green } });
  sl.addText("⬇ Falling Oil Prices", { x: 5.35, y: 1.2, w: 4.1, h: 0.48, fontSize: 17, bold: true, color: C.white, fontFace: "Georgia", margin: 0 });
  sl.addText("Impact on OIL-EXPORTING Nations (Saudi, Russia, UAE)", { x: 5.35, y: 1.7, w: 4.1, h: 0.32, fontSize: 10, color: C.gray, fontFace: "Calibri", italic: true, margin: 0 });
  ["🛢️  Oil revenue declines for exporters", "💰  Fewer USD petrodollar inflows", "📉  Exporter currency (Ruble, Riyal) weakens", "🏛️  Sovereign wealth funds may be drawn down", "📊  Government budget deficits widen", "📉  Equities & bonds in petrostates fall"].forEach((s, i) => {
    sl.addText(s, { x: 5.35, y: 2.1 + i * 0.48, w: 4.1, h: 0.4, fontSize: 11, color: C.white, fontFace: "Calibri", margin: 0 });
  });

  addFooter(sl);
}

// ─── SLIDE 13: INDIA'S CRUDE LANDSCAPE ────────────────────────────────
{
  const sl = pres.addSlide();
  lightBg(sl);
  addSlideTitle(sl, "India's Crude Oil Landscape & INR Vulnerability");

  [
    { v: "85%+", l: "Import\nDependence", bg: C.navy },
    { v: "#3", l: "World's 3rd Largest\nConsumer", bg: C.teal },
    { v: "$140B+", l: "Annual Crude\nBill (USD)", bg: C.navy },
    { v: "Russia", l: "Top Supplier\n(2024)", bg: C.teal },
  ].forEach((s, i) => { statCard(sl, 0.4 + i * 2.35, 1.1, 2.1, 1.5, s.v, s.l, s.bg); });

  sl.addText("Key Indian Institutions in the Crude Oil Sector", { x: 0.4, y: 2.78, w: 5.6, h: 0.38, fontSize: 14, bold: true, color: C.navy, fontFace: "Georgia" });

  sl.addTable([
    [{ text: "Institution", options: { bold: true, color: C.white, fill: { color: C.navy } } }, { text: "Role", options: { bold: true, color: C.white, fill: { color: C.navy } } }],
    ["ONGC", "Exploration & production of domestic crude oil"],
    ["Indian Oil Corporation (IOC)", "Refining, marketing & pipeline management"],
    ["Bharat Petroleum (BPCL)", "Refining & retail petroleum products"],
    ["PPAC", "Monitors oil prices, imports & subsidies"],
  ], {
    x: 0.4, y: 3.22, w: 5.5, h: 1.95,
    fontFace: "Calibri", fontSize: 11,
    border: { pt: 0.5, color: "D0D8E0" },
    fill: { color: C.white }, colW: [2.4, 3.1], rowH: 0.38
  });

  sl.addShape(pres.shapes.RECTANGLE, { x: 6.2, y: 2.78, w: 3.4, h: 2.4, fill: { color: C.navy }, shadow: makeShadow() });
  sl.addShape(pres.shapes.RECTANGLE, { x: 6.2, y: 2.78, w: 3.4, h: 0.07, fill: { color: C.gold } });
  sl.addText("Key Insight", { x: 6.35, y: 2.9, w: 3.1, h: 0.38, fontSize: 14, bold: true, color: C.goldLt, fontFace: "Georgia" });
  sl.addText("A $10/barrel rise in crude prices widens India's CAD by ~$15 billion annually.\n\nThis directly weakens the INR and pushes up domestic inflation via rising import costs.", {
    x: 6.35, y: 3.33, w: 3.1, h: 1.72,
    fontSize: 11.5, color: C.white, fontFace: "Calibri", wrap: true
  });

  addFooter(sl);
}

// ─── SLIDE 14: BRENT CRUDE vs INR/USD CHART ───────────────────────────
{
  const sl = pres.addSlide();
  lightBg(sl);
  addSlideTitle(sl, "Brent Crude Price vs INR/USD — Historical Correlation");

  const years = ["2018", "2019", "2020", "2021", "2022", "2023", "2024"];
  const brent  = [71.0, 64.0, 42.0, 70.0, 101.0, 83.0, 80.0];
  const inrUsd = [68.4, 70.4, 74.1, 73.9, 78.6, 82.6, 83.5];

  sl.addChart(pres.charts.BAR, [{ name: "Brent Crude (USD/bbl)", labels: years, values: brent }], {
    x: 0.4, y: 1.1, w: 5.6, h: 3.5,
    barDir: "col",
    chartColors: ["1B7FA6"],
    chartArea: { fill: { color: C.white }, roundedCorners: true },
    catAxisLabelColor: "64748B", valAxisLabelColor: "64748B",
    valGridLine: { color: "E2E8F0", size: 0.5 }, catGridLine: { style: "none" },
    showValue: true, dataLabelColor: C.navy,
    showLegend: true, legendPos: "b", legendFontSize: 10,
    showTitle: true, title: "Brent Crude Oil Price (USD/bbl)", titleFontSize: 12, titleColor: C.navy,
  });

  sl.addChart(pres.charts.LINE, [{ name: "INR/USD Exchange Rate", labels: years, values: inrUsd }], {
    x: 6.1, y: 1.1, w: 3.5, h: 3.5,
    chartColors: ["C9943A"],
    lineSize: 3, lineSmooth: true,
    chartArea: { fill: { color: C.white }, roundedCorners: true },
    catAxisLabelColor: "64748B", valAxisLabelColor: "64748B",
    valGridLine: { color: "E2E8F0", size: 0.5 }, catGridLine: { style: "none" },
    showValue: true, dataLabelColor: C.navy,
    showLegend: true, legendPos: "b", legendFontSize: 10,
    showTitle: true, title: "INR/USD Rate (Annual Avg)", titleFontSize: 12, titleColor: C.navy,
    valAxisMinVal: 60,
  });

  sl.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 4.72, w: 9.2, h: 0.52, fill: { color: C.navy } });
  sl.addText("📊 Key Insight: India's CAD widens when crude prices rise. A $10/bbl increase ≈ $15B wider CAD → INR depreciates further.", {
    x: 0.55, y: 4.74, w: 9.0, h: 0.46,
    fontSize: 11.5, color: C.goldLt, fontFace: "Calibri", valign: "middle"
  });

  addFooter(sl);
}

// ─── SLIDE 15: EXCHANGE RATE MECHANISMS ───────────────────────────────
{
  const sl = pres.addSlide();
  lightBg(sl);
  addSlideTitle(sl, "Exchange Rate Mechanisms & INR Determinants");

  const regimes = [
    { name: "Fixed Rate", color: C.red, desc: "Currency pegged to another (e.g. USD). Central bank buys/sells reserves to maintain the peg.", ex: "Example: Saudi Riyal pegged to USD at 3.75" },
    { name: "Floating Rate", color: C.teal, desc: "Market forces (supply & demand) freely determine the rate. Most major currencies — USD, EUR, JPY.", ex: "Example: US Dollar, Euro, British Pound" },
    { name: "Managed Float", color: C.gold, desc: "Currency floats but central bank intervenes to prevent excessive volatility — India's system.", ex: "INR floats freely; RBI intervenes when needed" },
  ];
  regimes.forEach((r, i) => {
    const x = 0.4 + i * 3.15;
    sl.addShape(pres.shapes.RECTANGLE, { x, y: 1.1, w: 2.9, h: 1.8, fill: { color: C.white }, shadow: makeShadow() });
    sl.addShape(pres.shapes.RECTANGLE, { x, y: 1.1, w: 2.9, h: 0.1, fill: { color: r.color } });
    sl.addText(r.name, { x: x + 0.15, y: 1.22, w: 2.6, h: 0.42, fontSize: 15, bold: true, color: r.color, fontFace: "Georgia", margin: 0 });
    sl.addText(r.desc, { x: x + 0.15, y: 1.65, w: 2.6, h: 0.82, fontSize: 10.5, color: C.darkGray, fontFace: "Calibri", wrap: true, margin: 0 });
    sl.addText(r.ex, { x: x + 0.15, y: 2.52, w: 2.6, h: 0.32, fontSize: 9.5, color: C.gray, fontFace: "Calibri", italic: true, wrap: true, margin: 0 });
  });

  sl.addText("Key Determinants of INR Exchange Rate", { x: 0.4, y: 3.05, w: 9.2, h: 0.4, fontSize: 14, bold: true, color: C.navy, fontFace: "Georgia" });
  sl.addTable([
    [{ text: "Factor", options: { bold: true, color: C.white, fill: { color: C.navy } } }, { text: "Effect on INR", options: { bold: true, color: C.white, fill: { color: C.navy } } }],
    ["Rising Crude Oil Prices", "INR weakens — more USD needed to pay oil import bill ↓"],
    ["Increase in FII / FDI Investments", "INR strengthens — USD inflows increase ↑"],
    ["Higher US Interest Rates (Fed)", "INR weakens — capital flows out to USD assets ↓"],
    ["India's Export Growth", "INR strengthens — more USD earned by exporters ↑"],
  ], {
    x: 0.4, y: 3.5, w: 9.2, h: 1.82,
    fontFace: "Calibri", fontSize: 11,
    border: { pt: 0.5, color: "D0D8E0" },
    fill: { color: C.white }, colW: [4.5, 4.7], rowH: 0.36
  });

  addFooter(sl);
}

// ─── SLIDE 16: HEDGING STRATEGIES ─────────────────────────────────────
{
  const sl = pres.addSlide();
  lightBg(sl);
  addSlideTitle(sl, "Hedging Against Crude Oil–Forex Risk");

  const strategies = [
    { name: "Currency Forward Contracts", who: "Oil PSUs, Airlines (Importers)", how: "Lock in USD/INR rate today for future oil payment — eliminates forex uncertainty at settlement.", risk: "Miss gains if INR strengthens", color: C.teal },
    { name: "Currency Options (Put/Call)", who: "Exporters & Corporate Treasuries", how: "Buy USD call option. If INR weakens → exercise option. If INR strengthens → let it lapse.", risk: "Premium cost; complex pricing", color: C.gold },
    { name: "Crude Oil Futures (MCX/NYMEX)", who: "Refineries, Aviation Companies", how: "Buy crude futures to lock in procurement price — protects against sudden price spikes.", risk: "Margin calls if price falls sharply", color: C.navy },
    { name: "Natural Hedging", who: "MNCs with India Operations", how: "Match USD revenues (exports) against USD costs (oil imports) to naturally offset FX exposure.", risk: "Requires balanced exposure across geographies", color: C.green },
  ];

  strategies.forEach((s, i) => {
    const col = i % 2 === 0 ? 0.4 : 5.1;
    const row = i < 2 ? 1.1 : 3.2;
    sl.addShape(pres.shapes.RECTANGLE, { x: col, y: row, w: 4.5, h: 1.88, fill: { color: C.white }, shadow: makeShadow() });
    sl.addShape(pres.shapes.RECTANGLE, { x: col, y: row, w: 4.5, h: 0.08, fill: { color: s.color } });
    sl.addText(s.name, { x: col + 0.15, y: row + 0.12, w: 4.2, h: 0.45, fontSize: 14, bold: true, color: s.color, fontFace: "Georgia", margin: 0, wrap: true });
    sl.addText(`👤 ${s.who}`, { x: col + 0.15, y: row + 0.58, w: 4.2, h: 0.3, fontSize: 10, color: C.gray, fontFace: "Calibri", italic: true, margin: 0 });
    sl.addText(s.how, { x: col + 0.15, y: row + 0.88, w: 4.2, h: 0.58, fontSize: 11, color: C.darkGray, fontFace: "Calibri", wrap: true, margin: 0 });
    sl.addShape(pres.shapes.RECTANGLE, { x: col + 0.15, y: row + 1.52, w: 4.2, h: 0.28, fill: { color: s.color, transparency: 88 } });
    sl.addText(`⚠ Risk: ${s.risk}`, { x: col + 0.15, y: row + 1.52, w: 4.2, h: 0.28, fontSize: 9.5, color: s.color, fontFace: "Calibri", bold: true, valign: "middle", margin: 0 });
  });

  addFooter(sl);
}

// ─── SLIDE 17: CONCLUSION ──────────────────────────────────────────────
{
  const sl = pres.addSlide();
  lightBg(sl);

  // Accent bar
  sl.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 0.2, h: 5.625, fill: { color: C.gold } });

  // Header band
  sl.addShape(pres.shapes.RECTANGLE, { x: 0.2, y: 0, w: 9.8, h: 0.75, fill: { color: C.navy } });

  // Header label
  sl.addText("FINANCIAL INSTITUTIONS, MARKETS & SERVICES  |  KEY CONCLUSIONS", {
    x: 0.4, y: 0, w: 9.4, h: 0.75,
    fontSize: 10, bold: true, color: C.goldLt, fontFace: "Calibri", charSpacing: 3, valign: "middle"
  });

  // Slide title
  sl.addText("Summary & Key Takeaways", {
    x: 0.4, y: 0.81, w: 9.4, h: 0.6,
    fontSize: 35, bold: true, color: C.navy, fontFace: "Georgia", align: "center", valign: "middle"
  });

  // Divider
  sl.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 1.4, w: 3, h: 0.05, fill: { color: C.gold } });

  // Card 1: Petrodollar System
  sl.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 1.5, w: 4.4, h: 1.2, fill: { color: C.white }, shadow: makeShadow() });
  sl.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 1.5, w: 4.4, h: 0.08, fill: { color: C.navy } });
  sl.addText("Petrodollar System", {
    x: 0.6, y: 1.62, w: 4, h: 0.3,
    fontSize: 18, bold: true, color: C.navy, fontFace: "Georgia"
  });
  sl.addText("Since 1973, crude oil is priced exclusively in USD, permanently linking oil markets to global FX dynamics and creating structural USD demand worldwide.", {
    x: 0.6, y: 1.95, w: 4, h: 0.7,
    fontSize: 11, color: C.darkGray, fontFace: "Calibri", wrap: true
  });

  // Card 2: India's INR Vulnerability
  sl.addShape(pres.shapes.RECTANGLE, { x: 4.9, y: 1.5, w: 4.9, h: 1.2, fill: { color: C.white }, shadow: makeShadow() });
  sl.addShape(pres.shapes.RECTANGLE, { x: 4.9, y: 1.5, w: 4.9, h: 0.08, fill: { color: C.navyMid } });
  sl.addText("India's INR Vulnerability", {
    x: 5.1, y: 1.62, w: 4.5, h: 0.3,
    fontSize: 18, bold: true, color: C.navyMid, fontFace: "Georgia"
  });
  sl.addText("Importing 85%+ of crude, a $10/bbl price rise widens India's CAD by ~$15B, directly depreciating the INR and stoking imported inflation.", {
    x: 5.1, y: 1.95, w: 4.5, h: 0.7,
    fontSize: 11, color: C.darkGray, fontFace: "Calibri", wrap: true
  });

  // Card 3: FX Market Instruments
  sl.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 2.75, w: 4.4, h: 1.2, fill: { color: C.white }, shadow: makeShadow() });
  sl.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 2.75, w: 4.4, h: 0.08, fill: { color: C.navyMid } });
  sl.addText("FX Market Instruments", {
    x: 0.6, y: 2.87, w: 4, h: 0.3,
    fontSize: 18, bold: true, color: C.navyMid, fontFace: "Georgia"
  });
  sl.addText("Spot, forwards, futures, options and swaps form the toolkit for FX exposure in the $7.5T daily market, regulated in India under FEMA, 1999.", {
    x: 0.6, y: 3.2, w: 4, h: 0.7,
    fontSize: 11, color: C.darkGray, fontFace: "Calibri", wrap: true
  });

  // Card 4: Hedging Strategies
  sl.addShape(pres.shapes.RECTANGLE, { x: 4.9, y: 2.75, w: 4.9, h: 1.2, fill: { color: C.white }, shadow: makeShadow() });
  sl.addShape(pres.shapes.RECTANGLE, { x: 4.9, y: 2.75, w: 4.9, h: 0.08, fill: { color: C.navy } });
  sl.addText("Hedging Strategies", {
    x: 5.1, y: 2.87, w: 4.5, h: 0.3,
    fontSize: 18, bold: true, color: C.navy, fontFace: "Georgia"
  });
  sl.addText("Forwards, currency options, crude futures (MCX/NYMEX), and natural hedging are essential tools for PSUs, airlines, and corporates managing oil-FX risk.", {
    x: 5.1, y: 3.2, w: 4.5, h: 0.7,
    fontSize: 11, color: C.darkGray, fontFace: "Calibri", wrap: true
  });

  // Policy Banner
  sl.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 4.05, w: 9.4, h: 0.65, fill: { color: C.offWht } });
  sl.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 4.05, w: 9.4, h: 0.08, fill: { color: C.gold } });
  sl.addText("Policy Implication: India's managed float (RBI), forex reserves, and INR-Ruble trade initiatives are strategic responses to structural petrodollar dependency.", {
    x: 0.6, y: 4.15, w: 9, h: 0.5,
    fontSize: 14, bold: true, color: C.navy, fontFace: "Georgia", align: "center", valign: "middle", wrap: true
  });

  // Footer
  addFooter(sl);
}

// ─── SLIDE 18: THANK YOU ───────────────────────────────────────────────
{
  const sl = pres.addSlide();
  lightBg(sl);

  // Accent bar
  sl.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 0.2, h: 5.625, fill: { color: C.teal } });

  // Header band
  sl.addShape(pres.shapes.RECTANGLE, { x: 0.2, y: 0, w: 9.8, h: 0.75, fill: { color: C.offWht } });

  // Header label
  sl.addText("FINANCIAL INSTITUTIONS, MARKETS & SERVICES", {
    x: 0.4, y: 0, w: 9.4, h: 0.75,
    fontSize: 10, bold: true, color: C.navy, fontFace: "Calibri", charSpacing: 3, valign: "middle"
  });

  // Main thank you text
  sl.addText("Thank You", {
    x: 0.4, y: 1.5, w: 9.4, h: 1.5,
    fontSize: 72, bold: true, color: C.teal, fontFace: "Georgia", align: "center", valign: "middle"
  });

  // Subtitle
  sl.addText("Questions & Discussion", {
    x: 0.4, y: 3.2, w: 9.4, h: 0.5,
    fontSize: 24, color: C.teal, fontFace: "Georgia", align: "center", valign: "middle"
  });

  // Decorative divider
  sl.addShape(pres.shapes.RECTANGLE, { x: 4, y: 3.8, w: 2, h: 0.08, fill: { color: C.teal } });

  // Footer
  addFooter(sl);
}

// ─── WRITE FILE ────────────────────────────────────────────────────────
pres.writeFile({ fileName: "D:/STUDY/Financial Institutions, Markets & Service/PPT/Oil_Impact_on_FX_Markets.pptx" })
  .then(() => console.log("Done!"))
  .catch(err => console.error("Error:", err));