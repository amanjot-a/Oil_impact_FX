const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  HeadingLevel, AlignmentType, LevelFormat, BorderStyle, WidthType,
  ShadingType, PageNumber, PageBreak, Footer, Header, TabStopType,
  TabStopPosition, PositionalTab, PositionalTabAlignment, PositionalTabRelativeTo,
  PositionalTabLeader
} = require('docx');
const fs = require('fs');

const CONTENT_WIDTH = 9360; // US Letter, 1" margins

const border = { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" };
const borders = { top: border, bottom: border, left: border, right: border };
const headerBorder = { style: BorderStyle.SINGLE, size: 1, color: "1F4E79" };
const headerBorders = { top: headerBorder, bottom: headerBorder, left: headerBorder, right: headerBorder };

function h1(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 360, after: 160 },
    children: [new TextRun({ text, bold: true, size: 32, font: "Calibri", color: "1F4E79" })]
  });
}

function h2(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 280, after: 120 },
    children: [new TextRun({ text, bold: true, size: 26, font: "Calibri", color: "2E75B6" })]
  });
}

function h3(text) {
  return new Paragraph({
    spacing: { before: 200, after: 80 },
    children: [new TextRun({ text, bold: true, size: 24, font: "Calibri", color: "375623" })]
  });
}

function para(text, opts = {}) {
  return new Paragraph({
    spacing: { before: 80, after: 160 },
    alignment: AlignmentType.JUSTIFIED,
    children: [new TextRun({ text, size: 22, font: "Calibri", ...opts })]
  });
}

function bullet(text, ref = "bullets") {
  return new Paragraph({
    numbering: { reference: ref, level: 0 },
    spacing: { before: 40, after: 40 },
    children: [new TextRun({ text, size: 22, font: "Calibri" })]
  });
}

function pageBreak() {
  return new Paragraph({ children: [new PageBreak()] });
}

function blankLine() {
  return new Paragraph({ children: [new TextRun("")], spacing: { before: 60, after: 60 } });
}

function makeTable(headers, rows, colWidths) {
  const total = colWidths.reduce((a, b) => a + b, 0);
  const headerRow = new TableRow({
    tableHeader: true,
    children: headers.map((h, i) => new TableCell({
      borders: headerBorders,
      width: { size: colWidths[i], type: WidthType.DXA },
      shading: { fill: "1F4E79", type: ShadingType.CLEAR },
      margins: { top: 80, bottom: 80, left: 120, right: 120 },
      children: [new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [new TextRun({ text: h, bold: true, size: 20, font: "Calibri", color: "FFFFFF" })]
      })]
    }))
  });
  const dataRows = rows.map((row, ri) => new TableRow({
    children: row.map((cell, ci) => new TableCell({
      borders,
      width: { size: colWidths[ci], type: WidthType.DXA },
      shading: { fill: ri % 2 === 0 ? "EBF3FB" : "FFFFFF", type: ShadingType.CLEAR },
      margins: { top: 80, bottom: 80, left: 120, right: 120 },
      children: [new Paragraph({
        children: [new TextRun({ text: cell, size: 20, font: "Calibri" })]
      })]
    }))
  }));
  return new Table({
    width: { size: total, type: WidthType.DXA },
    columnWidths: colWidths,
    rows: [headerRow, ...dataRows]
  });
}

const doc = new Document({
  numbering: {
    config: [
      {
        reference: "bullets",
        levels: [{
          level: 0, format: LevelFormat.BULLET, text: "\u2022", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } } }
        }]
      },
      {
        reference: "numbers",
        levels: [{
          level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } } }
        }]
      }
    ]
  },
  styles: {
    default: { document: { run: { font: "Calibri", size: 22 } } },
    paragraphStyles: [
      {
        id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 32, bold: true, font: "Calibri", color: "1F4E79" },
        paragraph: { spacing: { before: 360, after: 160 }, outlineLevel: 0 }
      },
      {
        id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 26, bold: true, font: "Calibri", color: "2E75B6" },
        paragraph: { spacing: { before: 280, after: 120 }, outlineLevel: 1 }
      },
      {
        id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 24, bold: true, font: "Calibri", color: "375623" },
        paragraph: { spacing: { before: 200, after: 80 }, outlineLevel: 2 }
      }
    ]
  },
  sections: [{
    properties: {
      page: {
        size: { width: 12240, height: 15840 },
        margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 }
      }
    },
    headers: {
      default: new Header({
        children: [
          new Paragraph({
            border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: "1F4E79", space: 1 } },
            spacing: { after: 120 },
            children: [
              new TextRun({ text: "Impact of Crude Oil on Foreign Exchange Markets", size: 18, font: "Calibri", color: "1F4E79", italics: true }),
              new TextRun({ text: "  |  Financial Institutions, Markets & Services", size: 18, font: "Calibri", color: "888888" })
            ]
          })
        ]
      })
    },
    footers: {
      default: new Footer({
        children: [
          new Paragraph({
            border: { top: { style: BorderStyle.SINGLE, size: 6, color: "1F4E79", space: 1 } },
            spacing: { before: 120 },
            tabStops: [{ type: TabStopType.RIGHT, position: CONTENT_WIDTH }],
            children: [
              new TextRun({ text: "Chitkara Business School  |  B.Com & B.Com ACCA  |  4th Semester", size: 18, font: "Calibri", color: "888888" }),
              new TextRun({ text: "\tPage ", size: 18, font: "Calibri", color: "888888" }),
              new TextRun({ children: [PageNumber.CURRENT], size: 18, font: "Calibri", color: "888888" }),
              new TextRun({ text: " of ", size: 18, font: "Calibri", color: "888888" }),
              new TextRun({ children: [PageNumber.TOTAL_PAGES], size: 18, font: "Calibri", color: "888888" })
            ]
          })
        ]
      })
    },
    children: [

      // ─── TITLE PAGE ───────────────────────────────────────────────────────────
      new Paragraph({
        spacing: { before: 1200, after: 80 },
        alignment: AlignmentType.CENTER,
        children: [new TextRun({ text: "CHITKARA BUSINESS SCHOOL", size: 24, font: "Calibri", bold: true, color: "888888", allCaps: true })]
      }),
      new Paragraph({
        spacing: { before: 80, after: 80 },
        alignment: AlignmentType.CENTER,
        children: [new TextRun({ text: "B.Com & B.Com ACCA  |  4th Semester  |  Jan–Jun 2026", size: 22, font: "Calibri", color: "888888" })]
      }),
      blankLine(),
      new Paragraph({
        spacing: { before: 480, after: 160 },
        alignment: AlignmentType.CENTER,
        border: {
          top: { style: BorderStyle.SINGLE, size: 12, color: "1F4E79", space: 12 },
          bottom: { style: BorderStyle.SINGLE, size: 12, color: "1F4E79", space: 12 }
        },
        children: [
          new TextRun({ text: "Impact of Crude Oil on Foreign Exchange Markets", size: 52, bold: true, font: "Calibri", color: "1F4E79" })
        ]
      }),
      new Paragraph({
        spacing: { before: 160, after: 80 },
        alignment: AlignmentType.CENTER,
        children: [new TextRun({ text: "Petrodollar System  •  FX Mechanisms  •  INR Impact  •  Hedging Strategies", size: 24, font: "Calibri", color: "2E75B6", italics: true })]
      }),
      blankLine(),
      new Paragraph({
        spacing: { before: 480, after: 80 },
        alignment: AlignmentType.CENTER,
        children: [new TextRun({ text: "A Written Project Submitted for", size: 22, font: "Calibri", color: "555555" })]
      }),
      new Paragraph({
        spacing: { before: 40, after: 40 },
        alignment: AlignmentType.CENTER,
        children: [new TextRun({ text: "Financial Institutions, Markets & Services (24COM4373)", size: 22, bold: true, font: "Calibri", color: "1F4E79" })]
      }),
      new Paragraph({
        spacing: { before: 40, after: 40 },
        alignment: AlignmentType.CENTER,
        children: [new TextRun({ text: "Unit 3  |  Lectures 31–33", size: 22, font: "Calibri", color: "555555" })]
      }),
      blankLine(),
      blankLine(),
      new Paragraph({
        spacing: { before: 600, after: 40 },
        alignment: AlignmentType.CENTER,
        children: [new TextRun({ text: "Academic Year 2025–2026", size: 22, font: "Calibri", color: "888888" })]
      }),

      pageBreak(),

      // ─── ABSTRACT ─────────────────────────────────────────────────────────────
      h1("Abstract"),
      para("This project examines the intricate relationship between global crude oil prices and foreign exchange markets, with a specific focus on the Indian Rupee (INR). Crude oil and currency markets are two of the most interconnected forces in the global economy. The petrodollar system — established in 1973 — created a permanent structural link between oil pricing and the US Dollar, making oil price fluctuations a critical driver of exchange rate movements across the world."),
      para("The project explores the structure of crude oil markets, the functioning of global forex markets, the mechanisms through which oil prices transmit to exchange rates, India's particular vulnerability due to its heavy import dependence, and the hedging instruments available to corporates and public sector units to manage oil-forex risk. The analysis draws on data from 2018 to 2024, demonstrating a consistent correlation between rising Brent crude prices and INR depreciation."),
      para("Key findings include: (1) every USD 10/barrel rise in crude prices widens India's Current Account Deficit by approximately USD 15 billion; (2) India's managed float regime and forex reserves act as critical buffers; and (3) instruments such as forward contracts, currency options, and crude futures are essential risk management tools for oil-exposed entities."),

      pageBreak(),

      // ─── TABLE OF CONTENTS (manual) ───────────────────────────────────────────
      h1("Table of Contents"),
      ...[
        ["1.", "Introduction", "4"],
        ["2.", "Crude Oil Markets", "4"],
        ["  2.1", "Definition and Market Size", "4"],
        ["  2.2", "Major Benchmarks", "5"],
        ["  2.3", "Market Structure and Pricing Factors", "5"],
        ["3.", "Foreign Exchange Markets", "6"],
        ["  3.1", "Overview and Structure", "6"],
        ["  3.2", "Foreign Exchange Instruments", "6"],
        ["4.", "The Crude Oil–Forex Nexus", "7"],
        ["  4.1", "The Petrodollar System", "7"],
        ["  4.2", "Mechanism: How Oil Prices Move Exchange Rates", "8"],
        ["  4.3", "India's Crude Oil Landscape and INR Vulnerability", "8"],
        ["  4.4", "Historical Data: Brent Crude vs. INR/USD", "9"],
        ["  4.5", "Exchange Rate Mechanisms and INR Determinants", "9"],
        ["5.", "Hedging Against Crude Oil–Forex Risk", "10"],
        ["6.", "Policy Implications", "11"],
        ["7.", "Conclusion", "11"],
        ["8.", "References", "12"],
      ].map(([num, title, pg]) => new Paragraph({
        spacing: { before: 60, after: 60 },
        tabStops: [{ type: TabStopType.RIGHT, position: CONTENT_WIDTH }],
        children: [
          new TextRun({ text: `${num}  ${title}`, size: 22, font: "Calibri" }),
          new TextRun({ text: `\t${pg}`, size: 22, font: "Calibri" })
        ]
      })),

      pageBreak(),

      // ─── 1. INTRODUCTION ──────────────────────────────────────────────────────
      h1("1. Introduction"),
      para("Crude oil and foreign exchange markets are two of the most powerful forces shaping the global economy. Individually, each commands enormous scale — the forex market trades approximately USD 7.5 trillion per day, while the global crude oil market is valued at around USD 3 trillion per year. Yet, their true significance lies in how deeply and structurally they are interconnected."),
      para("The connection between oil and currency markets is not incidental; it is architectural. Since the Petrodollar Agreement of 1973, all crude oil in the world has been priced and traded in US Dollars. This single fact means that any country importing oil must first acquire US Dollars, creating a permanent, structural demand for the greenback. When oil prices rise, oil-importing nations must spend more dollars — weakening their domestic currencies — while oil exporters accumulate dollar surpluses."),
      para("For India, this relationship is particularly consequential. India imports over 85% of its crude oil and is the world's third-largest consumer. Its annual crude oil bill exceeds USD 140 billion. The Indian Rupee (INR) is therefore highly sensitive to international oil price movements. A USD 10 per barrel rise in crude prices alone widens India's Current Account Deficit by approximately USD 15 billion, placing downward pressure on the Rupee."),
      para("This project systematically examines the oil-forex nexus through four lenses: (1) the structure and pricing dynamics of crude oil markets; (2) the functioning of global foreign exchange markets and the instruments used; (3) the direct channels through which oil prices influence exchange rates, with specific attention to India's INR; and (4) the hedging strategies available to manage this risk."),

      // ─── 2. CRUDE OIL MARKETS ─────────────────────────────────────────────────
      h1("2. Crude Oil Markets"),
      h2("2.1 Definition and Market Size"),
      para("Crude oil is a naturally occurring, unrefined petroleum product composed of hydrocarbon deposits found underground. It is the world's most actively traded commodity and is central to global economic activity — influencing energy prices, exchange rates, industrial output, and national GDPs. Upon refining, crude oil yields petrol, diesel, jet fuel, LPG, and a wide range of petrochemical products."),

      makeTable(
        ["Parameter", "Details"],
        [
          ["Annual Global Market Value", "~USD 3 Trillion"],
          ["Daily Global Production", "~100 million barrels per day"],
          ["Unit of Trade", "1 Barrel (bbl) = 159 litres"],
          ["Pricing Currency", "US Dollar (USD) — Petrodollar system"],
          ["Top Producers", "USA, Saudi Arabia, Russia"],
        ],
        [4200, 5160]
      ),
      blankLine(),

      h2("2.2 Major Crude Oil Benchmarks"),
      para("Crude oil is not a homogeneous commodity — it varies in density (measured in API gravity) and sulphur content. Different grades are priced off different benchmark crudes:"),

      makeTable(
        ["Benchmark", "Origin", "API Gravity", "Significance"],
        [
          ["Brent Crude", "North Sea (UK/Norway)", "~38° — Sweet", "Global benchmark; ~67% of world oil priced off Brent"],
          ["WTI (West Texas Intermediate)", "Midland, Texas, USA", "~39.6° — Sweet", "US benchmark; traded on NYMEX; typically USD 2–4 cheaper than Brent"],
          ["Dubai / Oman Crude", "Middle East", "~31° — Sour", "Asian benchmark; India's imports primarily priced off this"],
          ["OPEC Basket", "Multi-country blend", "Varies", "Average of OPEC member prices; reference for production decisions"],
        ],
        [2200, 2000, 1800, 3360]
      ),
      blankLine(),

      h2("2.3 Market Structure and Pricing Factors"),
      para("The crude oil market involves a wide ecosystem of participants, each influencing prices in different ways:"),
      bullet("OPEC+ — A producer cartel controlling approximately 40% of global oil supply. Production cut decisions by OPEC+ have an immediate and significant effect on global prices."),
      bullet("National Oil Companies (NOCs) — e.g., Saudi Aramco, ADNOC, ONGC — state-owned entities that manage the bulk of oil production in their respective countries."),
      bullet("International Oil Companies (IOCs) — e.g., Shell, BP, ExxonMobil — major private sector players involved in exploration, production, and refining globally."),
      bullet("Commodity Exchanges — NYMEX (New York), ICE (London), MCX (India) — where crude oil futures and options are listed and traded."),
      bullet("Hedge Funds and Speculators — Participate via futures markets, contributing to short-term price volatility."),
      blankLine(),
      para("Key pricing factors include supply-demand dynamics, geopolitical developments (sanctions on Iran or Russia), the strength of the US Dollar (inverse relationship — a stronger dollar makes oil cheaper for dollar-holders, reducing demand), speculative activity in derivatives markets, and weekly US Energy Information Administration (EIA) inventory reports."),

      pageBreak(),

      // ─── 3. FOREIGN EXCHANGE MARKETS ──────────────────────────────────────────
      h1("3. Foreign Exchange Markets"),
      h2("3.1 Overview and Structure"),
      para("The Foreign Exchange (Forex or FX) market is the world's largest and most liquid financial market, with a daily trading volume of approximately USD 7.5 trillion — dwarfing all global stock markets combined. It operates 24 hours a day, five days a week, in a continuous cycle across time zones: Sydney → Tokyo → London → New York."),
      para("Unlike stock exchanges, the forex market is entirely decentralised (Over-the-Counter / OTC). There is no single physical exchange. Trades occur electronically between banks, financial institutions, corporations, and retail brokers through platforms such as EBS (Electronic Broking Services) and Reuters Matching."),
      para("In India, the foreign exchange market is regulated by the Reserve Bank of India (RBI) under the Foreign Exchange Management Act (FEMA), 1999. Commercial banks, authorised by the RBI as Authorised Dealers, serve as the key intermediaries for foreign exchange transactions."),

      h2("3.2 Foreign Exchange Instruments"),
      para("Participants in the forex market use a range of instruments to exchange currencies, manage risk, and speculate:"),

      makeTable(
        ["Instrument", "Settlement", "Primary Users", "Key Feature"],
        [
          ["Spot Transactions", "T+2 days", "Importers, Exporters", "Exchange at current market rate for near-immediate settlement"],
          ["Forward Contracts", "Weeks to months ahead", "Corporations, Oil PSUs", "Rate locked today for future settlement — eliminates FX uncertainty"],
          ["Currency Futures", "Last Thursday of expiry month", "Speculators and Hedgers", "Standardised forward contracts traded on NSE/BSE"],
          ["Currency Options", "Exercise date", "Exporters, Treasuries", "Right (not obligation) to exchange at predetermined rate; flexible"],
          ["Currency Swaps", "As per agreement", "MNCs and Banks", "Simultaneous buy-and-sell of same currency for different maturities"],
        ],
        [2200, 1600, 2000, 3560]
      ),
      blankLine(),

      pageBreak(),

      // ─── 4. CRUDE OIL–FOREX NEXUS ─────────────────────────────────────────────
      h1("4. The Crude Oil–Forex Nexus"),
      h2("4.1 The Petrodollar System"),
      para("The Petrodollar system is the foundational mechanism linking oil and currency markets. Established in 1973 following an agreement between the United States and OPEC nations — primarily Saudi Arabia — the system mandates that all crude oil globally be priced and transacted exclusively in US Dollars."),
      para("The structural loop of the Petrodollar system works as follows:"),
      bullet("An oil-importing country needs crude oil and must acquire US Dollars to pay for it."),
      bullet("The country sells its domestic currency in the foreign exchange market to purchase USD."),
      bullet("The USD payment is received by the oil-exporting nation (e.g., Saudi Aramco)."),
      bullet("The exporter reinvests the dollar surplus — typically in US Treasury bonds, dollar-denominated assets, and sovereign wealth funds."),
      blankLine(),
      para("This creates a permanent, structural demand for the US Dollar globally. Nations that export oil accumulate vast dollar reserves, and those that import oil face a chronic requirement to earn or purchase dollars. The Petrodollar system thereby makes the USD the world's primary reserve currency and gives the United States a unique structural advantage in international finance."),
      para("For India, the Petrodollar system translates into a chronic structural demand for USD to fund oil imports. The RBI holds foreign exchange reserves specifically to buffer against sudden shocks in oil payment obligations. India has also explored alternative arrangements — such as INR-Ruble trade with Russia — as a strategic effort to reduce reliance on the petrodollar, particularly after Russia's discounted crude became available post-2022 sanctions."),

      h2("4.2 Mechanism: How Oil Prices Move Exchange Rates"),
      para("The transmission of oil price changes to exchange rate movements operates through distinct channels, varying depending on whether a country is an oil importer or exporter:"),

      makeTable(
        ["Direction", "Impact on Oil-Importing Nations (e.g., India, Japan, EU)", "Impact on Oil-Exporting Nations (e.g., Saudi Arabia, Russia, UAE)"],
        [
          ["Rising Oil Prices", "Higher import bill → more USD demand → domestic currency weakens; inflation rises; Current Account Deficit (CAD) widens; central bank may intervene", "Greater petrodollar inflows → currency strengthens; sovereign wealth funds accumulate; government budgets see surplus"],
          ["Falling Oil Prices", "Reduced import burden → CAD narrows; lower imported inflation; currency may stabilise or appreciate", "Revenue shortfall → currency weakens (Ruble, Riyal); sovereign wealth funds drawn down; government budget deficits widen"],
        ],
        [1560, 3900, 3900]
      ),
      blankLine(),

      h2("4.3 India's Crude Oil Landscape and INR Vulnerability"),
      para("India's position in the global oil market makes it uniquely exposed to oil-price-driven exchange rate volatility. Several key statistics illustrate the scale of this exposure:"),

      makeTable(
        ["Parameter", "Figure"],
        [
          ["Import Dependence", "85%+ of crude oil requirements met via imports"],
          ["Global Ranking", "3rd largest consumer of crude oil in the world"],
          ["Annual Crude Oil Bill", "USD 140 billion+ (2024)"],
          ["Top Supplier (2024)", "Russia (discounted crude post-sanctions)"],
          ["CAD Sensitivity", "USD 10/bbl price rise → CAD widens by ~USD 15 billion"],
        ],
        [4200, 5160]
      ),
      blankLine(),
      para("Key Indian institutions involved in the crude oil sector include:"),
      bullet("ONGC (Oil and Natural Gas Corporation) — Responsible for exploration and domestic production of crude oil."),
      bullet("Indian Oil Corporation (IOC) — India's largest refiner; manages pipeline networks and petroleum marketing."),
      bullet("Bharat Petroleum Corporation Ltd. (BPCL) — Engaged in refining and retail distribution of petroleum products."),
      bullet("PPAC (Petroleum Planning and Analysis Cell) — Monitors oil prices, imports, and government subsidy policy."),

      h2("4.4 Historical Data: Brent Crude vs. INR/USD (2018–2024)"),
      para("The following data illustrates the strong historical correlation between international crude oil prices and the INR/USD exchange rate:"),

      makeTable(
        ["Year", "Brent Crude (USD/bbl)", "INR/USD (Annual Avg.)", "Key Context"],
        [
          ["2018", "71.0", "68.4", "Trade tensions; oil demand healthy"],
          ["2019", "64.0", "70.4", "Moderate oil prices; INR mild depreciation"],
          ["2020", "42.0", "74.1", "COVID-19 demand crash; rupee fell due to capital outflows"],
          ["2021", "70.0", "73.9", "Recovery; oil and INR relatively stable"],
          ["2022", "101.0", "78.6", "Russia–Ukraine war; oil spike → sharp INR depreciation"],
          ["2023", "83.0", "82.6", "Elevated oil; persistent INR weakness"],
          ["2024", "80.0", "83.5", "Stable oil; continued INR depreciation trend"],
        ],
        [900, 2160, 2160, 4140]
      ),
      blankLine(),
      para("The data confirms the theory: as Brent crude prices surged in 2022 to USD 101/bbl, the INR/USD rate climbed to 78.6 — one of its sharpest annual depreciations. The structural trend from 2018 to 2024 shows the rupee declining from 68.4 to 83.5 per dollar, broadly tracking the upward trajectory in oil prices (excluding the temporary COVID dip)."),

      h2("4.5 Exchange Rate Mechanisms and INR Determinants"),
      para("Countries manage their exchange rates using one of three broad regimes:"),
      bullet("Fixed Rate System: The currency is pegged to another (usually the USD) at a set value. The central bank buys or sells reserves to maintain the peg. Example: Saudi Riyal pegged to USD at 3.75."),
      bullet("Floating Rate System: Market forces of supply and demand freely determine the exchange rate. Example: US Dollar, Euro, British Pound."),
      bullet("Managed Float (Dirty Float): The currency floats in line with market forces, but the central bank intervenes to prevent excessive volatility. This is India's system — the INR floats, but the RBI steps in when movements are sharp or disorderly."),
      blankLine(),
      para("Beyond oil prices, the INR exchange rate is influenced by several other factors:"),

      makeTable(
        ["Determinant", "Effect on INR"],
        [
          ["Rising Crude Oil Prices", "INR weakens — more USD required for oil import bill"],
          ["Higher FII / FDI Inflows", "INR strengthens — increased dollar inflows"],
          ["US Federal Reserve Rate Hikes", "INR weakens — capital flows out of India to USD-denominated assets"],
          ["India's Export Growth", "INR strengthens — more USD earned by exporters"],
          ["RBI Forex Reserve Interventions", "Stabilises INR — RBI sells USD to cap depreciation"],
        ],
        [4200, 5160]
      ),
      blankLine(),

      pageBreak(),

      // ─── 5. HEDGING ───────────────────────────────────────────────────────────
      h1("5. Hedging Against Crude Oil–Forex Risk"),
      para("Given the significant exposure of Indian corporates, airlines, and public sector oil companies to the dual risks of oil price volatility and currency depreciation, effective hedging is essential. Four primary strategies are widely employed:"),

      h3("5.1 Currency Forward Contracts"),
      para("An oil PSU or airline that knows it will need to buy USD to pay for an oil import in 3 months can enter a forward contract with a bank today — locking in the USD/INR rate. This eliminates the uncertainty of what the exchange rate will be at the time of payment."),
      bullet("Best suited for: Indian Oil Corporation, BPCL, airlines like IndiGo and Air India."),
      bullet("Risk: If the INR strengthens, the company misses out on the gains from a more favourable spot rate."),

      h3("5.2 Currency Options (Put / Call)"),
      para("A currency option gives the buyer the right — but not the obligation — to exchange currency at a predetermined rate. An importer expecting to pay USD would buy a USD call option: if the INR weakens beyond the strike rate, they exercise the option and buy USD at the favourable agreed rate; if the INR strengthens, they let the option lapse and buy at the cheaper market rate."),
      bullet("Best suited for: Exporters, corporate treasuries managing FX exposure."),
      bullet("Risk: The option premium adds to cost; pricing can be complex for large or long-dated positions."),

      h3("5.3 Crude Oil Futures (MCX / NYMEX)"),
      para("Refineries and aviation companies that need to purchase crude oil in the future can buy crude futures contracts to lock in the procurement price today. If oil prices spike by the delivery date, the gains on the futures contract offset the higher cost of physical crude."),
      bullet("Best suited for: Refineries (IOC, Reliance), aviation companies (IndiGo, Air India)."),
      bullet("Risk: If oil prices fall, the company faces losses on the futures position and may also receive margin calls."),

      h3("5.4 Natural Hedging"),
      para("Companies with both USD revenues and USD costs can offset their forex exposure without using any financial derivative. For instance, an Indian IT company that earns in USD and also has USD-denominated procurement costs has a natural hedge — the two exposures offset each other. In the oil context, a company that both exports refined petroleum products and imports crude has a partial natural hedge."),
      bullet("Best suited for: Multinational corporations with diversified geographies and balanced currency exposures."),
      bullet("Risk: Requires a carefully balanced revenue and cost structure across geographies; difficult to achieve perfectly."),

      pageBreak(),

      // ─── 6. POLICY IMPLICATIONS ───────────────────────────────────────────────
      h1("6. Policy Implications"),
      para("The structural vulnerability of the Indian economy to the oil-forex nexus calls for coordinated policy responses at the macroeconomic and institutional level:"),
      bullet("Managed Float and RBI Intervention: India's managed float regime allows the RBI to intervene in the forex market — selling USD from reserves — when the INR depreciates excessively due to oil shocks. India's forex reserves (over USD 600 billion as of 2024) provide a substantial buffer against short-term oil payment volatility."),
      bullet("Diversification of Oil Suppliers: India's shift toward Russian crude post-2022 — which is available at a significant discount — is a strategic hedge against benchmark oil price volatility. Supplier diversification reduces dependence on any single source and opens pricing alternatives."),
      bullet("INR-Ruble and Bilateral Currency Trade: India's experiments with settling oil trade in INR and Roubles represent attempts to reduce petrodollar dependency. While not yet at scale, such mechanisms could meaningfully reduce USD demand over the long run."),
      bullet("Domestic Energy Transition: India's investments in renewable energy (solar, wind) and electric vehicles are a long-term structural hedge — reducing oil import volumes and, consequently, structural USD demand."),
      bullet("Strategic Petroleum Reserve (SPR): India's SPR capacity provides a short-term buffer against oil supply shocks, reducing the need for emergency market purchases at peak prices."),

      pageBreak(),

      // ─── 7. CONCLUSION ────────────────────────────────────────────────────────
      h1("7. Conclusion"),
      para("The relationship between crude oil prices and foreign exchange markets is one of the most consequential linkages in the global financial system. Rooted in the Petrodollar framework established in 1973, this connection creates a structural dependency between the world's largest energy commodity and the world's dominant reserve currency — the US Dollar."),
      para("For India, this relationship is not abstract — it is felt in everyday prices, monetary policy decisions, and the Reserve Bank's management of forex reserves. With over 85% of crude oil imported, India's Current Account Deficit, inflation trajectory, and Rupee valuation are all materially influenced by Brent crude prices. The historical evidence from 2018-2024 demonstrates the structural linkage between oil shocks and INR depreciation."),

      para("Managing this risk requires active coordination among energy policy, macroeconomic policy, and financial markets. Continued diversification of crude suppliers, hedging with forwards and options, accumulated forex reserves, and accelerating the transition to renewable energy can together reduce India's currency vulnerability to crude price shocks."),

      h1("8. References"),
      bullet("Reserve Bank of India (RBI), 'India\'s Foreign Exchange Reserves' publications, 2024."),
      bullet("Ministry of Petroleum & Natural Gas, 'India Petroleum & Natural Gas Statistics' 2023-24."),
      bullet("International Energy Agency (IEA), 'Oil Market Report' 2024."),
      bullet("BP Statistical Review of World Energy, 2024."),
      bullet("IMF World Economic Outlook, October 2024."),

    ]
  }]
});

Packer.toBuffer(doc).then((buffer) => {
  fs.writeFileSync("Impact_of_Crude_Oil_on_Forex_Markets.docx", buffer);
  console.log("Document generated: Impact_of_Crude_Oil_on_Forex_Markets.docx");
}).catch((err) => {
  console.error("Error generating document:", err);
});