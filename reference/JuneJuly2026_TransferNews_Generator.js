// ORIGINAL STYLING REFERENCE — June/July 2026 Transfer News (docx.js)
// This file is the design source of truth for the printed Transfer News.
// The Python generator (generate_transfer_document.py) mirrors this palette,
// typography and layout. If the design changes, update both.
//
// Requires: npm install docx
// Run:      node reference/JuneJuly2026_TransferNews_Generator.js

const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, BorderStyle, WidthType, ShadingType,
  VerticalAlign, TabStopType, TabStopPosition, PageBreak, SectionType
} = require('docx');
const fs = require('fs');

// --------------------------------------
//  PALETTE
// --------------------------------------
const G      = "1A5C38"; // banner green
const GMID   = "236B44";
const GLIT   = "E4F2EB";
const GOLD   = "C9A227";
const GOLDT  = "7A6010";
const GOLDLIT= "FBF5DF";
const RD     = "A93226"; // new missionary red
const RDLIT  = "FDECEA";
const OR     = "B7560E"; // trainer orange
const ORLIT  = "FEF0E3";
const BL     = "1A4D8F"; // STL blue
const BLIT   = "EAF0FB";
const TE     = "0D6B55"; // DL/DT teal
const TELIT  = "E2F4EF";
const PU     = "6C3483"; // AP purple
const PULIT  = "F5EEF8";
const W      = "FFFFFF";
const K      = "111111";
const K2     = "3D3D3D";
const K3     = "888888";
const GRY    = "F6F6F6";
const GRY2   = "E8E8E8";
const STRIPE = "F0F7F3"; // zebra row tint

// --------------------------------------
//  BORDER / CELL / TEXT HELPERS
// --------------------------------------
const edge  = (c, s=4) => ({ style: BorderStyle.SINGLE, size: s, color: c });
const thick = (c, s=16) => ({ style: BorderStyle.THICK, size: s, color: c });
const none  = () => ({ style: BorderStyle.NIL, size: 0, color: W });
const box   = (c, s=4) => ({ top:edge(c,s), bottom:edge(c,s), left:edge(c,s), right:edge(c,s) });
const noBox = () => ({ top:none(), bottom:none(), left:none(), right:none() });

const mk = (children, w, o={}) => new TableCell({
  width: { size: w, type: WidthType.DXA },
  borders: o.borders || box("DEDEDE", 3),
  shading: { fill: o.bg || W, type: ShadingType.CLEAR },
  margins: o.m || { top:80, bottom:80, left:120, right:120 },
  verticalAlign: o.va || VerticalAlign.CENTER,
  columnSpan: o.span,
  children,
});

const t = (text, o={}) => new TextRun({
  text, font: "Calibri",
  size:   o.sz  || 17,
  bold:   o.b,
  italics:o.i,
  color:  o.c   || K,
  allCaps:o.cap,
  characterSpacing: o.cs,
});

const p = (runs, o={}) => new Paragraph({
  alignment:  o.align || AlignmentType.LEFT,
  spacing:    o.sp    || { before:0, after:0 },
  shading:    o.bg    ? { fill:o.bg, type:ShadingType.CLEAR } : undefined,
  border:     o.bdr,
  tabStops:   o.tabs,
  children:   Array.isArray(runs) ? runs : [runs],
});

const sp = (b=80, a=80) => p(t(""), { sp:{before:b, after:a} });

// --------------------------------------
//  ROLE BADGE CONFIG
// --------------------------------------
function roleStyle(role) {
  const base = role.replace(/[0-9]/g, "").trim();
  switch(base) {
    case "AP":  return { bg: PULIT, fg: PU,    label: role };
    case "ZL":  return { bg: GLIT,  fg: G,     label: role };
    case "STL": return { bg: BLIT,  fg: BL,    label: role };
    case "DL":  return { bg: TELIT, fg: TE,    label: role };
    case "DT":  return { bg: TELIT, fg: TE,    label: role };
    case "TR":  return { bg: ORLIT, fg: OR,    label: role };
    case "SA":  return { bg: GLIT,  fg: GMID,  label: role };
    default:    return { bg: GRY2,  fg: K3,    label: role };
  }
}

// --------------------------------------
//  TRANSFER DATA — UPDATE EACH CYCLE
//  (data omitted in this reference copy; see the sample PDF or use the
//   Python generator, which fills this automatically from the two inputs)
//  Format per row:
//  ["Name", "Role", "Zone", "Area", "Companion or NEW MISSIONARY"]
// --------------------------------------
const CYCLE = "June / July 2026";
const zones = []; // <- the Python generator derives this automatically

// --------------------------------------
//  STATS
// --------------------------------------
const totalM   = zones.reduce((s,z)=>s+z.rows.length,0);
const totalNew = zones.reduce((s,z)=>s+z.rows.filter(r=>r[4]==="NEW MISSIONARY").length,0);
const totalL   = zones.reduce((s,z)=>s+z.rows.filter(r=>/^(AP|ZL|STL|DL|DT)/.test(r[1])).length,0);

// --------------------------------------
//  TABLE — per zone
//  Columns (DXA): #, MISSIONARY, ROLE, ZONE, AREA, COMPANION(S)
// --------------------------------------
const CW = [440, 2250, 1100, 1800, 3500, 5110];
const TW = CW.reduce((a,b)=>a+b,0);

function zoneTable(zone) {
  const hdrs = ["#","MISSIONARY","ROLE","ZONE","AREA","COMPANION(S)"];
  const headerRow = new TableRow({
    tableHeader: true,
    children: hdrs.map((l,i) => mk(
      [p(t(l,{sz:15,b:true,c:W,cap:true}),{align:AlignmentType.CENTER})],
      CW[i],
      { bg:G, borders:box(G,6), m:{top:95,bottom:95,left:80,right:80} }
    ))
  });

  const dataRows = zone.rows.map(([name,role,zn,area,comp], i) => {
    const stripe = i%2!==0;
    const bg = stripe ? STRIPE : W;
    const rs = roleStyle(role);
    const isNew = comp === "NEW MISSIONARY";
    return new TableRow({ children:[
      mk([p(t(String(i+1),{sz:14,c:K3}),{align:AlignmentType.CENTER})], CW[0],
         {bg, borders:box("E0E0E0",3), m:{top:80,bottom:80,left:50,right:50}}),
      mk([p(t(name,{sz:16,b:true,c:K}))], CW[1],
         {bg, borders:box("E0E0E0",3), m:{top:80,bottom:80,left:110,right:110}}),
      mk([p(t(rs.label,{sz:14,b:true,c:rs.fg}),{align:AlignmentType.CENTER})], CW[2],
         {bg:rs.bg, borders:box("E0E0E0",3), m:{top:80,bottom:80,left:60,right:60}}),
      mk([p(t(zn,{sz:15,b:true,c:G}),{align:AlignmentType.CENTER})], CW[3],
         {bg, borders:box("E0E0E0",3), m:{top:80,bottom:80,left:80,right:80}}),
      mk([p(t(area,{sz:15,c:K2}))], CW[4],
         {bg, borders:box("E0E0E0",3), m:{top:80,bottom:80,left:110,right:110}}),
      mk([p(t(isNew?"⊕  NEW MISSIONARY":comp,
             {sz:15,c:isNew?RD:K2,b:isNew}))], CW[5],
         {bg:isNew?RDLIT:bg,
          borders: isNew
            ? {top:edge("E0E0E0",3),bottom:edge("E0E0E0",3),right:edge("E0E0E0",3),left:thick(RD,10)}
            : box("E0E0E0",3),
          m:{top:80,bottom:80,left:110,right:110}}),
    ]});
  });
  return new Table({ width:{size:TW,type:WidthType.DXA}, columnWidths:CW, rows:[headerRow,...dataRows] });
}

// -- SECTION HEADER BLOCK --
function zoneDivider(num, zone) {
  return [
    new Table({
      width:{size:TW,type:WidthType.DXA}, columnWidths:[700,13500],
      rows:[new TableRow({children:[
        mk([p(t(num.toString().padStart(2,"0"),{sz:28,b:true,c:W,cap:true}),{align:AlignmentType.CENTER})],
          700, { bg:G, borders:noBox(), m:{top:100,bottom:100,left:60,right:60} }),
        mk([p([
          t(zone.name,{sz:22,b:true,c:G,cap:true,cs:20}),
          t(`   ·   ${zone.rows.length} missionaries`,{sz:14,c:K3,i:true}),
        ])], 13500, { bg:W,
          borders:{top:none(),bottom:edge(GOLD,10),left:none(),right:none()},
          m:{top:90,bottom:90,left:160,right:160} }),
      ]})]
    }),
    sp(100,80),
  ];
}

// --------------------------------------
//  COVER PAGE
// --------------------------------------
function statCard(w, count, label, accent, bg) {
  return mk([
    p(t(String(count),{sz:52,b:true,c:accent}),{align:AlignmentType.CENTER,sp:{before:60,after:20}}),
    p(t(label,{sz:13,b:true,c:K2,cap:true}),{align:AlignmentType.CENTER,sp:{before:0,after:60}}),
  ], w, {
    borders:{top:thick(accent,20),bottom:edge("DDDDDD",4),left:edge("DDDDDD",4),right:edge("DDDDDD",4)},
    bg, m:{top:0,bottom:0,left:100,right:100},
  });
}

function buildZoneIndex() {
  const rows = [];
  for (let i=0;i<zones.length;i+=3) {
    rows.push(new TableRow({ children: [0,1,2].map(j => {
      const z = zones[i+j];
      const idx = i+j+1;
      return z
        ? mk([p([
            t(`${idx.toString().padStart(2,"0")}`,{sz:14,b:true,c:GOLDT}),
            t(`  ${z.name}`,{sz:14,b:true,c:G}),
            t(`  (${z.rows.length})`,{sz:13,c:K3,i:true}),
          ])], 4666, {bg:GRY,m:{top:70,bottom:70,left:130,right:130}})
        : mk([p(t(""))], 4666, {bg:GRY});
    })}));
  }
  return new Table({ width:{size:14000,type:WidthType.DXA}, columnWidths:[4666,4667,4667], rows:[
    new TableRow({children:[mk(
      [p(t("ZONE INDEX",{sz:15,b:true,c:W,cap:true}),{align:AlignmentType.CENTER})],
      14000,{span:3,bg:G,m:{top:80,bottom:80,left:100,right:100}})
    ]}),
    ...rows,
  ]});
}

const coverChildren = [
  p(t("NIGERIA UYO MISSION",{sz:52,b:true,c:GOLD,cap:true,cs:60}),
    {align:AlignmentType.CENTER, bg:G, sp:{before:500,after:40}}),
  p(t("THE CHURCH OF JESUS CHRIST OF LATTER-DAY SAINTS",{sz:16,c:W,cap:true,cs:20}),
    {align:AlignmentType.CENTER, bg:G, sp:{before:0,after:60},
     bdr:{bottom:edge(GOLD,18)}}),
  p(t(`${CYCLE.toUpperCase()} TRANSFER NEWS`,{sz:58,b:true,c:G,cap:true,cs:20}),
    {align:AlignmentType.CENTER, sp:{before:320,after:50}}),
  p(t("Complete Zone-by-Zone Transfer Assignments  ·  Mission Leadership Council",{sz:17,c:K3,i:true}),
    {align:AlignmentType.CENTER, sp:{before:0,after:260}}),
  new Table({ width:{size:13200,type:WidthType.DXA}, columnWidths:[3300,3300,3300,3300],
    rows:[new TableRow({children:[
      statCard(3300, zones.length,  "Zones",             G,     GLIT),
      statCard(3300, totalM,        "Missionaries",      BL,    BLIT),
      statCard(3300, totalL,        "Leadership Roles",  GOLDT, GOLDLIT),
      statCard(3300, totalNew,      "New Missionaries",  RD,    RDLIT),
    ]})],
  }),
  sp(200,140),
  buildZoneIndex(),
  sp(200,100),
  new Table({ width:{size:13200,type:WidthType.DXA}, columnWidths:[4400,4400,4400],
    rows:[new TableRow({children:[
      mk([
        p(t("PREPARED BY",{sz:12,cap:true,c:K3,b:true})),
        p(t("Elder ThankGod Andrew",{sz:17,b:true,c:G}),{sp:{before:30,after:6}}),
        p(t("Assistant to the President",{sz:13,c:K3,i:true})),
      ],4400,{bg:GRY,m:{top:130,bottom:130,left:160,right:160}}),
      mk([
        p(t("APPROVED BY",{sz:12,cap:true,c:K3,b:true})),
        p(t("President Richard Paapa Dadzie",{sz:17,b:true,c:G}),{sp:{before:30,after:6}}),
        p(t("Nigeria Uyo Mission President",{sz:13,c:K3,i:true})),
      ],4400,{bg:GRY,m:{top:130,bottom:130,left:160,right:160}}),
      mk([
        p(t("EFFECTIVE",{sz:12,cap:true,c:K3,b:true})),
        p(t(`${CYCLE} Transfer`,{sz:17,b:true,c:G}),{sp:{before:30,after:6}}),
        p(t("Document Date: _______________",{sz:13,c:K3,i:true})),
      ],4400,{bg:GRY,m:{top:130,bottom:130,left:160,right:160}}),
    ]})]
  }),
  p(t("STRICTLY CONFIDENTIAL — For Mission Presidency Use Only",{sz:13,i:true,c:K3}),
    {align:AlignmentType.CENTER, sp:{before:220,after:180}}),
  p(t(""),{bg:GOLD,sp:{before:0,after:0}}),
  p(t(""),{bg:G,sp:{before:0,after:120}}),
];

// --------------------------------------
//  HEADER / FOOTER
// --------------------------------------
const mkHeader = () => new Header({ children:[
  p([
    t("NIGERIA UYO MISSION",{sz:17,b:true,c:W,cap:true}),
    t("\t"),
    t(`${CYCLE.toUpperCase()}  ·  TRANSFER NEWS`,{sz:17,b:true,c:GOLD,cap:true}),
    t("\t"),
    t("All Zones",{sz:13,c:W,i:true}),
  ],{
    bg:G, sp:{before:60,after:60},
    bdr:{bottom:edge(GOLD,10)},
    tabs:[
      {type:TabStopType.CENTER, position: TabStopPosition.MAX/2},
      {type:TabStopType.RIGHT,  position: TabStopPosition.MAX},
    ]
  }),
]});

const mkFooter = () => new Footer({ children:[
  p([
    t("STRICTLY CONFIDENTIAL — Mission Presidency Use Only",{sz:12,i:true,c:K3}),
    t("\t"),
    t("President Richard Paapa Dadzie  ·  Nigeria Uyo Mission",{sz:12,b:true,c:G}),
    t("\t"),
    t("Page ",{sz:12,c:K3}),
  ],{
    sp:{before:60,after:0},
    bdr:{top:edge(G,8)},
    tabs:[
      {type:TabStopType.CENTER, position: TabStopPosition.MAX/2},
      {type:TabStopType.RIGHT,  position: TabStopPosition.MAX-400},
    ]
  }),
]});

// --------------------------------------
//  ZONE BLOCKS + LEGEND
// --------------------------------------
const zoneBlocks = [];
zones.forEach((zone, idx) => {
  if (idx > 0) zoneBlocks.push(new Paragraph({ children:[ new PageBreak() ] }));
  zoneBlocks.push(...zoneDivider(idx+1, zone));
  zoneBlocks.push(zoneTable(zone));
});

zoneBlocks.push(sp(200, 80));
zoneBlocks.push(
  new Table({ width:{size:TW,type:WidthType.DXA}, columnWidths:[2366,2367,2367,2367,2367,2366],
    rows:[
      new TableRow({children:[mk([p(t("ASSIGNMENT KEY",{sz:14,b:true,c:W,cap:true}),
        {align:AlignmentType.CENTER})],TW,{span:6,bg:G,m:{top:80,bottom:80,left:100,right:100}})]}),
      new TableRow({children:[
        mk([p([t("AP",{sz:14,b:true,c:PU}),t("  Assistant to President",{sz:14,c:K2})])],2366,{bg:PULIT}),
        mk([p([t("ZL",{sz:14,b:true,c:G}),  t("  Zone Leader",{sz:14,c:K2})])],2367,{bg:GLIT}),
        mk([p([t("STL",{sz:14,b:true,c:BL}),t("  Sister Training Leader",{sz:14,c:K2})])],2367,{bg:BLIT}),
        mk([p([t("DL",{sz:14,b:true,c:TE}), t("  District Leader",{sz:14,c:K2})])],2367,{bg:TELIT}),
        mk([p([t("DT",{sz:14,b:true,c:TE}), t("  District Trainer",{sz:14,c:K2})])],2367,{bg:TELIT}),
        mk([p([t("TR",{sz:14,b:true,c:OR}), t("  Trainer",{sz:14,c:K2})])],2366,{bg:ORLIT}),
      ]}),
      new TableRow({children:[
        mk([p([t("SA",{sz:14,b:true,c:GMID}),t("  Special Assignment",{sz:14,c:K2})])],2366,{bg:GLIT}),
        mk([p([t("JC",{sz:14,b:true,c:K3}), t("  Junior Companion",{sz:14,c:K2})])],2367,{bg:GRY}),
        mk([p([t("SC",{sz:14,b:true,c:K3}), t("  Senior Companion",{sz:14,c:K2})])],2367,{bg:GRY}),
        mk([p([t("⊕",{sz:14,b:true,c:RD}),  t("  New Missionary (incoming)",{sz:14,c:K2})])],2367,{bg:RDLIT}),
        mk([p([t("",{sz:14})])],2367,{bg:GRY}),
        mk([p([t("",{sz:14})])],2366,{bg:GRY}),
      ]}),
    ]
  })
);

// --------------------------------------
//  ASSEMBLE & WRITE
// --------------------------------------
const doc = new Document({
  styles:{ default:{ document:{ run:{ font:"Calibri", size:17 } } } },
  sections:[
    {
      properties:{ page:{size:{width:15840,height:12240}, margin:{top:0,right:0,bottom:0,left:0}, orientation:"landscape"}, type:SectionType.NEXT_PAGE },
      children: coverChildren,
    },
    {
      properties:{ page:{size:{width:15840,height:12240}, margin:{top:720,right:720,bottom:720,left:720}, orientation:"landscape"}, type:SectionType.NEXT_PAGE },
      headers:{ default: mkHeader() },
      footers:{ default: mkFooter() },
      children: zoneBlocks,
    }
  ]
});

Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync("Nigeria_Uyo_Mission_TransferNews.docx", buf);
  console.log(`Done — ${zones.length} zones | ${totalM} missionaries | ${totalNew} new | ${totalL} leadership`);
});
