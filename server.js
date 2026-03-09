const express = require("express");
const pptxgen = require("pptxgenjs");
const sharp = require("sharp");

const app = express();
app.use(express.json({ limit: "50mb" }));

const NAVY       = "156082";
const TEAL       = "04AF87";
const TRANSIT_BG = "FFCCCC";
const DIRECT_BG  = "CCFFCC";
const WHITE      = "FFFFFF";
const GREY       = "F2F2F2";
const BORDER     = "C8D8E8";
const DARK       = "222222";
const FONT       = "DIN Next LT Arabic";
const TABLE_X    = 0.18;
const HEADER_H   = 0.75;
const ROW_H      = 0.48;
const COL        = [0.22, 1.28, 1.28, 1.65, 1.22, 1.22, 0.98, 1.14, 1.82, 2.15];

async function mkIcon(svg) {
  const buf = await sharp(Buffer.from(svg)).resize(160, 160).png().toBuffer();
  return "image/png;base64," + buf.toString("base64");
}

async function generatePPTX(flights, legLabel, missionDate) {
  const pinPng = await mkIcon(`<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 32 32">
    <path d="M16 2C10.48 2 6 6.48 6 12c0 7.5 10 18 10 18S26 19.5 26 12c0-5.52-4.48-10-10-10z" fill="#11A77C"/>
    <circle cx="16" cy="12" r="4" fill="white"/><circle cx="16" cy="12" r="2" fill="#11A77C"/></svg>`);

  const calPng = await mkIcon(`<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 32 32">
    <rect x="2" y="5" width="28" height="24" rx="3" fill="#04AF87"/>
    <rect x="2" y="5" width="28" height="10" rx="3" fill="#038060"/>
    <rect x="2" y="12" width="28" height="3" fill="#038060"/>
    <rect x="9" y="1" width="3" height="7" rx="1.5" fill="#04AF87"/>
    <rect x="20" y="1" width="3" height="7" rx="1.5" fill="#04AF87"/>
    <rect x="6"  y="18" width="4" height="4" rx="1" fill="white"/>
    <rect x="14" y="18" width="4" height="4" rx="1" fill="white"/>
    <rect x="22" y="18" width="4" height="4" rx="1" fill="white"/>
    <rect x="6"  y="24" width="4" height="3" rx="1" fill="white"/>
    <rect x="14" y="24" width="4" height="3" rx="1" fill="white"/></svg>`);

  const starPng = await mkIcon(`<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 32 32">
    <polygon points="16,3 19.5,11.5 29,12.5 22.5,19 24.5,28.5 16,24 7.5,28.5 9.5,19 3,12.5 12.5,11.5"
             fill="#156082" stroke="#156082" stroke-width="0.5"/></svg>`);

  const bagPng = await mkIcon(`<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 32 32">
    <rect x="3" y="11" width="26" height="17" rx="3" fill="#04AF87"/>
    <rect x="10" y="5" width="12" height="7" rx="2" fill="#04AF87" stroke="#038060" stroke-width="1.5"/>
    <rect x="14" y="10" width="4" height="8" rx="1" fill="white"/>
    <rect x="3"  y="17" width="26" height="3" fill="white" opacity="0.35"/></svg>`);

  const pres = new pptxgen();
  pres.layout = "LAYOUT_WIDE";

  const dateMap = {};
  let di = 0;
  for (const f of flights) {
    if (!(f.date in dateMap)) dateMap[f.date] = di++ % 2;
  }

  const s = pres.addSlide();
  s.background = { color: WHITE };

  // HEADER
  s.addText(`Suggested Flight Schedule – ${legLabel}`, {
    x:0.3, y:0, w:6.8, h:HEADER_H,
    fontSize:17, bold:true, color:NAVY, fontFace:FONT, align:"left", valign:"middle"
  });
  s.addShape(pres.shapes.LINE, { x:7.32, y:0.14, w:0, h:HEADER_H-0.28, line:{color:"CCCCCC",width:1} });
  s.addShape(pres.shapes.LINE, { x:10.24, y:0.14, w:0, h:HEADER_H-0.28, line:{color:"CCCCCC",width:1} });

  s.addImage({ data:pinPng, x:7.48, y:0.14, w:0.34, h:0.42 });
  s.addText("Visit Location", { x:7.88, y:0.02, w:2.3, h:0.24, fontSize:11, bold:true, color:"888888", fontFace:FONT });
  s.addText(legLabel.split(" to ")[1] || legLabel, { x:7.88, y:0.30, w:2.3, h:0.26, fontSize:10.5, bold:true, color:NAVY, fontFace:FONT });

  s.addImage({ data:calPng, x:10.38, y:0.14, w:0.36, h:0.42 });
  s.addText("Mission Date", { x:10.80, y:0.02, w:2.3, h:0.24, fontSize:11, bold:true, color:"888888", fontFace:FONT });
  s.addText(missionDate, { x:10.80, y:0.30, w:2.4, h:0.26, fontSize:10, bold:true, color:NAVY, fontFace:FONT });

  const TABLE_W = COL.reduce((a,b)=>a+b,0);
  s.addShape(pres.shapes.LINE, { x:TABLE_X, y:HEADER_H, w:TABLE_W, h:0, line:{color:BORDER,width:1.2} });

  // TABLE
  const ho = (t) => ({ text:t, options:{ fill:{color:NAVY}, color:WHITE, bold:true, fontSize:12, align:"center", valign:"middle", fontFace:FONT }});
  const splitHdr = {
    text:[
      { text:"Total Duration", options:{ bold:true, color:WHITE, breakLine:true } },
      { text:"──────────────", options:{ color:"7AADBD", fontSize:5, breakLine:true } },
      { text:"Transit Time",   options:{ bold:true, color:WHITE } }
    ],
    options:{ fill:{color:NAVY}, fontSize:12, align:"center", valign:"middle", fontFace:FONT }
  };

  const rows = [[ ho("#"), ho("Day"), ho("Date"), ho("Airline"),
                  ho("Departure"), ho("Arrival"), ho("Flight"), ho("Duration"),
                  splitHdr, ho("Class") ]];
  const rowHArr = [0.72];

  for (const f of flights) {
    const bg = dateMap[f.date] === 0 ? WHITE : GREY;
    const c_ = (t, ex={}) => ({ text:t, options:{ fill:{color:bg}, color:DARK, fontSize:12, align:"center", valign:"middle", fontFace:FONT, ...ex }});
    const m_ = (t, ex={}) => ({ text:t, options:{ fill:{color:bg}, color:DARK, fontSize:12, align:"center", valign:"middle", fontFace:FONT, rowspan:2, ...ex }});

    if (f.direct) {
      const l = f.legs[0];
      rows.push([
        m_(String(f.num), { color:DARK, fontSize:13 }),
        m_(f.day), m_(f.date), m_(f.airline, { bold:true }),
        m_(`${l.from}\n${l.ft}`), m_(`${l.to}\n${l.at}`),
        m_(l.no), m_(l.dur),
        c_(l.total),
        m_(" ")
      ]);
      rows.push([
        { text:"Direct →", options:{ fill:{color:DIRECT_BG}, color:"2E7D32", bold:true, fontSize:12, align:"center", valign:"middle", fontFace:FONT }}
      ]);
    } else {
      const l1 = f.legs[0], l2 = f.legs[1];
      rows.push([
        m_(String(f.num), { color:DARK, fontSize:13 }),
        m_(f.day), m_(f.date), m_(f.airline, { bold:true }),
        c_(`${l1.from}\n${l1.ft}`), c_(`${l1.to}\n${l1.at}`),
        c_(l1.no), c_(l1.dur),
        c_(l1.total),
        m_(" ")
      ]);
      rows.push([
        c_(`${l2.from}\n${l2.ft}`), c_(`${l2.to}\n${l2.at}`),
        c_(l2.no), c_(l2.dur),
        { text:l2.transit, options:{ fill:{color:TRANSIT_BG}, color:"CC0000", fontSize:12, align:"center", valign:"middle", fontFace:FONT }}
      ]);
    }
    rowHArr.push(ROW_H, ROW_H);
  }

  s.addTable(rows, {
    x:TABLE_X, y:HEADER_H+0.07, w:TABLE_W,
    colW:COL, rowH:rowHArr,
    border:{pt:0.5, color:BORDER},
    autoPage:true, autoPageRepeatHeader:true, autoPageHeaderRows:1
  });

  // CLASS ICONS
  const clsColX = TABLE_X + COL.slice(0,9).reduce((a,b)=>a+b,0);
  const clsW    = COL[9];
  const ICON_S  = 0.26;
  const GAP     = 0.08;
  const TXT_W   = 0.95;
  const GROUP   = ICON_S + GAP + TXT_W;
  const colMid  = clsColX + clsW / 2;
  const iconX   = colMid - GROUP / 2;
  const textX   = iconX + ICON_S + GAP;
  const cellH   = 2 * ROW_H;
  let cy = HEADER_H + 0.07 + rowHArr[0];

  for (const f of flights) {
    const bg = dateMap[f.date] === 0 ? WHITE : GREY;
    const iy = cy + (cellH - ICON_S) / 2;
    if (f.cls === "First")    s.addImage({ data:starPng, x:iconX, y:iy, w:ICON_S, h:ICON_S });
    if (f.cls === "Business") s.addImage({ data:bagPng,  x:iconX, y:iy, w:ICON_S, h:ICON_S });
    s.addText(f.cls, {
      x:textX, y:cy, w:TXT_W, h:cellH,
      fontSize:12, color:DARK, fontFace:FONT,
      align:"left", valign:"middle",
      fill:{ type:"none" }, line:{ type:"none" }
    });
    cy += cellH;
  }

  // FOOTER
  s.addShape(pres.shapes.LINE, { x:0.2, y:7.35, w:12.9, h:0, line:{color:BORDER,width:1} });
  s.addText("Confidential — For Internal Use Only",  { x:0.3, y:7.38, w:6,   h:0.18, fontSize:8, color:"888888", fontFace:FONT, align:"left"  });
  s.addText("Auto-generated by Flight Brief System", { x:6.8, y:7.38, w:6.3, h:0.18, fontSize:8, color:"888888", fontFace:FONT, align:"right" });

  const buffer = await pres.write({ outputType: "base64" });
  return buffer;
}

// MAIN ENDPOINT
app.post("/generate", async (req, res) => {
  try {
    const { flights, legLabel, missionDate } = req.body;
    if (!flights || !flights.length) {
      return res.status(400).json({ error: "No flights provided" });
    }
    const base64 = await generatePPTX(flights, legLabel || "Flight Schedule", missionDate || "");
    res.json({ success: true, pptx: base64 });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message });
  }
});

app.get("/", (req, res) => res.send("PPTX Server is running ✅"));

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Server running on port ${PORT}`));
