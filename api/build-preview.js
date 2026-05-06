// api/build-preview.js
// Preview Generator – Excel Builder v2 (Schema mit Matrix/Freitext-Support)
//
// Layout-Konzept:
//   Z1-4:    Header (Studio, Kunde, Projekt) — aus Template
//   Z5:      Matrix-Mutter-Header (gemerged über alle Item-Spalten einer Matrix)
//   Z6:      Frage-Header / Item-Sub-Header
//   Z7..N:   TN-Eintragszeilen (N = brutto, dynamisch)
//   ZN+1:    "X für Y"-Label in Spalte E (lfd. Nr.)
//   ZN+2..:  Antwort-Codes (Screenout = dunkelrot, Off-Target = hellrot, Quote = grün)
//
// Tracking-Spalten (A bis QUOTE_START-1) bekommen ab Zeile N+1 keinen Inhalt mehr
// und werden visuell sauber abgeschlossen.
import ExcelJS from "exceljs";
import { LOGOS } from './logos.js';

// ---------------------------------------------------------------------------
// 1) KONFIGURATION
// ---------------------------------------------------------------------------

const SHEET_CONFIG = {
  GD:  { sheetName: 'GD',   quoteStartCol: 14 },
  IDI: { sheetName: 'IDIs', quoteStartCol: 16 },
  VGD: { sheetName: 'VGDs', quoteStartCol: 20 },
  VDI: { sheetName: 'VDIs', quoteStartCol: 22 },
};

// Header-Zellen je Setting (offline = GD/IDI, online = VGD/VDI)
const HEADER_MAP = {
  offline: {
    studioLabel: 'H1',  // "F&T Standort"
    studioWert:  'I1',  // "Bochum"
    terminLabel: 'L1', terminWert: 'L1',  // im offline kein Label, Wert direkt
    kunde:       'I2',  // "Psyma"
    zielgruppe:  'L2',
    projekt:     'I3',
    projNr:      'I4',
    incentive:   'L4',
  },
  online: {
    studioLabel: null,  // online kein Studio
    studioWert:  null,
    terminLabel: 'L1', terminWert: 'M1',
    kunde:       'J1',
    zielgruppe:  'J3',
    projekt:     'J2',
    projNr:      'J4',
    incentive:   'M4',
  }
};

const TN_START = 7;
const MATRIX_HEADER_ROW = 5;
const HEADER_ROW = 6;
const QUOTE_HINT_AFTER_TN_OFFSET = 1;  // "X für Y"-Label = TN_END + 1
const ANT_OFFSET = 2;                  // Antworten beginnen TN_END + 2

// ---------------------------------------------------------------------------
// 2) FARBEN
// ---------------------------------------------------------------------------

const COLORS = {
  DUNKELROT:   'FFC00000', // Screenout
  HELLROT:     'FFF4CCCC', // Off-Target (außerhalb Zielgruppe der Gruppe)
  GRUEN:       'FFA9D08E', // Quote-Hinweis
  HELLGRUEN:   'FFC6E0B4',
  HELLBLAU:    'FFBDD7EE',
  GELB:        'FFFFE699',
  ORANGE:      'FFFF5050',
  ROT:         'FFFF0000',
  GRAU:        'FF808080',
  WEISS:       'FFFFFFFF',
  BORDER:      'FFCCCCCC',
  MATRIX_HDR:  'FFDAE3F3', // zartes Blau für Matrix-Mutter-Header
  HEADER_GREY: 'FFF2F2F2',
  QUOTE_FONT:  'FF2E7D32', // dunkles Grün für Quote-Text
};

const THIN_BORDER = {
  top:    { style: 'thin', color: { argb: COLORS.BORDER } },
  bottom: { style: 'thin', color: { argb: COLORS.BORDER } },
  left:   { style: 'thin', color: { argb: COLORS.BORDER } },
  right:  { style: 'thin', color: { argb: COLORS.BORDER } },
};

const HEADER_BORDER = {
  top:    { style: 'thin', color: { argb: 'FF000000' } },
  bottom: { style: 'thin', color: { argb: 'FF000000' } },
  left:   { style: 'thin', color: { argb: COLORS.BORDER } },
  right:  { style: 'thin', color: { argb: COLORS.BORDER } },
};

// ---------------------------------------------------------------------------
// 3) HILFSFUNKTIONEN
// ---------------------------------------------------------------------------

function copyWorksheet(srcWs, dstWs) {
  srcWs.columns.forEach((col, i) => {
    const dstCol = dstWs.getColumn(i + 1);
    if (col.width)  dstCol.width  = col.width;
    if (col.hidden) dstCol.hidden = col.hidden;
  });

  srcWs.eachRow({ includeEmpty: true }, (srcRow, rn) => {
    const dstRow = dstWs.getRow(rn);
    if (srcRow.height) dstRow.height = srcRow.height;
    if (srcRow.hidden) dstRow.hidden = srcRow.hidden;
    srcRow.eachCell({ includeEmpty: true }, (srcCell, cn) => {
      const dstCell = dstWs.getCell(rn, cn);
      dstCell.value = srcCell.value;
      if (srcCell.font)      dstCell.font      = { ...srcCell.font };
      if (srcCell.fill)      dstCell.fill      = { ...srcCell.fill };
      if (srcCell.border)    dstCell.border    = { ...srcCell.border };
      if (srcCell.alignment) dstCell.alignment = { ...srcCell.alignment };
      if (srcCell.numFmt)    dstCell.numFmt    = srcCell.numFmt;
    });
    dstRow.commit();
  });

  if (srcWs._merges) {
    Object.keys(srcWs._merges).forEach(key => {
      try { dstWs.mergeCells(key); } catch (e) {}
    });
  }

  if (srcWs.dataValidations?.model) {
    Object.entries(srcWs.dataValidations.model).forEach(([sqref, dv]) => {
      try { dstWs.dataValidations.add(sqref, { ...dv }); } catch (e) {}
    });
  }
}

// Findet die erste "Quote"-Spalte in Zeile 6 dynamisch
function findQuoteStartCol(ws, fallback) {
  const row6 = ws.getRow(HEADER_ROW);
  let foundCol = null;
  row6.eachCell({ includeEmpty: false }, (cell, colNum) => {
    if (foundCol !== null) return;
    const v = String(cell.value ?? '').trim();
    if (/^Quote\b/i.test(v)) foundCol = colNum;
  });
  return foundCol || fallback;
}

// Bedingte Formatierung Spalte A + B (sauber ohne Repair-Warnings)
function addConditionalFormats(ws, tnEnd) {
  let prio = 1;
  const cfRule = (formula, fillArgb, opts = {}) => {
    const rule = {
      type: 'expression',
      formulae: [formula],
      priority: prio++,
      style: {
        fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: fillArgb } },
      },
    };
    if (opts.bold || opts.fontColor) {
      rule.style.font = {};
      if (opts.bold)      rule.style.font.bold  = true;
      if (opts.fontColor) rule.style.font.color = { argb: opts.fontColor };
    }
    return rule;
  };

  // Spalte A — Freigabe-Status
  ws.addConditionalFormatting({
    ref: `A${TN_START}:A${tnEnd}`,
    rules: [
      cfRule(`$A${TN_START}="ausgeladen"`,                COLORS.DUNKELROT, { bold: true, fontColor: 'FFFFFFFF' }),
      cfRule(`$A${TN_START}="Ausfall, da kein Reminder"`, COLORS.DUNKELROT, { bold: true, fontColor: 'FFFFFFFF' }),
      cfRule(`$A${TN_START}="abgesagt"`,                  COLORS.ROT,       { bold: true, fontColor: 'FFFFFFFF' }),
      cfRule(`$A${TN_START}="onhold (nicht ins Update)"`, COLORS.ORANGE),
      cfRule(`$A${TN_START}="Interviewer Freigabe"`,      COLORS.GELB),
      cfRule(`$A${TN_START}="Admin Freigabe"`,            COLORS.GRUEN),
      cfRule(`$A${TN_START}="Ersatz"`,                    COLORS.HELLGRUEN),
      cfRule(`$A${TN_START}="Umterminierung"`,            COLORS.HELLBLAU),
      cfRule(`$A${TN_START}="Umterminierung (Kunde)"`,    COLORS.HELLBLAU),
    ],
  });

  // Spalte B — Projektabschluss
  ws.addConditionalFormatting({
    ref: `B${TN_START}:B${tnEnd}`,
    rules: [
      cfRule(`$B${TN_START}="teilgenommen"`,                 COLORS.GRUEN),
      cfRule(`$B${TN_START}="ausgezahlt"`,                   'FFE2EFDA'),
      cfRule(`$B${TN_START}="nicht erschienen"`,             COLORS.GRAU, { fontColor: 'FFFFFFFF' }),
      cfRule(`$B${TN_START}="kam zu spät ohne Ankündigung"`, COLORS.GRAU, { fontColor: 'FFFFFFFF' }),
      cfRule(`$B${TN_START}="kam zu spät mit Ankündigung"`,  COLORS.GRAU, { fontColor: 'FFFFFFFF' }),
      cfRule(`$B${TN_START}="kurzfristig abgesagt"`,         COLORS.GRAU, { fontColor: 'FFFFFFFF' }),
      cfRule(`$B${TN_START}="abgesagt durch Kunde"`,         COLORS.GRAU, { fontColor: 'FFFFFFFF' }),
    ],
  });
}

// Logo aus logos.js einfügen
function addLogo(dstWs, dstWorkbook, unternehmen) {
  const logo = LOGOS[unternehmen];
  if (!logo?.base64 || logo.base64.startsWith('HIER_')) return;
  try {
    const logoId = dstWorkbook.addImage({ base64: logo.base64, extension: logo.ext });
    dstWs.addImage(logoId, {
      tl: { col: 0, row: 0 },
      ext: { width: logo.width || 200, height: logo.height || 60 },
      editAs: 'oneCell',
    });
  } catch (e) {
    console.error('Logo Fehler:', e.message);
  }
}

// Tracking-Bereich aufräumen: Zeilen TN_END+1 bis ca. 30 in Spalten 1..QUOTE_START-1 leeren
function clearTrackingArea(ws, tnEnd, quoteStartCol) {
  for (let r = tnEnd + 1; r <= 30; r++) {
    for (let col = 1; col < quoteStartCol; col++) {
      const c = ws.getCell(r, col);
      c.value = null;
      c.fill = { type: 'pattern', pattern: 'none' };
      c.border = {};
      c.font = undefined;
      c.alignment = undefined;
    }
  }
}

// Vorbefüllte TN-Zeilen aus Template entfernen (z.B. "1..10" in Spalte E, "Admin Freigabe" in A7)
function clearPrefilledTNCells(ws, tnEnd, quoteStartCol) {
  for (let r = TN_START; r <= 30; r++) {
    for (let col = 1; col < quoteStartCol; col++) {
      const c = ws.getCell(r, col);
      // Alle Werte in Spalten links der Quotes leeren — diese werden später
      // entweder neu befüllt (lfd. Nr.) oder bleiben leer (Vorname, etc.)
      c.value = null;
    }
  }
}

// Cell-Styling für Antwort-Codes
function styleAnswerCell(cell, ant, isOffTarget) {
  const screenout = !!ant.screenout;
  const offTarget = !screenout && isOffTarget;
  cell.value = `${ant.code} | ${ant.text}`;
  cell.fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: screenout ? COLORS.DUNKELROT : (offTarget ? COLORS.HELLROT : COLORS.WEISS) },
  };
  cell.font = {
    name: 'Arial', size: 9,
    bold: screenout,
    color: { argb: screenout ? 'FFFFFFFF' : (offTarget ? 'FF9C0006' : 'FF000000') },
  };
  cell.border = { ...THIN_BORDER };
  cell.alignment = { wrapText: true, vertical: 'top' };
}

// Off-Target-Erkennung pro Frage und Gruppe
// Markiert Antwort-Codes als off_target wenn sie nicht zur Zielgruppe der Gruppe passen
// (z.B. "weiblich" in einer Männer-Gruppe, oder "47-50" in einer 35-46-Gruppe)
function isAntOffTarget(frage, ant, gruppe) {
  const ftLower = (frage.fragetext || '').toLowerCase();
  const antLower = (ant.text || '').toLowerCase();

  // Geschlecht
  if (ftLower.match(/geschlecht/)) {
    if (gruppe.geschlecht === 'männlich' && antLower.match(/weiblich/)) return true;
    if (gruppe.geschlecht === 'weiblich' && antLower.match(/männlich/)) return true;
  }

  // Alter — versuche Range im Text zu erkennen "35-38", "47-50"
  if (ftLower.match(/alt|alter/) && gruppe.alter_min && gruppe.alter_max) {
    const m = ant.text.match(/(\d{2})\s*[-–]\s*(\d{2})/);
    if (m) {
      const lo = parseInt(m[1], 10);
      const hi = parseInt(m[2], 10);
      if (hi < gruppe.alter_min || lo > gruppe.alter_max) return true;
    }
  }

  return false;
}

// Schreibt eine einzelne Frage-Spalte (oder Item-Spalte einer Matrix)
function writeQuestionColumn(ws, col, label, note, antList, quoteText, tnEnd, gruppe, frage) {
  ws.getColumn(col).width = 22;

  const h = ws.getCell(HEADER_ROW, col);
  h.value = label;
  h.font = { bold: true, name: 'Arial', size: 9, color: { argb: 'FF000000' } };
  h.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
  h.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.WEISS } };
  h.border = HEADER_BORDER;
  if (note) h.note = note;

  // TN-Zellen (genau brutto-Zeilen)
  for (let r = TN_START; r <= tnEnd; r++) {
    const c = ws.getCell(r, col);
    c.value = null;
    c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.WEISS } };
    c.border = { ...THIN_BORDER };
    c.alignment = { wrapText: true, vertical: 'top' };
  }

  // Antwort-Codes
  if (antList && antList.length) {
    let row = tnEnd + ANT_OFFSET;
    for (const ant of antList) {
      const offTarget = isAntOffTarget(frage, ant, gruppe);
      const cell = ws.getCell(row, col);
      styleAnswerCell(cell, ant, offTarget);
      row++;
    }
    // Quote-Hinweis (grün) als letzte Zeile
    if (quoteText) {
      const qc = ws.getCell(row, col);
      qc.value = quoteText;
      qc.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.GRUEN } };
      qc.font = { name: 'Arial', size: 9, bold: true, color: { argb: COLORS.QUOTE_FONT } };
      qc.border = { ...THIN_BORDER };
      qc.alignment = { wrapText: true, vertical: 'top' };
    }
  }
}

// ---------------------------------------------------------------------------
// 4) HAUPT-FILL-LOGIK
// ---------------------------------------------------------------------------

function fillSheet(ws, gruppe, fragen, projektnummer, projektname, kundenname, setting, methode, quoteStartCol) {
  const hmap = HEADER_MAP[setting] || HEADER_MAP.offline;
  const brutto = gruppe.brutto || 8;
  const netto = gruppe.netto || 6;
  const tnEnd = TN_START + brutto - 1;
  const labelRow = tnEnd + QUOTE_HINT_AFTER_TN_OFFSET;
  const lfdCol = 5;

  // 1) Vorbefüllte Tracking-Werte aus Template leeren
  clearPrefilledTNCells(ws, tnEnd, quoteStartCol);

  // 2) Header befüllen
  if (hmap.studioLabel) {
    ws.getCell(hmap.studioLabel).value = 'F&T Standort';
  }
  if (hmap.studioWert && gruppe.standort) {
    ws.getCell(hmap.studioWert).value = gruppe.standort;
  }
  ws.getCell(hmap.terminWert).value     = gruppe.termin || `${gruppe.datum || ''} ${gruppe.uhrzeit || ''}`.trim();
  ws.getCell(hmap.kunde).value          = kundenname || '';
  ws.getCell(hmap.zielgruppe).value     = gruppe.zielgruppe || '';
  ws.getCell(hmap.projekt).value        = projektname;
  ws.getCell(hmap.projNr).value         = projektnummer;
  ws.getCell(hmap.incentive).value      = gruppe.incentive || '';

  // 3) lfd. Nr. neu setzen + "X für Y"-Label
  const lfdHeader = ws.getCell(HEADER_ROW, lfdCol);
  lfdHeader.note = `Brutto: ${brutto} TN\nNetto: ${netto} TN\n→ ${brutto} für ${netto}`;
  for (let r = TN_START; r <= tnEnd; r++) {
    const c = ws.getCell(r, lfdCol);
    c.value = r - TN_START + 1;
    c.alignment = { horizontal: 'center', vertical: 'center' };
    c.font = { name: 'Arial', size: 10, bold: true };
  }
  const lblCell = ws.getCell(labelRow, lfdCol);
  lblCell.value = `${brutto} für ${netto}`;
  lblCell.font = { name: 'Arial', size: 9, bold: true, italic: true, color: { argb: COLORS.QUOTE_FONT } };
  lblCell.alignment = { horizontal: 'center', vertical: 'center' };

  // 4) Tracking-Bereich rechts und unter TN ausnullen
  clearTrackingArea(ws, tnEnd, quoteStartCol);

  // 5) Quote-N-Reste in Header-Zeile leeren
  for (let col = quoteStartCol; col <= quoteStartCol + 30; col++) {
    const c = ws.getCell(HEADER_ROW, col);
    if (c.value && /^Quote/i.test(String(c.value))) {
      c.value = null;
    }
  }

  // 6) Fragen-Spalten aufbauen
  let currentCol = quoteStartCol;
  for (const frage of (fragen || [])) {
    if (frage.typ === 'entfaellt') {
      // Diese Frage wird übersprungen
      continue;
    }
    if (frage.typ === 'matrix' && Array.isArray(frage.items) && frage.items.length) {
      // Matrix: pro Item eine Spalte + Mutter-Header in Zeile 5 gemerged
      const matrixStart = currentCol;
      for (const item of frage.items) {
        const itemNote = `MUTTERFRAGE: ${frage.fragetext || ''}\n\n${frage.quotenkommentar || ''}\n\n${item.note || ''}\n\n${item.marker ? 'Marker: ' + item.marker : ''}`.trim();
        const itemQuote = item.quote_text || (item.is_quote_relevant && frage.quotenkommentar ? frage.quotenkommentar : '');
        // Antworten pro Item: screenout wird aus item.screenout_codes berechnet wenn nicht direkt gesetzt
        const ants = (item.antworten || []).map(a => ({
          ...a,
          screenout: a.screenout === true ? true :
                     (item.screenout_codes || []).includes(a.code),
        }));
        writeQuestionColumn(ws, currentCol, item.item_label || '', itemNote,
                            ants, itemQuote, tnEnd, gruppe, frage);
        currentCol++;
      }
      const matrixEnd = currentCol - 1;
      // Mutter-Header in Z5 gemerged
      if (matrixEnd > matrixStart) {
        try {
          ws.mergeCells(MATRIX_HEADER_ROW, matrixStart, MATRIX_HEADER_ROW, matrixEnd);
        } catch (e) {}
      }
      const mh = ws.getCell(MATRIX_HEADER_ROW, matrixStart);
      mh.value = frage.fragetext || frage.id;
      mh.font = { name: 'Arial', size: 10, bold: true };
      mh.alignment = { horizontal: 'center', vertical: 'center', wrapText: true };
      mh.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.MATRIX_HDR } };
      mh.border = HEADER_BORDER;
      if (frage.quotenkommentar) mh.note = frage.quotenkommentar;
    } else if (frage.typ === 'freitext' || frage.typ === 'numerisch') {
      // Freitext / Numerisch: nur Header, keine Antwort-Codes
      const note = frage.bedingung ? `Bedingung: ${frage.bedingung}` : (frage.fragetext || '');
      const label = frage.id ? `${frage.id}. ${frage.kurzlabel || frage.fragetext || ''}`.substring(0, 60) : (frage.fragetext || '');
      writeQuestionColumn(ws, currentCol, label, note, null, null, tnEnd, gruppe, frage);
      currentCol++;
    } else {
      // single_choice, multi_choice, ranking
      const label = frage.id ? `${frage.id}. ${frage.fragetext || ''}` : (frage.fragetext || '');
      const note = frage.bedingung ? `Bedingung: ${frage.bedingung}` : '';
      writeQuestionColumn(ws, currentCol, label, note,
                          frage.antworten, frage.quotenkommentar, tnEnd, gruppe, frage);
      currentCol++;
    }
  }

  // Höhere Zeilen für Matrix-Header und Frage-Header
  ws.getRow(MATRIX_HEADER_ROW).height = 32;
  ws.getRow(HEADER_ROW).height = 75;

  // Freeze Panes (Spalten A-E fix, ab Vorname F scrollt; Zeile 6 fix vertikal)
  ws.views = [{ state: 'frozen', xSplit: 5, ySplit: 6 }];
}

// ---------------------------------------------------------------------------
// 5) HANDLER
// ---------------------------------------------------------------------------

export default async function handler(req, res) {
  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'Method not allowed' });
  }
  try {
    const {
      templateBase64,
      projektnummer,
      projektname,
      kundenname,
      methode,
      isIDI,
      gruppen,
      fragen,
    } = req.body ?? {};

    const auftraggeber = kundenname || projektname || '';

    if (!templateBase64) return res.status(400).json({ error: 'Missing templateBase64' });
    if (!gruppen?.length) return res.status(400).json({ error: 'Missing gruppen' });

    const fragenArr = Array.isArray(fragen) ? fragen : [];
    const setting = (methode === 'VGD' || methode === 'VDI') ? 'online' : 'offline';
    const cfg = SHEET_CONFIG[methode] || SHEET_CONFIG['GD'];

    const template = new ExcelJS.Workbook();
    await template.xlsx.load(Buffer.from(templateBase64, 'base64'));
    const srcWs = template.getWorksheet(cfg.sheetName);
    if (!srcWs) {
      return res.status(400).json({
        error: `Sheet '${cfg.sheetName}' not found`,
        available: template.worksheets.map(w => w.name),
      });
    }

    const result = new ExcelJS.Workbook();
    result.creator = 'Preview Generator';
    result.created = new Date();
    const newSheetNames = [];

    const processGruppe = (gruppe, sheetName, fragenFG) => {
      const dstWs = result.addWorksheet(sheetName);
      copyWorksheet(srcWs, dstWs);

      const methodeFG = gruppe.methode || methode;
      const cfgFG = SHEET_CONFIG[methodeFG] || SHEET_CONFIG['GD'];
      const quoteStartCol = findQuoteStartCol(dstWs, cfgFG.quoteStartCol);
      const brutto = gruppe.brutto || 8;
      const tnEnd = TN_START + brutto - 1;

      addConditionalFormats(dstWs, tnEnd);
      addLogo(dstWs, result, gruppe.unternehmen);
      fillSheet(
        dstWs, gruppe, fragenFG, projektnummer, projektname,
        auftraggeber, setting, methodeFG, quoteStartCol
      );
      newSheetNames.push({ sheet: sheetName, quoteStartCol, brutto, tnEnd });
    };

    if (isIDI) {
      const hauptGruppe = { ...gruppen[0], methode };
      processGruppe(hauptGruppe, methode === 'VDI' ? 'VDIs' : 'IDIs', fragenArr);
    } else {
      for (const gruppe of gruppen) {
        gruppe.methode = gruppe.methode || methode;
        const sheetName = gruppe.id.replace(/[:\\/\?\*\[\]]/g, '').substring(0, 31);
        const fragenFG = fragenArr.filter(f =>
          !f.relevantFuerGruppen ||
          f.relevantFuerGruppen.includes('alle') ||
          f.relevantFuerGruppen.includes(gruppe.id)
        );
        processGruppe(gruppe, sheetName, fragenFG);
      }
    }

    const buffer = await result.xlsx.writeBuffer();
    const dateiname = `${projektnummer}_${projektname}_Preview.xlsx`
      .replace(/[^a-zA-Z0-9_\-\.äöüÄÖÜß ]/g, '_');

    let downloadUrl = null;
    let excelBase64 = null;
    let blobError = null;

    if (process.env.BLOB_READ_WRITE_TOKEN) {
      try {
        const { put } = await import('@vercel/blob');
        const blob = await put(`previews/${dateiname}`, buffer, {
          access: 'public',
          contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
          addRandomSuffix: true,
        });
        downloadUrl = blob.url;
      } catch (e) {
        console.error('Blob Upload fehlgeschlagen:', e.message);
        blobError = e.message;
        excelBase64 = Buffer.from(buffer).toString('base64');
      }
    } else {
      excelBase64 = Buffer.from(buffer).toString('base64');
    }

    return res.status(200).json({
      downloadUrl,
      excelBase64,
      dateiname,
      success: true,
      debug: {
        fragenCount: fragenArr.length,
        gruppenCount: gruppen.length,
        sheets: newSheetNames,
        setting,
        kundenname: auftraggeber,
        deliveryMode: downloadUrl ? 'blob' : 'base64',
        blobError,
        version: 'v2-matrix-schema',
      },
    });
  } catch (err) {
    console.error('Error:', err);
    return res.status(500).json({ error: err?.message, stack: err?.stack });
  }
}
