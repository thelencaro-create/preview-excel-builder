// api/build-preview.js
// Preview Generator – Excel Builder
// Fixt: (1) Kundenname-Feld, (2) CF Spalte A, (3) CF Spalte B, (4) Logo, (5) Rahmen
import ExcelJS from "exceljs";

// ---------------------------------------------------------------------------
// 1) KONFIGURATION
// ---------------------------------------------------------------------------

const SHEET_CONFIG = {
  GD:  { sheetName: 'GD',   quoteStartCol: 14 },
  IDI: { sheetName: 'IDIs', quoteStartCol: 16 },
  VGD: { sheetName: 'VGDs', quoteStartCol: 20 },
  VDI: { sheetName: 'VDIs', quoteStartCol: 22 },
};

const HEADER_MAP = {
  offline: {
    studio:     'H1',
    termin:     'L1',
    kunde:      'I2',
    zielgruppe: 'L2',
    projekt:    'I3',
    projNr:     'I4',
    incentive:  'L4',
  },
  online: {
    studio:     null,
    termin:     'M1',
    kunde:      'J1',
    zielgruppe: 'J3',
    projekt:    'J2',
    projNr:     'J4',
    incentive:  'M4',
  }
};

const TN_START  = 7;
const TN_END    = 16;
const ANT_START = TN_END + 2;

// ---------------------------------------------------------------------------
// 2) FARBEN (exakt aus Übergabedokument-Palette)
// ---------------------------------------------------------------------------

const COLORS = {
  DUNKELROT: 'FFC00000', // ausgeladen / Ausfall
  ROT:       'FFFF0000', // abgesagt
  ORANGE:    'FFFF5050', // onhold
  GELB:      'FFFFE699', // Interviewer Freigabe
  GRUEN:     'FFA9D08E', // Admin Freigabe / teilgenommen
  HELLGRUEN: 'FFC6E0B4', // Ersatz
  SEHRHELL:  'FFE2EFDA', // ausgezahlt
  HELLBLAU:  'FFBDD7EE', // Umterminierung (auch Kunde)
  GRAU:      'FF808080', // alle anderen B
  WEISS:     'FFFFFFFF',
  BORDER:    'FFCCCCCC',
};

const THIN_BORDER = {
  top:    { style: 'thin', color: { argb: COLORS.BORDER } },
  bottom: { style: 'thin', color: { argb: COLORS.BORDER } },
  left:   { style: 'thin', color: { argb: COLORS.BORDER } },
  right:  { style: 'thin', color: { argb: COLORS.BORDER } },
};

// ---------------------------------------------------------------------------
// 3) LOGOS – Base64 hart eingebettet (Bug #4)
//    Hier die echten Base64-Strings einfügen. Bis dahin werden Logos
//    übersprungen ohne den Build zu brechen.
// ---------------------------------------------------------------------------

const LOGOS = {
  'F&T': { base64: 'HIER_FT_BASE64', ext: 'png' },
  'm-s': { base64: 'HIER_MS_BASE64', ext: 'jpg' },
  'H+G': { base64: 'HIER_HG_BASE64', ext: 'jpg' },
};

// ---------------------------------------------------------------------------
// 4) HILFSFUNKTIONEN
// ---------------------------------------------------------------------------

function copyWorksheet(srcWs, dstWs) {
  // Spaltenbreiten
  srcWs.columns.forEach((col, i) => {
    const dstCol = dstWs.getColumn(i + 1);
    if (col.width)  dstCol.width  = col.width;
    if (col.hidden) dstCol.hidden = col.hidden;
  });

  // Zellen + Styles
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

  // Merges
  if (srcWs._merges) {
    Object.keys(srcWs._merges).forEach(key => {
      try { dstWs.mergeCells(key); } catch (e) {}
    });
  }

  // Dropdowns / Data Validations
  if (srcWs.dataValidations?.model) {
    Object.entries(srcWs.dataValidations.model).forEach(([sqref, dv]) => {
      try { dstWs.dataValidations.add(sqref, { ...dv }); } catch (e) {}
    });
  }

  // WICHTIG (Bug #2 + #3):
  // Bedingte Formatierung NICHT vom Template kopieren.
  // ExcelJS überträgt strikethrough/bold-Attribute fehlerhaft.
  // Stattdessen: addConditionalFormats() manuell aufrufen.
}

// Bug #2 + #3: Bedingte Formatierung manuell sauber setzen
//
// Wichtig für saubere DXF-Erzeugung (sonst meckert Excel beim Öffnen):
//  - jede Rule braucht eine explizite, eindeutige `priority`
//  - solid fill in DXF nutzt `bgColor` (NICHT `fgColor` – das ist nur bei
//    normalen Cell-Fills so; in DXF dreht Excel die Bedeutung um)
//  - font im DXF minimal halten: nur die Properties setzen, die
//    wirklich überschrieben werden sollen. `name`, `size`,
//    `italic: false`, `strike: false` als unnötige Overrides
//    triggern Repair-Warnings.
function addConditionalFormats(ws) {
  let prio = 1;

  const cfRule = (formula, fillArgb, opts = {}) => {
    const rule = {
      type: 'expression',
      formulae: [formula],
      priority: prio++,
      style: {
        fill: {
          type: 'pattern',
          pattern: 'solid',
          bgColor: { argb: fillArgb },
        },
      },
    };
    // Font NUR setzen wenn wirklich override nötig
    if (opts.bold || opts.fontColor) {
      rule.style.font = {};
      if (opts.bold)      rule.style.font.bold  = true;
      if (opts.fontColor) rule.style.font.color = { argb: opts.fontColor };
    }
    return rule;
  };

  // Spalte A – Freigabe-Status
  ws.addConditionalFormatting({
    ref: `A${TN_START}:A${TN_END}`,
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

  // Spalte B – Projektabschluss
  ws.addConditionalFormatting({
    ref: `B${TN_START}:B${TN_END}`,
    rules: [
      cfRule(`$B${TN_START}="teilgenommen"`,                 COLORS.GRUEN),
      cfRule(`$B${TN_START}="ausgezahlt"`,                   COLORS.SEHRHELL),
      cfRule(`$B${TN_START}="nicht erschienen"`,             COLORS.GRAU, { fontColor: 'FFFFFFFF' }),
      cfRule(`$B${TN_START}="kam zu spät ohne Ankündigung"`, COLORS.GRAU, { fontColor: 'FFFFFFFF' }),
      cfRule(`$B${TN_START}="kam zu spät mit Ankündigung"`,  COLORS.GRAU, { fontColor: 'FFFFFFFF' }),
      cfRule(`$B${TN_START}="kurzfristig abgesagt"`,         COLORS.GRAU, { fontColor: 'FFFFFFFF' }),
      cfRule(`$B${TN_START}="abgesagt durch Kunde"`,         COLORS.GRAU, { fontColor: 'FFFFFFFF' }),
    ],
  });
}

// Bug #4: Logo aus eingebettetem Base64 einfügen
function addLogo(dstWs, dstWorkbook, unternehmen) {
  const logo = LOGOS[unternehmen];
  if (!logo?.base64 || logo.base64.startsWith('HIER_')) {
    return; // Platzhalter – einfach skippen
  }
  try {
    const logoId = dstWorkbook.addImage({ base64: logo.base64, extension: logo.ext });
    dstWs.addImage(logoId, {
      tl: { col: 0, row: 0 },
      br: { col: 3, row: 4 },
      editAs: 'oneCell',
    });
  } catch (e) {
    console.error('Logo Fehler:', e.message);
  }
}

function styleAnswer(cell, isScreenout) {
  cell.fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: isScreenout ? COLORS.DUNKELROT : COLORS.WEISS },
  };
  cell.font = {
    name: 'Arial', size: 9,
    bold: isScreenout,
    color: { argb: isScreenout ? 'FFFFFFFF' : 'FF000000' },
  };
  cell.border    = { ...THIN_BORDER };
  cell.alignment = { wrapText: true, vertical: 'top' };
}

function fillSheet(ws, gruppe, fragen, projektnummer, projektname, kundenname, setting, methode) {
  const hmap   = HEADER_MAP[setting] || HEADER_MAP.offline;
  const cfg    = SHEET_CONFIG[methode] || SHEET_CONFIG['GD'];
  const termin = gruppe.termin || `${gruppe.datum || ''} ${gruppe.uhrzeit || ''}`.trim();

  // Header-Zellen
  if (hmap.studio) {
    ws.getCell(hmap.studio).value = `${gruppe.unternehmen || ''} ${gruppe.standort || ''}`.trim();
  }
  ws.getCell(hmap.termin).value     = termin;
  // Bug #1: Kundenname statt Projektname ins Kunde-Feld
  ws.getCell(hmap.kunde).value      = kundenname || projektname;
  ws.getCell(hmap.zielgruppe).value = gruppe.zielgruppe || 'Allgemein';
  ws.getCell(hmap.projekt).value    = projektname;
  ws.getCell(hmap.projNr).value     = projektnummer;
  ws.getCell(hmap.incentive).value  = gruppe.incentive || '';

  const startCol = cfg.quoteStartCol;

  // Fragen-Spalten aufbauen
  fragen.forEach((frage, fi) => {
    const col = startCol + fi;
    ws.getColumn(col).width = 22;

    // Header-Zelle
    const hCell = ws.getCell(6, col);
    hCell.value     = `${frage.id}. ${frage.fragetext}`;
    hCell.font      = { bold: true, name: 'Arial', size: 9, color: { argb: 'FF000000' } };
    hCell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
    hCell.fill      = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.WEISS } };
    hCell.border    = {
      top:    { style: 'thin', color: { argb: 'FF000000' } },
      bottom: { style: 'thin', color: { argb: 'FF000000' } },
      left:   { style: 'thin', color: { argb: COLORS.BORDER } },
      right:  { style: 'thin', color: { argb: COLORS.BORDER } },
    };
    if (frage.quotenKommentar) hCell.note = `Quoten:\n${frage.quotenKommentar}`;

    // Bug #5: TN-Zellen mit Rahmen
    for (let r = TN_START; r <= TN_END; r++) {
      const tnCell     = ws.getCell(r, col);
      tnCell.value     = null;
      tnCell.fill      = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.WEISS } };
      tnCell.border    = { ...THIN_BORDER };
      tnCell.alignment = { wrapText: true, vertical: 'top' };
    }
  });

  // Antwort-Referenz-Zeilen
  const maxAnt = Math.max(...fragen.map(f => (f.antworten || []).length), 0);
  for (let a = 0; a < maxAnt; a++) {
    const row = ANT_START + a;
    ws.getRow(row).height = 14;
    fragen.forEach((frage, fi) => {
      const col     = startCol + fi;
      const antwort = (frage.antworten || [])[a];
      if (!antwort) return;
      const cell = ws.getCell(row, col);
      cell.value = `${antwort.code} | ${antwort.text}`;
      styleAnswer(cell, !!antwort.screenout);
    });
  }

  // Übrige Quote-Header aufräumen
  for (let col = startCol + fragen.length; col <= startCol + 25; col++) {
    const h = ws.getCell(6, col);
    if (h.value && String(h.value).startsWith('Quote')) {
      h.value = null;
    } else {
      break;
    }
  }
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
      kundenname,        // Bug #1: jetzt eigenes Feld
      methode,
      isIDI,
      gruppen,
      fragen,
    } = req.body ?? {};

    // Fallback: alter Workflow ohne kundenname-Feld
    const auftraggeber = kundenname || projektname || '';

    if (!templateBase64) return res.status(400).json({ error: 'Missing templateBase64' });
    if (!gruppen?.length) return res.status(400).json({ error: 'Missing gruppen' });

    const fragenArr = Array.isArray(fragen) ? fragen : [];
    const setting   = (methode === 'VGD' || methode === 'VDI') ? 'online' : 'offline';
    const cfg       = SHEET_CONFIG[methode] || SHEET_CONFIG['GD'];

    // Template laden
    const template = new ExcelJS.Workbook();
    await template.xlsx.load(Buffer.from(templateBase64, 'base64'));
    const srcWs = template.getWorksheet(cfg.sheetName);
    if (!srcWs) {
      return res.status(400).json({
        error: `Sheet '${cfg.sheetName}' not found`,
        available: template.worksheets.map(w => w.name),
      });
    }

    // Output-Workbook
    const result   = new ExcelJS.Workbook();
    result.creator = 'Preview Generator';
    result.created = new Date();
    const newSheetNames = [];

    const processGruppe = (gruppe, sheetName, fragenFG) => {
      const dstWs = result.addWorksheet(sheetName);
      copyWorksheet(srcWs, dstWs);
      addConditionalFormats(dstWs);                           // Bug #2 + #3
      addLogo(dstWs, result, gruppe.unternehmen);             // Bug #4
      fillSheet(
        dstWs,
        gruppe,
        fragenFG,
        projektnummer,
        projektname,
        auftraggeber,                                          // Bug #1
        setting,
        gruppe.methode || methode,
      );
      newSheetNames.push(sheetName);
    };

    if (isIDI) {
      const hauptGruppe = { ...gruppen[0], methode };
      processGruppe(hauptGruppe, methode === 'VDI' ? 'VDIs' : 'IDIs', fragenArr);
    } else {
      for (const gruppe of gruppen) {
        gruppe.methode = gruppe.methode || methode;
        const sheetName = gruppe.id.replace(/[:\\/\?\*\[\]]/g, '').substring(0, 31);
        const fragenFG  = fragenArr.filter(f =>
          !f.relevantFuerGruppen ||
          f.relevantFuerGruppen.includes('alle') ||
          f.relevantFuerGruppen.includes(gruppe.id)
        );
        processGruppe(gruppe, sheetName, fragenFG);
      }
    }

    const buffer      = await result.xlsx.writeBuffer();
    const excelBase64 = Buffer.from(buffer).toString('base64');
    const dateiname   = `${projektnummer}_${projektname}_Preview.xlsx`
      .replace(/[^a-zA-Z0-9_\-\.äöüÄÖÜß ]/g, '_');

    return res.status(200).json({
      excelBase64,
      dateiname,
      success: true,
      debug: {
        fragenCount: fragenArr.length,
        gruppenCount: gruppen.length,
        sheets: newSheetNames,
        quoteStartCol: cfg.quoteStartCol,
        setting,
        kundenname: auftraggeber,
        bugfixes: ['kundenname', 'cf-spalte-a', 'cf-spalte-b', 'logo', 'rahmen'],
      },
    });
  } catch (err) {
    console.error('Error:', err);
    return res.status(500).json({ error: err?.message, stack: err?.stack });
  }
}
