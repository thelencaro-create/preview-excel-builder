import ExcelJS from "exceljs";

// ── KONFIGURATION ─────────────────────────────────────────────────────────────
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

const LOGO_BASE64 = ''; // <-- Base64-String des Logos einfügen

// ── SHEET KOPIEREN ────────────────────────────────────────────────────────────
function copyWorksheet(srcWs, dstWs, srcWorkbook, dstWorkbook) {

  // Spaltenbreiten
  srcWs.columns.forEach((col, i) => {
    const dstCol = dstWs.getColumn(i + 1);
    if (col.width)  dstCol.width  = col.width;
    if (col.hidden) dstCol.hidden = col.hidden;
  });

  // Zeilen + Zellen
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

  // Merged cells
  if (srcWs._merges) {
    Object.keys(srcWs._merges).forEach(key => {
      try { dstWs.mergeCells(key); } catch(e) {}
    });
  }

  // Dropdowns
  if (srcWs.dataValidations?.model) {
    Object.entries(srcWs.dataValidations.model).forEach(([sqref, dv]) => {
      try { dstWs.dataValidations.add(sqref, { ...dv }); } catch(e) {}
    });
  }

  // Bedingte Formatierung: aus Template kopieren, aber strikethrough entfernen
  // (ExcelJS's addConditionalFormatting mit style-Objekten erzeugt kaputtes XML –
  //  daher kopieren wir die Original-CF und bereinigen nur das strikethrough-Attribut)
  if (srcWs.conditionalFormattings?.length) {
    srcWs.conditionalFormattings.forEach(cf => {
      try {
        // Deep-clone damit wir das Original nicht verändern
        const cfClean = JSON.parse(JSON.stringify(cf));
        // strikethrough aus allen font-Objekten in den Rules entfernen
        if (cfClean.rules) {
          cfClean.rules.forEach(rule => {
            if (rule.style?.font) {
              delete rule.style.font.strikethrough;
              delete rule.style.font.strike;
            }
          });
        }
        dstWs.addConditionalFormatting(cfClean);
      } catch(e) {
        console.error('CF copy Fehler:', e.message);
      }
    });
  }

  // Logo
  if (LOGO_BASE64) {
    try {
      const logoId = dstWorkbook.addImage({ base64: LOGO_BASE64, extension: 'png' });
      dstWs.addImage(logoId, { tl: { col: 0, row: 0 }, br: { col: 3, row: 4 }, editAs: 'oneCell' });
    } catch(e) { console.error('Logo Fehler:', e.message); }
  } else if (srcWs._images?.length && srcWorkbook && dstWorkbook) {
    srcWs._images.forEach(img => {
      try {
        const srcImage = srcWorkbook.getImage(img.imageId);
        if (srcImage?.buffer) {
          const ext   = srcImage.extension || srcImage.type || 'png';
          const newId = dstWorkbook.addImage({ buffer: srcImage.buffer, extension: ext });
          const range = img.range ?? {
            tl: { col: img.col ?? 0, row: img.row ?? 0 },
            br: { col: img.col2 ?? 3, row: img.row2 ?? 4 },
          };
          dstWs.addImage(newId, range);
        }
      } catch(e) { console.error('Logo (dynamisch) Fehler:', e.message); }
    });
  }
}

// ── ANTWORT-STYLING ───────────────────────────────────────────────────────────
function styleAnswer(cell, isScreenout) {
  cell.fill = {
    type: 'pattern', pattern: 'solid',
    fgColor: { argb: isScreenout ? 'FFC00000' : 'FFFFFFFF' }
  };
  cell.font = {
    name: 'Arial', size: 9, bold: isScreenout,
    color: { argb: isScreenout ? 'FFFFFFFF' : 'FF000000' }
  };
  cell.border = {
    top:    { style: 'thin', color: { argb: 'FFCCCCCC' } },
    bottom: { style: 'thin', color: { argb: 'FFCCCCCC' } },
    left:   { style: 'thin', color: { argb: 'FFCCCCCC' } },
    right:  { style: 'thin', color: { argb: 'FFCCCCCC' } },
  };
  cell.alignment = { wrapText: true, vertical: 'top' };
}

const THIN_BORDER = {
  top:    { style: 'thin', color: { argb: 'FFCCCCCC' } },
  bottom: { style: 'thin', color: { argb: 'FFCCCCCC' } },
  left:   { style: 'thin', color: { argb: 'FFCCCCCC' } },
  right:  { style: 'thin', color: { argb: 'FFCCCCCC' } },
};

// ── SHEET BEFÜLLEN ────────────────────────────────────────────────────────────
function fillSheet(ws, gruppe, fragen, projektnummer, projektname, kundenname, setting, methode) {
  const hmap   = HEADER_MAP[setting] || HEADER_MAP.offline;
  const cfg    = SHEET_CONFIG[methode] || SHEET_CONFIG['GD'];
  const termin = gruppe.termin || `${gruppe.datum||''} ${gruppe.uhrzeit||''}`.trim();

  if (hmap.studio) {
    ws.getCell(hmap.studio).value = `${gruppe.unternehmen||''} ${gruppe.standort||''}`.trim();
  }
  ws.getCell(hmap.termin).value     = termin;
  ws.getCell(hmap.kunde).value      = kundenname || projektname; // BUG 1
  ws.getCell(hmap.zielgruppe).value = gruppe.zielgruppe || 'Allgemein';
  ws.getCell(hmap.projekt).value    = projektname;
  ws.getCell(hmap.projNr).value     = projektnummer;
  ws.getCell(hmap.incentive).value  = gruppe.incentive || '';

  const startCol = cfg.quoteStartCol;

  fragen.forEach((frage, fi) => {
    const col = startCol + fi;
    ws.getColumn(col).width = 22;

    const hCell = ws.getCell(6, col);
    hCell.value     = `${frage.id}. ${frage.fragetext}`;
    hCell.font      = { bold: true, name: 'Arial', size: 9, color: { argb: 'FF000000' } };
    hCell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
    hCell.fill      = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFFFF' } };
    hCell.border    = {
      top:    { style: 'thin', color: { argb: 'FF000000' } },
      bottom: { style: 'thin', color: { argb: 'FF000000' } },
      left:   { style: 'thin', color: { argb: 'FFCCCCCC' } },
      right:  { style: 'thin', color: { argb: 'FFCCCCCC' } },
    };

    if (frage.quotenKommentar) {
      hCell.note = `Quoten:\n${frage.quotenKommentar}`;
    }

    // BUG 5: Rahmen in TN-Zeilen
    for (let r = TN_START; r <= TN_END; r++) {
      const tnCell     = ws.getCell(r, col);
      tnCell.value     = null;
      tnCell.fill      = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFFFF' } };
      tnCell.border    = { ...THIN_BORDER };
      tnCell.alignment = { wrapText: true, vertical: 'top' };
    }
  });

  const maxAnt = Math.max(...fragen.map(f => (f.antworten||[]).length), 0);
  for (let a = 0; a < maxAnt; a++) {
    const row = ANT_START + a;
    ws.getRow(row).height = 14;

    fragen.forEach((frage, fi) => {
      const col     = startCol + fi;
      const antwort = (frage.antworten||[])[a];
      if (!antwort) return;
      const cell = ws.getCell(row, col);
      cell.value = `${antwort.code} | ${antwort.text}`;
      styleAnswer(cell, !!antwort.screenout);
    });
  }

  for (let col = startCol + fragen.length; col <= startCol + 25; col++) {
    const h = ws.getCell(6, col);
    if (h.value && String(h.value).startsWith('Quote')) {
      h.value = null;
    } else break;
  }
}

// ── MAIN HANDLER ──────────────────────────────────────────────────────────────
export default async function handler(req, res) {
  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'Method not allowed' });
  }

  try {
    const {
      templateBase64, projektnummer, projektname,
      kundenname,
      methode, isIDI, gruppen, fragen
    } = req.body ?? {};

    const auftraggeber = kundenname || projektname || '';

    console.log('Request:', {
      hasTemplate:  !!templateBase64,
      projektnummer, projektname,
      kundenname:   auftraggeber,
      methode, isIDI,
      gruppenCount: gruppen?.length,
      fragenCount:  fragen?.length,
    });

    if (!templateBase64) return res.status(400).json({ error: 'Missing templateBase64' });
    if (!gruppen?.length) return res.status(400).json({ error: 'Missing gruppen' });

    const fragenArr = Array.isArray(fragen) ? fragen : [];
    const setting   = (methode === 'VGD' || methode === 'VDI') ? 'online' : 'offline';
    const cfg       = SHEET_CONFIG[methode] || SHEET_CONFIG['GD'];

    const template = new ExcelJS.Workbook();
    await template.xlsx.load(Buffer.from(templateBase64, 'base64'));

    const srcWs = template.getWorksheet(cfg.sheetName);
    if (!srcWs) {
      return res.status(400).json({
        error:     `Sheet '${cfg.sheetName}' not found`,
        available: template.worksheets.map(w => w.name),
      });
    }

    const result   = new ExcelJS.Workbook();
    result.creator = 'Preview Generator';
    result.created = new Date();

    const newSheetNames = [];

    if (isIDI) {
      const hauptGruppe = { ...gruppen[0], methode };
      const sheetName   = methode === 'VDI' ? 'VDIs' : 'IDIs';
      const dstWs       = result.addWorksheet(sheetName);
      copyWorksheet(srcWs, dstWs, template, result);
      fillSheet(dstWs, hauptGruppe, fragenArr, projektnummer, projektname, auftraggeber, setting, methode);
      newSheetNames.push(sheetName);

    } else {
      for (const gruppe of gruppen) {
        gruppe.methode  = gruppe.methode || methode;
        const sheetName = gruppe.id.replace(/[:\\/\?\*\[\]]/g, '').substring(0, 31);
        const fragenFG  = fragenArr.filter(f =>
          !f.relevantFuerGruppen ||
          f.relevantFuerGruppen.includes('alle') ||
          f.relevantFuerGruppen.includes(gruppe.id)
        );
        const dstWs = result.addWorksheet(sheetName);
        copyWorksheet(srcWs, dstWs, template, result);
        fillSheet(dstWs, gruppe, fragenFG, projektnummer, projektname, auftraggeber, setting, gruppe.methode);
        newSheetNames.push(sheetName);
      }
    }

    const buffer      = await result.xlsx.writeBuffer();
    const excelBase64 = Buffer.from(buffer).toString('base64');
    const dateiname   = `${projektnummer}_${projektname}_Preview.xlsx`
      .replace(/[^a-zA-Z0-9_\-\.äöüÄÖÜß ]/g, '_');

    return res.status(200).json({
      excelBase64, dateiname, success: true,
      debug: {
        fragenCount:   fragenArr.length,
        gruppenCount:  gruppen.length,
        sheets:        newSheetNames,
        quoteStartCol: cfg.quoteStartCol,
        setting,
        kundenname:    auftraggeber,
      },
    });

  } catch (err) {
    console.error('Error:', err);
    return res.status(500).json({ error: err?.message, stack: err?.stack });
  }
}
