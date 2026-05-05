import ExcelJS from "exceljs";

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

// Logos als Base64 – HIER DIE STRINGS EINFÜGEN:
const LOGOS = {
  'F&T': { base64: 'HIER_FT_BASE64', ext: 'png' },
  'm-s': { base64: 'HIER_MS_BASE64', ext: 'jpg' },
  'H+G': { base64: 'HIER_HG_BASE64', ext: 'jpg' },
};

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
      try { dstWs.mergeCells(key); } catch(e) {}
    });
  }

  // Dropdowns kopieren
  if (srcWs.dataValidations?.model) {
    Object.entries(srcWs.dataValidations.model).forEach(([sqref, dv]) => {
      try { dstWs.dataValidations.add(sqref, { ...dv }); } catch(e) {}
    });
  }

  // Bedingte Formatierung aus Template kopieren (funktioniert stabil)
  if (srcWs.conditionalFormattings?.length) {
    srcWs.conditionalFormattings.forEach(cf => {
      try { dstWs.addConditionalFormatting(cf); } catch(e) {}
    });
  }
}

function addLogo(dstWs, dstWorkbook, unternehmen) {
  const logo = LOGOS[unternehmen];
  if (!logo?.base64 || logo.base64.startsWith('HIER_')) return;
  try {
    const logoId = dstWorkbook.addImage({ base64: logo.base64, extension: logo.ext });
    dstWs.addImage(logoId, { tl: { col: 0, row: 0 }, br: { col: 3, row: 4 }, editAs: 'oneCell' });
  } catch(e) { console.error('Logo Fehler:', e.message); }
}

function styleAnswer(cell, isScreenout) {
  cell.fill      = { type: 'pattern', pattern: 'solid', fgColor: { argb: isScreenout ? 'FFC00000' : 'FFFFFFFF' } };
  cell.font      = { name: 'Arial', size: 9, bold: isScreenout, color: { argb: isScreenout ? 'FFFFFFFF' : 'FF000000' } };
  cell.border    = {
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

function fillSheet(ws, gruppe, fragen, projektnummer, projektname, kundenname, setting, methode) {
  const hmap   = HEADER_MAP[setting] || HEADER_MAP.offline;
  const cfg    = SHEET_CONFIG[methode] || SHEET_CONFIG['GD'];
  const termin = gruppe.termin || `${gruppe.datum||''} ${gruppe.uhrzeit||''}`.trim();

  if (hmap.studio) {
    ws.getCell(hmap.studio).value = `${gruppe.unternehmen||''} ${gruppe.standort||''}`.trim();
  }
  ws.getCell(hmap.termin).value     = termin;
  ws.getCell(hmap.kunde).value      = kundenname || projektname;
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
    if (frage.quotenKommentar) hCell.note = `Quoten:\n${frage.quotenKommentar}`;

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
    if (h.value && String(h.value).startsWith('Quote')) { h.value = null; } else break;
  }
}

export default async function handler(req, res) {
  if (req.method !== 'POST') return res.status(405).json({ error: 'Method not allowed' });
  try {
    const { templateBase64, projektnummer, projektname, kundenname,
            methode, isIDI, gruppen, fragen } = req.body ?? {};
    const auftraggeber = kundenname || projektname || '';
    if (!templateBase64) return res.status(400).json({ error: 'Missing templateBase64' });
    if (!gruppen?.length) return res.status(400).json({ error: 'Missing gruppen' });
    const fragenArr = Array.isArray(fragen) ? fragen : [];
    const setting   = (methode === 'VGD' || methode === 'VDI') ? 'online' : 'offline';
    const cfg       = SHEET_CONFIG[methode] || SHEET_CONFIG['GD'];
    const template  = new ExcelJS.Workbook();
    await template.xlsx.load(Buffer.from(templateBase64, 'base64'));
    const srcWs = template.getWorksheet(cfg.sheetName);
    if (!srcWs) return res.status(400).json({ error: `Sheet '${cfg.sheetName}' not found`, available: template.worksheets.map(w => w.name) });
    const result   = new ExcelJS.Workbook();
    result.creator = 'Preview Generator';
    result.created = new Date();
    const newSheetNames = [];

    const processGruppe = (gruppe, sheetName, fragenFG) => {
      const dstWs = result.addWorksheet(sheetName);
      copyWorksheet(srcWs, dstWs);
      addLogo(dstWs, result, gruppe.unternehmen);
      fillSheet(dstWs, gruppe, fragenFG, projektnummer, projektname, auftraggeber, setting, gruppe.methode || methode);
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
    const dateiname   = `${projektnummer}_${projektname}_Preview.xlsx`.replace(/[^a-zA-Z0-9_\-\.äöüÄÖÜß ]/g, '_');
    return res.status(200).json({
      excelBase64, dateiname, success: true,
      debug: { fragenCount: fragenArr.length, gruppenCount: gruppen.length, sheets: newSheetNames, quoteStartCol: cfg.quoteStartCol, setting, kundenname: auftraggeber },
    });
  } catch (err) {
    console.error('Error:', err);
    return res.status(500).json({ error: err?.message, stack: err?.stack });
  }
}
