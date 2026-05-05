import ExcelJS from "exceljs";

// ── TEMPLATE KONFIGURATION ────────────────────────────────────────────────────
const SHEET_CONFIG = {
  GD:   { sheetName: 'GD',   quoteStartCol: 14 },
  IDI:  { sheetName: 'IDIs', quoteStartCol: 16 },
  VGD:  { sheetName: 'VGDs', quoteStartCol: 20 },
  VDI:  { sheetName: 'VDIs', quoteStartCol: 22 },
};

const HEADER_MAP = {
  offline: {
    studio:     'H1',
    termin:     ['L1','M1'],
    kunde:      ['I2','J2'],
    zielgruppe: ['L2','M2'],
    projekt:    ['I3','J3'],
    projNr:     ['I4','J4'],
    incentive:  ['L4','M4'],
  },
  online: {
    studio:     null,
    termin:     ['M1','N1'],
    kunde:      ['J1','K1'],
    zielgruppe: ['J3','K3'],
    projekt:    ['J2','K2'],
    projNr:     ['J4','K4'],
    incentive:  ['M4','N4'],
  }
};

const TN_START_ROW = 7;
const TN_END_ROW   = 16;
const ANT_START_ROW = TN_END_ROW + 2;

// ── HILFSFUNKTIONEN ───────────────────────────────────────────────────────────
function setCellValue(ws, cellOrRange, value) {
  if (Array.isArray(cellOrRange)) {
    ws.getCell(cellOrRange[0]).value = value;
  } else {
    ws.getCell(cellOrRange).value = value;
  }
}

function styleAnswerCell(cell, isScreenout) {
  if (isScreenout) {
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFC00000' } };
    cell.font = { bold: true, name: 'Arial', size: 9, color: { argb: 'FFFFFFFF' } };
  } else {
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFFFF' } };
    cell.font = { name: 'Arial', size: 9, color: { argb: 'FF000000' } };
  }
  cell.border = {
    top:    { style: 'thin', color: { argb: 'FFCCCCCC' } },
    bottom: { style: 'thin', color: { argb: 'FFCCCCCC' } },
    left:   { style: 'thin', color: { argb: 'FFCCCCCC' } },
    right:  { style: 'thin', color: { argb: 'FFCCCCCC' } },
  };
  cell.alignment = { wrapText: true, vertical: 'top' };
}

async function copySheet(workbook, sourceSheetName, newSheetName) {
  const sourceWs = workbook.getWorksheet(sourceSheetName);
  if (!sourceWs) throw new Error(`Sheet '${sourceSheetName}' not found in template`);

  const newWs = workbook.addWorksheet(newSheetName);

  sourceWs.columns.forEach((col, idx) => {
    if (col.width) newWs.getColumn(idx + 1).width = col.width;
  });

  sourceWs.eachRow({ includeEmpty: true }, (row, rowNum) => {
    const newRow = newWs.getRow(rowNum);
    newRow.height = row.height;
    row.eachCell({ includeEmpty: true }, (cell, colNum) => {
      const newCell = newRow.getCell(colNum);
      newCell.value = cell.value;
      if (cell.style) {
        try { newCell.style = JSON.parse(JSON.stringify(cell.style)); } catch(e) {}
      }
      if (cell.alignment) {
        newCell.alignment = { ...cell.alignment };
      }
    });
    newRow.commit();
  });

  if (sourceWs._merges) {
    Object.keys(sourceWs._merges).forEach(key => {
      try { newWs.mergeCells(key); } catch(e) {}
    });
  }

  if (sourceWs.dataValidations?.model) {
    Object.entries(sourceWs.dataValidations.model).forEach(([sqref, dv]) => {
      try { newWs.dataValidations.add(sqref, dv); } catch(e) {}
    });
  }

  return newWs;
}

async function buildGroupSheet(workbook, templateSheetName, sheetName, gruppe, fragen, projektnummer, projektname, setting) {
  const ws = await copySheet(workbook, templateSheetName, sheetName);

  const termin = gruppe.termin || `${gruppe.datum || ''} ${gruppe.uhrzeit || ''}`.trim();
  const hmap = setting === 'online' ? HEADER_MAP.online : HEADER_MAP.offline;

  if (hmap.studio && gruppe.unternehmen) {
    ws.getCell(hmap.studio).value = `${gruppe.unternehmen} ${gruppe.standort || ''}`.trim();
  }
  setCellValue(ws, hmap.termin,     termin);
  setCellValue(ws, hmap.kunde,      projektname);
  setCellValue(ws, hmap.zielgruppe, gruppe.zielgruppe || 'Allgemein');
  setCellValue(ws, hmap.projekt,    projektname);
  setCellValue(ws, hmap.projNr,     projektnummer);
  setCellValue(ws, hmap.incentive,  gruppe.incentive || '');

  const config = SHEET_CONFIG[gruppe.methode] || SHEET_CONFIG['GD'];
  const startCol = config.quoteStartCol;

  fragen.forEach((frage, fi) => {
    const colNum = startCol + fi;
    ws.getColumn(colNum).width = 22;

    const headerCell = ws.getCell(6, colNum);
    headerCell.value = `${frage.id}. ${frage.fragetext}`;
    headerCell.font = { bold: true, name: 'Arial', size: 9, color: { argb: 'FF000000' } };
    headerCell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };

    if (frage.quotenKommentar) {
      headerCell.note = `Quoten:\n${frage.quotenKommentar}`;
    }

    for (let t = TN_START_ROW; t <= TN_END_ROW; t++) {
      ws.getCell(t, colNum).value = null;
    }
  });

  const maxAntworten = Math.max(...fragen.map(f => (f.antworten || []).length), 0);
  for (let a = 0; a < maxAntworten; a++) {
    const rowNum = ANT_START_ROW + a;
    ws.getRow(rowNum).height = 14;
    fragen.forEach((frage, fi) => {
      const colNum = startCol + fi;
      const antwort = (frage.antworten || [])[a];
      if (!antwort) return;
      const cell = ws.getCell(rowNum, colNum);
      cell.value = `${antwort.code} | ${antwort.text}`;
      styleAnswerCell(cell, !!antwort.screenout);
    });
  }

  // Alte Quote-Spalten leeren
  for (let col = startCol + fragen.length; col <= startCol + 30; col++) {
    const h = ws.getCell(6, col);
    if (h.value && String(h.value).startsWith('Quote')) {
      h.value = null;
    } else break;
  }

  return ws;
}

// ── MAIN HANDLER ──────────────────────────────────────────────────────────────
export default async function handler(req, res) {
  try {
    if (req.method !== 'POST') {
      return res.status(405).json({ error: 'Method not allowed' });
    }

    const { templateBase64, projektnummer, projektname, methode, isIDI, gruppen, fragen } = req.body ?? {};

    if (!templateBase64) return res.status(400).json({ error: 'Missing templateBase64' });
    if (!gruppen?.length) return res.status(400).json({ error: 'Missing gruppen' });

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(Buffer.from(templateBase64, 'base64'));

    const setting = (methode === 'VGD' || methode === 'VDI') ? 'online' : 'offline';
    const config  = SHEET_CONFIG[methode] || SHEET_CONFIG['GD'];

    const newSheetNames = [];

    if (isIDI) {
      const hauptGruppe = { ...gruppen[0], methode };
      const sheetName = methode === 'VDI' ? 'VDIs' : 'IDIs';
      await buildGroupSheet(workbook, config.sheetName, sheetName, hauptGruppe, fragen, projektnummer, projektname, setting);
      newSheetNames.push(sheetName);
    } else {
      for (const gruppe of gruppen) {
        gruppe.methode = gruppe.methode || methode;
        const sheetName = gruppe.id.replace(/[:\\/\?\*\[\]]/g, '').substring(0, 31);
        const fragenFuerGruppe = (fragen || []).filter(f =>
          !f.relevantFuerGruppen ||
          f.relevantFuerGruppen.includes('alle') ||
          f.relevantFuerGruppen.includes(gruppe.id)
        );
        await buildGroupSheet(workbook, config.sheetName, sheetName, gruppe, fragenFuerGruppe, projektnummer, projektname, setting);
        newSheetNames.push(sheetName);
      }
    }

    // Template-Sheets entfernen
    workbook.worksheets
      .filter(ws => !newSheetNames.includes(ws.name))
      .map(ws => ws.id)
      .forEach(id => workbook.removeWorksheet(id));

    const buffer = await workbook.xlsx.writeBuffer();
    const excelBase64 = Buffer.from(buffer).toString('base64');
    const dateiname = `${projektnummer}_${projektname}_Preview.xlsx`
      .replace(/[^a-zA-Z0-9_\-\.äöüÄÖÜß ]/g, '_');

    return res.status(200).json({ excelBase64, dateiname, success: true });

  } catch (err) {
    console.error('preview-excel-builder error:', err);
    return res.status(500).json({ error: String(err?.message || err), stack: err?.stack });
  }
}
