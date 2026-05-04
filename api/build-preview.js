import ExcelJS from "exceljs";

// ── FARBEN ───────────────────────────────────────────────────────────────────
const C = {
  headerBg:   '1F4E79', headerFont: 'FFFFFF',
  fragenBg:   '2E75B6', fragenFont: 'FFFFFF',
  screenout:  'C00000', screenoutFg: 'FFFFFF',
  adminBg:    'D9D9D9', lightBlue:  'D6E4F0',
  lightGray:  'F2F2F2', white:      'FFFFFF',
  dark:       '333333',
};

function styleHeader(cell, bg, fg, bold = true) {
  cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + bg } };
  cell.font = { bold, color: { argb: 'FF' + fg }, name: 'Arial', size: 10 };
  cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
  cell.border = {
    top: { style: 'thin' }, bottom: { style: 'thin' },
    left: { style: 'thin' }, right: { style: 'thin' }
  };
}

function styleData(cell, bg = 'FFFFFF') {
  cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + bg } };
  cell.font = { name: 'Arial', size: 10 };
  cell.border = {
    top:    { style: 'thin', color: { argb: 'FFCCCCCC' } },
    bottom: { style: 'thin', color: { argb: 'FFCCCCCC' } },
    left:   { style: 'thin', color: { argb: 'FFCCCCCC' } },
    right:  { style: 'thin', color: { argb: 'FFCCCCCC' } }
  };
  cell.alignment = { vertical: 'middle', wrapText: true };
}

function buildSheet(ws, gruppe, fragenFuerGruppe, projektnummer, projektname, methode) {
  const tn = gruppe.tnBrutto || 8;
  const termin = gruppe.termin || `${gruppe.datum || ''} ${gruppe.uhrzeit || ''}`.trim();

  // Spaltenbreiten
  const widths = { A:12, B:14, C:14, D:14, E:7, F:15, G:15, H:14, I:22, J:12, K:16, L:12, M:16 };
  Object.entries(widths).forEach(([col, w]) => { ws.getColumn(col).width = w; });

  // ── HEADER BLOCK (Zeilen 1-4) ─────────────────────────────────────────────
  const headerInfos = [
    ['Studio:',     `${gruppe.unternehmen || ''} ${gruppe.standort || ''}`.trim()],
    ['Termin:',     termin],
    ['Kunde:',      projektname],
    ['Zielgruppe:', gruppe.zielgruppe || 'Allgemein'],
  ];
  const rightLabels = [
    `Proj.-Nr.: ${projektnummer}`,
    `Incentive: ${gruppe.incentive || ''}`,
    `Methode: ${gruppe.methode || methode}`,
    '',
  ];

  headerInfos.forEach(([label, value], i) => {
    const row = i + 1;
    ['A','B','C','D'].forEach(col => {
      const cell = ws.getCell(`${col}${row}`);
      cell.fill = { type:'pattern', pattern:'solid', fgColor:{ argb:'FF'+C.adminBg } };
      cell.border = { top:{style:'thin'}, bottom:{style:'thin'}, left:{style:'thin'}, right:{style:'thin'} };
    });
    ws.mergeCells(`E${row}:G${row}`);
    const lc = ws.getCell(`E${row}`);
    lc.value = label;
    lc.font = { bold:true, name:'Arial', size:10, color:{ argb:'FF'+C.headerBg } };
    lc.fill = { type:'pattern', pattern:'solid', fgColor:{ argb:'FF'+C.lightBlue } };
    lc.alignment = { vertical:'middle' };
    lc.border = { top:{style:'thin'}, bottom:{style:'thin'}, left:{style:'thin'}, right:{style:'thin'} };

    ws.mergeCells(`H${row}:L${row}`);
    const vc = ws.getCell(`H${row}`);
    vc.value = value;
    vc.font = { name:'Arial', size:10 };
    vc.fill = { type:'pattern', pattern:'solid', fgColor:{ argb:'FFFFFFFF' } };
    vc.alignment = { vertical:'middle' };
    vc.border = { top:{style:'thin'}, bottom:{style:'thin'}, left:{style:'thin'}, right:{style:'thin'} };

    ws.mergeCells(`M${row}:N${row}`);
    const rc = ws.getCell(`M${row}`);
    rc.value = rightLabels[i];
    rc.font = { bold:true, name:'Arial', size:9, color:{ argb:'FF'+C.headerBg } };
    rc.fill = { type:'pattern', pattern:'solid', fgColor:{ argb:'FF'+C.lightBlue } };
    rc.alignment = { vertical:'middle' };
    rc.border = { top:{style:'thin'}, bottom:{style:'thin'}, left:{style:'thin'}, right:{style:'thin'} };
  });

  // Leerzeile 5
  ws.getRow(5).height = 6;

  // ── SPALTEN-HEADER (Zeile 6) ──────────────────────────────────────────────
  ws.getRow(6).height = 40;
  const adminLabels = ['Freigabe','Projektabschluss','Feedback TN','Incentive final'];
  ['A','B','C','D'].forEach((col, i) => {
    const cell = ws.getCell(`${col}6`);
    cell.value = adminLabels[i];
    styleHeader(cell, C.adminBg, C.dark);
  });

  styleHeader(ws.getCell('E6'), C.headerBg, C.headerFont);
  ws.getCell('E6').value = 'lfd. Nr.';

  const personalien = ['Vorname','Nachname','Rufnummer','E-Mail','Reminder','Kommentar','Anonymität','letzte TN'];
  ['F','G','H','I','J','K','L','M'].forEach((col, i) => {
    const cell = ws.getCell(`${col}6`);
    cell.value = personalien[i];
    styleHeader(cell, C.headerBg, C.headerFont);
  });

  // Screener-Fragen ab Spalte N (=14)
  const startCol = 14;
  fragenFuerGruppe.forEach((frage, fi) => {
    const colNum = startCol + fi;
    const cell = ws.getCell(6, colNum);
    cell.value = `${frage.id}. ${frage.fragetext}`;
    styleHeader(cell, C.fragenBg, C.fragenFont);
    ws.getColumn(colNum).width = 24;
    if (frage.quotenKommentar) {
      cell.note = { texts: [{ text: `Quoten:\n${frage.quotenKommentar}` }] };
    }
  });

  // ── TN-ZEILEN ─────────────────────────────────────────────────────────────
  for (let t = 0; t < tn; t++) {
    const rowNum = 7 + t;
    ws.getRow(rowNum).height = 20;
    ['A','B','C','D'].forEach(col => styleData(ws.getCell(`${col}${rowNum}`), C.adminBg));
    const lfd = ws.getCell(`E${rowNum}`);
    lfd.value = t + 1;
    styleData(lfd, C.lightBlue);
    lfd.alignment = { horizontal:'center', vertical:'middle' };
    ['F','G','H','I','J','K','L','M'].forEach(col => styleData(ws.getCell(`${col}${rowNum}`)));
    fragenFuerGruppe.forEach((_, fi) => styleData(ws.getCell(rowNum, startCol + fi)));
  }

  // ── ANTWORT-REFERENZZEILEN ────────────────────────────────────────────────
  const antStart = 7 + tn + 1;
  ws.getRow(antStart - 1).height = 6;
  const maxAntworten = Math.max(...fragenFuerGruppe.map(f => (f.antworten || []).length), 0);

  for (let a = 0; a < maxAntworten; a++) {
    const rowNum = antStart + a;
    ws.getRow(rowNum).height = 18;
    if (a === 0) {
      const lbl = ws.getCell(`E${rowNum}`);
      lbl.value = 'Antworten:';
      lbl.font = { bold:true, name:'Arial', size:9, color:{ argb:'FF'+C.headerBg } };
      lbl.fill = { type:'pattern', pattern:'solid', fgColor:{ argb:'FF'+C.lightBlue } };
      lbl.border = { top:{style:'thin'}, bottom:{style:'thin'}, left:{style:'thin'}, right:{style:'thin'} };
    }
    fragenFuerGruppe.forEach((frage, fi) => {
      const antwort = (frage.antworten || [])[a];
      if (!antwort) return;
      const cell = ws.getCell(rowNum, startCol + fi);
      cell.value = `${antwort.code} | ${antwort.text}`;
      cell.font = {
        name: 'Arial', size: 9, bold: !!antwort.screenout,
        color: { argb: antwort.screenout ? 'FF'+C.screenoutFg : 'FF'+C.dark }
      };
      cell.fill = {
        type: 'pattern', pattern: 'solid',
        fgColor: { argb: antwort.screenout ? 'FF'+C.screenout : 'FF'+C.lightGray }
      };
      cell.border = {
        top:{style:'thin',color:{argb:'FFCCCCCC'}}, bottom:{style:'thin',color:{argb:'FFCCCCCC'}},
        left:{style:'thin',color:{argb:'FFCCCCCC'}}, right:{style:'thin',color:{argb:'FFCCCCCC'}}
      };
      cell.alignment = { wrapText: true, vertical: 'middle' };
    });
  }

  ws.views = [{ state: 'frozen', xSplit: 5, ySplit: 6 }];
}

// ── MAIN HANDLER ─────────────────────────────────────────────────────────────
export default async function handler(req, res) {
  try {
    if (req.method !== 'POST') {
      return res.status(405).json({ error: 'Method not allowed' });
    }

    const {
      templateBase64, projektnummer, projektname,
      methode, isIDI, gruppen, fragen
    } = req.body ?? {};

    if (!templateBase64) {
      return res.status(400).json({ error: 'Missing templateBase64' });
    }

    // Template laden
    const workbook = new ExcelJS.Workbook();
    const templateBuffer = Buffer.from(templateBase64, 'base64');
    await workbook.xlsx.load(templateBuffer);

    // Alle alten Sheets entfernen
    while (workbook.worksheets.length > 0) {
      workbook.removeWorksheet(workbook.worksheets[0].id);
    }

    if (isIDI) {
      const ws = workbook.addWorksheet('IDI');
      const hauptGruppe = gruppen?.[0] || {
        id: 'IDI', methode, standort: '', unternehmen: '',
        datum: '', uhrzeit: '', termin: '', tnBrutto: 8,
        incentive: '', zielgruppe: 'Allgemein'
      };
      buildSheet(ws, hauptGruppe, fragen || [], projektnummer, projektname, methode);
    } else {
      for (const gruppe of (gruppen || [])) {
        const sheetName = (gruppe.id || 'GD1')
          .replace(/[:\\/\?\*\[\]]/g, '')
          .substring(0, 31);
        const ws = workbook.addWorksheet(sheetName);
        const fragenFuerGruppe = (fragen || []).filter(f =>
          !f.relevantFuerGruppen ||
          f.relevantFuerGruppen.includes('alle') ||
          f.relevantFuerGruppen.includes(gruppe.id)
        );
        buildSheet(ws, gruppe, fragenFuerGruppe, projektnummer, projektname, methode);
      }
    }

    // Excel als Buffer
    const buffer = await workbook.xlsx.writeBuffer();
    const excelBase64 = Buffer.from(buffer).toString('base64');
    const dateiname = `${projektnummer}_${projektname}_Preview.xlsx`
      .replace(/[^a-zA-Z0-9_\-\.äöüÄÖÜß ]/g, '_');

    return res.status(200).json({ excelBase64, dateiname, success: true });

  } catch (err) {
    console.error('preview-excel-builder error:', err);
    return res.status(500).json({ error: String(err?.message || err) });
  }
}
