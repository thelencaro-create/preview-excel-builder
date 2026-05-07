// api/build-preview.js
// Preview Generator – Excel Builder v6 (Schema mit Matrix/Freitext-Support)
import ExcelJS from "exceljs";
import { LOGOS } from './logos.js';

const SHEET_CONFIG = {
  GD:  { sheetName: 'GD',   quoteStartCol: 14 },
  IDI: { sheetName: 'IDIs', quoteStartCol: 16 },
  VGD: { sheetName: 'VGDs', quoteStartCol: 20 },
  VDI: { sheetName: 'VDIs', quoteStartCol: 22 },
};

const HEADER_MAP = {
  offline: {
    studioLabel: 'H1',
    studioWert:  'I1',
    terminLabel: 'L1', terminWert: 'L1',
    kunde:       'I2',
    zielgruppe:  'L2',
    projekt:     'I3',
    projNr:      'I4',
    incentive:   'L4',
  },
  online: {
    studioLabel: null,
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
const QUOTE_HINT_AFTER_TN_OFFSET = 1;
const ANT_OFFSET = 2;

const COLORS = {
  DUNKELROT:   'FFC00000',
  HELLROT:     'FFF4CCCC',
  GRUEN:       'FFA9D08E',
  HELLGRUEN:   'FFC6E0B4',
  HELLBLAU:    'FFBDD7EE',
  GELB:        'FFFFE699',
  ORANGE:      'FFFF5050',
  ROT:         'FFFF0000',
  GRAU:        'FF808080',
  WEISS:       'FFFFFFFF',
  BORDER:      'FFCCCCCC',
  THICK_BORDER: 'FF606060',
  MATRIX_HDR:  'FFE0E0E0',
  HEADER_GREY: 'FFF2F2F2',
  QUOTE_FONT:  'FF2E7D32',
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

function borderWithThickRight(baseBorder) {
  return {
    ...baseBorder,
    right: { style: 'medium', color: { argb: COLORS.THICK_BORDER } },
  };
}

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

function addLogo(dstWs, dstWorkbook, unternehmen, setting) {
  const logo = LOGOS[unternehmen];
  if (!logo?.base64 || logo.base64.startsWith('HIER_')) return;
  try {
    const logoId = dstWorkbook.addImage({ base64: logo.base64, extension: logo.ext });
    const startCol = (setting === 'online') ? 5 : 4;
    dstWs.addImage(logoId, {
      tl: { col: startCol, row: 0 },
      ext: { width: logo.width || 200, height: logo.height || 60 },
      editAs: 'oneCell',
    });
  } catch (e) {
    console.error('Logo Fehler:', e.message);
  }
}

function findLfdNrCol(ws, fallback) {
  const row6 = ws.getRow(HEADER_ROW);
  let foundCol = null;
  row6.eachCell({ includeEmpty: false }, (cell, colNum) => {
    if (foundCol !== null) return;
    const v = String(cell.value ?? '').trim().toLowerCase();
    if (v === 'lfd. nr.' || v === 'lfd.nr.' || v === 'lfd nr.' || v.startsWith('lfd')) {
      foundCol = colNum;
    }
  });
  return foundCol || fallback;
}

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

function clearPrefilledTNCells(ws, tnEnd, quoteStartCol) {
  for (let r = TN_START; r <= 30; r++) {
    for (let col = 1; col < quoteStartCol; col++) {
      const c = ws.getCell(r, col);
      c.value = null;
    }
  }
}

function styleAnswerCell(cell, ant, isOffTarget) {
  const screenout = !!ant.screenout;
  const offTarget = !screenout && isOffTarget;
  const isRedHighlighted = screenout || offTarget;
  cell.value = `${ant.code} | ${ant.text}`;
  cell.fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: isRedHighlighted ? COLORS.DUNKELROT : COLORS.WEISS },
  };
  cell.font = {
    name: 'Arial', size: 9,
    bold: isRedHighlighted,
    color: { argb: isRedHighlighted ? 'FFFFFFFF' : 'FF000000' },
  };
 // Sichtbare schwarze Borders rundum (statt zartes Grau, das in Excel kaum sichtbar ist)
  const blackThin = { style: 'thin', color: { argb: 'FF000000' } };
  cell.border = {
    top: blackThin, bottom: blackThin,
    left: blackThin, right: blackThin,
  };
  cell.alignment = { wrapText: true, vertical: 'top' };
}

function cleanQuoteText(raw) {
  if (!raw) return '';
  let s = String(raw).trim();
  const hasCodeSyntax = /\b(IN|AND|OR)\b\s*[\[\(]|\.code\b|\.item\d+\b/i.test(s);
  if (hasCodeSyntax) {
    return 'Quotenrelevant — siehe Mutterfrage-Hinweis (Hover auf Header)';
  }
  s = s.replace(/^\s*ALLE\s+(müssen|sollen)\s+/i, 'Alle TN ');
  s = s.replace(/\bCode\s+(\d)/gi, 'C$1');
  s = s.replace(/\s+/g, ' ').trim();
  if (!/^quote/i.test(s)) {
    s = 'Quote: ' + s;
  }
  return s;
}

function isAntOffTarget(frage, ant, gruppe) {
  const ftLower = (frage.fragetext || '').toLowerCase();
  const antLower = (ant.text || '').toLowerCase();
  if (ftLower.match(/geschlecht/)) {
    if (gruppe.geschlecht === 'männlich' && antLower.match(/weiblich/)) return true;
    if (gruppe.geschlecht === 'weiblich' && antLower.match(/männlich/)) return true;
  }
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

function writeQuestionColumn(ws, col, label, note, antList, quoteText, tnEnd, gruppe, frage, isLastInGroup) {
  ws.getColumn(col).width = 22;
  const cellBorder = isLastInGroup ? borderWithThickRight(THIN_BORDER) : { ...THIN_BORDER };
  const headerBorder = isLastInGroup ? borderWithThickRight(HEADER_BORDER) : HEADER_BORDER;
  const h = ws.getCell(HEADER_ROW, col);
  h.value = label;
  h.font = { bold: true, name: 'Arial', size: 9, color: { argb: 'FF000000' } };
  h.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
  h.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.HEADER_GREY } };
  h.border = headerBorder;
  if (note) h.note = note;
  const blackBottom = { style: 'medium', color: { argb: 'FF000000' } };
  for (let r = TN_START; r <= tnEnd; r++) {
    const c = ws.getCell(r, col);
    c.value = null;
    c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.WEISS } };
    const baseBorder = isLastInGroup ? borderWithThickRight(THIN_BORDER) : { ...THIN_BORDER };
    if (r === tnEnd) {
      c.border = {
        top: baseBorder.top, left: baseBorder.left, right: baseBorder.right,
        bottom: blackBottom,
      };
    } else {
      c.border = baseBorder;
    }
    c.alignment = { wrapText: true, vertical: 'top' };
  }
  let row = tnEnd + ANT_OFFSET;
  if (antList && antList.length) {
    for (const ant of antList) {
      const offTarget = isAntOffTarget(frage, ant, gruppe);
      const cell = ws.getCell(row, col);
      styleAnswerCell(cell, ant, offTarget);
      if (isLastInGroup) {
        cell.border = borderWithThickRight(cell.border || THIN_BORDER);
      }
      row++;
    }
  }
  const cleanedQuote = cleanQuoteText(quoteText);
  if (cleanedQuote) {
    const qc = ws.getCell(row, col);
    qc.value = cleanedQuote;
    qc.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.GRUEN } };
    qc.font = { name: 'Arial', size: 9, bold: true, color: { argb: COLORS.QUOTE_FONT } };
    qc.border = cellBorder;
    qc.alignment = { wrapText: true, vertical: 'top' };
  }
}

function fillSheet(ws, gruppe, fragen, projektnummer, projektname, kundenname, setting, methode, quoteStartCol) {
  const hmap = HEADER_MAP[setting] || HEADER_MAP.offline;
  const brutto = gruppe.brutto || 8;
  const netto = gruppe.netto || 6;
  const tnEnd = TN_START + brutto - 1;
  const labelRow = tnEnd + QUOTE_HINT_AFTER_TN_OFFSET;
  const lfdCol = findLfdNrCol(ws, setting === 'online' ? 6 : 5);
  clearPrefilledTNCells(ws, tnEnd, quoteStartCol);
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
  const headerValueCells = [
    hmap.terminWert, hmap.kunde, hmap.zielgruppe,
    hmap.projekt, hmap.projNr, hmap.incentive,
  ].filter(Boolean);
  for (const addr of headerValueCells) {
    const c = ws.getCell(addr);
    c.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
  }
  if (hmap.studioWert) {
    ws.getCell(hmap.studioWert).alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
  }
  const lfdHeader = ws.getCell(HEADER_ROW, lfdCol);
  lfdHeader.note = `Brutto: ${brutto} TN\nNetto: ${netto} TN\n→ ${brutto} für ${netto}`;
  const blackBottom = { style: 'medium', color: { argb: 'FF000000' } };
  for (let r = TN_START; r <= tnEnd; r++) {
    const c = ws.getCell(r, lfdCol);
    c.value = r - TN_START + 1;
    c.alignment = { horizontal: 'center', vertical: 'center' };
    c.font = { name: 'Arial', size: 10, bold: true };
    if (r === tnEnd) {
      const eb = c.border || {};
      c.border = {
        top: eb.top, left: eb.left, right: eb.right,
        bottom: blackBottom,
      };
    }
  }
  for (let col = 1; col < quoteStartCol; col++) {
    const c = ws.getCell(tnEnd, col);
    const eb = c.border || {};
    c.border = {
      top: eb.top, left: eb.left, right: eb.right,
      bottom: blackBottom,
    };
  }
  const lblCell = ws.getCell(labelRow, lfdCol);
  lblCell.value = `${brutto} für ${netto}`;
  lblCell.font = { name: 'Arial', size: 9, bold: true, italic: true, color: { argb: COLORS.QUOTE_FONT } };
  lblCell.alignment = { horizontal: 'center', vertical: 'center' };
  clearTrackingArea(ws, tnEnd, quoteStartCol);
  for (let col = quoteStartCol; col <= quoteStartCol + 30; col++) {
    const c = ws.getCell(HEADER_ROW, col);
    if (c.value && /^Quote/i.test(String(c.value))) {
      c.value = null;
    }
  }
  let currentCol = quoteStartCol;
  for (const frage of (fragen || [])) {
    if (frage.typ === 'entfaellt' ||
        frage.kategorie === 'entfaellt' ||
        frage.kategorie === 'verfuegbarkeit') {
      continue;
    }
    const istEffektivMatrix = frage.typ === 'matrix' ||
      (Array.isArray(frage.items) && frage.items.length > 0 &&
       frage.items[0].antworten && frage.items[0].antworten.length > 0);
    if (istEffektivMatrix && Array.isArray(frage.items) && frage.items.length) {
      const matrixQuoteText = frage.quotenkommentar || frage.bedingung || '';
      const matrixStart = currentCol;
      const sharedAnswers = Array.isArray(frage.antworten) ? frage.antworten : [];
      const matrixQuoteMentionsItem = (itemLabel) => {
        if (!matrixQuoteText || !itemLabel) return false;
        const il = itemLabel.toLowerCase().trim();
        return matrixQuoteText.toLowerCase().includes(il);
      };
      for (const item of frage.items) {
        const itemNote = `MUTTERFRAGE: ${frage.fragetext || ''}\n\n${matrixQuoteText}\n\n${item.note || ''}\n\n${item.marker ? 'Marker: ' + item.marker : ''}`.trim();
        let itemQuote = item.quote_text || '';
        if (!itemQuote && matrixQuoteMentionsItem(item.item_label)) {
          itemQuote = matrixQuoteText;
        }
        if (!itemQuote && item.marker === 'X') {
          const codes = (item.screenout_codes || []).join(', ');
          itemQuote = codes
            ? `Quotenrelevant: Code ${codes} = Screenout`
            : 'Quotenrelevant: Item muss zutreffen (Marker X)';
        }
        if (!itemQuote && item.marker === 'Y') {
          itemQuote = 'Quotenrelevant: Item darf NICHT zutreffen (Marker Y)';
        }
        const rawAnts = (item.antworten && item.antworten.length) ? item.antworten : sharedAnswers;
        const ants = rawAnts.map(a => ({
          ...a,
          screenout: a.screenout === true ? true :
                     (item.screenout_codes || []).map(String).includes(String(a.code)),
        }));
        const isLast = (item === frage.items[frage.items.length - 1]);
        writeQuestionColumn(ws, currentCol, item.item_label || '', itemNote,
                            ants, itemQuote, tnEnd, gruppe, frage, isLast);
        currentCol++;
      }
      const matrixEnd = currentCol - 1;
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
      mh.border = borderWithThickRight(HEADER_BORDER);
      if (matrixQuoteText) mh.note = matrixQuoteText;
    } else if (frage.typ === 'freitext' || frage.typ === 'numerisch') {
      const note = frage.bedingung ? `Bedingung: ${frage.bedingung}` : (frage.fragetext || '');
      const label = frage.id ? `${frage.id}. ${frage.kurzlabel || frage.fragetext || ''}`.substring(0, 60) : (frage.fragetext || '');
      writeQuestionColumn(ws, currentCol, label, note, null, null, tnEnd, gruppe, frage, true);
      currentCol++;
    } else {
      const label = frage.id ? `${frage.id}. ${frage.fragetext || ''}` : (frage.fragetext || '');
      const note = frage.bedingung ? `Bedingung: ${frage.bedingung}` : '';
      let quoteText = frage.quotenkommentar || frage.bedingung || '';
      if (!quoteText) {
        const ftLower = (frage.fragetext || '').toLowerCase();
        if (ftLower.includes('geschlecht') && gruppe.geschlecht && gruppe.geschlecht !== 'gemischt') {
          quoteText = `Quote: nur ${gruppe.geschlecht}`;
        } else if ((ftLower.includes('alt sind') || ftLower.includes('alter')) && gruppe.alter_min && gruppe.alter_max) {
          quoteText = `Quote: ${gruppe.alter_min}-${gruppe.alter_max} Jahre`;
        }
      }
      writeQuestionColumn(ws, currentCol, label, note,
                          frage.antworten, quoteText, tnEnd, gruppe, frage, true);
      currentCol++;
    }
  }
  ws.getRow(MATRIX_HEADER_ROW).height = 32;
  ws.getRow(HEADER_ROW).height = 75;
  ws.views = [{ state: 'normal' }];
}

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
      addLogo(dstWs, result, gruppe.unternehmen, setting);
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
        version: 'v6-defensive-borders',
      },
    });
  } catch (err) {
    console.error('Error in build-preview:', err);
    console.error('Stack:', err?.stack);
    return res.status(500).json({
      success: false,
      error: err?.message || 'Unknown error',
      errorType: err?.name || 'Error',
      stack: err?.stack ? String(err.stack).split('\n').slice(0, 8) : null,
      version: 'v6-defensive-borders',
    });
  }
}
