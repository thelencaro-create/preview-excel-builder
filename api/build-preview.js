// api/build-preview.js
// Preview Generator – Excel Builder v8 (template-aware headers)
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
//
// v8 (07.05.2026): Header-Werte werden dynamisch in die Zelle direkt rechts
// vom Label geschrieben — funktioniert für F&T, m-s, H+G und alle künftigen
// Online-Templates ohne Code-Änderung. Fallback auf statische HEADER_MAP.
//
// v9 (07.05.2026): Sub-Quoten pro Antwort-Code aus dem Screener werden in
// der grünen Quote-Zelle aufgelistet (z.B. "• 35-38: Brutto 2"). Builder
// filtert automatisch nach Gruppen-Kontext (jüngere/ältere Gruppe).
//
// v10 (07.05.2026): Vollständige Screener-Robustheit
// - Patch 1: Compact-Matrix für riesige Matrizen (Option C: quotenrelevante
//   Items eigene Spalte, Rest in Sammelspalte) — Toggle via compactMatrix
// - Patch 2: Soll-Quoten gruppen-spezifisch (soll_quote_gilt_fuer)
// - Patch 3: Bedingungs-Markierung (gelber Rahmen + 🔀-Symbol bei
//   conditional Fragen)
// - Patch 4: Off-Target-Erkennung erweitert (Region/Stadt, Lifestage,
//   Marken-Verwendung)
// - Patch 5: Quotengrid-Hinweis im Sheet-Header (zuordnungs_kriterien)
// - Patch 6: Visueller Hinweis bei Fragen mit relevantFuerGruppen != alle
//
// v11 (08.05.2026): Universelle Screener-Robustheit
// Die 8 Patches aus der Diagnose-Tabelle (Raucher GD + JPM IDI):
// - P1: Erweiterung Screenout-Vokabular im Parser (BEENDEN/Schließen/X/*)
// - P2: Gruppen-Tabellen mit Vererbung leerer Zellen (Parser)
// - P3: Gruppen-spezifische Quoten ("Gruppen X & Y" → soll_quote_gilt_fuer)
// - P4: soll_quote als Liste mit gilt_fuer_gruppen pro Eintrag (Schema)
// - P5: kurz_label für Frage-Header (max 30 Zeichen, Parser-generiert)
// - P6: idiProfile[] mit Segment, Cluster, profile_quoten (Schema)
// - P7: segment_beschreibungen + studien_quoten in Spalte G ab Z+8 (Builder)
// - P8: typ='tool_input' für Algorithmus-Skalen (1 Sammelspalte)
import ExcelJS from "exceljs";
import { LOGOS } from './logos.js';
import { createRequire } from 'module';

// JSZip ist Dependency von exceljs, daher sollte es immer verfügbar sein.
// Wir laden über createRequire um async-import-Probleme zu vermeiden.
let JSZip;
try {
  const require_ = createRequire(import.meta.url);
  JSZip = require_('jszip');
} catch (e) {
  console.warn('JSZip nicht verfügbar — Merge-Parser nutzt ExcelJS-Fallback');
}

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
// Online-Map ist nur Fallback — primär greift detectHeaderPositions().
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

// ---------------------------------------------------------------------------
// 1b) DYNAMISCHE HEADER-ERKENNUNG (template-übergreifend)
// ---------------------------------------------------------------------------
//
// Hintergrund: F&T/H+G haben Labels in I/L (Werte J/M), m-s hat sie in H/K
// (Werte I/L). Statt eine Map pro Template zu pflegen, scannen wir die
// Header-Zeilen 1-4 nach bekannten Labels und schreiben den Wert in die
// Zelle direkt rechts daneben.
const HEADER_LABEL_PATTERNS = {
  kunde:      [/^kunde$/i],
  projekt:    [/^projekt$/i],            // exakt "Projekt", NICHT "Projekt-Nr."
  zielgruppe: [/^zielgruppe$/i],
  projNr:     [/projekt[-\s]?nr/i],      // "F&T Projekt-Nr.", "ms Projekt-Nr.", "HG Projekt-Nr."
  terminWert: [/^termin$/i],
  incentive:  [/^incentive$/i],
};

function colLetterFromIndex(n) {
  let s = '';
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

function readCellText(cell) {
  let v = cell.value;
  if (v && typeof v === 'object') {
    if (v.richText) v = v.richText.map(t => t.text).join('');
    else if (v.text) v = v.text;
    else return '';
  }
  return String(v ?? '').trim();
}

// Scannt Z1-4 / Spalten A-O nach Header-Labels und liefert Wert-Zellen
// (= jeweils die Spalte direkt rechts vom Label-Treffer).
function detectHeaderPositions(ws) {
  const found = {};
  for (let r = 1; r <= 4; r++) {
    for (let c = 1; c <= 15; c++) {
      const txt = readCellText(ws.getCell(r, c));
      if (!txt) continue;
      for (const [key, patterns] of Object.entries(HEADER_LABEL_PATTERNS)) {
        if (found[key]) continue; // erstes Vorkommen gewinnt
        if (patterns.some(p => p.test(txt))) {
          found[key] = `${colLetterFromIndex(c + 1)}${r}`;
          break;
        }
      }
    }
  }
  return found;
}

const TN_START = 7;
const MATRIX_HEADER_ROW = 5;
const HEADER_ROW = 6;
const QUOTE_HINT_AFTER_TN_OFFSET = 1;  // "X für Y"-Label = TN_END + 1
const ANT_OFFSET = 2;                  // Antworten beginnen TN_END + 2

// ---------------------------------------------------------------------------
// 2) FARBEN
// ---------------------------------------------------------------------------

const COLORS = {
  DUNKELROT:   'FFC00000', // Screenout UND Off-Target (außerhalb Zielgruppe)
  HELLROT:     'FFF4CCCC', // Reserve / nicht mehr aktiv genutzt
  GRUEN:       'FFA9D08E', // Quote-Hinweis
  HELLGRUEN:   'FFC6E0B4',
  HELLBLAU:    'FFBDD7EE',
  GELB:        'FFFFE699',
  ORANGE:      'FFFF5050',
  ROT:         'FFFF0000',
  GRAU:        'FF808080',
  WEISS:       'FFFFFFFF',
  BORDER:      'FFCCCCCC',
  THICK_BORDER: 'FF606060',  // dunkler Trenner zwischen Frage-Bereichen
  MATRIX_HDR:  'FFE0E0E0',   // mittleres Grau für Matrix-Mutter-Header (vorher zartes Blau)
  HEADER_GREY: 'FFF2F2F2',
  QUOTE_FONT:  'FF2E7D32',   // dunkles Grün für Quote-Text
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

// Border-Variante mit dickem dunklem rechten Rand — markiert das Ende eines Fragen-Bereichs
function borderWithThickRight(baseBorder) {
  return {
    ...baseBorder,
    right: { style: 'medium', color: { argb: COLORS.THICK_BORDER } },
  };
}

// ---------------------------------------------------------------------------
// 3) HILFSFUNKTIONEN
// ---------------------------------------------------------------------------

// v11.3: Liest Merge-Ranges direkt aus dem rohen XLSX-Buffer (ZIP+XML).
// Das umgeht alle ExcelJS-API-Inkonsistenzen — die mergeCell-Tags stehen
// garantiert in der sheet*.xml und können über Regex extrahiert werden.
//
// sheetIndex: 1-basiert (sheet1.xml = erstes Sheet, sheet2.xml = zweites...)
// sheetName:  Optionaler Name; wenn gegeben, wird der Index aus workbook.xml geholt
async function readMergesFromBuffer(buffer, sheetName) {
  if (!JSZip) return [];
  try {
    const zip = await JSZip.loadAsync(buffer);
    // 1) sheet-Index ermitteln aus workbook.xml
    let sheetIndex = 1;
    if (sheetName) {
      const wbXml = await zip.file('xl/workbook.xml')?.async('string');
      if (wbXml) {
        // <sheet name="VDIs" sheetId="3" r:id="rId3"/>
        // Match: <sheet name="VDIs" ... r:id="rIdX"/>
        const re = new RegExp(`<sheet[^>]*name="${sheetName.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}"[^>]*r:id="rId(\\d+)"`, 'i');
        const m = wbXml.match(re);
        if (m) {
          // r:id ist nicht 1:1 sheet-Index! Wir lesen es aus workbook.xml.rels
          const relsXml = await zip.file('xl/_rels/workbook.xml.rels')?.async('string');
          if (relsXml) {
            const relRe = new RegExp(`<Relationship[^>]*Id="rId${m[1]}"[^>]*Target="(?:worksheets/)?sheet(\\d+)\\.xml"`, 'i');
            const relMatch = relsXml.match(relRe);
            if (relMatch) sheetIndex = parseInt(relMatch[1], 10);
          }
        }
      }
    }
    const sheetXml = await zip.file(`xl/worksheets/sheet${sheetIndex}.xml`)?.async('string');
    if (!sheetXml) return [];
    // <mergeCell ref="L1:M1"/>
    const merges = [];
    const re = /<mergeCell\s+ref="([^"]+)"\s*\/>/gi;
    let match;
    while ((match = re.exec(sheetXml)) !== null) {
      merges.push(match[1]);
    }
    return merges;
  } catch (e) {
    return [];
  }
}

function copyWorksheet(srcWs, dstWs, preMerges) {
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
      // v11.3: Wenn Zelle Teil eines Merges ist, aber NICHT der Master,
      // dann nur Style kopieren, NICHT den Wert (sonst wird der Wert
      // in alle Zellen des Merges geschrieben und sieht nach Re-Merge doppelt aus)
      const isMergedSlave = srcCell.isMerged && srcCell.master &&
                            srcCell.master.address !== srcCell.address;
      if (!isMergedSlave) {
        dstCell.value = srcCell.value;
      }
      if (srcCell.font)      dstCell.font      = { ...srcCell.font };
      if (srcCell.fill)      dstCell.fill      = { ...srcCell.fill };
      if (srcCell.border)    dstCell.border    = { ...srcCell.border };
      if (srcCell.alignment) dstCell.alignment = { ...srcCell.alignment };
      if (srcCell.numFmt)    dstCell.numFmt    = srcCell.numFmt;
    });
    dstRow.commit();
  });

  // v11.3: Merges robust kopieren — Priorität:
  //  1. preMerges (vom Caller via XML-Parser geliefert) ← garantiert vollständig
  //  2. ExcelJS-API (drei Varianten: _merges Object, Map, model.merges)
  //  3. Cell-Iteration (cell.isMerged + master)
  const mergeRanges = new Set();
  // Variante 0: vom Caller vorab geparst (XML-Buffer-Lesung)
  if (Array.isArray(preMerges) && preMerges.length > 0) {
    for (const m of preMerges) mergeRanges.add(m);
  }
  // Variante 1: _merges als Object mit Range-Keys
  if (mergeRanges.size === 0 && srcWs._merges && typeof srcWs._merges === 'object' && !(srcWs._merges instanceof Map)) {
    for (const key of Object.keys(srcWs._merges)) {
      mergeRanges.add(key);
    }
  }
  // Variante 2: _merges als Map
  if (mergeRanges.size === 0 && srcWs._merges instanceof Map) {
    for (const key of srcWs._merges.keys()) {
      mergeRanges.add(key);
    }
  }
  // Variante 3: model.merges als Array
  if (mergeRanges.size === 0 && Array.isArray(srcWs.model?.merges)) {
    for (const m of srcWs.model.merges) {
      if (typeof m === 'string') mergeRanges.add(m);
    }
  }
  // Variante 4: Cell-Iteration über isMerged
  if (mergeRanges.size === 0) {
    const merges = new Map();
    srcWs.eachRow({ includeEmpty: true }, (row, rn) => {
      row.eachCell({ includeEmpty: true }, (cell, cn) => {
        if (cell.isMerged && cell.master) {
          const masterAddr = cell.master.address;
          if (!merges.has(masterAddr)) {
            merges.set(masterAddr, { minR: rn, maxR: rn, minC: cn, maxC: cn });
          }
          const m = merges.get(masterAddr);
          if (rn > m.maxR) m.maxR = rn;
          if (cn > m.maxC) m.maxC = cn;
          if (rn < m.minR) m.minR = rn;
          if (cn < m.minC) m.minC = cn;
        }
      });
    });
    for (const m of merges.values()) {
      try {
        const range = `${colLetterFromIndex(m.minC)}${m.minR}:${colLetterFromIndex(m.maxC)}${m.maxR}`;
        mergeRanges.add(range);
      } catch (e) {}
    }
  }
  for (const range of mergeRanges) {
    try { dstWs.mergeCells(range); } catch (e) {}
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
// WICHTIG: Position pro Setting unterschiedlich, weil offline und online Templates
// das Logo in unterschiedlichen Spalten haben.
function addLogo(dstWs, dstWorkbook, unternehmen, setting) {
  const logo = LOGOS[unternehmen];
  if (!logo?.base64 || logo.base64.startsWith('HIER_')) return;
  try {
    const logoId = dstWorkbook.addImage({ base64: logo.base64, extension: logo.ext });
    // Offline: Spalten E-H (Index 4-7), Online: Spalten F-H (Index 5-7)
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

// Findet dynamisch die "lfd. Nr."-Spalte im Template (Zeile 6).
// Offline = Spalte E (5), Online = Spalte F (6) — variiert je nach Template.
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

// v11.4: Findet Spalten-Indizes für Datum / Uhrzeit-Spalten in Z6
// Liefert {datum, uhrzeit} mit Spalten-Indizes (1-basiert) oder null
function findDatumUhrzeitCols(ws) {
  const row6 = ws.getRow(HEADER_ROW);
  let datumCol = null, uhrzeitCol = null;
  row6.eachCell({ includeEmpty: false }, (cell, colNum) => {
    const v = String(cell.value ?? '').trim().toLowerCase();
    if (datumCol === null && /^(datum|date)$/i.test(v)) datumCol = colNum;
    if (uhrzeitCol === null && /^(uhrzeit|zeit|time)$/i.test(v)) uhrzeitCol = colNum;
  });
  return { datum: datumCol, uhrzeit: uhrzeitCol };
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
// Off-Target-Codes (außerhalb Zielgruppe) und Screenout-Codes haben jetzt
// die gleiche Optik: dunkelroter Hintergrund, weiße fette Schrift.
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
  cell.alignment = { wrapText: true, vertical: 'middle' };
}

// Bereinigt Quote-Texte von Code-Syntax (z.B. "(F8.item2.code IN [1,2]) OR ...")
// und erzeugt lesbaren Klartext. Wenn der Text nach dem Bereinigen leer wäre,
// fällt die Funktion auf einen generischen Hinweis zurück.
function cleanQuoteText(raw) {
  if (!raw) return '';
  let s = String(raw).trim();

  // Wenn der Text Programmier-Syntax enthält (Klammern + IN/OR/AND/codes),
  // ist er für den User unbrauchbar — kompletten Hinweis durch Klartext ersetzen.
  const hasCodeSyntax = /\b(IN|AND|OR)\b\s*[\[\(]|\.code\b|\.item\d+\b/i.test(s);
  if (hasCodeSyntax) {
    return 'Quotenrelevant — siehe Mutterfrage-Hinweis (Hover auf Header)';
  }

  // Sonst: leichte Bereinigung von redundanten Pre-/Suffixen
  s = s.replace(/^\s*ALLE\s+(müssen|sollen)\s+/i, 'Alle TN ');
  s = s.replace(/\bCode\s+(\d)/gi, 'C$1');
  s = s.replace(/\s+/g, ' ').trim();

  // Wenn nicht schon mit "Quote" startet, vorne ergänzen
  if (!/^quote/i.test(s)) {
    s = 'Quote: ' + s;
  }

  return s;
}

// Off-Target-Erkennung pro Frage und Gruppe
// Markiert Antwort-Codes als off_target wenn sie nicht zur Zielgruppe der Gruppe passen
// (z.B. "weiblich" in einer Männer-Gruppe, oder "47-50" in einer 35-46-Gruppe)
//
// PATCH 4 (v10): Off-Target-Erkennung erweitert um:
//  - Stadt/Region (gruppe.standort)
//  - Code-basierte Off-Targets über Whitelist `ant.off_target_fuer_gruppen`
//    (vom Parser befüllt, falls Antwort-Text nicht eindeutig)
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

  // Stadt/Region — wenn Frage nach Wohnort fragt und Gruppe einen Standort hat,
  // sind andere Standort-Antworten off-target
  if (ftLower.match(/wohnen|stadt|standort|region/) && gruppe.standort) {
    const std = gruppe.standort.toLowerCase();
    if (std && std !== 'online' && antLower.length > 2 && !antLower.includes(std) && !std.includes(antLower)) {
      // Nur prüfen wenn die Antwort kein Freitext ist sondern eine konkrete Stadt
      // (Heuristik: Antwort enthält eine andere bekannte Stadt)
      const cities = ['köln', 'münchen', 'hannover', 'berlin', 'hamburg', 'bochum', 'frankfurt'];
      if (cities.some(c => antLower === c || antLower.startsWith(c + ' ') || antLower.endsWith(' ' + c))) {
        return true;
      }
    }
  }

  // Whitelist vom Parser: ant.off_target_fuer_gruppen = ["GD1", "GD3"]
  // → Antwort ist off-target für genau diese Gruppen
  if (Array.isArray(ant.off_target_fuer_gruppen) && ant.off_target_fuer_gruppen.includes(gruppe.id)) {
    return true;
  }

  return false;
}

// ---------------------------------------------------------------------------
// 3b) SUB-QUOTEN-LOGIK
// ---------------------------------------------------------------------------
//
// Sub-Quoten sind Verteilungs-Hinweise pro Antwort-Code aus dem Screener,
// z.B. "Brutto 2 je jüngere Gruppe" oder "je Gruppe 1-2" oder
// "Schwerpunkt: mind. brutto n=4 je Gruppe".
//
// Sie kommen aus dem Parser im Feld `ant.soll_quote` und werden in der
// grünen Quote-Zelle als Bulletpoint-Liste angehängt.

// Erkennt, ob ein soll_quote-Text auf die aktuelle Gruppe anwendbar ist.
// Schlüsselwörter "jüngere"/"ältere" werden gegen alter_min/alter_max geprüft.
// Ohne Gruppen-Kontext-Wörter: gilt für alle Gruppen.
function isSollQuoteApplicable(sollQuote, gruppe, allGruppen) {
  if (!sollQuote) return false;
  const s = String(sollQuote).toLowerCase();

  // Kein Gruppen-Kontext → gilt für alle
  const hasYoungContext = /(j[üu]ngere?|j[üu]ngeren)/i.test(s);
  const hasOldContext  = /([äa]ltere?|[äa]lteren)/i.test(s);
  if (!hasYoungContext && !hasOldContext) return true;

  // Mit Kontext: bestimmen ob aktuelle Gruppe "jung" oder "alt" ist.
  // Strategie: median des alter_min aller Gruppen → unter median = jung
  if (!gruppe.alter_min || !Array.isArray(allGruppen) || allGruppen.length < 2) {
    return true; // im Zweifel anzeigen
  }
  const mins = allGruppen
    .map(g => g.alter_min)
    .filter(v => typeof v === 'number');
  if (mins.length < 2) return true;
  const sorted = [...mins].sort((a, b) => a - b);
  const median = sorted[Math.floor(sorted.length / 2)];
  const isYoung = gruppe.alter_min < median;
  const isOld   = gruppe.alter_min >= median;

  if (hasYoungContext && !hasOldContext) return isYoung;
  if (hasOldContext && !hasYoungContext) return isOld;
  return true;
}

// Hilfsfunktion: Liefert die Liste der soll_quote-Einträge,
// egal ob als String (alt v9) oder als Liste von {text, gilt_fuer_gruppen} (v11)
function getSollQuoteList(raw) {
  if (!raw) return [];
  if (typeof raw === 'string') {
    const t = raw.trim();
    return t ? [{ text: t, gilt_fuer_gruppen: [] }] : [];
  }
  if (Array.isArray(raw)) {
    return raw
      .map(q => {
        if (typeof q === 'string') return { text: q.trim(), gilt_fuer_gruppen: [] };
        if (!q || typeof q !== 'object') return null;
        return {
          text: String(q.text || '').trim(),
          gilt_fuer_gruppen: Array.isArray(q.gilt_fuer_gruppen)
            ? q.gilt_fuer_gruppen.map(String)
            : [],
        };
      })
      .filter(q => q && q.text);
  }
  return [];
}

// Sammelt Sub-Quoten der nicht-screenout, nicht-off-target Antworten,
// gefiltert auf den aktuellen Gruppen-Kontext.
// v11: soll_quote ist jetzt eine Liste pro Antwort, mit gruppen-spezifischen Einträgen.
// Liefert sauber formatierte Bullet-Liste oder leeren String.
function buildSubQuoteList(antList, frage, gruppe, allGruppen) {
  if (!Array.isArray(antList) || antList.length === 0) return '';
  const lines = [];
  const seen = new Set();
  for (const ant of antList) {
    if (ant.screenout) continue;
    if (isAntOffTarget(frage, ant, gruppe)) continue;

    // v11: soll_quote ist Liste, jeder Eintrag kann eigene Gruppen-Whitelist haben
    const quotes = getSollQuoteList(ant.soll_quote);
    if (quotes.length === 0) continue;

    for (const q of quotes) {
      // Gruppen-spezifischer Filter: Eintrag gilt nur für bestimmte Gruppen
      if (Array.isArray(q.gilt_fuer_gruppen) && q.gilt_fuer_gruppen.length > 0) {
        if (!q.gilt_fuer_gruppen.includes(gruppe.id)) continue;
      }
      // Backward-Compat v10: ant.soll_quote_gilt_fuer (selten, aber möglich)
      if (Array.isArray(ant.soll_quote_gilt_fuer) && ant.soll_quote_gilt_fuer.length > 0) {
        if (!ant.soll_quote_gilt_fuer.includes(gruppe.id)) continue;
      }
      // Klartext-Schlüsselwort-Filter (jüngere/ältere) als Backup
      if (!isSollQuoteApplicable(q.text, gruppe, allGruppen)) continue;

      const txt = q.text.replace(/\s+/g, ' ');
      const key = `${ant.text}|${txt}`;
      if (seen.has(key)) continue;
      seen.add(key);
      lines.push(`• ${ant.text}: ${txt}`);
    }
  }
  return lines.join('\n');
}

// ---------------------------------------------------------------------------
// 3c) BEDINGUNGS-ERKENNUNG (Patch 3 + 6)
// ---------------------------------------------------------------------------
// Erkennt, ob eine Frage konditional ist (nur für bestimmte Gruppen oder
// abhängig von vorherigen Antworten). Wird visuell mit gelbem Rahmen
// markiert.
function isConditionalFrage(frage) {
  // Hat eine Bedingung
  if (frage.bedingung && String(frage.bedingung).trim().length > 0) return true;
  // Ist auf bestimmte Gruppen beschränkt
  if (Array.isArray(frage.relevantFuerGruppen)
      && frage.relevantFuerGruppen.length > 0
      && !frage.relevantFuerGruppen.includes('alle')) return true;
  return false;
}

// ---------------------------------------------------------------------------
// 3d) MATRIX-ITEM-PARTITIONIERUNG (Patch 1: Compact-Matrix)
// ---------------------------------------------------------------------------
// Teilt Matrix-Items in zwei Gruppen:
//   - quotenrelevant (Marker X/Y oder is_quote_relevant=true)
//   - "Sonstige" (alles andere)
// Wird nur angewendet, wenn:
//   - compactMatrix-Toggle ist true UND
//   - Anzahl Items >= compactThreshold (Default 5)
function partitionMatrixItems(items, compactMatrix, compactThreshold) {
  if (!Array.isArray(items)) return { quotaItems: [], otherItems: [] };
  if (!compactMatrix || items.length < compactThreshold) {
    // Kein Compact-Mode: alle bekommen eigene Spalte (= heutiges Verhalten)
    return { quotaItems: items.slice(), otherItems: [] };
  }
  const quotaItems = [];
  const otherItems = [];
  for (const it of items) {
    const isQuotaRelevant = it.is_quote_relevant === true
      || it.marker === 'X' || it.marker === 'Y'
      || (Array.isArray(it.screenout_codes) && it.screenout_codes.length > 0);
    if (isQuotaRelevant) quotaItems.push(it);
    else otherItems.push(it);
  }
  return { quotaItems, otherItems };
}

// Schreibt eine einzelne Frage-Spalte (oder Item-Spalte einer Matrix)
function writeQuestionColumn(ws, col, label, note, antList, quoteText, tnEnd, gruppe, frage, isLastInGroup, allGruppen, quoteRow) {
  ws.getColumn(col).width = 22;

  // Border-Variante je nachdem ob diese Spalte das Ende einer Frage/Matrix ist
  const cellBorder = isLastInGroup ? borderWithThickRight(THIN_BORDER) : { ...THIN_BORDER };
  let headerBorder = isLastInGroup ? borderWithThickRight(HEADER_BORDER) : HEADER_BORDER;

  // PATCH 3+6: Bedingungs-Markierung — gelber Rahmen bei conditional Fragen
  const isConditional = frage && isConditionalFrage(frage);
  if (isConditional) {
    const yellowMedium = { style: 'medium', color: { argb: 'FFD4A017' } };
    headerBorder = {
      top: yellowMedium,
      bottom: yellowMedium,
      left: yellowMedium,
      right: isLastInGroup ? borderWithThickRight({}).right : yellowMedium,
    };
  }

  const h = ws.getCell(HEADER_ROW, col);
  // v12.0: Bei Skala-Fragen mit Tool-Input-IDs (H1_5, C12_8, D3, E8_1 etc.) den
  // vollen Fragetext direkt unter dem Label anzeigen — sonst kann der Recruiter
  // bei einer reinen ID + 5er-Skala nicht zuordnen, WAS bewertet werden soll.
  const isToolInputId = /^[HCDE]\d/i.test((frage && frage.id) || '');
  const ftClean = (frage && frage.fragetext) ? frage.fragetext.trim() : '';
  const krzClean = (frage && (frage.kurz_label || frage.kurzlabel) || '').trim();
  const showFragetextInHeader = isToolInputId && ftClean.length > 20 &&
    ftClean.toLowerCase() !== krzClean.toLowerCase();

  const prefix = isConditional ? '🔀 ' : '';
  if (showFragetextInHeader) {
    h.value = {
      richText: [
        { text: prefix + label,           font: { bold: true,  name: 'Arial', size: 9, color: { argb: 'FF000000' } } },
        { text: '\n',                     font: { name: 'Arial', size: 8 } },
        { text: ftClean,                  font: { italic: true, name: 'Arial', size: 8, color: { argb: 'FF555555' } } },
      ],
    };
  } else {
    h.value = prefix + label;
    h.font = { bold: true, name: 'Arial', size: 9, color: { argb: 'FF000000' } };
  }
  h.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
  h.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: isConditional ? 'FFFFF2CC' : COLORS.HEADER_GREY } };
  h.border = headerBorder;
  if (note) h.note = note;

  // TN-Zellen (genau brutto-Zeilen)
  // Letzte TN-Zeile bekommt dickere schwarze Bottom-Border (visueller Abschluss TN-Bereich)
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
    c.alignment = { wrapText: true, vertical: 'middle' };
  }

  // Antwort-Codes
  let row = tnEnd + ANT_OFFSET;
  if (antList && antList.length) {
    for (const ant of antList) {
      const offTarget = isAntOffTarget(frage, ant, gruppe);
      const cell = ws.getCell(row, col);
      styleAnswerCell(cell, ant, offTarget);
      // Border ggf. mit dickem rechten Rand überschreiben
      if (isLastInGroup) {
        cell.border = borderWithThickRight(cell.border || THIN_BORDER);
      }
      row++;
    }
  }
  // Quote-Hinweis (grün) als letzte Zeile — auch wenn keine Antworten existieren
  // Text bereinigt von Code-Syntax + Sub-Quoten pro Antwort-Code (falls vorhanden)
  const cleanedQuote = cleanQuoteText(quoteText);
  const subQuoteList = buildSubQuoteList(antList, frage, gruppe, allGruppen);
  const finalQuoteText = [cleanedQuote, subQuoteList].filter(Boolean).join('\n');
  if (finalQuoteText) {
    // v11.3: Wenn quoteRow vom Caller übergeben wurde, an dieser einheitlichen
    // Position schreiben (1 Zeile Abstand unter der längsten Antwort-Liste);
    // sonst wie bisher direkt nach den Antworten dieser Spalte
    const targetRow = quoteRow || row;
    const qc = ws.getCell(targetRow, col);
    qc.value = finalQuoteText;
    qc.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.GRUEN } };
    qc.font = { name: 'Arial', size: 9, bold: true, color: { argb: COLORS.QUOTE_FONT } };
    qc.border = cellBorder;
    qc.alignment = { wrapText: true, vertical: 'middle' };
    // Zeilenhöhe der Quote-Zeile dynamisch erhöhen, wenn mehrzeilig
    const lineCount = finalQuoteText.split('\n').length;
    if (lineCount > 1) {
      const currentH = ws.getRow(targetRow).height || 15;
      const neededH = Math.min(15 + lineCount * 14, 200);
      if (neededH > currentH) ws.getRow(targetRow).height = neededH;
    }
  }
}

// ---------------------------------------------------------------------------
// 4) HAUPT-FILL-LOGIK
// ---------------------------------------------------------------------------

// v11: IDI-Profile, Segment-Beschreibungen und Studien-Quoten in Spalte G ausspielen
//
// Layout im JPM-Vorbild:
//   Z22+ Spalte G: Segment-Beschreibungen untereinander
//     "Segment 1 — Ambitious Maximisers"
//     "Jüngere bis mittlere Altersgruppe (18-50)"
//     "Eltern, die Vollzeit arbeiten"
//     "..."
//     (leere Zeile)
//     "Segment 2 — Experienced Optimisers"
//     "..."
//
// Die studien_quoten werden ans Ende angehängt unter Überschrift "Studien-Quoten".
function writeIdiInfoBlock(ws, opts) {
  const { idiProfile, segment_beschreibungen, studien_quoten, tnEnd } = opts;
  if (!segment_beschreibungen && !studien_quoten?.length && !idiProfile?.length) return;

  // Spalte G ist im IDI-Sheet "Feedback zum TN" — dort schreiben
  const col = 7; // G
  let row = tnEnd + 8; // genug Abstand zu TN-Daten und Quote-Hinweisen

  ws.getColumn(col).width = Math.max(ws.getColumn(col).width || 20, 30);

  // 1) Segment-Beschreibungen
  if (segment_beschreibungen && Object.keys(segment_beschreibungen).length > 0) {
    let segIdx = 1;
    for (const [name, desc] of Object.entries(segment_beschreibungen)) {
      const headerCell = ws.getCell(row, col);
      headerCell.value = `Segment ${segIdx} — ${name}`;
      headerCell.font = { name: 'Arial', size: 10, bold: true, color: { argb: 'FF1F4E79' } };
      headerCell.alignment = { wrapText: true, vertical: 'middle' };
      row++;
      // Beschreibung (kann mehrzeilig sein)
      const descCell = ws.getCell(row, col);
      descCell.value = String(desc);
      descCell.font = { name: 'Arial', size: 9 };
      descCell.alignment = { wrapText: true, vertical: 'middle' };
      const lineCount = String(desc).split('\n').length;
      ws.getRow(row).height = Math.min(15 + lineCount * 14, 200);
      row++;
      // Leere Zeile als Trenner
      row++;
      segIdx++;
    }
  }

  // 2) IDI-Profile (wenn idiProfile-Liste vorhanden)
  if (Array.isArray(idiProfile) && idiProfile.length > 0) {
    const headerCell = ws.getCell(row, col);
    headerCell.value = `IDI-Profile (${idiProfile.length} Interviews)`;
    headerCell.font = { name: 'Arial', size: 10, bold: true, color: { argb: 'FF1F4E79' } };
    row++;
    for (const p of idiProfile) {
      const segLabel = p.segment_oder?.length
        ? `${p.segment} (oder ${p.segment_oder.join(', ')})`
        : p.segment;
      const profileText = `${p.id}: ${segLabel}` +
        (p.cluster ? ` — ${p.cluster}` : '') +
        (p.profile_quoten?.length
          ? '\n  ' + p.profile_quoten.map(q => `• ${q.text}`).join('\n  ')
          : '');
      const c = ws.getCell(row, col);
      c.value = profileText;
      c.font = { name: 'Arial', size: 9 };
      c.alignment = { wrapText: true, vertical: 'middle' };
      const lc = profileText.split('\n').length;
      ws.getRow(row).height = Math.min(15 + lc * 14, 150);
      row++;
    }
    row++; // Trenner
  }

  // 3) Studien-Quoten (gelten studienweit)
  if (Array.isArray(studien_quoten) && studien_quoten.length > 0) {
    const headerCell = ws.getCell(row, col);
    headerCell.value = 'Studien-Quoten (gelten für alle TN)';
    headerCell.font = { name: 'Arial', size: 10, bold: true, color: { argb: 'FF1F4E79' } };
    row++;
    for (const q of studien_quoten) {
      const c = ws.getCell(row, col);
      c.value = `• ${q}`;
      c.font = { name: 'Arial', size: 9 };
      c.alignment = { wrapText: true, vertical: 'middle' };
      row++;
    }
  }
}

// v11: Frage-Header-Label bauen.
//  - Bevorzugt frage.kurz_label (Parser-generiert, max ~30 Zeichen)
//  - Fallback: frage.fragetext (gekürzt auf 60 Zeichen)
//  - Frage-Nummer (id) wird vorangestellt, wenn nicht schon im kurz_label enthalten
// v11.4: Soziodemographische Fragen sortieren — Geschlecht und Alter
// IMMER ganz vorne, dann andere persoenliche_daten, dann themenbezogen.
//
// Die Ranking-Logik:
//   0 = Geschlecht
//   1 = Alter
//   2 = andere persoenliche_daten (Familienstand, Kinder, Bildung, Beruf, ...)
//   3 = themenbezogen
//   4 = sonstige (verfuegbarkeit etc., werden vorher gefiltert)
//
// Innerhalb jeder Gruppe bleibt die ursprüngliche Reihenfolge (stabile Sortierung
// per Index-Tag), damit z.B. Q1 Beruf vor Q5 Bildung erscheint, falls beide
// "andere persoenliche_daten" sind.
function sortFragenSoziodemoFirst(fragen) {
  if (!Array.isArray(fragen)) return [];
  const REGEX_GESCHLECHT = /(geschlecht|geschlechts[\s-]?identit|m[äa]nnlich.*weiblich)/i;
  const REGEX_ALTER = /\balter\b|wie alt sind sie|altersgruppe/i;

  const score = (f) => {
    const txt = ((f.fragetext || '') + ' ' + (f.kurz_label || '')).toLowerCase();
    if (REGEX_GESCHLECHT.test(txt)) return 0;
    if (REGEX_ALTER.test(txt)) return 1;
    if (f.kategorie === 'persoenliche_daten') return 2;
    if (f.kategorie === 'themenbezogen') return 3;
    return 4;
  };

  // Stabile Sortierung mit ursprünglichem Index als Tiebreaker
  return [...fragen]
    .map((f, i) => ({ f, i, s: score(f) }))
    .sort((a, b) => a.s - b.s || a.i - b.i)
    .map(x => x.f);
}

// v11.4: Reduziert ein zu langes kurz_label intelligent auf 1-2 Wörter.
// Sinnvolle Strategien:
//   - "Q1. Berufliches Ausschlusskriterium" → "Beruf"
//   - "Q2g. Haushalts-Brutto-Einkommen" → "HHI"  (wenn als Abk. parsbar)
//   - "Q5b. Berufstätigkeit" → "Berufstätigkeit"
//   - "Q2d. Beziehungsstatus" → "Beziehung"
//
// Hauptstrategie: nimm das erste "Substantiv-artige" Wort, das mindestens
// 4 Zeichen lang ist und kein Füllwort.
function compactKurzLabel(label) {
  if (!label) return label;
  // Q-Nummer-Präfix entfernen (wird beim buildHeaderLabel wieder ergänzt)
  let txt = label.replace(/^[FQ]\d+[a-z]?\.\s*/i, '').trim();
  if (txt.length <= 20) return txt; // schon kurz genug
  // Bekannte lange Phrasen → kurze Form
  const REPLACEMENTS = [
    [/\bBerufliche?s?\s+Ausschlusskriteri\w*/i, 'Beruf'],
    [/\bHaushalts[\s-]?(Brutto[\s-]?)?Einkommen\b/i, 'HHI'],
    [/\bPers[öo]nliche\s+Ersparnisse\b/i, 'Ersparnisse'],
    [/\bAnlage[\s-]?Verm[öo]gen\b/i, 'IA'],
    [/\bRegion\s+Deutschlands?\b/i, 'Region'],
    [/\bBeziehungsstatus\b/i, 'Beziehung'],
    [/\bKinder\s+im\s+Haushalt\b/i, 'Kinder'],
    [/\bArbeitsstatus\b/i, 'Arbeit'],
    [/\bGeschlechts[\s-]?Identit[äa]t\b/i, 'Geschlecht'],
    [/\bAltersgruppe\b/i, 'Alter'],
  ];
  for (const [re, repl] of REPLACEMENTS) {
    if (re.test(txt)) return repl;
  }
  // Fallback: erstes Wort über 5 Zeichen
  const words = txt.split(/[\s\-]+/).filter(w => w.length >= 5 && !/^(eine?|der|die|das|den|dem|für|mit|von|bei|aus|und|oder|als|wie|was|sind|hat|hatte)$/i.test(w));
  if (words.length > 0) return words[0];
  return txt.substring(0, 25);
}

function buildHeaderLabel(frage) {
  const id = (frage.id || '').trim();
  let kurz = (frage.kurz_label || frage.kurzlabel || '').trim();
  const lang = (frage.fragetext || '').trim();
  // v11.4: kurz_label intelligent reduzieren wenn zu lang
  kurz = compactKurzLabel(kurz);
  if (kurz) {
    // Wenn kurz_label schon mit der ID startet, nicht doppelt
    if (id && !kurz.toLowerCase().startsWith(id.toLowerCase())) {
      return `${id}. ${kurz}`.substring(0, 40);
    }
    return kurz.substring(0, 40);
  }
  if (id && lang) return `${id}. ${lang}`.substring(0, 40);
  return (lang || id).substring(0, 40);
}

// v12.3: Formatiert Segment-Quotenkommentare aus dem Parser als mehrzeilige
// Liste mit aufgelösten Segment-Namen aus idiProfile.
//
// Input:  "Quote: Segment 1: 4 oder 5; Segment 2: 0,1,2 oder 3; Segment 5: 3 oder 4"
// Output: "Quote pro Segment:
//          Segment 1 (Ambitious Maximisers): 4 oder 5
//          Segment 2 (Experienced Optimisers): 0,1,2 oder 3
//          Segment 5 (Aspiring Apprentice): 3 oder 4"
//
// Bei Texten ohne "Segment N:" Pattern wird der Input unverändert zurückgegeben.
function formatSegmentQuotes(text, idiProfile) {
  if (!text || typeof text !== 'string') return text;
  const segMentions = text.match(/Segment\s*\d+\s*:/gi);
  if (!segMentions || segMentions.length < 2) return text;

  // "Quote:" prefix entfernen
  let body = text.replace(/^\s*Quote\s*:\s*/i, '');

  // Tail abtrennen (z.B. "Hinweis für Rekrutierer: ...")
  let tail = '';
  const tailRe = /\s*[;|]\s*((?:Hinweis|Note|Anmerkung|HINWEIS)[^]*)$/i;
  const tm = body.match(tailRe);
  if (tm) {
    body = body.substring(0, tm.index);
    tail = tm[1].trim();
  }

  // Split bei "; Segment N:" oder "| Segment N:"
  const parts = body
    .split(/\s*[;|]\s*(?=Segment\s*\d+\s*:)/i)
    .map(p => p.trim())
    .filter(Boolean);
  if (parts.length < 2) return text;

  // Segment-Nr → Name Mapping aus idiProfile
  // Quelle 1: explizite segment_nr-Felder (vom Parser via PATCH 10 erweitert)
  // Quelle 2: Auftreten-Reihenfolge der unique Segments in idiProfile
  const segNumToName = {};
  if (Array.isArray(idiProfile) && idiProfile.length > 0) {
    // 1) explizite segment_nr
    for (const p of idiProfile) {
      if (p.segment && p.segment_nr != null) {
        const n = parseInt(p.segment_nr);
        if (!isNaN(n) && !segNumToName[n]) segNumToName[n] = p.segment;
      }
    }
    // 2) Fallback: Auftreten-Reihenfolge mit Nrs aus dem Text matchen
    if (Object.keys(segNumToName).length === 0) {
      const uniqueSegs = [];
      for (const p of idiProfile) {
        if (p.segment && !uniqueSegs.includes(p.segment)) uniqueSegs.push(p.segment);
      }
      const numsInText = [];
      for (const part of parts) {
        const m = part.match(/Segment\s*(\d+)/i);
        if (m) {
          const num = parseInt(m[1]);
          if (!numsInText.includes(num)) numsInText.push(num);
        }
      }
      if (numsInText.length === uniqueSegs.length) {
        numsInText.forEach((num, idx) => { segNumToName[num] = uniqueSegs[idx]; });
      } else {
        uniqueSegs.forEach((name, idx) => { segNumToName[idx + 1] = name; });
      }
    }
  }

  const lines = parts.map(part => {
    const m = part.match(/^Segment\s*(\d+)\s*:\s*(.+)$/i);
    if (!m) return part;
    const num = parseInt(m[1]);
    const value = m[2].trim();
    const name = segNumToName[num];
    return name ? `Segment ${num} (${name}): ${value}` : `Segment ${num}: ${value}`;
  });

  let result = 'Quote pro Segment:\n' + lines.join('\n');
  if (tail) result += '\n\n' + tail;
  return result;
}

// v12.6: Standort -> Firma Mapping. Die Frontend-Cowork-Form sendet aktuell
// pauschal unternehmen='F&T' fuer alle Gruppen. Da unsere Studio-Locations
// fest auf eine Firma gemappt sind, ueberschreiben wir das hier basierend
// auf gruppe.standort. Bei unbekanntem Standort bleibt gruppe.unternehmen
// erhalten (Fallback).
const STANDORT_FIRMA_MAPPING = {
  'F&T': ['Bochum', 'Düsseldorf', 'Mannheim', 'Hannover', 'Hamburg (Spitalerstrasse)'],
  'm-s': ['Berlin', 'Köln', 'Nürnberg', 'Stuttgart', 'Hamburg (Mönckebergstrasse)'],
  'H+G': ['Essen', 'München', 'Leipzig', 'Frankfurt (Roßmarkt)', 'Frankfurt (Holzgraben)'],
};

function resolveUnternehmenFromStandort(standort, fallback) {
  if (!standort || typeof standort !== 'string') return fallback || 'F&T';
  const s = standort.toLowerCase().trim();
  // Tolerant: Sonderzeichen entfernen, damit Mönkebergstr/Mönckebergstr etc. matchen
  const norm = s.replace(/[ßẞ]/g, 'ss').replace(/[^a-zäöü0-9 ()]/g, '');
  for (const [firma, locations] of Object.entries(STANDORT_FIRMA_MAPPING)) {
    for (const loc of locations) {
      const locNorm = loc.toLowerCase().replace(/[ßẞ]/g, 'ss').replace(/[^a-zäöü0-9 ()]/g, '');
      // Substring-Match in beide Richtungen (toleriert "Frankfurt" alleine, ebenso ausfuehrliche Schreibweise)
      if (norm === locNorm) return firma;
      if (norm.includes(locNorm) || locNorm.includes(norm)) return firma;
      // Tippfehler-Toleranz: auch matchen wenn nach Streichen von 'c' identisch
      // (fängt "Mönckeberg" vs "Mönkeberg" ab). Nur für längere Strings, damit
      // kein Random-Match passiert.
      if (norm.length >= 12 && locNorm.length >= 12) {
        const normSimple = norm.replace(/c/g, '');
        const locSimple = locNorm.replace(/c/g, '');
        if (normSimple.includes(locSimple) || locSimple.includes(normSimple)) return firma;
      }
    }
  }
  // Spezialfall Hamburg: ohne Strassen-Suffix -> nicht eindeutig zuordbar
  // -> bei nacktem 'hamburg' fallback nehmen
  if (/^hamburg\s*$/.test(norm)) return fallback || 'F&T';
  // Spezialfall Frankfurt: ohne Suffix -> H+G (alle Frankfurter Studios sind H+G)
  if (/^frankfurt\s*$/.test(norm)) return 'H+G';
  return fallback || 'F&T';
}

function fillSheet(ws, gruppe, fragen, projektnummer, projektname, kundenname, setting, methode, quoteStartCol, allGruppen, options) {
  const opts = options || {};

  // v12.6: Unternehmen aus Standort ableiten (Frontend sendet pauschal 'F&T')
  const resolvedFirma = resolveUnternehmenFromStandort(gruppe.standort, gruppe.unternehmen);
  if (resolvedFirma !== gruppe.unternehmen) {
    gruppe.unternehmen = resolvedFirma;
  }
  const compactMatrix = opts.compactMatrix === true;
  const compactThreshold = opts.compactMatrixThreshold || 5;
  const baseHmap = HEADER_MAP[setting] || HEADER_MAP.offline;
  // Online: dynamisch erkannte Positionen überschreiben die statische Map.
  // Offline: bleibt bei der statischen Map (Studio-Logik mit eigener Struktur).
  const hmap = (setting === 'online')
    ? { ...baseHmap, ...detectHeaderPositions(ws) }
    : baseHmap;
  const brutto = gruppe.brutto || 8;
  const netto = gruppe.netto || 6;
  const tnEnd = TN_START + brutto - 1;
  const labelRow = tnEnd + QUOTE_HINT_AFTER_TN_OFFSET;
  // lfd. Nr. Spalte dynamisch erkennen — offline = E (5), online = F (6)
  const lfdCol = findLfdNrCol(ws, setting === 'online' ? 6 : 5);

  // 1) Vorbefüllte Tracking-Werte aus Template leeren
  clearPrefilledTNCells(ws, tnEnd, quoteStartCol);

  // 2) Header befüllen
  // v12.5: Studio-Label und Projekt-Nr.-Label je Unternehmen anpassen
  // (F&T / m-s / H+G) — alle Templates basieren auf dem F&T-Layout, wir
  // überschreiben nur die zwei Studio-/Projekt-Nr.-Labels und das Logo.
  const UNT = gruppe.unternehmen || 'F&T';
  const studioPrefix = (UNT === 'm-s') ? 'm-s' : (UNT === 'H+G' ? 'H+G' : 'F&T');
  if (hmap.studioLabel) {
    ws.getCell(hmap.studioLabel).value = `${studioPrefix} Standort`;
  }
  if (hmap.studioWert && gruppe.standort) {
    ws.getCell(hmap.studioWert).value = gruppe.standort;
  }
  // v12.5: Termin-Header sauber aus datum + uhrzeit bauen.
  // gruppe.termin kommt vom Parser oft als "KW21, Mai 2026" oder Range
  // ("Di. 19.05. – Do. 21.05.") — das ist für GD-Sheets (1 Termin pro Sheet)
  // nicht das, was hier hin soll. Wir bevorzugen pro Sheet den exakten
  // Termin "Datum Uhrzeit".
  let terminText = '';
  if (gruppe.datum && gruppe.uhrzeit) {
    terminText = `${gruppe.datum} ${gruppe.uhrzeit}`.trim();
  } else if (gruppe.termin && !/KW\s*\d|Januar|Februar|März|April|Mai|Juni|Juli|August|September|Oktober|November|Dezember/i.test(gruppe.termin)) {
    terminText = gruppe.termin;
  } else {
    terminText = `${gruppe.datum || ''} ${gruppe.uhrzeit || ''}`.trim();
  }
  ws.getCell(hmap.terminWert).value     = terminText;
  ws.getCell(hmap.kunde).value          = kundenname || '';
  // PATCH 5: Zielgruppe + Zuordnungs-Kriterien als Tooltip am Zielgruppe-Feld
  ws.getCell(hmap.zielgruppe).value     = gruppe.zielgruppe || '';
  if (gruppe.zuordnungs_kriterien) {
    ws.getCell(hmap.zielgruppe).note = `Zuordnungs-Kriterien:\n${gruppe.zuordnungs_kriterien}`;
  }
  ws.getCell(hmap.projekt).value        = projektname;
  // v12.5: Projekt-Nr.-Label je Unternehmen überschreiben
  // hmap.projNr ist die Wert-Zelle; das Label sitzt 1 Spalte links davon
  ws.getCell(hmap.projNr).value         = projektnummer;
  try {
    const projNrMatch = String(hmap.projNr).match(/^([A-Z]+)(\d+)$/);
    if (projNrMatch) {
      const valColLetter = projNrMatch[1];
      const valRow = parseInt(projNrMatch[2]);
      // Spalten-Index aus Buchstabe ermitteln
      let valColIdx = 0;
      for (const c of valColLetter) valColIdx = valColIdx * 26 + (c.charCodeAt(0) - 64);
      if (valColIdx > 1) {
        const labelCell = ws.getCell(valRow, valColIdx - 1);
        // Nur überschreiben, wenn aktuell ein Projekt-Nr-artiges Label drinsteht
        const curLabel = String(labelCell.value || '').trim();
        if (/projekt[-\s]?nr/i.test(curLabel)) {
          labelCell.value = `${studioPrefix} Projekt-Nr.`;
        }
      }
    }
  } catch (e) { /* fail-soft */ }
  ws.getCell(hmap.incentive).value      = gruppe.incentive || '';

  // Header-Werte links-bündig ausrichten (vereinheitlicht für alle Templates)
  const headerValueCells = [
    hmap.terminWert, hmap.kunde, hmap.zielgruppe,
    hmap.projekt, hmap.projNr, hmap.incentive,
  ].filter(Boolean);
  for (const addr of headerValueCells) {
    const c = ws.getCell(addr);
    c.alignment = { horizontal: 'left', vertical: 'middle', wrapText: true, indent: 1 };
  }
  if (hmap.studioWert) {
    ws.getCell(hmap.studioWert).alignment = { horizontal: 'left', vertical: 'middle', wrapText: true, indent: 1 };
  }

  // 3) lfd. Nr. neu setzen + "X für Y"-Label
  // Letzte TN-Zeile auch in lfd-Nr-Spalte mit dicker Bottom-Border abschließen
  const lfdHeader = ws.getCell(HEADER_ROW, lfdCol);
  lfdHeader.note = `Brutto: ${brutto} TN\nNetto: ${netto} TN\n→ ${brutto} für ${netto}`;
  const blackBottom = { style: 'medium', color: { argb: 'FF000000' } };

  // v11.4: Wenn brutto > Anzahl vorformatierter TN-Zeilen im Template, müssen
  // wir das Format der letzten formatierten Zeile auf alle weiteren propagieren
  // (Höhe, Borders, Fills, Fonts pro Spalte)
  //
  // Detektion: alle TN-Zeilen mit row.height >= 20 zählen als formatiert
  let lastFormattedTnRow = TN_START;
  for (let r = TN_START; r <= TN_START + 30; r++) {
    const h = ws.getRow(r).height;
    if (h && h >= 20) {
      lastFormattedTnRow = r;
    } else if (r > TN_START && (!h || h < 20)) {
      break; // erste unformatierte Zeile gefunden, abbrechen
    }
  }
  // Wenn brutto mehr Zeilen braucht als vorformatiert → kopiere Style der
  // letzten formatierten Zeile auf alle neuen Zeilen
  if (tnEnd > lastFormattedTnRow) {
    const templateRow = ws.getRow(lastFormattedTnRow);
    const templateHeight = templateRow.height;
    // Höchste Spalte mit Style/Inhalt in der Vorlagen-Zeile finden
    let maxCol = quoteStartCol; // mindestens bis Frage-Bereich
    templateRow.eachCell({ includeEmpty: true }, (cell, colNum) => {
      if (cell.font || cell.fill?.fgColor || cell.border?.top || cell.value !== null) {
        if (colNum > maxCol) maxCol = colNum;
      }
    });
    // Pro Spalte den Style der letzten formatierten Zelle merken
    const colStyles = new Map();
    for (let col = 1; col <= maxCol; col++) {
      const tcell = ws.getCell(lastFormattedTnRow, col);
      colStyles.set(col, {
        font: tcell.font ? { ...tcell.font } : undefined,
        fill: tcell.fill ? JSON.parse(JSON.stringify(tcell.fill)) : undefined,
        border: tcell.border ? JSON.parse(JSON.stringify(tcell.border)) : undefined,
        alignment: tcell.alignment ? { ...tcell.alignment } : undefined,
        numFmt: tcell.numFmt,
      });
    }
    // Auf neue Zeilen anwenden
    for (let r = lastFormattedTnRow + 1; r <= tnEnd; r++) {
      const newRow = ws.getRow(r);
      if (templateHeight) newRow.height = templateHeight;
      for (let col = 1; col <= maxCol; col++) {
        const cell = ws.getCell(r, col);
        const s = colStyles.get(col);
        if (!s) continue;
        if (s.font)      cell.font      = { ...s.font };
        if (s.fill)      cell.fill      = JSON.parse(JSON.stringify(s.fill));
        if (s.border)    cell.border    = JSON.parse(JSON.stringify(s.border));
        if (s.alignment) cell.alignment = { ...s.alignment };
        if (s.numFmt)    cell.numFmt    = s.numFmt;
      }
    }

    // v11.6: Merges und Data Validations aus dem Template-TN-Bereich extrahieren
    // und für die neuen Zeilen replizieren.
    //
    // (a) Merges: alle Merges im Bereich TN_START..lastFormattedTnRow analysieren
    //     und je Zeile-Block (z.B. G7:H7) für jede neue Zeile ergänzen.
    //     Merge-Pattern (G7:H7) wird repliziert auf G17:H17, G18:H18, ...
    const templateMergePatterns = [];
    // Quelle 1: preMerges (vom XML-Parser, zuverlässig)
    const preMergesList = Array.isArray(opts.preMerges) ? opts.preMerges : [];
    // Quelle 2: model.merges
    const modelMergesList = Array.isArray(ws.model?.merges) ? ws.model.merges : [];
    // Quelle 3: _merges-Object
    const oldMergesList = (ws._merges && typeof ws._merges === 'object' && !(ws._merges instanceof Map))
      ? Object.keys(ws._merges) : [];
    const allMergeStrings = [...new Set([...preMergesList, ...modelMergesList, ...oldMergesList])];
    for (const m of allMergeStrings) {
      if (typeof m !== 'string') continue;
      const match = m.match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/i);
      if (!match) continue;
      const [_, c1, r1str, c2, r2str] = match;
      const r1 = parseInt(r1str), r2 = parseInt(r2str);
      // Nur einzeilige Merges innerhalb des TN-Bereichs sind Kandidaten
      if (r1 === r2 && r1 >= TN_START && r1 <= lastFormattedTnRow) {
        templateMergePatterns.push({ c1: c1.toUpperCase(), c2: c2.toUpperCase(), srcRow: r1 });
      }
    }
    // Pattern: für jede Spalten-Kombo, die in genau EINER Template-Zeile vorkommt,
    // auf alle neuen Zeilen anwenden. Dedup-Set über Spalten-Kombo.
    const uniqueMergePatterns = new Map(); // key "c1-c2" → pattern
    for (const p of templateMergePatterns) {
      uniqueMergePatterns.set(`${p.c1}-${p.c2}`, p);
    }
    for (const p of uniqueMergePatterns.values()) {
      for (let r = lastFormattedTnRow + 1; r <= tnEnd; r++) {
        try { ws.mergeCells(`${p.c1}${r}:${p.c2}${r}`); } catch (e) {}
      }
    }

    // (b) Data Validations: für Spalten, die im Template eine DV haben,
    //     diese auf die neuen Zeilen erweitern.
    //     ExcelJS speichert DVs als Map: address-string → DV-Objekt.
    //     Wir suchen DVs auf Zellen wie "A7", "A10" etc. im Template-Bereich
    //     und kopieren sie auf die neuen Zeilen mit gleicher Spalte.
    const dvModel = ws.dataValidations?.model || ws.dataValidations || {};
    const dvEntries = (typeof dvModel === 'object' && !Array.isArray(dvModel))
      ? Object.entries(dvModel)
      : [];
    // Welche Spalten haben im Template eine DV?
    const dvByCol = new Map(); // col-letter → DV-Objekt
    for (const [sqref, dv] of dvEntries) {
      if (!sqref || !dv) continue;
      // sqref kann sein: "A7", "A7:A16", "A7 B7" oder mehrere Bereiche.
      // Wir splitten an Leerzeichen und Komma.
      const refs = String(sqref).split(/[\s,]+/);
      for (const ref of refs) {
        const m = ref.match(/^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/i);
        if (!m) continue;
        const c1 = m[1].toUpperCase();
        const r1 = parseInt(m[2]);
        const c2 = (m[3] || c1).toUpperCase();
        const r2 = m[4] ? parseInt(m[4]) : r1;
        // Nur einspaltige DVs im TN-Bereich sind Kandidaten
        if (c1 === c2 && r1 >= TN_START && r2 <= lastFormattedTnRow) {
          dvByCol.set(c1, dv);
        }
      }
    }
    // DVs auf neue Zeilen anwenden
    for (const [colLetter, dv] of dvByCol.entries()) {
      let dvClone;
      try { dvClone = JSON.parse(JSON.stringify(dv)); } catch (e) { dvClone = { ...dv }; }
      for (let r = lastFormattedTnRow + 1; r <= tnEnd; r++) {
        try {
          const addr = `${colLetter}${r}`;
          if (typeof ws.dataValidations?.add === 'function') {
            ws.dataValidations.add(addr, dvClone);
          } else if (ws.dataValidations?.model) {
            ws.dataValidations.model[addr] = dvClone;
          }
        } catch (e) {}
      }
    }
  }

  // v11.6: Vertikales Alignment ALLER TN-Zellen final auf 'top' setzen
  // (auch nach allen Style-Propagationen). Erstreckt sich über alle Zeilen
  // und alle Spalten links der Quote-Spalten.
  for (let r = TN_START; r <= tnEnd; r++) {
    for (let col = 1; col < quoteStartCol; col++) {
      const c = ws.getCell(r, col);
      const a = c.alignment || {};
      c.alignment = {
        ...a,
        vertical: 'middle',
      };
    }
  }

  for (let r = TN_START; r <= tnEnd; r++) {
    const c = ws.getCell(r, lfdCol);
    c.value = r - TN_START + 1;
    c.alignment = { horizontal: 'center', vertical: 'middle' };
    c.font = { name: 'Arial', size: 10, bold: true };
    if (r === tnEnd) {
      // Defensive: explizites Border-Objekt statt Spread auf ExcelJS-Property
      const eb = c.border || {};
      c.border = {
        top: eb.top, left: eb.left, right: eb.right,
        bottom: blackBottom,
      };
    }
  }
  // Auch Tracking-Spalten links der lfd. Nr. mit dicker Abschlusslinie versehen
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

  // v11.4: Termin-Verteilung pro TN-Zeile (nur bei IDI mit termin_blocks)
  // Beispiel: termin_blocks = [{tn_von:1, tn_bis:5, datum:'5.5.', uhrzeit:'10-15h'}, ...]
  // → TN-Zeilen 1-5 (= rows TN_START..TN_START+4) bekommen Datum 5.5.
  if (Array.isArray(opts.termin_blocks) && opts.termin_blocks.length > 0) {
    const cols = findDatumUhrzeitCols(ws);
    if (cols.datum || cols.uhrzeit) {
      for (const block of opts.termin_blocks) {
        const von = parseInt(block.tn_von) || 1;
        const bis = parseInt(block.tn_bis) || von;
        for (let tn = von; tn <= bis; tn++) {
          const row = TN_START + tn - 1;
          if (row > tnEnd) break;
          if (cols.datum && block.datum) {
            const c = ws.getCell(row, cols.datum);
            c.value = String(block.datum);
            c.alignment = { horizontal: 'center', vertical: 'middle' };
            c.font = { name: 'Arial', size: 10 };
          }
          if (cols.uhrzeit && block.uhrzeit) {
            const c = ws.getCell(row, cols.uhrzeit);
            c.value = String(block.uhrzeit);
            c.alignment = { horizontal: 'center', vertical: 'middle' };
            c.font = { name: 'Arial', size: 10 };
          }
        }
      }
    }
  }

  // 4) Tracking-Bereich rechts und unter TN ausnullen
  clearTrackingArea(ws, tnEnd, quoteStartCol);

  // 5) Quote-N-Reste in Header-Zeile leeren
  for (let col = quoteStartCol; col <= quoteStartCol + 30; col++) {
    const c = ws.getCell(HEADER_ROW, col);
    if (c.value && /^Quote/i.test(String(c.value))) {
      c.value = null;
    }
  }

  // v12.4: Phantom-Spalten links der quoteStartCol leeren
  // (z.B. T6='TB', U6='PaySite' aus alten Template-Versionen, die nicht zu den
  // offiziellen Tracking-Headern gehören). Konservativ: nur bekannte Phantom-Labels.
  const PHANTOM_HEADERS = ['tb', 'paysite', 'pay site', 'pay-site', 'paysite '];
  for (let col = 1; col < quoteStartCol; col++) {
    for (const row of [MATRIX_HEADER_ROW, HEADER_ROW]) {
      const c = ws.getCell(row, col);
      if (!c.value) continue;
      const v = String(c.value).toLowerCase().trim();
      if (PHANTOM_HEADERS.includes(v)) {
        c.value = null;
      }
    }
  }

  // 6) Fragen-Spalten aufbauen
  //
  // v11.4: Soziodemographische Fragen sortieren — Geschlecht + Alter immer
  // zuerst, dann andere persönliche Daten, dann themenbezogen.
  const sortedFragen = sortFragenSoziodemoFirst(fragen || []);

  // v11.3: Vorab Max-Antwort-Anzahl ermitteln, damit alle grünen Quote-Zellen
  // auf einer einheitlichen Höhe stehen (1 Zeile unter der längsten Antwort-Liste)
  let maxAnswers = 0;
  for (const f of sortedFragen) {
    if (f.typ === 'entfaellt' || f.kategorie === 'entfaellt' || f.kategorie === 'verfuegbarkeit') continue;
    // Filter: relevantFuerGruppen
    if (Array.isArray(f.relevantFuerGruppen) && f.relevantFuerGruppen.length
        && !f.relevantFuerGruppen.includes('alle')
        && !f.relevantFuerGruppen.includes(gruppe.id)) continue;
    if (Array.isArray(f.antworten)) {
      maxAnswers = Math.max(maxAnswers, f.antworten.length);
    }
    if (Array.isArray(f.items)) {
      for (const it of f.items) {
        const itemAnts = (it.antworten && it.antworten.length) ? it.antworten : (f.antworten || []);
        maxAnswers = Math.max(maxAnswers, itemAnts.length);
      }
    }
  }
  // Quote-Zeile = 1 Leerzeile nach der längsten Antwort-Liste
  const uniformQuoteRow = tnEnd + ANT_OFFSET + maxAnswers + 1;

  let currentCol = quoteStartCol;
  for (const frage of sortedFragen) {
    // Sicherheitsnetz: Fragen mit kategorie='entfaellt' oder 'verfuegbarkeit'
    // werden übersprungen, auch wenn typ noch single_choice ist.
    if (frage.typ === 'entfaellt' || 
        frage.kategorie === 'entfaellt' || 
        frage.kategorie === 'verfuegbarkeit') {
      continue;
    }
    // v12.6: Auto-Fix - wenn typ='numerisch' oder 'freitext' aber antworten[] mit
    // Codes UND Texten gefuellt ist, ist es eigentlich single_choice (typischer
    // Parser-Fehler bei Fragen wie 'Alter: ___' mit Optionen 1-3 darunter).
    if ((frage.typ === 'numerisch' || frage.typ === 'freitext') &&
        Array.isArray(frage.antworten) && frage.antworten.length > 0 &&
        frage.antworten[0].code && frage.antworten[0].text) {
      frage.typ = 'single_choice';
    }
    // Wenn typ='single_choice' aber items[] vorhanden ist (Parser-Inkonsistenz),
    // behandle als Matrix.
    const istEffektivMatrix = frage.typ === 'matrix' || 
      (Array.isArray(frage.items) && frage.items.length > 0 && 
       frage.items[0].antworten && frage.items[0].antworten.length > 0);
    if (istEffektivMatrix && Array.isArray(frage.items) && frage.items.length) {
      // Matrix: pro Item eine Spalte + Mutter-Header in Zeile 5 gemerged
      // Quote-Text fällt zurück auf bedingung (Parser legt Quoten-Logik dort ab)
      const matrixQuoteText = frage.quotenkommentar || frage.bedingung || '';
      const matrixStart = currentCol;
      // Fallback-Antworten von Frage-Ebene (z.B. Skala 1=sehr gern...6=kenne ich nicht)
      // Bei Matrizen liegen Antworten oft auf Frage-Ebene, nicht pro Item dupliziert
      const sharedAnswers = Array.isArray(frage.antworten) ? frage.antworten : [];
      // Erkennt ob der matrixQuoteText nur EIN bestimmtes Item meint (z.B. "Jörg Pilawa darf nicht...")
      // damit der Hinweis nicht generisch in alle Item-Spalten geschrieben wird
      const matrixQuoteMentionsItem = (itemLabel) => {
        if (!matrixQuoteText || !itemLabel) return false;
        const il = itemLabel.toLowerCase().trim();
        return matrixQuoteText.toLowerCase().includes(il);
      };

      // PATCH 1: Compact-Matrix — bei vielen Items quotenrelevante einzeln,
      // andere als Sammelspalte. Per Toggle aktivierbar.
      const partitioned = partitionMatrixItems(frage.items, compactMatrix, compactThreshold);
      const itemsToColumns = partitioned.quotaItems;
      const otherItems = partitioned.otherItems;

      const totalCols = itemsToColumns.length + (otherItems.length > 0 ? 1 : 0);
      let writtenInMatrix = 0;

      for (const item of itemsToColumns) {
        writtenInMatrix++;
        const isLast = (writtenInMatrix === totalCols);
        const itemNote = `MUTTERFRAGE: ${frage.fragetext || ''}\n\n${matrixQuoteText}\n\n${item.note || ''}\n\n${item.marker ? 'Marker: ' + item.marker : ''}`.trim();
        // Quote-Text-Logik mit Auto-Fallback aus Marker:
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
        // Item-Soll-Quote (v11: Liste mit gilt_fuer_gruppen-Filter)
        const itemQuotes = getSollQuoteList(item.soll_quote);
        for (const q of itemQuotes) {
          if (Array.isArray(q.gilt_fuer_gruppen) && q.gilt_fuer_gruppen.length > 0) {
            if (!q.gilt_fuer_gruppen.includes(gruppe.id)) continue;
          }
          itemQuote = (itemQuote ? itemQuote + '\n' : '') + `• ${q.text}`;
        }
        // Antworten pro Item: nutze item.antworten, fallback auf frage.antworten (Matrix-Skala)
        const rawAnts = (item.antworten && item.antworten.length) ? item.antworten : sharedAnswers;
        const ants = rawAnts.map(a => ({
          ...a,
          screenout: a.screenout === true ? true :
                     (item.screenout_codes || []).map(String).includes(String(a.code)),
        }));
        itemQuote = formatSegmentQuotes(itemQuote, opts.idiProfile);
        writeQuestionColumn(ws, currentCol, item.item_label || '', itemNote,
                            ants, itemQuote, tnEnd, gruppe, frage, isLast, allGruppen, uniformQuoteRow);
        currentCol++;
      }

      // PATCH 1: Sammelspalte für nicht-quotenrelevante Items
      if (otherItems.length > 0) {
        writtenInMatrix++;
        const summaryLabel = `Weitere Items (${otherItems.length})`;
        const summaryNote = `MUTTERFRAGE: ${frage.fragetext || ''}\n\nNicht quotenrelevante Items (zur Info, ohne Quote):\n\n` +
          otherItems.map(it => `• ${it.item_label || ''}`).join('\n') + 
          `\n\nDiese Items haben keine Screenouts/Marker — Antworten werden in einer Sammelspalte erfasst.`;
        // Antwort-Skala: gemeinsame Frage-Antworten verwenden, falls vorhanden
        const summaryAnts = sharedAnswers;
        // Quote-Text: Hinweis auf Sammelspalte
        const summaryQuote = `Sammelspalte: ${otherItems.length} weitere Items (siehe Tooltip)\nKein Screenout, kein Quoten-Marker.`;
        const summaryQuoteFmt = formatSegmentQuotes(summaryQuote, opts.idiProfile);
        writeQuestionColumn(ws, currentCol, summaryLabel, summaryNote,
                            summaryAnts, summaryQuoteFmt, tnEnd, gruppe, frage, true, allGruppen, uniformQuoteRow);
        // Sammelspalte etwas breiter machen
        ws.getColumn(currentCol).width = 28;
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
      // Mutter-Header zeigt zusätzlich Hinweis bei Compact-Mode
      const compactHint = (compactMatrix && otherItems.length > 0)
        ? ` [kompakt: ${itemsToColumns.length}/${frage.items.length} relevante Items]`
        : '';
      mh.value = (frage.kurz_label || frage.fragetext || frage.id) + compactHint;
      mh.font = { name: 'Arial', size: 10, bold: true };
      mh.alignment = { horizontal: 'center', vertical: 'center', wrapText: true };
      mh.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.MATRIX_HDR } };
      mh.border = borderWithThickRight(HEADER_BORDER);
      if (matrixQuoteText) mh.note = matrixQuoteText;
    } else if (frage.typ === 'freitext' || frage.typ === 'numerisch') {
      // Freitext / Numerisch: nur Header, keine Antwort-Codes
      const note = frage.bedingung ? `Bedingung: ${frage.bedingung}` : (frage.fragetext || '');
      const label = buildHeaderLabel(frage);
      writeQuestionColumn(ws, currentCol, label, note, null, null, tnEnd, gruppe, frage, true, allGruppen, uniformQuoteRow);
      currentCol++;
    } else if (frage.typ === 'tool_input') {
      // v11: Algorithmus-Eingabe-Skala (z.B. JPM Q12a) — 1 Sammelspalte
      const label = buildHeaderLabel(frage);
      const itemList = (frage.items || []).map(it => `• ${it.item_label || ''}`).join('\n');
      const note = `MUTTERFRAGE: ${frage.fragetext || ''}\n\n` +
                   `Algorithmus-Eingabe — Antworten in das externe Segmentierungs-Tool eingeben.\n\n` +
                   (itemList ? `Items:\n${itemList}\n\n` : '') +
                   (frage.quotenkommentar || '');
      const quoteText = frage.quotenkommentar
        || 'Eingabe in Segmentierungs-Tool — Ergebnis: Segment';
      const tiQuoteFmt = formatSegmentQuotes(quoteText, opts.idiProfile);
      writeQuestionColumn(ws, currentCol, label + ' (Tool)', note,
                          [], tiQuoteFmt, tnEnd, gruppe, frage, true, allGruppen, uniformQuoteRow);
      ws.getColumn(currentCol).width = 30;
      currentCol++;
    } else {
      // single_choice, multi_choice, ranking
      const label = buildHeaderLabel(frage);
      const note = frage.bedingung ? `Bedingung: ${frage.bedingung}` : '';
      let quoteText = frage.quotenkommentar || frage.bedingung || '';
      // Auto-Quote-Hinweis für F1 (Geschlecht) und F2 (Alter) aus Gruppen-Daten
      const ftLower = (frage.fragetext || '').toLowerCase();
      if (!quoteText && ftLower.includes('geschlecht') && gruppe.geschlecht && gruppe.geschlecht !== 'gemischt') {
        quoteText = `Quote: nur ${gruppe.geschlecht}`;
      }
      // v12.2: Alter-Spalte — IMMER Segment-Liste versuchen (auch ergänzend zu
      // einem ggf. vorhandenen quotenkommentar). Zwei Quellen mit Fallback:
      //   1) idiProfile[i].alter_min/alter_max (strukturierte Felder vom Parser)
      //   2) segment_beschreibungen[name] via Regex (Freitext-Fallback —
      //      greift wenn Sonnet die strukturierten Felder vergisst)
      if (ftLower.includes('alt sind') || ftLower.includes('alter')) {
        const idiProfile = Array.isArray(opts.idiProfile) ? opts.idiProfile : [];
        const segBeschr = (opts.segment_beschreibungen && typeof opts.segment_beschreibungen === 'object')
          ? opts.segment_beschreibungen : null;
        const segAge = {};
        // Quelle 1: idiProfile mit alter_min/alter_max
        for (const p of idiProfile) {
          if (p.alter_min != null || p.alter_max != null) {
            const segKey = p.segment || p.id || '?';
            if (!segAge[segKey]) segAge[segKey] = [p.alter_min, p.alter_max];
          }
        }
        // Quelle 2: Regex auf segment_beschreibungen (Fallback)
        if (Object.keys(segAge).length === 0 && segBeschr) {
          const segNames = idiProfile.length > 0
            ? [...new Set(idiProfile.map(p => p.segment).filter(Boolean))]
            : Object.keys(segBeschr);
          for (const seg of segNames) {
            const desc = segBeschr[seg];
            if (!desc) continue;
            const ds = String(desc);
            // Range: "18-50", "18–50", "18 bis 50"
            let m = ds.match(/(\d{2})\s*[-–]\s*(\d{2})/);
            if (!m) m = ds.match(/(\d{2})\s*bis\s*(\d{2})/i);
            if (m) { segAge[seg] = [parseInt(m[1]), parseInt(m[2])]; continue; }
            // Open-ended: "55+", "ab 55", "über 50"
            m = ds.match(/(\d{2})\s*\+/);
            if (!m) m = ds.match(/(?:ab|[üu]ber)\s*(\d{2})/i);
            if (m) segAge[seg] = [parseInt(m[1]), null];
          }
        }
        // Segment-Nr aus idiProfile bestimmen (explizit segment_nr oder Auftreten-Reihenfolge)
        const segNrMap = {};
        for (const p of idiProfile) {
          if (p.segment && p.segment_nr != null && !segNrMap[p.segment]) {
            segNrMap[p.segment] = parseInt(p.segment_nr);
          }
        }
        if (Object.keys(segNrMap).length === 0) {
          const uniqSegs = [];
          for (const p of idiProfile) {
            if (p.segment && !uniqSegs.includes(p.segment)) uniqSegs.push(p.segment);
          }
          uniqSegs.forEach((s, idx) => { segNrMap[s] = idx + 1; });
        }
        const lines = Object.entries(segAge)
          .map(([seg, [mn, mx]]) => {
            const range = (mx != null) ? `${mn}-${mx}` : (mn != null ? `${mn}+` : '?');
            const nr = segNrMap[seg];
            const label = nr != null ? `Segment ${nr} (${seg})` : seg;
            return { nr: nr || 999, text: `${label}: ${range}` };
          })
          .sort((a, b) => a.nr - b.nr)
          .map(x => x.text);
        if (lines.length > 0) {
          const segText = 'Quote pro Segment:\n' + lines.join('\n');
          quoteText = quoteText ? `${quoteText}\n\n${segText}` : segText;
        } else if (!quoteText && gruppe.alter_min && gruppe.alter_max) {
          quoteText = `Quote: ${gruppe.alter_min}-${gruppe.alter_max} Jahre`;
        }
      }
      quoteText = formatSegmentQuotes(quoteText, opts.idiProfile);
      writeQuestionColumn(ws, currentCol, label, note,
                          frage.antworten, quoteText, tnEnd, gruppe, frage, true, allGruppen, uniformQuoteRow);
      currentCol++;
    }
  }

  // Höhere Zeilen für Matrix-Header und Frage-Header
  ws.getRow(MATRIX_HEADER_ROW).height = 32;
  ws.getRow(HEADER_ROW).height = 75;

  // v11: IDI-Profile, Segment-Beschreibungen und Studien-Quoten in Spalte G ausspielen
  // (wenn vorhanden in opts)
  if (opts.idiProfile || opts.segment_beschreibungen || opts.studien_quoten) {
    writeIdiInfoBlock(ws, {
      idiProfile: opts.idiProfile,
      segment_beschreibungen: opts.segment_beschreibungen,
      studien_quoten: opts.studien_quoten,
      tnEnd,
    });
  }

  // Freeze Panes komplett deaktiviert — User scrollt frei in alle Richtungen
  ws.views = [{ state: 'normal' }];
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
      compactMatrix,            // PATCH 1: optionaler Toggle aus dem Form
      compactMatrixThreshold,   // optional: Item-Schwellwert (Default 5)
      // v11: IDI-Profile + Segment-Beschreibungen + Studien-Quoten
      idiProfile,
      segment_beschreibungen,
      studien_quoten,
      // v11.4: Termin-Blöcke (für Datum/Uhrzeit pro TN-Zeile bei IDI)
      termin_blocks,
      laufzeitVon,
      laufzeitBis,
    } = req.body ?? {};

    const auftraggeber = kundenname || projektname || '';
    const builderOptions = {
      compactMatrix: compactMatrix === true || compactMatrix === 'true',
      compactMatrixThreshold: parseInt(compactMatrixThreshold) || 5,
      // v11: nur in IDI-Sheets durchreichen
      idiProfile: Array.isArray(idiProfile) ? idiProfile : null,
      segment_beschreibungen: (segment_beschreibungen && typeof segment_beschreibungen === 'object')
        ? segment_beschreibungen : null,
      studien_quoten: Array.isArray(studien_quoten) ? studien_quoten : null,
      // v11.4: Termin-Blöcke für TN-Zeilen-Verteilung
      termin_blocks: Array.isArray(termin_blocks) ? termin_blocks : null,
      laufzeitVon: laufzeitVon || '',
      laufzeitBis: laufzeitBis || '',
    };

    if (!templateBase64) return res.status(400).json({ error: 'Missing templateBase64' });
    if (!gruppen?.length) return res.status(400).json({ error: 'Missing gruppen' });

    const fragenArr = Array.isArray(fragen) ? fragen : [];
    const setting = (methode === 'VGD' || methode === 'VDI') ? 'online' : 'offline';
    const cfg = SHEET_CONFIG[methode] || SHEET_CONFIG['GD'];

    const templateBuffer = Buffer.from(templateBase64, 'base64');
    const template = new ExcelJS.Workbook();
    await template.xlsx.load(templateBuffer);
    const srcWs = template.getWorksheet(cfg.sheetName);
    if (!srcWs) {
      return res.status(400).json({
        error: `Sheet '${cfg.sheetName}' not found`,
        available: template.worksheets.map(w => w.name),
      });
    }

    // v11.3: Merges aus rohem XLSX-Buffer parsen (umgeht ExcelJS-API-Quirks)
    const preMerges = await readMergesFromBuffer(templateBuffer, cfg.sheetName);
    // v11.6: preMerges auch fillSheet zugänglich machen, damit es die TN-Bereich-
    // Merges (z.B. G7:H7) auf neue Zeilen 11+ replizieren kann.
    builderOptions.preMerges = preMerges;

    const result = new ExcelJS.Workbook();
    result.creator = 'Preview Generator';
    result.created = new Date();
    const newSheetNames = [];

    const processGruppe = (gruppe, sheetName, fragenFG, includeIdiInfo) => {
      const dstWs = result.addWorksheet(sheetName);
      copyWorksheet(srcWs, dstWs, preMerges);

      const methodeFG = gruppe.methode || methode;
      const cfgFG = SHEET_CONFIG[methodeFG] || SHEET_CONFIG['GD'];
      const quoteStartCol = findQuoteStartCol(dstWs, cfgFG.quoteStartCol);
      const brutto = gruppe.brutto || 8;
      const tnEnd = TN_START + brutto - 1;

      addConditionalFormats(dstWs, tnEnd);

      // v12.6: Unternehmen anhand Standort ueberschreiben, BEVOR Logo gesetzt wird
      gruppe.unternehmen = resolveUnternehmenFromStandort(gruppe.standort, gruppe.unternehmen);
      addLogo(dstWs, result, gruppe.unternehmen, setting);

      // v11: nur dem IDI/VDI-Sheet die IDI-Info-Blöcke geben
      // v11.4: termin_blocks ebenfalls nur bei IDI (bei GD pro Sheet 1 Termin im Header)
      const sheetOptions = includeIdiInfo
        ? builderOptions
        : { ...builderOptions, idiProfile: null, segment_beschreibungen: null,
            studien_quoten: null, termin_blocks: null };

      fillSheet(
        dstWs, gruppe, fragenFG, projektnummer, projektname,
        auftraggeber, setting, methodeFG, quoteStartCol, gruppen, sheetOptions
      );
      newSheetNames.push({ sheet: sheetName, quoteStartCol, brutto, tnEnd });
    };

    if (isIDI) {
      const hauptGruppe = { ...gruppen[0], methode };
      processGruppe(hauptGruppe, methode === 'VDI' ? 'VDIs' : 'IDIs', fragenArr, true);
    } else {
      for (const gruppe of gruppen) {
        gruppe.methode = gruppe.methode || methode;
        const sheetName = gruppe.id.replace(/[:\\/\?\*\[\]]/g, '').substring(0, 31);
        const fragenFG = fragenArr.filter(f =>
          !f.relevantFuerGruppen ||
          f.relevantFuerGruppen.includes('alle') ||
          f.relevantFuerGruppen.includes(gruppe.id)
        );
        processGruppe(gruppe, sheetName, fragenFG, false);
      }
    }

    const buffer = await result.xlsx.writeBuffer();
    // v12.5: Dateiname kompakt — Spaces aus Projektnummer entfernen,
    // einheitlich mit _ verbinden: "26 1051 6264" -> "26_1051_6264"
    const projNrCompact = String(projektnummer || '').replace(/\s+/g, '_').trim();
    const projNameClean = String(projektname || '').replace(/[^a-zA-Z0-9äöüÄÖÜß]+/g, '_').replace(/^_|_$/g, '');
    const dateiname = `${projNrCompact}_${projNameClean}_Preview.xlsx`
      .replace(/[^a-zA-Z0-9_\-\.äöüÄÖÜß]/g, '_')
      .replace(/_+/g, '_');

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
        compactMatrix: builderOptions.compactMatrix,
        compactMatrixThreshold: builderOptions.compactMatrixThreshold,
        // v11
        idiProfileCount: builderOptions.idiProfile?.length || 0,
        segmentBeschreibungenCount: builderOptions.segment_beschreibungen
          ? Object.keys(builderOptions.segment_beschreibungen).length : 0,
        studienQuotenCount: builderOptions.studien_quoten?.length || 0,
        // v11.3 NEU: Merges aus dem rohen Template-Buffer
        preMergesCount: preMerges?.length || 0,
        preMerges: preMerges || [],
        // v11.4 NEU: Termin-Blöcke + Laufzeit
        terminBlocksCount: builderOptions.termin_blocks?.length || 0,
        laufzeitVon: builderOptions.laufzeitVon,
        laufzeitBis: builderOptions.laufzeitBis,
        version: 'v12.5-multi-firma-termin-dateiname',
      },
    });
  } catch (err) {
    console.error('Error in build-preview:', err);
    console.error('Stack:', err?.stack);
    // Strukturierte Fehlerantwort, damit n8n nicht nur "500" sieht
    return res.status(500).json({
      success: false,
      error: err?.message || 'Unknown error',
      errorType: err?.name || 'Error',
      stack: err?.stack ? String(err.stack).split('\n').slice(0, 8) : null,
      version: 'v12.5-multi-firma-termin-dateiname',
    });
  }
}
