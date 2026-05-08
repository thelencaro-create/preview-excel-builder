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
  cell.alignment = { wrapText: true, vertical: 'top' };
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
function writeQuestionColumn(ws, col, label, note, antList, quoteText, tnEnd, gruppe, frage, isLastInGroup, allGruppen) {
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
  // Bei conditional Fragen: Wechsel-Symbol vor Frage-Titel (Patch 3)
  h.value = isConditional ? `🔀 ${label}` : label;
  h.font = { bold: true, name: 'Arial', size: 9, color: { argb: 'FF000000' } };
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
    c.alignment = { wrapText: true, vertical: 'top' };
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
    const qc = ws.getCell(row, col);
    qc.value = finalQuoteText;
    qc.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.GRUEN } };
    qc.font = { name: 'Arial', size: 9, bold: true, color: { argb: COLORS.QUOTE_FONT } };
    qc.border = cellBorder;
    qc.alignment = { wrapText: true, vertical: 'top' };
    // Zeilenhöhe der Quote-Zeile dynamisch erhöhen, wenn mehrzeilig
    const lineCount = finalQuoteText.split('\n').length;
    if (lineCount > 1) {
      ws.getRow(row).height = Math.min(15 + lineCount * 14, 200);
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
      headerCell.alignment = { wrapText: true, vertical: 'top' };
      row++;
      // Beschreibung (kann mehrzeilig sein)
      const descCell = ws.getCell(row, col);
      descCell.value = String(desc);
      descCell.font = { name: 'Arial', size: 9 };
      descCell.alignment = { wrapText: true, vertical: 'top' };
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
      c.alignment = { wrapText: true, vertical: 'top' };
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
      c.alignment = { wrapText: true, vertical: 'top' };
      row++;
    }
  }
}

// v11: Frage-Header-Label bauen.
//  - Bevorzugt frage.kurz_label (Parser-generiert, max ~30 Zeichen)
//  - Fallback: frage.fragetext (gekürzt auf 60 Zeichen)
//  - Frage-Nummer (id) wird vorangestellt, wenn nicht schon im kurz_label enthalten
function buildHeaderLabel(frage) {
  const id = (frage.id || '').trim();
  const kurz = (frage.kurz_label || frage.kurzlabel || '').trim();
  const lang = (frage.fragetext || '').trim();
  if (kurz) {
    // Wenn kurz_label schon mit der ID startet, nicht doppelt
    if (id && !kurz.toLowerCase().startsWith(id.toLowerCase())) {
      return `${id}. ${kurz}`.substring(0, 60);
    }
    return kurz.substring(0, 60);
  }
  if (id && lang) return `${id}. ${lang}`.substring(0, 60);
  return (lang || id).substring(0, 60);
}

function fillSheet(ws, gruppe, fragen, projektnummer, projektname, kundenname, setting, methode, quoteStartCol, allGruppen, options) {
  const opts = options || {};
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
  if (hmap.studioLabel) {
    ws.getCell(hmap.studioLabel).value = 'F&T Standort';
  }
  if (hmap.studioWert && gruppe.standort) {
    ws.getCell(hmap.studioWert).value = gruppe.standort;
  }
  ws.getCell(hmap.terminWert).value     = gruppe.termin || `${gruppe.datum || ''} ${gruppe.uhrzeit || ''}`.trim();
  ws.getCell(hmap.kunde).value          = kundenname || '';
  // PATCH 5: Zielgruppe + Zuordnungs-Kriterien als Tooltip am Zielgruppe-Feld
  ws.getCell(hmap.zielgruppe).value     = gruppe.zielgruppe || '';
  if (gruppe.zuordnungs_kriterien) {
    ws.getCell(hmap.zielgruppe).note = `Zuordnungs-Kriterien:\n${gruppe.zuordnungs_kriterien}`;
  }
  ws.getCell(hmap.projekt).value        = projektname;
  ws.getCell(hmap.projNr).value         = projektnummer;
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
  for (let r = TN_START; r <= tnEnd; r++) {
    const c = ws.getCell(r, lfdCol);
    c.value = r - TN_START + 1;
    c.alignment = { horizontal: 'center', vertical: 'center' };
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
    // Sicherheitsnetz: Fragen mit kategorie='entfaellt' oder 'verfuegbarkeit'
    // werden übersprungen, auch wenn typ noch single_choice ist.
    if (frage.typ === 'entfaellt' || 
        frage.kategorie === 'entfaellt' || 
        frage.kategorie === 'verfuegbarkeit') {
      continue;
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
        writeQuestionColumn(ws, currentCol, item.item_label || '', itemNote,
                            ants, itemQuote, tnEnd, gruppe, frage, isLast, allGruppen);
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
        writeQuestionColumn(ws, currentCol, summaryLabel, summaryNote,
                            summaryAnts, summaryQuote, tnEnd, gruppe, frage, true, allGruppen);
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
      writeQuestionColumn(ws, currentCol, label, note, null, null, tnEnd, gruppe, frage, true, allGruppen);
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
      writeQuestionColumn(ws, currentCol, label + ' (Tool)', note,
                          [], quoteText, tnEnd, gruppe, frage, true, allGruppen);
      ws.getColumn(currentCol).width = 30;
      currentCol++;
    } else {
      // single_choice, multi_choice, ranking
      const label = buildHeaderLabel(frage);
      const note = frage.bedingung ? `Bedingung: ${frage.bedingung}` : '';
      let quoteText = frage.quotenkommentar || frage.bedingung || '';
      // Auto-Quote-Hinweis für F1 (Geschlecht) und F2 (Alter) aus Gruppen-Daten
      if (!quoteText) {
        const ftLower = (frage.fragetext || '').toLowerCase();
        if (ftLower.includes('geschlecht') && gruppe.geschlecht && gruppe.geschlecht !== 'gemischt') {
          quoteText = `Quote: nur ${gruppe.geschlecht}`;
        } else if ((ftLower.includes('alt sind') || ftLower.includes('alter')) && gruppe.alter_min && gruppe.alter_max) {
          quoteText = `Quote: ${gruppe.alter_min}-${gruppe.alter_max} Jahre`;
        }
      }
      writeQuestionColumn(ws, currentCol, label, note,
                          frage.antworten, quoteText, tnEnd, gruppe, frage, true, allGruppen);
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
    };

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

    const processGruppe = (gruppe, sheetName, fragenFG, includeIdiInfo) => {
      const dstWs = result.addWorksheet(sheetName);
      copyWorksheet(srcWs, dstWs);

      const methodeFG = gruppe.methode || methode;
      const cfgFG = SHEET_CONFIG[methodeFG] || SHEET_CONFIG['GD'];
      const quoteStartCol = findQuoteStartCol(dstWs, cfgFG.quoteStartCol);
      const brutto = gruppe.brutto || 8;
      const tnEnd = TN_START + brutto - 1;

      addConditionalFormats(dstWs, tnEnd);
      addLogo(dstWs, result, gruppe.unternehmen, setting);

      // v11: nur dem IDI/VDI-Sheet die IDI-Info-Blöcke geben
      const sheetOptions = includeIdiInfo
        ? builderOptions
        : { ...builderOptions, idiProfile: null, segment_beschreibungen: null, studien_quoten: null };

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
        compactMatrix: builderOptions.compactMatrix,
        compactMatrixThreshold: builderOptions.compactMatrixThreshold,
        // v11
        idiProfileCount: builderOptions.idiProfile?.length || 0,
        segmentBeschreibungenCount: builderOptions.segment_beschreibungen
          ? Object.keys(builderOptions.segment_beschreibungen).length : 0,
        studienQuotenCount: builderOptions.studien_quoten?.length || 0,
        version: 'v11-universal',
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
      version: 'v11-universal',
    });
  }
}
