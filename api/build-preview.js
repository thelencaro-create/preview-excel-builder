// api/build-preview.js
// Preview Generator – Excel Builder v12.22.24 (Bug-9-Fix: VGD-Sektionen)
//
// v12.22.24 (21.05.2026): Hotfix Bug 9 nach Test-Run 26_2345_3657_AI:
// - Bug 9 (Regress aus v12.22.21 Bug 2): VGD-Sektionen in der Quoten-
//   uebersicht zeigten 3x '(keine Gruppen-spezifischen Anforderungen)',
//   obwohl GD1/GD2/GD3 nachweislich unterschiedliche Quoten haben (siehe
//   QTRUSTTOOL: GD1=Codes 1-3, GD2=Codes 4-5, GD3=Natural Fallout).
//   Ursache: Exakt-Match Dedupe gegen globale Hinweise war zu strikt —
//   zusammengesetzte Quote-Strings 'GD1: ... GD2: ... GD3: ...' sind
//   global UND in jeder VGD-Sektion textgleich → komplette Sektion
//   fiel raus.
//   Fix: filterQuoteByGruppe() (existiert seit v12.15.2) pro Gruppe
//   anwenden. So bleibt in VGD1 nur der GD1-Teil, in VGD2 nur der GD2-
//   Teil etc. Header umbenannt von 'HINWEISE FUER DIESE GRUPPE' zu
//   'ANFORDERUNGEN SPEZIFISCH FUER {gruppe.id}' (klarer).
//
// v12.22.23 (21.05.2026): Hotfix nach Test-Run 26_1234_2345_AI:
// - Bug 7 v12.22.22b war zu aggressiv: 'mindestens einmal im Monat' aus
//   frage.quotenkommentar triggert maxCode=3 fuer alle Codes der Frage —
//   bei QAIUSED (single_choice mit Tool-Labels, KEINE Frequenz-Skala)
//   wurden Codes 4-15 (Google Gemini, Meta AI, ...) faelschlich rot.
//   Neue Regel: parseExclusionCodes Pattern 'mindestens einmal im Monat'
//   wird unterdrueckt, wenn der Text 'qualifizieren'/'ENDE'/'HOLD'
//   enthaelt (= Frage-weite Qualifizierungsanweisung, kein Code-Filter).
//   Zusaetzlich: isAntOffTarget priorisiert items[].soll_quote +
//   ant.soll_quote (PATCH 24/25 Schema) ueber frage.quotenkommentar.
//   Bei single_choice ohne items wird frage.quotenkommentar nicht mehr
//   ausgewertet — zu viele false positives.
// - Bug 6 analog konservativer: nur strukturierte Quellen.
//
// v12.22.22b (20.05.2026): Hotfix nach erstem Test-Run AI:
// - Bug 5: #VALUE!-Errors in Template-Headern (Z1-Z4) werden defensiv geleert.
//          Templates hatten in F1/E1 einen gespeicherten Excel-Error — bis zur
//          Re-Generation der Templates fängt der Builder das ab.
// - Bug 6: Quotengruppen-Marker fuer quoten-relevante Matrix-Items: 📌 +
//          amber Header-Hintergrund, sodass Recruiter sofort sieht WO eine
//          Quotengruppe entsteht (z.B. 'ChatGPT' bei QAIFREQUENCY).
// - Bug 7: ENDE-Code-Heuristik: parseExclusionCodes() extrahiert aus
//          quoteText-Strings Muster wie 'MUSS Code X' / 'nur Code X' und
//          markiert die NICHT-erlaubten Codes rot. Konservativ: lieber kein
//          Marker als ein falscher.
// - Bug 8: Forced-Choice A/B-Matrizen: Wenn JSON 'aussage_a' + 'aussage_b'
//          pro Item liefert, werden beide Aussagen mit A/B-Pills gestapelt
//          in Z6 dargestellt (vorher nur Aussage A sichtbar).
//
// v12.22.21 (20.05.2026): 4 Bugfixes aus Test-Run AI-Preview:
// - Bug 1: Quotenübersicht — Frage-Header pro Gruppe einmal, Codes eingerückt
//          statt 'QGENDER. Geschlecht' 5x wiederholt
// - Bug 2: VGD-Fallback (HINWEISE FÜR DIESE GRUPPE) — exakt-Match Dedupe gegen
//          globale HINWEISE FÜR DEN REKRUTIERER (keine wortgleiche Doppelung)
// - Bug 3: stripQCode() entfernt 'Qxxxxx.' Präfix aus Labels (Deutsch-only).
//          Comments im TN-Tabellen-Header (Z7) entfernt — Z6 ist die Quelle.
// - Bug 4: Matrix-Items: Fragetext jetzt in Z6 gemerged über Item-Spalten
//          (vorher leer, da Z5-Suppression versehentlich auch Z6 traf)
//
// Builder v8 (template-aware headers)
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
//
// v12.22.0 (15.05.2026): QUOTENGRUPPEN-LOGIK
// - Neue Spalte F "Quotengruppe" in den Vorlagen (vom User händisch eingefügt)
// - Builder erkennt die Spalte am Header-Text "Quotengruppe" in Z6
// - Wenn ein Sheet >1 Quotengruppe enthält:
//     * Dropdown (Data Validation) in Spalte F pro TN-Zeile
//     * Quoten-Übersichts-Box ab quoteStartCol in Z1-Z4 (Label / Soll / Ist / Ampel)
//     * COUNTIF-Formeln für Live-Soll-Ist-Abgleich
// - Wenn ein Sheet =1 Quotengruppe enthält: Spalte F per hidden=true ausgeblendet
// - Backward-Compat: Vorlagen ohne Spalte F (Header != "Quotengruppe") werden
//   unverändert verarbeitet — keine QG-Features, alles wie v12.21.7
// - Backward-Compat: Datenformat aus altem Schema (gruppe.zielgruppe als String,
//   idiProfile[] als separate Liste) wird intern zu quotengruppen[] gemappt
//   via buildQuotengruppenForSheet(). Neues Schema (gruppe.quotengruppen[]) hat
//   Vorrang, Fallback ist Single-QG aus gruppe.zielgruppe.
// - matchesGruppe() versteht hierarchische QG-IDs ("GD1.A", "IDI.S1") zusätzlich
//   zu Gruppen-IDs und Zielgruppen-Namen.
//
// v12.22.1 (15.05.2026): QG-DIAGNOSE-FIX
// - findQuotengruppenCol robuster: nutzt cell.text als kanonische Wert-Quelle,
//   Fallback cell.value (mit RichText-Unterstützung)
// - Spalte ausblenden: zusätzlich width=0.1 setzen (ExcelJS-hidden allein
//   ist nicht immer zuverlässig)
// - Debug-Info pro Sheet in Response-JSON unter debug.qgDebug — zeigt
//   ob/wie QG-Spalte gefunden wurde, qgListLength, action
//
// v12.22.2 (15.05.2026): QG-BLOCK AN RICHTIGE STELLE
// - Bug-Fix: in v12.22.0 wurde der QG-Block per str_replace versehentlich
//   in buildOverviewSheet eingefügt (gleicher "return ws;"-Anker), nicht in
//   fillSheet. Resultat: QG-Logik lief nur für das Quotenübersicht-Sheet
//   (mit Fehler "gruppe is not defined"), nicht für die echten Gruppen-Sheets.
// - Korrektur: QG-Block aus buildOverviewSheet entfernt, ans ECHTE Ende von
//   fillSheet eingefügt (nach ws.views-Setzung). fillSheet hatte kein expli-
//   zites "return ws;" — daher gab es keinen eindeutigen Anker.
//
// v12.22.3 (15.05.2026): QG-DROPDOWN-STYLING
// - Schriftgröße der QG-Zellen auf 10pt (war 9pt) — angeglichen an Nachbar-
//   zellen wie "lfd. Nr." und "Datum".
// - Horizontale Ausrichtung center (war left) — gleich wie Nachbarzellen.
// - Bedingte Formatierung pro QG-Wert: jeder Eintrag bekommt eine eigene
//   Hintergrundfarbe aus der Vorlagen-Palette (HELLBLAU, HELLGRUEN, GELB,
//   GRUEN, ORANGE). ROT/DUNKELROT werden NICHT genutzt (Screenout-Semantik).
//   Recruiter sieht auf einen Blick die Quotengruppen-Verteilung.
//
// v12.22.5 (15.05.2026): VORLAGEN-LOGO ÜBERNEHMEN
// - copyWorksheet kopiert jetzt die Bilder aus der Vorlage (statt zu überspringen).
//   Das Logo erscheint in genau der Position und Größe, die in der Vorlage steht.
// - editAs='oneCell' erzwungen → Logo wächst NICHT mit Spaltenbreiten-Änderungen
//   (verhindert das alte Skalierungsproblem mit langen IDI-Terminen).
// - addLogo() ist jetzt FALLBACK: greift nur wenn Vorlage kein Logo hat oder
//   die Übernahme scheitert (z.B. ExcelJS-Quirks beim Image-Lesen).
// - Marker dstWs._templateLogoCopied steuert den Fallback-Pfad.
//
// v12.22.6 (15.05.2026): LOGO PIXEL-GRÖSSE PINNEN
// - Bug in v12.22.5: ExcelJS ignoriert editAs='oneCell' bei twoCellAnchor —
//   das Logo verhält sich weiter wie in der Vorlage und skaliert mit den
//   Output-Spaltenbreiten. Folge: Wenn QG-Spalte F auf width=0.1 ausgeblendet
//   wird ODER IDI-Termin-Spalten verbreitert werden, wird das Logo verzerrt.
// - Fix: Statt twoCellAnchor übernehmen wir die Source-Pixel-Größe einmalig
//   (berechnet aus VORLAGEN-Spaltenbreiten) und schreiben sie als ext={w,h}
//   in einen oneCellAnchor. Damit ist die Logo-Größe fix in Pixel, egal was
//   im Output mit den Spalten passiert.
// - computeImagePixelSize() rechnet Spaltenbreite (chars) → Pixel und
//   Zeilenhöhe (pt) → Pixel mit Excel-Standardformeln.
//
// v12.22.7 (15.05.2026): LOGO ECHTE BILD-GRÖSSE NUTZEN
// - Bug in v12.22.6: computeImagePixelSize summierte VORLAGEN-Spaltenbreiten,
//   was ÜBERSCHÄTZT — die echte Bild-Anzeigegröße steht in <a:ext> innerhalb
//   <xdr:spPr><a:xfrm> der DrawingML. Z.B. F&T-Logo: Vorlage hat E1:H3 als
//   Anker, aber die echte <a:ext> ist 268×77 px (Logo-Original-Maße).
//   v12.22.6 hat 287 px berechnet → Logo zu breit, überdeckte "F&T Standort".
// - Fix: img.range.ext primär nutzen (ExcelJS liest <a:ext> mit), nur wenn
//   das fehlt, auf Spalten-Berechnung zurückfallen. EMU↔Pixel-Umrechnung
//   per Heuristik (Werte >10000 = EMU).
//
// v12.22.20 (20.05.2026): DEAD-CODE-CLEANUP
// - Drei Funktionen entfernt die seit v12.22.15 nicht mehr aufgerufen wurden:
//   * renderQuotenUebersicht (vertikale QG-Box, ersetzt durch writeQgBoxUnderTn)
//   * renderQuotenUebersichtHorizontal (horizontale QG-Box, v12.22.14-Experiment)
//   * writeIdiInfoBlock (No-Op-Stub, war Segment/IDI-Liste/Studien-Quoten-Renderer)
// - Reduziert die Datei um ~230 Zeilen. Keine funktionale Aenderung.
//
// v12.22.19 (20.05.2026): LAYOUT-KONSISTENZ-PASS (PRE-DEPLOY-AUDIT)
// - User-Hinweis: doppelte/veraltete Anweisungen zu Logo, Frage-Header,
//   TN-Start im Code pruefen. Diese Version: keine funktionalen Aenderungen,
//   nur Cleanups.
// - Audit-Ergebnisse:
//   * Layout-Konstanten (TN_START, HEADER_ROW, LONG_QUESTION_ROW,
//     MATRIX_HEADER_ROW) konsequent ueberall referenziert — KEIN hardcoded
//     Z6/Z7 im Code.
//   * Logo-Pipeline: ein-Weg-Pfad, kein Doppel-Add (copyWorksheet kopiert
//     Vorlagen-Logo, addLogo greift nur als Fallback bei _templateLogoCopied=false).
//   * Frage-Header: writeQuestionColumn ist die EINZIGE Stelle die Z6+Z7
//     beschreibt (Frage-Text + Kurzlabel).
//   * Veraltete Kommentare/Variable-Namen (row6 → headerRow) aktualisiert.
//   * Beispiel-Zelladressen in DV-Kommentaren von Z7→Z8 angepasst.
//
// v12.22.18 (20.05.2026): VOLLE SCREENER-FRAGE IN Z6 (STRUKTURELLE LOESUNG)
// - Templates wurden angepasst: alle 6 Vorlagen (offline/online x F_T/HG/ms)
//   haben jetzt eine zusaetzliche leere Zeile zwischen Matrix-Header (Z5)
//   und TN-Tabellen-Header (Z7 alt 6).
// - Konstanten-Verschiebung:
//   Z5 = MATRIX_HEADER_ROW (unveraendert)
//   Z6 = LONG_QUESTION_ROW (NEU - lange Screener-Frage)
//   Z7 = HEADER_ROW (war 6 - Kurzlabel + Tabellen-Header)
//   Z8 = TN_START (war 7 - erste TN-Zeile)
// - writeQuestionColumn schreibt volle frage.fragetext in Z6 (Smart-Suppression
//   bei Matrix-Items + Redundanz mit kurz_label).
// - Layout-stabil: keine spliceRows-Tricks, keine Bild-Anchor-Manipulation.
//   Templates und Code wachsen synchron.
//
// v12.22.17 (20.05.2026): REVERT v12.22.16 (Z4-FRAGE)
// - Problem: v12.22.16 hat lange Fragen in Z4 geschrieben + Z4-Hoehe auf
//   60pt erhoeht. Bei VGD-Templates mit Sheet-Header bis Spalte M und Logo
//   im Merge F1:H4 fuehrte das zu Layout-Chaos: Logo verschoben, Frage-
//   Texte direkt neben den Sheet-Header-Eintraegen (A4 'Rekru Standorte',
//   I4 'F&T Projekt-Nr.', L4 'Incentive').
// - Fix: Z4-Frage-Block aus writeQuestionColumn entfernt, Z4-Hoehe nicht
//   mehr veraendert. Layout identisch zu v12.22.15.
// - (Veraltet seit v12.22.18: LONG_QUESTION_ROW wieder aktiv, jetzt aber
//   in Z6 mit Template-Unterstuetzung statt Z4-Quick-Fix.)
//
// v12.22.16 (17.05.2026): VOLLE SCREENER-FRAGE IN Z4 (LONG_QUESTION_ROW) — REVERTED
//
// v12.22.15 (17.05.2026): QG-BOX UNTER TN IN SPALTE G + SEGMENTE RAUS
// - User-Wunsch: QG-Counter unter TN-Bereich in Spalte G (dort wo frueher
//   Segment-Beschreibungen, IDI-Profile-Liste und Studien-Quoten standen).
//   Recruiter sieht ihn direkt nach dem Scrollen unter den letzten TN-Zeilen.
// - Segment-Beschreibungen, IDI-Profile und Studien-Quoten werden NICHT mehr
//   im Sheet-Body angezeigt — der Recruiter findet sie in der Quotenübersicht
//   (Sheet 1). writeIdiInfoBlock entfernt (v12.22.20).
// - QG-Box-Layout (Z(tnEnd+8) abwaerts ab Spalte G):
//     Row 1: "Quotengruppen-Counter" (Titel)
//     Row 2: | Quotengruppe | Soll | Ist | Status |
//     Row 3+: Daten-Zeile pro QG mit COUNTIF + Ampel
// - Horizontale Box-Variante (v12.22.14) entfernt (v12.22.20).
//
// v12.22.14 (17.05.2026): HORIZONTALES QG-BOX-LAYOUT BEI VIELEN QGs
// - User-Feedback: bei v12.22.11 wurde die QG-Box bei ≥5 QGs nach rechts
//   hinter den letzten Frage-Bereich (BG+) verschoben. Damit war der
//   Soll/Ist-Counter im Recruiting-Workflow nicht mehr sichtbar — der
//   Recruiter musste extra dorthin scrollen.
// - Fix: Box bleibt IMMER an quoteStartCol (Spalte Q oberhalb der Frage-
//   Spalten). Bei ≤4 QGs: vertikales Layout wie bisher (Z1-Z(N+1) in
//   4 Spalten breit). Bei ≥5 QGs: HORIZONTALES Layout in Z1-Z4 ueber N
//   Spalten breit (Z1=Label rotiert, Z2=Soll, Z3=Ist, Z4=Status-Ampel).
//   Belegt damit Z1-Z4 ueber den ersten N Frage-Spalten, kollidiert nicht
//   mit HEADER_ROW=6.
// - Recruiter sieht damit immer den Counter "noch X uebrig pro QG"
//   direkt im Blickfeld, egal wie viele QGs.
//
// v12.22.13 (17.05.2026): LABEL-KONSISTENZ ZWISCHEN BOX, DROPDOWN, COUNTIF
// - Problem: Box-Label zeigte Kommas ("got2b-Nutzer, weiblich, 16-24 Jahre"),
//   Dropdown und COUNTIF Slashes ("got2b-Nutzer / weiblich / 16-24 Jahre").
//   Sah visuell inkonsistent aus (Box vs. Dropdown-Auswahl). Counter zaehlte
//   technisch korrekt (Dropdown und COUNTIF beide mit Slash) — aber das
//   Box-Label entsprach NICHT dem was im Dropdown stand.
// - Fix: Box-Label benutzt jetzt dieselbe Sanitize-Logik wie Dropdown
//   (Komma -> ' /'). Alle drei Stellen (Box-Anzeige, Dropdown-Liste,
//   COUNTIF-Suchstring) sind jetzt deckungsgleich.
//
// v12.22.12 (17.05.2026): IDI-PROFIL N_TARGET-FIX
// - Problem: Bei got2b lieferte der Parser nur 4 idiProfile-Eintraege
//   (Block-Stellvertreter IDI1, IDI4, IDI6, IDI9 mit "3 Teilnehmer"-Text
//   in profile_quoten) fuer 12 echte Interview-Slots. Builder zaehlte
//   info.count++ pro Eintrag -> Soll 2 statt 6 pro QG-Segment.
// - Fix: 3-stufige n_target-Berechnung im idiProfile-Pfad:
//   a) Extrahiere "N Teilnehmer/TN/Personen" aus profile_quoten.text und
//      summiere pro Segment (got2b-Fall: 3+3=6 pro Segment)
//   b) Fallback: gruppe.brutto durch Anzahl Segmente teilen (12/2=6)
//   c) Fallback: info.count wie bisher (klassisch JPM mit individuellen
//      idiProfile-Eintraegen)
// - Logging: qgDebug._meta='idi_profile_qg' enthaelt nTargetSource +
//   sumNTarget zur Forensik (sollte gleich brutto sein).
//
// v12.22.11 (17.05.2026): QG-BOX KOLLISIONS-FIX
// - Problem: Bei >= 5 Quotengruppen (z.B. got2b-IDI mit 8 Demographie-
//   Sub-Profilen) lief die QG-Uebersichts-Box von Z1 bis Z(N+1) und
//   ueberschrieb damit ab Z6 die Frage-Header-Zeile. In den 8-QG-Faellen
//   stand "Nicht-Nutzer got2b 16-24 weiblich" auf Q6 statt der ersten
//   Frage.
// - Fix: Box-Position dynamisch. Bei <= 4 QGs bleibt sie wie bisher an
//   quoteStartCol (Spalte Q ueber den Antwort-Spalten, kompakt). Bei
//   >= 5 QGs wird sie nach rechts hinter den letzten Frage-Bereich
//   verschoben (currentCol+1), damit keine Kollision mit Z6 entsteht.
// - Logging: qgDebug.boxPosition fasst die Entscheidung zusammen.
//
// v12.22.10 (17.05.2026): AUTO-EXTRACT QUOTENGRUPPEN AUS FREITEXT + DV-FIX
// - Problem A: Parser ignoriert PATCH 22 + Extract-Regel (v12.14.2) bei VELOXX
//   weil der "4 Milchnutzer + 2 PBB-Nutzer" Hinweis im quotenkommentar einer
//   Frage steht statt in gruppe.quotengruppen[]. Claude entscheidet bei
//   Unsicherheit konservativ und laesst quotengruppen[] weg.
// - Fix A: Builder-seitiger deterministischer Extractor ergaenzt quotengruppen[]
//   automatisch, wenn er im quotenkommentar oder studien_quoten ein
//   "N [Label] + M [Label]" Pattern findet, das Gruppen-IDs der aktuellen
//   Gruppe nennt (z.B. "GD3+GD4: davon 4 Milchnutzer + 2 PBB-Nutzer").
// - Problem B: Bei langen QG-Listen (>250 Zeichen gesamt, z.B. got2b-IDI mit
//   8 Demographie-Sub-Profilen) fiel Excel-Inline-DataValidation aus
//   (Limit), Dropdown fehlte. Box wurde trotzdem rendered.
// - Fix B: Fallback auf Range-Reference: Labels werden in Hilfsspalte 250
//   (weit rechts, ausserhalb sichtbaren Bereich) geschrieben, DataValidation
//   referenziert per "=$IP$1:$IP$N". Damit unlimitierte Listen-Laenge.
// - Konservativ: Greift nur wenn gruppe.quotengruppen leer/nicht vorhanden ist.
//   Wenn Parser was geliefert hat, hat das Vorrang.
// - Logging: Extrahierte QGs landen in qgDebug mit _meta='auto_extracted' fuer
//   Forensik bei Fehlmatches.
//
// v12.22.9 (16.05.2026): SORTIER-FIX + LETZTE-TEILNAHME-FILTER
// - Bug 1: F11 "Alter der Kinder" wurde vor F1-F6 sortiert, weil REGEX_ALTER
//   `\balter\b` auch auf "Alter der Kinder" matchte. Fix: spezifischere
//   Regex die nur die Hauptfrage "Wie alt sind Sie?" / "Alter [von Antworten]"
//   matcht — Folgefragen mit "Alter der Kinder" / "Alter der Eltern" etc.
//   bekommen Score 2 (andere persönliche Daten) und landen in
//   Screener-Reihenfolge nach F10.
// - Bug 2: F6 "Letzte Studie 3 Monate" / "Wann zuletzt an Marktforschung
//   teilgenommen" wurde vom Parser durchgeschleift trotz PATCH 20-Anweisung.
//   Fix: Builder-seitiger Fallback-Filter mit Pattern auf kurz_label und
//   fragetext — fängt typische Studien-Teilnahme-Historie-Fragen ab und
//   filtert sie analog zu kategorie='entfaellt'.
//
// v12.22.8 (15.05.2026): QG-SPALTE GRAU STATT AUSGEBLENDET
// - Wenn ein Sheet nur 1 Quotengruppe enthält: Spalte F NICHT mehr verstecken
//   (hidden + width=0.1 kollabierten den Logo-Bereich → Logo verzerrt).
// - Stattdessen: Header- und TN-Zellen der QG-Spalte grau einfärben als
//   Signal "hier nicht genutzt". User kann die Spalte bei Bedarf manuell
//   löschen — kostet 2 Klicks (Rechtsklick → Spalte löschen).
// - Vorteil: Logo-Anker E1:H3 hat alle Spaltenbreiten zur Verfügung, das
//   Vorlagen-Logo wird in voller Größe korrekt angezeigt.
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

// v12.22.18: Template hat jetzt eine zusaetzliche leere Zeile (Z6) zwischen
// Matrix-Header (Z5) und Tabellen-Header (Z7). Damit wandern alle Layout-
// Zeilen um 1 runter:
//   Z5 = MATRIX_HEADER_ROW (unveraendert)
//   Z6 = LONG_QUESTION_ROW (NEU - lange Screener-Frage)
//   Z7 = HEADER_ROW (war Z6 - Kurzlabel + Tabellen-Header)
//   Z8 = TN_START (war Z7 - erste TN-Zeile)
const TN_START = 8;
const MATRIX_HEADER_ROW = 5;
const LONG_QUESTION_ROW = 6;
const HEADER_ROW = 7;
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

  // v12.22.6: Bilder der Vorlage werden übernommen, ABER als oneCellAnchor
  // mit fester Pixelgröße — NICHT als twoCellAnchor.
  // Grund: twoCellAnchor skaliert das Logo mit den Spaltenbreiten des
  // Outputs. Wenn Builder die Spalten breiter macht (z.B. IDI-Termine in G)
  // oder eine Spalte ausblendet (QG-Spalte F mit width=0.1), wird das Logo
  // verzerrt/winzig dargestellt.
  // Lösung: Pixelgröße einmalig aus der VORLAGEN-Spaltenbreite berechnen
  // (das ist die "gemeinte" Größe) und dann fix in den Anker schreiben.
  try {
    const srcImages = (typeof srcWs.getImages === 'function') ? srcWs.getImages() : [];
    const srcWb = srcWs.workbook;
    let copiedCount = 0;
    for (const img of srcImages) {
      try {
        const mediaEntry = srcWb.model?.media?.[img.imageId];
        if (!mediaEntry || !mediaEntry.buffer) continue;
        const ext = mediaEntry.extension || 'png';
        const newImageId = dstWs.workbook.addImage({
          buffer: mediaEntry.buffer,
          extension: ext,
        });
        // Anker zur Build-Zeit fest-pinnen:
        const tl = {
          col: img.range.tl.nativeCol,
          row: img.range.tl.nativeRow,
        };
        // Pixel-Größe ermitteln:
        // 1. PRIORITÄT: img.range.ext (= das tatsächliche Bild-Anzeigemaß aus
        //    der Vorlage, sowohl bei oneCellAnchor als auch bei twoCellAnchor
        //    in der <a:ext> XML-Struktur — ExcelJS liest das mit).
        //    Excel-EMU: 1 px = 9525 EMU. ExcelJS gibt Werte teils in px,
        //    teils in EMU zurück — wir erkennen das anhand der Größenordnung.
        // 2. FALLBACK: aus Spalten- und Zeilengrößen der Source-Vorlage rechnen
        //    (twoCellAnchor ohne <a:ext>).
        let logoPx = null;
        if (img.range.ext &&
            typeof img.range.ext.width === 'number' &&
            typeof img.range.ext.height === 'number' &&
            img.range.ext.width > 0 && img.range.ext.height > 0) {
          let w = img.range.ext.width;
          let h = img.range.ext.height;
          // Heuristik: wenn Werte > 10000, sind sie in EMU (1px = 9525 EMU)
          if (w > 10000 || h > 10000) {
            w = Math.round(w / 9525);
            h = Math.round(h / 9525);
          }
          logoPx = { width: w, height: h };
        } else if (img.range.br && img.range.tl) {
          // Fallback: Größe aus Source-Spaltenbreiten und Zeilenhöhen
          logoPx = computeImagePixelSize(srcWs, img.range.tl, img.range.br);
        }
        if (!logoPx || logoPx.width <= 0 || logoPx.height <= 0) {
          // Letzter Fallback: 268×77 (typisches Querformat-Logo F&T)
          logoPx = { width: 268, height: 77 };
        }
        dstWs.addImage(newImageId, {
          tl,
          ext: logoPx,
          editAs: 'oneCell',
        });
        copiedCount++;
      } catch (e) {
        console.warn(`[Logo] Vorlagen-Image-Übernahme fehlgeschlagen: ${e.message}`);
      }
    }
    dstWs._templateLogoCopied = copiedCount > 0;
  } catch (e) {
    console.warn(`[Logo] Image-Iteration fehlgeschlagen: ${e.message}`);
  }
}

// Hilfsfunktion: berechnet die Pixel-Größe eines Logos zwischen tl und br
// anhand der Spaltenbreiten und Zeilenhöhen der Source-Vorlage.
// Excel-Maße:
//   Spaltenbreite in "characters" → Pixel ≈ width * 7 + 5 (für Calibri 11pt)
//   Zeilenhöhe in points → Pixel = height * 4/3
function computeImagePixelSize(srcWs, tl, br) {
  let widthPx = 0;
  let heightPx = 0;
  // Anteilig: tl.col ist 0-basiert, nativeColOff in EMU (1 px = 9525 EMU)
  // Wir nehmen erst mal die ganzen Spalten/Zeilen, Offsets sind meist klein
  for (let c = tl.nativeCol; c < br.nativeCol; c++) {
    const col = srcWs.getColumn(c + 1);
    const w = col?.width ?? 8.43;  // Excel-Default
    widthPx += Math.round(w * 7 + 5);
  }
  for (let r = tl.nativeRow; r < br.nativeRow; r++) {
    const row = srcWs.getRow(r + 1);
    const h = row?.height ?? 15;  // Excel-Default in pt
    heightPx += Math.round(h * 4 / 3);
  }
  return { width: widthPx, height: heightPx };
}

// Findet die erste "Quote"-Spalte in HEADER_ROW (Z7) dynamisch
function findQuoteStartCol(ws, fallback) {
  const headerRow = ws.getRow(HEADER_ROW);
  let foundCol = null;
  headerRow.eachCell({ includeEmpty: false }, (cell, colNum) => {
    if (foundCol !== null) return;
    const v = String(cell.value ?? '').trim();
    if (/^Quote\b/i.test(v)) foundCol = colNum;
  });
  return foundCol || fallback;
}

// ---------------------------------------------------------------------------
// 1c) QUOTENGRUPPEN-LOGIK (v12.22.0)
// ---------------------------------------------------------------------------
//
// Eine Quotengruppe (QG) ist die kleinste Rekrutierungs-Einheit, für die eigene
// Soll-Zahlen gelten. Jedes Sheet enthält 1..K Quotengruppen.
//
// Datenmodell pro QG:
//   { id, label, n_target, demographics?, zuordnungs_regel?, segment?, cluster? }
//
// Schema-Quellen (in Prioritäts-Reihenfolge):
//   1) gruppe.quotengruppen[]   — neues Schema (ab v12.22), direkt verwendet
//   2) idiProfile[]             — bei konsolidiertem IDI/VDI-Sheet aus idiProfile
//                                 ableiten (1 QG pro Segment-Profil)
//   3) gruppe.zielgruppe        — Fallback: 1 QG, label = zielgruppe, n = brutto

// Findet die Spalte mit Header "Quotengruppe" in HEADER_ROW (Z7).
// Rückgabe: Spalten-Nummer oder null wenn nicht vorhanden (alte Vorlage).
// v12.22.1: robuster gegen RichText/Formula-Werte — nutzt cell.text als Fallback
function findQuotengruppenCol(ws) {
  const headerRow = ws.getRow(HEADER_ROW);
  let foundCol = null;
  headerRow.eachCell({ includeEmpty: false }, (cell, colNum) => {
    if (foundCol !== null) return;
    let v = '';
    // 1) cell.text ist die kanonische Plain-Text-Repräsentation in ExcelJS
    if (cell.text != null) v = String(cell.text);
    // 2) Fallback: cell.value (kann String, Number, Object sein)
    if (!v && cell.value != null) {
      if (typeof cell.value === 'object' && Array.isArray(cell.value.richText)) {
        v = cell.value.richText.map(r => r.text || '').join('');
      } else {
        v = String(cell.value);
      }
    }
    v = v.trim();
    if (/^Quotengruppe$/i.test(v)) foundCol = colNum;
  });
  return foundCol;
}

// v12.22.10: Auto-Extract von Quotengruppen aus Freitext-Hinweisen.
// Sucht in quotenkommentar und studien_quoten nach Pattern wie:
//   "4 Milchnutzer + 2 PBB-Nutzer"
//   "davon 4 [Label] + 2 [Label] pro Gruppe"
//   "je 4 [Label], 4 [Label]"
//   "5 [Label] / 5 [Label]"
// Wenn die aktuelle gruppe.id im selben Text genannt ist (GD3, GD3+GD4 etc.),
// werden die Pattern als quotengruppen[] zurueckgegeben.
// Gibt null zurueck wenn nichts gefunden wurde.
function extractQuotengruppenFromFreitext(gruppe, fragenArr, studienQuoten) {
  if (!gruppe || !gruppe.id) return null;
  const gruppeId = String(gruppe.id);
  const brutto = gruppe.brutto || 8;

  // Sammle alle Texte die fuer diese Gruppe relevant sein koennten.
  // Jeder Bucket hat: source (debug), text (zu scannen), implicitGruppenIds (optional),
  // antwortenMapping (optional, fuer kontext-aware Zuordnung).
  const textBuckets = [];
  if (Array.isArray(fragenArr)) {
    for (const f of fragenArr) {
      if (!f) continue;
      const qk = (f.quotenkommentar || '').trim();
      if (qk) {
        // Sammle alle gilt_fuer_gruppen-Sets pro Antwort + die Antwort-Texte
        const sollQuoteSets = [];
        const antwortContext = [];  // [{antwortText, sollQuoteText, gruppenIds}]
        if (Array.isArray(f.antworten)) {
          for (const ant of f.antworten) {
            if (Array.isArray(ant.soll_quote)) {
              for (const sq of ant.soll_quote) {
                if (Array.isArray(sq.gilt_fuer_gruppen) && sq.gilt_fuer_gruppen.length > 0) {
                  const ids = sq.gilt_fuer_gruppen.map(g => String(g).toUpperCase());
                  sollQuoteSets.push(ids);
                  antwortContext.push({
                    antwortText: (ant.text || '').toLowerCase(),
                    sollQuoteText: (sq.text || '').toLowerCase(),
                    code: ant.code,
                    gruppenIds: ids,
                  });
                }
              }
            }
          }
        }
        // Eindeutige Zuordnung: nur ein Set ODER alle identisch
        let implicit = null;
        if (sollQuoteSets.length === 1) {
          implicit = sollQuoteSets[0];
        } else if (sollQuoteSets.length > 1) {
          const first = sollQuoteSets[0].slice().sort().join(',');
          if (sollQuoteSets.every(s => s.slice().sort().join(',') === first)) implicit = sollQuoteSets[0];
        }
        textBuckets.push({
          source: `frage.${f.id || '?'}.quotenkommentar`,
          text: qk,
          implicitGruppenIds: implicit,
          antwortContext: antwortContext.length > 1 ? antwortContext : null,
        });
      }
      // Auch soll_quote-Texte scannen: hier ist gilt_fuer_gruppen PRO Eintrag klar
      if (Array.isArray(f.antworten)) {
        for (const ant of f.antworten) {
          if (Array.isArray(ant.soll_quote)) {
            for (const sq of ant.soll_quote) {
              const sqt = (sq.text || '').trim();
              if (sqt) textBuckets.push({
                source: `frage.${f.id}.antwort.${ant.code}.soll_quote`,
                text: sqt,
                implicitGruppenIds: Array.isArray(sq.gilt_fuer_gruppen) && sq.gilt_fuer_gruppen.length > 0
                  ? sq.gilt_fuer_gruppen.map(g => String(g).toUpperCase())
                  : null,
                antwortContext: null,
              });
            }
          }
        }
      }
    }
  }
  if (Array.isArray(studienQuoten)) {
    studienQuoten.forEach((sq, i) => {
      const t = (typeof sq === 'string' ? sq : (sq && sq.text) || '').trim();
      if (t) textBuckets.push({ source: `studien_quoten[${i}]`, text: t, implicitGruppenIds: null, antwortContext: null });
    });
  }
  if (textBuckets.length === 0) return null;

  // Pattern fuer "N [Label] + M [Label]" — Label endet vor Bindewort/Punkt/Komma
  // Erlaubt auch Variationen mit "x"/"×"/"TN"/"davon" davor
  // Gruppen-Capture: (zahl1) (label1) (zahl2) (label2)
  const PAT_PLUS = /(?:^|[\s,;:.(])(?:davon\s+|je\s+)?(\d+)\s*(?:x|×|TN\s+)?\s*([\p{L}][\p{L}\d\s\-_/]*?)\s*\+\s*(\d+)\s*(?:x|×|TN\s+)?\s*([\p{L}][\p{L}\d\s\-_/]*?)(?=\s*(?:pro\s+gruppe|je\s+gruppe|anderer?\s+sorten?|sorten|[,;.()\n]|$))/giu;
  // Pattern fuer "N / N" Splits ("5 got2b-Nutzer / 5 Nicht-Nutzer") — etwas restriktiver
  const PAT_SLASH = /(?:^|[\s,;:.(])(\d+)\s+([\p{L}][\p{L}\d\s\-_]*?)\s*\/\s*(\d+)\s+([\p{L}][\p{L}\d\s\-_]*?)(?=[,;.()\n]|$)/giu;

  // Helper: bestimmt fuer einen Match welche Gruppen-IDs gelten.
  // Konservativ — nur sichere Fälle:
  //   1. bucket.implicitGruppenIds gesetzt (eindeutige soll_quote.gilt_fuer_gruppen)
  //   2. Explizite Gruppen-IDs im umgebenden Hauptsatz
  // Heuristik (Token-Match auf Antwort-Text) wurde entfernt — zu unzuverlässig
  // bei paraphrasierten Quotenkommentaren ("2 Gruppen erwägen" vs.
  // "konsumiere keine zuckerfreie Option").
  function resolveGruppenIdsForMatch(bucket, matchPos, matchLen) {
    if (Array.isArray(bucket.implicitGruppenIds)) return bucket.implicitGruppenIds;
    // Suche Gruppen-IDs im Hauptsatz um den Match (zwischen . und .)
    const HARD_SENT = /[.;!\n]/g;
    let segStart = 0, segEnd = bucket.text.length;
    HARD_SENT.lastIndex = 0;
    let m;
    while ((m = HARD_SENT.exec(bucket.text)) !== null && m.index < matchPos) {
      segStart = m.index + 1;
    }
    HARD_SENT.lastIndex = matchPos + matchLen;
    const next = HARD_SENT.exec(bucket.text);
    if (next) segEnd = next.index;
    const segment = bucket.text.substring(segStart, segEnd);
    const ids = (segment.match(/\b(?:GD|IDI|VGD|VDI)\d+\b/gi) || []).map(s => s.toUpperCase());
    return ids.length > 0 ? ids : null;
  }


  const cleanLabel = (s) => {
    if (!s) return '';
    return String(s)
      .replace(/\s+/g, ' ')
      .replace(/^[\s\-_/]+|[\s\-_/]+$/g, '')
      .replace(/^(TN|Teilnehmer)\s+/i, '')
      .trim();
  };

  // Plausi: ein extrahierter Match ist nur dann ein echter QG-Split wenn:
  //   - Summe N+M == brutto  (z.B. 4+2=6, 5+5=10, 4+4=8)
  //   - ODER Summe == brutto - safetyBuffer (manche Studien planen Backups)
  //   - Beide Labels sind plausibel (mehr als 2 chars, kein reines Zahlwort)
  const plausibleSum = (n1, n2) => {
    const sum = n1 + n2;
    return sum === brutto || sum === brutto - 1 || sum === brutto - 2 || sum === brutto + 2;
  };
  const plausibleLabel = (s) => s && s.length >= 3 && s.length <= 40 && !/^\d+$/.test(s);

  for (const bucket of textBuckets) {
    const text = bucket.text;
    // Plus-Pattern
    PAT_PLUS.lastIndex = 0;
    let m;
    while ((m = PAT_PLUS.exec(text)) !== null) {
      const n1 = parseInt(m[1], 10);
      const l1 = cleanLabel(m[2]);
      const n2 = parseInt(m[3], 10);
      const l2 = cleanLabel(m[4]);
      if (!plausibleLabel(l1) || !plausibleLabel(l2)) continue;
      if (!plausibleSum(n1, n2)) continue;
      // Bestimme welche Gruppen-IDs fuer DIESEN Match gelten
      const matchGruppenIds = resolveGruppenIdsForMatch(bucket, m.index, m[0].length)
        || (text.match(/\b(?:GD|IDI|VGD|VDI)\d+\b/gi) || []).map(s => s.toUpperCase());
      if (matchGruppenIds.length === 0) continue;  // keine Zuordnung moeglich
      if (!matchGruppenIds.includes(gruppeId.toUpperCase())) continue;  // nicht diese Gruppe
      // Match!
      const qgs = [
        { id: `${gruppeId}.A`, label: l1, n_target: n1, demographics: {
            alter_min: gruppe.alter_min ?? null, alter_max: gruppe.alter_max ?? null,
            geschlecht: gruppe.geschlecht ?? null, standort: gruppe.standort ?? null },
          zuordnungs_regel: `Auto-extrahiert aus: "${text.substring(0, 100)}${text.length > 100 ? '...' : ''}"`,
        },
        { id: `${gruppeId}.B`, label: l2, n_target: n2, demographics: {
            alter_min: gruppe.alter_min ?? null, alter_max: gruppe.alter_max ?? null,
            geschlecht: gruppe.geschlecht ?? null, standort: gruppe.standort ?? null },
          zuordnungs_regel: `Auto-extrahiert aus: "${text.substring(0, 100)}${text.length > 100 ? '...' : ''}"`,
        },
      ];
      if (!globalThis.__QG_DEBUG__) globalThis.__QG_DEBUG__ = [];
      globalThis.__QG_DEBUG__.push({
        _meta: 'auto_extracted', gruppe: gruppeId, source: bucket.source,
        pattern: 'plus', match: `${n1} ${l1} + ${n2} ${l2}`, sumOk: n1 + n2,
        resolvedFor: matchGruppenIds,
      });
      return qgs;
    }
    // Slash-Pattern
    PAT_SLASH.lastIndex = 0;
    while ((m = PAT_SLASH.exec(text)) !== null) {
      const n1 = parseInt(m[1], 10);
      const l1 = cleanLabel(m[2]);
      const n2 = parseInt(m[3], 10);
      const l2 = cleanLabel(m[4]);
      if (!plausibleLabel(l1) || !plausibleLabel(l2)) continue;
      if (!plausibleSum(n1, n2)) continue;
      const matchGruppenIds = resolveGruppenIdsForMatch(bucket, m.index, m[0].length)
        || (text.match(/\b(?:GD|IDI|VGD|VDI)\d+\b/gi) || []).map(s => s.toUpperCase());
      if (matchGruppenIds.length === 0) continue;
      if (!matchGruppenIds.includes(gruppeId.toUpperCase())) continue;
      const qgs = [
        { id: `${gruppeId}.A`, label: l1, n_target: n1, demographics: {
            alter_min: gruppe.alter_min ?? null, alter_max: gruppe.alter_max ?? null,
            geschlecht: gruppe.geschlecht ?? null, standort: gruppe.standort ?? null },
          zuordnungs_regel: `Auto-extrahiert aus: "${text.substring(0, 100)}${text.length > 100 ? '...' : ''}"`,
        },
        { id: `${gruppeId}.B`, label: l2, n_target: n2, demographics: {
            alter_min: gruppe.alter_min ?? null, alter_max: gruppe.alter_max ?? null,
            geschlecht: gruppe.geschlecht ?? null, standort: gruppe.standort ?? null },
          zuordnungs_regel: `Auto-extrahiert aus: "${text.substring(0, 100)}${text.length > 100 ? '...' : ''}"`,
        },
      ];
      if (!globalThis.__QG_DEBUG__) globalThis.__QG_DEBUG__ = [];
      globalThis.__QG_DEBUG__.push({
        _meta: 'auto_extracted', gruppe: gruppeId, source: bucket.source,
        pattern: 'slash', match: `${n1} ${l1} / ${n2} ${l2}`, sumOk: n1 + n2,
        resolvedFor: matchGruppenIds,
      });
      return qgs;
    }
  }
  return null;
}

// Baut die QG-Liste für ein Sheet aus den vorhandenen Daten.
// Berücksichtigt das neue Schema (gruppe.quotengruppen) und Backward-Compat
// (idiProfile bei konsolidierten IDI-Sheets, gruppe.zielgruppe als Single-QG).
function buildQuotengruppenForSheet(gruppe, opts) {
  const o = opts || {};
  // 1) Neues Schema: gruppe.quotengruppen[] direkt verwenden
  if (Array.isArray(gruppe.quotengruppen) && gruppe.quotengruppen.length > 0) {
    return gruppe.quotengruppen.map((qg, i) => normalizeQG(qg, gruppe, i));
  }
  // v12.22.10: 1b) Auto-Extract aus Freitext-Hinweisen (PATCH 22-Fallback)
  // Greift wenn Parser keine quotengruppen[] geliefert hat, aber im
  // quotenkommentar/studien_quoten ein "N [A] + M [B]"-Pattern fuer diese
  // Gruppe steht.
  const autoExtracted = extractQuotengruppenFromFreitext(gruppe, o.fragenArr, o.studienQuoten);
  if (autoExtracted && autoExtracted.length > 1) {
    return autoExtracted.map((qg, i) => normalizeQG(qg, gruppe, i));
  }
  // 2) IDI/VDI mit idiProfile[] -> 1 QG pro eindeutigem Segment
  const isIDI = gruppe.methode === 'IDI' || gruppe.methode === 'VDI';
  if (isIDI && Array.isArray(o.idiProfile) && o.idiProfile.length > 0) {
    // Unique Segments mit Anzahl
    const segCount = new Map();
    for (const p of o.idiProfile) {
      const key = p.segment || 'Allgemein';
      const cur = segCount.get(key) || { count: 0, sample: p, idis: [] };
      cur.count++;
      cur.idis.push(p);
      segCount.set(key, cur);
    }

    // v12.22.11: Bessere n_target-Berechnung.
    // Problem: Wenn der Parser nicht pro Einzelinterview einen idiProfile-Eintrag
    // liefert sondern pro Termin-Block einen (Block-Stellvertreter), zaehlt
    // info.count nur die Bloecke (z.B. 2+2=4 statt der echten 6+6=12).
    // Strategie:
    //   a) Versuche aus profile_quoten Text "N Teilnehmer/TN" zu extrahieren
    //      (z.B. "3 Teilnehmer zwischen 16 und 24 Jahre alt") -> Pro Eintrag
    //      diese Zahl statt 1.
    //   b) Wenn Summe aller info.count < gruppe.brutto und brutto > 0,
    //      verteile gruppe.brutto proportional auf die Segmente.
    function extractTNCountFromProfileQuoten(idi) {
      if (!Array.isArray(idi.profile_quoten)) return null;
      for (const pq of idi.profile_quoten) {
        const t = (pq && pq.text) || '';
        // Pattern: "3 Teilnehmer" / "3 TN" / "je 3 TN" / "3 Personen"
        const m = t.match(/(?:^|\s|^je\s+)(\d+)\s*(?:teilnehmer|tn\b|personen)/i);
        if (m) {
          const n = parseInt(m[1], 10);
          if (n > 0 && n <= 50) return n;  // Plausi-Cap
        }
      }
      return null;
    }

    // Strategie a: pro Segment versuchen wir, n_target aus den profile_quoten
    // aller Block-Stellvertreter zu summieren
    for (const [, info] of segCount.entries()) {
      let summe = 0;
      let alleHabenN = true;
      for (const idi of info.idis) {
        const n = extractTNCountFromProfileQuoten(idi);
        if (n == null) { alleHabenN = false; break; }
        summe += n;
      }
      if (alleHabenN && summe > 0) {
        info.extractedN = summe;
      }
    }

    // Strategie b: Fallback per Brutto-Verteilung
    const sumCount = [...segCount.values()].reduce((s, v) => s + v.count, 0);
    const brutto = gruppe.brutto || gruppe.tnBrutto || 0;
    const useBruttoFallback = (!segCount.values().next().value?.extractedN)
      && brutto > sumCount && segCount.size > 0;
    let bruttoPerSeg = 0;
    if (useBruttoFallback) {
      bruttoPerSeg = Math.floor(brutto / segCount.size);
    }

    let i = 0;
    const out = [];
    for (const [segName, info] of segCount.entries()) {
      const idSuffix = info.sample.segment_nr != null
        ? `S${info.sample.segment_nr}`
        : `S${i + 1}`;
      // n_target: bevorzuge extractedN, dann Brutto-Verteilung, dann count
      const nTarget = info.extractedN != null
        ? info.extractedN
        : (useBruttoFallback ? bruttoPerSeg : info.count);
      out.push({
        id: `${gruppe.id}.${idSuffix}`,
        label: buildQGLabel({ segment: segName, ...info.sample.demographics }),
        n_target: nTarget,
        demographics: {
          alter_min: info.sample.alter_min ?? null,
          alter_max: info.sample.alter_max ?? null,
          geschlecht: info.sample.geschlecht ?? null,
          standort:   gruppe.standort ?? null,
        },
        zuordnungs_regel: '',
        segment: segName,
        cluster: info.sample.cluster || null,
        segment_nr: info.sample.segment_nr ?? null,
      });
      i++;
    }
    // Plausi-Log
    if (out.length > 0 && !globalThis.__QG_DEBUG__) globalThis.__QG_DEBUG__ = [];
    if (out.length > 0) {
      globalThis.__QG_DEBUG__.push({
        _meta: 'idi_profile_qg', gruppe: gruppe.id,
        segCount: out.length,
        nTargetSource: out[0].n_target === segCount.values().next().value.extractedN
          ? 'profile_quoten' : (useBruttoFallback ? 'brutto_division' : 'idi_count'),
        nTargets: out.map(q => ({ label: q.label, n: q.n_target })),
        sumNTarget: out.reduce((s, q) => s + q.n_target, 0),
        brutto: brutto,
      });
    }
    if (out.length > 0) return out;
  }
  // 3) Fallback: 1 QG aus gruppe.zielgruppe
  return [{
    id: `${gruppe.id}.A`,
    label: buildQGLabel({
      zielgruppe: gruppe.zielgruppe,
      alter_min: gruppe.alter_min,
      alter_max: gruppe.alter_max,
      geschlecht: gruppe.geschlecht,
    }),
    n_target: gruppe.brutto || 8,
    demographics: {
      alter_min: gruppe.alter_min ?? null,
      alter_max: gruppe.alter_max ?? null,
      geschlecht: gruppe.geschlecht ?? null,
      standort:   gruppe.standort ?? null,
    },
    zuordnungs_regel: gruppe.zuordnungs_kriterien || '',
  }];
}

// Normalisiert einen QG-Eintrag aus dem neuen Schema (defensiv).
function normalizeQG(qg, gruppe, idx) {
  const id = qg.id || `${gruppe.id}.${String.fromCharCode(65 + idx)}`;
  const label = qg.label || buildQGLabel(qg.demographics || qg) || id;
  return {
    id,
    label,
    n_target: parseInt(qg.n_target, 10) || 0,
    demographics: qg.demographics || {
      alter_min: qg.alter_min ?? null,
      alter_max: qg.alter_max ?? null,
      geschlecht: qg.geschlecht ?? null,
      standort:   qg.standort ?? gruppe.standort ?? null,
    },
    zuordnungs_regel: qg.zuordnungs_regel || '',
    segment: qg.segment || null,
    cluster: qg.cluster || null,
    segment_nr: qg.segment_nr ?? null,
  };
}

// Baut ein lesbares Klartext-Label aus demographischen Attributen.
// Format C (Konzept-Doku): "weiblich, Heavy User" / "männlich, 30-50, Heavy User"
// Keine technischen Präfixe wie "QG-A". Wenn Segment vorhanden, das nutzen.
function buildQGLabel(attrs) {
  if (!attrs) return '';
  const parts = [];
  // Segment hat Vorrang
  if (attrs.segment) {
    parts.push(String(attrs.segment));
  } else if (attrs.zielgruppe && attrs.zielgruppe !== 'Allgemein') {
    parts.push(String(attrs.zielgruppe));
  }
  // Geschlecht
  if (attrs.geschlecht && attrs.geschlecht !== 'gemischt') {
    parts.push(String(attrs.geschlecht));
  }
  // Alter
  if (attrs.alter_min || attrs.alter_max) {
    const lo = attrs.alter_min;
    const hi = attrs.alter_max;
    if (lo && hi) parts.push(`${lo}-${hi}`);
    else if (lo)  parts.push(`${lo}+`);
    else if (hi)  parts.push(`bis ${hi}`);
  }
  // Dedupe und join
  const seen = new Set();
  return parts.filter(p => {
    const k = p.toLowerCase();
    if (seen.has(k)) return false;
    seen.add(k);
    return true;
  }).join(', ') || 'Allgemein';
}

// Schreibt das Dropdown (Data Validation), Styling und bedingte Formatierung
// für die QG-Spalte zwischen tnStart und tnEnd.
// qgLabels ist die Liste aller QG-Labels dieses Sheets (für die Dropdown-Werte).
//
// v12.22.3: Styling angeglichen an Nachbar-Zellen (Arial 10pt, center/center).
// Bedingte Formatierung pro QG-Wert mit Farben aus der Vorlagen-Palette
// (HELLBLAU, HELLGRUEN, GELB, GRUEN — die "freundlichen" Farben aus COLORS,
// nicht ROT/DUNKELROT die Screenout-Bedeutung haben).
// v12.22.8: Bei nur 1 QG im Sheet wird die Spalte NICHT mehr ausgeblendet
// (das kollabierte den Logo-Bereich und verzerrte das Vorlagen-Logo).
// Stattdessen wird die Spalte grau eingefärbt und gesperrt — visuelles Signal
// "hier nicht genutzt", der User kann sie bei Bedarf manuell löschen.
function grayOutQuotengruppenSpalte(ws, qgCol, headerRow, tnEnd) {
  if (!qgCol) return;
  // Header-Zelle bleibt sichtbar, aber leicht grau hinterlegt
  const headerCell = ws.getCell(headerRow, qgCol);
  headerCell.fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FFE0E0E0' },  // helles Grau (analog COLORS.GRAU)
  };
  headerCell.font = {
    name: 'Arial', size: 9, bold: true, color: { argb: 'FF999999' },
  };
  headerCell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
  // TN-Zellen: grau einfärben, locked, kein Dropdown
  for (let r = (headerRow + 1); r <= tnEnd; r++) {
    const cell = ws.getCell(r, qgCol);
    cell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FFEEEEEE' },  // sehr helles Grau, signalisiert "leer/inaktiv"
    };
    cell.protection = { locked: true };
    cell.border = {
      top:    { style: 'thin', color: { argb: 'FFCCCCCC' } },
      bottom: { style: 'thin', color: { argb: 'FFCCCCCC' } },
      left:   { style: 'thin', color: { argb: 'FFCCCCCC' } },
      right:  { style: 'thin', color: { argb: 'FFCCCCCC' } },
    };
  }
  // Sheet-Protection aktivieren mit minimalen Restriktionen:
  // - Nur die grauen QG-Zellen sind locked
  // - Alle anderen Zellen sind editierbar (Excel-Default ist locked=false bei
  //   Zellen ohne explizite Setzung; protection auf Sheet-Ebene aktiviert das nur)
  // ACHTUNG: Protection global zu aktivieren würde die Spalten-Lösch-Funktion
  // ggf. blockieren. Daher protection NICHT aktivieren — die Zellen sind nur
  // visuell grau, der User kann sie manuell löschen.
}


function applyQuotengruppenDropdown(ws, qgCol, qgLabels, tnStart, tnEnd) {
  if (!qgCol || !Array.isArray(qgLabels) || qgLabels.length === 0) return;
  // Data Validation Liste: Excel erwartet "WERT1,WERT2,..." in Quotes.
  // Kommas in Labels würden die Liste zerreißen — deshalb ersetzen wir Kommas
  // in Labels durch "/", was im Konzept ohnehin als Trennzeichen vorgesehen ist.
  // Maximale Listenlänge in Excel: 255 Zeichen für Inline-Listen.
  const sanitizedLabels = qgLabels.map(l => String(l).replace(/,/g, ' /'));
  const listStr = sanitizedLabels.join(',');
  let formulae;
  if (listStr.length <= 250) {
    // Kurz genug für Inline-Liste
    formulae = [`"${listStr}"`];
  } else {
    // v12.22.10: Lange Listen -> Hilfsbereich rechts neben Sheet anlegen und
    // per Range-Reference einbinden. Excel-Limit fuer Inline-Listen umgehen.
    // Wir nehmen Spalte 250+ (weit rechts, außerhalb sichtbarer Bereich),
    // schreiben dort die Labels untereinander und referenzieren absolut.
    const helperCol = 250;  // Spalte mit Index 250 ist sehr weit rechts (~IP)
    const helperColLetter = columnNumberToLetter(helperCol);
    for (let i = 0; i < sanitizedLabels.length; i++) {
      const helperRow = i + 1;
      const cell = ws.getCell(helperRow, helperCol);
      cell.value = sanitizedLabels[i];
      // Optional: Hilfszellen ausblenden via white-on-white wäre möglich,
      // aber die Spalte ist außerhalb des Druckbereichs und visuell unauffällig.
    }
    formulae = [`=$${helperColLetter}$1:$${helperColLetter}$${sanitizedLabels.length}`];
  }

  for (let r = tnStart; r <= tnEnd; r++) {
    const cell = ws.getCell(r, qgCol);
    cell.dataValidation = {
      type: 'list',
      allowBlank: true,
      formulae: formulae,
      showErrorMessage: true,
      errorStyle: 'warning',
      errorTitle: 'Quotengruppe',
      error: 'Bitte einen Eintrag aus der Liste wählen.',
      showInputMessage: true,
      promptTitle: 'Quotengruppe',
      prompt: 'Welche Quotengruppe nach Screening?',
    };
    // v12.22.3: Styling angeglichen an Nachbar-Zellen (Arial 10pt, center/center)
    cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
    cell.font = { name: 'Arial', size: 10 };
    cell.border = {
      top:    { style: 'thin', color: { argb: 'FFCCCCCC' } },
      bottom: { style: 'thin', color: { argb: 'FFCCCCCC' } },
      left:   { style: 'thin', color: { argb: 'FFCCCCCC' } },
      right:  { style: 'thin', color: { argb: 'FFCCCCCC' } },
    };
  }

  // v12.22.3: Bedingte Formatierung pro QG-Wert.
  // Farben aus der Vorlagen-Palette (COLORS), rotierend bei >4 QGs.
  // Bewusst NICHT genutzt: ROT/DUNKELROT (Screenout-Semantik), GRAU (Inaktiv).
  const QG_PALETTE = [
    COLORS.HELLBLAU,   // 1. QG
    COLORS.HELLGRUEN,  // 2. QG
    COLORS.GELB,       // 3. QG
    COLORS.GRUEN,      // 4. QG
    COLORS.ORANGE,     // 5. QG (Fallback bei mehr als 4 — orange ist heller als rot)
  ];
  const qgColLetter = columnNumberToLetter(qgCol);
  const range = `${qgColLetter}${tnStart}:${qgColLetter}${tnEnd}`;
  const rules = sanitizedLabels.map((label, idx) => {
    const color = QG_PALETTE[idx % QG_PALETTE.length];
    // Excel-Formel: Doppelte Quotes für eingebettete Quotes im Label
    const escapedLabel = String(label).replace(/"/g, '""');
    return {
      type: 'expression',
      priority: idx + 1,
      formulae: [`$${qgColLetter}${tnStart}="${escapedLabel}"`],
      style: {
        fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: color } },
      },
    };
  });
  ws.addConditionalFormatting({ ref: range, rules });
}

// v12.22.20: renderQuotenUebersicht + renderQuotenUebersichtHorizontal
// entfernt. Beide Funktionen waren seit v12.22.15 (Wechsel auf
// writeQgBoxUnderTn unter dem TN-Bereich in Spalte G) dead code.

function columnNumberToLetter(n) {
  let s = '';
  while (n > 0) {
    const rem = (n - 1) % 26;
    s = String.fromCharCode(65 + rem) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
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

// v12.21.7: addLogo zurueck zur Original-Logik (vor unserer Session).
// Der Builder fuegt sein eigenes Logo (aus logos.js) ein mit fester
// Pixelgroesse und editAs:'oneCell' = "Move but don't size with cells".
// Das ist robust gegen Spaltenbreiten-Aenderungen — egal wie breit die
// IDI-Termin-Spalte wird, das Logo behaelt seine Groesse.
//
// Das Logo der Vorlage selbst wird von copyWorksheet NICHT mitkopiert
// (siehe dort), sonst haetten wir zwei Logos.
// v12.22.5: addLogo ist jetzt Fallback. Nur wenn copyWorksheet kein Vorlagen-
// Logo übernommen hat (z.B. weil die Vorlage gar keines hat oder die Übernahme
// fehlgeschlagen ist), wird ein eigenes Logo mit festen Pixelmaßen aus
// LOGOS[unternehmen] eingefügt.
function addLogo(dstWs, dstWorkbook, unternehmen, setting) {
  // Wenn Vorlagen-Logo bereits übernommen wurde: nichts tun
  if (dstWs._templateLogoCopied) {
    return;
  }
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

// Findet dynamisch die "lfd. Nr."-Spalte im Template (HEADER_ROW = Z7).
// Offline = Spalte E (5), Online = Spalte F (6) — variiert je nach Template.
function findLfdNrCol(ws, fallback) {
  const headerRow = ws.getRow(HEADER_ROW);
  let foundCol = null;
  headerRow.eachCell({ includeEmpty: false }, (cell, colNum) => {
    if (foundCol !== null) return;
    const v = String(cell.value ?? '').trim().toLowerCase();
    if (v === 'lfd. nr.' || v === 'lfd.nr.' || v === 'lfd nr.' || v.startsWith('lfd')) {
      foundCol = colNum;
    }
  });
  return foundCol || fallback;
}

// v11.4: Findet Spalten-Indizes für Datum / Uhrzeit-Spalten in HEADER_ROW (Z7)
// Liefert {datum, uhrzeit} mit Spalten-Indizes (1-basiert) oder null
function findDatumUhrzeitCols(ws) {
  const headerRow = ws.getRow(HEADER_ROW);
  let datumCol = null, uhrzeitCol = null;
  headerRow.eachCell({ includeEmpty: false }, (cell, colNum) => {
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
// v12.14: Quote-Text pro Gruppe filtern
// Wenn ein Quotenkommentar mehrere Gruppen-spezifische Zeilen enthält (z.B.
// "GD1+3: hohes C Nutzer. GD2+4: Ehemalige. GD5+6: Mischung"),
// behält diese Funktion nur die Zeilen, die für die aktuelle Gruppe gelten.
// Zeilen ohne Gruppen-Erwähnung gelten als allgemein und bleiben drin.
//
// v12.15: Auch innerhalb eines Satzes wird Komma/Semikolon-getrennt gefiltert.
// Beispiel: "50% Nutzer (GD1+3), 50% ehemalige Nutzer (GD2+4)" wird bei GD1
// reduziert zu "50% Nutzer (GD1+3)".
//
// v12.20.1 (universal): zusätzlich Zielgruppen-Substring-Erkennung. Wenn ein
// Satz die zielgruppe einer ANDEREN Gruppe nennt UND nicht die eigene, wird
// er weggefiltert. So funktioniert Filtering ohne explizite Gruppen-IDs in
// den Hinweisen — der Parser kann zielgruppenspezifische Hinweise als
// einfachen Klartext schreiben (z.B. "got2b-Nutzer: ... / Nicht-Nutzer: ...").
//
// Erkennt verschiedene Schreibweisen:
//   - "GD1, GD2, GD3" oder "GD1, 2, 3"
//   - "GD1+3" oder "GD1+GD3" oder "GD1 und GD3"
//   - "GD1-GD3" (Bereich)
//   - "got2b-Nutzer:", "Active Buyers:" (Zielgruppen-Substring aus gruppe.zielgruppe)
function filterQuoteByGruppe(text, gruppeId, allGruppen) {
  if (!text) return '';
  // v12.15.2: Normalisiere "GD 4", "GD  4" → "GD4" damit das Filter-Regex matched.
  text = String(text).replace(/\b(GD|IDI|VGD|VDI)\s+(\d)/gi, '$1$2');
  const targetUpper = String(gruppeId || '').toUpperCase();
  const targetMatch = targetUpper.match(/^(GD|IDI|VGD|VDI)(\d+)$/i);

  // v12.20.1: Zielgruppen-Map aus allGruppen aufbauen.
  // Ignoriere generische Zielgruppen ("Allgemein", "Mix", leer) — die sind
  // kein semantischer Filter.
  function isGenericZG(zg) {
    if (!zg) return true;
    const z = String(zg).toLowerCase().trim();
    if (!z) return true;
    return /^(allgemein|mix|verschiedene|alle)\b/.test(z) || z.length < 4;
  }
  const ownGruppe = Array.isArray(allGruppen)
    ? allGruppen.find(g => String(g.id || '').toUpperCase() === targetUpper)
    : null;
  const ownZG = (ownGruppe && !isGenericZG(ownGruppe.zielgruppe))
    ? String(ownGruppe.zielgruppe).trim() : '';
  const otherZGs = Array.isArray(allGruppen)
    ? Array.from(new Set(
        allGruppen
          .filter(g => String(g.id || '').toUpperCase() !== targetUpper)
          .map(g => String(g.zielgruppe || '').trim())
          .filter(z => !isGenericZG(z) && z !== ownZG)
      ))
    : [];

  // Substring-Match einer Zielgruppe in einem Satz — case-insensitive,
  // normalisiert Trennzeichen (Hyphen, Slash, Space → gleichbehandelt).
  // So matched "got2b-Nutzer" auch wenn Sonnet "got2b Nutzer" schreibt.
  function normalizeForZG(s) {
    return String(s || '').toLowerCase().replace(/[\s\-\/]+/g, ' ').trim();
  }
  function zgMatch(text, zg) {
    if (!zg) return false;
    const tNorm = normalizeForZG(text);
    const zNorm = normalizeForZG(zg);
    if (!zNorm) return false;
    // 1) Voll-Match normalisiert (löst Schreibvarianten Hyphen↔Space)
    if (tNorm.includes(zNorm)) return true;
    // 2) Token-Fallback: einzelne markante Wörter (≥5 chars), die NICHT in
    //    ownZG vorkommen — sonst False-Positive bei Subset-Beziehung.
    const tokens = zNorm.split(' ').filter(t => t.length >= 5);
    const ownNorm = normalizeForZG(ownZG);
    for (const tok of tokens) {
      if (ownNorm && ownNorm.includes(tok)) continue;
      if (tNorm.includes(tok)) return true;
    }
    return false;
  }

  // Prüft ob ein Satz/Sub den Zielgruppen-Filter passiert.
  // - Wenn fremde ZG erwähnt UND eigene NICHT erwähnt → false (raus)
  // - Sonst → true (behalten)
  function zgFilterPass(s) {
    if (otherZGs.length === 0) return true; // kein Zielgruppen-Filter möglich
    const ownMentioned = ownZG ? zgMatch(s, ownZG) : false;
    const fremdMentioned = otherZGs.some(zg => zgMatch(s, zg));
    if (fremdMentioned && !ownMentioned) return false;
    return true;
  }

  // v12.21.2: Satz-Splitter, der Abkürzungen wie "mind.", "z.B.", "ca." nicht
  // als Satzende behandelt. Splittet nur an .!?-Stellen, an denen NICHT direkt
  // davor eine bekannte Abkürzung steht.
  function splitSentences(text) {
    const ABBREV = /\b(z\.\s*B|u\.\s*a|d\.\s*h|i\.\s*d\.\s*R|mind|bzw|etc|ca|inkl|exkl|usw|evtl|ggf|max|min|Nr|Tel|Mr|Mrs|Dr|St|vs|Co|GmbH|AG)\.?$/i;
    const out = [];
    let buf = '';
    // Erst auf Newlines splitten — die sind eindeutige Trennungen
    for (const line of String(text).split(/\n+/)) {
      // Dann auf .!? + Whitespace, aber Abkürzungs-Lookback prüfen
      const parts = line.split(/([.!?])\s+/);
      let acc = '';
      for (let i = 0; i < parts.length; i++) {
        const p = parts[i];
        if (p === '.' || p === '!' || p === '?') {
          acc += p;
          // Prüfen ob acc auf Abkürzung endet
          if (ABBREV.test(acc)) {
            acc += ' '; // Abkürzung — Satz geht weiter
          } else {
            const s = acc.trim();
            if (s) out.push(s);
            acc = '';
          }
        } else {
          acc += p;
        }
      }
      if (acc.trim()) out.push(acc.trim());
    }
    return out.filter(Boolean);
  }

  // Wenn keine Gruppen-ID-Match (z.B. bei konsolidierter "IDI"-Gruppe ohne Nummer),
  // läuft die ID-basierte Filterung nicht — aber Zielgruppen-Filter trotzdem.
  if (!targetMatch) {
    if (otherZGs.length === 0) return text;
    const sentences = splitSentences(text);
    const kept = sentences.filter(s => zgFilterPass(s));
    return kept.join(' ').replace(/\s+/g, ' ').trim();
  }
  const targetPrefix = targetMatch[1].toUpperCase();
  const targetNum = parseInt(targetMatch[2], 10);

  // In Sätze splitten (v12.21.2: abkürzungs-aware)
  const sentences = splitSentences(text);

  // Extrahiert ALLE Gruppen-Nummern aus einem Satz für den gegebenen Prefix.
  // Handhabt: "GD1, GD2", "GD1+3", "GD1, 2, 3", "GD2-4" (Bereich)
  function extractGroupNumbers(sentence, prefix) {
    const numbers = new Set();
    const upper = sentence.toUpperCase();
    // Schritt 1: Alle Vorkommen von PREFIX + Zahl als Anker finden
    const re = new RegExp('\\b' + prefix + '(\\d+)', 'g');
    let m;
    while ((m = re.exec(upper)) !== null) {
      const baseNum = parseInt(m[1], 10);
      numbers.add(baseNum);
      // Schaue ob nach der Zahl direkt +/-/und/,/Zahl kommt (Kettung)
      let pos = m.index + m[0].length;
      const rest = upper.substring(pos);
      const fm = rest.match(/^[\s]*([+,\-]|UND)[\s]*(\d+)(?:[\s]*(?:[+,\-]|UND)[\s]*(\d+))*/i);
      if (fm) {
        const nums = fm[0].match(/\d+/g) || [];
        if (fm[0].includes('-') && nums.length === 1) {
          const end = parseInt(nums[0], 10);
          for (let i = Math.min(baseNum, end); i <= Math.max(baseNum, end); i++) {
            numbers.add(i);
          }
        } else {
          for (const n of nums) numbers.add(parseInt(n, 10));
        }
      }
    }
    return numbers;
  }

  const keptLines = [];
  for (const s of sentences) {
    // v12.20.1: Zielgruppen-Filter zuerst — wenn fremde ZG ohne eigene erwähnt,
    // raus. Wenn passt, weiter zur ID-Logik.
    if (!zgFilterPass(s)) continue;

    // Erst grobe Prüfung: hat dieser Satz ÜBERHAUPT Gruppen-Erwähnungen?
    const anyMention = /\b(GD|IDI|VGD|VDI)\d+\b/i.test(s);
    if (!anyMention) {
      keptLines.push(s);
      continue;
    }

    // v12.15: Wenn Satz an Komma/Semikolon getrennt mehrere Gruppen-spezifische
    // Sub-Klauseln hat, jede Sub-Klausel einzeln filtern.
    const subParts = splitOnCommaIfMultiGroup(s);
    if (subParts.length > 1) {
      const keptSubs = [];
      for (const sub of subParts) {
        // v12.20.1: ZG-Filter auch auf Sub-Ebene
        if (!zgFilterPass(sub)) continue;
        const subHasMention = /\b(GD|IDI|VGD|VDI)\d+\b/i.test(sub);
        if (!subHasMention) {
          keptSubs.push(sub);
          continue;
        }
        const subNums = extractGroupNumbers(sub, targetPrefix);
        if (subNums.has(targetNum)) {
          keptSubs.push(sub);
        }
      }
      if (keptSubs.length > 0) {
        let joined = keptSubs.join(', ');
        const lastChar = s.slice(-1);
        if (/[.!?]/.test(lastChar) && !/[.!?]$/.test(joined)) {
          joined += lastChar;
        }
        keptLines.push(joined);
      }
      continue;
    }

    // Sonst: ganzen Satz als Ganzes prüfen
    const nums = extractGroupNumbers(s, targetPrefix);
    if (nums.has(targetNum)) {
      keptLines.push(s);
    }
  }
  return keptLines.join(' ').replace(/\s+/g, ' ').trim();
}

// v12.15: Helper - splittet einen Satz an Komma/Semikolon nur wenn jede Teil
// für sich genommen "echten Text" enthält (nicht nur Gruppen-IDs).
// Damit wird "50% Nutzer (GD1+3), 50% ehemalige Nutzer (GD2+4)" in zwei Teile
// gesplittet, aber "GD1, GD2 und GD3 sind familienorientiert" bleibt ein Teil.
function splitOnCommaIfMultiGroup(s) {
  // Erstmal an , oder ; splitten
  const raw = s.split(/[,;]/).map(p => p.trim()).filter(Boolean);
  if (raw.length <= 1) return [s];

  // Jede Sub-Klausel muss mindestens 3 nicht-Gruppen-Wörter enthalten, sonst
  // ist es eher eine Aufzählung ("GD1, GD2, GD3").
  function nonGroupWordCount(t) {
    // Entferne alle Gruppen-Erwähnungen und Zahlen
    const stripped = t
      .replace(/\b(GD|IDI|VGD|VDI)\d+(?:[\s]*[+,\-]\s*\d+)*\b/gi, '')
      .replace(/[\(\)]/g, '')
      .replace(/\d+%?/g, '')
      .trim();
    const words = stripped.split(/\s+/).filter(w => w.length >= 2);
    return words.length;
  }

  const allHaveText = raw.every(p => nonGroupWordCount(p) >= 1);
  if (!allHaveText) return [s];  // Aufzählung → nicht splitten
  return raw;
}

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

  // v12.22.22 (Bug 7): Heuristik aus quoteText-String. Wenn die Quote-Anweisung
  // 'MUSS Code X' o.ae. enthaelt, sind alle anderen Codes off-target. Bei Matrix-
  // Items wird zusaetzlich die Item-Bedingung beruecksichtigt (requiredItem).
  // Konservativ: nur greifen wenn parseExclusionCodes ein konkretes Match hat.
  //
  // v12.22.23 (Bug 7 Hotfix): Heuristik dramatisch eingeschraenkt nach Test-Run
  // 26_1234_2345 (QAIUSED hatte Codes 4-15 faelschlich rot wegen 'mindestens
  // einmal im Monat' Match in quotenkommentar). Neue Regeln:
  //
  //   1. parseExclusionCodes wird NUR auf candidates angewandt, die explizit
  //      einen Code-Filter signalisieren (MUSS Code X, Range, hauptsaechlich
  //      in). NICHT auf das alleinige Pattern "mindestens einmal im Monat".
  //      Letzteres ist eine Quoten-Beschreibung, kein Code-Filter pro Antwort.
  //   2. requiredItem-Match muss ueber strukturiertes Feld (frage.item_label)
  //      laufen, nicht ueber kurz_label-Fallback (das macht z.B. die ganze
  //      Frage QAIUSED zu "Genutzte KI-Tools" → matched gegen "ChatGPT" via
  //      includes()=false, dann mismatching-Item, dann mein Code-2-Trigger).
  const candidates = [];
  // Item-bezogenes soll_quote ist die SICHERSTE Quelle (PATCH 24/25 Schema)
  if (frage && Array.isArray(frage.items)) {
    for (const item of frage.items) {
      if (item && Array.isArray(item.soll_quote)) {
        for (const sq of item.soll_quote) {
          if (sq && typeof sq.text === 'string') candidates.push({text: sq.text, source: 'item'});
        }
      }
    }
  }
  // ant.soll_quote ist ebenfalls strukturiert (PATCH 25)
  if (ant && Array.isArray(ant.soll_quote)) {
    for (const sq of ant.soll_quote) {
      if (sq && typeof sq.text === 'string') candidates.push({text: sq.text, source: 'ant'});
    }
  }
  // frage.quotenkommentar ist die LETZTE Quelle — wird nur ausgewertet wenn die
  // Frage Matrix-Items hat (also keine reine single_choice mit Tool-Labels)
  if (frage && frage.quotenkommentar && Array.isArray(frage.items) && frage.items.length > 0) {
    if (typeof frage.quotenkommentar === 'string') candidates.push({text: frage.quotenkommentar, source: 'frage'});
    else if (Array.isArray(frage.quotenkommentar)) {
      for (const qk of frage.quotenkommentar) {
        if (typeof qk === 'string') candidates.push({text: qk, source: 'frage'});
        else if (qk && typeof qk.text === 'string') candidates.push({text: qk.text, source: 'frage'});
      }
    }
  }

  for (const cand of candidates) {
    const parsed = parseExclusionCodes(cand.text);
    if (!parsed) continue;
    // Code-Nummer der aktuellen Antwort ermitteln
    const antCode = typeof ant.code === 'number' ? ant.code
                  : typeof ant.code === 'string' ? parseInt(ant.code, 10)
                  : null;
    if (antCode === null || isNaN(antCode)) continue;

    // Item-Filter — strenger als v12.22.22b: nur ueber frage.item_label,
    // KEIN Fallback auf kurz_label (gibt false positives bei single_choice).
    const itemLabel = (frage.item_label || '').toLowerCase().trim();
    let isMatchingItem = true;
    if (parsed.requiredItem) {
      if (!itemLabel) {
        // Kein Item-Kontext (single_choice / Frage-Ebene): Heuristik nur greifen
        // lassen, wenn cand.source !== 'frage' (also strukturiert pro Item/Ant)
        if (cand.source === 'frage') continue;
        isMatchingItem = true;
      } else {
        const reqItem = parsed.requiredItem.toLowerCase().trim();
        isMatchingItem = itemLabel.includes(reqItem) || reqItem.includes(itemLabel);
      }
    }

    // Logik 1: requiredCodes-Liste (exakte Codes)
    if (parsed.requiredCodes && parsed.requiredCodes.length > 0) {
      if (isMatchingItem) {
        if (!parsed.requiredCodes.includes(antCode)) return true;
      } else if (parsed.requiredItem) {
        // Bei mismatching Item: der requiredCode IST off-target
        if (parsed.requiredCodes.includes(antCode)) return true;
      }
    }
    // Logik 2: Range (minCode/maxCode) — nur greifen wenn auch ein
    // requiredItem ODER nicht aus frage.quotenkommentar (sonst zu aggressiv)
    if (parsed.minCode !== null || parsed.maxCode !== null) {
      // Konservativ: Range nur anwenden wenn ein Item-Bezug besteht oder die
      // Quelle strukturiert ist (item/ant soll_quote).
      const safeToApply = parsed.requiredItem || cand.source !== 'frage';
      if (safeToApply && isMatchingItem) {
        const min = parsed.minCode !== null ? parsed.minCode : 1;
        const max = parsed.maxCode !== null ? parsed.maxCode : 99;
        if (antCode < min || antCode > max) return true;
      }
    }
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

  // PATCH 3+6: Bedingungs-Markierung bei conditional Fragen
  // v12.13: Statt farbigem Rahmen wird der Header-HINTERGRUND in Cluster-Farbe
  // (pastel-aufgehellt) eingefärbt. Rahmen bleibt normaler header border.
  const isConditional = frage && isConditionalFrage(frage);
  let clusterColorArgb = null;
  let clusterLightArgb = null;
  if (isConditional) {
    const clusterEntry = (typeof __CLUSTER_MAP__ !== 'undefined' && __CLUSTER_MAP__ && frage && frage.id) ?
      __CLUSTER_MAP__[frage.id] : null;
    clusterColorArgb = clusterEntry ? clusterEntry.colorArgb : 'FFD4A017';
    clusterLightArgb = clusterEntry ? clusterEntry.lightArgb : 'FFFFF2CC';
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

  // v12.22.18: Volle Screener-Frage in Z6 (LONG_QUESTION_ROW). Template hat
  // Z6 fuer diese Zeile vorbereitet (zwischen MATRIX_HEADER_ROW=5 und
  // HEADER_ROW=7). Smart-Suppression:
  //   - Bei Matrix-Item-Spalten: leer, da Z6 ueber die Matrix-Spannweite
  //     vom Matrix-Loop GEMERGED + befuellt wird (v12.22.21 / Bug 4).
  //     Vorher war Z6 bei Matrix-Items komplett leer.
  //   - Wenn fragetext == kurz_label oder leer: leer
  const isMatrixItemColumn = Array.isArray(frage && frage.items) && frage.items.length > 0
    && label !== (frage.kurz_label || frage.kurzlabel)
    && label !== (frage.fragetext || '');
  if (!isMatrixItemColumn && ftClean && ftClean.length > 0
      && ftClean.toLowerCase() !== krzClean.toLowerCase()) {
    const longQCell = ws.getCell(LONG_QUESTION_ROW, col);
    longQCell.value = ftClean;
    longQCell.font = { name: 'Arial', size: 9, italic: true, color: { argb: 'FF555555' } };
    longQCell.alignment = { horizontal: 'left', vertical: 'top', wrapText: true };
    longQCell.border = { ...THIN_BORDER };
  }

  const prefix = isConditional ? '🔀 ' : '';
  // v12.22.21 (Bug 3): Label wird ohne Q-Code-Präfix angezeigt — buildHeaderLabel
  // hat das bereits gestrippt. Hier zusätzliche Sicherung für Labels, die nicht
  // durch buildHeaderLabel gegangen sind (Matrix-Items, summaryLabel etc.):
  const cleanLabel = stripQCode(label);

  // v12.22.22 (Bug 6): Quotengruppen-Marker fuer quoten-relevante Matrix-Items.
  // Wenn diese Spalte das requiredItem der Quote-Anweisung ist (z.B. ChatGPT-
  // Spalte bei QAIFREQUENCY mit 'als ChatGPT User qualifizieren'), bekommt sie
  // ein 📌 + amber Header-Hintergrund, sodass der Recruiter sofort erkennt
  // dass hier eine Quotengruppe entsteht.
  //
  // v12.22.23: Wie Bug 7 — nur strukturierte Quellen (item.soll_quote,
  // ant.soll_quote, frage.quotenkommentar nur bei Matrix). frage.bedingung
  // entfaellt (das ist nur "Nur wenn QXY = ..." und kein Item-Marker).
  let isQuotaRelevantColumn = false;
  if (Array.isArray(frage && frage.items) && cleanLabel) {
    const candidates = [];
    if (quoteText) candidates.push(String(quoteText));
    if (frage.quotenkommentar) {
      if (typeof frage.quotenkommentar === 'string') candidates.push(frage.quotenkommentar);
      else if (Array.isArray(frage.quotenkommentar)) {
        for (const qk of frage.quotenkommentar) {
          if (typeof qk === 'string') candidates.push(qk);
          else if (qk && typeof qk.text === 'string') candidates.push(qk.text);
        }
      }
    }
    // PATCH 24-Schema: items[].soll_quote enthaelt das Item-bezogene Min/Max
    if (frage.items) {
      for (const item of frage.items) {
        if (item && Array.isArray(item.soll_quote)) {
          for (const sq of item.soll_quote) {
            if (sq && typeof sq.text === 'string') candidates.push(sq.text);
          }
        }
      }
    }
    if (frage.quotenkommentar_pro_zielgruppe && gruppe && gruppe.zielgruppe) {
      const perZG = frage.quotenkommentar_pro_zielgruppe[gruppe.zielgruppe];
      if (typeof perZG === 'string') candidates.push(perZG);
    }
    const labelLower = cleanLabel.toLowerCase().trim();
    for (const cand of candidates) {
      const parsedHeader = parseExclusionCodes(cand);
      if (parsedHeader && parsedHeader.requiredItem) {
        const reqItem = parsedHeader.requiredItem.toLowerCase().trim();
        if (labelLower.includes(reqItem) || reqItem.includes(labelLower)) {
          isQuotaRelevantColumn = true;
          break;
        }
      }
    }
  }
  const quotaPrefix = isQuotaRelevantColumn ? '📌 ' : '';

  if (showFragetextInHeader) {
    h.value = {
      richText: [
        { text: quotaPrefix + prefix + cleanLabel,      font: { bold: true,  name: 'Arial', size: 9, color: { argb: isQuotaRelevantColumn ? 'FF412402' : 'FF000000' } } },
        { text: '\n',                     font: { name: 'Arial', size: 8 } },
        { text: ftClean,                  font: { italic: true, name: 'Arial', size: 8, color: { argb: 'FF555555' } } },
      ],
    };
  } else {
    h.value = quotaPrefix + prefix + cleanLabel;
    h.font = { bold: true, name: 'Arial', size: 9, color: { argb: isQuotaRelevantColumn ? 'FF412402' : 'FF000000' } };
  }
  h.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
  // v12.22.22 (Bug 6): Amber-Hintergrund bei quoten-relevanten Spalten
  // (uberschreibt den regulaeren COLORS.HEADER_GREY / clusterLightArgb)
  const headerFill = isQuotaRelevantColumn ? 'FFFAEEDA'  // amber (c-amber 50)
                   : (isConditional ? (clusterLightArgb || 'FFFFF2CC') : COLORS.HEADER_GREY);
  h.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: headerFill } };
  h.border = headerBorder;
  // v12.22.21 (Bug 3 / User-Wunsch): Comment am Header entfernt — Z6 ist die
  // Quelle für den Fragetext. Bedingungen erscheinen weiterhin in der
  // Cluster-Legende der Quotenübersicht (Sektion 3.5).
  // (Alt: if (note) h.note = note;  — Param `note` wird ignoriert.)

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
  let cleanedQuote = cleanQuoteText(quoteText);
  // v12.14: Quote-Text pro Gruppe filtern - andere Gruppen-spezifische Anweisungen raus
  if (cleanedQuote && gruppe && gruppe.id) {
    cleanedQuote = filterQuoteByGruppe(cleanedQuote, gruppe.id, allGruppen);
  }
  const subQuoteList = buildSubQuoteList(antList, frage, gruppe, allGruppen);
  // v12.11: Bedingungs-Hinweis bei 🔀-Fragen explizit als ERSTE Zeile, damit der
  // Recruiter sofort sieht WANN diese Frage gestellt wird
  let bedingungHinweis = '';
  if (frage && frage.bedingung && String(frage.bedingung).trim()) {
    const bed = String(frage.bedingung).trim();
    // Nicht doppeln falls schon im quoteText drin
    if (!cleanedQuote || !cleanedQuote.includes(bed)) {
      bedingungHinweis = '⚠ Nur stellen wenn: ' + bed;
    }
  }
  const finalQuoteText = [bedingungHinweis, cleanedQuote, subQuoteList].filter(Boolean).join('\n\n');
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

// v12.22.15: QG-Box in Spalte G unter dem TN-Bereich rendern.
// Frueher hat diese Funktion Segment-Beschreibungen, IDI-Profile und Studien-
// Quoten dort gerendert. Auf User-Wunsch (Image 2 vom 17.05.) ersetzen wir
// das durch den QG-Counter (Quotengruppe / Soll / Ist / Status).
//
// Layout:
//   Z(tnEnd+8) Spalte G: "Quotengruppen-Counter" (Header)
//   Z(tnEnd+9): | Quotengruppe | Soll | Ist | Status |
//   Z(tnEnd+10+i): Daten-Zeile pro QG mit COUNTIF
//
// Segmente und Studien-Quoten kommen nicht mehr ins Sheet — der Recruiter
// findet sie ohnehin in der Quotenübersicht (Sheet 1).
function writeQgBoxUnderTn(ws, qgList, qgColLetter, tnStart, tnEnd) {
  if (!Array.isArray(qgList) || qgList.length < 2) return;

  const col = 7; // G
  ws.getColumn(col).width = Math.max(ws.getColumn(col).width || 20, 32);
  // Spalten H, I, J fuer Soll/Ist/Status sicherstellen
  for (let dc = 0; dc < 4; dc++) {
    if (!ws.getColumn(col + dc).width || ws.getColumn(col + dc).width < 8) {
      ws.getColumn(col + dc).width = dc === 0 ? 32 : 10;
    }
  }

  const startRow = tnEnd + 8;

  // Header
  const h1 = ws.getCell(startRow, col);
  h1.value = 'Quotengruppen-Counter';
  h1.font = { name: 'Arial', size: 11, bold: true, color: { argb: 'FF1F4E79' } };
  h1.alignment = { vertical: 'middle', horizontal: 'left' };
  // Subheader-Zeile mit Spalten-Labels
  const subRow = startRow + 1;
  const subHeaders = ['Quotengruppe', 'Soll', 'Ist', 'Status'];
  const fillHeader = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFEAF1F8' } };
  const thinGrey = { style: 'thin', color: { argb: 'FFB8C5D6' } };
  for (let i = 0; i < 4; i++) {
    const c = ws.getCell(subRow, col + i);
    c.value = subHeaders[i];
    c.font = { name: 'Arial', size: 10, bold: true, color: { argb: 'FF1F4E79' } };
    c.alignment = { vertical: 'middle', horizontal: i === 0 ? 'left' : 'center' };
    c.fill = fillHeader;
    c.border = { top: thinGrey, bottom: thinGrey, left: thinGrey, right: thinGrey };
  }

  // Daten-Zeilen
  const fillBox = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF7FAFC' } };
  qgList.forEach((qg, idx) => {
    const r = subRow + 1 + idx;
    const labelDisplay = String(qg.label).replace(/,/g, ' /');

    const labelCell = ws.getCell(r, col);
    labelCell.value = labelDisplay;
    labelCell.font = { name: 'Arial', size: 10 };
    labelCell.alignment = { vertical: 'middle', horizontal: 'left', wrapText: true };
    labelCell.fill = fillBox;
    labelCell.border = { top: thinGrey, bottom: thinGrey, left: thinGrey, right: thinGrey };

    const sollCell = ws.getCell(r, col + 1);
    sollCell.value = qg.n_target;
    sollCell.font = { name: 'Arial', size: 10, bold: true };
    sollCell.alignment = { vertical: 'middle', horizontal: 'center' };
    sollCell.fill = fillBox;
    sollCell.border = { top: thinGrey, bottom: thinGrey, left: thinGrey, right: thinGrey };

    const sanitizedLabel = labelDisplay.replace(/"/g, '""');
    const istCell = ws.getCell(r, col + 2);
    istCell.value = {
      formula: `COUNTIF(${qgColLetter}${tnStart}:${qgColLetter}${tnEnd},"${sanitizedLabel}")`,
    };
    istCell.font = { name: 'Arial', size: 10, bold: true };
    istCell.alignment = { vertical: 'middle', horizontal: 'center' };
    istCell.fill = fillBox;
    istCell.border = { top: thinGrey, bottom: thinGrey, left: thinGrey, right: thinGrey };

    const statusCell = ws.getCell(r, col + 3);
    statusCell.value = '';
    statusCell.alignment = { vertical: 'middle', horizontal: 'center' };
    statusCell.fill = fillBox;
    statusCell.border = { top: thinGrey, bottom: thinGrey, left: thinGrey, right: thinGrey };
  });

  // Bedingte Formatierung pro Status-Zelle (Ampel)
  qgList.forEach((qg, idx) => {
    const r = subRow + 1 + idx;
    const colLetter = columnNumberToLetter(col + 3);
    const sollRef = `${columnNumberToLetter(col + 1)}${r}`;
    const istRef  = `${columnNumberToLetter(col + 2)}${r}`;
    const statusRef = `${colLetter}${r}`;
    ws.addConditionalFormatting({
      ref: statusRef,
      rules: [
        { type: 'expression', priority: 1, formulae: [`${istRef}=${sollRef}`],
          style: { fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: 'FFA9D08E' } } } },
        { type: 'expression', priority: 2, formulae: [`${istRef}>${sollRef}`],
          style: { fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: 'FFFFC7CE' } } } },
        { type: 'expression', priority: 3, formulae: [`AND(${istRef}>0,${istRef}<${sollRef})`],
          style: { fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: 'FFFFEB9C' } } } },
      ],
    });
  });
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
  // v12.22.9: praeziser — nur die Hauptfrage "Wie alt sind Sie?" / "Alter" /
  // "Altersgruppe" matcht. Folgefragen wie "Alter der Kinder", "Alter der
  // Eltern", "Alter beim ersten X" matchen NICHT (kommen in Screener-Reihenfolge
  // nach ihrer Mutter-Frage wie F10 "Kinder im Haushalt").
  const REGEX_ALTER = /^(f\s*\d+\.?\s*)?alter\b(?!\s+(der|von|beim?))|wie alt sind sie|altersgruppe/i;

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

// v12.22.9: Studien-Teilnahme-Historie-Filter (defense in depth zu PATCH 20).
// Auch wenn der Parser PATCH 20 nicht befolgt und so eine Frage durchschleift,
// filtert der Builder sie zusaetzlich. Erkennt Fragen wie "F6 Letzte Studie",
// "Wann zuletzt an Marktforschung teilgenommen", "Anzahl Teilnahmen letzte X Monate".
// Diese Fragen gehoeren in den allgemeinen MaFo-Block (Spalte M Anonymitaet/
// Aufzeichnung), nicht als eigene Spalte in den TN-Bereich.
function isStudienTeilnahmeFrage(f) {
  if (!f) return false;
  const txt = ((f.fragetext || '') + ' ' + (f.kurz_label || '')).toLowerCase();
  if (!txt.trim()) return false;
  // Muster: "Teilnahme" oder "teilgenommen" oder "Letzte Studie" in Kombination
  // mit "Studie" / "Befragung" / "Marktforschung" / "Umfrage"
  const PAT = /(letzte[rs]?\s+(studie|befragung|marktforschung|umfrage|teilnahme)|zuletzt\s+(an einer|teilgenommen)|teilgenommen.*(studie|befragung|marktforschung|umfrage)|(studie|befragung|marktforschung|umfrage).*teilgenommen|wie oft.*(studie|befragung|teilgenommen)|anzahl.*teilnahmen|teilnahme[ -]?historie)/i;
  return PAT.test(txt);
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
  // v12.22.21 (Bug 3): Q-Code-Präfix aus dem kurz_label entfernen
  // (Deutsch-only Label gewünscht: 'QGENDER. Geschlecht' → 'Geschlecht')
  kurz = stripQCode(kurz);
  if (kurz) {
    // Wenn kurz_label schon mit der ID startet, nicht doppelt
    if (id && !kurz.toLowerCase().startsWith(id.toLowerCase())) {
      return `${id}. ${kurz}`.substring(0, 40);
    }
    return kurz.substring(0, 40);
  }
  // v12.22.21 (Bug 3): Wenn nach stripQCode nichts übrig (z.B. kurz_label war
  // nur 'QEMPLOYMENT.'), Fallback auf compactKurzLabel(fragetext) statt
  // wieder die englische Q-ID anzuzeigen.
  if (lang) {
    const fromLang = stripQCode(compactKurzLabel(lang));
    if (fromLang) {
      if (id && !fromLang.toLowerCase().startsWith(id.toLowerCase())) {
        return `${id}. ${fromLang}`.substring(0, 40);
      }
      return fromLang.substring(0, 40);
    }
    if (id) return `${id}. ${lang}`.substring(0, 40);
    return lang.substring(0, 40);
  }
  return (lang || id).substring(0, 40);
}

// v12.22.21 (Bug 3): Entfernt führenden 'Qxxxxx.'-Präfix aus einem Label
// und erhält dabei conditional-marker-Präfixe (🔀, ●).
// Beispiele:
//   'QGENDER. Geschlecht'              → 'Geschlecht'
//   'QAGE. Alter'                      → 'Alter'
//   '🔀 QOCCUPATION.'                  → '🔀' (leer dahinter — Caller-Fallback)
//   'QFAMILIARITY. Vertrautheit Tech.' → 'Vertrautheit Tech.'
//   'STATE_DE. Bundesland'             → 'Bundesland'
//   'Q1. Geschlecht'                   → 'Q1. Geschlecht' (echte deutsche Frage-Nr. bleibt!)
// Wird NICHT angewandt auf F-Codes (F1, F2) oder reine Q-Nr. (Q1, Q2a).
function stripQCode(label) {
  if (label == null) return '';
  let s = String(label).trim();
  if (!s) return '';
  // Conditional-/Cluster-Marker am Anfang erhalten
  let prefix = '';
  const markers = ['🔀 ', '● ', '⚠ '];
  for (const m of markers) {
    if (s.startsWith(m)) { prefix = m; s = s.substring(m.length); break; }
  }
  // Strip nur ECHTE Variablen-Codes: Q + ≥2 Großbuchstaben/Unterstriche/Ziffern + Punkt
  // Beispiele die matchen: QGENDER. QAGE. QFAMILIARITY. STATE_DE. QINCOME_DE.
  // Beispiele die NICHT matchen: Q1. Q2a. F1. (echte Frage-Nummern bleiben)
  const stripped = s.replace(/^(?:Q|STATE_DE|[A-Z]+_[A-Z]+)[A-Z_][A-Z0-9_]*\.\s*/, '');
  // Wenn nichts gestripped wurde (Frage-Nr-Pattern wie 'Q1.'), Original zurück
  if (stripped === s) return prefix + s;
  return (prefix + stripped).trim();
}

// v12.22.22 (Bug 7): Heuristik fuer ENDE-Codes aus Quote-Text-Strings.
// Liest Muster wie 'MUSS Code 02', 'nur Code 1 erlaubt', 'mindestens einmal im
// Monat' und gibt strukturierte Markierungen zurueck.
//
// Rueckgabe: { requiredCodes: number[], requiredItem: string | null,
//              minCode: number | null, maxCode: number | null }
//
// requiredCodes : Codes die OK sind (alle anderen → ENDE)
// requiredItem  : Wenn die Quote nur fuer EIN Matrix-Item gilt (z.B. "ChatGPT-
//                 User"), Name des Items. Sonst null = gilt fuer alle Items.
// minCode/maxCode: Bei Range-Patterns ('mindestens einmal im Monat' → Code <=3)
//
// Konservativ: lieber leeres Ergebnis als falsches Match. Liefert null wenn
// nichts erkannt wurde.
function parseExclusionCodes(quoteText) {
  if (!quoteText || typeof quoteText !== 'string') return null;
  const s = quoteText;
  const result = { requiredCodes: [], requiredItem: null, minCode: null, maxCode: null };
  let matched = false;

  // 1) "MUSS Code X" / "MUSS C0X" / "nur Code X"
  const reqMatch = s.match(/(?:MUSS|MUSS\s+nur|nur)\s*C(?:ode?)?\s*0?(\d{1,2})/i);
  if (reqMatch) {
    result.requiredCodes.push(parseInt(reqMatch[1], 10));
    matched = true;
  }

  // 2) Range: "1-3", "Code 1-3", "MUSS 1-3 wählen", "1-3 wählen"
  const rangeMatch = s.match(/(?:Code\s+|MUSS\s+|\b)0?(\d{1,2})\s*[-\u2013\u2014]\s*0?(\d{1,2})\s*(?:w[äa]hlen)?/i);
  if (rangeMatch) {
    const lo = parseInt(rangeMatch[1], 10), hi = parseInt(rangeMatch[2], 10);
    if (lo <= hi && hi <= 99) {
      result.minCode = lo;
      result.maxCode = hi;
      matched = true;
    }
  }

  // 3) "mindestens einmal im Monat" → maxCode=3 (Häufigkeits-Skala)
  // v12.22.23: Nur wenn der String NICHT Wort "qualifizieren" oder "ENDE" enthaelt
  // — sonst ist es eine Frage-weite Qualifizierungsanweisung ("muss mind. einmal
  // im Monat ein Tool nutzen um sich zu qualifizieren"), kein Code-Filter pro
  // Antwort. False positives bei QAIUSED-Frage wo dieser Satz im
  // quotenkommentar steht aber nicht die Skala der Frage betrifft.
  if (/mindestens\s+(?:einmal\s+)?(?:pro|im|am)\s+(?:Tag|Woche|Monat)/i.test(s)
      && !/qualifizieren|ENDE\b|HOLD\b/i.test(s)) {
    result.maxCode = 3;
    matched = true;
  }

  // 4) Item-Bezug — mehrere Patterns, erstes Match gewinnt:

  // 4a) "hauptsaechlich in X" → impliziert requiredCode=1 + item=X
  if (!result.requiredItem) {
    const hauptMatch = s.match(/haupts[äa]chlich\s+in\s+(?:der\s+|die\s+|das\s+|den\s+|dem\s+)?([A-Z][A-Za-zÄÖÜäöüß0-9 -]{2,30})/i);
    if (hauptMatch) {
      if (!result.requiredCodes.includes(1)) result.requiredCodes.push(1);
      let item = hauptMatch[1].trim().replace(/\s+(?:ein|kauft|kaufen|nutzen|verwenden).*$/i, '').trim();
      if (item) { result.requiredItem = item; matched = true; }
    }
  }

  // 4b) "X User: Min" oder "X-User: Min" — Doppelpunkt-Notation aus Quote-Listen
  if (!result.requiredItem) {
    const userColon = s.match(/([A-Z][A-Za-zÄÖÜäöüß0-9-]+(?:\s+[A-Za-zÄÖÜäöüß0-9-]+)?)\s+-?\s*User\s*:\s*Min/);
    if (userColon) { result.requiredItem = userColon[1].trim(); matched = true; }
  }

  // 4c) "als X User" / "als X zu qualifizieren"
  if (!result.requiredItem) {
    const alsM = s.match(/als\s+([A-Z][A-Za-zÄÖÜäöüß0-9 -]{1,30}?)\s+(?:User|zu\s+qualifizieren)/i);
    if (alsM) { result.requiredItem = alsM[1].trim(); matched = true; }
  }

  // 4d) "für X Quote" / "für X zu qualifizieren"
  if (!result.requiredItem) {
    const fuerM = s.match(/f[üu]r\s+([A-Z][A-Za-zÄÖÜäöüß0-9 -]{1,30}?)\s+(?:Quote|zu\s+qualifizieren)/i);
    if (fuerM) {
      let item = fuerM[1].trim().replace(/\s+(?:Quote)\s*$/i, '').trim();
      if (item) { result.requiredItem = item; matched = true; }
    }
  }

  return matched ? result : null;
}

// v12.17: Erkennt Alter-/Geschlecht-Fragen, damit bei diesen KEIN Fragetext
// in die Excel-Notiz wandert (User-Wunsch: Fragetext-Notiz fuer ALLE ausser
// Alter und Geschlecht). Matched bewusst tolerant via fragetext/kurz_label.
function isAlterOrGeschlechtFrage(frage) {
  if (!frage) return false;
  const ft = String(frage.fragetext || '').toLowerCase();
  const kl = String(frage.kurz_label || frage.kurzlabel || '').toLowerCase();
  const both = ft + ' ' + kl;
  if (/\bgeschlecht\b/.test(both)) return true;
  // "alt sind", "wie alt", "Alter:" (auch Standalone-Label)
  if (/\balt sind\b|\bwie alt\b|(^|[^a-zäöü])alter([^a-zäöü]|$)/.test(both)) return true;
  return false;
}

// v12.17: Baut die Header-Notiz konsistent. Default: Fragetext + (optional)
// Bedingung + (optional) extras. Bei Alter/Geschlecht entfaellt der Fragetext-
// Anteil — nur Bedingung/extras werden uebernommen (meist leer = keine Notiz).
function buildHeaderNote(frage, extras) {
  const parts = [];
  const ft = String((frage && frage.fragetext) || '').trim();
  const kl = String((frage && (frage.kurz_label || frage.kurzlabel)) || '').trim();
  if (!isAlterOrGeschlechtFrage(frage)) {
    const text = ft || kl;
    if (text) parts.push(text);
  }
  const bd = String((frage && frage.bedingung) || '').trim();
  if (bd) parts.push(`Bedingung: ${bd}`);
  if (extras) {
    const ex = String(extras).trim();
    if (ex) parts.push(ex);
  }
  return parts.join('\n\n').trim();
}

// v12.18: Wählt den passenden quotenkommentar für eine Gruppe.
// Bevorzugt frage.quotenkommentar_pro_zielgruppe[gruppe.zielgruppe] (PATCH 17),
// fällt zurück auf frage.quotenkommentar. Match ist tolerant gegen Whitespace
// und Case (damit "Active Buyers TikTok Shop" auch matched wenn der Parser
// "active buyers tiktok shop" liefert).
//
// v12.21: NEU - universelles gilt_fuer-Konzept (PATCH 17 vereinfacht).
//   frage.quotenkommentar kann jetzt auch ARRAY sein:
//     [{ text: "...", gilt_fuer: ["GD1","GD2"] }, ...]
//   - Eintraege ohne gilt_fuer (oder gilt_fuer=[]) gelten studienweit
//   - gilt_fuer kann Gruppen-IDs (GD1, IDI, VDI) oder Zielgruppen-Namen enthalten
//   - Match ist case/whitespace-tolerant
//   - Mehrere passende Eintraege werden mit '\n' verbunden
function matchesGruppe(giltFuer, gruppe) {
  // Leer/undefined = studienweit = passt fuer alle Gruppen
  if (!Array.isArray(giltFuer) || giltFuer.length === 0) return true;
  const norm = (s) => String(s || '').toLowerCase().replace(/\s+/g, ' ').trim();
  const gid = norm(gruppe && gruppe.id);
  const gzg = norm(gruppe && gruppe.zielgruppe);
  // v12.22.0: QG-IDs (hierarchisch, z.B. "GD1.A", "IDI.S1") werden als
  // Gruppen-Match interpretiert, wenn die Gruppen-ID dem Prefix entspricht.
  // Beispiel: gilt_fuer=["GD1.A"] passt zu gruppe.id="GD1", weil der QG-Eintrag
  // semantisch zu dieser Gruppe gehört (die Filterung pro QG passiert separat).
  for (const entry of giltFuer) {
    const e = norm(entry);
    if (!e) continue;
    if (e === gid || e === gzg) return true;
    // QG-ID-Match: entry hat Punkt-Notation, gruppen-Teil davor matched gid
    const dotIdx = e.indexOf('.');
    if (dotIdx > 0 && e.substring(0, dotIdx) === gid) return true;
  }
  return false;
}

function pickQuotenkommentarFuerGruppe(frage, gruppe) {
  if (!frage) return '';
  // Fall 1: quotenkommentar ist ein ARRAY (v12.21 universelles gilt_fuer)
  if (Array.isArray(frage.quotenkommentar)) {
    return frage.quotenkommentar
      .filter(entry => entry && matchesGruppe(entry.gilt_fuer || entry.gilt_fuer_gruppen, gruppe))
      .map(entry => String(entry.text || '').trim())
      .filter(Boolean)
      .join('\n');
  }
  // Fall 2: quotenkommentar_pro_zielgruppe (v12.18 backward compat)
  const perZG = frage.quotenkommentar_pro_zielgruppe;
  const zielgruppe = (gruppe && gruppe.zielgruppe) ? String(gruppe.zielgruppe).trim() : '';
  if (perZG && typeof perZG === 'object' && zielgruppe) {
    if (perZG[zielgruppe]) return String(perZG[zielgruppe]).trim();
    const norm = (s) => String(s || '').toLowerCase().replace(/\s+/g, ' ').trim();
    const target = norm(zielgruppe);
    for (const key of Object.keys(perZG)) {
      if (norm(key) === target) return String(perZG[key]).trim();
    }
  }
  // Fall 3: einfacher String (Default, studienweit)
  return String(frage.quotenkommentar || '').trim();
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

// v12.12: Cluster-Farbsystem für Bedingte Fragen
// Jede unique Bedingung bekommt eine eigene Farbe. Fragen mit gleicher
// Bedingung teilen die Farbe — so erkennt der Recruiter visuell welche
// Fragen zusammengehören (z.B. "Q11.2a + Q11.2b sind beide nur für
// hohes C Nutzer relevant").
const CLUSTER_COLORS = [
  'FF1E88E5', // blau
  'FF43A047', // grün
  'FFFB8C00', // orange
  'FF8E24AA', // lila
  'FFE53935', // rot
  'FF00897B', // teal
  'FFFFB300', // gelb-orange
  'FF6D4C41', // braun
];
// v12.13: Aufgehellte Pastell-Varianten passend zu CLUSTER_COLORS (gleiche Index-Position)
// Für Hintergrund von Frage-Headern, damit Schrift lesbar bleibt.
const CLUSTER_COLORS_LIGHT = [
  'FFD3E7FA', // blau pastel
  'FFD8EFD7', // grün pastel
  'FFFEE4CC', // orange pastel
  'FFEBD7F3', // lila pastel
  'FFFAD3D2', // rot pastel
  'FFCCE7E4', // teal pastel
  'FFFFEFC9', // gelb-orange pastel
  'FFDDD0CC', // braun pastel
];

// Normalisiert eine Bedingung für den Cluster-Key (case-insensitive, ohne
// Whitespace/Sonderzeichen-Variationen). So matchen "Nur fragen wenn Q11 = X"
// und "Nur wenn Q11 = X" auf den gleichen Cluster.
function normalizeBedingung(b) {
  if (!b) return '';
  return String(b).toLowerCase()
    .replace(/\s+/g, ' ')
    .replace(/[„""']/g, '')
    .replace(/^nur (fragen )?wenn[,:]?\s*/, '')
    .replace(/^nur (für|fuer)\s*/, '')
    .trim();
}

// Berechnet die Cluster-Map: {frageId: {colorArgb, key}}
// Reihenfolge der Cluster basiert auf erstem Auftreten in fragen[].
// v12.12: Modulscope-Variable für aktuelle Cluster-Map (wird pro Build gesetzt)
let __CLUSTER_MAP__ = null;

function computeClusterMap(fragen) {
  const map = {};
  const keyToIdx = {};
  let colorIdx = 0;
  for (const frage of fragen || []) {
    if (!frage || !frage.bedingung) continue;
    const key = normalizeBedingung(frage.bedingung);
    if (!key) continue;
    if (!(key in keyToIdx)) {
      keyToIdx[key] = colorIdx % CLUSTER_COLORS.length;
      colorIdx++;
    }
    const i = keyToIdx[key];
    map[frage.id] = {
      colorArgb: CLUSTER_COLORS[i],
      lightArgb: CLUSTER_COLORS_LIGHT[i],
      key
    };
  }
  return map;
}

// v12.7: Sheet-Sortierung nach Datum + Uhrzeit (chronologisch)
// Frontend liefert datum als "Mo. 19.05.", "Di. 20.05." (Wochentag + DD.MM.).
// Uhrzeit als "14:30 - 16:30". Wir extrahieren Tag, Monat, Stunde, Minute zum
// Sortieren. Wenn Datum/Uhrzeit fehlt, bleibt die Gruppe am Ende.
function getSortKey(gruppe) {
  const datum = String(gruppe.datum || '').trim();
  const uhrzeit = String(gruppe.uhrzeit || '').trim();
  // Datum parsen: irgendwo "DD.MM." oder "DD.MM.YYYY"
  const dm = datum.match(/(\d{1,2})\.(\d{1,2})\.(\d{4})?/);
  if (!dm) return Number.MAX_SAFE_INTEGER;
  const day = parseInt(dm[1]);
  const month = parseInt(dm[2]);
  const year = dm[3] ? parseInt(dm[3]) : 2026; // Fallback aktuelles Jahr
  // Uhrzeit parsen: erste HH:MM Gruppe
  const um = uhrzeit.match(/(\d{1,2}):(\d{2})/);
  const hour = um ? parseInt(um[1]) : 0;
  const minute = um ? parseInt(um[2]) : 0;
  // Numerischer Sort-Key: YYYYMMDDHHMM
  return year * 100000000 + month * 1000000 + day * 10000 + hour * 100 + minute;
}

// v12.9: Quotenübersicht-Sheet als erstes Sheet der Excel
//
// Struktur:
//   1. STUDIEN-WEITE QUOTEN (1x, gelten für alle Gruppen)
//   2. GLOBALE AUSSCHLUSSKRITERIEN (1x, gelten für alle Gruppen)
//   3. ALLGEMEINE HINWEISE (1x, Quotenkommentare auf Frage-Ebene)
//   4. PRO GRUPPE: nur Gruppen-spezifische Anforderungen
//   5. IDI-SEGMENTE falls vorhanden
//
// Ziel: Recruiter ohne tiefes Quotenverständnis kann die Zielgruppe einer
// Gruppe sofort erfassen. Globale Daten nur einmal, Gruppen-Block nur das
// Spezifische.
function buildOverviewSheet(workbook, gruppen, fragen, studienQuoten, idiProfile, methode) {
  const ws = workbook.addWorksheet('Quotenübersicht', {
    properties: { tabColor: { argb: 'FFFFC000' } }
  });

  // v12.10: Outline-Properties für Ein-/Ausklappen
  // summaryBelow:false = Plus/Minus erscheint OBERHALB des gruppierten Bereichs
  // (also direkt am Section-Header, nicht am Ende)
  ws.properties.outlineProperties = {
    summaryBelow: false,
    summaryRight: false
  };

  ws.getColumn(1).width = 4;
  ws.getColumn(2).width = 24;
  ws.getColumn(3).width = 60;
  ws.getColumn(4).width = 35;
  ws.getColumn(5).width = 4;

  let row = 2;

  // =========================================================================
  // HAUPT-TITEL
  // =========================================================================
  ws.mergeCells(`B${row}:D${row}`);
  const titleCell = ws.getCell(`B${row}`);
  titleCell.value = 'QUOTENÜBERSICHT';
  titleCell.font = { name: 'Calibri', size: 18, bold: true, color: { argb: 'FF1F4E78' } };
  titleCell.alignment = { horizontal: 'left', vertical: 'middle' };
  ws.getRow(row).height = 30;
  row += 2;

  // =========================================================================
  // SEKTION 1: STUDIEN-WEITE QUOTEN
  // =========================================================================
  if (Array.isArray(studienQuoten) && studienQuoten.length > 0) {
    ws.mergeCells(`B${row}:D${row}`);
    const c = ws.getCell(`B${row}`);
    c.value = '📌 STUDIEN-WEITE QUOTEN (gelten für ALLE Gruppen)';
    c.font = { name: 'Calibri', size: 12, bold: true, color: { argb: 'FFFFFFFF' } };
    c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1F4E78' } };
    c.alignment = { horizontal: 'left', vertical: 'middle', indent: 1 };
    ws.getRow(row).height = 22;
    row++;

    for (const q of studienQuoten) {
      ws.mergeCells(`B${row}:D${row}`);
      const cc = ws.getCell(`B${row}`);
      cc.value = '• ' + q;
      cc.font = { name: 'Calibri', size: 11 };
      cc.alignment = { horizontal: 'left', vertical: 'middle', wrapText: true, indent: 1 };
      cc.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF2F2F2' } };
      ws.getRow(row).height = Math.max(20, Math.ceil(q.length / 90) * 18);
      ws.getRow(row).outlineLevel = 1;
      row++;
    }
    row += 2;
  }

  // =========================================================================
  // SEKTION 2: GLOBALE AUSSCHLUSSKRITERIEN
  // =========================================================================
  const globalScreenouts = collectGlobalScreenouts(fragen);
  if (globalScreenouts.length > 0) {
    ws.mergeCells(`B${row}:D${row}`);
    const c = ws.getCell(`B${row}`);
    c.value = '❌ AUSSCHLUSSKRITERIEN (gelten für ALLE Gruppen)';
    c.font = { name: 'Calibri', size: 12, bold: true, color: { argb: 'FFFFFFFF' } };
    c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFC00000' } };
    c.alignment = { horizontal: 'left', vertical: 'middle', indent: 1 };
    ws.getRow(row).height = 22;
    row++;

    // Spalten-Header
    ws.getCell(`B${row}`).value = 'Frage';
    ws.getCell(`C${row}`).value = 'Antwort/Item';
    ws.getCell(`D${row}`).value = 'Konsequenz';
    for (const col of ['B','C','D']) {
      const h = ws.getCell(`${col}${row}`);
      h.font = { name: 'Calibri', size: 10, bold: true, color: { argb: 'FF595959' } };
      h.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF2F2F2' } };
      h.alignment = { horizontal: 'left', vertical: 'middle', indent: 1 };
    }
    ws.getRow(row).height = 20;
    ws.getRow(row).outlineLevel = 1;
    row++;

    for (const eintrag of globalScreenouts) {
      ws.getCell(`B${row}`).value = eintrag.frageLabel;
      ws.getCell(`B${row}`).font = { name: 'Calibri', size: 10, bold: true };
      ws.getCell(`B${row}`).alignment = { horizontal: 'left', vertical: 'top', wrapText: true, indent: 1 };

      ws.getCell(`C${row}`).value = eintrag.codeText;
      ws.getCell(`C${row}`).font = { name: 'Calibri', size: 10 };
      ws.getCell(`C${row}`).alignment = { horizontal: 'left', vertical: 'top', wrapText: true, indent: 1 };

      ws.getCell(`D${row}`).value = '→ Screenout';
      ws.getCell(`D${row}`).font = { name: 'Calibri', size: 10, bold: true, color: { argb: 'FFC00000' } };
      ws.getCell(`D${row}`).alignment = { horizontal: 'left', vertical: 'top', indent: 1 };

      const maxLen = (eintrag.codeText || '').length;
      ws.getRow(row).height = Math.max(18, Math.min(60, Math.ceil(maxLen / 50) * 18));
      ws.getRow(row).outlineLevel = 1;
      row++;
    }
    row += 2;
  }

  // =========================================================================
  // SEKTION 3: ALLGEMEINE HINWEISE (Quotenkommentare auf Frage-Ebene)
  // =========================================================================
  const globalHinweise = collectGlobalHinweise(fragen);
  if (globalHinweise.length > 0) {
    ws.mergeCells(`B${row}:D${row}`);
    const c = ws.getCell(`B${row}`);
    c.value = '💡 HINWEISE FÜR DEN REKRUTIERER (gelten für ALLE Gruppen)';
    c.font = { name: 'Calibri', size: 12, bold: true, color: { argb: 'FFFFFFFF' } };
    c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFBF8F00' } };
    c.alignment = { horizontal: 'left', vertical: 'middle', indent: 1 };
    ws.getRow(row).height = 22;
    row++;

    // v12.22.21 (Bug 1): Gruppieren nach frageLabel — Frage einmal als Header,
    // Codes/Items eingerückt darunter. Vorher: 'QGENDER. Geschlecht' 5x in B-Spalte.
    let lastLabel = null;
    for (const h of globalHinweise) {
      // v12.22.21 (Bug 3): Q-Code-Präfix aus frageLabel entfernen
      const cleanLabel = stripQCode(h.frageLabel) || h.frageLabel;
      const isNewQuestion = cleanLabel !== lastLabel;
      if (isNewQuestion) {
        // Header-Zeile mit Frage-Label (bold, B-Spalte)
        ws.mergeCells(`B${row}:D${row}`);
        const hb = ws.getCell(`B${row}`);
        hb.value = cleanLabel;
        hb.font = { name: 'Calibri', size: 10, bold: true, color: { argb: 'FF1F4E78' } };
        hb.alignment = { horizontal: 'left', vertical: 'middle', indent: 1 };
        hb.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFF9E6' } };
        ws.getRow(row).height = 18;
        ws.getRow(row).outlineLevel = 1;
        row++;
        lastLabel = cleanLabel;
      }
      // Code/Text-Zeile (eingerückt, C+D gemerged)
      ws.getCell(`B${row}`).value = '';
      ws.getCell(`B${row}`).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFF9E6' } };
      ws.mergeCells(`C${row}:D${row}`);
      const t = ws.getCell(`C${row}`);
      t.value = '  ' + h.text;
      t.font = { name: 'Calibri', size: 10, italic: true };
      t.alignment = { horizontal: 'left', vertical: 'top', wrapText: true, indent: 2 };
      t.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFF9E6' } };

      ws.getRow(row).height = Math.max(18, Math.min(80, Math.ceil(h.text.length / 70) * 18));
      ws.getRow(row).outlineLevel = 1;
      row++;
    }
    row += 2;
  }

  // =========================================================================
  // SEKTION 3.5: CLUSTER-LEGENDE (Farbcodierung bedingter Fragen)
  // =========================================================================
  if (__CLUSTER_MAP__ && Object.keys(__CLUSTER_MAP__).length > 0) {
    ws.mergeCells(`B${row}:D${row}`);
    const ch = ws.getCell(`B${row}`);
    ch.value = '🎨 FARB-CLUSTER: Zusammenhängende bedingte Fragen';
    ch.font = { name: 'Calibri', size: 12, bold: true, color: { argb: 'FFFFFFFF' } };
    ch.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF595959' } };
    ch.alignment = { horizontal: 'left', vertical: 'middle', indent: 1 };
    ws.getRow(row).height = 22;
    row++;

    // Pro Cluster eine Zeile: Farbe + Bedingung + welche Fragen
    const clusterToFragen = {};
    for (const [fid, info] of Object.entries(__CLUSTER_MAP__)) {
      const key = info.key;
      if (!clusterToFragen[key]) {
        clusterToFragen[key] = { color: info.colorArgb, key, fragen: [] };
      }
      // Frage-Label nachschlagen
      const frage = fragen.find(f => f.id === fid);
      if (frage) {
        clusterToFragen[key].fragen.push({
          id: fid,
          label: frage.kurz_label || fid,
          bedingung: frage.bedingung || ''
        });
      }
    }
    for (const cluster of Object.values(clusterToFragen)) {
      const fragenList = cluster.fragen.map(f => f.label).join(', ');
      const bedingung = cluster.fragen[0] ? cluster.fragen[0].bedingung : '';

      // Spalte B: farbiger Indikator
      const bc = ws.getCell(`B${row}`);
      bc.value = '●';
      bc.font = { name: 'Calibri', size: 18, bold: true, color: { argb: cluster.color } };
      bc.alignment = { horizontal: 'center', vertical: 'middle' };

      // Spalte C: Fragen-Liste
      const cc = ws.getCell(`C${row}`);
      cc.value = fragenList;
      cc.font = { name: 'Calibri', size: 10, bold: true, color: { argb: cluster.color } };
      cc.alignment = { horizontal: 'left', vertical: 'middle', wrapText: true, indent: 1 };

      // Spalte D: Bedingung
      const dc = ws.getCell(`D${row}`);
      dc.value = '⚠ ' + bedingung;
      dc.font = { name: 'Calibri', size: 10, italic: true };
      dc.alignment = { horizontal: 'left', vertical: 'middle', wrapText: true, indent: 1 };

      ws.getRow(row).height = Math.max(22, Math.ceil(bedingung.length / 60) * 18);
      ws.getRow(row).outlineLevel = 1;
      row++;
    }
    row += 2;
  }

  // =========================================================================
  // SEKTION 4: PRO GRUPPE EIN BLOCK (nur Gruppen-spezifische Anforderungen)
  // =========================================================================
  const COLORS = [
    { header: 'FF2E75B6', body: 'FFDEEBF7' }, // Blau
    { header: 'FF548235', body: 'FFE2EFDA' }, // Grün
    { header: 'FFBF8F00', body: 'FFFFF2CC' }, // Gelb-Gold
    { header: 'FFC65911', body: 'FFFCE4D6' }, // Orange
    { header: 'FF7030A0', body: 'FFE4D6F0' }, // Violett
    { header: 'FFA52A2A', body: 'FFF5DEDE' }, // Braun-Rot
  ];

  for (let i = 0; i < gruppen.length; i++) {
    const gruppe = gruppen[i];
    const farbe = COLORS[i % COLORS.length];

    // Header
    ws.mergeCells(`B${row}:D${row}`);
    const h = ws.getCell(`B${row}`);
    const groupLabel = gruppe.id + (gruppe.zielgruppe ? ' — ' + gruppe.zielgruppe : '');
    h.value = groupLabel;
    h.font = { name: 'Calibri', size: 14, bold: true, color: { argb: 'FFFFFFFF' } };
    h.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: farbe.header } };
    h.alignment = { horizontal: 'left', vertical: 'middle', indent: 1 };
    ws.getRow(row).height = 26;
    row++;

    // Metadaten
    const metaParts = [];
    if (gruppe.standort) metaParts.push('📍 ' + gruppe.standort);
    if (gruppe.datum || gruppe.uhrzeit) {
      metaParts.push('🕐 ' + ((gruppe.datum || '') + ' ' + (gruppe.uhrzeit || '')).trim());
    }
    const brutto = gruppe.brutto || gruppe.tnBrutto || 8;
    const netto = gruppe.netto || 6;
    metaParts.push('👥 ' + brutto + ' Brutto / ' + netto + ' Netto');
    if (gruppe.incentive) metaParts.push('💰 ' + gruppe.incentive);

    ws.mergeCells(`B${row}:D${row}`);
    const mc = ws.getCell(`B${row}`);
    mc.value = metaParts.join('   ·   ');
    mc.font = { name: 'Calibri', size: 10, italic: true, color: { argb: 'FF595959' } };
    mc.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: farbe.body } };
    mc.alignment = { horizontal: 'left', vertical: 'middle', indent: 1 };
    ws.getRow(row).height = 20;
    ws.getRow(row).outlineLevel = 1;
    row++;

    // Zuordnungs-Kriterien (Klartext)
    if (gruppe.zuordnungs_kriterien) {
      ws.mergeCells(`B${row}:D${row}`);
      const zc = ws.getCell(`B${row}`);
      zc.value = '📋 Zuordnung: ' + gruppe.zuordnungs_kriterien;
      zc.font = { name: 'Calibri', size: 10, color: { argb: 'FF1F4E78' } };
      zc.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: farbe.body } };
      zc.alignment = { horizontal: 'left', vertical: 'middle', wrapText: true, indent: 1 };
      ws.getRow(row).height = Math.max(30, Math.ceil(gruppe.zuordnungs_kriterien.length / 90) * 18);
      ws.getRow(row).outlineLevel = 1;
      row++;
    }
    ws.getRow(row).outlineLevel = 1;
    row++;

    // Anforderungen
    const anforderungen = collectAnforderungen(fragen, gruppe.id, __CLUSTER_MAP__);
    if (anforderungen.length > 0) {
      ws.mergeCells(`B${row}:D${row}`);
      const ah = ws.getCell(`B${row}`);
      ah.value = '✅ ANFORDERUNGEN AN TEILNEHMER (spezifisch für ' + gruppe.id + ')';
      ah.font = { name: 'Calibri', size: 11, bold: true, color: { argb: 'FF548235' } };
      ah.alignment = { horizontal: 'left', vertical: 'middle', indent: 1 };
      ws.getRow(row).height = 22;
      ws.getRow(row).outlineLevel = 1;
      row++;

      // Header-Zeile
      ws.getCell(`B${row}`).value = 'Frage';
      ws.getCell(`C${row}`).value = 'Antwort/Item';
      ws.getCell(`D${row}`).value = 'Quote/Anforderung';
      for (const col of ['B','C','D']) {
        const hc = ws.getCell(`${col}${row}`);
        hc.font = { name: 'Calibri', size: 10, bold: true, color: { argb: 'FF595959' } };
        hc.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF2F2F2' } };
        hc.alignment = { horizontal: 'left', vertical: 'middle', indent: 1 };
      }
      ws.getRow(row).height = 20;
      ws.getRow(row).outlineLevel = 1;
      row++;

      // v12.22.21 (Bug 1): Gruppieren nach frageLabel — Frage einmal als
      // Header, Codes/Items eingerückt darunter.
      let lastAnforderungLabel = null;
      for (const eintrag of anforderungen) {
        // v12.22.21 (Bug 3): Q-Code-Präfix aus frageLabel entfernen
        const cleanLabel = stripQCode(eintrag.frageLabel) || eintrag.frageLabel;
        const labelPrefix = eintrag.clusterColor ? '● ' : '';
        const isNewQuestion = cleanLabel !== lastAnforderungLabel;
        const fillSpec = eintrag.clusterLight
          ? { type: 'pattern', pattern: 'solid', fgColor: { argb: eintrag.clusterLight } }
          : null;

        if (isNewQuestion) {
          // Header-Zeile mit Frage-Label
          ws.mergeCells(`B${row}:D${row}`);
          const hb = ws.getCell(`B${row}`);
          hb.value = labelPrefix + cleanLabel;
          hb.font = { name: 'Calibri', size: 10, bold: true,
                      color: { argb: eintrag.clusterColor || 'FF1F4E78' } };
          hb.alignment = { horizontal: 'left', vertical: 'middle', indent: 1 };
          if (fillSpec) hb.fill = fillSpec;
          ws.getRow(row).height = 18;
          ws.getRow(row).outlineLevel = 1;
          row++;
          lastAnforderungLabel = cleanLabel;
        }
        // Code/Soll-Zeile (eingerückt)
        ws.getCell(`B${row}`).value = '';
        if (fillSpec) ws.getCell(`B${row}`).fill = fillSpec;

        ws.getCell(`C${row}`).value = '  ' + eintrag.codeText;
        ws.getCell(`C${row}`).font = { name: 'Calibri', size: 10 };
        ws.getCell(`C${row}`).alignment = { horizontal: 'left', vertical: 'top', wrapText: true, indent: 2 };
        if (fillSpec) ws.getCell(`C${row}`).fill = fillSpec;

        ws.getCell(`D${row}`).value = eintrag.soll;
        ws.getCell(`D${row}`).font = { name: 'Calibri', size: 10, bold: true, color: { argb: 'FF548235' } };
        ws.getCell(`D${row}`).alignment = { horizontal: 'left', vertical: 'top', wrapText: true, indent: 1 };
        if (fillSpec) ws.getCell(`D${row}`).fill = fillSpec;

        const maxLen = Math.max((eintrag.codeText || '').length, (eintrag.soll || '').length);
        ws.getRow(row).height = Math.max(18, Math.min(60, Math.ceil(maxLen / 50) * 18));
        ws.getRow(row).outlineLevel = 1;
        row++;
      }
    } else {
      // v12.20: Statt nur einem Hinweis-Text — zeige die globalen Hinweise
      // nochmal pro Gruppe als einklappbare Liste (outlineLevel=1). So sieht
      // der Recruiter sie direkt im Gruppen-Block ohne nach oben scrollen zu
      // müssen. Bei Bedarf zuklappen.
      // v12.22.21 (Bug 2): Exakt-Match-Dedupe gegen die globalen HINWEISE
      // FÜR DEN REKRUTIERER weiter oben. Wenn alle Einträge identisch sind
      // (typischer Fall), fällt der Block weg und die '(keine Gruppen-
      // spezifischen)'-Meldung wird angezeigt.
      // v12.22.24 (Bug 9): Exakt-Match war zu strikt — bei zusammengesetzten
      // Quote-Texten wie "GD1: Codes 1-3 ... GD2: Codes 4-5 ... GD3: ..."
      // war der String global UND in jeder VGD-Sektion identisch → komplette
      // Sektion ging verloren ('keine Gruppen-spezifischen Anforderungen').
      // Neue Logik: filterQuoteByGruppe() pro Gruppe anwenden — Texte werden
      // satz-weise gefiltert, sodass GD1-Sätze nur in VGD1 erscheinen, GD2 nur
      // in VGD2 etc. Wenn nach Filter etwas übrig bleibt UND es nicht 1:1
      // identisch zum globalen Text ist → anzeigen.
      const globalHinweiseFallback = collectGlobalHinweise(fragen);
      const gruppenSpezifischeHinweise = [];
      for (const h of globalHinweiseFallback) {
        const filtered = filterQuoteByGruppe(h.text, gruppe.id, allGruppen);
        // Anzeigen wenn:
        //  (a) nach Filter wurde etwas weggekürzt (= war GD-spezifisch),
        //  (b) UND der gefilterte Rest ist nicht leer
        if (filtered && filtered.trim() !== h.text.trim()) {
          gruppenSpezifischeHinweise.push({
            frageLabel: h.frageLabel,
            text: filtered,
          });
        }
      }
      if (gruppenSpezifischeHinweise.length === 0) {
        // Wirklich nichts Gruppen-Spezifisches da — der ursprüngliche Hinweis
        ws.mergeCells(`B${row}:D${row}`);
        const empty = ws.getCell(`B${row}`);
        empty.value = '(keine Gruppen-spezifischen Anforderungen — siehe globale Quoten oben)';
        empty.font = { name: 'Calibri', size: 10, italic: true, color: { argb: 'FF808080' } };
        empty.alignment = { horizontal: 'left', vertical: 'middle', indent: 1 };
        ws.getRow(row).height = 20;
        ws.getRow(row).outlineLevel = 1;
        row++;
      } else {
        ws.mergeCells(`B${row}:D${row}`);
        const ah = ws.getCell(`B${row}`);
        ah.value = `✅ ANFORDERUNGEN SPEZIFISCH FÜR ${(gruppe.id || '').toString().toUpperCase()}`;
        ah.font = { name: 'Calibri', size: 11, bold: true, color: { argb: 'FF548235' } };
        ah.alignment = { horizontal: 'left', vertical: 'middle', indent: 1 };
        ws.getRow(row).height = 22;
        ws.getRow(row).outlineLevel = 1;
        row++;

        // v12.22.21 (Bug 1): Auch hier nach frageLabel gruppieren
        let lastFallbackLabel = null;
        for (const h of gruppenSpezifischeHinweise) {
          const cleanLabel = stripQCode(h.frageLabel) || h.frageLabel;
          const isNewQuestion = cleanLabel !== lastFallbackLabel;
          if (isNewQuestion) {
            ws.mergeCells(`B${row}:D${row}`);
            const hb = ws.getCell(`B${row}`);
            hb.value = cleanLabel;
            hb.font = { name: 'Calibri', size: 9, bold: true, color: { argb: 'FF1F4E78' } };
            hb.alignment = { horizontal: 'left', vertical: 'middle', indent: 1 };
            hb.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFF9E6' } };
            ws.getRow(row).height = 16;
            ws.getRow(row).outlineLevel = 1;
            row++;
            lastFallbackLabel = cleanLabel;
          }
          ws.getCell(`B${row}`).value = '';
          ws.getCell(`B${row}`).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFF9E6' } };
          ws.mergeCells(`C${row}:D${row}`);
          const t = ws.getCell(`C${row}`);
          t.value = '  ' + h.text;
          t.font = { name: 'Calibri', size: 9, italic: true };
          t.alignment = { horizontal: 'left', vertical: 'top', wrapText: true, indent: 2 };
          t.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFF9E6' } };
          ws.getRow(row).height = Math.max(16, Math.min(60, Math.ceil(h.text.length / 80) * 14));
          ws.getRow(row).outlineLevel = 1;
          row++;
        }
      }
    }

    row += 2;
  }

  // =========================================================================
  // SEKTION 5: IDI-SEGMENTE
  // =========================================================================
  if (Array.isArray(idiProfile) && idiProfile.length > 0) {
    ws.mergeCells(`B${row}:D${row}`);
    const c = ws.getCell(`B${row}`);
    c.value = '🎯 IDI-SEGMENT-ZUORDNUNG';
    c.font = { name: 'Calibri', size: 14, bold: true, color: { argb: 'FFFFFFFF' } };
    c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1F4E78' } };
    c.alignment = { horizontal: 'left', vertical: 'middle', indent: 1 };
    ws.getRow(row).height = 26;
    row++;

    ws.getCell(`B${row}`).value = 'IDI';
    ws.getCell(`C${row}`).value = 'Segment';
    ws.getCell(`D${row}`).value = 'Profil-Anforderungen';
    for (const col of ['B','C','D']) {
      const hc = ws.getCell(`${col}${row}`);
      hc.font = { name: 'Calibri', size: 10, bold: true };
      hc.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF2F2F2' } };
      hc.alignment = { horizontal: 'left', vertical: 'middle', indent: 1 };
    }
    row++;

    for (const idi of idiProfile) {
      ws.getCell(`B${row}`).value = idi.id || '';
      ws.getCell(`B${row}`).font = { name: 'Calibri', size: 10, bold: true };
      ws.getCell(`B${row}`).alignment = { horizontal: 'left', vertical: 'top', indent: 1 };

      let segText = idi.segment || '';
      if (idi.segment_nr) segText = 'Segment ' + idi.segment_nr + ' (' + segText + ')';
      if (idi.cluster) segText += '\nCluster: ' + idi.cluster;
      ws.getCell(`C${row}`).value = segText;
      ws.getCell(`C${row}`).font = { name: 'Calibri', size: 10 };
      ws.getCell(`C${row}`).alignment = { horizontal: 'left', vertical: 'top', wrapText: true, indent: 1 };

      const profilParts = [];
      if (idi.alter_min || idi.alter_max) {
        const rangeStr = (idi.alter_max != null) ?
          (idi.alter_min + '-' + idi.alter_max + ' Jahre') :
          (idi.alter_min + '+ Jahre');
        profilParts.push('Alter: ' + rangeStr);
      }
      if (Array.isArray(idi.profile_quoten)) {
        for (const pq of idi.profile_quoten) {
          if (pq.text) profilParts.push(pq.text);
        }
      }
      ws.getCell(`D${row}`).value = profilParts.join('\n');
      ws.getCell(`D${row}`).font = { name: 'Calibri', size: 10 };
      ws.getCell(`D${row}`).alignment = { horizontal: 'left', vertical: 'top', wrapText: true, indent: 1 };

      ws.getRow(row).height = Math.max(30, profilParts.length * 16);
      row++;
    }
  }

  return ws;
}

// ----------------------------------------------------------------------------
// Sammler-Helfer
// ----------------------------------------------------------------------------
function isQuoteForGruppe(quoteArr, gruppeId) {
  if (!Array.isArray(quoteArr) || quoteArr.length === 0) return null;
  for (const sq of quoteArr) {
    if (!sq) continue;
    const fuerGruppen = Array.isArray(sq.gilt_fuer_gruppen) ? sq.gilt_fuer_gruppen : [];
    // Nur match wenn explizit für diese Gruppe — leere Liste = global, NICHT pro-Gruppe
    if (fuerGruppen.includes(gruppeId) || fuerGruppen.includes('alle')) {
      return sq.text || '';
    }
  }
  return null;
}

function isGlobalQuote(quoteArr) {
  // Globale Quote: soll_quote vorhanden aber gilt_fuer_gruppen leer
  if (!Array.isArray(quoteArr) || quoteArr.length === 0) return null;
  for (const sq of quoteArr) {
    if (!sq) continue;
    const fuerGruppen = Array.isArray(sq.gilt_fuer_gruppen) ? sq.gilt_fuer_gruppen : [];
    if (fuerGruppen.length === 0) {
      return sq.text || '';
    }
  }
  return null;
}

function collectAnforderungen(fragen, gruppeId, clusterMap) {
  // Nur Gruppen-SPEZIFISCHE Anforderungen (soll_quote mit gilt_fuer_gruppen-Eintrag)
  // v12.12: clusterColor mitgeben für bedingte Fragen
  const ergebnisse = [];
  for (const frage of fragen || []) {
    if (frage.kategorie === 'entfaellt') continue;
    if (isStudienTeilnahmeFrage(frage)) continue;  // v12.22.9
    const label = frage.kurz_label || frage.id;
    const clusterEntry = (clusterMap && frage.id) ? clusterMap[frage.id] : null;
    const clusterColor = clusterEntry ? clusterEntry.colorArgb : null;
    const clusterLight = clusterEntry ? clusterEntry.lightArgb : null;
    const bedingungText = frage.bedingung || '';

    if (Array.isArray(frage.antworten)) {
      for (const ant of frage.antworten) {
        if (ant.screenout) continue;
        const soll = isQuoteForGruppe(ant.soll_quote, gruppeId);
        if (soll != null) {
          ergebnisse.push({
            frageLabel: label,
            codeText: 'Code ' + ant.code + ': ' + ant.text,
            soll: '✓ ' + soll,
            clusterColor,
            clusterLight,
            bedingungText
          });
        }
      }
    }

    if (Array.isArray(frage.items)) {
      for (const item of frage.items) {
        const soll = isQuoteForGruppe(item.soll_quote, gruppeId);
        if (soll != null) {
          ergebnisse.push({
            frageLabel: label,
            codeText: 'Item: ' + (item.item_label || ''),
            soll: '✓ ' + soll,
            clusterColor,
            clusterLight,
            bedingungText
          });
        }
      }
    }
  }
  return ergebnisse;
}

function collectGlobalScreenouts(fragen) {
  // Alle screenout=true Antworten + item.screenout_codes (gelten für alle Gruppen)
  const ergebnisse = [];
  for (const frage of fragen || []) {
    if (frage.kategorie === 'entfaellt') continue;
    if (isStudienTeilnahmeFrage(frage)) continue;  // v12.22.9
    const label = frage.kurz_label || frage.id;

    if (Array.isArray(frage.antworten)) {
      for (const ant of frage.antworten) {
        if (ant.screenout) {
          ergebnisse.push({
            frageLabel: label,
            codeText: 'Code ' + ant.code + ': ' + ant.text
          });
        }
      }
    }

    if (Array.isArray(frage.items)) {
      for (const item of frage.items) {
        if (Array.isArray(item.screenout_codes) && item.screenout_codes.length > 0) {
          const codeTexte = item.screenout_codes.map(c => {
            const ant = (frage.antworten || []).find(a => String(a.code) === String(c));
            return 'Code ' + c + (ant ? ': ' + ant.text : '');
          });
          ergebnisse.push({
            frageLabel: label,
            codeText: 'Item "' + (item.item_label || '') + '" bei → ' + codeTexte.join(' / ')
          });
        }
      }
    }
  }
  return ergebnisse;
}

function collectGlobalHinweise(fragen) {
  // Quotenkommentare auf Frage-Ebene plus globale soll_quote-Hinweise (ohne Gruppen-Filter)
  const ergebnisse = [];
  for (const frage of fragen || []) {
    if (frage.kategorie === 'entfaellt') continue;
    if (isStudienTeilnahmeFrage(frage)) continue;  // v12.22.9
    const label = frage.kurz_label || frage.id;

    // Quotenkommentar auf Frage-Ebene
    if (frage.quotenkommentar && String(frage.quotenkommentar).trim()) {
      ergebnisse.push({
        frageLabel: label,
        text: String(frage.quotenkommentar).trim()
      });
    }

    // Globale soll_quote (gilt_fuer_gruppen leer) — z.B. F1 Geschlecht 50/50
    if (Array.isArray(frage.antworten)) {
      for (const ant of frage.antworten) {
        if (ant.screenout) continue;
        const global = isGlobalQuote(ant.soll_quote);
        if (global != null) {
          ergebnisse.push({
            frageLabel: label,
            text: 'Code ' + ant.code + ' (' + ant.text + '): ' + global
          });
        }
      }
    }
    if (Array.isArray(frage.items)) {
      for (const item of frage.items) {
        const global = isGlobalQuote(item.soll_quote);
        if (global != null) {
          ergebnisse.push({
            frageLabel: label,
            text: 'Item "' + (item.item_label || '') + '": ' + global
          });
        }
      }
    }
  }
  return ergebnisse;
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
    //     Wir suchen DVs auf Zellen wie "A8", "A11" etc. im Template-Bereich
    //     (ab TN_START) und kopieren sie auf die neuen Zeilen mit gleicher Spalte.
    const dvModel = ws.dataValidations?.model || ws.dataValidations || {};
    const dvEntries = (typeof dvModel === 'object' && !Array.isArray(dvModel))
      ? Object.entries(dvModel)
      : [];
    // Welche Spalten haben im Template eine DV?
    const dvByCol = new Map(); // col-letter → DV-Objekt
    for (const [sqref, dv] of dvEntries) {
      if (!sqref || !dv) continue;
      // sqref kann sein: "A8", "A8:A17", "A8 B8" oder mehrere Bereiche.
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

  // v12.21.2: Phantom-Cleanup ENTFERNT.
  // Frueher wurden T6='TB' und U6='PaySite' als veraltete Header-Reste
  // geloescht. Seit dem Vorlagen-Refresh (14.05.2026) sind diese Labels
  // offizieller Bestandteil der Vorlagen und muessen erhalten bleiben.
  // (Quote-N-Reste in Header-Zeile werden weiterhin oben entfernt.)

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
    if (isStudienTeilnahmeFrage(f)) continue;  // v12.22.9
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
    // v12.22.9: Studien-Teilnahme-Historie raus (PATCH 20 defense in depth)
    if (isStudienTeilnahmeFrage(frage)) continue;
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
      // v12.18: quotenkommentar_pro_zielgruppe vorrangig (PATCH 17)
      const matrixQuoteText = pickQuotenkommentarFuerGruppe(frage, gruppe) || frage.bedingung || '';
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

      // v12.11: Globale Matrix-Quote: bei nicht-genannten Items in 1. Spalte zeigen
      let matrixQuoteShownInColumn = false;
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
          const markerText = codes
            ? `Quotenrelevant: Code ${codes} = Screenout`
            : 'Quotenrelevant: Item muss zutreffen (Marker X)';
          // v12.11: Marker-Text NUR als Zusatz zu globaler Quote, nicht als Ersatz
          itemQuote = markerText;
        }
        if (!itemQuote && item.marker === 'Y') {
          itemQuote = 'Quotenrelevant: Item darf NICHT zutreffen (Marker Y)';
        }
        // v12.11: Bei der ERSTEN Item-Spalte zusätzlich die globale matrixQuoteText
        // anhängen wenn sie noch nirgends angezeigt wurde. So sieht der Recruiter
        // bei Q10/Q12 die Frage-weite Anweisung.
        if (!matrixQuoteShownInColumn && matrixQuoteText && !itemQuote.includes(matrixQuoteText)) {
          itemQuote = itemQuote
            ? `📌 Gesamt-Quote: ${matrixQuoteText}\n\n${itemQuote}`
            : `📌 Gesamt-Quote: ${matrixQuoteText}`;
          matrixQuoteShownInColumn = true;
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
        // v12.22.22 (Bug 8): Bei VS-Forced-Choice das Item-Label fuer Z7 kuerzen.
        // Das volle "1: Aussage A VS Aussage B" wird zu kompaktem "Item 1" (oder
        // mit Nummer-Prefix wenn vorhanden), die kompletten Aussagen landen in
        // Z6 mit A/B-Pills (siehe unten).
        let z7Label = item.item_label || '';
        if (z7Label && / VS /i.test(z7Label)) {
          // Versuche Nummern-Prefix zu extrahieren ("1: ..." → "Item 1")
          const numMatch = z7Label.match(/^(\d+)\s*[:.]/);
          z7Label = numMatch ? `Item ${numMatch[1]}` : 'Item';
        }
        writeQuestionColumn(ws, currentCol, z7Label, itemNote,
                            ants, itemQuote, tnEnd, gruppe, frage, isLast, allGruppen, uniformQuoteRow);

        // v12.22.22 (Bug 8): Forced-Choice A/B — zwei Quellen:
        //  Variante 1: Parser liefert getrennte Felder item.aussage_a + item.aussage_b
        //  Variante 2: Parser bündelt beide im item_label mit ' VS '-Trenner
        //              (z.B. "1: Ich ergreife die Initiative VS Ich lasse mich treiben")
        // Beide Wege werden hier abgefangen. Z6 wird mit A/B-Pills bestueckt,
        // sodass der Recruiter beide Aussagen vorlesen kann.
        let aussageA = null, aussageB = null;
        if (item.aussage_a && item.aussage_b) {
          aussageA = String(item.aussage_a).trim();
          aussageB = String(item.aussage_b).trim();
        } else if (item.item_label && / VS /i.test(item.item_label)) {
          // VS-Trenner: optional Nummern-Prefix "1: ..." beibehalten in A
          const parts = item.item_label.split(/\s+VS\s+/i);
          if (parts.length === 2) {
            aussageA = parts[0].trim();
            aussageB = parts[1].trim();
          }
        }
        if (aussageA && aussageB) {
          const longCell = ws.getCell(LONG_QUESTION_ROW, currentCol);
          longCell.value = {
            richText: [
              { text: 'A', font: { name: 'Arial', size: 8, bold: true, color: { argb: 'FF042C53' } } },
              { text: '  ' + aussageA + '\n\n',
                font: { name: 'Arial', size: 9, italic: true, color: { argb: 'FF333333' } } },
              { text: 'B', font: { name: 'Arial', size: 8, bold: true, color: { argb: 'FF4B1528' } } },
              { text: '  ' + aussageB,
                font: { name: 'Arial', size: 9, italic: true, color: { argb: 'FF333333' } } },
            ],
          };
          longCell.alignment = { horizontal: 'left', vertical: 'top', wrapText: true };
          longCell.border = { ...THIN_BORDER };
          const lenSum = aussageA.length + aussageB.length;
          const estLines = Math.max(4, Math.ceil(lenSum / 35) + 2);
          const targetH = Math.min(160, 14 * estLines);
          const curH = ws.getRow(LONG_QUESTION_ROW).height || 15;
          if (targetH > curH) ws.getRow(LONG_QUESTION_ROW).height = targetH;
        }

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
      // v12.22.21 (Bug 3): Z5-Label ohne Q-Code-Präfix
      const motherLabel = stripQCode(frage.kurz_label || frage.fragetext || frage.id);
      mh.value = motherLabel + compactHint;
      mh.font = { name: 'Arial', size: 10, bold: true };
      mh.alignment = { horizontal: 'center', vertical: 'center', wrapText: true };
      mh.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.MATRIX_HDR } };
      mh.border = borderWithThickRight(HEADER_BORDER);
      // v12.22.21 (Bug 3 / User-Wunsch): Comment am Mutter-Header entfernt
      // (Alt: const mhNote = buildHeaderNote(frage, matrixQuoteText); if (mhNote) mh.note = mhNote;)

      // v12.22.21 (Bug 4): Mutter-FRAGE in Z6 gemerged über die Matrix-Spannweite.
      // Vorher war Z6 bei Matrix-Items komplett leer, weil writeQuestionColumn
      // die Z6-Befüllung bei isMatrixItemColumn=true skippt. Jetzt schreiben wir
      // den Fragetext einmal zentral hier.
      // v12.22.22 (Bug 8): Skip wenn Forced-Choice A/B — dann hat jedes Item
      // bereits seine eigene Z6 mit A/B-Pills (siehe Item-Loop oben).
      const isForcedChoiceAB = Array.isArray(frage.items) && frage.items.some(
        it => it && it.aussage_a && it.aussage_b
      );
      const matrixFragetext = (frage.fragetext || '').trim();
      const matrixKurzClean = stripQCode(frage.kurz_label || frage.kurzlabel || '').trim();
      if (!isForcedChoiceAB && matrixFragetext && matrixFragetext.toLowerCase() !== matrixKurzClean.toLowerCase()) {
        if (matrixEnd > matrixStart) {
          try {
            ws.mergeCells(LONG_QUESTION_ROW, matrixStart, LONG_QUESTION_ROW, matrixEnd);
          } catch (e) {}
        }
        const longCell = ws.getCell(LONG_QUESTION_ROW, matrixStart);
        longCell.value = matrixFragetext;
        longCell.font = { name: 'Arial', size: 9, italic: true, color: { argb: 'FF555555' } };
        longCell.alignment = { horizontal: 'left', vertical: 'top', wrapText: true };
        longCell.border = { ...THIN_BORDER };
      }
    } else if (frage.typ === 'freitext' || frage.typ === 'numerisch') {
      // Freitext / Numerisch: nur Header, keine Antwort-Codes
      // v12.17: Note via buildHeaderNote — Fragetext + Bedingung (ausser Alter/Geschlecht)
      const note = buildHeaderNote(frage);
      const label = buildHeaderLabel(frage);
      writeQuestionColumn(ws, currentCol, label, note, null, null, tnEnd, gruppe, frage, true, allGruppen, uniformQuoteRow);
      currentCol++;
    } else if (frage.typ === 'tool_input') {
      // v11: Algorithmus-Eingabe-Skala (z.B. JPM Q12a) — 1 Sammelspalte
      const label = buildHeaderLabel(frage);
      const itemList = (frage.items || []).map(it => `• ${it.item_label || ''}`).join('\n');
      // v12.18: quotenkommentar_pro_zielgruppe vorrangig (PATCH 17)
      const tiQuotenkommentar = pickQuotenkommentarFuerGruppe(frage, gruppe);
      const note = `MUTTERFRAGE: ${frage.fragetext || ''}\n\n` +
                   `Algorithmus-Eingabe — Antworten in das externe Segmentierungs-Tool eingeben.\n\n` +
                   (itemList ? `Items:\n${itemList}\n\n` : '') +
                   (tiQuotenkommentar || '');
      const quoteText = tiQuotenkommentar
        || 'Eingabe in Segmentierungs-Tool — Ergebnis: Segment';
      const tiQuoteFmt = formatSegmentQuotes(quoteText, opts.idiProfile);
      writeQuestionColumn(ws, currentCol, label + ' (Tool)', note,
                          [], tiQuoteFmt, tnEnd, gruppe, frage, true, allGruppen, uniformQuoteRow);
      ws.getColumn(currentCol).width = 30;
      currentCol++;
    } else {
      // single_choice, multi_choice, ranking
      const label = buildHeaderLabel(frage);
      // v12.17: Note via buildHeaderNote — Fragetext in Excel-Notiz fuer ALLE
      // Fragen ausser Alter/Geschlecht (User-Wunsch). Vorher war hier nur die
      // Bedingung drin, sodass einfache Single-Choice-Fragen keine Notiz hatten.
      const note = buildHeaderNote(frage);
      // v12.18: quotenkommentar_pro_zielgruppe vorrangig (PATCH 17)
      let quoteText = pickQuotenkommentarFuerGruppe(frage, gruppe) || frage.bedingung || '';
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
  // v12.22.18: Z6 (LONG_QUESTION_ROW) bekommt 72pt fuer mehrzeilige Original-Frage.
  // Template hat 27pt vorgesehen, das reicht nur fuer eine Zeile — wir vergroessern.
  ws.getRow(LONG_QUESTION_ROW).height = 72;
  ws.getRow(HEADER_ROW).height = 75;

  // v12.22.15: writeIdiInfoBlock entfernt - QG-Box ersetzt Segmente/IDI-Profile/Studien-Quoten
  // in Spalte G. Aufruf erfolgt nun unten im QG-Block via writeQgBoxUnderTn.

  // Freeze Panes komplett deaktiviert — User scrollt frei in alle Richtungen
  ws.views = [{ state: 'normal' }];

  // v12.22.2: Quotengruppen-Spalte und Übersichts-Box (wenn Vorlage Spalte F hat)
  // Hier am ECHTEN Ende von fillSheet, NICHT in buildOverviewSheet (Bug v12.22.0/1).
  try {
    const qgCol = findQuotengruppenCol(ws);
    const dbg = {
      sheet: ws.name,
      qgCol: qgCol,
      qgColLetter: qgCol ? columnNumberToLetter(qgCol) : null,
      gruppeId: gruppe && gruppe.id,
      hasIdiProfile: Array.isArray(opts.idiProfile) && opts.idiProfile.length > 0,
      hasQuotengruppen: Array.isArray(gruppe.quotengruppen) && gruppe.quotengruppen.length > 0,
    };
    if (qgCol) {
      // v12.22.10: fragenArr + studienQuoten an Auto-Extract durchreichen
      const qgList = buildQuotengruppenForSheet(gruppe, {
        ...opts,
        fragenArr: fragen,
        studienQuoten: opts.studien_quoten,
      });
      const qgColLetter = columnNumberToLetter(qgCol);
      dbg.qgListLength = qgList.length;
      dbg.qgLabels = qgList.map(q => q.label);
      if (qgList.length > 1) {
        // Mehrere QGs → Dropdown + Übersichts-Box
        const qgLabels = qgList.map(q => q.label);
        applyQuotengruppenDropdown(ws, qgCol, qgLabels, TN_START, tnEnd);
        // v12.22.15: QG-Box jetzt unten in Spalte G (statt oben in Z1-Z4 oder
        // ganz rechts in BG). Recruiter findet sie unter dem TN-Bereich wo
        // frueher die Segment/Studien-Quoten-Liste stand.
        writeQgBoxUnderTn(ws, qgList, qgColLetter, TN_START, tnEnd);
        dbg.action = 'rendered_dropdown_and_box_under_tn';
      } else {
        // v12.22.8: Bei nur 1 QG die Spalte NICHT mehr ausblenden — sonst kollabiert
        // die Spalte unter dem Logo-Bereich und das Logo wird verzerrt dargestellt.
        // Stattdessen: Zellen grau einfärben + sperren (read-only) als visuelles
        // Signal, dass diese Spalte hier nicht genutzt wird. Der User kann sie
        // bei Bedarf manuell löschen.
        grayOutQuotengruppenSpalte(ws, qgCol, HEADER_ROW, tnEnd);
        dbg.action = 'grayed_out_single_qg';
      }
    } else {
      dbg.action = 'no_qg_column_found';
    }
    if (!globalThis.__QG_DEBUG__) globalThis.__QG_DEBUG__ = [];
    globalThis.__QG_DEBUG__.push(dbg);
    console.log(`[QG] ${ws.name}: ${JSON.stringify(dbg)}`);
  } catch (e) {
    console.warn(`[QG] ${ws.name} ERROR:`, e.message, e.stack);
    if (!globalThis.__QG_DEBUG__) globalThis.__QG_DEBUG__ = [];
    globalThis.__QG_DEBUG__.push({ sheet: ws.name, error: e.message });
  }
}

// ---------------------------------------------------------------------------
// 5) HANDLER
// ---------------------------------------------------------------------------

export default async function handler(req, res) {
  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'Method not allowed' });
  }
  // v12.22.1: QG-Debug pro Request frisch initialisieren
  globalThis.__QG_DEBUG__ = [];
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

    const templateBuffer = Buffer.from(templateBase64, 'base64');
    const template = new ExcelJS.Workbook();
    await template.xlsx.load(templateBuffer);

    // v12.22.22 (Bug 5): Defensive Reinigung gespeicherter Excel-Errors
    // (#VALUE!, #REF!, #NAME?, #DIV/0!, #N/A, #NULL!, #NUM!) in den Header-
    // Zeilen Z1-Z4 ALLER Template-Sheets. Die aktuell deployten Templates
    // haben in F1/E1 einen geerbten #VALUE!-Error (Reste einer nicht
    // aufgeloesten Formel beim Template-Speichern). Bis die Templates
    // re-generiert sind, faengt der Builder das hier ab.
    for (const tplSheet of template.worksheets) {
      for (let r = 1; r <= 4; r++) {
        for (let c = 1; c <= 20; c++) {
          const cell = tplSheet.getCell(r, c);
          if (cell && cell.type === ExcelJS.ValueType.Error) {
            cell.value = null;
          } else if (typeof cell.value === 'string' && /^#(VALUE|REF|NAME|DIV\/0|N\/A|NULL|NUM)[!?]$/i.test(cell.value)) {
            cell.value = null;
          }
        }
      }
    }

    // v12.16: Template-Sheet wird PRO GRUPPE ausgewaehlt (Mixed GD+IDI Projekte).
    // Frueher wurde global ein srcWs gewaehlt (basierend auf req.body.methode),
    // sodass IDI-Gruppen in gemischten Projekten faelschlich das GD-Sheet bekamen.
    // Jetzt: Cache pro Sheet-Name + Lookup ueber gruppe.methode.
    const srcCache = new Map();
    async function getSrcAndMerges(methodeForGruppe) {
      const cfgFG = SHEET_CONFIG[methodeForGruppe] || SHEET_CONFIG['GD'];
      const sheetName = cfgFG.sheetName;
      if (srcCache.has(sheetName)) return srcCache.get(sheetName);
      const ws = template.getWorksheet(sheetName);
      if (!ws) {
        const err = new Error(`Sheet '${sheetName}' nicht im Template (verfuegbar: ${template.worksheets.map(w => w.name).join(', ')})`);
        err.statusCode = 400;
        throw err;
      }
      const merges = await readMergesFromBuffer(templateBuffer, sheetName);
      const entry = { srcWs: ws, preMerges: merges, cfgFG, sheetName };
      srcCache.set(sheetName, entry);
      return entry;
    }

    // Pre-Validate: das Default-Sheet (gemaess globalem methode) muss existieren,
    // damit ein offensichtlich falsches Template frueh erkannt wird.
    let defaultEntry;
    try {
      defaultEntry = await getSrcAndMerges(methode);
    } catch (e) {
      return res.status(e.statusCode || 400).json({
        error: e.message,
        available: template.worksheets.map(w => w.name),
      });
    }

    const result = new ExcelJS.Workbook();
    result.creator = 'Preview Generator';
    result.created = new Date();
    const newSheetNames = [];

    // v12.12: Cluster-Map berechnen (vor Overview + Gruppen-Sheets)
    __CLUSTER_MAP__ = computeClusterMap(fragenArr);

    // v12.8: Quotenübersicht als ERSTES Sheet
    // Wird VOR den Gruppen-Sheets erstellt, damit es ganz vorne erscheint.
    // Sortierung der Gruppen passiert weiter unten — wir geben die unsortierte
    // Liste rein, damit die Blöcke später in der gleichen Reihenfolge stehen wie
    // die Sheets selbst. Achtung: gruppen wird unten in-place sortiert, also
    // erst SORTIEREN, dann Overview bauen, dann durch sortierte Liste iterieren.
    if (!isIDI) {
      gruppen.sort((a, b) => getSortKey(a) - getSortKey(b));
    }
    try {
      buildOverviewSheet(result, gruppen, fragenArr, studien_quoten, idiProfile, methode);
    } catch (e) {
      console.warn('Overview sheet failed (non-fatal):', e.message);
    }

    const processGruppe = async (gruppe, sheetName, fragenFG, includeIdiInfo) => {
      const methodeFG = gruppe.methode || methode;
      // v12.16: Template-Sheet PRO GRUPPE aus dem Cache holen (GD vs IDIs vs VGDs vs VDIs).
      const { srcWs: srcWsFG, preMerges: preMergesFG, cfgFG } = await getSrcAndMerges(methodeFG);

      const dstWs = result.addWorksheet(sheetName);
      copyWorksheet(srcWsFG, dstWs, preMergesFG);

      const quoteStartCol = findQuoteStartCol(dstWs, cfgFG.quoteStartCol);
      const brutto = gruppe.brutto || 8;
      const tnEnd = TN_START + brutto - 1;

      addConditionalFormats(dstWs, tnEnd);

      // v12.6: Unternehmen anhand Standort ueberschreiben, BEVOR Logo gesetzt wird
      gruppe.unternehmen = resolveUnternehmenFromStandort(gruppe.standort, gruppe.unternehmen);
      addLogo(dstWs, result, gruppe.unternehmen, setting);

      // v11: nur dem IDI/VDI-Sheet die IDI-Info-Blöcke geben
      // v11.4: termin_blocks ebenfalls nur bei IDI (bei GD pro Sheet 1 Termin im Header)
      // v12.16: preMerges PRO GRUPPE in sheetOptions, damit fillSheet die korrekten
      // Template-Merges fuer die TN-Bereich-Replikation hat.
      // v12.22.10: studien_quoten wird auch bei GD durchgereicht, damit
      // extractQuotengruppenFromFreitext darauf zugreifen kann. Sonst nur
      // bei IDIs noetig (Profile/Termin-Blocks).
      const baseOpts = { ...builderOptions, preMerges: preMergesFG };
      const sheetOptions = includeIdiInfo
        ? baseOpts
        : { ...baseOpts, idiProfile: null, segment_beschreibungen: null,
            termin_blocks: null };

      fillSheet(
        dstWs, gruppe, fragenFG, projektnummer, projektname,
        auftraggeber, setting, methodeFG, quoteStartCol, gruppen, sheetOptions
      );
      newSheetNames.push({ sheet: sheetName, methode: methodeFG, quoteStartCol, brutto, tnEnd });
    };

    if (isIDI) {
      const hauptGruppe = { ...gruppen[0], methode };
      await processGruppe(hauptGruppe, methode === 'VDI' ? 'VDIs' : 'IDIs', fragenArr, true);
    } else {
      // v12.7: Sortierung passiert schon oben vor Overview
      for (const gruppe of gruppen) {
        gruppe.methode = gruppe.methode || methode;
        const sheetName = gruppe.id.replace(/[:\\/\?\*\[\]]/g, '').substring(0, 31);
        // v12.14: Frage-Filter pro Gruppe — relevantFuerGruppen als harter Filter (1C).
        // Screenout-Fragen (Branchenausschluss, Alter unter 18 etc.) IMMER zeigen (2A),
        // auch wenn relevantFuerGruppen sie ausschließt. Erkannt durch:
        //   a) frage.antworten enthält mindestens eine Antwort mit screenout:true
        //   b) frage.items hat mindestens ein Item mit nicht-leeren screenout_codes
        // v12.14: Eine "klassische Screenout-Frage" (Branchenausschluss, Alter, etc.)
        // hat MEHRERE Screenout-Antworten. Bedingte Fragen mit 1 Ja/1 Nein (Screenout)
        // sind KEINE generellen Filter und folgen relevantFuerGruppen normal.
        const isScreenoutFrage = (f) => {
          // Antwort-Ebene: mind. 2 Screenouts ODER ≥50% der Antworten sind Screenout
          if (Array.isArray(f.antworten) && f.antworten.length > 0) {
            const screenoutCount = f.antworten.filter(a => a && a.screenout).length;
            if (screenoutCount >= 2) return true;
            // Edge-Case: 1 Screenout bei 2 Antworten (50/50) zählt NICHT als Filter
            // — das wäre z.B. die typische Ja/Nein-Bedingungsfrage
          }
          // Item-Ebene: mind. 1 Item mit screenout_codes ist ein echter Filter
          if (Array.isArray(f.items) && f.items.some(it =>
            Array.isArray(it.screenout_codes) && it.screenout_codes.length > 0
          )) return true;
          return false;
        };
        const fragenFG = fragenArr.filter(f => {
          // Screenout-Fragen IMMER drin
          if (isScreenoutFrage(f)) return true;
          // Ohne relevantFuerGruppen oder mit 'alle' -> immer drin
          if (!f.relevantFuerGruppen) return true;
          if (!Array.isArray(f.relevantFuerGruppen) || f.relevantFuerGruppen.length === 0) return true;
          if (f.relevantFuerGruppen.includes('alle')) return true;
          // Sonst nur wenn explizit für diese Gruppe
          return f.relevantFuerGruppen.includes(gruppe.id);
        });
        // v12.19: includeIdiInfo basiert auf gruppe.methode, damit eine
        // konsolidierte IDI-Gruppe im Mixed-Pfad (GD-Sheets + 1 IDI-Sheet)
        // ihre termin_blocks/idiProfile/segment_beschreibungen bekommt,
        // während die GD-Sheets sie weiterhin NICHT bekommen.
        const includeIdiInfoForThis = gruppe.methode === 'IDI' || gruppe.methode === 'VDI';
        await processGruppe(gruppe, sheetName, fragenFG, includeIdiInfoForThis);
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
        // v12.16: preMerges PRO SHEET-TYP cachen (Mixed-Projekte). Default-Sheet
        // war frueher die einzige Quelle; jetzt zeigen wir alle gecachten Sheets.
        preMergesPerSheet: Array.from(srcCache.entries()).map(([name, entry]) => ({
          sheet: name, count: entry.preMerges?.length || 0,
        })),
        preMergesDefault: defaultEntry?.preMerges || [],
        // v11.4 NEU: Termin-Blöcke + Laufzeit
        terminBlocksCount: builderOptions.termin_blocks?.length || 0,
        laufzeitVon: builderOptions.laufzeitVon,
        laufzeitBis: builderOptions.laufzeitBis,
        // v12.22.1: Quotengruppen-Debug pro Sheet
        qgDebug: globalThis.__QG_DEBUG__ || [],
        version: 'v12.22.20-cleanup',
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
      qgDebug: globalThis.__QG_DEBUG__ || [],
      version: 'v12.22.20-cleanup',
    });
  }
}
