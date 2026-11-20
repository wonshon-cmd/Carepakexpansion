/**
 * Loan Paste & Go — builds the whole workbook inside Google Sheets.
 *
 * How to use:
 * 1. Open a BLANK Google Sheet (sheets.google.com → Blank).
 * 2. Extensions → Apps Script. Delete what is there, paste this whole file, press Ctrl/⌘+S.
 * 3. Pick "buildLoanPasteAndGo" in the function menu and click ▶ Run (not Deploy).
 *    First time: Review permissions → your account → Advanced → Go to project → Allow.
 * 4. Go back to the sheet: Summary + tabs 1–49 appear (about 1 minute).
 *    Running it again rebuilds everything from scratch (clears all typed data).
 *
 * Paste zones are tall columns. Copying from the bank panel often brings blank lines
 * between the amounts; the sheet skips blanks and labels and reads the numbers in order.
 */
const N = 49;
const NAMES = ["Rishi", "Senthil", "Tobe", "Gabriela", "Joe"];
const GREEN = ["#C6EFCE", "#006100"], YELLOW = ["#FFEB9C", "#9C5700"], RED = ["#FFC7CE", "#9C0006"];
const PASTE_BG = "#EAF1FF", CALC_BG = "#F3F5F8", HEAD_BG = "#18202C", BLUE = "#1F5FD6", MUTED = "#5D6878";
const Z1 = 6, Z2 = 19;                 // paste zone rows 6..19
const RES = 21, STATUS = 24, STEP3 = 27;

function clean_(ref) {
  const t = 'SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(TRIM(' + ref + '),"$",""),",","")," ",""),CHAR(160),"")';
  return '=IF(' + ref + '="","",IFERROR(VALUE(' + t + '),""))';
}
// Everything pasted in a zone (two columns, all rows) joined into one text, with "282,240" → "282240" and "$" removed.
// Works whether the paste lands in one cell, in separate rows, or label | number side by side.
function joined_(c1, c2) {
  return '=REGEXREPLACE(REGEXREPLACE(REGEXREPLACE(TEXTJOIN(" ",TRUE,' + c1 + Z1 + ':' + c2 + Z2 + '),"(?i)application\\s*\\d+|round to nearest percentage[^%]*%|\\d+\\s*%"," "),"(\\d),(\\d)","$1$2"),"\\$","")';
}
// Value for one field: the number right after its label if the label was pasted too,
// otherwise the k-th number in reading order (labels contain no digits, so they are skipped).
function pick_(cell, k, labelRe) {
  const nth = 'INDEX(SPLIT(TRIM(REGEXREPLACE(' + cell + ',"[^0-9.]+"," "))," "),1,' + k + ')';
  return '=IF(' + cell + '="","",IFERROR(VALUE(REGEXEXTRACT(' + cell + ',"(?i)(?:' + labelRe + ')[^0-9]*([0-9.]+)")),IFERROR(VALUE(' + nth + '),"")))';
}
function head_(sh, a1, text) {
  const r = sh.getRange(a1).merge().setBackground(HEAD_BG).setFontColor("#FFFFFF").setFontWeight("bold").setFontSize(12);
  r.getCell(1, 1).setValue(text);
}
function cf_(sh, a1, formula, colors) {
  return SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied(formula)
    .setBackground(colors[0]).setFontColor(colors[1]).setRanges([sh.getRange(a1)]).build();
}
function pasteZone_(sh, col, label) {
  const col2 = String.fromCharCode(col.charCodeAt(0) + 1);   // the column to the right catches "label | number" pastes
  sh.getRange(col2 + Z1 + ":" + col2 + Z2).setBackground(PASTE_BG).setFontColor(MUTED);
  const z = sh.getRange(col + Z1 + ":" + col + Z2);
  z.setBackground(PASTE_BG).setHorizontalAlignment("center")
   .setBorder(true, true, true, true, null, null, BLUE, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  sh.getRange(col + (Z1 - 1)).setValue(label).setFontColor(BLUE).setFontWeight("bold").setFontSize(9).setHorizontalAlignment("center");
}

function buildTemplate_(sh) {
  // A label | B value | C PASTE ZONE | D gap | E label | F value | G PASTE ZONE | H gap | I state | J,K helpers
  [180, 95, 130, 70, 200, 95, 130, 70, 30, 60, 60].forEach((w, i) => sh.setColumnWidth(i + 1, w));
  sh.getRange("A1").setValue("Application 1").setFontSize(20).setFontWeight("bold");
  sh.getRange("A2:G2").merge().setFontColor(MUTED).setFontSize(10);
  sh.getRange("A2").setValue("Copy the blue $ amounts in the bank panel → click the TOP cell of the blue paste box → Ctrl/⌘+Shift+V. Blank lines are fine.");
  sh.getRange("A3").setValue("Owner (pick your name first)").setFontWeight("bold");
  sh.getRange("B3").setBackground(PASTE_BG).setHorizontalAlignment("center")
    .setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(NAMES, true).setAllowInvalid(true).build());
  sh.getRange("E3").setValue("Clicked in bank panel?").setFontWeight("bold");
  sh.getRange("F3").setBackground(PASTE_BG).setHorizontalAlignment("center")
    .setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(["Yes", "No"], true).build());

  // STEP 1 · LTV
  head_(sh, "A4:C4", "STEP 1 · Loan-to-Value page");
  pasteZone_(sh, "C", "▼ PASTE HERE (top cell)");
  ["Loan Amount (L)", "House Price (H)", "Appraisal Value (V)"].forEach((lab, i) => {
    sh.getRange("A" + (Z1 + i)).setValue(lab);
    sh.getRange("B" + (Z1 + i)).setFormula(pick_("$J$" + Z1, i + 1, ["\\(L\\)|loan amount", "\\(H\\)|house price", "\\(V\\)|appraisal"][i])).setNumberFormat("#,##0").setBackground(CALC_BG).setHorizontalAlignment("right");
  });
  // STEP 2 · Front / Back
  head_(sh, "E4:G4", "STEP 2 · Front / Back Ratio pages (only if YELLOW)");
  pasteZone_(sh, "G", "▼ PASTE HERE (top cell)");
  ["Yearly Income (I)", "Additional Income (A)", "Yearly Mortgage Payment (P)", "Debt Expenses (DE)", "Other Taxes (OT)"].forEach((lab, i) => {
    sh.getRange("E" + (Z1 + i)).setValue(lab);
    sh.getRange("F" + (Z1 + i)).setFormula(pick_("$K$" + Z1, i + 1, ["\\(I\\)|yearly income", "\\(A\\)|additional income", "\\(P\\)|mortgage payment", "\\(DE\\)|debt expense", "\\(OT\\)|other tax"][i])).setNumberFormat("#,##0").setBackground(CALC_BG).setHorizontalAlignment("right");
  });
  sh.getRange("E12").setValue("Front page: paste 3 amounts.\nBack page: paste all 5 (replaces the 3).").setFontColor(BLUE).setFontSize(9).setWrap(true);
  sh.getRange("A10").setValue("Paste anything: labels, blank lines or one cell.\nEach number goes to its label (L, H, V).").setFontColor(BLUE).setFontSize(9).setWrap(true);

  // hidden helpers: cleaned paste zones
  sh.getRange("J" + Z1).setFormula(joined_("C", "D"));
  sh.getRange("K" + Z1).setFormula(joined_("G", "H"));

  // results
  sh.getRange("A" + RES).setValue("LTV → type in panel").setFontWeight("bold");
  sh.getRange("C" + RES).setFormula('=IF(COUNT(B6:B8)<3,"",B6/MIN(B7:B8))').setNumberFormat("0.00%").setFontColor(MUTED);
  sh.getRange("B" + RES).setFormula('=IF(C' + RES + '="","–",ROUND(C' + RES + '*100,0))').setFontSize(26).setFontWeight("bold").setHorizontalAlignment("center");
  sh.getRange("E" + RES).setValue("Front → type in panel").setFontWeight("bold");
  sh.getRange("G" + RES).setFormula('=IF(COUNT(F6:F8)<3,"",F8/(F6+F7))').setNumberFormat("0.00%").setFontColor(MUTED);
  sh.getRange("F" + RES).setFormula('=IF(G' + RES + '="","–",ROUND(G' + RES + '*100,0))');
  sh.getRange("E" + (RES + 1)).setValue("Back → type in panel").setFontWeight("bold");
  sh.getRange("G" + (RES + 1)).setFormula('=IF(COUNT(F6:F10)<5,"",(F9+F10+F8)/(F6+F7))').setNumberFormat("0.00%").setFontColor(MUTED);
  sh.getRange("F" + (RES + 1)).setFormula('=IF(G' + (RES + 1) + '="","–",ROUND(G' + (RES + 1) + '*100,0))');
  sh.getRange("F" + RES + ":F" + (RES + 1)).setFontSize(22).setFontWeight("bold").setHorizontalAlignment("center");
  sh.setRowHeight(RES, 44); sh.setRowHeight(RES + 1, 40);

  // state + status bar
  const L = "C" + RES, F = "G" + RES, B = "G" + (RES + 1), ST = "$I$" + STATUS;
  sh.getRange("I" + STATUS).setFormula('=IF(' + L + '="","",IF(' + L + '>1,"R",IF(' + L + '<0.85,"G",IF(OR(' + F + '="",' + B + '=""),"Y",IF(AND(' + F + '<0.25,' + B + '<0.35),"G","R")))))');
  sh.getRange("A" + STATUS + ":G" + (STATUS + 1)).merge().setFontSize(18).setFontWeight("bold").setHorizontalAlignment("center").setVerticalAlignment("middle");
  sh.getRange("A" + STATUS).setFormula('=IF(I' + STATUS + '="","PASTE L, H, V ABOVE",IF(I' + STATUS + '="G","✓ APPROVED — click Approve Loan",IF(I' + STATUS + '="Y","NEED DATA — paste Front/Back in STEP 2","✗ NOT QUALIFIED — submit Reject Loan, then this one is complete")))');
  sh.setRowHeight(STATUS, 34); sh.setRowHeight(STATUS + 1, 34);

  // STEP 3 · puzzle + decision
  head_(sh, "A" + STEP3 + ":C" + STEP3, "STEP 3 · Risk Assessment + Decision");
  sh.getRange("A" + (STEP3 + 1)).setValue("Words found (0–5)");
  sh.getRange("B" + (STEP3 + 1)).setBackground(PASTE_BG).setHorizontalAlignment("center")
    .setBorder(true, true, true, true, null, null, BLUE, SpreadsheetApp.BorderStyle.SOLID_MEDIUM)
    .setDataValidation(SpreadsheetApp.newDataValidation().requireNumberBetween(0, 5).build());
  sh.getRange("A" + (STEP3 + 2) + ":C" + (STEP3 + 2)).merge().setFontWeight("bold").setFontSize(13);
  const W = "B" + (STEP3 + 1);
  sh.getRange("A" + (STEP3 + 2)).setFormula('=IF(I' + STATUS + '<>"G","",IF(' + W + '="","Do the puzzle, type words found","Click "&' + W + '&"% → Approve Loan → Next Application"))');
  sh.hideColumns(9, 3);   // I, J, K

  sh.setConditionalFormatRules([
    cf_(sh, "A" + STATUS + ":G" + (STATUS + 1), '=' + ST + '="G"', GREEN),
    cf_(sh, "A" + STATUS + ":G" + (STATUS + 1), '=' + ST + '="Y"', YELLOW),
    cf_(sh, "A" + STATUS + ":G" + (STATUS + 1), '=' + ST + '="R"', RED),
    cf_(sh, "B" + RES, '=AND($C$' + RES + '<>"",$C$' + RES + '>1)', RED),
    cf_(sh, "B" + RES, '=AND($C$' + RES + '<>"",$C$' + RES + '>=0.85)', YELLOW),
    cf_(sh, "B" + RES, '=$C$' + RES + '<>""', GREEN),
    cf_(sh, "F" + RES, '=AND($G$' + RES + '<>"",$G$' + RES + '>=0.25)', RED),
    cf_(sh, "F" + RES, '=$G$' + RES + '<>""', GREEN),
    cf_(sh, "F" + (RES + 1), '=AND($G$' + (RES + 1) + '<>"",$G$' + (RES + 1) + '>=0.35)', RED),
    cf_(sh, "F" + (RES + 1), '=$G$' + (RES + 1) + '<>""', GREEN),
    cf_(sh, "E4:G4", '=' + ST + '="Y"', YELLOW),
    cf_(sh, "F3", '=$F$3="Yes"', GREEN)
  ]);
}

function buildSummary_(sh) {
  [60, 100, 70, 70, 70, 320, 60, 80].forEach((w, i) => sh.setColumnWidth(i + 1, w));
  sh.getRange("A1").setValue("Huntsman Loan Floor · Summary").setFontSize(18).setFontWeight("bold");
  sh.getRange("A2").setValue("Each application has its own tab (1–49). Pick your name in Owner on that tab before you start.").setFontColor(MUTED);
  const last = 5 + N;
  const stats = [["Approved", '=COUNTIF(F6:F' + last + ',"✓*")'], ["Not qualified", '=COUNTIF(F6:F' + last + ',"✗*")'],
                 ["Need data", '=COUNTIF(F6:F' + last + ',"NEED*")'], ["Not started", '=COUNTIF(F6:F' + last + ',"PASTE*")'],
                 ["Clicked in panel", '=COUNTIF(H6:H' + last + ',"Yes")']];
  stats.forEach(([lab, f], i) => {
    const col = [1, 3, 5, 6, 8][i];
    sh.getRange(3, col).setValue(lab).setFontColor(MUTED).setFontSize(10);
    sh.getRange(4, col).setFormula(f).setFontSize(16).setFontWeight("bold");
  });
  sh.getRange("A5:H5").setValues([["App", "Owner", "LTV", "Front", "Back", "Status", "Rate", "In panel"]])
    .setBackground(HEAD_BG).setFontColor("#FFFFFF").setFontWeight("bold").setHorizontalAlignment("center");
  const rows = [];
  for (let n = 1; n <= N; n++) {
    const t = "'" + n + "'!";
    rows.push(["=" + n, '=IF(' + t + 'B3="","",' + t + 'B3)', '=' + t + 'B' + RES, '=' + t + 'F' + RES, '=' + t + 'F' + (RES + 1),
               '=' + t + 'A' + STATUS,
               '=IF(AND(' + t + 'I' + STATUS + '="G",' + t + 'B' + (STEP3 + 1) + '<>""),' + t + 'B' + (STEP3 + 1) + '&"%","")',
               '=IF(' + t + 'F3="","",' + t + 'F3)']);
  }
  sh.getRange(6, 1, N, 8).setFormulas(rows).setHorizontalAlignment("center");
  sh.getRange(6, 6, N, 1).setHorizontalAlignment("left");
  const rng = "A6:H" + last;
  sh.setConditionalFormatRules([
    cf_(sh, rng, '=LEFT($F6,1)="✓"', GREEN), cf_(sh, rng, '=LEFT($F6,4)="NEED"', YELLOW), cf_(sh, rng, '=LEFT($F6,1)="✗"', RED)
  ]);
  sh.setFrozenRows(5);
}

function buildLoanPasteAndGo() {
  const ss = SpreadsheetApp.getActive();
  const temp = ss.insertSheet("tmp_build_" + Date.now());
  ss.getSheets().forEach(s => { if (s.getSheetId() !== temp.getSheetId()) ss.deleteSheet(s); });
  const first = ss.insertSheet("1");
  ss.deleteSheet(temp);
  buildTemplate_(first);
  SpreadsheetApp.flush();
  for (let n = 2; n <= N; n++) {
    const s = first.copyTo(ss).setName(String(n));
    s.getRange("A1").setValue("Application " + n);
    s.hideColumns(9, 3);
  }
  const summary = ss.insertSheet("Summary", 0);
  buildSummary_(summary);
  ss.setActiveSheet(summary);
  SpreadsheetApp.flush();
  ss.toast("Done: Summary + 49 application tabs", "Loan Paste & Go", 8);
}
