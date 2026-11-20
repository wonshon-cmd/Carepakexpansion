/**
 * Loan Paste & Go — builds the whole workbook inside Google Sheets.
 *
 * How to use:
 * 1. Open a BLANK Google Sheet (sheets.google.com → Blank).
 * 2. Extensions → Apps Script. Delete what is there, paste this whole file, click Save.
 * 3. Pick "buildLoanPasteAndGo" in the function menu and click Run.
 *    Google asks for permission the first time (Review permissions → your account →
 *    Advanced → Go to project → Allow). It only edits this spreadsheet.
 * 4. Go back to the sheet: a Summary tab and tabs 1–49 appear (about 1 minute).
 */
const N = 49;
const NAMES = ["Rishi", "Senthil", "Tobe", "Gabriela", "Joe"];
const GREEN = ["#C6EFCE", "#006100"], YELLOW = ["#FFEB9C", "#9C5700"], RED = ["#FFC7CE", "#9C0006"];
const PASTE_BG = "#EAF1FF", CALC_BG = "#F3F5F8", HEAD_BG = "#18202C", BLUE = "#1F5FD6", MUTED = "#5D6878";

function clean_(ref) {
  const t = 'SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(TRIM(' + ref + '),"$",""),",","")," ",""),CHAR(160),"")';
  return '=IF(' + ref + '="","",IFERROR(VALUE(' + t + '),""))';
}

function head_(sh, a1, text) {
  const r = sh.getRange(a1).merge().setBackground(HEAD_BG).setFontColor("#FFFFFF").setFontWeight("bold").setFontSize(12);
  r.getCell(1, 1).setValue(text);
}

function pasteCell_(rng, top) {
  rng.setBackground(PASTE_BG).setHorizontalAlignment("center");
  if (top) rng.setBorder(true, true, true, true, null, null, BLUE, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
}

function cf_(sh, a1, formula, colors) {
  return SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied(formula)
    .setBackground(colors[0]).setFontColor(colors[1]).setRanges([sh.getRange(a1)]).build();
}

function buildTemplate_(sh) {
  [190, 120, 90, 20, 220, 120, 90, 20, 30].forEach((w, i) => sh.setColumnWidth(i + 1, w));
  sh.getRange("A1").setValue("Application 1").setFontSize(20).setFontWeight("bold");
  sh.getRange("A2:G2").merge().setFontColor(MUTED).setFontSize(10);
  sh.getRange("A2").setValue("Copy the blue $ amounts in the bank panel → click the BLUE-BORDERED cell → paste with Ctrl/⌘+Shift+V.");
  sh.getRange("A3").setValue("Owner (pick your name first)").setFontWeight("bold");
  pasteCell_(sh.getRange("B3"), false);
  sh.getRange("B3").setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(NAMES, true).setAllowInvalid(true).build());
  sh.getRange("E3").setValue("Clicked in bank panel?").setFontWeight("bold");
  pasteCell_(sh.getRange("F3"), false);
  sh.getRange("F3").setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(["Yes", "No"], true).build());

  // STEP 1 · LTV (paste into B6 → fills B6:B8)
  head_(sh, "A5:C5", "STEP 1 · Loan-to-Value page");
  ["Loan Amount (L)", "House Price (H)", "Appraisal Value (V)"].forEach((lab, i) => {
    const r = 6 + i;
    sh.getRange("A" + r).setValue(lab);
    pasteCell_(sh.getRange("B" + r), i === 0);
    sh.getRange("C" + r).setFormula(clean_("B" + r)).setNumberFormat("#,##0").setBackground(CALC_BG).setFontColor(MUTED);
  });
  sh.getRange("B9").setValue("▲ paste in top cell").setFontColor(BLUE).setFontSize(9).setFontWeight("bold").setHorizontalAlignment("center");
  sh.getRange("A11").setValue("LTV → type in panel").setFontWeight("bold");
  sh.getRange("C11").setFormula('=IF(COUNT(C6:C8)<3,"",C6/MIN(C7:C8))').setNumberFormat("0.00%").setFontColor(MUTED);
  sh.getRange("B11").setFormula('=IF(C11="","–",ROUND(C11*100,0))').setFontSize(26).setFontWeight("bold").setHorizontalAlignment("center");
  sh.setRowHeight(11, 48);

  // STEP 2 · Front/Back (paste into F6 → fills F6:F10)
  head_(sh, "E5:G5", "STEP 2 · Front / Back Ratio pages (only if YELLOW)");
  ["Yearly Income (I)", "Additional Income (A)", "Yearly Mortgage Payment (P)", "Debt Expenses (DE)", "Other Taxes (OT)"].forEach((lab, i) => {
    const r = 6 + i;
    sh.getRange("E" + r).setValue(lab);
    pasteCell_(sh.getRange("F" + r), i === 0);
    sh.getRange("G" + r).setFormula(clean_("F" + r)).setNumberFormat("#,##0").setBackground(CALC_BG).setFontColor(MUTED);
  });
  sh.getRange("F11:G11").merge().setFontColor(BLUE).setFontSize(9).setFontWeight("bold");
  sh.getRange("F11").setValue("▲ Front page: paste 3 · Back page: paste all 5");
  sh.getRange("E12").setValue("Front → type in panel").setFontWeight("bold");
  sh.getRange("G12").setFormula('=IF(COUNT(G6:G8)<3,"",G8/(G6+G7))').setNumberFormat("0.00%").setFontColor(MUTED);
  sh.getRange("F12").setFormula('=IF(G12="","–",ROUND(G12*100,0))');
  sh.getRange("E13").setValue("Back → type in panel").setFontWeight("bold");
  sh.getRange("G13").setFormula('=IF(COUNT(G6:G10)<5,"",(G9+G10+G8)/(G6+G7))').setNumberFormat("0.00%").setFontColor(MUTED);
  sh.getRange("F13").setFormula('=IF(G13="","–",ROUND(G13*100,0))');
  sh.getRange("F12:F13").setFontSize(22).setFontWeight("bold").setHorizontalAlignment("center");
  sh.setRowHeight(12, 40); sh.setRowHeight(13, 40);

  // hidden state: G approved / Y need data / R not qualified
  sh.getRange("I15").setFormula('=IF(C11="","",IF(C11>1,"R",IF(C11<0.85,"G",IF(OR(G12="",G13=""),"Y",IF(AND(G12<0.25,G13<0.35),"G","R")))))');
  sh.getRange("A15:G16").merge().setFontSize(18).setFontWeight("bold").setHorizontalAlignment("center").setVerticalAlignment("middle");
  sh.getRange("A15").setFormula('=IF(I15="","PASTE L, H, V ABOVE",IF(I15="G","✓ APPROVED — click Approve Loan",IF(I15="Y","NEED DATA — paste Front/Back in STEP 2","✗ NOT QUALIFIED — click Reject Loan")))');
  sh.setRowHeight(15, 34); sh.setRowHeight(16, 34);

  // STEP 3 · puzzle + decision
  head_(sh, "A18:C18", "STEP 3 · Risk Assessment + Decision");
  sh.getRange("A19").setValue("Words found (0–5)");
  pasteCell_(sh.getRange("B19"), true);
  sh.getRange("B19").setDataValidation(SpreadsheetApp.newDataValidation().requireNumberBetween(0, 5).build());
  sh.getRange("A20:C20").merge().setFontWeight("bold").setFontSize(13);
  sh.getRange("A20").setFormula('=IF(I15<>"G","",IF(B19="","Do the puzzle, type words found","Click "&B19&"% → Approve Loan → Next Application"))');
  sh.hideColumns(9);

  sh.setConditionalFormatRules([
    cf_(sh, "A15:G16", '=$I$15="G"', GREEN), cf_(sh, "A15:G16", '=$I$15="Y"', YELLOW), cf_(sh, "A15:G16", '=$I$15="R"', RED),
    cf_(sh, "B11", '=AND($C$11<>"",$C$11>1)', RED), cf_(sh, "B11", '=AND($C$11<>"",$C$11>=0.85)', YELLOW), cf_(sh, "B11", '=$C$11<>""', GREEN),
    cf_(sh, "F12", '=AND($G$12<>"",$G$12>=0.25)', RED), cf_(sh, "F12", '=$G$12<>""', GREEN),
    cf_(sh, "F13", '=AND($G$13<>"",$G$13>=0.35)', RED), cf_(sh, "F13", '=$G$13<>""', GREEN),
    cf_(sh, "E5:G5", '=$I$15="Y"', YELLOW), cf_(sh, "F3", '=$F$3="Yes"', GREEN)
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
    rows.push(["=" + n, '=IF(' + t + 'B3="","",' + t + 'B3)', '=' + t + 'B11', '=' + t + 'F12', '=' + t + 'F13', '=' + t + 'A15',
               '=IF(AND(' + t + 'I15="G",' + t + 'B19<>""),' + t + 'B19&"%","")', '=IF(' + t + 'F3="","",' + t + 'F3)']);
  }
  sh.getRange(6, 1, N, 8).setFormulas(rows.map(r => r.map(String))).setHorizontalAlignment("center");
  sh.getRange(6, 6, N, 1).setHorizontalAlignment("left");
  const rng = "A6:H" + last;
  sh.setConditionalFormatRules([
    cf_(sh, rng, '=LEFT($F6,1)="✓"', GREEN), cf_(sh, rng, '=LEFT($F6,4)="NEED"', YELLOW), cf_(sh, rng, '=LEFT($F6,1)="✗"', RED)
  ]);
  sh.setFrozenRows(5);
}

function buildLoanPasteAndGo() {
  const ss = SpreadsheetApp.getActive();
  // Start clean: keep one sheet, remove the rest.
  const sheets = ss.getSheets();
  const first = sheets[0];
  sheets.slice(1).forEach(s => ss.deleteSheet(s));
  first.clear(); first.clearConditionalFormatRules(); first.getRange("A:Z").breakApart();
  first.setName("1");
  buildTemplate_(first);
  SpreadsheetApp.flush();
  for (let n = 2; n <= N; n++) {
    const s = first.copyTo(ss).setName(String(n));
    s.getRange("A1").setValue("Application " + n);
    s.hideColumns(9);
  }
  const summary = ss.insertSheet("Summary", 0);
  buildSummary_(summary);
  ss.setActiveSheet(summary);
  SpreadsheetApp.flush();
  ss.toast("Done: Summary + 49 application tabs", "Loan Paste & Go", 8);
}
