// ESPRESSO SHOT feedback -> Google Sheet
// Paste this in: Sheet > Extensions > Apps Script, then Deploy > New deployment > Web app
// Execute as: Me   |   Who has access: Anyone

var HEADERS = ['التاريخ', 'الفرع', 'جودة الطلب', 'سرعة الخدمة', 'النظافة', 'التجربة العامة', 'الملاحظة', 'الاسم', 'الهاتف', 'موافقة التواصل'];

function clean(v, max) {
  var s = String(v == null ? '' : v).slice(0, max);
  // stop text like "=..." or "+970..." from being read as a formula/number
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function rating(v) {
  var n = parseInt(v, 10);
  return n >= 1 && n <= 5 ? n : '';
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    var d = JSON.parse(e.postData.contents);
    if (d.website) return reply({ ok: true }); // bot trap: silently drop
    if (!d.branch || !rating(d.overall)) return reply({ ok: false, error: 'missing fields' });

    lock.waitLock(10000);
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(HEADERS);
      sheet.setFrozenRows(1);
    }
    sheet.appendRow([
      new Date(), clean(d.branch, 60),
      rating(d.quality), rating(d.speed), rating(d.clean), rating(d.overall),
      clean(d.note, 1500), clean(d.name, 80), clean(d.phone, 30),
      d.consent === true ? 'نعم' : 'لا'
    ]);
    return reply({ ok: true });
  } catch (err) {
    return reply({ ok: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
}

function reply(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// Run from the editor to (re)style the responses tab and rebuild the dashboard tab.
var INK = '#262626', MUTED = '#737373', LINE = '#e0e0e0', DARK = '#3d3d3d', SOFT = '#f3f3f3', RED = '#b3261e';
var FONT = 'Readex Pro', DATA = 'الردود', SUM = 'ملخص';
var BRANCHES = ['الماصيون', 'الطيرة', 'وسط البلد'];

function formatSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheets()[0];
  styleResponses_(sh);
  buildDashboard_(ss);
  ss.setActiveSheet(ss.getSheetByName(SUM));
}

function rule_(formula, bg, fg, range) {
  return SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied(formula)
    .setBackground(bg).setFontColor(fg).setRanges([range]).build();
}

function styleResponses_(sh) {
  var N = HEADERS.length;
  sh.setName(DATA).setTabColor(DARK);
  sh.setRightToLeft(true);
  if (sh.getLastRow() === 0) sh.appendRow(HEADERS);
  if (sh.getMaxColumns() > N) sh.deleteColumns(N + 1, sh.getMaxColumns() - N);
  var rows = sh.getMaxRows(), body = rows - 1;
  sh.getBandings().forEach(function (b) { b.remove(); });
  if (sh.getFilter()) sh.getFilter().remove();

  sh.getRange(1, 1, rows, N).setFontFamily(FONT).setFontSize(11).setFontColor(INK).setFontWeight('normal')
    .setVerticalAlignment('middle').setHorizontalAlignment('center').setWrap(false).setBackground('#ffffff')
    .setBorder(null, null, null, null, false, true, LINE, SpreadsheetApp.BorderStyle.SOLID);
  sh.getRange(1, 1, 1, N).setBackground(DARK).setFontColor('#ffffff').setFontWeight('bold');
  sh.setFrozenRows(1);
  sh.setRowHeight(1, 46);
  [150, 120, 135, 140, 120, 150, 380, 150, 140, 160].forEach(function (w, i) { sh.setColumnWidth(i + 1, w); });

  sh.getRange(2, 1, body, 1).setNumberFormat('dd/mm/yyyy  hh:mm').setFontColor(MUTED).setFontSize(10);
  sh.getRange(2, 2, body, 1).setFontWeight('bold');
  sh.getRange(2, 3, body, 4).setFontWeight('bold').setFontSize(12);
  sh.getRange(2, 7, body, 1).setHorizontalAlignment('right').setWrap(true);
  sh.getRange(2, 8, body, 1).setHorizontalAlignment('right');
  sh.getRange(2, 9, body, 1).setNumberFormat('@');
  sh.setHiddenGridlines(true);

  var ratings = sh.getRange(2, 3, body, 4), contact = sh.getRange(2, 8, body, 3);
  sh.setConditionalFormatRules([
    rule_('=AND(ISNUMBER(C2),C2<=2)', '#fbeceb', RED, ratings),
    rule_('=AND(ISNUMBER(C2),C2=3)', '#fff6dd', '#7a5a00', ratings),
    rule_('=AND(ISNUMBER(C2),C2>=4)', '#eaf3ec', '#1e6b3a', ratings),
    rule_('=$J2="نعم"', '#fff6dd', INK, contact)
  ]);
  sh.getRange(1, 1, rows, N).createFilter();
}

function buildDashboard_(ss) {
  var sm = ss.getSheetByName(SUM) || ss.insertSheet(SUM, 1);
  var C = 9, R = 30, q = "'" + DATA + "'!";
  if (sm.getMaxColumns() < C) sm.insertColumnsAfter(sm.getMaxColumns(), C - sm.getMaxColumns());
  if (sm.getMaxColumns() > C) sm.deleteColumns(C + 1, sm.getMaxColumns() - C);
  if (sm.getMaxRows() < R) sm.insertRowsAfter(sm.getMaxRows(), R - sm.getMaxRows());
  if (sm.getMaxRows() > R) sm.deleteRows(R + 1, sm.getMaxRows() - R);
  var all = sm.getRange(1, 1, R, C);
  all.breakApart();
  sm.clear();
  sm.clearConditionalFormatRules();
  sm.setRightToLeft(true).setTabColor(DARK);
  sm.setHiddenGridlines(true);
  sm.setFrozenRows(0);
  all.setFontFamily(FONT).setFontSize(11).setFontColor(INK).setVerticalAlignment('middle').setHorizontalAlignment('center');
  sm.setColumnWidth(1, 28);
  for (var c = 2; c <= C; c++) sm.setColumnWidth(c, 124);
  var heights = { 1: 14, 2: 44, 3: 26, 4: 16, 5: 34, 6: 62, 7: 26, 8: 36, 9: 40, 14: 26, 15: 36, 21: 26, 22: 36, 23: 36 };
  for (var r = 1; r <= R; r++) sm.setRowHeight(r, heights[r] || 36);
  var SOLID = SpreadsheetApp.BorderStyle.SOLID, THICK = SpreadsheetApp.BorderStyle.SOLID_THICK;

  // Title
  sm.getRange(2, 2, 1, 8).merge().setValue('ESPRESSO SHOT').setFontFamily('Jost').setFontSize(22)
    .setFontColor('#4a4a4a').setHorizontalAlignment('right');
  sm.getRange(3, 2, 1, 8).merge().setValue('ملخص تقييمات الزبائن · يتحدّث تلقائياً مع كل رد جديد')
    .setFontColor(MUTED).setFontSize(10).setHorizontalAlignment('right');

  // Four tiles
  var tiles = [
    ['عدد الردود', '=COUNTA(' + q + 'B2:B)', '0'],
    ['متوسط التجربة العامة (من 5)', '=IFERROR(AVERAGE(' + q + 'F2:F),"—")', '0.0'],
    ['طلبوا التواصل معهم', '=COUNTIF(' + q + 'J2:J,"نعم")', '0'],
    ['تقييم عام 2 أو أقل', '=COUNTIF(' + q + 'F2:F,"<=2")', '0']
  ];
  tiles.forEach(function (t, i) {
    var col = 2 + i * 2;
    sm.getRange(5, col, 1, 2).merge().setValue(t[0]).setFontColor(MUTED).setFontSize(10).setVerticalAlignment('bottom');
    sm.getRange(6, col, 1, 2).merge().setFormula(t[1]).setNumberFormat(t[2]).setFontSize(28);
    sm.getRange(5, col, 2, 2).setBackground(SOFT).setBorder(true, true, true, true, null, null, '#ffffff', THICK);
  });

  // Per-branch table
  sm.getRange(8, 2, 1, 8).merge().setValue('حسب الفرع').setFontSize(13).setFontWeight('bold').setHorizontalAlignment('right');
  sm.getRange(9, 2, 1, 2).merge().setValue('الفرع');
  sm.getRange(9, 4, 1, 6).setValues([['الردود', 'جودة الطلب', 'سرعة الخدمة', 'النظافة', 'التجربة العامة', 'من 5']]);
  sm.getRange(9, 2, 1, 8).setBackground(DARK).setFontColor('#ffffff').setFontWeight('bold');
  BRANCHES.concat(['كل الفروع']).forEach(function (b, i) {
    var r = 10 + i, isTotal = i === BRANCHES.length;
    sm.getRange(r, 2, 1, 2).merge().setValue(b).setFontWeight('bold').setHorizontalAlignment('right');
    sm.getRange(r, 4).setFormula(isTotal ? '=COUNTA(' + q + 'B2:B)' : '=COUNTIF(' + q + '$B$2:$B,$B' + r + ')');
    ['C', 'D', 'E', 'F'].forEach(function (col, k) {
      var rng = q + col + '2:' + col;
      sm.getRange(r, 5 + k).setFormula(isTotal ? '=IFERROR(AVERAGE(' + rng + '),"—")'
        : '=IFERROR(AVERAGEIF(' + q + '$B$2:$B,$B' + r + ',' + rng + '),"—")');
    });
    sm.getRange(r, 9).setFormula('=IF(ISNUMBER(H' + r + '),SPARKLINE(H' + r + ',{"charttype","bar";"max",5;"rtl",TRUE;"color1","#4a4a4a"}),"")');
    if (isTotal) sm.getRange(r, 2, 1, 8).setBackground('#ececec').setFontWeight('bold');
  });
  sm.getRange(10, 5, 4, 4).setNumberFormat('0.0');
  sm.getRange(10, 2, 4, 8).setBorder(null, null, true, null, null, true, LINE, SOLID);

  // Distribution of the overall rating
  sm.getRange(15, 2, 1, 8).merge().setValue('توزيع التجربة العامة').setFontSize(13).setFontWeight('bold').setHorizontalAlignment('right');
  [[5, 'ممتاز', '#4a4a4a'], [4, 'جيد جداً', '#4a4a4a'], [3, 'مقبول', '#c28a00'], [2, 'ضعيف', RED], [1, 'بحاجة لتحسين', RED]]
    .forEach(function (d, i) {
      var r = 16 + i;
      sm.getRange(r, 2, 1, 2).merge().setValue(d[0] + '  ·  ' + d[1]).setHorizontalAlignment('right');
      sm.getRange(r, 4).setFormula('=COUNTIF(' + q + 'F2:F,' + d[0] + ')').setFontWeight('bold');
      sm.getRange(r, 5, 1, 5).merge().setFormula('=IF(D' + r + '>0,SPARKLINE(D' + r +
        ',{"charttype","bar";"max",MAX($D$16:$D$20);"rtl",TRUE;"color1","' + d[2] + '"}),"")');
    });
  sm.getRange(16, 2, 5, 8).setBorder(null, null, true, null, null, true, LINE, SOLID);

  // Latest notes
  sm.getRange(22, 2, 1, 8).merge().setValue('آخر الملاحظات').setFontSize(13).setFontWeight('bold').setHorizontalAlignment('right');
  sm.getRange(23, 2, 1, 3).setValues([['التاريخ', 'الفرع', 'التقييم']]);
  sm.getRange(23, 5, 1, 5).merge().setValue('الملاحظة');
  sm.getRange(23, 2, 1, 8).setBackground(DARK).setFontColor('#ffffff').setFontWeight('bold');
  var src = 'SORT(FILTER({' + q + 'A2:A,' + q + 'B2:B,' + q + 'F2:F,' + q + 'G2:G},' + q + 'G2:G<>""),1,FALSE)';
  for (var n = 1; n <= 5; n++) {
    var row = 23 + n;
    for (var k = 1; k <= 3; k++) sm.getRange(row, 1 + k).setFormula('=IFERROR(INDEX(' + src + ',' + n + ',' + k + '),"")');
    sm.getRange(row, 5, 1, 5).merge().setFormula('=IFERROR(INDEX(' + src + ',' + n + ',4),"")')
      .setHorizontalAlignment('right').setWrap(true);
  }
  sm.getRange(24, 2, 5, 1).setNumberFormat('dd/mm  hh:mm').setFontColor(MUTED).setFontSize(10);
  sm.getRange(24, 4, 5, 1).setFontWeight('bold');
  sm.getRange(24, 2, 5, 8).setBorder(null, null, true, null, null, true, LINE, SOLID);

  sm.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=AND(ISNUMBER(E10),E10<3)')
      .setFontColor(RED).setBold(true).setRanges([sm.getRange(10, 5, 4, 4)]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$H$6>0')
      .setFontColor(RED).setRanges([sm.getRange(6, 8, 1, 2)]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=AND(ISNUMBER($D24),$D24<=2)')
      .setFontColor(RED).setRanges([sm.getRange(24, 4, 5, 1)]).build()
  ]);
}
