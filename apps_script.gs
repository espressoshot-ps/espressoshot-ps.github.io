// ESPRESSO SHOT feedback -> Google Sheet
// Paste this in: Sheet > Extensions > Apps Script, then Deploy > New deployment > Web app
// Execute as: Me   |   Who has access: Anyone

var HEADERS = ['التاريخ', 'الفرع', 'جودة الطلب', 'سرعة الخدمة', 'النظافة', 'التجربة العامة', 'الملاحظة', 'الاسم', 'الهاتف', 'موافقة التواصل', 'جرّب التطبيق', 'تقييم التطبيق', 'سبب عدم التجربة', 'رأيه بالتطبيق'];

// Palestinian/Israeli mobile -> 05XXXXXXXX, or '' if not valid
function normPhone_(raw) {
  var s = String(raw == null ? '' : raw).replace(/[\s\-().]/g, '');
  if (s.charAt(0) === '+') s = '00' + s.slice(1);
  var m = s.match(/^(?:00)?(?:970|972)(5\d{8})$/);
  if (m) s = '0' + m[1];
  if (/^5\d{8}$/.test(s)) s = '0' + s;
  return /^05[012345689]\d{7}$/.test(s) ? s : '';
}

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
      sheet.appendRow(['ESPRESSO SHOT']);
      sheet.appendRow(HEADERS);
      sheet.setFrozenRows(2);
    }
    sheet.appendRow([
      new Date(), clean(d.branch, 60),
      rating(d.quality), rating(d.speed), rating(d.clean), rating(d.overall),
      clean(d.note, 1500), clean(d.name, 80), normPhone_(d.phone) ? "'" + normPhone_(d.phone) : '',
      d.consent === true ? 'نعم' : 'لا',
      d.app === 'yes' ? 'نعم' : (d.app === 'no' ? 'لا' : ''),
      d.app === 'yes' ? rating(d.appRating) : '',
      d.app === 'no' ? clean(d.appWhy, 60) : '',
      clean(d.app === 'yes' ? d.appNote : d.appWhyNote, 1000)
    ]);
    sheet.getRange(sheet.getLastRow(), 1).setNumberFormat('dd/mm/yyyy  hh:mm');
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

var LOGO_URL = 'https://espressoshot-ps.github.io/logo.png';

// Responses tab: logo banner in row 1, column headers in row 2, data from row 3.
function styleResponses_(sh) {
  var N = HEADERS.length;
  var P = { head: '#ebe7e1', headInk: '#3d3d3d', ink: '#3a3a3a', muted: '#8a857e', band: '#faf8f5', rule: '#ece8e2',
            lowBg: '#f7e4e1', lowInk: '#a4433a', midBg: '#f8efdc', midInk: '#86681f', hiBg: '#e6efe8', hiInk: '#3f6e4f', callBg: '#fbf4e4' };
  sh.setName(DATA).setTabColor(DARK);
  sh.setRightToLeft(true);
  if (sh.getMaxColumns() < N) sh.insertColumnsAfter(sh.getMaxColumns(), N - sh.getMaxColumns());
  if (sh.getMaxColumns() > N) sh.deleteColumns(N + 1, sh.getMaxColumns() - N);
  if (sh.getLastRow() === 0) sh.appendRow(HEADERS);
  if (String(sh.getRange(1, 1).getValue()) === HEADERS[0]) sh.insertRowBefore(1);   // one-time move to make room for the banner
  sh.getRange(2, 1, 1, N).setValues([HEADERS]);
  var rows = sh.getMaxRows(), body = rows - 2;
  sh.getBandings().forEach(function (x) { x.remove(); });
  if (sh.getFilter()) sh.getFilter().remove();
  sh.getRange(1, 1, 1, N).breakApart();

  sh.getRange(1, 1, rows, N).setFontFamily(FONT).setFontSize(11).setFontColor(P.ink).setFontWeight('normal')
    .setVerticalAlignment('middle').setHorizontalAlignment('center').setWrap(false).setBackground('#ffffff')
    .setBorder(false, false, false, false, false, false);
  [150, 120, 135, 140, 120, 150, 460, 150, 140, 160, 130, 130, 200, 320].forEach(function (w, i) { sh.setColumnWidth(i + 1, w); });

  // Banner
  sh.setRowHeight(1, 84);
  sh.getRange(1, 1, 1, 2).merge().setFormula('=IMAGE("' + LOGO_URL + '",1)');
  var title = 'ردود الزبائن', sub = 'كل رد جديد بيظهر هون تلقائياً';
  sh.getRange(1, 3, 1, 5).merge().setHorizontalAlignment('right').setWrap(true)
    .setRichTextValue(SpreadsheetApp.newRichTextValue().setText(title + '\n' + sub)
      .setTextStyle(0, title.length, SpreadsheetApp.newTextStyle().setFontSize(16).setBold(true).setForegroundColor(P.headInk).build())
      .setTextStyle(title.length + 1, title.length + 1 + sub.length, SpreadsheetApp.newTextStyle().setFontSize(10).setForegroundColor(P.muted).build())
      .build());
  sh.getRange(1, 1, 1, N).setBorder(null, null, true, null, null, null, P.rule, SpreadsheetApp.BorderStyle.SOLID);

  // Column headers
  sh.getRange(2, 1, 1, N).setBackground(P.head).setFontColor(P.headInk).setFontWeight('bold').setFontSize(11);
  sh.setRowHeight(2, 44);
  sh.setFrozenRows(2);

  // Body
  sh.getRange(3, 1, body, N).setBorder(null, null, null, null, null, true, P.rule, SpreadsheetApp.BorderStyle.SOLID);
  sh.getRange(3, 1, body, 1).setNumberFormat('dd/mm/yyyy  hh:mm').setFontColor(P.muted).setFontSize(10);
  sh.getRange(3, 2, body, 1).setFontWeight('bold');
  sh.getRange(3, 3, body, 4).setFontWeight('bold').setFontSize(12);
  sh.getRange(3, 7, body, 1).setHorizontalAlignment('right').setWrap(true);
  sh.getRange(3, 8, body, 1).setHorizontalAlignment('right');
  sh.getRange(3, 9, body, 1).setNumberFormat('@');
  sh.getRange(3, 12, body, 1).setFontWeight('bold').setFontSize(12);
  sh.getRange(3, 13, body, 2).setHorizontalAlignment('right');
  sh.getRange(3, 14, body, 1).setWrap(true);
  sh.setRowHeights(3, body, 44);
  sh.getRange(3, 1, body, N).applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, false, false)
    .setFirstRowColor('#ffffff').setSecondRowColor(P.band);
  sh.setHiddenGridlines(true);

  var ratings = sh.getRange(3, 3, body, 4), contact = sh.getRange(3, 8, body, 3), appR = sh.getRange(3, 12, body, 1);
  sh.setConditionalFormatRules([
    rule_('=AND(ISNUMBER(C3),C3<=2)', P.lowBg, P.lowInk, ratings),
    rule_('=AND(ISNUMBER(C3),C3=3)', P.midBg, P.midInk, ratings),
    rule_('=AND(ISNUMBER(C3),C3>=4)', P.hiBg, P.hiInk, ratings),
    rule_('=$J3="نعم"', P.callBg, P.ink, contact),
    rule_('=AND(ISNUMBER(L3),L3<=2)', P.lowBg, P.lowInk, appR),
    rule_('=AND(ISNUMBER(L3),L3=3)', P.midBg, P.midInk, appR),
    rule_('=AND(ISNUMBER(L3),L3>=4)', P.hiBg, P.hiInk, appR)
  ]);
  sh.getRange(2, 1, rows - 1, N).createFilter();
}

function buildDashboard_(ss) {
  var sm = ss.getSheetByName(SUM) || ss.insertSheet(SUM, 1);
  var C = 9, R = 35, q = "'" + DATA + "'!";
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
    ['عدد الردود', '=COUNTA(' + q + 'B3:B)', '0'],
    ['متوسط التجربة العامة (من 5)', '=IFERROR(AVERAGE(' + q + 'F3:F),"—")', '0.0'],
    ['طلبوا التواصل معهم', '=COUNTIF(' + q + 'J3:J,"نعم")', '0'],
    ['تقييم عام 2 أو أقل', '=COUNTIF(' + q + 'F3:F,"<=2")', '0']
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
  sm.getRange(9, 4, 1, 6).setValues([['الردود', 'جودة الطلب', 'سرعة الخدمة', 'النظافة', 'التجربة العامة', 'من 5 نقاط']]);
  sm.getRange(9, 2, 1, 8).setBackground(DARK).setFontColor('#ffffff').setFontWeight('bold');
  BRANCHES.concat(['كل الفروع']).forEach(function (b, i) {
    var r = 10 + i, isTotal = i === BRANCHES.length;
    sm.getRange(r, 2, 1, 2).merge().setValue(b).setFontWeight('bold').setHorizontalAlignment('right');
    sm.getRange(r, 4).setFormula(isTotal ? '=COUNTA(' + q + 'B3:B)' : '=COUNTIF(' + q + '$B$3:$B,$B' + r + ')');
    ['C', 'D', 'E', 'F'].forEach(function (col, k) {
      var rng = q + col + '3:' + col;
      sm.getRange(r, 5 + k).setFormula(isTotal ? '=IFERROR(AVERAGE(' + rng + '),"—")'
        : '=IFERROR(AVERAGEIF(' + q + '$B$3:$B,$B' + r + ',' + rng + '),"—")');
    });
    sm.getRange(r, 9).setFormula('=IF(ISNUMBER(H' + r + '),REPT("●",ROUND(H' + r + '))&REPT("○",5-ROUND(H' + r + ')),"")')
      .setFontColor('#4a4a4a').setFontSize(14);
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
      sm.getRange(r, 4).setFormula('=COUNTIF(' + q + 'F3:F,' + d[0] + ')').setFontWeight('bold');
      sm.getRange(r, 5, 1, 5).merge().setFormula('=IF(D' + r + '>0,SPARKLINE(D' + r +
        ',{"charttype","bar";"max",MAX($D$16:$D$20);"rtl",TRUE;"color1","' + d[2] + '"}),"")');
    });
  sm.getRange(16, 2, 5, 8).setBorder(null, null, true, null, null, true, LINE, SOLID);

  // Latest notes: rating, note, branch, date (read right to left)
  sm.getRange(22, 2, 1, 8).merge().setValue('آخر الملاحظات').setFontSize(13).setFontWeight('bold').setHorizontalAlignment('right');
  sm.getRange(23, 2).setValue('التقييم');
  sm.getRange(23, 3, 1, 5).merge().setValue('الملاحظة').setHorizontalAlignment('right');
  sm.getRange(23, 8, 1, 2).setValues([['الفرع', 'التاريخ']]);
  sm.getRange(23, 2, 1, 8).setBackground(DARK).setFontColor('#ffffff').setFontWeight('bold');
  var src = 'SORT(FILTER({' + q + 'A3:A,' + q + 'B3:B,' + q + 'F3:F,' + q + 'G3:G},' + q + 'G3:G<>""),1,FALSE)';
  var pick = function (n, k) { return '=IFERROR(INDEX(' + src + ',' + n + ',' + k + '),"")'; };
  for (var n = 1; n <= 5; n++) {
    var row = 23 + n;
    sm.getRange(row, 2).setFormula(pick(n, 3));
    sm.getRange(row, 3, 1, 5).merge().setFormula(pick(n, 4)).setHorizontalAlignment('right').setWrap(true);
    sm.getRange(row, 8).setFormula(pick(n, 2));
    sm.getRange(row, 9).setFormula(pick(n, 1));
  }
  sm.setRowHeights(24, 5, 48);
  sm.getRange(24, 2, 5, 1).setFontWeight('bold').setFontSize(14).setNumberFormat('0');
  sm.getRange(24, 8, 5, 1).setNumberFormat('@');
  sm.getRange(24, 8, 5, 1).setFontColor(MUTED);
  sm.getRange(24, 9, 5, 1).setNumberFormat('dd/mm  hh:mm').setFontColor(MUTED).setFontSize(10);
  sm.getRange(24, 2, 5, 8).setBorder(null, null, true, null, null, true, LINE, SOLID);

  // App section
  sm.getRange(30, 2, 1, 8).merge().setValue('تطبيق ESPRESSO SHOT').setFontSize(13).setFontWeight('bold').setHorizontalAlignment('right');
  var appRows = [
    ['نسبة اللي جرّبوا التطبيق', '=IFERROR(COUNTIF(' + q + 'K3:K,"نعم")/COUNTA(' + q + 'K3:K),"—")', '0%'],
    ['متوسط تقييم التطبيق (من 5)', '=IFERROR(AVERAGE(' + q + 'L3:L),"—")', '0.0'],
    ['أكثر سبب لعدم التجربة', '=IFERROR(INDEX(QUERY({' + q + 'M3:M},"select Col1, count(Col1) where Col1 <> \'\' group by Col1 order by count(Col1) desc limit 1",0),2,1),"—")', '@']
  ];
  appRows.forEach(function (a, i) {
    var r = 31 + i;
    sm.getRange(r, 2, 1, 3).merge().setValue(a[0]).setHorizontalAlignment('right').setFontColor(MUTED);
    sm.getRange(r, 5, 1, 5).merge().setFormula(a[1]).setNumberFormat(a[2]).setHorizontalAlignment('right').setFontWeight('bold').setFontSize(13);
  });
  sm.getRange(31, 2, 3, 8).setBorder(null, null, true, null, null, true, LINE, SOLID);

  sm.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=AND(ISNUMBER(E10),E10<3)')
      .setFontColor(RED).setBold(true).setRanges([sm.getRange(10, 5, 4, 4)]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$H$6>0')
      .setFontColor(RED).setRanges([sm.getRange(6, 8, 1, 2)]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=AND(ISNUMBER($B24),$B24<=2)')
      .setFontColor(RED).setRanges([sm.getRange(24, 2, 5, 1)]).build()
  ]);
}
