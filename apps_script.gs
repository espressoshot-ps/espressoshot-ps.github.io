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
  // Same palette as the dashboard tab: dark header, white rows, grey lines, red only for low ratings
  var P = { head: DARK, headInk: '#ffffff', ink: INK, muted: MUTED, band: '#ffffff', rule: LINE,
            lowBg: '#ffffff', lowInk: RED, midBg: '#ffffff', midInk: INK, hiBg: '#ffffff', hiInk: INK, callBg: SOFT };
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
  sh.getRange(1, 1, 1, 2).merge().setValue(
    SpreadsheetApp.newCellImage().setSourceUrl(LOGO_URL).setAltTextTitle('ESPRESSO SHOT').build());
  var title = 'ردود الزبائن', sub = 'كل رد جديد بيظهر هون تلقائياً';
  sh.getRange(1, 3, 1, 5).merge().setHorizontalAlignment('right').setVerticalAlignment('middle').setWrap(true)
    .setRichTextValue(SpreadsheetApp.newRichTextValue().setText(title + '\n' + sub)
      .setTextStyle(0, title.length, SpreadsheetApp.newTextStyle().setFontSize(16).setBold(true).setForegroundColor(INK).build())
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

// Dashboard laid out to fit one screen: two blocks side by side (RTL: right block first).
// Right block B:I = tiles, per-branch table, latest notes. Left block K:R = tiles, rating spread, app stats.
function buildDashboard_(ss) {
  var sm = ss.getSheetByName(SUM) || ss.insertSheet(SUM, 1);
  var C = 18, R = 21, q = "'" + DATA + "'!";
  if (sm.getMaxColumns() < C) sm.insertColumnsAfter(sm.getMaxColumns(), C - sm.getMaxColumns());
  if (sm.getMaxColumns() > C) sm.deleteColumns(C + 1, sm.getMaxColumns() - C);
  if (sm.getMaxRows() < R) sm.insertRowsAfter(sm.getMaxRows(), R - sm.getMaxRows());
  if (sm.getMaxRows() > R) sm.deleteRows(R + 1, sm.getMaxRows() - R);
  var all = sm.getRange(1, 1, R, C);
  all.breakApart();
  sm.clear();
  sm.clearConditionalFormatRules();
  sm.setRightToLeft(true).setTabColor(DARK).setHiddenGridlines(true).setFrozenRows(0);
  all.setFontFamily(FONT).setFontSize(10).setFontColor(INK).setVerticalAlignment('middle').setHorizontalAlignment('center').setBackground('#ffffff');
  sm.setColumnWidth(1, 16);
  for (var c = 2; c <= C; c++) sm.setColumnWidth(c, 80);
  sm.setColumnWidth(10, 24);
  var H = { 1: 10, 2: 34, 3: 18, 4: 8, 5: 22, 6: 44, 7: 12, 8: 26, 9: 28, 10: 28, 11: 28, 12: 28, 13: 28, 14: 12, 15: 26, 16: 28, 17: 32, 18: 32, 19: 32, 20: 10, 21: 10 };
  for (var r = 1; r <= R; r++) sm.setRowHeight(r, H[r]);
  var SOLID = SpreadsheetApp.BorderStyle.SOLID, THICK = SpreadsheetApp.BorderStyle.SOLID_THICK;
  var title = function (row, col, n, text) {
    sm.getRange(row, col, 1, n).merge().setValue(text).setFontSize(12).setFontWeight('bold').setHorizontalAlignment('right');
  };
  var head = function (row, col, n) { sm.getRange(row, col, 1, n).setBackground(DARK).setFontColor('#ffffff').setFontWeight('bold'); };
  var lines = function (row, col, rows, n) { sm.getRange(row, col, rows, n).setBorder(null, null, true, null, null, true, LINE, SOLID); };

  // Title
  sm.getRange(2, 2, 1, 8).merge().setValue('ESPRESSO SHOT').setFontFamily('Jost').setFontSize(20).setFontColor('#4a4a4a').setHorizontalAlignment('right');
  sm.getRange(3, 2, 1, 8).merge().setValue('ملخص تقييمات الزبائن · يتحدّث تلقائياً مع كل رد جديد').setFontColor(MUTED).setHorizontalAlignment('right');

  // Four tiles across both blocks
  [['عدد الردود', '=COUNTA(' + q + 'B3:B)', '0', 2],
   ['متوسط التجربة العامة (من 5)', '=IFERROR(AVERAGE(' + q + 'F3:F),"—")', '0.0', 6],
   ['طلبوا التواصل معهم', '=COUNTIF(' + q + 'J3:J,"نعم")', '0', 11],
   ['تقييم عام 2 أو أقل', '=COUNTIF(' + q + 'F3:F,"<=2")', '0', 15]
  ].forEach(function (t) {
    sm.getRange(5, t[3], 1, 4).merge().setValue(t[0]).setFontColor(MUTED).setVerticalAlignment('bottom');
    sm.getRange(6, t[3], 1, 4).merge().setFormula(t[1]).setNumberFormat(t[2]).setFontSize(22);
    sm.getRange(5, t[3], 2, 4).setBackground(SOFT).setBorder(true, true, true, true, null, null, '#ffffff', THICK);
  });

  // Right block: per-branch table (rows 8-13)
  title(8, 2, 8, 'حسب الفرع');
  sm.getRange(9, 2, 1, 2).merge().setValue('الفرع');
  sm.getRange(9, 4, 1, 6).setValues([['الردود', 'الجودة', 'السرعة', 'النظافة', 'العامة', 'من 5']]);
  head(9, 2, 8);
  BRANCHES.concat(['كل الفروع']).forEach(function (b, i) {
    var r = 10 + i, isTotal = i === BRANCHES.length;
    sm.getRange(r, 2, 1, 2).merge().setValue(b).setFontWeight('bold').setHorizontalAlignment('right');
    sm.getRange(r, 4).setFormula(isTotal ? '=COUNTA(' + q + 'B3:B)' : '=COUNTIF(' + q + '$B$3:$B,$B' + r + ')');
    ['C', 'D', 'E', 'F'].forEach(function (col, k) {
      var rng = q + col + '3:' + col;
      sm.getRange(r, 5 + k).setFormula(isTotal ? '=IFERROR(AVERAGE(' + rng + '),"—")'
        : '=IFERROR(AVERAGEIF(' + q + '$B$3:$B,$B' + r + ',' + rng + '),"—")');
    });
    sm.getRange(r, 9).setFormula('=IF(ISNUMBER(H' + r + '),REPT("●",ROUND(H' + r + '))&REPT("○",5-ROUND(H' + r + ')),"")').setFontColor('#4a4a4a').setFontSize(11);
    if (isTotal) sm.getRange(r, 2, 1, 8).setBackground('#ececec').setFontWeight('bold');
  });
  sm.getRange(10, 5, 4, 4).setNumberFormat('0.0');
  lines(10, 2, 4, 8);

  // Left block: spread of the overall rating (rows 8-13)
  title(8, 11, 8, 'توزيع التجربة العامة');
  [[5, 'ممتاز', '#4a4a4a'], [4, 'جيد جداً', '#4a4a4a'], [3, 'مقبول', '#9a9a9a'], [2, 'ضعيف', RED], [1, 'بحاجة لتحسين', RED]]
    .forEach(function (d, i) {
      var r = 9 + i;
      sm.getRange(r, 11, 1, 2).merge().setValue(d[0] + '  ·  ' + d[1]).setHorizontalAlignment('right');
      sm.getRange(r, 13).setFormula('=COUNTIF(' + q + 'F3:F,' + d[0] + ')').setFontWeight('bold');
      sm.getRange(r, 14, 1, 5).merge().setFormula('=IF(M' + r + '>0,SPARKLINE(M' + r +
        ',{"charttype","bar";"max",MAX($M$9:$M$13);"rtl",TRUE;"color1","' + d[2] + '"}),"")');
    });
  lines(9, 11, 5, 8);

  // Right block: latest notes (rows 15-19)
  title(15, 2, 8, 'آخر الملاحظات');
  sm.getRange(16, 2).setValue('التقييم');
  sm.getRange(16, 3, 1, 5).merge().setValue('الملاحظة').setHorizontalAlignment('right');
  sm.getRange(16, 8, 1, 2).merge().setValue('الفرع · التاريخ');
  head(16, 2, 8);
  var src = 'SORT(FILTER({' + q + 'A3:A,' + q + 'B3:B,' + q + 'F3:F,' + q + 'G3:G},' + q + 'G3:G<>""),1,FALSE)';
  var pick = function (n, k) { return 'INDEX(' + src + ',' + n + ',' + k + ')'; };
  for (var n = 1; n <= 3; n++) {
    var row = 16 + n;
    sm.getRange(row, 2).setFormula('=IFERROR(' + pick(n, 3) + ',"")').setNumberFormat('0').setFontWeight('bold').setFontSize(12);
    sm.getRange(row, 3, 1, 5).merge().setFormula('=IFERROR(' + pick(n, 4) + ',"")').setHorizontalAlignment('right').setWrap(true);
    sm.getRange(row, 8, 1, 2).merge().setFormula('=IFERROR(' + pick(n, 2) + '&"  ·  "&TEXT(' + pick(n, 1) + ',"dd/mm"),"")').setFontColor(MUTED);
  }
  lines(17, 2, 3, 8);

  // Left block: app stats (rows 15-19)
  title(15, 11, 8, 'تطبيق ESPRESSO SHOT');
  [['نسبة اللي جرّبوا التطبيق', '=IFERROR(COUNTIF(' + q + 'K3:K,"نعم")/COUNTA(' + q + 'K3:K),"—")', '0%'],
   ['متوسط تقييم التطبيق (من 5)', '=IFERROR(AVERAGE(' + q + 'L3:L),"—")', '0.0'],
   ['أكثر سبب لعدم التجربة', '=IFERROR(INDEX(QUERY({' + q + 'M3:M},"select Col1, count(Col1) where Col1 <> \'\' group by Col1 order by count(Col1) desc limit 1",0),2,1),"—")', '@']
  ].forEach(function (a, i) {
    var r = 17 + i;
    sm.getRange(r, 11, 1, 4).merge().setValue(a[0]).setHorizontalAlignment('right').setFontColor(MUTED);
    sm.getRange(r, 15, 1, 4).merge().setFormula(a[1]).setNumberFormat(a[2]).setHorizontalAlignment('right').setFontWeight('bold').setFontSize(12);
  });
  sm.getRange(16, 11, 1, 8).setBorder(null, null, true, null, null, null, DARK, SOLID);
  lines(17, 11, 3, 8);

  sm.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=AND(ISNUMBER(E10),E10<3)')
      .setFontColor(RED).setBold(true).setRanges([sm.getRange(10, 5, 4, 4)]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$O$6>0')
      .setFontColor(RED).setRanges([sm.getRange(6, 15, 1, 4)]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=AND(ISNUMBER($B17),$B17<=2)')
      .setFontColor(RED).setRanges([sm.getRange(17, 2, 3, 1)]).build()
  ]);
}
