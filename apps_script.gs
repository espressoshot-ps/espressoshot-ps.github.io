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

// Dashboard: white cards on a light grey canvas, one screen tall.
// Grid (RTL): B..I = right card column, J = gap, K..R = left card column. B, I, K, R are 14px inner padding.
function buildDashboard_(ss) {
  var sm = ss.getSheetByName(SUM) || ss.insertSheet(SUM, 1);
  var C = 19, R = 34, q = "'" + DATA + "'!";
  var CANVAS = '#f3f3f3', CARD = '#ffffff', HAIR = '#ececec';
  if (sm.getMaxColumns() < C) sm.insertColumnsAfter(sm.getMaxColumns(), C - sm.getMaxColumns());
  if (sm.getMaxColumns() > C) sm.deleteColumns(C + 1, sm.getMaxColumns() - C);
  if (sm.getMaxRows() < R) sm.insertRowsAfter(sm.getMaxRows(), R - sm.getMaxRows());
  if (sm.getMaxRows() > R) sm.deleteRows(R + 1, sm.getMaxRows() - R);
  sm.getCharts().forEach(function (ch) { sm.removeChart(ch); });
  var all = sm.getRange(1, 1, R, C);
  all.breakApart();
  sm.clear();
  sm.clearConditionalFormatRules();
  sm.setRightToLeft(true).setTabColor(DARK).setHiddenGridlines(true).setFrozenRows(0);
  all.setFontFamily(FONT).setFontSize(10).setFontColor(INK).setVerticalAlignment('middle').setHorizontalAlignment('center')
     .setBackground(CANVAS).setBorder(false, false, false, false, false, false);

  var W = { 1: 20, 2: 14, 9: 14, 10: 20, 11: 14, 18: 14, 19: 20 };
  for (var c = 1; c <= C; c++) sm.setColumnWidth(c, W[c] || 102);
  var H = { 1: 10, 2: 46, 3: 10, 4: 22, 5: 38, 6: 10, 7: 30, 8: 24, 9: 26, 10: 26, 11: 26, 12: 26, 13: 8,
            14: 10, 15: 30, 16: 24, 17: 26, 18: 26, 19: 26, 20: 8, 21: 10 };
  for (var r = 1; r <= R; r++) sm.setRowHeight(r, H[r] || 20);

  var box = function (r1, c1, rows, cols) { return sm.getRange(r1, c1, rows, cols).setBackground(CARD); };
  var cardTitle = function (row, col, text, sub) {
    var t = text + (sub ? '   ' + sub : '');
    var rt = SpreadsheetApp.newRichTextValue().setText(t)
      .setTextStyle(0, text.length, SpreadsheetApp.newTextStyle().setBold(true).setFontSize(12).setForegroundColor(INK).build());
    if (sub) rt.setTextStyle(text.length + 3, t.length, SpreadsheetApp.newTextStyle().setFontSize(9).setForegroundColor(MUTED).build());
    sm.getRange(row, col, 1, 6).merge().setRichTextValue(rt.build()).setHorizontalAlignment('right').setVerticalAlignment('bottom');
  };
  var colHead = function (row, col, n) {
    sm.getRange(row, col, 1, n).setFontColor(MUTED).setFontSize(9).setFontWeight('bold')
      .setBorder(null, null, true, null, null, null, LINE, SpreadsheetApp.BorderStyle.SOLID);
  };
  var hairlines = function (row, col, rows, n) {
    sm.getRange(row, col, rows, n).setBorder(null, null, true, null, null, true, HAIR, SpreadsheetApp.BorderStyle.SOLID);
  };

  // Header card: logo, page name, time of the latest response
  box(2, 2, 1, 17);
  sm.getRange(2, 3, 1, 2).merge().setValue(SpreadsheetApp.newCellImage().setSourceUrl(LOGO_URL).setAltTextTitle('ESPRESSO SHOT').build());
  sm.getRange(2, 5, 1, 9).merge().setRichTextValue(SpreadsheetApp.newRichTextValue().setText('لوحة تقييمات الزبائن')
    .setTextStyle(SpreadsheetApp.newTextStyle().setBold(true).setFontSize(14).setForegroundColor(INK).build()).build())
    .setHorizontalAlignment('right');
  sm.getRange(2, 14, 1, 4).merge().setFormula('=IFERROR("آخر رد  " & TEXT(MAX(' + q + 'A3:A),"dd/mm  hh:mm"),"")')
    .setFontColor(MUTED).setFontSize(9).setHorizontalAlignment('left');

  // KPI cards
  [['عدد الردود', '=COUNTA(' + q + 'B3:B)', '0', 2],
   ['متوسط التجربة العامة', '=IFERROR(AVERAGE(' + q + 'F3:F),"—")', '0.0" / 5"', 6],
   ['طلبوا نتواصل معهم', '=COUNTIF(' + q + 'J3:J,"نعم")', '0', 11],
   ['تقييمات منخفضة (1–2)', '=COUNTIF(' + q + 'F3:F,"<=2")', '0', 15]
  ].forEach(function (t) {
    box(4, t[3], 2, 4);
    sm.getRange(4, t[3], 1, 4).merge().setValue(t[0]).setFontColor(MUTED).setFontSize(9).setVerticalAlignment('bottom');
    sm.getRange(5, t[3], 1, 4).merge().setFormula(t[1]).setNumberFormat(t[2]).setFontSize(22).setFontColor(INK);
    sm.getRange(4, t[3], 2, 4).setBorder(true, true, true, true, null, null, CANVAS, SpreadsheetApp.BorderStyle.SOLID_THICK);
  });

  // Card 1 (right, top): by branch
  box(7, 2, 7, 8);
  cardTitle(7, 3, 'حسب الفرع', 'متوسط من 5');
  sm.getRange(8, 3, 1, 6).setValues([['الفرع', 'الردود', 'الجودة', 'السرعة', 'النظافة', 'التجربة العامة']]);
  sm.getRange(8, 3).setHorizontalAlignment('right');
  colHead(8, 3, 6);
  BRANCHES.concat(['كل الفروع']).forEach(function (b, i) {
    var r = 9 + i, isTotal = i === BRANCHES.length;
    sm.getRange(r, 3).setValue(b).setFontWeight('bold').setHorizontalAlignment('right');
    sm.getRange(r, 4).setFormula(isTotal ? '=COUNTA(' + q + 'B3:B)' : '=COUNTIF(' + q + '$B$3:$B,$C' + r + ')').setFontColor(MUTED);
    ['C', 'D', 'E', 'F'].forEach(function (col, k) {
      var rng = q + col + '3:' + col;
      sm.getRange(r, 5 + k).setFormula(isTotal ? '=IFERROR(AVERAGE(' + rng + '),"—")'
        : '=IFERROR(AVERAGEIF(' + q + '$B$3:$B,$C' + r + ',' + rng + '),"—")');
    });
    if (isTotal) sm.getRange(r, 3, 1, 6).setFontWeight('bold').setBorder(true, null, null, null, null, null, LINE, SpreadsheetApp.BorderStyle.SOLID);
  });
  sm.getRange(9, 5, 4, 4).setNumberFormat('0.0');
  sm.getRange(9, 8, 4, 1).setFontWeight('bold').setFontSize(11);
  hairlines(9, 3, 2, 6);

  // Card 2 (left, top): spread of the overall rating as a donut chart
  box(7, 11, 7, 8);
  cardTitle(7, 12, 'توزيع التجربة العامة', 'كم زبون أعطى كل تقييم');
  var helper = sm.getRange(30, 3, 4, 2);   // chart data, kept out of sight below the dashboard
  helper.setValues([['ممتاز · 5', ''], ['جيد جداً · 4', ''], ['مقبول · 3', ''], ['ضعيف · 1–2', '']]);
  sm.getRange(30, 4).setFormula('=COUNTIF(' + q + 'F3:F,5)');
  sm.getRange(31, 4).setFormula('=COUNTIF(' + q + 'F3:F,4)');
  sm.getRange(32, 4).setFormula('=COUNTIF(' + q + 'F3:F,3)');
  sm.getRange(33, 4).setFormula('=COUNTIF(' + q + 'F3:F,"<=2")');
  helper.setFontColor(CANVAS);
  sm.insertChart(sm.newChart().setChartType(Charts.ChartType.PIE).addRange(helper)
    .setPosition(8, 12, 0, 0)
    .setOption('pieHole', 0.6)
    .setOption('colors', ['#262626', '#737373', '#c4c4c4', RED])
    .setOption('legend', { position: 'right' })
    .setOption('pieSliceText', 'none')
    .setOption('width', 600).setOption('height', 136)
    .build());

  // Card 3 (right, bottom): latest comments
  box(15, 2, 6, 8);
  cardTitle(15, 3, 'آخر الملاحظات', 'أحدث 3');
  sm.getRange(16, 3).setValue('التقييم');
  sm.getRange(16, 4, 1, 4).merge().setValue('الملاحظة').setHorizontalAlignment('right');
  sm.getRange(16, 8).setValue('الفرع · اليوم');
  colHead(16, 3, 6);
  var src = 'SORT(FILTER({' + q + 'A3:A,' + q + 'B3:B,' + q + 'F3:F,' + q + 'G3:G},' + q + 'G3:G<>""),1,FALSE)';
  var pick = function (n, k) { return 'INDEX(' + src + ',' + n + ',' + k + ')'; };
  for (var n = 1; n <= 3; n++) {
    var row = 16 + n;
    sm.getRange(row, 3).setFormula('=IFERROR(' + pick(n, 3) + ',"")').setNumberFormat('0').setFontWeight('bold').setFontSize(12);
    sm.getRange(row, 4, 1, 4).merge().setFormula('=IFERROR(' + pick(n, 4) + ',"")').setHorizontalAlignment('right').setWrap(false);
    sm.getRange(row, 8).setFormula('=IFERROR(' + pick(n, 2) + '&" · "&TEXT(' + pick(n, 1) + ',"dd/mm"),"")').setFontColor(MUTED).setFontSize(9);
  }
  hairlines(17, 3, 2, 6);

  // Card 4 (left, bottom): the app
  box(15, 11, 6, 8);
  cardTitle(15, 12, 'تطبيق ESPRESSO SHOT', 'من سؤال آخر الاستبيان');
  sm.getRange(16, 12, 1, 3).merge().setValue('المؤشر').setHorizontalAlignment('right');
  sm.getRange(16, 15, 1, 3).merge().setValue('القيمة').setHorizontalAlignment('left');
  colHead(16, 12, 6);
  [['جرّبوا التطبيق', '=IFERROR(COUNTIF(' + q + 'K3:K,"نعم")/COUNTA(' + q + 'K3:K),"—")', '0%'],
   ['تقييم التطبيق', '=IFERROR(AVERAGE(' + q + 'L3:L),"—")', '0.0" / 5"'],
   ['أكثر سبب لعدم التجربة', '=IFERROR(INDEX(QUERY({' + q + 'M3:M},"select Col1, count(Col1) where Col1 <> \'\' group by Col1 order by count(Col1) desc limit 1",0),2,1),"—")', '@']
  ].forEach(function (a, i) {
    var r = 17 + i;
    sm.getRange(r, 12, 1, 3).merge().setValue(a[0]).setHorizontalAlignment('right').setFontColor(MUTED);
    sm.getRange(r, 15, 1, 3).merge().setFormula(a[1]).setNumberFormat(a[2]).setHorizontalAlignment('left').setFontWeight('bold').setFontSize(12);
  });
  hairlines(17, 12, 2, 6);

  sm.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=AND(ISNUMBER(E9),E9<3)')
      .setFontColor(RED).setRanges([sm.getRange(9, 5, 4, 4)]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$O$5>0')
      .setFontColor(RED).setRanges([sm.getRange(5, 15, 1, 4)]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=AND(ISNUMBER($C17),$C17<=2)')
      .setFontColor(RED).setRanges([sm.getRange(17, 3, 3, 1)]).build()
  ]);
}
