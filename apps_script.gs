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

// Run once from the editor to style the sheet and build the summary tab.
function formatSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheets()[0];
  var DATA = 'الردود', SUM = 'ملخص', N = HEADERS.length;
  sh.setName(DATA);
  sh.setRightToLeft(true);
  if (sh.getLastRow() === 0) sh.appendRow(HEADERS);
  if (sh.getMaxColumns() > N) sh.deleteColumns(N + 1, sh.getMaxColumns() - N);
  var rows = sh.getMaxRows();
  sh.getBandings().forEach(function (b) { b.remove(); });
  if (sh.getFilter()) sh.getFilter().remove();

  sh.getRange(1, 1, rows, N).setFontFamily('Readex Pro').setFontSize(10).setFontColor('#262626')
    .setVerticalAlignment('middle').setHorizontalAlignment('center').setWrap(false);
  sh.getRange(1, 1, 1, N).setBackground('#3d3d3d').setFontColor('#ffffff').setFontWeight('bold').setFontSize(11);
  sh.setFrozenRows(1);
  sh.setRowHeight(1, 42);
  [150, 120, 135, 140, 120, 150, 380, 150, 140, 160].forEach(function (w, i) { sh.setColumnWidth(i + 1, w); });

  var body = rows - 1;
  sh.getRange(2, 1, body, 1).setNumberFormat('yyyy-mm-dd  hh:mm').setFontColor('#737373');
  sh.getRange(2, 3, body, 4).setFontWeight('bold').setFontSize(11);
  sh.getRange(2, 7, body, 1).setHorizontalAlignment('right').setWrap(true);
  sh.getRange(2, 8, body, 1).setHorizontalAlignment('right');
  sh.getRange(2, 9, body, 1).setNumberFormat('@');
  sh.getRange(2, 1, body, N).applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, false, false)
    .setFirstRowColor('#ffffff').setSecondRowColor('#f5f5f5');
  sh.setHiddenGridlines(true);

  var ratings = sh.getRange(2, 3, body, 4), contact = sh.getRange(2, 8, body, 3);
  sh.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=AND(ISNUMBER(C2),C2<=2)')
      .setBackground('#fbeceb').setFontColor('#b3261e').setRanges([ratings]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=AND(ISNUMBER(C2),C2=3)')
      .setBackground('#fff6dd').setFontColor('#7a5a00').setRanges([ratings]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$J2="نعم"')
      .setBackground('#fff6dd').setBold(true).setRanges([contact]).build()
  ]);
  sh.getRange(1, 1, rows, N).createFilter();

  var sm = ss.getSheetByName(SUM) || ss.insertSheet(SUM, 1);
  sm.clear();
  sm.setRightToLeft(true);
  sm.setHiddenGridlines(true);
  if (sm.getMaxColumns() > 6) sm.deleteColumns(7, sm.getMaxColumns() - 6);
  var q = "'" + DATA + "'!";
  var branches = ['الماصيون', 'الطيرة', 'وسط البلد'];
  var table = [['الفرع', 'عدد الردود', 'جودة الطلب', 'سرعة الخدمة', 'النظافة', 'التجربة العامة']];
  branches.forEach(function (b, i) {
    var r = i + 4, row = [b, '=COUNTIF(' + q + 'B:B,A' + r + ')'];
    ['C', 'D', 'E', 'F'].forEach(function (c) {
      row.push('=IFERROR(AVERAGEIF(' + q + '$B:$B,$A' + r + ',' + q + c + ':' + c + '),"—")');
    });
    table.push(row);
  });
  var total = ['كل الفروع', '=COUNTA(' + q + 'B2:B)'];
  ['C', 'D', 'E', 'F'].forEach(function (c) { total.push('=IFERROR(AVERAGE(' + q + c + '2:' + c + '),"—")'); });
  table.push(total);

  sm.getRange('A1').setValue('ملخص التقييمات').setFontSize(18).setFontWeight('normal');
  sm.getRange('A2').setValue('المتوسط من 5. يتحدّث تلقائياً مع كل رد جديد.').setFontColor('#737373');
  sm.getRange(3, 1, table.length, 6).setValues(table);
  sm.getRange('A9').setValue('طلبوا التواصل معهم');
  sm.getRange('B9').setFormula('=COUNTIF(' + q + 'J:J,"نعم")');
  sm.getRange('A10').setValue('تقييم عام 2 أو أقل');
  sm.getRange('B10').setFormula('=COUNTIF(' + q + 'F2:F,"<=2")');

  sm.getRange(1, 1, sm.getMaxRows(), 6).setFontFamily('Readex Pro').setFontColor('#262626').setVerticalAlignment('middle');
  sm.getRange('A3:F3').setBackground('#3d3d3d').setFontColor('#ffffff').setFontWeight('bold').setHorizontalAlignment('center');
  sm.getRange('A4:F7').setHorizontalAlignment('center').setFontSize(11)
    .setBorder(null, null, true, null, null, true, '#e0e0e0', SpreadsheetApp.BorderStyle.SOLID);
  sm.getRange('A4:A7').setHorizontalAlignment('right').setFontWeight('bold');
  sm.getRange('A7:F7').setBackground('#ececec').setFontWeight('bold');
  sm.getRange('C4:F7').setNumberFormat('0.0');
  sm.getRange('A9:A10').setFontWeight('bold');
  sm.getRange('B9:B10').setHorizontalAlignment('center').setFontSize(12).setFontWeight('bold');
  sm.setRowHeight(1, 44); sm.setRowHeight(3, 38);
  [170, 110, 110, 110, 110, 120].forEach(function (w, i) { sm.setColumnWidth(i + 1, w); });
  sm.setFrozenRows(0);
  ss.setActiveSheet(sh);
}
