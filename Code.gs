var SHEET_ID = "1wn0xkxn6om9hIAsEoQYac6Gr9lSoN477wSL5weFHt5U";

var SUMMARY_HEADERS = ["الوقت", "الاسم", "العلامة", "من", "النسبة", "التصنيف", "نتاج أ", "نتاج ب", "نتاج ج"];
var DETAIL_HEADERS = ["الوقت", "الاسم", "رقم السؤال", "الرمز", "النتاج", "النتيجة", "اختيار الطالب"];

function ensureSheet(ss, name, headers) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  if (sheet.getLastRow() === 0 || String(sheet.getRange(1, 1).getValue()) !== headers[0]) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold");
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function doPost(e) {
  var data = {};
  try {
    data = JSON.parse(e.postData.contents);
  } catch (err) {
    data = {};
  }
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var when = new Date();
  var summary = ensureSheet(ss, "النتائج", SUMMARY_HEADERS);
  summary.appendRow([
    when,
    data.name || "",
    Number(data.correct) || 0,
    Number(data.total) || 0,
    Number(data.percent) || 0,
    data.band || "",
    data.outcomeA || "",
    data.outcomeB || "",
    data.outcomeC || ""
  ]);
  var items = data.items || [];
  if (items.length) {
    var detail = ensureSheet(ss, "التفاصيل", DETAIL_HEADERS);
    var rows = [];
    for (var i = 0; i < items.length; i++) {
      var item = items[i] || {};
      rows.push([
        when,
        data.name || "",
        Number(item.n) || 0,
        item.code || "",
        item.lo || "",
        Number(item.ok) ? "صحيح" : "خطأ",
        String(item.pick || "").slice(0, 180)
      ]);
    }
    detail.getRange(detail.getLastRow() + 1, 1, rows.length, DETAIL_HEADERS.length).setValues(rows);
  }
  return ContentService.createTextOutput("ok");
}

function doGet() {
  return ContentService.createTextOutput("ready");
}
