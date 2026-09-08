/**
 * ICTD - نموذج تقييم مقابلة شخصية وفنية
 * Apps Script Web App Backend — v1.0
 *
 * الإعداد:
 * 1) الشيت جاهز بالفعل باسم "سجل تقييم المقابلات - ICTD":
 *    https://docs.google.com/spreadsheets/d/1T54mh8JjtVzvYK5m7LUgy1ZClFrRZ7vQBBKiy2JZQG4/edit
 *    افتحه، ثم من القائمة: الإضافات (Extensions) > Apps Script
 * 3) الصق هذا الكود كاملاً بدلاً من أي كود موجود، ثم احفظ (Ctrl+S)
 * 4) من زر Deploy > New deployment:
 *      - Type: Web app
 *      - Execute as: Me
 *      - Who has access: Anyone (أو Anyone with the link)
 * 5) انسخ رابط الـ Web App (ينتهي بـ /exec) والصقه في متغير SHEET_URL
 *    داخل ملف الـ HTML (أعلى وسم <script> في نهاية الصفحة)
 * 6) عند أي تعديل على الكود يجب عمل Deploy > Manage deployments > تحديث (Edit) > New version
 */

const SHEET_NAME = "التقييمات"; // اسم الشيت (Tab) الذي سيتم الكتابة عليه

const HEADERS = [
  "الرقم المرجعي",
  "تاريخ الإرسال",
  "اسم المرشح",
  "تاريخ المقابلة",
  "الوظيفة",
  "القائم بالمقابلة",
  "المستوى العام",
  "ملخص التقييم العام",
  "نص التوصية",
  "نقاط التوصية",
  "تفاصيل التقييم الفني (JSON)",
  "اسم رئيس اللجنة",
  "تاريخ توقيع رئيس اللجنة",
  "اسم الرئيس التنفيذي",
  "تاريخ اعتماد الرئيس التنفيذي"
];

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight("bold");
    sheet.setRightToLeft(true);
  }
  return sheet;
}

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const sheet = getSheet_();

    // منع التكرار: إذا كان الرقم المرجعي موجودًا مسبقًا، حدّث الصف بدلاً من الإضافة
    const formIdCol = 1;
    const lastRow = sheet.getLastRow();
    let targetRow = -1;
    if (lastRow > 1) {
      const ids = sheet.getRange(2, formIdCol, lastRow - 1, 1).getValues().flat();
      const idx = ids.indexOf(data.formId);
      if (idx !== -1) targetRow = idx + 2;
    }

    const rowValues = [
      data.formId || "",
      data.submittedAt || new Date().toISOString(),
      data.candidateName || "",
      data.interviewDate || "",
      data.position || "",
      data.interviewer || "",
      data.overallLevel || "",
      data.overallSummary || "",
      data.recommendationIntro || "",
      (data.recommendationPoints || []).join(" | "),
      JSON.stringify(data.evaluationRows || []),
      data.sign1Name || "",
      data.sign1Date || "",
      data.sign2Name || "",
      data.sign2Date || ""
    ];

    if (targetRow > 0) {
      sheet.getRange(targetRow, 1, 1, rowValues.length).setValues([rowValues]);
    } else {
      sheet.appendRow(rowValues);
    }

    return ContentService
      .createTextOutput(JSON.stringify({ result: "success", formId: data.formId }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ result: "error", message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({ status: "ICTD Interview Evaluation Form API — v1.0 running" }))
    .setMimeType(ContentService.MimeType.JSON);
}
