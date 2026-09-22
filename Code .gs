/**
 * ICTD - نموذج تقييم مقابلة شخصية وفنية
 * Apps Script Web App Backend — v1.2
 *
 * الإعداد:
 * 1) الشيت جاهز بالفعل باسم "سجل تقييم المقابلات - ICTD":
 *    https://docs.google.com/spreadsheets/d/1T54mh8JjtVzvYK5m7LUgy1ZClFrRZ7vQBBKiy2JZQG4/edit
 *    افتحه، ثم من القائمة: الإضافات (Extensions) > Apps Script
 * 3) الصق هذا الكود كاملاً بدلاً من أي كود موجود، ثم احفظ (Ctrl+S)
 * 4) من زر Deploy > New deployment (أو Manage deployments > Edit > New version إذا كان منشورًا من قبل):
 *      - Type: Web app
 *      - Execute as: Me
 *      - Who has access: Anyone (أو Anyone with the link)
 * 5) انسخ رابط الـ Web App (ينتهي بـ /exec) والصقه في متغير SHEET_URL
 *    داخل ملف الـ HTML (أعلى وسم <script> في نهاية الصفحة)
 * 6) عند أي تعديل على الكود يجب عمل Deploy > Manage deployments > تحديث (Edit) > New version
 *
 * جديد في v1.2:
 * - doGet يدعم الآن:
 *     ?action=list  → يعيد قائمة مختصرة بكل السجلات (للقائمة المنسدلة في الفورم)
 *     ?action=get&formId=XXX → يعيد تفاصيل سجل واحد كاملة (لتحميلها وتعديلها في الفورم)
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

function jsonOut_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function rowToRecord_(row) {
  let evalRows = [];
  try { evalRows = JSON.parse(row[10] || "[]"); } catch (e) { evalRows = []; }
  const recoPoints = (row[9] || "")
    .toString()
    .split(" | ")
    .map(s => s.trim())
    .filter(Boolean);

  return {
    formId: row[0] || "",
    submittedAt: row[1] || "",
    candidateName: row[2] || "",
    interviewDate: formatDate_(row[3]),
    position: row[4] || "",
    interviewer: row[5] || "",
    overallLevel: row[6] || "",
    overallSummary: row[7] || "",
    recommendationIntro: row[8] || "",
    recommendationPoints: recoPoints,
    evaluationRows: evalRows,
    sign1Name: row[11] || "",
    sign1Date: formatDate_(row[12]),
    sign2Name: row[13] || "",
    sign2Date: formatDate_(row[14])
  };
}

// يوحّد صيغة التاريخ إلى YYYY-MM-DD سواء كانت القيمة كائن تاريخ أو نصًا
function formatDate_(val) {
  if (!val) return "";
  if (Object.prototype.toString.call(val) === "[object Date]") {
    return Utilities.formatDate(val, Session.getScriptTimeZone(), "yyyy-MM-dd");
  }
  const s = val.toString();
  // إذا كانت القيمة تحتوي على توقيت ISO كامل، خذ الجزء الخاص بالتاريخ فقط عند الحاجة
  return s;
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

    return jsonOut_({ result: "success", formId: data.formId });

  } catch (err) {
    return jsonOut_({ result: "error", message: err.message });
  }
}

function doGet(e) {
  try {
    const action = (e.parameter && e.parameter.action) || "status";
    const sheet = getSheet_();
    const lastRow = sheet.getLastRow();

    if (action === "list") {
      if (lastRow <= 1) return jsonOut_({ result: "success", records: [] });
      const values = sheet.getRange(2, 1, lastRow - 1, HEADERS.length).getValues();
      const records = values
        .filter(r => r[0]) // له رقم مرجعي
        .map(r => ({
          formId: r[0],
          candidateName: r[2] || "",
          interviewDate: formatDate_(r[3]),
          position: r[4] || "",
          submittedAt: r[1] || ""
        }))
        // الأحدث أولاً
        .sort((a, b) => (a.submittedAt < b.submittedAt ? 1 : -1));
      return jsonOut_({ result: "success", records: records });
    }

    if (action === "get") {
      const formId = e.parameter.formId;
      if (!formId) return jsonOut_({ result: "error", message: "formId مطلوب" });
      if (lastRow <= 1) return jsonOut_({ result: "error", message: "لا توجد سجلات" });
      const values = sheet.getRange(2, 1, lastRow - 1, HEADERS.length).getValues();
      const row = values.find(r => r[0] === formId);
      if (!row) return jsonOut_({ result: "error", message: "لم يتم العثور على السجل" });
      return jsonOut_({ result: "success", record: rowToRecord_(row) });
    }

    return jsonOut_({ status: "ICTD Interview Evaluation Form API — v1.2 running" });

  } catch (err) {
    return jsonOut_({ result: "error", message: err.message });
  }
}
