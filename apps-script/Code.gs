/**
 * =====================================================================
 *  نموذج تقييم مقابلة شخصية وفنية — ICTD
 *  Google Apps Script Web App backend (v1.6)
 *  يضيف هذا الإصدار: SRU Secure Access Kit — تقييد الدخول على حسابات
 *  جامعة سليمان الراجحي (@sr.edu.sa) فقط، مع سجل دخول وقائمة مصرَّح لهم.
 * ---------------------------------------------------------------------
 *  الملفات: Code.gs · Index.html · Logos.html · Denied.html
 *  النشر: Deploy ▸ New deployment ▸ Web app
 *         Execute as: Me  ·  Who has access: Anyone within Sulaiman Al Rajhi University
 * =====================================================================
 */

/* ▼▼▼ إعدادات المنصّة ▼▼▼ */
const APP_NAME       = 'نموذج تقييم مقابلة شخصية وفنية';
const APP_NAME_EN    = 'Interview Evaluation Form';
const APP_VERSION    = '1.6';
const ALLOWED_DOMAIN = 'sr.edu.sa';
const FAVICON_URL    = 'https://ictsru.github.io/ITOC/assets/itoc-logo.png';

/* سجل الدخول (Access log) */
const LOG_ENABLED        = true;
const LOG_SPREADSHEET_ID = '';            // '' = نفس الشيت المرتبط بهذا السكربت (سجل تقييم المقابلات - ICTD)
const SHEET_LOG          = 'AccessLog';
const LOG_TIMEZONE       = 'Asia/Riyadh';
const LOG_DEDUP_MINUTES  = 30;            // نفس المستخدم خلال N دقيقة = سطر واحد (0 = كل فتح صفحة)
const LOG_RETENTION_DAYS = 180;

/* قائمة المصرَّح لهم (Access list) */
const ACCESS_MODE           = 'LIST';     // 'LIST' = فقط من في AllowedUsers (نشِط) · 'DOMAIN' = أي حساب @sr.edu.sa
const ACCESS_SPREADSHEET_ID = '';         // '' = نفس الشيت
const SHEET_ACCESS          = 'AllowedUsers';
const OWNER_EMAILS          = ['m.elmahdy@sr.edu.sa']; // دائمًا مصرَّح — لا يمكن استبعاد المالك

/* القائمة الافتراضية — تُضاف عند تشغيل setupAccessList() (الصفوف الموجودة تبقى كما هي) */
const DEFAULT_ALLOWED = [
  // [email, name, role, notes]
  ['nabounar@sr.edu.sa',    'د. ناصر أبو نار', 'admin', 'CIO — اعتماد الرئيس التنفيذي'],
  ['m.elmahdy@sr.edu.sa',   'Mohamed ElMahdy', 'admin', 'Owner — IT Operations Manager'],
  ['n.nawaz@sr.edu.sa',     'Muhammad Nasir Nawaz', 'user', 'القائم بالمقابلات'],
  ['m.omar@sr.edu.sa',      '', 'user', ''],
  ['m.saleh@sr.edu.sa',     '', 'user', ''],
  ['a.alrowedan@sr.edu.sa', '', 'user', ''],
  ['n.alrashidi@sr.edu.sa', '', 'user', ''],
  ['a.alawaji@sr.edu.sa',   '', 'user', ''],
  ['a.alaqeel@sr.edu.sa',   '', 'user', ''],
  ['a.alolayan@sr.edu.sa',  '', 'user', ''],
  ['s.albusayli@sr.edu.sa', '', 'user', ''],
  ['i.almutairi@sr.edu.sa', '', 'user', ''],
  ['a.elnadi@sr.edu.sa',    '', 'user', '']
];
const ACCESS_CACHE_SECONDS = 120;
/* ▲▲▲ إعدادات المنصّة ▲▲▲ */

const LOG_HEADERS = ['timestamp', 'date', 'time', 'day', 'email', 'result', 'version'];
const DAYS_AR_    = ['الأحد', 'الاثنين', 'الثلاثاء', 'الأربعاء', 'الخميس', 'الجمعة', 'السبت'];

/* ----------------------------- بيانات التقييمات (كما في v1.5) ----------------------------- */

const SHEET_NAME = 'التقييمات';
const HEADERS = [
  'الرقم المرجعي', 'تاريخ الإرسال', 'اسم المرشح', 'تاريخ المقابلة', 'الوظيفة',
  'القائم بالمقابلة', 'المستوى العام', 'ملخص التقييم العام', 'نص التوصية', 'نقاط التوصية',
  'تفاصيل التقييم الفني (JSON)', 'اسم رئيس اللجنة', 'تاريخ توقيع رئيس اللجنة',
  'اسم الرئيس التنفيذي', 'تاريخ اعتماد الرئيس التنفيذي'
];

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
  }
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS])
      .setFontWeight('bold').setFontColor('#ffffff').setBackground('#501e8c');
    sh.setFrozenRows(1);
    sh.setRightToLeft(true);
  }
  return sh;
}

function rowToRecord_(row) {
  let evaluationRows = [];
  try { evaluationRows = row[10] ? JSON.parse(row[10]) : []; } catch (e) { evaluationRows = []; }
  const recommendationPoints = row[9] ? String(row[9]).split('|').map(s => s.trim()).filter(Boolean) : [];
  return {
    formId: row[0],
    submittedAt: formatDate_(row[1]),
    candidateName: row[2],
    interviewDate: formatDate_(row[3]),
    position: row[4],
    interviewer: row[5],
    overallLevel: row[6],
    overallSummary: row[7],
    recommendationIntro: row[8],
    recommendationPoints: recommendationPoints,
    evaluationRows: evaluationRows,
    sign1Name: row[11],
    sign1Date: formatDate_(row[12]),
    sign2Name: row[13],
    sign2Date: formatDate_(row[14])
  };
}

function formatDate_(val) {
  if (!val) return '';
  if (val instanceof Date) return Utilities.formatDate(val, LOG_TIMEZONE, 'yyyy-MM-dd');
  return String(val);
}

/* ----------------------------- Web app: doGet / doPost ----------------------------- */

function doGet(e) {
  const action = e && e.parameter && e.parameter.action;

  // واجهات JSON (تُستدعى عبر fetch من داخل الصفحة نفسها بعد تسجيل الدخول)
  if (action === 'list' || action === 'get') {
    return handleApiGet_(e, action);
  }

  // عرض الصفحة نفسها — يخضع للتحقق من حساب الجامعة/قائمة المصرَّح لهم
  const email = currentEmail_();
  const status = accessStatus_(email);
  if (status !== 'ok') {
    logAccess_(email || '(unknown)', status === 'list' ? 'not-listed' : 'denied');
    const d = HtmlService.createTemplateFromFile('Denied');
    d.appName = APP_NAME; d.appNameEn = APP_NAME_EN;
    d.userEmail = email || '';
    d.reason = status;
    return d.evaluate().setTitle(APP_NAME + ' | غير مصرح')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }
  logAccess_(email, 'granted');

  const t = HtmlService.createTemplateFromFile('Index');
  t.appName   = APP_NAME;
  t.appNameEn = APP_NAME_EN;
  t.version   = APP_VERSION;
  t.userEmail = email;
  const out = t.evaluate().setTitle(APP_NAME_EN + ' | ' + APP_NAME)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  if (FAVICON_URL) out.setFaviconUrl(FAVICON_URL);
  return out;
}

/** يُنفَّذ فقط بعد التأكد من هوية المستخدم عبر requireSruUser_() */
function handleApiGet_(e, action) {
  requireSruUser_(); // يرمي خطأ إن لم يكن مصرَّحًا — سيظهر كـ JSON خطأ أدناه عبر try/catch
  try {
    const sh = getSheet_();
    const last = sh.getLastRow();
    if (action === 'list') {
      const records = [];
      if (last > 1) {
        const rows = sh.getRange(2, 1, last - 1, HEADERS.length).getValues();
        rows.forEach(r => {
          if (!r[0]) return;
          records.push({ formId: r[0], candidateName: r[2], interviewDate: formatDate_(r[3]), position: r[4], submittedAt: formatDate_(r[1]) });
        });
      }
      records.sort((a, b) => String(b.submittedAt).localeCompare(String(a.submittedAt)));
      return jsonOut_({ result: 'success', records: records });
    }
    if (action === 'get') {
      const formId = e.parameter.formId;
      if (!formId) return jsonOut_({ result: 'error', message: 'formId مطلوب' });
      if (last > 1) {
        const rows = sh.getRange(2, 1, last - 1, HEADERS.length).getValues();
        for (let i = 0; i < rows.length; i++) {
          if (String(rows[i][0]) === String(formId)) {
            return jsonOut_({ result: 'success', record: rowToRecord_(rows[i]) });
          }
        }
      }
      return jsonOut_({ result: 'error', message: 'السجل غير موجود' });
    }
  } catch (err) {
    return jsonOut_({ result: 'error', message: String(err) });
  }
}

function doPost(e) {
  try {
    requireSruUser_(); // يرمي خطأ إن لم يكن الحساب مصرَّحًا بالدخول
    const data = JSON.parse(e.postData.contents);
    const sh = getSheet_();
    const last = sh.getLastRow();
    const recommendationPoints = Array.isArray(data.recommendationPoints) ? data.recommendationPoints.join(' | ') : (data.recommendationPoints || '');
    const evaluationRowsJson = JSON.stringify(data.evaluationRows || []);
    const now = new Date();
    const rowValues = [
      data.formId, now, data.candidateName, data.interviewDate, data.position,
      data.interviewer, data.overallLevel, data.overallSummary, data.recommendationIntro,
      recommendationPoints, evaluationRowsJson, data.sign1Name, data.sign1Date,
      data.sign2Name, data.sign2Date
    ];

    let updated = false;
    if (last > 1) {
      const ids = sh.getRange(2, 1, last - 1, 1).getValues();
      for (let i = 0; i < ids.length; i++) {
        if (String(ids[i][0]) === String(data.formId)) {
          sh.getRange(i + 2, 1, 1, HEADERS.length).setValues([rowValues]);
          updated = true;
          break;
        }
      }
    }
    if (!updated) {
      sh.appendRow(rowValues);
    }
    return jsonOut_({ result: 'success', formId: data.formId, updated: updated });
  } catch (err) {
    return jsonOut_({ result: 'error', message: String(err) });
  }
}

function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/* ------------------------------ الوصول (Access) ------------------------------ */

function currentEmail_() {
  return String(Session.getActiveUser().getEmail() || '').toLowerCase();
}
function isAllowed_(email) {
  return !!email && email.endsWith('@' + ALLOWED_DOMAIN);
}
/** 'ok' | 'domain' (ليس @sr.edu.sa) | 'list' (ليس في القائمة النشِطة) */
function accessStatus_(email) {
  if (!isAllowed_(email)) return 'domain';
  if (ACCESS_MODE !== 'LIST') return 'ok';
  if (OWNER_EMAILS.map(x => x.toLowerCase()).indexOf(email) !== -1) return 'ok';
  return allowedEmails_().indexOf(email) !== -1 ? 'ok' : 'list';
}
function requireSruUser_() {
  const email = currentEmail_();
  const status = accessStatus_(email);
  if (status === 'domain') throw new Error('غير مصرح — الدخول لمنسوبي الجامعة فقط');
  if (status === 'list') throw new Error('غير مصرح — حسابك غير مضاف إلى قائمة المصرّح لهم');
  return email;
}

/* ---------------------------- قائمة المصرَّح لهم (AllowedUsers) --------------------------- */

const ACCESS_HEADERS = ['email', 'name', 'role', 'active', 'notes'];

function allowedEmails_() {
  const cache = CacheService.getScriptCache();
  const hit = cache.get('acl_v1');
  if (hit) return JSON.parse(hit);
  let list = [];
  try {
    const sh = accessSheet_(false);
    if (sh && sh.getLastRow() > 1) {
      list = sh.getRange(2, 1, sh.getLastRow() - 1, 4).getValues()
        .filter(r => r[0] && r[3] !== false && String(r[3]).toUpperCase() !== 'FALSE')
        .map(r => String(r[0]).trim().toLowerCase());
    }
  } catch (e) { console.error('allowedEmails_ failed: ' + e); }
  cache.put('acl_v1', JSON.stringify(list), ACCESS_CACHE_SECONDS);
  return list;
}

function accessSheet_(create) {
  const ss = ACCESS_SPREADSHEET_ID ? SpreadsheetApp.openById(ACCESS_SPREADSHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_ACCESS);
  if (!sh && create) {
    sh = ss.insertSheet(SHEET_ACCESS);
    sh.getRange(1, 1, 1, ACCESS_HEADERS.length).setValues([ACCESS_HEADERS])
      .setFontWeight('bold').setFontColor('#ffffff').setBackground('#501e8c');
    sh.setFrozenRows(1);
    sh.setRightToLeft(true);
    [260, 200, 100, 70, 260].forEach((w, i) => sh.setColumnWidth(i + 1, w));
    sh.getRange(2, 4, 300, 1).insertCheckboxes();
    sh.getRange('A1').setNote('البريد الجامعي كاملًا — name@sr.edu.sa');
    sh.getRange('C1').setNote('admin / user (للتوثيق)');
    sh.getRange('D1').setNote('ألغِ العلامة لإيقاف الوصول دون حذف الصف');
  }
  return sh;
}

/** شغّلها مرة واحدة من محرر Apps Script: تُنشئ تبويب AllowedUsers وتضيف المالك والقائمة الافتراضية. */
function setupAccessList() {
  const sh = accessSheet_(true);
  const n = Math.max(sh.getLastRow() - 1, 0);
  const data = n ? sh.getRange(2, 1, n, 5).getValues() : [];

  const kept = data.filter(r => String(r[0]).trim())
    .map(r => [String(r[0]).trim(), r[1], r[2], r[3] === true || String(r[3]).toUpperCase() === 'TRUE', r[4]]);
  const have = kept.map(r => r[0].toLowerCase());

  const seed = OWNER_EMAILS.map(e => [e, 'Owner', 'admin', 'owner'])
    .concat(typeof DEFAULT_ALLOWED !== 'undefined' ? DEFAULT_ALLOWED : []);
  let added = 0;
  seed.forEach(r => {
    const e = String(r[0]).trim().toLowerCase();
    if (e && have.indexOf(e) === -1) {
      kept.push([e, r[1] || '', r[2] || 'user', true, r[3] || 'added by setupAccessList']);
      have.push(e); added++;
    }
  });

  if (n) sh.getRange(2, 1, n, 5).clearContent();
  const rows = Math.max(n, kept.length, 50);
  sh.getRange(2, 4, rows, 1).insertCheckboxes();
  if (kept.length) sh.getRange(2, 1, kept.length, 5).setValues(kept);

  CacheService.getScriptCache().remove('acl_v1');
  console.log('AllowedUsers ready: ' + kept.length + ' entries (' + added + ' added).');
  try { SpreadsheetApp.getActive().toast('قائمة المصرّح لهم جاهزة: ' + kept.length + ' بريدًا (أُضيف ' + added + ').', SHEET_ACCESS, 8); } catch (e) {}
}

/* ---------------------------- سجل الدخول (AccessLog) ---------------------------- */

function logAccess_(email, result) {
  if (!LOG_ENABLED) return;
  try {
    if (LOG_DEDUP_MINUTES > 0 && result === 'granted') {
      const cache = CacheService.getScriptCache();
      const key = 'log_' + email;
      if (cache.get(key)) return;
      cache.put(key, '1', LOG_DEDUP_MINUTES * 60);
    }
    const now = new Date();
    const f = p => Utilities.formatDate(now, LOG_TIMEZONE, p);
    getLogSheet_().appendRow([now, f('yyyy-MM-dd'), f('HH:mm:ss'), DAYS_AR_[Number(f('u')) % 7],
      email, result === 'granted' ? 'دخول' : (result === 'not-listed' ? 'غير مدرج' : 'مرفوض'), APP_VERSION]);
  } catch (e) {
    console.error('logAccess_ failed: ' + e);
  }
}

function getLogSheet_() {
  const ss = LOG_SPREADSHEET_ID ? SpreadsheetApp.openById(LOG_SPREADSHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('No spreadsheet: bind the script to a Sheet or set LOG_SPREADSHEET_ID');
  let sh = ss.getSheetByName(SHEET_LOG);
  if (!sh) {
    sh = ss.insertSheet(SHEET_LOG);
    sh.getRange(1, 1, 1, LOG_HEADERS.length).setValues([LOG_HEADERS])
      .setFontWeight('bold').setFontColor('#ffffff').setBackground('#3a1464');
    sh.setFrozenRows(1);
    sh.setRightToLeft(true);
    sh.getRange('A:A').setNumberFormat('yyyy-mm-dd hh:mm:ss');
    [170, 100, 80, 80, 260, 80, 70].forEach((w, i) => sh.setColumnWidth(i + 1, w));
  }
  return sh;
}

function showLogSummary() {
  const sh = getLogSheet_();
  const ui = SpreadsheetApp.getUi();
  const last = sh.getLastRow();
  if (last < 2) { ui.alert('سجل الدخول', 'لا توجد سجلات بعد.', ui.ButtonSet.OK); return; }
  const since = new Date(Date.now() - 30 * 864e5);
  const rows = sh.getRange(2, 1, last - 1, LOG_HEADERS.length).getValues()
    .filter(r => r[0] instanceof Date && r[0] >= since && r[5] === 'دخول');
  const users = {};
  rows.forEach(r => { users[r[4]] = (users[r[4]] || 0) + 1; });
  const top = Object.keys(users).sort((a, b) => users[b] - users[a]).slice(0, 10)
    .map((u, i) => (i + 1) + '. ' + u + ' — ' + users[u]).join('\n');
  ui.alert('سجل الدخول — آخر 30 يومًا — ' + APP_NAME,
    'عدد الزيارات: ' + rows.length + '\nعدد المستخدمين: ' + Object.keys(users).length + '\n\nالأكثر دخولًا:\n' + (top || '—'),
    ui.ButtonSet.OK);
}

function purgeOldLogs() {
  const sh = getLogSheet_();
  const last = sh.getLastRow();
  let removed = 0;
  if (last > 1) {
    const cutoff = new Date(Date.now() - LOG_RETENTION_DAYS * 864e5);
    const stamps = sh.getRange(2, 1, last - 1, 1).getValues();
    for (let i = stamps.length - 1; i >= 0; i--) {
      if (stamps[i][0] instanceof Date && stamps[i][0] < cutoff) { sh.deleteRow(i + 2); removed++; }
    }
  }
  SpreadsheetApp.getActive().toast('حُذف ' + removed + ' سجل أقدم من ' + LOG_RETENTION_DAYS + ' يومًا.', APP_NAME_EN, 6);
}

function onOpen() {
  try {
    SpreadsheetApp.getUi().createMenu(APP_NAME_EN)
      .addItem('إعداد قائمة المصرّح لهم (AllowedUsers)', 'setupAccessList')
      .addSeparator()
      .addItem('ملخص سجل الدخول (30 يومًا)', 'showLogSummary')
      .addItem('حذف سجلات الدخول القديمة', 'purgeOldLogs')
      .addSeparator()
      .addItem('عرض رابط المنصة', 'showWebAppUrl')
      .addToUi();
  } catch (e) { /* standalone script */ }
}

function showWebAppUrl() {
  const ui = SpreadsheetApp.getUi();
  ui.alert('رابط ' + APP_NAME, ScriptApp.getService().getUrl() || 'لم يُنشر بعد: Deploy ▸ New deployment ▸ Web app', ui.ButtonSet.OK);
}

/* ----------------------------- Helpers ------------------------------ */

function include(name) {
  return HtmlService.createHtmlOutputFromFile(name).getContent();
}
