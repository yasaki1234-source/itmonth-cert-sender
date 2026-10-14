/**
 * 大專碩博 AI × 學習產業見習證書｜批次發送工具（Google Apps Script 網頁版）
 * v1.2 · 2026-10
 *
 * 架構：網頁介面放在 GitHub Pages（docs/index.html），本程式只當資料 API（doPost）。
 *       網頁呼叫時不帶 Google 登入資訊，瀏覽器登入幾個 Google 帳號都不影響。
 * 安裝：請依《安裝與操作手冊》。第一次請在編輯器上方選擇函式「setup」→ 按「執行」。
 * 本檔不需修改；所有設定都在網頁的「設定」頁完成。
 */

const APP_TITLE = '見習證書發送工具';
const LIST_SHEET = '核發名單';
const LOG_SHEET = '寄送紀錄';
const TZ = 'Asia/Taipei';
const BATCH_SECONDS = 25;          // 每次伺服器呼叫最多處理秒數（網頁會自動接續）
const SENT_STATUSES = ['已寄出', '已補寄'];
const PLACEHOLDERS = ['{{姓名}}', '{{年度}}', '{{時數}}', '{{發證年月}}', '{{證書編號}}', '{{學校}}'];

const DEFAULT_CONFIG = {
  year: '2026',
  hours: '5',
  issueMonth: '2026年11月',
  certPrefix: 'ITM2026-',
  senderName: '資訊月活動委員會',
  replyTo: 'service@itmonth.tw',
  fromAlias: '',          // 選填：以別名寄出（需先在 Gmail 設定「以這個地址寄送郵件」）
  subject: '【{{年度}}資訊月】大專碩博AI×學習產業見習證書｜{{姓名}}',
  body: [
    '{{姓名}} 同學您好：',
    '',
    '感謝您參與「{{年度}}資訊月暨臺灣教育科技展」AI × 學習產業見習活動。',
    '隨信附上您的見習證書（證書編號：{{證書編號}}），請下載留存。',
    '',
    '本信由系統寄出，如有疑問請直接回覆本信。',
    '',
    '台北市電腦商業同業公會',
    '資訊月活動委員會'
  ].join('\n'),
  requireCheckin: true,   // 需要紙本簽到核對
  requireOut: true,       // 需要簽退
  requireSurvey: true,    // 需要填問卷
  identityKeywords: '大學,大專,碩士,博士,研究所,研究生,學士,科大,技術學院'
};

// 試算表欄位：[程式用名稱, 試算表標題]
const FIELDS = [
  ['seq', '編號'], ['certNo', '證書編號'], ['name', '姓名'], ['email', 'Email'],
  ['school', '學校'], ['phone', '手機'], ['source', '來源'],
  ['checkIn', '簽到'], ['checkOut', '簽退'], ['survey', '問卷'],
  ['result', '比對結果'], ['note', '備註'], ['approved', '核發'],
  ['status', '寄送狀態'], ['sentAt', '寄送時間'], ['error', '錯誤訊息'],
  ['id', '_id'], ['pdfId', '_pdfId'], ['inFile', '_inFile'], ['outFile', '_outFile'],
  ['surveyFile', '_surveyFile'], ['inManual', '_inManual'], ['outManual', '_outManual'],
  ['approve', '_approve'], ['flags', '_flags'], ['edited', '_edited']
];

/* =========================================================
 * 1. 網頁入口與初始化
 * ========================================================= */

/** 舊網址：顯示「已搬到新網址」說明頁 */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle(APP_TITLE)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** 網頁（GitHub Pages）呼叫的 API 入口：POST {fn, args} → {ok, data} 或 {ok:false, error} */
function doPost(e) {
  let out;
  try {
    const req = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const fn = apiFunctions_()[req.fn];
    if (!fn) throw new Error('不支援的操作：' + req.fn);
    out = { ok: true, data: fn.apply(null, Array.isArray(req.args) ? req.args : []) };
  } catch (err) {
    out = { ok: false, error: (err && err.message) || String(err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

/** 允許從網頁呼叫的函式（白名單） */
function apiFunctions_() {
  return {
    apiLogin: apiLogin, apiLogout: apiLogout, apiChangeLogin: apiChangeLogin,
    apiBootstrap: apiBootstrap, apiBuildList: apiBuildList, apiUpdateRow: apiUpdateRow,
    apiAddRow: apiAddRow, apiRemoveRow: apiRemoveRow, apiPreview: apiPreview, apiSendTest: apiSendTest,
    apiGenerateBatch: apiGenerateBatch, apiResetPdfs: apiResetPdfs, apiGetPdf: apiGetPdf,
    apiSendBatch: apiSendBatch, apiResend: apiResend, apiRequeueFailed: apiRequeueFailed,
    apiLogs: apiLogs, apiSaveConfig: apiSaveConfig, apiUploadTemplate: apiUploadTemplate,
    apiSetTemplateUrl: apiSetTemplateUrl
  };
}

/** 第一次安裝時在編輯器執行一次。重複執行不會覆蓋既有資料。 */
function setup() {
  const p = props_();
  let root;
  if (p.getProperty('FOLDER_ID')) {
    root = DriveApp.getFolderById(p.getProperty('FOLDER_ID'));
  } else {
    root = DriveApp.createFolder('見習證書發送工具');
    p.setProperty('FOLDER_ID', root.getId());
  }
  ensureSubFolder_('PDF_FOLDER_ID', '證書PDF', root);
  ensureSubFolder_('PREVIEW_FOLDER_ID', '預覽與測試', root);
  ensureSubFolder_('TEMP_FOLDER_ID', '暫存', root);

  if (!p.getProperty('SS_ID')) {
    const ss = SpreadsheetApp.create('見習證書發送工具_資料');
    DriveApp.getFileById(ss.getId()).moveTo(root);
    p.setProperty('SS_ID', ss.getId());
  }
  ensureSheets_();
  if (!p.getProperty('CONFIG')) p.setProperty('CONFIG', JSON.stringify(DEFAULT_CONFIG));
  initLogin_();

  try {
    Drive.Files.list({ pageSize: 1 });
  } catch (e) {
    throw new Error('Drive 進階服務未啟用：請確認 appsscript.json 已依手冊貼上。原始訊息：' + e.message);
  }
  // 觸發寄信與簡報權限授權（第一次執行時會跳出授權視窗）
  MailApp.getRemainingDailyQuota();

  Logger.log('✅ 初始化完成');
  Logger.log('登入帳號：' + p.getProperty('LOGIN_USER') + (p.getProperty('LOGIN_DEFAULT') === 'Y' ? '（預設密碼 itmonth，登入後請到「設定」修改）' : ''));
  Logger.log('資料夾：' + root.getUrl());
}

function ensureSubFolder_(key, name, root) {
  const p = props_();
  if (p.getProperty(key)) {
    try { DriveApp.getFolderById(p.getProperty(key)); return; } catch (e) { /* 重建 */ }
  }
  p.setProperty(key, root.createFolder(name).getId());
}

function ensureSheets_() {
  const ss = ss_();
  let list = ss.getSheetByName(LIST_SHEET);
  if (!list) {
    list = ss.insertSheet(LIST_SHEET);
    list.getRange(1, 1, 1, FIELDS.length).setValues([FIELDS.map(f => f[1])])
      .setFontWeight('bold').setBackground('#1F3A66').setFontColor('#FFFFFF');
    list.setFrozenRows(1);
    list.getRange(1, 1, list.getMaxRows(), FIELDS.length).setNumberFormat('@');
  }
  let log = ss.getSheetByName(LOG_SHEET);
  if (!log) {
    log = ss.insertSheet(LOG_SHEET);
    log.getRange(1, 1, 1, 7).setValues([['時間', '動作', '證書編號', '姓名', 'Email', '結果', '說明']])
      .setFontWeight('bold').setBackground('#1F3A66').setFontColor('#FFFFFF');
    log.setFrozenRows(1);
  }
  ['工作表1', 'Sheet1'].forEach(n => {
    const s = ss.getSheetByName(n);
    if (s && ss.getSheets().length > 1) ss.deleteSheet(s);
  });
}

/* =========================================================
 * 2. 網頁呼叫的 API（除登入外，每個都先檢查登入權杖）
 * ========================================================= */

function apiBootstrap(token) {
  guard_(token);
  const rows = loadRows_();
  return payload_(rows, true);
}

function apiBuildList(token, payload) {
  guard_(token);
  return withLock_(() => {
    const cfg = getConfig_();
    const existing = loadRows_();
    const rows = buildList(payload, cfg, existing);
    rows.forEach(r => computeRow(r, cfg));
    writeAll_(rows);
    const s = stats(rows);
    log_('產生名單', '', '', '', '完成', `共 ${s.total} 筆：符合 ${s.ok}、待確認 ${s.review}、不符合 ${s.fail}`);
    return payload_(rows, false);
  });
}

function apiUpdateRow(token, id, patch) {
  guard_(token);
  return withLock_(() => {
    const cfg = getConfig_();
    const rows = loadRows_();
    const r = rows.find(x => x.id === id);
    if (!r) throw new Error('找不到這筆資料，請重新整理頁面');
    applyPatch(r, patch);
    if (r.pdfStale) { trashPdf_(r); delete r.pdfStale; }
    computeRow(r, cfg);
    writeOne_(r);
    return { row: publicRow_(r), stats: stats(rows) };
  });
}

function apiAddRow(token, data) {
  guard_(token);
  return withLock_(() => {
    const cfg = getConfig_();
    const rows = loadRows_();
    const name = cleanText(data.name);
    if (!name) throw new Error('請填姓名');
    const r = blankRow_();
    Object.assign(r, {
      id: 'M:' + Date.now(), seq: nextSeq(rows), name: name, email: normEmail(data.email),
      school: cleanText(data.school), phone: normPhone(data.phone), source: '手動新增',
      inManual: data.checkIn ? 'Y' : '', outManual: data.checkOut ? 'Y' : '',
      surveyFile: data.survey ? 'Y' : '', flags: 'MANUAL_ADD'
    });
    computeRow(r, cfg);
    rows.push(r);
    writeAll_(rows);
    log_('手動新增', '', r.name, r.email, '完成', '');
    return payload_(rows, false);
  });
}

function apiRemoveRow(token, id) {
  guard_(token);
  return withLock_(() => {
    const rows = loadRows_();
    const r = rows.find(x => x.id === id);
    if (!r) throw new Error('找不到這筆資料');
    if (SENT_STATUSES.indexOf(r.status) > -1) throw new Error('已寄出的資料不能刪除');
    trashPdf_(r);
    const left = rows.filter(x => x.id !== id);
    writeAll_(left);
    log_('刪除', r.certNo, r.name, r.email, '完成', '');
    return payload_(left, false);
  });
}

/** 預覽：產生 1 張 PDF 回傳給網頁顯示，不改變任何狀態 */
function apiPreview(token, id) {
  guard_(token);
  const cfg = getConfig_();
  requireTemplate_();
  const r = previewRow_(id, cfg);
  const file = makePdf_(r, cfg, folder_('PREVIEW_FOLDER_ID'), '預覽_' + r.name + '.pdf');
  const out = { base64: Utilities.base64Encode(file.getBlob().getBytes()), fileName: file.getName() };
  file.setTrashed(true);
  return out;
}

function apiSendTest(token, id, to) {
  guard_(token);
  to = normEmail(to);
  if (!isEmail(to)) throw new Error('測試收件信箱格式有誤');
  const cfg = getConfig_();
  requireTemplate_();
  const r = previewRow_(id, cfg);
  const file = makePdf_(r, cfg, folder_('PREVIEW_FOLDER_ID'), '測試_' + r.name + '.pdf');
  sendMail_(to, r, cfg, file.getBlob(), '[測試] ');
  file.setTrashed(true);
  log_('測試寄送', r.certNo, r.name, to, '成功', '');
  return { to: to, quota: MailApp.getRemainingDailyQuota() };
}

/** 產出 PDF（分段執行，網頁會重複呼叫直到 remaining = 0） */
function apiGenerateBatch(token) {
  guard_(token);
  return withLock_(() => {
    const cfg = getConfig_();
    requireTemplate_();
    const rows = loadRows_();
    const start = Date.now();
    let done = 0, failed = 0;
    const targets = rows.filter(r => r.approved === '✓' && !r.pdfId);
    for (let i = 0; i < targets.length; i++) {
      if ((Date.now() - start) / 1000 > BATCH_SECONDS) break;
      const r = targets[i];
      try {
        ensurePdf_(r, cfg, rows);
        r.error = '';
        done++;
      } catch (e) {
        r.error = 'PDF 產出失敗：' + e.message;
        r.pdfId = '';
        failed++;
      }
      writeOne_(r);
    }
    const remaining = rows.filter(r => r.approved === '✓' && !r.pdfId && !r.error).length;
    return { done: done, failed: failed, remaining: remaining, stats: stats(rows) };
  });
}

/** 範本或設定修改後：清掉「尚未寄出」者的 PDF，下次會重新產出 */
function apiResetPdfs(token) {
  guard_(token);
  return withLock_(() => {
    const rows = loadRows_();
    let n = 0;
    rows.forEach(r => {
      if (SENT_STATUSES.indexOf(r.status) === -1 && (r.pdfId || r.error)) {
        trashPdf_(r); r.error = ''; n++; writeOne_(r);
      }
    });
    log_('清除未寄出PDF', '', '', '', '完成', `${n} 筆`);
    return { cleared: n, stats: stats(rows) };
  });
}

/** 取得單張正式 PDF（打包下載用）；尚未產出會先產出 */
function apiGetPdf(token, id) {
  guard_(token);
  return withLock_(() => {
    const cfg = getConfig_();
    const rows = loadRows_();
    const r = rows.find(x => x.id === id);
    if (!r) throw new Error('找不到這筆資料');
    const file = ensurePdf_(r, cfg, rows);
    writeOne_(r);
    return { base64: Utilities.base64Encode(file.getBlob().getBytes()), fileName: file.getName() };
  });
}

/** 正式寄出（分段執行）：只寄「核發 ✓ 且 尚未寄出」 */
function apiSendBatch(token) {
  guard_(token);
  return withLock_(() => {
    const cfg = getConfig_();
    requireTemplate_();
    const rows = loadRows_();
    const start = Date.now();
    let quota = MailApp.getRemainingDailyQuota();
    let sent = 0, failed = 0;
    const targets = rows.filter(isSendTarget);
    for (let i = 0; i < targets.length; i++) {
      if ((Date.now() - start) / 1000 > BATCH_SECONDS) break;
      if (quota < 1) break;
      const r = targets[i];
      try {
        if (!isEmail(r.email)) throw new Error('Email 格式有誤');
        const file = ensurePdf_(r, cfg, rows);
        sendMail_(r.email, r, cfg, file.getBlob(), '');
        r.status = '已寄出'; r.sentAt = now_(); r.error = '';
        sent++; quota--;
        log_('正式寄出', r.certNo, r.name, r.email, '成功', '');
      } catch (e) {
        r.status = '失敗'; r.error = e.message;
        failed++;
        log_('正式寄出', r.certNo, r.name, r.email, '失敗', e.message);
      }
      writeOne_(r);
    }
    return {
      sent: sent, failed: failed,
      remaining: rows.filter(isSendTarget).length,
      quota: MailApp.getRemainingDailyQuota(),
      stats: stats(rows)
    };
  });
}

/** 補寄單筆；to 有填就改寄到新信箱（並更新名單上的 Email） */
function apiResend(token, id, to) {
  guard_(token);
  return withLock_(() => {
    const cfg = getConfig_();
    requireTemplate_();
    const rows = loadRows_();
    const r = rows.find(x => x.id === id);
    if (!r) throw new Error('找不到這筆資料');
    if (to) applyPatch(r, { email: to });
    if (r.pdfStale) { trashPdf_(r); delete r.pdfStale; }
    if (!isEmail(r.email)) throw new Error('Email 格式有誤');
    try {
      const file = ensurePdf_(r, cfg, rows);
      sendMail_(r.email, r, cfg, file.getBlob(), '');
      r.status = '已補寄'; r.sentAt = now_(); r.error = '';
      log_('補寄', r.certNo, r.name, r.email, '成功', '');
    } catch (e) {
      r.status = '失敗'; r.error = e.message;
      log_('補寄', r.certNo, r.name, r.email, '失敗', e.message);
    }
    computeRow(r, cfg);
    writeOne_(r);
    return { row: publicRow_(r), stats: stats(rows) };
  });
}

/** 把「失敗」的全部重新排入寄送 */
function apiRequeueFailed(token) {
  guard_(token);
  return withLock_(() => {
    const rows = loadRows_();
    let n = 0;
    rows.forEach(r => { if (r.status === '失敗') { r.status = ''; n++; writeOne_(r); } });
    return { requeued: n, stats: stats(rows) };
  });
}

function apiLogs(token) {
  guard_(token);
  const sh = ss_().getSheetByName(LOG_SHEET);
  const last = sh.getLastRow();
  if (last < 2) return [];
  const n = Math.min(300, last - 1);
  return sh.getRange(last - n + 1, 1, n, 7).getDisplayValues().reverse();
}

function apiSaveConfig(token, input) {
  guard_(token);
  return withLock_(() => {
    const cfg = getConfig_();
    Object.keys(DEFAULT_CONFIG).forEach(k => {
      if (input[k] === undefined) return;
      cfg[k] = typeof DEFAULT_CONFIG[k] === 'boolean' ? !!input[k] : String(input[k]);
    });
    if (cfg.replyTo && !isEmail(cfg.replyTo)) throw new Error('回覆信箱格式有誤');
    cfg.fromAlias = normEmail(cfg.fromAlias);
    if (cfg.fromAlias && GmailApp.getAliases().map(a => a.toLowerCase()).indexOf(cfg.fromAlias) === -1) {
      throw new Error('「' + cfg.fromAlias + '」不是寄件帳號可用的別名，請先在 Gmail 設定「以這個地址寄送郵件」');
    }
    props_().setProperty('CONFIG', JSON.stringify(cfg));
    const rows = loadRows_();
    rows.forEach(r => computeRow(r, cfg));
    writeAll_(rows);
    log_('修改設定', '', '', '', '完成', '');
    return payload_(rows, true);
  });
}

/** 登入：成功回傳 token 與初始資料 */
function apiLogin(user, pass) {
  const p = props_();
  if (!p.getProperty('SS_ID')) throw new Error('工具尚未初始化：請在程式編輯器執行一次 setup');
  initLogin_();
  const cache = CacheService.getScriptCache();
  const fails = Number(cache.get('loginFails') || 0);
  if (fails >= 10) throw new Error('登入失敗次數過多，請 15 分鐘後再試');
  const okUser = cleanText(user).toLowerCase() === String(p.getProperty('LOGIN_USER')).toLowerCase();
  const okPass = hash_(String(pass == null ? '' : pass), p.getProperty('LOGIN_SALT')) === p.getProperty('LOGIN_HASH');
  if (!okUser || !okPass) {
    cache.put('loginFails', String(fails + 1), 900);
    throw new Error('帳號或密碼錯誤');
  }
  const token = Utilities.getUuid();
  cache.put('tok_' + token, p.getProperty('LOGIN_VER'), 21600);
  const out = payload_(loadRows_(), true);
  out.token = token;
  return out;
}

function apiLogout(token) {
  if (token) CacheService.getScriptCache().remove('tok_' + token);
  return true;
}

/** 修改帳號密碼：成功後其他已登入的人需重新登入 */
function apiChangeLogin(token, curPass, newUser, newPass) {
  guard_(token);
  const p = props_();
  if (hash_(String(curPass || ''), p.getProperty('LOGIN_SALT')) !== p.getProperty('LOGIN_HASH')) throw new Error('目前密碼不正確');
  newUser = cleanText(newUser);
  newPass = String(newPass || '');
  if (!newUser) throw new Error('帳號不能空白');
  if (newPass.length < 6) throw new Error('新密碼至少 6 碼');
  const salt = Utilities.getUuid();
  const ver = String(Number(p.getProperty('LOGIN_VER') || 1) + 1);
  p.setProperties({ LOGIN_USER: newUser, LOGIN_SALT: salt, LOGIN_HASH: hash_(newPass, salt), LOGIN_VER: ver, LOGIN_DEFAULT: '' });
  const cache = CacheService.getScriptCache();
  cache.remove('tok_' + token);
  const fresh = Utilities.getUuid();
  cache.put('tok_' + fresh, ver, 21600);
  log_('修改帳號密碼', '', '', '', '完成', '帳號：' + newUser);
  return { token: fresh, user: newUser };
}

/** 忘記密碼時：在程式編輯器執行這個函式，帳密會回到 itmonth / itmonth */
function resetLogin() {
  props_().deleteProperty('LOGIN_HASH');
  initLogin_();
  Logger.log('已重設：帳號 itmonth／密碼 itmonth');
}

/** 上傳 pptx 範本 → 自動轉成 Google 簡報 */
function apiUploadTemplate(token, base64, fileName) {
  guard_(token);
  const blob = Utilities.newBlob(Utilities.base64Decode(base64),
    'application/vnd.openxmlformats-officedocument.presentationml.presentation', fileName);
  const created = Drive.Files.create({
    name: '證書範本_' + Utilities.formatDate(new Date(), TZ, 'yyyyMMdd_HHmm'),
    mimeType: 'application/vnd.google-apps.presentation',
    parents: [props_().getProperty('FOLDER_ID')]
  }, blob);
  const info = checkTemplate_(created.id);
  if (!info.ok) {
    DriveApp.getFileById(created.id).setTrashed(true);
    throw new Error('範本裡找不到 {{姓名}}，請確認文字框內有 {{姓名}}（全半形需一致）');
  }
  props_().setProperty('TEMPLATE_ID', created.id);
  log_('更換範本', '', '', '', '完成', fileName);
  return info;
}

/** 改用既有的 Google 簡報當範本（貼網址） */
function apiSetTemplateUrl(token, url) {
  guard_(token);
  const m = String(url || '').match(/\/d\/([a-zA-Z0-9_-]{20,})/);
  const id = m ? m[1] : String(url || '').trim();
  let info;
  try { info = checkTemplate_(id); } catch (e) { throw new Error('打不開這份簡報，請確認網址與權限'); }
  if (!info.ok) throw new Error('範本裡找不到 {{姓名}}');
  props_().setProperty('TEMPLATE_ID', id);
  log_('更換範本', '', '', '', '完成', url);
  return info;
}

/* =========================================================
 * 3. 比對邏輯（純函式，不呼叫 Google 服務，方便測試）
 * ========================================================= */

function cleanText(v) {
  return String(v == null ? '' : v).normalize('NFKC').replace(/\s+/g, ' ').trim();
}
function normName(v) {
  return cleanText(v).replace(/\s|·|‧|・|\./g, '').toLowerCase();
}
function normEmail(v) {
  return cleanText(v).replace(/\s/g, '').replace(/^mailto:/i, '').toLowerCase();
}
function normPhone(v) {
  const d = String(v == null ? '' : v).normalize('NFKC').replace(/\D/g, '');
  if (!d) return '';
  return d.length >= 9 ? '0' + d.slice(-9) : d;
}
function isEmail(v) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/.test(String(v || ''));
}
function isMark(v) {
  const s = cleanText(v);
  if (!s) return false;
  return !/^(n|no|否|x|✗|✘|×|0|false|未|無|-|—)$/i.test(s);
}
function splitList(s) {
  return String(s || '').split(/[,，、;；\n]/).map(x => x.trim()).filter(Boolean);
}
function matchIdentity(value, keywords) {
  const v = cleanText(value);
  if (!v) return false;
  return keywords.some(k => v.indexOf(k) > -1);
}
function fmtSeq(n) { return 'R' + String(n).padStart(3, '0'); }
function seqNum(s) { const m = String(s || '').match(/(\d+)$/); return m ? parseInt(m[1], 10) : 0; }
function nextSeq(rows) { return fmtSeq(rows.reduce((m, r) => Math.max(m, seqNum(r.seq)), 0) + 1); }

function blankRow_() {
  const r = {};
  FIELDS.forEach(f => { r[f[0]] = ''; });
  return r;
}

/** 從勾選欄或「簽到/簽退」類型欄判斷進出 */
function readInOut(c, map) {
  if (map && map.type) {
    const t = cleanText(c.type);
    const out = /退|離|out/i.test(t);
    const inn = !out && /到|入|in/i.test(t);
    return { in: inn, out: out };
  }
  if (map && (map.in || map.out)) {
    return { in: map.in ? isMark(c.in) : false, out: map.out ? isMark(c.out) : false };
  }
  return { in: true, out: true }; // 只有名單、沒有勾選欄 → 視為有出席
}

/**
 * 產生核發名單
 * payload = { mode:'match'|'direct', reg:[...], filterIdentity:bool,
 *             checkin:[...]|null, checkinMap:{type,in,out}, survey:[...]|null }
 * existing = 目前試算表裡的名單（保留手動勾選、證書編號、寄送狀態）
 */
function buildList(payload, cfg, existing) {
  existing = existing || [];
  const exById = {};
  existing.forEach(e => { exById[e.id] = e; });
  const keywords = splitList(payload.identityKeywords || cfg.identityKeywords);
  const hasCheckin = Array.isArray(payload.checkin);
  const hasSurvey = Array.isArray(payload.survey);
  const direct = payload.mode === 'direct';

  // ① 報名名單 → 篩身分、去重
  const rows = [];
  const byId = {};
  (payload.reg || []).forEach(src => {
    const name = cleanText(src.name);
    const email = normEmail(src.email);
    const phone = normPhone(src.phone);
    if (!name && !email) return;
    if (!direct && payload.filterIdentity && !matchIdentity(src.identity, keywords)) return;
    const id = email ? 'E:' + email : 'N:' + normName(name) + '|' + phone;
    if (byId[id]) { addFlag_(byId[id], 'DUP'); return; }
    const r = blankRow_();
    Object.assign(r, { id: id, name: name, email: email, phone: phone, school: cleanText(src.school), source: '報名' });
    if (direct) { r.inFile = 'Y'; r.outFile = 'Y'; r.surveyFile = 'Y'; }
    byId[id] = r;
    rows.push(r);
  });

  // ② 合併舊資料（編號、證書編號、手動修改、寄送狀態都保留）
  let maxSeq = existing.reduce((m, e) => Math.max(m, seqNum(e.seq)), 0);
  rows.forEach(r => {
    const e = exById[r.id];
    if (!e) { r.seq = fmtSeq(++maxSeq); return; }
    ['seq', 'certNo', 'status', 'sentAt', 'error', 'pdfId', 'inManual', 'outManual', 'approve', 'edited']
      .forEach(k => { r[k] = e[k]; });
    splitList(e.edited).forEach(k => { r[k] = e[k]; });
    if (!direct && !hasCheckin) { r.inFile = e.inFile; r.outFile = e.outFile; }
    if (!direct && !hasSurvey) r.surveyFile = e.surveyFile;
    const keep = splitList(e.flags).filter(f => (f === 'NAME_MATCH' && !hasCheckin) || (f === 'SURVEY_BY_NAME' && !hasSurvey));
    keep.forEach(f => addFlag_(r, f));
  });

  // 姓名唯一性（同名者不用姓名比對）
  const nameCount = {};
  rows.forEach(r => { const n = normName(r.name); nameCount[n] = (nameCount[n] || 0) + 1; });

  // ③ 簽到退比對：編號 → Email → 手機 → 姓名
  if (!direct && hasCheckin) {
    const chk = payload.checkin.map(c => {
      const io = readInOut(c, payload.checkinMap);
      return {
        code: cleanText(c.code).toUpperCase(), name: cleanText(c.name), nname: normName(c.name),
        email: normEmail(c.email), phone: normPhone(c.phone), school: cleanText(c.school),
        in: io.in, out: io.out, used: false
      };
    }).filter(c => c.code || c.nname || c.email || c.phone);

    rows.forEach(r => {
      let hits = chk.filter(c => c.code && c.code === r.seq);
      let byName = false;
      if (!hits.length && r.email) hits = chk.filter(c => c.email && c.email === r.email);
      if (!hits.length && r.phone) hits = chk.filter(c => c.phone && c.phone === r.phone);
      if (!hits.length) {
        const n = normName(r.name);
        if (n && nameCount[n] === 1) {
          hits = chk.filter(c => !c.used && c.nname === n && !(c.code && c.code !== r.seq && /^R\d+$/.test(c.code)));
          byName = hits.length > 0;
        }
      }
      hits.forEach(c => { c.used = true; });
      r.inFile = hits.some(c => c.in) ? 'Y' : '';
      r.outFile = hits.some(c => c.out) ? 'Y' : '';
      if (byName) addFlag_(r, 'NAME_MATCH');
    });

    // 有簽到但不在報名名單 → 現場（未報名）
    const walk = {};
    chk.filter(c => !c.used).forEach(c => {
      const key = c.email || (c.phone ? 'P' + c.phone : '') || 'N' + c.nname;
      if (!walk[key]) walk[key] = { c: c, in: false, out: false };
      walk[key].in = walk[key].in || c.in;
      walk[key].out = walk[key].out || c.out;
    });
    Object.keys(walk).forEach(key => {
      const w = walk[key];
      if (!w.c.name && !w.c.email) return;
      const id = 'W:' + key;
      const r = blankRow_();
      Object.assign(r, {
        id: id, name: w.c.name, email: w.c.email, phone: w.c.phone, school: w.c.school,
        source: '現場（未報名）', inFile: w.in ? 'Y' : '', outFile: w.out ? 'Y' : '', flags: 'WALK_IN'
      });
      const e = exById[id];
      if (e) {
        ['seq', 'certNo', 'status', 'sentAt', 'error', 'pdfId', 'inManual', 'outManual', 'approve', 'edited', 'surveyFile']
          .forEach(k => { r[k] = e[k]; });
        splitList(e.edited).forEach(k => { r[k] = e[k]; });
      } else {
        r.seq = fmtSeq(++maxSeq);
      }
      byId[id] = r;
      rows.push(r);
    });
  }

  // ④ 舊資料中要保留的：已寄出、手動新增、（沒重新上傳簽到檔時的）現場名單
  existing.forEach(e => {
    if (byId[e.id]) return;
    const sent = SENT_STATUSES.indexOf(e.status) > -1;
    const manual = e.source === '手動新增';
    const walkKeep = e.source === '現場（未報名）' && !hasCheckin && !direct;
    if (sent || manual || walkKeep) {
      const r = Object.assign(blankRow_(), e);
      delete r._row;
      if (sent && !manual && !walkKeep) addFlag_(r, 'NOT_IN_LIST');
      byId[r.id] = r;
      rows.push(r);
    }
  });

  // ⑤ 問卷比對：Email → 手機 → 姓名
  if (!direct && hasSurvey) {
    const sv = payload.survey.map(s => ({ email: normEmail(s.email), phone: normPhone(s.phone), nname: normName(s.name) }));
    const svEmail = {}, svPhone = {}, svName = {};
    sv.forEach(s => {
      if (s.email) svEmail[s.email] = true;
      if (s.phone) svPhone[s.phone] = true;
      if (s.nname) svName[s.nname] = (svName[s.nname] || 0) + 1;
    });
    const allNames = {};
    rows.forEach(r => { const n = normName(r.name); allNames[n] = (allNames[n] || 0) + 1; });
    rows.forEach(r => {
      removeFlag_(r, 'SURVEY_BY_NAME');
      if (r.email && svEmail[r.email]) { r.surveyFile = 'Y'; return; }
      if (r.phone && svPhone[r.phone]) { r.surveyFile = 'Y'; return; }
      const n = normName(r.name);
      if (n && svName[n] === 1 && allNames[n] === 1) { r.surveyFile = 'Y'; addFlag_(r, 'SURVEY_BY_NAME'); return; }
      if (r.source !== '手動新增') r.surveyFile = '';
    });
  }

  rows.sort((a, b) => seqNum(a.seq) - seqNum(b.seq));
  return rows;
}

function addFlag_(r, f) {
  const s = splitList(r.flags);
  if (s.indexOf(f) === -1) s.push(f);
  r.flags = s.join(',');
}
function removeFlag_(r, f) {
  r.flags = splitList(r.flags).filter(x => x !== f).join(',');
}

/** 依旗標與設定計算：簽到/簽退/問卷/比對結果/備註/核發 */
function computeRow(r, cfg) {
  const flags = splitList(r.flags);
  const has = f => flags.indexOf(f) > -1;
  const inOk = r.inManual ? r.inManual === 'Y' : r.inFile === 'Y';
  const outOk = r.outManual ? r.outManual === 'Y' : r.outFile === 'Y';
  const svOk = r.surveyFile === 'Y';
  const human = !!(r.inManual || r.outManual);

  const needIn = cfg.requireCheckin !== false;
  const reasons = [];
  if (needIn && !inOk) reasons.push('未簽到');
  if (needIn && cfg.requireOut && !outOk) reasons.push('未簽退');
  if (cfg.requireSurvey && !svOk) reasons.push('未填問卷');
  if (!r.email) reasons.push('缺 Email');
  else if (!isEmail(r.email)) reasons.push('Email 格式有誤');

  const review = [];
  if (has('NAME_MATCH') && !human) review.push('簽到以姓名比對，請核對紙本');
  if (has('WALK_IN')) review.push('未在報名名單');
  if (has('MANUAL_ADD')) review.push('手動新增');

  const info = [];
  if (has('DUP')) info.push('報名重複，已合併');
  if (has('SURVEY_BY_NAME')) info.push('問卷以姓名比對');
  if (has('NOT_IN_LIST')) info.push('已不在最新報名名單');

  const noAttendance = needIn && !inOk && !outOk;        // 完全沒出席紀錄
  const noSurveyOnly = !needIn && cfg.requireSurvey && !svOk; // 不核對簽到時，沒問卷 = 無任何完成證據
  let result;
  if (noAttendance || noSurveyOnly) result = '不符合';
  else if (reasons.length || review.length) result = '待確認';
  else result = '符合';

  r.checkIn = inOk ? '✓' : '';
  r.checkOut = outOk ? '✓' : '';
  r.survey = svOk ? '✓' : '';
  r.result = result;
  r.note = reasons.concat(review, info).join('；');
  r.approved = (r.approve ? r.approve === 'Y' : result === '符合') ? '✓' : '';
  return r;
}

/** 網頁上的修改（姓名、Email、學校、勾選） */
function applyPatch(r, patch) {
  const edited = splitList(r.edited);
  const mark = k => { if (edited.indexOf(k) === -1) edited.push(k); };
  if (patch.name !== undefined) {
    const v = cleanText(patch.name);
    if (!v) throw new Error('姓名不能空白');
    if (v !== r.name) { r.name = v; mark('name'); if (r.pdfId) r.pdfStale = true; }
  }
  if (patch.school !== undefined) {
    const v = cleanText(patch.school);
    if (v !== r.school) { r.school = v; mark('school'); if (r.pdfId) r.pdfStale = true; }
  }
  if (patch.email !== undefined) {
    const v = normEmail(patch.email);
    if (v && !isEmail(v)) throw new Error('Email 格式有誤：' + v);
    if (v !== r.email) {
      r.email = v; mark('email');
      if (r.status === '失敗') { r.status = ''; r.error = ''; }
    }
  }
  ['inManual', 'outManual', 'approve'].forEach(k => {
    if (patch[k] === undefined) return;
    const v = String(patch[k]);
    if (['', 'Y', 'N'].indexOf(v) === -1) throw new Error('參數錯誤');
    r[k] = v;
  });
  r.edited = edited.join(',');
  return r;
}

function isSendTarget(r) {
  return r.approved === '✓' && !r.status;
}

function stats(rows) {
  const s = { total: rows.length, ok: 0, review: 0, fail: 0, approved: 0, sent: 0, failed: 0, toSend: 0, withPdf: 0 };
  rows.forEach(r => {
    if (r.result === '符合') s.ok++;
    else if (r.result === '待確認') s.review++;
    else s.fail++;
    if (r.approved === '✓') s.approved++;
    if (SENT_STATUSES.indexOf(r.status) > -1) s.sent++;
    if (r.status === '失敗') s.failed++;
    if (isSendTarget(r)) s.toSend++;
    if (r.approved === '✓' && r.pdfId) s.withPdf++;
  });
  return s;
}

/** 套用 {{欄位}} 文字 */
function fillText(text, r, cfg) {
  const map = valueMap(r, cfg);
  return Object.keys(map).reduce((t, k) => t.split(k).join(map[k]), String(text || ''));
}
function valueMap(r, cfg) {
  return {
    '{{姓名}}': r.name || '', '{{年度}}': cfg.year || '', '{{時數}}': cfg.hours || '',
    '{{發證年月}}': cfg.issueMonth || '', '{{證書編號}}': r.certNo || '', '{{學校}}': r.school || ''
  };
}

/** 姓名太長時自動縮小字級 */
function nameFontSize(name, base) {
  let w = 0;
  for (const ch of String(name || '')) w += ch.charCodeAt(0) > 0x2E80 ? 1 : 0.55;
  if (w <= 8) return base;
  return Math.max(18, Math.floor(base * 8 / w));
}

function nextCertNo(rows, prefix) {
  let max = 0;
  rows.forEach(r => {
    if (r.certNo && String(r.certNo).indexOf(prefix) === 0) max = Math.max(max, seqNum(r.certNo));
  });
  return prefix + String(max + 1).padStart(3, '0');
}

/* =========================================================
 * 4. PDF 與寄信
 * ========================================================= */

function makePdf_(r, cfg, folder, fileName) {
  const tplId = props_().getProperty('TEMPLATE_ID');
  const copy = DriveApp.getFileById(tplId).makeCopy('tmp_' + (r.certNo || r.name), folder_('TEMP_FOLDER_ID'));
  try {
    const pres = SlidesApp.openById(copy.getId());
    // 找出放 {{姓名}} 的文字框，記下原字級
    const nameShapes = [];
    pres.getSlides().forEach(slide => eachShape_(slide.getPageElements(), sh => {
      const t = sh.getText();
      if (t.asString().indexOf('{{姓名}}') > -1) nameShapes.push({ sh: sh, size: t.getTextStyle().getFontSize() || 50 });
    }));
    const map = valueMap(r, cfg);
    Object.keys(map).forEach(k => pres.replaceAllText(k, String(map[k])));
    nameShapes.forEach(x => {
      const size = nameFontSize(r.name, x.size);
      if (size !== x.size) x.sh.getText().getTextStyle().setFontSize(size);
    });
    pres.saveAndClose();
    const pdf = DriveApp.getFileById(copy.getId()).getAs('application/pdf').setName(fileName);
    return folder.createFile(pdf);
  } finally {
    copy.setTrashed(true);
  }
}

function eachShape_(elements, fn) {
  elements.forEach(el => {
    const type = el.getPageElementType();
    if (type === SlidesApp.PageElementType.SHAPE) fn(el.asShape());
    else if (type === SlidesApp.PageElementType.GROUP) eachShape_(el.asGroup().getChildren(), fn);
  });
}

function ensurePdf_(r, cfg, rows) {
  if (r.pdfId) {
    try {
      const f = DriveApp.getFileById(r.pdfId);
      if (!f.isTrashed()) return f;
    } catch (e) { /* 檔案不見了就重做 */ }
  }
  if (!r.certNo) r.certNo = nextCertNo(rows, cfg.certPrefix);
  const file = makePdf_(r, cfg, folder_('PDF_FOLDER_ID'), safeName_(r.certNo + '_' + r.name) + '.pdf');
  r.pdfId = file.getId();
  return file;
}

function trashPdf_(r) {
  if (!r.pdfId) return;
  try { DriveApp.getFileById(r.pdfId).setTrashed(true); } catch (e) { /* 忽略 */ }
  r.pdfId = '';
}

function sendMail_(to, r, cfg, pdfBlob, subjectPrefix) {
  const subject = subjectPrefix + fillText(cfg.subject, r, cfg);
  const body = fillText(cfg.body, r, cfg);
  const opts = { name: cfg.senderName, attachments: [pdfBlob], htmlBody: emailHtml_(body) };
  if (cfg.replyTo) opts.replyTo = cfg.replyTo;
  if (cfg.fromAlias) opts.from = cfg.fromAlias;
  GmailApp.sendEmail(to, subject, body, opts);
}

function emailHtml_(text) {
  const esc = String(text).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  return '<div style="font-family:\'Noto Sans TC\',\'PingFang TC\',\'Microsoft JhengHei\',sans-serif;'
    + 'font-size:15px;line-height:1.8;color:#333">' + esc.replace(/\n/g, '<br>') + '</div>';
}

function previewRow_(id, cfg) {
  if (id) {
    const r = loadRows_().find(x => x.id === id);
    if (r) {
      const c = Object.assign({}, r);
      if (!c.certNo) c.certNo = cfg.certPrefix + '000';
      return c;
    }
  }
  return Object.assign(blankRow_(), { name: '王小明', school: '範例大學', certNo: cfg.certPrefix + '000' });
}

function requireTemplate_() {
  const id = props_().getProperty('TEMPLATE_ID');
  if (!id) throw new Error('尚未設定證書範本，請到「設定」頁上傳');
}

function checkTemplate_(id) {
  const pres = SlidesApp.openById(id);
  let text = '';
  pres.getSlides().forEach(s => eachShape_(s.getPageElements(), sh => { text += sh.getText().asString(); }));
  const found = PLACEHOLDERS.filter(p => text.indexOf(p) > -1);
  return { id: id, name: pres.getName(), url: pres.getUrl(), found: found, ok: found.indexOf('{{姓名}}') > -1 };
}

/* =========================================================
 * 5. 試算表存取與共用工具
 * ========================================================= */

function props_() { return PropertiesService.getScriptProperties(); }

function ss_() {
  const id = props_().getProperty('SS_ID');
  if (!id) throw new Error('工具尚未初始化：請在程式編輯器執行一次 setup');
  return SpreadsheetApp.openById(id);
}

function folder_(key) { return DriveApp.getFolderById(props_().getProperty(key)); }

function getConfig_() {
  let saved = {};
  try { saved = JSON.parse(props_().getProperty('CONFIG') || '{}'); } catch (e) { saved = {}; }
  return Object.assign({}, DEFAULT_CONFIG, saved);
}

function guard_(token) {
  const p = props_();
  if (!p.getProperty('LOGIN_HASH')) throw new Error('工具尚未初始化：請在程式編輯器執行一次 setup');
  const cache = CacheService.getScriptCache();
  const v = token ? cache.get('tok_' + token) : null;
  if (!v || v !== p.getProperty('LOGIN_VER')) throw new Error('登入已逾時，請重新登入');
  cache.put('tok_' + token, v, 21600); // 使用中自動延長（閒置 6 小時登出）
}

function initLogin_() {
  const p = props_();
  if (p.getProperty('LOGIN_HASH')) return;
  const salt = Utilities.getUuid();
  p.setProperties({
    LOGIN_USER: 'itmonth', LOGIN_SALT: salt, LOGIN_HASH: hash_('itmonth', salt),
    LOGIN_VER: String(Number(p.getProperty('LOGIN_VER') || 0) + 1), LOGIN_DEFAULT: 'Y'
  });
}

function hash_(pass, salt) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + '|' + pass, Utilities.Charset.UTF_8);
  return bytes.map(b => ('0' + (b & 0xff).toString(16)).slice(-2)).join('');
}

function withLock_(fn) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) throw new Error('有其他同仁正在執行作業，請稍候再試');
  try { return fn(); } finally { lock.releaseLock(); }
}

function loadRows_() {
  const sh = ss_().getSheetByName(LIST_SHEET);
  const values = sh.getDataRange().getDisplayValues();
  if (values.length < 2) return [];
  const idx = {};
  values[0].forEach((h, i) => { idx[h] = i; });
  const rows = [];
  for (let i = 1; i < values.length; i++) {
    if (values[i].join('') === '') continue;
    const r = { _row: i + 1 };
    FIELDS.forEach(f => { const c = idx[f[1]]; r[f[0]] = c === undefined ? '' : values[i][c]; });
    rows.push(r);
  }
  return rows;
}

function rowArray_(r) { return FIELDS.map(f => (r[f[0]] == null ? '' : String(r[f[0]]))); }

function writeAll_(rows) {
  const sh = ss_().getSheetByName(LIST_SHEET);
  const last = sh.getLastRow();
  if (last > 1) sh.getRange(2, 1, last - 1, FIELDS.length).clearContent();
  if (rows.length) {
    const range = sh.getRange(2, 1, rows.length, FIELDS.length);
    range.setNumberFormat('@');
    range.setValues(rows.map(rowArray_));
  }
  rows.forEach((r, i) => { r._row = i + 2; });
}

function writeOne_(r) {
  if (!r._row) throw new Error('資料列位置遺失，請重新整理頁面');
  ss_().getSheetByName(LIST_SHEET).getRange(r._row, 1, 1, FIELDS.length).setValues([rowArray_(r)]);
}

function log_(action, certNo, name, email, result, detail) {
  try {
    ss_().getSheetByName(LOG_SHEET).appendRow([now_(), action, certNo, name, email, result, detail]);
  } catch (e) { /* 紀錄失敗不影響主流程 */ }
}

function now_() { return Utilities.formatDate(new Date(), TZ, 'yyyy/MM/dd HH:mm:ss'); }

function safeName_(s) { return String(s).replace(/[\\/:*?"<>|]/g, '_'); }

function publicRow_(r) {
  return {
    id: r.id, seq: r.seq, certNo: r.certNo, name: r.name, email: r.email, school: r.school,
    phone: r.phone, source: r.source, checkIn: r.checkIn === '✓', checkOut: r.checkOut === '✓',
    survey: r.survey === '✓', result: r.result, note: r.note, approved: r.approved === '✓',
    status: r.status, sentAt: r.sentAt, error: r.error, hasPdf: !!r.pdfId,
    inManual: r.inManual, outManual: r.outManual, approve: r.approve
  };
}

function payload_(rows, withMeta) {
  const out = { rows: rows.map(publicRow_), stats: stats(rows) };
  if (withMeta) {
    const p = props_();
    out.config = getConfig_();
    out.quota = MailApp.getRemainingDailyQuota();
    out.sender = Session.getEffectiveUser().getEmail();
    out.aliases = GmailApp.getAliases();
    out.user = p.getProperty('LOGIN_USER');
    out.isDefaultLogin = p.getProperty('LOGIN_DEFAULT') === 'Y';
    out.links = {
      sheet: ss_().getUrl(),
      folder: DriveApp.getFolderById(p.getProperty('FOLDER_ID')).getUrl()
    };
    out.template = null;
    const tid = p.getProperty('TEMPLATE_ID');
    if (tid) {
      try { out.template = checkTemplate_(tid); } catch (e) { out.template = { id: tid, ok: false, error: '範本無法開啟' }; }
    }
  }
  return out;
}
