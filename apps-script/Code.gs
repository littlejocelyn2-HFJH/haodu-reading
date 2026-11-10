/**
 * 好讀・讀報日記 — Apps Script 後端（只負責收資料，網頁放在 GitHub Pages）
 * 放在「好讀讀報日記資料」試算表的 擴充功能 → Apps Script。
 * 部署：網頁應用程式｜執行身分：我｜誰可以存取：所有人
 * 改了這個檔案後：部署 → 管理部署作業 → 編輯 → 版本選「新版本」→ 部署（網址不變）
 * 直接打開 /exec 網址看到 {"ok":true,...} 是正常的，那是後端。
 */
const TEACHER_PIN = 'change-me';      // ★ 第 8 行：改成你的教師密碼（保持 change-me 會一律登入失敗）
const IMG_FOLDER = '好讀讀報版面';     // 雲端硬碟中放版面圖的資料夾（不用分享）
const SEND_MAIL = true;                // 學生送出時也寄一封讀報日記到老師信箱

const DATA = '讀報資料';   // 給程式讀：每人每期一列
const LOG = '作答紀錄';    // 給老師看：每次送出一列
const DATA_HEAD = ['最後更新','代號','姓名','期別','答對題數','原始資料'];
const LOG_HEAD = ['送出時間','代號','姓名','期別','答對題數','作答','用時(秒)','讀報(分)','疑似亂答','擷取','推論','統整','評鑑','讀前熟悉','難度','有趣','自評讀懂','讀報方法','第1名','原因1','第2名','原因2','第3名','原因3','最想追問'];

const API = { loadByCode: loadByCode, saveIssue: saveIssue, teacherLogin: teacherLogin, getAll: getAll,
  setShowcase: setShowcase, setType: setType, setBadges: setBadges, getBadges: getBadges, getImage: getImage, clearClass: clearClass };

function doGet() { return json_({ ok: true, app: '好讀・讀報日記 API' }); }
function doPost(e) {
  try {
    const req = JSON.parse(e.postData.contents);
    const fn = API[req.fn];
    if (!fn) throw new Error('未知的動作');
    return json_({ ok: true, data: fn.apply(null, req.args || []) });
  } catch (err) { return json_({ ok: false, error: String(err.message || err) }); }
}
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

function pinOk_(pin) { return TEACHER_PIN !== 'change-me' && !!TEACHER_PIN && String(pin) === String(TEACHER_PIN); }
function teacherLogin(pin) {
  if (TEACHER_PIN === 'change-me' || !TEACHER_PIN) return '尚未設定教師密碼：請改 Code.gs 第 8 行，儲存後重新部署（新版本）';
  Utilities.sleep(800);
  return pinOk_(pin) ? true : '密碼錯誤';
}
function sheet_(name, head) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) { sh = ss.insertSheet(name); sh.appendRow(head); sh.setFrozenRows(1); }
  return sh;
}
function codeOk_(c) { return /^\d{3}-\d{2}$/.test(String(c)); }

/** 學生：用代號找回自己的紀錄 */
function loadByCode(code) {
  if (!codeOk_(code)) return {};
  const v = sheet_(DATA, DATA_HEAD).getDataRange().getValues(), out = {};
  for (let r = 1; r < v.length; r++) if (String(v[r][1]) === String(code)) { try { out[v[r][3]] = JSON.parse(v[r][5]); } catch (e) {} }
  return out;
}

/** 學生送出：讀報資料（同一代號同一期只留最新）＋ 作答紀錄（每次一列） */
function saveIssue(code, name, issue, json, sum) {
  if (!codeOk_(code)) throw new Error('代號格式不正確');
  issue = Number(issue);
  const x = JSON.parse(json);
  name = String(name || '').slice(0, 20);
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = sheet_(DATA, DATA_HEAD), v = sh.getDataRange().getValues();
    let row = -1;
    for (let r = 1; r < v.length; r++) if (String(v[r][1]) === String(code) && Number(v[r][3]) === issue) row = r + 1;
    const d = [new Date(), code, name, issue, sum.score, json];
    if (row > 0) sh.getRange(row, 1, 1, d.length).setValues([d]); else sh.appendRow(d);
    const ab = (sum.abil || []).map(function (p) { return p == null ? '' : Math.round(p * 100) + '%'; });
    const why = x.why || [], rank = x.rank || [];
    sheet_(LOG, LOG_HEAD).appendRow([new Date(), code, name, issue, sum.score, sum.letters, x.elapsed || '',
      Math.round((x.sec || 0) / 6) / 10, sum.sus ? '是' : '', ab[0], ab[1], ab[2], ab[3], x.pre, x.diff, x.int, x.conf,
      (x.strat || []).join('、'), rank[0], why[0], rank[1], why[1], rank[2], why[2], x.ask || '']);
    if (SEND_MAIL) {
      const body = '代號：' + code + (name ? '（' + name + '）' : '') + '\n期別：第' + issue + '期\n答對：' + sum.score + '/3（' + sum.letters + '）\n' +
        '自評讀懂：' + x.conf + '　讀報時間：' + Math.round((x.sec || 0) / 60) + ' 分鐘　用時：' + (x.elapsed || 0) + ' 秒\n\n' +
        [0, 1, 2].map(function (i) { return '第' + (i + 1) + '名：' + (rank[i] || '') + '\n原因：' + (why[i] || ''); }).join('\n\n') +
        '\n\n最想追問：' + (x.ask || '');
      MailApp.sendEmail(Session.getEffectiveUser().getEmail(), '【讀報日記】' + code + ' 第' + issue + '期 答對 ' + sum.score + '/3', body);
    }
  } finally { lock.releaseLock(); }
  return { ok: true, code: code };
}

/** 教師端：全部資料（需密碼） */
function getAll(pin) {
  if (!pinOk_(pin)) throw new Error('只有老師可以讀取');
  const v = sheet_(DATA, DATA_HEAD).getDataRange().getValues(), by = {}, names = JSON.parse(JSON.stringify(NAMES));
  for (let r = 1; r < v.length; r++) {
    const c = String(v[r][1]); if (!c) continue;
    const o = by[c] || (by[c] = { code: c, d: {} });
    try { o.d[v[r][3]] = JSON.parse(v[r][5]); } catch (e) {}
    if (v[r][2] && !names[c]) names[c] = String(v[r][2]);
  }
  const p = PropertiesService.getScriptProperties();
  return { rows: Object.keys(by).map(function (k) { return { code: k, data: JSON.stringify(by[k].d) }; }),
    mail: MAIL, names: names, showcase: p.getProperty('showcase') || '', types: p.getProperty('types') || '',
    badges: p.getProperty('badges') || '', sheetUrl: SpreadsheetApp.getActiveSpreadsheet().getUrl() };
}
function setShowcase(pin, s) { if (!pinOk_(pin)) throw new Error('no'); PropertiesService.getScriptProperties().setProperty('showcase', String(s)); return true; }
function setType(pin, key, val) {
  if (!pinOk_(pin)) throw new Error('no');
  const p = PropertiesService.getScriptProperties(), t = JSON.parse(p.getProperty('types') || '{}');
  if (val) t[String(key)] = String(val); else delete t[String(key)];
  p.setProperty('types', JSON.stringify(t)); return true;
}
function setBadges(pin, j) { if (!pinOk_(pin)) throw new Error('no'); JSON.parse(j); PropertiesService.getScriptProperties().setProperty('badges', String(j)); return true; }
function getBadges(code) { return JSON.parse(PropertiesService.getScriptProperties().getProperty('badges') || '{}')[String(code)] || []; }

/** 教師端：清除某一班的讀報資料（需密碼；作答紀錄工作表保留） */
function clearClass(pin, cls) {
  if (!pinOk_(pin)) throw new Error('密碼錯誤');
  if (!/^\d{3}$/.test(String(cls))) throw new Error('班級格式不正確');
  const sh = sheet_(DATA, DATA_HEAD), v = sh.getDataRange().getValues();
  let n = 0;
  for (let r = v.length - 1; r >= 1; r--) if (String(v[r][1]).indexOf(cls + '-') === 0) { sh.deleteRow(r + 1); n++; }
  return n;
}

/** 版面圖：由老師帳號讀雲端硬碟（資料夾不用分享） */
function getImage(name) {
  if (!/^[\w-]+$/.test(String(name))) return '';
  const folders = DriveApp.getFoldersByName(IMG_FOLDER);
  while (folders.hasNext()) {
    const files = folders.next().getFilesByName(name + '.jpg');
    if (files.hasNext()) { const b = files.next().getBlob(); return 'data:' + b.getContentType() + ';base64,' + Utilities.base64Encode(b.getBytes()); }
  }
  return '';
}
function testImage() { const u = getImage('894-nobel'); Logger.log(u ? '成功，圖片大小：' + u.length : '找不到圖片'); }

/** 學生姓名（只回傳給教師端）：完整版請用老師私人備份中的 Code.gs */
const NAMES = {};

/** 舊版信箱資料：完整版請用老師私人備份中的 Code.gs */
const MAIL = [];
