// 侯氏韧性评估 · 班组互评 —— Google 表格后端
// 用法：在 Google 表格里打开 扩展程序 → Apps Script，把本文件内容全部粘贴进去替换默认代码，
// 然后 部署 → 新建部署 → 类型选“Web 应用” → 执行身份“我” → 访问权限“任何人” → 部署，复制以 /exec 结尾的网址。

const SHEET_NAME = 'submissions';

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(['time', 'group', 'rater', 'raterIdx', 'names', 'blocks', 'code']);
  }
  return sh;
}

// 学员页提交：POST 一段 JSON
function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const d = JSON.parse(e.postData.contents || '{}');
    if (!d.group || !d.rater || !Array.isArray(d.blocks) || !Array.isArray(d.names)) {
      return out_({ ok: false, error: 'bad payload' });
    }
    sheet_().appendRow([
      new Date(), String(d.group), String(d.rater), Number(d.raterIdx),
      JSON.stringify(d.names), JSON.stringify(d.blocks), String(d.code || '')
    ]);
    return out_({ ok: true });
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (e2) {}
  }
}

// 后台读取：GET ?group=xxx（不带 group 则返回全部）
function doGet(e) {
  const g = (e && e.parameter && e.parameter.group) || '';
  const rows = sheet_().getDataRange().getValues().slice(1);
  const list = rows
    .filter(r => r[1] && (!g || String(r[1]) === g))
    .map(r => ({
      time: r[0] instanceof Date ? r[0].getTime() : r[0],
      group: String(r[1]), rater: String(r[2]), raterIdx: Number(r[3]),
      names: safe_(r[4]), blocks: safe_(r[5]), code: String(r[6] || '')
    }));
  return out_({ ok: true, count: list.length, submissions: list });
}

function safe_(s) { try { return JSON.parse(s); } catch (e) { return []; } }
function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
