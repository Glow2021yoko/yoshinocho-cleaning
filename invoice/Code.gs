/**
 * 清掃費の請求 — Googleスプレッドシート側のプログラム
 *
 * このスプレッドシートが保存先になります。
 * スタッフは GitHub Pages の画面から入力し、内容はここに書き込まれます。
 * 使い方は README.md を見てください。
 */

const SHEET = { staff: 'スタッフ', items: '記録', bills: '請求', config: '設定' };
const HEAD = {
  staff:  ['トークン', '名前', '銀行名', '支店名', '口座種別', '口座番号', '口座名義', '有効', '登録日時'],
  items:  ['ID', 'トークン', '勤務日', '物件', '清掃費', '駐車場代', 'その他', 'その他内容', 'メモ', '合計', '支払', '支払日', '登録日時', '更新日時'],
  bills:  ['トークン', '月', '請求日', '更新日時'],
  config: ['項目', '値'],
};
const TZ = 'Asia/Tokyo';
const PROPERTIES = ['Grand House（睦町）', 'Good Tiny House（中村町）', 'その他'];

/* ================= メニュー ================= */

function onOpen() {
  SpreadsheetApp.getUi().createMenu('清掃請求')
    .addItem('① 初期設定（最初に1回）', 'setup')
    .addItem('スタッフを追加してURLを発行', 'addStaff')
    .addItem('スタッフのURLを再表示', 'showStaffUrl')
    .addItem('管理画面のURLを表示', 'showAdminUrl')
    .addToUi();
}

function setup() {
  const ss = SpreadsheetApp.getActive();
  Object.keys(SHEET).forEach(key => {
    const sh = ss.getSheetByName(SHEET[key]) || ss.insertSheet(SHEET[key]);
    if (sh.getLastRow() === 0) sh.appendRow(HEAD[key]);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, HEAD[key].length).setFontWeight('bold').setBackground('#e9f0ef');
  });
  // 先頭の0が消えたり日付に変換されたりしないよう、文字列として扱う列
  const text = (key, cols) => {
    const sh = ss.getSheetByName(SHEET[key]);
    cols.forEach(c => sh.getRange(2, c, sh.getMaxRows() - 1, 1).setNumberFormat('@'));
  };
  text('staff', [1, 6]);
  text('items', [1, 2, 3]);
  text('bills', [1, 2, 3]);
  // チェックボックス（有効・支払）
  const cb = SpreadsheetApp.newDataValidation().requireCheckbox().build();
  ss.getSheetByName(SHEET.staff).getRange('H2:H').setDataValidation(cb);
  ss.getSheetByName(SHEET.items).getRange('K2:K').setDataValidation(cb);

  const cfg = ss.getSheetByName(SHEET.config);
  if (cfg.getLastRow() < 2) {
    cfg.appendRow(['宛名', '']);
    cfg.appendRow(['管理者トークン', newToken_()]);
    cfg.appendRow(['ページURL', '']);
    cfg.getRange('B2:B4').setNumberFormat('@');
  }
  notify_(
    '初期設定が終わりました。\n\n' +
    '「設定」シートの「宛名」に請求書の宛名（例：〇〇合同会社 御中）を、\n' +
    '「ページURL」に GitHub Pages のURLを入れてください。\n\n' +
    'そのあと「デプロイ → 新しいデプロイ → ウェブアプリ」で公開します（README参照）。');
}

function addStaff() {
  const ui = SpreadsheetApp.getUi();
  const r = ui.prompt('スタッフを追加', 'スタッフの名前を入れてください（あとで本人が修正できます）', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  const name = r.getResponseText().trim();
  const t = newToken_();
  const sh = sheet_('staff');
  sh.appendRow([t, name, '', '', '普通', '', '', true, now_()]);
  ui.alert('「' + name + '」さん専用のURLです。LINEで送ってください。\n\n' + staffUrl_(t) +
    '\n\n※このURLを知っている人は、この人として入力できます。他の人に転送しないよう伝えてください。');
}

function showStaffUrl() {
  const ui = SpreadsheetApp.getUi();
  const list = table_('staff').rows.map(r => '・' + (r[1] || '(名前なし)') + '\n' + staffUrl_(r[0])).join('\n\n');
  ui.alert(list || 'まだスタッフがいません。');
}

function showAdminUrl() {
  SpreadsheetApp.getUi().alert('管理画面のURLです（リサさん専用、他の人に渡さないでください）\n\n' +
    pageUrl_() + '?admin=' + config_('管理者トークン'));
}

/* ================= ウェブAPI ================= */

function doGet() {
  return ContentService.createTextOutput('清掃費の請求 API は動いています。');
}

function doPost(e) {
  let out;
  try {
    const req = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const fn = ACTIONS[req.action];
    if (!fn) throw new Error('不明な操作です');
    out = Object.assign({ ok: true }, fn(req) || {});
  } catch (err) {
    out = { ok: false, error: String((err && err.message) || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

const ACTIONS = {
  /* ---- スタッフ ---- */
  me(req) {
    const s = requireStaff_(req.token);
    return {
      profile: staffObj_(s),
      items: table_('items').rows.filter(r => r[1] === s.values[0]).map(itemObj_),
      billed: billedMap_(s.values[0]),
      payee: config_('宛名'),
    };
  },
  saveProfile(req) {
    return withLock_(() => {
      const s = requireStaff_(req.token);
      const p = req.profile || {};
      const name = str_(p.name, 40), bank = str_(p.bank, 40), branch = str_(p.branch, 40), holder = str_(p.holder, 40);
      const type = p.type === '当座' ? '当座' : '普通';
      const number = String(p.number || '').replace(/\D/g, '');
      if (!name || !bank || !branch || !holder) throw new Error('すべての項目を入れてください');
      if (!/^\d{7,8}$/.test(number)) throw new Error('口座番号は7桁の数字で入れてください');
      sheet_('staff').getRange(s.row, 2, 1, 6).setValues([[name, bank, branch, type, number, holder]]);
      return {};
    });
  },
  addItem(req) {
    return withLock_(() => {
      const s = requireStaff_(req.token);
      if (!s.values[2] || !s.values[5]) throw new Error('先にお名前と振込先を保存してください');
      const d = itemData_(req.item);
      const id = Utilities.getUuid().replace(/-/g, '').slice(0, 16);
      const ts = now_();
      sheet_('items').appendRow([id, s.values[0], d.date, d.property, d.cleaning, d.parking, d.otherAmt, d.otherNote, d.memo, d.total, false, '', ts, ts]);
      return { id: id };
    });
  },
  updateItem(req) {
    return withLock_(() => {
      const s = requireStaff_(req.token);
      const it = ownUnpaidItem_(s, req.id);
      const d = itemData_(req.item);
      sheet_('items').getRange(it.row, 3, 1, 8).setValues([[d.date, d.property, d.cleaning, d.parking, d.otherAmt, d.otherNote, d.memo, d.total]]);
      sheet_('items').getRange(it.row, 14).setValue(now_());
      return {};
    });
  },
  deleteItem(req) {
    return withLock_(() => {
      const s = requireStaff_(req.token);
      const it = ownUnpaidItem_(s, req.id);
      sheet_('items').deleteRow(it.row);
      return {};
    });
  },
  bill(req) {
    return withLock_(() => {
      const s = requireStaff_(req.token);
      const month = month_(req.month), date = date_(req.date);
      const t = table_('bills');
      const found = t.rows.find(r => r[0] === s.values[0] && r[1] === month);
      if (found) t.sh.getRange(found._row, 3, 1, 2).setValues([[date, now_()]]);
      else t.sh.appendRow([s.values[0], month, date, now_()]);
      return {};
    });
  },
  unbill(req) {
    return withLock_(() => {
      const s = requireStaff_(req.token);
      const month = month_(req.month);
      const paid = table_('items').rows.some(r => r[1] === s.values[0] && String(r[2]).startsWith(month) && isTrue_(r[10]));
      if (paid) throw new Error('支払済の記録があるため取り消せません');
      const t = table_('bills');
      const found = t.rows.find(r => r[0] === s.values[0] && r[1] === month);
      if (found) t.sh.deleteRow(found._row);
      return {};
    });
  },

  /* ---- 管理者 ---- */
  adminList(req) {
    requireAdmin_(req.admin);
    return {
      staff: table_('staff').rows.map(r => Object.assign({ token: r[0] }, staffObj_({ values: r }))),
      items: table_('items').rows.map(r => Object.assign({ token: r[1] }, itemObj_(r))),
      bills: table_('bills').rows.map(r => ({ token: r[0], month: fmt_(r[1]).slice(0, 7), date: fmt_(r[2]) })),
      payee: config_('宛名'),
    };
  },
  setPaid(req) {
    requireAdmin_(req.admin);
    return withLock_(() => {
      const ids = new Set([].concat(req.ids || []));
      const paid = !!req.paid;
      const t = table_('items');
      t.rows.forEach(r => {
        if (ids.has(r[0])) t.sh.getRange(r._row, 11, 1, 2).setValues([[paid, paid ? fmt_(new Date()) : '']]);
      });
      return {};
    });
  },
  setPayee(req) {
    requireAdmin_(req.admin);
    return withLock_(() => { setConfig_('宛名', str_(req.payee, 60)); return {}; });
  },
};

/* ================= 内部の処理 ================= */

// スプレッドシートの画面から実行したときはダイアログ、エディタから実行したときは実行ログに表示
function notify_(msg) {
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) { Logger.log(msg); }
}

function sheet_(key) {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET[key]);
  if (!sh) throw new Error('シート「' + SHEET[key] + '」がありません。メニューの「初期設定」を実行してください');
  return sh;
}
function table_(key) {
  const sh = sheet_(key);
  const n = sh.getLastRow() - 1;
  const width = HEAD[key].length;
  const vals = n > 0 ? sh.getRange(2, 1, n, width).getValues() : [];
  const rows = vals.map((v, i) => { const r = v.map(fmt_); r._row = i + 2; return r; }).filter(r => r[0] !== '');
  return { sh: sh, rows: rows };
}
function fmt_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, TZ, 'yyyy-MM-dd');
  if (typeof v === 'boolean' || typeof v === 'number') return v;
  return String(v);
}
function config_(key) {
  const r = table_('config').rows.find(r => r[0] === key);
  return r ? String(r[1]) : '';
}
function setConfig_(key, value) {
  const t = table_('config');
  const r = t.rows.find(r => r[0] === key);
  if (r) t.sh.getRange(r._row, 2).setValue(value); else t.sh.appendRow([key, value]);
}
function pageUrl_() {
  const u = config_('ページURL').replace(/\/?$/, '/');
  if (u === '/') throw new Error('「設定」シートの「ページURL」に GitHub Pages のURLを入れてください');
  return u;
}
function staffUrl_(t) { return pageUrl_() + '?t=' + t; }
function newToken_() { return Utilities.getUuid().replace(/-/g, '').slice(0, 24); }
function now_() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss'); }
function isTrue_(v) { return v === true || String(v).toUpperCase() === 'TRUE'; }
function str_(v, max) { return String(v == null ? '' : v).trim().slice(0, max); }
function int_(v) { const n = Math.round(Number(v)); return isFinite(n) && n > 0 ? n : 0; }
function date_(v) { const s = String(v || ''); if (!/^\d{4}-\d{2}-\d{2}$/.test(s)) throw new Error('日付の形式が正しくありません'); return s; }
function month_(v) { const s = String(v || ''); if (!/^\d{4}-\d{2}$/.test(s)) throw new Error('月の形式が正しくありません'); return s; }

function requireStaff_(token) {
  const t = String(token || '');
  if (t.length < 10) throw new Error('このURLは使えません。オーナーに専用URLをもらってください');
  const r = table_('staff').rows.find(r => r[0] === t);
  if (!r || !isTrue_(r[7])) throw new Error('このURLは使えません。オーナーに専用URLをもらってください');
  return { row: r._row, values: r };
}
function requireAdmin_(admin) {
  const a = String(admin || '');
  if (a.length < 10 || a !== config_('管理者トークン')) throw new Error('管理者用のURLが正しくありません');
}
function staffObj_(s) {
  const v = s.values;
  return { name: v[1], bank: v[2] ? { bank: v[2], branch: v[3], type: v[4] || '普通', number: String(v[5]), holder: v[6] } : null };
}
function itemObj_(r) {
  return {
    id: r[0], date: fmt_(r[2]), property: r[3], cleaning: Number(r[4]) || 0, parking: Number(r[5]) || 0,
    otherAmt: Number(r[6]) || 0, otherNote: r[7], memo: r[8], paid: isTrue_(r[10]), paidAt: fmt_(r[11]), createdAt: r[12],
  };
}
function billedMap_(token) {
  const m = {};
  table_('bills').rows.filter(r => r[0] === token).forEach(r => { m[fmt_(r[1]).slice(0, 7)] = fmt_(r[2]); });
  return m;
}
function itemData_(it) {
  it = it || {};
  const d = {
    date: date_(it.date),
    property: PROPERTIES.indexOf(it.property) >= 0 ? it.property : 'その他',
    cleaning: int_(it.cleaning), parking: int_(it.parking), otherAmt: int_(it.otherAmt),
    otherNote: str_(it.otherNote, 60), memo: str_(it.memo, 300),
  };
  d.total = d.cleaning + d.parking + d.otherAmt;
  if (!d.total) throw new Error('金額を1つ以上入れてください');
  if (d.otherAmt && !d.otherNote) throw new Error('その他の内容を書いてください');
  return d;
}
function ownUnpaidItem_(s, id) {
  const r = table_('items').rows.find(r => r[0] === String(id || ''));
  if (!r || r[1] !== s.values[0]) throw new Error('記録が見つかりません');
  if (isTrue_(r[10])) throw new Error('支払済の記録は変更できません');
  return { row: r._row, values: r };
}
function withLock_(fn) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) throw new Error('混み合っています。少し待ってもう一度お試しください');
  try { return fn(); } finally { lock.releaseLock(); }
}
