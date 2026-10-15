/**
 * 清掃費の請求 — Googleスプレッドシート側のプログラム
 *
 * スタッフは共通URLの画面で、名前・振込先・勤務日を毎回入力して送信します。
 * 送信された請求はこのスプレッドシートに書き込まれるだけで、スタッフ側には何も表示されません。
 * 一覧・支払管理はオーナー用の管理画面（またはこのシート）で行います。
 */

const SHEET = { inv: '請求書', lines: '明細', config: '設定' };
const HEAD = {
  inv:    ['受付番号', '受付日時', '名前', '銀行名', '支店名', '口座種別', '口座番号', '口座名義', '対象月', '請求日',
           '清掃費', '駐車場代', 'その他', '合計', '日数', '備考', '支払', '支払日'],
  lines:  ['受付番号', '勤務日', '物件', '清掃費', '駐車場代', 'その他', 'その他内容', '小計'],
  config: ['項目', '値'],
};
const TZ = 'Asia/Tokyo';
const PROPERTIES = ['Grand House（睦町）', 'Good Tiny House（中村町）', 'その他'];
const MAX_LINES = 40;

/* ================= メニュー ================= */

function onOpen() {
  SpreadsheetApp.getUi().createMenu('清掃請求')
    .addItem('初期設定（最初と更新後に1回）', 'setup')
    .addItem('スタッフ用のURLを表示', 'showStaffUrl')
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
  text('inv', [1, 7, 9, 10]);
  text('lines', [1, 2]);
  ss.getSheetByName(SHEET.inv).getRange('Q2:Q').setDataValidation(SpreadsheetApp.newDataValidation().requireCheckbox().build());

  const cfg = ss.getSheetByName(SHEET.config);
  const has = key => cfg.getLastRow() > 1 && cfg.getRange(2, 1, cfg.getLastRow() - 1, 1).getValues().some(r => r[0] === key);
  if (!has('宛名')) cfg.appendRow(['宛名', '']);
  if (!has('管理者トークン')) cfg.appendRow(['管理者トークン', newToken_()]);
  if (!has('ページURL')) cfg.appendRow(['ページURL', '']);
  cfg.getRange('B2:B4').setNumberFormat('@');

  notify_('初期設定が終わりました。\n\n' +
    '「請求書」「明細」シートに、スタッフから届いた請求が入ります。\n' +
    '（以前の「スタッフ」「記録」「請求」シートはもう使いません。不要なら削除して大丈夫です）');
}

function showStaffUrl() {
  SpreadsheetApp.getUi().alert('スタッフ全員共通のURLです。LINEで送ってください。\n\n' + pageUrl_());
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
  /* ---- スタッフ（書き込みだけ。誰の情報も返さない） ---- */
  payee() {
    return { payee: config_('宛名') };
  },
  submit(req) {
    const v = invoiceData_(req.invoice);
    return withLock_(() => {
      const id = Utilities.formatDate(new Date(), TZ, 'yyMMdd') + '-' + String(Math.floor(Math.random() * 9000) + 1000);
      const b = v.bank;
      sheet_('inv').appendRow([id, now_(), v.name, b.bank, b.branch, b.type, b.number, b.holder, v.month, v.date,
        v.sum.cleaning, v.sum.parking, v.sum.other, v.total, v.lines.length, v.note, false, '']);
      const rows = v.lines.map(l => [id, l.date, l.property, l.cleaning, l.parking, l.otherAmt, l.otherNote, l.total]);
      const sh = sheet_('lines');
      sh.getRange(sh.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
      return { id: id };
    });
  },

  /* ---- 管理者 ---- */
  adminList(req) {
    requireAdmin_(req.admin);
    const lines = {};
    table_('lines').rows.forEach(r => {
      (lines[r[0]] = lines[r[0]] || []).push({
        date: fmt_(r[1]), property: r[2], cleaning: Number(r[3]) || 0, parking: Number(r[4]) || 0,
        otherAmt: Number(r[5]) || 0, otherNote: r[6],
      });
    });
    return {
      invoices: table_('inv').rows.map(r => ({
        id: r[0], receivedAt: r[1], name: r[2],
        bank: { bank: r[3], branch: r[4], type: r[5], number: String(r[6]), holder: r[7] },
        month: fmt_(r[8]).slice(0, 7), date: fmt_(r[9]), total: Number(r[13]) || 0, note: r[15],
        paid: isTrue_(r[16]), paidAt: fmt_(r[17]), lines: (lines[r[0]] || []).sort((a, b) => a.date < b.date ? -1 : 1),
      })),
      payee: config_('宛名'),
    };
  },
  setPaid(req) {
    requireAdmin_(req.admin);
    return withLock_(() => {
      const ids = new Set([].concat(req.ids || []));
      const paid = !!req.paid;
      const t = table_('inv');
      t.rows.forEach(r => { if (ids.has(r[0])) t.sh.getRange(r._row, 17, 1, 2).setValues([[paid, paid ? fmt_(new Date()) : '']]); });
      return {};
    });
  },
  deleteInvoice(req) {
    requireAdmin_(req.admin);
    return withLock_(() => {
      const id = String(req.id || '');
      ['lines', 'inv'].forEach(key => {
        const t = table_(key);
        t.rows.filter(r => r[0] === id).map(r => r._row).sort((a, b) => b - a).forEach(row => t.sh.deleteRow(row));
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
  const vals = n > 0 ? sh.getRange(2, 1, n, HEAD[key].length).getValues() : [];
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
function newToken_() { return Utilities.getUuid().replace(/-/g, '').slice(0, 24); }
function now_() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss'); }
function isTrue_(v) { return v === true || String(v).toUpperCase() === 'TRUE'; }
function str_(v, max) { return String(v == null ? '' : v).trim().slice(0, max); }
function int_(v) { const n = Math.round(Number(v)); return isFinite(n) && n > 0 ? Math.min(n, 10000000) : 0; }
function date_(v, label) { const s = String(v || ''); if (!/^\d{4}-\d{2}-\d{2}$/.test(s)) throw new Error((label || '日付') + 'を入れてください'); return s; }
function month_(v) { const s = String(v || ''); if (!/^\d{4}-\d{2}$/.test(s)) throw new Error('何月分かを選んでください'); return s; }

function requireAdmin_(admin) {
  const a = String(admin || '');
  if (a.length < 10 || a !== config_('管理者トークン')) throw new Error('管理者用のURLが正しくありません');
}
function invoiceData_(x) {
  x = x || {};
  const p = x.bank || {};
  const name = str_(x.name, 40);
  const bank = { bank: str_(p.bank, 40), branch: str_(p.branch, 40), type: p.type === '当座' ? '当座' : '普通',
                 number: String(p.number || '').replace(/\D/g, ''), holder: str_(p.holder, 40) };
  if (!name) throw new Error('お名前を入れてください');
  if (!bank.bank || !bank.branch || !bank.holder) throw new Error('振込先をすべて入れてください');
  if (!/^\d{7,8}$/.test(bank.number)) throw new Error('口座番号は7桁の数字で入れてください');
  const raw = [].concat(x.lines || []);
  if (!raw.length) throw new Error('勤務日を1日以上入れてください');
  if (raw.length > MAX_LINES) throw new Error('勤務日は1回に' + MAX_LINES + '日までです');
  const lines = raw.map((l, i) => {
    l = l || {};
    const d = {
      date: date_(l.date, (i + 1) + '行目の勤務日'),
      property: PROPERTIES.indexOf(l.property) >= 0 ? l.property : 'その他',
      cleaning: int_(l.cleaning), parking: int_(l.parking), otherAmt: int_(l.otherAmt), otherNote: str_(l.otherNote, 60),
    };
    d.total = d.cleaning + d.parking + d.otherAmt;
    if (!d.total) throw new Error((i + 1) + '行目の金額を入れてください');
    if (d.otherAmt && !d.otherNote) throw new Error((i + 1) + '行目の「その他の内容」を書いてください');
    return d;
  }).sort((a, b) => a.date < b.date ? -1 : 1);
  const sum = lines.reduce((s, l) => ({ cleaning: s.cleaning + l.cleaning, parking: s.parking + l.parking, other: s.other + l.otherAmt }),
    { cleaning: 0, parking: 0, other: 0 });
  return { name: name, bank: bank, month: month_(x.month), date: date_(x.date, '請求日'), note: str_(x.note, 300),
           lines: lines, sum: sum, total: sum.cleaning + sum.parking + sum.other };
}
function withLock_(fn) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) throw new Error('混み合っています。少し待ってもう一度お試しください');
  try { return fn(); } finally { lock.releaseLock(); }
}
