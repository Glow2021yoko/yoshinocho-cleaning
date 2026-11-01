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
           '清掃費', '駐車場代', 'その他', '合計', '日数', '備考', '支払', '支払日', '請求書PDF'],
  lines:  ['受付番号', '名前', '勤務日', '物件', '清掃費', '駐車場代', 'その他', 'その他内容', '小計'],
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
    .addItem('PDFがない請求書のPDFを作る', 'makeMissingPdfs')
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
  migrate_();
  // 先頭の0が消えたり日付に変換されたりしないよう、文字列として扱う列
  const text = (key, cols) => {
    const sh = ss.getSheetByName(SHEET[key]);
    cols.forEach(c => sh.getRange(2, c, sh.getMaxRows() - 1, 1).setNumberFormat('@'));
  };
  text('inv', [1, 7, 9, 10]);
  text('lines', [1, 3]);
  ss.getSheetByName(SHEET.inv).getRange('Q2:Q').setDataValidation(SpreadsheetApp.newDataValidation().requireCheckbox().build());

  const cfg = ss.getSheetByName(SHEET.config);
  const has = key => cfg.getLastRow() > 1 && cfg.getRange(2, 1, cfg.getLastRow() - 1, 1).getValues().some(r => r[0] === key);
  if (!has('宛名')) cfg.appendRow(['宛名', '']);
  if (!has('管理者トークン')) cfg.appendRow(['管理者トークン', newToken_()]);
  if (!has('ページURL')) cfg.appendRow(['ページURL', '']);
  if (!has('通知メール')) cfg.appendRow(['通知メール', '']);
  if (!has('PDF保存フォルダID')) cfg.appendRow(['PDF保存フォルダID', '']);
  cfg.getRange('B2:B6').setNumberFormat('@');
  pdfFolder_();  // 保存先フォルダを用意（ここでドライブの許可を求められます）
  fixMonths_();
  setupPaySheet_();

  notify_('初期設定が終わりました。\n\n' +
    '「請求書」「明細」シートに、スタッフから届いた請求が入ります。\n' +
    '請求書PDFはドライブの「清掃費の請求書」フォルダに保存されます。\n' +
    '「支払い」シートに、月・支払先ごとの合計金額が自動で集計されます。\n' +
    '「設定」シートの「通知メール」にアドレスを入れると、届くたびにPDF付きのメールが届きます。\n\n' +
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
      migrate_();
      const id = Utilities.formatDate(new Date(), TZ, 'yyMMdd') + '-' + String(Math.floor(Math.random() * 9000) + 1000);
      const b = v.bank;
      sheet_('inv').appendRow([id, now_(), v.name, b.bank, b.branch, b.type, b.number, b.holder, v.month, v.date,
        v.sum.cleaning, v.sum.parking, v.sum.other, v.total, v.lines.length, v.note, false, '']);
      const rows = v.lines.map(l => [id, v.name, l.date, l.property, l.cleaning, l.parking, l.otherAmt, l.otherNote, l.total]);
      const sh = sheet_('lines');
      sh.getRange(sh.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
      afterSubmit_(id);   // PDF作成・メール通知（失敗しても請求自体は受け付ける）
      return { id: id };
    });
  },

  /* ---- 管理者 ---- */
  adminList(req) {
    requireAdmin_(req.admin);
    migrate_();
    const lines = {};
    table_('lines').rows.forEach(r => {
      (lines[r[0]] = lines[r[0]] || []).push({
        date: fmt_(r[2]), property: r[3], cleaning: Number(r[4]) || 0, parking: Number(r[5]) || 0,
        otherAmt: Number(r[6]) || 0, otherNote: r[7],
      });
    });
    return {
      invoices: table_('inv').rows.map(r => ({
        id: r[0], receivedAt: r[1], name: r[2],
        bank: { bank: r[3], branch: r[4], type: r[5], number: String(r[6]), holder: r[7] },
        month: fmt_(r[8]).slice(0, 7), date: fmt_(r[9]), total: Number(r[13]) || 0, note: r[15],
        paid: isTrue_(r[16]), paidAt: fmt_(r[17]), pdf: String(r[18] || ''), lines: (lines[r[0]] || []).sort((a, b) => a.date < b.date ? -1 : 1),
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

/* ================= 支払いシート ================= */

// 「請求書」シートから、月×支払先（名前＋口座）ごとの合計を自動集計するシート。数式なので常に最新。
function setupPaySheet_() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('支払い') || ss.insertSheet('支払い', 0);
  sh.clear();
  const q = "select I, C, D, E, F, G, H, sum(N), sum(O) where A is not null and C is not null " +
            "group by I, C, D, E, F, G, H order by I desc, C " +
            "label I '勤務月', C '支払先', D '銀行名', E '支店名', F '種別', G '口座番号', H '口座名義', sum(N) '合計金額', sum(O) '勤務日数'";
  sh.getRange('A1').setValue('支払月');
  sh.getRange('A2').setFormula('=ARRAYFORMULA(IF(B2:B="", "", TEXT(EDATE(DATEVALUE(B2:B&"-01"), 1), "yyyy-mm")))');
  sh.getRange('B1').setFormula('=QUERY(\'請求書\'!A:Q, "' + q + '", 1)');
  sh.getRange('K1:L1').setValues([['未払い金額', '状態']]);
  sh.getRange('K2').setFormula("=ARRAYFORMULA(IF(B2:B=\"\", \"\", SUMIFS('請求書'!N2:N, '請求書'!I2:I, B2:B, '請求書'!C2:C, C2:C, '請求書'!Q2:Q, FALSE)))");
  sh.getRange('L2').setFormula("=ARRAYFORMULA(IF(B2:B=\"\", \"\", IF(K2:K=0, \"支払済\", \"未払い\")))");
  sh.getRange('I:I').setNumberFormat('¥#,##0');
  sh.getRange('K:K').setNumberFormat('¥#,##0');
  sh.setFrozenRows(1);
  sh.getRange('A1:L1').setFontWeight('bold').setBackground('#e9f0ef');
  sh.getRange('N1').setValue('※自動で集計されます（月末締め・翌月払い）。直接書き込まないでください。支払済にするときは「請求書」シートの「支払」にチェックを入れるか、管理画面で切り替えてください。')
    .setFontColor('#666666');
}

// 届いている請求の「対象月（＝勤務月）」を、明細の勤務日から付け直す
function fixMonths_() {
  const byId = {};
  table_('lines').rows.forEach(r => { const m = fmt_(r[2]).slice(0, 7); (byId[r[0]] = byId[r[0]] || {})[m] = ((byId[r[0]] || {})[m] || 0) + 1; });
  const t = table_('inv');
  t.rows.forEach(r => {
    const counts = byId[r[0]]; if (!counts) return;
    const m = Object.keys(counts).sort((a, b) => counts[b] - counts[a] || (a < b ? -1 : 1))[0];  // いちばん多い月
    if (fmt_(r[8]).slice(0, 7) !== m) t.sh.getRange(r._row, 9).setValue(m);
  });
}

/* ================= 請求書PDF・メール ================= */

function makeMissingPdfs() {
  migrate_();
  const t = table_('inv');
  let n = 0;
  t.rows.forEach(r => { if (!/^https?:/.test(String(r[18]))) { makePdfFor_(r[0], t); n++; } });
  notify_(n ? n + '件の請求書PDFを作りました。「請求書」シートの「請求書PDF」列から開けます。' : 'PDFがない請求書はありません。');
}

function afterSubmit_(id) {
  let file = null;
  try { file = makePdfFor_(id); } catch (e) { Logger.log(e); }
  const to = config_('通知メール').trim();
  if (!to) return;
  try {
    const inv = loadInvoice_(id);
    const body = inv.name + ' さんから ' + jpMonth_(inv.month) + '分の清掃費の請求が届きました。\n\n' +
      '受付番号：' + inv.id + '\n請求日：' + inv.date.replace(/-/g, '/') + '\nご請求金額：' + yen_(inv.total) + '（' + inv.lines.length + '日）\n' +
      '振込先：' + [inv.bank.bank, inv.bank.branch, inv.bank.type, inv.bank.number, inv.bank.holder].join(' ') + '\n\n' +
      (file ? '請求書PDF：' + file.getUrl() + '\n' : '') +
      '管理画面：' + pageUrl_() + '?admin=' + config_('管理者トークン');
    MailApp.sendEmail({ to: to, subject: '【清掃費の請求】' + inv.name + ' ' + jpMonth_(inv.month) + '分 ' + yen_(inv.total), body: body,
      attachments: file ? [file.getAs(MimeType.PDF)] : [] });
  } catch (e) { Logger.log(e); }
}

function makePdfFor_(id, t) {
  const inv = loadInvoice_(id);
  const blob = Utilities.newBlob(invoiceHtml_(inv), MimeType.HTML, 'invoice.html').getAs(MimeType.PDF)
    .setName('請求書_' + inv.name.replace(/\s+/g, '') + '_' + jpMonth_(inv.month) + '_' + inv.id + '.pdf');
  const file = pdfFolder_().createFile(blob);
  t = t || table_('inv');
  const r = t.rows.find(x => x[0] === id);
  if (r) t.sh.getRange(r._row, 19).setValue(file.getUrl());
  return file;
}

function loadInvoice_(id) {
  const r = table_('inv').rows.find(x => x[0] === id);
  if (!r) throw new Error('請求が見つかりません: ' + id);
  const lines = table_('lines').rows.filter(x => x[0] === id).map(x => ({
    date: fmt_(x[2]), property: x[3], cleaning: Number(x[4]) || 0, parking: Number(x[5]) || 0, otherAmt: Number(x[6]) || 0, otherNote: x[7],
  })).sort((a, b) => a.date < b.date ? -1 : 1);
  return { id: r[0], name: r[2], bank: { bank: r[3], branch: r[4], type: r[5], number: String(r[6]), holder: r[7] },
           month: fmt_(r[8]).slice(0, 7), date: fmt_(r[9]), note: r[15], lines: lines,
           total: lines.reduce((s, l) => s + l.cleaning + l.parking + l.otherAmt, 0) };
}

function pdfFolder_() {
  const id = config_('PDF保存フォルダID').trim();
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) {} }
  const ss = DriveApp.getFileById(SpreadsheetApp.getActive().getId());
  const parents = ss.getParents();
  const base = parents.hasNext() ? parents.next() : DriveApp.getRootFolder();
  const folder = base.createFolder('清掃費の請求書');
  setConfig_('PDF保存フォルダID', folder.getId());
  return folder;
}

function invoiceHtml_(inv) {
  const e = s => String(s == null ? '' : s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
  const md = d => { const p = d.split('-'); return Number(p[1]) + '/' + Number(p[2]); };
  const sum = k => inv.lines.reduce((s, l) => s + (Number(l[k]) || 0), 0);
  const b = inv.bank;
  const rows = inv.lines.map(l => '<tr><td class="c">' + md(l.date) + '</td><td>' + e(l.property) + '</td>' +
    '<td class="n">' + (l.cleaning ? yen_(l.cleaning) : '') + '</td><td class="n">' + (l.parking ? yen_(l.parking) : '') + '</td>' +
    '<td class="n">' + (l.otherAmt ? yen_(l.otherAmt) : '') + (l.otherNote ? '<br><span class="s">' + e(l.otherNote) + '</span>' : '') + '</td>' +
    '<td class="n">' + yen_(l.cleaning + l.parking + l.otherAmt) + '</td></tr>').join('');
  return '<html><head><meta charset="utf-8"><style>' +
    '@page{size:A4;margin:18mm 16mm}' +
    'body{font-family:"Noto Sans JP","Hiragino Kaku Gothic ProN",sans-serif;color:#1a1a1a;font-size:10.5pt;line-height:1.6}' +
    'h1{text-align:center;font-size:22pt;letter-spacing:12pt;margin:0 0 18pt}' +
    '.meta{text-align:right;font-size:9.5pt;color:#444}' +
    '.top{width:100%;border-collapse:collapse;margin:14pt 0}.top td{vertical-align:bottom;border:none;padding:0}' +
    '.to{font-size:15pt;border-bottom:1px solid #1a1a1a;padding-bottom:3pt}' +
    '.from{text-align:right;font-size:12pt;font-weight:bold}' +
    '.amt{border:2px solid #1f5f5b;padding:8pt 14pt;display:inline-block;margin:6pt 0 4pt}.amt b{font-size:18pt;margin-left:14pt}' +
    'table.d{width:100%;border-collapse:collapse;margin-top:16pt;font-size:9.5pt}' +
    'table.d th{background:#e9f0ef;border:1px solid #c9d3d1;padding:5pt}table.d td{border:1px solid #d6dcdb;padding:5pt}' +
    '.n{text-align:right;white-space:nowrap}.c{text-align:center;white-space:nowrap}.s{color:#666;font-size:8.5pt}' +
    'tr.t td{font-weight:bold;background:#f6f8f8}' +
    '.bank{border:1px solid #c9d3d1;padding:8pt 12pt;margin-top:16pt}.foot{color:#666;font-size:9pt;margin-top:10pt}' +
    '</style></head><body>' +
    '<h1>請求書</h1>' +
    '<div class="meta">受付番号 ' + e(inv.id) + '　　請求日 ' + e(inv.date.replace(/-/g, '/')) + '</div>' +
    '<table class="top"><tr><td style="width:55%"><div class="to">' + e(config_('宛名') || '　　　　　　 御中') + '</div></td>' +
    '<td class="from">' + e(inv.name) + '</td></tr></table>' +
    '<p>' + e(jpMonth_(inv.month)) + '分の清掃業務について、下記のとおりご請求申し上げます。</p>' +
    '<div class="amt">ご請求金額<b>' + yen_(inv.total) + '</b></div>' +
    '<table class="d"><tr><th style="width:11%">勤務日</th><th>物件</th><th style="width:14%">清掃費</th><th style="width:13%">駐車場代</th><th style="width:18%">その他</th><th style="width:14%">小計</th></tr>' +
    rows +
    '<tr class="t"><td class="c" colspan="2">合計（' + inv.lines.length + '日）</td><td class="n">' + yen_(sum('cleaning')) + '</td><td class="n">' + yen_(sum('parking')) +
    '</td><td class="n">' + yen_(sum('otherAmt')) + '</td><td class="n">' + yen_(inv.total) + '</td></tr></table>' +
    '<div class="bank"><b>お振込先</b><br>' + e([b.bank, b.branch, b.type, b.number].join('　')) + '<br>口座名義：' + e(b.holder) + '</div>' +
    (inv.note ? '<p class="foot">備考：' + e(inv.note) + '</p>' : '') +
    '</body></html>';
}

function yen_(n) { return '¥' + Number(n || 0).toLocaleString('ja-JP'); }
function jpMonth_(m) { const p = String(m).split('-'); return p[0] + '年' + Number(p[1]) + '月'; }

// 以前の形のシートを、今の形（明細に「名前」列、請求書に「請求書PDF」列）にそろえる
function migrate_() {
  const ss = SpreadsheetApp.getActive();
  const lines = ss.getSheetByName(SHEET.lines), inv = ss.getSheetByName(SHEET.inv);
  if (!lines || !inv) return;
  if (lines.getLastRow() > 0 && lines.getRange(1, 2).getValue() !== '名前') {
    lines.insertColumnAfter(1);
    lines.getRange(1, 2).setValue('名前');
    const n = lines.getLastRow() - 1;
    if (n > 0) {
      const names = {};
      if (inv.getLastRow() > 1) inv.getRange(2, 1, inv.getLastRow() - 1, 3).getValues().forEach(r => { names[String(r[0])] = r[2]; });
      const ids = lines.getRange(2, 1, n, 1).getValues();
      lines.getRange(2, 2, n, 1).setValues(ids.map(r => [names[String(r[0])] || '']));
    }
  }
  if (inv.getLastRow() > 0 && inv.getRange(1, 19).getValue() !== '請求書PDF') inv.getRange(1, 19).setValue('請求書PDF');
  [[lines, HEAD.lines.length], [inv, HEAD.inv.length]].forEach(([sh, w]) => sh.getRange(1, 1, 1, w).setFontWeight('bold').setBackground('#e9f0ef'));
}

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
  // 月末締め：勤務日の月がその請求の「何月分」。月をまたぐ請求は受け付けない
  const months = Array.from(new Set(lines.map(l => l.date.slice(0, 7))));
  if (months.length > 1) throw new Error('勤務日は同じ月の分だけにしてください（' + months.map(jpMonth_).join('と') + 'が混ざっています。月ごとに分けて送ってください）');
  return { name: name, bank: bank, month: months[0], date: date_(x.date, '請求日'), note: str_(x.note, 300),
           lines: lines, sum: sum, total: sum.cleaning + sum.parking + sum.other };
}
function withLock_(fn) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) throw new Error('混み合っています。少し待ってもう一度お試しください');
  try { return fn(); } finally { lock.releaseLock(); }
}
