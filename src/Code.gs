/**
 * 車の試乗予約システム（Google Apps Script + Google スプレッドシート）
 *
 * スプレッドシートに紐づけた（コンテナバインド）スクリプトとして動作し、
 * 予約画面（GitHub Pages の docs/index.html）向けの JSON API を提供します。
 *   - 「設定」シート : 試乗時間・受付時間・定休日・規約などを管理
 *   - 「試乗車」シート: 試乗車の一覧（複数台）
 *   - 「予約」シート : 予約データ（1枠につき1名）
 */

const SHEET = {
  SETTINGS: '設定',
  CARS: '試乗車',
  BOOKINGS: '予約',
};

const STATUS_BOOKED = '予約';
const STATUS_CANCELED = 'キャンセル';

const DEFAULT_TERMS = [
  '【試乗に関する規約】',
  '1. 試乗には有効な運転免許証が必要です。当日受付にてご提示ください。',
  '2. 試乗コースはスタッフが指定します。スタッフが同乗します。',
  '3. 飲酒・体調不良時の試乗はお断りします。',
  '4. 交通法規を遵守し、安全運転をお願いします。',
  '5. お客様の故意・重大な過失による事故の損害は、お客様のご負担となる場合があります。',
  '6. 予約時刻に15分以上遅れた場合、予約を取り消すことがあります。',
  '7. ご入力いただいた個人情報は、試乗のご案内およびご連絡のみに利用します。',
].join('\n');

/** 設定項目の定義（label がシート上の項目名） */
const SETTING_DEFS = [
  { key: 'storeName',    label: '店舗名',            value: '○○モータース',  note: '画面上部に表示する名称' },
  { key: 'storePhone',   label: '店舗電話番号',      value: '00-0000-0000',   note: '予約完了画面に表示（変更・キャンセル連絡先）' },
  { key: 'duration',     label: '試乗時間（分）',    value: '40',             note: '1回の試乗時間' },
  { key: 'interval',     label: 'インターバル（分）', value: '10',            note: '枠と枠の間の準備時間（0も可）' },
  { key: 'openTime',     label: '受付開始時刻',      value: '10:00',          note: '最初の枠の開始時刻（HH:MM）' },
  { key: 'closeTime',    label: '受付終了時刻',      value: '18:00',          note: '最終枠がこの時刻までに終わるよう枠を作成（HH:MM）' },
  { key: 'daysAhead',    label: '予約可能日数',      value: '14',             note: '今日から何日先まで予約できるか' },
  { key: 'cutoff',       label: '締切（分前）',      value: '60',             note: '開始時刻の何分前まで予約を受け付けるか' },
  { key: 'closedDays',   label: '定休日（曜日）',    value: '水',             note: 'カンマ区切り 例: 水,木（空欄なら定休日なし）' },
  { key: 'holidays',     label: '休業日',            value: '',               note: 'カンマ区切り 例: 2026-12-31,2027-01-01' },
  { key: 'notifyEmail',  label: '通知先メール',      value: '',               note: '新規予約の通知先（空欄なら通知しない）' },
  { key: 'terms',        label: '利用規約',          value: DEFAULT_TERMS,    note: '予約画面に表示する規約' },
];

const CAR_HEADERS = ['車両ID', '車名', '説明', '画像URL', '受付中'];
const SAMPLE_CARS = [
  ['CAR01', 'サンプルセダン', 'ハイブリッド / 5人乗り', '', true],
  ['CAR02', 'サンプルSUV', '4WD / 7人乗り', '', true],
  ['CAR03', 'サンプルコンパクト', 'ガソリン / 5人乗り', '', true],
];

const BOOKING_HEADERS = ['予約ID', '受付日時', '試乗日', '開始', '終了', '車両ID', '車名', '氏名', '連絡先', '規約同意', 'ステータス', 'メモ'];
const COL = { id: 0, createdAt: 1, date: 2, start: 3, end: 4, carId: 5, carName: 6, name: 7, contact: 8, agreed: 9, status: 10, memo: 11 };

const WEEKDAYS = ['日', '月', '火', '水', '木', '金', '土'];

// ---------------------------------------------------------------------------
// Web アプリ
// ---------------------------------------------------------------------------
// 予約画面は GitHub Pages（docs/index.html）に置き、ここは JSON API のみ提供する。
// script.google.com を画面として直接開くと、複数の Google アカウントでログイン中の
// ブラウザ（特に iPhone）で「ファイルを開くことができません」となるため。
//
//   GET  ?action=init                       … 店舗情報・試乗車・予約可能日
//   GET  ?action=slots&carId=..&date=..     … 指定した車・日付の空き状況
//   POST {"action":"book","form":{...}}     … 予約登録（Content-Type: text/plain）

function doGet(e) {
  const p = (e && e.parameter) || {};
  return api_(p.action, p);
}

function doPost(e) {
  let body = {};
  try {
    body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
  } catch (err) {
    // 不正な JSON は下の「不明なリクエスト」として扱う
  }
  return api_(body.action, body);
}

function api_(action, p) {
  let result;
  try {
    let data;
    if (action === 'init') data = getInitData();
    else if (action === 'slots') data = getSlots(String(p.carId || ''), String(p.date || ''));
    else if (action === 'book') data = book(p.form);
    else throw new Error('不明なリクエストです。');
    result = { ok: true, data: data };
  } catch (err) {
    result = { ok: false, error: (err && err.message) || String(err) };
  }
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

/** 画面の初期表示に必要なデータ */
function getInitData() {
  const s = getSettings_();
  return {
    storeName: s.storeName,
    storePhone: s.storePhone,
    duration: s.duration,
    terms: s.terms,
    cars: getActiveCars_().map(function (c) {
      return { id: c.id, name: c.name, description: c.description, imageUrl: c.imageUrl };
    }),
    dates: getAvailableDates_(s, nowInfo_()),
  };
}

/** 指定した車・日付の枠と空き状況（個人情報は返さない） */
function getSlots(carId, date) {
  const s = getSettings_();
  const now = nowInfo_();
  const car = findActiveCar_(carId);
  if (!car) throw new Error('選択された試乗車は現在受付していません。');
  if (!isBookableDate_(s, now, date)) throw new Error('選択された日付は予約できません。');

  const booked = getActiveBookings_().filter(function (b) {
    return b.carId === carId && b.date === date;
  });
  return buildSlots_(s).map(function (slot) {
    return {
      start: slot.start,
      end: slot.end,
      available: !isPast_(s, now, date, slot.start) && !hasOverlap_(booked, slot),
    };
  });
}

/** 予約の登録 */
function book(form) {
  form = form || {};
  const name = cleanText_(form.name, 50);
  const contact = toHalfWidth_(cleanText_(form.contact, 100));
  const carId = String(form.carId || '');
  const date = String(form.date || '');
  const start = String(form.start || '');

  if (!name) throw new Error('氏名を入力してください。');
  if (!isValidContact_(contact)) throw new Error('連絡先には電話番号またはメールアドレスを入力してください。');
  if (form.agreed !== true) throw new Error('規約への同意が必要です。');

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) throw new Error('混み合っています。時間をおいて再度お試しください。');
  try {
    const s = getSettings_();
    const now = nowInfo_();
    const car = findActiveCar_(carId);
    if (!car) throw new Error('選択された試乗車は現在受付していません。');
    if (!isBookableDate_(s, now, date)) throw new Error('選択された日付は予約できません。');

    const slot = buildSlots_(s).filter(function (x) { return x.start === start; })[0];
    if (!slot) throw new Error('選択された時間枠は存在しません。');
    if (isPast_(s, now, date, slot.start)) throw new Error('選択された時間枠は受付を締め切りました。');

    const booked = getActiveBookings_().filter(function (b) {
      return b.carId === carId && b.date === date;
    });
    if (hasOverlap_(booked, slot)) {
      throw new Error('申し訳ありません。この枠は先に予約が入りました。別の枠をお選びください。');
    }

    const id = 'TD' + date.replace(/-/g, '') + '-' + Utilities.getUuid().slice(0, 4).toUpperCase();
    const row = [];
    row[COL.id] = id;
    row[COL.createdAt] = now.dateTime;
    row[COL.date] = date;
    row[COL.start] = slot.start;
    row[COL.end] = slot.end;
    row[COL.carId] = car.id;
    row[COL.carName] = car.name;
    row[COL.name] = name;
    row[COL.contact] = contact;
    row[COL.agreed] = '同意済';
    row[COL.status] = STATUS_BOOKED;
    row[COL.memo] = '';

    const sheet = getSheet_(SHEET.BOOKINGS);
    const rowIndex = sheet.getLastRow() + 1;
    if (rowIndex > sheet.getMaxRows()) sheet.insertRowsAfter(sheet.getMaxRows(), 100);
    const range = sheet.getRange(rowIndex, 1, 1, BOOKING_HEADERS.length);
    range.setNumberFormat('@').setValues([row.map(escapeFormula_)]);
    SpreadsheetApp.flush();

    const result = {
      id: id,
      carName: car.name,
      date: date,
      dateLabel: dateLabel_(date),
      start: slot.start,
      end: slot.end,
      name: name,
      contact: contact,
      storeName: s.storeName,
      storePhone: s.storePhone,
    };
    sendMails_(s, result);
    return result;
  } finally {
    lock.releaseLock();
  }
}

// ---------------------------------------------------------------------------
// スプレッドシート
// ---------------------------------------------------------------------------

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('試乗予約')
    .addItem('初期設定（シート作成）', 'setup')
    .addToUi();
}

/** 必要なシートを作成し、初期値を入れる（既存データは上書きしない） */
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let sheet = ss.getSheetByName(SHEET.SETTINGS);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET.SETTINGS);
    const rows = [['項目', '値', '説明']].concat(SETTING_DEFS.map(function (d) {
      return [d.label, d.value, d.note];
    }));
    sheet.getRange(1, 1, rows.length, 3).setNumberFormat('@').setValues(rows);
    sheet.getRange(1, 1, 1, 3).setFontWeight('bold').setBackground('#e8eaf6');
    sheet.setColumnWidth(1, 160).setColumnWidth(2, 360).setColumnWidth(3, 420);
    sheet.getRange(2, 2, rows.length - 1, 1).setWrap(true).setVerticalAlignment('top');
    sheet.setFrozenRows(1);
  } else {
    // 後から追加された設定項目を補う
    const labels = sheet.getRange(1, 1, Math.max(sheet.getLastRow(), 1), 1).getDisplayValues()
      .map(function (r) { return r[0]; });
    SETTING_DEFS.forEach(function (d) {
      if (labels.indexOf(d.label) === -1) {
        sheet.appendRow([d.label, d.value, d.note]);
      }
    });
  }

  sheet = ss.getSheetByName(SHEET.CARS);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET.CARS);
    sheet.getRange(1, 1, 1, CAR_HEADERS.length).setValues([CAR_HEADERS])
      .setFontWeight('bold').setBackground('#e8eaf6');
    sheet.getRange(2, 1, SAMPLE_CARS.length, CAR_HEADERS.length).setValues(SAMPLE_CARS);
    sheet.getRange(2, 5, 200, 1).insertCheckboxes();
    sheet.setColumnWidth(2, 200).setColumnWidth(3, 260).setColumnWidth(4, 260);
    sheet.setFrozenRows(1);
  }

  sheet = ss.getSheetByName(SHEET.BOOKINGS);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET.BOOKINGS);
    sheet.getRange(1, 1, 1, BOOKING_HEADERS.length).setValues([BOOKING_HEADERS])
      .setFontWeight('bold').setBackground('#e8eaf6');
    sheet.getRange(2, 1, sheet.getMaxRows() - 1, BOOKING_HEADERS.length).setNumberFormat('@');
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInList([STATUS_BOOKED, STATUS_CANCELED], true).build();
    sheet.getRange(2, COL.status + 1, sheet.getMaxRows() - 1, 1).setDataValidation(rule);
    sheet.setFrozenRows(1);
  }

  const first = ss.getSheetByName('シート1') || ss.getSheetByName('Sheet1');
  if (first && first.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(first);
}

function getSheet_(name) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sheet) throw new Error('「' + name + '」シートがありません。メニュー「試乗予約 > 初期設定」を実行してください。');
  return sheet;
}

function getSettings_() {
  const values = getSheet_(SHEET.SETTINGS).getDataRange().getDisplayValues();
  const byLabel = {};
  values.slice(1).forEach(function (r) { byLabel[String(r[0]).trim()] = r[1]; });

  const raw = {};
  SETTING_DEFS.forEach(function (d) {
    raw[d.key] = byLabel.hasOwnProperty(d.label) ? String(byLabel[d.label]) : d.value;
  });
  return normalizeSettings_(raw);
}

function getActiveCars_() {
  const values = getSheet_(SHEET.CARS).getDataRange().getValues();
  return values.slice(1)
    .filter(function (r) { return String(r[0]).trim() && String(r[1]).trim() && isTruthy_(r[4]); })
    .map(function (r) {
      return {
        id: String(r[0]).trim(),
        name: String(r[1]).trim(),
        description: String(r[2] || '').trim(),
        imageUrl: String(r[3] || '').trim(),
      };
    });
}

function findActiveCar_(carId) {
  return getActiveCars_().filter(function (c) { return c.id === carId; })[0] || null;
}

/** キャンセル以外の予約（手入力の行も考慮して正規化） */
function getActiveBookings_() {
  const sheet = getSheet_(SHEET.BOOKINGS);
  if (sheet.getLastRow() < 2) return [];
  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, BOOKING_HEADERS.length).getDisplayValues();
  return values
    .filter(function (r) { return String(r[COL.status]).trim() !== STATUS_CANCELED; })
    .map(function (r) {
      return {
        carId: String(r[COL.carId]).trim(),
        date: normalizeDate_(r[COL.date]),
        start: normalizeTime_(r[COL.start]),
        end: normalizeTime_(r[COL.end]),
      };
    })
    .filter(function (b) { return b.carId && b.date && b.start && b.end; });
}

// ---------------------------------------------------------------------------
// 予約枠の計算（純粋関数）
// ---------------------------------------------------------------------------

function normalizeSettings_(raw) {
  const s = {
    storeName: raw.storeName || '試乗予約',
    storePhone: raw.storePhone || '',
    duration: toPositiveInt_(raw.duration, 40),
    interval: toNonNegativeInt_(raw.interval, 0),
    openTime: normalizeTime_(raw.openTime) || '10:00',
    closeTime: normalizeTime_(raw.closeTime) || '18:00',
    daysAhead: toPositiveInt_(raw.daysAhead, 14),
    cutoff: toNonNegativeInt_(raw.cutoff, 0),
    closedDays: splitList_(raw.closedDays).map(function (x) { return x.charAt(0); })
      .filter(function (x) { return WEEKDAYS.indexOf(x) !== -1; }),
    holidays: splitList_(raw.holidays).map(normalizeDate_).filter(Boolean),
    notifyEmail: String(raw.notifyEmail || '').trim(),
    terms: raw.terms || '',
  };
  return s;
}

function buildSlots_(s) {
  const slots = [];
  const open = toMinutes_(s.openTime);
  const close = toMinutes_(s.closeTime);
  const step = s.duration + s.interval;
  for (let t = open; t + s.duration <= close; t += step) {
    slots.push({ start: fromMinutes_(t), end: fromMinutes_(t + s.duration) });
  }
  return slots;
}

function hasOverlap_(bookings, slot) {
  const a1 = toMinutes_(slot.start);
  const a2 = toMinutes_(slot.end);
  return bookings.some(function (b) {
    return toMinutes_(b.start) < a2 && a1 < toMinutes_(b.end);
  });
}

function getAvailableDates_(s, now) {
  const dates = [];
  for (let i = 0; i < s.daysAhead; i++) {
    const date = addDays_(now.date, i);
    if (isOpenDay_(s, date)) {
      // 当日で受付可能な枠が残っていない場合は表示しない
      if (i === 0 && !buildSlots_(s).some(function (x) { return !isPast_(s, now, date, x.start); })) continue;
      dates.push({ date: date, label: dateLabel_(date), weekday: weekday_(date) });
    }
  }
  return dates;
}

function isBookableDate_(s, now, date) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) return false;
  const diff = daysBetween_(now.date, date);
  return diff >= 0 && diff < s.daysAhead && isOpenDay_(s, date);
}

function isOpenDay_(s, date) {
  return s.closedDays.indexOf(WEEKDAYS[weekday_(date)]) === -1 && s.holidays.indexOf(date) === -1;
}

function isPast_(s, now, date, start) {
  if (date > now.date) return false;
  if (date < now.date) return true;
  return toMinutes_(start) - s.cutoff < now.minutes;
}

// ---------------------------------------------------------------------------
// メール
// ---------------------------------------------------------------------------

function sendMails_(s, r) {
  const body = [
    '予約ID: ' + r.id,
    '試乗車: ' + r.carName,
    '日時　: ' + r.dateLabel + ' ' + r.start + '〜' + r.end,
    '氏名　: ' + r.name,
    '連絡先: ' + r.contact,
  ].join('\n');

  try {
    if (s.notifyEmail) {
      MailApp.sendEmail(s.notifyEmail, '【試乗予約】' + r.dateLabel + ' ' + r.start + ' ' + r.carName, '新しい試乗予約が入りました。\n\n' + body);
    }
    if (isEmail_(r.contact)) {
      MailApp.sendEmail(r.contact, '【' + s.storeName + '】試乗予約を受け付けました',
        r.name + ' 様\n\nこの度は試乗のご予約ありがとうございます。\n以下の内容で受け付けました。\n\n' + body +
        '\n\n当日は運転免許証をお持ちください。\n変更・キャンセルは下記までご連絡ください。\n' +
        s.storeName + (s.storePhone ? ' TEL: ' + s.storePhone : ''));
    }
  } catch (e) {
    // メール送信に失敗しても予約自体は成立させる
    console.error('mail failed: ' + e);
  }
}

// ---------------------------------------------------------------------------
// ユーティリティ
// ---------------------------------------------------------------------------

function nowInfo_() {
  const tz = Session.getScriptTimeZone() || 'Asia/Tokyo';
  const d = new Date();
  const hm = Utilities.formatDate(d, tz, 'HH:mm');
  return {
    date: Utilities.formatDate(d, tz, 'yyyy-MM-dd'),
    minutes: toMinutes_(hm),
    dateTime: Utilities.formatDate(d, tz, 'yyyy-MM-dd HH:mm:ss'),
  };
}

function toMinutes_(hm) {
  const m = /^(\d{1,2}):(\d{2})/.exec(String(hm));
  return m ? Number(m[1]) * 60 + Number(m[2]) : NaN;
}

function fromMinutes_(n) {
  return ('0' + Math.floor(n / 60)).slice(-2) + ':' + ('0' + (n % 60)).slice(-2);
}

function normalizeTime_(v) {
  const m = /(\d{1,2})[:：](\d{2})/.exec(String(v || ''));
  return m ? fromMinutes_(Number(m[1]) * 60 + Number(m[2])) : '';
}

function normalizeDate_(v) {
  const m = /(\d{4})[-\/年](\d{1,2})[-\/月](\d{1,2})/.exec(String(v || ''));
  return m ? m[1] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[3]).slice(-2) : '';
}

function parseYmd_(ymd) {
  const p = ymd.split('-').map(Number);
  return new Date(Date.UTC(p[0], p[1] - 1, p[2], 12));
}

function addDays_(ymd, n) {
  const d = parseYmd_(ymd);
  d.setUTCDate(d.getUTCDate() + n);
  return d.getUTCFullYear() + '-' + ('0' + (d.getUTCMonth() + 1)).slice(-2) + '-' + ('0' + d.getUTCDate()).slice(-2);
}

function daysBetween_(a, b) {
  return Math.round((parseYmd_(b) - parseYmd_(a)) / 86400000);
}

function weekday_(ymd) {
  return parseYmd_(ymd).getUTCDay();
}

function dateLabel_(ymd) {
  const p = ymd.split('-').map(Number);
  return p[1] + '月' + p[2] + '日(' + WEEKDAYS[weekday_(ymd)] + ')';
}

function splitList_(v) {
  return String(v || '').split(/[,、，\s]+/).map(function (x) { return x.trim(); }).filter(Boolean);
}

function toPositiveInt_(v, def) {
  const n = parseInt(String(v).replace(/[^\d]/g, ''), 10);
  return n > 0 ? n : def;
}

function toNonNegativeInt_(v, def) {
  const n = parseInt(String(v).replace(/[^\d]/g, ''), 10);
  return n >= 0 ? n : def;
}

function isTruthy_(v) {
  return v === true || /^(true|yes|1|○|◯|はい|受付|有効)$/i.test(String(v).trim());
}

function cleanText_(v, max) {
  return String(v || '').replace(/[\u0000-\u001f\u007f]/g, ' ').trim().slice(0, max);
}

/** 全角英数字・記号を半角に変換（iPhone の全角入力対策） */
function toHalfWidth_(v) {
  return String(v).replace(/[！-～]/g, function (c) {
    return String.fromCharCode(c.charCodeAt(0) - 0xFEE0);
  }).replace(/[ー－―]/g, '-').replace(/　/g, ' ');
}

function isEmail_(v) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(v);
}

function isValidContact_(v) {
  if (!v) return false;
  if (isEmail_(v)) return true;
  const digits = v.replace(/[\s\-()（）ー－]/g, '').replace(/^\+81/, '0');
  return /^0\d{9,10}$/.test(digits);
}

/** スプレッドシートの数式として解釈されないようにする */
function escapeFormula_(v) {
  const s = String(v);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}
