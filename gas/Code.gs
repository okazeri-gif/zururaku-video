/**
 * 動画の視聴ページ（実装指示_84）— 期限の判定と視聴の記録
 *
 * ページ（GitHub Pages の index.html）から呼ばれる入口は2つ:
 *   GET  {Webアプリ}?action=open&uid=…（または bid=…）&src=…&device=…
 *        … 「この視聴者は見てよいか」。初めてなら視聴者の行を作り、期限を数え始める。
 *        戻り値（JSON）: { ok, status: 'open' | 'expired', videoId, remainingMs, consultSec, consultUrl, consultShown }
 *        （videoId などは status が open のときだけ入る）
 *   POST {Webアプリ}（本文は JSON の文字列。Content-Type は text/plain）
 *        { action:'event', event: play|progress|pause|ended|consult, pos, t, sid, uid|bid, src, device }
 *        … 視聴ログに1行足し、視聴者の行を更新する。t はページ側の経過ミリ秒（見た合計を足すのに使う）。
 *
 * 期限の判定は GAS 側だけでやる（ページ側の時計は使わない）。
 * 設定はスプレッドシートの「設定」タブから読む（コードに直書きしない）。
 * IP アドレス・ユーザーエージェントの全文・名前は受け取らないし保存しない（端末は「スマホ／PC」だけをページが送る）。
 */

var SETTINGS_TAB = '設定';
var LOG_TAB = '視聴ログ';
var VIEWER_TAB = '視聴者';
var SHEET_ID_PROP = 'SHEET_ID';

/** 仮の動画ID（YouTube IFrame Player API の説明で使われているサンプル動画）。司令塔が「設定」タブで書き換える */
var PLACEHOLDER_VIDEO_ID = 'M7lc1UVf-VE';

var SETTING_NAMES = {
  videoId: '動画ID',
  hours: '期限（時間）',
  consultSec: '相談ボタンを出す秒数',
  keywordSec: '合言葉の秒数',
  consultUrl: '相談フォームのURL',
  accepting: '受け付け中',
};
var SETTINGS_ROWS = [
  [SETTING_NAMES.videoId, PLACEHOLDER_VIDEO_ID],
  [SETTING_NAMES.hours, 72],
  [SETTING_NAMES.consultSec, 0],
  [SETTING_NAMES.keywordSec, 0],
  [SETTING_NAMES.consultUrl, ''],
  [SETTING_NAMES.accepting, true],
];

var LOG_HEADERS = ['時刻', '視聴者ID', 'IDの種類', '経路', 'イベント', '動画の位置（秒）', '端末'];
var VIEWER_HEADERS = [
  '視聴者ID', 'IDの種類', '経路', '初めて開いた時刻', '期限', '最後に開いた時刻', '開いた回数',
  '見た一番先の位置（秒）', '見た合計（秒）', '合言葉の位置を通過', '最後まで見た', '相談ボタンを押した',
];
var V = {
  id: 0, type: 1, src: 2, first: 3, deadline: 4, last: 5, opens: 6, maxPos: 7, watched: 8, keyword: 9, ended: 10, consult: 11,
};

var ID_TYPE_UID = 'uid';
var ID_TYPE_BROWSER = 'browser';

var EVENT_LABELS = {
  open: '開いた',
  openExpired: '期限切れで開いた',
  play: '再生',
  progress: '再生中',
  pause: '停止',
  keyword: '合言葉の位置を通過',
  ended: '最後まで見た',
  consult: '相談ボタンを押した',
};
/** ページから送ってよいイベント（合言葉の通過は GAS が位置から判定する） */
var CLIENT_EVENTS = { play: true, progress: true, pause: true, ended: true, consult: true };
var DEVICES = { 'スマホ': true, 'PC': true };

/** 「再生中」の間隔として足す上限（秒）。ページは15秒ごとに送るので、これを超える間隔は異常として切り詰める */
var MAX_INTERVAL_SEC = 60;
/** 動画の位置として受け付ける上限（秒） */
var MAX_POS_SEC = 6 * 60 * 60;
/** 再生の状態（直前の位置と時刻）を覚えておく時間（秒）。CacheService の上限 */
var STATE_TTL_SEC = 21600;

var DATE_FORMAT = 'yyyy-mm-dd hh:mm:ss';

// ── 準備 ──────────────────────────────────────────────

/** 1回だけ動かす: 記録用のスプレッドシートを作り、ID をスクリプトプロパティに保存する */
function setup() {
  var props = PropertiesService.getScriptProperties();
  var existing = props.getProperty(SHEET_ID_PROP);
  if (existing) {
    Logger.log('もう作ってあります: https://docs.google.com/spreadsheets/d/' + existing);
    return existing;
  }
  var ss = SpreadsheetApp.create('動画の視聴ページ_記録');
  ss.setSpreadsheetTimeZone('Asia/Tokyo');

  var settings = ss.getSheets()[0];
  settings.setName(SETTINGS_TAB);
  settings.getRange(1, 1, 1, 2).setValues([['項目', '値']]);
  settings.getRange(2, 1, SETTINGS_ROWS.length, 2).setValues(SETTINGS_ROWS);
  settings.setFrozenRows(1);

  var log = ss.insertSheet(LOG_TAB);
  log.getRange(1, 1, 1, LOG_HEADERS.length).setValues([LOG_HEADERS]);
  log.getRange('A:A').setNumberFormat(DATE_FORMAT);
  log.getRange('B:B').setNumberFormat('@');
  log.setFrozenRows(1);

  var viewers = ss.insertSheet(VIEWER_TAB);
  viewers.getRange(1, 1, 1, VIEWER_HEADERS.length).setValues([VIEWER_HEADERS]);
  viewers.getRange('A:A').setNumberFormat('@');
  viewers.getRange('D:F').setNumberFormat(DATE_FORMAT);
  viewers.setFrozenRows(1);

  props.setProperty(SHEET_ID_PROP, ss.getId());
  Logger.log('作りました: ' + ss.getUrl());
  return ss.getId();
}

// ── 入口 ──────────────────────────────────────────────

function doGet(e) {
  var p = (e && e.parameter) || {};
  try {
    if (p.action === 'open') return json_(handleOpen_(p, new Date()));
    return json_({ ok: false, error: 'unknown_action' });
  } catch (err) {
    console.error(err);
    return json_({ ok: false, error: 'server_error' });
  }
}

function doPost(e) {
  try {
    var body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (body.action === 'event') return json_(handleEvent_(body, new Date()));
    return json_({ ok: false, error: 'unknown_action' });
  } catch (err) {
    console.error(err);
    return json_({ ok: false, error: 'server_error' });
  }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// ── 「見てよいか」 ────────────────────────────────────────

function handleOpen_(p, now) {
  var who = normalizeViewer_(p);
  if (!who) return { ok: false, error: 'no_viewer' };
  var ss = openSheet_();
  var settings = readSettings_(ss);

  return withLock_(function () {
    var sheet = ss.getSheetByName(VIEWER_TAB);
    var found = findViewer_(sheet, who.id, who.idType);
    var row = found ? found.values : null;

    // 受け付けを止めているときは、新しい人の期限を数え始めない（行を作らない）
    if (!row && settings.accepting) row = newViewerRow_(who, now, settings.hours);
    if (row) {
      row[V.last] = now;
      row[V.opens] = toNumber_(row[V.opens]) + 1;
      if (!row[V.src] && who.src) row[V.src] = who.src;
      writeViewer_(sheet, found, row);
    }

    var verdict = settings.accepting && row ? judge_(row[V.deadline], now) : { status: 'expired', remainingMs: 0 };
    appendLog_(ss, now, who, verdict.status === 'open' ? EVENT_LABELS.open : EVENT_LABELS.openExpired, '');

    if (verdict.status !== 'open') return { ok: true, status: 'expired' };
    return {
      ok: true,
      status: 'open',
      videoId: settings.videoId,
      remainingMs: verdict.remainingMs,
      consultSec: settings.consultSec,
      consultUrl: settings.consultUrl,
      consultShown: consultShown_(row, settings),
    };
  });
}

/**
 * 期限の判定。deadline ちょうどの時刻からは見られない（初めて開いた時刻＋72時間「未満」だけ見られる）。
 * @return {{status:'open'|'expired', remainingMs:number}}
 */
function judge_(deadline, now) {
  var end = toTime_(deadline);
  if (!isFinite(end)) return { status: 'expired', remainingMs: 0 };
  var remaining = end - now.getTime();
  return remaining > 0 ? { status: 'open', remainingMs: remaining } : { status: 'expired', remainingMs: 0 };
}

function newViewerRow_(who, now, hours) {
  var row = [];
  row[V.id] = who.id;
  row[V.type] = who.idType;
  row[V.src] = who.src;
  row[V.first] = now;
  row[V.deadline] = new Date(now.getTime() + hours * 3600 * 1000);
  row[V.last] = now;
  row[V.opens] = 0;
  row[V.maxPos] = 0;
  row[V.watched] = 0;
  row[V.keyword] = false;
  row[V.ended] = false;
  row[V.consult] = false;
  return row;
}

/** 相談ボタンを最初から出しておくか（一度出る条件を満たした人） */
function consultShown_(row, settings) {
  if (toBool_(row[V.ended]) || toBool_(row[V.consult])) return true;
  return settings.consultSec > 0 && toNumber_(row[V.maxPos]) >= settings.consultSec;
}

// ── 記録 ──────────────────────────────────────────────

function handleEvent_(b, now) {
  var who = normalizeViewer_(b);
  if (!who) return { ok: false, error: 'no_viewer' };
  var event = String(b.event || '');
  if (!CLIENT_EVENTS[event]) return { ok: false, error: 'unknown_event' };
  var pos = normalizePos_(b.pos);
  var ev = { event: event, pos: pos, t: Number(b.t) };
  var sid = String(b.sid || '').replace(/[^A-Za-z0-9]/g, '').slice(0, 40);

  var ss = openSheet_();
  var settings = readSettings_(ss);

  return withLock_(function () {
    appendLog_(ss, now, who, EVENT_LABELS[event], pos);

    var sheet = ss.getSheetByName(VIEWER_TAB);
    var found = findViewer_(sheet, who.id, who.idType);
    if (!found) return { ok: true, viewer: false };
    var row = found.values;

    // 見た合計: 直前の記録（同じページの開き方＝sid ごと）との間隔から足す
    if (sid && isFinite(ev.t)) {
      var cache = CacheService.getScriptCache();
      var key = 'w:' + who.idType + ':' + who.id + ':' + sid;
      var acc = accumulateWatched_(parseState_(cache.get(key)), ev);
      if (acc.add > 0) row[V.watched] = round1_(toNumber_(row[V.watched]) + acc.add);
      cache.put(key, JSON.stringify(acc.state), STATE_TTL_SEC);
    }

    if (event !== 'consult' && pos > toNumber_(row[V.maxPos])) row[V.maxPos] = pos;
    if (event === 'ended') row[V.ended] = true;
    if (event === 'consult') row[V.consult] = true;

    // 合言葉の位置を通過（秒数が0なら、最後まで見たときに通過とする）
    if (!toBool_(row[V.keyword]) && keywordPassed_(ev, settings.keywordSec)) {
      row[V.keyword] = true;
      appendLog_(ss, now, who, EVENT_LABELS.keyword, pos);
    }

    writeViewer_(sheet, found, row);
    return { ok: true };
  });
}

function keywordPassed_(ev, keywordSec) {
  if (ev.event === 'consult') return false;
  if (ev.event === 'ended') return true;
  return keywordSec > 0 && ev.pos >= keywordSec;
}

/**
 * 「見た合計（秒）」に足す秒数を、直前の記録との間隔から求める。
 *   直前が再生中（再生／再生中）だったときだけ足す。足すのは「ページ側の経過時間」と「動画の位置の進み」の小さいほう。
 *   - 早送り: 位置は大きく進むが経過時間は短い → 経過時間だけ足す（飛ばした分は入らない）
 *   - 停止の記録が抜けた: 経過時間は長いが位置は進まない → 位置の進みだけ足す
 *   - 巻き戻し: 位置が戻る → 足さない
 *   - 順番が入れ替わって届いた（t が直前より前）: 足さず、状態も変えない
 * @param {{t:number,pos:number,playing:boolean}|null} prev
 * @param {{event:string,pos:number,t:number}} ev  t はページ側の経過ミリ秒
 * @return {{add:number, state:{t:number,pos:number,playing:boolean}}}
 */
function accumulateWatched_(prev, ev) {
  if (prev && !(ev.t > prev.t)) return { add: 0, state: prev };
  // 相談ボタンは再生の状態を変えない
  if (ev.event === 'consult') return { add: 0, state: prev || { t: ev.t, pos: ev.pos, playing: false } };

  var add = 0;
  if (prev && prev.playing) {
    var dt = (ev.t - prev.t) / 1000;
    var dp = ev.pos - prev.pos;
    if (dp > 0) add = Math.min(dt, dp, MAX_INTERVAL_SEC);
  }
  var playing = ev.event === 'play' || ev.event === 'progress';
  return { add: round1_(add), state: { t: ev.t, pos: ev.pos, playing: playing } };
}

// ── シート ────────────────────────────────────────────

function openSheet_() {
  var id = PropertiesService.getScriptProperties().getProperty(SHEET_ID_PROP);
  if (!id) throw new Error('SHEET_ID がありません。setup() を先に動かしてください');
  return SpreadsheetApp.openById(id);
}

/** 「設定」タブを項目名で読む（行の順番が変わっても読める） */
function readSettings_(ss) {
  var sheet = ss.getSheetByName(SETTINGS_TAB);
  var map = {};
  var last = sheet.getLastRow();
  if (last >= 2) {
    sheet.getRange(2, 1, last - 1, 2).getValues().forEach(function (r) {
      map[String(r[0]).trim()] = r[1];
    });
  }
  var hours = toNumber_(map[SETTING_NAMES.hours]);
  var url = String(map[SETTING_NAMES.consultUrl] == null ? '' : map[SETTING_NAMES.consultUrl]).trim();
  var accepting = map[SETTING_NAMES.accepting];
  return {
    videoId: String(map[SETTING_NAMES.videoId] == null ? '' : map[SETTING_NAMES.videoId]).trim(),
    hours: hours > 0 ? hours : 72,
    consultSec: Math.max(0, toNumber_(map[SETTING_NAMES.consultSec])),
    keywordSec: Math.max(0, toNumber_(map[SETTING_NAMES.keywordSec])),
    consultUrl: /^https:\/\//i.test(url) ? url : '',
    // 空や読めない値は「受け付け中」とみなす。FALSE と書いたときだけ止める
    accepting: !(accepting === false || String(accepting).trim().toUpperCase() === 'FALSE'),
  };
}

function findViewer_(sheet, id, idType) {
  var last = sheet.getLastRow();
  if (last < 2) return null;
  var keys = sheet.getRange(2, 1, last - 1, 2).getValues();
  for (var i = 0; i < keys.length; i++) {
    if (String(keys[i][0]) === id && String(keys[i][1]) === idType) {
      var rowIndex = i + 2;
      return { rowIndex: rowIndex, values: sheet.getRange(rowIndex, 1, 1, VIEWER_HEADERS.length).getValues()[0] };
    }
  }
  return null;
}

function writeViewer_(sheet, found, row) {
  var values = [];
  for (var i = 0; i < VIEWER_HEADERS.length; i++) values.push(row[i] === undefined ? '' : row[i]);
  var rowIndex = found ? found.rowIndex : sheet.getLastRow() + 1;
  sheet.getRange(rowIndex, 1, 1, values.length).setValues([values]);
  if (found) found.values = values;
}

function appendLog_(ss, now, who, label, pos) {
  ss.getSheetByName(LOG_TAB).appendRow([now, who.id, who.idType, who.src, label, pos, who.device]);
}

function withLock_(fn) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) return { ok: false, error: 'busy' };
  try {
    return fn();
  } finally {
    lock.releaseLock();
  }
}

// ── 値の整え ──────────────────────────────────────────

/**
 * 視聴者を決める。uid があれば uid（IDの種類＝uid）、無ければブラウザのID（bid・IDの種類＝browser）。
 * iLoop の差し込みに失敗した「#_userParamId_#」は uid 無しとして扱う。
 * 先頭は英数字だけ許す（シートで数式として読まれないように）。
 */
function normalizeViewer_(p) {
  var uid = cleanId_(p.uid);
  var bid = cleanId_(p.bid);
  var id = uid || bid;
  if (!id) return null;
  var src = String(p.src == null ? '' : p.src).trim();
  if (!/^[A-Za-z0-9_-]{1,40}$/.test(src)) src = '';
  var device = String(p.device || '');
  return {
    id: id,
    idType: uid ? ID_TYPE_UID : ID_TYPE_BROWSER,
    src: src,
    device: DEVICES[device] ? device : '',
  };
}

function cleanId_(v) {
  var s = String(v == null ? '' : v).trim();
  if (!s || s.indexOf('userParamId') >= 0) return '';
  return /^[A-Za-z0-9][A-Za-z0-9_.~=+\/-]{0,127}$/.test(s) ? s : '';
}

function parseState_(raw) {
  if (!raw) return null;
  try {
    var s = JSON.parse(raw);
    return s && isFinite(s.t) && isFinite(s.pos) ? s : null;
  } catch (e) {
    return null;
  }
}

function normalizePos_(v) {
  var n = Number(v);
  if (!isFinite(n) || n < 0) return 0;
  return round1_(Math.min(n, MAX_POS_SEC));
}

function toNumber_(v) {
  var n = Number(v);
  return isFinite(n) ? n : 0;
}

function toBool_(v) {
  return v === true || String(v).trim().toUpperCase() === 'TRUE';
}

function toTime_(v) {
  if (v instanceof Date) return v.getTime();
  if (typeof v === 'number') return v;
  if (v === '' || v == null) return NaN;
  return new Date(v).getTime();
}

function round1_(n) {
  return Math.round(n * 10) / 10;
}
