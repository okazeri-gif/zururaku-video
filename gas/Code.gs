/**
 * 動画の視聴ページ（実装指示_84・85）— 期限の判定と視聴の記録
 *
 * ページ（GitHub Pages の index.html）から呼ばれる入口は2つ:
 *   GET  {Webアプリ}?action=open&uid=…（または bid=…）&src=…&device=…&again=1（任意）
 *        … 「この視聴者は見てよいか」。初めてなら視聴者の行を作り、期限を数え始める。
 *        again=1 は「LINEに『動画視聴希望』と送った人」の合図。期限切れなら1回だけ再公開する（実装指示_85）。
 *        戻り値（JSON）:
 *          見られる  { ok, status:'open', videoId, remainingMs, consultSec, consultUrl, consultShown, reopenUsed }
 *          期限切れ  { ok, status:'expired', reopenUsed }（reopenUsed: true＝再公開を使い終わった／使えない）
 *          uid が要る { ok, status:'needUid' }（設定「uid必須」が TRUE で、使える uid が無い。視聴者の行は作らない）
 *   POST {Webアプリ}（本文は JSON の文字列。Content-Type は text/plain）
 *        { action:'event', event: play|progress|pause|ended|consult, pos, t, sid, uid|bid, src, device }
 *        … 視聴ログに1行足し、視聴者の行を更新する。t はページ側の経過ミリ秒（見た合計を足すのに使う）。
 *
 * 期限の判定は GAS 側だけでやる（ページ側の時計は使わない）。
 * 設定はスプレッドシートの「設定」タブから項目名で読む（コードに直書きしない）。
 * 「視聴者」「視聴ログ」タブは見出しの名前で読み書きする（列の位置で決め打ちしない）。
 * IP アドレス・ユーザーエージェントの全文・名前は受け取らないし保存しない（端末は「スマホ／PC」だけをページが送る）。
 *
 * 今ある本番のシートには migrate85()／migrateUidRequired() を動かす（足りない設定の行と視聴者の列だけを足す。何回動かしても同じ）。
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
  reopenHours: '再公開（時間）',
  uidRequired: 'uid必須',
};
var SETTINGS_ROWS = [
  [SETTING_NAMES.videoId, PLACEHOLDER_VIDEO_ID],
  [SETTING_NAMES.hours, 72],
  [SETTING_NAMES.consultSec, 0],
  [SETTING_NAMES.keywordSec, 0],
  [SETTING_NAMES.consultUrl, ''],
  [SETTING_NAMES.accepting, true],
  [SETTING_NAMES.reopenHours, 24],
  [SETTING_NAMES.uidRequired, false],
];

/** 視聴ログの列（キー → 見出し）。並びは setup() で作るときの順番 */
var LOG_COLS = {
  time: '時刻', id: '視聴者ID', type: 'IDの種類', src: '経路', event: 'イベント', pos: '動画の位置（秒）', device: '端末',
};
/** 視聴者の列（キー → 見出し）。並びは setup() で作るときの順番。reopenedAt・reopenDeadline は実装指示_85 で右端に足した */
var VIEWER_COLS = {
  id: '視聴者ID',
  type: 'IDの種類',
  src: '経路',
  first: '初めて開いた時刻',
  deadline: '期限',
  last: '最後に開いた時刻',
  opens: '開いた回数',
  maxPos: '見た一番先の位置（秒）',
  watched: '見た合計（秒）',
  keyword: '合言葉の位置を通過',
  ended: '最後まで見た',
  consult: '相談ボタンを押した',
  reopenedAt: '再公開した時刻',
  reopenDeadline: '再公開の期限',
};
var LOG_HEADERS = values_(LOG_COLS);
var VIEWER_HEADERS = values_(VIEWER_COLS);
/** 日時の書式にする列・文字列の書式にする列 */
var VIEWER_DATE_KEYS = ['first', 'deadline', 'last', 'reopenedAt', 'reopenDeadline'];

var ID_TYPE_UID = 'uid';
var ID_TYPE_BROWSER = 'browser';

var EVENT_LABELS = {
  open: '開いた',
  openExpired: '期限切れで開いた',
  openNoUid: 'uid無しで開いた',
  reopen: '再公開で開いた',
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
var HOUR_MS = 3600 * 1000;

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
  log.getRange(1, 1, log.getMaxRows(), 1).setNumberFormat(DATE_FORMAT);
  log.getRange(1, 2, log.getMaxRows(), 1).setNumberFormat('@');
  log.setFrozenRows(1);

  var viewers = ss.insertSheet(VIEWER_TAB);
  viewers.getRange(1, 1, 1, VIEWER_HEADERS.length).setValues([VIEWER_HEADERS]);
  formatViewerColumns_(viewers);
  viewers.setFrozenRows(1);

  props.setProperty(SHEET_ID_PROP, ss.getId());
  Logger.log('作りました: ' + ss.getUrl());
  return ss.getId();
}

/**
 * 実装指示_85: 今ある本番のシートに、足りないものだけを足す（作り直さない・何回動かしても同じ）。
 *   - 「設定」タブ: 「再公開（時間）」＝24 の行（無ければ末尾に）
 *   - 「視聴者」タブ: 「再公開した時刻」「再公開の期限」の列（無ければ右端に）
 * @return {{settingsAdded:string[], viewerColumnsAdded:string[]}}
 */
function migrate85() {
  return migrateSheets_();
}

/**
 * 「uid必須」: 今ある本番のシートの「設定」タブに「uid必須」＝FALSE の行を足す（無ければ末尾に）。
 * 中身は migrate85 と同じ（足りない設定の行と視聴者の列を全部足す。何回動かしても同じ）。
 */
function migrateUidRequired() {
  return migrateSheets_();
}

/** 足りない設定の行（SETTINGS_ROWS の順で末尾に）と、視聴者の列（右端に）だけを足す。今ある値は書き換えない */
function migrateSheets_() {
  var ss = openSheet_();
  var result = { settingsAdded: [], viewerColumnsAdded: [] };

  var settings = ss.getSheetByName(SETTINGS_TAB);
  var last = settings.getLastRow();
  var names = last >= 2 ? settings.getRange(2, 1, last - 1, 1).getValues().map(function (r) { return String(r[0]).trim(); }) : [];
  SETTINGS_ROWS.forEach(function (row) {
    if (names.indexOf(row[0]) >= 0) return;
    settings.getRange(settings.getLastRow() + 1, 1, 1, 2).setValues([row]);
    names.push(row[0]);
    result.settingsAdded.push(row[0]);
  });

  var viewers = ss.getSheetByName(VIEWER_TAB);
  var headers = headerRow_(viewers);
  VIEWER_HEADERS.forEach(function (h) {
    if (headers.indexOf(h) >= 0) return;
    var col = headers.length + 1;
    if (viewers.getMaxColumns() < col) viewers.insertColumnsAfter(viewers.getMaxColumns(), col - viewers.getMaxColumns());
    viewers.getRange(1, col, 1, 1).setValues([[h]]);
    headers.push(h);
    result.viewerColumnsAdded.push(h);
  });
  formatViewerColumns_(viewers);

  Logger.log('足した設定: ' + (result.settingsAdded.join('、') || 'なし') + '／足した列: ' + (result.viewerColumnsAdded.join('、') || 'なし'));
  return result;
}

function formatViewerColumns_(sheet) {
  var cols = headerIndex_(sheet);
  var rows = sheet.getMaxRows();
  if (cols[VIEWER_COLS.id] != null) sheet.getRange(1, cols[VIEWER_COLS.id] + 1, rows, 1).setNumberFormat('@');
  VIEWER_DATE_KEYS.forEach(function (k) {
    var c = cols[VIEWER_COLS[k]];
    if (c != null) sheet.getRange(2, c + 1, Math.max(1, rows - 1), 1).setNumberFormat(DATE_FORMAT);
  });
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
  var again = String(p.again || '') === '1';
  var ss = openSheet_();
  var settings = readSettings_(ss);

  // uid必須のときは、使える uid が無い人に動画を出さない。視聴者の行は作らない（ログに1行だけ残す）
  if (settings.uidRequired && who.idType !== ID_TYPE_UID) {
    return withLock_(function () {
      appendLog_(ss, now, who, EVENT_LABELS.openNoUid, '');
      return { ok: true, status: 'needUid' };
    });
  }

  return withLock_(function () {
    var sheet = ss.getSheetByName(VIEWER_TAB);
    var table = viewerTable_(sheet);
    // 「視聴者」タブに再公開の列がまだ無い（migrate85 の前）ときは、再公開しない
    if (!table.has.reopenedAt || !table.has.reopenDeadline) settings.reopenHours = 0;

    var found = findViewer_(table, who.id, who.idType);
    var d = decideOpen_(found ? found.v : null, now, settings, again);

    var v = found ? found.v : null;
    if (d.action === 'create') v = newViewer_(who, now, settings.hours);
    if (d.action === 'reopen') {
      v.reopenedAt = now;
      v.reopenDeadline = new Date(now.getTime() + settings.reopenHours * HOUR_MS);
    }
    if (v) {
      v.last = now;
      v.opens = toNumber_(v.opens) + 1;
      if (!v.src && who.src) v.src = who.src;
      writeViewer_(sheet, table, found, v);
    }
    appendLog_(ss, now, who, d.label, '');

    if (d.status !== 'open') return { ok: true, status: 'expired', reopenUsed: d.reopenUsed };
    return {
      ok: true,
      status: 'open',
      videoId: settings.videoId,
      remainingMs: d.remainingMs,
      consultSec: settings.consultSec,
      consultUrl: settings.consultUrl,
      consultShown: consultShown_(v, settings),
      // この期間が終わったときに出す期限切れの画面（true＝使い終わった版）。再公開中か、再公開（時間）が0なら true
      reopenUsed: d.reopenUsed || !(settings.reopenHours > 0),
    };
  });
}

/**
 * 開いたときにどうするかを決める（シートには触らない）。実装指示_85 の表のとおり。
 *   受け付け中が FALSE … 全員に期限切れ。再公開もしない（reopenUsed=true：「LINEに送って」を出さない）
 *   行が無い          … 初回（again があっても普通の初回。72時間が始まる）
 *   期限内            … 見せる（again は使わない）
 *   再公開の期限がある … その期限内なら見せる（again が無くても）。過ぎたら使い終わった版
 *   期限切れ・未使用   … again=1 かつ 再公開（時間）> 0 なら再公開。そうでなければ期限切れ
 *                        （再公開（時間）が 0 なら使い終わった版、それ以外はまだ使える版）
 * @return {{action:'create'|'reopen'|'none', status:'open'|'expired', remainingMs:number, reopenUsed:boolean, label:string}}
 */
function decideOpen_(v, now, settings, again) {
  var expired = function (used) {
    return { action: 'none', status: 'expired', remainingMs: 0, reopenUsed: used, label: EVENT_LABELS.openExpired };
  };
  if (!settings.accepting) return expired(true);
  if (!v) return { action: 'create', status: 'open', remainingMs: settings.hours * HOUR_MS, reopenUsed: false, label: EVENT_LABELS.open };

  var first = judge_(v.deadline, now);
  if (first.status === 'open') return { action: 'none', status: 'open', remainingMs: first.remainingMs, reopenUsed: false, label: EVENT_LABELS.open };

  if (isFinite(toTime_(v.reopenDeadline))) {
    var second = judge_(v.reopenDeadline, now);
    if (second.status === 'open') return { action: 'none', status: 'open', remainingMs: second.remainingMs, reopenUsed: true, label: EVENT_LABELS.open };
    return expired(true);
  }

  if (settings.reopenHours > 0 && again) {
    return { action: 'reopen', status: 'open', remainingMs: settings.reopenHours * HOUR_MS, reopenUsed: true, label: EVENT_LABELS.reopen };
  }
  return expired(!(settings.reopenHours > 0));
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

function newViewer_(who, now, hours) {
  return {
    id: who.id,
    type: who.idType,
    src: who.src,
    first: now,
    deadline: new Date(now.getTime() + hours * HOUR_MS),
    last: now,
    opens: 0,
    maxPos: 0,
    watched: 0,
    keyword: false,
    ended: false,
    consult: false,
    reopenedAt: '',
    reopenDeadline: '',
  };
}

/** 相談ボタンを最初から出しておくか（一度出る条件を満たした人） */
function consultShown_(v, settings) {
  if (toBool_(v.ended) || toBool_(v.consult)) return true;
  return settings.consultSec > 0 && toNumber_(v.maxPos) >= settings.consultSec;
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
  // uid必須のときは、uid の無い人の再生の記録は受け取らない（動画を出していないので、来るのはおかしな記録だけ）
  if (settings.uidRequired && who.idType !== ID_TYPE_UID) return { ok: true, ignored: 'need_uid' };

  return withLock_(function () {
    appendLog_(ss, now, who, EVENT_LABELS[event], pos);

    var sheet = ss.getSheetByName(VIEWER_TAB);
    var table = viewerTable_(sheet);
    var found = findViewer_(table, who.id, who.idType);
    if (!found) return { ok: true, viewer: false };
    var v = found.v;

    // 見た合計: 直前の記録（同じページの開き方＝sid ごと）との間隔から足す
    if (sid && isFinite(ev.t)) {
      var cache = CacheService.getScriptCache();
      var key = 'w:' + who.idType + ':' + who.id + ':' + sid;
      var acc = accumulateWatched_(parseState_(cache.get(key)), ev);
      if (acc.add > 0) v.watched = round1_(toNumber_(v.watched) + acc.add);
      cache.put(key, JSON.stringify(acc.state), STATE_TTL_SEC);
    }

    if (event !== 'consult' && pos > toNumber_(v.maxPos)) v.maxPos = pos;
    if (event === 'ended') v.ended = true;
    if (event === 'consult') v.consult = true;

    // 合言葉の位置を通過（秒数が0なら、最後まで見たときに通過とする）
    if (!toBool_(v.keyword) && keywordPassed_(ev, settings.keywordSec)) {
      v.keyword = true;
      appendLog_(ss, now, who, EVENT_LABELS.keyword, pos);
    }

    writeViewer_(sheet, table, found, v);
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
    // 行が無い（migrate85 の前）ときは 0＝再公開しない
    reopenHours: Math.max(0, toNumber_(map[SETTING_NAMES.reopenHours])),
    // 行が無い・空・TRUE 以外は FALSE（今までどおり、uid の無い人にもブラウザのIDで見せる）
    uidRequired: toBool_(map[SETTING_NAMES.uidRequired]),
    // 空や読めない値は「受け付け中」とみなす。FALSE と書いたときだけ止める
    accepting: !(accepting === false || String(accepting).trim().toUpperCase() === 'FALSE'),
  };
}

/** 1行目の見出し（右端の空の見出しは除く） */
function headerRow_(sheet) {
  var width = sheet.getLastColumn();
  if (width < 1) return [];
  var row = sheet.getRange(1, 1, 1, width).getValues()[0].map(function (h) { return String(h).trim(); });
  while (row.length && !row[row.length - 1]) row.pop();
  return row;
}

/** 見出し → 列の位置（0始まり） */
function headerIndex_(sheet) {
  var idx = {};
  headerRow_(sheet).forEach(function (h, i) { if (h && idx[h] == null) idx[h] = i; });
  return idx;
}

/** 「視聴者」タブを見出しの名前で読む準備。has[キー] はその列があるか */
function viewerTable_(sheet) {
  var idx = headerIndex_(sheet);
  var cols = {};
  var has = {};
  Object.keys(VIEWER_COLS).forEach(function (k) {
    if (idx[VIEWER_COLS[k]] != null) { cols[k] = idx[VIEWER_COLS[k]]; has[k] = true; }
  });
  if (!has.id || !has.type) throw new Error('「視聴者」タブに「視聴者ID」「IDの種類」の見出しがありません');
  return { sheet: sheet, cols: cols, has: has, width: Math.max(headerRow_(sheet).length, 1) };
}

function findViewer_(table, id, idType) {
  var sheet = table.sheet;
  var last = sheet.getLastRow();
  if (last < 2) return null;
  var rows = sheet.getRange(2, 1, last - 1, table.width).getValues();
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i][table.cols.id]) === id && String(rows[i][table.cols.type]) === idType) {
      var v = {};
      Object.keys(table.cols).forEach(function (k) { v[k] = rows[i][table.cols[k]]; });
      return { rowIndex: i + 2, raw: rows[i], v: v };
    }
  }
  return null;
}

/** 見出しのある列だけを書く（知らない列・見出しの無い列の値はそのまま残す） */
function writeViewer_(sheet, table, found, v) {
  var raw = found ? found.raw.slice() : [];
  while (raw.length < table.width) raw.push('');
  Object.keys(table.cols).forEach(function (k) {
    if (v[k] !== undefined) raw[table.cols[k]] = v[k];
  });
  var rowIndex = found ? found.rowIndex : sheet.getLastRow() + 1;
  sheet.getRange(rowIndex, 1, 1, raw.length).setValues([raw]);
  if (found) { found.raw = raw; found.v = v; }
}

function appendLog_(ss, now, who, label, pos) {
  var sheet = ss.getSheetByName(LOG_TAB);
  var idx = headerIndex_(sheet);
  var values = { time: now, id: who.id, type: who.idType, src: who.src, event: label, pos: pos, device: who.device };
  var row = [];
  var width = headerRow_(sheet).length;
  for (var i = 0; i < width; i++) row.push('');
  Object.keys(LOG_COLS).forEach(function (k) {
    if (idx[LOG_COLS[k]] != null) row[idx[LOG_COLS[k]]] = values[k];
  });
  sheet.appendRow(row);
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

function values_(obj) {
  return Object.keys(obj).map(function (k) { return obj[k]; });
}

function round1_(n) {
  return Math.round(n * 10) / 10;
}
