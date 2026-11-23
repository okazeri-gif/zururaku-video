/**
 * GAS（gas/Code.gs）のテスト。実行: npm test（= node --test）
 * GAS のソースを vm で読み、スプレッドシート・キャッシュ・ロックはすべて偽物で通す（本物のシートには触らない）。
 */
import test from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import vm from 'node:vm';

const SRC = readFileSync(new URL('../gas/Code.gs', import.meta.url), 'utf8');
const HOUR = 3600 * 1000;
const T0 = new Date('2026-11-23T10:00:00+09:00');
const at = ms => new Date(T0.getTime() + ms);

// ── 偽物 ──────────────────────────────────────────────

function fakeSheet(name) {
  const data = [];
  const range = (r, c, nr = 1, nc = 1) => ({
    getValues: () => Array.from({ length: nr }, (_, i) => Array.from({ length: nc }, (_, j) => (data[r - 1 + i] || [])[c - 1 + j] ?? '')),
    setValues: vals => vals.forEach((row, i) => row.forEach((v, j) => {
      while (data.length < r + i) data.push([]);
      data[r - 1 + i][c - 1 + j] = v;
    })),
    setNumberFormat: () => range(r, c, nr, nc),
  });
  return {
    name, data,
    setName(n) { this.name = n; },
    getRange: (r, c, nr, nc) => (typeof r === 'string' ? range(1, 1) : range(r, c, nr, nc)),
    getLastRow: () => data.length,
    appendRow: row => data.push([...row]),
    setFrozenRows: () => {},
  };
}

function load() {
  const props = {};
  const cache = {};
  let ss = null;
  const ctx = {
    console: { error: () => {}, log: () => {} },
    Logger: { log: () => {} },
    JSON, Math, Date, String, Number, isFinite,
    PropertiesService: { getScriptProperties: () => ({ getProperty: k => props[k] ?? null, setProperty: (k, v) => { props[k] = v; } }) },
    SpreadsheetApp: {
      create: () => {
        const sheets = [fakeSheet('シート1')];
        ss = {
          sheets,
          getId: () => 'sheet-1',
          getUrl: () => 'https://example/sheet-1',
          setSpreadsheetTimeZone: () => {},
          getSheets: () => sheets,
          insertSheet: n => { const s = fakeSheet(n); sheets.push(s); return s; },
          getSheetByName: n => sheets.find(s => s.name === n) || null,
        };
        return ss;
      },
      openById: () => ss,
    },
    CacheService: { getScriptCache: () => ({ get: k => cache[k] ?? null, put: (k, v) => { cache[k] = v; } }) },
    LockService: { getScriptLock: () => ({ tryLock: () => true, releaseLock: () => {} }) },
    ContentService: { MimeType: { JSON: 'json' }, createTextOutput: s => ({ setMimeType: () => ({ body: s }) }) },
  };
  vm.createContext(ctx);
  vm.runInContext(SRC, ctx);
  ctx.setup();
  const tab = n => ss.getSheetByName(n);
  const setSetting = (name, value) => {
    const row = tab('設定').data.find(r => r[0] === name);
    row[1] = value;
  };
  const viewers = () => tab('視聴者').data.slice(1);
  const logs = () => tab('視聴ログ').data.slice(1);
  return { ctx, ss, tab, setSetting, viewers, logs, cache };
}

// ── setup ─────────────────────────────────────────────

test('setup: 3つのタブと見出し・設定の最初の値を作り、2回目は作り直さない', () => {
  const { ctx, ss, tab } = load();
  assert.deepEqual(ss.sheets.map(s => s.name), ['設定', '視聴ログ', '視聴者']);
  assert.deepEqual(tab('視聴ログ').data[0], ['時刻', '視聴者ID', 'IDの種類', '経路', 'イベント', '動画の位置（秒）', '端末']);
  assert.equal(tab('視聴者').data[0].length, 12);
  const settings = Object.fromEntries(tab('設定').data.slice(1));
  assert.equal(settings['期限（時間）'], 72);
  assert.equal(settings['相談ボタンを出す秒数'], 0);
  assert.equal(settings['合言葉の秒数'], 0);
  assert.equal(settings['相談フォームのURL'], '');
  assert.equal(settings['受け付け中'], true);
  assert.ok(settings['動画ID']);
  assert.equal(ctx.setup(), 'sheet-1');
  assert.equal(ss.sheets.length, 3);
});

// ── 期限の判定 ─────────────────────────────────────────

test('期限: 初めて開いたら見られる・期限は初回＋72時間・残りは72時間', () => {
  const { ctx, viewers, logs } = load();
  const res = ctx.handleOpen_({ action: 'open', uid: 'U123', src: 'taisei', device: 'スマホ' }, T0);
  assert.equal(res.status, 'open');
  assert.equal(res.remainingMs, 72 * HOUR);
  assert.ok(res.videoId);
  const v = viewers();
  assert.equal(v.length, 1);
  assert.deepEqual(v[0].slice(0, 3), ['U123', 'uid', 'taisei']);
  assert.equal(v[0][3].getTime(), T0.getTime());
  assert.equal(v[0][4].getTime(), at(72 * HOUR).getTime());
  assert.equal(v[0][6], 1);
  assert.deepEqual(logs()[0].slice(1), ['U123', 'uid', 'taisei', '開いた', '', 'スマホ']);
});

test('期限: 72時間以内に開き直すと見られる・残りは減る・期限は数え直さない', () => {
  const { ctx, viewers } = load();
  ctx.handleOpen_({ uid: 'U1' }, T0);
  const res = ctx.handleOpen_({ uid: 'U1' }, at(71 * HOUR + 59 * 60 * 1000));
  assert.equal(res.status, 'open');
  assert.equal(res.remainingMs, 60 * 1000);
  assert.equal(viewers().length, 1);
  assert.equal(viewers()[0][4].getTime(), at(72 * HOUR).getTime());
  assert.equal(viewers()[0][6], 2);
});

test('期限: 72時間ちょうどからは見られない', () => {
  const { ctx, logs } = load();
  ctx.handleOpen_({ uid: 'U1' }, T0);
  const res = ctx.handleOpen_({ uid: 'U1' }, at(72 * HOUR));
  assert.deepEqual({ ...res }, { ok: true, status: 'expired' });
  assert.equal(logs()[1][4], '期限切れで開いた');
});

test('期限: 過ぎたら見られない・動画IDを返さない', () => {
  const { ctx } = load();
  ctx.handleOpen_({ uid: 'U1' }, T0);
  const res = ctx.handleOpen_({ uid: 'U1' }, at(100 * HOUR));
  assert.equal(res.status, 'expired');
  assert.equal(res.videoId, undefined);
});

test('期限: judge_ 単体（直前・ちょうど・過ぎた・空）', () => {
  const { ctx } = load();
  const end = at(72 * HOUR);
  assert.equal(ctx.judge_(end, at(72 * HOUR - 1)).status, 'open');
  assert.equal(ctx.judge_(end, at(72 * HOUR)).status, 'expired');
  assert.equal(ctx.judge_(end, at(73 * HOUR)).status, 'expired');
  assert.equal(ctx.judge_('', T0).status, 'expired');
});

test('期限: 設定の時間を変えると、新しく開いた人から効く', () => {
  const { ctx, setSetting } = load();
  setSetting('期限（時間）', 24);
  const res = ctx.handleOpen_({ uid: 'U9' }, T0);
  assert.equal(res.remainingMs, 24 * HOUR);
});

test('受け付け中が FALSE: 全員に期限切れ・新しい人の期限は数え始めない', () => {
  const { ctx, setSetting, viewers, logs } = load();
  ctx.handleOpen_({ uid: 'U1' }, T0);
  setSetting('受け付け中', 'FALSE');
  assert.equal(ctx.handleOpen_({ uid: 'U1' }, at(HOUR)).status, 'expired');
  assert.equal(ctx.handleOpen_({ uid: 'NEW' }, at(HOUR)).status, 'expired');
  assert.equal(viewers().length, 1);
  assert.equal(logs().filter(r => r[4] === '期限切れで開いた').length, 2);
  setSetting('受け付け中', true);
  assert.equal(ctx.handleOpen_({ uid: 'NEW' }, at(2 * HOUR)).remainingMs, 72 * HOUR);
});

// ── 視聴者ID ──────────────────────────────────────────

test('視聴者ID: uid が無ければブラウザのID・置き換え失敗の uid は無いものとして扱う', () => {
  const { ctx, viewers } = load();
  ctx.handleOpen_({ uid: '#_userParamId_#', bid: 'abc123' }, T0);
  ctx.handleOpen_({ uid: '', bid: 'abc123' }, T0);
  assert.equal(viewers().length, 1);
  assert.deepEqual(viewers()[0].slice(0, 2), ['abc123', 'browser']);
  assert.equal(ctx.handleOpen_({}, T0).error, 'no_viewer');
});

test('視聴者ID: 数式に読める値・おかしな経路は受け付けない', () => {
  const { ctx } = load();
  assert.equal(ctx.normalizeViewer_({ uid: '=HYPERLINK("x")' }), null);
  assert.equal(ctx.normalizeViewer_({ uid: 'U1', src: '=1+1' }).src, '');
  assert.equal(ctx.normalizeViewer_({ uid: 'U1', device: 'Mozilla/5.0' }).device, '');
});

// ── 見た合計（秒） ─────────────────────────────────────

test('見た合計: accumulateWatched_（再生→再生中→停止）', () => {
  const { ctx } = load();
  let s = ctx.accumulateWatched_(null, { event: 'play', pos: 0, t: 1000 });
  assert.equal(s.add, 0);
  s = ctx.accumulateWatched_(s.state, { event: 'progress', pos: 15, t: 16000 });
  assert.equal(s.add, 15);
  s = ctx.accumulateWatched_(s.state, { event: 'pause', pos: 20, t: 21000 });
  assert.equal(s.add, 5);
  // 停止のあと、時間がたって再生 → 停止中の時間は足さない
  s = ctx.accumulateWatched_(s.state, { event: 'play', pos: 20, t: 300000 });
  assert.equal(s.add, 0);
  s = ctx.accumulateWatched_(s.state, { event: 'progress', pos: 35, t: 315000 });
  assert.equal(s.add, 15);
});

test('見た合計: 早送りで飛ばした分は入れない', () => {
  const { ctx } = load();
  let s = ctx.accumulateWatched_(null, { event: 'play', pos: 0, t: 0 });
  // 5秒見て、600秒の位置まで早送り → 次の「再生中」までの経過15秒だけ足す
  s = ctx.accumulateWatched_(s.state, { event: 'progress', pos: 610, t: 15000 });
  assert.equal(s.add, 15);
});

test('見た合計: 停止の記録が抜けても、位置が進んでいなければ足さない', () => {
  const { ctx } = load();
  let s = ctx.accumulateWatched_(null, { event: 'play', pos: 100, t: 0 });
  s = ctx.accumulateWatched_(s.state, { event: 'play', pos: 103, t: 3600000 });
  assert.equal(s.add, 3);
});

test('見た合計: 巻き戻し・順番の入れ替わり・相談ボタンでは足さない', () => {
  const { ctx } = load();
  let s = ctx.accumulateWatched_(null, { event: 'play', pos: 100, t: 0 });
  const back = ctx.accumulateWatched_(s.state, { event: 'progress', pos: 50, t: 15000 });
  assert.equal(back.add, 0);
  const late = ctx.accumulateWatched_(back.state, { event: 'progress', pos: 99, t: 14000 });
  assert.equal(late.add, 0);
  assert.equal(late.state.t, 15000);
  const consult = ctx.accumulateWatched_(back.state, { event: 'consult', pos: 60, t: 20000 });
  assert.equal(consult.add, 0);
  assert.equal(consult.state.playing, true);
});

test('見た合計: 2倍速は実際の経過時間だけ足す', () => {
  const { ctx } = load();
  let s = ctx.accumulateWatched_(null, { event: 'play', pos: 0, t: 0 });
  s = ctx.accumulateWatched_(s.state, { event: 'progress', pos: 30, t: 15000 });
  assert.equal(s.add, 15);
});

test('記録: 再生→再生中→早送り→停止→最後まで見た を通して視聴者の行に入る', () => {
  const { ctx, viewers, logs } = load();
  ctx.handleOpen_({ uid: 'U1', src: 'taisei', device: 'スマホ' }, T0);
  const ev = (event, pos, t) => ctx.handleEvent_({ action: 'event', event, pos, t, uid: 'U1', src: 'taisei', device: 'スマホ', sid: 's1' }, T0);
  ev('play', 0, 1000);
  ev('progress', 15, 16000);
  ev('progress', 300, 31000); // 早送りした（足すのは 15）
  ev('pause', 305, 36000);
  ev('play', 305, 90000);
  ev('ended', 320, 105000);
  const v = viewers()[0];
  assert.equal(v[7], 320); // 見た一番先の位置
  assert.equal(v[8], 15 + 15 + 5 + 15); // 見た合計
  assert.equal(v[9], true); // 合言葉の秒数が0 → 最後まで見たときに通過
  assert.equal(v[10], true);
  assert.equal(v[11], false);
  assert.deepEqual(logs().map(r => r[4]), ['開いた', '再生', '再生中', '再生中', '停止', '再生', '最後まで見た', '合言葉の位置を通過']);
});

test('記録: ページの開き直し（sid が違う）をまたいでは足さない', () => {
  const { ctx, viewers } = load();
  ctx.handleOpen_({ uid: 'U1' }, T0);
  ctx.handleEvent_({ event: 'play', pos: 0, t: 0, uid: 'U1', sid: 'a' }, T0);
  ctx.handleEvent_({ event: 'progress', pos: 15, t: 15000, uid: 'U1', sid: 'b' }, T0);
  assert.equal(viewers()[0][8], 0);
});

test('記録: 合言葉の秒数を超えたら1回だけ記録・相談ボタンを押した', () => {
  const { ctx, setSetting, viewers, logs } = load();
  setSetting('合言葉の秒数', 600);
  ctx.handleOpen_({ uid: 'U1' }, T0);
  const ev = (event, pos, t) => ctx.handleEvent_({ event, pos, t, uid: 'U1', sid: 's' }, T0);
  ev('play', 590, 0);
  ev('progress', 605, 15000);
  ev('progress', 620, 30000);
  ev('consult', 625, 35000);
  assert.equal(logs().filter(r => r[4] === '合言葉の位置を通過').length, 1);
  assert.equal(viewers()[0][9], true);
  assert.equal(viewers()[0][10], false);
  assert.equal(viewers()[0][11], true);
});

test('記録: 知らないイベントは受け付けない', () => {
  const { ctx, logs } = load();
  ctx.handleOpen_({ uid: 'U1' }, T0);
  assert.equal(ctx.handleEvent_({ event: 'keyword', pos: 1, t: 0, uid: 'U1' }, T0).error, 'unknown_event');
  assert.equal(logs().length, 1);
});

// ── 相談ボタン ─────────────────────────────────────────

test('相談ボタン: 条件を満たした人には、次に開いたとき最初から出す', () => {
  const { ctx, setSetting } = load();
  setSetting('相談ボタンを出す秒数', 500);
  setSetting('相談フォームのURL', 'https://example.com/form');
  const first = ctx.handleOpen_({ uid: 'U1' }, T0);
  assert.equal(first.consultShown, false);
  assert.equal(first.consultUrl, 'https://example.com/form');
  assert.equal(first.consultSec, 500);
  ctx.handleEvent_({ event: 'progress', pos: 510, t: 0, uid: 'U1', sid: 's' }, T0);
  assert.equal(ctx.handleOpen_({ uid: 'U1' }, at(HOUR)).consultShown, true);
});

test('相談ボタン: https 以外のURLは返さない', () => {
  const { ctx, setSetting } = load();
  setSetting('相談フォームのURL', 'javascript:alert(1)');
  assert.equal(ctx.handleOpen_({ uid: 'U1' }, T0).consultUrl, '');
});

// ── 入口 ──────────────────────────────────────────────

test('doGet / doPost: JSON を返す', () => {
  const { ctx } = load();
  const get = JSON.parse(ctx.doGet({ parameter: { action: 'open', uid: 'U1' } }).body);
  assert.equal(get.status, 'open');
  const post = JSON.parse(ctx.doPost({ postData: { contents: JSON.stringify({ action: 'event', event: 'play', pos: 0, t: 0, uid: 'U1', sid: 's' }) } }).body);
  assert.equal(post.ok, true);
  assert.equal(JSON.parse(ctx.doPost({ postData: { contents: 'not json' } }).body).error, 'server_error');
});
