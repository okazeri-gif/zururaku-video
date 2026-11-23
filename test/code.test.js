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
    getRange: (r, c, nr, nc) => {
      if (typeof r !== 'number' || typeof c !== 'number') throw new Error('getRange は数字で呼ぶ');
      return range(r, c, nr, nc);
    },
    getLastRow: () => data.length,
    getLastColumn: () => Math.max(0, ...data.map(row => row.length)),
    getMaxRows: () => Math.max(1000, data.length),
    getMaxColumns: () => Math.max(26, ...data.map(row => row.length)),
    insertColumnsAfter: () => {},
    appendRow: row => data.push([...row]),
    setFrozenRows: () => {},
  };
}

function load({ sheet84 = false } = {}) {
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
  if (sheet84) makeSheet84(ss);
  const tab = n => ss.getSheetByName(n);
  const setSetting = (name, value) => {
    const row = tab('設定').data.find(r => r[0] === name);
    row[1] = value;
  };
  const viewers = () => tab('視聴者').data.slice(1);
  const logs = () => tab('視聴ログ').data.slice(1);
  return { ctx, ss, tab, setSetting, viewers, logs, cache };
}

/** 実装指示_84 の時点の本番シート（設定6行・視聴者12列・テストの行2行）に作り変える */
const HEADERS_84 = ['視聴者ID', 'IDの種類', '経路', '初めて開いた時刻', '期限', '最後に開いた時刻', '開いた回数',
  '見た一番先の位置（秒）', '見た合計（秒）', '合言葉の位置を通過', '最後まで見た', '相談ボタンを押した'];
function makeSheet84(ss) {
  const settings = ss.getSheetByName('設定');
  settings.data.splice(0, settings.data.length, ['項目', '値'], ['動画ID', 'M7lc1UVf-VE'], ['期限（時間）', 72],
    ['相談ボタンを出す秒数', 0], ['合言葉の秒数', 0], ['相談フォームのURL', ''], ['受け付け中', true]);
  const viewers = ss.getSheetByName('視聴者');
  viewers.data.splice(0, viewers.data.length, [...HEADERS_84],
    ['TEST1', 'uid', 'taisei', T0, at(72 * HOUR), T0, 1, 0, 0, false, false, false],
    ['TEST2', 'browser', '', at(-100 * HOUR), at(-28 * HOUR), at(-100 * HOUR), 2, 30, 30, false, false, false]);
}

// ── setup ─────────────────────────────────────────────

test('setup: 3つのタブと見出し・設定の最初の値を作り、2回目は作り直さない', () => {
  const { ctx, ss, tab } = load();
  assert.deepEqual(ss.sheets.map(s => s.name), ['設定', '視聴ログ', '視聴者']);
  assert.deepEqual(tab('視聴ログ').data[0], ['時刻', '視聴者ID', 'IDの種類', '経路', 'イベント', '動画の位置（秒）', '端末']);
  assert.equal(tab('視聴者').data[0].length, 14);
  assert.deepEqual(tab('視聴者').data[0].slice(12), ['再公開した時刻', '再公開の期限']);
  const settings = Object.fromEntries(tab('設定').data.slice(1));
  assert.equal(settings['期限（時間）'], 72);
  assert.equal(settings['相談ボタンを出す秒数'], 0);
  assert.equal(settings['合言葉の秒数'], 0);
  assert.equal(settings['相談フォームのURL'], '');
  assert.equal(settings['受け付け中'], true);
  assert.equal(settings['再公開（時間）'], 24);
  assert.equal(settings['uid必須'], false);
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
  assert.deepEqual({ ...res }, { ok: true, status: 'expired', reopenUsed: false });
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

// ── 再公開（実装指示_85） ───────────────────────────────

/** 状態を作る: 'none' 初めて／'within' 期限内／'expired' 期限切れ・未使用／'reopen' 再公開の期限内／'done' 再公開の期限も切れた */
function stateAt(ctx, state) {
  if (state === 'none') return T0;
  ctx.handleOpen_({ uid: 'U1' }, T0);
  if (state === 'within') return at(10 * HOUR);
  if (state === 'expired') return at(80 * HOUR);
  ctx.handleOpen_({ uid: 'U1', again: '1' }, at(80 * HOUR)); // 再公開（80時間目から24時間）
  if (state === 'reopen') return at(90 * HOUR);
  return at(104 * HOUR); // 再公開の期限ちょうど
}

const TABLE = [
  // 状態, again, 期待する status, reopenUsed（期限切れのとき）, 残り時間, ログのイベント
  ['none', true, 'open', undefined, 72 * HOUR, '開いた'],
  ['none', false, 'open', undefined, 72 * HOUR, '開いた'],
  ['within', true, 'open', undefined, 62 * HOUR, '開いた'],
  ['within', false, 'open', undefined, 62 * HOUR, '開いた'],
  ['expired', true, 'open', undefined, 24 * HOUR, '再公開で開いた'],
  ['expired', false, 'expired', false, undefined, '期限切れで開いた'],
  ['reopen', true, 'open', undefined, 14 * HOUR, '開いた'],
  ['reopen', false, 'open', undefined, 14 * HOUR, '開いた'],
  ['done', true, 'expired', true, undefined, '期限切れで開いた'],
  ['done', false, 'expired', true, undefined, '期限切れで開いた'],
];
for (const [state, again, status, reopenUsed, remaining, label] of TABLE) {
  test(`再公開の表: ${state} × again ${again ? 'あり' : 'なし'} → ${status}`, () => {
    const { ctx, viewers, logs } = load();
    const now = stateAt(ctx, state);
    const res = ctx.handleOpen_(again ? { uid: 'U1', again: '1' } : { uid: 'U1' }, now);
    assert.equal(res.status, status);
    if (status === 'expired') {
      assert.equal(res.reopenUsed, reopenUsed);
      assert.equal(res.videoId, undefined);
    } else {
      assert.equal(res.remainingMs, remaining);
    }
    assert.equal(logs().at(-1)[4], label);
    assert.equal(viewers().length, 1);
    assert.equal(logs().filter(r => r[4] === '再公開で開いた').length, ['reopen', 'done'].includes(state) || (state === 'expired' && again) ? 1 : 0);
  });
}

test('再公開: 視聴者タブに再公開した時刻と期限が入る・初めの期限は変えない', () => {
  const { ctx, viewers } = load();
  ctx.handleOpen_({ uid: 'U1' }, T0);
  ctx.handleOpen_({ uid: 'U1', again: '1' }, at(80 * HOUR));
  const v = viewers()[0];
  assert.equal(v[4].getTime(), at(72 * HOUR).getTime());
  assert.equal(v[12].getTime(), at(80 * HOUR).getTime());
  assert.equal(v[13].getTime(), at(104 * HOUR).getTime());
  // 期限内にもう一度 again で開いても、再公開した時刻は変わらない（1回だけ）
  ctx.handleOpen_({ uid: 'U1', again: '1' }, at(90 * HOUR));
  assert.equal(viewers()[0][12].getTime(), at(80 * HOUR).getTime());
});

test('再公開中の残り時間: 再公開の期限から出す・見られる返事の reopenUsed は true', () => {
  const { ctx } = load();
  ctx.handleOpen_({ uid: 'U1' }, T0);
  const r1 = ctx.handleOpen_({ uid: 'U1', again: '1' }, at(80 * HOUR));
  assert.equal(r1.remainingMs, 24 * HOUR);
  assert.equal(r1.reopenUsed, true);
  const r2 = ctx.handleOpen_({ uid: 'U1' }, at(103 * HOUR + 30 * 60 * 1000));
  assert.equal(r2.remainingMs, 30 * 60 * 1000);
  // 最初の期間に見ている人は、終わったら「まだ使える版」
  assert.equal(ctx.handleOpen_({ uid: 'U2' }, T0).reopenUsed, false);
});

test('再公開（時間）が 0: 再公開しない・期限切れは使い終わった版', () => {
  const { ctx, setSetting, viewers } = load();
  setSetting('再公開（時間）', 0);
  ctx.handleOpen_({ uid: 'U1' }, T0);
  const res = ctx.handleOpen_({ uid: 'U1', again: '1' }, at(80 * HOUR));
  assert.deepEqual({ ...res }, { ok: true, status: 'expired', reopenUsed: true });
  assert.equal(viewers()[0][12], '');
  assert.equal(ctx.handleOpen_({ uid: 'U2' }, T0).reopenUsed, true);
});

test('再公開（時間）を変えると、新しく再公開する人から効く', () => {
  const { ctx, setSetting } = load();
  setSetting('再公開（時間）', 6);
  ctx.handleOpen_({ uid: 'U1' }, T0);
  assert.equal(ctx.handleOpen_({ uid: 'U1', again: '1' }, at(80 * HOUR)).remainingMs, 6 * HOUR);
});

test('受け付け中が FALSE: again があっても再公開しない・使い終わった版', () => {
  const { ctx, setSetting, viewers } = load();
  ctx.handleOpen_({ uid: 'U1' }, T0);
  setSetting('受け付け中', false);
  const res = ctx.handleOpen_({ uid: 'U1', again: '1' }, at(80 * HOUR));
  assert.deepEqual({ ...res }, { ok: true, status: 'expired', reopenUsed: true });
  assert.equal(viewers()[0][12], '');
  // 期限内の人も期限切れ
  assert.equal(ctx.handleOpen_({ uid: 'U1' }, at(HOUR)).status, 'expired');
});

test('decideOpen_ 単体: 再公開の期限ちょうどからは見られない', () => {
  const { ctx } = load();
  const settings = { accepting: true, hours: 72, reopenHours: 24 };
  const v = { deadline: at(72 * HOUR), reopenDeadline: at(104 * HOUR) };
  assert.equal(ctx.decideOpen_(v, at(104 * HOUR - 1), settings, false).status, 'open');
  assert.equal(ctx.decideOpen_(v, at(104 * HOUR), settings, true).status, 'expired');
});

// ── migrate85 ────────────────────────────────────────

test('migrate85: 84の本番シートに、足りない設定の行と視聴者の列だけを足す（テストの行はそのまま）', () => {
  const { ctx, tab } = load({ sheet84: true });
  const before = tab('視聴者').data.map(r => [...r]);
  const r = ctx.migrate85();
  assert.deepEqual([...r.settingsAdded], ['再公開（時間）', 'uid必須']);
  assert.deepEqual([...r.viewerColumnsAdded], ['再公開した時刻', '再公開の期限']);
  assert.deepEqual(tab('設定').data.slice(-2), [['再公開（時間）', 24], ['uid必須', false]]);
  assert.equal(tab('設定').data.length, 9);
  assert.deepEqual(tab('視聴者').data[0], [...HEADERS_84, '再公開した時刻', '再公開の期限']);
  assert.deepEqual(tab('視聴者').data.slice(1), before.slice(1));
});

test('migrate85: 2回動かしても同じ・司令塔が書き換えた値を上書きしない', () => {
  const { ctx, tab, setSetting } = load({ sheet84: true });
  ctx.migrate85();
  setSetting('再公開（時間）', 12);
  const snapshot = JSON.stringify([tab('設定').data, tab('視聴者').data, tab('視聴ログ').data]);
  const r = ctx.migrate85();
  assert.equal(r.settingsAdded.length, 0);
  assert.equal(r.viewerColumnsAdded.length, 0);
  assert.equal(JSON.stringify([tab('設定').data, tab('視聴者').data, tab('視聴ログ').data]), snapshot);
});

test('migrate85 のあと: 84のテストの行（期限切れ）が again で再公開できる', () => {
  const { ctx, viewers } = load({ sheet84: true });
  ctx.migrate85();
  const res = ctx.handleOpen_({ bid: 'TEST2', again: '1' }, T0);
  assert.equal(res.status, 'open');
  assert.equal(res.remainingMs, 24 * HOUR);
  assert.equal(viewers()[1][13].getTime(), at(24 * HOUR).getTime());
});

test('migrate85 の前: 再公開の列が無ければ再公開しない（壊れずに期限切れを返す）', () => {
  const { ctx, viewers } = load({ sheet84: true });
  const res = ctx.handleOpen_({ bid: 'TEST2', again: '1' }, T0);
  assert.equal(res.status, 'expired');
  assert.equal(res.reopenUsed, true);
  assert.equal(viewers()[1].length, 12);
});

test('列は見出しの名前で読む: 列の順番を入れ替えても動く', () => {
  const { ctx, tab } = load();
  const sheet = tab('視聴者');
  // 「経路」と「相談ボタンを押した」の列を入れ替え、知らない列を間に足す
  sheet.data[0] = ['IDの種類', '視聴者ID', 'メモ', '相談ボタンを押した', '初めて開いた時刻', '期限', '最後に開いた時刻', '開いた回数',
    '見た一番先の位置（秒）', '見た合計（秒）', '合言葉の位置を通過', '最後まで見た', '経路', '再公開の期限', '再公開した時刻'];
  ctx.handleOpen_({ uid: 'U1', src: 'taisei' }, T0);
  sheet.data[1][2] = '司令塔のメモ';
  ctx.handleEvent_({ event: 'consult', pos: 10, t: 0, uid: 'U1', sid: 's' }, T0);
  ctx.handleOpen_({ uid: 'U1', again: '1' }, at(80 * HOUR));
  const row = sheet.data[1];
  assert.deepEqual([row[0], row[1], row[2], row[3], row[12]], ['uid', 'U1', '司令塔のメモ', true, 'taisei']);
  assert.equal(row[13].getTime(), at(104 * HOUR).getTime());
  assert.equal(row[14].getTime(), at(80 * HOUR).getTime());
  // 視聴ログも見出しの名前で書く
  const log = tab('視聴ログ');
  log.data[0] = ['イベント', '時刻', '視聴者ID', 'IDの種類', '経路', '動画の位置（秒）', '端末'];
  ctx.handleOpen_({ uid: 'U1' }, at(81 * HOUR));
  assert.deepEqual(log.data.at(-1).slice(0, 1).concat(log.data.at(-1).slice(2, 4)), ['開いた', 'U1', 'uid']);
});

// ── uid必須 ─────────────────────────────────────────

test('uid必須が FALSE（最初の値）: uid が無い人もブラウザのIDで見られる（今までどおり）', () => {
  const { ctx, viewers } = load();
  assert.equal(ctx.handleOpen_({ bid: 'b1' }, T0).status, 'open');
  assert.equal(viewers().length, 1);
});

test('uid必須が TRUE: uid が無い人には動画を出さず、視聴者の行を作らない', () => {
  const { ctx, setSetting, viewers, logs } = load();
  setSetting('uid必須', true);
  const res = ctx.handleOpen_({ bid: 'b1', src: 'taisei', device: 'スマホ' }, T0);
  assert.deepEqual({ ...res }, { ok: true, status: 'needUid' });
  assert.equal(res.videoId, undefined);
  assert.equal(viewers().length, 0);
  assert.deepEqual(logs().at(-1).slice(1), ['b1', 'browser', 'taisei', 'uid無しで開いた', '', 'スマホ']);
});

test('uid必須が TRUE: 置き換え失敗の uid・使えない形の uid も「uid が無い」扱い', () => {
  const { ctx, setSetting, viewers } = load();
  setSetting('uid必須', 'TRUE');
  assert.equal(ctx.handleOpen_({ uid: '#_userParamId_#', bid: 'b1' }, T0).status, 'needUid');
  assert.equal(ctx.handleOpen_({ uid: '=1+1', bid: 'b1' }, T0).status, 'needUid');
  assert.equal(ctx.handleOpen_({ uid: '', bid: 'b1', again: '1' }, T0).status, 'needUid');
  assert.equal(viewers().length, 0);
});

test('uid必須が TRUE: uid がある人は今までどおり（初回・期限切れ・再公開）', () => {
  const { ctx, setSetting } = load();
  setSetting('uid必須', true);
  assert.equal(ctx.handleOpen_({ uid: 'U1' }, T0).status, 'open');
  assert.equal(ctx.handleOpen_({ uid: 'U1' }, at(80 * HOUR)).status, 'expired');
  assert.equal(ctx.handleOpen_({ uid: 'U1', again: '1' }, at(80 * HOUR)).remainingMs, 24 * HOUR);
});

test('uid必須が TRUE: 前にブラウザのIDで見ていた人も、uid が無ければ見られない', () => {
  const { ctx, setSetting, viewers } = load();
  ctx.handleOpen_({ bid: 'b1' }, T0);
  setSetting('uid必須', true);
  assert.equal(ctx.handleOpen_({ bid: 'b1' }, at(HOUR)).status, 'needUid');
  assert.equal(viewers()[0][6], 1); // 開いた回数は増やさない
});

test('uid必須が TRUE: uid の無い人の再生の記録は受け取らない', () => {
  const { ctx, setSetting, logs } = load();
  setSetting('uid必須', true);
  const n = logs().length;
  assert.equal(ctx.handleEvent_({ event: 'play', pos: 0, t: 0, bid: 'b1', sid: 's' }, T0).ignored, 'need_uid');
  assert.equal(logs().length, n);
  ctx.handleOpen_({ uid: 'U1' }, T0);
  assert.equal(ctx.handleEvent_({ event: 'play', pos: 0, t: 0, uid: 'U1', sid: 's' }, T0).ok, true);
  assert.equal(logs().at(-1)[4], '再生');
});

test('migrateUidRequired: 85まで入ったシートに「uid必須」＝FALSE だけを足す・2回動かしても同じ', () => {
  const { ctx, tab } = load({ sheet84: true });
  ctx.migrate85();
  // 85まで入った本番のシート（uid必須の行がまだ無い）にする
  const settings = tab('設定');
  settings.data.splice(settings.data.findIndex(r => r[0] === 'uid必須'), 1);
  const viewersBefore = JSON.stringify(tab('視聴者').data);
  const r1 = ctx.migrateUidRequired();
  assert.deepEqual([...r1.settingsAdded], ['uid必須']);
  assert.equal(r1.viewerColumnsAdded.length, 0);
  assert.deepEqual(settings.data.at(-1), ['uid必須', false]);
  assert.equal(JSON.stringify(tab('視聴者').data), viewersBefore);
  // 司令塔が TRUE にしたあとで動かしても、書き換えない
  settings.data.at(-1)[1] = true;
  const snapshot = JSON.stringify(settings.data);
  const r2 = ctx.migrateUidRequired();
  assert.equal(r2.settingsAdded.length, 0);
  assert.equal(JSON.stringify(settings.data), snapshot);
});

test('migrateUidRequired の前: 「uid必須」の行が無ければ FALSE として動く', () => {
  const { ctx } = load({ sheet84: true });
  assert.equal(ctx.handleOpen_({ bid: 'TEST9' }, T0).status, 'open');
});
