/**
 * ページ（index.html）の小さな部品のテスト。<script> の中身を vm で読む（画面は動かさない）。
 */
import test from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import vm from 'node:vm';

const HTML = readFileSync(new URL('../index.html', import.meta.url), 'utf8');
const SCRIPT = HTML.match(/<script>([\s\S]*?)<\/script>/)[1];

function load() {
  const ctx = { window: { __ZV_TEST__: true }, URLSearchParams, Uint8Array, Math, crypto: globalThis.crypto };
  vm.createContext(ctx);
  vm.runInContext(SCRIPT, ctx);
  return ctx;
}

function memoryStorage() {
  const m = {};
  return { getItem: k => m[k] ?? null, setItem: (k, v) => { m[k] = String(v); }, removeItem: k => { delete m[k]; } };
}

test('ページ: 動画IDを HTML に直接書いていない', () => {
  assert.ok(!/videoId\s*:\s*['"][A-Za-z0-9_-]{11}['"]/.test(HTML));
  assert.ok(!HTML.includes('M7lc1UVf-VE'));
});

test('残り時間の文言: 日＋時間 → 時間 → 分', () => {
  const { formatRemaining } = load();
  const H = 3600000;
  assert.equal(formatRemaining(72 * H), 'あと3日0時間');
  assert.equal(formatRemaining(71 * H + 59 * 60000), 'あと2日23時間');
  assert.equal(formatRemaining(24 * H), 'あと1日0時間');
  assert.equal(formatRemaining(24 * H - 1), 'あと23時間');
  assert.equal(formatRemaining(H), 'あと1時間');
  assert.equal(formatRemaining(H - 1), 'あと59分');
  assert.equal(formatRemaining(30 * 1000), 'あと1分');
});

test('視聴者ID: uid があれば uid・置き換え失敗や空はブラウザのID', () => {
  const { resolveViewer } = load();
  assert.deepEqual({ ...resolveViewer('?src=taisei&uid=U123', memoryStorage()) }, { uid: 'U123' });
  // 置き換えに失敗すると「#」から後ろは URL のハッシュになり、uid は空になる
  const s = memoryStorage();
  const a = resolveViewer('?src=taisei&uid=', s);
  assert.match(a.bid, /^[0-9a-f]{24}$/);
  assert.equal(resolveViewer('?uid=%23_userParamId_%23', s).bid, a.bid); // 同じブラウザなら同じID
});

test('視聴者ID: localStorage が使えなくても落ちず、その回だけのIDで動く', () => {
  const { resolveViewer, safeStorage } = load();
  const broken = { getItem() { throw new Error('denied'); }, setItem() { throw new Error('denied'); } };
  assert.match(resolveViewer('', broken).bid, /^[0-9a-f]{24}$/);
  assert.equal(safeStorage({ get localStorage() { throw new Error('denied'); } }), null);
});

test('端末: スマホ／PC', () => {
  const { deviceOf } = load();
  assert.equal(deviceOf('Mozilla/5.0 (iPhone; CPU iPhone OS 17_0 like Mac OS X) Line/13.0'), 'スマホ');
  assert.equal(deviceOf('Mozilla/5.0 (Linux; Android 14) Mobile Safari'), 'スマホ');
  assert.equal(deviceOf('Mozilla/5.0 (Windows NT 10.0; Win64; x64) Chrome/120'), 'PC');
});

// ── 再公開（実装指示_85） ───────────────────────────────

test('again=1 を読む', () => {
  const { readAgain } = load();
  assert.equal(readAgain('?src=taisei&uid=U1&again=1'), true);
  assert.equal(readAgain('?src=taisei&uid=U1'), false);
  assert.equal(readAgain('?again=0'), false);
});

test('期限切れの画面: reopenUsed が false のときだけ「まだ使える版」', () => {
  const { expiredVariant } = load();
  assert.equal(expiredVariant({ reopenUsed: false }), 'canReopen');
  assert.equal(expiredVariant({ reopenUsed: true }), 'used');
  assert.equal(expiredVariant({}), 'used');
  assert.equal(expiredVariant(null), 'used');
});

test('期限切れの文言が指示書のとおり入っている', () => {
  const text = HTML.replace(/<br>/g, '\n').replace(/<[^>]+>/g, '');
  for (const line of [
    'この動画の公開は終了しました',
    '受け取ってから3日がたったため、見られなくなっています。',
    '「やっぱり見たい！」という方へ。', 'LINEに「動画視聴希望」と送ってください。', '1回だけ、もう一度見られるようにします。',
    'もう一度の公開も、終了しました。', 'このあともLINEで、Shopee物販のお役立ち情報をお届けしていきます。',
  ]) assert.ok(text.includes(line), line);
});
