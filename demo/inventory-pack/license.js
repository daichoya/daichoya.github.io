/*
 * 台帳屋 棚卸し・変更監視パック：ライセンスキーの確認
 * キーは台帳屋の鍵で署名してあり、ここに入れた公開鍵で確かめる。台帳屋のサーバーとは通信しない。
 * キーの形：DCY1-<内容（base64url のJSON）>.<署名（base64url）>
 */
(function (root) {
  'use strict';

  var PUBLIC_KEY = 'MIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEAq9XleCkCAyUymAhcuM6N8yyQRCPBc0AnY9aSTzg1PIyF60mDplUXvLhOxydeNNXdeL5FE2SK8vaLs+ODRWnozWyfR8JHuWFEvPTnu47BDwOr4231VUdLVYKQr+yrHkHMY4rlOYOw7Ws8zPaO5LG8j9Y3cLyldEFvjFysL32uaUKPdpiNc8ht4lsbZ+N3Zw8yb5k3eVrEmZyYfn++ES4jxaJTTThoyE93CJZ5Xtg8Ads/55ihk5cTnA+QqyS3t3lWo6Yw1OCaQyNvJWbaQpXu6BV5iCfxharM8eo7PCxAz9gDacsFWBRXR+dq6Q1PjGYy5f2P9M3yCGTh0bVhCpCRkwIDAQAB';
  var DEFAULT_PRODUCT = 'inventory';
  var ALL_PRODUCTS = 'all';  // すべての有料機能が使えるキー（今後のまとめ買いのプラン用）
  var GRACE_DAYS = 7;   // 期限が過ぎても、この日数は使える（月額の更新が少し遅れても止まらないように）
  var KIND_LABELS = { trial: '無料試用', month: '月額', year: '年額', demo: 'デモ' };

  function b64urlToBytes(text) {
    var s = String(text).replace(/-/g, '+').replace(/_/g, '/');
    while (s.length % 4) s += '=';
    var bin = atob(s);
    var out = new Uint8Array(bin.length);
    for (var i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i);
    return out;
  }

  // kintoneの環境名をそろえる。セキュアアクセス（○○.s.cybozu.com）も同じ環境として扱う
  function normalizeDomain(host) {
    var parts = String(host || '').trim().toLowerCase().split('.');
    if (parts.length > 3 && parts[1] === 's') parts.splice(1, 1);
    return parts.join('.');
  }

  // 日本時間の日付（YYYY-MM-DD）
  function todayJst(now) {
    var d = new Date((now ? now.getTime() : Date.now()) + 9 * 3600 * 1000);
    return d.toISOString().slice(0, 10);
  }

  function daysBetween(from, to) {
    return Math.round((Date.parse(to + 'T00:00:00Z') - Date.parse(from + 'T00:00:00Z')) / 86400000);
  }

  function parse(key) {
    var m = /^DCY1-([A-Za-z0-9_-]+)\.([A-Za-z0-9_-]+)$/.exec(String(key || '').replace(/\s+/g, ''));
    if (!m) return null;
    try {
      var payload = JSON.parse(new TextDecoder().decode(b64urlToBytes(m[1])));
      return { body: m[1], signature: b64urlToBytes(m[2]), payload: payload };
    } catch (e) {
      return null;
    }
  }

  function subtle() {
    var c = (typeof crypto !== 'undefined' && crypto.subtle) ? crypto : (typeof require === 'function' ? require('node:crypto').webcrypto : null);
    return c && c.subtle;
  }

  var keyPromise = null;
  function publicKey() {
    if (!keyPromise) {
      keyPromise = subtle().importKey('spki', b64urlToBytes(PUBLIC_KEY.replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '')),
        { name: 'RSASSA-PKCS1-v1_5', hash: 'SHA-256' }, false, ['verify']);
    }
    return keyPromise;
  }

  // 結果：{ ok, state: 'valid'|'grace'|'expired'|'domain'|'invalid'|'none', message, payload, daysLeft }
  // product：確かめる製品。キーの製品がそれと同じか、すべての製品用（all）なら使える
  function check(key, host, now, product) {
    var want = product || DEFAULT_PRODUCT;
    if (!key || !String(key).trim()) {
      return Promise.resolve({ ok: false, state: 'none', message: 'ライセンスキーが設定されていません。' });
    }
    var parsed = parse(key);
    if (!parsed || (parsed.payload.p !== want && parsed.payload.p !== ALL_PRODUCTS)) {
      return Promise.resolve({ ok: false, state: 'invalid', message: 'ライセンスキーの形が正しくありません。メールに記載のキーを、そのまま貼り付けてください。' });
    }
    return publicKey().then(function (k) {
      return subtle().verify('RSASSA-PKCS1-v1_5', k, parsed.signature, new TextEncoder().encode(parsed.body));
    }).then(function (verified) {
      var p = parsed.payload;
      if (!verified) return { ok: false, state: 'invalid', message: 'ライセンスキーを確認できませんでした。メールに記載のキーを、そのまま貼り付けてください。' };
      if (normalizeDomain(p.d) !== normalizeDomain(host)) {
        return { ok: false, state: 'domain', payload: p, message: 'このライセンスキーは「' + p.d + '」用です。お使いの環境（' + normalizeDomain(host) + '）では使えません。' };
      }
      var today = todayJst(now);
      var left = daysBetween(today, p.e);
      var kind = KIND_LABELS[p.t] || p.t;
      if (left >= 0) {
        if (p.t === 'demo') return { ok: true, state: 'valid', payload: p, daysLeft: left, message: 'デモ用（見本データでお試しいただけます）' };
        return { ok: true, state: 'valid', payload: p, daysLeft: left, message: kind + '・' + p.e + 'まで有効' + (p.a ? '（アプリ' + p.a + '個までのプラン）' : '') };
      }
      if (-left <= GRACE_DAYS) {
        return { ok: true, state: 'grace', payload: p, daysLeft: left, message: kind + '・' + p.e + 'で期限が切れました。新しいキーを設定してください（' + (GRACE_DAYS + left) + '日後に使えなくなります）' };
      }
      return { ok: false, state: 'expired', payload: p, daysLeft: left, message: kind + 'のライセンスは' + p.e + 'で期限が切れています。' };
    }).catch(function () {
      return { ok: false, state: 'invalid', message: 'ライセンスキーを確認できませんでした。' };
    });
  }

  var api = { check: check, parse: parse, normalizeDomain: normalizeDomain, todayJst: todayJst, GRACE_DAYS: GRACE_DAYS };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.DaichoyaLicense = api;
})(typeof window !== 'undefined' ? window : this);
