/*
 * 台帳屋 異動・退職チェッカー for kintone
 * 一覧画面に「異動・退職チェック」ボタンを追加し、選んだアプリの設定（アクセス権・プロセス管理・通知・
 * フィールドの初期値・一覧の絞り込みなど）に出てくるユーザー・組織・グループを一覧にする。
 * もういないユーザーや、利用停止中のユーザーの指定も見つける。
 * 読み取るのはアプリの設定だけで、データはすべてお客様のkintoneの中で処理し、外部には送らない。
 * 集める部分は共通部品 js/refs.js（window.DaichoyaRefs）を使う。
 */
(function () {
  'use strict';

  var Refs = typeof module !== 'undefined' && module.exports ? require('../../../common/js/refs.js') : window.DaichoyaRefs;

  // ===== 画面に依存しない部分（テスト対象） =====

  // 読み取る設定と、そのAPI
  var SETTINGS = [
    ['appAcl', '/k/v1/app/acl.json', 'アプリのアクセス権'],
    ['recordAcl', '/k/v1/record/acl.json', 'レコードのアクセス権'],
    ['fieldAcl', '/k/v1/field/acl.json', 'フィールドのアクセス権'],
    ['status', '/k/v1/app/status.json', 'プロセス管理'],
    ['general', '/k/v1/app/notifications/general.json', 'アプリの条件通知'],
    ['perRecord', '/k/v1/app/notifications/perRecord.json', 'レコードの条件通知'],
    ['reminder', '/k/v1/app/notifications/reminder.json', 'リマインダーの条件通知'],
    ['fields', '/k/v1/app/form/fields.json', 'フィールド'],
    ['views', '/k/v1/app/views.json', '一覧'],
    ['reports', '/k/v1/app/reports.json', 'グラフ'],
    ['actions', '/k/v1/app/actions.json', 'アクション']
  ];

  function errorText(error) {
    var code = error && error.code;
    if (code === 'CB_NO02' || code === 'GAIA_NO01') return '権限がありません';
    return (error && error.message) || '原因不明のエラー';
  }

  // results: { key: { ok, value } | { ok: false, error } } → 集めた指定と、読めなかった設定
  function analyze(results) {
    var settings = {};
    var errors = [];
    SETTINGS.forEach(function (s) {
      var r = results[s[0]];
      if (r && r.ok) settings[s[0]] = r.value;
      else { settings[s[0]] = null; errors.push(s[2] + '（' + errorText(r && r.error) + '）'); }
    });
    return { refs: Refs.collect(settings), sole: Refs.soleManager(settings.appAcl), errors: errors, allFailed: errors.length === SETTINGS.length };
  }

  function summaryLines(list, sole, errors) {
    var sum = Refs.summary(list);
    var lines = [];
    if (!list.length) lines.push({ alert: false, text: ((errors && errors.length) ? '読み取れた設定' : 'このアプリの設定') + 'には、ユーザー・組織・グループを指定しているところがありません。' });
    else {
      var count = function (n) { return n < 10 ? n + 'つ' : n + '個'; };
      var parts = [];
      if (sum.users) parts.push('ユーザー' + sum.users + '人');
      if (sum.organizations) parts.push('組織' + count(sum.organizations));
      if (sum.groups) parts.push('グループ' + count(sum.groups));
      lines.push({ alert: false, text: 'このアプリの設定には、' + parts.join('・') + 'が出てきます（全' + list.length + 'か所）。' });
    }
    if (sum.missing) lines.push({ alert: true, text: '見つからないユーザー・組織・グループの指定が' + sum.missing + '件あります。削除されたか、コードが変わった可能性があります。' });
    if (sum.stopped) lines.push({ alert: true, text: '利用停止中のユーザーの指定が' + sum.stopped + '件あります。' });
    if (sole) lines.push({ alert: true, text: 'アプリ管理の権限が、ユーザー1人（' + sole + '）だけに付いています。異動や退職の前に、ほかの人やグループにも付けておくと安心です。' });
    if (errors.length) lines.push({ alert: false, text: '読み取れなかった設定：' + errors.join('、') });
    return lines;
  }

  function filterRows(list, word) {
    var w = String(word || '').trim().toLowerCase();
    if (!w) return list;
    return list.filter(function (r) { return (r.code + '\n' + r.name).toLowerCase().indexOf(w) >= 0; });
  }

  function csvEscape(value) {
    var s = value === undefined || value === null ? '' : String(value);
    if (/[",\r\n]/.test(s)) s = '"' + s.replace(/"/g, '""') + '"';
    return s;
  }

  var CSV_HEADER = ['状態', '種類', '名前', 'コード', '設定の場所', '内容'];

  // Excelで文字化けしないよう、BOM付きUTF-8・CRLF改行で作る
  function buildCsv(list) {
    var lines = [CSV_HEADER].concat(list.map(function (r) { return [r.state, r.typeLabel, r.name, r.code, r.place, r.detail]; }));
    return '﻿' + lines.map(function (row) { return row.map(csvEscape).join(','); }).join('\r\n') + '\r\n';
  }

  // ===== 提出用のExcel（共通部品 js/xlsx.js。見出しの色・固定行・フィルター・A4の印刷設定つき） =====
  var Xlsx = typeof module !== 'undefined' && module.exports ? require('../../../common/js/xlsx.js') : (typeof window !== 'undefined' ? window.DaichoyaXlsx : null);
  function excelAtText(d) {
    var pad = function (n) { return ('0' + n).slice(-2); };
    return d.getFullYear() + '-' + pad(d.getMonth() + 1) + '-' + pad(d.getDate()) + ' ' + pad(d.getHours()) + ':' + pad(d.getMinutes());
  }
  function excelSub(info, product) {
    return 'kintone環境：' + (info.host || '') + '　／　作成日時：' + excelAtText(info.at || new Date()) + '　／　作成：台帳屋 ' + product;
  }

  // info: { app: アプリ名, host: kintoneの環境, at: 作成日時 }
  // Excelの先頭に入れる、結果のまとめの1文
  function excelLead(list) {
    if (!list.length) return 'ユーザー・組織・グループを指定している設定はありません。';
    var missing = list.filter(function (r) { return r.state === '見つかりません'; }).length;
    var stopped = list.filter(function (r) { return r.state === '利用停止中'; }).length;
    var parts = [];
    if (missing) parts.push('見つからない指定が' + missing + '件');
    if (stopped) parts.push('利用停止中のユーザーの指定が' + stopped + '件');
    return '設定に出てくるユーザー・組織・グループは全' + list.length + 'か所です。' +
      (parts.length ? parts.join('、') + 'あります。' : '見つからない指定や、利用停止中のユーザーの指定はありません。');
  }

  function buildExcel(list, info) {
    return Xlsx.listBook({
      title: 'ユーザー・組織・グループの指定', target: info.app, sub: excelSub(info, '異動・退職チェッカー'), sheet: 'ユーザーの指定', lead: excelLead(list),
      header: CSV_HEADER, rows: list.map(function (r) { return [r.state, r.typeLabel, r.name, r.code, r.place, r.detail]; }),
      styles: list.map(function (r) { return r.state === '見つかりません' ? Xlsx.STYLE.important : r.state === '利用停止中' ? Xlsx.STYLE.warn : Xlsx.STYLE.body; }),
      widths: [15, 10, 18, 22, 30, 60], creator: '台帳屋 異動・退職チェッカー',
      note: '赤の行：もういないユーザーや、なくなった組織・グループ。黄の行：利用停止中のユーザー。'
    });
  }

  function safeFileName(name) {
    var s = String(name || '').replace(/[\\/:*?"<>|]/g, '_').trim();
    return s || 'app';
  }

  var Core = { SETTINGS: SETTINGS, analyze: analyze, summaryLines: summaryLines, filterRows: filterRows, buildCsv: buildCsv, buildExcel: buildExcel, safeFileName: safeFileName };

  if (typeof module !== 'undefined' && module.exports) {
    module.exports = Core;
    return;
  }
  if (typeof kintone === 'undefined' || !Refs) return;

  // ===== 画面（kintone上でだけ動く部分） =====

  var P = 'dcy-tc';
  var current = { appName: '', list: null };

  function el(tag, attrs, children) {
    var node = document.createElement(tag);
    Object.keys(attrs || {}).forEach(function (key) {
      if (key === 'text') node.textContent = attrs[key];
      else node.setAttribute(key, attrs[key]);
    });
    (children || []).forEach(function (child) {
      node.appendChild(typeof child === 'string' ? document.createTextNode(child) : child);
    });
    return node;
  }

  function api(path, params) { return kintone.api(kintone.api.url(path, true), 'GET', params); }
  // cybozu.com の User API（ユーザー・組織・グループの確認）は、kintoneのパスを通さずに呼ぶ
  function userApi(path, params) { return kintone.api(path, 'GET', params); }
  function settle(promise) {
    return promise.then(function (value) { return { ok: true, value: value }; }, function (error) { return { ok: false, error: error }; });
  }

  function getAllApps() {
    var apps = [];
    function next(offset) {
      return api('/k/v1/apps.json', { limit: 100, offset: offset }).then(function (resp) {
        apps = apps.concat(resp.apps);
        return resp.apps.length === 100 ? next(offset + 100) : apps;
      });
    }
    return next(0);
  }

  function inspect(appId) {
    return Promise.all(SETTINGS.map(function (s) { return settle(api(s[1], { app: appId })); })).then(function (list) {
      var results = {};
      SETTINGS.forEach(function (s, i) { results[s[0]] = list[i]; });
      var found = analyze(results);
      if (found.allFailed) return found;
      return Refs.lookup(userApi, found.refs).then(function (dir) {
        found.list = Refs.rows(found.refs, dir);
        return found;
      });
    });
  }

  function renderTable(container, list) {
    container.innerHTML = '';
    var headers = ['状態', '種類', '名前', 'コード', '設定の場所', '内容'];
    var thead = el('thead', {}, [el('tr', {}, headers.map(function (h) { return el('th', { text: h }); }))]);
    var body = list.length ? list.map(function (r) {
      return el('tr', {}, [
        el('td', { class: P + '-state-' + r.level, text: r.state }),
        el('td', { text: r.typeLabel }),
        el('td', { text: r.name || '－' }),
        el('td', { text: r.code }),
        el('td', { text: r.place }),
        el('td', { class: P + '-detail', text: r.detail })
      ]);
    }) : [el('tr', {}, [el('td', { colspan: '6', class: P + '-empty', text: '該当するものはありません' })])];
    container.appendChild(el('table', { class: P + '-table' }, [thead, el('tbody', {}, body)]));
  }


  function saveExcel(bytes, fileName) {
    var blob = new Blob([bytes], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    var url = URL.createObjectURL(blob);
    var a = el('a', { href: url, download: fileName });
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    setTimeout(function () { URL.revokeObjectURL(url); }, 1000);
  }
  // 「Excelで保存」は、「CSVで書き出す」と同じときに押せるようにする
  function syncButton(from, to) {
    to.disabled = from.disabled;
    new MutationObserver(function () { to.disabled = from.disabled; }).observe(from, { attributes: true, attributeFilter: ['disabled'] });
  }

  function downloadExcel() {
    if (!current.list) return;
    var d = new Date(); var stamp = d.getFullYear() + ('0' + (d.getMonth() + 1)).slice(-2) + ('0' + d.getDate()).slice(-2);
    saveExcel(buildExcel(current.list, { app: current.displayName, host: location.hostname, at: d }), 'ユーザーの指定_' + current.appName + '_' + stamp + '.xlsx');
  }

  function downloadCsv() {
    if (!current.list) return;
    var blob = new Blob([buildCsv(current.list)], { type: 'text/csv;charset=utf-8' });
    var url = URL.createObjectURL(blob);
    var d = new Date();
    var stamp = d.getFullYear() + ('0' + (d.getMonth() + 1)).slice(-2) + ('0' + d.getDate()).slice(-2);
    var a = el('a', { href: url, download: 'ユーザーの指定_' + current.appName + '_' + stamp + '.csv' });
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    setTimeout(function () { URL.revokeObjectURL(url); }, 1000);
  }

  // 有料版のご案内（画面の下に小さく1行）。アプリの一覧を読んだら、この環境のアプリの数を入れる
  var moreText = null;
  function moreLink() {
    moreText = document.createTextNode('すべてのアプリを、まとめて無料で点検できます：');
    return el('p', { class: P + '-more' }, [
      moreText,
      el('a', { href: 'https://daichoya.github.io/pro.html#free', target: '_blank', rel: 'noopener', text: '棚卸し・変更監視パック（件数の点検は無料。一覧とExcelは30日間無料で試せます）' })
    ]);
  }
  function moreCount(n) { if (moreText && n > 1) moreText.textContent = 'この環境の' + n + '個のアプリを、まとめて無料で点検できます：'; }

  function openModal() {
    if (document.getElementById(P + '-overlay')) return;
    var appNames = {};
    var select = el('select', { class: P + '-select', 'aria-label': '調べるアプリ' });
    var runBtn = el('button', { type: 'button', class: P + '-btn ' + P + '-primary', text: 'ユーザーの指定を調べる' });
    var excelBtn = el('button', { type: 'button', class: P + '-btn', text: 'Excelで保存', disabled: 'disabled' });
    var csvBtn = el('button', { type: 'button', class: P + '-btn', text: 'CSVで書き出す', disabled: 'disabled' });
    var closeBtn = el('button', { type: 'button', class: P + '-close', 'aria-label': '閉じる', text: '×' });
    var status = el('p', { class: P + '-status', role: 'status' });
    var summary = el('div', { class: P + '-summary', hidden: 'hidden' });
    var filter = el('input', { type: 'search', class: P + '-filter', placeholder: '名前・コードで絞り込む', 'aria-label': '名前・コードで絞り込む' });
    var tools = el('div', { class: P + '-tools', hidden: 'hidden' }, [filter]);
    var result = el('div', { class: P + '-result' });
    var dialog = el('div', { class: P + '-dialog', role: 'dialog', 'aria-modal': 'true', 'aria-label': '異動・退職チェッカー' }, [
      el('div', { class: P + '-head' }, [el('h2', { text: '異動・退職チェッカー' }), closeBtn]),
      el('p', { class: P + '-lead', text: 'アプリの設定に出てくるユーザー・組織・グループを一覧にします。異動・退職・組織変更の前に、見直す場所を確かめられます。' }),
      el('div', { class: P + '-controls' }, [select, runBtn, excelBtn, csvBtn]),
      status, summary, tools, result,
      el('p', { class: P + '-note', text: '調べる設定：アクセス権（アプリ・レコード・フィールド）、プロセス管理、通知（アプリ・レコード・リマインダー）、ユーザー選択などのフィールドの初期値と選択肢、一覧・グラフ・アクセス権・通知の条件、アクション。プラグインの設定やJavaScriptの中は調べません。' }),
      moreLink()
    ]);
    var overlay = el('div', { id: P + '-overlay', class: P + '-overlay' }, [dialog]);

    function close() {
      document.removeEventListener('keydown', onKey);
      overlay.parentNode && overlay.parentNode.removeChild(overlay);
    }
    function onKey(e) { if (e.key === 'Escape') close(); }

    closeBtn.addEventListener('click', close);
    overlay.addEventListener('click', function (e) { if (e.target === overlay) close(); });
    document.addEventListener('keydown', onKey);
    csvBtn.addEventListener('click', downloadCsv);
    excelBtn.addEventListener('click', downloadExcel);
    syncButton(csvBtn, excelBtn);
    filter.addEventListener('input', function () { if (current.list) renderTable(result, filterRows(current.list, filter.value)); });

    runBtn.addEventListener('click', function () {
      var appId = select.value;
      if (!appId) return;
      runBtn.disabled = true;
      csvBtn.disabled = true;
      summary.hidden = true;
      tools.hidden = true;
      result.innerHTML = '';
      current.list = null;
      status.textContent = '設定を読み込み中…';
      inspect(appId).then(function (found) {
        if (found.allFailed) {
          status.textContent = 'このアプリの設定を読み取れませんでした。そのアプリの管理権限が必要です（' + found.errors[0] + '）。';
          return;
        }
        current.appName = safeFileName(appNames[appId] || ('app' + appId));
        current.displayName = appNames[appId] || ('アプリID ' + appId);
        current.list = found.list;
        status.textContent = '';
        summary.innerHTML = '';
        summaryLines(found.list, found.sole, found.errors).forEach(function (line) {
          summary.appendChild(el('p', line.alert ? { class: P + '-alert', text: line.text } : { text: line.text }));
        });
        summary.hidden = false;
        tools.hidden = !found.list.length;
        filter.value = '';
        renderTable(result, found.list);
        csvBtn.disabled = !found.list.length;
      }).catch(function (error) {
        status.textContent = '読み込みに失敗しました（' + errorText(error) + '）。';
      }).then(function () {
        runBtn.disabled = false;
      });
    });

    document.body.appendChild(overlay);
    status.textContent = 'アプリの一覧を読み込み中…';
    getAllApps().then(function (apps) {
      moreCount(apps.length);
      var thisApp = String(kintone.app.getId());
      apps.forEach(function (app) {
        appNames[app.appId] = app.name;
        var option = el('option', { value: app.appId, text: app.name + '（ID: ' + app.appId + '）' });
        if (String(app.appId) === thisApp) option.selected = true;
        select.appendChild(option);
      });
      status.textContent = 'アプリを選んで「ユーザーの指定を調べる」を押してください。';
    }).catch(function (error) {
      status.textContent = 'アプリの一覧を取得できませんでした（' + errorText(error) + '）。';
    });
  }

  kintone.events.on('app.record.index.show', function (event) {
    var space = kintone.app.getHeaderMenuSpaceElement();
    if (!space || document.getElementById(P + '-open')) return event;
    var button = el('button', { id: P + '-open', type: 'button', class: P + '-btn', text: '異動・退職チェック' });
    button.addEventListener('click', openModal);
    space.appendChild(button);
    return event;
  });
})();
