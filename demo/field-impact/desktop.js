/*
 * 台帳屋 フィールド影響チェッカー for kintone
 * 一覧画面に「フィールドの影響」ボタンを追加し、選んだフィールドが使われている場所
 * （計算式・一覧・グラフ・アクセス権・プロセス管理・通知・アクション・JavaScript、
 *   ほかのアプリのルックアップ・関連レコード一覧・アクション）を一覧にする。
 * 読み取るのはアプリの設定とカスタマイズのファイルだけで、データはすべてお客様のkintoneの中で処理し、外部には送らない。
 * 探す部分は共通部品 js/fieldrefs.js（window.DaichoyaFieldRefs）を使う。
 */
(function () {
  'use strict';

  var F = typeof module !== 'undefined' && module.exports ? require('../../../common/js/fieldrefs.js') : window.DaichoyaFieldRefs;

  // ===== 画面に依存しない部分（テスト対象） =====

  var SETTINGS = [
    ['fields', '/k/v1/app/form/fields.json', 'フィールド'],
    ['views', '/k/v1/app/views.json', '一覧'],
    ['reports', '/k/v1/app/reports.json', 'グラフ'],
    ['recordAcl', '/k/v1/record/acl.json', 'レコードのアクセス権'],
    ['fieldAcl', '/k/v1/field/acl.json', 'フィールドのアクセス権'],
    ['status', '/k/v1/app/status.json', 'プロセス管理'],
    ['general', '/k/v1/app/notifications/general.json', 'アプリの条件通知'],
    ['perRecord', '/k/v1/app/notifications/perRecord.json', 'レコードの条件通知'],
    ['reminder', '/k/v1/app/notifications/reminder.json', 'リマインダーの条件通知'],
    ['actions', '/k/v1/app/actions.json', 'アクション'],
    ['appSettings', '/k/v1/app/settings.json', 'アプリの一般設定'],
    ['customize', '/k/v1/app/customize.json', 'JavaScript・CSSでカスタマイズ']
  ];

  function errorText(error) {
    var code = error && error.code;
    if (code === 'CB_NO02' || code === 'GAIA_NO01') return '権限がありません';
    return (error && error.message) || '原因不明のエラー';
  }

  // カスタマイズのうち、中身を読めるJavaScriptのファイル（URLで読み込むものは調べない）
  function scriptFiles(customize) {
    var out = { files: [], urls: [] };
    if (!customize) return out;
    [['desktop', 'PC'], ['mobile', 'モバイル']].forEach(function (pair) {
      (((customize[pair[0]] || {}).js) || []).forEach(function (x) {
        if (x.type === 'FILE' && x.file) out.files.push({ screen: pair[1] + '・JavaScript', name: x.file.name, fileKey: x.file.fileKey, size: Number(x.file.size || 0) });
        else if (x.type === 'URL') out.urls.push(pair[1] + '：' + x.url);
      });
    });
    return out;
  }

  function summaryLines(target, list, info) {
    var name = '「' + target.label + '」';
    var here = list.filter(function (r) { return r.scope === 'このアプリ'; }).length;
    var otherApps = {};
    list.forEach(function (r) { if (r.scope === 'ほかのアプリ') otherApps[r.appId] = true; });
    var there = list.length - here;
    var lines = [];
    if (!list.length) lines.push({ alert: false, text: name + 'が使われている場所は、見つかりませんでした。' });
    else {
      var parts = [];
      if (here) parts.push('このアプリの' + here + 'か所');
      if (there) parts.push('ほかのアプリ（' + Object.keys(otherApps).length + '個）の' + there + 'か所');
      var note = !here ? '（このアプリの中では使われていません）' : (info.others && !there ? '（調べたほかのアプリでは使われていません）' : '');
      lines.push({ alert: false, text: name + 'は、' + parts.join('と、') + 'で使われています' + note + '。' });
    }
    if (list.some(function (r) { return r.blocks; })) lines.push({ alert: true, text: '★の付いた場所で使われている間は、kintoneでこのフィールドを削除できません。先にその設定を直してください。' });
    if (list.length) lines.push({ alert: false, text: 'フィールドコードを変えると、計算式の中のコードは自動で変わります。JavaScriptのカスタマイズやプラグインの設定の中は変わらないため、合わせて直す必要があります。' });
    if (!info.others) lines.push({ alert: false, text: 'ほかのアプリは調べていません。' });
    else {
      lines.push({ alert: false, text: 'ほかのアプリは、操作した人が見られるアプリだけを調べています。kintoneシステム管理者でない人が調べたときは、見られないアプリで使われていても、ここには出ません（その場合も、kintoneが削除を止めることがあります）。' });
      if (info.otherErrors) lines.push({ alert: false, text: '設定を読み取れなかったアプリが' + info.otherErrors + '個あります（閲覧や管理の権限がないため）。' });
    }
    if (info.errors.length) lines.push({ alert: false, text: '読み取れなかった設定：' + info.errors.join('、') });
    if (info.urls.length) lines.push({ alert: false, text: 'URLで読み込んでいるJavaScript（' + info.urls.length + '件）の中は調べていません。' });
    return lines;
  }

  function csvEscape(value) {
    var s = value === undefined || value === null ? '' : String(value);
    if (/[",\r\n]/.test(s)) s = '"' + s.replace(/"/g, '""') + '"';
    return s;
  }

  var CSV_HEADER = ['範囲', 'アプリID', 'アプリ名', '場所', '内容', 'kintoneが削除を止める'];

  function buildCsv(list) {
    var lines = [CSV_HEADER].concat(list.map(function (r) { return [r.scope, r.appId, r.appName, r.place, r.detail, r.blocks ? '○' : '']; }));
    return '﻿' + lines.map(function (row) { return row.map(csvEscape).join(','); }).join('\r\n') + '\r\n';
  }

  function safeFileName(name) {
    var s = String(name || '').replace(/[\\/:*?"<>|]/g, '_').trim();
    return s || 'field';
  }

  var Core = { SETTINGS: SETTINGS, scriptFiles: scriptFiles, summaryLines: summaryLines, buildCsv: buildCsv, safeFileName: safeFileName, errorText: errorText };

  if (typeof module !== 'undefined' && module.exports) {
    module.exports = Core;
    return;
  }
  if (typeof kintone === 'undefined' || !F) return;

  // ===== 画面（kintone上でだけ動く部分） =====

  var P = 'dcy-fi';
  var MAX_SCRIPT = 2 * 1024 * 1024; // 2MBより大きいファイルは読まない
  var current = { name: '', list: null };

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

  function readScript(file) {
    if (file.size > MAX_SCRIPT) return Promise.resolve(null);
    return fetch(kintone.api.url('/k/v1/file', true) + '?fileKey=' + encodeURIComponent(file.fileKey), { headers: { 'X-Requested-With': 'XMLHttpRequest' }, credentials: 'same-origin' })
      .then(function (r) { return r.ok ? r.text() : null; })
      .then(function (text) { return text === null ? null : { screen: file.screen, name: file.name, text: text }; }, function () { return null; });
  }

  function loadSettings(appId) {
    return Promise.all(SETTINGS.map(function (s) { return settle(api(s[1], { app: appId })); })).then(function (list) {
      var settings = {};
      var errors = [];
      SETTINGS.forEach(function (s, i) {
        if (list[i].ok) settings[s[0]] = list[i].value;
        else { settings[s[0]] = null; if (s[0] !== 'fields') errors.push(s[2] + '（' + errorText(list[i].error) + '）'); }
      });
      return { settings: settings, errors: errors, fieldsError: list[0].ok ? null : list[0].error };
    });
  }

  // ほかのアプリのフィールドとアクションを、3つずつ読む
  function loadOthers(apps, skipId, onProgress) {
    var targets = apps.filter(function (a) { return String(a.appId) !== String(skipId); });
    var out = [];
    var errors = 0;
    var index = 0;
    var done = 0;
    function worker() {
      if (index >= targets.length) return Promise.resolve();
      var app = targets[index++];
      return Promise.all([settle(api('/k/v1/app/form/fields.json', { app: app.appId })), settle(api('/k/v1/app/actions.json', { app: app.appId }))]).then(function (r) {
        if (!r[0].ok) errors += 1;
        out.push({ appId: app.appId, name: app.name, fields: r[0].ok ? r[0].value : null, actions: r[1].ok ? r[1].value : null });
        done += 1;
        onProgress(done, targets.length);
        return worker();
      });
    }
    var workers = [];
    for (var i = 0; i < Math.min(3, targets.length); i++) workers.push(worker());
    return Promise.all(workers).then(function () { return { list: out, errors: errors }; });
  }

  function renderTable(container, list) {
    container.innerHTML = '';
    var headers = ['範囲', 'アプリ', '場所', '内容'];
    var thead = el('thead', {}, [el('tr', {}, headers.map(function (h) { return el('th', { text: h }); }))]);
    var body = list.length ? list.map(function (r) {
      var place = el('td', {}, [r.place]);
      if (r.blocks) place.appendChild(el('span', { class: P + '-star', text: ' ★', title: 'kintoneがフィールドの削除を止める設定です' }));
      return el('tr', {}, [
        el('td', { text: r.scope }),
        el('td', { text: r.appName + '（ID: ' + r.appId + '）' }),
        place,
        el('td', { class: P + '-detail', text: r.detail })
      ]);
    }) : [el('tr', {}, [el('td', { colspan: '4', class: P + '-empty', text: '使われている場所は見つかりませんでした' })])];
    container.appendChild(el('table', { class: P + '-table' }, [thead, el('tbody', {}, body)]));
  }

  function downloadCsv() {
    if (!current.list) return;
    var blob = new Blob([buildCsv(current.list)], { type: 'text/csv;charset=utf-8' });
    var url = URL.createObjectURL(blob);
    var d = new Date();
    var stamp = d.getFullYear() + ('0' + (d.getMonth() + 1)).slice(-2) + ('0' + d.getDate()).slice(-2);
    var a = el('a', { href: url, download: 'フィールドの影響_' + current.name + '_' + stamp + '.csv' });
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    setTimeout(function () { URL.revokeObjectURL(url); }, 1000);
  }

  function moreLink() {
    return el('p', { class: P + '-more' }, [
      'すべてのアプリをまとめて点検するなら：',
      el('a', { href: 'https://daichoya.github.io/pro.html', target: '_blank', rel: 'noopener', text: '棚卸し・変更監視パック（有料版・30日間の無料試用あり）' })
    ]);
  }

  function openModal() {
    if (document.getElementById(P + '-overlay')) return;
    var apps = [];
    var appSelect = el('select', { class: P + '-select', 'aria-label': '調べるアプリ' });
    var fieldSelect = el('select', { class: P + '-select ' + P + '-select-field', 'aria-label': '調べるフィールド', disabled: 'disabled' });
    var runBtn = el('button', { type: 'button', class: P + '-btn ' + P + '-primary', text: '影響を調べる', disabled: 'disabled' });
    var csvBtn = el('button', { type: 'button', class: P + '-btn', text: 'CSVで書き出す', disabled: 'disabled' });
    var closeBtn = el('button', { type: 'button', class: P + '-close', 'aria-label': '閉じる', text: '×' });
    var othersBox = el('input', { type: 'checkbox', id: P + '-others', checked: 'checked' });
    var scriptBox = el('input', { type: 'checkbox', id: P + '-scripts', checked: 'checked' });
    var status = el('p', { class: P + '-status', role: 'status' });
    var summary = el('div', { class: P + '-summary', hidden: 'hidden' });
    var result = el('div', { class: P + '-result' });
    var dialog = el('div', { class: P + '-dialog', role: 'dialog', 'aria-modal': 'true', 'aria-label': 'フィールド影響チェッカー' }, [
      el('div', { class: P + '-head' }, [el('h2', { text: 'フィールド影響チェッカー' }), closeBtn]),
      el('p', { class: P + '-lead', text: 'フィールドを削除したり、フィールドコードを変えたりする前に、そのフィールドが使われている場所を一覧にします。' }),
      el('div', { class: P + '-controls' }, [appSelect, fieldSelect, runBtn, csvBtn]),
      el('div', { class: P + '-options' }, [
        el('label', { for: P + '-others' }, [othersBox, 'ほかのアプリのルックアップ・関連レコード一覧・アクションも調べる（すべてのアプリのフィールドとアクションの設定を読み取ります）']),
        el('label', { for: P + '-scripts' }, [scriptBox, 'JavaScriptのカスタマイズの中も調べる（アプリに読み込まれたファイルを読み取ります）'])
      ]),
      status, summary, result,
      el('p', { class: P + '-note', text: 'プラグインの設定の中は調べません。JavaScriptは、フィールドコードが文字列として書かれている場所を見つけます（実際に使っているかは、ファイルの中身で確かめてください）。' }),
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

    function loadFields() {
      fieldSelect.innerHTML = '';
      fieldSelect.setAttribute('disabled', 'disabled');
      runBtn.setAttribute('disabled', 'disabled');
      status.textContent = 'フィールドを読み込み中…';
      api('/k/v1/app/form/fields.json', { app: appSelect.value }).then(function (resp) {
        F.choices(resp.properties).sort(function (a, b) { return a.label < b.label ? -1 : a.label > b.label ? 1 : 0; }).forEach(function (c) {
          fieldSelect.appendChild(el('option', { value: c.code, text: c.label + '（' + c.code + '）' }));
        });
        fieldSelect.removeAttribute('disabled');
        runBtn.removeAttribute('disabled');
        status.textContent = 'フィールドを選んで「影響を調べる」を押してください。';
      }).catch(function (error) {
        status.textContent = 'フィールドを読み込めませんでした（' + errorText(error) + '）。';
      });
    }
    appSelect.addEventListener('change', loadFields);

    runBtn.addEventListener('click', function () {
      var appId = appSelect.value;
      var code = fieldSelect.value;
      if (!appId || !code) return;
      var option = fieldSelect.options[fieldSelect.selectedIndex];
      var target = { appId: appId, appName: (apps.filter(function (a) { return String(a.appId) === String(appId); })[0] || {}).name || '', code: code, label: option ? option.textContent.replace(/（[^（]*）$/, '') : code };
      runBtn.setAttribute('disabled', 'disabled');
      csvBtn.setAttribute('disabled', 'disabled');
      summary.hidden = true;
      result.innerHTML = '';
      status.textContent = 'このアプリの設定を読み込み中…';
      var info = { errors: [], urls: [], others: othersBox.checked, otherErrors: 0 };
      var loaded;
      loadSettings(appId).then(function (r) {
        loaded = r;
        info.errors = r.errors;
        var files = Core.scriptFiles(r.settings.customize);
        info.urls = files.urls;
        if (!scriptBox.checked) return [];
        status.textContent = 'JavaScriptのファイルを読み込み中…';
        return Promise.all(files.files.map(readScript));
      }).then(function (scripts) {
        var readable = scripts.filter(function (x) { return x; });
        if (!info.others) return { scripts: readable, others: [] };
        return loadOthers(apps, appId, function (done, total) { status.textContent = 'ほかのアプリの設定を読み込み中…（' + done + ' / ' + total + '）'; })
          .then(function (o) { info.otherErrors = o.errors; return { scripts: readable, others: o.list }; });
      }).then(function (x) {
        var list = F.find(target, loaded.settings, x.scripts, x.others);
        current.name = safeFileName(target.label);
        current.list = list;
        status.textContent = '';
        summary.innerHTML = '';
        summaryLines(target, list, info).forEach(function (line) {
          summary.appendChild(el('p', line.alert ? { class: P + '-alert', text: line.text } : { text: line.text }));
        });
        summary.hidden = false;
        renderTable(result, list);
        if (list.length) csvBtn.removeAttribute('disabled');
      }).catch(function (error) {
        status.textContent = '調べられませんでした（' + errorText(error) + '）。';
      }).then(function () {
        runBtn.removeAttribute('disabled');
      });
    });

    document.body.appendChild(overlay);
    status.textContent = 'アプリの一覧を読み込み中…';
    getAllApps().then(function (list) {
      apps = list;
      var thisApp = String(kintone.app.getId());
      list.forEach(function (app) {
        var option = el('option', { value: app.appId, text: app.name + '（ID: ' + app.appId + '）' });
        if (String(app.appId) === thisApp) option.selected = true;
        appSelect.appendChild(option);
      });
      loadFields();
    }).catch(function (error) {
      status.textContent = 'アプリの一覧を取得できませんでした（' + errorText(error) + '）。';
    });
  }

  kintone.events.on('app.record.index.show', function (event) {
    var space = kintone.app.getHeaderMenuSpaceElement();
    if (!space || document.getElementById(P + '-open')) return event;
    var button = el('button', { id: P + '-open', type: 'button', class: P + '-btn', text: 'フィールドの影響' });
    button.addEventListener('click', openModal);
    space.appendChild(button);
    return event;
  });
})();
