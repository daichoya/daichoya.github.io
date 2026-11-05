/*
 * 台帳屋 権限チェッカー for kintone
 * 一覧画面に「権限チェッカー」ボタンを追加し、選んだアプリの
 * アプリ・レコード・フィールドのアクセス権を1つの画面に表示して、CSVで書き出す。
 * データはすべてお客様のkintoneの中で処理し、外部には送らない。
 */
(function () {
  'use strict';

  // ===== 表示用の変換（kintoneに依存しない純粋な関数。テスト対象） =====

  var ENTITY_LABELS = {
    USER: 'ユーザー',
    GROUP: 'グループ',
    ORGANIZATION: '組織',
    CREATOR: '作成者',
    FIELD_ENTITY: 'フィールド'
  };

  var ACCESS_LABELS = { WRITE: '閲覧・編集', READ: '閲覧のみ', NONE: '権限なし' };

  function mark(value) {
    return value ? '○' : '－';
  }

  function entityLabel(entity) {
    if (!entity) return { type: '', code: '' };
    return {
      type: ENTITY_LABELS[entity.type] || entity.type || '',
      code: entity.type === 'CREATOR' ? '（レコード作成者）' : (entity.code || '')
    };
  }

  function appRows(appAcl) {
    return ((appAcl && appAcl.rights) || []).map(function (right, index) {
      var e = entityLabel(right.entity);
      return {
        order: index + 1,
        type: e.type,
        code: e.code,
        includeSubs: mark(right.includeSubs),
        appEditable: mark(right.appEditable),
        recordViewable: mark(right.recordViewable),
        recordAddable: mark(right.recordAddable),
        recordEditable: mark(right.recordEditable),
        recordDeletable: mark(right.recordDeletable),
        recordImportable: mark(right.recordImportable),
        recordExportable: mark(right.recordExportable)
      };
    });
  }

  function recordRows(recordAcl) {
    var rows = [];
    ((recordAcl && recordAcl.rights) || []).forEach(function (rule, index) {
      var condition = rule.filterCond ? rule.filterCond : 'すべてのレコード';
      (rule.entities || []).forEach(function (item) {
        var e = entityLabel(item.entity);
        rows.push({
          order: index + 1,
          condition: condition,
          type: e.type,
          code: e.code,
          viewable: mark(item.viewable),
          editable: mark(item.editable),
          deletable: mark(item.deletable),
          includeSubs: mark(item.includeSubs)
        });
      });
    });
    return rows;
  }

  // テーブル内のフィールドも含めて、フィールドコード→表示名の対応表を作る
  function fieldLabelMap(properties) {
    var map = {};
    Object.keys(properties || {}).forEach(function (code) {
      var prop = properties[code];
      map[code] = prop.label || code;
      if (prop.type === 'SUBTABLE' && prop.fields) {
        Object.keys(prop.fields).forEach(function (inner) {
          map[inner] = (prop.label || code) + ' > ' + (prop.fields[inner].label || inner);
        });
      }
    });
    return map;
  }

  function fieldRows(fieldAcl, labels) {
    var rows = [];
    ((fieldAcl && fieldAcl.rights) || []).forEach(function (rule) {
      var label = (labels && labels[rule.code]) || '（不明なフィールド）';
      (rule.entities || []).forEach(function (item) {
        var e = entityLabel(item.entity);
        rows.push({
          fieldLabel: label,
          fieldCode: rule.code,
          type: e.type,
          code: e.code,
          accessibility: ACCESS_LABELS[item.accessibility] || item.accessibility || '',
          includeSubs: mark(item.includeSubs)
        });
      });
    });
    return rows;
  }

  function csvEscape(value) {
    var s = value === undefined || value === null ? '' : String(value);
    if (/[",\r\n]/.test(s)) s = '"' + s.replace(/"/g, '""') + '"';
    return s;
  }

  var CSV_HEADER = [
    '区分', '優先順位', '条件・フィールド名', 'フィールドコード', '対象の種類', '対象',
    '下位組織に継承', 'アプリ管理', '閲覧', '追加', '編集', '削除',
    'ファイル読み込み', 'ファイル書き出し', 'フィールドの権限'
  ];

  // Excelで文字化けしないよう、BOM付きUTF-8・CRLF改行で作る
  // 権限の表の行（CSVとExcelで同じ中身）
  function permissionLines(data) {
    var lines = [];
    data.app.forEach(function (r) {
      lines.push(['アプリ', r.order, '', '', r.type, r.code, r.includeSubs, r.appEditable,
        r.recordViewable, r.recordAddable, r.recordEditable, r.recordDeletable,
        r.recordImportable, r.recordExportable, '']);
    });
    data.record.forEach(function (r) {
      lines.push(['レコード', r.order, r.condition, '', r.type, r.code, r.includeSubs, '',
        r.viewable, '', r.editable, r.deletable, '', '', '']);
    });
    data.field.forEach(function (r) {
      lines.push(['フィールド', '', r.fieldLabel, r.fieldCode, r.type, r.code, r.includeSubs,
        '', '', '', '', '', '', '', r.accessibility]);
    });
    return lines;
  }

  function buildCsv(data) {
    var lines = [CSV_HEADER].concat(permissionLines(data));
    return '﻿' + lines.map(function (row) {
      return row.map(csvEscape).join(',');
    }).join('\r\n') + '\r\n';
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
  function excelLead(rows) {
    var count = function (area) { return rows.filter(function (r) { return r[0] === area; }).length; };
    var text = 'アクセス権の設定は、アプリ ' + count('アプリ') + '件・レコード ' + count('レコード') + '件・フィールド ' + count('フィールド') + '件です。';
    var everyone = rows.some(function (r) { return r[0] === 'アプリ' && /^everyone$/i.test(String(r[5])) && r[7] === '○'; });
    return text + (everyone ? 'Everyoneに「アプリ管理」の権限があります（だれでもアプリの設定を変えられます）。' : '');
  }

  function buildExcel(data, info) {
    var rows = permissionLines(data);
    return Xlsx.listBook({
      title: '権限一覧', target: info.app, sub: excelSub(info, '権限チェッカー'), sheet: '権限一覧', lead: excelLead(rows),
      header: CSV_HEADER, rows: rows, widths: [10, 8, 26, 16, 10, 22, 9, 8, 6, 6, 6, 6, 9, 9, 12],
      numbers: [1], centers: [6, 7, 8, 9, 10, 11, 12, 13], creator: '台帳屋 権限チェッカー',
      note: '○：権限あり　－：権限なし。アプリとレコードのアクセス権は、優先順位の数の小さい行が優先されます。'
    });
  }

  // Windowsのファイル名に使えない文字を置き換える
  function safeFileName(name) {
    var s = String(name || '').replace(/[\\/:*?"<>|]/g, '_').trim();
    return s || 'app';
  }

  function errorMessage(error) {
    var code = error && error.code;
    if (code === 'CB_NO02' || code === 'GAIA_NO01') {
      return 'このアプリの権限を見るには、そのアプリの管理権限が必要です。';
    }
    var detail = (error && error.message) || '原因不明のエラー';
    return '権限の取得に失敗しました（' + detail + '）。';
  }

  var Core = {
    entityLabel: entityLabel,
    appRows: appRows,
    recordRows: recordRows,
    fieldLabelMap: fieldLabelMap,
    fieldRows: fieldRows,
    csvEscape: csvEscape,
    buildCsv: buildCsv,
    buildExcel: buildExcel,
    safeFileName: safeFileName,
    errorMessage: errorMessage
  };

  if (typeof module !== 'undefined' && module.exports) {
    module.exports = Core;
    return;
  }
  if (typeof kintone === 'undefined') return;

  // ===== 画面（kintone上でだけ動く部分） =====

  var P = 'dcy-pc';
  var current = { appId: null, appName: '', data: null };

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

  function api(path, params) {
    return kintone.api(kintone.api.url(path, true), 'GET', params);
  }

  // 閲覧できるアプリをすべて取得する（1回100件まで）
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

  function fetchPermissions(appId) {
    return Promise.all([
      api('/k/v1/app/form/fields.json', { app: appId }),
      api('/k/v1/app/acl.json', { app: appId }),
      api('/k/v1/record/acl.json', { app: appId }),
      api('/k/v1/field/acl.json', { app: appId })
    ]).then(function (res) {
      var labels = fieldLabelMap(res[0].properties);
      return {
        app: appRows(res[1]),
        record: recordRows(res[2]),
        field: fieldRows(res[3], labels)
      };
    });
  }

  function table(headers, rows, keys) {
    var thead = el('thead', {}, [el('tr', {}, headers.map(function (h) {
      return el('th', { text: h });
    }))]);
    var body = rows.length ? rows.map(function (row) {
      return el('tr', {}, keys.map(function (k) {
        var v = row[k];
        var cls = v === '○' ? P + '-yes' : (v === '－' ? P + '-no' : '');
        return el('td', cls ? { class: cls, text: String(v) } : { text: String(v) });
      }));
    }) : [el('tr', {}, [el('td', { colspan: String(headers.length), class: P + '-empty', text: '設定はありません' })])];
    return el('table', { class: P + '-table' }, [thead, el('tbody', {}, body)]);
  }

  function renderResult(container, data) {
    container.innerHTML = '';
    container.appendChild(el('h3', { text: 'アプリのアクセス権（上にあるものほど優先）' }));
    container.appendChild(table(
      ['順', '種類', '対象', '下位組織', 'アプリ管理', '閲覧', '追加', '編集', '削除', '読み込み', '書き出し'],
      data.app,
      ['order', 'type', 'code', 'includeSubs', 'appEditable', 'recordViewable', 'recordAddable',
        'recordEditable', 'recordDeletable', 'recordImportable', 'recordExportable']
    ));
    container.appendChild(el('h3', { text: 'レコードのアクセス権（上の条件ほど優先）' }));
    container.appendChild(table(
      ['順', '条件', '種類', '対象', '閲覧', '編集', '削除', '下位組織'],
      data.record,
      ['order', 'condition', 'type', 'code', 'viewable', 'editable', 'deletable', 'includeSubs']
    ));
    container.appendChild(el('h3', { text: 'フィールドのアクセス権' }));
    container.appendChild(table(
      ['フィールド名', 'フィールドコード', '種類', '対象', '権限', '下位組織'],
      data.field,
      ['fieldLabel', 'fieldCode', 'type', 'code', 'accessibility', 'includeSubs']
    ));
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
    if (!current.data) return;
    var d = new Date(); var stamp = d.getFullYear() + ('0' + (d.getMonth() + 1)).slice(-2) + ('0' + d.getDate()).slice(-2);
    saveExcel(buildExcel(current.data, { app: current.displayName, host: location.hostname, at: d }), '権限一覧_' + current.appName + '_' + stamp + '.xlsx');
  }

  function downloadCsv() {
    if (!current.data) return;
    var blob = new Blob([buildCsv(current.data)], { type: 'text/csv;charset=utf-8' });
    var url = URL.createObjectURL(blob);
    var d = new Date();
    var stamp = d.getFullYear() + ('0' + (d.getMonth() + 1)).slice(-2) + ('0' + d.getDate()).slice(-2);
    var a = el('a', { href: url, download: '権限一覧_' + current.appName + '_' + stamp + '.csv' });
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
    var select = el('select', { class: P + '-select', 'aria-label': '確認するアプリ' });
    var runBtn = el('button', { type: 'button', class: P + '-btn ' + P + '-primary', text: '権限を表示' });
    var excelBtn = el('button', { type: 'button', class: P + '-btn', text: 'Excelで保存', disabled: 'disabled' });
    var csvBtn = el('button', { type: 'button', class: P + '-btn', text: 'CSVで書き出す', disabled: 'disabled' });
    var closeBtn = el('button', { type: 'button', class: P + '-close', 'aria-label': '閉じる', text: '×' });
    var status = el('p', { class: P + '-status', role: 'status' });
    var result = el('div', { class: P + '-result' });
    var dialog = el('div', { class: P + '-dialog', role: 'dialog', 'aria-modal': 'true', 'aria-label': '権限チェッカー' }, [
      el('div', { class: P + '-head' }, [el('h2', { text: '権限チェッカー' }), closeBtn]),
      el('div', { class: P + '-controls' }, [select, runBtn, excelBtn, csvBtn]),
      status,
      result,
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

    runBtn.addEventListener('click', function () {
      var appId = select.value;
      if (!appId) return;
      runBtn.disabled = true;
      csvBtn.disabled = true;
      status.textContent = '読み込み中…';
      result.innerHTML = '';
      fetchPermissions(appId).then(function (data) {
        current.appId = appId;
        current.appName = safeFileName(appNames[appId] || ('app' + appId));
        current.displayName = appNames[appId] || ('アプリID ' + appId);
        current.data = data;
        status.textContent = '';
        renderResult(result, data);
        csvBtn.disabled = false;
      }).catch(function (error) {
        current.data = null;
        status.textContent = errorMessage(error);
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
      status.textContent = 'アプリを選んで「権限を表示」を押してください。';
    }).catch(function (error) {
      status.textContent = 'アプリの一覧を取得できませんでした（' + ((error && error.message) || '原因不明') + '）。';
    });
  }

  kintone.events.on('app.record.index.show', function (event) {
    var space = kintone.app.getHeaderMenuSpaceElement();
    if (!space || document.getElementById(P + '-open')) return event;
    var button = el('button', { id: P + '-open', type: 'button', class: P + '-btn', text: '権限チェッカー' });
    button.addEventListener('click', openModal);
    space.appendChild(button);
    return event;
  });
})();
