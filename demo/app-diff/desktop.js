/*
 * 台帳屋 アプリ設定の比較 for kintone
 * 一覧画面に「設定の比較」ボタンを追加し、このアプリと、選んだもう1つのアプリの設定の違いを一覧にする。
 * 本番と検証のアプリ、コピー元とコピー先のアプリなどを比べる用途。
 * 読み取るのはアプリの設定だけで、レコードは読まない。データはすべてお客様のkintoneの中で処理し、外部には送らない。
 */
(function () {
  'use strict';

  // ===== 画面に依存しない部分（テスト対象） =====

  var SETTINGS = [
    ['fields', '/k/v1/app/form/fields.json', 'フィールド'],
    ['views', '/k/v1/app/views.json', '一覧'],
    ['reports', '/k/v1/app/reports.json', 'グラフ'],
    ['appAcl', '/k/v1/app/acl.json', 'アプリのアクセス権'],
    ['recordAcl', '/k/v1/record/acl.json', 'レコードのアクセス権'],
    ['fieldAcl', '/k/v1/field/acl.json', 'フィールドのアクセス権'],
    ['status', '/k/v1/app/status.json', 'プロセス管理'],
    ['general', '/k/v1/app/notifications/general.json', 'アプリの条件通知'],
    ['perRecord', '/k/v1/app/notifications/perRecord.json', 'レコードの条件通知'],
    ['reminder', '/k/v1/app/notifications/reminder.json', 'リマインダーの条件通知'],
    ['actions', '/k/v1/app/actions.json', 'アプリアクション'],
    ['appSettings', '/k/v1/app/settings.json', 'アプリの設定'],
    ['plugins', '/k/v1/app/plugins.json', 'プラグイン'],
    ['customize', '/k/v1/app/customize.json', 'JavaScript・CSSでカスタマイズ']
  ];

  var FIELD_TYPES = {
    SINGLE_LINE_TEXT: '文字列（1行）', MULTI_LINE_TEXT: '文字列（複数行）', RICH_TEXT: 'リッチエディター',
    NUMBER: '数値', CALC: '計算', RADIO_BUTTON: 'ラジオボタン', CHECK_BOX: 'チェックボックス',
    MULTI_SELECT: '複数選択', DROP_DOWN: 'ドロップダウン', DATE: '日付', TIME: '時刻', DATETIME: '日時',
    LINK: 'リンク', FILE: '添付ファイル', USER_SELECT: 'ユーザー選択', ORGANIZATION_SELECT: '組織選択',
    GROUP_SELECT: 'グループ選択', REFERENCE_TABLE: '関連レコード一覧', SUBTABLE: 'テーブル', GROUP: 'グループ',
    RECORD_NUMBER: 'レコード番号', CREATOR: '作成者', CREATED_TIME: '作成日時', MODIFIER: '更新者',
    UPDATED_TIME: '更新日時', STATUS: 'ステータス', STATUS_ASSIGNEE: '作業者', CATEGORY: 'カテゴリー'
  };
  // アプリごとに必ずあり、違いを比べる意味のないフィールド
  var SKIP_FIELDS = { __ID__: 1, __REVISION__: 1, STATUS: 1, STATUS_ASSIGNEE: 1, CATEGORY: 1 };
  var VIEW_TYPES = { LIST: '表形式', CALENDAR: 'カレンダー形式', CUSTOM: 'カスタマイズ形式' };
  var CHART_TYPES = { BAR: '横棒', COLUMN: '縦棒', PIE: '円', LINE: '折れ線', PIVOT_TABLE: 'クロス集計表', TABLE: '表', AREA: '面', SPLINE: '曲線', SPLINE_AREA: '曲線面' };
  var AGG_TYPES = { COUNT: '件数', SUM: '合計', AVERAGE: '平均', MAX: '最大', MIN: '最小' };
  var APP_RIGHTS = [['appEditable', 'アプリ管理'], ['recordViewable', '閲覧'], ['recordAddable', '追加'], ['recordEditable', '編集'],
    ['recordDeletable', '削除'], ['recordImportable', '読み込み'], ['recordExportable', '書き出し']];
  var ACCESS = { WRITE: '閲覧・編集', READ: '閲覧のみ', NONE: '権限なし' };

  function errorText(error) {
    var code = error && error.code;
    if (code === 'CB_NO02' || code === 'GAIA_NO01') return '権限がありません';
    return (error && error.message) || '原因不明のエラー';
  }

  function byIndex(obj) {
    return Object.keys(obj || {}).map(function (k) { return obj[k]; })
      .sort(function (a, b) { return Number(a.index || 0) - Number(b.index || 0); });
  }

  function entityText(e, subs) {
    if (!e) return '';
    if (e.type === 'CREATOR') return '作成者';
    if (e.type === 'GROUP' && e.code === 'everyone') return 'Everyone';
    var kind = { USER: 'ユーザー', GROUP: 'グループ', ORGANIZATION: '組織', FIELD_ENTITY: 'フィールド', CUSTOM_FIELD: 'フィールド' }[e.type] || e.type;
    return kind + '：' + (e.code || '') + (subs ? '（下位組織を含む）' : '');
  }

  function appRef(id, names) {
    if (!id) return '';
    return names && names[String(id)] ? '「' + names[String(id)] + '」' : 'アプリ' + id;
  }

  function join(parts) { return parts.filter(function (p) { return p !== '' && p !== null && p !== undefined; }).join(' ／ '); }

  function fieldText(p, names, table) {
    var parts = [(FIELD_TYPES[p.type] || p.type) + '「' + (p.label || '') + '」'];
    if (table) parts.push('テーブル「' + table + '」の中');
    if (p.required) parts.push('必須');
    if (p.unique) parts.push('重複禁止');
    if (p.options) parts.push('選択肢：' + byIndex(p.options).map(function (o) { return o.label; }).join('、'));
    if (p.expression) parts.push('計算式：' + p.expression);
    if (p.hideExpression) parts.push('計算式を表示しない');
    if (p.defaultValue !== undefined && p.defaultValue !== '' && !(Array.isArray(p.defaultValue) && !p.defaultValue.length)) {
      var d = p.defaultValue;
      parts.push('初期値：' + (Array.isArray(d) ? d.map(function (x) { return typeof x === 'object' ? (x.code || '') : x; }).join('、') : d));
    }
    if (p.defaultNowValue) parts.push('初期値：現在の日時');
    ['minLength', 'maxLength', 'minValue', 'maxValue'].forEach(function (k) {
      if (p[k] !== undefined && p[k] !== '') parts.push({ minLength: '最小の文字数', maxLength: '最大の文字数', minValue: '最小値', maxValue: '最大値' }[k] + '：' + p[k]);
    });
    if (p.displayScale !== undefined && p.displayScale !== '') parts.push('小数点以下の桁数：' + p.displayScale);
    if (p.unit) parts.push('単位：' + p.unit);
    if (p.format && p.type === 'CALC') parts.push('表示形式：' + p.format);
    if (p.protocol && p.type === 'LINK') parts.push('種類：' + p.protocol);
    if (p.lookup) {
      var lk = p.lookup;
      parts.push('ルックアップ：' + appRef(lk.relatedApp && lk.relatedApp.app, names) + 'の' + lk.relatedKeyField + 'をキーに');
      if ((lk.fieldMappings || []).length) parts.push('コピー：' + lk.fieldMappings.map(function (m) { return m.relatedField + '→' + m.field; }).join('、'));
      if ((lk.lookupPickerFields || []).length) parts.push('選択画面：' + lk.lookupPickerFields.join('、'));
      if (lk.filterCond) parts.push('絞り込み：' + lk.filterCond);
      if (lk.sort) parts.push('並び順：' + lk.sort);
    }
    if (p.referenceTable) {
      var rt = p.referenceTable;
      parts.push('関連レコード：' + appRef(rt.relatedApp && rt.relatedApp.app, names));
      if (rt.condition) parts.push('条件：' + rt.condition.field + '＝' + rt.condition.relatedField);
      if (rt.filterCond) parts.push('絞り込み：' + rt.filterCond);
      if ((rt.displayFields || []).length) parts.push('表示：' + rt.displayFields.join('、'));
      if (rt.sort) parts.push('並び順：' + rt.sort);
      if (rt.size) parts.push('表示する件数：' + rt.size);
    }
    return join(parts);
  }

  function stripHtml(html) {
    return String(html || '').replace(/<br\s*\/?>/gi, ' ').replace(/<[^>]+>/g, '').replace(/&nbsp;/g, ' ').replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/\s+/g, ' ').trim();
  }

  function targetsText(list) {
    return (list || []).map(function (t) { return entityText(t.entity, t.includeSubs); }).join('、');
  }

  // アプリの設定を、「種類｜項目 → 内容」の並びにそろえる。names：アプリIDとアプリ名の対応（ルックアップの参照先などを名前で比べる）
  function normalize(s, names) {
    s = s || {};
    var out = [];
    function add(group, item, text) { out.push({ group: group, item: item, text: text === undefined || text === null ? '' : String(text) }); }

    var props = s.fields && s.fields.properties;
    if (props) {
      Object.keys(props).forEach(function (code) {
        var p = props[code];
        if (SKIP_FIELDS[p.type]) return;
        add('フィールド', (p.label || code) + '（' + code + '）', fieldText(p, names, null));
        if (p.type === 'SUBTABLE') Object.keys(p.fields || {}).forEach(function (inner) {
          var q = p.fields[inner];
          add('フィールド', (q.label || inner) + '（' + inner + '）', fieldText(q, names, p.label || code));
        });
      });
    }

    byIndex(s.views && s.views.views).forEach(function (v) {
      add('一覧', v.name, join([VIEW_TYPES[v.type] || v.type,
        v.type === 'LIST' ? '表示項目：' + (v.fields || []).join('、') : '',
        v.type === 'CALENDAR' ? '日付：' + (v.date || '') + '、タイトル：' + (v.title || '') : '',
        v.type === 'CUSTOM' ? 'HTMLあり' + (v.pager === false ? '、ページネーションなし' : '') : '',
        v.filterCond ? '絞り込み：' + v.filterCond : '', v.sort ? '並び順：' + v.sort : '']));
    });

    byIndex(s.reports && s.reports.reports).forEach(function (r) {
      add('グラフ', r.name, join([(CHART_TYPES[r.chartType] || r.chartType) + (r.chartMode ? '（' + r.chartMode + '）' : ''),
        (r.groups || []).length ? '分類：' + r.groups.map(function (g) { return g.code + (g.per ? '（' + g.per + '）' : ''); }).join('、') : '',
        (r.aggregations || []).length ? '集計：' + r.aggregations.map(function (a) { return (AGG_TYPES[a.type] || a.type) + (a.code ? '（' + a.code + '）' : ''); }).join('、') : '',
        r.filterCond ? '絞り込み：' + r.filterCond : '',
        (r.sorts || []).length ? '並び順：' + r.sorts.map(function (x) { return x.by + ' ' + x.order; }).join('、') : '']));
    });

    // アプリのアクセス権：対象ごとの権限と、上から順の優先順位を分けて比べる（1つ足しても、ほかの行が「違う」にならないように）
    var appRights = (s.appAcl && s.appAcl.rights) || [];
    if (s.appAcl) add('アプリのアクセス権', '優先順位（上から順）', appRights.map(function (r) { return entityText(r.entity, r.includeSubs); }).join(' → '));
    appRights.forEach(function (r) {
      var granted = APP_RIGHTS.filter(function (x) { return r[x[0]]; }).map(function (x) { return x[1]; });
      add('アプリのアクセス権', entityText(r.entity, r.includeSubs), granted.length ? granted.join('・') : '権限なし');
    });
    ((s.recordAcl && s.recordAcl.rights) || []).forEach(function (r, i) {
      add('レコードのアクセス権', (i + 1) + '番目の条件', join(['条件：' + (r.filterCond || 'すべてのレコード'),
        (r.entities || []).map(function (e) {
          var g = [['viewable', '閲覧'], ['editable', '編集'], ['deletable', '削除']].filter(function (x) { return e[x[0]]; }).map(function (x) { return x[1]; });
          return entityText(e.entity, e.includeSubs) + '＝' + (g.length ? g.join('・') : '権限なし');
        }).join('、')]));
    });
    ((s.fieldAcl && s.fieldAcl.rights) || []).forEach(function (r) {
      add('フィールドのアクセス権', r.code, (r.entities || []).map(function (e) { return entityText(e.entity, e.includeSubs) + '＝' + (ACCESS[e.accessibility] || e.accessibility); }).join('、'));
    });

    if (s.status) {
      add('プロセス管理', '使う／使わない', s.status.enable ? '使う' : '使わない');
      if (s.status.enable) {
        byIndex(s.status.states).forEach(function (st) {
          var as = st.assignee || {};
          add('プロセス管理', 'ステータス「' + st.name + '」', join([(Number(st.index) + 1) + '番目',
            (as.entities || []).length ? '作業者：' + targetsText(as.entities) + '（' + ({ ONE: '次のユーザーから作業者を選択', ALL: '全員', ANY: '1人' }[as.type] || as.type) + '）' : '作業者なし']));
        });
        (s.status.actions || []).forEach(function (a) {
          add('プロセス管理', 'アクション「' + a.name + '」', join([a.from + '→' + a.to, a.filterCond ? '条件：' + a.filterCond : '',
            a.executableUser && (a.executableUser.entities || []).length ? '実行できる人：' + targetsText(a.executableUser.entities) : '']));
        });
      }
    }

    ((s.general && s.general.notifications) || []).forEach(function (n) {
      var on = [['recordAdded', 'レコード追加'], ['recordEdited', 'レコード編集'], ['commentAdded', 'コメント書き込み'], ['fileImported', 'ファイル読み込み']]
        .filter(function (x) { return n[x[0]]; }).map(function (x) { return x[1]; });
      add('アプリの条件通知', entityText(n.entity, n.includeSubs), on.join('・') || '通知なし');
    });
    ((s.perRecord && s.perRecord.notifications) || []).forEach(function (n, i) {
      add('レコードの条件通知', n.title || (i + 1) + '番目', join(['条件：' + (n.filterCond || 'すべてのレコード'), '通知先：' + targetsText(n.targets)]));
    });
    ((s.reminder && s.reminder.notifications) || []).forEach(function (n, i) {
      var t = n.timing || {};
      var when = t.code ? t.code + ' の' + (t.daysLater ? t.daysLater + '日後' : '当日') + (t.time ? ' ' + t.time : '') + (t.hoursLater ? ' ' + t.hoursLater + '時間後' : '') : '';
      add('リマインダーの条件通知', n.title || (i + 1) + '番目', join(['日時：' + when, n.filterCond ? '条件：' + n.filterCond : '', '通知先：' + targetsText(n.targets)]));
    });

    byIndex(s.actions && s.actions.actions).forEach(function (a) {
      add('アプリアクション', a.name, join(['作成先：' + (a.destApp && a.destApp.app ? appRef(a.destApp.app, names) : '（見られないアプリ）'),
        (a.mappings || []).length ? 'コピー：' + a.mappings.map(function (m) { return (m.srcType === 'FIELD' ? m.srcField : m.srcType === 'RECORD_URL' ? 'レコードのURL' : m.srcType) + '→' + m.destField; }).join('、') : '',
        a.filterCond ? '条件：' + a.filterCond : '',
        (a.entities || []).length ? '使える人：' + a.entities.map(function (e) { return entityText(e); }).join('、') : '']));
    });

    var g = s.appSettings;
    if (g) {
      add('アプリの設定', 'アプリ名', g.name || '');
      add('アプリの設定', '説明', stripHtml(g.description));
      if (g.theme) add('アプリの設定', 'デザインテーマ', g.theme);
      if (g.titleField) add('アプリの設定', 'レコードのタイトル', g.titleField.selectionMode === 'MANUAL' ? g.titleField.code : '自動');
      [['enableThumbnails', 'サムネイルを表示'], ['enableBulkDeletion', 'レコードの一括削除'], ['enableComments', 'コメント機能'],
        ['enableDuplicateRecord', 'レコードの再利用'], ['enableInlineRecordEditing', '一覧での編集']].forEach(function (x) {
        if (g[x[0]] !== undefined) add('アプリの設定', x[1], g[x[0]] ? '有効' : '無効');
      });
      if (g.firstMonthOfFiscalYear) add('アプリの設定', '年度の開始月', g.firstMonthOfFiscalYear + '月');
    }

    ((s.plugins && s.plugins.plugins) || []).forEach(function (p) {
      add('プラグイン', p.name, p.enabled === false ? '無効' : '有効');
    });

    var c = s.customize;
    if (c) {
      add('カスタマイズ', '適用するユーザー', { ALL: 'すべてのユーザー', ADMIN: 'アプリの管理者だけ', NONE: '適用しない' }[c.scope] || c.scope);
      [['desktop', 'PC'], ['mobile', 'モバイル']].forEach(function (screen) {
        [['js', 'JavaScript'], ['css', 'CSS']].forEach(function (kind) {
          var list = ((c[screen[0]] || {})[kind[0]]) || [];
          if (list.length) add('カスタマイズ', screen[1] + '・' + kind[1], list.map(function (x) { return x.type === 'URL' ? x.url : (x.file && x.file.name) || ''; }).join('、'));
        });
      });
    }
    return out;
  }

  // 2つのアプリの並びを比べる。同じ「種類｜項目」が2つ以上あるときは、出てきた順に対応づける
  function compare(a, b) {
    function keyed(list) {
      var map = {};
      var order = [];
      list.forEach(function (x) {
        var base = x.group + '\n' + x.item;
        var k = base;
        var n = 2;
        while (map[k]) { k = base + '\n' + n; n += 1; }
        map[k] = x;
        order.push(k);
      });
      return { map: map, order: order };
    }
    var A = keyed(a);
    var B = keyed(b);
    var rows = [];
    A.order.forEach(function (k) {
      var x = A.map[k];
      var y = B.map[k];
      if (!y) rows.push({ group: x.group, item: x.item, kind: 'このアプリだけ', a: x.text, b: '' });
      else if (x.text !== y.text) rows.push({ group: x.group, item: x.item, kind: '違う', a: x.text, b: y.text });
    });
    B.order.forEach(function (k) {
      if (!A.map[k]) rows.push({ group: B.map[k].group, item: B.map[k].item, kind: '比べるアプリだけ', a: '', b: B.map[k].text });
    });
    var groupOrder = {};
    ['アプリの設定', 'フィールド', '一覧', 'グラフ', 'アプリのアクセス権', 'レコードのアクセス権', 'フィールドのアクセス権', 'プロセス管理',
      'アプリの条件通知', 'レコードの条件通知', 'リマインダーの条件通知', 'アプリアクション', 'プラグイン', 'カスタマイズ'].forEach(function (g, i) { groupOrder[g] = i; });
    return rows.map(function (r, i) { r.i = i; return r; })
      .sort(function (x, y) { return (groupOrder[x.group] - groupOrder[y.group]) || (x.i - y.i); })
      .map(function (r) { delete r.i; return r; });
  }

  function summaryLines(rows, info) {
    var lines = [];
    if (!rows.length) lines.push({ alert: false, text: '読み取れた設定の中に、違いは見つかりませんでした。' });
    else {
      var counts = {};
      rows.forEach(function (r) { counts[r.group] = (counts[r.group] || 0) + 1; });
      lines.push({ alert: false, text: '違いは' + rows.length + '件です（' + Object.keys(counts).map(function (g) { return g + ' ' + counts[g] + '件'; }).join('、') + '）。' });
    }
    if (info.errors.length) lines.push({ alert: true, text: '読み取れなかった設定があります：' + info.errors.join('、') + '。その部分は比べていません。' });
    lines.push({ alert: false, text: 'フィールドはフィールドコードで、一覧・グラフ・アクションなどは名前で対応づけています。名前を変えたものは「だけ」の行に2つに分かれて出ます。' });
    return lines;
  }

  function csvEscape(value) {
    var s = value === undefined || value === null ? '' : String(value);
    if (/[",\r\n]/.test(s)) s = '"' + s.replace(/"/g, '""') + '"';
    return s;
  }

  function buildCsv(rows, nameA, nameB) {
    var lines = [['種類', '項目', '違い', 'このアプリ（' + nameA + '）', '比べるアプリ（' + nameB + '）']]
      .concat(rows.map(function (r) { return [r.group, r.item, r.kind, r.a, r.b]; }));
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

  // info: { host: kintoneの環境, at: 作成日時 }
  function buildExcel(rows, nameA, nameB, info) {
    return Xlsx.listBook({
      title: 'アプリ設定の比較', target: nameA + ' と ' + nameB, sub: excelSub(info, 'アプリ設定の比較'), sheet: '設定の違い',
      lead: summaryLines(rows, { errors: [] })[0].text,
      header: ['種類', '項目', '違い', 'このアプリ（' + nameA + '）', '比べるアプリ（' + nameB + '）'],
      rows: rows.map(function (r) { return [r.group, r.item, r.kind, r.a, r.b]; }),
      widths: [14, 26, 16, 50, 50], creator: '台帳屋 アプリ設定の比較',
      note: rows.length ? 'フィールドはフィールドコードで、一覧・グラフ・アクションなどは名前で対応づけています。' : '違いは見つかりませんでした。'
    });
  }

  function safeFileName(name) {
    var s = String(name || '').replace(/[\\/:*?"<>|]/g, '_').trim();
    return s || 'app';
  }

  var Core = { SETTINGS: SETTINGS, normalize: normalize, compare: compare, summaryLines: summaryLines, buildCsv: buildCsv, buildExcel: buildExcel, safeFileName: safeFileName, errorText: errorText, fieldText: fieldText };

  if (typeof module !== 'undefined' && module.exports) {
    module.exports = Core;
    return;
  }
  if (typeof kintone === 'undefined') return;

  // ===== 画面（kintone上でだけ動く部分） =====

  var P = 'dcy-ad';
  var current = { rows: null, nameA: '', nameB: '' };

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

  function loadSettings(appId, label) {
    return Promise.all(SETTINGS.map(function (s) { return settle(api(s[1], { app: appId })); })).then(function (list) {
      var settings = {};
      var errors = [];
      SETTINGS.forEach(function (s, i) {
        if (list[i].ok) settings[s[0]] = list[i].value;
        else { settings[s[0]] = null; errors.push(label + 'の' + s[2] + '（' + errorText(list[i].error) + '）'); }
      });
      return { settings: settings, errors: errors };
    });
  }

  // 片方だけ読めなかった種類は、違いとして出さない
  function dropUnreadable(rows, a, b) {
    var skip = {};
    SETTINGS.forEach(function (s) {
      if (!a.settings[s[0]] || !b.settings[s[0]]) skip[{ fields: 'フィールド', views: '一覧', reports: 'グラフ', appAcl: 'アプリのアクセス権', recordAcl: 'レコードのアクセス権', fieldAcl: 'フィールドのアクセス権', status: 'プロセス管理', general: 'アプリの条件通知', perRecord: 'レコードの条件通知', reminder: 'リマインダーの条件通知', actions: 'アプリアクション', appSettings: 'アプリの設定', plugins: 'プラグイン', customize: 'カスタマイズ' }[s[0]]] = true;
    });
    return rows.filter(function (r) { return !skip[r.group]; });
  }

  var KIND_CLASS = { 'このアプリだけ': 'a', '比べるアプリだけ': 'b', '違う': 'diff' };

  function renderTable(container, rows, nameA, nameB) {
    container.innerHTML = '';
    var thead = el('thead', {}, [el('tr', {}, ['種類', '項目', '違い', 'このアプリ（' + nameA + '）', '比べるアプリ（' + nameB + '）'].map(function (h) { return el('th', { text: h }); }))]);
    var body = rows.length ? rows.map(function (r) {
      return el('tr', {}, [
        el('td', { text: r.group }),
        el('td', { text: r.item }),
        el('td', { class: P + '-kind ' + P + '-kind-' + KIND_CLASS[r.kind], text: r.kind }),
        el('td', { class: P + '-detail', text: r.a }),
        el('td', { class: P + '-detail', text: r.b })
      ]);
    }) : [el('tr', {}, [el('td', { colspan: '5', class: P + '-empty', text: '違いは見つかりませんでした' })])];
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
    if (!current.rows) return;
    var d = new Date(); var stamp = d.getFullYear() + ('0' + (d.getMonth() + 1)).slice(-2) + ('0' + d.getDate()).slice(-2);
    saveExcel(buildExcel(current.rows, current.nameA, current.nameB, { host: location.hostname, at: d }), '設定の比較_' + safeFileName(current.nameA) + '_' + safeFileName(current.nameB) + '_' + stamp + '.xlsx');
  }

  function downloadCsv() {
    if (!current.rows) return;
    var blob = new Blob([buildCsv(current.rows, current.nameA, current.nameB)], { type: 'text/csv;charset=utf-8' });
    var url = URL.createObjectURL(blob);
    var d = new Date();
    var stamp = d.getFullYear() + ('0' + (d.getMonth() + 1)).slice(-2) + ('0' + d.getDate()).slice(-2);
    var a = el('a', { href: url, download: '設定の比較_' + safeFileName(current.nameA) + '_' + safeFileName(current.nameB) + '_' + stamp + '.csv' });
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    setTimeout(function () { URL.revokeObjectURL(url); }, 1000);
  }

  // 有料版のご案内（画面の下に小さく1行）。アプリの一覧を読んだら、この環境のアプリの数を入れる
  var moreText = null;
  function moreLink() {
    moreText = document.createTextNode('すべてのアプリをまとめて点検するなら：');
    return el('p', { class: P + '-more' }, [
      moreText,
      el('a', { href: 'https://daichoya.github.io/pro.html', target: '_blank', rel: 'noopener', text: '有料の「棚卸し・変更監視パック」（月額3,300円から。件数だけなら無料、30日間すべて無料で試せます）' })
    ]);
  }
  function moreCount(n) { if (moreText && n > 1) moreText.textContent = 'この環境の' + n + '個のアプリをまとめて点検するなら：'; }

  function openModal() {
    if (document.getElementById(P + '-overlay')) return;
    var thisApp = String(kintone.app.getId());
    var apps = [];
    var names = {};
    var appSelect = el('select', { class: P + '-select', 'aria-label': '比べるアプリ', disabled: 'disabled' });
    var runBtn = el('button', { type: 'button', class: P + '-btn ' + P + '-primary', text: '比べる', disabled: 'disabled' });
    var excelBtn = el('button', { type: 'button', class: P + '-btn', text: 'Excelで保存', disabled: 'disabled' });
    var csvBtn = el('button', { type: 'button', class: P + '-btn', text: 'CSVで書き出す', disabled: 'disabled' });
    var closeBtn = el('button', { type: 'button', class: P + '-close', 'aria-label': '閉じる', text: '×' });
    var status = el('p', { class: P + '-status', role: 'status' });
    var summary = el('div', { class: P + '-summary', hidden: 'hidden' });
    var result = el('div', { class: P + '-result' });
    var dialog = el('div', { class: P + '-dialog', role: 'dialog', 'aria-modal': 'true', 'aria-label': 'アプリ設定の比較' }, [
      el('div', { class: P + '-head' }, [el('h2', { text: 'アプリ設定の比較' }), closeBtn]),
      el('p', { class: P + '-lead', text: 'このアプリと、もう1つのアプリの設定を比べて、違うところを一覧にします。本番と検証のアプリ、コピー元とコピー先のアプリの確認に使えます。' }),
      el('div', { class: P + '-controls' }, [el('span', { class: P + '-label', text: '比べるアプリ' }), appSelect, runBtn, excelBtn, csvBtn]),
      status, summary, result,
      el('p', { class: P + '-note', text: '比べるのは、フィールド・一覧・グラフ・アクセス権・プロセス管理・通知・アプリアクション・アプリの設定・プラグイン・カスタマイズです。フォームの配置（レイアウト）、プラグインの設定の中、JavaScriptのファイルの中身は比べません。両方のアプリの管理権限が必要です。' }),
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
      var other = appSelect.value;
      if (!other) return;
      var nameA = names[thisApp] || ('アプリ' + thisApp);
      var nameB = names[other] || ('アプリ' + other);
      runBtn.setAttribute('disabled', 'disabled');
      csvBtn.setAttribute('disabled', 'disabled');
      summary.hidden = true;
      result.innerHTML = '';
      status.textContent = '2つのアプリの設定を読み込み中…';
      Promise.all([loadSettings(thisApp, 'このアプリ'), loadSettings(other, '比べるアプリ')]).then(function (loaded) {
        var a = loaded[0];
        var b = loaded[1];
        if (!a.settings.fields || !b.settings.fields) throw { message: 'フィールドの設定を読み取れませんでした。両方のアプリの管理権限があるか確かめてください' };
        var rows = dropUnreadable(compare(normalize(a.settings, names), normalize(b.settings, names)), a, b);
        current = { rows: rows, nameA: nameA, nameB: nameB };
        status.textContent = '';
        summary.innerHTML = '';
        summaryLines(rows, { errors: a.errors.concat(b.errors) }).forEach(function (line) {
          summary.appendChild(el('p', line.alert ? { class: P + '-alert', text: line.text } : { text: line.text }));
        });
        summary.hidden = false;
        renderTable(result, rows, nameA, nameB);
        if (rows.length) csvBtn.removeAttribute('disabled');
      }).catch(function (error) {
        status.textContent = '比べられませんでした（' + errorText(error) + '）。';
      }).then(function () {
        runBtn.removeAttribute('disabled');
      });
    });

    document.body.appendChild(overlay);
    status.textContent = 'アプリの一覧を読み込み中…';
    getAllApps().then(function (list) {
      moreCount(list.length);
      apps = list;
      list.forEach(function (app) { names[String(app.appId)] = app.name; });
      var others = list.filter(function (app) { return String(app.appId) !== thisApp; });
      others.forEach(function (app) { appSelect.appendChild(el('option', { value: app.appId, text: app.name + '（ID: ' + app.appId + '）' })); });
      if (!others.length) { status.textContent = '比べられるアプリがありません。'; return; }
      appSelect.removeAttribute('disabled');
      runBtn.removeAttribute('disabled');
      status.textContent = '比べるアプリを選んで「比べる」を押してください。';
    }).catch(function (error) {
      status.textContent = 'アプリの一覧を取得できませんでした（' + errorText(error) + '）。';
    });
  }

  kintone.events.on('app.record.index.show', function (event) {
    var space = kintone.app.getHeaderMenuSpaceElement();
    if (!space || document.getElementById(P + '-open')) return event;
    var button = el('button', { id: P + '-open', type: 'button', class: P + '-btn', text: '設定の比較' });
    button.addEventListener('click', openModal);
    space.appendChild(button);
    return event;
  });
})();
