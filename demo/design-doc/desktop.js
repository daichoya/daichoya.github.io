/*
 * 台帳屋 設計書ジェネレーター for kintone
 * 一覧画面に「設計書を作成」ボタンを追加し、選んだアプリの設定を
 * 人が読める設計書（Excel形式 .xlsx）として出力する。
 * Excelファイルは外部ライブラリを使わずブラウザ内で組み立て、データは外部に送らない。
 */
(function () {
  'use strict';

  // ===== 1. Excel（xlsx）ファイルの組み立て =====
  // 共通部品（js/xlsx.js）で作る。表紙・目次・見出しの色・固定行・フィルター・印刷設定は、台帳屋のどの製品でも同じ体裁にそろえる。
  var Xlsx = typeof module !== 'undefined' && module.exports ? require('../../../common/js/xlsx.js') : window.DaichoyaXlsx;
  var STYLE = Xlsx.STYLE;
  var crc32 = Xlsx.crc32;
  var buildZip = Xlsx.buildZip;
  var sheetName = Xlsx.sheetName;
  var columnName = Xlsx.columnName;

  function buildXlsx(sheets) {
    var title = (sheets[0] && sheets[0].docTitle) || 'アプリ設計書';
    return Xlsx.build(sheets, { title: title, creator: '台帳屋 設計書ジェネレーター', footer: '台帳屋 設計書ジェネレーター' });
  }

  // ===== 2. kintoneの設定を設計書の表に変換する（純粋な関数。テスト対象） =====

  var FIELD_TYPES = {
    SINGLE_LINE_TEXT: '文字列（1行）', MULTI_LINE_TEXT: '文字列（複数行）', RICH_TEXT: 'リッチエディター',
    NUMBER: '数値', CALC: '計算', RADIO_BUTTON: 'ラジオボタン', CHECK_BOX: 'チェックボックス',
    MULTI_SELECT: '複数選択', DROP_DOWN: 'ドロップダウン', DATE: '日付', TIME: '時刻', DATETIME: '日時',
    LINK: 'リンク', FILE: '添付ファイル', USER_SELECT: 'ユーザー選択', ORGANIZATION_SELECT: '組織選択',
    GROUP_SELECT: 'グループ選択', REFERENCE_TABLE: '関連レコード一覧', SUBTABLE: 'テーブル', GROUP: 'グループ',
    RECORD_NUMBER: 'レコード番号', CREATOR: '作成者', CREATED_TIME: '作成日時', MODIFIER: '更新者',
    UPDATED_TIME: '更新日時', STATUS: 'ステータス', STATUS_ASSIGNEE: '作業者', CATEGORY: 'カテゴリー'
  };

  var ENTITY_LABELS = { USER: 'ユーザー', GROUP: 'グループ', ORGANIZATION: '組織', CREATOR: '作成者', FIELD_ENTITY: 'フィールド', CUSTOM_FIELD: 'フィールド' };
  var ASSIGNEE_TYPES = { ONE: '次のユーザーから作業者を選択', ALL: '次のユーザー全員', ANY: '次のユーザーのうち1人' };
  var VIEW_TYPES = { LIST: '表形式', CALENDAR: 'カレンダー', CUSTOM: 'カスタマイズ' };
  var ACCESS_LABELS = { WRITE: '閲覧・編集', READ: '閲覧のみ', NONE: '権限なし' };

  function mark(value) { return value ? '○' : ''; }

  function entityText(entity) {
    if (!entity) return '';
    if (entity.type === 'CREATOR') return '作成者';
    return (ENTITY_LABELS[entity.type] || entity.type) + '：' + (entity.code || '');
  }

  function stripHtml(html) {
    return String(html || '').replace(/<br\s*\/?>/gi, '\n').replace(/<[^>]*>/g, '')
      .replace(/&nbsp;/g, ' ').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&amp;/g, '&').trim();
  }

  function sortedOptions(options) {
    return Object.keys(options || {}).map(function (k) { return options[k]; })
      .sort(function (a, b) { return Number(a.index) - Number(b.index); })
      .map(function (o) { return o.label; });
  }

  function fieldTypeLabel(prop) {
    var base = FIELD_TYPES[prop.type] || prop.type;
    return prop.lookup ? 'ルックアップ（' + base + '）' : base;
  }

  // names：アプリIDからアプリ名（参照先を名前で書く。分からないときはIDだけ）
  var appLabel = function (id, names) { return names && names[id] ? '「' + names[id] + '」（アプリID ' + id + '）' : 'アプリID ' + id; };
  function fieldDetail(prop, names) {
    var parts = [];
    if (prop.lookup) {
      parts.push('参照先：' + appLabel(prop.lookup.relatedApp.app, names) + ' の「' + prop.lookup.relatedKeyField + '」');
      var maps = (prop.lookup.fieldMappings || []).map(function (m) { return m.relatedField + '→' + m.field; });
      if (maps.length) parts.push('コピー：' + maps.join('、'));
    }
    if (prop.options) parts.push('選択肢：' + sortedOptions(prop.options).join('、'));
    if (prop.expression) parts.push('計算式：' + prop.expression);
    if (prop.referenceTable) {
      var ref = prop.referenceTable;
      parts.push('参照先：' + appLabel(ref.relatedApp.app, names));
      if (ref.condition) parts.push('条件：' + ref.condition.field + ' ＝ 参照先の' + ref.condition.relatedField);
      if (ref.filterCond) parts.push('絞り込み：' + ref.filterCond);
    }
    if (prop.minValue || prop.maxValue) parts.push('範囲：' + (prop.minValue || '') + '〜' + (prop.maxValue || ''));
    if (prop.minLength || prop.maxLength) parts.push('文字数：' + (prop.minLength || '') + '〜' + (prop.maxLength || ''));
    if (prop.unit) parts.push('単位：' + prop.unit);
    if (prop.protocol) parts.push('リンクの種類：' + prop.protocol);
    return parts.join('\n');
  }

  function defaultText(prop) {
    var v = prop.defaultValue;
    if (Array.isArray(v)) {
      return v.map(function (x) { return typeof x === 'object' ? (x.code || '') : x; }).join('、');
    }
    if (prop.defaultNowValue) return '登録時の日時';
    return v === undefined || v === null ? '' : String(v);
  }

  // フォームの配置順にフィールドを並べ、配置にないもの（システム項目など）を最後に付ける
  function fieldRows(properties, layout, names) {
    var order = [];
    function walk(items, container) {
      (items || []).forEach(function (item) {
        if (item.type === 'ROW') {
          (item.fields || []).forEach(function (f) { if (f.code) order.push({ code: f.code, container: container }); });
        } else if (item.type === 'SUBTABLE') {
          order.push({ code: item.code, container: container });
          (item.fields || []).forEach(function (f) { if (f.code) order.push({ code: f.code, container: 'テーブル：' + item.code, table: item.code }); });
        } else if (item.type === 'GROUP') {
          order.push({ code: item.code, container: container });
          walk(item.layout, 'グループ：' + item.code);
        }
      });
    }
    walk(layout && layout.layout, '');

    var seen = {};
    var rows = [];
    function push(prop, container) {
      seen[prop.code] = true;
      rows.push([
        rows.length + 1, prop.label || '', prop.code, fieldTypeLabel(prop),
        mark(prop.required), mark(prop.unique), defaultText(prop), fieldDetail(prop, names), container
      ]);
    }
    order.forEach(function (o) {
      var prop = o.table ? (properties[o.table] && properties[o.table].fields && properties[o.table].fields[o.code]) : properties[o.code];
      if (prop && !seen[o.code]) push(prop, o.container);
    });
    Object.keys(properties || {}).forEach(function (code) {
      var prop = properties[code];
      if (!seen[code] && FIELD_TYPES[prop.type]) push(prop, '（システム項目）');
    });
    return rows;
  }

  function viewRows(views, labels) {
    return Object.keys((views && views.views) || {}).map(function (k) { return views.views[k]; })
      .sort(function (a, b) { return Number(a.index) - Number(b.index); })
      .map(function (v, i) {
        var fields = (v.fields || []).map(function (c) { return (labels && labels[c]) || c; }).join('、');
        return [i + 1, v.name, VIEW_TYPES[v.type] || v.type, v.filterCond || '（すべてのレコード）', v.sort || '', fields];
      });
  }

  function processBlocks(status) {
    if (!status || !status.enable) {
      return [{ title: 'プロセス管理', header: [], rows: [], note: 'このアプリではプロセス管理を使っていません。' }];
    }
    var states = Object.keys(status.states || {}).map(function (k) { return status.states[k]; })
      .sort(function (a, b) { return Number(a.index) - Number(b.index); })
      .map(function (s, i) {
        var a = s.assignee || {};
        var who = (a.entities || []).map(function (e) { return entityText(e.entity); }).join('、');
        return [i + 1, s.name, (a.entities && a.entities.length) ? (ASSIGNEE_TYPES[a.type] || a.type || '') : '（作業者なし）', who];
      });
    var actions = (status.actions || []).map(function (a, i) {
      return [i + 1, a.name, a.from, a.to, a.filterCond || '（条件なし）'];
    });
    return [
      { title: 'ステータス', header: ['順', 'ステータス名', '作業者の決め方', '作業者'], rows: states },
      { title: 'アクション（ステータスの移り方）', header: ['順', 'アクション名', '実行前', '実行後', '実行できる条件'], rows: actions }
    ];
  }

  function timingText(t) {
    if (!t) return '';
    var parts = ['「' + t.code + '」の'];
    var days = Number(t.daysLater || 0);
    if (days) parts.push(Math.abs(days) + '日' + (days > 0 ? '後' : '前'));
    if (t.time) parts.push(t.time);
    var hours = Number(t.hoursLater || 0);
    if (hours) parts.push(Math.abs(hours) + '時間' + (hours > 0 ? '後' : '前'));
    var minutes = Number(t.minutesLater || 0);
    if (minutes) parts.push(Math.abs(minutes) + '分' + (minutes > 0 ? '後' : '前'));
    if (parts.length === 1) parts.push('当日');
    return parts.join(' ');
  }

  function notificationBlocks(general, perRecord, reminder) {
    var generalRows = ((general && general.notifications) || []).map(function (n) {
      return [entityText(n.entity), mark(n.recordAdded), mark(n.recordEdited), mark(n.commentAdded), mark(n.fileImported)];
    });
    var perRecordRows = ((perRecord && perRecord.notifications) || []).map(function (n) {
      return [n.filterCond || '（すべてのレコード）', n.title || '', (n.targets || []).map(function (t) { return entityText(t.entity); }).join('、')];
    });
    var reminderRows = ((reminder && reminder.notifications) || []).map(function (n) {
      return [timingText(n.timing), n.filterCond || '（すべてのレコード）', n.title || '', (n.targets || []).map(function (t) { return entityText(t.entity); }).join('、')];
    });
    return [
      { title: 'アプリの条件通知', header: ['通知先', 'レコード追加', 'レコード編集', 'コメント書き込み', 'ファイル読み込み'], rows: generalRows, note: general && general.notifyToCommenter ? 'コメントの書き込み者にも通知します。' : null },
      { title: 'レコードの条件通知', header: ['条件', '通知内容', '通知先'], rows: perRecordRows },
      { title: 'リマインダーの条件通知', header: ['タイミング', '条件', '通知内容', '通知先'], rows: reminderRows }
    ];
  }

  function permissionBlocks(appAcl, recordAcl, fieldAcl, labels) {
    var appRows = ((appAcl && appAcl.rights) || []).map(function (r, i) {
      return [i + 1, entityText(r.entity), mark(r.includeSubs), mark(r.appEditable), mark(r.recordViewable), mark(r.recordAddable),
        mark(r.recordEditable), mark(r.recordDeletable), mark(r.recordImportable), mark(r.recordExportable)];
    });
    var recordRows = [];
    ((recordAcl && recordAcl.rights) || []).forEach(function (rule, i) {
      (rule.entities || []).forEach(function (e) {
        recordRows.push([i + 1, rule.filterCond || '（すべてのレコード）', entityText(e.entity), mark(e.viewable), mark(e.editable), mark(e.deletable), mark(e.includeSubs)]);
      });
    });
    var fieldRowsAcl = [];
    ((fieldAcl && fieldAcl.rights) || []).forEach(function (rule) {
      (rule.entities || []).forEach(function (e) {
        fieldRowsAcl.push([(labels && labels[rule.code]) || rule.code, rule.code, entityText(e.entity), ACCESS_LABELS[e.accessibility] || e.accessibility, mark(e.includeSubs)]);
      });
    });
    return [
      { title: 'アプリのアクセス権（上にあるものほど優先）', header: ['順', '対象', '下位組織', 'アプリ管理', '閲覧', '追加', '編集', '削除', '読み込み', '書き出し'], rows: appRows },
      { title: 'レコードのアクセス権（上の条件ほど優先）', header: ['順', '条件', '対象', '閲覧', '編集', '削除', '下位組織'], rows: recordRows },
      { title: 'フィールドのアクセス権', header: ['フィールド名', 'フィールドコード', '対象', '権限', '下位組織'], rows: fieldRowsAcl }
    ];
  }

  // JavaScript・CSSによるカスタマイズ（/k/v1/app/customize.json）。中身ではなく、読み込んでいるファイルの一覧を書く
  var SCOPE_TEXT = { ALL: 'すべてのユーザー', ADMIN: 'アプリの管理者だけ', NONE: '適用しない' };
  function customizeBlocks(customize) {
    var c = customize || {};
    var rows = [];
    [['desktop', 'PC'], ['mobile', 'モバイル']].forEach(function (screen) {
      var part = c[screen[0]] || {};
      [['js', 'JavaScript'], ['css', 'CSS']].forEach(function (kind) {
        (part[kind[0]] || []).forEach(function (item, i) {
          var isUrl = item.type === 'URL';
          rows.push([screen[1], kind[1], String(i + 1), isUrl ? 'URL' : 'ファイル', isUrl ? (item.url || '') : ((item.file && item.file.name) || '')]);
        });
      });
    });
    return [
      { title: '適用範囲', header: ['項目', '内容'], rows: [['カスタマイズを適用するユーザー', SCOPE_TEXT[c.scope] || c.scope || '不明']] },
      { title: '読み込んでいるファイル（同じ画面・種類の中では、順の小さいものから読み込まれます）', header: ['画面', '種類', '順', '形式', 'ファイル名・URL'], rows: rows, note: rows.length ? 'ファイルの中身は設計書に含めていません。' : 'JavaScript・CSSによるカスタマイズはありません。' }
    ];
  }

  function labelMap(properties) {
    var map = {};
    Object.keys(properties || {}).forEach(function (code) {
      var p = properties[code];
      map[code] = p.label || code;
      if (p.type === 'SUBTABLE' && p.fields) {
        Object.keys(p.fields).forEach(function (inner) { map[inner] = p.fields[inner].label || inner; });
      }
    });
    return map;
  }

  var NO_PERMISSION = 'この項目を取得できませんでした。アプリの管理権限が必要です。';

  function sectionError(error) {
    var code = error && error.code;
    if (code === 'CB_NO02' || code === 'GAIA_NO01') return NO_PERMISSION;
    return 'この項目を取得できませんでした（' + ((error && error.message) || '原因不明') + '）。';
  }

  // results: 各APIの結果 { ok: true, value } か { ok: false, error }
  // このアプリが参照しているアプリ（ルックアップ・関連レコード一覧・アクションの作成先）。names：アプリIDからアプリ名
  function relatedRows(properties, actions, names) {
    var rows = [];
    var name = function (id) { return id ? ((names && names[id]) || 'アプリID ' + id) : '見られないアプリ'; };
    function walk(props, table) {
      Object.keys(props || {}).forEach(function (code) {
        var p = props[code];
        var place = (p.label || code) + (table ? '（テーブル「' + table + '」内）' : '');
        if (p.lookup && p.lookup.relatedApp) rows.push([name(p.lookup.relatedApp.app), 'ルックアップ：' + place]);
        if (p.referenceTable && p.referenceTable.relatedApp) rows.push([name(p.referenceTable.relatedApp.app), '関連レコード一覧：' + place]);
        if (p.type === 'SUBTABLE') walk(p.fields, p.label || code);
      });
    }
    walk(properties, null);
    var list = (actions && actions.actions) || {};
    Object.keys(list).map(function (k) { return list[k]; }).sort(function (a, b) { return Number(a.index) - Number(b.index); }).forEach(function (a) {
      rows.push([name(a.destApp && a.destApp.app), 'アクション（レコードの作成先）：' + (a.name || '')]);
    });
    return rows;
  }

  function buildSheets(results, generatedAt, names) {
    function val(key) { return results[key] && results[key].ok ? results[key].value : null; }
    function err(key) { return results[key] && !results[key].ok ? sectionError(results[key].error) : null; }
    var app = val('app') || {};
    var fields = val('fields');
    var properties = (fields && fields.properties) || {};
    var labels = labelMap(properties);
    var fieldList = fields ? fieldRows(properties, val('layout'), names) : [];
    var views = viewRows(val('views'), labels);
    var plugins = ((val('plugins') && val('plugins').plugins) || []).map(function (p) {
      return [p.name, p.id, p.enabled === false ? '無効' : '有効'];
    });

    var overview = [
      ['アプリ名', app.name || ''],
      ['アプリID', app.appId || ''],
      ['説明', stripHtml(app.description)],
      ['アプリ管理者用メモ', val('adminNotes') ? (stripHtml(val('adminNotes').content) || '（なし）') : '（読み取れませんでした）'],
      ['作成者', app.creator ? (app.creator.name || app.creator.code) : ''],
      ['作成日時', dateText(app.createdAt)],
      ['最終更新者', app.modifier ? (app.modifier.name || app.modifier.code) : ''],
      ['最終更新日時', dateText(app.modifiedAt)],
      ['フィールド数', String(fieldList.length)],
      ['一覧の数', String(views.length)],
      ['プロセス管理', val('status') ? (val('status').enable ? '使用' : '未使用') : '不明'],
      ['設計書の作成日時', generatedAt]
    ];

    function withError(blocks, key) {
      var e = err(key);
      if (!e) return blocks;
      return [{ title: blocks[0] && blocks[0].title, header: [], rows: [], note: e }];
    }
    var permission = permissionBlocks(val('appAcl'), val('recordAcl'), val('fieldAcl'), labels);
    ['appAcl', 'recordAcl', 'fieldAcl'].forEach(function (key, i) {
      if (err(key)) permission[i] = { title: permission[i].title, header: [], rows: [], note: err(key) };
    });
    var notify = notificationBlocks(val('general'), val('perRecord'), val('reminder'));
    ['general', 'perRecord', 'reminder'].forEach(function (key, i) {
      if (err(key)) notify[i] = { title: notify[i].title, header: [], rows: [], note: err(key) };
    });

    var related = relatedRows(properties, val('actions'), names);
    var subText = 'アプリ設計書：' + (app.name || '') + '（アプリID ' + (app.appId || '') + '）　／　作成日時：' + generatedAt;
    var head = function (title) { return { title: title, sub: subText }; };

    return [
      { name: '概要', widths: [20, 12, 46, 18], orientation: 'portrait', tabColor: 'FF1F3A5F', docTitle: 'アプリ設計書：' + (app.name || ''),
        heading: { title: 'アプリ設計書', sub: (app.name || '') + '（アプリID ' + (app.appId || '') + '）　／　作成日時：' + generatedAt + '　／　作成：台帳屋 設計書ジェネレーター' },
        blocks: [
          { title: 'アプリの概要', header: ['項目', '内容'], rows: overview, labels: true, span: 3 },
          { title: '改訂履歴', header: ['版', '日付', '内容', '作成者'], rows: [['1.0', String(generatedAt || '').slice(0, 10), '設計書ジェネレーターで作成', ''], ['', '', '', ''], ['', '', '', '']],
            note: '設計書を直したときは、版・日付・内容・作成者を書き足してください。' },
          Object.assign(Xlsx.contents([
            ['フィールド', '全フィールドの名前・コード・種類・必須・初期値・選択肢や計算式・配置'],
            ['一覧', '一覧の種類・絞り込み条件・並び順・表示するフィールド'],
            ['プロセス管理', 'ステータス・作業者・アクション'],
            ['通知', 'アプリの条件通知・レコードの条件通知・リマインダーの条件通知'],
            ['アクセス権', 'アプリ・レコード・フィールドのアクセス権'],
            ['プラグイン', '追加しているプラグインと、有効・無効'],
            ['カスタマイズ', 'JavaScript・CSSのファイルと、適用するユーザー']
          ]), { span: 3 }),
          // 2ページ目：運用の情報（手で書く欄）と、関係するアプリ。表紙の1ページ目は概要・改訂履歴・目次に収める
          { title: '運用の情報', header: ['項目', '内容'], rows: [['目的・使い方', ''], ['利用する部署・人', ''], ['アプリの担当者', ''], ['見直しの時期', '']], labels: true, span: 3, pageBreak: true,
            note: 'kintoneの設定にはない情報です。引き継ぎや監査の前に書き足してください。' },
          { title: '関係するアプリ（このアプリが参照している先）', header: ['相手のアプリ', 'つながりの種類と、設定のある場所'], rows: related, span: 3,
            note: related.length ? 'このアプリを参照しているアプリは、アプリ相関マップで確かめられます。' : 'ルックアップ・関連レコード一覧・アクションで参照しているアプリはありません。' }
        ] },
      { name: 'フィールド', list: true, heading: head('フィールド'), widths: [6, 22, 22, 20, 7, 9, 16, 50, 22], blocks: withError([{ title: null, ids: [0], centers: [4, 5], header: ['順', 'フィールド名', 'フィールドコード', '種類', '必須', '重複禁止', '初期値', '詳細（選択肢・計算式・参照先など）', '配置'], rows: fieldList }], 'fields') },
      { name: '一覧', list: true, heading: head('一覧'), widths: [6, 22, 14, 40, 30, 50], blocks: withError([{ title: null, ids: [0], header: ['順', '一覧名', '種類', '絞り込み条件', '並び順', '表示するフィールド'], rows: views }], 'views') },
      { name: 'プロセス管理', heading: head('プロセス管理'), widths: [6, 24, 28, 40, 40], blocks: withError(processBlocks(val('status')), 'status') },
      { name: '通知', heading: head('通知'), widths: [30, 30, 30, 30, 16], blocks: notify },
      { name: 'アクセス権', heading: head('アクセス権'), widths: [22, 30, 20, 16, 10, 10, 10, 10, 10, 10], blocks: permission },
      { name: 'プラグイン', list: true, heading: head('プラグイン'), widths: [40, 36, 10], blocks: withError([{ title: null, header: ['プラグイン名', 'プラグインID', '状態'], rows: plugins }], 'plugins') },
      { name: 'カスタマイズ', heading: head('カスタマイズ'), widths: [30, 14, 6, 10, 70], blocks: err('customize') ? [{ title: null, header: [], rows: [], note: err('customize') }] : customizeBlocks(val('customize')) }
    ];
  }

  // kintoneの日時（UTCのISO形式）を、見ている人の時刻で「2026-11-04 23:03」の形にする
  function dateText(value) {
    if (!value) return '';
    var d = value instanceof Date ? value : new Date(value);
    if (isNaN(d.getTime())) return String(value);
    var pad = function (n) { return ('0' + n).slice(-2); };
    return d.getFullYear() + '-' + pad(d.getMonth() + 1) + '-' + pad(d.getDate()) + ' ' + pad(d.getHours()) + ':' + pad(d.getMinutes());
  }

  function safeFileName(name) {
    var s = String(name || '').replace(/[\\/:*?"<>|]/g, '_').trim();
    return s || 'app';
  }

  function summaryText(sheets) {
    var count = function (name) {
      var s = sheets.filter(function (x) { return x.name === name; })[0];
      return s ? s.blocks.reduce(function (n, b) { return n + (b.rows ? b.rows.length : 0); }, 0) : 0;
    };
    var failed = [];
    sheets.forEach(function (s) {
      if (s.blocks.some(function (b) { return /^この項目を取得できませんでした/.test(b.note || ''); })) failed.push(s.name);
    });
    var part = function (name) { return failed.indexOf(name) >= 0 ? name + '（読み取れませんでした）' : name + count(name) + '件'; };
    var custom = sheets.filter(function (x) { return x.name === 'カスタマイズ'; })[0];
    var files = custom && custom.blocks[1] ? custom.blocks[1].rows.length : 0;
    return part('フィールド') + '、' + part('一覧') + '、' + part('プラグイン') +
      (files ? '、カスタマイズのファイル' + files + '件' : '') + 'を設計書にまとめました。' +
      (failed.length ? '権限がないため読み取れなかった項目があります（' + failed.join('・') + 'のシートに書いてあります）。' : '');
  }

  var Core = {
    crc32: crc32,
    buildZip: buildZip,
    buildXlsx: buildXlsx,
    sheetName: sheetName,
    columnName: columnName,
    fieldRows: fieldRows,
    viewRows: viewRows,
    processBlocks: processBlocks,
    notificationBlocks: notificationBlocks,
    timingText: timingText,
    permissionBlocks: permissionBlocks,
    customizeBlocks: customizeBlocks,
    buildSheets: buildSheets,
    stripHtml: stripHtml,
    safeFileName: safeFileName,
    dateText: dateText,
    summaryText: summaryText
  };

  if (typeof module !== 'undefined' && module.exports) {
    module.exports = Core;
    return;
  }
  if (typeof kintone === 'undefined') return;

  // ===== 3. 画面（kintone上でだけ動く部分） =====

  var P = 'dcy-dd';

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

  function settle(promise) {
    return promise.then(function (value) { return { ok: true, value: value }; },
      function (error) { return { ok: false, error: error }; });
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

  function collect(appId) {
    var requests = {
      app: api('/k/v1/app.json', { id: appId }),
      fields: api('/k/v1/app/form/fields.json', { app: appId }),
      layout: api('/k/v1/app/form/layout.json', { app: appId }),
      views: api('/k/v1/app/views.json', { app: appId }),
      status: api('/k/v1/app/status.json', { app: appId }),
      general: api('/k/v1/app/notifications/general.json', { app: appId }),
      perRecord: api('/k/v1/app/notifications/perRecord.json', { app: appId }),
      reminder: api('/k/v1/app/notifications/reminder.json', { app: appId }),
      appAcl: api('/k/v1/app/acl.json', { app: appId }),
      recordAcl: api('/k/v1/record/acl.json', { app: appId }),
      fieldAcl: api('/k/v1/field/acl.json', { app: appId }),
      plugins: api('/k/v1/app/plugins.json', { app: appId }),
      customize: api('/k/v1/app/customize.json', { app: appId }),
      adminNotes: api('/k/v1/app/adminNotes.json', { app: appId }),
      actions: api('/k/v1/app/actions.json', { app: appId })
    };
    var keys = Object.keys(requests);
    return Promise.all(keys.map(function (k) { return settle(requests[k]); })).then(function (list) {
      var results = {};
      keys.forEach(function (k, i) { results[k] = list[i]; });
      return results;
    });
  }

  function nowText() {
    return dateText(new Date());
  }

  function download(bytes, fileName) {
    var blob = new Blob([bytes], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    var url = URL.createObjectURL(blob);
    var a = el('a', { href: url, download: fileName });
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
    var select = el('select', { class: P + '-select', 'aria-label': '設計書を作るアプリ' });
    var runBtn = el('button', { type: 'button', class: P + '-btn ' + P + '-primary', text: 'Excelで設計書を作成' });
    var closeBtn = el('button', { type: 'button', class: P + '-close', 'aria-label': '閉じる', text: '×' });
    var status = el('p', { class: P + '-status', role: 'status' });
    var dialog = el('div', { class: P + '-dialog', role: 'dialog', 'aria-modal': 'true', 'aria-label': '設計書ジェネレーター' }, [
      el('div', { class: P + '-head' }, [el('h2', { text: '設計書ジェネレーター' }), closeBtn]),
      el('p', { class: P + '-lead', text: 'フィールド・一覧・プロセス管理・通知・アクセス権・プラグイン・カスタマイズを、1つのExcelファイルにまとめます。' }),
      el('div', { class: P + '-controls' }, [select, runBtn]),
      status,
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

    runBtn.addEventListener('click', function () {
      var appId = select.value;
      if (!appId) return;
      runBtn.disabled = true;
      status.textContent = '設定を読み込み中…';
      collect(appId).then(function (results) {
        var sheets = buildSheets(results, nowText(), appNames);
        var stamp = nowText().slice(0, 10).replace(/-/g, '');
        download(buildXlsx(sheets), '設計書_' + safeFileName(appNames[appId] || ('app' + appId)) + '_' + stamp + '.xlsx');
        status.textContent = summaryText(sheets);
      }).catch(function (error) {
        status.textContent = '設計書を作成できませんでした（' + ((error && error.message) || '原因不明') + '）。';
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
      status.textContent = 'アプリを選んで「Excelで設計書を作成」を押してください。';
    }).catch(function (error) {
      status.textContent = 'アプリの一覧を取得できませんでした（' + ((error && error.message) || '原因不明') + '）。';
    });
  }

  kintone.events.on('app.record.index.show', function (event) {
    var space = kintone.app.getHeaderMenuSpaceElement();
    if (!space || document.getElementById(P + '-open')) return event;
    var button = el('button', { id: P + '-open', type: 'button', class: P + '-btn', text: '設計書を作成' });
    button.addEventListener('click', openModal);
    space.appendChild(button);
    return event;
  });
})();
