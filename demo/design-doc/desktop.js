/*
 * 台帳屋 設計書ジェネレーター for kintone
 * 一覧画面に「設計書を作成」ボタンを追加し、選んだアプリの設定を
 * 人が読める設計書（Excel形式 .xlsx）として出力する。
 * Excelファイルは外部ライブラリを使わずブラウザ内で組み立て、データは外部に送らない。
 */
(function () {
  'use strict';

  // ===== 1. Excel（xlsx）ファイルの組み立て =====
  // xlsxは「XMLファイルをまとめたzip」なので、無圧縮のzipを自前で作る。

  var CRC_TABLE = (function () {
    var table = new Uint32Array(256);
    for (var n = 0; n < 256; n++) {
      var c = n;
      for (var k = 0; k < 8; k++) c = (c & 1) ? (0xEDB88320 ^ (c >>> 1)) : (c >>> 1);
      table[n] = c >>> 0;
    }
    return table;
  })();

  function crc32(bytes) {
    var c = 0xFFFFFFFF;
    for (var i = 0; i < bytes.length; i++) c = CRC_TABLE[(c ^ bytes[i]) & 0xFF] ^ (c >>> 8);
    return (c ^ 0xFFFFFFFF) >>> 0;
  }

  function utf8(text) {
    return new TextEncoder().encode(text);
  }

  function concatBytes(chunks) {
    var total = chunks.reduce(function (sum, c) { return sum + c.length; }, 0);
    var out = new Uint8Array(total);
    var pos = 0;
    chunks.forEach(function (c) { out.set(c, pos); pos += c.length; });
    return out;
  }

  // files: [{ name: 'xl/workbook.xml', data: Uint8Array }]
  function buildZip(files) {
    var chunks = [];
    var central = [];
    var offset = 0;
    files.forEach(function (file) {
      var name = utf8(file.name);
      var data = file.data;
      var crc = crc32(data);

      var local = new Uint8Array(30 + name.length);
      var lv = new DataView(local.buffer);
      lv.setUint32(0, 0x04034b50, true);
      lv.setUint16(4, 20, true);
      lv.setUint16(6, 0x0800, true); // ファイル名はUTF-8
      lv.setUint16(8, 0, true);      // 無圧縮
      lv.setUint16(10, 0, true);
      lv.setUint16(12, 0x21, true);  // 1980-01-01
      lv.setUint32(14, crc, true);
      lv.setUint32(18, data.length, true);
      lv.setUint32(22, data.length, true);
      lv.setUint16(26, name.length, true);
      lv.setUint16(28, 0, true);
      local.set(name, 30);

      var entry = new Uint8Array(46 + name.length);
      var cv = new DataView(entry.buffer);
      cv.setUint32(0, 0x02014b50, true);
      cv.setUint16(4, 20, true);
      cv.setUint16(6, 20, true);
      cv.setUint16(8, 0x0800, true);
      cv.setUint16(10, 0, true);
      cv.setUint16(12, 0, true);
      cv.setUint16(14, 0x21, true);
      cv.setUint32(16, crc, true);
      cv.setUint32(20, data.length, true);
      cv.setUint32(24, data.length, true);
      cv.setUint16(28, name.length, true);
      cv.setUint16(30, 0, true);
      cv.setUint16(32, 0, true);
      cv.setUint16(34, 0, true);
      cv.setUint16(36, 0, true);
      cv.setUint32(38, 0, true);
      cv.setUint32(42, offset, true);
      entry.set(name, 46);

      chunks.push(local, data);
      central.push(entry);
      offset += local.length + data.length;
    });
    var centralBytes = concatBytes(central);
    var end = new Uint8Array(22);
    var ev = new DataView(end.buffer);
    ev.setUint32(0, 0x06054b50, true);
    ev.setUint16(8, files.length, true);
    ev.setUint16(10, files.length, true);
    ev.setUint32(12, centralBytes.length, true);
    ev.setUint32(16, offset, true);
    return concatBytes(chunks.concat([centralBytes, end]));
  }

  function xmlEscape(value) {
    return String(value === undefined || value === null ? '' : value)
      .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '')
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;');
  }

  function columnName(index) {
    var name = '';
    var n = index + 1;
    while (n > 0) {
      var m = (n - 1) % 26;
      name = String.fromCharCode(65 + m) + name;
      n = Math.floor((n - 1) / 26);
    }
    return name;
  }

  // シート名に使えない文字を除き、31文字以内にする
  function sheetName(name) {
    return String(name).replace(/[\[\]:*?\/\\]/g, '_').slice(0, 31) || 'Sheet';
  }

  var STYLE = { plain: 0, header: 1, body: 2, title: 3 };

  function cell(ref, value, style) {
    return '<c r="' + ref + '" t="inlineStr" s="' + style + '"><is><t xml:space="preserve">' +
      xmlEscape(value) + '</t></is></c>';
  }

  // sheet: { name, widths: [数値], blocks: [{ title, header: [], rows: [[]], note }] }
  function sheetXml(sheet) {
    var rowsXml = [];
    var r = 1;
    function addRow(values, style) {
      var cells = values.map(function (v, i) { return cell(columnName(i) + r, v, style); }).join('');
      rowsXml.push('<row r="' + r + '">' + cells + '</row>');
      r += 1;
    }
    sheet.blocks.forEach(function (block, i) {
      if (i > 0) r += 1; // 表と表の間を1行あける
      if (block.title) addRow([block.title], STYLE.title);
      if (block.header && block.header.length) addRow(block.header, STYLE.header);
      (block.rows || []).forEach(function (row) { addRow(row, STYLE.body); });
      if (block.note) addRow([block.note], STYLE.plain);
    });
    var cols = (sheet.widths || []).map(function (w, i) {
      return '<col min="' + (i + 1) + '" max="' + (i + 1) + '" width="' + w + '" customWidth="1"/>';
    }).join('');
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
      (cols ? '<cols>' + cols + '</cols>' : '') +
      '<sheetData>' + rowsXml.join('') + '</sheetData></worksheet>';
  }

  var STYLES_XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
    '<fonts count="2"><font><sz val="11"/><name val="Yu Gothic"/></font><font><b/><sz val="11"/><name val="Yu Gothic"/></font></fonts>' +
    '<fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>' +
    '<fill><patternFill patternType="solid"><fgColor rgb="FFDDEBF7"/><bgColor indexed="64"/></patternFill></fill></fills>' +
    '<borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border>' +
    '<border><left style="thin"><color rgb="FFBFBFBF"/></left><right style="thin"><color rgb="FFBFBFBF"/></right>' +
    '<top style="thin"><color rgb="FFBFBFBF"/></top><bottom style="thin"><color rgb="FFBFBFBF"/></bottom><diagonal/></border></borders>' +
    '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
    '<cellXfs count="4">' +
    '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>' +
    '<xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1"/>' +
    '<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf>' +
    '<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>' +
    '</cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>';

  function buildXlsx(sheets) {
    var names = sheets.map(function (s) { return sheetName(s.name); });
    var contentTypes = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
      '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' +
      sheets.map(function (s, i) {
        return '<Override PartName="/xl/worksheets/sheet' + (i + 1) + '.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>';
      }).join('') + '</Types>';
    var rootRels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>' +
      '</Relationships>';
    var workbook = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>' +
      names.map(function (n, i) {
        return '<sheet name="' + xmlEscape(n) + '" sheetId="' + (i + 1) + '" r:id="rId' + (i + 1) + '"/>';
      }).join('') + '</sheets></workbook>';
    var workbookRels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      sheets.map(function (s, i) {
        return '<Relationship Id="rId' + (i + 1) + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet' + (i + 1) + '.xml"/>';
      }).join('') +
      '<Relationship Id="rId' + (sheets.length + 1) + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' +
      '</Relationships>';
    var files = [
      { name: '[Content_Types].xml', data: utf8(contentTypes) },
      { name: '_rels/.rels', data: utf8(rootRels) },
      { name: 'xl/workbook.xml', data: utf8(workbook) },
      { name: 'xl/_rels/workbook.xml.rels', data: utf8(workbookRels) },
      { name: 'xl/styles.xml', data: utf8(STYLES_XML) }
    ];
    sheets.forEach(function (s, i) {
      files.push({ name: 'xl/worksheets/sheet' + (i + 1) + '.xml', data: utf8(sheetXml(s)) });
    });
    return buildZip(files);
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

  function fieldDetail(prop) {
    var parts = [];
    if (prop.lookup) {
      parts.push('参照先：アプリID ' + prop.lookup.relatedApp.app + ' の「' + prop.lookup.relatedKeyField + '」');
      var maps = (prop.lookup.fieldMappings || []).map(function (m) { return m.relatedField + '→' + m.field; });
      if (maps.length) parts.push('コピー：' + maps.join('、'));
    }
    if (prop.options) parts.push('選択肢：' + sortedOptions(prop.options).join('、'));
    if (prop.expression) parts.push('計算式：' + prop.expression);
    if (prop.referenceTable) {
      var ref = prop.referenceTable;
      parts.push('参照先：アプリID ' + ref.relatedApp.app);
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
  function fieldRows(properties, layout) {
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
        mark(prop.required), mark(prop.unique), defaultText(prop), fieldDetail(prop), container
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
  function buildSheets(results, generatedAt) {
    function val(key) { return results[key] && results[key].ok ? results[key].value : null; }
    function err(key) { return results[key] && !results[key].ok ? sectionError(results[key].error) : null; }
    var app = val('app') || {};
    var fields = val('fields');
    var properties = (fields && fields.properties) || {};
    var labels = labelMap(properties);
    var fieldList = fields ? fieldRows(properties, val('layout')) : [];
    var views = viewRows(val('views'), labels);
    var plugins = ((val('plugins') && val('plugins').plugins) || []).map(function (p) {
      return [p.name, p.id, p.enabled === false ? '無効' : '有効'];
    });

    var overview = [
      ['アプリ名', app.name || ''],
      ['アプリID', app.appId || ''],
      ['説明', stripHtml(app.description)],
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

    return [
      { name: '概要', widths: [20, 70], blocks: [{ title: null, header: ['項目', '内容'], rows: overview }] },
      { name: 'フィールド', widths: [6, 22, 22, 20, 7, 9, 16, 50, 22], blocks: withError([{ title: null, header: ['順', 'フィールド名', 'フィールドコード', '種類', '必須', '重複禁止', '初期値', '詳細（選択肢・計算式・参照先など）', '配置'], rows: fieldList }], 'fields') },
      { name: '一覧', widths: [6, 22, 14, 40, 30, 50], blocks: withError([{ title: null, header: ['順', '一覧名', '種類', '絞り込み条件', '並び順', '表示するフィールド'], rows: views }], 'views') },
      { name: 'プロセス管理', widths: [6, 24, 28, 40, 40], blocks: withError(processBlocks(val('status')), 'status') },
      { name: '通知', widths: [30, 30, 30, 30, 16], blocks: notify },
      { name: 'アクセス権', widths: [22, 30, 20, 16, 10, 10, 10, 10, 10, 10], blocks: permission },
      { name: 'プラグイン', widths: [40, 36, 10], blocks: withError([{ title: null, header: ['プラグイン名', 'プラグインID', '状態'], rows: plugins }], 'plugins') },
      { name: 'カスタマイズ', widths: [30, 14, 6, 10, 70], blocks: err('customize') ? [{ title: null, header: [], rows: [], note: err('customize') }] : customizeBlocks(val('customize')) }
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
      customize: api('/k/v1/app/customize.json', { app: appId })
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

  // 有料版のご案内（画面の下に小さく1行）
  function moreLink() {
    return el('p', { class: P + '-more' }, [
      'すべてのアプリをまとめて点検するなら：',
      el('a', { href: 'https://daichoya.github.io/pro.html', target: '_blank', rel: 'noopener', text: '棚卸し・変更監視パック（有料版・30日間の無料試用あり）' })
    ]);
  }

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
        var sheets = buildSheets(results, nowText());
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
