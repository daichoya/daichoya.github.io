/*
 * 台帳屋 棚卸し・変更監視パック for kintone
 * 一覧画面に「棚卸し」ボタンを追加し、すべてのアプリの設定を読み取って点検台帳（Excel）を作る。
 * 前回の点検台帳を選ぶと、変わったところも一覧にする。
 * 通信するのはkintone（cybozu.com）だけ。読み取るのはアプリの設定と、レコードの件数・最終更新日時だけ（レコードの中身は読まない）。
 */
(function (PLUGIN_ID) {
  'use strict';

  // 共通部品：設定に出てくるユーザー・組織・グループを集める（js/refs.js）
  var Refs = typeof module !== 'undefined' && module.exports ? require('../../../common/js/refs.js') : window.DaichoyaRefs;
  // 共通部品：フィールドが使われている場所を探す（js/fieldrefs.js）
  var FieldRefs = typeof module !== 'undefined' && module.exports ? require('../../../common/js/fieldrefs.js') : window.DaichoyaFieldRefs;

  // ===== 1. Excel（xlsx）の書き出し =====
  // 共通部品（js/xlsx.js）で作る。見出しの色・固定行・フィルター・印刷設定は、台帳屋のどの製品でも同じ体裁にそろえる。
  var Xlsx = typeof module !== 'undefined' && module.exports ? require('../../../common/js/xlsx.js') : window.DaichoyaXlsx;
  var STYLE = Xlsx.STYLE;
  var utf8 = Xlsx.utf8;

  function buildXlsx(sheets) {
    return Xlsx.build(sheets, { title: 'kintoneアプリ 点検台帳', creator: '台帳屋 棚卸し・変更監視パック', footer: '台帳屋 棚卸し・変更監視パック' });
  }

  // ===== 2. 前回の点検台帳の読み込み =====
  // 点検台帳の「記録」シート（非表示）に入れた設定の記録を取り出す。
  // Excelで開いて保存し直したファイル（圧縮あり・共有文字列あり）にも対応する。

  function readZip(bytes) {
    var view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
    var end = -1;
    for (var i = bytes.length - 22; i >= Math.max(0, bytes.length - 65557); i--) {
      if (view.getUint32(i, true) === 0x06054b50) { end = i; break; }
    }
    if (end < 0) throw new Error('zip');
    var count = view.getUint16(end + 10, true);
    var pos = view.getUint32(end + 16, true);
    var entries = {};
    for (var n = 0; n < count; n++) {
      if (view.getUint32(pos, true) !== 0x02014b50) throw new Error('zip');
      var method = view.getUint16(pos + 10, true);
      var size = view.getUint32(pos + 20, true);
      var nameLen = view.getUint16(pos + 28, true);
      var extraLen = view.getUint16(pos + 30, true);
      var commentLen = view.getUint16(pos + 32, true);
      var localOffset = view.getUint32(pos + 42, true);
      var name = new TextDecoder().decode(bytes.subarray(pos + 46, pos + 46 + nameLen));
      var localNameLen = view.getUint16(localOffset + 26, true);
      var localExtraLen = view.getUint16(localOffset + 28, true);
      var start = localOffset + 30 + localNameLen + localExtraLen;
      entries[name] = { method: method, data: bytes.subarray(start, start + size) };
      pos += 46 + nameLen + extraLen + commentLen;
    }
    return entries;
  }

  function inflate(entry) {
    if (entry.method === 0) return Promise.resolve(entry.data);
    if (entry.method !== 8 || typeof DecompressionStream === 'undefined') return Promise.reject(new Error('compress'));
    var stream = new Blob([entry.data]).stream().pipeThrough(new DecompressionStream('deflate-raw'));
    return new Response(stream).arrayBuffer().then(function (buf) { return new Uint8Array(buf); });
  }

  function readText(entries, name) {
    var entry = entries[name];
    if (!entry) return Promise.resolve(null);
    return inflate(entry).then(function (b) { return new TextDecoder().decode(b); });
  }

  function decodeXml(s) {
    return s.replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'")
      .replace(/&#(\d+);/g, function (_, d) { return String.fromCharCode(Number(d)); })
      .replace(/&#x([0-9a-fA-F]+);/g, function (_, h) { return String.fromCharCode(parseInt(h, 16)); })
      .replace(/&amp;/g, '&');
  }

  function textsOf(xml) {
    var out = [];
    var re = /<t(?:\s[^>]*)?>([\s\S]*?)<\/t>|<t(?:\s[^>]*)?\/>/g;
    var m;
    while ((m = re.exec(xml))) out.push(decodeXml(m[1] || ''));
    return out.join('');
  }

  var RECORD_SHEET = '記録';
  var RECORD_MARK = '台帳屋 点検記録 v1';

  // A列の文字を上から順につなげる
  function columnATexts(sheetXmlText, shared) {
    var out = [];
    var re = /<c\s([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g;
    var m;
    while ((m = re.exec(sheetXmlText))) {
      var attrs = m[1];
      var ref = /r="([A-Z]+)(\d+)"/.exec(attrs);
      if (!ref || ref[1] !== 'A') continue;
      var body = m[2] || '';
      var type = (/t="([^"]+)"/.exec(attrs) || [])[1];
      var value;
      if (type === 's') value = shared[Number((/<v>(\d+)<\/v>/.exec(body) || [])[1])] || '';
      else if (type === 'inlineStr') value = textsOf(body);
      else value = decodeXml((/<v>([\s\S]*?)<\/v>/.exec(body) || [])[1] || '');
      out.push(value);
    }
    return out;
  }

  function readSnapshot(bytes) {
    var entries;
    try { entries = readZip(bytes); } catch (e) { return Promise.reject(new Error('Excelファイルとして読み込めませんでした。')); }
    return Promise.all([readText(entries, 'xl/workbook.xml'), readText(entries, 'xl/_rels/workbook.xml.rels'), readText(entries, 'xl/sharedStrings.xml')])
      .then(function (r) {
        var workbook = r[0], rels = r[1], sst = r[2];
        if (!workbook || !rels) throw new Error('Excelファイルとして読み込めませんでした。');
        var sheet = new RegExp('<sheet\\s[^>]*name="' + RECORD_SHEET + '"[^>]*>').exec(workbook);
        if (!sheet) throw new Error('台帳屋の点検台帳ではないか、記録のシートが削除されています。');
        var rid = /r:id="([^"]+)"/.exec(sheet[0])[1];
        var target = new RegExp('<Relationship\\s[^>]*Id="' + rid + '"[^>]*>').exec(rels);
        var path = /Target="([^"]+)"/.exec(target[0])[1].replace(/^\/?(xl\/)?/, '');
        var shared = [];
        if (sst) {
          var re = /<si>([\s\S]*?)<\/si>/g;
          var m;
          while ((m = re.exec(sst))) shared.push(textsOf(m[1]));
        }
        return readText(entries, 'xl/' + path).then(function (xml) {
          var texts = columnATexts(xml || '', shared);
          if (texts[0] !== RECORD_MARK) throw new Error('台帳屋の点検台帳ではないか、記録のシートが書き換えられています。');
          return JSON.parse(texts.slice(1).join(''));
        });
      });
  }

  // ===== 3. 設定を「記録」の形にそろえる（純粋な関数。テスト対象） =====

  var ENTITY_LABELS = { USER: 'ユーザー', GROUP: 'グループ', ORGANIZATION: '組織', CREATOR: '作成者', FIELD_ENTITY: 'フィールド', CUSTOM_FIELD: 'フィールド' };
  var ACCESS_LABELS = { WRITE: '閲覧・編集', READ: '閲覧のみ', NONE: '権限なし' };
  var APP_RIGHTS = [['appEditable', 'アプリ管理'], ['recordViewable', '閲覧'], ['recordAddable', '追加'], ['recordEditable', '編集'],
    ['recordDeletable', '削除'], ['recordImportable', '読み込み'], ['recordExportable', '書き出し']];
  var SCOPE_TEXT = { ALL: 'すべてのユーザー', ADMIN: 'アプリの管理者だけ', NONE: '適用しない' };

  function entityText(entity, includeSubs) {
    if (!entity) return '';
    if (entity.type === 'CREATOR') return '作成者';
    if (entity.type === 'GROUP' && entity.code === 'everyone') return 'Everyone';
    return (ENTITY_LABELS[entity.type] || entity.type) + '：' + (entity.code || '') + (includeSubs ? '（下位組織を含む）' : '');
  }

  function isEveryone(entity) { return entity && entity.type === 'GROUP' && entity.code === 'everyone'; }

  function appAclRows(acl) {
    return ((acl && acl.rights) || []).map(function (r, i) {
      var granted = APP_RIGHTS.filter(function (x) { return r[x[0]]; }).map(function (x) { return x[1]; });
      return { order: i + 1, entity: r.entity, target: entityText(r.entity, r.includeSubs), text: granted.length ? granted.join('・') : '権限なし', raw: r };
    });
  }

  function recordAclRows(acl) {
    var rows = [];
    ((acl && acl.rights) || []).forEach(function (right, i) {
      (right.entities || []).forEach(function (e) {
        var granted = [['viewable', '閲覧'], ['editable', '編集'], ['deletable', '削除']].filter(function (x) { return e[x[0]]; }).map(function (x) { return x[1]; });
        rows.push({ order: i + 1, condition: right.filterCond || '（すべてのレコード）', entity: e.entity, target: entityText(e.entity, e.includeSubs), text: granted.length ? granted.join('・') : '権限なし' });
      });
    });
    return rows;
  }

  function fieldAclRows(acl, labels) {
    var rows = [];
    ((acl && acl.rights) || []).forEach(function (right) {
      (right.entities || []).forEach(function (e, i) {
        rows.push({ order: i + 1, field: (labels[right.code] || right.code) + '（' + right.code + '）', entity: e.entity, target: entityText(e.entity, e.includeSubs), text: ACCESS_LABELS[e.accessibility] || e.accessibility });
      });
    });
    return rows;
  }

  function labelMap(properties) {
    var map = {};
    Object.keys(properties || {}).forEach(function (code) {
      var p = properties[code];
      map[code] = p.label || code;
      if (p.type === 'SUBTABLE' && p.fields) Object.keys(p.fields).forEach(function (inner) { map[inner] = p.fields[inner].label || inner; });
    });
    return map;
  }

  // フィールドの設定を、変わったかどうかを比べられる短い文字列にする
  function fieldSignature(p) {
    var sig = { t: p.type, l: p.label };
    if (p.required) sig.r = 1;
    if (p.unique) sig.u = 1;
    if (p.options) sig.o = Object.keys(p.options).sort(function (a, b) { return Number(p.options[a].index) - Number(p.options[b].index); });
    if (p.expression) sig.x = p.expression;
    if (p.defaultValue !== undefined && p.defaultValue !== '' && !(Array.isArray(p.defaultValue) && !p.defaultValue.length)) sig.d = p.defaultValue;
    if (p.lookup) {
      sig.k = { a: p.lookup.relatedApp && p.lookup.relatedApp.app, f: p.lookup.relatedKeyField, m: (p.lookup.fieldMappings || []).map(function (m) { return m.relatedField + '>' + m.field; }) };
    }
    if (p.referenceTable) {
      var rt = p.referenceTable;
      sig.rt = { a: rt.relatedApp && rt.relatedApp.app, c: rt.condition && (rt.condition.field + '=' + rt.condition.relatedField), q: rt.filterCond || '' };
    }
    return JSON.stringify(sig);
  }

  function flattenFields(properties) {
    var out = {};
    Object.keys(properties || {}).forEach(function (code) {
      var p = properties[code];
      if (p.type === 'SUBTABLE') {
        out[code] = fieldSignature({ type: p.type, label: p.label });
        Object.keys(p.fields || {}).forEach(function (inner) { out[inner] = fieldSignature(p.fields[inner]); });
      } else {
        out[code] = fieldSignature(p);
      }
    });
    return out;
  }

  function relatedId(relatedApp) { return relatedApp && relatedApp.app ? String(relatedApp.app) : null; }

  function extractLinks(appId, properties, actions) {
    var links = [];
    function visit(props, tableLabel) {
      Object.keys(props || {}).forEach(function (code) {
        var p = props[code];
        if (p.type === 'SUBTABLE') { visit(p.fields, p.label || code); return; }
        var place = (p.label || code) + (tableLabel ? '（テーブル「' + tableLabel + '」内）' : '');
        if (p.lookup && relatedId(p.lookup.relatedApp)) links.push({ from: String(appId), to: relatedId(p.lookup.relatedApp), kind: 'ルックアップ', place: place });
        if (p.referenceTable && relatedId(p.referenceTable.relatedApp)) links.push({ from: String(appId), to: relatedId(p.referenceTable.relatedApp), kind: '関連レコード一覧', place: place });
      });
    }
    visit(properties, null);
    Object.keys((actions && actions.actions) || {}).forEach(function (name) {
      var a = actions.actions[name];
      if (!a) return;
      // 作成先のアプリを見る権限がないと、kintoneは destApp を null で返す。つながりは残して「見られないアプリ」と書く
      links.push({ from: String(appId), to: a.destApp && a.destApp.app ? String(a.destApp.app) : '', kind: 'アクション', place: a.name || name });
    });
    return links;
  }

  function customizeFiles(customize) {
    var rows = [];
    var c = customize || {};
    [['desktop', 'PC'], ['mobile', 'モバイル']].forEach(function (screen) {
      [['js', 'JavaScript'], ['css', 'CSS']].forEach(function (kind) {
        (((c[screen[0]] || {})[kind[0]]) || []).forEach(function (item, i) {
          var isUrl = item.type === 'URL';
          rows.push({ screen: screen[1], kind: kind[1], order: i + 1, form: isUrl ? 'URL' : 'ファイル', name: isUrl ? (item.url || '') : ((item.file && item.file.name) || '') });
        });
      });
    });
    return rows;
  }

  // フィールドの使われ方（このアプリの中）。レコードの中身は見ず、設定だけで判断する
  var NOT_DATA_FIELDS = { SUBTABLE: 1, REFERENCE_TABLE: 1 };
  var SYSTEM_FIELDS = { RECORD_NUMBER: 1, CREATOR: 1, CREATED_TIME: 1, MODIFIER: 1, UPDATED_TIME: 1 };

  function fieldUse(app, val) {
    var fields = val('fields');
    if (!fields) return null;
    var settings = {
      fields: fields, views: val('views'), reports: val('reports'), recordAcl: val('recordAcl'), fieldAcl: val('fieldAcl'),
      status: val('status'), general: val('general'), perRecord: val('perRecord'), reminder: val('reminder'),
      actions: val('actions'), appSettings: val('appSettings')
    };
    var flat = FieldRefs.flatFields(fields.properties);
    return FieldRefs.choices(fields.properties).filter(function (c) { return !NOT_DATA_FIELDS[c.type]; }).map(function (c) {
      var places = [];
      var own = flat[c.code] && flat[c.code].prop;
      // ルックアップで、ほかのフィールドに値をコピーしているなら、それ自体が働いている設定
      if (own && own.lookup && (own.lookup.fieldMappings || []).length) places.push('ルックアップ（ほかのフィールドへのコピー）');
      FieldRefs.find({ appId: app.appId, appName: app.name, code: c.code }, settings, [], []).forEach(function (h) {
        if (h.place === 'ルックアップ') return; // そのフィールド自身がルックアップであること（使われている場所ではない）
        if (places.indexOf(h.place) < 0) places.push(h.place);
      });
      return { c: c.code, l: c.label, t: c.type, u: places };
    });
  }

  // ほかのアプリのフィールドを参照している設定（ルックアップ・関連レコード一覧・アクション）だけを、短い形で残す
  function crossRefs(properties, actions) {
    var props = {};
    var labels = labelMap(properties);
    function stub(code) { if (code && !props[code]) props[code] = { label: labels[code] || code }; }
    function keep(code, p) {
      var lk = p.lookup;
      var rt = p.referenceTable;
      if (lk && relatedId(lk.relatedApp)) {
        props[code] = { label: p.label || code, type: p.type, lookup: {
          relatedApp: { app: relatedId(lk.relatedApp) }, relatedKeyField: lk.relatedKeyField,
          fieldMappings: (lk.fieldMappings || []).map(function (m) { return { field: m.field, relatedField: m.relatedField }; }),
          lookupPickerFields: lk.lookupPickerFields || [], filterCond: lk.filterCond || '', sort: lk.sort || '' } };
        (lk.fieldMappings || []).forEach(function (m) { stub(m.field); });
      } else if (rt && relatedId(rt.relatedApp)) {
        props[code] = { label: p.label || code, type: p.type, referenceTable: {
          relatedApp: { app: relatedId(rt.relatedApp) },
          condition: rt.condition ? { field: rt.condition.field, relatedField: rt.condition.relatedField } : null,
          displayFields: rt.displayFields || [], filterCond: rt.filterCond || '', sort: rt.sort || '' } };
      }
    }
    Object.keys(properties || {}).forEach(function (code) {
      var p = properties[code];
      if (p.type === 'SUBTABLE') Object.keys(p.fields || {}).forEach(function (inner) { keep(inner, p.fields[inner]); });
      else keep(code, p);
    });
    var acts = {};
    Object.keys((actions && actions.actions) || {}).forEach(function (name) {
      var a = actions.actions[name];
      if (!a || !a.destApp || !a.destApp.app) return;
      acts[name] = { name: a.name || name, index: a.index, destApp: { app: String(a.destApp.app) },
        mappings: (a.mappings || []).filter(function (m) { return m.destField; }).map(function (m) { return { destField: m.destField }; }) };
    });
    return { properties: props, actions: acts };
  }

  // results: アプリごとの API の結果 { ok, value } / { ok: false, error }
  function appSnapshot(app, results) {
    function val(k) { return results[k] && results[k].ok ? results[k].value : null; }
    var errors = {};
    Object.keys(results).forEach(function (k) { if (!results[k].ok) errors[k] = errorText(results[k].error); });
    var properties = (val('fields') && val('fields').properties) || null;
    var labels = labelMap(properties || {});
    function rowsText(rows, withCondition) {
      return rows.map(function (r) { return r.order + '｜' + (withCondition ? (r.condition || r.field) + '｜' : '') + r.target + '｜' + r.text; });
    }
    return {
      id: String(app.appId), name: app.name, spaceId: app.spaceId || '',
      creator: app.creator ? (app.creator.name || app.creator.code) : '', createdAt: app.createdAt || '',
      modifier: app.modifier ? (app.modifier.name || app.modifier.code) : '', modifiedAt: app.modifiedAt || '',
      fields: properties ? flattenFields(properties) : null,
      labels: properties ? labels : {},
      acl: {
        app: val('appAcl') ? rowsText(appAclRows(val('appAcl')), false) : null,
        record: val('recordAcl') ? rowsText(recordAclRows(val('recordAcl')), true) : null,
        field: val('fieldAcl') ? rowsText(fieldAclRows(val('fieldAcl'), labels), true) : null
      },
      plugins: val('plugins') ? (val('plugins').plugins || []).map(function (p) { return { id: p.id, name: p.name, enabled: p.enabled !== false }; }) : null,
      customize: val('customize') ? { scope: val('customize').scope, files: customizeFiles(val('customize')) } : null,
      links: properties ? extractLinks(app.appId, properties, val('actions')) : [],
      records: recordActivity(val('records')),
      refs: Refs.collect({ appAcl: val('appAcl'), recordAcl: val('recordAcl'), fieldAcl: val('fieldAcl'), status: val('status'), general: val('general'), perRecord: val('perRecord'), reminder: val('reminder'), fields: val('fields'), views: val('views'), reports: val('reports'), actions: val('actions') }),
      soleManager: Refs.soleManager(val('appAcl')),
      fieldUse: fieldUse(app, val),
      xref: properties ? crossRefs(properties, val('actions')) : null,
      warnings: warningsFor(val('appAcl'), val('recordAcl'), val('fieldAcl'), labels),
      notes: val('adminNotes') ? !!String(val('adminNotes').content || '').replace(/<[^>]*>/g, '').replace(/&nbsp;/g, ' ').trim() : null,
      health: healthChecks(app, properties, val('customize'), val('adminNotes') ? !!String(val('adminNotes').content || '').replace(/<[^>]*>/g, '').replace(/&nbsp;/g, ' ').trim() : null),
      errors: errors
    };
  }

  // レコードの件数と、最後に更新された日時（レコードの中身は記録しない）
  function recordActivity(value) {
    if (!value) return null;
    var last = '';
    var rec = (value.records || [])[0] || {};
    Object.keys(rec).forEach(function (k) { if (rec[k] && rec[k].type === 'UPDATED_TIME') last = rec[k].value; });
    return { count: Number(value.totalCount || 0), lastUpdated: last };
  }

  var DORMANT_DAYS = 180;

  // referencedBy: ほかのアプリからルックアップや関連レコード一覧で参照されているアプリ（マスタなど）
  function usageWarnings(a, now, referencedBy) {
    var r = a.records;
    if (!r) return [];
    var refs = (referencedBy && referencedBy[a.id]) || [];
    var note = refs.length ? '（ただし、ほかの' + refs.length + '個のアプリから参照されています。マスタとして使われている場合は問題ありません）' : '';
    var item = function (text) { return [{ level: '確認', kind: '利用状況', order: '', target: '', text: text + note }]; };
    if (r.count === 0) return item('レコードが1件もありません。使われていないアプリかもしれません。');
    var days = Number(daysSince(r.lastUpdated, now));
    if (r.lastUpdated && days >= DORMANT_DAYS) return item('レコードが' + days + '日間更新されていません。使われていないアプリかもしれません。');
    return [];
  }

  // dir：User API で確かめた結果（Refs.lookup）。ないときは状態を判定しない
  function refWarnings(a, dir) {
    var out = [];
    var byCode = {};
    Refs.rows(a.refs || [], dir).forEach(function (r) {
      if (!r.level) return;
      var key = r.type + '\n' + r.code;
      if (!byCode[key]) { byCode[key] = { r: r, places: [] }; }
      if (byCode[key].places.indexOf(r.place) < 0) byCode[key].places.push(r.place);
    });
    Object.keys(byCode).forEach(function (key) {
      var x = byCode[key];
      var target = x.r.typeLabel + '：' + x.r.code;
      if (x.r.level === 2) out.push({ level: '要確認', kind: 'ユーザーの指定', order: '', target: target, text: 'もういない' + x.r.typeLabel + 'が指定されています（' + x.places.join('、') + '）。' });
      else out.push({ level: '確認', kind: 'ユーザーの指定', order: '', target: target, text: '利用停止中のユーザーが指定されています（' + x.places.join('、') + '）。' });
    });
    if (a.soleManager) out.push({ level: '確認', kind: 'アプリ', order: '', target: 'ユーザー：' + a.soleManager, text: 'アプリ管理の権限が、このユーザー1人だけに付いています。異動や退職の前に、ほかの人やグループにも付けておくと安心です。' });
    return out;
  }

  function referencedApps(apps) {
    var out = {};
    apps.forEach(function (a) {
      (a.links || []).forEach(function (l) {
        if (l.kind === 'アクション' || l.to === a.id) return;
        out[l.to] = out[l.to] || [];
        if (out[l.to].indexOf(a.id) < 0) out[l.to].push(a.id);
      });
    });
    return out;
  }

  function errorText(error) {
    var code = error && error.code;
    if (code === 'CB_NO02' || code === 'GAIA_NO01') return '権限がありません';
    return (error && error.message) || '原因不明';
  }

  // 権限の注意点
  function warningsFor(appAcl, recordAcl, fieldAcl, labels) {
    var out = [];
    appAclRows(appAcl).forEach(function (r) {
      if (isEveryone(r.entity) && r.raw.appEditable) out.push({ level: '要確認', kind: 'アプリ', order: r.order, target: r.target, text: 'Everyoneに「アプリ管理」の権限があります。だれでもアプリの設定を変えられます。' });
      if (isEveryone(r.entity) && r.raw.recordDeletable) out.push({ level: '要確認', kind: 'アプリ', order: r.order, target: r.target, text: 'Everyoneに「削除」の権限があります。' });
      if (isEveryone(r.entity) && r.raw.recordExportable) out.push({ level: '確認', kind: 'アプリ', order: r.order, target: r.target, text: 'Everyoneに「書き出し」の権限があります。だれでもレコードをファイルに書き出せます。' });
      if (r.entity && r.entity.type === 'USER') out.push({ level: '確認', kind: 'アプリ', order: r.order, target: r.target, text: 'ユーザー個人を指定しています。異動や退職のときに見直しが必要です。' });
    });
    recordAclRows(recordAcl).forEach(function (r) {
      if (r.entity && r.entity.type === 'USER') out.push({ level: '確認', kind: 'レコード', order: r.order, target: r.target, text: 'ユーザー個人を指定しています（条件：' + r.condition + '）。' });
    });
    fieldAclRows(fieldAcl, labels || {}).forEach(function (r) {
      if (r.entity && r.entity.type === 'USER') out.push({ level: '確認', kind: 'フィールド', order: r.order, target: r.target, text: 'ユーザー個人を指定しています（' + r.field + '）。' });
    });
    return out;
  }

  // ===== 3-2. 健康診断（サイボウズの「kintoneガバナンスガイドライン」の点検項目と、保守のしやすさ） =====

  // kintoneが最初に付けるフィールドコード（日本語・英語の画面）。末尾に _0 などが付くこともある
  var DEFAULT_CODE = /^(文字列__1行_|文字列__複数行_|リッチエディター|数値|計算|ラジオボタン|チェックボックス|複数選択|ドロップダウン|日付|時刻|日時|添付ファイル|リンク|ユーザー選択|組織選択|グループ選択|テーブル|関連レコード一覧|Text|Text_area|Rich_text|Number|Calculated|Radio_button|Check_box|Multi_choice|Drop_down|Date|Time|Date_and_time|Attachment|Link|User_selection|Department_selection|Group_selection|Table|Related_records)(_\d+)?$/;
  var KINTONE_HOST = /^https:\/\/[^\/]+\.(cybozu\.com|kintone\.com|cybozu\.cn)\//;

  function defaultCodes(properties) {
    var out = [];
    Object.keys(properties || {}).forEach(function (code) {
      var p = properties[code];
      if (DEFAULT_CODE.test(code)) out.push(code);
      if (p.type === 'SUBTABLE') Object.keys(p.fields || {}).forEach(function (inner) { if (DEFAULT_CODE.test(inner)) out.push(inner); });
    });
    return out;
  }

  function hostOf(url) {
    var m = /^https?:\/\/([^\/?#]+)/.exec(String(url || ''));
    return m ? m[1] : String(url || '');
  }

  // app：アプリ一覧の1件（spaceId・description）。properties・customize：読み取れなかったときは null
  // notes：アプリ管理者用メモがあるか（true / false。読み取れなかったときは null）
  function healthChecks(app, properties, customize, notes) {
    var out = [];
    var item = function (level, target, text) { out.push({ level: level, kind: '健康診断', order: '', target: target, text: text }); };
    if (!app.spaceId) item('確認', 'スペース', 'スペースに入っていないアプリです。kintoneガバナンスガイドラインでは、スペースに所属しないアプリを作らない運用が例に挙げられています。');
    var noDescription = app.description !== undefined && !String(app.description || '').replace(/<[^>]*>/g, '').trim();
    if (noDescription && notes === false) item('確認', 'アプリの説明', 'アプリの説明も、アプリ管理者用メモもありません。引き継ぎのときに何のアプリか分かるよう、目的や担当を書いておくと安心です（サイボウズの運用ルールでは、機能を足した背景や要件をアプリ管理者用メモに残すことがすすめられています）。');
    else if (noDescription && notes !== true) item('確認', 'アプリの説明', 'アプリの説明がありません。引き継ぎのときに何のアプリか分かるよう、目的や担当を書いておくと安心です。');
    if (customize) {
      var outside = {};
      customizeFiles(customize).forEach(function (f) { if (f.form === 'URL' && !KINTONE_HOST.test(f.name)) outside[hostOf(f.name)] = true; });
      var hosts = Object.keys(outside);
      if (hosts.length) item('要確認', 'JavaScript・CSS', '外部のURL（' + hosts.join('、') + '）からJavaScript・CSSを読み込んでいます。読み込み先が変わると、アプリの動きも変わります。読み込み先を管理しているか確かめてください。');
    }
    if (properties) {
      var codes = defaultCodes(properties);
      if (codes.length) item('確認', 'フィールドコード', '初期のままのフィールドコードが' + codes.length + '個あります（例：' + codes.slice(0, 3).join('、') + '）。計算式やJavaScriptで使うときに分かりにくいため、意味の分かるコードにしておくと保守が楽になります。');
    }
    return out;
  }

  // 環境全体の点検。env：{ administrator: true（使える）/ false（停止中）/ null（確かめられなかった） }
  function envWarnings(env) {
    var out = [];
    if (env && env.administrator === true) out.push({ level: '確認', kind: '健康診断', order: '', target: 'ユーザー：Administrator', text: '最初からある管理者のアカウント「Administrator」が使える状態です。kintoneガバナンスガイドラインでは、業務で使わないなら停止することがすすめられています。' });
    return out;
  }

  // kintoneに読み込んであるのに、点検したどのアプリにも追加されていないプラグイン。
  // installed：GET /k/v1/plugins.json の一覧（読めなかったときは null）。selfId：このプラグイン自身（数えない）
  function unusedPlugins(apps, installed, selfId) {
    if (!installed || !installed.length) return [];
    var used = {};
    apps.forEach(function (a) { (a.plugins || []).forEach(function (p) { used[p.id] = true; }); });
    var unread = apps.filter(function (a) { return !a.plugins; }).length;
    return installed.filter(function (p) { return p.id !== selfId && !used[p.id]; }).map(function (p) {
      return { level: '確認', kind: '健康診断', order: '', target: 'プラグイン：' + p.name,
        text: 'kintoneに読み込んであるのに、点検したどのアプリにも追加されていないプラグインです。使っていなければ、kintoneシステム管理の「プラグイン」から削除できます（サイボウズの運用ルールでは、使われていないプラグインを棚卸しで確かめることがすすめられています）。' +
          (unread ? 'プラグインを読み取れなかったアプリが' + unread + '個あるため、そこで使われている場合もあります。' : '') };
    });
  }

  // 名前が同じか、よく似たアプリ（「（旧）」「コピー」「テスト」「検証」などを除くと同じ名前）。
  // 新旧のアプリが並んでいると、利用者が古いほうを使ってしまうことがある
  var NAME_WORDS = '(旧版|旧|コピー|copy|テスト|test|検証|old|bak|backup|削除予定|不要|使用禁止)\\d*';
  var NAME_NOISE = new RegExp('[\\s　]|[（(【\\[]' + NAME_WORDS + '[）)】\\]]|[_\\-－ー・]?' + NAME_WORDS + '$|のコピー\\d*$', 'gi');
  function similarApps(apps) {
    var groups = {};
    apps.forEach(function (a) {
      var key = String(a.name || '').replace(NAME_NOISE, '').toLowerCase();
      if (key) (groups[key] = groups[key] || []).push(a);
    });
    return Object.keys(groups).filter(function (k) { return groups[k].length > 1; }).map(function (k) {
      var list = groups[k];
      return { app: list[0], level: '確認', kind: '健康診断', order: '', target: 'アプリ名',
        text: '名前が同じか、よく似たアプリがあります（' + list.map(function (a) { return '「' + a.name + '」（ID ' + a.id + '）'; }).join('、') + '）。新旧のアプリが並んでいると、利用者が古いほうを使ってしまうことがあります。使っていないほうは、閲覧の権限を外すか、削除を検討してください。' };
    });
  }

  // 許可するプラグイン（設定画面で選んだもの）にないプラグインが入っていないか。
  // allowed：プラグインIDの配列。指定がない（null・空）ときは点検しない。selfId：このプラグイン自身（いつも許可として扱う）
  function pluginWarnings(a, allowed, selfId) {
    if (!allowed || !allowed.length || !a.plugins) return [];
    var ok = {};
    allowed.forEach(function (id) { ok[id] = true; });
    if (selfId) ok[selfId] = true;
    return a.plugins.filter(function (p) { return !ok[p.id]; }).map(function (p) {
      return { level: '要確認', kind: '健康診断', order: '', target: 'プラグイン：' + p.name,
        text: '許可するプラグインに入っていないプラグインです' + (p.enabled ? '' : '（いまは無効）') + '。誰が何のために入れたかを確かめ、使い続けるなら、棚卸し・変更監視パックの設定画面で許可に加えてください。' };
    });
  }

  function isAllowedPlugin(id, allowed, selfId) {
    if (!allowed || !allowed.length) return null;
    return id === selfId || allowed.indexOf(id) >= 0;
  }

  // ===== 3-3. 点検の推移と報告書 =====
  // 点検台帳の記録に、これまでの点検の件数を残す（前回の点検台帳を選んで作るたびに1行ずつ増える）

  var HISTORY_MAX = 24;

  function inspectionSummary(at, appCount, warnings, diff) {
    var groups = {};
    FREE_GROUPS.forEach(function (g) { groups[g.name] = warnings.filter(function (w) { return g.kinds.indexOf(w[3]) >= 0; }).length; });
    return { at: at, apps: appCount, total: warnings.length, important: warnings.filter(function (w) { return w[0] === '要確認'; }).length, groups: groups, changes: diff ? diff.length : null };
  }

  function historyOf(prev, host, current) {
    var past = [];
    if (prev && (!prev.host || !host || prev.host === host)) {
      past = Array.isArray(prev.history) && prev.history.length ? prev.history.slice()
        // 1.2以前の点検台帳には件数の記録がないため、点検日とアプリの数だけ残す
        : [{ at: prev.takenAt, apps: Object.keys(prev.apps || {}).length, total: null, important: null, groups: null, changes: null }];
    }
    return past.concat([current]).slice(-HISTORY_MAX);
  }

  function known(n) { return n !== null && n !== undefined; }

  function historyText(h) {
    return ['アプリ ' + h.apps + '個', known(h.total) ? '注意点 ' + h.total + '件（うち要確認 ' + h.important + '件）' : '注意点の記録なし',
      known(h.changes) ? '前回からの変更 ' + h.changes + '件' : ''].filter(Boolean).join('、');
  }

  var DIFF_ITEMS = [['アプリ', ['アプリ', 'アプリ名']], ['フィールド', ['フィールド']], ['アクセス権', ['アプリのアクセス権', 'レコードのアクセス権', 'フィールドのアクセス権']], ['プラグイン', ['プラグイン']], ['カスタマイズ', ['カスタマイズ']]];
  var REPORT_IMPORTANT_MAX = 8;

  // 報告書に載せる内容。sheets は buildSheets の結果（注意点は sheets[3]、変更点は sheets[2]）
  function reportData(snapshot, prev, sheets, now, user) {
    var warnings = sheets[3].blocks[0].rows;
    var history = snapshot.history || [];
    var current = history[history.length - 1] || inspectionSummary(snapshot.takenAt, Object.keys(snapshot.apps).length, warnings, null);
    var before = history.length > 1 ? history[history.length - 2] : null;
    var groups = FREE_GROUPS.map(function (g) {
      var hit = warnings.filter(function (w) { return g.kinds.indexOf(w[3]) >= 0; });
      return { name: g.name, count: hit.length, important: hit.filter(function (w) { return w[0] === '要確認'; }).length,
        before: before && before.groups && known(before.groups[g.name]) ? before.groups[g.name] : null };
    });
    var diff = null;
    if (prev) {
      var rows = sheets[2].blocks[0].rows;
      diff = DIFF_ITEMS.map(function (item) {
        var hit = rows.filter(function (r) { return item[1].indexOf(r[3]) >= 0; });
        var count = function (kind) { return hit.filter(function (r) { return r[0] === kind; }).length; };
        return { name: item[0], add: count('追加'), change: count('変更'), remove: count('削除') };
      });
    }
    var important = warnings.filter(function (w) { return w[0] === '要確認'; });
    var apps = Object.keys(snapshot.apps).map(function (id) { return snapshot.apps[id]; });
    return {
      at: now, host: snapshot.host, user: user || '', apps: apps.length,
      unreadable: apps.filter(function (a) { return Object.keys(a.errors || {}).length; }).length,
      total: current.total, important: current.important, before: before,
      groups: groups, diff: diff, prevAt: prev ? prev.takenAt : null,
      history: history.slice(-6),
      items: important.slice(0, REPORT_IMPORTANT_MAX).map(function (w) { return { app: w[2], target: w[5], text: w[6] }; }),
      more: Math.max(0, important.length - REPORT_IMPORTANT_MAX),
      plan: snapshot.plan || ''
    };
  }

  function esc(value) {
    return String(known(value) ? value : '').replace(/[&<>"']/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]; });
  }

  function signed(n) { return n > 0 ? '+' + n : n < 0 ? '−' + Math.abs(n) : '±0'; }

  function firstSentence(text) { var m = /^[^。]*。/.exec(String(text || '')); return m ? m[0] : String(text || ''); }

  function clip(text, max) { text = String(text || ''); return text.length > max ? text.slice(0, max - 1) + '…' : text; }

  var REPORT_CSS = '@page{size:A4 portrait;margin:12mm}' +
    'body{margin:0;color:#222;font-family:"BIZ UDPGothic","Yu Gothic","Hiragino Sans","Meiryo",sans-serif;font-size:9.5pt;line-height:1.45;background:#fff}' +
    '.bar{position:sticky;top:0;display:flex;flex-wrap:wrap;gap:12px;align-items:center;padding:10px 16px;background:#f5f6f7;border-bottom:1px solid #ccc}' +
    '.bar button{height:40px;padding:0 18px;border:none;background:#3498db;color:#fff;font:inherit;font-size:11pt;cursor:pointer}' +
    '.sheet{max-width:186mm;margin:0 auto;padding:8mm 4mm 12mm}' +
    'h1{margin:0;padding-top:5mm;font-size:16pt;text-align:center;letter-spacing:.15em;color:#1f3a5f}' +
    '.sheet{position:relative}.date{position:absolute;top:0;right:0;margin:0;font-size:9pt}' +
    '.parties{display:flex;justify-content:space-between;gap:8mm;margin:3mm 0 2mm}.parties div{flex:1}' +
    '.fill{display:inline-block;min-width:52mm;border-bottom:1px solid #666;padding:0 2px}.fill:empty::before{content:attr(data-placeholder);color:#aaa}' +
    '.summary{margin:1.5mm 0 1mm;padding:2mm 3mm;border:1px solid #1f3a5f;border-left:5px solid #1f3a5f;background:#f4f7fb;font-size:10pt}' +
    '.summary b{color:#1f3a5f}.scope{margin:1mm 0 0;font-size:8.5pt;color:#444}' +
    'h2{margin:3.5mm 0 1.5mm;padding-left:6px;border-left:4px solid #1f3a5f;font-size:10.5pt;color:#1f3a5f}' +
    'table{width:100%;border-collapse:collapse}th,td{padding:1.5px 5px;border:1px solid #999;text-align:left;vertical-align:top}tr{break-inside:avoid}' +
    'th{background:#1f3a5f;color:#fff;border-color:#1f3a5f}.n{text-align:right;white-space:nowrap}.total td{font-weight:bold;background:#eef2f6}.imp{color:#b42318;font-weight:bold}' +
    '.meta td{padding:1px 0;border:none}.note{margin:1.5mm 0 0;font-size:8pt;color:#555}' +
    '.two{display:flex;gap:5mm}.two>section{flex:1;min-width:0}.items td{font-size:8pt}' +
    '.end{display:flex;gap:4mm;align-items:stretch}.memo{flex:1;min-height:15mm;padding:4px 6px;border:1px solid #999;white-space:pre-wrap}.memo:empty::before{content:attr(data-placeholder);color:#999}' +
    '.sign{display:flex;flex:none}.sign div{width:22mm;margin-left:-1px;border:1px solid #999;text-align:center}' +
    '.sign span{display:block;border-bottom:1px solid #999;background:#eef1f3;font-size:8.5pt}.sign i{display:block;height:14mm}' +
    '.foot{margin-top:3mm;font-size:8pt;color:#666}' +
    '@media screen and (max-width:700px){.two{display:block}}' +
    '@media print{.bar{display:none}.sheet{max-width:none;padding:0}.memo:empty::before{content:""}.fill:empty::before{content:""}.sheet{padding:0 1px}}';

  // 報告書の冒頭に置く総評（数字は点検の結果だけから作る）
  function summarySentence(d) {
    var out = '<b>総評：</b>点検したアプリ' + d.apps + '個で、見直したほうがよい設定（注意点）が<b>' + d.total + '件</b>見つかりました（うち要確認 <span class="imp">' + d.important + '件</span>）。';
    if (d.before && known(d.before.total)) {
      var diff = d.total - d.before.total;
      out += '前回の点検（' + esc(dateText(d.before.at).slice(0, 10)) + '）と比べて、' + (diff < 0 ? Math.abs(diff) + '件減りました。' : diff > 0 ? diff + '件増えました。' : '件数は変わっていません。');
    }
    if (d.diff) {
      var changes = d.diff.reduce(function (n, r) { return n + r.add + r.change + r.remove; }, 0);
      out += '前回からの設定の変更は' + changes + '件です。';
    }
    out += d.important ? '「4. 要確認の項目」から、優先して確かめてください。' : '要確認の項目はありませんでした。';
    return out;
  }

  function reportHtml(d) {
    var day = dateText(d.at).slice(0, 10);
    var h = [];
    h.push('<!doctype html><html lang="ja"><head><meta charset="utf-8"><title>kintoneアプリ点検報告書（' + esc(day) + '）</title><style>' + REPORT_CSS + '</style></head><body>');
    h.push('<div class="bar"><button type="button" id="dcy-print">印刷・PDFで保存</button><span class="note">「提出先」「報告者」「所見・対応」の欄は、この画面で書き込んでから印刷できます。印刷の画面で「PDFに保存」を選ぶとPDFになります。</span></div>');
    h.push('<div class="sheet"><p class="date">' + esc(day.replace(/^(\d+)-(\d+)-(\d+)$/, function (_, y, m, dd) { return y + '年' + Number(m) + '月' + Number(dd) + '日'; })) + '</p><h1>kintoneアプリ 点検報告書</h1>');
    h.push('<div class="parties"><div>提出先：<span class="fill" contenteditable="true" data-placeholder="（例）情報システム部長 様"></span></div>' +
      '<div style="text-align:right">報告者：<span class="fill" contenteditable="true" data-placeholder="部署・氏名">' + esc(d.user) + '</span></div></div>');
    h.push('<div class="summary">' + summarySentence(d) + '</div>');
    h.push('<p class="scope">点検の範囲：kintone環境「' + esc(d.host) + '」で、点検した人が見られるアプリ' + d.apps + '個。点検日時：' + esc(dateText(d.at)) +
      '。方法：台帳屋 棚卸し・変更監視パックで、アプリの設定を読み取って点検しました（設定は変えていません。レコードの中身は読んでいません）。</p>');

    h.push('<h2>1. 点検の結果</h2><table><tr><th>区分</th><th class="n">件数</th><th class="n">うち要確認</th><th class="n">前回</th><th class="n">増減</th></tr>');
    d.groups.forEach(function (g) {
      h.push('<tr><td>' + esc(g.name) + '</td><td class="n">' + g.count + '件</td><td class="n">' + g.important + '件</td><td class="n">' +
        (known(g.before) ? g.before + '件' : '－') + '</td><td class="n">' + (known(g.before) ? signed(g.count - g.before) : '－') + '</td></tr>');
    });
    var bt = d.before && known(d.before.total) ? d.before.total : null;
    h.push('<tr class="total"><td>合計</td><td class="n">' + d.total + '件</td><td class="n">' + d.important + '件</td><td class="n">' + (known(bt) ? bt + '件' : '－') + '</td><td class="n">' + (known(bt) ? signed(d.total - bt) : '－') + '</td></tr></table>');
    if (d.unreadable) h.push('<p class="note">権限がなく、設定の一部を読み取れなかったアプリが' + d.unreadable + '個あります。kintoneシステム管理者が点検すると、すべて読み取れます。</p>');

    h.push('<div class="two"><section><h2>2. 前回の点検からの変更</h2>');
    if (d.diff) {
      h.push('<table><tr><th>項目</th><th class="n">追加</th><th class="n">変更</th><th class="n">削除</th></tr>');
      var sum = { add: 0, change: 0, remove: 0 };
      d.diff.forEach(function (r) {
        sum.add += r.add; sum.change += r.change; sum.remove += r.remove;
        h.push('<tr><td>' + esc(r.name) + '</td><td class="n">' + r.add + '件</td><td class="n">' + r.change + '件</td><td class="n">' + r.remove + '件</td></tr>');
      });
      h.push('<tr class="total"><td>合計</td><td class="n">' + sum.add + '件</td><td class="n">' + sum.change + '件</td><td class="n">' + sum.remove + '件</td></tr></table>');
      h.push('<p class="note">前回の点検：' + esc(dateText(d.prevAt)) + '。中身は点検台帳の「変更点」シートにあります。</p>');
    } else {
      h.push('<p class="note">今回は前回の点検台帳を選ばずに作成したため、比べていません。次回、今回の点検台帳を選んで作成すると、変更の件数が入ります。</p>');
    }

    h.push('</section><section><h2>3. 点検の推移</h2><table><tr><th>点検日</th><th class="n">アプリ</th><th class="n">注意点</th><th class="n">要確認</th><th class="n">変更</th></tr>');
    d.history.forEach(function (x) {
      h.push('<tr><td>' + esc(dateText(x.at).slice(0, 10)) + '</td><td class="n">' + x.apps + '個</td><td class="n">' + (known(x.total) ? x.total + '件' : '－') + '</td><td class="n">' +
        (known(x.important) ? x.important + '件' : '－') + '</td><td class="n">' + (known(x.changes) ? x.changes + '件' : '－') + '</td></tr>');
    });
    h.push('</table>');
    if (d.history.length < 2) h.push('<p class="note">前回の点検台帳を選んで点検を続けると、ここに点検ごとの件数が並びます（最新の6回分）。</p>');
    h.push('</section></div>');

    h.push('<h2>4. 要確認の項目</h2>');
    if (d.items.length) {
      h.push('<table class="items"><tr><th style="width:17%">アプリ</th><th style="width:19%">対象</th><th>内容</th></tr>');
      d.items.forEach(function (x) { h.push('<tr><td>' + esc(x.app) + '</td><td>' + esc(x.target) + '</td><td>' + esc(clip(firstSentence(x.text), 44)) + '</td></tr>'); });
      h.push('</table>');
      if (d.more) h.push('<p class="note">ほかに' + d.more + '件あります。すべての項目は、点検台帳の「注意点」シートにあります。</p>');
    } else {
      h.push('<p class="note">要確認の項目はありませんでした。確認の項目は、点検台帳の「注意点」シートにあります。</p>');
    }

    h.push('<h2>5. 所見・対応</h2><div class="end"><div class="memo" contenteditable="true" data-placeholder="対応したこと、対応の予定、担当などを書き込めます"></div>');
    h.push('<div class="sign"><div><span>点検者</span><i></i></div><div><span>確認者</span><i></i></div><div><span>承認者</span><i></i></div></div></div>');
    h.push('<p class="foot">台帳屋 棚卸し・変更監視パックで作成。詳しい内容は、同じときに作成した点検台帳（Excel）の各シートをご覧ください。' + (d.plan ? 'ご契約のプラン：' + esc(d.plan) + '。' : '') + '</p>');
    h.push('</div></body></html>');
    return h.join('');
  }

  // ===== 4. 前回との比較 =====

  var FIELD_TYPES = {
    SINGLE_LINE_TEXT: '文字列（1行）', MULTI_LINE_TEXT: '文字列（複数行）', RICH_TEXT: 'リッチエディター',
    NUMBER: '数値', CALC: '計算', RADIO_BUTTON: 'ラジオボタン', CHECK_BOX: 'チェックボックス',
    MULTI_SELECT: '複数選択', DROP_DOWN: 'ドロップダウン', DATE: '日付', TIME: '時刻', DATETIME: '日時',
    LINK: 'リンク', FILE: '添付ファイル', USER_SELECT: 'ユーザー選択', ORGANIZATION_SELECT: '組織選択',
    GROUP_SELECT: 'グループ選択', REFERENCE_TABLE: '関連レコード一覧', SUBTABLE: 'テーブル', GROUP: 'グループ',
    RECORD_NUMBER: 'レコード番号', CREATOR: '作成者', CREATED_TIME: '作成日時', MODIFIER: '更新者',
    UPDATED_TIME: '更新日時', STATUS: 'ステータス', STATUS_ASSIGNEE: '作業者', CATEGORY: 'カテゴリー'
  };

  function fieldSummary(sig) {
    try {
      var s = JSON.parse(sig);
      var parts = [(FIELD_TYPES[s.t] || s.t) + (s.l ? '「' + s.l + '」' : '')];
      if (s.r) parts.push('必須');
      if (s.u) parts.push('重複禁止');
      if (s.o) parts.push('選択肢：' + s.o.join('、'));
      if (s.x) parts.push('計算式：' + s.x);
      if (s.d !== undefined) parts.push('初期値：' + (Array.isArray(s.d) ? s.d.join('、') : s.d));
      if (s.k) parts.push('ルックアップ：アプリ' + s.k.a + '（キー ' + s.k.f + '）');
      if (s.rt) parts.push('関連レコード：アプリ' + s.rt.a);
      return parts.join(' ／ ');
    } catch (e) { return sig; }
  }

  function fieldName(snap, code) { return ((snap.labels && snap.labels[code]) || code) + '（' + code + '）'; }

  function diffSnapshots(prev, curr) {
    var rows = [];
    function add(kind, app, item, content, before, after) { rows.push([kind, app.id, app.name, item, content, before || '', after || '']); }
    var prevApps = (prev && prev.apps) || {};
    var currApps = curr.apps;
    Object.keys(currApps).forEach(function (id) {
      var c = currApps[id];
      var p = prevApps[id];
      if (!p) { add('追加', c, 'アプリ', '新しく作られたアプリです'); return; }
      if (p.name !== c.name) add('変更', c, 'アプリ名', 'アプリ名が変わりました', p.name, c.name);
      if (p.fields && c.fields) {
        Object.keys(c.fields).forEach(function (code) {
          if (!(code in p.fields)) add('追加', c, 'フィールド', fieldName(c, code), '', fieldSummary(c.fields[code]));
          else if (p.fields[code] !== c.fields[code]) add('変更', c, 'フィールド', fieldName(c, code), fieldSummary(p.fields[code]), fieldSummary(c.fields[code]));
        });
        Object.keys(p.fields).forEach(function (code) {
          if (!(code in c.fields)) add('削除', c, 'フィールド', fieldName(p, code), fieldSummary(p.fields[code]), '');
        });
      }
      [['app', 'アプリのアクセス権'], ['record', 'レコードのアクセス権'], ['field', 'フィールドのアクセス権']].forEach(function (pair) {
        var a = p.acl && p.acl[pair[0]];
        var b = c.acl && c.acl[pair[0]];
        if (a && b && a.join('\n') !== b.join('\n')) add('変更', c, pair[1], '設定が変わりました', a.join('\n') || '（設定なし）', b.join('\n') || '（設定なし）');
      });
      if (p.plugins && c.plugins) {
        var pp = {};
        p.plugins.forEach(function (x) { pp[x.id] = x; });
        var cp = {};
        c.plugins.forEach(function (x) { cp[x.id] = x; });
        c.plugins.forEach(function (x) {
          if (!pp[x.id]) add('追加', c, 'プラグイン', x.name, '', x.enabled ? '有効' : '無効');
          else if (pp[x.id].enabled !== x.enabled) add('変更', c, 'プラグイン', x.name, pp[x.id].enabled ? '有効' : '無効', x.enabled ? '有効' : '無効');
        });
        p.plugins.forEach(function (x) { if (!cp[x.id]) add('削除', c, 'プラグイン', x.name, pp[x.id].enabled ? '有効' : '無効', ''); });
      }
      if (p.customize && c.customize) {
        if (p.customize.scope !== c.customize.scope) add('変更', c, 'カスタマイズ', '適用するユーザー', SCOPE_TEXT[p.customize.scope] || p.customize.scope, SCOPE_TEXT[c.customize.scope] || c.customize.scope);
        var key = function (f) { return f.screen + '・' + f.kind + '：' + f.name; };
        var pf = p.customize.files.map(key);
        var cf = c.customize.files.map(key);
        cf.forEach(function (f) { if (pf.indexOf(f) < 0) add('追加', c, 'カスタマイズ', f); });
        pf.forEach(function (f) { if (cf.indexOf(f) < 0) add('削除', c, 'カスタマイズ', f); });
      }
    });
    Object.keys(prevApps).forEach(function (id) {
      if (!currApps[id]) add('削除', prevApps[id], 'アプリ', '削除されたか、見られなくなったアプリです');
    });
    return rows;
  }

  // ===== 5. 点検台帳のシート =====

  function daysSince(iso, now) {
    if (!iso) return '';
    var t = Date.parse(iso);
    if (isNaN(t)) return '';
    return String(Math.floor(((now ? now.getTime() : Date.now()) - t) / 86400000));
  }

  function dateText(value) {
    if (!value) return '';
    var d = value instanceof Date ? value : new Date(value);
    if (isNaN(d.getTime())) return String(value);
    var pad = function (n) { return ('0' + n).slice(-2); };
    return d.getFullYear() + '-' + pad(d.getMonth() + 1) + '-' + pad(d.getDate()) + ' ' + pad(d.getHours()) + ':' + pad(d.getMinutes());
  }

  var SNAPSHOT_CHUNK = 30000; // Excelの1つのセルに入る文字数（32,767）より少なくする

  var USE_ORDER = { ほかのアプリから使われています: 0, 設定では使われていません: 1, このアプリの中で使われています: 2 };

  // フィールドの使われ方：ほかのアプリから使われているフィールドと、設定のどこにも出てこないフィールド
  function fieldUseRows(apps) {
    var byTarget = {};
    apps.forEach(function (a) {
      if (!a.xref) return;
      var targets = {};
      Object.keys(a.xref.properties).forEach(function (code) {
        var p = a.xref.properties[code];
        var t = (p.lookup && p.lookup.relatedApp.app) || (p.referenceTable && p.referenceTable.relatedApp.app);
        if (t) targets[t] = true;
      });
      Object.keys(a.xref.actions).forEach(function (n) { targets[a.xref.actions[n].destApp.app] = true; });
      Object.keys(targets).forEach(function (t) {
        if (t === a.id) return;
        (byTarget[t] = byTarget[t] || []).push({ appId: a.id, name: a.name, fields: { properties: a.xref.properties }, actions: { actions: a.xref.actions } });
      });
    });
    var rows = [];
    var count = { other: 0, blocked: 0, unused: 0 };
    apps.forEach(function (a) {
      if (!a.fieldUse) return;
      var props = {};
      Object.keys(a.labels || {}).forEach(function (code) { props[code] = { label: a.labels[code] }; });
      var others = byTarget[a.id] || [];
      var hasJs = !!(a.customize && a.customize.files.some(function (f) { return f.kind === 'JavaScript'; }));
      var hasPlugins = !!(a.plugins && a.plugins.length);
      var list = [];
      a.fieldUse.forEach(function (f) {
        var from = others.length ? FieldRefs.find({ appId: a.id, appName: a.name, code: f.c }, { fields: { properties: props } }, [], others)
          .filter(function (h) { return h.scope === 'ほかのアプリ'; }) : [];
        if (SYSTEM_FIELDS[f.t] && !from.length) return; // 消せないフィールドは、ほかのアプリから使われているときだけ載せる
        var blocks = from.some(function (h) { return h.blocks; });
        var judge = from.length ? 'ほかのアプリから使われています' : f.u.length ? 'このアプリの中で使われています' : '設定では使われていません';
        var fromText = [];
        from.forEach(function (h) {
          var t = 'アプリ' + h.appId + '「' + h.appName + '」の' + h.place + (h.blocks ? '★' : '');
          if (fromText.indexOf(t) < 0) fromText.push(t);
        });
        var check = '';
        if (blocks) check = 'ほかのアプリの設定を直すまで、このフィールドは削除できません。';
        else if (from.length) check = '削除や変更の前に、相手のアプリの設定を確かめてください。';
        else if (!f.u.length) {
          check = [hasJs ? 'JavaScriptのカスタマイズの中で使っていないか' : '', hasPlugins ? 'プラグインの設定で使っていないか' : '']
            .filter(Boolean).join('、');
          check = check ? check + '確かめてください。' : '';
        }
        if (from.length) count.other += 1;
        if (blocks) count.blocked += 1;
        if (judge === '設定では使われていません') count.unused += 1;
        list.push([a.id, a.name, f.l, f.c, FIELD_TYPES[f.t] || f.t, judge, f.u.join('、'), fromText.join('、'), check]);
      });
      list.sort(function (x, y) { return USE_ORDER[x[5]] - USE_ORDER[y[5]]; });
      rows = rows.concat(list);
    });
    return { rows: rows, count: count };
  }

  function buildSheets(snapshot, prev, now, licenseText, dir) {
    var apps = Object.keys(snapshot.apps).map(function (id) { return snapshot.apps[id]; })
      .sort(function (a, b) { return Number(a.id) - Number(b.id); });
    var names = {};
    apps.forEach(function (a) { names[a.id] = a.name; });
    var warnings = [];
    var referencedBy = referencedApps(apps);
    var env = snapshot.env || {};
    var allowed = env.allowedPlugins || null;
    var noteOf = function (a) { return a.warnings.concat(usageWarnings(a, now, referencedBy), refWarnings(a, dir), a.health || [], pluginWarnings(a, allowed, env.selfPlugin)); };
    envWarnings(snapshot.env).concat(unusedPlugins(apps, env.installed, env.selfPlugin)).forEach(function (w) { warnings.push([w.level, '', '（環境全体）', w.kind, '', w.target, w.text]); });
    similarApps(apps).forEach(function (w) { warnings.push([w.level, w.app.id, w.app.name, w.kind, '', w.target, w.text]); });
    apps.forEach(function (a) { noteOf(a).forEach(function (w) { warnings.push([w.level, a.id, a.name, w.kind, String(w.order), w.target, w.text]); }); });
    var diff = prev ? diffSnapshots(prev, snapshot) : null;
    var healthCount = warnings.filter(function (w) { return w[3] === '健康診断'; }).length;
    var withErrors = apps.filter(function (a) { return Object.keys(a.errors).length; });
    // 点検の推移（記録シートに入れて、次回の点検台帳に引き継ぐ）
    snapshot.history = historyOf(prev, snapshot.host, inspectionSummary(snapshot.takenAt || (now ? now.toISOString() : ''), apps.length, warnings, diff));
    var notAllowedApps = allowed && allowed.length ? apps.filter(function (a) { return pluginWarnings(a, allowed, env.selfPlugin).length; }).length : 0;

    var cover = [
      ['点検日時', dateText(now)],
      ['kintone環境', snapshot.host],
      ['点検したアプリ', apps.length + '個（点検した人が見られるアプリだけです。見られないアプリは含みません）'],
      ['読み取れなかった項目があるアプリ', withErrors.length + '個' + (withErrors.length ? '（「アプリ一覧」シートの右端の列を参照）' : '')],
      ['注意点（権限・利用状況・ユーザーの指定・健康診断）', warnings.length + '件（うち健康診断 ' + healthCount + '件）'],
      ['前回の点検', prev ? dateText(prev.takenAt) + '（' + (prev.host || '') + '）' : '比較していません'],
      ['前回からの変更', diff ? diff.length + '件' : '－'],
      ['許可するプラグイン', allowed && allowed.length ? allowed.length + '個を指定（許可していないプラグインが入っているアプリ ' + notAllowedApps + '個）' : '指定していません（設定画面で指定すると、健康診断で許可していないプラグインを見つけます）'],
      ['ライセンス', licenseText || '']
    ];
    if (snapshot.plan) cover.push(['ご契約のプラン', snapshot.plan]);
    var historyRows = snapshot.history.slice().reverse().map(function (x) { return [dateText(x.at), historyText(x)]; });
    var day = function (v) { return dateText(v).slice(0, 10); };
    var appRows = apps.map(function (a) {
      return [a.id, a.name, a.spaceId, a.creator, day(a.createdAt), a.modifier, day(a.modifiedAt), daysSince(a.modifiedAt, now),
        a.fields ? String(Object.keys(a.fields).length) : '', a.plugins ? String(a.plugins.length) : '', a.customize ? String(a.customize.files.length) : '',
        a.records ? String(a.records.count) : '', a.records ? day(a.records.lastUpdated) : '', a.records ? daysSince(a.records.lastUpdated, now) : '',
        String(noteOf(a).length), Object.keys(a.errors).map(function (k) { return ERROR_KEYS[k] + '（' + a.errors[k] + '）'; }).join('、')];
    });
    var aclRows = [];
    apps.forEach(function (a) {
      [['app', 'アプリ'], ['record', 'レコード'], ['field', 'フィールド']].forEach(function (pair) {
        (a.acl[pair[0]] || []).forEach(function (line) {
          var parts = line.split('｜');
          if (pair[0] === 'app') aclRows.push([a.id, a.name, pair[1], parts[0], '', parts[1], parts[2]]);
          else aclRows.push([a.id, a.name, pair[1], parts[0], parts[1], parts[2], parts[3]]);
        });
      });
    });
    var pluginRows = [];
    apps.forEach(function (a) {
      (a.plugins || []).forEach(function (p) {
        var ok = isAllowedPlugin(p.id, allowed, env.selfPlugin);
        pluginRows.push([a.id, a.name, p.name, p.id, p.enabled ? '有効' : '無効', ok === null ? '' : ok ? '許可' : '許可していない']);
      });
    });
    var customRows = [];
    apps.forEach(function (a) {
      if (!a.customize) return;
      a.customize.files.forEach(function (f) { customRows.push([a.id, a.name, SCOPE_TEXT[a.customize.scope] || a.customize.scope, f.screen, f.kind, String(f.order), f.form, f.name]); });
    });
    var refRows = [];
    apps.forEach(function (a) { Refs.rows(a.refs || [], dir).forEach(function (r) { refRows.push([r.state, r.typeLabel, r.name, r.code, a.id, a.name, r.place, r.detail]); }); });
    refRows.sort(function (x, y) { var lv = { 見つかりません: 0, 利用停止中: 1 }; return ((x[0] in lv ? lv[x[0]] : 2) - (y[0] in lv ? lv[y[0]] : 2)) || (x[3] < y[3] ? -1 : x[3] > y[3] ? 1 : 0) || (Number(x[4]) - Number(y[4])); });
    var linkRows = [];
    apps.forEach(function (a) { a.links.forEach(function (l) { linkRows.push([l.from, names[l.from] || '', l.kind, l.to, names[l.to] || '（見られないアプリ）', l.place]); }); });
    var use = fieldUseRows(apps);
    cover.splice(5, 0, ['フィールドの使われ方', 'ほかのアプリから使われている ' + use.count.other + '個（うち、kintoneが削除を止めるもの ' + use.count.blocked + '個）、設定では使われていない ' + use.count.unused + '個']);

    var json = JSON.stringify(snapshot);
    var chunks = [[RECORD_MARK]];
    for (var i = 0; i < json.length; i += SNAPSHOT_CHUNK) chunks.push([json.slice(i, i + SNAPSHOT_CHUNK)]);

    var sub = 'kintoneアプリ 点検台帳　／　kintone環境：' + snapshot.host + '　／　点検日時：' + dateText(now);
    var head = function (title, lead) { return { title: title, sub: sub, lead: lead }; };
    var importantCount = warnings.filter(function (w) { return w[0] === '要確認'; }).length;
    var diffBlock = diff
      ? { title: null, header: ['区分', 'アプリID', 'アプリ名', '項目', '内容', '前回', '今回'], rows: diff, note: diff.length ? null : '前回の点検から、変わったところはありません。' }
      : { title: null, header: [], rows: [], note: '前回の点検台帳を選ばずに作成したため、比較していません。次回の点検で、このファイルを選ぶと変更点が分かります。' };
    return [
      { name: '点検報告', widths: [34, 84], orientation: 'portrait', tabColor: 'FF1F3A5F',
        heading: { title: 'kintoneアプリ 点検台帳', sub: 'kintone環境：' + snapshot.host + '　／　点検日時：' + dateText(now) + '　／　作成：台帳屋 棚卸し・変更監視パック' },
        blocks: [
          { title: '点検の概要', header: ['項目', '内容'], rows: cover, labels: true },
          { title: '点検の推移（新しい順）', header: ['点検日時', '結果'], rows: historyRows, labels: true,
            note: historyRows.length < 2 ? '次回の点検で、このファイルを「前回の点検台帳」に選ぶと、点検ごとの件数がここに並びます。' : null },
          Xlsx.contents([
            ['アプリ一覧', '全アプリの作成者・設定の最終更新・フィールド数・レコードの件数と最終更新'],
            ['変更点', '前回の点検台帳と比べて、追加・削除・変更された設定'],
            ['注意点', '見直したほうがよい設定（要確認は赤、確認は白）'],
            ['権限一覧', '全アプリのアプリ・レコード・フィールドのアクセス権'],
            ['ユーザーの指定', '設定に出てくるユーザー・組織・グループ（もういない人を先頭に）'],
            ['プラグイン', 'どのアプリに、どのプラグインが入っているか'],
            ['カスタマイズ', 'どのアプリに、どのJavaScript・CSSが入っているか'],
            ['つながり', 'ルックアップ・関連レコード一覧・アクションでつながるアプリ'],
            ['フィールドの使われ方', 'ほかのアプリから使われているフィールドと、設定では使われていないフィールド']
          ])
        ] },
      { name: 'アプリ一覧', list: true, heading: head('アプリ一覧', '点検したアプリは' + apps.length + '個です。'), widths: [6, 22, 6, 12, 11, 12, 11, 8, 7, 7, 8, 8, 11, 8, 6, 22], blocks: [{ title: null, ids: [0, 2], numbers: [7, 8, 9, 10, 11, 13, 14], header: ['アプリID', 'アプリ名', 'スペースID', '作成者', '作成日時', '設定の最終更新者', '設定の最終更新', '設定の更新からの日数', 'フィールド数', 'プラグイン数', 'カスタマイズ数', 'レコード数', 'レコードの最終更新', 'レコードの更新からの日数', '注意点', '読み取れなかった項目'], rows: appRows }] },
      { name: '変更点', list: true, heading: head('変更点', diff ? '前回の点検（' + dateText(prev.takenAt) + '）からの変更は' + diff.length + '件です。' : '前回の点検台帳を選ばずに作成したため、比較していません。'), widths: [8, 8, 24, 18, 36, 40, 40], blocks: [Object.assign({ ids: [1] }, diffBlock)] },
      { name: '注意点', list: true, heading: head('注意点', '見直したほうがよい設定は' + warnings.length + '件です（うち要確認 ' + importantCount + '件。赤の行）。'), widths: [8, 7, 22, 13, 5, 24, 62], blocks: [{ title: null, ids: [1], numbers: [4], header: ['重要度', 'アプリID', 'アプリ名', '種類', '順', '対象', '内容'], rows: warnings, styles: warnings.map(function (w) { return w[0] === '要確認' ? STYLE.important : STYLE.body; }), note: warnings.length ? null : '注意点は見つかりませんでした。' }] },
      { name: '権限一覧', list: true, heading: head('権限一覧'), widths: [7, 24, 10, 5, 30, 28, 40], blocks: [{ title: null, ids: [0], numbers: [3], header: ['アプリID', 'アプリ名', '種類', '順', '条件・フィールド', '対象', '権限'], rows: aclRows }] },
      { name: 'ユーザーの指定', list: true, heading: head('ユーザーの指定'), widths: [11, 8, 14, 18, 7, 20, 20, 48], blocks: [{ title: null, ids: [4], header: ['状態', '種類', '名前', 'コード', 'アプリID', 'アプリ名', '設定の場所', '内容'], rows: refRows, styles: refRows.map(function (r) { return r[0] === '見つかりません' ? STYLE.warn : STYLE.body; }), note: refRows.length ? null : 'ユーザー・組織・グループを指定している設定はありません。' }] },
      { name: 'プラグイン', list: true, heading: head('プラグイン'), widths: [7, 24, 36, 36, 8, 14], blocks: [{ title: null, ids: [0], header: ['アプリID', 'アプリ名', 'プラグイン名', 'プラグインID', '状態', '許可'], rows: pluginRows,
        styles: pluginRows.map(function (r) { return r[5] === '許可していない' ? STYLE.warn : STYLE.body; }),
        note: pluginRows.length ? (allowed && allowed.length ? null : '「許可」の列は、設定画面で許可するプラグインを指定すると入ります。') : 'プラグインを追加しているアプリはありません。' }] },
      { name: 'カスタマイズ', list: true, heading: head('カスタマイズ'), widths: [7, 24, 18, 10, 12, 5, 8, 60], blocks: [{ title: null, ids: [0], numbers: [5], header: ['アプリID', 'アプリ名', '適用するユーザー', '画面', '種類', '順', '形式', 'ファイル名・URL'], rows: customRows, note: customRows.length ? null : 'JavaScript・CSSによるカスタマイズはありません。' }] },
      { name: 'つながり', list: true, heading: head('つながり'), widths: [10, 24, 16, 10, 24, 36], blocks: [{ title: null, ids: [0, 3], header: ['元のアプリID', '元のアプリ', 'つながりの種類', '相手のアプリID', '相手のアプリ', '設定のある場所'], rows: linkRows, note: linkRows.length ? null : 'ルックアップ・関連レコード一覧・アクションでつながっているアプリはありません。' }] },
      { name: 'フィールドの使われ方', list: true, heading: head('フィールドの使われ方'), widths: [6, 18, 18, 14, 11, 18, 28, 32, 30], blocks: [{ title: null, ids: [0], header: ['アプリID', 'アプリ名', 'フィールド名', 'フィールドコード', '種類', '判定', 'このアプリでの使われ方', 'ほかのアプリからの使われ方', '確かめること'], rows: use.rows,
        styles: use.rows.map(function (r) { return /★/.test(r[7]) ? STYLE.warn : STYLE.body; }),
        note: use.rows.length
          ? '★：使われている間は、kintoneがフィールドの削除を止める設定です（黄の行）。'
          : 'フィールドの設定を読み取れたアプリがありません。' }].concat(use.rows.length ? [
        // 長い注記を1つのセルに入れると、印刷の幅が広がって字が小さくなるため、短く分ける
        { title: null, header: [], rows: [], note: '「設定では使われていません」：計算式・一覧・グラフ・アクセス権・プロセス管理・通知・アクション・レコードのタイトルと、ほかのアプリのどれにも出てきません。' },
        { title: null, header: [], rows: [], note: '利用者が値を入力して使っている場合や、JavaScript・プラグインの設定の中で使っている場合があるため、削除の前に確かめてください。' },
        { title: null, header: [], rows: [], note: '点検した人が見られないアプリで使われているフィールドは、ここに出ません（kintoneシステム管理者でない人が点検したときは、特に注意してください）。' }
      ] : []) },
      { name: RECORD_SHEET, hidden: true, widths: [20], blocks: [{ title: null, header: [], rows: chunks }] }
    ];
  }

  var ERROR_KEYS = { appSettings: 'レコードのタイトル', records: 'レコードの件数', status: 'プロセス管理', general: 'アプリの条件通知', perRecord: 'レコードの条件通知', reminder: 'リマインダーの条件通知', views: '一覧', reports: 'グラフ', fields: 'フィールド', appAcl: 'アプリのアクセス権', recordAcl: 'レコードのアクセス権', fieldAcl: 'フィールドのアクセス権', plugins: 'プラグイン', customize: 'カスタマイズ', actions: 'アクション', adminNotes: 'アプリ管理者用メモ' };

  function summaryText(snapshot, sheets, compared) {
    var count = Object.keys(snapshot.apps).length;
    var warn = sheets[3].blocks[0].rows.length;
    var diff = sheets[2].blocks[0].rows.length;
    return count + '個のアプリを点検しました。注意点は' + warn + '件' + (compared ? '、前回からの変更は' + diff + '件' : '') + 'です。';
  }

  // 無料の点検で見せるまとめ：件数と、種類ごとに例を少しだけ（一覧とExcelは有料版）
  var FREE_GROUPS = [
    { name: '権限', kinds: ['アプリ', 'レコード', 'フィールド'] },
    { name: 'ユーザーの指定（退職・異動）', kinds: ['ユーザーの指定'] },
    { name: '使われていないかもしれないアプリ', kinds: ['利用状況'] },
    { name: '健康診断（ガバナンスガイドラインの点検項目など）', kinds: ['健康診断'] }
  ];

  function freeSummary(snapshot, sheets, examplesPerGroup) {
    var per = examplesPerGroup === undefined ? 2 : examplesPerGroup;
    var rows = sheets[3].blocks[0].rows;
    var groups = FREE_GROUPS.map(function (g) {
      var hit = rows.filter(function (r) { return g.kinds.indexOf(r[3]) >= 0; });
      hit = hit.slice().sort(function (x, y) { return (x[0] === '要確認' ? 0 : 1) - (y[0] === '要確認' ? 0 : 1); });
      return {
        name: g.name, count: hit.length, important: hit.filter(function (r) { return r[0] === '要確認'; }).length,
        examples: hit.slice(0, per).map(function (r) { return { app: r[2], target: r[5], text: r[6] }; })
      };
    });
    var use = sheets[9] && sheets[9].blocks[0].rows;
    var fieldCount = { other: 0, blocked: 0, unused: 0 };
    (use || []).forEach(function (r) {
      if (r[5] === 'ほかのアプリから使われています') { fieldCount.other += 1; if (/★/.test(r[7])) fieldCount.blocked += 1; }
      if (r[5] === '設定では使われていません') fieldCount.unused += 1;
    });
    var apps = Object.keys(snapshot.apps).map(function (id) { return snapshot.apps[id]; });
    return {
      apps: apps.length,
      total: rows.length,
      important: rows.filter(function (r) { return r[0] === '要確認'; }).length,
      groups: groups,
      fields: fieldCount,
      unreadable: apps.filter(function (a) { return Object.keys(a.errors).length; }).length
    };
  }

  // キーに「アプリの数の上限（a）」があるときの確かめ。上限を超えても止めず、プランの変更をお願いする
  function planNote(payload, appCount) {
    var max = payload && Number(payload.a);
    if (!max) return { over: false, text: '' };
    var over = appCount > max;
    return { over: over, text: 'アプリ' + max + '個まで' + (over ? '（この環境で点検したアプリは' + appCount + '個です。プランの変更をお願いします）' : '') };
  }

  // 料金のプラン（税込）。pro.html の料金と同じにする。apps：アプリの数の上限（0は上限なし）
  var PLANS = [
    { apps: 100, name: 'アプリ100個まで', month: 3300, year: 33000 },
    { apps: 500, name: 'アプリ500個まで', month: 9900, year: 99000 },
    { apps: 0, name: 'アプリの数の上限なし', month: 19800, year: 198000 }
  ];

  function recommendPlan(appCount) {
    for (var i = 0; i < PLANS.length; i++) if (!PLANS[i].apps || appCount <= PLANS[i].apps) return PLANS[i];
    return PLANS[PLANS.length - 1];
  }

  function yen(n) { return String(n).replace(/\B(?=(\d{3})+(?!\d))/g, ',') + '円'; }

  function planText(appCount) {
    var p = recommendPlan(appCount);
    return 'この環境のアプリは' + appCount + '個なので、プランは「' + p.name + '」（月額' + yen(p.month) + '・年額' + yen(p.year) + '、税込）です。';
  }

  var Core = {
    freeSummary: freeSummary, planNote: planNote, recommendPlan: recommendPlan, planText: planText, healthChecks: healthChecks, envWarnings: envWarnings,
    unusedPlugins: unusedPlugins, similarApps: similarApps,
    pluginWarnings: pluginWarnings, historyOf: historyOf, inspectionSummary: inspectionSummary, reportData: reportData, reportHtml: reportHtml,
    buildXlsx: buildXlsx, readSnapshot: readSnapshot, readZip: readZip,
    appSnapshot: appSnapshot, diffSnapshots: diffSnapshots, buildSheets: buildSheets, summaryText: summaryText,
    fieldSignature: fieldSignature, warningsFor: warningsFor, extractLinks: extractLinks, entityText: entityText, dateText: dateText
  };
  if (typeof module !== 'undefined' && module.exports) { module.exports = Core; return; }
  if (typeof kintone === 'undefined') return;

  // ===== 6. 画面（kintone上でだけ動く部分） =====

  var P = 'dcy-ip';

  function el(tag, attrs, children) {
    var node = document.createElement(tag);
    Object.keys(attrs || {}).forEach(function (key) {
      if (key === 'text') node.textContent = attrs[key];
      else node.setAttribute(key, attrs[key]);
    });
    (children || []).forEach(function (child) { node.appendChild(typeof child === 'string' ? document.createTextNode(child) : child); });
    return node;
  }

  function api(path, params) { return kintone.api(kintone.api.url(path, true), 'GET', params); }
  // cybozu.com の User API（ユーザー・組織・グループがいまもあるか）
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

  // レコードの中身は読まない。更新日時のフィールドだけを1件取り、件数と合わせて確かめる
  function recordRequest(appId, fieldsResult) {
    var props = (fieldsResult.ok && fieldsResult.value.properties) || {};
    var code = Object.keys(props).filter(function (k) { return props[k].type === 'UPDATED_TIME'; })[0];
    if (!code) return Promise.resolve(null);
    return settle(api('/k/v1/records.json', { app: appId, fields: [code], query: 'order by ' + code + ' desc limit 1', totalCount: true }));
  }

  function collect(appId, withRecords) {
    var requests = {
      fields: api('/k/v1/app/form/fields.json', { app: appId }),
      appAcl: api('/k/v1/app/acl.json', { app: appId }),
      recordAcl: api('/k/v1/record/acl.json', { app: appId }),
      fieldAcl: api('/k/v1/field/acl.json', { app: appId }),
      plugins: api('/k/v1/app/plugins.json', { app: appId }),
      customize: api('/k/v1/app/customize.json', { app: appId }),
      actions: api('/k/v1/app/actions.json', { app: appId }),
      status: api('/k/v1/app/status.json', { app: appId }),
      general: api('/k/v1/app/notifications/general.json', { app: appId }),
      perRecord: api('/k/v1/app/notifications/perRecord.json', { app: appId }),
      reminder: api('/k/v1/app/notifications/reminder.json', { app: appId }),
      views: api('/k/v1/app/views.json', { app: appId }),
      reports: api('/k/v1/app/reports.json', { app: appId }),
      appSettings: api('/k/v1/app/settings.json', { app: appId }),
      adminNotes: api('/k/v1/app/adminNotes.json', { app: appId })
    };
    var keys = Object.keys(requests);
    return Promise.all(keys.map(function (k) { return settle(requests[k]); })).then(function (list) {
      var results = {};
      keys.forEach(function (k, i) { results[k] = list[i]; });
      if (!withRecords) return results;
      return recordRequest(appId, results.fields).then(function (r) { if (r) results.records = r; return results; });
    });
  }

  // 同時に調べるアプリは3つまでにして、kintoneに負担をかけない
  function scanAll(apps, withRecords, onProgress) {
    var snapshots = {};
    var index = 0;
    var done = 0;
    function worker() {
      if (index >= apps.length) return Promise.resolve();
      var app = apps[index++];
      return collect(app.appId, withRecords).then(function (results) {
        snapshots[app.appId] = appSnapshot(app, results);
        done += 1;
        onProgress(done, apps.length);
        return worker();
      });
    }
    var workers = [];
    for (var i = 0; i < Math.min(3, apps.length); i++) workers.push(worker());
    return Promise.all(workers).then(function () { return snapshots; });
  }

  function download(bytes, fileName, type) {
    var blob = new Blob([bytes], { type: type || 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    var url = URL.createObjectURL(blob);
    var a = el('a', { href: url, download: fileName });
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    setTimeout(function () { URL.revokeObjectURL(url); }, 1000);
  }

  function readFile(file) {
    return new Promise(function (resolve, reject) {
      var reader = new FileReader();
      reader.onload = function () { resolve(new Uint8Array(reader.result)); };
      reader.onerror = function () { reject(new Error('ファイルを読み込めませんでした。')); };
      reader.readAsArrayBuffer(file);
    });
  }

  function licenseKey() {
    try { return (kintone.plugin.app.getConfig(PLUGIN_ID) || {}).license || ''; } catch (e) { return ''; }
  }

  // 設定画面で選んだ「許可するプラグイン」。指定がなければ null（点検しない）
  function allowedPlugins() {
    try {
      var list = JSON.parse((kintone.plugin.app.getConfig(PLUGIN_ID) || {}).allowedPlugins || '[]');
      return Array.isArray(list) && list.length ? list : null;
    } catch (e) { return null; }
  }

  function loginName() {
    try { var u = kintone.getLoginUser(); return u.name || u.code || ''; } catch (e) { return ''; }
  }

  // 報告書は新しいタブに開く。開けないとき（ポップアップを止める設定など）は、HTMLのファイルとして保存する
  function openReport(html, stamp) {
    var w = window.open('', '_blank');
    if (!w) { download(utf8(html), '点検報告書_' + stamp + '.html', 'text/html'); return false; }
    w.document.open();
    w.document.write(html);
    w.document.close();
    var btn = w.document.getElementById('dcy-print');
    if (btn) btn.addEventListener('click', function () { w.print(); });
    return true;
  }

  var TRIAL_FORM = 'https://docs.google.com/forms/d/e/1FAIpQLSeazyhYt9ZZf3fJufgTl02byKzbnpDxQjxyjDhJAbuV_heLbQ/viewform';
  var TRIAL_DOMAIN_ENTRY = 'entry.1870255768';
  var PRICE_URL = 'https://daichoya.github.io/pro.html#price';

  function trialUrl(domain) { return TRIAL_FORM + '?usp=pp_url&' + TRIAL_DOMAIN_ENTRY + '=' + encodeURIComponent(domain); }

  // 前回の点検の日（このブラウザに残すだけ。読めないときは何も表示しない）
  function lastRunKey() { return P + '-last-' + location.hostname; }
  function readLastRun() { try { return window.localStorage.getItem(lastRunKey()) || ''; } catch (e) { return ''; } }
  function saveLastRun(date) { try { window.localStorage.setItem(lastRunKey(), date); } catch (e) { /* 保存できなくても点検は続ける */ } }

  // kintoneに読み込んであるプラグインの一覧（id と名前だけ）。読めないときは null
  function installedPlugins() {
    var list = [];
    function next(offset) {
      return api('/k/v1/plugins.json', { offset: offset, limit: 100 }).then(function (resp) {
        list = list.concat((resp.plugins || []).map(function (p) { return { id: p.id, name: p.name }; }));
        return (resp.plugins || []).length === 100 ? next(offset + 100) : list;
      });
    }
    return next(0).catch(function () { return null; });
  }

  // 初期の管理者アカウント（Administrator）が使えるか。確かめられないときは null
  function administratorState() {
    return userApi('/v1/users.json', { codes: ['Administrator'] }).then(function (resp) {
      var u = (resp.users || [])[0];
      return u ? u.valid !== false : false;
    }, function () { return null; });
  }

  function scan(withRecords, status) {
    var env = { allowedPlugins: allowedPlugins(), selfPlugin: PLUGIN_ID };
    status.textContent = 'アプリの一覧を読み込み中…';
    return Promise.all([getAllApps(), administratorState(), installedPlugins()]).then(function (r) {
      env.administrator = r[1];
      env.installed = r[2];
      return scanAll(r[0], withRecords, function (done, total) { status.textContent = 'アプリの設定を読み込み中…（' + done + ' / ' + total + '）'; });
    }).then(function (apps) {
      status.textContent = '設定に出てくるユーザー・組織・グループを確認中…';
      var all = [];
      Object.keys(apps).forEach(function (id) { all = all.concat(apps[id].refs || []); });
      return Refs.lookup(userApi, all).then(function (dir) { return { apps: apps, dir: dir, env: env }; });
    });
  }

  function renderFree(box, sum, domain) {
    box.innerHTML = '';
    box.appendChild(el('h3', { text: '無料の点検の結果' }));
    box.appendChild(el('p', { class: P + '-lead', text: sum.apps + '個のアプリを点検し、見直したほうがよい設定が' + sum.total + '件見つかりました（うち要確認 ' + sum.important + '件）。' }));
    var table = el('table', { class: P + '-table' });
    table.appendChild(el('tr', {}, [el('th', { text: '種類' }), el('th', { text: '件数' }), el('th', { text: '例（すべての一覧はご契約で）' })]));
    sum.groups.forEach(function (g) {
      var ex = el('td', {});
      if (!g.examples.length) ex.appendChild(document.createTextNode('－'));
      g.examples.forEach(function (x) { ex.appendChild(el('div', { text: '「' + x.app + '」' + (x.target ? x.target + '：' : '') + x.text })); });
      if (g.count > g.examples.length) ex.appendChild(el('div', { class: P + '-more-count', text: 'ほか ' + (g.count - g.examples.length) + '件' }));
      table.appendChild(el('tr', {}, [el('td', { text: g.name }), el('td', { text: g.count + '件' + (g.important ? '（要確認 ' + g.important + '）' : '') }), ex]));
    });
    table.appendChild(el('tr', {}, [el('td', { text: 'フィールドの使われ方' }), el('td', { text: (sum.fields.other + sum.fields.unused) + '個' }),
      el('td', { text: 'ほかのアプリから使われている ' + sum.fields.other + '個（うち削除できないもの ' + sum.fields.blocked + '個）、設定では使われていない ' + sum.fields.unused + '個' })]));
    box.appendChild(table);
    if (sum.unreadable) box.appendChild(el('p', { class: P + '-note', text: '権限がなく、設定の一部を読み取れなかったアプリが' + sum.unreadable + '個あります。kintoneシステム管理者が点検すると、すべて読み取れます。' }));
    box.appendChild(el('div', { class: P + '-cta' }, [
      el('p', { text: 'ご契約（30日間は無料で試せます）で、すべての一覧と、どのアプリのどの設定かをExcelの点検台帳にまとめ、上司や監査に出せる1枚の報告書も作ります。次回からは、前回の点検から変わったところと、件数の推移も分かります。' }),
      el('p', { class: P + '-note', text: planText(sum.apps) }),
      el('div', { class: P + '-controls' }, [
        el('a', { class: P + '-btn ' + P + '-primary ' + P + '-link', href: trialUrl(domain), target: '_blank', rel: 'noopener', text: '30日間無料で、すべてを見る' }),
        el('a', { class: P + '-btn ' + P + '-link', href: PRICE_URL, target: '_blank', rel: 'noopener', text: '料金を見る' })
      ]),
      el('p', { class: P + '-note', text: '無料試用は、会社名・お名前・メールアドレスを入れるだけで申し込めます（kintoneの環境名は入力済みです）。数分でライセンスキーがメールで届きます。カードの登録は要りません。' })
    ]));
  }

  // 有料版：点検台帳を作ったあと、画面にも結果のまとめを出す（d：reportData の結果）
  function renderPaid(box, d) {
    box.innerHTML = '';
    box.appendChild(el('h3', { text: '点検の結果' }));
    box.appendChild(el('p', { class: P + '-lead', text: d.apps + '個のアプリを点検し、見直したほうがよい設定が' + d.total + '件見つかりました（うち要確認 ' + d.important + '件）。' +
      (d.before && known(d.before.total) ? '前回（' + dateText(d.before.at).slice(0, 10) + '）は' + d.before.total + '件でした。' : '') }));
    var withBefore = d.groups.some(function (g) { return known(g.before); });
    var table = el('table', { class: P + '-table' });
    table.appendChild(el('tr', {}, [el('th', { text: '種類' }), el('th', { text: '件数' }), el('th', { text: 'うち要確認' })].concat(withBefore ? [el('th', { text: '前回' }), el('th', { text: '増減' })] : [])));
    d.groups.forEach(function (g) {
      table.appendChild(el('tr', {}, [el('td', { text: g.name }), el('td', { text: g.count + '件' }), el('td', { text: g.important ? g.important + '件' : '－' })]
        .concat(withBefore ? [el('td', { text: known(g.before) ? g.before + '件' : '－' }), el('td', { text: known(g.before) ? signed(g.count - g.before) : '－' })] : [])));
    });
    box.appendChild(table);
    if (d.diff) {
      var parts = d.diff.filter(function (x) { return x.add || x.change || x.remove; }).map(function (x) {
        return x.name + '（' + [x.add ? '追加 ' + x.add : '', x.change ? '変更 ' + x.change : '', x.remove ? '削除 ' + x.remove : ''].filter(Boolean).join('・') + '）';
      });
      box.appendChild(el('p', { class: P + '-lead', text: '前回の点検からの変更：' + (parts.length ? parts.join('、') : 'ありません') }));
    }
    if (d.items.length) {
      box.appendChild(el('p', { class: P + '-lead', text: '要確認の項目（先に確かめてください）' }));
      var list = el('ul', { class: P + '-items' });
      d.items.slice(0, 5).forEach(function (x) { list.appendChild(el('li', { text: '「' + x.app + '」' + (x.target ? x.target + '：' : '') + firstSentence(x.text) })); });
      box.appendChild(list);
    }
    box.appendChild(el('p', { class: P + '-note', text: 'すべての一覧は、保存した点検台帳（Excel）の「注意点」のシートにあります。' + (d.unreadable ? '権限がなく、設定の一部を読み取れなかったアプリが' + d.unreadable + '個あります。' : '') }));
  }

  function openModal() {
    if (document.getElementById(P + '-overlay')) return;
    var domain = window.DaichoyaLicense.normalizeDomain(location.hostname);
    var closeBtn = el('button', { type: 'button', class: P + '-close', 'aria-label': '閉じる', text: '×' });
    var licenseLine = el('p', { class: P + '-license', text: 'ライセンスを確認しています…' });
    var fileInput = el('input', { type: 'file', accept: '.xlsx', class: P + '-file', id: P + '-prev' });
    var runBtn = el('button', { type: 'button', class: P + '-btn ' + P + '-primary', text: '点検台帳を作成', disabled: 'disabled' });
    var freeBtn = el('button', { type: 'button', class: P + '-btn', text: '無料で点検する（件数だけ）', disabled: 'disabled' });
    var recordsBox = el('input', { type: 'checkbox', id: P + '-records', checked: 'checked' });
    var status = el('p', { class: P + '-status', role: 'status' });
    var result = el('div', { class: P + '-result' });
    var last = readLastRun();
    var prevBlock = el('div', { class: P + '-prev' }, [el('label', { for: P + '-prev', text: '前回の点検台帳（任意）' }), fileInput]);
    var dialog = el('div', { class: P + '-dialog', role: 'dialog', 'aria-modal': 'true', 'aria-label': '棚卸し・変更監視パック' }, [
      el('div', { class: P + '-head' }, [el('h2', { text: '棚卸し・変更監視パック' }), closeBtn]),
      licenseLine,
      el('p', { class: P + '-lead', text: '全アプリをまとめて点検します。件数だけの点検は無料です。ご契約（30日間の無料試用あり）で、すべての一覧を点検台帳（Excel）と報告書にまとめ、前回の点検台帳を選ぶと、変わったところと点検の推移も分かります。' }),
      last ? el('p', { class: P + '-note', text: '前回の点検：' + last + '（このブラウザで点検した日）' }) : el('span', {}),
      prevBlock,
      el('div', { class: P + '-option' }, [recordsBox, el('label', { for: P + '-records', text: '使われていないアプリも調べる（レコードの件数と最終更新日時だけを見ます。中身は読みません）' })]),
      el('div', { class: P + '-controls' }, [runBtn, freeBtn]),
      status,
      result
    ]);
    var overlay = el('div', { id: P + '-overlay', class: P + '-overlay' }, [dialog]);
    function close() {
      document.removeEventListener('keydown', onKey);
      if (overlay.parentNode) overlay.parentNode.removeChild(overlay);
    }
    function onKey(e) { if (e.key === 'Escape') close(); }
    closeBtn.addEventListener('click', close);
    overlay.addEventListener('click', function (e) { if (e.target === overlay) close(); });
    document.addEventListener('keydown', onKey);
    document.body.appendChild(overlay);

    var licenseText = '';
    var lic = null;
    window.DaichoyaLicense.check(licenseKey(), location.hostname).then(function (l) {
      lic = l;
      licenseText = l.message;
      licenseLine.textContent = 'ライセンス：' + l.message;
      licenseLine.className = P + '-license ' + P + '-license-' + l.state;
      freeBtn.removeAttribute('disabled');
      if (l.ok) {
        runBtn.removeAttribute('disabled');
        status.textContent = '「点検台帳を作成」を押してください。アプリの数によっては数分かかります。';
      } else {
        runBtn.style.display = 'none';
        prevBlock.style.display = 'none';
        freeBtn.className = P + '-btn ' + P + '-primary';
        licenseLine.textContent = l.state === 'none' ? '無料の点検（全アプリの件数と例）は、申し込みやライセンスキーなしで使えます。' : 'ライセンス：' + l.message;
        status.textContent = '「無料で点検する」を押すと、すべてのアプリを点検して、見直したほうがよい設定の件数と例を表示します。';
      }
    });

    function done(text) {
      status.textContent = text;
      saveLastRun(dateText(new Date()).slice(0, 10));
      runBtn.removeAttribute('disabled');
      freeBtn.removeAttribute('disabled');
    }
    function fail(error) {
      status.textContent = '点検できませんでした（' + ((error && error.message) || '原因不明') + '）。';
      if (lic && lic.ok) runBtn.removeAttribute('disabled');
      freeBtn.removeAttribute('disabled');
    }

    freeBtn.addEventListener('click', function () {
      freeBtn.setAttribute('disabled', 'disabled');
      runBtn.setAttribute('disabled', 'disabled');
      result.innerHTML = '';
      scan(recordsBox.checked, status).then(function (sc) {
        var now = new Date();
        var snapshot = { v: 1, host: domain, takenAt: now.toISOString(), apps: sc.apps, env: sc.env };
        var sheets = buildSheets(snapshot, null, now, licenseText, sc.dir);
        renderFree(result, freeSummary(snapshot, sheets), domain);
        done(Object.keys(sc.apps).length + '個のアプリを点検しました。');
      }).catch(fail);
    });

    runBtn.addEventListener('click', function () {
      runBtn.setAttribute('disabled', 'disabled');
      freeBtn.setAttribute('disabled', 'disabled');
      result.innerHTML = '';
      var file = fileInput.files && fileInput.files[0];
      var prev;
      status.textContent = file ? '前回の点検台帳を読み込み中…' : 'アプリの一覧を読み込み中…';
      (file ? readFile(file).then(readSnapshot) : Promise.resolve(null)).then(function (p) {
        prev = p;
        return scan(recordsBox.checked, status);
      }).then(function (sc) {
        var now = new Date();
        var count = Object.keys(sc.apps).length;
        var plan = planNote(lic && lic.payload, count);
        var snapshot = { v: 1, host: domain, takenAt: now.toISOString(), apps: sc.apps, env: sc.env, plan: plan.text || undefined };
        var sheets = buildSheets(snapshot, prev, now, licenseText, sc.dir);
        var stamp = dateText(now).slice(0, 10).replace(/-/g, '');
        download(buildXlsx(sheets), '点検台帳_' + stamp + '.xlsx');
        var data = reportData(snapshot, prev, sheets, now, loginName());
        var report = reportHtml(data);
        renderPaid(result, data);
        var reportBtn = el('button', { type: 'button', class: P + '-btn ' + P + '-primary', text: '点検の報告書を開く（印刷・PDF）' });
        reportBtn.addEventListener('click', function () {
          if (!openReport(report, stamp)) status.textContent = '新しいタブを開けなかったため、報告書をHTMLのファイルで保存しました。ブラウザで開いて印刷できます。';
        });
        // 報告書のボタンは、結果のまとめより上に置く（いちばん使うもの）
        result.insertBefore(el('div', { class: P + '-report' }, [
          el('p', { text: '上司や監査に出せる1枚の報告書です。点検の結果・前回からの変更・点検の推移・要確認の項目と、確認の欄があります。' }),
          el('div', { class: P + '-controls' }, [reportBtn])
        ]), result.firstChild);
        if (plan.over) result.appendChild(el('p', { class: P + '-note ' + P + '-warn', text: 'ご契約のプランは' + plan.text + '。' + planText(count) + 'プランの変更は、daichoya.plugins@gmail.com までメールでお知らせください。' }));
        done(summaryText(snapshot, sheets, !!prev) + 'ファイルは次回の点検で「前回の点検台帳」として使えます。');
      }).catch(fail);
    });
  }

  kintone.events.on('app.record.index.show', function (event) {
    var space = kintone.app.getHeaderMenuSpaceElement();
    if (!space || document.getElementById(P + '-open')) return event;
    var button = el('button', { id: P + '-open', type: 'button', class: P + '-btn', text: '棚卸し' });
    button.addEventListener('click', openModal);
    space.appendChild(button);
    return event;
  });
})(typeof kintone !== 'undefined' ? kintone.$PLUGIN_ID : null);
