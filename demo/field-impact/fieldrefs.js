/*
 * 台帳屋 共通部品：フィールドが使われている場所を探す
 * フィールドの影響調査と、棚卸し・変更監視パックで使う。読み取るのは設定だけで、レコードの中身は読まない。
 */
(function (root) {
  'use strict';

  // フィールドコードに使える文字（ひらがな・カタカナ・漢字・英数字・_＿・･・全角の通貨記号）
  var CODE_CHAR = '[\\p{L}\\p{N}_＿･・\\p{Sc}]';

  function escapeRe(s) { return String(s).replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); }

  // 式や条件の中に、そのフィールドコードが「ことば」として出てくるか（"…" の中の文字は除く）
  function mentions(text, code) {
    if (!text || !code) return false;
    var plain = String(text).replace(/"(?:[^"\\]|\\.)*"/g, '""');
    var re = new RegExp('(^|[^' + CODE_CHAR.slice(1, -1) + '])' + escapeRe(code) + '(?=$|[^' + CODE_CHAR.slice(1, -1) + '])', 'u');
    return re.test(plain);
  }

  // JavaScriptの中に、フィールドコードが文字列として出てくるか（'コード' "コード" `コード`）
  function mentionsInScript(text, code) {
    if (!text || !code) return false;
    var c = escapeRe(code);
    return new RegExp('([\'"`])' + c + '\\1').test(text);
  }

  function flatFields(properties) {
    var map = {};
    Object.keys(properties || {}).forEach(function (code) {
      var p = properties[code];
      map[code] = { prop: p, table: null };
      if (p.type === 'SUBTABLE' && p.fields) Object.keys(p.fields).forEach(function (inner) { map[inner] = { prop: p.fields[inner], table: p.label || code }; });
    });
    return map;
  }

  function label(fields, code) {
    var f = fields[code];
    return f ? (f.prop.label || code) : code;
  }

  function byIndex(obj) {
    return Object.keys(obj || {}).map(function (k) { return obj[k]; })
      .sort(function (a, b) { return Number(a.index || 0) - Number(b.index || 0); });
  }

  // 調べる対象：選んだフィールド。テーブルを選んだときは、テーブルの中のフィールドも対象にする
  function targetCodes(fields, code) {
    var f = fields[code];
    var codes = [code];
    if (f && f.prop.type === 'SUBTABLE') Object.keys(f.prop.fields || {}).forEach(function (inner) { codes.push(inner); });
    return codes;
  }

  // settings：調べるアプリの設定のAPIの応答（読めなかったものは null）
  //   { fields, views, reports, recordAcl, fieldAcl, status, general, perRecord, reminder, actions, appSettings }
  // scripts：[{ screen, name, text }]（読めたJavaScriptの中身）
  // others：[{ appId, name, fields, actions }]（ほかのアプリの設定）
  // 戻り値：[{ scope, appId, appName, place, detail, blocks }]  blocks=true は、kintoneが削除を止める設定
  function find(target, settings, scripts, others) {
    var s = settings || {};
    var fields = flatFields(s.fields && s.fields.properties);
    var codes = targetCodes(fields, target.code);
    var out = [];
    var seen = {};
    function add(scope, appId, appName, place, detail, blocks) {
      var key = [scope, appId, place, detail].join('\n');
      if (seen[key]) return;
      seen[key] = true;
      out.push({ scope: scope, appId: String(appId), appName: appName || '', place: place, detail: detail, blocks: !!blocks });
    }
    var here = function (place, detail, blocks) { add('このアプリ', target.appId, target.appName, place, detail, blocks); };
    var named = function (code) { return '「' + label(fields, code) + '」'; };
    var any = function (fn) { return codes.some(fn); };
    var hit = function (text) { return codes.filter(function (c) { return mentions(text, c); }); };

    // フォーム：計算式、ルックアップ、関連レコード一覧
    Object.keys(fields).forEach(function (code) {
      var p = fields[code].prop;
      if (codes.indexOf(code) >= 0) return;
      if (p.expression && hit(p.expression).length) here('計算式', 'フィールド' + named(code) + '：' + p.expression);
      if (p.lookup) {
        (p.lookup.fieldMappings || []).forEach(function (m) {
          if (codes.indexOf(m.field) >= 0) here('ルックアップのコピー先', 'ルックアップ' + named(code) + 'が、' + named(m.field) + 'に値をコピーしています');
        });
      }
      if (p.referenceTable && p.referenceTable.condition && codes.indexOf(p.referenceTable.condition.field) >= 0) {
        here('関連レコード一覧の条件', '関連レコード一覧' + named(code) + 'が、' + named(p.referenceTable.condition.field) + 'の値で絞り込んでいます');
      }
    });
    codes.forEach(function (c) {
      var p = fields[c] && fields[c].prop;
      if (p && p.lookup) here('ルックアップ', named(c) + 'はルックアップです（参照先：アプリ' + ((p.lookup.relatedApp && p.lookup.relatedApp.app) || '') + '）');
    });

    // 一覧・グラフ
    byIndex(s.views && s.views.views).forEach(function (v) {
      if ((v.fields || []).some(function (f) { return codes.indexOf(f) >= 0; })) here('一覧の表示項目', '一覧「' + v.name + '」');
      if (hit(v.filterCond).length) here('一覧の絞り込み', '一覧「' + v.name + '」：' + v.filterCond);
      if (hit(v.sort).length) here('一覧の並び順', '一覧「' + v.name + '」：' + v.sort);
    });
    byIndex(s.reports && s.reports.reports).forEach(function (r) {
      if ((r.groups || []).some(function (g) { return codes.indexOf(g.code) >= 0; })) here('グラフの分類', 'グラフ「' + r.name + '」');
      if ((r.aggregations || []).some(function (g) { return codes.indexOf(g.code) >= 0; })) here('グラフの集計', 'グラフ「' + r.name + '」');
      if (hit(r.filterCond).length) here('グラフの絞り込み', 'グラフ「' + r.name + '」：' + r.filterCond);
    });

    // アクセス権
    ((s.recordAcl && s.recordAcl.rights) || []).forEach(function (rule, i) {
      if (hit(rule.filterCond).length) here('レコードのアクセス権の条件', (i + 1) + '番目の条件：' + rule.filterCond);
      (rule.entities || []).forEach(function (e) {
        if (e.entity && e.entity.type === 'FIELD_ENTITY' && codes.indexOf(e.entity.code) >= 0) here('レコードのアクセス権の対象', (i + 1) + '番目の条件で、' + named(e.entity.code) + 'に入っている人に権限を付けています');
      });
    });
    ((s.fieldAcl && s.fieldAcl.rights) || []).forEach(function (rule) {
      if (codes.indexOf(rule.code) >= 0) here('フィールドのアクセス権', named(rule.code) + 'にアクセス権を設定しています');
      (rule.entities || []).forEach(function (e) {
        if (e.entity && e.entity.type === 'FIELD_ENTITY' && codes.indexOf(e.entity.code) >= 0) here('フィールドのアクセス権の対象', named(rule.code) + 'の権限を、' + named(e.entity.code) + 'に入っている人に付けています');
      });
    });

    // プロセス管理
    if (s.status && s.status.enable) {
      byIndex(s.status.states).forEach(function (st) {
        ((st.assignee && st.assignee.entities) || []).forEach(function (e) {
          if (e.entity && e.entity.type === 'FIELD_ENTITY' && codes.indexOf(e.entity.code) >= 0) here('プロセス管理の作業者', 'ステータス「' + st.name + '」の作業者に' + named(e.entity.code) + 'を使っています', true);
        });
      });
      (s.status.actions || []).forEach(function (a) {
        if (hit(a.filterCond).length) here('プロセス管理の条件', 'アクション「' + a.name + '」：' + a.filterCond);
        ((a.executableUser && a.executableUser.entities) || []).forEach(function (e) {
          if (e.entity && e.entity.type === 'FIELD_ENTITY' && codes.indexOf(e.entity.code) >= 0) here('プロセス管理の実行者', 'アクション「' + a.name + '」');
        });
      });
    }

    // 通知
    ((s.general && s.general.notifications) || []).forEach(function (n) {
      if (n.entity && n.entity.type === 'FIELD_ENTITY' && codes.indexOf(n.entity.code) >= 0) here('アプリの条件通知', named(n.entity.code) + 'に入っている人に通知しています');
    });
    [['perRecord', 'レコードの条件通知'], ['reminder', 'リマインダーの条件通知']].forEach(function (pair) {
      ((s[pair[0]] && s[pair[0]].notifications) || []).forEach(function (n, i) {
        var name = n.title ? '「' + n.title + '」' : (i + 1) + '番目';
        if (hit(n.filterCond).length) here(pair[1] + 'の条件', name + '：' + n.filterCond);
        if (n.timing && codes.indexOf(n.timing.code) >= 0) here(pair[1] + 'の日時', name + 'は、' + named(n.timing.code) + 'の日時を基準にしています');
        (n.targets || []).forEach(function (t) {
          if (t.entity && t.entity.type === 'FIELD_ENTITY' && codes.indexOf(t.entity.code) >= 0) here(pair[1] + 'の通知先', name + 'の通知先に' + named(t.entity.code) + 'を使っています');
        });
      });
    });

    // アクション（このアプリから、ほかのアプリへコピー）
    byIndex(s.actions && s.actions.actions).forEach(function (a) {
      (a.mappings || []).forEach(function (m) {
        if (m.srcType === 'FIELD' && codes.indexOf(m.srcField) >= 0) here('アクションのコピー元', 'アクション「' + a.name + '」が、' + named(m.srcField) + 'の値をアプリ' + ((a.destApp && a.destApp.app) || '') + 'にコピーしています');
      });
      if (hit(a.filterCond).length) here('アクションの条件', 'アクション「' + a.name + '」：' + a.filterCond);
    });

    // レコードのタイトル
    var title = s.appSettings && s.appSettings.titleField;
    if (title && title.selectionMode === 'MANUAL' && codes.indexOf(title.code) >= 0) here('レコードのタイトル', 'レコードのタイトルに' + named(title.code) + 'を使っています');

    // JavaScriptのカスタマイズ（文字列として出てくる場所。実際に使っているかは、中身を見て確かめる）
    (scripts || []).forEach(function (f) {
      codes.forEach(function (c) {
        if (mentionsInScript(f.text, c)) here('JavaScript（可能性）', f.screen + '「' + f.name + '」の中に、文字列 \'' + c + '\' があります');
      });
    });

    // ほかのアプリ
    (others || []).forEach(function (o) {
      if (String(o.appId) === String(target.appId)) return;
      var of = flatFields(o.fields && o.fields.properties);
      var there = function (place, detail, blocks) { add('ほかのアプリ', o.appId, o.name, place, detail, blocks); };
      var theirName = function (code) { return '「' + label(of, code) + '」'; };
      Object.keys(of).forEach(function (code) {
        var p = of[code].prop;
        var lk = p.lookup;
        if (lk && lk.relatedApp && String(lk.relatedApp.app) === String(target.appId)) {
          if (codes.indexOf(lk.relatedKeyField) >= 0) there('ルックアップのキー', 'ルックアップ' + theirName(code) + 'が、' + named(lk.relatedKeyField) + 'をキーにしています', true);
          (lk.fieldMappings || []).forEach(function (m) {
            if (codes.indexOf(m.relatedField) >= 0) there('ルックアップのコピー元', 'ルックアップ' + theirName(code) + 'が、' + named(m.relatedField) + 'の値を' + theirName(m.field) + 'にコピーしています', true);
          });
          if ((lk.lookupPickerFields || []).some(function (f) { return codes.indexOf(f) >= 0; })) there('ルックアップの選択画面', 'ルックアップ' + theirName(code) + 'の選択画面に表示しています');
          if (hit(lk.filterCond).length) there('ルックアップの絞り込み', 'ルックアップ' + theirName(code) + '：' + lk.filterCond);
          if (hit(lk.sort).length) there('ルックアップの並び順', 'ルックアップ' + theirName(code) + '：' + lk.sort);
        }
        var rt = p.referenceTable;
        if (rt && rt.relatedApp && String(rt.relatedApp.app) === String(target.appId)) {
          if (rt.condition && codes.indexOf(rt.condition.relatedField) >= 0) there('関連レコード一覧の条件', '関連レコード一覧' + theirName(code) + 'が、' + named(rt.condition.relatedField) + 'で絞り込んでいます', true);
          if ((rt.displayFields || []).some(function (f) { return codes.indexOf(f) >= 0; })) there('関連レコード一覧の表示項目', '関連レコード一覧' + theirName(code), true);
          if (hit(rt.filterCond).length) there('関連レコード一覧の絞り込み', '関連レコード一覧' + theirName(code) + '：' + rt.filterCond);
          if (hit(rt.sort).length) there('関連レコード一覧の並び順', '関連レコード一覧' + theirName(code) + '：' + rt.sort);
        }
      });
      byIndex(o.actions && o.actions.actions).forEach(function (a) {
        if (!a.destApp || String(a.destApp.app) !== String(target.appId)) return;
        (a.mappings || []).forEach(function (m) {
          if (codes.indexOf(m.destField) >= 0) there('アクションのコピー先', 'アクション「' + a.name + '」が、' + named(m.destField) + 'に値をコピーしています');
        });
      });
    });
    return out;
  }

  // 選べるフィールドの一覧（テーブルの中のフィールドは「テーブル名 > フィールド名」）
  function choices(properties) {
    var list = [];
    var skip = { __ID__: 0, __REVISION__: 0, STATUS: 0, STATUS_ASSIGNEE: 0, CATEGORY: 0, LABEL: 0, SPACER: 0, HR: 0, GROUP: 0 };
    Object.keys(properties || {}).forEach(function (code) {
      var p = properties[code];
      if (p.type in skip) return;
      list.push({ code: code, label: (p.label || code), type: p.type });
      if (p.type === 'SUBTABLE') Object.keys(p.fields || {}).forEach(function (inner) {
        list.push({ code: inner, label: (p.label || code) + ' > ' + (p.fields[inner].label || inner), type: p.fields[inner].type });
      });
    });
    return list;
  }

  var FieldRefs = { mentions: mentions, mentionsInScript: mentionsInScript, find: find, choices: choices, flatFields: flatFields };
  if (typeof module !== 'undefined' && module.exports) module.exports = FieldRefs;
  else root.DaichoyaFieldRefs = FieldRefs;
})(typeof window !== 'undefined' ? window : this);
