/*
 * 台帳屋 共通部品：アプリの設定に出てくるユーザー・組織・グループを集める
 * 異動・退職チェッカーと、棚卸し・変更監視パックで使う。読み取るのは設定だけで、レコードの中身は読まない。
 */
(function (root) {
  'use strict';

  var TYPE_LABELS = { USER: 'ユーザー', ORGANIZATION: '組織', GROUP: 'グループ' };
  // 条件の文字列の中で、ユーザー・組織・グループを指定できるフィールド
  var FIELD_ENTITY_TYPES = {
    USER_SELECT: 'USER', CREATOR: 'USER', MODIFIER: 'USER', STATUS_ASSIGNEE: 'USER',
    ORGANIZATION_SELECT: 'ORGANIZATION', GROUP_SELECT: 'GROUP'
  };
  var APP_RIGHTS = [['appEditable', 'アプリ管理'], ['recordViewable', '閲覧'], ['recordAddable', '追加'], ['recordEditable', '編集'],
    ['recordDeletable', '削除'], ['recordImportable', 'ファイル読み込み'], ['recordExportable', 'ファイル書き出し']];
  var ACCESS = { WRITE: '閲覧・編集', READ: '閲覧のみ', NONE: '権限なし' };
  var GENERAL_EVENTS = [['recordAdded', 'レコード追加'], ['recordEdited', 'レコード編集'], ['commentAdded', 'コメント書き込み'],
    ['statusChanged', 'ステータスの更新'], ['fileImported', 'ファイル読み込み']];

  // テーブルの中のフィールドも含めて、フィールドコード→設定 の対応表を作る
  function flatFields(properties) {
    var map = {};
    Object.keys(properties || {}).forEach(function (code) {
      var p = properties[code];
      map[code] = p;
      if (p.type === 'SUBTABLE' && p.fields) Object.keys(p.fields).forEach(function (inner) { map[inner] = p.fields[inner]; });
    });
    return map;
  }

  function fieldLabel(fields, code) {
    var p = fields[code];
    return (p && p.label) || code;
  }

  function isTarget(entity) {
    return !!(entity && TYPE_LABELS[entity.type] && entity.code);
  }

  // 条件の文字列から、ユーザー・組織・グループの指定を取り出す
  // 例：担当者 in ("sato", LOGINUSER()) → [{ type: 'USER', code: 'sato' }]
  function condRefs(cond, fields) {
    var out = [];
    if (!cond) return out;
    var re = /([^\s()"=!<>,]+)\s+(?:not\s+in|in)\s*\(((?:[^()"]|"(?:[^"\\]|\\.)*"|\(\))*)\)/g;
    var m;
    while ((m = re.exec(cond))) {
      var field = fields[m[1]];
      var type = field && FIELD_ENTITY_TYPES[field.type];
      if (!type) continue;
      var vre = /"((?:[^"\\]|\\.)*)"/g;
      var v;
      while ((v = vre.exec(m[2]))) {
        var code = v[1].replace(/\\(.)/g, '$1');
        if (code && !/\(\)$/.test(code)) out.push({ type: type, code: code, field: m[1] });
      }
    }
    return out;
  }

  function rightsText(right, pairs) {
    var list = pairs.filter(function (p) { return right[p[0]]; }).map(function (p) { return p[1]; });
    return list.length ? list.join('・') : '権限なし';
  }

  function byIndex(obj) {
    return Object.keys(obj || {}).map(function (k) { return obj[k]; })
      .sort(function (a, b) { return Number(a.index || 0) - Number(b.index || 0); });
  }

  // settings: アプリの設定のAPIの応答（読めなかったものは null）
  //   { appAcl, recordAcl, fieldAcl, status, general, perRecord, reminder, fields, views, reports, actions }
  // 戻り値：[{ type, code, place, detail }]
  function collect(settings) {
    var s = settings || {};
    var refs = [];
    var seen = {};
    var fields = flatFields(s.fields && s.fields.properties);
    function add(entity, place, detail) {
      if (!isTarget(entity)) return;
      var key = entity.type + '\n' + entity.code + '\n' + place + '\n' + detail;
      if (seen[key]) return;
      seen[key] = true;
      refs.push({ type: entity.type, code: entity.code, place: place, detail: detail || '' });
    }
    function addCond(cond, place, detail) {
      condRefs(cond, fields).forEach(function (r) {
        add(r, place, detail + '（フィールド「' + fieldLabel(fields, r.field) + '」）：' + cond);
      });
    }
    var subs = function (item) { return item && item.includeSubs ? '（下位組織を含む）' : ''; };

    ((s.appAcl && s.appAcl.rights) || []).forEach(function (r, i) {
      add(r.entity, 'アプリのアクセス権', (i + 1) + '番目：' + rightsText(r, APP_RIGHTS) + subs(r));
    });
    ((s.recordAcl && s.recordAcl.rights) || []).forEach(function (rule, i) {
      var cond = rule.filterCond || 'すべてのレコード';
      (rule.entities || []).forEach(function (e) {
        add(e.entity, 'レコードのアクセス権', (i + 1) + '番目の条件「' + cond + '」：' + rightsText(e, [['viewable', '閲覧'], ['editable', '編集'], ['deletable', '削除']]) + subs(e));
      });
      addCond(rule.filterCond, 'レコードのアクセス権の条件', (i + 1) + '番目の条件');
    });
    ((s.fieldAcl && s.fieldAcl.rights) || []).forEach(function (rule) {
      (rule.entities || []).forEach(function (e) {
        add(e.entity, 'フィールドのアクセス権', 'フィールド「' + fieldLabel(fields, rule.code) + '」：' + (ACCESS[e.accessibility] || e.accessibility || '') + subs(e));
      });
    });
    if (s.status && s.status.enable) {
      byIndex(s.status.states).forEach(function (st) {
        ((st.assignee && st.assignee.entities) || []).forEach(function (e) {
          add(e.entity, 'プロセス管理の作業者', 'ステータス「' + st.name + '」' + subs(e));
        });
      });
      (s.status.actions || []).forEach(function (a) {
        ((a.executableUser && a.executableUser.entities) || []).forEach(function (e) {
          add(e.entity, 'プロセス管理の実行者', 'アクション「' + a.name + '」' + subs(e));
        });
        addCond(a.filterCond, 'プロセス管理の条件', 'アクション「' + a.name + '」の条件');
      });
    }
    ((s.general && s.general.notifications) || []).forEach(function (n) {
      add(n.entity, 'アプリの条件通知', rightsText(n, GENERAL_EVENTS).replace('権限なし', '通知なし') + subs(n));
    });
    [['perRecord', 'レコードの条件通知'], ['reminder', 'リマインダーの条件通知']].forEach(function (pair) {
      ((s[pair[0]] && s[pair[0]].notifications) || []).forEach(function (n, i) {
        var name = n.title ? '「' + n.title + '」' : (i + 1) + '番目';
        (n.targets || []).forEach(function (t) { add(t.entity, pair[1], name + 'の通知先' + subs(t)); });
        addCond(n.filterCond, pair[1] + 'の条件', name + 'の条件');
      });
    });
    Object.keys(fields).forEach(function (code) {
      var p = fields[code];
      if (!/^(USER|ORGANIZATION|GROUP)_SELECT$/.test(p.type)) return;
      (p.entities || []).forEach(function (e) { add(e, 'フィールドの選択肢', 'フィールド「' + fieldLabel(fields, code) + '」で選べる候補'); });
      (p.defaultValue || []).forEach(function (e) { if (e && e.type !== 'FUNCTION') add(e, 'フィールドの初期値', 'フィールド「' + fieldLabel(fields, code) + '」'); });
    });
    byIndex(s.views && s.views.views).forEach(function (v) { addCond(v.filterCond, '一覧の絞り込み', '一覧「' + v.name + '」'); });
    byIndex(s.reports && s.reports.reports).forEach(function (r) { addCond(r.filterCond, 'グラフの絞り込み', 'グラフ「' + r.name + '」'); });
    byIndex(s.actions && s.actions.actions).forEach(function (a) {
      (a.entities || []).forEach(function (e) { add(e.entity || e, 'アクションを使えるユーザー', 'アクション「' + a.name + '」'); });
    });
    return refs;
  }

  // アプリ管理の権限が、ユーザー1人だけに付いているか（組織・グループでの指定がない）
  function soleManager(appAcl) {
    var managers = ((appAcl && appAcl.rights) || []).filter(function (r) { return r.appEditable && r.entity; });
    if (managers.length === 1 && managers[0].entity.type === 'USER') return managers[0].entity.code;
    return null;
  }

  // 出てきたユーザー・組織・グループが、いまもあるかを cybozu.com の User API で確かめる
  // get(path, params) は Promise を返す関数（kintone.api など）。100件ずつ問い合わせる
  function lookup(get, refs) {
    var codes = { USER: [], ORGANIZATION: [], GROUP: [] };
    refs.forEach(function (r) {
      if (r.type === 'USER' && /^guest\//.test(r.code)) return;
      if (codes[r.type] && codes[r.type].indexOf(r.code) < 0) codes[r.type].push(r.code);
    });
    var paths = { USER: ['/v1/users.json', 'users'], ORGANIZATION: ['/v1/organizations.json', 'organizations'], GROUP: ['/v1/groups.json', 'groups'] };
    var dir = { USER: {}, ORGANIZATION: {}, GROUP: {}, failed: {} };
    var jobs = [];
    Object.keys(codes).forEach(function (type) {
      for (var i = 0; i < codes[type].length; i += 100) {
        jobs.push(get(paths[type][0], { codes: codes[type].slice(i, i + 100) }).then(function (resp) {
          ((resp && resp[paths[type][1]]) || []).forEach(function (x) { dir[type][x.code] = { name: x.name || x.code, valid: x.valid !== false }; });
        }, function () { dir.failed[type] = true; }));
      }
    });
    return Promise.all(jobs).then(function () { return dir; });
  }

  // level 2：見つからない、1：利用停止中、0：問題なし（または確かめられない）
  function stateOf(ref, dir) {
    if (ref.type === 'USER' && /^guest\//.test(ref.code)) return { level: 0, text: 'ゲスト（確認していません）', name: '' };
    if (!dir || dir.failed[ref.type]) return { level: 0, text: '確認できませんでした', name: '' };
    var x = dir[ref.type][ref.code];
    if (!x) return { level: 2, text: '見つかりません', name: '' };
    if (ref.type === 'USER' && !x.valid) return { level: 1, text: '利用停止中', name: x.name };
    return { level: 0, text: '有効', name: x.name };
  }

  // 表示用：問題のあるものを先に、同じ人の行をまとめて並べる
  function rows(refs, dir) {
    var order = { USER: 0, ORGANIZATION: 1, GROUP: 2 };
    return refs.map(function (r) {
      var st = stateOf(r, dir);
      return { level: st.level, state: st.text, type: r.type, typeLabel: TYPE_LABELS[r.type], code: r.code, name: st.name, place: r.place, detail: r.detail };
    }).sort(function (a, b) {
      return (b.level - a.level) || (order[a.type] - order[b.type]) || (a.code < b.code ? -1 : a.code > b.code ? 1 : 0);
    });
  }

  function summary(list) {
    var who = {};
    var counts = { USER: 0, ORGANIZATION: 0, GROUP: 0 };
    var missing = {};
    var stopped = {};
    list.forEach(function (r) {
      var key = r.type + '\n' + r.code;
      if (!who[key]) { who[key] = true; counts[r.type] += 1; }
      if (r.level === 2) missing[key] = true;
      if (r.level === 1) stopped[key] = true;
    });
    return { users: counts.USER, organizations: counts.ORGANIZATION, groups: counts.GROUP, missing: Object.keys(missing).length, stopped: Object.keys(stopped).length };
  }

  var Refs = { TYPE_LABELS: TYPE_LABELS, flatFields: flatFields, condRefs: condRefs, collect: collect, soleManager: soleManager, lookup: lookup, stateOf: stateOf, rows: rows, summary: summary };
  if (typeof module !== 'undefined' && module.exports) module.exports = Refs;
  else root.DaichoyaRefs = Refs;
})(typeof window !== 'undefined' ? window : this);
