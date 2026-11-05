// 製品画像の撮影用：kintoneの代わりに、架空の会社の見本データを返す
(function () {
  var apps = [
    { appId: '11', name: '顧客マスタ', spaceId: '2', createdAt: '2024-04-01T00:30:00Z', creator: { code: 'kanri', name: '情報システム課' }, modifiedAt: '2026-09-12T02:10:00Z', modifier: { code: 'kanri', name: '情報システム課' } },
    { appId: '12', name: '案件管理', spaceId: '2', createdAt: '2025-04-01T00:30:00Z', creator: { code: 'kanri', name: '情報システム課' }, modifiedAt: '2026-10-28T06:12:00Z', modifier: { code: 'kanri', name: '情報システム課' } },
    { appId: '13', name: '対応履歴', spaceId: '2', createdAt: '2025-04-01T00:40:00Z', creator: { code: 'kanri', name: '情報システム課' }, modifiedAt: '2026-06-03T08:00:00Z', modifier: { code: 'sato', name: '佐藤' } },
    { appId: '14', name: '見積書', spaceId: '3', createdAt: '2023-10-02T01:00:00Z', creator: { code: 'kanri', name: '情報システム課' }, modifiedAt: '2025-02-14T05:30:00Z', modifier: { code: 'suzuki', name: '鈴木' } },
    { appId: '15', name: '請求管理', spaceId: '3', createdAt: '2023-10-02T01:10:00Z', creator: { code: 'kanri', name: '情報システム課' }, modifiedAt: '2026-08-20T03:00:00Z', modifier: { code: 'keiri', name: '経理課' } },
    { appId: '16', name: '商品マスタ', spaceId: '', createdAt: '2022-05-10T00:00:00Z', creator: { code: 'kanri', name: '情報システム課' }, modifiedAt: '2024-03-01T00:00:00Z', modifier: { code: 'kanri', name: '情報システム課' } }
  ];
  var lookup = function (app, key, maps) {
    return { relatedApp: { app: app, code: '' }, relatedKeyField: key, fieldMappings: maps || [], lookupPickerFields: [], filterCond: '', sort: '' };
  };
  var fields12 = {
    レコード番号: { type: 'RECORD_NUMBER', code: 'レコード番号', label: 'レコード番号' },
    作成者: { type: 'CREATOR', code: '作成者', label: '作成者' },
    作成日時: { type: 'CREATED_TIME', code: '作成日時', label: '作成日時' },
    更新者: { type: 'MODIFIER', code: '更新者', label: '更新者' },
    更新日時: { type: 'UPDATED_TIME', code: '更新日時', label: '更新日時' },
    案件名: { type: 'SINGLE_LINE_TEXT', code: '案件名', label: '案件名', required: true, unique: false, defaultValue: '', maxLength: '100', minLength: '' },
    顧客ID: { type: 'SINGLE_LINE_TEXT', code: '顧客ID', label: '顧客ID', required: true, lookup: lookup('11', '顧客ID', [{ field: '顧客名', relatedField: '会社名' }, { field: '業種', relatedField: '業種' }]) },
    顧客名: { type: 'SINGLE_LINE_TEXT', code: '顧客名', label: '顧客名' },
    業種: { type: 'SINGLE_LINE_TEXT', code: '業種', label: '業種' },
    受注確度: { type: 'DROP_DOWN', code: '受注確度', label: '受注確度', required: false, defaultValue: 'B', options: { A: { label: 'A', index: '0' }, B: { label: 'B', index: '1' }, C: { label: 'C', index: '2' }, 失注: { label: '失注', index: '3' } } },
    担当者: { type: 'USER_SELECT', code: '担当者', label: '担当者', required: true, defaultValue: [{ type: 'FUNCTION', code: 'LOGINUSER()' }] },
    受注予定日: { type: 'DATE', code: '受注予定日', label: '受注予定日' },
    明細: { type: 'SUBTABLE', code: '明細', label: '明細', fields: {
      商品コード: { type: 'SINGLE_LINE_TEXT', code: '商品コード', label: '商品コード', lookup: lookup('16', '商品コード', [{ field: '単価', relatedField: '標準単価' }]) },
      単価: { type: 'NUMBER', code: '単価', label: '単価', unit: '円' },
      数量: { type: 'NUMBER', code: '数量', label: '数量', defaultValue: '1' },
      小計: { type: 'CALC', code: '小計', label: '小計', expression: '単価*数量', unit: '円' }
    } },
    金額: { type: 'CALC', code: '金額', label: '金額', expression: 'SUM(小計)', unit: '円' },
    対応履歴: { type: 'REFERENCE_TABLE', code: '対応履歴', label: '対応履歴', referenceTable: { relatedApp: { app: '13', code: '' }, condition: { field: 'レコード番号', relatedField: '案件番号' }, filterCond: '', displayFields: ['対応日', '内容'], sort: '対応日 desc', size: '5' } }
  };
  var layout12 = { layout: [
    { type: 'ROW', fields: [{ type: 'RECORD_NUMBER', code: 'レコード番号' }, { type: 'SINGLE_LINE_TEXT', code: '案件名' }] },
    { type: 'ROW', fields: [{ type: 'SINGLE_LINE_TEXT', code: '顧客ID' }, { type: 'SINGLE_LINE_TEXT', code: '顧客名' }, { type: 'SINGLE_LINE_TEXT', code: '業種' }] },
    { type: 'ROW', fields: [{ type: 'DROP_DOWN', code: '受注確度' }, { type: 'USER_SELECT', code: '担当者' }, { type: 'DATE', code: '受注予定日' }] },
    { type: 'SUBTABLE', code: '明細', fields: [{ type: 'SINGLE_LINE_TEXT', code: '商品コード' }, { type: 'NUMBER', code: '単価' }, { type: 'NUMBER', code: '数量' }, { type: 'CALC', code: '小計' }] },
    { type: 'ROW', fields: [{ type: 'CALC', code: '金額' }] },
    { type: 'ROW', fields: [{ type: 'REFERENCE_TABLE', code: '対応履歴' }] }
  ] };
  var data = {
    '/k/v1/app/form/fields.json': {
      11: { properties: {
        顧客ID: { type: 'SINGLE_LINE_TEXT', code: '顧客ID', label: '顧客ID', unique: true },
        会社名: { type: 'SINGLE_LINE_TEXT', code: '会社名', label: '会社名' },
        案件一覧: { type: 'REFERENCE_TABLE', code: '案件一覧', label: 'この顧客の案件', referenceTable: { relatedApp: { app: '12', code: '' }, condition: { field: '顧客ID', relatedField: '顧客ID' }, filterCond: '' } }
      } },
      12: { properties: fields12 },
      13: { properties: { 案件番号: { type: 'NUMBER', code: '案件番号', label: '案件番号', lookup: lookup('12', 'レコード番号') } } },
      14: { properties: { 見積番号: { type: 'SINGLE_LINE_TEXT', code: '見積番号', label: '見積番号' } } },
      15: { properties: { 案件番号: { type: 'NUMBER', code: '案件番号', label: '案件番号', lookup: lookup('12', 'レコード番号') } } },
      16: { properties: { 商品コード: { type: 'SINGLE_LINE_TEXT', code: '商品コード', label: '商品コード', unique: true } } }
    },
    '/k/v1/app/actions.json': {
      12: { actions: { 対応を記録: { name: '対応を記録', id: '1', index: '0', destApp: { app: '13', code: '' }, mappings: [{ srcType: 'RECORD_URL', destField: '案件URL' }], entities: [], filterCond: '' } } },
      14: { actions: { 受注にする: { name: '受注にする', id: '2', index: '0', destApp: { app: '12', code: '' }, mappings: [{ srcType: 'FIELD', srcField: '見積番号', destField: '案件名' }], entities: [], filterCond: '' } } }
    },
    '/k/v1/app/form/layout.json': { 12: layout12 },
    '/k/v1/app/settings.json': { 12: { name: '案件管理', titleField: { selectionMode: 'MANUAL', code: '案件名' } } },
    '/k/v1/app/reports.json': { 12: { reports: { 担当者別の受注見込み: { name: '担当者別の受注見込み', index: '0', chartType: 'BAR', groups: [{ code: '担当者' }], aggregations: [{ type: 'SUM', code: '金額' }], filterCond: '受注確度 in ("A", "B")', sorts: [] } } } },
    '/k/v1/app.json': { 12: { appId: '12', name: '案件管理', description: '<div>営業部の案件を管理するアプリです。<br />受注確度は毎週月曜に見直します。</div>', createdAt: '2025-04-01T00:30:00Z', creator: { code: 'kanri', name: '情報システム課' }, modifiedAt: '2026-10-28T06:12:00Z', modifier: { code: 'kanri', name: '情報システム課' } } },
    '/k/v1/app/views.json': { 12: { views: {
      すべての案件: { type: 'LIST', name: 'すべての案件', index: '0', fields: ['案件名', '顧客名', '受注確度', '担当者', '金額'], filterCond: '', sort: 'レコード番号 desc' },
      今月の受注予定: { type: 'LIST', name: '今月の受注予定', index: '1', fields: ['案件名', '顧客名', '受注予定日', '金額'], filterCond: '受注予定日 = THIS_MONTH()', sort: '受注予定日 asc' },
      山本さんの担当: { type: 'LIST', name: '山本さんの担当', index: '2', fields: ['案件名', '顧客名', '受注確度', '金額'], filterCond: '担当者 in ("yamamoto")', sort: 'レコード番号 desc' }
    } } },
    '/k/v1/app/status.json': { 12: { enable: true, states: {
      商談中: { name: '商談中', index: '0', assignee: { type: 'ONE', entities: [{ entity: { type: 'FIELD_ENTITY', code: '担当者' } }] } },
      上長確認: { name: '上長確認', index: '1', assignee: { type: 'ONE', entities: [{ entity: { type: 'ORGANIZATION', code: '営業部' } }] } },
      受注: { name: '受注', index: '2', assignee: { type: 'ONE', entities: [] } }
    }, actions: [
      { name: '確認を依頼', from: '商談中', to: '上長確認', filterCond: '金額 > 0' },
      { name: '承認する', from: '上長確認', to: '受注', filterCond: '', executableUser: { entities: [{ entity: { type: 'USER', code: 'kimura' }, includeSubs: false }, { entity: { type: 'ORGANIZATION', code: '営業部' }, includeSubs: false }] } },
      { name: '差し戻す', from: '上長確認', to: '商談中', filterCond: '' }
    ] } },
    '/k/v1/app/notifications/general.json': { 12: { notifications: [{ entity: { type: 'FIELD_ENTITY', code: '担当者' }, includeSubs: false, recordAdded: false, recordEdited: true, commentAdded: true, fileImported: false }], notifyToCommenter: true } },
    '/k/v1/app/notifications/perRecord.json': { 12: { notifications: [{ filterCond: '金額 >= 1000000', title: '100万円以上の案件が登録されました', targets: [{ entity: { type: 'ORGANIZATION', code: '営業部' }, includeSubs: true }] }] } },
    '/k/v1/app/notifications/reminder.json': { 12: { notifications: [{ timing: { code: '受注予定日', daysLater: -3, time: '09:00' }, filterCond: '受注確度 not in ("失注")', title: '受注予定日の3日前です', targets: [{ entity: { type: 'FIELD_ENTITY', code: '担当者' } }] }] } },
    '/k/v1/app/acl.json': {
      14: { rights: [
        { entity: { type: 'USER', code: 'suzuki' }, includeSubs: false, appEditable: true, recordViewable: true, recordAddable: true, recordEditable: true, recordDeletable: true, recordImportable: true, recordExportable: true },
        { entity: { type: 'GROUP', code: 'everyone' }, includeSubs: false, appEditable: true, recordViewable: true, recordAddable: true, recordEditable: true, recordDeletable: false, recordImportable: false, recordExportable: true }
      ] },
      15: { rights: [
        { entity: { type: 'ORGANIZATION', code: '経理課' }, includeSubs: false, appEditable: true, recordViewable: true, recordAddable: true, recordEditable: true, recordDeletable: true, recordImportable: true, recordExportable: true },
        { entity: { type: 'USER', code: 'tanaka' }, includeSubs: false, appEditable: false, recordViewable: true, recordAddable: false, recordEditable: false, recordDeletable: false, recordImportable: false, recordExportable: false }
      ] },
      12: { rights: [
      { entity: { type: 'GROUP', code: 'Administrators' }, includeSubs: false, appEditable: true, recordViewable: true, recordAddable: true, recordEditable: true, recordDeletable: true, recordImportable: true, recordExportable: true },
      { entity: { type: 'ORGANIZATION', code: '営業部' }, includeSubs: true, appEditable: false, recordViewable: true, recordAddable: true, recordEditable: true, recordDeletable: false, recordImportable: false, recordExportable: true },
      { entity: { type: 'CREATOR' }, includeSubs: false, appEditable: false, recordViewable: true, recordAddable: true, recordEditable: true, recordDeletable: true, recordImportable: false, recordExportable: false },
      { entity: { type: 'GROUP', code: 'everyone' }, includeSubs: false, appEditable: false, recordViewable: true, recordAddable: false, recordEditable: false, recordDeletable: false, recordImportable: false, recordExportable: false }
    ] } },
    '/k/v1/record/acl.json': { 12: { rights: [
      { filterCond: '受注確度 in ("失注")', entities: [
        { entity: { type: 'ORGANIZATION', code: '営業部' }, viewable: true, editable: false, deletable: false, includeSubs: true },
        { entity: { type: 'GROUP', code: 'everyone' }, viewable: false, editable: false, deletable: false, includeSubs: false }
      ] }
    ] } },
    '/k/v1/field/acl.json': { 12: { rights: [
      { code: '金額', entities: [
        { accessibility: 'WRITE', entity: { type: 'ORGANIZATION', code: '経理部' }, includeSubs: false },
        { accessibility: 'READ', entity: { type: 'ORGANIZATION', code: '営業部' }, includeSubs: true },
        { accessibility: 'NONE', entity: { type: 'GROUP', code: 'everyone' }, includeSubs: false }
      ] }
    ] } },
    '/k/v1/app/plugins.json': { 12: { plugins: [
      { id: 'pnihnmdhjnomgnnfhibjgpkefifmpbbi', name: '台帳屋 設計書ジェネレーター', enabled: true },
      { id: 'gnncoaohinigebkngpnpekmpapfnebfn', name: '台帳屋 権限チェッカー', enabled: true },
      { id: 'oengaljieegejcfdbmbkijobeodheoah', name: '台帳屋 アプリ相関マップ', enabled: true }
    ] } },
    '/k/v1/app/customize.json': { 12: { scope: 'ALL', desktop: {
      js: [{ type: 'FILE', file: { name: '金額の入力チェック.js', fileKey: 'demo1', contentType: 'text/javascript', size: '2048' } }],
      css: [{ type: 'FILE', file: { name: '一覧の見た目.css', fileKey: 'demo2', contentType: 'text/css', size: '512' } }]
    }, mobile: { js: [], css: [] }, revision: '5' } }
  };
  // kintoneのアプリには、更新日時のフィールドが必ずある
  Object.keys(data['/k/v1/app/form/fields.json']).forEach(function (id) {
    var props = data['/k/v1/app/form/fields.json'][id].properties;
    if (!props.更新日時) props.更新日時 = { type: 'UPDATED_TIME', code: '更新日時', label: '更新日時' };
  });
  // cybozu.com のユーザー・組織・グループ（User API の見本）。木村さんは利用停止中、山本さんと「経理部」はもういない
  var directory = {
    users: [
      { code: 'kanri', name: '情報システム課', valid: true }, { code: 'sato', name: '佐藤', valid: true }, { code: 'suzuki', name: '鈴木', valid: true },
      { code: 'tanaka', name: '田中', valid: true }, { code: 'keiri', name: '経理課', valid: true }, { code: 'kimura', name: '木村', valid: false }
    ],
    organizations: [{ code: '営業部', name: '営業部' }, { code: '経理課', name: '経理課' }, { code: '情報システム課', name: '情報システム課' }],
    groups: [{ code: 'Administrators', name: 'Administrators' }, { code: 'everyone', name: 'Everyone' }]
  };
  // 見本データのないアプリには、設定が空の応答を返す
  function emptyFor(path, appId) {
    var app = apps.filter(function (a) { return a.appId === String(appId); })[0] || { appId: String(appId), name: 'アプリ' };
    switch (path) {
      case '/k/v1/app.json': return { appId: app.appId, name: app.name, description: '', createdAt: '2025-04-01T00:30:00Z', creator: { code: 'kanri', name: '情報システム課' }, modifiedAt: '2026-10-01T00:00:00Z', modifier: { code: 'kanri', name: '情報システム課' } };
      case '/k/v1/app/form/fields.json': return { properties: { 更新日時: { type: 'UPDATED_TIME', code: '更新日時', label: '更新日時' } } };
      case '/k/v1/app/form/layout.json': return { layout: [] };
      case '/k/v1/app/views.json': return { views: {} };
      case '/k/v1/app/status.json': return { enable: false, states: {}, actions: [] };
      case '/k/v1/app/notifications/general.json': return { notifications: [], notifyToCommenter: false };
      case '/k/v1/app/notifications/perRecord.json':
      case '/k/v1/app/notifications/reminder.json': return { notifications: [] };
      case '/k/v1/app/acl.json':
      case '/k/v1/record/acl.json':
      case '/k/v1/field/acl.json': return { rights: [] };
      case '/k/v1/app/plugins.json': return { plugins: [] };
      case '/k/v1/app/actions.json': return { actions: {} };
      case '/k/v1/app/reports.json': return { reports: {} };
      case '/k/v1/app/settings.json': return { name: app.name, titleField: { selectionMode: 'AUTO' } };
      case '/k/v1/records.json': return null;
      case '/k/v1/app/customize.json': return { scope: 'ALL', desktop: { js: [], css: [] }, mobile: { js: [], css: [] } };
      default: return null;
    }
  }
  // 複数のプラグインを同じ画面で動かせるよう、イベントの処理は配列で持つ
  var handlers = {};
  var space = document.getElementById('header-space');
  window.kintone = {
    events: { on: function (name, fn) { (handlers[name] = handlers[name] || []).push(fn); } },
    app: { getId: function () { return 12; }, getHeaderMenuSpaceElement: function () { return space; } },
    getLoginUser: function () { return { code: 'kanri' }; },
    // 有料版（棚卸し・変更監視パック）のデモ用：プラグインの設定（ライセンスキー）を返す
    $PLUGIN_ID: 'demo',
    plugin: { app: { getConfig: function () { return window.__dcyConfig || {}; } } },
    api: function (path, method, params) {
      return new Promise(function (resolve, reject) {
        setTimeout(function () {
          if (path === '/k/v1/apps.json') { resolve({ apps: apps }); return; }
          if (/^\/v1\//.test(path)) {
            var key = path.replace('/v1/', '').replace('.json', '');
            var out = {};
            out[key] = (directory[key] || []).filter(function (x) { return !params.codes || params.codes.indexOf(x.code) >= 0; });
            resolve(JSON.parse(JSON.stringify(out)));
            return;
          }
          if (path === '/k/v1/records.json') {
            var activity = { 11: ['2026-11-02T09:00:00Z', 318], 12: ['2026-11-04T08:30:00Z', 105], 13: ['2026-11-04T05:10:00Z', 1240], 14: ['2025-12-18T02:00:00Z', 57], 15: ['2026-10-31T07:00:00Z', 96], 16: ['2026-01-20T00:00:00Z', 24] }[String(params.app)];
            var rec = {};
            if (activity) rec[(params.fields || ['更新日時'])[0]] = { type: 'UPDATED_TIME', value: activity[0] };
            resolve({ records: activity ? [rec] : [], totalCount: String(activity ? activity[1] : 0) });
            return;
          }
          var byApp = data[path];
          var appId = String(params.app || params.id);   // app.json だけは id で指定する
          var value = (byApp && byApp[appId]) || emptyFor(path, appId);
          if (value) resolve(JSON.parse(JSON.stringify(value)));
          else reject({ code: 'GAIA_NO01', message: 'not found' });
        }, 20);
      });
    }
  };
  kintone.api.url = function (path) { return path; };
  window.__fire = function () { (handlers['app.record.index.show'] || []).forEach(function (fn) { fn({}); }); };
})();
