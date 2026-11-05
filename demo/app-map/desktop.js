/*
 * 台帳屋 アプリ相関マップ for kintone
 * 一覧画面に「相関マップ」ボタンを追加し、選んだアプリとつながっているアプリ
 * （ルックアップ・関連レコード一覧・アクション）を図で表示する。図は画像（PNG）、
 * つながりの一覧はCSVで保存できる。データはすべてお客様のkintoneの中で処理し、外部には送らない。
 */
(function () {
  'use strict';

  // ===== 1. つながりの抽出（kintoneに依存しない純粋な関数。テスト対象） =====

  var KIND_LABELS = { LOOKUP: 'ルックアップ', REFERENCE: '関連レコード一覧', ACTION: 'アクション' };

  function kindLabel(kind) { return KIND_LABELS[kind] || kind; }

  function fieldLabel(properties, code) {
    if (!code) return '';
    var direct = properties[code];
    if (direct) return direct.label || code;
    var found = '';
    Object.keys(properties).forEach(function (key) {
      var prop = properties[key];
      if (!found && prop.type === 'SUBTABLE' && prop.fields && prop.fields[code]) found = prop.fields[code].label || code;
    });
    return found || code;
  }

  function relatedAppId(relatedApp) {
    return relatedApp && relatedApp.app ? String(relatedApp.app) : null;
  }

  function lookupLink(appId, prop, properties, tableLabel) {
    var lookup = prop.lookup;
    var place = (prop.label || prop.code) + (tableLabel ? '（テーブル「' + tableLabel + '」内）' : '');
    if (lookup === null) {
      return { from: appId, to: null, kind: 'LOOKUP', place: place, detail: '参照先のアプリを見る権限がないため、参照先が分かりません' };
    }
    var copies = (lookup.fieldMappings || []).map(function (m) {
      return m.relatedField + '→' + fieldLabel(properties, m.field);
    });
    var detail = 'キー：' + (lookup.relatedKeyField || '') + '→' + (prop.label || prop.code);
    if (copies.length) detail += '／コピー：' + copies.join('、');
    if (lookup.filterCond) detail += '／絞り込み：' + lookup.filterCond;
    return { from: appId, to: relatedAppId(lookup.relatedApp), kind: 'LOOKUP', place: place, detail: detail };
  }

  function referenceLink(appId, prop, properties) {
    var ref = prop.referenceTable;
    var place = prop.label || prop.code;
    if (ref === null) {
      return { from: appId, to: null, kind: 'REFERENCE', place: place, detail: '参照先のアプリを見る権限がないため、参照先が分かりません' };
    }
    var cond = ref.condition || {};
    var detail = '条件：' + fieldLabel(properties, cond.field) + ' ＝ 相手の ' + (cond.relatedField || '');
    if (ref.filterCond) detail += '／絞り込み：' + ref.filterCond;
    return { from: appId, to: relatedAppId(ref.relatedApp), kind: 'REFERENCE', place: place, detail: detail };
  }

  // properties: フォームのフィールド設定、actions: アクション設定（取得できなければ null）
  function extractLinks(appId, properties, actions) {
    var id = String(appId);
    var links = [];
    var props = properties || {};
    Object.keys(props).forEach(function (code) {
      var prop = props[code];
      if (prop.lookup !== undefined) links.push(lookupLink(id, prop, props, null));
      if (prop.type === 'REFERENCE_TABLE') links.push(referenceLink(id, prop, props));
      if (prop.type === 'SUBTABLE' && prop.fields) {
        Object.keys(prop.fields).forEach(function (inner) {
          var child = prop.fields[inner];
          if (child.lookup !== undefined) links.push(lookupLink(id, child, props, prop.label || prop.code));
        });
      }
    });
    var list = Object.keys((actions && actions.actions) || {}).map(function (key) { return actions.actions[key]; });
    list.sort(function (a, b) { return Number(a.index) - Number(b.index); });
    list.forEach(function (action) {
      if (!action.destApp) {
        links.push({ from: id, to: null, kind: 'ACTION', place: action.name, detail: 'コピー先のアプリを見る権限がないため、コピー先が分かりません' });
        return;
      }
      var copies = (action.mappings || []).map(function (m) {
        var src = m.srcType === 'RECORD_URL' ? 'レコードのURL' : fieldLabel(props, m.srcField);
        return src + '→' + m.destField;
      });
      links.push({
        from: id,
        to: String(action.destApp.app),
        kind: 'ACTION',
        place: action.name,
        detail: copies.length ? 'コピー：' + copies.join('、') : 'コピーするフィールドなし'
      });
    });
    return links;
  }

  // 中心のアプリから見たつながりを、相手のアプリごとにまとめる
  function neighborhood(links, centerId) {
    var center = String(centerId);
    var byApp = {};
    var self = [];
    var unknown = [];
    function entry(app) {
      if (!byApp[app]) byApp[app] = { app: app, outgoing: [], incoming: [] };
      return byApp[app];
    }
    links.forEach(function (link) {
      if (link.from === center && link.to === null) { unknown.push(link); return; }
      if (link.from === center && link.to === center) { self.push(link); return; }
      if (link.from === center) entry(link.to).outgoing.push(link);
      else if (link.to === center) entry(link.from).incoming.push(link);
    });
    var apps = Object.keys(byApp).map(function (k) { return byApp[k]; });
    return {
      center: center,
      incomingOnly: apps.filter(function (a) { return a.incoming.length && !a.outgoing.length; }),
      outgoing: apps.filter(function (a) { return a.outgoing.length; }),
      self: self,
      unknown: unknown
    };
  }

  // 図の線に添える短いラベル（種類ごとに1行）
  var SHORT_LABELS = { LOOKUP: 'ルックアップ', REFERENCE: '関連レコード', ACTION: 'アクション' };

  function edgeLines(links) {
    var counts = {};
    var order = [];
    links.forEach(function (l) {
      if (!counts[l.kind]) { counts[l.kind] = 0; order.push(l.kind); }
      counts[l.kind] += 1;
    });
    return order.map(function (k) { return (SHORT_LABELS[k] || k) + (counts[k] > 1 ? '×' + counts[k] : ''); });
  }

  // 漢字の名前は読み順に並べられないため、管理者が見慣れたアプリIDの順に並べる
  function byAppId(a, b) {
    return Number(a.app) - Number(b.app);
  }

  // ===== 2. 図の配置とSVG =====

  var NODE_W = 220;
  var NODE_H = 48;
  var GAP_Y = 28;
  var COL_X = [24, 404, 784];
  var MAP_W = 1028;
  var TOP = 56;

  function layoutMap(hood, names) {
    var left = hood.incomingOnly.slice().sort(byAppId);
    var right = hood.outgoing.slice().sort(byAppId);
    var rows = Math.max(left.length, right.length, 1);
    var height = TOP + rows * NODE_H + (rows - 1) * GAP_Y + 32;
    function columnY(count, i) {
      var block = count * NODE_H + (count - 1) * GAP_Y;
      var start = TOP + (rows * NODE_H + (rows - 1) * GAP_Y - block) / 2;
      return start + i * (NODE_H + GAP_Y);
    }
    var centerNode = { app: hood.center, x: COL_X[1], y: columnY(1, 0), role: 'center' };
    var nodes = [centerNode];
    var edges = [];
    left.forEach(function (item, i) {
      var node = { app: item.app, x: COL_X[0], y: columnY(left.length, i), role: 'incoming' };
      nodes.push(node);
      edges.push({ from: node, to: centerNode, lines: edgeLines(item.incoming), offset: 0 });
    });
    right.forEach(function (item, i) {
      var node = { app: item.app, x: COL_X[2], y: columnY(right.length, i), role: 'outgoing' };
      nodes.push(node);
      var both = item.incoming.length > 0;
      edges.push({ from: centerNode, to: node, lines: edgeLines(item.outgoing), offset: both ? -9 : 0 });
      if (both) edges.push({ from: node, to: centerNode, lines: edgeLines(item.incoming), offset: 9 });
    });
    return {
      width: MAP_W, height: height + (hood.self.length ? 20 : 0), nodes: nodes, edges: edges,
      leftCount: left.length, rightCount: right.length, centerY: centerNode.y, selfCount: hood.self.length
    };
  }

  function xmlEscape(value) {
    return String(value).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  }

  function truncate(text, max) {
    var chars = Array.from(String(text));
    return chars.length > max ? chars.slice(0, max - 1).join('') + '…' : chars.join('');
  }

  var COLORS = {
    center: { fill: '#3498db', stroke: '#217dbb', text: '#ffffff' },
    incoming: { fill: '#ffffff', stroke: '#8fa3ad', text: '#333333' },
    outgoing: { fill: '#ffffff', stroke: '#8fa3ad', text: '#333333' }
  };

  function edgePoints(edge) {
    var a = edge.from;
    var b = edge.to;
    var leftToRight = a.x < b.x;
    var x1 = leftToRight ? a.x + NODE_W : a.x;
    var x2 = leftToRight ? b.x : b.x + NODE_W;
    return { x1: x1, y1: a.y + NODE_H / 2 + edge.offset, x2: x2, y2: b.y + NODE_H / 2 + edge.offset };
  }

  function buildSvg(layout, names, title) {
    var font = "Meiryo, 'Hiragino Sans', 'Yu Gothic', sans-serif";
    var out = [];
    out.push('<svg xmlns="http://www.w3.org/2000/svg" width="' + layout.width + '" height="' + layout.height +
      '" viewBox="0 0 ' + layout.width + ' ' + layout.height + '" font-family="' + xmlEscape(font) + '">');
    out.push('<defs><marker id="dcy-am-arrow" viewBox="0 0 10 10" refX="10" refY="5" markerWidth="8" markerHeight="8" orient="auto-start-reverse">' +
      '<path d="M0,0 L10,5 L0,10 z" fill="#5d7480"/></marker></defs>');
    out.push('<rect x="0" y="0" width="' + layout.width + '" height="' + layout.height + '" fill="#ffffff"/>');
    out.push('<text x="24" y="28" font-size="15" font-weight="bold" fill="#333333">' + xmlEscape(title) + '</text>');
    out.push('<text x="' + (COL_X[0] + NODE_W / 2) + '" y="48" font-size="12" text-anchor="middle" fill="#777777">このアプリを参照しているアプリ</text>');
    out.push('<text x="' + (COL_X[2] + NODE_W / 2) + '" y="48" font-size="12" text-anchor="middle" fill="#777777">このアプリが参照しているアプリ</text>');
    [[layout.leftCount, COL_X[0]], [layout.rightCount, COL_X[2]]].forEach(function (col) {
      if (col[0] === 0 && layout.edges.length > 0) {
        out.push('<text x="' + (col[1] + NODE_W / 2) + '" y="' + (layout.centerY + NODE_H / 2 + 4) + '" font-size="12" text-anchor="middle" fill="#aaaaaa">（なし）</text>');
      }
    });
    layout.edges.forEach(function (edge) {
      var p = edgePoints(edge);
      out.push('<line x1="' + p.x1 + '" y1="' + p.y1 + '" x2="' + p.x2 + '" y2="' + p.y2 +
        '" stroke="#5d7480" stroke-width="1.5" marker-end="url(#dcy-am-arrow)"/>');
      // 線は中心のアプリに集まるため、ラベルは相手のアプリの手前に寄せて重なりを避ける。
      // 上へ向かう線は線の上、下へ向かう線は線の下に置くと、文字と線が交わらない。
      var fromCenter = edge.from.role === 'center';
      var nx = fromCenter ? p.x2 : p.x1;
      var ny = fromCenter ? p.y2 : p.y1;
      var cy = fromCenter ? p.y1 : p.y2;
      var below = edge.offset > 0 || (edge.offset === 0 && ny > cy);
      var rightSide = edge.from.role === 'outgoing' || edge.to.role === 'outgoing';
      var tx = rightSide ? nx - 14 : nx + 14;
      // 線の上に置くときは上へ、下に置くときは下へ行を積む
      var lines = below ? edge.lines : edge.lines.slice().reverse();
      lines.forEach(function (line, i) {
        var text = truncate(line, 10);
        var w = Array.from(text).length * 11.5 + 8;
        var ty = below ? ny + 14 + i * 15 : ny - 6 - i * 15;
        var rx = rightSide ? tx - w + 4 : tx - 4;
        out.push('<rect x="' + rx + '" y="' + (ty - 12) + '" width="' + w + '" height="15" rx="3" fill="#ffffff" fill-opacity="0.9"/>');
        out.push('<text x="' + tx + '" y="' + ty + '" font-size="11" text-anchor="' + (rightSide ? 'end' : 'start') + '" fill="#3d5561">' + xmlEscape(text) + '</text>');
      });
    });
    layout.nodes.forEach(function (node) {
      var c = COLORS[node.role];
      var name = names[node.app] || ('アプリID ' + node.app);
      out.push('<g class="dcy-am-node" data-app="' + xmlEscape(node.app) + '">');
      out.push('<title>' + xmlEscape(name + '（ID: ' + node.app + '）') + '</title>');
      out.push('<rect x="' + node.x + '" y="' + node.y + '" width="' + NODE_W + '" height="' + NODE_H +
        '" rx="6" fill="' + c.fill + '" stroke="' + c.stroke + '" stroke-width="1.5"/>');
      out.push('<text x="' + (node.x + NODE_W / 2) + '" y="' + (node.y + 21) + '" font-size="13" font-weight="bold" text-anchor="middle" fill="' + c.text + '">' +
        xmlEscape(truncate(name, 16)) + '</text>');
      out.push('<text x="' + (node.x + NODE_W / 2) + '" y="' + (node.y + 38) + '" font-size="11" text-anchor="middle" fill="' + c.text + '">ID: ' + xmlEscape(node.app) + '</text>');
      out.push('</g>');
    });
    if (layout.selfCount) {
      out.push('<text x="' + (COL_X[1] + NODE_W / 2) + '" y="' + (layout.centerY + NODE_H + 18) + '" font-size="11" text-anchor="middle" fill="#3d5561">' +
        '同じアプリ内のつながり：' + layout.selfCount + '件（一覧を参照）</text>');
    }
    if (layout.edges.length === 0) {
      out.push('<text x="' + (layout.width / 2) + '" y="' + (layout.height - 12) + '" font-size="12" text-anchor="middle" fill="#777777">' +
        'ほかのアプリとのつながりは見つかりませんでした</text>');
    }
    out.push('</svg>');
    return out.join('');
  }

  // ===== 3. 一覧表とCSV =====

  function linkRows(hood, names) {
    function name(app) { return app === null ? '（不明）' : (names[app] || ('アプリID ' + app)); }
    var rows = [];
    function push(direction, link, other) {
      rows.push([direction, kindLabel(link.kind), name(other), other === null ? '' : other, link.place, link.detail]);
    }
    hood.outgoing.slice().sort(byAppId).forEach(function (item) {
      item.outgoing.forEach(function (l) { push('参照している先', l, item.app); });
    });
    hood.unknown.forEach(function (l) { push('参照している先', l, null); });
    hood.incomingOnly.concat(hood.outgoing.filter(function (a) { return a.incoming.length; }))
      .sort(byAppId).forEach(function (item) {
        item.incoming.forEach(function (l) { push('参照されている元', l, item.app); });
      });
    hood.self.forEach(function (l) { push('同じアプリ内', l, l.to); });
    return rows;
  }

  var ROW_HEADER = ['向き', 'つながりの種類', '相手のアプリ', '相手のアプリID', '設定のある場所', '詳細'];

  function csvEscape(value) {
    var s = value === null || value === undefined ? '' : String(value);
    return /[",\r\n]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s;
  }

  function buildCsv(header, rows) {
    return '﻿' + [header].concat(rows).map(function (r) { return r.map(csvEscape).join(','); }).join('\r\n') + '\r\n';
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
  function buildExcel(rows, info) {
    return Xlsx.listBook({
      title: 'アプリのつながり', target: info.app, sub: excelSub(info, 'アプリ相関マップ'), sheet: 'つながり',
      header: ROW_HEADER, rows: rows, widths: [14, 18, 26, 12, 30, 50], ids: [3], creator: '台帳屋 アプリ相関マップ',
      note: rows.length ? 'ルックアップ・関連レコード一覧・アクションでつながるアプリです。図は「図を画像で保存」で保存できます。' : 'つながっているアプリはありません。'
    });
  }

  function safeFileName(name) {
    var s = String(name || '').replace(/[\\/:*?"<>|]/g, '_').trim();
    return s || 'app';
  }

  function dateStamp(d) {
    var pad = function (n) { return ('0' + n).slice(-2); };
    return d.getFullYear() + pad(d.getMonth() + 1) + pad(d.getDate());
  }

  function scanNote(scan) {
    var notes = [];
    notes.push(scan.total + '個のアプリを調べました（操作した人が見られるアプリだけです。見られないアプリは図に入りません）。');
    if (scan.fieldErrors) notes.push('フィールドを読めなかったアプリが' + scan.fieldErrors + '個あります（閲覧の権限がないため）。');
    if (scan.actionErrors) notes.push('アクションを読めなかったアプリが' + scan.actionErrors + '個あります（アクションはアプリの管理権限が必要です）。');
    return notes.join('');
  }

  var Core = {
    extractLinks: extractLinks,
    neighborhood: neighborhood,
    edgeLines: edgeLines,
    layoutMap: layoutMap,
    buildSvg: buildSvg,
    linkRows: linkRows,
    ROW_HEADER: ROW_HEADER,
    buildCsv: buildCsv,
    buildExcel: buildExcel,
    truncate: truncate,
    safeFileName: safeFileName,
    dateStamp: dateStamp,
    scanNote: scanNote
  };

  if (typeof module !== 'undefined' && module.exports) {
    module.exports = Core;
    return;
  }
  if (typeof kintone === 'undefined') return;

  // ===== 4. 画面（kintone上でだけ動く部分） =====

  var P = 'dcy-am';
  var PARALLEL = 4;

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

  // すべてのアプリのフィールドとアクションを、同時に4つずつ読む
  function scanAll(apps, onProgress) {
    var links = [];
    var scan = { total: apps.length, fieldErrors: 0, actionErrors: 0 };
    var index = 0;
    var done = 0;
    function worker() {
      if (index >= apps.length) return Promise.resolve();
      var app = apps[index++];
      return Promise.all([
        settle(api('/k/v1/app/form/fields.json', { app: app.appId })),
        settle(api('/k/v1/app/actions.json', { app: app.appId }))
      ]).then(function (res) {
        if (!res[0].ok) scan.fieldErrors += 1;
        if (!res[1].ok) scan.actionErrors += 1;
        var props = res[0].ok ? res[0].value.properties : {};
        var actions = res[1].ok ? res[1].value : null;
        links = links.concat(extractLinks(app.appId, props, actions));
        done += 1;
        onProgress(done, apps.length);
      }).then(worker);
    }
    var workers = [];
    for (var i = 0; i < PARALLEL; i++) workers.push(worker());
    return Promise.all(workers).then(function () { return { links: links, scan: scan }; });
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

  function saveBlob(blob, fileName) {
    var url = URL.createObjectURL(blob);
    var a = el('a', { href: url, download: fileName });
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    setTimeout(function () { URL.revokeObjectURL(url); }, 1000);
  }

  function svgToPng(svgText, width, height) {
    return new Promise(function (resolve, reject) {
      var img = new Image();
      var url = URL.createObjectURL(new Blob([svgText], { type: 'image/svg+xml;charset=utf-8' }));
      img.onload = function () {
        var scale = 2;
        var canvas = document.createElement('canvas');
        canvas.width = width * scale;
        canvas.height = height * scale;
        var ctx = canvas.getContext('2d');
        ctx.scale(scale, scale);
        ctx.drawImage(img, 0, 0);
        URL.revokeObjectURL(url);
        canvas.toBlob(function (blob) { blob ? resolve(blob) : reject(new Error('画像に変換できませんでした')); }, 'image/png');
      };
      img.onerror = function () { URL.revokeObjectURL(url); reject(new Error('画像に変換できませんでした')); };
      img.src = url;
    });
  }

  function renderTable(rows) {
    var thead = el('thead', {}, [el('tr', {}, ROW_HEADER.map(function (h) { return el('th', { text: h }); }))]);
    var tbody = el('tbody');
    if (!rows.length) {
      tbody.appendChild(el('tr', {}, [el('td', { colspan: String(ROW_HEADER.length), class: P + '-empty', text: 'つながりはありません' })]));
    }
    rows.forEach(function (r) {
      tbody.appendChild(el('tr', {}, r.map(function (v) { return el('td', { text: v }); })));
    });
    return el('table', { class: P + '-table' }, [thead, tbody]);
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
    var names = {};
    var cache = null;
    var current = null;
    var select = el('select', { class: P + '-select', 'aria-label': '中心にするアプリ' });
    var runBtn = el('button', { type: 'button', class: P + '-btn ' + P + '-primary', text: '相関マップを表示' });
    var pngBtn = el('button', { type: 'button', class: P + '-btn', text: '画像で保存（PNG）', disabled: 'disabled' });
    var excelBtn = el('button', { type: 'button', class: P + '-btn', text: '一覧をExcelで保存', disabled: 'disabled' });
    var csvBtn = el('button', { type: 'button', class: P + '-btn', text: '一覧をCSVで書き出す', disabled: 'disabled' });
    var closeBtn = el('button', { type: 'button', class: P + '-close', 'aria-label': '閉じる', text: '×' });
    var status = el('p', { class: P + '-status', role: 'status' });
    var mapBox = el('div', { class: P + '-map' });
    var tableBox = el('div', { class: P + '-list' });
    var dialog = el('div', { class: P + '-dialog', role: 'dialog', 'aria-modal': 'true', 'aria-label': 'アプリ相関マップ' }, [
      el('div', { class: P + '-head' }, [el('h2', { text: 'アプリ相関マップ' }), closeBtn]),
      el('p', { class: P + '-lead', text: 'ルックアップ・関連レコード一覧・アクションでつながっているアプリを図にします。図のアプリを押すと、そのアプリを中心に切り替えます。' }),
      el('div', { class: P + '-controls' }, [select, runBtn, pngBtn, excelBtn, csvBtn]),
      status, mapBox, tableBox,
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

    function show(appId) {
      var hood = neighborhood(cache.links, appId);
      var layout = layoutMap(hood, names);
      var title = (names[appId] || ('アプリID ' + appId)) + ' の相関マップ';
      var svg = buildSvg(layout, names, title);
      var rows = linkRows(hood, names);
      current = { appId: String(appId), svg: svg, layout: layout, rows: rows };
      mapBox.innerHTML = svg;
      Array.prototype.forEach.call(mapBox.querySelectorAll('.' + P + '-node'), function (g) {
        var target = g.getAttribute('data-app');
        if (target === String(appId) || !names[target]) return;
        g.setAttribute('tabindex', '0');
        g.setAttribute('role', 'button');
        g.classList.add(P + '-clickable');
        var go = function () { select.value = target; show(target); };
        g.addEventListener('click', go);
        g.addEventListener('keydown', function (e) { if (e.key === 'Enter') go(); });
      });
      tableBox.innerHTML = '';
      tableBox.appendChild(el('h3', { text: 'つながりの一覧（' + rows.length + '件）' }));
      tableBox.appendChild(renderTable(rows));
      status.textContent = scanNote(cache.scan);
      pngBtn.disabled = false;
      csvBtn.disabled = false;
    }

    runBtn.addEventListener('click', function () {
      var appId = select.value;
      if (!appId) return;
      if (cache) { show(appId); return; }
      runBtn.disabled = true;
      var apps = Object.keys(names).map(function (id) { return { appId: id }; });
      status.textContent = 'アプリのつながりを調べています（0/' + apps.length + '）…';
      scanAll(apps, function (done, total) {
        status.textContent = 'アプリのつながりを調べています（' + done + '/' + total + '）…';
      }).then(function (result) {
        cache = result;
        show(appId);
      }).catch(function (error) {
        status.textContent = '相関マップを作れませんでした（' + ((error && error.message) || '原因不明') + '）。';
      }).then(function () {
        runBtn.disabled = false;
      });
    });

    pngBtn.addEventListener('click', function () {
      if (!current) return;
      svgToPng(current.svg, current.layout.width, current.layout.height).then(function (blob) {
        saveBlob(blob, '相関マップ_' + safeFileName(names[current.appId]) + '_' + dateStamp(new Date()) + '.png');
      }).catch(function (error) { status.textContent = error.message; });
    });

    excelBtn.addEventListener('click', function () {
      if (!current) return;
      saveExcel(buildExcel(current.rows, { app: names[current.appId] || ('アプリID ' + current.appId), host: location.hostname, at: new Date() }), '相関マップ_' + safeFileName(names[current.appId]) + '_' + dateStamp(new Date()) + '.xlsx');
    });
    syncButton(csvBtn, excelBtn);
    csvBtn.addEventListener('click', function () {
      if (!current) return;
      var blob = new Blob([buildCsv(ROW_HEADER, current.rows)], { type: 'text/csv;charset=utf-8' });
      saveBlob(blob, '相関マップ_' + safeFileName(names[current.appId]) + '_' + dateStamp(new Date()) + '.csv');
    });

    document.body.appendChild(overlay);
    status.textContent = 'アプリの一覧を読み込み中…';
    getAllApps().then(function (apps) {
      moreCount(apps.length);
      var thisApp = String(kintone.app.getId());
      apps.forEach(function (app) {
        names[String(app.appId)] = app.name;
        var option = el('option', { value: app.appId, text: app.name + '（ID: ' + app.appId + '）' });
        if (String(app.appId) === thisApp) option.selected = true;
        select.appendChild(option);
      });
      status.textContent = 'アプリを選んで「相関マップを表示」を押してください。最初の1回は、すべてのアプリを調べるため少し時間がかかります。';
    }).catch(function (error) {
      status.textContent = 'アプリの一覧を取得できませんでした（' + ((error && error.message) || '原因不明') + '）。';
    });
  }

  kintone.events.on('app.record.index.show', function (event) {
    var space = kintone.app.getHeaderMenuSpaceElement();
    if (!space || document.getElementById(P + '-open')) return event;
    var button = el('button', { id: P + '-open', type: 'button', class: P + '-btn', text: '相関マップ' });
    button.addEventListener('click', openModal);
    space.appendChild(button);
    return event;
  });
})();
