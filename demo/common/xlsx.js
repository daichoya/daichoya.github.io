/*
 * 台帳屋 共通部品：Excel（.xlsx）を作る
 * 上司や監査に出す資料として見やすいよう、どの製品でも同じ体裁にそろえる。
 *  - 表紙の見出し（文書名・対象・作成日時）、目次から各シートへのリンク
 *  - 見出し行は濃い紺に白い文字、罫線は細い灰色、長い文字は折り返す
 *  - 一覧のシートは見出し行を固定し、フィルターを付ける
 *  - 印刷はA4で、幅を1ページに合わせ、見出し行をページごとに繰り返す。ヘッダーに文書名とシート名、フッターにページ番号
 * 外部のライブラリを使わず、ブラウザの中だけで作る（通信しない）。
 *
 * sheets: [{
 *   name, widths: [列幅], hidden, orientation: 'landscape'|'portrait', list: true（固定行とフィルター）,
 *   heading: { title, sub },         // シートの先頭に大きな見出し
 *   blocks: [{ title, header: [], rows: [[]], note, styles: [行ごとのスタイル], labels: true（1列目を項目名の色に）, numbers: [数にする列（桁区切り）], ids: [IDの列（数・桁区切りなし）], centers: [中央にそろえる列] }]
 * }]
 * span：その表の最後の列を、右へ何列分つなげるか（例：項目名と内容の2列の表を4列の幅に）。
 * セルの値は文字・数のほか { text, link: 'シート名' }（そのシートへのリンク）も使える。
 */
(function (root) {
  'use strict';

  // ===== zip（無圧縮） =====
  var CRC_TABLE = (function () {
    var table = new Uint32Array(256);
    for (var n = 0; n < 256; n++) {
      var c = n;
      for (var k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
      table[n] = c >>> 0;
    }
    return table;
  })();

  function crc32(bytes) {
    var crc = 0xffffffff;
    for (var i = 0; i < bytes.length; i++) crc = CRC_TABLE[(crc ^ bytes[i]) & 0xff] ^ (crc >>> 8);
    return (crc ^ 0xffffffff) >>> 0;
  }

  function utf8(text) { return new TextEncoder().encode(text); }

  function concatBytes(chunks) {
    var total = chunks.reduce(function (sum, c) { return sum + c.length; }, 0);
    var out = new Uint8Array(total);
    var pos = 0;
    chunks.forEach(function (c) { out.set(c, pos); pos += c.length; });
    return out;
  }

  function buildZip(files) {
    var local = [];
    var central = [];
    var offset = 0;
    files.forEach(function (file) {
      var name = utf8(file.name);
      var data = file.data;
      var crc = crc32(data);
      var head = new DataView(new ArrayBuffer(30));
      head.setUint32(0, 0x04034b50, true);
      head.setUint16(4, 20, true);
      head.setUint16(6, 0x0800, true); // ファイル名はUTF-8
      head.setUint16(8, 0, true);
      head.setUint32(14, crc, true);
      head.setUint32(18, data.length, true);
      head.setUint32(22, data.length, true);
      head.setUint16(26, name.length, true);
      local.push(new Uint8Array(head.buffer), name, data);
      var dir = new DataView(new ArrayBuffer(46));
      dir.setUint32(0, 0x02014b50, true);
      dir.setUint16(4, 20, true);
      dir.setUint16(6, 20, true);
      dir.setUint16(8, 0x0800, true);
      dir.setUint32(16, crc, true);
      dir.setUint32(20, data.length, true);
      dir.setUint32(24, data.length, true);
      dir.setUint16(28, name.length, true);
      dir.setUint32(42, offset, true);
      central.push(new Uint8Array(dir.buffer), name);
      offset += 30 + name.length + data.length;
    });
    var centralBytes = concatBytes(central);
    var end = new DataView(new ArrayBuffer(22));
    end.setUint32(0, 0x06054b50, true);
    end.setUint16(8, files.length, true);
    end.setUint16(10, files.length, true);
    end.setUint32(12, centralBytes.length, true);
    end.setUint32(16, offset, true);
    return concatBytes(local.concat([centralBytes, new Uint8Array(end.buffer)]));
  }

  // ===== 文字とセル =====
  function xmlEscape(value) {
    return String(value === undefined || value === null ? '' : value)
      .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '')
      .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
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
  function sheetName(name) { return String(name).replace(/[\[\]:*?\/\\]/g, '_').slice(0, 31) || 'Sheet'; }

  function quoteSheet(name) { return "'" + sheetName(name).replace(/'/g, "''") + "'"; }

  // スタイルの番号（cellXfs の順）。0〜4 はこれまでの製品と同じ意味
  var STYLE = { plain: 0, header: 1, body: 2, title: 3, warn: 4, docTitle: 5, label: 6, number: 7, note: 8, important: 9, sub: 10, link: 11, good: 12, center: 13, id: 14 };

  var NAVY = 'FF1F3A5F';
  var LINE = 'FFB7C0CB';
  var FONT = '<name val="Yu Gothic"/><family val="3"/><charset val="128"/>';

  var STYLES_XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
    '<numFmts count="1"><numFmt numFmtId="164" formatCode="#,##0"/></numFmts>' +
    '<fonts count="8">' +
    '<font><sz val="10"/>' + FONT + '</font>' +                                     // 0 本文
    '<font><b/><sz val="10"/><color rgb="FFFFFFFF"/>' + FONT + '</font>' +         // 1 見出し行
    '<font><b/><sz val="16"/><color rgb="' + NAVY + '"/>' + FONT + '</font>' +     // 2 文書名
    '<font><b/><sz val="11"/><color rgb="' + NAVY + '"/>' + FONT + '</font>' +     // 3 表の題
    '<font><sz val="9"/><color rgb="FF5F6368"/>' + FONT + '</font>' +              // 4 注記
    '<font><b/><sz val="10"/><color rgb="FF1F2933"/>' + FONT + '</font>' +         // 5 項目名
    '<font><u/><sz val="10"/><color rgb="FF1A5FB4"/>' + FONT + '</font>' +         // 6 リンク
    '<font><sz val="10"/><color rgb="FF5F6368"/>' + FONT + '</font>' +             // 7 副題
    '</fonts>' +
    '<fills count="7"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>' +
    '<fill><patternFill patternType="solid"><fgColor rgb="' + NAVY + '"/><bgColor indexed="64"/></patternFill></fill>' +  // 2 見出し行
    '<fill><patternFill patternType="solid"><fgColor rgb="FFEEF2F6"/><bgColor indexed="64"/></patternFill></fill>' +     // 3 項目名
    '<fill><patternFill patternType="solid"><fgColor rgb="FFFFF2CC"/><bgColor indexed="64"/></patternFill></fill>' +     // 4 確認（黄）
    '<fill><patternFill patternType="solid"><fgColor rgb="FFFCE4E1"/><bgColor indexed="64"/></patternFill></fill>' +     // 5 要確認（赤）
    '<fill><patternFill patternType="solid"><fgColor rgb="FFE6F4EA"/><bgColor indexed="64"/></patternFill></fill>' +     // 6 問題なし（緑）
    '</fills>' +
    '<borders count="3"><border><left/><right/><top/><bottom/><diagonal/></border>' +
    '<border><left style="thin"><color rgb="' + LINE + '"/></left><right style="thin"><color rgb="' + LINE + '"/></right>' +
    '<top style="thin"><color rgb="' + LINE + '"/></top><bottom style="thin"><color rgb="' + LINE + '"/></bottom><diagonal/></border>' +
    '<border><left/><right/><top/><bottom style="medium"><color rgb="' + NAVY + '"/></bottom><diagonal/></border></borders>' +
    '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
    '<cellXfs count="15">' +
    '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>' +                                                                                                              // 0 plain
    '<xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment vertical="center" wrapText="1"/></xf>' + // 1 header
    '<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf>' +                                // 2 body
    '<xf numFmtId="0" fontId="3" fillId="0" borderId="0" xfId="0" applyFont="1"/>' +                                                                                                 // 3 title
    '<xf numFmtId="0" fontId="0" fillId="4" borderId="1" xfId="0" applyFill="1" applyBorder="1" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf>' +                  // 4 warn
    '<xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1"/>' +                                                                                                 // 5 docTitle
    '<xf numFmtId="0" fontId="5" fillId="3" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf>' +    // 6 label
    '<xf numFmtId="164" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1" applyAlignment="1"><alignment horizontal="right" vertical="top"/></xf>' +  // 7 number
    '<xf numFmtId="0" fontId="4" fillId="0" borderId="0" xfId="0" applyFont="1"/>' +                                                                                                 // 8 note
    '<xf numFmtId="0" fontId="0" fillId="5" borderId="1" xfId="0" applyFill="1" applyBorder="1" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf>' +                  // 9 important
    '<xf numFmtId="0" fontId="7" fillId="0" borderId="0" xfId="0" applyFont="1"/>' +                                                                                                 // 10 sub
    '<xf numFmtId="0" fontId="6" fillId="0" borderId="1" xfId="0" applyFont="1" applyBorder="1" applyAlignment="1"><alignment vertical="top"/></xf>' +                               // 11 link
    '<xf numFmtId="0" fontId="0" fillId="6" borderId="1" xfId="0" applyFill="1" applyBorder="1" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf>' +                  // 12 good
    '<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="top"/></xf>' +                         // 13 center
    '<xf numFmtId="1" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1" applyAlignment="1"><alignment horizontal="right" vertical="top"/></xf>' +   // 14 id（桁区切りなし）
    '</cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>';

  var NUMERIC = /^-?(0|[1-9]\d{0,11})(\.\d+)?$/;

  // 1つのセルのXML。数（または「数にする列」の数字）は数のセルにする
  function cellXml(ref, value, style, asNumber, links, numberStyle) {
    if (value && typeof value === 'object') {
      if (value.link) links.push({ ref: ref, location: quoteSheet(value.link) + '!A1', display: value.text });
      if (value.style !== undefined) style = value.style;
      else if (value.link) style = STYLE.link;
      value = value.text;
    }
    var isNumber = typeof value === 'number' ? isFinite(value) : (asNumber && typeof value === 'string' && NUMERIC.test(value));
    if (isNumber) {
      var s = style === STYLE.body || style === STYLE.plain ? (numberStyle || STYLE.number) : style;
      return '<c r="' + ref + '" s="' + s + '"><v>' + Number(value) + '</v></c>';
    }
    if (value === '' || value === undefined || value === null) return '<c r="' + ref + '" s="' + style + '"/>';
    return '<c r="' + ref + '" t="inlineStr" s="' + style + '"><is><t xml:space="preserve">' + xmlEscape(value) + '</t></is></c>';
  }

  // 文字の量から、行の高さ（ポイント）を見積もる。全角は2、半角は1の幅として数える
  function rowHeight(value, width) {
    var text = value && typeof value === 'object' ? value.text : value;
    var lines = 0;
    String(text === undefined || text === null ? '' : text).split('\n').forEach(function (line) {
      var w = 0;
      for (var i = 0; i < line.length; i++) w += line.charCodeAt(i) > 0xff ? 2 : 1;
      lines += Math.max(1, Math.ceil(w / Math.max(4, width - 1)));
    });
    return lines > 1 ? Math.min(409, lines * 14 + 4) : 0;
  }

  function headerFooterText(text) { return String(text || '').replace(/&/g, '&&'); }

  // 1つのシートのXMLと、印刷の見出し行・フィルターの範囲
  function sheetXml(sheet, index, opts) {
    var rowsXml = [];
    var links = [];
    var merges = [];
    var r = 1;
    var maxCol = 1;
    var listHeader = 0;
    var listEnd = 0;
    var listCols = 0;
    // span：最後のセルを右へ何列分つなげるか（項目名と内容の表を、下の広い表と同じ幅にそろえる）
    function addRow(values, style, numbers, labels, height, ids, span, centers) {
      if (span > 1) {
        var last = values.length - 1;
        merges.push(columnName(last) + r + ':' + columnName(last + span - 1) + r);
        // Excelは、つなげたセルの高さを文字に合わせて自動で広げないため、文字の量から高さを決めておく
        if (!height && style !== STYLE.header) {
          var width = 0;
          for (var c = last; c < last + span; c++) width += (sheet.widths && sheet.widths[c]) || 10;
          height = rowHeight(values[last], width);
        }
        values = values.concat(new Array(span - 1).join(',').split(',').map(function () { return ''; }));
      }
      maxCol = Math.max(maxCol, values.length);
      var cells = values.map(function (v, i) {
        var s = Array.isArray(style) ? style[i] : style;
        if (labels && i === 0 && s === STYLE.body) s = STYLE.label;
        if (centers && centers.indexOf(i) >= 0 && s === STYLE.body) s = STYLE.center;
        var isId = ids && ids.indexOf(i) >= 0;
        return cellXml(columnName(i) + r, v, s, isId || (numbers && numbers.indexOf(i) >= 0), links, isId ? STYLE.id : null);
      }).join('');
      rowsXml.push('<row r="' + r + '"' + (height ? ' ht="' + height + '" customHeight="1"' : '') + '>' + cells + '</row>');
      r += 1;
    }
    if (sheet.heading) {
      addRow([sheet.heading.title || ''], STYLE.docTitle, null, false, 30);
      if (sheet.heading.sub) addRow([sheet.heading.sub], STYLE.sub);
      r += 1;
    }
    sheet.blocks.forEach(function (block, i) {
      if (i > 0) r += 1; // 表と表の間を1行あける
      if (block.title) addRow([block.title], STYLE.title, null, false, 20);
      var hasHeader = block.header && block.header.length;
      if (hasHeader) {
        if (sheet.list && !listHeader) { listHeader = r; listCols = block.header.length; }
        addRow(block.header, STYLE.header, null, false, null, null, block.span);
      }
      (block.rows || []).forEach(function (row, j) { addRow(row, (block.styles && block.styles[j]) || STYLE.body, block.numbers, block.labels, null, block.ids, block.span, block.centers); });
      if (sheet.list && listHeader && !listEnd) listEnd = r - 1;
      if (block.note) addRow([block.note], STYLE.note);
    });
    var cols = (sheet.widths || []).map(function (w, i) {
      return '<col min="' + (i + 1) + '" max="' + (i + 1) + '" width="' + w + '" customWidth="1"/>';
    }).join('');
    var frozen = sheet.list && listHeader;
    var filter = frozen && listEnd > listHeader ? 'A' + listHeader + ':' + columnName(listCols - 1) + listEnd : '';
    var view = '<sheetView workbookViewId="0"' + (index === 0 ? ' tabSelected="1"' : '') + ' zoomScaleNormal="100">' +
      (frozen ? '<pane ySplit="' + listHeader + '" topLeftCell="A' + (listHeader + 1) + '" activePane="bottomLeft" state="frozen"/><selection pane="bottomLeft" activeCell="A' + (listHeader + 1) + '" sqref="A' + (listHeader + 1) + '"/>' : '') +
      '</sheetView>';
    var head = headerFooterText(opts.title || '');
    var foot = headerFooterText(opts.footer || '');
    var xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
      '<sheetPr>' + (sheet.tabColor ? '<tabColor rgb="' + sheet.tabColor + '"/>' : '') + '<pageSetUpPr fitToPage="1"/></sheetPr>' +
      '<sheetViews>' + view + '</sheetViews>' +
      '<sheetFormatPr defaultRowHeight="18"/>' +
      (cols ? '<cols>' + cols + '</cols>' : '') +
      '<sheetData>' + rowsXml.join('') + '</sheetData>' +
      (filter ? '<autoFilter ref="' + filter + '"/>' : '') +
      (merges.length ? '<mergeCells count="' + merges.length + '">' + merges.map(function (m) { return '<mergeCell ref="' + m + '"/>'; }).join('') + '</mergeCells>' : '') +
      (links.length ? '<hyperlinks>' + links.map(function (l) { return '<hyperlink ref="' + l.ref + '" location="' + xmlEscape(l.location) + '" display="' + xmlEscape(l.display) + '"/>'; }).join('') + '</hyperlinks>' : '') +
      '<printOptions horizontalCentered="1"/>' +
      '<pageMargins left="0.4" right="0.4" top="0.7" bottom="0.6" header="0.3" footer="0.3"/>' +
      '<pageSetup paperSize="9" orientation="' + (sheet.orientation === 'portrait' ? 'portrait' : 'landscape') + '" fitToWidth="1" fitToHeight="0"/>' +
      '<headerFooter><oddHeader>' + xmlEscape('&L&9' + head + '&R&9&A') + '</oddHeader>' +
      '<oddFooter>' + xmlEscape('&L&9' + foot + '&C&9&P / &N&R&9&D') + '</oddFooter></headerFooter>' +
      '</worksheet>';
    return { xml: xml, printTitle: frozen ? listHeader : 0, filter: filter };
  }

  function build(sheets, opts) {
    opts = opts || {};
    var names = sheets.map(function (s) { return sheetName(s.name); });
    var built = sheets.map(function (s, i) { return sheetXml(s, i, opts); });
    var firstVisible = Math.max(0, sheets.findIndex(function (s) { return !s.hidden; }));
    var defined = [];
    built.forEach(function (b, i) {
      if (b.printTitle) defined.push('<definedName name="_xlnm.Print_Titles" localSheetId="' + i + '">' + xmlEscape(quoteSheet(names[i])) + '!$' + b.printTitle + ':$' + b.printTitle + '</definedName>');
      if (b.filter) {
        var m = /^A(\d+):([A-Z]+)(\d+)$/.exec(b.filter);
        defined.push('<definedName name="_xlnm._FilterDatabase" localSheetId="' + i + '" hidden="1">' + xmlEscape(quoteSheet(names[i])) + '!$A$' + m[1] + ':$' + m[2] + '$' + m[3] + '</definedName>');
      }
    });
    var contentTypes = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
      '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' +
      '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>' +
      sheets.map(function (s, i) {
        return '<Override PartName="/xl/worksheets/sheet' + (i + 1) + '.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>';
      }).join('') + '</Types>';
    var rootRels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>' +
      '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>' +
      '</Relationships>';
    var now = new Date().toISOString().replace(/\.\d+Z$/, 'Z');
    var core = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">' +
      '<dc:title>' + xmlEscape(opts.title || '') + '</dc:title><dc:creator>' + xmlEscape(opts.creator || '台帳屋') + '</dc:creator>' +
      '<dcterms:created xsi:type="dcterms:W3CDTF">' + now + '</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">' + now + '</dcterms:modified>' +
      '</cp:coreProperties>';
    var workbook = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
      '<bookViews><workbookView activeTab="' + firstVisible + '"/></bookViews><sheets>' +
      names.map(function (n, i) {
        return '<sheet name="' + xmlEscape(n) + '" sheetId="' + (i + 1) + '"' + (sheets[i].hidden ? ' state="hidden"' : '') + ' r:id="rId' + (i + 1) + '"/>';
      }).join('') + '</sheets>' +
      (defined.length ? '<definedNames>' + defined.join('') + '</definedNames>' : '') + '</workbook>';
    var workbookRels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      sheets.map(function (s, i) {
        return '<Relationship Id="rId' + (i + 1) + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet' + (i + 1) + '.xml"/>';
      }).join('') +
      '<Relationship Id="rId' + (sheets.length + 1) + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>';
    var files = [
      { name: '[Content_Types].xml', data: utf8(contentTypes) },
      { name: '_rels/.rels', data: utf8(rootRels) },
      { name: 'docProps/core.xml', data: utf8(core) },
      { name: 'xl/workbook.xml', data: utf8(workbook) },
      { name: 'xl/_rels/workbook.xml.rels', data: utf8(workbookRels) },
      { name: 'xl/styles.xml', data: utf8(STYLES_XML) }
    ];
    built.forEach(function (b, i) { files.push({ name: 'xl/worksheets/sheet' + (i + 1) + '.xml', data: utf8(b.xml) }); });
    return buildZip(files);
  }

  // 一覧1枚の資料（無料のプラグインの「Excelで保存」）。見出し・固定行・フィルター・印刷設定は build と同じ
  // o: { title, target, sub, sheet, header, rows, widths, styles, ids, numbers, centers, note, creator, orientation }
  function listBook(o) {
    return build([{ name: o.sheet || '一覧', list: true, widths: o.widths, orientation: o.orientation, tabColor: 'FF1F3A5F',
      heading: { title: o.title + (o.target ? '：' + o.target : ''), sub: o.sub },
      blocks: [{ title: null, header: o.header, rows: o.rows, styles: o.styles, ids: o.ids, numbers: o.numbers, centers: o.centers, note: o.note }] }],
    { title: o.title + (o.target ? '：' + o.target : ''), creator: o.creator, footer: o.creator });
  }

  // 目次の表（シート名はそのシートへのリンク）
  function contents(items) {
    return { title: '目次', header: ['シート', '内容'], rows: items.map(function (x) { return [{ text: x[0], link: x[0] }, x[1]]; }) };
  }

  var api = { build: build, listBook: listBook, contents: contents, STYLE: STYLE, buildZip: buildZip, crc32: crc32, utf8: utf8, sheetName: sheetName, columnName: columnName };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.DaichoyaXlsx = api;
})(typeof window !== 'undefined' ? window : this);
