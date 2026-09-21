/* ==========================================================
   Word / Excel 出力
   - 外部ライブラリ不要（ブラウザだけで完結）
   - Word : HTML 形式の .doc（Word でそのまま開けます）
   - Excel: HTML テーブル形式の .xls（Excel でそのまま開けます）
   より厳密な .docx / .xlsx が必要な場合は tools/export.py を使用してください。
   ========================================================== */
(function () {
  'use strict';

  var WORD_CSS =
    'body{font-family:"Yu Gothic","Hiragino Sans","Meiryo",sans-serif;font-size:11pt;line-height:1.8;}' +
    'h1{font-size:18pt;margin:0 0 4pt;}' +
    'h2{font-size:13pt;margin:16pt 0 6pt;border-left:4pt solid #2f80ed;padding-left:6pt;}' +
    '.doc__meta{color:#666;font-size:9pt;margin:0 0 16pt;}' +
    'table{border-collapse:collapse;width:100%;margin:0 0 12pt;}' +
    'th,td{border:1px solid #bbb;padding:5pt 7pt;text-align:left;vertical-align:top;font-size:10.5pt;}' +
    'th{background:#eef4fd;}' +
    '.rules th{width:30%;}' +
    '.note{border-left:4pt solid #2f80ed;background:#f2f7ff;padding:8pt 10pt;margin:12pt 0;}' +
    '.note--warn{border-left-color:#f5a623;background:#fff8ec;}' +
    '.step{background:#2f80ed;color:#fff;padding:1pt 6pt;border-radius:8pt;font-size:9pt;}';

  function escapeHtml(s) {
    return String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;')
      .replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  }

  function download(blob, filename) {
    var url = URL.createObjectURL(blob);
    var a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    setTimeout(function () { URL.revokeObjectURL(url); }, 1000);
  }

  /** 本文 HTML を Word ファイルとしてダウンロードする */
  function downloadWord(bodyHtml, filename, title) {
    var html =
      '\uFEFF<html xmlns:o="urn:schemas-microsoft-com:office:office" ' +
      'xmlns:w="urn:schemas-microsoft-com:office:word" xmlns="http://www.w3.org/TR/REC-html40">' +
      '<head><meta charset="utf-8"><title>' + escapeHtml(title || filename) + '</title>' +
      '<!--[if gte mso 9]><xml><w:WordDocument><w:View>Print</w:View>' +
      '<w:Zoom>100</w:Zoom><w:DoNotOptimizeForBrowser/></w:WordDocument></xml><![endif]-->' +
      '<style>@page{size:A4;margin:20mm;}' + WORD_CSS + '</style></head>' +
      '<body>' + bodyHtml + '</body></html>';
    download(new Blob([html], { type: 'application/msword;charset=utf-8' }), filename + '.doc');
  }

  /** 2次元配列（1行目が見出し）を Excel ファイルとしてダウンロードする */
  function downloadExcel(rows, filename, sheetName) {
    var body = rows.map(function (r, i) {
      var tag = i === 0 ? 'th' : 'td';
      return '<tr>' + r.map(function (c) {
        return '<' + tag + '>' + escapeHtml(c) + '</' + tag + '>';
      }).join('') + '</tr>';
    }).join('');

    var name = String(sheetName || 'Sheet1').replace(/[\\\/\?\*\[\]:]/g, ' ').slice(0, 31);

    var html =
      '\uFEFF<html xmlns:o="urn:schemas-microsoft-com:office:office" ' +
      'xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40">' +
      '<head><meta charset="utf-8">' +
      '<!--[if gte mso 9]><xml><x:ExcelWorkbook><x:ExcelWorksheets><x:ExcelWorksheet>' +
      '<x:Name>' + escapeHtml(name) + '</x:Name><x:WorksheetOptions><x:DisplayGridlines/>' +
      '</x:WorksheetOptions></x:ExcelWorksheet></x:ExcelWorksheets></x:ExcelWorkbook></xml><![endif]-->' +
      '<style>td,th{border:1px solid #bbb;padding:4px 6px;vertical-align:top;' +
      'mso-number-format:"\\@";}th{background:#eef4fd;font-weight:bold;}</style></head>' +
      '<body><table>' + body + '</table></body></html>';

    download(new Blob([html], { type: 'application/vnd.ms-excel;charset=utf-8' }), filename + '.xls');
  }

  /** ページ上部にツールバーを作る。buttons = [[ラベル, 処理], ...] */
  function buildToolbar(buttons) {
    var root = document.querySelector('.doc');
    if (!root) { return; }

    var bar = document.createElement('div');
    bar.className = 'doc__toolbar';
    bar.setAttribute('role', 'toolbar');
    bar.setAttribute('aria-label', '出力');

    buttons.forEach(function (item) {
      var btn = document.createElement('button');
      btn.type = 'button';
      btn.className = 'doc__btn';
      btn.textContent = item[0];
      btn.addEventListener('click', item[1]);
      bar.appendChild(btn);
    });

    var anchor = root.querySelector('.doc__meta') || root.querySelector('h1');
    if (anchor && anchor.nextSibling) {
      root.insertBefore(bar, anchor.nextSibling);
    } else {
      root.appendChild(bar);
    }
    return bar;
  }

  window.DocExport = {
    escapeHtml: escapeHtml,
    download: download,
    downloadWord: downloadWord,
    downloadExcel: downloadExcel,
    buildToolbar: buildToolbar
  };

  /* ---------- 通常のドキュメントページ用の既定動作 ---------- */
  function docRoot() { return document.querySelector('.doc'); }
  function fileBase() {
    var root = docRoot();
    return (root && root.dataset.file) || 'document';
  }

  function contentHtml() {
    var clone = docRoot().cloneNode(true);
    clone.querySelectorAll('.doc__toolbar, .doc__back').forEach(function (el) {
      el.parentNode.removeChild(el);
    });
    return clone.innerHTML;
  }

  function collectRows() {
    var rows = [['区分', '項目', '内容']];
    var section = '';

    docRoot().querySelectorAll('h1, h2, p, li, .rules tr').forEach(function (el) {
      if (el.closest('.doc__toolbar')) { return; }
      var text = (el.textContent || '').replace(/\s+/g, ' ').trim();

      if (el.tagName === 'H1') {
        rows.push(['タイトル', '', text]);
      } else if (el.tagName === 'H2') {
        section = text;
        rows.push(['見出し', text, '']);
      } else if (el.tagName === 'TR') {
        var th = el.querySelector('th');
        var td = el.querySelector('td');
        rows.push([section, th ? th.textContent.trim() : '', td ? td.textContent.trim() : '']);
      } else if (el.tagName === 'LI') {
        rows.push([section, '箇条書き', text]);
      } else if (text) {
        if (el.classList.contains('doc__meta')) {
          rows.push(['メタ情報', '', text]);
        } else {
          rows.push([section, el.closest('.note') ? '注記' : '本文', text]);
        }
      }
    });
    return rows;
  }

  function init() {
    var root = docRoot();
    // 検索ページなど、独自のツールバーを持つページでは何もしない
    if (!root || root.dataset.toolbar === 'custom') { return; }

    buildToolbar([
      ['Word で出力', function () { downloadWord(contentHtml(), fileBase(), document.title); }],
      ['Excel で出力', function () { downloadExcel(collectRows(), fileBase(), document.title); }],
      ['印刷 / PDF', function () { window.print(); }]
    ]);
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init);
  } else {
    init();
  }
})();
