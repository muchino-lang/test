/* ==========================================================
   ルール一覧の検索・絞り込み
   - データは docs/data/rules.js（window.RULES）
   - 検索条件は URL に反映されるので、そのまま共有できます
   ========================================================== */
(function () {
  'use strict';

  var RULES = window.RULES || [];
  var COLUMNS = [
    { key: 'category', label: 'カテゴリ' },
    { key: 'doc', label: 'ドキュメント' },
    { key: 'target', label: '対象' },
    { key: 'item', label: '項目' },
    { key: 'content', label: '内容' }
  ];

  var esc = window.DocExport.escapeHtml;
  var state = { q: '', categories: [], docs: [] };

  var els = {};

  /* ---------- URL との同期 ---------- */
  function readUrl() {
    var params = new URLSearchParams(location.search);
    state.q = params.get('q') || '';
    state.categories = (params.get('cat') || '').split(',').filter(Boolean);
    state.docs = (params.get('doc') || '').split(',').filter(Boolean);
  }

  function writeUrl() {
    var params = new URLSearchParams();
    if (state.q) { params.set('q', state.q); }
    if (state.categories.length) { params.set('cat', state.categories.join(',')); }
    if (state.docs.length) { params.set('doc', state.docs.join(',')); }
    var qs = params.toString();
    history.replaceState(null, '', qs ? '?' + qs : location.pathname);
  }

  /* ---------- 絞り込み ---------- */
  function normalize(s) {
    return String(s || '').toLowerCase();
  }

  function terms() {
    return state.q.trim().split(/[\s　]+/).filter(Boolean).map(normalize);
  }

  function filtered() {
    var words = terms();
    return RULES.filter(function (r) {
      if (state.categories.length && state.categories.indexOf(r.category) === -1) { return false; }
      if (state.docs.length && state.docs.indexOf(r.doc) === -1) { return false; }
      if (!words.length) { return true; }
      var haystack = normalize([r.category, r.doc, r.section, r.target, r.item, r.content].join(' '));
      return words.every(function (w) { return haystack.indexOf(w) !== -1; });
    });
  }

  /* ---------- 描画 ---------- */
  function highlight(text) {
    var html = esc(text);
    terms().forEach(function (w) {
      if (!w) { return; }
      var re = new RegExp(w.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'gi');
      html = html.replace(re, function (m) { return '<mark>' + m + '</mark>'; });
    });
    return html;
  }

  function renderChips() {
    function chipSet(container, values, selected, onToggle) {
      container.innerHTML = '';
      values.forEach(function (value) {
        var btn = document.createElement('button');
        btn.type = 'button';
        btn.className = 'chip' + (selected.indexOf(value) !== -1 ? ' chip--on' : '');
        btn.textContent = value;
        btn.setAttribute('aria-pressed', selected.indexOf(value) !== -1 ? 'true' : 'false');
        btn.addEventListener('click', function () { onToggle(value); });
        container.appendChild(btn);
      });
    }

    var categories = [];
    var docs = [];
    RULES.forEach(function (r) {
      if (categories.indexOf(r.category) === -1) { categories.push(r.category); }
      if (docs.indexOf(r.doc) === -1) { docs.push(r.doc); }
    });
    categories.sort(function (a, b) { return a.localeCompare(b, 'ja'); });

    chipSet(els.categories, categories, state.categories, function (value) {
      toggle(state.categories, value);
      render();
    });
    chipSet(els.docs, docs, state.docs, function (value) {
      toggle(state.docs, value);
      render();
    });
  }

  function toggle(list, value) {
    var i = list.indexOf(value);
    if (i === -1) { list.push(value); } else { list.splice(i, 1); }
  }

  function renderTable(rows) {
    if (!rows.length) {
      els.tbody.innerHTML = '<tr><td colspan="' + COLUMNS.length +
        '" class="rules-table__empty">該当するルールがありません。条件を変えてお試しください。</td></tr>';
      return;
    }

    els.tbody.innerHTML = rows.map(function (r) {
      var link = '<a href="' + esc(r.href) + '">' + highlight(r.doc) + '</a>' +
        '<span class="rules-table__section">' + esc(r.section) + '</span>';
      return '<tr>' +
        '<td data-label="カテゴリ"><span class="tag">' + highlight(r.category) + '</span></td>' +
        '<td data-label="ドキュメント">' + link + '</td>' +
        '<td data-label="対象">' + highlight(r.target) + '</td>' +
        '<td data-label="項目"><strong>' + highlight(r.item) + '</strong></td>' +
        '<td data-label="内容">' + highlight(r.content) + '</td>' +
        '</tr>';
    }).join('');
  }

  function render() {
    var rows = filtered();
    renderChips();
    renderTable(rows);
    els.count.textContent = rows.length === RULES.length
      ? '全 ' + RULES.length + ' 件'
      : rows.length + ' 件 / 全 ' + RULES.length + ' 件';
    els.reset.hidden = !(state.q || state.categories.length || state.docs.length);
    els.input.value = state.q;
    writeUrl();
    els.current = rows;
  }

  /* ---------- 出力 ---------- */
  function exportRows() {
    var header = COLUMNS.map(function (c) { return c.label; });
    header.splice(2, 0, 'セクション');
    return [header].concat(els.current.map(function (r) {
      return [r.category, r.doc, r.section, r.target, r.item, r.content];
    }));
  }

  function exportWord() {
    var head = '<tr>' + ['カテゴリ', 'ドキュメント', 'セクション', '対象', '項目', '内容']
      .map(function (h) { return '<th>' + h + '</th>'; }).join('') + '</tr>';
    var body = els.current.map(function (r) {
      return '<tr>' + [r.category, r.doc, r.section, r.target, r.item, r.content]
        .map(function (c) { return '<td>' + esc(c) + '</td>'; }).join('') + '</tr>';
    }).join('');
    var html = '<h1>ルール一覧</h1><p class="doc__meta">' + esc(conditionText()) +
      '</p><table>' + head + body + '</table>';
    window.DocExport.downloadWord(html, 'ルール一覧', 'ルール一覧');
  }

  function conditionText() {
    var parts = [];
    if (state.q) { parts.push('キーワード: ' + state.q); }
    if (state.categories.length) { parts.push('カテゴリ: ' + state.categories.join('、')); }
    if (state.docs.length) { parts.push('ドキュメント: ' + state.docs.join('、')); }
    parts.push(els.current.length + ' 件');
    return parts.join(' / ');
  }

  /* ---------- 初期化 ---------- */
  function init() {
    els.input = document.getElementById('rules-search');
    els.categories = document.getElementById('rules-categories');
    els.docs = document.getElementById('rules-docs');
    els.tbody = document.getElementById('rules-body');
    els.count = document.getElementById('rules-count');
    els.reset = document.getElementById('rules-reset');

    els.input.addEventListener('input', function () {
      state.q = els.input.value;
      render();
    });

    els.reset.addEventListener('click', function () {
      state.q = '';
      state.categories = [];
      state.docs = [];
      render();
      els.input.focus();
    });

    window.DocExport.buildToolbar([
      ['Word で出力', exportWord],
      ['Excel で出力', function () {
        window.DocExport.downloadExcel(exportRows(), 'ルール一覧', 'ルール一覧');
      }],
      ['印刷 / PDF', function () { window.print(); }]
    ]);

    readUrl();
    render();
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init);
  } else {
    init();
  }
})();
