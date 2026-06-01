/* ─── Table Sort / Filter Engine ──────────────────────────────────────────
   Activates on every <table data-sortfilter> in the page.
   • Click any column header → opens a panel with sort + text filter controls.
   • Multiple columns can each have an active filter simultaneously (AND logic).
   • Active sort  → header turns bright blue + ▲/▼ icon.
   • Active filter → header gets amber underline + ● icon.
   • A status bar appears above each table while filters/sort are in use.
────────────────────────────────────────────────────────────────────────── */
(function () {
  'use strict';

  /* ── Shared floating panel (one instance, reused across all tables) ─── */
  var _panel = null, _ctrl = null, _col = -1;

  function buildPanel() {
    var p = document.createElement('div');
    p.className = 'sf-panel';
    p.innerHTML =
      '<div class="sf-panel-col-name" id="sf-col-name"></div>' +
      '<button class="sf-panel-btn" id="sf-asc">' +
        '<span class="sf-panel-icon">↑</span><span id="sf-asc-lbl">Sort A → Z</span>' +
      '</button>' +
      '<button class="sf-panel-btn" id="sf-desc">' +
        '<span class="sf-panel-icon">↓</span><span id="sf-desc-lbl">Sort Z → A</span>' +
      '</button>' +
      '<div class="sf-panel-sep"></div>' +
      '<div class="sf-filter-label">Filter rows containing:</div>' +
      '<div class="sf-filter-row">' +
        '<input type="text" class="sf-filter-input" id="sf-finput" placeholder="e.g. Sunrise, 90%…">' +
        '<button class="sf-filter-x" id="sf-fx" title="Clear filter">✕</button>' +
      '</div>' +
      '<div class="sf-panel-footer">' +
        '<button class="sf-clear-col-btn" id="sf-clearcol">Clear this column</button>' +
      '</div>';

    /* Stop clicks inside the panel from closing it via the document handler */
    p.addEventListener('click', function (e) { e.stopPropagation(); });

    p.querySelector('#sf-asc').addEventListener('click', function (e) {
      e.stopPropagation();
      if (_ctrl) {
        var cur = _ctrl.sortCol && _ctrl.sortCol.idx === _col ? _ctrl.sortCol.dir : null;
        _ctrl.setSort(_col, cur === 'asc' ? null : 'asc');
      }
      closePanel();
    });
    p.querySelector('#sf-desc').addEventListener('click', function (e) {
      e.stopPropagation();
      if (_ctrl) {
        var cur = _ctrl.sortCol && _ctrl.sortCol.idx === _col ? _ctrl.sortCol.dir : null;
        _ctrl.setSort(_col, cur === 'desc' ? null : 'desc');
      }
      closePanel();
    });
    p.querySelector('#sf-finput').addEventListener('input', function () {
      if (_ctrl) _ctrl.setFilter(_col, this.value);
    });
    p.querySelector('#sf-fx').addEventListener('click', function (e) {
      e.stopPropagation();
      p.querySelector('#sf-finput').value = '';
      if (_ctrl) _ctrl.setFilter(_col, '');
    });
    p.querySelector('#sf-clearcol').addEventListener('click', function (e) {
      e.stopPropagation();
      if (_ctrl) _ctrl.clearColumn(_col);
      closePanel();
    });

    document.body.appendChild(p);
    return p;
  }

  function openPanel(ctrl, colIdx, th) {
    if (!_panel) _panel = buildPanel();
    _ctrl = ctrl; _col = colIdx;

    var isNum = ctrl.isNumeric(colIdx);
    _panel.querySelector('#sf-asc-lbl').textContent  = isNum ? 'Sort Low → High'  : 'Sort A → Z';
    _panel.querySelector('#sf-desc-lbl').textContent = isNum ? 'Sort High → Low'  : 'Sort Z → A';

    /* Strip any icon characters from the raw th text for the title */
    var colName = th.textContent.replace(/[▲▼▾●]/g, '').trim();
    _panel.querySelector('#sf-col-name').textContent = colName;

    var sortDir = ctrl.sortCol && ctrl.sortCol.idx === colIdx ? ctrl.sortCol.dir : null;
    _panel.querySelector('#sf-asc').classList.toggle('sf-active',  sortDir === 'asc');
    _panel.querySelector('#sf-desc').classList.toggle('sf-active', sortDir === 'desc');
    _panel.querySelector('#sf-finput').value = ctrl.filters[colIdx] || '';

    /* Position below the header cell, clamped to viewport */
    var rect = th.getBoundingClientRect();
    var panelW = 220;
    _panel.style.top  = (rect.bottom + 3) + 'px';
    _panel.style.left = Math.min(rect.left, window.innerWidth - panelW - 8) + 'px';
    _panel.style.display = 'block';

    setTimeout(function () { _panel.querySelector('#sf-finput').focus(); }, 30);
  }

  function closePanel() {
    if (_panel) _panel.style.display = 'none';
    _ctrl = null; _col = -1;
  }

  /* ── Controller — one per table ─────────────────────────────────────── */
  function SortFilter(table) {
    var self = this;
    this.table  = table;
    this.tbody  = table.querySelector('tbody');
    if (!this.tbody) return;

    this.rows      = Array.from(this.tbody.querySelectorAll('tr'));
    this.origOrder = this.rows.slice();

    var headerTr = table.querySelector('thead tr');
    this.ths = headerTr ? Array.from(headerTr.querySelectorAll('th')) : [];

    this.sortCol = null;   /* { idx, dir } | null */
    this.filters = {};     /* colIdx (number) → filterText (string) */
    this.bar     = null;

    this._makeStatusBar();
    this._initHeaders();
  }

  SortFilter.prototype._makeStatusBar = function () {
    var self = this;
    var bar = document.createElement('div');
    bar.className = 'sf-status-bar';
    bar.style.display = 'none';
    bar.innerHTML =
      '<span class="sf-status-text"></span>' +
      '<button class="sf-status-clear-all">↺ Clear all filters &amp; sort</button>';
    bar.querySelector('.sf-status-clear-all').addEventListener('click', function () {
      self.clearAll();
    });
    /* Insert immediately before the scroll/overflow wrapper that contains the table */
    var wrapper = this.table.closest('.table-scroll') ||
                  this.table.closest('[style*="overflow"]') ||
                  this.table.parentNode;
    wrapper.parentNode.insertBefore(bar, wrapper);
    this.bar = bar;
  };

  SortFilter.prototype._initHeaders = function () {
    var self = this;
    this.ths.forEach(function (th, idx) {
      if (th.textContent.trim() === 'Detail') return; /* skip link-only column */
      th.classList.add('sf-th');
      var icon = document.createElement('span');
      icon.className = 'sf-icon';
      icon.setAttribute('aria-hidden', 'true');
      icon.textContent = '▾';
      th.appendChild(icon);
      th.addEventListener('click', function (e) {
        e.stopPropagation();
        /* Toggle: click open header closes it; click elsewhere opens */
        if (_ctrl === self && _col === idx && _panel && _panel.style.display !== 'none') {
          closePanel();
        } else {
          openPanel(self, idx, th);
        }
      });
    });
  };

  /* Returns true if this column should sort numerically.
     Primary: any cell in the column carries class="currency" or class="pct" — these
              classes are applied by the server-side templates to every money/percent cell,
              so the check is reliable even when most cells show "—".
     Fallback: scan all rows and apply the ≥50% heuristic. */
  SortFilter.prototype.isNumeric = function (colIdx) {
    for (var i = 0; i < this.rows.length; i++) {
      var cell = this.rows[i].cells[colIdx];
      if (!cell) continue;
      var cls = ' ' + cell.className + ' ';
      if (cls.indexOf(' currency ') >= 0 || cls.indexOf(' pct ') >= 0) return true;
    }
    /* Heuristic fallback (catches ad-hoc tables without explicit class markers) */
    var hits = 0, checks = 0;
    for (var i = 0; i < this.rows.length; i++) {
      var cell = this.rows[i].cells[colIdx];
      if (!cell) continue;
      var txt = cell.textContent.trim();
      if (!txt || txt === '—' || txt === 'N/A' || txt === 'Not Available') continue;
      checks++;
      var clean = txt.replace(/[$%\s,]/g, '').replace(/[()]/g, '');
      if (!isNaN(parseFloat(clean))) hits++;
    }
    return checks > 0 && hits / checks >= 0.5;
  };

  /* Parse a displayed cell value into a signed float.
     Handles: "$1,234"  →  1234
              "$-1,234" → -1234   (Python formats negatives as $-N, not ($N))
              "($1,234)"→ -1234   (Excel-style parenthesised negative)
              "95.5%"   →  0.955  */
  SortFilter.prototype._num = function (txt) {
    if (!txt || txt === '—' || txt === 'N/A' || txt === 'Not Available') return null;
    var t = txt.trim();
    /* Detect parenthesised-negative form BEFORE any stripping */
    var parens = /^\(.*\)$/.test(t);
    /* Strip currency symbol, percent, spaces, and commas — keep "-" so parseFloat
       can handle "$-1,234" → "-1234" correctly */
    var clean = t.replace(/[$%\s,]/g, '').replace(/[()]/g, '');
    /* Convert "95.5%" to a decimal fraction */
    var isPct = txt.indexOf('%') >= 0;
    var v = parseFloat(clean);
    if (isNaN(v)) return null;
    if (isPct) v = v / 100;
    /* Parenthesised form is always negative regardless of its numeric sign */
    return (parens && v > 0) ? -v : v;
  };

  SortFilter.prototype.setSort = function (colIdx, dir) {
    this.sortCol = dir ? { idx: colIdx, dir: dir } : null;
    this._render();
  };

  SortFilter.prototype.setFilter = function (colIdx, text) {
    if (text && text.trim()) this.filters[colIdx] = text;
    else delete this.filters[colIdx];
    this._render();
  };

  SortFilter.prototype.clearColumn = function (colIdx) {
    if (this.sortCol && this.sortCol.idx === colIdx) this.sortCol = null;
    delete this.filters[colIdx];
    this._render();
  };

  SortFilter.prototype.clearAll = function () {
    this.sortCol = null;
    this.filters = {};
    this._render();
  };

  SortFilter.prototype._render = function () {
    var self = this;
    /* Always start from original order so clearing sort restores original */
    var rows = this.origOrder.slice();

    /* ── 1. Sort ─────────────────────────────────────────────────────── */
    if (this.sortCol) {
      var sc    = this.sortCol;
      var isNum = self.isNumeric(sc.idx);
      rows.sort(function (a, b) {
        var aT = a.cells[sc.idx] ? a.cells[sc.idx].textContent.trim() : '';
        var bT = b.cells[sc.idx] ? b.cells[sc.idx].textContent.trim() : '';
        var cmp;
        if (isNum) {
          var aV = self._num(aT); if (aV === null) aV = -Infinity;
          var bV = self._num(bT); if (bV === null) bV = -Infinity;
          cmp = aV - bV;
        } else {
          cmp = aT.localeCompare(bT, undefined, { sensitivity: 'base' });
        }
        return sc.dir === 'asc' ? cmp : -cmp;
      });
    }
    /* Re-append in (sorted) order — restores original order when no sort */
    rows.forEach(function (r) { self.tbody.appendChild(r); });

    /* ── 2. Filter visibility ─────────────────────────────────────────── */
    var fEntries = [];
    Object.keys(self.filters).forEach(function (k) {
      var txt = self.filters[k];
      if (txt && txt.trim())
        fEntries.push({ idx: parseInt(k, 10), lower: txt.trim().toLowerCase() });
    });
    rows.forEach(function (row) {
      var show = fEntries.every(function (f) {
        var cell = row.cells[f.idx];
        return cell && cell.textContent.toLowerCase().indexOf(f.lower) !== -1;
      });
      row.style.display = show ? '' : 'none';
    });

    /* ── 3. Update header chrome ──────────────────────────────────────── */
    self.ths.forEach(function (th, idx) {
      if (!th.classList.contains('sf-th')) return;
      var sortDir  = self.sortCol && self.sortCol.idx === idx ? self.sortCol.dir : null;
      var filtered = !!self.filters[idx];

      th.classList.toggle('sf-sorted-asc',  sortDir === 'asc');
      th.classList.toggle('sf-sorted-desc', sortDir === 'desc');
      th.classList.toggle('sf-filtered', filtered);

      var icon = th.querySelector('.sf-icon');
      if (!icon) return;
      if      (sortDir === 'asc'  && filtered) icon.textContent = '▲●';
      else if (sortDir === 'desc' && filtered) icon.textContent = '▼●';
      else if (sortDir === 'asc')              icon.textContent = '▲';
      else if (sortDir === 'desc')             icon.textContent = '▼';
      else if (filtered)                       icon.textContent = '●';
      else                                     icon.textContent = '▾';
    });

    /* ── 4. Update status bar ─────────────────────────────────────────── */
    var fCount = Object.keys(self.filters).filter(function (k) { return self.filters[k]; }).length;
    if ((self.sortCol || fCount > 0) && self.bar) {
      var parts = [];
      if (self.sortCol) {
        var sTh  = self.ths[self.sortCol.idx];
        var sLbl = sTh ? sTh.textContent.replace(/[▲▼▾●]/g, '').trim() : '';
        parts.push('Sorted by "' + sLbl + '" ' + (self.sortCol.dir === 'asc' ? '↑' : '↓'));
      }
      if (fCount > 0)
        parts.push(fCount + (fCount === 1 ? ' column filter' : ' column filters') + ' active');
      self.bar.querySelector('.sf-status-text').textContent = parts.join(' · ');
      self.bar.style.display = 'flex';
    } else if (self.bar) {
      self.bar.style.display = 'none';
    }
  };

  /* ── Bootstrap ──────────────────────────────────────────────────────── */
  document.addEventListener('DOMContentLoaded', function () {
    document.querySelectorAll('[data-sortfilter]').forEach(function (tbl) {
      new SortFilter(tbl);
    });
    /* Close panel on outside click or Escape */
    document.addEventListener('click', function () { closePanel(); });
    document.addEventListener('keydown', function (e) {
      if (e.key === 'Escape') closePanel();
    });
    /* Close panel if the page scrolls (panel is fixed, header may move) */
    document.addEventListener('scroll', function () { closePanel(); }, true);
  });
}());
