/* ── AI Impact on Industries — Spreadsheet Viewer ─────── */
(function () {
  "use strict";

  const XLSX_FILE = "AI_Impact_on_Industries_Data.xlsx";
  const DEFAULT_COL_W = 100;
  const DEFAULT_ROW_H = 26;
  const HEADER_H = 28;
  const COL_W_SCALE = 8; // approximate px per Excel width unit

  let DATA = null;
  let currentSheet = null;
  let selectedCell = null; // {r, c}
  let searchMatches = [];
  let searchIdx = -1;
  let cellElements = {};  // "r,c" -> DOM element
  let colWidths = [];     // px per column (1-indexed)
  let rowHeights = [];    // px per row (1-indexed)
  let colOffsets = [];    // cumulative X offsets
  let rowOffsets = [];    // cumulative Y offsets

  /* ── DOM refs ────────────────────────────────────────── */
  const $ = (s) => document.querySelector(s);
  const colHeadersEl = $("#col-headers");
  const rowHeadersEl = $("#row-headers");
  const gridEl = $("#grid");
  const gridWrapper = $("#grid-wrapper");
  const cellRefEl = $("#cell-ref");
  const cellValueEl = $("#cell-value-display");
  const sheetTabsEl = $("#sheet-tabs");
  const sheetInfoEl = $("#sheet-info");
  const searchInput = $("#search-input");
  const searchCount = $("#search-count");

  /* ── Load Data ───────────────────────────────────────── */
  fetch("data.json")
    .then((r) => r.json())
    .then((data) => {
      DATA = data;
      buildSheetTabs();
      switchSheet(DATA.sheetNames[0]);
    });

  /* ── Sheet Tabs ──────────────────────────────────────── */
  function buildSheetTabs() {
    sheetTabsEl.innerHTML = "";
    DATA.sheetNames.forEach((name) => {
      const btn = document.createElement("button");
      btn.className = "sheet-tab";
      btn.textContent = name;
      btn.addEventListener("click", () => switchSheet(name));
      sheetTabsEl.appendChild(btn);
    });
  }

  function switchSheet(name) {
    currentSheet = name;
    // Clear search
    searchInput.value = "";
    searchMatches = [];
    searchIdx = -1;
    searchCount.textContent = "";

    // Update tab active state
    document.querySelectorAll(".sheet-tab").forEach((t) => {
      t.classList.toggle("active", t.textContent === name);
    });

    renderSheet();
    selectCell(1, 1);
  }

  /* ── Render Sheet ────────────────────────────────────── */
  function renderSheet() {
    const sheet = DATA.sheets[currentSheet];
    const maxR = sheet.maxRow;
    const maxC = sheet.maxCol;

    // Compute column widths
    colWidths = [0]; // index 0 unused
    for (let c = 1; c <= maxC; c++) {
      const ew = sheet.colWidths[String(c)];
      colWidths[c] = ew ? Math.max(50, Math.round(ew * COL_W_SCALE)) : DEFAULT_COL_W;
    }

    // Compute row heights
    rowHeights = [0];
    for (let r = 1; r <= maxR; r++) {
      const eh = sheet.rowHeights[String(r)];
      rowHeights[r] = eh ? Math.max(20, Math.round(eh * 1.1)) : DEFAULT_ROW_H;
    }

    // Cumulative offsets
    colOffsets = [0, 0];
    for (let c = 2; c <= maxC; c++) {
      colOffsets[c] = colOffsets[c - 1] + colWidths[c - 1];
    }
    rowOffsets = [0, 0];
    for (let r = 2; r <= maxR; r++) {
      rowOffsets[r] = rowOffsets[r - 1] + rowHeights[r - 1];
    }

    const totalW = colOffsets[maxC] + colWidths[maxC];
    const totalH = rowOffsets[maxR] + rowHeights[maxR];

    // Build merged cell lookup
    const mergedCoverage = {}; // "r,c" -> merge owner key or "hidden"
    const mergeMap = {};       // owner "r,c" -> {colspan, rowspan}
    sheet.merges.forEach((m) => {
      const ownerKey = `${m.startRow},${m.startCol}`;
      mergeMap[ownerKey] = {
        colspan: m.endCol - m.startCol + 1,
        rowspan: m.endRow - m.startRow + 1,
        endRow: m.endRow,
        endCol: m.endCol,
      };
      for (let r = m.startRow; r <= m.endRow; r++) {
        for (let c = m.startCol; c <= m.endCol; c++) {
          if (r === m.startRow && c === m.startCol) continue;
          mergedCoverage[`${r},${c}`] = "hidden";
        }
      }
    });

    // Column headers
    colHeadersEl.innerHTML = "";
    for (let c = 1; c <= maxC; c++) {
      const div = document.createElement("div");
      div.className = "col-header";
      div.style.width = colWidths[c] + "px";
      div.textContent = colLetter(c);
      div.dataset.col = c;
      colHeadersEl.appendChild(div);
    }

    // Row headers
    rowHeadersEl.innerHTML = "";
    for (let r = 1; r <= maxR; r++) {
      const div = document.createElement("div");
      div.className = "row-header";
      div.style.height = rowHeights[r] + "px";
      div.textContent = r;
      div.dataset.row = r;
      rowHeadersEl.appendChild(div);
    }

    // Grid
    gridEl.innerHTML = "";
    gridEl.style.width = totalW + "px";
    gridEl.style.height = totalH + "px";
    cellElements = {};

    for (let r = 1; r <= maxR; r++) {
      for (let c = 1; c <= maxC; c++) {
        const key = `${r},${c}`;
        if (mergedCoverage[key] === "hidden") continue;

        const cellData = sheet.cells[key];
        const div = document.createElement("div");
        div.className = "cell";

        let w = colWidths[c];
        let h = rowHeights[r];
        if (mergeMap[key]) {
          const mg = mergeMap[key];
          w = 0;
          for (let cc = c; cc <= mg.endCol; cc++) w += colWidths[cc];
          h = 0;
          for (let rr = r; rr <= mg.endRow; rr++) h += rowHeights[rr];
        }

        div.style.left = colOffsets[c] + "px";
        div.style.top = rowOffsets[r] + "px";
        div.style.width = w + "px";
        div.style.height = h + "px";

        if (cellData) {
          div.textContent = formatCellValue(cellData);
          applyCellStyle(div, cellData);
        }

        div.dataset.row = r;
        div.dataset.col = c;
        div.addEventListener("click", () => selectCell(r, c));

        gridEl.appendChild(div);
        cellElements[key] = div;
      }
    }

    // Sheet info
    sheetInfoEl.textContent = `${maxR} rows × ${maxC} cols · ${Object.keys(sheet.cells).length} cells`;

    // Sync scroll
    gridWrapper.onscroll = syncScroll;
  }

  /* ── Format cell value ───────────────────────────────── */
  function formatCellValue(cd) {
    if (cd.v === null || cd.v === undefined) return "";
    const val = cd.v;
    const nf = cd.nf || "";

    if (cd.type === "number") {
      // Percentage
      if (nf.includes("%")) {
        return (val * 100).toFixed(nf.includes("0.0%") ? 1 : 0) + "%";
      }
      // Currency
      if (nf.includes("$") || nf.includes("€")) {
        const prefix = nf.includes("$") ? "$" : "€";
        return prefix + commify(val, getDecimalPlaces(nf));
      }
      // Comma formatting
      if (nf.includes(",") || nf.includes("#,")) {
        return commify(val, getDecimalPlaces(nf));
      }
      // Plain number with decimal
      const dp = getDecimalPlaces(nf);
      if (dp > 0) return commify(val, dp);
      // Integer-ish
      if (Number.isInteger(val)) return commify(val, 0);
      return commify(val, 2);
    }
    return String(val);
  }

  function getDecimalPlaces(nf) {
    const m = nf.match(/\.(0+)/);
    return m ? m[1].length : 0;
  }

  function commify(n, dp) {
    const parts = Number(n).toFixed(dp).split(".");
    parts[0] = parts[0].replace(/\B(?=(\d{3})+(?!\d))/g, ",");
    return parts.join(".");
  }

  /* ── Apply cell styling ──────────────────────────────── */
  function applyCellStyle(el, cd) {
    // Type class
    if (cd.type === "number") {
      el.classList.add("num");
      // Percentage coloring
      const nf = cd.nf || "";
      if (nf.includes("%") && typeof cd.v === "number") {
        el.classList.add(cd.v >= 0 ? "pct-positive" : "pct-negative");
      }
    } else {
      el.classList.add("str");
    }

    // Font
    if (cd.font) {
      if (cd.font.b) el.style.fontWeight = "700";
      if (cd.font.i) el.style.fontStyle = "italic";
      if (cd.font.u) el.style.textDecoration = "underline";
      if (cd.font.sz) el.style.fontSize = Math.min(cd.font.sz, 16) + "px";
      if (cd.font.color) el.style.color = remapColor(cd.font.color);
    }

    // Background
    if (cd.bg) {
      el.style.backgroundColor = cd.bg;
    }

    // Alignment
    if (cd.align) {
      if (cd.align.h === "center") {
        el.style.justifyContent = "center";
        el.style.textAlign = "center";
      } else if (cd.align.h === "right") {
        el.style.justifyContent = "flex-end";
        el.style.textAlign = "right";
      } else if (cd.align.h === "left") {
        el.style.justifyContent = "flex-start";
        el.style.textAlign = "left";
      }
      if (cd.align.wrap) {
        el.style.whiteSpace = "normal";
        el.style.wordBreak = "break-word";
      }
    }

    // Borders
    if (cd.border) {
      if (cd.border.bottom) el.style.borderBottom = "1px solid var(--border-lt)";
      if (cd.border.right) el.style.borderRight = "1px solid var(--border-lt)";
      if (cd.border.top) el.style.borderTop = "1px solid var(--border-lt)";
      if (cd.border.left) el.style.borderLeft = "1px solid var(--border-lt)";
    }
  }

  /* ── Cell Selection ──────────────────────────────────── */
  function selectCell(r, c) {
    // Remove old selection
    if (selectedCell) {
      const old = cellElements[`${selectedCell.r},${selectedCell.c}`];
      if (old) old.classList.remove("selected");
      document.querySelectorAll(".col-header.active, .row-header.active").forEach((e) =>
        e.classList.remove("active")
      );
    }

    selectedCell = { r, c };

    // Highlight new
    const el = cellElements[`${r},${c}`];
    if (el) el.classList.add("selected");

    // Header highlights
    const ch = colHeadersEl.querySelector(`[data-col="${c}"]`);
    if (ch) ch.classList.add("active");
    const rh = rowHeadersEl.querySelector(`[data-row="${r}"]`);
    if (rh) rh.classList.add("active");

    // Formula bar
    cellRefEl.textContent = colLetter(c) + r;
    const sheet = DATA.sheets[currentSheet];
    const cd = sheet.cells[`${r},${c}`];
    cellValueEl.textContent = cd ? formatCellValue(cd) : "";
  }

  /* ── Scroll Sync ─────────────────────────────────────── */
  function syncScroll() {
    const sl = gridWrapper.scrollLeft;
    const st = gridWrapper.scrollTop;
    document.getElementById("col-headers-wrapper").scrollLeft = sl;
    document.getElementById("row-headers-wrapper").scrollTop = st;
  }

  /* ── Search ──────────────────────────────────────────── */
  searchInput.addEventListener("input", runSearch);
  $("#search-next").addEventListener("click", () => navigateSearch(1));
  $("#search-prev").addEventListener("click", () => navigateSearch(-1));
  $("#search-clear").addEventListener("click", () => {
    searchInput.value = "";
    runSearch();
    searchInput.focus();
  });

  // Keyboard shortcuts
  searchInput.addEventListener("keydown", (e) => {
    if (e.key === "Enter") {
      navigateSearch(e.shiftKey ? -1 : 1);
      e.preventDefault();
    }
    if (e.key === "Escape") {
      searchInput.value = "";
      runSearch();
    }
  });

  document.addEventListener("keydown", (e) => {
    // Ctrl/Cmd+F to focus search
    if ((e.ctrlKey || e.metaKey) && e.key === "f") {
      e.preventDefault();
      searchInput.focus();
      searchInput.select();
    }
    // Arrow key navigation
    if (!selectedCell || document.activeElement === searchInput) return;
    const sheet = DATA.sheets[currentSheet];
    let { r, c } = selectedCell;
    if (e.key === "ArrowDown") { r = Math.min(r + 1, sheet.maxRow); e.preventDefault(); }
    else if (e.key === "ArrowUp") { r = Math.max(r - 1, 1); e.preventDefault(); }
    else if (e.key === "ArrowRight" || e.key === "Tab") { c = Math.min(c + 1, sheet.maxCol); e.preventDefault(); }
    else if (e.key === "ArrowLeft") { c = Math.max(c - 1, 1); e.preventDefault(); }
    else return;
    selectCell(r, c);
    scrollCellIntoView(r, c);
  });

  function runSearch() {
    // Clear old highlights
    searchMatches.forEach((key) => {
      const el = cellElements[key];
      if (el) {
        el.classList.remove("search-match");
        el.classList.remove("search-active");
      }
    });
    searchMatches = [];
    searchIdx = -1;

    const q = searchInput.value.trim().toLowerCase();
    if (!q) {
      searchCount.textContent = "";
      return;
    }

    const sheet = DATA.sheets[currentSheet];
    Object.keys(sheet.cells).forEach((key) => {
      const cd = sheet.cells[key];
      const display = formatCellValue(cd).toLowerCase();
      const raw = cd.v !== null && cd.v !== undefined ? String(cd.v).toLowerCase() : "";
      if (display.includes(q) || raw.includes(q)) {
        searchMatches.push(key);
        const el = cellElements[key];
        if (el) el.classList.add("search-match");
      }
    });

    // Sort matches by row,col
    searchMatches.sort((a, b) => {
      const [ar, ac] = a.split(",").map(Number);
      const [br, bc] = b.split(",").map(Number);
      return ar - br || ac - bc;
    });

    searchCount.textContent = searchMatches.length > 0 ? `${searchMatches.length} found` : "0 found";

    if (searchMatches.length > 0) navigateSearch(1);
  }

  function navigateSearch(dir) {
    if (searchMatches.length === 0) return;

    // Remove active from current
    if (searchIdx >= 0 && searchIdx < searchMatches.length) {
      const el = cellElements[searchMatches[searchIdx]];
      if (el) el.classList.remove("search-active");
    }

    searchIdx += dir;
    if (searchIdx >= searchMatches.length) searchIdx = 0;
    if (searchIdx < 0) searchIdx = searchMatches.length - 1;

    const key = searchMatches[searchIdx];
    const el = cellElements[key];
    if (el) {
      el.classList.add("search-active");
      const [r, c] = key.split(",").map(Number);
      selectCell(r, c);
      scrollCellIntoView(r, c);
    }

    searchCount.textContent = `${searchIdx + 1} / ${searchMatches.length}`;
  }

  function scrollCellIntoView(r, c) {
    const x = colOffsets[c];
    const y = rowOffsets[r];
    const w = colWidths[c];
    const h = rowHeights[r];
    const vw = gridWrapper.clientWidth;
    const vh = gridWrapper.clientHeight;
    const sl = gridWrapper.scrollLeft;
    const st = gridWrapper.scrollTop;

    if (x < sl) gridWrapper.scrollLeft = x;
    else if (x + w > sl + vw) gridWrapper.scrollLeft = x + w - vw;

    if (y < st) gridWrapper.scrollTop = y;
    else if (y + h > st + vh) gridWrapper.scrollTop = y + h - vh;
  }

  /* ── Download ────────────────────────────────────────── */
  $("#download-btn").addEventListener("click", () => {
    const a = document.createElement("a");
    a.href = XLSX_FILE;
    a.download = XLSX_FILE;
    a.click();
  });

  /* ── Helpers ─────────────────────────────────────────── */
  /* Remap dark Excel font colors to readable light colors */
  function remapColor(hex) {
    if (!hex) return null;
    const r = parseInt(hex.slice(1, 3), 16);
    const g = parseInt(hex.slice(3, 5), 16);
    const b = parseInt(hex.slice(5, 7), 16);
    const lum = (0.299 * r + 0.587 * g + 0.114 * b);
    // Dark colors become white; mid-dark become light blue
    if (lum < 60) return "#e2eafc";   // very dark → soft white-blue
    if (lum < 120) return "#89b4fa";  // mid-dark → light blue
    return hex; // already light enough
  }

  function colLetter(c) {
    let s = "";
    while (c > 0) {
      c--;
      s = String.fromCharCode(65 + (c % 26)) + s;
      c = Math.floor(c / 26);
    }
    return s;
  }
})();
