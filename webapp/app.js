(function () {
  "use strict";

  /** @typedef {{ id: string, label: string, kind: 'no'|'field' }} ColumnDef */
  /** @typedef {{ mode: 'text'|'aria'|'attr', attrName?: string, relativePath: string }} Mapping */

  const DEFAULT_COLUMNS = /** @type {ColumnDef[]} */ ([
    { id: "no", label: "No.", kind: "no" },
    { id: "name", label: "名前", kind: "field" },
    { id: "rating", label: "評価", kind: "field" },
    { id: "date", label: "いつ頃", kind: "field" },
    { id: "comment", label: "コメント", kind: "field" },
  ]);

  let columns = DEFAULT_COLUMNS.map((c) => ({ ...c }));
  /** @type {Record<string, Mapping|null>} */
  let mappings = {};
  columns.forEach((c) => {
    if (c.kind === "field") mappings[c.id] = null;
  });

  let customColumnCounter = 0;

  let previewRoot = null;
  let previewInner = null;
  let lastDoc = null;
  let lastContainerSelector = "";

  const htmlInput = document.getElementById("htmlInput");
  const containerSelector = document.getElementById("containerSelector");
  const btnPreview = document.getElementById("btnPreview");
  const parseError = document.getElementById("parseError");
  const previewHost = document.getElementById("previewHost");
  const columnList = document.getElementById("columnList");
  const btnAddColumn = document.getElementById("btnAddColumn");
  const btnDownload = document.getElementById("btnDownload");
  const includeBom = document.getElementById("includeBom");
  const extractInfo = document.getElementById("extractInfo");

  const previewPadHost = document.getElementById("previewPadHost");
  const previewPadRoot = document.getElementById("previewPadRoot");
  const previewPadHostVal = document.getElementById("previewPadHostVal");
  const previewPadRootVal = document.getElementById("previewPadRootVal");

  const LS_PAD_HOST = "googleReviewPreviewPadHost";
  const LS_PAD_ROOT = "googleReviewPreviewPadRoot";

  const modal = document.getElementById("modal");
  const modalSnippet = document.getElementById("modalSnippet");
  const modalColumn = document.getElementById("modalColumn");
  const modalMode = document.getElementById("modalMode");
  const modalAttrName = document.getElementById("modalAttrName");
  const attrRow = document.getElementById("attrRow");
  const modalSave = document.getElementById("modalSave");

  /** @type {{ path: string, element: Element } | null} */
  let pendingPick = null;

  function showError(msg) {
    parseError.hidden = !msg;
    parseError.textContent = msg || "";
  }

  /**
   * コンテナ root から el までの nth-child 連結セレクタ（プレビュー内で一意）
   * @param {Element} root
   * @param {Element} el
   */
  function relativePathFromRoot(root, el) {
    if (!root.contains(el) || el === root) return "";
    const parts = [];
    let node = el;
    while (node && node !== root) {
      const parent = node.parentElement;
      if (!parent) return null;
      const index = Array.prototype.indexOf.call(parent.children, node) + 1;
      const tag = node.tagName.toLowerCase();
      parts.unshift(tag + ":nth-child(" + index + ")");
      node = parent;
    }
    return parts.join(" > ");
  }

  /**
   * @param {Element} root
   * @param {string} path
   */
  function queryByRelativePath(root, path) {
    if (!path) return root;
    return root.querySelector(path);
  }

  /**
   * @param {Element} el
   * @param {Mapping} map
   */
  function extractValue(el, map) {
    if (!el) return "";
    if (map.mode === "text") return (el.textContent || "").trim();
    if (map.mode === "aria") return (el.getAttribute("aria-label") || "").trim();
    if (map.mode === "attr") {
      const name = (map.attrName || "").trim();
      if (!name) return "";
      return (el.getAttribute(name) || "").trim();
    }
    return "";
  }

  function parseHtmlToDocument(html) {
    const parser = new DOMParser();
    const doc = parser.parseFromString(html, "text/html");
    const p = doc.querySelector("parsererror");
    if (p) throw new Error("HTMLの解析に失敗しました。タグの欠けがないか確認してください。");
    return doc;
  }

  function buildPreview() {
    showError("");
    previewHost.innerHTML = "";
    previewHost.removeAttribute("data-empty");
    previewRoot = null;
    previewInner = null;
    lastDoc = null;
    btnDownload.disabled = true;
    extractInfo.textContent = "";

    const raw = htmlInput.value.trim();
    if (!raw) {
      showError("HTMLを貼り付けてください。");
      previewHost.setAttribute("data-empty", "");
      previewHost.innerHTML =
        '<p class="preview-placeholder">HTMLを貼り付けてから「プレビュー生成」を押してください。</p>';
      return;
    }

    const sel = containerSelector.value.trim();
    if (!sel) {
      showError("口コミ1件のコンテナーのCSSセレクターを入力してください。");
      previewHost.setAttribute("data-empty", "");
      previewHost.innerHTML = '<p class="preview-placeholder">セレクターを入力してください。</p>';
      return;
    }

    let doc;
    try {
      doc = parseHtmlToDocument(raw);
    } catch (e) {
      showError(e.message || String(e));
      previewHost.setAttribute("data-empty", "");
      return;
    }

    let nodes;
    try {
      nodes = doc.querySelectorAll(sel);
    } catch (e) {
      showError("セレクターが不正です: " + (e.message || String(e)));
      previewHost.setAttribute("data-empty", "");
      return;
    }

    if (!nodes.length) {
      showError("セレクターに一致する要素がありません。セレクターを見直してください。");
      previewHost.setAttribute("data-empty", "");
      previewHost.innerHTML =
        '<p class="preview-placeholder">一致する要素がありません。</p>';
      return;
    }

    const first = nodes[0];
    lastDoc = doc;
    lastContainerSelector = sel;

    const inner = document.createElement("div");
    inner.className = "preview-inner";
    inner.appendChild(first.cloneNode(true));

    const wrap = document.createElement("div");
    wrap.className = "preview-wrap";
    wrap.appendChild(inner);
    previewHost.appendChild(wrap);

    previewInner = inner;
    previewRoot = inner.firstElementChild;
    if (!previewRoot) {
      showError("最初の口コミ枠に要素がありません。");
      return;
    }

    inner.addEventListener("click", onPreviewClick);

    refreshMappedHighlights();
    applyPreviewPadding();
    btnDownload.disabled = false;
    extractInfo.textContent =
      "一致件数: " + nodes.length + " 件（CSVは全件に適用されます）";
    renderColumnList();
  }

  function applyPreviewPadding() {
    const hostPx = Math.max(0, Math.min(48, parseInt(previewPadHost.value, 10) || 0));
    const rootPx = Math.max(0, Math.min(48, parseInt(previewPadRoot.value, 10) || 0));
    previewPadHostVal.textContent = String(hostPx);
    previewPadRootVal.textContent = String(rootPx);
    previewHost.style.padding = hostPx + "px";
    if (previewRoot && previewRoot.style) {
      previewRoot.style.boxSizing = "border-box";
      previewRoot.style.padding = rootPx > 0 ? rootPx + "px" : "";
    }
  }

  function initPreviewPaddingControls() {
    try {
      const h = localStorage.getItem(LS_PAD_HOST);
      const r = localStorage.getItem(LS_PAD_ROOT);
      if (h !== null && h !== "") previewPadHost.value = h;
      if (r !== null && r !== "") previewPadRoot.value = r;
    } catch (_) {}

    function onChange() {
      applyPreviewPadding();
      try {
        localStorage.setItem(LS_PAD_HOST, previewPadHost.value);
        localStorage.setItem(LS_PAD_ROOT, previewPadRoot.value);
      } catch (_) {}
    }

    previewPadHost.addEventListener("input", onChange);
    previewPadRoot.addEventListener("input", onChange);
    applyPreviewPadding();
  }

  /** @param {MouseEvent} e */
  function onPreviewClick(e) {
    e.preventDefault();
    e.stopPropagation();
    let el = /** @type {Node} */ (e.target);
    if (el.nodeType !== Node.ELEMENT_NODE) el = el.parentElement;
    if (!el || !previewRoot || !previewInner.contains(el)) return;

    const path = relativePathFromRoot(previewRoot, el);
    if (path === null) return;

    pendingPick = { path, element: el };
    const outer = el.cloneNode(false);
    modalSnippet.textContent = path + " — <" + outer.outerHTML.slice(0, 200) + ">";

    fillModalColumnOptions();
    modalColumn.value = guessColumnId(el);
    modalMode.value = guessMode(el);
    syncAttrRow();
    modal.removeAttribute("hidden");
    modalSave.focus();
  }

  function guessColumnId(el) {
    const al = (el.getAttribute("aria-label") || "").toLowerCase();
    if (al.includes("star") || al.includes("つ星") || al.includes("星")) return "rating";
    const fieldIds = columns.filter((c) => c.kind === "field").map((c) => c.id);
    for (const id of fieldIds) {
      if (mappings[id] && mappings[id].relativePath === relativePathFromRoot(previewRoot, el))
        return id;
    }
    return fieldIds[0] || "name";
  }

  function guessMode(el) {
    if (el.getAttribute("aria-label")) return "aria";
    return "text";
  }

  function fillModalColumnOptions() {
    modalColumn.innerHTML = "";
    columns
      .filter((c) => c.kind === "field")
      .forEach((c) => {
        const opt = document.createElement("option");
        opt.value = c.id;
        opt.textContent = c.label;
        modalColumn.appendChild(opt);
      });
  }

  function syncAttrRow() {
    const show = modalMode.value === "attr";
    attrRow.hidden = !show;
  }

  modalMode.addEventListener("change", syncAttrRow);

  function closeModal() {
    modal.setAttribute("hidden", "");
    pendingPick = null;
  }

  document.addEventListener("keydown", (e) => {
    if (e.key === "Escape" && !modal.hasAttribute("hidden")) {
      closeModal();
    }
  });

  modal.querySelectorAll("[data-close]").forEach((n) => {
    n.addEventListener("click", closeModal);
  });

  modalSave.addEventListener("click", () => {
    if (!pendingPick || !previewRoot) {
      closeModal();
      return;
    }
    const colId = modalColumn.value;
    const mode = /** @type {'text'|'aria'|'attr'} */ (modalMode.value);
    const attrName = mode === "attr" ? modalAttrName.value.trim() : undefined;

    const map = /** @type {Mapping} */ ({
      relativePath: pendingPick.path,
      mode,
      attrName: attrName || undefined,
    });

    mappings[colId] = map;
    closeModal();
    refreshMappedHighlights();
    renderColumnList();
  });

  function refreshMappedHighlights() {
    if (!previewInner || !previewRoot) return;
    previewInner.querySelectorAll("[data-mapped]").forEach((n) => {
      n.removeAttribute("data-mapped");
    });
    columns.forEach((c) => {
      if (c.kind !== "field") return;
      const m = mappings[c.id];
      if (!m || !m.relativePath) return;
      const el = queryByRelativePath(previewRoot, m.relativePath);
      if (el && el instanceof HTMLElement) {
        el.setAttribute("data-mapped", "1");
      }
    });
  }

  function renderColumnList() {
    columnList.innerHTML = "";
    columns.forEach((col) => {
      const li = document.createElement("li");
      li.className = "column-item";
      li.draggable = true;
      li.dataset.columnId = col.id;

      const handle = document.createElement("span");
      handle.className = "column-item__handle";
      handle.textContent = "≡";
      handle.title = "ドラッグで並べ替え";

      const label = document.createElement("span");
      label.className = "column-item__label";
      label.textContent = col.label + (col.kind === "no" ? "（自動）" : "");

      const meta = document.createElement("span");
      meta.className = "column-item__meta";
      if (col.kind === "field") {
        const m = mappings[col.id];
        meta.textContent = m
          ? m.mode === "text"
            ? "text · " + m.relativePath
            : m.mode === "aria"
              ? "aria-label · " + m.relativePath
              : (m.attrName || "?") + " · " + m.relativePath
          : "未設定";
      } else {
        meta.textContent = "1, 2, 3…";
      }

      li.appendChild(handle);
      li.appendChild(label);
      li.appendChild(meta);

      if (col.kind === "field") {
        const clearBtn = document.createElement("button");
        clearBtn.type = "button";
        clearBtn.className = "btn btn--ghost column-item__clear";
        clearBtn.textContent = "解除";
        clearBtn.addEventListener("click", () => {
          mappings[col.id] = null;
          refreshMappedHighlights();
          renderColumnList();
        });
        li.appendChild(clearBtn);
      }

      columnList.appendChild(li);
    });

    bindDragSort();
  }

  function bindDragSort() {
    let dragEl = null;
    const items = columnList.querySelectorAll(".column-item");

    items.forEach((item) => {
      item.addEventListener("dragstart", (e) => {
        dragEl = item;
        item.classList.add("dragging");
        e.dataTransfer.effectAllowed = "move";
        e.dataTransfer.setData("text/plain", item.dataset.columnId || "");
      });

      item.addEventListener("dragend", () => {
        item.classList.remove("dragging");
        dragEl = null;
        applyColumnOrderFromDom();
      });

      item.addEventListener("dragover", (e) => {
        e.preventDefault();
        e.dataTransfer.dropEffect = "move";
        if (!dragEl || dragEl === item) return;
        const rect = item.getBoundingClientRect();
        const mid = rect.top + rect.height / 2;
        if (e.clientY < mid) {
          columnList.insertBefore(dragEl, item);
        } else {
          columnList.insertBefore(dragEl, item.nextSibling);
        }
      });

      item.addEventListener("drop", (e) => {
        e.preventDefault();
      });
    });
  }

  function applyColumnOrderFromDom() {
    const ids = Array.prototype.map.call(
      columnList.querySelectorAll(".column-item"),
      (li) => li.dataset.columnId
    );
    const byId = new Map(columns.map((c) => [c.id, c]));
    columns = ids.map((id) => byId.get(id)).filter(Boolean);
    renderColumnList();
  }

  btnAddColumn.addEventListener("click", () => {
    customColumnCounter += 1;
    const id = "custom_" + customColumnCounter;
    columns.push({ id, label: "カスタム" + customColumnCounter, kind: "field" });
    mappings[id] = null;
    renderColumnList();
  });

  btnPreview.addEventListener("click", buildPreview);

  btnDownload.addEventListener("click", () => {
    if (!lastDoc || !lastContainerSelector) {
      buildPreview();
      if (!lastDoc) return;
    }

    let items;
    try {
      items = lastDoc.querySelectorAll(lastContainerSelector);
    } catch (e) {
      showError("CSV出力時にセレクターエラー: " + (e.message || String(e)));
      return;
    }

    const fieldColumns = columns.filter((c) => c.kind === "field");
    const headers = columns.map((c) => c.label);

    const rows = [];
    items.forEach((item, rowIdx) => {
      const row = [];
      columns.forEach((col) => {
        if (col.kind === "no") {
          row.push(String(rowIdx + 1));
          return;
        }
        const m = mappings[col.id];
        if (!m || !m.relativePath) {
          row.push("");
          return;
        }
        const el = queryByRelativePath(item, m.relativePath);
        row.push(extractValue(el, m));
      });
      rows.push(row);
    });

    const csv = rowsToCsv(headers, rows, includeBom.checked);
    const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = "reviews-" + Date.now() + ".csv";
    a.click();
    URL.revokeObjectURL(a.href);
  });

  /**
   * @param {string[]} headers
   * @param {string[][]} rows
   * @param {boolean} bom
   */
  function rowsToCsv(headers, rows, bom) {
    const esc = (s) => {
      const t = String(s);
      if (/[",\r\n]/.test(t)) return '"' + t.replace(/"/g, '""') + '"';
      return t;
    };
    const lines = [headers.map(esc).join(",")];
    rows.forEach((r) => lines.push(r.map(esc).join(",")));
    const body = lines.join("\r\n");
    return bom ? "\uFEFF" + body : body;
  }

  initPreviewPaddingControls();
  renderColumnList();
})();
