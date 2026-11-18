(() => {
  'use strict';

  if (window.__dtableInitialized) {
    return;
  }
  window.__dtableInitialized = true;

  const SHADOW_HOST_ID = 'dtable-shadow-host';
  const SCROLL_HIGHLIGHT_CLASS = 'dtable-scroll-highlight';
  const HEADING_SELECTOR = 'h1, h2, h3, h4, h5, h6';
  const ARIA_TABLE_SELECTOR = '[role="table"], [role="grid"], [role="treegrid"]';
  const ARIA_CELL_SELECTOR = '[role="columnheader"], [role="rowheader"], [role="gridcell"], [role="cell"]';
  const ROW_VERTICAL_TOLERANCE_PX = 4;
  const COL_ALIGNMENT_TOLERANCE_PX = 12;
  const SECTION_TAGS = new Set([
    'SECTION',
    'ARTICLE',
    'MAIN',
    'ASIDE',
    'NAV',
    'DIALOG',
    'BODY',
    'HTML'
  ]);
  const BLOCK_TAGS = new Set([
    'ADDRESS',
    'ARTICLE',
    'ASIDE',
    'BLOCKQUOTE',
    'DD',
    'DIV',
    'DL',
    'DT',
    'FIELDSET',
    'FIGCAPTION',
    'FIGURE',
    'FOOTER',
    'FORM',
    'H1',
    'H2',
    'H3',
    'H4',
    'H5',
    'H6',
    'HEADER',
    'HR',
    'LI',
    'MAIN',
    'NAV',
    'OL',
    'P',
    'PRE',
    'SECTION',
    'TABLE',
    'TR',
    'UL'
  ]);
  const SKIP_TAGS = new Set(['SCRIPT', 'STYLE', 'NOSCRIPT', 'TEMPLATE']);

  let lastScannedTables = [];
  let scrollHighlightTimer = null;
  let activeScrollHighlightedEl = null;

  const selectionState = {
    active: false,
    format: 'xlsx',
    host: null,
    shadow: null,
    banner: null,
    highlightBox: null,
    tooltip: null,
    toast: null,
    toastTimer: null,
    hoveredTable: null,
    hoveredMeta: null,
    rafId: null
  };

  /**
   * Clean text inside a DOM node: removes script/style/hidden nodes and normalizes whitespace.
   */
  function cleanNodeText(root) {
    if (!root) return '';

    const chunks = [];

    function walk(node) {
      if (!node) return;
      if (node.nodeType === Node.TEXT_NODE) {
        chunks.push(node.nodeValue || '');
        return;
      }
      if (node.nodeType !== Node.ELEMENT_NODE) {
        return;
      }

      const tag = node.tagName;
      if (SKIP_TAGS.has(tag)) {
        return;
      }
      if (node.hidden || (node.style && (node.style.display === 'none' || node.style.visibility === 'hidden'))) {
        return;
      }
      if (tag === 'BR') {
        chunks.push('\n');
        return;
      }

      const isBlock = BLOCK_TAGS.has(tag);
      if (isBlock) {
        chunks.push('\n');
      }

      const children = node.childNodes;
      for (let i = 0; i < children.length; i += 1) {
        walk(children[i]);
      }

      if (isBlock) {
        chunks.push('\n');
      }
    }

    walk(root);

    return chunks
      .join('')
      .replace(/\u00A0/g, ' ')
      .replace(/[\u200B-\u200D\uFEFF]/g, '')
      .split('\n')
      .map((line) => line.replace(/[ \t\f\v\r]+/g, ' ').trim())
      .filter((line, idx, arr) => line !== '' || (idx > 0 && arr[idx - 1] !== ''))
      .join('\n')
      .trim();
  }

  /**
   * Format an element ID attribute into readable title-cased words.
   */
  function formatIdToWords(rawId) {
    if (!rawId || typeof rawId !== 'string') return '';
    const spaced = rawId
      .replace(/[-_]+/g, ' ')
      .replace(/([a-z0-9])([A-Z])/g, '$1 $2')
      .replace(/([A-Z]+)([A-Z][a-z])/g, '$1 $2')
      .replace(/\s+/g, ' ')
      .trim();

    if (!spaced) return '';
    return spaced
      .split(' ')
      .map((word) => word.charAt(0).toUpperCase() + word.slice(1))
      .join(' ');
  }

  /**
   * Check whether a container marks a semantic section boundary.
   */
  function isSectionBoundary(el) {
    if (!el || el.nodeType !== Node.ELEMENT_NODE) return true;
    if (SECTION_TAGS.has(el.tagName)) return true;
    const role = (el.getAttribute('role') || '').toLowerCase();
    return role === 'region' || role === 'tabpanel' || role === 'dialog' || role === 'main';
  }

  /**
   * Find the nearest preceding heading (h1-h6) within the same parent section or container.
   */
  function findNearestPrecedingHeading(table) {
    let current = table;
    let depth = 0;

    while (current && current.parentElement && depth < 5) {
      const parent = current.parentElement;
      let prev = current.previousElementSibling;

      while (prev) {
        if (prev.matches && prev.matches(HEADING_SELECTOR)) {
          const text = cleanNodeText(prev);
          if (text) return text;
        }

        if (
          prev.tagName === 'TABLE' ||
          (prev.matches && prev.matches(ARIA_TABLE_SELECTOR)) ||
          (prev.querySelector && prev.querySelector(`table, ${ARIA_TABLE_SELECTOR}`))
        ) {
          break;
        }

        if (prev.querySelectorAll) {
          const headings = prev.querySelectorAll(HEADING_SELECTOR);
          for (let i = headings.length - 1; i >= 0; i -= 1) {
            const text = cleanNodeText(headings[i]);
            if (text) return text;
          }
        }

        prev = prev.previousElementSibling;
      }

      if (isSectionBoundary(parent)) {
        break;
      }

      current = parent;
      depth += 1;
    }

    return '';
  }

  /**
   * Resolve a table's display/sheet name using the required 6-tier priority order.
   */
  function resolveTableName(table, fallbackIndex) {
    // 1. Text inside <caption>
    const caption = table.caption || table.querySelector(':scope > caption');
    if (caption) {
      const captionText = cleanNodeText(caption);
      if (captionText) return captionText;
    }

    // 2. aria-label or title attribute on <table>
    const ariaLabel = (table.getAttribute('aria-label') || '').replace(/\s+/g, ' ').trim();
    if (ariaLabel) return ariaLabel;

    const titleAttr = (table.getAttribute('title') || '').replace(/\s+/g, ' ').trim();
    if (titleAttr) return titleAttr;

    // 3. aria-labelledby referenced element(s)
    const labelledBy = (table.getAttribute('aria-labelledby') || '').trim();
    if (labelledBy) {
      const parts = labelledBy
        .split(/\s+/)
        .map((id) => document.getElementById(id))
        .filter(Boolean)
        .map((el) => cleanNodeText(el))
        .filter(Boolean);
      if (parts.length > 0) {
        return parts.join(' ');
      }
    }

    // 4. Nearest preceding heading (h1-h6) in same parent section/container
    const headingText = findNearestPrecedingHeading(table);
    if (headingText) return headingText;

    // 5. Formatted id attribute
    const idText = formatIdToWords(table.id);
    if (idText) return idText;

    // 6. Fallback
    return `Table ${fallbackIndex}`;
  }

  /**
   * Collect direct rows of a table grouped by section (thead, tbody, tfoot) in visual order.
   */
  function getTableRowSections(table) {
    const sections = [];

    if (table.tHead) {
      const headRows = Array.from(table.tHead.rows).filter((tr) => tr.closest('table') === table);
      if (headRows.length > 0) sections.push(headRows);
    }

    const bodySet = new Set();
    for (const tbody of Array.from(table.tBodies)) {
      const bodyRows = Array.from(tbody.rows).filter((tr) => tr.closest('table') === table);
      if (bodyRows.length > 0) {
        sections.push(bodyRows);
        bodyRows.forEach((tr) => bodySet.add(tr));
      }
    }

    // Include any direct child <tr> elements outside tHead/tBodies/tFoot
    const footRows = table.tFoot
      ? Array.from(table.tFoot.rows).filter((tr) => tr.closest('table') === table)
      : [];
    const headSet = new Set(table.tHead ? Array.from(table.tHead.rows) : []);
    const footSet = new Set(footRows);
    const orphanRows = Array.from(table.children).filter(
      (el) => el.tagName === 'TR' && !headSet.has(el) && !bodySet.has(el) && !footSet.has(el)
    );
    if (orphanRows.length > 0) {
      sections.push(orphanRows);
    }

    if (footRows.length > 0) {
      sections.push(footRows);
    }

    return sections;
  }

  /**
   * Extract normalized 2D cell grid and SheetJS merge ranges from a <table>.
   */
  function extractTableMatrix(table) {
    const sections = getTableRowSections(table);
    const grid = [];
    const occupied = [];
    const merges = [];
    let globalRowIndex = 0;
    let maxCols = 0;
    let hasContent = false;

    const ensureRow = (r) => {
      if (!grid[r]) grid[r] = [];
      if (!occupied[r]) occupied[r] = [];
    };

    for (const sectionRows of sections) {
      const sectionStart = globalRowIndex;
      const sectionEnd = sectionStart + sectionRows.length;

      for (let localR = 0; localR < sectionRows.length; localR += 1) {
        const r = sectionStart + localR;
        const tr = sectionRows[localR];
        ensureRow(r);

        let c = 0;
        const cells = Array.from(tr.cells);

        for (let cellIdx = 0; cellIdx < cells.length; cellIdx += 1) {
          const cell = cells[cellIdx];

          while (occupied[r][c]) {
            c += 1;
          }

          const rawColspan = parseInt(cell.getAttribute('colspan') || '1', 10);
          const colspan = Number.isFinite(rawColspan) && rawColspan > 0 ? rawColspan : 1;

          const rawRowspanAttr = cell.getAttribute('rowspan');
          let rowspan = 1;
          if (rawRowspanAttr === '0') {
            rowspan = sectionEnd - r;
          } else {
            const parsedRowspan = parseInt(rawRowspanAttr || '1', 10);
            rowspan = Number.isFinite(parsedRowspan) && parsedRowspan > 0 ? parsedRowspan : 1;
          }
          const effectiveRowspan = Math.min(rowspan, sectionEnd - r);

          const text = cleanNodeText(cell);
          if (text) {
            hasContent = true;
          }

          for (let dr = 0; dr < effectiveRowspan; dr += 1) {
            const targetRow = r + dr;
            ensureRow(targetRow);
            for (let dc = 0; dc < colspan; dc += 1) {
              const targetCol = c + dc;
              occupied[targetRow][targetCol] = true;
              grid[targetRow][targetCol] = dr === 0 && dc === 0 ? text : '';
            }
          }

          if (effectiveRowspan > 1 || colspan > 1) {
            merges.push({
              s: { r, c },
              e: { r: r + effectiveRowspan - 1, c: c + colspan - 1 }
            });
          }

          c += colspan;
          if (c > maxCols) {
            maxCols = c;
          }
        }

        // Account for trailing spanned columns from previous rows
        if (occupied[r].length > maxCols) {
          maxCols = occupied[r].length;
        }
      }

      globalRowIndex = sectionEnd;
    }

    // Pad all rows to uniform column length
    for (let r = 0; r < grid.length; r += 1) {
      ensureRow(r);
      for (let c = 0; c < maxCols; c += 1) {
        if (grid[r][c] === undefined) {
          grid[r][c] = '';
        }
      }
    }

    return {
      rows: grid,
      merges,
      rowCount: grid.length,
      colCount: maxCols,
      hasContent
    };
  }

  /**
   * Check if a table element is visible in the viewport/layout.
   */
  function isTableVisible(table) {
    if (!table || !table.isConnected) return false;
    if (table.closest('[hidden]')) return false;

    const style = window.getComputedStyle(table);
    if (
      style.display === 'none' ||
      style.visibility === 'hidden' ||
      style.visibility === 'collapse'
    ) {
      return false;
    }

    if (table.getClientRects().length === 0) {
      return false;
    }

    const rect = table.getBoundingClientRect();
    return rect.width > 0 && rect.height > 0;
  }

  /**
   * Check whether a candidate element overlaps (contains or is inside) any already-detected table.
   */
  function overlapsDetectedElement(el, detectedSet) {
    for (const existing of detectedSet) {
      if (existing === el || existing.contains(el) || el.contains(existing)) {
        return true;
      }
    }
    return false;
  }

  /**
   * Build a uniform 2D cell grid and SheetJS merge ranges from row cell arrays.
   */
  function buildMatrixFromCellRows(rowCellLists, getSpanAndColIndex) {
    const grid = [];
    const occupied = [];
    const merges = [];
    const totalRows = rowCellLists.length;
    let maxCols = 0;
    let hasContent = false;

    const ensureRow = (r) => {
      if (!grid[r]) grid[r] = [];
      if (!occupied[r]) occupied[r] = [];
    };

    for (let r = 0; r < totalRows; r += 1) {
      ensureRow(r);
      let c = 0;
      const cells = rowCellLists[r];

      for (let i = 0; i < cells.length; i += 1) {
        const cell = cells[i];
        const info = getSpanAndColIndex ? getSpanAndColIndex(cell) : null;
        const explicitCol = info && info.colIndex >= 0 ? info.colIndex : -1;

        if (explicitCol > c) {
          c = explicitCol;
        }
        while (occupied[r][c]) {
          c += 1;
        }

        const colspan = info && info.colspan > 0 ? info.colspan : 1;
        const rawRowspan = info && info.rowspan > 0 ? info.rowspan : 1;
        const rowspan = Math.min(rawRowspan, totalRows - r);

        const text = cleanNodeText(cell);
        if (text) {
          hasContent = true;
        }

        for (let dr = 0; dr < rowspan; dr += 1) {
          const targetRow = r + dr;
          ensureRow(targetRow);
          for (let dc = 0; dc < colspan; dc += 1) {
            const targetCol = c + dc;
            occupied[targetRow][targetCol] = true;
            grid[targetRow][targetCol] = dr === 0 && dc === 0 ? text : '';
          }
        }

        if (rowspan > 1 || colspan > 1) {
          merges.push({
            s: { r, c },
            e: { r: r + rowspan - 1, c: c + colspan - 1 }
          });
        }

        c += colspan;
        if (c > maxCols) {
          maxCols = c;
        }
      }

      if (occupied[r].length > maxCols) {
        maxCols = occupied[r].length;
      }
    }

    for (let r = 0; r < grid.length; r += 1) {
      ensureRow(r);
      for (let c = 0; c < maxCols; c += 1) {
        if (grid[r][c] === undefined) {
          grid[r][c] = '';
        }
      }
    }

    return {
      rows: grid,
      merges,
      rowCount: grid.length,
      colCount: maxCols,
      hasContent
    };
  }

  /**
   * Filter out card grids, media galleries, and non-tabular blocks from heuristic results.
   */
  function passesDataTableHeuristics(container, matrix, rowElements, cellElements, hasHeaderSignal) {
    if (!matrix.hasContent || matrix.colCount < 2) {
      return false;
    }
    const minRows = hasHeaderSignal ? 2 : 3;
    if (matrix.rowCount < minRows) {
      return false;
    }

    // Reject navigation bars or footer links unless clearly multi-row data
    if (container.closest('nav') && matrix.rowCount < 4) {
      return false;
    }

    // Reject media/product card grids (large images/videos inside cells)
    let totalTextLen = 0;
    let nonEmptyCells = 0;
    for (let i = 0; i < cellElements.length; i += 1) {
      const cell = cellElements[i];
      if (!cell) continue;
      const media = cell.querySelector('img, video, canvas');
      if (media) {
        const mediaRect = media.getBoundingClientRect();
        if (mediaRect.width >= 64 && mediaRect.height >= 64) {
          return false;
        }
      }
      const text = cleanNodeText(cell);
      if (text) {
        totalTextLen += text.length;
        nonEmptyCells += 1;
      }
    }

    if (nonEmptyCells === 0) {
      return false;
    }

    // Reject tall multi-paragraph card layouts
    const avgLen = totalTextLen / nonEmptyCells;
    if (avgLen > 280) {
      return false;
    }

    if (rowElements && rowElements.length > 0) {
      const heights = rowElements.map((r) => r.getBoundingClientRect().height).sort((a, b) => a - b);
      const medianHeight = heights[Math.floor(heights.length / 2)];
      if (medianHeight > 200 && avgLen > 120) {
        return false;
      }
    }

    return true;
  }

  /**
   * Check whether the first row looks visually distinct as a table header.
   */
  function hasDistinctHeaderRow(rowElements) {
    if (!rowElements || rowElements.length < 2) return false;
    const first = rowElements[0];
    const second = rowElements[1];
    if (first.querySelector('[role="columnheader"]')) {
      return true;
    }

    const s1 = window.getComputedStyle(first);
    const s2 = window.getComputedStyle(second);
    const w1 = parseInt(s1.fontWeight, 10) || 400;
    const w2 = parseInt(s2.fontWeight, 10) || 400;
    if (w1 >= 600 && w1 > w2) return true;
    if (s1.backgroundColor !== s2.backgroundColor) return true;
    if (parseFloat(s1.borderBottomWidth) > 0 && s1.borderBottomWidth !== s2.borderBottomWidth) {
      return true;
    }
    return false;
  }

  /**
   * Extract an ARIA table/grid ([role="table"|"grid"|"treegrid"]).
   */
  function extractAriaTableMatrix(container) {
    const allRows = Array.from(container.querySelectorAll('[role="row"]')).filter(
      (row) => row.closest(`table, ${ARIA_TABLE_SELECTOR}`) === container && isTableVisible(row)
    );
    if (allRows.length < 2) return null;

    const rowCellLists = [];
    for (const row of allRows) {
      const cells = Array.from(row.querySelectorAll(ARIA_CELL_SELECTOR)).filter(
        (cell) => cell.closest('[role="row"]') === row && isTableVisible(cell)
      );
      if (cells.length > 0) {
        rowCellLists.push(cells);
      }
    }
    if (rowCellLists.length < 2) return null;

    const matrix = buildMatrixFromCellRows(rowCellLists, (cell) => {
      const colIndexAttr = parseInt(cell.getAttribute('aria-colindex') || '0', 10);
      const colspanAttr = parseInt(cell.getAttribute('aria-colspan') || '1', 10);
      const rowspanAttr = parseInt(cell.getAttribute('aria-rowspan') || '1', 10);
      return {
        colIndex: Number.isFinite(colIndexAttr) && colIndexAttr > 0 ? colIndexAttr - 1 : -1,
        colspan: Number.isFinite(colspanAttr) && colspanAttr > 0 ? colspanAttr : 1,
        rowspan: Number.isFinite(rowspanAttr) && rowspanAttr > 0 ? rowspanAttr : 1
      };
    });

    if (!matrix.hasContent || matrix.colCount < 2 || matrix.rowCount < 2) {
      return null;
    }
    return matrix;
  }

  /**
   * Extract a CSS display:table / inline-table structure.
   */
  function extractCssTableMatrix(container) {
    const headerRows = [];
    const bodyRows = [];
    const footerRows = [];

    for (const child of Array.from(container.children)) {
      if (!isTableVisible(child)) continue;
      const d = window.getComputedStyle(child).display;
      if (d === 'table-header-group') {
        for (const r of Array.from(child.children)) {
          if (isTableVisible(r) && window.getComputedStyle(r).display === 'table-row') {
            headerRows.push(r);
          }
        }
      } else if (d === 'table-row-group') {
        for (const r of Array.from(child.children)) {
          if (isTableVisible(r) && window.getComputedStyle(r).display === 'table-row') {
            bodyRows.push(r);
          }
        }
      } else if (d === 'table-footer-group') {
        for (const r of Array.from(child.children)) {
          if (isTableVisible(r) && window.getComputedStyle(r).display === 'table-row') {
            footerRows.push(r);
          }
        }
      } else if (d === 'table-row') {
        bodyRows.push(child);
      }
    }

    const allRows = [...headerRows, ...bodyRows, ...footerRows];
    if (allRows.length < 2) return null;

    const rowCellLists = allRows
      .map((row) =>
        Array.from(row.children).filter(
          (cell) => isTableVisible(cell) && window.getComputedStyle(cell).display === 'table-cell'
        )
      )
      .filter((cells) => cells.length > 0);

    if (rowCellLists.length < 2) return null;

    const matrix = buildMatrixFromCellRows(rowCellLists, null);
    if (!matrix.hasContent || matrix.colCount < 2 || matrix.rowCount < 2) {
      return null;
    }
    return matrix;
  }

  /**
   * Flatten direct children while unwrapping CSS display:contents wrappers.
   */
  function getVisibleFlowChildren(container, maxDepth = 2) {
    const result = [];
    for (const child of Array.from(container.children)) {
      if (SKIP_TAGS.has(child.tagName) || child.id === SHADOW_HOST_ID) continue;
      const style = window.getComputedStyle(child);
      if (style.display === 'none' || style.visibility === 'hidden') continue;
      if (style.position === 'absolute' || style.position === 'fixed') continue;
      if (style.display === 'contents' && maxDepth > 0) {
        result.push(...getVisibleFlowChildren(child, maxDepth - 1));
      } else if (isTableVisible(child)) {
        result.push(child);
      }
    }
    return result;
  }

  /**
   * Extract a flat CSS Grid table (display: grid with grid-template-columns).
   */
  function extractCssGridMatrix(container, containerStyle) {
    const rawTracks = (containerStyle.gridTemplateColumns || '').trim();
    if (!rawTracks || rawTracks === 'none') return null;

    const trackWidths = rawTracks
      .split(/\s+/)
      .map((token) => parseFloat(token))
      .filter((n) => Number.isFinite(n) && n > 0);

    if (trackWidths.length < 2) return null;

    const items = getVisibleFlowChildren(container);
    if (items.length < trackWidths.length * 2) return null;

    const itemEntries = items.map((el) => ({ el, rect: el.getBoundingClientRect() }));
    itemEntries.sort((a, b) => a.rect.top - b.rect.top || a.rect.left - b.rect.left);

    // Group grid items into visual rows by vertical top coordinate
    const visualRows = [];
    for (const item of itemEntries) {
      const lastRow = visualRows[visualRows.length - 1];
      if (!lastRow || Math.abs(item.rect.top - lastRow.top) > ROW_VERTICAL_TOLERANCE_PX) {
        visualRows.push({ top: item.rect.top, items: [item] });
      } else {
        lastRow.items.push(item);
      }
    }

    if (visualRows.length < 2) return null;

    // Compute pixel track boundaries anchored at the leftmost grid item
    let minLeft = Infinity;
    for (const item of itemEntries) {
      if (item.rect.left < minLeft) minLeft = item.rect.left;
    }
    const colGap = parseFloat(containerStyle.columnGap) || 0;
    const trackBounds = [];
    let cursorX = minLeft;
    for (let i = 0; i < trackWidths.length; i += 1) {
      trackBounds.push({ start: cursorX, end: cursorX + trackWidths[i] });
      cursorX += trackWidths[i] + colGap;
    }

    const rowCellLists = [];
    const rowHeaderWrappers = [];
    for (const vRow of visualRows) {
      vRow.items.sort((a, b) => a.rect.left - b.rect.left);
      rowCellLists.push(vRow.items.map((x) => x.el));
      rowHeaderWrappers.push(vRow.items[0].el);
    }

    const itemRectMap = new Map(itemEntries.map((x) => [x.el, x.rect]));
    const matrix = buildMatrixFromCellRows(rowCellLists, (cell) => {
      const rect = itemRectMap.get(cell);
      if (!rect) return null;
      let startCol = 0;
      let bestDist = Infinity;
      for (let i = 0; i < trackBounds.length; i += 1) {
        const dist = Math.abs(rect.left - trackBounds[i].start);
        if (dist < bestDist) {
          bestDist = dist;
          startCol = i;
        }
      }
      let span = 1;
      for (let i = startCol + 1; i < trackBounds.length; i += 1) {
        if (rect.right > trackBounds[i].start + trackWidths[i] * 0.35) {
          span += 1;
        } else {
          break;
        }
      }
      return { colIndex: startCol, colspan: span, rowspan: 1 };
    });

    const headerSignal = hasDistinctHeaderRow(rowHeaderWrappers);
    if (!passesDataTableHeuristics(container, matrix, null, items, headerSignal)) {
      return null;
    }
    return matrix;
  }

  /**
   * Unwrap single-child anchor/row wrapper around cells if a row has exactly one wrapper child.
   */
  function getRowCellElements(rowEl) {
    let current = rowEl;
    for (let depth = 0; depth < 2; depth += 1) {
      const children = getVisibleFlowChildren(current, 1);
      if (children.length === 1 && children[0].children.length >= 2) {
        current = children[0];
      } else {
        return children;
      }
    }
    return getVisibleFlowChildren(current, 1);
  }

  /**
   * Collect candidate row elements for flex/div tables, flattening 1-level header/body group wrappers.
   */
  function collectGeometricRowCandidates(container) {
    const direct = getVisibleFlowChildren(container, 1);
    if (direct.length >= 2 && direct.length <= 3) {
      const flattened = [];
      let usedFlattening = false;
      for (const child of direct) {
        const grandChildren = getVisibleFlowChildren(child, 1);
        const childCells = getRowCellElements(child);
        if (
          grandChildren.length >= 2 &&
          grandChildren.every((gc) => getRowCellElements(gc).length >= 2) &&
          childCells.length !== getRowCellElements(grandChildren[0]).length
        ) {
          flattened.push(...grandChildren);
          usedFlattening = true;
        } else {
          flattened.push(child);
        }
      }
      if (usedFlattening && flattened.length >= 2) {
        return flattened;
      }
    }
    return direct;
  }

  /**
   * Check whether cells in every row lie horizontally and columns align vertically across rows.
   */
  function verifyColumnAlignment(rowCellsWithRects) {
    const colCount = rowCellsWithRects[0].length;
    if (colCount < 2) return false;

    for (const row of rowCellsWithRects) {
      if (row.length !== colCount) return false;
      for (let c = 1; c < colCount; c += 1) {
        if (row[c].rect.left < row[c - 1].rect.right - 4) {
          return false;
        }
      }
    }

    for (let c = 0; c < colCount; c += 1) {
      let minLeft = Infinity;
      let maxLeft = -Infinity;
      let minRight = Infinity;
      let maxRight = -Infinity;
      let minCenter = Infinity;
      let maxCenter = -Infinity;

      for (const row of rowCellsWithRects) {
        const r = row[c].rect;
        const center = (r.left + r.right) / 2;
        if (r.left < minLeft) minLeft = r.left;
        if (r.left > maxLeft) maxLeft = r.left;
        if (r.right < minRight) minRight = r.right;
        if (r.right > maxRight) maxRight = r.right;
        if (center < minCenter) minCenter = center;
        if (center > maxCenter) maxCenter = center;
      }

      const aligned =
        maxLeft - minLeft <= COL_ALIGNMENT_TOLERANCE_PX ||
        maxRight - minRight <= COL_ALIGNMENT_TOLERANCE_PX ||
        maxCenter - minCenter <= COL_ALIGNMENT_TOLERANCE_PX;
      if (!aligned) {
        return false;
      }
    }

    return true;
  }

  /**
   * Extract a geometric column-aligned div/flex/list table UI.
   */
  function extractFlexTableMatrix(container) {
    const candidates = collectGeometricRowCandidates(container);
    if (candidates.length < 2) return null;

    const rowData = [];
    let prevBottom = -Infinity;
    for (const rowEl of candidates) {
      const rect = rowEl.getBoundingClientRect();
      if (rect.top < prevBottom - ROW_VERTICAL_TOLERANCE_PX) {
        return null;
      }
      prevBottom = rect.bottom;

      const cells = getRowCellElements(rowEl);
      if (cells.length < 2) return null;
      const cellEntries = cells
        .map((el) => ({ el, rect: el.getBoundingClientRect() }))
        .sort((a, b) => a.rect.left - b.rect.left);
      rowData.push({ rowEl, cells: cellEntries });
    }

    if (rowData.length < 2) return null;
    if (!verifyColumnAlignment(rowData.map((r) => r.cells))) {
      return null;
    }

    const rowCellLists = rowData.map((r) => r.cells.map((c) => c.el));
    const allCells = rowCellLists.flat();
    const rowElements = rowData.map((r) => r.rowEl);
    const matrix = buildMatrixFromCellRows(rowCellLists, null);
    const headerSignal = hasDistinctHeaderRow(rowElements);

    if (!passesDataTableHeuristics(container, matrix, rowElements, allCells, headerSignal)) {
      return null;
    }
    return matrix;
  }

  /**
   * Try extracting a non-<table> element as a table structure by priority.
   */
  function tryExtractNonNativeTable(el) {
    if (!el || el.tagName === 'TABLE' || el.tagName === 'SVG') return null;
    if (SKIP_TAGS.has(el.tagName) || el.id === SHADOW_HOST_ID) return null;
    if (!isTableVisible(el)) return null;

    if (el.matches && el.matches(ARIA_TABLE_SELECTOR)) {
      const matrix = extractAriaTableMatrix(el);
      if (matrix) return { kind: 'aria', matrix };
    }

    const style = window.getComputedStyle(el);
    if (style.display === 'table' || style.display === 'inline-table') {
      const matrix = extractCssTableMatrix(el);
      if (matrix) return { kind: 'css-table', matrix };
    }

    if (style.display === 'grid' || style.display === 'inline-grid') {
      const matrix = extractCssGridMatrix(el, style);
      if (matrix) return { kind: 'css-grid', matrix };
    }

    if (el.children.length >= 2) {
      const matrix = extractFlexTableMatrix(el);
      if (matrix) return { kind: 'flex-table', matrix };
    }

    return null;
  }

  /**
   * Re-extract fresh matrix for a matched entry of any detected table kind.
   */
  function extractMatchedEntryMatrix(entry) {
    if (!entry || !entry.element) return null;
    if (entry.kind === 'table' || entry.element.tagName === 'TABLE') {
      return extractTableMatrix(entry.element);
    }
    if (entry.kind === 'aria') {
      return extractAriaTableMatrix(entry.element);
    }
    if (entry.kind === 'css-table') {
      return extractCssTableMatrix(entry.element);
    }
    if (entry.kind === 'css-grid') {
      return extractCssGridMatrix(entry.element, window.getComputedStyle(entry.element));
    }
    if (entry.kind === 'flex-table') {
      return extractFlexTableMatrix(entry.element);
    }
    return null;
  }

  /**
   * Scan the active document: native <table> tags first, then ARIA/CSS Grid/flex tables.
   */
  function scanValidTables() {
    const valid = [];
    const detectedElements = new Set();

    const registerDetected = (element, kind, matrix) => {
      const fallbackIndex = valid.length + 1;
      const name = resolveTableName(element, fallbackIndex);
      detectedElements.add(element);
      valid.push({
        index: valid.length,
        kind,
        element,
        name,
        rowCount: matrix.rowCount,
        colCount: matrix.colCount,
        rows: matrix.rows,
        merges: matrix.merges
      });
    };

    // 1. Native <table> tags first (primary implementation)
    const allTables = Array.from(document.querySelectorAll('table'));
    for (const table of allTables) {
      if (!isTableVisible(table)) continue;
      const matrix = extractTableMatrix(table);
      if (!matrix.hasContent || matrix.rowCount === 0 || matrix.colCount === 0) {
        continue;
      }
      registerDetected(table, 'table', matrix);
    }

    // 2. ARIA table/grid/treegrid elements
    const ariaTables = Array.from(document.querySelectorAll(ARIA_TABLE_SELECTOR));
    for (const el of ariaTables) {
      if (overlapsDetectedElement(el, detectedElements)) continue;
      if (!isTableVisible(el)) continue;
      const matrix = extractAriaTableMatrix(el);
      if (matrix) {
        registerDetected(el, 'aria', matrix);
      }
    }

    // 3. CSS display:table, flat CSS Grid, and geometric flex/div tables
    if (document.body) {
      const isInsideDetected = (el) => {
        for (const existing of detectedElements) {
          if (existing === el || existing.contains(el)) return true;
        }
        return false;
      };

      const walk = (node) => {
        if (!node || node.nodeType !== Node.ELEMENT_NODE) return;
        if (SKIP_TAGS.has(node.tagName) || node.tagName === 'SVG' || node.id === SHADOW_HOST_ID) {
          return;
        }
        if (isInsideDetected(node)) {
          return;
        }
        if (node.children.length === 0 || !isTableVisible(node)) {
          return;
        }

        if (node.children.length >= 2 && !overlapsDetectedElement(node, detectedElements)) {
          const style = window.getComputedStyle(node);
          if (style.display === 'table' || style.display === 'inline-table') {
            const matrix = extractCssTableMatrix(node);
            if (matrix) {
              registerDetected(node, 'css-table', matrix);
              return;
            }
          } else if (style.display === 'grid' || style.display === 'inline-grid') {
            const matrix = extractCssGridMatrix(node, style);
            if (matrix) {
              registerDetected(node, 'css-grid', matrix);
              return;
            }
          } else {
            const matrix = extractFlexTableMatrix(node);
            if (matrix) {
              registerDetected(node, 'flex-table', matrix);
              return;
            }
          }
        }

        for (const child of Array.from(node.children)) {
          walk(child);
        }
      };

      walk(document.body);
    }

    lastScannedTables = valid;
    return valid;
  }

  /**
   * Build serializable payload for a scanned table entry.
   */
  function serializeTablePayload(entry) {
    return {
      name: entry.name,
      rows: entry.rows,
      merges: entry.merges,
      rowCount: entry.rowCount,
      colCount: entry.colCount
    };
  }

  /**
   * Scroll a detected table into view and apply a brief highlight ring.
   */
  function scrollToTableByIndex(index) {
    if (lastScannedTables.length === 0) {
      scanValidTables();
    }
    const entry = lastScannedTables[index];
    if (!entry || !entry.element || !entry.element.isConnected) {
      return;
    }

    if (activeScrollHighlightedEl) {
      activeScrollHighlightedEl.classList.remove(SCROLL_HIGHLIGHT_CLASS);
      activeScrollHighlightedEl = null;
    }
    if (scrollHighlightTimer) {
      clearTimeout(scrollHighlightTimer);
      scrollHighlightTimer = null;
    }

    entry.element.scrollIntoView({ behavior: 'smooth', block: 'center', inline: 'nearest' });
    entry.element.classList.add(SCROLL_HIGHLIGHT_CLASS);
    activeScrollHighlightedEl = entry.element;

    scrollHighlightTimer = setTimeout(() => {
      if (activeScrollHighlightedEl) {
        activeScrollHighlightedEl.classList.remove(SCROLL_HIGHLIGHT_CLASS);
        activeScrollHighlightedEl = null;
      }
      scrollHighlightTimer = null;
    }, 1600);
  }

  /**
   * Create or return the isolated Shadow DOM root for picker UI and toasts.
   */
  function ensureShadowOverlay() {
    if (selectionState.host && selectionState.host.isConnected && selectionState.shadow) {
      return selectionState.shadow;
    }

    const existing = document.getElementById(SHADOW_HOST_ID);
    if (existing) {
      existing.remove();
    }

    const host = document.createElement('div');
    host.id = SHADOW_HOST_ID;
    host.style.cssText =
      'all: initial; position: fixed; inset: 0; width: 100vw; height: 100vh; z-index: 2147483647; pointer-events: none;';

    const shadow = host.attachShadow({ mode: 'open' });
    const style = document.createElement('style');
    style.textContent = `
      :host {
        all: initial;
        font-family: -apple-system, BlinkMacSystemFont, "Inter", "Segoe UI", system-ui, sans-serif;
      }
      *, *::before, *::after {
        box-sizing: border-box;
      }
      .dtable-banner {
        position: fixed;
        top: 16px;
        left: 50%;
        transform: translateX(-50%);
        pointer-events: auto;
        display: inline-flex;
        align-items: center;
        gap: 12px;
        padding: 10px 12px 10px 16px;
        background: rgba(15, 23, 42, 0.94);
        backdrop-filter: blur(12px);
        -webkit-backdrop-filter: blur(12px);
        color: #f8fafc;
        border: 1px solid rgba(255, 255, 255, 0.14);
        border-radius: 9999px;
        box-shadow: 0 12px 32px rgba(2, 6, 23, 0.32), 0 0 0 1px rgba(79, 70, 229, 0.28);
        font-size: 13px;
        line-height: 1.3;
        font-weight: 500;
        letter-spacing: -0.01em;
        white-space: nowrap;
        max-width: calc(100vw - 32px);
        animation: dtableSlideDown 180ms cubic-bezier(0.16, 1, 0.3, 1);
      }
      .dtable-banner-dot {
        width: 8px;
        height: 8px;
        border-radius: 9999px;
        background: #6366f1;
        box-shadow: 0 0 0 3px rgba(99, 102, 241, 0.28);
        flex-shrink: 0;
      }
      .dtable-banner-text {
        overflow: hidden;
        text-overflow: ellipsis;
      }
      .dtable-cancel-btn {
        appearance: none;
        border: 1px solid rgba(255, 255, 255, 0.16);
        background: rgba(255, 255, 255, 0.08);
        color: #e2e8f0;
        border-radius: 9999px;
        padding: 4px 10px;
        font-size: 12px;
        font-weight: 600;
        font-family: inherit;
        cursor: pointer;
        transition: background 140ms ease, border-color 140ms ease, color 140ms ease;
        flex-shrink: 0;
      }
      .dtable-cancel-btn:hover {
        background: rgba(255, 255, 255, 0.16);
        border-color: rgba(255, 255, 255, 0.28);
        color: #ffffff;
      }
      .dtable-highlight {
        position: fixed;
        pointer-events: none;
        border: 2px solid #4f46e5;
        background: rgba(79, 70, 229, 0.09);
        border-radius: 6px;
        box-shadow: 0 0 0 4px rgba(79, 70, 229, 0.16);
        transition: top 70ms ease-out, left 70ms ease-out, width 70ms ease-out, height 70ms ease-out, opacity 120ms ease;
        opacity: 0;
      }
      .dtable-highlight.visible {
        opacity: 1;
      }
      .dtable-tooltip {
        position: fixed;
        pointer-events: none;
        display: inline-flex;
        align-items: center;
        gap: 6px;
        padding: 5px 10px;
        border-radius: 7px;
        background: rgba(15, 23, 42, 0.95);
        color: #f8fafc;
        border: 1px solid rgba(255, 255, 255, 0.14);
        box-shadow: 0 8px 20px rgba(2, 6, 23, 0.28);
        font-size: 12px;
        font-weight: 500;
        line-height: 1.25;
        max-width: min(360px, calc(100vw - 24px));
        white-space: nowrap;
        overflow: hidden;
        text-overflow: ellipsis;
        opacity: 0;
        transition: opacity 100ms ease;
      }
      .dtable-tooltip.visible {
        opacity: 1;
      }
      .dtable-tooltip-title {
        font-weight: 600;
        color: #ffffff;
        overflow: hidden;
        text-overflow: ellipsis;
      }
      .dtable-tooltip-meta {
        color: #a5b4fc;
        font-variant-numeric: tabular-nums;
        flex-shrink: 0;
      }
      .dtable-toast {
        position: fixed;
        bottom: 24px;
        right: 24px;
        pointer-events: none;
        display: flex;
        align-items: center;
        gap: 10px;
        padding: 12px 16px;
        background: rgba(15, 23, 42, 0.96);
        backdrop-filter: blur(12px);
        color: #f8fafc;
        border: 1px solid rgba(255, 255, 255, 0.14);
        border-radius: 12px;
        box-shadow: 0 14px 32px rgba(2, 6, 23, 0.35);
        font-size: 13px;
        line-height: 1.35;
        max-width: 340px;
        animation: dtableToastIn 180ms cubic-bezier(0.16, 1, 0.3, 1);
      }
      .dtable-toast-icon {
        width: 20px;
        height: 20px;
        border-radius: 9999px;
        background: rgba(16, 185, 129, 0.18);
        color: #34d399;
        display: flex;
        align-items: center;
        justify-content: center;
        flex-shrink: 0;
      }
      .dtable-toast-body {
        display: flex;
        flex-direction: column;
        min-width: 0;
      }
      .dtable-toast-title {
        font-weight: 600;
        color: #ffffff;
      }
      .dtable-toast-sub {
        font-size: 12px;
        color: #94a3b8;
        overflow: hidden;
        text-overflow: ellipsis;
        white-space: nowrap;
      }
      @keyframes dtableSlideDown {
        from {
          opacity: 0;
          transform: translate(-50%, -10px);
        }
        to {
          opacity: 1;
          transform: translate(-50%, 0);
        }
      }
      @keyframes dtableToastIn {
        from {
          opacity: 0;
          transform: translateY(8px) scale(0.97);
        }
        to {
          opacity: 1;
          transform: translateY(0) scale(1);
        }
      }
    `;
    shadow.appendChild(style);
    document.documentElement.appendChild(host);

    selectionState.host = host;
    selectionState.shadow = shadow;
    return shadow;
  }

  /**
   * Position the highlight overlay and floating tooltip for the currently hovered table.
   */
  function updateHighlightPosition() {
    const { active, hoveredTable, hoveredMeta, highlightBox, tooltip } = selectionState;
    if (!active || !highlightBox || !tooltip) return;

    if (!hoveredTable || !hoveredTable.isConnected) {
      highlightBox.classList.remove('visible');
      tooltip.classList.remove('visible');
      return;
    }

    const rect = hoveredTable.getBoundingClientRect();
    if (rect.width <= 0 || rect.height <= 0) {
      highlightBox.classList.remove('visible');
      tooltip.classList.remove('visible');
      return;
    }

    highlightBox.style.top = `${rect.top}px`;
    highlightBox.style.left = `${rect.left}px`;
    highlightBox.style.width = `${rect.width}px`;
    highlightBox.style.height = `${rect.height}px`;
    highlightBox.classList.add('visible');

    if (hoveredMeta) {
      const rowLabel = hoveredMeta.rowCount === 1 ? '1 row' : `${hoveredMeta.rowCount} rows`;
      const colLabel = hoveredMeta.colCount === 1 ? '1 col' : `${hoveredMeta.colCount} cols`;
      tooltip.querySelector('.dtable-tooltip-title').textContent = hoveredMeta.name;
      tooltip.querySelector('.dtable-tooltip-meta').textContent = `${rowLabel} × ${colLabel}`;

      const tooltipRect = tooltip.getBoundingClientRect();
      const tooltipHeight = tooltipRect.height || 28;
      const tooltipWidth = tooltipRect.width || 180;

      let top = rect.top - tooltipHeight - 8;
      if (top < 64) {
        top = rect.bottom + 8;
      }
      if (top + tooltipHeight > window.innerHeight - 12) {
        top = Math.max(68, rect.top + 8);
      }

      const maxLeft = Math.max(12, window.innerWidth - tooltipWidth - 12);
      const left = Math.min(Math.max(12, rect.left), maxLeft);

      tooltip.style.top = `${top}px`;
      tooltip.style.left = `${left}px`;
      tooltip.classList.add('visible');
    }
  }

  /**
   * Find the innermost valid table element under a given DOM event target.
   * Always checks native <table> elements first, then ARIA/CSS Grid/flex tables.
   */
  function findValidTableFromTarget(target) {
    if (!target || !(target instanceof Element)) return null;
    if (target.closest(`#${SHADOW_HOST_ID}`)) return null;

    // 1. Native <table> tags first (primary implementation)
    let candidate = target.closest('table');
    while (candidate) {
      if (isTableVisible(candidate)) {
        const cached = lastScannedTables.find((t) => t.element === candidate);
        if (cached) {
          return cached;
        }
        const matrix = extractTableMatrix(candidate);
        if (matrix.hasContent && matrix.rowCount > 0 && matrix.colCount > 0) {
          const fallbackIndex = lastScannedTables.length + 1;
          return {
            index: lastScannedTables.length,
            kind: 'table',
            element: candidate,
            name: resolveTableName(candidate, fallbackIndex),
            rowCount: matrix.rowCount,
            colCount: matrix.colCount,
            rows: matrix.rows,
            merges: matrix.merges
          };
        }
      }
      candidate = candidate.parentElement ? candidate.parentElement.closest('table') : null;
    }

    // 2. Cached non-<table> detections (ARIA, CSS table, CSS Grid, flex table), innermost first
    let bestCached = null;
    for (const entry of lastScannedTables) {
      if (entry.kind === 'table' || !entry.element || !entry.element.isConnected) continue;
      if (entry.element === target || entry.element.contains(target)) {
        if (!bestCached || bestCached.element.contains(entry.element)) {
          bestCached = entry;
        }
      }
    }
    if (bestCached) {
      return bestCached;
    }

    // 3. Dynamic ancestor walk for newly rendered ARIA/CSS Grid/flex tables
    let cursor = target;
    let depth = 0;
    while (cursor && cursor !== document.documentElement && depth < 12) {
      const extracted = tryExtractNonNativeTable(cursor);
      if (extracted) {
        const fallbackIndex = lastScannedTables.length + 1;
        return {
          index: lastScannedTables.length,
          kind: extracted.kind,
          element: cursor,
          name: resolveTableName(cursor, fallbackIndex),
          rowCount: extracted.matrix.rowCount,
          colCount: extracted.matrix.colCount,
          rows: extracted.matrix.rows,
          merges: extracted.matrix.merges
        };
      }
      cursor = cursor.parentElement;
      depth += 1;
    }

    return null;
  }

  function handleSelectionMouseMove(e) {
    if (!selectionState.active) return;

    const target = e.target;
    if (!(target instanceof Element) || target.closest(`#${SHADOW_HOST_ID}`)) {
      selectionState.hoveredTable = null;
      selectionState.hoveredMeta = null;
      updateHighlightPosition();
      return;
    }

    const closestNativeTable = target.closest('table');
    if (
      selectionState.hoveredTable &&
      (selectionState.hoveredTable === target || selectionState.hoveredTable.contains(target)) &&
      (!closestNativeTable || closestNativeTable === selectionState.hoveredTable)
    ) {
      updateHighlightPosition();
      return;
    }

    const matched = findValidTableFromTarget(target);
    if (!matched) {
      selectionState.hoveredTable = null;
      selectionState.hoveredMeta = null;
    } else {
      selectionState.hoveredTable = matched.element;
      selectionState.hoveredMeta = matched;
    }
    updateHighlightPosition();
  }

  function handleSelectionScrollOrResize() {
    if (!selectionState.active || selectionState.rafId !== null) return;
    selectionState.rafId = requestAnimationFrame(() => {
      selectionState.rafId = null;
      updateHighlightPosition();
    });
  }

  function handleSelectionKeyDown(e) {
    if (!selectionState.active) return;
    if (e.key === 'Escape') {
      e.preventDefault();
      e.stopPropagation();
      e.stopImmediatePropagation();
      stopSelectionMode(true);
    }
  }

  function handleSelectionClick(e) {
    if (!selectionState.active) return;

    // Allow clicks on our own banner cancel button
    const path = e.composedPath ? e.composedPath() : [];
    if (selectionState.banner && path.includes(selectionState.banner)) {
      return;
    }

    const matched = findValidTableFromTarget(e.target);
    if (!matched) {
      return;
    }

    e.preventDefault();
    e.stopPropagation();
    e.stopImmediatePropagation();

    const exportFormat = selectionState.format;
    const freshMatrix = extractMatchedEntryMatrix(matched) || {
      rows: matched.rows,
      merges: matched.merges,
      rowCount: matched.rowCount,
      colCount: matched.colCount
    };
    const payload = {
      name: matched.name,
      rows: freshMatrix.rows,
      merges: freshMatrix.merges,
      rowCount: freshMatrix.rowCount,
      colCount: freshMatrix.colCount
    };

    // Exit selection UI cleanly before showing toast
    stopSelectionMode(false);

    chrome.runtime.sendMessage(
      {
        action: 'EXPORT_TABLES',
        tables: [payload],
        mode: 'single',
        pageTitle: document.title,
        format: exportFormat
      },
      (response) => {
        if (chrome.runtime.lastError || !response || !response.ok) {
          const errMsg =
            (chrome.runtime.lastError && chrome.runtime.lastError.message) ||
            (response && response.error) ||
            'Download failed';
          showToast('Export Failed', errMsg);
          return;
        }
        const ext = exportFormat === 'csv' ? '.csv' : '.xlsx';
        showToast('Table Exported!', `${payload.name} (${ext})`);
      }
    );
  }

  /**
   * Show a brief toast notification inside the isolated Shadow DOM.
   */
  function showToast(titleText, subtitleText) {
    const shadow = ensureShadowOverlay();
    if (selectionState.toast) {
      selectionState.toast.remove();
      selectionState.toast = null;
    }
    if (selectionState.toastTimer) {
      clearTimeout(selectionState.toastTimer);
      selectionState.toastTimer = null;
    }

    const toast = document.createElement('div');
    toast.className = 'dtable-toast';

    const icon = document.createElement('div');
    icon.className = 'dtable-toast-icon';
    icon.innerHTML =
      '<svg width="13" height="13" viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><polyline points="3.5 8.5 6.5 11.5 12.5 4.5"></polyline></svg>';

    const body = document.createElement('div');
    body.className = 'dtable-toast-body';

    const title = document.createElement('div');
    title.className = 'dtable-toast-title';
    title.textContent = titleText;

    const sub = document.createElement('div');
    sub.className = 'dtable-toast-sub';
    sub.textContent = subtitleText || '';

    body.appendChild(title);
    if (subtitleText) {
      body.appendChild(sub);
    }
    toast.appendChild(icon);
    toast.appendChild(body);
    shadow.appendChild(toast);

    selectionState.toast = toast;
    selectionState.toastTimer = setTimeout(() => {
      if (selectionState.toast) {
        selectionState.toast.remove();
        selectionState.toast = null;
      }
      selectionState.toastTimer = null;
      if (!selectionState.active && selectionState.host) {
        selectionState.host.remove();
        selectionState.host = null;
        selectionState.shadow = null;
      }
    }, 2600);
  }

  /**
   * Activate interactive table selection mode on the active webpage.
   */
  function startSelectionMode(format) {
    if (selectionState.active) {
      stopSelectionMode(false);
    }

    selectionState.active = true;
    selectionState.format = format === 'csv' ? 'csv' : 'xlsx';
    selectionState.hoveredTable = null;
    selectionState.hoveredMeta = null;

    const shadow = ensureShadowOverlay();

    const highlightBox = document.createElement('div');
    highlightBox.className = 'dtable-highlight';

    const tooltip = document.createElement('div');
    tooltip.className = 'dtable-tooltip';
    const tooltipTitle = document.createElement('span');
    tooltipTitle.className = 'dtable-tooltip-title';
    const tooltipSep = document.createElement('span');
    tooltipSep.textContent = '•';
    tooltipSep.style.opacity = '0.6';
    const tooltipMeta = document.createElement('span');
    tooltipMeta.className = 'dtable-tooltip-meta';
    tooltip.append(tooltipTitle, tooltipSep, tooltipMeta);

    const banner = document.createElement('div');
    banner.className = 'dtable-banner';
    const dot = document.createElement('span');
    dot.className = 'dtable-banner-dot';
    const bannerText = document.createElement('span');
    bannerText.className = 'dtable-banner-text';
    bannerText.textContent =
      'DTable Selection Mode: Click any table to export, or press ESC to cancel.';
    const cancelBtn = document.createElement('button');
    cancelBtn.type = 'button';
    cancelBtn.className = 'dtable-cancel-btn';
    cancelBtn.textContent = 'Cancel';
    cancelBtn.addEventListener('click', (e) => {
      e.preventDefault();
      e.stopPropagation();
      stopSelectionMode(true);
    });
    banner.append(dot, bannerText, cancelBtn);

    shadow.append(highlightBox, tooltip, banner);
    selectionState.highlightBox = highlightBox;
    selectionState.tooltip = tooltip;
    selectionState.banner = banner;

    document.addEventListener('mousemove', handleSelectionMouseMove, true);
    document.addEventListener('click', handleSelectionClick, true);
    document.addEventListener('keydown', handleSelectionKeyDown, true);
    window.addEventListener('scroll', handleSelectionScrollOrResize, true);
    window.addEventListener('resize', handleSelectionScrollOrResize, true);
  }

  /**
   * Cleanly remove selection mode listeners, overlays, and styles.
   */
  function stopSelectionMode(removeHostIfNoToast) {
    selectionState.active = false;
    selectionState.hoveredTable = null;
    selectionState.hoveredMeta = null;

    if (selectionState.rafId !== null) {
      cancelAnimationFrame(selectionState.rafId);
      selectionState.rafId = null;
    }

    document.removeEventListener('mousemove', handleSelectionMouseMove, true);
    document.removeEventListener('click', handleSelectionClick, true);
    document.removeEventListener('keydown', handleSelectionKeyDown, true);
    window.removeEventListener('scroll', handleSelectionScrollOrResize, true);
    window.removeEventListener('resize', handleSelectionScrollOrResize, true);

    if (selectionState.banner) {
      selectionState.banner.remove();
      selectionState.banner = null;
    }
    if (selectionState.highlightBox) {
      selectionState.highlightBox.remove();
      selectionState.highlightBox = null;
    }
    if (selectionState.tooltip) {
      selectionState.tooltip.remove();
      selectionState.tooltip = null;
    }

    if (removeHostIfNoToast && !selectionState.toast && selectionState.host) {
      selectionState.host.remove();
      selectionState.host = null;
      selectionState.shadow = null;
    }
  }

  /**
   * Handle messages from the extension popup.
   */
  chrome.runtime.onMessage.addListener((message, _sender, sendResponse) => {
    if (!message || !message.action) return false;

    if (message.action === 'SCAN_TABLES') {
      const tables = scanValidTables();
      sendResponse({
        ok: true,
        pageTitle: document.title,
        tables: tables.map((t) => ({
          index: t.index,
          name: t.name,
          rowCount: t.rowCount,
          colCount: t.colCount
        }))
      });
      return false;
    }

    if (message.action === 'EXTRACT_ALL_TABLES') {
      const tables = scanValidTables();
      sendResponse({
        ok: true,
        pageTitle: document.title,
        tables: tables.map(serializeTablePayload)
      });
      return false;
    }

    if (message.action === 'EXTRACT_SINGLE_TABLE') {
      if (lastScannedTables.length === 0) {
        scanValidTables();
      }
      const entry = lastScannedTables[message.index];
      if (!entry) {
        sendResponse({ ok: false, error: 'Selected table is no longer available on the page.' });
        return false;
      }
      sendResponse({
        ok: true,
        pageTitle: document.title,
        table: serializeTablePayload(entry)
      });
      return false;
    }

    if (message.action === 'SCROLL_TO_TABLE') {
      scrollToTableByIndex(message.index);
      sendResponse({ ok: true });
      return false;
    }

    if (message.action === 'START_SELECTION') {
      startSelectionMode(message.format);
      sendResponse({ ok: true });
      return false;
    }

    return false;
  });
})();
