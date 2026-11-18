'use strict';

importScripts('lib/xlsx.full.min.js');

const STORAGE_FORMAT_KEY = 'exportFormat';
const DEFAULT_FORMAT = 'xlsx';
const EXCEL_SHEET_MAX_LEN = 31;
const FORBIDDEN_SHEET_CHARS_RE = /[\\/?*[\]:]/g;
const FORBIDDEN_FILENAME_CHARS_RE = /[<>:"/\\|?*\x00-\x1F]/g;
const MIME_XLSX = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
const MIME_CSV = 'text/csv;charset=utf-8';

/**
 * Ensure default format preference is initialized on install.
 */
chrome.runtime.onInstalled.addListener(() => {
  chrome.storage.local.get([STORAGE_FORMAT_KEY], (result) => {
    if (!result[STORAGE_FORMAT_KEY]) {
      chrome.storage.local.set({ [STORAGE_FORMAT_KEY]: DEFAULT_FORMAT });
    }
  });
});

/**
 * Listen for export requests from popup or content script.
 */
chrome.runtime.onMessage.addListener((message, _sender, sendResponse) => {
  if (!message || message.action !== 'EXPORT_TABLES') {
    return false;
  }

  handleExport(message)
    .then((result) => sendResponse({ ok: true, ...result }))
    .catch((err) => {
      console.error('[DTable] Export failed:', err);
      sendResponse({ ok: false, error: err && err.message ? err.message : 'Export failed' });
    });

  return true;
});

/**
 * Sanitize a string for safe use as a filename segment.
 */
function sanitizeFilenamePart(raw, fallback, maxLen = 60) {
  const cleaned = String(raw || '')
    .replace(FORBIDDEN_FILENAME_CHARS_RE, '_')
    .replace(/\s+/g, '_')
    .replace(/_+/g, '_')
    .replace(/^[._]+|[._]+$/g, '');

  const truncated = cleaned.slice(0, maxLen).replace(/[._]+$/g, '');
  return truncated || fallback;
}

/**
 * Deduplicate a filename stem while preserving a concise maximum length.
 */
function getUniqueFilenameStem(baseStem, usedLowerStems, maxLen = 50) {
  const normalizedBase = baseStem.slice(0, maxLen).replace(/[._]+$/g, '') || 'Table';
  if (!usedLowerStems.has(normalizedBase.toLowerCase())) {
    usedLowerStems.add(normalizedBase.toLowerCase());
    return normalizedBase;
  }

  let counter = 2;
  while (true) {
    const suffix = `_(${counter})`;
    const allowedBase = Math.max(1, maxLen - suffix.length);
    const candidate = `${normalizedBase.slice(0, allowedBase).replace(/[._]+$/g, '') || 'Table'}${suffix}`;
    const lower = candidate.toLowerCase();
    if (!usedLowerStems.has(lower)) {
      usedLowerStems.add(lower);
      return candidate;
    }
    counter += 1;
  }
}

/**
 * Sanitize and deduplicate an Excel worksheet name according to Excel's 31-char limit.
 */
function getUniqueSheetName(rawName, usedLowerNames, fallbackIndex) {
  let base = String(rawName || '')
    .replace(FORBIDDEN_SHEET_CHARS_RE, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .replace(/^'+|'+$/g, '');

  if (!base) {
    base = `Table ${fallbackIndex}`;
  }

  const truncated = base.slice(0, EXCEL_SHEET_MAX_LEN).trim().replace(/^'+|'+$/g, '') || `Table ${fallbackIndex}`;
  if (!usedLowerNames.has(truncated.toLowerCase())) {
    usedLowerNames.add(truncated.toLowerCase());
    return truncated;
  }

  let counter = 2;
  while (true) {
    const suffix = ` (${counter})`;
    const maxBaseLen = Math.max(1, EXCEL_SHEET_MAX_LEN - suffix.length);
    const trimmedBase = truncated.slice(0, maxBaseLen).trimEnd().replace(/^'+|'+$/g, '') || 'T';
    const candidate = `${trimmedBase}${suffix}`;
    const lower = candidate.toLowerCase();
    if (!usedLowerNames.has(lower)) {
      usedLowerNames.add(lower);
      return candidate;
    }
    counter += 1;
  }
}

/**
 * Compute readable column widths (in character units) from 2D cell rows.
 */
function computeColumnWidths(rows) {
  if (!rows || rows.length === 0) {
    return [];
  }
  let maxCols = 0;
  for (const row of rows) {
    if (row.length > maxCols) {
      maxCols = row.length;
    }
  }

  const widths = new Array(maxCols).fill(8);
  for (const row of rows) {
    for (let c = 0; c < maxCols; c += 1) {
      const cell = row[c] == null ? '' : String(row[c]);
      if (!cell) continue;
      const lines = cell.split(/\r?\n/);
      for (const line of lines) {
        const len = line.length + 2;
        if (len > widths[c]) {
          widths[c] = Math.min(48, len);
        }
      }
    }
  }
  return widths.map((wch) => ({ wch }));
}

/**
 * Serialize 2D array rows into an RFC 4180 compliant CSV string.
 */
function serializeCsv(rows) {
  return rows
    .map((row) =>
      row
        .map((cell) => {
          const text = cell == null ? '' : String(cell);
          if (/[",\r\n]/.test(text)) {
            return `"${text.replace(/"/g, '""')}"`;
          }
          return text;
        })
        .join(',')
    )
    .join('\r\n');
}

/**
 * Encode a UTF-8 string into base64 safely in a Service Worker context.
 */
function utf8ToBase64(text) {
  const bytes = new TextEncoder().encode(text);
  const chunkSize = 0x8000;
  const parts = [];
  for (let i = 0; i < bytes.length; i += chunkSize) {
    parts.push(String.fromCharCode.apply(null, bytes.subarray(i, i + chunkSize)));
  }
  return btoa(parts.join(''));
}

/**
 * Trigger a Chrome automatic download for a data URL.
 */
function downloadDataUrl(url, filename) {
  return new Promise((resolve, reject) => {
    chrome.downloads.download(
      {
        url,
        filename,
        saveAs: false,
        conflictAction: 'uniquify'
      },
      (downloadId) => {
        if (chrome.runtime.lastError) {
          reject(new Error(chrome.runtime.lastError.message));
          return;
        }
        resolve(downloadId);
      }
    );
  });
}

/**
 * Build and download either an Excel workbook (.xlsx) or individual CSV files.
 */
async function handleExport({ tables, mode, pageTitle, format }) {
  if (!Array.isArray(tables) || tables.length === 0) {
    throw new Error('No table data provided for export.');
  }

  const resolvedFormat = format === 'csv' ? 'csv' : 'xlsx';
  const safePageTitle = sanitizeFilenamePart(pageTitle, 'Webpage', 60);

  if (resolvedFormat === 'xlsx') {
    const wb = XLSX.utils.book_new();
    const usedSheetNames = new Set();

    tables.forEach((table, idx) => {
      const sheetName = getUniqueSheetName(table.name, usedSheetNames, idx + 1);
      const rows = Array.isArray(table.rows) ? table.rows : [];
      const ws = XLSX.utils.aoa_to_sheet(rows);

      if (Array.isArray(table.merges) && table.merges.length > 0) {
        ws['!merges'] = table.merges;
      }
      const cols = computeColumnWidths(rows);
      if (cols.length > 0) {
        ws['!cols'] = cols;
      }

      XLSX.utils.book_append_sheet(wb, ws, sheetName);
    });

    const base64 = XLSX.write(wb, {
      bookType: 'xlsx',
      type: 'base64',
      compression: true
    });

    const filename =
      mode === 'single'
        ? `${safePageTitle}_${sanitizeFilenamePart(tables[0].name, 'Table_1', 50)}.xlsx`
        : `${safePageTitle}_DTable.xlsx`;

    await downloadDataUrl(`data:${MIME_XLSX};base64,${base64}`, filename);
    return { count: 1, filenames: [filename] };
  }

  // CSV mode: export one file per table with deduplicated filenames.
  const usedCsvStems = new Set();
  const filenames = [];

  for (let i = 0; i < tables.length; i += 1) {
    const table = tables[i];
    const rawTableStem = sanitizeFilenamePart(table.name, `Table_${i + 1}`, 50);
    const uniqueStem = getUniqueFilenameStem(rawTableStem, usedCsvStems, 50);
    const filename = `${safePageTitle}_${uniqueStem}.csv`;

    const csvContent = serializeCsv(Array.isArray(table.rows) ? table.rows : []);
    const base64 = utf8ToBase64(csvContent);
    await downloadDataUrl(`data:${MIME_CSV};base64,${base64}`, filename);
    filenames.push(filename);
  }

  return { count: filenames.length, filenames };
}
