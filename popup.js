(() => {
  'use strict';

  const STORAGE_FORMAT_KEY = 'exportFormat';

  const statusPill = document.getElementById('statusPill');
  const statusText = document.getElementById('statusText');
  const formatXlsxBtn = document.getElementById('formatXlsxBtn');
  const formatCsvBtn = document.getElementById('formatCsvBtn');
  const exportAllBtn = document.getElementById('exportAllBtn');
  const exportAllLabel = document.getElementById('exportAllLabel');
  const selectTableBtn = document.getElementById('selectTableBtn');
  const tablesHint = document.getElementById('tablesHint');
  const tablesList = document.getElementById('tablesList');
  const emptyState = document.getElementById('emptyState');
  const emptyTitle = document.getElementById('emptyTitle');
  const emptyDesc = document.getElementById('emptyDesc');
  const feedbackBanner = document.getElementById('feedbackBanner');

  let currentFormat = 'xlsx';
  let activeTabId = null;
  let detectedTables = [];
  let feedbackTimer = null;

  const DOWNLOAD_SVG =
    '<svg width="13" height="13" viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.75" stroke-linecap="round" stroke-linejoin="round"><path d="M8 2.25V10.5"/><path d="M4.75 7.5L8 10.75L11.25 7.5"/><path d="M2.75 11.75V12.5C2.75 13.05 3.2 13.5 3.75 13.5H12.25C12.8 13.5 13.25 13.05 13.25 12.5V11.75"/></svg>';
  const CHECK_SVG =
    '<svg width="13" height="13" viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polyline points="3.5 8.5 6.5 11.5 12.5 4.5"/></svg>';

  /**
   * Show a temporary inline status message at the bottom of the popup.
   */
  function showFeedback(text, type) {
    if (feedbackTimer) {
      clearTimeout(feedbackTimer);
      feedbackTimer = null;
    }
    feedbackBanner.textContent = text;
    feedbackBanner.className = `feedback-banner is-${type}`;
    feedbackBanner.hidden = false;

    feedbackTimer = setTimeout(() => {
      feedbackBanner.hidden = true;
      feedbackTimer = null;
    }, 3000);
  }

  /**
   * Update the format segmented control UI and persist selection.
   */
  function setFormat(format, persist) {
    currentFormat = format === 'csv' ? 'csv' : 'xlsx';
    const isXlsx = currentFormat === 'xlsx';

    formatXlsxBtn.classList.toggle('is-active', isXlsx);
    formatXlsxBtn.setAttribute('aria-checked', String(isXlsx));
    formatCsvBtn.classList.toggle('is-active', !isXlsx);
    formatCsvBtn.setAttribute('aria-checked', String(!isXlsx));

    if (persist) {
      chrome.storage.local.set({ [STORAGE_FORMAT_KEY]: currentFormat });
    }
  }

  /**
   * Load persisted export format preference from chrome.storage.local.
   */
  function loadFormatPreference() {
    return new Promise((resolve) => {
      chrome.storage.local.get([STORAGE_FORMAT_KEY], (result) => {
        const saved = result[STORAGE_FORMAT_KEY] === 'csv' ? 'csv' : 'xlsx';
        setFormat(saved, false);
        resolve(saved);
      });
    });
  }

  /**
   * Check whether Chrome allows scripting on the given tab URL.
   */
  function isRestrictedUrl(url) {
    if (!url) return true;
    return (
      url.startsWith('chrome://') ||
      url.startsWith('chrome-extension://') ||
      url.startsWith('edge://') ||
      url.startsWith('about:') ||
      url.startsWith('https://chrome.google.com/webstore') ||
      url.startsWith('https://chromewebstore.google.com')
    );
  }

  /**
   * Ensure content.css and content.js are injected into the active tab.
   */
  async function ensureContentScripts(tabId) {
    await chrome.scripting.insertCSS({
      target: { tabId },
      files: ['content.css']
    });
    await chrome.scripting.executeScript({
      target: { tabId },
      files: ['content.js']
    });
  }

  /**
   * Send a message to the active tab's content script.
   */
  function sendTabMessage(tabId, payload) {
    return new Promise((resolve, reject) => {
      chrome.tabs.sendMessage(tabId, payload, (response) => {
        if (chrome.runtime.lastError) {
          reject(new Error(chrome.runtime.lastError.message));
          return;
        }
        resolve(response);
      });
    });
  }

  /**
   * Send an export request to the background service worker.
   */
  function sendExportRequest(payload) {
    return new Promise((resolve, reject) => {
      chrome.runtime.sendMessage({ action: 'EXPORT_TABLES', ...payload }, (response) => {
        if (chrome.runtime.lastError) {
          reject(new Error(chrome.runtime.lastError.message));
          return;
        }
        if (!response || !response.ok) {
          reject(new Error((response && response.error) || 'Export failed'));
          return;
        }
        resolve(response);
      });
    });
  }

  /**
   * Render the detected tables mini-list and empty state.
   */
  function renderTables(tables) {
    detectedTables = tables;
    tablesList.innerHTML = '';

    const count = tables.length;
    statusText.textContent = count === 1 ? '1 Table Found' : `${count} Tables Found`;
    statusPill.classList.toggle('has-tables', count > 0);
    exportAllBtn.disabled = count === 0;
    tablesHint.hidden = count === 0;

    if (count === 0) {
      emptyTitle.textContent = 'No tables found';
      emptyDesc.textContent = 'This page does not contain any visible data tables.';
      emptyState.hidden = false;
      return;
    }

    emptyState.hidden = true;

    for (const table of tables) {
      const item = document.createElement('div');
      item.className = 'table-item';
      item.setAttribute('role', 'listitem');

      const info = document.createElement('div');
      info.className = 'table-item-info';

      const nameEl = document.createElement('div');
      nameEl.className = 'table-item-name';
      nameEl.textContent = table.name;
      nameEl.title = table.name;

      const rowLabel = table.rowCount === 1 ? '1 row' : `${table.rowCount} rows`;
      const colLabel = table.colCount === 1 ? '1 col' : `${table.colCount} cols`;
      const metaEl = document.createElement('div');
      metaEl.className = 'table-item-meta';
      metaEl.textContent = `${rowLabel} × ${colLabel}`;

      info.append(nameEl, metaEl);

      const exportBtn = document.createElement('button');
      exportBtn.type = 'button';
      exportBtn.className = 'table-export-btn';
      exportBtn.title = `Export "${table.name}"`;
      exportBtn.setAttribute('aria-label', `Export ${table.name}`);
      exportBtn.innerHTML = DOWNLOAD_SVG;

      item.addEventListener('mouseenter', () => {
        if (activeTabId !== null) {
          sendTabMessage(activeTabId, { action: 'SCROLL_TO_TABLE', index: table.index }).catch(() => {});
        }
      });

      exportBtn.addEventListener('click', (e) => {
        e.stopPropagation();
        handleSingleExport(table.index, exportBtn);
      });

      item.append(info, exportBtn);
      tablesList.appendChild(item);
    }
  }

  /**
   * Show restricted or unreachable page state.
   */
  function renderUnavailableState(reason) {
    statusText.textContent = 'Unavailable';
    statusPill.classList.remove('has-tables');
    exportAllBtn.disabled = true;
    selectTableBtn.disabled = true;
    tablesHint.hidden = true;
    tablesList.innerHTML = '';
    emptyTitle.textContent = 'Page cannot be scanned';
    emptyDesc.textContent = reason;
    emptyState.hidden = false;
  }

  /**
   * Export a single table directly from the popup mini-list.
   */
  async function handleSingleExport(tableIndex, buttonEl) {
    if (activeTabId === null) return;

    buttonEl.disabled = true;
    buttonEl.innerHTML = '<span class="mini-spinner" aria-hidden="true"></span>';

    try {
      const extractRes = await sendTabMessage(activeTabId, {
        action: 'EXTRACT_SINGLE_TABLE',
        index: tableIndex
      });
      if (!extractRes || !extractRes.ok || !extractRes.table) {
        throw new Error((extractRes && extractRes.error) || 'Failed to read table');
      }

      const exportRes = await sendExportRequest({
        tables: [extractRes.table],
        mode: 'single',
        pageTitle: extractRes.pageTitle,
        format: currentFormat
      });

      buttonEl.classList.add('is-done');
      buttonEl.innerHTML = CHECK_SVG;
      showFeedback(`Exported ${exportRes.filenames[0]}`, 'success');

      setTimeout(() => {
        buttonEl.classList.remove('is-done');
        buttonEl.innerHTML = DOWNLOAD_SVG;
        buttonEl.disabled = false;
      }, 1500);
    } catch (err) {
      buttonEl.innerHTML = DOWNLOAD_SVG;
      buttonEl.disabled = false;
      showFeedback(err.message || 'Export failed', 'error');
    }
  }

  /**
   * Export all detected tables on the active page.
   */
  async function handleExportAll() {
    if (activeTabId === null || detectedTables.length === 0) return;

    exportAllBtn.disabled = true;
    exportAllBtn.classList.add('is-loading');
    exportAllLabel.textContent = 'Exporting...';

    try {
      const extractRes = await sendTabMessage(activeTabId, { action: 'EXTRACT_ALL_TABLES' });
      if (!extractRes || !extractRes.ok || !Array.isArray(extractRes.tables) || extractRes.tables.length === 0) {
        throw new Error('No tables available to export.');
      }

      const exportRes = await sendExportRequest({
        tables: extractRes.tables,
        mode: 'all',
        pageTitle: extractRes.pageTitle,
        format: currentFormat
      });

      const msg =
        currentFormat === 'xlsx'
          ? `Exported ${exportRes.filenames[0]}`
          : `Exported ${exportRes.count} CSV ${exportRes.count === 1 ? 'file' : 'files'}`;
      showFeedback(msg, 'success');
    } catch (err) {
      showFeedback(err.message || 'Export failed', 'error');
    } finally {
      exportAllBtn.classList.remove('is-loading');
      exportAllLabel.textContent = 'Export All Tables';
      exportAllBtn.disabled = detectedTables.length === 0;
    }
  }

  /**
   * Activate interactive table picker on the active page and close popup.
   */
  async function handleStartSelection() {
    if (activeTabId === null) return;

    try {
      await ensureContentScripts(activeTabId);
      await sendTabMessage(activeTabId, {
        action: 'START_SELECTION',
        format: currentFormat
      });
      window.close();
    } catch (err) {
      showFeedback(err.message || 'Could not start selection mode', 'error');
    }
  }

  /**
   * Initialize popup state on open.
   */
  async function init() {
    await loadFormatPreference();

    formatXlsxBtn.addEventListener('click', () => setFormat('xlsx', true));
    formatCsvBtn.addEventListener('click', () => setFormat('csv', true));
    exportAllBtn.addEventListener('click', handleExportAll);
    selectTableBtn.addEventListener('click', handleStartSelection);

    const tabs = await chrome.tabs.query({ active: true, currentWindow: true });
    const activeTab = tabs && tabs[0];

    if (!activeTab || activeTab.id === undefined) {
      renderUnavailableState('No active browser tab detected.');
      return;
    }

    activeTabId = activeTab.id;

    if (isRestrictedUrl(activeTab.url)) {
      renderUnavailableState('Chrome restricts extensions on browser system pages.');
      return;
    }

    try {
      await ensureContentScripts(activeTabId);
      const response = await sendTabMessage(activeTabId, { action: 'SCAN_TABLES' });
      if (!response || !response.ok) {
        throw new Error('Could not scan page tables.');
      }
      renderTables(response.tables || []);
    } catch (_err) {
      renderUnavailableState('Refresh this page or open a standard webpage to scan tables.');
    }
  }

  document.addEventListener('DOMContentLoaded', init);
})();
