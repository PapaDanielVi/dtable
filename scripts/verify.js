'use strict';

const fs = require('fs');
const vm = require('vm');
const assert = require('assert');
const path = require('path');

const rootDir = path.resolve(__dirname, '..');
const downloads = [];
let messageListener = null;

function verifyManifestAndAssets() {
  const manifestPath = path.join(rootDir, 'manifest.json');
  const manifest = JSON.parse(fs.readFileSync(manifestPath, 'utf8'));
  assert.strictEqual(manifest.manifest_version, 3);
  assert.strictEqual(manifest.name, 'DTable');

  const requiredFiles = [
    'background.js',
    'content.js',
    'content.css',
    'popup.html',
    'popup.css',
    'popup.js',
    'lib/xlsx.full.min.js',
    'icons/icon16.png',
    'icons/icon32.png',
    'icons/icon48.png',
    'icons/icon128.png'
  ];

  for (const rel of requiredFiles) {
    const full = path.join(rootDir, rel);
    assert(fs.existsSync(full), `Missing required file: ${rel}`);
  }
}

function createBackgroundContext() {
  const sandbox = {
    console,
    Uint8Array,
    ArrayBuffer,
    DataView,
    TextEncoder,
    TextDecoder,
    Date,
    Math,
    RegExp,
    Error,
    TypeError,
    RangeError,
    JSON,
    Object,
    Array,
    String,
    Number,
    Boolean,
    Set,
    Map,
    Promise,
    btoa: (s) => Buffer.from(s, 'binary').toString('base64'),
    parseInt,
    parseFloat,
    isNaN,
    isFinite,
    importScripts: (relPath) => {
      const code = fs.readFileSync(path.join(rootDir, relPath), 'utf8');
      vm.runInContext(code, context);
    },
    chrome: {
      runtime: {
        lastError: null,
        onInstalled: { addListener: () => {} },
        onMessage: {
          addListener: (fn) => {
            messageListener = fn;
          }
        }
      },
      storage: {
        local: { get: (_k, cb) => cb({}), set: () => {} }
      },
      downloads: {
        download: (opts, cb) => {
          downloads.push(opts);
          cb(downloads.length);
        }
      }
    }
  };
  sandbox.self = sandbox;
  sandbox.globalThis = sandbox;
  const context = vm.createContext(sandbox);
  vm.runInContext(fs.readFileSync(path.join(rootDir, 'background.js'), 'utf8'), context);
  return sandbox;
}

function dispatchExport(payload) {
  return new Promise((resolve, reject) => {
    messageListener(payload, {}, (res) => {
      if (res && res.ok) {
        resolve(res);
      } else {
        reject(new Error((res && res.error) || 'Export failed'));
      }
    });
  });
}

async function main() {
  verifyManifestAndAssets();
  const sandbox = createBackgroundContext();

  // Sheet name sanitization, 31-character cap, and case-insensitive deduplication.
  const used = new Set();
  assert.strictEqual(sandbox.getUniqueSheetName('Sales/Data: [2026]*?\\', used, 1), 'Sales Data 2026');
  const longName = 'Very Long Worksheet Name Over Thirty One Chars';
  const firstLong = sandbox.getUniqueSheetName(longName, used, 2);
  assert.strictEqual(firstLong, 'Very Long Worksheet Name Over T');
  assert(firstLong.length <= 31);
  const dupLong = sandbox.getUniqueSheetName(longName, used, 3);
  assert.strictEqual(dupLong, 'Very Long Worksheet Name Ov (2)');
  assert(dupLong.length <= 31);
  assert.strictEqual(sandbox.getUniqueSheetName('sales data 2026', used, 4), 'sales data 2026 (2)');

  // RFC 4180 CSV escaping.
  const csv = sandbox.serializeCsv([
    ['Name', 'Note, "Quoted"', 'Multi\nLine'],
    ['Alice', 'OK', 'Line1\nLine2']
  ]);
  assert.strictEqual(csv, 'Name,"Note, ""Quoted""","Multi\nLine"\r\nAlice,OK,"Line1\nLine2"');

  // Full-page XLSX export and SheetJS roundtrip parse.
  downloads.length = 0;
  await dispatchExport({
    action: 'EXPORT_TABLES',
    mode: 'all',
    pageTitle: 'Q4 Report: Financial / Summary',
    format: 'xlsx',
    tables: [
      {
        name: 'Revenue/2026',
        rows: [['Quarter', 'Amount'], ['Q1', '$100']],
        merges: [{ s: { r: 0, c: 0 }, e: { r: 0, c: 1 } }]
      },
      {
        name: 'Revenue/2026',
        rows: [['Region', 'Units'], ['US', '50']],
        merges: []
      }
    ]
  });
  assert.strictEqual(downloads.length, 1);
  assert.strictEqual(downloads[0].filename, 'Q4_Report_Financial_Summary_DTable.xlsx');
  const parsedWb = sandbox.XLSX.read(downloads[0].url.split(',')[1], { type: 'base64' });
  assert.deepStrictEqual(Array.from(parsedWb.SheetNames), ['Revenue 2026', 'Revenue 2026 (2)']);
  assert.strictEqual(parsedWb.Sheets['Revenue 2026'].A2.v, 'Q1');
  assert.strictEqual(parsedWb.Sheets['Revenue 2026']['!merges'].length, 1);

  // Single-table XLSX filename check.
  downloads.length = 0;
  await dispatchExport({
    action: 'EXPORT_TABLES',
    mode: 'single',
    pageTitle: 'Dashboard',
    format: 'xlsx',
    tables: [{ name: 'Quarterly Revenue', rows: [['A', 'B'], ['1', '2']], merges: [] }]
  });
  assert.strictEqual(downloads[0].filename, 'Dashboard_Quarterly_Revenue.xlsx');

  // Multi-table CSV export with UTF-8 preservation and filename deduplication.
  downloads.length = 0;
  await dispatchExport({
    action: 'EXPORT_TABLES',
    mode: 'all',
    pageTitle: 'Unicode Test',
    format: 'csv',
    tables: [
      { name: 'Summary', rows: [['Col1', 'Col2'], ['Café', '1,234']] },
      { name: 'Summary', rows: [['A'], ['B']] }
    ]
  });
  assert.strictEqual(downloads.length, 2);
  assert.strictEqual(downloads[0].filename, 'Unicode_Test_Summary.csv');
  assert.strictEqual(downloads[1].filename, 'Unicode_Test_Summary_(2).csv');
  assert.strictEqual(
    Buffer.from(downloads[0].url.split(',')[1], 'base64').toString('utf8'),
    'Col1,Col2\r\nCafé,"1,234"'
  );

  // Filename safety: surrogate-pair boundary, control chars, reserved names, dot-dot, byte cap.
  const surrogateBoundaryTitle = 'A'.repeat(59) + '📊Tail';
  const safeSurrogate = sandbox.sanitizeFilenamePart(surrogateBoundaryTitle, 'Webpage', 60);
  assert.strictEqual(safeSurrogate.isWellFormed(), true);

  assert.strictEqual(sandbox.sanitizeFilenamePart('CON', 'Webpage'), 'Webpage');
  assert.strictEqual(sandbox.sanitizeFilenamePart('nul', 'Table_1'), 'Table_1');
  assert.strictEqual(sandbox.sanitizeFilenamePart('...\u200B\u202E\x7F...', 'Webpage'), 'Webpage');
  assert.strictEqual(sandbox.sanitizeFilenamePart('foo/../bar', 'Webpage'), 'foo_bar');

  // Long multi-byte CJK title + table name must stay well within OS 255-byte limit.
  const longCjk = '日本語テスト表データ'.repeat(10);
  const cjkFile = sandbox.buildSafeFilename(
    sandbox.sanitizeFilenamePart(longCjk, 'Webpage'),
    sandbox.sanitizeFilenamePart(longCjk, 'Table_1'),
    'xlsx'
  );
  assert(Buffer.byteLength(cjkFile, 'utf8') <= 140, `Filename too long: ${cjkFile}`);
  assert.strictEqual(cjkFile.isWellFormed(), true);

  // Automatic ASCII fallback retry if Chrome rejects a Unicode filename as Invalid filename.
  downloads.length = 0;
  let rejectFirst = true;
  const origDownload = sandbox.chrome.downloads.download;
  sandbox.chrome.downloads.download = (opts, cb) => {
    if (rejectFirst) {
      rejectFirst = false;
      sandbox.chrome.runtime.lastError = { message: 'Invalid filename' };
      cb(undefined);
      sandbox.chrome.runtime.lastError = null;
      return;
    }
    origDownload(opts, cb);
  };
  const retryRes = await dispatchExport({
    action: 'EXPORT_TABLES',
    mode: 'single',
    pageTitle: 'Café 📊 Report',
    format: 'csv',
    tables: [{ name: '日本語', rows: [['A'], ['1']] }]
  });
  sandbox.chrome.downloads.download = origDownload;
  assert.strictEqual(retryRes.filenames[0], 'Caf_Report.csv');

  verifyContentDetection();

  console.log('All verification checks passed.');
}

function verifyContentDetection() {
  const NodeConsts = { ELEMENT_NODE: 1, TEXT_NODE: 3 };

  class MockText {
    constructor(val) {
      this.nodeType = NodeConsts.TEXT_NODE;
      this.nodeValue = val;
    }
  }

  class MockElement {
    constructor(tagName, attrs = {}, rect = { left: 0, top: 0, width: 400, height: 100 }, style = {}) {
      this.nodeType = NodeConsts.ELEMENT_NODE;
      this.tagName = tagName.toUpperCase();
      this.attrs = { ...attrs };
      this.id = attrs.id || '';
      this.children = [];
      this.childNodes = [];
      this.parentElement = null;
      this.isConnected = true;
      this.hidden = false;
      this.rect = { ...rect, right: rect.left + rect.width, bottom: rect.top + rect.height };
      this.computedStyle = {
        display: 'block',
        visibility: 'visible',
        position: 'static',
        fontWeight: '400',
        backgroundColor: 'rgba(0, 0, 0, 0)',
        borderBottomWidth: '0px',
        gridTemplateColumns: 'none',
        columnGap: '0px',
        ...style
      };
      this.style = {};
      this.classList = { add: () => {}, remove: () => {} };
    }

    append(...nodes) {
      for (const n of nodes) {
        if (typeof n === 'string') {
          this.childNodes.push(new MockText(n));
        } else {
          n.parentElement = this;
          this.children.push(n);
          this.childNodes.push(n);
        }
      }
      return this;
    }

    getAttribute(name) {
      return Object.prototype.hasOwnProperty.call(this.attrs, name) ? this.attrs[name] : null;
    }

    getBoundingClientRect() {
      return this.rect;
    }

    getClientRects() {
      return this.rect.width > 0 && this.rect.height > 0 ? [this.rect] : [];
    }

    contains(other) {
      let cur = other;
      while (cur) {
        if (cur === this) return true;
        cur = cur.parentElement;
      }
      return false;
    }

    matches(selector) {
      const parts = selector.split(',').map((s) => s.trim()).filter(Boolean);
      return parts.some((part) => {
        const roleMatch = part.match(/^\[role="([^"]+)"\]$/);
        if (roleMatch) {
          return this.getAttribute('role') === roleMatch[1];
        }
        return this.tagName === part.toUpperCase();
      });
    }

    closest(selector) {
      let cur = this;
      while (cur) {
        if (cur.matches(selector)) return cur;
        cur = cur.parentElement;
      }
      return null;
    }

    querySelector(selector) {
      return this.querySelectorAll(selector)[0] || null;
    }

    querySelectorAll(selector) {
      const results = [];
      const visit = (el) => {
        for (const child of el.children) {
          if (child.matches(selector)) {
            results.push(child);
          }
          visit(child);
        }
      };
      visit(this);
      return results;
    }
  }

  class MockTable extends MockElement {
    constructor(attrs, rect) {
      super('table', attrs, rect, { display: 'table' });
      this.tHead = null;
      this.tBodies = [];
      this.tFoot = null;
      this.caption = null;
    }
  }

  const body = new MockElement('body', {}, { left: 0, top: 0, width: 1000, height: 1000 });

  // 1. Native <table>
  const nativeTable = new MockTable({ 'aria-label': 'Native Sales' }, { left: 0, top: 0, width: 400, height: 80 });
  const tbody = new MockElement('tbody');
  tbody.rows = [];
  const tr1 = new MockElement('tr');
  const th1 = new MockElement('th');
  th1.append('Item');
  const th2 = new MockElement('th');
  th2.append('Qty');
  tr1.cells = [th1, th2];
  tr1.append(th1, th2);
  const tr2 = new MockElement('tr');
  const td1 = new MockElement('td');
  td1.append('Widget');
  const td2 = new MockElement('td');
  td2.append('42');
  tr2.cells = [td1, td2];
  tr2.append(td1, td2);
  tbody.rows.push(tr1, tr2);
  tbody.append(tr1, tr2);
  nativeTable.tBodies.push(tbody);
  nativeTable.append(tbody);

  // Wrapper around native table should NOT produce a duplicate flex/div detection.
  const nativeWrapper = new MockElement('div', {}, { left: 0, top: 0, width: 400, height: 80 });
  nativeWrapper.append(nativeTable);

  // 2. ARIA Grid
  const ariaGrid = new MockElement('div', { role: 'grid', 'aria-label': 'ARIA Metrics' }, { left: 0, top: 100, width: 400, height: 80 });
  const aRow1 = new MockElement('div', { role: 'row' }, { left: 0, top: 100, width: 400, height: 30 });
  aRow1.append(
    new MockElement('div', { role: 'columnheader' }, { left: 0, top: 100, width: 200, height: 30 }).append('Metric'),
    new MockElement('div', { role: 'columnheader' }, { left: 200, top: 100, width: 200, height: 30 }).append('Value')
  );
  const aRow2 = new MockElement('div', { role: 'row' }, { left: 0, top: 130, width: 400, height: 30 });
  aRow2.append(
    new MockElement('div', { role: 'gridcell' }, { left: 0, top: 130, width: 200, height: 30 }).append('Latency'),
    new MockElement('div', { role: 'gridcell' }, { left: 200, top: 130, width: 200, height: 30 }).append('12ms')
  );
  ariaGrid.append(aRow1, aRow2);

  // 3. Flat CSS Grid (no row wrappers)
  const cssGrid = new MockElement(
    'div',
    { id: 'pricing-grid' },
    { left: 0, top: 200, width: 300, height: 90 },
    { display: 'grid', gridTemplateColumns: '150px 150px' }
  );
  cssGrid.append(
    new MockElement('div', {}, { left: 0, top: 200, width: 150, height: 30 }, { fontWeight: '700' }).append('Plan'),
    new MockElement('div', {}, { left: 150, top: 200, width: 150, height: 30 }, { fontWeight: '700' }).append('Price'),
    new MockElement('div', {}, { left: 0, top: 230, width: 150, height: 30 }).append('Pro'),
    new MockElement('div', {}, { left: 150, top: 230, width: 150, height: 30 }).append('$20'),
    new MockElement('div', {}, { left: 0, top: 260, width: 150, height: 30 }).append('Team'),
    new MockElement('div', {}, { left: 150, top: 260, width: 150, height: 30 }).append('$50')
  );

  // 4. Geometric <div> / <p> flex-row table
  const flexTable = new MockElement('div', { id: 'team_directory' }, { left: 0, top: 320, width: 400, height: 90 });
  const flexRows = [
    ['Name', 'Role', 'Top', 320],
    ['Alice', 'Engineer', 'Top', 350],
    ['Bob', 'Designer', 'Top', 380]
  ];
  for (const [colA, colB, _t, y] of flexRows) {
    const row = new MockElement('div', {}, { left: 0, top: y, width: 400, height: 28 }, { display: 'flex' });
    const p1 = new MockElement('p', {}, { left: 10, top: y, width: 160, height: 24 });
    p1.append(colA);
    const p2 = new MockElement('p', {}, { left: 200, top: y, width: 160, height: 24 });
    p2.append(colB);
    row.append(p1, p2);
    flexTable.append(row);
  }

  // 5. Misaligned non-table <div> list (must be rejected)
  const misalignedCards = new MockElement('div', { id: 'not-a-table' }, { left: 0, top: 440, width: 400, height: 90 });
  const cardOffsets = [
    [10, 120, 440],
    [45, 210, 470],
    [90, 290, 500]
  ];
  for (const [x1, x2, y] of cardOffsets) {
    const row = new MockElement('div', {}, { left: 0, top: y, width: 400, height: 28 });
    row.append(
      new MockElement('div', {}, { left: x1, top: y, width: 80, height: 24 }).append('Tag'),
      new MockElement('div', {}, { left: x2, top: y, width: 80, height: 24 }).append('Desc')
    );
    misalignedCards.append(row);
  }

  body.append(nativeWrapper, ariaGrid, cssGrid, flexTable, misalignedCards);

  let contentMessageListener = null;
  const doc = {
    title: 'Test Page',
    body,
    documentElement: body,
    querySelectorAll: (sel) => body.querySelectorAll(sel),
    getElementById: () => null,
    addEventListener: () => {},
    removeEventListener: () => {}
  };

  const contentSandbox = {
    console,
    Set,
    Map,
    Array,
    Object,
    String,
    Number,
    Math,
    Infinity,
    parseInt,
    parseFloat,
    setTimeout,
    clearTimeout,
    requestAnimationFrame: (fn) => fn(),
    cancelAnimationFrame: () => {},
    Node: NodeConsts,
    Element: MockElement,
    document: doc,
    window: null,
    chrome: {
      runtime: {
        onMessage: {
          addListener: (fn) => {
            contentMessageListener = fn;
          }
        },
        sendMessage: () => {}
      }
    }
  };
  contentSandbox.window = contentSandbox;
  contentSandbox.getComputedStyle = (el) => el.computedStyle;

  vm.runInNewContext(fs.readFileSync(path.join(rootDir, 'content.js'), 'utf8'), contentSandbox);

  let scanResult = null;
  contentMessageListener({ action: 'EXTRACT_ALL_TABLES' }, {}, (res) => {
    scanResult = res;
  });

  assert(scanResult && scanResult.ok, 'Content scan failed');
  assert.strictEqual(scanResult.tables.length, 4, `Expected 4 tables, got ${scanResult.tables.length}`);
  assert.strictEqual(scanResult.tables[0].name, 'Native Sales');
  assert.deepStrictEqual(Array.from(scanResult.tables[0].rows[1]), ['Widget', '42']);
  assert.strictEqual(scanResult.tables[1].name, 'ARIA Metrics');
  assert.deepStrictEqual(Array.from(scanResult.tables[1].rows[1]), ['Latency', '12ms']);
  assert.strictEqual(scanResult.tables[2].name, 'Pricing Grid');
  assert.deepStrictEqual(Array.from(scanResult.tables[2].rows[2]), ['Team', '$50']);
  assert.strictEqual(scanResult.tables[3].name, 'Team Directory');
  assert.deepStrictEqual(Array.from(scanResult.tables[3].rows[2]), ['Bob', 'Designer']);
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
