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

  console.log('All verification checks passed.');
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
