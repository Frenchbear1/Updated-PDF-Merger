const { app, BrowserWindow, dialog, ipcMain, Menu, nativeImage } = require('electron');
const path = require('path');
const fsSync = require('fs');
const fs = require('fs/promises');
const { spawn } = require('child_process');
const os = require('os');
const { imageSize } = require('image-size');
const JSZip = require('jszip');
const { PDFArray, PDFDocument, PDFHexString, PDFName, PDFNumber } = require('pdf-lib');
const { default: Automizer } = require('pptx-automizer');
const PptxGenJS = require('pptxgenjs');

const SPLIT_TRIGGER_BYTES = 500 * 1024 * 1024;
const QPDF_VERSION = '12.3.2';
const QPDF_FOLDER = `qpdf-${QPDF_VERSION}-mingw64`;
const QPDF_DOWNLOAD_URL = `https://github.com/qpdf/qpdf/releases/download/v${QPDF_VERSION}/${QPDF_FOLDER}.zip`;
const MAX_INPUTS_PER_MERGE_CALL = 24;
const DIRECT_ARG_CHAR_LIMIT = 26000;
const FAST_FREE_MEMORY_RATIO = 0.22;
const FAST_TOTAL_MEMORY_RATIO = 0.08;
const MIN_FAST_MEMORY_BYTES = 384 * 1024 * 1024;
const MAX_FAST_MEMORY_BYTES = 1200 * 1024 * 1024;
const APP_ICON_ICO = path.join(__dirname, 'assets', 'app-icon.ico');
const APP_ICON_PNG = path.join(__dirname, 'assets', 'app-icon.png');
const TEMP_ROOT = path.join(app.getPath('temp'), 'updated-file-merger');
const BUNDLED_TOOLS_ROOT = app.isPackaged
  ? path.join(process.resourcesPath, 'tools')
  : path.join(__dirname, 'tools');
const CACHE_TOOLS_ROOT = path.join(app.getPath('userData'), 'tools');
const BUNDLED_QPDF_EXE = path.join(BUNDLED_TOOLS_ROOT, QPDF_FOLDER, 'bin', 'qpdf.exe');
const CACHE_QPDF_EXE = path.join(CACHE_TOOLS_ROOT, QPDF_FOLDER, 'bin', 'qpdf.exe');
const DEFAULT_OUTPUT_NAME = 'merged-output.pptx';
const DEFAULT_OUTPUT_FORMAT = 'pptx';
const DEFAULT_SLIDE_SIZE = { width: 13.333, height: 7.5 };
const EMU_PER_INCH = 914400;
const IMAGE_PX_PER_INCH = 96;
const MIN_SLIDE_INCH = 1;
const MAX_SLIDE_INCH = 56;
const IMAGE_TEMPLATE_LABEL = '__image_template__';
const SUPPORTED_PDF_EXTENSIONS = new Set(['.pdf']);
const SUPPORTED_IMAGE_EXTENSIONS = new Set([
  '.png',
  '.jpg',
  '.jpeg',
  '.bmp',
  '.gif',
  '.webp',
  '.tif',
  '.tiff',
]);
const SUPPORTED_PRESENTATION_EXTENSIONS = new Set(['.pptx', '.ppt']);
const SUPPORTED_EXTENSIONS = new Set([
  ...SUPPORTED_PDF_EXTENSIONS,
  ...SUPPORTED_PRESENTATION_EXTENSIONS,
  ...SUPPORTED_IMAGE_EXTENSIONS,
]);

let mainWindow;
let activeMerge = null;
let qpdfSetupPromise = null;

function getAppIconPath() {
  const candidates = [APP_ICON_ICO, APP_ICON_PNG];
  for (const p of candidates) {
    if (!fsSync.existsSync(p)) continue;
    try {
      const img = nativeImage.createFromPath(p);
      if (!img.isEmpty()) return p;
    } catch {
      // Try next candidate
    }
  }
  return undefined;
}

if (process.platform === 'win32') {
  app.setAppUserModelId('com.labar.filemerger');
}

function sendWindowState() {
  if (mainWindow && !mainWindow.isDestroyed()) {
    mainWindow.webContents.send('window-state', {
      maximized: mainWindow.isMaximized(),
    });
  }
}

function createWindow() {
  mainWindow = new BrowserWindow({
    width: 1180,
    height: 760,
    minWidth: 980,
    minHeight: 640,
    backgroundColor: '#111417',
    frame: false,
    titleBarStyle: 'hidden',
    autoHideMenuBar: true,
    icon: getAppIconPath(),
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      contextIsolation: true,
      nodeIntegration: false,
    },
  });

  mainWindow.setMenuBarVisibility(false);
  mainWindow.removeMenu();
  mainWindow.loadFile(path.join(__dirname, 'index.html'));
  if (process.platform === 'win32') {
    const iconPath = getAppIconPath();
    if (iconPath) {
      try {
        mainWindow.setIcon(iconPath);
      } catch {
        // Do not fail app startup if icon loading fails.
      }
    }
  }
  mainWindow.on('maximize', sendWindowState);
  mainWindow.on('unmaximize', sendWindowState);
  mainWindow.webContents.on('did-finish-load', sendWindowState);
}

function sendMergeProgress(payload) {
  if (mainWindow && !mainWindow.isDestroyed()) {
    mainWindow.webContents.send('merge-progress', payload);
  }
}

function runProcess(command, args, state, options = {}) {
  const successCodes = Array.isArray(options.successCodes) && options.successCodes.length
    ? options.successCodes
    : [0];

  return new Promise((resolve, reject) => {
    const child = spawn(command, args, {
      windowsHide: true,
      stdio: ['ignore', 'pipe', 'pipe'],
    });

    if (state) {
      state.activeChild = child;
    }

    let stdout = '';
    let stderr = '';

    child.stdout.on('data', (d) => {
      stdout += String(d);
    });

    child.stderr.on('data', (d) => {
      stderr += String(d);
    });

    child.on('error', (err) => {
      if (state && state.activeChild === child) {
        state.activeChild = null;
      }
      reject(err);
    });

    child.on('close', (code) => {
      if (state && state.activeChild === child) {
        state.activeChild = null;
      }
      if (successCodes.includes(code)) {
        resolve({ stdout, stderr, code });
      } else {
        const err = new Error(stderr || stdout || `${command} exited with code ${code}`);
        err.code = code;
        reject(err);
      }
    });
  });
}

async function downloadFile(url, outputPath) {
  const res = await fetch(url, {
    headers: {
      'User-Agent': 'UpdatedFileMerger',
    },
  });

  if (!res.ok) {
    throw new Error(`Failed to download qpdf (${res.status} ${res.statusText})`);
  }

  const arrayBuffer = await res.arrayBuffer();
  await fs.writeFile(outputPath, Buffer.from(arrayBuffer));
}

async function ensureQpdf() {
  try {
    await fs.access(BUNDLED_QPDF_EXE);
    return BUNDLED_QPDF_EXE;
  } catch {
    // continue
  }

  try {
    await fs.access(CACHE_QPDF_EXE);
    return CACHE_QPDF_EXE;
  } catch {
    // continue
  }

  if (!qpdfSetupPromise) {
    qpdfSetupPromise = (async () => {
      await fs.mkdir(CACHE_TOOLS_ROOT, { recursive: true });
      const zipPath = path.join(CACHE_TOOLS_ROOT, `${QPDF_FOLDER}.zip`);

      await downloadFile(QPDF_DOWNLOAD_URL, zipPath);
      try {
        await runProcess('tar', ['-xf', zipPath, '-C', CACHE_TOOLS_ROOT]);
      } catch {
        const psZip = zipPath.replace(/'/g, "''");
        const psOut = CACHE_TOOLS_ROOT.replace(/'/g, "''");
        await runProcess('powershell.exe', [
          '-NoProfile',
          '-ExecutionPolicy',
          'Bypass',
          '-Command',
          `Expand-Archive -Path '${psZip}' -DestinationPath '${psOut}' -Force`,
        ]);
      }
      await fs.rm(zipPath, { force: true });

      await fs.access(CACHE_QPDF_EXE);
      return CACHE_QPDF_EXE;
    })().finally(() => {
      qpdfSetupPromise = null;
    });
  }

  return qpdfSetupPromise;
}

async function qpdfGetPageCount(qpdfPath, pdfPath, state) {
  const { stdout } = await runProcess(qpdfPath, ['--show-npages', pdfPath], state, { successCodes: [0, 3] });
  const parsed = Number.parseInt(String(stdout).trim(), 10);
  if (!Number.isFinite(parsed) || parsed <= 0) {
    throw new Error(`Unable to determine page count for ${path.basename(pdfPath)}.`);
  }
  return parsed;
}

async function qpdfMakeRangePdf(qpdfPath, srcPath, pageRange, outPath, state) {
  await runProcess(qpdfPath, ['--empty', '--pages', srcPath, pageRange, '--', outPath], state, { successCodes: [0, 3] });
}

async function qpdfMergeFiles(qpdfPath, inputPaths, outPath, state) {
  if (!inputPaths.length) {
    throw new Error('No input files for qpdf merge.');
  }
  await runProcess(qpdfPath, ['--empty', '--pages', ...inputPaths, '--', outPath], state, { successCodes: [0, 3] });
}

function getInputKind(filePath) {
  const ext = path.extname(String(filePath || '')).toLowerCase();
  if (SUPPORTED_PDF_EXTENSIONS.has(ext)) return 'pdf';
  if (ext === '.pptx') return 'pptx';
  if (ext === '.ppt') return 'ppt';
  if (SUPPORTED_IMAGE_EXTENSIONS.has(ext)) return 'image';
  return null;
}

function isSupportedPath(filePath) {
  const ext = path.extname(String(filePath || '')).toLowerCase();
  return SUPPORTED_EXTENSIONS.has(ext);
}

function sanitizeOutputName(outputName, outputFormat) {
  const requestedFormat = String(outputFormat || DEFAULT_OUTPUT_FORMAT).toLowerCase() === 'pdf' ? 'pdf' : 'pptx';
  const sanitized = String(outputName || 'merged-output').trim().replace(/[<>:"/\\|?*]+/g, '-');
  const withoutKnownExt = sanitized.replace(/\.(pptx|ppt|pdf)$/i, '');
  const base = withoutKnownExt.trim() || 'merged-output';
  return requestedFormat === 'pdf' ? `${base}.pdf` : `${base}.pptx`;
}

function maybeCancelError(state) {
  if (!state || !state.cancelRequested) return null;
  const err = new Error('Merge canceled by user.');
  err.code = 'MERGE_CANCELED';
  return err;
}

function assertNotCanceled(state) {
  const err = maybeCancelError(state);
  if (err) throw err;
}

function killActiveChild(state) {
  const child = state?.activeChild;
  if (!child) return;

  try {
    child.kill('SIGTERM');
  } catch {
    // ignore
  }

  if (process.platform === 'win32' && child.pid) {
    spawn('taskkill', ['/PID', String(child.pid), '/T', '/F'], {
      windowsHide: true,
      stdio: 'ignore',
    });
  }
}

function emitStatus(state, filesTotal) {
  sendMergeProgress({
    completed: state.completedFiles,
    total: filesTotal,
    percent: Math.min(95, Math.round((state.completedFiles / Math.max(1, filesTotal)) * 95)),
    etaSeconds: 0,
    fileName: state.currentFileName,
    phase: 'merging',
  });
}

function estimateDirectArgChars(inputPaths, outputPath) {
  const parts = ['--empty', '--pages', ...inputPaths, '--', outputPath];
  return parts.reduce((sum, part) => sum + String(part || '').length + 3, 0);
}

function computeFastMemoryLimitBytes() {
  const freeBytes = Math.max(0, Number(os.freemem() || 0));
  const totalBytes = Math.max(0, Number(os.totalmem() || 0));
  const byFree = freeBytes > 0 ? Math.floor(freeBytes * FAST_FREE_MEMORY_RATIO) : Number.POSITIVE_INFINITY;
  const byTotal = totalBytes > 0 ? Math.floor(totalBytes * FAST_TOTAL_MEMORY_RATIO) : Number.POSITIVE_INFINITY;
  const candidate = Math.min(byFree, byTotal, MAX_FAST_MEMORY_BYTES);
  return Math.max(MIN_FAST_MEMORY_BYTES, candidate);
}

function computePdfMergePlan(files, outputPath) {
  const validFiles = Array.isArray(files) ? files : [];
  const inputPaths = validFiles.map((f) => f.path).filter(Boolean);
  const totalBytes = validFiles.reduce((sum, f) => sum + Math.max(1, Number(f?.size || 1)), 0);
  const memoryLimitBytes = computeFastMemoryLimitBytes();
  const argChars = estimateDirectArgChars(inputPaths, outputPath);

  if (inputPaths.length === 0) {
    return {
      mode: 'none',
      totalBytes: 0,
      memoryLimitBytes,
      argChars,
      argLimitChars: DIRECT_ARG_CHAR_LIMIT,
      reason: 'no_files',
    };
  }

  if (argChars > DIRECT_ARG_CHAR_LIMIT) {
    return {
      mode: 'slow',
      totalBytes,
      memoryLimitBytes,
      argChars,
      argLimitChars: DIRECT_ARG_CHAR_LIMIT,
      reason: 'command_length',
    };
  }

  if (totalBytes > memoryLimitBytes) {
    return {
      mode: 'slow',
      totalBytes,
      memoryLimitBytes,
      argChars,
      argLimitChars: DIRECT_ARG_CHAR_LIMIT,
      reason: 'memory',
    };
  }

  return {
    mode: 'fast',
    totalBytes,
    memoryLimitBytes,
    argChars,
    argLimitChars: DIRECT_ARG_CHAR_LIMIT,
    reason: 'within_limits',
  };
}

function summarizeInputKinds(files) {
  const summary = {
    total: 0,
    pdf: 0,
    ppt: 0,
    image: 0,
    unsupported: 0,
  };

  for (const file of Array.isArray(files) ? files : []) {
    const kind = getInputKind(file?.path || file?.name);
    if (!kind) {
      summary.unsupported += 1;
      continue;
    }

    summary.total += 1;
    if (kind === 'pdf') summary.pdf += 1;
    if (kind === 'image') summary.image += 1;
    if (kind === 'ppt' || kind === 'pptx') summary.ppt += 1;
  }

  summary.allPdf = summary.total > 0 && summary.pdf === summary.total;
  summary.hasPdf = summary.pdf > 0;
  summary.hasNonPdf = summary.ppt > 0 || summary.image > 0;
  summary.mixedPdfWithOther = summary.hasPdf && summary.hasNonPdf;
  return summary;
}

function clampSlideSize(value) {
  const n = Number(value);
  if (!Number.isFinite(n) || n <= 0) return MIN_SLIDE_INCH;
  return Math.max(MIN_SLIDE_INCH, Math.min(MAX_SLIDE_INCH, n));
}

function normalizeSlideSize(width, height) {
  let w = Number(width);
  let h = Number(height);
  if (!Number.isFinite(w) || !Number.isFinite(h) || w <= 0 || h <= 0) {
    return { ...DEFAULT_SLIDE_SIZE };
  }

  if (w > MAX_SLIDE_INCH || h > MAX_SLIDE_INCH) {
    const scaleDown = Math.min(MAX_SLIDE_INCH / w, MAX_SLIDE_INCH / h);
    w *= scaleDown;
    h *= scaleDown;
  }

  if (w < MIN_SLIDE_INCH || h < MIN_SLIDE_INCH) {
    const scaleUp = Math.max(MIN_SLIDE_INCH / w, MIN_SLIDE_INCH / h);
    w *= scaleUp;
    h *= scaleUp;
  }

  return {
    width: clampSlideSize(w),
    height: clampSlideSize(h),
  };
}

function getImageDimensions(imagePath) {
  const dimensions = imageSize(imagePath);
  const width = Number(dimensions?.width || 0);
  const height = Number(dimensions?.height || 0);
  if (!Number.isFinite(width) || !Number.isFinite(height) || width <= 0 || height <= 0) {
    throw new Error(`Unable to read image dimensions for ${path.basename(imagePath)}.`);
  }
  return { width, height };
}

function getSlideSizeFromImage(imagePath) {
  const dimensions = getImageDimensions(imagePath);
  return normalizeSlideSize(
    dimensions.width / IMAGE_PX_PER_INCH,
    dimensions.height / IMAGE_PX_PER_INCH
  );
}

function computeCoverRect(imageWidthPx, imageHeightPx, slideWidthIn, slideHeightIn) {
  const width = Number(imageWidthPx);
  const height = Number(imageHeightPx);
  if (!Number.isFinite(width) || !Number.isFinite(height) || width <= 0 || height <= 0) {
    return { x: 0, y: 0, w: slideWidthIn, h: slideHeightIn };
  }

  const imageRatio = width / height;
  const slideRatio = slideWidthIn / slideHeightIn;

  if (imageRatio >= slideRatio) {
    const h = slideHeightIn;
    const w = h * imageRatio;
    return {
      x: (slideWidthIn - w) / 2,
      y: 0,
      w,
      h,
    };
  }

  const w = slideWidthIn;
  const h = w / imageRatio;
  return {
    x: 0,
    y: (slideHeightIn - h) / 2,
    w,
    h,
  };
}

async function readPptxSlideSize(pptxPath) {
  const buffer = await fs.readFile(pptxPath);
  const zip = await JSZip.loadAsync(buffer);
  const file = zip.file('ppt/presentation.xml');
  if (!file) return { ...DEFAULT_SLIDE_SIZE };

  const xml = await file.async('string');
  const sldTagMatch = xml.match(/<p:sldSz\b[^>]*>/i);
  if (!sldTagMatch) return { ...DEFAULT_SLIDE_SIZE };

  const tag = sldTagMatch[0];
  const cxMatch = tag.match(/\bcx="(\d+)"/i);
  const cyMatch = tag.match(/\bcy="(\d+)"/i);

  const cx = Number(cxMatch?.[1] || 0);
  const cy = Number(cyMatch?.[1] || 0);
  if (!Number.isFinite(cx) || !Number.isFinite(cy) || cx <= 0 || cy <= 0) {
    return { ...DEFAULT_SLIDE_SIZE };
  }

  return normalizeSlideSize(cx / EMU_PER_INCH, cy / EMU_PER_INCH);
}

async function createBlankTemplatePptx(outPath, slideSize) {
  const pptx = new PptxGenJS();
  const customLayout = 'CUSTOM_LAYOUT';
  pptx.defineLayout({
    name: customLayout,
    width: slideSize.width,
    height: slideSize.height,
  });
  pptx.layout = customLayout;
  pptx.addSlide();
  await pptx.writeFile({ fileName: outPath });
}

async function convertPptToPptx(inputPath, outputPath, state) {
  const psIn = String(inputPath).replace(/'/g, "''");
  const psOut = String(outputPath).replace(/'/g, "''");
  const MSO_FALSE = '0';
  const PP_ALERTS_NONE = '1';
  const command = [
    "$ErrorActionPreference = 'Stop'",
    '$ppt = $null',
    '$presentation = $null',
    'try {',
    '  $ppt = New-Object -ComObject PowerPoint.Application',
    `  $ppt.DisplayAlerts = ${PP_ALERTS_NONE}`,
    `  $presentation = $ppt.Presentations.Open('${psIn}', ${MSO_FALSE}, ${MSO_FALSE}, ${MSO_FALSE})`,
    `  $presentation.SaveAs('${psOut}', 24)`,
    '} finally {',
    '  if ($presentation -ne $null) { $presentation.Close() }',
    '  if ($ppt -ne $null) { $ppt.Quit() }',
    '}',
  ].join('; ');

  try {
    await runProcess(
      'powershell.exe',
      ['-NoProfile', '-ExecutionPolicy', 'Bypass', '-Command', command],
      state
    );
  } catch (err) {
    const wrapped = new Error(
      `Unable to convert "${path.basename(inputPath)}" from .ppt to .pptx. This feature requires Microsoft PowerPoint on Windows.`
    );
    wrapped.cause = err;
    throw wrapped;
  }
}

async function convertPptxToPdf(inputPath, outputPath, state) {
  const psIn = String(inputPath).replace(/'/g, "''");
  const psOut = String(outputPath).replace(/'/g, "''");
  const MSO_FALSE = '0';
  const PP_ALERTS_NONE = '1';
  const command = [
    "$ErrorActionPreference = 'Stop'",
    '$ppt = $null',
    '$presentation = $null',
    'try {',
    '  $ppt = New-Object -ComObject PowerPoint.Application',
    `  $ppt.DisplayAlerts = ${PP_ALERTS_NONE}`,
    `  $presentation = $ppt.Presentations.Open('${psIn}', ${MSO_FALSE}, ${MSO_FALSE}, ${MSO_FALSE})`,
    `  $presentation.SaveAs('${psOut}', 32)`,
    '} finally {',
    '  if ($presentation -ne $null) { $presentation.Close() }',
    '  if ($ppt -ne $null) { $ppt.Quit() }',
    '}',
  ].join('; ');

  try {
    await runProcess(
      'powershell.exe',
      ['-NoProfile', '-ExecutionPolicy', 'Bypass', '-Command', command],
      state
    );
  } catch (err) {
    const wrapped = new Error(
      'Unable to export PDF. PDF export requires Microsoft PowerPoint on Windows.'
    );
    wrapped.cause = err;
    throw wrapped;
  }
}

async function getPdfImagePayload(imagePath) {
  const ext = path.extname(String(imagePath || '')).toLowerCase();

  if (ext === '.jpg' || ext === '.jpeg') {
    const dimensions = getImageDimensions(imagePath);
    return {
      width: dimensions.width,
      height: dimensions.height,
      format: 'jpg',
      buffer: await fs.readFile(imagePath),
    };
  }

  if (ext === '.png') {
    const dimensions = getImageDimensions(imagePath);
    return {
      width: dimensions.width,
      height: dimensions.height,
      format: 'png',
      buffer: await fs.readFile(imagePath),
    };
  }

  // Normalize other image formats (webp, tiff, etc.) to PNG for PDF embedding.
  const image = nativeImage.createFromPath(imagePath);
  if (image.isEmpty()) {
    throw new Error(`Unsupported image format for PDF output: ${path.basename(imagePath)}.`);
  }
  const size = image.getSize();
  const width = Number(size?.width || 0);
  const height = Number(size?.height || 0);
  if (!Number.isFinite(width) || !Number.isFinite(height) || width <= 0 || height <= 0) {
    throw new Error(`Unable to decode image for PDF output: ${path.basename(imagePath)}.`);
  }

  return {
    width,
    height,
    format: 'png',
    buffer: image.toPNG(),
  };
}

async function writeImagesToPdf(imageFiles, outputPath, state, filesTotal) {
  const pdfDoc = await PDFDocument.create();

  for (const item of imageFiles) {
    assertNotCanceled(state);
    state.currentFileName = `Adding image page: ${item.name}`;
    emitStatus(state, filesTotal);

    const payload = await getPdfImagePayload(item.path);
    const page = pdfDoc.addPage([payload.width, payload.height]);
    const embedded = payload.format === 'jpg'
      ? await pdfDoc.embedJpg(payload.buffer)
      : await pdfDoc.embedPng(payload.buffer);

    page.drawImage(embedded, {
      x: 0,
      y: 0,
      width: payload.width,
      height: payload.height,
    });

    state.completedFiles += 1;
    emitStatus(state, filesTotal);
  }

  assertNotCanceled(state);
  sendMergeProgress({
    completed: filesTotal,
    total: filesTotal,
    percent: 99,
    etaSeconds: 0,
    fileName: 'Writing .pdf file',
    phase: 'writing',
  });

  const bytes = await pdfDoc.save();
  await fs.writeFile(outputPath, bytes);
}

async function writeSingleImageToPdf(imageFile, outputPath) {
  const payload = await getPdfImagePayload(imageFile.path);
  const pdfDoc = await PDFDocument.create();
  const page = pdfDoc.addPage([payload.width, payload.height]);
  const embedded = payload.format === 'jpg'
    ? await pdfDoc.embedJpg(payload.buffer)
    : await pdfDoc.embedPng(payload.buffer);

  page.drawImage(embedded, {
    x: 0,
    y: 0,
    width: payload.width,
    height: payload.height,
  });

  const bytes = await pdfDoc.save();
  await fs.writeFile(outputPath, bytes);
}

function buildDeckBookmarkTitle(fileName) {
  const base = path.basename(String(fileName || '').trim());
  if (!base) return 'Deck';
  return base.replace(/\.[^/.\\]+$/, '') || base;
}

async function addPdfDeckBookmarks(pdfPath, entries) {
  const rawEntries = Array.isArray(entries) ? entries : [];
  if (!rawEntries.length) return;

  const sourceBytes = await fs.readFile(pdfPath);
  const pdfDoc = await PDFDocument.load(sourceBytes, { updateMetadata: false });
  const pages = pdfDoc.getPages();
  if (!pages.length) return;

  const normalizedEntries = rawEntries
    .map((entry) => ({
      title: String(entry?.title || '').trim(),
      startPageIndex: Number(entry?.startPageIndex || 0),
    }))
    .filter((entry) => entry.title.length > 0)
    .filter((entry) => Number.isFinite(entry.startPageIndex))
    .map((entry) => ({
      title: entry.title,
      startPageIndex: Math.max(0, Math.min(pages.length - 1, Math.floor(entry.startPageIndex))),
    }))
    .sort((a, b) => a.startPageIndex - b.startPageIndex);

  if (!normalizedEntries.length) return;

  const deduped = [];
  let lastPage = -1;
  for (const entry of normalizedEntries) {
    if (entry.startPageIndex === lastPage) continue;
    deduped.push(entry);
    lastPage = entry.startPageIndex;
  }
  if (!deduped.length) return;

  const context = pdfDoc.context;
  const outlines = context.obj({});
  const outlinesRef = context.register(outlines);

  let firstRef = null;
  let lastRef = null;
  let previousRef = null;

  for (const entry of deduped) {
    const dest = PDFArray.withContext(context);
    dest.push(pages[entry.startPageIndex].ref);
    dest.push(PDFName.of('Fit'));

    const item = context.obj({});
    item.set(PDFName.of('Title'), PDFHexString.fromText(entry.title));
    item.set(PDFName.of('Parent'), outlinesRef);
    item.set(PDFName.of('Dest'), dest);
    if (previousRef) {
      item.set(PDFName.of('Prev'), previousRef);
    }

    const itemRef = context.register(item);
    if (!firstRef) firstRef = itemRef;
    if (previousRef) {
      const previous = context.lookup(previousRef);
      previous.set(PDFName.of('Next'), itemRef);
    }
    previousRef = itemRef;
    lastRef = itemRef;
  }

  outlines.set(PDFName.of('Type'), PDFName.of('Outlines'));
  outlines.set(PDFName.of('First'), firstRef);
  outlines.set(PDFName.of('Last'), lastRef);
  outlines.set(PDFName.of('Count'), PDFNumber.of(deduped.length));

  pdfDoc.catalog.set(PDFName.of('Outlines'), outlinesRef);
  pdfDoc.catalog.set(PDFName.of('PageMode'), PDFName.of('UseOutlines'));

  const saved = await pdfDoc.save({
    useObjectStreams: false,
    updateFieldAppearances: false,
  });
  await fs.writeFile(pdfPath, saved);
}

async function preparePresentationInputs(files, state, tempDir, filesTotal) {
  const prepared = [];

  for (let i = 0; i < files.length; i += 1) {
    assertNotCanceled(state);
    const file = files[i];
    const kind = getInputKind(file.path);
    if (!kind) {
      throw new Error(`Unsupported file type: ${file.name}`);
    }

    if (kind === 'pdf') {
      throw new Error('PDF files can only be merged in a PDF-only batch.');
    }

    if (kind === 'ppt') {
      const convertedPath = path.join(tempDir, `converted-${String(i + 1).padStart(4, '0')}.pptx`);
      state.currentFileName = `Converting ${file.name} to PPTX`;
      emitStatus(state, filesTotal);
      await convertPptToPptx(file.path, convertedPath, state);
      prepared.push({
        name: file.name,
        path: convertedPath,
        kind: 'pptx',
        sourceKind: 'ppt',
      });
      continue;
    }

    prepared.push({
      name: file.name,
      path: file.path,
      kind,
      sourceKind: kind,
    });
  }

  return prepared;
}

function computeMergePlan(files, outputName, outputFormat) {
  const inputFiles = Array.isArray(files) ? files : [];
  if (!inputFiles.length) {
    return {
      mode: 'none',
      reason: 'no_files',
    };
  }

  const summary = summarizeInputKinds(inputFiles);
  const requestedFormat = String(outputFormat || DEFAULT_OUTPUT_FORMAT).toLowerCase() === 'pdf' ? 'pdf' : 'pptx';
  if (summary.unsupported > 0) {
    return {
      mode: 'unsupported',
      reason: 'unsupported_files',
      unsupported: summary.unsupported,
    };
  }

  if (summary.mixedPdfWithOther) {
    if (requestedFormat !== 'pdf') {
      return {
        mode: 'unsupported',
        reason: 'mixed_requires_pdf',
      };
    }

    return {
      mode: 'ready',
      reason: 'mixed_to_pdf',
      route: 'mixed',
    };
  }

  if (summary.allPdf) {
    const defaultOutput = path.join(process.cwd(), sanitizeOutputName(outputName || DEFAULT_OUTPUT_NAME, 'pdf'));
    const pdfPlan = computePdfMergePlan(inputFiles, defaultOutput);
    return {
      ...pdfPlan,
      route: 'pdf',
    };
  }

  return {
    mode: 'ready',
    reason: 'supported',
    route: 'presentation',
  };
}

async function mergePdfOnlyFiles(files, finalOutputPath, state, tempDir, options = {}) {
  const filesTotal = files.length;
  const emitDone = options.emitDone !== false;
  const qpdfPath = await ensureQpdf();
  const plan = computePdfMergePlan(files, finalOutputPath);
  let currentMergedPath = null;
  let mergeStep = 0;
  let pendingSmallFiles = [];

  emitStatus(state, filesTotal);

  const flushPending = async () => {
    if (!pendingSmallFiles.length) return;

    const chunks = [];
    for (let i = 0; i < pendingSmallFiles.length; i += MAX_INPUTS_PER_MERGE_CALL) {
      chunks.push(pendingSmallFiles.slice(i, i + MAX_INPUTS_PER_MERGE_CALL));
    }

    for (const chunk of chunks) {
      assertNotCanceled(state);
      const chunkPaths = chunk.map((x) => x.path);

      if (!currentMergedPath && chunkPaths.length === 1) {
        currentMergedPath = chunkPaths[0];
      } else {
        const mergeInputs = currentMergedPath ? [currentMergedPath, ...chunkPaths] : chunkPaths;
        mergeStep += 1;
        const outPath = path.join(tempDir, `merged-step-${String(mergeStep).padStart(5, '0')}.pdf`);

        state.currentFileName = `Merging files (${state.completedFiles}/${filesTotal})`;
        emitStatus(state, filesTotal);

        await qpdfMergeFiles(qpdfPath, mergeInputs, outPath, state);

        if (currentMergedPath && currentMergedPath.startsWith(tempDir)) {
          await fs.rm(currentMergedPath, { force: true }).catch(() => {});
        }

        currentMergedPath = outPath;
      }

      state.completedFiles += chunk.length;
      emitStatus(state, filesTotal);
    }

    pendingSmallFiles = [];
  };

  const splitLargeFileToTemp = async (file, fileIndex) => {
    const pageCount = await qpdfGetPageCount(qpdfPath, file.path, state);
    const fileSize = Math.max(1, Number(file.size || 1));
    const bySize = Math.max(1, Math.ceil(fileSize / SPLIT_TRIGGER_BYTES));
    const pagesPerSegment = Math.max(1, Math.ceil(pageCount / bySize));

    const splitPaths = [];
    for (let seg = 0; seg < bySize; seg += 1) {
      assertNotCanceled(state);

      const start = seg * pagesPerSegment + 1;
      if (start > pageCount) break;
      const end = Math.min(pageCount, start + pagesPerSegment - 1);

      state.currentFileName = `Splitting ${file.name} (${start}-${end})`;
      emitStatus(state, filesTotal);

      const splitPath = path.join(
        tempDir,
        `split-${String(fileIndex + 1).padStart(4, '0')}-${String(seg + 1).padStart(4, '0')}.pdf`
      );
      await qpdfMakeRangePdf(qpdfPath, file.path, `${start}-${end}`, splitPath, state);
      splitPaths.push(splitPath);
    }

    return splitPaths;
  };

  if (plan.mode === 'fast') {
    state.currentFileName = 'Fast mode: merging directly';
    emitStatus(state, filesTotal);

    try {
      await qpdfMergeFiles(qpdfPath, files.map((f) => f.path), finalOutputPath, state);
      state.completedFiles = filesTotal;
      if (emitDone) {
        sendMergeProgress({
          completed: filesTotal,
          total: filesTotal,
          percent: 100,
          etaSeconds: 0,
          fileName: 'Finished',
          phase: 'done',
          done: true,
        });
      }
      return;
    } catch (err) {
      if ((err && err.code === 'MERGE_CANCELED') || state.cancelRequested) {
        throw err;
      }

      // Fallback if direct merge fails even though plan predicted fast mode.
      await fs.rm(finalOutputPath, { force: true }).catch(() => {});
      state.currentFileName = 'Switching to safe mode (slower)';
      emitStatus(state, filesTotal);
    }
  }

  for (let i = 0; i < files.length; i += 1) {
    assertNotCanceled(state);
    const file = files[i];
    const fileSize = Math.max(1, Number(file.size || 1));

    if (fileSize < SPLIT_TRIGGER_BYTES) {
      state.currentFileName = `Queued: ${file.name}`;
      emitStatus(state, filesTotal);

      pendingSmallFiles.push({ path: file.path, name: file.name });
      if (pendingSmallFiles.length >= MAX_INPUTS_PER_MERGE_CALL) {
        await flushPending();
      }
      continue;
    }

    // Merge all prior small files first, then process this large file only.
    await flushPending();

    const splitPaths = await splitLargeFileToTemp(file, i);

    if (splitPaths.length) {
      let remaining = [...splitPaths];

      while (remaining.length > 0) {
        assertNotCanceled(state);
        const capacity = currentMergedPath ? MAX_INPUTS_PER_MERGE_CALL - 1 : MAX_INPUTS_PER_MERGE_CALL;
        const chunk = remaining.splice(0, Math.max(1, capacity));
        mergeStep += 1;
        const outPath = path.join(tempDir, `merged-step-${String(mergeStep).padStart(5, '0')}.pdf`);
        const mergeInputs = currentMergedPath ? [currentMergedPath, ...chunk] : chunk;

        state.currentFileName = `Merging split parts: ${file.name}`;
        emitStatus(state, filesTotal);

        await qpdfMergeFiles(qpdfPath, mergeInputs, outPath, state);

        if (currentMergedPath && currentMergedPath.startsWith(tempDir)) {
          await fs.rm(currentMergedPath, { force: true }).catch(() => {});
        }

        currentMergedPath = outPath;
      }
    }

    for (const splitPath of splitPaths) {
      await fs.rm(splitPath, { force: true }).catch(() => {});
    }

    state.completedFiles += 1;
    emitStatus(state, filesTotal);
  }

  await flushPending();
  assertNotCanceled(state);

  if (!currentMergedPath) {
    throw new Error('No merged output was produced.');
  }

  sendMergeProgress({
    completed: filesTotal,
    total: filesTotal,
    percent: 99,
    etaSeconds: 0,
    fileName: 'Writing file to disk',
    phase: 'writing',
  });

  await fs.copyFile(currentMergedPath, finalOutputPath);

  if (emitDone) {
    sendMergeProgress({
      completed: filesTotal,
      total: filesTotal,
      percent: 100,
      etaSeconds: 0,
      fileName: 'Finished',
      phase: 'done',
      done: true,
    });
  }
}

async function mergeMixedFilesToPdf(files, finalOutputPath, state, tempDir) {
  const filesTotal = files.length;
  const conversionDir = path.join(tempDir, 'mixed-pdf-inputs');
  await fs.mkdir(conversionDir, { recursive: true });

  const pdfInputs = [];

  state.completedFiles = 0;
  emitStatus(state, filesTotal);

  for (let i = 0; i < files.length; i += 1) {
    assertNotCanceled(state);
    const file = files[i];
    const kind = getInputKind(file.path);
    if (!kind) {
      throw new Error(`Unsupported file type: ${file.name}`);
    }

    if (kind === 'pdf') {
      state.currentFileName = `Using PDF: ${file.name}`;
      emitStatus(state, filesTotal);
      pdfInputs.push({
        name: file.name,
        path: file.path,
        size: Number(file.size || 0),
      });
      state.completedFiles += 1;
      emitStatus(state, filesTotal);
      continue;
    }

    if (kind === 'image') {
      state.currentFileName = `Converting image to PDF: ${file.name}`;
      emitStatus(state, filesTotal);
      const outPath = path.join(conversionDir, `img-${String(i + 1).padStart(4, '0')}.pdf`);
      await writeSingleImageToPdf(file, outPath);
      const stats = await fs.stat(outPath);
      pdfInputs.push({
        name: file.name,
        path: outPath,
        size: Number(stats.size || 0),
      });
      state.completedFiles += 1;
      emitStatus(state, filesTotal);
      continue;
    }

    let pptxPath = file.path;
    if (kind === 'ppt') {
      state.currentFileName = `Converting ${file.name} to PPTX`;
      emitStatus(state, filesTotal);
      pptxPath = path.join(conversionDir, `deck-${String(i + 1).padStart(4, '0')}.pptx`);
      await convertPptToPptx(file.path, pptxPath, state);
    }

    state.currentFileName = `Converting ${file.name} to PDF`;
    emitStatus(state, filesTotal);
    const outPath = path.join(conversionDir, `deck-${String(i + 1).padStart(4, '0')}.pdf`);
    await convertPptxToPdf(pptxPath, outPath, state);
    const stats = await fs.stat(outPath);
    pdfInputs.push({
      name: file.name,
      path: outPath,
      size: Number(stats.size || 0),
    });
    state.completedFiles += 1;
    emitStatus(state, filesTotal);
  }

  assertNotCanceled(state);
  if (!pdfInputs.length) {
    throw new Error('No files were converted to PDF.');
  }

  state.completedFiles = 0;
  state.currentFileName = 'Merging converted PDF files';
  emitStatus(state, filesTotal);

  await mergePdfOnlyFiles(pdfInputs, finalOutputPath, state, tempDir, { emitDone: false });
  assertNotCanceled(state);

  const qpdfPath = await ensureQpdf();
  const bookmarkEntries = [];
  let pageCursor = 0;
  for (const item of pdfInputs) {
    const pageCount = await qpdfGetPageCount(qpdfPath, item.path, state);
    bookmarkEntries.push({
      title: buildDeckBookmarkTitle(item.name),
      startPageIndex: pageCursor,
    });
    pageCursor += pageCount;
  }

  if (bookmarkEntries.length > 0) {
    state.currentFileName = 'Adding PDF bookmarks';
    emitStatus(state, filesTotal);
    await addPdfDeckBookmarks(finalOutputPath, bookmarkEntries);
    assertNotCanceled(state);
  }

  sendMergeProgress({
    completed: filesTotal,
    total: filesTotal,
    percent: 100,
    etaSeconds: 0,
    fileName: 'Finished',
    phase: 'done',
    done: true,
  });
}

async function mergePresentationFiles(files, finalOutputPath, requestedFormat, state, tempDir) {
  const preparedFiles = await preparePresentationInputs(files, state, tempDir, files.length);
  assertNotCanceled(state);

  if (!preparedFiles.length) {
    throw new Error('No supported files were selected.');
  }

  if (requestedFormat === 'pdf' && preparedFiles.every((item) => item.kind === 'image')) {
    await writeImagesToPdf(preparedFiles, finalOutputPath, state, files.length);
    assertNotCanceled(state);

    sendMergeProgress({
      completed: files.length,
      total: files.length,
      percent: 100,
      etaSeconds: 0,
      fileName: 'Finished',
      phase: 'done',
      done: true,
    });
    return;
  }

  const firstItem = preparedFiles[0];
  let slideSize = { ...DEFAULT_SLIDE_SIZE };
  if (firstItem.kind === 'pptx') {
    state.currentFileName = `Reading slide size: ${firstItem.name}`;
    emitStatus(state, files.length);
    slideSize = await readPptxSlideSize(firstItem.path);
  } else if (firstItem.kind === 'image') {
    slideSize = getSlideSizeFromImage(firstItem.path);
  }

  const imageTemplatePath = path.join(tempDir, 'image-template.pptx');
  await createBlankTemplatePptx(imageTemplatePath, slideSize);

  const rootTemplatePath = firstItem.kind === 'pptx' ? firstItem.path : imageTemplatePath;
  const outputDir = requestedFormat === 'pdf' ? tempDir : path.dirname(finalOutputPath);
  const outputNameOnly = requestedFormat === 'pdf' ? 'merged-output.pptx' : path.basename(finalOutputPath);
  const mergedPptxPath = path.join(outputDir, outputNameOnly);

  const automizer = new Automizer({
    templateDir: process.cwd(),
    outputDir,
    removeExistingSlides: true,
    autoImportSlideMasters: true,
    cleanup: true,
    verbosity: 0,
  });

  let presentation = automizer
    .loadRoot(rootTemplatePath)
    .load(imageTemplatePath, IMAGE_TEMPLATE_LABEL);

  const deckBookmarks = [];
  let nextSlideIndex = 0;

  const templateLabelByPath = new Map();
  let templateIndex = 0;
  const ensureTemplateLabel = (filePath) => {
    if (templateLabelByPath.has(filePath)) {
      return templateLabelByPath.get(filePath);
    }
    templateIndex += 1;
    const label = `source_${String(templateIndex).padStart(3, '0')}`;
    templateLabelByPath.set(filePath, label);
    presentation = presentation.load(filePath, label);
    return label;
  };

  for (const item of preparedFiles) {
    assertNotCanceled(state);

    if (item.kind === 'pptx') {
      state.currentFileName = `Adding slides: ${item.name}`;
      emitStatus(state, files.length);

      const label = ensureTemplateLabel(item.path);
      const slideNumbers = await presentation.getTemplate(label).getAllSlideNumbers();
      if (slideNumbers.length > 0) {
        deckBookmarks.push({
          title: buildDeckBookmarkTitle(item.name),
          startPageIndex: nextSlideIndex,
        });
      }
      for (const slideNumber of slideNumbers) {
        assertNotCanceled(state);
        presentation = presentation.addSlide(label, slideNumber);
      }

      nextSlideIndex += slideNumbers.length;
      state.completedFiles += 1;
      emitStatus(state, files.length);
      continue;
    }

    if (item.kind === 'image') {
      state.currentFileName = `Adding image slide: ${item.name}`;
      emitStatus(state, files.length);

      const dimensions = getImageDimensions(item.path);
      const rect = computeCoverRect(
        dimensions.width,
        dimensions.height,
        slideSize.width,
        slideSize.height
      );

      presentation = presentation.addSlide(IMAGE_TEMPLATE_LABEL, 1, (slide) => {
        slide.generate((pptxSlide) => {
          pptxSlide.addImage({
            path: item.path,
            x: rect.x,
            y: rect.y,
            w: rect.w,
            h: rect.h,
          });
        }, `img_${state.completedFiles + 1}`);
      });

      nextSlideIndex += 1;
      state.completedFiles += 1;
      emitStatus(state, files.length);
    }
  }

  assertNotCanceled(state);
  sendMergeProgress({
    completed: files.length,
    total: files.length,
    percent: 99,
    etaSeconds: 0,
    fileName: requestedFormat === 'pdf' ? 'Writing temporary .pptx file' : 'Writing .pptx file',
    phase: 'writing',
  });

  await presentation.write(outputNameOnly);
  assertNotCanceled(state);

  if (requestedFormat === 'pdf') {
    state.currentFileName = 'Converting merged presentation to PDF';
    emitStatus(state, files.length);
    await convertPptxToPdf(mergedPptxPath, finalOutputPath, state);
    assertNotCanceled(state);

    if (deckBookmarks.length > 0) {
      state.currentFileName = 'Adding PDF bookmarks';
      emitStatus(state, files.length);
      await addPdfDeckBookmarks(finalOutputPath, deckBookmarks);
      assertNotCanceled(state);
    }
  }

  sendMergeProgress({
    completed: files.length,
    total: files.length,
    percent: 100,
    etaSeconds: 0,
    fileName: 'Finished',
    phase: 'done',
    done: true,
  });
}

app.whenReady().then(() => {
  Menu.setApplicationMenu(null);
  createWindow();

  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) {
      createWindow();
    }
  });
});

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') {
    app.quit();
  }
});

ipcMain.handle('pick-merge-files', async () => {
  const res = await dialog.showOpenDialog(mainWindow, {
    title: 'Select PDF, PowerPoint, or image files',
    properties: ['openFile', 'multiSelections'],
    filters: [
      {
        name: 'All Supported',
        extensions: ['pdf', 'pptx', 'ppt', 'png', 'jpg', 'jpeg', 'bmp', 'gif', 'webp', 'tif', 'tiff'],
      },
      {
        name: 'PDF Files',
        extensions: ['pdf'],
      },
      {
        name: 'PowerPoint Files',
        extensions: ['pptx', 'ppt'],
      },
      {
        name: 'Image Files',
        extensions: ['png', 'jpg', 'jpeg', 'bmp', 'gif', 'webp', 'tif', 'tiff'],
      },
    ],
  });

  if (res.canceled) return [];
  return res.filePaths;
});

ipcMain.handle('read-file-metadata', async (_, inputPaths) => {
  const results = [];

  for (const filePath of inputPaths) {
    const kind = getInputKind(filePath);
    if (!kind) continue;

    try {
      const stats = await fs.stat(filePath);
      const name = path.basename(filePath);
      results.push({
        id: `${name}-${stats.size}-${stats.mtimeMs}`,
        path: filePath,
        name,
        kind,
        size: stats.size,
        createdMs: stats.birthtimeMs || stats.ctimeMs,
        modifiedMs: stats.mtimeMs,
      });
    } catch {
      // Ignore unreadable files.
    }
  }

  return results;
});

ipcMain.handle('get-merge-plan', async (_, payload) => {
  const files = Array.isArray(payload?.files) ? payload.files : [];
  return computeMergePlan(files, payload?.outputName, payload?.outputFormat);
});

ipcMain.handle('cancel-merge', async () => {
  if (!activeMerge) return false;
  activeMerge.cancelRequested = true;
  killActiveChild(activeMerge);
  return true;
});

ipcMain.handle('window-minimize', () => {
  if (mainWindow && !mainWindow.isDestroyed()) {
    mainWindow.minimize();
  }
});

ipcMain.handle('window-toggle-maximize', () => {
  if (!mainWindow || mainWindow.isDestroyed()) return { maximized: false };
  if (mainWindow.isMaximized()) {
    mainWindow.unmaximize();
  } else {
    mainWindow.maximize();
  }
  return { maximized: mainWindow.isMaximized() };
});

ipcMain.handle('window-close', () => {
  if (mainWindow && !mainWindow.isDestroyed()) {
    mainWindow.close();
  }
});

ipcMain.handle('window-is-maximized', () => {
  if (!mainWindow || mainWindow.isDestroyed()) return { maximized: false };
  return { maximized: mainWindow.isMaximized() };
});

ipcMain.handle('merge-files', async (_, payload) => {
  const { files, outputName, outputFormat } = payload;
  if (!Array.isArray(files) || files.length === 0) {
    throw new Error('No files selected.');
  }
  if (activeMerge) {
    throw new Error('A merge is already in progress.');
  }

  const unsupportedFile = files.find((f) => !isSupportedPath(f.path));
  if (unsupportedFile) {
    throw new Error(`Unsupported file type: ${unsupportedFile.name}`);
  }

  const summary = summarizeInputKinds(files);
  const mergeRoute = summary.allPdf
    ? 'pdf'
    : (summary.mixedPdfWithOther ? 'mixed' : 'presentation');
  const requestedFormat = mergeRoute === 'presentation'
    ? (String(outputFormat || DEFAULT_OUTPUT_FORMAT).toLowerCase() === 'pdf' ? 'pdf' : 'pptx')
    : 'pdf';

  const save = await dialog.showSaveDialog(mainWindow, {
    title: requestedFormat === 'pdf' ? 'Save merged PDF' : 'Save merged PowerPoint',
    defaultPath: sanitizeOutputName(outputName || DEFAULT_OUTPUT_NAME, requestedFormat),
    filters: requestedFormat === 'pdf'
      ? [{ name: 'PDF Files', extensions: ['pdf'] }]
      : [{ name: 'PowerPoint Files', extensions: ['pptx'] }],
  });

  if (save.canceled || !save.filePath) {
    return { canceled: true };
  }
  const finalOutputPath = save.filePath;

  const state = {
    cancelRequested: false,
    completedFiles: 0,
    currentFileName: 'Preparing merge...',
    activeChild: null,
  };

  activeMerge = state;

  const mergeJobId = `${Date.now()}-${Math.random().toString(16).slice(2, 8)}`;
  const tempDir = path.join(TEMP_ROOT, mergeJobId);

  try {
    await fs.mkdir(tempDir, { recursive: true });
    emitStatus(state, files.length);

    if (mergeRoute === 'pdf') {
      await mergePdfOnlyFiles(files, finalOutputPath, state, tempDir);
    } else if (mergeRoute === 'mixed') {
      await mergeMixedFilesToPdf(files, finalOutputPath, state, tempDir);
    } else {
      await mergePresentationFiles(files, finalOutputPath, requestedFormat, state, tempDir);
    }

    return { canceled: false, outputPath: finalOutputPath };
  } catch (err) {
    if ((err && err.code === 'MERGE_CANCELED') || state.cancelRequested) {
      sendMergeProgress({
        completed: state.completedFiles,
        total: files.length,
        percent: 0,
        etaSeconds: 0,
        fileName: 'Canceled',
        phase: 'canceled',
      });
      return { canceled: true };
    }

    throw err;
  } finally {
    killActiveChild(state);
    await fs.rm(tempDir, { recursive: true, force: true }).catch(() => {});
    activeMerge = null;
  }
});
