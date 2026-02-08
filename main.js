const { app, BrowserWindow, dialog, ipcMain, Menu, nativeImage } = require('electron');
const path = require('path');
const fsSync = require('fs');
const fs = require('fs/promises');
const { spawn } = require('child_process');
const os = require('os');

const SPLIT_TRIGGER_BYTES = 500 * 1024 * 1024;
const TEMP_ROOT = path.join(__dirname, 'temp-processing');
const TOOLS_ROOT = path.join(__dirname, 'tools');
const QPDF_VERSION = '12.3.2';
const QPDF_FOLDER = `qpdf-${QPDF_VERSION}-mingw64`;
const QPDF_EXE = path.join(TOOLS_ROOT, QPDF_FOLDER, 'bin', 'qpdf.exe');
const QPDF_DOWNLOAD_URL = `https://github.com/qpdf/qpdf/releases/download/v${QPDF_VERSION}/${QPDF_FOLDER}.zip`;
const MAX_INPUTS_PER_MERGE_CALL = 24;
const DIRECT_ARG_CHAR_LIMIT = 26000;
const FAST_FREE_MEMORY_RATIO = 0.22;
const FAST_TOTAL_MEMORY_RATIO = 0.08;
const MIN_FAST_MEMORY_BYTES = 384 * 1024 * 1024;
const MAX_FAST_MEMORY_BYTES = 1200 * 1024 * 1024;
const APP_ICON_ICO = path.join(__dirname, 'assets', 'app-icon.ico');
const APP_ICON_PNG = path.join(__dirname, 'assets', 'app-icon.png');

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
  app.setAppUserModelId('com.labar.pdfmerger');
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
      'User-Agent': 'UpdatedPDFMerger',
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
    await fs.access(QPDF_EXE);
    return QPDF_EXE;
  } catch {
    // continue
  }

  if (!qpdfSetupPromise) {
    qpdfSetupPromise = (async () => {
      await fs.mkdir(TOOLS_ROOT, { recursive: true });
      const zipPath = path.join(TOOLS_ROOT, `${QPDF_FOLDER}.zip`);

      await downloadFile(QPDF_DOWNLOAD_URL, zipPath);
      try {
        await runProcess('tar', ['-xf', zipPath, '-C', TOOLS_ROOT]);
      } catch {
        const psZip = zipPath.replace(/'/g, "''");
        const psOut = TOOLS_ROOT.replace(/'/g, "''");
        await runProcess('powershell.exe', [
          '-NoProfile',
          '-ExecutionPolicy',
          'Bypass',
          '-Command',
          `Expand-Archive -Path '${psZip}' -DestinationPath '${psOut}' -Force`,
        ]);
      }
      await fs.rm(zipPath, { force: true });

      await fs.access(QPDF_EXE);
      return QPDF_EXE;
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

function sanitizeOutputName(outputName) {
  const sanitized = String(outputName || 'merged-output').trim().replace(/[<>:"/\\|?*]+/g, '-');
  return sanitized.toLowerCase().endsWith('.pdf') ? sanitized : `${sanitized}.pdf`;
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

function computeMergePlan(files, outputPath) {
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

ipcMain.handle('pick-pdf-files', async () => {
  const res = await dialog.showOpenDialog(mainWindow, {
    title: 'Select PDF files',
    properties: ['openFile', 'multiSelections'],
    filters: [{ name: 'PDF Files', extensions: ['pdf'] }],
  });

  if (res.canceled) return [];
  return res.filePaths;
});

ipcMain.handle('read-file-metadata', async (_, inputPaths) => {
  const results = [];

  for (const filePath of inputPaths) {
    try {
      const stats = await fs.stat(filePath);
      const name = path.basename(filePath);
      results.push({
        id: `${name}-${stats.size}-${stats.mtimeMs}`,
        path: filePath,
        name,
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
  const defaultOutput = path.join(process.cwd(), sanitizeOutputName(payload?.outputName || 'merged-output'));
  return computeMergePlan(files, defaultOutput);
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

ipcMain.handle('merge-pdfs', async (_, payload) => {
  const { files, outputName } = payload;
  if (!Array.isArray(files) || files.length === 0) {
    throw new Error('No files selected.');
  }
  if (activeMerge) {
    throw new Error('A merge is already in progress.');
  }

  const save = await dialog.showSaveDialog(mainWindow, {
    title: 'Save merged PDF',
    defaultPath: sanitizeOutputName(outputName),
    filters: [{ name: 'PDF Files', extensions: ['pdf'] }],
  });

  if (save.canceled || !save.filePath) {
    return { canceled: true };
  }

  const state = {
    cancelRequested: false,
    completedFiles: 0,
    currentFileName: 'Preparing merge engine...',
    activeChild: null,
  };

  activeMerge = state;

  const mergeJobId = `${Date.now()}-${Math.random().toString(16).slice(2, 8)}`;
  const tempDir = path.join(TEMP_ROOT, mergeJobId);
  let currentMergedPath = null;
  let mergeStep = 0;
  let pendingSmallFiles = [];

  const flushPending = async (qpdfPath) => {
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

        state.currentFileName = `Merging files (${state.completedFiles}/${files.length})`;
        emitStatus(state, files.length);

        await qpdfMergeFiles(qpdfPath, mergeInputs, outPath, state);

        if (currentMergedPath && currentMergedPath.startsWith(tempDir)) {
          await fs.rm(currentMergedPath, { force: true }).catch(() => {});
        }

        currentMergedPath = outPath;
      }

      for (const item of chunk) {
        state.completedFiles += 1;
      }
      emitStatus(state, files.length);
    }

    pendingSmallFiles = [];
  };

  const splitLargeFileToTemp = async (qpdfPath, file, fileIndex) => {
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
      emitStatus(state, files.length);

      const splitPath = path.join(
        tempDir,
        `split-${String(fileIndex + 1).padStart(4, '0')}-${String(seg + 1).padStart(4, '0')}.pdf`
      );
      await qpdfMakeRangePdf(qpdfPath, file.path, `${start}-${end}`, splitPath, state);
      splitPaths.push(splitPath);
    }

    return splitPaths;
  };

  try {
    const qpdfPath = await ensureQpdf();
    const plan = computeMergePlan(files, save.filePath);
    emitStatus(state, files.length);

    if (plan.mode === 'fast') {
      state.currentFileName = 'Fast mode: merging directly';
      emitStatus(state, files.length);

      try {
        await qpdfMergeFiles(qpdfPath, files.map((f) => f.path), save.filePath, state);
        state.completedFiles = files.length;
        sendMergeProgress({
          completed: files.length,
          total: files.length,
          percent: 100,
          etaSeconds: 0,
          fileName: 'Finished',
          phase: 'done',
          done: true,
        });
        return { canceled: false, outputPath: save.filePath };
      } catch (err) {
        if ((err && err.code === 'MERGE_CANCELED') || state.cancelRequested) {
          throw err;
        }

        // Fallback if direct merge fails even though plan predicted fast mode.
        await fs.rm(save.filePath, { force: true }).catch(() => {});
        state.currentFileName = 'Switching to safe mode (slower)';
        emitStatus(state, files.length);
      }
    }

    await fs.mkdir(tempDir, { recursive: true });

    for (let i = 0; i < files.length; i += 1) {
      assertNotCanceled(state);
      const file = files[i];
      const fileSize = Math.max(1, Number(file.size || 1));

      if (fileSize < SPLIT_TRIGGER_BYTES) {
        state.currentFileName = `Queued: ${file.name}`;
        emitStatus(state, files.length);

        pendingSmallFiles.push({ path: file.path, name: file.name });
        if (pendingSmallFiles.length >= MAX_INPUTS_PER_MERGE_CALL) {
          await flushPending(qpdfPath);
        }
        continue;
      }

      // Merge all prior small files first, then process this large file only.
      await flushPending(qpdfPath);

      const splitPaths = await splitLargeFileToTemp(qpdfPath, file, i);

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
          emitStatus(state, files.length);

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
      emitStatus(state, files.length);
    }

    await flushPending(qpdfPath);

    assertNotCanceled(state);

    if (!currentMergedPath) {
      throw new Error('No merged output was produced.');
    }

    sendMergeProgress({
      completed: files.length,
      total: files.length,
      percent: 99,
      etaSeconds: 0,
      fileName: 'Writing file to disk',
      phase: 'writing',
    });

    await fs.copyFile(currentMergedPath, save.filePath);

    sendMergeProgress({
      completed: files.length,
      total: files.length,
      percent: 100,
      etaSeconds: 0,
      fileName: 'Finished',
      phase: 'done',
      done: true,
    });

    return { canceled: false, outputPath: save.filePath };
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
