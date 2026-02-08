const clearBtn = document.getElementById('clearBtn');
const mergeBtn = document.getElementById('mergeBtn');
const fileListEl = document.getElementById('fileList');
const emptyStateEl = document.getElementById('emptyState');
const dropZone = document.getElementById('dropZone');
const sortModeEl = document.getElementById('sortMode');
const sortDirectionEl = document.getElementById('sortDirection');
const outputNameEl = document.getElementById('outputName');
const mergeModeHintEl = document.getElementById('mergeModeHint');
const statusTextEl = document.getElementById('statusText');
const progressFillEl = document.getElementById('progressFill');
const stopMergeBtn = document.getElementById('stopMergeBtn');
const winMinBtn = document.getElementById('winMinBtn');
const winMaxBtn = document.getElementById('winMaxBtn');
const winCloseBtn = document.getElementById('winCloseBtn');

let files = [];
let unsubscribeProgress;
let dragIndex = null;
let hoverMarkerKey = '';
let mergeInProgress = false;
let cancelPending = false;
let mergePlanRequestSeq = 0;
let unsubscribeWindowState;

function formatDateTime(ms) {
  const date = new Date(ms);
  return date.toLocaleString(undefined, {
    year: 'numeric',
    month: '2-digit',
    day: '2-digit',
    hour: '2-digit',
    minute: '2-digit',
    second: '2-digit',
    hour12: false,
  });
}

function formatSize(bytes) {
  if (bytes < 1024) return `${bytes} B`;
  const kb = bytes / 1024;
  if (kb < 1024) return `${kb.toFixed(1)} KB`;
  const mb = kb / 1024;
  if (mb < 1024) return `${mb.toFixed(2)} MB`;
  const gb = mb / 1024;
  if (gb < 1024) return `${gb.toFixed(2)} GB`;
  return `${(gb / 1024).toFixed(2)} TB`;
}

function parseNumericBase(name) {
  const match = name.match(/\d+(?:\.\d+)?/);
  return match ? Number(match[0]) : Number.POSITIVE_INFINITY;
}

function dedupeAndAppend(newItems) {
  const seen = new Set(files.map((f) => f.path.toLowerCase()));
  const merged = [...files];

  for (const item of newItems) {
    const key = item.path.toLowerCase();
    if (!seen.has(key)) {
      seen.add(key);
      merged.push(item);
    }
  }

  files = merged;
  applySort();
  refreshMergePlan();
}

function normalizeIncomingPath(p) {
  const raw = String(p || '').trim();
  if (!raw) return '';

  if (/^file:\/\//i.test(raw)) {
    try {
      const withoutScheme = raw.replace(/^file:\/\//i, '');
      const decoded = decodeURIComponent(withoutScheme);
      const windowsPath = decoded.replace(/\//g, '\\');
      return windowsPath.replace(/^\\+/, '');
    } catch {
      return raw;
    }
  }

  return raw;
}

async function addFilesFromPaths(paths) {
  if (!paths || paths.length === 0) return;

  const normalized = paths
    .map(normalizeIncomingPath)
    .filter(Boolean)
    .filter((p) => /\.pdf$/i.test(p));

  if (!normalized.length) return;

  const metadata = await window.pdfMergerAPI.readFileMetadata(normalized);
  dedupeAndAppend(metadata);
}

function extractDroppedPaths(dataTransfer) {
  const normalizeUnique = (values) => {
    const out = [];
    const seen = new Set();
    for (const value of values) {
      const normalized = normalizeIncomingPath(value);
      if (!normalized) continue;
      const key = normalized.toLowerCase();
      if (seen.has(key)) continue;
      seen.add(key);
      out.push(normalized);
    }
    return out;
  };

  const fromFiles = normalizeUnique(
    Array.from(dataTransfer?.files || []).map((file) => {
      let filePath = file?.path || '';
      if (!filePath && file) {
        try {
          filePath = window.pdfMergerAPI.getPathForFile(file) || '';
        } catch {
          filePath = '';
        }
      }
      return filePath;
    })
  );

  const fromItems = normalizeUnique(
    Array.from(dataTransfer?.items || [])
      .filter((item) => item.kind === 'file')
      .map((item) => item.getAsFile?.())
      .filter(Boolean)
      .map((f) => {
        let itemPath = f.path || '';
        if (!itemPath) {
          try {
            itemPath = window.pdfMergerAPI.getPathForFile(f) || '';
          } catch {
            itemPath = '';
          }
        }
        return itemPath;
      })
  );

  const fromUriList = normalizeUnique(
    (dataTransfer?.getData?.('text/uri-list') || '')
      .split(/\r?\n/)
      .map((line) => line.trim())
      .filter((line) => line && !line.startsWith('#'))
  );

  const fromPlain = normalizeUnique(
    (dataTransfer?.getData?.('text/plain') || '')
      .split(/\r?\n/)
      .map((line) => line.trim())
      .filter((line) => /^file:\/\//i.test(line) || /^[a-zA-Z]:\\/.test(line))
  );

  const candidates = [
    { paths: fromUriList, priority: 4 },
    { paths: fromFiles, priority: 3 },
    { paths: fromItems, priority: 2 },
    { paths: fromPlain, priority: 1 },
  ].filter((candidate) => candidate.paths.length > 0);

  if (!candidates.length) return [];

  candidates.sort((a, b) => {
    if (b.paths.length !== a.paths.length) return b.paths.length - a.paths.length;
    return b.priority - a.priority;
  });

  return candidates[0].paths;
}

function render() {
  fileListEl.innerHTML = '';
  emptyStateEl.style.display = files.length ? 'none' : 'block';

  files.forEach((file, idx) => {
    const li = document.createElement('li');
    li.className = 'file-item';
    li.draggable = true;
    li.dataset.index = String(idx);

    li.innerHTML = `
      <div class="drag-handle" title="Drag to reorder">☰</div>
      <div class="file-main">
        <div class="file-name" title="${file.name}">${idx + 1}. ${file.name}</div>
        <div class="file-meta">Created: ${formatDateTime(file.createdMs)} | Size: ${formatSize(file.size)}</div>
      </div>
      <button class="remove-btn" data-remove="${idx}">Remove</button>
    `;

    li.addEventListener('dragstart', () => {
      dragIndex = idx;
      li.classList.add('dragging');
      hoverMarkerKey = '';
    });

    li.addEventListener('dragend', () => {
      dragIndex = null;
      li.classList.remove('dragging');
      clearDropMarkers();
    });

    li.addEventListener('dragover', (e) => {
      e.preventDefault();
      if (dragIndex == null) return;

      const rect = li.getBoundingClientRect();
      const placeBefore = e.clientY < rect.top + rect.height / 2;
      const markerClass = placeBefore ? 'drop-before' : 'drop-after';
      const markerKey = `${idx}-${markerClass}`;

      if (hoverMarkerKey !== markerKey) {
        clearDropMarkers();
        li.classList.add(markerClass);
        hoverMarkerKey = markerKey;
      }

      let targetIndex = placeBefore ? idx : idx + 1;
      if (targetIndex > dragIndex) targetIndex -= 1;
      targetIndex = Math.max(0, Math.min(files.length - 1, targetIndex));

      if (targetIndex !== dragIndex) {
        reorderManual(dragIndex, targetIndex);
        dragIndex = targetIndex;
      }
    });

    li.addEventListener('drop', (e) => {
      e.preventDefault();
      clearDropMarkers();
    });

    fileListEl.appendChild(li);
  });
}

function reorderManual(fromIndex, toIndex) {
  if (fromIndex == null || toIndex == null || fromIndex === toIndex) return;
  const clone = [...files];
  const [moved] = clone.splice(fromIndex, 1);
  clone.splice(toIndex, 0, moved);
  files = clone;
  sortModeEl.value = 'manual';
  render();
}

function clearDropMarkers() {
  hoverMarkerKey = '';
  for (const el of fileListEl.querySelectorAll('.file-item.drop-before, .file-item.drop-after')) {
    el.classList.remove('drop-before', 'drop-after');
  }
}

function applySort() {
  const mode = sortModeEl.value;
  const direction = sortDirectionEl.value === 'asc' ? 1 : -1;

  if (mode === 'manual') {
    render();
    return;
  }

  files = [...files].sort((a, b) => {
    if (mode === 'alpha') {
      return direction * a.name.localeCompare(b.name, undefined, { numeric: false, sensitivity: 'base' });
    }

    if (mode === 'numeric') {
      const numA = parseNumericBase(a.name);
      const numB = parseNumericBase(b.name);
      if (numA === numB) return direction * a.name.localeCompare(b.name, undefined, { sensitivity: 'base' });
      return direction * (numA - numB);
    }

    if (mode === 'created') {
      return direction * (a.createdMs - b.createdMs);
    }

    return 0;
  });

  render();
}

function setStatus(text) {
  statusTextEl.textContent = text;
}

function setProgress(percent) {
  progressFillEl.style.width = `${Math.max(0, Math.min(100, percent))}%`;
}

function setMergeUiState(active) {
  mergeInProgress = active;
  stopMergeBtn.hidden = !active;
  mergeBtn.disabled = active;
  mergeBtn.textContent = active ? 'Merging...' : 'Merge PDFs';
}

function setMergeModeHint(text) {
  const value = String(text || '').trim();
  if (!value) {
    mergeModeHintEl.hidden = true;
    mergeModeHintEl.textContent = '';
    return;
  }
  mergeModeHintEl.hidden = false;
  mergeModeHintEl.textContent = value;
}

async function refreshMergePlan() {
  if (!files.length) {
    setMergeModeHint('');
    return;
  }

  const requestId = ++mergePlanRequestSeq;

  try {
    const plan = await window.pdfMergerAPI.getMergePlan({
      files,
      outputName: outputNameEl.value,
    });

    if (requestId !== mergePlanRequestSeq) return;

    if (plan?.mode !== 'slow') {
      setMergeModeHint('');
      return;
    }

    if (plan.reason === 'memory') {
      setMergeModeHint(
        `Safe mode: total size ${formatSize(plan.totalBytes)} exceeds fast memory budget ${formatSize(plan.memoryLimitBytes)}. Merge will be slower.`
      );
      return;
    }

    if (plan.reason === 'command_length') {
      setMergeModeHint('Safe mode: file path list is very long, so merge will run in slower compatibility mode.');
      return;
    }

    setMergeModeHint('Safe mode: merge will run in slower compatibility mode.');
  } catch {
    if (requestId !== mergePlanRequestSeq) return;
    setMergeModeHint('');
  }
}

function resetAfterCancelImmediate() {
  setProgress(0);
  setStatus('Merge canceled.');
  setMergeUiState(false);
}

function applyWindowState(state) {
  if (!winMaxBtn) return;
  const maximized = Boolean(state && state.maximized);
  winMaxBtn.classList.toggle('maximized', maximized);
  winMaxBtn.setAttribute('aria-label', maximized ? 'Restore' : 'Maximize');
}

async function openFilePicker() {
  const paths = await window.pdfMergerAPI.pickPdfFiles();
  await addFilesFromPaths(paths);
}

async function merge() {
  if (cancelPending) {
    setStatus('Cancel in progress. Please wait a moment.');
    return;
  }

  if (!files.length) {
    setStatus('No files to merge. Add PDF files first.');
    return;
  }

  cancelPending = false;
  setMergeUiState(true);
  setProgress(0);
  setStatus('Preparing merge...');

  try {
    const result = await window.pdfMergerAPI.mergePdfs({
      files,
      outputName: outputNameEl.value,
    });

    if (result.canceled) {
      cancelPending = false;
      setStatus('Merge canceled.');
      setProgress(0);
    } else {
      setStatus(`Done. Saved to: ${result.outputPath}`);
      setProgress(100);
    }
  } catch (err) {
    if (!cancelPending) {
      setStatus(`Merge failed: ${err.message || err}`);
    }
  } finally {
    cancelPending = false;
    setMergeUiState(false);
  }
}

clearBtn.addEventListener('click', () => {
  files = [];
  render();
  setProgress(0);
  setStatus('Ready');
  refreshMergePlan();
});
mergeBtn.addEventListener('click', merge);
stopMergeBtn.addEventListener('click', async () => {
  if (!mergeInProgress) return;
  cancelPending = true;
  resetAfterCancelImmediate();
  try {
    await window.pdfMergerAPI.cancelMerge();
  } catch {
    // Ignore cancellation transport errors; UI has already been reset.
  }
});

if (winMinBtn) {
  winMinBtn.addEventListener('click', async () => {
    await window.pdfMergerAPI.windowMinimize();
  });
}

if (winMaxBtn) {
  winMaxBtn.addEventListener('click', async () => {
    const state = await window.pdfMergerAPI.windowToggleMaximize();
    applyWindowState(state);
  });
}

if (winCloseBtn) {
  winCloseBtn.addEventListener('click', async () => {
    await window.pdfMergerAPI.windowClose();
  });
}
sortModeEl.addEventListener('change', applySort);
sortDirectionEl.addEventListener('change', applySort);
outputNameEl.addEventListener('input', refreshMergePlan);

dropZone.addEventListener('click', openFilePicker);

fileListEl.addEventListener('click', (e) => {
  const target = e.target;
  if (!(target instanceof HTMLElement)) return;

  const removeIndex = target.dataset.remove;
  if (removeIndex != null) {
    files.splice(Number(removeIndex), 1);
    render();
    refreshMergePlan();
  }
});

dropZone.addEventListener('dragover', (e) => {
  e.preventDefault();
  dropZone.classList.add('drag-active');
});

dropZone.addEventListener('dragleave', () => {
  dropZone.classList.remove('drag-active');
});

dropZone.addEventListener('drop', async (e) => {
  e.preventDefault();
  dropZone.classList.remove('drag-active');
  const droppedPaths = extractDroppedPaths(e.dataTransfer);
  await addFilesFromPaths(droppedPaths);
});

window.addEventListener('dragover', (e) => {
  e.preventDefault();
});

window.addEventListener('drop', (e) => {
  if (!dropZone.contains(e.target)) {
    e.preventDefault();
  }
});

unsubscribeProgress = window.pdfMergerAPI.onMergeProgress((update) => {
  if (cancelPending && update.phase !== 'canceled' && update.phase !== 'done') {
    return;
  }

  const pct = Number(update.percent || 0);
  setProgress(pct);

  if (update.done) {
    setStatus('Merge completed.');
    setMergeUiState(false);
    return;
  }

  if (update.phase === 'canceled') {
    cancelPending = false;
    setStatus('Merge canceled.');
    setProgress(0);
    setMergeUiState(false);
    return;
  }

  if (update.phase === 'finalizing') {
    setStatus('Finalizing merged PDF...');
    return;
  }

  if (update.phase === 'writing') {
    setStatus('Writing merged PDF to disk...');
    return;
  }

  setStatus(
    `Merging ${update.completed}/${update.total}: ${update.fileName}`
  );
});

unsubscribeWindowState = window.pdfMergerAPI.onWindowState((state) => {
  applyWindowState(state);
});

window.pdfMergerAPI.windowIsMaximized().then((state) => {
  applyWindowState(state);
}).catch(() => {});

window.addEventListener('beforeunload', () => {
  if (typeof unsubscribeProgress === 'function') {
    unsubscribeProgress();
  }
  if (typeof unsubscribeWindowState === 'function') {
    unsubscribeWindowState();
  }
});

render();
setMergeUiState(false);
refreshMergePlan();
