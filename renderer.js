const clearBtn = document.getElementById('clearBtn');
const mergeBtn = document.getElementById('mergeBtn');
const fileListEl = document.getElementById('fileList');
const emptyStateEl = document.getElementById('emptyState');
const dropZone = document.getElementById('dropZone');
const sortModeEl = document.getElementById('sortMode');
const sortDirectionEl = document.getElementById('sortDirection');
const outputFormatEl = document.getElementById('outputFormat');
const outputNameEl = document.getElementById('outputName');
const mergeModeHintEl = document.getElementById('mergeModeHint');
const statusTextEl = document.getElementById('statusText');
const progressFillEl = document.getElementById('progressFill');
const stopMergeBtn = document.getElementById('stopMergeBtn');
const winMinBtn = document.getElementById('winMinBtn');
const winMaxBtn = document.getElementById('winMaxBtn');
const winCloseBtn = document.getElementById('winCloseBtn');
const imagePreviewModalEl = document.getElementById('imagePreviewModal');
const imagePreviewImgEl = document.getElementById('imagePreviewImg');
const imagePreviewCaptionEl = document.getElementById('imagePreviewCaption');
const imagePreviewCloseBtn = document.getElementById('imagePreviewCloseBtn');
const SUPPORTED_FILE_PATTERN = /\.(pdf|pptx|ppt|png|jpe?g|bmp|gif|webp|tiff?)$/i;

let files = [];
let unsubscribeProgress;
let dragIndex = null;
let hoverMarkerKey = '';
let mergeInProgress = false;
let cancelPending = false;
let mergePlanRequestSeq = 0;
let unsubscribeWindowState;
let replaceBatchOnNextAdd = false;

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
  updateMergeButtonLabel();
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

function isSupportedInputPath(inputPath) {
  return SUPPORTED_FILE_PATTERN.test(String(inputPath || ''));
}

async function addFilesFromPaths(paths) {
  if (!paths || paths.length === 0) return;

  const normalized = paths
    .map(normalizeIncomingPath)
    .filter(Boolean)
    .filter(isSupportedInputPath);

  if (!normalized.length) return;

  if (replaceBatchOnNextAdd) {
    files = [];
    replaceBatchOnNextAdd = false;
    setProgress(0);
    setStatus('Ready');
    setMergeModeHint('');
  }

  const metadata = await window.fileMergerAPI.readFileMetadata(normalized);
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
          filePath = window.fileMergerAPI.getPathForFile(file) || '';
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
            itemPath = window.fileMergerAPI.getPathForFile(f) || '';
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

  const sources = [
    { name: 'items', priority: 1, paths: fromItems },
    { name: 'files', priority: 2, paths: fromFiles },
    { name: 'uri', priority: 3, paths: fromUriList },
    { name: 'plain', priority: 4, paths: fromPlain },
  ].filter((source) => source.paths.length > 0);

  if (!sources.length) return [];
  if (sources.length === 1) return sources[0].paths;

  const sameMembers = (a, b) => {
    if (a.length !== b.length) return false;
    const setA = new Set(a.map((p) => p.toLowerCase()));
    for (const p of b) {
      if (!setA.has(String(p).toLowerCase())) return false;
    }
    return true;
  };

  const pairwiseDisagreement = (a, b) => {
    const indexB = new Map();
    b.forEach((value, idx) => {
      indexB.set(String(value).toLowerCase(), idx);
    });

    let inversions = 0;
    for (let i = 0; i < a.length; i += 1) {
      const ai = String(a[i]).toLowerCase();
      const idxI = indexB.get(ai);
      if (idxI == null) continue;
      for (let j = i + 1; j < a.length; j += 1) {
        const aj = String(a[j]).toLowerCase();
        const idxJ = indexB.get(aj);
        if (idxJ == null) continue;
        if (idxI > idxJ) inversions += 1;
      }
    }
    return inversions;
  };

  const scored = sources.map((source) => {
    let comparisons = 0;
    let inversions = 0;
    for (const other of sources) {
      if (other === source) continue;
      if (!sameMembers(source.paths, other.paths)) continue;
      comparisons += 1;
      inversions += pairwiseDisagreement(source.paths, other.paths);
    }
    return { source, comparisons, inversions };
  });

  scored.sort((a, b) => {
    if (b.comparisons !== a.comparisons) return b.comparisons - a.comparisons;
    if (a.inversions !== b.inversions) return a.inversions - b.inversions;
    return a.source.priority - b.source.priority;
  });

  return scored[0].source.paths;
}

function escapeHtml(value) {
  return String(value || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function toFileUrl(filePath) {
  const normalized = String(filePath || '').replace(/\\/g, '/').replace(/^\/+/, '');
  return `file:///${encodeURI(normalized)}`;
}

function render() {
  fileListEl.innerHTML = '';
  emptyStateEl.style.display = files.length ? 'none' : 'block';

  files.forEach((file, idx) => {
    const li = document.createElement('li');
    li.className = 'file-item';
    li.draggable = true;
    li.dataset.index = String(idx);

    const fileName = String(file?.name || '');
    const escapedName = escapeHtml(fileName);
    const kind = getFileKind(file);
    const previewHtml = kind === 'image'
      ? `
        <button
          type="button"
          class="preview-btn"
          data-preview-index="${idx}"
          title="Preview image"
          aria-label="Preview ${escapedName}"
        >
          <img class="file-thumb" alt="" src="${toFileUrl(file.path)}" loading="lazy" />
        </button>
      `
      : '';

    li.innerHTML = `
      <div class="drag-handle" title="Drag to reorder">☰</div>
      <div class="file-main">
        <div class="file-name" title="${escapedName}">
          <span class="file-index">${idx + 1}.</span>
          <span class="file-name-text">${escapedName}</span>
        </div>
        <div class="file-meta">Modified: ${formatDateTime(file.modifiedMs || file.createdMs)} | Size: ${formatSize(file.size)}</div>
      </div>
      <div class="file-actions">
        ${previewHtml}
        <button class="remove-btn" data-remove="${idx}">Remove</button>
      </div>
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

  updateOutputFormatControl();
  updateMergeButtonLabel();
}

function openImagePreview(file) {
  if (!imagePreviewModalEl || !imagePreviewImgEl || !file) return;
  imagePreviewImgEl.src = toFileUrl(file.path);
  imagePreviewImgEl.alt = String(file.name || 'Image preview');
  if (imagePreviewCaptionEl) {
    imagePreviewCaptionEl.textContent = String(file.name || '');
  }
  imagePreviewModalEl.hidden = false;
}

function closeImagePreview() {
  if (!imagePreviewModalEl || !imagePreviewImgEl) return;
  imagePreviewModalEl.hidden = true;
  imagePreviewImgEl.src = '';
  imagePreviewImgEl.alt = '';
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
      return direction * (a.modifiedMs - b.modifiedMs);
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

function getFileKind(file) {
  const explicitKind = String(file?.kind || '').toLowerCase();
  if (explicitKind) return explicitKind;

  const candidatePath = String(file?.path || file?.name || '').toLowerCase();
  if (/\.(pdf)$/.test(candidatePath)) return 'pdf';
  if (/\.(pptx)$/.test(candidatePath)) return 'pptx';
  if (/\.(ppt)$/.test(candidatePath)) return 'ppt';
  if (/\.(png|jpe?g|bmp|gif|webp|tiff?)$/.test(candidatePath)) return 'image';
  return '';
}

function getFileTypeGroup() {
  if (!files.length) return 'empty';

  const allPdfs = files.every((f) => getFileKind(f) === 'pdf');
  if (allPdfs) return 'pdfs';

  const hasPdf = files.some((f) => getFileKind(f) === 'pdf');
  if (hasPdf) return 'mixed-pdf';

  const imageKinds = new Set(['image']);
  const pptKinds = new Set(['ppt', 'pptx']);

  const allImages = files.every((f) => imageKinds.has(getFileKind(f)));
  if (allImages) return 'images';

  const allPowerPoints = files.every((f) => pptKinds.has(getFileKind(f)));
  if (allPowerPoints) return 'powerpoints';

  return 'mixed';
}

function getIdleMergeLabel() {
  const group = getFileTypeGroup();
  if (group === 'pdfs') return 'Merge PDFs';
  if (group === 'mixed-pdf') return 'Merge to PDF';
  if (group === 'images') return 'Merge Photos';
  if (group === 'powerpoints') return 'Merge PowerPoints';
  return 'Merge Files';
}

function updateOutputFormatControl() {
  if (!outputFormatEl) return;
  const group = getFileTypeGroup();
  const forcePdf = group === 'pdfs' || group === 'mixed-pdf';
  outputFormatEl.disabled = forcePdf;
  if (forcePdf) {
    outputFormatEl.value = 'pdf';
  }
}

function updateMergeButtonLabel() {
  if (mergeInProgress) {
    mergeBtn.textContent = 'Merging...';
    mergeBtn.disabled = true;
    return;
  }

  mergeBtn.disabled = false;
  mergeBtn.textContent = getIdleMergeLabel();
}

function setMergeUiState(active) {
  mergeInProgress = active;
  stopMergeBtn.hidden = !active;
  mergeBtn.disabled = active;
  updateMergeButtonLabel();
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
    const plan = await window.fileMergerAPI.getMergePlan({
      files,
      outputName: outputNameEl.value,
      outputFormat: outputFormatEl?.value,
    });

    if (requestId !== mergePlanRequestSeq) return;

    if (plan?.mode === 'unsupported') {
      if (plan.reason === 'mixed_requires_pdf') {
        setMergeModeHint('Mixed PDF + PowerPoint/image batches can only be exported as PDF.');
        return;
      }

      const unsupportedCount = Number(plan.unsupported || 0);
      setMergeModeHint(`Unsupported files detected (${unsupportedCount}). Remove them to continue.`);
      return;
    }

    if (plan?.route === 'mixed') {
      setMergeModeHint('Mixed batch will be converted and merged as PDF.');
      return;
    }

    if (plan?.route === 'pdf' && plan?.mode === 'slow') {
      if (plan.reason === 'memory') {
        setMergeModeHint(
          `PDF safe mode: total size ${formatSize(plan.totalBytes)} exceeds fast memory budget ${formatSize(plan.memoryLimitBytes)}.`
        );
        return;
      }

      if (plan.reason === 'command_length') {
        setMergeModeHint('PDF safe mode: input list is too long for direct command mode.');
        return;
      }

      setMergeModeHint('PDF safe mode: merge will run in slower compatibility mode.');
      return;
    }

    setMergeModeHint('');
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
  const paths = await window.fileMergerAPI.pickFiles();
  await addFilesFromPaths(paths);
}

async function merge() {
  if (cancelPending) {
    setStatus('Cancel in progress. Please wait a moment.');
    return;
  }

  if (!files.length) {
    setStatus('No files to merge. Add PDFs, PowerPoints, or images first.');
    return;
  }

  cancelPending = false;
  setMergeUiState(true);
  setProgress(0);
  setStatus('Preparing merge...');

  try {
    const result = await window.fileMergerAPI.mergeFiles({
      files,
      outputName: outputNameEl.value,
      outputFormat: outputFormatEl?.value || 'pptx',
    });

    if (result.canceled) {
      cancelPending = false;
      setStatus('Merge canceled.');
      setProgress(0);
    } else {
      replaceBatchOnNextAdd = true;
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
  replaceBatchOnNextAdd = false;
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
    await window.fileMergerAPI.cancelMerge();
  } catch {
    // Ignore cancellation transport errors; UI has already been reset.
  }
});

if (winMinBtn) {
  winMinBtn.addEventListener('click', async () => {
    await window.fileMergerAPI.windowMinimize();
  });
}

if (winMaxBtn) {
  winMaxBtn.addEventListener('click', async () => {
    const state = await window.fileMergerAPI.windowToggleMaximize();
    applyWindowState(state);
  });
}

if (winCloseBtn) {
  winCloseBtn.addEventListener('click', async () => {
    await window.fileMergerAPI.windowClose();
  });
}
sortModeEl.addEventListener('change', applySort);
sortDirectionEl.addEventListener('change', applySort);
outputNameEl.addEventListener('input', refreshMergePlan);
if (outputFormatEl) {
  outputFormatEl.addEventListener('change', refreshMergePlan);
}

if (imagePreviewCloseBtn) {
  imagePreviewCloseBtn.addEventListener('click', closeImagePreview);
}

if (imagePreviewModalEl) {
  imagePreviewModalEl.addEventListener('click', (e) => {
    const target = e.target;
    if (!(target instanceof HTMLElement)) return;
    if (target.dataset.closePreview === '1') {
      closeImagePreview();
    }
  });
}

window.addEventListener('keydown', (e) => {
  if (e.key === 'Escape' && imagePreviewModalEl && !imagePreviewModalEl.hidden) {
    closeImagePreview();
  }
});

dropZone.addEventListener('click', openFilePicker);

fileListEl.addEventListener('click', (e) => {
  const target = e.target;
  if (!(target instanceof HTMLElement)) return;

  const previewBtn = target.closest('[data-preview-index]');
  if (previewBtn instanceof HTMLElement && previewBtn.dataset.previewIndex != null) {
    const idx = Number(previewBtn.dataset.previewIndex);
    const file = files[idx];
    if (file && getFileKind(file) === 'image') {
      openImagePreview(file);
    }
    return;
  }

  const removeBtn = target.closest('[data-remove]');
  const removeIndex = removeBtn instanceof HTMLElement ? removeBtn.dataset.remove : null;
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

unsubscribeProgress = window.fileMergerAPI.onMergeProgress((update) => {
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
    setStatus('Finalizing merged file...');
    return;
  }

  if (update.phase === 'writing') {
    setStatus('Writing merged file to disk...');
    return;
  }

  setStatus(
    `Merging ${update.completed}/${update.total}: ${update.fileName}`
  );
});

unsubscribeWindowState = window.fileMergerAPI.onWindowState((state) => {
  applyWindowState(state);
});

window.fileMergerAPI.windowIsMaximized().then((state) => {
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
