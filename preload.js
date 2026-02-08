const { contextBridge, ipcRenderer, webUtils } = require('electron');

contextBridge.exposeInMainWorld('pdfMergerAPI', {
  pickPdfFiles: () => ipcRenderer.invoke('pick-pdf-files'),
  readFileMetadata: (paths) => ipcRenderer.invoke('read-file-metadata', paths),
  getMergePlan: (payload) => ipcRenderer.invoke('get-merge-plan', payload),
  mergePdfs: (payload) => ipcRenderer.invoke('merge-pdfs', payload),
  cancelMerge: () => ipcRenderer.invoke('cancel-merge'),
  windowMinimize: () => ipcRenderer.invoke('window-minimize'),
  windowToggleMaximize: () => ipcRenderer.invoke('window-toggle-maximize'),
  windowClose: () => ipcRenderer.invoke('window-close'),
  windowIsMaximized: () => ipcRenderer.invoke('window-is-maximized'),
  getPathForFile: (file) => webUtils.getPathForFile(file),
  onWindowState: (handler) => {
    const listener = (_, state) => handler(state);
    ipcRenderer.on('window-state', listener);
    return () => ipcRenderer.removeListener('window-state', listener);
  },
  onMergeProgress: (handler) => {
    const listener = (_, data) => handler(data);
    ipcRenderer.on('merge-progress', listener);
    return () => ipcRenderer.removeListener('merge-progress', listener);
  },
});
