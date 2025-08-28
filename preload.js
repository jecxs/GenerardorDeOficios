const { contextBridge, ipcRenderer } = require("electron");

contextBridge.exposeInMainWorld("electronAPI", {
    // Funciones de selección de archivos
    selectBaseFolder: () => ipcRenderer.invoke("select-base-folder"),
    loadSavedFolder: () => ipcRenderer.invoke("load-saved-folder"), // 🔥 NUEVA función
    selectExcel: () => ipcRenderer.invoke("select-excel"),
    selectDocx: () => ipcRenderer.invoke("select-docx"),
    selectDocxMultiple: () => ipcRenderer.invoke("select-docx-multiple"),

    sendEmails: (args) => ipcRenderer.invoke("send-emails", args),

    saveSmtpConfig: (config) => ipcRenderer.invoke("save-smtp-config", config),

    selectExtraPdfs: () => ipcRenderer.invoke("select-extra-pdfs"),

    // Lectura de Excel
    readExcel: (filePath) => ipcRenderer.invoke("read-excel", filePath),

    // Generación de PDFs con LibreOffice
    generatePdfsLibreOffice: (args) => ipcRenderer.invoke("generate-pdfs-libreoffice", args),

    // Verificar si LibreOffice está disponible
    checkLibreOffice: () => ipcRenderer.invoke("check-libreoffice"),

    // Seleccionar LibreOffice manualmente
    selectLibreOffice: () => ipcRenderer.invoke("select-libreoffice"),


    loadConfig: () => ipcRenderer.invoke("load-config"),
    saveConfig: (config) => ipcRenderer.invoke("save-config", config),

    selectSingleDocx: () => ipcRenderer.invoke("select-single-docx"),
    // Escuchar progreso
    onPdfProgress: (callback) => {
        ipcRenderer.on('pdf-progress', (event, data) => callback(data));
    },

    generateAndSendIntegrated: (args) => ipcRenderer.invoke("generate-and-send-integrated", args),
    onIntegratedProgress: (callback) => {
        ipcRenderer.on('integrated-progress', (event, data) => callback(data));
    },


    // Limpiar listeners
    removeAllListeners: () => {
        ipcRenderer.removeAllListeners('pdf-progress');
        ipcRenderer.removeAllListeners('integrated-progress');
    }
});