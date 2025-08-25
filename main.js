const { app, BrowserWindow, dialog, ipcMain } = require("electron");

const fs = require("fs");
const path = require("path");
const XLSX = require("xlsx");
const Docxtemplater = require("docxtemplater");
const PizZip = require("pizzip");
const mammoth = require("mammoth");
const libre = require('libreoffice-convert');
const { exec } = require('child_process');
const os = require('os');
const nodemailer = require("nodemailer");
const configPath = path.join(app.getPath('userData'), 'config.json');


let smtpConfig = null;

let mainWindow;

function createWindow() {
    mainWindow = new BrowserWindow({
        width: 1000,
        height: 700,
        webPreferences: {
            preload: path.join(__dirname, "preload.js"),
        },
    });

    mainWindow.loadFile("renderer/index.html");

    // 🔥 CARGAR CONFIGURACIÓN AL INICIO
    mainWindow.webContents.once('dom-ready', () => {
        const config = loadConfig();
        if (config.smtpConfig) {
            smtpConfig = config.smtpConfig;
            console.log("📨 Configuración SMTP cargada desde archivo");
        }
    });
}



// 🔹 Función para crear subcarpetas si no existen
function ensureSubfolders(basePath) {
    const subfolders = ["excel", "plantillas", "adjuntos", "salida"];
    let created = [];

    subfolders.forEach((folder) => {
        const dir = path.join(basePath, folder);
        if (!fs.existsSync(dir)) {
            fs.mkdirSync(dir);
            created.push(folder);
        }
    });

    return created;
}

function loadConfig() {
    try {
        if (fs.existsSync(configPath)) {
            const configData = fs.readFileSync(configPath, 'utf8');
            return JSON.parse(configData);
        }
    } catch (error) {
        console.error('Error cargando configuración:', error);
    }
    return {
        baseFolder: null,
        smtpConfig: null
    };
}

function saveConfig(config) {
    try {
        const existingConfig = loadConfig();
        const newConfig = { ...existingConfig, ...config };
        fs.writeFileSync(configPath, JSON.stringify(newConfig, null, 2), 'utf8');
        return { success: true };
    } catch (error) {
        console.error('Error guardando configuración:', error);
        return { success: false, error: error.message };
    }
}

ipcMain.handle("load-config", async () => {
    return loadConfig();
});

// Handler para guardar configuración
ipcMain.handle("save-config", async (event, config) => {
    return saveConfig(config);
});

// 🔹 Función reutilizable para leer Excel
function readExcelFile(filePath) {
    try {
        if (!fs.existsSync(filePath)) {
            throw new Error(`El archivo Excel no existe: ${filePath}`);
        }
        const workbook = XLSX.readFile(filePath);
        const sheetName = workbook.SheetNames[0]; // Primera hoja
        const sheet = workbook.Sheets[sheetName];
        const data = XLSX.utils.sheet_to_json(sheet, { header: 1 }); // Como array de arrays

        // Headers son la primera fila
        const headers = data[0] || [];
        // Filas de datos (excluyendo header)
        let rows = data.slice(1);

        // 🔧 Filtrar filas vacías
        rows = rows.filter(row =>
            row.some(cell => cell !== null && cell !== undefined && String(cell).trim() !== "")
        );

        // Previsualización: primeras 5 filas
        const previewRows = rows.slice(0, 5);

        return { headers, previewRows, rows, totalRows: rows.length };
    } catch (error) {
        console.error("Error leyendo Excel:", error);
        return { error: error.message };
    }
}

// Rutas comunes donde se instala LibreOffice
const getLibreOfficePaths = () => {
    const platform = os.platform();

    if (platform === 'win32') {
        return [
            'C:\\Program Files\\LibreOffice\\program\\soffice.exe',
            'C:\\Program Files (x86)\\LibreOffice\\program\\soffice.exe',
            'C:\\Users\\%USERNAME%\\AppData\\Local\\Programs\\LibreOffice\\program\\soffice.exe',
            process.env.PROGRAMFILES + '\\LibreOffice\\program\\soffice.exe',
            process.env['PROGRAMFILES(X86)'] + '\\LibreOffice\\program\\soffice.exe'
        ];
    } else if (platform === 'darwin') {
        return [
            '/Applications/LibreOffice.app/Contents/MacOS/soffice',
            '/usr/local/bin/soffice',
            '/opt/homebrew/bin/soffice'
        ];
    } else {
        return [
            '/usr/bin/soffice',
            '/usr/local/bin/soffice',
            '/opt/libreoffice/program/soffice',
            '/snap/bin/libreoffice'
        ];
    }
};

// Verificar si un archivo existe y es ejecutable
const checkExecutable = (filePath) => {
    try {
        const stats = fs.statSync(filePath);
        return stats.isFile();
    } catch (error) {
        return false;
    }
};

// Función mejorada para detectar LibreOffice
async function findLibreOffice() {
    return new Promise((resolve) => {
        console.log('🔍 Buscando LibreOffice...');

        // 1. Intentar comandos básicos primero
        const basicCommands = ['soffice --version', 'libreoffice --version'];

        const tryBasicCommand = (index = 0) => {
            if (index >= basicCommands.length) {
                // 2. Si no funcionan los comandos básicos, buscar en rutas específicas
                searchInPaths();
                return;
            }

            const command = basicCommands[index];
            console.log(`📍 Probando comando: ${command}`);

            exec(command, { timeout: 3000 }, (error, stdout) => {
                if (!error && stdout && stdout.toLowerCase().includes('libreoffice')) {
                    console.log(`✅ LibreOffice encontrado via comando: ${stdout.trim()}`);
                    resolve({ found: true, method: 'command', path: command.split(' ')[0] });
                } else {
                    tryBasicCommand(index + 1);
                }
            });
        };

        // 3. Buscar en rutas específicas del sistema
        const searchInPaths = () => {
            const paths = getLibreOfficePaths();
            console.log('🗂️ Buscando en rutas específicas...');

            for (const execPath of paths) {
                if (execPath && checkExecutable(execPath)) {
                    console.log(`✅ LibreOffice encontrado en: ${execPath}`);

                    // Verificar que realmente funcione
                    exec(`"${execPath}" --headless --terminate_after_init`, { timeout: 5000 }, (error, stdout, stderr) => {
                        if (!error) {
                            console.log(`✅ LibreOffice válido en: ${execPath}`);
                            resolve({ found: true, method: 'path', path: execPath });
                        } else {
                            console.log(`⚠️ Ejecutable encontrado pero no funcional: ${execPath}`);
                            console.error(stderr || error.message);
                        }
                    });
                    return; // Salir del loop al encontrar el primero
                }
            }

            console.log('❌ LibreOffice no encontrado automáticamente');
            resolve({ found: false, method: null, path: null });
        };

        tryBasicCommand();
    });
}

// Variable global para almacenar la ruta de LibreOffice
let libreOfficePath = null;


// Guardar config SMTP
ipcMain.handle("save-smtp-config", async (event, config) => {
    smtpConfig = config;

    // También guardar en archivo de configuración
    saveConfig({ smtpConfig: config });

    console.log("📨 Configuración SMTP guardada:", smtpConfig);
    return { success: true };
});


// Función para crear transporter dinámico
function getTransporter() {
    if (!smtpConfig) throw new Error("SMTP no configurado");

    return nodemailer.createTransport({
        host: smtpConfig.host,
        port: smtpConfig.port,
        secure: smtpConfig.secure, // true = 465, false = 587
        auth: {
            user: smtpConfig.user,
            pass: smtpConfig.pass
        }
    });
}

// Handler para leer Excel (para previsualización)
ipcMain.handle("read-excel", async (event, filePath) => {
    return readExcelFile(filePath);
});

// Seleccionar carpeta base
ipcMain.handle("select-base-folder", async (event, useLastPath = false) => {
    if (useLastPath) {
        const config = loadConfig();
        if (config.baseFolder && fs.existsSync(config.baseFolder)) {
            const created = ensureSubfolders(config.baseFolder);
            return { basePath: config.baseFolder, created, fromConfig: true };
        }
        // Si no existe la carpeta guardada, retornar null sin mostrar diálogo
        return null;
    }

    const result = await dialog.showOpenDialog(mainWindow, {
        properties: ["openDirectory"]
    });

    if (result.canceled || result.filePaths.length === 0) return null;

    const basePath = result.filePaths[0];
    const created = ensureSubfolders(basePath);

    // Guardar en configuración
    saveConfig({ baseFolder: basePath });

    return { basePath, created, fromConfig: false };
});
ipcMain.handle("load-saved-folder", async () => {
    const config = loadConfig();
    if (config.baseFolder && fs.existsSync(config.baseFolder)) {
        const created = ensureSubfolders(config.baseFolder);
        return { basePath: config.baseFolder, created, fromConfig: true };
    }
    return null;
});



// Seleccionar archivo Excel
ipcMain.handle("select-excel", async () => {
    const result = await dialog.showOpenDialog(mainWindow, {
        filters: [{ name: "Excel Files", extensions: ["xlsx"] }],
        properties: ["openFile"],
    });

    if (result.canceled || result.filePaths.length === 0) return null;
    return result.filePaths[0];
});

// Seleccionar plantilla DOCX
ipcMain.handle("select-docx", async () => {
    const result = await dialog.showOpenDialog(mainWindow, {
        filters: [{ name: "Word Files", extensions: ["docx"] }],
        properties: ["openFile"],
    });

    if (result.canceled || result.filePaths.length === 0) return null;
    return result.filePaths[0];
});

function getUniqueFilePath(baseDir, fileName) {
    const ext = path.extname(fileName);      // .pdf
    const name = path.basename(fileName, ext); // JuanPerez
    let counter = 0;
    let finalPath = path.join(baseDir, fileName);

    while (fs.existsSync(finalPath)) {
        counter++;
        finalPath = path.join(baseDir, `${name}(${counter})${ext}`);
    }

    return finalPath;
}


// Generar PDFs
// Handler actualizado para generación con LibreOffice
ipcMain.handle("generate-pdfs-libreoffice", async (event, { baseFolder, excelFile, docxFile }) => {
    try {
        console.log('🚀 Iniciando generación con LibreOffice...');

        if (!libreOfficePath) {
            throw new Error('LibreOffice no está configurado. Por favor configúralo primero.');
        }

        const excelData = readExcelFile(excelFile);
        if (excelData.error) throw new Error(excelData.error);

        const headers = excelData.headers;
        const rows = excelData.rows;
        const salidaPath = path.join(baseFolder, "salida");
        const generatedFiles = [];

        // 🔥 Convertir docxFile a array siempre
        const docxFiles = Array.isArray(docxFile) ? docxFile : [docxFile];

        console.log(`📊 Procesando ${rows.length} filas con ${docxFiles.length} plantillas...`);

        for (let i = 0; i < rows.length; i++) {
            const row = rows[i];

            const data = {};
            headers.forEach((header, idx) => {
                const cleanHeader = String(header).toLowerCase().trim();
                data[cleanHeader] = String(row[idx] || "").trim();
            });

            // 🔄 Para cada plantilla, generar un PDF
            for (let j = 0; j < docxFiles.length; j++) {
                const docxFilePath = docxFiles[j];
                const templateContent = fs.readFileSync(docxFilePath, "binary");

                const zip = new PizZip(templateContent);
                const doc = new Docxtemplater(zip, {
                    paragraphLoop: true,
                    linebreaks: true,
                    nullGetter(part) {
                        return `{${part.value}}`;
                    }
                });

                doc.render(data);

                const docxBuffer = doc.getZip().generate({ type: "nodebuffer" });
                const pdfBuffer = await convertWithLibreOffice(docxBuffer);

                const nombreField = data["nombre"] || data["name"] || "";
                const plantillaName = path.basename(docxFilePath, ".docx");

                const fileName = nombreField
                    ? `${nombreField.replace(/[^a-zA-Z0-9\s\-_]/g, "").trim().replace(/\s+/g, "_")}_${plantillaName}.pdf`
                    : `documento_${String(i + 1).padStart(3, '0')}_${plantillaName}.pdf`;

                const pdfPath = getUniqueFilePath(salidaPath, fileName);
                fs.writeFileSync(pdfPath, pdfBuffer);

                generatedFiles.push(pdfPath);

                // Progreso
                event.sender.send('pdf-progress', {
                    current: i + 1,
                    total: rows.length,
                    fileName: fileName
                });
            }
        }

        return { success: true, files: generatedFiles, total: generatedFiles.length };

    } catch (error) {
        console.error("❌ Error generando PDFs:", error);
        return { success: false, error: error.message };
    }
});

ipcMain.handle("select-docx-multiple", async () => {
    const result = await dialog.showOpenDialog(mainWindow, {
        filters: [{ name: "Word Files", extensions: ["docx"] }],
        properties: ["openFile", "multiSelections"],
    });

    if (result.canceled || result.filePaths.length === 0) return null;
    return result.filePaths; // ← Array de rutas
});

// Handler para verificar LibreOffice
ipcMain.handle("check-libreoffice", async () => {
    try {
        const result = await findLibreOffice();
        if (result.found) {
            libreOfficePath = result.path;
            console.log(`🎯 LibreOffice configurado: ${libreOfficePath}`);
            return { available: true, path: result.path, method: result.method };
        } else {
            return { available: false, path: null, method: null };
        }
    } catch (error) {
        console.error('Error verificando LibreOffice:', error);
        return { available: false, path: null, method: null };
    }
});

// Handler para que el usuario seleccione manualmente LibreOffice
ipcMain.handle("select-libreoffice", async () => {
    try {
        const platform = os.platform();
        let filters = [];

        if (platform === 'win32') {
            filters = [{ name: 'LibreOffice', extensions: ['exe'] }];
        } else {
            filters = [{ name: 'Ejecutables', extensions: ['*'] }];
        }

        const result = await dialog.showOpenDialog(mainWindow, {
            title: 'Selecciona el ejecutable de LibreOffice (soffice.exe o soffice)',
            filters: filters,
            properties: ['openFile'],
            defaultPath: platform === 'win32' ? 'C:\\Program Files\\LibreOffice\\program' : '/usr/bin'
        });

        if (result.canceled || result.filePaths.length === 0) {
            return { success: false, path: null };
        }

        const selectedPath = result.filePaths[0];

        // Verificar que el archivo seleccionado sea válido
        return new Promise((resolve) => {
            exec(`"${selectedPath}" --version`, { timeout: 5000 }, (error, stdout) => {
                if (!error && stdout && stdout.toLowerCase().includes('libreoffice')) {
                    libreOfficePath = selectedPath;
                    console.log(`✅ LibreOffice configurado manualmente: ${selectedPath}`);
                    resolve({
                        success: true,
                        path: selectedPath,
                        version: stdout.trim()
                    });
                } else {
                    console.log(`❌ Archivo seleccionado no es LibreOffice válido: ${selectedPath}`);
                    resolve({
                        success: false,
                        path: null,
                        error: 'El archivo seleccionado no es un ejecutable válido de LibreOffice'
                    });
                }
            });
        });

    } catch (error) {
        console.error('Error seleccionando LibreOffice:', error);
        return { success: false, path: null, error: error.message };
    }
});

ipcMain.handle("select-extra-pdfs", async () => {
    const result = await dialog.showOpenDialog({
        properties: ["openFile", "multiSelections"],
        filters: [{ name: "PDFs", extensions: ["pdf"] }]
    });
    return result.canceled ? [] : result.filePaths;
});

ipcMain.handle("select-single-docx", async () => {
    const result = await dialog.showOpenDialog(mainWindow, {
        filters: [{ name: "Word Files", extensions: ["docx"] }],
        properties: ["openFile"], // Solo una selección
    });

    if (result.canceled || result.filePaths.length === 0) return null;
    return result.filePaths[0]; // Solo devolver la primera (y única) ruta
});

ipcMain.handle("select-single-pdf", async () => {
    const result = await dialog.showOpenDialog(mainWindow, {
        filters: [{ name: "PDF Files", extensions: ["pdf"] }],
        properties: ["openFile"], // Solo una selección
    });

    if (result.canceled || result.filePaths.length === 0) return null;
    return result.filePaths[0]; // Solo devolver la primera (y única) ruta
});

ipcMain.handle("generate-and-send-integrated", async (event, { baseFolder, excelFile, docxFile, emailColumn, subject, body, extraFiles }) => {
    try {
        console.log('🚀 Iniciando proceso integrado: Generar + Enviar...');

        // Validaciones iniciales
        if (!smtpConfig) throw new Error("SMTP no configurado. Guarda la configuración primero.");
        if (!libreOfficePath) throw new Error('LibreOffice no está configurado. Por favor configúralo primero.');

        const transporter = getTransporter();
        const excelData = readExcelFile(excelFile);
        if (excelData.error) throw new Error(excelData.error);

        const headers = excelData.headers;
        const rows = excelData.rows;
        const docxFiles = Array.isArray(docxFile) ? docxFile : [docxFile];

        const headersLower = headers.map(h => String(h).toLowerCase().trim());
        const emailIndex = headersLower.indexOf(emailColumn.toLowerCase().trim());
        if (emailIndex === -1) throw new Error(`No se encontró la columna: ${emailColumn}`);

        const salidaPath = path.join(baseFolder, "salida");
        if (!fs.existsSync(salidaPath)) fs.mkdirSync(salidaPath, { recursive: true });

        console.log(`📊 Procesando ${rows.length} personas con ${docxFiles.length} plantillas...`);

        // Mapa para almacenar PDFs en memoria por persona
        const personaPdfBuffers = new Map();
        const generatedFiles = []; // Para guardar a disco opcionalmente
        let enviados = 0;
        let erroresEnvio = [];

        // FASE 1: GENERAR TODOS LOS PDFs EN MEMORIA
        console.log('📄 FASE 1: Generando PDFs...');

        for (let i = 0; i < rows.length; i++) {
            const row = rows[i];
            const email = row[emailIndex];

            // Skip si no hay email válido
            if (!email || String(email).trim() === "") {
                console.log(`⚠️ Fila ${i+1}: Email vacío, saltando...`);
                continue;
            }

            // Preparar datos de la persona
            const data = {};
            headers.forEach((header, idx) => {
                const cleanHeader = String(header).toLowerCase().trim();
                data[cleanHeader] = String(row[idx] || "").trim();
            });

            const nombreField = data["nombre"] || data["name"] || "";
            const safeName = nombreField
                ? nombreField.replace(/[^a-zA-Z0-9\s\-_]/g, "").trim().replace(/\s+/g, "_")
                : `documento_${String(i + 1).padStart(3, '0')}`;

            // Array para almacenar los PDFs de esta persona
            const personaPdfs = [];

            // Generar PDF para cada plantilla
            for (let j = 0; j < docxFiles.length; j++) {
                const docxFilePath = docxFiles[j];

                try {
                    // Leer y procesar plantilla
                    const templateContent = fs.readFileSync(docxFilePath, "binary");
                    const zip = new PizZip(templateContent);
                    const doc = new Docxtemplater(zip, {
                        paragraphLoop: true,
                        linebreaks: true,
                        nullGetter(part) {
                            return `{${part.value}}`;
                        }
                    });

                    doc.render(data);
                    const docxBuffer = doc.getZip().generate({ type: "nodebuffer" });

                    // Convertir a PDF con LibreOffice
                    const pdfBuffer = await convertWithLibreOffice(docxBuffer);

                    const plantillaName = path.basename(docxFilePath, ".docx");
                    const fileName = `${safeName}_${plantillaName}.pdf`;

                    // Almacenar en memoria
                    personaPdfs.push({
                        filename: fileName,
                        buffer: pdfBuffer,
                        path: path.join(salidaPath, fileName) // Para guardado opcional
                    });

                    // Progreso de generación
                    const totalOperaciones = rows.length * docxFiles.length;
                    const operacionActual = (i * docxFiles.length) + (j + 1);

                    event.sender.send('integrated-progress', {
                        phase: 'generating',
                        current: operacionActual,
                        total: totalOperaciones,
                        person: nombreField || `Persona ${i+1}`,
                        template: plantillaName
                    });

                } catch (pdfError) {
                    console.error(`❌ Error generando PDF para ${email} con plantilla ${docxFilePath}:`, pdfError.message);
                    // Continuar con la siguiente plantilla
                }
            }

            // Almacenar PDFs de la persona en el mapa
            if (personaPdfs.length > 0) {
                personaPdfBuffers.set(email, personaPdfs);
            }
        }

        console.log(`✅ FASE 1 completada. PDFs generados para ${personaPdfBuffers.size} personas.`);

        // FASE 2: ENVIAR CORREOS CON PDFs DESDE MEMORIA
        console.log('📨 FASE 2: Enviando correos...');

        for (const [email, personaPdfs] of personaPdfBuffers) {
            try {
                // Preparar adjuntos: PDFs generados + PDFs extra
                let attachments = [];

                // Agregar PDFs generados desde memoria
                personaPdfs.forEach(pdf => {
                    attachments.push({
                        filename: pdf.filename,
                        content: pdf.buffer,
                        contentType: 'application/pdf'
                    });
                });

                // Agregar PDFs extra desde disco
                if (extraFiles && extraFiles.length > 0) {
                    extraFiles
                        .filter(file => typeof file === "string" && file.trim() !== "" && fs.existsSync(file))
                        .forEach(file => {
                            attachments.push({
                                filename: path.basename(file),
                                path: file
                            });
                        });
                }

                // Enviar email
                await transporter.sendMail({
                    from: smtpConfig.user,
                    to: email,
                    subject: subject,
                    text: body,
                    attachments: attachments
                });

                enviados++;
                console.log(`✅ Enviado a ${email} (${attachments.length} adjuntos)`);

                // Progreso de envío
                event.sender.send('integrated-progress', {
                    phase: 'sending',
                    current: enviados,
                    total: personaPdfBuffers.size,
                    email: email,
                    attachments: attachments.length
                });

            } catch (emailError) {
                console.error(`❌ Error enviando a ${email}:`, emailError.message);
                erroresEnvio.push({ email, error: emailError.message });
            }
        }

        // FASE 3: GUARDAR PDFs A DISCO (OPCIONAL)
        console.log('💾 FASE 3: Guardando PDFs a disco...');

        for (const [email, personaPdfs] of personaPdfBuffers) {
            for (const pdf of personaPdfs) {
                try {
                    const finalPath = getUniqueFilePath(salidaPath, pdf.filename);
                    fs.writeFileSync(finalPath, pdf.buffer);
                    generatedFiles.push(finalPath);
                } catch (saveError) {
                    console.error(`⚠️ Error guardando ${pdf.filename}:`, saveError.message);
                }
            }
        }

        // Resultado final
        const resultado = {
            success: true,
            enviados: enviados,
            totalPersonas: personaPdfBuffers.size,
            pdfGenerados: generatedFiles.length,
            erroresEnvio: erroresEnvio.length,
            files: generatedFiles
        };

        console.log('🎉 PROCESO INTEGRADO COMPLETADO:', resultado);
        return resultado;

    } catch (error) {
        console.error("❌ Error en proceso integrado:", error);
        return {
            success: false,
            error: error.message,
            enviados: enviados || 0,
            erroresEnvio: erroresEnvio || []
        };
    }
});


// Función para usar LibreOffice con la ruta configurada
const convertWithLibreOffice = (docxBuffer) => {
    return new Promise((resolve, reject) => {
        if (!libreOfficePath) {
            reject(new Error('LibreOffice no está configurado'));
            return;
        }

        // Si tenemos una ruta específica, intentar usarla directamente
        if (libreOfficePath.includes('soffice')) {
            // Método directo con subprocess si es necesario
            const libre = require('libreoffice-convert');
            libre.convert(docxBuffer, '.pdf', undefined, (err, result) => {
                if (err) {
                    console.error('Error LibreOffice:', err);
                    reject(new Error(`Error de conversión: ${err.message}`));
                } else {
                    resolve(result);
                }
            });
        } else {
            reject(new Error('Ruta de LibreOffice no válida'));
        }
    });
};



app.whenReady().then(() => {
    createWindow();

    app.on("activate", () => {
        if (BrowserWindow.getAllWindows().length === 0) createWindow();
    });
});

app.on("window-all-closed", () => {
    if (process.platform !== "darwin") app.quit();
});