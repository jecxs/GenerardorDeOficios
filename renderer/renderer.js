const baseFolderEl = document.getElementById("baseFolder");
const excelFileEl = document.getElementById("excelFile");
const docxFileEl = document.getElementById("docxFile");
const logEl = document.getElementById("log");

let baseFolder = null;
let excelFile = null;
let docxFile = null;
let libreOfficeStatus = { available: false, path: null, method: null };
let extraFiles = [];
let appConfig = null;
let selectedTemplates = [];
let selectedExtraPdfs = []

// Verificar LibreOffice al cargar la página
window.addEventListener('DOMContentLoaded', async () => {
    // Cargar configuración guardada
    await loadAppConfig();

    // Verificar LibreOffice
    await checkLibreOfficeStatus();

    // Intentar cargar carpeta base automáticamente si existe en config
    await autoLoadBaseFolder();
});


// Función para cargar configuración de la aplicación
async function loadAppConfig() {
    try {
        appConfig = await window.electronAPI.loadConfig();
        console.log("📂 Configuración cargada:", appConfig);

        // Si hay configuración SMTP, cargar en la interfaz
        if (appConfig.smtpConfig) {
            loadSmtpConfigToUI(appConfig.smtpConfig);
        }

    } catch (error) {
        console.error("Error cargando configuración:", error);
        appConfig = { baseFolder: null, smtpConfig: null };
    }
}

function loadSmtpConfigToUI(smtpConfig) {
    document.getElementById("smtpHost").value = smtpConfig.host || "";
    document.getElementById("smtpPort").value = smtpConfig.port || 587;
    document.getElementById("smtpSecure").value = String(smtpConfig.secure || false);
    document.getElementById("smtpUser").value = smtpConfig.user || "";
    document.getElementById("smtpPass").value = smtpConfig.pass || "";

    logEl.textContent = `✅ Configuración SMTP cargada automáticamente:\nServidor: ${smtpConfig.host}:${smtpConfig.port}\nUsuario: ${smtpConfig.user}`;
}
// Función para cargar automáticamente la carpeta base
async function autoLoadBaseFolder() {
    if (appConfig && appConfig.baseFolder) {
        logEl.textContent = "🔍 Verificando última carpeta base utilizada...";

        try {
            // 🔥 USAR LA NUEVA FUNCIÓN que NO abre diálogo
            const result = await window.electronAPI.loadSavedFolder();

            if (result && result.basePath) {
                baseFolder = result.basePath;
                baseFolderEl.textContent = baseFolder;

                logEl.textContent = `✅ Carpeta base cargada automáticamente: ${baseFolder}`;
                if (result.created && result.created.length > 0) {
                    logEl.textContent += `\n📁 Subcarpetas verificadas: ${result.created.join(", ")}`;
                }

                // Efecto visual de carga automática
                baseFolderEl.style.backgroundColor = "#d4edda";
                baseFolderEl.style.padding = "5px";
                baseFolderEl.style.borderRadius = "3px";
                baseFolderEl.style.border = "1px solid #c3e6cb";

                setTimeout(() => {
                    baseFolderEl.style.backgroundColor = "";
                    baseFolderEl.style.padding = "";
                    baseFolderEl.style.borderRadius = "";
                    baseFolderEl.style.border = "";
                }, 3000);

            } else {
                // Si no hay carpeta guardada o no existe, no mostrar error
                logEl.textContent = `📂 Listo para comenzar. Selecciona una carpeta base.`;
            }
        } catch (error) {
            console.error("Error cargando carpeta base:", error);
            logEl.textContent = "📂 Listo para comenzar. Selecciona una carpeta base.";
        }
    } else {
        logEl.textContent = "📂 Listo para comenzar. Selecciona una carpeta base.";
    }
}

// Función para verificar estado de LibreOffice
async function checkLibreOfficeStatus() {
    logEl.textContent = "🔍 Verificando LibreOffice...";
    try {
        libreOfficeStatus = await window.electronAPI.checkLibreOffice();
        updateLibreOfficeUI();
    } catch (error) {
        logEl.textContent = "⚠️ Error verificando LibreOffice.";
        console.error("Error verificando LibreOffice:", error);
    }
}

// Actualizar interfaz según estado de LibreOffice
function updateLibreOfficeUI() {
    if (libreOfficeStatus.available) {
        logEl.textContent = `✅ LibreOffice detectado correctamente!\n` +
            `📍 Método: ${libreOfficeStatus.method}\n` +
            `🎯 Los PDFs mantendrán el formato original de tu plantilla.`;

        // Cambiar color del botón generar a verde
        const generateBtn = document.getElementById("btnGenerate");
        if (generateBtn) {
            generateBtn.style.backgroundColor = "#48bb78";
            generateBtn.style.borderColor = "#38a169";
        }
    } else {
        logEl.textContent = `❌ LibreOffice no detectado automáticamente.\n\n` +
            `🔧 Opciones:\n` +
            `1. Instalar LibreOffice desde: https://www.libreoffice.org/\n` +
            `2. O hacer clic en "Configurar LibreOffice" para seleccionarlo manualmente.`;

        // Mostrar botón de configuración manual
        showLibreOfficeConfigButton();
    }
}

// Mostrar botón para configurar LibreOffice manualmente
function showLibreOfficeConfigButton() {
    const existingBtn = document.getElementById("btnConfigLibre");
    if (existingBtn) return; // Ya existe

    const configSection = document.createElement("div");
    configSection.className = "section";
    configSection.style.backgroundColor = "#fff3cd";
    configSection.style.border = "1px solid #ffeaa7";
    configSection.style.borderRadius = "5px";
    configSection.style.padding = "15px";
    configSection.style.margin = "10px 0";

    const configBtn = document.createElement("button");
    configBtn.id = "btnConfigLibre";
    configBtn.textContent = "🔧 Configurar LibreOffice manualmente";
    configBtn.style.backgroundColor = "#f39c12";
    configBtn.style.color = "white";
    configBtn.style.border = "none";
    configBtn.style.padding = "10px 20px";
    configBtn.style.borderRadius = "5px";
    configBtn.style.cursor = "pointer";
    configBtn.style.fontSize = "14px";

    const configText = document.createElement("p");
    configText.innerHTML = `<strong>💡 Ayuda:</strong> Busca el archivo <code>soffice.exe</code> en:<br>` +
        `📂 <code>C:\\Program Files\\LibreOffice\\program\\soffice.exe</code>`;
    configText.style.marginTop = "10px";
    configText.style.fontSize = "12px";

    configBtn.addEventListener("click", async () => {
        configBtn.disabled = true;
        configBtn.textContent = "🔍 Buscando...";

        try {
            const result = await window.electronAPI.selectLibreOffice();

            if (result.success) {
                libreOfficeStatus = {
                    available: true,
                    path: result.path,
                    method: 'manual'
                };

                logEl.textContent = `✅ LibreOffice configurado correctamente!\n` +
                    `📍 Ubicación: ${result.path}\n` +
                    `🎯 Ahora puedes generar PDFs con formato original.`;

                // Ocultar sección de configuración
                configSection.style.display = "none";

                // Cambiar color del botón generar
                const generateBtn = document.getElementById("btnGenerate");
                if (generateBtn) {
                    generateBtn.style.backgroundColor = "#48bb78";
                    generateBtn.style.borderColor = "#38a169";
                }

            } else {
                logEl.textContent = `❌ Error configurando LibreOffice:\n${result.error || "Archivo no válido"}\n\n` +
                    `💡 Asegúrate de seleccionar el archivo "soffice.exe" correcto.`;

                configBtn.disabled = false;
                configBtn.textContent = "🔧 Configurar LibreOffice manualmente";
            }
        } catch (error) {
            logEl.textContent = `❌ Error inesperado: ${error.message}`;
            configBtn.disabled = false;
            configBtn.textContent = "🔧 Configurar LibreOffice manualmente";
        }
    });

    configSection.appendChild(configBtn);
    configSection.appendChild(configText);

    // Insertar después del log
    logEl.parentNode.insertBefore(configSection, logEl.nextSibling);
}
function updateExtraPdfsUI() {
    const pdfsList = document.getElementById("extraPdfsList");
    const noPdfs = document.getElementById("noExtraPdfs");
    const pdfCount = document.getElementById("extraPdfCount");

    // Actualizar contador
    pdfCount.textContent = selectedExtraPdfs.length;

    if (selectedExtraPdfs.length === 0) {
        // Mostrar mensaje de "no hay PDFs"
        pdfsList.innerHTML = '<p id="noExtraPdfs" style="color: #666; font-style: italic;">No hay PDFs extra seleccionados</p>';
        extraFiles = []; // Limpiar variable global
        return;
    }

    // Mostrar lista de PDFs
    let listHTML = '<div style="border: 1px solid #ddd; border-radius: 5px; padding: 10px; background-color: #f8f9fa;">';

    selectedExtraPdfs.forEach((pdfPath, index) => {
        const fileName = pdfPath.split(/[\\/]/).pop();

        listHTML += `
            <div style="display: flex; justify-content: space-between; align-items: center; padding: 8px; margin: 5px 0; background-color: white; border-radius: 3px; border: 1px solid #e9ecef;">
                <span style="flex: 1; font-weight: 500;">
                    📄 ${fileName}
                </span>
                <div style="display: flex; gap: 5px;">
                    <small style="color: #666; padding: 0 10px;">${pdfPath}</small>
                    <button onclick="removeExtraPdf(${index})" 
                            style="background-color: #dc3545; color: white; border: none; padding: 4px 8px; border-radius: 3px; cursor: pointer; font-size: 12px;">
                        ❌ Quitar
                    </button>
                </div>
            </div>
        `;
    });

    listHTML += '</div>';
    pdfsList.innerHTML = listHTML;

    // Actualizar variable global para compatibilidad
    extraFiles = [...selectedExtraPdfs];
}
function removeExtraPdf(index) {
    const removedFile = selectedExtraPdfs[index];
    selectedExtraPdfs.splice(index, 1);
    updateExtraPdfsUI();

    const fileName = removedFile.split(/[\\/]/).pop();
    logEl.textContent = `🗑️ PDF eliminado: ${fileName}\nTotal: ${selectedExtraPdfs.length} PDFs extra`;

    // Si no quedan PDFs, mostrar mensaje
    if (selectedExtraPdfs.length === 0) {
        logEl.textContent += "\n📎 Los PDFs extra se adjuntarán a todos los correos enviados.";
    }
}
function clearAllExtraPdfs() {
    if (selectedExtraPdfs.length === 0) {
        logEl.textContent = "⚠️ No hay PDFs extra para limpiar.";
        return;
    }

    const confirmed = confirm(`¿Estás seguro de eliminar todos los ${selectedExtraPdfs.length} PDFs extra seleccionados?`);
    if (confirmed) {
        selectedExtraPdfs = [];
        updateExtraPdfsUI();
        logEl.textContent = "🗑️ Todos los PDFs extra han sido eliminados.";
    }
}

function updateTemplatesUI() {
    const templatesList = document.getElementById("templatesList");
    const noTemplates = document.getElementById("noTemplates");
    const templateCount = document.getElementById("templateCount");

    // Actualizar contador
    templateCount.textContent = selectedTemplates.length;

    if (selectedTemplates.length === 0) {
        // Mostrar mensaje de "no hay plantillas"
        templatesList.innerHTML = '<p id="noTemplates" style="color: #666; font-style: italic;">No hay plantillas seleccionadas</p>';
        docxFile = null; // Limpiar variable global
        return;
    }

    // Ocultar mensaje y mostrar lista
    let listHTML = '<div style="border: 1px solid #ddd; border-radius: 5px; padding: 10px; background-color: #f8f9fa;">';

    selectedTemplates.forEach((template, index) => {
        const fileName = template.split(/[\\/]/).pop();
        listHTML += `
            <div style="display: flex; justify-content: space-between; align-items: center; padding: 8px; margin: 5px 0; background-color: white; border-radius: 3px; border: 1px solid #e9ecef;">
                <span style="flex: 1; font-weight: 500;">
                    📄 ${fileName}
                </span>
                <div style="display: flex; gap: 5px;">
                    <small style="color: #666; padding: 0 10px;">${template}</small>
                    <button onclick="removeTemplate(${index})" 
                            style="background-color: #dc3545; color: white; border: none; padding: 4px 8px; border-radius: 3px; cursor: pointer; font-size: 12px;">
                        ❌ Quitar
                    </button>
                </div>
            </div>
        `;
    });

    listHTML += '</div>';
    templatesList.innerHTML = listHTML;

    // Actualizar variable global para compatibilidad
    docxFile = selectedTemplates.length > 0 ? selectedTemplates : null;
}
function removeTemplate(index) {
    selectedTemplates.splice(index, 1);
    updateTemplatesUI();

    logEl.textContent = `🗑️ Plantilla eliminada. Total: ${selectedTemplates.length} plantillas`;

    // Si no quedan plantillas, mostrar mensaje
    if (selectedTemplates.length === 0) {
        logEl.textContent += "\n📝 Agrega al menos una plantilla para continuar.";
    }
}

function clearAllTemplates() {
    if (selectedTemplates.length === 0) {
        logEl.textContent = "⚠️ No hay plantillas para limpiar.";
        return;
    }

    const confirmed = confirm(`¿Estás seguro de eliminar todas las ${selectedTemplates.length} plantillas seleccionadas?`);
    if (confirmed) {
        selectedTemplates = [];
        updateTemplatesUI();
        logEl.textContent = "🗑️ Todas las plantillas han sido eliminadas.";
    }
}

document.getElementById("btnAddExtraPdf").addEventListener("click", async () => {
    try {
        // Usar la función existente que ya permite selección múltiple
        const newPdfs = await window.electronAPI.selectExtraPdfs();

        if (newPdfs && newPdfs.length > 0) {
            let addedCount = 0;
            let duplicateCount = 0;

            newPdfs.forEach(pdfPath => {
                if (!selectedExtraPdfs.includes(pdfPath)) {
                    selectedExtraPdfs.push(pdfPath);
                    addedCount++;
                } else {
                    duplicateCount++;
                }
            });

            updateExtraPdfsUI();

            let message;
            if (addedCount === 1) {
                const fileName = newPdfs[0].split(/[\\/]/).pop();
                message = `✅ PDF agregado: ${fileName}\n📊 Total: ${selectedExtraPdfs.length} PDFs extra`;
            } else {
                message = `✅ ${addedCount} PDFs agregados correctamente.\n📊 Total: ${selectedExtraPdfs.length} PDFs extra`;
            }

            if (duplicateCount > 0) {
                message += `\n⚠️ ${duplicateCount} archivos omitidos (ya estaban agregados)`;
            }

            if (selectedExtraPdfs.length === addedCount && addedCount > 0) {
                message += "\n💡 Estos PDFs se adjuntarán a todos los correos que envíes.";
            }

            logEl.textContent = message;
        }
    } catch (error) {
        logEl.textContent = `❌ Error agregando PDFs: ${error.message}`;
    }
});

// Event listener para limpiar todos los PDFs extra
document.getElementById("btnClearExtraPdfs").addEventListener("click", () => {
    clearAllExtraPdfs();
});


// Seleccionar carpeta base
document.getElementById("btnBaseFolder").addEventListener("click", async () => {
    const result = await window.electronAPI.selectBaseFolder(); // Sin parámetros = mostrar diálogo
    baseFolder = result?.basePath || null;
    baseFolderEl.textContent = baseFolder || "Ninguna";

    if (baseFolder) {
        logEl.textContent = `📂 Nueva carpeta seleccionada: ${baseFolder}`;
        if (result.created && result.created.length > 0) {
            logEl.textContent += `\n📁 Subcarpetas creadas: ${result.created.join(", ")}`;
        }
        logEl.textContent += `\n💾 Carpeta guardada para próximas sesiones.`;

        // Guardar en configuración
        await window.electronAPI.saveConfig({ baseFolder: baseFolder });
    }
});

// Seleccionar Excel
document.getElementById("btnExcel").addEventListener("click", async () => {
    excelFile = await window.electronAPI.selectExcel();
    excelFileEl.textContent = excelFile ? excelFile.split(/[\\/]/).pop() : "Ninguno";

    if (excelFile) {
        logEl.textContent = "📊 Cargando Excel...";
        const data = await window.electronAPI.readExcel(excelFile);

        if (data.error) {
            logEl.textContent = `❌ Error cargando Excel: ${data.error}`;
        } else {
            renderPreview(data.headers, data.previewRows);
            logEl.textContent = `✅ Excel cargado correctamente.\n📈 ${data.totalRows} filas de datos encontradas.`;

            // 🔥 Llenar combo con headers
            const emailColumnSelect = document.getElementById("emailColumn");
            emailColumnSelect.innerHTML = "";
            data.headers.forEach((header, idx) => {
                const opt = document.createElement("option");
                opt.value = header;
                opt.textContent = header;
                emailColumnSelect.appendChild(opt);
            });
        }
    }
});

document.getElementById("btnSaveSMTP").addEventListener("click", async () => {
    const smtpConfig = {
        host: document.getElementById("smtpHost").value,
        port: parseInt(document.getElementById("smtpPort").value, 10),
        secure: document.getElementById("smtpSecure").value === "true",
        user: document.getElementById("smtpUser").value,
        pass: document.getElementById("smtpPass").value
    };

    await window.electronAPI.saveSmtpConfig(smtpConfig);

    // Actualizar configuración local
    if (!appConfig) appConfig = {};
    appConfig.smtpConfig = smtpConfig;

    logEl.textContent = `✅ Configuración SMTP guardada y recordada:\nServidor: ${smtpConfig.host}:${smtpConfig.port}\nUsuario: ${smtpConfig.user}`;
});


// Seleccionar plantilla DOCX
document.getElementById("btnDocx").addEventListener("click", async () => {
    try {
        const newTemplate = await window.electronAPI.selectSingleDocx();

        if (newTemplate) {
            // Verificar si ya está agregada
            if (selectedTemplates.includes(newTemplate)) {
                logEl.textContent = "⚠️ Esta plantilla ya está agregada.";
                return;
            }

            // Agregar a la lista
            selectedTemplates.push(newTemplate);
            updateTemplatesUI();

            const fileName = newTemplate.split(/[\\/]/).pop();
            logEl.textContent = `✅ Plantilla agregada: ${fileName}\n📊 Total: ${selectedTemplates.length} plantillas`;

            if (selectedTemplates.length === 1) {
                logEl.textContent += "\n💡 Puedes agregar más plantillas o continuar con la generación.";
            }
        }
    } catch (error) {
        logEl.textContent = `❌ Error agregando plantilla: ${error.message}`;
    }
});

document.getElementById("btnClearTemplates").addEventListener("click", () => {
    clearAllTemplates();
});
// Generar PDFs
document.getElementById("btnGenerate").addEventListener("click", async () => {
    if (!baseFolder || !excelFile || !selectedTemplates || selectedTemplates.length === 0) {
        logEl.textContent = "⚠️ ERROR: Faltan elementos por seleccionar:\n" +
            `${!baseFolder ? "❌ Carpeta base\n" : ""}` +
            `${!excelFile ? "❌ Archivo Excel\n" : ""}` +
            `${!selectedTemplates || selectedTemplates.length === 0 ? "❌ Al menos una plantilla DOCX\n" : ""}`;
        return;
    }

    // Verificar LibreOffice antes de generar
    if (!libreOfficeStatus.available) {
        const tryAgain = confirm(
            "❌ LibreOffice no está configurado.\n\n" +
            "Sin LibreOffice no se pueden generar PDFs con el formato original.\n\n" +
            "¿Quieres intentar detectarlo de nuevo?"
        );

        if (tryAgain) {
            await checkLibreOfficeStatus();
            return;
        } else {
            logEl.textContent = "❌ Generación cancelada. Configura LibreOffice primero.";
            return;
        }
    }

    // Configurar listener de progreso
    window.electronAPI.onPdfProgress((progress) => {
        const percentage = Math.round((progress.current / progress.total) * 100);
        logEl.textContent = `⏳ Generando PDFs... ${percentage}%\n` +
            `📄 Procesando ${progress.current}/${progress.total}: ${progress.fileName}\n\n` +
            `🔧 Usando LibreOffice: ${libreOfficeStatus.path}`;
    });

    // Iniciar generación
    logEl.textContent = `🚀 Iniciando generación de PDFs...\n⏳ Esto puede tomar unos minutos...\n\n` +
        `📝 ${selectedTemplates.length} plantillas × personas = múltiples PDFs\n` +
        `🔧 LibreOffice: ${libreOfficeStatus.path}`;

    try {
        const result = await window.electronAPI.generatePdfsLibreOffice({
            baseFolder,
            excelFile,
            docxFile: selectedTemplates // Pasar el array de plantillas
        });

        if (result.success) {
            logEl.textContent = `🎉 ¡GENERACIÓN COMPLETADA EXITOSAMENTE!\n\n` +
                `✅ ${result.total} PDFs generados con formato original\n` +
                `📁 Ubicación: ${baseFolder}\\salida\n\n`;

            // Mostrar primeros archivos generados
            if (result.files && result.files.length > 0) {
                const firstFiles = result.files.slice(0, 3).map(f => f.split(/[\\/]/).pop());
                logEl.textContent += `📄 Archivos: ${firstFiles.join(", ")}`;
                if (result.files.length > 3) {
                    logEl.textContent += ` y ${result.files.length - 3} más...`;
                }
            }

            logEl.textContent += `\n\n🎯 Los PDFs mantienen el formato exacto de tus plantillas DOCX.`;

            // Efecto visual de éxito
            logEl.style.backgroundColor = "#d4edda";
            logEl.style.border = "1px solid #c3e6cb";


        } else {
            logEl.textContent = `❌ ERROR EN LA GENERACIÓN:\n\n${result.error}\n\n` +
                `🔧 Posibles soluciones:\n` +
                `• Verifica que las variables en la plantilla sean como {nombre}\n` +
                `• Asegúrate que las columnas del Excel coincidan\n` +
                `• Cierra archivos abiertos en la carpeta de salida\n` +
                `• Verifica permisos de escritura en la carpeta\n` +
                `• Reinicia LibreOffice si está abierto`;

            // Efecto visual de error
            logEl.style.backgroundColor = "#f8d7da";
            logEl.style.border = "1px solid #f5c6cb";
            logEl.style.color = "#721c24";
        }
    } catch (error) {
        logEl.textContent = `❌ ERROR CRÍTICO:\n\n${error.message}\n\n` +
            `🚨 Acciones recomendadas:\n` +
            `• Reinicia la aplicación\n` +
            `• Verifica que LibreOffice funcione correctamente\n` +
            `• Contacta soporte técnico si persiste`;
        console.error("Error completo:", error);

        // Efecto visual de error crítico
        logEl.style.backgroundColor = "#f8d7da";
        logEl.style.border = "1px solid #f5c6cb";
        logEl.style.color = "#721c24";
    } finally {
        // Limpiar listeners de progreso
        window.electronAPI.removeAllListeners();

        // Restaurar estilo después de 10 segundos
        setTimeout(() => {
            logEl.style.backgroundColor = "";
            logEl.style.border = "";
            logEl.style.color = "#ffffff";
            logEl.classList.remove('text-green-400', 'text-red-400');
            logEl.classList.add('text-green');
        }, 10000);
    }
});

document.getElementById("btnGenerateAndSend").addEventListener("click", async () => {
    // Validaciones iniciales
    if (!baseFolder || !excelFile || !selectedTemplates || selectedTemplates.length === 0) {
        logEl.textContent = "⚠️ ERROR: Faltan elementos por seleccionar:\n" +
            `${!baseFolder ? "❌ Carpeta base\n" : ""}` +
            `${!excelFile ? "❌ Archivo Excel\n" : ""}` +
            `${!selectedTemplates || selectedTemplates.length === 0 ? "❌ Al menos una plantilla DOCX\n" : ""}`;
        return;
    }

    // Validar configuración SMTP
    const emailColumn = document.getElementById("emailColumn").value;
    const subject = document.getElementById("emailSubject").value || "Documento generado";
    const body = document.getElementById("emailBody").value || "Adjunto encontrarás tus documentos.";

    if (!emailColumn) {
        logEl.textContent = "⚠️ ERROR: Selecciona la columna de correos antes de enviar.";
        return;
    }

    // Verificar LibreOffice
    if (!libreOfficeStatus.available) {
        const tryAgain = confirm(
            "❌ LibreOffice no está configurado.\n\n" +
            "Sin LibreOffice no se pueden generar PDFs.\n\n" +
            "¿Quieres intentar detectarlo de nuevo?"
        );

        if (tryAgain) {
            await checkLibreOfficeStatus();
            return;
        } else {
            logEl.textContent = "❌ Proceso cancelado. Configura LibreOffice primero.";
            return;
        }
    }

    // Confirmación antes de procesar
    const excelData = await window.electronAPI.readExcel(excelFile);
    const totalPersonas = excelData.totalRows || 0;
    const totalPdfs = totalPersonas * selectedTemplates.length;

    const confirmed = confirm(
        `🚀 PROCESO INTEGRADO: Generar + Enviar\n\n` +
        `📊 Resumen:\n` +
        `• ${totalPersonas} personas\n` +
        `• ${selectedTemplates.length} plantillas\n` +
        `• ${totalPdfs} PDFs a generar\n` +
        `• ${selectedExtraPdfs.length} PDFs extra por correo\n\n` +
        `⏱️ Tiempo estimado: ${Math.ceil(totalPersonas / 60)} - ${Math.ceil(totalPersonas / 30)} minutos\n\n` +
        `¿Continuar con el proceso completo?`
    );

    if (!confirmed) {
        logEl.textContent = "❌ Proceso cancelado por el usuario.";
        return;
    }

    // Configurar listeners de progreso
    window.electronAPI.onIntegratedProgress((progress) => {
        if (progress.phase === 'generating') {
            const percentage = Math.round((progress.current / progress.total) * 100);
            logEl.textContent = `📄 GENERANDO PDFs... ${percentage}%\n` +
                `⏳ Procesando: ${progress.person}\n` +
                `📝 Plantilla: ${progress.template}\n` +
                `📊 Progreso: ${progress.current}/${progress.total}\n\n` +
                `🔧 LibreOffice: ${libreOfficePath}`;
        } else if (progress.phase === 'sending') {
            const percentage = Math.round((progress.current / progress.total) * 100);
            logEl.textContent = `📨 ENVIANDO CORREOS... ${percentage}%\n` +
                `📧 Enviando a: ${progress.email}\n` +
                `📎 Adjuntos: ${progress.attachments}\n` +
                `📊 Progreso: ${progress.current}/${progress.total}`;
        }
    });

    // Iniciar proceso integrado
    logEl.textContent = `🚀 INICIANDO PROCESO INTEGRADO...\n\n` +
        `📋 FASE 1: Generando ${totalPdfs} PDFs...\n` +
        `📨 FASE 2: Enviando ${totalPersonas} correos...\n` +
        `💾 FASE 3: Guardando archivos...\n\n` +
        `⏳ Esto puede tomar varios minutos, por favor espera...`;

    try {
        const result = await window.electronAPI.generateAndSendIntegrated({
            baseFolder,
            excelFile,
            docxFile: selectedTemplates,
            emailColumn,
            subject,
            body,
            extraFiles: selectedExtraPdfs
        });

        if (result.success) {
            logEl.textContent = `🎉 ¡PROCESO COMPLETADO EXITOSAMENTE!\n\n` +
                `✅ Resultados:\n` +
                `📨 ${result.enviados} correos enviados\n` +
                `📄 ${result.pdfGenerados} PDFs generados y guardados\n` +
                `👥 ${result.totalPersonas} personas procesadas\n\n`;

            if (result.erroresEnvio > 0) {
                logEl.textContent += `⚠️ ${result.erroresEnvio} errores de envío (ver consola para detalles)\n`;
            }

            logEl.textContent += `📁 Archivos guardados en: ${baseFolder}\\salida\n\n` +
                `🎯 Todos los PDFs se generaron y enviaron automáticamente.`;

            // Efecto visual de éxito
            logEl.style.backgroundColor = "#d4edda";
            logEl.style.border = "1px solid #c3e6cb";
            logEl.style.color = "#155724";

        } else {
            logEl.textContent = `❌ ERROR EN EL PROCESO INTEGRADO:\n\n${result.error}\n\n`;

            if (result.enviados > 0) {
                logEl.textContent += `✅ Correos enviados antes del error: ${result.enviados}\n\n`;
            }

            logEl.textContent += `🔧 Posibles soluciones:\n` +
                `• Verifica configuración SMTP\n` +
                `• Asegúrate que LibreOffice funcione correctamente\n` +
                `• Revisa que las variables en plantillas sean {nombre}\n` +
                `• Verifica conexión a internet\n` +
                `• Cierra archivos abiertos en carpeta salida`;

            // Efecto visual de error
            logEl.style.backgroundColor = "#f8d7da";
            logEl.style.border = "1px solid #f5c6cb";
            logEl.style.color = "#721c24";
        }

    } catch (error) {
        logEl.textContent = `❌ ERROR CRÍTICO EN PROCESO INTEGRADO:\n\n${error.message}\n\n` +
            `🚨 Acciones recomendadas:\n` +
            `• Reinicia la aplicación\n` +
            `• Verifica configuración completa\n` +
            `• Contacta soporte técnico si persiste`;
        console.error("Error completo:", error);

        // Efecto visual de error crítico
        logEl.style.backgroundColor = "#f8d7da";
        logEl.style.border = "1px solid #f5c6cb";
        logEl.style.color = "#721c24";
    } finally {
        // Limpiar listeners
        window.electronAPI.removeAllListeners();

        // Restaurar estilo después de 10 segundos
        setTimeout(() => {
            logEl.style.backgroundColor = "";
            logEl.style.border = "";
            logEl.style.color = "#ffffff";
            logEl.classList.remove('text-green-400', 'text-red-400');
            logEl.classList.add('text-green');
        }, 10000);
    }
});


// Función para renderizar vista previa del Excel
function renderPreview(headers, rows) {
    const headersEl = document.getElementById("previewHeaders");
    const bodyEl = document.getElementById("previewBody");
    const noPreviewEl = document.getElementById("noPreview");

    // Limpiar contenido previo
    headersEl.innerHTML = "";
    bodyEl.innerHTML = "";
    noPreviewEl.style.display = "none";

    if (!headers || headers.length === 0) {
        noPreviewEl.style.display = "block";
        noPreviewEl.textContent = "❌ No se encontraron columnas en el Excel.";
        return;
    }

    // Renderizar headers
    const headerRow = document.createElement("tr");
    headers.forEach((header, index) => {
        const th = document.createElement("th");
        th.textContent = header || `Columna ${index + 1}`;
        th.style.backgroundColor = "#2c5282";
        th.style.color = "white";
        th.style.fontWeight = "bold";
        th.style.padding = "12px 8px";
        th.style.border = "1px solid #2d3748";
        th.style.textAlign = "left";
        headerRow.appendChild(th);
    });
    headersEl.appendChild(headerRow);

    // Renderizar filas de datos
    if (!rows || rows.length === 0) {
        const noDataRow = document.createElement("tr");
        const noDataCell = document.createElement("td");
        noDataCell.colSpan = headers.length;
        noDataCell.textContent = "⚠️ No hay datos para mostrar en la vista previa";
        noDataCell.style.textAlign = "center";
        noDataCell.style.fontStyle = "italic";
        noDataCell.style.padding = "20px";
        noDataCell.style.color = "#666";
        noDataRow.appendChild(noDataCell);
        bodyEl.appendChild(noDataRow);
        return;
    }

    rows.forEach((row, rowIndex) => {
        const tr = document.createElement("tr");

        // Alternar colores de filas
        tr.style.backgroundColor = rowIndex % 2 === 0 ? "#f7fafc" : "white";

        headers.forEach((header, colIndex) => {
            const td = document.createElement("td");
            const cellValue = row[colIndex];
            td.textContent = cellValue !== null && cellValue !== undefined ? String(cellValue) : "";
            td.style.padding = "8px";
            td.style.border = "1px solid #e2e8f0";
            td.style.maxWidth = "200px";
            td.style.overflow = "hidden";
            td.style.textOverflow = "ellipsis";
            td.style.whiteSpace = "nowrap";
            td.title = td.textContent; // Tooltip para texto largo
            tr.appendChild(td);
        });

        bodyEl.appendChild(tr);
    });

    // Agregar información adicional
    if (rows.length >= 5) {
        const infoRow = document.createElement("tr");
        const infoCell = document.createElement("td");
        infoCell.colSpan = headers.length;
        infoCell.textContent = "📋 Mostrando solo las primeras 5 filas como vista previa";
        infoCell.style.textAlign = "center";
        infoCell.style.fontStyle = "italic";
        infoCell.style.color = "#4a5568";
        infoCell.style.padding = "12px";
        infoCell.style.backgroundColor = "#edf2f7";
        infoCell.style.border = "1px solid #e2e8f0";
        infoRow.appendChild(infoCell);
        bodyEl.appendChild(infoRow);
    }
}