/**
 * ==========================================================================
 * Convertidor de PDF a Word (.docx) Editable - Versión Masiva
 * ==========================================================================
 * Permite la conversión de uno o múltiples archivos PDF (10, 20 o más) a
 * documentos Microsoft Word (.docx) editables directamente en el navegador.
 */

// Configuración del worker de PDF.js
if (typeof pdfjsLib !== 'undefined') {
  pdfjsLib.GlobalWorkerOptions.workerSrc =
    'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
}

// Estado de la aplicación
const AppState = {
  queue: [], // Lista de objetos { id, file, name, sizeFormatted, numPages, status, progress, statusText, pdfDoc, extractedPages, editedText, docxBlob, errorMsg }
  isConverting: false,
  activePreviewId: null,
  currentModalPage: 1,
  currentModalZoom: 1.0,
  currentModalPdfDoc: null,
  currentRenderTask: null
};

// Elementos del DOM
const DOM = {
  dropzone: document.getElementById('dropzone'),
  fileInput: document.getElementById('fileInput'),
  browseBtn: document.getElementById('browseBtn'),
  batchControlsSection: document.getElementById('batchControlsSection'),
  fileQueueSection: document.getElementById('fileQueueSection'),
  fileGrid: document.getElementById('fileGrid'),
  queueCount: document.getElementById('queueCount'),
  statTotal: document.getElementById('statTotal'),
  statPending: document.getElementById('statPending'),
  statCompleted: document.getElementById('statCompleted'),
  addMoreBtn: document.getElementById('addMoreBtn'),
  convertAllBtn: document.getElementById('convertAllBtn'),
  downloadZipBtn: document.getElementById('downloadZipBtn'),
  clearAllBtn: document.getElementById('clearAllBtn'),
  globalProgressContainer: document.getElementById('globalProgressContainer'),
  globalProgressText: document.getElementById('globalProgressText'),
  globalProgressPct: document.getElementById('globalProgressPct'),
  globalProgressBar: document.getElementById('globalProgressBar'),
  toastContainer: document.getElementById('toastContainer'),

  // Modal / Editor
  previewModal: document.getElementById('previewModal'),
  modalFileName: document.getElementById('modalFileName'),
  modalFileDetails: document.getElementById('modalFileDetails'),
  modalCloseBtn: document.getElementById('modalCloseBtn'),
  modalCloseBtn2: document.getElementById('modalCloseBtn2'),
  modalDownloadDocxBtn: document.getElementById('modalDownloadDocxBtn'),
  modalSaveAndDownloadBtn: document.getElementById('modalSaveAndDownloadBtn'),
  pdfCanvas: document.getElementById('pdfCanvas'),
  canvasLoading: document.getElementById('canvasLoading'),
  canvasViewport: document.getElementById('canvasViewport'),
  prevPageBtn: document.getElementById('prevPageBtn'),
  nextPageBtn: document.getElementById('nextPageBtn'),
  pageIndicator: document.getElementById('pageIndicator'),
  zoomInBtn: document.getElementById('zoomInBtn'),
  zoomOutBtn: document.getElementById('zoomOutBtn'),
  zoomLevel: document.getElementById('zoomLevel'),
  extractedTextEditor: document.getElementById('extractedTextEditor'),
  copyTextBtn: document.getElementById('copyTextBtn'),
  refreshExtractedBtn: document.getElementById('refreshExtractedBtn')
};

// Inicialización de eventos al cargar la página
document.addEventListener('DOMContentLoaded', () => {
  setupUploadEvents();
  setupBatchEvents();
  setupModalEvents();
});

// ==========================================================================
// 1. GESTIÓN DE SUBIDA Y DRAG & DROP
// ==========================================================================
function setupUploadEvents() {
  DOM.browseBtn.addEventListener('click', (e) => {
    e.stopPropagation();
    DOM.fileInput.click();
  });

  DOM.dropzone.addEventListener('click', () => {
    DOM.fileInput.click();
  });

  DOM.fileInput.addEventListener('change', (e) => {
    if (e.target.files && e.target.files.length > 0) {
      handleFiles(Array.from(e.target.files));
      DOM.fileInput.value = ''; // Reset para permitir volver a seleccionar los mismos
    }
  });

  // Prevenir apertura de archivos al soltar fuera del dropzone
  window.addEventListener('dragover', (e) => e.preventDefault(), false);
  window.addEventListener('drop', (e) => e.preventDefault(), false);

  // Efectos visuales de drag & drop
  DOM.dropzone.addEventListener('dragover', (e) => {
    e.preventDefault();
    DOM.dropzone.classList.add('dragover');
  });

  DOM.dropzone.addEventListener('dragleave', (e) => {
    e.preventDefault();
    DOM.dropzone.classList.remove('dragover');
  });

  DOM.dropzone.addEventListener('drop', (e) => {
    e.preventDefault();
    DOM.dropzone.classList.remove('dragover');
    if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
      handleFiles(Array.from(e.dataTransfer.files));
    }
  });

  DOM.addMoreBtn.addEventListener('click', () => {
    DOM.fileInput.click();
  });
}

/**
 * Procesa la lista de archivos recibidos y los agrega a la cola
 */
async function handleFiles(files) {
  const pdfFiles = files.filter(f => f.type === 'application/pdf' || f.name.toLowerCase().endsWith('.pdf'));

  if (pdfFiles.length === 0) {
    showToast('Por favor selecciona archivos con formato PDF válido.', 'error');
    return;
  }

  const addedCount = pdfFiles.length;
  showToast(`Se agregaron ${addedCount} archivo(s) PDF a la cola.`, 'info');

  for (const file of pdfFiles) {
    const fileId = 'file_' + Date.now() + '_' + Math.random().toString(36).substring(2, 9);
    const item = {
      id: fileId,
      file: file,
      name: file.name,
      sizeFormatted: formatBytes(file.size),
      numPages: null,
      status: 'pending', // pending, processing, completed, error
      progress: 0,
      statusText: 'En espera',
      pdfDoc: null,
      extractedPages: [],
      editedText: null,
      docxBlob: null,
      errorMsg: null
    };

    AppState.queue.push(item);
    renderFileCard(item);

    // Cargar metadatos iniciales (número de páginas) en segundo plano
    inspectPdfMetadata(item);
  }

  updateStats();
  toggleViewSections();
}

/**
 * Lee el PDF para saber cuántas páginas tiene antes de convertir
 */
async function inspectPdfMetadata(item) {
  try {
    const arrayBuffer = await item.file.arrayBuffer();
    const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
    item.pdfDoc = pdf;
    item.numPages = pdf.numPages;
    updateFileCard(item);
  } catch (err) {
    console.warn(`No se pudieron precargar las páginas de ${item.name}:`, err);
  }
}

// ==========================================================================
// 2. CONVERSIÓN INDIVIDUAL Y EN MASA (BATCH)
// ==========================================================================
function setupBatchEvents() {
  DOM.convertAllBtn.addEventListener('click', () => {
    convertAllPending();
  });

  DOM.downloadZipBtn.addEventListener('click', () => {
    downloadAllAsZip();
  });

  DOM.clearAllBtn.addEventListener('click', () => {
    if (AppState.isConverting) {
      if (!confirm('Hay conversiones en curso. ¿Deseas detener y vaciar la cola?')) return;
    }
    AppState.queue = [];
    AppState.isConverting = false;
    DOM.fileGrid.innerHTML = '';
    updateStats();
    toggleViewSections();
    showToast('Lista de archivos limpiada.', 'info');
  });
}

/**
 * Convierte todos los archivos en espera de forma continua
 */
async function convertAllPending() {
  const pendingFiles = AppState.queue.filter(f => f.status === 'pending' || f.status === 'error');
  if (pendingFiles.length === 0) {
    showToast('No hay archivos pendientes por convertir.', 'info');
    return;
  }

  AppState.isConverting = true;
  DOM.convertAllBtn.disabled = true;
  DOM.globalProgressContainer.style.display = 'flex';
  updateStats();

  let completedSoFar = 0;
  const totalToConvert = pendingFiles.length;

  for (let i = 0; i < pendingFiles.length; i++) {
    const item = pendingFiles[i];
    updateGlobalProgress(
      `Convirtiendo archivo ${i + 1} de ${totalToConvert}: ${item.name}`,
      Math.round((completedSoFar / totalToConvert) * 100)
    );

    await convertSinglePdf(item);
    completedSoFar++;

    updateGlobalProgress(
      `Completado ${completedSoFar} de ${totalToConvert}`,
      Math.round((completedSoFar / totalToConvert) * 100)
    );
  }

  AppState.isConverting = false;
  DOM.convertAllBtn.disabled = false;
  DOM.globalProgressContainer.style.display = 'none';
  updateStats();

  showToast(`¡Lote completado! Se convirtieron ${completedSoFar} archivos con éxito.`, 'success');
}

/**
 * Procesa y convierte un PDF individual a Word (.docx)
 */
async function convertSinglePdf(item) {
  item.status = 'processing';
  item.progress = 5;
  item.statusText = 'Iniciando lectura...';
  updateFileCard(item);

  try {
    // 1. Cargar documento PDF si no está en memoria
    if (!item.pdfDoc) {
      const arrayBuffer = await item.file.arrayBuffer();
      item.pdfDoc = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
      item.numPages = item.pdfDoc.numPages;
    }

    const numPages = item.pdfDoc.numPages;
    item.extractedPages = [];

    // 2. Extraer capa de texto estructurada página a página
    for (let p = 1; p <= numPages; p++) {
      item.statusText = `Extrayendo pág ${p} de ${numPages}...`;
      item.progress = Math.min(85, Math.round((p / numPages) * 75) + 5);
      updateFileCard(item);

      const page = await item.pdfDoc.getPage(p);
      const textContent = await page.getTextContent();
      const pageData = processPageTextContent(textContent, p);
      item.extractedPages.push(pageData);
    }

    // 3. Generar el documento Microsoft Word (.docx) real
    item.statusText = 'Generando Word (.docx)...';
    item.progress = 90;
    updateFileCard(item);

    const docxBlob = await buildDocxFromExtracted(item);
    item.docxBlob = docxBlob;
    item.status = 'completed';
    item.progress = 100;
    item.statusText = 'Listo para descargar';
    updateFileCard(item);
    updateStats();

  } catch (error) {
    console.error(`Error al convertir ${item.name}:`, error);
    item.status = 'error';
    item.errorMsg = error.message || 'Error al procesar el archivo PDF';
    item.statusText = 'Error en conversión';
    updateFileCard(item);
    updateStats();
  }
}

/**
 * Analiza los items de texto devueltos por PDF.js:
 * Agrupa por coordenadas Y (líneas) y coordenadas X (palabras),
 * detectando tamaños de fuente para clasificar títulos y párrafos normales.
 */
function processPageTextContent(textContent, pageNumber) {
  const items = textContent.items;
  if (!items || items.length === 0) {
    return {
      pageNumber,
      lines: [],
      rawText: ''
    };
  }

  // Filtrar elementos vacíos
  const validItems = items.filter(it => it.str && it.str.trim().length > 0);
  if (validItems.length === 0) {
    return { pageNumber, lines: [], rawText: '' };
  }

  // Agrupar items en la misma línea usando la coordenada Y (transform[5]) con un margen de tolerancia
  const linesMap = [];
  const Y_TOLERANCE = 4.0; // Puntos de tolerancia para la misma línea base

  for (const it of validItems) {
    const x = it.transform[4];
    const y = it.transform[5];
    const fontSize = Math.abs(it.transform[0]) || Math.abs(it.transform[3]) || it.height || 11;
    const fontName = (it.fontName || '').toLowerCase();
    const isBold = fontName.includes('bold') || fontName.includes('black') || fontName.includes('heavy');
    const isItalic = fontName.includes('italic') || fontName.includes('oblique');

    // Buscar si ya existe una línea cercana en Y
    let foundLine = linesMap.find(line => Math.abs(line.y - y) <= Y_TOLERANCE);
    if (!foundLine) {
      foundLine = {
        y: y,
        fontSize: fontSize,
        items: []
      };
      linesMap.push(foundLine);
    }

    foundLine.items.push({
      text: it.str,
      x: x,
      fontSize: fontSize,
      isBold: isBold,
      isItalic: isItalic
    });
  }

  // Ordenar las líneas de arriba hacia abajo (Y descendente en el sistema de coordenadas PDF)
  linesMap.sort((a, b) => b.y - a.y);

  // Calcular tamaño de fuente promedio de la página para detectar encabezados
  const allFontSizes = validItems.map(it => Math.abs(it.transform[0]) || 11);
  const avgFontSize = allFontSizes.reduce((a, b) => a + b, 0) / allFontSizes.length;

  const processedLines = [];
  const rawTextLines = [];

  for (const line of linesMap) {
    // Ordenar items de izquierda a derecha (X ascendente)
    line.items.sort((a, b) => a.x - b.x);

    // Concatenar texto respetando espacios
    let lineString = '';
    let maxFontSize = 0;
    let hasBold = false;
    let hasItalic = false;

    for (let i = 0; i < line.items.length; i++) {
      const it = line.items[i];
      if (it.fontSize > maxFontSize) maxFontSize = it.fontSize;
      if (it.isBold) hasBold = true;
      if (it.isItalic) hasItalic = true;

      if (i > 0) {
        const prev = line.items[i - 1];
        // Si hay separación horizontal mayor a 3pt o el texto no tiene espacio al inicio
        if (it.x - (prev.x + (prev.text.length * 4)) > 2 && !it.text.startsWith(' ')) {
          lineString += ' ';
        }
      }
      lineString += it.text;
    }

    lineString = lineString.trim();
    if (lineString.length > 0) {
      // Detección de títulos / encabezados según el tamaño de fuente relativo
      let isHeading = false;
      let headingLevel = 0;

      if (maxFontSize >= avgFontSize * 1.55 || (maxFontSize >= 16 && hasBold)) {
        isHeading = true;
        headingLevel = 1;
      } else if (maxFontSize >= avgFontSize * 1.25 || (maxFontSize >= 13 && hasBold)) {
        isHeading = true;
        headingLevel = 2;
      }

      processedLines.push({
        text: lineString,
        fontSize: Math.round(maxFontSize),
        isBold: hasBold,
        isItalic: hasItalic,
        isHeading: isHeading,
        headingLevel: headingLevel
      });

      rawTextLines.push(lineString);
    }
  }

  return {
    pageNumber,
    lines: processedLines,
    rawText: rawTextLines.join('\n')
  };
}

/**
 * Construye un documento DOCX nativo usando la librería docx.js
 */
async function buildDocxFromExtracted(item) {
  if (typeof docx === 'undefined') {
    throw new Error('La librería docx.js no está disponible en este momento.');
  }

  const { Document, Paragraph, TextRun, HeadingLevel } = docx;

  const docChildren = [];

  // Si el usuario editó el texto en el modal, usar el texto personalizado
  if (item.editedText !== null) {
    const paragraphs = item.editedText.split(/\n\s*\n/); // Dividir por líneas vacías
    for (let i = 0; i < paragraphs.length; i++) {
      const pText = paragraphs[i].trim();
      if (pText.length > 0) {
        // Chequear si es un salto de página manual marcado como [PÁGINA X]
        if (pText.startsWith('--- [PÁGINA') && i > 0) {
          docChildren.push(new Paragraph({
            children: [new TextRun({ text: '' })],
            pageBreakBefore: true
          }));
          continue;
        }

        docChildren.push(new Paragraph({
          children: [
            new TextRun({
              text: pText,
              font: 'Calibri',
              size: 24 // 12pt (docx usa medios puntos: 24 = 12pt)
            })
          ],
          spacing: { after: 200, line: 276 }
        }));
      }
    }
  } else {
    // Usar la extracción estructurada por páginas
    for (let pIdx = 0; pIdx < item.extractedPages.length; pIdx++) {
      const page = item.extractedPages[pIdx];

      // Salto de página entre hojas del documento original
      if (pIdx > 0) {
        docChildren.push(new Paragraph({
          children: [new TextRun({ text: '' })],
          pageBreakBefore: true
        }));
      }

      if (page.lines.length === 0) {
        docChildren.push(new Paragraph({
          children: [new TextRun({ text: '', font: 'Calibri', size: 24 })]
        }));
        continue;
      }

      // Agrupar líneas en párrafos continuos
      let currentParagraphRuns = [];
      let isCurrentHeading = false;
      let currentHeadingLvl = 0;

      for (let lIdx = 0; lIdx < page.lines.length; lIdx++) {
        const line = page.lines[lIdx];

        if (line.isHeading) {
          // Si había un párrafo acumulándose, vaciarlo
          if (currentParagraphRuns.length > 0) {
            docChildren.push(new Paragraph({
              children: currentParagraphRuns,
              spacing: { after: 160, line: 276 }
            }));
            currentParagraphRuns = [];
          }

          // Crear párrafo de encabezado
          docChildren.push(new Paragraph({
            heading: line.headingLevel === 1 ? HeadingLevel.HEADING_1 : HeadingLevel.HEADING_2,
            children: [
              new TextRun({
                text: line.text,
                font: 'Calibri',
                bold: true,
                size: line.headingLevel === 1 ? 32 : 28
              })
            ],
            spacing: { before: 240, after: 120 }
          }));
        } else {
          // Línea de texto normal
          currentParagraphRuns.push(new TextRun({
            text: (currentParagraphRuns.length > 0 ? ' ' : '') + line.text,
            font: 'Calibri',
            bold: line.isBold,
            italics: line.isItalic,
            size: 24 // 12pt
          }));

          // Si la línea termina con punto final o dos puntos, cerrar el párrafo para un formato limpio
          if (line.text.endsWith('.') || line.text.endsWith(':') || lIdx === page.lines.length - 1) {
            docChildren.push(new Paragraph({
              children: currentParagraphRuns,
              spacing: { after: 180, line: 276 }
            }));
            currentParagraphRuns = [];
          }
        }
      }

      if (currentParagraphRuns.length > 0) {
        docChildren.push(new Paragraph({
          children: currentParagraphRuns,
          spacing: { after: 180, line: 276 }
        }));
      }
    }
  }

  // Si no se extrajo texto (ej. PDF escaneado sin OCR)
  if (docChildren.length === 0) {
    docChildren.push(new Paragraph({
      children: [
        new TextRun({
          text: 'Nota: Este documento PDF parece contener únicamente imágenes escaneadas sin capa de texto seleccionable.',
          italics: true,
          font: 'Calibri',
          size: 22
        })
      ]
    }));
  }

  // Crear la estructura de secciones del documento DOCX
  const doc = new Document({
    creator: 'Convertidor PDF a Word Web',
    title: item.name.replace(/\.pdf$/i, ''),
    description: 'Documento Word convertido desde PDF',
    sections: [
      {
        properties: {
          page: {
            margin: {
              top: 1440, // 1 pulgada (1440 twips)
              right: 1440,
              bottom: 1440,
              left: 1440
            }
          }
        },
        children: docChildren
      }
    ]
  });

  return await docx.Packer.toBlob(doc);
}

/**
 * Función robusta para descargar Blobs en el navegador con soporte nativo
 */
function downloadBlob(blob, filename) {
  if (typeof saveAs === 'function') {
    saveAs(blob, filename);
  } else {
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.style.display = 'none';
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    setTimeout(() => URL.revokeObjectURL(url), 5000);
  }
}

/**
 * Empaqueta todos los documentos Word completados en un archivo .ZIP y lo descarga
 */
async function downloadAllAsZip() {
  const completed = AppState.queue.filter(f => f.status === 'completed' && f.docxBlob);
  if (completed.length === 0) {
    showToast('Aún no hay archivos convertidos para descargar.', 'error');
    return;
  }

  if (typeof JSZip === 'undefined') {
    showToast('La librería de compresión JSZip no está disponible.', 'error');
    return;
  }

  showToast(`Comprimiendo ${completed.length} documentos Word en un archivo .ZIP...`, 'info');

  const zip = new JSZip();
  const folder = zip.folder('Documentos_Word_Convertidos');

  for (const item of completed) {
    const docxName = item.name.replace(/\.pdf$/i, '') + '.docx';
    const arrayBuffer = await item.docxBlob.arrayBuffer();
    folder.file(docxName, arrayBuffer, { binary: true });
  }

  const zipBlob = await zip.generateAsync({
    type: 'blob',
    compression: 'DEFLATE',
    compressionOptions: { level: 6 }
  });

  const dateStr = new Date().toISOString().slice(0, 10);
  downloadBlob(zipBlob, `Word_Convertidos_${dateStr}.zip`);
  showToast('¡Descarga de archivo .ZIP iniciada!', 'success');
}

/**
 * Descarga individual de un archivo .docx
 */
function downloadSingleDocx(item) {
  if (!item.docxBlob) {
    showToast('El archivo aún no ha sido convertido.', 'error');
    return;
  }
  const docxName = item.name.replace(/\.pdf$/i, '') + '.docx';
  downloadBlob(item.docxBlob, docxName);
  showToast(`Descargando "${docxName}"`, 'success');
}

// ==========================================================================
// 3. RENDERIZADO DE LA INTERFAZ Y TARJETAS
// ==========================================================================
function renderFileCard(item) {
  const card = document.createElement('div');
  card.className = 'file-card';
  card.id = `card_${item.id}`;

  card.innerHTML = `
    <div class="file-card-top">
      <div class="file-icon-box">
        <i class="fa-solid fa-file-pdf"></i>
      </div>
      <div class="file-meta">
        <div class="file-name" title="${item.name}">${escapeHtml(item.name)}</div>
        <div class="file-details">
          <span>${item.sizeFormatted}</span>
          <span>&bull;</span>
          <span id="pages_${item.id}">${item.numPages ? item.numPages + ' pág(s)' : 'Leyendo páginas...'}</span>
        </div>
      </div>
      <span class="file-status-badge status-pending" id="badge_${item.id}">
        ${item.statusText}
      </span>
    </div>

    <div class="card-progress-track">
      <div class="card-progress-bar" id="prog_${item.id}" style="width: ${item.progress}%;"></div>
    </div>

    <div class="file-card-actions">
      <button type="button" class="btn btn-outline btn-sm btn-view" id="btn_view_${item.id}" title="Ver original y editar texto">
        <i class="fa-solid fa-eye"></i> Ver / Editar
      </button>
      <button type="button" class="btn btn-primary btn-sm btn-download" id="btn_down_${item.id}" disabled title="Descargar archivo Word editable">
        <i class="fa-solid fa-file-word"></i> Descargar Word
      </button>
      <button type="button" class="btn btn-danger-soft btn-sm btn-icon-only btn-remove" id="btn_rem_${item.id}" title="Eliminar de la lista">
        <i class="fa-solid fa-trash"></i>
      </button>
    </div>
  `;

  // Asignar eventos de la tarjeta
  card.querySelector(`#btn_view_${item.id}`).addEventListener('click', () => openPreviewModal(item.id));
  card.querySelector(`#btn_down_${item.id}`).addEventListener('click', () => downloadSingleDocx(item));
  card.querySelector(`#btn_rem_${item.id}`).addEventListener('click', () => removeFileFromQueue(item.id));

  DOM.fileGrid.appendChild(card);
}

function updateFileCard(item) {
  const card = document.getElementById(`card_${item.id}`);
  if (!card) return;

  const pagesSpan = document.getElementById(`pages_${item.id}`);
  if (pagesSpan && item.numPages) {
    pagesSpan.textContent = `${item.numPages} pág(s)`;
  }

  const badge = document.getElementById(`badge_${item.id}`);
  if (badge) {
    badge.className = `file-status-badge status-${item.status}`;
    if (item.status === 'processing') {
      badge.innerHTML = `<i class="fa-solid fa-spinner fa-spin"></i> ${escapeHtml(item.statusText)}`;
    } else if (item.status === 'completed') {
      badge.innerHTML = `<i class="fa-solid fa-check"></i> Convertido`;
    } else if (item.status === 'error') {
      badge.innerHTML = `<i class="fa-solid fa-triangle-exclamation"></i> Error`;
    } else {
      badge.textContent = item.statusText;
    }
  }

  const progBar = document.getElementById(`prog_${item.id}`);
  if (progBar) {
    progBar.style.width = `${item.progress}%`;
  }

  const downBtn = document.getElementById(`btn_down_${item.id}`);
  if (downBtn) {
    downBtn.disabled = item.status !== 'completed' || !item.docxBlob;
    if (item.status === 'completed') {
      downBtn.classList.remove('btn-primary');
      downBtn.classList.add('btn-success');
    }
  }

  if (item.status === 'completed') {
    card.classList.add('completed');
  } else {
    card.classList.remove('completed');
  }
}

function removeFileFromQueue(fileId) {
  const idx = AppState.queue.findIndex(f => f.id === fileId);
  if (idx !== -1) {
    AppState.queue.splice(idx, 1);
    const el = document.getElementById(`card_${fileId}`);
    if (el) el.remove();
    updateStats();
    toggleViewSections();
  }
}

function updateStats() {
  const total = AppState.queue.length;
  const pending = AppState.queue.filter(f => f.status === 'pending' || f.status === 'processing').length;
  const completed = AppState.queue.filter(f => f.status === 'completed').length;

  DOM.statTotal.textContent = total;
  DOM.statPending.textContent = pending;
  DOM.statCompleted.textContent = completed;
  DOM.queueCount.textContent = total;

  DOM.downloadZipBtn.disabled = completed === 0;
  DOM.convertAllBtn.disabled = AppState.isConverting || pending === 0;
}

function updateGlobalProgress(text, pct) {
  DOM.globalProgressText.textContent = text;
  DOM.globalProgressPct.textContent = `${pct}%`;
  DOM.globalProgressBar.style.width = `${pct}%`;
}

function toggleViewSections() {
  const hasFiles = AppState.queue.length > 0;
  DOM.batchControlsSection.style.display = hasFiles ? 'flex' : 'none';
  DOM.fileQueueSection.style.display = hasFiles ? 'flex' : 'none';
}

// ==========================================================================
// 4. MODAL / VISOR & EDITOR INTERACTIVO
// ==========================================================================
function setupModalEvents() {
  DOM.modalCloseBtn.addEventListener('click', closePreviewModal);
  DOM.modalCloseBtn2.addEventListener('click', closePreviewModal);

  DOM.prevPageBtn.addEventListener('click', () => {
    if (AppState.currentModalPage > 1) {
      AppState.currentModalPage--;
      renderModalPdfPage(AppState.currentModalPage);
    }
  });

  DOM.nextPageBtn.addEventListener('click', () => {
    if (AppState.currentModalPdfDoc && AppState.currentModalPage < AppState.currentModalPdfDoc.numPages) {
      AppState.currentModalPage++;
      renderModalPdfPage(AppState.currentModalPage);
    }
  });

  DOM.zoomInBtn.addEventListener('click', () => {
    if (AppState.currentModalZoom < 2.5) {
      AppState.currentModalZoom += 0.25;
      DOM.zoomLevel.textContent = `${Math.round(AppState.currentModalZoom * 100)}%`;
      renderModalPdfPage(AppState.currentModalPage);
    }
  });

  DOM.zoomOutBtn.addEventListener('click', () => {
    if (AppState.currentModalZoom > 0.5) {
      AppState.currentModalZoom -= 0.25;
      DOM.zoomLevel.textContent = `${Math.round(AppState.currentModalZoom * 100)}%`;
      renderModalPdfPage(AppState.currentModalPage);
    }
  });

  DOM.copyTextBtn.addEventListener('click', async () => {
    const text = DOM.extractedTextEditor.value;
    if (!text) return;
    try {
      await navigator.clipboard.writeText(text);
      showToast('¡Texto copiado al portapapeles!', 'success');
    } catch {
      DOM.extractedTextEditor.select();
      document.execCommand('copy');
      showToast('¡Texto copiado!', 'success');
    }
  });

  DOM.refreshExtractedBtn.addEventListener('click', () => {
    const item = AppState.queue.find(f => f.id === AppState.activePreviewId);
    if (!item) return;
    item.editedText = null;
    fillEditorWithExtracted(item);
    showToast('Texto restablecido a la extracción original.', 'info');
  });

  DOM.modalSaveAndDownloadBtn.addEventListener('click', async () => {
    await saveEditedAndDownload();
  });

  DOM.modalDownloadDocxBtn.addEventListener('click', async () => {
    await saveEditedAndDownload();
  });
}

async function openPreviewModal(fileId) {
  const item = AppState.queue.find(f => f.id === fileId);
  if (!item) return;

  AppState.activePreviewId = fileId;
  DOM.modalFileName.textContent = item.name;
  DOM.modalFileDetails.textContent = `${item.sizeFormatted} • ${item.numPages ? item.numPages + ' página(s)' : 'Cargando...'}`;

  DOM.previewModal.style.display = 'flex';
  AppState.currentModalPage = 1;
  AppState.currentModalZoom = 1.0;
  DOM.zoomLevel.textContent = '100%';

  // Cargar PDF en el Canvas
  if (!item.pdfDoc) {
    DOM.canvasLoading.style.display = 'flex';
    const arrayBuffer = await item.file.arrayBuffer();
    item.pdfDoc = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
    item.numPages = item.pdfDoc.numPages;
  }

  AppState.currentModalPdfDoc = item.pdfDoc;
  renderModalPdfPage(AppState.currentModalPage);

  // Si no se ha extraído texto aún, extraerlo ahora para este archivo
  if (item.extractedPages.length === 0) {
    DOM.extractedTextEditor.value = 'Extrayendo texto del documento PDF, por favor espera...';
    DOM.extractedTextEditor.disabled = true;

    for (let p = 1; p <= item.pdfDoc.numPages; p++) {
      const page = await item.pdfDoc.getPage(p);
      const textContent = await page.getTextContent();
      const pageData = processPageTextContent(textContent, p);
      item.extractedPages.push(pageData);
    }
    DOM.extractedTextEditor.disabled = false;
  }

  fillEditorWithExtracted(item);
}

function fillEditorWithExtracted(item) {
  if (item.editedText !== null) {
    DOM.extractedTextEditor.value = item.editedText;
  } else {
    const fullText = item.extractedPages.map((page, idx) => {
      const header = item.extractedPages.length > 1 ? `--- [PÁGINA ${idx + 1}] ---\n\n` : '';
      return header + page.rawText;
    }).join('\n\n');

    DOM.extractedTextEditor.value = fullText || '(No se detectó capa de texto seleccionable en el PDF)';
  }
}

async function renderModalPdfPage(pageNum) {
  if (!AppState.currentModalPdfDoc) return;

  DOM.canvasLoading.style.display = 'flex';
  DOM.pageIndicator.textContent = `${pageNum} / ${AppState.currentModalPdfDoc.numPages}`;

  try {
    const page = await AppState.currentModalPdfDoc.getPage(pageNum);
    const viewport = page.getViewport({ scale: AppState.currentModalZoom * 1.35 });

    const canvas = DOM.pdfCanvas;
    const context = canvas.getContext('2d');
    canvas.height = viewport.height;
    canvas.width = viewport.width;

    if (AppState.currentRenderTask) {
      try { AppState.currentRenderTask.cancel(); } catch { }
    }

    const renderContext = {
      canvasContext: context,
      viewport: viewport
    };

    AppState.currentRenderTask = page.render(renderContext);
    await AppState.currentRenderTask.promise;
    DOM.canvasLoading.style.display = 'none';
  } catch (err) {
    if (err.name !== 'RenderingCancelledException') {
      console.error('Error al renderizar canvas:', err);
    }
    DOM.canvasLoading.style.display = 'none';
  }
}

async function saveEditedAndDownload() {
  const item = AppState.queue.find(f => f.id === AppState.activePreviewId);
  if (!item) return;

  const currentEditedText = DOM.extractedTextEditor.value;
  item.editedText = currentEditedText;

  showToast('Generando documento Word con los cambios...', 'info');

  try {
    const docxBlob = await buildDocxFromExtracted(item);
    item.docxBlob = docxBlob;
    item.status = 'completed';
    item.statusText = 'Listo para descargar';
    updateFileCard(item);
    updateStats();

    downloadSingleDocx(item);
    closePreviewModal();
  } catch (err) {
    console.error('Error al regenerar DOCX:', err);
    showToast('Error al crear el documento Word.', 'error');
  }
}

function closePreviewModal() {
  DOM.previewModal.style.display = 'none';
  AppState.activePreviewId = null;
  AppState.currentModalPdfDoc = null;
  if (AppState.currentRenderTask) {
    try { AppState.currentRenderTask.cancel(); } catch { }
  }
}

// ==========================================================================
// 5. UTILIDADES Y NOTIFICACIONES
// ==========================================================================
function formatBytes(bytes) {
  if (bytes === 0) return '0 Bytes';
  const k = 1024;
  const sizes = ['Bytes', 'KB', 'MB', 'GB'];
  const i = Math.floor(Math.log(bytes) / Math.log(k));
  return parseFloat((bytes / Math.pow(k, i)).toFixed(1)) + ' ' + sizes[i];
}

function escapeHtml(text) {
  const div = document.createElement('div');
  div.textContent = text;
  return div.innerHTML;
}

function showToast(message, type = 'info') {
  const toast = document.createElement('div');
  toast.className = `toast toast-${type}`;

  const iconClass = type === 'success' ? 'fa-circle-check text-success' :
    type === 'error' ? 'fa-circle-exclamation text-danger' :
      'fa-circle-info text-primary';

  toast.innerHTML = `
    <i class="fa-solid ${iconClass}"></i>
    <span>${escapeHtml(message)}</span>
  `;

  DOM.toastContainer.appendChild(toast);

  setTimeout(() => {
    toast.style.opacity = '0';
    toast.style.transform = 'translateY(10px)';
    toast.style.transition = 'all 0.3s ease';
    setTimeout(() => toast.remove(), 300);
  }, 4000);
}
