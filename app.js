const MATRIX_ROWS = 1000;
const MATRIX_COLS = 20;

const dropzone = document.getElementById("dropzone");
const fileInput = document.getElementById("file-input");
const browseButton = document.getElementById("browse-button");
const clearButton = document.getElementById("clear-button");
const resetViewButton = document.getElementById("reset-view-button");
const downloadButton = document.getElementById("download-button");
const downloadWordSheetButton = document.getElementById("download-word-sheet-button");
const downloadWordRowButton = document.getElementById("download-word-row-button");
const wordRowInput = document.getElementById("word-row-input");
const sheetSelect = document.getElementById("sheet-select");
const statusText = document.getElementById("status-text");
const fileMeta = document.getElementById("file-meta");
const gridWrapper = document.getElementById("grid-wrapper");

const allowedExtensions = new Set(["xlsx", "xls", "csv"]);
let sheetsData = { Hoja1: createEmptyMatrix() };
let activeSheetName = "Hoja1";

function createEmptyMatrix() {
  return Array.from({ length: MATRIX_ROWS }, () => Array(MATRIX_COLS).fill(""));
}

function getActiveMatrix() {
  if (!sheetsData[activeSheetName]) {
    sheetsData[activeSheetName] = createEmptyMatrix();
  }
  return sheetsData[activeSheetName];
}

function formatBytes(bytes) {
  if (bytes < 1024) return `${bytes} B`;
  const kb = bytes / 1024;
  if (kb < 1024) return `${kb.toFixed(1)} KB`;
  return `${(kb / 1024).toFixed(2)} MB`;
}

function getExtension(fileName) {
  const parts = fileName.split(".");
  return parts.length > 1 ? parts.pop().toLowerCase() : "";
}

function setStatus(message, isError = false) {
  statusText.textContent = message;
  statusText.classList.toggle("error", isError);
}

function renderFileMeta(file, extraMessage = "") {
  fileMeta.classList.remove("hidden");
  fileMeta.innerHTML = "";

  const details = [
    `Nombre: ${file.name}`,
    `Tamano: ${formatBytes(file.size)}`,
    `Tipo: ${file.type || "No especificado"}`,
    `Hojas detectadas: ${Object.keys(sheetsData).length}`
  ];

  if (extraMessage) {
    details.push(extraMessage);
  }

  details.forEach((item) => {
    const li = document.createElement("li");
    li.textContent = item;
    fileMeta.appendChild(li);
  });
}

function hideFileMeta() {
  fileMeta.classList.add("hidden");
  fileMeta.innerHTML = "";
}

function normalizeToMatrix(rawData) {
  const normalized = createEmptyMatrix();

  for (let r = 0; r < Math.min(rawData.length, MATRIX_ROWS); r += 1) {
    const sourceRow = Array.isArray(rawData[r]) ? rawData[r] : [];
    for (let c = 0; c < Math.min(sourceRow.length, MATRIX_COLS); c += 1) {
      normalized[r][c] = sourceRow[c] == null ? "" : String(sourceRow[c]);
    }
  }

  return normalized;
}

function renderGrid() {
  const matrix = getActiveMatrix();
  const headerCells = Array.from({ length: MATRIX_COLS }, (_, index) => `<th scope="col">C${index + 1}</th>`).join("");

  const bodyRows = matrix
    .map((row, rowIndex) => {
      const cells = row
        .map((cellValue, colIndex) => (`<td><input class="cell-input" data-row="${rowIndex}" data-col="${colIndex}" value="${escapeHtml(cellValue)}" /></td>`))
        .join("");

      return `<tr><th class="row-header" scope="row">${rowIndex + 1}</th>${cells}</tr>`;
    })
    .join("");

  gridWrapper.innerHTML = `
    <table class="matrix-table">
      <thead>
        <tr>
          <th class="row-header corner" scope="col">#</th>
          ${headerCells}
        </tr>
      </thead>
      <tbody>${bodyRows}</tbody>
    </table>
  `;
}

function escapeHtml(value) {
  return value
    .replaceAll("&", "&amp;")
    .replaceAll("\"", "&quot;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;");
}

function setSheetOptions(sheetNames) {
  sheetSelect.innerHTML = "";
  sheetNames.forEach((sheetName) => {
    const option = document.createElement("option");
    option.value = sheetName;
    option.textContent = sheetName;
    sheetSelect.appendChild(option);
  });

  sheetSelect.disabled = sheetNames.length === 0;
  if (sheetNames.length > 0) {
    sheetSelect.value = activeSheetName;
  }
}

function parseWorkbookFile(file) {
  return new Promise((resolve, reject) => {
    if (!window.XLSX) {
      reject(new Error("No se pudo cargar la libreria XLSX."));
      return;
    }

    const reader = new FileReader();

    reader.onload = (event) => {
      try {
        const data = new Uint8Array(event.target.result);
        const workbook = XLSX.read(data, { type: "array" });
        const parsedSheets = {};

        workbook.SheetNames.forEach((sheetName) => {
          const sheet = workbook.Sheets[sheetName];
          const rawData = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
          parsedSheets[sheetName] = normalizeToMatrix(rawData);
        });

        if (workbook.SheetNames.length === 0) {
          parsedSheets.Hoja1 = createEmptyMatrix();
        }

        resolve({ parsedSheets, sheetNames: Object.keys(parsedSheets) });
      } catch (error) {
        reject(error);
      }
    };

    reader.onerror = () => reject(new Error("No se pudo leer el archivo."));
    reader.readAsArrayBuffer(file);
  });
}

async function processFile(file) {
  if (!file) return;

  const extension = getExtension(file.name);
  if (!allowedExtensions.has(extension)) {
    setStatus("Formato invalido. Usa .xlsx, .xls o .csv.", true);
    hideFileMeta();
    return;
  }

  setStatus("Procesando archivo...");

  try {
    const { parsedSheets, sheetNames } = await parseWorkbookFile(file);

    sheetsData = parsedSheets;
    activeSheetName = sheetNames[0];
    setSheetOptions(sheetNames);
    renderGrid();
    downloadButton.disabled = false;
    downloadWordSheetButton.disabled = false;
    downloadWordRowButton.disabled = false;

    setStatus(`Archivo cargado. Hoja activa: ${activeSheetName}.`);
    renderFileMeta(file, `Matriz por hoja normalizada a ${MATRIX_ROWS} x ${MATRIX_COLS}.`);
  } catch (error) {
    setStatus(`Error al cargar archivo: ${error.message}`, true);
    hideFileMeta();
  }
}

function downloadWorkbook() {
  if (!window.XLSX) {
    setStatus("No se pudo cargar la libreria XLSX para exportar.", true);
    return;
  }

  const workbook = XLSX.utils.book_new();
  Object.entries(sheetsData).forEach(([sheetName, matrix]) => {
    const worksheet = XLSX.utils.aoa_to_sheet(matrix);
    XLSX.utils.book_append_sheet(workbook, worksheet, sheetName.slice(0, 31) || "Hoja");
  });

  XLSX.writeFile(workbook, "cartera_seguros_editada.xlsx");
  setStatus("Descarga iniciada: cartera_seguros_editada.xlsx");
}

function sanitizeFilePart(value) {
  return String(value).replace(/[^a-zA-Z0-9_-]/g, "_");
}

function buildWordHtml(title, rows) {
  const header = Array.from({ length: MATRIX_COLS }, (_, idx) => `<th>C${idx + 1}</th>`).join("");
  const body = rows
    .map((row) => `<tr>${row.map((cell) => `<td>${escapeHtml(cell || "")}</td>`).join("")}</tr>`)
    .join("");

  return `<!doctype html>
<html>
<head>
  <meta charset="utf-8">
  <title>${escapeHtml(title)}</title>
  <style>
    body { font-family: Calibri, Arial, sans-serif; font-size: 11pt; }
    h1 { font-size: 14pt; margin-bottom: 12px; }
    table { border-collapse: collapse; width: 100%; }
    th, td { border: 1px solid #444; padding: 4px; vertical-align: top; }
    th { background: #efefef; }
  </style>
</head>
<body>
  <h1>${escapeHtml(title)}</h1>
  <table>
    <thead><tr>${header}</tr></thead>
    <tbody>${body}</tbody>
  </table>
</body>
</html>`;
}

function downloadWordDocument(fileName, html) {
  const blob = new Blob(["\ufeff", html], { type: "application/msword;charset=utf-8" });
  const link = document.createElement("a");
  link.href = URL.createObjectURL(blob);
  link.download = fileName;
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(link.href);
}

function getNonEmptyRows(matrix) {
  return matrix.filter((row) => row.some((cell) => String(cell || "").trim() !== ""));
}

function downloadActiveSheetAsWord() {
  const matrix = getActiveMatrix();
  const rows = getNonEmptyRows(matrix);

  if (rows.length === 0) {
    setStatus("La hoja activa no tiene datos para exportar a Word.", true);
    return;
  }

  const title = `Sura - Hoja ${activeSheetName}`;
  const html = buildWordHtml(title, rows);
  const safeSheet = sanitizeFilePart(activeSheetName);
  downloadWordDocument(`sura_${safeSheet}_hoja.doc`, html);
  setStatus(`Word generado para hoja ${activeSheetName}.`);
}

function downloadRowAsWord() {
  const rowIndex = Number(wordRowInput.value) - 1;
  if (!Number.isInteger(rowIndex) || rowIndex < 0 || rowIndex >= MATRIX_ROWS) {
    setStatus("Fila invalida. Usa un numero entre 1 y 1000.", true);
    return;
  }

  const matrix = getActiveMatrix();
  const row = matrix[rowIndex];
  const hasData = row.some((cell) => String(cell || "").trim() !== "");

  if (!hasData) {
    setStatus(`La fila ${rowIndex + 1} esta vacia en la hoja activa.`, true);
    return;
  }

  const title = `Sura - Hoja ${activeSheetName} - Fila ${rowIndex + 1}`;
  const html = buildWordHtml(title, [row]);
  const safeSheet = sanitizeFilePart(activeSheetName);
  downloadWordDocument(`sura_${safeSheet}_fila_${rowIndex + 1}.doc`, html);
  setStatus(`Word generado para fila ${rowIndex + 1} de ${activeSheetName}.`);
}

browseButton.addEventListener("click", () => fileInput.click());

downloadButton.addEventListener("click", downloadWorkbook);
downloadWordSheetButton.addEventListener("click", downloadActiveSheetAsWord);
downloadWordRowButton.addEventListener("click", downloadRowAsWord);
resetViewButton.addEventListener("click", () => {
  gridWrapper.scrollTo({ top: 0, left: 0, behavior: "smooth" });
});

sheetSelect.addEventListener("change", () => {
  activeSheetName = sheetSelect.value;
  renderGrid();
  setStatus(`Hoja activa: ${activeSheetName}.`);
});

clearButton.addEventListener("click", () => {
  sheetsData[activeSheetName] = createEmptyMatrix();
  renderGrid();
  setStatus(`Hoja ${activeSheetName} reiniciada a ${MATRIX_ROWS} x ${MATRIX_COLS}.`);
  downloadButton.disabled = false;
  downloadWordSheetButton.disabled = false;
  downloadWordRowButton.disabled = false;
});

fileInput.addEventListener("change", (event) => {
  const [file] = event.target.files;
  processFile(file);
});

dropzone.addEventListener("dragover", (event) => {
  event.preventDefault();
  dropzone.classList.add("drag-over");
});

dropzone.addEventListener("dragleave", () => {
  dropzone.classList.remove("drag-over");
});

dropzone.addEventListener("drop", (event) => {
  event.preventDefault();
  dropzone.classList.remove("drag-over");

  const [file] = event.dataTransfer.files;
  processFile(file);
});

dropzone.addEventListener("keydown", (event) => {
  if (event.key === "Enter" || event.key === " ") {
    event.preventDefault();
    fileInput.click();
  }
});

gridWrapper.addEventListener("input", (event) => {
  const target = event.target;
  if (!(target instanceof HTMLInputElement)) return;

  const row = Number(target.dataset.row);
  const col = Number(target.dataset.col);

  if (Number.isNaN(row) || Number.isNaN(col)) return;
  sheetsData[activeSheetName][row][col] = target.value;
});

setSheetOptions([activeSheetName]);
renderGrid();
setStatus(`Matriz vacia disponible (${MATRIX_ROWS} x ${MATRIX_COLS}).`);
