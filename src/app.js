const MATRIX_ROWS = 1000;
const MATRIX_COLS = 20;

const dropzone = document.getElementById("dropzone");
const fileInput = document.getElementById("file-input");
const browseButton = document.getElementById("browse-button");
const clearButton = document.getElementById("clear-button");
const resetViewButton = document.getElementById("reset-view-button");
const downloadButton = document.getElementById("download-button");
const statusText = document.getElementById("status-text");
const fileMeta = document.getElementById("file-meta");
const gridWrapper = document.getElementById("grid-wrapper");

const allowedExtensions = new Set(["xlsx", "xls", "csv"]);
let matrix = createEmptyMatrix();

function createEmptyMatrix() {
  return Array.from({ length: MATRIX_ROWS }, () => Array(MATRIX_COLS).fill(""));
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
    `Tamaño: ${formatBytes(file.size)}`,
    `Tipo: ${file.type || "No especificado"}`
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
  const headerCells = Array.from({ length: MATRIX_COLS }, (_, index) => `<th scope="col">C${index + 1}</th>`).join("");

  const bodyRows = matrix
    .map((row, rowIndex) => {
      const cells = row
        .map((cellValue, colIndex) => (
          `<td><input class="cell-input" data-row="${rowIndex}" data-col="${colIndex}" value="${escapeHtml(cellValue)}" /></td>`
        ))
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

function parseWorkbookFile(file) {
  return new Promise((resolve, reject) => {
    if (!window.XLSX) {
      reject(new Error("No se pudo cargar la librería XLSX."));
      return;
    }

    const reader = new FileReader();

    reader.onload = (event) => {
      try {
        const data = new Uint8Array(event.target.result);
        const workbook = XLSX.read(data, { type: "array" });
        const firstSheetName = workbook.SheetNames[0];
        const firstSheet = workbook.Sheets[firstSheetName];
        const rawData = XLSX.utils.sheet_to_json(firstSheet, { header: 1, defval: "" });

        resolve(rawData);
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
    setStatus("Formato inválido. Usa .xlsx, .xls o .csv.", true);
    hideFileMeta();
    return;
  }

  setStatus("Procesando archivo...");

  try {
    const rawData = await parseWorkbookFile(file);
    matrix = normalizeToMatrix(rawData);
    renderGrid();
    downloadButton.disabled = false;

    const trimmedRows = Math.max(0, rawData.length - MATRIX_ROWS);
    const maxColsInSource = rawData.reduce((max, row) => Math.max(max, Array.isArray(row) ? row.length : 0), 0);
    const trimmedCols = Math.max(0, maxColsInSource - MATRIX_COLS);

    const trimMessage =
      trimmedRows > 0 || trimmedCols > 0
        ? `Aviso: se recortó a ${MATRIX_ROWS} x ${MATRIX_COLS}. Filas recortadas: ${trimmedRows}, columnas recortadas: ${trimmedCols}.`
        : `Matriz normalizada a ${MATRIX_ROWS} x ${MATRIX_COLS}.`;

    setStatus("Archivo cargado y listo para editar.");
    renderFileMeta(file, trimMessage);
  } catch (error) {
    setStatus(`Error al cargar archivo: ${error.message}`, true);
    hideFileMeta();
  }
}

function downloadWorkbook() {
  if (!window.XLSX) {
    setStatus("No se pudo cargar la librería XLSX para exportar.", true);
    return;
  }

  const worksheet = XLSX.utils.aoa_to_sheet(matrix);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, "CarteraSeguros");
  XLSX.writeFile(workbook, "cartera_seguros_editada.xlsx");
  setStatus("Descarga iniciada: cartera_seguros_editada.xlsx");
}

browseButton.addEventListener("click", () => fileInput.click());

downloadButton.addEventListener("click", downloadWorkbook);
resetViewButton.addEventListener("click", () => {
  gridWrapper.scrollTo({ top: 0, left: 0, behavior: "smooth" });
});

clearButton.addEventListener("click", () => {
  matrix = createEmptyMatrix();
  renderGrid();
  hideFileMeta();
  downloadButton.disabled = false;
  setStatus(`Matriz reiniciada a ${MATRIX_ROWS} x ${MATRIX_COLS}.`);
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
  matrix[row][col] = target.value;
});

renderGrid();
setStatus(`Matriz vacía disponible (${MATRIX_ROWS} x ${MATRIX_COLS}).`);
