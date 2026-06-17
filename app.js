import * as pdfjsLib from "https://cdn.jsdelivr.net/npm/pdfjs-dist@6.0.227/build/pdf.mjs";

pdfjsLib.GlobalWorkerOptions.workerSrc = "https://cdn.jsdelivr.net/npm/pdfjs-dist@6.0.227/build/pdf.worker.mjs";

const WEEK_COLUMNS = ["MonA", "TueA", "WedA", "ThuA", "FriA", "MonB", "TueB", "WedB", "ThuB", "FriB"];
const HEADER_ROW = ["Student", ...WEEK_COLUMNS];
const A_DAYS = ["MonA", "TueA", "WedA", "ThuA", "FriA"];
const B_DAYS = ["MonB", "TueB", "WedB", "ThuB", "FriB"];
const DEFAULT_COLUMN_EDGES = [30, 115, 200, 285, 370, 455];
const DEFAULT_BANDS = {
  A: [40, 265],
  B: [280, 520],
};

const fileInput = document.querySelector("#pdf-input");
const dropzone = document.querySelector("#dropzone");
const controls = document.querySelector("#controls");
const processButton = document.querySelector("#process-button");
const clearButton = document.querySelector("#clear-button");
const subjectsInput = document.querySelector("#subjects-input");
const uploadIndicator = document.querySelector("#upload-indicator");
const uploadIndicatorTitle = document.querySelector("#upload-indicator-title");
const uploadIndicatorDetail = document.querySelector("#upload-indicator-detail");
const statusLine = document.querySelector("#status-line");
const resultTitle = document.querySelector("#result-title");
const tableWrap = document.querySelector("#table-wrap");
const resultsTable = document.querySelector("#results-table");
const xlsxButton = document.querySelector("#xlsx-button");
const csvButton = document.querySelector("#csv-button");

let selectedFiles = [];
let spreadsheetRows = [];

window.addEventListener("DOMContentLoaded", () => {
  updateUploadIndicator();
  if (window.lucide) {
    window.lucide.createIcons();
  }
});

fileInput.addEventListener("change", () => {
  setFiles(Array.from(fileInput.files || []));
});

dropzone.addEventListener("dragover", (event) => {
  event.preventDefault();
  dropzone.classList.add("dragging");
});

dropzone.addEventListener("dragleave", () => {
  dropzone.classList.remove("dragging");
});

dropzone.addEventListener("drop", (event) => {
  event.preventDefault();
  dropzone.classList.remove("dragging");
  const pdfs = Array.from(event.dataTransfer.files || []).filter(isPdfFile);
  setFiles(pdfs);
});

controls.addEventListener("submit", async (event) => {
  event.preventDefault();
  await processFiles();
});

clearButton.addEventListener("click", () => {
  selectedFiles = [];
  spreadsheetRows = [];
  fileInput.value = "";
  processButton.disabled = true;
  clearButton.disabled = true;
  xlsxButton.disabled = true;
  csvButton.disabled = true;
  tableWrap.hidden = true;
  resultsTable.innerHTML = "";
  resultTitle.textContent = "No timetable loaded yet";
  updateUploadIndicator();
  setStatus("Choose a PDF to begin.");
});

xlsxButton.addEventListener("click", () => {
  if (!spreadsheetRows.length || !window.XLSX) return;

  const worksheet = window.XLSX.utils.aoa_to_sheet(toSheetData(spreadsheetRows));
  worksheet["!cols"] = [{ wch: 24 }, ...WEEK_COLUMNS.map(() => ({ wch: 28 }))];
  const workbook = window.XLSX.utils.book_new();
  window.XLSX.utils.book_append_sheet(workbook, worksheet, "PE Uniform Days");
  window.XLSX.writeFile(workbook, "pe-uniform-days.xlsx");
});

csvButton.addEventListener("click", () => {
  if (!spreadsheetRows.length) return;
  downloadBlob(toCsv(toSheetData(spreadsheetRows)), "pe-uniform-days.csv", "text/csv;charset=utf-8");
});

function setFiles(files) {
  selectedFiles = files.filter(isPdfFile);
  processButton.disabled = selectedFiles.length === 0;
  clearButton.disabled = selectedFiles.length === 0 && spreadsheetRows.length === 0;
  updateUploadIndicator();

  if (!selectedFiles.length) {
    setStatus("Choose at least one PDF timetable.");
    return;
  }

  const fileWord = selectedFiles.length === 1 ? "file" : "files";
  setStatus(`${selectedFiles.length} PDF ${fileWord} ready.`);
}

function updateUploadIndicator() {
  const hasFiles = selectedFiles.length > 0;
  uploadIndicator.classList.toggle("uploaded", hasFiles);

  if (!hasFiles) {
    uploadIndicator.querySelector(".upload-indicator-icon").setAttribute("data-lucide", "circle-dashed");
    uploadIndicatorTitle.textContent = "No PDF uploaded";
    uploadIndicatorDetail.textContent = "Choose a timetable PDF to show it here.";
  } else {
    const fileWord = selectedFiles.length === 1 ? "PDF uploaded" : "PDFs uploaded";
    uploadIndicator.querySelector(".upload-indicator-icon").setAttribute("data-lucide", "circle-check");
    uploadIndicatorTitle.textContent = `${selectedFiles.length} ${fileWord}`;
    uploadIndicatorDetail.textContent = summarizeFileNames(selectedFiles);
  }

  if (window.lucide) {
    window.lucide.createIcons();
  }
}

function summarizeFileNames(files) {
  const visibleNames = files.slice(0, 3).map((file) => file.name);
  const hiddenCount = files.length - visibleNames.length;

  if (hiddenCount > 0) {
    visibleNames.push(`+${hiddenCount} more`);
  }

  return visibleNames.join(", ");
}

async function processFiles() {
  const subjects = getSubjects();

  if (!selectedFiles.length) {
    setStatus("Choose at least one PDF timetable.", true);
    return;
  }

  if (!subjects.length) {
    setStatus("Add at least one class name to search for.", true);
    return;
  }

  processButton.disabled = true;
  processButton.textContent = "Reading PDFs...";
  setStatus("Reading timetable pages in your browser...");

  try {
    const allRows = [];

    for (const file of selectedFiles) {
      const rows = await parsePdfFile(file, subjects);
      allRows.push(...rows);
    }

    spreadsheetRows = allRows.sort((a, b) => a.Student.localeCompare(b.Student));
    renderTable(spreadsheetRows);
    const studentWord = spreadsheetRows.length === 1 ? "student" : "students";
    resultTitle.textContent = `${spreadsheetRows.length} ${studentWord} found`;
    setStatus(`Ready. Download the spreadsheet as XLSX or CSV.`);
    xlsxButton.disabled = spreadsheetRows.length === 0 || !window.XLSX;
    csvButton.disabled = spreadsheetRows.length === 0;
    clearButton.disabled = false;
  } catch (error) {
    console.error(error);
    setStatus(error.message || "The PDF could not be read.", true);
  } finally {
    processButton.innerHTML = '<span data-lucide="sparkles"></span>Make spreadsheet';
    processButton.disabled = selectedFiles.length === 0;
    if (window.lucide) {
      window.lucide.createIcons();
    }
  }
}

async function parsePdfFile(file, subjects) {
  const bytes = await file.arrayBuffer();
  const loadingTask = pdfjsLib.getDocument({ data: bytes });
  const pdf = await loadingTask.promise;
  const rows = [];

  for (let pageNumber = 1; pageNumber <= pdf.numPages; pageNumber += 1) {
    const page = await pdf.getPage(pageNumber);
    const textContent = await page.getTextContent();
    const viewport = page.getViewport({ scale: 1 });
    const items = textContent.items
      .filter((item) => item.str && item.str.trim())
      .map((item) => normalizeTextItem(item, viewport.height));

    const pageLines = groupItemsIntoLines(items);
    const studentName = extractStudentName(pageLines, pageNumber);
    const geometry = detectTimetableGeometry(items);
    const row = buildStudentRow(studentName, items, geometry, subjects);
    rows.push(row);
  }

  return rows;
}

function normalizeTextItem(item, pageHeight) {
  const transform = item.transform || [1, 0, 0, 1, 0, 0];
  const x = transform[4] || 0;
  const rawY = transform[5] || 0;
  const height = Math.abs(transform[3] || item.height || 0);
  const top = pageHeight - rawY - height;

  return {
    text: item.str.trim(),
    x,
    top,
    width: item.width || 0,
  };
}

function groupItemsIntoLines(items) {
  const sorted = [...items].sort((a, b) => a.top - b.top || a.x - b.x);
  const lines = [];

  for (const item of sorted) {
    let line = lines.find((candidate) => Math.abs(candidate.top - item.top) < 3);

    if (!line) {
      line = { top: item.top, items: [] };
      lines.push(line);
    }

    line.items.push(item);
    line.top = (line.top * (line.items.length - 1) + item.top) / line.items.length;
  }

  return lines.map((line) => ({
    top: line.top,
    text: line.items.sort((a, b) => a.x - b.x).map((item) => item.text).join(" ").replace(/\s+/g, " ").trim(),
  }));
}

function extractStudentName(lines, pageNumber) {
  const titleLine = lines.find((line) => /^Timetable\s+For\s+/i.test(line.text));
  const name = titleLine?.text.replace(/^Timetable\s+For\s+/i, "").trim();
  return name || `Student ${pageNumber}`;
}

function detectTimetableGeometry(items) {
  const headers = new Map();

  for (const item of items) {
    if (WEEK_COLUMNS.includes(item.text)) {
      headers.set(item.text, item);
    }
  }

  const headerXs = A_DAYS.map((day) => headers.get(day)?.x).filter(isFiniteNumber);
  const bHeaderXs = B_DAYS.map((day) => headers.get(day)?.x).filter(isFiniteNumber);
  const allHeaderXs = headerXs.length === 5 ? headerXs : bHeaderXs;
  const columnEdges = allHeaderXs.length === 5 ? centersToEdges(allHeaderXs) : DEFAULT_COLUMN_EDGES;

  const monA = headers.get("MonA");
  const monB = headers.get("MonB");
  const friB = headers.get("FriB");
  const bandAStart = monA ? monA.top + 8 : DEFAULT_BANDS.A[0];
  const bandBStart = monB ? monB.top + 8 : DEFAULT_BANDS.B[0];
  const bandAEnd = monB ? monB.top - 8 : DEFAULT_BANDS.A[1];
  const bandBEnd = friB ? Math.min(friB.top + 250, 780) : DEFAULT_BANDS.B[1];

  return {
    columnEdges,
    bands: {
      A: [bandAStart, bandAEnd],
      B: [bandBStart, bandBEnd],
    },
  };
}

function centersToEdges(centers) {
  const sorted = [...centers].sort((a, b) => a - b);
  const edges = [];
  const averageGap = average(sorted.slice(1).map((value, index) => value - sorted[index])) || 85;

  edges.push(sorted[0] - averageGap * 0.63);

  for (let i = 0; i < sorted.length - 1; i += 1) {
    edges.push((sorted[i] + sorted[i + 1]) / 2);
  }

  edges.push(sorted[sorted.length - 1] + averageGap * 0.63);
  return edges;
}

function buildStudentRow(studentName, items, geometry, subjects) {
  const row = { Student: studentName };
  const mode = getCellMode();

  for (const day of WEEK_COLUMNS) {
    const text = getDayText(day, items, geometry);
    const matches = findSubjectMatches(text, subjects);
    row[day] = formatCell(matches, mode);
  }

  return row;
}

function getDayText(day, items, geometry) {
  const week = day.endsWith("A") ? "A" : "B";
  const dayIndex = (week === "A" ? A_DAYS : B_DAYS).indexOf(day);
  const [x0, x1] = [geometry.columnEdges[dayIndex], geometry.columnEdges[dayIndex + 1]];
  const [y0, y1] = geometry.bands[week];

  return items
    .filter((item) => item.x >= x0 && item.x < x1 && item.top >= y0 && item.top < y1)
    .sort((a, b) => a.top - b.top || a.x - b.x)
    .map((item) => item.text)
    .join(" ")
    .replace(/\s+/g, " ")
    .trim();
}

function findSubjectMatches(text, subjects) {
  const normalizedText = normalizeForMatching(text);
  const matches = [];

  for (const subject of subjects) {
    if (normalizedText.includes(normalizeForMatching(subject))) {
      matches.push(subject);
    }
  }

  return [...new Set(matches)];
}

function formatCell(matches, mode) {
  if (!matches.length) {
    return "No";
  }

  const subjectList = matches.join("; ");

  if (mode === "yes-no") {
    return "Yes";
  }

  if (mode === "subjects") {
    return subjectList;
  }

  return `Yes - ${subjectList}`;
}

function getSubjects() {
  return subjectsInput.value
    .split(/\r?\n|,/)
    .map((subject) => subject.trim())
    .filter(Boolean);
}

function getCellMode() {
  return document.querySelector('input[name="cell-mode"]:checked')?.value || "yes-subjects";
}

function renderTable(rows) {
  tableWrap.hidden = rows.length === 0;
  resultsTable.innerHTML = "";

  if (!rows.length) {
    return;
  }

  const thead = document.createElement("thead");
  const headRow = document.createElement("tr");

  for (const heading of HEADER_ROW) {
    const th = document.createElement("th");
    th.textContent = heading;
    headRow.appendChild(th);
  }

  thead.appendChild(headRow);
  resultsTable.appendChild(thead);

  const tbody = document.createElement("tbody");

  for (const row of rows) {
    const tr = document.createElement("tr");

    for (const heading of HEADER_ROW) {
      const td = document.createElement("td");
      td.textContent = row[heading] || "";

      if (heading !== "Student") {
        td.className = /^No$/i.test(td.textContent) ? "no-class" : "has-class";
      }

      tr.appendChild(td);
    }

    tbody.appendChild(tr);
  }

  resultsTable.appendChild(tbody);
}

function toSheetData(rows) {
  return [HEADER_ROW, ...rows.map((row) => HEADER_ROW.map((heading) => row[heading] || ""))];
}

function toCsv(data) {
  return data
    .map((row) => row.map((cell) => `"${String(cell).replace(/"/g, '""')}"`).join(","))
    .join("\r\n");
}

function downloadBlob(content, filename, type) {
  const blob = new Blob([content], { type });
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = filename;
  document.body.appendChild(anchor);
  anchor.click();
  anchor.remove();
  URL.revokeObjectURL(url);
}

function isPdfFile(file) {
  return file && (file.type === "application/pdf" || /\.pdf$/i.test(file.name));
}

function setStatus(message, isError = false) {
  statusLine.textContent = message;
  statusLine.classList.toggle("error", isError);
}

function normalizeForMatching(value) {
  return String(value)
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function average(values) {
  if (!values.length) return 0;
  return values.reduce((sum, value) => sum + value, 0) / values.length;
}

function isFiniteNumber(value) {
  return Number.isFinite(value);
}
