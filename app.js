"use strict";

const OUTPUT_HEADERS = [
  "Invoice Number",
  "Invoice Date",
  "Customer GSTIN",
  "Customer Name",
  "Customer State",
  "Place of Supply",
  "HSN/SAC",
  "Taxable Value",
  "GST Rate (%)",
  "CGST",
  "SGST",
  "IGST",
  "Cess",
  "Invoice Total"
];

const FIELDS = [
  { key: "invoiceNumber", label: "Invoice number", required: true, aliases: ["invoice number", "invoice no", "invoice #", "order id", "order number", "sale id", "transaction id"] },
  { key: "invoiceDate", label: "Invoice date", aliases: ["invoice date", "date", "order date", "sale date", "transaction date"] },
  { key: "customerGstin", label: "Customer GSTIN", aliases: ["customer gstin", "buyer gstin", "gstin", "gst number", "customer gst no"] },
  { key: "customerName", label: "Customer name", aliases: ["customer name", "buyer name", "customer", "billing name", "recipient"] },
  { key: "customerState", label: "Customer state", aliases: ["customer state", "buyer state", "billing state", "state", "ship state", "shipping state"] },
  { key: "placeOfSupply", label: "Place of supply", aliases: ["place of supply", "pos", "place of supply state", "supply state"] },
  { key: "hsnSac", label: "HSN / SAC", aliases: ["hsn", "hsn code", "hsn/sac", "sac", "product hsn"] },
  { key: "taxableValue", label: "Taxable value", required: true, aliases: ["taxable value", "taxable amount", "taxable", "subtotal", "net amount", "assessable value"] },
  { key: "gstRate", label: "GST rate (%)", aliases: ["gst rate", "tax rate", "gst %", "tax %", "rate", "gst percentage"] },
  { key: "cgst", label: "CGST", aliases: ["cgst", "cgst amount"] },
  { key: "sgst", label: "SGST", aliases: ["sgst", "sgst amount", "utgst", "utgst amount"] },
  { key: "igst", label: "IGST", aliases: ["igst", "igst amount"] },
  { key: "cess", label: "Cess", aliases: ["cess", "cess amount"] },
  { key: "invoiceTotal", label: "Invoice total", aliases: ["invoice total", "total amount", "grand total", "total", "order total", "gross amount"] }
];

const STATES = [
  ["01", "Jammu and Kashmir"], ["02", "Himachal Pradesh"], ["03", "Punjab"], ["04", "Chandigarh"],
  ["05", "Uttarakhand"], ["06", "Haryana"], ["07", "Delhi"], ["08", "Rajasthan"],
  ["09", "Uttar Pradesh"], ["10", "Bihar"], ["11", "Sikkim"], ["12", "Arunachal Pradesh"],
  ["13", "Nagaland"], ["14", "Manipur"], ["15", "Mizoram"], ["16", "Tripura"],
  ["17", "Meghalaya"], ["18", "Assam"], ["19", "West Bengal"], ["20", "Jharkhand"],
  ["21", "Odisha"], ["22", "Chhattisgarh"], ["23", "Madhya Pradesh"], ["24", "Gujarat"],
  ["26", "Dadra and Nagar Haveli and Daman and Diu"], ["27", "Maharashtra"], ["28", "Andhra Pradesh (old code)"],
  ["29", "Karnataka"], ["30", "Goa"], ["31", "Lakshadweep"], ["32", "Kerala"],
  ["33", "Tamil Nadu"], ["34", "Puducherry"], ["35", "Andaman and Nicobar Islands"],
  ["36", "Telangana"], ["37", "Andhra Pradesh"], ["38", "Ladakh"],
  ["97", "Other Territory"]
];

const stateByCode = new Map(STATES.map(([code, name]) => [code, name]));
const stateCodeByName = new Map(STATES.map(([code, name]) => [normalize(name), code]));

const elements = {
  fileInput: document.getElementById("file-input"),
  fileLabel: document.getElementById("file-label"),
  fileStatus: document.getElementById("file-status"),
  sheetControl: document.getElementById("sheet-control"),
  sheetSelect: document.getElementById("sheet-select"),
  setupSection: document.getElementById("setup-section"),
  sellerState: document.getElementById("seller-state"),
  mappingGrid: document.getElementById("mapping-grid"),
  convertButton: document.getElementById("convert-button"),
  convertStatus: document.getElementById("convert-status"),
  resultSection: document.getElementById("result-section"),
  resultSummary: document.getElementById("result-summary"),
  errorBox: document.getElementById("error-box"),
  previewHead: document.getElementById("preview-head"),
  previewBody: document.getElementById("preview-body"),
  downloadButton: document.getElementById("download-button")
};

let workbook = null;
let currentRows = [];
let currentFilename = "sales-report";
let convertedCsv = "";

STATES.forEach(([code, name]) => {
  const option = document.createElement("option");
  option.value = code;
  option.textContent = `${code} · ${name}`;
  elements.sellerState.append(option);
});

elements.fileInput.addEventListener("change", handleFileSelection);
elements.sheetSelect.addEventListener("change", () => loadSelectedSheet(true));
elements.convertButton.addEventListener("click", convertReport);
elements.downloadButton.addEventListener("click", downloadReport);

function normalize(value) {
  return String(value == null ? "" : value)
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, " ")
    .trim();
}

function handleFileSelection(event) {
  const file = event.target.files && event.target.files[0];
  if (!file) return;
  if (!window.XLSX) {
    showFileError("The spreadsheet reader did not load. Check your internet connection and reload this page.");
    return;
  }

  currentFilename = file.name.replace(/\.[^.]+$/, "") || "sales-report";
  elements.fileLabel.textContent = file.name;
  elements.fileStatus.textContent = "Reading file in your browser…";
  elements.setupSection.classList.add("hidden");
  elements.resultSection.classList.add("hidden");
  elements.sheetControl.classList.add("hidden");
  elements.convertStatus.textContent = "";
  convertedCsv = "";

  const reader = new FileReader();
  reader.onerror = () => showFileError("The selected file could not be read. Please try another copy.");
  reader.onload = () => {
    try {
      workbook = XLSX.read(reader.result, { type: "array", cellDates: true });
      if (!workbook.SheetNames || workbook.SheetNames.length === 0) {
        showFileError("No worksheets were found in this file.");
        return;
      }
      elements.sheetSelect.replaceChildren();
      workbook.SheetNames.forEach((name) => {
        const option = document.createElement("option");
        option.value = name;
        option.textContent = name;
        elements.sheetSelect.append(option);
      });
      elements.sheetControl.classList.toggle("hidden", workbook.SheetNames.length < 2);
      loadSelectedSheet(false);
    } catch (error) {
      showFileError(`Could not read this spreadsheet: ${error.message}`);
    }
  };
  reader.readAsArrayBuffer(file);
}

function showFileError(message) {
  elements.fileStatus.textContent = message;
  elements.fileStatus.classList.add("error-text");
}

function loadSelectedSheet(preserveMappings) {
  if (!workbook) return;
  const sheet = workbook.Sheets[elements.sheetSelect.value];
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "", raw: false, blankrows: false });
  const headerIndex = rows.findIndex((row) => row.some((cell) => String(cell).trim() !== ""));
  if (headerIndex < 0) {
    showFileError("The selected worksheet is empty.");
    elements.setupSection.classList.add("hidden");
    return;
  }

  const headers = rows[headerIndex].map((value, index) => {
    const text = String(value == null ? "" : value).trim();
    return text || `Column ${index + 1}`;
  });
  currentRows = rows.slice(headerIndex + 1)
    .filter((row) => row.some((cell) => String(cell == null ? "" : cell).trim() !== ""))
    .map((row) => headers.map((_, index) => row[index] == null ? "" : row[index]));

  const previous = {};
  if (preserveMappings) {
    FIELDS.forEach((field) => {
      const selected = document.getElementById(`map-${field.key}`);
      if (selected) previous[field.key] = selected.value;
    });
  }

  renderMappings(headers, previous);
  elements.setupSection.classList.remove("hidden");
  elements.resultSection.classList.add("hidden");
  elements.fileStatus.classList.remove("error-text");
  elements.fileStatus.textContent = `${currentRows.length} data row${currentRows.length === 1 ? "" : "s"} found · ${headers.length} columns detected.`;
}

function renderMappings(headers, previous) {
  elements.mappingGrid.replaceChildren();
  FIELDS.forEach((field) => {
    const wrapper = document.createElement("div");
    wrapper.className = "mapping-field";
    const label = document.createElement("label");
    label.htmlFor = `map-${field.key}`;
    label.textContent = field.label + (field.required ? " *" : "");

    const select = document.createElement("select");
    select.id = `map-${field.key}`;
    select.dataset.field = field.key;
    const noColumn = document.createElement("option");
    noColumn.value = "";
    noColumn.textContent = "— Not mapped —";
    select.append(noColumn);

    headers.forEach((header, index) => {
      const option = document.createElement("option");
      option.value = String(index);
      option.textContent = header;
      select.append(option);
    });

    if (previous[field.key] !== undefined && [...select.options].some((option) => option.value === previous[field.key])) {
      select.value = previous[field.key];
    } else {
      const match = guessColumn(headers, field.aliases);
      if (match >= 0) select.value = String(match);
    }

    wrapper.append(label, select);
    elements.mappingGrid.append(wrapper);
  });
}

function guessColumn(headers, aliases) {
  const normalizedHeaders = headers.map(normalize);
  for (const alias of aliases) {
    const exact = normalizedHeaders.findIndex((header) => header === normalize(alias));
    if (exact >= 0) return exact;
  }
  for (const alias of aliases) {
    const target = normalize(alias);
    const match = normalizedHeaders.findIndex((header) => header && (header.includes(target) || target.includes(header)));
    if (match >= 0) return match;
  }
  return -1;
}

function selectedColumns() {
  return Object.fromEntries(FIELDS.map((field) => {
    const value = document.getElementById(`map-${field.key}`).value;
    return [field.key, value === "" ? -1 : Number(value)];
  }));
}

function cellValue(row, columns, key) {
  const index = columns[key];
  return index < 0 ? "" : row[index];
}

function parseAmount(value, label, rowNumber, errors, required) {
  if (value == null || String(value).trim() === "") {
    if (required) errors.push(`Row ${rowNumber}: ${label} is missing.`);
    return null;
  }
  const original = String(value).trim();
  const parenthesized = /^\(.*\)$/.test(original);
  const cleaned = original
    .replace(/[₹$£€,\s]/g, "")
    .replace(/^\((.*)\)$/, "$1")
    .replace(/%$/, "");
  const number = Number(cleaned);
  if (!Number.isFinite(number)) {
    errors.push(`Row ${rowNumber}: ${label} "${original}" is not a valid number.`);
    return null;
  }
  return parenthesized ? -number : number;
}

function resolveState(value) {
  const raw = String(value == null ? "" : value).trim();
  if (!raw) return null;
  const codeMatch = raw.match(/^\s*(\d{1,2})(?:\D|$)/);
  if (codeMatch) {
    const code = codeMatch[1].padStart(2, "0");
    if (stateByCode.has(code)) return { code, name: stateByCode.get(code) };
  }
  const normalized = normalize(raw);
  const code = stateCodeByName.get(normalized);
  if (code) return { code, name: stateByCode.get(code) };
  const startsWith = STATES.find(([, name]) => {
    const candidate = normalize(name);
    return normalized.startsWith(candidate) || candidate.startsWith(normalized);
  });
  return startsWith ? { code: startsWith[0], name: startsWith[1] } : null;
}

function stateFromGstin(value) {
  const gstin = String(value == null ? "" : value).trim().toUpperCase();
  if (!gstin) return null;
  if (!/^\d{2}[A-Z0-9]{13}$/.test(gstin)) return null;
  return resolveState(gstin.slice(0, 2));
}

function convertReport() {
  elements.convertStatus.textContent = "";
  elements.resultSection.classList.add("hidden");

  const sellerCode = elements.sellerState.value;
  const columns = selectedColumns();
  const errors = [];
  if (!sellerCode) errors.push("Select the seller state or union territory.");
  FIELDS.filter((field) => field.required).forEach((field) => {
    if (columns[field.key] < 0) errors.push(`Map the required "${field.label}" column.`);
  });
  if (currentRows.length === 0) errors.push("The selected worksheet has no data rows.");
  if (errors.length) {
    elements.convertStatus.textContent = errors.join(" ");
    return;
  }

  const seller = { code: sellerCode, name: stateByCode.get(sellerCode) };
  const outputRows = [];
  const rowErrors = [];

  currentRows.forEach((row, index) => {
    const rowNumber = index + 2;
    const invoiceNumber = String(cellValue(row, columns, "invoiceNumber")).trim();
    const invoiceDate = String(cellValue(row, columns, "invoiceDate")).trim();
    const customerGstin = String(cellValue(row, columns, "customerGstin")).trim().toUpperCase();
    const customerName = String(cellValue(row, columns, "customerName")).trim();
    const customerStateValue = cellValue(row, columns, "customerState");
    const posValue = cellValue(row, columns, "placeOfSupply");
    const customerState = resolveState(customerStateValue) || stateFromGstin(customerGstin);
    const placeOfSupply = resolveState(posValue) || customerState;
    const rowIssues = [];

    if (!invoiceNumber) rowIssues.push(`Row ${rowNumber}: invoice number is missing.`);
    if (customerGstin && !stateFromGstin(customerGstin)) rowIssues.push(`Row ${rowNumber}: customer GSTIN "${customerGstin}" is invalid or has an unknown state code.`);

    const taxable = parseAmount(cellValue(row, columns, "taxableValue"), "taxable value", rowNumber, rowIssues, true);
    const rate = columns.gstRate < 0
      ? null
      : parseAmount(cellValue(row, columns, "gstRate"), "GST rate", rowNumber, rowIssues, false);
    if (rate !== null && (rate < 0 || rate > 100)) rowIssues.push(`Row ${rowNumber}: GST rate must be between 0 and 100.`);

    const sourceTax = ["cgst", "sgst", "igst"].map((key) => {
      const raw = cellValue(row, columns, key);
      return raw == null || String(raw).trim() === "" ? null : parseAmount(raw, key.toUpperCase(), rowNumber, rowIssues, false);
    });
    const hasSourceTax = sourceTax.some((amount) => amount !== null);
    let cgst = 0;
    let sgst = 0;
    let igst = 0;

    if (hasSourceTax) {
      [cgst, sgst, igst] = sourceTax.map((amount) => amount === null ? 0 : amount);
    } else if (taxable !== null) {
      if (rate === null) {
        rowIssues.push(`Row ${rowNumber}: provide a GST rate or at least one source tax amount (CGST, SGST, or IGST).`);
      } else if (rate > 0) {
        if (!placeOfSupply) {
          rowIssues.push(`Row ${rowNumber}: place of supply is missing or unrecognized; it is required to calculate GST.`);
        } else {
          const totalTax = roundMoney(taxable * rate / 100);
          if (placeOfSupply.code === seller.code) {
            cgst = roundMoney(totalTax / 2);
            sgst = roundMoney(totalTax - cgst);
          } else {
            igst = totalTax;
          }
        }
      }
    }

    const cess = columns.cess < 0
      ? 0
      : parseAmount(cellValue(row, columns, "cess"), "cess", rowNumber, rowIssues, false) || 0;
    const suppliedTotal = columns.invoiceTotal < 0
      ? null
      : parseAmount(cellValue(row, columns, "invoiceTotal"), "invoice total", rowNumber, rowIssues, false);
    const invoiceTotal = suppliedTotal !== null
      ? suppliedTotal
      : taxable === null ? null : roundMoney(taxable + cgst + sgst + igst + cess);

    if (rowIssues.length) {
      rowErrors.push(...rowIssues);
      return;
    }

    outputRows.push([
      invoiceNumber,
      invoiceDate,
      customerGstin,
      customerName,
      customerState ? customerState.name : String(customerStateValue || "").trim(),
      placeOfSupply ? placeOfSupply.name : String(posValue || "").trim(),
      String(cellValue(row, columns, "hsnSac")).trim(),
      formatMoney(taxable),
      rate === null ? "" : formatMoney(rate),
      formatMoney(cgst),
      formatMoney(sgst),
      formatMoney(igst),
      formatMoney(cess),
      formatMoney(invoiceTotal)
    ]);
  });

  renderResults(outputRows, rowErrors);
}

function roundMoney(value) {
  return Math.round((value + Number.EPSILON) * 100) / 100;
}

function formatMoney(value) {
  return value == null ? "" : roundMoney(value).toFixed(2);
}

function escapeCsv(value, numeric) {
  let text = String(value == null ? "" : value);
  if (!numeric && /^[\s]*[=+\-@\t\r]/.test(text)) text = `'${text}`;
  return `"${text.replace(/"/g, '""')}"`;
}

function renderResults(outputRows, errors) {
  elements.previewHead.replaceChildren();
  elements.previewBody.replaceChildren();
  const headerRow = document.createElement("tr");
  OUTPUT_HEADERS.forEach((header) => {
    const th = document.createElement("th");
    th.textContent = header;
    headerRow.append(th);
  });
  elements.previewHead.append(headerRow);

  outputRows.slice(0, 5).forEach((row) => {
    const tr = document.createElement("tr");
    row.forEach((value) => {
      const td = document.createElement("td");
      td.textContent = value;
      tr.append(td);
    });
    elements.previewBody.append(tr);
  });

  convertedCsv = [
    OUTPUT_HEADERS.map((header) => escapeCsv(header, false)).join(","),
    ...outputRows.map((row) => row.map((value, index) => escapeCsv(value, index >= 7)).join(","))
  ].join("\r\n");
  elements.resultSummary.textContent = `${outputRows.length} row${outputRows.length === 1 ? "" : "s"} ready${errors.length ? ` · ${errors.length} row${errors.length === 1 ? "" : "s"} need attention` : ""}. Preview shows up to 5 rows.`;

  if (errors.length) {
    const intro = document.createElement("strong");
    intro.textContent = "Some rows were not included. Fix these source data issues and convert again:";
    const list = document.createElement("ul");
    errors.slice(0, 12).forEach((error) => {
      const item = document.createElement("li");
      item.textContent = error;
      list.append(item);
    });
    if (errors.length > 12) {
      const item = document.createElement("li");
      item.textContent = `${errors.length - 12} more issue(s) not shown.`;
      list.append(item);
    }
    elements.errorBox.replaceChildren(intro, list);
    elements.errorBox.classList.remove("hidden");
  } else {
    elements.errorBox.classList.add("hidden");
  }

  elements.downloadButton.disabled = outputRows.length === 0;
  elements.downloadButton.classList.toggle("disabled", outputRows.length === 0);
  elements.resultSection.classList.remove("hidden");
}

function downloadReport() {
  if (!convertedCsv) return;
  const blob = new Blob(["\uFEFF", convertedCsv], { type: "text/csv;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = `${currentFilename}-gst-register.csv`;
  document.body.append(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
}
