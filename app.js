const DATA_FILE = "data/database.xlsx";

const els = {
  destinationDate: document.getElementById("destinationDate"),
  destinationDatePicker: document.getElementById("destinationDatePicker"),
  mallFilter: document.getElementById("mallFilter"),
  mallList: document.getElementById("mallList"),
  productFilter: document.getElementById("productFilter"),
  resultSection: document.getElementById("resultSection"),
  resultBody: document.getElementById("resultBody"),
  formMessage: document.getElementById("formMessage"),
};

let databaseRows = [];
let mallOptions = [];
let productOptions = [];
let allProductOptions = [];
let databaseReady = false;

function toArabicDigits(value) {
  const thai = "๐๑๒๓๔๕๖๗๘๙";
  return String(value ?? "").replace(/[๐-๙]/g, (digit) => String(thai.indexOf(digit)));
}

function cleanText(value) {
  return String(value ?? "").replace(/\s+/g, " ").trim();
}

function normalizeHeader(value) {
  return cleanText(value).toLowerCase().replace(/\s+/g, "");
}

function normalizeChoice(value) {
  return cleanText(value).toLowerCase();
}

function toNumber(value) {
  if (typeof value === "number" && Number.isFinite(value)) return value;
  const text = toArabicDigits(value).replace(/,/g, "").trim();
  const number = Number(text);
  return Number.isFinite(number) ? number : null;
}

function pad2(value) {
  return String(value).padStart(2, "0");
}

function formatDate(date) {
  if (!(date instanceof Date) || Number.isNaN(date.getTime())) return "";
  return `${pad2(date.getDate())}/${pad2(date.getMonth() + 1)}/${date.getFullYear()}`;
}

function formatISODate(date) {
  if (!(date instanceof Date) || Number.isNaN(date.getTime())) return "";
  return `${date.getFullYear()}-${pad2(date.getMonth() + 1)}-${pad2(date.getDate())}`;
}

function dateOnly(date) {
  if (!(date instanceof Date) || Number.isNaN(date.getTime())) return null;
  return new Date(date.getFullYear(), date.getMonth(), date.getDate());
}

function todayDateOnly() {
  const now = new Date();
  return new Date(now.getFullYear(), now.getMonth(), now.getDate());
}

function parseDateInput(value) {
  const text = toArabicDigits(value).trim();
  if (!text) return null;

  const slashMatch = text.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  const dashMatch = text.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
  const compactMatch = text.match(/^(\d{2})(\d{2})(\d{4})$/);

  let day;
  let month;
  let year;

  if (slashMatch) {
    day = Number(slashMatch[1]);
    month = Number(slashMatch[2]);
    year = Number(slashMatch[3]);
  } else if (dashMatch) {
    year = Number(dashMatch[1]);
    month = Number(dashMatch[2]);
    day = Number(dashMatch[3]);
  } else if (compactMatch) {
    day = Number(compactMatch[1]);
    month = Number(compactMatch[2]);
    year = Number(compactMatch[3]);
  } else {
    return null;
  }

  const date = new Date(year, month - 1, day);
  const isValid = date.getFullYear() === year && date.getMonth() === month - 1 && date.getDate() === day;
  return isValid ? date : null;
}

function parseExcelDate(value) {
  if (value instanceof Date) return dateOnly(value);

  if (typeof value === "number" && Number.isFinite(value)) {
    if (window.XLSX?.SSF?.parse_date_code) {
      const parsed = window.XLSX.SSF.parse_date_code(value);
      if (parsed && parsed.y && parsed.m && parsed.d) {
        return new Date(parsed.y, parsed.m - 1, parsed.d);
      }
    }
    const excelEpoch = new Date(Date.UTC(1899, 11, 30));
    const ms = excelEpoch.getTime() + Math.round(value) * 86400000;
    const utc = new Date(ms);
    return new Date(utc.getUTCFullYear(), utc.getUTCMonth(), utc.getUTCDate());
  }

  return parseDateInput(value);
}

function setDateField(textInput, datePicker, date) {
  textInput.value = formatDate(date);
  datePicker.value = formatISODate(date);
}

function normalizeDateField(textInput, datePicker) {
  const parsed = parseDateInput(textInput.value);
  if (parsed) {
    setDateField(textInput, datePicker, parsed);
  } else if (!textInput.value.trim()) {
    datePicker.value = "";
  }
}

function addDays(date, days) {
  const result = new Date(date.getFullYear(), date.getMonth(), date.getDate());
  result.setDate(result.getDate() + Number(days));
  return result;
}

function findValueByExact(row, headers) {
  const wanted = headers.map(normalizeHeader);
  const key = Object.keys(row).find((candidate) => wanted.includes(normalizeHeader(candidate)));
  return key ? row[key] : undefined;
}

function findValueByContains(row, includeTerms, excludeTerms = []) {
  const include = includeTerms.map(normalizeHeader);
  const exclude = excludeTerms.map(normalizeHeader);

  const key = Object.keys(row).find((candidate) => {
    const normalized = normalizeHeader(candidate);
    return include.every((term) => normalized.includes(term)) && !exclude.some((term) => normalized.includes(term));
  });

  return key ? row[key] : undefined;
}

function readRow(row) {
  const mall = cleanText(findValueByExact(row, ["ห้าง"]) ?? findValueByContains(row, ["ห้าง"]));
  const product = cleanText(findValueByExact(row, ["สินค้า"]) ?? findValueByContains(row, ["สินค้า"]));

  const shelfLifeDays = toNumber(
    findValueByExact(row, ["อายุสินค้า", "อายุสินค้า (นับจากวันผลิต)"]) ??
    findValueByContains(row, ["อายุสินค้า"])
  );

  const productionDate = parseExcelDate(
    findValueByExact(row, ["Input (วันผลิต)", "วันผลิต"]) ??
    findValueByContains(row, ["วันผลิต"], ["จำนวน", "ต้องไม่เกิน", "ปัจจุบัน", "เทียบ"])
  );

  const criteriaDaysFromFile = toNumber(
    findValueByExact(row, ["เกณฑ์การรับสินค้า ของดีซีห้าง ต้องไม่เกินกี่วัน นับจากวันผลิต"]) ??
    findValueByContains(row, ["เกณฑ์", "กี่วัน"]) ??
    findValueByContains(row, ["ต้องไม่เกินกี่วัน"])
  );

  const criteriaPercent = toNumber(
    findValueByContains(row, ["เกณฑ์", "%"]) ??
    findValueByContains(row, ["กี่%"])
  );

  let criteriaDays = criteriaDaysFromFile;
  if (criteriaDays === null && shelfLifeDays !== null && criteriaPercent !== null) {
    criteriaDays = Math.floor(shelfLifeDays * criteriaPercent) - (product === "สด" ? 3 : 7);
  }
  if (criteriaDays === null && shelfLifeDays !== null) {
    criteriaDays = Math.floor(shelfLifeDays * 0.25) - 1;
  }

  const codeStartDateFromFile = parseExcelDate(
    findValueByExact(row, ["วันผลิต (รหัสจ่ายได้ตั้งแต่วันที่)", "วันผลิต(รหัสจ่ายได้ตั้งแต่วันที่)", "วันผลิต+จำนวนวันที่ต้องไม่เกิน นับจากวันผลิต"]) ??
    findValueByContains(row, ["รหัสจ่ายได้ตั้งแต่วันที่"]) ??
    findValueByContains(row, ["วันผลิต", "จำนวน", "ต้องไม่เกิน"])
  );

  const codeStartDate = codeStartDateFromFile || (productionDate && criteriaDays !== null ? addDays(productionDate, criteriaDays) : null);

  if (!mall || !product || shelfLifeDays === null || criteriaDays === null || !productionDate || !codeStartDate) return null;

  return {
    mall,
    product,
    shelfLifeDays,
    productionDate,
    criteriaDays,
    codeStartDate,
  };
}

function uniqueSorted(values) {
  return [...new Set(values.filter(Boolean))].sort((a, b) => a.localeCompare(b, "th"));
}

function exactChoice(value, options) {
  const normalized = normalizeChoice(value);
  return options.find((option) => normalizeChoice(option) === normalized) || "";
}

function filterChoices(query, options) {
  const normalizedQuery = normalizeChoice(query);
  if (!normalizedQuery) return options;
  return options.filter((option) => normalizeChoice(option).includes(normalizedQuery));
}

function setFormMessage(message = "", state = "") {
  els.formMessage.textContent = message;
  els.formMessage.classList.toggle("is-error", state === "error");
}

function openCombo(combo, input, list, options) {
  if (input.disabled) return;
  renderComboList(combo, input, list, options);
  combo.classList.add("is-open");
  input.setAttribute("aria-expanded", "true");
}

function closeCombo(combo, input) {
  combo.classList.remove("is-open");
  input.setAttribute("aria-expanded", "false");
}

function renderComboList(combo, input, list, options) {
  const filtered = filterChoices(input.value, options);
  list.innerHTML = "";

  if (!filtered.length) {
    const empty = document.createElement("div");
    empty.className = "combo-empty";
    empty.textContent = "ไม่พบรายการ";
    list.appendChild(empty);
    return;
  }

  filtered.forEach((option) => {
    const button = document.createElement("button");
    button.type = "button";
    button.className = "combo-option";
    button.setAttribute("role", "option");
    button.textContent = option;
    button.addEventListener("mousedown", (event) => event.preventDefault());
    button.addEventListener("click", () => {
      input.value = option;
      closeCombo(combo, input);
      if (input === els.mallFilter) {
        refreshProductOptions(true);
      }
      syncCodeStartDate();
      renderResults();
    });
    list.appendChild(button);
  });
}

function setupCombo(input, list, getOptions, onInputChange) {
  const combo = input.closest(".combo");
  const toggle = combo.querySelector(".combo-toggle");

  input.addEventListener("focus", () => openCombo(combo, input, list, getOptions()));
  input.addEventListener("input", () => {
    openCombo(combo, input, list, getOptions());
    if (onInputChange) onInputChange();
    syncCodeStartDate();
    renderResults();
  });
  input.addEventListener("blur", () => {
    window.setTimeout(() => closeCombo(combo, input), 130);
  });

  toggle.addEventListener("click", () => {
    if (combo.classList.contains("is-open")) {
      closeCombo(combo, input);
    } else {
      input.focus();
      openCombo(combo, input, list, getOptions());
    }
  });
}

function refreshFilters() {
  mallOptions = uniqueSorted(databaseRows.map((row) => row.mall));
  allProductOptions = uniqueSorted(databaseRows.map((row) => row.product));
  refreshProductOptions(false);
}

function setProductSelectOptions(options, selectedValue = "") {
  els.productFilter.innerHTML = "";

  const placeholder = document.createElement("option");
  placeholder.value = "";
  placeholder.textContent = "กรุณาเลือก";
  els.productFilter.appendChild(placeholder);

  options.forEach((option) => {
    const item = document.createElement("option");
    item.value = option;
    item.textContent = option;
    els.productFilter.appendChild(item);
  });

  els.productFilter.value = exactChoice(selectedValue, options) || "";
  els.productFilter.disabled = false;
}

function refreshProductOptions(clearProduct) {
  const selectedMall = exactChoice(els.mallFilter.value, mallOptions);

  productOptions = selectedMall
    ? uniqueSorted(
        databaseRows
          .filter((row) => row.mall === selectedMall)
          .map((row) => row.product)
      )
    : allProductOptions;

  const previousValue = clearProduct ? "" : els.productFilter.value;
  setProductSelectOptions(productOptions, previousValue);
}

function getMatchingRows() {
  const mall = exactChoice(els.mallFilter.value, mallOptions);
  const product = exactChoice(els.productFilter.value, productOptions);
  if (!mall || !product) return [];
  return databaseRows.filter((row) => row.mall === mall && row.product === product);
}

function syncCodeStartDate() {
  const rows = getMatchingRows();
  const firstRow = rows[0];
}

function getValidation() {
  if (!databaseReady) {
    return { valid: false, message: "กำลังโหลดฐานข้อมูล..." };
  }

  const mallText = els.mallFilter.value.trim();
  const productText = els.productFilter.value.trim();
  const destinationDateText = els.destinationDate.value.trim();

  // วันที่ถึงลูกค้าปลายทางมีค่า default เป็นวันที่ปัจจุบันแล้ว
  // แสดงผลทันทีเมื่อผู้ใช้เลือกห้างและสินค้า
  if (!mallText || !productText) {
    return { valid: false, message: "" };
  }

  const mall = exactChoice(mallText, mallOptions);
  if (!mall) {
    return { valid: false, message: "กรุณาเลือกห้างจากรายการ", error: true };
  }

  refreshProductOptions(false);

  const product = exactChoice(productText, productOptions);
  if (!product) {
    return { valid: false, message: "กรุณาเลือกสินค้าจากรายการ", error: true };
  }

  const destinationDate = parseDateInput(destinationDateText);
  if (!destinationDate) {
    return { valid: false, message: "รูปแบบวันที่ถึงลูกค้าปลายทางไม่ถูกต้อง กรุณากรอกเป็น วัน/เดือน/ปี เช่น 29/06/2026", error: true };
  }

  const rows = getMatchingRows();
  if (!rows.length) {
    return { valid: false, message: "ไม่พบข้อมูลที่ตรงกับตัวกรอง", error: true };
  }

  return { valid: true, destinationDate, codeStartDate: rows[0].codeStartDate, mall, product, rows };
}

function createCell(label, value, extraClass = "") {
  const cell = document.createElement("td");
  cell.setAttribute("data-label", label);
  if (extraClass) cell.className = extraClass;

  if (value instanceof Node) {
    cell.appendChild(value);
  } else {
    cell.textContent = value;
  }

  return cell;
}

function createBadge(result) {
  const badge = document.createElement("span");
  const normalized = result === "Yes" ? "yes" : "no";
  badge.className = `badge ${normalized}`;
  badge.textContent = result;
  return badge;
}

function renderResults() {
  const validation = getValidation();

  if (!validation.valid) {
    els.resultSection.classList.add("is-hidden");
    els.resultBody.innerHTML = "";
    setFormMessage(validation.message, validation.error ? "error" : "");
    return;
  }

  setFormMessage("");
  els.resultSection.classList.remove("is-hidden");
  els.resultBody.innerHTML = "";

  validation.rows.forEach((row) => {
    const result = validation.destinationDate < row.codeStartDate ? "Yes" : "No";

    const tr = document.createElement("tr");
    tr.appendChild(createCell("รหัสจ่ายได้ตั้งแต่วันที่", formatDate(row.codeStartDate)));
    tr.appendChild(createCell("ห้าง", row.mall));
    tr.appendChild(createCell("สินค้า", row.product));
    tr.appendChild(createCell("อายุสินค้า", `${row.shelfLifeDays.toLocaleString("th-TH")} วัน`));
    tr.appendChild(createCell("เกณฑ์รับสินค้า (วัน)", `${row.criteriaDays.toLocaleString("th-TH")} วัน`));
    tr.appendChild(createCell("วันที่ถึงลูกค้าปลายทาง", formatDate(validation.destinationDate)));
    els.resultBody.appendChild(tr);
  });
}

async function loadDatabase() {
  try {
    setFormMessage("กำลังโหลดฐานข้อมูล...");

    if (!window.XLSX) {
      throw new Error("ไม่สามารถโหลด XLSX library ได้ กรุณาตรวจสอบ internet connection");
    }

    const response = await fetch(DATA_FILE, { cache: "no-store" });
    if (!response.ok) {
      throw new Error(`โหลดไฟล์ฐานข้อมูลไม่สำเร็จ (${response.status})`);
    }

    const arrayBuffer = await response.arrayBuffer();
    const workbook = XLSX.read(arrayBuffer, { type: "array", cellDates: false });
    const firstSheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[firstSheetName];
    const rawRows = XLSX.utils.sheet_to_json(sheet, { defval: "", raw: true });

    databaseRows = rawRows.map(readRow).filter(Boolean);

    if (!databaseRows.length) {
      throw new Error("ไม่พบ row ที่อ่านได้จาก Excel กรุณาตรวจสอบ header ของไฟล์ database.xlsx");
    }

    databaseReady = true;
    refreshFilters();
    syncCodeStartDate();
    renderResults();
  } catch (error) {
    console.error(error);
    databaseReady = false;
    els.resultSection.classList.add("is-hidden");
    setFormMessage(error.message || "โหลดฐานข้อมูลไม่สำเร็จ", "error");
  }
}

function bindDateField(textInput, datePicker) {
  textInput.addEventListener("blur", () => {
    normalizeDateField(textInput, datePicker);
    renderResults();
  });

  textInput.addEventListener("input", () => {
    const parsed = parseDateInput(textInput.value);
    datePicker.value = parsed ? formatISODate(parsed) : "";
    renderResults();
  });

  datePicker.addEventListener("change", () => {
    const parsed = parseDateInput(datePicker.value);
    if (parsed) {
      setDateField(textInput, datePicker, parsed);
    }
    renderResults();
  });
}

function updateProductColor() {
  if (els.productFilter.value) {
    els.productFilter.classList.add("has-value");
  } else {
    els.productFilter.classList.remove("has-value");
  }
}

function bindEvents() {
  bindDateField(els.destinationDate, els.destinationDatePicker);

  setupCombo(
    els.mallFilter,
    els.mallList,
    () => mallOptions,
    () => refreshProductOptions(true)
  );

  els.productFilter.addEventListener("change", () => {
    updateProductColor();
    syncCodeStartDate();
    renderResults();
  });

  document.addEventListener("click", (event) => {
    document.querySelectorAll(".combo.is-open").forEach((combo) => {
      if (!combo.contains(event.target)) {
        const input = combo.querySelector("input");
        closeCombo(combo, input);
      }
    });
  });
}

function init() {
  setDateField(els.destinationDate, els.destinationDatePicker, todayDateOnly());
  updateProductColor();
  bindEvents();
  loadDatabase();
}

document.addEventListener("DOMContentLoaded", init);
