
"use strict";

/* =========================================
   1. CẤU HÌNH VÀ TRẠNG THÁI
========================================= */

const STORAGE_KEY = "crm_excel_sync_v3";

const state = {
  customers: [],
  query: "",
  type: "",
  symbol: "",
  deadline: "",
  fileName: "",
  syncedAt: ""
};

const $ = selector => document.querySelector(selector);

const el = {
  excelFile: $("#excelFile"),
  syncBtn: $("#syncExcelBtn"),
  exportBtn: $("#exportBtn"),
  exportExcelBtn: $("#exportExcelBtn"),
  addCustomerBtn: $("#addCustomerBtn"),
  dialog: $("#customerDialog"),
  form: $("#customerForm"),
  dialogTitle: $("#dialogTitle"),
  closeDialogBtn: $("#closeDialogBtn"),
  cancelDialogBtn: $("#cancelDialogBtn"),
  rows: $("#customerRows"),
  empty: $("#emptyState"),
  total: $("#totalCustomers"),
  newCustomers: $("#newCustomers"),
  revenue: $("#expectedRevenue"),
  follow: $("#needFollowUp"),
  search: $("#searchInput"),
  typeFilter: $("#typeFilter"),
  symbolFilter: $("#symbolFilter"),
  deadlineFilter: $("#deadlineFilter"),
  clearFilter: $("#clearFilterBtn"),
  resultCount: $("#resultCount"),
  filterIndicator: $("#filterIndicator"),
  fileInfo: $("#fileInfo"),
  pipeline: $("#pipelineGrid"),
  lastUpdated: $("#lastUpdated"),
  importStatus: $("#importStatus"),
  sideTotal: $("#sideTotal"),
  sideUrgent: $("#sideUrgent")
};

/* =========================================
   2. HÀM TIỆN ÍCH
========================================= */

function escapeHTML(value) {
  return String(value ?? "").replace(/[&<>"']/g, char => ({
    "&": "&amp;",
    "<": "&lt;",
    ">": "&gt;",
    '"': "&quot;",
    "'": "&#039;"
  })[char]);
}

function createId() {
  if (window.crypto?.randomUUID) {
    return crypto.randomUUID();
  }

  return "customer-" + Date.now() + "-" +
    Math.random().toString(16).slice(2);
}

function normalizeSearch(value) {
  return String(value ?? "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase()
    .trim();
}

function formatMoney(value) {
  return new Intl.NumberFormat("vi-VN", {
    maximumFractionDigits: 0
  }).format(Number(value) || 0) + " ₫";
}

function parseMoney(value) {
  if (typeof value === "number") {
    return Number.isFinite(value) ? value : 0;
  }

  return Number(
    String(value ?? "").replace(/[^\d-]/g, "")
  ) || 0;
}

/* =========================================
   3. CHUẨN HÓA NGÀY THÁNG
   Định dạng đầu ra: dd/mm/yyyy
========================================= */

function makeDateText(day, month, year) {
  const d = Number(day);
  const m = Number(month);
  const y = Number(year);

  if (
    !Number.isInteger(d) ||
    !Number.isInteger(m) ||
    !Number.isInteger(y) ||
    y < 1000 ||
    m < 1 || m > 12 ||
    d < 1 || d > 31
  ) {
    return "";
  }

  const date = new Date(y, m - 1, d);

  if (
    date.getFullYear() !== y ||
    date.getMonth() !== m - 1 ||
    date.getDate() !== d
  ) {
    return "";
  }

  return [
    String(d).padStart(2, "0"),
    String(m).padStart(2, "0"),
    String(y).padStart(4, "0")
  ].join("/");
}

function formatDateOnly(value) {
  if (value === null || value === undefined || value === "") {
    return "";
  }

  // Đối tượng Date từ JavaScript hoặc Excel.
  if (value instanceof Date) {
    if (Number.isNaN(value.getTime())) return "";

    return makeDateText(
      value.getDate(),
      value.getMonth() + 1,
      value.getFullYear()
    );
  }

  const text = String(value).trim();

  if (!text) return "";

  // Giữ nguyên dạng ngày/tháng không có năm, ví dụ 6/7.
  const shortDM = text.match(/^(\d{1,2})\s*\/\s*(\d{1,2})$/);

  if (shortDM) {
    return `${shortDM[1].padStart(2, "0")}/${shortDM[2].padStart(2, "0")}`;
  }

  // dd/mm/yyyy hoặc dd-mm-yyyy.
  const vietnamDate = text.match(
    /^(\d{1,2})[/-](\d{1,2})[/-](\d{4})$/
  );

  if (vietnamDate) {
    return makeDateText(
      vietnamDate[1],
      vietnamDate[2],
      vietnamDate[3]
    ) || text;
  }

  // yyyy-mm-dd hoặc yyyy-mm-ddTHH:mm:ss...
  const isoDate = text.match(
    /^(\d{4})-(\d{1,2})-(\d{1,2})(?:$|T|\s)/
  );

  if (isoDate) {
    return makeDateText(
      isoDate[3],
      isoDate[2],
      isoDate[1]
    ) || text;
  }

  // Chuỗi kiểu:
  // Thu Jan 15 2026 23:59:56 GMT+0700 (Giờ Đông Dương)
  // Chỉ chuẩn hóa khi nhận diện được chuỗi ngày giờ.
  const looksLikeDateTime =
    /\b(?:Mon|Tue|Wed|Thu|Fri|Sat|Sun)\b/i.test(text) ||
    /\bGMT[+-]\d{4}\b/i.test(text) ||
    /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}/.test(text);

  if (looksLikeDateTime) {
    const date = new Date(text);

    if (!Number.isNaN(date.getTime())) {
      return makeDateText(
        date.getDate(),
        date.getMonth() + 1,
        date.getFullYear()
      );
    }
  }

  // Không đoán những dữ liệu không xác định được.
  return text;
}

/* =========================================
   4. CHUẨN HÓA TIẾN TRÌNH
   - Nếu là chuỗi ngày giờ: chỉ hiện ngày.
   - Nếu là chữ bình thường: giữ nguyên.
========================================= */

function formatProgress(value) {
  if (value === null || value === undefined || value === "") {
    return "";
  }

  const text = String(value).trim();

  if (!text) return "";

  const looksLikeDateTime =
    /\b(?:Mon|Tue|Wed|Thu|Fri|Sat|Sun)\b/i.test(text) ||
    /\bGMT[+-]\d{4}\b/i.test(text) ||
    /\b(?:Giờ Đông Dương|Indochina Time)\b/i.test(text) ||
    /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}/.test(text);

  if (looksLikeDateTime) {
    return formatDateOnly(text);
  }

  // Giữ nguyên tiến trình dạng văn bản.
  return text;
}

function dateText() {
  return formatDateOnly(new Date());
}

/* =========================================
   5. LƯU VÀ KHÔI PHỤC DỮ LIỆU
========================================= */

function saveData() {
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify({
      customers: state.customers,
      fileName: state.fileName,
      syncedAt: state.syncedAt
    }));
  } catch (error) {
    console.error("Lỗi lưu dữ liệu:", error);
    alert("Không thể lưu dữ liệu. Hãy kiểm tra dung lượng trình duyệt.");
  }
}

function restoreData() {
  try {
    const saved = JSON.parse(localStorage.getItem(STORAGE_KEY));

    if (!saved) return;

    state.customers = Array.isArray(saved.customers)
      ? saved.customers.map(customer => ({
          ...customer,
          closeDate: formatDateOnly(customer.closeDate),
          deadline: formatDateOnly(customer.deadline),
          progress: formatProgress(customer.progress)
        }))
      : [];

    state.fileName = saved.fileName || "";
    state.syncedAt = formatDateOnly(saved.syncedAt);
  } catch (error) {
    console.error("Không đọc được dữ liệu đã lưu:", error);
    state.customers = [];
  }
}

/* =========================================
   6. TRẠNG THÁI HẠN CHỐT
========================================= */

function getDeadlineStatus(customer) {
  const deadline = String(customer.deadline ?? "").trim();

  if (!deadline || deadline === "0" || deadline === "-") {
    return "zero";
  }

  let day;
  let month;
  let year;

  const fullDate = deadline.match(
    /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/
  );

  const shortDate = deadline.match(
    /^(\d{1,2})\/(\d{1,2})$/
  );

  if (fullDate) {
    day = Number(fullDate[1]);
    month = Number(fullDate[2]);
    year = Number(fullDate[3]);
  } else if (shortDate) {
    day = Number(shortDate[1]);
    month = Number(shortDate[2]);
    year = new Date().getFullYear();
  } else {
    return "normal";
  }

  let deadlineDate = new Date(year, month - 1, day, 23, 59, 59);

  if (
    deadlineDate.getDate() !== day ||
    deadlineDate.getMonth() !== month - 1
  ) {
    return "normal";
  }

  const now = new Date();

  // Với ngày/tháng không có năm, nếu đã qua lâu thì xét năm sau.
  if (
    shortDate &&
    deadlineDate.getTime() < now.getTime() - 180 * 86400000
  ) {
    deadlineDate = new Date(year + 1, month - 1, day, 23, 59, 59);
  }

  const diff = deadlineDate.getTime() - now.getTime();

  if (diff < 0) return "overdue";
  if (diff <= 2 * 86400000) return "urgent";

  return "normal";
}

function deadlineLabel(customer) {
  const labels = {
    zero: ["Không hạn", "tag-zero"],
    normal: ["Có hạn", "tag-normal"],
    urgent: ["Khẩn cấp", "tag-urgent"],
    overdue: ["Quá hạn", "tag-overdue"]
  };

  return labels[getDeadlineStatus(customer)];
}

function isNewCustomer(customer) {
  return normalizeSearch(customer.customerType).includes("moi");
}

/* =========================================
   7. LỌC KHÁCH HÀNG
========================================= */

function getFilteredCustomers() {
  return state.customers.filter(customer => {
    const searchable = normalizeSearch([
      customer.name,
      customer.customerType,
      customer.symbol,
      customer.explanation,
      customer.progress,
      customer.closeDate,
      customer.deadline
    ].join(" "));

    if (state.query && !searchable.includes(state.query)) {
      return false;
    }

    if (state.type && customer.customerType !== state.type) {
      return false;
    }

    if (state.symbol && customer.symbol !== state.symbol) {
      return false;
    }

    if (
      state.deadline &&
      getDeadlineStatus(customer) !== state.deadline
    ) {
      return false;
    }

    return true;
  });
}

/* =========================================
   8. HIỂN THỊ GIAO DIỆN
========================================= */

function render() {
  renderFilters();
  renderMetrics();
  renderRows();
  renderPipeline();
  renderSummary();
  syncQuickButtons();
}

function renderFilters() {
  const types = [...new Set(
    state.customers.map(c => c.customerType).filter(Boolean)
  )].sort();

  const symbols = [...new Set(
    state.customers.map(c => c.symbol).filter(Boolean)
  )].sort();

  const previousType = state.type;
  const previousSymbol = state.symbol;

  el.typeFilter.innerHTML =
    '<option value="">Tất cả loại khách</option>' +
    types.map(type =>
      `<option value="${escapeHTML(type)}">${escapeHTML(type)}</option>`
    ).join("");

  el.symbolFilter.innerHTML =
    '<option value="">Tất cả ký hiệu</option>' +
    symbols.map(symbol =>
      `<option value="${escapeHTML(symbol)}">${escapeHTML(symbol)}</option>`
    ).join("");

  el.typeFilter.value = types.includes(previousType) ? previousType : "";
  el.symbolFilter.value = symbols.includes(previousSymbol)
    ? previousSymbol
    : "";

  state.type = el.typeFilter.value;
  state.symbol = el.symbolFilter.value;
}

function renderMetrics() {
  const customers = state.customers;

  const revenue = customers.reduce(
    (sum, c) => sum + (Number(c.value) || 0),
    0
  );

  const followCount = customers.filter(c => {
    const status = getDeadlineStatus(c);
    return status === "urgent" || status === "overdue";
  }).length;

  el.total.textContent = customers.length;
  el.newCustomers.textContent = customers.filter(isNewCustomer).length;
  el.revenue.textContent = formatMoney(revenue);
  el.follow.textContent = followCount;

  el.sideTotal.textContent = customers.length;
  el.sideUrgent.textContent = followCount;
}

function renderRows() {
  const customers = getFilteredCustomers();

  el.rows.innerHTML = customers.map(customer => {
    const [deadlineText, deadlineClass] = deadlineLabel(customer);

    const typeClass = normalizeSearch(customer.customerType).includes("vip")
      ? "tag-vip"
      : isNewCustomer(customer)
        ? "tag-new"
        : "";

    return `
      <tr>
        <td>${escapeHTML(customer.stt)}</td>

        <td>
          <div class="customer-name">
            ${escapeHTML(customer.name)}
          </div>
        </td>

        <td>
          ${escapeHTML(formatDateOnly(customer.closeDate) || "—")}
        </td>

        <td>
          ${escapeHTML(formatProgress(customer.progress) || "—")}
        </td>

        <td>
          <span class="tag ${deadlineClass}">
            ${escapeHTML(formatDateOnly(customer.deadline) || "Không hạn")}
            · ${escapeHTML(deadlineText)}
          </span>
        </td>

        <td>
          <span class="amount">${formatMoney(customer.value)}</span>
        </td>

        <td>
          <span class="tag ${typeClass}">
            ${escapeHTML(customer.customerType || "Chưa phân loại")}
          </span>
        </td>

        <td>${escapeHTML(customer.symbol || "—")}</td>

        <td>
          <div class="customer-note">
            ${escapeHTML(customer.explanation || "—")}
          </div>
        </td>

        <td>
          <div class="row-actions">
            <button
              data-action="receipt"
              data-id="${escapeHTML(customer.id)}"
              class="receipt-btn"
              type="button"
            >Biên lai</button>

            <button
              data-action="edit"
              data-id="${escapeHTML(customer.id)}"
              type="button"
            >Sửa</button>

            <button
              data-action="delete"
              data-id="${escapeHTML(customer.id)}"
              class="delete-btn"
              type="button"
            >Xóa</button>
          </div>
        </td>
      </tr>
    `;
  }).join("");

  el.empty.classList.toggle("show", customers.length === 0);
}

function renderPipeline() {
  const groups = {};

  state.customers.forEach(customer => {
    const type = customer.customerType || "Chưa phân loại";

    if (!groups[type]) {
      groups[type] = { count: 0, revenue: 0 };
    }

    groups[type].count++;
    groups[type].revenue += Number(customer.value) || 0;
  });

  const entries = Object.entries(groups);

  if (!entries.length) {
    el.pipeline.innerHTML = `
      <div class="pipeline-card">
        <span>Chưa có dữ liệu</span>
        <strong>0</strong>
        <small>Thêm khách hàng hoặc nhập Excel.</small>
      </div>
    `;
    return;
  }

  el.pipeline.innerHTML = entries.map(([type, data]) => `
    <article class="pipeline-card">
      <span>${escapeHTML(type)}</span>
      <strong>${data.count} khách</strong>
      <small>Doanh số: ${formatMoney(data.revenue)}</small>
    </article>
  `).join("");
}

function renderSummary() {
  const customers = getFilteredCustomers();

  el.resultCount.textContent = `${customers.length} khách hàng`;

  const hasFilters =
    state.query || state.type || state.symbol || state.deadline;

  el.filterIndicator.textContent = hasFilters
    ? "Đang áp dụng bộ lọc"
    : "Đang hiển thị tất cả";

  el.fileInfo.textContent = state.fileName
    ? `File: ${state.fileName}`
    : "Chưa đồng bộ Excel.";

  el.lastUpdated.textContent = state.syncedAt
    ? `Cập nhật: ${formatDateOnly(state.syncedAt)}`
    : "Chưa cập nhật";
}

function syncQuickButtons() {
  document.querySelectorAll(".quick-btn").forEach(button => {
    button.classList.toggle(
      "active",
      button.dataset.deadline === state.deadline
    );
  });
}

/* =========================================
   9. THÊM / SỬA KHÁCH HÀNG
========================================= */

function openCustomerForm(customer = null) {
  el.form.reset();

  el.dialogTitle.textContent = customer
    ? "Chỉnh sửa khách hàng"
    : "Thêm dữ liệu";

  $("#customerId").value = customer?.id || "";
  $("#sttInput").value = customer?.stt ?? (state.customers.length + 1);
  $("#nameInput").value = customer?.name || "";
  $("#closeDateInput").value = formatDateOnly(customer?.closeDate);
  $("#progressInput").value = formatProgress(customer?.progress);
  $("#deadlineInput").value = formatDateOnly(customer?.deadline);
  $("#valueInput").value = customer?.value ?? "";
  $("#customerTypeInput").value = customer?.customerType || "";
  $("#symbolInput").value = customer?.symbol || "";
  $("#explanationInput").value = customer?.explanation || "";

  if (!el.dialog.open) {
    el.dialog.showModal();
  }
}

function closeDialog() {
  if (el.dialog.open) {
    el.dialog.close();
  }
}

function saveCustomer(event) {
  event.preventDefault();

  const id = $("#customerId").value || createId();

  const customer = {
    id,
    stt: Number($("#sttInput").value) || (state.customers.length + 1),
    name: $("#nameInput").value.trim(),
    closeDate: formatDateOnly($("#closeDateInput").value),
    progress: formatProgress($("#progressInput").value),
    deadline: formatDateOnly($("#deadlineInput").value),
    value: parseMoney($("#valueInput").value),
    customerType: $("#customerTypeInput").value.trim(),
    symbol: $("#symbolInput").value.trim(),
    explanation: $("#explanationInput").value.trim()
  };

  if (!customer.name) {
    alert("Vui lòng nhập tên khách hàng.");
    return;
  }

  const index = state.customers.findIndex(c => c.id === id);

  if (index >= 0) {
    state.customers[index] = customer;
  } else {
    state.customers.push(customer);
  }

  saveData();
  closeDialog();
  render();
}

/* =========================================
   10. SỬA / XÓA / XUẤT BIÊN LAI
========================================= */

function handleRowAction(event) {
  const button = event.target.closest("button[data-action]");

  if (!button) return;

  const { action, id } = button.dataset;
  const customer = state.customers.find(c => c.id === id);

  if (!customer) return;

  if (action === "edit") {
    openCustomerForm(customer);
    return;
  }

  if (action === "receipt") {
    exportReceipt(id);
    return;
  }

  if (action === "delete") {
    if (!confirm(`Bạn có chắc muốn xóa khách hàng "${customer.name}" không?`)) {
      return;
    }

    state.customers = state.customers.filter(c => c.id !== id);

    saveData();
    render();
  }
}

/* =========================================
   11. NHẬP EXCEL / CSV
========================================= */

function normalizeHeader(value) {
  return normalizeSearch(value).replace(/[^a-z0-9]/g, "");
}

function findValue(row, aliases) {
  const keys = Object.keys(row);

  for (const alias of aliases) {
    const normalizedAlias = normalizeHeader(alias);

    const found = keys.find(key =>
      normalizeHeader(key) === normalizedAlias
    );

    if (found !== undefined) return row[found];
  }

  return "";
}

function normalizeImportedDate(value) {
  if (value === null || value === undefined || value === "") {
    return "";
  }

  return formatDateOnly(value);
}

function normalizeCustomer(row, index) {
  return {
    id: createId(),

    stt: Number(findValue(row, [
      "STT", "Số thứ tự"
    ])) || index + 1,

    name: String(findValue(row, [
      "Khách Hàng", "Khách hàng", "Tên khách hàng", "Tên"
    ]) ?? "").trim(),

    closeDate: normalizeImportedDate(findValue(row, [
      "Ngày Chốt", "Ngày chốt", "Ngày"
    ])),

    // Quan trọng: chuẩn hóa cả cột Tiến trình khi nhập Excel.
    progress: formatProgress(findValue(row, [
      "Tiến Trình", "Tiến trình", "Tiến độ"
    ])),

    deadline: normalizeImportedDate(findValue(row, [
      "Hạn Chốt", "Hạn chốt", "Deadline"
    ])),

    value: parseMoney(findValue(row, [
      "Doanh Số", "Doanh số", "Giá trị", "Số tiền"
    ])),

    customerType: String(findValue(row, [
      "Loại Khách", "Loại khách", "Phân loại"
    ]) ?? "").trim(),

    symbol: String(findValue(row, [
      "Ký Hiệu", "Ký hiệu", "Nền tảng"
    ]) ?? "").trim(),

    explanation: String(findValue(row, [
      "Giải Thích", "Giải thích", "Ghi chú", "Ghi chu"
    ]) ?? "").trim()
  };
}

async function syncFromExcelFile() {
  const file = el.excelFile.files[0];

  if (!file) {
    alert("Vui lòng chọn file Excel hoặc CSV trước.");
    return;
  }

  if (typeof XLSX === "undefined") {
    alert("Không tải được thư viện Excel. Hãy kiểm tra kết nối Internet.");
    return;
  }

  try {
    el.importStatus.textContent = "Đang đọc dữ liệu...";

    const buffer = await file.arrayBuffer();

    const workbook = XLSX.read(buffer, {
      type: "array",
      cellDates: true
    });

    const sheet = workbook.Sheets[workbook.SheetNames[0]];

    const rows = XLSX.utils.sheet_to_json(sheet, {
      defval: "",
      raw: true
    });

    if (!rows.length) {
      alert("File không có dữ liệu hoặc không tìm thấy dòng tiêu đề.");
      el.importStatus.textContent = "Không có dữ liệu để nhập.";
      return;
    }

    const imported = rows
      .map(normalizeCustomer)
      .filter(customer => customer.name);

    if (!imported.length) {
      alert("Không tìm thấy cột tên khách hàng phù hợp.");
      el.importStatus.textContent = "Không nhận diện được dữ liệu.";
      return;
    }

    const replace = confirm(
      `Đã đọc ${imported.length} khách hàng.\n\n` +
      "Nhấn OK để thay thế danh sách hiện tại.\n" +
      "Nhấn Cancel để giữ dữ liệu cũ và thêm khách mới."
    );

    if (replace) {
      state.customers = imported;
    } else {
      state.customers.push(...imported);
    }

    state.fileName = file.name;
    state.syncedAt = dateText();

    saveData();
    render();

    el.importStatus.textContent =
      `Đã nhập ${imported.length} khách hàng thành công.`;
  } catch (error) {
    console.error("Lỗi nhập Excel:", error);
    alert("Không thể đọc file. Hãy kiểm tra định dạng Excel/CSV.");
    el.importStatus.textContent = "Nhập file thất bại.";
  }
}

/* =========================================
   12. XUẤT FILE CSV / EXCEL
========================================= */

function downloadBlob(blob, fileName) {
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");

  link.href = url;
  link.download = fileName;

  document.body.appendChild(link);
  link.click();
  link.remove();

  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

function exportCsv() {
  const customers = getFilteredCustomers();

  if (!customers.length) {
    alert("Không có dữ liệu để xuất.");
    return;
  }

  const headers = [
    "STT",
    "Khách Hàng",
    "Ngày Chốt",
    "Tiến Trình",
    "Hạn Chốt",
    "Doanh Số",
    "Loại Khách",
    "Ký Hiệu",
    "Giải Thích"
  ];

  const rows = customers.map(c => [
    c.stt,
    c.name,
    formatDateOnly(c.closeDate),
    formatProgress(c.progress),
    formatDateOnly(c.deadline),
    c.value,
    c.customerType,
    c.symbol,
    c.explanation
  ]);

  const csvCell = value =>
    '"' + String(value ?? "").replace(/"/g, '""') + '"';

  const csv = "\uFEFF" + [
    headers,
    ...rows
  ].map(row => row.map(csvCell).join(",")).join("\r\n");

  downloadBlob(
    new Blob([csv], { type: "text/csv;charset=utf-8;" }),
    "CRM_Khach_Hang.csv"
  );
}

function exportExcel() {
  const customers = getFilteredCustomers();

  if (!customers.length) {
    alert("Không có dữ liệu để xuất.");
    return;
  }

  if (typeof XLSX === "undefined") {
    alert("Không tải được thư viện Excel. Hãy kiểm tra Internet.");
    return;
  }

  const rows = customers.map(c => ({
    "STT": c.stt,
    "Khách Hàng": c.name,
    "Ngày Chốt": formatDateOnly(c.closeDate),
    "Tiến Trình": formatProgress(c.progress),
    "Hạn Chốt": formatDateOnly(c.deadline),
    "Doanh Số": c.value,
    "Loại Khách": c.customerType,
    "Ký Hiệu": c.symbol,
    "Giải Thích": c.explanation
  }));

  const worksheet = XLSX.utils.json_to_sheet(rows);
  const workbook = XLSX.utils.book_new();

  worksheet["!cols"] = [
    { wch: 8 },
    { wch: 25 },
    { wch: 15 },
    { wch: 20 },
    { wch: 15 },
    { wch: 18 },
    { wch: 18 },
    { wch: 15 },
    { wch: 35 }
  ];

  XLSX.utils.book_append_sheet(workbook, worksheet, "KhachHang");
  XLSX.writeFile(workbook, "CRM_Khach_Hang.xlsx");
}

/* =========================================
   13. ĐỌC SỐ TIỀN THÀNH CHỮ
========================================= */

function numberToVietnameseWords(number) {
  number = Math.floor(Math.abs(Number(number) || 0));

  if (number === 0) return "Không đồng";

  const digits = [
    "không", "một", "hai", "ba", "bốn",
    "năm", "sáu", "bảy", "tám", "chín"
  ];

  const units = [
    "", "nghìn", "triệu", "tỷ",
    "nghìn tỷ", "triệu tỷ"
  ];

  function readThreeDigits(n, full) {
    const hundred = Math.floor(n / 100);
    const ten = Math.floor((n % 100) / 10);
    const unit = n % 10;
    const words = [];

    if (hundred > 0 || full) {
      words.push(digits[hundred], "trăm");
    }

    if (ten > 1) {
      words.push(digits[ten], "mươi");

      if (unit === 1) words.push("mốt");
      else if (unit === 5) words.push("lăm");
      else if (unit > 0) words.push(digits[unit]);
    } else if (ten === 1) {
      words.push("mười");

      if (unit === 5) words.push("lăm");
      else if (unit > 0) words.push(digits[unit]);
    } else if (unit > 0) {
      if (hundred > 0 || full) words.push("lẻ");
      words.push(digits[unit]);
    }

    return words.join(" ");
  }

  const groups = [];
  let remaining = number;
  let index = 0;

  while (remaining > 0) {
    const group = remaining % 1000;

    if (group > 0) {
      groups.unshift({
        value: group,
        unit: units[index] || ""
      });
    }

    remaining = Math.floor(remaining / 1000);
    index++;
  }

  const words = groups.map((group, i) => {
    const full = i > 0 && group.value < 100;
    const text = readThreeDigits(group.value, full);

    return [text, group.unit].filter(Boolean).join(" ");
  });

  let result = words.join(" ").replace(/\s+/g, " ").trim();

  result = result.charAt(0).toUpperCase() + result.slice(1);

  return result + " đồng";
}

/* =========================================
   14. XUẤT BIÊN LAI
========================================= */

function exportReceipt(customerId) {
  const customer = state.customers.find(c => c.id === customerId);

  if (!customer) {
    alert("Không tìm thấy khách hàng.");
    return;
  }

  const amount = Number(customer.value) || 0;
  const now = new Date();

  const receiptNumber =
    "CSGN-" +
    now.getFullYear() +
    String(now.getMonth() + 1).padStart(2, "0") +
    String(now.getDate()).padStart(2, "0") +
    "-" +
    String(customer.stt || 1).padStart(4, "0");

  const receiptWindow = window.open(
    "",
    "_blank",
    "width=900,height=850"
  );

  if (!receiptWindow) {
    alert("Trình duyệt đã chặn cửa sổ biên lai. Hãy cho phép cửa sổ bật lên.");
    return;
  }

  const receiptHTML = `
<!doctype html>
<html lang="vi">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Biên lai ${escapeHTML(receiptNumber)}</title>

<style>
* { box-sizing: border-box; }

body {
  margin: 0;
  padding: 28px;
  color: #172033;
  background: #f0f2f5;
  font-family: Arial, Helvetica, sans-serif;
}

.receipt {
  width: 100%;
  max-width: 760px;
  margin: 0 auto;
  padding: 42px;
  background: white;
  border: 1px solid #d9dee7;
}

.header {
  display: flex;
  justify-content: space-between;
  gap: 20px;
  padding-bottom: 22px;
  border-bottom: 2px solid #146c94;
}

.brand {
  color: #146c94;
  font-size: 23px;
  font-weight: 900;
}

.muted {
  color: #667085;
  font-size: 12px;
  line-height: 1.6;
}

.receipt-no {
  text-align: right;
  font-size: 12px;
  line-height: 1.8;
}

h1 {
  margin: 28px 0 8px;
  text-align: center;
  font-size: 27px;
  letter-spacing: 2px;
}

.subtitle {
  text-align: center;
  color: #667085;
  font-size: 12px;
}

.info { margin-top: 28px; }

.info-row {
  display: grid;
  grid-template-columns: 175px 1fr;
  gap: 14px;
  padding: 11px 0;
  border-bottom: 1px dashed #d9dee7;
  font-size: 13px;
  line-height: 1.5;
}

.info-row span:first-child { color: #667085; }
.info-row strong { overflow-wrap: anywhere; }

.amount-box {
  margin-top: 24px;
  padding: 20px;
  border: 1px solid #b8dce9;
  border-radius: 10px;
  background: #f2f9fc;
}

.amount-box span {
  display: block;
  color: #475467;
  font-size: 12px;
}

.amount {
  margin: 8px 0;
  color: #146c94;
  font-size: 28px;
  font-weight: 900;
}

.amount-words {
  font-size: 12px;
  font-style: italic;
  line-height: 1.6;
}

.note {
  margin-top: 20px;
  padding: 12px;
  border-left: 3px solid #d87a2c;
  background: #fff8ef;
  color: #69410d;
  font-size: 11px;
  line-height: 1.6;
}

.signatures {
  display: grid;
  grid-template-columns: 1fr 1fr;
  gap: 35px;
  margin-top: 48px;
  text-align: center;
}

.signatures strong {
  display: block;
  font-size: 12px;
}

.signatures span {
  display: block;
  margin-top: 6px;
  color: #667085;
  font-size: 11px;
  font-style: italic;
}

.sign-space { height: 90px; }

.footer {
  margin-top: 30px;
  padding-top: 12px;
  border-top: 1px solid #e1e7ef;
  color: #667085;
  text-align: center;
  font-size: 10px;
  line-height: 1.6;
}

.toolbar {
  display: flex;
  justify-content: center;
  gap: 10px;
  margin: 0 auto 20px;
}

.toolbar button {
  padding: 11px 18px;
  border: 0;
  border-radius: 7px;
  background: #146c94;
  color: white;
  font-weight: bold;
  cursor: pointer;
}

.toolbar .close { background: #667085; }

@media print {
  body {
    padding: 0;
    background: white;
  }

  .receipt {
    max-width: none;
    padding: 15mm;
    border: 0;
  }

  .toolbar { display: none; }

  @page {
    size: A4;
    margin: 10mm;
  }
}

@media (max-width: 600px) {
  body { padding: 10px; }
  .receipt { padding: 20px; }
  .header { flex-direction: column; }
  .receipt-no { text-align: left; }
  .info-row { grid-template-columns: 1fr; gap: 4px; }
}
</style>
</head>

<body>

<div class="toolbar">
  <button onclick="window.print()">In / Lưu PDF</button>
  <button class="close" onclick="window.close()">Đóng</button>
</div>

<main class="receipt">

  <header class="header">
    <div>
      <div class="brand">CHỢ SALE GIA NGHĨA</div>
      <p class="muted">
        Hệ thống quản lý khách hàng<br>
        Biên lai nội bộ
      </p>
    </div>

    <div class="receipt-no">
      <strong>Mã biên lai</strong><br>
      ${escapeHTML(receiptNumber)}<br>
      <span class="muted">Ngày lập: ${escapeHTML(dateText())}</span>
    </div>
  </header>

  <h1>BIÊN LAI</h1>
  <div class="subtitle">Xác nhận thông tin giao dịch nội bộ</div>

  <section class="info">

    <div class="info-row">
      <span>Khách hàng</span>
      <strong>${escapeHTML(customer.name)}</strong>
    </div>

    <div class="info-row">
      <span>Mã khách hàng / STT</span>
      <strong>${escapeHTML(customer.stt || "—")}</strong>
    </div>

    <div class="info-row">
      <span>Loại khách</span>
      <strong>${escapeHTML(customer.customerType || "—")}</strong>
    </div>

    <div class="info-row">
      <span>Ký hiệu</span>
      <strong>${escapeHTML(customer.symbol || "—")}</strong>
    </div>

    <div class="info-row">
      <span>Ngày chốt</span>
      <strong>${escapeHTML(formatDateOnly(customer.closeDate) || "—")}</strong>
    </div>

    <div class="info-row">
      <span>Tiến trình</span>
      <strong>${escapeHTML(formatProgress(customer.progress) || "—")}</strong>
    </div>

    <div class="info-row">
      <span>Hạn chốt</span>
      <strong>${escapeHTML(formatDateOnly(customer.deadline) || "Không hạn")}</strong>
    </div>

    <div class="info-row">
      <span>Nội dung / giải thích</span>
      <strong>${escapeHTML(customer.explanation || "Không có ghi chú")}</strong>
    </div>

  </section>

  <section class="amount-box">
    <span>Tổng số tiền</span>
    <div class="amount">${formatMoney(amount)}</div>
    <div class="amount-words">
      Bằng chữ: ${escapeHTML(numberToVietnameseWords(amount))}.
    </div>
  </section>

  <div class="note">
    <strong>Lưu ý:</strong>
    Đây là mẫu biên lai nội bộ được tạo từ dữ liệu CRM.
    Biên lai không thay thế hóa đơn điện tử hoặc chứng từ thuế.
    Chỉ xác nhận thanh toán khi giao dịch thực tế đã được kiểm tra.
  </div>

  <section class="signatures">
    <div>
      <strong>NGƯỜI NỘP TIỀN</strong>
      <span>(Ký và ghi rõ họ tên)</span>
      <div class="sign-space"></div>
    </div>

    <div>
      <strong>NGƯỜI THU TIỀN</strong>
      <span>(Ký và ghi rõ họ tên)</span>
      <div class="sign-space"></div>
    </div>
  </section>

  <footer class="footer">
    Chợ Sale Gia Nghĩa · Biên lai được tạo từ hệ thống CRM<br>
    Mã biên lai: ${escapeHTML(receiptNumber)}
  </footer>

</main>
</body>
</html>
`;

  receiptWindow.document.open();
  receiptWindow.document.write(receiptHTML);
  receiptWindow.document.close();
}

/* =========================================
   15. GẮN SỰ KIỆN
========================================= */

function bindEvents() {
  el.syncBtn.addEventListener("click", syncFromExcelFile);
  el.excelFile.addEventListener("change", syncFromExcelFile);

  el.exportBtn.addEventListener("click", exportCsv);
  el.exportExcelBtn.addEventListener("click", exportExcel);

  el.addCustomerBtn.addEventListener("click", () => {
    openCustomerForm();
  });

  el.closeDialogBtn.addEventListener("click", closeDialog);
  el.cancelDialogBtn.addEventListener("click", closeDialog);

  el.form.addEventListener("submit", saveCustomer);
  el.rows.addEventListener("click", handleRowAction);

  el.search.addEventListener("input", event => {
    state.query = normalizeSearch(event.target.value);
    renderRows();
    renderSummary();
  });

  el.typeFilter.addEventListener("change", event => {
    state.type = event.target.value;
    renderRows();
    renderSummary();
  });

  el.symbolFilter.addEventListener("change", event => {
    state.symbol = event.target.value;
    renderRows();
    renderSummary();
  });

  el.deadlineFilter.addEventListener("change", event => {
    state.deadline = event.target.value;
    syncQuickButtons();
    renderRows();
    renderSummary();
  });

  el.clearFilter.addEventListener("click", () => {
    state.query = "";
    state.type = "";
    state.symbol = "";
    state.deadline = "";

    el.search.value = "";
    el.typeFilter.value = "";
    el.symbolFilter.value = "";
    el.deadlineFilter.value = "";

    render();
  });

  document.querySelectorAll(".quick-btn").forEach(button => {
    button.addEventListener("click", () => {
      state.deadline = button.dataset.deadline || "";
      el.deadlineFilter.value = state.deadline;

      syncQuickButtons();
      renderRows();
      renderSummary();
    });
  });
}

/* =========================================
   16. KHỞI ĐỘNG ỨNG DỤNG
========================================= */

restoreData();
bindEvents();
render();
