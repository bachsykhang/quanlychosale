const STORAGE_KEY = "crm_excel_sync_v3";

const COLUMNS = [
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

const state = {
  customers: [],
  query: "",
  type: "",
  symbol: "",
  deadline: "",
  fileName: "",
  syncedAt: "",
  dirty: false
};

let searchTimer = null;

const el = {
  excelFile: document.querySelector("#excelFile"),
  syncBtn: document.querySelector("#syncExcelBtn"),

  exportBtn: document.querySelector("#exportBtn"),
  exportExcelBtn: document.querySelector("#exportExcelBtn"),
  addCustomerBtn: document.querySelector("#addCustomerBtn"),

  dialog: document.querySelector("#customerDialog"),
  form: document.querySelector("#customerForm"),
  dialogTitle: document.querySelector("#dialogTitle"),
  closeDialogBtn: document.querySelector("#closeDialogBtn"),
  cancelDialogBtn: document.querySelector("#cancelDialogBtn"),

  rows: document.querySelector("#customerRows"),
  empty: document.querySelector("#emptyState"),

  total: document.querySelector("#totalCustomers"),
  newCustomers: document.querySelector("#newCustomers"),
  revenue: document.querySelector("#expectedRevenue"),
  follow: document.querySelector("#needFollowUp"),

  search: document.querySelector("#searchInput"),
  typeFilter: document.querySelector("#typeFilter"),
  symbolFilter: document.querySelector("#symbolFilter"),
  deadlineFilter: document.querySelector("#deadlineFilter"),
  clearFilter: document.querySelector("#clearFilterBtn"),

  resultCount: document.querySelector("#resultCount"),
  fileInfo: document.querySelector("#fileInfo"),

  pipeline: document.querySelector("#pipelineGrid"),
  lastUpdated: document.querySelector("#lastUpdated"),

  importStatus: document.querySelector("#importStatus")
};

restoreData();
bindEvents();
render();


/* =========================================================
   EVENTS
========================================================= */

function bindEvents() {

  el.syncBtn.addEventListener(
    "click",
    syncFromExcelFile
  );

  el.excelFile.addEventListener(
    "change",
    syncFromExcelFile
  );

  el.exportBtn.addEventListener(
    "click",
    exportCsv
  );

  el.exportExcelBtn.addEventListener(
    "click",
    exportExcel
  );

  el.addCustomerBtn.addEventListener(
    "click",
    () => openCustomerForm()
  );

  el.closeDialogBtn.addEventListener(
    "click",
    closeDialog
  );

  el.cancelDialogBtn.addEventListener(
    "click",
    closeDialog
  );

  el.form.addEventListener(
    "submit",
    saveCustomer
  );

  el.clearFilter.addEventListener(
    "click",
    clearFilters
  );

  el.rows.addEventListener(
    "click",
    handleRowAction
  );


  /* TÌM KIẾM */

  el.search.addEventListener(
    "input",
    event => {

      clearTimeout(searchTimer);

      searchTimer = setTimeout(() => {

        state.query = normalizeSearch(
          event.target.value
        );

        render();

      }, 120);

    }
  );


  /* LỌC LOẠI KHÁCH */

  el.typeFilter.addEventListener(
    "change",
    event => {

      state.type = event.target.value;

      render();

    }
  );


  /* LỌC KÝ HIỆU */

  el.symbolFilter.addEventListener(
    "change",
    event => {

      state.symbol = event.target.value;

      render();

    }
  );


  /* LỌC HẠN CHỐT */

  el.deadlineFilter.addEventListener(
    "change",
    event => {

      state.deadline = event.target.value;

      render();

    }
  );
}


/* =========================================================
   STORAGE
========================================================= */

function restoreData() {

  try {

    const saved =
      JSON.parse(
        localStorage.getItem(STORAGE_KEY)
      );

    if (!saved) {
      return;
    }

    state.customers =
      Array.isArray(saved.customers)
        ? saved.customers
        : [];

    state.fileName =
      saved.fileName || "";

    state.syncedAt =
      saved.syncedAt || "";

    state.dirty =
      Boolean(saved.dirty);

  } catch {

    state.customers = [];

  }
}


function persistData() {

  localStorage.setItem(
    STORAGE_KEY,
    JSON.stringify({
      customers: state.customers,
      fileName: state.fileName,
      syncedAt: state.syncedAt,
      dirty: state.dirty
    })
  );

}


/* =========================================================
   IMPORT EXCEL
========================================================= */

async function syncFromExcelFile() {

  const file =
    el.excelFile.files?.[0];

  if (!file) {

    setImportStatus(
      "Vui lòng chọn file Excel trước khi đồng bộ.",
      true
    );

    return;
  }

  setImportStatus(
    `Đang đọc file ${file.name}...`
  );

  try {

    const rows =
      await readWorkbookRows(file);

    const customers =
      rowsToCustomers(rows);

    if (!customers.length) {

      throw new Error(
        "Không tìm thấy dòng khách hàng hợp lệ trong file."
      );

    }

    state.customers = customers;

    state.fileName = file.name;

    state.syncedAt =
      new Date().toISOString();

    state.dirty = false;

    persistData();

    clearFilters(false);

    setImportStatus(
      `Đã đồng bộ ${customers.length} khách hàng từ file Excel.`
    );

    render();

  } catch (error) {

    setImportStatus(
      error.message,
      true
    );

  }
}


function readWorkbookRows(file) {

  const extension =
    file.name
      .split(".")
      .pop()
      .toLowerCase();

  if (extension === "csv") {

    return readTextFile(file)
      .then(parseCsvRows);

  }

  if (!window.XLSX) {

    return Promise.reject(
      new Error(
        "Chưa tải được thư viện đọc Excel. Hãy kiểm tra kết nối mạng rồi mở lại trang."
      )
    );

  }

  return readArrayBuffer(file)
    .then(buffer => {

      const workbook =
        XLSX.read(
          buffer,
          {
            type: "array",
            cellDates: false
          }
        );

      const sheetName =
        workbook.SheetNames[0];

      const sheet =
        workbook.Sheets[sheetName];

      return XLSX.utils.sheet_to_json(
        sheet,
        {
          header: 1,
          raw: false,
          defval: "",
          blankrows: false
        }
      );

    });

}


function readArrayBuffer(file) {

  return new Promise(
    (resolve, reject) => {

      const reader =
        new FileReader();

      reader.onload = () =>
        resolve(reader.result);

      reader.onerror = () =>
        reject(
          new Error(
            "Không đọc được file Excel."
          )
        );

      reader.readAsArrayBuffer(file);

    }
  );

}


function readTextFile(file) {

  return new Promise(
    (resolve, reject) => {

      const reader =
        new FileReader();

      reader.onload = () =>
        resolve(reader.result);

      reader.onerror = () =>
        reject(
          new Error(
            "Không đọc được file CSV."
          )
        );

      reader.readAsText(
        file,
        "utf-8"
      );

    }
  );

}


/* =========================================================
   CSV
========================================================= */

function parseCsvRows(text) {

  const rows = [];

  let row = [];
  let cell = "";
  let quoted = false;

  for (
    let index = 0;
    index < text.length;
    index += 1
  ) {

    const char =
      text[index];

    const next =
      text[index + 1];

    if (
      char === '"' &&
      quoted &&
      next === '"'
    ) {

      cell += '"';
      index += 1;

    }

    else if (char === '"') {

      quoted = !quoted;

    }

    else if (
      char === "," &&
      !quoted
    ) {

      row.push(cell);
      cell = "";

    }

    else if (
      (char === "\n" ||
       char === "\r") &&
      !quoted
    ) {

      if (
        char === "\r" &&
        next === "\n"
      ) {
        index += 1;
      }

      row.push(cell);

      if (
        row.some(
          value =>
            String(value).trim() !== ""
        )
      ) {
        rows.push(row);
      }

      row = [];
      cell = "";

    }

    else {

      cell += char;

    }
  }

  row.push(cell);

  if (
    row.some(
      value =>
        String(value).trim() !== ""
    )
  ) {
    rows.push(row);
  }

  return rows;
}


/* =========================================================
   CONVERT EXCEL -> CUSTOMER
========================================================= */

function rowsToCustomers(rows) {

  const headerIndex =
    findHeaderIndex(rows);

  if (headerIndex < 0) {

    throw new Error(
      `File cần có hàng tiêu đề gồm: ${COLUMNS.join(", ")}.`
    );

  }

  const headers =
    rows[headerIndex]
      .map(normalizeKey);

  return rows
    .slice(headerIndex + 1)
    .map(
      (row, index) =>
        normalizeCustomer(
          rowToRecord(headers, row),
          index
        )
    )
    .filter(
      customer =>
        customer.name &&
        normalizeKey(customer.name) !==
          "khach_hang"
    );

}


function findHeaderIndex(rows) {

  return rows.findIndex(row => {

    const keys =
      row.map(normalizeKey);

    return (
      keys.includes("stt") &&
      keys.includes("khach_hang")
    );

  });

}


function rowToRecord(headers, row) {

  return row.reduce(
    (record, value, index) => {

      record[
        headers[index] ||
        `cot_${index}`
      ] =
        String(
          value ?? ""
        ).trim();

      return record;

    },
    {}
  );

}


function normalizeCustomer(
  record,
  index
) {

  const find = keys => {

    const key =
      keys.find(
        item =>
          record[item] !== undefined &&
          record[item] !== ""
      );

    return key
      ? record[key]
      : "";

  };

  return {

    id:
      crypto.randomUUID(),

    stt:
      find([
        "stt",
        "so_thu_tu",
        "thu_tu",
        "cot_0"
      ]) ||
      String(index + 1),

    name:
      find([
        "khach_hang",
        "ten_khach_hang",
        "ho_ten",
        "ten",
        "cot_1"
      ]),

    closeDate:
      find([
        "ngay_chot",
        "chot_ngay",
        "close_date",
        "cot_2"
      ]),

    progress:
      find([
        "tien_trinh",
        "progress",
        "cot_3"
      ]),

    deadline:
      find([
        "han_chot",
        "deadline",
        "cot_4"
      ]),

    value:
      parseMoney(
        find([
          "doanh_so",
          "doanh_thu",
          "gia_tri",
          "amount",
          "cot_5"
        ])
      ),

    customerType:
      find([
        "loai_khach",
        "type",
        "customer_type",
        "cot_6"
      ]) ||
      "Mới",

    symbol:
      find([
        "ky_hieu",
        "symbol",
        "tag",
        "cot_7"
      ]),

    explanation:
      find([
        "giai_thich",
        "explanation",
        "description",
        "cot_8"
      ])
  };
}


/* =========================================================
   RENDER
========================================================= */

function render() {

  const customers =
    getFilteredCustomers();

  renderRows(customers);
  renderMetrics();
  renderFilters();
  renderPipeline();
  renderSyncInfo();

}


/* =========================================================
   FILTER
========================================================= */

function getFilteredCustomers() {

  return state.customers.filter(
    customer => {

      const haystack =
        normalizeSearch(
          customerToRow(customer).join(" ")
        );

      const matchesQuery =
        !state.query ||
        haystack.includes(state.query);

      const matchesType =
        !state.type ||
        customer.customerType === state.type;

      const matchesSymbol =
        !state.symbol ||
        customer.symbol === state.symbol;

      let matchesDeadline = true;


      switch (state.deadline) {

        /* Không có hạn chốt */

        case "none":

          matchesDeadline =
            !String(
              customer.deadline || ""
            ).trim() ||
            String(
              customer.deadline || ""
            ).trim() === "0";

          break;


        /* Có hạn chốt */

        case "has":

          matchesDeadline =
            String(
              customer.deadline || ""
            ).trim() !== "" &&
            String(
              customer.deadline || ""
            ).trim() !== "0";

          break;


        /* Sắp đến hạn */

        case "urgent":

          matchesDeadline =
            isDeadlineUrgent(
              customer.deadline
            ) &&
            !normalizeKey(
              customer.deadline
            ).includes("qua_han");

          break;


        /* Quá hạn */

        case "overdue":

          matchesDeadline =
            normalizeKey(
              customer.deadline
            ).includes("qua_han");

          break;

      }

      return (
        matchesQuery &&
        matchesType &&
        matchesSymbol &&
        matchesDeadline
      );

    }
  );

}


/* =========================================================
   RENDER ROWS
========================================================= */

function renderRows(customers) {

  el.resultCount.textContent =
    `${customers.length} / ${state.customers.length} khách hàng`;

  el.rows.innerHTML =
    customers
      .map(
        customer => `
          <tr>

            <td class="center">
              ${escapeHtml(
                customer.stt || "-"
              )}
            </td>

            <td>
              <strong>
                ${escapeHtml(
                  customer.name || "-"
                )}
              </strong>
            </td>

            <td>
              ${escapeHtml(
                customer.closeDate || "-"
              )}
            </td>

            <td>
              ${renderProgress(
                customer.progress
              )}
            </td>

            <td>
              ${renderDeadline(
                customer.deadline
              )}
            </td>

            <td>
              ${formatMoney(
                customer.value
              )}
            </td>

            <td>
              ${renderCustomerType(
                customer.customerType
              )}
            </td>

            <td>
              ${escapeHtml(
                customer.symbol || "-"
              )}
            </td>

            <td>
              ${escapeHtml(
                customer.explanation || "-"
              )}
            </td>

            <td>

              <div class="row-actions">

                <button
                  type="button"
                  data-action="edit"
                  data-id="${customer.id}"
                >
                  Sửa
                </button>

                <button
                  type="button"
                  data-action="delete"
                  data-id="${customer.id}"
                >
                  Xóa
                </button>

              </div>

            </td>

          </tr>
        `
      )
      .join("");

  el.empty.style.display =
    customers.length
      ? "none"
      : "block";
}


/* =========================================================
   METRICS
========================================================= */

function renderMetrics() {

  const revenue =
    state.customers.reduce(
      (sum, customer) =>
        sum +
        Number(customer.value || 0),
      0
    );

  el.total.textContent =
    state.customers.length;

  el.newCustomers.textContent =
    state.customers.filter(
      customer =>
        normalizeKey(
          customer.customerType
        ).includes("moi")
    ).length;

  el.revenue.textContent =
    formatMoney(revenue);

  el.follow.textContent =
    state.customers.filter(
      customer =>
        isDeadlineUrgent(
          customer.deadline
        )
    ).length;
}


/* =========================================================
   FILTER OPTIONS
========================================================= */

function renderFilters() {

  fillSelect(
    el.typeFilter,
    "Tất cả loại khách",
    uniqueValues(
      state.customers.map(
        customer =>
          customer.customerType
      )
    ),
    state.type
  );

  fillSelect(
    el.symbolFilter,
    "Tất cả ký hiệu",
    uniqueValues(
      state.customers.map(
        customer =>
          customer.symbol
      )
    ),
    state.symbol
  );

  el.deadlineFilter.value =
    state.deadline;

}


function fillSelect(
  select,
  label,
  values,
  current
) {

  select.innerHTML =
    `<option value="">
      ${label}
    </option>` +

    values
      .map(
        value =>
          `<option value="${escapeHtml(value)}">
            ${escapeHtml(value)}
          </option>`
      )
      .join("");

  select.value =
    values.includes(current)
      ? current
      : "";
}


/* =========================================================
   PIPELINE
========================================================= */

function renderPipeline() {

  const groups =
    uniqueValues(
      state.customers.map(
        customer =>
          customer.customerType
      )
    );

  const total =
    Math.max(
      state.customers.length,
      1
    );

  el.pipeline.innerHTML =
    (
      groups.length
        ? groups
        : ["Chưa có dữ liệu"]
    )
      .map(type => {

        const count =
          state.customers.filter(
            customer =>
              customer.customerType === type
          ).length;

        const percent =
          Math.round(
            (count / total) * 100
          );

        return `
          <div class="pipe-item">

            <strong>
              ${escapeHtml(type)}
            </strong>

            <span>
              ${count} khách - ${percent}%
            </span>

            <div class="bar">
              <span
                style="width:${percent}%"
              ></span>
            </div>

          </div>
        `;

      })
      .join("");
}


/* =========================================================
   SYNC INFO
========================================================= */

function renderSyncInfo() {

  if (!state.syncedAt) {

    el.fileInfo.textContent =
      state.dirty
        ? "Đã có dữ liệu mới, hãy tải Excel đã cập nhật."
        : "Chưa đồng bộ file Excel.";

    el.lastUpdated.textContent =
      "Chưa cập nhật";

    return;
  }

  const syncedAt =
    new Date(
      state.syncedAt
    ).toLocaleString("vi-VN");

  const dirtyText =
    state.dirty
      ? " - Có thay đổi chưa xuất Excel"
      : "";

  el.fileInfo.textContent =
    `Nguồn dữ liệu: ${
      state.fileName || "Excel"
    }${dirtyText}`;

  el.lastUpdated.textContent =
    `Đồng bộ lúc ${syncedAt}`;
}


/* =========================================================
   EDIT / DELETE
========================================================= */

function handleRowAction(event) {

  const button =
    event.target.closest(
      "button[data-action]"
    );

  if (!button) {
    return;
  }

  if (
    button.dataset.action === "edit"
  ) {

    openCustomerForm(
      button.dataset.id
    );

  }

  if (
    button.dataset.action === "delete"
  ) {

    deleteCustomer(
      button.dataset.id
    );

  }
}


function openCustomerForm(id = "") {

  const customer =
    state.customers.find(
      item => item.id === id
    );

  el.dialogTitle.textContent =
    customer
      ? "Sửa dữ liệu"
      : "Thêm dữ liệu";

  setFormValue(
    "#customerId",
    customer?.id || ""
  );

  setFormValue(
    "#sttInput",
    customer?.stt || nextStt()
  );

  setFormValue(
    "#nameInput",
    customer?.name || ""
  );

  setFormValue(
    "#closeDateInput",
    customer?.closeDate || ""
  );

  setFormValue(
    "#progressInput",
    customer?.progress || ""
  );

  setFormValue(
    "#deadlineInput",
    customer?.deadline || ""
  );

  setFormValue(
    "#valueInput",
    customer?.value
      ? formatPlainMoney(customer.value)
      : ""
  );

  setFormValue(
    "#customerTypeInput",
    customer?.customerType || "Mới"
  );

  setFormValue(
    "#symbolInput",
    customer?.symbol || ""
  );

  setFormValue(
    "#explanationInput",
    customer?.explanation || ""
  );

  el.dialog.showModal();
}


function closeDialog() {

  el.dialog.close();

  el.form.reset();

}


function saveCustomer(event) {

  event.preventDefault();

  const id =
    document.querySelector(
      "#customerId"
    ).value ||
    crypto.randomUUID();

  const customer = {

    id,

    stt:
      document.querySelector(
        "#sttInput"
      ).value.trim() ||
      nextStt(),

    name:
      document.querySelector(
        "#nameInput"
      ).value.trim(),

    closeDate:
      document.querySelector(
        "#closeDateInput"
      ).value.trim(),

    progress:
      document.querySelector(
        "#progressInput"
      ).value.trim(),

    deadline:
      document.querySelector(
        "#deadlineInput"
      ).value.trim(),

    value:
      parseMoney(
        document.querySelector(
          "#valueInput"
        ).value
      ),

    customerType:
      document.querySelector(
        "#customerTypeInput"
      ).value.trim() ||
      "Mới",

    symbol:
      document.querySelector(
        "#symbolInput"
      ).value.trim(),

    explanation:
      document.querySelector(
        "#explanationInput"
      ).value.trim()
  };

  const index =
    state.customers.findIndex(
      item => item.id === id
    );

  if (index >= 0) {

    state.customers[index] =
      customer;

  } else {

    state.customers.push(
      customer
    );

  }

  markChanged(
    "Đã lưu dữ liệu. Bấm “Tải Excel đã cập nhật” để đồng bộ ra file Excel."
  );

  closeDialog();
}


function deleteCustomer(id) {

  const customer =
    state.customers.find(
      item => item.id === id
    );

  if (
    !customer ||
    !confirm(
      `Xóa khách hàng "${customer.name}"?`
    )
  ) {
    return;
  }

  state.customers =
    state.customers.filter(
      item => item.id !== id
    );

  markChanged(
    "Đã xóa dữ liệu. Bấm “Tải Excel đã cập nhật” để đồng bộ ra file Excel."
  );
}


/* =========================================================
   CHANGE
========================================================= */

function markChanged(message) {

  state.dirty = true;

  state.syncedAt =
    state.syncedAt ||
    new Date().toISOString();

  persistData();

  render();

  setImportStatus(
    message
  );
}


/* =========================================================
   CLEAR FILTER
========================================================= */

function clearFilters(
  shouldRender = true
) {

  state.query = "";
  state.type = "";
  state.symbol = "";
  state.deadline = "";

  el.search.value = "";
  el.typeFilter.value = "";
  el.symbolFilter.value = "";
  el.deadlineFilter.value = "";

  if (shouldRender) {
    render();
  }
}


/* =========================================================
   EXPORT CSV
========================================================= */

function exportCsv() {

  if (!state.customers.length) {

    setImportStatus(
      "Chưa có dữ liệu để xuất CSV.",
      true
    );

    return;
  }

  const csv =
    [
      COLUMNS,
      ...state.customers.map(
        customerToRow
      )
    ]
      .map(
        row =>
          row
            .map(csvCell)
            .join(",")
      )
      .join("\n");

  const blob =
    new Blob(
      ["\ufeff" + csv],
      {
        type:
          "text/csv;charset=utf-8"
      }
    );

  downloadBlob(
    blob,
    `khach-hang-${todayStamp()}.csv`
  );
}


/* =========================================================
   EXPORT EXCEL
========================================================= */

function exportExcel() {

  if (!state.customers.length) {

    setImportStatus(
      "Chưa có dữ liệu để xuất Excel.",
      true
    );

    return;
  }

  if (!window.XLSX) {

    setImportStatus(
      "Chưa tải được thư viện xuất Excel. Hãy kiểm tra kết nối mạng rồi mở lại trang.",
      true
    );

    return;
  }

  const worksheet =
    XLSX.utils.aoa_to_sheet(
      [
        COLUMNS,
        ...state.customers.map(
          customerToRow
        )
      ]
    );

  worksheet["!cols"] = [
    { wch: 8 },
    { wch: 24 },
    { wch: 12 },
    { wch: 12 },
    { wch: 12 },
    { wch: 14 },
    { wch: 14 },
    { wch: 12 },
    { wch: 34 }
  ];

  const workbook =
    XLSX.utils.book_new();

  XLSX.utils.book_append_sheet(
    workbook,
    worksheet,
    "KhachHang"
  );

  const fileName =
    `khach-hang-da-cap-nhat-${todayStamp()}.xlsx`;

  XLSX.writeFile(
    workbook,
    fileName
  );

  state.dirty = false;
  state.fileName = fileName;

  state.syncedAt =
    new Date().toISOString();

  persistData();

  render();

  setImportStatus(
    "Đã tải file Excel đã cập nhật."
  );
}


/* =========================================================
   CUSTOMER ROW
========================================================= */

function customerToRow(customer) {

  return [
    customer.stt,
    customer.name,
    customer.closeDate,
    customer.progress,
    customer.deadline,
    Number(customer.value || 0),
    customer.customerType,
    customer.symbol,
    customer.explanation
  ];

}


/* =========================================================
   UNIQUE
========================================================= */

function uniqueValues(values) {

  return [
    ...new Set(
      values
        .map(
          value =>
            String(
              value || ""
            ).trim()
        )
        .filter(Boolean)
    )
  ].sort(
    (a, b) =>
      a.localeCompare(
        b,
        "vi"
      )
  );
}


/* =========================================================
   PROGRESS
========================================================= */

function renderProgress(value) {

  const text =
    String(
      value || ""
    ).trim();

  if (
    !text ||
    text === "0"
  ) {
    return "-";
  }

  const percent =
    ratioToPercent(text);

  return `
    <div class="ratio-cell">

      <strong>
        ${escapeHtml(text)}
      </strong>

      ${
        percent > 0
          ? `
            <div class="mini-bar">
              <span
                style="width:${percent}%"
              ></span>
            </div>
          `
          : ""
      }

    </div>
  `;
}


/* =========================================================
   DEADLINE
========================================================= */

function renderDeadline(value) {

  const text =
    String(
      value || ""
    ).trim();

  if (
    !text ||
    text === "0"
  ) {

    return `
      <span class="soft-tag">
        0
      </span>
    `;

  }

  const isUrgent =
    isDeadlineUrgent(text);

  return `
    <span
      class="deadline-tag ${
        isUrgent
          ? "urgent"
          : ""
      }"
    >
      ${escapeHtml(text)}
    </span>
  `;
}


/* =========================================================
   CUSTOMER TYPE
========================================================= */

function renderCustomerType(value) {

  const text =
    String(
      value || "Mới"
    ).trim();

  const key =
    normalizeKey(text);

  const className =
    key.includes("vip")
      ? "vip"
      : key.includes("cu")
        ? "old"
        : "new";

  return `
    <span
      class="type-tag ${className}"
    >
      ${escapeHtml(text)}
    </span>
  `;
}


/* =========================================================
   RATIO
========================================================= */

function ratioToPercent(value) {

  const match =
    String(value).match(
      /(\d+(?:[.,]\d+)?)\s*\/\s*(\d+(?:[.,]\d+)?)/
    );

  if (!match) {
    return 0;
  }

  const current =
    Number(
      match[1].replace(",", ".")
    );

  const total =
    Number(
      match[2].replace(",", ".")
    );

  if (!total) {
    return 0;
  }

  return Math.max(
    0,
    Math.min(
      100,
      Math.round(
        (current / total) * 100
      )
    )
  );
}


/* =========================================================
   DEADLINE URGENT
========================================================= */

function isDeadlineUrgent(value) {

  const text =
    String(
      value || ""
    ).trim();

  if (
    !text ||
    text === "0"
  ) {
    return false;
  }

  return (
    ratioToPercent(text) >= 75 ||
    normalizeKey(text).includes(
      "qua_han"
    )
  );
}


/* =========================================================
   MONEY
========================================================= */

function parseMoney(value) {

  const text =
    String(
      value || ""
    ).trim();

  const normalized =
    text
      .replace(/\s/g, "")
      .replace(
        /\.(?=\d{3}(\D|$))/g,
        ""
      )
      .replace(
        /,(?=\d{3}(\D|$))/g,
        ""
      )
      .replace(",", ".")
      .replace(
        /[^\d.-]/g,
        ""
      );

  return Number(
    normalized || 0
  );
}


/* =========================================================
   SEARCH
========================================================= */

function normalizeSearch(value) {

  return normalizeKey(value)
    .replace(
      /_/g,
      " "
    );
}


function normalizeKey(value) {

  return String(
    value || ""
  )
    .trim()
    .toLowerCase()
    .normalize("NFD")
    .replace(
      /[\u0300-\u036f]/g,
      ""
    )
    .replace(
      /đ/g,
      "d"
    )
    .replace(
      /[^a-z0-9]+/g,
      "_"
    )
    .replace(
      /^_|_$/g,
      "");
}


/* =========================================================
   FORM
========================================================= */

function setFormValue(
  selector,
  value
) {

  document.querySelector(
    selector
  ).value = value;

}


function nextStt() {

  const max =
    state.customers.reduce(
      (value, customer) =>
        Math.max(
          value,
          Number(
            customer.stt
          ) || 0
        ),
      0
    );

  return String(
    max + 1
  );
}


/* =========================================================
   STATUS
========================================================= */

function setImportStatus(
  message,
  isError = false
) {

  el.importStatus.textContent =
    message;

  el.importStatus.style.color =
    isError
      ? "#ffd0d0"
      : "#aebfca";
}


/* =========================================================
   MONEY FORMAT
========================================================= */

function formatMoney(value) {

  return new Intl.NumberFormat(
    "vi-VN",
    {
      style: "currency",
      currency: "VND",
      maximumFractionDigits: 0
    }
  ).format(
    Number(value || 0)
  );
}


function formatPlainMoney(value) {

  return new Intl.NumberFormat(
    "vi-VN",
    {
      maximumFractionDigits: 0
    }
  ).format(
    Number(value || 0)
  );
}


/* =========================================================
   DATE
========================================================= */

function todayStamp() {

  return new Date()
    .toISOString()
    .slice(0, 10);
}


/* =========================================================
   DOWNLOAD
========================================================= */

function downloadBlob(
  blob,
  fileName
) {

  const link =
    document.createElement("a");

  link.href =
    URL.createObjectURL(blob);

  link.download =
    fileName;

  link.click();

  URL.revokeObjectURL(
    link.href
  );
}


/* =========================================================
   CSV CELL
========================================================= */

function csvCell(value) {

  return `"${String(
    value ?? ""
  ).replace(
    /"/g,
    '""'
  )}"`;
}


/* =========================================================
   ESCAPE HTML
========================================================= */

function escapeHtml(value) {

  return String(
    value ?? ""
  )
    .replace(
      /&/g,
      "&amp;"
    )
    .replace(
      /</g,
      "&lt;"
    )
    .replace(
      />/g,
      "&gt;"
    )
    .replace(
      /"/g,
      "&quot;"
    )
    .replace(
      /'/g,
      "&#039;"
    );
}