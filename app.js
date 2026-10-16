"use strict";

const STORAGE_KEY = "choSaleGiaNghiaCRM_v33";

let customers = [];
let toastTimer;

const $ = id => document.getElementById(id);

const todayISO = () => {
  const d = new Date();
  d.setMinutes(d.getMinutes() - d.getTimezoneOffset());
  return d.toISOString().slice(0, 10);
};

function esc(value) {
  return String(value ?? "").replace(/[&<>"']/g, c => ({
    "&": "&amp;",
    "<": "&lt;",
    ">": "&gt;",
    '"': "&quot;",
    "'": "&#39;"
  })[c]);
}

function makeId() {
  return "KH-" +
    Date.now().toString(36).toUpperCase() +
    "-" +
    Math.random().toString(36).slice(2, 6).toUpperCase();
}

/* CHUẨN HÓA NGÀY */

function dateISO(value) {
  if (value === 0 || value === "0") {
    return "0";
  }

  if (value === null || value === undefined || value === "") {
    return "";
  }

  if (value instanceof Date && !isNaN(value)) {
    return `${value.getFullYear()}-${String(value.getMonth() + 1).padStart(2, "0")}-${String(value.getDate()).padStart(2, "0")}`;
  }

  if (
    typeof value === "number" &&
    Number.isFinite(value) &&
    typeof XLSX !== "undefined"
  ) {
    const parts = XLSX.SSF.parse_date_code(value);

    if (parts && parts.y && parts.m && parts.d) {
      return `${parts.y}-${String(parts.m).padStart(2, "0")}-${String(parts.d).padStart(2, "0")}`;
    }
  }

  const s = String(value).trim();

  let m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);

  if (m) {
    return `${m[1]}-${m[2].padStart(2, "0")}-${m[3].padStart(2, "0")}`;
  }

  m = s.match(/^(\d{1,2})[./-](\d{1,2})[./-](\d{4})$/);

  if (m) {
    return `${m[3]}-${m[2].padStart(2, "0")}-${m[1].padStart(2, "0")}`;
  }

  m = s.match(/^(\d{1,2})[./-](\d{1,2})$/);

  if (m) {
    return `${new Date().getFullYear()}-${m[2].padStart(2, "0")}-${m[1].padStart(2, "0")}`;
  }

  const d = new Date(s);

  if (isNaN(d)) {
    return "";
  }

  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
}

function dateText(value) {
  // Giá trị 0 phải giữ nguyên, không chuyển thành ngày.
  if (value === 0 || value === "0") {
    return "0";
  }

  const raw = String(value ?? "").trim();

  if (raw === "0") {
    return "0";
  }

  const date = dateISO(value);

  if (date && date !== "0") {
    const [year, month, day] = date.split("-");
    return `${day}/${month}/${year}`;
  }

  return raw || "—";
}

function money(value) {
  return new Intl.NumberFormat("vi-VN", {
    style: "currency",
    currency: "VND",
    maximumFractionDigits: 0
  }).format(Number(value) || 0);
}

function normalizeType(value) {
  const s = String(value ?? "").toLowerCase();

  if (s.includes("vip")) {
    return "Cũ (Vip)";
  }

  if (s === "cũ" || s === "cu" || s.includes("khách cũ")) {
    return "Cũ";
  }

  return "Mới";
}

function normalizeSymbol(value) {
  const s = String(value ?? "lite").toLowerCase();

  if (s.includes("zalo")) return "zalo";
  if (s.includes("mess")) return "mess";
  if (s.includes("lite")) return "lite";

  return s || "lite";
}

function normalizeProgress(value) {
  // Không cho đối tượng Date hiển thị thành chuỗi ngày tháng.
  if (value instanceof Date) {
    return "Chưa liên hệ";
  }

  const s = String(value ?? "").trim();

  if (!s) {
    return "Chưa liên hệ";
  }

  // Nếu dữ liệu là chuỗi ngày tháng dài thì không dùng làm tiến trình.
  const looksLikeDate =
    /^\d{4}-\d{1,2}-\d{1,2}(T.*)?$/.test(s) ||
    /^\d{1,2}[/-]\d{1,2}[/-]\d{2,4}$/.test(s) ||
    /^(Mon|Tue|Wed|Thu|Fri|Sat|Sun)\s/i.test(s);

  if (looksLikeDate) {
    return "Chưa liên hệ";
  }

  // Ngăn số serial ngày Excel bị hiển thị như tiến trình.
  if (/^\d{5}(\.\d+)?$/.test(s)) {
    return "Chưa liên hệ";
  }

  const known = [
    "Chưa liên hệ",
    "Đã liên hệ",
    "Đang tư vấn",
    "Đang chốt",
    "Thành công",
    "Không thành công"
  ];

  const found = known.find(
    item => item.toLowerCase() === s.toLowerCase()
  );

  // Giữ các tiến trình hợp lệ khác, ví dụ 14/14.
  return found || s;
}

/* CHUẨN HÓA THÔNG TIN KHÁCH HÀNG */

function normalizeCustomer(c = {}) {
  return {
    id: String(c.id || c.ID || makeId()),

    name: String(
      c.name ??
      c["Khách Hàng"] ??
      c["Khách hàng"] ??
      c["Tên khách hàng"] ??
      ""
    ).trim(),

    closeDate: dateISO(
      c.closeDate ??
      c["Ngày Chốt"] ??
      c["Ngày chốt"] ??
      ""
    ),

    progress: normalizeProgress(
      c.progress ??
      c["Tiến Trình"] ??
      c["Tiến trình"] ??
      ""
    ),

    deadline: dateISO(
      c.deadline ??
      c["Hạn Chốt"] ??
      c["Hạn chốt"] ??
      ""
    ),

    revenue: Math.max(
      0,
      Number(
        String(
          c.revenue ??
          c["Doanh Số"] ??
          c["Doanh số"] ??
          0
        ).replace(/[,.](?=\d{3}(\D|$))/g, "")
      ) || 0
    ),

    type: normalizeType(
      c.type ??
      c["Loại Khách"] ??
      c["Loại khách"] ??
      ""
    ),

    symbol: normalizeSymbol(
      c.symbol ??
      c["Ký Hiệu"] ??
      c["Ký hiệu"] ??
      "lite"
    ),

    explanation: String(
      c.explanation ??
      c["Giải Thích"] ??
      c["Giải thích"] ??
      ""
    ),

    createdAt: c.createdAt || new Date().toISOString(),
    updatedAt: c.updatedAt || ""
  };
}

/* THÔNG BÁO */

function toast(message) {
  $("toast").textContent = message;
  $("toast").classList.add("show");

  clearTimeout(toastTimer);

  toastTimer = setTimeout(() => {
    $("toast").classList.remove("show");
  }, 3200);
}

/* LƯU VÀ ĐỌC DỮ LIỆU */

function save() {
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(customers));
  } catch (error) {
    console.error(error);
    toast("Không thể lưu dữ liệu trên trình duyệt.");
  }
}

function load() {
  try {
    const value = JSON.parse(localStorage.getItem(STORAGE_KEY) || "[]");

    customers = Array.isArray(value)
      ? value.map(normalizeCustomer)
      : [];
  } catch (error) {
    customers = [];
    toast("Không đọc được dữ liệu đã lưu.");
  }
}

/* ĐIỀU HƯỚNG */

function page(id) {
  document.querySelectorAll(".page").forEach(element => {
    element.classList.toggle("hidden", element.id !== id);
  });

  document.querySelectorAll(".nav").forEach(element => {
    element.classList.toggle("active", element.dataset.page === id);
  });
}

/* TÌM KIẾM VÀ LỌC */

function filtered() {
  const q = $("searchInput").value.trim().toLowerCase();
  const type = $("typeFilter").value;
  const symbol = $("symbolFilter").value;

  return customers.filter(c =>
    (
      !q ||
      [
        c.name,
        c.explanation,
        c.progress,
        c.type,
        c.symbol
      ].join(" ").toLowerCase().includes(q)
    ) &&
    (!type || c.type === type) &&
    (!symbol || c.symbol === symbol)
  );
}

/* THỐNG KÊ */

function renderStats() {
  const today = todayISO();

  $("statTotal").textContent =
    customers.length.toLocaleString("vi-VN");

  $("statNew").textContent =
    customers.filter(c => c.type === "Mới")
      .length.toLocaleString("vi-VN");

  $("statRevenue").textContent =
    money(customers.reduce((sum, c) => sum + c.revenue, 0));

  $("statUrgent").textContent =
    customers.filter(c =>
      c.deadline &&
      c.deadline <= today &&
      !["Thành công", "Không thành công"].includes(c.progress)
    ).length.toLocaleString("vi-VN");

  $("year").textContent = new Date().getFullYear();
}

function typeClass(type) {
  return type === "Cũ (Vip)"
    ? "type-vip"
    : type === "Cũ"
      ? "type-old"
      : "type-new";
}

function progressClass(progress) {
  return progress === "Thành công"
    ? "progress-success"
    : progress === "Không thành công"
      ? "progress-failed"
      : "progress-normal";
}

/* HIỂN THỊ BẢNG TẤT CẢ KHÁCH HÀNG
   Ngày chốt và Hạn chốt được hiển thị trực tiếp tại đây. */

function renderTable() {
  const list = filtered();
  const today = todayISO();

  $("resultCount").textContent = `${list.length} khách hàng`;

  $("tableWrap").classList.toggle("hidden", !list.length);
  $("emptyState").classList.toggle("hidden", !!list.length);

  $("customerBody").innerHTML = list.map((c, i) => {
    const urgent =
      c.deadline &&
      c.deadline <= today &&
      !["Thành công", "Không thành công"].includes(c.progress);

    return `
      <tr>
        <td>${i + 1}</td>

        <td class="customer-name">
          ${esc(c.name || "Chưa đặt tên")}
        </td>

        <td>${dateText(c.closeDate)}</td>

        <td>
          <span class="pill ${progressClass(c.progress)}">
            ${esc(c.progress)}
          </span>
        </td>

        <td class="${urgent ? "urgent" : ""}">
          ${dateText(c.deadline)}${urgent ? "" : ""}
        </td>

        <td class="money">${money(c.revenue)}</td>

        <td>
          <span class="pill ${typeClass(c.type)}">
            ${esc(c.type)}
          </span>
        </td>

        <td>${esc(c.symbol)}</td>

        <td>${esc(c.explanation || "—")}</td>

        <td>
          <div class="row-actions">
            <button
              class="action"
              data-action="receipt"
              data-id="${esc(c.id)}"
            >Biên lai</button>

            <button
              class="action"
              data-action="edit"
              data-id="${esc(c.id)}"
            >Sửa</button>

            <button
              class="action delete"
              data-action="delete"
              data-id="${esc(c.id)}"
            >Xóa</button>
          </div>
        </td>
      </tr>
    `;
  }).join("");
}

/* KHÁCH HÀNG GẦN ĐÂY */

function renderRecent() {
  const list = [...customers]
    .sort((a, b) =>
      (b.updatedAt || b.createdAt)
        .localeCompare(a.updatedAt || a.createdAt)
    )
    .slice(0, 5);

  $("recentList").innerHTML = list.length
    ? list.map(c => `
      <div class="recent-item">
        <div>
          <div class="recent-name">${esc(c.name)}</div>
          <div class="recent-meta">
            ${esc(c.type)} · ${esc(c.progress)}
          </div>
        </div>
        <div class="recent-revenue">${money(c.revenue)}</div>
      </div>
    `).join("")
    : `
      <div class="empty">
        <h3>Chưa có khách hàng</h3>
        <p>Thêm khách hàng hoặc tải file Excel lên để bắt đầu.</p>
      </div>
    `;
}

function render() {
  renderStats();
  renderTable();
  renderRecent();
}

/* FORM THÊM VÀ SỬA */

function openModal(c = null) {
  $("customerForm").reset();

  $("customerId").value = c?.id || "";
  $("modalTitle").textContent = c ? "Sửa khách hàng" : "Thêm khách hàng";

  $("name").value = c?.name || "";
  $("closeDate").value = c?.closeDate || todayISO();
  $("deadline").value = c?.deadline || "";
  $("progress").value = c?.progress || "Chưa liên hệ";
  $("revenue").value = c?.revenue ?? 0;
  $("type").value = c?.type || "Mới";
  $("symbol").value = c?.symbol || "lite";
  $("explanation").value = c?.explanation || "";

  $("modalBackdrop").classList.remove("hidden");
  $("name").focus();
}

function closeModal() {
  $("modalBackdrop").classList.add("hidden");
}

function bindForm() {
  $("addBtn").onclick = () => openModal();
  $("addBtn2").onclick = () => openModal();
  $("emptyAddBtn").onclick = () => openModal();

  $("closeModal").onclick = closeModal;
  $("cancelModal").onclick = closeModal;

  $("modalBackdrop").addEventListener("click", event => {
    if (event.target === $("modalBackdrop")) {
      closeModal();
    }
  });

  document.addEventListener("keydown", event => {
    if (event.key === "Escape") closeModal();
  });

  $("customerForm").addEventListener("submit", event => {
    event.preventDefault();

    const name = $("name").value.trim();

    if (!name) {
      toast("Vui lòng nhập tên khách hàng.");
      return;
    }

    const id = $("customerId").value;
    const old = customers.find(c => c.id === id);

    const customer = normalizeCustomer({
      id: id || makeId(),
      name,
      closeDate: $("closeDate").value,
      deadline: $("deadline").value,
      progress: $("progress").value,
      revenue: $("revenue").value,
      type: $("type").value,
      symbol: $("symbol").value,
      explanation: $("explanation").value.trim(),
      createdAt: old?.createdAt || new Date().toISOString(),
      updatedAt: new Date().toISOString()
    });

    if (old) {
      customers = customers.map(c => c.id === id ? customer : c);
      toast("Đã cập nhật khách hàng.");
    } else {
      customers.unshift(customer);
      toast("Đã thêm khách hàng.");
    }

    save();
    render();
    closeModal();
  });
}

/* THAO TÁC BẢNG */

function bindTable() {
  document.querySelectorAll(".nav").forEach(button => {
    button.addEventListener("click", () => page(button.dataset.page));
  });

  $("viewAllBtn").onclick = () => page("customers");

  ["searchInput", "typeFilter", "symbolFilter"].forEach(id => {
    $(id).addEventListener("input", renderTable);
    $(id).addEventListener("change", renderTable);
  });

  $("scrollTopBtn").onclick = () => {
    $("tableWrap").scrollTo({ top: 0, behavior: "smooth" });
  };

  $("scrollBottomBtn").onclick = () => {
    $("tableWrap").scrollTo({
      top: $("tableWrap").scrollHeight,
      behavior: "smooth"
    });
  };

  $("customerBody").addEventListener("click", event => {
    const button = event.target.closest("[data-action]");

    if (!button) return;

    const customer = customers.find(c => c.id === button.dataset.id);

    if (!customer) return;

    if (button.dataset.action === "edit") {
      openModal(customer);
    }

    if (
      button.dataset.action === "delete" &&
      confirm(`Xóa khách hàng "${customer.name}"?`)
    ) {
      customers = customers.filter(c => c.id !== customer.id);
      save();
      render();
      toast("Đã xóa khách hàng.");
    }

    if (button.dataset.action === "receipt") {
      printReceipt(customer);
    }
  });
}

/* XUẤT DỮ LIỆU */

function rowsForExport() {
  return customers.map((c, i) => ({
    "STT": i + 1,
    "Khách Hàng": c.name,
    "Ngày Chốt": c.closeDate,
    "Tiến Trình": c.progress,
    "Hạn Chốt": c.deadline,
    "Doanh Số": c.revenue,
    "Loại Khách": c.type,
    "Ký Hiệu": c.symbol,
    "Giải Thích": c.explanation
  }));
}

function download(blob, name) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");

  a.href = url;
  a.download = name;

  document.body.appendChild(a);
  a.click();
  a.remove();

  setTimeout(() => URL.revokeObjectURL(url), 1200);
}

function exportCSV() {
  const rows = rowsForExport();

  const heads = [
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

  const quote = value => {
    let s = String(value ?? "");

    if (/^[=+@-]/.test(s)) {
      s = "'" + s;
    }

    return '"' + s.replace(/"/g, '""') + '"';
  };

  const csv = [
    heads.map(quote).join(","),
    ...rows.map(row => heads.map(h => quote(row[h])).join(","))
  ].join("\r\n");

  download(
    new Blob(["\ufeff" + csv], {
      type: "text/csv;charset=utf-8"
    }),
    `Cho_Sale_${todayISO()}.csv`
  );

  toast("Đã xuất CSV.");
}

function exportExcel() {
  if (typeof XLSX === "undefined") {
    toast("Không tải được thư viện Excel. Hãy kiểm tra Internet.");
    return;
  }

  const ws = XLSX.utils.json_to_sheet(rowsForExport());

  ws["!cols"] = [
    { wch: 6 },
    { wch: 24 },
    { wch: 15 },
    { wch: 22 },
    { wch: 15 },
    { wch: 18 },
    { wch: 15 },
    { wch: 14 },
    { wch: 40 }
  ];

  const wb = XLSX.utils.book_new();

  XLSX.utils.book_append_sheet(wb, ws, "KhachHang");

  XLSX.writeFile(wb, `Cho_Sale_${todayISO()}.xlsx`);

  toast("Đã xuất Excel.");
}

/* NHẬN DIỆN TIÊU ĐỀ EXCEL */

function normalizeHeader(value) {
  return String(value ?? "")
    .trim()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[đĐ]/g, "d")
    .toLowerCase()
    .replace(/[^a-z0-9]/g, "");
}

function headerValue(row, aliases) {
  const wanted = new Set(aliases.map(normalizeHeader));

  const key = Object.keys(row).find(k =>
    wanted.has(normalizeHeader(k))
  );

  return key === undefined ? "" : row[key];
}

function findHeaderRow(rows) {
  for (let i = 0; i < Math.min(rows.length, 30); i++) {
    const headers = (rows[i] || []).map(normalizeHeader);

    const hasName = headers.some(h =>
      ["khachhang", "tenkhachhang", "ten", "hoten"].includes(h)
    );

    const hasClose = headers.some(h =>
      [
        "ngaychot",
        "ngaychotkh",
        "ngaychotkhachhang",
        "chotngay",
        "closedate"
      ].includes(h)
    );

    const hasDeadline = headers.some(h =>
      [
        "hanchot",
        "hanchotkh",
        "hanchotkhachhang",
        "deadline",
        "ngayhethan"
      ].includes(h)
    );

    if (hasName && (hasClose || hasDeadline)) {
      return i;
    }
  }

  return 0;
}

/* NHẬP EXCEL
   Đọc Ngày Chốt và Hạn Chốt vào từng khách hàng. */

function importExcel(file) {
  if (typeof XLSX === "undefined") {
    toast("Không tải được thư viện Excel. Hãy kiểm tra Internet.");
    return;
  }

  const reader = new FileReader();

  reader.onload = event => {
    try {
      const wb = XLSX.read(event.target.result, {
        type: "array",
        cellDates: false
      });

      const sheet = wb.Sheets[wb.SheetNames[0]];

      if (!sheet) {
        throw Error("Không tìm thấy sheet");
      }

      const rows = XLSX.utils.sheet_to_json(sheet, {
        header: 1,
        defval: "",
        raw: true,
        blankrows: false
      });

      if (!rows.length) {
        throw Error("File không có dữ liệu.");
      }

      const headerIndex = findHeaderRow(rows);
      const headers = rows[headerIndex] || [];

      const source = rows.slice(headerIndex + 1).map(row => {
        const obj = {};

        headers.forEach((header, i) => {
          if (String(header ?? "").trim()) {
            obj[String(header).trim()] = row[i] ?? "";
          }
        });

        return obj;
      });

      const imported = source.map(row => normalizeCustomer({
        id: headerValue(row, [
          "ID",
          "Mã khách hàng",
          "Mã KH"
        ]) || makeId(),

        name: headerValue(row, [
          "Khách Hàng",
          "Khách hàng",
          "Tên khách hàng",
          "Tên KH",
          "Họ tên",
          "Tên"
        ]),

        closeDate: headerValue(row, [
          "Ngày Chốt",
          "Ngày chốt KH",
          "Ngày chốt khách hàng",
          "Chốt ngày",
          "Close date"
        ]),

        progress: headerValue(row, [
          "Tiến Trình",
          "Tiến trình",
          "Trạng thái",
          "Tiến độ"
        ]),

        deadline: headerValue(row, [
          "Hạn Chốt",
          "Hạn chốt KH",
          "Hạn chốt khách hàng",
          "Deadline",
          "Ngày hết hạn"
        ]),

        revenue: headerValue(row, [
          "Doanh Số",
          "Doanh số",
          "Doanh thu",
          "Giá trị"
        ]),

        type: headerValue(row, [
          "Loại Khách",
          "Loại khách",
          "Loại KH",
          "Loại"
        ]),

        symbol: headerValue(row, [
          "Ký Hiệu",
          "Ký hiệu",
          "Nguồn",
          "Kênh"
        ]),

        explanation: headerValue(row, [
          "Giải Thích",
          "Giải thích",
          "Ghi chú",
          "Mô tả"
        ])
      })).filter(c => c.name);

      if (!imported.length) {
        toast(
          "Không thấy khách hàng hợp lệ. Kiểm tra tiêu đề cột Khách Hàng, Ngày Chốt và Hạn Chốt."
        );
        return;
      }

      const dateCount = imported.filter(c =>
        c.closeDate || c.deadline
      ).length;

      const replace = confirm(
        `Đọc được ${imported.length} khách hàng, trong đó ${dateCount} dòng có Ngày Chốt hoặc Hạn Chốt.\n\n` +
        "OK: Thay thế dữ liệu hiện tại.\n" +
        "Cancel: Gộp dữ liệu vào danh sách."
      );

      if (replace) {
        customers = imported;
      } else {
        const map = new Map(
          customers.map(c => [c.name.toLowerCase(), c])
        );

        for (const c of imported) {
          const key = c.name.toLowerCase();
          const old = map.get(key);

          if (old) {
            Object.assign(old, c, {
              id: old.id,
              updatedAt: new Date().toISOString()
            });
          } else {
            customers.push(c);
            map.set(key, c);
          }
        }
      }

      save();
      render();
      page("customers");

      toast(
        `Đã nhập ${imported.length} khách hàng. Ngày chốt/hạn chốt đã được đưa vào bảng Tất cả khách hàng.`
      );
    } catch (error) {
      console.error(error);

      toast(
        "Không đọc được file Excel/CSV: " +
        (error.message || "Kiểm tra định dạng file.")
      );
    } finally {
      $("importFile").value = "";
    }
  };

  reader.onerror = () => {
    toast("Không thể đọc file.");
    $("importFile").value = "";
  };

  reader.readAsArrayBuffer(file);
}

/* SAO LƯU VÀ KHÔI PHỤC JSON */

function backup() {
  download(
    new Blob([
      JSON.stringify({
        app: "Cho Sale Gia Nghia CRM",
        version: "3.5",
        exportedAt: new Date().toISOString(),
        customers
      }, null, 2)
    ], {
      type: "application/json;charset=utf-8"
    }),
    `CRM_Backup_${todayISO()}.json`
  );

  toast("Đã sao lưu JSON.");
}

function restore(file) {
  const reader = new FileReader();

  reader.onload = () => {
    try {
      const data = JSON.parse(reader.result);
      const list = Array.isArray(data) ? data : data.customers;

      if (!Array.isArray(list)) {
        throw Error("JSON không hợp lệ");
      }

      if (!confirm(
        `Khôi phục ${list.length} khách hàng và thay thế dữ liệu hiện tại?`
      )) {
        return;
      }

      customers = list.map(normalizeCustomer);

      save();
      render();

      toast("Đã khôi phục dữ liệu.");
    } catch (error) {
      toast("File JSON không hợp lệ.");
    } finally {
      $("restoreFile").value = "";
    }
  };

  reader.onerror = () => toast("Không đọc được file JSON.");

  reader.readAsText(file, "utf-8");
}

function bindFiles() {
  $("uploadExcelBtn").onclick = () => $("importFile").click();

  $("importFile").addEventListener("change", event => {
    const file = event.target.files?.[0];

    if (!file) return;

    if (!/\.(xlsx|xls|csv)$/i.test(file.name)) {
      toast("Chọn file .xlsx, .xls hoặc .csv.");
      event.target.value = "";
      return;
    }

    importExcel(file);
  });

  $("exportExcelBtn").onclick = exportExcel;
  $("exportCsvBtn").onclick = exportCSV;
  $("backupBtn").onclick = backup;

  $("restoreBtn").onclick = () => $("restoreFile").click();

  $("restoreFile").addEventListener("change", event => {
    if (event.target.files?.[0]) {
      restore(event.target.files[0]);
    }
  });
}

/* ĐỔI SỐ TIỀN THÀNH CHỮ */

const ones = [
  "không", "một", "hai", "ba", "bốn",
  "năm", "sáu", "bảy", "tám", "chín"
];

function readTriple(n, full) {
  const h = Math.floor(n / 100);
  const t = Math.floor((n % 100) / 10);
  const u = n % 10;

  let out = "";

  if (h) {
    out += ones[h] + " trăm";
  } else if (full && (t || u)) {
    out += "không trăm";
  }

  if (t > 1) {
    out += (out ? " " : "") + ones[t] + " mươi";

    if (u === 1) out += " mốt";
    else if (u === 4) out += " tư";
    else if (u === 5) out += " lăm";
    else if (u) out += " " + ones[u];
  } else if (t === 1) {
    out += (out ? " " : "") + "mười";

    if (u === 5) out += " lăm";
    else if (u) out += " " + ones[u];
  } else if (u) {
    if (h || full) {
      out += (out ? " " : "") + "lẻ " + ones[u];
    } else {
      out += ones[u];
    }
  }

  return out.trim();
}

function amountWords(value) {
  let n = Math.floor(Math.max(0, Number(value) || 0));

  if (!n) return "Không đồng";

  const scales = [
    "",
    " nghìn",
    " triệu",
    " tỷ",
    " nghìn tỷ",
    " triệu tỷ"
  ];

  const groups = [];

  while (n > 0) {
    groups.push(n % 1000);
    n = Math.floor(n / 1000);
  }

  const parts = [];

  for (let i = groups.length - 1; i >= 0; i--) {
    if (!groups[i]) continue;

    parts.push(
      readTriple(
        groups[i],
        parts.length > 0 && groups[i] < 100
      ) + (scales[i] || "")
    );
  }

  let result = parts.join(" ").replace(/\s+/g, " ").trim();

  return result.charAt(0).toUpperCase() + result.slice(1) + " đồng";
}

/* TẠO BIÊN LAI VÀ IN / LƯU PDF */

function printReceipt(c) {
  const no =
    "BL-" +
    todayISO().replace(/-/g, "") +
    "-" +
    c.id.replace(/\W/g, "").slice(-6).toUpperCase();

  const amount = Number(c.revenue) || 0;

  const w = window.open("", "_blank", "width=900,height=1000");

  if (!w) {
    toast("Trình duyệt chặn cửa sổ biên lai. Cho phép pop-up rồi thử lại.");
    return;
  }

  const html = `<!doctype html>
<html lang="vi">
<head>
<meta charset="utf-8">
<title>Biên lai ${esc(no)}</title>
<style>
*{box-sizing:border-box}
body{margin:0;padding:22px;background:#eef2f6;color:#243447;font:13px/1.55 Arial,sans-serif}
.tools{display:flex;justify-content:center;gap:10px;margin:0 auto 18px}
.tools button{padding:10px 16px;border:0;border-radius:7px;background:#146c94;color:#fff;font-weight:bold;cursor:pointer}
.tools .secondary{background:#fff;color:#102f43;border:1px solid #dce5ec}
.receipt{max-width:800px;min-height:950px;margin:auto;padding:42px 48px;background:white;box-shadow:0 10px 35px #102f4315}
.stripe{height:5px;margin:-42px -48px 30px;background:linear-gradient(90deg,#102f43,#146c94,#4db5dc)}
header{display:flex;justify-content:space-between;gap:20px;border-bottom:2px solid #102f43;padding-bottom:22px}
.brand{display:flex;gap:12px;align-items:center}
.logo{width:52px;height:52px;display:grid;place-items:center;background:#102f43;color:#fff;border-radius:12px;font-size:19px;font-weight:bold}
.brand h2{margin:0;color:#102f43;font-size:20px}
.brand p{margin:4px 0 0;color:#64748b;font-size:10px;letter-spacing:1px}
.meta{text-align:right;font-size:11px}
.meta b{display:block;color:#146c94;font-size:13px;margin-bottom:6px}
.title{text-align:center;margin:30px 0 4px;font-size:26px;letter-spacing:1px;color:#102f43}
.subtitle{text-align:center;color:#64748b;font-size:11px;margin-bottom:25px}
.sect{margin:22px 0 10px;font-weight:bold;font-size:12px;color:#102f43;letter-spacing:.5px}
.customer{display:grid;grid-template-columns:1fr 1fr;gap:13px 22px;padding:17px;border:1px solid #dce5ec;border-radius:9px;background:#fbfdfe}
.label{display:block;color:#64748b;font-size:10px;text-transform:uppercase}
.value{font-weight:bold;overflow-wrap:anywhere}
table{width:100%;border-collapse:collapse;margin-top:10px}
th{background:#102f43;color:white;text-align:left;padding:11px;font-size:11px}
td{padding:14px 10px;border-bottom:1px solid #dce5ec;vertical-align:top}
td:last-child,th:last-child{text-align:right}
.amount{margin-top:18px;padding:18px 20px;border:1px solid #d3e8f2;border-radius:9px;background:#eaf4f9}
.amount .label{font-weight:bold;letter-spacing:.5px}
.big{font-size:26px;font-weight:900;color:#102f43;margin:6px 0 12px;overflow-wrap:anywhere}
.words{padding-top:10px;border-top:1px dashed #b8d4e2;font-style:italic}
.note{margin-top:20px;padding:12px 14px;border-left:3px solid #146c94;background:#f7fafc;color:#536779;font-size:10px}
.signs{display:grid;grid-template-columns:1fr 1fr;gap:25px;text-align:center;margin-top:40px}
.signs b{display:block;color:#102f43}
.signs small{color:#64748b;font-style:italic}
.space{height:90px}
.foot{text-align:center;border-top:1px solid #dce5ec;padding-top:12px;margin-top:22px;color:#64748b;font-size:10px}
@page{size:A4;margin:12mm}
@media print{
 body{padding:0;background:white;-webkit-print-color-adjust:exact;print-color-adjust:exact}
 .tools{display:none}
 .receipt{max-width:none;min-height:0;padding:12mm;box-shadow:none}
 .stripe{margin:-12mm -12mm 9mm}
 .customer,.amount,.table,.signs{break-inside:avoid}
}
@media(max-width:600px){
 body{padding:8px}
 .receipt{padding:20px;min-height:0}
 .stripe{margin:-20px -20px 20px}
 header{flex-direction:column}
 .meta{text-align:left}
 .customer{grid-template-columns:1fr}
 .title{font-size:22px}
}
</style>
</head>
<body>
<div class="tools">
  <button onclick="window.print()">In / Lưu PDF</button>
  <button class="secondary" onclick="window.close()">Đóng</button>
</div>

<main class="receipt">
  <div class="stripe"></div>

  <header>
    <div class="brand">
      <div class="logo">CS</div>
      <div>
        <h2>CHỢ SALE GIA NGHĨA</h2>
        <p>HỆ THỐNG QUẢN LÝ KHÁCH HÀNG</p>
      </div>
    </div>

    <div class="meta">
      <b>BIÊN LAI NỘI BỘ</b>
      <div>Mã chứng từ: ${esc(no)}</div>
      <div>Ngày lập: ${dateText(todayISO())}</div>
    </div>
  </header>

  <h1 class="title">BIÊN LAI THANH TOÁN</h1>
  <div class="subtitle">Chứng từ ghi nhận thông tin giao dịch nội bộ</div>

  <div class="sect">01 / THÔNG TIN KHÁCH HÀNG</div>

  <div class="customer">
    <div>
      <span class="label">Khách hàng</span>
      <span class="value">${esc(c.name)}</span>
    </div>

    <div>
      <span class="label">Mã khách hàng</span>
      <span class="value">${esc(c.id)}</span>
    </div>

    <div>
      <span class="label">Ngày chốt</span>
      <span class="value">${dateText(c.closeDate)}</span>
    </div>

    <div>
      <span class="label">Tiến trình</span>
      <span class="value">${esc(c.progress)}</span>
    </div>

    <div>
      <span class="label">Loại khách</span>
      <span class="value">${esc(c.type)}</span>
    </div>

    <div>
      <span class="label">Kênh / Ký hiệu</span>
      <span class="value">${esc(c.symbol)}</span>
    </div>
  </div>

  <div class="sect">02 / CHI TIẾT GIAO DỊCH</div>

  <table>
    <thead>
      <tr>
        <th>STT</th>
        <th>Nội dung</th>
        <th>Số lượng</th>
        <th>Thành tiền</th>
      </tr>
    </thead>

    <tbody>
      <tr>
        <td>01</td>
        <td>
          <b>Giá trị giao dịch / doanh số</b><br>
          <span style="color:#64748b">
            ${esc(c.explanation || "Theo thông tin ghi nhận trong CRM")}
          </span>
        </td>
        <td>1</td>
        <td><b>${money(amount)}</b></td>
      </tr>
    </tbody>
  </table>

  <div class="amount">
    <span class="label">TỔNG GIÁ TRỊ CHỨNG TỪ</span>
    <div class="big">${money(amount)}</div>
    <div class="words">
      <b>Bằng chữ:</b> ${esc(amountWords(amount))}.
    </div>
  </div>

  <div class="note">
    <b>Lưu ý:</b> Biên lai này được tạo từ dữ liệu quản lý nội bộ.
    Cần kiểm tra và xác nhận giao dịch trước khi sử dụng.
    Không thay thế hóa đơn điện tử hoặc chứng từ thuế theo quy định.
  </div>

  <div class="signs">
    <div>
      <b>NGƯỜI NỘP TIỀN</b>
      <small>(Ký và ghi rõ họ tên)</small>
      <div class="space"></div>
    </div>

    <div>
      <b>NGƯỜI LẬP / THU TIỀN</b>
      <small>(Ký và ghi rõ họ tên)</small>
      <div class="space"></div>
    </div>
  </div>

  <div class="foot">
    CHỢ SALE GIA NGHĨA · BIÊN LAI NỘI BỘ · ${esc(no)}
  </div>
</main>
</body>
</html>`;

  w.document.open();
  w.document.write(html);
  w.document.close();

  toast("Đã tạo biên lai. Chọn In / Lưu PDF.");
}

/* KHỞI ĐỘNG */

function init() {
  load();
  bindForm();
  bindTable();
  bindFiles();
  bindExcelDateReader();
  render();
}

document.addEventListener("DOMContentLoaded", init);

/* ĐỌC RIÊNG CỘT NGÀY CHỐT VÀ HẠN CHỐT TỪ EXCEL/CSV */

function normalizeExcelHeader(value) {
  return String(value ?? "")
    .trim()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[đĐ]/g, "d")
    .toLowerCase()
    .replace(/[^a-z0-9]/g, "");
}

function findExcelColumn(headers, aliases) {
  const names = aliases.map(normalizeExcelHeader);

  return headers.findIndex(value =>
    names.includes(normalizeExcelHeader(value))
  );
}

function excelDateDisplay(value) {
  if (value === 0 || String(value).trim() === "0") {
    return "0";
  }

  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return `${String(value.getDate()).padStart(2, "0")}/${String(value.getMonth() + 1).padStart(2, "0")}/${value.getFullYear()}`;
  }

  if (
    typeof value === "number" &&
    Number.isFinite(value) &&
    typeof XLSX !== "undefined"
  ) {
    const parts = XLSX.SSF.parse_date_code(value);

    if (parts && parts.y && parts.m && parts.d) {
      return `${String(parts.d).padStart(2, "0")}/${String(parts.m).padStart(2, "0")}/${parts.y}`;
    }

    return String(value);
  }

  const text = String(value).trim();

  if (!text) return "—";

  let match = text.match(/^(\d{1,2})[./-](\d{1,2})[./-](\d{4})$/);

  if (match) {
    return `${match[1].padStart(2, "0")}/${match[2].padStart(2, "0")}/${match[3]}`;
  }

  match = text.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);

  if (match) {
    return `${match[3].padStart(2, "0")}/${match[2].padStart(2, "0")}/${match[1]}`;
  }

  // Giữ nguyên dạng ngày như 31/5 nếu file không ghi năm.
  return text;
}

function setExcelDateMessage(text, type = "") {
  const el = $("dateExcelMessage");

  if (!el) return;

  el.textContent = text;

  el.classList.toggle("excel-date-message-error", type === "error");
  el.classList.toggle("excel-date-message-success", type === "success");
}

function renderExcelDateRows(rows) {
  const body = $("dateExcelTableBody");

  body.replaceChildren();

  if (!rows.length) {
    const tr = document.createElement("tr");
    const td = document.createElement("td");

    td.colSpan = 4;
    td.className = "muted";
    td.textContent = "Không có dòng dữ liệu phù hợp.";

    tr.appendChild(td);
    body.appendChild(tr);

    return;
  }

  rows.forEach((item, index) => {
    const tr = document.createElement("tr");

    [
      String(index + 1),
      item.customer || "—",
      excelDateDisplay(item.closeDate),
      excelDateDisplay(item.deadline)
    ].forEach(value => {
      const td = document.createElement("td");
      td.textContent = value;
      tr.appendChild(td);
    });

    body.appendChild(tr);
  });
}

function readDateColumnsFromExcel(file) {
  if (typeof XLSX === "undefined") {
    setExcelDateMessage(
      "Không tải được thư viện Excel. Kiểm tra kết nối mạng rồi tải lại trang.",
      "error"
    );
    return;
  }

  const reader = new FileReader();

  reader.onload = event => {
    try {
      const workbook = XLSX.read(event.target.result, {
        type: "array",
        cellDates: false
      });

      if (!workbook.SheetNames.length) {
        throw new Error("File không có trang tính.");
      }

      const sheet = workbook.Sheets[workbook.SheetNames[0]];

      const rows = XLSX.utils.sheet_to_json(sheet, {
        header: 1,
        defval: "",
        raw: false,
        blankrows: false
      });

      if (!rows.length) {
        throw new Error("File không có dữ liệu.");
      }

      let headerRow = -1;
      let closeIndex = -1;
      let deadlineIndex = -1;
      let nameIndex = -1;

      for (let i = 0; i < Math.min(rows.length, 30); i++) {
        const candidate = rows[i] || [];

        const close = findExcelColumn(candidate, [
          "Ngày Chốt",
          "Ngay chot",
          "Ngày chốt khách hàng",
          "Ngày"
        ]);

        const deadline = findExcelColumn(candidate, [
          "Hạn Chốt",
          "Han chot",
          "Hạn chốt khách hàng",
          "Deadline",
          "Hạn"
        ]);

        if (close !== -1 && deadline !== -1) {
          headerRow = i;
          closeIndex = close;
          deadlineIndex = deadline;

          nameIndex = findExcelColumn(candidate, [
            "Khách Hàng",
            "Khách hàng",
            "Tên khách hàng",
            "Tên"
          ]);

          break;
        }
      }

      if (headerRow === -1) {
        throw new Error(
          "Không tìm thấy đủ hai cột “Ngày Chốt” và “Hạn Chốt” trong 30 dòng đầu."
        );
      }

      const result = [];

      rows.slice(headerRow + 1).forEach(row => {
        const customer = nameIndex !== -1
          ? String(row[nameIndex] ?? "").trim()
          : "";

        const closeDate = row[closeIndex] ?? "";
        const deadline = row[deadlineIndex] ?? "";

        if (
          !customer &&
          !String(closeDate).trim() &&
          !String(deadline).trim()
        ) {
          return;
        }

        result.push({
          customer,
          closeDate,
          deadline
        });
      });

      renderExcelDateRows(result);

      setExcelDateMessage(
        `Đã đọc ${result.length} dòng từ "${file.name}" (sheet: ${workbook.SheetNames[0]}).`,
        "success"
      );
    } catch (error) {
      console.error("Lỗi đọc cột ngày từ Excel:", error);

      renderExcelDateRows([]);

      setExcelDateMessage(
        error.message || "Không thể đọc file Excel/CSV.",
        "error"
      );
    }
  };

  reader.onerror = () => {
    setExcelDateMessage("Không thể đọc file được chọn.", "error");
  };

  reader.readAsArrayBuffer(file);
}

function bindExcelDateReader() {
  const input = $("dateExcelFile");
  const button = $("readDateExcelBtn");
  const clear = $("clearDateExcelBtn");

  if (!input || !button || !clear) return;

  button.addEventListener("click", () => {
    const file = input.files && input.files[0];

    if (!file) {
      setExcelDateMessage(
        "Vui lòng chọn file Excel hoặc CSV trước.",
        "error"
      );
      return;
    }

    if (!/\.(xlsx|xls|csv)$/i.test(file.name)) {
      setExcelDateMessage(
        "Chỉ hỗ trợ file .xlsx, .xls hoặc .csv.",
        "error"
      );
      return;
    }

    readDateColumnsFromExcel(file);
  });

  clear.addEventListener("click", () => {
    input.value = "";

    renderExcelDateRows([]);

    setExcelDateMessage(
      "Đã xóa bảng tra cứu. Dữ liệu CRM không bị thay đổi."
    );
  });
}