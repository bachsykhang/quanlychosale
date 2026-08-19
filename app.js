const STORAGE_KEY = "crm_customers_v1";
const STATUSES = ["Moi", "Dang cham soc", "Da mua", "Tam dung"];

const state = {
  customers: loadCustomers(),
  query: "",
  status: "",
  source: ""
};

const el = {
  rows: document.querySelector("#customerRows"),
  empty: document.querySelector("#emptyState"),
  total: document.querySelector("#totalCustomers"),
  active: document.querySelector("#activeCustomers"),
  revenue: document.querySelector("#expectedRevenue"),
  follow: document.querySelector("#needFollowUp"),
  search: document.querySelector("#searchInput"),
  statusFilter: document.querySelector("#statusFilter"),
  sourceFilter: document.querySelector("#sourceFilter"),
  pipeline: document.querySelector("#pipelineGrid"),
  lastUpdated: document.querySelector("#lastUpdated"),
  dialog: document.querySelector("#customerDialog"),
  form: document.querySelector("#customerForm"),
  dialogTitle: document.querySelector("#dialogTitle"),
  importStatus: document.querySelector("#importStatus")
};

document.querySelector("#openFormBtn").addEventListener("click", () => openCustomerForm());
document.querySelector("#closeDialogBtn").addEventListener("click", closeDialog);
document.querySelector("#cancelBtn").addEventListener("click", closeDialog);
document.querySelector("#exportBtn").addEventListener("click", exportCsv);
document.querySelector("#sampleBtn").addEventListener("click", loadSampleData);
document.querySelector("#loadSheetBtn").addEventListener("click", importFromSheet);
el.search.addEventListener("input", event => {
  state.query = event.target.value.trim().toLowerCase();
  render();
});
el.statusFilter.addEventListener("change", event => {
  state.status = event.target.value;
  render();
});
el.sourceFilter.addEventListener("change", event => {
  state.source = event.target.value;
  render();
});
el.form.addEventListener("submit", saveCustomer);

render();

function loadCustomers() {
  try {
    return JSON.parse(localStorage.getItem(STORAGE_KEY)) || [];
  } catch {
    return [];
  }
}

function persist() {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(state.customers));
  el.lastUpdated.textContent = `Cap nhat luc ${new Date().toLocaleString("vi-VN")}`;
}

function render() {
  const customers = getFilteredCustomers();
  renderRows(customers);
  renderMetrics();
  renderSources();
  renderPipeline();
}

function getFilteredCustomers() {
  return state.customers.filter(customer => {
    const haystack = [
      customer.name,
      customer.phone,
      customer.email,
      customer.source,
      customer.owner,
      customer.note
    ].join(" ").toLowerCase();
    const matchesQuery = !state.query || haystack.includes(state.query);
    const matchesStatus = !state.status || customer.status === state.status;
    const matchesSource = !state.source || customer.source === state.source;
    return matchesQuery && matchesStatus && matchesSource;
  });
}

function renderRows(customers) {
  el.rows.innerHTML = customers.map(customer => `
    <tr>
      <td>
        <strong>${escapeHtml(customer.name || "Chua co ten")}</strong>
        <span>${escapeHtml(customer.note || "Khong co ghi chu")}</span>
      </td>
      <td>
        <strong>${escapeHtml(customer.phone || "-")}</strong>
        <span>${escapeHtml(customer.email || "-")}</span>
      </td>
      <td>${escapeHtml(customer.source || "-")}</td>
      <td><span class="status ${statusClass(customer.status)}">${escapeHtml(customer.status || "Moi")}</span></td>
      <td>${formatMoney(customer.value)}</td>
      <td>${customer.followUp || "-"}</td>
      <td>
        <div class="row-actions">
          <button type="button" onclick="openCustomerForm('${customer.id}')">Sua</button>
          <button type="button" onclick="deleteCustomer('${customer.id}')">Xoa</button>
        </div>
      </td>
    </tr>
  `).join("");
  el.empty.style.display = customers.length ? "none" : "block";
}

function renderMetrics() {
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const needFollow = state.customers.filter(customer => {
    if (!customer.followUp || customer.status === "Da mua") return false;
    return new Date(customer.followUp) <= today;
  }).length;
  const revenue = state.customers.reduce((sum, customer) => sum + Number(customer.value || 0), 0);

  el.total.textContent = state.customers.length;
  el.active.textContent = state.customers.filter(customer => customer.status === "Dang cham soc").length;
  el.revenue.textContent = formatMoney(revenue);
  el.follow.textContent = needFollow;
}

function renderSources() {
  const current = el.sourceFilter.value;
  const sources = [...new Set(state.customers.map(customer => customer.source).filter(Boolean))].sort();
  el.sourceFilter.innerHTML = '<option value="">Tat ca nguon</option>' + sources
    .map(source => `<option value="${escapeHtml(source)}">${escapeHtml(source)}</option>`)
    .join("");
  el.sourceFilter.value = sources.includes(current) ? current : "";
  state.source = el.sourceFilter.value;
}

function renderPipeline() {
  const total = Math.max(state.customers.length, 1);
  el.pipeline.innerHTML = STATUSES.map(status => {
    const count = state.customers.filter(customer => customer.status === status).length;
    const percent = Math.round((count / total) * 100);
    return `
      <div class="pipe-item">
        <strong>${status}</strong>
        <span>${count} khach - ${percent}%</span>
        <div class="bar"><span style="width:${percent}%"></span></div>
      </div>
    `;
  }).join("");
}

function openCustomerForm(id = "") {
  const customer = state.customers.find(item => item.id === id);
  el.dialogTitle.textContent = customer ? "Sua khach hang" : "Them khach hang";
  document.querySelector("#customerId").value = customer?.id || "";
  document.querySelector("#nameInput").value = customer?.name || "";
  document.querySelector("#phoneInput").value = customer?.phone || "";
  document.querySelector("#emailInput").value = customer?.email || "";
  document.querySelector("#sourceInput").value = customer?.source || "";
  document.querySelector("#statusInput").value = customer?.status || "Moi";
  document.querySelector("#valueInput").value = customer?.value || "";
  document.querySelector("#followUpInput").value = customer?.followUp || "";
  document.querySelector("#ownerInput").value = customer?.owner || "";
  document.querySelector("#noteInput").value = customer?.note || "";
  el.dialog.showModal();
}

function closeDialog() {
  el.dialog.close();
}

function saveCustomer(event) {
  event.preventDefault();
  const id = document.querySelector("#customerId").value || crypto.randomUUID();
  const nextCustomer = {
    id,
    name: document.querySelector("#nameInput").value.trim(),
    phone: document.querySelector("#phoneInput").value.trim(),
    email: document.querySelector("#emailInput").value.trim(),
    source: document.querySelector("#sourceInput").value.trim(),
    status: document.querySelector("#statusInput").value,
    value: Number(document.querySelector("#valueInput").value || 0),
    followUp: document.querySelector("#followUpInput").value,
    owner: document.querySelector("#ownerInput").value.trim(),
    note: document.querySelector("#noteInput").value.trim()
  };

  const existingIndex = state.customers.findIndex(customer => customer.id === id);
  if (existingIndex >= 0) {
    state.customers[existingIndex] = nextCustomer;
  } else {
    state.customers.unshift(nextCustomer);
  }

  persist();
  render();
  closeDialog();
}

function deleteCustomer(id) {
  const customer = state.customers.find(item => item.id === id);
  if (!customer || !confirm(`Xoa khach hang "${customer.name}"?`)) return;
  state.customers = state.customers.filter(item => item.id !== id);
  persist();
  render();
}

async function importFromSheet() {
  const url = document.querySelector("#sheetUrl").value.trim();
  const sheetId = getSheetId(url);
  if (!sheetId) {
    setImportStatus("Khong nhan dien duoc ID Google Sheet.", true);
    return;
  }

  setImportStatus("Dang doc du lieu tu Google Sheet...");
  try {
    const json = await loadSheetJsonp(sheetId);
    const customers = sheetRowsToCustomers(json.table);
    if (!customers.length) throw new Error("Sheet khong co dong du lieu hop le.");
    mergeCustomers(customers);
    setImportStatus(`Da nhap ${customers.length} khach hang tu Google Sheet.`);
  } catch (error) {
    setImportStatus(`Chua doc duoc Sheet: ${error.message}. Hay chia se Sheet o che do Anyone with the link can view.`, true);
  }
}

function loadSheetJsonp(sheetId) {
  return new Promise((resolve, reject) => {
    const callbackName = `sheetCallback_${Date.now()}`;
    const script = document.createElement("script");
    const timeout = setTimeout(() => {
      cleanup();
      reject(new Error("Qua thoi gian cho phan hoi tu Google Sheet"));
    }, 12000);

    window[callbackName] = data => {
      cleanup();
      if (data?.status === "error") {
        reject(new Error(data.errors?.[0]?.detailed_message || "Google Sheet tra ve loi"));
      } else {
        resolve(data);
      }
    };

    script.onerror = () => {
      cleanup();
      reject(new Error("Khong tai duoc Google Sheet"));
    };

    script.src = `https://docs.google.com/spreadsheets/d/${sheetId}/gviz/tq?tqx=out:json;responseHandler:${callbackName}`;
    document.body.appendChild(script);

    function cleanup() {
      clearTimeout(timeout);
      delete window[callbackName];
      script.remove();
    }
  });
}

function sheetRowsToCustomers(table) {
  const headers = table.cols.map(col => normalizeKey(col.label || col.id));
  return table.rows.map(row => {
    const record = {};
    row.c.forEach((cell, index) => {
      record[headers[index] || `cot_${index}`] = cell?.f || cell?.v || "";
    });
    return normalizeCustomer(record);
  }).filter(customer => customer.name || customer.phone || customer.email);
}

function normalizeCustomer(record) {
  const find = names => {
    const key = names.find(name => record[name] !== undefined && record[name] !== "");
    return key ? String(record[key]).trim() : "";
  };

  return {
    id: crypto.randomUUID(),
    name: find(["ten_khach_hang", "khach_hang", "ho_ten", "ten", "name", "customer"]),
    phone: find(["so_dien_thoai", "dien_thoai", "sdt", "phone", "mobile"]),
    email: find(["email", "mail"]),
    source: find(["nguon", "source", "kenh", "channel"]) || "Google Sheet",
    status: normalizeStatus(find(["trang_thai", "status", "tinh_trang"])) || "Moi",
    value: parseNumber(find(["gia_tri", "doanh_thu", "value", "revenue", "amount"])),
    followUp: parseDate(find(["ngay_hen", "lich_hen", "follow_up", "next_follow_up"])),
    owner: find(["nhan_vien", "phu_trach", "owner", "sale"]),
    note: find(["ghi_chu", "note", "notes", "mo_ta"])
  };
}

function mergeCustomers(customers) {
  const known = new Set(state.customers.map(customer => `${customer.phone}|${customer.email}`));
  const fresh = customers.filter(customer => {
    const key = `${customer.phone}|${customer.email}`;
    if (known.has(key) && key !== "|") return false;
    known.add(key);
    return true;
  });
  state.customers = [...fresh, ...state.customers];
  persist();
  render();
}

function loadSampleData() {
  mergeCustomers([
    {
      id: crypto.randomUUID(),
      name: "Nguyen Minh Anh",
      phone: "0901234567",
      email: "minhanh@example.com",
      source: "Facebook",
      status: "Dang cham soc",
      value: 12000000,
      followUp: new Date().toISOString().slice(0, 10),
      owner: "Sale A",
      note: "Quan tam goi dich vu theo thang"
    },
    {
      id: crypto.randomUUID(),
      name: "Tran Quoc Bao",
      phone: "0987654321",
      email: "bao@example.com",
      source: "Website",
      status: "Moi",
      value: 5500000,
      followUp: "",
      owner: "Sale B",
      note: "Can goi tu van lan dau"
    },
    {
      id: crypto.randomUUID(),
      name: "Le Hoang Linh",
      phone: "0911222333",
      email: "linh@example.com",
      source: "Zalo",
      status: "Da mua",
      value: 24000000,
      followUp: "",
      owner: "Sale A",
      note: "Khach da chot don"
    }
  ]);
  setImportStatus("Da nap du lieu mau.");
}

function exportCsv() {
  const headers = ["Ten khach hang", "So dien thoai", "Email", "Nguon", "Trang thai", "Gia tri", "Ngay hen", "Nhan vien", "Ghi chu"];
  const rows = state.customers.map(customer => [
    customer.name,
    customer.phone,
    customer.email,
    customer.source,
    customer.status,
    customer.value,
    customer.followUp,
    customer.owner,
    customer.note
  ]);
  const csv = [headers, ...rows].map(row => row.map(csvCell).join(",")).join("\n");
  const blob = new Blob([csv], { type: "text/csv;charset=utf-8" });
  const link = document.createElement("a");
  link.href = URL.createObjectURL(blob);
  link.download = `khach-hang-${new Date().toISOString().slice(0, 10)}.csv`;
  link.click();
  URL.revokeObjectURL(link.href);
}

function getSheetId(url) {
  return url.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/)?.[1] || "";
}

function normalizeKey(value) {
  return String(value || "")
    .trim()
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[^a-z0-9]+/g, "_")
    .replace(/^_|_$/g, "");
}

function normalizeStatus(value) {
  const raw = normalizeKey(value);
  if (!raw) return "";
  if (["da_mua", "chot", "won", "closed"].includes(raw)) return "Da mua";
  if (["dang_cham_soc", "dang_tu_van", "active", "follow"].includes(raw)) return "Dang cham soc";
  if (["tam_dung", "ngung", "lost", "pause"].includes(raw)) return "Tam dung";
  return "Moi";
}

function parseNumber(value) {
  const number = String(value || "").replace(/[^\d.-]/g, "");
  return Number(number || 0);
}

function parseDate(value) {
  if (!value) return "";
  const date = new Date(value);
  if (!Number.isNaN(date.getTime())) return date.toISOString().slice(0, 10);
  return "";
}

function setImportStatus(message, isError = false) {
  el.importStatus.textContent = message;
  el.importStatus.style.color = isError ? "#ffd0d0" : "#aebfca";
}

function statusClass(status) {
  if (status === "Dang cham soc") return "active";
  if (status === "Da mua") return "won";
  if (status === "Tam dung") return "pause";
  return "new";
}

function formatMoney(value) {
  return new Intl.NumberFormat("vi-VN", {
    style: "currency",
    currency: "VND",
    maximumFractionDigits: 0
  }).format(Number(value || 0));
}

function csvCell(value) {
  return `"${String(value ?? "").replace(/"/g, '""')}"`;
}

function escapeHtml(value) {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}
