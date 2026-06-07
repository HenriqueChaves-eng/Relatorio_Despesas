const CATEGORIES = [
  "Hospedagem",
  "Refeições",
  "Pedágio",
  "Frete",
  "Combustível",
  "Material de Uso e Consumo",
  "Locações de Veículos",
  "Transporte/ taxi",
  "Outras Despesas"
];

const DB_NAME = "relatorio-despesas-mobile";
const DB_VERSION = 1;
const STORE_NAME = "expenses";
const MAX_DATES = 8;
const MAX_OTHER_EXPENSES = 4;

function displayCategory(category) {
  if (category === "Transporte/ taxi") return "Transporte / Táxi";
  return category;
}

const elements = {
  form: document.querySelector("#expense-form"),
  editingId: document.querySelector("#editing-id"),
  date: document.querySelector("#expense-date"),
  amount: document.querySelector("#expense-amount"),
  category: document.querySelector("#expense-category"),
  description: document.querySelector("#expense-description"),
  photo: document.querySelector("#expense-photo"),
  photoStatus: document.querySelector("#photo-status"),
  saveExpense: document.querySelector("#save-expense"),
  cancelEdit: document.querySelector("#cancel-edit"),
  formMessage: document.querySelector("#form-message"),
  grandTotal: document.querySelector("#grand-total"),
  headerCount: document.querySelector("#header-count"),
  dateCount: document.querySelector("#date-count"),
  expenseCount: document.querySelector("#expense-count"),
  summaryList: document.querySelector("#summary-list"),
  entriesList: document.querySelector("#entries-list"),
  limitWarning: document.querySelector("#limit-warning"),
  exportPackage: document.querySelector("#export-package"),
  clearAll: document.querySelector("#clear-all"),
  exportMessage: document.querySelector("#export-message"),
  photoPreviewWrap: document.querySelector("#photo-preview-wrap"),
  photoPreview: document.querySelector("#photo-preview"),
  clearConfirmDialog: document.querySelector("#clear-confirm-dialog"),
  clearConfirmSummary: document.querySelector("#clear-confirm-summary"),
  clearConfirmInput: document.querySelector("#clear-confirm-input"),
  cancelClear: document.querySelector("#cancel-clear"),
  confirmClear: document.querySelector("#confirm-clear"),
  entryTemplate: document.querySelector("#entry-template")
};

let expenses = [];
let retainedPhoto = null;
let retainedPhotoName = "";
let retainedPhotoType = "";
let previewUrl = "";
const objectUrls = [];

function openDatabase() {
  return new Promise((resolve, reject) => {
    const request = indexedDB.open(DB_NAME, DB_VERSION);
    request.onupgradeneeded = () => {
      const db = request.result;
      if (!db.objectStoreNames.contains(STORE_NAME)) {
        db.createObjectStore(STORE_NAME, { keyPath: "id" });
      }
    };
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error);
  });
}

async function withStore(mode, operation) {
  const db = await openDatabase();
  return new Promise((resolve, reject) => {
    const transaction = db.transaction(STORE_NAME, mode);
    const store = transaction.objectStore(STORE_NAME);
    const request = operation(store);
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error);
    transaction.oncomplete = () => db.close();
  });
}

function loadExpenses() {
  return withStore("readonly", store => store.getAll());
}

function saveExpenseRecord(record) {
  return withStore("readwrite", store => store.put(record));
}

function deleteExpenseRecord(id) {
  return withStore("readwrite", store => store.delete(id));
}

function clearExpenseRecords() {
  return withStore("readwrite", store => store.clear());
}

function todayIso() {
  const now = new Date();
  const offset = now.getTimezoneOffset();
  return new Date(now.getTime() - offset * 60000).toISOString().slice(0, 10);
}

function createRecordId() {
  if (globalThis.crypto?.randomUUID) return globalThis.crypto.randomUUID();
  const randomPart = Math.random().toString(36).slice(2, 12);
  return `expense-${Date.now()}-${randomPart}`;
}

function parseAmount(value) {
  const text = String(value || "").trim().replace(/[R$\s]/g, "");
  if (!text) return 0;
  const normalized = text.includes(",")
    ? text.replace(/\./g, "").replace(",", ".")
    : text;
  const amount = Number(normalized);
  return Number.isFinite(amount) ? Math.round(amount * 100) / 100 : 0;
}

function formatMoney(value) {
  return new Intl.NumberFormat("pt-BR", {
    style: "currency",
    currency: "BRL"
  }).format(Number(value || 0));
}

function formatDate(value) {
  if (!value) return "Sem data";
  const [year, month, day] = value.split("-");
  return `${day}/${month}/${year}`;
}

function categoryIndex(category) {
  const index = CATEGORIES.indexOf(category);
  return index >= 0 ? index : CATEGORIES.length;
}

function sortedExpenses(records = expenses) {
  return [...records].sort((a, b) =>
    String(a.date).localeCompare(String(b.date)) ||
    categoryIndex(a.category) - categoryIndex(b.category) ||
    String(a.createdAt).localeCompare(String(b.createdAt))
  );
}

function slug(value) {
  return String(value || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toUpperCase()
    .replace(/[^A-Z0-9]+/g, "_")
    .replace(/^_+|_+$/g, "")
    .slice(0, 42) || "NOTA";
}

function fileExtension(name, type) {
  const match = String(name || "").match(/(\.[a-zA-Z0-9]{2,5})$/);
  if (match) return match[1].toLowerCase();
  if (type === "image/png") return ".png";
  if (type === "image/webp") return ".webp";
  if (type === "image/heic" || type === "image/heif") return ".heic";
  return ".jpg";
}

function formulaForAmounts(amounts) {
  return "=" + amounts.map(value => Number(value).toFixed(2)).join("+");
}

function groupedSummary(records = expenses) {
  const groups = new Map();
  sortedExpenses(records).forEach(expense => {
    const key = `${expense.date}|${expense.category}`;
    if (!groups.has(key)) {
      groups.set(key, {
        date: expense.date,
        category: expense.category,
        amounts: [],
        total: 0
      });
    }
    const group = groups.get(key);
    group.amounts.push(Number(expense.amount));
    group.total = Math.round((group.total + Number(expense.amount)) * 100) / 100;
  });
  return [...groups.values()];
}

function releaseObjectUrls() {
  while (objectUrls.length) URL.revokeObjectURL(objectUrls.pop());
}

function hidePhotoPreview() {
  if (previewUrl) URL.revokeObjectURL(previewUrl);
  previewUrl = "";
  elements.photoPreview.removeAttribute("src");
  elements.photoPreviewWrap.classList.add("hidden");
}

function showPhotoPreview(blob) {
  hidePhotoPreview();
  if (!blob) return;
  previewUrl = URL.createObjectURL(blob);
  elements.photoPreview.src = previewUrl;
  elements.photoPreviewWrap.classList.remove("hidden");
}

function closeClearConfirmation() {
  elements.clearConfirmDialog.close();
  elements.clearConfirmInput.value = "";
  elements.confirmClear.disabled = true;
}

function openClearConfirmation() {
  if (!expenses.length) {
    elements.exportMessage.textContent = "Não existem lançamentos para apagar.";
    return;
  }
  const total = expenses.reduce((sum, item) => sum + Number(item.amount), 0);
  elements.clearConfirmSummary.textContent =
    `${expenses.length} ${expenses.length === 1 ? "lançamento será apagado" : "lançamentos serão apagados"}, totalizando ${formatMoney(total)}.`;
  elements.clearConfirmInput.value = "";
  elements.confirmClear.disabled = true;
  elements.clearConfirmDialog.showModal();
  elements.clearConfirmInput.focus();
}

function render() {
  releaseObjectUrls();
  const ordered = sortedExpenses();
  const grandTotal = ordered.reduce((total, item) => total + Number(item.amount), 0);
  elements.grandTotal.textContent = formatMoney(grandTotal);
  elements.headerCount.textContent = String(ordered.length);
  elements.expenseCount.textContent = `${ordered.length} ${ordered.length === 1 ? "nota" : "notas"}`;

  const uniqueDates = new Set(ordered.map(item => item.date)).size;
  elements.dateCount.textContent = String(uniqueDates);
  const otherCount = ordered.filter(item => item.category === "Outras Despesas").length;
  const warnings = [];
  if (uniqueDates > MAX_DATES) warnings.push(`O modelo aceita no máximo ${MAX_DATES} datas diferentes.`);
  if (otherCount > MAX_OTHER_EXPENSES) warnings.push(`O modelo aceita no máximo ${MAX_OTHER_EXPENSES} linhas em Outras Despesas.`);
  elements.limitWarning.textContent = warnings.join(" ");
  elements.limitWarning.classList.toggle("hidden", warnings.length === 0);

  const summary = groupedSummary(ordered);
  elements.summaryList.replaceChildren();
  if (!summary.length) {
    elements.summaryList.textContent = "Nenhuma despesa registrada.";
    elements.summaryList.className = "summary-list empty-state";
  } else {
    elements.summaryList.className = "summary-list";
    summary.forEach(group => {
      const row = document.createElement("div");
      row.className = "summary-row";
      const label = document.createElement("span");
      label.textContent = `${formatDate(group.date)} · ${displayCategory(group.category)}`;
      const total = document.createElement("strong");
      total.textContent = formatMoney(group.total);
      row.append(label, total);
      elements.summaryList.append(row);
    });
  }

  elements.entriesList.replaceChildren();
  if (!ordered.length) {
    elements.entriesList.textContent = "As notas aparecerão aqui.";
    elements.entriesList.className = "entries-list empty-state";
  } else {
    elements.entriesList.className = "entries-list";
    ordered.forEach(expense => {
      const fragment = elements.entryTemplate.content.cloneNode(true);
      const photo = fragment.querySelector(".entry-photo");
      if (expense.photoBlob) {
        const photoUrl = URL.createObjectURL(expense.photoBlob);
        objectUrls.push(photoUrl);
        photo.src = photoUrl;
        photo.alt = `Nota de ${expense.category}`;
      }
      fragment.querySelector(".entry-date").textContent = formatDate(expense.date);
      fragment.querySelector(".entry-amount").textContent = formatMoney(expense.amount);
      fragment.querySelector(".entry-category").textContent = displayCategory(expense.category);
      fragment.querySelector(".entry-description").textContent = expense.description || "Sem descrição";
      fragment.querySelector(".edit-button").addEventListener("click", () => startEdit(expense.id));
      fragment.querySelector(".delete-button").addEventListener("click", () => removeExpense(expense.id));
      elements.entriesList.append(fragment);
    });
  }
}

function resetForm() {
  elements.form.reset();
  elements.editingId.value = "";
  elements.date.value = todayIso();
  elements.category.value = CATEGORIES[1];
  elements.saveExpense.textContent = "Adicionar Despesa";
  elements.saveExpense.innerHTML = '<span class="button-step">3</span><span>Adicionar Despesa</span>';
  elements.cancelEdit.classList.add("hidden");
  elements.photoStatus.textContent = "Nenhuma Foto Selecionada";
  retainedPhoto = null;
  retainedPhotoName = "";
  retainedPhotoType = "";
  hidePhotoPreview();
}

function startEdit(id) {
  const expense = expenses.find(item => item.id === id);
  if (!expense) return;
  elements.editingId.value = expense.id;
  elements.date.value = expense.date;
  elements.amount.value = Number(expense.amount).toFixed(2).replace(".", ",");
  elements.category.value = expense.category;
  elements.description.value = expense.description || "";
  retainedPhoto = expense.photoBlob;
  retainedPhotoName = expense.photoName;
  retainedPhotoType = expense.photoType;
  elements.photoStatus.textContent = `Foto Atual: ${expense.photoName}`;
  elements.saveExpense.innerHTML = '<span class="button-step">3</span><span>Salvar Alteração</span>';
  elements.cancelEdit.classList.remove("hidden");
  showPhotoPreview(expense.photoBlob);
  window.scrollTo({ top: 0, behavior: "smooth" });
}

async function removeExpense(id) {
  if (!confirm("Excluir este lançamento?")) return;
  await deleteExpenseRecord(id);
  expenses = await loadExpenses();
  render();
}

async function handleSubmit(event) {
  event.preventDefault();
  elements.formMessage.textContent = "";
  const amount = parseAmount(elements.amount.value);
  const selectedPhoto = elements.photo.files[0] || retainedPhoto;
  if (amount <= 0) {
    elements.formMessage.textContent = "Informe um valor maior que zero.";
    return;
  }
  if (!selectedPhoto) {
    elements.formMessage.textContent = "Selecione a foto da nota.";
    return;
  }

  elements.saveExpense.disabled = true;
  try {
    const editingId = elements.editingId.value;
    const existing = expenses.find(item => item.id === editingId);
    const selectedFile = elements.photo.files[0];
    const record = {
      id: editingId || createRecordId(),
      date: elements.date.value,
      amount,
      category: elements.category.value,
      description: elements.description.value.trim(),
      photoBlob: selectedFile || retainedPhoto,
      photoName: selectedFile?.name || retainedPhotoName || "nota.jpg",
      photoType: selectedFile?.type || retainedPhotoType || "image/jpeg",
      createdAt: existing?.createdAt || new Date().toISOString(),
      updatedAt: new Date().toISOString()
    };

    await saveExpenseRecord(record);
    expenses = await loadExpenses();
    resetForm();
    render();
    elements.formMessage.textContent = "Despesa salva no aparelho.";
  } catch (error) {
    elements.formMessage.textContent = `Não foi possível salvar: ${error.message || "erro desconhecido"}.`;
  } finally {
    elements.saveExpense.disabled = false;
  }
}

function uint16(value) {
  return new Uint8Array([value & 255, (value >>> 8) & 255]);
}

function uint32(value) {
  return new Uint8Array([
    value & 255,
    (value >>> 8) & 255,
    (value >>> 16) & 255,
    (value >>> 24) & 255
  ]);
}

function concatBytes(parts) {
  const length = parts.reduce((sum, part) => sum + part.length, 0);
  const output = new Uint8Array(length);
  let offset = 0;
  parts.forEach(part => {
    output.set(part, offset);
    offset += part.length;
  });
  return output;
}

const crcTable = (() => {
  const table = new Uint32Array(256);
  for (let n = 0; n < 256; n++) {
    let c = n;
    for (let k = 0; k < 8; k++) c = (c & 1) ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    table[n] = c >>> 0;
  }
  return table;
})();

function crc32(bytes) {
  let crc = 0xffffffff;
  for (const byte of bytes) crc = crcTable[(crc ^ byte) & 255] ^ (crc >>> 8);
  return (crc ^ 0xffffffff) >>> 0;
}

function dosDateTime(date = new Date()) {
  const year = Math.max(1980, date.getFullYear());
  const dosTime = (date.getHours() << 11) | (date.getMinutes() << 5) | Math.floor(date.getSeconds() / 2);
  const dosDate = ((year - 1980) << 9) | ((date.getMonth() + 1) << 5) | date.getDate();
  return { dosTime, dosDate };
}

async function createZip(files) {
  const encoder = new TextEncoder();
  const localParts = [];
  const centralParts = [];
  let offset = 0;
  const { dosTime, dosDate } = dosDateTime();

  for (const file of files) {
    const nameBytes = encoder.encode(file.name);
    const data = file.data instanceof Uint8Array ? file.data : new Uint8Array(await file.data.arrayBuffer());
    const checksum = crc32(data);
    const localHeader = concatBytes([
      uint32(0x04034b50), uint16(20), uint16(0x0800), uint16(0), uint16(dosTime), uint16(dosDate),
      uint32(checksum), uint32(data.length), uint32(data.length), uint16(nameBytes.length), uint16(0), nameBytes
    ]);
    localParts.push(localHeader, data);

    const centralHeader = concatBytes([
      uint32(0x02014b50), uint16(20), uint16(20), uint16(0x0800), uint16(0), uint16(dosTime), uint16(dosDate),
      uint32(checksum), uint32(data.length), uint32(data.length), uint16(nameBytes.length), uint16(0), uint16(0),
      uint16(0), uint16(0), uint32(0), uint32(offset), nameBytes
    ]);
    centralParts.push(centralHeader);
    offset += localHeader.length + data.length;
  }

  const centralDirectory = concatBytes(centralParts);
  const localDirectory = concatBytes(localParts);
  const end = concatBytes([
    uint32(0x06054b50), uint16(0), uint16(0), uint16(files.length), uint16(files.length),
    uint32(centralDirectory.length), uint32(localDirectory.length), uint16(0)
  ]);
  return new Blob([localDirectory, centralDirectory, end], { type: "application/zip" });
}

function csvEscape(value) {
  const text = String(value ?? "");
  return `"${text.replace(/"/g, '""')}"`;
}

async function exportPackage() {
  elements.exportMessage.textContent = "";
  const ordered = sortedExpenses();
  if (!ordered.length) {
    elements.exportMessage.textContent = "Registre pelo menos uma despesa.";
    return;
  }
  const missingPhoto = ordered.find(item => !item.photoBlob);
  if (missingPhoto) {
    elements.exportMessage.textContent = "Existe lançamento sem foto.";
    return;
  }

  const files = [];
  const manifestExpenses = [];
  for (let index = 0; index < ordered.length; index++) {
    const expense = ordered[index];
    const sequence = index + 1;
    const extension = fileExtension(expense.photoName, expense.photoType);
    const photoPath = `notas/${String(sequence).padStart(3, "0")}_${expense.date}_${slug(expense.category)}${extension}`;
    files.push({ name: photoPath, data: expense.photoBlob });
    manifestExpenses.push({
      sequence,
      date: expense.date,
      category: expense.category,
      description: expense.description || "",
      amount: Number(expense.amount).toFixed(2),
      photo: photoPath,
      created_at: expense.createdAt
    });
  }

  const summary = groupedSummary(ordered).map(group => ({
    date: group.date,
    category: group.category,
    components: group.amounts.map(value => Number(value).toFixed(2)),
    formula: formulaForAmounts(group.amounts),
    total: group.total.toFixed(2)
  }));
  const manifest = {
    format: "relatorio-despesas-mobile",
    version: 1,
    exported_at: new Date().toISOString(),
    expenses: manifestExpenses,
    summary
  };
  const csvRows = [
    ["sequencia", "data", "categoria", "descricao", "valor", "foto"].map(csvEscape).join(";"),
    ...manifestExpenses.map(item => [
      item.sequence,
      item.date,
      item.category,
      item.description,
      item.amount.replace(".", ","),
      item.photo
    ].map(csvEscape).join(";"))
  ];
  const encoder = new TextEncoder();
  files.unshift(
    { name: "dados_relatorio_mobile.json", data: encoder.encode(JSON.stringify(manifest, null, 2)) },
    { name: "resumo_lancamentos.csv", data: encoder.encode("\ufeff" + csvRows.join("\r\n")) }
  );

  elements.exportPackage.disabled = true;
  elements.exportPackage.textContent = "Gerando pacote...";
  try {
    const zip = await createZip(files);
    const url = URL.createObjectURL(zip);
    const link = document.createElement("a");
    link.href = url;
    const lastExpenseDate = ordered[ordered.length - 1].date.replaceAll("-", "");
    link.download = `${lastExpenseDate}_RELATÓRIO_DESPESAS_HENRIQUE.zip`;
    document.body.append(link);
    link.click();
    link.remove();
    setTimeout(() => URL.revokeObjectURL(url), 30000);
    elements.exportMessage.textContent = "Pacote gerado. Envie o ZIP para o PC.";
  } catch (error) {
    elements.exportMessage.textContent = `Não foi possível gerar o pacote: ${error.message}`;
  } finally {
    elements.exportPackage.disabled = false;
    elements.exportPackage.textContent = "Gerar Pacote para o PC";
  }
}

async function initialize() {
  CATEGORIES.forEach(category => {
    const option = document.createElement("option");
    option.value = category;
    option.textContent = displayCategory(category);
    elements.category.append(option);
  });
  elements.date.value = todayIso();
  elements.category.value = CATEGORIES[1];
  expenses = await loadExpenses();
  render();

  elements.form.addEventListener("submit", handleSubmit);
  elements.cancelEdit.addEventListener("click", resetForm);
  elements.photo.addEventListener("change", () => {
    const file = elements.photo.files[0];
    elements.photoStatus.textContent = file ? `Selecionada: ${file.name}` : "Nenhuma Foto Selecionada";
    if (file) showPhotoPreview(file);
    else if (retainedPhoto) showPhotoPreview(retainedPhoto);
    else hidePhotoPreview();
  });
  elements.exportPackage.addEventListener("click", exportPackage);
  elements.clearAll.addEventListener("click", openClearConfirmation);
  elements.cancelClear.addEventListener("click", closeClearConfirmation);
  elements.clearConfirmInput.addEventListener("input", () => {
    elements.confirmClear.disabled = elements.clearConfirmInput.value.trim().toUpperCase() !== "APAGAR";
  });
  elements.clearConfirmDialog.addEventListener("cancel", event => {
    event.preventDefault();
    closeClearConfirmation();
  });
  elements.confirmClear.addEventListener("click", async () => {
    if (elements.clearConfirmInput.value.trim().toUpperCase() !== "APAGAR") return;
    await clearExpenseRecords();
    expenses = [];
    resetForm();
    render();
    closeClearConfirmation();
    elements.exportMessage.textContent = "Lançamentos apagados.";
  });

  if ("serviceWorker" in navigator) {
    navigator.serviceWorker.register("./sw.js").catch(() => {});
  }
}

initialize().catch(error => {
  elements.formMessage.textContent = `Falha ao abrir o armazenamento local: ${error.message}`;
});
