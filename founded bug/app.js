const state = {
  bugs: Array.isArray(window.BUG_DATA) ? window.BUG_DATA.map(normalizeRow) : [],
  manifest: [
    { method: "SQLess", file: "SQLess.xlsx" },
    { method: "QTRAN", file: "QTRAN.xlsx" },
    { method: "ValScope", file: "ValScope.xlsx" },
    { method: "SmartFuzz", file: "SmartFuzz.xlsx" },
    { method: "FPMT", file: "FPMT.xlsx" },
    { method: "DMLScope", file: "DMLScope.xlsx" },
    { method: "SchemaMorph", file: "SchemaMorph.xlsx" },
    { method: "FMU", file: "FMU.xlsx" },
    { method: "RIFT", file: "RIFT.xlsx" },
    { method: "VECT", file: "VECT.xlsx" },
  ],
  search: "",
  program: "all",
  type: "all",
  status: "all",
  method: "all",
};

const elements = {
  body: document.querySelector("#bug-table-body"),
  empty: document.querySelector("#empty-state"),
  resultCount: document.querySelector("#result-count"),
  totalCount: document.querySelector("#total-count"),
  search: document.querySelector("#search-input"),
  program: document.querySelector("#program-filter"),
  type: document.querySelector("#type-filter"),
  status: document.querySelector("#status-filter"),
  methodTabs: document.querySelector("#method-tabs"),
  downloads: document.querySelector("#downloads"),
};

function normalizeRow(row) {
  return {
    program: String(row.Program ?? row.program ?? ""),
    id: row.ID ?? row.id ?? "",
    bugId: String(row["Bug ID"] ?? row.bugId ?? ""),
    link: String(row["Bug Link"] ?? row.link ?? ""),
    type: String(row["Bug Type"] ?? row.type ?? ""),
    status: String(row.Status ?? row.status ?? ""),
    method: String(row.Method ?? row.method ?? ""),
  };
}

function escapeHtml(value) {
  return String(value).replace(/[&<>'"]/g, (character) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", "'": "&#39;", '"': "&quot;" })[character]);
}

function fillSelect(select, values) {
  select.querySelectorAll("option:not(:first-child)").forEach((option) => option.remove());
  [...new Set(values)].filter(Boolean).sort().forEach((value) => {
    const option = document.createElement("option");
    option.value = value;
    option.textContent = value;
    select.append(option);
  });
}

function filteredBugs() {
  const term = state.search.trim().toLowerCase();
  return state.bugs.filter((bug) => {
    const matchesSearch = !term || [bug.program, bug.bugId, bug.type, bug.status, bug.method]
      .some((value) => value.toLowerCase().includes(term));
    return matchesSearch
      && (state.program === "all" || bug.program === state.program)
      && (state.type === "all" || bug.type === state.type)
      && (state.status === "all" || bug.status === state.status)
      && (state.method === "all" || bug.method === state.method);
  });
}

function render() {
  const bugs = filteredBugs();
  elements.totalCount.textContent = state.bugs.length;
  elements.resultCount.textContent = `${bugs.length} of ${state.bugs.length} records`;
  elements.body.innerHTML = bugs.map((bug, index) => `<tr>
    <td>${index + 1}</td>
    <td><span class="bug-id">${escapeHtml(bug.bugId)}</span></td>
    <td>${escapeHtml(bug.program)}</td>
    <td>${escapeHtml(bug.type)}</td>
    <td><span class="status status-${bug.status.toLowerCase().replaceAll(" ", "-")}">${escapeHtml(bug.status)}</span></td>
    <td>${escapeHtml(bug.method)}</td>
    <td><a class="bug-link" href="${escapeHtml(bug.link)}" target="_blank" rel="noreferrer">Open ↗</a></td>
  </tr>`).join("");
  elements.empty.hidden = bugs.length !== 0;
}

function renderMethodTabs() {
  const methods = [...new Set(state.bugs.map((bug) => bug.method))].filter(Boolean);
  const items = [
    { method: "all", label: "All", count: state.bugs.length },
    ...methods.map((method) => ({ method, label: method, count: state.bugs.filter((bug) => bug.method === method).length })),
  ];
  elements.methodTabs.innerHTML = items.map((item) => `<button class="method-tab${state.method === item.method ? " is-active" : ""}" type="button" data-method="${escapeHtml(item.method)}">${escapeHtml(item.label)} <span>${item.count}</span></button>`).join("");
  elements.downloads.innerHTML = `<span>Excel:</span>${state.manifest.map((item) => `<a href="./data/${escapeHtml(item.file)}" download>${escapeHtml(item.method)}</a>`).join("")}`;
}

function refreshOptions() {
  fillSelect(elements.program, state.bugs.map((bug) => bug.program));
  fillSelect(elements.type, state.bugs.map((bug) => bug.type));
  fillSelect(elements.status, state.bugs.map((bug) => bug.status));
}

function rowsFromExcelBuffer(buffer) {
  const workbook = XLSX.read(buffer, { type: "array" });
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  return XLSX.utils.sheet_to_json(sheet, { defval: "" }).map(normalizeRow);
}

function applyExcelData(manifest, datasets) {
  const rows = datasets.flat().filter((bug) => bug.program && bug.bugId);
  if (!rows.length) throw new Error("No valid bug records found");

  state.manifest = manifest;
  state.bugs = rows;
  refreshOptions();
  renderMethodTabs();
  render();
  return rows.length;
}

async function loadExcelFiles() {
  if (!window.XLSX) return false;

  const requestOptions = { cache: "no-store" };
  const cacheKey = `?reload=${Date.now()}`;
  try {
    const manifestResponse = await fetch(`./data/workbooks.json${cacheKey}`, requestOptions);
    if (!manifestResponse.ok) throw new Error("workbooks.json");
    const manifest = await manifestResponse.json();
    state.manifest = manifest;
    const datasets = await Promise.all(manifest.map(async (item) => {
      const response = await fetch(`./data/${item.file}${cacheKey}`, requestOptions);
      if (!response.ok) throw new Error(item.file);
      return rowsFromExcelBuffer(await response.arrayBuffer());
    }));
    applyExcelData(manifest, datasets);
    return true;
  } catch (error) {
    // The local fallback remains visible if Excel loading is unavailable.
    console.error("Excel reload failed:", error);
    return false;
  }
}

elements.search.addEventListener("input", (event) => { state.search = event.target.value; render(); });
elements.program.addEventListener("change", (event) => { state.program = event.target.value; render(); });
elements.type.addEventListener("change", (event) => { state.type = event.target.value; render(); });
elements.status.addEventListener("change", (event) => { state.status = event.target.value; render(); });
elements.methodTabs.addEventListener("click", (event) => {
  const button = event.target.closest("[data-method]");
  if (!button) return;
  state.method = button.dataset.method;
  document.querySelectorAll(".method-tab").forEach((tab) => tab.classList.toggle("is-active", tab === button));
  render();
});
document.querySelector("#reset-button").addEventListener("click", () => {
  Object.assign(state, { search: "", program: "all", type: "all", status: "all", method: "all" });
  elements.search.value = "";
  elements.program.value = "all";
  elements.type.value = "all";
  elements.status.value = "all";
  document.querySelectorAll(".method-tab").forEach((tab) => tab.classList.toggle("is-active", tab.dataset.method === "all"));
  render();
});

refreshOptions();
renderMethodTabs();
render();
if (window.location.protocol !== "file:") loadExcelFiles();
