const state = {
  data: null,
  filtered: [],
};

const elements = {
  sourcePath: document.getElementById("source-path"),
  generatedAt: document.getElementById("generated-at"),
  statUsers: document.getElementById("stat-users"),
  statDoors: document.getElementById("stat-doors"),
  statVisible: document.getElementById("stat-visible"),
  rowCount: document.getElementById("row-count"),
  status: document.getElementById("status"),
  search: document.getElementById("search"),
  doorFilter: document.getElementById("door-filter"),
  doorOnly: document.getElementById("door-only"),
  refresh: document.getElementById("refresh"),
  exportBtn: document.getElementById("export"),
  syncBtn: document.getElementById("sync-btn"),
  syncInput: document.getElementById("sync-input"),
  table: document.getElementById("dashboard-table"),
};

function setStatus(message, tone = "") {
  elements.status.textContent = message;
  elements.status.dataset.tone = tone;
}

function updateMeta(meta) {
  elements.sourcePath.textContent = meta.doors_path || "--";
  elements.generatedAt.textContent = meta.generated_at || "--";
  elements.statUsers.textContent = meta.user_count ?? "--";
  elements.statDoors.textContent = meta.door_count ?? "--";
}

function updateCounts(visible, total) {
  const totalText = total ?? visible;
  elements.rowCount.textContent = `${visible} / ${totalText} rows`;
  elements.statVisible.textContent = visible;
}

function buildDoorOptions(doors) {
  const current = elements.doorFilter.value;
  elements.doorFilter.innerHTML = "";

  const allOption = document.createElement("option");
  allOption.value = "all";
  allOption.textContent = "All doors";
  elements.doorFilter.appendChild(allOption);

  doors.forEach((door) => {
    const option = document.createElement("option");
    option.value = door;
    option.textContent = door;
    elements.doorFilter.appendChild(option);
  });

  if (doors.includes(current)) {
    elements.doorFilter.value = current;
  } else {
    elements.doorFilter.value = "all";
  }

  updateDoorToggle();
}

function updateDoorToggle() {
  const isAll = elements.doorFilter.value === "all";
  elements.doorOnly.disabled = isAll;
}

function renderTable(columns, rows) {
  const thead = elements.table.querySelector("thead");
  const tbody = elements.table.querySelector("tbody");
  thead.innerHTML = "";
  tbody.innerHTML = "";

  if (!columns || columns.length === 0) {
    return;
  }

  const headerRow = document.createElement("tr");
  columns.forEach((col) => {
    const th = document.createElement("th");
    th.textContent = col;
    headerRow.appendChild(th);
  });
  thead.appendChild(headerRow);

  rows.forEach((row) => {
    const tr = document.createElement("tr");
    row.forEach((cell, index) => {
      const td = document.createElement("td");
      td.textContent = cell;
      if (index === 1) {
        td.classList.add("cell-mono");
      }
      if (index >= 4) {
        td.classList.add("cell-door");
      }
      tr.appendChild(td);
    });
    tbody.appendChild(tr);
  });
}

function applyFilters() {
  if (!state.data) {
    return;
  }

  const query = elements.search.value.trim().toLowerCase();
  const selectedDoor = elements.doorFilter.value;
  const onlyAccess = elements.doorOnly.checked;
  const doorIndex = selectedDoor === "all" ? -1 : state.data.columns.indexOf(selectedDoor);

  const filtered = state.data.rows.filter((row) => {
    if (query) {
      const haystack = row.join(" ").toLowerCase();
      if (!haystack.includes(query)) {
        return false;
      }
    }

    if (selectedDoor !== "all" && onlyAccess) {
      if (doorIndex === -1 || row[doorIndex] !== "X") {
        return false;
      }
    }

    return true;
  });

  state.filtered = filtered;
  renderTable(state.data.columns, filtered);
  updateCounts(filtered.length, state.data.meta?.user_count ?? state.data.rows.length);
}

async function exportXlsx() {
  if (!state.data) {
    return;
  }

  const params = new URLSearchParams();
  const query = elements.search.value.trim();
  const door = elements.doorFilter.value;

  if (query) {
    params.set("q", query);
  }
  if (door && door !== "all") {
    params.set("door", door);
  }
  if (elements.doorOnly.checked) {
    params.set("only_access", "1");
  }
  params.set("ts", Date.now());

  const url = `/api/export.xlsx?${params.toString()}`;

  setStatus("Preparing export...", "loading");
  try {
    const response = await fetch(url);
    if (!response.ok) {
      const payload = await response.json();
      throw new Error(payload.error || "Export failed.");
    }

    const blob = await response.blob();
    const link = document.createElement("a");
    const stamp = new Date().toISOString().slice(0, 19).replace(/[:T]/g, "-");
    const disposition = response.headers.get("content-disposition") || "";
    const match = disposition.match(/filename=([^;]+)/i);
    const filename = match ? match[1].replace(/\"/g, "") : `door_access_${stamp}.xlsx`;

    link.href = URL.createObjectURL(blob);
    link.download = filename;
    document.body.appendChild(link);
    link.click();
    link.remove();
    URL.revokeObjectURL(link.href);
    setStatus("");
  } catch (error) {
    setStatus(error.message, "error");
  }
}

async function handleUpload() {
  const files = elements.syncInput.files;
  if (!files.length) return;

  const formData = new FormData();
  for (const file of files) {
    formData.append("files", file);
  }

  setStatus(`Syncing ${files.length} files...`, "loading");
  try {
    const response = await fetch("/api/upload", {
      method: "POST",
      body: formData,
    });
    const payload = await response.json();

    if (!response.ok) {
      throw new Error(payload.error || "Sync failed.");
    }

    setStatus(payload.message || "Sync successful!", "success");
    elements.syncInput.value = ""; // reset
    setTimeout(() => {
      loadDashboard();
    }, 1500);
  } catch (error) {
    setStatus(error.message, "error");
    elements.syncInput.value = ""; // reset
  }
}

async function loadDashboard() {
  setStatus("Loading data...", "loading");
  try {
    const response = await fetch(`/api/dashboard?ts=${Date.now()}`);
    const payload = await response.json();

    if (!response.ok) {
      throw new Error(payload.error || "Failed to load dashboard data.");
    }

    state.data = payload;
    updateMeta(payload.meta || {});
    buildDoorOptions(payload.doors || []);
    applyFilters();
    setStatus("");
  } catch (error) {
    setStatus(error.message, "error");
  }
}

window.addEventListener("DOMContentLoaded", () => {
  elements.search.addEventListener("input", applyFilters);
  elements.doorFilter.addEventListener("change", () => {
    updateDoorToggle();
    applyFilters();
  });
  elements.doorOnly.addEventListener("change", applyFilters);
  elements.refresh.addEventListener("click", loadDashboard);
  elements.exportBtn.addEventListener("click", exportXlsx);

  if (elements.syncBtn && elements.syncInput) {
    elements.syncBtn.addEventListener("click", () => elements.syncInput.click());
    elements.syncInput.addEventListener("change", handleUpload);
  }

  loadDashboard();
});
