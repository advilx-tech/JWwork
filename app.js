"use strict";

/* ===================== State ===================== */

const state = {
  transactions: [],      // normalized {date, dept, user, merchant, category, amount, memo}
  headers: [],
  rawRows: [],
  mapping: {},
  filters: { start: null, end: null, depts: null, categories: null, search: "" },
  sort: {},               // per-table sort state: {tableId: {key, dir}}
  charts: {},              // Chart.js instances keyed by canvas id
  currentTab: "upload",
};

const FIELD_LABELS = {
  date: "거래일자 (필수)",
  amount: "금액 (필수)",
  dept: "부서",
  user: "사용자",
  merchant: "가맹점",
  category: "카테고리/업종",
  memo: "메모",
};

const FIELD_SYNONYMS = {
  date: ["거래일자", "승인일자", "이용일자", "사용일자", "결제일자", "일자", "날짜", "date"],
  amount: ["이용금액", "승인금액", "사용금액", "결제금액", "청구금액", "금액", "amount"],
  dept: ["부서명", "부서", "소속", "팀명", "팀", "부문", "department"],
  user: ["사용자명", "사용자", "카드소지자", "카드사용자", "이용자", "성명", "이름", "user", "name"],
  merchant: ["가맹점명", "가맹점", "사용처", "상호명", "상호", "merchant"],
  category: ["업종명", "업종", "카테고리", "구분", "분류", "category"],
  memo: ["적요", "메모", "비고사항", "비고", "memo"],
};

const CHART_COLORS = ["#2f5fed", "#17b3a3", "#f2994a", "#9b59d0", "#e0473e", "#2ea1e0", "#d9b31c", "#6b7688", "#1f9d55", "#c2185b"];

const fmtKRW = (n) => Math.round(n).toLocaleString("ko-KR") + "원";
const fmtNum = (n) => Math.round(n).toLocaleString("ko-KR");
const fmtDate = (d) => `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
const monthKey = (d) => `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}`;

/* ===================== Parsing helpers ===================== */

function guessField(header) {
  const h = String(header).replace(/\s/g, "").toLowerCase();
  for (const [field, syns] of Object.entries(FIELD_SYNONYMS)) {
    for (const syn of syns) {
      if (h.includes(syn.toLowerCase())) return field;
    }
  }
  return null;
}

function guessMapping(headers) {
  const mapping = {};
  const used = new Set();
  for (const field of Object.keys(FIELD_SYNONYMS)) mapping[field] = null;
  for (const header of headers) {
    const f = guessField(header);
    if (f && !used.has(f)) {
      mapping[f] = header;
      used.add(f);
    }
  }
  return mapping;
}

function excelSerialToDate(n) {
  const utcDays = Math.floor(n - 25569);
  const utcValue = utcDays * 86400;
  return new Date(utcValue * 1000);
}

function parseDateFlexible(value) {
  if (value instanceof Date && !isNaN(value)) return value;
  if (typeof value === "number") return excelSerialToDate(value);
  if (typeof value === "string") {
    const s = value.trim();
    let m = s.match(/^(\d{4})[.\-\/년]\s*(\d{1,2})[.\-\/월]\s*(\d{1,2})/);
    if (m) return new Date(+m[1], +m[2] - 1, +m[3]);
    m = s.match(/^(\d{4})(\d{2})(\d{2})$/);
    if (m) return new Date(+m[1], +m[2] - 1, +m[3]);
    const d = new Date(s);
    if (!isNaN(d)) return d;
  }
  return null;
}

function parseAmountFlexible(value) {
  if (typeof value === "number") return value;
  if (typeof value === "string") {
    let s = value.replace(/[,\s원₩]/g, "");
    let neg = false;
    if (/^\(.*\)$/.test(s)) {
      neg = true;
      s = s.slice(1, -1);
    }
    if (s === "" || s === "-") return NaN;
    const n = parseFloat(s);
    if (isNaN(n)) return NaN;
    return neg ? -n : n;
  }
  return NaN;
}

/* ===================== File upload / parsing ===================== */

const dropZone = document.getElementById("dropZone");
const fileInput = document.getElementById("fileInput");
const uploadError = document.getElementById("uploadError");
const mappingCard = document.getElementById("mappingCard");

dropZone.addEventListener("click", () => fileInput.click());
dropZone.addEventListener("dragover", (e) => { e.preventDefault(); dropZone.classList.add("dragover"); });
dropZone.addEventListener("dragleave", () => dropZone.classList.remove("dragover"));
dropZone.addEventListener("drop", (e) => {
  e.preventDefault();
  dropZone.classList.remove("dragover");
  if (e.dataTransfer.files.length) handleFile(e.dataTransfer.files[0]);
});
fileInput.addEventListener("change", (e) => {
  if (e.target.files.length) handleFile(e.target.files[0]);
});

document.getElementById("sampleDataBtn").addEventListener("click", () => {
  loadSampleData();
});

function showUploadError(msg) {
  uploadError.textContent = msg;
  uploadError.hidden = false;
}

function handleFile(file) {
  uploadError.hidden = true;
  const reader = new FileReader();
  reader.onload = (e) => {
    try {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: "array", cellDates: true });
      const sheetName = workbook.SheetNames[0];
      const sheet = workbook.Sheets[sheetName];
      const rows = XLSX.utils.sheet_to_json(sheet, { defval: "" });
      if (!rows.length) {
        showUploadError("파일에서 데이터를 찾을 수 없습니다. 첫 행이 헤더(열 이름)인지 확인해주세요.");
        return;
      }
      const headers = Object.keys(rows[0]);
      state.headers = headers;
      state.rawRows = rows;
      state.mapping = guessMapping(headers);
      renderMappingUI();
    } catch (err) {
      console.error(err);
      showUploadError("파일을 읽는 중 오류가 발생했습니다: " + err.message);
    }
  };
  reader.readAsArrayBuffer(file);
}

function renderMappingUI() {
  mappingCard.hidden = false;
  const fieldsEl = document.getElementById("mappingFields");
  fieldsEl.innerHTML = "";
  const requiredFields = ["date", "amount"];

  Object.keys(FIELD_LABELS).forEach((field) => {
    const wrap = document.createElement("div");
    wrap.className = "mapping-field";
    const label = document.createElement("label");
    label.textContent = FIELD_LABELS[field];
    wrap.appendChild(label);

    const select = document.createElement("select");
    select.dataset.field = field;
    const noneOpt = document.createElement("option");
    noneOpt.value = "";
    noneOpt.textContent = requiredFields.includes(field) ? "-- 선택하세요 --" : "-- 없음 --";
    select.appendChild(noneOpt);
    state.headers.forEach((h) => {
      const opt = document.createElement("option");
      opt.value = h;
      opt.textContent = h;
      if (state.mapping[field] === h) opt.selected = true;
      select.appendChild(opt);
    });
    wrap.appendChild(select);
    fieldsEl.appendChild(wrap);
  });

  const previewEl = document.getElementById("mappingPreview");
  const previewRows = state.rawRows.slice(0, 5);
  let html = "<table><thead><tr>" + state.headers.map((h) => `<th>${escapeHtml(h)}</th>`).join("") + "</tr></thead><tbody>";
  previewRows.forEach((r) => {
    html += "<tr>" + state.headers.map((h) => `<td>${escapeHtml(String(r[h]))}</td>`).join("") + "</tr>";
  });
  html += "</tbody></table>";
  previewEl.innerHTML = html;

  document.getElementById("confirmMappingBtn").onclick = confirmMapping;
  document.getElementById("cancelMappingBtn").onclick = () => {
    mappingCard.hidden = true;
    fileInput.value = "";
  };
}

function confirmMapping() {
  const selects = document.querySelectorAll("#mappingFields select");
  const mapping = {};
  selects.forEach((s) => { mapping[s.dataset.field] = s.value || null; });

  if (!mapping.date || !mapping.amount) {
    showUploadError("거래일자와 금액 열은 반드시 지정해야 합니다.");
    return;
  }

  const transactions = [];
  let skipped = 0;
  state.rawRows.forEach((row) => {
    const date = parseDateFlexible(row[mapping.date]);
    const amount = parseAmountFlexible(row[mapping.amount]);
    if (!date || isNaN(amount)) { skipped++; return; }
    transactions.push({
      date,
      amount,
      dept: mapping.dept ? String(row[mapping.dept] || "").trim() || "미분류" : "미분류",
      user: mapping.user ? String(row[mapping.user] || "").trim() || "미상" : "미상",
      merchant: mapping.merchant ? String(row[mapping.merchant] || "").trim() || "미상" : "미상",
      category: mapping.category ? String(row[mapping.category] || "").trim() || "기타" : "기타",
      memo: mapping.memo ? String(row[mapping.memo] || "").trim() : "",
    });
  });

  if (!transactions.length) {
    showUploadError("유효한 거래 데이터를 찾지 못했습니다. 열 매핑을 확인해주세요.");
    return;
  }

  state.transactions = transactions;
  mappingCard.hidden = true;
  document.getElementById("uploadError").hidden = true;
  if (skipped > 0) {
    console.warn(`${skipped}개 행을 건너뛰었습니다 (날짜/금액 파싱 실패).`);
  }
  onDataLoaded();
}

/* ===================== Sample data generator ===================== */

function loadSampleData() {
  const depts = ["영업1팀", "영업2팀", "마케팅팀", "개발팀", "경영지원팀"];
  const usersByDept = {
    "영업1팀": ["김민준", "이서연", "박도윤"],
    "영업2팀": ["최지우", "정하은"],
    "마케팅팀": ["강수아", "조은우"],
    "개발팀": ["윤서준", "임지호", "한소율"],
    "경영지원팀": ["오유진", "신하윤"],
  };
  const categories = ["식비", "교통비", "숙박비", "사무용품", "접대비", "복리후생", "기타"];
  const merchantsByCategory = {
    "식비": ["한식당 미소", "스타벅스", "김밥천국", "이디야커피", "본죽"],
    "교통비": ["카카오T", "SRT", "대한항공", "타임리스주차장"],
    "숙박비": ["롯데호텔", "신라스테이", "야놀자"],
    "사무용품": ["오피스디포", "다이소", "교보문고"],
    "접대비": ["강남고깃집", "일식오마카세", "와인바 루체"],
    "복리후생": ["GS25", "헬스클럽 짐박스"],
    "기타": ["다이소", "우체국"],
  };

  const transactions = [];
  const today = new Date();
  const startDate = new Date(today.getFullYear(), today.getMonth() - 5, 1);
  const dayRange = Math.floor((today - startDate) / 86400000);

  for (let i = 0; i < 420; i++) {
    const dept = depts[Math.floor(Math.random() * depts.length)];
    const users = usersByDept[dept];
    const user = users[Math.floor(Math.random() * users.length)];
    const category = categories[Math.floor(Math.random() * categories.length)];
    const merchants = merchantsByCategory[category];
    const merchant = merchants[Math.floor(Math.random() * merchants.length)];
    const date = new Date(startDate.getTime() + Math.floor(Math.random() * dayRange) * 86400000);
    let base = { "식비": 25000, "교통비": 18000, "숙박비": 150000, "사무용품": 40000, "접대비": 120000, "복리후생": 20000, "기타": 15000 }[category];
    let amount = Math.round((base * (0.5 + Math.random())) / 100) * 100;
    transactions.push({ date, amount, dept, user, merchant, category, memo: "" });
  }

  // inject statistical outliers
  for (let i = 0; i < 6; i++) {
    const dept = depts[Math.floor(Math.random() * depts.length)];
    const user = usersByDept[dept][0];
    const date = new Date(startDate.getTime() + Math.floor(Math.random() * dayRange) * 86400000);
    transactions.push({ date, amount: 2500000 + Math.random() * 1500000, dept, user, merchant: "프리미엄 오마카세", category: "접대비", memo: "" });
  }

  // inject duplicate-looking transactions
  for (let i = 0; i < 4; i++) {
    const dept = depts[0];
    const user = usersByDept[dept][1];
    const date = new Date(startDate.getTime() + Math.floor(Math.random() * dayRange) * 86400000);
    const dup = { date, amount: 88000, dept, user, merchant: "한식당 미소", category: "식비", memo: "" };
    transactions.push({ ...dup });
    transactions.push({ ...dup, date: new Date(date) });
  }

  state.transactions = transactions;
  onDataLoaded();
}

/* ===================== Data loaded / filter setup ===================== */

function onDataLoaded() {
  document.querySelectorAll(".tab-btn").forEach((btn) => { btn.disabled = false; });
  document.getElementById("filterBar").hidden = false;
  document.getElementById("resetDataBtn").hidden = false;
  setupFilterOptions();
  updateDataSummary();
  switchTab("dashboard");
}

function setupFilterOptions() {
  const depts = [...new Set(state.transactions.map((t) => t.dept))].sort();
  const categories = [...new Set(state.transactions.map((t) => t.category))].sort();

  const deptSel = document.getElementById("filterDept");
  const catSel = document.getElementById("filterCategory");
  deptSel.innerHTML = depts.map((d) => `<option value="${escapeHtml(d)}">${escapeHtml(d)}</option>`).join("");
  catSel.innerHTML = categories.map((c) => `<option value="${escapeHtml(c)}">${escapeHtml(c)}</option>`).join("");

  const minDate = state.transactions.reduce((min, t) => (t.date < min ? t.date : min), state.transactions[0].date);
  const maxDate = state.transactions.reduce((max, t) => (t.date > max ? t.date : max), state.transactions[0].date);
  document.getElementById("filterStart").value = fmtDate(minDate);
  document.getElementById("filterEnd").value = fmtDate(maxDate);

  state.filters = { start: minDate, end: maxDate, depts: null, categories: null, search: "" };
}

function updateDataSummary() {
  const el = document.getElementById("dataSummary");
  el.textContent = `총 ${fmtNum(state.transactions.length)}건 로드됨`;
}

document.getElementById("applyFilterBtn").addEventListener("click", () => {
  const start = document.getElementById("filterStart").value;
  const end = document.getElementById("filterEnd").value;
  const deptSel = [...document.getElementById("filterDept").selectedOptions].map((o) => o.value);
  const catSel = [...document.getElementById("filterCategory").selectedOptions].map((o) => o.value);
  const search = document.getElementById("filterSearch").value.trim().toLowerCase();

  state.filters.start = start ? new Date(start + "T00:00:00") : null;
  state.filters.end = end ? new Date(end + "T23:59:59") : null;
  state.filters.depts = deptSel.length ? new Set(deptSel) : null;
  state.filters.categories = catSel.length ? new Set(catSel) : null;
  state.filters.search = search;

  renderCurrentTab();
});

document.getElementById("clearFilterBtn").addEventListener("click", () => {
  document.getElementById("filterDept").selectedIndex = -1;
  document.getElementById("filterCategory").selectedIndex = -1;
  document.getElementById("filterSearch").value = "";
  setupFilterOptions();
  renderCurrentTab();
});

document.getElementById("resetDataBtn").addEventListener("click", () => {
  if (!confirm("현재 데이터를 지우고 새 파일을 업로드하시겠습니까?")) return;
  state.transactions = [];
  document.querySelectorAll(".tab-btn").forEach((btn) => {
    if (btn.dataset.tab !== "upload") btn.disabled = true;
  });
  document.getElementById("filterBar").hidden = true;
  document.getElementById("resetDataBtn").hidden = true;
  document.getElementById("dataSummary").textContent = "";
  fileInput.value = "";
  switchTab("upload");
});

function getFilteredData() {
  const f = state.filters;
  return state.transactions.filter((t) => {
    if (f.start && t.date < f.start) return false;
    if (f.end && t.date > f.end) return false;
    if (f.depts && !f.depts.has(t.dept)) return false;
    if (f.categories && !f.categories.has(t.category)) return false;
    if (f.search) {
      const hay = (t.user + " " + t.merchant).toLowerCase();
      if (!hay.includes(f.search)) return false;
    }
    return true;
  });
}

/* ===================== Tab switching ===================== */

document.getElementById("tabNav").addEventListener("click", (e) => {
  const btn = e.target.closest(".tab-btn");
  if (!btn || btn.disabled) return;
  switchTab(btn.dataset.tab);
});

function switchTab(tab) {
  state.currentTab = tab;
  document.querySelectorAll(".tab-btn").forEach((b) => b.classList.toggle("active", b.dataset.tab === tab));
  document.querySelectorAll(".tab-panel").forEach((p) => p.classList.toggle("active", p.id === "tab-" + tab));
  renderCurrentTab();
}

function renderCurrentTab() {
  if (state.transactions.length === 0) return;
  const data = getFilteredData();
  switch (state.currentTab) {
    case "dashboard": renderDashboard(data); break;
    case "dept": renderDeptTab(data); break;
    case "category": renderCategoryTab(data); break;
    case "trend": renderTrendTab(data); break;
    case "anomaly": renderAnomalyTab(data); break;
    case "table": renderTableTab(data); break;
  }
}

/* ===================== Aggregation helpers ===================== */

function sumBy(data, keyFn) {
  const map = new Map();
  for (const t of data) {
    const key = keyFn(t);
    if (!map.has(key)) map.set(key, { sum: 0, count: 0 });
    const entry = map.get(key);
    entry.sum += t.amount;
    entry.count += 1;
  }
  return map;
}

function topN(map, n) {
  return [...map.entries()].sort((a, b) => b[1].sum - a[1].sum).slice(0, n);
}

function destroyChart(id) {
  if (state.charts[id]) {
    state.charts[id].destroy();
    delete state.charts[id];
  }
}

function makeChart(id, config) {
  destroyChart(id);
  const ctx = document.getElementById(id);
  if (!ctx) return;
  state.charts[id] = new Chart(ctx, config);
}

function sortableTable(wrapId, tableId, columns, rows) {
  const sortState = state.sort[tableId] || { key: null, dir: 1 };
  let sortedRows = rows;
  if (sortState.key) {
    const col = columns.find((c) => c.key === sortState.key);
    sortedRows = [...rows].sort((a, b) => {
      const av = col.value(a), bv = col.value(b);
      if (typeof av === "number") return (av - bv) * sortState.dir;
      return String(av).localeCompare(String(bv)) * sortState.dir;
    });
  }

  let html = `<table id="${tableId}"><thead><tr>`;
  columns.forEach((c) => {
    const arrow = sortState.key === c.key ? (sortState.dir === 1 ? " ▲" : " ▼") : "";
    html += `<th class="${c.num ? "num" : ""}" data-key="${c.key}">${escapeHtml(c.label)}${arrow}</th>`;
  });
  html += "</tr></thead><tbody>";
  sortedRows.forEach((row) => {
    html += "<tr>" + columns.map((c) => `<td class="${c.num ? "num" : ""}">${c.render ? c.render(row) : escapeHtml(String(c.value(row)))}</td>`).join("") + "</tr>";
  });
  html += "</tbody></table>";

  const wrap = document.getElementById(wrapId);
  if (!rows.length) {
    wrap.innerHTML = `<div class="empty-state">표시할 데이터가 없습니다.</div>`;
    return;
  }
  wrap.innerHTML = html;
  wrap.querySelectorAll("th").forEach((th) => {
    th.addEventListener("click", () => {
      const key = th.dataset.key;
      const prev = state.sort[tableId] || { key: null, dir: 1 };
      state.sort[tableId] = { key, dir: prev.key === key ? -prev.dir : -1 };
      sortableTable(wrapId, tableId, columns, rows);
    });
  });
}

/* ===================== Dashboard tab ===================== */

function renderDashboard(data) {
  const kpiGrid = document.getElementById("kpiGrid");
  const total = data.reduce((s, t) => s + t.amount, 0);
  const count = data.length;
  const avg = count ? total / count : 0;
  const deptCount = new Set(data.map((t) => t.dept)).size;
  const userCount = new Set(data.map((t) => t.user)).size;

  kpiGrid.innerHTML = `
    <div class="kpi-card"><div class="kpi-label">총 지출액</div><div class="kpi-value">${fmtKRW(total)}</div></div>
    <div class="kpi-card"><div class="kpi-label">거래 건수</div><div class="kpi-value">${fmtNum(count)}건</div></div>
    <div class="kpi-card"><div class="kpi-label">건당 평균 지출</div><div class="kpi-value">${fmtKRW(avg)}</div></div>
    <div class="kpi-card"><div class="kpi-label">부서 수</div><div class="kpi-value">${deptCount}</div></div>
    <div class="kpi-card"><div class="kpi-label">사용자 수</div><div class="kpi-value">${userCount}</div></div>
  `;

  const deptMap = sumBy(data, (t) => t.dept);
  const deptTop = topN(deptMap, 10);
  makeChart("chartDeptOverview", {
    type: "bar",
    data: {
      labels: deptTop.map((d) => d[0]),
      datasets: [{ label: "지출액", data: deptTop.map((d) => d[1].sum), backgroundColor: CHART_COLORS[0] }],
    },
    options: baseBarOptions(),
  });

  const catMap = sumBy(data, (t) => t.category);
  const catEntries = [...catMap.entries()].sort((a, b) => b[1].sum - a[1].sum);
  makeChart("chartCategoryOverview", {
    type: "doughnut",
    data: {
      labels: catEntries.map((c) => c[0]),
      datasets: [{ data: catEntries.map((c) => c[1].sum), backgroundColor: CHART_COLORS }],
    },
    options: { responsive: true, plugins: { legend: { position: "right" } } },
  });

  renderMonthlyTrendChart("chartTrendOverview", data);
}

function baseBarOptions(indexAxis) {
  return {
    responsive: true,
    indexAxis: indexAxis || "x",
    plugins: { legend: { display: false }, tooltip: { callbacks: { label: (ctx) => fmtKRW(ctx.parsed.y ?? ctx.parsed.x) } } },
    scales: {
      x: indexAxis === "y" ? { ticks: { callback: (v) => fmtNum(v) } } : {},
      y: indexAxis === "y" ? {} : { ticks: { callback: (v) => fmtNum(v) } },
    },
  };
}

function renderMonthlyTrendChart(canvasId, data) {
  const monthMap = new Map();
  data.forEach((t) => {
    const key = monthKey(t.date);
    if (!monthMap.has(key)) monthMap.set(key, 0);
    monthMap.set(key, monthMap.get(key) + t.amount);
  });
  const months = [...monthMap.keys()].sort();
  makeChart(canvasId, {
    type: "line",
    data: {
      labels: months,
      datasets: [{
        label: "월별 지출액",
        data: months.map((m) => monthMap.get(m)),
        borderColor: CHART_COLORS[0],
        backgroundColor: CHART_COLORS[0] + "33",
        fill: true,
        tension: 0.3,
      }],
    },
    options: baseBarOptions(),
  });
}

/* ===================== Dept/User tab ===================== */

function renderDeptTab(data) {
  const deptMap = sumBy(data, (t) => t.dept);
  const deptEntries = [...deptMap.entries()].sort((a, b) => b[1].sum - a[1].sum);

  makeChart("chartDeptFull", {
    type: "bar",
    data: {
      labels: deptEntries.map((d) => d[0]),
      datasets: [{ label: "지출액", data: deptEntries.map((d) => d[1].sum), backgroundColor: CHART_COLORS[0] }],
    },
    options: baseBarOptions("y"),
  });

  const deptSelect = document.getElementById("deptDrilldown");
  const currentDepts = [...deptMap.keys()].sort();
  const prevVal = deptSelect.value;
  deptSelect.innerHTML = `<option value="">전체 부서</option>` + currentDepts.map((d) => `<option value="${escapeHtml(d)}">${escapeHtml(d)}</option>`).join("");
  if (currentDepts.includes(prevVal)) deptSelect.value = prevVal;
  deptSelect.onchange = () => renderUserChart(data, deptSelect.value);
  renderUserChart(data, deptSelect.value);

  sortableTable("deptTableWrap", "deptTable", [
    { key: "dept", label: "부서", value: (r) => r[0] },
    { key: "sum", label: "지출액", value: (r) => r[1].sum, num: true, render: (r) => fmtKRW(r[1].sum) },
    { key: "count", label: "건수", value: (r) => r[1].count, num: true, render: (r) => fmtNum(r[1].count) },
    { key: "avg", label: "건당 평균", value: (r) => r[1].sum / r[1].count, num: true, render: (r) => fmtKRW(r[1].sum / r[1].count) },
  ], deptEntries);
}

function renderUserChart(data, dept) {
  const filtered = dept ? data.filter((t) => t.dept === dept) : data;
  const userMap = sumBy(filtered, (t) => t.user);
  const top = topN(userMap, 15);
  makeChart("chartUserFull", {
    type: "bar",
    data: {
      labels: top.map((d) => d[0]),
      datasets: [{ label: "지출액", data: top.map((d) => d[1].sum), backgroundColor: CHART_COLORS[1] }],
    },
    options: baseBarOptions("y"),
  });
}

/* ===================== Category/Merchant tab ===================== */

function renderCategoryTab(data) {
  const catMap = sumBy(data, (t) => t.category);
  const catEntries = [...catMap.entries()].sort((a, b) => b[1].sum - a[1].sum);

  makeChart("chartCategoryFull", {
    type: "bar",
    data: {
      labels: catEntries.map((d) => d[0]),
      datasets: [{ label: "지출액", data: catEntries.map((d) => d[1].sum), backgroundColor: CHART_COLORS[2] }],
    },
    options: baseBarOptions("y"),
  });

  const merchMap = sumBy(data, (t) => t.merchant);
  const merchTop = topN(merchMap, 15);
  makeChart("chartMerchantFull", {
    type: "bar",
    data: {
      labels: merchTop.map((d) => d[0]),
      datasets: [{ label: "지출액", data: merchTop.map((d) => d[1].sum), backgroundColor: CHART_COLORS[3] }],
    },
    options: baseBarOptions("y"),
  });

  sortableTable("categoryTableWrap", "categoryTable", [
    { key: "category", label: "카테고리", value: (r) => r[0] },
    { key: "sum", label: "지출액", value: (r) => r[1].sum, num: true, render: (r) => fmtKRW(r[1].sum) },
    { key: "count", label: "건수", value: (r) => r[1].count, num: true, render: (r) => fmtNum(r[1].count) },
    { key: "avg", label: "건당 평균", value: (r) => r[1].sum / r[1].count, num: true, render: (r) => fmtKRW(r[1].sum / r[1].count) },
  ], catEntries);
}

/* ===================== Trend tab ===================== */

function renderTrendTab(data) {
  renderMonthlyTrendChart("chartTrendFull", data);

  const weekdayNames = ["일", "월", "화", "수", "목", "금", "토"];
  const weekdaySum = new Array(7).fill(0);
  data.forEach((t) => { weekdaySum[t.date.getDay()] += t.amount; });
  makeChart("chartWeekday", {
    type: "bar",
    data: {
      labels: weekdayNames,
      datasets: [{
        label: "지출액",
        data: weekdaySum,
        backgroundColor: weekdayNames.map((_, i) => (i === 0 || i === 6) ? CHART_COLORS[4] : CHART_COLORS[0]),
      }],
    },
    options: baseBarOptions(),
  });

  const monthMap = sumBy(data, (t) => monthKey(t.date));
  const monthEntries = [...monthMap.entries()].sort((a, b) => a[0].localeCompare(b[0]));
  sortableTable("trendTableWrap", "trendTable", [
    { key: "month", label: "월", value: (r) => r[0] },
    { key: "sum", label: "지출액", value: (r) => r[1].sum, num: true, render: (r) => fmtKRW(r[1].sum) },
    { key: "count", label: "건수", value: (r) => r[1].count, num: true, render: (r) => fmtNum(r[1].count) },
  ], monthEntries);
}

/* ===================== Anomaly detection tab ===================== */

function detectAnomalies(data) {
  // stats per category for z-score
  const catStats = new Map();
  const catGroups = sumBy(data, (t) => t.category);
  catGroups.forEach((_, cat) => {
    const amounts = data.filter((t) => t.category === cat).map((t) => t.amount);
    const mean = amounts.reduce((s, a) => s + a, 0) / amounts.length;
    const variance = amounts.reduce((s, a) => s + (a - mean) ** 2, 0) / amounts.length;
    catStats.set(cat, { mean, std: Math.sqrt(variance) });
  });

  // duplicate detection key
  const dupMap = new Map();
  data.forEach((t) => {
    const key = [t.user, t.merchant, t.amount, fmtDate(t.date)].join("|");
    dupMap.set(key, (dupMap.get(key) || 0) + 1);
  });

  const results = [];
  data.forEach((t) => {
    const reasons = [];
    let severity = 0;

    const stats = catStats.get(t.category);
    let z = 0;
    if (stats && stats.std > 0) {
      z = (t.amount - stats.mean) / stats.std;
      if (z > 2.5) {
        reasons.push(`통계적 이상치 (${t.category} 평균 대비 ${z.toFixed(1)}σ)`);
        severity += Math.min(z, 10);
      }
    }

    const day = t.date.getDay();
    if (day === 0 || day === 6) {
      reasons.push("주말 사용");
      severity += 1;
    }

    const dupKey = [t.user, t.merchant, t.amount, fmtDate(t.date)].join("|");
    if (dupMap.get(dupKey) > 1) {
      reasons.push("중복 의심 거래 (동일 사용자·가맹점·금액·날짜)");
      severity += 3;
    }

    if (reasons.length) {
      results.push({ ...t, reasons, severity, zscore: z });
    }
  });

  results.sort((a, b) => b.severity - a.severity);
  return results;
}

function renderAnomalyTab(data) {
  const anomalies = detectAnomalies(data);
  const kpiGrid = document.getElementById("anomalyKpiGrid");
  const anomalyTotal = anomalies.reduce((s, a) => s + a.amount, 0);
  const highSeverity = anomalies.filter((a) => a.severity >= 3).length;

  kpiGrid.innerHTML = `
    <div class="kpi-card"><div class="kpi-label">탐지된 이상 거래</div><div class="kpi-value danger">${fmtNum(anomalies.length)}건</div><div class="kpi-sub">전체 ${fmtNum(data.length)}건 중 ${data.length ? ((anomalies.length / data.length) * 100).toFixed(1) : 0}%</div></div>
    <div class="kpi-card"><div class="kpi-label">이상 거래 총액</div><div class="kpi-value danger">${fmtKRW(anomalyTotal)}</div></div>
    <div class="kpi-card"><div class="kpi-label">고위험 거래</div><div class="kpi-value danger">${fmtNum(highSeverity)}건</div><div class="kpi-sub">중복 의심 또는 이상치 강도 높음</div></div>
  `;

  sortableTable("anomalyTableWrap", "anomalyTable", [
    { key: "date", label: "날짜", value: (r) => r.date.getTime(), render: (r) => fmtDate(r.date) },
    { key: "dept", label: "부서", value: (r) => r.dept },
    { key: "user", label: "사용자", value: (r) => r.user },
    { key: "merchant", label: "가맹점", value: (r) => r.merchant },
    { key: "category", label: "카테고리", value: (r) => r.category },
    { key: "amount", label: "금액", value: (r) => r.amount, num: true, render: (r) => fmtKRW(r.amount) },
    {
      key: "reasons", label: "탐지 사유", value: (r) => r.reasons.join(", "),
      render: (r) => r.reasons.map((reason) => {
        const cls = reason.includes("중복") ? "badge-danger" : reason.includes("이상치") ? "badge-warning" : "badge-muted";
        return `<span class="badge ${cls}">${escapeHtml(reason)}</span>`;
      }).join(""),
    },
  ], anomalies);
}

/* ===================== Full table tab ===================== */

function renderTableTab(data) {
  document.getElementById("tableCount").textContent = `(${fmtNum(data.length)}건)`;
  const sorted = [...data].sort((a, b) => b.date - a.date);
  sortableTable("fullTableWrap", "fullTable", [
    { key: "date", label: "날짜", value: (r) => r.date.getTime(), render: (r) => fmtDate(r.date) },
    { key: "dept", label: "부서", value: (r) => r.dept },
    { key: "user", label: "사용자", value: (r) => r.user },
    { key: "merchant", label: "가맹점", value: (r) => r.merchant },
    { key: "category", label: "카테고리", value: (r) => r.category },
    { key: "amount", label: "금액", value: (r) => r.amount, num: true, render: (r) => fmtKRW(r.amount) },
    { key: "memo", label: "메모", value: (r) => r.memo },
  ], sorted);

  document.getElementById("exportCsvBtn").onclick = () => exportCsv(sorted);
}

function exportCsv(rows) {
  const header = ["날짜", "부서", "사용자", "가맹점", "카테고리", "금액", "메모"];
  const lines = [header.join(",")];
  rows.forEach((r) => {
    const cells = [fmtDate(r.date), r.dept, r.user, r.merchant, r.category, r.amount, r.memo]
      .map((c) => `"${String(c).replace(/"/g, '""')}"`);
    lines.push(cells.join(","));
  });
  const csv = "﻿" + lines.join("\n");
  const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = "법인카드_사용내역_필터결과.csv";
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

/* ===================== Utils ===================== */

function escapeHtml(str) {
  return String(str).replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
}
