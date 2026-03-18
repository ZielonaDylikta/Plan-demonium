const comboInput = document.getElementById("comboInput");
const comboList = document.getElementById("comboList");
const comboWrap = document.getElementById("comboWrap");
const overviewBtn = document.getElementById("overviewBtn");
const summaryDiv = document.getElementById("summary");
const drilldownDiv = document.getElementById("drilldown");
const chartHint = document.getElementById("chartHint");
const overviewDiv = document.getElementById("overview");
const changeStripDiv = document.getElementById("changeStrip");
const portfolioDiv = document.getElementById("portfolio");
const riskStripDiv = document.getElementById("riskStrip");
const insightStripDiv = document.getElementById("insightStrip");
const triagePanelDiv = document.getElementById("triagePanel");
const overviewShell = document.getElementById("overviewShell");
const modeBadge = document.getElementById("modeBadge");
const modeTitle = document.getElementById("modeTitle");
const modeSubtitle = document.getElementById("modeSubtitle");
const portfolioView = document.getElementById("portfolioView");
const planView = document.getElementById("planView");
const workloadView = document.getElementById("workloadView");
const workloadBtn = document.getElementById("workloadBtn");
const workloadBackBtn = document.getElementById("workloadBackBtn");
const backBtn = document.getElementById("backBtn");
const csvBtn = document.getElementById("csvBtn");
const pdfBtn = document.getElementById("pdfBtn");
const pngBtn = document.getElementById("pngBtn");

let chart = null;
let workloadChart = null;
let currentStats = [];
let currentPlanId = null;
let overviewData = [];
let allPlans = [];
let trendData = {};
let workloadSnapshot = [];
let deltasLoaded = false;
let deltaSnapshot = {
  currentDate: null,
  previousDate: null,
  portfolio: null,
  byPlan: {},
};
let currentView = "overview"; // overview | plan | workload
let currentPlanTitle = "";

let pinnedPlanIds = JSON.parse(localStorage.getItem("pinnedPlans") || "[]");

let overviewSortField = null;
let overviewSortDir = "asc";
let drillSortField = null;
let drillSortDir = "asc";
let drillFilterStatus = "all";
let drillFilterAssignee = "all";
let activeDrillData = null;
let selectedBucketIndex = null;

const COLORS = {
  notStarted: "#d1d5db",
  inProgress: "#0284c7",
  completed: "#16a34a",
};

const prefersReducedMotion = window.matchMedia("(prefers-reduced-motion: reduce)").matches;

function escapeHtml(value) {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/\"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function hexToRgba(hex, alpha) {
  const normalized = hex.replace("#", "");
  const bigint = Number.parseInt(normalized.length === 3
    ? normalized
        .split("")
        .map((c) => c + c)
        .join("")
    : normalized, 16);
  const r = (bigint >> 16) & 255;
  const g = (bigint >> 8) & 255;
  const b = bigint & 255;
  return `rgba(${r}, ${g}, ${b}, ${alpha})`;
}

function applyStaggerReveal(container, selector) {
  if (!container) return;
  const nodes = container.querySelectorAll(selector);
  nodes.forEach((node, idx) => {
    node.classList.add("reveal-item");
    node.style.setProperty("--reveal-delay", `${Math.min(idx * 55, 320)}ms`);
  });
}

function animateKpiNumbers(root) {
  if (!root) return;
  const nodes = root.querySelectorAll("[data-anim-number]");
  nodes.forEach((node) => {
    const target = Number.parseInt(node.dataset.animNumber, 10);
    const suffix = node.dataset.suffix || "";
    if (!Number.isFinite(target)) return;
    if (prefersReducedMotion) {
      node.textContent = `${target}${suffix}`;
      return;
    }
    const start = performance.now();
    const duration = 560;
    const tick = (ts) => {
      const progress = Math.min((ts - start) / duration, 1);
      const eased = 1 - Math.pow(1 - progress, 3);
      const value = Math.round(target * eased);
      node.textContent = `${value}${suffix}`;
      if (progress < 1) requestAnimationFrame(tick);
    };
    requestAnimationFrame(tick);
  });
}

function syncTriageActiveState() {
  const activeFilter = activeDrillData?.type === "filter" ? activeDrillData.filterType : null;
  triagePanelDiv.querySelectorAll(".triage-chip[data-filter]").forEach((chip) => {
    chip.classList.toggle("active", chip.dataset.filter === activeFilter);
  });
}

function applyChartSelection() {
  if (!chart) return;
  const datasetStyles = [
    { key: "notStarted", alpha: 0.95 },
    { key: "inProgress", alpha: 0.9 },
    { key: "completed", alpha: 0.9 },
  ];

  chart.data.datasets.forEach((dataset, datasetIdx) => {
    const style = datasetStyles[datasetIdx];
    const baseColor = COLORS[style.key];
    dataset.backgroundColor = chart.data.labels.map((_, idx) =>
      selectedBucketIndex === null || selectedBucketIndex === idx
        ? hexToRgba(baseColor, style.alpha)
        : hexToRgba(baseColor, 0.24)
    );
    dataset.borderColor = chart.data.labels.map((_, idx) =>
      selectedBucketIndex === null || selectedBucketIndex === idx
        ? hexToRgba(baseColor, 1)
        : hexToRgba(baseColor, 0.28)
    );
    dataset.borderWidth = chart.data.labels.map((_, idx) =>
      selectedBucketIndex === null || selectedBucketIndex === idx ? 1 : 0
    );
  });

  chart.update("none");
}

function computeHealthScore(plan) {
  if (plan.total === 0) return 100;
  const completion = plan.percentComplete;
  const overduePenalty = Math.round((plan.overdue / plan.total) * 100);
  const dueSoonPenalty = Math.round(((plan.dueSoon || 0) / plan.total) * 50);
  const unassignedPenalty = Math.round(((plan.unassigned || 0) / plan.total) * 60);
  return Math.max(0, Math.min(100, completion - overduePenalty - dueSoonPenalty - unassignedPenalty + 35));
}

function getPlanTotals(stats) {
  return stats.reduce(
    (acc, s) => ({
      notStarted: acc.notStarted + s.notStarted,
      inProgress: acc.inProgress + s.inProgress,
      completed: acc.completed + s.completed,
      total: acc.total + s.total,
      overdue: acc.overdue + s.overdue,
      dueSoon: acc.dueSoon + s.dueSoon,
      noDueDate: acc.noDueDate + s.noDueDate,
    }),
    {
      notStarted: 0,
      inProgress: 0,
      completed: 0,
      total: 0,
      overdue: 0,
      dueSoon: 0,
      noDueDate: 0,
    }
  );
}

function renderInsightStrip(stats) {
  if (!stats || stats.length === 0) {
    insightStripDiv.innerHTML = "";
    return;
  }

  const totals = getPlanTotals(stats);
  const pct = totals.total > 0 ? Math.round((totals.completed / totals.total) * 100) : 0;
  const bucketsAtRisk = stats.filter((b) => b.overdue > 0 || b.dueSoon > 0).length;
  const trendPoints = trendData[currentPlanId] || [];
  const trendDelta = trendPoints.length >= 2 ? trendPoints[trendPoints.length - 1].percentComplete - trendPoints[0].percentComplete : 0;
  const trendLabel = trendPoints.length >= 2 ? `${trendDelta >= 0 ? "+" : ""}${trendDelta}% since first snapshot` : "Collecting trend history";

  const assigneeLoad = {};
  stats.forEach((bucket) => {
    bucket.tasks.forEach((task) => {
      task.assignees.forEach((a) => {
        assigneeLoad[a.displayName] = (assigneeLoad[a.displayName] || 0) + 1;
      });
    });
  });
  const topAssignee =
    Object.entries(assigneeLoad).sort((a, b) => b[1] - a[1])[0] || null;

  insightStripDiv.innerHTML = `
    <article class="insight-card">
      <div class="kicker">Narrative</div>
      <div class="headline">${pct}% complete across ${stats.length} buckets</div>
      <div class="sub">${totals.completed} completed of ${totals.total} tasks</div>
    </article>
    <article class="insight-card">
      <div class="kicker">Risk Focus</div>
      <div class="headline">${bucketsAtRisk} buckets need attention</div>
      <div class="sub">${totals.overdue} overdue · ${totals.dueSoon} due soon</div>
    </article>
    <article class="insight-card">
      <div class="kicker">Momentum</div>
      <div class="headline">${trendLabel}</div>
      <div class="sub">Trend window: last 14 snapshots</div>
    </article>
    <article class="insight-card">
      <div class="kicker">Load Signal</div>
      <div class="headline">${topAssignee ? escapeHtml(topAssignee[0]) : "No assignees yet"}</div>
      <div class="sub">${topAssignee ? `${topAssignee[1]} assigned tasks` : "Assign owners to surface workload"}</div>
    </article>
  `;
  applyStaggerReveal(insightStripDiv, ".insight-card");
}

function renderTriagePanel(stats) {
  if (!stats || stats.length === 0) {
    triagePanelDiv.innerHTML = "";
    return;
  }

  const totals = getPlanTotals(stats);
  triagePanelDiv.innerHTML = `
    <button type="button" class="triage-chip" data-action="filter-tasks" data-filter="overdue">
      <span>Overdue</span><strong>${totals.overdue}</strong>
    </button>
    <button type="button" class="triage-chip" data-action="filter-tasks" data-filter="dueSoon">
      <span>Due Soon</span><strong>${totals.dueSoon}</strong>
    </button>
    <button type="button" class="triage-chip" data-action="filter-tasks" data-filter="inProgress">
      <span>In Progress</span><strong>${totals.inProgress}</strong>
    </button>
    <button type="button" class="triage-chip" data-action="filter-tasks" data-filter="notStarted">
      <span>Not Started</span><strong>${totals.notStarted}</strong>
    </button>
    <button type="button" class="triage-chip" data-action="filter-tasks" data-filter="unassigned">
      <span>Unassigned</span><strong>${stats.flatMap((b) => b.tasks).filter((t) => t.assignees.length === 0).length}</strong>
    </button>
  `;
  applyStaggerReveal(triagePanelDiv, ".triage-chip");
  syncTriageActiveState();
}

function getFilteredCurrentTasks(filterType) {
  const allTasks = currentStats.flatMap((b) => b.tasks);
  if (filterType === "overdue" || filterType === "dueSoon") {
    return allTasks.filter((t) => t.dueStatus === filterType);
  }
  if (filterType === "unassigned") {
    return allTasks.filter((t) => t.assignees.length === 0);
  }
  return allTasks.filter((t) => t.status === filterType);
}

function showView(view) {
  currentView = view;
  overviewShell.style.display = view === "overview" ? "block" : "none";
  portfolioView.style.display = view === "overview" ? "block" : "none";
  planView.style.display = view === "plan" ? "block" : "none";
  workloadView.style.display = view === "workload" ? "block" : "none";
  workloadBtn.classList.toggle("active", view === "workload");
  overviewBtn.classList.toggle("active", view === "overview");

  document.body.classList.remove("mode-overview", "mode-plan", "mode-workload");
  document.body.classList.add(`mode-${view}`);
  setModeContext(view);
}

function setModeContext(view, customTitle = "") {
  if (view === "overview") {
    modeBadge.textContent = "Overview";
    modeTitle.textContent = "Portfolio Overview";
    modeSubtitle.textContent = "Cross-project KPIs and plan health.";
    return;
  }

  if (view === "workload") {
    modeBadge.textContent = "Workload";
    modeTitle.textContent = "Team Workload";
    modeSubtitle.textContent = "Assignee load and overdue distribution.";
    return;
  }

  modeBadge.textContent = "Project";
  modeTitle.textContent = customTitle || (currentPlanTitle ? `Project ${currentPlanTitle}` : "Project Details");
  modeSubtitle.textContent = "Bucket progress and task drill-down.";
}

function togglePin(planId) {
  const idx = pinnedPlanIds.indexOf(planId);
  if (idx >= 0) pinnedPlanIds.splice(idx, 1);
  else pinnedPlanIds.push(planId);
  localStorage.setItem("pinnedPlans", JSON.stringify(pinnedPlanIds));
  renderOverview();
  renderPortfolio();
  renderRiskStrip();
  renderComboList();
}

function sortWithPins(arr) {
  return [...arr].sort((a, b) => {
    const aPin = pinnedPlanIds.includes(a.planId) ? 0 : 1;
    const bPin = pinnedPlanIds.includes(b.planId) ? 0 : 1;
    return aPin - bPin;
  });
}

let comboOpen = false;
let highlightIdx = -1;
let comboFilterText = "";

function getFilteredOverview() {
  const q = comboFilterText.toLowerCase().trim();
  let data = overviewData;
  if (q) data = data.filter((p) => p.title.toLowerCase().includes(q));
  return data;
}

function getFilteredPlans() {
  const q = comboFilterText.toLowerCase().trim();
  let plans = allPlans;
  if (q) plans = plans.filter((p) => p.title.toLowerCase().includes(q));
  return [...plans].sort((a, b) => {
    const aPin = pinnedPlanIds.includes(a.id) ? 0 : 1;
    const bPin = pinnedPlanIds.includes(b.id) ? 0 : 1;
    return aPin - bPin;
  });
}

function renderComboList() {
  const plans = getFilteredPlans();
  const q = comboFilterText.toLowerCase().trim();
  const showAllPlansOption = !q || "all plans overview".includes(q);
  const items = [
    ...(showAllPlansOption
      ? [
          {
            id: "__all__",
            title: "All Plans (Overview)",
            isOverview: true,
          },
        ]
      : []),
    ...plans.map((p) => ({ id: p.id, title: p.title, isOverview: false })),
  ];

  if (items.length === 0) {
    comboList.innerHTML =
      '<div style="padding:12px;color:var(--text-secondary);font-size:13px">No plans found</div>';
    return;
  }
  comboList.innerHTML = items
    .map(
      (item, i) => `<div class="combo-item${item.isOverview ? " combo-item-overview" : ""}${i === highlightIdx ? " highlighted" : ""}" data-plan-id="${item.id}" data-idx="${i}" role="option" tabindex="0" aria-selected="${
        i === highlightIdx
      }">
        <span class="plan-name">${escapeHtml(item.title)}</span>
      </div>`
    )
    .join("");
}

function openCombo() {
  comboOpen = true;
  highlightIdx = -1;
  comboList.classList.add("open");
  renderComboList();
}

function closeCombo() {
  comboOpen = false;
  comboList.classList.remove("open");
}

comboInput.addEventListener("focus", openCombo);
comboInput.addEventListener("input", () => {
  comboFilterText = comboInput.value;
  if (!comboOpen) openCombo();
  renderComboList();
  renderOverview();
  renderRiskStrip();
});

comboList.addEventListener("click", (e) => {
  const item = e.target.closest(".combo-item");
  if (!item) return;
  const planId = item.dataset.planId;
  if (planId === "__all__") {
    comboInput.value = "";
    comboFilterText = "";
    closeCombo();
    goBack();
    renderOverview();
    renderRiskStrip();
    return;
  }
  const plan = allPlans.find((p) => p.id === planId);
  if (!plan) return;
  comboInput.value = plan.title;
  comboFilterText = "";
  closeCombo();
  loadStats(planId);
});

comboList.addEventListener("keydown", (e) => {
  const item = e.target.closest(".combo-item");
  if (!item) return;
  if (e.key === "Enter" || e.key === " ") {
    e.preventDefault();
    item.click();
  }
});

comboInput.addEventListener("keydown", (e) => {
  const items = comboList.querySelectorAll(".combo-item");
  if (e.key === "ArrowDown") {
    e.preventDefault();
    highlightIdx = Math.min(highlightIdx + 1, items.length - 1);
    renderComboList();
  } else if (e.key === "ArrowUp") {
    e.preventDefault();
    highlightIdx = Math.max(highlightIdx - 1, 0);
    renderComboList();
  } else if (e.key === "Enter" && highlightIdx >= 0 && items[highlightIdx]) {
    e.preventDefault();
    items[highlightIdx].click();
  } else if (e.key === "Escape") {
    closeCombo();
    comboInput.blur();
  }
});

document.addEventListener("click", (e) => {
  if (!comboWrap.contains(e.target)) closeCombo();
});

async function loadPlans() {
  try {
    const res = await fetch("/api/plans");
    const plans = await res.json();
    if (plans.error) return;
    allPlans = plans;
  } catch {
    // silent
  }
}

function renderRiskStrip() {
  const d = getFilteredOverview();
  const totalOverdue = d.reduce((s, p) => s + p.overdue, 0);
  const totalDueSoon = d.reduce((s, p) => s + (p.dueSoon || 0), 0);
  const totalUnassigned = d.reduce((s, p) => s + (p.unassigned || 0), 0);
  const atRiskPlans = d.filter((p) => p.total > 0 && p.overdue / p.total > 0.15).length;
  const mostLoaded = workloadSnapshot.length > 0 ? workloadSnapshot[0] : null;
  const namedOwners = workloadSnapshot.filter((owner) => owner.displayName !== "Unassigned");
  const leastLoaded = namedOwners.length > 0 ? namedOwners[namedOwners.length - 1] : null;

  riskStripDiv.innerHTML = `
    <section class="risk-card" aria-live="polite">
      <div class="label">Overdue Tasks</div>
      <div class="value" style="color:var(--red)">${totalOverdue}</div>
      <div class="subtext">Across ${d.length} visible plans</div>
    </section>
    <section class="risk-card" aria-live="polite">
      <div class="label">Due Soon (7 days)</div>
      <div class="value" style="color:var(--amber)">${totalDueSoon}</div>
      <div class="subtext">Early warning window</div>
    </section>
    <section class="risk-card" aria-live="polite">
      <div class="label">Unassigned Work</div>
      <div class="value" style="color:var(--accent)">${totalUnassigned}</div>
      <div class="subtext">Tasks without owners</div>
    </section>
    <section class="risk-card" aria-live="polite">
      <div class="label">Most Loaded Owner</div>
      <div class="value" style="font-size:20px;color:var(--text)">${mostLoaded ? escapeHtml(mostLoaded.displayName) : "-"}</div>
      <div class="subtext">${mostLoaded ? `${mostLoaded.total} total tasks` : `${atRiskPlans} at-risk plans (>15% overdue)`}</div>
    </section>
    <section class="risk-card" aria-live="polite">
      <div class="label">Least Loaded Owner</div>
      <div class="value" style="font-size:20px;color:var(--text)">${leastLoaded ? escapeHtml(leastLoaded.displayName) : "-"}</div>
      <div class="subtext">${leastLoaded ? `${leastLoaded.total} total tasks` : "Waiting for workload data"}</div>
    </section>
  `;
  applyStaggerReveal(riskStripDiv, ".risk-card");
}

function formatDelta(value, positiveIsGood = true) {
  const num = Number(value || 0);
  const sign = num > 0 ? "+" : "";
  const tone =
    num === 0
      ? "trend-flat"
      : positiveIsGood
      ? num > 0
        ? "trend-up"
        : "trend-down"
      : num > 0
      ? "trend-down"
      : "trend-up";

  return {
    text: `${sign}${num}`,
    tone,
  };
}

function renderChangeStrip() {
  if (!changeStripDiv) return;

  if (!deltasLoaded) {
    changeStripDiv.innerHTML = '<span class="guided-state">Loading snapshot changes</span>';
    return;
  }

  if (!deltaSnapshot?.portfolio || !deltaSnapshot.previousDate || !deltaSnapshot.currentDate) {
    changeStripDiv.innerHTML = '<span class="guided-state">Collecting snapshot history</span>';
    return;
  }

  const completed = formatDelta(deltaSnapshot.portfolio.completedDelta, true);
  const overdue = formatDelta(deltaSnapshot.portfolio.overdueDelta, false);
  const unassigned = formatDelta(deltaSnapshot.portfolio.unassignedDelta, false);
  const dueSoon = formatDelta(deltaSnapshot.portfolio.dueSoonDelta, false);

  changeStripDiv.innerHTML = `
    <article class="change-card">
      <div class="kicker">Since last snapshot</div>
      <div class="headline">${escapeHtml(deltaSnapshot.previousDate)} → ${escapeHtml(deltaSnapshot.currentDate)}</div>
      <div class="sub">Portfolio movement across tracked plans</div>
    </article>
    <article class="change-card">
      <div class="kicker">Completed</div>
      <div class="headline ${completed.tone}">${completed.text}</div>
      <div class="sub">Tasks moved to done</div>
    </article>
    <article class="change-card">
      <div class="kicker">Overdue</div>
      <div class="headline ${overdue.tone}">${overdue.text}</div>
      <div class="sub">Change in overdue tasks</div>
    </article>
    <article class="change-card">
      <div class="kicker">Unassigned</div>
      <div class="headline ${unassigned.tone}">${unassigned.text}</div>
      <div class="sub">Ownership risk movement</div>
    </article>
    <article class="change-card">
      <div class="kicker">Due Soon</div>
      <div class="headline ${dueSoon.tone}">${dueSoon.text}</div>
      <div class="sub">7-day warning movement</div>
    </article>
  `;

  applyStaggerReveal(changeStripDiv, ".change-card");
}

function renderPortfolio() {
  const d = overviewData;
  const totalProjects = d.length;
  const totalTasks = d.reduce((s, p) => s + p.total, 0);
  const totalCompleted = d.reduce((s, p) => s + p.completed, 0);
  const totalOverdue = d.reduce((s, p) => s + p.overdue, 0);
  const totalDueSoon = d.reduce((s, p) => s + (p.dueSoon || 0), 0);
  const totalUnassigned = d.reduce((s, p) => s + (p.unassigned || 0), 0);
  const pct = totalTasks > 0 ? Math.round((totalCompleted / totalTasks) * 100) : 0;

  portfolioDiv.innerHTML = `
    <div class="portfolio-card">
      <div class="number" style="color:var(--accent)" data-anim-number="${totalProjects}">${totalProjects}</div>
      <div class="label">Projects</div>
    </div>
    <div class="portfolio-card">
      <div class="number" style="color:var(--text)" data-anim-number="${totalTasks}">${totalTasks}</div>
      <div class="label">Total Tasks</div>
    </div>
    <div class="portfolio-card">
      <div class="number" style="color:var(--green)" data-anim-number="${pct}" data-suffix="%">${pct}%</div>
      <div class="label">Completion</div>
    </div>
    <button class="portfolio-card clickable" type="button" data-action="show-overdue" aria-label="Show plans with overdue tasks">
      <div class="number" style="color:var(--red)" data-anim-number="${totalOverdue}">${totalOverdue}</div>
      <div class="label">Overdue</div>
    </button>
    <button class="portfolio-card clickable" type="button" data-action="show-due-soon" aria-label="Show plans with due soon tasks">
      <div class="number" style="color:var(--amber)" data-anim-number="${totalDueSoon}">${totalDueSoon}</div>
      <div class="label">Due Soon</div>
    </button>
    <button class="portfolio-card clickable" type="button" data-action="show-unassigned" aria-label="Show unassigned tasks">
      <div class="number" style="color:var(--grey)" data-anim-number="${totalUnassigned}">${totalUnassigned}</div>
      <div class="label">Unassigned</div>
    </button>
  `;
  applyStaggerReveal(portfolioDiv, ".portfolio-card");
  animateKpiNumbers(portfolioDiv);
}

async function showUnassigned() {
  drilldownDiv.innerHTML = '<div class="loading">Loading unassigned tasks...</div>';
  showView("plan");
  setModeContext("plan", "Project Unassigned Tasks");
  insightStripDiv.innerHTML = "";
  triagePanelDiv.innerHTML = "";
  planView.querySelector(".chart-container").style.display = "none";
  chartHint.textContent = "";
  summaryDiv.innerHTML = "";
  csvBtn.style.display = "none";

  try {
    const res = await fetch("/api/unassigned");
    const data = await res.json();
    if (data.error) {
      drilldownDiv.innerHTML = `<div class="error">${escapeHtml(data.error)}</div>`;
      return;
    }

    const statusLabel = {
      completed: "Completed",
      inProgress: "In Progress",
      notStarted: "Not Started",
    };
    const dueLabel = {
      overdue: "Overdue",
      dueSoon: "Due Soon",
      onTrack: "On Track",
      noDueDate: "No Due Date",
    };

    const rows = data.tasks
      .map(
        (t) => `<tr>
      <td>${escapeHtml(t.title)}</td>
      <td>${escapeHtml(t.planTitle)}</td>
      <td><span class="badge badge-${t.status}">${statusLabel[t.status]}</span></td>
      <td>${
        t.dueDateTime
          ? new Date(t.dueDateTime).toLocaleDateString()
          : '<span style="color:var(--grey)">None</span>'
      }</td>
      <td><span class="badge badge-${t.dueStatus}">${dueLabel[t.dueStatus]}</span></td>
    </tr>`
      )
      .join("");

    drilldownDiv.innerHTML = `
      <div class="drilldown-header">
        <h3>All Unassigned Tasks (${data.count})</h3>
        <button class="drilldown-close" type="button" data-action="go-back" aria-label="Close unassigned tasks">&times;</button>
      </div>
      <table>
        <thead><tr><th>Task</th><th>Plan</th><th>Status</th><th>Due Date</th><th>Due Status</th></tr></thead>
        <tbody>${rows}</tbody>
      </table>
    `;
  } catch {
    drilldownDiv.innerHTML = '<div class="error">Failed to load unassigned tasks</div>';
  }
}

function showOverdueAll() {
  const overduePlans = overviewData.filter((p) => p.overdue > 0);
  if (overduePlans.length === 0) return;

  drilldownDiv.innerHTML = "";
  showView("plan");
  setModeContext("plan", "Project Overdue Tasks");
  insightStripDiv.innerHTML = "";
  triagePanelDiv.innerHTML = "";
  planView.querySelector(".chart-container").style.display = "none";
  chartHint.textContent = "";
  summaryDiv.innerHTML = "";
  csvBtn.style.display = "none";

  const rows = overduePlans
    .map(
      (p) => `<tr class="overview-row" data-plan-id="${p.planId}" tabindex="0" role="button" aria-label="Open ${escapeHtml(
        p.title
      )} overdue tasks">
      <td><strong>${escapeHtml(p.title)}</strong></td>
      <td style="color:var(--red);font-weight:700">${p.overdue}</td>
      <td>${p.total}</td>
    </tr>`
    )
    .join("");

  drilldownDiv.innerHTML = `
    <div class="drilldown-header">
      <h3>Plans with Overdue Tasks</h3>
      <button class="drilldown-close" type="button" data-action="go-back" aria-label="Close overdue plans">&times;</button>
    </div>
    <table>
      <thead><tr><th>Plan</th><th>Overdue</th><th>Total</th></tr></thead>
      <tbody>${rows}</tbody>
    </table>
  `;
}

function showDueSoonAll() {
  const d = getFilteredOverview();
  const dueSoonPlans = d.filter((p) => (p.dueSoon || 0) > 0);

  if (dueSoonPlans.length === 0) {
    chartHint.textContent = "No due soon tasks in current view.";
    return;
  }

  drilldownDiv.innerHTML = "";
  showView("plan");
  setModeContext("plan", "Project Due Soon Tasks");
  insightStripDiv.innerHTML = "";
  triagePanelDiv.innerHTML = "";
  planView.querySelector(".chart-container").style.display = "none";
  chartHint.textContent = "";
  summaryDiv.innerHTML = "";
  csvBtn.style.display = "none";

  const rows = dueSoonPlans
    .map(
      (p) => `<tr class="overview-row" data-plan-id="${p.planId}" tabindex="0" role="button" aria-label="Open ${escapeHtml(
        p.title
      )} due soon tasks">
      <td><strong>${escapeHtml(p.title)}</strong></td>
      <td style="color:var(--amber);font-weight:700">${p.dueSoon || 0}</td>
      <td>${p.total}</td>
    </tr>`
    )
    .join("");

  drilldownDiv.innerHTML = `
    <div class="drilldown-header">
      <h3>Plans with Due Soon Tasks</h3>
      <button class="drilldown-close" type="button" data-action="go-back" aria-label="Close due soon plans">&times;</button>
    </div>
    <table>
      <thead><tr><th>Plan</th><th>Due Soon</th><th>Total</th></tr></thead>
      <tbody>${rows}</tbody>
    </table>
  `;
}

function selectPlan(planId) {
  const plan = allPlans.find((p) => p.id === planId);
  if (plan) comboInput.value = plan.title;
  comboFilterText = "";
  loadStats(planId);
}

function goBack() {
  showView("overview");
  currentPlanId = null;
  currentPlanTitle = "";
  comboInput.value = "";
  comboFilterText = "";
  drilldownDiv.innerHTML = "";
  activeDrillData = null;
  selectedBucketIndex = null;
  insightStripDiv.innerHTML = "";
  triagePanelDiv.innerHTML = "";
  chartHint.textContent = "";
  if (chart) {
    applyChartSelection();
  }
  renderRiskStrip();
}

async function loadOverview() {
  overviewDiv.innerHTML = '<div class="loading">Loading overview...</div>';
  deltasLoaded = false;
  renderChangeStrip();
  try {
    const res = await fetch("/api/overview");
    const data = await res.json();
    if (data.error) {
      overviewDiv.innerHTML = `<div class="error">${escapeHtml(data.error)}</div>`;
      return;
    }
    overviewData = data;
    renderPortfolio();
    renderOverview();
    loadDeltas();
    loadTrends();
    loadRiskWorkload();
  } catch {
    overviewDiv.innerHTML = '<div class="error">Failed to load overview</div>';
    deltasLoaded = true;
    renderChangeStrip();
  }
}

async function loadDeltas() {
  try {
    const res = await fetch("/api/deltas");
    const data = await res.json();
    if (!data.error) {
      deltaSnapshot = data;
    }
  } catch {
    deltaSnapshot = {
      currentDate: null,
      previousDate: null,
      portfolio: null,
      byPlan: {},
    };
  } finally {
    deltasLoaded = true;
    renderChangeStrip();
  }
}

async function loadRiskWorkload() {
  try {
    const res = await fetch("/api/workload");
    const data = await res.json();
    if (!data.error) {
      workloadSnapshot = data;
      renderRiskStrip();
    }
  } catch {
    renderRiskStrip();
  }
}

function renderOverview() {
  let data = getFilteredOverview();
  data = data.map((p) => ({
    ...p,
    healthScore: computeHealthScore(p),
  }));
  data = sortWithPins(data);

  if (overviewSortField) {
    const dir = overviewSortDir === "asc" ? 1 : -1;
    const pinned = data.filter((p) => pinnedPlanIds.includes(p.planId));
    const unpinned = data.filter((p) => !pinnedPlanIds.includes(p.planId));
    const sortFn = (a, b) => {
      const va = a[overviewSortField];
      const vb = b[overviewSortField];
      if (typeof va === "string") return dir * va.localeCompare(vb);
      return dir * (va - vb);
    };
    pinned.sort(sortFn);
    unpinned.sort(sortFn);
    data = [...pinned, ...unpinned];
  }

  if (data.length === 0) {
    overviewDiv.innerHTML = '<div class="guided-state" style="margin:10px">No plans found for current search</div>';
    return;
  }

  const arrow = (field) => {
    if (overviewSortField !== field) return "";
    return `<span class="sort-arrow">${overviewSortDir === "asc" ? "▲" : "▼"}</span>`;
  };

  const rows = data
    .map((p) => {
      const isPinned = pinnedPlanIds.includes(p.planId);
      const star = isPinned ? "★" : "☆";
      let dotColor = "var(--green)";
      if (p.healthScore < 60) dotColor = "var(--red)";
      else if (p.healthScore < 80) dotColor = "var(--amber)";

      return `<tr class="overview-row" data-plan-id="${p.planId}" tabindex="0" role="button" aria-label="Open plan ${escapeHtml(
        p.title
      )}">
      <td><button type="button" class="pin-star" data-action="toggle-pin" data-plan-id="${p.planId}" aria-label="${
        isPinned ? "Unpin" : "Pin"
      } ${escapeHtml(p.title)}">${star}</button></td>
      <td><span class="health-dot" style="background:${dotColor}" title="Health score ${p.healthScore}"></span></td>
      <td><strong>${escapeHtml(p.title)}</strong></td>
      <td>${p.total}</td>
      <td>${p.completed}</td>
      <td>${p.inProgress}</td>
      <td style="color:${
        p.overdue > 0 ? "var(--red)" : "var(--text-secondary)"
      };font-weight:${p.overdue > 0 ? "700" : "400"}">${p.overdue}</td>
      <td><span class="trend-badge ${
        p.healthScore >= 75 ? "trend-up" : p.healthScore >= 50 ? "trend-flat" : "trend-down"
      }">${p.healthScore}</span></td>
      <td>
        <span class="progress-cell">
          <span class="pct-bar"><span class="pct-bar-fill" style="width:${p.percentComplete}%"></span></span>
          <span class="progress-value">${p.percentComplete}%</span>
        </span>
      </td>
      <td class="sparkline-cell" data-plan-id="${p.planId}"><span class="guided-state">Collecting trend</span></td>
    </tr>`;
    })
    .join("");

  overviewDiv.innerHTML = `<table>
    <thead><tr>
      <th style="width:30px"></th>
      <th style="width:30px">Health</th>
      <th data-sort="title">Plan${arrow("title")}</th>
      <th data-sort="total">Total${arrow("total")}</th>
      <th data-sort="completed">Done${arrow("completed")}</th>
      <th data-sort="inProgress">In Prog${arrow("inProgress")}</th>
      <th data-sort="overdue">Overdue${arrow("overdue")}</th>
      <th data-sort="healthScore">Health${arrow("healthScore")}</th>
      <th data-sort="percentComplete">Progress${arrow("percentComplete")}</th>
      <th style="width:90px">Trend</th>
    </tr></thead>
    <tbody>${rows}</tbody>
  </table>`;

  renderSparklines();
  renderRiskStrip();
}

function sortOverview(field) {
  if (overviewSortField === field) {
    overviewSortDir = overviewSortDir === "asc" ? "desc" : "asc";
  } else {
    overviewSortField = field;
    overviewSortDir = "asc";
  }
  renderOverview();
}

async function loadTrends() {
  try {
    const res = await fetch("/api/trends");
    trendData = await res.json();
    renderSparklines();
  } catch {
    // silent fail for trends
  }
}

function renderSparklines() {
  document.querySelectorAll(".sparkline-cell").forEach((cell) => {
    const planId = cell.dataset.planId;
    const points = trendData[planId];
    if (!points || points.length < 2) {
      cell.innerHTML = '<span class="guided-state">Collecting trend</span>';
      return;
    }

    const first = points[0].percentComplete;
    const last = points[points.length - 1].percentComplete;
    const delta = last - first;
    const cls = delta > 0 ? "trend-up" : delta < 0 ? "trend-down" : "trend-flat";
    const icon = delta > 0 ? "↑" : delta < 0 ? "↓" : "→";
    const deltaText = `${delta >= 0 ? "+" : ""}${delta}%`;
    cell.innerHTML = `<span class="trend-badge ${cls}" title="${first}% to ${last}% over ${points.length} snapshots">${icon} ${deltaText}</span>`;
  });
}

async function loadStats(planId) {
  if (!planId) {
    goBack();
    return;
  }

  drilldownDiv.innerHTML = "";
  chartHint.textContent = "";

  try {
    const res = await fetch(`/api/plans/${encodeURIComponent(planId)}/stats`);
    const stats = await res.json();
    if (stats.error) {
      summaryDiv.innerHTML = `<div class="error">${escapeHtml(stats.error)}</div>`;
      return;
    }

    currentStats = stats;
    currentPlanId = planId;
    selectedBucketIndex = null;
    activeDrillData = null;
    const activePlan = allPlans.find((p) => p.id === planId);
    currentPlanTitle = activePlan ? activePlan.title : "";
    showView("plan");
    setModeContext("plan", currentPlanTitle ? `Project ${currentPlanTitle}` : "Project Details");
    planView.querySelector(".chart-container").style.display = "";
    csvBtn.style.display = "";
    renderInsightStrip(stats);
    renderTriagePanel(stats);
    renderChart(stats);
    renderSummary(stats);
    chartHint.textContent = "Click a bar to see task details";
  } catch {
    summaryDiv.innerHTML = '<div class="error">Failed to load stats</div>';
  }
}

function renderChart(stats) {
  const ctx = document.getElementById("chart").getContext("2d");
  if (chart) chart.destroy();

  chart = new Chart(ctx, {
    type: "bar",
    data: {
      labels: stats.map((s) => s.bucketName),
      datasets: [
        {
          label: "Not Started",
          data: stats.map((s) => s.notStarted),
          backgroundColor: stats.map(() => hexToRgba(COLORS.notStarted, 0.95)),
          borderRadius: 4,
        },
        {
          label: "In Progress",
          data: stats.map((s) => s.inProgress),
          backgroundColor: stats.map(() => hexToRgba(COLORS.inProgress, 0.9)),
          borderRadius: 4,
        },
        {
          label: "Completed",
          data: stats.map((s) => s.completed),
          backgroundColor: stats.map(() => hexToRgba(COLORS.completed, 0.9)),
          borderRadius: 4,
        },
      ],
    },
    options: {
      responsive: true,
      animation: { duration: 800, easing: "easeOutQuart" },
      onClick: (event, elements) => {
        if (elements.length === 0) return;
        selectedBucketIndex = elements[0].index;
        applyChartSelection();
        renderDrillDown(currentStats[elements[0].index]);
      },
      plugins: {
        title: {
          display: true,
          text: "Tasks per Bucket",
          font: { size: 16, weight: "700", family: "Space Grotesk" },
          color: "#1f2937",
          padding: { bottom: 16 },
        },
        legend: {
          position: "bottom",
          labels: {
            font: { size: 12, family: "Manrope" },
            color: "#6b7280",
            usePointStyle: true,
            pointStyle: "circle",
            padding: 20,
          },
        },
      },
      scales: {
        x: {
          stacked: true,
          grid: { display: false },
          ticks: { font: { size: 12, family: "Manrope" }, color: "#6b7280" },
        },
        y: {
          stacked: true,
          beginAtZero: true,
          ticks: { stepSize: 1, font: { size: 12, family: "Manrope" }, color: "#6b7280" },
          grid: { color: "#eee7da" },
        },
      },
    },
  });
  applyChartSelection();
}

function renderDrillDown(bucket) {
  activeDrillData = { type: "bucket", bucket };
  selectedBucketIndex = currentStats.findIndex((s) => s.bucketName === bucket.bucketName);
  applyChartSelection();
  drillSortField = null;
  drillFilterStatus = "all";
  drillFilterAssignee = "all";
  syncTriageActiveState();
  renderDrillTable();
}

function renderFilteredTasks(filterType) {
  activeDrillData = { type: "filter", filterType };
  selectedBucketIndex = null;
  applyChartSelection();
  drillSortField = null;
  drillFilterStatus = "all";
  drillFilterAssignee = "all";
  syncTriageActiveState();
  renderDrillTable();
}

function renderDrillTable() {
  if (!activeDrillData) return;

  const statusLabel = {
    completed: "Completed",
    inProgress: "In Progress",
    notStarted: "Not Started",
  };
  const dueLabel = {
    overdue: "Overdue",
    dueSoon: "Due Soon",
    onTrack: "On Track",
    noDueDate: "No Due Date",
  };
  const filterLabels = {
    overdue: "Overdue",
    dueSoon: "Due Soon",
    notStarted: "Not Started",
    inProgress: "In Progress",
    unassigned: "Unassigned",
  };

  let tasks;
  let title;
  let showBucket;

  if (activeDrillData.type === "bucket") {
    const b = activeDrillData.bucket;
    tasks = [...b.tasks];
    title = `${b.bucketName} - ${b.total} task${b.total !== 1 ? "s" : ""}`;
    showBucket = false;
  } else {
    const ft = activeDrillData.filterType;
    tasks = getFilteredCurrentTasks(ft);
    title = `All ${filterLabels[ft]} Tasks (${tasks.length})`;
    showBucket = true;
  }

  if (tasks.length === 0) {
    drilldownDiv.innerHTML = "";
    return;
  }

  const allAssignees = new Set();
  tasks.forEach((t) => t.assignees.forEach((a) => allAssignees.add(a.displayName)));

  if (drillFilterStatus !== "all") tasks = tasks.filter((t) => t.status === drillFilterStatus);
  if (drillFilterAssignee !== "all") {
    tasks = tasks.filter((t) => t.assignees.some((a) => a.displayName === drillFilterAssignee));
  }

  if (drillSortField) {
    const dir = drillSortDir === "asc" ? 1 : -1;
    tasks.sort((a, b) => {
      let va;
      let vb;
      if (drillSortField === "title" || drillSortField === "bucketName") {
        va = a[drillSortField];
        vb = b[drillSortField];
        return dir * va.localeCompare(vb);
      }
      if (drillSortField === "status") {
        const order = { notStarted: 0, inProgress: 1, completed: 2 };
        return dir * (order[a.status] - order[b.status]);
      }
      if (drillSortField === "dueDateTime") {
        va = a.dueDateTime || "9999";
        vb = b.dueDateTime || "9999";
        return dir * va.localeCompare(vb);
      }
      if (drillSortField === "dueStatus") {
        const order = { overdue: 0, dueSoon: 1, onTrack: 2, noDueDate: 3 };
        return dir * (order[a.dueStatus] - order[b.dueStatus]);
      }
      return 0;
    });
  }

  const arrow = (field) => {
    if (drillSortField !== field) return "";
    return `<span class="sort-arrow">${drillSortDir === "asc" ? "▲" : "▼"}</span>`;
  };

  const rows = tasks
    .map((t) => {
      const assignees =
        t.assignees.length > 0
          ? t.assignees.map((a) => escapeHtml(a.displayName)).join(", ")
          : '<span style="color:var(--grey)">Unassigned</span>';
      const dueDate = t.dueDateTime
        ? new Date(t.dueDateTime).toLocaleDateString()
        : '<span style="color:var(--grey)">None</span>';
      const expander =
        t.hasDescription === false
          ? '<span class="expand-placeholder" aria-hidden="true"></span>'
          : `<button type="button" class="expand-icon" data-task-id="${t.id}" aria-expanded="false" aria-label="Toggle task notes"><span class="expand-chevron" aria-hidden="true">▸</span></button>`;

      return `<tr id="row-${t.id}">
      <td>${expander}${escapeHtml(t.title)}</td>
      ${showBucket ? `<td>${escapeHtml(t.bucketName)}</td>` : ""}
      <td><span class="badge badge-${t.status}">${statusLabel[t.status]}</span></td>
      <td>${assignees}</td>
      <td>${dueDate}</td>
      <td><span class="badge badge-${t.dueStatus}">${dueLabel[t.dueStatus]}</span></td>
    </tr>`;
    })
    .join("");

  const assigneeOpts = [...allAssignees]
    .sort()
    .map(
      (a) =>
        `<option value="${escapeHtml(a)}" ${
          drillFilterAssignee === a ? "selected" : ""
        }>${escapeHtml(a)}</option>`
    )
    .join("");

  drilldownDiv.innerHTML = `
    <div class="drilldown-header">
      <h3>${escapeHtml(title)}</h3>
      <button class="drilldown-close" type="button" data-action="clear-drill" aria-label="Close details">&times;</button>
    </div>
    <div class="drilldown-filters">
      <select aria-label="Filter tasks by status" id="drillStatusFilter">
        <option value="all">All Statuses</option>
        <option value="notStarted" ${drillFilterStatus === "notStarted" ? "selected" : ""}>Not Started</option>
        <option value="inProgress" ${drillFilterStatus === "inProgress" ? "selected" : ""}>In Progress</option>
        <option value="completed" ${drillFilterStatus === "completed" ? "selected" : ""}>Completed</option>
      </select>
      <select aria-label="Filter tasks by assignee" id="drillAssigneeFilter">
        <option value="all">All Assignees</option>
        ${assigneeOpts}
      </select>
    </div>
    <table>
      <thead><tr>
        <th data-drill-sort="title">Task${arrow("title")}</th>
        ${showBucket ? `<th data-drill-sort="bucketName">Bucket${arrow("bucketName")}</th>` : ""}
        <th data-drill-sort="status">Status${arrow("status")}</th>
        <th>Assignee(s)</th>
        <th data-drill-sort="dueDateTime">Due Date${arrow("dueDateTime")}</th>
        <th data-drill-sort="dueStatus">Due Status${arrow("dueStatus")}</th>
      </tr></thead>
      <tbody>${rows}</tbody>
    </table>
  `;

  drilldownDiv.scrollIntoView({ behavior: "smooth", block: "nearest" });
}

function sortDrill(field) {
  if (drillSortField === field) {
    drillSortDir = drillSortDir === "asc" ? "desc" : "asc";
  } else {
    drillSortField = field;
    drillSortDir = "asc";
  }
  renderDrillTable();
}

async function toggleNotes(button, taskId) {
  const existingRow = document.getElementById(`notes-${taskId}`);
  if (existingRow) {
    existingRow.remove();
    button.classList.remove("open");
    button.setAttribute("aria-expanded", "false");
    return;
  }

  button.classList.add("open");
  button.setAttribute("aria-expanded", "true");
  const taskRow = document.getElementById(`row-${taskId}`);
  const colSpan = taskRow.cells.length;
  const tempRow = taskRow.insertAdjacentElement(
    "afterend",
    Object.assign(document.createElement("tr"), {
      id: `notes-${taskId}`,
      className: "notes-row",
      innerHTML: `<td colspan="${colSpan}"><em>Loading...</em></td>`,
    })
  );

  try {
    const res = await fetch(`/api/tasks/${encodeURIComponent(taskId)}/details`);
    const data = await res.json();
    const desc = data.description || "No description available.";
    tempRow.innerHTML = `<td colspan="${colSpan}">${escapeHtml(desc)}</td>`;
  } catch {
    tempRow.innerHTML = `<td colspan="${colSpan}" style="color:var(--red)">Failed to load notes</td>`;
  }
}

function renderSummary(stats) {
  const totals = getPlanTotals(stats);
  const pct = totals.total > 0 ? Math.round((totals.completed / totals.total) * 100) : 0;

  summaryDiv.innerHTML = `
    <div class="summary-card"><div class="number" style="color:var(--text)" data-anim-number="${totals.total}">${totals.total}</div><div class="label">Total Tasks</div></div>
    <div class="summary-card"><div class="number" style="color:var(--green)" data-anim-number="${totals.completed}">${totals.completed}</div><div class="label">Completed</div></div>
    <button type="button" class="summary-card clickable" data-action="filter-tasks" data-filter="inProgress"><div class="number" style="color:var(--blue)" data-anim-number="${totals.inProgress}">${totals.inProgress}</div><div class="label">In Progress</div></button>
    <button type="button" class="summary-card clickable" data-action="filter-tasks" data-filter="notStarted"><div class="number" style="color:var(--grey)" data-anim-number="${totals.notStarted}">${totals.notStarted}</div><div class="label">Not Started</div></button>
    <div class="summary-card"><div class="number" style="color:var(--accent)" data-anim-number="${pct}" data-suffix="%">${pct}%</div><div class="label">Complete</div></div>
    <button type="button" class="summary-card clickable" data-action="filter-tasks" data-filter="overdue"><div class="number" style="color:var(--red)" data-anim-number="${totals.overdue}">${totals.overdue}</div><div class="label">Overdue</div></button>
    <button type="button" class="summary-card clickable" data-action="filter-tasks" data-filter="dueSoon"><div class="number" style="color:var(--amber)" data-anim-number="${totals.dueSoon}">${totals.dueSoon}</div><div class="label">Due Soon</div></button>
    <div class="summary-card"><div class="number" style="color:var(--grey)" data-anim-number="${totals.noDueDate}">${totals.noDueDate}</div><div class="label">No Due Date</div></div>
  `;
  applyStaggerReveal(summaryDiv, ".summary-card");
  animateKpiNumbers(summaryDiv);
}

csvBtn.addEventListener("click", () => {
  if (!currentStats || currentStats.length === 0) return;
  const tasks = currentStats.flatMap((b) => b.tasks);
  const headers = ["Task", "Bucket", "Status", "Assignees", "Due Date", "Due Status"];
  const csvRows = [headers.join(",")];
  tasks.forEach((t) => {
    const assignees = t.assignees.map((a) => a.displayName).join("; ");
    const due = t.dueDateTime ? new Date(t.dueDateTime).toLocaleDateString() : "";
    csvRows.push(
      [
        `"${t.title.replace(/\"/g, '""')}"`,
        `"${t.bucketName}"`,
        t.status,
        `"${assignees}"`,
        due,
        t.dueStatus,
      ].join(",")
    );
  });
  const blob = new Blob([csvRows.join("\n")], { type: "text/csv" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = "plan-tasks.csv";
  a.click();
  URL.revokeObjectURL(url);
});

async function exportSnapshotPng() {
  if (!currentStats || currentStats.length === 0 || typeof window.html2canvas !== "function") return;
  const target = document.getElementById("planView");
  const canvas = await window.html2canvas(target, {
    backgroundColor: "#f7faf8",
    scale: 2,
  });
  const link = document.createElement("a");
  link.href = canvas.toDataURL("image/png");
  link.download = `${(currentPlanTitle || "plan").replace(/\s+/g, "-").toLowerCase()}-snapshot.png`;
  link.click();
}

async function exportCurrentViewPdfSnapshot(doc, title) {
  if (typeof window.html2canvas !== "function") return false;
  const target = document.getElementById("planView");
  const canvas = await window.html2canvas(target, {
    backgroundColor: "#f7faf8",
    scale: 2,
  });
  const imgData = canvas.toDataURL("image/png", 1.0);
  const pageWidth = doc.internal.pageSize.getWidth();
  const pageHeight = doc.internal.pageSize.getHeight();
  const margin = 32;
  const printableWidth = pageWidth - margin * 2;
  const scaledHeight = (canvas.height * printableWidth) / canvas.width;

  doc.setFont("helvetica", "bold");
  doc.setFontSize(14);
  doc.setTextColor(31, 41, 55);
  const wrappedTitle = doc.splitTextToSize(title, printableWidth);
  doc.text(wrappedTitle, margin, 34);
  const titleBottom = 34 + (wrappedTitle.length - 1) * 16;
  const firstPageY = titleBottom + 12;
  const printableHeight = pageHeight - firstPageY - 20;
  const renderHeight = Math.min(scaledHeight, printableHeight);

  doc.addImage(imgData, "PNG", margin, firstPageY, printableWidth, renderHeight);

  return true;
}

async function exportPlanPdf() {
  if (!window.jspdf || !window.jspdf.jsPDF) return;
  const { jsPDF } = window.jspdf;
  const doc = new jsPDF({ unit: "pt", format: "a4" });
  const specialView = currentView === "plan" && (!currentPlanId || planView.querySelector(".chart-container").style.display === "none");
  if (specialView) {
    const title = modeTitle.textContent?.trim() || "Martin's Plan-demonium Snapshot";
    const rendered = await exportCurrentViewPdfSnapshot(doc, title);
    if (rendered) {
      doc.save(`${title.replace(/\s+/g, "-").toLowerCase()}-brief.pdf`);
    }
    return;
  }

  if (!currentStats || currentStats.length === 0) return;
  const pageHeight = doc.internal.pageSize.getHeight();
  const pageWidth = doc.internal.pageSize.getWidth();
  const horizontalMargin = 40;
  const contentWidth = pageWidth - horizontalMargin * 2;
  const totals = getPlanTotals(currentStats);
  const pct = totals.total > 0 ? Math.round((totals.completed / totals.total) * 100) : 0;
  const unassignedCount = currentStats
    .flatMap((bucket) => bucket.tasks)
    .filter((task) => task.assignees.length === 0).length;
  const title = currentPlanTitle ? `Project ${currentPlanTitle}` : "Project Plan";

  doc.setFillColor(15, 118, 110);
  doc.rect(0, 0, pageWidth, 24, "F");
  doc.setFont("helvetica", "bold");
  doc.setFontSize(10);
  doc.setTextColor(255, 255, 255);
  doc.text("Project Brief", horizontalMargin, 16);

  doc.setFont("helvetica", "bold");
  doc.setFontSize(18);
  doc.setTextColor(31, 41, 55);
  const wrappedTitle = doc.splitTextToSize(title, contentWidth);
  const titleTopY = 52;
  doc.text(wrappedTitle, horizontalMargin, titleTopY);
  const titleBottomY = titleTopY + (wrappedTitle.length - 1) * 22;

  doc.setFont("helvetica", "normal");
  doc.setFontSize(10);
  doc.setTextColor(90, 90, 90);
  const generatedY = titleBottomY + 18;
  doc.text(`Generated ${new Date().toLocaleString()}`, horizontalMargin, generatedY);

  doc.setFont("helvetica", "bold");
  doc.setFontSize(11);
  doc.setTextColor(15, 118, 110);
  doc.text("Portfolio Snapshot", horizontalMargin, generatedY + 18);

  doc.setFontSize(11);
  doc.setFont("helvetica", "normal");
  doc.setTextColor(31, 41, 55);
  const metricsLine = `Total: ${totals.total}   Completed: ${totals.completed}   In Progress: ${totals.inProgress}   Overdue: ${totals.overdue}   Completion: ${pct}%`;
  const wrappedMetrics = doc.splitTextToSize(metricsLine, contentWidth);
  const metricsTopY = generatedY + 36;
  doc.text(wrappedMetrics, horizontalMargin, metricsTopY);
  const metricsBottomY = metricsTopY + (wrappedMetrics.length - 1) * 14;

  const keyActions = [];
  if (totals.overdue > 0) keyActions.push(`${totals.overdue} overdue task${totals.overdue === 1 ? "" : "s"} require immediate follow-up.`);
  if (unassignedCount > 0) keyActions.push(`${unassignedCount} task${unassignedCount === 1 ? "" : "s"} are unassigned; assign ownership to reduce delivery risk.`);
  if (totals.dueSoon > 0) keyActions.push(`${totals.dueSoon} task${totals.dueSoon === 1 ? "" : "s"} are due soon within the next 7 days.`);
  if (keyActions.length === 0) keyActions.push("No urgent delivery risks detected in current plan snapshot.");

  doc.setFont("helvetica", "bold");
  doc.setFontSize(11);
  doc.setTextColor(15, 118, 110);
  const keyActionsHeaderY = metricsBottomY + 20;
  doc.text("Key Actions", horizontalMargin, keyActionsHeaderY);

  doc.setFont("helvetica", "normal");
  doc.setFontSize(10);
  doc.setTextColor(31, 41, 55);
  const actionLines = keyActions.flatMap((action) =>
    doc.splitTextToSize(`• ${action}`, contentWidth)
  );
  const keyActionsTopY = keyActionsHeaderY + 14;
  doc.text(actionLines, horizontalMargin, keyActionsTopY);
  const keyActionsBottomY = keyActionsTopY + (actionLines.length - 1) * 12;

  const chartCanvas = document.getElementById("chart");
  const chartTopY = keyActionsBottomY + 30;
  if (chartCanvas) {
    const imgData = chartCanvas.toDataURL("image/png", 1.0);
    doc.setFont("helvetica", "bold");
    doc.setFontSize(11);
    doc.setTextColor(15, 118, 110);
    doc.text("Bucket Distribution", horizontalMargin, chartTopY - 6);
    doc.addImage(imgData, "PNG", horizontalMargin, chartTopY, contentWidth, 170);
  }

  const rows = currentStats.flatMap((bucket) =>
    bucket.tasks.map((t) => [
      t.title,
      bucket.bucketName,
      t.status,
      t.assignees.map((a) => a.displayName).join(", ") || "Unassigned",
      t.dueDateTime ? new Date(t.dueDateTime).toLocaleDateString() : "",
      t.dueStatus,
    ])
  );

  if (typeof doc.autoTable === "function") {
    doc.setFont("helvetica", "bold");
    doc.setFontSize(11);
    doc.setTextColor(15, 118, 110);
    const tableStartY = chartTopY + 170 + 24;
    doc.text("Task Register", horizontalMargin, tableStartY - 8);

    doc.autoTable({
      startY: tableStartY,
      head: [["Task", "Bucket", "Status", "Assignees", "Due Date", "Due"]],
      body: rows,
      styles: { fontSize: 8, cellPadding: 4 },
      headStyles: { fillColor: [15, 118, 110] },
      margin: { left: horizontalMargin, right: horizontalMargin },
      theme: "grid",
    });

    const footerY = pageHeight - 18;
    doc.setFont("helvetica", "normal");
    doc.setFontSize(9);
    doc.setTextColor(120, 120, 120);
    doc.text("Generated by Martin's Plan-demonium", horizontalMargin, footerY);
  }

  doc.save(`${title.replace(/\s+/g, "-").toLowerCase()}-brief.pdf`);
}

pdfBtn.addEventListener("click", exportPlanPdf);
pngBtn.addEventListener("click", () => {
  exportSnapshotPng().catch(() => {
    chartHint.textContent = "Snapshot export failed. Try again after chart loads.";
  });
});

workloadBtn.addEventListener("click", loadWorkload);
workloadBackBtn.addEventListener("click", goBack);
backBtn.addEventListener("click", goBack);
overviewBtn.addEventListener("click", goBack);

async function loadWorkload() {
  showView("workload");
  document.getElementById("workloadTable").innerHTML =
    '<div class="loading" style="padding:20px">Loading workload...</div>';

  try {
    const res = await fetch("/api/workload");
    const data = await res.json();
    if (data.error) {
      document.getElementById("workloadTable").innerHTML = `<div class="error">${escapeHtml(
        data.error
      )}</div>`;
      return;
    }

    workloadSnapshot = data;
    renderRiskStrip();

    const ctx = document.getElementById("workloadChart").getContext("2d");
    if (workloadChart) workloadChart.destroy();
    workloadChart = new Chart(ctx, {
      type: "bar",
      data: {
        labels: data.map((d) => d.displayName),
        datasets: [
          {
            label: "Not Started",
            data: data.map((d) => d.notStarted),
            backgroundColor: COLORS.notStarted,
            borderRadius: 4,
          },
          {
            label: "In Progress",
            data: data.map((d) => d.inProgress),
            backgroundColor: COLORS.inProgress,
            borderRadius: 4,
          },
          {
            label: "Completed",
            data: data.map((d) => d.completed),
            backgroundColor: COLORS.completed,
            borderRadius: 4,
          },
        ],
      },
      options: {
        indexAxis: "y",
        responsive: true,
        animation: { duration: 800, easing: "easeOutQuart" },
        plugins: {
          title: {
            display: true,
            text: "Team Workload",
            font: { size: 16, weight: "700", family: "Space Grotesk" },
            color: "#1f2937",
            padding: { bottom: 16 },
          },
          legend: {
            position: "bottom",
            labels: {
              font: { size: 12, family: "Manrope" },
              color: "#6b7280",
              usePointStyle: true,
              pointStyle: "circle",
              padding: 20,
            },
          },
        },
        scales: {
          x: {
            stacked: true,
            beginAtZero: true,
            ticks: { stepSize: 1, font: { size: 12, family: "Manrope" }, color: "#6b7280" },
            grid: { color: "#eee7da" },
          },
          y: {
            stacked: true,
            grid: { display: false },
            ticks: { font: { size: 12, family: "Manrope" }, color: "#6b7280" },
          },
        },
      },
    });

    const rows = data
      .map(
        (d) => `<tr>
      <td><strong>${escapeHtml(d.displayName)}</strong></td>
      <td>${d.total}</td>
      <td style="color:var(--green)">${d.completed}</td>
      <td style="color:var(--blue)">${d.inProgress}</td>
      <td style="color:var(--grey)">${d.notStarted}</td>
      <td style="color:${d.overdue > 0 ? "var(--red)" : "var(--text-secondary)"};font-weight:${
          d.overdue > 0 ? "700" : "400"
        }">${d.overdue}</td>
    </tr>`
      )
      .join("");

    document.getElementById("workloadTable").innerHTML = `<table>
      <thead><tr><th>Assignee</th><th>Total</th><th>Completed</th><th>In Progress</th><th>Not Started</th><th>Overdue</th></tr></thead>
      <tbody>${rows}</tbody>
    </table>`;
  } catch {
    document.getElementById("workloadTable").innerHTML =
      '<div class="error">Failed to load workload</div>';
  }
}

document.body.addEventListener("click", (e) => {
  const overviewRow = e.target.closest(".overview-row[data-plan-id]");
  if (overviewRow && !e.target.closest("[data-action='toggle-pin']")) {
    selectPlan(overviewRow.dataset.planId);
    return;
  }

  const expander = e.target.closest(".expand-icon[data-task-id]");
  if (expander) {
    toggleNotes(expander, expander.dataset.taskId);
    return;
  }

  const actionTarget = e.target.closest("[data-action]");
  if (!actionTarget) return;

  const action = actionTarget.dataset.action;
  if (action === "toggle-pin") {
    togglePin(actionTarget.dataset.planId);
  } else if (action === "show-overdue") {
    showOverdueAll();
  } else if (action === "show-due-soon") {
    showDueSoonAll();
  } else if (action === "show-unassigned") {
    showUnassigned();
  } else if (action === "go-back") {
    goBack();
  } else if (action === "filter-tasks") {
    renderFilteredTasks(actionTarget.dataset.filter);
  } else if (action === "clear-drill") {
    drilldownDiv.innerHTML = "";
    activeDrillData = null;
    selectedBucketIndex = null;
    applyChartSelection();
    syncTriageActiveState();
  }
});

document.body.addEventListener("keydown", (e) => {
  const row = e.target.closest(".overview-row[data-plan-id]");
  if (row && (e.key === "Enter" || e.key === " ")) {
    e.preventDefault();
    selectPlan(row.dataset.planId);
    return;
  }

  const comboItem = e.target.closest(".combo-item");
  if (comboItem && (e.key === "Enter" || e.key === " ")) {
    e.preventDefault();
    comboItem.click();
  }
});

overviewDiv.addEventListener("click", (e) => {
  const sortable = e.target.closest("th[data-sort]");
  if (sortable) sortOverview(sortable.dataset.sort);
});

drilldownDiv.addEventListener("click", (e) => {
  const sortable = e.target.closest("th[data-drill-sort]");
  if (sortable) sortDrill(sortable.dataset.drillSort);
});

drilldownDiv.addEventListener("change", (e) => {
  if (e.target.id === "drillStatusFilter") {
    drillFilterStatus = e.target.value;
    renderDrillTable();
  } else if (e.target.id === "drillAssigneeFilter") {
    drillFilterAssignee = e.target.value;
    renderDrillTable();
  }
});

loadPlans();
loadOverview();
