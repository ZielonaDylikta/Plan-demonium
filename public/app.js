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
const feedbackBtn = document.getElementById("feedbackBtn");
const adminBtn = document.getElementById("adminBtn");
const workloadBackBtn = document.getElementById("workloadBackBtn");
const backBtn = document.getElementById("backBtn");
const csvBtn = document.getElementById("csvBtn");
const pdfBtn = document.getElementById("pdfBtn");
const pngBtn = document.getElementById("pngBtn");
const feedbackModal = document.getElementById("feedbackModal");
const feedbackForm = document.getElementById("feedbackForm");
const feedbackTypeInput = document.getElementById("feedbackType");
const feedbackMessageInput = document.getElementById("feedbackMessage");
const feedbackContactInput = document.getElementById("feedbackContact");
const feedbackStatus = document.getElementById("feedbackStatus");
const feedbackSubmitBtn = document.getElementById("feedbackSubmitBtn");
const feedbackCloseBtn = document.getElementById("feedbackCloseBtn");
const feedbackCancelBtn = document.getElementById("feedbackCancelBtn");
const feedbackExportSection = document.getElementById("feedbackExportSection");
const feedbackExportLockedView = document.getElementById("feedbackExportLockedView");
const feedbackExportUnlockedView = document.getElementById("feedbackExportUnlockedView");
const feedbackExportStatus = document.getElementById("feedbackExportStatus");
const feedbackExportUnlockedStatus = document.getElementById("feedbackExportUnlockedStatus");
const feedbackExportSessionMeta = document.getElementById("feedbackExportSessionMeta");
const feedbackExportUnlockBtn = document.getElementById("feedbackExportUnlockBtn");
const feedbackExportLockBtn = document.getElementById("feedbackExportLockBtn");
const feedbackExportDownloadBtn = document.getElementById("feedbackExportDownloadBtn");
const feedbackExportKeyModal = document.getElementById("feedbackExportKeyModal");
const feedbackExportKeyCloseBtn = document.getElementById("feedbackExportKeyCloseBtn");
const feedbackExportModalKeyInput = document.getElementById("feedbackExportModalKeyInput");
const feedbackExportModalStatus = document.getElementById("feedbackExportModalStatus");
const feedbackExportModalCancelBtn = document.getElementById("feedbackExportModalCancelBtn");
const feedbackExportModalSubmitBtn = document.getElementById("feedbackExportModalSubmitBtn");
const adminModal = document.getElementById("adminModal");
const adminCloseBtn = document.getElementById("adminCloseBtn");
const adminLockedView = document.getElementById("adminLockedView");
const adminUnlockedView = document.getElementById("adminUnlockedView");
const adminPinInput = document.getElementById("adminPinInput");
const adminUnlockBtn = document.getElementById("adminUnlockBtn");
const adminStatus = document.getElementById("adminStatus");
const adminSessionMeta = document.getElementById("adminSessionMeta");
const adminTemplateSelect = document.getElementById("adminTemplateSelect");
const adminTemplateNameInput = document.getElementById("adminTemplateNameInput");
const adminTemplateCreateBtn = document.getElementById("adminTemplateCreateBtn");
const adminTemplateDeleteBtn = document.getElementById("adminTemplateDeleteBtn");
const adminChannelBindingList = document.getElementById("adminChannelBindingList");
const adminChannelStatus = document.getElementById("adminChannelStatus");
const adminIncludeAll = document.getElementById("adminIncludeAll");
const adminPlanSearch = document.getElementById("adminPlanSearch");
const adminPlanList = document.getElementById("adminPlanList");
const adminSaveStatus = document.getElementById("adminSaveStatus");
const adminSaveBtn = document.getElementById("adminSaveBtn");
const adminLogoutBtn = document.getElementById("adminLogoutBtn");

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
const DEV_KEY_STORAGE_KEY = "planDemoniumDevKey";
const devKeyFromQuery = new URLSearchParams(window.location.search).get("devKey");
if (devKeyFromQuery) localStorage.setItem(DEV_KEY_STORAGE_KEY, devKeyFromQuery);

let authToken = null;
let teamsAuthEnabled = false;
let devAccessKey = localStorage.getItem(DEV_KEY_STORAGE_KEY) || "";
let uiConfig = {
  feedbackEnabled: false,
  feedbackExportEnabled: false,
  adminEnabled: false,
  scope: null,
};
let adminPlans = [];
let adminTemplates = [];
let adminActiveTemplateId = null;
let adminChannels = [];
let adminSelection = {
  includeAll: true,
  selectedPlanIds: [],
};
let adminPlanFilterText = "";
let adminSessionExpiresAt = null;
let feedbackExportSessionExpiresAt = null;
let teamsChannelId = "";
const localAuthDebugRequest =
  new URLSearchParams(window.location.search).get("authDebug") === "1" ||
  localStorage.getItem("authDebug") === "1";
let serverAuthDebugEnabled = false;
const authDiagnostics = [];

function shouldShowAuthDiagnostics() {
  return serverAuthDebugEnabled && localAuthDebugRequest;
}

async function loadUiConfig() {
  try {
    const response = await authFetch("/ui-config", { cache: "no-store" });
    if (!response.ok) return;
    const data = await response.json();
    uiConfig = {
      feedbackEnabled: data?.feedbackEnabled === true,
      feedbackExportEnabled: data?.feedbackExportEnabled === true,
      adminEnabled: data?.adminEnabled === true,
      scope: data?.scope || null,
    };
  } catch {
    uiConfig = {
      feedbackEnabled: false,
      feedbackExportEnabled: false,
      adminEnabled: false,
      scope: null,
    };
  }
}

function setFeedbackStatus(message, tone = "") {
  if (!feedbackStatus) return;
  feedbackStatus.textContent = message || "";
  feedbackStatus.classList.remove("success", "error");
  if (tone) feedbackStatus.classList.add(tone);
}

function setFeedbackExportStatus(message, tone = "", target = feedbackExportStatus) {
  if (!target) return;
  target.textContent = message || "";
  target.classList.remove("success", "error");
  if (tone) target.classList.add(tone);
}

function setFeedbackExportModalStatus(message, tone = "") {
  setFeedbackExportStatus(message, tone, feedbackExportModalStatus);
}

function openFeedbackExportKeyModal() {
  if (!feedbackExportKeyModal) return;
  feedbackExportKeyModal.style.display = "flex";
  feedbackExportKeyModal.setAttribute("aria-hidden", "false");
  setFeedbackExportModalStatus("");
  if (feedbackExportModalKeyInput) {
    feedbackExportModalKeyInput.value = "";
    feedbackExportModalKeyInput.focus();
  }
}

function closeFeedbackExportKeyModal() {
  if (!feedbackExportKeyModal) return;
  feedbackExportKeyModal.style.display = "none";
  feedbackExportKeyModal.setAttribute("aria-hidden", "true");
  setFeedbackExportModalStatus("");
  if (feedbackExportModalKeyInput) feedbackExportModalKeyInput.value = "";
}

function setFeedbackExportUnlockedState(isUnlocked) {
  if (!feedbackExportLockedView || !feedbackExportUnlockedView) return;
  feedbackExportLockedView.style.display = isUnlocked ? "none" : "";
  feedbackExportUnlockedView.style.display = isUnlocked ? "" : "none";
  if (feedbackExportUnlockBtn) {
    feedbackExportUnlockBtn.style.display = isUnlocked ? "none" : "";
  }
}

function setFeedbackExportSessionMetaText() {
  if (!feedbackExportSessionMeta) return;
  if (!feedbackExportSessionExpiresAt) {
    feedbackExportSessionMeta.textContent = "";
    return;
  }
  const expires = new Date(feedbackExportSessionExpiresAt);
  feedbackExportSessionMeta.textContent = `Unlocked until ${expires.toLocaleString()}`;
}

async function verifyFeedbackExportSession() {
  const session = await fetchJson("/api/feedback/export/session");
  if (session.error) {
    feedbackExportSessionExpiresAt = null;
    setFeedbackExportUnlockedState(false);
    setFeedbackExportSessionMetaText();
    return {
      ok: false,
      error: session.error,
    };
  }

  feedbackExportSessionExpiresAt = session.expiresAt || null;
  setFeedbackExportUnlockedState(true);
  setFeedbackExportSessionMetaText();
  return { ok: true };
}

async function unlockFeedbackExport() {
  const key = String(feedbackExportModalKeyInput?.value || "").trim();
  if (!key) {
    setFeedbackExportModalStatus("Developer key is required.", "error");
    if (feedbackExportModalKeyInput) feedbackExportModalKeyInput.focus();
    return { ok: false };
  }

  if (feedbackExportModalSubmitBtn) feedbackExportModalSubmitBtn.disabled = true;
  if (feedbackExportUnlockBtn) feedbackExportUnlockBtn.disabled = true;
  setFeedbackExportModalStatus("Unlocking...");
  setFeedbackExportStatus("Unlocking...");

  try {
    const data = await fetchJson("/api/feedback/export/unlock", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ key }),
    });

    if (data.error) {
      setFeedbackExportStatus(data.error, "error");
      setFeedbackExportModalStatus(data.error, "error");
      return { ok: false };
    }

    const sessionState = await verifyFeedbackExportSession();
    if (!sessionState.ok) {
      setFeedbackExportStatus(
        sessionState.error || "Feedback export session could not be established.",
        "error"
      );
      setFeedbackExportModalStatus(
        sessionState.error || "Feedback export session could not be established.",
        "error"
      );
      return { ok: false };
    }
    setFeedbackExportStatus("");
    setFeedbackExportStatus("Export unlocked.", "success", feedbackExportUnlockedStatus);
    setFeedbackExportModalStatus("");
    return { ok: true };
  } catch {
    setFeedbackExportStatus("Failed to unlock export.", "error");
    setFeedbackExportModalStatus("Failed to unlock export.", "error");
    return { ok: false };
  } finally {
    if (feedbackExportModalSubmitBtn) feedbackExportModalSubmitBtn.disabled = false;
    if (feedbackExportUnlockBtn) feedbackExportUnlockBtn.disabled = false;
  }
}

async function submitFeedbackExportUnlock() {
  const result = await unlockFeedbackExport();
  if (result.ok) {
    closeFeedbackExportKeyModal();
  }
}

async function lockFeedbackExport() {
  if (feedbackExportLockBtn) feedbackExportLockBtn.disabled = true;
  try {
    await fetchJson("/api/feedback/export/logout", {
      method: "POST",
    });
  } finally {
    feedbackExportSessionExpiresAt = null;
    setFeedbackExportUnlockedState(false);
    setFeedbackExportSessionMetaText();
    setFeedbackExportStatus("Session locked.", "success");
    setFeedbackExportStatus("", "", feedbackExportUnlockedStatus);
    if (feedbackExportLockBtn) feedbackExportLockBtn.disabled = false;
  }
}

async function downloadFeedbackExportCsv() {
  if (feedbackExportDownloadBtn) feedbackExportDownloadBtn.disabled = true;
  setFeedbackExportStatus("Preparing CSV...", "", feedbackExportUnlockedStatus);

  try {
    const response = await authFetch("/api/feedback/export.csv", {
      cache: "no-store",
    });

    if (!response.ok) {
      let errorText = "Failed to export CSV.";
      try {
        const data = await response.json();
        if (data?.error) errorText = data.error;
      } catch {
        // Keep generic error.
      }

      if (response.status === 403) {
        feedbackExportSessionExpiresAt = null;
        setFeedbackExportUnlockedState(false);
        setFeedbackExportSessionMetaText();
      }
      setFeedbackExportStatus(errorText, "error", feedbackExportUnlockedStatus);
      return;
    }

    const blob = await response.blob();
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = `feedback-${new Date().toISOString().slice(0, 10)}.csv`;
    document.body.appendChild(link);
    link.click();
    link.remove();
    URL.revokeObjectURL(url);
    setFeedbackExportStatus("CSV downloaded.", "success", feedbackExportUnlockedStatus);
  } catch {
    setFeedbackExportStatus("Failed to export CSV.", "error", feedbackExportUnlockedStatus);
  } finally {
    if (feedbackExportDownloadBtn) feedbackExportDownloadBtn.disabled = false;
  }
}

async function refreshFeedbackExportView() {
  if (!feedbackExportSection) return;

  const enabled = uiConfig.feedbackEnabled && uiConfig.feedbackExportEnabled;
  feedbackExportSection.style.display = enabled ? "" : "none";
  if (!enabled) {
    feedbackExportSessionExpiresAt = null;
    setFeedbackExportUnlockedState(false);
    if (feedbackExportUnlockBtn) feedbackExportUnlockBtn.style.display = "none";
    return;
  }

  setFeedbackExportStatus("");
  setFeedbackExportStatus("", "", feedbackExportUnlockedStatus);
  const sessionState = await verifyFeedbackExportSession();
  if (!sessionState.ok) return;
}

function openFeedbackModal() {
  if (!feedbackModal) return;
  setFeedbackStatus("");
  setFeedbackExportStatus("");
  setFeedbackExportStatus("", "", feedbackExportUnlockedStatus);
  feedbackModal.style.display = "flex";
  feedbackModal.setAttribute("aria-hidden", "false");
  if (feedbackTypeInput) feedbackTypeInput.value = "general";
  if (feedbackMessageInput) feedbackMessageInput.focus();
  refreshFeedbackExportView().catch(() => {
    setFeedbackExportStatus("Failed to load export access state.", "error");
  });
}

function closeFeedbackModal() {
  if (!feedbackModal) return;
  feedbackModal.style.display = "none";
  feedbackModal.setAttribute("aria-hidden", "true");
  if (feedbackForm) feedbackForm.reset();
  setFeedbackStatus("");
  setFeedbackExportStatus("");
  setFeedbackExportStatus("", "", feedbackExportUnlockedStatus);
}

function setAdminStatus(message, tone = "", target = adminStatus) {
  if (!target) return;
  target.textContent = message || "";
  target.classList.remove("success", "error");
  if (tone) target.classList.add(tone);
}

function setAdminUnlockedState(isUnlocked) {
  if (!adminLockedView || !adminUnlockedView) return;
  adminLockedView.style.display = isUnlocked ? "none" : "";
  adminUnlockedView.style.display = isUnlocked ? "" : "none";
}

function renderAdminPlanList() {
  if (!adminPlanList) return;

  if (!adminPlans.length) {
    adminPlanList.innerHTML = '<div class="guided-state" style="margin:6px 0">No plans available</div>';
    return;
  }

  const selected = new Set(adminSelection.selectedPlanIds);
  const includeAll = Boolean(adminSelection.includeAll);
  const q = adminPlanFilterText.trim().toLowerCase();
  const visiblePlans = q
    ? adminPlans.filter((plan) => String(plan.title || "").toLowerCase().includes(q))
    : adminPlans;

  if (!visiblePlans.length) {
    adminPlanList.innerHTML = '<div class="guided-state" style="margin:6px 0">No plans match your search</div>';
    return;
  }

  adminPlanList.innerHTML = visiblePlans
    .map((plan) => {
      const checked = includeAll || selected.has(plan.id);
      return `<label class="admin-plan-item">
        <input type="checkbox" class="admin-plan-checkbox" data-plan-id="${escapeHtml(plan.id)}" ${
          checked ? "checked" : ""
        } ${includeAll ? "disabled" : ""} />
        <span>${escapeHtml(plan.title || plan.id)}</span>
      </label>`;
    })
    .join("");
}

function setAdminSessionMetaText() {
  if (!adminSessionMeta) return;
  if (!adminSessionExpiresAt) {
    adminSessionMeta.textContent = "";
    return;
  }
  const expires = new Date(adminSessionExpiresAt);
  adminSessionMeta.textContent = `Unlocked until ${expires.toLocaleString()}`;
}

function renderAdminTemplateControls() {
  if (!adminTemplateSelect) return;

  adminTemplateSelect.innerHTML = adminTemplates
    .map((template) => {
      const selected = template.id === adminActiveTemplateId ? " selected" : "";
      return `<option value="${escapeHtml(template.id)}"${selected}>${escapeHtml(
        template.name || template.id
      )}</option>`;
    })
    .join("");

  if (adminTemplateDeleteBtn) {
    adminTemplateDeleteBtn.disabled = adminTemplates.length <= 1;
  }
}

function getAdminTemplateName(templateId) {
  return adminTemplates.find((template) => template.id === templateId)?.name || templateId;
}

function setAdminChannelStatus(message, tone = "") {
  setAdminStatus(message, tone, adminChannelStatus);
}

function renderAdminChannelBindings() {
  if (!adminChannelBindingList) return;

  if (!adminChannels.length) {
    adminChannelBindingList.innerHTML =
      '<div class="guided-state" style="margin:6px 0">No channels discovered for this Team</div>';
    return;
  }

  const activeTemplateName = getAdminTemplateName(adminActiveTemplateId) || "active template";
  const templateOptions = [
    `<option value="">Use active template (${escapeHtml(activeTemplateName)})</option>`,
    ...adminTemplates.map(
      (template) =>
        `<option value="${escapeHtml(template.id)}">${escapeHtml(template.name || template.id)}</option>`
    ),
  ].join("");

  adminChannelBindingList.innerHTML = adminChannels
    .map((channel) => {
      const selectedValue = String(channel.templateId || "");
      return `<label class="admin-plan-item" style="display:block">
        <span>
          <strong>${escapeHtml(channel.displayName || channel.id)}</strong>
          <small style="color:var(--text-secondary);display:block">${escapeHtml(
            channel.membershipType || "standard"
          )}</small>
        </span>
        <select class="admin-channel-template-select" data-channel-id="${escapeHtml(channel.id)}">${templateOptions}</select>
      </label>`;
    })
    .join("");

  adminChannelBindingList
    .querySelectorAll(".admin-channel-template-select[data-channel-id]")
    .forEach((selectNode) => {
      const channelId = selectNode.dataset.channelId;
      const selected = String(
        adminChannels.find((channel) => channel.id === channelId)?.templateId || ""
      );
      selectNode.value = selected;
    });
}

async function loadAdminScopeData() {
  const [plansData, selectionData, templatesData, channelsData] = await Promise.all([
    fetchJson("/api/admin/plans"),
    fetchJson("/api/admin/selection"),
    fetchJson("/api/admin/templates"),
    fetchJson("/api/admin/channels"),
  ]);

  if (plansData.error) throw new Error(plansData.error);
  if (selectionData.error) throw new Error(selectionData.error);
  if (templatesData.error) throw new Error(templatesData.error);

  adminPlans = Array.isArray(plansData) ? plansData : [];
  adminTemplates = Array.isArray(templatesData.templates) ? templatesData.templates : [];
  adminActiveTemplateId = String(
    templatesData.activeTemplateId || selectionData.templateId || ""
  );
  adminChannels = Array.isArray(channelsData.channels)
    ? channelsData.channels.map((channel) => ({
        id: String(channel.id || ""),
        displayName: String(channel.displayName || channel.id || "").trim(),
        membershipType: String(channel.membershipType || "standard"),
        templateId: channel.templateId ? String(channel.templateId) : null,
      }))
    : [];
  adminSelection = {
    includeAll: Boolean(selectionData.includeAll),
    selectedPlanIds: Array.isArray(selectionData.selectedPlanIds)
      ? selectionData.selectedPlanIds.map((id) => String(id))
      : [],
  };
  adminPlanFilterText = "";

  renderAdminTemplateControls();
  renderAdminChannelBindings();
  if (adminIncludeAll) adminIncludeAll.checked = adminSelection.includeAll;
  if (adminPlanSearch) adminPlanSearch.value = "";
  if (channelsData.error) {
    setAdminChannelStatus(channelsData.error, "error");
  } else {
    setAdminChannelStatus("");
  }
  renderAdminPlanList();
}

async function saveAdminChannelBinding(channelId, templateId) {
  const normalizedChannelId = String(channelId || "").trim();
  if (!normalizedChannelId) return;

  setAdminChannelStatus("Saving channel binding...");
  try {
    const data = await fetchJson(
      `/api/admin/channel-bindings/${encodeURIComponent(normalizedChannelId)}`,
      {
        method: "PUT",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ templateId: templateId || "" }),
      }
    );
    if (data.error) {
      setAdminChannelStatus(data.error, "error");
      await loadAdminScopeData();
      return;
    }

    adminChannels = adminChannels.map((channel) =>
      channel.id === normalizedChannelId
        ? {
            ...channel,
            templateId: templateId || null,
          }
        : channel
    );
    renderAdminChannelBindings();
    setAdminChannelStatus("Channel binding saved.", "success");
    await refreshDashboardDataForScope();
  } catch {
    setAdminChannelStatus("Failed to save channel binding.", "error");
  }
}

async function activateAdminTemplate(templateId) {
  if (!templateId) return;
  if (adminTemplateSelect) adminTemplateSelect.disabled = true;
  setAdminStatus("Switching template...", "", adminSaveStatus);

  try {
    const data = await fetchJson(`/api/admin/templates/${encodeURIComponent(templateId)}/activate`, {
      method: "POST",
    });
    if (data.error) {
      setAdminStatus(data.error, "error", adminSaveStatus);
      return;
    }

    await loadAdminScopeData();
    setAdminStatus("Template activated. Refreshing dashboard...", "success", adminSaveStatus);
    await refreshDashboardDataForScope();
  } catch {
    setAdminStatus("Failed to activate template.", "error", adminSaveStatus);
  } finally {
    if (adminTemplateSelect) adminTemplateSelect.disabled = false;
  }
}

async function createAdminTemplateFromCurrent() {
  const name = String(adminTemplateNameInput?.value || "").trim();
  if (!name) {
    setAdminStatus("Enter a template name.", "error", adminSaveStatus);
    if (adminTemplateNameInput) adminTemplateNameInput.focus();
    return;
  }

  if (adminTemplateCreateBtn) adminTemplateCreateBtn.disabled = true;
  setAdminStatus("Creating template...", "", adminSaveStatus);

  try {
    const includeAll = Boolean(adminIncludeAll?.checked);
    const selectedPlanIds = includeAll ? [] : [...new Set(adminSelection.selectedPlanIds)];
    const data = await fetchJson("/api/admin/templates", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ name, includeAll, selectedPlanIds }),
    });

    if (data.error) {
      setAdminStatus(data.error, "error", adminSaveStatus);
      return;
    }

    if (adminTemplateNameInput) adminTemplateNameInput.value = "";
    await loadAdminScopeData();
    setAdminStatus("Template created and activated.", "success", adminSaveStatus);
    await refreshDashboardDataForScope();
  } catch {
    setAdminStatus("Failed to create template.", "error", adminSaveStatus);
  } finally {
    if (adminTemplateCreateBtn) adminTemplateCreateBtn.disabled = false;
  }
}

async function deleteActiveAdminTemplate() {
  const templateId = String(adminTemplateSelect?.value || "");
  if (!templateId) return;
  if (!window.confirm("Delete this scope template?")) return;

  if (adminTemplateDeleteBtn) adminTemplateDeleteBtn.disabled = true;
  setAdminStatus("Deleting template...", "", adminSaveStatus);

  try {
    const data = await fetchJson(`/api/admin/templates/${encodeURIComponent(templateId)}`, {
      method: "DELETE",
    });
    if (data.error) {
      setAdminStatus(data.error, "error", adminSaveStatus);
      return;
    }

    await loadAdminScopeData();
    setAdminStatus("Template deleted. Refreshing dashboard...", "success", adminSaveStatus);
    await refreshDashboardDataForScope();
  } catch {
    setAdminStatus("Failed to delete template.", "error", adminSaveStatus);
  } finally {
    if (adminTemplateDeleteBtn) adminTemplateDeleteBtn.disabled = adminTemplates.length <= 1;
  }
}

async function refreshDashboardDataForScope() {
  await loadPlans();
  await loadOverview();
  if (currentView === "workload") {
    await loadWorkload();
  }
}

async function verifyAdminSession() {
  const session = await fetchJson("/api/admin/session");
  if (session.error) {
    adminSessionExpiresAt = null;
    setAdminUnlockedState(false);
    return {
      ok: false,
      error: session.error,
    };
  }

  adminSessionExpiresAt = session.expiresAt || null;
  setAdminUnlockedState(true);
  setAdminSessionMetaText();
  try {
    await loadAdminScopeData();
    return { ok: true };
  } catch (err) {
    setAdminStatus(err?.message || "Admin scope failed to load.", "error", adminSaveStatus);
    return { ok: true };
  }
}

async function openAdminModal() {
  if (!adminModal) return;
  adminModal.style.display = "flex";
  adminModal.setAttribute("aria-hidden", "false");
  setAdminStatus("");
  setAdminStatus("", "", adminSaveStatus);

  try {
    const sessionState = await verifyAdminSession();
    if (!sessionState.ok && adminPinInput) {
      adminPinInput.value = "";
      adminPinInput.focus();
      if (sessionState.error) setAdminStatus(sessionState.error, "error");
    }
  } catch (err) {
    setAdminUnlockedState(false);
    setAdminStatus(err?.message || "Failed to open admin controls.", "error");
  }
}

function closeAdminModal() {
  if (!adminModal) return;
  adminModal.style.display = "none";
  adminModal.setAttribute("aria-hidden", "true");
  setAdminStatus("");
  setAdminStatus("", "", adminSaveStatus);
}

async function unlockAdmin() {
  if (!adminPinInput) return;
  const pin = adminPinInput.value.trim();
  if (!/^\d{6}$/.test(pin)) {
    setAdminStatus("PIN must be 6 digits.", "error");
    adminPinInput.focus();
    return;
  }

  if (adminUnlockBtn) adminUnlockBtn.disabled = true;
  setAdminStatus("Unlocking...");

  try {
    const data = await fetchJson("/api/admin/unlock", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ pin }),
    });

    if (data.error) {
      setAdminStatus(data.error, "error");
      return;
    }

    adminPinInput.value = "";
    const sessionState = await verifyAdminSession();
    if (!sessionState.ok) {
      setAdminStatus(sessionState.error || "Admin session could not be established.", "error");
      return;
    }
    setAdminStatus("", "");
    setAdminStatus("Admin unlocked.", "success", adminSaveStatus);
  } catch {
    setAdminStatus("Failed to unlock admin session.", "error");
  } finally {
    if (adminUnlockBtn) adminUnlockBtn.disabled = false;
  }
}

async function saveAdminSelection() {
  const includeAll = Boolean(adminIncludeAll?.checked);
  const selectedPlanIds = includeAll ? [] : [...new Set(adminSelection.selectedPlanIds)];

  if (!includeAll && selectedPlanIds.length === 0) {
    setAdminStatus("Select at least one plan, or enable View all plans.", "error", adminSaveStatus);
    return;
  }

  if (adminSaveBtn) adminSaveBtn.disabled = true;
  setAdminStatus("Saving scope...", "", adminSaveStatus);

  try {
    const data = await fetchJson("/api/admin/selection", {
      method: "PUT",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        includeAll,
        selectedPlanIds,
      }),
    });

    if (data.error) {
      setAdminStatus(data.error, "error", adminSaveStatus);
      return;
    }

    adminSelection = {
      includeAll: Boolean(data.includeAll),
      selectedPlanIds: Array.isArray(data.selectedPlanIds) ? data.selectedPlanIds : [],
    };
    if (adminIncludeAll) adminIncludeAll.checked = adminSelection.includeAll;
    renderAdminPlanList();
    setAdminStatus("Scope saved. Refreshing dashboard...", "success", adminSaveStatus);
    await refreshDashboardDataForScope();
  } catch {
    setAdminStatus("Failed to save scope.", "error", adminSaveStatus);
  } finally {
    if (adminSaveBtn) adminSaveBtn.disabled = false;
  }
}

async function logoutAdmin() {
  if (adminLogoutBtn) adminLogoutBtn.disabled = true;
  try {
    await fetchJson("/api/admin/logout", {
      method: "POST",
    });
  } finally {
    adminSessionExpiresAt = null;
    setAdminUnlockedState(false);
    setAdminSessionMetaText();
    setAdminStatus("Session locked.", "success");
    if (adminLogoutBtn) adminLogoutBtn.disabled = false;
  }
}

async function handleFeedbackSubmit(e) {
  e.preventDefault();
  if (!feedbackMessageInput || !feedbackTypeInput || !feedbackContactInput) return;

  const message = feedbackMessageInput.value.trim();
  if (!message) {
    setFeedbackStatus("Please enter feedback before sending.", "error");
    feedbackMessageInput.focus();
    return;
  }

  if (feedbackSubmitBtn) feedbackSubmitBtn.disabled = true;
  setFeedbackStatus("Sending...");

  try {
    const data = await fetchJson("/api/feedback", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        type: feedbackTypeInput.value === "bug" ? "bug" : "general",
        message,
        contact: feedbackContactInput.value.trim(),
      }),
    });

    if (data.error) {
      setFeedbackStatus(data.error, "error");
      return;
    }

    setFeedbackStatus("Thanks! Your feedback was submitted.", "success");
    setTimeout(() => {
      closeFeedbackModal();
    }, 650);
  } catch {
    setFeedbackStatus("Failed to submit feedback. Please try again.", "error");
  } finally {
    if (feedbackSubmitBtn) feedbackSubmitBtn.disabled = false;
  }
}

async function loadAuthConfig() {
  try {
    const response = await fetch("/auth-config", { cache: "no-store" });
    if (!response.ok) {
      serverAuthDebugEnabled = false;
      return;
    }
    const data = await response.json();
    serverAuthDebugEnabled = data?.authDebug === true;
  } catch {
    serverAuthDebugEnabled = false;
  }
}

function recordAuthDiagnostic(event, details = {}) {
  const timestamp = new Date().toISOString();
  authDiagnostics.push({ timestamp, event, ...details });
  if (authDiagnostics.length > 20) authDiagnostics.shift();
  if (shouldShowAuthDiagnostics()) {
    console.log("[auth-ui]", event, details);
  }
}

function buildAuthDebugHtml() {
  if (!shouldShowAuthDiagnostics()) return "";
  if (!authDiagnostics.length) return "";
  const rows = authDiagnostics
    .slice()
    .reverse()
    .map((entry) => {
      const { timestamp, event, ...rest } = entry;
      const detail = Object.entries(rest)
        .map(([k, v]) => `${k}=${v}`)
        .join(" | ");
      return `<li><strong>${escapeHtml(event)}</strong> <small>${escapeHtml(timestamp)}</small><br><span>${escapeHtml(detail || "-")}</span></li>`;
    })
    .join("");
  return `<details open style="margin-top:12px"><summary>Auth diagnostics</summary><ol style="margin-top:8px;padding-left:20px">${rows}</ol></details>`;
}

function renderAuthError(message) {
  const html = `<div class="error">${escapeHtml(message)}${buildAuthDebugHtml()}</div>`;
  overviewDiv.innerHTML = html;
  drilldownDiv.innerHTML = "";
  summaryDiv.innerHTML = "";
  chartHint.textContent = "";
}

async function initAuthSession() {
  recordAuthDiagnostic("auth_init_start", {
    teamsSdkPresent: Boolean(window.microsoftTeams?.app && window.microsoftTeams?.authentication),
    hasDevKey: Boolean(devAccessKey),
  });

  if (window.microsoftTeams?.app && window.microsoftTeams?.authentication) {
    try {
      await window.microsoftTeams.app.initialize();
      recordAuthDiagnostic("teams_init_ok");
      try {
        const teamsContext = await window.microsoftTeams.app.getContext();
        teamsChannelId =
          String(
            teamsContext?.channel?.id ||
              teamsContext?.channelId ||
              teamsContext?.chat?.id ||
              ""
          ).trim();
        recordAuthDiagnostic("teams_context_ok", {
          hasChannelId: Boolean(teamsChannelId),
        });
      } catch {
        teamsChannelId = "";
        recordAuthDiagnostic("teams_context_failed");
      }
      authToken = await window.microsoftTeams.authentication.getAuthToken();
      teamsAuthEnabled = true;
      recordAuthDiagnostic("teams_token_ok", {
        tokenLength: String(authToken || "").length,
      });
      return true;
    } catch (err) {
      const message = err?.message || "unknown_error";
      const code = err?.errorCode || err?.code || "none";
      recordAuthDiagnostic("teams_token_failed", { code, message });
      console.warn("[auth] Teams SSO initialization failed:", err?.message || err);
    }
  }

  if (devAccessKey) {
    recordAuthDiagnostic("dev_key_mode_enabled");
    return true;
  }

  renderAuthError(
    "Authentication required. Open this dashboard as a Teams tab."
  );
  return false;
}

async function authFetch(input, init = {}, allowRetry = true) {
  const headers = new Headers(init.headers || {});
  const requestPath = typeof input === "string" ? input : String(input?.url || "");
  if (teamsChannelId) {
    headers.set("x-teams-channel-id", teamsChannelId);
  }
  if (authToken) {
    headers.set("Authorization", `Bearer ${authToken}`);
    recordAuthDiagnostic("api_request", { path: requestPath, mode: "bearer" });
  } else if (devAccessKey) {
    headers.set("x-dev-access-key", devAccessKey);
    recordAuthDiagnostic("api_request", { path: requestPath, mode: "dev-key" });
  } else {
    recordAuthDiagnostic("api_request", { path: requestPath, mode: "none" });
  }

  const response = await fetch(input, {
    ...init,
    headers,
  });

  if (response.status === 401 && allowRetry && teamsAuthEnabled) {
    recordAuthDiagnostic("api_401_retry", { path: requestPath });
    try {
      authToken = await window.microsoftTeams.authentication.getAuthToken();
      recordAuthDiagnostic("teams_token_refresh_ok", {
        tokenLength: String(authToken || "").length,
      });
      return authFetch(input, init, false);
    } catch {
      recordAuthDiagnostic("teams_token_refresh_failed", { path: requestPath });
      // Let caller handle final 401 response.
    }
  }

  if (response.status === 401) {
    recordAuthDiagnostic("api_401_final", {
      path: requestPath,
      traceId: response.headers.get("x-auth-trace-id") || "none",
    });
  }

  return response;
}

async function fetchJson(input, init = {}) {
  const response = await authFetch(input, init);
  let data = {};
  try {
    data = await response.json();
  } catch {
    data = {};
  }

  if (!response.ok && !data.error) {
    data.error = `Request failed (${response.status})`;
  }

  return data;
}

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
    const plans = await fetchJson("/api/plans");
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
    const data = await fetchJson("/api/unassigned");
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
    const data = await fetchJson("/api/overview");
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
    const data = await fetchJson("/api/deltas");
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
    const data = await fetchJson("/api/workload");
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
    const data = await fetchJson("/api/trends");
    if (data.error) return;
    trendData = data;
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
    const stats = await fetchJson(`/api/plans/${encodeURIComponent(planId)}/stats`);
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
    const data = await fetchJson(`/api/tasks/${encodeURIComponent(taskId)}/details`);
    if (data.error) {
      tempRow.innerHTML = `<td colspan="${colSpan}" style="color:var(--red)">${escapeHtml(data.error)}</td>`;
      return;
    }
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
if (feedbackBtn) {
  feedbackBtn.addEventListener("click", openFeedbackModal);
}
if (feedbackCloseBtn) {
  feedbackCloseBtn.addEventListener("click", closeFeedbackModal);
}
if (feedbackCancelBtn) {
  feedbackCancelBtn.addEventListener("click", closeFeedbackModal);
}
if (feedbackModal) {
  feedbackModal.addEventListener("click", (e) => {
    if (e.target === feedbackModal) closeFeedbackModal();
  });
}
if (feedbackForm) {
  feedbackForm.addEventListener("submit", handleFeedbackSubmit);
}
if (feedbackExportUnlockBtn) {
  feedbackExportUnlockBtn.addEventListener("click", openFeedbackExportKeyModal);
}
if (feedbackExportLockBtn) {
  feedbackExportLockBtn.addEventListener("click", lockFeedbackExport);
}
if (feedbackExportDownloadBtn) {
  feedbackExportDownloadBtn.addEventListener("click", downloadFeedbackExportCsv);
}
if (feedbackExportKeyCloseBtn) {
  feedbackExportKeyCloseBtn.addEventListener("click", closeFeedbackExportKeyModal);
}
if (feedbackExportModalCancelBtn) {
  feedbackExportModalCancelBtn.addEventListener("click", closeFeedbackExportKeyModal);
}
if (feedbackExportModalSubmitBtn) {
  feedbackExportModalSubmitBtn.addEventListener("click", submitFeedbackExportUnlock);
}
if (feedbackExportModalKeyInput) {
  feedbackExportModalKeyInput.addEventListener("keydown", (e) => {
    if (e.key === "Enter") {
      e.preventDefault();
      submitFeedbackExportUnlock();
    }
  });
}
if (feedbackExportKeyModal) {
  feedbackExportKeyModal.addEventListener("click", (e) => {
    if (e.target === feedbackExportKeyModal) closeFeedbackExportKeyModal();
  });
}
if (adminBtn) {
  adminBtn.addEventListener("click", openAdminModal);
}
if (adminCloseBtn) {
  adminCloseBtn.addEventListener("click", closeAdminModal);
}
if (adminModal) {
  adminModal.addEventListener("click", (e) => {
    if (e.target === adminModal) closeAdminModal();
  });
}
if (adminUnlockBtn) {
  adminUnlockBtn.addEventListener("click", unlockAdmin);
}
if (adminPinInput) {
  adminPinInput.addEventListener("keydown", (e) => {
    if (e.key === "Enter") {
      e.preventDefault();
      unlockAdmin();
    }
  });
}
if (adminIncludeAll) {
  adminIncludeAll.addEventListener("change", () => {
    adminSelection.includeAll = adminIncludeAll.checked;
    renderAdminPlanList();
    setAdminStatus("", "", adminSaveStatus);
  });
}
if (adminTemplateSelect) {
  adminTemplateSelect.addEventListener("change", () => {
    const templateId = String(adminTemplateSelect.value || "").trim();
    activateAdminTemplate(templateId);
  });
}
if (adminTemplateCreateBtn) {
  adminTemplateCreateBtn.addEventListener("click", createAdminTemplateFromCurrent);
}
if (adminTemplateDeleteBtn) {
  adminTemplateDeleteBtn.addEventListener("click", deleteActiveAdminTemplate);
}
if (adminTemplateNameInput) {
  adminTemplateNameInput.addEventListener("keydown", (e) => {
    if (e.key === "Enter") {
      e.preventDefault();
      createAdminTemplateFromCurrent();
    }
  });
}
if (adminChannelBindingList) {
  adminChannelBindingList.addEventListener("change", (e) => {
    const selectNode = e.target.closest(".admin-channel-template-select[data-channel-id]");
    if (!selectNode) return;
    const channelId = String(selectNode.dataset.channelId || "");
    const templateId = String(selectNode.value || "");
    saveAdminChannelBinding(channelId, templateId);
  });
}
if (adminPlanList) {
  adminPlanList.addEventListener("change", (e) => {
    const checkbox = e.target.closest(".admin-plan-checkbox[data-plan-id]");
    if (!checkbox) return;
    const selected = new Set(adminSelection.selectedPlanIds);
    if (checkbox.checked) selected.add(checkbox.dataset.planId);
    else selected.delete(checkbox.dataset.planId);
    adminSelection.selectedPlanIds = [...selected];
    setAdminStatus("", "", adminSaveStatus);
  });
}
if (adminPlanSearch) {
  adminPlanSearch.addEventListener("input", () => {
    adminPlanFilterText = adminPlanSearch.value || "";
    renderAdminPlanList();
  });
}
if (adminSaveBtn) {
  adminSaveBtn.addEventListener("click", saveAdminSelection);
}
if (adminLogoutBtn) {
  adminLogoutBtn.addEventListener("click", logoutAdmin);
}

async function loadWorkload() {
  showView("workload");
  document.getElementById("workloadTable").innerHTML =
    '<div class="loading" style="padding:20px">Loading workload...</div>';

  try {
    const data = await fetchJson("/api/workload");
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
  if (e.key === "Escape" && feedbackExportKeyModal?.getAttribute("aria-hidden") === "false") {
    closeFeedbackExportKeyModal();
    return;
  }

  if (e.key === "Escape" && adminModal?.getAttribute("aria-hidden") === "false") {
    closeAdminModal();
    return;
  }

  if (e.key === "Escape" && feedbackModal?.getAttribute("aria-hidden") === "false") {
    closeFeedbackModal();
    return;
  }

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

async function bootstrapApp() {
  await loadAuthConfig();
  const isAuthorized = await initAuthSession();
  if (!isAuthorized) return;

  await loadUiConfig();
  if (feedbackBtn) {
    feedbackBtn.style.display = uiConfig.feedbackEnabled ? "" : "none";
  }
  if (adminBtn) {
    adminBtn.style.display = uiConfig.adminEnabled ? "" : "none";
  }

  await loadPlans();
  await loadOverview();
}

bootstrapApp();
