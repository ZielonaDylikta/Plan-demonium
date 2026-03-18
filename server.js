require("dotenv").config();
const express = require("express");
const fs = require("fs");
const path = require("path");
const msal = require("@azure/msal-node");
const { Client } = require("@microsoft/microsoft-graph-client");

const app = express();
app.use(express.static("public"));

const API_CACHE_TTL_MS = 60 * 1000;
const apiCache = new Map();

function getCachedValue(key) {
  const cached = apiCache.get(key);
  if (!cached) return null;
  if (Date.now() - cached.createdAt > API_CACHE_TTL_MS) {
    apiCache.delete(key);
    return null;
  }
  return cached.value;
}

function setCachedValue(key, value) {
  apiCache.set(key, {
    value,
    createdAt: Date.now(),
  });
}

async function withCache(cacheKey, producer) {
  const cached = getCachedValue(cacheKey);
  if (cached) return cached;
  const value = await producer();
  setCachedValue(cacheKey, value);
  return value;
}

// MSAL client credentials config
const msalConfig = {
  auth: {
    clientId: process.env.CLIENT_ID,
    authority: `https://login.microsoftonline.com/${process.env.TENANT_ID}`,
    clientSecret: process.env.CLIENT_SECRET,
  },
};

const cca = new msal.ConfidentialClientApplication(msalConfig);

async function getGraphClient() {
  const result = await cca.acquireTokenByClientCredential({
    scopes: ["https://graph.microsoft.com/.default"],
  });

  return Client.init({
    authProvider: (done) => done(null, result.accessToken),
  });
}

// Fetch all pages from a paginated Graph API endpoint
async function fetchAllPages(client, url) {
  let results = [];
  let nextUrl = url;
  while (nextUrl) {
    const res = await client.api(nextUrl).get();
    results.push(...res.value);
    nextUrl = res["@odata.nextLink"] || null;
  }
  return results;
}

// Classify a due date relative to today
function getDueStatus(dueDateTime, percentComplete) {
  if (!dueDateTime) return "noDueDate";
  if (percentComplete === 100) return "onTrack";
  const now = new Date();
  now.setHours(0, 0, 0, 0);
  const due = new Date(dueDateTime);
  if (due < now) return "overdue";
  const weekFromNow = new Date(now);
  weekFromNow.setDate(weekFromNow.getDate() + 7);
  if (due <= weekFromNow) return "dueSoon";
  return "onTrack";
}

// Resolve user display names from IDs (graceful fallback)
async function resolveUsers(client, userIds) {
  const userMap = new Map();
  const results = await Promise.allSettled(
    userIds.map((id) =>
      client.api(`/users/${id}`).select("displayName").get()
    )
  );
  userIds.forEach((id, i) => {
    if (results[i].status === "fulfilled") {
      userMap.set(id, results[i].value.displayName);
    } else {
      userMap.set(id, id);
    }
  });
  return userMap;
}

// --- Snapshot helpers for sparkline trends ---
const SNAPSHOT_DIR = path.join(__dirname, "data");
const SNAPSHOT_FILE = path.join(SNAPSHOT_DIR, "snapshots.json");
const SNAPSHOT_INTERVAL_MS = 24 * 60 * 60 * 1000;
const SNAPSHOT_TIME_UTC = process.env.SNAPSHOT_TIME_UTC || "";
let snapshotTimer = null;

function parseSnapshotTimeUtc(value) {
  if (!value) return null;
  const match = String(value).trim().match(/^([01]\d|2[0-3]):([0-5]\d)$/);
  if (!match) return null;
  return {
    hours: Number.parseInt(match[1], 10),
    minutes: Number.parseInt(match[2], 10),
  };
}

function getSnapshotCapturedAtMs(dateKey, dayData) {
  const fromMeta = dayData?.__meta?.capturedAt;
  if (typeof fromMeta === "string") {
    const parsed = Date.parse(fromMeta);
    if (!Number.isNaN(parsed)) return parsed;
  }

  const legacy = Date.parse(`${dateKey}T00:00:00.000Z`);
  return Number.isNaN(legacy) ? null : legacy;
}

function getSnapshotPlanEntries(dayData) {
  if (!dayData || typeof dayData !== "object") return [];
  return Object.entries(dayData).filter(
    ([key, stats]) => !key.startsWith("__") && stats && typeof stats === "object"
  );
}

function getLatestSnapshotInfo(snapshots) {
  let latest = null;

  Object.entries(snapshots || {}).forEach(([dateKey, dayData]) => {
    const capturedAtMs = getSnapshotCapturedAtMs(dateKey, dayData);
    if (capturedAtMs === null) return;
    if (!latest || capturedAtMs > latest.capturedAtMs) {
      latest = { dateKey, capturedAtMs };
    }
  });

  return latest;
}

function getNextConfiguredSnapshotUtcDate(hours, minutes, fromDate = new Date()) {
  const next = new Date(fromDate);
  next.setUTCHours(hours, minutes, 0, 0);
  if (next <= fromDate) next.setUTCDate(next.getUTCDate() + 1);
  return next;
}

function scheduleSnapshotTimeout(runAt, modeLabel) {
  if (snapshotTimer) clearTimeout(snapshotTimer);
  const delayMs = Math.max(0, runAt.getTime() - Date.now());

  console.log(`[snapshot] Scheduler mode: ${modeLabel}. Next run at ${runAt.toISOString()}`);

  snapshotTimer = setTimeout(async () => {
    await captureSnapshot();
    await scheduleNextSnapshot();
  }, delayMs);
}

async function scheduleRollingSnapshot() {
  const snapshots = await readSnapshots();
  const latest = getLatestSnapshotInfo(snapshots);

  if (!latest) {
    console.log("[snapshot] No previous snapshot found; capturing on startup.");
    await captureSnapshot();
    const refreshed = await readSnapshots();
    const latestAfterCapture = getLatestSnapshotInfo(refreshed);
    const nextRunAt = new Date((latestAfterCapture?.capturedAtMs || Date.now()) + SNAPSHOT_INTERVAL_MS);
    scheduleSnapshotTimeout(nextRunAt, "rolling-24h");
    return;
  }

  const nextRunAt = new Date(latest.capturedAtMs + SNAPSHOT_INTERVAL_MS);
  if (nextRunAt <= new Date()) {
    console.log("[snapshot] Next rolling snapshot is due now; capturing immediately.");
    await captureSnapshot();
    const refreshed = await readSnapshots();
    const latestAfterCapture = getLatestSnapshotInfo(refreshed);
    const delayedRunAt = new Date((latestAfterCapture?.capturedAtMs || Date.now()) + SNAPSHOT_INTERVAL_MS);
    scheduleSnapshotTimeout(delayedRunAt, "rolling-24h");
    return;
  }

  scheduleSnapshotTimeout(nextRunAt, "rolling-24h");
}

async function scheduleConfiguredSnapshot() {
  const parsed = parseSnapshotTimeUtc(SNAPSHOT_TIME_UTC);
  if (!parsed) {
    console.warn(
      `[snapshot] Invalid SNAPSHOT_TIME_UTC="${SNAPSHOT_TIME_UTC}" (expected HH:mm). Falling back to rolling 24h mode.`
    );
    await scheduleRollingSnapshot();
    return;
  }

  const nextRunAt = getNextConfiguredSnapshotUtcDate(parsed.hours, parsed.minutes);
  scheduleSnapshotTimeout(nextRunAt, `configured-utc (${SNAPSHOT_TIME_UTC})`);
}

async function scheduleNextSnapshot() {
  if (SNAPSHOT_TIME_UTC) {
    await scheduleConfiguredSnapshot();
    return;
  }

  await scheduleRollingSnapshot();
}

async function startSnapshotScheduler() {
  await scheduleNextSnapshot();
}

async function readSnapshots() {
  try {
    const raw = await fs.promises.readFile(SNAPSHOT_FILE, "utf8");
    return JSON.parse(raw);
  } catch {
    return {};
  }
}

async function writeSnapshots(data) {
  await fs.promises.mkdir(SNAPSHOT_DIR, { recursive: true });
  await fs.promises.writeFile(SNAPSHOT_FILE, JSON.stringify(data, null, 2));
}

async function captureSnapshot() {
  const today = new Date().toISOString().split("T")[0];
  const snapshots = await readSnapshots();
  if (snapshots[today]) return; // already captured today

  try {
    const client = await getGraphClient();
    const groupId = process.env.GROUP_ID;
    const plansRes = await client.api(`/groups/${groupId}/planner/plans`).get();

    const capturedAt = new Date().toISOString();
    const dayData = {
      __meta: {
        capturedAt,
      },
    };
    for (const plan of plansRes.value) {
      const tasks = await fetchAllPages(client, `/planner/plans/${plan.id}/tasks`);
      const total = tasks.length;
      const completed = tasks.filter((t) => t.percentComplete === 100).length;
      const overdue = tasks.filter(
        (t) => getDueStatus(t.dueDateTime, t.percentComplete) === "overdue"
      ).length;
      const dueSoon = tasks.filter(
        (t) => getDueStatus(t.dueDateTime, t.percentComplete) === "dueSoon"
      ).length;
      const unassigned = tasks.filter(
        (t) => !t.assignments || Object.keys(t.assignments).length === 0
      ).length;
      dayData[plan.id] = {
        total,
        completed,
        overdue,
        dueSoon,
        unassigned,
        percentComplete: total > 0 ? Math.round((completed / total) * 100) : 0,
      };
    }

    snapshots[today] = dayData;
    await writeSnapshots(snapshots);
    console.log(`Snapshot captured for ${today} at ${capturedAt}`);
  } catch (err) {
    console.error("Snapshot capture failed:", err.message);
  }
}

// Start snapshot scheduler
startSnapshotScheduler().catch((err) => {
  console.error("Snapshot scheduler failed:", err.message);
});

// Get all plans for the group
app.get("/api/plans", async (req, res) => {
  try {
    const client = await getGraphClient();
    const groupId = process.env.GROUP_ID;
    const plans = await client.api(`/groups/${groupId}/planner/plans`).get();
    res.json(plans.value);
  } catch (err) {
    console.error("Error fetching plans:", err);
    res.status(500).json({ error: err.message });
  }
});

// All-plans overview — now includes unassigned count
app.get("/api/overview", async (req, res) => {
  try {
    const overview = await withCache("overview", async () => {
      const client = await getGraphClient();
      const groupId = process.env.GROUP_ID;
      const plansRes = await client.api(`/groups/${groupId}/planner/plans`).get();

      return Promise.all(
        plansRes.value.map(async (plan) => {
          const tasks = await fetchAllPages(client, `/planner/plans/${plan.id}/tasks`);

          const total = tasks.length;
          const completed = tasks.filter((t) => t.percentComplete === 100).length;
          const inProgress = tasks.filter(
            (t) => t.percentComplete > 0 && t.percentComplete < 100
          ).length;
          const notStarted = tasks.filter((t) => t.percentComplete === 0).length;
          const overdue = tasks.filter(
            (t) => getDueStatus(t.dueDateTime, t.percentComplete) === "overdue"
          ).length;
          const dueSoon = tasks.filter(
            (t) => getDueStatus(t.dueDateTime, t.percentComplete) === "dueSoon"
          ).length;
          const unassigned = tasks.filter(
            (t) => !t.assignments || Object.keys(t.assignments).length === 0
          ).length;

          return {
            planId: plan.id,
            title: plan.title,
            total,
            completed,
            inProgress,
            notStarted,
            overdue,
            dueSoon,
            unassigned,
            percentComplete: total > 0 ? Math.round((completed / total) * 100) : 0,
          };
        })
      );
    });

    res.json(overview);
  } catch (err) {
    console.error("Error fetching overview:", err);
    res.status(500).json({ error: err.message });
  }
});

// Unassigned tasks across all plans
app.get("/api/unassigned", async (req, res) => {
  try {
    const data = await withCache("unassigned", async () => {
      const client = await getGraphClient();
      const groupId = process.env.GROUP_ID;
      const plansRes = await client.api(`/groups/${groupId}/planner/plans`).get();

      const allUnassigned = [];
      for (const plan of plansRes.value) {
        const tasks = await fetchAllPages(client, `/planner/plans/${plan.id}/tasks`);
        tasks
          .filter((t) => !t.assignments || Object.keys(t.assignments).length === 0)
          .forEach((t) =>
            allUnassigned.push({
              id: t.id,
              title: t.title,
              planTitle: plan.title,
              dueDateTime: t.dueDateTime,
              dueStatus: getDueStatus(t.dueDateTime, t.percentComplete),
              status:
                t.percentComplete === 100
                  ? "completed"
                  : t.percentComplete > 0
                  ? "inProgress"
                  : "notStarted",
            })
          );
      }

      return { count: allUnassigned.length, tasks: allUnassigned };
    });

    res.json(data);
  } catch (err) {
    console.error("Error fetching unassigned:", err);
    res.status(500).json({ error: err.message });
  }
});

// Team workload — assignee stats across all plans
app.get("/api/workload", async (req, res) => {
  try {
    const result = await withCache("workload", async () => {
      const client = await getGraphClient();
      const groupId = process.env.GROUP_ID;
      const plansRes = await client.api(`/groups/${groupId}/planner/plans`).get();

      const allTasks = [];
      for (const plan of plansRes.value) {
        const tasks = await fetchAllPages(client, `/planner/plans/${plan.id}/tasks`);
        allTasks.push(...tasks);
      }

      // Collect unique user IDs
      const allUserIds = new Set();
      allTasks.forEach((t) => {
        if (t.assignments) {
          Object.keys(t.assignments).forEach((uid) => allUserIds.add(uid));
        }
      });

      const userMap = await resolveUsers(client, [...allUserIds]);

      // Build workload map
      const workload = {};
      allTasks.forEach((t) => {
        const assigneeIds = t.assignments ? Object.keys(t.assignments) : [];
        const status =
          t.percentComplete === 100
            ? "completed"
            : t.percentComplete > 0
            ? "inProgress"
            : "notStarted";
        const dueStatus = getDueStatus(t.dueDateTime, t.percentComplete);
        const isOverdue = dueStatus === "overdue";

        const names =
          assigneeIds.length > 0
            ? assigneeIds.map((uid) => userMap.get(uid) || uid)
            : ["Unassigned"];

        names.forEach((name) => {
          if (!workload[name]) {
            workload[name] = {
              displayName: name,
              completed: 0,
              inProgress: 0,
              notStarted: 0,
              total: 0,
              overdue: 0,
            };
          }
          workload[name][status]++;
          workload[name].total++;
          if (isOverdue) workload[name].overdue++;
        });
      });

      return Object.values(workload).sort((a, b) => b.total - a.total);
    });

    res.json(result);
  } catch (err) {
    console.error("Error fetching workload:", err);
    res.status(500).json({ error: err.message });
  }
});

// Task details (description/notes) — lazy loaded
app.get("/api/tasks/:taskId/details", async (req, res) => {
  try {
    const client = await getGraphClient();
    const details = await client
      .api(`/planner/tasks/${req.params.taskId}/details`)
      .get();
    res.json({ description: details.description || "" });
  } catch (err) {
    console.error("Error fetching task details:", err);
    res.status(500).json({ error: err.message });
  }
});

// Trend data — last 14 days of snapshots
app.get("/api/trends", async (req, res) => {
  try {
    const snapshots = await readSnapshots();
    const dates = Object.keys(snapshots).sort().slice(-14);

    const trends = {};
    dates.forEach((date) => {
      const dayData = snapshots[date];
      getSnapshotPlanEntries(dayData).forEach(([planId, stats]) => {
        if (!trends[planId]) trends[planId] = [];
        trends[planId].push({ date, percentComplete: stats.percentComplete || 0 });
      });
    });

    res.json(trends);
  } catch (err) {
    console.error("Error fetching trends:", err);
    res.status(500).json({ error: err.message });
  }
});

// Snapshot deltas — latest day vs previous captured day
app.get("/api/deltas", async (req, res) => {
  try {
    const snapshots = await readSnapshots();
    const dates = Object.keys(snapshots).sort();

    if (dates.length < 2) {
      res.json({
        currentDate: dates.length === 1 ? dates[0] : null,
        previousDate: null,
        portfolio: null,
        byPlan: {},
      });
      return;
    }

    const currentDate = dates[dates.length - 1];
    const previousDate = dates[dates.length - 2];
    const current = snapshots[currentDate] || {};
    const previous = snapshots[previousDate] || {};
    const planIds = new Set([
      ...getSnapshotPlanEntries(current).map(([planId]) => planId),
      ...getSnapshotPlanEntries(previous).map(([planId]) => planId),
    ]);

    const byPlan = {};
    for (const planId of planIds) {
      const now = current[planId] || {};
      const prev = previous[planId] || {};
      byPlan[planId] = {
        completedDelta: (now.completed || 0) - (prev.completed || 0),
        overdueDelta: (now.overdue || 0) - (prev.overdue || 0),
        dueSoonDelta: (now.dueSoon || 0) - (prev.dueSoon || 0),
        unassignedDelta: (now.unassigned || 0) - (prev.unassigned || 0),
      };
    }

    const portfolio = Object.values(byPlan).reduce(
      (acc, planDelta) => ({
        completedDelta: acc.completedDelta + planDelta.completedDelta,
        overdueDelta: acc.overdueDelta + planDelta.overdueDelta,
        dueSoonDelta: acc.dueSoonDelta + planDelta.dueSoonDelta,
        unassignedDelta: acc.unassignedDelta + planDelta.unassignedDelta,
      }),
      { completedDelta: 0, overdueDelta: 0, dueSoonDelta: 0, unassignedDelta: 0 }
    );

    res.json({
      currentDate,
      previousDate,
      portfolio,
      byPlan,
    });
  } catch (err) {
    console.error("Error fetching deltas:", err);
    res.status(500).json({ error: err.message });
  }
});

// Get buckets and tasks for a specific plan
app.get("/api/plans/:planId/stats", async (req, res) => {
  try {
    const client = await getGraphClient();
    const { planId } = req.params;

    const [buckets, tasks] = await Promise.all([
      fetchAllPages(client, `/planner/plans/${planId}/buckets`),
      fetchAllPages(client, `/planner/plans/${planId}/tasks`),
    ]);

    const allUserIds = new Set();
    tasks.forEach((t) => {
      if (t.assignments) {
        Object.keys(t.assignments).forEach((uid) => allUserIds.add(uid));
      }
    });

    const userMap = await resolveUsers(client, [...allUserIds]);

    const stats = buckets.map((bucket) => {
      const bucketTasks = tasks.filter((t) => t.bucketId === bucket.id);

      const enrichedTasks = bucketTasks.map((t) => {
        const assigneeIds = t.assignments ? Object.keys(t.assignments) : [];
        const status =
          t.percentComplete === 100
            ? "completed"
            : t.percentComplete > 0
            ? "inProgress"
            : "notStarted";

        return {
          id: t.id,
          title: t.title,
          bucketName: bucket.name,
          status,
          percentComplete: t.percentComplete,
          hasDescription: t.hasDescription !== false,
          assignees: assigneeIds.map((uid) => ({
            userId: uid,
            displayName: userMap.get(uid) || uid,
          })),
          dueDateTime: t.dueDateTime || null,
          dueStatus: getDueStatus(t.dueDateTime, t.percentComplete),
        };
      });

      const assigneeSummary = {};
      enrichedTasks.forEach((t) => {
        t.assignees.forEach((a) => {
          assigneeSummary[a.displayName] =
            (assigneeSummary[a.displayName] || 0) + 1;
        });
      });

      return {
        bucketId: bucket.id,
        bucketName: bucket.name,
        notStarted: enrichedTasks.filter((t) => t.status === "notStarted").length,
        inProgress: enrichedTasks.filter((t) => t.status === "inProgress").length,
        completed: enrichedTasks.filter((t) => t.status === "completed").length,
        total: enrichedTasks.length,
        overdue: enrichedTasks.filter((t) => t.dueStatus === "overdue").length,
        dueSoon: enrichedTasks.filter((t) => t.dueStatus === "dueSoon").length,
        onTrack: enrichedTasks.filter((t) => t.dueStatus === "onTrack").length,
        noDueDate: enrichedTasks.filter((t) => t.dueStatus === "noDueDate").length,
        assigneeSummary,
        tasks: enrichedTasks,
      };
    });

    res.json(stats);
  } catch (err) {
    console.error("Error fetching plan stats:", err);
    res.status(500).json({ error: err.message });
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`Planner Stats running on http://localhost:${PORT}`);
});
