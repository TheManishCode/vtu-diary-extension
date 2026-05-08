importScripts("lib/jspdf.umd.min.js");

const { jsPDF } = self.jspdf;

const VTU_WEB_HOST = "vtu.internyet.in";
const VTU_API_HOST = "vtuapi.internyet.in";
const VTU_DIARY_PATH_HINTS = [
  "student-diary",
  "create-diary-entry",
  "diary-entries",
  "edit-diary-entry"
];

// ============================================================================
// Job Queue Infrastructure
// ============================================================================
const DEFAULT_BATCH_SIZE = 5;
const DEFAULT_MAX_ATTEMPTS = 50;
const DEFAULT_RETRY_DELAY_MS = 1500;
const DEFAULT_REQUEST_DELAY_MS = 300;
const MAX_CONCURRENT_UPLOADS = 1;

let activeUploadCount = 0;
let cancelUploadFlag = false; // Flag to cancel ongoing uploads
const jobQueue = [];
const activeDedupKeys = new Map();
const activeJobStates = new Map();
// Runtime-config storage key and connected progress ports (SSE-like)
const RUNTIME_CONFIG_KEY = "vtu_runtime_config";
const progressPorts = new Set();
// Skill cache (per-internship) to avoid repeated skill fetches
const SKILL_CACHE_KEY = "vtu_skill_cache";
const skillCache = new Map();

// Simple statistics
const STATS_KEY = "vtu_stats";
const DEDUP_CACHE_KEY = "vtu_dedup_cache";
let runtimeStats = { totalUploaded: 0, totalFailed: 0, totalSkipped: 0, lastUploadAt: 0 };

// Load persisted caches/stats (best-effort)
chrome.storage.local.get([SKILL_CACHE_KEY, STATS_KEY], (res) => {
  try {
    const sc = res?.[SKILL_CACHE_KEY] || {};
    Object.keys(sc || {}).forEach((k) => skillCache.set(k, sc[k]));
    runtimeStats = Object.assign(runtimeStats, res?.[STATS_KEY] || {});
  } catch (_) {}
});

function saveSkillCache() {
  try {
    const obj = {};
    for (const [k, v] of skillCache.entries()) obj[k] = v;
    chrome.storage.local.set({ [SKILL_CACHE_KEY]: obj });
  } catch (_) {}
}

function saveStats() {
  try {
    chrome.storage.local.set({ [STATS_KEY]: runtimeStats });
  } catch (_) {}
}

function recordSuccess() {
  runtimeStats.totalUploaded = (runtimeStats.totalUploaded || 0) + 1;
  runtimeStats.lastUploadAt = Date.now();
  saveStats();
}

function recordFailure() {
  runtimeStats.totalFailed = (runtimeStats.totalFailed || 0) + 1;
  runtimeStats.lastUploadAt = Date.now();
  saveStats();
}

// Persistent deduplication cache with TTL (7 days)
const DEDUP_TTL_MS = 7 * 24 * 60 * 60 * 1000; // 7 days

async function getDedupCache() {
  return new Promise((resolve) => {
    chrome.storage.local.get([DEDUP_CACHE_KEY], (res) => {
      const cache = res?.[DEDUP_CACHE_KEY] || {};
      // Clean up expired entries
      const now = Date.now();
      const cleaned = {};
      for (const [key, entry] of Object.entries(cache)) {
        if (entry && entry.ts && (now - entry.ts) < DEDUP_TTL_MS) {
          cleaned[key] = entry;
        }
      }
      resolve(cleaned);
    });
  });
}

async function saveDedupCache(cache) {
  return new Promise((resolve) => {
    chrome.storage.local.set({ [DEDUP_CACHE_KEY]: cache }, resolve);
  });
}

function generateDedupEntryKey(email, internshipId, entryDate, entryHours) {
  // Create a unique key from entry metadata
  const parts = [
    String(email || "").toLowerCase().trim(),
    String(internshipId || "").trim(),
    String(entryDate || "").trim(),
    String(entryHours || "").trim()
  ];
  return parts.filter(Boolean).join(':');
}

async function isEntryDuplicated(email, internshipId, entryDate, entryHours) {
  const cache = await getDedupCache();
  const key = generateDedupEntryKey(email, internshipId, entryDate, entryHours);
  return !!cache[key];
}

async function markEntryAsUploaded(email, internshipId, entryDate, entryHours) {
  const cache = await getDedupCache();
  const key = generateDedupEntryKey(email, internshipId, entryDate, entryHours);
  cache[key] = { ts: Date.now() };
  await saveDedupCache(cache);
}


async function getRuntimeConfig() {
  return new Promise((resolve) => {
    chrome.storage.local.get([RUNTIME_CONFIG_KEY], (res) => {
      const stored = res?.[RUNTIME_CONFIG_KEY] || {};
      resolve({
        batchSize: Number(stored.batchSize) || DEFAULT_BATCH_SIZE,
        maxAttempts: Number(stored.maxAttempts) || DEFAULT_MAX_ATTEMPTS,
        retryDelayMs: Number(stored.retryDelayMs) || DEFAULT_RETRY_DELAY_MS,
        requestDelayMs: Number(stored.requestDelayMs) || DEFAULT_REQUEST_DELAY_MS,
        maxConcurrentUploads: Number(stored.maxConcurrentUploads) || MAX_CONCURRENT_UPLOADS
      });
    });
  });
}

function generateDedupKey(internshipId, tabId) {
  return `${String(internshipId || "").trim()}-${String(tabId || "").trim()}`;
}

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

let uploadInProgress = false;
const LOG_STORAGE_KEY = "vtu_runtime_logs";
const MAX_STORED_LOGS = 500;

function getStoredLogs() {
  return new Promise((resolve) => {
    chrome.storage.local.get([LOG_STORAGE_KEY], (result) => {
      const logs = Array.isArray(result?.[LOG_STORAGE_KEY]) ? result[LOG_STORAGE_KEY] : [];
      resolve(logs);
    });
  });
}

function setStoredLogs(logs) {
  return new Promise((resolve) => {
    chrome.storage.local.set({ [LOG_STORAGE_KEY]: logs }, () => {
      resolve();
    });
  });
}

async function persistLog(text) {
  const entryText = String(text || "").trim();
  if (!entryText) {
    return;
  }

  const current = await getStoredLogs();
  current.push({ text: entryText, ts: Date.now() });
  const trimmed = current.slice(-MAX_STORED_LOGS);
  await setStoredLogs(trimmed);
}

function clearPersistedLogs() {
  return setStoredLogs([]);
}

function send(text) {
  void persistLog(text);
  chrome.runtime.sendMessage({ type: "log", text, source: "background" }, () => {
    void chrome.runtime.lastError;
  });
}

// ============================================================================
// Queue Management Functions
// ============================================================================

async function tryDequeueNextJob() {
  if (activeUploadCount >= MAX_CONCURRENT_UPLOADS) {
    return;
  }

  if (jobQueue.length === 0) {
    return;
  }

  const job = jobQueue.shift();
  if (!job) {
    return;
  }

  activeUploadCount += 1;
  const dedupKey = generateDedupKey(job.internshipId, job.tabId);
  activeDedupKeys.set(dedupKey, true);

  try {
    const jobStateId = `${Date.now()}-${Math.random()}`;
    activeJobStates.set(jobStateId, {
      status: "processing",
      startTime: Date.now(),
      internshipId: job.internshipId,
      tabId: job.tabId,
      entryCount: Array.isArray(job.entries) ? job.entries.length : 0
    });

    await uploadEntriesWithBatching(job.tabId, job.entries, job.options || {});

    activeJobStates.set(jobStateId, {
      ...activeJobStates.get(jobStateId),
      status: "completed",
      endTime: Date.now()
    });
  } catch (error) {
    activeJobStates.set(jobStateId, {
      ...activeJobStates.get(jobStateId),
      status: "failed",
      endTime: Date.now(),
      error: error?.message || "Unknown error"
    });
    send(`❌ Job processing failed: ${error?.message || "Unknown error"}`);
  } finally {
    activeUploadCount -= 1;
    activeDedupKeys.delete(dedupKey);

    // Recursively process next job in queue
    await tryDequeueNextJob();
  }
}

async function uploadEntriesWithBatching(tabId, rawEntries, options = {}) {
  const parsed = normalizeUploadEntries(rawEntries);

  const receivedCount = Array.isArray(rawEntries)
    ? rawEntries.length
    : (rawEntries && typeof rawEntries === "object" ? 1 : 0);
  send(`📦 Received ${receivedCount} upload rows`);

  if (parsed.rejected.length) {
    send(`⚠️ Skipped ${parsed.rejected.length} invalid rows`);
    parsed.rejected.slice(0, 5).forEach((r) => {
      send(`Row ${r.index}: ${r.reason}`);
    });
  }

  if (!parsed.valid.length) {
    send("❌ No valid entries to upload");
    return;
  }

  const dedupedByDate = [];
  const seenDates = new Set();
  let duplicateDateRows = 0;

  parsed.valid.forEach((entry) => {
    if (seenDates.has(entry.date)) {
      duplicateDateRows += 1;
      return;
    }
    seenDates.add(entry.date);
    dedupedByDate.push(entry);
  });

  if (duplicateDateRows > 0) {
    send(`⚠️ Skipped ${duplicateDateRows} duplicate date row(s) from input`);
  }

  send(`✅ Valid rows ready: ${dedupedByDate.length}`);

  // Process entries in batches with round-robin retry logic
  const config = await getRuntimeConfig();
  const batchSize = config.batchSize || DEFAULT_BATCH_SIZE;
  const batches = [];

  for (let i = 0; i < dedupedByDate.length; i += batchSize) {
    batches.push(dedupedByDate.slice(i, i + batchSize));
  }

  send(`📤 Starting upload with ${batches.length} batch(es)...`);

  // Attach cached skillIds to options if available (dynamic caching)
  const internshipKey = String(options?.internshipId || "");
  if (internshipKey && skillCache.has(internshipKey)) {
    options = Object.assign({}, options, { skillIds: skillCache.get(internshipKey) });
    send(`ℹ️ Using cached skills for internship ${internshipKey}`);
  }

  for (let batchIndex = 0; batchIndex < batches.length; batchIndex += 1) {
    // Check if cancel flag is set
    if (cancelUploadFlag) {
      send("⛔ Upload cancelled by user");
      cancelUploadFlag = false;
      notifyPopupProgress(0, 0);
      return;
    }

    const batch = batches[batchIndex];
    send(`📋 Processing batch ${batchIndex + 1}/${batches.length} (${batch.length} entries)`);

    for (let entryIndex = 0; entryIndex < batch.length; entryIndex += 1) {
      // Check cancel flag again before each entry
      if (cancelUploadFlag) {
        send("⛔ Upload cancelled by user");
        cancelUploadFlag = false;
        notifyPopupProgress(0, 0);
        return;
      }

      const entry = batch[entryIndex];
      await tryUploadEntry(tabId, entry, options, batchIndex, entryIndex, batch.length);

      // Request delay between entries
      if (entryIndex < batch.length - 1) {
        await sleep(config.requestDelayMs || DEFAULT_REQUEST_DELAY_MS);
      }
    }

    // Batch delay
    if (batchIndex < batches.length - 1) {
      await sleep(DEFAULT_REQUEST_DELAY_MS * 2);
    }
  }

  notifyPopupProgress(dedupedByDate.length, dedupedByDate.length);
}

async function tryUploadEntry(tabId, entry, options = {}, batchIndex, entryIndex, batchSize) {
  const maxAttempts = DEFAULT_MAX_ATTEMPTS;
  let lastError = null;

  // Check if entry is already deduplicated (from a previous run)
  try {
    const email = options?.email || "unknown@vtu.edu.in";
    const internshipId = options?.internshipId || "unknown";
    if (await isEntryDuplicated(email, internshipId, entry.date, entry.hours)) {
      send(`⏭️  Skipped duplicate: ${entry.date}`);
      runtimeStats.totalSkipped = (runtimeStats.totalSkipped || 0) + 1;
      saveStats();
      return;
    }
  } catch (e) {
    // Continue even if dedup check fails
  }

  for (let attempt = 0; attempt < maxAttempts; attempt += 1) {
    try {
      const result = await chrome.scripting.executeScript({
        target: { tabId },
        args: [entry, options, attempt],
        func: async (entry, options, attemptNum) => {
          const STORE_URL = "https://vtuapi.internyet.in/api/v1/student/internship-diaries/store";
          const LIST_URL = "https://vtuapi.internyet.in/api/v1/student/internship-diaries";

          const clean = (v) => String(v || "").replace(/\s+/g, " ").trim();
          const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

          // Fetch internship ID if needed
          let internshipId = clean(options?.internshipId || "");
          // Accept pre-cached skills passed via options to avoid repeated fetches
          let skillIds = Array.isArray(options?.skillIds) && options.skillIds.length ? options.skillIds : [];

          if (!internshipId) {
            try {
              const diaryListRes = await fetch(LIST_URL, {
                credentials: "include",
                headers: { Accept: "application/json" }
              });

              if (diaryListRes.ok) {
                const diaryListData = await diaryListRes.json();
                const diaryList = diaryListData?.data?.data || diaryListData?.data || [];
                if (Array.isArray(diaryList) && diaryList.length > 0) {
                  internshipId = String(diaryList[0].internship_id || "");
                }
              }
            } catch (_) {
              // Continue without internship ID
            }
          }

          if (!internshipId) {
            return {
              success: false,
              status: 400,
              message: "Could not resolve internship ID",
              internshipId: internshipId,
              skillIds: skillIds
            };
          }

          // Fetch skills only if not provided via options/cache
          if (!Array.isArray(skillIds) || skillIds.length === 0) {
            try {
              const skillsRes = await fetch("https://vtuapi.internyet.in/api/v1/master/skills", {
                credentials: "include",
                headers: { Accept: "application/json" }
              });

              if (skillsRes.ok) {
                const skillsData = await skillsRes.json();
                const skillsList = skillsData?.data?.data || skillsData?.data || [];
                if (Array.isArray(skillsList) && skillsList.length > 0) {
                  skillIds = [String(skillsList[0].id)];
                }
              }
            } catch (_) {
              // Continue without skills
            }
          }

          // Check for existing entry
          let existingEntryId = null;
          try {
            const listRes = await fetch(LIST_URL, {
              credentials: "include",
              headers: { Accept: "application/json" }
            });

            if (listRes.ok) {
              const listData = await listRes.json();
              const diaryList = listData?.data?.data || listData?.data || [];
              if (Array.isArray(diaryList)) {
                const existing = diaryList.find((e) => e?.date === entry.date);
                if (existing?.id) {
                  existingEntryId = String(existing.id);
                }
              }
            }
          } catch (_) {
            // Continue without checking existing entries
          }

          // Prepare payload
          const payload = {
            date: entry.date,
            hours: Number(entry.hours),
            description: entry.description,
            learnings: entry.learnings,
            skill_ids: skillIds,
            mood_slider: 5,
            internship_id: internshipId
          };

          // Upload or update entry
          const targetUrl = existingEntryId
            ? `https://vtuapi.internyet.in/api/v1/student/internship-diaries/${existingEntryId}`
            : STORE_URL;

          const targetMethod = existingEntryId ? "PATCH" : "POST";

          const controller = new AbortController();
          const timeoutId = setTimeout(() => controller.abort(), 15000);

          const res = await fetch(targetUrl, {
            method: targetMethod,
            credentials: "include",
            headers: {
              Accept: "application/json",
              "Content-Type": "application/json"
            },
            body: JSON.stringify(payload),
            signal: controller.signal
          });

          clearTimeout(timeoutId);

          if (res.ok) {
            return {
              success: true,
              status: res.status,
              message: "Entry uploaded successfully",
              internshipId: internshipId,
              skillIds: skillIds
            };
          }

          // Check for 401 - session expired
          if (res.status === 401) {
            return {
              success: false,
              status: 401,
              message: "Session expired - requires re-login",
              shouldRefreshSession: true,
              internshipId: internshipId,
              skillIds: skillIds
            };
          }

          // Check for 429 - rate limit
          if (res.status === 429) {
            const retryAfter = Number(res.headers.get("Retry-After") || "0");
            return {
              success: false,
              status: 429,
              message: "Rate limited",
              retryAfter: retryAfter > 0 ? retryAfter * 1000 : 5000,
              internshipId: internshipId,
              skillIds: skillIds
            };
          }

          return {
            success: false,
            status: res.status,
            message: `Failed with HTTP ${res.status}`,
            internshipId: internshipId,
            skillIds: skillIds
          };
        }
      });

      const scriptResult = result?.[0]?.result;

      if (!scriptResult) {
        lastError = "No response from tab script";
        await sleep(DEFAULT_RETRY_DELAY_MS);
        continue;
      }

      if (scriptResult.success) {
        send(`✅ Entry ${entry.date} uploaded successfully`);
        notifyPopupProgress(batchIndex * DEFAULT_BATCH_SIZE + entryIndex + 1, batchSize);
        // record stats
        try { recordSuccess(); } catch (_) {}
        // mark entry as uploaded (for persistent dedup)
        try {
          const email = options?.email || "unknown@vtu.edu.in";
          const internshipId = options?.internshipId || "unknown";
          await markEntryAsUploaded(email, internshipId, entry.date, entry.hours);
        } catch (_) {}
        // cache skillIds if returned
        try {
          const iid = String(scriptResult.internshipId || internshipKey || "");
          if (iid && Array.isArray(scriptResult.skillIds) && scriptResult.skillIds.length) {
            skillCache.set(iid, scriptResult.skillIds);
            saveSkillCache();
          }
        } catch (_) {}
        return;
      }

      // Handle 401 - auto-refresh session
      if (scriptResult.status === 401 && scriptResult.shouldRefreshSession) {
        send(`⚠️ Session expired. Attempting to refresh...`);
        try {
          await chrome.scripting.executeScript({
            target: { tabId },
            func: async () => {
              // Attempt to navigate to refresh session
              window.location.href = "https://vtu.internyet.in/dashboard/student";
              await new Promise((resolve) => setTimeout(resolve, 3000));
            }
          });
        } catch (_) {
          // Continue with retry
        }
        lastError = "Session refresh attempted";
        await sleep(DEFAULT_RETRY_DELAY_MS * 3);
        continue;
      }

      // Handle 429 - rate limit
      if (scriptResult.status === 429) {
        const waitMs = scriptResult.retryAfter || 5000;
        send(`⏳ Rate limited. Waiting ${Math.ceil(waitMs / 1000)}s...`);
        await sleep(waitMs);
        continue;
      }

      lastError = scriptResult.message || "Upload failed";

      // Exponential backoff for retries
      const backoffMs = DEFAULT_RETRY_DELAY_MS * Math.pow(2, Math.min(attempt, 5));
      await sleep(backoffMs);

    } catch (error) {
      lastError = error?.message || "Unknown error";
      await sleep(DEFAULT_RETRY_DELAY_MS);
    }
  }

  send(`❌ Entry ${entry.date} failed after ${maxAttempts} attempts: ${lastError}`);
  try { recordFailure(); } catch (_) {}
}

function broadcastProgressUpdate(payload) {
  // Send to connected ports (SSE-like) and also send a runtime message
  try {
    for (const port of progressPorts) {
      try {
        port.postMessage({ type: 'upload_progress', ...payload });
      } catch (_) {
        // ignore
      }
    }
  } catch (_) {}

  // Also keep existing message channel for popups that listen with runtime.onMessage
  try {
    chrome.runtime.sendMessage({ type: 'upload_progress', ...payload }, () => void chrome.runtime.lastError);
  } catch (_) {}
}

// Keep notifyPopupProgress compatible: also broadcast via ports
function notifyPopupProgress(uploaded, total) {
  const progress = total > 0 ? Math.round((uploaded / total) * 100) : 0;
  const payload = { uploaded, total, progress };
  broadcastProgressUpdate(payload);
}

// Accept long-lived connections from popup (SSE-like)
chrome.runtime.onConnect.addListener((port) => {
  if (!port || !port.name) return;
  if (port.name !== 'vtu-progress') return;
  progressPorts.add(port);
  // Send current queue state immediately
  try {
    port.postMessage({ type: 'connected', queueLength: jobQueue.length });
  } catch (_) {}

  port.onDisconnect.addListener(() => {
    progressPorts.delete(port);
  });
});

function pruneJobStates(maxItems = 200) {
  try {
    if (activeJobStates.size <= maxItems) return;
    const arr = Array.from(activeJobStates.entries()).map(([k, v]) => ({ k, v }));
    arr.sort((a, b) => (a.v.startTime || 0) - (b.v.startTime || 0));
    const toRemove = arr.slice(0, Math.max(0, arr.length - maxItems));
    for (const r of toRemove) {
      activeJobStates.delete(r.k);
    }
  } catch (_) {}
}

  // Keepalive pings to connected ports to keep popup informed and encourage SW stay-alive
  setInterval(() => {
    try {
      const payload = { type: 'keepalive', ts: Date.now(), queueLength: jobQueue.length };
      for (const p of progressPorts) {
        try { p.postMessage(payload); } catch (_) {}
      }
      // also prune job states periodically
      pruneJobStates(200);
    } catch (_) {}
  }, 25000);

function parseUrl(rawUrl) {
  try {
    return new URL(String(rawUrl || ""));
  } catch (error) {
    return null;
  }
}

function getVtuTabContext(url) {
  const parsed = parseUrl(url);
  if (!parsed) {
    return {
      isSupportedHost: false,
      isApiHost: false,
      isDiaryPath: false
    };
  }

  const host = parsed.hostname;
  const isApiHost = host === VTU_API_HOST;
  const isWebHost = host === VTU_WEB_HOST;
  const isSupportedHost = isApiHost || isWebHost;

  const fullPath = `${parsed.pathname}${parsed.search}${parsed.hash}`.toLowerCase();
  const isDiaryPath = VTU_DIARY_PATH_HINTS.some((hint) => fullPath.includes(hint));

  return {
    isSupportedHost,
    isApiHost,
    isDiaryPath
  };
}

function scoreUploadTab(tab) {
  const ctx = getVtuTabContext(tab?.url);
  if (!ctx.isSupportedHost) {
    return -1;
  }

  let score = 0;
  if (ctx.isApiHost) {
    score += 100;
  }
  if (ctx.isDiaryPath) {
    score += 100;
  }
  if (tab?.active) {
    score += 10;
  }
  if (typeof tab?.lastAccessed === "number") {
    score += Math.floor(tab.lastAccessed / 1000000000);
  }

  return score;
}

async function resolveUploadTab() {
  const activeInWindow = await chrome.tabs.query({ active: true, currentWindow: true });
  const activeTab = activeInWindow?.[0] || null;
  const activeContext = getVtuTabContext(activeTab?.url);
  if (activeTab?.id && activeContext.isSupportedHost) {
    return { tab: activeTab, fromFallback: false };
  }

  const allTabs = await chrome.tabs.query({});
  const candidates = allTabs
    .filter((t) => t && t.id)
    .map((t) => ({ tab: t, score: scoreUploadTab(t) }))
    .filter((x) => x.score >= 0)
    .sort((a, b) => b.score - a.score);

  if (!candidates.length) {
    return { tab: activeTab, fromFallback: false };
  }

  return { tab: candidates[0].tab, fromFallback: true };
}

async function fetchInternshipsForUpload(tabId) {
  const scriptResult = await chrome.scripting.executeScript({
    target: { tabId },
    func: async () => {
      const clean = (value) => String(value || "").replace(/\s+/g, " ").trim();

      async function fetchJson(url, timeoutMs = 10000) {
        const controller = new AbortController();
        const timeoutId = setTimeout(() => controller.abort(), timeoutMs);
        try {
          const response = await fetch(url, {
            credentials: "include",
            headers: { Accept: "application/json" },
            signal: controller.signal
          });
          clearTimeout(timeoutId);
          if (!response.ok) {
            return null;
          }
          return await response.json();
        } catch (_) {
          clearTimeout(timeoutId);
          return null;
        }
      }

      function addByMap(map, idValue, nameValue, source) {
        const id = clean(idValue);
        if (!id || !/^\d+$/.test(id)) {
          return;
        }

        const name = clean(nameValue) || `Internship ${id}`;
        const existing = map.get(id);

        if (!existing) {
          map.set(id, {
            id,
            name,
            source,
            order: map.size
          });
          return;
        }

        if ((!existing.name || /^Internship\s+\d+$/i.test(existing.name)) && name) {
          existing.name = name;
        }
      }

      const internshipMap = new Map();

      const applyUrls = [
        "https://vtuapi.internyet.in/api/v1/student/internship-applys?page=1&status=6",
        "https://vtuapi.internyet.in/api/v1/student/internship-applys?page=1",
        "https://vtuapi.internyet.in/api/v1/student/internship-applys"
      ];

      for (const url of applyUrls) {
        const json = await fetchJson(url);
        if (!json) {
          continue;
        }

        const list = json?.data?.data || json?.data || [];

        if (!Array.isArray(list) || !list.length) {
          continue;
        }

        list.forEach((item) => {
          addByMap(internshipMap, item.internship_id, item.internship_details?.name, "apply");
          if (item.company) {
            addByMap(internshipMap, item.company?.id, item.company?.name, "company");
          }
        });

        if (internshipMap.size > 0) {
          break;
        }
      }

      return Array.from(internshipMap.values());
    }
  });

  return scriptResult?.[0]?.result || [];
}

function normalizeText(text) {
  return String(text || "").replace(/\s+/g, " ").trim();
}

function normalizeProfile(profile) {
  return {
    name: normalizeText(profile.name),
    usn: normalizeText(profile.usn),
    college: normalizeText(profile.college),
    internship: normalizeText(profile.internship)
  };
}

function mergeProfiles(a, b) {
  return {
    name: normalizeText(b.name || a.name),
    usn: normalizeText(b.usn || a.usn),
    college: normalizeText(b.college || a.college),
    internship: normalizeText(b.internship || a.internship)
  };
}

function generateLearning(activityText) {
  const text = normalizeText(activityText);
  if (!text) {
    return "Practical learning and skill development.";
  }
  return `Gained practical understanding of ${text.slice(0, 100)}. Improved technical competencies and problem-solving skills.`;
}

function escapeHtml(text) {
  const map = {
    "&": "&amp;",
    "<": "&lt;",
    ">": "&gt;",
    '"': "&quot;",
    "'": "&#039;"
  };
  return String(text || "").replace(/[&<>"']/g, (m) => map[m]);
}

function safeText(text) {
  try {
    return String(text || "").slice(0, 5000);
  } catch (_) {
    return "";
  }
}

function buildDocHtml(entries, profile) {
  const p = normalizeProfile(profile);
  return `
  <!DOCTYPE html>
  <html>
  <head>
    <meta charset="UTF-8">
    <title>Internship Diary</title>
    <style>
      body {
        font-family: "Times New Roman", serif;
        padding: 20px;
        line-height: 1.6;
      }

      .page {
        page-break-after: always;
        margin-bottom: 40px;
      }

      .entry-page {
        margin-top: 20px;
      }

      .meta-line {
        margin: 5px 0;
      }

      .section-title {
        font-weight: bold;
        margin-top: 10px;
        margin-bottom: 5px;
      }

      .section-text {
        margin-bottom: 15px;
        text-align: justify;
      }

      .footer {
        display: flex;
        justify-content: space-between;
        margin-top: 60px;
        font-size: 12pt;
      }

      .page-number {
        text-align: center;
        margin-top: 6mm;
        font-size: 10pt;
      }

      .cover {
        text-align: center;
        padding: 80px 20px;
      }

      .cover h1 {
        font-size: 20pt;
        margin-bottom: 20pt;
      }

      .cover h2 {
        font-size: 16pt;
        margin-bottom: 40pt;
      }

      .cover p {
        font-size: 14pt;
        margin: 10pt 0;
      }
    </style>
  </head>

  <body>

  <div class="page cover">
    <h1>VISVESVARAYA TECHNOLOGICAL UNIVERSITY</h1>
    <h2>INTERNSHIP DIARY</h2>

    <p><b>Name:</b> ${escapeHtml(p.name)}</p>
    <p><b>USN:</b> ${escapeHtml(p.usn)}</p>
    <p><b>College:</b> ${escapeHtml(p.college)}</p>
    <p><b>Internship:</b> ${escapeHtml(p.internship)}</p>
  </div>

  ${entries.map((e, i) => `
    <div class="page entry-page">

      <h1>VISVESVARAYA TECHNOLOGICAL UNIVERSITY</h1>
      <h2>INTERNSHIP DIARY</h2>

      <p class="meta-line"><b>Entry:</b> ${escapeHtml(i + 1)} of ${escapeHtml(entries.length)}</p>

      <p class="meta-line"><b>DATE:</b> ${escapeHtml(e.date || "-")}</p>
      <p class="meta-line"><b>Number of hours:</b> ${escapeHtml(e.hours || "-")}</p>

      <p class="section-title"><b>Description:</b></p>
      <p class="section-text">${escapeHtml(e.activity || "-")}</p>

      <p class="section-title"><b>Learnings/outcomes:</b></p>
      <p class="section-text">${escapeHtml(e.learning || generateLearning(e.activity || ""))}</p>

      <div class="footer">
        <div>Signature of External Coordinator</div>
        <div>Signature of Internship Coordinator</div>
      </div>

      <div class="page-number">
        Page ${i + 1} of ${entries.length}
      </div>

    </div>
  `).join("")}

  </body>
  </html>
  `;
}

function drawPdf(entries, profile) {
  const p = normalizeProfile(profile);
  const doc = new jsPDF("p", "mm", "a4");

  function centerText(text, y, size, bold) {
    doc.setFont("times", bold ? "bold" : "normal");
    doc.setFontSize(size);
    doc.text(String(text || ""), 105, y, { align: "center" });
  }

  // Cover page styling
  centerText("VISVESVARAYA TECHNOLOGICAL UNIVERSITY", 42, 18, true);
  centerText("INTERNSHIP DIARY", 56, 15, true);
  doc.setLineWidth(0.6);
  doc.line(25, 63, 185, 63);

  doc.setFont("times", "italic");
  doc.setFontSize(11);
  centerText("Internship Progress Record", 72, 11, false);

  doc.setDrawColor(0, 0, 0);
  doc.setLineWidth(0.4);
  doc.roundedRect(22, 86, 166, 66, 2, 2);

  doc.setFont("times", "bold");
  doc.setFontSize(12);
  doc.text("Student Details", 30, 97);

  doc.setFont("times", "normal");
  doc.setFontSize(12);
  doc.text(`Name: ${p.name}`, 30, 109);
  doc.text(`USN: ${p.usn}`, 30, 120);
  doc.text(`College: ${p.college}`, 30, 131, { maxWidth: 150 });
  doc.text(`Internship: ${p.internship}`, 30, 142, { maxWidth: 150 });

  entries.forEach((e, i) => {
    doc.addPage();

    centerText("VISVESVARAYA TECHNOLOGICAL UNIVERSITY", 14, 13, true);
    centerText("INTERNSHIP DIARY", 21, 12, true);

    let y = 30;
    const left = 20;
    const maxWidth = 170;

    doc.setFont("times", "normal");
    doc.setFontSize(11);
    doc.text(`Entry: ${i + 1} of ${entries.length}`, left, y);

    y += 10;

    doc.setFont("times", "bold");
    doc.setFontSize(14);
    doc.text(`DATE: ${safeText(e.date) || "-"}`, left, y);

    y += 10;

    doc.setFont("times", "normal");
    doc.setFontSize(12);
    doc.text(`Number of hours: ${safeText(e.hours) || "-"}`, left, y);

    y += 12;

    doc.setFont("times", "bold");
    doc.text("Description:", left, y);

    y += 8;

    doc.setFont("times", "normal");
    const desc = doc.splitTextToSize(safeText(e.activity) || "-", maxWidth);
    doc.text(desc, left, y);

    y += desc.length * 6 + 6;

    doc.setFont("times", "bold");
    doc.text("Learnings/outcomes:", left, y);

    y += 8;

    doc.setFont("times", "normal");
    const learn = doc.splitTextToSize(
      safeText(e.learning || generateLearning(e.activity)) || "-",
      maxWidth
    );
    doc.text(learn, left, y);

    doc.setFont("times", "normal");
    doc.setFontSize(10);
    doc.text("Signature of External Coordinator", 10, 252);
    doc.text("Signature of Internship Coordinator", 200, 252, { align: "right" });

    doc.setFontSize(9);
    doc.text(`Page ${i + 1} of ${entries.length}`, 105, 286, { align: "center" });
  });

  return doc;
}

async function downloadPdf(entries, profile) {
  const doc = drawPdf(entries, profile);
  const pdfDataUrl = doc.output("datauristring");

  await chrome.downloads.download({
    url: pdfDataUrl,
    filename: "Internship_Diary.pdf",
    saveAs: false,
    conflictAction: "uniquify"
  });
}

async function downloadDoc(entries, profile) {
  const html = buildDocHtml(entries, profile);
  const docData = "\ufeff" + html;
  const docDataUrl = "data:application/msword;charset=utf-8," + encodeURIComponent(docData);

  await chrome.downloads.download({
    url: docDataUrl,
    filename: "Internship_Diary.doc",
    saveAs: false,
    conflictAction: "uniquify"
  });
}

function normalizeUploadEntries(rawEntries) {
  const valid = [];
  const rejected = [];

  const sourceEntries = Array.isArray(rawEntries)
    ? rawEntries
    : (rawEntries && typeof rawEntries === "object" ? [rawEntries] : null);

  if (!sourceEntries) {
    return { valid, rejected: [{ index: 0, reason: "Payload must be an object or array" }] };
  }

  sourceEntries.forEach((item, index) => {
    if (!item || typeof item !== "object") {
      rejected.push({ index: index + 1, reason: "Entry must be an object" });
      return;
    }

    const date = normalizeText(item.date || item.Date || "");
    const hours = normalizeText(item.hours || item.Hours || item.hours_worked || "");

    const workDescription = normalizeText(
      item.description ||
      item.activity ||
      item.work_description ||
      item.workDescription ||
      item["Work Description"] ||
      ""
    );
    const learnings = normalizeText(item.learnings || item.Learnings || "");

    const skillsRaw = item.skills || item.Skills || [];
    const skillsList = Array.isArray(skillsRaw)
      ? skillsRaw.map((s) => normalizeText(s)).filter(Boolean)
      : normalizeText(skillsRaw)
        ? [normalizeText(skillsRaw)]
        : [];

    const description = workDescription;

    if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) {
      rejected.push({ index: index + 1, reason: "Date must be YYYY-MM-DD" });
      return;
    }

    const hoursNum = Number(hours);
    if (!Number.isFinite(hoursNum) || hoursNum <= 0 || hoursNum > 24) {
      rejected.push({ index: index + 1, reason: "Hours must be a number between 1 and 24" });
      return;
    }

    if (!description) {
      rejected.push({ index: index + 1, reason: "Description (work summary) is required" });
      return;
    }

    // Fallback learnings text if not provided
    const learningsText = learnings || ("Gained practical understanding of " + description.slice(0, 120));

    valid.push({
      date,
      hours: hoursNum,
      description,
      learnings: learningsText,
      skills: skillsList
    });
  });

  return { valid, rejected };
}

// Runs inside VTU tab to use session cookies, returns both entries and profile.
async function extractDataAndProfile(tabId) {
  const result = await chrome.scripting.executeScript({
    target: { tabId },
    func: async () => {
      const API_URL = "https://vtuapi.internyet.in/api/v1/student/internship-diaries";

      function clean(v) {
        return String(v || "").replace(/\s+/g, " ").trim();
      }

      function setIfEmpty(obj, key, value) {
        let raw = value;
        if (raw && typeof raw === "object") {
          raw = raw.name || raw.title || raw.company_name || raw.organization || raw.value || "";
        }

        const val = clean(raw);
        if (val && !obj[key]) {
          obj[key] = val;
        }
      }

      function extractFromObject(source, profile) {
        if (!source || typeof source !== "object") {
          return;
        }

        const buckets = [
          source,
          source.data,
          source.student,
          source.profile,
          source.user,
          source.internship,
          source.college
        ].filter(Boolean);

        for (const obj of buckets) {
          setIfEmpty(profile, "name", obj.name || obj.student_name || obj.full_name);
          setIfEmpty(profile, "usn", obj.usn || obj.university_seat_number || obj.reg_no || obj.registration_number);
          setIfEmpty(profile, "college", obj.college || obj.college_name || obj.institution || obj.institute);
          setIfEmpty(profile, "internship", obj.internship || obj.company || obj.company_name || obj.organization);
        }
      }

      function parseKeyValueText(text, profile) {
        if (!text) {
          return;
        }

        const t = String(text);
        const nameMatch = t.match(/(?:^|\n)\s*(?:Student\s*Name|Name)\s*[:\-]\s*([^\n]+)/i);
        const usnMatch = t.match(/(?:^|\n)\s*(?:USN|University\s*Seat\s*Number|Reg(?:istration)?\s*No)\s*[:\-]\s*([^\n]+)/i);
        const collegeMatch = t.match(/(?:^|\n)\s*(?:College(?:\s*Name)?|Institution|Institute)\s*[:\-]\s*([^\n]+)/i);
        const internshipMatch = t.match(/(?:^|\n)\s*(?:Internship|Company|Organization)\s*[:\-]\s*([^\n]+)/i);

        setIfEmpty(profile, "name", nameMatch && nameMatch[1]);
        setIfEmpty(profile, "usn", usnMatch && usnMatch[1]);
        setIfEmpty(profile, "college", collegeMatch && collegeMatch[1]);
        setIfEmpty(profile, "internship", internshipMatch && internshipMatch[1]);
      }

      function parseProfileFromDocument(doc, profile) {
        try {
          const rows = Array.from(doc.querySelectorAll("tr"));
          for (const row of rows) {
            const cells = Array.from(row.querySelectorAll("th,td")).map((c) => clean(c.textContent));
            if (cells.length < 2) {
              continue;
            }

            const key = cells[0].toLowerCase();
            const value = cells[1];

            if (/student\s*name|^name$/.test(key)) setIfEmpty(profile, "name", value);
            if (/usn|seat\s*number|reg/.test(key)) setIfEmpty(profile, "usn", value);
            if (/college|institution|institute/.test(key)) setIfEmpty(profile, "college", value);
            if (/internship|company|organization/.test(key)) setIfEmpty(profile, "internship", value);
          }

          parseKeyValueText(doc.body ? doc.body.innerText : "", profile);
        } catch (e) {
          // Ignore parse failures for this page and continue fallbacks.
        }
      }

      const profile = {};
      const all = [];

      let page = 1;
      while (true) {
        const res = await fetch(`${API_URL}?page=${page}`, {
          credentials: "include",
          headers: { Accept: "application/json" }
        });

        if (!res.ok) {
          break;
        }

        const json = await res.json();
        const entries = json.data?.data || json.data || [];

        extractFromObject(json, profile);

        if (!entries.length) {
          break;
        }

        for (const e of entries) {
          extractFromObject(e, profile);
        }

        all.push(...entries);

        if (json.meta?.last_page && page >= json.meta.last_page) {
          break;
        }

        page += 1;
      }

      parseProfileFromDocument(document, profile);

      const profilePages = [
        "/dashboard/student/profile",
        "/dashboard/student/details",
        "/dashboard/student/internship-details",
        "/dashboard/student"
      ];

      for (const path of profilePages) {
        if (profile.name && profile.usn && profile.college && profile.internship) {
          break;
        }

        try {
          const res = await fetch(path, { credentials: "include" });
          if (!res.ok) {
            continue;
          }

          const html = await res.text();
          const doc = new DOMParser().parseFromString(html, "text/html");
          parseProfileFromDocument(doc, profile);
        } catch (e) {
          // Ignore and try next source.
        }
      }

      const profileApis = [
        "https://vtuapi.internyet.in/api/v1/student/profile",
        "https://vtuapi.internyet.in/api/v1/student/me",
        "https://vtuapi.internyet.in/api/v1/student/details"
      ];

      for (const endpoint of profileApis) {
        if (profile.name && profile.usn && profile.college && profile.internship) {
          break;
        }

        try {
          const res = await fetch(endpoint, {
            credentials: "include",
            headers: { Accept: "application/json" }
          });
          if (!res.ok) {
            continue;
          }

          const json = await res.json();
          extractFromObject(json, profile);
        } catch (e) {
          // Ignore and continue fallback chain.
        }
      }

      return { entries: all, profile };
    }
  });

  return result?.[0]?.result || { entries: [], profile: {} };
}

async function startExport(profileInput) {
  try {
    const tab = (await chrome.tabs.query({ active: true, currentWindow: true }))[0];
    if (!tab?.id) {
      send("❌ No active tab");
      return;
    }

    send("📡 Fetching data...");
    const payload = await extractDataAndProfile(tab.id);
    const raw = Array.isArray(payload?.entries) ? payload.entries : [];
    const profile = mergeProfiles(payload?.profile || {}, profileInput || {});

    send(`📊 Found ${raw.length} entries`);

    if (!raw.length) {
      send("❌ No entries (check login)");
      return;
    }

    send(`👤 ${profile.name} | ${profile.usn}`);
    send(`🏫 ${profile.college}`);

    const entries = raw.map((e) => ({
      date: normalizeText(e.date || e.created_at || "-"),
      hours: normalizeText(e.hours || e.hours_worked || "7 hrs"),
      activity: normalizeText(e.description || e.activity || ""),
      learning: generateLearning(e.description || e.activity || "")
    })).sort((a, b) => String(a.date).localeCompare(String(b.date)));

    send("📄 Generating PDF...");
    await downloadPdf(entries, profile);

    send("📝 Generating DOC...");
    await downloadDoc(entries, profile);

    send("✅ Done!");
  } catch (e) {
    send("❌ Error: " + (e?.message || "Unknown error"));
  }
}

chrome.runtime.onMessage.addListener((msg, sender, sendResponse) => {
  if (msg?.type === "log") {
    if (msg?.source === "background") {
      return false;
    }
    void persistLog(msg.text);
    return false;
  }

  if (msg?.type === "get_logs") {
    (async () => {
      const logs = await getStoredLogs();
      sendResponse({ ok: true, logs });
    })().catch((error) => {
      sendResponse({ ok: false, error: error?.message || "Could not fetch logs" });
    });
    return true;
  }

  if (msg?.type === "clear_logs") {
    (async () => {
      await clearPersistedLogs();
      sendResponse({ ok: true });
    })().catch((error) => {
      sendResponse({ ok: false, error: error?.message || "Could not clear logs" });
    });
    return true;
  }

  if (msg?.type === "get_upload_internships") {
    (async () => {
      const resolved = await resolveUploadTab();
      const { tab } = resolved;
      if (!tab || !tab.id) {
        sendResponse({ ok: false, error: "No active tab for internship fetch" });
        return;
      }

      const tabContext = getVtuTabContext(tab.url);
      if (!tabContext.isSupportedHost) {
        sendResponse({ ok: false, error: "Open a VTU tab before loading internships" });
        return;
      }

      const internships = await fetchInternshipsForUpload(tab.id);
      sendResponse({
        ok: true,
        internships,
        tabUrl: String(tab.url || "about:blank"),
        fromFallback: !!resolved.fromFallback
      });
    })().catch((error) => {
      sendResponse({ ok: false, error: error?.message || "Could not fetch internships" });
    });
    return true;
  }

  if (msg?.type === "get_runtime_config") {
    (async () => {
      try {
        const config = await getRuntimeConfig();
        sendResponse({ ok: true, config });
      } catch (e) {
        sendResponse({ ok: false, error: e?.message || 'Could not read config' });
      }
    })().catch((error) => {
      sendResponse({ ok: false, error: error?.message || 'Could not read config' });
    });
    return true;
  }

  if (msg?.type === "set_runtime_config") {
    (async () => {
      try {
        const toStore = msg?.config || {};
        await new Promise((resolve) => chrome.storage.local.set({ [RUNTIME_CONFIG_KEY]: toStore }, resolve));
        sendResponse({ ok: true });
      } catch (e) {
        sendResponse({ ok: false, error: e?.message || 'Could not save config' });
      }
    })().catch((error) => {
      sendResponse({ ok: false, error: error?.message || 'Could not save config' });
    });
    return true;
  }

  if (msg?.type === "get_job_queue") {
    (async () => {
      try {
        const jobs = jobQueue.map((j, idx) => ({ pos: idx + 1, tabId: j.tabId, internshipId: j.internshipId, entries: Array.isArray(j.entries) ? j.entries.length : 0 }));
        const active = Array.from(activeJobStates.entries()).map(([k, v]) => ({ id: k, ...v }));
        sendResponse({ ok: true, queue: jobs, active });
      } catch (e) {
        sendResponse({ ok: false, error: e?.message || 'Could not fetch queue' });
      }
    })().catch((error) => {
      sendResponse({ ok: false, error: error?.message || 'Could not fetch queue' });
    });
    return true;
  }

  if (msg?.type === "get_stats") {
    (async () => {
      try {
        sendResponse({ ok: true, stats: runtimeStats });
      } catch (e) {
        sendResponse({ ok: false, error: e?.message || 'Could not read stats' });
      }
    })().catch((error) => {
      sendResponse({ ok: false, error: error?.message || 'Could not read stats' });
    });
    return true;
  }

  if (msg?.type === "cancel_upload") {
    try {
      cancelUploadFlag = true;
      sendResponse({ ok: true });
    } catch (e) {
      sendResponse({ ok: false, error: e?.message || 'Could not cancel upload' });
    }
    return false;
  }

  return false;
});

chrome.runtime.onMessage.addListener((msg) => {
  if (msg.type === "start") {
    startExport(msg.profileInput);
  }

  if (msg.type === "upload_entries") {
    (async () => {
      try {
        const resolved = await resolveUploadTab();
        const { tab } = resolved;
        if (!tab || !tab.id) {
          send("❌ No active tab for upload");
          return;
        }

        send(`🔎 Active tab URL: ${String(tab.url || "about:blank")}`);
        const tabContext = getVtuTabContext(tab.url);

        if (!tabContext.isSupportedHost) {
          send("❌ Open the VTU diary page in the active tab before uploading");
          return;
        }

        if (resolved.fromFallback) {
          send("ℹ️ Active tab was not VTU. Using the most recent VTU tab for upload.");
        }

        if (!tabContext.isApiHost && !tabContext.isDiaryPath) {
          send("⚠️ Active tab is VTU but not a recognized diary route; attempting upload anyway");
        }

        // Check for deduplication - prevent concurrent uploads of same internship
        const dedupKey = generateDedupKey(msg.internshipId, tab.id);
        if (activeDedupKeys.has(dedupKey)) {
          send(`⚠️ Upload for this internship is already in progress. Please wait for completion.`);
          return;
        }

        // Enqueue the job instead of immediate upload
        send("📦 Enqueueing upload job...");
        jobQueue.push({
          tabId: tab.id,
          entries: msg.data,
          internshipId: msg.internshipId || "",
          options: {
            internshipId: msg.internshipId || "",
            internshipName: msg.internshipName || ""
          }
        });

        send(`📊 Job queue length: ${jobQueue.length}`);

        // Process queue
        await tryDequeueNextJob();
        
        send("✅ Upload job completed");
      } catch (error) {
        send(`❌ Upload setup failed: ${error?.message || "Unknown error"}`);
      }
    })().catch((error) => {
      send(`❌ Upload error: ${error?.message || "Unknown error"}`);
    });
  }
});
