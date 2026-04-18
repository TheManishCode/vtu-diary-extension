const logEl = document.getElementById("log");
const statusPillEl = document.getElementById("statusPill");
const nameEl = document.getElementById("name");
const usnEl = document.getElementById("usn");
const collegeEl = document.getElementById("college");
const internshipEl = document.getElementById("internship");
const jsonFileEl = document.getElementById("jsonFile");
const generateBtn = document.getElementById("generate");
const uploadBtn = document.getElementById("upload");
const saveProfileBtn = document.getElementById("saveProfile");
const clearFileBtn = document.getElementById("clearFile");
const clearLogsBtn = document.getElementById("clearLogs");
const downloadExampleBtn = document.getElementById("downloadExample");
const openPersistentBtn = document.getElementById("openPersistent");

const PROFILE_KEY = "vtu_profile_overrides";
const LOG_STORAGE_KEY = "vtu_runtime_logs";
const LOCAL_LOGS_KEY = "vtu_popup_logs_local";
const MAX_LOCAL_LOGS = 500;

let isBusy = false;

function now() {
  const d = new Date();
  return d.toLocaleTimeString([], { hour: "2-digit", minute: "2-digit", second: "2-digit" });
}

function detectLevel(msg) {
  if (/❌|error|failed/i.test(msg)) return "error";
  if (/⚠️|warn|skipped/i.test(msg)) return "warn";
  if (/✅|done|completed|uploaded/i.test(msg)) return "success";
  return "info";
}

function setBusy(busy, label) {
  isBusy = busy;
  generateBtn.disabled = busy;
  uploadBtn.disabled = busy;
  saveProfileBtn.disabled = busy;
  clearFileBtn.disabled = busy;
  if (downloadExampleBtn) {
    downloadExampleBtn.disabled = busy;
  }
  if (openPersistentBtn) {
    openPersistentBtn.disabled = busy;
  }

  statusPillEl.classList.toggle("busy", busy);
  statusPillEl.textContent = label || (busy ? "Working" : "Idle");

  // Show informational message when starting long operation
  if (busy && label && label.includes("Uploading")) {
    appendLog(
      "ℹ️ Operation running in background. You can safely switch tabs or close this popup—logs will resume here when ready.",
      "info"
    );
  }
}

function appendLog(msg, levelHint, timestampMs) {
  const cleanMsg = sanitizeLogMessage(msg);
  const level = levelHint || detectLevel(cleanMsg);
  const resolvedTs = Number(timestampMs || Date.now());

  const item = document.createElement("div");
  item.className = `log-item log-${level}`;

  const meta = document.createElement("span");
  meta.className = "log-meta";
  meta.textContent = new Date(resolvedTs).toLocaleTimeString([], { hour: "2-digit", minute: "2-digit", second: "2-digit" });

  const body = document.createElement("span");
  body.textContent = cleanMsg;

  item.appendChild(meta);
  item.appendChild(body);
  logEl.appendChild(item);
  logEl.scrollTop = logEl.scrollHeight;

  cacheLocalLog(cleanMsg, resolvedTs);
}

function readLocalLogs() {
  try {
    const raw = localStorage.getItem(LOCAL_LOGS_KEY);
    if (!raw) {
      return [];
    }
    const parsed = JSON.parse(raw);
    return Array.isArray(parsed) ? parsed : [];
  } catch (_) {
    return [];
  }
}

function writeLocalLogs(logs) {
  try {
    localStorage.setItem(LOCAL_LOGS_KEY, JSON.stringify(logs.slice(-MAX_LOCAL_LOGS)));
  } catch (_) {
    // Ignore local log write failures.
  }
}

function cacheLocalLog(text, ts) {
  const entryText = String(text || "").trim();
  if (!entryText) {
    return;
  }
  const current = readLocalLogs();
  current.push({ text: entryText, ts: Number(ts || Date.now()) });
  writeLocalLogs(current);
}

function clearLocalLogs() {
  try {
    localStorage.removeItem(LOCAL_LOGS_KEY);
  } catch (_) {
    // Ignore local clear failures.
  }
}

function sendRuntimeMessage(payload) {
  return new Promise((resolve, reject) => {
    chrome.runtime.sendMessage(payload, (response) => {
      if (chrome.runtime.lastError) {
        reject(new Error(chrome.runtime.lastError.message || "Runtime message failed"));
        return;
      }
      resolve(response);
    });
  });
}

function storageGet(key) {
  return new Promise((resolve, reject) => {
    chrome.storage.local.get([key], (result) => {
      if (chrome.runtime.lastError) {
        reject(new Error(chrome.runtime.lastError.message || "Storage get failed"));
        return;
      }
      resolve(result?.[key]);
    });
  });
}

function storageSet(obj) {
  return new Promise((resolve, reject) => {
    chrome.storage.local.set(obj, () => {
      if (chrome.runtime.lastError) {
        reject(new Error(chrome.runtime.lastError.message || "Storage set failed"));
        return;
      }
      resolve();
    });
  });
}

function openPersistentWorkspaceTab() {
  return new Promise((resolve, reject) => {
    const url = chrome.runtime.getURL("popup.html") + "?mode=persistent";
    chrome.tabs.create({ url, active: true }, (tab) => {
      if (chrome.runtime.lastError) {
        reject(new Error(chrome.runtime.lastError.message || "Could not open persistent workspace tab"));
        return;
      }
      resolve(tab);
    });
  });
}

function sanitizeLogMessage(msg) {
  let text = String(msg || "").trim();

  if (!text) {
    return "";
  }

  text = text.replace(/https?:\/\/\S+/gi, "[secure endpoint]");
  text = text.replace(/\b(?:GET|POST|PUT|PATCH|DELETE)\s+\[secure endpoint\]/gi, "API request");

  if (/Active tab URL:/i.test(text)) {
    return "Session check complete. VTU tab detected.";
  }
  if (/Upload method:/i.test(text)) {
    return "Upload service selected automatically.";
  }
  if (/Payload:/i.test(text)) {
    return "Preparing secure request payload.";
  }
  if (/Skills available:/i.test(text)) {
    return "Skills catalog fetched successfully.";
  }
  if (/internship-applys/i.test(text)) {
    return "Internship context resolved.";
  }
  if (/Internship ID:/i.test(text)) {
    return "Internship mapping resolved.";
  }
  if (/Skill IDs:/i.test(text)) {
    return "Skill mapping resolved.";
  }
  if (/Response:\s*\d+/i.test(text)) {
    return "Server response received.";
  }
  if (/POSTing to/i.test(text)) {
    return "Sending request to upload service.";
  }

  return text;
}

function appendStartupGuidanceLogs() {
  appendLog("Quick Start: 1) Keep VTU session logged in 2) Select JSON 3) Upload Entries.", "note");
  appendLog("Persistent mode: use Open Persistent Workspace to continue in a non-closing extension tab.", "note");
}

async function restorePersistedLogs() {
  try {
    let entries = [];

    try {
      const response = await sendRuntimeMessage({ type: "get_logs" });
      entries = Array.isArray(response?.logs) ? response.logs : [];
    } catch (_) {
      try {
        const storageLogs = await storageGet(LOG_STORAGE_KEY);
        entries = Array.isArray(storageLogs) ? storageLogs : [];
      } catch (_) {
        entries = [];
      }
    }

    if (!entries.length) {
      entries = readLocalLogs();
    }

    if (!entries.length) {
      return 0;
    }

    logEl.textContent = "";
    entries.forEach((entry) => {
      appendLog(entry.text || "", null, entry.ts || Date.now());
    });
    return entries.length;
  } catch (error) {
    appendLog("Could not restore previous logs. Starting a fresh session log.", "warn");
    return 0;
  }
}

function saveProfile() {
  const profileInput = {
    name: nameEl.value.trim(),
    usn: usnEl.value.trim(),
    college: collegeEl.value.trim(),
    internship: internshipEl.value.trim()
  };

  localStorage.setItem(PROFILE_KEY, JSON.stringify(profileInput));
  appendLog("Profile overrides saved", "success");
}

function loadProfile() {
  try {
    const raw = localStorage.getItem(PROFILE_KEY);
    if (!raw) {
      return;
    }

    const data = JSON.parse(raw);
    nameEl.value = data.name || "";
    usnEl.value = data.usn || "";
    collegeEl.value = data.college || "";
    internshipEl.value = data.internship || "";
  } catch (error) {
    appendLog("Could not load saved profile", "warn");
  }
}

function profileInput() {
  return {
    name: nameEl.value,
    usn: usnEl.value,
    college: collegeEl.value,
    internship: internshipEl.value
  };
}

function finishIfTerminal(msg) {
  if (/✅ Done!|✅ Upload flow completed|❌/i.test(msg)) {
    setBusy(false, /❌/i.test(msg) ? "Needs Attention" : "Idle");
  }
}

generateBtn.onclick = () => {
  if (isBusy) {
    return;
  }

  saveProfile();
  appendLog("Starting export...", "info");
  setBusy(true, "Exporting");
  chrome.runtime.sendMessage({ type: "start", profileInput: profileInput() });
};

uploadBtn.onclick = async () => {
  if (isBusy) {
    return;
  }

  setBusy(true, "Preparing Upload");

  try {
    const file = jsonFileEl.files && jsonFileEl.files[0];
    if (!file) {
      appendLog("Select a JSON file first", "error");
      setBusy(false, "Needs Attention");
      return;
    }

    const text = await file.text();
    const data = JSON.parse(text);

    const normalized = Array.isArray(data) ? data : (data && typeof data === "object" ? [data] : null);

    if (!normalized) {
      appendLog("JSON must be an entry object or an array of entry objects", "error");
      setBusy(false, "Needs Attention");
      return;
    }

    appendLog(`Uploading ${normalized.length} entries...`, "info");
    setBusy(true, "Uploading");
    chrome.runtime.sendMessage({ type: "upload_entries", data: normalized });
  } catch (error) {
    appendLog("Invalid JSON: " + (error?.message || "Parse failed"), "error");
    setBusy(false, "Needs Attention");
  }
};

saveProfileBtn.onclick = () => {
  saveProfile();
};

clearFileBtn.onclick = () => {
  jsonFileEl.value = "";
  appendLog("File selection cleared", "info");
};

clearLogsBtn.onclick = () => {
  void sendRuntimeMessage({ type: "clear_logs" }).catch(async () => {
    try {
      await storageSet({ [LOG_STORAGE_KEY]: [] });
    } catch (_) {
      // Ignore fallback clear failures in UI.
    }
  });
  clearLocalLogs();
  logEl.textContent = "";
  appendLog("Log cleared", "info");
  appendStartupGuidanceLogs();
};

if (openPersistentBtn) {
  openPersistentBtn.onclick = async () => {
    if (isBusy) {
      return;
    }

    try {
      await openPersistentWorkspaceTab();
      appendLog("Persistent workspace opened in a non-closing tab.", "success");
    } catch (error) {
      const detail = error && error.message ? ` (${error.message})` : "";
      appendLog(`Could not open persistent workspace${detail}.`, "error");
    }
  };
}

if (downloadExampleBtn) {
  downloadExampleBtn.onclick = async () => {
    if (isBusy) {
      return;
    }

    try {
      const url = chrome.runtime.getURL("examples/2026-04-14.json");
      await chrome.downloads.download({
        url,
        filename: "VTU_Upload_Example_2026-04-14.json",
        saveAs: true,
        conflictAction: "uniquify"
      });
      appendLog("Example JSON downloaded successfully", "success");
    } catch (error) {
      appendLog("Could not download example file", "error");
    }
  };
}

chrome.runtime.onMessage.addListener((msg) => {
  if (msg.type === "log") {
    appendLog(msg.text);
    finishIfTerminal(msg.text);
  }
});

loadProfile();
setBusy(false, "Idle");

(async () => {
  const restored = await restorePersistedLogs();
  if (!restored) {
    appendLog("Ready", "success");
    appendStartupGuidanceLogs();
  } else {
    appendLog(`Restored ${restored} previous log entries`, "info");
    appendStartupGuidanceLogs();
  }
})();
