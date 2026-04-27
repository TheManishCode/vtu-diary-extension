document.addEventListener("DOMContentLoaded", () => {
  const logEl = document.getElementById("log");
  const logCountEl = document.getElementById("logCount");
  const statusPillEl = document.getElementById("statusPill");
  const statusTxtEl = document.getElementById("statusTxt");
  const statusLedEl = document.getElementById("led");
  const tabs = Array.from(document.querySelectorAll(".tab[data-view]"));
  const views = Array.from(document.querySelectorAll(".view"));
  const tabInkEl = document.getElementById("ink");
  const logBadgeEl = document.getElementById("logBadge");
  const logsViewEl = document.getElementById("view-lg");
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
  const affiliateCardEl = document.getElementById("affiliateCard");

  const PROFILE_KEY = "vtu_profile_overrides";
  const LOG_STORAGE_KEY = "vtu_runtime_logs";
  const LOCAL_LOGS_KEY = "vtu_popup_logs_local";
  const AFFILIATE_UNLOCK_KEY = "vtu_affiliate_unlocked";
  const MAX_LOCAL_LOGS = 500;

  let isBusy = false;

  function moveInk(tabEl) {
    if (!tabInkEl || !tabEl) {
      return;
    }
    tabInkEl.style.left = tabEl.offsetLeft + "px";
    tabInkEl.style.width = tabEl.offsetWidth + "px";
  }

  function isLogsViewActive() {
    return !!logsViewEl?.classList.contains("on");
  }

  function clearLogBadge() {
    if (!logBadgeEl) {
      return;
    }
    logBadgeEl.style.display = "none";
  }

  function switchTab(tabEl) {
    if (!tabEl) {
      return;
    }

    tabs.forEach((tab) => tab.classList.remove("on"));
    views.forEach((view) => view.classList.remove("on"));

    tabEl.classList.add("on");

    const nextView = document.getElementById("view-" + tabEl.dataset.view);
    if (nextView) {
      nextView.classList.add("on");
    }

    moveInk(tabEl);

    if (tabEl.dataset.view === "lg") {
      clearLogBadge();
    }
  }

  function initializeTabs() {
    if (!tabs.length) {
      return;
    }

    tabs.forEach((tab) => {
      tab.addEventListener("click", () => switchTab(tab));
    });

    const activeTab = tabs.find((tab) => tab.classList.contains("on")) || tabs[0];
    requestAnimationFrame(() => moveInk(activeTab));
    window.addEventListener("resize", () => moveInk(tabs.find((tab) => tab.classList.contains("on")) || activeTab));
  }

  function detectLevel(msg) {
    if (/❌|error|failed/i.test(msg)) return "error";
    if (/⚠️|warn|skipped/i.test(msg)) return "warn";
    if (/✅|done|completed|uploaded/i.test(msg)) return "success";
    return "info";
  }

  function updateStatus(label, busy) {
    const statusLabel = label || (busy ? "Working" : "Idle");
    const state = busy ? "busy" : (statusLabel === "Needs Attention" ? "err" : "ok");

    if (statusPillEl) {
      statusPillEl.classList.toggle("busy", !!busy);
    }
    if (statusTxtEl) {
      statusTxtEl.textContent = statusLabel;
    }
    if (statusLedEl) {
      statusLedEl.className = "led" + (state ? " " + state : "");
    }
  }

  function setBusy(busy, label) {
    isBusy = !!busy;

    if (generateBtn) generateBtn.disabled = isBusy;
    if (uploadBtn) uploadBtn.disabled = isBusy;
    if (saveProfileBtn) saveProfileBtn.disabled = isBusy;
    if (clearFileBtn) clearFileBtn.disabled = isBusy;
    if (downloadExampleBtn) downloadExampleBtn.disabled = isBusy;

    updateStatus(label, isBusy);

    if (isBusy && label && label.includes("Uploading")) {
      appendLog(
        "ℹ️ Operation running in background. You can safely switch tabs or close this popup—logs will resume here when ready.",
        "info"
      );
    }
  }

  function updateLogCounter() {
    if (!logEl || !logCountEl) {
      return;
    }
    const count = logEl.querySelectorAll(".li").length;
    logCountEl.textContent = count + (count === 1 ? " entry" : " entries");

    if (!logBadgeEl) {
      return;
    }

    if (count <= 0 || isLogsViewActive()) {
      clearLogBadge();
      return;
    }

    logBadgeEl.style.display = "inline";
    logBadgeEl.textContent = String(count);
  }

  function appendLog(msg, levelHint, timestampMs, options) {
    const persist = options?.persist !== false;
    const cleanMsg = sanitizeLogMessage(msg);
    if (!cleanMsg) {
      return;
    }

    const level = levelHint || detectLevel(cleanMsg);
    const resolvedTs = Number(timestampMs || Date.now());

    if (logEl) {
      const empty = logEl.querySelector(".log-empty");
      if (empty) {
        logEl.innerHTML = "";
      }

      const item = document.createElement("div");
      const levelClassMap = {
        info: "li-info",
        success: "li-ok",
        warn: "li-warn",
        error: "li-err",
        note: ""
      };
      item.className = `li ${levelClassMap[level] || ""}`.trim();

      const bar = document.createElement("div");
      bar.className = "li-bar";

      const body = document.createElement("div");
      body.className = "li-body";

      const time = document.createElement("span");
      time.className = "li-time";
      time.textContent = new Date(resolvedTs).toLocaleTimeString([], {
        hour: "2-digit",
        minute: "2-digit",
        second: "2-digit"
      });

      const text = document.createElement("div");
      text.className = "li-msg";
      text.textContent = cleanMsg;

      body.appendChild(time);
      body.appendChild(text);
      item.appendChild(bar);
      item.appendChild(body);
      logEl.appendChild(item);
      logEl.scrollTop = logEl.scrollHeight;

      updateLogCounter();
    }

    if (persist) {
      cacheLocalLog(cleanMsg, resolvedTs);
    }
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
    const entryTs = Number(ts || Date.now());
    const last = current[current.length - 1];
    if (last && last.text === entryText && Number(last.ts) === entryTs) {
      return;
    }
    current.push({ text: entryText, ts: entryTs });
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

      if (logEl) {
        logEl.innerHTML = "";
      }

      entries.forEach((entry) => {
        appendLog(entry.text || "", null, entry.ts || Date.now(), { persist: false });
      });
      return entries.length;
    } catch (_) {
      appendLog("Could not restore previous logs. Starting a fresh session log.", "warn");
      return 0;
    }
  }

  function saveProfile() {
    if (!nameEl || !usnEl || !collegeEl || !internshipEl) {
      return;
    }

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
    if (!nameEl || !usnEl || !collegeEl || !internshipEl) {
      return;
    }

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
    } catch (_) {
      appendLog("Could not load saved profile", "warn");
    }
  }

  function profileInput() {
    return {
      name: nameEl?.value || "",
      usn: usnEl?.value || "",
      college: collegeEl?.value || "",
      internship: internshipEl?.value || ""
    };
  }

  function showAffiliateCard() {
    if (!affiliateCardEl) {
      return;
    }

    affiliateCardEl.classList.remove("hidden");
    affiliateCardEl.classList.add("show");
    localStorage.setItem(AFFILIATE_UNLOCK_KEY, "1");
  }

  function restoreAffiliateCardState() {
    try {
      const unlocked = localStorage.getItem(AFFILIATE_UNLOCK_KEY) === "1";
      if (unlocked) {
        showAffiliateCard();
      }
    } catch (_) {
      // Ignore localStorage failures.
    }
  }

  function finishIfTerminal(msg) {
    if (/✅ Done!/i.test(msg)) {
      showAffiliateCard();
    }

    if (/✅ Done!|✅ Upload flow completed|❌/i.test(msg)) {
      setBusy(false, /❌/i.test(msg) ? "Needs Attention" : "Idle");
    }
  }

  if (generateBtn) {
    generateBtn.onclick = () => {
      if (isBusy) {
        return;
      }

      saveProfile();
      appendLog("Starting export...", "info");
      setBusy(true, "Exporting");
      chrome.runtime.sendMessage({ type: "start", profileInput: profileInput() });
    };
  }

  if (uploadBtn) {
    uploadBtn.onclick = async () => {
      if (isBusy) {
        return;
      }

      setBusy(true, "Preparing Upload");

      try {
        const file = jsonFileEl?.files && jsonFileEl.files[0];
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
  }

  if (saveProfileBtn) {
    saveProfileBtn.onclick = () => {
      saveProfile();
    };
  }

  if (clearFileBtn) {
    clearFileBtn.onclick = () => {
      if (jsonFileEl) {
        jsonFileEl.value = "";
      }
      appendLog("File selection cleared", "info");
    };
  }

  if (clearLogsBtn) {
    clearLogsBtn.onclick = () => {
      void sendRuntimeMessage({ type: "clear_logs" }).catch(async () => {
        try {
          await storageSet({ [LOG_STORAGE_KEY]: [] });
        } catch (_) {
          // Ignore fallback clear failures in UI.
        }
      });

      clearLocalLogs();

      if (logEl) {
        logEl.innerHTML = `
      <div class="log-empty">
        <div class="log-empty-icon">◎</div>
        No activity yet
      </div>`;
      }

      updateLogCounter();
      clearLogBadge();
      appendLog("Log cleared", "info", Date.now(), { persist: false });
      appendLog("Quick Start: 1) Keep VTU session logged in 2) Select JSON 3) Upload Entries.", "note", Date.now(), { persist: false });
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
      } catch (_) {
        appendLog("Could not download example file", "error");
      }
    };
  }

  chrome.runtime.onMessage.addListener((msg) => {
    if (msg?.type === "log") {
      appendLog(msg.text);
      finishIfTerminal(msg.text || "");
    }
  });

  loadProfile();
  initializeTabs();
  setBusy(false, "Idle");
  restoreAffiliateCardState();

  (async () => {
    const restored = await restorePersistedLogs();
    if (!restored) {
      appendLog("Ready", "success", Date.now(), { persist: false });
      appendLog("Quick Start: 1) Keep VTU session logged in 2) Select JSON 3) Upload Entries.", "note", Date.now(), { persist: false });
    } else {
      appendLog(`Restored ${restored} previous log entries`, "info", Date.now(), { persist: false });
      appendLog("Quick Start: 1) Keep VTU session logged in 2) Select JSON 3) Upload Entries.", "note", Date.now(), { persist: false });
    }
  })();
});
