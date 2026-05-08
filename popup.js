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
  const internshipIdEl = document.getElementById("internshipId");
  const internshipSelectEl = document.getElementById("internshipSelect");
  const refreshInternshipsBtn = document.getElementById("refreshInternships");
  const jsonFileEl = document.getElementById("jsonFile");
  const fileDropZoneEl = document.getElementById("fileDropZone");
  const fileSelectedEl = document.getElementById("fileSelected");
  const fileSelectedNameEl = document.getElementById("fileSelectedName");
  const fileSelectedClearEl = document.getElementById("fileSelectedClear");
  const jsonPreviewEl = document.getElementById("jsonPreview");
  const uploadProgressEl = document.getElementById("uploadProgress");
  const progressBarFillEl = document.getElementById("progressBarFill");
  const progressTextEl = document.getElementById("progressText");
  const sessionBadgeEl = document.getElementById("sessionBadge");
  const generateBtn = document.getElementById("generate");
  const uploadBtn = document.getElementById("upload");
  const saveProfileBtn = document.getElementById("saveProfile");
  const clearFileBtn = document.getElementById("clearFile");
  const clearLogsBtn = document.getElementById("clearLogs");
  const downloadExampleBtn = document.getElementById("downloadExample");

  const PROFILE_KEY = "vtu_profile_overrides";
  const LOG_STORAGE_KEY = "vtu_runtime_logs";
  const LOCAL_LOGS_KEY = "vtu_popup_logs_local";
  const MAX_LOCAL_LOGS = 500;
  const INTERNSHIP_OPTIONS_KEY = "vtu_upload_internship_options";

  let isBusy = false;
  let internshipFetchInProgress = false;
  let currentUploadTotal = 0;
  let currentUploadCount = 0;

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

    if (tabEl.dataset.view === "up") {
      const hasLoadedOptions = !!(internshipSelectEl && internshipSelectEl.options.length > 1);
      if (!hasLoadedOptions && !internshipFetchInProgress) {
        void refreshInternshipOptions({ silent: true });
      }
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

    // Show/hide cancel button based on upload state
    const cancelBtn = document.getElementById('cancelUpload');
    if (isBusy && label && label.includes("Uploading")) {
      if (uploadBtn) uploadBtn.style.display = 'none';
      if (cancelBtn) cancelBtn.style.display = 'block';
    } else {
      if (uploadBtn) uploadBtn.style.display = 'block';
      if (cancelBtn) cancelBtn.style.display = 'none';
    }

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
    if (/internship_id still not found/i.test(text) || /will attempt upload anyway/i.test(text)) {
      return "Internship ID could not be resolved. Upload was stopped.";
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
    if (!nameEl || !usnEl || !collegeEl || !internshipEl || !internshipIdEl) {
      return;
    }

    const profileInput = {
      name: nameEl.value.trim(),
      usn: usnEl.value.trim(),
      college: collegeEl.value.trim(),
      internship: internshipEl.value.trim(),
      internshipId: internshipIdEl.value.trim(),
      selectedInternshipId: internshipSelectEl?.value || ""
    };

    localStorage.setItem(PROFILE_KEY, JSON.stringify(profileInput));
    appendLog("Profile overrides saved", "success");
  }

  function loadProfile() {
    if (!nameEl || !usnEl || !collegeEl || !internshipEl || !internshipIdEl) {
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
      internshipIdEl.value = data.internshipId || "";
      if (internshipSelectEl && data.selectedInternshipId) {
        internshipSelectEl.dataset.pendingSelectedId = String(data.selectedInternshipId);
      }
    } catch (_) {
      appendLog("Could not load saved profile", "warn");
    }
  }

  function getSelectedInternship() {
    if (!internshipSelectEl) {
      return { id: "", name: "" };
    }

    const selectedId = internshipSelectEl.value || "";
    if (!selectedId) {
      return { id: "", name: "" };
    }

    const option = internshipSelectEl.options[internshipSelectEl.selectedIndex];
    return {
      id: selectedId,
      name: option?.textContent?.trim() || ""
    };
  }

  function readCachedInternshipOptions() {
    try {
      const raw = localStorage.getItem(INTERNSHIP_OPTIONS_KEY);
      if (!raw) {
        return [];
      }
      const parsed = JSON.parse(raw);
      if (!Array.isArray(parsed)) {
        return [];
      }
      return parsed
        .filter((item) => item && /^\d+$/.test(String(item.id || "")))
        .map((item) => ({ id: String(item.id), name: String(item.name || `Internship ${item.id}`) }));
    } catch (_) {
      return [];
    }
  }

  function writeCachedInternshipOptions(options) {
    try {
      localStorage.setItem(INTERNSHIP_OPTIONS_KEY, JSON.stringify(options));
    } catch (_) {
      // Ignore cache write failures.
    }
  }

  function renderInternshipOptions(options) {
    if (!internshipSelectEl) {
      return;
    }

    const list = Array.isArray(options) ? options : [];
    const pendingFromProfile = internshipSelectEl.dataset.pendingSelectedId || "";
    const currentSelection = internshipSelectEl.value || pendingFromProfile || "";

    internshipSelectEl.innerHTML = "";

    const placeholder = document.createElement("option");
    placeholder.value = "";
    placeholder.textContent = list.length
      ? "Select enrolled internship"
      : "No internships loaded";
    internshipSelectEl.appendChild(placeholder);

    list.forEach((item) => {
      const option = document.createElement("option");
      option.value = String(item.id);
      option.textContent = String(item.name || `Internship ${item.id}`);
      internshipSelectEl.appendChild(option);
    });

    if (currentSelection && list.some((item) => String(item.id) === String(currentSelection))) {
      internshipSelectEl.value = String(currentSelection);
    } else if (list.length === 1) {
      internshipSelectEl.value = String(list[0].id);
    }

    delete internshipSelectEl.dataset.pendingSelectedId;
    internshipSelectEl.disabled = !list.length;
  }

  function setInternshipSelectLoading() {
    if (!internshipSelectEl) {
      return;
    }

    internshipSelectEl.innerHTML = "";
    const loading = document.createElement("option");
    loading.value = "";
    loading.textContent = "Loading internships...";
    internshipSelectEl.appendChild(loading);
    internshipSelectEl.disabled = true;
  }

  // NEW HELPER FUNCTIONS

  function updateSessionBadge() {
    if (!sessionBadgeEl) {
      return;
    }

    chrome.tabs.query({ url: "*://internship.vtu.edu.in/*" }, (tabs) => {
      const isActive = tabs && tabs.length > 0;
      if (isActive) {
        sessionBadgeEl.classList.add("active");
        sessionBadgeEl.textContent = "VTU Active";
      } else {
        sessionBadgeEl.classList.remove("active");
        sessionBadgeEl.textContent = "VTU Not Active";
      }
    });
  }

  function validateJsonEntry(entry) {
    if (!entry || typeof entry !== "object") {
      return false;
    }

    const date = String(entry.date || entry.Date || "").trim();
    const hours = Number(entry.hours || entry.Hours || entry.hours_worked || "");
    const description = String(
      entry.description ||
      entry.activity ||
      entry.work_description ||
      entry.workDescription ||
      entry["Work Description"] ||
      ""
    ).trim();

    return /^\d{4}-\d{2}-\d{2}$/.test(date) && Number.isFinite(hours) && hours > 0 && hours <= 24 && !!description;
  }

  function showJsonPreview(data) {
    if (!jsonPreviewEl) {
      return;
    }

    try {
      const normalized = Array.isArray(data) ? data : (data && typeof data === "object" ? [data] : []);
      
      if (!normalized.length) {
        jsonPreviewEl.innerHTML = '<div class="preview-message">Invalid JSON structure</div>';
        return;
      }

      const validEntries = normalized.filter((entry) => validateJsonEntry(entry));
      const invalidCount = normalized.length - validEntries.length;

      let previewHtml = `<div class="preview-header">Preview: ${validEntries.length} valid entries`;
      if (invalidCount > 0) {
        previewHtml += ` <span class="preview-warning">(${invalidCount} invalid)</span>`;
      }
      previewHtml += "</div>";

      previewHtml += '<div class="preview-list">';
      validEntries.slice(0, 5).forEach((entry, index) => {
        const description = String(
          entry.description ||
          entry.activity ||
          entry.work_description ||
          entry.workDescription ||
          entry["Work Description"] ||
          "Untitled"
        ).trim();
        previewHtml += `
          <div class="preview-item">
            <span class="preview-index">${index + 1}.</span>
            <span class="preview-title">${description || "Untitled"}</span>
            <span class="preview-date">${String(entry.date || entry.Date || "No date")}</span>
          </div>
        `;
      });

      if (validEntries.length > 5) {
        previewHtml += `<div class="preview-more">... and ${validEntries.length - 5} more entries</div>`;
      }

      previewHtml += "</div>";
      jsonPreviewEl.innerHTML = previewHtml;
    } catch (error) {
      jsonPreviewEl.innerHTML = '<div class="preview-message error">Failed to parse preview</div>';
    }
  }

  function handleFileSelect(file) {
    if (!file) {
      return;
    }

    if (!file.name.endsWith(".json")) {
      appendLog("Please select a valid JSON file", "error");
      return;
    }

    const reader = new FileReader();
    reader.onload = (event) => {
      try {
        const data = JSON.parse(event.target.result);
        
        if (fileSelectedEl && fileSelectedNameEl) {
          fileSelectedEl.style.display = "block";
          fileSelectedNameEl.textContent = file.name;
        }

        showJsonPreview(data);
        appendLog(`JSON file loaded: ${file.name}`, "success");
      } catch (error) {
        appendLog("Invalid JSON file: " + (error?.message || "Parse failed"), "error");
        if (fileSelectedEl) {
          fileSelectedEl.style.display = "none";
        }
        if (jsonPreviewEl) {
          jsonPreviewEl.innerHTML = "";
        }
      }
    };
    reader.readAsText(file);
  }

  function setupDragAndDrop() {
    if (!fileDropZoneEl) {
      return;
    }

    fileDropZoneEl.addEventListener("dragover", (e) => {
      e.preventDefault();
      e.stopPropagation();
      fileDropZoneEl.classList.add("dragover");
    });

    fileDropZoneEl.addEventListener("dragleave", (e) => {
      e.preventDefault();
      e.stopPropagation();
      fileDropZoneEl.classList.remove("dragover");
    });

    fileDropZoneEl.addEventListener("drop", (e) => {
      e.preventDefault();
      e.stopPropagation();
      fileDropZoneEl.classList.remove("dragover");

      const files = e.dataTransfer?.files;
      if (files && files[0]) {
        if (jsonFileEl) {
          const dataTransfer = new DataTransfer();
          dataTransfer.items.add(files[0]);
          jsonFileEl.files = dataTransfer.files;
        }
        handleFileSelect(files[0]);
      }
    });
  }

  function updateUploadProgress(current, total) {
    currentUploadCount = current;
    currentUploadTotal = total;

    if (!uploadProgressEl || !progressBarFillEl || !progressTextEl) {
      return;
    }

    const percentage = total > 0 ? (current / total) * 100 : 0;
    progressBarFillEl.style.width = percentage + "%";
    progressTextEl.textContent = `${current}/${total}`;

    if (percentage === 100) {
      uploadProgressEl.classList.add("complete");
    } else {
      uploadProgressEl.classList.remove("complete");
    }
  }

  async function refreshInternshipOptions(options = {}) {
    if (!refreshInternshipsBtn) {
      return;
    }

    const silent = options?.silent === true;
    if (internshipFetchInProgress) {
      return;
    }

    internshipFetchInProgress = true;
    setInternshipSelectLoading();

    refreshInternshipsBtn.disabled = true;
    if (!silent) {
      appendLog("Fetching enrolled internships...", "info");
    }

    try {
      const response = await sendRuntimeMessage({ type: "get_upload_internships" });
      if (!response?.ok) {
        throw new Error(response?.error || "Could not load internships");
      }

      const internships = Array.isArray(response.internships) ? response.internships : [];
      if (!internships.length) {
        renderInternshipOptions([]);
        if (!silent) {
          appendLog("No enrolled internships found. Use manual internship ID fallback.", "warn");
        }
        return;
      }

      const compact = internships.map((item) => ({
        id: String(item.id),
        name: String(item.name || `Internship ${item.id}`)
      }));

      renderInternshipOptions(compact);
      writeCachedInternshipOptions(compact);
      if (!silent) {
        appendLog(`Loaded ${compact.length} enrolled internship(s)`, "success");
      }
    } catch (error) {
      const cached = readCachedInternshipOptions();
      renderInternshipOptions(cached);
      if (!silent) {
        appendLog("Could not fetch internships from VTU: " + (error?.message || "Unknown error"), "warn");
      }
    } finally {
      internshipFetchInProgress = false;
      refreshInternshipsBtn.disabled = false;
    }
  }

  function profileInput() {
    return {
      name: nameEl?.value || "",
      usn: usnEl?.value || "",
      college: collegeEl?.value || "",
      internship: internshipEl?.value || "",
      internshipId: internshipIdEl?.value || ""
    };
  }

  function finishIfTerminal(msg) {
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

        const selected = getSelectedInternship();
        const manualInternshipId = internshipIdEl?.value.trim() || "";
        const internshipIdForUpload = selected.id || manualInternshipId;

        if (!internshipIdForUpload) {
          appendLog("Select an internship from dropdown (or use manual internship ID).", "error");
          setBusy(false, "Needs Attention");
          return;
        }

        chrome.runtime.sendMessage({
          type: "upload_entries",
          data: normalized,
          internshipId: internshipIdForUpload,
          internshipName: selected.name || ""
        });
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

  if (refreshInternshipsBtn) {
    refreshInternshipsBtn.onclick = () => {
      void refreshInternshipOptions({ silent: false });
    };
  }

  if (clearFileBtn) {
    clearFileBtn.onclick = () => {
      if (jsonFileEl) {
        jsonFileEl.value = "";
      }
      if (fileSelectedEl) {
        fileSelectedEl.style.display = "none";
      }
      if (jsonPreviewEl) {
        jsonPreviewEl.innerHTML = "";
      }
      currentUploadTotal = 0;
      currentUploadCount = 0;
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
        const url = chrome.runtime.getURL("examples/entries.json");
        await chrome.downloads.download({
          url,
          filename: "VTU_Upload_Example.json",
          saveAs: true,
          conflictAction: "uniquify"
        });
        appendLog("Example JSON downloaded successfully", "success");
      } catch (_) {
        appendLog("Could not download example file", "error");
      }
    };
  }

  // EVENT LISTENER FOR CLEARING SELECTED FILE
  if (fileSelectedClearEl) {
    fileSelectedClearEl.onclick = () => {
      if (jsonFileEl) {
        jsonFileEl.value = "";
      }
      if (fileSelectedEl) {
        fileSelectedEl.style.display = "none";
      }
      if (jsonPreviewEl) {
        jsonPreviewEl.innerHTML = "";
      }
      currentUploadTotal = 0;
      currentUploadCount = 0;
      appendLog("File cleared", "info");
    };
  }

  // EVENT LISTENER FOR FILE SELECTION
  if (jsonFileEl) {
    jsonFileEl.addEventListener("change", () => {
      const file = jsonFileEl.files && jsonFileEl.files[0];
      if (file) {
        handleFileSelect(file);
      }
    });
  }

  chrome.runtime.onMessage.addListener((msg) => {
    if (msg?.type === "log") {
      appendLog(msg.text);
      finishIfTerminal(msg.text || "");
    }
    // UPLOAD PROGRESS MESSAGE LISTENER
    if (msg?.type === "upload_progress") {
      updateUploadProgress(msg.current || 0, msg.total || 0);
    }
  });

  loadProfile();
  renderInternshipOptions(readCachedInternshipOptions());
  initializeTabs();
  setBusy(false, "Idle");
  void refreshInternshipOptions({ silent: true });

  // INITIALIZATION: SETUP DRAG AND DROP
  setupDragAndDrop();

  // INITIALIZATION: UPDATE SESSION BADGE
  updateSessionBadge();

  // INITIALIZATION: Open a long-lived progress port (SSE-like)
  try {
    const progressPort = chrome.runtime.connect({ name: "vtu-progress" });
    progressPort.onMessage.addListener((msg) => {
      if (!msg) return;
      if (msg.type === "upload_progress") {
        updateUploadProgress(msg.uploaded || 0, msg.total || 0);
      }
      if (msg.type === "connected") {
        // optional: show queue length
        if (typeof msg.queueLength === "number") {
          appendLog(`Queue length: ${msg.queueLength}`, "info");
        }
      }
    });
  } catch (e) {
    // ignore if connect not available
  }

  // INITIALIZATION: LISTEN FOR TAB CHANGES TO UPDATE SESSION BADGE
  chrome.tabs.onActivated.addListener(() => {
    updateSessionBadge();
  });

  (async () => {
    const restored = await restorePersistedLogs();
    if (!restored) {
      appendLog("Ready", "success", Date.now(), { persist: false });
      appendLog("Quick Start: 1) Keep VTU session logged in 2) Select JSON 3) Upload Entries.", "note", Date.now(), { persist: false });
    } else {
      appendLog(`Restored ${restored} previous log entries`, "info", Date.now(), { persist: false });
      appendLog("Quick Start: 1) Keep VTU session logged in 2) Select JSON 3) Upload Entries.", "note", Date.now(), { persist: false });
    }
    // Fetch runtime stats and config (admin endpoints)
    try {
      const statsRes = await sendRuntimeMessage({ type: 'get_stats' });
      if (statsRes?.ok && statsRes.stats) {
        appendLog(`Stats — Uploaded: ${statsRes.stats.totalUploaded || 0}, Failed: ${statsRes.stats.totalFailed || 0}`, 'info');
        // Update stats display in settings tab
        document.getElementById('statUploaded').textContent = statsRes.stats.totalUploaded || 0;
        document.getElementById('statFailed').textContent = statsRes.stats.totalFailed || 0;
        document.getElementById('statSkipped').textContent = statsRes.stats.totalSkipped || 0;
      }
    } catch (_) {}

    try {
      const cfgRes = await sendRuntimeMessage({ type: 'get_runtime_config' });
      if (cfgRes?.ok && cfgRes.config) {
        appendLog(`Config — batchSize: ${cfgRes.config.batchSize}, concurrency: ${cfgRes.config.maxConcurrentUploads}`, 'info');
        // Update settings sliders with current config
        document.getElementById('batchSizeRange').value = cfgRes.config.batchSize || 5;
        document.getElementById('batchSizeVal').textContent = cfgRes.config.batchSize || 5;
        document.getElementById('maxConcurrentRange').value = cfgRes.config.maxConcurrentUploads || 2;
        document.getElementById('maxConcurrentVal').textContent = cfgRes.config.maxConcurrentUploads || 2;
        document.getElementById('requestDelayRange').value = cfgRes.config.requestDelayMs || 300;
        document.getElementById('requestDelayVal').textContent = (cfgRes.config.requestDelayMs || 300) + ' ms';
        document.getElementById('maxRetriesRange').value = cfgRes.config.maxAttempts || 3;
        document.getElementById('maxRetriesVal').textContent = cfgRes.config.maxAttempts || 3;
      }
    } catch (_) {}
  })();

  // SETTINGS TAB: RANGE SLIDER UPDATES
  document.getElementById('batchSizeRange').addEventListener('input', (e) => {
    document.getElementById('batchSizeVal').textContent = e.target.value;
  });
  document.getElementById('maxConcurrentRange').addEventListener('input', (e) => {
    document.getElementById('maxConcurrentVal').textContent = e.target.value;
  });
  document.getElementById('requestDelayRange').addEventListener('input', (e) => {
    document.getElementById('requestDelayVal').textContent = e.target.value + ' ms';
  });
  document.getElementById('maxRetriesRange').addEventListener('input', (e) => {
    document.getElementById('maxRetriesVal').textContent = e.target.value;
  });

  // SETTINGS TAB: SAVE SETTINGS
  document.getElementById('saveSettings').addEventListener('click', async () => {
    const newConfig = {
      batchSize: parseInt(document.getElementById('batchSizeRange').value),
      maxConcurrentUploads: parseInt(document.getElementById('maxConcurrentRange').value),
      requestDelayMs: parseInt(document.getElementById('requestDelayRange').value),
      maxAttempts: parseInt(document.getElementById('maxRetriesRange').value)
    };
    try {
      const res = await sendRuntimeMessage({ type: 'set_runtime_config', config: newConfig });
      if (res?.ok) {
        appendLog('Settings saved successfully', 'success');
      } else {
        appendLog('Failed to save settings', 'error');
      }
    } catch (err) {
      appendLog('Error saving settings: ' + err.message, 'error');
    }
  });

  // SETTINGS TAB: CLEAR DEDUP CACHE
  document.getElementById('clearDedup').addEventListener('click', async () => {
    try {
      await chrome.storage.local.remove('vtu_dedup_cache');
      appendLog('Deduplication cache cleared', 'success');
    } catch (err) {
      appendLog('Error clearing cache: ' + err.message, 'error');
    }
  });

  // UPLOAD TAB: CANCEL UPLOAD
  document.getElementById('cancelUpload').addEventListener('click', async () => {
    try {
      await sendRuntimeMessage({ type: 'cancel_upload' });
      document.getElementById('cancelUpload').style.display = 'none';
      document.getElementById('upload').style.display = 'block';
      appendLog('Upload cancelled', 'warn');
    } catch (err) {
      appendLog('Error cancelling upload: ' + err.message, 'error');
    }
  });
});
