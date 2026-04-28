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

function chunk(arr, size) {
  const out = [];
  for (let i = 0; i < arr.length; i += size) {
    out.push(arr.slice(i, i + size));
  }
  return out;
}

function safeText(value) {
  return normalizeText(value);
}

function normalizeText(value) {
  let text = String(value || "");

  // Strip HTML tags that sometimes leak from API-rich text fields.
  text = text.replace(/<[^>]*>/g, " ");

  // Decode common entities to avoid literal entity noise in outputs.
  text = text
    .replace(/&nbsp;/gi, " ")
    .replace(/&amp;/gi, "&")
    .replace(/&lt;/gi, "<")
    .replace(/&gt;/gi, ">")
    .replace(/&quot;/gi, '"')
    .replace(/&#39;/gi, "'");

  // Normalize unicode punctuation to plain ASCII for built-in jsPDF fonts.
  text = text
    .replace(/[\u2018\u2019]/g, "'")
    .replace(/[\u201C\u201D]/g, '"')
    .replace(/[\u2013\u2014]/g, "-")
    .replace(/\u2026/g, "...")
    .replace(/[\u00A0\u200B-\u200D\uFEFF]/g, " ");

  // Drop combining marks and unsupported high unicode that prints as garbage in PDF.
  text = text.normalize("NFKD").replace(/[\u0300-\u036f]/g, "");
  text = text.replace(/[^\x09\x0A\x0D\x20-\x7E]/g, " ");

  return text.replace(/\s+/g, " ").trim();
}

function escapeHtml(value) {
  return String(value || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/\"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function generateLearning(text) {
  const clean = normalizeText(text);
  if (!clean) {
    return "Improved understanding of practical concepts.";
  }

  // Use first meaningful sentence when possible for cleaner summary text.
  const firstSentence = clean.split(/[.!?]/).map((s) => s.trim()).find(Boolean) || clean;
  return "Gained practical understanding of " + firstSentence.slice(0, 140);
}

function normalizeProfile(profile) {
  return {
    name: safeText(profile?.name) || "N/A",
    usn: safeText(profile?.usn) || "N/A",
    college: safeText(profile?.college) || "N/A",
    internship: safeText(profile?.internship) || "Bharat Unnati AI Fellowship"
  };
}

function mergeProfiles(fetchedProfile, userInput) {
  const fetched = normalizeProfile(fetchedProfile || {});
  const manual = {
    name: safeText(userInput?.name),
    usn: safeText(userInput?.usn),
    college: safeText(userInput?.college),
    internship: safeText(userInput?.internship)
  };

  return normalizeProfile({
    name: manual.name || fetched.name,
    usn: manual.usn || fetched.usn,
    college: manual.college || fetched.college,
    internship: manual.internship || fetched.internship
  });
}

function buildDocHtml(entries, profile) {
  const p = normalizeProfile(profile);

  return `
  <html>
  <head>
    <meta charset="UTF-8" />
    <style>
      body {
        font-family: "Times New Roman";
        font-size: 12pt;
        color: #111;
        margin: 0;
        padding: 0;
      }

      .page {
        width: 170mm;
        min-height: 257mm;
        margin: 0 auto;
        padding: 16mm 14mm;
        box-sizing: border-box;
        page-break-after: always;
      }

      .entry-page {
        page-break-before: always;
      }

      p {
        margin: 6pt 0;
        font-size: 12pt;
        line-height: 1.45;
      }

      h1 {
        text-align: center;
        font-size: 15pt;
        font-weight: bold;
        margin: 0;
      }

      h2 {
        text-align: center;
        font-size: 13pt;
        font-weight: bold;
        margin: 4pt 0 18pt;
      }

      .meta-line {
        margin: 4pt 0;
      }

      .section-title {
        margin-top: 12pt;
        margin-bottom: 4pt;
      }

      .section-text {
        text-align: justify;
      }

      .footer {
        margin-top: 26mm;
        display: flex;
        justify-content: space-between;
        font-size: 11pt;
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

async function uploadEntries(tabId, rawEntries) {
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

  // Keep first occurrence per date to avoid duplicate writes and extra API traffic.
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
  send("📤 Starting upload to VTU API...");

  const result = await chrome.scripting.executeScript({
    target: { tabId },
    args: [dedupedByDate],
    func: async (entries) => {
      const logs = [];
      const STORE_URL = "https://vtuapi.internyet.in/api/v1/student/internship-diaries/store";
      const LIST_URL = "https://vtuapi.internyet.in/api/v1/student/internship-diaries";
      const MAX_RETRIES = 2;
      const INTER_ENTRY_DELAY_MS = 2500;
      const INTER_ENTRY_JITTER_MS = 700;
      const BASE_RETRY_DELAY_MS = 900;

      const clean = (v) => String(v || "").replace(/\s+/g, " ").trim();
      const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

      function rlog(text) {
        logs.push(text);
        try { chrome.runtime.sendMessage({ type: "log", text }); } catch (_) {}
      }

      async function fetchJson(url, timeoutMs = 6000) {
        const controller = new AbortController();
        const tid = setTimeout(() => controller.abort(), timeoutMs);
        try {
          const res = await fetch(url, {
            credentials: "include",
            headers: { Accept: "application/json" },
            signal: controller.signal
          });
          clearTimeout(tid);
          if (!res.ok) return null;
          return await res.json();
        } catch (e) {
          clearTimeout(tid);
          return null;
        }
      }

      function pickInternshipIdFromDiaryList(json) {
        const list = json?.data?.data || json?.data || [];
        if (!Array.isArray(list) || list.length === 0) {
          return null;
        }

        const withId = list.find((item) => item?.internship_id);
        return withId?.internship_id ? String(withId.internship_id) : null;
      }

      function pickInternshipIdFromDom() {
        const selects = Array.from(document.querySelectorAll("select"));

        for (const select of selects) {
          const value = clean(select.value);
          if (value && /^\d+$/.test(value)) {
            return value;
          }

          const labelText = clean(
            select.name ||
            select.id ||
            select.getAttribute("aria-label") ||
            select.closest("label")?.textContent ||
            ""
          );

          const internshipOption = Array.from(select.options || []).find((option) => {
            const optionValue = clean(option.value);
            const optionText = clean(option.textContent || option.label || "");
            return /^\d+$/.test(optionValue) && /internship/i.test(optionText);
          });

          if (internshipOption?.value) {
            return clean(internshipOption.value);
          }

          if (/internship/i.test(labelText) && value) {
            return value;
          }
        }

        const attrEl = document.querySelector(
          '[data-internship-id], meta[name="internship-id"], meta[name="internship_id"]'
        );
        if (attrEl) {
          return clean(attrEl.getAttribute("content") || attrEl.dataset.internshipId || "");
        }

        return null;
      }

      rlog("🔍 Resolving required IDs...");
      let internshipId = null;

      if (!internshipId) {
        rlog("  trying authenticated diary list first...");
        const diaryListJson = await fetchJson("https://vtuapi.internyet.in/api/v1/student/internship-diaries");
        internshipId = pickInternshipIdFromDiaryList(diaryListJson);
        if (internshipId) {
          rlog(`  ✅ internship_id from diary list: ${internshipId}`);
        }
      }

      if (!internshipId) {
        try {
          const candidates = [
            window.__NEXT_DATA__,
            window.__INITIAL_STATE__,
            window.__APP_STATE__,
            window.initialData
          ];
          for (const state of candidates) {
            if (!state) continue;
            const raw = JSON.stringify(state);
            const m = raw.match(/"internship_id"\s*:\s*(\d+)/);
            if (m) { internshipId = m[1]; break; }
          }
        } catch (_) {}
      }

      if (!internshipId) {
        const metaEl = document.querySelector(
          'meta[name="internship-id"], meta[name="internship_id"], [data-internship-id]'
        );
        if (metaEl) internshipId = metaEl.getAttribute("content") || metaEl.dataset.internshipId;
      }

      if (!internshipId) {
        internshipId = pickInternshipIdFromDom();
        if (internshipId) {
          rlog(`  ✅ internship_id from DOM: ${internshipId}`);
        }
      }

      if (!internshipId) {
        const DIARY_INTERNSHIP_NAME = "Bharat Unnati AI Fellowship";
        const applyUrls = [
          "https://vtuapi.internyet.in/api/v1/student/internship-applys?page=1&status=6",
          "https://vtuapi.internyet.in/api/v1/student/internship-applys?page=1",
          "https://vtuapi.internyet.in/api/v1/student/internship-applys"
        ];
        for (const url of applyUrls) {
          const json = await fetchJson(url);
          if (!json) continue;
          const list = json?.data?.data || json?.data || [];
          if (!Array.isArray(list) || list.length === 0) continue;
          rlog(`  internship-applys → ${list.length} application(s) found`);
          // Prefer the one whose name matches diary entries, else take first paid/active
          let best = list.find(item =>
            item.internship_payment_status === true &&
            (item.internship_details?.name || "").includes(DIARY_INTERNSHIP_NAME)
          ) || list.find(item => item.internship_payment_status === true)
            || list[0];
          if (best?.internship_id) {
            internshipId = String(best.internship_id);
            rlog(`  ✅ internship_id: ${internshipId} (${best.internship_details?.name || ""})`);
            break;
          }
        }
      }

      rlog(`Internship ID: ${internshipId ?? "NOT FOUND"}`);
      let allSkillsList = [];
      let fallbackSkillIds = [];
      const skillByNormalizedName = new Map();
      const skillByCompactName = new Map();
      const skillsJson = await fetchJson("https://vtuapi.internyet.in/api/v1/master/skills");
      if (skillsJson) {
        const skillsList = skillsJson?.data?.data || skillsJson?.data || [];
        rlog(`Skills available: ${JSON.stringify(Array.isArray(skillsList) ? skillsList.slice(0, 5) : skillsList)}`);
        if (Array.isArray(skillsList) && skillsList.length > 0) {
          allSkillsList = skillsList;

          // Build deterministic lookup indexes for robust name->ID mapping.
          skillsList.forEach((skill) => {
            const normalized = normalizeSkillName(skill?.name || "");
            const compact = normalized.replace(/\s+/g, "");
            if (normalized && !skillByNormalizedName.has(normalized)) {
              skillByNormalizedName.set(normalized, skill);
            }
            if (compact && !skillByCompactName.has(compact)) {
              skillByCompactName.set(compact, skill);
            }
          });

          // Prefer Python as default for this internship flow, then safe alternates.
          const preferredFallbackNames = ["python", "machine learning", "javascript", "react"];
          const picked = [];
          preferredFallbackNames.forEach((name) => {
            const found = skillByNormalizedName.get(normalizeSkillName(name));
            if (found?.id) {
              const id = String(found.id);
              if (!picked.includes(id)) {
                picked.push(id);
              }
            }
          });

          fallbackSkillIds = picked.length > 0
            ? picked
            : [String(skillsList[0].id)];
        }
      }

      function normalizeSkillName(value) {
        return clean(value).toLowerCase();
      }

      const skillAliases = {
        python: ["python", "python programming", "python development"],
        js: ["javascript"],
        javascript: ["javascript", "js"],
        ml: ["machine learning", "machinelearning"],
        ai: ["artificial intelligence", "ai"]
      };

      function compactSkillName(value) {
        return normalizeSkillName(value).replace(/\s+/g, "");
      }

      function expandedSkillTerms(desired) {
        const base = normalizeSkillName(desired);
        const alias = skillAliases[base] || [];
        const all = [base, ...alias].map((x) => normalizeSkillName(x)).filter(Boolean);
        return Array.from(new Set(all));
      }

      function resolveSkillIdsForEntry(entrySkills) {
        const requested = Array.isArray(entrySkills)
          ? entrySkills
          : (entrySkills ? [entrySkills] : []);

        const requestedNormalized = requested
          .map((s) => normalizeSkillName(s))
          .filter(Boolean);

        if (requestedNormalized.length > 0 && allSkillsList.length > 0) {
          const resolved = [];

          for (const desired of requestedNormalized) {
            const terms = expandedSkillTerms(desired);
            let pick = null;

            for (const term of terms) {
              pick = skillByNormalizedName.get(term) || null;
              if (pick) {
                break;
              }
            }

            if (!pick) {
              for (const term of terms) {
                const compactTerm = term.replace(/\s+/g, "");
                pick = skillByCompactName.get(compactTerm) || null;
                if (pick) {
                  break;
                }
              }
            }

            if (!pick) {
              for (const term of terms) {
                const compactTerm = compactSkillName(term);
                pick = allSkillsList.find((skill) => {
                  const skillName = normalizeSkillName(skill.name);
                  const compactName = compactSkillName(skill.name);
                  return (
                    skillName.includes(term) ||
                    term.includes(skillName) ||
                    compactName.includes(compactTerm) ||
                    compactTerm.includes(compactName)
                  );
                });
                if (pick) {
                  break;
                }
              }
            }

            if (pick && pick.id) {
              const id = String(pick.id);
              if (!resolved.includes(id)) {
                resolved.push(id);
              }
            }
          }

          if (resolved.length > 0) {
            return resolved;
          }

          // Explicit skills were requested but none mapped; avoid injecting unrelated fallback skills.
          rlog(`  ⚠️ Could not map requested skills: [${requestedNormalized.join(", ")}]. Sending empty skill list.`);
          return [];
        }

        return fallbackSkillIds;
      }

      rlog(`Default Skill IDs: [${fallbackSkillIds.join(", ") || "NOT FOUND"}]`);
      const moodDefault = "5";

      if (!internshipId) {
        rlog("❌ internship_id could not be resolved. Open the VTU internship diary/application page and retry.");
        throw new Error("internship_id not resolved");
      }
      
      const summary = {
        total: entries.length,
        uploaded: 0,
        failed: 0,
        failures: [],
        uploadedVia: null
      };

      rlog("🧩 Uploader build: 2026-04-19-ratelimit-resume-v3");
      rlog(`Starting upload of ${entries.length} entry(ies)...`);

      const existingEntryByDate = new Map();
      try {
        const listRes = await fetch(LIST_URL, {
          credentials: "include",
          headers: { Accept: "application/json" }
        });

        if (listRes.ok) {
          const listData = await listRes.json();
          const diaryList = listData?.data?.data || listData?.data || [];
          if (Array.isArray(diaryList)) {
            diaryList.forEach((entry) => {
              if (entry?.date && entry?.id) {
                existingEntryByDate.set(String(entry.date), String(entry.id));
              }
            });
            rlog(`📚 Loaded ${existingEntryByDate.size} existing diary date(s) in one request`);
          }
        } else {
          rlog(`⚠️ Could not pre-load existing entries (HTTP ${listRes.status}). Continuing...`);
        }
      } catch (err) {
        rlog(`⚠️ Could not pre-load existing entries: ${err?.message || "unknown error"}`);
      }

      for (let i = 0; i < entries.length; i += 1) {
        const e = entries[i];
        rlog(`\n--- Processing entry ${i + 1}/${entries.length}: ${e.date} ---`);

        const existingEntryId = existingEntryByDate.get(e.date) || null;
        rlog(`🔍 Checking for existing entry on ${e.date}...`);
        if (existingEntryId) {
          rlog(`  📝 Found existing entry (ID: ${existingEntryId})`);
        }

        const entrySkillIds = resolveSkillIdsForEntry(e.skills);
        rlog(`  🧠 Skills requested: [${(Array.isArray(e.skills) ? e.skills.join(", ") : "")}], resolved IDs: [${entrySkillIds.join(", ") || "NOT FOUND"}]`);

        const payload = {
          date: e.date,
          hours: Number(e.hours),
          description: e.description,
          learnings: e.learnings,
          skill_ids: entrySkillIds,
          mood_slider: Number(moodDefault)
        };
        if (internshipId) {
          payload.internship_id = internshipId;
        }

        let success = false;
        let lastStatus = "network error";
        let lastDetail = "no response";
        let rateLimitPauseCount = 0;
        const MAX_RATE_LIMIT_PAUSES_PER_ENTRY = 5;
        const BASE_RATE_LIMIT_COOLDOWN_MS = 15000;

        const requestTargets = existingEntryId
          ? [
              { method: "PATCH", url: `https://vtuapi.internyet.in/api/v1/student/internship-diaries/${existingEntryId}` },
              { method: "PUT", url: `https://vtuapi.internyet.in/api/v1/student/internship-diaries/${existingEntryId}` },
              { method: "PATCH", url: `https://vtuapi.internyet.in/api/v1/student/internship-diaries/update/${existingEntryId}` },
              { method: "PUT", url: `https://vtuapi.internyet.in/api/v1/student/internship-diaries/update/${existingEntryId}` }
            ]
          : [
              { method: "POST", url: STORE_URL }
            ];

        if (existingEntryId) {
          rlog(`✏️ Existing entry found. Updating content for ${e.date}...`);
        }

        for (let attempt = 0; attempt <= MAX_RETRIES; attempt += 1) {
          const attemptNo = attempt + 1;
          let rateLimitedThisAttempt = false;
          try {
            rlog(`📤 Attempt ${attemptNo}/${MAX_RETRIES + 1}`);

            for (let targetIndex = 0; targetIndex < requestTargets.length; targetIndex += 1) {
              const target = requestTargets[targetIndex];
              const controller = new AbortController();
              const timeoutId = setTimeout(() => controller.abort(), 15000);

              const res = await fetch(target.url, {
                method: target.method,
                credentials: "include",
                headers: {
                  Accept: "application/json",
                  "Content-Type": "application/json"
                },
                body: JSON.stringify(payload),
                signal: controller.signal
              });

              clearTimeout(timeoutId);

              let responseText = "";
              try {
                responseText = clean((await res.text()).slice(0, 300));
              } catch (_) {
                responseText = "";
              }

              if (res.ok) {
                success = true;
                summary.uploaded += 1;
                if (!summary.uploadedVia) {
                  summary.uploadedVia = `${target.method} ${target.url}`;
                }
                rlog(`✅ Success (${summary.uploaded}/${summary.total})`);
                break;
              }

              lastStatus = String(res.status);
              lastDetail = responseText || "no response";

              if (res.status === 429) {
                const retryAfterSec = Number(res.headers.get("Retry-After") || "0");
                const waitMs = retryAfterSec > 0
                  ? retryAfterSec * 1000
                  : Math.min(12000, BASE_RETRY_DELAY_MS * Math.pow(2, attemptNo));
                rlog(`⏳ Rate limited (429). Waiting ${Math.ceil(waitMs / 1000)}s before retry...`);
                await sleep(waitMs);
                rateLimitedThisAttempt = true;
                break;
              }

              const isLastTarget = targetIndex === requestTargets.length - 1;
              if (isLastTarget) {
                rlog(`❌ Failed attempt ${attemptNo}: ${target.method} ${target.url} -> ${lastStatus} — ${lastDetail}`);
              }
            }

            if (success) {
              break;
            }
          } catch (error) {
            lastStatus = String(error?.message || "network error");
            lastDetail = "exception";
            rlog(`❌ Exception on attempt ${attemptNo}: ${lastStatus}`);
          }

          if (rateLimitedThisAttempt && attempt === MAX_RETRIES) {
            rateLimitPauseCount += 1;

            if (rateLimitPauseCount > MAX_RATE_LIMIT_PAUSES_PER_ENTRY) {
              lastStatus = "429";
              lastDetail = `rate limit persisted after ${MAX_RATE_LIMIT_PAUSES_PER_ENTRY} cooldown(s)`;
              rlog("❌ Rate limit persisted for this entry after multiple cooldowns. Skipping this date and continuing.");
              break;
            }

            const cooldownMs = Math.min(
              120000,
              BASE_RATE_LIMIT_COOLDOWN_MS * rateLimitPauseCount
            );

            rlog(
              `⏸️ Rate limit persisted. Cooling down ${Math.ceil(cooldownMs / 1000)}s then resuming this entry... (${rateLimitPauseCount}/${MAX_RATE_LIMIT_PAUSES_PER_ENTRY})`
            );
            await sleep(cooldownMs);

            // Restart attempts for the same entry after cooldown.
            attempt = -1;
            continue;
          }

          if (rateLimitedThisAttempt) {
            // Already waited based on Retry-After / backoff inside the request loop.
            continue;
          }

          if (attempt < MAX_RETRIES) {
            await sleep(BASE_RETRY_DELAY_MS);
          }
        }

        if (!success) {
          summary.failed += 1;
          summary.failures.push({
            index: i + 1,
            date: e.date,
            status: lastStatus,
            detail: lastDetail
          });
        }

        rlog(`📊 Progress: ${summary.uploaded + summary.failed}/${summary.total} processed`);
        const jitter = Math.floor(Math.random() * (INTER_ENTRY_JITTER_MS + 1));
        await sleep(INTER_ENTRY_DELAY_MS + jitter);
      }

      rlog(`\n=== Upload Complete ===`);
      rlog(`Uploaded: ${summary.uploaded}/${summary.total}`);
      if (summary.failed > 0) {
        rlog(`Failed: ${summary.failed}`);
      }
      
      return summary;
    }
  });

  const summary = result?.[0]?.result;
  if (!summary) {
    send("❌ Upload failed: no response from tab script");
    return;
  }

  send(`✅ Uploaded: ${summary.uploaded}/${summary.total}`);
  if (summary.uploadedVia) {
    send(`🧩 Upload method: ${summary.uploadedVia}`);
  }
  if (summary.failed) {
    send(`❌ Failed: ${summary.failed}`);
    summary.failures.slice(0, 8).forEach((f) => {
      send(`  Row ${f.index} (${f.date}): ${f.status} | ${f.detail}`);
    });
  }
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

  return false;
});

chrome.runtime.onMessage.addListener((msg) => {
  if (msg.type === "start") {
    startExport(msg.profileInput);
  }

  if (msg.type === "upload_entries") {
    (async () => {
      if (uploadInProgress) {
        send("⚠️ Upload already in progress. Please wait for completion.");
        return;
      }

      uploadInProgress = true;
      try {
        const resolved = await resolveUploadTab();
        const tab = resolved.tab;
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

        send("📤 Starting bulk upload...");
        await uploadEntries(tab.id, msg.data);
        send("✅ Upload flow completed");
      } finally {
        uploadInProgress = false;
      }
    })().catch((error) => {
      send(`❌ Upload setup failed: ${error?.message || "Unknown error"}`);
    });
  }
});
