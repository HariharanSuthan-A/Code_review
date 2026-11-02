// Excela AI — text-only chat frontend.
//
// Transport note (voice-ready): the chat composer below is currently the ONLY
// way to talk to the agent. Live status reuses the /ws/agent status-stream
// contract ("Thinking…" / "Figuring out the edits…" / "Reading workbook…" /
// "Editing Excel…" / "Verifying…" / "✓ Completed") plus a typing indicator,
// so a future voice (STT/TTS) transport can re-attach around the same frames
// without rearchitecting the agent or tool layer.

const API = "/api";
const SESSION_ID = "default";

// ---------------------------------------------------------------- state
let currentSheets = [];
let activeSheet = null;
let pendingConfirmMessage = null; // set when the agent asked "are you sure?"
const PAGE_SIZE = 25; // rows per preview page
const PREVIEW_COLS = 50; // columns fetched per page (grid scrolls horizontally)
let previewPage = 1;
let previewTotalRows = 0;

// ---------------------------------------------------------------- dom refs
const el = (id) => document.getElementById(id);
const wbName = el("wbName");
const wbSheet = el("wbSheet");
const sheetTabs = el("sheetTabs");
const sheetGrid = el("sheetGrid");
const gridMeta = el("gridMeta");
const chatLog = el("chatLog");
const composerForm = el("composerForm");
const composerInput = el("composerInput");
const statusDot = el("statusDot");
const statusText = el("statusText");
const confirmBox = el("confirmBox");
const confirmText = el("confirmText");
const activityLog = el("activityLog");

// ---------------------------------------------------------------- status helpers
// Live working phases shown in the status line AND inside the typing bubble
// while the agent is thinking / reading / editing.
const WORKING_PHASES = [
  "Thinking…",
  "Figuring out the edits…",
  "Reading workbook…",
  "Editing Excel…",
  "Verifying…",
];
let statusCycleTimer = null;
let statusPhaseIdx = 0;
let typingEl = null;
let typingTextEl = null;
let typingCycleTimer = null;
let typingPhaseIdx = 0;

function setStatus(text, mode = "") {
  statusText.textContent = text;
  statusDot.className = "status-dot" + (mode ? " " + mode : "");
  activityLog.textContent = text.toLowerCase();
}

function stopStatusCycle() {
  if (statusCycleTimer) {
    clearInterval(statusCycleTimer);
    statusCycleTimer = null;
  }
}

function startStatusCycle() {
  stopStatusCycle();
  statusPhaseIdx = 0;
  setStatus(WORKING_PHASES[0], "active");
  statusCycleTimer = setInterval(() => {
    statusPhaseIdx = (statusPhaseIdx + 1) % WORKING_PHASES.length;
    // Don't stomp a fresher server-pushed status that isn't in our cycle;
    // the WS handler calls setStatus directly for server truth.
    setStatus(WORKING_PHASES[statusPhaseIdx], "active");
  }, 2200);
}

function showTyping() {
  hideTyping();
  if (chatLog.querySelector(".chat-empty")) chatLog.innerHTML = "";
  typingEl = document.createElement("div");
  typingEl.className = "chat-msg assistant typing";
  typingEl.setAttribute("aria-live", "polite");
  const dots = document.createElement("span");
  dots.className = "typing-dots";
  dots.innerHTML = "<span></span><span></span><span></span>";
  typingTextEl = document.createElement("span");
  typingTextEl.className = "typing-text";
  typingPhaseIdx = 0;
  typingTextEl.textContent = WORKING_PHASES[0];
  typingEl.appendChild(dots);
  typingEl.appendChild(typingTextEl);
  chatLog.appendChild(typingEl);
  chatLog.scrollTop = chatLog.scrollHeight;
  typingCycleTimer = setInterval(() => {
    typingPhaseIdx = (typingPhaseIdx + 1) % WORKING_PHASES.length;
    if (typingTextEl) typingTextEl.textContent = WORKING_PHASES[typingPhaseIdx];
  }, 2200);
}

function setTypingText(text) {
  if (typingTextEl) typingTextEl.textContent = text;
}

function hideTyping() {
  if (typingCycleTimer) {
    clearInterval(typingCycleTimer);
    typingCycleTimer = null;
  }
  if (typingEl && typingEl.parentNode) typingEl.parentNode.removeChild(typingEl);
  typingEl = null;
  typingTextEl = null;
}

function setComposerEnabled(enabled) {
  composerInput.disabled = !enabled;
  const btn = composerForm.querySelector('button[type="submit"]');
  if (btn) btn.disabled = !enabled;
}

function addChatMsg(role, text) {
  if (chatLog.querySelector(".chat-empty")) chatLog.innerHTML = "";
  const div = document.createElement("div");
  div.className = `chat-msg ${role}`;
  div.textContent = text;
  chatLog.appendChild(div);
  chatLog.scrollTop = chatLog.scrollHeight;
}

function addSystemNote(text) {
  const div = document.createElement("div");
  div.className = "chat-msg system";
  div.textContent = text;
  chatLog.appendChild(div);
  chatLog.scrollTop = chatLog.scrollHeight;
}

// ---------------------------------------------------------------- workbook state / preview
async function refreshWorkbook() {
  const res = await fetch(`${API}/workbook`);
  const data = await res.json();
  if (!data.loaded) {
    wbName.textContent = "— none open —";
    wbSheet.textContent = "";
    currentSheets = [];
    return;
  }
  wbName.textContent = (data.path || "untitled").split("/").pop();
  activeSheet = data.active_sheet;
  currentSheets = data.sheets;
  wbSheet.textContent = activeSheet;
  previewPage = 1;
  renderSheetTabs();
  await renderPreview();
}

function renderSheetTabs() {
  sheetTabs.innerHTML = "";
  currentSheets.forEach((name) => {
    const tab = document.createElement("button");
    tab.className = "sheet-tab" + (name === activeSheet ? " active" : "");
    tab.textContent = name;
    tab.onclick = async () => {
      activeSheet = name;
      wbSheet.textContent = name;
      previewPage = 1;
      renderSheetTabs();
      await renderPreview();
    };
    sheetTabs.appendChild(tab);
  });
}

function colLetter(n) {
  let s = "";
  while (n > 0) {
    const rem = (n - 1) % 26;
    s = String.fromCharCode(65 + rem) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

async function renderPreview() {
  if (!activeSheet) return;
  const startRow = (previewPage - 1) * PAGE_SIZE + 1;
  const res = await fetch(
    `${API}/workbook/preview?sheet=${encodeURIComponent(activeSheet)}&start_row=${startRow}&max_rows=${PAGE_SIZE}&max_cols=${PREVIEW_COLS}`
  );
  if (!res.ok) return;
  const data = await res.json();
  const rows = data.preview || [];
  const firstRow = data.start_row || startRow;
  const dataStartRow = data.data_start_row || firstRow;
  const headers = data.headers || data.header_row || [];
  previewTotalRows = data.max_row || 0;
  const colCount = Math.max(
    rows.reduce((m, r) => Math.max(m, r.length), 1),
    headers.length,
    1
  );

  const thead = document.createElement("thead");
  const headRow = document.createElement("tr");
  headRow.appendChild(document.createElement("th"));
  for (let c = 1; c <= colCount; c++) {
    const th = document.createElement("th");
    th.textContent = colLetter(c);
    headRow.appendChild(th);
  }
  thead.appendChild(headRow);

  const tbody = document.createElement("tbody");
  if (rows.length === 0) {
    const tr = document.createElement("tr");
    const td = document.createElement("td");
    td.className = "empty-state";
    td.colSpan = colCount + 1;
    td.textContent = previewPage === 1
      ? "This sheet is empty. Try: \u201cAdd a row with name, age, salary\u201d."
      : "No more rows on this page.";
    tr.appendChild(td);
    tbody.appendChild(tr);
  } else {
    rows.forEach((row, i) => {
      // Row 1 is the column-names row: pinned on EVERY page (backend
      // prepends it when start_row > 1) and rendered as a highlighted,
      // bordered header row.
      const isHeaderRow = i === 0;
      const tr = document.createElement("tr");
      if (isHeaderRow) tr.className = "header-row";
      const rowNumTd = document.createElement("td");
      rowNumTd.className = "row-num" + (isHeaderRow ? " header-row-num" : "");
      rowNumTd.textContent = isHeaderRow ? 1 : (firstRow <= 1 ? firstRow + i : dataStartRow + (i - 1));
      rowNumTd.title = isHeaderRow ? "Column names (row 1 — pinned on every page)" : "";
      tr.appendChild(rowNumTd);
      for (let c = 0; c < colCount; c++) {
        const td = document.createElement("td");
        const v = row[c];
        td.textContent = v === null || v === undefined ? "" : v;
        if (isHeaderRow) td.className = "header-cell";
        tr.appendChild(td);
      }
      tbody.appendChild(tr);
    });
  }

  sheetGrid.innerHTML = "";
  sheetGrid.appendChild(thead);
  sheetGrid.appendChild(tbody);
  el("gridWrap").scrollTop = 0;
  updatePager(firstRow, rows.length);
}

function totalPages() {
  return Math.max(1, Math.ceil(previewTotalRows / PAGE_SIZE));
}

function updatePager(firstRow, rowsShown) {
  const pages = totalPages();
  if (previewPage > pages) previewPage = pages;
  const lastRow = rowsShown ? firstRow + rowsShown - 1 : firstRow - 1;
  gridMeta.textContent = previewTotalRows
    ? `Rows ${rowsShown ? `${firstRow}–${lastRow}` : "—"} of ${previewTotalRows} • ${PAGE_SIZE}/page • row 1 = column names (pinned)`
    : "empty sheet";
  el("pageInfo").textContent = `Page ${previewPage} of ${pages}`;
  el("prevPageBtn").disabled = previewPage <= 1;
  el("nextPageBtn").disabled = previewPage >= pages;
}

el("prevPageBtn").addEventListener("click", async () => {
  if (previewPage > 1) {
    previewPage -= 1;
    await renderPreview();
  }
});

el("nextPageBtn").addEventListener("click", async () => {
  if (previewPage < totalPages()) {
    previewPage += 1;
    await renderPreview();
  }
});

// ---------------------------------------------------------------- upload / new / save / undo / download
el("fileInput").addEventListener("change", async (e) => {
  const file = e.target.files[0];
  if (!file) return;
  setStatus("Uploading workbook…", "active");
  const form = new FormData();
  form.append("file", file);
  const res = await fetch(`${API}/workbook/upload`, { method: "POST", body: form });
  if (!res.ok) {
    const err = await res.json().catch(() => ({}));
    setStatus("Upload failed", "error");
    addSystemNote(`Upload failed: ${err.detail || res.statusText}`);
    return;
  }
  setStatus("Workbook loaded", "done");
  addSystemNote(`Opened ${file.name}`);
  await refreshWorkbook();
});

el("newWbBtn").addEventListener("click", async () => {
  const name = prompt("New workbook filename:", "untitled.xlsx");
  if (!name) return;
  await fetch(`${API}/workbook/new?filename=${encodeURIComponent(name)}`, { method: "POST" });
  addSystemNote(`Created ${name}`);
  await refreshWorkbook();
});

el("saveBtn").addEventListener("click", async () => {
  setStatus("Saving…", "active");
  const res = await fetch(`${API}/workbook/save`, { method: "POST" });
  if (res.ok) {
    setStatus("Saved", "done");
    addSystemNote("Workbook saved");
  } else {
    setStatus("Save failed", "error");
  }
});

el("downloadBtn").addEventListener("click", () => {
  window.location.href = `${API}/workbook/download`;
});

el("undoBtn").addEventListener("click", async () => {
  const res = await fetch(`${API}/undo`, { method: "POST" });
  if (res.ok) {
    addSystemNote("Reverted last change");
    await refreshWorkbook();
  } else {
    const err = await res.json().catch(() => ({}));
    addSystemNote(`Nothing to undo${err.detail ? ": " + err.detail : ""}`);
  }
});

// ---------------------------------------------------------------- agent messaging (text via WS status stream, REST fallback)
let ws = null;
let wsReady = false;

function ensureSocket() {
  return new Promise((resolve) => {
    if (ws && wsReady) return resolve(ws);
    try {
      const proto = location.protocol === "https:" ? "wss" : "ws";
      ws = new WebSocket(`${proto}://${location.host}/api/ws/agent`);
    } catch (err) {
      return resolve(null);
    }
    const timeout = setTimeout(() => resolve(null), 3000);
    ws.onopen = () => {
      wsReady = true;
      clearTimeout(timeout);
      resolve(ws);
    };
    ws.onerror = () => {
      clearTimeout(timeout);
      resolve(null);
    };
    ws.onclose = () => {
      wsReady = false;
      ws = null;
    };
  });
}

function sendViaSocket(message, confirmDestructive) {
  return new Promise((resolve, reject) => {
    const socket = ws;
    const onMessage = (ev) => {
      let data;
      try {
        data = JSON.parse(ev.data);
      } catch (err) {
        return;
      }
      if (data.status && !data.reply) {
        // Server heartbeat (Thinking… / Figuring… / Reading… / Editing…).
        // Mirror it in both the status line and the typing bubble.
        stopStatusCycle();
        setStatus(data.status, data.status.includes("Completed") ? "done" : "active");
        setTypingText(data.status);
        return;
      }
      socket.removeEventListener("message", onMessage);
      resolve(data);
    };
    socket.addEventListener("message", onMessage);
    socket.send(JSON.stringify({
      message,
      session_id: SESSION_ID,
      confirm_destructive: confirmDestructive,
    }));
    setTimeout(() => {
      socket.removeEventListener("message", onMessage);
      reject(new Error("ws timeout"));
    }, 120000);
  });
}

async function sendViaRest(message, confirmDestructive) {
  // No server push over REST — the local status/typing cyclers (started by
  // the caller) carry the "Thinking / Figuring / Editing" animation.
  const res = await fetch(`${API}/agent/message`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      message,
      session_id: SESSION_ID,
      confirm_destructive: confirmDestructive,
    }),
  });
  setTypingText("Editing Excel…");
  setStatus("Updating Excel…", "active");
  const data = await res.json();
  if (!res.ok) throw new Error(data.detail || "Something went wrong.");
  return data;
}

function handleAgentResult(message, data) {
  hideTyping();
  stopStatusCycle();
  addChatMsg("assistant", data.reply);
  if (data.needs_confirmation) {
    pendingConfirmMessage = message;
    confirmText.textContent = data.reply;
    confirmBox.hidden = false;
    setStatus("Waiting for confirmation", "active");
  } else {
    pendingConfirmMessage = null;
    setStatus("✓ Completed", "done");
    refreshWorkbook();
  }
  setComposerEnabled(true);
}

async function sendAgentMessage(message, confirmDestructive = false) {
  addChatMsg("user", message);
  confirmBox.hidden = true;
  showTyping();
  startStatusCycle();
  setComposerEnabled(false);

  // Prefer the WS status-stream contract (live Thinking/Editing/Completed);
  // fall back to REST if the socket is unavailable.
  const socket = await ensureSocket();
  if (socket) {
    try {
      const data = await sendViaSocket(message, confirmDestructive);
      handleAgentResult(message, data);
      return;
    } catch (err) {
      addSystemNote("Live channel dropped — retrying over REST.");
    }
  }
  try {
    const data = await sendViaRest(message, confirmDestructive);
    handleAgentResult(message, data);
  } catch (err) {
    hideTyping();
    stopStatusCycle();
    setStatus("Error", "error");
    addChatMsg("assistant", err.message || "I couldn't reach the agent backend.");
    setComposerEnabled(true);
  }
}

composerForm.addEventListener("submit", (e) => {
  e.preventDefault();
  const text = composerInput.value.trim();
  if (!text) return;
  composerInput.value = "";
  sendAgentMessage(text);
});

el("confirmYes").addEventListener("click", () => {
  if (pendingConfirmMessage) {
    const msg = pendingConfirmMessage;
    pendingConfirmMessage = null;
    confirmBox.hidden = true;
    sendAgentMessage(msg, true);
  }
});

el("confirmNo").addEventListener("click", () => {
  pendingConfirmMessage = null;
  confirmBox.hidden = true;
  addSystemNote("Cancelled");
  setStatus("Idle");
});

// ---------------------------------------------------------------- init
refreshWorkbook();
setStatus("Idle");
