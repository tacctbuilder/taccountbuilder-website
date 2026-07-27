/* global document, window, localStorage, fetch, XLSX, history */

const API_URL = "https://t-account-builder-api.onrender.com";
const LOCAL_STORAGE_KEY = "tacct_web_session";

let cachedAllValues = [];
let cachedHeaderRowIndex = 0;
let cachedMapping = {};
let cachedFormatType = "";
let cachedRows = [];
let cachedAccountTypeConfig = [];

// ---- Pure helpers copied from the Excel add-in's taskpane.js (no Office.js dependency) ----

function formatAccounting(number) {
  if (number === null || number === undefined || number === 0) return "";
  const formatted = Math.abs(number).toLocaleString("en-US", {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  });
  if (number < 0) return `($${formatted})`;
  return `$${formatted}`;
}

function formatDate(dateVal) {
  if (typeof dateVal === "number") {
    const date = new Date(Math.round((dateVal - 25569) * 86400 * 1000));
    return `${String(date.getMonth() + 1).padStart(2, "0")}/${String(date.getDate()).padStart(2, "0")}/${date.getFullYear()}`;
  } else if (dateVal instanceof Date) {
    return `${String(dateVal.getMonth() + 1).padStart(2, "0")}/${String(dateVal.getDate()).padStart(2, "0")}/${dateVal.getFullYear()}`;
  }
  return dateVal ? dateVal.toString() : "";
}

function guessDebitNormal(name) {
  const lower = name.toLowerCase();
  const debitKeywords = ["asset", "deferred outflow", "expenditure", "expense", "loss", "draw", "prepaid", "receivable"];
  const creditKeywords = ["liability", "revenue", "fund balance", "deferred inflow", "equity", "retained", "income", "gain", "payable", "unearned"];
  for (const kw of debitKeywords) if (lower.includes(kw)) return true;
  for (const kw of creditKeywords) if (lower.includes(kw)) return false;
  return true;
}

function analyzeDataFormat(allValues, headerRowIndex, mapping) {
  const DC_VALUES = new Set(["d", "c", "dr", "cr", "debit", "credit"]);
  const sampleStart = headerRowIndex + 1;
  const sampleEnd = Math.min(sampleStart + 100, allValues.length);
  const numCols = allValues[headerRowIndex] ? allValues[headerRowIndex].length : 0;

  if (mapping.debit_credit_indicator !== null && mapping.debit_credit_indicator !== undefined) {
    let dcCount = 0, total = 0;
    for (let i = sampleStart; i < sampleEnd; i++) {
      const val = ((allValues[i] || [])[mapping.debit_credit_indicator] || "").toString().trim().toLowerCase();
      if (val) { total++; if (DC_VALUES.has(val)) dcCount++; }
    }
    if (total > 5 && dcCount / total > 0.5) {
      return { format: "combined", mapping };
    }
  }

  for (let col = 0; col < numCols; col++) {
    let dcCount = 0, total = 0;
    for (let i = sampleStart; i < sampleEnd; i++) {
      const val = ((allValues[i] || [])[col] !== undefined ? (allValues[i] || [])[col] : "").toString().trim().toLowerCase();
      if (val) { total++; if (DC_VALUES.has(val)) dcCount++; }
    }
    if (total >= 10 && dcCount / total >= 0.8) {
      const updatedMapping = Object.assign({}, mapping, {
        debit_credit_indicator: col,
        ledger_debit_amount: null,
        ledger_credit_amount: null,
      });
      return { format: "combined", mapping: updatedMapping };
    }
  }

  const hasDebitCol = mapping.ledger_debit_amount !== null && mapping.ledger_debit_amount !== undefined;
  const hasCreditCol = mapping.ledger_credit_amount !== null && mapping.ledger_credit_amount !== undefined;
  if (hasDebitCol || hasCreditCol) {
    for (let i = sampleStart; i < sampleEnd; i++) {
      const d = hasDebitCol ? parseFloat((allValues[i] || [])[mapping.ledger_debit_amount]) : NaN;
      const c = hasCreditCol ? parseFloat((allValues[i] || [])[mapping.ledger_credit_amount]) : NaN;
      if ((!isNaN(d) && d !== 0) || (!isNaN(c) && c !== 0)) {
        return { format: "separate", mapping };
      }
    }
  }

  return { format: null, mapping };
}

function readDataWithMapping(allValues, headerRowIndex, mapping, formatType) {
  const rows = [];

  for (let i = headerRowIndex + 1; i < allValues.length; i++) {
    const row = allValues[i];

    const journalNumber = (mapping.journal_number !== null && mapping.journal_number !== undefined)
      ? (row[mapping.journal_number] || "").toString().trim()
      : "";

    const dateVal = (mapping.accounting_date !== null && mapping.accounting_date !== undefined) ? row[mapping.accounting_date] : "";
    const accountType = (mapping.ledger_account_type !== null && mapping.ledger_account_type !== undefined) ? (row[mapping.ledger_account_type] || "").toString() : "";
    const account = (mapping.ledger_account !== null && mapping.ledger_account !== undefined) ? (row[mapping.ledger_account] || "").toString() : "";
    const journal = (mapping.journal !== null && mapping.journal !== undefined) ? (row[mapping.journal] || "").toString() : "";

    if (!account) continue;

    const rowData = {
      journal_number: journalNumber.toString(),
      ledger_account_type: accountType,
      ledger_account: account,
      journal: journal,
      accounting_date: formatDate(dateVal),
      ledger_debit_amount: 0,
      ledger_credit_amount: 0,
      debit_credit_indicator: null,
      amount: null,
    };

    if (formatType === "separate") {
      rowData.ledger_debit_amount = mapping.ledger_debit_amount !== null ? (parseFloat(row[mapping.ledger_debit_amount]) || 0) : 0;
      rowData.ledger_credit_amount = mapping.ledger_credit_amount !== null ? (parseFloat(row[mapping.ledger_credit_amount]) || 0) : 0;
    } else if (formatType === "combined") {
      rowData.debit_credit_indicator = mapping.debit_credit_indicator !== null ? (row[mapping.debit_credit_indicator] || "").toString() : "";
      rowData.amount = mapping.amount !== null ? (parseFloat(row[mapping.amount]) || 0) : 0;
    }

    rows.push(rowData);
  }

  return rows;
}

function detectAccountTypes(allValues, headerRowIndex, accountTypeColIndex) {
  const seen = [];
  const seenSet = new Set();
  for (let i = headerRowIndex + 1; i < allValues.length; i++) {
    const val = ((allValues[i] || [])[accountTypeColIndex] || "").toString().trim();
    if (val && !seenSet.has(val)) {
      seenSet.add(val);
      seen.push(val);
    }
  }
  return seen.map((name, i) => ({
    name,
    display_order: i,
    normal_debit_balance: guessDebitNormal(name),
  }));
}

// ---- DOM wiring ----

const fileInput = document.getElementById("file-input");
const headerRowInput = document.getElementById("header-row");
const detectColumnsBtn = document.getElementById("detect-columns-btn");
const uploadStatus = document.getElementById("upload-status");
const previewCard = document.getElementById("preview-card");
const previewStatus = document.getElementById("preview-status");
const previewContainer = document.getElementById("t-account-preview");
const downloadCard = document.getElementById("download-card");
const buyBtn = document.getElementById("buy-btn");
const downloadBtn = document.getElementById("download-btn");
const creditsBadge = document.getElementById("credits-badge");
const downloadStatus = document.getElementById("download-status");

function setStatus(el, message, type) {
  el.textContent = message;
  el.className = "status-message" + (type ? ` ${type}` : "");
}

function getSavedSession() {
  const saved = localStorage.getItem(LOCAL_STORAGE_KEY);
  return saved ? JSON.parse(saved) : null;
}

function saveSession(session) {
  localStorage.setItem(LOCAL_STORAGE_KEY, JSON.stringify(session));
}

function updateDownloadUI() {
  const session = getSavedSession();
  if (session && session.remaining > 0) {
    buyBtn.style.display = "none";
    downloadBtn.style.display = "inline-block";
    creditsBadge.style.display = "inline-block";
    creditsBadge.textContent = `${session.remaining} download${session.remaining === 1 ? "" : "s"} left`;
  } else {
    buyBtn.style.display = "inline-block";
    downloadBtn.style.display = "none";
    creditsBadge.style.display = "none";
  }
}

fileInput.addEventListener("change", () => {
  detectColumnsBtn.disabled = !fileInput.files.length;
  setStatus(uploadStatus, "", "");
});

detectColumnsBtn.addEventListener("click", async () => {
  const file = fileInput.files[0];
  if (!file) return;

  setStatus(uploadStatus, "Reading file...", "info");

  try {
    const buffer = await file.arrayBuffer();
    const workbook = XLSX.read(buffer, { type: "array" });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    cachedAllValues = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true });

    cachedHeaderRowIndex = Math.max(parseInt(headerRowInput.value, 10) - 1, 0);
    const headers = (cachedAllValues[cachedHeaderRowIndex] || []).map((h) => (h || "").toString());

    if (!headers.length) {
      setStatus(uploadStatus, "Could not find any columns at that header row. Check the header row number.", "error");
      return;
    }

    const mapResponse = await fetch(`${API_URL}/map-columns`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ headers }),
    });
    const mapResult = await mapResponse.json();

    const analyzed = analyzeDataFormat(cachedAllValues, cachedHeaderRowIndex, mapResult.mapping);
    cachedMapping = analyzed.mapping;
    cachedFormatType = analyzed.format;

    if (cachedMapping.ledger_account_type === null || cachedMapping.ledger_account_type === undefined) {
      setStatus(uploadStatus, "Could not detect an Account Type column in your file. Please check your column headers.", "error");
      return;
    }
    if (cachedMapping.ledger_account === null || cachedMapping.ledger_account === undefined) {
      setStatus(uploadStatus, "Could not detect a Ledger Account column in your file. Please check your column headers.", "error");
      return;
    }
    if (!cachedFormatType) {
      setStatus(uploadStatus, "Could not detect whether this file uses separate debit/credit columns or a D/C indicator column. Please check your data.", "error");
      return;
    }

    cachedAccountTypeConfig = detectAccountTypes(cachedAllValues, cachedHeaderRowIndex, cachedMapping.ledger_account_type);
    cachedRows = readDataWithMapping(cachedAllValues, cachedHeaderRowIndex, cachedMapping, cachedFormatType);

    if (!cachedRows.length) {
      setStatus(uploadStatus, "No data rows found below the header row.", "error");
      return;
    }

    setStatus(uploadStatus, `Detected ${cachedRows.length} rows across ${cachedAccountTypeConfig.length} account types.`, "info");
    await generatePreview();
  } catch (err) {
    setStatus(uploadStatus, `Error reading file: ${err.message}`, "error");
  }
});

async function generatePreview() {
  setStatus(previewStatus, "Generating preview...", "info");
  previewCard.classList.remove("is-disabled");

  try {
    const response = await fetch(`${API_URL}/generate`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        rows: cachedRows,
        account_type_config: cachedAccountTypeConfig,
        journal_numbers: null,
        format_type: cachedFormatType,
      }),
    });
    const data = await response.json();

    if (data.error) {
      setStatus(previewStatus, data.error, "error");
      return;
    }

    renderPreview(data);
    setStatus(previewStatus, "", "");
    downloadCard.classList.remove("is-disabled");
    updateDownloadUI();
  } catch (err) {
    setStatus(previewStatus, `Error generating preview: ${err.message}`, "error");
  }
}

function renderPreview(data) {
  previewContainer.innerHTML = "";

  for (const accountType of data.sorted_types) {
    const accounts = data.account_groups[accountType] || [];
    const balance = data.type_balances[accountType] || 0;

    const typeGroup = document.createElement("div");
    typeGroup.className = "t-type-group";

    const typeHeader = document.createElement("div");
    typeHeader.className = "t-type-header";
    typeHeader.innerHTML = `<span class="type-name">${accountType}</span><span class="type-balance">${formatAccounting(balance)}</span>`;
    typeGroup.appendChild(typeHeader);

    for (const acct of accounts) {
      const box = document.createElement("div");
      box.className = "t-box";

      const nameEl = document.createElement("div");
      nameEl.className = "account-name";
      nameEl.textContent = acct.account_name;
      box.appendChild(nameEl);

      const table = document.createElement("table");
      const thead = document.createElement("thead");
      thead.innerHTML = "<tr><th>Journal Entry - Date</th><th>Debit</th><th>Credit</th></tr>";
      table.appendChild(thead);

      const tbody = document.createElement("tbody");
      for (const entry of acct.entries) {
        const tr = document.createElement("tr");
        tr.innerHTML = `<td>${entry.journal_number} - ${entry.date || ""}</td><td>${entry.debit ? formatAccounting(entry.debit) : ""}</td><td>${entry.credit ? formatAccounting(entry.credit) : ""}</td>`;
        tbody.appendChild(tr);
      }

      const balanceRow = document.createElement("tr");
      balanceRow.className = "balance-row";
      const debitCell = acct.balance_type === "debit" ? formatAccounting(acct.balance) : "";
      const creditCell = acct.balance_type === "credit" ? formatAccounting(acct.balance) : "";
      balanceRow.innerHTML = `<td>Balance:</td><td>${debitCell}</td><td>${creditCell}</td>`;
      tbody.appendChild(balanceRow);

      table.appendChild(tbody);
      box.appendChild(table);
      typeGroup.appendChild(box);
    }

    previewContainer.appendChild(typeGroup);
  }
}

buyBtn.addEventListener("click", async () => {
  setStatus(downloadStatus, "Redirecting to checkout...", "info");
  try {
    const response = await fetch(`${API_URL}/create-checkout-session`, { method: "POST" });
    const result = await response.json();
    if (result.checkout_url) {
      window.location.href = result.checkout_url;
    } else {
      setStatus(downloadStatus, "Could not start checkout. Please try again.", "error");
    }
  } catch (err) {
    setStatus(downloadStatus, `Error starting checkout: ${err.message}`, "error");
  }
});

downloadBtn.addEventListener("click", async () => {
  const session = getSavedSession();
  if (!session) {
    setStatus(downloadStatus, "No active purchase found. Please buy a download bundle first.", "error");
    updateDownloadUI();
    return;
  }

  setStatus(downloadStatus, "Generating your Excel file...", "info");

  try {
    const response = await fetch(`${API_URL}/generate-excel`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        rows: cachedRows,
        account_type_config: cachedAccountTypeConfig,
        journal_numbers: null,
        format_type: cachedFormatType,
        session_id: session.session_id,
      }),
    });

    if (response.status === 402) {
      setStatus(downloadStatus, "No downloads remaining on this purchase. Buy another bundle to continue.", "error");
      session.remaining = 0;
      saveSession(session);
      updateDownloadUI();
      return;
    }

    if (!response.ok) {
      setStatus(downloadStatus, "Something went wrong generating the file. Please try again.", "error");
      return;
    }

    const remainingHeader = response.headers.get("X-Downloads-Remaining");
    const blob = await response.blob();
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = "T-Accounts.xlsx";
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);

    if (remainingHeader !== null) {
      session.remaining = parseInt(remainingHeader, 10);
      saveSession(session);
    }
    updateDownloadUI();
    setStatus(downloadStatus, "Download started.", "info");
  } catch (err) {
    setStatus(downloadStatus, `Error downloading file: ${err.message}`, "error");
  }
});

async function checkForReturningPayment() {
  const params = new URLSearchParams(window.location.search);
  const sessionId = params.get("session_id");
  if (!sessionId) return;

  setStatus(downloadStatus, "Confirming your payment...", "info");

  try {
    const response = await fetch(`${API_URL}/verify-session`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ session_id: sessionId }),
    });
    const result = await response.json();

    if (result.paid) {
      saveSession({ session_id: sessionId, remaining: result.remaining });
      setStatus(downloadStatus, "Payment confirmed! You're ready to download.", "info");
    } else {
      setStatus(downloadStatus, "Payment not confirmed yet. If you completed checkout, refresh this page.", "error");
    }
  } catch (err) {
    setStatus(downloadStatus, `Error confirming payment: ${err.message}`, "error");
  }

  params.delete("session_id");
  const newUrl = window.location.pathname + (params.toString() ? `?${params.toString()}` : "");
  history.replaceState({}, "", newUrl);
  updateDownloadUI();
}

checkForReturningPayment();
updateDownloadUI();
