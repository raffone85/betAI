const fileInput = document.getElementById("file-o05");
const analyzeBtn = document.getElementById("analyze-o05");
const statusBox = document.getElementById("o05-status");

const premiumBox = document.getElementById("o05-premium");
const listaABox = document.getElementById("o05-lista-a");
const listaBBox = document.getElementById("o05-lista-b");
const finalBox = document.getElementById("results-box");

const o05Count = document.getElementById("o05-count");
const finalCount = document.getElementById("final-count");

let selectedFile = null;

fileInput.addEventListener("change", (e) => {
  selectedFile = e.target.files[0] || null;
  statusBox.textContent = selectedFile
    ? `File caricato: ${selectedFile.name}`
    : "Nessun file caricato.";
});

analyzeBtn.addEventListener("click", async () => {
  if (!selectedFile) {
    statusBox.textContent = "Carica prima un file Excel.";
    return;
  }

  if (typeof XLSX === "undefined") {
    statusBox.textContent = "Libreria Excel non caricata.";
    return;
  }

  statusBox.textContent = "Analisi in corso...";

  try {
    const rows = await readExcelRows(selectedFile);
    const results = analyzeOver05(rows);

    renderList(premiumBox, results.premium, "O0.5 PT", false);
    renderList(listaABox, results.listaA, "O0.5 PT", false);
    renderList(listaBBox, results.listaB, "O0.5 PT", false);

    const allResults = [
      ...results.premium.map(item => ({ ...item, rank: 1 })),
      ...results.listaA.map(item => ({ ...item, rank: 2 })),
      ...results.listaB.map(item => ({ ...item, rank: 3 }))
    ].sort((a, b) => a.rank - b.rank || (b.score || 0) - (a.score || 0));

    renderList(finalBox, allResults, "O0.5 PT", true);

    const total = allResults.length;
    o05Count.textContent = `${total} esiti`;
    finalCount.textContent = `${total} esiti`;

    statusBox.textContent =
      `Analisi completata. Righe lette: ${rows.length} | Modalità 1: ${results.premium.length} | Lista A: ${results.listaA.length} | Lista B: ${results.listaB.length}`;
  } catch (error) {
    statusBox.textContent = `Errore analisi: ${error.message}`;
  }
});

async function readExcelRows(file) {
  const buffer = await readFileAsArrayBuffer(file);
  const workbook = XLSX.read(buffer, { type: "array" });
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const rawRows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });

  if (!rawRows || !rawRows.length) return [];

  let headerIndex = -1;

  for (let i = 0; i < Math.min(rawRows.length, 6); i++) {
    const row = rawRows[i].map(v => String(v).trim());
    if (row.includes("ORA") && row.includes("EVENTO")) {
      headerIndex = i;
      break;
    }
  }

  if (headerIndex === -1) headerIndex = 0;

  const headers = rawRows[headerIndex].map(v => String(v).trim());
  const data = [];

  for (let i = headerIndex + 1; i < rawRows.length; i++) {
    const row = rawRows[i];
    if (!row || row.every(cell => String(cell).trim() === "")) continue;

    const obj = {};
    headers.forEach((header, index) => {
      obj[header || `col_${index}`] = row[index];
    });
    data.push(obj);
  }

  return data;
}

function readFileAsArrayBuffer(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = e => resolve(e.target.result);
    reader.onerror = () => reject(new Error("Impossibile leggere il file Excel"));
    reader.readAsArrayBuffer(file);
  });
}

function analyzeOver05(rows) {
  const premium = [];
  const listaA = [];
  const listaB = [];

  rows.forEach((row) => {
    const ora = readValue(row, ["ORA"]);
    const evento = cleanText(readValue(row, ["EVENTO"]));
    if (!evento) return;

    const ind = toNumber(readValue(row, ["IND"]));
    const delta = toNumber(readValue(row, ["Delta"]));
    const diff = toNumber(readValue(row, ["Diff"]));
    const mge = toNumber(readValue(row, ["MGE"]));
    const cov = toPercent(readValue(row, ["COv0.5%"]));
    const oov = toPercent(readValue(row, ["OOv0.5%"]));
    const qrgg = toNumber(readValue(row, ["QrGG"]));
    const qro25 = toNumber(readValue(row, ["QrO25"]));
    const qro05 = toNumber(readValue(row, ["QROv0.5pt"]));
    const u5 = splitPair(readValue(row, ["U5[C|O]"]));
    const u5ct = splitPair(readValue(row, ["U5CT[C|O]"]));

    if (isFinite(ind) && isFinite(delta) && isFinite(diff)) {
      if (delta < -0.20 && ind < 320 && diff < 0) {
        premium.push({
          ora,
          evento,
          quality: "Forte",
          score: (400 - ind) + (-delta * 100) + (Math.abs(diff) * 10)
        });
      }
    }

    let score = 0;

    if (isFinite(ind)) score += ind < 280 ? 2 : ind < 320 ? 1.5 : ind < 330 ? 1 : 0;
    if (isFinite(delta)) score += delta <= -0.20 ? 2 : delta < -0.05 ? 1 : delta < 0 ? 0.5 : 0;
    if (isFinite(diff)) score += diff < 0 ? 1.5 : diff <= 0.2 ? 0.5 : 0;
    if (isFinite(mge)) score += mge >= 3.2 ? 1.5 : mge >= 2.8 ? 1 : mge >= 2.5 ? 0.5 : 0;

    if (isFinite(cov) && cov >= 75) score += 0.75;
    if (isFinite(oov) && oov >= 75) score += 0.75;
    if (isFinite(qrgg) && isFinite(qro25) && qrgg <= qro25) score += 0.75;
    if (isFinite(qro05)) score += qro05 <= 1.30 ? 0.75 : qro05 <= 1.40 ? 0.5 : 0;

    const u5Sum = sumFinite(u5);
    const u5ctSum = sumFinite(u5ct);

    if (u5Sum >= 9) score += 1;
    else if (u5Sum >= 6) score += 0.5;

    if (u5ctSum >= 8) score += 1.5;
    else if (u5ctSum >= 6) score += 1;

    if (u5ct[0] >= 3 && u5ct[1] >= 3) score += 0.5;
    if (u5ct[0] === 0 || u5ct[1] === 0) score -= 0.5;

    const item = {
      ora,
      evento,
      quality: score >= 7 ? "Forte" : "Media",
      score
    };

    if (score >= 7) listaA.push(item);
    else if (score >= 5.5) listaB.push(item);
  });

  premium.sort((a, b) => b.score - a.score);
  listaA.sort((a, b) => b.score - a.score);
  listaB.sort((a, b) => b.score - a.score);

  return { premium, listaA, listaB };
}

function renderList(container, items, badgeText, showQuality) {
  if (!items.length) {
    container.innerHTML = `<div class="empty-state">Nessun esito disponibile.</div>`;
    return;
  }

  container.innerHTML = items.map(item => {
    const qualityBadge = showQuality
      ? `<span class="result-badge ${item.quality === "Forte" ? "strong" : "medium"}">${item.quality}</span>`
      : "";

    return `
      <div class="result-item">
        <div class="result-left">
          <div class="result-time">${escapeHtml(item.ora || "-")}</div>
          <div class="result-match">${escapeHtml(item.evento || "-")}</div>
        </div>
        <div class="result-right">
          <span class="result-badge o05">${badgeText}</span>
          ${qualityBadge}
        </div>
      </div>
    `;
  }).join("");
}

function readValue(obj, keys) {
  for (const key of keys) {
    if (obj[key] !== undefined && obj[key] !== null && obj[key] !== "") {
      return obj[key];
    }
  }
  return "";
}

function toNumber(value) {
  if (value === null || value === undefined || value === "") return NaN;
  const match = String(value).replace(",", ".").match(/-?\d+(\.\d+)?/);
  return match ? parseFloat(match[0]) : NaN;
}

function toPercent(value) {
  const n = toNumber(value);
  if (!isFinite(n)) return NaN;
  return n <= 1 ? n * 100 : n;
}

function splitPair(value) {
  if (value === null || value === undefined || value === "") return [NaN, NaN];
  const parts = String(value).split("|");
  return [
    toNumber(parts[0] || ""),
    toNumber(parts[1] || "")
  ];
}

function sumFinite(arr) {
  return arr.filter(n => isFinite(n)).reduce((sum, n) => sum + n, 0);
}

function cleanText(value) {
  return String(value || "").replace(/\s+/g, " ").trim();
}

function escapeHtml(value) {
  return String(value).replace(/[&<>"']/g, function(char) {
    const map = {
      "&": "&amp;",
      "<": "&lt;",
      ">": "&gt;",
      '"': "&quot;",
      "'": "&#39;"
    };
    return map[char];
  });
}
