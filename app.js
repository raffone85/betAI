const fileInputO05 = document.getElementById("file-o05");
const analyzeBtnO05 = document.getElementById("analyze-o05");
const statusBoxO05 = document.getElementById("o05-status");

const premiumBox = document.getElementById("o05-premium");
const listaABox = document.getElementById("o05-lista-a");
const listaBBox = document.getElementById("o05-lista-b");

const fileInputGG25 = document.getElementById("file-gg25");
const analyzeBtnGG25 = document.getElementById("analyze-gg25");
const statusBoxGG25 = document.getElementById("gg25-status");

const ggStrongBox = document.getElementById("gg-strong");
const ggMediumBox = document.getElementById("gg-medium");
const o25StrongBox = document.getElementById("o25-strong");
const o25MediumBox = document.getElementById("o25-medium");
const comboBox = document.getElementById("gg25-combo");

const finalBox = document.getElementById("results-box");
const o05Count = document.getElementById("o05-count");
const gg25Count = document.getElementById("gg25-count");
const finalCount = document.getElementById("final-count");

let selectedFileO05 = null;
let selectedFileGG25 = null;

const appState = {
  over05: [],
  gg25: []
};

fileInputO05.addEventListener("change", (e) => {
  selectedFileO05 = e.target.files[0] || null;
  statusBoxO05.textContent = selectedFileO05
    ? `File caricato: ${selectedFileO05.name}`
    : "Nessun file caricato.";
});

fileInputGG25.addEventListener("change", (e) => {
  selectedFileGG25 = e.target.files[0] || null;
  statusBoxGG25.textContent = selectedFileGG25
    ? `File caricato: ${selectedFileGG25.name}`
    : "Nessun file caricato.";
});

analyzeBtnO05.addEventListener("click", async () => {
  if (!selectedFileO05) {
    statusBoxO05.textContent = "Carica prima un file Excel.";
    return;
  }

  if (typeof XLSX === "undefined") {
    statusBoxO05.textContent = "Libreria Excel non caricata.";
    return;
  }

  statusBoxO05.textContent = "Analisi in corso...";

  try {
    const rows = await readExcelRows(selectedFileO05);
    const results = analyzeOver05(rows);

    renderList(premiumBox, results.premium, "O0.5 PT", false, "o05");
    renderList(listaABox, results.listaA, "O0.5 PT", false, "o05");
    renderList(listaBBox, results.listaB, "O0.5 PT", false, "o05");

    const allResults = [
      ...results.premium.map(item => ({ ...item, market: "O0.5 PT", qualityClass: "strong", rank: 1 })),
      ...results.listaA.map(item => ({ ...item, market: "O0.5 PT", qualityClass: "strong", rank: 2 })),
      ...results.listaB.map(item => ({ ...item, market: "O0.5 PT", qualityClass: "medium", rank: 3 }))
    ].sort((a, b) => a.rank - b.rank || (b.score || 0) - (a.score || 0));

    appState.over05 = allResults;
    o05Count.textContent = `${allResults.length} esiti`;
    updateFinalSummary();

    statusBoxO05.textContent =
      `Analisi completata. Righe lette: ${rows.length} | Modalità 1: ${results.premium.length} | Lista A: ${results.listaA.length} | Lista B: ${results.listaB.length}`;
  } catch (error) {
    statusBoxO05.textContent = `Errore analisi: ${error.message}`;
  }
});

analyzeBtnGG25.addEventListener("click", async () => {
  if (!selectedFileGG25) {
    statusBoxGG25.textContent = "Carica prima un file Excel.";
    return;
  }

  if (typeof XLSX === "undefined") {
    statusBoxGG25.textContent = "Libreria Excel non caricata.";
    return;
  }

  statusBoxGG25.textContent = "Analisi in corso...";

  try {
    const rows = await readExcelRows(selectedFileGG25);
    const results = analyzeGGOver25(rows);

    renderList(ggStrongBox, results.ggStrong, "GG", false, "gg");
    renderList(ggMediumBox, results.ggMedium, "GG", false, "gg");
    renderList(o25StrongBox, results.o25Strong, "Over 2.5", false, "o25");
    renderList(o25MediumBox, results.o25Medium, "Over 2.5", false, "o25");
    renderList(comboBox, results.combo, "Combo", false, "combo");

    const allResults = [
      ...results.ggStrong.map(item => ({ ...item, market: "GG", qualityClass: "strong", rank: 1 })),
      ...results.ggMedium.map(item => ({ ...item, market: "GG", qualityClass: "medium", rank: 2 })),
      ...results.o25Strong.map(item => ({ ...item, market: "Over 2.5", qualityClass: "strong", rank: 3 })),
      ...results.o25Medium.map(item => ({ ...item, market: "Over 2.5", qualityClass: "medium", rank: 4 })),
      ...results.combo.map(item => ({ ...item, market: "Combo", qualityClass: "strong", rank: 0 }))
    ].sort((a, b) => a.rank - b.rank || (b.score || 0) - (a.score || 0));

    appState.gg25 = allResults;
    gg25Count.textContent = `${allResults.length} esiti`;
    updateFinalSummary();

    statusBoxGG25.textContent =
      `Analisi completata. Righe lette: ${rows.length} | GG Forte: ${results.ggStrong.length} | GG Medio: ${results.ggMedium.length} | O2.5 Forte: ${results.o25Strong.length} | O2.5 Medio: ${results.o25Medium.length} | Combo: ${results.combo.length}`;
  } catch (error) {
    statusBoxGG25.textContent = `Errore analisi: ${error.message}`;
  }
});

function updateFinalSummary() {
  const merged = [...appState.over05, ...appState.gg25];

  if (!merged.length) {
    finalBox.innerHTML = `<div class="empty-state">Nessun risultato disponibile.</div>`;
    finalCount.textContent = "0 esiti";
    return;
  }

  finalCount.textContent = `${merged.length} esiti`;

  finalBox.innerHTML = merged.map(item => `
    <div class="result-item">
      <div class="result-left">
        <div class="result-time">${escapeHtml(item.ora || "-")}</div>
        <div class="result-match">${escapeHtml(item.evento || "-")}</div>
      </div>
      <div class="result-right">
        <span class="result-badge ${badgeClassForMarket(item.market)}">${escapeHtml(item.market)}</span>
        <span class="result-badge ${item.qualityClass}">${escapeHtml(item.quality || "Forte")}</span>
      </div>
    </div>
  `).join("");
}

async function readExcelRows(file) {
  const buffer = await readFileAsArrayBuffer(file);
  const workbook = XLSX.read(buffer, { type: "array" });
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const rawRows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });

  if (!rawRows || !rawRows.length) return [];

  let headerIndex = -1;

  for (let i = 0; i < Math.min(rawRows.length, 8); i++) {
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
    const ora = normalizeTime(readValue(row, ["ORA"]));
    const evento = normalizeEvent(readValue(row, ["EVENTO"]));
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

function analyzeGGOver25(rows) {
  const ggStrong = [];
  const ggMedium = [];
  const o25Strong = [];
  const o25Medium = [];
  const combo = [];

  rows.forEach((row) => {
    const ora = normalizeTime(readValue(row, ["ORA"]));
    const evento = normalizeEvent(readValue(row, ["EVENTO"]));
    if (!evento) return;

    const igbc = toNumber(readValue(row, ["IGBc"]));
    const igbo = toNumber(readValue(row, ["IGBo"]));
    const igbt = toNumber(readValue(row, ["IGBt"]));
    const mge = toNumber(readValue(row, ["MGE"]));
    const diff = toNumber(readValue(row, ["Diff"]));
    const pgg = toNumber(readValue(row, ["PGG"]));
    const spread = toNumber(readValue(row, ["S1-S2"]));

    const ggStrongCheck =
      isFinite(igbc) && isFinite(igbo) && isFinite(pgg) && isFinite(diff) && isFinite(spread) &&
      igbc > 120 && igbo > 100 && pgg > 60 && diff <= 0 && spread < 200;

    const ggMediumCheck =
      isFinite(igbc) && isFinite(igbo) && isFinite(pgg) && isFinite(diff) && isFinite(spread) &&
      igbc > 100 && igbo > 80 && pgg > 55 && diff <= 0.5 && spread < 300;

    const o25StrongCheck =
      isFinite(igbt) && isFinite(mge) && isFinite(diff) && isFinite(spread) &&
      igbt >= 300 && mge >= 3.2 && diff <= 0 && spread < 200;

    const o25MediumCheck =
      isFinite(igbt) && isFinite(mge) && isFinite(diff) && isFinite(spread) &&
      igbt >= 260 && mge >= 2.8 && diff <= 0.5 && spread < 300;

    if (ggStrongCheck) {
      ggStrong.push({
        ora,
        evento,
        quality: "Forte",
        score: igbc + igbo + pgg - spread
      });
    } else if (ggMediumCheck) {
      ggMedium.push({
        ora,
        evento,
        quality: "Media",
        score: igbc + igbo + pgg - spread
      });
    }

    if (o25StrongCheck) {
      o25Strong.push({
        ora,
        evento,
        quality: "Forte",
        score: igbt + (mge * 10) - spread
      });
    } else if (o25MediumCheck) {
      o25Medium.push({
        ora,
        evento,
        quality: "Media",
        score: igbt + (mge * 10) - spread
      });
    }

    if (ggStrongCheck && o25StrongCheck) {
      combo.push({
        ora,
        evento,
        quality: "Forte",
        score: igbc + igbo + igbt + (mge * 10) + pgg - spread
      });
    }
  });

  ggStrong.sort((a, b) => b.score - a.score);
  ggMedium.sort((a, b) => b.score - a.score);
  o25Strong.sort((a, b) => b.score - a.score);
  o25Medium.sort((a, b) => b.score - a.score);
  combo.sort((a, b) => b.score - a.score);

  return { ggStrong, ggMedium, o25Strong, o25Medium, combo };
}

function renderList(container, items, badgeText, showQuality, marketClass) {
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
          <span class="result-badge ${marketClass}">${escapeHtml(badgeText)}</span>
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

function normalizeTime(value) {
  const text = String(value || "").replace(/\s+/g, " ").trim();
  const full = text.match(/\d{2}-\d{2}-\d{4}\s+\d{2}:\d{2}/);
  if (full) return full[0];
  const hm = text.match(/\d{2}:\d{2}/);
  return hm ? hm[0] : text || "-";
}

function normalizeEvent(value) {
  const raw = String(value || "").replace(/\r|\n/g, " ").trim();
  const match = raw.match(/^(.*? - .*?)(?:\s{2,}.*)?$/);
  if (match) return cleanText(match[1]);
  return cleanText(raw);
}

function cleanText(value) {
  return String(value || "").replace(/\s+/g, " ").trim();
}

function badgeClassForMarket(market) {
  if (market === "O0.5 PT") return "o05";
  if (market === "GG") return "gg";
  if (market === "Over 2.5") return "o25";
  return "combo";
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
