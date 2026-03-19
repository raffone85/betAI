const fileO05 = document.getElementById("file-o05");
const btnO05 = document.getElementById("btn-o05");
const statusO05 = document.getElementById("status-o05");
const countO05 = document.getElementById("count-o05");
const boxO05Premium = document.getElementById("box-o05-premium");
const boxO05A = document.getElementById("box-o05-a");
const boxO05B = document.getElementById("box-o05-b");

const fileGG25 = document.getElementById("file-gg25");
const btnGG25 = document.getElementById("btn-gg25");
const statusGG25 = document.getElementById("status-gg25");
const countGG25 = document.getElementById("count-gg25");
const boxGGStrong = document.getElementById("box-gg-strong");
const boxGGMedium = document.getElementById("box-gg-medium");
const boxO25Strong = document.getElementById("box-o25-strong");
const boxO25Medium = document.getElementById("box-o25-medium");
const boxCombo = document.getElementById("box-combo");

const boxFinal = document.getElementById("box-final");
const countFinal = document.getElementById("count-final");
const btnExport = document.getElementById("btn-export");

let selectedO05 = null;
let selectedGG25 = null;

const state = {
  over05: [],
  gg25: []
};

fileO05.addEventListener("change", (e) => {
  selectedO05 = e.target.files && e.target.files[0] ? e.target.files[0] : null;
  statusO05.textContent = selectedO05
    ? `File caricato: ${selectedO05.name}`
    : "Nessun file caricato.";
});

fileGG25.addEventListener("change", (e) => {
  selectedGG25 = e.target.files && e.target.files[0] ? e.target.files[0] : null;
  statusGG25.textContent = selectedGG25
    ? `File caricato: ${selectedGG25.name}`
    : "Nessun file caricato.";
});

bindPress(btnO05, analyzeO05);
bindPress(btnGG25, analyzeGG25);
bindPress(btnExport, exportTxt);

async function analyzeO05() {
  if (!selectedO05) {
    statusO05.textContent = "Seleziona prima un file Excel.";
    return;
  }

  if (typeof XLSX === "undefined") {
    statusO05.textContent = "Libreria XLSX non caricata.";
    return;
  }

  setLoading(btnO05, true, "Analisi in corso...");
  statusO05.textContent = "Lettura file in corso...";

  try {
    const rows = await readExcelRows(selectedO05, ["ORA", "EVENTO"]);
    const results = buildOver05Results(rows);

    renderList(boxO05Premium, results.premium, "O0.5 PT", "o05", false);
    renderList(boxO05A, results.listaA, "O0.5 PT", "o05", false);
    renderList(boxO05B, results.listaB, "O0.5 PT", "o05", false);

    const merged = [
      ...results.premium.map(x => ({ ...x, market: "O0.5 PT", qualityClass: "strong", rank: 1 })),
      ...results.listaA.map(x => ({ ...x, market: "O0.5 PT", qualityClass: "strong", rank: 2 })),
      ...results.listaB.map(x => ({ ...x, market: "O0.5 PT", qualityClass: "medium", rank: 3 }))
    ].sort((a, b) => a.rank - b.rank || (b.score || 0) - (a.score || 0));

    state.over05 = merged;
    countO05.textContent = `${merged.length} esiti`;
    updateFinal();

    statusO05.textContent =
      `Analisi completata. Righe lette: ${rows.length} | Modalità 1: ${results.premium.length} | Lista A: ${results.listaA.length} | Lista B: ${results.listaB.length}`;
  } catch (error) {
    statusO05.textContent = `Errore: ${error.message}`;
  } finally {
    setLoading(btnO05, false, "Analizza Over 0.5 PT");
  }
}

async function analyzeGG25() {
  if (!selectedGG25) {
    statusGG25.textContent = "Seleziona prima un file Excel.";
    return;
  }

  if (typeof XLSX === "undefined") {
    statusGG25.textContent = "Libreria XLSX non caricata.";
    return;
  }

  setLoading(btnGG25, true, "Analisi in corso...");
  statusGG25.textContent = "Lettura file in corso...";

  try {
    const rows = await readExcelRows(selectedGG25, ["ORA", "EVENTO"]);
    const results = buildGG25Results(rows);

    renderList(boxGGStrong, results.ggStrong, "GG", "gg", false);
    renderList(boxGGMedium, results.ggMedium, "GG", "gg", false);
    renderList(boxO25Strong, results.o25Strong, "Over 2.5", "o25", false);
    renderList(boxO25Medium, results.o25Medium, "Over 2.5", "o25", false);
    renderList(boxCombo, results.combo, "Combo", "combo", false);

    const merged = [
      ...results.combo.map(x => ({ ...x, market: "Combo", qualityClass: "strong", rank: 0 })),
      ...results.ggStrong.map(x => ({ ...x, market: "GG", qualityClass: "strong", rank: 1 })),
      ...results.ggMedium.map(x => ({ ...x, market: "GG", qualityClass: "medium", rank: 2 })),
      ...results.o25Strong.map(x => ({ ...x, market: "Over 2.5", qualityClass: "strong", rank: 3 })),
      ...results.o25Medium.map(x => ({ ...x, market: "Over 2.5", qualityClass: "medium", rank: 4 }))
    ].sort((a, b) => a.rank - b.rank || (b.score || 0) - (a.score || 0));

    state.gg25 = merged;
    countGG25.textContent = `${merged.length} esiti`;
    updateFinal();

    statusGG25.textContent =
      `Analisi completata. Righe lette: ${rows.length} | GG Forte: ${results.ggStrong.length} | GG Medio: ${results.ggMedium.length} | O2.5 Forte: ${results.o25Strong.length} | O2.5 Medio: ${results.o25Medium.length} | Combo: ${results.combo.length}`;
  } catch (error) {
    statusGG25.textContent = `Errore: ${error.message}`;
  } finally {
    setLoading(btnGG25, false, "Analizza GG / Over 2.5");
  }
}

async function readExcelRows(file, requiredHeaders = []) {
  const buffer = await getFileBuffer(file);
  const workbook = XLSX.read(buffer, { type: "array" });
  const sheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];

  const matrix = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
  if (!matrix || !matrix.length) return [];

  const headerIndex = findHeaderRow(matrix, requiredHeaders);
  if (headerIndex === -1) {
    throw new Error("Intestazioni non trovate nel file Excel.");
  }

  const headers = matrix[headerIndex].map((v) => normalizeHeader(v));
  const rows = [];

  for (let i = headerIndex + 1; i < matrix.length; i++) {
    const row = matrix[i];
    if (!row || row.every(cell => cleanText(cell) === "")) continue;

    const obj = {};
    headers.forEach((header, index) => {
      if (!header) return;
      obj[header] = row[index];
    });

    rows.push(obj);
  }

  return rows;
}

function findHeaderRow(matrix, requiredHeaders) {
  const wanted = requiredHeaders.map(normalizeHeader);

  for (let i = 0; i < Math.min(matrix.length, 12); i++) {
    const row = matrix[i].map(normalizeHeader);
    const ok = wanted.every(h => row.includes(h));
    if (ok) return i;
  }

  return -1;
}

async function getFileBuffer(file) {
  if (file && typeof file.arrayBuffer === "function") {
    return await file.arrayBuffer();
  }

  return await new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = e => resolve(e.target.result);
    reader.onerror = () => reject(new Error("Impossibile leggere il file"));
    reader.readAsArrayBuffer(file);
  });
}

function buildOver05Results(rows) {
  const premium = [];
  const listaA = [];
  const listaB = [];

  rows.forEach((row) => {
    const ora = normalizeTime(getVal(row, ["ORA"]));
    const evento = normalizeEvent(getVal(row, ["EVENTO"]));
    if (!evento) return;

    const ind = toNumber(getVal(row, ["IND"]));
    const delta = toPercentNumber(getVal(row, ["DELTA"]));
    const diff = toNumber(getVal(row, ["DIFF"]));
    const mge = toNumber(getVal(row, ["MGE"]));
    const cov = toPercentNumber(getVal(row, ["COV0.5%"]));
    const oov = toPercentNumber(getVal(row, ["OOV0.5%"]));
    const qrgg = toNumber(getVal(row, ["QRGG"]));
    const qro25 = toNumber(getVal(row, ["QRO25"]));
    const qro05 = toNumber(getVal(row, ["QROV0.5PT"]));

    const u5 = splitPairFlexible(getVal(row, ["U5[C|O]"]));
    const u5ct = splitPairFlexible(getVal(row, ["U5CT[C|O]"]));

    if (isFinite(ind) && isFinite(delta) && isFinite(diff)) {
      if (delta < -20 && ind < 320 && diff < 0) {
        premium.push({
          ora,
          evento,
          quality: "Forte",
          score: (400 - ind) + Math.abs(delta) + (Math.abs(diff) * 10)
        });
      }
    }

    let score = 0;

    if (isFinite(ind)) score += ind < 280 ? 2 : ind < 320 ? 1.5 : ind < 360 ? 1 : 0;
    if (isFinite(delta)) score += delta <= -20 ? 2 : delta <= -10 ? 1.25 : delta < 0 ? 0.5 : 0;
    if (isFinite(diff)) score += diff < 0 ? 1.5 : diff <= 0.2 ? 0.5 : 0;
    if (isFinite(mge)) score += mge >= 3.2 ? 1.5 : mge >= 2.8 ? 1 : mge >= 2.4 ? 0.5 : 0;

    if (isFinite(cov)) score += cov >= 80 ? 1 : cov >= 72 ? 0.6 : 0;
    if (isFinite(oov)) score += oov >= 80 ? 1 : oov >= 72 ? 0.6 : 0;
    if (isFinite(qrgg) && isFinite(qro25) && qrgg <= qro25) score += 0.7;
    if (isFinite(qro05)) score += qro05 <= 1.25 ? 0.8 : qro05 <= 1.35 ? 0.5 : 0;

    const u5sum = sumFinite(u5);
    const u5ctsum = sumFinite(u5ct);

    if (u5sum >= 6) score += 0.6;
    if (u5sum >= 8) score += 0.4;
    if (u5ctsum >= 4) score += 0.8;
    if (u5ctsum >= 6) score += 0.6;

    const item = {
      ora,
      evento,
      quality: score >= 6.6 ? "Forte" : "Media",
      score
    };

    if (score >= 6.6) listaA.push(item);
    else if (score >= 5.1) listaB.push(item);
  });

  premium.sort((a, b) => b.score - a.score);
  listaA.sort((a, b) => b.score - a.score);
  listaB.sort((a, b) => b.score - a.score);

  return { premium, listaA, listaB };
}

function buildGG25Results(rows) {
  const ggStrong = [];
  const ggMedium = [];
  const o25Strong = [];
  const o25Medium = [];
  const combo = [];

  rows.forEach((row) => {
    const ora = normalizeTime(getVal(row, ["ORA"]));
    const evento = normalizeEvent(getVal(row, ["EVENTO"]));
    if (!evento) return;

    const igbc = toNumber(getVal(row, ["IGBC"]));
    const igbo = toNumber(getVal(row, ["IGBO"]));
    const igbt = toNumber(getVal(row, ["IGBT"]));
    const mge = toNumber(getVal(row, ["MGE"]));
    const diff = toNumber(getVal(row, ["DIFF"]));
    const pgg = toNumber(getVal(row, ["PGG"]));
    const s1s2 = toNumber(getVal(row, ["S1-S2"]));

    const ggStrongCheck =
      isFinite(igbc) && isFinite(igbo) && isFinite(pgg) && isFinite(diff) && isFinite(s1s2) &&
      igbc > 120 && igbo > 100 && pgg > 60 && diff <= 0 && s1s2 < 200;

    const ggMediumCheck =
      isFinite(igbc) && isFinite(igbo) && isFinite(pgg) && isFinite(diff) && isFinite(s1s2) &&
      igbc > 100 && igbo > 80 && pgg > 52 && diff <= 0.5 && s1s2 < 300;

    const o25StrongCheck =
      isFinite(igbt) && isFinite(mge) && isFinite(diff) && isFinite(s1s2) &&
      igbt >= 300 && mge >= 2.8 && diff <= 0 && s1s2 < 220;

    const o25MediumCheck =
      isFinite(igbt) && isFinite(mge) && isFinite(diff) && isFinite(s1s2) &&
      igbt >= 260 && mge >= 2.5 && diff <= 0.6 && s1s2 < 350;

    if (ggStrongCheck) {
      ggStrong.push({
        ora,
        evento,
        quality: "Forte",
        score: igbc + igbo + pgg - s1s2
      });
    } else if (ggMediumCheck) {
      ggMedium.push({
        ora,
        evento,
        quality: "Media",
        score: igbc + igbo + pgg - s1s2
      });
    }

    if (o25StrongCheck) {
      o25Strong.push({
        ora,
        evento,
        quality: "Forte",
        score: igbt + (mge * 20) - s1s2
      });
    } else if (o25MediumCheck) {
      o25Medium.push({
        ora,
        evento,
        quality: "Media",
        score: igbt + (mge * 20) - s1s2
      });
    }

    if (ggStrongCheck && o25StrongCheck) {
      combo.push({
        ora,
        evento,
        quality: "Forte",
        score: igbc + igbo + igbt + (mge * 20) + pgg - s1s2
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

function updateFinal() {
  const merged = [...state.over05, ...state.gg25];

  if (!merged.length) {
    boxFinal.innerHTML = `<div class="empty-state">Nessun risultato disponibile.</div>`;
    countFinal.textContent = "0 esiti";
    return;
  }

  countFinal.textContent = `${merged.length} esiti`;

  boxFinal.innerHTML = merged.map(item => `
    <div class="result-item">
      <div class="result-left">
        <div class="result-time">${escapeHtml(item.ora || "-")}</div>
        <div class="result-match">${escapeHtml(item.evento || "-")}</div>
      </div>
      <div class="result-right">
        <span class="badge ${badgeClass(item.market)}">${escapeHtml(item.market)}</span>
        <span class="badge ${item.qualityClass}">${escapeHtml(item.quality || "Forte")}</span>
      </div>
    </div>
  `).join("");
}

function renderList(container, items, label, marketClass, showQuality = false) {
  if (!items.length) {
    container.innerHTML = `<div class="empty-state">Nessun esito disponibile.</div>`;
    return;
  }

  container.innerHTML = items.map(item => `
    <div class="result-item">
      <div class="result-left">
        <div class="result-time">${escapeHtml(item.ora || "-")}</div>
        <div class="result-match">${escapeHtml(item.evento || "-")}</div>
      </div>
      <div class="result-right">
        <span class="badge ${marketClass}">${escapeHtml(label)}</span>
        ${showQuality ? `<span class="badge ${item.quality === "Forte" ? "strong" : "medium"}">${escapeHtml(item.quality)}</span>` : ""}
      </div>
    </div>
  `).join("");
}

function exportTxt() {
  const merged = [...state.over05, ...state.gg25];

  if (!merged.length) {
    alert("Nessun risultato da esportare.");
    return;
  }

  const lines = [
    "betAI - Risultati",
    "",
    ...merged.map((item, i) =>
      `${i + 1}. ${item.ora} | ${item.evento} | ${item.market} | ${item.quality || "Forte"}`
    )
  ];

  const blob = new Blob([lines.join("\n")], { type: "text/plain;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = "betai-risultati.txt";
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 800);
}

function bindPress(element, handler) {
  let touchStamp = 0;
  let locked = false;

  const run = async (event, fromTouch) => {
    if (locked) return;

    if (fromTouch) {
      touchStamp = Date.now();
    } else if (Date.now() - touchStamp < 700) {
      return;
    }

    locked = true;
    try {
      await handler(event);
    } finally {
      setTimeout(() => {
        locked = false;
      }, 250);
    }
  };

  element.addEventListener("touchend", async (event) => {
    event.preventDefault();
    await run(event, true);
  }, { passive: false });

  element.addEventListener("click", async (event) => {
    event.preventDefault();
    await run(event, false);
  });
}

function setLoading(button, isLoading, text) {
  button.disabled = isLoading;
  button.textContent = text;
}

function getVal(obj, keys) {
  for (const key of keys) {
    if (obj[key] !== undefined && obj[key] !== null && obj[key] !== "") {
      return obj[key];
    }
  }
  return "";
}

function normalizeHeader(value) {
  return cleanText(value).toUpperCase();
}

function cleanText(value) {
  return String(value ?? "").replace(/\s+/g, " ").trim();
}

function normalizeTime(value) {
  const text = cleanText(value);
  const full = text.match(/\d{2}-\d{2}-\d{4}\s+\d{2}:\d{2}/);
  if (full) return full[0];
  const hm = text.match(/\d{2}:\d{2}/);
  return hm ? hm[0] : (text || "-");
}

function normalizeEvent(value) {
  const text = cleanText(value);
  const match = text.match(/^(.*? - .*?)(?:\s{2,}.*)?$/);
  return match ? cleanText(match[1]) : text;
}

function toNumber(value) {
  if (value === null || value === undefined || value === "") return NaN;
  const match = String(value).replace(",", ".").match(/-?\d+(\.\d+)?/);
  return match ? parseFloat(match[0]) : NaN;
}

function toPercentNumber(value) {
  if (value === null || value === undefined || value === "") return NaN;
  const text = String(value).trim().replace(",", ".");
  if (text.includes("%")) {
    const n = parseFloat(text.replace("%", ""));
    return Number.isFinite(n) ? n : NaN;
  }
  const n = parseFloat(text);
  if (!Number.isFinite(n)) return NaN;
  return n <= 1 ? n * 100 : n;
}

function splitPairFlexible(value) {
  if (value === null || value === undefined || value === "") return [NaN, NaN];

  const text = String(value);
  if (text.includes("|")) {
    const parts = text.split("|");
    return [toNumber(parts[0]), toNumber(parts[1])];
  }

  const nums = text.match(/-?\d+(\.\d+)?/g);
  if (!nums || !nums.length) return [NaN, NaN];
  if (nums.length === 1) return [toNumber(nums[0]), NaN];
  return [toNumber(nums[0]), toNumber(nums[1])];
}

function sumFinite(arr) {
  return arr.filter(Number.isFinite).reduce((a, b) => a + b, 0);
}

function badgeClass(market) {
  if (market === "O0.5 PT") return "o05";
  if (market === "GG") return "gg";
  if (market === "Over 2.5") return "o25";
  return "combo";
}

function escapeHtml(value) {
  return String(value).replace(/[&<>"']/g, (char) => {
    const map = {
      "&": "&amp;",
      "<": "&lt;",
      ">": "&gt;",
      "\"": "&quot;",
      "'": "&#39;"
    };
    return map[char];
  });
}
