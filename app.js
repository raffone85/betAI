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

const isIPhoneSafari =
  /iPhone|iPad|iPod/i.test(navigator.userAgent) &&
  /Safari/i.test(navigator.userAgent) &&
  !/CriOS|FxiOS|EdgiOS/i.test(navigator.userAgent);

const state = {
  over05: [],
  gg25: []
};

const HEADER_SCHEMAS = {
  O05: {
    required: ["ORA", "EVENTO", "IND", "DELTA"],
    aliases: {
      ORA: ["ORA"],
      EVENTO: ["EVENTO"],
      RPT: ["RPT", "RIS", "RIS."],
      IND: ["IND"],
      QROV05PT: ["QROV0.5PT", "QROV05PT"],
      DELTA: ["DELTA"],
      U5CO: ["U5[C|O]", "U5CO", "U5C|O"],
      U5CTCO: ["U5CT[C|O]", "U5CTCO", "U5CTC|O"],
      DIFF: ["DIFF"],
      MGE: ["MGE"],
      COV05: ["COV0.5%", "COV05", "COV05%"],
      OOV05: ["OOV0.5%", "OOV05", "OOV05%"],
      QRGG: ["QRGG"],
      QRO25: ["QRO25", "QROV25", "QOV25"]
    }
  },
  GG25: {
    required: ["ORA", "NAZIONE", "EVENTO", "IGBC", "IGBO", "IGBT"],
    aliases: {
      ORA: ["ORA"],
      NAZIONE: ["NAZIONE"],
      EVENTO: ["EVENTO"],
      RIS: ["RIS.", "RIS"],
      IGBC: ["IGBC"],
      IGBO: ["IGBO"],
      IGBT: ["IGBT"],
      MGE: ["MGE"],
      DIFF: ["DIFF"],
      PGG: ["PGG"],
      S1S2: ["S1-S2", "S1S2"],
      QOV25: ["QOV2.5", "QOV25", "QOV2_5"]
    }
  }
};

fileO05.addEventListener("click", () => {
  fileO05.value = "";
});

fileGG25.addEventListener("click", () => {
  fileGG25.value = "";
});

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

btnO05.addEventListener("click", analyzeO05);
btnGG25.addEventListener("click", analyzeGG25);
btnExport.addEventListener("click", exportTxt);

async function analyzeO05() {
  if (!selectedO05) {
    statusO05.textContent = "Seleziona prima un file CSV o Excel.";
    return;
  }

  setLoading(btnO05, true, "Analisi in corso...");
  statusO05.textContent = "Lettura file in corso...";

  try {
    const rows = await readRowsFromFile(selectedO05, "O05");
    const results = buildOver05Results(rows);

    renderList(boxO05Premium, results.premium, "O0.5 PT", "o05");
    renderList(boxO05A, results.listaA, "O0.5 PT", "o05");
    renderList(boxO05B, results.listaB, "O0.5 PT", "o05");

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
    statusGG25.textContent = "Seleziona prima un file CSV o Excel.";
    return;
  }

  setLoading(btnGG25, true, "Analisi in corso...");
  statusGG25.textContent = "Lettura file in corso...";

  try {
    const rows = await readRowsFromFile(selectedGG25, "GG25");
    const results = buildGG25Results(rows);

    renderList(boxGGStrong, results.ggStrong, "GG", "gg");
    renderList(boxGGMedium, results.ggMedium, "GG", "gg");
    renderList(boxO25Strong, results.o25Strong, "Over 2.5", "o25");
    renderList(boxO25Medium, results.o25Medium, "Over 2.5", "o25");
    renderList(boxCombo, results.combo, "Combo", "combo");

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

async function readRowsFromFile(file, schemaName) {
  const ext = getExtension(file.name);

  if (ext === "csv" || ext === "txt") {
    const text = await readFileAsText(file);
    const matrix = parseCSV(text);
    return matrixToObjects(matrix, schemaName);
  }

  if (ext === "xlsx" || ext === "xls") {
    if (isIPhoneSafari) {
      throw new Error("Su iPhone usa il CSV esportato da Excel, non il file .xlsx.");
    }
    if (typeof XLSX === "undefined") {
      throw new Error("Libreria XLSX non caricata.");
    }
    const matrix = await readExcelMatrix(file);
    return matrixToObjects(matrix, schemaName);
  }

  throw new Error("Formato non supportato. Usa CSV su iPhone oppure XLSX/CSV su PC.");
}

async function readExcelMatrix(file) {
  const ab = await readFileAsArrayBuffer(file);
  const workbook = XLSX.read(new Uint8Array(ab), { type: "array" });
  const sheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];

  return XLSX.utils.sheet_to_json(sheet, {
    header: 1,
    defval: "",
    raw: false,
    blankrows: false
  });
}

function matrixToObjects(matrix, schemaName) {
  if (!matrix || !matrix.length) {
    throw new Error("File vuoto o non leggibile.");
  }

  const schema = HEADER_SCHEMAS[schemaName];
  const headerIndex = findBestHeaderRow(matrix, schema);

  if (headerIndex === -1) {
    const preview = matrix
      .slice(0, 6)
      .map(row => (row || []).map(cell => cleanText(cell)).join(" | "))
      .join(" || ");
    throw new Error("Intestazioni non trovate. Preview: " + preview);
  }

  const rawHeaderRow = matrix[headerIndex] || [];
  const headers = rawHeaderRow.map(cell => canonicalHeader(cell));
  const rows = [];

  for (let i = headerIndex + 1; i < matrix.length; i++) {
    const row = matrix[i];
    if (!row || row.every(cell => cleanText(cell) === "")) continue;

    const obj = {};
    headers.forEach((header, index) => {
      if (!header) return;
      obj[header] = row[index];
    });

    if (Object.keys(obj).length) rows.push(obj);
  }

  return rows;
}

function findBestHeaderRow(matrix, schema) {
  let bestIndex = -1;
  let bestScore = -1;

  for (let i = 0; i < Math.min(matrix.length, 20); i++) {
    const score = scoreHeaderRow(matrix[i] || [], schema);
    if (score > bestScore) {
      bestScore = score;
      bestIndex = i;
    }
  }

  return bestScore >= 6 ? bestIndex : -1;
}

function scoreHeaderRow(row, schema) {
  const normalizedCells = row.map(cell => canonicalHeader(cell)).filter(Boolean);
  let score = 0;

  schema.required.forEach(req => {
    const aliases = (schema.aliases[req] || [req]).map(canonicalHeader);
    if (aliases.some(alias => normalizedCells.includes(alias))) score += 2;
  });

  Object.keys(schema.aliases).forEach(key => {
    const aliases = schema.aliases[key].map(canonicalHeader);
    if (aliases.some(alias => normalizedCells.includes(alias))) score += 1;
  });

  return score;
}

function buildOver05Results(rows) {
  const premium = [];
  const listaA = [];
  const listaB = [];

  rows.forEach((row) => {
    const ora = normalizeTime(getValue(row, ["ORA"]));
    const evento = normalizeEvent(getValue(row, ["EVENTO"]));
    if (!evento) return;

    const ind = toNumber(getValue(row, ["IND"]));
    const delta = toPercentNumber(getValue(row, ["DELTA"]));
    const diff = toNumber(getValue(row, ["DIFF"]));
    const mge = toNumber(getValue(row, ["MGE"]));
    const cov = toPercentNumber(getValue(row, ["COV05"]));
    const oov = toPercentNumber(getValue(row, ["OOV05"]));
    const qrgg = toNumber(getValue(row, ["QRGG"]));
    const qro25 = toNumber(getValue(row, ["QRO25"]));
    const qro05 = toNumber(getValue(row, ["QROV05PT"]));
    const u5 = splitPairFlexible(getValue(row, ["U5CO"]));
    const u5ct = splitPairFlexible(getValue(row, ["U5CTCO"]));

    let score = 0;

    if (isFinite(ind)) {
      if (ind <= 280) score += 2.4;
      else if (ind <= 320) score += 2;
      else if (ind <= 350) score += 1.5;
      else if (ind <= 390) score += 0.9;
    }

    if (isFinite(delta)) {
      if (delta <= -20) score += 2.1;
      else if (delta <= -10) score += 1.5;
      else if (delta < 0) score += 0.8;
      else if (delta <= 5) score += 0.2;
    }

    if (isFinite(diff)) {
      if (diff <= -1) score += 1.8;
      else if (diff < 0) score += 1.3;
      else if (diff <= 0.2) score += 0.8;
      else if (diff <= 0.6) score += 0.3;
    }

    if (isFinite(mge)) {
      if (mge >= 3.4) score += 2;
      else if (mge >= 3.0) score += 1.5;
      else if (mge >= 2.7) score += 1;
      else if (mge >= 2.4) score += 0.5;
    }

    if (isFinite(cov)) {
      if (cov >= 85) score += 1.3;
      else if (cov >= 78) score += 1;
      else if (cov >= 72) score += 0.6;
      else if (cov >= 65) score += 0.2;
    }

    if (isFinite(oov)) {
      if (oov >= 85) score += 1.3;
      else if (oov >= 78) score += 1;
      else if (oov >= 72) score += 0.6;
      else if (oov >= 65) score += 0.2;
    }

    if (isFinite(qro05)) {
      if (qro05 <= 1.20) score += 1.2;
      else if (qro05 <= 1.25) score += 0.9;
      else if (qro05 <= 1.32) score += 0.5;
      else if (qro05 <= 1.40) score += 0.2;
    }

    if (isFinite(qrgg) && isFinite(qro25)) {
      if (qrgg <= qro25) score += 0.6;
      if (qrgg <= 1.60 && qro25 <= 1.70) score += 0.6;
    }

    const u5sum = sumFinite(u5);
    const u5ctsum = sumFinite(u5ct);

    if (u5sum >= 6) score += 0.7;
    else if (u5sum >= 4) score += 0.3;

    if (u5ctsum >= 4) score += 0.8;
    else if (u5ctsum >= 3) score += 0.4;

    const premiumCheck =
      isFinite(ind) && isFinite(delta) && isFinite(diff) && isFinite(mge) &&
      ind <= 330 &&
      delta <= -8 &&
      diff <= 0.2 &&
      mge >= 2.7;

    const item = {
      ora,
      evento,
      quality: "Media",
      score
    };

    if (premiumCheck || score >= 7.1) {
      item.quality = "Forte";
      premium.push(item);
    } else if (score >= 5.6) {
      item.quality = "Forte";
      listaA.push(item);
    } else if (score >= 4.2) {
      item.quality = "Media";
      listaB.push(item);
    }
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
    const ora = normalizeTime(getValue(row, ["ORA"]));
    const evento = normalizeEvent(getValue(row, ["EVENTO"]));
    if (!evento) return;

    const igbc = toNumber(getValue(row, ["IGBC"]));
    const igbo = toNumber(getValue(row, ["IGBO"]));
    const igbt = toNumber(getValue(row, ["IGBT"]));
    const mge = toNumber(getValue(row, ["MGE"]));
    const diff = toNumber(getValue(row, ["DIFF"]));
    const pgg = toNumber(getValue(row, ["PGG"]));
    const s1s2 = toNumber(getValue(row, ["S1S2"]));

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

function renderList(container, items, label, marketClass) {
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

function setLoading(button, isLoading, text) {
  button.disabled = isLoading;
  button.textContent = text;
}

function getExtension(name) {
  const parts = String(name || "").toLowerCase().split(".");
  return parts.length > 1 ? parts.pop() : "";
}

function getValue(obj, keys) {
  for (const key of keys) {
    if (obj[key] !== undefined && obj[key] !== null && cleanText(obj[key]) !== "") {
      return obj[key];
    }
  }
  return "";
}

function canonicalHeader(value) {
  return cleanText(value)
    .replace(/^\uFEFF/, "")
    .normalize("NFKD")
    .replace(/[\u0300-\u036f]/g, "")
    .toUpperCase()
    .replace(/\s+/g, "")
    .replace(/[^A-Z0-9.\-%\\[\\]\|]/g, "");
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
  let text = cleanText(value);
  if (!text) return "";

  text = text.replace(/\s{2,}/g, " ");
  if (!text.includes(" - ")) return text;

  const dashIndex = text.indexOf(" - ");
  const left = cleanText(text.slice(0, dashIndex));
  let right = cleanText(text.slice(dashIndex + 3));

  if (!left || !right) return cleanText(text);

  const leftEscaped = escapeRegExp(left);
  const repeatedLeft = new RegExp(`\\b${leftEscaped}\\b`, "i");
  const repeatedLeftMatch = right.match(repeatedLeft);

  if (repeatedLeftMatch && repeatedLeftMatch.index > 0) {
    right = cleanText(right.slice(0, repeatedLeftMatch.index));
  }

  right = trimRepeatedTail(right);

  return cleanText(`${left} - ${right}`);
}

function trimRepeatedTail(text) {
  const tokens = cleanText(text).split(" ").filter(Boolean);
  if (tokens.length < 2) return cleanText(text);

  for (let size = Math.floor(tokens.length / 2); size >= 1; size--) {
    const a = tokens.slice(tokens.length - size).join(" ").toLowerCase();
    const b = tokens.slice(tokens.length - size * 2, tokens.length - size).join(" ").toLowerCase();
    if (a && a === b) {
      return tokens.slice(0, tokens.length - size).join(" ");
    }
  }

  return cleanText(text);
}

function toNumber(value) {
  if (value === null || value === undefined || value === "") return NaN;
  const text = String(value).trim().replace(/\s/g, "").replace(",", ".");
  const match = text.match(/-?\d+(\.\d+)?/);
  return match ? parseFloat(match[0]) : NaN;
}

function toPercentNumber(value) {
  if (value === null || value === undefined || value === "") return NaN;
  const text = String(value).trim().replace(/\s/g, "").replace(",", ".");
  const hasPercent = text.includes("%");
  const n = parseFloat(text.replace("%", ""));
  if (!Number.isFinite(n)) return NaN;
  if (hasPercent) return n;
  return n <= 1 ? n * 100 : n;
}

function splitPairFlexible(value) {
  if (value === null || value === undefined || value === "") return [NaN, NaN];

  const text = String(value).replace(/;/g, "|");
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

function escapeRegExp(value) {
  return String(value).replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function readFileAsArrayBuffer(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = e => resolve(e.target.result);
    reader.onerror = () => reject(new Error("Impossibile leggere il file."));
    reader.readAsArrayBuffer(file);
  });
}

function readFileAsText(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = e => resolve(String(e.target.result || ""));
    reader.onerror = () => reject(new Error("Impossibile leggere il CSV."));
    reader.readAsText(file, "utf-8");
  });
}

function parseCSV(text) {
  const clean = String(text || "").replace(/^\uFEFF/, "").replace(/\r\n/g, "\n").replace(/\r/g, "\n");
  const lines = clean.split("\n").filter(line => line.trim() !== "");
  const sample = lines.slice(0, 5).join("\n");

  const commaCount = (sample.match(/,/g) || []).length;
  const semicolonCount = (sample.match(/;/g) || []).length;
  const tabCount = (sample.match(/\t/g) || []).length;

  let delimiter = ",";
  if (semicolonCount >= commaCount && semicolonCount >= tabCount) delimiter = ";";
  else if (tabCount > commaCount && tabCount > semicolonCount) delimiter = "\t";

  return splitCSVRows(clean, delimiter);
}

function splitCSVRows(text, delimiter) {
  const rows = [];
  let row = [];
  let cell = "";
  let inQuotes = false;

  for (let i = 0; i < text.length; i++) {
    const ch = text[i];
    const next = text[i + 1];

    if (ch === '"') {
      if (inQuotes && next === '"') {
        cell += '"';
        i++;
      } else {
        inQuotes = !inQuotes;
      }
      continue;
    }

    if (!inQuotes && ch === delimiter) {
      row.push(cell);
      cell = "";
      continue;
    }

    if (!inQuotes && ch === "\n") {
      row.push(cell);
      if (row.some(item => cleanText(item) !== "")) rows.push(row);
      row = [];
      cell = "";
      continue;
    }

    cell += ch;
  }

  if (cell.length || row.length) {
    row.push(cell);
    if (row.some(item => cleanText(item) !== "")) rows.push(row);
  }

  return rows;
}
