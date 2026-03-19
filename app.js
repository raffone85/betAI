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

const GREEN_MAX_TOTAL = 10;
const GREEN_SECTION_LIMITS = {
  combo: 3,
  ggStrong: 4,
  o25Strong: 4,
  ggMedium: 2,
  o25Medium: 2
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
    const rawResults = buildGG25Results(rows);
    const limited = limitGreenResults(rawResults);

    renderList(boxCombo, limited.combo, "Combo", "combo");
    renderList(boxGGStrong, limited.ggStrong, "GG", "gg");
    renderList(boxO25Strong, limited.o25Strong, "Over 2.5", "o25");
    renderList(boxGGMedium, limited.ggMedium, "GG", "gg");
    renderList(boxO25Medium, limited.o25Medium, "Over 2.5", "o25");

    const merged = [
      ...limited.combo.map(x => ({ ...x, market: "Combo", qualityClass: "strong", rank: 0 })),
      ...limited.ggStrong.map(x => ({ ...x, market: "GG", qualityClass: "strong", rank: 1 })),
      ...limited.o25Strong.map(x => ({ ...x, market: "Over 2.5", qualityClass: "strong", rank: 2 })),
      ...limited.ggMedium.map(x => ({ ...x, market: "GG", qualityClass: "medium", rank: 3 })),
      ...limited.o25Medium.map(x => ({ ...x, market: "Over 2.5", qualityClass: "medium", rank: 4 }))
    ].sort((a, b) => a.rank - b.rank || (b.score || 0) - (a.score || 0));

    state.gg25 = merged;
    countGG25.textContent = `${merged.length} esiti`;
    updateFinal();

    statusGG25.textContent =
      `Analisi completata. Righe lette: ${rows.length} | Candidati grezzi: ${limited.rawTotal} | Selezionati: ${limited.total} | Combo: ${limited.combo.length} | GG Forte: ${limited.ggStrong.length} | O2.5 Forte: ${limited.o25Strong.length} | GG Medio: ${limited.ggMedium.length} | O2.5 Medio: ${limited.o25Medium.length}`;
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
    const qov25 = toNumber(getValue(row, ["QOV25"]));

    if (!isFinite(igbc) && !isFinite(igbo) && !isFinite(igbt)) return;

    let ggScore = 0;
    let o25Score = 0;

    if (isFinite(igbc)) {
      if (igbc >= 220) ggScore += 2.2;
      else if (igbc >= 170) ggScore += 1.7;
      else if (igbc >= 130) ggScore += 1.2;
      else if (igbc >= 100) ggScore += 0.7;
    }

    if (isFinite(igbo)) {
      if (igbo >= 150) ggScore += 2.1;
      else if (igbo >= 115) ggScore += 1.6;
      else if (igbo >= 90) ggScore += 1.1;
      else if (igbo >= 70) ggScore += 0.6;
    }

    if (isFinite(pgg)) {
      if (pgg >= 70) ggScore += 1.8;
      else if (pgg >= 62) ggScore += 1.3;
      else if (pgg >= 55) ggScore += 0.9;
      else if (pgg >= 48) ggScore += 0.4;
    }

    if (isFinite(diff)) {
      if (diff <= -1.2) {
        ggScore += 1.6;
        o25Score += 1.5;
      } else if (diff <= -0.4) {
        ggScore += 1.1;
        o25Score += 1.1;
      } else if (diff <= 0.2) {
        ggScore += 0.7;
        o25Score += 0.7;
      } else if (diff <= 0.8) {
        ggScore += 0.3;
        o25Score += 0.3;
      }
    }

    if (isFinite(s1s2)) {
      if (s1s2 <= 90) {
        ggScore += 1.8;
        o25Score += 1.6;
      } else if (s1s2 <= 160) {
        ggScore += 1.2;
        o25Score += 1.1;
      } else if (s1s2 <= 260) {
        ggScore += 0.6;
        o25Score += 0.6;
      } else if (s1s2 <= 360) {
        ggScore += 0.2;
        o25Score += 0.2;
      }
    }

    if (isFinite(igbt)) {
      if (igbt >= 400) o25Score += 2.5;
      else if (igbt >= 340) o25Score += 1.9;
      else if (igbt >= 290) o25Score += 1.4;
      else if (igbt >= 240) o25Score += 0.9;
      else if (igbt >= 200) o25Score += 0.4;
    }

    if (isFinite(mge)) {
      if (mge >= 3.5) o25Score += 2;
      else if (mge >= 3.0) o25Score += 1.5;
      else if (mge >= 2.7) o25Score += 1;
      else if (mge >= 2.45) o25Score += 0.5;
    } else {
      if (isFinite(igbt) && igbt >= 280) o25Score += 0.3;
      if (isFinite(pgg) && pgg >= 60) ggScore += 0.2;
    }

    if (isFinite(qov25)) {
      if (qov25 <= 1.45) o25Score += 1.1;
      else if (qov25 <= 1.60) o25Score += 0.8;
      else if (qov25 <= 1.80) o25Score += 0.5;
      else if (qov25 <= 2.00) o25Score += 0.2;
    }

    const ggStrongCheck =
      ggScore >= 5.4 &&
      isFinite(igbc) && igbc >= 120 &&
      isFinite(igbo) && igbo >= 85 &&
      isFinite(pgg) && pgg >= 52 &&
      (!isFinite(s1s2) || s1s2 <= 260);

    const ggMediumCheck =
      !ggStrongCheck &&
      ggScore >= 4.0 &&
      (
        (isFinite(igbc) && igbc >= 105) ||
        (isFinite(igbo) && igbo >= 80)
      ) &&
      (!isFinite(s1s2) || s1s2 <= 420);

    const o25StrongCheck =
      o25Score >= 5.2 &&
      isFinite(igbt) && igbt >= 260 &&
      (!isFinite(mge) || mge >= 2.6) &&
      (!isFinite(s1s2) || s1s2 <= 320);

    const o25MediumCheck =
      !o25StrongCheck &&
      o25Score >= 4.0 &&
      isFinite(igbt) && igbt >= 220 &&
      (!isFinite(s1s2) || s1s2 <= 450);

    if (ggStrongCheck) {
      ggStrong.push({
        ora,
        evento,
        quality: "Forte",
        score: ggScore
      });
    } else if (ggMediumCheck) {
      ggMedium.push({
        ora,
        evento,
        quality: "Media",
        score: ggScore
      });
    }

    if (o25StrongCheck) {
      o25Strong.push({
        ora,
        evento,
        quality: "Forte",
        score: o25Score
      });
    } else if (o25MediumCheck) {
      o25Medium.push({
        ora,
        evento,
        quality: "Media",
        score: o25Score
      });
    }

    const comboScore = ggScore + o25Score;
    const comboCheck =
      (
        (ggStrongCheck && o25StrongCheck) ||
        (comboScore >= 10 && (ggStrongCheck || ggMediumCheck) && (o25StrongCheck || o25MediumCheck))
      );

    if (comboCheck) {
      combo.push({
        ora,
        evento,
        quality: comboScore >= 11 ? "Forte" : "Media",
        score: comboScore
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

function limitGreenResults(results) {
  const selected = {
    combo: [],
    ggStrong: [],
    ggMedium: [],
    o25Strong: [],
    o25Medium: []
  };

  const rawTotal =
    results.combo.length +
    results.ggStrong.length +
    results.ggMedium.length +
    results.o25Strong.length +
    results.o25Medium.length;

  const seen = new Set();
  let total = 0;

  const priorityOrder = [
    "combo",
    "ggStrong",
    "o25Strong",
    "ggMedium",
    "o25Medium"
  ];

  for (const key of priorityOrder) {
    for (const item of results[key]) {
      if (total >= GREEN_MAX_TOTAL) break;

      const eventKey = normalizeKey(item.evento);
      if (seen.has(eventKey)) continue;
      if (selected[key].length >= GREEN_SECTION_LIMITS[key]) continue;

      selected[key].push(item);
      seen.add(eventKey);
      total++;
    }
  }

  return {
    ...selected,
    rawTotal,
    total
  };
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

  const allTimes = text.match(/\d{2}:\d{2}/g);
  if (allTimes && allTimes.length) return allTimes[allTimes.length - 1];

  return text || "-";
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
  right = dedupeWordTail(right);

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

function dedupeWordTail(text) {
  const tokens = cleanText(text).split(" ").filter(Boolean);
  if (tokens.length < 3) return cleanText(text);

  const out = [];
  for (let i = 0; i < tokens.length; i++) {
    if (i > 0 && tokens[i].toLowerCase() === tokens[i - 1].toLowerCase()) continue;
    out.push(tokens[i]);
  }
  return out.join(" ");
}

function normalizeKey(text) {
  return cleanText(text)
    .toLowerCase()
    .normalize("NFKD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[^a-z0-9]+/g, "-");
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
