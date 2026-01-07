/* global XLSX, ExcelJS */

window.addEventListener("DOMContentLoaded", () => {
  // ==============================
  // UI
  // ==============================
  const elPresenze = document.getElementById("filePresenze");
  const btnRun     = document.getElementById("btnRun");
  const statusEl   = document.getElementById("status");
  const checksEl   = document.getElementById("checks");

  // Se anche uno solo è null, la pagina non potrà mai validare
  if (!elPresenze || !btnRun || !statusEl || !checksEl) {
    console.error("ID mancanti in index.html:", {
      filePresenze: !!elPresenze,
      btnRun: !!btnRun,
      status: !!statusEl,
      checks: !!checksEl,
    });
    if (statusEl) statusEl.textContent = "❌ Errore pagina: mancano elementi HTML (ID non trovati).";
    return;
  }

  // evita race condition: solo l’ultima validazione può abilitare/disabilitare
  let VALIDATION_SEQ = 0;

  function setStatus(msg, kind="") {
    statusEl.className = "status " + (kind || "");
    statusEl.textContent = msg;
  }
  function setChecks(lines) {
    checksEl.innerHTML = "";
    for (const l of lines) {
      const div = document.createElement("div");
      div.textContent = l;
      checksEl.appendChild(div);
    }
  }
  function readyBtn(enabled) { btnRun.disabled = !enabled; }

  // ✅ debug minimo: ti dice subito se l’evento change parte
  console.log("✅ app.js init ok");

  // ------------------------------------------------
  // QUI SOTTO incolla TUTTO il resto del tuo app.js
  // (funzioni readXlsxToAOA, validateIfPossible, btnRun click, ecc.)
  // e NON lasciare più codice fuori da questo blocco.
  // ------------------------------------------------

  // ... INCOLLA QUI IL RESTO DEL CODICE CHE AVEVI (dal normStr in giù) ...

});


/* global XLSX, ExcelJS */

// ==============================
// UI
// ==============================
const elPresenze = document.getElementById("filePresenze");
const btnRun     = document.getElementById("btnRun");
const statusEl   = document.getElementById("status");
const checksEl   = document.getElementById("checks");

// evita race condition: solo l’ultima validazione può abilitare/disabilitare
let VALIDATION_SEQ = 0;

function setStatus(msg, kind="") {
  statusEl.className = "status " + (kind || "");
  statusEl.textContent = msg;
}
function setChecks(lines) {
  checksEl.innerHTML = "";
  for (const l of lines) {
    const div = document.createElement("div");
    div.textContent = l;
    checksEl.appendChild(div);
  }
}
function readyBtn(enabled) { btnRun.disabled = !enabled; }

function normStr(v) { return (v ?? "").toString().trim(); }

// numero robusto: gestisce 1567,5 e 1.567,5
function toNum(v) {
  if (v === null || v === undefined || v === "") return null;
  if (typeof v === "number" && Number.isFinite(v)) return v;

  let s = String(v).trim();
  if (!s) return null;

  // se contiene sia . che , assumo . come separatore migliaia e , come decimali
  if (s.includes(".") && s.includes(",")) {
    s = s.replace(/\./g, "").replace(",", ".");
  } else {
    s = s.replace(",", ".");
  }

  const n = Number(s);
  return Number.isFinite(n) ? n : null;
}

function assert(cond, msg) { if (!cond) throw new Error(msg); }

function isReadyFiles() {
  return !!(elPresenze.files?.[0]);
}

function readXlsxToAOA(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onerror = () => reject(new Error("Errore lettura file: " + file.name));
    reader.onload = (e) => {
      const data = new Uint8Array(e.target.result);
      const wb = XLSX.read(data, { type: "array" });
      const sheetName = wb.SheetNames[0];
      const ws = wb.Sheets[sheetName];
      const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });
      resolve({ wb, sheetName, aoa });
    };
    reader.readAsArrayBuffer(file);
  });
}

// ==============================
// DIZIONARIO + REGOLE (come Python)
// categorie: presenza | cig | festivita | ferie | assenza_varie | infortunio | escludi
// ==============================
const CAUSALE_CATEGORY = {
  // CIG
  "CIG CALDO": "cig",
  "CIG CALDO 1 GIORNO": "cig",
  "CIG GELO": "cig",
  "CIG GELO 1 GIORNO": "cig",
  "CIG PIOGGIA": "cig",
  "CIG PIOGGIA 1 GIORNO": "cig",
  "CIG PIOGGIA CON VITTO": "cig",
  "CIG VENTO": "cig",
  "CIG VENTO 1 GIORNO": "cig",
  "CIG VENTO CON VITTO": "cig",
  "LLUVIA": "cig",
  "INTEMPERIE": "cig",

  // Festività (NEUTRE)
  "FESTIVITA": "festivita",
  "FESTIVITA 4 NOVEMBRE": "festivita",
  "FESTIVITA DOMENICA": "festivita",

  // Ferie
  "FERIE": "ferie",

  // Infortunio
  "INFORTUNIO": "infortunio",

  // Assenze varie
  "ALLATTAMENTO": "assenza_varie",
  "ASPETTATIVA": "assenza_varie",
  "ASSENZA INGIUSTIFICATA": "assenza_varie",
  "ASSENZA NON RETRIBUITA": "assenza_varie",
  "ASSEMBLEA SINDACALE": "assenza_varie",
  "CONGEDO MATRIMONIALE": "assenza_varie",
  "CONGEDO PARENTALE": "assenza_varie",
  "CONGEDO PATERNITA": "assenza_varie",
  "DONAZIONE SANGUE": "assenza_varie",
  "PATERNITA": "assenza_varie",
  "LUTTO": "assenza_varie",
  "MALATTIA": "assenza_varie",
  "PERMESSO 104": "assenza_varie",
  "PERMESSO NON RETRIBUITO": "assenza_varie",
  "PERMESSO RETRIBUITO": "assenza_varie",
  "PERMESSO SEA": "assenza_varie",
  "PERMISO JUSTIFICADO/RETRIBUIDO": "assenza_varie",
  "SOSPENSIONE CAUTELATIVA": "assenza_varie",
  "SOSPENSIONE DISCIPLINARE": "assenza_varie",
  "CASSA INTEGRAZIONE": "assenza_varie",

  // Presenza
  "CORSI": "presenza",
  "FORMAZIONE": "presenza",
  "GIORNATA DI RIPOSO": "presenza",
  "GUASTO MEZZO": "presenza",
  "LAVORO FESTIVO": "presenza",
  "LAVORO NOTTURNO": "presenza",
  "ORDINARIO": "presenza",
  "ORDINARIO GARANZIA": "presenza",
  "ORDINARIO SABATO": "presenza",
  "ORD. SABATO TRASFERTA": "presenza",
  "ORD. SABATO TRASFERTA CON VITTO": "presenza",
  "PER PREVENTIVO": "presenza",
  "RIPOSO COMPENSATIVO": "presenza",
  "SOSPENSIONE PER INIDONEITA": "presenza",
  "STRAORDINARIO": "presenza",
  "STRAORDINARIO FESTIVO": "presenza",
  "STRAORDINARIO GARANZIA": "presenza",
  "STRAORDINARIO IN TRASFERTA": "presenza",
  "STRAORDINARIO NOTTURNO": "presenza",
  "STRAORDINARIO NOTTURNO FESTIVO": "presenza",
  "TRASF CON VITTO": "presenza",
  "TRASFERTA": "presenza",
  "VIAGGIO": "presenza",
  "VISITA MEDICA": "presenza",
  "LAVORO STRAORDINARIO FESTIVO": "presenza",
  "HEURE SUPPL. > 43": "presenza",

  // Distacco
  "DISTACCO": "presenza",
  "DISTACCO ORD. SABATO": "presenza",
  "DISTACCO CON VITTO": "presenza",
  "DISTACCO CON VITTO ORD. SABATO": "presenza",
  "DISTACCO CON VITTO STRAORDINARIO": "presenza",
  "DISTACCO IN TRASFERTA": "presenza",
  "DISTACCO IN TRASFERTA ORD SABATO": "presenza",
  "DISTACCO IN TRASFERTA STRAORDINARIO": "presenza",
  "DISTACCO PER SUBAPPALTO": "presenza",
  "DISTACCO STRAORDINARIO": "presenza",
  "ENERGY DISTACCO CON VITTO": "presenza",
};

const BAD_ROWS = new Set(["TOTAL","TOTALE","GRAND TOTAL","TOTALS","TOTALE GENERALE"]);

function categorizeWithRules(cu) {
  if (cu.startsWith("CIG")) return "cig";
  if (cu.includes("LLUVIA") || cu.includes("INTEMPERIE")) return "cig";
  if (cu.startsWith("FESTIVIT")) return "festivita";
  if (cu.includes("HEURE SUPPL")) return "presenza";
  if (cu.includes("PERMISO") && cu.includes("RETRIBUIDO")) return "assenza_varie";
  if (cu.includes("DISTACCO")) return "presenza";
  if (cu.includes("PATERNIT")) return "assenza_varie";
  if (cu.includes("FORMAZION")) return "presenza";
  return null;
}

// ==============================
// VALIDAZIONE + PARSING
// ==============================
function findHeaderRow(aoa, maxRows=200) {
  for (let r=0; r<Math.min(maxRows, aoa.length); r++) {
    const row = aoa[r] || [];
    const up = row.map(x => normStr(x).toUpperCase());
    if (up.includes("RISORSA") && up.includes("CAUSALE")) return r;
  }
  return -1;
}

function findCols(aoa, headerRow) {
  const row = (aoa[headerRow] || []).map(x => normStr(x).toUpperCase());
  const risCol = row.indexOf("RISORSA");
  const cauCol = row.indexOf("CAUSALE");

  let totalCol = -1;
  const from = Math.max(0, headerRow - 20);
  for (let r=from; r<=headerRow; r++) {
    const rr = (aoa[r] || []).map(x => normStr(x).toUpperCase());
    const idx = rr.indexOf("TOTAL");
    if (idx >= 0) totalCol = idx;
  }
  if (totalCol < 0) totalCol = (aoa[headerRow] || []).length - 1;
  return { risCol, cauCol, totalCol };
}

function validatePresenzeAOA(aoa) {
  assert(Array.isArray(aoa) && aoa.length > 5, "Presenze_Base: file vuoto o non leggibile.");

  const headerRow = findHeaderRow(aoa);
  assert(headerRow >= 0, "Presenze_Base: non trovo intestazioni con 'RISORSA' e 'CAUSALE'.");

  const { risCol, cauCol, totalCol } = findCols(aoa, headerRow);
  assert(risCol >= 0 && cauCol >= 0, "Presenze_Base: colonne RISORSA/CAUSALE non trovate.");
  assert(totalCol >= 0, "Presenze_Base: colonna TOTAL non trovata.");

  // prova a leggere qualche riga valida
  let currentRis = null;
  let nValid = 0;

  for (let i = headerRow + 1; i < aoa.length; i++) {
    const row = aoa[i] || [];
    const rRaw = row[risCol];
    const cRaw = row[cauCol];
    const tRaw = row[totalCol];

    const tip = normStr(cRaw);
    const ore = toNum(tRaw);

    if (normStr(rRaw)) currentRis = normStr(rRaw).trim();
    if (!currentRis) continue;
    if (!tip) continue;
    if (BAD_ROWS.has(currentRis.toUpperCase()) || BAD_ROWS.has(tip.toUpperCase())) continue;

    if (ore !== null) nValid++;
    if (nValid >= 5) break;
  }

  assert(nValid > 0, "Presenze_Base: non trovo ore numeriche nella colonna TOTAL (o righe valide).");

  return { headerRow, ...findCols(aoa, headerRow) };
}

// Pivot: Risorsa -> Tipologia -> Ore
function parseToPivot(aoa, meta) {
  const { headerRow, risCol, cauCol, totalCol } = meta;

  const pivot = new Map();
  const dettaglioRows = [];

  let currentRis = null;

  for (let i = headerRow + 1; i < aoa.length; i++) {
    const row = aoa[i] || [];
    const rRaw = row[risCol];
    const cRaw = row[cauCol];
    const tRaw = row[totalCol];

    if (normStr(rRaw)) currentRis = normStr(rRaw).trim();

    const ris = currentRis;
    const tip = normStr(cRaw).trim();
    const ore = toNum(tRaw);

    if (!ris || !tip || ore === null) continue;
    if (BAD_ROWS.has(ris.toUpperCase()) || BAD_ROWS.has(tip.toUpperCase())) continue;

    if (!pivot.has(ris)) pivot.set(ris, new Map());
    const m = pivot.get(ris);
    m.set(tip, (m.get(tip) || 0) + ore);

    dettaglioRows.push({ Risorsa: ris, Tipologia: tip, Ore: ore });
  }

  // dettaglio aggregato (come python groupby)
  const detAgg = new Map(); // key ris||tip -> ore
  for (const r of dettaglioRows) {
    const k = r.Risorsa + "||" + r.Tipologia;
    detAgg.set(k, (detAgg.get(k) || 0) + r.Ore);
  }
  const dettaglio = Array.from(detAgg.entries()).map(([k, v]) => {
    const [Risorsa, Tipologia] = k.split("||");
    return { Risorsa, Tipologia, Ore: v };
  }).sort((a,b) => a.Risorsa.localeCompare(b.Risorsa) || a.Tipologia.localeCompare(b.Tipologia));

  // tipologie globali
  const tipiSet = new Set();
  for (const [,m] of pivot.entries()) for (const tip of m.keys()) tipiSet.add(tip);
  const tipologie = Array.from(tipiSet).sort((a,b) => a.localeCompare(b));

  const risorse = Array.from(pivot.keys()).sort((a,b) => a.localeCompare(b));

  // pivot matrix (righe risorsa, colonne tipologie)
  const pivotRows = [];
  for (const ris of risorse) {
    const row = { Risorsa: ris };
    let tot = 0;
    const m = pivot.get(ris);
    for (const tip of tipologie) {
      const v = m.get(tip) || 0;
      row[tip] = v;
      tot += v;
    }
    row["TOTALE_ANNO"] = tot;
    pivotRows.push(row);
  }

  return { risorse, tipologie, pivotRows, dettaglio };
}

function classifyTipologia(tip) {
  const cu = normStr(tip).toUpperCase();
  const byRule = categorizeWithRules(cu);
  if (byRule) return byRule;
  const byDict = CAUSALE_CATEGORY[cu];
  if (byDict) return byDict;
  return null;
}

// ==============================
// CALCOLI percentuali (ordine colonne richiesto)
// ==============================
function computePercentuali(pivotRows, tipologie) {
  // pre-classifica tipologie
  const catByTip = new Map();
  const unmapped = [];

  for (const tip of tipologie) {
    const cat = classifyTipologia(tip);
    if (!cat) unmapped.push(normStr(tip).toUpperCase());
    catByTip.set(tip, cat || "presenza"); // fallback
  }

  const out = [];
  for (const r of pivotRows) {
    const oreTot = r["TOTALE_ANNO"] || 0;

    let oreCig = 0, oreFest = 0, oreFerie = 0, oreAssVarie = 0, oreInf = 0;

    for (const tip of tipologie) {
      const v = r[tip] || 0;
      const cat = catByTip.get(tip);

      if (cat === "cig") oreCig += v;
      else if (cat === "festivita") oreFest += v;
      else if (cat === "ferie") oreFerie += v;
      else if (cat === "assenza_varie") oreAssVarie += v;
      else if (cat === "infortunio") oreInf += v;
    }

    const orePerse = oreAssVarie + oreInf + oreFerie;
    const oreLav = oreTot - oreCig - oreFest;

    const pct = (num, den) => (den && den !== 0) ? (num / den) : null;

    out.push({
      "Risorsa": r.Risorsa,
      "Ore_Totali_Anno": oreTot,
      "Ore_Lavorabili": oreLav,
      "%_Presenza_su_ore_lavorabili": oreLav ? (1 - (orePerse / oreLav)) : null,
      "%_Assenza_su_ore_lavorabili": pct(orePerse, oreLav),
      "%_CIG_su_totale": pct(oreCig, oreTot),
      "%_Festivita_su_totale": pct(oreFest, oreTot),
      "Ore_Assenze_Varie": oreAssVarie,
      "Ore_Infortunio": oreInf,
      "Ore_Ferie": oreFerie,
      "Ore_Perse_Persona": orePerse,
    });
  }

  return { percRows: out, unmapped: Array.from(new Set(unmapped)).sort() };
}

// ==============================
// EXCEL OUTPUT (ExcelJS)
// ==============================
function autoFitWorksheet(ws, minW=8, maxW=60, padding=2) {
  ws.columns.forEach(col => {
    let maxLen = 0;
    col.eachCell({ includeEmpty:true }, (cell) => {
      const v = cell.value;
      let s = "";
      if (v === null || v === undefined) s = "";
      else if (typeof v === "object" && v.richText) s = "";
      else s = String(v);
      maxLen = Math.max(maxLen, s.length);
    });
    col.width = Math.min(maxW, Math.max(minW, maxLen + padding));
  });
}

function setThinBorder(cell) {
  cell.border = {
    top:    { style: "thin" },
    left:   { style: "thin" },
    bottom: { style: "thin" },
    right:  { style: "thin" },
  };
}

function applyPercentFormatting(ws) {
  // Formatta tutte le colonne che iniziano con "%_"
  const headerRow = ws.getRow(1);
  headerRow.eachCell((cell, colNumber) => {
    const h = normStr(cell.value);
    if (h.startsWith("%_")) {
      const col = ws.getColumn(colNumber);
      col.eachCell({ includeEmpty: false }, (c, rowNumber) => {
        if (rowNumber === 1) return;
        c.numFmt = "0.00%";
      });
    }
  });
}

function freezeAtB2(ws) {
  ws.views = [{ state: "frozen", xSplit: 1, ySplit: 1 }];
}

function addLegendBox(ws, dataHeadersCount) {
  const legendTitle = "DIDASCALIA – significato colonne";
  const legendLines = [
    ["Ore_Totali_Anno", "Somma di tutte le ore registrate nell’anno per la risorsa (tutte le causali)."],
    ["Ore_Lavorabili", "Ore potenzialmente lavorabili: Ore_Totali_Anno − Ore_CIG − Ore_Festivita."],
    ["%_Presenza_su_ore_lavorabili", "1 − (Ore_Perse_Persona / Ore_Lavorabili)."],
    ["%_Assenza_su_ore_lavorabili", "Ore_Perse_Persona / Ore_Lavorabili."],
    ["%_CIG_su_totale", "Ore_CIG / Ore_Totali_Anno (solo indicatore)."],
    ["%_Festivita_su_totale", "Ore_Festivita / Ore_Totali_Anno (solo indicatore)."],
    ["Ore_Assenze_Varie", "Somma ore assenze imputabili (malattia, permessi, donazione sangue, assemblea sindacale, ecc.)."],
    ["Ore_Infortunio", "Somma ore di infortunio (tracciato separatamente)."],
    ["Ore_Ferie", "Somma ore di ferie."],
    ["Ore_Perse_Persona", "Ore perse imputabili: Ore_Assenze_Varie + Ore_Infortunio + Ore_Ferie."],
  ];

  const startCol = dataHeadersCount + 2; // 2 colonne di spazio dopo la tabella
  const c1 = startCol;
  const c2 = startCol + 1;

  // titolo unito
  ws.mergeCells(1, c1, 1, c2);
  const tCell = ws.getCell(1, c1);
  tCell.value = legendTitle;
  tCell.font = { bold: true };
  tCell.alignment = { vertical: "middle", wrapText: true };

  // header legenda
  ws.getCell(2, c1).value = "Colonna";
  ws.getCell(2, c2).value = "Descrizione";
  ws.getCell(2, c1).font = { bold: true };
  ws.getCell(2, c2).font = { bold: true };

  // righe
  let r = 3;
  for (const [name, desc] of legendLines) {
    ws.getCell(r, c1).value = name;
    ws.getCell(r, c2).value = desc;
    ws.getCell(r, c1).alignment = { vertical: "top" };
    ws.getCell(r, c2).alignment = { vertical: "top", wrapText: true };
    r++;
  }

  ws.getColumn(c1).width = 30;
  ws.getColumn(c2).width = 72;

  // bordi su tutto il riquadro
  const endRow = 2 + legendLines.length;
  for (let rr = 1; rr <= endRow; rr++) {
    setThinBorder(ws.getCell(rr, c1));
    setThinBorder(ws.getCell(rr, c2));
  }
}

async function buildWorkbookAndDownload({ percRows, pivotRows, dettaglio, risorse, tipologie }) {
  const wb = new ExcelJS.Workbook();
  wb.creator = "Presenze Report (client-side)";
  wb.created = new Date();
  wb.views = [{ activeTab: 0 }];

  // ==========================
  // 1) percentuali presenza
  // ==========================
  const ws1 = wb.addWorksheet("percentuali presenza");

  const headers1 = [
    "Risorsa",
    "Ore_Totali_Anno",
    "Ore_Lavorabili",
    "%_Presenza_su_ore_lavorabili",
    "%_Assenza_su_ore_lavorabili",
    "%_CIG_su_totale",
    "%_Festivita_su_totale",
    "Ore_Assenze_Varie",
    "Ore_Infortunio",
    "Ore_Ferie",
    "Ore_Perse_Persona",
  ];

  ws1.addRow(headers1);
  ws1.getRow(1).font = { bold: true };
  freezeAtB2(ws1);

  for (const r of percRows) {
    ws1.addRow(headers1.map(h => (r[h] ?? null)));
  }

  applyPercentFormatting(ws1);
  autoFitWorksheet(ws1);
  addLegendBox(ws1, headers1.length);

  // ==========================
  // 2) Ore_Annue_Pivot
  // ==========================
  const ws2 = wb.addWorksheet("Ore_Annue_Pivot");
  const headers2 = ["Risorsa", ...tipologie, "TOTALE_ANNO"];
  ws2.addRow(headers2);
  ws2.getRow(1).font = { bold: true };
  freezeAtB2(ws2);

  for (const r of pivotRows) {
    ws2.addRow(headers2.map(h => (r[h] ?? 0)));
  }
  autoFitWorksheet(ws2);

  // ==========================
  // 3) Dettaglio_Annuale
  // ==========================
  const ws3 = wb.addWorksheet("Dettaglio_Annuale");
  ws3.addRow(["Risorsa", "Tipologia", "Ore"]);
  ws3.getRow(1).font = { bold: true };
  ws3.views = [{ state: "frozen", ySplit: 1 }];

  for (const r of dettaglio) {
    ws3.addRow([r.Risorsa, r.Tipologia, r.Ore]);
  }
  autoFitWorksheet(ws3);

  // ==========================
  // 4) Liste
  // ==========================
  const ws4 = wb.addWorksheet("Liste");
  ws4.addRow(["Risorse", "", "Tipologie"]);
  ws4.getRow(1).font = { bold: true };
  ws4.views = [{ state: "frozen", ySplit: 1 }];

  const maxLen = Math.max(risorse.length, tipologie.length);
  for (let i = 0; i < maxLen; i++) {
    ws4.addRow([risorse[i] ?? "", "", tipologie[i] ?? ""]);
  }
  autoFitWorksheet(ws4);

  // download
  const buffer = await wb.xlsx.writeBuffer();
  const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });

  const outName = "Ore_Annue_per_Risorsa_e_Tipologia.xlsx";
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = outName;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(a.href);

  return outName;
}

// ==============================
// VALIDAZIONE LIVE
// ==============================
async function validateIfPossible() {
  const mySeq = ++VALIDATION_SEQ;

  setStatus("", "");
  setChecks([]);
  readyBtn(false);

  if (!isReadyFiles()) return;

  try {
    setStatus("Controllo file Presenze_Base...", "");

    const pres = await readXlsxToAOA(elPresenze.files[0]);
    if (mySeq !== VALIDATION_SEQ) return;

    const meta = validatePresenzeAOA(pres.aoa);

    // parsing veloce per warning causali
    const { tipologie, pivotRows } = parseToPivot(pres.aoa, meta);
    const { unmapped } = computePercentuali(pivotRows, tipologie);

    const lines = [
      "✅ Presenze_Base: struttura OK (RISORSA/CAUSALE/TOTAL)",
      `✅ Tipologie trovate: ${tipologie.length}`,
    ];
    if (unmapped.length) {
      lines.push(`⚠️ Causali non mappate (fallback=presenza): ${unmapped.length}`);
      // mostro max 12 per non spaccare la UI
      const show = unmapped.slice(0, 12);
      for (const u of show) lines.push("   - " + u);
      if (unmapped.length > show.length) lines.push(`   ... +${unmapped.length - show.length} altre`);
    } else {
      lines.push("✅ Tutte le causali sono nel dizionario/regole.");
    }

    setChecks(lines);
    setStatus("✅ File OK. Puoi generare l’Excel.", "ok");
    readyBtn(true);
  } catch (e) {
    if (mySeq !== VALIDATION_SEQ) return;
    setChecks([]);
    setStatus("❌ " + (e?.message || String(e)), "err");
    readyBtn(false);
  }
}

elPresenze.addEventListener("change", validateIfPossible);

// ==============================
// RUN
// ==============================
btnRun.addEventListener("click", async () => {
  try {
    readyBtn(false);
    setStatus("Elaboro...", "");

    const pres = await readXlsxToAOA(elPresenze.files[0]);
    const meta = validatePresenzeAOA(pres.aoa);

    const parsed = parseToPivot(pres.aoa, meta);
    const { percRows, unmapped } = computePercentuali(parsed.pivotRows, parsed.tipologie);

    // se vuoi bloccare su causali non mappate invece di warning, basta:
    // assert(unmapped.length === 0, "Ci sono causali non mappate: " + unmapped.join(", "));

    const outName = await buildWorkbookAndDownload({
      percRows,
      pivotRows: parsed.pivotRows,
      dettaglio: parsed.dettaglio,
      risorse: parsed.risorse,
      tipologie: parsed.tipologie,
    });

    setStatus(`✅ Creato: ${outName}`, "ok");
    readyBtn(true);
  } catch (e) {
    console.error(e);
    setStatus("❌ Errore: " + (e?.message || String(e)), "err");
    readyBtn(false);
  }
});
