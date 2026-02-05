/**
 * Baixa um XLSX de um URL (OneDrive) e gera data/data.json
 *
 * ✅ Você só precisa configurar um Secret no GitHub:
 *   ONEDRIVE_XLSX_URL = link de DOWNLOAD direto do seu .xlsx
 *
 * ⚠️ Importante:
 * O link que você mandou (com /doc.aspx) geralmente é uma página de visualização.
 * Para automação, você precisa de um URL que devolva o arquivo XLSX (não HTML).
 */

import fs from "node:fs";
import path from "node:path";
import XLSX from "xlsx";

const url = process.env.ONEDRIVE_XLSX_URL;

if (!url) {
  console.error("ERRO: Secret ONEDRIVE_XLSX_URL não configurado.");
  console.error("GitHub -> Settings -> Secrets and variables -> Actions -> New repository secret");
  process.exit(1);
}

async function fetchArrayBuffer(u) {
  const res = await fetch(u, { redirect: "follow" });
  if (!res.ok) throw new Error(`Falha ao baixar XLSX: ${res.status} ${res.statusText}`);
  const ct = res.headers.get("content-type") || "";
  if (ct.includes("text/html")) {
    throw new Error(
      "O URL retornou HTML (provável página de visualização/login).\n" +
      "Use um link de DOWNLOAD direto do XLSX no Secret ONEDRIVE_XLSX_URL."
    );
  }
  return await res.arrayBuffer();
}

function brlToNumber(v) {
  if (v == null) return null;
  if (typeof v === "number") return v;
  const s = String(v).trim();
  if (!s) return null;
  const cleaned = s.replace(/R\$\s?/g, "").replace(/\./g, "").replace(/,/g, ".");
  const n = Number(cleaned);
  return Number.isFinite(n) ? n : null;
}

const monthMap = {
  "janeiro": 1,
  "fevereiro": 2,
  "março": 3,
  "marco": 3,
  "abril": 4,
  "maio": 5,
  "junho": 6,
  "julho": 7,
  "agosto": 8,
  "setembro": 9,
  "outubro": 10,
  "novembro": 11,
  "dezembro": 12
};

function normalize(s) {
  return String(s || "").trim().toLowerCase();
}

function isoDateFromMonth(monthName, year=2026) {
  const mm = monthMap[normalize(monthName)];
  if (!mm) return null;
  return `${year}-${String(mm).padStart(2,"0")}-01`;
}

function parseWorkbook(wb) {
  const out = {
    meta: { year: 2026, updatedAt: new Date().toISOString() },
    dashboard: { tabela1: [], tabela2: [], tabela3: [] },
    detalheMensal: {}
  };

  const sheetName = wb.SheetNames[0];
  const ws = wb.Sheets[sheetName];
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });

  // -------- TABELA 1 (Dashboard - linha com Entrada/Saída/Líquido/Diferença/Crescimento) --------
  let headerRow = -1;
  for (let r = 0; r < Math.min(rows.length, 80); r++) {
    const row = rows[r].map(x => String(x).trim().toLowerCase());
    const hasEntrada = row.includes("entrada");
    const hasSaida = row.includes("saída") || row.includes("saida");
    const hasLiquido = row.includes("líquido") || row.includes("liquido");
    if (hasEntrada && hasSaida && hasLiquido) { headerRow = r; break; }
  }
  if (headerRow >= 0) {
    for (let r = headerRow + 1; r < rows.length; r++) {
      const month = String(rows[r][0]).trim();
      if (!month) continue;
      const mk = normalize(month);
      if (!monthMap[mk]) {
        if (mk === "total") break;
        continue;
      }
      out.dashboard.tabela1.push({
        month,
        date: isoDateFromMonth(month, 2026),
        entrada: brlToNumber(rows[r][1]),
        saida: brlToNumber(rows[r][2]),
        liquido: brlToNumber(rows[r][3]),
        diferenca_m1: brlToNumber(rows[r][4]),
        crescimento: String(rows[r][5]).trim() || null
      });
    }
  }

  // -------- TABELA 2 e 3 (Nubank/Santander) --------
  function parseMini(title) {
    const t = title.toLowerCase();
    for (let r = 0; r < rows.length; r++) {
      const row = rows[r].map(x => String(x).trim());
      if (row.some(c => c.toLowerCase() === t)) {
        const data = [];
        for (let rr = r + 1; rr < Math.min(rows.length, r + 40); rr++) {
          const m = String(rows[rr][0]).trim();
          if (!m) continue;
          const mk = normalize(m);
          if (mk === "total") break;
          if (!monthMap[mk]) continue;
          data.push({ month: m, date: isoDateFromMonth(m, 2026), saida: brlToNumber(rows[rr][1]) });
        }
        return data;
      }
    }
    return [];
  }

  out.dashboard.tabela2 = parseMini("Nubank");
  out.dashboard.tabela3 = parseMini("Santander");

  // -------- TABELA 4 (Detalhe mensal - blocos por mês com Entrada/Saída e listas) --------
  function findMonthAnchors() {
    const anchors = [];
    for (let r = 0; r < rows.length; r++) {
      for (let c = 0; c < Math.min(rows[r].length, 250); c++) {
        const val = String(rows[r][c]).trim();
        if (!val) continue;
        if (!monthMap[normalize(val)]) continue;

        const next = (rows[r+1] || []).slice(c, c+20).map(x => String(x).trim().toLowerCase());
        const hasEntrada = next.includes("entrada");
        const hasSaida = next.includes("saída") || next.includes("saida");
        if (hasEntrada && hasSaida) anchors.push({ month: val, r, c });
      }
    }
    return anchors;
  }

  const anchors = findMonthAnchors();
  for (const a of anchors) {
    const monthName = a.month;
    const header = (rows[a.r+1] || []).slice(a.c, a.c+30).map(x => String(x).trim().toLowerCase());
    const eRel = header.indexOf("entrada");
    const sRel = header.findIndex(x => x === "saída" || x === "saida");
    if (eRel < 0 || sRel < 0) continue;

    const entradaDescCol = a.c + eRel;
    const entradaValCol  = entradaDescCol + 2;
    const saidaDescCol   = a.c + sRel;
    const saidaValCol    = saidaDescCol + 2;

    const entradas = [];
    const saidas = [];

    for (let r = a.r + 2; r < Math.min(rows.length, a.r + 60); r++) {
      const descE = String((rows[r]||[])[entradaDescCol] ?? "").trim();
      const valE  = (rows[r]||[])[entradaValCol];
      const descS = String((rows[r]||[])[saidaDescCol] ?? "").trim();
      const valS  = (rows[r]||[])[saidaValCol];

      if (normalize(descE) === "total" || normalize(descS) === "total") break;

      if (descE && normalize(descE) !== "r$") {
        const n = brlToNumber(valE);
        if (n != null) entradas.push({ desc: descE, amount: n });
      }
      if (descS && normalize(descS) !== "r$") {
        const n = brlToNumber(valS);
        if (n != null) saidas.push({ desc: descS, amount: n });
      }
    }

    saidas.sort((a,b) => (b.amount||0) - (a.amount||0));

    out.detalheMensal[monthName] = {
      date: isoDateFromMonth(monthName, 2026),
      entradas,
      saidas
    };
  }

  return out;
}

const ab = await fetchArrayBuffer(url);
const wb = XLSX.read(Buffer.from(ab), { type: "buffer" });
const data = parseWorkbook(wb);

const outPath = path.join(process.cwd(), "data", "data.json");
fs.mkdirSync(path.dirname(outPath), { recursive: true });
fs.writeFileSync(outPath, JSON.stringify(data, null, 2), "utf-8");
console.log("OK: data/data.json atualizado.");
