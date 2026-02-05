/**
 * Fix 403 (OneDrive) - adiciona headers "browser-like" e estratégia de retry.
 * Substitua o seu arquivo: scripts/sync_onedrive_to_datajson.mjs
 *
 * Obs: Se ainda retornar 403, o link NÃO está realmente público para requisições server-side
 * e a solução recomendada é usar Power Automate para publicar data.json no GitHub.
 */

import fs from "node:fs";
import path from "node:path";
import XLSX from "xlsx";

const url = process.env.ONEDRIVE_XLSX_URL;

if (!url) {
  console.error("ERRO: Secret ONEDRIVE_XLSX_URL não configurado.");
  process.exit(1);
}

function withDownload(u) {
  // garante download=1
  const has = /[?&]download=1/.test(u);
  if (has) return u;
  return u + (u.includes("?") ? "&" : "?") + "download=1";
}

async function fetchArrayBuffer(u) {
  const candidates = [u, withDownload(u)];
  let lastErr = null;

  for (const cu of candidates) {
    for (let attempt = 1; attempt <= 3; attempt++) {
      try {
        const res = await fetch(cu, {
          redirect: "follow",
          headers: {
            // OneDrive às vezes bloqueia user-agent vazio/"node"
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/121.0 Safari/537.36",
            "Accept": "*/*",
            "Accept-Language": "pt-BR,pt;q=0.9,en;q=0.8"
          }
        });

        if (!res.ok) {
          // em 403, espere um pouco e tente de novo
          if (res.status === 403 && attempt < 3) {
            await new Promise(r => setTimeout(r, 1200 * attempt));
            continue;
          }
          throw new Error(`Falha ao baixar XLSX: ${res.status} ${res.statusText}`);
        }

        const ct = res.headers.get("content-type") || "";
        if (ct.includes("text/html")) {
          throw new Error(
            "O URL retornou HTML (visualização/login). Use um link de download direto e público."
          );
        }
        return await res.arrayBuffer();
      } catch (e) {
        lastErr = e;
      }
    }
  }
  throw lastErr ?? new Error("Falha ao baixar XLSX.");
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
  "janeiro": 1, "fevereiro": 2, "março": 3, "marco": 3, "abril": 4, "maio": 5, "junho": 6,
  "julho": 7, "agosto": 8, "setembro": 9, "outubro": 10, "novembro": 11, "dezembro": 12
};

function normalize(s) { return String(s||"").trim().toLowerCase(); }
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

  // Tabela1
  let headerRow = -1;
  for (let r=0;r<Math.min(rows.length,80);r++){
    const row = rows[r].map(x=>String(x).trim().toLowerCase());
    const hasE=row.includes("entrada");
    const hasS=row.includes("saída")||row.includes("saida");
    const hasL=row.includes("líquido")||row.includes("liquido");
    if (hasE&&hasS&&hasL){ headerRow=r; break; }
  }
  if (headerRow>=0){
    for (let r=headerRow+1;r<rows.length;r++){
      const month=String(rows[r][0]).trim();
      if(!month) continue;
      const mk=normalize(month);
      if(!monthMap[mk]){ if(mk==="total") break; continue; }
      out.dashboard.tabela1.push({
        month, date: isoDateFromMonth(month, 2026),
        entrada: brlToNumber(rows[r][1]),
        saida: brlToNumber(rows[r][2]),
        liquido: brlToNumber(rows[r][3]),
        diferenca_m1: brlToNumber(rows[r][4]),
        crescimento: String(rows[r][5]).trim() || null
      });
    }
  }

  // Tabela2/3
  function parseMini(title){
    const t=title.toLowerCase();
    for(let r=0;r<rows.length;r++){
      const row=rows[r].map(x=>String(x).trim());
      if(row.some(c=>c.toLowerCase()===t)){
        const data=[];
        for(let rr=r+1;rr<Math.min(rows.length,r+40);rr++){
          const m=String(rows[rr][0]).trim();
          if(!m) continue;
          const mk=normalize(m);
          if(mk==="total") break;
          if(!monthMap[mk]) continue;
          data.push({month:m, date: isoDateFromMonth(m, 2026), saida: brlToNumber(rows[rr][1])});
        }
        return data;
      }
    }
    return [];
  }
  out.dashboard.tabela2=parseMini("Nubank");
  out.dashboard.tabela3=parseMini("Santander");

  // Tabela4 (blocos por mês)
  function findMonthAnchors(){
    const anchors=[];
    for(let r=0;r<rows.length;r++){
      for(let c=0;c<Math.min(rows[r].length,250);c++){
        const val=String(rows[r][c]).trim();
        if(!val) continue;
        if(!monthMap[normalize(val)]) continue;
        const next=(rows[r+1]||[]).slice(c,c+20).map(x=>String(x).trim().toLowerCase());
        const hasE=next.includes("entrada");
        const hasS=next.includes("saída")||next.includes("saida");
        if(hasE&&hasS) anchors.push({month:val,r,c});
      }
    }
    return anchors;
  }

  for (const a of findMonthAnchors()){
    const header=(rows[a.r+1]||[]).slice(a.c,a.c+30).map(x=>String(x).trim().toLowerCase());
    const eRel=header.indexOf("entrada");
    const sRel=header.findIndex(x=>x==="saída"||x==="saida");
    if(eRel<0||sRel<0) continue;

    const entradaDescCol=a.c+eRel, entradaValCol=entradaDescCol+2;
    const saidaDescCol=a.c+sRel, saidaValCol=saidaDescCol+2;

    const entradas=[], saidas=[];
    for(let r=a.r+2;r<Math.min(rows.length,a.r+60);r++){
      const descE=String((rows[r]||[])[entradaDescCol]??"").trim();
      const valE=(rows[r]||[])[entradaValCol];
      const descS=String((rows[r]||[])[saidaDescCol]??"").trim();
      const valS=(rows[r]||[])[saidaValCol];

      if(normalize(descE)==="total"||normalize(descS)==="total") break;

      if(descE && normalize(descE)!=="r$"){
        const n=brlToNumber(valE);
        if(n!=null) entradas.push({desc:descE, amount:n});
      }
      if(descS && normalize(descS)!=="r$"){
        const n=brlToNumber(valS);
        if(n!=null) saidas.push({desc:descS, amount:n});
      }
    }
    saidas.sort((a,b)=>(b.amount||0)-(a.amount||0));
    out.detalheMensal[a.month]={ date: isoDateFromMonth(a.month, 2026), entradas, saidas };
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
