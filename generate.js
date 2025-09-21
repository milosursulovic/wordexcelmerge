import fs from "node:fs/promises";
import fssync from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import XLSX from "xlsx";
import PizZip from "pizzip";
import Docxtemplater from "docxtemplater";
import dayjs from "dayjs";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const TEMPLATE_PATH = path.resolve(__dirname, "Template.docx");
const ULAZ_XLSX = path.resolve(__dirname, "Ulaz.xlsx");
const SIFARNIK_XLSX = path.resolve(__dirname, "Sifarnik.xlsx");
const OUTPUT_DIR = path.resolve(__dirname, "out");

const SHEET_ULAZ = "Podaci";
const SHEET_SIFARNIK = "RM";

const deburrMap = {
  š: "s",
  đ: "d",
  ž: "z",
  č: "c",
  ć: "c",
  Š: "s",
  Đ: "d",
  Ž: "z",
  Č: "c",
  Ć: "c",
};
function deburrSR(s) {
  return String(s).replace(/[šđžčćŠĐŽČĆ]/g, (m) => deburrMap[m] || m);
}

function norm(v) {
  return String(v ?? "")
    .replace(/\u00A0/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}
function nkey(v) {
  return deburrSR(norm(v)).toLowerCase();
}
function guessColumns(headerRow, wantedNames) {
  const map = {};
  for (const [logical, aliases] of Object.entries(wantedNames)) {
    const target = aliases.map(nkey);
    const found = headerRow.find((h) => target.includes(nkey(h)));
    map[logical] = found || null;
  }
  return map;
}
function readSheetAsJsonWithHeader(xlsxPath, sheetName) {
  const wb = XLSX.readFile(xlsxPath, { cellDates: false });
  const ws = wb.Sheets[sheetName || wb.SheetNames[0]];
  if (!ws) return { header: [], rows: [] };
  const rows = XLSX.utils.sheet_to_json(ws, {
    defval: "",
    raw: false,
    header: 1,
  });
  if (!rows.length) return { header: [], rows: [] };
  const [header, ...body] = rows;
  const objs = body.map((r) => {
    const o = {};
    header.forEach((h, i) => (o[h] = r[i] ?? ""));
    return o;
  });
  return { header, rows: objs };
}
function padTo(sampleMap) {
  let maxLen = 0;
  for (const k of sampleMap.keys()) maxLen = Math.max(maxLen, norm(k).length);
  return (s) => {
    const x = norm(s);
    if (x.length >= maxLen) return x;
    return /^\d+$/.test(x) ? x.padStart(maxLen, "0") : x;
  };
}
function sanitizeFilename(name) {
  return String(name)
    .replace(/[\\/:*?"<>|]+/g, "_")
    .trim();
}

const UL_COL_ALIASES = {
  ime: ["ime"],
  prezime: ["prezime"],
  jmbg: ["jmbg"],
  sifraRM: ["sifrarm", "sifra rm", "sifra", "šifra", "šifra rm", "code"],
  opisRM: [
    "opisrm",
    "opis",
    "opis posla",
    "opis radnog mesta",
    "opisradnogmesta",
  ],
};
const SIF_COL_ALIASES = {
  sifra: ["sifrarm", "sifra", "šifra", "code", "rm", "sifra rm"],
  opis: ["opisrm", "opis", "opis posla", "opis radnog mesta"],
};

async function renderOne(templateBuffer, data, outPath) {
  const zip = new PizZip(templateBuffer);
  const doc = new Docxtemplater(zip, {
    paragraphLoop: true,
    linebreaks: true,
    delimiters: { start: "[[", end: "]]" },
  });
  doc.render(data);
  const buf = doc.getZip().generate({ type: "nodebuffer" });
  await fs.writeFile(outPath, buf);
}

(async () => {
  await fs.mkdir(OUTPUT_DIR, { recursive: true });

  const { header: sh, rows: srows } = readSheetAsJsonWithHeader(
    SIFARNIK_XLSX,
    SHEET_SIFARNIK
  );
  if (!srows.length)
    throw new Error("Sifarnik.xlsx je prazan ili sheet ne postoji");
  const sCols = guessColumns(sh, SIF_COL_ALIASES);
  if (!sCols.sifra || !sCols.opis) {
    console.warn("Header šifarnika:", sh);
    throw new Error(
      "Nisam našao kolone šifarnika (tražim npr. 'SifraRM' i 'OpisRM')."
    );
  }
  const sifraToOpis = new Map();
  for (const r of srows) {
    const k = norm(r[sCols.sifra]);
    const v = norm(r[sCols.opis]);
    if (k) sifraToOpis.set(k, v);
  }
  const pad = padTo(sifraToOpis);

  const { header: uh, rows: urows } = readSheetAsJsonWithHeader(
    ULAZ_XLSX,
    SHEET_ULAZ
  );
  if (!urows.length)
    throw new Error("Ulaz.xlsx je prazan ili sheet ne postoji");
  const uCols = guessColumns(uh, UL_COL_ALIASES);
  if (!uCols.ime || !uCols.prezime || !uCols.jmbg || !uCols.sifraRM) {
    console.warn("Header ulaza:", uh);
    throw new Error(
      "Nisam našao ključne kolone u Ulaz.xlsx (Ime, Prezime, JMBG, SifraRM)."
    );
  }

  const templateBuf = fssync.readFileSync(TEMPLATE_PATH);
  let ok = 0,
    miss = 0;

  for (const row of urows) {
    const Ime = norm(row[uCols.ime]);
    const Prezime = norm(row[uCols.prezime]);
    const JMBG = norm(row[uCols.jmbg]);

    let OpisRadnogMesta = norm(row[uCols.opisRM]);

    if (!OpisRadnogMesta) {
      const rawKey = norm(row[uCols.sifraRM]);
      const paddedKey = pad(rawKey);
      OpisRadnogMesta =
        sifraToOpis.get(rawKey) || sifraToOpis.get(paddedKey) || "";
      if (!OpisRadnogMesta) {
        console.warn(
          `[MISS] ${Ime} ${Prezime} | SifraRM="${rawKey}" (padded="${paddedKey}") nije nađena u šifarniku`
        );
        miss++;
      } else {
        ok++;
      }
    } else {
      ok++;
    }

    const DatumCell = row["Datum"] ?? row["datum"] ?? "";
    const Datum = norm(DatumCell) || dayjs().format("DD.MM.YYYY.");

    const data = {
      Ime,
      Prezime,
      JMBG,
      OpisRadnogMesta,
      OpisRM: OpisRadnogMesta,
      Datum,
    };

    const fname =
      sanitizeFilename(`${Ime || "BezImena"} ${Prezime || ""}`) || "Zaposleni";
    const outPath = path.join(OUTPUT_DIR, `${fname} - generisano.docx`);
    await renderOne(templateBuf, data, outPath);
    console.log("✔", outPath, "| OpisRM:", OpisRadnogMesta || "(prazno)");
  }

  console.log(
    `Gotovo. Pogodaka opisa: ${ok}, Promašaja (bez opisa): ${miss}. Izlaz: ${OUTPUT_DIR}`
  );
})().catch((e) => {
  if (e.properties?.errors) {
    e.properties.errors.forEach((err) =>
      console.error("TEMPLATE:", err.properties?.explanation || err.message)
    );
  }
  console.error("Greška:", e.message);
  process.exitCode = 1;
});
