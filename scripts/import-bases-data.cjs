
const XLSX = require("xlsx");
const fs = require("fs");
const path = require("path");

const ROOT = path.resolve(__dirname, "..");
const SOURCE = path.join(ROOT, "data-source");
const OUTPUT = path.join(ROOT, "public", "data");

if (!fs.existsSync(OUTPUT)) {
  fs.mkdirSync(OUTPUT, { recursive: true });
}

function readExcel(fileName) {
  const filePath = path.join(SOURCE, fileName);

  if (!fs.existsSync(filePath)) {
    throw new Error(`Arquivo não encontrado: ${filePath}`);
  }

  const workbook = XLSX.readFile(filePath);
  const sheet = workbook.Sheets[workbook.SheetNames[0]];

  return XLSX.utils.sheet_to_json(sheet, {
    defval: null,
    raw: false,
  });
}

function normalizeNumber(value) {
  if (value === null || value === undefined || value === "") return 0;
  if (typeof value === "number") return value;

  let text = String(value).trim().replace(/\s/g, "");

  if (text.includes(",") && text.includes(".")) {
    text = text.replace(/\./g, "").replace(",", ".");
  } else if (text.includes(",")) {
    text = text.replace(",", ".");
  }

  const number = Number(text);
  return Number.isFinite(number) ? number : 0;
}

function normalizeQuantidadeBrasileira(value) {
  if (value === null || value === undefined || value === "") return 0;
  if (typeof value === "number") return value;

  let text = String(value).trim().replace(/\s/g, "");

  if (text.includes(".") && text.includes(",")) {
    text = text.replace(/\./g, "").replace(",", ".");
  } else if (text.includes(".")) {
    text = text.replace(/\./g, "");
  } else if (text.includes(",")) {
    text = text.replace(",", ".");
  }

  const number = Number(text);
  return Number.isFinite(number) ? number : 0;
}

function normalizeText(value) {
  if (value === null || value === undefined) return "";
  return String(value).trim();
}

function getColumn(row, ...names) {
  const normalizeHeader = (value) =>
    String(value ?? "")
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .toLowerCase()
      .replace(/[^a-z0-9]/g, "");

  const keys = Object.keys(row || {});

  for (const name of names) {
    const expected = normalizeHeader(name);
    const found = keys.find(
      (key) => normalizeHeader(key) === expected
    );

    if (found !== undefined) return row[found];
  }

  return null;
}

function convertEstoque(fileName, local) {
  const rows = readExcel(fileName);

  return rows.map((r) => ({
    produto: normalizeText(getColumn(r, "Produto")),
    desc_completa: normalizeText(
      getColumn(r, "Desc. Completa", "Desc.completa")
    ),
    qtd_fisica: normalizeNumber(
      getColumn(r, "Qtd.física", "Qtd física")
    ),
    grupo: normalizeText(getColumn(r, "Grupo")),
    familia: normalizeText(getColumn(r, "Familia", "Família")),
    local_estoque: local,
  }));
}

function convertCompras() {
  const rows = readExcel("COMPRABASES.xlsx");

  return rows.map((r) => ({
    num_oc: normalizeText(getColumn(r, "Núm.OC")),
    produto: normalizeText(getColumn(r, "Produto")),
    desc_completa: normalizeText(
      getColumn(r, "Desc. Completa", "Desc.completa")
    ),
    grupo: normalizeText(getColumn(r, "Grupo")),
    familia: normalizeText(getColumn(r, "Familia", "Família")),
    qtd_aberto: normalizeNumber(getColumn(r, "Qtd.aberto")),
    vlr_un: normalizeNumber(getColumn(r, "Vlr.un")),
    dt_prazo_ent: normalizeText(getColumn(r, "Dt.prazo ent")),
  }));
}

function convertComprasParaguai() {
  const rows = readExcel("COMPRAPARAGUAI.xlsx");

  return rows.map((r) => ({
    produto: normalizeText(getColumn(r, "Produto")),
    desc_completa: normalizeText(
      getColumn(r, "Desc. Completa", "Desc.completa")
    ),
    grupo: normalizeText(getColumn(r, "Grupo")),
    familia: normalizeText(getColumn(r, "Familia", "Família")),
    qtd_aberto: normalizeQuantidadeBrasileira(
      getColumn(r, "Qtd.aberto")
    ),
  }));
}

function convertConsumo() {
  const rows = readExcel("CONSUMOBASES.xlsx");

  return rows.map((r) => ({
    cod_prod: normalizeText(getColumn(r, "Cód.prod")),
    desc_completa: normalizeText(
      getColumn(r, "Desc. Completa", "Desc.completa")
    ),
    grupo: normalizeText(getColumn(r, "Grupo")),
    descricao: normalizeText(getColumn(r, "Descrição")),
    orig_movto: normalizeText(getColumn(r, "Orig.movto")),
    descr_orig_movto: normalizeText(
      getColumn(r, "Descr.Orig.movto")
    ),
    qtd_movimentada: normalizeNumber(
      getColumn(r, "Qtd.movimentada")
    ),
    dt_movto: normalizeText(getColumn(r, "Dt.movto")),
  }));
}

function writeJson(fileName, data) {
  const filePath = path.join(OUTPUT, fileName);

  fs.writeFileSync(
    filePath,
    JSON.stringify(data, null, 2),
    "utf8"
  );

  console.log(
    `${fileName}: ${data.length.toLocaleString("pt-BR")} registros`
  );
}

console.log("====================================");
console.log(" SMART GROUP - IMPORTAÇÃO DE BASES");
console.log("====================================");

const estoqueLocal = convertEstoque(
  "ESTOQUEBASESLOCAL.xlsx",
  "Estoque Local"
);

const estoqueParaguai = convertEstoque(
  "ESTOQUEPARAGUAI.xlsx",
  "Paraguai"
);

const estoqueEadi = convertEstoque(
  "ESTOQUEEADI.xlsx",
  "EADI"
);

const compras = convertCompras();
const comprasParaguai = convertComprasParaguai();
const consumo = convertConsumo();

writeJson("ESTOQUEBASESLOCAL.json", estoqueLocal);
writeJson("ESTOQUEPARAGUAI.json", estoqueParaguai);
writeJson("ESTOQUEEADI.json", estoqueEadi);
writeJson("COMPRABASES.json", compras);
writeJson("COMPRAPARAGUAI.json", comprasParaguai);
writeJson("CONSUMOBASES.json", consumo);

console.log("");
console.log("Importação concluída.");
