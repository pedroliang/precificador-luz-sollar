// Lê a planilha de precificação (Google Sheets exportado como .xlsx) e devolve
// a estrutura que o painel consome. O layout NÃO é fixo: as colunas são
// descobertas pelos rótulos do cabeçalho e, principalmente, pelas próprias
// fórmulas de preço mínimo e preço de venda — é delas que sai quais colunas são
// taxa (%), quais são custo fixo (R$) e qual é o lucro alvo.

/** A biblioteca de leitura de .xlsx entra por parâmetro: no Supabase vem do npm,
 *  no navegador vem do CDN. Assim este arquivo roda igual nos dois. */
export interface LeitorXlsx {
  read(dados: Uint8Array, opcoes: Record<string, unknown>): {
    SheetNames: string[];
    Sheets: Record<string, Record<string, unknown>>;
  };
}

export interface Loja {
  id: string;
  nome: string;
  plataforma: string;
  defaults: Record<string, number>;
  fixos: { key: string; label: string }[];
  canais: { key: string; label: string }[];
  rows: LinhaLoja[];
}
export interface LinhaLoja {
  sku: string;
  nome: string;
  cr: unknown[] | null;
  cv: number | null;
  desc: number;
  eov: number | null;
  rates: Record<string, number>;
  promoC: number | null;
  rebC: number | null;
  promoP: number | null;
  rebP: number | null;
  ov: Record<string, number>;
  fixos: Record<string, number>;
  mkt: (number | null)[];
  obs: string;
  r: number;
  [k: string]: unknown;
}
export interface Resultado {
  produtos: Record<string, [string, number | null]>;
  stores: Loja[];
  avisos: string[];
  abas: { nome: string; tipo: string; linhas: number }[];
}

/* ---------- utilidades ---------- */

const norm = (s: unknown) =>
  String(s ?? "")
    .normalize("NFD").replace(/[\u0300-\u036f]/g, "")
    .toUpperCase().replace(/\s+/g, " ").trim();

const slug = (s: unknown) =>
  norm(s).toLowerCase().replace(/[^a-z0-9]+/g, "_").replace(/^_|_$/g, "");

function colName(c: number): string {
  let s = "";
  c += 1;
  while (c > 0) { const m = (c - 1) % 26; s = String.fromCharCode(65 + m) + s; c = ((c - m) / 26) | 0; }
  return s;
}
function colIndex(name: string): number {
  let n = 0;
  for (const ch of name.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n - 1;
}

/** Números malformados na planilha (ex.: "1,100,00") viram null e são reportados. */
export const numerosInvalidos: { texto: string; onde: string }[] = [];
let ondeAtual = "";

function parseBRL(v: unknown): number | null {
  if (v == null) return null;
  if (typeof v === "number") return Number.isFinite(v) ? round6(v) : null;
  const bruto = String(v).trim();
  if (!bruto || bruto === "-") return null;
  const s = bruto.replace(/R\$/gi, "").replace(/[\s\u00a0]/g, "").replace(/%/g, "");
  let n = NaN;
  if (/^-?\d{1,3}(\.\d{3})+(,\d+)?$/.test(s)) n = parseFloat(s.replace(/\./g, "").replace(",", "."));
  else if (/^-?\d+,\d+$/.test(s)) n = parseFloat(s.replace(",", "."));
  else if (/^-?\d{1,3}(,\d{3})+(\.\d+)?$/.test(s)) n = parseFloat(s.replace(/,/g, ""));
  else if (/^-?\d+(\.\d+)?$/.test(s)) n = parseFloat(s);
  if (!Number.isFinite(n)) {
    const erroDeFormula = bruto.startsWith("#");
    if (/\d/.test(s) && !erroDeFormula && numerosInvalidos.length < 200) {
      numerosInvalidos.push({ texto: bruto, onde: ondeAtual });
    }
    return null;
  }
  return round6(n);
}
const round6 = (n: number) => Math.round(n * 1e6) / 1e6;

/** Todas as referências de coluna (A, AB…) que aparecem num trecho de fórmula. */
function refCols(expr: string): Set<number> {
  const out = new Set<number>();
  for (const m of expr.matchAll(/\$?([A-Z]{1,3})\$?\d+/g)) out.add(colIndex(m[1]));
  return out;
}

/* ---------- acesso à planilha ---------- */

interface Cel { v?: unknown; f?: string; z?: string; t?: string }
type Planilha = Record<string, unknown> & { "!ref"?: string };

class Aba {
  ws: Planilha;
  nRows: number;
  nCols: number;
  constructor(ws: Planilha) {
    this.ws = ws;
    const ref = String(ws["!ref"] ?? "A1:A1");
    const fim = ref.split(":").pop() ?? "A1";
    const m = fim.match(/^([A-Z]+)(\d+)$/);
    this.nRows = m ? parseInt(m[2], 10) : 1;
    this.nCols = m ? colIndex(m[1]) + 1 : 1;
  }
  cel(r: number, c: number): Cel | null {
    return (this.ws[colName(c) + (r + 1)] as Cel) ?? null;
  }
  val(r: number, c: number): unknown {
    return this.cel(r, c)?.v ?? null;
  }
  /** valor só quando digitado à mão (a coluna normalmente tem fórmula) */
  literal(r: number, c: number): number | null {
    const cel = this.cel(r, c);
    if (!cel || cel.f) return null;
    return parseBRL(cel.v);
  }
  formula(r: number, c: number): string | null {
    return this.cel(r, c)?.f ?? null;
  }
  ehPercent(r: number, c: number): boolean {
    return (this.cel(r, c)?.z ?? "").includes("%");
  }
}

/* ---------- descoberta do cabeçalho ---------- */

interface Mapa {
  headerRow: number;
  colSku: number;
  colProd: number;
  colCusto: number;
  colDesc: number;
  colCustoLiq: number;
  taxas: number[];
  fixos: number[];
  lucro: number;
  canais: { comissao: number; campos: Record<string, number> }[];
  mercado: number[];
  obs: number;
}

const CAMPO_POR_ROTULO: Record<string, string> = {
  "PRECO MINIMO": "min", "PRECO MÍNIMO": "min",
  "PRECO DE VENDA": "ven", "PRECO VENDA": "ven",
  "PRECO ORIGINAL": "ori",
  "PROMOCAO ATIVA": "promo", "PROMOCAO": "promo",
  "LUCRO": "lucro", "REBATE": "rebate", "LUCRO FINAL": "lucroFinal", "MARGEM": "margem",
};

function acharHeader(a: Aba): number {
  for (let r = 0; r < Math.min(8, a.nRows); r++) {
    for (let c = 0; c < Math.min(4, a.nCols); c++) {
      if (norm(a.val(r, c)) === "SKU" && norm(a.val(r, c + 1)) === "PRODUTO") return r;
    }
  }
  return -1;
}

function mapear(a: Aba, hr: number): Mapa | null {
  const rot: string[] = [];
  for (let c = 0; c < a.nCols; c++) rot.push(norm(a.val(hr, c)));
  const acha = (nome: string) => rot.indexOf(nome);

  const colSku = acha("SKU"), colProd = acha("PRODUTO"), colCusto = acha("CUSTO");
  if (colSku < 0 || colProd < 0 || colCusto < 0) return null;
  const colDesc = acha("DESCONTO");
  let colCustoLiq = rot.findIndex((x) => x.startsWith("CUSTO -") || x.startsWith("CUSTO-"));
  if (colCustoLiq < 0) colCustoLiq = colCusto;

  // blocos de canal: cada "PRECO MINIMO" (ou o primeiro "PRECO DE VENDA" sem mínimo) abre um bloco
  const blocos: { comissao: number; campos: Record<string, number> }[] = [];
  let atual: Record<string, number> | null = null;
  const mercado: number[] = [];
  let obs = -1;
  for (let c = colCustoLiq + 1; c < a.nCols; c++) {
    const l = rot[c];
    if (!l) continue;
    if (l === "OBSERVACAO") { obs = c; continue; }
    if (l === "PRECO MEDIO" || /^PRECO \d+$/.test(l)) { mercado.push(c); continue; }
    const campo = CAMPO_POR_ROTULO[l];
    if (!campo) continue;
    if (campo === "min" || (campo === "ven" && (!atual || "ven" in atual))) {
      atual = {}; blocos.push({ comissao: -1, campos: atual });
    }
    if (atual) atual[campo] = c;
  }
  if (!blocos.length) return null;
  const primeiroBloco = Math.min(...blocos.map((b) => Math.min(...Object.values(b.campos))));

  // colunas de taxa / custo fixo: tudo entre "custo -x%" e o primeiro bloco
  const taxas: number[] = [], fixos: number[] = [];
  for (let c = colCustoLiq + 1; c < primeiroBloco; c++) {
    let temDado = false, percent = false;
    for (let r = hr; r < Math.min(hr + 40, a.nRows); r++) {
      const cel = a.cel(r, c);
      if (!cel || cel.v == null || cel.v === "" || typeof cel.v === "string") continue;
      temDado = true;
      if ((cel.z ?? "").includes("%")) { percent = true; break; }
    }
    if (!temDado) continue;
    (percent ? taxas : fixos).push(c);
  }

  // as fórmulas dizem a verdade: mínimo x venda diferem exatamente pela coluna do lucro
  let lucro = -1;
  const linhaRef = acharLinhaComFormula(a, hr, blocos[0]);
  if (linhaRef >= 0) {
    const b0 = blocos[0].campos;
    const fMin = b0.min != null ? a.formula(linhaRef, b0.min) : null;
    const fVen = b0.ven != null ? a.formula(linhaRef, b0.ven) : null;
    if (fMin && fVen) {
      const dMin = denominador(fMin), dVen = denominador(fVen);
      for (const c of dVen) if (!dMin.has(c)) lucro = c;
      const num = numerador(fMin);
      num.delete(colCustoLiq);
      const numArr = [...num].filter((c) => c > colCustoLiq && c < primeiroBloco);
      if (numArr.length) { fixos.length = 0; fixos.push(...numArr.sort((x, y) => x - y)); }
      const taxasArr = [...dMin].filter((c) => c > colCustoLiq && c < primeiroBloco);
      if (taxasArr.length) { taxas.length = 0; taxas.push(...taxasArr.sort((x, y) => x - y)); }
    }
    // comissão de cada canal = a coluna de taxa que só aparece na fórmula daquele canal
    const porCanal = blocos.map((b) => {
      const f = b.campos.min != null ? a.formula(linhaRef, b.campos.min) : null;
      return f ? denominador(f) : new Set<number>();
    });
    const comuns = new Set<number>([...(porCanal[0] ?? [])]);
    porCanal.forEach((s) => { for (const c of [...comuns]) if (!s.has(c)) comuns.delete(c); });
    blocos.forEach((b, i) => {
      const so = [...(porCanal[i] ?? [])].filter((c) => !comuns.has(c));
      b.comissao = so.length ? so[0] : -1;
      if (b.comissao >= 0 && !taxas.includes(b.comissao)) taxas.push(b.comissao);
    });
    for (const c of comuns) if (c !== lucro && !taxas.includes(c)) taxas.push(c);
  }
  if (blocos.length === 1 && blocos[0].comissao < 0 && linhaRef >= 0) {
    // canal único: a comissão é a taxa cujo valor está na linha, não no cabeçalho
    for (const c of taxas) {
      if (c === lucro) continue;
      const cab = a.cel(hr, c);
      if (!cab || typeof cab.v !== "number") { blocos[0].comissao = c; break; }
    }
  }
  if (lucro >= 0 && !taxas.includes(lucro)) taxas.push(lucro);
  taxas.sort((x, y) => x - y);

  return { headerRow: hr, colSku, colProd, colCusto, colDesc, colCustoLiq, taxas, fixos, lucro, canais: blocos, mercado, obs };
}

function acharLinhaComFormula(a: Aba, hr: number, bloco: { campos: Record<string, number> }): number {
  const c = bloco.campos.min ?? bloco.campos.ven;
  if (c == null) return -1;
  for (let r = hr + 1; r < Math.min(hr + 200, a.nRows); r++) if (a.formula(r, c)) return r;
  return -1;
}
function denominador(f: string): Set<number> {
  const i = f.lastIndexOf("/");
  return refCols(i < 0 ? f : f.slice(i));
}
function numerador(f: string): Set<number> {
  const i = f.lastIndexOf("/");
  return refCols(i < 0 ? f : f.slice(0, i));
}

/* ---------- leitura de uma aba de loja ---------- */

const ROTULO_KEY: Record<string, string> = {
  NF: "nf", T: "t", ADS: "ads", LUCRO: "lucro", DEVOLUCAO: "dev", WEDO: "dev",
  CLASSICO: "cls", PREMIUM: "prm", COMISSAO: "com", COMISSAO_: "com",
};
const CANAL_LABEL: Record<string, string> = { cls: "Clássico", prm: "Premium", com: "Padrão" };
// chaves estáveis para os custos fixos mais comuns, para os ajustes salvos não se perderem
const KEY_FIXO: Record<string, string> = {
  "FRETE/FIXO": "frete", "FRETE": "frete", "FIXO": "fixo", "TAXA": "taxa",
  "TAXA FIXA": "fixo", "FRETE FIXO": "frete",
};
const LABEL_FIXO: Record<string, string> = {
  frete: "Frete/Fixo", fixo: "Taxa fixa (R$)", taxa: "Taxa (R$)",
};

function rotuloColuna(a: Aba, hr: number, c: number): string {
  for (let r = hr - 1; r >= 0 && r >= hr - 2; r--) {
    const v = norm(a.val(r, c));
    if (v) return v;
  }
  const v = norm(a.val(hr, c));
  return v || colName(c);
}
function chaveTaxa(rotulo: string, usadas: Set<string>): string {
  let k = ROTULO_KEY[rotulo] ?? slug(rotulo);
  if (!k) k = "taxa";
  let n = 2, base = k;
  while (usadas.has(k)) k = base + "_" + n++;
  usadas.add(k);
  return k;
}

function lerLoja(nome: string, a: Aba, m: Mapa, avisos: string[]): Loja {
  ondeAtual = nome;
  const hr = m.headerRow;
  const usadas = new Set<string>();
  const taxaKey = new Map<number, string>();
  for (const c of m.taxas) taxaKey.set(c, chaveTaxa(rotuloColuna(a, hr, c), usadas));
  if (m.lucro >= 0) taxaKey.set(m.lucro, "lucro");

  const usadasFix = new Set<string>();
  const fixoKey = new Map<number, string>();
  const fixos: { key: string; label: string }[] = [];
  for (const c of m.fixos) {
    const rotHeader = String(a.val(hr, c) ?? "").trim();
    const util = rotHeader && !/^[-–—\d.,%]*$/.test(rotHeader);
    const label = util ? rotHeader : rotuloColuna(a, hr, c);
    let k = KEY_FIXO[norm(label)] ?? slug(label) ?? "fixo";
    if (!k) k = "fixo";
    let n = 2; const base = k;
    while (usadasFix.has(k)) k = base + "_" + n++;
    usadasFix.add(k);
    fixoKey.set(c, k);
    fixos.push({ key: k, label: LABEL_FIXO[k] ?? label });
  }

  const canais = m.canais.map((b, i) => {
    const k = b.comissao >= 0 ? (taxaKey.get(b.comissao) ?? "com") : "com" + (i || "");
    return { key: k, label: CANAL_LABEL[k] ?? (rotuloColuna(a, hr, b.comissao) || "Canal " + (i + 1)) };
  });

  const defaults: Record<string, number> = {};
  for (const [c, k] of taxaKey) {
    const cel = a.cel(hr, c);
    if (cel && typeof cel.v === "number") defaults[k] = round6(cel.v);
  }

  const rows: LinhaLoja[] = [];
  const idxPorLinha = new Map<number, number>();
  for (let r = hr + 1; r < a.nRows; r++) {
    const skuRaw = a.val(r, m.colSku);
    const prod = a.val(r, m.colProd);
    const fCusto = a.formula(r, m.colCusto);
    const vCusto = a.val(r, m.colCusto);
    if (prod == null && fCusto == null && vCusto == null) continue;
    idxPorLinha.set(r, rows.length);

    const rates: Record<string, number> = {};
    for (const [c, k] of taxaKey) {
      const cel = a.cel(r, c);
      if (!cel) continue;
      if (cel.f && /^\$[A-Z]+\$\d+$/.test(cel.f)) continue; // herda o global
      if (typeof cel.v !== "number") continue;
      const v = round6(cel.v);
      if (defaults[k] == null || v !== defaults[k]) rates[k] = v;
    }
    for (const b of m.canais) {
      if (b.comissao < 0) continue;
      const k = taxaKey.get(b.comissao)!;
      if (rates[k] == null && defaults[k] == null) {
        const cel = a.cel(r, b.comissao);
        if (cel && typeof cel.v === "number") rates[k] = round6(cel.v);
      }
    }

    const ov: Record<string, number> = {};
    m.canais.forEach((b, i) => {
      const ck = canais[i].key;
      for (const [campo, sufixo] of [["min", "_min"], ["ven", "_ven"], ["ori", "_ori"]] as const) {
        const c = b.campos[campo];
        if (c == null) continue;
        const v = a.literal(r, c);
        if (v != null) ov[ck + sufixo] = v;
      }
    });

    const linha: LinhaLoja = {
      sku: normSku(skuRaw),
      nome: prod == null ? "" : String(prod),
      cr: regraCusto(fCusto, vCusto),
      cv: parseBRL(vCusto),
      desc: m.colDesc >= 0 ? (parseBRL(a.val(r, m.colDesc)) ?? 0) : 0,
      eov: m.colCustoLiq !== m.colCusto ? a.literal(r, m.colCustoLiq) : null,
      rates,
      promoC: null, rebC: null, promoP: null, rebP: null,
      ov,
      mkt: [0, 1, 2, 3, 4, 5].map((i) => (m.mercado[i] != null ? parseBRL(a.val(r, m.mercado[i])) : null)),
      obs: m.obs >= 0 ? String(a.val(r, m.obs) ?? "") : "",
      fixos: {},
      r: r + 1,
    };
    m.canais.forEach((b, i) => {
      const promo = b.campos.promo != null ? parseBRL(a.val(r, b.campos.promo)) : null;
      const reb = b.campos.rebate != null ? parseBRL(a.val(r, b.campos.rebate)) : null;
      if (i === 0) { linha.promoC = promo; linha.rebC = reb; }
      else if (i === 1) { linha.promoP = promo; linha.rebP = reb; }
    });
    for (const [c, k] of fixoKey) (linha.fixos as Record<string, number>)[k] = parseBRL(a.val(r, c)) ?? 0;
    rows.push(linha);
  }

  // referências de fórmula (C15*10) viram índices do array
  for (const linha of rows) {
    const cr = linha.cr as unknown[] | null;
    if (!cr) continue;
    if (cr[0] === "M") cr[1] = idxPorLinha.get((cr[1] as number) - 1) ?? -1;
    else if (cr[0] === "S") {
      cr[1] = idxPorLinha.get((cr[1] as number) - 1) ?? -1;
      cr[2] = idxPorLinha.get((cr[2] as number) - 1) ?? -1;
    }
  }

  if (!canais.length) avisos.push(`A aba "${nome}" não tem colunas de preço reconhecidas.`);

  return {
    id: nome,
    nome: nomeAmigavel(nome),
    plataforma: plataforma(nome),
    defaults,
    fixos,
    canais,
    rows,
  };
}

function normSku(v: unknown): string {
  if (v == null) return "";
  if (typeof v === "number") return Number.isInteger(v) ? String(v) : String(v);
  return String(v).trim();
}
function regraCusto(f: string | null, v: unknown): unknown[] | null {
  if (f) {
    const s = f.replace(/\s/g, "");
    if (/^VLOOKUP/i.test(s)) return ["L"];
    let m = s.match(/^([A-Z]{1,3})(\d+)\*([\d.]+)$/);
    if (m) return ["M", parseInt(m[2], 10), parseFloat(m[3])];
    m = s.match(/^([A-Z]{1,3})(\d+)\+([A-Z]{1,3})(\d+)$/);
    if (m) return ["S", parseInt(m[2], 10), parseInt(m[4], 10)];
    return null;
  }
  const n = parseBRL(v);
  return n == null ? null : ["F", n];
}
const NOMES_PROPRIOS: Record<string, string> = { Tiktok: "TikTok", Meli: "Meli", Shopee: "Shopee" };
function nomeAmigavel(id: string): string {
  const base = id.replace(/^GERAL\s+MELI\s+/i, "").replace(/^GERAL\s+/i, "").trim()
    .toLowerCase().replace(/(^|\s|-)([a-zà-ú])/g, (_, a, b) => a + b.toUpperCase());
  return base.split(" ").map((w) => NOMES_PROPRIOS[w] ?? w).join(" ");
}
function plataforma(id: string): string {
  const n = norm(id);
  if (n.includes("MELI") || n.includes("MERCADO LIVRE")) return "meli";
  if (n.includes("SHOPEE")) return "shopee";
  if (n.includes("TIKTOK")) return "tiktok";
  if (n.includes("AMAZON")) return "amazon";
  if (n.includes("MAGALU")) return "magalu";
  return "outro";
}

/* ---------- catálogo de custos ---------- */

function lerCatalogo(a: Aba): Record<string, [string, number | null]> | null {
  let hr = -1, cSku = -1, cDesc = -1, cCusto = -1;
  for (let r = 0; r < Math.min(6, a.nRows); r++) {
    for (let c = 0; c < Math.min(6, a.nCols); c++) {
      const l = norm(a.val(r, c));
      if (l === "DESCRICAO") { hr = r; cDesc = c; }
      if (hr === r && l === "CUSTO") cCusto = c;
    }
    if (hr === r && cDesc >= 0 && cCusto >= 0) { cSku = Math.max(0, cDesc - 1); break; }
    hr = -1; cDesc = -1; cCusto = -1;
  }
  if (hr < 0 || cCusto < 0) return null;
  const out: Record<string, [string, number | null]> = {};
  for (let r = hr + 1; r < a.nRows; r++) {
    const sku = normSku(a.val(r, cSku));
    if (!sku) continue;
    out[sku] = [String(a.val(r, cDesc) ?? ""), parseBRL(a.val(r, cCusto))];
  }
  return Object.keys(out).length ? out : null;
}

/* ---------- entrada principal ---------- */

/** Só o que é preciso para descobrir o tipo da aba, sem carregar a planilha toda. */
function pareceCatalogo(a: Aba): boolean {
  for (let r = 0; r < Math.min(6, a.nRows); r++) {
    let temDesc = false, temCusto = false;
    for (let c = 0; c < Math.min(6, a.nCols); c++) {
      const l = norm(a.val(r, c));
      if (l === "DESCRICAO") temDesc = true;
      if (l === "CUSTO") temCusto = true;
    }
    if (temDesc && temCusto) return true;
  }
  return false;
}

const OPCOES_LEITURA = {
  cellFormula: true, cellNF: true, cellText: false, cellHTML: false,
  cellStyles: false, cellDates: false, sheetStubs: false,
} as const;

export function parsePlanilha(buf: Uint8Array, XLSX: LeitorXlsx): Resultado {
  const produtos: Record<string, [string, number | null]> = {};
  const stores: Loja[] = [];
  const avisos: string[] = [];
  const abas: { nome: string; tipo: string; linhas: number }[] = [];
  const candidatos: { nome: string; cat: Record<string, [string, number | null]> }[] = [];

  numerosInvalidos.length = 0;

  // 1ª passada: só as primeiras linhas de cada aba, para saber o que é o quê.
  // Carregar as 21 abas inteiras de uma vez estoura a memória da Edge Function.
  const tipos = new Map<string, "loja" | "catalogo" | "ignorada">();
  {
    const cabecalhos = XLSX.read(buf, { ...OPCOES_LEITURA, cellFormula: false, sheetRows: 8 });
    for (const nome of cabecalhos.SheetNames) {
      const a = new Aba(cabecalhos.Sheets[nome]);
      if (acharHeader(a) >= 0) tipos.set(nome, "loja");
      else if (pareceCatalogo(a)) tipos.set(nome, "catalogo");
      else tipos.set(nome, "ignorada");
    }
  }

  // 2ª passada: uma aba de cada vez, liberando a anterior antes de abrir a próxima.
  for (const [nome, tipo] of tipos) {
    if (tipo === "ignorada") { abas.push({ nome, tipo: "ignorada", linhas: 0 }); continue; }

    const wb = XLSX.read(buf, { ...OPCOES_LEITURA, sheets: [nome] });
    const ws = wb.Sheets[nome];
    if (!ws) { abas.push({ nome, tipo: "ignorada", linhas: 0 }); continue; }
    const a = new Aba(ws);

    if (tipo === "loja") {
      const hr = acharHeader(a);
      const m = hr >= 0 ? mapear(a, hr) : null;
      if (m && m.canais.length) {
        const loja = lerLoja(nome, a, m, avisos);
        stores.push(loja);
        abas.push({ nome, tipo: "loja", linhas: loja.rows.length });
        continue;
      }
    }
    const cat = lerCatalogo(a);
    if (cat && Object.keys(cat).length >= 20) {
      candidatos.push({ nome, cat });
      abas.push({ nome, tipo: "catalogo", linhas: Object.keys(cat).length });
      continue;
    }
    abas.push({ nome, tipo: "ignorada", linhas: 0 });
  }

  candidatos.sort((x, y) => Object.keys(y.cat).length - Object.keys(x.cat).length);
  if (candidatos.length) {
    Object.assign(produtos, candidatos[0].cat);
    for (const extra of candidatos.slice(1)) {
      const i = abas.findIndex((x) => x.nome === extra.nome);
      if (i >= 0) { abas[i].tipo = "ignorada"; abas[i].linhas = 0; }
    }
  }

  if (!stores.length) avisos.push("Nenhuma aba de loja foi reconhecida nesta planilha.");
  if (!Object.keys(produtos).length) avisos.push("Nenhuma aba de catálogo de custos foi reconhecida.");
  if (numerosInvalidos.length) {
    const ex = numerosInvalidos.slice(0, 3).map((x) => `"${x.texto}" (${x.onde})`).join(", ");
    avisos.push(`${numerosInvalidos.length} célula(s) com número em formato inválido foram ignoradas — ex.: ${ex}.`);
  }

  return { produtos, stores, avisos, abas };
}
