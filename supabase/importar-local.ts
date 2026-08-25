// Importador de reserva, para rodar no seu computador quando você não quiser
// (ou não puder) usar a Edge Function. Faz exatamente a mesma coisa.
//
//   deno run -A supabase/importar-local.ts "<link da planilha>"
//
// Antes, defina as duas variáveis de ambiente (Project Settings → API):
//   SUPABASE_URL=https://seu-projeto.supabase.co
//   SUPABASE_SERVICE_ROLE_KEY=...        (chave secreta — nunca versione)

import * as XLSX from "npm:xlsx@0.18.5";
import { parsePlanilha } from "./functions/importar/parser.ts";

const URL_BASE = Deno.env.get("SUPABASE_URL");
const CHAVE = Deno.env.get("SUPABASE_SERVICE_ROLE_KEY");
const entrada = Deno.args[0];

if (!URL_BASE || !CHAVE) {
  console.error("Defina SUPABASE_URL e SUPABASE_SERVICE_ROLE_KEY no ambiente.");
  Deno.exit(1);
}
if (!entrada) {
  console.error('Uso: deno run -A supabase/importar-local.ts "<link da planilha>"');
  Deno.exit(1);
}

const id = entrada.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]{20,})/)?.[1] ??
  (/^[a-zA-Z0-9-_]{20,}$/.test(entrada.trim()) ? entrada.trim() : null);
if (!id) {
  console.error("Não reconheci esse link do Google Sheets.");
  Deno.exit(1);
}

console.log("Baixando a planilha…");
const r = await fetch(`https://docs.google.com/spreadsheets/d/${id}/export?format=xlsx`, { redirect: "follow" });
const buf = new Uint8Array(await r.arrayBuffer());
if (!(buf[0] === 0x50 && buf[1] === 0x4b)) {
  console.error(
    "A planilha não está pública. Em Compartilhar → Acesso geral, escolha " +
      '"Qualquer pessoa com o link" como Leitor.',
  );
  Deno.exit(1);
}
console.log(`Baixados ${Math.round(buf.length / 1024)} KB. Lendo as abas…`);

const dados = parsePlanilha(buf, XLSX as never);
for (const a of dados.abas) console.log(`  ${a.tipo.padEnd(9)} ${a.nome} (${a.linhas})`);
for (const aviso of dados.avisos) console.warn("  aviso: " + aviso);

const total = dados.stores.reduce((n, s) => n + s.rows.length, 0);
console.log(`${dados.stores.length} loja(s), ${total} linhas, ${Object.keys(dados.produtos).length} produtos. Gravando…`);

const resp = await fetch(`${URL_BASE}/rest/v1/rpc/importar_dados`, {
  method: "POST",
  headers: { apikey: CHAVE, Authorization: `Bearer ${CHAVE}`, "Content-Type": "application/json" },
  body: JSON.stringify({
    p: dados,
    p_url: entrada,
    p_resumo: {
      abas: dados.abas,
      avisos: dados.avisos,
      lojas: dados.stores.map((s) => ({ id: s.id, nome: s.nome, linhas: s.rows.length })),
      tamanho_kb: Math.round(buf.length / 1024),
      origem: "importar-local",
    },
  }),
});
if (!resp.ok) {
  console.error("Falhou:", await resp.text());
  Deno.exit(1);
}
console.log("Pronto:", await resp.text());
