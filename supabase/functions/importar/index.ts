// Edge Function "importar" — ponte para baixar a planilha do Google Sheets.
//
// O navegador não pode buscar o arquivo direto (o Google não envia cabeçalhos CORS),
// e ler a planilha aqui estourava o limite de CPU da Edge Function (2s no plano
// gratuito; a leitura leva uns 4s). Então esta função só baixa o .xlsx e devolve os
// bytes; quem lê a planilha e grava no banco é o painel, no navegador.
//
// Exige JWT: o Supabase valida antes de chegar aqui, e ainda conferimos abaixo que
// quem chamou é um usuário logado de verdade — a chave anônima sozinha não basta.

const CORS = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers": "authorization, x-client-info, apikey, content-type",
  "Access-Control-Allow-Methods": "POST, OPTIONS",
};

function erro(mensagem: string, status = 400) {
  return new Response(JSON.stringify({ status: "erro", erro: mensagem }), {
    status,
    headers: { ...CORS, "Content-Type": "application/json" },
  });
}

/** O JWT já foi validado pela plataforma; aqui só lemos quem ele diz ser. */
function ehUsuarioLogado(req: Request): boolean {
  const token = (req.headers.get("authorization") ?? "").replace(/^Bearer\s+/i, "");
  const payload = token.split(".")[1];
  if (!payload) return false;
  try {
    const dados = JSON.parse(atob(payload.replace(/-/g, "+").replace(/_/g, "/")));
    return dados.role === "authenticated" && typeof dados.sub === "string";
  } catch {
    return false;
  }
}

/** Aceita a URL completa de edição, a de compartilhamento ou só o ID. */
function idDaPlanilha(entrada: string): string | null {
  const s = entrada.trim();
  const m = s.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]{20,})/);
  if (m) return m[1];
  if (/^[a-zA-Z0-9-_]{20,}$/.test(s)) return s;
  return null;
}

Deno.serve(async (req) => {
  if (req.method === "OPTIONS") return new Response("ok", { headers: CORS });
  if (req.method !== "POST") return erro("Use POST.", 405);
  if (!ehUsuarioLogado(req)) return erro("Entre na sua conta para atualizar os dados.", 401);

  let url = "";
  try {
    const corpo = await req.json();
    url = String(corpo?.url ?? "");
  } catch {
    return erro("Envie um JSON com o campo url.");
  }

  const id = idDaPlanilha(url);
  if (!id) return erro("Não reconheci esse link. Cole a URL completa da planilha do Google Sheets.");

  let r: Response;
  try {
    r = await fetch(`https://docs.google.com/spreadsheets/d/${id}/export?format=xlsx`, {
      redirect: "follow",
    });
  } catch (e) {
    return erro("Não consegui falar com o Google: " + (e instanceof Error ? e.message : String(e)), 502);
  }
  if (!r.ok) {
    return erro(`O Google respondeu ${r.status} ao baixar a planilha. Confira se o link está correto.`, 502);
  }

  const buf = new Uint8Array(await r.arrayBuffer());
  // Sem permissão o Google devolve a página de login em HTML, não o arquivo .xlsx.
  if (!(buf[0] === 0x50 && buf[1] === 0x4b)) {
    return erro(
      "A planilha não está acessível publicamente. Abra o Google Sheets, clique em " +
        "Compartilhar, e em Acesso geral escolha “Qualquer pessoa com o link” como Leitor.",
    );
  }

  return new Response(buf, {
    headers: {
      ...CORS,
      "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      "X-Planilha-Id": id,
      "Cache-Control": "no-store",
    },
  });
});
