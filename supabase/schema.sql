-- ============================================================================
-- Precificador — banco de dados
-- Rode este arquivo inteiro no SQL Editor do Supabase (uma vez só).
-- Rodar de novo é seguro: tudo é "if not exists" / "create or replace".
-- ============================================================================

create extension if not exists pgcrypto;

-- ---------------------------------------------------------------- tabelas ---

create table if not exists public.lojas (
  id            text primary key,           -- nome da aba na planilha
  nome          text not null,              -- nome curto exibido no painel
  plataforma    text not null default 'outro',
  ordem         int  not null default 0,
  padroes       jsonb not null default '{}'::jsonb,  -- taxas globais da loja
  fixos         jsonb not null default '[]'::jsonb,  -- custos fixos em R$
  canais        jsonb not null default '[]'::jsonb,  -- clássico / premium / padrão
  atualizado_em timestamptz not null default now()
);

create table if not exists public.produtos (
  sku           text primary key,
  descricao     text not null default '',
  custo         numeric,
  atualizado_em timestamptz not null default now()
);

create table if not exists public.linhas (
  loja_id        text not null references public.lojas(id) on delete cascade,
  idx            int  not null,             -- posição da linha dentro da aba
  linha_planilha int,                       -- número da linha no Google Sheets
  sku            text not null default '',
  nome           text not null default '',
  custo_regra    jsonb,                     -- ["L"] | ["M",idx,fator] | ["S",a,b] | ["F",valor]
  custo_cache    numeric,                   -- custo que estava na planilha
  desconto       numeric not null default 0,
  custo_liquido  numeric,                   -- quando digitado à mão na planilha
  fixos          jsonb not null default '{}'::jsonb,
  taxas          jsonb not null default '{}'::jsonb,  -- exceções de taxa da linha
  promo_c        numeric,
  rebate_c       numeric,
  promo_p        numeric,
  rebate_p       numeric,
  travas         jsonb not null default '{}'::jsonb,  -- preços digitados à mão
  mercado        jsonb not null default '[]'::jsonb,  -- [média, p1..p5]
  observacao     text not null default '',
  primary key (loja_id, idx)
);

create index if not exists linhas_sku_idx  on public.linhas (sku);
create index if not exists linhas_nome_idx on public.linhas (nome);

-- ajustes de precificação feitos no painel (um documento por perfil)
create table if not exists public.estado (
  id            text primary key default 'padrao',
  dados         jsonb not null default '{}'::jsonb,
  atualizado_em timestamptz not null default now(),
  atualizado_por uuid
);

-- histórico de importações
create table if not exists public.importacoes (
  id            bigint generated always as identity primary key,
  planilha_url  text,
  planilha_id   text,
  status        text not null default 'ok',
  resumo        jsonb not null default '{}'::jsonb,
  criado_em     timestamptz not null default now(),
  criado_por    uuid
);

-- configurações simples (link atual da planilha etc.)
create table if not exists public.config (
  chave         text primary key,
  valor         jsonb not null default '{}'::jsonb,
  atualizado_em timestamptz not null default now()
);

-- ------------------------------------------------------------------- rls ---
-- Tudo exige usuário logado. A chave anônima sozinha não lê nada.

alter table public.lojas       enable row level security;
alter table public.produtos    enable row level security;
alter table public.linhas      enable row level security;
alter table public.estado      enable row level security;
alter table public.importacoes enable row level security;
alter table public.config      enable row level security;

do $$
declare t text;
begin
  foreach t in array array['lojas','produtos','linhas','importacoes','config'] loop
    execute format('drop policy if exists %I on public.%I', t || '_leitura', t);
    execute format(
      'create policy %I on public.%I for select to authenticated using (true)',
      t || '_leitura', t);
  end loop;

  -- o estado (ajustes do painel) é o único que o usuário logado escreve direto
  drop policy if exists estado_leitura on public.estado;
  drop policy if exists estado_escrita on public.estado;
  create policy estado_leitura on public.estado
    for select to authenticated using (true);
  create policy estado_escrita on public.estado
    for all to authenticated using (true) with check (true);

  -- LEITURA PÚBLICA: quem abre o site sem entrar enxerga os dados do painel.
  -- Só leitura, e só destas quatro tabelas — "config" e "importacoes" ficam de
  -- fora de propósito, porque guardam o link do Google Sheets.
  foreach t in array array['lojas','produtos','linhas','estado'] loop
    execute format('drop policy if exists %I on public.%I', t || '_leitura_anon', t);
    execute format(
      'create policy %I on public.%I for select to anon using (true)',
      t || '_leitura_anon', t);
  end loop;
end $$;

grant select on public.lojas, public.produtos, public.linhas, public.estado to anon;

-- --------------------------------------------------------------- funções ---

-- A leitura da planilha acontece no NAVEGADOR (a Edge Function estourava o limite
-- de CPU do plano gratuito). Esta função recebe o resultado já lido e grava tudo
-- de uma vez: dados, link da planilha e registro da importação, numa transação só.
create or replace function public.importar_dados(p jsonb, p_url text, p_resumo jsonb)
returns jsonb
language plpgsql
security definer
set search_path = public
as $$
declare
  n_lojas int := 0;
  n_linhas int := 0;
  n_produtos int := 0;
  v_id text;
begin
  if jsonb_array_length(coalesce(p->'stores', '[]'::jsonb)) = 0 then
    raise exception 'Nenhuma aba de loja foi reconhecida na planilha.';
  end if;

  -- TRUNCATE em vez de DELETE: o papel "authenticator" (PostgREST) carrega a
  -- biblioteca "safeupdate", que rejeita DELETE sem WHERE em qualquer consulta
  -- da sessão — inclusive dentro de funções SECURITY DEFINER.
  truncate table public.linhas, public.lojas, public.produtos;

  insert into public.lojas (id, nome, plataforma, ordem, padroes, fixos, canais)
  select
    loja->>'id',
    coalesce(loja->>'nome', loja->>'id'),
    coalesce(loja->>'plataforma', 'outro'),
    (ord - 1)::int,
    coalesce(loja->'defaults', '{}'::jsonb),
    coalesce(loja->'fixos', '[]'::jsonb),
    coalesce(loja->'canais', '[]'::jsonb)
  from jsonb_array_elements(p->'stores') with ordinality as t(loja, ord);
  get diagnostics n_lojas = row_count;

  insert into public.produtos (sku, descricao, custo)
  select chave, coalesce(valor->>0, ''), nullif(valor->>1, '')::numeric
  from jsonb_each(coalesce(p->'produtos', '{}'::jsonb)) as t(chave, valor);
  get diagnostics n_produtos = row_count;

  insert into public.linhas (
    loja_id, idx, linha_planilha, sku, nome, custo_regra, custo_cache, desconto,
    custo_liquido, fixos, taxas, promo_c, rebate_c, promo_p, rebate_p,
    travas, mercado, observacao)
  select
    loja->>'id',
    (pos - 1)::int,
    nullif(linha->>'r', '')::int,
    coalesce(linha->>'sku', ''),
    coalesce(linha->>'nome', ''),
    linha->'cr',
    nullif(linha->>'cv', '')::numeric,
    coalesce(nullif(linha->>'desc', '')::numeric, 0),
    nullif(linha->>'eov', '')::numeric,
    coalesce(linha->'fixos', '{}'::jsonb),
    coalesce(linha->'rates', '{}'::jsonb),
    nullif(linha->>'promoC', '')::numeric,
    nullif(linha->>'rebC', '')::numeric,
    nullif(linha->>'promoP', '')::numeric,
    nullif(linha->>'rebP', '')::numeric,
    coalesce(linha->'ov', '{}'::jsonb),
    coalesce(linha->'mkt', '[]'::jsonb),
    coalesce(linha->>'obs', '')
  from jsonb_array_elements(p->'stores') as loja,
       jsonb_array_elements(loja->'rows') with ordinality as t(linha, pos);
  get diagnostics n_linhas = row_count;

  v_id := substring(coalesce(p_url, '') from '/spreadsheets/d/([a-zA-Z0-9_-]+)');

  insert into public.config (chave, valor, atualizado_em)
  values ('planilha',
          jsonb_build_object('url', p_url, 'id', v_id, 'importado_em', now()),
          now())
  on conflict (chave) do update
    set valor = excluded.valor, atualizado_em = now();

  insert into public.importacoes (planilha_url, planilha_id, status, resumo, criado_por)
  values (p_url, v_id, 'ok',
          coalesce(p_resumo, '{}'::jsonb)
            || jsonb_build_object('lojas', n_lojas, 'produtos', n_produtos, 'linhas', n_linhas),
          auth.uid());

  return jsonb_build_object('lojas', n_lojas, 'produtos', n_produtos, 'linhas', n_linhas);
end $$;

revoke all on function public.importar_dados(jsonb, text, jsonb) from public, anon;
grant execute on function public.importar_dados(jsonb, text, jsonb) to authenticated;

-- registra tentativas que falharam
create or replace function public.registrar_falha(p_url text, p_erro text)
returns void
language sql
security definer
set search_path = public
as $$
  insert into public.importacoes (planilha_url, status, resumo, criado_por)
  values (p_url, 'erro', jsonb_build_object('erro', p_erro), auth.uid());
$$;

revoke all on function public.registrar_falha(text, text) from public, anon;
grant execute on function public.registrar_falha(text, text) to authenticated;

-- Salva os ajustes do painel sem risco de sobrescrever o trabalho de outra
-- sessão: se o registro mudou depois do "visto_em" informado, devolve conflito.
create or replace function public.salvar_estado(p_id text, p_dados jsonb, p_visto_em timestamptz)
returns jsonb
language plpgsql
security invoker
set search_path = public
as $$
declare
  atual timestamptz;
begin
  select atualizado_em into atual from public.estado where id = p_id;

  if atual is not null and p_visto_em is not null and atual > p_visto_em then
    return jsonb_build_object('status', 'conflito', 'atualizado_em', atual);
  end if;

  insert into public.estado (id, dados, atualizado_em, atualizado_por)
  values (p_id, p_dados, now(), auth.uid())
  on conflict (id) do update
    set dados = excluded.dados,
        atualizado_em = now(),
        atualizado_por = auth.uid();

  select atualizado_em into atual from public.estado where id = p_id;
  return jsonb_build_object('status', 'ok', 'atualizado_em', atual);
end $$;

grant execute on function public.salvar_estado(text, jsonb, timestamptz) to authenticated;

insert into public.estado (id) values ('padrao') on conflict (id) do nothing;
