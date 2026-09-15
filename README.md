# Precificador — Luz Sollar

Painel web para ajustar a precificação das lojas de marketplace a partir da planilha
do Google Sheets. Site estático, um arquivo só, sem build.

Funciona de dois jeitos:

- **Sozinho** — abra o `index.html` e pronto. Usa os dados embutidos na página e
  guarda seus ajustes no navegador.
- **Ligado ao Supabase** — com o `config.js` no lugar, o painel passa a ler os dados do
  banco, salvar seus ajustes lá (valem em qualquer computador) e ganhar uma aba
  **Dados** onde você cola o link novo da planilha e clica em *Atualizar dados*.

---

## O que o painel faz

- **8 lojas** num seletor: Luz Sollar, E-Steck, E-Brilia, E-Lumin, Apollo,
  Central Lux, Shopee e TikTok — o que existir na planilha.
- **Variáveis da loja** editáveis: NF/Imposto, ADS, Lucro alvo, Devolução e comissões.
- **Ajuste por produto**: custo, desconto, frete/taxa fixa, promoção ativa e rebate;
  ou travar manualmente preço mínimo, preço de venda e preço original.
- **Catálogo de custos**: mudar o custo de um SKU propaga para os kits e para todas as lojas.
- **Pesquisa de mercado**: média e os 5 preços concorrentes, com comparação de competitividade.
- **Filtros**: com promoção, margem negativa, preço acima do mercado, editados, sem custo.
- **Exportar CSV** separado por tabulação, com a coluna `Linha` apontando para a linha
  original da aba — pronto para colar de volta no Google Sheets.

### Fórmulas

Com `custoLíq = custo × (1 − desconto)` e `fixos` = frete + taxas fixas em R$:

```
taxas         = NF + ADS + Devolução + Comissão do canal

preçoMínimo   = (custoLíq + fixos) / (1 − taxas)
preçoVenda    = (custoLíq + fixos) / (1 − taxas − lucroAlvo)
preçoOriginal = preçoVenda / 0,8

lucro         = promoção − promoção × taxas − fixos − custoLíq
lucroFinal    = lucro + rebate
margem        = lucroFinal / promoção
```

O importador não assume um layout fixo: ele lê as **próprias fórmulas** da aba para
descobrir quais colunas são taxa (%), quais são custo fixo (R$) e qual é o lucro alvo.
Se você mudar a ordem das colunas ou criar uma loja nova, ele continua entendendo.

---

## Ligando ao Supabase

Faça uma vez só. Leva uns 15 minutos.

### 1. Projeto — já está pronto

Projeto `zvmpfsslhyjqwasvxlej`, na organização **Luz Sollar**, região us-east-1.

### 2. Tabelas — já estão criadas

As migrações `precificador_tabelas_e_rls` e `precificador_funcoes` já rodaram.
O arquivo [`supabase/schema.sql`](supabase/schema.sql) fica aqui como referência
e para recriar o banco do zero, se um dia precisar.

### 3. Criar o seu usuário

**Authentication → Users → Add user**: e-mail e senha, marcando *Auto Confirm User*.

Depois, em **Authentication → Sign In / Providers → Email**, desligue
**Allow new users to sign up**. Assim só quem você cadastrar entra.

### 4. Apontar o site para o projeto — já está feito

O `config.js` já vem preenchido com a *Project URL* e a chave *anon* deste projeto.

A chave `anon` pode ficar no repositório. Com ela se **lê** os dados do painel —
isso é intencional: quem abre o site sem entrar vê os preços da última importação.
Escrever, importar e ver o link da planilha exigem login.
**A chave `service_role` nunca entra aqui.**

### 5. Função de importação — já está publicada

A Edge Function `importar` está no ar. Ela **não lê** a planilha: só baixa o arquivo
do Google e devolve os bytes, porque o Google não envia cabeçalhos CORS e o navegador
sozinho não consegue buscar. Quem lê a planilha é o próprio painel, no seu navegador
— ler ali dentro da função estourava o limite de 2 segundos de CPU do plano gratuito.
Além do JWT que a plataforma valida, a função confere que quem chamou é um usuário
logado: a chave anônima sozinha não abre.

Para republicar depois de mexer no código:

```bash
npm install -g supabase
supabase login
supabase link --project-ref zvmpfsslhyjqwasvxlej
supabase functions deploy importar
```

Ela recebe `SUPABASE_URL` e `SUPABASE_SERVICE_ROLE_KEY` automaticamente — não há
segredo para configurar.

### 6. Liberar a planilha

No Google Sheets: **Compartilhar → Acesso geral → Qualquer pessoa com o link →
Leitor**. Sem isso o Google devolve a página de login e a importação falha com um
aviso explicando exatamente isso.

### 7. Importar

Abra o site, entre com o usuário do passo 3, vá na aba **Dados**, cole o link da
planilha e clique em **Atualizar dados**. Leva de 10 a 40 segundos.

A cada importação:

- os dados das lojas e do catálogo são substituídos por inteiro, numa transação só
  (ou entra tudo, ou nada muda);
- **seus ajustes de precificação não são apagados** — eles vivem em outra tabela;
- a tela mostra quais abas foram lidas como loja, quais como catálogo, quais foram
  ignoradas, e avisa sobre células com número malformado (ex.: `1,100,00`).

---

## Importar pela linha de comando

Se preferir não usar o painel, o mesmo importador roda no seu computador:

```bash
export SUPABASE_URL=https://seu-projeto.supabase.co
export SUPABASE_SERVICE_ROLE_KEY=...          # Project Settings → API
deno run -A supabase/importar-local.ts "https://docs.google.com/spreadsheets/d/…"
```

---

## Estrutura

```
index.html                            o painel inteiro, com os dados embutidos
config.example.js                     modelo do config.js (Supabase)
dados/dataset.json                    dados embutidos, usados sem Supabase
supabase/schema.sql                   tabelas, RLS e funções do banco
supabase/functions/importar/index.ts  Edge Function: baixa o .xlsx do Google
supabase/functions/importar/parser.ts leitor da planilha (usado no navegador e no CLI)
supabase/importar-local.ts            mesmo importador, para rodar no seu PC
vendor/xlsx.mjs                       SheetJS, carregado só na hora de importar
.nojekyll                             evita o processamento Jekyll no GitHub Pages
```

## Publicar no GitHub Pages

**Settings → Pages** → branch `main`, pasta `/ (root)`.

Se o repositório estiver conectado ao Supabase, lembre que o `config.js` fica visível
para quem abrir a página — o que é esperado, já que a proteção real são o login e as
políticas de RLS.
