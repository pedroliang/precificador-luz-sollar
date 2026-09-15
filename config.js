// Projeto Supabase do Precificador.
//
// A chave "anon" é pública por natureza. Com ela dá para LER os dados do painel
// (lojas, produtos, linhas e os ajustes) — a leitura é aberta de propósito, para
// que quem não tem login veja os números. Escrever, importar a planilha e ver o
// link do Google Sheets continuam exigindo usuário logado.
// NUNCA coloque aqui a service_role.

window.PRECIFICADOR_CONFIG = {
  url: "https://zvmpfsslhyjqwasvxlej.supabase.co",
  anonKey: "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Inp2bXBmc3NsaHlqcXdhc3Z4bGVqIiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODc2NzI0MjQsImV4cCI6MjEwMzI0ODQyNH0.Z9c4mFrWtsGn4n9lPrWInNZ-vornY63q25nFtQ0bQdI",
};
