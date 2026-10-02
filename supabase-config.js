// Configuração do cliente Supabase -- compartilhada por login.html e
// painel-mensal.html (e pelos próximos painéis: consolidado, sazonalidade).
//
// A chave abaixo é a PUBLISHABLE key (prefixo sb_publishable_), feita pra
// rodar no navegador -- ela só enxerga o que o Row Level Security do banco
// deixar (ver supabase/003_*.sql / 004_*.sql). Nunca coloque aqui a chave
// secreta (sb_secret_...) -- essa é só do ETL, roda no GitHub Actions.
const SUPABASE_URL = 'https://dbxncpmzdvkqrebdvqpe.supabase.co';
const SUPABASE_PUBLISHABLE_KEY = 'sb_publishable_9GZkZlBbaZe5l05g0wyo3A_wudUpws0';

const supabaseClient = window.supabase.createClient(SUPABASE_URL, SUPABASE_PUBLISHABLE_KEY);

// Garante que existe uma sessão logada; se não existir, manda pro login e
// interrompe a execução do resto da página (quem chamar deve checar o
// retorno antes de continuar a buscar dados).
async function exigirSessao() {
  const { data: { session } } = await supabaseClient.auth.getSession();
  if (!session) {
    window.location.href = 'login.html';
    return null;
  }
  return session;
}

async function sair() {
  await supabaseClient.auth.signOut();
  window.location.href = 'login.html';
}
