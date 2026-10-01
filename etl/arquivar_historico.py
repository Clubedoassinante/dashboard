#!/usr/bin/env python3
"""
Arquivamento único dos meses já carregados anteriores a 2026 — roda 1 vez
(e de novo, sem problema, sempre que sobrar algo pra arquivar: é
idempotente, pula mês que já não tem detalhe cru).

Decisão de Stela (01/10/2026): detalhe linha a linha (cpf_hash por
assinatura/voucher) só é mantido a partir de 2026. Meses anteriores
(jul/2023-dez/2025) já carregados em rodadas anteriores do backfill
continuam com o detalhe cru ocupando espaço no banco -- este script:

  1) calcula o resumo agregado de cada mês (mesmos cortes do dashboard:
     produto, plano_familia, faixa_etaria, faixa_tenure, safra_semestral)
     a partir do detalhe cru que já está no banco;
  2) grava esse resumo em resumo_kpis_mensal / resumo_composicao_mensal /
     resumo_quem_gerou_mensal (ver 008_resumo_mensal.sql -- rodar esse
     SQL no Supabase ANTES deste script);
  3) apaga o detalhe cru (carteira_mensal/vouchers_detalhados) daquele
     mês, em lotes, liberando o espaço que ele ocupava.

Depois disso, o histórico completo (jul/2023 em diante) continua
disponível pro dashboard através das views -- só que como resumo, não
mais linha a linha, pros meses anteriores a 2026.

Variáveis de ambiente esperadas (iguais aos outros scripts do etl/):
    SUPABASE_URL, SUPABASE_SERVICE_ROLE_KEY
"""
import os
import sys
from collections import Counter

sys.path.insert(0, os.path.dirname(__file__))
from carga_mensal import apaga_mes_em_lotes
from backfill_historico import PLANO_ANO, ANO_INICIO_DETALHE
from supabase import create_client

PAGINA = 1000

DIMENSOES_COMPOSICAO = [
    ("produto", "sigla_produto"),
    ("plano_familia", "plano_familia"),
    ("faixa_etaria", "faixa_etaria"),
    ("faixa_tenure", "faixa_tenure"),
    ("safra_semestral", "safra_semestral"),
]


def busca_tudo(supabase, tabela, colunas, mes_referencia):
    """Busca todas as linhas de um mês, paginando (a API do Supabase
    limita quantas linhas voltam por chamada)."""
    linhas = []
    inicio = 0
    while True:
        pagina = (
            supabase.table(tabela)
            .select(colunas)
            .eq("mes_referencia", mes_referencia)
            .range(inicio, inicio + PAGINA - 1)
            .execute()
            .data
        )
        if not pagina:
            break
        linhas.extend(pagina)
        if len(pagina) < PAGINA:
            break
        inicio += PAGINA
    return linhas


def calcula_resumo_do_banco(supabase, mes_referencia):
    """Mesma lógica de carga_mensal.calcula_resumo(), mas lendo o
    detalhe cru do banco em vez de uma lista em memória -- usado aqui
    porque, pra um mês já carregado antes, os registros originais não
    estão mais disponíveis em memória (vieram de uma rodada passada)."""
    carteira = busca_tudo(
        supabase, "carteira_mensal",
        "cpf_hash,situacao,sigla_produto,plano_familia,faixa_etaria,faixa_tenure,safra_semestral,mes_entrada",
        mes_referencia,
    )
    vouchers = busca_tudo(
        supabase, "vouchers_detalhados",
        "cpf_hash,excluido_farmacia,excluido_anomalia",
        mes_referencia,
    )
    oficial_rows = (
        supabase.table("vouchers_oficial_mensal")
        .select("vouchers_gerados_oficial,usuarios_unicos_oficial,frequencia_uso_oficial")
        .eq("mes_referencia", mes_referencia)
        .execute()
        .data
    )
    oficial = oficial_rows[0] if oficial_rows else {}

    ativos = [r for r in carteira if r["situacao"] == "Ativa"]
    canceladas = sum(1 for r in carteira if r["situacao"] == "Cancelada")
    novos_assinantes = sum(1 for r in ativos if r["mes_entrada"] == mes_referencia)
    plano_familia_ativo = sum(1 for r in ativos if r["plano_familia"] == "Ativo")
    carteira_ativa = len(ativos)
    plano_familia_ativo_pct = (
        round(plano_familia_ativo / carteira_ativa * 100, 2) if carteira_ativa else None
    )

    liquidos = [v for v in vouchers if not v["excluido_farmacia"] and not v["excluido_anomalia"]]
    vouchers_liquidos = len(liquidos)
    geradores_cpf = {v["cpf_hash"] for v in liquidos}
    usuarios_unicos_geradores = len(geradores_cpf)
    penetracao_pct = (
        round(usuarios_unicos_geradores / carteira_ativa * 100, 2) if carteira_ativa else None
    )

    kpis = {
        "mes_referencia": mes_referencia,
        "carteira_ativa": carteira_ativa,
        "canceladas": canceladas,
        "novos_assinantes": novos_assinantes,
        "plano_familia_ativo_pct": plano_familia_ativo_pct,
        "vouchers_liquidos": vouchers_liquidos,
        "usuarios_unicos_geradores": usuarios_unicos_geradores,
        "penetracao_pct": penetracao_pct,
        "vouchers_gerados_oficial": oficial.get("vouchers_gerados_oficial"),
        "usuarios_unicos_oficial": oficial.get("usuarios_unicos_oficial"),
        "frequencia_uso_oficial": oficial.get("frequencia_uso_oficial"),
    }

    composicao_rows = []
    quem_gerou_rows = []
    for dimensao, campo in DIMENSOES_COMPOSICAO:
        contagem_total = Counter()
        contagem_geradores = Counter()
        for r in ativos:
            categoria = r[campo]
            contagem_total[categoria] += 1
            if r["cpf_hash"] in geradores_cpf:
                contagem_geradores[categoria] += 1
        for categoria, total in contagem_total.items():
            composicao_rows.append({
                "mes_referencia": mes_referencia, "dimensao": dimensao,
                "categoria": categoria, "qtd": total,
            })
            quem_gerou_rows.append({
                "mes_referencia": mes_referencia, "dimensao": dimensao,
                "categoria": categoria, "carteira_ativa": total,
                "geradores": contagem_geradores.get(categoria, 0),
            })

    return kpis, composicao_rows, quem_gerou_rows, len(carteira), len(vouchers)


def grava_resumo(supabase, mes_referencia, kpis, composicao_rows, quem_gerou_rows):
    supabase.table("resumo_kpis_mensal").upsert(kpis).execute()
    supabase.table("resumo_composicao_mensal").delete().eq("mes_referencia", mes_referencia).execute()
    if composicao_rows:
        supabase.table("resumo_composicao_mensal").insert(composicao_rows).execute()
    supabase.table("resumo_quem_gerou_mensal").delete().eq("mes_referencia", mes_referencia).execute()
    if quem_gerou_rows:
        supabase.table("resumo_quem_gerou_mensal").insert(quem_gerou_rows).execute()


def main():
    supabase = create_client(os.environ["SUPABASE_URL"], os.environ["SUPABASE_SERVICE_ROLE_KEY"])

    meses = [
        f"{ano:04d}-{mes:02d}-01"
        for ano, lista_meses in PLANO_ANO.items() if ano < ANO_INICIO_DETALHE
        for mes in lista_meses
    ]
    print(f"Verificando {len(meses)} meses anteriores a {ANO_INICIO_DETALHE} "
          f"(candidatos a arquivamento)...")

    arquivados = 0
    for i, mes_referencia in enumerate(meses, start=1):
        kpis, composicao_rows, quem_gerou_rows, n_carteira, n_vouchers = calcula_resumo_do_banco(
            supabase, mes_referencia
        )
        if n_carteira == 0:
            print(f"[{i}/{len(meses)}] {mes_referencia}: sem detalhe cru no banco "
                  f"(já arquivado antes, ou nunca chegou a ser carregado) -- pulando.")
            continue

        print(f"[{i}/{len(meses)}] {mes_referencia}: {n_carteira} linhas de carteira, "
              f"{n_vouchers} de vouchers -- calculando resumo e liberando espaço...")
        grava_resumo(supabase, mes_referencia, kpis, composicao_rows, quem_gerou_rows)
        apaga_mes_em_lotes(supabase, "carteira_mensal", mes_referencia)
        apaga_mes_em_lotes(supabase, "vouchers_detalhados", mes_referencia)
        arquivados += 1
        print(f"  OK -- resumo gravado (carteira_ativa={kpis['carteira_ativa']}, "
              f"vouchers_liquidos={kpis['vouchers_liquidos']}) e detalhe cru liberado")

    print(f"\n== Arquivamento concluído: {arquivados} mês(es) tiveram o detalhe cru liberado "
          f"({len(meses) - arquivados} já estavam em dia) ==")


if __name__ == "__main__":
    main()
