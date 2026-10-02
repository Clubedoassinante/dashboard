#!/usr/bin/env python3
"""
Carga única (retroativa) das 3 métricas oficiais novas — carteira_clube_
oficial, penetracao_oficial_pct, media_uso_dia_oficial (ver Blueprint,
decisão de 02/10/2026 e supabase/009_metricas_oficiais_extras.sql) —
para os meses do histórico que já foram carregados ANTES dessas 3
colunas existirem.

Diferente de backfill_usuarios_unicos_hist.py, este script NÃO precisa
reprocessar nenhuma planilha de voucher/carteira: as 3 métricas novas
são números "oficiais" prontos (vêm da planilha de referência da
Alloyal, já copiados em etl/oficial_alloyal.json), não recalculados a
partir do detalhe cru. Por isso este script só faz um UPDATE simples
em resumo_kpis_mensal + vouchers_oficial_mensal para cada mês que já
existe em oficial_alloyal.json — não acessa o Supabase Storage, não
precisa de nenhuma planilha reenviada.

Seguro rodar mesmo com o site já em produção: só toca as 3 colunas
novas (e só nas linhas de mês que já existem), nunca reescreve carteira_
ativa, vouchers_liquidos, penetracao_pct nem qualquer outro campo já
calculado.

Variáveis de ambiente esperadas (iguais aos outros scripts de ETL):
    SUPABASE_URL, SUPABASE_SERVICE_ROLE_KEY
(CPF_PEPPER não é necessário aqui -- este script não lida com CPF.)
"""
import json
import os

from supabase import create_client


def main():
    supabase_url = os.environ["SUPABASE_URL"]
    service_role_key = os.environ["SUPABASE_SERVICE_ROLE_KEY"]
    supabase = create_client(supabase_url, service_role_key)

    oficial_path = os.path.join(os.path.dirname(__file__), "oficial_alloyal.json")
    with open(oficial_path, encoding="utf-8") as f:
        oficial = json.load(f)["referencia_mensal"]

    meses_existentes = {
        r["mes_referencia"]
        for r in supabase.table("resumo_kpis_mensal").select("mes_referencia").execute().data
    }
    print(f"Meses já carregados em resumo_kpis_mensal: {len(meses_existentes)}")

    atualizados = 0
    pulados_sem_mes = 0
    pulados_sem_campo_novo = 0

    for mes_curto, of in sorted(oficial.items()):
        mes_referencia = f"{mes_curto}-01"
        campos_novos = {
            "carteira_clube_oficial": of.get("carteira_clube_oficial"),
            "penetracao_oficial_pct": of.get("penetracao_oficial_pct"),
            "media_uso_dia_oficial": of.get("media_uso_dia_oficial"),
        }
        if all(v is None for v in campos_novos.values()):
            pulados_sem_campo_novo += 1
            continue
        if mes_referencia not in meses_existentes:
            # esse mês ainda não foi carregado (ex: ainda não rodou a carga
            # mensal normal) -- nada a atualizar aqui, vai entrar certo
            # quando a carga mensal normal rodar (carga_mensal.py já passa
            # os 3 campos novos desde a atualização de 02/10/2026)
            pulados_sem_mes += 1
            continue

        supabase.table("resumo_kpis_mensal").update(campos_novos).eq(
            "mes_referencia", mes_referencia
        ).execute()

        # vouchers_oficial_mensal é uma tabela de auditoria/referência à
        # parte (não é o que o front-end lê) -- só atualiza se a linha já
        # existir, pra não inventar um arquivo_origem genérico numa linha
        # que nunca foi gravada por uma carga de verdade.
        existe_oficial = supabase.table("vouchers_oficial_mensal").select("mes_referencia").eq(
            "mes_referencia", mes_referencia
        ).execute().data
        if existe_oficial:
            supabase.table("vouchers_oficial_mensal").update(campos_novos).eq(
                "mes_referencia", mes_referencia
            ).execute()

        atualizados += 1
        print(f"  {mes_curto}: atualizado ({campos_novos})")

    print(f"\n== Backfill de métricas oficiais concluído: {atualizados} meses atualizados "
          f"({pulados_sem_mes} ainda não carregados, {pulados_sem_campo_novo} sem campo novo no json) ==")


if __name__ == "__main__":
    main()
