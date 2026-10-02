#!/usr/bin/env python3
"""
Carga única (retroativa) de resumo_cpfs_geradores_mensal para os 38
meses do histórico (jul/2023 a ago/2026) — roda 1 vez só.

Por quê isso existe separado do backfill_historico.py: a tabela
resumo_cpfs_geradores_mensal (ver 008_usuarios_unicos_periodo.sql) foi
criada DEPOIS do backfill histórico original já ter rodado com sucesso
e já ter apagado as planilhas de voucher do Storage (passo normal de
limpeza). Esse script reprocessa só o que essa tabela nova precisa —
não toca carteira_mensal, vouchers_detalhados nem resumo_kpis_mensal,
que já estão corretos — então é seguro rodar mesmo com o site já em
produção.

Pré-requisito: reenviar ao Storage (bucket "uploads-planilhas") as
MESMAS planilhas de vouchers já usadas no backfill original:

    <ano>/vouchers/<ano>.xlsx           (1 por ano, 2023-2026)
    2024/vouchers/2024-05-especial.csv  (opcional -- ver nota abaixo)

Se "2024/vouchers/2024-05-especial.csv" não for reenviado, este script
usa o arquivo anual de 2024 também para maio/2024 -- o número de
usuários únicos daquele mês específico pode ficar levemente subestimado
(a fonte CSV é mais completa, ver Metodologia Padrão), mas isso afeta
só esse 1 mês dentro de uma contagem de um intervalo maior (ex: o ano
inteiro), nunca o resumo mensal oficial (resumo_kpis_mensal), que não é
tocado por este script.

Ao final, remove do Storage os arquivos que processou (mesmo
comportamento de limpeza do backfill_historico.py).

Variáveis de ambiente esperadas (iguais aos outros scripts de ETL):
    SUPABASE_URL, SUPABASE_SERVICE_ROLE_KEY, CPF_PEPPER
"""
import os
import sys

sys.path.insert(0, os.path.dirname(__file__))
from carga_mensal import baixar_do_storage, carrega_vouchers, grava_cpfs_geradores
from backfill_historico import PLANO_ANO, arquivo_existe_no_storage
from supabase import create_client

BUCKET = "uploads-planilhas"


def main():
    supabase_url = os.environ["SUPABASE_URL"]
    service_role_key = os.environ["SUPABASE_SERVICE_ROLE_KEY"]
    pepper = os.environ["CPF_PEPPER"]
    supabase = create_client(supabase_url, service_role_key)

    admin_rows = supabase.table("contas_administrativas").select("cpf_hash").execute().data
    cpfs_hash_admin = {r["cpf_hash"] for r in admin_rows}
    print(f"Contas administrativas carregadas do banco: {len(cpfs_hash_admin)}")

    total_meses = sum(len(v) for v in PLANO_ANO.values())
    processados = 0

    for ano, meses in PLANO_ANO.items():
        vouchers_path_ano = f"{ano}/vouchers/{ano}.xlsx"

        if not arquivo_existe_no_storage(supabase, BUCKET, vouchers_path_ano):
            print(f"\n== Ano {ano}: {vouchers_path_ano} não encontrado no Storage -- "
                  f"reenvie esse arquivo antes de rodar este script. Pulando o ano. ==")
            processados += len(meses)
            continue

        vouchers_local_ano = baixar_do_storage(supabase, BUCKET, vouchers_path_ano)

        tem_csv_especial_2024_05 = (
            ano == 2024 and arquivo_existe_no_storage(supabase, BUCKET, "2024/vouchers/2024-05-especial.csv")
        )
        if ano == 2024 and not tem_csv_especial_2024_05:
            print("  (aviso: 2024/vouchers/2024-05-especial.csv não encontrado -- "
                  "maio/2024 vai usar o arquivo anual, pode ficar levemente "
                  "subestimado nesta tabela específica, ver docstring)")

        for mes in meses:
            mes_referencia = f"{ano:04d}-{mes:02d}-01"
            print(f"\n== {mes:02d}/{ano} (mes_referencia={mes_referencia}) "
                  f"[{processados + 1}/{total_meses}] ==")

            if (ano, mes) == (2024, 5) and tem_csv_especial_2024_05:
                voucher_path_mes = "2024/vouchers/2024-05-especial.csv"
                voucher_local_mes = baixar_do_storage(supabase, BUCKET, voucher_path_mes)
                registros_vouchers = carrega_vouchers(
                    voucher_local_mes, ano, mes, pepper, mes_referencia,
                    voucher_path_mes, cpfs_hash_admin,
                )
                os.unlink(voucher_local_mes)
            else:
                registros_vouchers = carrega_vouchers(
                    vouchers_local_ano, ano, mes, pepper, mes_referencia,
                    vouchers_path_ano, cpfs_hash_admin,
                )

            liquidos = [v for v in registros_vouchers if not v["excluido_farmacia"] and not v["excluido_anomalia"]]
            geradores_cpf = {v["cpf_hash"] for v in liquidos}
            grava_cpfs_geradores(supabase, mes_referencia, geradores_cpf)
            print(f"  usuários únicos geradores no mês: {len(geradores_cpf)}")

            processados += 1
            print(f"  OK -- {mes:02d}/{ano} concluído ({processados}/{total_meses})")

        os.unlink(vouchers_local_ano)
        supabase.storage.from_(BUCKET).remove([vouchers_path_ano])
        if tem_csv_especial_2024_05:
            supabase.storage.from_(BUCKET).remove(["2024/vouchers/2024-05-especial.csv"])
        print(f"Arquivo(s) de {ano} removidos do Storage.")

    print(f"\n== Backfill de usuários únicos (histórico) concluído: "
          f"{processados}/{total_meses} meses ==")


if __name__ == "__main__":
    main()
