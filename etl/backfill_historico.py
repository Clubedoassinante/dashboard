#!/usr/bin/env python3
"""
Carga única do histórico (jul/2023 a ago/2026) — roda 1 vez só.

Espera que os arquivos já estejam no bucket "uploads-planilhas" do
Supabase Storage, na estrutura:

    <ano>/carteira/<ano>-<mes>.xlsx     (1 por mês, 38 arquivos)
    <ano>/vouchers/<ano>.xlsx           (1 por ano, arquivo consolidado)
    2024/vouchers/2024-05-especial.csv  (só maio/2024, fonte alternativa)

Processa os 38 meses em sequência, reaproveitando exatamente as mesmas
funções do carga_mensal.py (mesmas regras, mesmo hash de CPF). Ao final
de cada mês bem-sucedido, os arquivos daquele mês são removidos do
Storage (os arquivos anuais de voucher só são removidos depois do
último mês daquele ano que os usa).

Variáveis de ambiente esperadas (iguais ao carga_mensal.py):
    SUPABASE_URL, SUPABASE_SERVICE_ROLE_KEY, CPF_PEPPER
"""
import calendar
import json
import os
import sys

sys.path.insert(0, os.path.dirname(__file__))
from carga_mensal import (
    baixar_do_storage, carrega_carteira, carrega_vouchers, grava_em_lotes,
    apaga_mes_em_lotes, calcula_resumo, grava_resumo,
    JANELA_MESES_DETALHE, soma_meses, mantem_janela_detalhe,
)
from supabase import create_client
import datetime

BUCKET = "uploads-planilhas"

MESES_2023 = list(range(7, 13))
MESES_PADRAO = list(range(1, 13))
MESES_2026 = list(range(1, 9))

PLANO_ANO = {
    2023: MESES_2023,
    2024: MESES_PADRAO,
    2025: MESES_PADRAO,
    2026: MESES_2026,
}

# Último mês coberto por este backfill -- usado pra decidir, por mês,
# se ele cai dentro da janela móvel de detalhe cru (ver
# JANELA_MESES_DETALHE em carga_mensal.py e Blueprint, bug #8).
_ultimo_ano = max(PLANO_ANO)
ULTIMO_MES_BACKFILL = f"{_ultimo_ano:04d}-{max(PLANO_ANO[_ultimo_ano]):02d}-01"
CUTOFF_DETALHE = soma_meses(ULTIMO_MES_BACKFILL, -(JANELA_MESES_DETALHE - 1))


def arquivo_existe_no_storage(supabase, bucket, path):
    """Confere se um arquivo existe no Storage (sem baixar o conteúdo)."""
    pasta = os.path.dirname(path)
    nome = os.path.basename(path)
    arquivos = supabase.storage.from_(bucket).list(pasta)
    return any(a["name"] == nome for a in arquivos)


def main():
    supabase_url = os.environ["SUPABASE_URL"]
    service_role_key = os.environ["SUPABASE_SERVICE_ROLE_KEY"]
    pepper = os.environ["CPF_PEPPER"]
    supabase = create_client(supabase_url, service_role_key)

    oficial_path = os.path.join(os.path.dirname(__file__), "oficial_alloyal.json")
    with open(oficial_path, encoding="utf-8") as f:
        oficial = json.load(f)["referencia_mensal"]

    admin_rows = supabase.table("contas_administrativas").select("cpf_hash").execute().data
    cpfs_hash_admin = {r["cpf_hash"] for r in admin_rows}
    print(f"Contas administrativas carregadas do banco: {len(cpfs_hash_admin)}")

    total_meses = sum(len(v) for v in PLANO_ANO.values())
    processados = 0

    for ano, meses in PLANO_ANO.items():
        vouchers_path_ano = f"{ano}/vouchers/{ano}.xlsx"

        # Se esse script já rodou antes e concluiu esse ano inteiro com
        # sucesso, o arquivo anual já foi removido do Storage ao final
        # (ver limpeza no fim do loop de meses, abaixo). Reconhece esse
        # caso e pula o ano inteiro, em vez de quebrar com "Object not
        # found" -- permite re-rodar o backfill do zero com segurança
        # depois de uma falha em um ano posterior, sem reprocessar (nem
        # precisar re-subir arquivos de) anos já concluídos.
        if not arquivo_existe_no_storage(supabase, BUCKET, vouchers_path_ano):
            print(f"\n== Ano {ano}: {vouchers_path_ano} não existe mais no Storage -- "
                  f"esse ano já foi concluído e limpo em uma rodada anterior. Pulando. ==")
            processados += len(meses)
            continue

        vouchers_local_ano = baixar_do_storage(supabase, BUCKET, vouchers_path_ano)

        for mes in meses:
            mes_referencia = f"{ano:04d}-{mes:02d}-01"
            ultimo_dia = calendar.monthrange(ano, mes)[1]
            ref_date = datetime.datetime(ano, mes, ultimo_dia)
            carteira_path = f"{ano}/carteira/{ano}-{mes:02d}.xlsx"

            print(f"\n== {mes:02d}/{ano} (mes_referencia={mes_referencia}) "
                  f"[{processados + 1}/{total_meses}] ==")

            carteira_local = baixar_do_storage(supabase, BUCKET, carteira_path)
            registros_carteira = carrega_carteira(
                carteira_local, ref_date, pepper, mes_referencia, carteira_path
            )
            os.unlink(carteira_local)

            # maio/2024 usa uma fonte alternativa em CSV (mais completa --
            # ver Metodologia Padrão, decisão de 16/09/2026), em vez do
            # arquivo anual consolidado de 2024.
            if (ano, mes) == (2024, 5):
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

            of = oficial.get(mes_referencia[:7])
            kpis, composicao_rows, quem_gerou_rows = calcula_resumo(
                registros_carteira, registros_vouchers, mes_referencia,
                of["vouchers_gerados_oficial"] if of else None,
                of["usuarios_unicos_oficial"] if of else None,
                of["frequencia_uso_oficial"] if of else None,
            )
            grava_resumo(supabase, mes_referencia, kpis, composicao_rows, quem_gerou_rows)

            if mes_referencia >= CUTOFF_DETALHE:
                apaga_mes_em_lotes(supabase, "carteira_mensal", mes_referencia)
                apaga_mes_em_lotes(supabase, "vouchers_detalhados", mes_referencia)
                grava_em_lotes(supabase, "carteira_mensal", registros_carteira)
                grava_em_lotes(supabase, "vouchers_detalhados", registros_vouchers)
            else:
                print(f"  (fora da janela de {JANELA_MESES_DETALHE} meses: só o resumo foi "
                      f"gravado, sem detalhe linha a linha -- ver Blueprint, bug #8)")

            if of:
                supabase.table("vouchers_oficial_mensal").upsert({
                    "mes_referencia": mes_referencia,
                    "vouchers_gerados_oficial": of["vouchers_gerados_oficial"],
                    "usuarios_unicos_oficial": of["usuarios_unicos_oficial"],
                    "frequencia_uso_oficial": of["frequencia_uso_oficial"],
                    "arquivo_origem": "oficial_alloyal.json (referência histórica)",
                }).execute()
                print("  vouchers_oficial_mensal: atualizado")

            processados += 1
            print(f"  OK -- {mes:02d}/{ano} concluído ({processados}/{total_meses})")

        os.unlink(vouchers_local_ano)
        supabase.storage.from_(BUCKET).remove([vouchers_path_ano])
        if ano == 2024:
            supabase.storage.from_(BUCKET).remove(["2024/vouchers/2024-05-especial.csv"])
        carteira_paths_ano = [f"{ano}/carteira/{ano}-{m:02d}.xlsx" for m in meses]
        supabase.storage.from_(BUCKET).remove(carteira_paths_ano)
        print(f"Arquivos de {ano} removidos do Storage.")

    # Varredura final: garante que nenhum mês fora da janela ficou com
    # detalhe cru no banco -- cobre tanto o caso normal quanto o caso de
    # já ter rodado antes sob uma janela/regra diferente (como aconteceu
    # em 01/10/2026: os 8 meses de 2026 tinham sido carregados com
    # detalhe cru sob a regra antiga de "ano calendário inteiro", e essa
    # chamada é o que efetivamente arquiva jan/fev-2026, que ficaram de
    # fora da janela de 6 meses).
    mantem_janela_detalhe(supabase, ULTIMO_MES_BACKFILL)

    print(f"\n== Backfill histórico concluído: {processados}/{total_meses} meses ==")


if __name__ == "__main__":
    main()
