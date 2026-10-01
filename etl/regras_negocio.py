"""
Regras de negócio do Clube do Assinante GZH — portadas de
process_month.py (pipeline original, JSON) sem alterar nenhuma lógica.
Qualquer mudança de regra deve ser feita aqui E documentada na
Metodologia Padrão do projeto, exatamente como já era feito antes.
"""
import datetime
from collections import Counter

# ---------------------------------------------------------------------
# Farmácias excluídas dos vouchers (Panvel NÃO é excluída, de propósito)
# ---------------------------------------------------------------------
FARMACIAS_EXCLUIDAS = {
    "droga raia", "farmacia pague menos", "farmácia pague menos",
    "farmacias pague menos", "farmácias pague menos",
    "drogaria araujo", "drogaria araújo",
    "drogarias pacheco", "drogaria pacheco",
    "drogaria sao paulo", "drogaria são paulo",
    "farmacias droga raia e drogasil", "farmácias droga raia e drogasil",
}


def is_farmacia_excluida(marca):
    if marca is None:
        return False
    return str(marca).strip().lower() in FARMACIAS_EXCLUIDAS


# ---------------------------------------------------------------------
# Faixa etária (com sentinela dinâmica de "idade não informada")
# ---------------------------------------------------------------------
SENTINELA_IDADE_FIXA = {0.0}


def detecta_sentinela_dinamica(rows, col_idx=7, limiar_min_count=50):
    contagem = Counter()
    for row in rows:
        idade = row[col_idx] if len(row) > col_idx else None
        try:
            idade = float(idade)
        except (TypeError, ValueError):
            continue
        if idade > 115:
            contagem[idade] += 1
    if not contagem:
        return None
    valor, cnt = contagem.most_common(1)[0]
    return valor if cnt >= limiar_min_count else None


def faixa_etaria(idade, sentinela_dinamica=None):
    sentinelas = SENTINELA_IDADE_FIXA | ({sentinela_dinamica} if sentinela_dinamica is not None else set())
    if idade is None:
        return "Idade não informada"
    try:
        idade = float(idade)
    except (TypeError, ValueError):
        return "Idade não informada"
    if idade in sentinelas:
        return "Idade não informada"
    if idade < 18:
        return "Erro na idade"
    if idade <= 29:
        return "Gen Z"
    if idade <= 45:
        return "Millennials"
    if idade <= 61:
        return "Gen X"
    if idade <= 80:
        return "Baby Boomers"
    if idade <= 85:
        return "Pré-Boomer"
    if idade <= 115:
        return "86+ (fora da tabela padrão)"
    return "Erro na idade"


def faixa_tenure(data_assinatura, ref_date):
    if data_assinatura is None or not isinstance(data_assinatura, datetime.datetime):
        return "Erro na data de assinatura"
    if data_assinatura.year < 1950 or data_assinatura > ref_date:
        return "Erro na data de assinatura"
    dias = (ref_date - data_assinatura).days
    meses = dias / 30.4375
    anos = dias / 365.25
    if meses <= 6:
        return "Onboarding (0–6m)"
    if anos <= 2:
        return "Consolidação (7m–2a)"
    if anos <= 5:
        return "Maturidade (3–5a)"
    if anos <= 9:
        return "Lealdade (6–9a)"
    return "Legado (10a+)"


def safra_semestral(data_assinatura, ref_date=None, janela_anos=3):
    if data_assinatura is None or not isinstance(data_assinatura, datetime.datetime):
        return "Erro na data de assinatura"
    if data_assinatura.year < 1950:
        return "Erro na data de assinatura"
    if ref_date is not None and (ref_date - data_assinatura).days / 365.25 > janela_anos:
        return f"Safras anteriores (assinou há mais de {janela_anos} anos)"
    sem = "S1 (jan-jun)" if data_assinatura.month <= 6 else "S2 (jul-dez)"
    return f"{data_assinatura.year} {sem}"


def normaliza_cpf(cpf):
    """Retorna string de 11 dígitos zero-padded, ou None se inválido."""
    if cpf is None:
        return None
    if isinstance(cpf, str):
        s = cpf.strip().lower()
        if s in ("", "null", "none", "nan"):
            return None
        try:
            cpf = float(cpf)
        except ValueError:
            return None
    try:
        cpf_i = int(float(cpf))
    except (TypeError, ValueError):
        return None
    if cpf_i <= 0:
        return None
    return str(cpf_i).zfill(11)


PLANO_FAMILIA_MAP = {
    "PLANO FAMÍLIA ATIVO": "Ativo",
    "PLANO FAMILIA ATIVO": "Ativo",
    "AINDA TEM ESPAÇO NO PLANO FAMÍLIA": "Tem espaço",
    "AINDA TEM ESPACO NO PLANO FAMILIA": "Tem espaço",
    "PLANO FAMÍLIA INATIVO": "Inativo",
    "PLANO FAMILIA INATIVO": "Inativo",
    "NÃO TEM PLANO FAMÍLIA": "Não tem direito",
    "NAO TEM PLANO FAMILIA": "Não tem direito",
}


def mapeia_plano_familia(v):
    if v is None:
        return "Indefinido (dado ausente)"
    key = str(v).strip().upper()
    return PLANO_FAMILIA_MAP.get(key, "Indefinido (dado ausente)") if key != "#N/A" else "Indefinido (dado ausente)"


SITUACAO_CANCELADA = {"CANCELADO", "DESATIVADO PARA SEMPRE"}
SITUACAO_ATIVA = {"ATIVO", "SUSPENSA", "PERIODO ENTREGUE"}


def classifica_situacao(v):
    if v is None:
        return "Outros/Indefinido"
    s = str(v).strip().upper()
    if s in SITUACAO_ATIVA:
        return "Ativa"
    if s in SITUACAO_CANCELADA:
        return "Cancelada"
    return "Outros/Indefinido"


# ---------------------------------------------------------------------
# A partir de maio/2024, a fonte de vouchers passou a chegar em CSV
# (colunas: Nome, CPF, Marca, Título, Descrição do Cupom, Tipo de
# Cupom, Data em texto pt-BR "31 de mai. de 2024", Produtos), diferente
# do xlsx antigo usado em jul/2023-abr/2024. O ETL detecta o formato
# pela extensão do arquivo (ver carga_mensal.py).
# ---------------------------------------------------------------------
MESES_PT = {
    "jan": 1, "fev": 2, "mar": 3, "abr": 4, "mai": 5, "jun": 6,
    "jul": 7, "ago": 8, "set": 9, "out": 10, "nov": 11, "dez": 12,
}


def parse_data_pt(texto):
    if not texto:
        return None
    partes = str(texto).replace(".", "").strip().split(" de ")
    if len(partes) != 3:
        return None
    dia_s, mes_s, ano_s = partes
    mes = MESES_PT.get(mes_s.strip().lower()[:3])
    if mes is None:
        return None
    try:
        return datetime.datetime(int(ano_s), mes, int(dia_s))
    except ValueError:
        return None
