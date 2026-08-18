from __future__ import annotations

import argparse
from dataclasses import dataclass
from datetime import datetime
import json
from pathlib import Path
from typing import Iterable

import pandas as pd
from openpyxl.utils import get_column_letter


BASE_DIR = Path(__file__).resolve().parent
HISTORICO_DIR = BASE_DIR / "historico"
ENTRADA_DIR = BASE_DIR / "entrada"

RAW_RENAME_MAP = {
    "Núm. Orç.": "ID_Orcamento",
    "Num. Orc.": "ID_Orcamento",
    "Data de Criação": "Data",
    "Data de Criacao": "Data",
    "Data de Faturamento": "Data de Faturamento",
    "CNPJ/ CPF": "CNPJ",
    "CNPJ/CPF": "CNPJ",
}


@dataclass
class InputPaths:
    oportunidades: Path | None
    oportunidades_brutas: Path
    oportunidades_kpi: Path | None
    itens: Path


@dataclass
class AnalysisPackage:
    oportunidades: pd.DataFrame
    revisados: pd.DataFrame
    resumo: pd.DataFrame
    funil_geral: pd.DataFrame
    funil_vendedor: pd.DataFrame
    vendedores: pd.DataFrame
    itens_revisados: pd.DataFrame
    status_revisados: pd.DataFrame


def find_latest(pattern: str, directory: Path) -> Path:
    files = sorted(directory.glob(pattern), key=lambda p: p.stat().st_mtime, reverse=True)
    if not files:
        raise FileNotFoundError(f"Nenhum arquivo encontrado em {directory} com padrao {pattern}")
    return files[0]


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Gera analise de aprovacao do gestor para apresentacao comercial."
    )
    parser.add_argument(
        "-o",
        "--oportunidades",
        type=Path,
        help="Arquivo de oportunidades reais (.xlsx). Se omitido, usa o mais recente em historico.",
    )
    parser.add_argument(
        "--oportunidades-brutas",
        type=Path,
        help="Arquivo DAX1 tratado antes da remocao de ruidos (.xlsx). Se omitido, usa o mais recente em entrada.",
    )
    parser.add_argument(
        "--oportunidades-kpi",
        type=Path,
        help="Arquivo tratado do Bloco 2 para calculo dos KPIs de revisao.",
    )
    parser.add_argument(
        "-i",
        "--itens",
        type=Path,
        help="Arquivo de itens (.xlsx). Se omitido, usa o mais recente em entrada.",
    )
    parser.add_argument(
        "--sheet-oportunidades",
        default="Lista_Oportunidades_Reais",
        help="Nome da aba de oportunidades.",
    )
    parser.add_argument(
        "--output",
        type=Path,
        help="Arquivo Excel de saida.",
    )
    parser.add_argument(
        "--output-json",
        type=Path,
        help="Arquivo JSON de saida.",
    )
    return parser.parse_args()


def resolve_inputs(args: argparse.Namespace) -> InputPaths:
    oportunidades = args.oportunidades
    oportunidades_brutas = args.oportunidades_brutas or find_latest("dax1_remocao_ruidos_*.xlsx", ENTRADA_DIR)
    oportunidades_kpi = args.oportunidades_kpi
    itens = args.itens or find_latest("dax2_itens_orcamento_*.xlsx", ENTRADA_DIR)
    return InputPaths(
        oportunidades=oportunidades,
        oportunidades_brutas=oportunidades_brutas,
        oportunidades_kpi=oportunidades_kpi,
        itens=itens,
    )


def fmt_int(v: int | float) -> str:
    return f"{int(v):,}".replace(",", ".")


def fmt_brl(v: int | float) -> str:
    return "R$ " + f"{float(v):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def fmt_pct(v: int | float) -> str:
    return f"{float(v):.2f}%".replace(".", ",")


def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    return df


def load_oportunidades(path: Path, sheet_name: str) -> pd.DataFrame:
    df = pd.read_excel(path, sheet_name=sheet_name)
    df = normalize_columns(df)
    required = [
        "ID_Orcamento",
        "Data",
        "Cliente",
        "CNPJ",
        "Tipo Cliente",
        "Vendedor",
        "ETAPA 1 FUNIL Orçamentos Emitidos",
        "ETAPA 3 FUNIL Aprovados pelo Cliente",
        "ETAPA 4 FUNIL Pedidos Faturados",
        "Passou por Revisão Gestor",
        "Valor",
        "Status Atual",
    ]
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError(f"Colunas obrigatorias ausentes em oportunidades: {missing}")

    df["ID_Orcamento"] = pd.to_numeric(df["ID_Orcamento"], errors="coerce")
    df["Data"] = pd.to_datetime(df["Data"], errors="coerce")
    if "Data de Faturamento" in df.columns:
        df["Data de Faturamento"] = pd.to_datetime(df["Data de Faturamento"], errors="coerce")
    else:
        df["Data de Faturamento"] = pd.NaT
    for col in [
        "ETAPA 1 FUNIL Orçamentos Emitidos",
        "ETAPA 3 FUNIL Aprovados pelo Cliente",
        "ETAPA 4 FUNIL Pedidos Faturados",
        "Passou por Revisão Gestor",
    ]:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0).astype(int)
    if "Total de Revisões" not in df.columns:
        df["Total de Revisões"] = 0
    df["Total de Revisões"] = pd.to_numeric(df["Total de Revisões"], errors="coerce").fillna(0).astype(int)
    df["Valor"] = pd.to_numeric(df["Valor"], errors="coerce").fillna(0.0)
    df["Vendedor"] = df["Vendedor"].astype(str).str.strip()
    df["Tipo Cliente"] = df["Tipo Cliente"].astype(str).str.strip()
    df["Status Atual"] = df["Status Atual"].astype(str).str.strip()
    return df


def load_oportunidades_brutas(path: Path, sheet_name: str = "DAX1_Tratado") -> pd.DataFrame:
    df = pd.read_excel(path, sheet_name=sheet_name)
    df = normalize_columns(df).rename(columns=RAW_RENAME_MAP)
    required = [
        "ID_Orcamento",
        "Data",
        "Cliente",
        "CNPJ",
        "Tipo Cliente",
        "Vendedor",
        "ETAPA 1 FUNIL Orçamentos Emitidos",
        "ETAPA 3 FUNIL Aprovados pelo Cliente",
        "ETAPA 4 FUNIL Pedidos Faturados",
        "Passou por Revisão Gestor",
        "Valor",
        "Status Atual",
    ]
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError(f"Colunas obrigatorias ausentes em oportunidades brutas: {missing}")

    df["ID_Orcamento"] = pd.to_numeric(df["ID_Orcamento"], errors="coerce")
    df["Data"] = pd.to_datetime(df["Data"], errors="coerce")
    if "Data de Faturamento" in df.columns:
        df["Data de Faturamento"] = pd.to_datetime(df["Data de Faturamento"], errors="coerce")
    else:
        df["Data de Faturamento"] = pd.NaT

    for col in [
        "ETAPA 1 FUNIL Orçamentos Emitidos",
        "ETAPA 2 FUNIL Orçamentos em Negociação",
        "ETAPA 3 FUNIL Aprovados pelo Cliente",
        "ETAPA 4 FUNIL Pedidos Faturados",
        "Passou por Revisão Gestor",
    ]:
        if col not in df.columns:
            df[col] = 0
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0).astype(int)
    if "Total de Revisões" not in df.columns:
        df["Total de Revisões"] = 0
    df["Total de Revisões"] = pd.to_numeric(df["Total de Revisões"], errors="coerce").fillna(0).astype(int)

    df["Valor"] = pd.to_numeric(df["Valor"], errors="coerce").fillna(0.0)
    df["Vendedor"] = df["Vendedor"].astype(str).str.strip()
    df["Tipo Cliente"] = df["Tipo Cliente"].astype(str).str.strip()
    df["Status Atual"] = df["Status Atual"].astype(str).str.strip()
    df["Cliente"] = df["Cliente"].astype(str).str.strip()
    df["CNPJ"] = df["CNPJ"].astype(str).str.strip()
    return df


def load_itens(path: Path) -> pd.DataFrame:
    df = pd.read_excel(path)
    df = normalize_columns(df)
    required = ["IDOrcamentoPrinc", "CodigoErp", "Descricao", "Quantidade", "Valor", "VlTot"]
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError(f"Colunas obrigatorias ausentes em itens: {missing}")
    df["IDOrcamentoPrinc"] = pd.to_numeric(df["IDOrcamentoPrinc"], errors="coerce")
    df["Quantidade"] = pd.to_numeric(df["Quantidade"], errors="coerce").fillna(0.0)
    df["Valor"] = pd.to_numeric(df["Valor"], errors="coerce").fillna(0.0)
    df["VlTot"] = pd.to_numeric(df["VlTot"], errors="coerce").fillna(0.0)
    df["CodigoErp"] = df["CodigoErp"].astype(str).str.strip()
    df["Descricao"] = df["Descricao"].astype(str).str.strip()
    return df


def build_resumo(oportunidades: pd.DataFrame, revisados: pd.DataFrame) -> pd.DataFrame:
    total = oportunidades["ID_Orcamento"].nunique()
    total_valor = oportunidades.groupby("ID_Orcamento", as_index=False)["Valor"].max()["Valor"].sum()
    qtd_revisados = revisados["ID_Orcamento"].nunique()
    valor_revisados = revisados.groupby("ID_Orcamento", as_index=False)["Valor"].max()["Valor"].sum()
    qtd_faturados_revisados = revisados.loc[
        revisados["ETAPA 4 FUNIL Pedidos Faturados"] == 1, "ID_Orcamento"
    ].nunique()
    valor_faturados_revisados = revisados.loc[
        revisados["ETAPA 4 FUNIL Pedidos Faturados"] == 1
    ].groupby("ID_Orcamento", as_index=False)["Valor"].max()["Valor"].sum()

    rows = [
        ("Total de orcamentos analisados", total, fmt_int(total)),
        ("Valor total analisado", total_valor, fmt_brl(total_valor)),
        ("Orcamentos com aprovacao do gestor", qtd_revisados, fmt_int(qtd_revisados)),
        ("Valor dos orcamentos com aprovacao do gestor", valor_revisados, fmt_brl(valor_revisados)),
        (
            "% de orcamentos com aprovacao do gestor sobre o total",
            (qtd_revisados / total * 100) if total else 0,
            fmt_pct((qtd_revisados / total * 100) if total else 0),
        ),
        (
            "% de valor com aprovacao do gestor sobre o total",
            (valor_revisados / total_valor * 100) if total_valor else 0,
            fmt_pct((valor_revisados / total_valor * 100) if total_valor else 0),
        ),
        ("Orcamentos revisados que faturaram", qtd_faturados_revisados, fmt_int(qtd_faturados_revisados)),
        (
            "Valor faturado dos orcamentos revisados",
            valor_faturados_revisados,
            fmt_brl(valor_faturados_revisados),
        ),
        (
            "Win rate dos revisados por quantidade",
            (qtd_faturados_revisados / qtd_revisados * 100) if qtd_revisados else 0,
            fmt_pct((qtd_faturados_revisados / qtd_revisados * 100) if qtd_revisados else 0),
        ),
        (
            "Win rate dos revisados por valor",
            (valor_faturados_revisados / valor_revisados * 100) if valor_revisados else 0,
            fmt_pct((valor_faturados_revisados / valor_revisados * 100) if valor_revisados else 0),
        ),
    ]

    return pd.DataFrame(rows, columns=["Indicador", "Valor_Numero", "Valor_Formatado"])


def summarize_stage(df: pd.DataFrame) -> tuple[int, float]:
    if df.empty:
        return 0, 0.0
    base = df.groupby("ID_Orcamento", as_index=False)["Valor"].max()
    return int(base["ID_Orcamento"].nunique()), float(base["Valor"].sum())


def get_easter_sunday(year: int) -> pd.Timestamp:
    a = year % 19
    b = year // 100
    c = year % 100
    d = b // 4
    e = b % 4
    f = (b + 8) // 25
    g = (b - f + 1) // 3
    h = (19 * a + b - d - g + 15) % 30
    i = c // 4
    k = c % 4
    l = (32 + 2 * e + 2 * i - h - k) % 7
    m = (a + 11 * h + 22 * l) // 451
    month = (h + l - 7 * m + 114) // 31
    day = ((h + l - 7 * m + 114) % 31) + 1
    return pd.Timestamp(year=year, month=month, day=day)


def get_brazil_national_holidays(year: int) -> set[pd.Timestamp]:
    easter = get_easter_sunday(year)
    return {
        pd.Timestamp(year=year, month=1, day=1),
        easter - pd.Timedelta(days=2),
        pd.Timestamp(year=year, month=4, day=21),
        pd.Timestamp(year=year, month=5, day=1),
        pd.Timestamp(year=year, month=9, day=7),
        pd.Timestamp(year=year, month=10, day=12),
        pd.Timestamp(year=year, month=11, day=2),
        pd.Timestamp(year=year, month=11, day=15),
        pd.Timestamp(year=year, month=11, day=20),
        pd.Timestamp(year=year, month=12, day=25),
    }


def is_business_day(date: pd.Timestamp) -> bool:
    if pd.isna(date):
        return False
    if date.weekday() >= 5:
        return False
    return date.normalize() not in get_brazil_national_holidays(date.year)


def get_penultimate_business_day(year: int, month: int) -> pd.Timestamp:
    cursor = pd.Timestamp(year=year, month=month, day=1) + pd.offsets.MonthEnd(0)
    business_days: list[pd.Timestamp] = []
    while cursor.month == month:
        if is_business_day(cursor):
            business_days.append(cursor.normalize())
            if len(business_days) == 2:
                return business_days[1]
        cursor -= pd.Timedelta(days=1)
    raise ValueError(f"Nao foi possivel identificar o penultimo dia util de {month}/{year}")


def resolve_commercial_period(date: pd.Timestamp) -> tuple[pd.Timestamp, pd.Timestamp]:
    current_close = get_penultimate_business_day(date.year, date.month)
    if date.normalize() <= current_close:
        previous_month_date = date - pd.DateOffset(months=1)
        previous_close = get_penultimate_business_day(previous_month_date.year, previous_month_date.month)
        return previous_close + pd.Timedelta(days=1), current_close

    next_month_date = date + pd.DateOffset(months=1)
    next_close = get_penultimate_business_day(next_month_date.year, next_month_date.month)
    return current_close + pd.Timedelta(days=1), next_close


def format_month_year_pt(date: pd.Timestamp) -> str:
    meses = ["Jan", "Fev", "Mar", "Abr", "Mai", "Jun", "Jul", "Ago", "Set", "Out", "Nov", "Dez"]
    return f"{meses[date.month - 1]}/{str(date.year)[-2:]}"


def build_historico_6_meses(oportunidades: pd.DataFrame) -> list[dict]:
    base = oportunidades.copy()
    base = base[base["ETAPA 1 FUNIL Orçamentos Emitidos"] == 1].copy()
    base = base[pd.notna(base["Data"])].copy()
    if base.empty:
        return []

    base["Periodo_Comercial"] = base["Data"].apply(resolve_commercial_period)
    base["Periodo_Inicio"] = base["Periodo_Comercial"].apply(lambda p: p[0])
    base["Periodo_Fim"] = base["Periodo_Comercial"].apply(lambda p: p[1])

    periodos = (
        base[["Periodo_Inicio", "Periodo_Fim"]]
        .drop_duplicates()
        .sort_values(["Periodo_Fim", "Periodo_Inicio"])
        .tail(6)
    )

    rows: list[dict] = []
    for _, periodo in periodos.iterrows():
        inicio = periodo["Periodo_Inicio"]
        fim = periodo["Periodo_Fim"]
        grupo = base[(base["Periodo_Inicio"] == inicio) & (base["Periodo_Fim"] == fim)].copy()
        revisados = grupo[grupo["Passou por Revisão Gestor"] == 1].copy()
        faturados_com = revisados[revisados["ETAPA 4 FUNIL Pedidos Faturados"] == 1].copy()
        sem_gestor = grupo[grupo["Passou por Revisão Gestor"] == 0].copy()
        faturados_sem = sem_gestor[sem_gestor["ETAPA 4 FUNIL Pedidos Faturados"] == 1].copy()

        emitidos_qtd, emitidos_valor = summarize_stage(grupo)
        revisados_qtd, revisados_valor = summarize_stage(revisados)
        fat_com_qtd, fat_com_valor = summarize_stage(faturados_com)
        fat_sem_qtd, fat_sem_valor = summarize_stage(faturados_sem)

        rows.append(
            {
                "mes": format_month_year_pt(fim),
                "periodo": f"{inicio.strftime('%d/%m/%Y')} a {fim.strftime('%d/%m/%Y')}",
                "emitidos": emitidos_qtd,
                "emitidos_valor": emitidos_valor,
                "revisados": revisados_qtd,
                "revisados_valor": revisados_valor,
                "share": (revisados_qtd / emitidos_qtd * 100) if emitidos_qtd else 0.0,
                "faturados_qtd": fat_com_qtd,
                "faturados_valor": fat_com_valor,
                "sem_fat_qtd": fat_sem_qtd,
                "sem_fat_valor": fat_sem_valor,
            }
        )

    return rows


def build_funil_geral(oportunidades: pd.DataFrame) -> pd.DataFrame:
    emitidos = oportunidades[oportunidades["ETAPA 1 FUNIL Orçamentos Emitidos"] == 1].copy()
    com_gestor = emitidos[emitidos["Passou por Revisão Gestor"] == 1].copy()
    sem_gestor = emitidos[emitidos["Passou por Revisão Gestor"] == 0].copy()
    com_gestor_aprov_cliente = com_gestor[com_gestor["ETAPA 3 FUNIL Aprovados pelo Cliente"] == 1].copy()
    com_gestor_faturado = com_gestor[com_gestor["ETAPA 4 FUNIL Pedidos Faturados"] == 1].copy()
    sem_gestor_aprov_cliente = sem_gestor[sem_gestor["ETAPA 3 FUNIL Aprovados pelo Cliente"] == 1].copy()
    sem_gestor_faturado = sem_gestor[sem_gestor["ETAPA 4 FUNIL Pedidos Faturados"] == 1].copy()

    total_qtd, total_valor = summarize_stage(emitidos)

    etapas = [
        ("Orçamentos Emitidos", emitidos),
        ("Passaram por Aprovação do Gestor", com_gestor),
        ("Não Passaram por Aprovação do Gestor", sem_gestor),
        ("Com Gestor -> Aprovados pelo Cliente", com_gestor_aprov_cliente),
        ("Com Gestor -> Faturados", com_gestor_faturado),
        ("Sem Gestor -> Aprovados pelo Cliente", sem_gestor_aprov_cliente),
        ("Sem Gestor -> Faturados", sem_gestor_faturado),
    ]

    rows = []
    for etapa, base in etapas:
        qtd, valor = summarize_stage(base)
        rows.append(
            {
                "Etapa": etapa,
                "Qtd_Orcamentos": qtd,
                "Valor": valor,
                "Qtd_Orcamentos_Fmt": fmt_int(qtd),
                "Valor_Fmt": fmt_brl(valor),
                "%_Qtd_sobre_Emitidos": (qtd / total_qtd * 100) if total_qtd else 0.0,
                "%_Valor_sobre_Emitidos": (valor / total_valor * 100) if total_valor else 0.0,
                "%_Qtd_sobre_Emitidos_Fmt": fmt_pct((qtd / total_qtd * 100) if total_qtd else 0.0),
                "%_Valor_sobre_Emitidos_Fmt": fmt_pct((valor / total_valor * 100) if total_valor else 0.0),
            }
        )

    return pd.DataFrame(rows)


def build_funil_vendedor(oportunidades: pd.DataFrame) -> pd.DataFrame:
    emitidos = oportunidades[oportunidades["ETAPA 1 FUNIL Orçamentos Emitidos"] == 1].copy()
    rows = []

    for vendedor, base in emitidos.groupby("Vendedor", dropna=False):
        emitidos_qtd, emitidos_valor = summarize_stage(base)
        com_gestor = base[base["Passou por Revisão Gestor"] == 1]
        sem_gestor = base[base["Passou por Revisão Gestor"] == 0]
        com_gestor_aprov = com_gestor[com_gestor["ETAPA 3 FUNIL Aprovados pelo Cliente"] == 1]
        com_gestor_fat = com_gestor[com_gestor["ETAPA 4 FUNIL Pedidos Faturados"] == 1]
        sem_gestor_aprov = sem_gestor[sem_gestor["ETAPA 3 FUNIL Aprovados pelo Cliente"] == 1]
        sem_gestor_fat = sem_gestor[sem_gestor["ETAPA 4 FUNIL Pedidos Faturados"] == 1]

        com_gestor_qtd, com_gestor_valor = summarize_stage(com_gestor)
        sem_gestor_qtd, sem_gestor_valor = summarize_stage(sem_gestor)
        com_gestor_aprov_qtd, com_gestor_aprov_valor = summarize_stage(com_gestor_aprov)
        com_gestor_fat_qtd, com_gestor_fat_valor = summarize_stage(com_gestor_fat)
        sem_gestor_aprov_qtd, sem_gestor_aprov_valor = summarize_stage(sem_gestor_aprov)
        sem_gestor_fat_qtd, sem_gestor_fat_valor = summarize_stage(sem_gestor_fat)

        rows.append(
            {
                "Vendedor": vendedor,
                "Emitidos_Qtd": emitidos_qtd,
                "Emitidos_Valor": emitidos_valor,
                "Com_Gestor_Qtd": com_gestor_qtd,
                "Com_Gestor_Valor": com_gestor_valor,
                "Com_Gestor_Aprov_Cliente_Qtd": com_gestor_aprov_qtd,
                "Com_Gestor_Aprov_Cliente_Valor": com_gestor_aprov_valor,
                "Com_Gestor_Faturado_Qtd": com_gestor_fat_qtd,
                "Com_Gestor_Faturado_Valor": com_gestor_fat_valor,
                "Sem_Gestor_Qtd": sem_gestor_qtd,
                "Sem_Gestor_Valor": sem_gestor_valor,
                "Sem_Gestor_Aprov_Cliente_Qtd": sem_gestor_aprov_qtd,
                "Sem_Gestor_Aprov_Cliente_Valor": sem_gestor_aprov_valor,
                "Sem_Gestor_Faturado_Qtd": sem_gestor_fat_qtd,
                "Sem_Gestor_Faturado_Valor": sem_gestor_fat_valor,
                "%_Com_Gestor_sobre_Emitidos_Qtd": (com_gestor_qtd / emitidos_qtd * 100) if emitidos_qtd else 0.0,
                "%_Com_Gestor_sobre_Emitidos_Valor": (com_gestor_valor / emitidos_valor * 100) if emitidos_valor else 0.0,
                "%_Fat_Com_Gestor_sobre_Com_Gestor_Qtd": (com_gestor_fat_qtd / com_gestor_qtd * 100) if com_gestor_qtd else 0.0,
                "%_Fat_Com_Gestor_sobre_Com_Gestor_Valor": (com_gestor_fat_valor / com_gestor_valor * 100) if com_gestor_valor else 0.0,
                "%_Fat_Sem_Gestor_sobre_Sem_Gestor_Qtd": (sem_gestor_fat_qtd / sem_gestor_qtd * 100) if sem_gestor_qtd else 0.0,
                "%_Fat_Sem_Gestor_sobre_Sem_Gestor_Valor": (sem_gestor_fat_valor / sem_gestor_valor * 100) if sem_gestor_valor else 0.0,
            }
        )

    base = pd.DataFrame(rows)
    if base.empty:
        return base

    for col in [
        "Emitidos_Valor",
        "Com_Gestor_Valor",
        "Com_Gestor_Aprov_Cliente_Valor",
        "Com_Gestor_Faturado_Valor",
        "Sem_Gestor_Valor",
        "Sem_Gestor_Aprov_Cliente_Valor",
        "Sem_Gestor_Faturado_Valor",
    ]:
        fmt_col = f"{col}_Fmt"
        base[fmt_col] = base[col].map(fmt_brl)

    for col in [
        "%_Com_Gestor_sobre_Emitidos_Qtd",
        "%_Com_Gestor_sobre_Emitidos_Valor",
        "%_Fat_Com_Gestor_sobre_Com_Gestor_Qtd",
        "%_Fat_Com_Gestor_sobre_Com_Gestor_Valor",
        "%_Fat_Sem_Gestor_sobre_Sem_Gestor_Qtd",
        "%_Fat_Sem_Gestor_sobre_Sem_Gestor_Valor",
    ]:
        fmt_col = f"{col}_Fmt"
        base[fmt_col] = base[col].map(fmt_pct)

    base = base.sort_values(
        ["Com_Gestor_Qtd", "Emitidos_Qtd", "Com_Gestor_Faturado_Valor"],
        ascending=[False, False, False],
    )
    return base


def merge_revisoes_por_vendedor(
    funil_vendedor: pd.DataFrame, oportunidades_kpi: pd.DataFrame | None
) -> pd.DataFrame:
    """Soma o total de interacoes de revisao (Bloco 2) por vendedor e cruza no funil por vendedor."""
    base = funil_vendedor.copy()
    if base.empty:
        return base

    if (
        oportunidades_kpi is None
        or oportunidades_kpi.empty
        or "Total de Revisões" not in oportunidades_kpi.columns
    ):
        base["Total_Revisoes"] = 0
        return base

    revisoes_vendedor = (
        oportunidades_kpi.groupby("Vendedor", dropna=False)["Total de Revisões"]
        .sum()
        .rename("Total_Revisoes")
        .reset_index()
    )
    base = base.merge(revisoes_vendedor, on="Vendedor", how="left")
    base["Total_Revisoes"] = base["Total_Revisoes"].fillna(0).astype(int)
    return base


def build_vendedores(revisados: pd.DataFrame) -> pd.DataFrame:
    base = revisados.groupby("Vendedor", dropna=False).agg(
        Orcamentos_Revisados=("ID_Orcamento", pd.Series.nunique),
        Valor_Revisado=("Valor", "sum"),
        Orcamentos_Faturados=("ETAPA 4 FUNIL Pedidos Faturados", "sum"),
    ).reset_index()
    base["Win_Rate_Qtd_Revisados_%"] = (
        base["Orcamentos_Faturados"] / base["Orcamentos_Revisados"] * 100
    ).where(base["Orcamentos_Revisados"] > 0, 0)
    base = base.sort_values(["Orcamentos_Revisados", "Valor_Revisado"], ascending=[False, False])
    return base


def build_itens_revisados(revisados: pd.DataFrame, itens: pd.DataFrame) -> pd.DataFrame:
    ids = set(revisados["ID_Orcamento"].dropna().astype(int).tolist())
    itens_rev = itens[itens["IDOrcamentoPrinc"].isin(ids)].copy()
    if itens_rev.empty:
        return pd.DataFrame(
            columns=[
                "CodigoErp",
                "Descricao",
                "Qtd_Orcamentos_Com_Item",
                "Quantidade_Total",
                "Valor_Total_Item",
                "Preco_Medio_Item",
            ]
        )
    base = itens_rev.groupby(["CodigoErp", "Descricao"], dropna=False).agg(
        Qtd_Orcamentos_Com_Item=("IDOrcamentoPrinc", pd.Series.nunique),
        Quantidade_Total=("Quantidade", "sum"),
        Valor_Total_Item=("VlTot", "sum"),
    ).reset_index()
    base["Preco_Medio_Item"] = (
        base["Valor_Total_Item"] / base["Quantidade_Total"]
    ).where(base["Quantidade_Total"] > 0, 0)
    base = base.sort_values(
        ["Qtd_Orcamentos_Com_Item", "Valor_Total_Item"],
        ascending=[False, False],
    )
    return base


def build_status_revisados(revisados: pd.DataFrame) -> pd.DataFrame:
    base = revisados.groupby("Status Atual", dropna=False).agg(
        Orcamentos=("ID_Orcamento", pd.Series.nunique),
        Valor=("Valor", "sum"),
    ).reset_index()
    total = base["Orcamentos"].sum()
    if total > 0:
        base["Share_Qtd_%"] = base["Orcamentos"] / total * 100
    else:
        base["Share_Qtd_%"] = 0.0
    base = base.sort_values(["Orcamentos", "Valor"], ascending=[False, False])
    return base


def build_lacunas() -> pd.DataFrame:
    rows = [
        {
            "Pergunta": "Quais vendedores pedem mais descontos",
            "Status": "Parcial / proxy",
            "Leitura atual possivel": "Usar 'Passou por Revisão Gestor' como proxy de pedido de excecao comercial.",
            "Campo recomendado": "Indicador explicito de desconto solicitado ou motivo da revisao.",
        },
        {
            "Pergunta": "Quanto em desconto e dado no item pelo supervisor",
            "Status": "Nao respondida com a base atual",
            "Leitura atual possivel": "Nao ha valor base x valor aprovado para medir desconto concedido.",
            "Campo recomendado": "Preco original, preco aprovado, desconto solicitado %, desconto aprovado % ou valor do abatimento.",
        },
    ]
    return pd.DataFrame(rows)


def build_analysis_package(oportunidades: pd.DataFrame, itens: pd.DataFrame) -> AnalysisPackage:
    revisados = oportunidades[oportunidades["Passou por Revisão Gestor"] == 1].copy()
    return AnalysisPackage(
        oportunidades=oportunidades,
        revisados=revisados,
        resumo=build_resumo(oportunidades, revisados),
        funil_geral=build_funil_geral(oportunidades),
        funil_vendedor=build_funil_vendedor(oportunidades),
        vendedores=build_vendedores(revisados),
        itens_revisados=build_itens_revisados(revisados, itens),
        status_revisados=build_status_revisados(revisados),
    )


def filter_oportunidades_periodo(
    oportunidades: pd.DataFrame,
    inicio: pd.Timestamp | None,
    fim: pd.Timestamp | None,
) -> pd.DataFrame:
    base = oportunidades.copy()
    if inicio is not None:
        base = base[base["Data"] >= inicio]
    if fim is not None:
        base = base[base["Data"] <= fim]
    return base.copy()


def build_revisao_kpis(oportunidades_kpi: pd.DataFrame | None) -> dict:
    if oportunidades_kpi is None or oportunidades_kpi.empty:
        return {
            "qtd_orcamentos_com_revisao": 0,
            "qtd_revisoes_total": None,
            "media_revisoes_por_orcamento": None,
        }

    revisados = oportunidades_kpi[oportunidades_kpi["Passou por Revisão Gestor"] == 1].copy()
    qtd_orcamentos = int(revisados["ID_Orcamento"].nunique())

    if "Total de Revisões" not in revisados.columns:
        return {
            "qtd_orcamentos_com_revisao": qtd_orcamentos,
            "qtd_revisoes_total": None,
            "media_revisoes_por_orcamento": None,
        }

    revisoes_por_orcamento = revisados.groupby("ID_Orcamento", as_index=False)["Total de Revisões"].max()
    qtd_revisoes_total = int(revisoes_por_orcamento["Total de Revisões"].sum()) if not revisoes_por_orcamento.empty else 0
    media_revisoes = round(qtd_revisoes_total / qtd_orcamentos, 2) if qtd_orcamentos else 0.0
    return {
        "qtd_orcamentos_com_revisao": qtd_orcamentos,
        "qtd_revisoes_total": qtd_revisoes_total,
        "media_revisoes_por_orcamento": media_revisoes,
    }


def package_to_json_dict(package: AnalysisPackage) -> dict:
    periodo_inicio = package.oportunidades["Data"].min()
    periodo_fim = package.oportunidades["Data"].max()
    return {
        "periodo_analisado": {
            "inicio": periodo_inicio.strftime("%d/%m/%Y") if pd.notna(periodo_inicio) else "",
            "fim": periodo_fim.strftime("%d/%m/%Y") if pd.notna(periodo_fim) else "",
        },
        "resumo": {
            row["Indicador"]: row["Valor_Formatado"]
            for _, row in package.resumo.iterrows()
        },
        "funil_geral": package.funil_geral.to_dict(orient="records"),
        "funil_por_vendedor": package.funil_vendedor.head(50).to_dict(orient="records"),
        "top_vendedores_revisao": package.vendedores.head(10).to_dict(orient="records"),
        "top_itens_revisao": package.itens_revisados.head(20).to_dict(orient="records"),
        "status_revisados": package.status_revisados.to_dict(orient="records"),
        "qtd_orcamentos_com_revisao": int(package.revisados["ID_Orcamento"].nunique()),
    }


def build_comparativo_visoes(sem_ruido: AnalysisPackage, com_ruido: AnalysisPackage) -> pd.DataFrame:
    base = [
        "Total de orcamentos analisados",
        "Valor total analisado",
        "Orcamentos com aprovacao do gestor",
        "Valor dos orcamentos com aprovacao do gestor",
        "% de orcamentos com aprovacao do gestor sobre o total",
        "% de valor com aprovacao do gestor sobre o total",
        "Orcamentos revisados que faturaram",
        "Valor faturado dos orcamentos revisados",
        "Win rate dos revisados por quantidade",
        "Win rate dos revisados por valor",
    ]
    resumo_sem = sem_ruido.resumo.set_index("Indicador")
    resumo_com = com_ruido.resumo.set_index("Indicador")
    rows = []
    for indicador in base:
        sem_row = resumo_sem.loc[indicador]
        com_row = resumo_com.loc[indicador]
        rows.append(
            {
                "Indicador": indicador,
                "Sem_Ruido": sem_row["Valor_Formatado"],
                "Com_Ruido": com_row["Valor_Formatado"],
            }
        )
    return pd.DataFrame(rows)


def build_json_payload(
    oportunidades_path: Path | None,
    oportunidades_brutas_path: Path,
    oportunidades_kpi_path: Path | None,
    itens_path: Path,
    sem_ruido: AnalysisPackage,
    com_ruido: AnalysisPackage,
    historico_6_meses: list[dict],
    revisao_kpis: dict,
) -> dict:
    sem_ruido_json = package_to_json_dict(sem_ruido)
    com_ruido_json = package_to_json_dict(com_ruido)
    return {
        "gerado_em": datetime.now().strftime("%d/%m/%Y %H:%M"),
        "fonte": {
            "oportunidades_sem_ruido": str(oportunidades_path) if oportunidades_path else "",
            "oportunidades_com_ruido": str(oportunidades_brutas_path),
            "oportunidades_kpi": str(oportunidades_kpi_path) if oportunidades_kpi_path else "",
            "itens": str(itens_path),
        },
        "periodo_analisado": sem_ruido_json["periodo_analisado"],
        "resumo": sem_ruido_json["resumo"],
        "funil_geral": sem_ruido_json["funil_geral"],
        "funil_por_vendedor": sem_ruido_json["funil_por_vendedor"],
        "top_vendedores_revisao": sem_ruido_json["top_vendedores_revisao"],
        "top_itens_revisao": sem_ruido_json["top_itens_revisao"],
        "status_revisados": sem_ruido_json["status_revisados"],
        "observacao": (
            "Pedidos de desconto e valor de desconto concedido ainda nao estao disponiveis "
            "como dado explicito na base atual."
        ),
        "qtd_orcamentos_com_revisao": revisao_kpis["qtd_orcamentos_com_revisao"],
        "qtd_revisoes_total": revisao_kpis["qtd_revisoes_total"],
        "media_revisoes_por_orcamento": revisao_kpis["media_revisoes_por_orcamento"],
        "historico_6_meses": historico_6_meses,
        "visoes": {
            "sem_ruido": sem_ruido_json,
            "com_ruido": com_ruido_json,
        },
    }


def autosize_excel_columns(writer: pd.ExcelWriter, sheet_name: str, df: pd.DataFrame) -> None:
    worksheet = writer.sheets[sheet_name]
    for idx, col in enumerate(df.columns, start=1):
        max_len = max([len(str(col))] + [len(str(v)) for v in df[col].head(200).fillna("")])
        worksheet.column_dimensions[get_column_letter(idx)].width = min(max_len + 2, 40)


def main() -> None:
    args = parse_args()
    inputs = resolve_inputs(args)

    itens = load_itens(inputs.itens)
    oportunidades_com_ruido_amplas = load_oportunidades_brutas(inputs.oportunidades_brutas)
    historico_6_meses = build_historico_6_meses(oportunidades_com_ruido_amplas)

    if inputs.oportunidades:
        oportunidades_sem_ruido = load_oportunidades(inputs.oportunidades, args.sheet_oportunidades)
        periodo_inicio = oportunidades_sem_ruido["Data"].min()
        periodo_fim = oportunidades_sem_ruido["Data"].max()
        oportunidades_com_ruido = filter_oportunidades_periodo(
            oportunidades_com_ruido_amplas,
            periodo_inicio if pd.notna(periodo_inicio) else None,
            periodo_fim if pd.notna(periodo_fim) else None,
        )
    else:
        # Quando o fluxo roda em com_ruido sem base sem ruido explicita,
        # a trilha principal deve refletir apenas o mes comercial atual da rodada.
        data_maxima = oportunidades_com_ruido_amplas["Data"].max()
        if pd.isna(data_maxima):
            raise ValueError("Nao foi possivel identificar a data maxima da base com ruido.")

        periodo_comercial_inicio, _ = resolve_commercial_period(pd.Timestamp(data_maxima))
        oportunidades_sem_ruido = filter_oportunidades_periodo(
            oportunidades_com_ruido_amplas,
            periodo_comercial_inicio,
            pd.Timestamp(data_maxima),
        )
        oportunidades_com_ruido = oportunidades_sem_ruido.copy()

    pacote_sem_ruido = build_analysis_package(oportunidades_sem_ruido, itens)
    pacote_com_ruido = build_analysis_package(oportunidades_com_ruido, itens)
    oportunidades_kpi = load_oportunidades_brutas(inputs.oportunidades_kpi) if inputs.oportunidades_kpi else None
    if oportunidades_kpi is not None:
        periodo_kpi_inicio = oportunidades_sem_ruido["Data"].min()
        periodo_kpi_fim = oportunidades_sem_ruido["Data"].max()
        oportunidades_kpi = filter_oportunidades_periodo(
            oportunidades_kpi,
            periodo_kpi_inicio if pd.notna(periodo_kpi_inicio) else None,
            periodo_kpi_fim if pd.notna(periodo_kpi_fim) else None,
        )
    revisao_kpis = build_revisao_kpis(oportunidades_kpi)
    pacote_sem_ruido.funil_vendedor = merge_revisoes_por_vendedor(pacote_sem_ruido.funil_vendedor, oportunidades_kpi)
    pacote_com_ruido.funil_vendedor = merge_revisoes_por_vendedor(pacote_com_ruido.funil_vendedor, oportunidades_kpi)
    comparativo_visoes = build_comparativo_visoes(pacote_sem_ruido, pacote_com_ruido)
    lacunas = build_lacunas()

    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_excel = args.output or (HISTORICO_DIR / f"analise_aprovacao_gestor_{ts}.xlsx")
    output_json = args.output_json or (BASE_DIR / "alertas" / f"analise_aprovacao_gestor_{ts}.json")

    with pd.ExcelWriter(output_excel, engine="openpyxl") as writer:
        pacote_sem_ruido.resumo.to_excel(writer, sheet_name="Resumo_Executivo", index=False)
        comparativo_visoes.to_excel(writer, sheet_name="Resumo_Comparativo", index=False)
        pacote_sem_ruido.funil_geral.to_excel(writer, sheet_name="Funil_Gestor_Geral", index=False)
        pacote_com_ruido.funil_geral.to_excel(writer, sheet_name="Funil_Geral_Com_Ruido", index=False)
        pacote_sem_ruido.funil_vendedor.to_excel(writer, sheet_name="Funil_Gestor_Vendedor", index=False)
        pacote_com_ruido.funil_vendedor.to_excel(writer, sheet_name="Funil_Vend_Com_Ruido", index=False)
        pacote_sem_ruido.vendedores.to_excel(writer, sheet_name="Vendedores_Revisao", index=False)
        pacote_com_ruido.vendedores.to_excel(writer, sheet_name="Vendedores_Com_Ruido", index=False)
        pacote_sem_ruido.itens_revisados.to_excel(writer, sheet_name="Itens_Revisao", index=False)
        pacote_com_ruido.itens_revisados.to_excel(writer, sheet_name="Itens_Com_Ruido", index=False)
        pacote_sem_ruido.status_revisados.to_excel(writer, sheet_name="Status_Revisados", index=False)
        pacote_com_ruido.status_revisados.to_excel(writer, sheet_name="Status_Com_Ruido", index=False)
        lacunas.to_excel(writer, sheet_name="Lacunas_Desconto", index=False)

        for sheet_name, df in [
            ("Resumo_Executivo", pacote_sem_ruido.resumo),
            ("Resumo_Comparativo", comparativo_visoes),
            ("Funil_Gestor_Geral", pacote_sem_ruido.funil_geral),
            ("Funil_Geral_Com_Ruido", pacote_com_ruido.funil_geral),
            ("Funil_Gestor_Vendedor", pacote_sem_ruido.funil_vendedor),
            ("Funil_Vend_Com_Ruido", pacote_com_ruido.funil_vendedor),
            ("Vendedores_Revisao", pacote_sem_ruido.vendedores),
            ("Vendedores_Com_Ruido", pacote_com_ruido.vendedores),
            ("Itens_Revisao", pacote_sem_ruido.itens_revisados),
            ("Itens_Com_Ruido", pacote_com_ruido.itens_revisados),
            ("Status_Revisados", pacote_sem_ruido.status_revisados),
            ("Status_Com_Ruido", pacote_com_ruido.status_revisados),
            ("Lacunas_Desconto", lacunas),
        ]:
            autosize_excel_columns(writer, sheet_name, df)

    payload = build_json_payload(
        inputs.oportunidades,
        inputs.oportunidades_brutas,
        inputs.oportunidades_kpi,
        inputs.itens,
        pacote_sem_ruido,
        pacote_com_ruido,
        historico_6_meses,
        revisao_kpis,
    )
    output_json.parent.mkdir(parents=True, exist_ok=True)
    output_json.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")

    print(f"Excel gerado: {output_excel}")
    print(f"JSON gerado: {output_json}")
    print(f"Oportunidades analisadas (sem ruido): {inputs.oportunidades or 'nao informadas; usando base ampla da rodada'}")
    print(f"Oportunidades analisadas (com ruido): {inputs.oportunidades_brutas}")
    print(f"Itens analisados: {inputs.itens}")


if __name__ == "__main__":
    main()
