#!/usr/bin/env python
# -*- coding: utf-8 -*-

import argparse
import json
from datetime import datetime
from pathlib import Path

import pandas as pd


TARGET_COLUMNS = [
    "Núm. Orç.",
    "Data de Criação",
    "Data de Faturamento",
    "Cliente",
    "CNPJ/ CPF",
    "Tipo Cliente",
    "Vendedor",
    "ETAPA 1 FUNIL Orçamentos Emitidos",
    "ETAPA 2 FUNIL Orçamentos em Negociação",
    "ETAPA 3 FUNIL Aprovados pelo Cliente",
    "ETAPA 4 FUNIL Pedidos Faturados",
    "Passou por Revisão Gestor",
    "Total de Revisões",
    "Valor",
    "Status Atual",
]

RENAME_MAP = {
    "Pedido": "Núm. Orç.",
    "Número do Pedido": "Núm. Orç.",
    "Numero do Pedido": "Núm. Orç.",
    "Núm. Orç.": "Núm. Orç.",
    "Num. Orc.": "Núm. Orç.",
    "Data": "Data de Criação",
    "Data de Criação": "Data de Criação",
    "Data de Criacao": "Data de Criação",
    "Data de Faturamento": "Data de Faturamento",
    "Data de Faturamento ": "Data de Faturamento",
    "Data Faturamento": "Data de Faturamento",
    "Data_Faturamento": "Data de Faturamento",
    "Dt Faturamento": "Data de Faturamento",
    "Dt_Faturamento": "Data de Faturamento",
    "Cliente": "Cliente",
    "Cnpj": "CNPJ/ CPF",
    "CNPJ": "CNPJ/ CPF",
    "CNPJ/ CPF": "CNPJ/ CPF",
    "CNPJ/CPF": "CNPJ/ CPF",
    "Tipo_Cliente": "Tipo Cliente",
    "Tipo Cliente": "Tipo Cliente",
    "Vendedor": "Vendedor",
    "Nome do Vendedor": "Vendedor",
    "Status": "Status Atual",
    "Descricao": "Status Atual",
    "Status Atual": "Status Atual",
    "Qtd_Orcamentos": "ETAPA 1 FUNIL Orçamentos Emitidos",
    "Orcamentos_Totais": "ETAPA 1 FUNIL Orçamentos Emitidos",
    "ETAPA 1 FUNIL Orçamentos Emitidos": "ETAPA 1 FUNIL Orçamentos Emitidos",
    "Orcamentos_Negociacao": "ETAPA 2 FUNIL Orçamentos em Negociação",
    "ETAPA 2 FUNIL Orçamentos em Negociação": "ETAPA 2 FUNIL Orçamentos em Negociação",
    "Orcamentos_Pedido": "ETAPA 3 FUNIL Aprovados pelo Cliente",
    "ETAPA 3 FUNIL Aprovados pelo Cliente": "ETAPA 3 FUNIL Aprovados pelo Cliente",
    "Orcamentos_Faturado": "ETAPA 4 FUNIL Pedidos Faturados",
    "ETAPA 4 FUNIL Pedidos Faturados": "ETAPA 4 FUNIL Pedidos Faturados",
    "Orcamentos_Revisao_Gestor": "Passou por Revisão Gestor",
    "Passou por Revisão Gestor": "Passou por Revisão Gestor",
    "PassouRevisaoGestor": "Passou por Revisão Gestor",
    "QtdPassagensRevisaoGestor": "Total de Revisões",
    "TotalRevisoes": "Total de Revisões",
    "Total de Revisões": "Total de Revisões",
    "PassouAprovacaoCliente": "ETAPA 3 FUNIL Aprovados pelo Cliente",
    "QtdOrcamentos": "ETAPA 1 FUNIL Orçamentos Emitidos",
    "Faturou": "ETAPA 4 FUNIL Pedidos Faturados",
    "Valor": "Valor",
}

INTEGER_COLUMNS = [
    "Núm. Orç.",
    "ETAPA 1 FUNIL Orçamentos Emitidos",
    "ETAPA 2 FUNIL Orçamentos em Negociação",
    "ETAPA 3 FUNIL Aprovados pelo Cliente",
    "ETAPA 4 FUNIL Pedidos Faturados",
    "Passou por Revisão Gestor",
    "Total de Revisões",
]


def parse_args():
    parser = argparse.ArgumentParser(
        description="Converte o JSON gerado pelo Power Automate em XLSX tratado para o pipeline do funil."
    )
    parser.add_argument("-i", "--input", required=True, help="Arquivo JSON salvo pelo Power Automate.")
    parser.add_argument(
        "-o",
        "--output",
        default="",
        help="Arquivo XLSX de saida. Se vazio, salva em entrada/dax1_remocao_ruidos_<timestamp>.xlsx.",
    )
    parser.add_argument(
        "--timestamp",
        default="",
        help="Timestamp da rodada no formato YYYYMMDD_HHMMSS. Se vazio, gera automaticamente.",
    )
    return parser.parse_args()


def clean_column_name(name: str) -> str:
    text = str(name).strip()
    if text.startswith("[") and text.endswith("]"):
        text = text[1:-1]
    return text


def load_rows(path: Path) -> list[dict]:
    data = json.loads(path.read_text(encoding="utf-8-sig"))

    if isinstance(data, str):
        data = json.loads(data)

    if isinstance(data, list):
        return data

    if isinstance(data, dict):
        for key in ("firstTableRows", "rows", "value"):
            value = data.get(key)
            if isinstance(value, str):
                value = json.loads(value)
            if isinstance(value, list):
                return value

        tables = data.get("results", [{}])[0].get("tables", []) if isinstance(data.get("results"), list) else []
        if tables and isinstance(tables[0].get("rows"), list):
            return tables[0]["rows"]

    raise ValueError("Formato JSON nao reconhecido. Esperado: array de objetos ou objeto com firstTableRows/rows.")


def normalize_decimal_series(series: pd.Series) -> pd.Series:
    if pd.api.types.is_numeric_dtype(series):
        return pd.to_numeric(series, errors="coerce")

    cleaned = (
        series.astype("string")
        .str.strip()
        .replace({"<NA>": pd.NA, "nan": pd.NA, "None": pd.NA, "": pd.NA})
    )
    has_comma = cleaned.str.contains(",", regex=False, na=False)
    cleaned = cleaned.where(~has_comma, cleaned.str.replace(".", "", regex=False).str.replace(",", ".", regex=False))
    return pd.to_numeric(cleaned, errors="coerce")


def normalize_integer_series(series: pd.Series) -> pd.Series:
    return pd.to_numeric(series, errors="coerce").fillna(0).astype("Int64")


def convert_power_automate_json(input_path: Path, output_path: Path) -> pd.DataFrame:
    rows = load_rows(input_path)
    df = pd.DataFrame(rows)
    if df.empty:
        raise ValueError(f"Nenhuma linha encontrada em {input_path}")

    df.columns = [clean_column_name(col) for col in df.columns]
    df = df.rename(columns={col: RENAME_MAP.get(col, col) for col in df.columns})

    for col in TARGET_COLUMNS:
        if col not in df.columns:
            df[col] = 0 if col in INTEGER_COLUMNS or col == "Valor" else ""

    df = df[TARGET_COLUMNS].copy()
    df["Data de Criação"] = pd.to_datetime(df["Data de Criação"], errors="coerce")
    df["Data de Faturamento"] = pd.to_datetime(df["Data de Faturamento"], errors="coerce")
    df["Valor"] = normalize_decimal_series(df["Valor"]).fillna(0)

    for col in INTEGER_COLUMNS:
        df[col] = normalize_integer_series(df[col])

    df["CNPJ/ CPF"] = df["CNPJ/ CPF"].astype("string").str.replace(r"\.0$", "", regex=True)
    df["Cliente"] = df["Cliente"].astype("string").fillna("").str.strip()
    df["Tipo Cliente"] = df["Tipo Cliente"].astype("string").fillna("").str.strip()
    df["Vendedor"] = df["Vendedor"].astype("string").fillna("").str.strip()
    df["Status Atual"] = df["Status Atual"].astype("string").fillna("").str.strip()

    output_path.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="DAX1_Tratado")

    return df


def main():
    args = parse_args()
    input_path = Path(args.input).resolve()
    timestamp = args.timestamp or datetime.now().strftime("%Y%m%d_%H%M%S")
    output_path = (
        Path(args.output).resolve()
        if args.output
        else (Path.cwd() / "entrada" / f"dax1_remocao_ruidos_{timestamp}.xlsx")
    )

    df = convert_power_automate_json(input_path, output_path)
    print(f"Arquivo tratado gerado em: {output_path}")
    print(f"Linhas: {len(df)}")
    print(f"Colunas: {list(df.columns)}")


if __name__ == "__main__":
    main()
