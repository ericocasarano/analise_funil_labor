#!/usr/bin/env python
# -*- coding: utf-8 -*-

import argparse
import json
from datetime import datetime
from pathlib import Path

import pandas as pd


TARGET_COLUMNS = [
    "IDOrcamentoPrinc",
    "DtInclusao",
    "CodigoErp",
    "Descricao",
    "Quantidade",
    "Valor",
    "VlTot",
]

RENAME_MAP = {
    "IDOrcamentoPrinc": "IDOrcamentoPrinc",
    "ID_Orcamento": "IDOrcamentoPrinc",
    "Pedido": "IDOrcamentoPrinc",
    "Orcamento": "IDOrcamentoPrinc",
    "DtInclusao": "DtInclusao",
    "Data": "DtInclusao",
    "Data de Inclusao": "DtInclusao",
    "Data de Inclusão": "DtInclusao",
    "CodigoErp": "CodigoErp",
    "Codigo ERP": "CodigoErp",
    "Código ERP": "CodigoErp",
    "Cod Produto": "CodigoErp",
    "Codigo Produto": "CodigoErp",
    "Descricao": "Descricao",
    "Descrição": "Descricao",
    "Produto": "Descricao",
    "Quantidade": "Quantidade",
    "Qtd": "Quantidade",
    "Valor": "Valor",
    "VlTot": "VlTot",
    "Valor Total": "VlTot",
}

INTEGER_COLUMNS = ["IDOrcamentoPrinc"]
DECIMAL_COLUMNS = ["Quantidade", "Valor", "VlTot"]


def parse_args():
    parser = argparse.ArgumentParser(
        description="Converte o JSON de itens gerado pelo Power Automate em XLSX tratado para o pipeline do funil."
    )
    parser.add_argument("-i", "--input", required=True, help="Arquivo JSON salvo pelo Power Automate.")
    parser.add_argument(
        "-o",
        "--output",
        default="",
        help="Arquivo XLSX de saida. Se vazio, salva em entrada/dax2_itens_orcamento_<timestamp>.xlsx.",
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
    return pd.to_numeric(series, errors="coerce").astype("Int64")


def convert_power_automate_json(input_path: Path, output_path: Path) -> pd.DataFrame:
    rows = load_rows(input_path)
    df = pd.DataFrame(rows)
    if df.empty:
        raise ValueError(f"Nenhuma linha encontrada em {input_path}")

    df.columns = [clean_column_name(col) for col in df.columns]
    df = df.rename(columns={col: RENAME_MAP.get(col, col) for col in df.columns})

    for col in TARGET_COLUMNS:
        if col not in df.columns:
            df[col] = 0 if col in INTEGER_COLUMNS or col in DECIMAL_COLUMNS else ""

    df = df[TARGET_COLUMNS].copy()
    df["DtInclusao"] = pd.to_datetime(df["DtInclusao"], errors="coerce")

    for col in DECIMAL_COLUMNS:
        df[col] = normalize_decimal_series(df[col]).fillna(0)

    for col in INTEGER_COLUMNS:
        df[col] = normalize_integer_series(df[col])

    df["CodigoErp"] = df["CodigoErp"].astype("string").str.replace(r"\.0$", "", regex=True).fillna("").str.strip()
    df["Descricao"] = df["Descricao"].astype("string").fillna("").str.strip()

    output_path.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="DAX2_Tratado")

    return df


def main():
    args = parse_args()
    input_path = Path(args.input).resolve()
    timestamp = args.timestamp or datetime.now().strftime("%Y%m%d_%H%M%S")
    output_path = (
        Path(args.output).resolve()
        if args.output
        else (Path.cwd() / "entrada" / f"dax2_itens_orcamento_{timestamp}.xlsx")
    )

    df = convert_power_automate_json(input_path, output_path)
    print(f"Arquivo tratado gerado em: {output_path}")
    print(f"Linhas: {len(df)}")
    print(f"Colunas: {list(df.columns)}")


if __name__ == "__main__":
    main()
