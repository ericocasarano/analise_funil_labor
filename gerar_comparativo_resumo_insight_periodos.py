#!/usr/bin/env python
# -*- coding: utf-8 -*-

import argparse
import json
from datetime import datetime
from pathlib import Path


SUMMARY_PREFIX = "resumo_insight_*.json"

FIELDS_TO_COPY = [
    "titulo",
    "mes_comercial",
    "periodo",
    "atualizado_em",
    "qtd_enviados",
    "qtd_enviados_fmt",
    "valor_enviado",
    "valor_enviado_numero",
    "ticket_medio_gerado",
    "ticket_medio_gerado_numero",
    "qtd_faturados",
    "qtd_faturados_fmt",
    "valor_faturado",
    "valor_faturado_numero",
    "ticket_medio_convertido",
    "ticket_medio_convertido_numero",
    "qtd_nao_convertidas",
    "qtd_nao_convertidas_fmt",
    "valor_nao_convertido",
    "valor_nao_convertido_numero",
    "ticket_medio_nao_convertido",
    "ticket_medio_nao_convertido_numero",
    "win_rate_qtd",
    "win_rate_qtd_numero",
    "win_rate_valor",
    "win_rate_valor_numero",
    "win_rate_qtd_data_faturamento",
    "win_rate_qtd_data_faturamento_numero",
    "win_rate_valor_data_faturamento",
    "win_rate_valor_data_faturamento_numero",
    "qtd_enviados_data_faturamento",
    "qtd_enviados_data_faturamento_fmt",
    "valor_enviado_data_faturamento",
    "valor_enviado_numero_data_faturamento",
    "qtd_faturados_data_faturamento",
    "qtd_faturados_data_faturamento_fmt",
    "valor_faturado_data_faturamento",
    "valor_faturado_numero_data_faturamento",
]


def parse_args():
    parser = argparse.ArgumentParser(
        description="Gera um JSON comparativo de dois resumos de win rate e funil."
    )
    parser.add_argument(
        "--alertas-dir",
        default="alertas",
        help="Diretorio com os arquivos resumo_insight_*.json. Padrao: alertas",
    )
    parser.add_argument(
        "--arquivo-a",
        default="",
        help="Arquivo do periodo A. Se vazio, usa um dos 2 resumos mais recentes.",
    )
    parser.add_argument(
        "--arquivo-b",
        default="",
        help="Arquivo do periodo B. Se vazio, usa um dos 2 resumos mais recentes.",
    )
    parser.add_argument(
        "--dias-uteis-a",
        type=int,
        default=0,
        help="Dias uteis do periodo A. Opcional.",
    )
    parser.add_argument(
        "--dias-uteis-b",
        type=int,
        default=0,
        help="Dias uteis do periodo B. Opcional.",
    )
    parser.add_argument(
        "-o",
        "--output",
        default="",
        help="Arquivo JSON de saida. Se vazio, salva em alertas/comparativo_resumo_insight_card_teams_<timestamp>.json",
    )
    return parser.parse_args()


def load_json(path: Path) -> dict:
    return json.loads(path.read_text(encoding="utf-8-sig"))


def latest_two_files(alertas_dir: Path) -> tuple[Path, Path]:
    files = sorted(alertas_dir.glob(SUMMARY_PREFIX), key=lambda p: p.stat().st_mtime, reverse=True)
    if len(files) < 2:
        raise FileNotFoundError(
            f"Nao foi possivel localizar 2 arquivos {SUMMARY_PREFIX} em {alertas_dir}"
        )
    return files[1], files[0]


def resolve_input_files(args, alertas_dir: Path) -> tuple[Path, Path]:
    if args.arquivo_a and args.arquivo_b:
        return Path(args.arquivo_a).resolve(), Path(args.arquivo_b).resolve()
    return latest_two_files(alertas_dir)


def pct_change(base: float, current: float) -> float | None:
    if base == 0:
        return None
    return ((current - base) / base) * 100.0


def format_pct(value: float | None) -> str:
    if value is None:
        return ""
    return f"{value:.2f}%".replace(".", ",")


def extract_resumo(data: dict) -> dict:
    resumo = {field: data.get(field) for field in FIELDS_TO_COPY}
    resumo["principal_status_aberto"] = data.get("em_aberto", {}).get("principal_status", "")
    resumo["quantidade_principal_status_aberto"] = data.get("em_aberto", {}).get(
        "quantidade_principal_status", 0
    )
    resumo["valor_principal_status_aberto"] = data.get("em_aberto", {}).get(
        "valor_principal_status", ""
    )
    resumo["valor_principal_status_aberto_numero"] = 0.0
    if data.get("em_aberto", {}).get("status_totais"):
        resumo["valor_principal_status_aberto_numero"] = float(
            data["em_aberto"]["status_totais"][0].get("valor_numero", 0.0) or 0.0
        )

    resumo["principal_status_perdido"] = data.get("perdido", {}).get("principal_status", "")
    resumo["quantidade_principal_status_perdido"] = data.get("perdido", {}).get(
        "quantidade_principal_status", 0
    )
    resumo["valor_principal_status_perdido"] = data.get("perdido", {}).get(
        "valor_principal_status", ""
    )
    resumo["valor_principal_status_perdido_numero"] = 0.0
    if data.get("perdido", {}).get("status_totais"):
        resumo["valor_principal_status_perdido_numero"] = float(
            data["perdido"]["status_totais"][0].get("valor_numero", 0.0) or 0.0
        )
    return resumo


def build_comparison_metrics(a: dict, b: dict) -> dict:
    pairs = [
        ("qtd_enviados", "quantidade_enviada"),
        ("valor_enviado_numero", "valor_enviado"),
        ("qtd_faturados", "quantidade_faturada"),
        ("valor_faturado_numero", "valor_faturado"),
        ("qtd_nao_convertidas", "quantidade_nao_convertida"),
        ("valor_nao_convertido_numero", "valor_nao_convertido"),
        ("win_rate_qtd_numero", "win_rate_quantidade"),
        ("win_rate_valor_numero", "win_rate_valor"),
        ("win_rate_qtd_data_faturamento_numero", "win_rate_quantidade_data_faturamento"),
        ("win_rate_valor_data_faturamento_numero", "win_rate_valor_data_faturamento"),
        ("qtd_enviados_data_faturamento", "quantidade_enviada_data_faturamento"),
        ("valor_enviado_numero_data_faturamento", "valor_enviado_data_faturamento"),
        ("qtd_faturados_data_faturamento", "quantidade_faturada_data_faturamento"),
        ("valor_faturado_numero_data_faturamento", "valor_faturado_data_faturamento"),
        ("ticket_medio_gerado_numero", "ticket_medio_gerado"),
        ("ticket_medio_convertido_numero", "ticket_medio_convertido"),
        ("ticket_medio_nao_convertido_numero", "ticket_medio_nao_convertido"),
    ]
    metrics = {}
    for source_key, target_key in pairs:
        a_val = float(a.get(source_key, 0) or 0)
        b_val = float(b.get(source_key, 0) or 0)
        delta = b_val - a_val
        metrics[target_key] = {
            "periodo_a": a_val,
            "periodo_b": b_val,
            "variacao_absoluta": delta,
            "variacao_percentual": pct_change(a_val, b_val),
            "variacao_percentual_fmt": format_pct(pct_change(a_val, b_val)),
        }
    return metrics


def output_path(args, alertas_dir: Path) -> Path:
    if args.output:
        return Path(args.output).resolve()
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    return (alertas_dir / f"comparativo_resumo_insight_card_teams_{ts}.json").resolve()


def main():
    args = parse_args()
    root = Path.cwd()
    alertas_dir = (root / args.alertas_dir).resolve()
    alertas_dir.mkdir(parents=True, exist_ok=True)

    arquivo_a, arquivo_b = resolve_input_files(args, alertas_dir)
    data_a = load_json(arquivo_a)
    data_b = load_json(arquivo_b)

    resumo_a = extract_resumo(data_a)
    resumo_b = extract_resumo(data_b)

    comparativo = {
        "titulo": "Comparativo de Win Rate e Funil por Periodos",
        "atualizado_em": datetime.now().strftime("%d/%m/%Y %H:%M"),
        "fonte": {
            "arquivo_a": arquivo_a.name,
            "arquivo_b": arquivo_b.name,
        },
        "comparativo": {
            "periodo_a": {
                "mes_comercial": data_a.get("mes_comercial", ""),
                "label": data_a.get("periodo", ""),
                "dias_uteis": args.dias_uteis_a,
                "resumo": resumo_a,
            },
            "periodo_b": {
                "mes_comercial": data_b.get("mes_comercial", ""),
                "label": data_b.get("periodo", ""),
                "dias_uteis": args.dias_uteis_b,
                "resumo": resumo_b,
            },
            "metricas_comparativas": build_comparison_metrics(resumo_a, resumo_b),
        },
    }

    destination = output_path(args, alertas_dir)
    destination.write_text(json.dumps(comparativo, ensure_ascii=False, indent=2), encoding="utf-8")

    print(f"Comparativo gerado em: {destination}")
    print(f"Periodo A: {comparativo['comparativo']['periodo_a']['label']} ({arquivo_a.name})")
    print(f"Periodo B: {comparativo['comparativo']['periodo_b']['label']} ({arquivo_b.name})")


if __name__ == "__main__":
    main()
