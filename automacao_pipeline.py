#!/usr/bin/env python
# -*- coding: utf-8 -*-

import argparse
import json
import subprocess
import sys
from datetime import datetime
from pathlib import Path


STEP_ORDER = ["step1", "step2"]


def parse_args():
    parser = argparse.ArgumentParser(description="Orquestra o pipeline analitico do funil.")
    parser.add_argument("--config", required=True, help="Caminho do arquivo de configuracao JSON.")
    parser.add_argument(
        "--resume-from",
        default="step1",
        choices=STEP_ORDER,
        help="Etapa inicial para retomada do pipeline.",
    )
    parser.add_argument(
        "--no-alerts",
        action="store_true",
        help="Mantido por compatibilidade com o runner principal. Nao executa nada adicional.",
    )
    parser.add_argument(
        "--skip-step2",
        action="store_true",
        help="Pula a geracao do arquivo de itens e perdas quando apenas o step1 for necessario.",
    )
    return parser.parse_args()


def load_config(path: Path) -> dict:
    return json.loads(path.read_text(encoding="utf-8-sig"))


def run_command(args: list[str]) -> None:
    print("Executando:", " ".join(args))
    result = subprocess.run(args, check=False)
    if result.returncode != 0:
        raise RuntimeError(f"Falha ao executar comando: {' '.join(args)}")


def latest_output(history_dir: Path, stem: str, after: datetime) -> Path:
    candidates = [
        p for p in history_dir.glob(f"{stem}_*.xlsx") if datetime.fromtimestamp(p.stat().st_mtime) >= after
    ]
    if not candidates:
        candidates = list(history_dir.glob(f"{stem}_*.xlsx"))
    if not candidates:
        raise FileNotFoundError(f"Nenhum arquivo encontrado para o padrao {stem}_*.xlsx em {history_dir}")
    return max(candidates, key=lambda p: p.stat().st_mtime)


def main():
    args = parse_args()
    config_path = Path(args.config).resolve()
    cfg = load_config(config_path)

    root = Path.cwd()
    history_dir = (root / cfg["paths"]["history_dir"]).resolve()
    history_dir.mkdir(parents=True, exist_ok=True)

    scripts = cfg["scripts"]
    inputs = cfg["inputs"]
    outputs = cfg["outputs"]
    filters = cfg.get("filters", {})
    cluster = cfg.get("cluster", {})
    rankings = cfg.get("rankings", {})

    python_cmd = [sys.executable]

    resume_index = STEP_ORDER.index(args.resume_from)
    oportunidades_path = None

    if resume_index <= STEP_ORDER.index("step1"):
        started = datetime.now()
        step1_script = str((root / scripts["step1"]).resolve())
        run_command(
            python_cmd
            + [
                step1_script,
                "-i",
                str((root / inputs["ruidos"]).resolve()),
                "-it",
                str((root / inputs["itens"]).resolve()),
                "-o",
                outputs["oportunidades_stem"],
                "--delta_horas",
                str(cluster.get("delta_horas", 360)),
                "--sim_min",
                str(cluster.get("sim_min", 0.5)),
            ]
            + (["--start", str(filters["start"])] if filters.get("start") else [])
            + (["--end", str(filters["end"])] if filters.get("end") else [])
            + (
                ["--debug_id", str(cluster["debug_id"])]
                if cluster.get("debug_id") is not None
                else []
            )
        )
        oportunidades_path = latest_output(history_dir, outputs["oportunidades_stem"], started)
    else:
        oportunidades_path = latest_output(history_dir, outputs["oportunidades_stem"], datetime.min)

    if (not args.skip_step2) and resume_index <= STEP_ORDER.index("step2"):
        started = datetime.now()
        step2_script = str((root / scripts["step2"]).resolve())
        run_command(
            python_cmd
            + [
                step2_script,
                "-i",
                str(oportunidades_path),
                "-it",
                str((root / inputs["itens"]).resolve()),
                "-o",
                outputs["itens_perdas_stem"],
                "--tipo_perda",
                str(filters.get("tipo_perda", "todas")),
                "--modo_data",
                str(filters.get("modo_data_itens", "criacao")),
                "--top",
                str(rankings.get("top", 50)),
                "--top_vendedor",
                str(rankings.get("top_vendedor", 10)),
            ]
            + (["--start", str(filters["start"])] if filters.get("start") else [])
            + (["--end", str(filters["end"])] if filters.get("end") else [])
        )
        _ = latest_output(history_dir, outputs["itens_perdas_stem"], started)
    elif args.skip_step2:
        print("Step2 ignorado por parametro --skip-step2.")

    print("Pipeline concluido com sucesso.")


if __name__ == "__main__":
    main()
