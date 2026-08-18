from __future__ import annotations

import argparse
import json
from pathlib import Path


BASE_DIR = Path(__file__).resolve().parent
ALERTAS_DIR = BASE_DIR / "alertas"
DEFAULT_OUTPUT = BASE_DIR / "Slide de Win Rate Interativo" / "data" / "comparativo_latest.js"


def find_latest_json() -> Path:
    arquivos = sorted(
        ALERTAS_DIR.glob("comparativo_win_rate_periodos_*.json"),
        key=lambda p: p.stat().st_mtime,
        reverse=True,
    )
    if not arquivos:
        raise FileNotFoundError("Nenhum comparativo JSON encontrado em alertas.")
    return arquivos[0]


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Gera o arquivo JS usado pelo slide interativo de win rate."
    )
    parser.add_argument(
        "-i",
        "--input",
        type=Path,
        help="Caminho do JSON comparativo. Se omitido, usa o mais recente da pasta alertas.",
    )
    parser.add_argument(
        "-o",
        "--output",
        type=Path,
        default=DEFAULT_OUTPUT,
        help="Caminho do arquivo JS de saída.",
    )
    args = parser.parse_args()

    input_path = args.input if args.input else find_latest_json()
    output_path = args.output

    with input_path.open("r", encoding="utf-8") as f:
        data = json.load(f)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    payload = json.dumps(data, ensure_ascii=False, indent=2)
    content = "window.WIN_RATE_COMPARATIVO_DATA = " + payload + ";\n"

    with output_path.open("w", encoding="utf-8") as f:
        f.write(content)

    print(f"Arquivo de dados gerado: {output_path}")
    print(f"Fonte usada: {input_path}")


if __name__ == "__main__":
    main()
