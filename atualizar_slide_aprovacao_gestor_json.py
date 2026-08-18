from __future__ import annotations

import argparse
import json
from pathlib import Path


BASE_DIR = Path(__file__).resolve().parent
ALERTAS_DIR = BASE_DIR / "alertas"
DEFAULT_OUTPUT = BASE_DIR / "Slide de Win Rate Interativo" / "data" / "aprovacao_gestor_latest.js"


def find_latest_json() -> Path:
    arquivos = sorted(
        ALERTAS_DIR.glob("analise_aprovacao_gestor_*.json"),
        key=lambda p: p.stat().st_mtime,
        reverse=True,
    )
    if not arquivos:
        raise FileNotFoundError("Nenhum JSON de aprovacao do gestor encontrado em alertas.")
    return arquivos[0]


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Gera o arquivo JS usado pelo slide interativo de aprovacao do gestor."
    )
    parser.add_argument(
        "-i",
        "--input",
        type=Path,
        help="Caminho do JSON de aprovacao do gestor. Se omitido, usa o mais recente.",
    )
    parser.add_argument(
        "-o",
        "--output",
        type=Path,
        default=DEFAULT_OUTPUT,
        help="Caminho do arquivo JS de saida.",
    )
    parser.add_argument(
        "--visao",
        choices=("sem_ruido", "com_ruido"),
        default="sem_ruido",
        help="Visao padrao do slide. Default: sem_ruido.",
    )
    args = parser.parse_args()

    input_path = args.input if args.input else find_latest_json()
    output_path = args.output

    with input_path.open("r", encoding="utf-8") as f:
        data = json.load(f)

    data["visao_padrao"] = args.visao

    output_path.parent.mkdir(parents=True, exist_ok=True)
    payload = json.dumps(data, ensure_ascii=False, indent=2)
    content = "window.APROVACAO_GESTOR_DATA = " + payload + ";\n"

    with output_path.open("w", encoding="utf-8") as f:
        f.write(content)

    print(f"Arquivo de dados gerado: {output_path}")
    print(f"Fonte usada: {input_path}")
    print(f"Visao padrao: {args.visao}")


if __name__ == "__main__":
    main()
