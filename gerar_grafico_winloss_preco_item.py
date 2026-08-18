#!/usr/bin/env python
# -*- coding: utf-8 -*-

r"""
Gera PNG de Win/Loss por faixa de preco para um item.

Entrada esperada: arquivo do step 2 com abas Itens_Faturados e Itens_Nao_Convertidas.

Exemplo:
py -3.12 gerar_grafico_winloss_preco_item.py ^
  -i historico\itens_perdas_reais_auto_20260508_142409.xlsx ^
  --codigo 5468 ^
  --tipo-cliente Revendedor
"""

import argparse
import math
import os
from datetime import datetime

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import matplotlib.patheffects as pe
from matplotlib.patches import FancyBboxPatch


def positive_float(value: str) -> float:
    number = float(str(value).replace(",", "."))
    if number <= 0:
        raise argparse.ArgumentTypeError("deve ser maior que zero")
    return number


def parse_args():
    p = argparse.ArgumentParser()
    p.add_argument("-i", "--input", required=True, help="Arquivo .xlsx do step 2.")
    p.add_argument("--codigo", required=True, help="CodigoErp do item.")
    p.add_argument("--tipo-cliente", default=None, help="Filtro opcional de Tipo Cliente.")
    p.add_argument("--descricao", default=None, help="Filtro opcional por trecho da descricao.")
    p.add_argument("--bin-size", type=positive_float, default=None, help="Tamanho da faixa de preco. Ex.: 5")
    p.add_argument("--min-total-faixa", type=int, default=1, help="Minimo de orcamentos para manter a faixa.")
    p.add_argument("-o", "--output", default=None, help="Arquivo PNG de saida.")
    return p.parse_args()


def read_sheet(path: str, sheet_name: str) -> pd.DataFrame:
    try:
        return pd.read_excel(path, sheet_name=sheet_name, dtype={"CodigoErp": "string", "IDOrcamentoPrinc": "string"})
    except ValueError as exc:
        raise ValueError(f"Aba obrigatoria nao encontrada: {sheet_name}") from exc


def normalize_item_rows(df: pd.DataFrame, codigo: str, tipo_cliente: str | None, descricao: str | None) -> pd.DataFrame:
    required = ["IDOrcamentoPrinc", "CodigoErp", "Descricao", "Quantidade", "VlTot", "Tipo Cliente"]
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError("Colunas obrigatorias ausentes: " + ", ".join(missing))

    out = df.copy()
    out["CodigoErp"] = out["CodigoErp"].astype("string").str.strip()
    out["Tipo Cliente"] = out["Tipo Cliente"].astype("string").fillna("").str.strip()
    out["Descricao"] = out["Descricao"].astype("string").fillna("").str.strip()
    out = out[out["CodigoErp"] == str(codigo).strip()]

    if tipo_cliente:
        out = out[out["Tipo Cliente"].str.casefold() == str(tipo_cliente).strip().casefold()]
    if descricao:
        out = out[out["Descricao"].str.contains(str(descricao), case=False, na=False, regex=False)]

    out["Quantidade"] = pd.to_numeric(out["Quantidade"], errors="coerce").fillna(0.0)
    out["VlTot"] = pd.to_numeric(out["VlTot"], errors="coerce").fillna(0.0)
    out = out[out["Quantidade"] > 0].copy()
    return out


def build_prices_by_order(df: pd.DataFrame, result_label: str) -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame(columns=["IDOrcamentoPrinc", "Quantidade", "VlTot", "Preco_Orcamento", "Resultado"])

    grouped = (
        df.groupby("IDOrcamentoPrinc", dropna=False)
        .agg(
            Quantidade=("Quantidade", "sum"),
            VlTot=("VlTot", "sum"),
            Descricao=("Descricao", "first"),
            Tipo_Cliente=("Tipo Cliente", "first"),
        )
        .reset_index()
    )
    grouped["Preco_Orcamento"] = np.where(
        grouped["Quantidade"] > 0,
        grouped["VlTot"] / grouped["Quantidade"],
        np.nan,
    )
    grouped = grouped.dropna(subset=["Preco_Orcamento"]).copy()
    grouped["Resultado"] = result_label
    return grouped


def choose_bin_size(prices: pd.Series) -> float:
    span = float(prices.max() - prices.min())
    if span <= 0:
        return max(1.0, round(float(prices.max()) * 0.05, 2))

    raw = span / 7
    magnitude = 10 ** math.floor(math.log10(raw))
    normalized = raw / magnitude
    if normalized <= 1:
        nice = 1
    elif normalized <= 2:
        nice = 2
    elif normalized <= 5:
        nice = 5
    else:
        nice = 10
    return nice * magnitude


def build_bins(prices: pd.Series, bin_size: float | None) -> tuple[list[float], list[str]]:
    if bin_size is None:
        bin_size = choose_bin_size(prices)

    min_price = float(prices.min())
    max_price = float(prices.max())
    start = math.floor(min_price / bin_size) * bin_size
    end = math.ceil(max_price / bin_size) * bin_size
    if end <= start:
        end = start + bin_size

    edges = list(np.arange(start, end + bin_size * 1.01, bin_size))
    labels = []
    for i in range(len(edges) - 1):
        left = edges[i]
        right = edges[i + 1]
        labels.append(f"{left:,.2f} - {right:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
    return edges, labels


def summarize_by_price_band(data: pd.DataFrame, bin_size: float | None, min_total: int) -> pd.DataFrame:
    edges, labels = build_bins(data["Preco_Orcamento"], bin_size)
    data = data.copy()
    data["Faixa_Preco"] = pd.cut(
        data["Preco_Orcamento"],
        bins=edges,
        labels=labels,
        include_lowest=True,
        right=False,
    )

    summary = (
        data.groupby("Faixa_Preco", observed=False)
        .agg(
            Faturados=("Resultado", lambda s: int((s == "Faturado").sum())),
            Nao_Convertidos=("Resultado", lambda s: int((s == "Nao Convertido").sum())),
            Preco_Min=("Preco_Orcamento", "min"),
            Preco_Max=("Preco_Orcamento", "max"),
        )
        .reset_index()
    )
    summary["Total"] = summary["Faturados"] + summary["Nao_Convertidos"]
    summary = summary[summary["Total"] >= int(min_total)].copy()
    summary["Win_Rate_%"] = np.where(summary["Total"] > 0, summary["Faturados"] / summary["Total"] * 100, np.nan)
    return summary


def compact_price_band_label(value: str) -> str:
    text = str(value)
    if " - " not in text:
        return text
    left, right = text.split(" - ", 1)

    def compact_number(part: str) -> str:
        part = part.strip()
        if part.endswith(",00"):
            return part[:-3]
        return part

    return f"{compact_number(left)}-{compact_number(right)}"


def split_subtitle(subtitle: str) -> tuple[str, str]:
    parts = [p.strip() for p in str(subtitle).split("|")]
    description = parts[0] if parts else ""
    tipo_cliente = ""
    for part in parts:
        if part.lower().startswith("tipo cliente:"):
            tipo_cliente = part.split(":", 1)[1].strip()
            break
    return description, tipo_cliente


def add_summary_card(fig, x: float, y: float, w: float, h: float, label: str, value: str, bg: str, color: str):
    card = FancyBboxPatch(
        (x, y),
        w,
        h,
        boxstyle="round,pad=0.008,rounding_size=0.008",
        linewidth=0,
        facecolor=bg,
        transform=fig.transFigure,
        clip_on=False,
    )
    fig.add_artist(card)
    fig.text(x + 0.014, y + h - 0.020, label, fontsize=8, color=color, ha="left", va="top")
    fig.text(x + 0.014, y + 0.018, value, fontsize=16, fontweight="bold", color=color, ha="left", va="bottom")


def plot_chart(summary: pd.DataFrame, title: str, subtitle: str, output_path: str):
    if summary.empty:
        raise ValueError("Nao ha dados suficientes para gerar o grafico.")

    x = np.arange(len(summary))
    wins = summary["Faturados"].to_numpy()
    losses = summary["Nao_Convertidos"].to_numpy()
    win_rate = summary["Win_Rate_%"].to_numpy()

    n_bands = len(summary)
    fig_width = max(12, min(20, 8 + n_bands * 0.48))
    fig_height = 8.5 if n_bands <= 12 else 9.3
    fig, ax = plt.subplots(figsize=(fig_width, fig_height))
    fig.patch.set_facecolor("white")
    ax.set_facecolor("white")

    green = "#4CCF78"
    green_edge = "#34A85E"
    red = "#F7486A"
    red_edge = "#D72F50"
    navy = "#185FA5"
    grid = "#E7EAF0"
    dark = "#111827"
    muted = "#6B7280"

    bar_width = 0.58 if n_bands > 14 else 0.64
    bars_win = ax.bar(x, wins, width=bar_width, color=green, edgecolor=green_edge, label="Faturados (Win)")
    bars_loss = ax.bar(x, losses, width=bar_width, bottom=wins, color=red, edgecolor=red_edge, label="Nao Convertidos (Loss)")

    ax2 = ax.twinx()
    ax2.plot(x, win_rate, color=navy, marker="o", markersize=5, linewidth=2.0, label="Taxa de Aprovação (Win Rate)")

    description, tipo_cliente = split_subtitle(subtitle)
    clean_title = title.replace("WIN/LOSS POR PREÇO - ", "").replace("ITEM ", "Item ")
    total_win = int(np.nansum(wins))
    total_loss = int(np.nansum(losses))
    total = total_win + total_loss

    fig.text(0.055, 0.965, "WIN/LOSS POR PREÇO", fontsize=8, fontweight="bold", color=muted, ha="left", va="top")
    fig.text(0.055, 0.928, f"{clean_title} - {description}", fontsize=15, fontweight="bold", color=dark, ha="left", va="top")
    if tipo_cliente:
        fig.text(0.055, 0.892, f"Tipo Cliente: {tipo_cliente}", fontsize=9, color=muted, ha="left", va="top")

    card_y = 0.770
    card_w = 0.165
    card_h = 0.078
    add_summary_card(fig, 0.055, card_y, card_w, card_h, "Total de Orçamentos", f"{total}", "#F3F4F6", dark)
    add_summary_card(fig, 0.285, card_y, card_w, card_h, "Faturados (Win)", f"{total_win}", "#E9F9F0", "#167A3A")
    add_summary_card(fig, 0.515, card_y, card_w, card_h, "Não Convertidos (Loss)", f"{total_loss}", "#FDECF1", "#B01032")

    ax.set_ylabel("Qtd Orçamentos", fontsize=9, color=muted)
    ax2.set_ylabel("Taxa de Aprovação", fontsize=9, color=navy)
    ax.set_xlabel("Faixa de Preço (R$)", fontsize=9, color=muted, labelpad=12)

    ax.set_xticks(x)
    x_font = 8 if n_bands > 14 else 9
    x_rotation = 40 if n_bands > 10 else 0
    labels = [compact_price_band_label(v) for v in summary["Faixa_Preco"].astype(str).tolist()]
    ax.set_xticklabels(labels, rotation=x_rotation, ha="right" if x_rotation else "center", fontsize=x_font, color=muted)
    ax.grid(axis="y", color=grid, linewidth=0.8)
    ax.set_axisbelow(True)
    ax2.set_ylim(0, 112)
    ax2.set_yticks(np.arange(0, 101, 20))
    ax2.set_yticklabels([f"{v}%" for v in range(0, 101, 20)])

    max_total = max((wins + losses).max(), 1)
    label_font = 10 if n_bands > 14 else 11
    min_inside_height = max_total * 0.060
    text_outline = [pe.withStroke(linewidth=2, foreground="white")]
    for bar, value in zip(bars_win, wins):
        if value <= 0:
            continue
        x_pos = bar.get_x() + bar.get_width() / 2
        if value >= min_inside_height:
            txt = ax.text(
                x_pos,
                value / 2,
                f"{int(value)}",
                ha="center",
                va="center",
                color="white",
                fontsize=label_font,
                fontweight="bold",
                zorder=8,
            )
        else:
            txt = ax.text(
                x_pos,
                value + max_total * 0.018,
                f"{int(value)}",
                ha="center",
                va="bottom",
                color="#167A3A",
                fontsize=label_font,
                fontweight="bold",
                zorder=8,
            )
            txt.set_path_effects(text_outline)
    for bar, win_value, loss_value in zip(bars_loss, wins, losses):
        if loss_value <= 0:
            continue
        x_pos = bar.get_x() + bar.get_width() / 2
        y_mid = win_value + loss_value / 2
        if loss_value >= min_inside_height:
            txt = ax.text(
                x_pos,
                y_mid,
                f"{int(loss_value)}",
                ha="center",
                va="center",
                color="white",
                fontsize=label_font,
                fontweight="bold",
                zorder=8,
            )
        else:
            txt = ax.text(
                x_pos,
                win_value + loss_value + max_total * 0.018,
                f"{int(loss_value)}",
                ha="center",
                va="bottom",
                color="#B01032",
                fontsize=label_font,
                fontweight="bold",
                zorder=8,
            )
            txt.set_path_effects(text_outline)
    if n_bands <= 12:
        for idx, (xi, rate, total_band) in enumerate(zip(x, win_rate, wins + losses)):
            if np.isnan(rate):
                continue
            if rate >= 96:
                y_label = 94
                va = "top"
            elif rate <= 4:
                y_label = 4
                va = "bottom"
            else:
                y_label = min(rate + 4, 108)
                va = "bottom"
            ax2.text(
                xi,
                y_label,
                f"{rate:.0f}%",
                ha="center",
                va=va,
                color=navy,
                fontsize=8,
                fontweight="bold",
                zorder=9,
                path_effects=[pe.withStroke(linewidth=2, foreground="white")],
            )

    ax.set_ylim(0, max_total * 1.32)
    handles1, labels1 = ax.get_legend_handles_labels()
    handles2, labels2 = ax2.get_legend_handles_labels()
    fig.legend(
        handles1 + handles2,
        labels1 + labels2,
        loc="upper left",
        bbox_to_anchor=(0.055, 0.705),
        ncol=3,
        frameon=False,
        fontsize=8,
    )

    for spine in ["top", "right"]:
        ax.spines[spine].set_visible(False)
    ax2.spines["top"].set_visible(False)
    ax.spines["left"].set_color("#CBD5E1")
    ax.spines["bottom"].set_color("#CBD5E1")
    ax2.spines["right"].set_color("#CBD5E1")
    ax.tick_params(axis="y", labelsize=8, colors=muted)
    ax2.tick_params(axis="y", labelsize=8, colors=navy)

    fig.subplots_adjust(left=0.075, right=0.92, top=0.660, bottom=0.145)
    fig.savefig(output_path, dpi=180, bbox_inches="tight")
    plt.close(fig)


def main():
    args = parse_args()
    faturados = normalize_item_rows(read_sheet(args.input, "Itens_Faturados"), args.codigo, args.tipo_cliente, args.descricao)
    nao_convertidos = normalize_item_rows(read_sheet(args.input, "Itens_Nao_Convertidas"), args.codigo, args.tipo_cliente, args.descricao)

    wins = build_prices_by_order(faturados, "Faturado")
    losses = build_prices_by_order(nao_convertidos, "Nao Convertido")
    data = pd.concat([wins, losses], ignore_index=True)

    if data.empty:
        raise ValueError("Nenhum dado encontrado para o item/filtros informados.")

    descricao = data["Descricao"].dropna().astype(str).iloc[0] if "Descricao" in data and not data["Descricao"].dropna().empty else ""
    tipo_cliente = args.tipo_cliente or (data["Tipo_Cliente"].dropna().astype(str).iloc[0] if "Tipo_Cliente" in data and not data["Tipo_Cliente"].dropna().empty else "Todos")
    summary = summarize_by_price_band(data, args.bin_size, args.min_total_faixa)

    if args.output:
        output_path = args.output
    else:
        os.makedirs("graficos", exist_ok=True)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        safe_tipo = re_safe(tipo_cliente)
        output_path = os.path.join("graficos", f"winloss_preco_item_{args.codigo}_{safe_tipo}_{timestamp}.png")

    title = f"WIN/LOSS POR PREÇO - ITEM {args.codigo}"
    subtitle = f"{descricao} | Tipo Cliente: {tipo_cliente} | Win: {len(wins)} | Loss: {len(losses)}"
    plot_chart(summary, title, subtitle, output_path)
    print(f"Grafico gerado: {os.path.abspath(output_path)}")


def re_safe(value: str) -> str:
    safe = "".join(ch if ch.isalnum() else "_" for ch in str(value).strip())
    return "_".join(part for part in safe.split("_") if part) or "todos"


if __name__ == "__main__":
    main()
