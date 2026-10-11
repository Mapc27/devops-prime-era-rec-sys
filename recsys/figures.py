"""Графики для отчёта: python -m recsys.figures (после recsys.evaluate).

reports/figures/data.png         - длинный хвост и сезонность
reports/figures/metrics.png      - точность, покрытие и доля Trending по моделям
reports/figures/architecture.png - схема решения
"""
import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
from matplotlib.patches import FancyBboxPatch

from recsys.data import ROOT, prepare

OUT = ROOT / "reports" / "figures"
BLUE, ORANGE, GREEN, GREY = "#2b7bd6", "#ec6a35", "#1bb07a", "#4d4d4d"
plt.rcParams.update({"font.family": "DejaVu Sans", "font.size": 12, "axes.spines.top": False,
                     "axes.spines.right": False, "axes.grid": True, "grid.color": "#e5e5e5",
                     "axes.edgecolor": "#cccccc", "axes.labelcolor": GREY, "xtick.color": GREY,
                     "ytick.color": GREY, "axes.titlelocation": "left", "axes.axisbelow": True})


def data_figure(inter: pd.DataFrame):
    fig, (a, b) = plt.subplots(1, 2, figsize=(14, 5))
    cnt = inter["item_id"].value_counts().to_numpy()
    x = np.arange(1, len(cnt) + 1) / len(cnt) * 100
    y = np.cumsum(cnt) / cnt.sum() * 100
    a.plot(x, y, color=BLUE, lw=3)
    i20 = int(len(cnt) * 0.2) - 1
    a.scatter([x[i20]], [y[i20]], color=BLUE, zorder=3, edgecolor="white", s=60)
    a.annotate(f"20% товаров = {y[i20]:.0f}% покупок", (x[i20], y[i20]), (32, 55),
               arrowprops=dict(arrowstyle="-", color=GREY), color=GREY)
    a.set(xlim=(0, 100), ylim=(0, 102), xlabel="Доля товаров (от самых популярных), %", ylabel="Доля покупок, %",
          title="Длинный хвост: накопленная доля покупок")

    m = inter["ts"].dt.month.value_counts(normalize=True).sort_index() * 100
    b.bar(m.index, m.to_numpy(), color=BLUE, width=0.62)
    b.axhline(100 / 12, ls="--", color=GREY, lw=1)
    b.text(6.5, 100 / 12 + 0.25, "равномерно", ha="center", color=GREY)
    b.set_xticks(range(1, 13), ["Янв", "Фев", "Мар", "Апр", "Май", "Июн", "Июл", "Авг", "Сен", "Окт", "Ноя", "Дек"])
    b.grid(axis="x", visible=False)
    b.set_title("Сезонность: доля покупок по месяцам, %")
    fig.tight_layout()
    fig.savefig(OUT / "data.png", dpi=150)


def metrics_figure():
    df = pd.read_csv(ROOT / "reports" / "metrics.csv")
    df = df[df["segment"] == "warm"].iloc[::-1]
    heuristics = {"Popular", "Trending", "AlsoBought", "SameBrand"}
    colors = [GREEN if n.startswith("Hybrid") else BLUE if n in heuristics else ORANGE for n in df["model"]]
    fig, axes = plt.subplots(1, 3, figsize=(17, 6), sharey=True)
    panels = [("NDCG@10", "Точность: NDCG@10", 1, "{:.3f}"),
              ("Coverage", "Покрытие каталога, %", 100, "{:.1f}"),
              ("TrendShare", "Доля выдачи из топ-30 Trending, %", 100, "{:.0f}")]
    for ax, (col, title, scale, fmt) in zip(axes, panels):
        v = df[col].to_numpy() * scale
        ax.barh(df["model"], v, color=colors, height=0.62)
        for yi, vi in enumerate(v):
            ax.text(vi + v.max() * 0.015, yi, fmt.format(vi), va="center", color=GREY, fontsize=11)
        ax.set_xlim(0, v.max() * 1.18)
        ax.grid(axis="y", visible=False)
        ax.set_title(title)
    handles = [plt.Rectangle((0, 0), 1, 1, color=c) for c in (BLUE, ORANGE, GREEN)]
    fig.legend(handles, ["Эвристика", "ML", "Гибрид"], loc="lower center", ncol=3, frameon=False)
    fig.tight_layout(rect=(0, 0.06, 1, 1))
    fig.savefig(OUT / "metrics.png", dpi=150)


def architecture_figure():
    fig, ax = plt.subplots(figsize=(16, 8.6))
    ax.set(xlim=(0, 16), ylim=(0, 8.6))
    ax.axis("off")

    def box(x, y, w, h, title, body, accent=False):
        ax.add_patch(FancyBboxPatch((x, y), w, h, boxstyle="round,pad=0,rounding_size=0.15", lw=1.5,
                                    ec="#c8c8c0", fc="#e8f0fb" if accent else "#f3f2ee"))
        ax.text(x + w / 2, y + h - 0.5, title, ha="center", va="center", fontsize=15, weight="bold")
        ax.text(x + w / 2, y + h - 0.85, body, ha="center", va="top", fontsize=12.5, color=GREY, linespacing=1.6)

    def arrow(x1, y1, x2, y2, label=None):
        ax.annotate("", (x2, y2), (x1, y1), arrowprops=dict(arrowstyle="-|>", color=GREY, lw=2, mutation_scale=18))
        if label:
            ax.text((x1 + x2) / 2, (y1 + y2) / 2 + 0.2, label, ha="center", color=GREY, fontsize=12.5)

    box(0.3, 5.7, 3.1, 2.35, "HuggingFace", "Amazon Reviews 2023\nVideo_Games.csv\nmeta_Video_Games.jsonl")
    box(4.0, 5.7, 3.4, 2.35, "data.py", "очистка, фильтр\nметаданных → parquet\nсплит по времени (95/5)")
    box(8.05, 5.7, 3.4, 2.35, "Context", "индексы user/item\nразреженная матрица\nuser × item, история")
    box(12.1, 5.7, 3.55, 2.35, "models.py", "эвристики · content-based\nUserKNN · ItemKNN · ALS\nгибрид (switching + RRF)",
        accent=True)
    box(4.0, 1.4, 4.0, 2.8, "Клиент (витрина)",
        "«Вам может понравиться»\n«С этим покупают»\n«Похожие» · «Популярное»\n«Недавно покупали»")
    box(8.65, 1.4, 3.4, 2.8, "api.py (FastAPI)",
        "/recommend/{user}\n/item/{id}/also_bought\n/item/{id}/similar\n/same_brand · /popular\n/history/{user}",
        accent=True)
    box(12.85, 1.4, 2.9, 2.8, "evaluate.py",
        "подбор весов гибрида\nна валидации\nметрики @10 на тесте\n→ reports/metrics.md")
    arrow(3.45, 6.85, 3.9, 6.85)
    arrow(7.5, 6.85, 7.95, 6.85)
    arrow(11.55, 6.85, 12.0, 6.85)
    arrow(13.0, 5.6, 10.8, 4.3)
    arrow(14.3, 5.6, 14.3, 4.3)
    arrow(12.8, 2.8, 12.1, 2.8, "веса")
    arrow(8.55, 2.8, 8.1, 2.8, "JSON")
    ax.text(10.35, 0.9, "новый пользователь → fallback на Trending", ha="center", color=GREY, style="italic",
            fontsize=12.5)
    fig.savefig(OUT / "architecture.png", dpi=100, bbox_inches="tight")


def main():
    OUT.mkdir(parents=True, exist_ok=True)
    inter, _ = prepare()
    data_figure(inter)
    metrics_figure()
    architecture_figure()
    print(f"saved to {OUT}")


if __name__ == "__main__":
    main()
