"""Офлайн-сравнение моделей: python -m recsys.evaluate

1. train / test - глобальный сплит по времени (последние 5% покупок в тест).
2. Веса гибрида подбираются на валидации: train ещё раз делится по времени
   (последние 5% train), модели обучаются на оставшейся части. Тест в подборе не участвует.
3. Все модели обучаются на полном train и сравниваются на тесте.
"""
import itertools
import json
import time

import numpy as np
import pandas as pd

from recsys.data import ROOT, prepare, time_split
from recsys.metrics import beyond_accuracy, ranking_metrics
from recsys.models import (ALS, AlsoBought, ContentBased, Context, Hybrid, ItemKNN, Popular, SameBrand, Trending,
                           UserKNN, rank_contrib, recommend)

K = 10
WEIGHTS_PATH = ROOT / "reports" / "hybrid_weights.json"
# сетка весов: ALS - опорная модель с весом 1, остальные относительно неё
GRID = {"ItemKNN": [0.5, 1.0], "ContentBased": [0.0, 0.5, 1.0], "Trending": [0.0, 0.5, 1.0, 2.0, 3.0]}


def relevant_items(train: pd.DataFrame, test: pd.DataFrame, ctx: Context):
    """Релевантное = купленное в тестовом периоде, кроме уже купленного в train. Возвращает (warm, cold)."""
    test = test.merge(train[["user_id", "item_id"]], how="left", indicator=True)
    test = test[test["_merge"] == "left_only"]
    rel = test.groupby("user_id")["item_id"].apply(lambda s: {ctx.item2idx[i] for i in s})
    return rel[rel.index.isin(ctx.user2idx)], rel[~rel.index.isin(ctx.user2idx)]


def tune_hybrid(train: pd.DataFrame, items: pd.DataFrame) -> dict:
    """Перебор весов гибрида по NDCG@K на валидационном срезе.

    Ранжирование каждой компоненты считается один раз, дальше для каждой комбинации
    весов меняется только сумма вкладов RRF - перебор занимает секунды.
    """
    inner, val, val_cutoff = time_split(train)
    ctx = Context(inner, items)
    warm, _ = relevant_items(inner, val, ctx)
    uidx = warm.index.map(ctx.user2idx).to_numpy()
    models = {"ALS": ALS(), "ItemKNN": ItemKNN(), "ContentBased": ContentBased(), "Trending": Trending()}
    contrib = {name: rank_contrib(m.fit(ctx), ctx, uidx) for name, m in models.items()}
    print(f"validation: cutoff={val_cutoff:%Y-%m-%d}, warm users={len(uidx):,}")

    results = []
    for combo in itertools.product(*GRID.values()):
        w = {"ALS": 1.0, **dict(zip(GRID, combo))}
        S = sum(wi * contrib[name] for name, wi in w.items() if wi)
        recs = []
        for s in range(0, len(uidx), 1024):
            d = S[s:s + 1024].toarray()
            top = np.argpartition(-d, K, axis=1)[:, :K]
            order = np.argsort(-np.take_along_axis(d, top, axis=1), axis=1)
            recs.append(np.take_along_axis(top, order, axis=1))
        ndcg = ranking_metrics(np.vstack(recs), list(warm), K)[f"NDCG@{K}"]
        results.append((ndcg, w))

    best = max(results, key=lambda r: r[0])
    best_no_trend = max((r for r in results if r[1]["Trending"] == 0), key=lambda r: r[0])
    for label, (ndcg, w) in [("with Trending", best), ("without Trending", best_no_trend)]:
        print(f"best hybrid {label}: val NDCG@{K}={ndcg:.4f} weights={w}")
    return {"hybrid": best[1], "hybrid_no_trending": best_no_trend[1], "val_cutoff": f"{val_cutoff:%Y-%m-%d}",
            "val_ndcg": {"hybrid": best[0], "hybrid_no_trending": best_no_trend[0]}}


def dataset_stats(inter, items, train, test, ctx, warm, cold) -> dict:
    """Описательная статистика датасета для отчёта (reports/data_stats.json)."""
    per_user = inter.groupby("user_id").size()
    per_item = inter["item_id"].value_counts()
    cum = per_item.cumsum() / per_item.sum()
    months = inter["ts"].dt.month.value_counts(normalize=True)
    no_desc = items["description"].str.strip() == ""
    return {
        "interactions": len(inter), "users": int(inter["user_id"].nunique()), "items": len(ctx.item_ids),
        "density": ctx.X.nnz / np.prod(ctx.X.shape),
        "period": [f"{inter.ts.min():%Y}", f"{inter.ts.max():%Y}"],
        "median_per_user": float(per_user.median()), "users_le5": float((per_user <= 5).mean()),
        "items_le10": float((per_item <= 10).mean()),
        "top1_share": float(cum.iloc[int(len(cum) * 0.01) - 1]), "top20_share": float(cum.iloc[int(len(cum) * 0.2) - 1]),
        "low_rating": float((inter["rating"] <= 2).mean()),
        "dec_jan": float(months.get(12, 0) + months.get(1, 0)),
        "no_description": float(no_desc.mean()),
        "no_description_features": float((no_desc & (items["features"].str.strip() == "")).mean()),
        "train": len(train), "test_start": f"{test.ts.min():%d.%m.%Y}", "test_end": f"{test.ts.max():%d.%m.%Y}",
        "test_months": round((test.ts.max() - test.ts.min()).days / 30.4),
        "warm_users": len(warm), "cold_users": len(cold), "cold_share": len(cold) / (len(warm) + len(cold)),
    }


def build_models(weights: dict):
    base = {"ALS": ALS(), "ItemKNN": ItemKNN(), "ContentBased": ContentBased(), "Trending": Trending()}

    def hybrid(w, name):
        return Hybrid({base[n]: wi for n, wi in w.items() if wi}, name=name)

    return [Popular(), base["Trending"], AlsoBought(), SameBrand(), base["ContentBased"], base["ItemKNN"], UserKNN(),
            base["ALS"], hybrid(weights["hybrid_no_trending"], "Hybrid без Trending"),
            hybrid(weights["hybrid"], "Hybrid")]


def main():
    inter, items = prepare()
    train, test, cutoff = time_split(inter)
    # time_split уже убрал из теста новые товары, их долю считаем по полному срезу
    new_item_share = float((~inter.loc[inter["ts"] >= cutoff, "item_id"].isin(set(train["item_id"]))).mean())

    weights = tune_hybrid(train, items)
    WEIGHTS_PATH.parent.mkdir(exist_ok=True)
    WEIGHTS_PATH.write_text(json.dumps(weights, indent=2, ensure_ascii=False))

    ctx = Context(train, items)
    warm, cold = relevant_items(train, test, ctx)
    warm_uidx = warm.index.map(ctx.user2idx).to_numpy()

    print(f"interactions={len(inter):,} users={inter.user_id.nunique():,} items={len(ctx.item_ids):,}")
    print(f"test: {cutoff:%Y-%m-%d} - {test.ts.max():%Y-%m-%d}  train={len(train):,}  "
          f"test users: warm={len(warm):,} cold={len(cold):,}  "
          f"density={ctx.X.nnz / np.prod(ctx.X.shape):.5%}")

    stats = dataset_stats(inter, items, train, test, ctx, warm, cold)
    stats["test_new_items"] = new_item_share
    (ROOT / "reports" / "data_stats.json").write_text(json.dumps(stats, indent=2, ensure_ascii=False))

    rows = []
    top_trending = np.argsort(-Trending().fit(ctx).scores)[:30]
    for m in build_models(weights):
        t0 = time.perf_counter()
        m.fit(ctx)  # компоненты гибрида стоят в списке раньше него и уже обучены
        fit_s = time.perf_counter() - t0
        t0 = time.perf_counter()
        recs = recommend(m, ctx, warm_uidx, K)
        ms_per_user = (time.perf_counter() - t0) / len(warm_uidx) * 1000
        rows.append({"model": m.name, "segment": "warm",
                     **ranking_metrics(recs, list(warm), K),
                     **beyond_accuracy(recs, ctx.item_pop, len(ctx.user_ids), top_trending),
                     "fit_s": fit_s, "ms/user": ms_per_user})
        print(f"{m.name}: NDCG@{K}={rows[-1][f'NDCG@{K}']:.4f}")

    # холодные пользователи: истории нет, работают только не-персональные модели
    for m in [Popular().fit(ctx), Trending().fit(ctx)]:
        top = np.argsort(-m.scores)[:K]
        recs = np.tile(top, (len(cold), 1))
        rows.append({"model": m.name, "segment": "cold", **ranking_metrics(recs, list(cold), K),
                     **beyond_accuracy(recs, ctx.item_pop, len(ctx.user_ids), top_trending)})

    res = pd.DataFrame(rows)
    out = ROOT / "reports"
    res.to_csv(out / "metrics.csv", index=False)
    (out / "metrics.md").write_text(res.to_markdown(index=False, floatfmt=".4f"))
    print(res.to_string(index=False, float_format="%.4f"))


if __name__ == "__main__":
    main()
