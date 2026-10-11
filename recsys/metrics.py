"""Офлайн-метрики ранжирования и beyond-accuracy метрики."""
import numpy as np


def ranking_metrics(recs: np.ndarray, relevant: list[set], k: int = 10) -> dict:
    prec, rec, ndcg, ap, hit = [], [], [], [], []
    discounts = 1 / np.log2(np.arange(2, k + 2))
    for r, rel in zip(recs[:, :k], relevant):
        hits = np.array([i in rel for i in r], dtype=float)
        n_hit = hits.sum()
        prec.append(n_hit / k)
        rec.append(n_hit / len(rel))
        hit.append(float(n_hit > 0))
        ndcg.append((hits * discounts).sum() / discounts[:min(len(rel), k)].sum())
        ap.append((np.cumsum(hits) / np.arange(1, k + 1) * hits).sum() / min(len(rel), k))
    return {f"Precision@{k}": np.mean(prec), f"Recall@{k}": np.mean(rec), f"NDCG@{k}": np.mean(ndcg),
            f"MAP@{k}": np.mean(ap), f"HitRate@{k}": np.mean(hit)}


def beyond_accuracy(recs: np.ndarray, item_pop: np.ndarray, n_users: int, top_trending=None) -> dict:
    """Coverage - доля каталога, попавшая в рекомендации.
    Novelty - средняя «неожиданность» товара: -log2(доля купивших). Чем выше, тем менее заезжено.
    TrendShare - доля выдачи из топ-30 Trending: насколько «персональная» выдача совпадает с общей для всех.
    """
    p = np.maximum(item_pop, 1) / n_users
    out = {"Coverage": len(np.unique(recs)) / len(item_pop),
           "Novelty": float(np.mean(-np.log2(p[recs])))}
    if top_trending is not None:
        out["TrendShare"] = float(np.isin(recs, top_trending).mean())
    return out
