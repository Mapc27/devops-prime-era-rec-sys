"""REST-сервис рекомендаций: uvicorn recsys.api:app

Модели обучаются на всех данных при старте (~1 минута).
В проде вместо этого рекомендации считают офлайн и кладут в кэш (см. отчёт, часть II-b).
"""
import json
import time
from contextlib import asynccontextmanager

import numpy as np
from fastapi import FastAPI, HTTPException, Query

from recsys.data import ROOT, prepare
from recsys.models import ALS, AlsoBought, ContentBased, Context, Hybrid, ItemKNN, SameBrand, Trending, recommend

S: dict = {}
# веса гибрида, подобранные на валидации (python -m recsys.evaluate); без файла - значения по умолчанию
WEIGHTS_PATH = ROOT / "reports" / "hybrid_weights.json"
DEFAULT_WEIGHTS = {"ALS": 1.0, "ItemKNN": 1.0, "ContentBased": 0.5, "Trending": 2.0}


@asynccontextmanager
async def lifespan(app: FastAPI):
    inter, items = prepare()
    ctx = Context(inter, items)
    als, iknn, content = ALS().fit(ctx), ItemKNN().fit(ctx), ContentBased().fit(ctx)
    trending = Trending().fit(ctx)
    models = {"ALS": als, "ItemKNN": iknn, "ContentBased": content, "Trending": trending}
    weights = json.loads(WEIGHTS_PATH.read_text())["hybrid"] if WEIGHTS_PATH.exists() else DEFAULT_WEIGHTS
    S.update(ctx=ctx, trending=trending, also=AlsoBought().fit(ctx), brand=SameBrand().fit(ctx), content=content,
             hybrid=Hybrid({models[n]: w for n, w in weights.items() if w}).fit(ctx))
    yield


app = FastAPI(title="Marketplace RecSys MVP", lifespan=lifespan)
K = Query(10, ge=1, le=100, description="Сколько товаров вернуть")


def _items(idx) -> list[dict]:
    it = S["ctx"].items.iloc[list(idx)]
    return it[["item_id", "title", "brand", "category", "price"]].replace({np.nan: None}).to_dict("records")


def _item_idx(item_id: str) -> int:
    if item_id not in S["ctx"].item2idx:
        raise HTTPException(404, "unknown item")
    return S["ctx"].item2idx[item_id]


@app.get("/recommend/{user_id}")
def recommend_for_user(user_id: str, k: int = K):
    """Блок «Вам может понравиться». Новый пользователь получает Trending (switching-гибрид)."""
    ctx, t0 = S["ctx"], time.perf_counter()
    if user_id in ctx.user2idx:
        source, idx = "hybrid", recommend(S["hybrid"], ctx, np.array([ctx.user2idx[user_id]]), k)[0]
    else:
        source, idx = "trending_fallback", np.argsort(-S["trending"].scores)[:k]
    return {"source": source, "latency_ms": round((time.perf_counter() - t0) * 1000, 1), "items": _items(idx)}


@app.get("/history/{user_id}")
def recently_bought(user_id: str, k: int = K):
    """Блок «Недавно покупали» (просмотров в датасете нет, поэтому только покупки)."""
    ctx = S["ctx"]
    if user_id not in ctx.user2idx:
        return {"items": []}
    return {"items": _items(ctx.history[ctx.user2idx[user_id]][::-1][:k])}


@app.get("/item/{item_id}/also_bought")
def also_bought(item_id: str, k: int = K):
    """Блок «С этим товаром покупают»."""
    return {"items": _items(S["also"].similar(_item_idx(item_id), k))}


@app.get("/item/{item_id}/similar")
def similar(item_id: str, k: int = K):
    """Блок «Похожие товары» (content-based: похожесть по описанию, а не по покупкам)."""
    return {"items": _items(S["content"].similar(_item_idx(item_id), k))}


@app.get("/item/{item_id}/same_brand")
def same_brand(item_id: str, k: int = K):
    """Блок «Другие товары бренда»."""
    return {"items": _items(S["brand"].similar(_item_idx(item_id), k))}


@app.get("/popular")
def popular(k: int = K):
    """Блок «Популярное сейчас»."""
    return {"items": _items(np.argsort(-S["trending"].scores)[:k])}


@app.get("/health")
def health():
    return {"status": "ok" if S else "loading"}
