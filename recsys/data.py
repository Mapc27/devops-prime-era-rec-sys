"""Загрузка и подготовка датасета Amazon Reviews 2023 (Video Games, 5-core).

Отзыв трактуем как факт покупки (implicit feedback): оценка не важна,
важно, что пользователь купил товар.
"""
import json
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
RAW = ROOT / "data" / "raw"
PROCESSED = ROOT / "data" / "processed"


def load_interactions() -> pd.DataFrame:
    df = pd.read_csv(RAW / "Video_Games.csv")
    df = df.rename(columns={"parent_asin": "item_id"})
    df["ts"] = pd.to_datetime(df["timestamp"], unit="ms")
    return df[["user_id", "item_id", "rating", "ts"]]


def _clean_text(value) -> str:
    if isinstance(value, list):
        return " ".join(str(v) for v in value)
    return str(value) if value else ""


def load_items(item_ids: set[str]) -> pd.DataFrame:
    """Читает метаданные построчно и оставляет только товары из взаимодействий."""
    rows = []
    with open(RAW / "meta_Video_Games.jsonl") as f:
        for line in f:
            m = json.loads(line)
            if m["parent_asin"] not in item_ids:
                continue
            categories = m.get("categories") or []
            rows.append({
                "item_id": m["parent_asin"],
                "title": m.get("title") or "",
                "brand": (m.get("store") or "unknown").strip(),
                "category": categories[1] if len(categories) > 1 else (m.get("main_category") or "unknown"),
                "categories": " > ".join(categories),
                "features": _clean_text(m.get("features")),
                "description": _clean_text(m.get("description"))[:1000],
                "price": m.get("price"),
            })
    items = pd.DataFrame(rows)
    items["price"] = pd.to_numeric(items["price"], errors="coerce")
    return items


def prepare() -> tuple[pd.DataFrame, pd.DataFrame]:
    """Возвращает (interactions, items); результат кэшируется в parquet."""
    PROCESSED.mkdir(parents=True, exist_ok=True)
    inter_path, items_path = PROCESSED / "interactions.parquet", PROCESSED / "items.parquet"
    if inter_path.exists() and items_path.exists():
        return pd.read_parquet(inter_path), pd.read_parquet(items_path)

    inter = load_interactions()
    items = load_items(set(inter["item_id"]))
    inter = inter[inter["item_id"].isin(items["item_id"])]
    inter.to_parquet(inter_path, index=False)
    items.to_parquet(items_path, index=False)
    return inter, items


def time_split(inter: pd.DataFrame, test_share: float = 0.05):
    """Глобальное разбиение по времени: последние test_share взаимодействий идут в тест.

    Так честнее, чем случайный сплит: модель не видит будущего,
    а в тест естественно попадают новые (холодные) пользователи.
    """
    cutoff = inter["ts"].quantile(1 - test_share)
    train = inter[inter["ts"] < cutoff]
    test = inter[inter["ts"] >= cutoff]
    # товары, которых нет в train, модель рекомендовать не может - убираем из теста
    test = test[test["item_id"].isin(set(train["item_id"]))]
    return train, test, cutoff
