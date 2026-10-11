"""Рекомендательные модели: эвристики, content-based, коллаборативные, гибрид.

У всех моделей общий интерфейс:
    fit(ctx)            - обучение на train из контекста
    score(uidx)         - матрица скоров (len(uidx), n_items) для тёплых пользователей
    similar(iidx, k)    - похожие товары для item-to-item блоков (если модель это умеет)
"""
import os
import warnings

os.environ.setdefault("OPENBLAS_NUM_THREADS", "1")  # рекомендация implicit для ALS

import numpy as np
import pandas as pd
import scipy.sparse as sp
from implicit.cpu.als import AlternatingLeastSquares
from implicit.nearest_neighbours import CosineRecommender
from implicit.utils import ParameterWarning
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.preprocessing import normalize


class Context:
    """Общие для всех моделей структуры: индексы, матрица user x item, каталог."""

    def __init__(self, train: pd.DataFrame, items: pd.DataFrame):
        self.train = train
        self.item_ids = np.array(sorted(train["item_id"].unique()))
        self.user_ids = np.array(sorted(train["user_id"].unique()))
        self.item2idx = {i: n for n, i in enumerate(self.item_ids)}
        self.user2idx = {u: n for n, u in enumerate(self.user_ids)}
        self.items = items.set_index("item_id").loc[self.item_ids].reset_index()

        rows = train["user_id"].map(self.user2idx).to_numpy()
        cols = train["item_id"].map(self.item2idx).to_numpy()
        X = sp.csr_matrix((np.ones(len(rows), dtype=np.float32), (rows, cols)),
                          shape=(len(self.user_ids), len(self.item_ids)))
        X.data[:] = 1.0  # повторные покупки не учитываем
        self.X = X
        self.item_pop = np.asarray(X.sum(axis=0)).ravel()
        self.max_ts = train["ts"].max()

        # история пользователя по времени (для "недавно купленных" и профилей)
        hist = train.sort_values("ts")
        self.history = (hist.assign(i=hist["item_id"].map(self.item2idx))
                        .groupby(hist["user_id"].map(self.user2idx))["i"].apply(np.array).to_dict())

    @property
    def n_items(self):
        return len(self.item_ids)


def _topk_sparse_rows(M: sp.csr_matrix, k: int) -> sp.csr_matrix:
    """Оставляет k наибольших значений в каждой строке разреженной матрицы."""
    M = M.tocsr()
    rows, cols, vals = [], [], []
    for r in range(M.shape[0]):
        s, e = M.indptr[r], M.indptr[r + 1]
        if e == s:
            continue
        d, c = M.data[s:e], M.indices[s:e]
        top = np.argpartition(-d, min(k, len(d) - 1))[:k] if len(d) > k else np.arange(len(d))
        rows.append(np.full(len(top), r)); cols.append(c[top]); vals.append(d[top])
    if not rows:
        return sp.csr_matrix(M.shape, dtype=np.float32)
    return sp.csr_matrix((np.concatenate(vals), (np.concatenate(rows), np.concatenate(cols))), shape=M.shape)


# ---------------------------------------------------------------- эвристики

class Popular:
    """Самые продаваемые товары за всё время: бейзлайн и fallback."""
    name = "Popular"

    def fit(self, ctx):
        self.scores = ctx.item_pop.astype(np.float32)
        return self

    def score(self, uidx):
        return np.tile(self.scores, (len(uidx), 1))


class Trending(Popular):
    """Популярность с затуханием по времени: свежие покупки весят больше (half-life 30 дней)."""
    name = "Trending"

    def __init__(self, half_life_days: float = 30):
        self.half_life = half_life_days

    def fit(self, ctx):
        age = (ctx.max_ts - ctx.train["ts"]).dt.days.to_numpy()
        w = 0.5 ** (age / self.half_life)
        idx = ctx.train["item_id"].map(ctx.item2idx).to_numpy()
        self.scores = np.bincount(idx, weights=w, minlength=ctx.n_items).astype(np.float32)
        return self


class AlsoBought:
    """«С этим товаром также покупают»: пары товаров, купленные одним человеком в пределах окна.

    Сырые счётчики совместных покупок тянут в топ хиты («бананы» в любой корзине),
    поэтому делим на sqrt(популярности) рекомендуемого товара.
    Рекомендации пользователю = сумма соседей по последним купленным товарам.
    """
    name = "AlsoBought"

    def __init__(self, window_days: int = 30, max_lag: int = 5, min_support: int = 2, last_n: int = 5):
        self.window, self.max_lag, self.min_support, self.last_n = window_days, max_lag, min_support, last_n

    def fit(self, ctx):
        df = ctx.train.sort_values(["user_id", "ts"])
        u = df["user_id"].to_numpy()
        i = df["item_id"].map(ctx.item2idx).to_numpy()
        t = df["ts"].to_numpy()
        a_all, b_all = [], []
        for lag in range(1, self.max_lag + 1):
            same = (u[lag:] == u[:-lag]) & ((t[lag:] - t[:-lag]) <= np.timedelta64(self.window, "D"))
            a_all += [i[:-lag][same], i[lag:][same]]
            b_all += [i[lag:][same], i[:-lag][same]]
        a, b = np.concatenate(a_all), np.concatenate(b_all)
        C = sp.csr_matrix((np.ones(len(a), dtype=np.float32), (a, b)), shape=(ctx.n_items, ctx.n_items))
        C.sum_duplicates()
        C.data[C.data < self.min_support] = 0
        C.eliminate_zeros()
        C = C @ sp.diags(1 / np.sqrt(np.maximum(ctx.item_pop, 1)).astype(np.float32))
        self.sim = _topk_sparse_rows(C, 50)
        self.ctx = ctx
        return self

    def score(self, uidx):
        rows = [self.ctx.history[u][-self.last_n:] for u in uidx]
        P = sp.csr_matrix((np.ones(sum(map(len, rows)), dtype=np.float32),
                           (np.repeat(np.arange(len(rows)), [len(r) for r in rows]), np.concatenate(rows))),
                          shape=(len(uidx), self.ctx.n_items))
        return (P @ self.sim).toarray()

    def similar(self, iidx, k=10):
        row = self.sim[iidx]
        return row.indices[np.argsort(-row.data)][:k].tolist()


class SameBrand:
    """«Товары того же бренда»: популярное среди брендов, которые пользователь уже покупал."""
    name = "SameBrand"

    def fit(self, ctx):
        brands = ctx.items["brand"].where(ctx.items["brand"] != "unknown")
        codes, self.brand_names = pd.factorize(brands)
        known = codes >= 0
        self.B = sp.csr_matrix((np.ones(known.sum(), dtype=np.float32), (np.arange(ctx.n_items)[known], codes[known])),
                               shape=(ctx.n_items, len(self.brand_names)))
        self.codes = codes
        self.pop_norm = (ctx.item_pop / ctx.item_pop.max()).astype(np.float32)
        self.ctx = ctx
        return self

    def score(self, uidx):
        profile = self.ctx.X[uidx] @ self.B  # сколько раз покупал каждый бренд
        return (profile @ self.B.T).toarray() + self.pop_norm  # бренд важнее, популярность - тай-брейк

    def similar(self, iidx, k=10):
        same = np.where((self.codes == self.codes[iidx]) & (self.codes >= 0))[0]
        same = same[same != iidx]
        return same[np.argsort(-self.pop_norm[same])][:k].tolist()


# ---------------------------------------------------------------- content-based

class ContentBased:
    """TF-IDF по названию, бренду, категориям, характеристикам и описанию.

    Профиль пользователя - взвешенная сумма векторов последних покупок
    (свежие весят больше), скор = косинусная близость профиля и товара.
    """
    name = "ContentBased"

    def __init__(self, last_n: int = 10):
        self.last_n = last_n

    def fit(self, ctx):
        it = ctx.items
        text = (it["title"] + " " + it["brand"] + " " + it["categories"] + " "
                + it["features"] + " " + it["description"])
        vec = TfidfVectorizer(max_features=50_000, min_df=2, stop_words="english", sublinear_tf=True)
        self.V = normalize(vec.fit_transform(text)).astype(np.float32)
        self.ctx = ctx
        return self

    def score(self, uidx):
        rows, cols, vals = [], [], []
        for r, u in enumerate(uidx):
            h = self.ctx.history[u][-self.last_n:]
            rows += [r] * len(h); cols += list(h); vals += list(np.linspace(0.5, 1.0, len(h)))
        P = sp.csr_matrix((vals, (rows, cols)), shape=(len(uidx), self.ctx.n_items), dtype=np.float32)
        return (normalize(P @ self.V) @ self.V.T).toarray()

    def similar(self, iidx, k=10):
        s = (self.V[iidx] @ self.V.T).toarray().ravel()
        s[iidx] = -np.inf
        return np.argsort(-s)[:k].tolist()


# ---------------------------------------------------------------- коллаборативные

class ItemKNN:
    """Item-based CF: косинусная близость товаров по векторам покупателей (implicit)."""
    name = "ItemKNN"

    def __init__(self, k: int = 100):
        self.k = k

    def fit(self, ctx):
        model = CosineRecommender(K=self.k)
        with warnings.catch_warnings():  # implicit сам внутри передаёт себе coo вместо csr
            warnings.simplefilter("ignore", ParameterWarning)
            model.fit(ctx.X, show_progress=False)
        self.sim = model.similarity.tocsr().astype(np.float32)
        self.ctx = ctx
        return self

    def score(self, uidx):
        return (self.ctx.X[uidx] @ self.sim).toarray()

    def similar(self, iidx, k=10):
        row = self.sim[iidx].toarray().ravel()
        row[iidx] = -np.inf
        return [i for i in np.argsort(-row)[:k] if row[i] > 0]


class UserKNN:
    """User-based CF: находим k самых похожих покупателей и берём их покупки.

    Считается «на лету» для каждого запроса - это и есть главная проблема
    масштабирования user-based подхода (см. отчёт, часть II).
    """
    name = "UserKNN"

    def __init__(self, k: int = 50):
        self.k = k

    def fit(self, ctx):
        self.Xn = normalize(ctx.X).astype(np.float32)
        self.ctx = ctx
        return self

    def score(self, uidx):
        S = (self.Xn[uidx] @ self.Xn.T).toarray()
        S[np.arange(len(uidx)), uidx] = 0  # себя не считаем соседом
        drop = np.argpartition(-S, self.k, axis=1)[:, self.k:]
        np.put_along_axis(S, drop, 0, axis=1)
        return (sp.csr_matrix(S) @ self.ctx.X).toarray()


class ALS:
    """Матричная факторизация (Alternating Least Squares) для неявного фидбека."""
    name = "ALS"

    def __init__(self, factors: int = 64, iterations: int = 20, regularization: float = 0.05, alpha: float = 10.0):
        self.params = dict(factors=factors, iterations=iterations, regularization=regularization, alpha=alpha)

    def fit(self, ctx):
        self.model = AlternatingLeastSquares(**self.params, random_state=42)
        self.model.fit(ctx.X, show_progress=False)
        return self

    def score(self, uidx):
        return self.model.user_factors[uidx] @ self.model.item_factors.T

    def similar(self, iidx, k=10):
        ids, _ = self.model.similar_items(iidx, N=k + 1)
        return [i for i in ids.tolist() if i != iidx][:k]


# ---------------------------------------------------------------- гибрид

class Hybrid:
    """Гибрид двух видов.

    1) Switching - решение холодного старта: у нового пользователя нет истории,
       ему показываем Trending (см. recommend_for_user).
    2) Weighted - для тёплых пользователей смешиваем ALS, ItemKNN, ContentBased и Trending
       через Reciprocal Rank Fusion: складываем w / (60 + позиция товара).
       RRF не требует калибровки скоров разных моделей между собой.
       Trending добавляет «свежесть»: в играх спрос сильно смещён к новым релизам.
       Веса подбираются на валидационном срезе (см. evaluate.tune_hybrid), а не на тесте.
    """

    def __init__(self, components: dict, top_n: int = 200, rrf_k: int = 60, name: str = "Hybrid"):
        self.components, self.top_n, self.rrf_k, self.name = components, top_n, rrf_k, name

    def fit(self, ctx):
        self.ctx = ctx
        return self

    def score(self, uidx):
        out = np.zeros((len(uidx), self.ctx.n_items), dtype=np.float32)
        for model, w in self.components.items():
            if w:
                out += w * rank_contrib(model, self.ctx, uidx, self.top_n, self.rrf_k).toarray()
        return out


def rank_contrib(model, ctx, uidx, top_n: int = 200, rrf_k: int = 60) -> sp.csr_matrix:
    """Вклад модели в RRF: 1 / (rrf_k + позиция) для её top_n товаров.

    Уже купленное выкидываем до отбора кандидатов, иначе оно занимает места в top_n
    (у ALS в среднем ~5 из 200), а потом всё равно вычёркивается из выдачи.
    """
    s = model.score(uidx).astype(np.float32)
    s[ctx.X[uidx].nonzero()] = -np.inf
    top = np.argpartition(-s, top_n, axis=1)[:, :top_n]
    order = np.argsort(-np.take_along_axis(s, top, axis=1), axis=1)
    top = np.take_along_axis(top, order, axis=1)
    vals = np.tile(1 / (rrf_k + np.arange(1, top_n + 1, dtype=np.float32)), len(uidx))
    rows = np.repeat(np.arange(len(uidx)), top_n)
    return sp.csr_matrix((vals, (rows, top.ravel())), shape=(len(uidx), ctx.n_items))


def recommend(model, ctx: Context, uidx: np.ndarray, k: int = 10, batch: int = 512) -> np.ndarray:
    """Top-k товаров для тёплых пользователей, без уже купленного."""
    result = []
    for s in range(0, len(uidx), batch):
        b = uidx[s:s + batch]
        scores = model.score(b).astype(np.float32)
        seen = ctx.X[b]
        scores[seen.nonzero()] = -np.inf
        top = np.argpartition(-scores, k, axis=1)[:, :k]
        order = np.argsort(-np.take_along_axis(scores, top, axis=1), axis=1)
        result.append(np.take_along_axis(top, order, axis=1))
    return np.vstack(result)
