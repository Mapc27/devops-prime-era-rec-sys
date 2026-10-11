# RecSys MVP: рекомендации товаров для маркетплейса

Датасет: [Amazon Reviews 2023](https://huggingface.co/datasets/McAuley-Lab/Amazon-Reviews-2023), категория Video Games, 5-core
(814 тыс. покупок, 94.7 тыс. пользователей, 24 тыс. товаров, плотность матрицы 0.034%).
Отзыв считаем фактом покупки (implicit feedback).

## Запуск

```bash
python3.12 -m venv .venv && .venv/bin/pip install -r requirements.txt
./scripts/download_data.sh                   # данные -> data/raw/

.venv/bin/python -m recsys.evaluate          # подбор весов гибрида + сравнение (~2 мин), reports/metrics.md
.venv/bin/python -m recsys.figures           # графики для отчёта -> reports/figures/
.venv/bin/uvicorn recsys.api:app --port 8000 # REST API, документация на /docs
.venv/bin/python reports/make_docx.py --pdf   # отчёт docx + pdf (pdf через LibreOffice)
```

## Что реализовано

| Тип | Модель | Блок на сайте / эндпоинт |
|---|---|---|
| Эвристика | `Popular`, `Trending` (популярность с затуханием, half-life 30 дней) | «Популярное» `/popular`, fallback для новых пользователей |
| Эвристика | `AlsoBought` (совместные покупки в окне 30 дней, штраф за популярность) | «С этим покупают» `/item/{id}/also_bought` |
| Эвристика | `SameBrand` | «Другие товары бренда» `/item/{id}/same_brand` |
| Эвристика | История пользователя | «Недавно покупали» `/history/{user}` |
| Content-based | TF-IDF по названию, бренду, категориям, описанию | «Похожие товары» `/item/{id}/similar` |
| Collaborative | `UserKNN`, `ItemKNN` (cosine), `ALS` (библиотека implicit) | — |
| Hybrid | Switching (новый пользователь → Trending) + RRF-смешивание ALS + ItemKNN + Content + Trending | «Вам может понравиться» `/recommend/{user}` |

Код: [recsys/models.py](recsys/models.py), оценка: [recsys/evaluate.py](recsys/evaluate.py), сервис: [recsys/api.py](recsys/api.py).
Отчёт: [reports/RecSys_report.pdf](reports/RecSys_report.pdf).

## Результаты (сплит по времени: последние 5% покупок, после 2022-08-05; K=10)

Тёплые пользователи (есть история), 9 610 человек:

| Модель | Recall@10 | NDCG@10 | HitRate@10 | Coverage | Novelty | TrendShare | ms/user |
|---|---|---|---|---|---|---|---|
| Popular | 0.0066 | 0.0038 | 0.012 | 0.07% | 5.6 | 0% | 0.04 |
| Trending | 0.0367 | 0.0216 | 0.064 | 0.06% | 8.7 | 100% | 0.06 |
| AlsoBought | 0.0091 | 0.0060 | 0.018 | 37% | 10.5 | 2% | 0.04 |
| SameBrand | 0.0107 | 0.0060 | 0.019 | 16% | 8.0 | 3% | 0.09 |
| ContentBased | 0.0121 | 0.0075 | 0.022 | 56% | 12.3 | 1% | 1.01 |
| ItemKNN | 0.0126 | 0.0079 | 0.024 | 71% | 12.2 | 0% | 0.06 |
| UserKNN | 0.0143 | 0.0093 | 0.027 | 39% | 8.8 | 3% | 0.69 |
| ALS | 0.0175 | 0.0105 | 0.032 | 5% | 7.8 | 6% | 0.16 |
| Hybrid без Trending | 0.0204 | 0.0122 | 0.038 | 43% | 9.5 | 3% | 1.37 |
| **Hybrid** | **0.0464** | **0.0296** | **0.081** | 5% | 8.3 | 80% | 1.51 |

TrendShare — доля выдачи из топ-30 Trending. Веса гибрида подобраны на валидации (последние 5% train),
сохраняются в `reports/hybrid_weights.json` и используются API.

Холодные пользователи (нет истории), 1 302 человека: Popular NDCG@10 = 0.0071, Trending = 0.0232.

## Выводы (полный разбор частей II и III — в [REPORT.md](REPORT.md))

1. **Сильный бейзлайн.** Trending обходит все «чистые» ML-модели: в играх спрос сосредоточен вокруг новых релизов.
   Любую модель нужно сравнивать с бейзлайном, а не с нулём.
2. **Гибрид даёт +37% NDCG к Trending**, но платит разнообразием: coverage падает с 43% до 5%,
   а 80% выдачи — товары из топ-30 Trending (popularity bias). Это компромисс точность ↔ разнообразие из части II.
3. **Холодный старт.** Персональные модели не работают для 12% тестовых пользователей. Спасает switching на Trending.
4. **Масштабирование.** UserKNN и Content-based считают скоры на лету (0.7–1.0 мс на пользователя на 24 тыс. товаров).
   ALS и ItemKNN дешевле в инференсе, но ALS дольше обучается (10 с здесь, часы на реальном каталоге).
5. **Ограничения MVP.** Тестовое окно — 13 месяцев без переобучения. Отзыв ≠ покупка, просмотров в данных нет.
   Каталог ограничен товарами с покупками.
