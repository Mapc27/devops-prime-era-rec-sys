#!/usr/bin/env bash
# Скачивает Amazon Reviews 2023 (Video Games, 5-core) в data/raw/
set -euo pipefail
cd "$(dirname "$0")/.."
mkdir -p data/raw
B=https://huggingface.co/datasets/McAuley-Lab/Amazon-Reviews-2023/resolve/main
curl -fL -o data/raw/Video_Games.csv "$B/benchmark/5core/rating_only/Video_Games.csv"
curl -fL -o data/raw/meta_Video_Games.jsonl "$B/raw/meta_categories/meta_Video_Games.jsonl"
