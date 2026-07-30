import streamlit as st
from pathlib import Path

from picker import (
    load_products,
    available_genres,
    pick_products,
)

EXCEL_FILE = Path(__file__).with_name(
    "ドトールメニュー商品価格一覧_2026-07-30(1).xlsx"
)

st.title("ドトール商品ピッカー")

products = load_products(EXCEL_FILE)
genres = available_genres(products)

selected_genres = st.multiselect(
    "ジャンルを選択してください",
    genres,
)

genre_counts = {}

for genre in selected_genres:
    genre_counts[genre] = st.number_input(
        f"{genre}から選ぶ個数",
        min_value=0,
        value=1,
        step=1,
    )

budget = st.number_input(
    "上限金額",
    min_value=0,
    value=1500,
    step=10,
)

if st.button("商品を抽選する"):
    # ここでpick_productsを実行
    pass
