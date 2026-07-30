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
    
    try:
        if not selected_genres:
            st.warning("ジャンルを1つ以上選択してください。")
        elif sum(genre_counts.values()) == 0:
            st.warning("商品の個数を1個以上にしてください。")
        else:
            with st.spinner("商品を選んでいます…"):
                result = pick_products(
                    products=products,
                    genre_counts=genre_counts,
                    budget=int(budget),
                    rng=random.Random(),
                )

            st.success("抽選しました")

            for product in result:
                size = f"（{product.size}）" if product.size != "—" else ""
                st.write(
                    f"**{product.name}{size}**　"
                    f"{product.price:,}円"
                )

            normal_total = sum(product.price for product in result)
            payment = payable_total(result)
            discount = normal_total - payment

            st.divider()
            st.write(f"通常合計：{normal_total:,}円")
            st.write(f"割引：-{discount:,}円")
            st.write(f"支払額：**{payment:,}円**")

    except Exception as error:
        st.error(f"エラーが発生しました：{error}")
