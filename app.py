from pathlib import Path
import random

import streamlit as st

from picker import (
    available_genres,
    calculate_discount,
    load_products,
    payable_total,
    pick_products,
)


# =========================================================
# 基本設定
# =========================================================

st.set_page_config(
    page_title="ドトール商品ピッカー",
    page_icon="☕",
    layout="centered",
)

BASE_DIR = Path(__file__).resolve().parent

# 実際のExcelファイル名に合わせてください
EXCEL_FILE = BASE_DIR / "ドトールメニュー商品価格一覧_2026-07-30(1).xlsx"


# =========================================================
# 商品データの読み込み
# =========================================================

@st.cache_data
def get_products():
    """
    Excelから商品一覧を読み込む。
    Excelを変更した際は、Streamlitのキャッシュを削除するか、
    アプリを再起動してください。
    """
    return load_products(EXCEL_FILE)


try:
    products = get_products()
except FileNotFoundError:
    st.error(
        "Excelファイルが見つかりません。\n\n"
        f"探した場所：`{EXCEL_FILE.name}`\n\n"
        "app.pyとExcelファイルを同じフォルダに置いてください。"
    )
    st.stop()
except Exception as error:
    st.error(f"商品データの読み込みに失敗しました：{error}")
    st.stop()


genres = available_genres(products)


# =========================================================
# セッション情報
# =========================================================

if "result" not in st.session_state:
    st.session_state.result = None

if "last_conditions" not in st.session_state:
    st.session_state.last_conditions = None


# =========================================================
# 画面
# =========================================================

st.title("☕ ドトール商品ピッカー")

st.write(
    "ジャンルごとの個数と上限金額を指定すると、"
    "条件を満たす商品をランダムに抽選します。"
)

selected_genres = st.multiselect(
    "ジャンルを選択してください",
    options=genres,
    placeholder="複数選択できます",
)

genre_counts: dict[str, int] = {}

if selected_genres:
    st.subheader("ジャンルごとの個数")

    for genre in selected_genres:
        genre_counts[genre] = int(
            st.number_input(
                f"{genre}から選ぶ個数",
                min_value=0,
                max_value=20,
                value=1,
                step=1,
                key=f"count_{genre}",
            )
        )

budget = int(
    st.number_input(
        "割引後の合計金額の上限",
        min_value=0,
        max_value=100000,
        value=1500,
        step=10,
        format="%d",
    )
)

st.caption("ドリンク1点につき、セット割引は1回だけ適用されます。")


# =========================================================
# 抽選処理
# =========================================================

def run_lottery() -> None:
    """現在入力されている条件で商品を抽選する。"""

    if not selected_genres:
        st.warning("ジャンルを1つ以上選択してください。")
        return

    positive_counts = {
        genre: count
        for genre, count in genre_counts.items()
        if count > 0
    }

    if not positive_counts:
        st.warning("少なくとも1つのジャンルで、個数を1以上にしてください。")
        return

    if budget <= 0:
        st.warning("上限金額を1円以上にしてください。")
        return

    try:
        with st.spinner("条件を満たす商品を探しています…"):
            result = pick_products(
                products=products,
                genre_counts=positive_counts,
                budget=budget,
                rng=random.Random(),
            )

        st.session_state.result = result
        st.session_state.last_conditions = {
            "genre_counts": positive_counts.copy(),
            "budget": budget,
        }

    except ValueError as error:
        st.session_state.result = None
        st.error(str(error))

    except Exception as error:
        st.session_state.result = None
        st.exception(error)


if st.button(
    "商品を抽選する",
    type="primary",
    use_container_width=True,
):
    run_lottery()


# =========================================================
# 結果表示
# =========================================================

result = st.session_state.result

if result:
    st.divider()
    st.subheader("抽選結果")

    for index, product in enumerate(result, start=1):
        size = ""

        if product.size and str(product.size) not in {"—", "-", "なし", "None"}:
            size = f"（{product.size}）"

        st.markdown(
            f"**{index}. {product.name}{size}**  \n"
            f"{product.genre}・{product.price:,}円"
        )

    normal_total = sum(product.price for product in result)
    payment = payable_total(result)
    discount = normal_total - payment

    st.divider()

    col1, col2, col3 = st.columns(3)

    with col1:
        st.metric(
            "通常合計",
            f"{normal_total:,}円",
        )

    with col2:
        st.metric(
            "割引額",
            f"-{discount:,}円",
        )

    with col3:
        st.metric(
            "支払額",
            f"{payment:,}円",
        )

    remaining = budget - payment

    if remaining >= 0:
        st.success(f"上限金額まで残り {remaining:,}円です。")

    # 割引の詳しい内訳を表示
    try:
        discount_amount, discount_details = calculate_discount(result)

        if discount_amount > 0 and discount_details:
            with st.expander("割引の内訳"):
                for detail in discount_details:
                    st.write(detail)

    except Exception:
        # 割引内訳の表示だけ失敗しても、
        # 抽選結果そのものは表示する
        pass

    st.write("")

    if st.button(
        "同じ条件で再抽選する",
        use_container_width=True,
    ):
        run_lottery()
        st.rerun()
