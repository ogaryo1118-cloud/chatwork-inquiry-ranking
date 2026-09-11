"""Streamlit版 役員報酬比較シミュレーター。"""

import pandas as pd
import streamlit as st
from calc import simulate

st.set_page_config(
    page_title="役員報酬比較シミュレーター",
    page_icon="📊",
    layout="wide",
)

st.title("役員報酬比較シミュレーター")
st.caption("東京都・令和8年度版 / PoC（源泉所得税：甲欄・扶養0人）")

with st.container(border=True):
    st.subheader("入力")
    age = st.number_input(
        "年齢",
        min_value=20,
        max_value=80,
        value=39,
        step=1,
        help="40〜64歳は介護保険料率を反映します。",
    )
    if 40 <= age <= 64:
        st.info("40〜64歳のため、介護保険料を含む健康保険料率で計算します。")

    st.markdown("**比較する役員報酬（月額・最大3案）**")
    col1, col2, col3 = st.columns(3)
    with col1:
        salary1 = st.number_input("案①", min_value=0, value=80_000, step=1_000, format="%d")
    with col2:
        salary2 = st.number_input("案②", min_value=0, value=100_000, step=1_000, format="%d")
    with col3:
        salary3 = st.number_input("案③", min_value=0, value=150_000, step=1_000, format="%d")

calculate = st.button("比較する", type="primary", use_container_width=True)

if calculate:
    salaries = [salary for salary in (salary1, salary2, salary3) if salary > 0]

    if not salaries:
        st.warning("1つ以上、役員報酬額を入力してください。")
    else:
        try:
            results = [simulate(salary, age) for salary in salaries]
            df = pd.DataFrame(results).set_index("役員報酬").T

            st.subheader("比較結果")
            df.columns = [f"{int(c):,}円" for c in df.columns]

            display_df = df.copy()
            for idx in display_df.index:
                if idx != "介護保険該当":
                    display_df.loc[idx] = display_df.loc[idx].map(
                        lambda x: f"{int(x):,}円" if pd.notna(x) else ""
                    )

            st.dataframe(display_df, use_container_width=True)

            metric_cols = st.columns(len(salaries))
            for i, (salary, result) in enumerate(zip(salaries, results)):
                with metric_cols[i]:
                    st.metric(
                        label=f"{salary:,}円案：手取り概算",
                        value=f"{result['手取り概算']:,}円",
                    )
                    st.caption(f"会社総コスト：{result['会社総コスト']:,}円")

            csv_df = pd.DataFrame(results).set_index("役員報酬").T
            st.download_button(
                "比較結果をCSVでダウンロード",
                data=csv_df.to_csv(encoding="utf-8-sig"),
                file_name="yakuin_hoshu_hikaku.csv",
                mime="text/csv",
                use_container_width=True,
            )

        except ValueError as e:
            st.error(f"計算エラー：{e}")

st.divider()
with st.expander("このツールについて"):
    st.markdown(
        """
- 計算はすべてPythonで実行します。生成AIには金額計算をさせません。
- 標準報酬月額・保険料率・源泉徴収税額表はCSVマスタとして分離しています。
- 現在の対象は **東京都・令和8年度・源泉所得税は甲欄／扶養0人** です。
- 本ツールはPoC（試作）です。実際の申告・届出・手続きでは最新の公表資料との照合が必要です。
        """
    )
