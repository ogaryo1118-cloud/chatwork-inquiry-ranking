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
st.caption("東京都・令和8年4月分以降 / PoC（源泉所得税：甲欄・扶養0人）")

with st.container(border=True):
    st.subheader("入力")
    age = st.number_input(
        "年齢",
        min_value=20,
        max_value=74,
        value=39,
        step=1,
        help="40〜64歳は介護保険料を反映。70〜74歳は通常の厚生年金保険料を0円として計算します。75歳以上は対象外です。",
    )
    if 40 <= age <= 64:
        st.info("40〜64歳のため、介護保険料を含む健康保険料率で計算します。")
    elif 70 <= age <= 74:
        st.info("70歳以上のため、通常の厚生年金保険料・子ども子育て拠出金は0円として計算します（高齢任意加入は考慮しません）。")

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
- **令和8年4月分以降・東京都・協会けんぽ**を前提にしています。
- 健康保険料率9.85%、介護保険料率1.62%、厚生年金保険料率18.3%を使用します。
- **子ども・子育て支援金（0.23%）は本人と会社が原則折半**です。
- **子ども・子育て拠出金（0.36%）は会社負担のみ**です。
- 本人負担の円未満は、給与天引き時の公式端数処理（50銭以下切捨て、50銭超切上げ）を反映しています。
- 健康保険と厚生年金では標準報酬月額の下限・上限が異なるため、それぞれ別に判定しています。
- 70〜74歳は通常の厚生年金保険料を0円として計算します（高齢任意加入は考慮しません）。75歳以上は本ツール対象外です。
- 会社負担額は**1名分の概算**です。実際の納入告知額は事業所全体で合算後に端数処理されるため、数円の差が生じる場合があります。
- 源泉所得税は令和8年分の月額表・甲欄・扶養0人を前提としています。
        """
    )
