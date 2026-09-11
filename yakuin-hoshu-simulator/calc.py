"""役員報酬比較シミュレーター（東京都・令和8年度・扶養0人）計算ロジック。

前提:
- 協会けんぽ東京支部
- 令和8年4月分以降（子ども・子育て支援金0.23%適用後）
- 源泉所得税: 甲欄・扶養0人
- 給与から本人負担分を控除する場合の端数処理
- 会社負担額は1名分の概算（実際の納入告知額は事業所全体で合算後に端数処理）
"""

from decimal import Decimal, ROUND_CEILING, ROUND_FLOOR
from pathlib import Path

import pandas as pd

BASE_DIR = Path(__file__).resolve().parent

df_standard = pd.read_csv(BASE_DIR / "master_standard_remuneration.csv")
df_rates = pd.read_csv(
    BASE_DIR / "master_rates.csv",
    dtype={"rate": "string"},
).set_index("item")
df_tax_low = pd.read_csv(BASE_DIR / "master_tax_table_low.csv")

RATE_KENPO = Decimal(str(df_rates.loc["health_insurance", "rate"]))
RATE_KENPO_KAIGO = Decimal(str(df_rates.loc["health_insurance_with_kaigo", "rate"]))
RATE_KOSEI_NENKIN = Decimal(str(df_rates.loc["kosei_nenkin", "rate"]))
RATE_SHIENKIN = Decimal(str(df_rates.loc["kodomo_kosodate_shienkin", "rate"]))
RATE_KYOSHUTSUKIN = Decimal(str(df_rates.loc["kodomo_kosodate_kyoshutsukin", "rate"]))


def _employee_share(total_premium: Decimal) -> int:
    """給与天引き時の被保険者負担分。

    被保険者負担分の円未満が50銭以下なら切捨て、50銭超なら切上げ。
    """
    half = total_premium / Decimal("2")
    floor_value = half.to_integral_value(rounding=ROUND_FLOOR)
    fraction = half - floor_value
    if fraction <= Decimal("0.5"):
        return int(floor_value)
    return int(half.to_integral_value(rounding=ROUND_CEILING))


def _single_person_billed_total(total_premium: Decimal) -> int:
    """1名分として納付額を概算するため、全額の円未満を切り捨てる。"""
    return int(total_premium.to_integral_value(rounding=ROUND_FLOOR))


def get_standard_remuneration(salary: float) -> dict:
    """報酬月額から健康保険・厚生年金の標準報酬月額を返す。"""
    if salary < 0:
        raise ValueError("役員報酬は0円以上で入力してください。")

    row = df_standard[
        (df_standard["salary_lower"] <= salary)
        & (salary < df_standard["salary_upper"])
    ]
    if row.empty:
        raise ValueError(f"報酬月額 {salary:,.0f} 円に該当する等級が見つかりません。")

    row = row.iloc[0]
    return {
        "kenpo_grade": int(row["kenpo_grade"]),
        "kosei_nenkin_grade": int(row["kosei_nenkin_grade"]),
        "health_standard_remuneration": int(row["standard_remuneration"]),
        "pension_standard_remuneration": int(row["pension_standard_remuneration"]),
    }


def get_withholding_tax(salary_after_social_insurance: float) -> int:
    """社会保険料等控除後給与から源泉所得税（甲欄・扶養0人）を返す。"""
    s = float(salary_after_social_insurance)

    if s < 105_000:
        return 0

    if s <= 740_000:
        row = df_tax_low[
            (df_tax_low["salary_lower"] <= s)
            & (s < df_tax_low["salary_upper"])
        ]
        if row.empty:
            if s == 740_000:
                return 71_680
            raise ValueError(f"源泉徴収税額表に該当する行が見つかりません: {s:,.0f}円")
        return int(row.iloc[0]["tax_0nin"])

    if s < 790_000:
        return round(71_680 + (s - 740_000) * 0.2042)
    if s < 960_000:
        return round(81_890 + (s - 790_000) * 0.23483)
    if s < 1_710_000:
        return round(121_820 + (s - 960_000) * 0.33693)
    if s < 2_130_000:
        return round(374_520 + (s - 1_710_000) * 0.4084)
    if s < 2_170_000:
        return round(549_440 + (s - 2_130_000) * 0.4084)
    if s < 2_210_000:
        return round(571_220 + (s - 2_170_000) * 0.4084)
    if s < 2_250_000:
        return round(593_000 + (s - 2_210_000) * 0.4084)
    if s < 3_500_000:
        return round(614_770 + (s - 2_250_000) * 0.4084)
    return round(1_125_270 + (s - 3_500_000) * 0.45945)


def simulate(salary: float, age: int) -> dict:
    """役員報酬1案について本人負担・手取り・会社負担を試算する。"""
    age = int(age)
    if not 20 <= age <= 74:
        raise ValueError("本ツールの年齢対応範囲は20〜74歳です。75歳以上は後期高齢者医療制度となるため対象外です。")

    salary = int(salary)
    std = get_standard_remuneration(salary)

    health_std = std["health_standard_remuneration"]
    # 通常の厚生年金保険は70歳未満。70〜74歳は本PoCでは高齢任意加入を考慮しない。
    pension_std = std["pension_standard_remuneration"] if age < 70 else 0

    is_kaigo = 40 <= age <= 64
    kenpo_rate = RATE_KENPO_KAIGO if is_kaigo else RATE_KENPO

    health_total = Decimal(health_std) * kenpo_rate
    health_employee = _employee_share(health_total)
    health_company = _single_person_billed_total(health_total) - health_employee

    support_total = Decimal(health_std) * RATE_SHIENKIN
    support_employee = _employee_share(support_total)
    support_company = _single_person_billed_total(support_total) - support_employee

    if pension_std > 0:
        pension_total = Decimal(pension_std) * RATE_KOSEI_NENKIN
        pension_employee = _employee_share(pension_total)
        pension_company = _single_person_billed_total(pension_total) - pension_employee
        child_contribution_company = _single_person_billed_total(
            Decimal(pension_std) * RATE_KYOSHUTSUKIN
        )
    else:
        pension_employee = 0
        pension_company = 0
        child_contribution_company = 0

    salary_after_social = (
        salary
        - health_employee
        - pension_employee
        - support_employee
    )
    withholding_tax = get_withholding_tax(salary_after_social)

    employee_total = (
        health_employee
        + pension_employee
        + support_employee
        + withholding_tax
    )
    take_home = salary - employee_total

    company_total = (
        health_company
        + pension_company
        + child_contribution_company
        + support_company
    )
    company_total_cost = salary + company_total

    return {
        "役員報酬": salary,
        "健康保険 標準報酬月額": health_std,
        "厚生年金 標準報酬月額": pension_std,
        "健康保険(本人)": health_employee,
        "厚生年金(本人)": pension_employee,
        "子ども・子育て支援金(本人)": support_employee,
        "社会保険料等控除後給与": salary_after_social,
        "源泉所得税": withholding_tax,
        "本人負担合計": employee_total,
        "手取り概算": take_home,
        "健康保険(会社)": health_company,
        "厚生年金(会社)": pension_company,
        "子ども・子育て拠出金(会社)": child_contribution_company,
        "子ども・子育て支援金(会社)": support_company,
        "会社負担合計": company_total,
        "会社総コスト": company_total_cost,
        "介護保険該当": "○" if is_kaigo else "-",
    }


def compare(salaries: list, age: int) -> pd.DataFrame:
    """複数の役員報酬案を横並びの比較表にする。"""
    valid_salaries = [int(s) for s in salaries if int(s) > 0]
    if not valid_salaries:
        raise ValueError("1つ以上、役員報酬額を入力してください。")

    results = [simulate(salary, age) for salary in valid_salaries]
    return pd.DataFrame(results).set_index("役員報酬").T


if __name__ == "__main__":
    demo = compare([80_000, 100_000, 150_000], 39)
    print(demo.to_string())
