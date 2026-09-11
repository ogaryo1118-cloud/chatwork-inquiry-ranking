"""役員報酬比較シミュレーター（東京都・令和8年度・扶養0人）計算ロジック。"""

from pathlib import Path
import pandas as pd

BASE_DIR = Path(__file__).resolve().parent

df_standard = pd.read_csv(BASE_DIR / "master_standard_remuneration.csv")
df_rates = pd.read_csv(BASE_DIR / "master_rates.csv").set_index("item")
df_tax_low = pd.read_csv(BASE_DIR / "master_tax_table_low.csv")

RATE_KENPO = float(df_rates.loc["health_insurance", "rate"])
RATE_KENPO_KAIGO = float(df_rates.loc["health_insurance_with_kaigo", "rate"])
RATE_KOSEI_NENKIN = float(df_rates.loc["kosei_nenkin", "rate"])
RATE_SHIENKIN = float(df_rates.loc["kodomo_kosodate_shienkin", "rate"])
RATE_KYOSHUTSUKIN = float(df_rates.loc["kodomo_kosodate_kyoshutsukin", "rate"])


def get_standard_remuneration(salary: float) -> dict:
    """報酬月額から健康保険等級・厚生年金等級・標準報酬月額を返す。"""
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
        "kosei_nenkin_grade": (
            None
            if pd.isna(row["kosei_nenkin_grade"])
            else int(row["kosei_nenkin_grade"])
        ),
        "standard_remuneration": int(row["standard_remuneration"]),
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
    if not 20 <= int(age) <= 80:
        raise ValueError("年齢は20〜80歳で入力してください。")

    salary = float(salary)
    std = get_standard_remuneration(salary)
    standard_remuneration = std["standard_remuneration"]

    is_kaigo = 40 <= int(age) <= 64
    kenpo_rate = RATE_KENPO_KAIGO if is_kaigo else RATE_KENPO

    kenpo_total = round(standard_remuneration * kenpo_rate, 1)
    kenpo_half = round(kenpo_total / 2)

    if std["kosei_nenkin_grade"] is not None:
        nenkin_total = round(standard_remuneration * RATE_KOSEI_NENKIN, 1)
        nenkin_half = round(nenkin_total / 2)
        kyoshutsukin = round(standard_remuneration * RATE_KYOSHUTSUKIN, 1)
    else:
        nenkin_half = 0
        kyoshutsukin = 0

    shienkin = round(standard_remuneration * RATE_SHIENKIN, 1)

    salary_after_social = salary - kenpo_half - nenkin_half
    withholding_tax = get_withholding_tax(salary_after_social)

    honnin_futan = kenpo_half + nenkin_half + withholding_tax
    tedori = salary - honnin_futan

    kaisha_futan = kenpo_half + nenkin_half + kyoshutsukin + shienkin
    kaisha_total_cost = salary + kaisha_futan

    return {
        "役員報酬": int(salary),
        "標準報酬月額": standard_remuneration,
        "健康保険(本人)": int(kenpo_half),
        "厚生年金(本人)": int(nenkin_half),
        "源泉所得税": int(withholding_tax),
        "本人負担合計": int(honnin_futan),
        "手取り概算": int(tedori),
        "健康保険(会社)": int(kenpo_half),
        "厚生年金(会社)": int(nenkin_half),
        "子育て拠出金(会社)": int(kyoshutsukin),
        "子育て支援金(会社)": int(shienkin),
        "会社負担合計": int(kaisha_futan),
        "会社総コスト": int(kaisha_total_cost),
        "介護保険該当": "○" if is_kaigo else "-",
    }


def compare(salaries: list, age: int) -> pd.DataFrame:
    """複数の役員報酬案を横並びの比較表にする。"""
    valid_salaries = [float(s) for s in salaries if float(s) > 0]
    if not valid_salaries:
        raise ValueError("1つ以上、役員報酬額を入力してください。")

    results = [simulate(salary, age) for salary in valid_salaries]
    return pd.DataFrame(results).set_index("役員報酬").T


if __name__ == "__main__":
    demo = compare([80_000, 100_000, 150_000], 39)
    print(demo.to_string())
