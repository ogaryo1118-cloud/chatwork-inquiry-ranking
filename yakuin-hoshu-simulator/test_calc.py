"""Google Sheets版との基準一致確認。"""
from calc import simulate

EXPECTED = {
    80_000: {
        "標準報酬月額": 78_000,
        "健康保険(本人)": 3_842,
        "厚生年金(本人)": 0,
        "源泉所得税": 0,
        "手取り概算": 76_158,
        "会社負担合計": 4_021,
        "会社総コスト": 84_021,
    },
    100_000: {
        "標準報酬月額": 98_000,
        "健康保険(本人)": 4_826,
        "厚生年金(本人)": 8_967,
        "源泉所得税": 0,
        "手取り概算": 86_207,
        "会社負担合計": 14_371,
        "会社総コスト": 114_371,
    },
    150_000: {
        "標準報酬月額": 150_000,
        "健康保険(本人)": 7_388,
        "厚生年金(本人)": 13_725,
        "源泉所得税": 1_300,
        "手取り概算": 127_587,
        "会社負担合計": 21_998,
        "会社総コスト": 171_998,
    },
}

for salary, expected in EXPECTED.items():
    actual = simulate(salary, 39)
    for key, expected_value in expected.items():
        assert actual[key] == expected_value, (
            f"{salary:,}円 / {key}: expected={expected_value}, actual={actual[key]}"
        )

print("OK: Google Sheets版の基準値（39歳・8万/10万/15万円）と一致しました。")
