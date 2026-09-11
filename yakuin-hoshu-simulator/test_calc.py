"""主要ケースの回帰テスト。"""

from calc import simulate

EXPECTED_39 = {
    80_000: {
        "健康保険 標準報酬月額": 78_000,
        "厚生年金 標準報酬月額": 88_000,
        "健康保険(本人)": 3_841,
        "厚生年金(本人)": 8_052,
        "子ども・子育て支援金(本人)": 90,
        "源泉所得税": 0,
        "手取り概算": 68_017,
        "会社負担合計": 12_299,
        "会社総コスト": 92_299,
    },
    100_000: {
        "健康保険 標準報酬月額": 98_000,
        "厚生年金 標準報酬月額": 98_000,
        "健康保険(本人)": 4_826,
        "厚生年金(本人)": 8_967,
        "子ども・子育て支援金(本人)": 113,
        "源泉所得税": 0,
        "手取り概算": 86_094,
        "会社負担合計": 14_258,
        "会社総コスト": 114_258,
    },
    150_000: {
        "健康保険 標準報酬月額": 150_000,
        "厚生年金 標準報酬月額": 150_000,
        "健康保険(本人)": 7_387,
        "厚生年金(本人)": 13_725,
        "子ども・子育て支援金(本人)": 172,
        "源泉所得税": 1_300,
        "手取り概算": 127_416,
        "会社負担合計": 21_826,
        "会社総コスト": 171_826,
    },
}


def test_default_cases():
    for salary, expected in EXPECTED_39.items():
        actual = simulate(salary, 39)
        for key, value in expected.items():
            assert actual[key] == value, (
                f"{salary=} {key}: expected {value}, got {actual[key]}"
            )


def test_age_70_has_no_regular_pension_premium():
    actual = simulate(150_000, 70)
    assert actual["厚生年金 標準報酬月額"] == 0
    assert actual["厚生年金(本人)"] == 0
    assert actual["厚生年金(会社)"] == 0
    assert actual["子ども・子育て拠出金(会社)"] == 0


if __name__ == "__main__":
    test_default_cases()
    test_age_70_has_no_regular_pension_premium()
    print("OK: all tests passed")
