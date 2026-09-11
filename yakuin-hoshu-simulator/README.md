# 役員報酬比較シミュレーター（Python / Streamlit版）

Googleスプレッドシート版と同じ基準値になるように作成したPython版です。

## 対応範囲

- 東京都
- 令和8年度
- 源泉所得税：甲欄・扶養0人
- 年齢20〜80歳
- 役員報酬：最大3案比較

## Streamlit Community Cloud で公開する場合

- Repository: `ogaryo1118-cloud/chatwork-inquiry-ranking`
- Branch: `main`
- Main file path: `yakuin-hoshu-simulator/app.py`

## ファイル

- `app.py`：Streamlitの画面
- `calc.py`：計算ロジック
- `master_standard_remuneration.csv`：標準報酬月額マスタ
- `master_rates.csv`：保険料率マスタ
- `master_tax_table_low.csv`：源泉徴収税額表マスタ
- `test_calc.py`：Google Sheets版との基準一致テスト
- `requirements.txt`：必要ライブラリ

## ローカル起動

```bash
python -m pip install -r requirements.txt
python -m streamlit run app.py
```

## 計算一致テスト

```bash
python test_calc.py
```

`OK: Google Sheets版の基準値...と一致しました。` と表示されれば基準テスト合格です。

> 注意：PoC（試作）です。実際の申告・届出・手続きでは最新の公表資料との照合が必要です。
