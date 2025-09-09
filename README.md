# analytics.app

サーバーレスで動作する小売分析ダッシュボードの雛形です。

## ディレクトリ構成
```
analytics.app/
├─ data/
│   └─ product_master.csv   # 商品マスタ（JAN, 商品分類=部門コード, 商品名）
├─ README.md                # プロジェクト説明
```

## 商品マスタについて
- CSV形式（UTF-8-SIG）
- 列構成: 商品コード(JAN), 商品分類(部門コード), 商品名
- 未登録JANは「未分類」として扱います。
