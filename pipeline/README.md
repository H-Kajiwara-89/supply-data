# 供給データ パイプライン

厚生労働省「医療用医薬品供給状況」Excel を毎日取得し、アプリ用の `supply.sqlite` を作って GitHub Pages で配信する。
スキーマは [../docs/xcode-prompts.md](../docs/xcode-prompts.md) の「プロンプト0」と同一（schema_version = 1）。

## ローカル実行

```bash
cd pipeline
python3 -m unittest -v
python3 build_supply_db.py --out dist
```

- `dist/supply.sqlite`：アプリに同梱する初期データ（Xcode の Resources にドラッグ）
- `dist/supply-YYYYMMDD.sqlite.deflate`：配信用（raw DEFLATE。iOS では `(data as NSData).decompressed(using: .zlib)` で展開）
- `dist/manifest.json`：`url` は manifest からの相対パス

依存ライブラリはなし（Python 3.9+ の標準ライブラリのみ）。

## 配信（GitHub Actions）

`.github/workflows/supply-data.yml` が毎日 JST 19:00 と 7:00 に実行され、`https://<user>.github.io/<repo>/v1/manifest.json` に配信する。

初回セットアップ:
1. このフォルダ（SupplyChecker/）を GitHub リポジトリにする（無料の GitHub Pages を使う場合は public）
2. Settings → Pages → Source を「GitHub Actions」にする
3. Settings → Secrets and variables → Actions → Variables に `PAGES_URL` = `https://<user>.github.io/<repo>` を登録する（前回DBから code_map を引き継ぐため）
4. Actions タブで `supply-data` を手動実行（Run workflow）し、manifest.json が見えることを確認する

失敗時（列構成の変更・品目数の急減・想定外の値）は配信せず、ワークフローが失敗する。前回のデータは配信されたまま残り、GitHub から失敗通知メールが届く。

## レセ電算コード対応表（code_map）

処方箋QRの薬品コードの多くはレセプト電算コード（9桁）なので、支払基金「医薬品マスター」で YJ に変換する。
取得元は次の優先順:
1. リポジトリ変数 `SSK_MASTER_URL`（マスター zip の直リンク）
2. `pipeline/data/y_*.zip`（手動でダウンロードしてコミットしたもの）
3. 前回配信した DB の code_map を引き継ぐ

2026-10-01 時点では支払基金サイトがメンテナンス中で取得できず、code_map は空（レセ電算コードはアプリ側で名称照合にフォールバックする）。
サイトが復旧したら、医薬品マスター（y_ALL*.zip）を `pipeline/data/` に置いてコミットするか、`SSK_MASTER_URL` を設定する。

## 注意
- public リポジトリで 60 日間コミットがないと、スケジュール実行が自動で無効化される。Actions タブの通知に従って再有効化する
- 厚労省サイトが GitHub Actions からのアクセスを拒否した場合はワークフローが失敗する（初回の手動実行で確認する）
