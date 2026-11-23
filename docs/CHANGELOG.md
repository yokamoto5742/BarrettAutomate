# 変更履歴

このプロジェクトのすべての重要な変更は、このファイルに記録されます。

フォーマットは [Keep a Changelog](https://keepachangelog.com/ja/1.1.0/) に基づいており、
バージョン番号は [Semantic Versioning](https://semver.org/lang/ja/) に従っています。

## [Unreleased]

## [1.0.5] - 2026-11-23
### 変更
- `service/barrett_calculator.py` で患者1人を処理するごとに結果ファイルへ保存するよう変更(途中で処理が止まってもそれまでの計算結果が残るようにするため)
- 結果ファイルの保存に失敗しても処理を継続するよう変更(処理中に結果ファイルをExcelで開いていても止まらないようにするため)

### 修正
- 完了メッセージに表示する結果ファイル名が実際の保存ファイル名とずれることがある問題を修正

### 追加
- `tests/service/test_barrett_calculator.py` に結果保存処理のテストを追加

## [1.0.4] - 2026-11-23
### 削除
- `service/barrett_calculator.py` の代替抽出処理 `_extract_refraction_alternative` を削除。結果テーブルに目標の IOL Power が見つからない場合は「計算エラー」として扱う(ページ内の無関係な数値を誤って取得するのを防ぐため)

## [1.0.3] - 2026-11-23
### 修正
- `service/barrett_calculator.py` で IOL Power の読み込み列名を `IOLPower` から `IOL_power` に修正(列が見つからず IOL Power が 0.0 として扱われ、誤った Refraction を取得していた)

## [1.0.2] - 2026-11-23
### 修正
- `service/barrett_calculator.py` の読み込みExcelファイルを `data/Barrettdata.xlsx`(旧 `APACdata.xlsx`)に修正

## [1.0.1] - 2026-11-23
### 変更
- `service/barrett_calculator.py` の読み込みExcelファイルを `data/IOLCalcdata.xlsx` に変更

## [1.0.0] - 2026-07-10
- 初版リリース
