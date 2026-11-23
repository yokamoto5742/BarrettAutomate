# 変更履歴

このプロジェクトのすべての重要な変更は、このファイルに記録されます。

フォーマットは [Keep a Changelog](https://keepachangelog.com/ja/1.1.0/) に基づいており、
バージョン番号は [Semantic Versioning](https://semver.org/lang/ja/) に従っています。

## [Unreleased]

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
