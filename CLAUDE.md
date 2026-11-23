# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.必ず日本語で回答してください。

## Project overview

Barrett Universal II Formula(眼科のIOL度数計算式)の計算を自動化するツール。`service/barrett_calculator.py` の `BarrettCalculator` が Excel(`data/Barrettdata.xlsx` 等)から患者データを読み込み、Playwright で https://calc.apacrs.org/barrett_universal2105/ を操作して右眼データのみを入力・計算し、結果を新しい timestamped Excel ファイルに保存する(既存ファイルの上書き・バックアップは行わない)。Windows専用(`pywin32` で処理完了時にメッセージボックス通知)。

## Environment

- Python >= 3.13、パッケージ管理は `uv`(`uv.lock` あり)
- 主要依存: `pandas`、`openpyxl`、`playwright`
- pyright は `app`、`service`、`utils`、`tests` のみを型チェック対象とし、`scripts` は除外

## Gotchas

- `service/barrett_calculator.py` は `win32api`/`win32con`(pywin32)を import しているが、`pyproject.toml` の依存関係には含まれていない。Windows環境の別途インストールが前提。
- `service/barrett_calculator.py` の `main()` は Excelファイルパスが `"../data/Barrettdata.xlsx"` にハードコードされている。

## Commands

```bash
# 依存インストール
uv sync

# Playwrightブラウザのインストール(初回のみ)
playwright install chromium

# 型チェック
pyright

# Lint
ruff check .
```

テストコマンドは `.claude/rules/testing.md` を参照。コーディング規約は `.claude/rules/python-coding.md`、コミット規約は `.claude/rules/commit.md`、レビュー前の行動指針は `.claude/rules/coding-guidelines.md` を参照。
