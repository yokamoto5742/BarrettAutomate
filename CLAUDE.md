# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## House Rules:
- 文章ではなくパッチの差分を返すこと。Return patch diffs, not prose.
- 不明な点がある場合は、トレードオフを明記した2つの選択肢を提案すること（80語以内）。
- 変更範囲は最小限に抑えること
- Pythonコードのimport文は以下の適切な順序に並べ替えてください。
標準ライブラリ
サードパーティライブラリ
カスタムモジュール 
それぞれアルファベット順に並べます。importが先でfromは後です。

## Automatic Notifications (Hooks)
自動通知は`.claude/settings.local.json` で設定済：
- **Stop Hook**: ユーザーがClaude Codeを停止した時に「作業が完了しました」と通知
- **SessionEnd Hook**: セッション終了時に「Claude Code セッションが終了しました」と通知


## Project Overview

BarrettAutomate is a web automation tool that calculates Barrett Universal II Formula values for ophthalmology patients. It automates data entry and value extraction from the APACRS Barrett calculator website.

### Core Architecture

- **Main Script**: `barrett_calculator.py` - Contains the `BarrettCalculator` class that orchestrates the entire automation process
- **Data Flow**: Excel input → Web automation → Result extraction → Excel output
- **Browser Automation**: Uses Playwright to interact with https://calc.apacrs.org/barrett_universal2105/

### Key Components

1. **BarrettCalculator Class**: Main automation engine with methods for:
   - `load_patient_data()`: Reads patient data from Excel
   - `input_patient_data()`: Fills web form fields
   - `calculate_and_get_result()`: Executes calculation and extracts results
   - `save_patient_data()`: Saves results to output Excel file

2. **Data Processing**: Handles patient data including:
   - Patient Name, A Constant, Axial Length
   - Measured K1/K2, Optical ACD, Refraction
   - IOL Power input and Barrett value output

3. **Error Handling**: Comprehensive logging and error recovery for web automation failures

## Development Commands

### Environment Setup
```bash
# Activate virtual environment
source .venv/Scripts/activate  # Windows Git Bash
# or
.venv\Scripts\activate.bat     # Windows CMD

# Install dependencies
pip install -r requirements.txt

# Install Playwright browsers
playwright install
```

### Running the Application
```bash
# Run with browser visible (for debugging)
python barrett_calculator.py

# The script processes APACdata.xlsx by default
# Results saved to APACdata_results.xlsx
```

### Development Workflow
- Input data: Place patient data in `APACdata.xlsx`
- Main script: `barrett_calculator.py`
- Output: Results written to `APACdata_results.xlsx`
- Logs: Written to `barrett_calculator.log`

## Key Technical Details

### Dependencies
- **pandas**: Excel file processing
- **playwright**: Web browser automation
- **openpyxl**: Excel file reading/writing
- **uv**: Fast Python package installer

### Web Automation Strategy
- Uses Playwright with Chromium for stability
- Implements retry logic and error recovery
- Handles dynamic web content loading
- Extracts results from multiple table formats

### File Management
- Automatic backup creation for existing results
- Robust file path handling with pathlib
- UTF-8 encoding for international characters

## Common Tasks

### Debugging Web Automation
- Set `headless=False` in main() to watch browser execution
- Check `barrett_calculator.log` for detailed execution logs
- Adjust `slow_mo` parameter for slower automation speed

### Data Format Requirements
Excel input must contain columns:
- Patient Name, A Constant, Axial Length
- Measured K1, Measured K2, Optical ACD
- Refraction, IOL Power

### Error Recovery
- Failed calculations marked as "計算エラー" in output
- Input errors marked as "入力エラー"
- Partial results saved even if some patients fail