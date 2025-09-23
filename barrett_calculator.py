import logging
import sys
import time
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import pandas as pd
from playwright.sync_api import Playwright, sync_playwright


class BarrettCalculator:
    """Barrett Universal II Formula自動計算クラス"""

    def __init__(self, excel_file_path: str, headless: bool = False):
        self.excel_file_path = Path(excel_file_path)
        self.results_file_path = Path(excel_file_path).with_name('APACdata_results.xlsx')
        self.headless = headless
        self.url = "https://calc.apacrs.org/barrett_universal2105/"

        # ログ設定
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s',
            handlers=[
                logging.FileHandler('barrett_calculator.log', encoding='utf-8'),
                logging.StreamHandler(sys.stdout)
            ]
        )
        self.logger = logging.getLogger(__name__)

    def load_patient_data(self) -> pd.DataFrame:
        """患者データをExcelファイルから読み込み"""
        try:
            if not self.excel_file_path.exists():
                raise FileNotFoundError(f"ファイルが見つかりません: {self.excel_file_path}")

            df = pd.read_excel(self.excel_file_path)
            self.logger.info(f"患者データを読み込みました: {len(df)}件")

            # 必要な列が存在するかチェック
            required_columns = [
                'Patient Name', 'A Constant', 'Axial Length',
                'Measured K1', 'Measured K2', 'Optical ACD',
                'Refraction', 'IOL Power'
            ]

            missing_columns = [col for col in required_columns if col not in df.columns]
            if missing_columns:
                raise ValueError(f"必要な列が不足しています: {missing_columns}")

            # Barrett列が存在しない場合は作成
            if 'Barrett' not in df.columns:
                df['Barrett'] = None
                self.logger.info("Barrett列を作成しました")

            return df

        except Exception as e:
            self.logger.error(f"データ読み込みエラー: {e}")
            raise

    def save_patient_data(self, df: pd.DataFrame) -> None:
        """更新された患者データを結果ファイルに保存"""
        try:
            # 結果ファイルが既に存在する場合はバックアップを作成
            if self.results_file_path.exists():
                backup_path = self.results_file_path.with_suffix('.backup.xlsx')
                self.results_file_path.rename(backup_path)
                self.logger.info(f"既存結果ファイルのバックアップを作成: {backup_path}")

            # 更新されたデータを結果ファイルに保存
            df.to_excel(self.results_file_path, index=False)
            self.logger.info(f"処理結果を保存しました: {self.results_file_path}")

        except Exception as e:
            self.logger.error(f"データ保存エラー: {e}")
            raise

    def input_patient_data(self, page, patient_row: pd.Series) -> bool:
        """患者データをウェブフォームに入力"""
        try:
            # Patient Name入力
            patient_name_input = page.locator('input[placeholder*="Patient Name"], input[name*="patient"]').first
            if patient_name_input.is_visible():
                patient_name_input.fill(str(patient_row['Patient Name']))

            # A Constant入力
            a_constant_input = page.locator('input[value="119.0"]').first
            if a_constant_input.is_visible():
                a_constant_input.fill(str(patient_row['A Constant']))

            # Axial Length (R)入力
            axial_length_input = page.locator('input[value="23.04"]').first
            if axial_length_input.is_visible():
                axial_length_input.fill(str(patient_row['Axial Length']))

            # Measured K1 (R)入力
            k1_input = page.locator('input[value="44.75"]').first
            if k1_input.is_visible():
                k1_input.fill(str(patient_row['Measured K1']))

            # Measured K2 (R)入力
            k2_input = page.locator('input[value="44.25"]').first
            if k2_input.is_visible():
                k2_input.fill(str(patient_row['Measured K2']))

            # Optical ACD (R)入力
            acd_input = page.locator('input[value="2.18"]').first
            if acd_input.is_visible():
                acd_input.fill(str(patient_row['Optical ACD']))

            # Refraction (R)入力
            refraction_input = page.locator('input[value="-0.03"]').first
            if refraction_input.is_visible():
                refraction_input.fill(str(patient_row['Refraction']))

            self.logger.info(f"患者データを入力しました: {patient_row['Patient Name']}")
            return True

        except Exception as e:
            self.logger.error(f"データ入力エラー ({patient_row['Patient Name']}): {e}")
            return False

    def calculate_and_get_result(self, page, target_iol_power: float) -> Optional[float]:
        """計算実行とBarrett値の取得"""
        try:
            # Calculateボタンをクリック
            calculate_btn = page.locator('input[value="Calculate"], button:has-text("Calculate")').first
            if calculate_btn.is_visible():
                calculate_btn.click()
                page.wait_for_timeout(2000)  # 計算完了を待機

            # Universal Formulaボタンをクリック
            universal_formula_btn = page.locator(
                'a:has-text("Universal Formula"), button:has-text("Universal Formula")').first
            if universal_formula_btn.is_visible():
                universal_formula_btn.click()
                page.wait_for_timeout(3000)  # ページ遷移を待機

            # 結果テーブルからIOL Powerに対応するRefractionを取得
            refraction_value = self._extract_refraction_from_table(page, target_iol_power)

            if refraction_value is not None:
                self.logger.info(f"Barrett値を取得: IOL Power {target_iol_power} → Refraction {refraction_value}")
            else:
                self.logger.warning(f"Barrett値が見つかりません: IOL Power {target_iol_power}")

            return refraction_value

        except Exception as e:
            self.logger.error(f"計算処理エラー: {e}")
            return None

    def _extract_refraction_from_table(self, page, target_iol_power: float) -> Optional[float]:
        """結果テーブルからIOL Powerに対応するRefractionを抽出"""
        try:
            # テーブルが表示されるまで待機
            page.wait_for_selector('table, .table, [role="table"]', timeout=10000)

            # IOL PowerとRefractionの値を含むセルを探す
            table_rows = page.locator('tr').all()

            for row in table_rows:
                cells = row.locator('td, th').all()
                if len(cells) >= 3:  # IOL Power, Optic, Refraction
                    try:
                        # IOL Powerの値を取得
                        iol_power_text = cells[0].text_content().strip()
                        iol_power = float(iol_power_text)

                        # 目標のIOL Powerと一致するかチェック
                        if abs(iol_power - target_iol_power) < 0.1:  # 0.1の誤差を許容
                            refraction_text = cells[2].text_content().strip()
                            refraction_value = float(refraction_text)
                            return refraction_value

                    except (ValueError, IndexError):
                        continue

            # テーブルから直接値が見つからない場合、別の方法を試す
            return self._extract_refraction_alternative(page, target_iol_power)

        except Exception as e:
            self.logger.error(f"テーブル解析エラー: {e}")
            return None

    def _extract_refraction_alternative(self, page, target_iol_power: float) -> Optional[float]:
        """代替方法でRefractionを抽出"""
        try:
            # ページ内のすべてのテキストから数値ペアを探す
            page_content = page.content()

            # IOL PowerとRefractionのペアを正規表現で探す
            import re
            pattern = rf'{target_iol_power}\s*.*?(-?\d+\.?\d*)'
            matches = re.findall(pattern, page_content)

            if matches:
                try:
                    return float(matches[0])
                except ValueError:
                    pass

            return None

        except Exception as e:
            self.logger.error(f"代替抽出エラー: {e}")
            return None

    def process_all_patients(self) -> None:
        """全患者データの一括処理"""
        try:
            # データ読み込み
            df = self.load_patient_data()

            with sync_playwright() as p:
                browser = p.chromium.launch(headless=self.headless)
                context = browser.new_context()
                page = context.new_page()

                successful_count = 0
                error_count = 0

                for index, row in df.iterrows():
                    try:
                        self.logger.info(f"処理中: {row['Patient Name']} ({index + 1}/{len(df)})")

                        # ウェブサイトを開く
                        page.goto(self.url)
                        page.wait_for_load_state('networkidle')

                        # データ入力
                        if self.input_patient_data(page, row):
                            # 計算とBarrett値取得
                            barrett_value = self.calculate_and_get_result(page, float(row['IOL Power']))

                            if barrett_value is not None:
                                df.loc[index, 'Barrett'] = barrett_value
                                successful_count += 1
                                self.logger.info(f"成功: {row['Patient Name']} → Barrett: {barrett_value}")
                            else:
                                df.loc[index, 'Barrett'] = "計算エラー"
                                error_count += 1
                        else:
                            df.loc[index, 'Barrett'] = "入力エラー"
                            error_count += 1

                        # 次の処理前に少し待機
                        time.sleep(1)

                    except Exception as e:
                        df.loc[index, 'Barrett'] = f"エラー: {str(e)[:50]}"
                        error_count += 1
                        self.logger.error(f"患者処理エラー ({row['Patient Name']}): {e}")

                browser.close()

            # 結果保存
            self.save_patient_data(df)

            # サマリー出力
            self.logger.info(f"処理完了 - 成功: {successful_count}, エラー: {error_count}")
            print(f"\n=== 処理結果 ===")
            print(f"成功: {successful_count}件")
            print(f"エラー: {error_count}件")
            print(f"元ファイル: {self.excel_file_path}")
            print(f"結果ファイル: {self.results_file_path}")

        except Exception as e:
            self.logger.error(f"一括処理エラー: {e}")
            raise


def main():
    """メイン実行関数"""
    # 設定
    excel_file = "APACdata.xlsx"  # Excelファイルのパス
    headless = False  # ブラウザを表示する場合はFalse、非表示の場合はTrue

    try:
        calculator = BarrettCalculator(excel_file, headless=headless)
        calculator.process_all_patients()

    except Exception as e:
        print(f"実行エラー: {e}")
        logging.error(f"メイン実行エラー: {e}")


if __name__ == "__main__":
    main()