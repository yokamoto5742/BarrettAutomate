import logging
import re
import sys
import time
from pathlib import Path
from typing import Optional

import pandas as pd
from playwright.sync_api import sync_playwright


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
        """患者データをWebフォームに入力"""
        try:
            # ページが完全に読み込まれるまで待機
            page.wait_for_load_state('networkidle')
            time.sleep(2)

            # Patient Name入力（修正：2番目のテキスト入力フィールドを取得）
            try:
                # 全てのテキスト入力フィールドを取得し、Patient Name用の2番目を選択
                all_text_inputs = page.locator('input[type="text"]').all()
                if len(all_text_inputs) >= 2:
                    patient_name_input = all_text_inputs[1]  # 2番目のフィールド（0から始まるため1）
                    patient_name_input.clear()
                    patient_name_input.fill(str(patient_row['Patient Name']))
                    self.logger.info(f"Patient Name入力完了: {patient_row['Patient Name']}")
                else:
                    self.logger.warning("Patient Name用のフィールドが見つかりません")
            except Exception as e:
                self.logger.warning(f"Patient Name入力スキップ: {e}")

            # A Constant入力（Lens Factorの右側の入力フィールド）
            try:
                # A Constant範囲（112～125）の入力フィールドを探す
                a_constant_inputs = page.locator('input[type="text"]').all()
                for input_field in a_constant_inputs:
                    # フィールドの近くに"112~125"のテキストがあるかチェック
                    parent_text = input_field.locator('..').text_content()
                    if "112" in parent_text and "125" in parent_text:
                        input_field.clear()
                        input_field.fill(str(patient_row['A Constant']))
                        self.logger.info(f"A Constant入力完了: {patient_row['A Constant']}")
                        break
            except Exception as e:
                self.logger.warning(f"A Constant入力エラー: {e}")

            # 右眼（OD）の測定値入力
            od_inputs = page.locator('table tr').filter(has_text='OD').locator('input[type="text"]').all()

            if len(od_inputs) >= 5:  # Axial Length, K1, K2, ACD, Refraction
                try:
                    # Axial Length (R)
                    od_inputs[0].clear()
                    od_inputs[0].fill(str(patient_row['Axial Length']))

                    # Measured K1 (R)
                    od_inputs[1].clear()
                    od_inputs[1].fill(str(patient_row['Measured K1']))

                    # Measured K2 (R)
                    od_inputs[2].clear()
                    od_inputs[2].fill(str(patient_row['Measured K2']))

                    # Optical ACD (R)
                    od_inputs[3].clear()
                    od_inputs[3].fill(str(patient_row['Optical ACD']))

                    # Refraction (R)
                    od_inputs[4].clear()
                    od_inputs[4].fill(str(patient_row['Refraction']))

                    self.logger.info(f"右眼データ入力完了: {patient_row['Patient Name']}")

                except Exception as e:
                    self.logger.error(f"右眼データ入力エラー: {e}")

            # 少し待機してフォームが更新されるのを待つ
            time.sleep(1)
            return True

        except Exception as e:
            self.logger.error(f"データ入力エラー ({patient_row['Patient Name']}): {e}")
            return False

    def calculate_and_get_result(self, page, target_iol_power: float) -> Optional[float]:
        """計算実行とBarrett値の取得"""
        try:
            # Calculateボタンをクリック
            calculate_btn = page.locator('input[value="Calculate"]').first
            if calculate_btn.is_visible():
                calculate_btn.click()
                self.logger.info("Calculateボタンをクリックしました")
                page.wait_for_timeout(3000)  # 計算完了を待機
            else:
                self.logger.error("Calculateボタンが見つかりません")
                return None

            # Universal Formulaタブをクリック
            universal_formula_btn = page.locator('a:has-text("Universal Formula")').first
            if universal_formula_btn.is_visible():
                universal_formula_btn.click()
                self.logger.info("Universal Formulaタブをクリックしました")
                page.wait_for_timeout(3000)  # ページ遷移を待機
            else:
                self.logger.error("Universal Formulaタブが見つかりません")
                return None

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
            page.wait_for_timeout(2000)

            # IOL PowerとRefractionの値を含むテーブル行を探す
            table_rows = page.locator('table tr').all()

            for row in table_rows:
                cells = row.locator('td').all()
                if len(cells) >= 3:  # IOL Power, Optic, Refraction
                    try:
                        # IOL Powerの値を取得（最初の列）
                        iol_power_text = cells[0].text_content().strip()

                        # 数値のみを抽出
                        iol_match = re.search(r'(\d+\.?\d*)', iol_power_text)
                        if not iol_match:
                            continue

                        iol_power = float(iol_match.group(1))

                        # 目標のIOL Powerと一致するかチェック（0.1の誤差を許容）
                        if abs(iol_power - target_iol_power) < 0.1:
                            # Refractionの値を取得（3番目の列）
                            refraction_text = cells[2].text_content().strip()

                            # 負の値も含む数値を抽出
                            refraction_match = re.search(r'(-?\d+\.?\d*)', refraction_text)
                            if refraction_match:
                                refraction_value = float(refraction_match.group(1))
                                self.logger.info(
                                    f"テーブルから抽出: IOL Power {iol_power} → Refraction {refraction_value}")
                                return refraction_value

                    except (ValueError, IndexError) as e:
                        self.logger.debug(f"行解析スキップ: {e}")
                        continue

            # テーブルから直接値が見つからない場合、別の方法を試す
            return self._extract_refraction_alternative(page, target_iol_power)

        except Exception as e:
            self.logger.error(f"テーブル解析エラー: {e}")
            return None

    def _extract_refraction_alternative(self, page, target_iol_power: float) -> Optional[float]:
        """代替方法でRefractionを抽出"""
        try:
            # ページ内のすべてのテキストからIOL PowerとRefractionのペアを探す
            page_content = page.content()

            # IOL PowerとRefractionの組み合わせをより柔軟に検索
            patterns = [
                rf'{target_iol_power}.*?(-?\d+\.?\d*)',
                rf'>{target_iol_power}<.*?(-?\d+\.?\d*)',
                rf'{target_iol_power}\s*</td>.*?(-?\d+\.?\d*)</td>'
            ]

            for pattern in patterns:
                matches = re.findall(pattern, page_content, re.DOTALL)
                if matches:
                    try:
                        refraction_value = float(matches[0])
                        self.logger.info(
                            f"代替方法で抽出: IOL Power {target_iol_power} → Refraction {refraction_value}")
                        return refraction_value
                    except ValueError:
                        continue

            # ハイライトされた行（青い背景）を特別に探す
            highlighted_rows = page.locator('tr[style*="background"], tr.highlighted').all()
            for row in highlighted_rows:
                row_text = row.text_content()
                if str(target_iol_power) in row_text:
                    refraction_match = re.search(r'(-?\d+\.?\d*)\s*$', row_text.strip())
                    if refraction_match:
                        refraction_value = float(refraction_match.group(1))
                        self.logger.info(
                            f"ハイライト行から抽出: IOL Power {target_iol_power} → Refraction {refraction_value}")
                        return refraction_value

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
                browser = p.chromium.launch(headless=self.headless, slow_mo=500)  # slow_moで処理を少し遅くする
                context = browser.new_context(
                    viewport={'width': 1920, 'height': 1080},
                    user_agent='Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
                )
                page = context.new_page()

                successful_count = 0
                error_count = 0

                for index, row in df.iterrows():
                    try:
                        self.logger.info(f"処理中: {row['Patient Name']} ({index + 1}/{len(df)})")

                        # Webサイトを開く
                        page.goto(self.url)
                        page.wait_for_load_state('networkidle')
                        time.sleep(2)  # 追加待機

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
                        time.sleep(2)

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