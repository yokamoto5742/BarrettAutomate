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

            # 列名の確認とログ出力
            self.logger.info(f"データ列: {df.columns.tolist()}")

            return df

        except Exception as e:
            self.logger.error(f"データ読み込みエラー: {e}")
            raise

    def save_patient_data(self, df: pd.DataFrame) -> None:
        """更新された患者データを結果ファイルに保存"""
        try:
            # 結果ファイルが既に存在する場合はバックアップを作成
            if self.results_file_path.exists():
                backup_path = self.results_file_path.with_suffix('_backup.xlsx')
                self.results_file_path.rename(backup_path)
                self.logger.info(f"前回結果ファイルのバックアップを作成: {backup_path}")

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
            time.sleep(1)

            # 全てのテキスト入力フィールドを取得
            all_text_inputs = page.locator('input[type="text"]').all()
            self.logger.info(f"検出されたテキスト入力フィールド数: {len(all_text_inputs)}")

            # 入力するデータを順番に準備
            input_values = [
                str(patient_row.get('DoctorName', '')),  # Doctor Name
                str(patient_row.get('PatientName', '')),  # Patient Name
                str(patient_row.get('PatientID', '')),  # Patient ID
                str(patient_row.get('LensFactor', '')),  # Lens Factor (例: 1.83)
                str(patient_row.get('AConstant', '')),  # A Constant (例: 119.0)
            ]

            # 右眼データのみを入力（左眼の位置はスキップ）
            od_values = [
                str(patient_row.get('AxialLength_R', '')),  # Axial Length (R)
                str(patient_row.get('MeasuredK1_R', '')),   # Measured K1 (R)
                str(patient_row.get('MeasuredK2_R', '')),   # Measured K2 (R)
                str(patient_row.get('OpticalACD_R', '')),   # Optical ACD (R)
                str(patient_row.get('Refraction_R', '')),   # Refraction (R)
            ]

            # 基本情報を入力
            for i, value in enumerate(input_values):
                if i < len(all_text_inputs) and value:
                    try:
                        all_text_inputs[i].clear()
                        all_text_inputs[i].fill(value)
                        field_names = ['DoctorName', 'PatientName', 'PatientID', 'LensFactor', 'AConstant']
                        self.logger.info(f"{field_names[i] if i < len(field_names) else f'Field{i}'}: {value}を入力")
                    except Exception as e:
                        self.logger.warning(f"フィールド{i}への入力エラー: {e}")

            # 基本情報の次から右眼データを入力（左眼位置はスキップ）
            measurement_start_index = len(input_values)

            # 右眼データを適切な位置に入力
            for i, value in enumerate(od_values):
                # 右眼のフィールド位置：measurement_start_index + i*2（左眼をスキップするため2つ飛ばし）
                field_index = measurement_start_index + (i * 2)
                if field_index < len(all_text_inputs) and value:
                    try:
                        all_text_inputs[field_index].clear()
                        all_text_inputs[field_index].fill(value)
                        od_field_names = ['AxialLength_R', 'MeasuredK1_R', 'MeasuredK2_R', 'OpticalACD_R', 'Refraction_R']
                        self.logger.info(f"{od_field_names[i] if i < len(od_field_names) else f'ODField{i}'}: {value}を入力")
                    except Exception as e:
                        self.logger.warning(f"右眼フィールド{i}への入力エラー: {e}")

            # 入力後少し待機
            time.sleep(1)
            self.logger.info(f"全データ入力完了: {patient_row['PatientName']}")
            return True

        except Exception as e:
            self.logger.error(f"データ入力エラー ({patient_row.get('PatientName', 'Unknown')}): {e}")
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
            page.wait_for_timeout(1000)

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
                        patient_name = row.get('PatientName', f'Patient_{index}')
                        self.logger.info(f"処理中: {patient_name} ({index + 1}/{len(df)})")

                        # Webサイトを開く
                        page.goto(self.url)
                        page.wait_for_load_state('networkidle')
                        time.sleep(1)  # 追加待機

                        # データ入力
                        if self.input_patient_data(page, row):
                            # 計算とBarrett値取得
                            iol_power = float(row.get('IOLPower', 0))
                            barrett_value = self.calculate_and_get_result(page, iol_power)

                            if barrett_value is not None:
                                df.loc[index, 'Refraction'] = barrett_value
                                successful_count += 1
                                self.logger.info(f"成功: {patient_name} → Barrett: {barrett_value}")
                            else:
                                df.loc[index, 'Refraction'] = "計算エラー"
                                error_count += 1
                        else:
                            df.loc[index, 'Refraction'] = "入力エラー"
                            error_count += 1

                        # 次の処理前に少し待機
                        time.sleep(1)

                    except Exception as e:
                        df.loc[index, 'Refraction'] = f"エラー: {str(e)[:50]}"
                        error_count += 1
                        self.logger.error(f"患者処理エラー ({patient_name}): {e}")

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
    excel_file = "APACdata.xlsx"  # Excelファイルのパス
    headless = False

    try:
        calculator = BarrettCalculator(excel_file, headless=headless)
        calculator.process_all_patients()

    except Exception as e:
        print(f"実行エラー: {e}")
        logging.error(f"メイン実行エラー: {e}")


if __name__ == "__main__":
    main()
