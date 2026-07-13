import logging
import sys
import time
from datetime import datetime
from pathlib import Path
from typing import Optional

import pandas as pd
import win32api
import win32con
from playwright.sync_api import Locator, Page, sync_playwright


class KaneCalculator:
    """Kane Formula自動計算クラス"""

    def __init__(self, excel_file_path: str, headless: bool = False):
        self.file_path = Path(excel_file_path)
        self.headless = headless
        self.url = "https://www.iolformula.com/"
        self.agreement_url = "https://www.iolformula.com/agreement/"

        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s',
            handlers=[
                logging.FileHandler('kane_calculator.log', encoding='utf-8'),
                logging.StreamHandler(sys.stdout)
            ]
        )
        self.logger = logging.getLogger(__name__)

    def load_patient_data(self) -> pd.DataFrame:
        """患者データをExcelファイルから読み込み"""
        if not self.file_path.exists():
            raise FileNotFoundError(f"ファイルが見つかりません: {self.file_path}")

        df = pd.read_excel(self.file_path)
        self.logger.info(f"患者データを読み込みました: {len(df)}件")
        self.logger.info(f"データ列: {df.columns.tolist()}")
        return df

    def save_patient_data(self, df: pd.DataFrame) -> Path:
        """更新された患者データをタイムスタンプ付きの結果ファイルに保存"""
        timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
        output_path = self.file_path.parent / f"{self.file_path.stem}_results_{timestamp}{self.file_path.suffix}"
        df.to_excel(output_path, index=False, engine='openpyxl')
        self.logger.info(f"患者データをExcelファイルに保存しました: {output_path}")
        return output_path

    def open_input_form(self, page: Page) -> None:
        """入力フォームを開く（初回は利用規約に同意）"""
        page.goto(self.url)
        agree_button = page.get_by_text("I Agree", exact=True)
        if agree_button.count() > 0 and agree_button.first.is_visible():
            agree_button.first.click()
        page.wait_for_load_state('networkidle')

    def input_patient_data(self, page: Page, patient_row: pd.Series) -> bool:
        """患者データをWebフォームに入力（右眼のみ）"""
        try:
            page.get_by_role("textbox", name="Surgeon").fill(str(patient_row['Surgeon.1']))
            page.get_by_role("textbox", name="Patient").fill(str(patient_row['Patient.1']))
            page.get_by_role("textbox", name="ID").fill(str(patient_row['ID']))

            # 性別はボタン型ラベル（gender_1=M, gender_2=F）を親要素クリックで選択
            gender_name = "gender_1" if str(patient_row['Sex']).strip().upper() == "M" else "gender_2"
            page.locator(f'input[name={gender_name}]').locator('xpath=..').click()

            page.locator("#A-Constant1").fill(str(patient_row['AConstant']))
            page.locator("#right-target").fill(str(patient_row['Target refraction']))
            page.locator("#al-right").fill(str(patient_row['AL_OD']))

            # サイトは K2 >= K1 を要求するため、小さい方をK1に入力する
            k1, k2 = sorted([float(str(patient_row['K1_OD'])), float(str(patient_row['K2_OD']))])
            page.locator('input[name="k1_right"]').fill(str(k1))
            page.locator('input[name="k2_right"]').fill(str(k2))

            page.locator("#acd-right").fill(str(patient_row['ACD_OD']))
            self.logger.info(f"データ入力完了: {patient_row['Surgeon.1']}")
            return True

        except Exception as e:
            self.logger.error(f"データ入力エラー ({patient_row.get('Surgeon.1', 'Unknown')}): {e}")
            return False

    def calculate_and_get_result(self, page: Page, target_iol_power: float) -> Optional[float]:
        """計算実行と指定IOL Powerに対応するRefractionの取得"""
        try:
            page.get_by_role("button", name="Calculate").click()
            self.logger.info("Calculateボタンをクリックしました")

            result_table = self._wait_for_result_table(page)
            if result_table is None:
                self.logger.error("計算結果テーブルが表示されませんでした")
                return None

            for row in result_table.locator('tr').all()[1:]:
                cells = row.locator('td').all()
                if len(cells) < 2:
                    continue
                try:
                    iol_power = float((cells[0].text_content() or '').strip())
                    refraction = float((cells[1].text_content() or '').strip())
                except ValueError:
                    continue
                if abs(iol_power - target_iol_power) < 0.05:
                    self.logger.info(f"結果取得: IOL Power {iol_power} → Refraction {refraction}")
                    return refraction

            self.logger.warning(f"IOL Power {target_iol_power} が結果テーブルに見つかりません")
            return None

        except Exception as e:
            self.logger.error(f"計算処理エラー: {e}")
            return None

    def _wait_for_result_table(self, page: Page, timeout_seconds: int = 30) -> Optional[Locator]:
        """結果テーブルの値が確定するまで待機（計算中は '00' が表示される）"""
        for _ in range(timeout_seconds):
            page.wait_for_timeout(1000)
            table = page.locator('table:visible', has_text="IOL Power").first
            try:
                first_cell = table.locator('tr').nth(1).locator('td').first.text_content()
            except Exception:
                continue
            if first_cell and first_cell.strip() not in ("", "00"):
                return table
        return None

    def process_all_patients(self) -> None:
        """全患者データを上から順番に一括処理"""
        df = self.load_patient_data()

        # float64列に文字列（エラーメッセージ）を代入できるようobject型にする
        if 'Refraction' not in df.columns:
            df['Refraction'] = None
        df['Refraction'] = df['Refraction'].astype(object)

        successful_count = 0
        error_count = 0

        try:
            with sync_playwright() as p:
                browser = p.chromium.launch(headless=self.headless, slow_mo=500)
                context = browser.new_context(viewport={'width': 1920, 'height': 1080})
                page = context.new_page()

                for position, (index, row) in enumerate(df.iterrows(), start=1):
                    surgeon = row.get('Surgeon.1', f'Row_{index}')
                    try:
                        self.logger.info(f"処理中: {surgeon} ({position}/{len(df)})")
                        self.open_input_form(page)

                        if self.input_patient_data(page, row):
                            refraction = self.calculate_and_get_result(page, float(str(row['IOLPower'])))
                            if refraction is not None:
                                df.loc[index, 'Refraction'] = refraction
                                successful_count += 1
                                self.logger.info(f"成功: {surgeon} → Refraction: {refraction}")
                            else:
                                df.loc[index, 'Refraction'] = "計算エラー"
                                error_count += 1
                        else:
                            df.loc[index, 'Refraction'] = "入力エラー"
                            error_count += 1

                        time.sleep(1)

                    except Exception as e:
                        df.loc[index, 'Refraction'] = f"エラー: {str(e)[:50]}"
                        error_count += 1
                        self.logger.error(f"患者処理エラー ({surgeon}): {e}")

                browser.close()
        finally:
            # 途中で予期しない例外が発生しても、それまでの計算結果は必ず保存する
            output_path = self.save_patient_data(df)

            self.logger.info(f"処理完了 - 成功: {successful_count}, エラー: {error_count}")
            print("\n=== 処理結果 ===")
            print(f"成功: {successful_count}件")
            print(f"エラー: {error_count}件")
            print(f"元ファイル: {self.file_path}")
            print(f"結果ファイル: {output_path.name}")

            message = (f"Kane Calculator 処理完了\n\n成功: {successful_count}件\nエラー: {error_count}件"
                       f"\n\n結果ファイル: {output_path.name}")
            win32api.MessageBox(0, message, "処理完了", win32con.MB_OK | win32con.MB_ICONINFORMATION)


def main():
    """メイン実行関数"""
    excel_file = str(Path(__file__).resolve().parent.parent / "Kanedata.xlsx")
    headless = True

    try:
        calculator = KaneCalculator(excel_file, headless=headless)
        calculator.process_all_patients()

    except Exception as e:
        print(f"実行エラー: {e}")
        logging.error(f"メイン実行エラー: {e}")


if __name__ == "__main__":
    main()
