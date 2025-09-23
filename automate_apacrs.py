import pandas as pd
from playwright.sync_api import sync_playwright, Page, TimeoutError
import time

# --- 設定 ---
# 入力ファイルと出力ファイルの名前 (Excel形式)
INPUT_EXCEL_PATH = 'APACdata.xlsx'
OUTPUT_EXCEL_PATH = 'APACdata_results.xlsx'
# 計算サイトのURL
TARGET_URL = 'https://calc.apacrs.org/barrett_universal2105/'


def get_refraction_value(page: Page, target_iol_power: float) -> float | None:
    """
    Barrett Universal II Formulaの画面から、指定されたIOL Powerに
    対応するRefractionの値を取得します。
    """
    try:
        # 結果テーブルはiframe内に表示されるため、iframeを取得
        # ページの読み込みが遅い場合を考慮して、最大60秒待機
        frame = page.frame_locator('iframe[name="iframe_RIGHT"]').first
        frame.wait_for_load_state(timeout=60000)

        # テーブル内のすべての行を取得
        rows = frame.locator('table tr').all()

        # ヘッダー行を除いて、各行をループ処理
        for row in rows[1:]:
            cols = row.locator('td').all()
            if len(cols) >= 2:
                # 1列目 (IOL Power) と 2列目 (Refraction) のテキストを取得
                iol_power_text = cols[0].inner_text().strip()
                refraction_text = cols[1].inner_text().strip()

                try:
                    # テキストを浮動小数点数に変換
                    iol_power_in_table = float(iol_power_text)

                    # テーブルのIOL Powerと目標のIOL Powerが一致するか確認
                    if iol_power_in_table == target_iol_power:
                        print(f"  ✅ IOL Power {target_iol_power} に一致する屈折度: {refraction_text} を見つけました。")
                        return float(refraction_text)
                except ValueError:
                    # 数値に変換できないセルはスキップ
                    continue

        print(f"  ⚠️ IOL Power {target_iol_power} がテーブルに見つかりませんでした。")
        return None

    except TimeoutError:
        print("  ❌ タイムアウト: 結果ページの読み込みに失敗しました。")
        return None
    except Exception as e:
        print(f"  ❌ 結果取得中に予期せぬエラーが発生しました: {e}")
        return None


def main():
    """
    メインの自動化処理
    """
    print(f"処理を開始します: {INPUT_EXCEL_PATH}")

    # 1. Excelファイルをpandasで読み込む
    try:
        # read_excelを使用
        df = pd.read_excel(INPUT_EXCEL_PATH)
    except FileNotFoundError:
        print(f"エラー: 入力ファイル '{INPUT_EXCEL_PATH}' が見つかりません。")
        return
    except Exception as e:
        print(f"エラー: Excelファイルの読み込み中に問題が発生しました: {e}")
        return

    # Playwrightの実行ブロック
    with sync_playwright() as p:
        # ブラウザを起動 (headless=Trueにするとバックグラウンドで実行)
        browser = p.chromium.launch(headless=False)
        page = browser.new_page()

        # 各患者データについてループ処理
        for index, row in df.iterrows():
            # 'Patient Name' 列が存在するか確認
            if 'Patient Name' in row:
                patient_name = row['Patient Name']
            else:
                patient_name = f"Index {index}"  # 存在しない場合はインデックスを使用

            print(f"\n--- 患者 '{patient_name}' (データ {index + 1}/{len(df)}件目) の処理を開始 ---")

            try:
                # 2. 計算サイトを開く
                page.goto(TARGET_URL, timeout=60000)

                # 3. Patient dataの各項目に値を入力
                #    str()で文字列に変換して入力の確実性を高める
                page.locator('input[name="axiallength"]').fill(str(row['Axial Length']))
                page.locator('input[name="k1"]').fill(str(row['Measured K1']))
                page.locator('input[name="k2"]').fill(str(row['Measured K2']))
                page.locator('input[name="acd"]').fill(str(row['Optical ACD']))
                page.locator('input[name="aconstant"]').fill(str(row['A Constant']))
                print("  - データ入力完了")

                # 4. "Calculate" ボタンをクリック
                page.get_by_role("button", name="Calculate").click()
                print("  - 'Calculate' ボタンをクリック")

                # 5. "Universal Formula" ボタンをクリック
                page.get_by_role("button", name="Universal Formula").click()
                print("  - 'Universal Formula' ボタンをクリック")

                # 6. 表示されたテーブルからRefractionを取得
                target_iol = float(row['IOL Power'])
                refraction_val = get_refraction_value(page, target_iol)

                # 7. 取得した値をDataFrameの'Barrett'列に格納
                if refraction_val is not None:
                    df.at[index, 'Barrett'] = refraction_val

                # 次の処理のために少し待機（任意）
                time.sleep(1)

            except TimeoutError:
                print(
                    f"  ❌ タイムアウト: ページの読み込みまたは要素の検索に失敗しました。この患者の処理をスキップします。")
                continue
            except Exception as e:
                print(f"  ❌ 予期せぬエラーが発生しました: {e}。この患者の処理をスキップします。")
                continue

        # ブラウザを閉じる
        browser.close()

    # 8. 結果を新しいExcelファイルに保存
    # to_excelを使用し、index=Falseで不要な行番号の出力を防ぐ
    df.to_excel(OUTPUT_EXCEL_PATH, index=False)
    print(f"\n--- 全ての処理が完了しました ---")
    print(f"結果は '{OUTPUT_EXCEL_PATH}' に保存されました。")


if __name__ == "__main__":
    main()
