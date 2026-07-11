"""Hoffer Q / Haigis 術後予想屈折値バッチ計算.

IOLCalcdata.xlsx の各行を上から順に処理し、指定 IOL 度数(IOLPower 列)を
挿入した場合の術後予想等価球面度数(眼鏡面, D)を Hoffer Q 式と Haigis 式で
計算して HofferQ_SE / Haigis_SE 列に書き込み、タイムスタンプ付きの新しい
Excel ファイルに保存する(既存ファイルの上書きは行わない)。

入力列:
    AL          : 眼軸長 (mm)
    K1, K2      : 角膜屈折力 (D)
    ACD         : 前房深度 (mm, 角膜前面〜水晶体前面)  ※Haigis で使用
    Hoffer_pACD : Hoffer Q の personalized ACD
    a0, a1, a2  : Haigis 定数
    IOLPower    : 挿入する IOL 度数 (D)

使い方:
    python iol_power_calculator.py [Excelファイルパス]
        パス省略時は ../IOLCalcdata.xlsx を読み込む。
"""

import math
import sys
from datetime import datetime
from pathlib import Path

import pandas as pd

VERTEX = 0.012  # 頂点間距離 (m)
KERATOMETER_INDEX = 337.5  # K(D) と角膜曲率半径(mm)の換算値 (n=1.3375)


# ---- Hoffer Q ----
def _mg(al):
    return (1.0, 28.0) if al <= 23.0 else (-1.0, 23.5)


def _clamp_al(al):
    return max(18.5, min(31.0, al))


def _hofferq_pred_acd(al, k, pacd):
    m, g = _mg(al)
    a = _clamp_al(al)
    p2 = 0.3 * (a - 23.5)
    p3 = math.tan(math.radians(k)) ** 2
    p4 = 0.1 * m * (23.5 - a) ** 2
    p5 = math.tan(math.radians(0.1 * (g - a) ** 2))
    return pacd + p2 + p3 + p4 * p5 - 0.99166


def hofferq_se(al, r_mm, pacd, power):
    """IOL 度数 power を入れたときの予測等価球面度数(眼鏡面, D)を返す。"""
    k = KERATOMETER_INDEX / r_mm  # Hoffer Q は角膜屈折率 1.3375 を使用
    acd = _hofferq_pred_acd(al, k, pacd)
    rc = (
        1.336 / (1.336 / (1336.0 / (al - acd - 0.05) - power) + (acd + 0.05) / 1000.0)
        - k
    )
    return rc / (1 + VERTEX * rc)  # 角膜面 -> 眼鏡面


# ---- Haigis ----
def haigis_se(al, r_mm, acd_meas, a0, a1, a2, power):
    """IOL 度数 power を入れたときの予測等価球面度数(眼鏡面, D)を返す。"""
    n, nc = 1.336, 1.3315  # Haigis は角膜屈折率 1.3315 を使用
    d = a0 + a1 * acd_meas + a2 * al  # ELP (mm)
    big_r = r_mm / 1000.0
    length = al / 1000.0
    dd = d / 1000.0
    dc = (nc - 1) / big_r
    qnum = n * (n - power * (length - dd))
    qden = n * (length - dd) + dd * (n - power * (length - dd))
    q = qnum / qden
    return (q - dc) / (1 + VERTEX * (q - dc))


def process_excel(excel_path: Path) -> tuple[Path, int]:
    """Excel の全行を計算し、結果を新規ファイルに保存してパスと件数を返す。"""
    df = pd.read_excel(excel_path)

    for index, row in df.iterrows():
        values = row.to_dict()
        al = float(values["AL"])
        k_mean = (float(values["K1"]) + float(values["K2"])) / 2
        r_mm = KERATOMETER_INDEX / k_mean
        acd = float(values["ACD"])
        power = float(values["IOLPower"])

        hq = hofferq_se(al, r_mm, float(values["Hoffer_pACD"]), power)
        hg = haigis_se(
            al,
            r_mm,
            acd,
            float(values["a0"]),
            float(values["a1"]),
            float(values["a2"]),
            power,
        )
        df.loc[index, "HofferQ_SE"] = round(hq, 2)
        df.loc[index, "Haigis_SE"] = round(hg, 2)

    timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
    output_path = (
        excel_path.parent / f"{excel_path.stem}_results_{timestamp}{excel_path.suffix}"
    )
    df.to_excel(output_path, index=False, engine="openpyxl")
    return output_path, len(df)


def main() -> None:
    excel_file = sys.argv[1] if len(sys.argv) > 1 else "../IOLCalcdata.xlsx"
    excel_path = Path(excel_file)

    if not excel_path.exists():
        print(f"ファイルが見つかりません: {excel_path}")
        sys.exit(1)

    output_path, count = process_excel(excel_path)
    print(f"{count} 件を計算し {output_path} に出力しました。")


if __name__ == "__main__":
    main()
