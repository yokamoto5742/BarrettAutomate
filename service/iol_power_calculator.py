"""Hoffer Q / Haigis 眼内レンズ度数バッチ計算.

入力 CSV (UTF-8, ヘッダ必須) の各行について、目標屈折(既定=正視 0.00 D)を
得るための IOL 度数を Hoffer Q 式と Haigis 式で算出し、結果 CSV を書き出す。

入力列:
    patient_id, eye, al, r_mm, acd
        al    : 眼軸長 (mm)
        r_mm  : 平均角膜曲率半径 (mm)  ※度数(D)ではなく半径で渡す(下記注意参照)
        acd   : 前房深度 (mm, 角膜前面〜水晶体前面)  ※Haigis で使用
使い方:
    python iol_power_calculator.py input.csv output.csv
    python iol_power_calculator.py --predict AL R_MM ACD POWER
        例: python iol_power_calculator.py --predict 23.5 7.7 3.2 +23.5
        指定した IOL 度数を挿入した場合の術後予想屈折値(眼鏡面)を表示する。
"""

import csv
import math
import sys

from pathlib import Path

# ---- 定数(必ず OA-2000 の設定値に合わせて上書きすること)----
PACD = 5.0  # Hoffer Q: personalized ACD(= レンズの pACD 定数)
HAIGIS_A0 = 0.9  # Haigis: a0
HAIGIS_A1 = 0.4  # Haigis: a1(既定 0.4)
HAIGIS_A2 = 0.1  # Haigis: a2(既定 0.1)
TARGET_SE = 0.0  # 目標屈折(眼鏡面, D)。正視なら 0.00
VERTEX = 0.012  # 頂点間距離 (m)


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
    k = 337.5 / r_mm  # Hoffer Q は角膜屈折率 1.3375 を使用
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


# ---- 目標屈折を満たす IOL 度数を二分法で求める ----
def solve_power(se_func, target_se=0.0, lo=-10.0, hi=60.0, tol=1e-5):
    def f(p):
        return se_func(p) - target_se

    for _ in range(200):
        mid = (lo + hi) / 2
        if f(mid) > 0:  # SE は power に対し単調減少
            lo = mid
        else:
            hi = mid
        if hi - lo < tol:
            break
    return round((lo + hi) / 2, 2)


def process(in_path, out_path):
    rows_out = []
    with open(in_path, newline="", encoding="utf-8-sig") as f:
        for row in csv.DictReader(f):
            al = float(row["al"])
            r_mm = float(row["r_mm"])
            acd = float(row["acd"])
            hq = solve_power(lambda p: hofferq_se(al, r_mm, PACD, p), TARGET_SE)
            hg = solve_power(
                lambda p: haigis_se(al, r_mm, acd, HAIGIS_A0, HAIGIS_A1, HAIGIS_A2, p),
                TARGET_SE,
            )
            rows_out.append(
                {
                    "patient_id": row["patient_id"],
                    "eye": row.get("eye", ""),
                    "al": al,
                    "r_mm": r_mm,
                    "acd": acd,
                    "target_se": TARGET_SE,
                    "iol_hofferq": hq,
                    "iol_haigis": hg,
                }
            )

    fields = [
        "patient_id",
        "eye",
        "al",
        "r_mm",
        "acd",
        "target_se",
        "iol_hofferq",
        "iol_haigis",
    ]
    with open(out_path, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=fields)
        writer.writeheader()
        writer.writerows(rows_out)
    return len(rows_out)


def print_prediction(al: float, r_mm: float, acd: float, power: float) -> None:
    """指定 IOL 度数に対する術後予想屈折値(眼鏡面)を両式で表示する。"""
    hq = hofferq_se(al, r_mm, PACD, power)
    hg = haigis_se(al, r_mm, acd, HAIGIS_A0, HAIGIS_A1, HAIGIS_A2, power)
    print(f"AL={al}mm R={r_mm}mm (K={337.5 / r_mm:.2f}D) ACD={acd}mm IOL={power:+.2f}D")
    print(f"  Hoffer Q 予想屈折値: {hq:+.2f} D")
    print(f"  Haigis   予想屈折値: {hg:+.2f} D")


def main():
    if len(sys.argv) == 6 and sys.argv[1] == "--predict":
        al, r_mm, acd, power = (float(v) for v in sys.argv[2:6])
        print_prediction(al, r_mm, acd, power)
        return
    if len(sys.argv) != 3:
        print(
            "usage: python iol_power_calculator.py input.csv output.csv\n"
            "       python iol_power_calculator.py --predict AL R_MM ACD POWER"
        )
        sys.exit(1)
    n = process(Path(sys.argv[1]), Path(sys.argv[2]))
    print(f"{n} 件を計算し {sys.argv[2]} に出力しました。")


if __name__ == "__main__":
    main()
