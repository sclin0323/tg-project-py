"""
main.py — 主入口
用法: python main.py

每次執行會在 output/ 底下建立一個新的快照目錄：
    output/YYYYMMDD_NNN/
        ├── 專案管理總表SDC.xlsx
        ├── WBS_*.xlsx                (所有 WBS 來源檔)
        └── 專案報告_YYYYMMDD.pptx
NNN 從 001 開始，同一天執行多次會自動遞增。
"""
import os
import shutil
from datetime import date
from data_loader import build_report
from ppt_builder import build_ppt

# 以 main.py 所在位置為基準，避免「目前工作目錄」造成的路徑問題
HERE = os.path.dirname(os.path.abspath(__file__))

EXCEL_DIR  = os.path.join(HERE, "excel")    # Excel 來源（總表 + 6 個 WBS）
OUTPUT_DIR = os.path.join(HERE, "output")   # 輸出根目錄


# ──────────────────────────────────────────────────────
# 建立下個可用的快照目錄  output/YYYYMMDD_NNN
# ──────────────────────────────────────────────────────
def make_next_snapshot_dir(output_root: str, today: date) -> str:
    os.makedirs(output_root, exist_ok=True)
    date_tag = today.strftime("%Y%m%d")
    seq = 1
    while True:
        path = os.path.join(output_root, f"{date_tag}_{seq:03d}")
        if not os.path.exists(path):
            os.makedirs(path)
            return path
        seq += 1


# ──────────────────────────────────────────────────────
# 把 Excel 來源（總表 + 6 個 WBS）複製到快照目錄
# ──────────────────────────────────────────────────────
def copy_excel_snapshot(excel_dir: str, snapshot_dir: str) -> int:
    count = 0
    for fname in sorted(os.listdir(excel_dir)):
        if not fname.lower().endswith((".xlsx", ".xlsm")):
            continue
        # 跳過 Excel 暫存檔 ~$xxx.xlsx
        if fname.startswith("~$"):
            continue
        src = os.path.join(excel_dir, fname)
        dst = os.path.join(snapshot_dir, fname)
        shutil.copy2(src, dst)
        count += 1
    return count


# ──────────────────────────────────────────────────────
# 主流程
# ──────────────────────────────────────────────────────
def main():
    # 因為這是模擬資料，固定 today = 2026/05/12
    # 實務上要產報告時改為 date.today() 即可
    today = date(2026, 5, 20)

    print(f"=== 載入資料 (today = {today}) ===")
    print(f"    Excel 來源: {EXCEL_DIR}")
    if not os.path.isdir(EXCEL_DIR):
        raise FileNotFoundError(
            f"找不到 Excel 資料夾: {EXCEL_DIR}\n"
            f"請在 {HERE} 底下建立 excel 子資料夾，並把總表與 WBS 檔案放進去。"
        )

    # Step 1: 建立快照目錄
    snapshot_dir = make_next_snapshot_dir(OUTPUT_DIR, today)
    print(f"\n=== 建立快照目錄 ===")
    print(f"    📁 {snapshot_dir}")

    # Step 2: 複製 Excel 來源到快照目錄
    n = copy_excel_snapshot(EXCEL_DIR, snapshot_dir)
    print(f"    📄 已複製 {n} 個 Excel 檔")

    # Step 3: 載入與分析資料
    report = build_report(EXCEL_DIR, today=today)
    print(f"\n=== 分析結果 ===")
    print(f"    整體：{report['overall']} | {report['headline']}")
    print(f"    紅 {report['stats']['red_projects']}　·　"
          f"黃 {report['stats']['amber_projects']}　·　"
          f"綠 {report['stats']['green_projects']}")

    # Step 4: 產出 PPT 到快照目錄
    pptx_path = os.path.join(snapshot_dir,
                             f"專案報告_{today.strftime('%Y%m%d')}.pptx")
    build_ppt(report, pptx_path)
    print(f"\n=== 已產出 PPT ===")
    print(f"    🎯 {pptx_path}")
    return pptx_path


if __name__ == "__main__":
    main()