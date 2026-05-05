import os
import pandas as pd
import openpyxl
from datetime import datetime

EXCEL_PATH = "./excel/專案管理總表SDC.xlsx"
EXCEL_DIR  = "./excel/WBS2026"   # WBS 子檔案資料夾


def extract_wbs_links(path: str) -> dict:
    """從總表 G 欄讀取超連結，回傳 {專案名稱: WBS檔案路徑}"""
    wb = openpyxl.load_workbook(path)
    ws = wb["WBS清單總表"]
    links = {}
    for row in ws.iter_rows(min_row=2):
        project_name = row[1].value
        wbs_cell     = row[6]
        if not project_name:
            continue
        if wbs_cell.hyperlink:
            raw = wbs_cell.hyperlink.target
            filename = os.path.basename(raw.replace("\\", "/"))
            local_path = os.path.join(EXCEL_DIR, filename)
            links[project_name] = local_path
    return links


def read_master(path: str) -> pd.DataFrame:
    df = pd.read_excel(path, sheet_name="WBS清單總表")
    for col in ["開始日期", "目標日期"]:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col]).dt.strftime("%Y/%m/%d")
    return df


def read_wbs(path: str, project_name: str) -> pd.DataFrame | None:
    if not os.path.exists(path):
        print(f"  ⚠️  找不到檔案：{path}")
        return None
    try:
        df = pd.read_excel(path, sheet_name=0)
        df["_專案"] = project_name
        for col in ["開始日", "預計結束日", "實際完成日"]:
            if col in df.columns:
                df[col] = pd.to_datetime(df[col], errors="coerce").dt.strftime("%Y/%m/%d")
        return df
    except Exception as e:
        print(f"  ❌ 讀取失敗 {path}：{e}")
        return None


def check_risk(df: pd.DataFrame) -> list:
    risks = []
    today = datetime.today()
    for _, row in df.iterrows():
        task    = row.get("工作項目", "")
        status  = str(row.get("狀態", ""))
        end_str = row.get("預計結束日", "")
        actual  = row.get("實際完成日", "")

        if status in ["延遲", "風險"]:
            risks.append(f"🔴 [{row['_專案']}] {task}｜狀態：{status}")
            continue

        try:
            end_date = datetime.strptime(end_str, "%Y/%m/%d")
            if end_date < today and str(actual) in ["NaT", "nan", "", "None"]:
                risks.append(f"🟡 [{row['_專案']}] {task}｜截止 {end_str} 已過但未完成")
        except Exception:
            pass

    return risks


# ── 主流程 ────────────────────────────────────────────

print("=== 📋 專案管理總表 ===")
df_master = read_master(EXCEL_PATH)
print(df_master.to_string(index=False))

print("\n=== 🔗 讀取 WBS 子檔案 ===")
wbs_links = extract_wbs_links(EXCEL_PATH)
all_wbs = []

for project, wbs_path in wbs_links.items():
    print(f"\n▶ {project}  →  {wbs_path}")
    df_wbs = read_wbs(wbs_path, project)
    if df_wbs is not None:
        print(f"  ✅ 共 {len(df_wbs)} 筆任務")
        print(df_wbs.to_string(index=False))
        all_wbs.append(df_wbs)

# ── 風險彙整
if all_wbs:
    df_all = pd.concat(all_wbs, ignore_index=True)
    print("\n=== ⚠️  風險任務彙整 ===")
    risks = check_risk(df_all)
    if risks:
        for r in risks:
            print(f"  {r}")
    else:
        print("  ✅ 目前無風險任務")