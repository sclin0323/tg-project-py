"""
generate_ppt.py
從 WBS Excel 資料產生 JSON，再呼叫 generate_ppt.js 產出 PPT。
執行: python generate_ppt.py
"""

import os
import json
import subprocess
import tempfile
import pandas as pd
import openpyxl
from datetime import datetime

# ── 設定（與 main.py 保持一致）──────────────────────────
EXCEL_PATH    = "./excel/專案管理總表SDC.xlsx"
SNAPSHOT_ROOT = "./excel/snapshots"
PPT_JS        = os.path.join(os.path.dirname(__file__), "generate_ppt.js")


# ── 工具函數 ────────────────────────────────────────────

def fmt_date(val) -> str:
    if val is None or str(val) in ("NaT", "nan", "None", ""):
        return ""
    try:
        if isinstance(val, str):
            return val
        return pd.to_datetime(val).strftime("%Y/%m/%d")
    except Exception:
        return str(val)


def extract_wbs_links(path: str) -> dict:
    wb = openpyxl.load_workbook(path)
    ws = wb["WBS清單總表"]
    links = {}
    for row in ws.iter_rows(min_row=2):
        project_name = row[1].value
        wbs_cell     = row[6]
        if not project_name:
            continue
        if wbs_cell.hyperlink:
            src_path = wbs_cell.hyperlink.target.replace("\\", "/")
            links[project_name] = src_path
    return links


def read_master(path: str) -> pd.DataFrame:
    df = pd.read_excel(path, sheet_name="WBS清單總表").dropna(how="all")
    return df


def read_wbs(path: str) -> pd.DataFrame | None:
    if not os.path.exists(path):
        print(f"  ⚠️  找不到 WBS：{path}")
        return None
    try:
        return pd.read_excel(path, sheet_name=0)
    except Exception as e:
        print(f"  ❌ 讀取失敗：{e}")
        return None


def is_group_row(row: pd.Series) -> bool:
    """判斷是否為群組標題列（任務類型為空 & 工作項目含【】）"""
    task_type = row.get("任務類型", "")
    task_name = str(row.get("工作項目", "") or "")
    return (pd.isna(task_type) or task_type == "") and "【" in task_name


def build_project_data(master_row: pd.Series, wbs_df: pd.DataFrame) -> dict:
    tasks_raw = wbs_df.to_dict("records")
    tasks = []
    today = datetime.today()

    for t in tasks_raw:
        task_name = str(t.get("工作項目") or "")
        if not task_name.strip():
            continue

        group = is_group_row(pd.Series(t))
        status = str(t.get("狀態") or "")
        end_str = fmt_date(t.get("預計結束日"))
        actual_str = fmt_date(t.get("實際完成日"))

        # 自動偵測逾期
        if not group and status not in ("完成", "延遲", "風險") and end_str:
            try:
                end_dt = datetime.strptime(end_str, "%Y/%m/%d")
                if end_dt < today and not actual_str:
                    status = "延遲"
            except Exception:
                pass

        tasks.append({
            "wbsNo":      str(t.get("WBS 編號") or ""),
            "name":       task_name,
            "taskType":   str(t.get("任務類型") or ""),
            "owner":      str(t.get("負責人") or ""),
            "startDate":  fmt_date(t.get("開始日")),
            "endDate":    fmt_date(t.get("預計結束日")),
            "actualDate": fmt_date(t.get("實際完成日")),
            "status":     status,
            "risk":       str(t.get("風險等級") or ""),
            "note":       str(t.get("備註") or t.get("註記") or ""),
            "isGroup":    group,
        })

    non_group = [t for t in tasks if not t["isGroup"]]
    done       = [t for t in non_group if t["status"] == "完成"]
    in_prog    = [t for t in non_group if t["status"] == "進行中"]
    at_risk    = [t for t in non_group if t["status"] in ("延遲", "風險")]
    not_start  = [t for t in non_group if t["status"] == "待開始"]

    # 取最早開始日 / 最晚結束日 for 甘特圖
    dated = [t for t in non_group if t["startDate"]]
    start_dates = [t["startDate"] for t in dated if t["startDate"]]
    end_dates   = [t["endDate"]   for t in dated if t["endDate"]]

    return {
        "name":           str(master_row.get("專案名稱", "")),
        "owner":          str(master_row.get("負責人", "")),
        "stage":          str(master_row.get("階段目標", "")),
        "status":         str(master_row.get("狀態", "")),
        "startDate":      fmt_date(master_row.get("開始日期")),
        "targetDate":     fmt_date(master_row.get("目標日期")),
        "totalTasks":     len(non_group),
        "doneTasks":      len(done),
        "inProgressTasks":len(in_prog),
        "riskTasks":      len(at_risk),
        "notStartedTasks":len(not_start),
        "ganttStart":     min(start_dates) if start_dates else "",
        "ganttEnd":       max(end_dates)   if end_dates   else "",
        "tasks":          tasks,
    }


def collect_risks(projects: list) -> list:
    risks = []
    for proj in projects:
        for t in proj["tasks"]:
            if not t["isGroup"] and t["status"] in ("延遲", "風險"):
                risks.append({
                    "project": proj["name"],
                    "wbsNo":   t["wbsNo"],
                    "name":    t["name"],
                    "endDate": t["endDate"],
                    "status":  t["status"],
                    "owner":   t["owner"],
                    "note":    t["note"],
                })
    return risks


def get_latest_snapshot_dir() -> str | None:
    """找到最新的 snapshot 目錄"""
    if not os.path.exists(SNAPSHOT_ROOT):
        return None
    dirs = sorted([
        d for d in os.listdir(SNAPSHOT_ROOT)
        if os.path.isdir(os.path.join(SNAPSHOT_ROOT, d))
    ], reverse=True)
    return os.path.join(SNAPSHOT_ROOT, dirs[0]) if dirs else None


# ── 主流程 ────────────────────────────────────────────

def main():
    print("=== 📊 讀取專案資料 ===")
    df_master  = read_master(EXCEL_PATH)
    wbs_links  = extract_wbs_links(EXCEL_PATH)

    projects = []
    for _, master_row in df_master.iterrows():
        proj_name = str(master_row.get("專案名稱", "")).strip()
        if not proj_name:
            continue
        wbs_path = wbs_links.get(proj_name)
        if not wbs_path:
            print(f"  ⚠️  找不到 {proj_name} 的 WBS 連結")
            continue
        print(f"  ▶ {proj_name}")
        wbs_df = read_wbs(wbs_path)
        if wbs_df is not None:
            proj_data = build_project_data(master_row, wbs_df)
            projects.append(proj_data)

    risks = collect_risks(projects)
    total_tasks = sum(p["totalTasks"] for p in projects)

    payload = {
        "reportDate": datetime.today().strftime("%Y/%m/%d"),
        "stats": {
            "totalProjects":    len(projects),
            "totalTasks":       total_tasks,
            "doneTasks":        sum(p["doneTasks"]       for p in projects),
            "inProgressTasks":  sum(p["inProgressTasks"] for p in projects),
            "riskTasks":        sum(p["riskTasks"]       for p in projects),
        },
        "projects": projects,
        "risks":    risks,
    }

    # ── 決定輸出路徑（放在最新 snapshot 目錄，或當前目錄）
    snap_dir = get_latest_snapshot_dir()
    if snap_dir:
        date_tag = datetime.today().strftime("%Y%m%d")
        out_ppt  = os.path.join(snap_dir, f"專案報告_{date_tag}.pptx")
    else:
        out_ppt  = f"專案報告_{datetime.today().strftime('%Y%m%d')}.pptx"

    # ── 寫 JSON 暫存
    with tempfile.NamedTemporaryFile(mode="w", suffix=".json",
                                     delete=False, encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)
        json_path = f.name

    # ── 呼叫 JS 產 PPT
    print(f"\n=== 🖥️  產生 PPT ===")
    result = subprocess.run(
        ["node", PPT_JS, json_path, out_ppt],
        capture_output=True, text=True
    )
    os.unlink(json_path)

    if result.returncode == 0:
        print(f"  {result.stdout.strip()}")
        print(f"\n✅  PPT 已儲存：{out_ppt}")
    else:
        print(f"  ❌ 錯誤：{result.stderr}")

    return out_ppt


if __name__ == "__main__":
    main()