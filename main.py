import os
import json
import shutil
import subprocess
import tempfile
import pandas as pd
import openpyxl
from datetime import datetime, timedelta

EXCEL_PATH    = "./excel/專案管理總表SDC.xlsx"
SNAPSHOT_ROOT = "./excel/snapshots"
PPT_JS        = os.path.join(os.path.dirname(os.path.abspath(__file__)), "generate_ppt.js")

LOOKBACK_DAYS = 14   # 本期回顧視窗
LOOKAHEAD_DAYS = 14  # 下期前瞻視窗


# ── 基礎工具 ────────────────────────────────────────────

def get_next_snapshot_dir() -> str:
    os.makedirs(SNAPSHOT_ROOT, exist_ok=True)
    today = datetime.today().strftime("%Y%m%d")
    seq = 1
    while True:
        path = os.path.join(SNAPSHOT_ROOT, f"{today}_{seq:03d}")
        if not os.path.exists(path):
            os.makedirs(path)
            return path
        seq += 1


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
            links[project_name] = wbs_cell.hyperlink.target.replace("\\", "/")
    return links


def read_master(path: str) -> pd.DataFrame:
    df = pd.read_excel(path, sheet_name="WBS清單總表").dropna(how="all")
    for col in ["開始日期", "目標日期"]:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors="coerce").dt.strftime("%Y/%m/%d")
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
        print(f"  ❌ 讀取失敗：{e}")
        return None


def copy_wbs_to_snapshot(wbs_links: dict, snapshot_dir: str):
    for project, src_path in wbs_links.items():
        if not os.path.exists(src_path):
            print(f"  ⚠️  來源不存在：{src_path}")
            continue
        dest_path = os.path.join(snapshot_dir, os.path.basename(src_path))
        shutil.copy2(src_path, dest_path)
        print(f"  📄 [{project}]  →  {dest_path}")


def fmt_date(val) -> str:
    if val is None or str(val) in ("NaT", "nan", "None", ""):
        return ""
    try:
        return val if isinstance(val, str) else pd.to_datetime(val).strftime("%Y/%m/%d")
    except Exception:
        return str(val)


def parse_date(s: str) -> datetime | None:
    if not s:
        return None
    try:
        return datetime.strptime(s, "%Y/%m/%d")
    except Exception:
        return None


def is_group_row(row: dict) -> bool:
    """群組標題列：任務類型空 + 工作項目含【】"""
    task_type = row.get("任務類型", "")
    task_name = str(row.get("工作項目", "") or "")
    return (pd.isna(task_type) or task_type == "") and "【" in task_name


# ══════════════════════════════════════════════════════════
# 高階分析邏輯
# ══════════════════════════════════════════════════════════

def compute_health(proj_start: datetime, proj_target: datetime,
                   tasks: list, today: datetime) -> dict:
    """計算專案健康度（紅黃綠 + 進度差 + 預估完工日）"""
    non_group = [t for t in tasks if not t["isGroup"]]
    if not non_group or not proj_start or not proj_target:
        return {"level": "GREEN", "label": "正常", "drift": 0,
                "elapsedPct": 0, "donePct": 0,
                "forecastEnd": "", "forecastDrift": 0}

    total = len(non_group)
    done  = sum(1 for t in non_group if t["status"] == "完成")
    risk  = sum(1 for t in non_group if t["status"] in ("延遲", "風險"))

    # ── 進度差 (時間 vs 完成)
    total_days   = max(1, (proj_target - proj_start).days)
    elapsed_days = max(0, (today - proj_start).days)
    elapsed_pct  = min(1.0, elapsed_days / total_days)
    done_pct     = done / total

    drift = done_pct - elapsed_pct  # 負值 = 落後

    if drift >= -0.05:    level = "GREEN";  label = "健康"
    elif drift >= -0.15:  level = "AMBER";  label = "注意"
    else:                 level = "RED";    label = "警報"

    # 風險任務存在 → 自動降一級
    if risk > 0:
        if level == "GREEN":  level, label = "AMBER", "注意"
        elif level == "AMBER": level, label = "RED", "警報"

    # ── 預估完工日 (依當前 velocity 推算，前期加阻尼)
    forecast_end_str = ""
    forecast_drift = 0
    if done > 0 and elapsed_days > 0 and elapsed_pct >= 0.15:
        # 已過 15% 以上才信任 velocity
        velocity_days_per_task = elapsed_days / done
        remaining_tasks = total - done
        remaining_days = remaining_tasks * velocity_days_per_task
        forecast_end = today + timedelta(days=remaining_days)
        forecast_end_str = forecast_end.strftime("%Y/%m/%d")
        forecast_drift = (forecast_end - proj_target).days
        # cap 在 ±50% 專案總時間內
        max_drift = int(total_days * 0.5)
        forecast_drift = max(-max_drift, min(max_drift, forecast_drift))
    else:
        # 前期：用目標日期當預估
        forecast_end_str = proj_target.strftime("%Y/%m/%d")
        forecast_drift = 0

    return {
        "level":         level,
        "label":         label,
        "drift":         round(drift * 100, 1),
        "elapsedPct":    round(elapsed_pct * 100, 1),
        "donePct":       round(done_pct * 100, 1),
        "forecastEnd":   forecast_end_str,
        "forecastDrift": forecast_drift,   # 正 = 落後 N 天，負 = 提早
    }


def extract_milestones(tasks: list) -> list:
    """從群組列(【...】)萃取里程碑：每個 Phase 的起訖、進度、狀態"""
    milestones = []
    current = None
    bucket  = []

    for t in tasks:
        if t["isGroup"]:
            # 收尾上一個
            if current and bucket:
                _close_milestone(current, bucket, milestones)
            current = {"name": t["name"].replace("【", "").replace("】", "")}
            bucket  = []
        else:
            if current is not None:
                bucket.append(t)

    if current and bucket:
        _close_milestone(current, bucket, milestones)

    return milestones


def _close_milestone(meta: dict, children: list, out: list):
    starts = [parse_date(t["startDate"]) for t in children if t["startDate"]]
    ends   = [parse_date(t["endDate"])   for t in children if t["endDate"]]
    starts = [d for d in starts if d]
    ends   = [d for d in ends   if d]
    if not starts or not ends:
        return

    total = len(children)
    done  = sum(1 for t in children if t["status"] == "完成")
    risk  = sum(1 for t in children if t["status"] in ("延遲", "風險"))
    in_prog = sum(1 for t in children if t["status"] == "進行中")

    if done == total:        status = "完成"
    elif risk > 0:           status = "風險"
    elif in_prog > 0:        status = "進行中"
    else:                    status = "待開始"

    out.append({
        "name":      meta["name"].strip(),
        "startDate": min(starts).strftime("%Y/%m/%d"),
        "endDate":   max(ends).strftime("%Y/%m/%d"),
        "totalTasks": total,
        "doneTasks":  done,
        "progress":   round(done / total * 100) if total else 0,
        "status":     status,
    })


def collect_period_activity(tasks: list, today: datetime) -> dict:
    """彙整本期完成 + 進行中與即將到期任務"""
    recent_done = []
    upcoming    = []
    seen        = set()

    lookback  = today - timedelta(days=LOOKBACK_DAYS)
    lookahead = today + timedelta(days=LOOKAHEAD_DAYS)

    for t in tasks:
        if t["isGroup"]:
            continue

        # 本期完成
        actual = parse_date(t["actualDate"])
        if actual and actual >= lookback and actual <= today:
            recent_done.append({
                "wbsNo": t["wbsNo"], "name": t["name"],
                "owner": t["owner"], "date": t["actualDate"]
            })

        # 即將到期 / 進行中（兩種都列入）
        end = parse_date(t["endDate"])
        is_upcoming    = end and t["status"] != "完成" and today <= end <= lookahead
        is_in_progress = t["status"] in ("進行中", "風險", "延遲")

        if (is_upcoming or is_in_progress) and t["wbsNo"] not in seen:
            seen.add(t["wbsNo"])
            upcoming.append({
                "wbsNo": t["wbsNo"], "name": t["name"],
                "owner": t["owner"], "date": t["endDate"],
                "status": t["status"]
            })

    recent_done.sort(key=lambda x: x["date"], reverse=True)
    upcoming.sort(key=lambda x: x["date"] or "9999")

    return {
        "recentDone": recent_done[:6],
        "upcoming":   upcoming[:6],
    }


def derive_decisions(projects: list, today: datetime) -> list:
    """把風險彙整成主管需要決策的事項（含影響、建議）"""
    decisions = []

    for proj in projects:
        for t in proj["tasks"]:
            if t["isGroup"]:
                continue

            # 只篩風險/延遲
            if t["status"] not in ("延遲", "風險"):
                continue

            end = parse_date(t["endDate"])
            actual = parse_date(t["actualDate"])
            days_late = (today - end).days if (end and not actual) else 0

            # 嚴重度判斷
            if days_late > 14 or t["risk"] == "高":
                severity = "高"
            elif days_late > 3 or t["status"] == "風險":
                severity = "中"
            else:
                severity = "低"

            # 議題描述
            if days_late > 0:
                issue = f"{t['wbsNo']} {t['name']}　延遲 {days_late} 天"
            else:
                issue = f"{t['wbsNo']} {t['name']}　標記為{t['status']}"

            decisions.append({
                "severity": severity,
                "project":  proj["name"],
                "issue":    issue,
                "impact":   t["note"] or "可能影響後續任務啟動",
                "owner":    t["owner"],
                "ask":      _suggest_ask(t, days_late),
                "daysLate": days_late,
            })

    sev_order = {"高": 0, "中": 1, "低": 2}
    decisions.sort(key=lambda d: (sev_order.get(d["severity"], 3), -d["daysLate"]))
    return decisions


def _suggest_ask(task: dict, days_late: int) -> str:
    if task["risk"] == "高":
        return "需裁示：是否加派人力 / 縮減範圍"
    if days_late > 14:
        return "需裁示：是否調整目標日"
    if days_late > 7:
        return "需協助排除阻礙"
    if days_late > 0:
        return "請關注進度"
    return "請關注"


# ══════════════════════════════════════════════════════════
# 組裝專案 payload
# ══════════════════════════════════════════════════════════

def build_project_payload(master_row: pd.Series, wbs_df: pd.DataFrame,
                          today: datetime) -> dict:
    tasks = []
    for t in wbs_df.to_dict("records"):
        task_name = str(t.get("工作項目") or "").strip()
        if not task_name:
            continue
        group  = is_group_row(t)
        status = str(t.get("狀態") or "")
        end_str = fmt_date(t.get("預計結束日"))
        actual_str = fmt_date(t.get("實際完成日"))

        # 自動偵測逾期
        if not group and status not in ("完成", "延遲", "風險") and end_str:
            ed = parse_date(end_str)
            if ed and ed < today and not actual_str:
                status = "延遲"

        tasks.append({
            "wbsNo":      str(t.get("WBS 編號") or ""),
            "name":       task_name,
            "taskType":   str(t.get("任務類型") or ""),
            "owner":      str(t.get("負責人") or ""),
            "startDate":  fmt_date(t.get("開始日")),
            "endDate":    end_str,
            "actualDate": actual_str,
            "status":     status,
            "risk":       str(t.get("風險等級") or ""),
            "note":       str(t.get("備註") or t.get("註記") or ""),
            "isGroup":    group,
        })

    non_group = [t for t in tasks if not t["isGroup"]]
    proj_start  = parse_date(fmt_date(master_row.get("開始日期")))
    proj_target = parse_date(fmt_date(master_row.get("目標日期")))

    health     = compute_health(proj_start, proj_target, tasks, today)
    milestones = extract_milestones(tasks)
    activity   = collect_period_activity(tasks, today)

    return {
        "name":            str(master_row.get("專案名稱", "")),
        "owner":           str(master_row.get("負責人", "")),
        "stage":           str(master_row.get("階段目標", "")),
        "status":          str(master_row.get("狀態", "")),
        "startDate":       fmt_date(master_row.get("開始日期")),
        "targetDate":      fmt_date(master_row.get("目標日期")),
        "totalTasks":      len(non_group),
        "doneTasks":       sum(1 for t in non_group if t["status"] == "完成"),
        "inProgressTasks": sum(1 for t in non_group if t["status"] == "進行中"),
        "riskTasks":       sum(1 for t in non_group if t["status"] in ("延遲", "風險")),
        "notStartedTasks": sum(1 for t in non_group if t["status"] == "待開始"),
        "health":          health,
        "milestones":      milestones,
        "activity":        activity,
        "tasks":           tasks,
    }


def generate_ppt(projects: list, decisions: list, snapshot_dir: str) -> str | None:
    if not os.path.exists(PPT_JS):
        print(f"  ⚠️  找不到 {PPT_JS}")
        return None

    # 整體 headline
    red   = sum(1 for p in projects if p["health"]["level"] == "RED")
    amber = sum(1 for p in projects if p["health"]["level"] == "AMBER")
    green = sum(1 for p in projects if p["health"]["level"] == "GREEN")

    if red > 0:
        headline = f"⚠ {red} 個專案需立即關注"
        overall  = "RED"
    elif amber > 0:
        headline = f"{amber} 個專案需注意觀察"
        overall  = "AMBER"
    else:
        headline = "全部專案進度健康"
        overall  = "GREEN"

    payload = {
        "reportDate": datetime.today().strftime("%Y/%m/%d"),
        "nextReport": (datetime.today() + timedelta(days=7)).strftime("%Y/%m/%d"),
        "headline":   headline,
        "overallHealth": overall,
        "stats": {
            "totalProjects":   len(projects),
            "totalTasks":      sum(p["totalTasks"]      for p in projects),
            "doneTasks":       sum(p["doneTasks"]       for p in projects),
            "inProgressTasks": sum(p["inProgressTasks"] for p in projects),
            "riskTasks":       sum(p["riskTasks"]       for p in projects),
            "redProjects":     red,
            "amberProjects":   amber,
            "greenProjects":   green,
            "decisionCount":   sum(1 for d in decisions if d["severity"] in ("高", "中")),
        },
        "projects":  projects,
        "decisions": decisions,
    }

    out_ppt = os.path.join(snapshot_dir,
                           f"專案報告_{datetime.today().strftime('%Y%m%d')}.pptx")

    with tempfile.NamedTemporaryFile(mode="w", suffix=".json",
                                     delete=False, encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)
        json_path = f.name

    result = subprocess.run(["node", PPT_JS, json_path, out_ppt],
                            capture_output=True, text=True, encoding="utf-8")
    os.unlink(json_path)

    if result.returncode == 0:
        print(f"  ✅ PPT：{out_ppt}")
        return out_ppt
    else:
        print(f"  ❌ 失敗：{result.stderr.strip()}")
        return None


# ══════════════════════════════════════════════════════════
# 主流程
# ══════════════════════════════════════════════════════════

today = datetime.today()

print("=== 📋 專案管理總表 ===")
df_master = read_master(EXCEL_PATH)
print(df_master.to_string(index=False))

print("\n=== 🔗 讀取 WBS 子檔案 ===")
wbs_links    = extract_wbs_links(EXCEL_PATH)
ppt_projects = []

for _, master_row in df_master.iterrows():
    project  = str(master_row.get("專案名稱", "")).strip()
    wbs_path = wbs_links.get(project)
    if not project or not wbs_path:
        continue
    print(f"\n▶ {project}  →  {wbs_path}")
    df_wbs = read_wbs(wbs_path, project)
    if df_wbs is not None:
        print(f"  ✅ 共 {len(df_wbs)} 筆任務")
        ppt_projects.append(build_project_payload(master_row, df_wbs, today))

# 健康度報告
print("\n=== 🚦 專案健康度 ===")
for p in ppt_projects:
    h = p["health"]
    icon = {"RED": "🔴", "AMBER": "🟡", "GREEN": "🟢"}[h["level"]]
    drift_txt = f"{h['forecastDrift']:+d} 天" if h["forecastEnd"] else "—"
    print(f"  {icon} {p['name']}：{h['label']}  完成 {h['donePct']}%  預估{drift_txt}")

# 建立快照目錄並複製 WBS 檔案
print("\n=== 📁 建立快照 ===")
snapshot_dir = get_next_snapshot_dir()
print(f"  📂 {snapshot_dir}")
copy_wbs_to_snapshot(wbs_links, snapshot_dir)

# 決策需求
decisions = derive_decisions(ppt_projects, today)
print(f"\n=== ⚠️  需主管裁示事項：{len(decisions)} 件 ===")
for d in decisions[:5]:
    print(f"  [{d['severity']}] {d['project']}: {d['issue']}")
    print(f"      → {d['ask']}")

# 產生 PPT
print("\n=== 🖥️  產生 PPT ===")
if ppt_projects:
    generate_ppt(ppt_projects, decisions, snapshot_dir)

print(f"\n✅ 完成！快照位置：{snapshot_dir}")