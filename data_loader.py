"""
data_loader.py — 從 Excel 載入並計算 PPT 需要的所有衍生資料

回傳的結構 (report):
{
  "report_date":  date,
  "next_report":  date,
  "headline":     str,              # 一句總結
  "overall":      "RED|AMBER|GREEN",
  "stats": {
    "total_projects", "red_projects", "amber_projects", "green_projects",
    "total_tasks", "done_tasks", "in_progress_tasks", "risk_tasks",
    "decision_count",
  },
  "projects": [ project_obj, ... ],
  "decisions": [ decision, ... ],   # 跨專案排序好
}

project_obj: {
  "name", "owner", "stage", "status", "start", "target",
  "total", "done", "in_progress", "risk", "not_started",
  "health":   { "level", "label", "drift_pct", "done_pct", "elapsed_pct",
                "forecast_end", "forecast_drift" },
  "summary":  str,        # 一句話狀況描述（給主管 3 秒看的）
  "milestones": [...],
  "activity":   { "recent_done": [...], "upcoming": [...] },
  "decisions":  [...],
}
"""
from __future__ import annotations
import os
from datetime import datetime, timedelta, date
import openpyxl
import pandas as pd

# ── 設定 ─────────────────────────────────────────────
LOOKBACK_DAYS  = 14
LOOKAHEAD_DAYS = 14

STATUS_DONE   = "完成"
STATUS_ACTIVE = ("進行中",)
STATUS_RISK   = ("延遲", "風險")
STATUS_SKIP   = ("取消", "暫緩執行")
STATUS_WAIT   = "待開始"
RISK_HIGH = "高"
RISK_MID  = "中"

HEALTH_LABELS = {"RED": "警報", "AMBER": "注意", "GREEN": "健康"}


# ──────────────────────────────────────────────────────
# 基礎工具
# ──────────────────────────────────────────────────────
def _parse(v):
    """把任意值 (datetime/str/NaT) 轉成 date 或 None"""
    if v is None or (hasattr(v, '__class__') and v.__class__.__name__ == 'NaTType'):
        return None
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    if isinstance(v, str):
        v = v.strip()
        if not v or v in ("NaT", "nan", "None"):
            return None
        for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%m/%d/%Y"):
            try:
                return datetime.strptime(v, fmt).date()
            except ValueError:
                continue
    try:
        ts = pd.to_datetime(v, errors="coerce")
        if pd.isna(ts):
            return None
        return ts.date()
    except Exception:
        return None


def _str(v) -> str:
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    return str(v).strip()


def _is_group(name: str, task_type: str) -> bool:
    return (not task_type) and ("【" in (name or ""))


# ──────────────────────────────────────────────────────
# 載入資料
# ──────────────────────────────────────────────────────
def load_master(excel_dir: str) -> list[dict]:
    """讀總表 WBS清單總表"""
    path = os.path.join(excel_dir, "專案管理總表SDC.xlsx")
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb["WBS清單總表"]
    out = []
    for r in range(2, ws.max_row + 1):
        name = _str(ws.cell(row=r, column=2).value)
        if not name:
            continue
        out.append({
            "start":    _parse(ws.cell(row=r, column=1).value),
            "name":     name,
            "stage":    _str(ws.cell(row=r, column=3).value),
            "owner":    _str(ws.cell(row=r, column=4).value),
            "target":   _parse(ws.cell(row=r, column=5).value),
            "status":   _str(ws.cell(row=r, column=6).value),
            "wbs_file": _str(ws.cell(row=r, column=7).value),
        })
    return out


def load_wbs(excel_dir: str, wbs_file: str) -> list[dict]:
    """讀單個 WBS Excel；總表 G 欄的檔名是 'X.xlsx'，實體是 'WBS_X.xlsx'"""
    # 先嘗試實體檔名（WBS_ 前綴），找不到再試原檔名
    candidates = [f"WBS_{wbs_file}", wbs_file]
    src = None
    for c in candidates:
        p = os.path.join(excel_dir, c)
        if os.path.exists(p):
            src = p
            break
    if not src:
        raise FileNotFoundError(f"找不到 WBS 檔案: {wbs_file}")

    wb = openpyxl.load_workbook(src, data_only=True)
    ws = wb["WBS"]
    tasks = []
    for r in range(2, ws.max_row + 1):
        row = [ws.cell(row=r, column=c).value for c in range(1, 13)]
        wbs_no, t_type, work, deliv, owner, start, end, actual, pred, status, risk, note = row

        name = _str(work)
        if not name:
            continue

        tt   = _str(t_type)
        is_g = _is_group(name, tt)

        tasks.append({
            "wbs_no":     _str(wbs_no),
            "task_type":  tt,
            "name":       name,
            "deliv":      _str(deliv),
            "owner":      _str(owner),
            "start":      _parse(start),
            "end":        _parse(end),
            "actual":     _parse(actual),
            "pred":       _str(pred),
            "status":     _str(status),
            "risk":       _str(risk),
            "note":       _str(note),
            "is_group":   is_g,
        })
    return tasks


# ──────────────────────────────────────────────────────
# 健康度
# ──────────────────────────────────────────────────────
def compute_health(start: date, target: date, tasks: list[dict],
                   today: date) -> dict:
    real = [t for t in tasks if not t["is_group"] and t["status"] not in STATUS_SKIP]
    if not real or not start or not target:
        return {
            "level": "GREEN", "label": "健康",
            "done_pct": 0, "elapsed_pct": 0, "drift_pct": 0,
            "forecast_end": target, "forecast_drift": 0,
            "forecast_available": False,
        }

    total = len(real)
    done  = sum(1 for t in real if t["status"] == STATUS_DONE)
    risk  = sum(1 for t in real if t["status"] in STATUS_RISK)
    high  = sum(1 for t in real if t["status"] in STATUS_RISK and t["risk"] == RISK_HIGH)

    total_days = max(1, (target - start).days)
    elapsed    = max(0, min(total_days, (today - start).days))
    elapsed_pct = elapsed / total_days
    done_pct    = done / total
    drift       = done_pct - elapsed_pct

    # 進度差 → 燈號
    if   drift >= -0.05: level = "GREEN"
    elif drift >= -0.15: level = "AMBER"
    else:                level = "RED"

    # 風險加成（升級規則，單向不降級）
    #   1) 有任何「高」風險任務          → 直升 RED
    #   2) AMBER 狀態下風險任務 ≥ 3      → 升 RED
    #   3) GREEN 狀態下風險任務 ≥ 2      → 升 AMBER
    #      （只有 1 個一般風險視為可控雜訊，不影響整體燈號）
    if high > 0:
        level = "RED"
    elif risk >= 3 and level == "AMBER":
        level = "RED"
    elif risk >= 2 and level == "GREEN":
        level = "AMBER"

    # 預估完工
    forecast_end = target
    forecast_drift = 0
    forecast_available = False
    if done > 0 and elapsed >= 7 and done_pct >= 0.10:
        velocity = elapsed / done
        remaining = (total - done) * velocity
        forecast_end = today + timedelta(days=int(remaining))
        forecast_drift = (forecast_end - target).days
        cap = 60
        if forecast_drift > cap:
            forecast_drift, forecast_end = cap, target + timedelta(days=cap)
        elif forecast_drift < -cap:
            forecast_drift, forecast_end = -cap, target - timedelta(days=cap)
        forecast_available = True

    return {
        "level":              level,
        "label":              HEALTH_LABELS[level],
        "done_pct":           round(done_pct * 100, 1),
        "elapsed_pct":        round(elapsed_pct * 100, 1),
        "drift_pct":          round(drift * 100, 1),
        "forecast_end":       forecast_end,
        "forecast_drift":     forecast_drift,
        "forecast_available": forecast_available,
    }


# ──────────────────────────────────────────────────────
# 里程碑
# ──────────────────────────────────────────────────────
def extract_milestones(tasks: list[dict]) -> list[dict]:
    out, cur_name, bucket = [], None, []

    def close():
        if not cur_name or not bucket:
            return
        kids = [t for t in bucket if t["status"] not in STATUS_SKIP]
        starts = [t["start"] for t in kids if t["start"]]
        ends   = [t["end"]   for t in kids if t["end"]]
        if not starts or not ends:
            return
        total = len(kids)
        done  = sum(1 for t in kids if t["status"] == STATUS_DONE)
        risk  = sum(1 for t in kids if t["status"] in STATUS_RISK)
        inp   = sum(1 for t in kids if t["status"] in STATUS_ACTIVE)
        if done == total: st = "完成"
        elif risk > 0:    st = "風險"
        elif inp > 0:     st = "進行中"
        else:             st = "待開始"
        # 里程碑進度演算法（Earned Value 簡化版）：
        #   完成 = 1.0 個
        #   進行中 = 0.5 個（已啟動、正常推進中）
        #   風險 / 延遲 = 0   （已啟動但偏離計畫，不視為「半完成」避免誤導）
        #   待開始 = 0
        progress_count = done + inp * 0.5
        progress = round(progress_count / total * 100) if total else 0
        out.append({
            "name":   cur_name.replace("【", "").replace("】", "").strip(),
            "start":  min(starts),
            "end":    max(ends),
            "total":  total,
            "done":   done,
            "progress": progress,
            "status": st,
        })

    for t in tasks:
        if t["is_group"]:
            close()
            cur_name = t["name"]
            bucket = []
        elif cur_name is not None:
            bucket.append(t)
    close()
    return out


# ──────────────────────────────────────────────────────
# 本期動態
# ──────────────────────────────────────────────────────
def collect_activity(tasks: list[dict], today: date) -> dict:
    lookback  = today - timedelta(days=LOOKBACK_DAYS)
    lookahead = today + timedelta(days=LOOKAHEAD_DAYS)
    recent_done, upcoming, seen = [], [], set()

    for t in tasks:
        if t["is_group"] or t["status"] in STATUS_SKIP:
            continue

        if t["actual"] and lookback <= t["actual"] <= today:
            recent_done.append(t)

        is_up      = t["end"] and t["status"] != STATUS_DONE and today <= t["end"] <= lookahead
        is_running = t["status"] in (*STATUS_ACTIVE, *STATUS_RISK)
        if (is_up or is_running) and t["wbs_no"] not in seen:
            seen.add(t["wbs_no"])
            upcoming.append(t)

    recent_done.sort(key=lambda x: x["actual"], reverse=True)

    def _sort_key(x):
        end = x["end"] or date.max
        sev = 0 if x["status"] in STATUS_RISK else 1
        return (sev, end)
    upcoming.sort(key=_sort_key)

    return {"recent_done": recent_done[:8], "upcoming": upcoming[:8]}


# ──────────────────────────────────────────────────────
# 裁示事項
# ──────────────────────────────────────────────────────
def derive_decisions(project_name: str, tasks: list[dict],
                     today: date) -> list[dict]:
    out = []
    for t in tasks:
        if t["is_group"] or t["status"] not in STATUS_RISK:
            continue
        end = t["end"]
        days_late = (today - end).days if (end and not t["actual"]) else 0

        # 嚴重度
        if t["risk"] == RISK_HIGH or days_late > 21:
            sev = "高"
        elif t["risk"] == RISK_MID or days_late > 7 or t["status"] == "風險":
            sev = "中"
        else:
            sev = "低"

        # 議題敘述
        if days_late > 0:
            issue = f"{t['wbs_no']} {t['name']}（已逾期 {days_late} 天）"
        else:
            issue = f"{t['wbs_no']} {t['name']}（標記為{t['status']}）"

        # 建議行動
        if sev == "高":
            ask = "需裁示：加派人力 / 縮減範圍 / 調整上線時程"
        elif sev == "中":
            ask = "需協助：釐清阻礙因素，評估是否調整計畫"
        elif days_late > 7:
            ask = "需協助排除阻礙"
        else:
            ask = "請關注進度"

        out.append({
            "severity":  sev,
            "project":   project_name,
            "wbs_no":    t["wbs_no"],
            "name":      t["name"],
            "issue":     issue,
            "impact":    t["note"] or "可能影響後續任務啟動",
            "owner":     t["owner"],
            "ask":       ask,
            "days_late": days_late,
        })

    sev_order = {"高": 0, "中": 1, "低": 2}
    out.sort(key=lambda d: (sev_order[d["severity"]], -d["days_late"]))
    return out


# ──────────────────────────────────────────────────────
# 一句話總結
# ──────────────────────────────────────────────────────
def make_summary(p: dict) -> str:
    h = p["health"]
    n_risk = p["risk"]
    drift = h["drift_pct"]
    if h["level"] == "RED":
        if n_risk >= 3:
            return f"{n_risk} 個任務出現嚴重風險，影響整體上線時程，需立即裁示"
        if n_risk > 0:
            return f"進度落後 {abs(round(drift))}%，{n_risk} 個風險任務急需處理"
        return f"進度嚴重落後 {abs(round(drift))}%，需採取補救措施"
    if h["level"] == "AMBER":
        if n_risk > 0:
            return f"進度大致符合計畫，{n_risk} 個任務出現風險訊號，需持續關注"
        return f"進度略落後 {abs(round(drift))}%，仍可追回"
    if h["done_pct"] > h["elapsed_pct"] + 5:
        return f"進度超前約 {round(h['done_pct'] - h['elapsed_pct'])}%，依計畫順利推進"
    return "進度依計畫推進，整體狀況良好"


# ──────────────────────────────────────────────────────
# 主流程
# ──────────────────────────────────────────────────────
def build_report(excel_dir: str, today: date | None = None) -> dict:
    today = today or date.today()
    master = load_master(excel_dir)

    projects = []
    all_decisions = []

    for m in master:
        try:
            tasks = load_wbs(excel_dir, m["wbs_file"])
        except FileNotFoundError as e:
            print(f"  ⚠  {e}")
            continue

        # 自動偵測逾期
        for t in tasks:
            if (not t["is_group"] and t["status"] not in
                    (STATUS_DONE, *STATUS_RISK, *STATUS_SKIP)
                    and t["end"] and t["end"] < today and not t["actual"]):
                t["status"] = "延遲"

        real = [t for t in tasks if not t["is_group"] and t["status"] not in STATUS_SKIP]
        total = len(real)
        done  = sum(1 for t in real if t["status"] == STATUS_DONE)
        inp   = sum(1 for t in real if t["status"] in STATUS_ACTIVE)
        risk  = sum(1 for t in real if t["status"] in STATUS_RISK)
        wait  = sum(1 for t in real if t["status"] == STATUS_WAIT)

        health = compute_health(m["start"], m["target"], tasks, today)
        milestones = extract_milestones(tasks)
        activity = collect_activity(tasks, today)
        decisions = derive_decisions(m["name"], tasks, today)
        all_decisions.extend(decisions)

        p = {
            "name":        m["name"],
            "owner":       m["owner"],
            "stage":       m["stage"],
            "status":      m["status"],
            "start":       m["start"],
            "target":      m["target"],
            "total":       total,
            "done":        done,
            "in_progress": inp,
            "risk":        risk,
            "not_started": wait,
            "health":      health,
            "milestones":  milestones,
            "activity":    activity,
            "decisions":   decisions,
            "tasks":       tasks,
        }
        p["summary"] = make_summary(p)
        projects.append(p)

    # 跨專案排序裁示事項
    sev_order = {"高": 0, "中": 1, "低": 2}
    all_decisions.sort(key=lambda d: (sev_order[d["severity"]], -d["days_late"]))

    red   = sum(1 for p in projects if p["health"]["level"] == "RED")
    amber = sum(1 for p in projects if p["health"]["level"] == "AMBER")
    green = sum(1 for p in projects if p["health"]["level"] == "GREEN")

    if red > 0:
        headline = f"{red} 個專案需立即關注"
        overall = "RED"
    elif amber > 0:
        headline = f"{amber} 個專案需注意觀察"
        overall = "AMBER"
    else:
        headline = "全部專案進度健康"
        overall = "GREEN"

    return {
        "report_date":  today,
        "next_report":  today + timedelta(days=7),
        "headline":     headline,
        "overall":      overall,
        "stats": {
            "total_projects":   len(projects),
            "red_projects":     red,
            "amber_projects":   amber,
            "green_projects":   green,
            "total_tasks":      sum(p["total"]       for p in projects),
            "done_tasks":       sum(p["done"]        for p in projects),
            "in_progress_tasks":sum(p["in_progress"] for p in projects),
            "risk_tasks":       sum(p["risk"]        for p in projects),
            "decision_count":   sum(1 for d in all_decisions if d["severity"] in ("高","中")),
        },
        "projects":  projects,
        "decisions": all_decisions,
    }


# ── 自我驗證 ──────────────────────────────────────────
if __name__ == "__main__":
    r = build_report("/home/claude/work/excel", today=date(2026, 5, 12))
    print(f"\n📊  整體：{r['overall']} | {r['headline']}")
    print(f"     報告日：{r['report_date']} | 下次回報：{r['next_report']}")
    print(f"     紅 {r['stats']['red_projects']} / 黃 {r['stats']['amber_projects']} / 綠 {r['stats']['green_projects']}")
    print(f"\n📋  各專案：")
    for p in r["projects"]:
        h = p["health"]
        icon = {"RED": "🔴", "AMBER": "🟡", "GREEN": "🟢"}[h["level"]]
        print(f"  {icon} {p['name']:<22} | {h['label']} | 完成 {h['done_pct']:>4}% / 時間 {h['elapsed_pct']:>4}% | 偏差 {h['forecast_drift']:+4}d")
        print(f"     → {p['summary']}")
        print(f"     里程碑 {len(p['milestones'])} | 近 {len(p['activity']['recent_done'])} 完成 | 即將到期 {len(p['activity']['upcoming'])} | 風險 {len(p['decisions'])}")
    print(f"\n⚠   跨專案裁示事項：{r['stats']['decision_count']} 件（高/中嚴重度）")
    for d in r["decisions"][:6]:
        print(f"   [{d['severity']}] {d['project']:<15} | {d['issue']}")