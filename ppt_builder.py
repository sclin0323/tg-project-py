"""
ppt_builder.py — 從 data_loader 取得的 report 物件產出 PPT
"""
from __future__ import annotations
from datetime import date, datetime
from pptx import Presentation
from pptx.util import Inches

from ppt_components import (
    C, HEALTH, STATUS_COLOR, FONT, SLIDE_W, SLIDE_H,
    rect, round_rect, circle, text, divider,
    page_header, kpi_card, progress_bar, status_pill, health_badge,
    fmt_date, fmt_date_short,
)
from data_loader import HEALTH_LABELS, STATUS_DONE, STATUS_RISK


def _blank_slide(pres):
    """取一張完全空白的投影片"""
    return pres.slides.add_slide(pres.slide_layouts[6])


# ══════════════════════════════════════════════════════
# Slide 1: 封面
# ══════════════════════════════════════════════════════
def slide_cover(pres, r):
    s = _blank_slide(pres)
    overall = r["overall"]
    main_c  = HEALTH[overall]["main"]
    soft_c  = HEALTH[overall]["soft"]

    # 左側色條
    rect(s, 0, 0, 0.10, SLIDE_H, main_c)

    # 主標
    text(s, 0.55, 1.2, 6.0, 0.9, "專案管理週報",
         size=42, bold=True, color=C["text"])
    text(s, 0.55, 2.05, 6.0, 0.34, "Executive Status Report",
         size=13, italic=True, color=C["muted"])

    # 細分隔線
    rect(s, 0.55, 2.62, 2.8, 0.008, C["divider"])

    # Headline 膠囊
    round_rect(s, 0.55, 2.80, 4.0, 0.70, soft_c,
               line=main_c, line_w=1.2, radius=0.18)
    text(s, 0.55, 2.80, 4.0, 0.70, r["headline"],
         size=14, bold=True, color=main_c,
         align="center", valign="middle")

    # 日期資訊
    text(s, 0.55, 3.78, 5.0, 0.34,
         f"報告日期　{fmt_date(r['report_date'])}",
         size=12, color=C["muted"])
    text(s, 0.55, 4.16, 5.0, 0.34,
         f"下次回報　{fmt_date(r['next_report'])}",
         size=12, color=C["muted"])

    # 右側狀態統計欄
    rect(s, 7.0, 0, SLIDE_W - 7.0, SLIDE_H, C["bg"])
    rect(s, 7.0, 0, SLIDE_W - 7.0, 0.06, C["divider"])

    items = [
        ("健康", r["stats"]["green_projects"], C["green"]),
        ("注意", r["stats"]["amber_projects"], C["amber"]),
        ("警報", r["stats"]["red_projects"],   C["red"]),
    ]
    for i, (label, n, col) in enumerate(items):
        ty = 1.3 + i * 1.15
        circle(s, 7.45, ty + 0.18, 0.30, col)
        text(s, 7.75, ty, 1.6, 0.45, str(n),
             size=32, bold=True, color=col, valign="middle")
        text(s, 7.75, ty + 0.50, 1.8, 0.26, f"{label} 專案",
             size=10, color=C["muted"])


# ══════════════════════════════════════════════════════
# Slide 2: 執行摘要（所有專案一覽表）
# ══════════════════════════════════════════════════════
def slide_executive_summary(pres, r):
    s = _blank_slide(pres)
    page_header(s, "執行摘要",
                f"專案一覽　·　{fmt_date(r['report_date'])}",
                accent=C["text"])

    # 頂部三個統計卡
    cards = [
        ("需立即關注", r["stats"]["red_projects"],   C["red"],   C["redSoft"]),
        ("需要觀察",   r["stats"]["amber_projects"], C["amber"], C["amberSoft"]),
        ("進度健康",   r["stats"]["green_projects"], C["green"], C["greenSoft"]),
    ]
    for i, (label, n, col, soft) in enumerate(cards):
        x = 0.3 + i * 3.15
        rect(s, x, 0.92, 2.95, 0.85, soft, line=C["divider"], line_w=0.5)
        rect(s, x, 0.92, 0.07, 0.85, col)
        text(s, x + 0.2, 0.92, 0.9, 0.85, str(n),
             size=36, bold=True, color=col, valign="middle")
        text(s, x + 1.1, 0.92, 1.8, 0.85, f"{label}　專案",
             size=11, color=C["muted"], valign="middle")

    # 表格
    TX, TY, TW = 0.3, 1.98, 9.4
    COLS = [0.4, 1.85, 0.85, 1.55, 1.30, 0.90, 2.55]
    HDRS = ["", "專案 / PM", "完成率", "進度 / 時間", "預估完工", "偏差", "本期關鍵"]

    # 表頭
    hx = TX
    for h, w in zip(HDRS, COLS):
        rect(s, hx, TY, w, 0.32, C["bg"], line=C["divider"], line_w=0.4)
        text(s, hx, TY, w, 0.32, h,
             size=9, bold=True, color=C["muted"], align="center", valign="middle")
        hx += w

    ROW_H = 0.55
    for i, p in enumerate(r["projects"]):
        y = TY + 0.32 + i * ROW_H
        bg = C["white"] if i % 2 == 0 else C["rowAlt"]
        rect(s, TX, y, TW, ROW_H, bg, line=C["divider"], line_w=0.3)

        h = p["health"]
        hc = HEALTH[h["level"]]["main"]

        rx = TX

        # 燈號
        circle(s, rx + COLS[0]/2, y + ROW_H/2, 0.22, hc)
        rx += COLS[0]

        # 專案名 + PM
        text(s, rx + 0.08, y + 0.06, COLS[1] - 0.16, 0.27, p["name"],
             size=11, bold=True, color=C["text"], valign="middle")
        text(s, rx + 0.08, y + 0.30, COLS[1] - 0.16, 0.22,
             f"{p['owner']}  ·  {h['label']}",
             size=8.5, color=hc)
        rx += COLS[1]

        # 完成率
        text(s, rx, y, COLS[2], ROW_H, f"{round(h['done_pct'])}%",
             size=16, bold=True, color=C["text"],
             align="center", valign="middle")
        rx += COLS[2]

        # 進度條 + 時間軸刻度
        bx, bw = rx + 0.12, COLS[3] - 0.24
        by = y + ROW_H/2 - 0.05
        progress_bar(s, bx, by, bw, 0.13, h["done_pct"], hc,
                     vs_pct=h["elapsed_pct"], vs_color=C["text"])
        text(s, bx, y + ROW_H - 0.18, bw, 0.16,
             f"時間 {round(h['elapsed_pct'])}%",
             size=7, color=C["light"], align="center")
        rx += COLS[3]

        # 預估完工
        if h.get("forecast_available", False):
            fc_txt = fmt_date(h["forecast_end"])
            fc_col = C["text"]
        else:
            fc_txt = "—"
            fc_col = C["light"]
        text(s, rx, y + 0.06, COLS[4], 0.27,
             fc_txt,
             size=10, bold=True, color=fc_col,
             align="center", valign="middle")
        text(s, rx, y + 0.30, COLS[4], 0.22,
             f"目標 {fmt_date(p['target'])}",
             size=7.5, color=C["light"], align="center")
        rx += COLS[4]

        # 偏差
        if not h.get("forecast_available", False):
            txt, col = "資料不足", C["light"]
        else:
            d = h["forecast_drift"]
            if d > 0:
                txt, col = f"+{d} 天", C["red"]
            elif d < 0:
                txt, col = f"{d} 天", C["green"]
            else:
                txt, col = "準時", C["green"]
        text(s, rx, y, COLS[5], ROW_H, txt,
             size=12 if "資料" not in txt else 10,
             bold=True, color=col,
             align="center", valign="middle")
        rx += COLS[5]

        # 本期關鍵：取最嚴重的議題
        if p["decisions"]:
            top_d = p["decisions"][0]
            key_txt = top_d["issue"]
            key_col = C["red"] if top_d["severity"] == "高" else C["amber"]
            italic = False
        else:
            key_txt = "進度正常"
            key_col = C["green"]
            italic = True
        text(s, rx + 0.1, y, COLS[6] - 0.2, ROW_H, key_txt,
             size=9, color=key_col, italic=italic, valign="middle")

    # 底部備註
    bot_y = TY + 0.32 + len(r["projects"]) * ROW_H + 0.15
    text(s, 0.3, bot_y, 9.4, 0.26,
         f"需主管裁示事項：{r['stats']['decision_count']} 件　·　詳見各專案 Page 2 與彙整頁",
         size=9, color=C["light"], italic=True)


# ══════════════════════════════════════════════════════
# Slide 3..N：每專案 Page A — 一眼看狀況
# ══════════════════════════════════════════════════════
def slide_project_overview(pres, p, r):
    s = _blank_slide(pres)
    h  = p["health"]
    hc = HEALTH[h["level"]]["main"]
    soft = HEALTH[h["level"]]["soft"]
    deep = HEALTH[h["level"]]["deep"]

    period = f"{fmt_date(p['start'])} — {fmt_date(p['target'])}"
    page_header(s, p["name"],
                f"PM　{p['owner']}　·　{period}",
                page_tag="專案概況　1 / 2", accent=hc)

    # === 左半部佈局參數 ===
    LX, LW = 0.25, 4.40
    CY     = 0.85

    HERO_H   = 1.55     # Hero 區（燈號 + 階段 + 一句話）
    METRIC_Y = CY + HERO_H + 0.08
    METRIC_H = 0.90
    PROG_Y   = METRIC_Y + METRIC_H + 0.08
    PROG_H   = 0.36
    DIST_Y   = PROG_Y + PROG_H + 0.10
    DIST_H   = SLIDE_H - DIST_Y - 0.30

    # ─── Hero 區 ───
    rect(s, LX, CY, LW, HERO_H, C["white"], line=C["divider"], line_w=0.5)
    # 燈號 + 標籤
    health_badge(s, LX + 0.42, CY + 0.36, 0.34, h["level"], with_ring=True)
    text(s, LX + 0.70, CY + 0.14, 2.5, 0.46, h["label"],
         size=24, bold=True, color=hc, valign="middle")
    text(s, LX + 0.70, CY + 0.54, LW - 0.85, 0.24, p["stage"] or "—",
         size=10, color=C["muted"], valign="middle")

    # 一句話結論（軟色背景框）
    sub_x = LX + 0.18
    sub_y = CY + 0.86
    sub_w = LW - 0.36
    sub_h = HERO_H - 0.86 - 0.10
    round_rect(s, sub_x, sub_y, sub_w, sub_h, soft,
               line=hc, line_w=0.6, radius=0.12)
    text(s, sub_x + 0.15, sub_y, sub_w - 0.3, sub_h, p["summary"],
         size=11, bold=True, color=deep, valign="middle")

    # ─── 三大 KPI ───
    metrics = []
    # 完成率
    metrics.append({"label": "完成率", "value": f"{round(h['done_pct'])}%", "color": hc})
    # 預估完工
    if h.get("forecast_available", False):
        metrics.append({"label": "預估完工",
                        "value": fmt_date_short(h["forecast_end"]),
                        "color": C["text"]})
    else:
        metrics.append({"label": "預估完工（資料不足）",
                        "value": "—",
                        "color": C["light"]})
    # 偏差
    if not h.get("forecast_available", False):
        drift_v, drift_c = "—", C["light"]
    else:
        d = h["forecast_drift"]
        if d > 0:
            drift_v, drift_c = f"+{d} 天", C["red"]
        elif d < 0:
            drift_v, drift_c = f"{d} 天", C["green"]
        else:
            drift_v, drift_c = "準時", C["green"]
    metrics.append({"label": "與目標偏差", "value": drift_v, "color": drift_c})

    mw = LW / 3
    for i, m in enumerate(metrics):
        kpi_card(s, LX + i * mw, METRIC_Y, mw, METRIC_H,
                 m["value"], m["label"],
                 value_color=m["color"],
                 accent_color=None, value_size=22, label_size=9.5)

    # ─── 進度 vs 時間 ───
    text(s, LX, PROG_Y, 2.0, 0.18, "進度 vs 時間",
         size=8.5, bold=True, color=C["muted"])
    text(s, LX + 2.0, PROG_Y, LW - 2.0, 0.18,
         f"完成 {round(h['done_pct'])}%　·　時間 {round(h['elapsed_pct'])}%",
         size=8.5, color=C["muted"], align="right")
    progress_bar(s, LX, PROG_Y + 0.20, LW, 0.16,
                 h["done_pct"], hc,
                 vs_pct=h["elapsed_pct"], vs_color=C["text"])

    # ─── 任務分佈條（取代原本的「下個關鍵節點」） ───
    rect(s, LX, DIST_Y, LW, DIST_H, C["white"],
         line=C["divider"], line_w=0.5)
    text(s, LX + 0.18, DIST_Y + 0.06, LW - 0.36, 0.22,
         "任務分佈",
         size=9.5, bold=True, color=C["muted"], valign="middle")
    text(s, LX + 0.18, DIST_Y + 0.06, LW - 0.36, 0.22,
         f"共 {p['total']} 個任務",
         size=8.5, color=C["light"], align="right", valign="middle")

    # 堆疊長條
    bar_x = LX + 0.20
    bar_y = DIST_Y + 0.40
    bar_w = LW - 0.40
    bar_h = 0.20
    total = max(1, p["total"])
    segs = [
        ("完成",   p["done"],        C["green"]),
        ("進行中", p["in_progress"], C["blue"]),
        ("風險",   p["risk"],        C["red"]),
        ("待開始", p["not_started"], C["divider"]),
    ]
    cx = bar_x
    for _, val, col in segs:
        w = bar_w * val / total
        if w > 0.01:
            rect(s, cx, bar_y, w, bar_h, col)
        cx += w

    # 圖例（與堆疊一致）
    legend_y = bar_y + bar_h + 0.12
    lx = bar_x
    seg_w = bar_w / len(segs)
    for label, val, col in segs:
        rect(s, lx, legend_y + 0.04, 0.14, 0.12, col)
        text(s, lx + 0.18, legend_y, seg_w - 0.20, 0.22,
             f"{label} {val}",
             size=8.5, color=C["muted"], valign="middle")
        lx += seg_w

    # ─── 右半部：里程碑甘特圖 ───
    GX, GY = 4.80, CY
    GW = SLIDE_W - GX - 0.25
    GH = SLIDE_H - GY - 0.30
    _draw_gantt(s, GX, GY, GW, GH, p["milestones"], hc, r["report_date"])


def _draw_gantt(s, x, y, w, h, milestones, accent, today):
    """里程碑甘特圖（含今日線）"""
    rect(s, x, y, w, h, C["white"], line=C["divider"], line_w=0.5)
    text(s, x + 0.2, y + 0.10, 3.0, 0.30, "里程碑進度",
         size=12, bold=True, color=C["text"], valign="middle")
    divider(s, x, y + 0.46, w)

    if not milestones:
        text(s, x, y + h/2 - 0.2, w, 0.4, "（無里程碑資料）",
             size=11, italic=True, color=C["light"], align="center")
        return

    # 時間軸範圍
    all_dates = [d for m in milestones for d in (m["start"], m["end"])]
    minD, maxD = min(all_dates), max(all_dates)
    total_d = max(1, (maxD - minD).days)

    LABEL_W = 1.55
    AREA_X = x + LABEL_W
    AREA_W = w - LABEL_W - 0.20
    HDR_Y  = y + 0.50

    # 月份標尺
    cur = date(minD.year, minD.month, 1)
    while cur <= maxD:
        lx = AREA_X + (cur - minD).days / total_d * AREA_W
        if AREA_X - 0.02 <= lx <= AREA_X + AREA_W + 0.02:
            rect(s, lx, HDR_Y, 0.006, h - 0.5, C["divider"])
            text(s, lx + 0.02, HDR_Y + 0.02, 0.4, 0.20,
                 f"{cur.month}月",
                 size=7, color=C["light"])
        # 下個月
        if cur.month == 12:
            cur = date(cur.year + 1, 1, 1)
        else:
            cur = date(cur.year, cur.month + 1, 1)

    # 今日線
    if minD <= today <= maxD:
        tx = AREA_X + (today - minD).days / total_d * AREA_W
        rect(s, tx - 0.015, HDR_Y, 0.030, h - 0.60, accent)
        text(s, tx - 0.30, y + h - 0.26, 0.60, 0.20, "今日",
             size=7.5, bold=True, color=accent, align="center")

    # 各列
    rows = milestones[:8]
    bar_area_h = h - 0.78
    row_h = min(0.50, bar_area_h / len(rows))
    row_y0 = y + 0.74

    for i, m in enumerate(rows):
        ry = row_y0 + i * row_h
        if i % 2 == 1:
            rect(s, x, ry, w, row_h, C["rowAlt"])

        col = STATUS_COLOR.get(m["status"], C["blue"])

        # 左側名稱 + 進度%
        text(s, x + 0.12, ry, LABEL_W - 0.50, row_h, m["name"],
             size=8.5, bold=True, color=C["text"], valign="middle")
        text(s, x + LABEL_W - 0.50, ry, 0.45, row_h, f"{m['progress']}%",
             size=8.5, bold=True, color=col,
             align="right", valign="middle")

        # 計畫範圍
        x1 = AREA_X + (m["start"] - minD).days / total_d * AREA_W
        x2 = AREA_X + (m["end"]   - minD).days / total_d * AREA_W
        bw = max(0.12, x2 - x1)
        barH = min(0.16, row_h * 0.42)
        barY = ry + (row_h - barH) / 2

        # 外框（計畫範圍）+ 填色（完成進度）
        rect(s, x1, barY, bw, barH,
             HEALTH["GREEN"]["soft"] if m["status"] == "完成" else C["bg"],
             line=col, line_w=0.6)
        fill_w = max(0.04, bw * m["progress"] / 100)
        rect(s, x1, barY, fill_w, barH, col)

        # 完成/總數標籤
        if x1 + bw + 0.05 < AREA_X + AREA_W - 0.4:
            text(s, x1 + bw + 0.03, ry, 0.55, row_h,
                 f"{m['done']}/{m['total']}",
                 size=7, color=C["light"], valign="middle")


# ══════════════════════════════════════════════════════
# Slide 4..N：每專案 Page B — 風險與決策
# ══════════════════════════════════════════════════════
def slide_project_risks(pres, p, r):
    s = _blank_slide(pres)
    h  = p["health"]
    hc = HEALTH[h["level"]]["main"]
    page_header(s, p["name"],
                f"風險與決策事項　·　{fmt_date(r['report_date'])}",
                page_tag="風險決策　2 / 2", accent=hc)

    # ─── KPI 條（4 個） ───
    kpi_y, kpi_h = 0.86, 0.85
    items = [
        ("本期完成",       len(p["activity"]["recent_done"]), C["green"]),
        ("進行中 / 到期",  len(p["activity"]["upcoming"]),    C["blue"]),
        ("風險議題",       len(p["decisions"]),               C["red"]),
        ("待開始",         p["not_started"],                  C["light"]),
    ]
    kw = (SLIDE_W - 0.5) / len(items)
    for i, (lab, val, col) in enumerate(items):
        x = 0.25 + i * kw
        kpi_card(s, x, kpi_y, kw - 0.12, kpi_h,
                 val, lab, value_color=col,
                 accent_color=col, value_size=26, label_size=10)

    # ─── 三欄主面板 ───
    PY = kpi_y + kpi_h + 0.18
    PH = SLIDE_H - PY - 0.32

    cols = [
        (0.25, 3.05),   # 左：近 14 天完成
        (3.42, 3.05),   # 中：進行中 / 即將到期
        (6.59, 3.16),   # 右：風險議題 + 建議行動
    ]

    # 左：近 14 天完成 ──────────────────────────────────
    cx, cw = cols[0]
    _panel_header(s, cx, cw, PY,
                  f"近 14 天完成（{len(p['activity']['recent_done'])}）",
                  C["green"], PH)
    recent = p["activity"]["recent_done"]
    if not recent:
        text(s, cx, PY + PH/2 - 0.15, cw, 0.30,
             "本期無完成任務",
             size=10, italic=True, color=C["light"], align="center")
    else:
        item_h = 0.46
        max_items = int((PH - 0.60 - 0.30) // item_h)
        for i, t in enumerate(recent[:max_items]):
            iy = PY + 0.58 + i * item_h
            circle(s, cx + 0.22, iy + 0.16, 0.10, C["green"])
            text(s, cx + 0.36, iy + 0.02, cw - 0.50, 0.24, t["name"],
                 size=9.5, bold=True, color=C["text"], valign="middle")
            text(s, cx + 0.36, iy + 0.26, cw - 0.50, 0.18,
                 f"{t['owner']}  ·  完成 {fmt_date_short(t['actual'])}",
                 size=8, color=C["muted"])
        # 底部摘要
        text(s, cx + 0.18, PY + PH - 0.26, cw - 0.36, 0.20,
             f"本期完成率　{len(recent)} / {p['total']}",
             size=8, color=C["light"], italic=True)

    # 中：進行中 / 即將到期 ──────────────────────────────
    cx, cw = cols[1]
    _panel_header(s, cx, cw, PY,
                  f"進行中 / 即將到期（{len(p['activity']['upcoming'])}）",
                  C["blue"], PH)
    upc = p["activity"]["upcoming"]
    if not upc:
        text(s, cx, PY + PH/2 - 0.15, cw, 0.30,
             "目前無進行中任務",
             size=10, italic=True, color=C["light"], align="center")
    else:
        item_h = 0.46
        max_items = int((PH - 0.60 - 0.30) // item_h)
        for i, t in enumerate(upc[:max_items]):
            iy = PY + 0.58 + i * item_h
            stc = STATUS_COLOR.get(t["status"], C["blue"])
            is_risk = t["status"] in STATUS_RISK
            if is_risk:
                rect(s, cx + 0.10, iy - 0.02, cw - 0.20, item_h - 0.04,
                     C["redSoft"])
            circle(s, cx + 0.22, iy + 0.16, 0.10, stc)
            text(s, cx + 0.36, iy + 0.02, cw - 0.50, 0.24, t["name"],
                 size=9.5, bold=True, color=C["text"], valign="middle")
            # 狀態膠囊 + 截止日
            status_pill(s, cx + 0.36, iy + 0.28, 0.45, 0.18, t["status"])
            text(s, cx + 0.86, iy + 0.28, cw - 1.0, 0.18,
                 f"{t['owner']}  ·  截止 {fmt_date_short(t['end'])}",
                 size=7.5, color=C["muted"], valign="middle")

    # 右：風險議題與建議行動 ──────────────────────────────
    cx, cw = cols[2]
    _panel_header(s, cx, cw, PY,
                  f"風險議題與建議行動（{len(p['decisions'])}）",
                  C["red"], PH)
    decs = p["decisions"]
    if not decs:
        round_rect(s, cx + 0.30, PY + PH/2 - 0.35, cw - 0.60, 0.70,
                   C["greenSoft"], line=C["divider"], line_w=0.5, radius=0.12)
        text(s, cx + 0.30, PY + PH/2 - 0.35, cw - 0.60, 0.70,
             "✓  目前無風險議題",
             size=13, bold=True, color=C["green"],
             align="center", valign="middle")
    else:
        MAX_RISK = 3
        MIN_ITEM_H = 0.95  # 每張卡至少這麼高才看得清

        # 動態調整：先試顯示 3 張，若太擠則減少
        shown_n = min(MAX_RISK, len(decs))
        while True:
            extra = len(decs) - shown_n
            bot_h = 0.30 if extra > 0 else 0.0
            avail = PH - 0.60 - 0.15 - bot_h
            item_h = min(1.30, avail / shown_n)
            if item_h >= MIN_ITEM_H or shown_n <= 1:
                break
            shown_n -= 1
        shown = decs[:shown_n]

        for i, d in enumerate(shown):
            iy = PY + 0.62 + i * item_h
            card_h = item_h - 0.10
            if d["severity"] == "高":
                sc, ss, sd = C["red"], C["redSoft"], C["redDeep"]
            elif d["severity"] == "中":
                sc, ss, sd = C["amber"], C["amberSoft"], C["amberDeep"]
            else:
                sc, ss, sd = C["light"], C["rowAlt"], C["muted"]

            # 卡片
            rect(s, cx + 0.12, iy, cw - 0.24, card_h, ss,
                 line=C["divider"], line_w=0.4)
            rect(s, cx + 0.12, iy, 0.07, card_h, sc)

            content_x = cx + 0.26
            content_w = cw - 0.42

            # 行 1：嚴重度 + 負責人
            text(s, content_x, iy + 0.05, 0.45, 0.22,
                 d["severity"],
                 size=10, bold=True, color=sc, valign="middle")
            text(s, content_x + 0.50, iy + 0.05, content_w - 0.50, 0.22,
                 d["owner"],
                 size=9, color=C["muted"], valign="middle")
            # 行 2：議題（在中段）
            issue_y = iy + 0.30
            issue_h = card_h - 0.30 - 0.36   # 上 0.30 給標題 / 下 0.36 給建議
            text(s, content_x, issue_y, content_w, issue_h,
                 d["issue"],
                 size=9.5, bold=True, color=sd, valign="top",
                 line_space=1.15)
            # 行 3：建議（卡片底部往上固定 0.30 高度）
            ask_y = iy + card_h - 0.32
            text(s, content_x, ask_y, content_w, 0.28,
                 "→  " + d["ask"],
                 size=8.8, bold=True, color=sc, valign="top",
                 line_space=1.10)

        if extra > 0:
            text(s, cx + 0.18, PY + PH - bot_h + 0.02,
                 cw - 0.36, bot_h - 0.04,
                 f"另有 {extra} 件風險議題，詳見彙整頁",
                 size=8.5, italic=True, color=C["light"], valign="middle")


def _panel_header(s, cx, cw, PY, title, color, PH):
    """三欄式面板的共通標題列"""
    rect(s, cx, PY, cw, PH, C["white"], line=C["divider"], line_w=0.5)
    rect(s, cx, PY, cw, 0.06, color)
    text(s, cx + 0.18, PY + 0.10, cw - 0.36, 0.30, title,
         size=11, bold=True, color=color, valign="middle")
    divider(s, cx, PY + 0.48, cw)


# ══════════════════════════════════════════════════════
# Slide N: 跨專案裁示事項彙整
# ══════════════════════════════════════════════════════
def slide_decisions(pres, r):
    s = _blank_slide(pres)
    high_n = sum(1 for d in r["decisions"] if d["severity"] == "高")
    mid_n  = sum(1 for d in r["decisions"] if d["severity"] == "中")
    page_header(s, "需主管裁示事項",
                f"高嚴重度 {high_n}　·　中嚴重度 {mid_n}　·　{fmt_date(r['report_date'])}",
                accent=C["red"])

    if not r["decisions"]:
        round_rect(s, 2.5, 2.2, 5.0, 1.2, C["greenSoft"],
                   line=C["divider"], line_w=0.5, radius=0.1)
        text(s, 2.5, 2.2, 5.0, 1.2, "✓  目前無需裁示事項",
             size=18, bold=True, color=C["green"],
             align="center", valign="middle")
        return

    HX, HY, HW = 0.3, 0.92, 9.4
    COLW = [0.55, 1.55, 2.65, 2.30, 2.35]
    HDRS = ["嚴重度", "專案 / PM", "議題", "影響說明", "建議行動"]

    hx = HX
    for hd, w in zip(HDRS, COLW):
        rect(s, hx, HY, w, 0.34, C["bg"], line=C["divider"], line_w=0.4)
        text(s, hx, HY, w, 0.34, hd,
             size=9.5, bold=True, color=C["muted"],
             align="center", valign="middle")
        hx += w

    # 只顯示高/中
    show = [d for d in r["decisions"] if d["severity"] in ("高", "中")][:7]
    avail = SLIDE_H - HY - 0.34 - 0.45
    RH = min(0.78, max(0.52, avail / max(len(show), 1)))

    for i, d in enumerate(show):
        y = HY + 0.34 + i * RH
        bg = C["white"] if i % 2 == 0 else C["rowAlt"]
        rect(s, HX, y, HW, RH, bg, line=C["divider"], line_w=0.3)

        sc = C["red"] if d["severity"] == "高" else C["amber"]
        dc = HX

        # 嚴重度大色塊
        round_rect(s, dc + 0.10, y + 0.16, 0.36, RH - 0.32, sc, radius=0.20)
        text(s, dc + 0.10, y + 0.16, 0.36, RH - 0.32, d["severity"],
             size=10, bold=True, color=C["white"],
             align="center", valign="middle")
        dc += COLW[0]

        # 專案 + PM
        text(s, dc + 0.08, y + 0.07, COLW[1] - 0.16, 0.30, d["project"],
             size=10, bold=True, color=C["text"], valign="middle")
        text(s, dc + 0.08, y + 0.40, COLW[1] - 0.16, RH - 0.5,
             d["owner"], size=8.5, color=C["muted"])
        dc += COLW[1]

        # 議題
        text(s, dc + 0.10, y + 0.06, COLW[2] - 0.20, RH - 0.12,
             d["issue"],
             size=9.5, color=C["text"], valign="middle")
        dc += COLW[2]

        # 影響
        text(s, dc + 0.10, y + 0.06, COLW[3] - 0.20, RH - 0.12,
             d["impact"],
             size=8.5, italic=True, color=C["muted"], valign="middle")
        dc += COLW[3]

        # 建議
        text(s, dc + 0.10, y + 0.06, COLW[4] - 0.20, RH - 0.12,
             d["ask"],
             size=9, bold=True, color=sc, valign="middle")

    # 底部備註
    low_n = sum(1 for d in r["decisions"] if d["severity"] == "低")
    extra = (high_n + mid_n) - len(show)
    notes = []
    if extra > 0:
        notes.append(f"另有 {extra} 件高/中嚴重度事項")
    if low_n > 0:
        notes.append(f"另有 {low_n} 件低嚴重度事項")
    if notes:
        text(s, HX, HY + 0.34 + len(show) * RH + 0.10, HW, 0.22,
             "　·　".join(notes) + "，請參閱各專案 Page 2",
             size=9, color=C["light"], italic=True)


# ══════════════════════════════════════════════════════
# 結尾頁
# ══════════════════════════════════════════════════════
def slide_closing(pres, r):
    s = _blank_slide(pres)
    main_c = HEALTH[r["overall"]]["main"]

    rect(s, 0, 0, 0.10, SLIDE_H, main_c)
    rect(s, SLIDE_W - 0.06, 0, 0.06, SLIDE_H, C["divider"])

    text(s, 0, 1.9, SLIDE_W, 0.9, "謝謝聆聽",
         size=44, bold=True, color=C["text"], align="center")
    rect(s, 3.5, 3.0, 3.0, 0.008, C["divider"])
    text(s, 0, 3.20, SLIDE_W, 0.36,
         f"本期報告　{fmt_date(r['report_date'])}　　·　　下次回報　{fmt_date(r['next_report'])}",
         size=12, color=C["muted"], align="center")


# ══════════════════════════════════════════════════════
# 主入口
# ══════════════════════════════════════════════════════
def build_ppt(report: dict, output_path: str) -> None:
    pres = Presentation()
    pres.slide_width  = Inches(SLIDE_W)
    pres.slide_height = Inches(SLIDE_H)

    slide_cover(pres, report)
    slide_executive_summary(pres, report)
    for p in report["projects"]:
        slide_project_overview(pres, p, report)
        slide_project_risks(pres, p, report)
    slide_decisions(pres, report)
    slide_closing(pres, report)

    pres.save(output_path)