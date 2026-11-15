# -*- coding: utf-8 -*-
"""鉴定栏分数自动检查与修正。

规则（依据《zixin-nxy 德育分工具功能需求说明书-鉴定栏分数检查》）：
- 表扬信：每封 0.5 分，按换行分隔的每条记录计一封；整个学期封顶 10 封（5 分），
  先计前一个月份段，再用剩余额度计后一个月份段。
- 活动分：累加文本里所有「+数字」；含「新闻报道」的条目合计封顶 5 分，其余不封顶。
- 各鉴定栏：F = 第一段活动分 + 第一段表扬信分；K = 第二段活动分 + 第二段表扬信分；
  O = 附加分说明里所有分值之和。
- 总分 P 按表内公式 C+F-H+K-M+O 勾稽，只核对不修改。
- 一键修正只写 F/K/O 三个单元格，另存为副本。
"""
import os
import re

import openpyxl

# 单元格里「+2」「+0.5」「+0，5」「＋3」这类分值
POINT_RE = re.compile(r"[+＋]\s*(\d+(?:[.,，]\d+)?)")
# 表头里的月份段：9-10月 / X-X月 / 3月
MONTH_HEAD_RE = re.compile(r"(\d+\s*[-—~]\s*\d+\s*月|[Xx]\s*[-—~]\s*[Xx]\s*月|\d+\s*月)")

PRAISE_PER = 0.5          # 每封表扬信的分值
PRAISE_CAP_COUNT = 10     # 整学期封顶封数
NEWS_CAP = 5.0            # 新闻报道类活动合计封顶

EPS = 1e-6


# ---------------------------------------------------------------- 解析
def parse_points(text):
    """把一段文本里所有「+数字」加总；逗号小数点（+0,5）按 0.5 处理。"""
    if text is None:
        return 0.0
    total = 0.0
    for m in POINT_RE.finditer(str(text)):
        try:
            total += float(m.group(1).replace(",", ".").replace("，", "."))
        except ValueError:
            continue
    return round(total, 2)


def split_records(value):
    """单元格按换行拆成条目；空行忽略。"""
    if value is None:
        return []
    s = str(value).replace("\r\n", "\n").replace("\r", "\n")
    return [line.strip() for line in s.split("\n") if line.strip()]


def activity_score(value):
    """活动分：普通条目直接累加；含「新闻报道」的条目合计封顶 5 分。"""
    news, other = 0.0, 0.0
    for line in split_records(value):
        pts = parse_points(line)
        if "新闻报道" in line or "新闻" in line and "报道" in line:
            news += pts
        else:
            other += pts
    return round(other + min(news, NEWS_CAP), 2)


def praise_count(value):
    """表扬信封数：按非空换行计。"""
    return len(split_records(value))


# ---------------------------------------------------------------- 列定位
def _expand_row(ws, row, max_col):
    """把某一行因合并而空置的单元格，用合并区左上角标题填充。"""
    vals = {}
    for c in range(1, max_col + 1):
        v = ws.cell(row, c).value
        vals[c] = str(v).strip() if v is not None else ""
    for mc in ws.merged_cells.ranges:
        if mc.min_row <= row <= mc.max_row:
            top = ws.cell(mc.min_row, mc.min_col).value
            top = str(top).strip() if top is not None else ""
            if top:
                for c in range(mc.min_col, mc.max_col + 1):
                    vals[c] = top
    return vals


def locate_columns(ws):
    """按表头文字定位各列，返回 dict；定位不到关键列时抛 ValueError。

    返回：
      name        姓名列
      base        基础分鉴定（C）
      groups      [{"desc": 活动列, "praise": 表扬信列, "judge": 鉴定列, "label": 段名}]
      deduct      [扣分鉴定列...]（H、M）
      add_desc    附加分说明列（N）
      add_judge   附加分鉴定列（O）
      total       总分列（P）
    """
    max_col = ws.max_column or 0
    row3 = _expand_row(ws, 3, max_col)
    row4 = _expand_row(ws, 4, max_col)

    name_col = None
    for c in range(1, max_col + 1):
        if "姓名" in row3.get(c, "") or "姓名" in row4.get(c, ""):
            name_col = c
            break
    if name_col is None:
        raise ValueError("没找到「姓名」列（表头第 3 / 4 行）")

    base_col = None
    groups = {}
    deduct = []
    add_desc = add_judge = total = None

    for c in range(1, max_col + 1):
        big = row3.get(c, "")
        sub = row4.get(c, "")
        if not big and not sub:
            continue
        if "加分项" in big:
            g = groups.setdefault(big, {"label": big, "desc": None,
                                        "praise": None, "judge": None})
            if "表扬信" in sub:
                g["praise"] = c
            elif "鉴定" in sub:
                g["judge"] = c
            elif MONTH_HEAD_RE.search(sub) or not sub:
                if g["desc"] is None:
                    g["desc"] = c
        elif "扣分项" in big:
            if "鉴定" in sub:
                deduct.append(c)
        elif "附加分" in big:
            if "鉴定" in sub:
                add_judge = c
            elif add_desc is None:
                add_desc = c
        elif "基础分" in big:
            if "鉴定" in sub:
                base_col = c
        elif "总分" in big:
            total = c

    ordered = []
    for label, g in sorted(groups.items(), key=lambda kv: min(
            [x for x in (kv[1]["desc"], kv[1]["praise"], kv[1]["judge"]) if x] or [0])):
        if g["desc"] and g["judge"]:
            ordered.append(g)
    if not ordered:
        raise ValueError("没找到「加分项 / 鉴定」列（表头第 3 / 4 行）")

    return {"name": name_col, "base": base_col, "groups": ordered,
            "deduct": sorted(deduct), "add_desc": add_desc,
            "add_judge": add_judge, "total": total,
            "max_col": max_col}


# ---------------------------------------------------------------- 逐行计算
def _num(v):
    """单元格转数字；空值或非数字返回 None。"""
    if v is None:
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip()
    if not s:
        return None
    try:
        return float(s)
    except ValueError:
        return None


def _is_summary_row(name):
    return any(k in str(name) for k in ("合计", "总计", "小计", "说明", "备注", "签字"))


def audit_sheet(ws, first_row=5):
    """逐行核对，返回 (cols, rows, problems)。

    rows 每项：
      name, row, current {f,k,o}, expect {f,k,o}, diff {f,k,o},
      praise_used, total_ok, notes[]
    problems：整体性提示（如总分列缺失）。
    """
    cols = locate_columns(ws)
    rows, problems = [], []
    if cols["total"] is None:
        problems.append("没找到「总分」列，P 勾稽将跳过")

    for r in range(first_row, (ws.max_row or first_row) + 1):
        name = ws.cell(r, cols["name"]).value
        name = str(name).strip() if name is not None else ""
        if not name or _is_summary_row(name):
            continue

        budget = PRAISE_CAP_COUNT      # 剩余可用封数
        expect = {}
        praise_info = []
        for g in cols["groups"]:
            act = activity_score(ws.cell(r, g["desc"]).value)
            cnt = praise_count(ws.cell(r, g["praise"]).value)
            used = min(cnt, budget)
            budget -= used
            praise_score = round(used * PRAISE_PER, 2)
            expect[g["judge"]] = round(act + praise_score, 2)
            praise_info.append({"label": g["label"], "count": cnt,
                                "used": used, "activity": act,
                                "praise_score": praise_score})

        add_val = parse_points(ws.cell(r, cols["add_desc"]).value) \
            if cols["add_desc"] else 0.0
        if cols["add_judge"]:
            expect[cols["add_judge"]] = round(add_val, 2)

        current = {}
        for c in expect:
            current[c] = _num(ws.cell(r, c).value)

        diff = {}
        for c, exp in expect.items():
            cur = current[c]
            diff[c] = None if cur is None else round(cur - exp, 2)

        # 总分勾稽：用当前表里的数值算
        total_ok = None
        if cols["total"]:
            p_val = _num(ws.cell(r, cols["total"]).value)
            base = _num(ws.cell(r, cols["base"]).value) if cols["base"] else 0.0
            sub = 0.0
            for c in cols["deduct"]:
                v = _num(ws.cell(r, c).value)
                sub += v or 0.0
            cur_sum = (base or 0.0) + sum(current.get(c) or 0.0 for c in expect) - sub
            if p_val is None:
                total_ok = None   # 公式没算过（缓存为空），无法判断
            else:
                total_ok = abs(p_val - cur_sum) < 0.005

        rows.append({"name": name, "row": r, "current": current,
                     "expect": expect, "diff": diff,
                     "praise": praise_info,
                     "total_ok": total_ok})

    return cols, rows, problems


# ---------------------------------------------------------------- 差异汇总
def collect_issues(rows):
    """把每行的差异整理成待修正清单。

    规则：应得为 0 而单元格为空 → 视为正常（不报）；
    应得不为 0 而单元格为空 → 报「未填」；有值但不等 → 报「多计 / 少计」。
    返回 [(row, col, current, expected, reason), ...]
    """
    issues = []
    for r in rows:
        for c, exp in r["expect"].items():
            cur = r["current"].get(c)
            if cur is None:
                if abs(exp) < EPS:
                    continue
                issues.append((r["row"], c, None, exp, f"未填，应计 {exp:g}"))
                continue
            d = round(cur - exp, 2)
            if abs(d) < 0.005:
                continue
            issues.append((r["row"], c, cur, exp,
                           f"{'多计' if d > 0 else '少计'} {abs(d):g}"))
    return issues


# ---------------------------------------------------------------- 修正
def apply_fixes(src_path, dst_path, fixes):
    """把 fixes 里的 (行号, 列号, 新值) 写进副本，其余内容与格式保持不变。"""
    wb = openpyxl.load_workbook(src_path)
    try:
        ws = wb.active
        for row, col, value in fixes:
            ws.cell(row, col).value = value
        d = os.path.dirname(os.path.abspath(dst_path))
        if d:
            os.makedirs(d, exist_ok=True)
        wb.save(dst_path)
    finally:
        wb.close()
    return dst_path
