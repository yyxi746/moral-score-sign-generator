# -*- coding: utf-8 -*-
"""
德育分签字表生成器
- 不写死文件名 / 工作表 / 表头名 / 列位置 / 人数，通用适配。
- 表头自动识别（姓名 / 总分 / 表扬信），可手动调整。
- 表扬信按“月份段”分组，可只选某几个月份段；按条目条数累加，没有填 0。
- Word 标题自动识别为“年级 学年 第X学期德育分确认签字表”。
- 输出文件名：数据源名中的“统计表”替换为“签字表”。
- 模板路径、固定输出目录通过 config.json 记住，下次打开不重置。
"""
import os
import re
import sys
import json
import copy

import openpyxl
from openpyxl.utils import get_column_letter

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from docx import Document
from docx.shared import Pt
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT

HEADER_SCAN_ROWS = 8
NAME_KW = ["姓名", "名字", "学生姓名", "name"]
SCORE_KW = ["总分", "合计", "总成绩", "总得分", "最终得分", "德育总分", "德育分"]
PRAISE_KW = ["表扬信", "通报表扬", "表扬"]
ROW_BLACK = ["合计", "总计", "小计", "说明", "备注", "签字", "序号", "编号",
             "日期", "制表", "审核", "姓名", "班级", "统计表", "评分", "汇总",
             "一览", "学年", "学期", "表格"]
TEMPLATE_NAME = "签字表模板.docx"
CONFIG_NAME = "config.json"
APP_DIR_NAME = "德育分签字表生成器"
CN_NUM = {"1": "一", "2": "二", "3": "三", "4": "四",
          "一": "一", "二": "二", "三": "三", "四": "四"}


# ============================ 通用辅助 ============================
def app_dir():
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def is_number_text(text):
    try:
        float(str(text).strip())
        return True
    except (ValueError, TypeError):
        return False


def fmt_score(v):
    if v is None:
        return ""
    if isinstance(v, (int, float)):
        f = float(v)
        return str(int(f)) if f.is_integer() else str(f)
    s = str(v).strip()
    if is_number_text(s):
        f = float(s)
        return str(int(f)) if f.is_integer() else str(f)
    return s


def match_any(header, keywords):
    h = str(header).lower()
    return any(k.lower() in h for k in keywords)


# ============================ 配置持久化 ============================
def app_config_dir():
    """系统隐藏目录（%APPDATA%\\德育分签字表生成器），EXE 放哪都不影响。"""
    base = os.environ.get("APPDATA") or os.path.expanduser("~")
    return os.path.join(base, APP_DIR_NAME)


def config_file():
    d = app_config_dir()
    try:
        os.makedirs(d, exist_ok=True)
    except Exception:
        pass
    return os.path.join(d, CONFIG_NAME)


def load_config():
    try:
        with open(config_file(), "r", encoding="utf-8") as f:
            data = json.load(f)
            return data if isinstance(data, dict) else {}
    except Exception:
        return {}


def save_config(cfg):
    try:
        with open(config_file(), "w", encoding="utf-8") as f:
            json.dump(cfg, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


# ============================ Excel 读取与识别 ============================
def list_sheets(path):
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    names = list(wb.sheetnames)
    wb.close()
    return names


def scan_headers(ws):
    max_col = ws.max_column or 1
    max_row = ws.max_row or 1
    scan = min(HEADER_SCAN_ROWS, max_row)
    headers = {}
    for c in range(1, max_col + 1):
        parts = []
        for r in range(1, scan + 1):
            v = ws.cell(r, c).value
            if v is not None and str(v).strip():
                parts.append(str(v).strip())
        headers[c] = " ".join(parts)
    return headers


def column_labels(headers):
    out = []
    for c in sorted(headers):
        letter = get_column_letter(c)
        h = headers[c].replace("\n", " ")
        if len(h) > 18:
            h = h[:18] + "…"
        out.append((f"{letter}｜{h}", c))
    return out


def detect_columns(headers):
    cols = sorted(headers)
    name_cands = [c for c in cols if match_any(headers[c], NAME_KW)]
    score_cands = [c for c in cols
                   if match_any(headers[c], SCORE_KW) and "基础分" not in headers[c]]
    praise_cands = [c for c in cols if match_any(headers[c], PRAISE_KW)]
    name_col = name_cands[0] if name_cands else None
    score_col = score_cands[-1] if score_cands else None
    return name_col, score_col, praise_cands


def extract_class(ws):
    max_row = ws.max_row or 1
    scan = min(HEADER_SCAN_ROWS, max_row)
    for r in range(1, scan + 1):
        for c in range(1, (ws.max_column or 1) + 1):
            v = ws.cell(r, c).value
            if v and "班级" in str(v):
                t = str(v)
                for sep in ["：", ":"]:
                    if sep in t:
                        name = t.split(sep, 1)[1].strip()
                        if name:
                            return name
    return ""


# ---------- 表扬信：按月份段分组 ----------
def find_month_group(ws, col):
    cands = []
    for rng in ws.merged_cells.ranges:
        if rng.min_col <= col <= rng.max_col and rng.min_row <= HEADER_SCAN_ROWS:
            v = ws.cell(rng.min_row, rng.min_col).value
            if v:
                cands.append(str(v).strip())
    for t in cands:
        if re.search(r"\d+\s*月", t):
            return t
    return cands[0] if cands else None


def month_of_group(text):
    m = re.search(r"(\d+\s*[-—~]\s*\d+\s*月)", text)
    if m:
        return m.group(1).replace(" ", "")
    m = re.search(r"(\d+\s*月)", text)
    return m.group(1) if m else text


def group_praise_by_month(ws, praise_cols):
    groups, order = {}, []
    for c in praise_cols:
        g = find_month_group(ws, c)
        month = month_of_group(g) if g else (get_column_letter(c) + "列")
        if month not in groups:
            groups[month] = []
            order.append(month)
        groups[month].append(c)
    return [(m, groups[m]) for m in order]


# ---------- 标题信息 ----------
def extract_title(ws, sheet_name):
    parts_text = []
    max_row = ws.max_row or 1
    for r in range(1, min(HEADER_SCAN_ROWS, max_row) + 1):
        for c in range(1, (ws.max_column or 1) + 1):
            v = ws.cell(r, c).value
            if v:
                parts_text.append(str(v))
    blob = " ".join(parts_text)
    ym = re.search(r"(20\d{2}\s*[-—~]\s*20\d{2})", blob)
    tm = re.search(r"第\s*([1-4一二三四])\s*学期", blob)
    term = ("第" + CN_NUM[tm.group(1)] + "学期") if tm else ""
    head = " ".join(p for p in [sheet_name or "",
                                ym.group(1) if ym else "", term] if p)
    return head + "德育分确认签字表"


# ---------- 数据行 ----------
def is_valid_name(text):
    if not text:
        return False
    text = str(text).strip()
    if not text or is_number_text(text):
        return False
    if any(k in text for k in ROW_BLACK):
        return False
    return True


def count_praise_cell(v):
    if v is None:
        return 0
    return sum(1 for line in str(v).split("\n") if line.strip())


def extract_students(ws, name_col, score_col, praise_cols):
    students = []
    max_row = ws.max_row or 1
    max_col = ws.max_column or 1
    if not name_col or name_col > max_col:
        return students
    for r in range(1, max_row + 1):
        nv = ws.cell(r, name_col).value
        name = "" if nv is None else str(nv).strip()
        if not is_valid_name(name):
            continue
        score = fmt_score(ws.cell(r, score_col).value) if score_col and score_col <= max_col else ""
        praise = 0
        for pc in praise_cols:
            if pc and pc <= max_col:
                praise += count_praise_cell(ws.cell(r, pc).value)
        students.append({"name": name, "score": score, "praise": praise})
    return students


# ---------- 输出文件名 ----------
def output_file_name(excel_path):
    stem = os.path.splitext(os.path.basename(excel_path))[0]
    if "统计表" in stem:
        name = stem.replace("统计表", "签字表")
    else:
        name = stem + "签字表"
    return name + ".docx"


def compute_output_path(excel_path, output_dir):
    name = output_file_name(excel_path)
    d = output_dir.strip() if output_dir and output_dir.strip() else os.path.dirname(excel_path)
    return os.path.join(d, name)


# ============================ Word 填充 ============================
def set_run_font(run, name="宋体", size=12, bold=False):
    run.font.name = name
    run.font.size = Pt(size)
    run.font.bold = bold
    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.append(rfonts)
    for attr in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
        rfonts.set(qn(attr), name)


def fill_cell(cell, text, size=12):
    p = cell.paragraphs[0]
    run = p.add_run("" if text is None else str(text))
    set_run_font(run, size=size)
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER


def replace_paragraph(p, text, size=18, bold=True):
    for r in list(p.runs):
        r._element.getparent().remove(r._element)
    run = p.add_run(text)
    set_run_font(run, size=size, bold=bold)


def set_class_on_paragraph(p, class_name):
    t = p.text.strip()
    size = 12
    if p.runs and p.runs[0].font.size:
        size = p.runs[0].font.size.pt
    if t == "班级":
        run = p.add_run("：" + class_name)
        set_run_font(run, size=size)
    elif t in ("班级：", "班级:"):
        run = p.add_run(class_name)
        set_run_font(run, size=size)


def set_title_row(table, title, class_name):
    cell = table.rows[0].cells[0]
    for p in cell.paragraphs:
        t = p.text.strip()
        if t.startswith("班级"):
            if class_name:
                set_class_on_paragraph(p, class_name)
        elif "德育分确认签字表" in t or "签字表" in t:
            if title:
                replace_paragraph(p, title, size=18, bold=True)


def fill_word(template_path, output_path_, students, title, class_name):
    doc = Document(template_path)
    if not doc.tables:
        raise ValueError("模板中没有找到表格，请检查 Word 模板。")
    table = doc.tables[0]
    tbl = table._tbl
    all_tr = tbl.findall(qn("w:tr"))
    if len(all_tr) < 3:
        raise ValueError("模板表格结构异常（至少需要标题行、表头行、一条数据行）。")

    sample_tr = copy.deepcopy(all_tr[2])
    for tr in all_tr[2:]:
        tbl.remove(tr)

    for st in students:
        new_tr = copy.deepcopy(sample_tr)
        tbl.append(new_tr)
        row = table.rows[-1]
        values = [st["name"], st["score"], st["praise"], "", ""]
        for ci, val in enumerate(values):
            fill_cell(row.cells[ci], val, size=12)

    set_title_row(table, title, class_name)
    doc.save(output_path_)


# ============================ 图形界面 ============================
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("德育分签字表生成器")
        self.geometry("820x900")
        self.minsize(780, 840)

        cfg = load_config()

        self.excel_path = tk.StringVar()
        tpl = cfg.get("template_path", "")
        init_tpl = tpl if (tpl and os.path.exists(tpl)) else os.path.join(app_dir(), TEMPLATE_NAME)
        self.template_path = tk.StringVar(value=init_tpl)
        self.output_dir = tk.StringVar(value=cfg.get("output_dir", ""))
        self.output_path = tk.StringVar()
        self.class_name = tk.StringVar()
        self.sheet_name = tk.StringVar()
        self.title_text = tk.StringVar()

        self.wb = None
        self.headers = {}
        self.col_options = []
        self.label_to_col = {}
        self.name_label = tk.StringVar()
        self.score_label = tk.StringVar()
        self.month_groups = []
        self.month_vars = {}

        self._build_ui()

    def _build_ui(self):
        pad = {"padx": 8, "pady": 4}
        root = ttk.Frame(self)
        root.pack(fill="both", expand=True, padx=10, pady=10)

        r = 0
        ttk.Label(root, text="Excel 数据源：").grid(row=r, column=0, sticky="w", **pad)
        ttk.Entry(root, textvariable=self.excel_path).grid(
            row=r, column=1, columnspan=3, sticky="we", **pad)
        ttk.Button(root, text="浏览…", command=self.browse_excel).grid(row=r, column=4, **pad)

        r += 1
        ttk.Label(root, text="工作表：").grid(row=r, column=0, sticky="w", **pad)
        self.sheet_combo = ttk.Combobox(root, textvariable=self.sheet_name, state="readonly")
        self.sheet_combo.grid(row=r, column=1, columnspan=3, sticky="we", **pad)
        self.sheet_combo.bind("<<ComboboxSelected>>", lambda e: self.analyze())

        r += 1
        ttk.Label(root, text="Word 模板：").grid(row=r, column=0, sticky="w", **pad)
        ttk.Entry(root, textvariable=self.template_path).grid(
            row=r, column=1, columnspan=3, sticky="we", **pad)
        ttk.Button(root, text="浏览…", command=self.browse_template).grid(row=r, column=4, **pad)

        r += 1
        ttk.Label(root, text="Word 标题：").grid(row=r, column=0, sticky="w", **pad)
        ttk.Entry(root, textvariable=self.title_text).grid(
            row=r, column=1, columnspan=3, sticky="we", **pad)

        r += 1
        ttk.Label(root, text="班级：").grid(row=r, column=0, sticky="w", **pad)
        ttk.Entry(root, textvariable=self.class_name).grid(
            row=r, column=1, columnspan=3, sticky="we", **pad)

        r += 1
        ttk.Label(root, text="输出目录：").grid(row=r, column=0, sticky="w", **pad)
        ttk.Entry(root, textvariable=self.output_dir).grid(
            row=r, column=1, columnspan=2, sticky="we", **pad)
        ttk.Button(root, text="浏览…", command=self.browse_output_dir).grid(
            row=r, column=3, **pad)
        ttk.Button(root, text="清除", command=self.clear_output_dir).grid(
            row=r, column=4, **pad)

        r += 1
        ttk.Label(root, text="（输出目录留空 = 与数据源同目录）",
                  foreground="gray").grid(row=r, column=1, columnspan=3, sticky="w", padx=8)

        r += 1
        ttk.Label(root, text="输出文件：").grid(row=r, column=0, sticky="w", **pad)
        ttk.Entry(root, textvariable=self.output_path).grid(
            row=r, column=1, columnspan=3, sticky="we", **pad)
        ttk.Button(root, text="浏览…", command=self.browse_output).grid(row=r, column=4, **pad)

        r += 1
        map_frame = ttk.LabelFrame(root, text="列映射（已自动识别，可手动调整）")
        map_frame.grid(row=r, column=0, columnspan=5, sticky="we", padx=8, pady=6)
        ttk.Label(map_frame, text="姓名列：").grid(row=0, column=0, sticky="w", padx=6, pady=4)
        self.name_combo = ttk.Combobox(map_frame, textvariable=self.name_label,
                                      state="readonly", width=24)
        self.name_combo.grid(row=0, column=1, sticky="w", padx=6)
        self.name_combo.bind("<<ComboboxSelected>>", lambda e: self.refresh_preview())

        ttk.Label(map_frame, text="德育分(总分)列：").grid(row=0, column=2, sticky="w", padx=6)
        self.score_combo = ttk.Combobox(map_frame, textvariable=self.score_label,
                                       state="readonly", width=24)
        self.score_combo.grid(row=0, column=3, sticky="w", padx=6)
        self.score_combo.bind("<<ComboboxSelected>>", lambda e: self.refresh_preview())

        ttk.Label(map_frame, text="表扬信\n(按月份段多选)：").grid(
            row=1, column=0, sticky="nw", padx=6, pady=6)
        self.praise_frame = ttk.Frame(map_frame)
        self.praise_frame.grid(row=1, column=1, columnspan=3, sticky="w", padx=6)

        r += 1
        prev_frame = ttk.LabelFrame(root, text="数据预览")
        prev_frame.grid(row=r, column=0, columnspan=5, sticky="nsew", padx=8, pady=6)
        self.tree = ttk.Treeview(prev_frame, columns=("name", "score", "praise"),
                                 show="headings", height=8)
        self.tree.heading("name", text="姓名")
        self.tree.heading("score", text="德育分")
        self.tree.heading("praise", text="表扬信")
        self.tree.column("name", width=220, anchor="center")
        self.tree.column("score", width=170, anchor="center")
        self.tree.column("praise", width=170, anchor="center")
        self.tree.pack(side="left", fill="both", expand=True, padx=(6, 0), pady=6)
        sb = ttk.Scrollbar(prev_frame, orient="vertical", command=self.tree.yview)
        sb.pack(side="right", fill="y", pady=6)
        self.tree.configure(yscrollcommand=sb.set)

        r += 1
        self.count_label = ttk.Label(root, text="共识别到 0 人")
        self.count_label.grid(row=r, column=0, columnspan=5, sticky="w", padx=10)

        r += 1
        self.gen_btn = ttk.Button(root, text="生成签字表", command=self.generate)
        self.gen_btn.grid(row=r, column=0, columnspan=5, pady=10)

        r += 1
        self.status = ttk.Label(root, text="就绪", foreground="gray")
        self.status.grid(row=r, column=0, columnspan=5, sticky="w", padx=10)

        root.columnconfigure(1, weight=1)
        root.rowconfigure(9, weight=1)

    # ---------- 文件选择 ----------
    def browse_excel(self):
        try:
            path = filedialog.askopenfilename(
                title="选择德育分统计表",
                filetypes=[("Excel 文件", "*.xlsx *.xlsm"), ("所有文件", "*.*")])
            if path:
                self.excel_path.set(path)
                self.load_excel()
        except Exception as ex:
            self._error(ex)

    def browse_template(self):
        path = filedialog.askopenfilename(
            title="选择 Word 签字表模板",
            filetypes=[("Word 文件", "*.docx"), ("所有文件", "*.*")])
        if path:
            self.template_path.set(path)
            self._save_cfg()

    def browse_output_dir(self):
        d = filedialog.askdirectory(title="选择固定输出目录")
        if d:
            self.output_dir.set(d)
            self._save_cfg()
            self._recompute_output()

    def clear_output_dir(self):
        self.output_dir.set("")
        self._save_cfg()
        self._recompute_output()

    def browse_output(self):
        path = filedialog.asksaveasfilename(
            title="保存为", defaultextension=".docx",
            filetypes=[("Word 文件", "*.docx")])
        if path:
            self.output_path.set(path)

    # ---------- 加载 / 分析 ----------
    def load_excel(self):
        path = self.excel_path.get()
        if not os.path.exists(path):
            raise FileNotFoundError("所选 Excel 文件不存在。")
        if self.wb is not None:
            try:
                self.wb.close()
            except Exception:
                pass
        self.wb = openpyxl.load_workbook(path, data_only=True)
        names = list(self.wb.sheetnames)
        self.sheet_combo["values"] = names
        self.sheet_name.set(names[0] if names else "")
        self.analyze()

    def _current_ws(self):
        name = self.sheet_name.get()
        if not name or self.wb is None:
            return None
        return self.wb[name]

    def analyze(self):
        ws = self._current_ws()
        if ws is None:
            return
        self.headers = scan_headers(ws)
        self.col_options = column_labels(self.headers)
        self.label_to_col = {label: c for label, c in self.col_options}

        name_col, score_col, praise_cols = detect_columns(self.headers)
        labels_by_col = {c: label for label, c in self.col_options}
        self.name_label.set(labels_by_col.get(name_col, ""))
        self.score_label.set(labels_by_col.get(score_col, ""))

        self.month_groups = group_praise_by_month(ws, praise_cols)
        for w in self.praise_frame.winfo_children():
            w.destroy()
        self.month_vars = {}
        for i, (month, cols) in enumerate(self.month_groups):
            var = tk.IntVar(value=1)
            self.month_vars[month] = var
            ttk.Checkbutton(self.praise_frame, text=f"{month}（{len(cols)}列）",
                            variable=var,
                            command=self.refresh_preview).grid(
                row=i // 3, column=i % 3, sticky="w", padx=6, pady=3)

        self.title_text.set(extract_title(ws, self.sheet_name.get()))
        cls = extract_class(ws)
        if cls:
            self.class_name.set(cls)

        self._recompute_output()
        self.refresh_preview()

    def _recompute_output(self):
        excel = self.excel_path.get()
        if excel:
            self.output_path.set(compute_output_path(excel, self.output_dir.get()))

    def _selected_cols(self):
        name_col = self.label_to_col.get(self.name_label.get())
        score_col = self.label_to_col.get(self.score_label.get())
        praise_cols = []
        for month, cols in self.month_groups:
            if self.month_vars[month].get():
                praise_cols.extend(cols)
        return name_col, score_col, praise_cols

    def current_students(self):
        ws = self._current_ws()
        if ws is None:
            return []
        name_col, score_col, praise_cols = self._selected_cols()
        return extract_students(ws, name_col, score_col, praise_cols)

    def refresh_preview(self):
        try:
            students = self.current_students()
            for item in self.tree.get_children():
                self.tree.delete(item)
            for st in students[:8]:
                self.tree.insert("", "end",
                                 values=(st["name"], st["score"], st["praise"]))
            self.count_label.config(text=f"共识别到 {len(students)} 人")
        except Exception as ex:
            self._error(ex)

    def _save_cfg(self):
        save_config({"template_path": self.template_path.get(),
                     "output_dir": self.output_dir.get().strip()})

    # ---------- 生成 ----------
    def generate(self):
        try:
            excel = self.excel_path.get()
            template = self.template_path.get()
            output = self.output_path.get()
            if not excel or not os.path.exists(excel):
                messagebox.showwarning("提示", "请先选择有效的 Excel 文件。")
                return
            if not template or not os.path.exists(template):
                messagebox.showwarning("提示", "未找到 Word 模板，请选择“签字表模板.docx”。")
                return
            name_col, score_col, _ = self._selected_cols()
            if not name_col:
                messagebox.showwarning("提示", "请指定“姓名列”。")
                return
            if not score_col:
                messagebox.showwarning("提示", "请指定“德育分(总分)列”。")
                return
            students = self.current_students()
            if not students:
                messagebox.showwarning("提示", "没有识别到任何学生数据，请检查列映射。")
                return
            out_dir = os.path.dirname(output)
            if not output or not out_dir:
                messagebox.showwarning("提示", "请指定输出文件路径。")
                return
            if not os.path.isdir(out_dir):
                messagebox.showwarning("提示", f"输出目录不存在：\n{out_dir}")
                return

            self.gen_btn.config(state="disabled")
            self.status.config(text="正在生成…")
            self.update_idletasks()
            fill_word(template, output, students,
                      self.title_text.get(), self.class_name.get())
            self.status.config(text=f"已生成：{output}")
            self._ask_open(output, len(students))
        except Exception as ex:
            self._error(ex)
        finally:
            try:
                self.gen_btn.config(state="normal")
            except Exception:
                pass

    def _ask_open(self, output, n):
        top = tk.Toplevel(self)
        top.title("生成成功")
        top.transient(self)
        top.grab_set()
        top.geometry(f"360x150+{self.winfo_x()+200}+{self.winfo_y()+220}")
        ttk.Label(top, text=f"已生成，共 {n} 人。").pack(pady=(24, 8))
        btns = ttk.Frame(top)
        btns.pack(pady=8)

        def open_file():
            top.destroy()
            try:
                os.startfile(output)
            except Exception as ex:
                self._error(ex)

        def open_dir():
            top.destroy()
            try:
                os.startfile(os.path.dirname(output))
            except Exception as ex:
                self._error(ex)

        ttk.Button(btns, text="打开文件", command=open_file).grid(row=0, column=0, padx=8)
        ttk.Button(btns, text="打开文件夹", command=open_dir).grid(row=0, column=1, padx=8)
        ttk.Button(btns, text="关闭", command=top.destroy).grid(row=0, column=2, padx=8)

    def _error(self, ex):
        try:
            self.status.config(text="出错了")
            messagebox.showerror("出错了", f"{ex}")
        except Exception:
            pass


def main():
    App().mainloop()


if __name__ == "__main__":
    main()
