# -*- coding: utf-8 -*-
"""自动化回归测试（自包含，不依赖本机任何文件）。

运行：python test_core.py
全部通过时最后打印 ALL PASS；有失败会列出具体项目并返回退出码 1。

样本数据全部在临时目录里现场生成：
- 一张仿真实结构的德育分统计表（第3行大类、第4行细分、数据从第5行）
- 一个 5 列 Word 模板
"""
import os
import sys
import shutil
import traceback

sys.stdout.reconfigure(encoding="utf-8")
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import json
import time

import openpyxl
from docx import Document

import judge_audit
import main

TMP = None
RESULTS = []
SKIPPED = []


class Skip(Exception):
    """当前环境跑不了的检查（如没有图形界面）。"""


# ---------------------------------------------------------------- 测试框架
def check(name, func):
    """执行一个检查项，记录通过/失败。"""
    try:
        func()
        RESULTS.append((name, True, ""))
        print(f"  [OK]   {name}")
    except Skip as exc:
        SKIPPED.append(name)
        print(f"  [SKIP] {name}（{exc}）")
    except Exception as exc:
        RESULTS.append((name, False, str(exc)))
        print(f"  [FAIL] {name}")
        print(f"         {exc}")
        if os.environ.get("TEST_VERBOSE"):
            traceback.print_exc()


def eq(actual, expected, what=""):
    if actual != expected:
        raise AssertionError(f"{what} 期望 {expected!r}，实际 {actual!r}")


def truth(value, what=""):
    if not value:
        raise AssertionError(f"{what} 期望为真，实际 {value!r}")


# ---------------------------------------------------------------- 样本构造
def build_sample_xlsx(path):
    """构造一张与真实统计表结构一致的样本：
    A=姓名 B=基础分 C=基础分鉴定 D=3-4月加分描述 E=3-4月表扬信 F=鉴定
    G/H=扣分 I=5-8月加分描述 J=5-8月表扬信 K=鉴定 L/M=扣分 N=备注 O=序号 P=德育总分
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "大三下"

    ws["A1"] = "2025-2026学年第2学期德育分统计表"
    ws.merge_cells("A1:P1")
    ws["A2"] = "班级：农学23-1班"
    ws.merge_cells("A2:B2")

    # 第 3 行：大类（合并单元格，和真实表一致）
    for rng, text in (("A3:A4", "姓名"), ("B3:C3", "基础分"),
                      ("D3:F3", "3-4月加分项"), ("G3:H3", "3-4月扣分项"),
                      ("I3:K3", "5-8月加分项"), ("L3:M3", "5-8月扣分项"),
                      ("N3:N4", "备注"), ("O3:O4", "序号"), ("P3:P4", "德育总分")):
        ws[rng.split(":")[0]] = text
        ws.merge_cells(rng)

    # 第 4 行：细分
    for cell, text in (("C4", "鉴定"), ("D4", "3-4月"), ("E4", "表扬信"),
                       ("F4", "鉴定"), ("G4", "鉴定"), ("H4", "鉴定"),
                       ("I4", "5-8月"), ("J4", "表扬信"), ("K4", "鉴定"),
                       ("L4", "鉴定"), ("M4", "鉴定")):
        ws[cell] = text

    # 第 5 行起：数据
    ws.append(["张三", "基础分+72", "72", "院3月学霸宿舍奖+2",
               "3.6志愿服务表扬信+0.5\n3.10表彰大会表扬信+0.5", "", "", "", "",
               "7.5宿舍搬迁表扬信+0.5", "", "", "", "班长", "1", 82])
    ws.append(["李四", "基础分+75", "75", "", "", "", "", "", "", "", "", "",
               "", "", "2", None])
    ws.append(["王五", "基础分+60", "60", "", "3.8器材搬运表扬信+0.5", "", "", "", "",
               "6.5反诈活动表扬信+0.5", "", "", "", "", "3", -5])
    # 干扰行：不该被当成学生
    ws.append(["合计", "", "207", "", "", "", "", "", "", "", "", "", "", "", "", 207])
    wb.save(path)


def build_sample_template(path):
    """构造一个 5 列 Word 模板：标题行 + 表头行 + 空数据行。"""
    doc = Document()
    table = doc.add_table(rows=3, cols=5)
    cell = table.rows[0].cells[0]
    cell.text = "XX班 2025-2026 第二学期德育分确认签字表"
    cell.add_paragraph("班级")
    for i, head in enumerate(["姓名", "德育分", "表扬信", "确认签字", "备注"]):
        table.rows[1].cells[i].text = head
    doc.save(path)


# ---------------------------------------------------------------- 各项检查
def test_detect_columns():
    wb = openpyxl.load_workbook(os.path.join(TMP, "sample.xlsx"), data_only=True)
    ws = wb.active
    headers = main.scan_headers(ws)
    name_col, score_col, praise_cols = main.detect_columns(headers)
    wb.close()
    eq(name_col, 1, "姓名列")
    eq(score_col, 16, "总分列")
    eq(praise_cols, [5, 10], "表扬信列")


def test_month_groups():
    wb = openpyxl.load_workbook(os.path.join(TMP, "sample.xlsx"), data_only=True)
    ws = wb.active
    headers = main.scan_headers(ws)
    _, _, praise_cols = main.detect_columns(headers)
    groups = main.group_praise_by_month(ws, praise_cols)
    wb.close()
    eq(groups, [("3-4月", [5]), ("5-8月", [10])], "月份分组")


def test_title_and_class():
    wb = openpyxl.load_workbook(os.path.join(TMP, "sample.xlsx"), data_only=True)
    ws = wb.active
    title = main.extract_title(ws, "大三下")
    cls = main.extract_class(ws)
    wb.close()
    eq(title, "大三下 2025-2026 第二学期德育分确认签字表", "标题")
    eq(cls, "农学23-1班", "班级")


def test_extract_students():
    wb = openpyxl.load_workbook(os.path.join(TMP, "sample.xlsx"), data_only=True)
    ws = wb.active
    students = main.extract_students(ws, 1, 16, [5, 10])
    wb.close()
    by = {s["name"]: s for s in students}
    eq(len(students), 3, "学生人数（合计行应被排除）")
    eq(by["张三"]["praise"], 3, "张三表扬信条数")
    eq(by["李四"]["praise"], 0, "李四表扬信条数")
    eq(by["王五"]["praise"], 2, "王五表扬信条数")
    eq(by["张三"]["score"], "82", "张三分数字符串")


def test_validate_students():
    wb = openpyxl.load_workbook(os.path.join(TMP, "sample.xlsx"), data_only=True)
    ws = wb.active
    students = main.extract_students(ws, 1, 16, [5, 10])
    wb.close()
    problems = main.validate_students(students)
    kinds = {(p[0], p[2]) for p in problems}
    eq(len(problems), 2, "问题条数（李四空分数 + 王五负分）")
    truth(("empty", "李四") in kinds, "应检出李四分数为空")
    truth(("neg", "王五") in kinds, "应检出王五负分")


def test_cert_and_praise_extract():
    data = main.extract_persons_from_workbook(os.path.join(TMP, "sample.xlsx"))
    truth("张三" in data, "应提取到张三")
    certs, praises = data["张三"]
    eq(len(certs), 1, "张三证书条数")
    eq(certs[0], "院3月学霸宿舍奖", "张三证书内容（应去掉 +2）")
    eq(len(praises), 3, "张三表扬信条数（跨两个月份段合并）")


def test_add_year_to_date():
    sem = "2025-2026学年第2学期"
    eq(main.add_year_to_date("3.6志愿服务", sem), "2026年3月6日志愿服务", "补年份-月.日")
    eq(main.add_year_to_date("五月劳动周", sem), "2026年5月劳动周", "补年份-中文月份")
    eq(main.add_year_to_date("9月义务劳动", "2025-2026学年第一学期"),
       "2025年9月义务劳动", "第一学期 9 月用起始年")
    eq(main.add_year_to_date("3.6志愿服务", "某班统计表"),
       "3月6日志愿服务", "无法解析学期时仍规范化日期、不补年")


def test_no_semester_dates():
    """文件名里没有学期时：日期仍要规范化并留在句首，且能被界面提示出来。"""
    sem = "XX班20XX-20XX学年第X学期德育分统计表"
    eq(main.add_year_to_date("4.9关工委工作礼仪活动", sem),
       "4月9日关工委工作礼仪活动", "无学期时的 月.日")
    eq(main.add_year_to_date("10月22号百团大战志愿活动", sem),
       "10月22日百团大战志愿活动", "无学期时的 月日号")
    eq(main.add_year_to_date("大二：11.25农学院体测活动", sem),
       "11月25日农学院体测活动", "先去掉「大二：」再规范化日期")
    eq(main.to_praise_sentence("4月9日关工委工作礼仪活动"),
       "4月9日参与关工委工作礼仪活动", "没有年份时日期仍在句首")
    eq(main.to_cert_sentence("11月25日院太极训练证书"),
       "11月25日荣获院太极训练", "证书句子同样")

    # 列表提示：识别不出学期时要能收集出来
    no_sem = []
    rows, missing, multi = main.collect_person_rows(
        [os.path.join(TMP, "农学23-1班2025-2026学年第2学期德育分统计表.xlsx")],
        ["张三"], no_semester=no_sem)
    eq(no_sem, [], "文件名带学期时不应报警")

    bad_name = os.path.join(TMP, "XX班20XX-20XX学年第X学期德育分统计表.xlsx")
    shutil.copy(os.path.join(TMP, "sample.xlsx"), bad_name)
    no_sem2 = []
    main.collect_person_rows([bad_name], ["张三"], no_semester=no_sem2)
    eq(no_sem2, ["XX班20XX-20XX学年第X学期德育分统计表.xlsx"], "应收集识别不出学期的文件")


def test_lead_date_formats():
    """开头带年份的各种写法（2026.5 / 2026.4.2 / 2026-5-6）都要统一成中文日期并留在句首。"""
    sem = "2025-2026学年第2学期"
    eq(main.add_year_to_date("2026.5创意视频拍摄表扬信", sem),
       "2026年5月创意视频拍摄表扬信", "年.月 写法")
    eq(main.add_year_to_date("2026.4.2春季穿搭", sem),
       "2026年4月2日春季穿搭", "年.月.日 写法")
    eq(main.add_year_to_date("2026-5-6志愿服务", sem),
       "2026年5月6日志愿服务", "年-月-日 写法")
    eq(main.to_praise_sentence("2026年5月创意视频拍摄表扬信"),
       "2026年5月参与创意视频拍摄活动", "表扬信句子日期在句首")
    eq(main.to_cert_sentence("2026年4月2日春季穿搭证书"),
       "2026年4月2日荣获春季穿搭", "证书句子日期在句首")
    eq(main.add_year_to_date("9.9宿舍搬迁表扬信", sem),
       "2026年9月9日宿舍搬迁表扬信", "不带年份的写法仍按学期补年")


def test_sentence_trailing_semicolon():
    """句式转换后每一条都以「；」结尾（和源表书写习惯一致）。"""
    rows = [("张三", "2026年3月院级学霸宿舍奖\n2026年6月省志愿者证书",
             "参加校园积雪清扫活动；\n2026年7月5日宿舍搬迁表扬信",
             "2025-2026学年第2学期", False)]
    src = os.path.join(TMP, "分号测试.xlsx")
    dst = os.path.join(TMP, "分号测试_句式.xlsx")
    main.write_person_result_xlsx(rows, src)
    main.convert_result_to_sentences(src, dst)
    wb = openpyxl.load_workbook(dst)
    ws = wb.active
    certs = str(ws.cell(2, 2).value).split("\n")
    praises = str(ws.cell(2, 3).value).split("\n")
    wb.close()
    for line in certs + praises:
        truth(line.endswith("；"), f"每行应以「；」结尾，实际：{line!r}")
    eq(certs[0], "2026年3月荣获院级学霸宿舍奖；", "证书句子")
    eq(praises[0], "参与校园积雪清扫活动；", "表扬信句子")


def test_verb_before_date():
    """「参与2026年4.10…」这类动词在日期前面的写法，日期要提到句首。"""
    eq(main.to_praise_sentence("参与2026年4.10拟发展对象材料审核工作表扬信"),
       "2026年4月10日参与拟发展对象材料审核工作", "参与在日期前 + 年.月.日")
    eq(main.to_praise_sentence("参与2026年7.5宿舍搬迁表扬信"),
       "2026年7月5日参与宿舍搬迁活动", "参与在日期前 + 年.月")
    eq(main.to_praise_sentence(
        "参与2026年3月河南农业大学举办“挑战杯”大学生创业计划竞赛决赛志愿活动表扬信"),
       "2026年3月参与河南农业大学举办“挑战杯”大学生创业计划竞赛决赛志愿活动",
       "参与在「2026年3月」前")
    eq(main.normalize_lead_date("2026年4.10拟发展对象材料审核工作"),
       "2026年4月10日拟发展对象材料审核工作", "年.月.日 混写规范化")
    eq(main.to_cert_sentence("参与2026年4月10日志愿服务证书"),
       "2026年4月10日参与志愿服务", "证书里的动词要保留、日期提前")


def test_praise_verb_prefix():
    """源数据本身带「参加/参与」时，不能拼成「参与参加…」；结尾标点也要去掉。"""
    eq(main.to_praise_sentence("参加校园积雪清扫活动；表扬信"),
       "参与校园积雪清扫活动", "参加 → 换成参与，且不留分号")
    eq(main.to_praise_sentence("参与参加三下乡出征仪式授旗活动表扬信"),
       "参与三下乡出征仪式授旗活动", "叠加的动词只留一个")
    eq(main.to_praise_sentence("参加百团大战会场布置志愿者活动表扬信"),
       "参与百团大战会场布置志愿者活动", "普通「参加」条目")
    eq(main.to_cert_sentence("参与发表论文"),
       "参与发表论文", "证书里的参与类保持原样")
    eq(main.to_cert_sentence("2026年3月校级新闻；"),
       "2026年3月荣获校级新闻", "证书结尾标点也要去掉")


def test_cert_sentences():
    eq(main.to_cert_sentence("2026年3月参与发表论文"),
       "2026年3月参与发表论文", "参与类不加荣获")
    eq(main.to_cert_sentence("2026年6月省志愿者证书"),
       "2026年6月被评为省志愿者", "身份类用被评为")
    eq(main.to_cert_sentence("2026年3月校一等奖"),
       "2026年3月荣获校一等奖", "获奖类用荣获")
    eq(main.to_cert_sentence("2026年3月优秀团员"),
       "2026年3月被评为优秀团员", "荣誉称号用被评为")


def test_praise_sentences():
    eq(main.to_praise_sentence("2026年7月5日宿舍搬迁表扬信"),
       "2026年7月5日参与宿舍搬迁活动", "表扬信加参与与活动")
    eq(main.to_praise_sentence("2026年6月年级大会表扬信"),
       "2026年6月参与年级大会", "会议类不加活动")
    eq(main.to_praise_sentence("2026年6月参与义务劳动表扬信"),
       "2026年6月参与义务劳动活动", "不重复参与")


def test_parse_name_text():
    names, dup = main.parse_name_text("张三\n李四\n张三\n张三（男）\n李四")
    eq(names, ["张三", "李四"], "去重后的名单")
    eq(dup, 3, "被忽略的重复个数")
    names, dup = main.parse_name_text("甲，乙、丙;丁 戊")
    eq(names, ["甲", "乙", "丙", "丁", "戊"], "多种分隔符")
    eq(dup, 0, "无重复时计数为 0")


def test_collect_and_write_then_convert():
    xlsx = os.path.join(TMP, "农学23-1班2025-2026学年第2学期德育分统计表.xlsx")
    rows, missing, multi = main.collect_person_rows([xlsx], ["张三", "李四", "查无此人"])
    eq(len(rows), 2, "结果行数")
    eq(missing, ["查无此人"], "未找到的人员")
    eq(multi, [], "多源表人员")

    out1 = os.path.join(TMP, "提取结果.xlsx")
    main.write_person_result_xlsx(rows, out1)
    wb = openpyxl.load_workbook(out1)
    ws = wb.active
    eq([c.value for c in ws[1]], ["姓名", "证书", "表扬信", "来源学期"], "结果表头")
    eq(ws.cell(2, 4).value, "2025-2026学年第2学期", "来源学期列")
    wb.close()

    out2 = os.path.join(TMP, "提取结果_句式.xlsx")
    n = main.convert_result_to_sentences(out1, out2)
    eq(n, 2, "句式转换行数")
    wb = openpyxl.load_workbook(out2)
    ws = wb.active
    eq(ws.cell(2, 4).value, "2025-2026学年第2学期", "句式转换后仍保留来源学期")
    body = str(ws.cell(2, 2).value or "")
    truth("荣获" in body or "参与" in body, "证书应已转成句子")
    wb.close()


def test_template_structure():
    ok, msg = main.check_template_structure(os.path.join(TMP, "template.docx"))
    truth(ok, f"模板校验应通过（{msg}）")
    ok, msg = main.check_template_structure(os.path.join(TMP, "不存在.docx"))
    truth(not ok, "不存在的模板应校验失败")


def test_fill_word():
    tpl = os.path.join(TMP, "template.docx")
    out = os.path.join(TMP, "签字表.docx")
    wb = openpyxl.load_workbook(os.path.join(TMP, "sample.xlsx"), data_only=True)
    ws = wb.active
    students = main.extract_students(ws, 1, 16, [5, 10])
    title = main.extract_title(ws, "大三下")
    cls = main.extract_class(ws)
    wb.close()

    main.fill_word(tpl, out, students, title, cls)
    table = Document(out).tables[0]
    eq(len(table.rows), 5, "生成表格行数（标题+表头+3人）")
    paras = [p.text for p in table.rows[0].cells[0].paragraphs]
    eq(paras[0], title, "标题行")
    truth("班级：农学23-1班" in paras, "班级应写入标题行")
    eq([c.text for c in table.rows[2].cells], ["张三", "82", "3", "", ""], "第一行数据")
    eq([c.text for c in table.rows[4].cells], ["王五", "-5", "2", "", ""], "第三行数据")


def test_unique_path():
    base = os.path.join(TMP, "同名.xlsx")
    eq(main.unique_path_on_disk(base), base, "文件不存在时保持原名")
    open(base, "w").close()
    second = main.unique_path_on_disk(base)
    eq(os.path.basename(second), "同名_1.xlsx", "已存在时加 _1")
    open(second, "w").close()
    eq(os.path.basename(main.unique_path_on_disk(base)), "同名_2.xlsx", "再次冲突时加 _2")


def test_config_roundtrip():
    cfg_path = os.path.join(TMP, "config.json")
    orig = main.config_file
    main.config_file = lambda: cfg_path
    try:
        eq(main.load_config(), {}, "无配置时返回空字典")
        truth(main.save_config({"template_path": r"C:\t.docx", "output_dir": TMP}),
              "保存应成功")
        cfg = main.load_config()
        eq(cfg["template_path"], r"C:\t.docx", "模板路径往返")
        eq(cfg["output_dir"], TMP, "输出目录往返")
    finally:
        main.config_file = orig


def test_unreadable_source_reported():
    """混入损坏的源表时，必须把它报出来，不能静默跳过。"""
    good = os.path.join(TMP, "农学23-1班2025-2026学年第2学期德育分统计表.xlsx")
    bad = os.path.join(TMP, "损坏的统计表.xlsx")
    with open(bad, "w", encoding="utf-8") as f:
        f.write("这不是一个真正的 Excel 文件")
    failed = []
    rows, missing, multi = main.collect_person_rows([good, bad], ["张三"], failed=failed)
    eq(len(rows), 1, "正常源表仍应提取到 1 行")
    eq(len(failed), 1, "应记录 1 个读取失败的源表")
    eq(failed[0][0], "损坏的统计表.xlsx", "失败文件名")
    truth("损坏" in failed[0][1] or "Excel" in failed[0][1], "失败原因应为中文可读提示")


def test_friendly_excel_errors():
    """打不开的文件要给出中文原因，而不是 openpyxl 的英文异常。"""
    bad = os.path.join(TMP, "坏文件.xlsx")
    with open(bad, "w", encoding="utf-8") as f:
        f.write("not a workbook")
    try:
        main.load_workbook_checked(bad)
        raise AssertionError("损坏文件应当报错")
    except ValueError as exc:
        truth("损坏" in str(exc) or "Excel" in str(exc), f"中文提示，实际：{exc}")

    old = os.path.join(TMP, "老格式.xls")
    with open(old, "wb") as f:
        f.write(b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1")
    try:
        main.load_workbook_checked(old)
        raise AssertionError(".xls 应当报错")
    except ValueError as exc:
        truth("xls" in str(exc).lower() or "另存为" in str(exc), f"应提示另存为 xlsx，实际：{exc}")


def test_process_one_excel():
    outdir = os.path.join(TMP, "out")
    os.makedirs(outdir, exist_ok=True)
    xlsx = os.path.join(TMP, "农学23-1班2025-2026学年第2学期德育分统计表.xlsx")
    out, n = main.process_one_excel(xlsx, os.path.join(TMP, "template.docx"), outdir)
    eq(n, 3, "生成人数")
    truth(os.path.exists(out), "输出文件应存在")
    eq(os.path.basename(out),
       "农学23-1班2025-2026学年第2学期德育分签字表.docx", "输出文件名")


def test_process_one_excel_month_filter():
    """只统计 3-4月：张三应剩 2 条表扬信。"""
    outdir = os.path.join(TMP, "out2")
    os.makedirs(outdir, exist_ok=True)
    xlsx = os.path.join(TMP, "农学23-1班2025-2026学年第2学期德育分统计表.xlsx")
    out, n = main.process_one_excel(
        xlsx, os.path.join(TMP, "template.docx"), outdir, selected_months={"3-4月"})
    eq(n, 3, "人数")
    rows = Document(out).tables[0].rows
    eq(rows[2].cells[2].text, "2", "张三只统计 3-4月 后的表扬信条数")


# ---------------------------------------------------------------- 鉴定栏核对
def build_judge_sheet():
    """构造一张用于鉴定栏校对的样本：含表扬信封顶、新闻报道封顶、逗号笔误、附加分、未填。"""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "大三上"
    ws["A1"] = "2026-2027学年第一学期德育分评分统计"
    ws.merge_cells("A1:P1")
    ws["A2"] = "班级：农学24-1班"
    for rng, text in (("A3:A4", "姓名"), ("B3:C3", "基础分"),
                      ("D3:F3", "9-10月加分项"), ("G3:H3", "9-10月扣分项"),
                      ("I3:K3", "X-X月加分项"), ("L3:M3", "X-X月扣分项"),
                      ("N3:O3", "附加分"), ("P3:P4", "总分")):
        ws[rng.split(":")[0]] = text
        ws.merge_cells(rng)
    for c, t in (("C4", "鉴定"), ("D4", "9-10月"), ("E4", "表扬信"), ("F4", "鉴定"),
                 ("G4", "鉴定"), ("H4", "鉴定"), ("I4", "X-X月"), ("J4", "表扬信"),
                 ("K4", "鉴定"), ("L4", "鉴定"), ("M4", "鉴定"), ("O4", "鉴定")):
        ws[c] = t

    rows = [
        # 甲：第一段 12 封（封顶 10 封=5 分），第二段 3 封额度已用完；F 现值 2.5 应为 7.5
        ["甲", "基础分+70", 70, "校级证书+2.5", "\n".join(["表扬信"] * 12), 2.5,
         None, None, None, "\n".join(["表扬信"] * 3), 0, None, None, None, None,
         "=C5+F5-H5+K5-M5+O5"],
        # 乙：两条新闻报道各 3 分 → 封顶 5，F 现值正确
        ["乙", "基础分+70", 70, "新闻报道+3\n新闻报道+3", None, 5,
         None, None, None, None, 0, None, None, None, None,
         "=C6+F6-H6+K6-M6+O6"],
        # 丙：逗号笔误 +0,5，一封 0.5
        ["丙", "基础分+70", 70, None, "参加活动表扬信+0,5", 0.5,
         None, None, None, None, 0, None, None, None, None,
         "=C7+F7-H7+K7-M7+O7"],
        # 丁：附加分说明里两个分值
        ["丁", "基础分+70", 70, None, None, 0,
         None, None, None, None, 0, None, None, "班长+7 六级证书+4", 11,
         "=C8+F8-H8+K8-M8+O8"],
        # 戊：鉴定栏空着但应有 4 封表扬信 × 0.5 = 2
        ["戊", "基础分+70", 70, None, "\n".join(["表扬信"] * 4), None,
         None, None, None, None, 0, None, None, None, None,
         "=C9+F9-H9+K9-M9+O9"],
        # 己：附加分里有认不出来的条目（献血），该行附加分不参与比对
        ["己", "基础分+70", 70, None, None, 0,
         None, None, None, None, 0, None, None, "10.3献血+0.5", 0.5,
         "=C10+F10-H10+K10-M10+O10"],
    ]
    for r in rows:
        ws.append(r)
    return wb, ws


def test_judge_parse_rules():
    """分值解析：普通加分、新闻报道封顶、逗号笔误、封数统计。"""
    eq(judge_audit.parse_points("班长+7 六级证书+4"), 11, "两个分值累加")
    eq(judge_audit.parse_points("参加活动+0,5"), 0.5, "半角逗号小数")
    eq(judge_audit.parse_points("参加活动+0，5"), 0.5, "全角逗号小数")
    eq(judge_audit.parse_points("上限五分"), 0, "没有 + 数字时不计分")
    eq(judge_audit.activity_score("校级证书+2.5\n校级证书+2.5"), 5, "普通活动不封顶")
    eq(judge_audit.activity_score("新闻报道+3\n新闻报道+3"), 5, "新闻报道封顶 5")
    eq(judge_audit.activity_score("新闻报道+3\n新闻报道+3\n校级证书+2"), 7,
       "新闻封顶 5 后，其余活动照计（5+2）")
    eq(judge_audit.addon_score("文体委员+4\n英语四级证书+2\n青禾艺术团工作人员+5"), 7,
       "附加分：职务取最高 5 + 证书 2")
    eq(judge_audit.addon_score("班长+7\n团支书+9"), 9, "附加分：多个职务取最高")
    eq(judge_audit.addon_score("英语四级证书+2\n计算机二级证书+2"), 4,
       "附加分：证书累加")
    eq(judge_audit.addon_score("班长+7\n神秘项+99"), 7, "附加分：认不出来的不计分")
    eq(judge_audit.classify_addon("优秀志愿者证书+1"), "cert",
       "带「证书」二字的按证书算，不算职务")
    eq(judge_audit.addon_score("文体委员+4 英语四级证书+2 青禾艺术团工作人员+5"), 7,
       "附加分：同一行用空格隔开也能拆开")
    eq(judge_audit.praise_count("\n".join(["x"] * 12)), 12, "按换行计封数")
    eq(judge_audit.praise_count("a\n\nb\n"), 2, "空行不计")
    eq(judge_audit.praise_count(
        "10.31军训表扬信+0.5\n10.28彩旗队表扬信+0.5" + " " * 30 +
        "10.30彩旗队表扬信+0.5"), 3, "同一行用空格隔开的多条也要算封")
    eq(judge_audit.praise_count(
        "9.20 参加活动表扬信+0.5\n9.21 参加活动表扬信+0.5"), 2,
       "单个空格不拆条")


def test_judge_audit_and_fix():
    """整表核对：表扬信学期封顶、未填报错、附加分、修正后归零。"""
    wb, ws = build_judge_sheet()
    path = os.path.join(TMP, "鉴定栏样本.xlsx")
    wb.save(path)
    wb.close()

    wb = openpyxl.load_workbook(path, data_only=True)
    cols, rows, problems = judge_audit.audit_sheet(wb.active)
    wb.close()
    eq(len(rows), 6, "学生数")
    by = {r["name"]: r for r in rows}

    f_col = cols["groups"][0]["judge"]
    k_col = cols["groups"][1]["judge"] if len(cols["groups"]) > 1 else None
    o_col = cols["add_judge"]
    eq(by["甲"]["expect"][f_col], 7.5, "甲：2.5 活动 + 封顶 5 分表扬信")
    eq(by["乙"]["expect"][f_col], 5, "乙：新闻报道封顶 5")
    eq(by["丙"]["expect"][f_col], 0.5, "丙：+0,5 计 0.5")
    eq(by["丁"]["expect"][o_col], 11, "丁：附加分 7+4")
    eq(by["戊"]["expect"][f_col], 2, "戊：4 封 = 2 分")
    truth(by["己"].get("addon_unknown"), "己：应识别出未计入的「献血」条目")

    issues = judge_audit.collect_issues(rows)
    eq(len(issues), 2, f"应只有 2 处差异，实际：{issues}")
    reasons = {by_row["name"]: reason for by_row in rows
               for (row, _c, _cur, _exp, reason) in issues
               if row == by_row["row"]}
    truth("甲" in reasons and "少计" in reasons["甲"], "甲的差异应提示少计")
    truth("戊" in reasons and "未填" in reasons["戊"], "戊的差异应提示未填")

    dst = os.path.join(TMP, "鉴定栏样本_已修正.xlsx")
    judge_audit.apply_fixes(path, dst, [(r, c, exp) for r, c, _cur, exp, _ in issues])
    wb = openpyxl.load_workbook(dst, data_only=True)
    cols2, rows2, _ = judge_audit.audit_sheet(wb.active)
    wb.close()
    eq(len(judge_audit.collect_issues(rows2)), 0, "修正后不应再有差异")


# ---------------------------------------------------------------- 界面（线程调度）
def test_audit_tab_batch():
    """鉴定栏核对页：多文件去重、批量检查、明细联动、按输出目录批量修正。"""
    if os.environ.get("SKIP_GUI_TESTS"):
        raise Skip("已设置 SKIP_GUI_TESTS，跳过图形测试")
    try:
        app = main.App()
    except Exception as exc:
        raise Skip(f"无图形环境：{exc}")

    cfg_path = os.path.join(TMP, "gui_config_audit.json")
    orig_cfg = main.config_file
    orig_mb = (main.messagebox.askyesno, main.messagebox.showinfo,
               main.messagebox.showwarning)
    main.config_file = lambda: cfg_path
    main.messagebox.askyesno = lambda *a, **k: True
    main.messagebox.showinfo = lambda *a, **k: None
    main.messagebox.showwarning = lambda *a, **k: None
    try:
        with open(cfg_path, "w", encoding="utf-8") as f:
            json.dump({"about_seen": True}, f, ensure_ascii=False)
        app.withdraw()
        app._switch_tab("audit")

        wb, _ws = build_judge_sheet()
        p1 = os.path.join(TMP, "核对甲.xlsx")
        wb.save(p1)
        wb.close()
        p2 = os.path.join(TMP, "核对乙.xlsx")
        shutil.copy(p1, p2)

        app._audit_add_files([p1, p2, p1])       # 第三个是重复的
        eq(len(app.audit_tree.get_children()), 2, "重复文件应去重")

        app.audit_run_all()
        vals = [app.audit_tree.item(i, "values")
                for i in app.audit_tree.get_children()]
        eq(len(vals), 2, "两个文件都应检查")
        truth(all(v[2] == "2" for v in vals), f"每个文件应有 2 处差异：{vals}")

        app.audit_tree.selection_set(app.audit_tree.get_children()[0])
        app._audit_show_detail()
        eq(len(app.audit_detail.get_children()), 6, "明细应有 6 行")

        # 错误要高亮：红底加粗标签 + 结论条写明要改哪些文件
        bad_detail = [i for i in app.audit_detail.get_children()
                      if "bad" in app.audit_detail.item(i, "tags")]
        eq(len(bad_detail), 2, "明细里应有 2 行标红")
        bad_files = [i for i in app.audit_tree.get_children()
                     if "bad" in app.audit_tree.item(i, "tags")]
        eq(len(bad_files), 2, "列表里应有 2 个文件标红")
        summary = app.audit_summary.cget("text")
        truth("需要修正" in summary, f"结论条应提示需要修正：{summary}")
        first_bad = app.audit_detail.item(bad_detail[0], "values")
        truth("→" in first_bad[-1], f"结果列应写明改什么：{first_bad[-1]}")

        outdir = os.path.join(TMP, "audit_out")
        app.audit_output.set(outdir)
        app.audit_fix()
        made = [f for f in os.listdir(outdir) if f.endswith("_已修正.xlsx")]
        eq(len(made), 2, f"应生成 2 个已修正文件，实际：{made}")
    finally:
        main.config_file = orig_cfg
        (main.messagebox.askyesno, main.messagebox.showinfo,
         main.messagebox.showwarning) = orig_mb
        try:
            app.destroy()
        except Exception:
            pass


def test_batch_threads():
    """批量生成走后台线程 + 结果队列：两个同名源文件应各自成功并自动错开文件名。
    需要图形环境；CI 上通过 SKIP_GUI_TESTS=1 跳过。"""
    if os.environ.get("SKIP_GUI_TESTS"):
        raise Skip("已设置 SKIP_GUI_TESTS，跳过图形测试")
    try:
        app = main.App()
    except Exception as exc:
        raise Skip(f"无图形环境：{exc}")

    cfg_path = os.path.join(TMP, "gui_config.json")
    orig_cfg = main.config_file
    orig_mb = (main.messagebox.askyesno, main.messagebox.showwarning,
               main.messagebox.showinfo)
    main.config_file = lambda: cfg_path
    main.messagebox.askyesno = lambda *a, **k: True
    main.messagebox.showwarning = lambda *a, **k: None
    main.messagebox.showinfo = lambda *a, **k: None
    try:
        with open(cfg_path, "w", encoding="utf-8") as f:
            json.dump({"template_path": os.path.join(TMP, "template.docx"),
                       "about_seen": True}, f, ensure_ascii=False)
        app.withdraw()
        app._batch_finish = lambda cancelled=False: None  # 不弹完成窗
        outdir = os.path.join(TMP, "batch")
        os.makedirs(outdir, exist_ok=True)
        app.template_path.set(os.path.join(TMP, "template.docx"))
        app.output_dir.set(outdir)

        xlsx = os.path.join(TMP, "农学23-1班2025-2026学年第2学期德育分统计表.xlsx")
        app.batch_items = {}
        app.batch_info = {}
        for _ in range(2):  # 故意加同一个文件两次，验证同名自动错开
            item = app.batch_tree.insert(
                "", "end", values=(os.path.basename(xlsx), "", ""))
            app.batch_items[item] = xlsx
            app.batch_info[item] = {"students": [1], "problems": []}

        app.batch_generate()
        t0 = time.time()
        while time.time() - t0 < 60:
            app.update()
            time.sleep(0.05)
            if not getattr(app, "_gen_active", True):
                break
        eq(app._gen_results["ok"], 2, "批量成功数")
        eq(app._gen_results["fail"], 0, "批量失败数")
        eq(len(os.listdir(outdir)), 2, "输出文件数（同名应自动加 _2）")
    finally:
        main.config_file = orig_cfg
        (main.messagebox.askyesno, main.messagebox.showwarning,
         main.messagebox.showinfo) = orig_mb
        try:
            app.destroy()
        except Exception:
            pass


def test_batch_analyze_threads():
    """批量分析走后台线程：加入文件后列表能自动填好工作表、人数、状态。"""
    if os.environ.get("SKIP_GUI_TESTS"):
        raise Skip("已设置 SKIP_GUI_TESTS，跳过图形测试")
    try:
        app = main.App()
    except Exception as exc:
        raise Skip(f"无图形环境：{exc}")

    cfg_path = os.path.join(TMP, "gui_config2.json")
    orig_cfg = main.config_file
    main.config_file = lambda: cfg_path
    try:
        with open(cfg_path, "w", encoding="utf-8") as f:
            json.dump({"template_path": os.path.join(TMP, "template.docx"),
                       "about_seen": True}, f, ensure_ascii=False)
        app.withdraw()
        xlsx = os.path.join(TMP, "农学23-1班2025-2026学年第2学期德育分统计表.xlsx")
        app._batch_add_files([xlsx])
        t0 = time.time()
        while time.time() - t0 < 30:
            app.update()
            time.sleep(0.03)
            if not app._analyze_queue and app._analyze_pending == 0:
                break
        eq(len(app.batch_info), 1, "应完成 1 个文件的分析")
        vals = app.batch_tree.item(app.batch_tree.get_children()[0], "values")
        eq(vals[1], "大三下", "工作表")
        eq(vals[2], "3", "人数")
        eq(vals[3], "就绪（2项异常，双击查看）",
           "状态（样本里李四空分数、王五负分）")
    finally:
        main.config_file = orig_cfg
        try:
            app.destroy()
        except Exception:
            pass


def test_output_dir_required():
    """没选输出目录时，批量生成必须被拦下，不再默认放到数据源目录。"""
    if os.environ.get("SKIP_GUI_TESTS"):
        raise Skip("已设置 SKIP_GUI_TESTS，跳过图形测试")
    try:
        app = main.App()
    except Exception as exc:
        raise Skip(f"无图形环境：{exc}")

    cfg_path = os.path.join(TMP, "gui_config3.json")
    orig_cfg = main.config_file
    warns = []
    orig_warn = main.messagebox.showwarning
    main.config_file = lambda: cfg_path
    main.messagebox.showwarning = lambda *a, **k: warns.append(a[0] if a else "")
    try:
        with open(cfg_path, "w", encoding="utf-8") as f:
            json.dump({"template_path": os.path.join(TMP, "template.docx"),
                       "about_seen": True}, f, ensure_ascii=False)
        app.withdraw()
        app.output_dir.set("")
        app.batch_generate()
        truth("请先选择输出目录" in warns,
              f"应提示先选输出目录，实际弹窗：{warns}")
        truth(not getattr(app, "_gen_active", False), "不应开始生成")

        # 完成弹窗：单文件页要带「打开文件 / 打开文件夹 / 关闭」三个按钮
        top = app._show_done(
            "生成成功", "测试",
            ["已生成，共 3 人。", r"C:\temp\a.docx"],
            [("打开文件", lambda t: None, False),
             ("打开文件夹", lambda t: None, False),
             ("关闭", lambda t: t.destroy(), True)])
        texts = []

        def walk(w):
            for c in w.winfo_children():
                try:
                    t = c.cget("text")
                    if t:
                        texts.append(str(t))
                except Exception:
                    pass
                walk(c)

        walk(top)
        for label in ("打开文件", "打开文件夹", "关闭"):
            truth(label in texts, f"完成弹窗应包含「{label}」按钮")
        top.destroy()
    finally:
        main.config_file = orig_cfg
        main.messagebox.showwarning = orig_warn
        try:
            app.destroy()
        except Exception:
            pass



# ---------------------------------------------------------------- 主流程
def main_entry():
    global TMP
    # 用脚本目录下的固定工作目录，而不是系统临时目录：
    # 某些受限环境（如沙箱）不允许往 %TEMP% 里写，放在项目内哪儿都能跑。
    TMP = os.path.join(os.path.dirname(os.path.abspath(__file__)), "_test_work")
    shutil.rmtree(TMP, ignore_errors=True)
    os.makedirs(TMP, exist_ok=True)
    try:
        build_sample_xlsx(os.path.join(TMP, "sample.xlsx"))
        build_sample_template(os.path.join(TMP, "template.docx"))
        # 提取流程按文件名解析学期，另存一份带学期的文件名
        shutil.copy(os.path.join(TMP, "sample.xlsx"),
                    os.path.join(TMP, "农学23-1班2025-2026学年第2学期德育分统计表.xlsx"))

        print("\n=== 表头识别与数据提取 ===")
        check("识别姓名/总分/表扬信列", test_detect_columns)
        check("月份段分组", test_month_groups)
        check("标题与班级识别", test_title_and_class)
        check("学生数据提取（合计行不应混入）", test_extract_students)
        check("分数校验（空分数/负分）", test_validate_students)

        print("\n=== 证书 / 表扬信 ===")
        check("跨月份提取证书与表扬信", test_cert_and_praise_extract)
        check("日期补年份", test_add_year_to_date)
        check("开头自带年份的日期写法", test_lead_date_formats)
        check("识别不出学期时的日期处理", test_no_semester_dates)
        check("证书句子措辞", test_cert_sentences)
        check("表扬信句子措辞", test_praise_sentences)
        check("参与/参加 不重复、结尾标点清理", test_praise_verb_prefix)
        check("「参与」在日期前面时日期提前", test_verb_before_date)
        check("每条句子以「；」结尾", test_sentence_trailing_semicolon)

        print("\n=== 名单与结果文件 ===")
        check("名单解析与去重", test_parse_name_text)
        check("提取结果写出 / 句式转换保留来源学期", test_collect_and_write_then_convert)
        check("同名文件自动另存 _1/_2", test_unique_path)
        check("损坏源表会被报出来", test_unreadable_source_reported)
        check("打不开的文件给中文提示", test_friendly_excel_errors)

        print("\n=== Word 生成 ===")
        check("模板结构校验", test_template_structure)
        check("灌入学生数据", test_fill_word)
        check("端到端生成（含文件名替换）", test_process_one_excel)
        check("按月份段生成", test_process_one_excel_month_filter)

        print("\n=== 配置 ===")
        check("配置读写往返", test_config_roundtrip)

        print("\n=== 鉴定栏核对 ===")
        check("分值解析（新闻报道封顶 / 逗号笔误）", test_judge_parse_rules)
        check("整表核对与一键修正", test_judge_audit_and_fix)
        check("鉴定栏页：多文件批量核对与修正", test_audit_tab_batch)

        print("\n=== 界面 / 线程调度（无图形环境时跳过）===")
        check("批量生成后台线程 + 同名自动错开", test_batch_threads)
        check("批量分析后台线程", test_batch_analyze_threads)
        check("未选输出目录时拦截 + 完成弹窗按钮", test_output_dir_required)

        passed = sum(1 for _, ok, _ in RESULTS if ok)
        total = len(RESULTS)
        skip_txt = f"，跳过 {len(SKIPPED)} 项" if SKIPPED else ""
        print(f"\n{passed}/{total} 项通过{skip_txt}")
        if passed == total:
            print("ALL PASS")
            return 0
        print("\n失败项：")
        for name, ok, msg in RESULTS:
            if not ok:
                print(f"  - {name}: {msg}")
        return 1
    finally:
        shutil.rmtree(TMP, ignore_errors=True)


if __name__ == "__main__":
    sys.exit(main_entry())
