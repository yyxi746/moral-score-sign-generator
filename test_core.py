# -*- coding: utf-8 -*-
import sys, os
sys.stdout.reconfigure(encoding='utf-8')
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import openpyxl
import main
from docx import Document

xlsx = r'C:\Users\瑶\Desktop\文件\资信\4\农学23-1班2025-2026学年第2学期德育分统计表.xlsx'
tpl  = os.path.join(os.path.dirname(os.path.abspath(__file__)), '签字表模板.docx')
out  = os.path.join(os.path.dirname(os.path.abspath(__file__)), '_test_out.docx')

wb = openpyxl.load_workbook(xlsx, data_only=True)
ws = wb['大三下']
headers = main.scan_headers(ws)
name_col, score_col, praise_cols = main.detect_columns(headers)
print('列：', name_col, score_col, praise_cols)

mg = main.group_praise_by_month(ws, praise_cols)
print('月份分组 =', mg)
assert mg == [('3-4月', [5]), ('5-8月', [10])], mg

title = main.extract_title(ws, '大三下')
print('标题 =', title)
assert title == '大三下 2025-2026 第二学期德育分确认签字表', title

# 输出：同源目录
op = main.compute_output_path(xlsx, '')
print('输出(同源目录) =', os.path.basename(op))
assert os.path.basename(op) == '农学23-1班2025-2026学年第2学期德育分签字表.docx'
# 输出：固定目录
fixed = os.path.join(os.path.dirname(os.path.abspath(__file__)), '_fixed_test')
os.makedirs(fixed, exist_ok=True)
op2 = main.compute_output_path(xlsx, fixed)
print('输出(固定目录) =', op2)
assert os.path.dirname(op2) == fixed and os.path.basename(op2).endswith('签字表.docx')

# 配置持久化往返（重定向到临时目录，不碰真实 AppData）
import tempfile
tmpcfg = tempfile.mkdtemp(prefix='moral_cfg_')
_orig_cf = main.config_file
main.config_file = lambda: os.path.join(tmpcfg, 'config.json')
main.save_config({'template_path': r'C:\t.docx', 'output_dir': fixed})
c = main.load_config()
assert c['output_dir'] == fixed and c['template_path'] == r'C:\t.docx', c
real_cfg = _orig_cf()
assert 'AppData' in real_cfg or 'Roaming' in real_cfg, real_cfg
print('真实配置路径 =', real_cfg)
print('配置持久化 OK')
main.config_file = _orig_cf
import shutil
shutil.rmtree(tmpcfg, ignore_errors=True)

cls = main.extract_class(ws)
print('班级 =', cls)

def byname(students): return {s['name']: s for s in students}

allst = main.extract_students(ws, name_col, score_col, [5, 10])
b = byname(allst)
assert len(allst) == 30
assert b['李怡臻']['praise'] == 22 and b['杨凌']['praise'] == 1
assert b['孙文崇']['praise'] == 6 and b['席龙飞']['praise'] == 11

s34 = byname(main.extract_students(ws, name_col, score_col, [5]))
assert s34['李怡臻']['praise'] == 13 and s34['杨凌']['praise'] == 0
s58 = byname(main.extract_students(ws, name_col, score_col, [10]))
assert s58['李怡臻']['praise'] == 9 and s58['杨凌']['praise'] == 1

main.fill_word(tpl, out, allst, title, cls)
doc = Document(out)
t = doc.tables[0]
paras = [p.text for p in t.rows[0].cells[0].paragraphs]
print('标题行段落 =', paras)
assert paras[0] == '大三下 2025-2026 第二学期德育分确认签字表'
assert paras[1] == '班级：农学23-1班'
assert len(t.rows) == 32
wb.close()

# 清理临时固定目录
import shutil
shutil.rmtree(fixed, ignore_errors=True)
print('ALL PASS')
