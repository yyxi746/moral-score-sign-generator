# 德育分签字表生成器

GUI 小工具，读取德育分 Excel 统计表，自动生成格式化 Word 签字确认表。

Excel 文件名、工作表、表头名称、列位置不固定，程序自动识别。

## 下载

进入 [Releases 页面](https://github.com/yyxi746/moral-score-sign-generator/releases)，下载最新版 `default.exe`，双击即可使用，无需安装 Python。

## ✨ 功能

1. 图形界面，无需命令行操作；单文件 / 批量处理两个页签
2. 自动识别姓名、德育总分、多组月份表扬信列，表扬信支持按月份段多选，无表扬信填 0
3. 批量处理：一次添加多个 Excel 或整个文件夹，支持把文件/文件夹直接拖入窗口，逐个分析生成、失败不中断、完成后汇总
4. 批量列表双击查看文件详情：学生明细（姓名 / 德育分 / 表扬信条数）一目了然，并可手工修正该文件的 Word 标题与班级
5. 生成前数据校验：分数为空、负数、非数字或人数为 0 的文件提前标黄提醒，避免签完字才发现错误
6. 首次运行自动引导选择 Word 模板；模板缺失时弹窗引导，不再生硬报错
7. 启动时自动检查新版本，状态栏提示并可点击前往下载
8. 配置记忆：Word 模板路径、固定输出目录保存在系统 %APPDATA% 隐藏目录，程序目录不生成配置文件
9. 输出文件名自动将数据源文件名里的「统计表」替换为「签字表」
10. 自动生成 Word 标题，完整继承 Word 模板的表格样式、边框、字体

## 📦 环境依赖

Python >= 3.10

```bash
pip install -r requirements.txt
python main.py
```

## 打包 exe

```bash
pyinstaller --onefile --noconsole --name default --icon app.ico --add-data "app.ico;." --collect-all tkinterdnd2 main.py
```
