# Excel-Text-Exporter

将 Excel 导出为多种可读的文本和 Markdown 格式，一并处理单元格公式和条件格式。

## 用法

```bash
python export_excel.py example.xlsx
```

## 效果

脚本会在当前目录下创建一个 `output` 文件夹，并根据源文件名生成四个文件，每个文件都处理单元格公式和条件格式。

例如，处理 `example.xlsx` 会输出：

* **`example_archive.txt`**
    * 完整的结构化归档，不易阅读。

* **`example_visual.txt`**
    * 纯文本可视化视图，使用等宽字体对齐，模拟 Excel 的视觉布局。

* **`example_visual_plain.md`**
    * Markdown 表格，合并的单元格将仅在左上角显示内容。

* **`example_visual_rich.md`**
    * Markdown 表格。如果表格包含合并单元格，会自动使用 HTML 语法以保证视觉效果的呈现。