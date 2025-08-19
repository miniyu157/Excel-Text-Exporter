# Excel-Text-Exporter

将 Excel 导出为多种可读的可视化文件和结构化的数据归档文件。

## 依赖

```bash
pip install openpyxl toml pyyaml
```

## 用法

```bash
python export_excel.py <example.xlsx>
```

## 效果

脚本会在当前目录下创建一个 `output` 文件夹，并根据源文件名生成六个文件。所有文件都能处理单元格公式、条件格式、命名区域、超链接以及批注。

例如，处理 `example.xlsx` 会输出：

#### 可视化文件

* **`example_visual.txt`**
    * 纯文本可视化视图，使用等宽字体对齐，模拟 Excel 的视觉布局。

* **`example_visual_plain.md`**
    * 标准的 Markdown 表格。合并的单元格将仅在左上角单元格显示内容。

* **`example_visual_rich.md`**
    * 富文本 Markdown 表格。如果表格包含合并单元格，会自动使用 HTML 语法以保证视觉效果的完美呈现。

#### 数据归档文件

* **`example_archive.toml`**
    * TOML 格式的结构化数据归档，包含工作表的所有核心信息。

* **`example_archive.json`**
    * JSON 格式的结构化数据归档，内容与 TOML 版本相同。

* **`example_archive.yaml`**
    * YAML 格式的结构化数据归档，内容与 TOML 版本相同。