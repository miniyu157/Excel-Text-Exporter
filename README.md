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

## 配置

脚本的行为可以通过仓库根目录下的 `config.toml` 文件进行自定义。

脚本在首次运行时，如果未找到该文件，会自动生成一份默认配置。具体可配置项请直接查阅该文件中的注释。

## 效果

脚本会根据配置，在 `output` 文件夹内生成相应的文件。其中，**可视化文件**和**数据归档文件**会完整处理单元格公式、条件格式、命名区域、超链接及批注；而 **CSV 文件**仅包含纯净的最终显示值。

例如，处理 `example.xlsx` (假设内含 "Sheet1" 和 "Sheet2" 两个工作表) 会输出：

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

#### 数据交换文件

* **`example_Sheet1.csv`**
* **`example_Sheet2.csv`**
    * 每个工作表一个对应的纯净 CSV 文件，仅包含最终显示值，便于在其他程序中进行数据分析。