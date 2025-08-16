import sys
import os

# --- 1. 依赖库检测 ---
try:
    import openpyxl
    from openpyxl.utils import get_column_letter
except ImportError:
    print("错误: 缺少必要的库 'openpyxl'。")
    print("请使用此命令安装: pip install openpyxl")
    sys.exit(1)

import datetime

def get_display_width(s):
    """计算字符串的显示宽度，中文字符宽度为2，英文字符为1。"""
    width = 0
    if s is None: return 0
    text = str(s)
    for char in text:
        if '\u4e00' <= char <= '\u9fa5' or '\uff00' <= char <= '\uffef':
            width += 2
        else:
            width += 1
    return width

def get_rule_details(rule, is_markdown=False):
    """
    辅助函数，用于从单个 Rule 对象中提取详细信息。
    is_markdown 参数控制输出格式。
    """
    rule_type = type(rule).__name__
    details = [f"**类型**: {rule_type}"] if is_markdown else [f"      类型: {rule_type}"]
    
    prefix = "" if is_markdown else "      "
    
    if hasattr(rule, 'formula') and rule.formula:
        formula_str = ', '.join(map(str, rule.formula))
        details.append(f"{prefix}**公式**: `{formula_str}`" if is_markdown else f"{prefix}公式: {formula_str}")
    if hasattr(rule, 'operator') and rule.operator:
        details.append(f"{prefix}**运算符**: {rule.operator}" if is_markdown else f"{prefix}运算符: {rule.operator}")
    if hasattr(rule, 'text') and rule.text:
        details.append(f"{prefix}**文本内容**: {rule.text}" if is_markdown else f"{prefix}文本内容: {rule.text}")
    if hasattr(rule, 'dxf') and rule.dxf:
        if rule.dxf.font and hasattr(rule.dxf.font, 'color') and rule.dxf.font.color:
            details.append(f"{prefix}**字体颜色**: {rule.dxf.font.color.rgb}" if is_markdown else f"{prefix}字体颜色: {rule.dxf.font.color.rgb}")
        if rule.dxf.fill and hasattr(rule.dxf.fill, 'start_color') and rule.dxf.fill.start_color:
            details.append(f"{prefix}**背景填充色**: {rule.dxf.fill.start_color.rgb}" if is_markdown else f"{prefix}背景填充色: {rule.dxf.fill.start_color.rgb}")

    return "  \n".join(details) if is_markdown else "\n".join(details)


def export_excel_to_text(file_path, archive_file, visual_file_txt, visual_file_md_plain, visual_file_md_rich):
    """
    将Excel文件导出为四种格式的文本文件。
    """
    try:
        wb_formulas = openpyxl.load_workbook(file_path, data_only=False)
        wb_values = openpyxl.load_workbook(file_path, data_only=True)
    except Exception as e:
        print(f"\n错误：无法读取Excel文件 '{file_path}'。")
        print(f"详细信息: {e}")
        sys.exit(1)

    # --- Part 1: 生成详细归档文件 (archive_file) ---
    with open(archive_file, 'w', encoding='utf-8') as f:
        f.write(f"--- Excel文件完整归档 ---\n")
        f.write(f"文件: {file_path}\n")
        f.write(f"归档时间: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
        f.write("=" * 50 + "\n\n")
        for sheet_name in wb_formulas.sheetnames:
            sheet = wb_formulas[sheet_name]
            sheet_values = wb_values[sheet_name]
            f.write(f"工作表: {sheet_name}\n")
            f.write("-" * 40 + "\n\n")
            f.write("单元格数据:\n")
            if sheet.calculate_dimension() == 'A1:A1' and sheet['A1'].value is None:
                 f.write("(此工作表为空)\n")
            else:
                for row_idx in range(1, sheet.max_row + 1):
                    for col_idx in range(1, sheet.max_column + 1):
                        cell = sheet.cell(row=row_idx, column=col_idx)
                        if cell.value is not None:
                            display_value = sheet_values.cell(row=row_idx, column=col_idx).value
                            f.write(f"  - 单元格: {cell.coordinate}\n")
                            if cell.data_type == 'f':
                                f.write(f"    公式: {cell.value}\n")
                                f.write(f"    显示值: {display_value}\n")
                            else:
                                f.write(f"    值: {cell.value}\n")
            f.write("\n条件格式:\n")
            if sheet.conditional_formatting:
                for cf_obj in sheet.conditional_formatting:
                    f.write(f"  - 作用范围: {cf_obj.sqref}\n")
                    for i, rule in enumerate(cf_obj.rules):
                        f.write(f"    - 规则 #{i + 1}\n")
                        f.write(f"{get_rule_details(rule)}\n")
                    f.write("\n")
            else:
                f.write("  (此工作表无条件格式)\n")
            f.write("=" * 50 + "\n\n")
    # --- 实时输出成功信息 ---
    print(f"已生成: {archive_file}")

    # --- Part 2, 3, 4: 分析并生成三种可视化文件 ---
    visual_txt_content = f"--- Excel文件可视化视图 ---\n文件: {file_path}\n生成时间: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n" + "=" * 50 + "\n\n"
    visual_md_plain_content = ""
    visual_md_rich_content = ""

    for sheet_name in wb_formulas.sheetnames:
        sheet_formulas = wb_formulas[sheet_name]
        sheet_values = wb_values[sheet_name]
        
        full_grid_data, formulas_map = [], {}
        formula_counter = 1
        non_empty_rows, non_empty_cols = set(), set()
        
        for r_idx in range(1, sheet_formulas.max_row + 1):
            row_data = []
            for c_idx in range(1, sheet_formulas.max_column + 1):
                cell_formulas = sheet_formulas.cell(row=r_idx, column=c_idx)
                cell_values = sheet_values.cell(row=r_idx, column=c_idx)
                display_text = ""
                if cell_formulas.data_type == 'f':
                    tag = f"[{formula_counter}]"
                    formulas_map[tag] = cell_formulas.value
                    val = cell_values.value if cell_values.value is not None else ""
                    display_text = f"{val}{tag}"
                    formula_counter += 1
                else:
                    val = cell_values.value if cell_values.value is not None else ""
                    display_text = str(val)
                if display_text.strip() != "":
                    non_empty_rows.add(r_idx)
                    non_empty_cols.add(c_idx)
                row_data.append(display_text)
            full_grid_data.append(row_data)

        sheet_header_txt = f"工作表: {sheet_name}\n" + "-" * 40 + "\n\n"
        sheet_header_md = f"## 工作表: {sheet_name}\n\n"
        
        if not non_empty_rows or not non_empty_cols:
            visual_txt_content += sheet_header_txt + "(此工作表无数据)\n\n"
            visual_md_plain_content += sheet_header_md + "*(此工作表无数据)*\n\n"
            visual_md_rich_content += sheet_header_md + "*(此工作表无数据)*\n\n"
            continue

        min_r, max_r = min(non_empty_rows), max(non_empty_rows)
        min_c, max_c = min(non_empty_cols), max(non_empty_cols)
        
        merge_info = {}
        for merged_range in sheet_formulas.merged_cells.ranges:
            min_col, min_row, max_col, max_row = merged_range.bounds
            primary_cell = (min_row, min_col)
            merge_info[primary_cell] = {'primary': True, 'colspan': max_col - min_col + 1, 'rowspan': max_row - min_row + 1}
            for r in range(min_row, max_row + 1):
                for c in range(min_col, max_col + 1):
                    if (r, c) != primary_cell:
                        merge_info[(r, c)] = {'primary': False}

        col_widths = {c: get_display_width(get_column_letter(c)) for c in range(min_c, max_c + 1)}
        for r_idx in range(min_r, max_r + 1):
            for c_idx in range(min_c, max_c + 1):
                cell_text = full_grid_data[r_idx - 1][c_idx - 1]
                col_widths[c_idx] = max(col_widths.get(c_idx, 0), get_display_width(cell_text))
        
        row_header_width = len(str(max_r))
        headers_txt = [" " * (col_widths.get(c, 0) - len(get_column_letter(c))) + get_column_letter(c) for c in range(min_c, max_c + 1)]
        visual_txt_content += sheet_header_txt + " " * row_header_width + " | " + " | ".join(headers_txt) + "\n"
        separator_txt = ["-" * col_widths.get(c, 0) for c in range(min_c, max_c + 1)]
        visual_txt_content += "-" * row_header_width + "-+-" + "-+-".join(separator_txt) + "\n"
        for r_idx in range(min_r, max_r + 1):
            line_data = [full_grid_data[r_idx-1][c_idx-1] + " " * (col_widths.get(c_idx,0) - get_display_width(full_grid_data[r_idx-1][c_idx-1])) for c_idx in range(min_c, max_c + 1)]
            visual_txt_content += f"{str(r_idx).rjust(row_header_width)} | " + " | ".join(line_data) + "\n"

        headers_md = [""] + [get_column_letter(c) for c in range(min_c, max_c + 1)]
        visual_md_plain_content += sheet_header_md + "| " + " | ".join(headers_md) + " |\n"
        visual_md_plain_content += "|:" + "--:|:" + ":|".join(["--"] * (max_c - min_c + 1)) + "|\n"
        for r_idx in range(min_r, max_r + 1):
            line_data_md = [f"**{r_idx}**"]
            for c_idx in range(min_c, max_c + 1):
                info = merge_info.get((r_idx, c_idx))
                if info and not info['primary']: line_data_md.append("")
                else: line_data_md.append(full_grid_data[r_idx - 1][c_idx - 1])
            visual_md_plain_content += "| " + " | ".join(line_data_md) + " |\n"
        
        visual_md_rich_content += sheet_header_md
        if not sheet_formulas.merged_cells.ranges:
            visual_md_rich_content += "| " + " | ".join(headers_md) + " |\n"
            visual_md_rich_content += "|:" + "--:|:" + ":|".join(["--"] * (max_c - min_c + 1)) + "|\n"
            for r_idx in range(min_r, max_r + 1):
                line_data_md = [f"**{r_idx}**"] + [full_grid_data[r_idx-1][c_idx-1] for c_idx in range(min_c, max_c+1)]
                visual_md_rich_content += "| " + " | ".join(line_data_md) + " |\n"
        else:
            visual_md_rich_content += "<table>\n  <thead>\n    <tr>\n      <th></th>\n"
            for c in range(min_c, max_c+1): visual_md_rich_content += f"      <th>{get_column_letter(c)}</th>\n"
            visual_md_rich_content += "    </tr>\n  </thead>\n  <tbody>\n"
            for r_idx in range(min_r, max_r + 1):
                visual_md_rich_content += f"    <tr>\n      <td><b>{r_idx}</b></td>\n"
                for c_idx in range(min_c, max_c + 1):
                    info = merge_info.get((r_idx, c_idx))
                    if info and info['primary']:
                        visual_md_rich_content += f'      <td colspan="{info["colspan"]}" rowspan="{info["rowspan"]}">{full_grid_data[r_idx - 1][c_idx - 1]}</td>\n'
                    elif info and not info['primary']: continue
                    else: visual_md_rich_content += f"      <td>{full_grid_data[r_idx - 1][c_idx - 1]}</td>\n"
                visual_md_rich_content += "    </tr>\n"
            visual_md_rich_content += "  </tbody>\n</table>\n"

        txt_legend, md_legend = "", ""
        if formulas_map:
            txt_legend += "\n--- 引用列表 ---\n"
            md_legend += "\n### 引用列表\n"
            for tag, formula in sorted(formulas_map.items(), key=lambda item: int(item[0][1:-1])):
                txt_legend += f"{tag}: {formula}\n"
                md_legend += f"- **`{tag}`**: `{formula}`\n"
        if sheet_formulas.conditional_formatting:
            txt_legend += "\n--- 条件格式 ---\n"
            md_legend += "\n### 条件格式\n"
            for cf_obj in sheet_formulas.conditional_formatting:
                txt_legend += f"  - 作用范围: {cf_obj.sqref}\n"
                md_legend += f"- **作用范围**: `{cf_obj.sqref}`\n"
                for i, rule in enumerate(cf_obj.rules):
                    txt_legend += f"    - 规则 #{i + 1}\n{get_rule_details(rule)}\n"
                    md_legend += f"  - **规则 #{i + 1}**\n    - {get_rule_details(rule, is_markdown=True)}\n"
        visual_txt_content += txt_legend
        visual_md_plain_content += md_legend
        visual_md_rich_content += md_legend
        
    with open(visual_file_txt, 'w', encoding='utf-8') as f: f.write(visual_txt_content)
    print(f"已生成: {visual_file_txt}")
    
    with open(visual_file_md_plain, 'w', encoding='utf-8') as f: f.write(visual_md_plain_content)
    print(f"已生成: {visual_file_md_plain}")

    with open(visual_file_md_rich, 'w', encoding='utf-8') as f: f.write(visual_md_rich_content)
    print(f"已生成: {visual_file_md_rich}")

if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("用法: python export_excel.py <example.xlsx>")
        sys.exit(1)
    input_excel_file = sys.argv[1]
    if not os.path.exists(input_excel_file):
        print(f"错误: 文件 '{input_excel_file}' 不存在。")
        sys.exit(1)
    if not input_excel_file.lower().endswith('.xlsx'):
        print(f"错误: 文件 '{input_excel_file}' 不是 .xlsx 格式。")
        sys.exit(1)

    output_dir = 'output'
    os.makedirs(output_dir, exist_ok=True)
    base_name = os.path.basename(input_excel_file)
    name_without_ext = os.path.splitext(base_name)[0]

    archive_output_file = os.path.join(output_dir, f"{name_without_ext}_archive.txt")
    visual_output_file_txt = os.path.join(output_dir, f"{name_without_ext}_visual.txt")
    visual_output_file_md_plain = os.path.join(output_dir, f"{name_without_ext}_visual_plain.md")
    visual_output_file_md_rich = os.path.join(output_dir, f"{name_without_ext}_visual_rich.md")

    print(f"正在处理文件: {input_excel_file}")
    try:
        export_excel_to_text(input_excel_file, archive_output_file, visual_output_file_txt, visual_output_file_md_plain, visual_output_file_md_rich)
        print("\n处理完成！")
    except Exception as e:
        print(f"\n处理过程中发生未知错误: {e}")