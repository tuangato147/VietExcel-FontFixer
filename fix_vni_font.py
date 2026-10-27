# -*- coding: utf-8 -*-
"""
Script to fix VNI font encoding in Excel files (.xls, .xlsx)
Converts VNI-Times / VNI-Helve text to standard Unicode (Times New Roman / Arial)
Preserves all formatting: cell borders, column widths, row heights, merged cells, bold/italic, alignment.
"""

import os
import re
import sys
import xlrd
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# Ensure stdout can print UTF-8
if sys.stdout.encoding != 'utf-8':
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass

# Complete VNI-Windows to Unicode mapping table
VNI_TO_UNICODE = {
    # 3-char sequences (horn + vowel combinations)
    '\xf6\xf4\xf9': 'ướ',
    '\xf6\xf4\xf8': 'ườ',
    '\xf6\xf4\xfb': 'ưở',
    '\xf6\xf4\xfa': 'ưở',
    '\xf6\xf4\xf5': 'ưỡ',
    '\xf6\xf4\xef': 'ượ',
    '\xf6\xf4': 'ươ',
    '\xd6\xd4\xd9': 'ƯỚ',
    '\xd6\xd4\xd8': 'ƯỜ',
    '\xd6\xd4\xdb': 'ƯỞ',
    '\xd6\xd4\xda': 'ƯỞ',
    '\xd6\xd4\xd5': 'ƯỠ',
    '\xd6\xd4\xdf': 'ƯỢ',
    '\xd6\xd4': 'ƯƠ',

    # Horn combinations (ơ, ư)
    '\xf4\xf9': 'ớ',
    '\xf4\xf8': 'ờ',
    '\xf4\xfb': 'ở',
    '\xf4\xfa': 'ở',
    '\xf4\xf5': 'ỡ',
    '\xf4\xef': 'ợ',
    '\xf4': 'ơ',

    '\xf6\xf9': 'ứ',
    '\xf6\xf8': 'ừ',
    '\xf6\xfb': 'ử',
    '\xf6\xfa': 'ử',
    '\xf6\xf5': 'ữ',
    '\xf6\xef': 'ự',
    '\xf6': 'ư',

    '\xd4\xd9': 'Ớ',
    '\xd4\xd8': 'Ờ',
    '\xd4\xdb': 'Ở',
    '\xd4\xda': 'Ở',
    '\xd4\xd5': 'Ỡ',
    '\xd4\xdf': 'Ợ',
    '\xd4': 'Ơ',

    '\xd6\xd9': 'Ứ',
    '\xd6\xd8': 'Ừ',
    '\xd6\xdb': 'Ử',
    '\xd6\xda': 'Ử',
    '\xd6\xd5': 'Ữ',
    '\xd6\xdf': 'Ự',
    '\xd6': 'Ư',

    # Breve combinations (ă)
    'a\xe9': 'ắ',
    'a\xe8': 'ằ',
    'a\xec': 'ẳ',
    'a\xed': 'ẵ',
    'a\xeb': 'ặ',
    'a\xea': 'ă',
    'A\xc9': 'Ắ',
    'A\xc8': 'Ằ',
    'A\xcc': 'Ẳ',
    'A\xcd': 'Ẵ',
    'A\xcb': 'Ặ',
    'A\xca': 'Ă',

    # Circumflex combinations (â, ê, ô)
    'a\xe1': 'ấ',
    'a\xe0': 'ầ',
    'a\xe5': 'ẩ',
    'a\xe3': 'ẫ',
    'a\xe4': 'ậ',
    'a\xe2': 'â',

    'A\xd9': 'Á',
    'A\xd8': 'À',
    'A\xdb': 'Ả',
    'A\xd5': 'Ã',
    'A\xcf': 'Ạ',
    'A\xc1': 'Ấ',
    'A\xc0': 'Ầ',
    'A\xc5': 'Ẩ',
    'A\xc3': 'Ẫ',
    'A\xc4': 'Ậ',
    'A\xc2': 'Â',

    'e\xe1': 'ế',
    'e\xe0': 'ề',
    'e\xe5': 'ể',
    'e\xe3': 'ễ',
    'e\xe4': 'ệ',
    'e\xe2': 'ê',

    'E\xc1': 'Ế',
    'E\xc0': 'Ề',
    'E\xc5': 'Ể',
    'E\xc3': 'Ễ',
    'E\xc4': 'Ệ',
    'E\xc2': 'Ê',

    'o\xe1': 'ố',
    'o\xe0': 'ồ',
    'o\xe5': 'ổ',
    'o\xe3': 'ỗ',
    'o\xe4': 'ộ',
    'o\xe2': 'ô',

    'O\xc1': 'Ố',
    'O\xc0': 'Ồ',
    'O\xc5': 'Ổ',
    'O\xc3': 'Ỗ',
    'O\xc4': 'Ộ',
    'O\xc2': 'Ô',

    # Standard vowel + acute (sắc)
    'a\xf9': 'á',
    'e\xf9': 'é',
    'o\xf9': 'ó',
    'u\xf9': 'ú',
    'y\xf9': 'ý',

    # Standard vowel + grave (huyền)
    'a\xf8': 'à',
    'e\xf8': 'è',
    'o\xf8': 'ò',
    'u\xf8': 'ù',
    'y\xf8': 'ỳ',

    # Standard vowel + hook (hỏi)
    'a\xfb': 'ả',
    'e\xfb': 'ẻ',
    'o\xfb': 'ỏ',
    'u\xfb': 'ủ',
    'y\xfb': 'ỷ',

    # Standard vowel + tilde (ngã)
    'a\xf5': 'ã',
    'e\xf5': 'ẽ',
    'o\xf5': 'õ',
    'u\xf5': 'ũ',
    'y\xf5': 'ỹ',

    # Standard vowel + dot below (nặng)
    'a\xef': 'ạ',
    'e\xef': 'ẹ',
    'o\xef': 'ọ',
    'u\xef': 'ụ',
    'y\xef': 'ỵ',

    # Capital U + hook / accents
    'U\xdb': 'Ủ',
    'U\xd9': 'Ú',
    'U\xd8': 'Ù',
    'U\xd5': 'Ũ',
    'U\xdf': 'Ụ',

    # Standalone i and d
    '\xf1': 'đ',
    '\xd1': 'Đ',
    '\xec': 'ì',
    '\xed': 'í',
    '\xf2': 'ị',
    '\xe6': 'ỉ',
    '\xcc': 'Ì',
    '\xcd': 'Í',
}

# Compile single-pass regex, sorted by length descending
VNI_REGEX = re.compile('|'.join(re.escape(k) for k in sorted(VNI_TO_UNICODE.keys(), key=len, reverse=True)))

def convert_vni_to_unicode(text):
    if not isinstance(text, str):
        return text
    return VNI_REGEX.sub(lambda m: VNI_TO_UNICODE[m.group(0)], text)

BORDER_MAP = {
    0: None,
    1: 'thin',
    2: 'medium',
    3: 'dashed',
    4: 'dotted',
    5: 'thick',
    6: 'double',
    7: 'hair',
    8: 'mediumDashed',
    9: 'dashDot',
    10: 'mediumDashDot',
    11: 'dashDotDot',
    12: 'mediumDashDotDot',
    13: 'slantDashDot',
}

HOR_ALIGN_MAP = {
    0: None,
    1: 'left',
    2: 'center',
    3: 'right',
    4: 'fill',
    5: 'justify',
    6: 'centerContinuous',
    7: 'distributed',
}

VERT_ALIGN_MAP = {
    0: 'top',
    1: 'center',
    2: 'bottom',
    3: 'justify',
    4: 'distributed',
}

def convert_xls_to_xlsx(src_path, dst_path):
    print(f"Reading source XLS: {src_path}")
    wb_xls = xlrd.open_workbook(src_path, formatting_info=True)
    wb_xlsx = openpyxl.Workbook()
    # Remove default active sheet
    wb_xlsx.remove(wb_xlsx.active)

    for sheet_name in wb_xls.sheet_names():
        s_xls = wb_xls.sheet_by_name(sheet_name)
        new_sheet_name = convert_vni_to_unicode(sheet_name)
        ws_xlsx = wb_xlsx.create_sheet(title=new_sheet_name)
        print(f"Processing sheet: '{sheet_name}' -> '{new_sheet_name}' ({s_xls.nrows} rows, {s_xls.ncols} cols)")

        # Set column widths
        for col_idx in range(s_xls.ncols):
            col_letter = get_column_letter(col_idx + 1)
            if col_idx in s_xls.colinfo_map:
                width = s_xls.colinfo_map[col_idx].width / 256.0
                ws_xlsx.column_dimensions[col_letter].width = max(width + 1.5, 4.0)
            else:
                ws_xlsx.column_dimensions[col_letter].width = 10.0

        # Set row heights
        for row_idx in range(s_xls.nrows):
            if row_idx in s_xls.rowinfo_map:
                height = s_xls.rowinfo_map[row_idx].height / 20.0
                ws_xlsx.row_dimensions[row_idx + 1].height = height

        # Convert cells and copy styles
        for r in range(s_xls.nrows):
            for c in range(s_xls.ncols):
                cell_value = s_xls.cell_value(r, c)
                cell_type = s_xls.cell_type(r, c)

                # Format text
                if isinstance(cell_value, str):
                    converted_val = convert_vni_to_unicode(cell_value)
                elif cell_type == 2 and cell_value == int(cell_value):
                    # Preserve integer appearance if exact integer
                    converted_val = cell_value
                else:
                    converted_val = cell_value

                target_cell = ws_xlsx.cell(row=r + 1, column=c + 1, value=converted_val)

                # Style transfer
                xf_idx = s_xls.cell_xf_index(r, c)
                xf = wb_xls.xf_list[xf_idx]
                font_info = wb_xls.font_list[xf.font_index]

                # Modern font name mapping
                font_name = font_info.name
                if 'VNI-Times' in font_name:
                    font_name = 'Times New Roman'
                elif 'VNI-Helve' in font_name:
                    font_name = 'Arial'
                elif 'VNI' in font_name:
                    font_name = 'Times New Roman'

                font_size = font_info.height / 20.0 if font_info.height > 0 else 11.0
                bold = bool(font_info.bold)
                italic = bool(font_info.italic)

                target_cell.font = Font(
                    name=font_name,
                    size=font_size,
                    bold=bold,
                    italic=italic
                )

                # Alignment
                h_align = HOR_ALIGN_MAP.get(xf.alignment.hor_align, None)
                v_align = VERT_ALIGN_MAP.get(xf.alignment.vert_align, 'center')
                wrap = bool(xf.alignment.text_wrapped)
                target_cell.alignment = Alignment(horizontal=h_align, vertical=v_align, wrap_text=wrap)

                # Border
                b = xf.border
                top_s = BORDER_MAP.get(b.top_line_style, None)
                bot_s = BORDER_MAP.get(b.bottom_line_style, None)
                left_s = BORDER_MAP.get(b.left_line_style, None)
                right_s = BORDER_MAP.get(b.right_line_style, None)

                top_side = Side(style=top_s) if top_s else None
                bot_side = Side(style=bot_s) if bot_s else None
                left_side = Side(style=left_s) if left_s else None
                right_side = Side(style=right_s) if right_s else None

                target_cell.border = Border(top=top_side, bottom=bot_side, left=left_side, right=right_side)

        # Merged cells
        for crange in s_xls.merged_cells:
            rlo, rhi, clo, chi = crange
            ws_xlsx.merge_cells(
                start_row=rlo + 1,
                end_row=rhi,
                start_column=clo + 1,
                end_column=chi
            )

    print(f"Saving fixed XLSX to: {dst_path}")
    wb_xlsx.save(dst_path)
    print("XLSX conversion completed successfully!")


def convert_xls_to_xls(src_path, dst_path):
    """Also create a Unicode-compatible .xls file for legacy support"""
    import xlwt
    print(f"Creating legacy XLS: {dst_path}")
    wb_xls = xlrd.open_workbook(src_path, formatting_info=True)
    wb_out = xlwt.Workbook(encoding='utf-8')

    for sheet_name in wb_xls.sheet_names():
        s_xls = wb_xls.sheet_by_name(sheet_name)
        new_sheet_name = convert_vni_to_unicode(sheet_name)
        ws_out = wb_out.add_sheet(new_sheet_name, cell_overwrite_ok=True)

        # Set column widths
        for col_idx in range(s_xls.ncols):
            if col_idx in s_xls.colinfo_map:
                ws_out.col(col_idx).width = s_xls.colinfo_map[col_idx].width + 300
            else:
                ws_out.col(col_idx).width = 256 * 10

        # Styles cache
        style_cache = {}

        for r in range(s_xls.nrows):
            if r in s_xls.rowinfo_map:
                ws_out.row(r).height = s_xls.rowinfo_map[r].height

            for c in range(s_xls.ncols):
                val = s_xls.cell_value(r, c)
                if isinstance(val, str):
                    val = convert_vni_to_unicode(val)

                xf_idx = s_xls.cell_xf_index(r, c)
                if xf_idx not in style_cache:
                    xf = wb_xls.xf_list[xf_idx]
                    font_info = wb_xls.font_list[xf.font_index]

                    style = xlwt.XFStyle()
                    font = xlwt.Font()
                    font.name = 'Times New Roman'
                    font.height = font_info.height if font_info.height > 0 else 220
                    font.bold = bool(font_info.bold)
                    font.italic = bool(font_info.italic)
                    style.font = font

                    # Alignment
                    alignment = xlwt.Alignment()
                    alignment.horz = min(xf.alignment.hor_align, 7)
                    alignment.vert = min(xf.alignment.vert_align, 4)
                    alignment.wrap = xf.alignment.text_wrapped
                    style.alignment = alignment

                    # Borders
                    borders = xlwt.Borders()
                    borders.top = min(xf.border.top_line_style, 13)
                    borders.bottom = min(xf.border.bottom_line_style, 13)
                    borders.left = min(xf.border.left_line_style, 13)
                    borders.right = min(xf.border.right_line_style, 13)
                    style.borders = borders

                    style_cache[xf_idx] = style

                ws_out.write(r, c, val, style_cache[xf_idx])

        # Merges
        for crange in s_xls.merged_cells:
            rlo, rhi, clo, chi = crange
            ws_out.merge(rlo, rhi - 1, clo, chi - 1)

    wb_out.save(dst_path)
    print("XLS conversion completed successfully!")


if __name__ == '__main__':
    src = os.path.abspath('BT TINH TOAN NANG SUAT-THIET BI (1).xls')
    dst_xlsx = os.path.abspath('BT TINH TOAN NANG SUAT-THIET BI (1).xlsx')
    dst_xls = os.path.abspath('BT TINH TOAN NANG SUAT-THIET BI (1)_fixed.xls')

    convert_xls_to_xlsx(src, dst_xlsx)
    convert_xls_to_xls(src, dst_xls)
