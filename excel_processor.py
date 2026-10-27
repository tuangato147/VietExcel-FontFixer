# -*- coding: utf-8 -*-
"""
Excel Processor Module
Handles conversion of .xls and .xlsx files with complete style preservation.
Supports VNI and TCVN3 encoding conversions.
"""

import os
import xlrd
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

from viet_converter import (
    convert_vni,
    convert_tcvn3,
    detect_font_target,
    VNI_REGEX,
    TCVN3_REGEX
)

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

def convert_text_by_mode(text, font_name, mode='auto'):
    if not isinstance(text, str) or not text:
        return text

    if mode == 'vni':
        return convert_vni(text)
    elif mode == 'tcvn3':
        return convert_tcvn3(text)
    else:  # auto mode
        enc, _ = detect_font_target(font_name)
        if enc == 'VNI':
            return convert_vni(text)
        elif enc == 'TCVN3':
            return convert_tcvn3(text)
        else:
            # Fallback heuristic: check if string contains VNI patterns
            if VNI_REGEX.search(text):
                return convert_vni(text)
            elif TCVN3_REGEX.search(text):
                return convert_tcvn3(text)
            return text


def process_xls(src_path, dst_path, mode='auto', progress_cb=None):
    """Convert legacy .xls file to .xlsx with full style preservation."""
    if progress_cb:
        progress_cb(0.1, f"Đang đọc file XLS: {os.path.basename(src_path)}")

    wb_xls = xlrd.open_workbook(src_path, formatting_info=True)
    wb_xlsx = openpyxl.Workbook()
    wb_xlsx.remove(wb_xlsx.active)  # remove default empty sheet

    total_sheets = len(wb_xls.sheet_names())

    for s_idx, sheet_name in enumerate(wb_xls.sheet_names()):
        s_xls = wb_xls.sheet_by_name(sheet_name)
        new_sheet_name = convert_text_by_mode(sheet_name, '', mode)
        # Sheet title max 31 chars in Excel
        new_sheet_name = new_sheet_name[:31] if new_sheet_name else f"Sheet{s_idx + 1}"
        ws_xlsx = wb_xlsx.create_sheet(title=new_sheet_name)

        if progress_cb:
            sheet_ratio = 0.1 + (0.8 * (s_idx / max(total_sheets, 1)))
            progress_cb(sheet_ratio, f"Đang xử lý sheet '{sheet_name}' ({s_xls.nrows} dòng)")

        # Column widths
        for col_idx in range(s_xls.ncols):
            col_letter = get_column_letter(col_idx + 1)
            if col_idx in s_xls.colinfo_map:
                width = s_xls.colinfo_map[col_idx].width / 256.0
                ws_xlsx.column_dimensions[col_letter].width = max(width + 1.5, 4.0)
            else:
                ws_xlsx.column_dimensions[col_letter].width = 10.0

        # Row heights
        for row_idx in range(s_xls.nrows):
            if row_idx in s_xls.rowinfo_map:
                height = s_xls.rowinfo_map[row_idx].height / 20.0
                ws_xlsx.row_dimensions[row_idx + 1].height = height

        # Convert cells and styles
        for r in range(s_xls.nrows):
            for c in range(s_xls.ncols):
                cell_value = s_xls.cell_value(r, c)
                cell_type = s_xls.cell_type(r, c)

                xf_idx = s_xls.cell_xf_index(r, c)
                xf = wb_xls.xf_list[xf_idx]
                font_info = wb_xls.font_list[xf.font_index]
                raw_font_name = font_info.name

                # Target font and converted text
                enc, target_font = detect_font_target(raw_font_name)
                converted_val = convert_text_by_mode(cell_value, raw_font_name, mode)

                target_cell = ws_xlsx.cell(row=r + 1, column=c + 1, value=converted_val)

                # Modern font name mapping
                if mode == 'vni' and 'Times New Roman' not in target_font and 'Arial' not in target_font:
                    target_font = 'Times New Roman'
                elif mode == 'tcvn3' and 'Times New Roman' not in target_font and 'Arial' not in target_font:
                    target_font = 'Times New Roman'

                font_size = font_info.height / 20.0 if font_info.height > 0 else 11.0
                target_cell.font = Font(
                    name=target_font,
                    size=font_size,
                    bold=bool(font_info.bold),
                    italic=bool(font_info.italic)
                )

                # Alignment
                h_align = HOR_ALIGN_MAP.get(xf.alignment.hor_align, None)
                v_align = VERT_ALIGN_MAP.get(xf.alignment.vert_align, 'center')
                target_cell.alignment = Alignment(
                    horizontal=h_align,
                    vertical=v_align,
                    wrap_text=bool(xf.alignment.text_wrapped)
                )

                # Border
                b = xf.border
                top_s = BORDER_MAP.get(b.top_line_style, None)
                bot_s = BORDER_MAP.get(b.bottom_line_style, None)
                left_s = BORDER_MAP.get(b.left_line_style, None)
                right_s = BORDER_MAP.get(b.right_line_style, None)

                target_cell.border = Border(
                    top=Side(style=top_s) if top_s else None,
                    bottom=Side(style=bot_s) if bot_s else None,
                    left=Side(style=left_s) if left_s else None,
                    right=Side(style=right_s) if right_s else None
                )

        # Merged cells
        for crange in s_xls.merged_cells:
            rlo, rhi, clo, chi = crange
            ws_xlsx.merge_cells(
                start_row=rlo + 1,
                end_row=rhi,
                start_column=clo + 1,
                end_column=chi
            )

    if progress_cb:
        progress_cb(0.95, f"Đang lưu file: {os.path.basename(dst_path)}")

    wb_xlsx.save(dst_path)

    if progress_cb:
        progress_cb(1.0, f"Hoàn thành: {os.path.basename(dst_path)}")


def process_xlsx(src_path, dst_path, mode='auto', progress_cb=None):
    """Process modern .xlsx file, fixing fonts while preserving styles."""
    if progress_cb:
        progress_cb(0.1, f"Đang đọc file XLSX: {os.path.basename(src_path)}")

    wb = openpyxl.load_workbook(src_path)
    total_sheets = len(wb.worksheets)

    for s_idx, ws in enumerate(wb.worksheets):
        if progress_cb:
            sheet_ratio = 0.1 + (0.8 * (s_idx / max(total_sheets, 1)))
            progress_cb(sheet_ratio, f"Đang xử lý sheet '{ws.title}'")

        # Convert sheet name
        ws.title = convert_text_by_mode(ws.title, '', mode)[:31]

        # Convert cells
        for row in ws.iter_rows():
            for cell in row:
                if cell.value is None:
                    continue

                font_name = cell.font.name if cell.font and cell.font.name else ''
                enc, target_font = detect_font_target(font_name)

                if isinstance(cell.value, str):
                    new_val = convert_text_by_mode(cell.value, font_name, mode)
                    if new_val != cell.value:
                        cell.value = new_val

                # Update font if detected legacy font
                if enc in ('VNI', 'TCVN3') or mode in ('vni', 'tcvn3'):
                    if cell.font:
                        cell.font = Font(
                            name=target_font,
                            size=cell.font.size,
                            bold=cell.font.bold,
                            italic=cell.font.italic,
                            vertAlign=cell.font.vertAlign,
                            underline=cell.font.underline,
                            strike=cell.font.strike,
                            color=cell.font.color
                        )

    if progress_cb:
        progress_cb(0.95, f"Đang lưu file: {os.path.basename(dst_path)}")

    wb.save(dst_path)

    if progress_cb:
        progress_cb(1.0, f"Hoàn thành: {os.path.basename(dst_path)}")


def convert_file(src_path, dst_path=None, mode='auto', progress_cb=None):
    """
    Unified entry point for converting an Excel file (.xls or .xlsx).
    dst_path: if None, appends '_fixed.xlsx' in the same folder.
    """
    ext = os.path.splitext(src_path)[1].lower()
    if not dst_path:
        base, _ = os.path.splitext(src_path)
        dst_path = f"{base}_fixed.xlsx"

    # Make sure dst_path has .xlsx extension
    dst_base, _ = os.path.splitext(dst_path)
    if not dst_path.lower().endswith('.xlsx'):
        dst_path = f"{dst_base}.xlsx"

    if ext == '.xls':
        process_xls(src_path, dst_path, mode=mode, progress_cb=progress_cb)
    elif ext == '.xlsx':
        process_xlsx(src_path, dst_path, mode=mode, progress_cb=progress_cb)
    else:
        raise ValueError(f"Định dạng không được hỗ trợ: {ext} (chỉ hỗ trợ .xls và .xlsx)")

    return dst_path
