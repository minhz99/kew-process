import os
import io
import unicodedata
import re
from pathlib import Path
from typing import Sequence, Mapping
from openpyxl import load_workbook

def _nfc(s: object) -> str:
    """Chuẩn hoá tên hiển thị."""
    if s is None:
        return ""
    return unicodedata.normalize("NFC", str(s).strip())

def _parse_float(v: object) -> float | None:
    """Trích xuất float từ ô dữ liệu."""
    if v is None:
        return None
    try:
        import pandas as pd
        if isinstance(v, float) and pd.isna(v):
            return None
    except Exception:
        pass
    try:
        s = str(v).strip().replace(",", ".")
        if not s:
            return None
        return float(s)
    except (ValueError, TypeError):
        return None

def generate_mba_excel_from_devices(
    devices: Sequence[Mapping],
    output_path: str | Path,
    template_path: str | Path,
) -> Path:
    """
    Sinh báo cáo Excel MBA từ danh sách metadata thiết bị (devices).
    
    Args:
        devices: Danh sách metadata (name, excel_params, kind, ...).
        output_path: Đường dẫn lưu file Excel kết quả.
        template_path: Đường dẫn template Excel MBA.
    """
    from modules.report.gen_word import (
        _resolve_word_section_kind, _eval_voltage, _eval_pf, _eval_thd, _eval_unbalance,
        _V_DEV_LIMIT_PCT, _THDV_LIMIT_PCT, _TDD_LIMIT_PCT, _MBA_NOMINAL_VOLTAGE_V
    )

    # 1. Lọc danh sách MBA
    mbas = []
    found_kinds = []
    for d in devices:
        k_val = d.get("kind")
        kind = _resolve_word_section_kind({"kind": k_val}, name=d.get("name", ""), default_kind=None)
        found_kinds.append(f"{d.get('name')}: {kind} (original type: {k_val})")
        if kind == "mba":
            mbas.append(d)

    if not mbas:
        details = "; ".join(found_kinds)
        raise ValueError(f"Không tìm thấy máy biến áp nào trong dữ liệu. Danh sách thiết bị: {details}")

    # 2. Mở template
    wb = load_workbook(str(template_path), keep_vba=True)
    
    # Đảm bảo sheet "MBA1" tồn tại
    if "MBA1" not in wb.sheetnames:
        # Nếu template không có MBA1, cố gắng lấy sheet đầu tiên hoặc tạo mới (nhưng spec yêu cầu MBA1)
        raise ValueError("Template Excel không có sheet 'MBA1'.")

    # Xóa các sheet MBA khác (nếu có sẵn trong template cũ) để làm sạch
    for sname in list(wb.sheetnames):
        if sname.startswith("MBA") and sname != "MBA1":
            del wb[sname]

    source_ws = wb["MBA1"]
    
    # 3. Điền dữ liệu vào các sheet MBA
    for i, mba_data in enumerate(mbas):
        name = _nfc(mba_data.get("name") or f"MBA{i+1}")
        # Hạn chế độ dài tên sheet (Excel giới hạn 31 ký tự)
        sheet_name = name[:31]
        
        if i == 0:
            ws = source_ws
            ws.title = sheet_name
        else:
            ws = wb.copy_worksheet(source_ws)
            ws.title = sheet_name

        ep = mba_data.get("excel_params") or {}
        
        # --- Đánh giá và ghi vào các ô AE (col 31) ---
        # AE8: Điện áp
        u_max = _parse_float(ep.get("u_max"))
        u_min = _parse_float(ep.get("u_min"))
        # Giả định u_avg là trung bình nếu không có cột avg riêng
        u_avg = (u_max + u_min) / 2 if (u_max is not None and u_min is not None) else None
        u_eval, _, _, _ = _eval_voltage(u_max, u_min, u_avg, _MBA_NOMINAL_VOLTAGE_V)
        ws["AE8"] = u_eval
        
        # AE14: Lệch pha áp
        du = _parse_float(ep.get("delta_u"))
        ws["AE14"] = _eval_unbalance(du, du, _V_DEV_LIMIT_PCT)
        
        # AE16: PF
        pf_val = _parse_float(ep.get("cos_phi"))
        if pf_val is not None and abs(pf_val) > 1.0:
            pf_val /= 1000.0
        ws["AE16"] = _eval_pf(pf_val, pf_val, pf_val)
        
        # AE20: THD
        thd = _parse_float(ep.get("thd"))
        ws["AE20"] = _eval_thd([thd], [thd], _THDV_LIMIT_PCT)
        
        # AE23: TDD
        tdd = _parse_float(ep.get("tdd"))
        ws["AE23"] = _eval_thd([tdd], [tdd], _TDD_LIMIT_PCT)

    # 4. Cập nhật bảng tổng hợp tại sheet "Tổn thất MBA"
    if "Tổn thất MBA" in wb.sheetnames:
        ws_ton = wb["Tổn thất MBA"]
        
        # Dòng mẫu A6:J6
        # STT (1) - A6
        # Tên MBA - B6
        # Pdm - D6
        
        for i, mba_data in enumerate(mbas):
            row = 6 + i
            if i > 0:
                # Copy định dạng và công thức từ dòng 6 xuống dòng row
                for col in range(1, 11): # A (1) -> J (10)
                    source_cell = ws_ton.cell(row=6, column=col)
                    new_cell = ws_ton.cell(row=row, column=col)
                    new_cell.value = source_cell.value
                    if source_cell.has_style:
                        from copy import copy
                        new_cell.font = copy(source_cell.font)
                        new_cell.border = copy(source_cell.border)
                        new_cell.fill = copy(source_cell.fill)
                        new_cell.number_format = copy(source_cell.number_format)
                        new_cell.protection = copy(source_cell.protection)
                        new_cell.alignment = copy(source_cell.alignment)

            ep = mba_data.get("excel_params") or {}
            ws_ton.cell(row=row, column=1, value=i + 1) # STT
            ws_ton.cell(row=row, column=2, value=_nfc(mba_data.get("name"))) # Tên MBA
            pdm = _parse_float(ep.get("pdm"))
            ws_ton.cell(row=row, column=4, value=pdm) # Pdm

    # Lưu kết quả
    wb.save(str(output_path))
    return Path(output_path)
