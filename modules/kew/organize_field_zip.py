"""
Tổ chức hồ sơ đo KEW6315 từ ZIP: đọc Excel hiện trường, đổi tên thư mục Sxxxx,
chuyển ảnh PS-SDxxx.BMP vào đúng thư mục thiết bị, nén lại Project_Output.zip.
"""
from __future__ import annotations

import io
import os
import re
import shutil
import unicodedata
import zipfile
from dataclasses import dataclass
from typing import Any, Optional

import pandas as pd

_SKIP_DIR_NAMES = {"__MACOSX"}

_BMP_RE = re.compile(r"^PS-SD(\d{1,4})\.BMP$", re.IGNORECASE)
_S_DIR_RE = re.compile(r"^S(\d{4})$", re.IGNORECASE)

_WIN_RESERVED = {
    "CON", "PRN", "AUX", "NUL",
    *(f"COM{i}" for i in range(1, 10)),
    *(f"LPT{i}" for i in range(1, 10)),
}


def _norm_key(s: str) -> str:
    if s is None or (isinstance(s, float) and pd.isna(s)):
        return ""
    t = unicodedata.normalize("NFKC", str(s)).strip().lower()
    t = re.sub(r"\s+", " ", t)
    return t


def _is_skipped_path(path: str) -> bool:
    parts = path.split(os.sep)
    return any(p in _SKIP_DIR_NAMES or p.startswith("._") for p in parts)


def find_first_excel(root: str) -> Optional[str]:
    candidates: list[str] = []
    for dirpath, _, filenames in os.walk(root):
        if _is_skipped_path(dirpath):
            continue
        for fn in filenames:
            low = fn.lower()
            if low.endswith((".xlsx", ".xlsm")):
                candidates.append(os.path.join(dirpath, fn))
    if not candidates:
        return None
    candidates.sort()
    return candidates[0]


def scan_s_folders(root: str) -> tuple[dict[str, str], list[str]]:
    """Trả về map S0001 -> đường dẫn tuyệt đối."""
    mapping: dict[str, str] = {}
    errors: list[str] = []
    for dirpath, dirnames, _ in os.walk(root):
        if _is_skipped_path(dirpath):
            continue
        for d in dirnames:
            m = _S_DIR_RE.match(d)
            if not m:
                continue
            key = f"S{m.group(1)}"
            full = os.path.join(dirpath, d)
            if key in mapping and os.path.normpath(mapping[key]) != os.path.normpath(full):
                errors.append(f"Trùng thư mục {key}: {mapping[key]} và {full}")
            else:
                mapping[key] = full
    return mapping, errors


def scan_bmp_files(root: str) -> tuple[dict[int, str], list[str]]:
    """Số thứ tự ảnh (1-based như trong tên 001) -> một đường dẫn file."""
    by_num: dict[int, str] = {}
    dup: list[str] = []
    for dirpath, _, filenames in os.walk(root):
        if _is_skipped_path(dirpath):
            continue
        for fn in filenames:
            m = _BMP_RE.match(fn)
            if not m:
                continue
            n = int(m.group(1))
            full = os.path.join(dirpath, fn)
            if n in by_num and os.path.normpath(by_num[n]) != os.path.normpath(full):
                dup.append(f"Trùng PS-SD{n:03d}: {by_num[n]} và {full}")
            else:
                by_num[n] = full
    return by_num, dup


def file_code_to_s_name(raw: Any) -> Optional[str]:
    if raw is None or (isinstance(raw, float) and pd.isna(raw)):
        return None
    s = str(raw).strip()
    if not s:
        return None
    m = re.match(r"^S\s*(\d{1,4})$", s, re.I)
    if m:
        return f"S{int(m.group(1)):04d}"
    digits = re.sub(r"\D", "", s)
    if not digits:
        return None
    return f"S{int(digits):04d}"


def _to_int_img(v: Any) -> Optional[int]:
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return None
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        if isinstance(v, float) and v.is_integer():
            return int(v)
        if isinstance(v, int):
            return int(v)
    s = str(v).strip()
    if not s:
        return None
    digits = re.sub(r"\D", "", s)
    if not digits:
        return None
    return int(digits)


def bmp_basename_for_index(n: int) -> str:
    return f"PS-SD{n:03d}.BMP"


def sanitize_device_folder(name: Any) -> str:
    if name is None or (isinstance(name, float) and pd.isna(name)):
        raise ValueError("Tên thiết bị trống.")
    s = unicodedata.normalize("NFKC", str(name)).strip()
    s = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", s)
    s = re.sub(r"\s+", " ", s).strip()
    s = s.rstrip(" .")
    if not s or s in {".", ".."}:
        raise ValueError("Tên thiết bị không hợp lệ sau khi chuẩn hóa.")
    base = s.split(".")[0].upper()
    if base in _WIN_RESERVED:
        s = f"_{s}"
    if len(s) > 180:
        s = s[:180].rstrip(" .")
    return s


def _unique_name(base: str, used: set[str]) -> str:
    if base not in used:
        used.add(base)
        return base
    i = 2
    while True:
        cand = f"{base}_{i}"
        if cand not in used:
            used.add(cand)
            return cand
        i += 1


def _resolve_columns(df: pd.DataFrame) -> dict[str, Optional[str]]:
    """Map internal key -> actual column name (một số cột là tuỳ chọn → ``None``)."""
    cols = { _norm_key(c): c for c in df.columns }
    def pick(*aliases: str) -> Optional[str]:
        for a in aliases:
            k = _norm_key(a)
            if k in cols:
                return cols[k]
        for a in aliases:
            nk = _norm_key(a).replace(" ", "")
            for ck, orig in cols.items():
                if ck.replace(" ", "") == nk:
                    return orig
        return None

    device_col = pick("tên thiết bị", "ten thiet bi", "tên thiet bi", "thiết bị", "thiet bi")
    file_col = pick("file")
    img_col = pick("img")
    img_end_col = pick("img end", "img_end", "img end ", "img kết thúc", "img ket thuc")
    kind_col = pick("loại", "loai", "kiểu", "kieu", "type")
    nominal_v_col = pick("điện áp định mức", "dien ap dinh muc", "u định mức", "u dinh muc",
                        "vnom", "u_nom", "u nominal", "nominal voltage")
    remarks_col = pick("nhận xét", "nhan xet", "ghi chú", "ghi chu", "remarks", "notes")

    missing = []
    if not device_col:
        missing.append("Tên thiết bị")
    if not file_col:
        missing.append("File")
    if not img_col:
        missing.append("IMG")
    if not img_end_col:
        missing.append("IMG end")
    if missing:
        raise ValueError("Thiếu cột Excel: " + ", ".join(missing) + f". Các cột hiện có: {list(df.columns)}")
    return {
        "device": device_col,
        "file": file_col,
        "img": img_col,
        "img_end": img_end_col,
        "kind": kind_col,
        "nominal_voltage": nominal_v_col,
        "remarks": remarks_col,
    }


@dataclass
class RowPlan:
    device_raw: str
    folder_name: str
    s_key: str
    img_start: int
    img_end: int
    kind: Optional[str] = None
    nominal_voltage: Optional[float] = None
    remarks: str = ""


def read_plans_from_excel(excel_path: str) -> tuple[list[RowPlan], list[str]]:
    warnings: list[str] = []
    try:
        df = pd.read_excel(excel_path, header=0, engine="openpyxl")
    except Exception as e:
        raise ValueError(f"Không đọc được Excel: {e}") from e
    if df.empty:
        raise ValueError("File Excel không có dữ liệu.")
    colmap = _resolve_columns(df)
    used_names: set[str] = set()
    plans: list[RowPlan] = []
    used_s: set[str] = set()

    for idx, row in df.iterrows():
        try:
            dev_raw = row[colmap["device"]]
            s_name = file_code_to_s_name(row[colmap["file"]])
            i0 = _to_int_img(row[colmap["img"]])
            i1 = _to_int_img(row[colmap["img_end"]])
        except Exception as e:
            warnings.append(f"Dòng {int(idx) + 2}: bỏ qua ({e})")
            continue
        if s_name is None or i0 is None or i1 is None:
            warnings.append(f"Dòng {int(idx) + 2}: thiếu File/IMG/IMG end — bỏ qua.")
            continue
        if i1 < i0:
            raise ValueError(f"Dòng {int(idx) + 2}: IMG end ({i1}) nhỏ hơn IMG ({i0}).")
        if s_name in used_s:
            raise ValueError(f"Dòng {int(idx) + 2}: mã thư mục {s_name} bị lặp trong Excel.")
        used_s.add(s_name)
        try:
            folder = sanitize_device_folder(dev_raw)
        except ValueError as e:
            raise ValueError(f"Dòng {int(idx) + 2}: {e}") from e
        folder = _unique_name(folder, used_names)
        kind_raw = row[colmap["kind"]] if colmap.get("kind") else None
        kind = None
        if kind_raw is not None and not (isinstance(kind_raw, float) and pd.isna(kind_raw)):
            k = _norm_key(kind_raw)
            if k in {"mba", "máy biến áp", "may bien ap", "transformer", "tr"}:
                kind = "mba"
            elif k in {"device", "thiết bị", "thiet bi"}:
                kind = "device"
        nom_v = None
        if colmap.get("nominal_voltage"):
            nv_raw = row[colmap["nominal_voltage"]]
            if nv_raw is not None and not (isinstance(nv_raw, float) and pd.isna(nv_raw)):
                try:
                    nom_v = float(re.sub(r"[^0-9.]", "", str(nv_raw)) or "nan")
                    if nom_v != nom_v:
                        nom_v = None
                except Exception:
                    nom_v = None
        remarks = ""
        if colmap.get("remarks"):
            r_raw = row[colmap["remarks"]]
            if r_raw is not None and not (isinstance(r_raw, float) and pd.isna(r_raw)):
                remarks = str(r_raw).strip()
        plans.append(
            RowPlan(
                device_raw=str(dev_raw).strip(),
                folder_name=folder,
                s_key=s_name,
                img_start=i0,
                img_end=i1,
                kind=kind,
                nominal_voltage=nom_v,
                remarks=remarks,
            )
        )
    if not plans:
        raise ValueError("Không có dòng hợp lệ nào trong Excel (sau khi lọc).")
    return plans, warnings


def _ranges_overlap(a0: int, a1: int, b0: int, b1: int) -> bool:
    return not (a1 < b0 or b1 < a0)


def validate_plans_against_fs(
    plans: list[RowPlan],
    s_map: dict[str, str],
    bmp_map: dict[int, str],
) -> list[str]:
    errors: list[str] = []
    for p in plans:
        if p.s_key not in s_map:
            errors.append(f"Thiết bị «{p.device_raw}»: không có thư mục {p.s_key} trong ZIP.")
        for n in range(p.img_start, p.img_end + 1):
            if n not in bmp_map:
                errors.append(
                    f"Thiết bị «{p.device_raw}»: thiếu ảnh {bmp_basename_for_index(n)}."
                )
    for i, pa in enumerate(plans):
        for pb in plans[i + 1 :]:
            if _ranges_overlap(pa.img_start, pa.img_end, pb.img_start, pb.img_end):
                errors.append(
                    f"Trùng dải ảnh: «{pa.device_raw}» ({pa.img_start}–{pa.img_end}) "
                    f"và «{pb.device_raw}» ({pb.img_start}–{pb.img_end})."
                )
    return errors


def build_project_output(
    extract_root: str,
    output_parent: str,
    plans: list[RowPlan],
    s_map: dict[str, str],
    bmp_map: dict[int, str],
) -> str:
    """
    Tạo thư mục Project_Output trong output_parent, trả về đường dẫn Project_Output.
    """
    out_root = os.path.join(output_parent, "Project_Output")
    os.makedirs(out_root, exist_ok=True)
    for p in plans:
        src_dir = s_map[p.s_key]
        dest_dir = os.path.join(out_root, p.folder_name)
        if os.path.exists(dest_dir):
            shutil.rmtree(dest_dir)
        shutil.copytree(src_dir, dest_dir)

        for n in range(p.img_start, p.img_end + 1):
            src_bmp = bmp_map[n]
            norm_src = os.path.normpath(src_bmp)
            base = bmp_basename_for_index(n)
            dest_bmp = os.path.join(dest_dir, base)

            if os.path.normpath(dest_bmp) == norm_src:
                continue

            if os.path.dirname(norm_src) == os.path.normpath(dest_dir):
                continue

            shutil.copy2(src_bmp, dest_bmp)
            try:
                os.remove(src_bmp)
            except OSError:
                pass

    return out_root


def zip_directory(folder: str, zip_path: str) -> None:
    """Nén `folder` (ví dụ .../Project_Output) sao cho gốc ZIP là tên thư mục đó."""
    parent = os.path.dirname(folder)
    with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        for root, _, files in os.walk(folder):
            for fn in files:
                fp = os.path.join(root, fn)
                arc = os.path.relpath(fp, parent)
                zf.write(fp, arcname=arc)


_MBA_TEMPLATE_DOCX = os.path.normpath(
    os.path.join(os.path.dirname(os.path.abspath(__file__)),
                 "..", "..", "static", "word-template", "mba.docx")
)
_DEVICE_TEMPLATE_DOCX = os.path.normpath(
    os.path.join(os.path.dirname(os.path.abspath(__file__)),
                 "..", "..", "static", "word-template", "device.docx")
)


def _generate_field_word_report(
    project_out_dir: str,
    plans: list[RowPlan],
    output_docx: str,
) -> list[str]:
    """Sinh file Word tổng hợp tại ``output_docx`` từ ``Project_Output/``.

    Trả về danh sách cảnh báo. Nếu không thể tạo (thiếu template, không có
    plan hợp lệ, lỗi import…) → trả ``["..."]`` và **không** ném exception
    để pipeline chính vẫn tiếp tục.
    """
    warnings: list[str] = []
    if not plans:
        return ["Không có thiết bị nào để dựng báo cáo Word."]
    if not os.path.isfile(_MBA_TEMPLATE_DOCX):
        return [f"Thiếu template Word MBA: {_MBA_TEMPLATE_DOCX}"]
    if not os.path.isfile(_DEVICE_TEMPLATE_DOCX):
        return [f"Thiếu template Word device: {_DEVICE_TEMPLATE_DOCX}"]

    try:
        from modules.report.gen_word import build_field_word_report
    except Exception as e:
        return [f"Không thể import modules.report.gen_word: {e}"]

    devices = [
        {
            "name": p.device_raw or p.folder_name,
            "folder": p.folder_name,
            "kind": p.kind,
            "nominal_voltage": p.nominal_voltage,
            "remarks": p.remarks,
        }
        for p in plans
    ]
    try:
        _, w = build_field_word_report(
            project_out_dir,
            output_docx,
            mba_template=_MBA_TEMPLATE_DOCX,
            device_template=_DEVICE_TEMPLATE_DOCX,
            devices=devices,
        )
        warnings.extend(w)
    except Exception as e:
        warnings.append(f"Không sinh được báo cáo Word: {e}")
    return warnings


def process_field_zip_bytes(zip_bytes: bytes, work_dir: str) -> tuple[str, list[str], list[str]]:
    """
    Giải nén zip_bytes vào work_dir, xử lý, tạo file ZIP kết quả trong work_dir.
    Trả về (đường_dẫn_zip_kết_quả, warnings, errors_fatal).
    errors_fatal rỗng nếu thành công.

    Ngoài việc tổ chức ảnh + thư mục, bước này còn sinh thêm
    ``Project_Output/BaoCao_KEW.docx`` tổng hợp số liệu từ file INPS của
    từng thiết bị (nếu có).
    """
    warnings: list[str] = []
    extract = os.path.join(work_dir, "in")
    os.makedirs(extract, exist_ok=True)
    bio = io.BytesIO(zip_bytes)
    with zipfile.ZipFile(bio, "r") as zf:
        zf.extractall(extract)

    excel_path = find_first_excel(extract)
    if not excel_path:
        return "", warnings, ["Không tìm thấy file Excel (.xlsx/.xlsm/.xls) trong ZIP."]

    s_map, s_err = scan_s_folders(extract)
    warnings.extend(s_err)
    bmp_map, bmp_dup = scan_bmp_files(extract)
    warnings.extend(bmp_dup)

    try:
        plans, w2 = read_plans_from_excel(excel_path)
        warnings.extend(w2)
    except ValueError as e:
        return "", warnings, [str(e)]

    fatal = validate_plans_against_fs(plans, s_map, bmp_map)
    if fatal:
        return "", warnings, fatal

    staging = os.path.join(work_dir, "staging")
    os.makedirs(staging, exist_ok=True)
    proj = os.path.join(staging, "Project_Output")
    build_project_output(extract, staging, plans, s_map, bmp_map)

    # Sinh báo cáo Word tổng hợp (không chặn pipeline nếu lỗi).
    report_path = os.path.join(proj, "BaoCao_KEW.docx")
    warnings.extend(_generate_field_word_report(proj, plans, report_path))

    out_zip = os.path.join(work_dir, "KEW_HoSoDaXuLy.zip")
    zip_directory(proj, out_zip)
    return out_zip, warnings, []
