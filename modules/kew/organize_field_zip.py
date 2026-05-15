"""
Tổ chức hồ sơ đo KEW6315 từ ZIP: đọc Excel hiện trường (.xlsx), đổi tên thư mục Sxxxx,
chuyển ảnh PS-SDxxx.BMP vào đúng thư mục thiết bị, nén lại Project_Output.zip.

Excel hiện trường chỉ hỗ trợ một bộ cột cố định (tên cột không phân biệt hoa thường):
``stt``, ``name``, ``file``, ``img``, ``imgend``, ``imgomit``, ``type``, ``pdm``,
``p``, ``pf``, ``i1``, ``i2``, ``i3``, ``di``, ``thd``, ``tdd``,
``current_char``, ``u_min``, ``u_max``, ``delta_u``.

Bước tổ chức ZIP dùng ``name`` / ``file`` / ``img`` / ``imgend`` / ``imgomit``
(tuỳ chọn điền: chỉ số ảnh trong dải bị loại, vd ``944, 945`` hoặc ``PS-SD944``);
``stt`` và các cột còn lại dành cho bước Word (thứ tự thiết bị, metadata).

Các cột đã có sẵn từ trước, được tái dùng để sinh nhận xét (không dùng INPS):
- ``pf``: Hệ số công suất trung bình (cosφ).
- ``di``: Độ lệch pha (mất cân bằng) dòng điện lớn nhất (%).
- ``thd``: THD điện áp lớn nhất (%).
- ``tdd``: TDD dòng điện lớn nhất (%).

Các cột mới thêm để sinh nhận xét:
- ``current_char``: Đặc tính dòng điện — một trong: "Ổn định" / "Dao động nhẹ" /
  "Biến đổi liên tục" / "Chu kỳ Load-Unload".
- ``u_min``, ``u_max``: Điện áp đo được thấp nhất / cao nhất (V); tool tự tính δU.
- ``delta_u``: Độ lệch pha (mất cân bằng) điện áp lớn nhất (%).
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

# Một file Excel hiện trường duy nhất: đủ 20 cột (so khớp sau chuẩn hóa NFKC + chữ thường).
FIELD_XLSX_HEADERS: tuple[str, ...] = (
    "stt",
    "name",
    "file",
    "img",
    "imgend",
    "imgomit",
    "type",
    "pdm",
    "p",
    "pf",          # Hệ số công suất (cosφ) — tái dùng cho sinh nhận xét
    "i1",
    "i2",
    "i3",
    "di",          # Độ lệch pha dòng điện lớn nhất (%) — tái dùng cho sinh nhận xét
    "thd",         # THD điện áp lớn nhất (%) — tái dùng cho sinh nhận xét
    "tdd",         # TDD dòng điện lớn nhất (%) — tái dùng cho sinh nhận xét
    # ── Cột mới: sinh nhận xét từ hiện trường ───────────────────────────
    "current_char",  # Đặc tính dòng điện (Ổn định / Dao động nhẹ / Biến đổi liên tục / Chu kỳ Load-Unload)
    "u_min",         # Điện áp đo thấp nhất (V)
    "u_max",         # Điện áp đo cao nhất (V)
    "delta_u",       # Độ lệch pha (mất cân bằng) điện áp lớn nhất (%)
)

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


def _parse_img_omit(raw: Any) -> tuple[frozenset[int], list[str]]:
    """Đọc cột ``imgomit``: danh sách chỉ số ``PS-SDxxx`` cần bỏ qua trong dải ``img``–``imgend``.

    Hỗ trợ: ``944``, ``944,945``, ``944+945``, ``PS-SD944``, ``PS-SD944.BMP``, số nguyên trong ô Excel.
    """
    warnings: list[str] = []
    if raw is None or (isinstance(raw, float) and pd.isna(raw)):
        return frozenset(), warnings
    if isinstance(raw, (int, float)) and not isinstance(raw, bool):
        if isinstance(raw, float):
            if not raw.is_integer():
                warnings.append("imgomit: ô chứa số thập phân — chỉ dùng chỉ số ảnh nguyên, bỏ qua giá trị này.")
                return frozenset(), warnings
            return frozenset({int(raw)}), warnings
        return frozenset({raw}), warnings
    s = str(raw).strip()
    if not s:
        return frozenset(), warnings
    out: set[int] = set()
    for tok in re.split(r"[,;+/\s|]+", s):
        tok = tok.strip()
        if not tok:
            continue
        tm = re.search(r"PS-SD(\d{1,4})(?:\.BMP)?", tok, re.IGNORECASE)
        if tm:
            out.add(int(tm.group(1)))
            continue
        dm = re.fullmatch(r"(\d{1,4})", re.sub(r"\s+", "", tok))
        if dm:
            out.add(int(dm.group(1)))
            continue
        warnings.append(f"imgomit: không hiểu mục «{tok}» — bỏ qua token.")
    return frozenset(out), warnings


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
    # Dùng NFC để tên thư mục trùng khóa với Excel / macOS (HFS+ hay lưu NFD).
    return unicodedata.normalize("NFC", s)


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


def resolve_field_excel_column_map(df: pd.DataFrame) -> dict[str, Any]:
    """Map tên cột logic (chữ thường) → tên cột gốc trong DataFrame.

    Chỉ chấp nhận đúng :data:`FIELD_XLSX_HEADERS`; thiếu cột hoặc trùng tên sau
    chuẩn hóa → ``ValueError``.
    """
    seen: dict[str, Any] = {}
    for c in df.columns:
        k = _norm_key(str(c))
        if not k:
            continue
        if k in seen and seen[k] != c:
            raise ValueError(
                f"Hai cột trùng tên sau khi chuẩn hóa ({k!r}): {seen[k]!r} và {c!r}."
            )
        if k not in seen:
            seen[k] = c
    unknown = sorted(k for k in seen if k not in set(FIELD_XLSX_HEADERS))
    if unknown:
        raise ValueError(
            "File Excel có cột không được hỗ trợ: "
            + ", ".join(unknown)
            + ". Chỉ chấp nhận: "
            + ", ".join(FIELD_XLSX_HEADERS)
        )
    missing = [h for h in FIELD_XLSX_HEADERS if h not in seen]
    if missing:
        raise ValueError(
            "File Excel hiện trường cần đủ các cột: "
            + ", ".join(FIELD_XLSX_HEADERS)
            + f". Thiếu: {', '.join(missing)}. Các cột hiện có: {list(df.columns)}"
        )
    return {h: seen[h] for h in FIELD_XLSX_HEADERS}


@dataclass
class RowPlan:
    device_raw: str
    folder_name: str
    s_key: str
    img_start: int
    img_end: int
    img_omit: frozenset[int]


def read_plans_from_excel(excel_path: str) -> tuple[list[RowPlan], list[str]]:
    warnings: list[str] = []
    try:
        df = pd.read_excel(excel_path, header=0, engine="openpyxl")
    except Exception as e:
        raise ValueError(f"Không đọc được Excel: {e}") from e
    if df.empty:
        raise ValueError("File Excel không có dữ liệu.")
    colmap = resolve_field_excel_column_map(df)
    used_names: set[str] = set()
    plans: list[RowPlan] = []
    used_s: set[str] = set()

    for idx, row in df.iterrows():
        try:
            dev_raw = row[colmap["name"]]
            s_name = file_code_to_s_name(row[colmap["file"]])
            i0 = _to_int_img(row[colmap["img"]])
            i1 = _to_int_img(row[colmap["imgend"]])
        except Exception as e:
            warnings.append(f"Dòng {int(idx) + 2}: bỏ qua ({e})")
            continue
        if s_name is None or i0 is None or i1 is None:
            warnings.append(f"Dòng {int(idx) + 2}: thiếu file/img/imgend — bỏ qua.")
            continue
        if i1 < i0:
            raise ValueError(f"Dòng {int(idx) + 2}: IMG end ({i1}) nhỏ hơn IMG ({i0}).")
        if s_name in used_s:
            raise ValueError(f"Dòng {int(idx) + 2}: mã thư mục {s_name} bị lặp trong Excel.")
        used_s.add(s_name)
        omit_all, w_omit = _parse_img_omit(row[colmap["imgomit"]])
        warnings.extend(w_omit)
        omit_eff = frozenset(n for n in omit_all if i0 <= n <= i1)
        outs = sorted(n for n in omit_all if n < i0 or n > i1)
        if outs:
            warnings.append(
                f"Dòng {int(idx) + 2}: imgomit có chỉ số ngoài dải {i0}–{i1}: "
                f"{', '.join(str(x) for x in outs)} (bỏ qua các mục ngoài dải)."
            )
        try:
            folder = sanitize_device_folder(dev_raw)
        except ValueError as e:
            raise ValueError(f"Dòng {int(idx) + 2}: {e}") from e
        folder = _unique_name(folder, used_names)
        plans.append(
            RowPlan(
                device_raw=str(dev_raw).strip(),
                folder_name=folder,
                s_key=s_name,
                img_start=i0,
                img_end=i1,
                img_omit=omit_eff,
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
        kept = [n for n in range(p.img_start, p.img_end + 1) if n not in p.img_omit]
        if not kept:
            errors.append(
                f"Thiết bị «{p.device_raw}»: imgomit loại hết dải ảnh {p.img_start}–{p.img_end}."
            )
        for n in kept:
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
            if n in p.img_omit:
                continue
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


def process_field_zip_bytes(zip_bytes: bytes, work_dir: str) -> tuple[str, list[str], list[str]]:
    """
    Giải nén zip_bytes vào work_dir, xử lý, tạo file ZIP kết quả trong work_dir.
    Trả về (đường_dẫn_zip_kết_quả, warnings, errors_fatal).
    errors_fatal rỗng nếu thành công.
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
    build_project_output(extract, staging, plans, s_map, bmp_map)

    out_zip = os.path.join(work_dir, "KEW_HoSoDaXuLy.zip")
    proj = os.path.join(staging, "Project_Output")
    zip_directory(proj, out_zip)
    return out_zip, warnings, []
