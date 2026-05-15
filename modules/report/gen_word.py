"""Context docxtpl + sinh báo cáo Word cho luồng "Tạo báo cáo Word".

API chính:
* ``mba(doc, ...)``, ``device(doc, ...)`` — dựng context cho từng template.
* ``merge_rendered_docx`` / ``merge_mba_device_docx`` — ghép nhiều file đã render.
* ``mba_kwargs_from_inps`` / ``mba_kwargs_from_folder`` / ``device_kwargs_from_folder`` —
  tự chọn ảnh từ thư mục thiết bị (**bắt buộc** ``a.png`` + PS-SDxxx.BMP). Với MBA,
  ``mba_kwargs_from_folder`` tự tìm ``INPS*.KEW`` (cùng quy ước ``find_file`` như phân tích KEW
  và như luồng Excel MBA) rồi gọi ``mba_kwargs_from_inps``; thiếu INPS thì bảng dùng ``"—"``.
  **Nhận xét văn bản** (``remarks_mba`` / ``remarks_device``): tự sinh từ INPS (+ INIS khi có
  công suất định mức cho MBA) theo ``quy-tac.md`` / ``quy-tac-2.md``; có thể ghép thêm ghi chú
  cột Excel (P, PF, …); nếu ô Excel chứa đoạn bắt đầu bằng ``Nhận xét:`` thì dùng nguyên văn thủ công.
* ``build_field_word_report`` — quét một thư mục ``Project_Output/`` rồi xuất
  1 file Word duy nhất gồm nhiều MBA / device.
* ``build_word_report_from_zip`` — entry-point cho API: nhận ZIP đã tổ chức
  (output của "Xử lý file sơ bộ"), tự dò metadata trong Excel kèm (nếu có),
  trả về đường dẫn báo cáo Word.

Tham số / khóa template — xem ``modules/report/context_keys.json``.
"""

from __future__ import annotations

import io
import re
import unicodedata
import zipfile
from pathlib import Path
from shutil import copy2
from tempfile import TemporaryDirectory
from typing import Iterable, Literal, Mapping, Sequence

from docx import Document
from docx.enum.text import WD_BREAK
from docx.shared import Mm
from docxcompose.composer import Composer
from docxtpl import DocxTemplate, InlineImage

__all__ = [
    "mba",
    "device",
    "merge_rendered_docx",
    "merge_mba_device_docx",
    "mba_kwargs_from_inps",
    "mba_kwargs_from_folder",
    "device_kwargs_from_folder",
    "build_field_word_report",
    "build_word_report_from_zip",
    "DEFAULT_MBA_TEMPLATE",
    "DEFAULT_DEVICE_TEMPLATE",
]

_REPO_ROOT = Path(__file__).resolve().parents[2]
DEFAULT_MBA_TEMPLATE = _REPO_ROOT / "static" / "word-template" / "mba.docx"
DEFAULT_DEVICE_TEMPLATE = _REPO_ROOT / "static" / "word-template" / "device.docx"

WIDTH_LARGE, WIDTH_SMALL, HEIGHT_MBA = Mm(109.6), Mm(53.8), Mm(40.5)
WIDTH_A, WIDTH_SUB, HEIGHT_A, HEIGHT_SUB = Mm(166.3), Mm(54.3), Mm(60.0), Mm(41.3)

SectionKind = Literal["mba", "device"]

_DEVICE_KIND_LABELS = frozenset(
    {"device", "thiết bị", "thiet bi", "tủ", "tu", "khac", "khác"}
)
# Các cột ghi chú cũ — giữ lại để đọc nhưng không ghép vào remarks tự động nữa.
_EXCEL_METRIC_REMARKS: tuple[tuple[str, str], ...] = (
    ("p", "P"),
    ("pf", "PF"),
    ("i1", "I1"),
    ("i2", "I2"),
    ("i3", "I3"),
    ("di", "ΔI"),
    ("thd", "THD"),
    ("tdd", "TDD"),
)
# Các cột đặc tính hiện trường dùng để sinh nhận xét (thay thế INPS).
_FIELD_PARAM_COLS: tuple[str, ...] = (
    "current_char", "u_min", "u_max", "i_max",
    "delta_u", "delta_i", "p", "cos_phi",
    "thd", "tdd", "pdm",
)

# Giới hạn / tiêu chuẩn dùng cho cột "Đánh giá" trong báo cáo MBA.
_V_DEV_LIMIT_PCT = 5.0
_PF_LIMIT = 0.9
_THDV_LIMIT_PCT = 8.0
_TDD_LIMIT_PCT = 12.0
# Điện áp danh định cho cột đánh giá lệch % U12: -5% ≤ δ ≤ +5% (so với 400 V), không lấy từ Excel/metadata.
_MBA_NOMINAL_VOLTAGE_V = 400.0
# Biến động dòng (INPS min/max/avg) — cùng ý tưởng với analyse_kew.CONFIG defaults.
_I_SPREAD_STABLE_PCT = 15.0
_I_SPREAD_MODERATE_PCT = 50.0
# Ảnh theo dải Excel / organize_field_zip: PS-SD641.BMP …
_BMP_RE = re.compile(r"^PS-SD(\d+)\.BMP$", re.IGNORECASE)


# ════════════════════════════════════════════════════════════════════
#                     Context builders (template)
# ════════════════════════════════════════════════════════════════════


def _inline(doc: DocxTemplate, path: str, height, width) -> InlineImage:
    return InlineImage(doc, path, height=height, width=width)


def mba(
    doc: DocxTemplate,
    *,
    name: str,
    imga: str,
    img1: str,
    img2: str,
    img4: str,
    img6: str,
    cap_fig_mba: str,
    remarks_mba: str,
    cap_tab_mba: str,
    # Điện áp
    u12max: str,
    u12min: str,
    u12avg: str,
    u12eval: str,
    u23max: str,
    u23min: str,
    u23avg: str,
    u31max: str,
    u31min: str,
    u31avg: str,
    # Dòng điện
    i1max: str,
    i1min: str,
    i1avg: str,
    i2max: str,
    i2min: str,
    i2avg: str,
    i3max: str,
    i3min: str,
    i3avg: str,
    # Độ lệch pha điện áp / dòng & hệ số công suất
    dumax: str,
    dumin: str,
    duavg: str,
    dueval: str,
    dimax: str,
    dimin: str,
    diavg: str,
    pfmax: str,
    pfmin: str,
    pfavg: str,
    pfeval: str,
    # Công suất P, Q, S
    pmax: str,
    pmin: str,
    pavg: str,
    qmax: str,
    qmin: str,
    qavg: str,
    smax: str,
    smin: str,
    savg: str,
    # Sóng hài THD
    thd1max: str,
    thd1min: str,
    thd1avg: str,
    thdeval: str,
    thd2max: str,
    thd2min: str,
    thd2avg: str,
    thd3max: str,
    thd3min: str,
    thd3avg: str,
    # Sóng hài TDD
    tdd1max: str,
    tdd1min: str,
    tdd1avg: str,
    tddeval: str,
    tdd2max: str,
    tdd2min: str,
    tdd2avg: str,
    tdd3max: str,
    tdd3min: str,
    tdd3avg: str,
) -> dict:
    return {
        "mba_name": name,
        "imga": _inline(doc, imga, HEIGHT_MBA, WIDTH_LARGE),
        "img1": _inline(doc, img1, HEIGHT_MBA, WIDTH_SMALL),
        "img2": _inline(doc, img2, HEIGHT_MBA, WIDTH_SMALL),
        "img4": _inline(doc, img4, HEIGHT_MBA, WIDTH_SMALL),
        "img6": _inline(doc, img6, HEIGHT_MBA, WIDTH_SMALL),
        "cap_fig_mba": cap_fig_mba,
        "remarks_mba": remarks_mba,
        "cap_tab_mba": cap_tab_mba,
        "u12max": u12max, "u12min": u12min, "u12avg": u12avg, "u12eval": u12eval,
        "u23max": u23max, "u23min": u23min, "u23avg": u23avg,
        "u31max": u31max, "u31min": u31min, "u31avg": u31avg,
        "i1max": i1max, "i1min": i1min, "i1avg": i1avg,
        "i2max": i2max, "i2min": i2min, "i2avg": i2avg,
        "i3max": i3max, "i3min": i3min, "i3avg": i3avg,
        "dumax": dumax, "dumin": dumin, "duavg": duavg, "dueval": dueval,
        "dimax": dimax, "dimin": dimin, "diavg": diavg,
        "pfmax": pfmax, "pfmin": pfmin, "pfavg": pfavg, "pfeval": pfeval,
        "pmax": pmax, "pmin": pmin, "pavg": pavg,
        "qmax": qmax, "qmin": qmin, "qavg": qavg,
        "smax": smax, "smin": smin, "savg": savg,
        "thd1max": thd1max, "thd1min": thd1min, "thd1avg": thd1avg, "thdeval": thdeval,
        "thd2max": thd2max, "thd2min": thd2min, "thd2avg": thd2avg,
        "thd3max": thd3max, "thd3min": thd3min, "thd3avg": thd3avg,
        "tdd1max": tdd1max, "tdd1min": tdd1min, "tdd1avg": tdd1avg, "tddeval": tddeval,
        "tdd2max": tdd2max, "tdd2min": tdd2min, "tdd2avg": tdd2avg,
        "tdd3max": tdd3max, "tdd3min": tdd3min, "tdd3avg": tdd3avg,
    }


def device(
    doc: DocxTemplate,
    *,
    name: str,
    imga: str,
    img1: str,
    img2: str,
    img3: str,
    img4: str,
    img5: str,
    img6: str,
    cap_device: str,
    remarks_device: str,
) -> dict:
    return {
        "device_name": name,
        "imga": _inline(doc, imga, HEIGHT_A, WIDTH_A),
        "img1": _inline(doc, img1, HEIGHT_SUB, WIDTH_SUB),
        "img2": _inline(doc, img2, HEIGHT_SUB, WIDTH_SUB),
        "img3": _inline(doc, img3, HEIGHT_SUB, WIDTH_SUB),
        "img4": _inline(doc, img4, HEIGHT_SUB, WIDTH_SUB),
        "img5": _inline(doc, img5, HEIGHT_SUB, WIDTH_SUB),
        "img6": _inline(doc, img6, HEIGHT_SUB, WIDTH_SUB),
        "cap1": cap_device,
        "remarks": remarks_device,
    }


# ════════════════════════════════════════════════════════════════════
#                            Format helpers
# ════════════════════════════════════════════════════════════════════


def _f(v, d: int = 1) -> str:
    """Định dạng số kiểu Việt (dấu phẩy thập phân). Trả ``"—"`` nếu thiếu."""
    if v is None:
        return "—"
    try:
        x = float(v)
    except (TypeError, ValueError):
        return "—"
    if x != x:  # NaN
        return "—"
    return f"{x:.{d}f}".replace(".", ",")


def _tri(d: Mapping, dec: int = 1) -> tuple[str, str, str]:
    """min / max / avg → ba chuỗi định dạng ``_f``."""
    return _f(d.get("max"), dec), _f(d.get("min"), dec), _f(d.get("avg"), dec)


# ════════════════════════════════════════════════════════════════════
#                       INPS aggregation helpers
# ════════════════════════════════════════════════════════════════════


def _parse_inps(inps_path: str | Path) -> dict[str, dict]:
    """Đọc một file INPS .KEW và trả về ``{cột_avg: {min, max, avg}}``.

    Sử dụng lại ``parse_inps`` / ``analyse_inps`` trong module phân tích KEW;
    ``min/max`` ưu tiên cột ``recorded_min/max`` của thiết bị (giá trị cực trị
    được KEW6315 ghi lại trong từng khoảng 1 s, chứ không phải min/max của
    chuỗi trung bình).
    """
    # Lazy import để gen_word có thể đứng độc lập khi không cần phân tích.
    from modules.kew.analyse_kew import analyse_inps, parse_inps  # type: ignore

    _, df = parse_inps(str(inps_path))
    if df is None:
        return {}
    raw = analyse_inps(df)
    out: dict[str, dict] = {}
    for col, d in raw.items():
        if not isinstance(d, Mapping):
            continue
        avg = d.get("avg")
        rec_min = d.get("recorded_min")
        rec_max = d.get("recorded_max")
        out[col] = {
            "avg": avg,
            "min": rec_min if rec_min is not None else d.get("min"),
            "max": rec_max if rec_max is not None else d.get("max"),
        }
    return out


def _pick(stats: Mapping[str, dict], *keys: str) -> dict:
    """Trả về stat của khóa khớp đầu tiên trong ``stats`` (case-insensitive)."""
    upper = {k.upper(): v for k, v in stats.items()}
    for k in keys:
        v = upper.get(k.upper())
        if v:
            return v
    return {}


def _pick_total(stats: Mapping[str, dict], prefix: str, unit_substr: str) -> dict:
    """Lấy stat của cột tổng ``AVG_<prefix>[<unit>]`` (không có chỉ số pha)."""
    for col, d in stats.items():
        if not col.upper().startswith(f"AVG_{prefix.upper()}"):
            continue
        if unit_substr.upper() not in col.upper():
            continue
        if any(ch in col for ch in ("1", "2", "3")):
            continue
        return d
    return {}


def _scale(d: Mapping[str, float] | None, k: float) -> dict:
    if not d:
        return {}
    return {
        "avg": d.get("avg") * k if d.get("avg") is not None else None,
        "min": d.get("min") * k if d.get("min") is not None else None,
        "max": d.get("max") * k if d.get("max") is not None else None,
    }


def _eval_voltage(u_max, u_min, vref: float) -> tuple[str, float, float, float]:
    """Trả về (đánh giá, δmax, δmin, δavg) cho dải điện áp so với ``vref`` (±5%)."""
    if u_max is None or u_min is None or vref <= 0:
        return "—", None, None, None
    dmax = (u_max - vref) / vref * 100
    dmin = (u_min - vref) / vref * 100
    ok = abs(dmax) <= _V_DEV_LIMIT_PCT and abs(dmin) <= _V_DEV_LIMIT_PCT
    dabs_max = max(abs(dmax), abs(dmin))
    dabs_min = min(abs(dmax), abs(dmin))
    return ("Đạt" if ok else "Không đạt"), dabs_max, dabs_min, (dabs_max + dabs_min) / 2


def _eval_pf(pf_avg) -> str:
    if pf_avg is None:
        return "—"
    return "Đạt" if abs(pf_avg) >= _PF_LIMIT else "Không đạt"


def _eval_thd(values: Iterable[float | None], limit: float) -> str:
    vals = [v for v in values if v is not None]
    if not vals:
        return "—"
    return "Đạt" if max(vals) < limit else "Không đạt"


def _fmt_remark_pct(v: float | None, decimals: int = 2) -> str:
    if v is None or v != v:  # NaN
        return "—"
    return f"{v:.{decimals}f}".replace(".", ",")


def _fmt_remark_voltage(v: float | None, decimals: int = 1) -> str:
    if v is None or v != v:
        return "—"
    return f"{v:.{decimals}f}".replace(".", ",")


def _is_full_manual_remarks(text: str) -> bool:
    """Người dùng dán cả đoạn nhận xét (không ghép với sinh tự động)."""
    t = text.strip()
    if not t:
        return False
    if "Nhận xét:" in t:
        return True
    if len(t) > 160 and "=" not in t.split("\n", 1)[0]:
        return True
    return False


def _merge_auto_and_excel_notes(auto: str, excel_bits: str) -> str:
    auto = (auto or "").strip()
    bits = (excel_bits or "").strip()
    if not bits:
        return auto
    if not auto:
        return bits
    return f"{auto}\n\nGhi chú hiện trường (Excel): {bits}"


def _rated_kva_from_inis_folder(folder: str | Path | None) -> float | None:
    """Công suất định mức (kVA) từ INIS nếu đọc được."""
    if folder is None:
        return None
    try:
        from modules.kew.analyse_kew import find_file, parse_inis
    except ImportError:
        return None
    fp = find_file(str(folder), "INIS")
    if not fp:
        return None
    meta = parse_inis(fp)
    if not meta:
        return None

    def _one_val(text: str) -> float | None:
        t = str(text).strip().upper().replace(",", ".")
        m = re.search(r"(\d+(?:\.\d+)?)\s*(KVA|MVA|VA)?", t)
        if not m:
            return None
        num = float(m.group(1))
        unit = (m.group(2) or "").upper()
        if unit == "MVA":
            return num * 1000.0
        if unit == "VA":
            return num / 1000.0 if num > 500 else num
        if not unit and num > 2500:
            return num / 1000.0
        return num

    best: float | None = None
    for k, v in meta.items():
        ks = str(k).upper()
        if not any(
            x in ks
            for x in ("KVA", "MVA", "RATING", "CAPACITY", "POWER", "APPARENT", "S[RATED")
        ):
            continue
        val = _one_val(str(v))
        if val is not None and val > 0:
            best = val if best is None else max(best, val)
    if best is not None:
        return best
    for v in meta.values():
        val = _one_val(str(v))
        if val is not None and 10 <= val <= 50000:
            return val if val < 500 else val / 1000.0
    return None


def _current_spread_pct(
    i1: Mapping | None, i2: Mapping | None, i3: Mapping | None,
) -> float | None:
    """Biên độ biến thiên dòng (max-min)/avg theo pha, lấy max các pha."""
    best: float | None = None
    for ix in (i1, i2, i3):
        if not ix:
            continue
        avg = ix.get("avg")
        lo = ix.get("min")
        hi = ix.get("max")
        if avg is None or lo is None or hi is None:
            continue
        try:
            a = float(avg)
        except (TypeError, ValueError):
            continue
        if a <= 0:
            continue
        try:
            sp = (float(hi) - float(lo)) / a * 100.0
        except (TypeError, ValueError):
            continue
        best = sp if best is None else max(best, sp)
    return best


def _waveform_phrase_mba(spread_pct: float | None) -> str:
    if spread_pct is None or spread_pct != spread_pct:
        return "biến đổi liên tục"
    if spread_pct < _I_SPREAD_STABLE_PCT:
        return "biến đổi liên tục với biên độ nhỏ"
    if spread_pct < _I_SPREAD_MODERATE_PCT:
        return "biến đổi liên tục với biên độ nhỏ"
    return "biến đổi liên tục với biên độ lớn"


def _waveform_phrase_device(spread_pct: float | None) -> str:
    if spread_pct is None or spread_pct != spread_pct:
        return "biến đổi liên tục với biên độ nhỏ"
    if spread_pct < _I_SPREAD_STABLE_PCT:
        return "biến đổi liên tục với biên độ nhỏ"
    if spread_pct < _I_SPREAD_MODERATE_PCT:
        return "biến đổi liên tục với biên độ nhỏ"
    return "biến đổi liên tục với biên độ lớn"


def _pf_text_for_remarks(pf_avg: float | None) -> str:
    """Trong đoạn văn: mốc 0,8 — theo quy-tac-2 / quy-tac.md."""
    if pf_avg is None or pf_avg != pf_avg:
        return "chưa xác định đủ từ dữ liệu INPS"
    p = abs(float(pf_avg))
    if p >= 0.995:
        return "rất cao (cosφ ≈ 1, có thể đã lắp đặt tụ bù)"
    if p >= 0.8:
        return "cao (trên 0,8)"
    if p >= 0.5:
        return "trung bình (dưới 0,8)"
    return "thấp (dưới 0,8)"


def _du_rhetorical_mba(uv_max_pct: float | None) -> str:
    """Theo thói quan mẫu MBA: thường «ở mức cao» trừ khi cực nhỏ."""
    if uv_max_pct is None or uv_max_pct != uv_max_pct:
        return "cao"
    return "thấp" if uv_max_pct < 0.1 else "cao"


def _quality_level_mba(
    *,
    u12eval: str,
    dueval: str,
    pfeval: str,
    thdeval: str,
    tddeval: str,
) -> str:
    ok = [x for x in (u12eval, dueval, pfeval, thdeval, tddeval) if x == "Đạt"]
    bad = [x for x in (u12eval, dueval, pfeval, thdeval, tddeval) if x == "Không đạt"]
    if len(ok) >= 4 and not bad:
        return "tốt"
    if len(bad) >= 2:
        return "chưa tốt"
    return "tương đối tốt"


def _quality_level_device(
    *,
    delta_u_ok: bool,
    thd_ok: bool,
    tdd_ok: bool,
    di_ok: bool,
    pf_avg: float | None,
) -> str:
    p = abs(float(pf_avg)) if pf_avg is not None and pf_avg == pf_avg else None
    score = sum([delta_u_ok, thd_ok, tdd_ok, di_ok])
    if p is not None and p >= 0.8:
        score += 1
    if score >= 5:
        return "Tốt"
    if score == 4:
        return "Khá tốt"
    if score >= 2:
        return "Tương đối tốt"
    return "Chưa tốt"


def _line_voltage_range(
    u12: Mapping | None, u23: Mapping | None, u31: Mapping | None,
) -> tuple[float | None, float | None]:
    mins: list[float] = []
    maxs: list[float] = []
    for u in (u12, u23, u31):
        if not u:
            continue
        for key in ("min", "max"):
            v = u.get(key)
            if v is None:
                continue
            try:
                x = float(v)
            except (TypeError, ValueError):
                continue
            if x > 10:
                (mins if key == "min" else maxs).append(x)
    if not mins or not maxs:
        return None, None
    return min(mins), max(maxs)


def _delta_u_line_window_pct(
    u_min: float | None, u_max: float | None, vref: float,
) -> tuple[float | None, float | None, bool]:
    if u_min is None or u_max is None or vref <= 0:
        return None, None, False
    d_lo = (u_min - vref) / vref * 100.0
    d_hi = (u_max - vref) / vref * 100.0
    ok = abs(d_lo) <= _V_DEV_LIMIT_PCT and abs(d_hi) <= _V_DEV_LIMIT_PCT
    return min(d_lo, d_hi), max(d_lo, d_hi), ok


def _thd_tdd_maxes(
    thd1: Mapping | None,
    thd2: Mapping | None,
    thd3: Mapping | None,
    tdd1: Mapping | None,
    tdd2: Mapping | None,
    tdd3: Mapping | None,
) -> tuple[float | None, float | None]:
    thds: list[float] = []
    tdds: list[float] = []
    for d in (thd1, thd2, thd3, tdd1, tdd2, tdd3):
        if not d:
            continue
        m = d.get("max")
        if m is None:
            continue
        try:
            x = float(m)
        except (TypeError, ValueError):
            continue
        if d in (thd1, thd2, thd3):
            thds.append(x)
        else:
            tdds.append(x)
    return (max(thds) if thds else None), (max(tdds) if tdds else None)


def _compose_remarks_mba_intro(
    *,
    name: str,
    load_pct: float | None,
    wave: str,
    du_rhetorical: str,
    pf_txt: str,
    quality: str,
) -> str:
    load_seg = ""
    if load_pct is not None and load_pct == load_pct:
        load_seg = (
            f"Công suất tiêu thụ của {name} đạt {_fmt_remark_pct(load_pct, 2)}% công suất thiết kế. "
        )
    return (
        f"Nhận xét: {load_seg}"
        f"Biểu đồ dòng điện tiêu thụ tại thời điểm đo kiểm {wave}, "
        f"độ lệch pha điện áp ở mức {du_rhetorical}, "
        f"hệ số công suất cosφ ở mức {pf_txt}. "
        f"Chất lượng điện đo tại {name} ở mức {quality}. "
        f"Dưới đây là bảng tổng hợp thông số hoạt động của {name}:"
    )


def _compose_remarks_device_paragraph(
    *,
    name: str,
    vref: float,
    tdd_limit_pct: float,
    u_line_min: float | None,
    u_line_max: float | None,
    du_line_low: float | None,
    du_line_high: float | None,
    delta_u_ok: bool,
    uv_unb_max: float | None,
    ua_max: float | None,
    di_ok: bool,
    pf_avg: float | None,
    thd_max: float | None,
    tdd_max: float | None,
    spread_pct: float | None,
) -> str:
    """Một đoạn liên tục 6 ý — theo quy-tac-2.md phần thiết bị."""
    wave = _waveform_phrase_device(spread_pct)
    q = _quality_level_device(
        delta_u_ok=delta_u_ok,
        thd_ok=thd_max is not None and thd_max < _THDV_LIMIT_PCT,
        tdd_ok=tdd_max is not None and tdd_max < tdd_limit_pct,
        di_ok=di_ok,
        pf_avg=pf_avg,
    )
    pf_txt = _pf_text_for_remarks(pf_avg)
    umin = _fmt_remark_voltage(u_line_min, 1)
    umax = _fmt_remark_voltage(u_line_max, 1)
    d1 = _fmt_remark_pct(du_line_low, 2)
    d2 = _fmt_remark_pct(du_line_high, 2)
    du_s = _fmt_remark_pct(uv_unb_max, 2) if uv_unb_max is not None else "—"
    di_s = _fmt_remark_pct(ua_max, 2) if ua_max is not None else "—"

    if uv_unb_max is None or ua_max is None:
        return (
            f"Nhận xét: Dữ liệu INPS tại {name} thiếu độ lệch pha (UV%/UA%); "
            "không thể tự động lập nhận xét đầy đủ theo quy chuẩn."
        )

    du_num = float(uv_unb_max) if uv_unb_max is not None else 0.0
    di_num = float(ua_max) if ua_max is not None else 0.0
    du_pass = du_num < _V_DEV_LIMIT_PCT
    di_pass = di_ok
    if du_pass and di_pass:
        di_cmp = "<"
        di_part = f"đều ở mức thấp (ΔU = {du_s}% < 5,0%, ΔI = {di_s}% {di_cmp} 10,0%)."
    elif du_pass and not di_pass:
        di_cmp = ">" if di_num > 10.0 else ">"
        di_part = (
            f"điện áp ở mức thấp (ΔU = {du_s}% < 5,0%); tuy nhiên, "
            f"độ lệch pha dòng điện ở mức cao (ΔI = {di_s}% {di_cmp} 10,0%)."
        )
    elif not du_pass and di_pass:
        di_part = (
            f"độ lệch pha dòng điện ở mức thấp (ΔI = {di_s}% < 10,0%); tuy nhiên, "
            f"độ lệch pha điện áp vượt mức cho phép (ΔU = {du_s}% > 5,0%)."
        )
    else:
        di_cmp = ">" if di_num > 10.0 else ">"
        di_part = (
            f"điện áp và dòng điện đều cần chú ý (ΔU = {du_s}% > 5,0%, ΔI = {di_s}% {di_cmp} 10,0%)."
        )

    thd_ok = thd_max is not None and thd_max < _THDV_LIMIT_PCT
    tdd_ok = tdd_max is not None and tdd_max < tdd_limit_pct
    th_s = _fmt_remark_pct(thd_max, 2) if thd_max is not None else "—"
    td_s = _fmt_remark_pct(tdd_max, 2) if tdd_max is not None else "—"
    lim_s = _fmt_remark_pct(tdd_limit_pct, 1)

    if thd_ok and tdd_ok:
        harm = (
            f"Tổng biến dạng sóng hài điện áp và dòng điện đều ở mức cho phép "
            f"(THDmax = {th_s}% < 8,0% & TDDmax = {td_s}% < {lim_s}%)."
        )
    elif thd_ok and not tdd_ok:
        harm = (
            f"Tổng biến dạng sóng hài điện áp ở mức cho phép (THDmax = {th_s}% < 8,0%); "
            f"tuy nhiên, tổng biến dạng sóng hài dòng điện cao hơn mức cho phép "
            f"(TDDmax = {td_s}% > {lim_s}%)."
        )
    elif not thd_ok and tdd_ok:
        harm = (
            f"Tổng biến dạng sóng hài dòng điện ở mức cho phép (TDDmax = {td_s}% < {lim_s}%); "
            f"tuy nhiên, tổng biến dạng sóng hài điện áp cao hơn mức cho phép "
            f"(THDmax = {th_s}% > 8,0%)."
        )
    else:
        harm = (
            f"Tổng biến dạng sóng hài điện áp và dòng điện đều cao hơn mức cho phép "
            f"(THDmax = {th_s}% > 8,0% & TDDmax = {td_s}% > {lim_s}%)."
        )

    return (
        f"Nhận xét: Chất lượng điện cấp cho {name} ở mức {q}. "
        f"Biểu đồ dòng điện tiêu thụ tại {name} {wave} trong thời gian đo kiểm. "
        f"Hệ số công suất cosφ ở mức {pf_txt}. "
        f"Điện áp dao động từ {umin} ÷ {umax} V, độ lệch chuẩn của điện áp δU "
        f"(= {d1}% ÷ {d2}%) đạt tiêu chuẩn (-5,0% ≤ δ ≤ 5,0%). "
        f"Độ lệch pha điện áp và dòng điện {di_part} "
        f"{harm}"
    )


def _device_tdd_limit_from_name(name: str) -> float:
    n = _norm(name)
    if any(k in n for k in ("nén", "nen", "nghiền", "nghien", "máy ép", "may ep", "băng tải", "bang tai")):
        return 12.0
    return 20.0


def _build_auto_remarks_from_inps(
    stats: Mapping[str, dict],
    *,
    name: str,
    kind: SectionKind,
    folder: str | Path | None,
    nominal_voltage: float | None,
) -> str:
    """Sinh đoạn «Nhận xét:» từ cùng nguồn INPS với bả Word (quy-tac.md / quy-tac-2.md)."""
    if not stats:
        return ""

    u12 = _pick(stats, "AVG_VL1[V]", "AVG_V12[V]")
    u23 = _pick(stats, "AVG_VL2[V]", "AVG_V23[V]")
    u31 = _pick(stats, "AVG_VL3[V]", "AVG_V31[V]")
    if not u12:
        v1 = _pick(stats, "AVG_V1[V]")
        v2 = _pick(stats, "AVG_V2[V]")
        v3 = _pick(stats, "AVG_V3[V]")
        k = 3 ** 0.5
        u12, u23, u31 = _scale(v1, k), _scale(v2, k), _scale(v3, k)

    vref = float(nominal_voltage) if nominal_voltage and nominal_voltage > 0 else _MBA_NOMINAL_VOLTAGE_V
    u12eval, _, _, _ = _eval_voltage(u12.get("max"), u12.get("min"), vref)

    i1 = _pick(stats, "AVG_A1[A]")
    i2 = _pick(stats, "AVG_A2[A]")
    i3 = _pick(stats, "AVG_A3[A]")
    uv_unb = _pick(stats, "AVG_UV[%]", "AVG_VUNB[%]")
    ua_unb = _pick(stats, "AVG_UA[%]", "AVG_AUNB[%]")
    pf = _pick_total(stats, "PF", "[_]") or _pick(stats, "AVG_PF[_]")
    thd1 = _pick(stats, "AVG_THDVR1[%]", "AVG_VTHD1[%]")
    thd2 = _pick(stats, "AVG_THDVR2[%]", "AVG_VTHD2[%]")
    thd3 = _pick(stats, "AVG_THDVR3[%]", "AVG_VTHD3[%]")
    tdd1 = _pick(stats, "AVG_THDAR1[%]", "AVG_ATHD1[%]")
    tdd2 = _pick(stats, "AVG_THDAR2[%]", "AVG_ATHD2[%]")
    tdd3 = _pick(stats, "AVG_THDAR3[%]", "AVG_ATHD3[%]")

    dueval = "Đạt" if (uv_unb.get("max") is not None and uv_unb["max"] < _V_DEV_LIMIT_PCT) else (
        "Không đạt" if uv_unb.get("max") is not None else "—"
    )
    pfeval = _eval_pf(pf.get("avg"))
    thdeval = _eval_thd(
        [thd1.get("max"), thd2.get("max"), thd3.get("max")], _THDV_LIMIT_PCT
    )
    tdd_lim = _TDD_LIMIT_PCT if kind == "mba" else _device_tdd_limit_from_name(name)
    tddeval = _eval_thd(
        [tdd1.get("max"), tdd2.get("max"), tdd3.get("max")], tdd_lim
    )

    u_line_min, u_line_max = _line_voltage_range(u12, u23, u31)
    du_lo, du_hi, delta_u_ok = _delta_u_line_window_pct(u_line_min, u_line_max, vref)
    uv_max = uv_unb.get("max")
    try:
        uv_max_f = float(uv_max) if uv_max is not None else None
    except (TypeError, ValueError):
        uv_max_f = None
    ua_max = ua_unb.get("max")
    try:
        ua_max_f = float(ua_max) if ua_max is not None else None
    except (TypeError, ValueError):
        ua_max_f = None
    di_ok = ua_max_f is not None and ua_max_f < 10.0

    pf_avg = pf.get("avg")
    try:
        pf_avg_f = float(pf_avg) if pf_avg is not None else None
    except (TypeError, ValueError):
        pf_avg_f = None

    thd_max, tdd_max = _thd_tdd_maxes(thd1, thd2, thd3, tdd1, tdd2, tdd3)
    spread = _current_spread_pct(i1, i2, i3)

    if kind == "mba":
        s_total = _pick_total(stats, "S", "[VA]") or _pick(stats, "AVG_S[VA]")
        s_avg_va = s_total.get("avg") if s_total else None
        try:
            s_avg_kva = float(s_avg_va) / 1000.0 if s_avg_va is not None else None
        except (TypeError, ValueError):
            s_avg_kva = None
        rated = _rated_kva_from_inis_folder(folder)
        load_pct = None
        if s_avg_kva is not None and rated is not None and rated > 0:
            load_pct = s_avg_kva / rated * 100.0

        q = _quality_level_mba(
            u12eval=u12eval, dueval=dueval, pfeval=pfeval, thdeval=thdeval, tddeval=tddeval,
        )
        return _compose_remarks_mba_intro(
            name=name,
            load_pct=load_pct,
            wave=_waveform_phrase_mba(spread),
            du_rhetorical=_du_rhetorical_mba(uv_max_f),
            pf_txt=_pf_text_for_remarks(pf_avg_f),
            quality=q,
        )

    return _compose_remarks_device_paragraph(
        name=name,
        vref=vref,
        tdd_limit_pct=tdd_lim,
        u_line_min=u_line_min,
        u_line_max=u_line_max,
        du_line_low=du_lo,
        du_line_high=du_hi,
        delta_u_ok=delta_u_ok,
        uv_unb_max=uv_max_f,
        ua_max=ua_max_f,
        di_ok=di_ok,
        pf_avg=pf_avg_f,
        thd_max=thd_max,
        tdd_max=tdd_max,
        spread_pct=spread,
    )


# ──────────────────────────────────────────────────────────────────────────────
# Sinh nhận xét từ các trường Excel hiện trường (KHÔNG dùng INPS)
# ──────────────────────────────────────────────────────────────────────────────

_CURRENT_CHAR_MAP: dict[str, str] = {
    "on dinh": "ổn định",
    "on định": "ổn định",
    "ổn định": "ổn định",
    "dao dong nhe": "có sự dao động nhẹ",
    "dao động nhẹ": "có sự dao động nhẹ",
    "bien doi lien tuc": "biến đổi liên tục với biên độ nhỏ",
    "biến đổi liên tục": "biến đổi liên tục với biên độ nhỏ",
    "chu ky load-unload": "biến đổi theo chu kỳ Load/Unload",
    "chu kỳ load-unload": "biến đổi theo chu kỳ Load/Unload",
    "load-unload": "biến đổi theo chu kỳ Load/Unload",
    "load/unload": "biến đổi theo chu kỳ Load/Unload",
}


def _parse_float_field(v: object) -> float | None:
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


def _wave_phrase_from_char(current_char: str | None) -> str:
    """Chuyển giá trị cột current_char → cụm từ tiêu chuẩn."""
    if not current_char:
        return "biến đổi liên tục với biên độ nhỏ"
    key = unicodedata.normalize("NFKC", str(current_char).strip()).lower()
    for k, v in _CURRENT_CHAR_MAP.items():
        if k in key:
            return v
    return str(current_char).strip()


def _pf_phrase(cos_phi: float | None) -> str:
    if cos_phi is None or cos_phi != cos_phi:
        return "chưa xác định"
    p = abs(cos_phi)
    if p >= 0.995:
        return "rất cao (cosφ ≈ 1, có thể đã lắp đặt tụ bù)"
    if p >= 0.8:
        return "cao (trên 0,8)"
    if p >= 0.5:
        return "trung bình (dưới 0,8)"
    return "thấp (dưới 0,8)"


def _compose_remarks_from_excel_fields(
    *,
    name: str,
    kind: SectionKind,
    current_char: str | None,
    u_min: float | None,
    u_max: float | None,
    i_max: float | None = None,
    delta_u: float | None,
    delta_i: float | None,
    p_kw: float | None = None,
    cos_phi: float | None,
    thd_max: float | None,
    tdd_max: float | None,
    pdm_kva: float | None = None,
    nominal_voltage: float | None,
) -> str:
    """Sinh đoạn nhận xét từ các trường Excel hiện trường.

    - Không có tiền tố "Nhận xét: " (template Word tự thêm nếu cần).
    - Thuật toán Loi_Dem: delta_u vi phạm cộng 2, các vi phạm khác cộng 1.
    - Device: 6 câu theo thứ tự spec.
    - MBA: 3 câu (% tải, biểu đồ dòng, đánh giá tổng + dẫn bảng).
    """
    vref = float(nominal_voltage) if nominal_voltage and nominal_voltage > 0 else _MBA_NOMINAL_VOLTAGE_V
    wave = _wave_phrase_from_char(current_char)

    # ── Định dạng helper ─────────────────────────────────────────────────
    def _pct(v: float | None, d: int = 2) -> str:
        return "—" if v is None else f"{v:.{d}f}".replace(".", ",")

    def _volt(v: float | None, d: int = 1) -> str:
        return "—" if v is None else f"{v:.{d}f}".replace(".", ",")

    # ── δU lệch so với điện áp danh định (từ u_min / u_max) ─────────────
    du_lo: float | None = None
    du_hi: float | None = None
    if u_min is not None and u_max is not None and vref > 0:
        du_lo = (u_min - vref) / vref * 100.0
        du_hi = (u_max - vref) / vref * 100.0

    # ── Thuật toán "Cờ báo lỗi" (Loi_Dem) ───────────────────────────────
    tdd_lim = _TDD_LIMIT_PCT if kind == "mba" else _device_tdd_limit_from_name(name)
    loi_dem = 0

    if cos_phi is not None and abs(cos_phi) < _PF_LIMIT:
        loi_dem += 1
    if du_lo is not None and du_hi is not None:
        if du_lo < -_V_DEV_LIMIT_PCT or du_hi > _V_DEV_LIMIT_PCT:
            loi_dem += 1
    if delta_i is not None and delta_i > 10.0:
        loi_dem += 1
    if thd_max is not None and thd_max > _THDV_LIMIT_PCT:
        loi_dem += 1
    if tdd_max is not None and tdd_max > tdd_lim:
        loi_dem += 1
    if delta_u is not None and delta_u > _V_DEV_LIMIT_PCT:
        loi_dem += 2  # Mất cân bằng điện áp — lỗi nghiêm trọng, cộng 2

    # ── Đánh giá chất lượng tổng quan ────────────────────────────────────
    if kind == "mba":
        quality = "Tốt" if loi_dem == 0 else ("Tương đối tốt" if loi_dem == 1 else "Chưa thực sự tốt")
    else:
        quality = "Tốt" if loi_dem == 0 else ("Tương đối tốt" if loi_dem == 1 else "Chưa tốt")

    # ── Câu Hệ số công suất ───────────────────────────────────────────────
    pf_txt = _pf_phrase(cos_phi)

    # ── Câu Độ lệch pha ΔU / ΔI (Mẫu 1) ─────────────────────────────────
    du_num = float(delta_u) if delta_u is not None else None
    di_num = float(delta_i) if delta_i is not None else None
    du_pass = du_num is not None and du_num <= _V_DEV_LIMIT_PCT
    di_pass = di_num is not None and di_num <= 10.0
    du_s, di_s = _pct(du_num), _pct(di_num)

    if du_num is None and di_num is None:
        unbalance_sent = ""
    elif du_num is None:
        unbalance_sent = (
            f"Độ lệch pha dòng điện ở mức {'thấp' if di_pass else 'cao'} "
            f"(ΔI = {di_s}% {'<' if di_pass else '>'} 10,0%)."
        )
    elif di_num is None:
        unbalance_sent = (
            f"Độ lệch pha điện áp ở mức {'thấp' if du_pass else 'cao'} "
            f"(ΔU = {du_s}% {'<' if du_pass else '>'} 5,0%)."
        )
    elif du_pass and di_pass:
        unbalance_sent = (
            f"Độ lệch pha điện áp và dòng điện đều ở mức thấp "
            f"(ΔU = {du_s}% < 5,0%; ΔI = {di_s}% < 10,0%)."
        )
    elif du_pass and not di_pass:
        unbalance_sent = (
            f"Độ lệch pha điện áp ở mức thấp (ΔU = {du_s}% < 5,0%); "
            f"tuy nhiên, độ lệch pha dòng điện vượt mức cho phép (ΔI = {di_s}% > 10,0%)."
        )
    elif not du_pass and di_pass:
        unbalance_sent = (
            f"Độ lệch dòng điện ở mức thấp (ΔI = {di_s}% < 10,0%); "
            f"tuy nhiên, độ lệch pha điện áp vượt mức cho phép (ΔU = {du_s}% > 5,0%)."
        )
    else:
        unbalance_sent = (
            f"Độ lệch pha điện áp và dòng điện đều vượt mức cho phép "
            f"(ΔU = {du_s}% > 5,0%; ΔI = {di_s}% > 10,0%)."
        )

    # ── Câu Sóng hài THD / TDD (Mẫu 2) ──────────────────────────────────
    lim_s = _pct(tdd_lim, 1)
    th_s, td_s = _pct(thd_max), _pct(tdd_max)
    thd_ok = thd_max is not None and thd_max <= _THDV_LIMIT_PCT
    tdd_ok = tdd_max is not None and tdd_max <= tdd_lim

    if thd_max is None and tdd_max is None:
        harm_sent = ""
    elif thd_ok and tdd_ok:
        harm_sent = (
            f"Tổng biến dạng sóng hài điện áp và dòng điện đều ở mức cho phép "
            f"(THDmax = {th_s}% < 8,0%, TDDmax = {td_s}% < {lim_s}%)."
        )
    elif thd_ok and not tdd_ok:
        harm_sent = (
            f"Tổng biến dạng sóng hài điện áp ở mức cho phép (THDmax = {th_s}% < 8,0%); "
            f"tuy nhiên, tổng biến dạng sóng hài dòng điện vượt mức cho phép "
            f"(TDDmax = {td_s}% > {lim_s}%)."
        )
    elif not thd_ok and tdd_ok:
        harm_sent = (
            f"Tổng biến dạng sóng hài dòng điện ở mức cho phép (TDDmax = {td_s}% < {lim_s}%); "
            f"tuy nhiên, tổng biến dạng sóng hài điện áp vượt mức cho phép "
            f"(THDmax = {th_s}% > 8,0%)."
        )
    else:
        harm_sent = (
            f"Tổng biến dạng sóng hài điện áp và tổng biến dạng sóng hài dòng điện "
            f"đều vượt mức cho phép (THDmax = {th_s}% > 8,0%, TDDmax = {td_s}% > {lim_s}%)."
        )

    # ── MBA: 3 câu (% tải → biểu đồ dòng → đánh giá + dẫn bảng) ────────
    if kind == "mba":
        load_seg = ""
        if p_kw is not None and cos_phi is not None and abs(cos_phi) > 0.01 and pdm_kva is not None and pdm_kva > 0:
            s_kva = p_kw / abs(cos_phi)
            load_pct = s_kva / pdm_kva * 100.0
            load_seg = f"Công suất tiêu thụ của {name} đạt {_pct(load_pct, 2)}% công suất thiết kế. "
        return (
            f"{load_seg}"
            f"Biểu đồ dòng điện tiêu thụ tại thời điểm đo kiểm {wave}. "
            f"Chất lượng điện đo tại {name} ở mức {quality}. "
            f"Dưới đây là bảng tổng hợp thông số hoạt động của {name}:"
        )

    # ── Device: 6 câu theo thứ tự spec ───────────────────────────────────
    # Câu 4: Điện áp dao động + δU so danh định
    umin_s, umax_s = _volt(u_min), _volt(u_max)
    dlo_s, dhi_s = _pct(du_lo), _pct(du_hi)
    if du_lo is not None and du_hi is not None:
        both_in = abs(du_lo) <= _V_DEV_LIMIT_PCT and abs(du_hi) <= _V_DEV_LIMIT_PCT
        verdict = "đạt tiêu chuẩn (-5,0% ≤ δ ≤ 5,0%)" if both_in else "vượt giới hạn cho phép (-5,0% ≤ δ ≤ 5,0%)"
        volt_sent = (
            f"Điện áp dao động từ {umin_s} ÷ {umax_s} V, "
            f"độ lệch so với điện áp danh định δU = {dlo_s}% ÷ {dhi_s}%, {verdict}."
        )
    elif u_min is not None and u_max is not None:
        volt_sent = f"Điện áp dao động từ {umin_s} ÷ {umax_s} V."
    else:
        volt_sent = ""

    parts: list[str] = [
        f"Chất lượng điện cấp cho {name} ở mức {quality}.",
        f"Biểu đồ dòng điện tiêu thụ tại {name} {wave} trong thời gian đo kiểm.",
        f"Hệ số công suất cosφ ở mức {pf_txt}.",
    ]
    if volt_sent:
        parts.append(volt_sent)
    if unbalance_sent:
        parts.append(unbalance_sent)
    if harm_sent:
        parts.append(harm_sent)
    return " ".join(parts)


def _resolve_remarks_field(
    *,
    kind: SectionKind,
    folder: Path,
    name: str,
    user_remarks: str,
    nominal_voltage: float | None,
    excel_params: dict | None = None,
) -> str:
    """Sinh nhận xét từ các trường Excel hiện trường. Nếu người dùng ghi đoạn
    thủ công đầy đủ (bắt đầu bằng «Nhận xét:»), dùng nguyên văn.
    INPS không còn được tham chiếu ở đây nữa.
    """
    raw = (user_remarks or "").strip()
    if _is_full_manual_remarks(raw):
        return raw

    params = excel_params or {}
    current_char = params.get("current_char")
    u_min = _parse_float_field(params.get("u_min"))
    u_max = _parse_float_field(params.get("u_max"))
    i_max = _parse_float_field(params.get("i_max"))
    delta_u = _parse_float_field(params.get("delta_u"))
    delta_i = _parse_float_field(params.get("delta_i"))
    p_kw = _parse_float_field(params.get("p"))
    cos_phi = _parse_float_field(params.get("cos_phi"))
    thd_max = _parse_float_field(params.get("thd"))
    tdd_max_v = _parse_float_field(params.get("tdd"))
    pdm_kva = _parse_float_field(params.get("pdm"))

    has_data = any(v is not None for v in [
        current_char, u_min, u_max, delta_u, delta_i, cos_phi, thd_max, tdd_max_v
    ])
    if not has_data:
        return raw

    auto = _compose_remarks_from_excel_fields(
        name=name,
        kind=kind,
        current_char=current_char,
        u_min=u_min,
        u_max=u_max,
        i_max=i_max,
        delta_u=delta_u,
        delta_i=delta_i,
        p_kw=p_kw,
        cos_phi=cos_phi,
        thd_max=thd_max,
        tdd_max=tdd_max_v,
        pdm_kva=pdm_kva,
        nominal_voltage=nominal_voltage,
    )
    return _merge_auto_and_excel_notes(auto, raw)


# ════════════════════════════════════════════════════════════════════
#                       Image discovery helpers
# ════════════════════════════════════════════════════════════════════


def list_bmp_in_folder(folder: str | Path) -> list[Path]:
    """Trả về danh sách ``PS-SDxxx.BMP`` trong ``folder``, sắp theo số trong tên (641→642→…).

    Không gồm ``a.png`` — ảnh tổng quan dùng :func:`find_overview_png`.
    """
    p = Path(folder)
    if not p.is_dir():
        return []
    bmps: list[tuple[int, Path]] = []
    for f in p.iterdir():
        if not f.is_file():
            continue
        m = _BMP_RE.match(f.name)
        if m:
            bmps.append((int(m.group(1)), f))
    bmps.sort(key=lambda x: x[0])
    # Fallback: mọi file .bmp nếu không có tên PS-SD (hồ sơ cũ / tay)
    if not bmps:
        extra: list[tuple[int, Path]] = []
        for i, f in enumerate(sorted(p.glob("*.BMP")) + sorted(p.glob("*.bmp"))):
            if not f.is_file():
                continue
            extra.append((i, f))
        bmps = extra
    return [pp for _, pp in bmps]


def find_overview_png(folder: str | Path) -> Path | None:
    """Ảnh tổng quan ``a.png`` / ``A.png`` (stem ``a``) trong thư mục thiết bị, nếu có."""
    p = Path(folder)
    if not p.is_dir():
        return None
    for f in p.iterdir():
        if not f.is_file():
            continue
        if f.suffix.lower() == ".png" and f.stem.lower() == "a":
            return f
    return None


def _require_overview_png_and_bmps(folder: str | Path) -> tuple[Path, list[Path]]:
    """Trả về ``(a.png, danh_sách_PS-SD_BMP)`` hoặc ném lỗi nếu thiếu."""
    p = Path(folder)
    overview = find_overview_png(p)
    if overview is None:
        raise FileNotFoundError(
            f"Thiếu file a.png (ảnh tổng quan) trong {p!s}. "
            "Vui lòng đặt a.png trong thư mục thiết bị."
        )
    bmps = list_bmp_in_folder(p)
    if not bmps:
        raise FileNotFoundError(f"Không tìm thấy PS-SD*.BMP trong {p!s}.")
    return overview, bmps


def _take(lst: list[Path], idx: int, fallback: Path | None) -> Path | None:
    return lst[idx] if 0 <= idx < len(lst) else fallback


def auto_pick_mba_images(folder: str | Path) -> dict[str, str]:
    """Chọn ảnh cho template MBA (``imga`` + ``img1, img2, img4, img6``).

    * ``imga``: **bắt buộc** file ``a.png`` (stem ``a``, đuôi ``.png``) trong thư mục.
    * Các ô nhỏ: ``PS-SD*.BMP`` theo thứ tự số → ``img1``, ``img2``, ``img4``, ``img6``.
    """
    overview, bmps = _require_overview_png_and_bmps(folder)
    fb = bmps[0]
    imga = str(overview)
    return {
        "imga": imga,
        "img1": str(_take(bmps, 0, fb)),
        "img2": str(_take(bmps, 1, fb)),
        "img4": str(_take(bmps, 2, fb)),
        "img6": str(_take(bmps, 3, fb)),
    }


def auto_pick_device_images(folder: str | Path) -> dict[str, str]:
    """Chọn ảnh cho template device (``imga`` + ``img1``…``img6``).

    * ``imga``: **bắt buộc** file ``a.png`` trong thư mục.
    * ``img1``…``img6``: ``PS-SD*.BMP`` theo thứ tự số (vd 641–646 → ``img1``–``img6``).
    """
    overview, bmps = _require_overview_png_and_bmps(folder)
    fb = bmps[0]
    imga = str(overview)
    return {
        "imga": imga,
        **{f"img{i}": str(_take(bmps, i - 1, fb)) for i in range(1, 7)},
    }


# ════════════════════════════════════════════════════════════════════
#                  Kwargs builders từ thư mục thiết bị
# ════════════════════════════════════════════════════════════════════


def mba_kwargs_from_inps(
    inps_path: str | Path | None,
    *,
    name: str,
    imga: str,
    img1: str,
    img2: str,
    img4: str,
    img6: str,
    cap_fig_mba: str | None = None,
    remarks_mba: str = "",
    cap_tab_mba: str | None = None,
) -> dict:
    """Đọc INPS (nếu có), tổng hợp số liệu, trả về ``kwargs`` cho :func:`mba`.

    Khi ``inps_path`` là ``None`` hoặc không đọc được file, bảng số liệu dùng
    ``"—"`` (vẫn render template MBA — ví dụ thiếu INPS nhưng Excel ghi loại MBA).

    Đánh giá lệch điện áp U12: cố định so với **400 V** (mỗi lệch % tại Umax/Umin nằm trong ±5%),
    không dùng cột điện áp định mức từ Excel.

    ``imga, img1, img2, img4, img6`` là đường dẫn ảnh (đã tự chọn từ thư mục
    hoặc do người dùng chỉ định).
    """
    if inps_path is None:
        stats: dict[str, dict] = {}
    else:
        stats = _parse_inps(inps_path)

    # ─── Điện áp dây U12, U23, U31 (ưu tiên AVG_VLi[V]) ────────────
    u12 = _pick(stats, "AVG_VL1[V]", "AVG_V12[V]")
    u23 = _pick(stats, "AVG_VL2[V]", "AVG_V23[V]")
    u31 = _pick(stats, "AVG_VL3[V]", "AVG_V31[V]")

    # Nếu thiết bị chỉ ghi điện áp pha → quy đổi sang dây bằng √3.
    if not u12:
        v1 = _pick(stats, "AVG_V1[V]")
        v2 = _pick(stats, "AVG_V2[V]")
        v3 = _pick(stats, "AVG_V3[V]")
        k = 3 ** 0.5
        u12, u23, u31 = _scale(v1, k), _scale(v2, k), _scale(v3, k)

    vref = _MBA_NOMINAL_VOLTAGE_V
    u12eval, _, _, _ = _eval_voltage(u12.get("max"), u12.get("min"), vref)

    # ─── Dòng điện I1, I2, I3 ─────────────────────────────────────
    i1 = _pick(stats, "AVG_A1[A]")
    i2 = _pick(stats, "AVG_A2[A]")
    i3 = _pick(stats, "AVG_A3[A]")

    # ─── Độ lệch pha điện áp / dòng (%) ───────────────────────────
    uv_unb = _pick(stats, "AVG_UV[%]", "AVG_VUNB[%]")
    ua_unb = _pick(stats, "AVG_UA[%]", "AVG_AUNB[%]")
    dueval = "Đạt" if (uv_unb.get("max") is not None and uv_unb["max"] < _V_DEV_LIMIT_PCT) else (
        "Không đạt" if uv_unb.get("max") is not None else "—"
    )

    # ─── Hệ số công suất ─────────────────────────────────────────
    pf = _pick_total(stats, "PF", "[_]") or _pick(stats, "AVG_PF[_]")
    pfeval = _eval_pf(pf.get("avg"))

    # ─── P, Q, S (đổi sang kW/kvar/kVA) ──────────────────────────
    p_total = _pick_total(stats, "P", "[W]") or _pick(stats, "AVG_P[W]")
    q_total = _pick_total(stats, "Q", "[var]") or _pick(stats, "AVG_Q[var]")
    s_total = _pick_total(stats, "S", "[VA]") or _pick(stats, "AVG_S[VA]")
    p_k = _scale(p_total, 1e-3)
    q_k = _scale(q_total, 1e-3)
    s_k = _scale(s_total, 1e-3)

    # ─── THD V & TDD I theo pha ─────────────────────────────────
    thd1 = _pick(stats, "AVG_THDVR1[%]", "AVG_VTHD1[%]")
    thd2 = _pick(stats, "AVG_THDVR2[%]", "AVG_VTHD2[%]")
    thd3 = _pick(stats, "AVG_THDVR3[%]", "AVG_VTHD3[%]")
    tdd1 = _pick(stats, "AVG_THDAR1[%]", "AVG_ATHD1[%]")
    tdd2 = _pick(stats, "AVG_THDAR2[%]", "AVG_ATHD2[%]")
    tdd3 = _pick(stats, "AVG_THDAR3[%]", "AVG_ATHD3[%]")

    thdeval = _eval_thd(
        [thd1.get("max"), thd2.get("max"), thd3.get("max")], _THDV_LIMIT_PCT
    )
    tddeval = _eval_thd(
        [tdd1.get("max"), tdd2.get("max"), tdd3.get("max")], _TDD_LIMIT_PCT
    )

    u12max, u12min, u12avg = _tri(u12, 1)
    u23max, u23min, u23avg = _tri(u23, 1)
    u31max, u31min, u31avg = _tri(u31, 1)
    i1max, i1min, i1avg = _tri(i1, 0)
    i2max, i2min, i2avg = _tri(i2, 0)
    i3max, i3min, i3avg = _tri(i3, 0)
    dumax, dumin, duavg = _tri(uv_unb, 3)
    dimax, dimin, diavg = _tri(ua_unb, 3)
    pfmax, pfmin, pfavg = _tri(pf, 3)
    pmax, pmin, pavg = _tri(p_k, 1)
    qmax, qmin, qavg = _tri(q_k, 1)
    smax, smin, savg = _tri(s_k, 1)
    thd1max, thd1min, thd1avg = _tri(thd1, 2)
    thd2max, thd2min, thd2avg = _tri(thd2, 2)
    thd3max, thd3min, thd3avg = _tri(thd3, 2)
    tdd1max, tdd1min, tdd1avg = _tri(tdd1, 2)
    tdd2max, tdd2min, tdd2avg = _tri(tdd2, 2)
    tdd3max, tdd3min, tdd3avg = _tri(tdd3, 2)

    return {
        "name": name,
        "imga": imga, "img1": img1, "img2": img2, "img4": img4, "img6": img6,
        "cap_fig_mba": cap_fig_mba if cap_fig_mba is not None else f"Hình ảnh đo tại {name}",
        "remarks_mba": remarks_mba,
        "cap_tab_mba": cap_tab_mba if cap_tab_mba is not None else f"Bảng tổng hợp thông số đo tại {name}",
        "u12max": u12max, "u12min": u12min, "u12avg": u12avg, "u12eval": u12eval,
        "u23max": u23max, "u23min": u23min, "u23avg": u23avg,
        "u31max": u31max, "u31min": u31min, "u31avg": u31avg,
        "i1max": i1max, "i1min": i1min, "i1avg": i1avg,
        "i2max": i2max, "i2min": i2min, "i2avg": i2avg,
        "i3max": i3max, "i3min": i3min, "i3avg": i3avg,
        "dumax": dumax, "dumin": dumin, "duavg": duavg, "dueval": dueval,
        "dimax": dimax, "dimin": dimin, "diavg": diavg,
        "pfmax": pfmax, "pfmin": pfmin, "pfavg": pfavg, "pfeval": pfeval,
        "pmax": pmax, "pmin": pmin, "pavg": pavg,
        "qmax": qmax, "qmin": qmin, "qavg": qavg,
        "smax": smax, "smin": smin, "savg": savg,
        "thd1max": thd1max, "thd1min": thd1min, "thd1avg": thd1avg, "thdeval": thdeval,
        "thd2max": thd2max, "thd2min": thd2min, "thd2avg": thd2avg,
        "thd3max": thd3max, "thd3min": thd3min, "thd3avg": thd3avg,
        "tdd1max": tdd1max, "tdd1min": tdd1min, "tdd1avg": tdd1avg, "tddeval": tddeval,
        "tdd2max": tdd2max, "tdd2min": tdd2min, "tdd2avg": tdd2avg,
        "tdd3max": tdd3max, "tdd3min": tdd3min, "tdd3avg": tdd3avg,
    }


def mba_kwargs_from_folder(
    folder: str | Path,
    *,
    name: str,
    cap_fig_mba: str | None = None,
    remarks_mba: str = "",
    cap_tab_mba: str | None = None,
    excel_params: dict | None = None,
) -> dict:
    """Tự chọn ảnh trong ``folder``, tìm ``INPS*.KEW`` rồi dựng kwargs MBA.

    Nhận xét sinh từ các trường ``excel_params`` (không dùng INPS).
    """
    folder = Path(folder)
    from modules.kew.analyse_kew import find_file  # type: ignore

    inps_path = find_file(str(folder), "INPS")
    images = auto_pick_mba_images(folder)
    kw = mba_kwargs_from_inps(
        inps_path,
        name=name,
        cap_fig_mba=cap_fig_mba,
        remarks_mba="",
        cap_tab_mba=cap_tab_mba,
        **images,
    )
    kw["remarks_mba"] = _resolve_remarks_field(
        kind="mba",
        folder=folder,
        name=name,
        user_remarks=remarks_mba,
        nominal_voltage=None,
        excel_params=excel_params,
    )
    return kw


def device_kwargs_from_folder(
    folder: str | Path,
    *,
    name: str,
    cap_device: str | None = None,
    remarks_device: str = "",
    nominal_voltage: float | None = None,
    excel_params: dict | None = None,
) -> dict:
    """Trả về ``kwargs`` cho :func:`device` (ảnh + caption + nhận xét từ các trường Excel hiện trường)."""
    images = auto_pick_device_images(folder)
    return {
        "name": name,
        "cap_device": cap_device if cap_device is not None else f"Hình ảnh đo tại {name}",
        "remarks_device": _resolve_remarks_field(
            kind="device",
            folder=Path(folder),
            name=name,
            user_remarks=remarks_device,
            nominal_voltage=nominal_voltage,
            excel_params=excel_params,
        ),
        **images,
    }


# ════════════════════════════════════════════════════════════════════
#                         Render & merge
# ════════════════════════════════════════════════════════════════════


def merge_rendered_docx(
    docx_paths: Sequence[str | Path],
    output_path: str | Path,
) -> Path:
    """Ghép các file Word đã render thành một file, giữ quan hệ media."""
    paths = [Path(p) for p in docx_paths]
    if not paths:
        raise ValueError("merge_rendered_docx: cần ít nhất một đường dẫn .docx.")
    out = Path(output_path)
    out.parent.mkdir(parents=True, exist_ok=True)
    if len(paths) == 1:
        copy2(paths[0], out)
        return out
    composer = Composer(Document(str(paths[0])))
    for p in paths[1:]:
        composer.append(Document(str(p)))
    composer.save(str(out))
    return out


def merge_mba_device_docx(
    output_path: str | Path,
    *,
    mba_template: str | Path,
    device_template: str | Path,
    mba_sections: list[dict] | None = None,
    device_sections: list[dict] | None = None,
    sections: Sequence[tuple[SectionKind, dict]] | None = None,
) -> Path:
    """Render từng MBA / device rồi ghép thành ``output_path``.

    * ``sections``: danh sách ``("mba", kwargs)`` / ``("device", kwargs)`` theo
      đúng thứ tự cần xuất hiện trong file cuối.
    * Nếu ``sections=None``: lần lượt toàn bộ ``mba_sections`` rồi ``device_sections``.
    """
    mba_sections = mba_sections or []
    device_sections = device_sections or []
    mba_tpl, device_tpl = Path(mba_template), Path(device_template)

    if sections is not None:
        work: list[tuple[SectionKind, dict]] = list(sections)
    else:
        work = [("mba", s) for s in mba_sections] + [("device", s) for s in device_sections]

    if not work:
        raise ValueError(
            "merge_mba_device_docx: cần sections hoặc mba_sections/device_sections không rỗng."
        )

    kind_render = {"mba": (mba_tpl, mba), "device": (device_tpl, device)}

    with TemporaryDirectory() as td:
        tmp = Path(td)
        rendered: list[Path] = []
        for i, (kind, spec) in enumerate(work):
            try:
                tpl_path, render_fn = kind_render[kind]
            except KeyError as e:
                raise ValueError(
                    f"Loại section không hợp lệ: {kind!r} (chỉ 'mba' hoặc 'device')."
                ) from e
            tpl = DocxTemplate(str(tpl_path))
            tpl.render(render_fn(tpl, **spec))
            path = tmp / f"{kind}_{i}.docx"
            tpl.save(str(path))
            if i < len(work) - 1:
                doc_pb = Document(str(path))
                doc_pb.add_paragraph().add_run().add_break(WD_BREAK.PAGE)
                doc_pb.save(str(path))
            rendered.append(path)
        return merge_rendered_docx(rendered, output_path)


# ════════════════════════════════════════════════════════════════════
#                Pipeline cho luồng "Xử lý file sơ bộ"
# ════════════════════════════════════════════════════════════════════


_MBA_NAME_RE = re.compile(r"^(MBA|TR|TBA|T\d|MBT)\b|MÁY BIẾN ÁP|BIẾN ÁP", re.IGNORECASE)


def _guess_kind(name: str) -> SectionKind:
    """Đoán loại template từ tên thiết bị (chỉ khi không có cột ``type`` từ Excel)."""
    return "mba" if _MBA_NAME_RE.search(name or "") else "device"


def _resolve_word_section_kind(
    spec: Mapping,
    *,
    name: str,
    default_kind: SectionKind | None,
) -> SectionKind:
    """Chọn ``mba`` / ``device`` cho báo cáo Word.

    * Nếu mục có khóa ``kind`` (luồng ZIP + Excel): **chỉ** ``mba`` khi cột type
      nhận diện được là MBA; ô trống / không nhận diện / ``device`` → ``device``
      (không đoán theo tên).
    * Nếu không có khóa ``kind`` (quét thư mục thuần): ``default_kind`` hoặc
      :func:`_guess_kind`.
    """
    if "kind" in spec:
        raw = spec["kind"]
        if raw is not None and not (isinstance(raw, str) and not str(raw).strip()):
            nk = _norm_kind(raw)
            if nk == "mba":
                return "mba"
            if nk == "device":
                return "device"
        return "device"
    if default_kind in ("mba", "device"):
        return default_kind
    return _guess_kind(name)


def build_field_word_report(
    project_root: str | Path,
    output_path: str | Path,
    *,
    mba_template: str | Path,
    device_template: str | Path,
    devices: Sequence[Mapping] | None = None,
    default_kind: SectionKind | None = None,
) -> tuple[Path, list[str]]:
    """Quét ``project_root`` (= ``Project_Output/``) → xuất 1 file Word tổng hợp.

    ``devices`` (tuỳ chọn): danh sách ``{name, folder, kind?, remarks?, nominal_voltage?}``.
        * ``folder`` có thể là tên thư mục con trong ``project_root`` hoặc đường dẫn tuyệt đối.
        * ``nominal_voltage`` (tuỳ chọn, từ Excel ``pdm``): điện áp chuẩn (V) để tính δU trong nhận xét thiết bị.
        * ``remarks``: ghi chú/ghi tay; nếu không phải đoạn ``Nhận xét:`` đầy đủ thì ghép sau bản tự động từ INPS.
        * Nếu mục có khóa ``kind`` (metadata Excel từ ZIP): chỉ ``mba`` khi cột type
          nhận diện MBA; không ghi / trống / không nhận diện → ``device`` (không đoán tên).
        * Nếu mục **không** có khóa ``kind``: dùng ``default_kind`` nếu có, không thì
          đoán theo tên (``MBA…`` → ``mba``).
    Khi ``devices=None``: tự duyệt mọi thư mục con của ``project_root`` (sort theo tên).

    Trả về ``(đường_dẫn_báo_cáo, warnings)``.
    """
    root = Path(project_root)
    if not root.is_dir():
        raise FileNotFoundError(f"Không tìm thấy thư mục Project_Output: {root}")

    if devices is None:
        devices = [
            {"name": _nfc(d.name), "folder": d}
            for d in sorted(root.iterdir())
            if d.is_dir() and not d.name.startswith(".") and d.name != "__MACOSX"
        ]
    if not devices:
        raise ValueError("Không có thiết bị nào để dựng báo cáo Word.")

    warnings: list[str] = []
    sections: list[tuple[SectionKind, dict]] = []
    for spec in devices:
        name = _nfc(str(spec.get("name") or "").strip())
        folder_raw = spec.get("folder")
        if not name or not folder_raw:
            warnings.append(f"Bỏ qua mục thiếu name/folder: {spec!r}.")
            continue
        folder = Path(folder_raw)
        if not folder.is_absolute():
            folder = root / folder
        if not folder.is_dir():
            warnings.append(f"«{name}»: không tìm thấy thư mục {folder}.")
            continue

        kind = _resolve_word_section_kind(spec, name=name, default_kind=default_kind)
        remarks = str(spec.get("remarks") or "")
        nom_raw = spec.get("nominal_voltage")
        nom_v: float | None = None
        if isinstance(nom_raw, (int, float)) and not isinstance(nom_raw, bool):
            if not (isinstance(nom_raw, float) and nom_raw != nom_raw):
                nom_v = float(nom_raw)
        excel_params = spec.get("excel_params") or {}

        try:
            if kind == "mba":
                kwargs = mba_kwargs_from_folder(
                    folder, name=name, remarks_mba=remarks,
                    excel_params=excel_params,
                )
            else:
                kwargs = device_kwargs_from_folder(
                    folder,
                    name=name,
                    remarks_device=remarks,
                    nominal_voltage=nom_v,
                    excel_params=excel_params,
                )
            sections.append((kind, kwargs))
        except FileNotFoundError as e:
            warnings.append(f"«{name}»: {e}.")
        except Exception as e:
            warnings.append(f"«{name}»: lỗi dựng section ({e}).")

    if not sections:
        raise RuntimeError("Không dựng được section nào: " + "; ".join(warnings))

    out = Path(output_path)
    merge_mba_device_docx(
        out,
        mba_template=mba_template,
        device_template=device_template,
        sections=sections,
    )
    return out, warnings


# ════════════════════════════════════════════════════════════════════
#   Đọc Excel metadata (Loại / U định mức / Nhận xét) cho tab Word
# ════════════════════════════════════════════════════════════════════


def _norm(s: object) -> str:
    if s is None:
        return ""
    t = unicodedata.normalize("NFKC", str(s)).strip().lower()
    return re.sub(r"\s+", " ", t)


def _nfc(s: object) -> str:
    """Chuẩn hoá tên hiển thị (tránh ký tự tổ hợp kiểu Mac / Excel lệch dạng)."""
    if s is None:
        return ""
    return unicodedata.normalize("NFC", str(s).strip())


def _norm_kind(value: object) -> SectionKind | None:
    raw = value
    if raw is None or (isinstance(raw, float) and raw != raw):
        return None
    k = _norm(raw)
    if not k:
        return None
    if k == "mba":
        return "mba"
    if k in _DEVICE_KIND_LABELS:
        return "device"
    return None


def _metadata_first_hit(metadata: dict[str, dict], folder_name: str) -> dict | None:
    for v in (
        folder_name,
        unicodedata.normalize("NFC", folder_name),
        unicodedata.normalize("NFD", folder_name),
    ):
        hit = metadata.get(_norm(v))
        if hit:
            return hit
    return None


def _metadata_keys_for_excel_name(name: str) -> list[str]:
    """Các khóa tra cứu metadata cho một dòng Excel (tên gốc + tên thư mục sau sanitize)."""
    keys: list[str] = [_norm(name)]
    try:
        from modules.kew.organize_field_zip import sanitize_device_folder

        san = sanitize_device_folder(name)
    except (ValueError, ImportError):
        san = ""
    if san:
        for v in (san, unicodedata.normalize("NFC", san), unicodedata.normalize("NFD", san)):
            nk = _norm(v)
            if nk not in keys:
                keys.append(nk)
    return keys


def _lookup_device_metadata(metadata: dict[str, dict], folder_name: str) -> dict:
    """Ghép dòng Excel với thư mục ``Project_Output/<tên>/`` (NFC/NFD, hậu tố ``_2``…)."""
    if not metadata:
        return {}
    hit = _metadata_first_hit(metadata, folder_name)
    if hit:
        return hit
    m = re.fullmatch(r"(.+)_([1-9]\d*)$", folder_name)
    if m:
        base = m.group(1)
        hit = _metadata_first_hit(metadata, base)
        if hit:
            return hit
    m2 = re.match(r"^\s*([Ss]\d{1,4})\s*[-–—]?\s*", folder_name)
    if m2:
        code = m2.group(1).upper()
        matches: list[dict] = []
        seen: set[int] = set()
        for ent in metadata.values():
            i = id(ent)
            if i in seen:
                continue
            seen.add(i)
            dn = str(ent.get("name") or "")
            if re.match(rf"^\s*{re.escape(code)}\s*[-–—]?\s*", dn, re.IGNORECASE):
                matches.append(ent)
        if len(matches) == 1:
            return matches[0]
    return {}


def _norm_voltage(value: object) -> float | None:
    if value is None:
        return None
    s = str(value).strip()
    if not s:
        return None
    m = re.search(r"(\d+(?:[.,]\d+)?)", s.replace(",", "."))
    if not m:
        return None
    try:
        v = float(m.group(1))
    except ValueError:
        return None
    if "kv" in s.lower():
        v *= 1000.0
    return v if v > 0 else None


def _excel_stt(value: object) -> int | None:
    """Số thứ tự từ cột ``stt`` (số nguyên hoặc chuỗi có chữ số)."""
    if value is None:
        return None
    if isinstance(value, float) and value != value:  # NaN
        return None
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        if isinstance(value, float) and value.is_integer():
            return int(value)
        if isinstance(value, int):
            return int(value)
    s = str(value).strip()
    if not s:
        return None
    digits = re.sub(r"\D", "", s)
    if not digits:
        return None
    return int(digits)


def _cell_text(row: object, col: str | None) -> str:
    if col is None:
        return ""
    try:
        import pandas as pd

        v = row[col]
        if v is None or (isinstance(v, float) and pd.isna(v)):
            return ""
    except Exception:
        return ""
    return str(v).strip()


def read_device_metadata_from_excel(
    excel_path: str | Path,
) -> dict[str, dict]:
    """Đọc Excel hiện trường → ``{tên_chuẩn_hóa: {name, stt, kind, nominal_voltage, remarks}}``.

    Chỉ hỗ trợ bộ cột cố định (xem ``FIELD_XLSX_HEADERS`` trong
    ``modules.kew.organize_field_zip``), gồm ``imgomit`` (chỉ dùng bước ZIP, Word bỏ qua):
    ``stt`` → thứ tự thiết bị khi ghép Word;
    ``type`` → loại section; ``pdm`` → ``nominal_voltage`` (dùng cho δU trong nhận xét thiết bị;
    MBA vẫn so bảng với 400 V); ``p``, ``pf``, ``i1``–``i3``,
    ``di``, ``thd``, ``tdd`` ghép vào ``remarks`` (ghi chú Excel, ghép sau nhận xét tự động nếu có).

    Nếu thiếu cột hoặc không đọc được file → trả về ``{}`` (báo cáo Word vẫn chạy
    không metadata).
    """
    try:
        import pandas as pd

        from modules.kew.organize_field_zip import resolve_field_excel_column_map
    except ImportError:
        return {}
    try:
        df = pd.read_excel(str(excel_path), header=0, engine="openpyxl")
    except Exception:
        return {}
    if df.empty:
        return {}

    try:
        cm = resolve_field_excel_column_map(df)
    except ValueError:
        return {}

    name_col = cm["name"]
    kind_col = cm["type"]
    nom_v_col = cm["pdm"]

    out: dict[str, dict] = {}
    for _, row in df.iterrows():
        raw_name = row[name_col]
        if raw_name is None or (isinstance(raw_name, float) and raw_name != raw_name):
            continue
        name = _nfc(raw_name)
        if not name:
            continue

        # ── Ghi chú phụ (không dùng làm số trong sinh nhận xét) ──
        # Chỉ ghi lại những trường không có ý nghĩa số học rõ ràng
        remarks = ""

        # ── Thông số đo lường hiện trường — dùng để sinh nhận xét (không dùng INPS) ──
        def _cell_val(col_key: str) -> object:
            col = cm.get(col_key)
            if col is None:
                return None
            try:
                import pandas as _pd
                v = row[col]
                if v is None or (isinstance(v, float) and _pd.isna(v)):
                    return None
                return v
            except Exception:
                return None

        def _cell_str(col_key: str) -> str | None:
            v = _cell_val(col_key)
            if v is None:
                return None
            s = str(v).strip()
            return s if s else None

        excel_params = {
            "current_char": _cell_str("current_char"),
            "u_min":        _cell_val("u_min"),
            "u_max":        _cell_val("u_max"),
            "i_max":        _cell_val("i_max"),
            "delta_u":      _cell_val("delta_u"),
            "delta_i":      _cell_val("delta_i"),
            "p":            _cell_val("p"),
            "cos_phi":      _cell_val("cos_phi"),
            "thd":          _cell_val("thd"),
            "tdd":          _cell_val("tdd"),
            "pdm":          _cell_val("pdm"),
        }

        entry = {
            "name": name,
            "stt": _excel_stt(row[cm["stt"]]),
            "kind": _norm_kind(row[kind_col]),
            "nominal_voltage": _norm_voltage(row[nom_v_col]),
            "remarks": remarks,
            "excel_params": excel_params,
        }
        for mk in _metadata_keys_for_excel_name(name):
            out[mk] = entry
    return out


def _find_first_excel(root: Path) -> Path | None:
    for p in sorted(root.rglob("*")):
        if not p.is_file():
            continue
        parts = [pp.lower() for pp in p.parts]
        if any(pp == "__macosx" for pp in parts) or p.name.startswith("._"):
            continue
        if p.suffix.lower() in {".xlsx", ".xlsm", ".xls"}:
            return p
    return None


def _find_project_root(extract_root: Path) -> Path:
    """Tìm thư mục chứa các thư mục thiết bị.

    Ưu tiên một thư mục có tên ``Project_Output``; nếu không, dùng chính
    ``extract_root``.
    """
    direct = extract_root / "Project_Output"
    if direct.is_dir():
        return direct
    for p in extract_root.rglob("Project_Output"):
        if p.is_dir():
            return p
    return extract_root


def build_word_report_from_zip(
    zip_bytes: bytes,
    output_docx: str | Path,
    *,
    mba_template: str | Path | None = None,
    device_template: str | Path | None = None,
) -> tuple[Path, list[str]]:
    """Entry-point cho tab "Tạo báo cáo Word".

    Nhận ZIP đã tổ chức (output của "Xử lý file sơ bộ"):
        ``Project_Output/<Tên thiết bị>/a.png + PS-SDxxx.BMP``

    Có thể chấp nhận ZIP không có ``Project_Output/`` (các thư mục thiết bị nằm ngay dưới gốc giải nén).

    Nếu có Excel kèm theo đủ bộ cột hiện trường (``stt``, ``name``, ``file``, ``img``,
    ``imgend``, ``imgomit``, ``type``, ``pdm``, ``p``, ``pf``, ``i1``–``i3``, ``di``, ``thd``,
    ``tdd``), các giá trị được dùng làm metadata; thứ tự section trong Word theo ``stt``;
    ``pdm`` truyền vào nhận xét thiết bị (δU so với điện áp danh định).

    Trả về ``(đường_dẫn_báo_cáo_word, warnings)``.
    """
    mba_template = Path(mba_template or DEFAULT_MBA_TEMPLATE)
    device_template = Path(device_template or DEFAULT_DEVICE_TEMPLATE)
    if not mba_template.is_file():
        raise FileNotFoundError(f"Thiếu template Word MBA: {mba_template}")
    if not device_template.is_file():
        raise FileNotFoundError(f"Thiếu template Word device: {device_template}")

    with TemporaryDirectory(prefix="word_report_") as td:
        extract = Path(td) / "in"
        extract.mkdir()
        bio = io.BytesIO(zip_bytes)
        try:
            try:
                zf = zipfile.ZipFile(bio, "r", metadata_encoding="utf-8")
            except TypeError:
                bio.seek(0)
                zf = zipfile.ZipFile(bio, "r")
            with zf:
                zf.extractall(extract)
        except zipfile.BadZipFile as e:
            raise ValueError(f"File ZIP không hợp lệ: {e}") from e

        project_root = _find_project_root(extract)
        raw_dirs = [
            d for d in project_root.iterdir()
            if d.is_dir() and not d.name.startswith(".") and d.name != "__MACOSX"
        ]
        if not raw_dirs:
            raise ValueError(
                "ZIP không chứa thư mục thiết bị nào. Cấu trúc mong đợi: "
                "Project_Output/<Tên thiết bị>/ (a.png, PS-SDxxx.BMP)."
            )

        excel_path = _find_first_excel(extract)
        metadata = read_device_metadata_from_excel(excel_path) if excel_path else {}

        _stt_fallback = 10**9

        def _device_dir_sort_key(p: Path) -> tuple[int, str]:
            if not metadata:
                return (_stt_fallback, p.name.lower())
            m = _lookup_device_metadata(metadata, p.name)
            st = m.get("stt")
            if isinstance(st, int):
                return (st, p.name.lower())
            return (_stt_fallback, p.name.lower())

        device_dirs = sorted(raw_dirs, key=_device_dir_sort_key)

        devices: list[dict] = []
        for d in device_dirs:
            meta = _lookup_device_metadata(metadata, d.name)
            display = _nfc(meta.get("name") or d.name)
            devices.append({
                "name": display,
                "folder": d,
                "kind": meta.get("kind"),
                "remarks": meta.get("remarks", ""),
                "nominal_voltage": meta.get("nominal_voltage"),
                "excel_params": meta.get("excel_params") or {},
            })

        out = Path(output_docx)
        out.parent.mkdir(parents=True, exist_ok=True)
        path, warnings = build_field_word_report(
            project_root,
            out,
            mba_template=mba_template,
            device_template=device_template,
            devices=devices,
        )
        if excel_path:
            warnings.insert(0, f"Đã dùng metadata Excel: {excel_path.name}.")
        return path, warnings
