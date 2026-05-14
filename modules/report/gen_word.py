"""Context docxtpl + sinh báo cáo Word cho luồng "Tạo báo cáo Word".

API chính:
* ``mba(doc, ...)``, ``device(doc, ...)`` — dựng context cho từng template.
* ``merge_rendered_docx`` / ``merge_mba_device_docx`` — ghép nhiều file đã render.
* ``mba_kwargs_from_inps`` / ``mba_kwargs_from_folder`` / ``device_kwargs_from_folder`` —
  tự chọn ảnh từ thư mục thiết bị (**bắt buộc** ``a.png`` + PS-SDxxx.BMP). Với MBA,
  ``mba_kwargs_from_folder`` tự tìm ``INPS*.KEW`` (cùng quy ước ``find_file`` như phân tích KEW
  và như luồng Excel MBA) rồi gọi ``mba_kwargs_from_inps``; thiếu INPS thì bảng dùng ``"—"``.
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

# Giới hạn / tiêu chuẩn dùng cho cột "Đánh giá" trong báo cáo MBA.
_V_DEV_LIMIT_PCT = 5.0
_PF_LIMIT = 0.9
_THDV_LIMIT_PCT = 8.0
_TDD_LIMIT_PCT = 12.0
# Điện áp danh định cho cột đánh giá lệch % U12: -5% ≤ δ ≤ +5% (so với 400 V), không lấy từ Excel/metadata.
_MBA_NOMINAL_VOLTAGE_V = 400.0
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
        "cap_device": cap_device,
        "remarks_device": remarks_device,
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
    return "Đã đạt" if abs(pf_avg) >= _PF_LIMIT else "Cần lắp bù"


def _eval_thd(values: Iterable[float | None], limit: float) -> str:
    vals = [v for v in values if v is not None]
    if not vals:
        return "—"
    return "Đạt" if max(vals) < limit else "Chưa đạt"


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
) -> dict:
    """Tự chọn ảnh trong ``folder``, tìm ``INPS*.KEW`` rồi dựng kwargs MBA (giống cơ sở dữ liệu Excel MBA)."""
    folder = Path(folder)
    from modules.kew.analyse_kew import find_file  # type: ignore

    inps_path = find_file(str(folder), "INPS")
    images = auto_pick_mba_images(folder)
    return mba_kwargs_from_inps(
        inps_path,
        name=name,
        cap_fig_mba=cap_fig_mba,
        remarks_mba=remarks_mba,
        cap_tab_mba=cap_tab_mba,
        **images,
    )


def device_kwargs_from_folder(
    folder: str | Path,
    *,
    name: str,
    cap_device: str | None = None,
    remarks_device: str = "",
) -> dict:
    """Trả về ``kwargs`` cho :func:`device` (chỉ cần ảnh + caption)."""
    images = auto_pick_device_images(folder)
    return {
        "name": name,
        "cap_device": cap_device if cap_device is not None else f"Hình ảnh đo tại {name}",
        "remarks_device": remarks_device,
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

    ``devices`` (tuỳ chọn): danh sách ``{name, folder, kind?, remarks?}``.
        * ``folder`` có thể là tên thư mục con trong ``project_root`` hoặc đường dẫn tuyệt đối.
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

        try:
            if kind == "mba":
                kwargs = mba_kwargs_from_folder(
                    folder, name=name, remarks_mba=remarks,
                )
            else:
                kwargs = device_kwargs_from_folder(folder, name=name, remarks_device=remarks)
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
    ``modules.kew.organize_field_zip``): ``stt`` → thứ tự thiết bị khi ghép Word;
    ``type`` → loại section; ``pdm`` → ``nominal_voltage`` (chỉ lưu trong metadata,
    không dùng cho đánh giá lệch % điện áp MBA — MBA luôn so với 400 V); ``p``, ``pf``, ``i1``–``i3``,
    ``di``, ``thd``, ``tdd`` ghép vào ``remarks`` dạng ``P=…; PF=…``.

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
        extra_parts: list[str] = []
        for key, label in _EXCEL_METRIC_REMARKS:
            t = _cell_text(row, cm[key])
            if t:
                extra_parts.append(f"{label}={t}")
        remarks = "; ".join(extra_parts)

        entry = {
            "name": name,
            "stt": _excel_stt(row[cm["stt"]]),
            "kind": _norm_kind(row[kind_col]),
            "nominal_voltage": _norm_voltage(row[nom_v_col]),
            "remarks": remarks,
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
    ``imgend``, ``type``, ``pdm``, ``p``, ``pf``, ``i1``–``i3``, ``di``, ``thd``,
    ``tdd``), các giá trị được dùng làm metadata; thứ tự section trong Word theo ``stt``.

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
