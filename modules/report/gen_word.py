"""Context docxtpl + sinh báo cáo Word cho luồng "Tạo báo cáo Word".

API chính:
* ``mba(doc, ...)``, ``device(doc, ...)`` — dựng context cho từng template.
* ``merge_rendered_docx`` / ``merge_mba_device_docx`` — ghép nhiều file đã render.
* ``mba_kwargs_from_inps`` / ``device_kwargs_from_folder`` — tự tổng hợp số liệu &
  chọn ảnh từ một thư mục thiết bị (INPSxxxx.KEW + PS-SDxxx.BMP).
* ``build_field_word_report`` — quét một thư mục ``Project_Output/`` rồi xuất
  1 file Word duy nhất gồm nhiều MBA / device.
* ``build_word_report_from_zip`` — entry-point cho API: nhận ZIP đã tổ chức
  (output của "Xử lý file sơ bộ"), tự dò metadata trong Excel kèm (nếu có),
  trả về đường dẫn báo cáo Word.

Tham số / khóa template — xem ``modules/report/context_keys.json``.
"""

from __future__ import annotations

import io
import os
import re
import unicodedata
import zipfile
from pathlib import Path
from shutil import copy2
from tempfile import TemporaryDirectory
from typing import Iterable, Literal, Mapping, Sequence

from docx import Document
from docx.shared import Mm
from docxcompose.composer import Composer
from docxtpl import DocxTemplate, InlineImage

__all__ = [
    "mba",
    "device",
    "merge_rendered_docx",
    "merge_mba_device_docx",
    "mba_kwargs_from_inps",
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

# Giới hạn / tiêu chuẩn dùng cho cột "Đánh giá" trong báo cáo MBA.
_V_DEV_LIMIT_PCT = 5.0
_PF_LIMIT = 0.9
_THDV_LIMIT_PCT = 8.0
_TDD_LIMIT_PCT = 12.0
_STANDARD_VOLTAGES = (110, 127, 220, 230, 240, 380, 400, 415, 440, 480, 600, 690, 1000)
_BMP_RE = re.compile(r"PS-?SD?(\d{1,4})\.BMP$", re.IGNORECASE)


# ════════════════════════════════════════════════════════════════════
#                     Context builders (template)
# ════════════════════════════════════════════════════════════════════


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
    ii = lambda p, h, w: InlineImage(doc, p, height=h, width=w)
    return {
        "mba_name": name,
        "imga": ii(imga, HEIGHT_MBA, WIDTH_LARGE),
        "img1": ii(img1, HEIGHT_MBA, WIDTH_SMALL),
        "img2": ii(img2, HEIGHT_MBA, WIDTH_SMALL),
        "img4": ii(img4, HEIGHT_MBA, WIDTH_SMALL),
        "img6": ii(img6, HEIGHT_MBA, WIDTH_SMALL),
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
    ii = lambda p, h, w: InlineImage(doc, p, height=h, width=w)
    return {
        "device_name": name,
        "imga": ii(imga, HEIGHT_A, WIDTH_A),
        "img1": ii(img1, HEIGHT_SUB, WIDTH_SUB),
        "img2": ii(img2, HEIGHT_SUB, WIDTH_SUB),
        "img3": ii(img3, HEIGHT_SUB, WIDTH_SUB),
        "img4": ii(img4, HEIGHT_SUB, WIDTH_SUB),
        "img5": ii(img5, HEIGHT_SUB, WIDTH_SUB),
        "img6": ii(img6, HEIGHT_SUB, WIDTH_SUB),
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


def _fp(v, d: int = 2) -> str:
    """Tương tự ``_f`` nhưng có dấu ``+`` cho số dương."""
    if v is None:
        return "—"
    try:
        x = float(v)
    except (TypeError, ValueError):
        return "—"
    if x != x:
        return "—"
    sign = "+" if x > 0 else ""
    return f"{sign}{x:.{d}f}".replace(".", ",")


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


def _nearest_nominal(v_measured: float | None, hint: float | None = None) -> float:
    """Chọn điện áp định mức gần nhất từ giá trị đo (hoặc gợi ý nếu có)."""
    if hint is not None:
        try:
            h = float(hint)
            if h > 0:
                return h
        except (TypeError, ValueError):
            pass
    if v_measured is None or v_measured <= 0:
        return 400.0
    return min(_STANDARD_VOLTAGES, key=lambda sv: abs(sv - v_measured))


def _eval_voltage(u_max, u_min, vref: float) -> tuple[str, float, float, float]:
    """Trả về (đánh giá, δmax, δmin, δavg) cho dải điện áp."""
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
    """Trả về danh sách PS-SDxxx.BMP trong ``folder``, sắp theo số thứ tự."""
    p = Path(folder)
    if not p.is_dir():
        return []
    bmps: list[tuple[int, Path]] = []
    for f in p.iterdir():
        if not f.is_file():
            continue
        m = _BMP_RE.search(f.name)
        if m:
            bmps.append((int(m.group(1)), f))
    bmps.sort(key=lambda x: x[0])
    # Fallback: file BMP bất kỳ nếu không có tên theo chuẩn
    if not bmps:
        bmps = [(i, f) for i, f in enumerate(sorted(p.glob("*.BMP")))]
        bmps += [(i + 10_000, f) for i, f in enumerate(sorted(p.glob("*.bmp")))]
    return [pp for _, pp in bmps]


def _take(lst: list[Path], idx: int, fallback: Path | None) -> Path | None:
    return lst[idx] if 0 <= idx < len(lst) else fallback


def auto_pick_mba_images(folder: str | Path) -> dict[str, str]:
    """Chọn 5 ảnh BMP cho template MBA (imga, img1, img2, img4, img6).

    Quy ước: ảnh #1 là tổng quan (``imga``); 4 ảnh tiếp theo theo thứ tự cho
    ``img1, img2, img4, img6``. Nếu thiếu ảnh, dùng ảnh đầu tiên làm fallback.
    """
    bmps = list_bmp_in_folder(folder)
    if not bmps:
        raise FileNotFoundError(f"Không tìm thấy ảnh BMP trong {folder!s} để dựng MBA.")
    fb = bmps[0]
    return {
        "imga": str(_take(bmps, 0, fb)),
        "img1": str(_take(bmps, 1, fb)),
        "img2": str(_take(bmps, 2, fb)),
        "img4": str(_take(bmps, 3, fb)),
        "img6": str(_take(bmps, 4, fb)),
    }


def auto_pick_device_images(folder: str | Path) -> dict[str, str]:
    """Chọn 7 ảnh BMP cho template device (imga, img1..img6)."""
    bmps = list_bmp_in_folder(folder)
    if not bmps:
        raise FileNotFoundError(f"Không tìm thấy ảnh BMP trong {folder!s} để dựng device.")
    fb = bmps[0]
    return {
        "imga": str(_take(bmps, 0, fb)),
        "img1": str(_take(bmps, 1, fb)),
        "img2": str(_take(bmps, 2, fb)),
        "img3": str(_take(bmps, 3, fb)),
        "img4": str(_take(bmps, 4, fb)),
        "img5": str(_take(bmps, 5, fb)),
        "img6": str(_take(bmps, 6, fb)),
    }


def find_inps_file(folder: str | Path) -> Path | None:
    """Tìm file ``INPSxxxx.KEW`` (không phân biệt hoa thường) trong ``folder``."""
    p = Path(folder)
    if not p.is_dir():
        return None
    for f in p.iterdir():
        if f.is_file() and f.name.upper().startswith("INPS") and f.suffix.upper() == ".KEW":
            return f
    return None


# ════════════════════════════════════════════════════════════════════
#                  Kwargs builders từ thư mục thiết bị
# ════════════════════════════════════════════════════════════════════


def mba_kwargs_from_inps(
    inps_path: str | Path,
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
    nominal_voltage: float | None = None,
) -> dict:
    """Đọc INPS, tổng hợp số liệu, trả về ``kwargs`` cho :func:`mba`.

    ``imga, img1, img2, img4, img6`` là đường dẫn ảnh (đã tự chọn từ thư mục
    hoặc do người dùng chỉ định).
    """
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

    v_avg_meas = u12.get("avg") if u12 else None
    vref = _nearest_nominal(v_avg_meas, nominal_voltage)
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

    def s1(d, dec=1):
        return _f(d.get("max"), dec), _f(d.get("min"), dec), _f(d.get("avg"), dec)

    def s2(d, dec=2):
        return _f(d.get("max"), dec), _f(d.get("min"), dec), _f(d.get("avg"), dec)

    def s3(d, dec=3):
        return _f(d.get("max"), dec), _f(d.get("min"), dec), _f(d.get("avg"), dec)

    u12max, u12min, u12avg = s1(u12)
    u23max, u23min, u23avg = s1(u23)
    u31max, u31min, u31avg = s1(u31)
    i1max, i1min, i1avg = s2(i1)
    i2max, i2min, i2avg = s2(i2)
    i3max, i3min, i3avg = s2(i3)
    dumax, dumin, duavg = s3(uv_unb)
    dimax, dimin, diavg = s3(ua_unb)
    pfmax, pfmin, pfavg = s3(pf)
    pmax, pmin, pavg = s2(p_k)
    qmax, qmin, qavg = s2(q_k)
    smax, smin, savg = s2(s_k)
    thd1max, thd1min, thd1avg = s2(thd1)
    thd2max, thd2min, thd2avg = s2(thd2)
    thd3max, thd3min, thd3avg = s2(thd3)
    tdd1max, tdd1min, tdd1avg = s2(tdd1)
    tdd2max, tdd2min, tdd2avg = s2(tdd2)
    tdd3max, tdd3min, tdd3avg = s2(tdd3)

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
    nominal_voltage: float | None = None,
) -> dict:
    """Tiện ích: tự tìm INPS + ảnh trong ``folder`` rồi gọi ``mba_kwargs_from_inps``."""
    folder = Path(folder)
    inps = find_inps_file(folder)
    if inps is None:
        raise FileNotFoundError(f"Không tìm thấy file INPSxxxx.KEW trong {folder}")
    images = auto_pick_mba_images(folder)
    return mba_kwargs_from_inps(
        inps,
        name=name,
        cap_fig_mba=cap_fig_mba,
        remarks_mba=remarks_mba,
        cap_tab_mba=cap_tab_mba,
        nominal_voltage=nominal_voltage,
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

    with TemporaryDirectory() as td:
        tmp = Path(td)
        rendered: list[Path] = []
        for i, (kind, spec) in enumerate(work):
            if kind == "mba":
                tpl = DocxTemplate(str(mba_tpl))
                tpl.render(mba(tpl, **spec))
                path = tmp / f"mba_{i}.docx"
            elif kind == "device":
                tpl = DocxTemplate(str(device_tpl))
                tpl.render(device(tpl, **spec))
                path = tmp / f"device_{i}.docx"
            else:
                raise ValueError(f"Loại section không hợp lệ: {kind!r} (chỉ 'mba' hoặc 'device').")
            tpl.save(str(path))
            rendered.append(path)
        return merge_rendered_docx(rendered, output_path)


# ════════════════════════════════════════════════════════════════════
#                Pipeline cho luồng "Xử lý file sơ bộ"
# ════════════════════════════════════════════════════════════════════


_MBA_NAME_RE = re.compile(r"^(MBA|TR|TBA|T\d|MBT)\b|MÁY BIẾN ÁP|BIẾN ÁP", re.IGNORECASE)


def _guess_kind(name: str) -> SectionKind:
    """Đoán loại template từ tên thiết bị."""
    return "mba" if _MBA_NAME_RE.search(name or "") else "device"


def build_field_word_report(
    project_root: str | Path,
    output_path: str | Path,
    *,
    mba_template: str | Path,
    device_template: str | Path,
    devices: Sequence[Mapping] | None = None,
    default_kind: SectionKind | None = None,
    nominal_voltage: float | None = None,
) -> tuple[Path, list[str]]:
    """Quét ``project_root`` (= ``Project_Output/``) → xuất 1 file Word tổng hợp.

    ``devices`` (tuỳ chọn): danh sách ``{name, folder, kind?, nominal_voltage?, remarks?}``.
        * ``folder`` có thể là tên thư mục con trong ``project_root`` hoặc đường dẫn tuyệt đối.
        * Nếu ``kind`` không có:
          - dùng ``default_kind`` nếu được chỉ định,
          - ngược lại đoán theo tên (``MBA…`` → ``mba``, còn lại → ``device``).
    Khi ``devices=None``: tự duyệt mọi thư mục con của ``project_root`` (sort theo tên).

    Trả về ``(đường_dẫn_báo_cáo, warnings)``.
    """
    root = Path(project_root)
    if not root.is_dir():
        raise FileNotFoundError(f"Không tìm thấy thư mục Project_Output: {root}")

    if devices is None:
        devices = [
            {"name": d.name, "folder": d}
            for d in sorted(root.iterdir())
            if d.is_dir() and not d.name.startswith(".") and d.name != "__MACOSX"
        ]
    if not devices:
        raise ValueError("Không có thiết bị nào để dựng báo cáo Word.")

    warnings: list[str] = []
    sections: list[tuple[SectionKind, dict]] = []
    for spec in devices:
        name = str(spec.get("name") or "").strip()
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

        kind = spec.get("kind") or default_kind or _guess_kind(name)
        remarks = str(spec.get("remarks") or "")
        nom_v = spec.get("nominal_voltage", nominal_voltage)

        try:
            if kind == "mba":
                if find_inps_file(folder) is None:
                    warnings.append(f"«{name}»: thiếu INPS — chuyển sang template device.")
                    kind = "device"
            if kind == "mba":
                kwargs = mba_kwargs_from_folder(
                    folder, name=name, remarks_mba=remarks, nominal_voltage=nom_v,
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


def _norm_kind(value: object) -> SectionKind | None:
    k = _norm(value)
    if not k:
        return None
    if k in {"mba", "máy biến áp", "may bien ap", "transformer", "tr", "mbt", "tba"}:
        return "mba"
    if k in {"device", "thiết bị", "thiet bi", "tủ", "tu", "khac", "khác"}:
        return "device"
    return None


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


def read_device_metadata_from_excel(
    excel_path: str | Path,
) -> dict[str, dict]:
    """Đọc Excel hiện trường → ``{tên_chuẩn_hóa: {name, kind, nominal_voltage, remarks}}``.

    Các cột Excel: ``Tên thiết bị`` (bắt buộc); ``Loại``, ``Điện áp định mức``,
    ``Nhận xét`` (đều tuỳ chọn). Trả về dict rỗng nếu không đọc được hoặc thiếu
    cột tên — caller tự fallback.
    """
    try:
        import pandas as pd
    except ImportError:
        return {}
    try:
        df = pd.read_excel(str(excel_path), header=0, engine="openpyxl")
    except Exception:
        return {}
    if df.empty:
        return {}

    col_map = {_norm(c): c for c in df.columns}

    def pick(*aliases: str) -> str | None:
        for a in aliases:
            v = col_map.get(_norm(a))
            if v:
                return v
        return None

    name_col = pick("tên thiết bị", "ten thiet bi", "thiết bị", "thiet bi", "device")
    if not name_col:
        return {}
    kind_col = pick("loại", "loai", "kiểu", "kieu", "type")
    nom_v_col = pick("điện áp định mức", "dien ap dinh muc", "u định mức",
                     "u dinh muc", "vnom", "u_nom", "u nominal", "nominal voltage")
    remarks_col = pick("nhận xét", "nhan xet", "ghi chú", "ghi chu", "remarks", "notes")

    out: dict[str, dict] = {}
    for _, row in df.iterrows():
        raw_name = row[name_col]
        if raw_name is None or (isinstance(raw_name, float) and raw_name != raw_name):
            continue
        name = str(raw_name).strip()
        if not name:
            continue
        out[_norm(name)] = {
            "name": name,
            "kind": _norm_kind(row[kind_col]) if kind_col else None,
            "nominal_voltage": _norm_voltage(row[nom_v_col]) if nom_v_col else None,
            "remarks": str(row[remarks_col]).strip()
            if remarks_col and row[remarks_col] is not None
            and not (isinstance(row[remarks_col], float) and row[remarks_col] != row[remarks_col])
            else "",
        }
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
        ``Project_Output/<Tên thiết bị>/INPSxxxx.KEW + PS-SDxxx.BMP``

    Có thể chấp nhận ZIP không có ``Project_Output/`` (thư mục thiết bị nằm
    ngay ở gốc) — sẽ duyệt từ gốc. Nếu có Excel kèm theo (cùng cấu trúc cột
    của tab "Xử lý file sơ bộ"), các cột tuỳ chọn ``Loại`` / ``Điện áp định mức``
    / ``Nhận xét`` sẽ được dùng làm metadata cho từng thiết bị.

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
        try:
            with zipfile.ZipFile(io.BytesIO(zip_bytes), "r") as zf:
                zf.extractall(extract)
        except zipfile.BadZipFile as e:
            raise ValueError(f"File ZIP không hợp lệ: {e}") from e

        project_root = _find_project_root(extract)
        device_dirs = [
            d for d in sorted(project_root.iterdir())
            if d.is_dir() and not d.name.startswith(".") and d.name != "__MACOSX"
        ]
        if not device_dirs:
            raise ValueError(
                "ZIP không chứa thư mục thiết bị nào. Cấu trúc mong đợi: "
                "Project_Output/<Tên thiết bị>/INPSxxxx.KEW + PS-SDxxx.BMP."
            )

        excel_path = _find_first_excel(extract)
        metadata = read_device_metadata_from_excel(excel_path) if excel_path else {}

        devices: list[dict] = []
        for d in device_dirs:
            meta = metadata.get(_norm(d.name), {})
            devices.append({
                "name": meta.get("name") or d.name,
                "folder": d,
                "kind": meta.get("kind"),
                "nominal_voltage": meta.get("nominal_voltage"),
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
