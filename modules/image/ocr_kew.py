"""
Nhận dạng giá trị đo từ ảnh BMP KEW6315 bằng phương pháp Template Matching pixel-by-pixel.

Mỗi overlay trong SCREENS có tọa độ pixel cố định. Module này đọc ngược quy trình
``apply_text_to_image``:
- Cắt vùng ROI theo (x_right, y_bot, w_clear) đã biết từ layout.
- Quét từ phải sang trái, so khớp từng strip pixel với ảnh mẫu chữ số.
- Ghép lại thành chuỗi số → parse float.

Ánh xạ màn hình → cột Excel hiện trường:
  SD140 (offset 0): P → p, PF → cos_phi, A1/A2/A3 → i_max, V1/V2/V3 → u_min/u_max
  SD141 (offset 1): V_unb → delta_u, A_unb → delta_i
  SD144 (offset 4): THDV1/2/3 → thd (max)
  SD145 (offset 5): THDA1/2/3 → tdd (max)
"""

from __future__ import annotations

import os
import re
from typing import Optional

import numpy as np
from PIL import Image

from modules.image.kew6315_layout import SCREENS, SCREEN_BY_ID

# ── Đường dẫn mặc định đến thư mục ảnh mẫu chữ số ───────────────────────────
_DEFAULT_DIGITS_DIR = os.path.normpath(
    os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..", "static", "digits")
)

# ── Constants ─────────────────────────────────────────────────────────────────
# Chiều cao ROI (khớp với h_clear trong apply_text_to_image)
_ROI_H = 15
# Ngưỡng tương đồng tối thiểu (0-1): giá trị thấp hơn = bỏ qua digit đó
_MATCH_THRESHOLD = 0.70
# Khoảng cách cộng thêm sau mỗi chữ số (theo logic trong apply_text_to_image)
_SPACING_WIDE = 1   # khi digit_width >= 8
_SPACING_NARROW = 2 # khi digit_width < 8

# ── Map ký tự → tên file (giống CHAR_MAP trong image_api.py) ─────────────────
_CHAR_MAP = {".": "dot", "-": "minus"}

# ── Ánh xạ screen_id → offset BMP từ img_start ───────────────────────────────
SCREEN_OFFSET: dict[str, int] = {
    "SD140": 0,
    "SD141": 1,
    "SD142": 2,
    "SD143": 3,
    "SD144": 4,
    "SD145": 5,
}

# ── Ánh xạ cột Excel → (screen_id, overlay_ids, aggregation) ─────────────────
# aggregation: "first" | "max" | "min"
FIELD_SCREEN_MAP: dict[str, tuple[str, list[str], str]] = {
    "p":       ("SD140", ["P"],                    "first"),
    "cos_phi": ("SD140", ["PF"],                   "first"),
    "i_max":   ("SD140", ["A1", "A2", "A3"],        "max"),
    "u_min":   ("SD140", ["V1", "V2", "V3"],        "min"),
    "u_max":   ("SD140", ["V1", "V2", "V3"],        "max"),
    "delta_u": ("SD141", ["V_unb"],                "first"),
    "delta_i": ("SD141", ["A_unb"],                "first"),
    "thd":     ("SD144", ["THDV1", "THDV2", "THDV3"], "max"),
    "tdd":     ("SD145", ["THDA1", "THDA2", "THDA3"], "max"),
}


# ── Cache ảnh mẫu ─────────────────────────────────────────────────────────────
_tpl_cache: dict[str, Optional[np.ndarray]] = {}


def _load_template(char: str, color: str, digits_dir: str) -> Optional[np.ndarray]:
    """
    Tải ảnh mẫu chữ số dưới dạng mảng grayscale numpy.

    Args:
        char: Ký tự cần tải ('0'-'9', '.', '-').
        color: Màu nền ('w' trắng hoặc 'g' xanh).
        digits_dir: Thư mục chứa ảnh mẫu chữ số.

    Returns:
        Mảng numpy float32 (H×W) hoặc None nếu không tìm thấy file.
    """
    s = _CHAR_MAP.get(char, char)
    key = f"{s}_{color}_{digits_dir}"
    if key in _tpl_cache:
        return _tpl_cache[key]

    # Thử màu chính, sau đó fallback màu còn lại
    for c in [color, ("g" if color == "w" else "w")]:
        path = os.path.join(digits_dir, f"{s}{c}.bmp")
        if os.path.exists(path):
            arr = np.array(Image.open(path).convert("L"), dtype=np.float32)
            _tpl_cache[key] = arr
            return arr

    _tpl_cache[key] = None
    return None


def _load_all_templates(color: str, digits_dir: str) -> dict[str, np.ndarray]:
    """
    Tải tất cả ảnh mẫu chữ số cho một màu nhất định.

    Args:
        color: Màu nền ('w' hoặc 'g').
        digits_dir: Thư mục chứa ảnh mẫu.

    Returns:
        Dict mapping ký tự → mảng numpy template.
    """
    chars = list("0123456789") + [".", "-"]
    result: dict[str, np.ndarray] = {}
    for ch in chars:
        arr = _load_template(ch, color, digits_dir)
        if arr is not None:
            result[ch] = arr
    return result


def _match_score(roi_strip: np.ndarray, template: np.ndarray) -> float:
    """
    Tính điểm tương đồng giữa strip ROI và template.
    Sử dụng phạt đối xứng: phạt cả khi thiếu nét và khi thừa nét.
    """
    h, w = template.shape
    rh, rw = roi_strip.shape

    if rh < h or rw != w:
        return 0.0

    if rh > h:
        offset = (rh - h) // 2
        roi_strip = roi_strip[offset: offset + h, :]

    diff = np.abs(roi_strip - template)
    
    # ── CẢI TIẾN: Phạt đối xứng (Symmetric Penalty) ──────────────────────
    # 1. Nét chữ trong template (tối) nhưng ROI lại sáng -> Thiếu nét (phạt 1.4)
    penalty_missing = np.where((template < 100) & (roi_strip > 160), 1.4, 1.0)
    
    # 2. ROI có nét (tối) nhưng template lại sáng -> Thừa nét (phạt 1.2)
    penalty_extra = np.where((template > 180) & (roi_strip < 100), 1.2, 1.0)
    
    # Trọng số cho vùng nét chữ (nét chữ quan trọng hơn nền)
    weights = np.where(template < 128, 1.4, 1.0)
    
    weighted_diff = diff * weights * penalty_missing * penalty_extra
    
    score = 1.0 - float(weighted_diff.mean()) / 255.0
    return score


def _scan_digits_rtl(
    roi: np.ndarray,
    templates: dict[str, np.ndarray],
) -> Optional[str]:
    """
    Quét nhận dạng số từ phải sang trái.
    Bỏ qua bước nhảy cố định (spacing=0) để dò tìm linh hoạt hơn.
    """
    roi_h, roi_w = roi.shape
    result_chars: list[str] = []
    x_cursor = roi_w
    
    # Cho phép nhảy tối đa 10 pixel trống (để vượt qua dấu % hoặc đơn vị)
    MAX_SKIP = 10 
    skipped_pixels = 0

    while x_cursor > 2:
        best_char: Optional[str] = None
        best_score = _MATCH_THRESHOLD
        best_w = 0
        best_x_shift = 0

        # Cửa sổ trượt cho phép tìm ký tự trong phạm vi rộng xung quanh x_cursor
        # Thử dịch chuyển mạnh hơn (+/- 5px) để không bỏ sót số cuối
        search_range = [0, -1, 1, -2, 2, -3, 3, -4, 4, -5, 5]
        
        for x_shift in search_range:
            current_x = x_cursor + x_shift
            if current_x > roi_w: continue

            for ch, tpl in templates.items():
                tw = tpl.shape[1]
                x_start = current_x - tw
                if x_start < 0: continue

                # Dấu trừ phải nằm sát số
                if ch == "-" and (not result_chars or current_x < (x_cursor - 3)):
                    continue
                
                current_threshold = 0.92 if ch == "-" else _MATCH_THRESHOLD

                strip = roi[:, x_start:current_x]
                score = _match_score(strip, tpl)
                
                # Ưu tiên các score cao hơn vượt ngưỡng
                if score > best_score and score > current_threshold:
                    best_score = score
                    best_char = ch
                    best_w = tw
                    best_x_shift = x_shift

        if best_char is None:
            # Không tìm thấy ký tự nào đạt ngưỡng -> dịch cursor sang trái 1px và thử lại
            x_cursor -= 1
            # Nếu đã lùi quá nhiều mà không thấy gì thì dừng
            if len(result_chars) > 0 and x_cursor < (roi_w // 4):
                break
            continue

        result_chars.append(best_char)
        spacing = _SPACING_WIDE if best_w >= 8 else _SPACING_NARROW
        x_cursor = (x_cursor + best_x_shift) - (best_w + spacing)

    if not result_chars:
        return None

    result_chars.reverse()
    text = "".join(result_chars)

    # Kiểm tra chuỗi có phải là số hợp lệ không
    if re.match(r"^-?\d+(?:\.\d+)?$", text):
        return text

    # Thử xóa các ký tự rác ở hai đầu nếu chuỗi gần đúng
    cleaned = text.strip("-").strip(".")
    if re.match(r"^\d+(?:\.\d+)?$", cleaned) and cleaned:
        return cleaned

    return None


def extract_overlay_value(
    img_arr: np.ndarray,
    overlay: dict[str, Any],
    digits_dir: str = _DEFAULT_DIGITS_DIR,
) -> Optional[str]:
    """
    Trích xuất giá trị số từ một vùng overlay cụ thể trên mảng ảnh đã được xử lý.

    Args:
        img_arr: Mảng numpy float32 (Grayscale) của ảnh màn hình.
        overlay: Dict cấu hình overlay (x, y, w_clear, bg).
        digits_dir: Thư mục chứa các chữ số mẫu.

    Returns:
        Optional[str]: Chuỗi nhận dạng được (ví dụ "0.934") hoặc None.
    """
    x_right = overlay.get("x", 0)
    y_bot = overlay.get("y", 0)
    w_clear = overlay.get("w_clear", 50)
    bg_color = overlay.get("bg", "w")

    # Tính toán tọa độ ROI
    x_left = max(0, x_right - w_clear + 1)
    y_top = max(0, y_bot - _ROI_H + 1)

    # Giới hạn tọa độ trong ảnh
    h_img, w_img = img_arr.shape
    y_top = min(y_top, h_img - 1)
    y_bot = min(y_bot, h_img - 1)
    x_left = min(x_left, w_img - 1)
    x_right = min(x_right, w_img - 1)

    roi = img_arr[y_top: y_bot + 1, x_left: x_right + 1]
    if roi.size == 0:
        return None

    templates = _load_all_templates(color, digits_dir)
    if not templates:
        return None

    text = _scan_digits_rtl(roi, templates)
    if text is None:
        return None

    try:
        return float(text)
    except ValueError:
        return None


def read_screen_values(
    bmp_path: str,
    screen_id: str,
    digits_dir: str = _DEFAULT_DIGITS_DIR,
) -> dict[str, Optional[float]]:
    """
    Đọc tất cả các giá trị overlay từ một ảnh BMP của màn hình cụ thể.

    Args:
        bmp_path: Đường dẫn đến file ảnh BMP.
        screen_id: ID màn hình ('SD140', 'SD141', ...).
        digits_dir: Thư mục chứa ảnh mẫu chữ số.

    Returns:
        Dict mapping overlay_id → float (hoặc None nếu không đọc được).
    """
    if screen_id not in SCREEN_BY_ID:
        return {}

    screen = SCREEN_BY_ID[screen_id]
    overlays = screen.get("overlays", [])

    try:
        img = Image.open(bmp_path).convert("RGB")
    except Exception:
        return {ov.get("id", ""): None for ov in overlays}

    result: dict[str, Optional[float]] = {}
    for overlay in overlays:
        ov_id = overlay.get("id", "")
        if not ov_id:
            continue
        val = extract_overlay_value(img, overlay, digits_dir)
        result[ov_id] = val

    return result


def _aggregate(values: list[Optional[float]], mode: str) -> Optional[float]:
    """
    Tổng hợp danh sách giá trị theo chế độ: 'first', 'max', 'min'.

    Args:
        values: Danh sách giá trị (có thể chứa None).
        mode: Phương thức tổng hợp.

    Returns:
        Giá trị tổng hợp hoặc None nếu tất cả đều None.
    """
    valid = [v for v in values if v is not None]
    if not valid:
        return None
    if mode == "first":
        return valid[0]
    if mode == "max":
        return max(valid)
    if mode == "min":
        return min(valid)
    return valid[0]


def read_device_ocr(
    bmp_indices: list[int],
    bmp_map: dict[int, str],
    digits_dir: str = _DEFAULT_DIGITS_DIR,
) -> tuple[dict[str, Optional[float]], list[str]]:
    """
    Nhận dạng thông số kỹ thuật từ bộ ảnh BMP đã được lọc (theo thứ tự SD140-SD145).

    Args:
        bmp_indices: Danh sách các chỉ số ảnh thực tế (đã loại bỏ omit).
                     Phần tử 0 tương ứng SD140, phần tử 1 tương ứng SD141, ...
        bmp_map: Dict mapping chỉ số ảnh → đường dẫn file BMP.
        digits_dir: Thư mục chứa ảnh mẫu chữ số.

    Returns:
        tuple: (dict kết quả, list cảnh báo).
    """
    warnings: list[str] = []
    # Cache giá trị đọc từng màn hình: screen_id → {overlay_id: value}
    screen_cache: dict[str, dict[str, Optional[float]]] = {}

    # Map từ screen_id sang vị trí trong danh sách bmp_indices
    # SD140: 0, SD141: 1, ... SD145: 5
    screen_to_pos = {
        "SD140": 0, "SD141": 1, "SD142": 2,
        "SD143": 3, "SD144": 4, "SD145": 5
    }

    # Chỉ đọc các màn hình thực sự cần dùng
    needed_screens = {info[0] for info in FIELD_SCREEN_MAP.values()}

    for screen_id in needed_screens:
        pos = screen_to_pos.get(screen_id)
        if pos is None or pos >= len(bmp_indices):
            # Không có ảnh tương ứng cho màn hình này
            screen_cache[screen_id] = {}
            continue
        
        bmp_idx = bmp_indices[pos]
        bmp_path = bmp_map.get(bmp_idx)
        if bmp_path is None:
            warnings.append(
                f"Không tìm thấy ảnh PS-SD{bmp_idx:03d}.BMP cho màn hình {screen_id}."
            )
            screen_cache[screen_id] = {}
            continue

        if not os.path.exists(bmp_path):
            warnings.append(
                f"File ảnh PS-SD{bmp_idx:03d}.BMP không tồn tại tại: {bmp_path}"
            )
            screen_cache[screen_id] = {}
            continue

        screen_cache[screen_id] = read_screen_values(bmp_path, screen_id, digits_dir)

    # Tổng hợp kết quả theo FIELD_SCREEN_MAP
    result: dict[str, Optional[float]] = {}
    for field_name, (screen_id, overlay_ids, agg_mode) in FIELD_SCREEN_MAP.items():
        sv = screen_cache.get(screen_id, {})
        raw_vals = [sv.get(ov_id) for ov_id in overlay_ids]
        val = _aggregate(raw_vals, agg_mode)
        result[field_name] = val
        if val is None:
            warnings.append(
                f"OCR: không đọc được giá trị '{field_name}' "
                f"(screen {screen_id}, overlays {overlay_ids})."
            )

    return result, warnings
