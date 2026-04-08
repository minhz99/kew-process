from __future__ import annotations

import re
from dataclasses import dataclass
from functools import lru_cache
from pathlib import Path
from typing import Iterable

import numpy as np
from PIL import Image

from modules.image.kew6315_layout import (
    KEW6315_REF_HEIGHT,
    KEW6315_REF_WIDTH,
    SCREEN_BY_INDEX,
)

FIELD_HEIGHT = 15
MAX_FIELD_CHARS = 8
COLOR_DISTANCE_MAX = np.sqrt(3.0 * 255.0 * 255.0)
FOREGROUND_THRESHOLD = 0.085
ACTIVE_COLUMN_THRESHOLD = 0.10
MAX_CHAR_SCORE = 0.36
DEFAULT_DIGITS_DIR = Path(__file__).resolve().parents[2] / "static" / "digits"
CHAR_TO_FILENAME = {".": "dot", "-": "minus"}
SUPPORTED_CHARS = "0123456789.-"
BACKGROUND_BY_KEY = {
    "w": np.array([255.0, 255.0, 255.0], dtype=np.float32),
    "g": np.array([218.0, 255.0, 170.0], dtype=np.float32),
}


@dataclass(frozen=True)
class DigitTemplate:
    char: str
    color: str
    width: int
    height: int
    spacing: int
    fg_left: int
    fg_right: int
    foreground_mask: np.ndarray
    contrast_map: np.ndarray


def _as_rgb_array(image: Image.Image) -> np.ndarray:
    return np.asarray(image.convert("RGB"), dtype=np.float32)


def _contrast_from_background(arr: np.ndarray, background: np.ndarray) -> np.ndarray:
    return np.linalg.norm(arr - background, axis=2) / COLOR_DISTANCE_MAX


def _field_background(field_arr: np.ndarray) -> np.ndarray:
    left_strip = field_arr[:, : min(3, field_arr.shape[1])]
    samples = left_strip.reshape(-1, left_strip.shape[-1]).astype(np.float32)
    return np.median(samples, axis=0)


@lru_cache(maxsize=2)
def _load_templates(digits_dir: str = str(DEFAULT_DIGITS_DIR)) -> dict[str, list[DigitTemplate]]:
    result: dict[str, list[DigitTemplate]] = {"w": [], "g": []}
    base_dir = Path(digits_dir)

    for color in result:
        for char in SUPPORTED_CHARS:
            name = CHAR_TO_FILENAME.get(char, char)
            path = base_dir / f"{name}{color}.bmp"
            if not path.exists():
                continue

            image = Image.open(path).convert("RGB")
            arr = _as_rgb_array(image)
            background = BACKGROUND_BY_KEY[color]
            contrast = _contrast_from_background(arr, background)
            foreground_mask = contrast > FOREGROUND_THRESHOLD
            foreground_columns = np.where(foreground_mask.any(axis=0))[0]
            result[color].append(
                DigitTemplate(
                    char=char,
                    color=color,
                    width=image.width,
                    height=image.height,
                    spacing=1 if image.width >= 8 else 2,
                    fg_left=int(foreground_columns[0]),
                    fg_right=int(foreground_columns[-1]),
                    foreground_mask=foreground_mask,
                    contrast_map=contrast,
                )
            )

    return result


def _open_reference_image(image_or_path) -> Image.Image:
    if isinstance(image_or_path, Image.Image):
        image = image_or_path.convert("RGB")
    else:
        image = Image.open(image_or_path).convert("RGB")

    if image.size != (KEW6315_REF_WIDTH, KEW6315_REF_HEIGHT):
        image = image.resize((KEW6315_REF_WIDTH, KEW6315_REF_HEIGHT), Image.Resampling.BILINEAR)

    return image


def _match_template(
    field_contrast: np.ndarray,
    cursor: int,
    template: DigitTemplate,
) -> tuple[float, int] | None:
    left = cursor - template.fg_right
    right = left + template.width - 1
    if left < 0:
        return None
    if right >= field_contrast.shape[1]:
        return None

    patch_top = field_contrast.shape[0] - template.height
    patch_contrast = field_contrast[patch_top:, left:right + 1]
    if patch_contrast.shape != template.foreground_mask.shape:
        return None
    patch_mask = patch_contrast > FOREGROUND_THRESHOLD

    missing_foreground = np.logical_and(template.foreground_mask, ~patch_mask).mean()
    extra_foreground = np.logical_and(~template.foreground_mask, patch_mask).mean()
    contrast_error = np.abs(template.contrast_map - patch_contrast).mean()
    score = contrast_error + (missing_foreground * 2.4) + (extra_foreground * 1.2)

    if score > MAX_CHAR_SCORE:
        return None

    return score, left


def _read_field(field_arr: np.ndarray, color: str, digits_dir: str) -> str | None:
    templates = _load_templates(digits_dir).get(color, [])
    if not templates:
        return None

    background = BACKGROUND_BY_KEY.get(color, _field_background(field_arr))
    field_contrast = _contrast_from_background(field_arr, background)
    active_columns = field_contrast.max(axis=0) > ACTIVE_COLUMN_THRESHOLD

    def skip_blank(cursor: int) -> int:
        while cursor >= 0 and not active_columns[cursor]:
            cursor -= 1
        return cursor

    @lru_cache(maxsize=None)
    def solve(cursor: int, steps: int) -> tuple[float, str] | None:
        cursor = skip_blank(cursor)
        if cursor < 0:
            return 0.0, ""
        if steps >= MAX_FIELD_CHARS:
            return None

        best: tuple[float, str] | None = None
        for template in templates:
            matched = _match_template(field_contrast, cursor, template)
            if matched is None:
                continue

            score, left = matched
            next_result = solve(left - template.spacing, steps + 1)
            if next_result is None:
                continue

            total_score = score + next_result[0]
            text_reversed = template.char + next_result[1]
            if best is None or total_score < best[0]:
                best = (total_score, text_reversed)

        return best

    best = solve(field_arr.shape[1] - 1, 0)
    if best is None:
        return None

    text = best[1][::-1].strip()
    return text or None


def _crop_field(screen_arr: np.ndarray, overlay: dict) -> np.ndarray:
    x_right = overlay["x"]
    y_bottom = overlay["y"]
    width = overlay.get("w_clear", 50)
    x_left = max(0, x_right - width + 1)
    y_top = max(0, y_bottom - FIELD_HEIGHT + 1)
    return screen_arr[y_top:y_bottom + 1, x_left:x_right + 1]


def read_kew6315_screen_fields(
    image_or_path,
    screen_idx: int,
    field_ids: Iterable[str] | None = None,
    digits_dir: str | Path | None = None,
) -> dict[str, str | None]:
    screen = SCREEN_BY_INDEX.get(screen_idx)
    if screen is None:
        raise ValueError(f"Unsupported KEW6315 screen index: {screen_idx}")

    selected = set(field_ids or [])
    reference_img = _open_reference_image(image_or_path)
    screen_arr = _as_rgb_array(reference_img)
    digits_path = str(digits_dir or DEFAULT_DIGITS_DIR)

    result: dict[str, str | None] = {}
    for overlay in screen["overlays"]:
        field_id = overlay["id"]
        if selected and field_id not in selected:
            continue

        field_arr = _crop_field(screen_arr, overlay)
        result[field_id] = _read_field(field_arr, overlay.get("bg", "w"), digits_path)

    return result


def coerce_number(text: str | None) -> float | None:
    if text is None:
        return None

    cleaned = re.sub(r"[^0-9.\-]", "", text.replace(",", "."))
    if not cleaned:
        return None

    if cleaned.count("-") > 1:
        cleaned = ("-" if cleaned.startswith("-") else "") + cleaned.replace("-", "")
    if cleaned.count(".") > 1:
        head, *tail = cleaned.split(".")
        cleaned = head + "." + "".join(tail)

    try:
        return float(cleaned)
    except ValueError:
        return None
