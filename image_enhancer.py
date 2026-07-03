#!/usr/bin/env python3
"""
image_enhancer.py — Pre-OCR image enhancement pipeline.

Responsibilities:
  1. Deskew          — detect + correct rotational skew
  2. Denoise         — remove speckle / JPEG artifacts
  3. Border crop     — strip dark scan borders
  4. Contrast/levels — auto-level, CLAHE, global boost, sharpen
  5. Binarize        — adaptive Gaussian threshold → B&W
  6. Upscale         — bicubic upscale when source DPI too low
  7. Presets         — SCAN / PHOTO / MIXED / FAST / AUTO profiles

Public API:
    enhance(img, *, preset="auto", deskew=True, denoise=True,
            contrast=True, binarize=False, upscale=True,
            crop_border=True, target_dpi=300) -> PIL.Image.Image

    enhance_bytes(raw, *, ext=".jpg", **kwargs) -> PIL.Image.Image

    batch_enhance(images, *, workers=4, **kwargs) -> List[PIL.Image.Image]

Dependencies:
    pip install pillow numpy
    pip install opencv-python-headless   # optional — stronger deskew + denoise
"""

from __future__ import annotations

import io
import logging
import math
import os
import warnings
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import dataclass, field
from enum import Enum
from typing import Dict, List, Optional, Tuple

import numpy as np
from PIL import Image, ImageFilter, ImageOps, ImageEnhance

logger = logging.getLogger(__name__)

# ══════════════════════════════════════════════════════════════════════════════
# Constants & tunables
# ══════════════════════════════════════════════════════════════════════════════

DESKEW_MAX_ANGLE        = 15.0   # degrees — ignore detections beyond this
DESKEW_MIN_CONTOUR_AREA = 300    # px² — ignore tiny noise contours
DESKEW_KERNEL_W         = 30     # px — dilation kernel width for line detection
DESKEW_MIN_ANGLE_DEG    = 0.3    # degrees — skip rotation if smaller than this

DENOISE_NLM_H           = 10     # NlMeans filter strength
DENOISE_NLM_TEMPLATE    = 7      # NlMeans template window
DENOISE_NLM_SEARCH      = 21     # NlMeans search window
DENOISE_MEDIAN_RADIUS   = 1      # PIL MedianFilter radius (size = 2r+1)

CONTRAST_AUTOCONTRAST_CUT = 2    # % to clip at each end of histogram
CONTRAST_BOOST_FACTOR   = 1.4    # ImageEnhance.Contrast multiplier
SHARPNESS_FACTOR        = 1.5    # ImageEnhance.Sharpness multiplier
CLAHE_CLIP_LIMIT        = 2.0    # cv2 CLAHE clip limit
CLAHE_TILE_GRID         = (8, 8) # cv2 CLAHE tile grid size

BINARIZE_BLOCK_SIZE     = 31     # adaptive threshold block size (must be odd)
BINARIZE_C              = 10     # adaptive threshold constant
BINARIZE_SIMPLE_THRESH  = 180    # PIL fallback global threshold

BORDER_CROP_THRESHOLD   = 30     # px brightness below which is "border"
BORDER_CROP_MARGIN      = 4      # px extra margin after crop

UPSCALE_MIN_DPI         = 200    # DPI below which we upscale
UPSCALE_TARGET_DPI      = 300    # target DPI for upscale
UPSCALE_MIN_PX          = 1000   # if width < this, upscale regardless of DPI


# ══════════════════════════════════════════════════════════════════════════════
# Preset system
# ══════════════════════════════════════════════════════════════════════════════

class Preset(str, Enum):
    AUTO   = "auto"    # detect based on image characteristics
    SCAN   = "scan"    # clean flatbed scan — full pipeline + binarize
    PHOTO  = "photo"   # photographed document — deskew + denoise + contrast
    MIXED  = "mixed"   # mixed content (text + images) — no binarize
    FAST   = "fast"    # contrast only, skip deskew/denoise
    CUSTOM = "custom"  # caller controls every flag individually


@dataclass
class EnhanceConfig:
    """Full configuration for a single enhance() call."""
    preset:       Preset = Preset.AUTO
    deskew:       bool   = True
    denoise:      bool   = True
    contrast:     bool   = True
    binarize:     bool   = False
    upscale:      bool   = True
    crop_border:  bool   = True
    target_dpi:   int    = UPSCALE_TARGET_DPI

    # Advanced overrides
    contrast_factor:  float = CONTRAST_BOOST_FACTOR
    sharpness_factor: float = SHARPNESS_FACTOR
    deskew_max_angle: float = DESKEW_MAX_ANGLE

    @classmethod
    def from_preset(cls, preset: str | Preset) -> "EnhanceConfig":
        p = Preset(preset) if isinstance(preset, str) else preset
        if p == Preset.SCAN:
            return cls(preset=p, deskew=True, denoise=True, contrast=True,
                       binarize=True, upscale=True, crop_border=True)
        if p == Preset.PHOTO:
            return cls(preset=p, deskew=True, denoise=True, contrast=True,
                       binarize=False, upscale=True, crop_border=True)
        if p == Preset.MIXED:
            return cls(preset=p, deskew=True, denoise=True, contrast=True,
                       binarize=False, upscale=False, crop_border=True)
        if p == Preset.FAST:
            return cls(preset=p, deskew=False, denoise=False, contrast=True,
                       binarize=False, upscale=False, crop_border=False)
        # AUTO and CUSTOM — caller resolves
        return cls(preset=p)


@dataclass
class EnhanceResult:
    """Returned by enhance() — wraps the image with diagnostic metadata."""
    image:          Image.Image
    preset_used:    Preset
    skew_angle:     float   = 0.0
    was_deskewed:   bool    = False
    was_denoised:   bool    = False
    was_binarized:  bool    = False
    was_upscaled:   bool    = False
    upscale_factor: float   = 1.0
    border_crop:    Optional[Tuple[int,int,int,int]] = None  # (l,t,r,b) removed
    steps_applied:  List[str] = field(default_factory=list)
    warnings:       List[str] = field(default_factory=list)

    def __repr__(self) -> str:
        steps = ", ".join(self.steps_applied) or "none"
        return (f"<EnhanceResult preset={self.preset_used.value} "
                f"size={self.image.size} steps=[{steps}]>")


# ══════════════════════════════════════════════════════════════════════════════
# Internal helpers
# ══════════════════════════════════════════════════════════════════════════════

def _to_rgb(img: Image.Image) -> Image.Image:
    if img.mode == "RGB":
        return img
    return img.convert("RGB")


def _to_gray_array(img: Image.Image) -> np.ndarray:
    return np.array(img.convert("L"))


def _cv2_available() -> bool:
    try:
        import cv2  # noqa: F401
        return True
    except ImportError:
        return False


# ══════════════════════════════════════════════════════════════════════════════
# Step 1 — Deskew
# ══════════════════════════════════════════════════════════════════════════════

def _detect_skew_cv2(gray: np.ndarray, max_angle: float) -> float:
    """Detect skew angle using minAreaRect on text-line contours (cv2)."""
    import cv2
    _, binary = cv2.threshold(
        gray, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU
    )
    kernel = cv2.getStructuringElement(
        cv2.MORPH_RECT, (DESKEW_KERNEL_W, 1)
    )
    dilated = cv2.dilate(binary, kernel, iterations=2)
    contours, _ = cv2.findContours(
        dilated, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE
    )
    angles: List[float] = []
    for cnt in contours:
        if cv2.contourArea(cnt) < DESKEW_MIN_CONTOUR_AREA:
            continue
        rect = cv2.minAreaRect(cnt)
        angle = rect[-1]
        if angle < -45:
            angle += 90
        angles.append(angle)

    if not angles:
        return 0.0
    median = float(np.median(angles))
    return median if abs(median) <= max_angle else 0.0


def _detect_skew_hough(gray: np.ndarray, max_angle: float) -> float:
    """
    Fallback skew detection using Hough-line transform (numpy only).
    Less accurate but has no cv2 dependency.
    """
    # Edge-detect with a simple gradient
    gy = np.diff(gray.astype(np.int16), axis=0)
    edges = np.abs(gy) > 30
    # Collect row-wise pixel positions
    ys, xs = np.where(edges)
    if len(xs) < 50:
        return 0.0
    # Fit a line through a random sample to estimate angle
    rng = np.random.default_rng(42)
    idx = rng.choice(len(xs), size=min(500, len(xs)), replace=False)
    xs_s, ys_s = xs[idx].astype(float), ys[idx].astype(float)
    # Least-squares fit
    A = np.vstack([xs_s, np.ones(len(xs_s))]).T
    result = np.linalg.lstsq(A, ys_s, rcond=None)
    slope = result[0][0]
    angle = math.degrees(math.atan(slope))
    return angle if abs(angle) <= max_angle else 0.0


def _deskew(img: Image.Image, max_angle: float = DESKEW_MAX_ANGLE
            ) -> Tuple[Image.Image, float]:
    """
    Detect and correct document skew.
    Returns (corrected_image, angle_corrected_degrees).
    """
    gray = _to_gray_array(img)
    if _cv2_available():
        angle = _detect_skew_cv2(gray, max_angle)
    else:
        angle = _detect_skew_hough(gray, max_angle)

    if abs(angle) < DESKEW_MIN_ANGLE_DEG:
        return img, 0.0

    corrected = img.rotate(
        -angle, expand=True, fillcolor=255,
        resample=Image.Resampling.BICUBIC
    )
    logger.debug("Deskew: rotated %.2f°", angle)
    return corrected, angle


# ══════════════════════════════════════════════════════════════════════════════
# Step 2 — Denoise
# ══════════════════════════════════════════════════════════════════════════════

def _denoise(img: Image.Image) -> Tuple[Image.Image, bool]:
    """
    Remove noise. Prefers cv2 NlMeans; falls back to PIL MedianFilter.
    Returns (denoised_image, used_nlmeans).
    """
    if _cv2_available():
        import cv2
        arr = np.array(_to_rgb(img))
        try:
            denoised = cv2.fastNlMeansDenoisingColored(
                arr, None,
                DENOISE_NLM_H, DENOISE_NLM_H,
                DENOISE_NLM_TEMPLATE, DENOISE_NLM_SEARCH
            )
            return Image.fromarray(denoised), True
        except Exception as e:
            logger.warning("NlMeans failed (%s), falling back to median", e)

    size = DENOISE_MEDIAN_RADIUS * 2 + 1
    return img.filter(ImageFilter.MedianFilter(size=size)), False


# ══════════════════════════════════════════════════════════════════════════════
# Step 3 — Border crop
# ══════════════════════════════════════════════════════════════════════════════

def _detect_border(arr: np.ndarray, threshold: int = BORDER_CROP_THRESHOLD
                   ) -> Tuple[int, int, int, int]:
    """
    Return (left, top, right, bottom) border widths to crop.
    Scans from each edge inward until brightness > threshold.
    """
    gray = arr if arr.ndim == 2 else arr.mean(axis=2)
    h, w = gray.shape

    def scan_edge(values: np.ndarray) -> int:
        for i, v in enumerate(values):
            if v > threshold:
                return max(0, i - BORDER_CROP_MARGIN)
        return 0

    top    = scan_edge(gray.mean(axis=1))
    bottom = scan_edge(gray.mean(axis=1)[::-1])
    left   = scan_edge(gray.mean(axis=0))
    right  = scan_edge(gray.mean(axis=0)[::-1])

    # Safety: never crop more than 10% of dimension
    top    = min(top,    h // 10)
    bottom = min(bottom, h // 10)
    left   = min(left,   w // 10)
    right  = min(right,  w // 10)

    return left, top, right, bottom


def _crop_border(img: Image.Image
                 ) -> Tuple[Image.Image, Optional[Tuple[int,int,int,int]]]:
    arr = np.array(img.convert("L"))
    l, t, r, b = _detect_border(arr)
    if l == 0 and t == 0 and r == 0 and b == 0:
        return img, None
    w, h = img.size
    box = (l, t, w - r, h - b)
    if box[2] <= box[0] or box[3] <= box[1]:
        return img, None  # degenerate — skip
    logger.debug("Border crop: l=%d t=%d r=%d b=%d", l, t, r, b)
    return img.crop(box), (l, t, r, b)


# ══════════════════════════════════════════════════════════════════════════════
# Step 4 — Contrast, auto-levels, CLAHE, sharpening
# ══════════════════════════════════════════════════════════════════════════════

def _apply_clahe(img: Image.Image) -> Image.Image:
    """Apply CLAHE (Contrast Limited Adaptive Histogram Equalization) via cv2."""
    if not _cv2_available():
        return img
    import cv2
    gray = np.array(img.convert("L"))
    clahe = cv2.createCLAHE(
        clipLimit=CLAHE_CLIP_LIMIT, tileGridSize=CLAHE_TILE_GRID
    )
    eq = clahe.apply(gray)
    # Merge back to RGB
    rgb = np.array(_to_rgb(img))
    # Apply equalization only to luminance (convert to LAB)
    try:
        lab = cv2.cvtColor(rgb, cv2.COLOR_RGB2LAB)
        lab[:, :, 0] = clahe.apply(lab[:, :, 0])
        rgb_out = cv2.cvtColor(lab, cv2.COLOR_LAB2RGB)
        return Image.fromarray(rgb_out)
    except Exception:
        return Image.fromarray(eq).convert("RGB")


def _enhance_contrast(
    img: Image.Image,
    contrast_factor: float = CONTRAST_BOOST_FACTOR,
    sharpness_factor: float = SHARPNESS_FACTOR,
) -> Image.Image:
    """
    1. Auto-level (stretch histogram).
    2. CLAHE for local contrast (cv2).
    3. Global contrast boost.
    4. Sharpening pass.
    """
    img = ImageOps.autocontrast(img, cutoff=CONTRAST_AUTOCONTRAST_CUT)
    img = _apply_clahe(img)
    img = ImageEnhance.Contrast(img).enhance(contrast_factor)
    img = ImageEnhance.Sharpness(img).enhance(sharpness_factor)
    return img


# ══════════════════════════════════════════════════════════════════════════════
# Step 5 — Binarization
# ══════════════════════════════════════════════════════════════════════════════

def _binarize_cv2(img: Image.Image) -> Image.Image:
    """Adaptive Gaussian threshold → clean B&W via cv2."""
    import cv2
    gray = np.array(img.convert("L"))
    binary = cv2.adaptiveThreshold(
        gray, 255,
        cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
        cv2.THRESH_BINARY,
        blockSize=BINARIZE_BLOCK_SIZE,
        C=BINARIZE_C,
    )
    return Image.fromarray(binary).convert("RGB")


def _binarize_pil(img: Image.Image) -> Image.Image:
    """Simple global threshold via PIL (cv2 fallback)."""
    gray = img.convert("L")
    return gray.point(
        lambda x: 255 if x > BINARIZE_SIMPLE_THRESH else 0
    ).convert("RGB")


def _binarize(img: Image.Image) -> Image.Image:
    if _cv2_available():
        try:
            return _binarize_cv2(img)
        except Exception as e:
            logger.warning("cv2 binarize failed (%s), using PIL", e)
    return _binarize_pil(img)


# ══════════════════════════════════════════════════════════════════════════════
# Step 6 — Resolution upscale
# ══════════════════════════════════════════════════════════════════════════════

def _get_dpi(img: Image.Image) -> Optional[float]:
    """Extract DPI from image metadata if available."""
    try:
        info = img.info
        if "dpi" in info:
            dpi = info["dpi"]
            if isinstance(dpi, (tuple, list)):
                return float(dpi[0])
            return float(dpi)
        if "jfif_density" in info:
            return float(info["jfif_density"][0])
    except Exception:
        pass
    return None


def _needs_upscale(img: Image.Image, target_dpi: int = UPSCALE_TARGET_DPI) -> bool:
    """Return True if image needs upscaling for good OCR quality."""
    dpi = _get_dpi(img)
    if dpi is not None and dpi < UPSCALE_MIN_DPI:
        return True
    w, h = img.size
    if w < UPSCALE_MIN_PX:
        return True
    return False


def _upscale(img: Image.Image, target_dpi: int = UPSCALE_TARGET_DPI
             ) -> Tuple[Image.Image, float]:
    """
    Upscale image to reach target_dpi equivalent.
    Falls back to 2× upscale if DPI metadata is absent.
    """
    dpi = _get_dpi(img)
    w, h = img.size

    if dpi and dpi > 0:
        factor = target_dpi / dpi
    elif w < UPSCALE_MIN_PX:
        factor = UPSCALE_TARGET_DPI / UPSCALE_MIN_DPI  # ~1.5×
    else:
        factor = 1.5  # conservative default

    factor = max(1.1, min(factor, 4.0))  # clamp to reasonable range
    new_w = int(w * factor)
    new_h = int(h * factor)
    upscaled = img.resize((new_w, new_h), Image.Resampling.LANCZOS)
    logger.debug("Upscale: %.2f× → %dx%d", factor, new_w, new_h)
    return upscaled, factor


# ══════════════════════════════════════════════════════════════════════════════
# Step 7 — Auto preset detection
# ══════════════════════════════════════════════════════════════════════════════

def _auto_detect_preset(img: Image.Image) -> Preset:
    """
    Heuristically decide the best preset based on image statistics.
    - Low std-dev + mostly white → SCAN (clean flatbed)
    - High std-dev + many colors → PHOTO
    - Default → MIXED
    """
    arr = np.array(img.convert("L"), dtype=np.float32)
    std  = float(arr.std())
    mean = float(arr.mean())
    # High mean + low std = mostly-white clean scan
    if mean > 200 and std < 40:
        return Preset.SCAN
    # Very high std = noisy photo
    if std > 80:
        return Preset.PHOTO
    return Preset.MIXED


# ══════════════════════════════════════════════════════════════════════════════
# Public API
# ══════════════════════════════════════════════════════════════════════════════

def enhance(
    img: Image.Image,
    *,
    preset: str | Preset = Preset.AUTO,
    deskew: bool  = True,
    denoise: bool = True,
    contrast: bool = True,
    binarize: bool = False,
    upscale: bool  = True,
    crop_border: bool = True,
    target_dpi: int   = UPSCALE_TARGET_DPI,
    contrast_factor:  float = CONTRAST_BOOST_FACTOR,
    sharpness_factor: float = SHARPNESS_FACTOR,
    deskew_max_angle: float = DESKEW_MAX_ANGLE,
    return_result: bool = False,
) -> Image.Image | EnhanceResult:
    """
    Apply pre-OCR image enhancement pipeline.

    Args:
        img              : PIL.Image.Image (any mode — converted internally)
        preset           : "auto" | "scan" | "photo" | "mixed" | "fast" | "custom"
                           When not "custom", preset overrides the step flags.
        deskew           : correct rotational skew (requires opencv, skipped if absent)
        denoise          : NlMeans / median-filter denoise pass
        contrast         : auto-level + CLAHE + boost + sharpening
        binarize         : adaptive B&W threshold (best for printed text)
        upscale          : bicubic upscale when source resolution is too low
        crop_border      : detect and remove dark scan borders
        target_dpi       : target DPI for upscale calculation
        contrast_factor  : ImageEnhance.Contrast multiplier (default 1.4)
        sharpness_factor : ImageEnhance.Sharpness multiplier (default 1.5)
        deskew_max_angle : ignore skew detections beyond this many degrees
        return_result    : if True, return EnhanceResult instead of raw Image

    Returns:
        PIL.Image.Image (RGB) — or EnhanceResult if return_result=True
    """
    p = Preset(preset) if isinstance(preset, str) else preset

    # Resolve preset → step flags
    if p == Preset.AUTO:
        p = _auto_detect_preset(img)
        cfg = EnhanceConfig.from_preset(p)
    elif p == Preset.CUSTOM:
        cfg = EnhanceConfig(
            preset=p, deskew=deskew, denoise=denoise, contrast=contrast,
            binarize=binarize, upscale=upscale, crop_border=crop_border,
            target_dpi=target_dpi, contrast_factor=contrast_factor,
            sharpness_factor=sharpness_factor, deskew_max_angle=deskew_max_angle,
        )
    else:
        cfg = EnhanceConfig.from_preset(p)

    result = EnhanceResult(image=_to_rgb(img), preset_used=cfg.preset)
    warnings_list: List[str] = []

    # ── Step 1: Upscale (before all other ops — more pixels = better detection) ──
    if cfg.upscale and _needs_upscale(result.image, cfg.target_dpi):
        result.image, factor = _upscale(result.image, cfg.target_dpi)
        result.was_upscaled  = True
        result.upscale_factor = factor
        result.steps_applied.append(f"upscale×{factor:.2f}")

    # ── Step 2: Border crop ──
    if cfg.crop_border:
        result.image, crop_px = _crop_border(result.image)
        if crop_px:
            result.border_crop = crop_px
            result.steps_applied.append("border_crop")

    # ── Step 3: Deskew ──
    if cfg.deskew:
        result.image, angle = _deskew(result.image, cfg.deskew_max_angle)
        result.skew_angle = angle
        if abs(angle) >= DESKEW_MIN_ANGLE_DEG:
            result.was_deskewed = True
            result.steps_applied.append(f"deskew({angle:.1f}°)")

    # ── Step 4: Denoise ──
    if cfg.denoise:
        result.image, used_nlm = _denoise(result.image)
        result.was_denoised = True
        result.steps_applied.append("denoise_nlm" if used_nlm else "denoise_median")

    # ── Step 5: Contrast / auto-level ──
    if cfg.contrast:
        result.image = _enhance_contrast(
            result.image, cfg.contrast_factor, cfg.sharpness_factor
        )
        result.steps_applied.append("contrast+clahe+sharpen")

    # ── Step 6: Binarize ──
    if cfg.binarize:
        result.image = _binarize(result.image)
        result.was_binarized = True
        result.steps_applied.append("binarize")

    result.warnings = warnings_list

    if return_result:
        return result
    return result.image


def enhance_bytes(
    raw: bytes,
    *,
    ext: str = ".jpg",
    **kwargs,
) -> Image.Image:
    """
    Convenience wrapper: accept raw bytes, return enhanced PIL Image.

    Args:
        raw  : raw image bytes (JPEG, PNG, TIFF, BMP, ...)
        ext  : file extension hint (e.g. ".tiff") — not used by PIL but
               available for future format-specific routing
        **kwargs : passed to enhance()
    """
    img = Image.open(io.BytesIO(raw))
    return enhance(img, **kwargs)


def enhance_file(
    path: str,
    output_path: Optional[str] = None,
    **kwargs,
) -> Image.Image:
    """
    Load image from file path, enhance, optionally save to output_path.

    Args:
        path        : input file path
        output_path : if given, save result here
        **kwargs    : passed to enhance()

    Returns:
        Enhanced PIL Image
    """
    img = Image.open(path)
    result = enhance(img, **kwargs)
    if output_path:
        result.save(output_path)
        logger.info("Saved enhanced image to %s", output_path)
    return result


def batch_enhance(
    images: List[Image.Image],
    *,
    workers: int = 4,
    **kwargs,
) -> List[Image.Image]:
    """
    Enhance a list of PIL Images in parallel using a thread pool.

    Args:
        images  : list of PIL.Image.Image
        workers : max worker threads
        **kwargs : passed to enhance() for every image

    Returns:
        List of enhanced PIL Images, in the same order as input.
    """
    results: Dict[int, Image.Image] = {}

    def _worker(idx: int, img: Image.Image) -> Tuple[int, Image.Image]:
        return idx, enhance(img, **kwargs)

    with ThreadPoolExecutor(max_workers=workers) as pool:
        futures = {pool.submit(_worker, i, img): i for i, img in enumerate(images)}
        for future in as_completed(futures):
            idx, enhanced = future.result()
            results[idx] = enhanced

    return [results[i] for i in range(len(images))]


def describe(img: Image.Image) -> Dict[str, object]:
    """
    Return a diagnostic dict describing image characteristics.
    Useful for debugging preset auto-detection.
    """
    arr = np.array(img.convert("L"), dtype=np.float32)
    dpi = _get_dpi(img)
    return {
        "size":       img.size,
        "mode":       img.mode,
        "dpi":        dpi,
        "mean_brightness": float(arr.mean()),
        "std_brightness":  float(arr.std()),
        "min_brightness":  float(arr.min()),
        "max_brightness":  float(arr.max()),
        "suggested_preset": _auto_detect_preset(img).value,
        "needs_upscale":    _needs_upscale(img),
        "cv2_available":    _cv2_available(),
    }


# ══════════════════════════════════════════════════════════════════════════════
# CLI
# ══════════════════════════════════════════════════════════════════════════════

def _cli() -> None:
    import sys
    import argparse

    parser = argparse.ArgumentParser(
        description="Pre-OCR image enhancement pipeline",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python image_enhancer.py scan.jpg enhanced.jpg
  python image_enhancer.py scan.jpg enhanced.png --preset scan
  python image_enhancer.py photo.jpg out.jpg --preset photo --no-binarize
  python image_enhancer.py doc.tiff out.tiff --describe
        """,
    )
    parser.add_argument("input",  help="Input image path")
    parser.add_argument("output", nargs="?", help="Output image path (optional)")
    parser.add_argument("--preset", default="auto",
                        choices=["auto","scan","photo","mixed","fast","custom"],
                        help="Enhancement preset (default: auto)")
    parser.add_argument("--no-deskew",      dest="deskew",      action="store_false")
    parser.add_argument("--no-denoise",     dest="denoise",     action="store_false")
    parser.add_argument("--no-contrast",    dest="contrast",    action="store_false")
    parser.add_argument("--binarize",       action="store_true",
                        help="Force adaptive binarization")
    parser.add_argument("--no-upscale",     dest="upscale",     action="store_false")
    parser.add_argument("--no-crop-border", dest="crop_border", action="store_false")
    parser.add_argument("--target-dpi",     type=int, default=UPSCALE_TARGET_DPI)
    parser.add_argument("--describe",       action="store_true",
                        help="Print image diagnostics and exit")
    parser.add_argument("--verbose", "-v",  action="store_true")
    parser.set_defaults(deskew=True, denoise=True, contrast=True,
                        upscale=True, crop_border=True)
    args = parser.parse_args()

    if args.verbose:
        logging.basicConfig(level=logging.DEBUG)
    else:
        logging.basicConfig(level=logging.INFO)

    img = Image.open(args.input)

    if args.describe:
        info = describe(img)
        for k, v in info.items():
            print(f"  {k:<22}: {v}")
        return

    preset = Preset.CUSTOM if args.preset == "custom" else Preset(args.preset)
    result: EnhanceResult = enhance(
        img,
        preset=preset,
        deskew=args.deskew,
        denoise=args.denoise,
        contrast=args.contrast,
        binarize=args.binarize,
        upscale=args.upscale,
        crop_border=args.crop_border,
        target_dpi=args.target_dpi,
        return_result=True,
    )

    print(f"Preset used    : {result.preset_used.value}")
    print(f"Steps applied  : {', '.join(result.steps_applied) or 'none'}")
    if result.was_deskewed:
        print(f"Skew corrected : {result.skew_angle:.2f}°")
    if result.was_upscaled:
        print(f"Upscale factor : {result.upscale_factor:.2f}×")
    if result.border_crop:
        print(f"Border removed : l={result.border_crop[0]} t={result.border_crop[1]} "
              f"r={result.border_crop[2]} b={result.border_crop[3]}")
    if result.warnings:
        for w in result.warnings:
            print(f"WARNING: {w}")

    out_path = args.output or (
        os.path.splitext(args.input)[0] + "_enhanced" +
        os.path.splitext(args.input)[1]
    )
    result.image.save(out_path)
    print(f"Saved          : {out_path}")


if __name__ == "__main__":
    _cli()