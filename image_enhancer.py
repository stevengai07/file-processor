import cv2
import numpy as np
from dataclasses import dataclass
from typing import Optional

@dataclass
class EnhancerConfig:
    deskew: bool = True
    denoise: bool = True
    contrast: bool = True
    binarize: bool = True
    remove_borders: bool = True
    sharpen: bool = False
    morphology: bool = False
    upscale: bool = False
    adaptive_threshold: bool = True
    target_dpi: int = 300
    source_dpi: int = 150

PRESET_SCANNER      = EnhancerConfig(adaptive_threshold=False, sharpen=False, morphology=False, upscale=False)
PRESET_PHONE_CAMERA = EnhancerConfig(adaptive_threshold=True, sharpen=True, morphology=False, upscale=True, source_dpi=96, target_dpi=300)
PRESET_FADED        = EnhancerConfig(adaptive_threshold=True, sharpen=True, morphology=True, contrast=True, denoise=True, upscale=True, source_dpi=150)
PRESET_FAX          = EnhancerConfig(adaptive_threshold=False, sharpen=False, morphology=True, denoise=True)

PRESETS = {
    "scanner":      PRESET_SCANNER,
    "phone_camera": PRESET_PHONE_CAMERA,
    "faded":        PRESET_FADED,
    "fax":          PRESET_FAX,
}

def upscale_to_dpi(img, config):
    if not config.upscale or config.source_dpi >= config.target_dpi:
        return img
    scale = config.target_dpi / config.source_dpi
    h, w = img.shape[:2]
    return cv2.resize(img, (int(w * scale), int(h * scale)), interpolation=cv2.INTER_LANCZOS4)

def to_grayscale(img):
    return cv2.cvtColor(img, cv2.COLOR_BGR2GRAY) if len(img.shape) == 3 else img
def enhance_contrast(img):
    clahe = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8, 8))
    return clahe.apply(img)

def denoise(img):
    return cv2.fastNlMeansDenoising(img, h=15, templateWindowSize=7, searchWindowSize=21)

def deskew(img):
    edges = cv2.Canny(img, 50, 150, apertureSize=3)
    lines = cv2.HoughLines(edges, 1, np.pi / 180, threshold=200)
    if lines is None:
        return img
    angles = [theta - np.pi / 2 for rho, theta in lines[:, 0]
              if np.pi / 4 < theta < 3 * np.pi / 4]
    if not angles:
        return img
    angle = np.median(angles) * (180 / np.pi)
    if abs(angle) < 0.3:
        return img
    h, w = img.shape
    M = cv2.getRotationMatrix2D((w // 2, h // 2), angle, 1.0)
    return cv2.warpAffine(img, M, (w, h), flags=cv2.INTER_CUBIC,
                          borderMode=cv2.BORDER_REPLICATE)

def sharpen(img):
    blurred = cv2.GaussianBlur(img, (0, 0), sigmaX=2)
    return cv2.addWeighted(img, 1.5, blurred, -0.5, 0)

def binarize(img, adaptive=True):
    if adaptive:
        return cv2.adaptiveThreshold(img, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
                                     cv2.THRESH_BINARY, 31, 15)
    _, binary = cv2.threshold(img, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
    return binary

def apply_morphology(img):
    kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (2, 2))
    opened = cv2.morphologyEx(img, cv2.MORPH_OPEN, kernel, iterations=1)
    return cv2.morphologyEx(opened, cv2.MORPH_CLOSE, kernel, iterations=1)

def remove_borders(img):
    mask = np.zeros((img.shape[0] + 2, img.shape[1] + 2), np.uint8)
    flood = img.copy()
    for y, x in [(0,0),(0,img.shape[1]-1),(img.shape[0]-1,0),(img.shape[0]-1,img.shape[1]-1)]:
        cv2.floodFill(flood, mask, (x, y), 255)
    return cv2.bitwise_and(img, cv2.bitwise_not(flood))

def enhance_document_image(img: np.ndarray, config: Optional[EnhancerConfig] = None) -> np.ndarray:
    if config is None:
        config = EnhancerConfig()
    img = upscale_to_dpi(img, config)
    img = to_grayscale(img)
    if config.contrast:       img = enhance_contrast(img)
    if config.denoise:        img = denoise(img)
    if config.deskew:         img = deskew(img)
    if config.sharpen:        img = sharpen(img)
    if config.binarize:       img = binarize(img, adaptive=config.adaptive_threshold)
    if config.morphology:     img = apply_morphology(img)
    if config.remove_borders: img = remove_borders(img)
    return img