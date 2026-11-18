"""扫描件增强：找纸张四角、透视矫正、去阴影。"""

from __future__ import annotations

import logging
from pathlib import Path
from typing import Callable

import cv2
import numpy as np
from PIL import Image, ImageOps

from converter import ConversionError

LOGGER = logging.getLogger("docworkshop.scanner")

MAX_EDGE = 1600
MIN_QUAD_AREA_RATIO = 0.12
CANNY_LOW = 50
CANNY_HIGH = 150
APPROX_EPSILON = 0.02
TRIM_RATIO = 0.01
CLAHE_CLIP = 2.0
CLAHE_GRID = (8, 8)

ProgressFn = Callable[[int, str], None]


def _load_image(src: Path) -> tuple[np.ndarray, str]:
    try:
        with Image.open(src) as opened:
            opened.seek(0)
            frames = getattr(opened, "n_frames", 1)
            image = ImageOps.exif_transpose(opened)
            if image.mode in {"RGBA", "LA"}:
                background = Image.new("RGB", image.size, "white")
                alpha = image.getchannel("A")
                background.paste(image.convert("RGB"), mask=alpha)
                image = background
            else:
                image = image.convert("RGB")
            array = np.asarray(image, dtype=np.uint8)[:, :, ::-1].copy()
    except Exception as exc:
        raise ConversionError("无法读取图片，请确认文件没有损坏") from exc
    note = "仅处理了第一帧" if frames > 1 else ""
    return array, note


def _resize(image: np.ndarray) -> np.ndarray:
    height, width = image.shape[:2]
    longest = max(height, width)
    if longest <= MAX_EDGE:
        return image
    scale = MAX_EDGE / float(longest)
    return cv2.resize(
        image,
        (max(1, int(width * scale)), max(1, int(height * scale))),
        interpolation=cv2.INTER_AREA,
    )


def _order_points(points: np.ndarray) -> np.ndarray:
    ordered = np.zeros((4, 2), dtype="float32")
    sums = points.sum(axis=1)
    diffs = np.diff(points, axis=1).reshape(-1)
    ordered[0] = points[int(np.argmin(sums))]
    ordered[2] = points[int(np.argmax(sums))]
    ordered[1] = points[int(np.argmin(diffs))]
    ordered[3] = points[int(np.argmax(diffs))]
    return ordered


def _find_page_quad(image: np.ndarray) -> np.ndarray | None:
    gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
    blurred = cv2.GaussianBlur(gray, (5, 5), 0)
    edges = cv2.Canny(blurred, CANNY_LOW, CANNY_HIGH)
    kernel = np.ones((3, 3), np.uint8)
    edges = cv2.dilate(edges, kernel, iterations=1)
    contours, _ = cv2.findContours(edges, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    total_area = float(image.shape[0] * image.shape[1])
    best: np.ndarray | None = None
    best_area = 0.0
    for contour in sorted(contours, key=cv2.contourArea, reverse=True)[:12]:
        area = float(cv2.contourArea(contour))
        if area < total_area * MIN_QUAD_AREA_RATIO or area > total_area * 0.985:
            continue
        perimeter = cv2.arcLength(contour, True)
        approx = cv2.approxPolyDP(contour, APPROX_EPSILON * perimeter, True)
        if len(approx) != 4 or not cv2.isContourConvex(approx):
            continue
        if area > best_area:
            best = approx.reshape(4, 2).astype("float32")
            best_area = area
    return best


def _warp(image: np.ndarray, quad: np.ndarray) -> np.ndarray:
    top_left, top_right, bottom_right, bottom_left = _order_points(quad)
    width = int(
        max(
            np.linalg.norm(bottom_right - bottom_left),
            np.linalg.norm(top_right - top_left),
        )
    )
    height = int(
        max(
            np.linalg.norm(top_right - bottom_right),
            np.linalg.norm(top_left - bottom_left),
        )
    )
    width = max(10, width)
    height = max(10, height)
    destination = np.array(
        [[0, 0], [width - 1, 0], [width - 1, height - 1], [0, height - 1]],
        dtype="float32",
    )
    matrix = cv2.getPerspectiveTransform(
        np.array([top_left, top_right, bottom_right, bottom_left], dtype="float32"),
        destination,
    )
    return cv2.warpPerspective(image, matrix, (width, height))


def _trim(image: np.ndarray) -> np.ndarray:
    height, width = image.shape[:2]
    margin_x = int(width * TRIM_RATIO)
    margin_y = int(height * TRIM_RATIO)
    if width - margin_x * 2 < 10 or height - margin_y * 2 < 10:
        return image
    return image[margin_y : height - margin_y, margin_x : width - margin_x]


def _remove_shadow(image: np.ndarray) -> np.ndarray:
    lab = cv2.cvtColor(image, cv2.COLOR_BGR2LAB)
    lightness, channel_a, channel_b = cv2.split(lab)
    height, width = lightness.shape[:2]
    kernel_size = max(31, (min(height, width) // 14) | 1)
    background = cv2.GaussianBlur(lightness, (kernel_size, kernel_size), 0)
    normalized = cv2.divide(lightness, background, scale=255)
    clahe = cv2.createCLAHE(clipLimit=CLAHE_CLIP, tileGridSize=CLAHE_GRID)
    enhanced = clahe.apply(normalized)
    merged = cv2.merge([enhanced, channel_a, channel_b])
    return cv2.cvtColor(merged, cv2.COLOR_LAB2BGR)


def _save_png(image: np.ndarray, dst: Path) -> None:
    success, buffer = cv2.imencode(".png", image, [cv2.IMWRITE_PNG_COMPRESSION, 6])
    if not success:
        raise ConversionError("无法生成 PNG 结果")
    buffer.tofile(str(dst))


def _save_original(image: np.ndarray, dst: Path) -> None:
    _save_png(image, dst)


def enhance(src: Path, output_dir: Path, progress: ProgressFn) -> tuple[Path, str]:
    progress(5, "正在读取图片")
    image, frame_note = _load_image(src)
    output_dir.mkdir(parents=True, exist_ok=True)
    dst = output_dir / "scan.png"
    try:
        progress(20, "正在缩放图片")
        resized = _resize(image)
        progress(40, "正在寻找纸张边界")
        quad = _find_page_quad(resized)
        if quad is None:
            progress(60, "未找到纸张边界，正在按原图增强")
            note = "未检测到纸张边界，已按原图增强"
            flattened = _trim(resized)
        else:
            progress(60, "正在透视矫正")
            note = "已检测到纸张边界并完成透视矫正"
            flattened = _warp(resized, quad)
        progress(80, "正在去阴影并增强对比度")
        cleaned = _remove_shadow(flattened)
        _save_png(cleaned, dst)
        progress(95, "正在保存扫描件")
        if frame_note:
            note = f"{note}；{frame_note}"
        return dst, f"{note}；已去阴影并增强对比度"
    except Exception as exc:
        LOGGER.warning("扫描件增强失败，返回原图：%s", exc, exc_info=True)
        try:
            _save_original(image, dst)
            return dst, f"扫描件增强失败，已返回原图：{exc}"
        except Exception as fallback_exc:
            LOGGER.exception("返回原图也失败了")
            raise ConversionError("扫描件增强失败，且无法保存原图") from fallback_exc
