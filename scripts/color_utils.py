"""颜色工具：hex/RGB/Lab 转换、CIEDE2000 色差、WCAG 对比度、最近主题色匹配。

不依赖 colormath（其在新版 numpy 上有兼容问题），自己实现以保证稳定。
"""

from __future__ import annotations

import math
import re
from typing import Iterable, Optional

HEX_RE = re.compile(r"^#?([0-9A-Fa-f]{6})$")


# ---------- 基础转换 ----------

def hex_to_rgb(hex_str: str) -> tuple[int, int, int]:
    """'RRGGBB' 或 '#RRGGBB' -> (r, g, b)."""
    m = HEX_RE.match(hex_str.strip())
    if not m:
        raise ValueError(f"invalid hex color: {hex_str!r}")
    h = m.group(1)
    return int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)


def rgb_to_hex(rgb: tuple[int, int, int]) -> str:
    r, g, b = (max(0, min(255, int(round(c)))) for c in rgb)
    return f"{r:02X}{g:02X}{b:02X}"


def normalize_hex(hex_str: str) -> str:
    return rgb_to_hex(hex_to_rgb(hex_str))


def _srgb_to_linear(c: float) -> float:
    c = c / 255.0
    return c / 12.92 if c <= 0.04045 else ((c + 0.055) / 1.055) ** 2.4


def rgb_to_xyz(rgb: tuple[int, int, int]) -> tuple[float, float, float]:
    r, g, b = (_srgb_to_linear(v) for v in rgb)
    # sRGB D65
    x = r * 0.4124564 + g * 0.3575761 + b * 0.1804375
    y = r * 0.2126729 + g * 0.7151522 + b * 0.0721750
    z = r * 0.0193339 + g * 0.1191920 + b * 0.9503041
    return x * 100, y * 100, z * 100


_REF_X, _REF_Y, _REF_Z = 95.047, 100.000, 108.883  # D65


def xyz_to_lab(xyz: tuple[float, float, float]) -> tuple[float, float, float]:
    def f(t: float) -> float:
        return t ** (1 / 3) if t > 0.008856 else (7.787 * t + 16 / 116)

    x, y, z = xyz
    fx = f(x / _REF_X)
    fy = f(y / _REF_Y)
    fz = f(z / _REF_Z)
    L = 116 * fy - 16
    a = 500 * (fx - fy)
    b = 200 * (fy - fz)
    return L, a, b


def hex_to_lab(hex_str: str) -> tuple[float, float, float]:
    return xyz_to_lab(rgb_to_xyz(hex_to_rgb(hex_str)))


# ---------- CIEDE2000 ----------

def ciede2000(lab1: tuple[float, float, float], lab2: tuple[float, float, float]) -> float:
    """标准 CIEDE2000 实现，输入 Lab 三元组，返回 ΔE。"""
    L1, a1, b1 = lab1
    L2, a2, b2 = lab2

    avg_L = (L1 + L2) / 2
    C1 = math.hypot(a1, b1)
    C2 = math.hypot(a2, b2)
    avg_C = (C1 + C2) / 2

    G = 0.5 * (1 - math.sqrt(avg_C ** 7 / (avg_C ** 7 + 25 ** 7)))
    a1p = (1 + G) * a1
    a2p = (1 + G) * a2
    C1p = math.hypot(a1p, b1)
    C2p = math.hypot(a2p, b2)
    avg_Cp = (C1p + C2p) / 2

    def _hp(ap: float, bp: float) -> float:
        if ap == 0 and bp == 0:
            return 0.0
        h = math.degrees(math.atan2(bp, ap))
        return h + 360 if h < 0 else h

    h1p = _hp(a1p, b1)
    h2p = _hp(a2p, b2)

    if abs(h1p - h2p) > 180:
        avg_Hp = (h1p + h2p + 360) / 2
    else:
        avg_Hp = (h1p + h2p) / 2

    T = (
        1
        - 0.17 * math.cos(math.radians(avg_Hp - 30))
        + 0.24 * math.cos(math.radians(2 * avg_Hp))
        + 0.32 * math.cos(math.radians(3 * avg_Hp + 6))
        - 0.20 * math.cos(math.radians(4 * avg_Hp - 63))
    )

    delta_hp = h2p - h1p
    if abs(delta_hp) > 180:
        delta_hp = delta_hp - 360 if h2p > h1p else delta_hp + 360

    delta_Lp = L2 - L1
    delta_Cp = C2p - C1p
    delta_Hp = 2 * math.sqrt(C1p * C2p) * math.sin(math.radians(delta_hp / 2))

    SL = 1 + (0.015 * (avg_L - 50) ** 2) / math.sqrt(20 + (avg_L - 50) ** 2)
    SC = 1 + 0.045 * avg_Cp
    SH = 1 + 0.015 * avg_Cp * T

    delta_theta = 30 * math.exp(-(((avg_Hp - 275) / 25) ** 2))
    Rc = 2 * math.sqrt(avg_Cp ** 7 / (avg_Cp ** 7 + 25 ** 7))
    RT = -math.sin(math.radians(2 * delta_theta)) * Rc

    return math.sqrt(
        (delta_Lp / SL) ** 2
        + (delta_Cp / SC) ** 2
        + (delta_Hp / SH) ** 2
        + RT * (delta_Cp / SC) * (delta_Hp / SH)
    )


def hex_delta_e(hex_a: str, hex_b: str) -> float:
    return ciede2000(hex_to_lab(hex_a), hex_to_lab(hex_b))


# ---------- 最近主题色匹配 ----------

def nearest_theme_color(
    hex_color: str,
    theme_palette: dict[str, str],
) -> tuple[str, float]:
    """在主题色字典中找最接近的槽位。

    theme_palette: {"dk1": "RRGGBB", "accent1": "RRGGBB", ...}
    返回: (slot_name, delta_e)
    """
    target_lab = hex_to_lab(hex_color)
    best_slot, best_de = "", float("inf")
    for slot, hx in theme_palette.items():
        try:
            de = ciede2000(target_lab, hex_to_lab(hx))
        except ValueError:
            continue
        if de < best_de:
            best_de = de
            best_slot = slot
    return best_slot, best_de


# ---------- WCAG 对比度 ----------

def relative_luminance(rgb: tuple[int, int, int]) -> float:
    r, g, b = (_srgb_to_linear(v) for v in rgb)
    return 0.2126 * r + 0.7152 * g + 0.0722 * b


def wcag_contrast(hex_a: str, hex_b: str) -> float:
    la = relative_luminance(hex_to_rgb(hex_a))
    lb = relative_luminance(hex_to_rgb(hex_b))
    lighter, darker = max(la, lb), min(la, lb)
    return (lighter + 0.05) / (darker + 0.05)


# ---------- 明度变体（用于候选色不足时派生） ----------

def derive_variant(hex_color: str, lightness_delta: float) -> str:
    """在 Lab 空间调整 L 通道，正值变亮、负值变暗。"""
    L, a, b = hex_to_lab(hex_color)
    new_L = max(0.0, min(100.0, L + lightness_delta))
    # Lab -> XYZ -> RGB
    fy = (new_L + 16) / 116
    fx = a / 500 + fy
    fz = fy - b / 200

    def f_inv(t: float) -> float:
        t3 = t ** 3
        return t3 if t3 > 0.008856 else (t - 16 / 116) / 7.787

    x = _REF_X * f_inv(fx)
    y = _REF_Y * f_inv(fy)
    z = _REF_Z * f_inv(fz)

    x, y, z = x / 100, y / 100, z / 100
    r = x * 3.2404542 + y * -1.5371385 + z * -0.4985314
    g = x * -0.9692660 + y * 1.8760108 + z * 0.0415560
    bl = x * 0.0556434 + y * -0.2040259 + z * 1.0572252

    def linear_to_srgb(c: float) -> float:
        c = max(0.0, min(1.0, c))
        return 12.92 * c if c <= 0.0031308 else 1.055 * (c ** (1 / 2.4)) - 0.055

    return rgb_to_hex(
        (round(linear_to_srgb(r) * 255),
         round(linear_to_srgb(g) * 255),
         round(linear_to_srgb(bl) * 255))
    )


if __name__ == "__main__":
    # 自检
    assert normalize_hex("#ff0000") == "FF0000"
    assert normalize_hex("00ff00") == "00FF00"
    assert hex_delta_e("FFFFFF", "FFFFFF") < 0.01
    assert hex_delta_e("FF0000", "00FF00") > 50
    assert wcag_contrast("000000", "FFFFFF") > 20
    assert wcag_contrast("FFFFFF", "FFFFFF") < 1.01
    print("color_utils self-check OK")
    print("FF0000 vs F50000 delta_e =", round(hex_delta_e("FF0000", "F50000"), 2))
    print("derive_variant(4F81BD, +20) =", derive_variant("4F81BD", 20))
    print("derive_variant(4F81BD, -20) =", derive_variant("4F81BD", -20))
