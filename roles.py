"""Extensible assignment labels shared by forms, corrections and publications."""
import colorsys
import hashlib
import re
import unicodedata


def normalize_role(value: str) -> str:
    role = unicodedata.normalize("NFKC", str(value)).strip().upper()
    role = re.sub(r"\s*/\s*", "/", role)
    role = re.sub(r"\s+", " ", role)
    if not role or len(role) > 40 or not re.fullmatch(r"[^\W_][\w /&+().-]*", role, re.UNICODE):
        raise ValueError("Use a role of 1–40 letters or numbers, with spaces or / & + ( ) . - for combinations.")
    if any(not part.strip() for part in role.split("/")):
        raise ValueError("Each role in a / combination must have a name.")
    return role


def custom_role_colors(role: str) -> tuple[str, str]:
    """Stable colours across processes and transfers, with dark readable text."""
    hue = int.from_bytes(hashlib.sha256(role.encode("utf-8")).digest()[:4], "big") / 2**32
    def hex_color(lightness: float, saturation: float) -> str:
        return "".join(f"{round(c * 255):02X}" for c in colorsys.hls_to_rgb(hue, lightness, saturation))
    return hex_color(.93, .55), hex_color(.26, .65)
