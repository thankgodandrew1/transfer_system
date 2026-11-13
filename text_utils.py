"""Comparison keys that retain names outside the ASCII alphabet."""
import unicodedata


def comparison_key(value: object) -> str:
    text = unicodedata.normalize("NFKD", str(value or "").upper())
    return "".join(char for char in text if char.isalnum())
