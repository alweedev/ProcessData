import re
import unicodedata


def normalize_text(value):
    if value is None:
        return ""
    return str(value).strip()


def upper_no_accents(value):
    text = normalize_text(value).upper()
    return unicodedata.normalize("NFKD", text).encode("ASCII", "ignore").decode("utf-8")


def sanitize_output_text(value, maxlen=None):
    if value is None:
        return ""
    text = upper_no_accents(value)
    text = re.sub(r"[^A-Z0-9 \-/()]", "", text)
    text = re.sub(r"\s+", " ", text).strip()
    if maxlen:
        return text[:maxlen]
    return text


def split_name_first_last(fullname):
    if not fullname:
        return "", ""
    parts = [p for p in str(fullname).strip().split() if p]
    if not parts:
        return "", ""
    if len(parts) == 1:
        first, last = parts[0], ""
    else:
        first, last = parts[0], parts[-1]
    return sanitize_output_text(first, 20), sanitize_output_text(last, 20)
