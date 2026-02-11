import re


_MULTISPACE_RE = re.compile(r"\s+")
_NON_DIGIT_RE = re.compile(r"\D+")
_NON_RG_CHAR_RE = re.compile(r"[^0-9xX]+")


def normalize_upper_text(value: str | None) -> str:
    raw = _MULTISPACE_RE.sub(" ", (value or "").strip())
    return raw.upper()


def normalize_digits(value: str | None) -> str:
    return _NON_DIGIT_RE.sub("", value or "")


def format_cpf(digits: str | None) -> str:
    raw = normalize_digits(digits)[:11]
    if len(raw) <= 3:
        return raw
    if len(raw) <= 6:
        return f"{raw[:3]}.{raw[3:]}"
    if len(raw) <= 9:
        return f"{raw[:3]}.{raw[3:6]}.{raw[6:]}"
    return f"{raw[:3]}.{raw[3:6]}.{raw[6:9]}-{raw[9:]}"


def format_phone(digits: str | None) -> str:
    raw = normalize_digits(digits)[:11]
    if len(raw) <= 2:
        return raw
    if len(raw) <= 6:
        return f"({raw[:2]}) {raw[2:]}"
    if len(raw) <= 10:
        return f"({raw[:2]}) {raw[2:6]}-{raw[6:]}"
    return f"({raw[:2]}) {raw[2:7]}-{raw[7:]}"


def normalize_rg(value: str | None) -> str:
    raw = _NON_RG_CHAR_RE.sub("", value or "").upper()
    if not raw:
        return ""

    trailing_x = raw.endswith("X")
    digits = raw.replace("X", "")
    if trailing_x:
        return f"{digits[:9]}X"
    return digits[:10]


def format_rg(value: str | None) -> str:
    canon = normalize_rg(value)
    if not canon:
        return ""

    if len(canon) == 8:
        base = canon[:-1]
        dv = canon[-1]
        return f"{base[0]}.{base[1:4]}.{base[4:7]}-{dv}"

    if len(canon) == 9:
        base = canon[:-1]
        dv = canon[-1]
        if dv == "X":
            return f"{base[0]}.{base[1:4]}.{base[4:7]}-{dv}"
        return f"{base[0:2]}.{base[2:5]}.{base[5:8]}-{dv}"

    if len(canon) == 10:
        base = canon[:-1]
        dv = canon[-1]
        return f"{base[0:2]}.{base[2:5]}.{base[5:8]}-{dv}"

    return canon


def normalize_oficio_num(value: str | None) -> str:
    raw = (value or "").strip().replace("\\", "/")
    raw = raw.replace(" ", "")
    raw = re.sub(r"[^0-9/]", "", raw)
    if not raw:
        return ""

    if "/" not in raw:
        digits = normalize_digits(raw)
        return str(int(digits)) if digits else ""

    parts = [normalize_digits(part) for part in raw.split("/") if part is not None]
    numero = parts[0] if parts else ""
    ano = parts[1] if len(parts) > 1 else ""
    if not numero and not ano:
        return ""
    numero = str(int(numero)) if numero else ""
    if not ano:
        return numero
    if len(ano) > 4:
        ano = ano[-4:]
    return f"{numero}/{ano}" if numero else ano


def normalize_protocolo_num(value: str | None) -> str:
    return normalize_digits(value)


def format_protocolo_num(value: str | None) -> str:
    digits = normalize_protocolo_num(value)
    if len(digits) != 9:
        return digits
    return f"{digits[0:2]}.{digits[2:5]}.{digits[5:8]}-{digits[8]}"


def split_oficio_num(value: str | None) -> tuple[int | None, int | None]:
    normalized = normalize_oficio_num(value)
    if not normalized:
        return (None, None)
    if "/" in normalized:
        left, right = normalized.split("/", 1)
        numero = int(left) if left.isdigit() else None
        ano = int(right) if right.isdigit() else None
        return (numero, ano)
    if normalized.isdigit():
        return (int(normalized), None)
    return (None, None)


def format_oficio_num(numero: int | str | None, ano: int | str | None) -> str:
    if numero in (None, "") or ano in (None, ""):
        return ""
    try:
        numero_int = int(str(numero))
        ano_int = int(str(ano))
    except (TypeError, ValueError):
        return ""
    if numero_int <= 0 or ano_int <= 0:
        return ""
    return f"{numero_int:02d}/{ano_int}"


def apply_mask_pattern(digits: str | None, pattern: str | None) -> str:
    raw_digits = normalize_digits(digits)
    mask = (pattern or "").strip()
    if not mask:
        return raw_digits
    out: list[str] = []
    idx = 0
    for ch in mask:
        if ch == "0":
            if idx >= len(raw_digits):
                break
            out.append(raw_digits[idx])
            idx += 1
            continue
        if idx == 0 or idx >= len(raw_digits):
            continue
        out.append(ch)
    if idx < len(raw_digits):
        out.append(raw_digits[idx:])
    return "".join(out)
