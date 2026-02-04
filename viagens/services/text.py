def title_case_pt(texto: str) -> str:
    if not texto:
        return ""
    base = " ".join(texto.strip().split())
    if not base:
        return ""
    titled = base.title()
    parts = titled.split()
    particles = {"da", "de", "do", "dos", "das", "e"}
    normalized: list[str] = []
    for idx, part in enumerate(parts):
        lowered = part.lower()
        if idx > 0 and lowered in particles:
            normalized.append(lowered)
        else:
            normalized.append(part)
    return " ".join(normalized)
