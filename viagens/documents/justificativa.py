# viagens/documents/justificativa.py
"""Geração de documento de Justificativa a partir do modelo DOCX com placeholders {{chave}}."""
from __future__ import annotations

from io import BytesIO
from pathlib import Path

from django.conf import settings

from .document import (
    _extract_placeholders,
    _iter_docx_xml_parts_from_path,
    _sanitize_mapping_values,
    safe_replace_placeholders,
)
from docx import Document as DocxFactory


def get_justificativa_template_path() -> Path:
    """Retorna o caminho do modelo modelo_Justificativa.docx."""
    base_dir = Path(settings.BASE_DIR) / "viagens" / "documents"
    path = base_dir / "modelo_Justificativa.docx"
    if not path.exists():
        raise FileNotFoundError(
            f"Modelo de justificativa não encontrado: {path}"
        )
    return path


def get_justificativa_placeholders() -> list[str]:
    """Extrai do modelo DOCX a lista de placeholders (ordem alfabética)."""
    template_path = get_justificativa_template_path()
    parts = _iter_docx_xml_parts_from_path(str(template_path))
    keys = _extract_placeholders(parts)
    return sorted(keys)


def build_justificativa_docx_bytes(mapping: dict[str, str]) -> BytesIO:
    """
    Gera o DOCX da justificativa substituindo os placeholders do modelo.
    mapping: dicionário chave -> valor (chave = nome do placeholder sem chaves).
    Placeholders não informados são preenchidos com string vazia.
    """
    template_path = get_justificativa_template_path()
    doc = DocxFactory(str(template_path))

    template_placeholders = set(get_justificativa_placeholders())
    for key in template_placeholders:
        if key not in mapping:
            mapping[key] = ""
    mapping = _sanitize_mapping_values(mapping)
    safe_replace_placeholders(doc, mapping)

    buf = BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf
