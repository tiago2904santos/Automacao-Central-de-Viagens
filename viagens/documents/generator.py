from __future__ import annotations

import logging

from viagens.documents.document import (
    build_oficio_docx_and_pdf_bytes,
    build_oficio_docx_bytes,
    build_termo_autorizacao_docx_bytes,
    docx_bytes_to_pdf_bytes,
)
from viagens.documents.ordem_servico import build_ordem_servico_docx_bytes
from viagens.documents.plano_trabalho import build_plano_trabalho_docx_bytes
from viagens.models import Oficio

logger = logging.getLogger(__name__)


def _add_pdf_if_available(
    docs: dict[str, bytes],
    *,
    pdf_name: str,
    docx_bytes: bytes,
    oficio: Oficio,
) -> None:
    try:
        pdf_bytes = docx_bytes_to_pdf_bytes(docx_bytes, oficio_id=getattr(oficio, "id", None))
    except Exception:
        logger.exception("[docs-generator] PDF indisponivel para %s", pdf_name)
        return
    docs[pdf_name] = pdf_bytes


def generate_all_documents(oficio: Oficio, *, pdf_if_available: bool = True) -> dict[str, bytes]:
    docs: dict[str, bytes] = {}

    if pdf_if_available:
        try:
            oficio_docx, oficio_pdf = build_oficio_docx_and_pdf_bytes(oficio)
            docs["oficio.docx"] = oficio_docx
            docs["oficio.pdf"] = oficio_pdf
        except Exception:
            logger.exception("[docs-generator] Falha ao gerar PDF do oficio; seguindo com DOCX.")
            docs["oficio.docx"] = build_oficio_docx_bytes(oficio).getvalue()
    else:
        docs["oficio.docx"] = build_oficio_docx_bytes(oficio).getvalue()

    termo_docx = build_termo_autorizacao_docx_bytes(oficio).getvalue()
    docs["termo.docx"] = termo_docx
    if pdf_if_available:
        _add_pdf_if_available(
            docs,
            pdf_name="termo.pdf",
            docx_bytes=termo_docx,
            oficio=oficio,
        )

    has_plano = False
    try:
        has_plano = bool(oficio.plano_trabalho)
    except Exception:
        has_plano = False

    if has_plano:
        plano_docx = build_plano_trabalho_docx_bytes(oficio).getvalue()
        docs["plano_trabalho.docx"] = plano_docx
        if pdf_if_available:
            _add_pdf_if_available(
                docs,
                pdf_name="plano_trabalho.pdf",
                docx_bytes=plano_docx,
                oficio=oficio,
            )
    else:
        ordem_docx = build_ordem_servico_docx_bytes(oficio).getvalue()
        docs["ordem_servico.docx"] = ordem_docx
        if pdf_if_available:
            _add_pdf_if_available(
                docs,
                pdf_name="ordem_servico.pdf",
                docx_bytes=ordem_docx,
                oficio=oficio,
            )

    return docs
