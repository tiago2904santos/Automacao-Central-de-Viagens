# viagens/services/evento_compilacao.py
"""Compilação do PDF único do protocolo do evento (merge dos assinados)."""
from __future__ import annotations

import io
from pathlib import Path

from django.core.files.base import ContentFile

from viagens.models import DocumentoEventoArquivo, Evento, EventoProtocoloArquivo
from viagens.services.evento_assinados import get_status_assinados_evento, is_evento_pronto_para_compilar


def _image_to_pdf_bytes(image_path_or_file) -> bytes:
    """Converte imagem (JPG/PNG) em PDF de uma página usando Pillow."""
    from PIL import Image
    img = Image.open(image_path_or_file)
    if img.mode in ("RGBA", "P"):
        img = img.convert("RGB")
    buf = io.BytesIO()
    img.save(buf, "PDF")
    buf.seek(0)
    return buf.read()


def _ensure_pdf_bytes(arq: DocumentoEventoArquivo) -> bytes | None:
    """Retorna o conteúdo do arquivo como PDF (bytes). Se for imagem, converte para PDF."""
    if not arq or not arq.arquivo:
        return None
    path = arq.arquivo.path
    if not path or not Path(path).exists():
        # Pode estar em storage remoto; lê pelo arquivo
        arq.arquivo.open("rb")
        try:
            data = arq.arquivo.read()
        finally:
            arq.arquivo.close()
    else:
        with open(path, "rb") as f:
            data = f.read()
    mime = (arq.mime_type or "").lower()
    name = (arq.original_name or "").lower()
    if mime.startswith("image/") or name.endswith(".jpg") or name.endswith(".jpeg") or name.endswith(".png"):
        return _image_to_pdf_bytes(io.BytesIO(data))
    return data


def _ordenar_pdfs_para_protocolo(evento: Evento) -> list[DocumentoEventoArquivo]:
    """
    Ordem: para cada ofício (ofício assinado + justificativa assinada se houver);
    depois plano ou ordem assinado; depois termos (por nome do servidor).
    """
    status = get_status_assinados_evento(evento)
    resultado = []
    # 1) Por ofício: ofício assinado, depois justificativa desse ofício (se existir)
    oficios = status["oficios"]
    justificativas = {j["oficio_id"]: j for j in status["justificativas"]}
    for o in oficios:
        arq_of = _get_arquivo_oficio_assinado(evento, o["oficio_id"])
        if arq_of:
            resultado.append(arq_of)
        j = justificativas.get(o["oficio_id"])
        if j and j.get("arquivo") and j["arquivo"].get("id"):
            arq_j = DocumentoEventoArquivo.objects.filter(pk=j["arquivo"]["id"], is_active=True).first()
            if arq_j:
                resultado.append(arq_j)
    # 2) Plano ou Ordem (um só)
    if status["plano_ou_ordem"].get("arquivo_plano") and status["plano_ou_ordem"]["arquivo_plano"].get("id"):
        arq = DocumentoEventoArquivo.objects.filter(pk=status["plano_ou_ordem"]["arquivo_plano"]["id"], is_active=True).first()
        if arq:
            resultado.append(arq)
    elif status["plano_ou_ordem"].get("arquivo_ordem") and status["plano_ou_ordem"]["arquivo_ordem"].get("id"):
        arq = DocumentoEventoArquivo.objects.filter(pk=status["plano_ou_ordem"]["arquivo_ordem"]["id"], is_active=True).first()
        if arq:
            resultado.append(arq)
    # 3) Termos (já ordenados por nome no status)
    for t in status["termos"]:
        if t.get("arquivo") and t["arquivo"].get("id"):
            arq = DocumentoEventoArquivo.objects.filter(pk=t["arquivo"]["id"], is_active=True).first()
            if arq:
                resultado.append(arq)
    return resultado


def _get_arquivo_oficio_assinado(evento: Evento, oficio_id: int) -> DocumentoEventoArquivo | None:
    return (
        DocumentoEventoArquivo.objects.filter(
            evento=evento,
            tipo=DocumentoEventoArquivo.Tipo.OFICIO_ASSINADO,
            oficio_id=oficio_id,
            is_active=True,
        )
        .order_by("-uploaded_at")
        .first()
    )


def compilar_pdf_protocolo(evento: Evento, compilado_por_id: int | None = None) -> EventoProtocoloArquivo | None:
    """
    Junta todos os PDFs assinados na ordem do protocolo e salva em EventoProtocoloArquivo.
    Retorna a instância criada ou None se não estiver pronto para compilar.
    """
    if not is_evento_pronto_para_compilar(evento):
        return None
    from pypdf import PdfWriter
    ordenados = _ordenar_pdfs_para_protocolo(evento)
    writer = PdfWriter()
    for arq in ordenados:
        pdf_bytes = _ensure_pdf_bytes(arq)
        if pdf_bytes:
            writer.append(io.BytesIO(pdf_bytes))
    buf = io.BytesIO()
    writer.write(buf)
    buf.seek(0)
    # Versão
    ultimo = evento.protocolos_compilados.order_by("-versao").first()
    versao = (ultimo.versao + 1) if ultimo else 1
    nome_arquivo = f"protocolo_evento_{evento.id}_v{versao}.pdf"
    obj = EventoProtocoloArquivo(
        evento=evento,
        compilado_por_id=compilado_por_id,
        versao=versao,
    )
    obj.pdf_compilado.save(nome_arquivo, ContentFile(buf.getvalue()), save=True)
    return obj
