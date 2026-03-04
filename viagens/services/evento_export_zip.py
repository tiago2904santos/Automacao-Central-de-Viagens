# viagens/services/evento_export_zip.py
"""Exportação do pacote do evento em ZIP: compilado + arquivos separados + pasta Prestação de Contas."""
from __future__ import annotations

import io
import re
import zipfile

from django.core.files.base import ContentFile

from viagens.models import DocumentoEventoArquivo, Evento, EventoProtocoloArquivo
from viagens.services.evento_assinados import get_status_assinados_evento, is_evento_pronto_para_compilar
from viagens.services.evento_compilacao import compilar_pdf_protocolo, _ensure_pdf_bytes


def _sanitize_filename(s: str, max_len: int = 80) -> str:
    """Remove caracteres inválidos para nome de arquivo; limita tamanho."""
    if not s:
        return "evento"
    s = str(s).strip().replace("/", "_").replace("\\", "_")
    s = re.sub(r"[^\w\s\-]", "", s, flags=re.ASCII)
    s = re.sub(r"\s+", "_", s).strip("_")
    return (s[:max_len] if len(s) > max_len else s) or "evento"


def _read_file_bytes(arq: DocumentoEventoArquivo | EventoProtocoloArquivo, file_attr: str = "arquivo") -> bytes | None:
    """Lê o conteúdo do FileField; para DocumentoEventoArquivo pode precisar converter imagem para PDF."""
    if arq is None:
        return None
    field = getattr(arq, file_attr, None)
    if not field:
        return None
    try:
        field.open("rb")
        data = field.read()
        field.close()
    except Exception:
        return None
    if hasattr(arq, "mime_type") and (arq.mime_type or "").lower().startswith("image/"):
        from viagens.services.evento_compilacao import _image_to_pdf_bytes
        try:
            return _image_to_pdf_bytes(io.BytesIO(data))
        except Exception:
            return data
    return data


def build_evento_zip(evento: Evento, compilado_por_id: int | None = None) -> tuple[bytes, str] | None:
    """
    Monta o ZIP do pacote do evento. Retorna (zip_bytes, nome_arquivo) ou None se não estiver pronto.
    - 01_PACOTE_COMPILADO/Protocolo_Compilado.pdf
    - 02_ARQUIVOS_SEPARADOS/01_Oficios_Assinados/, 02_Solicitacao_ou_PT_OS/, 03_Justificativas/, 04_Termos/
    - 03_PRESTACAO_DE_CONTAS/ (vazio)
    """
    if not is_evento_pronto_para_compilar(evento):
        return None
    buf = io.BytesIO()
    zf = zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED)

    # 1) PDF compilado
    ultimo = evento.protocolos_compilados.order_by("-versao").first()
    if (not ultimo or not ultimo.pdf_compilado) and compilado_por_id is not None:
        try:
            compilar_pdf_protocolo(evento, compilado_por_id=compilado_por_id)
            ultimo = evento.protocolos_compilados.order_by("-versao").first()
        except Exception:
            ultimo = None
    if ultimo and ultimo.pdf_compilado:
        try:
            ultimo.pdf_compilado.open("rb")
            data = ultimo.pdf_compilado.read()
            ultimo.pdf_compilado.close()
        except Exception:
            data = None
        if data:
            zf.writestr("01_PACOTE_COMPILADO/Protocolo_Compilado.pdf", data)
    else:
        zf.writestr(
            "01_PACOTE_COMPILADO/LEIA-ME.txt",
            "Gere o PDF compilado no Pacote do Evento (botão Gerar PDF do protocolo).".encode("utf-8"),
        )

    # 2) Arquivos separados
    status = get_status_assinados_evento(evento)
    oficios = list(evento.oficios.order_by("id").prefetch_related("termos_autorizacao", "viajantes"))

    # 2.1 Ofícios assinados
    for o in status["oficios"]:
        if not o.get("assinado") or not o.get("arquivo", {}).get("id"):
            continue
        arq = DocumentoEventoArquivo.objects.filter(pk=o["arquivo"]["id"], is_active=True).first()
        if arq and arq.arquivo:
            data = _ensure_pdf_bytes(arq) if hasattr(arq, "mime_type") else _read_file_bytes(arq)
            if data:
                num = _sanitize_filename(str(o.get("numero_display") or o["oficio_id"]), 40)
                zf.writestr(f"02_ARQUIVOS_SEPARADOS/01_Oficios_Assinados/Oficio_{num}.pdf", data)

    # 2.2 Solicitação formal ou Plano/Ordem
    if status.get("solicitacao_formal", {}).get("assinado") and status["solicitacao_formal"].get("arquivo", {}).get("id"):
        arq = DocumentoEventoArquivo.objects.filter(pk=status["solicitacao_formal"]["arquivo"]["id"], is_active=True).first()
        if arq and arq.arquivo:
            data = _ensure_pdf_bytes(arq)
            if data:
                zf.writestr("02_ARQUIVOS_SEPARADOS/02_Solicitacao_ou_PT_OS/Solicitacao_Formal.pdf", data)
    for key, tipo, prefix in [
        ("arquivo_plano", DocumentoEventoArquivo.Tipo.PLANO_ASSINADO, "Plano_Trabalho"),
        ("arquivo_ordem", DocumentoEventoArquivo.Tipo.ORDEM_ASSINADO, "Ordem_Servico"),
    ]:
        arq_info = status["plano_ou_ordem"].get(key)
        if arq_info and arq_info.get("id"):
            arq = DocumentoEventoArquivo.objects.filter(pk=arq_info["id"], is_active=True).first()
            if arq and arq.arquivo:
                data = _ensure_pdf_bytes(arq)
                if data:
                    zf.writestr(f"02_ARQUIVOS_SEPARADOS/02_Solicitacao_ou_PT_OS/{prefix}.pdf", data)

    # 2.3 Justificativas
    for j in status["justificativas"]:
        if not j.get("assinado") or not j.get("arquivo", {}).get("id"):
            continue
        arq = DocumentoEventoArquivo.objects.filter(pk=j["arquivo"]["id"], is_active=True).first()
        if arq and arq.arquivo:
            data = _ensure_pdf_bytes(arq)
            if data:
                num_j = _sanitize_filename(str(j.get("numero_display") or j["oficio_id"]), 40)
                zf.writestr(f"02_ARQUIVOS_SEPARADOS/03_Justificativas/Justificativa_Oficio_{num_j}.pdf", data)

    # 2.4 Termos (não dispensados)
    termos_oficio = {}
    for o in oficios:
        termos_oficio[o.id] = {
            t.viajante_id: t
            for t in o.termos_autorizacao.select_related("viajante").all()
            if t.viajante_id and not t.dispensado
        }
    for t in status["termos"]:
        if t.get("dispensado"):
            continue
        if not t.get("assinado") or not t.get("arquivo", {}).get("id"):
            continue
        arq = DocumentoEventoArquivo.objects.filter(pk=t["arquivo"]["id"], is_active=True).first()
        if arq and arq.arquivo:
            data = _ensure_pdf_bytes(arq)
            if data:
                oficio_id = t.get("oficio_id")
                oficio_obj = next((x for x in oficios if x.id == oficio_id), None)
                num_of = _sanitize_filename((oficio_obj.numero_formatado if oficio_obj else str(oficio_id)), 30)
                nome_serv = _sanitize_filename((t.get("nome") or "servidor")[:50])
                zf.writestr(f"02_ARQUIVOS_SEPARADOS/04_Termos/Termo_{num_of}_{nome_serv}.pdf", data)

    # 3) Pasta Prestação de Contas (vazia)
    zf.writestr("03_PRESTACAO_DE_CONTAS/.gitkeep", b"")

    zf.close()
    buf.seek(0)
    titulo = _sanitize_filename(evento.titulo or "")
    nome_zip = f"EVENTO_{evento.id}_{titulo}.zip"
    return buf.getvalue(), nome_zip
