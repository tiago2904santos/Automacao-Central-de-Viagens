# viagens/services/documentos_manager.py
"""Status e metadados dos documentos por ofício para a Central de Documentos e drawer."""
from __future__ import annotations

from viagens.models import Oficio, OrdemServico, PlanoTrabalho
from viagens.services.justificativa_helpers import (
    exige_justificativa,
    justificativa_preenchida,
)


def _get_plano(oficio: Oficio) -> PlanoTrabalho | None:
    acao = getattr(oficio, "acao", None)
    if acao is None and hasattr(oficio, "ensure_acao"):
        try:
            acao = oficio.ensure_acao()
        except Exception:
            acao = None
    if acao is not None:
        try:
            return acao.plano_trabalho
        except PlanoTrabalho.DoesNotExist:
            pass
    try:
        return oficio.plano_trabalho
    except PlanoTrabalho.DoesNotExist:
        return None


def _get_ordem(oficio: Oficio) -> OrdemServico | None:
    acao = getattr(oficio, "acao", None)
    if acao is None and hasattr(oficio, "ensure_acao"):
        try:
            acao = oficio.ensure_acao()
        except Exception:
            acao = None
    if acao is not None:
        try:
            return acao.ordem_servico
        except OrdemServico.DoesNotExist:
            pass
    try:
        return oficio.ordem_servico
    except OrdemServico.DoesNotExist:
        return None


def build_documentos_status(oficio: Oficio) -> dict:
    """
    Retorna um dicionário com status, flags e dados para cada tipo de documento
    do ofício. Usado pelo partial de cards e pela Central de Documentos.
    """
    plano = _get_plano(oficio)
    ordem = _get_ordem(oficio)
    tem_plano = plano is not None
    tem_ordem = ordem is not None

    justificativa_exige = exige_justificativa(oficio)
    justificativa_ok = justificativa_preenchida(oficio)
    if justificativa_exige and not justificativa_ok:
        justificativa_status = "pendente"
    else:
        justificativa_status = "ok" if justificativa_ok else "nao_exigido"

    preview_justificativa = ""
    if oficio.justificativa_texto:
        raw = (oficio.justificativa_texto or "").strip()
        preview_justificativa = raw[:200] + ("..." if len(raw) > 200 else "")

    return {
        "oficio": {
            "status": "ok",
            "can_generate": True,
        },
        "termo": {
            "status": "ok",
            "can_generate": True,
        },
        "justificativa": {
            "status": justificativa_status,
            "exige": justificativa_exige,
            "preview": preview_justificativa,
            "can_generate_doc": justificativa_ok,
        },
        "plano": {
            "status": "ok" if tem_plano else "nao_cadastrado",
            "numero": getattr(plano, "numero", None) if plano else None,
            "ano": getattr(plano, "ano", None) if plano else None,
        },
        "ordem": {
            "visible": not tem_plano,
            "status": "ok" if tem_ordem else "nao_cadastrado",
            "numero": getattr(ordem, "numero", None) if ordem else None,
            "ano": getattr(ordem, "ano", None) if ordem else None,
        },
        "outros": _build_outros_links(oficio),
    }


def _build_outros_links(oficio: Oficio) -> list[dict]:
    """Links para outros documentos (listas, anexos, etc.)."""
    from django.urls import reverse

    links = []
    try:
        links.append({
            "label": "Planos de trabalho",
            "url": reverse("planos_trabalho_list"),
        })
        links.append({
            "label": "Ordens de serviço",
            "url": reverse("ordens_servico_list"),
        })
        links.append({
            "label": "Termos de autorização",
            "url": reverse("termos_autorizacao_lista"),
        })
    except Exception:
        pass
    return links
