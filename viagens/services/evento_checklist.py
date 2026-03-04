# viagens/services/evento_checklist.py
"""Checklist do pacote do evento conforme fluxograma: roteiro, plano/ordem, ofícios, termos, justificativas."""
from __future__ import annotations

from viagens.models import Evento, Oficio
from viagens.services.justificativa_helpers import (
    exige_justificativa,
    get_dias_antecedencia,
    justificativa_preenchida,
)
from viagens.services.documentos_manager import _get_ordem, _get_plano


def build_evento_checklist(evento: Evento) -> dict:
    """
    Retorna o checklist do pacote do evento para exibição e regra "pronto para protocolar".
    - required_docs: roteiro, plano_ou_ordem (se não tem convite), ofícios, justificativas por ofício, termos por servidor não-ASCOM
    - readiness: pronto_para_protocolar
    """
    oficios = list(
        evento.oficios.all().prefetch_related("trechos", "viajantes").order_by("id")
    )
    # Trechos: união dos trechos de todos os ofícios do evento
    total_trechos = sum(o.trechos.count() for o in oficios)
    roteiro_ok = total_trechos > 0

    tem_convite = getattr(evento, "tem_convite_ou_oficio_evento", True)
    exige_plano_ou_ordem = not tem_convite

    tem_plano = False
    tem_ordem = False
    for o in oficios:
        if _get_plano(o) is not None:
            tem_plano = True
        if _get_ordem(o) is not None:
            tem_ordem = True
    plano_ou_ordem_ok = (tem_plano or tem_ordem) if exige_plano_ou_ordem else True

    # Ofícios: pelo menos 1; cada um com justificativa ok se antecedência < 10
    oficios_check = []
    all_oficios_ok = len(oficios) > 0
    for o in oficios:
        exige = exige_justificativa(o)
        ok = justificativa_preenchida(o)
        all_oficios_ok = all_oficios_ok and (not exige or ok)
        oficios_check.append({
            "oficio_id": o.id,
            "numero_display": getattr(o, "numero_formatado", None) or getattr(o, "oficio", None) or str(o.id),
            "exige_justificativa": exige,
            "justificativa_ok": ok,
            "dias_antecedencia": get_dias_antecedencia(o),
        })

    # Termos: por (ofício, viajante) não-ASCOM; dispensado conta como ok
    termos_check = []
    termos_ok = True
    for o in oficios:
        termos_oficio = {
            t.viajante_id: t
            for t in o.termos_autorizacao.select_related("viajante").all()
            if t.viajante_id
        }
        for v in o.viajantes.all():
            if getattr(v, "is_ascom", True):
                continue
            termo = termos_oficio.get(v.id)
            ok = termo is not None  # existe (ativo ou dispensado)
            if not ok:
                termos_ok = False
            termos_check.append({
                "oficio_id": o.id,
                "viajante_id": v.id,
                "nome": v.nome or "",
                "is_ascom": False,
                "termo_ok": ok,
                "dispensado": termo.dispensado if termo else False,
            })
    if not any(not c["is_ascom"] for c in termos_check):
        termos_ok = True

    pronto_para_protocolar = (
        roteiro_ok
        and plano_ou_ordem_ok
        and all_oficios_ok
        and termos_ok
    )

    return {
        "required_docs": {
            "roteiro": {
                "status": "ok" if roteiro_ok else "vazio",
                "total_trechos": total_trechos,
            },
            "plano_ou_ordem": {
                "required": exige_plano_ou_ordem,
                "status": "ok" if plano_ou_ordem_ok else "pendente",
                "tem_plano": tem_plano,
                "tem_ordem": tem_ordem,
            },
            "oficios": oficios_check,
            "termos": termos_check,
        },
        "readiness": {
            "pronto_para_protocolar": pronto_para_protocolar,
        },
        "tem_convite_ou_oficio_evento": tem_convite,
    }
