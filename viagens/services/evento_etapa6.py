# viagens/services/evento_etapa6.py
"""Checklist final da Etapa 6 (Finalização): uploads guiados em ordem fixa e critério de OK."""
from __future__ import annotations

from viagens.models import DocumentoEventoArquivo, Evento
from viagens.services.evento_assinados import get_status_assinados_evento, is_evento_pronto_para_compilar
from viagens.services.justificativa_helpers import exige_justificativa


def build_etapa6_checklist(evento: Evento) -> dict:
    """
    Retorna o checklist da Etapa 6 para a página única de uploads guiados.
    Ordem fixa: 1) Ofícios assinados, 2) Solicitação formal ou Plano/Ordem, 3) Justificativas, 4) Termos.
    Cada bloco tem: titulo, status_ok, items (lista com label, obrigatorio, assinado, arquivo, oficio_id?, viajante_id?, dispensado?).
    """
    status = get_status_assinados_evento(evento)
    oficios = list(evento.oficios.order_by("id").prefetch_related("viajantes", "termos_autorizacao"))
    tem_convite = status["tem_convite_ou_oficio_evento"]

    # Bloco 1: Ofícios assinados
    items_oficios = []
    bloco1_ok = True
    for o in status["oficios"]:
        assinado = o["assinado"]
        if not assinado:
            bloco1_ok = False
        items_oficios.append({
            "label": f"Ofício {o['numero_display']}",
            "obrigatorio": True,
            "assinado": assinado,
            "arquivo": o.get("arquivo"),
            "oficio_id": o["oficio_id"],
            "tipo": DocumentoEventoArquivo.Tipo.OFICIO_ASSINADO,
        })
    bloco1 = {
        "titulo": "1. Ofícios assinados pela chefia",
        "status_ok": bloco1_ok and len(items_oficios) > 0,
        "items": items_oficios,
    }

    # Bloco 2: Solicitação formal OU Plano/Ordem
    if tem_convite:
        sol = status.get("solicitacao_formal", {})
        bloco2_ok = sol.get("assinado", False)
        items_bloco2 = [{
            "label": "Solicitação formal (convite/ofício solicitante) assinada",
            "obrigatorio": True,
            "assinado": bloco2_ok,
            "arquivo": sol.get("arquivo"),
            "oficio_id": None,
            "viajante_id": None,
            "tipo": DocumentoEventoArquivo.Tipo.SOLICITACAO_FORMAL_ASSINADA,
        }]
    else:
        plano_ok = status["plano_ou_ordem"].get("plano_assinado") or status["plano_ou_ordem"].get("ordem_assinado")
        items_bloco2 = [
            {
                "label": "Plano de trabalho assinado",
                "obrigatorio": False,
                "assinado": status["plano_ou_ordem"].get("plano_assinado", False),
                "arquivo": status["plano_ou_ordem"].get("arquivo_plano"),
                "oficio_id": None,
                "tipo": DocumentoEventoArquivo.Tipo.PLANO_ASSINADO,
            },
            {
                "label": "Ordem de serviço assinada",
                "obrigatorio": False,
                "assinado": status["plano_ou_ordem"].get("ordem_assinado", False),
                "arquivo": status["plano_ou_ordem"].get("arquivo_ordem"),
                "oficio_id": None,
                "tipo": DocumentoEventoArquivo.Tipo.ORDEM_ASSINADO,
            },
        ]
        bloco2_ok = plano_ok
    bloco2 = {
        "titulo": "2. Solicitação formal ou Plano/Ordem assinados",
        "status_ok": bloco2_ok,
        "items": items_bloco2,
    }

    # Bloco 3: Justificativas (condicional por ofício)
    items_just = []
    bloco3_ok = True
    for o in oficios:
        exige = exige_justificativa(o)
        num = getattr(o, "numero_formatado", None) or getattr(o, "oficio", None) or str(o.id)
        if not exige:
            items_just.append({
                "label": f"Justificativa (Ofício {num})",
                "obrigatorio": False,
                "nao_necessario": True,
                "assinado": True,
                "arquivo": None,
                "oficio_id": o.id,
                "tipo": DocumentoEventoArquivo.Tipo.JUSTIFICATIVA_ASSINADA,
            })
            continue
        j = next((x for x in status["justificativas"] if x["oficio_id"] == o.id), None)
        assinado = j and j.get("assinado", False)
        if not assinado:
            bloco3_ok = False
        items_just.append({
            "label": f"Justificativa (Ofício {num})",
            "obrigatorio": True,
            "nao_necessario": False,
            "assinado": assinado,
            "arquivo": j.get("arquivo") if j else None,
            "oficio_id": o.id,
            "tipo": DocumentoEventoArquivo.Tipo.JUSTIFICATIVA_ASSINADA,
        })
    bloco3 = {
        "titulo": "3. Justificativas assinadas (quando antecedência < 10 dias)",
        "status_ok": bloco3_ok,
        "items": items_just,
    }

    # Bloco 4: Termos (condicional por ofício/viajante; dispensado = OK)
    items_termos = []
    bloco4_ok = True
    for t in status["termos"]:
        if t.get("dispensado"):
            items_termos.append({
                "label": t.get("nome", ""),
                "obrigatorio": False,
                "dispensado": True,
                "assinado": True,
                "arquivo": None,
                "oficio_id": t.get("oficio_id"),
                "viajante_id": t.get("viajante_id"),
                "tipo": DocumentoEventoArquivo.Tipo.TERMO_ASSINADO,
            })
            continue
        assinado = t.get("assinado", False)
        if not assinado:
            bloco4_ok = False
        items_termos.append({
            "label": t.get("nome", ""),
            "obrigatorio": True,
            "dispensado": False,
            "assinado": assinado,
            "arquivo": t.get("arquivo"),
            "oficio_id": t.get("oficio_id"),
            "viajante_id": t.get("viajante_id"),
            "tipo": DocumentoEventoArquivo.Tipo.TERMO_ASSINADO,
        })
    bloco4 = {
        "titulo": "4. Termos de autorização assinados",
        "status_ok": bloco4_ok,
        "items": items_termos,
    }

    etapa6_ok = is_evento_pronto_para_compilar(evento)
    return {
        "blocos": [bloco1, bloco2, bloco3, bloco4],
        "etapa6_ok": etapa6_ok,
        "status": status,
    }


def listar_pendencias_etapa6(evento: Evento) -> list[str]:
    """Lista do que falta para a Etapa 6 estar OK (para bloqueio de export ZIP)."""
    from viagens.services.evento_assinados import listar_pendencias_compilacao
    return listar_pendencias_compilacao(evento)
