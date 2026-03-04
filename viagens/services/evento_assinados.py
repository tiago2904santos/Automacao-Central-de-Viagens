# viagens/services/evento_assinados.py
"""Status dos documentos ASSINADOS do pacote do evento e regra pronto_para_compilar."""
from __future__ import annotations

from viagens.models import DocumentoEventoArquivo, Evento, Oficio
from viagens.services.documentos_manager import _get_ordem, _get_plano
from viagens.services.justificativa_helpers import exige_justificativa


def _get_arquivo_assinado_ativo(evento: Evento, tipo: str, oficio_id: int | None = None, viajante_id: int | None = None) -> DocumentoEventoArquivo | None:
    """Retorna o arquivo assinado ativo para (evento, tipo, oficio?, viajante?)."""
    qs = DocumentoEventoArquivo.objects.filter(
        evento=evento,
        tipo=tipo,
        is_active=True,
    )
    if oficio_id is not None:
        qs = qs.filter(oficio_id=oficio_id)
    else:
        qs = qs.filter(oficio__isnull=True)
    if viajante_id is not None:
        qs = qs.filter(viajante_id=viajante_id)
    else:
        qs = qs.filter(viajante__isnull=True)
    return qs.order_by("-uploaded_at").first()


def get_status_assinados_evento(evento: Evento) -> dict:
    """
    Retorna o status de assinado por documento do evento.
    - oficios: lista { oficio_id, numero_display, assinado: bool, arquivo_info }
    - plano_ou_ordem: { required, plano_assinado, ordem_assinado, arquivo_plano?, arquivo_ordem? }
    - justificativas: lista por ofício que exige { oficio_id, assinado, arquivo_info }
    - termos: lista por viajante não-ASCOM { viajante_id, nome, assinado, arquivo_info }
    """
    oficios = list(evento.oficios.all().order_by("id").prefetch_related("viajantes"))
    tem_convite = getattr(evento, "tem_convite_ou_oficio_evento", True)
    exige_plano_ou_ordem = not tem_convite

    # Ofícios assinados
    oficios_status = []
    for o in oficios:
        arq = _get_arquivo_assinado_ativo(evento, DocumentoEventoArquivo.Tipo.OFICIO_ASSINADO, oficio_id=o.id)
        oficios_status.append({
            "oficio_id": o.id,
            "numero_display": getattr(o, "numero_formatado", None) or getattr(o, "oficio", None) or str(o.id),
            "assinado": arq is not None,
            "arquivo": _arquivo_info(arq),
        })

    # Plano / Ordem (1 por evento)
    plano_assinado = _get_arquivo_assinado_ativo(evento, DocumentoEventoArquivo.Tipo.PLANO_ASSINADO) is not None
    ordem_assinado = _get_arquivo_assinado_ativo(evento, DocumentoEventoArquivo.Tipo.ORDEM_ASSINADO) is not None
    arq_plano = _get_arquivo_assinado_ativo(evento, DocumentoEventoArquivo.Tipo.PLANO_ASSINADO)
    arq_ordem = _get_arquivo_assinado_ativo(evento, DocumentoEventoArquivo.Tipo.ORDEM_ASSINADO)
    plano_ou_ordem_status = {
        "required": exige_plano_ou_ordem,
        "plano_assinado": plano_assinado,
        "ordem_assinado": ordem_assinado,
        "arquivo_plano": _arquivo_info(arq_plano),
        "arquivo_ordem": _arquivo_info(arq_ordem),
    }

    # Justificativas (por ofício quando exige)
    justificativas_status = []
    for o in oficios:
        if not exige_justificativa(o):
            continue
        arq = _get_arquivo_assinado_ativo(evento, DocumentoEventoArquivo.Tipo.JUSTIFICATIVA_ASSINADA, oficio_id=o.id)
        justificativas_status.append({
            "oficio_id": o.id,
            "numero_display": getattr(o, "numero_formatado", None) or getattr(o, "oficio", None) or str(o.id),
            "assinado": arq is not None,
            "arquivo": _arquivo_info(arq),
        })

    # Viajantes não-ASCOM e termos assinados
    seen_v = set()
    viajantes_nao_ascom = []
    for o in oficios:
        for v in o.viajantes.all():
            if v.id not in seen_v:
                seen_v.add(v.id)
                if not getattr(v, "is_ascom", True):
                    viajantes_nao_ascom.append(v)
    termos_status = []
    for v in sorted(viajantes_nao_ascom, key=lambda x: (x.nome or "")):
        arq = _get_arquivo_assinado_ativo(evento, DocumentoEventoArquivo.Tipo.TERMO_ASSINADO, viajante_id=v.id)
        termos_status.append({
            "viajante_id": v.id,
            "nome": v.nome or "",
            "assinado": arq is not None,
            "arquivo": _arquivo_info(arq),
        })

    return {
        "oficios": oficios_status,
        "plano_ou_ordem": plano_ou_ordem_status,
        "justificativas": justificativas_status,
        "termos": termos_status,
        "tem_convite_ou_oficio_evento": tem_convite,
    }


def _arquivo_info(arq: DocumentoEventoArquivo | None) -> dict | None:
    if arq is None:
        return None
    return {
        "id": arq.id,
        "name": arq.original_name or (arq.arquivo.name.split("/")[-1] if arq.arquivo else ""),
        "uploaded_at": arq.uploaded_at,
        "url": arq.arquivo.url if arq.arquivo else None,
    }


def is_evento_pronto_para_compilar(evento: Evento) -> bool:
    """
    True somente se todos os documentos obrigatórios têm versão ASSINADA (upload).
    - Todos os ofícios do evento têm OFICIO_ASSINADO
    - Se evento exige Plano/Ordem: existe PLANO_ASSINADO ou ORDEM_ASSINADO
    - Para cada ofício que exige justificativa: existe JUSTIFICATIVA_ASSINADA
    - Para cada viajante não-ASCOM: existe TERMO_ASSINADO
    """
    status = get_status_assinados_evento(evento)
    if not status["oficios"]:
        return False
    for o in status["oficios"]:
        if not o["assinado"]:
            return False
    if status["plano_ou_ordem"]["required"]:
        if not (status["plano_ou_ordem"]["plano_assinado"] or status["plano_ou_ordem"]["ordem_assinado"]):
            return False
    for j in status["justificativas"]:
        if not j["assinado"]:
            return False
    for t in status["termos"]:
        if not t["assinado"]:
            return False
    return True


def listar_pendencias_compilacao(evento: Evento) -> list[str]:
    """Lista mensagens de pendências para exibir na UI (o que falta para compilar)."""
    status = get_status_assinados_evento(evento)
    pendencias = []
    for o in status["oficios"]:
        if not o["assinado"]:
            pendencias.append(f"Ofício {o['numero_display']}: falta upload do documento assinado.")
    if status["plano_ou_ordem"]["required"]:
        if not (status["plano_ou_ordem"]["plano_assinado"] or status["plano_ou_ordem"]["ordem_assinado"]):
            pendencias.append("Plano de Trabalho ou Ordem de Serviço: falta upload do documento assinado.")
    for j in status["justificativas"]:
        if not j["assinado"]:
            pendencias.append(f"Justificativa (ofício {j['numero_display']}): falta upload do documento assinado.")
    for t in status["termos"]:
        if not t["assinado"]:
            pendencias.append(f"Termo do servidor {t['nome']}: falta upload do documento assinado.")
    return pendencias
