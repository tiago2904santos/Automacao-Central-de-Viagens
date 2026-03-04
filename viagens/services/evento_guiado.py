# Progresso do fluxo guiado por Evento (wizard)
from __future__ import annotations

from django.urls import reverse

from viagens.models import Evento


def build_evento_guiado_progresso(evento: Evento) -> dict:
    """
    Retorna o progresso das etapas do fluxo guiado para o painel.
    - etapas: lista de { numero, label, status, url, em_construcao }
    - etapa1_ok: bool
    - proximo_passo_url, proximo_passo_label
    """
    etapa1_ok = bool(evento.titulo and str(evento.titulo).strip() and evento.titulo != "Rascunho")
    etapa2_ok = evento.roteiros.filter(ativo=True).exists()
    etapa3_ok = evento.oficios.exists()
    exige_plano_ou_ordem = not getattr(evento, "tem_convite_ou_oficio_evento", True)
    # Etapa 4 OK quando dispensada (tem convite) OU quando existe PT ou OS em algum ofício do evento
    tem_pt_ou_os = False
    if exige_plano_ou_ordem and evento.oficios.exists():
        from viagens.models import PlanoTrabalho, OrdemServico
        oficio_ids = list(evento.oficios.values_list("id", flat=True))
        tem_pt_ou_os = (
            PlanoTrabalho.objects.filter(oficio_id__in=oficio_ids).exists()
            or OrdemServico.objects.filter(oficio_id__in=oficio_ids).exists()
        )
    etapa4_ok = not exige_plano_ou_ordem or tem_pt_ou_os
    # Etapa 5: termos por (ofício, viajante); OK quando todos necessários têm termo ou dispensa; "não necessário" se não há viajantes não-ASCOM
    from viagens.services.evento_termos_etapa5 import etapa5_necessarios_ok
    _tem_necessarios, etapa5_ok = etapa5_necessarios_ok(evento)
    etapa5_nao_necessario = not _tem_necessarios

    # Etapa 6: finalização (uploads); OK quando todos obrigatórios têm upload (ou dispensa)
    from viagens.services.evento_assinados import is_evento_pronto_para_compilar
    etapa6_ok = is_evento_pronto_para_compilar(evento)

    def _url_placeholder(num: int):
        return reverse(
            "evento_guiado_etapa_placeholder",
            kwargs={"evento_id": evento.id, "etapa_num": num},
        )

    url_etapa1 = reverse("evento_guiado_etapa1", kwargs={"evento_id": evento.id})
    url_etapa2 = reverse("evento_guiado_etapa2", kwargs={"evento_id": evento.id})
    url_etapa3 = reverse("evento_guiado_etapa3", kwargs={"evento_id": evento.id})
    url_etapa4 = reverse("evento_guiado_etapa4", kwargs={"evento_id": evento.id})
    url_etapa5 = reverse("evento_guiado_etapa5", kwargs={"evento_id": evento.id})
    url_etapa6 = reverse("evento_guiado_etapa6", kwargs={"evento_id": evento.id})
    url_pacote = reverse("evento_pacote", kwargs={"evento_id": evento.id})

    etapas = [
        {
            "numero": 1,
            "label": "Cadastro do evento",
            "status": "ok" if etapa1_ok else "pendente",
            "url": url_etapa1,
            "em_construcao": False,
        },
        {
            "numero": 2,
            "label": "Roteiro",
            "status": "pendente" if not etapa2_ok else "ok",
            "url": url_etapa2,
            "em_construcao": False,
        },
        {
            "numero": 3,
            "label": "Ofícios do evento",
            "status": "pendente" if not etapa3_ok else "ok",
            "url": url_etapa3,
            "em_construcao": False,
        },
        {
            "numero": 4,
            "label": "Plano de trabalho / Ordem de serviço",
            "status": "nao_necessario" if not exige_plano_ou_ordem else ("ok" if etapa4_ok else "pendente"),
            "url": url_etapa4,
            "em_construcao": False,
        },
        {
            "numero": 5,
            "label": "Termos de autorização",
            "status": "nao_necessario" if etapa5_nao_necessario else ("ok" if etapa5_ok else "pendente"),
            "url": url_etapa5,
            "em_construcao": False,
        },
        {
            "numero": 6,
            "label": "Finalização (uploads e export)",
            "status": "ok" if etapa6_ok else "pendente",
            "url": url_etapa6,
            "em_construcao": False,
        },
    ]

    # Próximo passo recomendado
    if not etapa1_ok:
        proximo_passo_url = url_etapa1
        proximo_passo_label = "Completar cadastro do evento"
    elif not etapa2_ok:
        proximo_passo_url = url_etapa2
        proximo_passo_label = "Definir roteiro"
    elif not etapa3_ok:
        proximo_passo_url = url_etapa3
        proximo_passo_label = "Ofícios do evento"
    elif not etapa4_ok:
        proximo_passo_url = url_etapa4
        proximo_passo_label = "PT/OS ou base formal"
    elif not etapa5_ok:
        proximo_passo_url = url_etapa5
        proximo_passo_label = "Termos de autorização"
    elif not etapa6_ok:
        proximo_passo_url = url_etapa6
        proximo_passo_label = "Finalização (uploads)"
    else:
        proximo_passo_url = url_etapa6
        proximo_passo_label = "Finalização"

    return {
        "etapa1_ok": etapa1_ok,
        "etapas": etapas,
        "proximo_passo_url": proximo_passo_url,
        "proximo_passo_label": proximo_passo_label,
        "url_pacote": url_pacote,
    }
