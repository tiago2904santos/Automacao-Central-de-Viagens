from __future__ import annotations

from datetime import date
from typing import Any, Mapping, TypedDict

from django.utils import timezone
from django.utils.dateparse import parse_date

from viagens.models import Oficio


class JustificativaTemplate(TypedDict):
    label: str
    texto: str


JUSTIFICATIVA_DIAS_MINIMOS = 10

JUSTIFICATIVA_TEMPLATES: dict[str, JustificativaTemplate] = {
    "recebimento_tardio": {
        "label": "Demanda recebida tardiamente",
        "texto": (
            "Em atenção ao prazo de 10 dias estabelecido pelo Decreto nº 6.358/2024 e ao "
            "Ofício Circular nº 340/2024-GAF, informamos que o pedido de deslocamento "
            "referente ao Ofício nº X/ANO foi encaminhado na data em que a demanda foi "
            "formalmente recebida por esta unidade, razão pela qual o presente ofício está "
            "sendo enviado com prazo inferior ao estipulado.\n"
            "Esclarecemos que o envio ocorreu imediatamente após o recebimento da solicitação, "
            "não havendo possibilidade de cumprimento integral do prazo regulamentar, "
            "considerando a data e o horário em que o pedido foi repassado para providências."
        ),
    },
    "operacao_policial": {
        "label": "Operação policial (orientações DG)",
        "texto": (
            "Em atenção ao prazo de 10 dias estabelecido pelo Decreto nº 6.358/2024, e ofício "
            "Circular 340/2024-GAF, justificamos que o envio se deu em data próxima ao "
            "deslocamento em razão da necessidade de aguardar as orientações do Gabinete do "
            "Delegado-Geral acerca da operação policial, imprescindíveis para a definição das "
            "diretrizes e correta formalização da demanda."
        ),
    },
    "evento": {
        "label": "Confirmação tardia de evento",
        "texto": (
            "O prazo de 10 (dez) dias previsto no Decreto nº 6.358/2024 e no Ofício Circular "
            "nº 340/2024-GAF não pôde ser observado, uma vez que a equipe encontrava-se em "
            "tratativas para a confirmação da data e definição do local do evento, circunstância "
            "que inviabilizou o protocolo antecipado da solicitação de diárias e a adoção das "
            "demais providências administrativas pertinentes."
        ),
    },
    "servidores": {
        "label": "Documentos/autorização de servidores",
        "texto": (
            "Em atenção ao prazo de 10 dias estabelecido pelo Decreto nº 6.358/2024, e ofício "
            "Circular 340/2024-GAF, o deslocamento referente ao ofício X/ANO ocorreu de forma "
            "intempestiva. Justificamos que não foi possível encaminhar o ofício com a devida "
            "antecedência, uma vez que contamos com a participação de servidores de outras "
            "unidades que compõem a equipe de apoio a esta Assessoria de Comunicação, e essa "
            "readequação demanda a espera para recebimento das autorizações das chefias, bem "
            "como da documentação desses servidores, que nem sempre ocorrem em tempo hábil."
        ),
    },
}


def get_justificativa_template_text(modelo: str | None) -> str:
    if not modelo:
        return ""
    template = JUSTIFICATIVA_TEMPLATES.get(str(modelo).strip())
    if not template:
        return ""
    return template["texto"]


def _coerce_saida_data(value: Any) -> date | None:
    if isinstance(value, date):
        return value
    if value in (None, ""):
        return None
    return parse_date(str(value))


def get_primeira_saida_data(
    *,
    oficio: Oficio | None = None,
    trechos_payload: list[Mapping[str, Any]] | None = None,
) -> date | None:
    if trechos_payload is not None:
        datas: list[date] = []
        for trecho in trechos_payload:
            saida_data = _coerce_saida_data(trecho.get("saida_data"))
            if saida_data:
                datas.append(saida_data)
        return min(datas) if datas else None

    if oficio is None:
        return None
    primeiro_trecho = (
        oficio.trechos.exclude(saida_data__isnull=True)
        .order_by("ordem", "id")
        .only("saida_data")
        .first()
    )
    return primeiro_trecho.saida_data if primeiro_trecho else None


def get_antecedencia_dias(
    *,
    oficio: Oficio | None = None,
    trechos_payload: list[Mapping[str, Any]] | None = None,
    referencia_data: date | None = None,
) -> int | None:
    primeira_saida = get_primeira_saida_data(oficio=oficio, trechos_payload=trechos_payload)
    if not primeira_saida:
        return None
    data_base = referencia_data or timezone.localdate()
    return (primeira_saida - data_base).days


def requires_justificativa(
    *,
    oficio: Oficio | None = None,
    trechos_payload: list[Mapping[str, Any]] | None = None,
    referencia_data: date | None = None,
) -> bool:
    antecedencia = get_antecedencia_dias(
        oficio=oficio,
        trechos_payload=trechos_payload,
        referencia_data=referencia_data,
    )
    if antecedencia is None:
        return False
    return antecedencia < JUSTIFICATIVA_DIAS_MINIMOS


def has_justificativa_preenchida(oficio: Oficio) -> bool:
    return bool((oficio.justificativa_texto or "").strip())
