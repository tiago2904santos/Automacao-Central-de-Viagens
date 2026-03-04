# viagens/services/justificativa_helpers.py
"""Helpers para justificativa: antecedência, exige justificativa, dados para listagem."""
from __future__ import annotations

from datetime import date, datetime

from viagens.models import Oficio


DIAS_ANTECEDENCIA_MINIMO = 10


def _to_date(value) -> date | None:
    if value is None:
        return None
    if isinstance(value, date) and not isinstance(value, datetime):
        return value
    if isinstance(value, datetime):
        return value.date()
    if hasattr(value, "date"):
        return value.date()
    return None


def get_data_inicio_viagem(oficio: Oficio) -> date | None:
    """Data de início da viagem (menor saida_data entre os trechos)."""
    datas = []
    for t in oficio.trechos.all():
        saida = getattr(t, "saida_data", None)
        if saida is not None:
            d = _to_date(saida)
            if d is not None:
                datas.append(d)
    return min(datas) if datas else None


def get_dias_antecedencia(oficio: Oficio) -> int | None:
    """Dias entre data de criação do ofício e data de início da viagem. None se sem dados."""
    data_inicio = get_data_inicio_viagem(oficio)
    if data_inicio is None:
        return None
    created = getattr(oficio, "created_at", None)
    if created is None:
        return None
    data_criacao = _to_date(created)
    if data_criacao is None:
        return None
    return (data_inicio - data_criacao).days


def exige_justificativa(oficio: Oficio) -> bool:
    """True se antecedência < 10 dias (e há data de viagem)."""
    dias = get_dias_antecedencia(oficio)
    return dias is not None and dias < DIAS_ANTECEDENCIA_MINIMO


def justificativa_preenchida(oficio: Oficio) -> bool:
    """True se o ofício tem texto de justificativa preenchido."""
    texto = (getattr(oficio, "justificativa_texto", "") or "").strip()
    return bool(texto)
