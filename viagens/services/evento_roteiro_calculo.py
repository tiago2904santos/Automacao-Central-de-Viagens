# Cálculo de chegada a partir de saída + duração (Etapa 2 - Roteiro do Evento)
from __future__ import annotations

from datetime import date, datetime, time, timedelta

from django.utils.dateparse import parse_date, parse_time


def parse_duracao_minutos(value: str | int | None) -> int | None:
    """
    Converte duração para minutos.
    - "6:30" ou "6:30" -> 390
    - "90" -> 90 (apenas número = minutos)
    - None ou vazio -> None
    """
    if value is None:
        return None
    if isinstance(value, int):
        return value if value >= 0 else None
    raw = (value or "").strip()
    if not raw:
        return None
    if ":" in raw:
        # HH:MM ou H:MM
        parts = raw.split(":", 1)
        try:
            h = int(parts[0].strip())
            m = int(parts[1].strip()) if len(parts) > 1 else 0
            if h < 0 or m < 0 or m >= 60:
                return None
            return h * 60 + m
        except (ValueError, IndexError):
            return None
    try:
        n = int(raw)
        return n if n >= 0 else None
    except ValueError:
        return None


def calcular_chegada(
    saida_data: date | None,
    saida_hora: time | None,
    duracao_minutos: int | None,
) -> tuple[date | None, time | None]:
    """
    Retorna (chegada_data, chegada_hora) a partir de saída + duração em minutos.
    Se faltar algum dado, retorna (None, None).
    """
    if not saida_data or not saida_hora or duracao_minutos is None or duracao_minutos < 0:
        return None, None
    inicio = datetime.combine(saida_data, saida_hora)
    fim = inicio + timedelta(minutes=duracao_minutos)
    return fim.date(), fim.time().replace(second=0, microsecond=0)
