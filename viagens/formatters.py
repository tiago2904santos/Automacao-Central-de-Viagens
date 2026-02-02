from __future__ import annotations

from datetime import date, datetime, time

from .models import Cidade, Estado


def format_date(date_value: str | date | None) -> str:
    if not date_value:
        return ""
    if isinstance(date_value, date):
        return date_value.strftime("%d/%m/%Y")
    try:
        return datetime.strptime(date_value, "%Y-%m-%d").strftime("%d/%m/%Y")
    except ValueError:
        return str(date_value)


def format_time(time_value: str | time | None) -> str:
    if not time_value:
        return ""
    if isinstance(time_value, time):
        return time_value.strftime("%H:%M")
    return str(time_value).strip()


def format_date_time(date_value: str | date | None, time_value: str | time | None) -> str:
    date_part = format_date(date_value)
    time_part = format_time(time_value)
    if date_part and time_part:
        return f"{date_part} - {time_part}"
    return date_part or time_part


def format_trecho_local(cidade: Cidade | None, estado: Estado | None) -> str:
    if cidade and estado:
        return f"{cidade.nome}/{estado.sigla}"
    if cidade:
        return cidade.nome
    if estado:
        return estado.sigla
    return ""
