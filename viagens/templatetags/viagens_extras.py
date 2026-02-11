from django import template

from viagens.utils.normalize import format_cpf as _format_cpf
from viagens.utils.normalize import format_phone as _format_phone
from viagens.utils.normalize import format_rg as _format_rg

register = template.Library()


@register.filter
def first_name(value: str) -> str:
    if not value:
        return ""
    return value.strip().split()[0]


@register.filter
def format_cpf(value: str) -> str:
    return _format_cpf(value)


@register.filter
def format_phone(value: str) -> str:
    return _format_phone(value)


@register.filter
def format_rg(value: str) -> str:
    return _format_rg(value)
