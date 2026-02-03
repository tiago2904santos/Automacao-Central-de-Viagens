from django import template

register = template.Library()


@register.filter
def first_name(value: str) -> str:
    if not value:
        return ""
    return value.strip().split()[0]
