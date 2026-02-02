from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from decimal import Decimal, ROUND_HALF_UP


@dataclass(frozen=True)
class DiariasResultado:
    dias_100: int
    parcial: int
    quantidade_diarias_str: str
    valor_1_servidor: Decimal
    valor_total_oficio: Decimal


TABELA_DIARIAS = {
    "INTERIOR": {
        "24h": Decimal("290.55"),
        "15": Decimal("43.58"),
        "30": Decimal("87.17"),
    },
    "CAPITAL": {
        "24h": Decimal("371.26"),
        "15": Decimal("55.69"),
        "30": Decimal("111.38"),
    },
    "BRASILIA": {
        "24h": Decimal("468.12"),
        "15": Decimal("70.22"),
        "30": Decimal("140.43"),
    },
}


def _format_currency(valor: Decimal) -> str:
    quantizado = valor.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    return f"{quantizado:.2f}".replace(".", ",")


def calcular_diarias(
    tipo_destino: str,
    saida_sede: datetime | None,
    chegada_sede: datetime | None,
    quantidade_servidores: int,
) -> DiariasResultado:
    if not tipo_destino or not saida_sede or not chegada_sede:
        return DiariasResultado(
            dias_100=0,
            parcial=0,
            quantidade_diarias_str="",
            valor_1_servidor=Decimal("0.00"),
            valor_total_oficio=Decimal("0.00"),
        )

    total_seconds = (chegada_sede - saida_sede).total_seconds()
    if total_seconds <= 0:
        return DiariasResultado(
            dias_100=0,
            parcial=0,
            quantidade_diarias_str="",
            valor_1_servidor=Decimal("0.00"),
            valor_total_oficio=Decimal("0.00"),
        )

    horas_inteiras = int(total_seconds // 3600)
    dias_inteiros = horas_inteiras // 24
    resto_seconds = total_seconds - (dias_inteiros * 24 * 3600)

    parcial = 0
    if saida_sede.date() != chegada_sede.date() and total_seconds < 24 * 3600:
        dias_inteiros = 1
        resto_seconds = 0
    else:
        if resto_seconds <= 6 * 3600:
            parcial = 0
        elif resto_seconds <= 8 * 3600:
            parcial = 15
        else:
            parcial = 30

    partes = []
    if dias_inteiros > 0:
        partes.append(f"{dias_inteiros} x 100%")
    if parcial > 0:
        partes.append(f"1 x {parcial}%")

    tabela = TABELA_DIARIAS.get(tipo_destino, {})
    valor_24h = tabela.get("24h", Decimal("0.00"))
    valor_15 = tabela.get("15", Decimal("0.00"))
    valor_30 = tabela.get("30", Decimal("0.00"))

    valor_parcial = Decimal("0.00")
    if parcial == 15:
        valor_parcial = valor_15
    elif parcial == 30:
        valor_parcial = valor_30

    valor_1_servidor = (valor_24h * dias_inteiros) + valor_parcial
    servidores = quantidade_servidores if quantidade_servidores > 0 else 0
    valor_total = valor_1_servidor * servidores

    return DiariasResultado(
        dias_100=dias_inteiros,
        parcial=parcial,
        quantidade_diarias_str=" + ".join(partes),
        valor_1_servidor=valor_1_servidor,
        valor_total_oficio=valor_total,
    )


def formatar_valor_diarias(valor: Decimal) -> str:
    return _format_currency(valor)
