from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from decimal import Decimal, ROUND_HALF_UP
import unicodedata

from viagens.services.diarias import TABELA_DIARIAS, formatar_valor_diarias


CAPITAIS_POR_UF = {
    "AC": "RIO BRANCO",
    "AL": "MACEIO",
    "AP": "MACAPA",
    "AM": "MANAUS",
    "BA": "SALVADOR",
    "CE": "FORTALEZA",
    "DF": "BRASILIA",
    "ES": "VITORIA",
    "GO": "GOIANIA",
    "MA": "SAO LUIS",
    "MT": "CUIABA",
    "MS": "CAMPO GRANDE",
    "MG": "BELO HORIZONTE",
    "PA": "BELEM",
    "PB": "JOAO PESSOA",
    "PR": "CURITIBA",
    "PE": "RECIFE",
    "PI": "TERESINA",
    "RJ": "RIO DE JANEIRO",
    "RN": "NATAL",
    "RS": "PORTO ALEGRE",
    "RO": "PORTO VELHO",
    "RR": "BOA VISTA",
    "SC": "FLORIANOPOLIS",
    "SP": "SAO PAULO",
    "SE": "ARACAJU",
    "TO": "PALMAS",
}


@dataclass(frozen=True)
class PeriodMarker:
    saida: datetime
    destino_cidade: str
    destino_uf: str


def _normalize_city_name(value: str | None) -> str:
    raw = unicodedata.normalize("NFKD", (value or "").strip().upper())
    return "".join(ch for ch in raw if not unicodedata.combining(ch))


def classify(cidade: str | None, uf: str | None) -> str:
    uf_norm = (uf or "").strip().upper()
    cidade_norm = _normalize_city_name(cidade)
    if uf_norm == "DF" and cidade_norm == "BRASILIA":
        return "BRASILIA"
    if uf_norm and cidade_norm and CAPITAIS_POR_UF.get(uf_norm) == cidade_norm:
        return "CAPITAL"
    return "INTERIOR"


def _segment_breakdown(start: datetime, end: datetime) -> tuple[int, int, Decimal, Decimal]:
    total_seconds = (end - start).total_seconds()
    if total_seconds <= 0:
        raise ValueError("Periodo invalido para calculo de diarias.")

    total_horas = Decimal(str(total_seconds / 3600)).quantize(
        Decimal("0.01"), rounding=ROUND_HALF_UP
    )
    dias_inteiros = int(total_seconds // (24 * 3600))
    resto_seconds = total_seconds - (dias_inteiros * 24 * 3600)

    parcial = 0
    if start.date() != end.date() and total_seconds < 24 * 3600:
        dias_inteiros = 1
        resto_seconds = 0
    else:
        if resto_seconds <= 6 * 3600:
            parcial = 0
        elif resto_seconds <= 8 * 3600:
            parcial = 15
        else:
            parcial = 30

    horas_adicionais = Decimal(str(resto_seconds / 3600)).quantize(
        Decimal("0.01"), rounding=ROUND_HALF_UP
    )
    return dias_inteiros, parcial, horas_adicionais, total_horas


def _currency(value: Decimal) -> str:
    return formatar_valor_diarias(value.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP))


def _format_dt(value: datetime) -> tuple[str, str]:
    return value.strftime("%d/%m/%Y"), value.strftime("%H:%M")


def _total_diarias_resumo(periodos: list[dict]) -> str:
    full = sum(int(item.get("n_diarias", 0) or 0) for item in periodos)
    p15 = sum(1 for item in periodos if int(item.get("percentual_adicional", 0) or 0) == 15)
    p30 = sum(1 for item in periodos if int(item.get("percentual_adicional", 0) or 0) == 30)
    partes: list[str] = []
    if full:
        partes.append(f"{full} x 100%")
    if p15:
        partes.append(f"{p15} x 15%")
    if p30:
        partes.append(f"{p30} x 30%")
    return " + ".join(partes)


def build_periods(
    markers: list[PeriodMarker],
    chegada_final_sede: datetime,
    *,
    quantidade_servidores: int = 1,
) -> list[dict]:
    if not markers:
        raise ValueError("Preencha datas e horas para calcular.")
    if not chegada_final_sede:
        raise ValueError("Preencha datas e horas para calcular.")

    sorted_markers = sorted(markers, key=lambda item: item.saida)
    periodos: list[dict] = []

    for idx, marker in enumerate(sorted_markers):
        start = marker.saida
        if idx + 1 < len(sorted_markers):
            end = sorted_markers[idx + 1].saida
        else:
            end = chegada_final_sede
        if end <= start:
            raise ValueError("Preencha datas e horas para calcular.")

        tipo = classify(marker.destino_cidade, marker.destino_uf)
        dias_inteiros, parcial, horas_adicionais, total_horas = _segment_breakdown(start, end)

        tabela = TABELA_DIARIAS.get(tipo, TABELA_DIARIAS["INTERIOR"])
        valor_24h = tabela["24h"]
        valor_parcial = Decimal("0.00")
        if parcial == 15:
            valor_parcial = tabela["15"]
        elif parcial == 30:
            valor_parcial = tabela["30"]

        valor_1_servidor = (valor_24h * dias_inteiros) + valor_parcial
        subtotal = valor_1_servidor * max(0, int(quantidade_servidores or 0))

        data_saida, hora_saida = _format_dt(start)
        data_chegada, hora_chegada = _format_dt(end)
        periodos.append(
            {
                "tipo": tipo,
                "data_saida": data_saida,
                "hora_saida": hora_saida,
                "data_chegada": data_chegada,
                "hora_chegada": hora_chegada,
                "n_diarias": dias_inteiros,
                "horas_adicionais": float(horas_adicionais),
                "valor_diaria": _currency(valor_24h),
                "subtotal": _currency(subtotal),
                "subtotal_decimal": subtotal,
                "percentual_adicional": parcial,
                "total_horas_periodo": float(total_horas),
            }
        )

    return periodos


def calculate_periodized_diarias(
    markers: list[PeriodMarker],
    chegada_final_sede: datetime,
    *,
    quantidade_servidores: int = 1,
    valor_extenso_fn=None,
) -> dict:
    periodos = build_periods(
        markers,
        chegada_final_sede,
        quantidade_servidores=quantidade_servidores,
    )
    total_valor_decimal = sum(
        (item["subtotal_decimal"] for item in periodos),
        Decimal("0.00"),
    )
    total_horas = sum(float(item.get("total_horas_periodo", 0) or 0) for item in periodos)
    resumo_diarias = _total_diarias_resumo(periodos)
    servidores = max(1, int(quantidade_servidores or 1))
    valor_por_servidor = (
        (total_valor_decimal / Decimal(servidores)).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
        if servidores > 0
        else total_valor_decimal
    )
    valores_unitarios = [str(item.get("valor_diaria", "") or "").strip() for item in periodos]
    valores_unitarios = [item for item in valores_unitarios if item]
    if len(set(valores_unitarios)) == 1:
        valor_unitario_referencia = valores_unitarios[0]
    elif valores_unitarios:
        valor_unitario_referencia = f"{valores_unitarios[0]} (variavel por periodo)"
    else:
        valor_unitario_referencia = ""
    valor_extenso = ""
    if callable(valor_extenso_fn):
        valor_extenso = valor_extenso_fn(_currency(total_valor_decimal)) or ""

    periodos_out = []
    for item in periodos:
        row = dict(item)
        row.pop("subtotal_decimal", None)
        row.pop("total_horas_periodo", None)
        periodos_out.append(row)

    return {
        "periodos": periodos_out,
        "totais": {
            "total_diarias": resumo_diarias,
            "total_horas": round(total_horas, 2),
            "total_valor": _currency(total_valor_decimal),
            "valor_extenso": valor_extenso,
            "quantidade_servidores": servidores,
            "diarias_por_servidor": resumo_diarias,
            "valor_por_servidor": _currency(valor_por_servidor),
            "valor_unitario_referencia": valor_unitario_referencia,
        },
    }
