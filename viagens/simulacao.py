from __future__ import annotations

from decimal import Decimal, ROUND_HALF_UP
from datetime import date, datetime, time

from django.utils.dateparse import parse_date, parse_time

from .diarias import CAPITAIS_POR_UF, PeriodMarker, calculate_periodized_diarias
from .services.diarias import TABELA_DIARIAS, formatar_valor_diarias
from .services.oficio_helpers import valor_por_extenso_ptbr

TIPOS_MANUAIS_VALIDOS = {"AUTOMATICO", "INTERIOR", "CAPITAL", "BRASILIA"}


def _coerce_date(value: date | str | None) -> date | None:
    if isinstance(value, date):
        return value
    if isinstance(value, str):
        raw = value.strip()
        if not raw:
            return None
        return parse_date(raw)
    return None


def _coerce_time(value: time | str | None) -> time | None:
    if isinstance(value, time):
        return value
    if isinstance(value, str):
        raw = value.strip()
        if not raw:
            return None
        return parse_time(raw)
    return None


def _coerce_int(value: int | str | None, default: int = 1) -> int:
    try:
        return int(value)
    except (TypeError, ValueError):
        return default


def _normalize_manual_tipo(value: str | None) -> str:
    tipo = (value or "AUTOMATICO").strip().upper()
    if tipo not in TIPOS_MANUAIS_VALIDOS:
        return "AUTOMATICO"
    return tipo


def _normalize_tipo_rapido(value: str | None) -> str:
    tipo = _normalize_manual_tipo(value)
    if tipo == "AUTOMATICO":
        raise ValueError("Informe o tipo do destino no modo rapido.")
    return tipo


def _apply_manual_tipo(
    cidade_destino: str | None,
    uf_destino: str | None,
    tipo_manual: str,
) -> tuple[str, str]:
    cidade = (cidade_destino or "").strip()
    uf = (uf_destino or "").strip().upper()

    if tipo_manual == "AUTOMATICO":
        return cidade, uf
    if tipo_manual == "BRASILIA":
        return "BRASILIA", "DF"
    if tipo_manual == "CAPITAL":
        uf_ref = uf if uf in CAPITAIS_POR_UF else "PR"
        return CAPITAIS_POR_UF.get(uf_ref, "CURITIBA"), uf_ref

    uf_ref = uf if uf else "PR"
    cidade_ref = cidade if cidade else "INTERIOR"
    if CAPITAIS_POR_UF.get(uf_ref) == cidade_ref.upper():
        cidade_ref = f"INTERIOR {uf_ref}"
    return cidade_ref, uf_ref


def _currency(value: Decimal) -> str:
    return formatar_valor_diarias(value.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP))


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


def _format_dt(value: datetime) -> tuple[str, str]:
    return value.strftime("%d/%m/%Y"), value.strftime("%H:%M")


def calculate_simulacao_diaria(
    data_saida: date | str | None,
    hora_saida: time | str | None,
    data_chegada: date | str | None,
    hora_chegada: time | str | None,
    cidade_destino: str | None,
    uf_destino: str | None,
    quantidade_servidores: int | str | None = 1,
    *,
    tipo_manual: str | None = None,
) -> dict:
    saida_data = _coerce_date(data_saida)
    saida_hora = _coerce_time(hora_saida)
    chegada_data = _coerce_date(data_chegada)
    chegada_hora = _coerce_time(hora_chegada)
    if not saida_data or not saida_hora or not chegada_data or not chegada_hora:
        raise ValueError("Preencha datas e horas para calcular.")

    saida_dt = datetime.combine(saida_data, saida_hora)
    chegada_dt = datetime.combine(chegada_data, chegada_hora)
    if chegada_dt <= saida_dt:
        raise ValueError("A chegada precisa ser posterior a saida.")

    tipo_manual_norm = _normalize_manual_tipo(tipo_manual)
    destino_cidade, destino_uf = _apply_manual_tipo(
        cidade_destino,
        uf_destino,
        tipo_manual_norm,
    )
    marker = PeriodMarker(
        saida=saida_dt,
        destino_cidade=destino_cidade,
        destino_uf=destino_uf,
    )
    resultado = calculate_periodized_diarias(
        [marker],
        chegada_dt,
        quantidade_servidores=max(1, _coerce_int(quantidade_servidores, default=1)),
        valor_extenso_fn=valor_por_extenso_ptbr,
    )
    if tipo_manual_norm != "AUTOMATICO":
        for periodo in resultado.get("periodos", []):
            periodo["tipo"] = tipo_manual_norm
    return resultado


def calculate_periods_from_payload(
    periods: list[dict],
    quantidade_servidores: int | str | None = 1,
) -> dict:
    if not periods:
        raise ValueError("Adicione ao menos um periodo para calcular.")

    servidores = max(1, _coerce_int(quantidade_servidores, default=1))
    periodos_out: list[dict] = []
    total_valor_decimal = Decimal("0.00")
    total_horas = 0.0

    for idx, payload in enumerate(periods, start=1):
        if not isinstance(payload, dict):
            raise ValueError(f"Periodo {idx}: formato invalido.")
        tipo = _normalize_tipo_rapido(payload.get("tipo"))
        start_date = _coerce_date(payload.get("start_date"))
        start_time = _coerce_time(payload.get("start_time"))
        end_date = _coerce_date(payload.get("end_date"))
        end_time = _coerce_time(payload.get("end_time"))
        if not start_date or not start_time or not end_date or not end_time:
            raise ValueError(f"Periodo {idx}: preencha data e hora de inicio e fim.")

        start_dt = datetime.combine(start_date, start_time)
        end_dt = datetime.combine(end_date, end_time)
        if end_dt <= start_dt:
            raise ValueError(f"Periodo {idx}: inicio deve ser anterior ao fim.")

        dias_inteiros, parcial, horas_adicionais, total_horas_periodo = _segment_breakdown(
            start_dt,
            end_dt,
        )
        tabela = TABELA_DIARIAS.get(tipo, TABELA_DIARIAS["INTERIOR"])
        valor_24h = tabela["24h"]
        valor_parcial = Decimal("0.00")
        if parcial == 15:
            valor_parcial = tabela["15"]
        elif parcial == 30:
            valor_parcial = tabela["30"]

        valor_1_servidor = (valor_24h * dias_inteiros) + valor_parcial
        subtotal = valor_1_servidor * servidores
        data_saida, hora_saida = _format_dt(start_dt)
        data_chegada, hora_chegada = _format_dt(end_dt)

        out = {
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
            "total_horas_periodo": float(total_horas_periodo),
        }
        total_valor_decimal += Decimal(out["subtotal_decimal"])
        total_horas += float(out.get("total_horas_periodo", 0) or 0)
        out.pop("subtotal_decimal", None)
        out.pop("total_horas_periodo", None)
        periodos_out.append(out)

    total_valor_str = _currency(total_valor_decimal)
    valor_por_servidor = (
        (total_valor_decimal / Decimal(servidores)).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
        if servidores > 0
        else total_valor_decimal
    )
    valores_unitarios = [str(item.get("valor_diaria", "") or "").strip() for item in periodos_out]
    valores_unitarios = [item for item in valores_unitarios if item]
    if len(set(valores_unitarios)) == 1:
        valor_unitario_referencia = valores_unitarios[0]
    elif valores_unitarios:
        valor_unitario_referencia = f"{valores_unitarios[0]} (variavel por periodo)"
    else:
        valor_unitario_referencia = ""
    valor_extenso = valor_por_extenso_ptbr(total_valor_str) or ""
    return {
        "periodos": periodos_out,
        "totais": {
            "total_diarias": _total_diarias_resumo(periodos_out),
            "total_horas": round(total_horas, 2),
            "total_valor": total_valor_str,
            "valor_extenso": valor_extenso,
            "quantidade_servidores": servidores,
            "diarias_por_servidor": _total_diarias_resumo(periodos_out),
            "valor_por_servidor": _currency(valor_por_servidor),
            "valor_unitario_referencia": valor_unitario_referencia,
        },
    }
