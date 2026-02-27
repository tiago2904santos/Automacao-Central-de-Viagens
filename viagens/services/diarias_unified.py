from __future__ import annotations

from datetime import datetime
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP

from viagens.diarias import PeriodMarker, calculate_periodized_diarias
from viagens.simulacao import calculate_periods_from_payload
from viagens.services.oficio_helpers import valor_por_extenso_ptbr


def parse_decimal_br(value: str | Decimal | None) -> Decimal | None:
    if value in (None, ""):
        return None
    if isinstance(value, Decimal):
        return value
    normalized = str(value).strip().replace("R$", "").replace("r$", "").replace(" ", "")
    if not normalized:
        return None
    # Handle strings like "290,55 (variavel por periodo)".
    normalized = normalized.split("(", 1)[0].strip()
    try:
        if "," in normalized and "." in normalized:
            if normalized.rfind(",") > normalized.rfind("."):
                normalized = normalized.replace(".", "").replace(",", ".")
            else:
                normalized = normalized.replace(",", "")
        elif "," in normalized:
            normalized = normalized.replace(".", "").replace(",", ".")
        else:
            normalized = normalized.replace(",", "")
        return Decimal(normalized)
    except (InvalidOperation, TypeError, ValueError):
        return None


def _format_decimal_br(value: Decimal | None) -> str:
    if value is None:
        return ""
    quantized = value.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    return f"{quantized:.2f}".replace(".", ",")


def _normalize_diarias_resultado(resultado: dict) -> dict:
    totais = resultado.get("totais", {}) if isinstance(resultado, dict) else {}
    total_geral = str(totais.get("total_geral") or totais.get("total_valor") or "").strip()
    diarias_por_servidor = str(
        totais.get("diarias_por_servidor")
        or totais.get("quantidade_diarias_por_servidor")
        or totais.get("total_diarias")
        or ""
    ).strip()
    valor_unitario = str(
        totais.get("valor_unitario")
        or totais.get("valor_unitario_referencia")
        or ""
    ).strip()
    valor_por_servidor = str(totais.get("valor_por_servidor") or "").strip()
    servidores = int(totais.get("quantidade_servidores") or 0)

    total_decimal = parse_decimal_br(total_geral) or Decimal("0.00")
    if not valor_por_servidor and servidores > 0:
        valor_por_servidor = _format_decimal_br(
            (total_decimal / Decimal(servidores)).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
        )
    if not valor_unitario:
        valor_unitario = valor_por_servidor

    totais.update(
        {
            "total_geral": total_geral,
            "total_valor": total_geral,
            "valor_total": total_geral,
            "diarias_por_servidor": diarias_por_servidor,
            "quantidade_diarias_por_servidor": diarias_por_servidor,
            "valor_por_servidor": valor_por_servidor,
            "valor_unitario": valor_unitario,
            "valor_unitario_referencia": valor_unitario,
        }
    )
    resultado["totais"] = totais
    return resultado


def derive_financeiro_diarias(resultado: dict | None) -> dict[str, str]:
    totais = resultado.get("totais", {}) if isinstance(resultado, dict) else {}
    return {
        "diarias_por_servidor": str(
            totais.get("diarias_por_servidor")
            or totais.get("quantidade_diarias_por_servidor")
            or ""
        ).strip(),
        "valor_unitario": str(
            totais.get("valor_unitario")
            or totais.get("valor_unitario_referencia")
            or ""
        ).strip(),
        "valor_por_servidor": str(totais.get("valor_por_servidor") or "").strip(),
        "total_geral": str(totais.get("total_geral") or totais.get("total_valor") or "").strip(),
        "valor_extenso": str(totais.get("valor_extenso") or "").strip(),
    }


def calculate_diarias_from_markers(
    *,
    markers: list[PeriodMarker],
    chegada_final_sede: datetime,
    total_servidores: int,
) -> dict:
    servidores = int(total_servidores or 0)
    if servidores < 1:
        raise ValueError("Preencha o efetivo para calcular as diarias.")
    resultado = calculate_periodized_diarias(
        markers,
        chegada_final_sede,
        quantidade_servidores=servidores,
        valor_extenso_fn=valor_por_extenso_ptbr,
    )
    return _normalize_diarias_resultado(resultado)


def calculate_diarias_from_periods_payload(
    *,
    periods_payload: list[dict],
    total_servidores: int,
) -> dict:
    servidores = int(total_servidores or 0)
    if servidores < 1:
        raise ValueError("Preencha o efetivo para calcular as diarias.")
    resultado = calculate_periods_from_payload(
        periods_payload,
        quantidade_servidores=servidores,
    )
    return _normalize_diarias_resultado(resultado)

