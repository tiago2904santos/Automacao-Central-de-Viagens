from __future__ import annotations

from datetime import datetime
from decimal import Decimal

from django.test import TestCase

from viagens.diarias import PeriodMarker
from viagens.services.diarias_unified import (
    calculate_diarias_from_markers,
    calculate_diarias_from_periods_payload,
    parse_decimal_br,
)


class DiariasServiceUnificadoTests(TestCase):
    def test_calculate_diarias_from_markers_retorna_campos_obrigatorios(self) -> None:
        markers = [
            PeriodMarker(
                saida=datetime(2026, 3, 1, 8, 0),
                destino_cidade="Curitiba",
                destino_uf="PR",
            )
        ]
        resultado = calculate_diarias_from_markers(
            markers=markers,
            chegada_final_sede=datetime(2026, 3, 2, 18, 0),
            total_servidores=1,
        )
        totais = resultado["totais"]
        self.assertIn("valor_unitario", totais)
        self.assertIn("diarias_por_servidor", totais)
        self.assertIn("valor_por_servidor", totais)
        self.assertIn("total_geral", totais)
        self.assertTrue(totais["valor_por_servidor"])
        self.assertTrue(totais["total_geral"])

    def test_calculate_diarias_from_markers_proporcional_por_servidor(self) -> None:
        markers = [
            PeriodMarker(
                saida=datetime(2026, 3, 1, 8, 0),
                destino_cidade="Curitiba",
                destino_uf="PR",
            )
        ]
        resultado = calculate_diarias_from_markers(
            markers=markers,
            chegada_final_sede=datetime(2026, 3, 2, 18, 0),
            total_servidores=3,
        )
        totais = resultado["totais"]
        total_decimal = parse_decimal_br(totais["total_geral"]) or Decimal("0.00")
        por_servidor_decimal = parse_decimal_br(totais["valor_por_servidor"]) or Decimal("0.00")
        self.assertEqual(
            total_decimal,
            (por_servidor_decimal * Decimal("3")).quantize(Decimal("0.01")),
        )

    def test_calculate_diarias_from_periods_payload_retorna_valor_por_servidor(self) -> None:
        periods_payload = [
            {
                "tipo": "INTERIOR",
                "start_date": "2026-03-10",
                "start_time": "08:00",
                "end_date": "2026-03-11",
                "end_time": "18:00",
            }
        ]
        resultado = calculate_diarias_from_periods_payload(
            periods_payload=periods_payload,
            total_servidores=2,
        )
        totais = resultado["totais"]
        self.assertEqual(totais["quantidade_servidores"], 2)
        self.assertIn("valor_por_servidor", totais)
        self.assertIn("total_geral", totais)

