from __future__ import annotations

import json

from django.test import TestCase
from django.urls import reverse


class SimulacaoDiariasPorServidorTests(TestCase):
    def test_simulacao_retorna_detalhamento_por_servidor(self) -> None:
        periods = [
            {
                "tipo": "INTERIOR",
                "start_date": "2026-03-10",
                "start_time": "08:00",
                "end_date": "2026-03-11",
                "end_time": "18:00",
            }
        ]
        response = self.client.post(
            reverse("simulacao_diarias_calcular"),
            {
                "quantidade_servidores": "2",
                "periods_payload": json.dumps(periods),
            },
        )
        self.assertEqual(response.status_code, 200)
        payload = response.json()
        totais = payload["totais"]
        self.assertIn("valor_por_servidor", totais)
        self.assertIn("diarias_por_servidor", totais)
        self.assertIn("valor_unitario_referencia", totais)
        self.assertEqual(totais["quantidade_servidores"], 2)
