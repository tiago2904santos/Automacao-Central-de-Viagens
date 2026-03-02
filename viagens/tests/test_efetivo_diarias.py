from __future__ import annotations

from datetime import date
from decimal import Decimal

from django.test import TestCase
from django.urls import reverse

from viagens.diarias import calculate_simple_diarias_total
from viagens.models import Cargo, Efetivo


class EfetivoViewsTests(TestCase):
    def setUp(self) -> None:
        self.cargo_delegado = Cargo.objects.create(nome="Delegado", ordem=1, ativo=True)
        self.cargo_coord = Cargo.objects.create(
            nome="Coordenador de Operacao",
            ordem=2,
            ativo=True,
            is_coordenador=True,
        )

    def test_get_efetivo_cria_registros_e_lista_cargos(self) -> None:
        response = self.client.get(reverse("efetivo"))

        self.assertEqual(response.status_code, 200)
        self.assertContains(response, "Delegado")
        self.assertContains(response, "Coordenador de Operacao")
        self.assertTrue(Efetivo.objects.filter(cargo=self.cargo_delegado).exists())
        self.assertTrue(Efetivo.objects.filter(cargo=self.cargo_coord).exists())

    def test_post_efetivo_salva_quantidades_e_permite_novo_cargo(self) -> None:
        self.client.get(reverse("efetivo"))
        payload = {
            "efetivo-TOTAL_FORMS": "2",
            "efetivo-INITIAL_FORMS": "2",
            "efetivo-MIN_NUM_FORMS": "0",
            "efetivo-MAX_NUM_FORMS": "1000",
            "efetivo-0-cargo_id": str(self.cargo_delegado.id),
            "efetivo-0-cargo_nome": self.cargo_delegado.nome,
            "efetivo-0-is_coordenador": "",
            "efetivo-0-quantidade": "3",
            "efetivo-1-cargo_id": str(self.cargo_coord.id),
            "efetivo-1-cargo_nome": self.cargo_coord.nome,
            "efetivo-1-is_coordenador": "True",
            "efetivo-1-quantidade": "1",
            "novo-nome_cargo": "Papiloscopista",
            "novo-is_coordenador": "",
        }

        response = self.client.post(reverse("efetivo"), payload)

        self.assertEqual(response.status_code, 302)
        self.assertEqual(response.headers["Location"], reverse("efetivo"))
        self.assertEqual(Efetivo.objects.get(cargo=self.cargo_delegado).quantidade, 3)
        self.assertEqual(Efetivo.objects.get(cargo=self.cargo_coord).quantidade, 1)
        novo_cargo = Cargo.objects.get(nome="Papiloscopista")
        self.assertTrue(novo_cargo.ativo)
        self.assertFalse(novo_cargo.is_coordenador)
        self.assertTrue(Efetivo.objects.filter(cargo=novo_cargo).exists())


class DiariasCalculadoraTests(TestCase):
    def test_calculate_simple_diarias_total_aplica_meia_diaria(self) -> None:
        dias, total = calculate_simple_diarias_total(
            date(2026, 3, 1),
            date(2026, 3, 3),
            Decimal("100.00"),
            meia_diaria=True,
        )

        self.assertEqual(dias, 3)
        self.assertEqual(total, Decimal("250.00"))

    def test_view_diarias_exibe_resultado(self) -> None:
        response = self.client.post(
            reverse("diarias"),
            {
                "data_saida": "2026-03-01",
                "data_retorno": "2026-03-03",
                "valor_diaria": "100.00",
                "meia_diaria": "on",
            },
        )

        self.assertEqual(response.status_code, 200)
        self.assertContains(response, "R$ 250,00")
        self.assertContains(response, ">3<")

    def test_view_diarias_valida_data_retorno(self) -> None:
        response = self.client.post(
            reverse("diarias"),
            {
                "data_saida": "2026-03-03",
                "data_retorno": "2026-03-01",
                "valor_diaria": "100.00",
            },
        )

        self.assertEqual(response.status_code, 200)
        self.assertContains(response, "deve ser igual ou posterior")
