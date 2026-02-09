from django.test import TestCase

from viagens.models import Cargo, Viajante
from viagens.views import _get_cargo_choices, _normalizar_cargo_key


class CargoChoicesTests(TestCase):
    def setUp(self) -> None:
        Cargo.objects.get_or_create(nome="Assessor de Comunicação Social")
        Viajante.objects.create(
            nome="Ana Silva",
            rg="123456789",
            cpf="12345678901",
            cargo="Agente de Polícia Judiciária",
        )
        Viajante.objects.create(
            nome="Bruno Lima",
            rg="987654321",
            cpf="10987654321",
            cargo="agente de polícia judiciária",
        )
        Viajante.objects.create(
            nome="Carla Rocha",
            rg="111222333",
            cpf="10987650123",
            cargo="Delegado",
        )

    def test_choices_deduplicated_and_sorted(self):
        choices = _get_cargo_choices()
        normalized = {_normalizar_cargo_key(item) for item in choices}
        self.assertEqual(len(choices), len(normalized))
        self.assertIn("Assessor de Comunicação Social", choices)
        self.assertEqual(choices, sorted(choices, key=lambda value: value.casefold()))
