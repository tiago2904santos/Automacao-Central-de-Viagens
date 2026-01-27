from django.test import TestCase
from django.urls import reverse

from viagens.models import Cidade, Estado, Oficio, Trecho, Viajante


class RoteiroFormsetTests(TestCase):
    def setUp(self) -> None:
        self.estado_pr = Estado.objects.create(sigla="PR", nome="Parana")
        self.cidade_sede = Cidade.objects.create(
            nome="Curitiba", estado=self.estado_pr
        )
        self.cidade_intermediaria = Cidade.objects.create(
            nome="Maringa", estado=self.estado_pr
        )
        self.viajante = Viajante.objects.create(
            nome="Servidor Teste",
            rg="123456",
            cpf="000.000.000-00",
            cargo="Delegado de Policia",
        )

    def _set_wizard_session(self) -> None:
        session = self.client.session
        session["oficio_wizard"] = {
            "oficio": "123/2024",
            "protocolo": "456/2024",
            "placa": "ABC1234",
            "modelo": "Uno",
            "combustivel": "Gasolina",
            "viajantes_ids": [self.viajante.id],
        }
        session.save()

    def test_rota_sede_intermediaria_sede_salva_trechos(self) -> None:
        self._set_wizard_session()
        payload = {
            "trechos-TOTAL_FORMS": "3",
            "trechos-INITIAL_FORMS": "0",
            "trechos-MIN_NUM_FORMS": "0",
            "trechos-MAX_NUM_FORMS": "1000",
            "trechos-0-origem_estado": self.estado_pr.sigla,
            "trechos-0-origem_cidade": str(self.cidade_sede.id),
            "trechos-0-destino_estado": self.estado_pr.sigla,
            "trechos-0-destino_cidade": str(self.cidade_intermediaria.id),
            "trechos-1-origem_estado": self.estado_pr.sigla,
            "trechos-1-origem_cidade": str(self.cidade_intermediaria.id),
            "trechos-1-destino_estado": self.estado_pr.sigla,
            "trechos-1-destino_cidade": str(self.cidade_sede.id),
            "trechos-2-origem_estado": self.estado_pr.sigla,
            "trechos-2-origem_cidade": str(self.cidade_sede.id),
            "motivo": "Retorno a sede.",
        }

        response = self.client.post(reverse("oficio_step3"), payload)

        self.assertEqual(response.status_code, 200)
        self.assertEqual(Oficio.objects.count(), 1)
        oficio = Oficio.objects.first()
        trechos = list(Trecho.objects.filter(oficio=oficio).order_by("ordem"))
        self.assertEqual(len(trechos), 2)
        self.assertEqual(trechos[0].destino_cidade, self.cidade_intermediaria)
        self.assertEqual(trechos[1].origem_cidade, self.cidade_intermediaria)
        self.assertEqual(trechos[1].destino_cidade, self.cidade_sede)
