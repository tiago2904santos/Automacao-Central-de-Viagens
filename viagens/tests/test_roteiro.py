from datetime import timedelta

from django.test import TestCase
from django.urls import reverse
from django.utils import timezone

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
        self.cidade_final = Cidade.objects.create(
            nome="Londrina", estado=self.estado_pr
        )
        self.viajante = Viajante.objects.create(
            nome="Servidor Teste",
            rg="12345678X",
            cpf="000.000.000-00",
            cargo="Delegado de Policia",
        )

    def _set_wizard_session(self) -> None:
        session = self.client.session
        session["oficio_wizard"] = {
            "oficio": "123/2024",
            "protocolo": "121234567",
            "placa": "ABC1234",
            "modelo": "Uno",
            "combustivel": "Gasolina",
            "viajantes_ids": [self.viajante.id],
        }
        session.save()

    def test_rota_sede_intermediaria_sede_salva_trechos(self) -> None:
        self._set_wizard_session()
        saida_data = timezone.localdate() + timedelta(days=15)
        retorno_data = saida_data + timedelta(days=1)
        payload = {
            "trechos-TOTAL_FORMS": "2",
            "trechos-INITIAL_FORMS": "0",
            "trechos-MIN_NUM_FORMS": "0",
            "trechos-MAX_NUM_FORMS": "1000",
            "trechos-0-origem_estado": self.estado_pr.sigla,
            "trechos-0-origem_cidade": str(self.cidade_sede.id),
            "trechos-0-destino_estado": self.estado_pr.sigla,
            "trechos-0-destino_cidade": str(self.cidade_intermediaria.id),
            "trechos-0-saida_data": saida_data.isoformat(),
            "trechos-0-saida_hora": "07:00",
            "trechos-1-origem_estado": self.estado_pr.sigla,
            "trechos-1-origem_cidade": str(self.cidade_intermediaria.id),
            "retorno_saida_data": retorno_data.isoformat(),
            "retorno_saida_hora": "08:00",
            "retorno_chegada_data": retorno_data.isoformat(),
            "retorno_chegada_hora": "18:00",
            "tipo_destino": "INTERIOR",
            "motivo": "Retorno a sede.",
        }

        response = self.client.post(reverse("oficio_step3"), payload)
        self.assertEqual(response.status_code, 302)
        self.assertEqual(response.url, reverse("oficio_step4"))
        response = self.client.post(reverse("oficio_step4"))
        self.assertEqual(response.status_code, 302)

        self.assertEqual(Oficio.objects.count(), 1)
        oficio = Oficio.objects.first()
        trechos = list(Trecho.objects.filter(oficio=oficio).order_by("ordem"))
        self.assertEqual(len(trechos), 1)
        self.assertEqual(trechos[0].destino_cidade, self.cidade_intermediaria)
        self.assertEqual(
            oficio.retorno_saida_cidade,
            f"{self.cidade_intermediaria.nome}/{self.estado_pr.sigla}",
        )
        self.assertEqual(
            oficio.retorno_chegada_cidade,
            f"{self.cidade_sede.nome}/{self.estado_pr.sigla}",
        )

    def test_rota_com_dois_trechos_salva(self) -> None:
        self._set_wizard_session()
        saida_data = timezone.localdate() + timedelta(days=15)
        retorno_data = saida_data + timedelta(days=1)
        payload = {
            "trechos-TOTAL_FORMS": "2",
            "trechos-INITIAL_FORMS": "0",
            "trechos-MIN_NUM_FORMS": "0",
            "trechos-MAX_NUM_FORMS": "1000",
            "trechos-0-origem_estado": self.estado_pr.sigla,
            "trechos-0-origem_cidade": str(self.cidade_sede.id),
            "trechos-0-destino_estado": self.estado_pr.sigla,
            "trechos-0-destino_cidade": str(self.cidade_intermediaria.id),
            "trechos-0-saida_data": saida_data.isoformat(),
            "trechos-0-saida_hora": "06:30",
            "trechos-1-origem_estado": self.estado_pr.sigla,
            "trechos-1-origem_cidade": str(self.cidade_intermediaria.id),
            "trechos-1-destino_estado": self.estado_pr.sigla,
            "trechos-1-destino_cidade": str(self.cidade_final.id),
            "retorno_saida_data": retorno_data.isoformat(),
            "retorno_saida_hora": "08:00",
            "retorno_chegada_data": retorno_data.isoformat(),
            "retorno_chegada_hora": "20:00",
            "tipo_destino": "INTERIOR",
            "motivo": "Retorno a sede.",
        }

        response = self.client.post(reverse("oficio_step3"), payload)
        self.assertEqual(response.status_code, 302)
        self.assertEqual(response.url, reverse("oficio_step4"))
        response = self.client.post(reverse("oficio_step4"))
        self.assertEqual(response.status_code, 302)

        self.assertEqual(Oficio.objects.count(), 1)
        oficio = Oficio.objects.first()
        trechos = list(Trecho.objects.filter(oficio=oficio).order_by("ordem"))
        self.assertEqual(len(trechos), 2)
        self.assertEqual(trechos[0].destino_cidade, self.cidade_intermediaria)
        self.assertEqual(trechos[1].origem_cidade, self.cidade_intermediaria)
        self.assertEqual(trechos[1].destino_cidade, self.cidade_final)

    def test_rota_sede_intermediaria_sede_sem_placeholder(self) -> None:
        self._set_wizard_session()
        saida_data = timezone.localdate() + timedelta(days=15)
        retorno_data = saida_data + timedelta(days=1)
        payload = {
            "trechos-TOTAL_FORMS": "2",
            "trechos-INITIAL_FORMS": "0",
            "trechos-MIN_NUM_FORMS": "0",
            "trechos-MAX_NUM_FORMS": "1000",
            "trechos-0-origem_estado": self.estado_pr.sigla,
            "trechos-0-origem_cidade": str(self.cidade_sede.id),
            "trechos-0-destino_estado": self.estado_pr.sigla,
            "trechos-0-destino_cidade": str(self.cidade_intermediaria.id),
            "trechos-0-saida_data": saida_data.isoformat(),
            "trechos-0-saida_hora": "07:15",
            "trechos-1-origem_estado": self.estado_pr.sigla,
            "trechos-1-origem_cidade": str(self.cidade_intermediaria.id),
            "retorno_saida_data": retorno_data.isoformat(),
            "retorno_saida_hora": "07:30",
            "retorno_chegada_data": retorno_data.isoformat(),
            "retorno_chegada_hora": "18:00",
            "tipo_destino": "INTERIOR",
            "motivo": "Retorno a sede.",
        }

        response = self.client.post(reverse("oficio_step3"), payload)
        self.assertEqual(response.status_code, 302)
        self.assertEqual(response.url, reverse("oficio_step4"))
        response = self.client.post(reverse("oficio_step4"))
        self.assertEqual(response.status_code, 302)

        self.assertEqual(Oficio.objects.count(), 1)
        trechos = list(Trecho.objects.filter(oficio=Oficio.objects.first()).order_by("ordem"))
        self.assertEqual(len(trechos), 1)
        self.assertEqual(trechos[0].destino_cidade, self.cidade_intermediaria)
        self.assertEqual(
            Oficio.objects.first().retorno_chegada_cidade,
            f"{self.cidade_sede.nome}/{self.estado_pr.sigla}",
        )
