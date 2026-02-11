from datetime import datetime
from django.test import TestCase
from django.urls import reverse
from viagens.models import Cidade, Estado, Oficio, Trecho, Viajante


class Step3BackendContractTests(TestCase):
    def setUp(self) -> None:
        self.estado_pr = Estado.objects.create(sigla="PR", nome="Parana")
        self.cidade_sede = Cidade.objects.create(nome="Curitiba", estado=self.estado_pr)
        self.cidade_destino = Cidade.objects.create(nome="Cascavel", estado=self.estado_pr)

        self.viajante = Viajante.objects.create(
            nome="Servidor Teste",
            rg="12345678X",
            cpf="000.000.000-00",
            cargo="Delegado",
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

    def test_step3_permitem_trechos_total_forms_1_sem_placeholder(self) -> None:
        """
        Requisito: NÃO exigir trecho placeholder.
        Com apenas 1 trecho de ida, backend salva 1 Trecho e seta retorno (destino = sede).
        """
        self._set_wizard_session()

        payload = {
            "trechos-TOTAL_FORMS": "1",
            "trechos-INITIAL_FORMS": "0",
            "trechos-MIN_NUM_FORMS": "0",
            "trechos-MAX_NUM_FORMS": "1000",
            "trechos-0-origem_estado": self.estado_pr.sigla,
            "trechos-0-origem_cidade": str(self.cidade_sede.id),
            "trechos-0-destino_estado": self.estado_pr.sigla,
            "trechos-0-destino_cidade": str(self.cidade_destino.id),
            "trechos-0-saida_data": "2024-01-01",
            "trechos-0-saida_hora": "08:00",
            # retorno (1 trecho único)
            "retorno_saida_data": "2024-01-02",
            "retorno_saida_hora": "09:00",
            "retorno_chegada_data": "2024-01-02",
            "retorno_chegada_hora": "18:00",
            "tipo_destino": "INTERIOR",
            "motivo": "Teste retorno.",
        }

        resp = self.client.post(reverse("oficio_step3"), payload)
        self.assertEqual(resp.status_code, 302)
        self.assertEqual(resp.url, reverse("oficio_step4"))
        resp = self.client.post(reverse("oficio_step4"))
        self.assertEqual(resp.status_code, 302)

        self.assertEqual(Oficio.objects.count(), 1)
        oficio = Oficio.objects.first()

        trechos = list(Trecho.objects.filter(oficio=oficio).order_by("ordem"))
        self.assertEqual(len(trechos), 1)
        self.assertEqual(trechos[0].origem_cidade, self.cidade_sede)
        self.assertEqual(trechos[0].destino_cidade, self.cidade_destino)

        # retorno: origem = último destino, chegada = sede
        self.assertEqual(
            oficio.retorno_saida_cidade,
            f"{self.cidade_destino.nome}/{self.estado_pr.sigla}",
        )
        self.assertEqual(
            oficio.retorno_chegada_cidade,
            f"{self.cidade_sede.nome}/{self.estado_pr.sigla}",
        )
