from datetime import timedelta

from django.test import TestCase
from django.urls import reverse
from django.utils import timezone

from viagens.models import Cidade, Estado, Oficio, Trecho, Viajante


class Step3HorariosTests(TestCase):
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

    def test_oficio_step3_get_rehidrata_trechos_do_banco_quando_sessao_esta_sem_trechos(self) -> None:
        oficio = Oficio.objects.create(status=Oficio.Status.DRAFT)
        data_viagem = timezone.localdate() + timedelta(days=10)
        Trecho.objects.create(
            oficio=oficio,
            ordem=1,
            origem_estado=self.estado_pr,
            origem_cidade=self.cidade_sede,
            destino_estado=self.estado_pr,
            destino_cidade=self.cidade_destino,
            saida_data=data_viagem,
            saida_hora="08:00",
            chegada_data=data_viagem,
            chegada_hora="12:30",
        )

        session = self.client.session
        session["oficio_wizard_id"] = oficio.id
        session["oficio_wizard"] = {
            "oficio": "123/2024",
            "protocolo": "121234567",
            "sede_uf": self.estado_pr.sigla,
            "sede_cidade": str(self.cidade_sede.id),
            "destinos": [
                {
                    "uf": self.estado_pr.sigla,
                    "cidade": str(self.cidade_destino.id),
                }
            ],
            "trechos": [],
            "viajantes_ids": [self.viajante.id],
        }
        session.save()

        response = self.client.get(reverse("oficio_step3"))

        self.assertEqual(response.status_code, 200)
        self.assertContains(response, 'value="08:00"')
        self.assertContains(response, 'value="12:30"')

