import json
from datetime import date, time

from django.contrib.auth import get_user_model
from django.test import TestCase
from django.urls import reverse

from viagens.models import Cidade, Estado, Roteiro, TrechoRoteiro


User = get_user_model()


class RoteiroJsonFlowTest(TestCase):
    def setUp(self):
        self.user = User.objects.create_user(username="json-user", password="password123")
        self.client.login(username="json-user", password="password123")

        self.estado_pr = Estado.objects.create(sigla="PR", nome="Parana")
        self.curitiba = Cidade.objects.create(nome="Curitiba", estado=self.estado_pr)
        self.maringa = Cidade.objects.create(nome="Maringa", estado=self.estado_pr)
        self.londrina = Cidade.objects.create(nome="Londrina", estado=self.estado_pr)

    def test_roteiro_salvar_persiste_tempo_viagem_e_retorno(self):
        payload = {
            "sede_uf": "PR",
            "sede_cidade": "Curitiba",
            "tempo_viagem": "04:00",
            "destinos": [{"uf": "PR", "cidade": "Maringa"}],
            "trechos": [
                {
                    "origem_estado": "PR",
                    "origem_cidade": "Curitiba",
                    "destino_estado": "PR",
                    "destino_cidade": "Maringa",
                    "saida_data": "2026-03-15",
                    "saida_hora": "08:00",
                    "chegada_data": "",
                    "chegada_hora": "",
                }
            ],
            "retorno": {
                "saida_data": "2026-03-16",
                "saida_hora": "09:00",
                "chegada_data": "",
                "chegada_hora": "",
            },
        }

        response = self.client.post(
            reverse("roteiro_salvar"),
            data=json.dumps(payload),
            content_type="application/json",
            HTTP_X_REQUESTED_WITH="XMLHttpRequest",
        )

        self.assertEqual(response.status_code, 201)
        data = response.json()
        self.assertTrue(data["ok"])
        self.assertIn("/roteiros/", data["redirect_url"])

        roteiro = Roteiro.objects.get(pk=data["roteiro_id"])
        trecho = roteiro.trechos.get()
        self.assertEqual(roteiro.nome, "Curitiba > Maringa 15/03/2026 08:00")
        self.assertEqual(roteiro.retorno_saida_cidade, "Maringa")
        self.assertEqual(roteiro.tempo_viagem.strftime("%H:%M"), "04:00")
        self.assertEqual(trecho.tempo_viagem_minutos, 240)
        self.assertEqual(trecho.saida_hora.strftime("%H:%M"), "08:00")
        self.assertEqual(trecho.chegada_hora.strftime("%H:%M"), "12:00")
        self.assertEqual(roteiro.retorno_saida_hora.strftime("%H:%M"), "09:00")
        self.assertEqual(trecho.retorno_saida_hora.strftime("%H:%M"), "09:00")
        self.assertEqual(trecho.retorno_chegada_hora.strftime("%H:%M"), "13:00")

    def test_roteiro_editar_get_injeta_json_inicial(self):
        roteiro = Roteiro.objects.create(
            nome="(gerando...)",
            estado_sede=self.estado_pr,
            cidade_sede=self.curitiba,
            uf_origem="PR",
            cidade_origem="Curitiba",
            uf_destino="PR",
            cidade_destino="Maringa",
            tipo_deslocamento=Roteiro.TipoDeslocamentoChoices.INTERIOR,
            tempo_viagem=time(4, 0),
        )
        TrechoRoteiro.objects.create(
            roteiro=roteiro,
            ordem=1,
            origem_estado=self.estado_pr,
            origem_cidade=self.curitiba,
            destino_estado=self.estado_pr,
            destino_cidade=self.maringa,
            saida_data=date(2026, 3, 15),
        )

        response = self.client.get(reverse("roteiro_editar", args=[roteiro.pk]))

        self.assertEqual(response.status_code, 200)
        self.assertContains(response, "Editar Roteiro")
        self.assertContains(response, "const ROTEIRO_INICIAL =")
        self.assertContains(response, 'const SEDE_UF_INICIAL = "PR";')
        self.assertContains(response, "const RETORNO_INICIAL =")
        self.assertContains(response, 'const TEMPO_VIAGEM_INICIAL = "04:00";')
        self.assertNotContains(response, "Tipo de Deslocamento")

    def test_roteiro_editar_post_json_recria_trechos_sem_duplicar(self):
        roteiro = Roteiro.objects.create(
            nome="(gerando...)",
            estado_sede=self.estado_pr,
            cidade_sede=self.curitiba,
            uf_origem="PR",
            cidade_origem="Curitiba",
            uf_destino="PR",
            cidade_destino="Maringa",
            tipo_deslocamento=Roteiro.TipoDeslocamentoChoices.INTERIOR,
        )
        TrechoRoteiro.objects.create(
            roteiro=roteiro,
            ordem=1,
            origem_estado=self.estado_pr,
            origem_cidade=self.curitiba,
            destino_estado=self.estado_pr,
            destino_cidade=self.maringa,
        )
        TrechoRoteiro.objects.create(
            roteiro=roteiro,
            ordem=2,
            origem_estado=self.estado_pr,
            origem_cidade=self.maringa,
            destino_estado=self.estado_pr,
            destino_cidade=self.curitiba,
        )

        payload = {
            "sede_uf": "PR",
            "sede_cidade": "Curitiba",
            "tempo_viagem": "03:30",
            "destinos": [{"uf": "PR", "cidade": "Londrina"}],
            "trechos": [
                {
                    "origem_estado": "PR",
                    "origem_cidade": "Curitiba",
                    "destino_estado": "PR",
                    "destino_cidade": "Londrina",
                    "saida_data": "2026-04-01",
                    "saida_hora": "07:30",
                    "chegada_data": "",
                    "chegada_hora": "",
                }
            ],
        }

        response = self.client.post(
            reverse("roteiro_editar", args=[roteiro.pk]),
            data=json.dumps(payload),
            content_type="application/json",
            HTTP_X_REQUESTED_WITH="XMLHttpRequest",
        )

        self.assertEqual(response.status_code, 200)
        roteiro.refresh_from_db()
        self.assertEqual(roteiro.nome, "Curitiba > Londrina 01/04/2026 07:30")
        self.assertEqual(roteiro.tempo_viagem.strftime("%H:%M"), "03:30")
        self.assertEqual(roteiro.trechos.count(), 1)
        trecho = roteiro.trechos.get()
        self.assertEqual(trecho.destino_cidade.nome, "Londrina")
        self.assertEqual(trecho.tempo_viagem_minutos, 210)
        self.assertEqual(trecho.chegada_hora.strftime("%H:%M"), "11:00")
