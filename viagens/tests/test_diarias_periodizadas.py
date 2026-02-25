from datetime import timedelta

from django.test import TestCase
from django.urls import reverse
from django.utils import timezone

from viagens.models import Cidade, Estado, Oficio, Viajante, Veiculo


class DiariasPeriodizadasTests(TestCase):
    def setUp(self) -> None:
        self.estado_pr = Estado.objects.create(sigla="PR", nome="Parana")
        self.estado_sp = Estado.objects.create(sigla="SP", nome="Sao Paulo")
        self.cidade_curitiba = Cidade.objects.create(nome="Curitiba", estado=self.estado_pr)
        self.cidade_paranagua = Cidade.objects.create(nome="Paranagua", estado=self.estado_pr)
        self.cidade_sao_paulo = Cidade.objects.create(nome="Sao Paulo", estado=self.estado_sp)
        self.viajante = Viajante.objects.create(
            nome="Servidor Teste",
            rg="12345678X",
            cpf="000.000.000-00",
            cargo="Delegado",
        )
        self.veiculo = Veiculo.objects.create(
            placa="ABC1234",
            modelo="Uno",
            combustivel="Gasolina",
        )

    def _step1(self) -> None:
        response = self.client.post(
            reverse("formulario"),
            {
                "oficio": "01/2026",
                "protocolo": "121234567",
                "servidores": [str(self.viajante.id)],
            },
        )
        self.assertEqual(response.status_code, 302)

    def _step2(self) -> None:
        response = self.client.post(
            reverse("oficio_step2"),
            {
                "placa": self.veiculo.placa,
                "modelo": self.veiculo.modelo,
                "combustivel": self.veiculo.combustivel,
                "motorista": str(self.viajante.id),
                "motorista_nome": "",
            },
        )
        self.assertEqual(response.status_code, 302)

    def test_step3_salva_calculo_periodizado_sem_regra_de_maior_valor(self) -> None:
        self._step1()
        self._step2()
        saida_data_1 = timezone.localdate() + timedelta(days=20)
        saida_data_2 = saida_data_1 + timedelta(days=2)
        retorno_data = saida_data_1 + timedelta(days=4)

        response = self.client.post(
            reverse("oficio_step3"),
            {
                "trechos-TOTAL_FORMS": "2",
                "trechos-INITIAL_FORMS": "0",
                "trechos-MIN_NUM_FORMS": "0",
                "trechos-MAX_NUM_FORMS": "1000",
                "trechos-0-origem_estado": self.estado_pr.sigla,
                "trechos-0-origem_cidade": str(self.cidade_curitiba.id),
                "trechos-0-destino_estado": self.estado_pr.sigla,
                "trechos-0-destino_cidade": str(self.cidade_paranagua.id),
                "trechos-0-saida_data": saida_data_1.isoformat(),
                "trechos-0-saida_hora": "08:00",
                "trechos-1-origem_estado": self.estado_pr.sigla,
                "trechos-1-origem_cidade": str(self.cidade_paranagua.id),
                "trechos-1-destino_estado": self.estado_sp.sigla,
                "trechos-1-destino_cidade": str(self.cidade_sao_paulo.id),
                "trechos-1-saida_data": saida_data_2.isoformat(),
                "trechos-1-saida_hora": "08:00",
                "retorno_saida_data": retorno_data.isoformat(),
                "retorno_saida_hora": "09:00",
                "retorno_chegada_data": retorno_data.isoformat(),
                "retorno_chegada_hora": "18:00",
                "motivo": "Teste calculo periodizado.",
            },
        )

        self.assertEqual(response.status_code, 302)
        self.assertEqual(response.url, reverse("oficio_step4"))

        oficio = Oficio.objects.get()
        self.assertEqual(oficio.quantidade_diarias, "4 x 100% + 1 x 30%")
        self.assertEqual(oficio.valor_diarias, "1435,00")

    def test_endpoint_calcular_diarias_retorna_periodos_segmentados(self) -> None:
        session = self.client.session
        session["oficio_wizard"] = {"viajantes_ids": [str(self.viajante.id)]}
        session.save()

        response = self.client.post(
            reverse("oficio_calcular_diarias"),
            {
                "trechos-TOTAL_FORMS": "2",
                "trechos-INITIAL_FORMS": "0",
                "trechos-MIN_NUM_FORMS": "0",
                "trechos-MAX_NUM_FORMS": "1000",
                "trechos-0-origem_estado": self.estado_pr.sigla,
                "trechos-0-origem_cidade": str(self.cidade_curitiba.id),
                "trechos-0-destino_estado": self.estado_pr.sigla,
                "trechos-0-destino_cidade": str(self.cidade_paranagua.id),
                "trechos-0-saida_data": "2026-02-10",
                "trechos-0-saida_hora": "08:00",
                "trechos-1-origem_estado": self.estado_pr.sigla,
                "trechos-1-origem_cidade": str(self.cidade_paranagua.id),
                "trechos-1-destino_estado": self.estado_sp.sigla,
                "trechos-1-destino_cidade": str(self.cidade_sao_paulo.id),
                "trechos-1-saida_data": "2026-02-12",
                "trechos-1-saida_hora": "08:00",
                "retorno_chegada_data": "2026-02-14",
                "retorno_chegada_hora": "18:00",
            },
        )

        self.assertEqual(response.status_code, 200)
        payload = response.json()
        self.assertEqual(len(payload["periodos"]), 2)
        self.assertEqual(payload["periodos"][0]["tipo"], "INTERIOR")
        self.assertEqual(payload["periodos"][1]["tipo"], "CAPITAL")
        self.assertEqual(payload["totais"]["total_diarias"], "4 x 100% + 1 x 30%")
        self.assertEqual(payload["totais"]["total_valor"], "1435,00")

    def test_endpoint_calcular_diarias_retorna_erro_quando_falta_data(self) -> None:
        session = self.client.session
        session["oficio_wizard"] = {"viajantes_ids": [str(self.viajante.id)]}
        session.save()

        response = self.client.post(
            reverse("oficio_calcular_diarias"),
            {
                "trechos-TOTAL_FORMS": "1",
                "trechos-INITIAL_FORMS": "0",
                "trechos-MIN_NUM_FORMS": "0",
                "trechos-MAX_NUM_FORMS": "1000",
                "trechos-0-origem_estado": self.estado_pr.sigla,
                "trechos-0-origem_cidade": str(self.cidade_curitiba.id),
                "trechos-0-destino_estado": self.estado_pr.sigla,
                "trechos-0-destino_cidade": str(self.cidade_paranagua.id),
                "trechos-0-saida_data": "",
                "trechos-0-saida_hora": "",
                "retorno_chegada_data": "2026-02-14",
                "retorno_chegada_hora": "18:00",
            },
        )

        self.assertEqual(response.status_code, 400)
        self.assertIn("Preencha datas e horas para calcular.", response.json().get("error", ""))
