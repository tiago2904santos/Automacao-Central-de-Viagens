from __future__ import annotations

from datetime import timedelta
from urllib.parse import parse_qs, urlparse

from django.test import TestCase
from django.urls import reverse
from django.utils import timezone

from viagens.models import Cidade, Estado, Oficio, Viajante, Veiculo
from viagens.services.justificativas import requires_justificativa


class OficioJustificativaFlowTests(TestCase):
    def setUp(self) -> None:
        self.estado_pr = Estado.objects.create(sigla="PR", nome="Parana")
        self.cidade_sede = Cidade.objects.create(nome="Curitiba", estado=self.estado_pr)
        self.cidade_destino = Cidade.objects.create(nome="Maringa", estado=self.estado_pr)
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

    def _post_step1(self) -> None:
        response = self.client.post(
            reverse("formulario"),
            {
                "oficio": "123/2026",
                "protocolo": "121234567",
                "assunto": "Teste",
                "servidores": [str(self.viajante.id)],
            },
        )
        self.assertEqual(response.status_code, 302)

    def _post_step2(self) -> None:
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

    def _step3_payload(self, *, dias_ate_primeira_saida: int) -> dict[str, str]:
        saida_data = timezone.localdate() + timedelta(days=dias_ate_primeira_saida)
        retorno_data = saida_data + timedelta(days=1)
        return {
            "trechos-TOTAL_FORMS": "2",
            "trechos-INITIAL_FORMS": "0",
            "trechos-MIN_NUM_FORMS": "0",
            "trechos-MAX_NUM_FORMS": "1000",
            "trechos-0-origem_estado": self.estado_pr.sigla,
            "trechos-0-origem_cidade": str(self.cidade_sede.id),
            "trechos-0-destino_estado": self.estado_pr.sigla,
            "trechos-0-destino_cidade": str(self.cidade_destino.id),
            "trechos-0-saida_data": saida_data.isoformat(),
            "trechos-0-saida_hora": "08:00",
            "trechos-1-origem_estado": self.estado_pr.sigla,
            "trechos-1-origem_cidade": str(self.cidade_destino.id),
            "retorno_saida_data": retorno_data.isoformat(),
            "retorno_saida_hora": "08:00",
            "retorno_chegada_data": retorno_data.isoformat(),
            "retorno_chegada_hora": "18:00",
            "tipo_destino": "INTERIOR",
            "motivo": "Teste justificativa.",
        }

    def test_step3_redireciona_para_justificativa_quando_antecedencia_menor_10(self) -> None:
        self._post_step1()
        self._post_step2()

        response = self.client.post(
            reverse("oficio_step3"),
            self._step3_payload(dias_ate_primeira_saida=5),
        )

        self.assertEqual(response.status_code, 302)
        oficio = Oficio.objects.get()
        parsed = urlparse(response.url)
        self.assertEqual(parsed.path, reverse("oficio_justificativa", args=[oficio.id]))
        self.assertEqual(
            parse_qs(parsed.query).get("next", [""])[0],
            reverse("oficio_step4"),
        )

    def test_salvar_justificativa_permite_ir_para_step4_e_finalizar(self) -> None:
        self._post_step1()
        self._post_step2()

        response_step3 = self.client.post(
            reverse("oficio_step3"),
            self._step3_payload(dias_ate_primeira_saida=5),
        )
        self.assertEqual(response_step3.status_code, 302)
        oficio = Oficio.objects.get()

        response_justificativa = self.client.post(
            reverse("oficio_justificativa", args=[oficio.id]),
            {
                "justificativa_modelo": "evento",
                "justificativa_texto": "Justificativa de teste.",
                "next": reverse("oficio_step4"),
            },
        )
        self.assertEqual(response_justificativa.status_code, 302)
        self.assertEqual(response_justificativa.url, reverse("oficio_step4"))

        response_step4 = self.client.post(reverse("oficio_step4"))
        self.assertEqual(response_step4.status_code, 302)
        self.assertEqual(response_step4.url, reverse("oficios_lista"))

        oficio.refresh_from_db()
        self.assertEqual(oficio.status, Oficio.Status.FINAL)
        self.assertEqual(oficio.justificativa_modelo, "evento")
        self.assertEqual(oficio.justificativa_texto, "Justificativa de teste.")

    def test_step3_nao_exige_justificativa_quando_antecedencia_maior_ou_igual_10(self) -> None:
        self._post_step1()
        self._post_step2()

        response = self.client.post(
            reverse("oficio_step3"),
            self._step3_payload(dias_ate_primeira_saida=15),
        )

        self.assertEqual(response.status_code, 302)
        self.assertEqual(response.url, reverse("oficio_step4"))

    def test_helper_nao_exige_sem_saida_data(self) -> None:
        self.assertFalse(requires_justificativa(trechos_payload=[{"saida_data": ""}]))
