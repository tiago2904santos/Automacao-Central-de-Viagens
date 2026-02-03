import re
import shutil
from pathlib import Path

from django.conf import settings
from django.test import TestCase
from django.urls import reverse

from viagens.models import Cidade, Estado, Oficio, Trecho, Viajante, Veiculo


class OficioFlowTests(TestCase):
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
            cargo="Delegado",
        )
        self.veiculo = Veiculo.objects.create(
            placa="ABC1234",
            modelo="Uno",
            combustivel="Gasolina",
        )

    def _post_step1(self) -> None:
        payload = {
            "oficio": "123/2024",
            "protocolo": "456/2024",
            "assunto": "Teste",
            "servidores": [str(self.viajante.id)],
        }
        response = self.client.post(reverse("formulario"), payload)
        self.assertEqual(response.status_code, 302)

    def _post_step2(self) -> None:
        payload = {
            "placa": "ABC1234",
            "modelo": "Uno",
            "combustivel": "Gasolina",
            "motorista": str(self.viajante.id),
            "motorista_nome": "",
        }
        response = self.client.post(reverse("oficio_step2"), payload)
        self.assertEqual(response.status_code, 302)

    def _post_step3(self) -> None:
        payload = {
            "trechos-TOTAL_FORMS": "2",
            "trechos-INITIAL_FORMS": "0",
            "trechos-MIN_NUM_FORMS": "0",
            "trechos-MAX_NUM_FORMS": "1000",
            "trechos-0-origem_estado": self.estado_pr.sigla,
            "trechos-0-origem_cidade": str(self.cidade_sede.id),
            "trechos-0-destino_estado": self.estado_pr.sigla,
            "trechos-0-destino_cidade": str(self.cidade_intermediaria.id),
            "trechos-0-saida_data": "2024-01-01",
            "trechos-0-saida_hora": "08:00",
            "trechos-1-origem_estado": self.estado_pr.sigla,
            "trechos-1-origem_cidade": str(self.cidade_intermediaria.id),
            "retorno_saida_data": "2024-01-02",
            "retorno_saida_hora": "08:00",
            "retorno_chegada_data": "2024-01-02",
            "retorno_chegada_hora": "18:00",
            "tipo_destino": "INTERIOR",
            "motivo": "Retorno a sede.",
        }
        response = self.client.post(reverse("oficio_step3"), payload)
        self.assertEqual(response.status_code, 302)
        self.assertEqual(response.url, reverse("oficio_step4"))

    def _post_step4(self) -> None:
        response = self.client.post(reverse("oficio_step4"))
        self.assertEqual(response.status_code, 302)

    def test_get_formulario_limpa_session(self) -> None:
        session = self.client.session
        session["oficio_wizard"] = {
            "oficio": "TESTE",
            "protocolo": "TESTE",
            "viajantes_ids": [self.viajante.id],
        }
        session.save()

        response = self.client.get(reverse("formulario"))

        self.assertEqual(response.status_code, 200)
        self.assertEqual(response.context["oficio"], "")
        self.assertEqual(response.context["protocolo"], "")
        self.assertFalse(self.client.session.get("oficio_wizard"))

    def test_fluxo_create_persiste_oficio(self) -> None:
        self._post_step1()
        self._post_step2()
        self._post_step3()
        self._post_step4()

        self.assertEqual(Oficio.objects.count(), 1)
        oficio = Oficio.objects.first()
        self.assertEqual(oficio.oficio, "123/2024")
        self.assertEqual(oficio.veiculo, self.veiculo)
        self.assertEqual(oficio.motorista_viajante, self.viajante)
        self.assertEqual(oficio.viajantes.count(), 1)
        self.assertIsNotNone(oficio.created_at)

        trechos = list(Trecho.objects.filter(oficio=oficio).order_by("ordem"))
        self.assertEqual(len(trechos), 1)
        self.assertEqual(trechos[0].destino_cidade, self.cidade_intermediaria)
        self.assertEqual(
            oficio.retorno_chegada_cidade,
            f"{self.cidade_sede.nome}/{self.estado_pr.sigla}",
        )

    def test_update_oficio(self) -> None:
        oficio = Oficio.objects.create(
            oficio="000/2024",
            protocolo="111",
            destino="Curitiba",
            assunto="Original",
            placa="ABC1234",
            modelo="Uno",
            combustivel="Gasolina",
            motorista="Teste",
            tipo_destino="INTERIOR",
            retorno_saida_data="2024-01-04",
            retorno_saida_hora="07:00",
            retorno_chegada_data="2024-01-05",
            retorno_chegada_hora="18:00",
            motivo="Original",
        )
        oficio.viajantes.add(self.viajante)
        Trecho.objects.create(
            oficio=oficio,
            ordem=1,
            origem_estado=self.estado_pr,
            origem_cidade=self.cidade_sede,
            destino_estado=self.estado_pr,
            destino_cidade=self.cidade_intermediaria,
            saida_data="2024-01-01",
            saida_hora="08:00",
        )

        response = self.client.get(reverse("oficio_edit_step1", args=[oficio.id]))
        self.assertEqual(response.status_code, 200)
        session = self.client.session
        session_key = f"oficio_edit_wizard:{oficio.id}"
        draft = session[session_key]
        trechos_serialized = []
        for trecho in Trecho.objects.filter(oficio=oficio).order_by("ordem"):
            trechos_serialized.append(
                {
                    "origem_estado": trecho.origem_estado.sigla if trecho.origem_estado else "",
                    "origem_cidade": str(trecho.origem_cidade_id or ""),
                    "destino_estado": trecho.destino_estado.sigla if trecho.destino_estado else "",
                    "destino_cidade": str(trecho.destino_cidade_id or ""),
                    "saida_data": trecho.saida_data.isoformat() if trecho.saida_data else "",
                    "saida_hora": trecho.saida_hora.strftime("%H:%M")
                    if trecho.saida_hora
                    else "",
                    "chegada_data": trecho.chegada_data.isoformat() if trecho.chegada_data else "",
                    "chegada_hora": trecho.chegada_hora.strftime("%H:%M")
                    if trecho.chegada_hora
                    else "",
                }
            )
        destinos = [
            {
                "uf": self.estado_pr.sigla,
                "cidade": str(self.cidade_intermediaria.id),
            }
        ]
        draft.update(
            {
                "oficio": "999/2024",
                "protocolo": "999",
                "assunto": "Atualizado",
                "placa": "ABC1234",
                "modelo": "Uno",
                "combustivel": "Diesel",
                "motorista_id": str(self.viajante.id),
                "motorista_oficio": "",
                "motorista_protocolo": "",
                "motivo": "Teste",
                "tipo_destino": "INTERIOR",
                "valor_diarias_extenso": "",
                "destinos": destinos,
                "trechos": trechos_serialized,
                "retorno": {
                    "retorno_saida_data": "2024-01-05",
                    "retorno_saida_hora": "08:00",
                    "retorno_chegada_data": "2024-01-05",
                    "retorno_chegada_hora": "18:00",
                },
            }
        )
        session[session_key] = draft
        session.save()

        response = self.client.post(reverse("oficio_edit_save", args=[oficio.id]))
        self.assertRedirects(
            response, f"{reverse('oficio_edit_step4', args=[oficio.id])}?salvo=1"
        )

        oficio.refresh_from_db()
        self.assertEqual(oficio.oficio, "999/2024")
        self.assertEqual(oficio.combustivel, "Diesel")
        self.assertEqual(oficio.motorista_viajante, self.viajante)
        self.assertIn(self.viajante, oficio.viajantes.all())

    def test_editar_oficio_exibe_data_criacao_sem_input_editavel(self) -> None:
        oficio = Oficio.objects.create(
            oficio="111/2024",
            protocolo="222",
            destino="Curitiba",
            assunto="Teste",
        )

        response = self.client.get(reverse("oficio_edit_step1", args=[oficio.id]))

        self.assertEqual(response.status_code, 200)
        self.assertContains(response, "Data de criacao")
        self.assertNotContains(response, 'name="data"')

    def test_step4_renderiza_resumo_com_session(self) -> None:
        self._post_step1()
        self._post_step2()
        self._post_step3()

        response = self.client.get(reverse("oficio_step4"))

        self.assertEqual(response.status_code, 200)
        self.assertIn("placa", response.context)
        self.assertIn("trechos_summary", response.context)
        self.assertGreater(len(response.context["trechos_summary"]), 0)

    def test_step4_finaliza_salva_oficio_e_limpa_session(self) -> None:
        self._post_step1()
        self._post_step2()
        self._post_step3()

        response = self.client.post(reverse("oficio_step4"))

        self.assertEqual(Oficio.objects.count(), 1)
        self.assertIsNone(self.client.session.get("oficio_wizard"))
        self.assertRedirects(response, reverse("oficios_lista"))
