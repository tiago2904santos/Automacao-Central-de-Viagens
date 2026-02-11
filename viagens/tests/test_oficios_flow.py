import re
import shutil
from pathlib import Path
from unittest.mock import patch

from django.conf import settings
from django.contrib.messages import get_messages
from django.test import TestCase
from django.urls import reverse
from django.utils import timezone

from viagens.documents.document import MotoristaCaronaValidationError
from viagens.models import Cidade, Estado, Oficio, OficioConfig, Trecho, Viajante, Veiculo


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
        payload = {
            "oficio": "123/2024",
            "protocolo": "121234567",
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
        self.assertRegex(response.context["oficio"], r"^\d{2}/\d{4}$")
        self.assertEqual(response.context["protocolo"], "")
        self.assertNotEqual(self.client.session.get("oficio_wizard", {}).get("oficio"), "TESTE")

    def test_fluxo_create_persiste_oficio(self) -> None:
        self._post_step1()
        self._post_step2()
        self._post_step3()
        self._post_step4()

        self.assertEqual(Oficio.objects.count(), 1)
        oficio = Oficio.objects.first()
        self.assertEqual(oficio.oficio, f"01/{timezone.localdate().year}")
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

    def test_destino_automatico_pr(self) -> None:
        oficio = Oficio.objects.create(oficio="200/2024")
        Trecho.objects.create(
            oficio=oficio,
            ordem=1,
            origem_estado=self.estado_pr,
            origem_cidade=self.cidade_sede,
            destino_estado=self.estado_pr,
            destino_cidade=self.cidade_intermediaria,
        )
        oficio.refresh_from_db()
        self.assertEqual(oficio.destino, Oficio.DestinoChoices.GAB)
        self.assertEqual(
            oficio.get_destino_display(),
            "GABINETE DO DELEGADO GERAL ADJUNTO",
        )

    def test_destino_automatico_fora_pr(self) -> None:
        estado_sc = Estado.objects.create(sigla="SC", nome="Santa Catarina")
        cidade_sc = Cidade.objects.create(nome="Florianopolis", estado=estado_sc)
        oficio = Oficio.objects.create(oficio="201/2024")
        Trecho.objects.create(
            oficio=oficio,
            ordem=1,
            origem_estado=self.estado_pr,
            origem_cidade=self.cidade_sede,
            destino_estado=estado_sc,
            destino_cidade=cidade_sc,
        )
        oficio.refresh_from_db()
        self.assertEqual(oficio.destino, Oficio.DestinoChoices.SESP)
        self.assertEqual(oficio.get_destino_display(), "SESP")

    def test_update_oficio(self) -> None:
        oficio = Oficio.objects.create(
            oficio="000/2024",
            protocolo="111",
            destino=Oficio.DestinoChoices.GAB,
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

    def test_edit_step2_get_e_post_nao_quebram(self) -> None:
        oficio = Oficio.objects.create(
            oficio="301/2026",
            protocolo="302/2026",
            assunto="Teste status_context",
            placa="ABC1234",
            modelo="Uno",
            combustivel="Gasolina",
            tipo_destino="INTERIOR",
            retorno_saida_data="2026-02-10",
            retorno_saida_hora="08:00",
            retorno_chegada_data="2026-02-11",
            retorno_chegada_hora="18:00",
            motivo="Teste",
        )
        oficio.viajantes.add(self.viajante)

        response = self.client.get(reverse("oficio_edit_step2", args=[oficio.id]))
        self.assertEqual(response.status_code, 200)

        response = self.client.post(
            reverse("oficio_edit_step2", args=[oficio.id]),
            {
                "placa": "ABC1234",
                "modelo": "Uno",
                "combustivel": "Gasolina",
                "motorista": str(self.viajante.id),
                "motorista_nome": "",
                "motorista_oficio": "",
                "motorista_protocolo": "",
                "goto": "step3",
            },
        )
        self.assertEqual(response.status_code, 302)
        self.assertEqual(response.url, reverse("oficio_edit_step3", args=[oficio.id]))

    def test_edit_step2_persiste_campos_ao_voltar(self) -> None:
        referencia = Oficio.objects.create(
            oficio="REF-01/2026",
            protocolo="REF-PROT-01/2026",
        )
        oficio = Oficio.objects.create(
            oficio="401/2026",
            protocolo="402/2026",
            assunto="Teste persistencia step2",
            placa="ABC1234",
            modelo="Uno",
            combustivel="Gasolina",
            tipo_destino="INTERIOR",
            retorno_saida_data="2026-02-10",
            retorno_saida_hora="08:00",
            retorno_chegada_data="2026-02-11",
            retorno_chegada_hora="18:00",
            motivo="Teste",
        )
        oficio.viajantes.add(self.viajante)

        response = self.client.post(
            reverse("oficio_edit_step2", args=[oficio.id]),
            {
                "placa": "ABC1234",
                "modelo": "Uno",
                "combustivel": "Gasolina",
                "motorista_nome": "Motorista Externo",
                "motorista_oficio": "777/2026",
                "motorista_protocolo": "881234567",
                "carona_oficio_referencia": str(referencia.id),
                "goto": "step3",
            },
        )
        self.assertEqual(response.status_code, 302)
        self.assertEqual(response.url, reverse("oficio_edit_step3", args=[oficio.id]))

        response = self.client.get(reverse("oficio_edit_step2", args=[oficio.id]))
        self.assertEqual(response.status_code, 200)
        self.assertEqual(response.context["motorista_nome"], "MOTORISTA EXTERNO")
        self.assertEqual(response.context["motorista_oficio"], "777/2026")
        self.assertEqual(response.context["motorista_protocolo"], "88.123.456-7")
        self.assertEqual(response.context["carona_oficio_referencia_id"], referencia.id)

    def test_oficio_edit_save_persiste_custeio_e_referencia_carona(self) -> None:
        oficio = Oficio.objects.create(
            oficio="501/2026",
            protocolo="502/2026",
            assunto="Teste save edit wizard",
            placa="ABC1234",
            modelo="Uno",
            combustivel="Gasolina",
            tipo_destino="INTERIOR",
            retorno_saida_data="2026-02-10",
            retorno_saida_hora="08:00",
            retorno_chegada_data="2026-02-11",
            retorno_chegada_hora="18:00",
            motivo="Teste",
        )
        oficio.viajantes.add(self.viajante)
        Trecho.objects.create(
            oficio=oficio,
            ordem=1,
            origem_estado=self.estado_pr,
            origem_cidade=self.cidade_sede,
            destino_estado=self.estado_pr,
            destino_cidade=self.cidade_intermediaria,
            saida_data="2026-02-10",
            saida_hora="08:00",
            chegada_data="2026-02-10",
            chegada_hora="18:00",
        )
        referencia = Oficio.objects.create(oficio="REF-02/2026", protocolo="REF-02")

        response = self.client.get(reverse("oficio_edit_step1", args=[oficio.id]))
        self.assertEqual(response.status_code, 200)

        session = self.client.session
        session_key = f"oficio_edit_wizard:{oficio.id}"
        draft = session[session_key]
        draft.update(
            {
                "oficio": "999/2026",
                "protocolo": "991234567",
                "viajantes_ids": [str(self.viajante.id)],
                "placa": "ABC1234",
                "modelo": "Uno",
                "combustivel": "Gasolina",
                "tipo_viatura": "DESCARACTERIZADA",
                "motorista_id": "",
                "motorista_nome": "Motorista Externo",
                "motorista_oficio": "123/2026",
                "motorista_protocolo": "451234567",
                "motorista_carona": True,
                "carona_oficio_referencia_id": str(referencia.id),
                "custeio_tipo": Oficio.CusteioTipoChoices.OUTRA_INSTITUICAO,
                "nome_instituicao_custeio": "SESP",
                "tipo_destino": "INTERIOR",
                "retorno": {
                    "retorno_saida_data": "2026-02-10",
                    "retorno_saida_hora": "08:00",
                    "retorno_chegada_data": "2026-02-11",
                    "retorno_chegada_hora": "18:00",
                },
                "trechos": [
                    {
                        "origem_estado": self.estado_pr.sigla,
                        "origem_cidade": str(self.cidade_sede.id),
                        "destino_estado": self.estado_pr.sigla,
                        "destino_cidade": str(self.cidade_intermediaria.id),
                        "saida_data": "2026-02-10",
                        "saida_hora": "08:00",
                        "chegada_data": "2026-02-10",
                        "chegada_hora": "18:00",
                    }
                ],
            }
        )
        session[session_key] = draft
        session.save()

        response = self.client.post(reverse("oficio_edit_save", args=[oficio.id]))
        self.assertRedirects(
            response, f"{reverse('oficio_edit_step4', args=[oficio.id])}?salvo=1"
        )

        oficio.refresh_from_db()
        self.assertEqual(oficio.custeio_tipo, Oficio.CusteioTipoChoices.OUTRA_INSTITUICAO)
        self.assertEqual(oficio.nome_instituicao_custeio, "SESP")
        self.assertEqual(oficio.carona_oficio_referencia_id, referencia.id)

    def test_editar_oficio_exibe_data_criacao_sem_input_editavel(self) -> None:
        oficio = Oficio.objects.create(
            oficio="111/2024",
            protocolo="222",
            destino=Oficio.DestinoChoices.GAB,
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

    def test_download_docx_sem_assinante_bloqueia_com_mensagem(self) -> None:
        oficio = Oficio.objects.create(
            oficio="777/2026",
            protocolo="778/2026",
            assunto="Teste download sem assinante",
            placa="ABC1234",
            modelo="Uno",
            combustivel="Gasolina",
            motorista=self.viajante.nome,
            tipo_destino="INTERIOR",
            motivo="Teste",
        )
        oficio.viajantes.add(self.viajante)
        Trecho.objects.create(
            oficio=oficio,
            ordem=1,
            origem_estado=self.estado_pr,
            origem_cidade=self.cidade_sede,
            destino_estado=self.estado_pr,
            destino_cidade=self.cidade_intermediaria,
            saida_data="2026-02-10",
            saida_hora="08:00",
            chegada_data="2026-02-10",
            chegada_hora="12:00",
        )
        cfg = OficioConfig()
        cfg.assinante = None
        cfg.unidade_nome = ""
        cfg.origem_nome = ""

        with patch("viagens.documents.document.get_oficio_config", return_value=cfg):
            response = self.client.get(reverse("oficio_download_docx", args=[oficio.id]))

        self.assertEqual(response.status_code, 302)
        self.assertEqual(response.url, reverse("config_oficio"))
        messages = [m.message for m in get_messages(response.wsgi_request)]
        self.assertTrue(any("assinante" in message.lower() for message in messages))

    def test_download_docx_com_assinante_gera_arquivo(self) -> None:
        oficio = Oficio.objects.create(
            oficio="779/2026",
            protocolo="780/2026",
            assunto="Teste download com assinante",
            placa="ABC1234",
            modelo="Uno",
            combustivel="Gasolina",
            motorista=self.viajante.nome,
            tipo_destino="INTERIOR",
            motivo="Teste",
        )
        oficio.viajantes.add(self.viajante)
        Trecho.objects.create(
            oficio=oficio,
            ordem=1,
            origem_estado=self.estado_pr,
            origem_cidade=self.cidade_sede,
            destino_estado=self.estado_pr,
            destino_cidade=self.cidade_intermediaria,
            saida_data="2026-02-10",
            saida_hora="08:00",
            chegada_data="2026-02-10",
            chegada_hora="12:00",
        )
        cfg = OficioConfig()
        cfg.assinante = self.viajante
        cfg.unidade_nome = "POLICIA CIVIL"
        cfg.origem_nome = "POLICIA CIVIL"

        with patch("viagens.documents.document.get_oficio_config", return_value=cfg):
            response = self.client.get(reverse("oficio_download_docx", args=[oficio.id]))

        self.assertEqual(response.status_code, 200)
        self.assertEqual(
            response["Content-Type"],
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )

    def test_download_pdf_fallback_redireciona_para_docx_sem_500(self) -> None:
        oficio = Oficio.objects.create(
            oficio="781/2026",
            protocolo="782/2026",
            assunto="Teste fallback PDF",
            placa="ABC1234",
            modelo="Uno",
            combustivel="Gasolina",
            motorista=self.viajante.nome,
            tipo_destino="INTERIOR",
            motivo="Teste",
        )
        oficio.viajantes.add(self.viajante)
        Trecho.objects.create(
            oficio=oficio,
            ordem=1,
            origem_estado=self.estado_pr,
            origem_cidade=self.cidade_sede,
            destino_estado=self.estado_pr,
            destino_cidade=self.cidade_intermediaria,
            saida_data="2026-02-10",
            saida_hora="08:00",
            chegada_data="2026-02-10",
            chegada_hora="12:00",
        )

        with patch(
            "viagens.views.build_oficio_docx_and_pdf_bytes",
            side_effect=RuntimeError("boom"),
        ):
            response = self.client.get(reverse("oficio_download_pdf", args=[oficio.id]))

        self.assertEqual(response.status_code, 302)
        self.assertEqual(response.url, reverse("oficio_download_docx", args=[oficio.id]))
        msgs = [m.message for m in get_messages(response.wsgi_request)]
        self.assertTrue(any("Falha ao gerar PDF. Baixe o DOCX." in msg for msg in msgs))

    def test_download_pdf_carona_invalida_redireciona_para_step2(self) -> None:
        oficio = Oficio.objects.create(
            oficio="783/2026",
            protocolo="784/2026",
            assunto="Teste carona inválida",
            placa="ABC1234",
            modelo="Uno",
            combustivel="Gasolina",
            motorista=self.viajante.nome,
            tipo_destino="INTERIOR",
            motivo="Teste",
        )

        with patch(
            "viagens.views.build_oficio_docx_and_pdf_bytes",
            side_effect=MotoristaCaronaValidationError(
                "Informe o Ofício do motorista (carona)."
            ),
        ):
            response = self.client.get(reverse("oficio_download_pdf", args=[oficio.id]))

        self.assertEqual(response.status_code, 302)
        self.assertEqual(response.url, reverse("oficio_edit_step2", args=[oficio.id]))
        msgs = [m.message for m in get_messages(response.wsgi_request)]
        self.assertIn("Informe o Ofício do motorista (carona).", msgs)
