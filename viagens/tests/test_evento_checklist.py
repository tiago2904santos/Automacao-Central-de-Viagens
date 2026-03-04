# viagens/tests/test_evento_checklist.py
"""Testes do gerenciador por evento (checklist e pronto_para_protocolar)."""
from __future__ import annotations

from datetime import time, timedelta

from django.test import TestCase
from django.urls import reverse
from django.utils import timezone

from viagens.models import Cidade, Estado, Evento, Oficio, Trecho, Viajante
from viagens.services.evento_checklist import build_evento_checklist


class EventoChecklistTests(TestCase):
    def setUp(self) -> None:
        self.estado_pr = Estado.objects.create(sigla="PR", nome="Paraná")
        self.cidade_sede = Cidade.objects.create(nome="Curitiba", estado=self.estado_pr)
        self.cidade_destino = Cidade.objects.create(nome="Maringá", estado=self.estado_pr)
        self.viajante_ascom = Viajante.objects.create(
            nome="Servidor ASCOM",
            rg="12345678X",
            cpf="00000000000",
            cargo="Delegado",
            is_ascom=True,
        )
        self.viajante_nao_ascom = Viajante.objects.create(
            nome="Servidor Não ASCOM",
            rg="87654321X",
            cpf="11111111111",
            cargo="Analista",
            is_ascom=False,
        )

    def _trecho(self, oficio: Oficio, saida_em_dias: int):
        saida_data = timezone.localdate() + timedelta(days=saida_em_dias)
        Trecho.objects.create(
            oficio=oficio,
            ordem=1,
            origem_estado=self.estado_pr,
            origem_cidade=self.cidade_sede,
            destino_estado=self.estado_pr,
            destino_cidade=self.cidade_destino,
            saida_data=saida_data,
            saida_hora=time(8, 0),
            chegada_data=saida_data,
            chegada_hora=time(12, 0),
        )

    def test_evento_com_dois_oficios_e_viajante_nao_ascom_exige_termo(self) -> None:
        evento = Evento.objects.create(titulo="Evento Teste", tem_convite_ou_oficio_evento=True)
        of1 = Oficio.objects.create(
            oficio="1/2026",
            protocolo="001/2026",
            assunto="Of1",
            assunto_tipo=Oficio.AssuntoTipo.AUTORIZACAO,
            placa="ABC1234",
            modelo="Viatura",
            combustivel="Gasolina",
            motorista=self.viajante_ascom.nome,
            motivo="Teste",
            tipo_destino="INTERIOR",
            estado_sede=self.estado_pr,
            cidade_sede=self.cidade_sede,
            estado_destino=self.estado_pr,
            cidade_destino=self.cidade_destino,
            status=Oficio.Status.FINAL,
            evento=evento,
        )
        of1.viajantes.add(self.viajante_nao_ascom)
        self._trecho(of1, 15)
        of2 = Oficio.objects.create(
            oficio="2/2026",
            protocolo="002/2026",
            assunto="Of2",
            assunto_tipo=Oficio.AssuntoTipo.AUTORIZACAO,
            placa="ABC1234",
            modelo="Viatura",
            combustivel="Gasolina",
            motorista=self.viajante_ascom.nome,
            motivo="Teste",
            tipo_destino="INTERIOR",
            estado_sede=self.estado_pr,
            cidade_sede=self.cidade_sede,
            estado_destino=self.estado_pr,
            cidade_destino=self.cidade_destino,
            status=Oficio.Status.FINAL,
            evento=evento,
        )
        of2.viajantes.add(self.viajante_nao_ascom)
        self._trecho(of2, 15)
        checklist = build_evento_checklist(evento)
        self.assertIn("termos", checklist["required_docs"])
        termos = checklist["required_docs"]["termos"]
        # Termos por (ofício, viajante): 2 ofícios × 1 viajante não-ASCOM = 2 itens
        self.assertEqual(len(termos), 2)
        self.assertTrue(all(not t["termo_ok"] for t in termos))
        self.assertFalse(checklist["readiness"]["pronto_para_protocolar"])

    def test_evento_sem_convite_exige_plano_ou_ordem(self) -> None:
        evento = Evento.objects.create(titulo="Evento Sem Convite", tem_convite_ou_oficio_evento=False)
        of1 = Oficio.objects.create(
            oficio="10/2026",
            protocolo="010/2026",
            assunto="Of",
            assunto_tipo=Oficio.AssuntoTipo.AUTORIZACAO,
            placa="X",
            modelo="Y",
            combustivel="Gasolina",
            motorista=self.viajante_ascom.nome,
            motivo="M",
            tipo_destino="INTERIOR",
            estado_sede=self.estado_pr,
            cidade_sede=self.cidade_sede,
            estado_destino=self.estado_pr,
            cidade_destino=self.cidade_destino,
            status=Oficio.Status.FINAL,
            evento=evento,
        )
        of1.viajantes.add(self.viajante_ascom)
        self._trecho(of1, 20)
        checklist = build_evento_checklist(evento)
        self.assertTrue(checklist["required_docs"]["plano_ou_ordem"]["required"])
        self.assertEqual(checklist["required_docs"]["plano_ou_ordem"]["status"], "pendente")
        self.assertFalse(checklist["readiness"]["pronto_para_protocolar"])

    def test_oficio_antecedencia_menor_10_exige_justificativa(self) -> None:
        evento = Evento.objects.create(titulo="Evento Just", tem_convite_ou_oficio_evento=True)
        of1 = Oficio.objects.create(
            oficio="20/2026",
            protocolo="020/2026",
            assunto="Of",
            assunto_tipo=Oficio.AssuntoTipo.AUTORIZACAO,
            placa="X",
            modelo="Y",
            combustivel="Gasolina",
            motorista=self.viajante_ascom.nome,
            motivo="M",
            tipo_destino="INTERIOR",
            estado_sede=self.estado_pr,
            cidade_sede=self.cidade_sede,
            estado_destino=self.estado_pr,
            cidade_destino=self.cidade_destino,
            status=Oficio.Status.FINAL,
            evento=evento,
        )
        of1.viajantes.add(self.viajante_ascom)
        self._trecho(of1, 5)
        checklist = build_evento_checklist(evento)
        self.assertEqual(len(checklist["required_docs"]["oficios"]), 1)
        self.assertTrue(checklist["required_docs"]["oficios"][0]["exige_justificativa"])
        self.assertFalse(checklist["required_docs"]["oficios"][0]["justificativa_ok"])
        self.assertFalse(checklist["readiness"]["pronto_para_protocolar"])

    def test_checklist_pronto_para_protocolar_somente_quando_tudo_ok(self) -> None:
        evento = Evento.objects.create(titulo="Evento Completo", tem_convite_ou_oficio_evento=True)
        of1 = Oficio.objects.create(
            oficio="30/2026",
            protocolo="030/2026",
            assunto="Of",
            assunto_tipo=Oficio.AssuntoTipo.AUTORIZACAO,
            placa="X",
            modelo="Y",
            combustivel="Gasolina",
            motorista=self.viajante_ascom.nome,
            motivo="M",
            tipo_destino="INTERIOR",
            estado_sede=self.estado_pr,
            cidade_sede=self.cidade_sede,
            estado_destino=self.estado_pr,
            cidade_destino=self.cidade_destino,
            status=Oficio.Status.FINAL,
            evento=evento,
            justificativa_texto="Justificativa preenchida.",
        )
        of1.viajantes.add(self.viajante_ascom)
        self._trecho(of1, 15)
        checklist = build_evento_checklist(evento)
        self.assertTrue(checklist["required_docs"]["roteiro"]["status"] == "ok")
        self.assertTrue(checklist["required_docs"]["plano_ou_ordem"]["status"] == "ok")
        self.assertTrue(all(
            not o["exige_justificativa"] or o["justificativa_ok"]
            for o in checklist["required_docs"]["oficios"]
        ))
        self.assertTrue(checklist["readiness"]["pronto_para_protocolar"])

    def test_eventos_lista_e_pacote_retornam_200(self) -> None:
        evento = Evento.objects.create(titulo="Evento UI", tem_convite_ou_oficio_evento=True)
        response_lista = self.client.get(reverse("eventos_lista"))
        self.assertEqual(response_lista.status_code, 200)
        self.assertContains(response_lista, "Evento UI")
        response_pacote = self.client.get(reverse("evento_pacote", args=[evento.id]))
        self.assertEqual(response_pacote.status_code, 200)
        self.assertContains(response_pacote, "Evento UI")

    def test_evento_redirect_from_oficio_cria_evento_se_nao_tem(self) -> None:
        oficio = Oficio.objects.create(
            oficio="40/2026",
            protocolo="040/2026",
            assunto="Ofício sem evento",
            assunto_tipo=Oficio.AssuntoTipo.AUTORIZACAO,
            placa="X",
            modelo="Y",
            combustivel="Gasolina",
            motorista=self.viajante_ascom.nome,
            motivo="M",
            tipo_destino="INTERIOR",
            estado_sede=self.estado_pr,
            cidade_sede=self.cidade_sede,
            estado_destino=self.estado_pr,
            cidade_destino=self.cidade_destino,
            status=Oficio.Status.FINAL,
        )
        self.assertIsNone(oficio.evento_id)
        response = self.client.get(reverse("evento_redirect_from_oficio", args=[oficio.id]))
        self.assertEqual(response.status_code, 302)
        oficio.refresh_from_db()
        self.assertIsNotNone(oficio.evento_id)
        self.assertIn(f"/eventos/{oficio.evento_id}/pacote/", response.url)
