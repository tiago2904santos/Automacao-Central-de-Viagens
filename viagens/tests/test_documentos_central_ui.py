from __future__ import annotations

from datetime import time, timedelta

from django.conf import settings
from django.test import TestCase
from django.urls import reverse
from django.utils import timezone

from viagens.models import Cidade, Estado, Oficio, Trecho, Viajante


class DocumentosCentralUiTests(TestCase):
    def setUp(self) -> None:
        self.estado_pr = Estado.objects.create(sigla="PR", nome="Parana")
        self.cidade_sede = Cidade.objects.create(nome="Curitiba", estado=self.estado_pr)
        self.cidade_destino = Cidade.objects.create(nome="Maringa", estado=self.estado_pr)
        self.viajante = Viajante.objects.create(
            nome="Servidor Teste",
            rg="12345678X",
            cpf="00000000000",
            cargo="Delegado",
        )

    def _build_oficio(self, *, saida_em_dias: int = 5) -> Oficio:
        oficio = Oficio.objects.create(
            oficio="123/2026",
            protocolo="456/2026",
            assunto="Teste central documentos",
            assunto_tipo=Oficio.AssuntoTipo.AUTORIZACAO,
            placa="ABC1234",
            modelo="Viatura Teste",
            combustivel="Gasolina",
            motorista=self.viajante.nome,
            motivo="Teste",
            tipo_destino="INTERIOR",
            estado_sede=self.estado_pr,
            cidade_sede=self.cidade_sede,
            estado_destino=self.estado_pr,
            cidade_destino=self.cidade_destino,
            status=Oficio.Status.FINAL,
        )
        oficio.viajantes.add(self.viajante)
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
        return oficio

    def test_debug_settings_permite_testserver_em_allowed_hosts(self) -> None:
        self.assertIn("testserver", settings.ALLOWED_HOSTS)
        response = self.client.get("/planos-trabalho/")
        self.assertEqual(response.status_code, 200)

    def test_oficio_documentos_retorna_200_e_renderiza_abas(self) -> None:
        oficio = self._build_oficio(saida_em_dias=5)

        response = self.client.get(reverse("oficio_documentos", args=[oficio.id]))

        self.assertEqual(response.status_code, 200)
        self.assertContains(response, "Central de Documentos")
        self.assertContains(response, "?tab=oficio")
        self.assertContains(response, "?tab=termo")
        self.assertContains(response, "?tab=plano")
        self.assertContains(response, "?tab=ordem")
        self.assertContains(response, "?tab=justificativa")

    def test_menu_base_contem_links_globais_de_documentos(self) -> None:
        response = self.client.get(reverse("oficios_lista"))

        self.assertEqual(response.status_code, 200)
        self.assertContains(response, reverse("planos_trabalho_list"))
        self.assertContains(response, reverse("justificativas_list"))
        self.assertContains(response, reverse("ordens_servico_list"))

    def test_listagem_oficios_exibe_link_documentos(self) -> None:
        oficio = self._build_oficio(saida_em_dias=12)

        response = self.client.get(reverse("oficios_lista"))

        self.assertEqual(response.status_code, 200)
        self.assertContains(response, "Documentos")
        self.assertContains(response, reverse("oficio_documentos", args=[oficio.id]))
