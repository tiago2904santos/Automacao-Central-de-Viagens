from __future__ import annotations

from datetime import date, time, timedelta
from io import BytesIO

from django.test import TestCase
from django.urls import reverse
from django.utils import timezone
from docx import Document

from viagens.models import Cidade, Estado, Oficio, OrdemServico, PlanoTrabalho, Trecho, Viajante


class DocumentosListagensTests(TestCase):
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

    def _build_oficio(self, numero: str = "123/2026", *, saida_em_dias: int = 12) -> Oficio:
        oficio = Oficio.objects.create(
            oficio=numero,
            protocolo="456/2026",
            assunto_tipo=Oficio.AssuntoTipo.AUTORIZACAO,
            placa="ABC1234",
            modelo="Viatura Teste",
            combustivel="Gasolina",
            motorista=self.viajante.nome,
            motivo="Teste",
            tipo_destino="INTERIOR",
            quantidade_diarias="1",
            valor_diarias="290,55",
            retorno_saida_data=date(2026, 2, 1),
            retorno_saida_hora=time(8, 0),
            retorno_chegada_data=date(2026, 2, 2),
            retorno_chegada_hora=time(18, 0),
            estado_sede=self.estado_pr,
            cidade_sede=self.cidade_sede,
            estado_destino=self.estado_pr,
            cidade_destino=self.cidade_destino,
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

    def _doc_paragraphs_text(self, payload: bytes) -> str:
        doc = Document(BytesIO(payload))
        return "\n".join(paragraph.text for paragraph in doc.paragraphs)

    def test_listagem_planos_trabalho_retorna_200_e_lista_itens(self) -> None:
        oficio = self._build_oficio("100/2026")
        PlanoTrabalho.objects.create(
            oficio=oficio,
            numero=1,
            ano=2026,
            local="Curitiba/PR",
            data_inicio=date(2026, 2, 1),
            data_fim=date(2026, 2, 2),
            efetivo_por_dia=1,
            valor_total="R$ 290,55",
            coordenador_nome="Coordenador",
            coordenador_cargo="Cargo",
        )

        response = self.client.get(reverse("planos_trabalho_list"))
        self.assertEqual(response.status_code, 200)
        self.assertContains(response, "1/2026")

    def test_listagem_justificativas_pendentes_filtra_corretamente(self) -> None:
        oficio_pendente = self._build_oficio("101/2026", saida_em_dias=5)
        oficio_completo = self._build_oficio("102/2026", saida_em_dias=5)
        oficio_completo.justificativa_texto = "Justificativa preenchida"
        oficio_completo.save(update_fields=["justificativa_texto"])

        response = self.client.get(reverse("justificativas_list"), {"status": "pendentes"})
        self.assertEqual(response.status_code, 200)
        self.assertContains(response, oficio_pendente.numero_formatado)
        self.assertNotContains(response, oficio_completo.numero_formatado)

    def test_listagem_ordens_servico_so_traz_oficios_sem_plano(self) -> None:
        oficio_sem_plano = self._build_oficio("103/2026")
        oficio_com_plano = self._build_oficio("104/2026")
        PlanoTrabalho.objects.create(
            oficio=oficio_com_plano,
            numero=2,
            ano=2026,
            local="Curitiba/PR",
            data_inicio=date(2026, 2, 1),
            data_fim=date(2026, 2, 2),
            efetivo_por_dia=1,
            valor_total="R$ 290,55",
            coordenador_nome="Coordenador",
            coordenador_cargo="Cargo",
        )

        response = self.client.get(reverse("ordens_servico_list"))
        self.assertEqual(response.status_code, 200)
        self.assertContains(response, oficio_sem_plano.numero_formatado)
        self.assertNotContains(response, oficio_com_plano.numero_formatado)

    def test_download_docx_de_plano_e_ordem_retorna_200_e_titulo(self) -> None:
        oficio_plano = self._build_oficio("105/2026")
        PlanoTrabalho.objects.create(
            oficio=oficio_plano,
            numero=3,
            ano=2026,
            local="Curitiba/PR",
            data_inicio=date(2026, 2, 1),
            data_fim=date(2026, 2, 2),
            efetivo_por_dia=1,
            valor_total="R$ 290,55",
            coordenador_nome="Coordenador",
            coordenador_cargo="Cargo",
        )

        oficio_ordem = self._build_oficio("106/2026")
        OrdemServico.objects.create(
            oficio=oficio_ordem,
            numero=5,
            ano=2026,
            referencia="Diligências",
            determinante_nome="Maria Assinante",
            determinante_cargo="Delegada Adjunta",
            finalidade="para visita técnica.",
        )

        plano_response = self.client.get(
            reverse("plano_trabalho_download_docx", args=[oficio_plano.id])
        )
        self.assertEqual(plano_response.status_code, 200)
        self.assertIn("PLANO DE TRABALHO", self._doc_paragraphs_text(plano_response.content))

        ordem_response = self.client.get(
            reverse("ordem_servico_download_docx", args=[oficio_ordem.id])
        )
        self.assertEqual(ordem_response.status_code, 200)
        self.assertIn("ORDEM DE SERVIÇO Nº", self._doc_paragraphs_text(ordem_response.content))
