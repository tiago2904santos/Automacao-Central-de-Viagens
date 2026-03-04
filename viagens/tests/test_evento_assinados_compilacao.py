# Testes: upload assinados, pronto_para_compilar, compilação PDF
from django.core.files.base import ContentFile
from django.test import TestCase

from viagens.models import (
    DocumentoEventoArquivo,
    Evento,
    EventoProtocoloArquivo,
    Oficio,
    Viajante,
)
from viagens.services.evento_assinados import (
    get_status_assinados_evento,
    is_evento_pronto_para_compilar,
    listar_pendencias_compilacao,
)
from viagens.services.evento_compilacao import compilar_pdf_protocolo


class EventoAssinadosCompilacaoTests(TestCase):
    def setUp(self):
        self.evento = Evento.objects.create(
            titulo="Evento Teste",
            tem_convite_ou_oficio_evento=True,
        )
        self.oficio = Oficio.objects.create(
            assunto="Ofício 1",
            evento=self.evento,
        )
        self.viajante_ascom = Viajante.objects.create(
            nome="Servidor ASCOM",
            rg="123456789",
            cpf="12345678901",
            cargo="Analista",
            is_ascom=True,
        )
        self.viajante_nao_ascom = Viajante.objects.create(
            nome="Servidor Não ASCOM",
            rg="987654321",
            cpf="98765432109",
            cargo="Técnico",
            is_ascom=False,
        )
        self.oficio.viajantes.add(self.viajante_ascom, self.viajante_nao_ascom)

    def test_pronto_para_compilar_false_sem_uploads(self):
        self.assertFalse(is_evento_pronto_para_compilar(self.evento))
        pendencias = listar_pendencias_compilacao(self.evento)
        self.assertIn("Ofício", pendencias[0])
        self.assertTrue(any("falta upload" in p for p in pendencias))

    def test_status_assinados_retorna_estrutura(self):
        status = get_status_assinados_evento(self.evento)
        self.assertIn("oficios", status)
        self.assertIn("plano_ou_ordem", status)
        self.assertIn("justificativas", status)
        self.assertIn("termos", status)
        self.assertEqual(len(status["oficios"]), 1)
        self.assertFalse(status["oficios"][0]["assinado"])

    def test_upload_pdf_cria_arquivo_e_marca_assinado(self):
        pdf_minimo = b"%PDF-1.4 1 0 obj<</Type/Catalog/Pages 2 0 R>>endobj trailer<</Root 1 0 R>>"
        arq = DocumentoEventoArquivo.objects.create(
            evento=self.evento,
            tipo=DocumentoEventoArquivo.Tipo.OFICIO_ASSINADO,
            oficio=self.oficio,
            original_name="oficio.pdf",
            mime_type="application/pdf",
            is_active=True,
        )
        arq.arquivo.save("oficio.pdf", ContentFile(pdf_minimo), save=True)
        status = get_status_assinados_evento(self.evento)
        self.assertTrue(status["oficios"][0]["assinado"])

    def test_compilar_retorna_none_se_nao_pronto(self):
        self.assertIsNone(compilar_pdf_protocolo(self.evento))
