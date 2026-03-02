from datetime import timedelta

from django.test import TestCase
from django.utils import timezone

from viagens.models import Cidade, Estado, Oficio, PlanoTrabalho, Trecho


class AcaoInstitucionalModelTests(TestCase):
    def setUp(self) -> None:
        self.estado = Estado.objects.create(sigla="PR", nome="Parana")
        self.cidade = Cidade.objects.create(nome="Curitiba", estado=self.estado)

    def test_oficio_novo_cria_acao_automaticamente(self) -> None:
        oficio = Oficio.objects.create(
            numero=901,
            ano=2026,
            protocolo="123456789",
        )

        self.assertIsNotNone(oficio.acao)
        self.assertIn("901/2026", oficio.acao.titulo)

    def test_plano_trabalho_novo_vincula_acao_do_oficio(self) -> None:
        oficio = Oficio.objects.create(
            numero=902,
            ano=2026,
            protocolo="223456789",
        )

        plano = PlanoTrabalho.objects.create(
            oficio=oficio,
            numero=1,
            ano=2026,
            local="Curitiba/PR",
            data_inicio=timezone.localdate(),
            data_fim=timezone.localdate(),
        )

        self.assertEqual(plano.acao, oficio.acao)
        self.assertEqual(oficio.acao.plano_trabalho, plano)

    def test_precisa_justificativa_e_obrigatoria_respeitam_antecedencia(self) -> None:
        oficio = Oficio.objects.create(
            numero=903,
            ano=2026,
            protocolo="323456789",
        )
        Oficio.objects.filter(pk=oficio.pk).update(created_at=timezone.now())
        oficio.refresh_from_db()

        Trecho.objects.create(
            oficio=oficio,
            ordem=1,
            origem_estado=self.estado,
            origem_cidade=self.cidade,
            destino_estado=self.estado,
            destino_cidade=self.cidade,
            saida_data=timezone.localdate() + timedelta(days=5),
        )

        self.assertTrue(oficio.precisa_justificativa)
        self.assertTrue(oficio.justificativa_obrigatoria)

        oficio.justificativa_texto = "Antecedencia reduzida por demanda urgente."
        oficio.save(update_fields=["justificativa_texto"])

        self.assertTrue(oficio.precisa_justificativa)
        self.assertFalse(oficio.justificativa_obrigatoria)
