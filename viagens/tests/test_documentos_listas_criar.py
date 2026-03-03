from __future__ import annotations

from django.test import TestCase
from django.urls import reverse

from viagens.models import Oficio, OrdemServico, PlanoTrabalho


class DocumentosListasCriarTests(TestCase):
    def test_listas_documentos_retornam_200(self) -> None:
        self.assertEqual(self.client.get(reverse("planos_trabalho_list")).status_code, 200)
        self.assertEqual(self.client.get(reverse("ordens_servico_list")).status_code, 200)

    def test_plano_trabalho_criar_get_e_post_criam_base_e_redirecionam(self) -> None:
        base_oficios = Oficio.objects.count()
        base_planos = PlanoTrabalho.objects.count()

        response_get = self.client.get(reverse("plano_trabalho_criar"))
        self.assertEqual(response_get.status_code, 302)
        oficio_get = Oficio.objects.order_by("-id").first()
        self.assertIsNotNone(oficio_get)
        self.assertEqual(Oficio.objects.count(), base_oficios + 1)
        self.assertEqual(PlanoTrabalho.objects.count(), base_planos + 1)
        self.assertIn(reverse("plano_trabalho_step1", args=[oficio_get.id]), response_get.url)

        response_post = self.client.post(reverse("plano_trabalho_criar"))
        self.assertEqual(response_post.status_code, 302)
        oficio_post = Oficio.objects.order_by("-id").first()
        self.assertIsNotNone(oficio_post)
        self.assertEqual(Oficio.objects.count(), base_oficios + 2)
        self.assertEqual(PlanoTrabalho.objects.count(), base_planos + 2)
        self.assertIn(reverse("plano_trabalho_step1", args=[oficio_post.id]), response_post.url)

    def test_ordem_servico_criar_get_e_post_criam_base_e_redirecionam(self) -> None:
        base_oficios = Oficio.objects.count()
        base_ordens = OrdemServico.objects.count()

        response_get = self.client.get(reverse("ordem_servico_criar"))
        self.assertEqual(response_get.status_code, 302)
        oficio_get = Oficio.objects.order_by("-id").first()
        self.assertIsNotNone(oficio_get)
        self.assertEqual(Oficio.objects.count(), base_oficios + 1)
        self.assertEqual(OrdemServico.objects.count(), base_ordens + 1)
        self.assertIn(reverse("ordem_servico_editar", args=[oficio_get.id]), response_get.url)

        response_post = self.client.post(reverse("ordem_servico_criar"))
        self.assertEqual(response_post.status_code, 302)
        oficio_post = Oficio.objects.order_by("-id").first()
        self.assertIsNotNone(oficio_post)
        self.assertEqual(Oficio.objects.count(), base_oficios + 2)
        self.assertEqual(OrdemServico.objects.count(), base_ordens + 2)
        self.assertIn(reverse("ordem_servico_editar", args=[oficio_post.id]), response_post.url)

    def test_links_de_editar_apontam_para_rotas_corretas(self) -> None:
        response_plano_create = self.client.get(reverse("plano_trabalho_criar"))
        self.assertEqual(response_plano_create.status_code, 302)
        oficio_plano = Oficio.objects.order_by("-id").first()
        self.assertIsNotNone(oficio_plano)
        response_planos = self.client.get(reverse("planos_trabalho_list"))
        self.assertEqual(response_planos.status_code, 200)
        self.assertContains(response_planos, reverse("plano_trabalho_editar", args=[oficio_plano.id]))

        response_ordem_create = self.client.get(reverse("ordem_servico_criar"))
        self.assertEqual(response_ordem_create.status_code, 302)
        oficio_ordem = Oficio.objects.order_by("-id").first()
        self.assertIsNotNone(oficio_ordem)
        response_ordens = self.client.get(reverse("ordens_servico_list"))
        self.assertEqual(response_ordens.status_code, 200)
        self.assertContains(response_ordens, reverse("ordem_servico_editar", args=[oficio_ordem.id]))

