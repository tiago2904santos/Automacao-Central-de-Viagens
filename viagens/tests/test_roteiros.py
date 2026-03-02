from datetime import timedelta
import json
from decimal import Decimal

from django.contrib.auth import get_user_model
from django.test import TestCase
from django.urls import reverse
from django.utils import timezone

from viagens.forms import RoteiroForm
from viagens.models import Cidade, Estado, Oficio, OficioRoteiro, Roteiro, TrechoRoteiro, Viajante


User = get_user_model()


class RoteiroModelTest(TestCase):
    def test_roteiro_creation(self):
        roteiro = Roteiro.objects.create(
            nome="Teste Roteiro",
            uf_origem="PR",
            cidade_origem="Curitiba",
            uf_destino="SP",
            cidade_destino="Sao Paulo",
            distancia_km=Decimal("400.50"),
            tipo_deslocamento=Roteiro.TipoDeslocamentoChoices.INTERIOR,
        )
        self.assertEqual(roteiro.nome, "Teste Roteiro")
        self.assertTrue(roteiro.ativo)
        self.assertEqual(roteiro.get_distancia_total(), Decimal("400.50"))

    def test_trecho_roteiro_creation(self):
        roteiro = Roteiro.objects.create(
            nome="Roteiro com Trechos",
            uf_origem="PR",
            cidade_origem="Curitiba",
            uf_destino="SP",
            cidade_destino="Sao Paulo",
            tipo_deslocamento=Roteiro.TipoDeslocamentoChoices.INTERIOR,
        )
        TrechoRoteiro.objects.create(
            roteiro=roteiro,
            ordem=1,
            uf_origem="PR",
            cidade_origem="Curitiba",
            uf_destino="PR",
            cidade_destino="Maringa",
            distancia_km=Decimal("400"),
        )
        TrechoRoteiro.objects.create(
            roteiro=roteiro,
            ordem=2,
            uf_origem="PR",
            cidade_origem="Maringa",
            uf_destino="SP",
            cidade_destino="Sao Paulo",
            distancia_km=Decimal("550"),
        )
        self.assertEqual(roteiro.trechos.count(), 2)
        self.assertEqual(roteiro.get_distancia_total(), Decimal("950"))

    def test_criar_roteiro_valido_gera_n_mais_um_cards(self):
        roteiro = Roteiro.objects.create(
            nome="Curitiba -> Maringa -> Londrina",
            uf_origem="PR",
            cidade_origem="Curitiba",
            uf_destino="PR",
            cidade_destino="Londrina",
            tipo_deslocamento=Roteiro.TipoDeslocamentoChoices.INTERIOR,
        )
        TrechoRoteiro.objects.create(
            roteiro=roteiro,
            ordem=1,
            cidade_origem="Curitiba",
            cidade_destino="Maringa",
        )
        TrechoRoteiro.objects.create(
            roteiro=roteiro,
            ordem=2,
            cidade_origem="Maringa",
            cidade_destino="Londrina",
        )

        cards = roteiro.get_cards_gerados()
        self.assertEqual(len(cards), 3)

    def test_cards_ordem_correta_ultimo_card_termina_na_sede(self):
        roteiro = Roteiro.objects.create(
            nome="Curitiba -> Maringa",
            uf_origem="PR",
            cidade_origem="Curitiba",
            uf_destino="PR",
            cidade_destino="Maringa",
            tipo_deslocamento=Roteiro.TipoDeslocamentoChoices.INTERIOR,
        )
        TrechoRoteiro.objects.create(
            roteiro=roteiro,
            ordem=1,
            cidade_origem="Curitiba",
            cidade_destino="Maringa",
        )

        cards = roteiro.get_cards_gerados()
        self.assertEqual(cards[-1]["destino_cidade"], "Curitiba")
        self.assertEqual(cards[-1]["label"], "Retorno")

    def test_roteiro_invalido_mesma_origem_destino(self):
        form = RoteiroForm(
            data={
                "nome": "Roteiro Invalido",
                "uf_origem": "PR",
                "cidade_origem": "Curitiba",
                "uf_destino": "PR",
                "cidade_destino": "curitiba",
                "distancia_km": "10",
                "tipo_deslocamento": Roteiro.TipoDeslocamentoChoices.INTERIOR,
                "ativo": True,
            }
        )
        self.assertFalse(form.is_valid())
        self.assertIn(
            "A cidade de origem nao pode ser a mesma que a cidade de destino",
            form.errors["__all__"][0],
        )


class RoteiroViewTest(TestCase):
    def setUp(self):
        self.user = User.objects.create_user(username="testuser", password="password123")
        self.client.login(username="testuser", password="password123")
        self.estado_pr = Estado.objects.create(sigla="PR", nome="Parana")
        self.cidade_curitiba = Cidade.objects.create(nome="Curitiba", estado=self.estado_pr)
        self.cidade_maringa = Cidade.objects.create(nome="Maringa", estado=self.estado_pr)
        self.cidade_londrina = Cidade.objects.create(nome="Londrina", estado=self.estado_pr)
        self.viajante = Viajante.objects.create(
            nome="Servidor Teste",
            rg="12345678X",
            cpf="000.000.000-00",
            cargo="Delegado",
        )
        self.roteiro1 = Roteiro.objects.create(
            nome="Roteiro A",
            uf_origem="PR",
            cidade_origem="Curitiba",
            uf_destino="SC",
            cidade_destino="Florianopolis",
            tipo_deslocamento=Roteiro.TipoDeslocamentoChoices.INTERIOR,
        )
        self.roteiro2 = Roteiro.objects.create(
            nome="Roteiro B",
            uf_origem="SP",
            cidade_origem="Sao Paulo",
            uf_destino="RJ",
            cidade_destino="Rio de Janeiro",
            tipo_deslocamento=Roteiro.TipoDeslocamentoChoices.CAPITAL,
        )
        self.oficio = Oficio.objects.create(
            numero=1,
            ano=2026,
            protocolo="12345/2026",
        )

    def _set_wizard_session(self):
        session = self.client.session
        session["oficio_wizard"] = {
            "oficio": "123/2026",
            "protocolo": "121234567",
            "placa": "ABC1234",
            "modelo": "Uno",
            "combustivel": "Gasolina",
            "viajantes_ids": [self.viajante.id],
        }
        session.save()

    def test_roteiro_lista_view(self):
        response = self.client.get(reverse("roteiro_lista"))
        self.assertEqual(response.status_code, 200)
        self.assertContains(response, "Roteiro A")
        self.assertContains(response, "Roteiro B")

    def test_roteiro_lista_busca(self):
        response = self.client.get(reverse("roteiro_lista"), {"q": "Curitiba"})
        self.assertEqual(response.status_code, 200)
        self.assertContains(response, "Roteiro A")
        self.assertNotContains(response, "Roteiro B")

    def test_roteiro_create_view(self):
        response = self.client.get(reverse("roteiro_create"))
        self.assertEqual(response.status_code, 200)
        self.assertContains(response, "Novo roteiro")

        form_data = {
            "nome": "Novo Roteiro Teste",
            "descricao": "Descricao do novo roteiro",
            "uf_origem": "RS",
            "cidade_origem": "Porto Alegre",
            "uf_destino": "SC",
            "cidade_destino": "Florianopolis",
            "distancia_km": "500",
            "tipo_deslocamento": Roteiro.TipoDeslocamentoChoices.INTERIOR,
            "ativo": "on",
            "trechos_roteiro-TOTAL_FORMS": "1",
            "trechos_roteiro-INITIAL_FORMS": "0",
            "trechos_roteiro-MIN_NUM_FORMS": "0",
            "trechos_roteiro-MAX_NUM_FORMS": "1000",
            "trechos_roteiro-0-ordem": "1",
            "trechos_roteiro-0-uf_origem": "RS",
            "trechos_roteiro-0-cidade_origem": "Porto Alegre",
            "trechos_roteiro-0-uf_destino": "SC",
            "trechos_roteiro-0-cidade_destino": "Florianopolis",
            "trechos_roteiro-0-distancia_km": "500",
        }
        response = self.client.post(reverse("roteiro_create"), form_data, follow=True)
        self.assertEqual(response.status_code, 200)
        self.assertTrue(Roteiro.objects.filter(nome="Novo Roteiro Teste").exists())
        self.assertContains(response, "Novo Roteiro Teste")
        self.assertContains(response, "criado com sucesso")

    def test_api_roteiros_buscar_retorna_json(self):
        response = self.client.get(reverse("api_roteiros_buscar"), {"q": "Roteiro"})
        self.assertEqual(response.status_code, 200)
        data = response.json()
        self.assertIn("roteiros", data)
        self.assertEqual(len(data["roteiros"]), 2)
        self.assertEqual(data["roteiros"][0]["nome"], "Roteiro B")
        self.assertEqual(data["roteiros"][1]["nome"], "Roteiro A")

    def test_api_roteiro_json_retorna_trechos(self):
        TrechoRoteiro.objects.create(
            roteiro=self.roteiro1,
            ordem=1,
            uf_origem="PR",
            cidade_origem="Curitiba",
            uf_destino="SC",
            cidade_destino="Florianopolis",
            distancia_km=Decimal("400"),
        )
        response = self.client.get(reverse("api_roteiro_json", args=[self.roteiro1.pk]))
        self.assertEqual(response.status_code, 200)
        data = response.json()
        self.assertEqual(data["nome"], "Roteiro A")
        self.assertEqual(len(data["trechos"]), 1)
        self.assertEqual(data["trechos"][0]["cidade_origem"], "Curitiba")

    def test_api_roteiros_cards(self):
        TrechoRoteiro.objects.create(
            roteiro=self.roteiro1,
            ordem=1,
            cidade_origem="Curitiba",
            cidade_destino="Florianopolis",
        )
        response = self.client.get(reverse("api_roteiros_cards", args=[self.roteiro1.pk]))
        self.assertEqual(response.status_code, 200)
        data = response.json()
        self.assertEqual(data["id"], self.roteiro1.pk)
        self.assertEqual(data["uf_sede"], "PR")
        self.assertEqual(len(data["cards"]), 2)

    def test_api_salvar_roteiro(self):
        payload = {
            "nome": "Curitiba -> Maringa -> Londrina",
            "uf_sede": "PR",
            "cidade_sede": "Curitiba",
            "destinos": [
                {"uf": "PR", "cidade": "Maringa"},
                {"uf": "PR", "cidade": "Londrina"},
            ],
        }
        response = self.client.post(
            reverse("api_roteiros_salvar"),
            data=json.dumps(payload),
            content_type="application/json",
        )
        self.assertEqual(response.status_code, 201)
        data = response.json()
        self.assertEqual(data["nome"], payload["nome"])
        self.assertEqual(len(data["cards"]), 3)

    def test_api_salvar_sem_destino(self):
        response = self.client.post(
            reverse("api_roteiros_salvar"),
            data=json.dumps(
                {
                    "nome": "Roteiro sem destino",
                    "uf_sede": "PR",
                    "cidade_sede": "Curitiba",
                    "destinos": [],
                }
            ),
            content_type="application/json",
        )
        self.assertEqual(response.status_code, 400)
        self.assertEqual(
            response.json()["error"],
            "O roteiro deve ter pelo menos um destino.",
        )

    def test_api_salvar_sem_nome(self):
        response = self.client.post(
            reverse("api_roteiros_salvar"),
            data=json.dumps(
                {
                    "nome": "",
                    "uf_sede": "PR",
                    "cidade_sede": "Curitiba",
                    "destinos": [{"uf": "PR", "cidade": "Maringa"}],
                }
            ),
            content_type="application/json",
        )
        self.assertEqual(response.status_code, 400)
        self.assertEqual(
            response.json()["error"],
            "O nome do roteiro e obrigatorio.",
        )

    def test_uf_padrao_pr(self):
        roteiro = Roteiro.objects.create(
            nome="Teste Default",
            uf_origem="PR",
            cidade_origem="Curitiba",
            uf_destino="PR",
            cidade_destino="Maringa",
            tipo_deslocamento=Roteiro.TipoDeslocamentoChoices.INTERIOR,
        )
        trecho = TrechoRoteiro.objects.create(
            roteiro=roteiro,
            ordem=1,
            cidade_origem="Curitiba",
            cidade_destino="Maringa",
        )
        self.assertEqual(trecho.uf_origem, "PR")
        self.assertEqual(trecho.uf_destino, "PR")

    def test_oficio_step3_aceita_roteiro_id(self):
        self._set_wizard_session()
        TrechoRoteiro.objects.create(
            roteiro=self.roteiro1,
            ordem=1,
            cidade_origem="Curitiba",
            cidade_destino="Florianopolis",
        )
        saida_data = timezone.localdate() + timedelta(days=15)
        retorno_data = saida_data + timedelta(days=1)
        payload = {
            "roteiro_id": str(self.roteiro1.pk),
            "trechos-TOTAL_FORMS": "1",
            "trechos-INITIAL_FORMS": "0",
            "trechos-MIN_NUM_FORMS": "0",
            "trechos-MAX_NUM_FORMS": "1000",
            "trechos-0-origem_estado": self.estado_pr.sigla,
            "trechos-0-origem_cidade": str(self.cidade_curitiba.id),
            "trechos-0-destino_estado": self.estado_pr.sigla,
            "trechos-0-destino_cidade": str(self.cidade_maringa.id),
            "trechos-0-saida_data": saida_data.isoformat(),
            "trechos-0-saida_hora": "08:00",
            "retorno_saida_data": retorno_data.isoformat(),
            "retorno_saida_hora": "09:00",
            "retorno_chegada_data": retorno_data.isoformat(),
            "retorno_chegada_hora": "18:00",
            "tipo_destino": "INTERIOR",
            "motivo": "Teste com roteiro",
        }

        response = self.client.post(reverse("oficio_step3"), payload)
        self.assertEqual(response.status_code, 302)
        oficio = Oficio.objects.filter(status=Oficio.Status.DRAFT).order_by("-id").first()
        self.assertEqual(oficio.roteiro_id, self.roteiro1.pk)

    def test_oficio_step3_nao_aceita_roteiro_invalido(self):
        self._set_wizard_session()
        saida_data = timezone.localdate() + timedelta(days=15)
        retorno_data = saida_data + timedelta(days=1)
        payload = {
            "roteiro_id": "999999",
            "trechos-TOTAL_FORMS": "1",
            "trechos-INITIAL_FORMS": "0",
            "trechos-MIN_NUM_FORMS": "0",
            "trechos-MAX_NUM_FORMS": "1000",
            "trechos-0-origem_estado": self.estado_pr.sigla,
            "trechos-0-origem_cidade": str(self.cidade_curitiba.id),
            "trechos-0-destino_estado": self.estado_pr.sigla,
            "trechos-0-destino_cidade": str(self.cidade_maringa.id),
            "trechos-0-saida_data": saida_data.isoformat(),
            "trechos-0-saida_hora": "08:00",
            "retorno_saida_data": retorno_data.isoformat(),
            "retorno_saida_hora": "09:00",
            "retorno_chegada_data": retorno_data.isoformat(),
            "retorno_chegada_hora": "18:00",
            "tipo_destino": "INTERIOR",
            "motivo": "Teste com roteiro invalido",
        }

        response = self.client.post(reverse("oficio_step3"), payload, follow=True)
        self.assertEqual(response.status_code, 200)
        oficio = Oficio.objects.filter(status=Oficio.Status.DRAFT).order_by("-id").first()
        self.assertIsNone(oficio.roteiro_id)

    def test_vincular_roteiro_ao_oficio(self):
        url = reverse("oficio_vincular_roteiro", args=[self.oficio.pk])
        payload = json.dumps({"roteiro": self.roteiro1.pk, "observacao": "Teste de vinculo"})
        response = self.client.post(url, payload, content_type="application/json")
        self.assertEqual(response.status_code, 200)
        json_response = response.json()
        self.assertTrue(json_response["success"])
        self.assertTrue(
            OficioRoteiro.objects.filter(oficio=self.oficio, roteiro=self.roteiro1).exists()
        )

    def test_vincular_roteiro_duplicado_ao_oficio(self):
        OficioRoteiro.objects.create(oficio=self.oficio, roteiro=self.roteiro1)
        url = reverse("oficio_vincular_roteiro", args=[self.oficio.pk])
        payload = json.dumps({"roteiro": self.roteiro1.pk, "observacao": "Teste de vinculo"})
        response = self.client.post(url, payload, content_type="application/json")
        self.assertEqual(response.status_code, 400)
        json_response = response.json()
        self.assertFalse(json_response["success"])
        self.assertIn("ja esta vinculado", json_response["message"])

    def test_desvincular_roteiro_do_oficio(self):
        OficioRoteiro.objects.create(oficio=self.oficio, roteiro=self.roteiro1)
        url = reverse("oficio_desvincular_roteiro", args=[self.oficio.pk, self.roteiro1.pk])
        response = self.client.post(url)
        self.assertEqual(response.status_code, 200)
        json_response = response.json()
        self.assertTrue(json_response["success"])
        self.assertFalse(
            OficioRoteiro.objects.filter(oficio=self.oficio, roteiro=self.roteiro1).exists()
        )

    def test_delete_roteiro_com_vinculo_retorna_erro(self):
        OficioRoteiro.objects.create(oficio=self.oficio, roteiro=self.roteiro1)
        response = self.client.post(reverse("roteiro_delete", args=[self.roteiro1.pk]), follow=True)
        self.assertEqual(response.status_code, 200)
        self.assertContains(
            response,
            "Nao e possivel excluir o roteiro pois ele esta vinculado a um ou mais oficios.",
        )
        self.assertTrue(Roteiro.objects.filter(pk=self.roteiro1.pk, ativo=True).exists())

    def test_soft_delete_roteiro_sem_vinculo(self):
        response = self.client.post(reverse("roteiro_delete", args=[self.roteiro2.pk]), follow=True)
        self.assertEqual(response.status_code, 200)
        self.assertContains(response, "Roteiro B")
        self.assertContains(response, "desativado com sucesso")
        self.assertFalse(Roteiro.objects.filter(pk=self.roteiro2.pk, ativo=True).exists())
        self.assertTrue(Roteiro.objects.filter(pk=self.roteiro2.pk, ativo=False).exists())
