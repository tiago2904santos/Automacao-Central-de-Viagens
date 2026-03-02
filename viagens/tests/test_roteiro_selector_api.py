import json

from django.contrib.auth import get_user_model
from django.test import Client, TestCase
from django.urls import reverse

from viagens.models import Cidade, Estado, RoteiroViagem, TrechoRoteiro


User = get_user_model()


class RoteiroSelectorAPITests(TestCase):
    """Testa os endpoints AJAX de roteiro usados na Etapa 3 do Oficio."""

    def setUp(self):
        self.client = Client()
        self.user = User.objects.create_user(
            username="testuser", password="testpassword"
        )
        self.client.login(username="testuser", password="testpassword")

        self.estado_pr = Estado.objects.create(sigla="PR", nome="Parana")
        self.cidade_curitiba = Cidade.objects.create(
            nome="Curitiba", estado=self.estado_pr
        )
        self.cidade_maringa = Cidade.objects.create(
            nome="Maringa", estado=self.estado_pr
        )
        self.cidade_londrina = Cidade.objects.create(
            nome="Londrina", estado=self.estado_pr
        )

        self.roteiro1 = RoteiroViagem.objects.create(
            nome="Curitiba-Maringa-Londrina",
            uf_origem="PR",
            cidade_origem="Curitiba",
            uf_destino="PR",
            cidade_destino="Londrina",
            tipo_deslocamento=RoteiroViagem.TipoDeslocamentoChoices.INTERIOR,
            ativo=True,
        )
        TrechoRoteiro.objects.create(
            roteiro=self.roteiro1,
            ordem=1,
            uf_origem="PR",
            cidade_origem="Curitiba",
            uf_destino="PR",
            cidade_destino="Maringa",
            modal=TrechoRoteiro.ModalChoices.VEICULO_PROPRIO,
        )
        TrechoRoteiro.objects.create(
            roteiro=self.roteiro1,
            ordem=2,
            uf_origem="PR",
            cidade_origem="Maringa",
            uf_destino="PR",
            cidade_destino="Londrina",
            modal=TrechoRoteiro.ModalChoices.VEICULO_PROPRIO,
        )

        self.roteiro2 = RoteiroViagem.objects.create(
            nome="Maringa-Curitiba",
            uf_origem="PR",
            cidade_origem="Maringa",
            uf_destino="PR",
            cidade_destino="Curitiba",
            tipo_deslocamento=RoteiroViagem.TipoDeslocamentoChoices.INTERIOR,
            ativo=True,
        )
        TrechoRoteiro.objects.create(
            roteiro=self.roteiro2,
            ordem=1,
            uf_origem="PR",
            cidade_origem="Maringa",
            uf_destino="PR",
            cidade_destino="Curitiba",
            modal=TrechoRoteiro.ModalChoices.VEICULO_PROPRIO,
        )

    def test_buscar_roteiros_retorna_lista(self):
        response = self.client.get(reverse("api_roteiros_buscar"), {"q": "Curitiba"})
        self.assertEqual(response.status_code, 200)
        data = response.json()
        self.assertIn("roteiros", data)
        self.assertGreater(len(data["roteiros"]), 0)
        nomes = [item["nome"] for item in data["roteiros"]]
        self.assertIn(self.roteiro1.nome, nomes)

    def test_buscar_roteiros_por_cidade_destino(self):
        response = self.client.get(reverse("api_roteiros_buscar"), {"q": "Londrina"})
        self.assertEqual(response.status_code, 200)
        data = response.json()
        self.assertIn("roteiros", data)
        self.assertEqual(len(data["roteiros"]), 1)
        self.assertEqual(data["roteiros"][0]["nome"], self.roteiro1.nome)

    def test_buscar_roteiro_q_curto_retorna_lista_vazia(self):
        response = self.client.get(reverse("api_roteiros_buscar"), {"q": "C"})
        self.assertEqual(response.status_code, 200)
        data = response.json()
        self.assertEqual(len(data["roteiros"]), 0)

    def test_detalhe_roteiro_json(self):
        response = self.client.get(
            reverse("api_roteiro_detalhe_json", kwargs={"pk": self.roteiro1.pk})
        )
        self.assertEqual(response.status_code, 200)
        data = response.json()
        self.assertIn("trechos", data)
        self.assertEqual(data["id"], self.roteiro1.pk)
        self.assertEqual(len(data["trechos"]), 2)
        self.assertEqual(data["trechos"][0]["cidade_destino_nome"], "Maringa")

    def test_detalhe_roteiro_nao_encontrado(self):
        response = self.client.get(
            reverse("api_roteiro_detalhe_json", kwargs={"pk": 99999})
        )
        self.assertEqual(response.status_code, 404)
        self.assertIn("erro", response.json())

    def test_criar_roteiro_inline_sucesso(self):
        payload = {
            "nome": "Novo Roteiro Teste",
            "uf_sede_id": self.estado_pr.pk,
            "cidade_sede_id": self.cidade_curitiba.pk,
            "trechos": [
                {
                    "uf_destino": "PR",
                    "cidade_destino_id": self.cidade_maringa.pk,
                    "modal": "veiculo_proprio",
                },
                {
                    "uf_destino": "PR",
                    "cidade_destino_id": self.cidade_londrina.pk,
                    "modal": "veiculo_proprio",
                },
            ],
        }
        response = self.client.post(
            reverse("api_roteiro_criar_inline"),
            data=json.dumps(payload),
            content_type="application/json",
        )
        self.assertEqual(response.status_code, 201)
        data = response.json()
        self.assertTrue(data.get("sucesso"))
        self.assertIn("roteiro_id", data)
        self.assertEqual(RoteiroViagem.objects.count(), 3)
        self.assertEqual(
            TrechoRoteiro.objects.filter(roteiro_id=data["roteiro_id"]).count(),
            2,
        )

    def test_criar_roteiro_sem_nome_retorna_400(self):
        payload = {
            "nome": "",
            "uf_sede_id": self.estado_pr.pk,
            "cidade_sede_id": self.cidade_curitiba.pk,
            "trechos": [
                {"uf_destino": "PR", "cidade_destino_id": self.cidade_maringa.pk}
            ],
        }
        response = self.client.post(
            reverse("api_roteiro_criar_inline"),
            data=json.dumps(payload),
            content_type="application/json",
        )
        self.assertEqual(response.status_code, 400)
        data = response.json()
        self.assertIn("erro", data)
        self.assertEqual(data["erro"], "Nome do roteiro e obrigatorio.")

    def test_criar_roteiro_sem_trechos_retorna_400(self):
        payload = {
            "nome": "Roteiro Sem Trechos",
            "uf_sede_id": self.estado_pr.pk,
            "cidade_sede_id": self.cidade_curitiba.pk,
            "trechos": [],
        }
        response = self.client.post(
            reverse("api_roteiro_criar_inline"),
            data=json.dumps(payload),
            content_type="application/json",
        )
        self.assertEqual(response.status_code, 400)
        data = response.json()
        self.assertIn("erro", data)
        self.assertEqual(
            data["erro"],
            "Adicione ao menos um destino antes de salvar o roteiro.",
        )

    def test_criar_roteiro_sem_uf_sede_retorna_400(self):
        payload = {
            "nome": "Roteiro Sem UF Sede",
            "uf_sede_id": None,
            "cidade_sede_id": self.cidade_curitiba.pk,
            "trechos": [
                {"uf_destino": "PR", "cidade_destino_id": self.cidade_maringa.pk}
            ],
        }
        response = self.client.post(
            reverse("api_roteiro_criar_inline"),
            data=json.dumps(payload),
            content_type="application/json",
        )
        self.assertEqual(response.status_code, 400)
        data = response.json()
        self.assertIn("erro", data)
        self.assertEqual(data["erro"], "UF e Cidade da sede sao obrigatorios.")

    def test_criar_roteiro_com_cidade_destino_invalida_rollback(self):
        initial_roteiro_count = RoteiroViagem.objects.count()
        initial_trecho_count = TrechoRoteiro.objects.count()
        payload = {
            "nome": "Roteiro com Erro",
            "uf_sede_id": self.estado_pr.pk,
            "cidade_sede_id": self.cidade_curitiba.pk,
            "trechos": [
                {"uf_destino": "PR", "cidade_destino_id": self.cidade_maringa.pk},
                {"uf_destino": "PR", "cidade_destino_id": 999999},
            ],
        }
        response = self.client.post(
            reverse("api_roteiro_criar_inline"),
            data=json.dumps(payload),
            content_type="application/json",
        )
        self.assertEqual(response.status_code, 400)
        self.assertEqual(RoteiroViagem.objects.count(), initial_roteiro_count)
        self.assertEqual(TrechoRoteiro.objects.count(), initial_trecho_count)
