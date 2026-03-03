from datetime import date, timedelta
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

    def test_roteiro_lista_redireciona_para_login_admin_quando_deslogado(self):
        self.client.logout()
        response = self.client.get(reverse("roteiro_lista"))
        self.assertEqual(response.status_code, 302)
        self.assertIn("/admin/login/", response["Location"])

    def test_roteiro_lista_busca(self):
        response = self.client.get(reverse("roteiro_lista"), {"q": "Curitiba"})
        self.assertEqual(response.status_code, 200)
        self.assertContains(response, "Roteiro A")
        self.assertNotContains(response, "Roteiro B")

    def test_roteiro_create_view(self):
        response = self.client.get(reverse("roteiro_novo"))
        self.assertEqual(response.status_code, 200)
        self.assertContains(response, "Novo Roteiro de Viagem")
        self.assertContains(response, 'id="sedeUf"')
        self.assertContains(response, 'id="sedeCidade"')
        self.assertContains(response, 'id="destinosList"')
        self.assertContains(response, 'id="tempoViagem"')
        self.assertContains(response, 'name="tempo_viagem"')
        self.assertContains(response, 'name="retorno_saida_hora"')
        self.assertContains(response, 'name="retorno_chegada_hora"')
        self.assertContains(response, 'placeholder="00:00"')
        self.assertContains(response, 'name="trechos-TOTAL_FORMS"')
        self.assertContains(response, 'id="salvarRoteiroBtn"')

        form_data = {
            "sede_uf": "PR",
            "sede_cidade": str(self.cidade_curitiba.pk),
            "destinos-TOTAL_FORMS": "1",
            "destinos-INITIAL_FORMS": "0",
            "destinos-order": "0",
            "destinos-0-uf": "PR",
            "destinos-0-cidade": str(self.cidade_maringa.pk),
            "trechos-TOTAL_FORMS": "1",
            "trechos-INITIAL_FORMS": "0",
            "trechos-MIN_NUM_FORMS": "0",
            "trechos-MAX_NUM_FORMS": "1000",
            "trechos-0-id": "",
            "trechos-0-ordem": "1",
            "trechos-0-origem_estado": str(self.estado_pr.pk),
            "trechos-0-origem_cidade": str(self.cidade_curitiba.pk),
            "trechos-0-destino_estado": str(self.estado_pr.pk),
            "trechos-0-destino_cidade": str(self.cidade_maringa.pk),
            "trechos-0-saida_data": "2026-03-15",
            "trechos-0-saida_hora": "08:00",
            "trechos-0-chegada_data": "2026-03-15",
            "trechos-0-chegada_hora": "11:00",
            "retorno_saida_cidade": "Maringa",
            "retorno_saida_data": "2026-03-16",
            "retorno_saida_hora": "09:00",
            "retorno_chegada_cidade": "Curitiba",
            "retorno_chegada_data": "2026-03-16",
            "retorno_chegada_hora": "12:00",
        }
        response = self.client.post(reverse("roteiro_novo"), form_data, follow=True)
        self.assertEqual(response.status_code, 200)
        self.assertTrue(Roteiro.objects.filter(nome="Curitiba > Maringa 15/03/2026 08:00").exists())
        self.assertContains(response, "Curitiba &gt; Maringa 15/03/2026 08:00")
        self.assertContains(response, "salvo com sucesso")

    def test_roteiro_editar_view_e_post(self):
        trecho = TrechoRoteiro.objects.create(
            roteiro=self.roteiro1,
            ordem=1,
            origem_estado=self.estado_pr,
            origem_cidade=self.cidade_curitiba,
            destino_estado=self.estado_pr,
            destino_cidade=self.cidade_maringa,
            saida_data=timezone.localdate(),
        )

        response = self.client.get(reverse("roteiro_editar", args=[self.roteiro1.pk]))
        self.assertEqual(response.status_code, 200)
        self.assertContains(response, "Editar Roteiro")
        self.assertContains(response, 'id="sedeUf"')
        self.assertNotContains(response, "Tipo de Deslocamento")

        form_data = {
            "sede_uf": "PR",
            "sede_cidade": str(self.cidade_curitiba.pk),
            "destinos-TOTAL_FORMS": "1",
            "destinos-INITIAL_FORMS": "0",
            "destinos-order": "0",
            "destinos-0-uf": "PR",
            "destinos-0-cidade": str(self.cidade_londrina.pk),
            "trechos-TOTAL_FORMS": "1",
            "trechos-INITIAL_FORMS": "1",
            "trechos-MIN_NUM_FORMS": "0",
            "trechos-MAX_NUM_FORMS": "1000",
            "trechos-0-id": str(trecho.pk),
            "trechos-0-ordem": "1",
            "trechos-0-origem_estado": str(self.estado_pr.pk),
            "trechos-0-origem_cidade": str(self.cidade_curitiba.pk),
            "trechos-0-destino_estado": str(self.estado_pr.pk),
            "trechos-0-destino_cidade": str(self.cidade_londrina.pk),
            "trechos-0-saida_data": "2026-04-01",
            "trechos-0-saida_hora": "07:30",
            "trechos-0-chegada_data": "2026-04-01",
            "trechos-0-chegada_hora": "12:00",
            "retorno_saida_cidade": "Londrina",
            "retorno_saida_data": "2026-04-02",
            "retorno_saida_hora": "14:00",
            "retorno_chegada_cidade": "Curitiba",
            "retorno_chegada_data": "2026-04-02",
            "retorno_chegada_hora": "18:00",
        }
        response = self.client.post(
            reverse("roteiro_editar", args=[self.roteiro1.pk]),
            form_data,
            follow=True,
        )
        self.assertEqual(response.status_code, 200)
        self.roteiro1.refresh_from_db()
        self.assertEqual(self.roteiro1.nome, "Curitiba > Londrina 01/04/2026 07:30")
        self.assertEqual(self.roteiro1.trechos.count(), 1)
        self.assertEqual(self.roteiro1.trechos.first().destino_cidade.nome, "Londrina")

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
        response = self.client.get(reverse("api_roteiro_cards", args=[self.roteiro1.pk]))
        self.assertEqual(response.status_code, 200)
        data = response.json()
        self.assertEqual(data["id"], self.roteiro1.pk)
        self.assertEqual(data["uf_sede"], "PR")
        self.assertEqual(len(data["cards"]), 2)

    def test_api_salvar_roteiro(self):
        payload = {
            "nome": "",
            "sede_uf": "PR",
            "sede_cidade": str(self.cidade_curitiba.pk),
            "destinos": [
                {"uf": "PR", "cidade": str(self.cidade_maringa.pk)},
                {"uf": "PR", "cidade": str(self.cidade_londrina.pk)},
            ],
            "trechos": [
                {
                    "saida_data": "2026-03-15",
                    "saida_hora": "08:00",
                    "chegada_data": "2026-03-15",
                    "chegada_hora": "11:30",
                },
                {
                    "saida_data": "2026-03-16",
                    "saida_hora": "09:00",
                    "chegada_data": "2026-03-16",
                    "chegada_hora": "12:15",
                },
            ],
            "retorno": {
                "saida_data": "2026-03-17",
                "saida_hora": "14:00",
                "chegada_data": "2026-03-17",
                "chegada_hora": "18:45",
            },
        }
        response = self.client.post(
            reverse("api_roteiro_salvar"),
            data=json.dumps(payload),
            content_type="application/json",
        )
        self.assertEqual(response.status_code, 201)
        data = response.json()
        self.assertEqual(data["nome"], "Curitiba > Maringa 15/03/2026 08:00")
        self.assertTrue(data["ok"])
        roteiro = Roteiro.objects.get(pk=data["id"])
        self.assertEqual(roteiro.retorno_saida_data.isoformat(), "2026-03-17")
        self.assertEqual(
            roteiro.trechos.order_by("ordem").first().saida_hora.strftime("%H:%M"),
            "08:00",
        )
        self.assertEqual(len(data["cards"]), 3)

    def test_salvar_roteiro_endpoint_retorna_json(self):
        payload = {
            "nome": "",
            "sede_uf": "PR",
            "sede_cidade": str(self.cidade_curitiba.pk),
            "destinos": [{"uf": "PR", "cidade": str(self.cidade_maringa.pk)}],
            "trechos": [{"saida_data": "2026-03-20", "saida_hora": "09:30"}],
        }
        response = self.client.post(
            reverse("api_roteiro_salvar"),
            data=json.dumps(payload),
            content_type="application/json",
        )
        self.assertEqual(response.status_code, 201)
        self.assertEqual(response["Content-Type"], "application/json")
        data = response.json()
        self.assertTrue(data["ok"])
        self.assertTrue(data["sucesso"])
        self.assertIn("message", data)
        self.assertEqual(data["nome"], "Curitiba > Maringa 20/03/2026 09:30")

    def test_api_cidades_por_estado_alias_retorna_json(self):
        response = self.client.get(
            reverse("api_cidades_por_estado"),
            {"uf": "PR"},
        )
        self.assertEqual(response.status_code, 200)
        data = response.json()
        self.assertIn("cidades", data)
        self.assertTrue(any(item["nome"] == "Curitiba" for item in data["cidades"]))

    def test_api_salvar_sem_destino(self):
        response = self.client.post(
            reverse("api_roteiro_salvar"),
            data=json.dumps(
                {
                    "nome": "",
                    "sede_uf": "PR",
                    "sede_cidade": str(self.cidade_curitiba.pk),
                    "destinos": [],
                }
            ),
            content_type="application/json",
        )
        self.assertEqual(response.status_code, 400)
        self.assertEqual(
            response.json()["error"],
            "Informe ao menos um destino.",
        )

    def test_api_salvar_sem_nome_gera_nome_automatico(self):
        response = self.client.post(
            reverse("api_roteiro_salvar"),
            data=json.dumps(
                {
                    "nome": "",
                    "sede_uf": "PR",
                    "sede_cidade": str(self.cidade_curitiba.pk),
                    "destinos": [{"uf": "PR", "cidade": str(self.cidade_maringa.pk)}],
                    "trechos": [{"saida_data": "2026-03-21", "saida_hora": "10:15"}],
                }
            ),
            content_type="application/json",
        )
        self.assertEqual(response.status_code, 201)
        self.assertEqual(
            response.json()["nome"],
            "Curitiba > Maringa 21/03/2026 10:15",
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
        saida_data = timezone.localdate() + timedelta(days=15)
        retorno_data = saida_data + timedelta(days=1)
        self.roteiro1.estado_sede = self.estado_pr
        self.roteiro1.cidade_sede = self.cidade_curitiba
        self.roteiro1.retorno_saida_data = retorno_data
        self.roteiro1.retorno_saida_hora = "09:00"
        self.roteiro1.retorno_chegada_data = retorno_data
        self.roteiro1.retorno_chegada_hora = "18:00"
        self.roteiro1.save()
        TrechoRoteiro.objects.create(
            roteiro=self.roteiro1,
            ordem=1,
            origem_estado=self.estado_pr,
            origem_cidade=self.cidade_curitiba,
            destino_estado=self.estado_pr,
            destino_cidade=self.cidade_maringa,
            saida_data=saida_data,
            saida_hora="08:00",
        )
        payload = {
            "roteiro_id": str(self.roteiro1.pk),
            "roteiro_origem_id": str(self.roteiro1.pk),
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

    def test_oficio_step3_clona_roteiro_ao_modificar_trechos_de_um_roteiro_existente(self):
        self._set_wizard_session()
        data_original = date(2026, 3, 15)
        data_modificada = date(2026, 3, 20)
        retorno_data = data_modificada + timedelta(days=1)

        self.roteiro1.estado_sede = self.estado_pr
        self.roteiro1.cidade_sede = self.cidade_curitiba
        self.roteiro1.retorno_saida_data = data_original + timedelta(days=1)
        self.roteiro1.retorno_saida_hora = "09:00"
        self.roteiro1.retorno_chegada_data = data_original + timedelta(days=1)
        self.roteiro1.retorno_chegada_hora = "18:00"
        self.roteiro1.save()

        TrechoRoteiro.objects.create(
            roteiro=self.roteiro1,
            ordem=1,
            origem_estado=self.estado_pr,
            origem_cidade=self.cidade_curitiba,
            destino_estado=self.estado_pr,
            destino_cidade=self.cidade_maringa,
            saida_data=data_original,
            saida_hora="08:00",
            chegada_data=data_original,
            chegada_hora="12:00",
        )

        payload = {
            "roteiro_id": str(self.roteiro1.pk),
            "roteiro_origem_id": str(self.roteiro1.pk),
            "trechos-TOTAL_FORMS": "1",
            "trechos-INITIAL_FORMS": "0",
            "trechos-MIN_NUM_FORMS": "0",
            "trechos-MAX_NUM_FORMS": "1000",
            "trechos-0-origem_estado": self.estado_pr.sigla,
            "trechos-0-origem_cidade": str(self.cidade_curitiba.id),
            "trechos-0-destino_estado": self.estado_pr.sigla,
            "trechos-0-destino_cidade": str(self.cidade_maringa.id),
            "trechos-0-saida_data": data_modificada.isoformat(),
            "trechos-0-saida_hora": "08:00",
            "trechos-0-chegada_data": data_modificada.isoformat(),
            "trechos-0-chegada_hora": "12:00",
            "retorno_saida_data": retorno_data.isoformat(),
            "retorno_saida_hora": "09:00",
            "retorno_chegada_data": retorno_data.isoformat(),
            "retorno_chegada_hora": "18:00",
            "tipo_destino": "INTERIOR",
            "motivo": "Teste clone automatico",
        }

        response = self.client.post(reverse("oficio_step3"), payload)

        self.assertEqual(response.status_code, 302)
        oficio = Oficio.objects.filter(status=Oficio.Status.DRAFT).order_by("-id").first()
        self.assertIsNotNone(oficio.roteiro)
        self.assertNotEqual(oficio.roteiro_id, self.roteiro1.pk)
        self.assertTrue(oficio.roteiro.criado_automaticamente)
        self.assertEqual(oficio.roteiro.nome, "Curitiba > Maringa 20/03/2026 08:00")

        novo_trecho = oficio.roteiro.trechos.get()
        self.assertEqual(novo_trecho.saida_data, data_modificada)
        self.assertEqual(novo_trecho.saida_hora.strftime("%H:%M"), "08:00")

        self.roteiro1.refresh_from_db()
        self.assertFalse(self.roteiro1.criado_automaticamente)
        trecho_original = self.roteiro1.trechos.get()
        self.assertEqual(trecho_original.saida_data, data_original)
        self.assertEqual(trecho_original.saida_hora.strftime("%H:%M"), "08:00")

        self.assertEqual(
            self.client.session["oficio_wizard"]["roteiro_origem_id"],
            str(oficio.roteiro_id),
        )

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

    def test_delete_roteiro_com_vinculo_realiza_soft_delete(self):
        OficioRoteiro.objects.create(oficio=self.oficio, roteiro=self.roteiro1)
        response = self.client.post(reverse("roteiro_excluir", args=[self.roteiro1.pk]), follow=True)
        self.assertEqual(response.status_code, 200)
        self.assertContains(response, "Roteiro removido.")
        self.assertFalse(Roteiro.objects.filter(pk=self.roteiro1.pk, ativo=True).exists())

    def test_soft_delete_roteiro_sem_vinculo(self):
        response = self.client.post(reverse("roteiro_excluir", args=[self.roteiro2.pk]), follow=True)
        self.assertEqual(response.status_code, 200)
        self.assertContains(response, "Roteiro removido.")
        self.assertFalse(Roteiro.objects.filter(pk=self.roteiro2.pk, ativo=True).exists())
        self.assertTrue(Roteiro.objects.filter(pk=self.roteiro2.pk, ativo=False).exists())
