from django.test import TestCase
from django.urls import reverse

from viagens.models import Cidade, Estado, Oficio, Trecho, Viajante


class EditWizardReorderTests(TestCase):
    def setUp(self) -> None:
        self.estado_pr = Estado.objects.create(sigla="PR", nome="Parana")
        self.sede = Cidade.objects.create(nome="Curitiba", estado=self.estado_pr)
        self.d1 = Cidade.objects.create(nome="Cascavel", estado=self.estado_pr)
        self.d2 = Cidade.objects.create(nome="Londrina", estado=self.estado_pr)
        self.d3 = Cidade.objects.create(nome="Maringa", estado=self.estado_pr)
        self.d4 = Cidade.objects.create(nome="Ponta Grossa", estado=self.estado_pr)

        self.viajante = Viajante.objects.create(
            nome="Servidor Teste",
            rg="123456",
            cpf="000.000.000-00",
            cargo="Delegado",
        )

        self.oficio = Oficio.objects.create(
            oficio="123/2026",
            protocolo="456/2026",
            assunto="Teste reorder",
            placa="ABC1234",
            modelo="Uno",
            combustivel="Gasolina",
            tipo_destino="INTERIOR",
            retorno_saida_cidade=f"{self.d4.nome}/{self.estado_pr.sigla}",
            retorno_saida_data="2026-02-01",
            retorno_saida_hora="08:00",
            retorno_chegada_cidade=f"{self.sede.nome}/{self.estado_pr.sigla}",
            retorno_chegada_data="2026-02-02",
            retorno_chegada_hora="18:00",
            quantidade_diarias="",
            valor_diarias="",
            valor_diarias_extenso="",
            motivo="Teste",
            estado_sede=self.estado_pr,
            cidade_sede=self.sede,
        )
        self.oficio.viajantes.add(self.viajante)

        # Initial chain: sede -> d1 -> d2 -> d3 -> d4
        Trecho.objects.create(
            oficio=self.oficio,
            ordem=1,
            origem_estado=self.estado_pr,
            origem_cidade=self.sede,
            destino_estado=self.estado_pr,
            destino_cidade=self.d1,
        )
        Trecho.objects.create(
            oficio=self.oficio,
            ordem=2,
            origem_estado=self.estado_pr,
            origem_cidade=self.d1,
            destino_estado=self.estado_pr,
            destino_cidade=self.d2,
        )
        Trecho.objects.create(
            oficio=self.oficio,
            ordem=3,
            origem_estado=self.estado_pr,
            origem_cidade=self.d2,
            destino_estado=self.estado_pr,
            destino_cidade=self.d3,
        )
        Trecho.objects.create(
            oficio=self.oficio,
            ordem=4,
            origem_estado=self.estado_pr,
            origem_cidade=self.d3,
            destino_estado=self.estado_pr,
            destino_cidade=self.d4,
        )

    def test_edit_wizard_persiste_ordem_dos_trechos(self) -> None:
        """
        Reorder via Step3 must persist to DB as Trecho.ordem = 1..N following the draft order.
        """
        url = reverse("oficio_edit_step3", args=[self.oficio.id])

        # New destinos order: d3, d1, d4, d2  (permutation of original indices: 2,0,3,1)
        payload = {
            "sede_uf": self.estado_pr.sigla,
            "sede_cidade": str(self.sede.id),
            "destinos-TOTAL_FORMS": "4",
            "destinos-INITIAL_FORMS": "0",
            "destinos-0-uf": self.estado_pr.sigla,
            "destinos-0-cidade": str(self.d1.id),
            "destinos-1-uf": self.estado_pr.sigla,
            "destinos-1-cidade": str(self.d2.id),
            "destinos-2-uf": self.estado_pr.sigla,
            "destinos-2-cidade": str(self.d3.id),
            "destinos-3-uf": self.estado_pr.sigla,
            "destinos-3-cidade": str(self.d4.id),
            "destinos-order": "2,0,3,1",
            "trechos-TOTAL_FORMS": "4",
            "trechos-INITIAL_FORMS": "0",
            "trechos-MIN_NUM_FORMS": "0",
            "trechos-MAX_NUM_FORMS": "1000",
            # New chain: sede->d3, d3->d1, d1->d4, d4->d2
            "trechos-0-origem_estado": self.estado_pr.sigla,
            "trechos-0-origem_cidade": str(self.sede.id),
            "trechos-0-destino_estado": self.estado_pr.sigla,
            "trechos-0-destino_cidade": str(self.d3.id),
            "trechos-0-saida_data": "",
            "trechos-0-saida_hora": "",
            "trechos-0-chegada_data": "",
            "trechos-0-chegada_hora": "",
            "trechos-1-origem_estado": self.estado_pr.sigla,
            "trechos-1-origem_cidade": str(self.d3.id),
            "trechos-1-destino_estado": self.estado_pr.sigla,
            "trechos-1-destino_cidade": str(self.d1.id),
            "trechos-1-saida_data": "",
            "trechos-1-saida_hora": "",
            "trechos-1-chegada_data": "",
            "trechos-1-chegada_hora": "",
            "trechos-2-origem_estado": self.estado_pr.sigla,
            "trechos-2-origem_cidade": str(self.d1.id),
            "trechos-2-destino_estado": self.estado_pr.sigla,
            "trechos-2-destino_cidade": str(self.d4.id),
            "trechos-2-saida_data": "",
            "trechos-2-saida_hora": "",
            "trechos-2-chegada_data": "",
            "trechos-2-chegada_hora": "",
            "trechos-3-origem_estado": self.estado_pr.sigla,
            "trechos-3-origem_cidade": str(self.d4.id),
            "trechos-3-destino_estado": self.estado_pr.sigla,
            "trechos-3-destino_cidade": str(self.d2.id),
            "trechos-3-saida_data": "",
            "trechos-3-saida_hora": "",
            "trechos-3-chegada_data": "",
            "trechos-3-chegada_hora": "",
            # Required draft fields in this POST (otherwise Step3 would overwrite with blanks and fail validation).
            "retorno_saida_data": "2026-02-01",
            "retorno_saida_hora": "08:00",
            "retorno_chegada_data": "2026-02-02",
            "retorno_chegada_hora": "18:00",
            "tipo_destino": "INTERIOR",
            "motivo": "Teste reorder",
            "action": "save",
        }

        resp = self.client.post(url, payload)
        self.assertEqual(resp.status_code, 302)

        trechos = list(Trecho.objects.filter(oficio=self.oficio).order_by("ordem"))
        self.assertEqual([t.ordem for t in trechos], [1, 2, 3, 4])
        self.assertEqual(
            [t.destino_cidade_id for t in trechos],
            [self.d3.id, self.d1.id, self.d4.id, self.d2.id],
        )

        # Reopen edit wizard -> session hydration must follow DB ordem (not id/creation order).
        resp = self.client.get(reverse("oficio_edit_step3", args=[self.oficio.id]))
        self.assertEqual(resp.status_code, 200)
        session = self.client.session
        key = f"oficio_edit_wizard:{self.oficio.id}"
        draft = session.get(key) or {}
        self.assertTrue(draft.get("trechos"))
        self.assertEqual(
            [t.get("destino_cidade") for t in draft["trechos"]],
            [str(self.d3.id), str(self.d1.id), str(self.d4.id), str(self.d2.id)],
        )

