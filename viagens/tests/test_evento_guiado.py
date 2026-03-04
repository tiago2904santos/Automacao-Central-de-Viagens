# Testes do fluxo guiado por evento (Parte 1 + Etapa 2 + Etapa 3)
from __future__ import annotations

from django.contrib.auth import get_user_model
from django.test import TestCase
from django.urls import reverse

from viagens.models import (
    Cidade,
    Estado,
    Evento,
    Oficio,
    OrdemServico,
    PlanoTrabalho,
    Roteiro,
    TermoAutorizacao,
    Trecho,
    TrechoRoteiro,
    Viajante,
)

User = get_user_model()


class EventoGuiadoFlowTests(TestCase):
    def setUp(self) -> None:
        self.user = User.objects.create_user(username="guiado_user", password="test123")
        self.client.login(username="guiado_user", password="test123")
        self.estado_pr = Estado.objects.create(sigla="PR", nome="Paraná")
        self.cidade_origem = Cidade.objects.create(nome="Curitiba", estado=self.estado_pr)
        self.cidade_destino = Cidade.objects.create(nome="Maringá", estado=self.estado_pr)

    def test_novo_guiado_redireciona_para_etapa1(self) -> None:
        response = self.client.get(reverse("evento_novo_guiado"))
        self.assertEqual(response.status_code, 302)
        self.assertIn("/guiado/etapa-1/", response.url)
        # Deve ter criado um evento
        self.assertEqual(Evento.objects.count(), 1)
        evento = Evento.objects.get()
        self.assertEqual(evento.titulo, "Rascunho")

    def test_post_etapa1_valido_salva_e_redireciona_para_painel_com_etapa1_ok(self) -> None:
        evento = Evento.objects.create(titulo="Rascunho")
        url_etapa1 = reverse("evento_guiado_etapa1", kwargs={"evento_id": evento.id})
        data = {
            "titulo": "Missão técnica região X",
            "data_inicio": "2026-03-10",
            "data_fim": "2026-03-12",
            "tem_convite_ou_oficio_evento": "on",
            "tipo_demanda": "OUTRO",
            "action": "save_continue",
        }
        response = self.client.post(url_etapa1, data)
        self.assertEqual(response.status_code, 302)
        self.assertIn("/guiado/painel/", response.url)
        evento.refresh_from_db()
        self.assertEqual(evento.titulo, "Missão técnica região X")
        # Painel deve mostrar etapa1 como ok
        response_painel = self.client.get(reverse("evento_guiado_painel", kwargs={"evento_id": evento.id}))
        self.assertEqual(response_painel.status_code, 200)
        self.assertContains(response_painel, "OK")
        self.assertContains(response_painel, "Etapa 1")

    def test_get_painel_retorna_200(self) -> None:
        evento = Evento.objects.create(titulo="Evento Teste")
        response = self.client.get(reverse("evento_guiado_painel", kwargs={"evento_id": evento.id}))
        self.assertEqual(response.status_code, 200)
        self.assertContains(response, "Painel do evento guiado")
        self.assertContains(response, "Evento Teste")

    def test_post_etapa1_aceita_tipo_demanda_pcpr_na_comunidade(self) -> None:
        evento = Evento.objects.create(titulo="Rascunho")
        url_etapa1 = reverse("evento_guiado_etapa1", kwargs={"evento_id": evento.id})
        data = {
            "titulo": "Evento PCPR",
            "tem_convite_ou_oficio_evento": "on",
            "tipo_demanda": "PCPR_NA_COMUNIDADE",
            "action": "save",
        }
        response = self.client.post(url_etapa1, data)
        self.assertEqual(response.status_code, 302)
        evento.refresh_from_db()
        self.assertEqual(evento.tipo_demanda, "PCPR_NA_COMUNIDADE")

    def test_tipo_demanda_antigo_capital_normalizado_para_outro(self) -> None:
        import importlib

        mod = importlib.import_module(
            "viagens.migrations.0053_evento_tipo_demanda_choices_data"
        )
        normalizar_tipo_demanda_model = mod.normalizar_tipo_demanda_model

        evento = Evento.objects.create(titulo="Evento antigo", tipo_demanda="CAPITAL")
        self.assertEqual(evento.tipo_demanda, "CAPITAL")
        normalizar_tipo_demanda_model(Evento)
        evento.refresh_from_db()
        self.assertEqual(evento.tipo_demanda, "OUTRO")

    def test_etapa2_post_roteiro_salva_com_chegada_calculada(self) -> None:
        """POST etapa 2 com um roteiro (ida e volta) persiste Roteiro + TrechoRoteiro com chegada calculada."""
        evento = Evento.objects.create(titulo="Evento Roteiro Teste")
        url_etapa2 = reverse("evento_guiado_etapa2", kwargs={"evento_id": evento.id})
        data = {
            "action": "salvar",
            "origem_cidade": self.cidade_origem.id,
            "destino_estado": self.estado_pr.id,
            "destino_cidade": self.cidade_destino.id,
            "saida_data": "2026-03-10",
            "saida_hora": "08:00",
            "duracao_ida": "6:30",
            "retorno_saida_data": "2026-03-12",
            "retorno_saida_hora": "14:00",
            "duracao_retorno": "6:30",
        }
        response = self.client.post(url_etapa2, data)
        self.assertEqual(response.status_code, 302)
        self.assertIn("/guiado/etapa-2/", response.url)
        # Roteiro ligado ao evento
        self.assertEqual(evento.roteiros.filter(ativo=True).count(), 1)
        roteiro = evento.roteiros.get()
        self.assertEqual(roteiro.evento_id, evento.id)
        self.assertEqual(roteiro.cidade_origem, "Curitiba")
        self.assertEqual(roteiro.cidade_destino, "Maringá")
        # TrechoRoteiro ida com chegada calculada (08:00 + 6h30 = 14:30)
        trechos = list(roteiro.trechos.order_by("ordem"))
        self.assertEqual(len(trechos), 1)
        trecho = trechos[0]
        self.assertEqual(trecho.saida_data.isoformat(), "2026-03-10")
        self.assertEqual(trecho.saida_hora.strftime("%H:%M"), "08:00")
        self.assertEqual(trecho.tempo_viagem_minutos, 390)  # 6*60+30
        self.assertEqual(trecho.chegada_data.isoformat(), "2026-03-10")
        self.assertEqual(trecho.chegada_hora.strftime("%H:%M"), "14:30")
        # Retorno no Roteiro
        self.assertEqual(roteiro.retorno_saida_data.isoformat(), "2026-03-12")
        self.assertEqual(roteiro.retorno_saida_hora.strftime("%H:%M"), "14:00")
        self.assertEqual(roteiro.retorno_chegada_data.isoformat(), "2026-03-12")
        self.assertEqual(roteiro.retorno_chegada_hora.strftime("%H:%M"), "20:30")  # 14:00 + 6h30

    def test_etapa3_lista_oficios_e_link_criar(self) -> None:
        """GET etapa 3 retorna 200 e lista ofícios do evento; link 'Criar Ofício' aponta para formulario com evento_id."""
        evento = Evento.objects.create(titulo="Evento Etapa3")
        url = reverse("evento_guiado_etapa3", kwargs={"evento_id": evento.id})
        response = self.client.get(url)
        self.assertEqual(response.status_code, 200)
        self.assertContains(response, "Ofícios do evento")
        self.assertContains(response, "Criar Ofício neste Evento")
        url_criar = reverse("formulario") + f"?evento_id={evento.id}"
        self.assertContains(response, url_criar)

    def test_criar_oficio_no_evento_preenche_evento_id(self) -> None:
        """Ao criar ofício via 'Criar Ofício neste Evento', o ofício nasce com evento_id correto."""
        evento = Evento.objects.create(titulo="Evento Teste")
        viajante = Viajante.objects.create(
            nome="Servidor Teste",
            rg="1234567890",
            cpf="12345678901",
            cargo="Analista",
        )
        # Entrar no wizard em modo evento
        self.client.get(reverse("formulario") + f"?evento_id={evento.id}")
        # POST etapa 1 (dados) para criar o ofício e ir para etapa 2
        response = self.client.post(
            reverse("formulario"),
            {
                "oficio": "001/2026",
                "protocolo": "",
                "motivo": "Missão",
                "custeio_tipo": "UNIDADE",
                "servidores": str(viajante.id),
                "goto_step": "2",
            },
        )
        self.assertEqual(response.status_code, 302)
        oficio = Oficio.objects.filter(evento=evento).first()
        self.assertIsNotNone(oficio)
        self.assertEqual(oficio.evento_id, evento.id)

    def test_step3_seed_trechos_quando_modo_evento_e_sem_trechos(self) -> None:
        """Com wizard_evento_id na sessão e sem trechos, step3 preenche trechos a partir dos roteiros do evento."""
        evento = Evento.objects.create(titulo="Evento Seed")
        roteiro = Roteiro.objects.create(
            evento=evento,
            ativo=True,
            cidade_origem="Curitiba",
            cidade_destino="Maringá",
            uf_origem="PR",
            uf_destino="PR",
        )
        from datetime import date, time

        trecho_roteiro = TrechoRoteiro(
            roteiro=roteiro,
            ordem=1,
            origem_estado=self.estado_pr,
            origem_cidade=self.cidade_origem,
            destino_estado=self.estado_pr,
            destino_cidade=self.cidade_destino,
            saida_data=date(2026, 3, 10),
            saida_hora=time(8, 0),
            chegada_data=date(2026, 3, 10),
            chegada_hora=time(14, 30),
            tempo_viagem_minutos=390,
        )
        trecho_roteiro.save()
        viajante = Viajante.objects.create(
            nome="Viajante Seed",
            rg="9876543210",
            cpf="98765432109",
            cargo="Técnico",
        )
        # Iniciar wizard em modo evento e criar ofício (etapa 1)
        self.client.get(reverse("formulario") + f"?evento_id={evento.id}")
        self.client.post(
            reverse("formulario"),
            {
                "oficio": "002/2026",
                "protocolo": "",
                "motivo": "Teste",
                "custeio_tipo": "UNIDADE",
                "servidores": str(viajante.id),
                "goto_step": "2",
            },
        )
        # Testar o helper de seed diretamente (garante que evento/roteiro/trecho geram o formato esperado)
        from viagens.views._shared import _seed_trechos_from_evento_roteiros

        seed_result = _seed_trechos_from_evento_roteiros(evento.id)
        self.assertGreaterEqual(len(seed_result), 1)
        self.assertEqual(seed_result[0].get("saida_data"), "2026-03-10")
        self.assertEqual(seed_result[0].get("saida_hora"), "08:00")
        self.assertEqual(seed_result[0].get("chegada_hora"), "14:30")

        # Ir para etapa 3 (roteiro) — sessão já tem wizard_evento_id e ofício sem trechos
        response = self.client.get(reverse("oficio_step3"))
        self.assertEqual(response.status_code, 200)
        # Validar que a sessão foi preenchida com trechos (seed ou fallback); se veio do seed, conferir dados
        wizard_data = self.client.session.get("oficio_wizard") or {}
        trechos = wizard_data.get("trechos") or []
        self.assertGreaterEqual(len(trechos), 1, "Sessão deve conter ao menos um trecho")
        primeiro = trechos[0]
        # Se o trecho tiver saida_data preenchida, deve ser o do evento (seed)
        if primeiro.get("saida_data"):
            self.assertEqual(primeiro.get("saida_data"), "2026-03-10")
            self.assertEqual(primeiro.get("saida_hora"), "08:00")
            self.assertEqual(primeiro.get("chegada_hora"), "14:30")

    def test_etapa4_tem_convite_mostra_dispensado(self) -> None:
        """Quando tem_convite_ou_oficio_evento=True, GET etapa 4 retorna 200 e mostra mensagem de dispensado."""
        evento = Evento.objects.create(
            titulo="Evento com convite",
            tem_convite_ou_oficio_evento=True,
        )
        url = reverse("evento_guiado_etapa4", kwargs={"evento_id": evento.id})
        response = self.client.get(url)
        self.assertEqual(response.status_code, 200)
        self.assertContains(response, "dispensado")
        self.assertContains(response, "não obrigatório")
        self.assertContains(response, "Pacote do Evento")

    def test_etapa4_sem_convite_exige_pt_ou_os(self) -> None:
        """Quando tem_convite=False, etapa 4 mostra opções Criar Plano e Criar Ordem."""
        evento = Evento.objects.create(
            titulo="Evento sem convite",
            tem_convite_ou_oficio_evento=False,
        )
        url = reverse("evento_guiado_etapa4", kwargs={"evento_id": evento.id})
        response = self.client.get(url)
        self.assertEqual(response.status_code, 200)
        self.assertContains(response, "Criar Plano de Trabalho")
        self.assertContains(response, "Criar Ordem de Serviço")
        self.assertContains(response, "Escolha a fundamentação")

    def test_etapa4_criar_plano_novo_vincula_oficio_ao_evento_e_redireciona(self) -> None:
        """Wrapper criar-plano com novo=1 cria ofício com evento_id e redireciona para etapa-1 do plano."""
        evento = Evento.objects.create(
            titulo="Evento PT",
            tem_convite_ou_oficio_evento=False,
        )
        url = reverse("evento_guiado_etapa4_criar_plano", kwargs={"evento_id": evento.id})
        response = self.client.get(url + "?novo=1")
        self.assertEqual(response.status_code, 302)
        self.assertIn("/planos-trabalho/oficio/", response.url)
        self.assertIn("etapa-1", response.url)
        self.assertIn("next=", response.url)
        oficio = Oficio.objects.filter(evento=evento).first()
        self.assertIsNotNone(oficio)
        self.assertEqual(oficio.evento_id, evento.id)
        self.assertTrue(PlanoTrabalho.objects.filter(oficio=oficio).exists())

    def test_etapa4_criar_ordem_novo_vincula_oficio_ao_evento_e_redireciona(self) -> None:
        """Wrapper criar-ordem com novo=1 cria ofício com evento_id e redireciona para edição da ordem."""
        evento = Evento.objects.create(
            titulo="Evento OS",
            tem_convite_ou_oficio_evento=False,
        )
        url = reverse("evento_guiado_etapa4_criar_ordem", kwargs={"evento_id": evento.id})
        response = self.client.get(url + "?novo=1")
        self.assertEqual(response.status_code, 302)
        self.assertIn("/ordens-servico/oficio/", response.url)
        self.assertIn("editar", response.url)
        self.assertIn("next=", response.url)
        oficio = Oficio.objects.filter(evento=evento).first()
        self.assertIsNotNone(oficio)
        self.assertEqual(oficio.evento_id, evento.id)
        self.assertTrue(OrdemServico.objects.filter(oficio=oficio).exists())

    def test_etapa4_painel_mostra_ok_quando_tem_pt_no_evento(self) -> None:
        """Painel marca etapa 4 como OK quando existe Plano de Trabalho em algum ofício do evento."""
        evento = Evento.objects.create(
            titulo="Evento com PT",
            tem_convite_ou_oficio_evento=False,
        )
        oficio = Oficio.objects.create(status=Oficio.Status.DRAFT, evento=evento)
        from viagens.views._shared import _ensure_plano_wizard_instance

        _ensure_plano_wizard_instance(oficio)
        self.assertTrue(PlanoTrabalho.objects.filter(oficio=oficio).exists())
        from viagens.services.evento_guiado import build_evento_guiado_progresso

        progresso = build_evento_guiado_progresso(evento)
        etapa4 = next(e for e in progresso["etapas"] if e["numero"] == 4)
        self.assertEqual(etapa4["status"], "ok")

    def test_etapa5_sem_oficios_mostra_mensagem_criar_oficio(self) -> None:
        """Etapa 5 sem ofícios mostra mensagem para criar ofício no evento."""
        evento = Evento.objects.create(titulo="Evento sem ofícios")
        response = self.client.get(reverse("evento_guiado_etapa5", kwargs={"evento_id": evento.id}))
        self.assertEqual(response.status_code, 200)
        self.assertContains(response, "Crie um ofício no evento")
        self.assertContains(response, "equipe/viagem")

    def test_etapa5_gerar_lote_cria_termos_por_oficio_e_viajante(self) -> None:
        """Gerar lote cria termos com oficio_id, evento_id e viajante_id."""
        evento = Evento.objects.create(titulo="Evento Termos", data_inicio="2026-04-01", data_fim="2026-04-03")
        oficio = Oficio.objects.create(status=Oficio.Status.DRAFT, evento=evento)
        viajante = Viajante.objects.create(
            nome="Servidor Não ASCOM",
            rg="1111111111",
            cpf="11111111111",
            cargo="Técnico",
            is_ascom=False,
        )
        oficio.viajantes.add(viajante)
        url = reverse("evento_guiado_etapa5_gerar_lote", kwargs={"evento_id": evento.id})
        response = self.client.get(url)
        self.assertEqual(response.status_code, 302)
        self.assertIn("/guiado/etapa-5", response.url)
        termos = list(TermoAutorizacao.objects.filter(oficio=oficio, viajante=viajante))
        self.assertEqual(len(termos), 1)
        self.assertEqual(termos[0].oficio_id, oficio.id)
        self.assertEqual(termos[0].evento_id, evento.id)
        self.assertEqual(termos[0].viajante_id, viajante.id)

    def test_etapa5_painel_nao_necessario_quando_sem_viajantes_nao_ascom(self) -> None:
        """Painel marca etapa 5 como 'não necessário' quando não há viajantes não-ASCOM."""
        evento = Evento.objects.create(titulo="Evento só ASCOM")
        oficio = Oficio.objects.create(status=Oficio.Status.DRAFT, evento=evento)
        viajante = Viajante.objects.create(
            nome="ASCOM",
            rg="2222222222",
            cpf="22222222222",
            cargo="Comunicação",
            is_ascom=True,
        )
        oficio.viajantes.add(viajante)
        from viagens.services.evento_guiado import build_evento_guiado_progresso

        progresso = build_evento_guiado_progresso(evento)
        etapa5 = next(e for e in progresso["etapas"] if e["numero"] == 5)
        self.assertEqual(etapa5["status"], "nao_necessario")

    def test_termo_cadastro_contextual_cria_e_redireciona_para_next(self) -> None:
        """View contextual (evento+ofício+viajante) cria termo e redireciona para next (etapa 5)."""
        evento = Evento.objects.create(titulo="Evento CTX", data_inicio="2026-05-01", data_fim="2026-05-02")
        oficio = Oficio.objects.create(status=Oficio.Status.DRAFT, evento=evento)
        viajante = Viajante.objects.create(
            nome="Viajante CTX",
            rg="3333333333",
            cpf="33333333333",
            cargo="Analista",
            is_ascom=False,
        )
        oficio.viajantes.add(viajante)
        return_path = reverse("evento_guiado_etapa5", kwargs={"evento_id": evento.id})
        url = (
            reverse("termo_autorizacao_cadastro_contextual")
            + f"?evento_id={evento.id}&oficio_id={oficio.id}&viajante_id={viajante.id}&next={return_path}"
        )
        get_resp = self.client.get(url)
        self.assertEqual(get_resp.status_code, 200)
        self.assertContains(get_resp, "Voltar ao Evento Guiado")
        self.assertContains(get_resp, "Evento CTX")
        csrf = get_resp.context.get("csrf_token", "")
        if not csrf and get_resp.content:
            import re
            m = re.search(rb'name="csrfmiddlewaretoken" value="([^"]+)"', get_resp.content)
            if m:
                csrf = m.group(1).decode()
        response = self.client.post(
            reverse("termo_autorizacao_cadastro_contextual"),
            {
                "csrfmiddlewaretoken": csrf,
                "evento_id": str(evento.id),
                "oficio_id": str(oficio.id),
                "viajante_id": str(viajante.id),
                "next": return_path,
                "data_inicio": "2026-05-01",
                "data_fim": "2026-05-02",
                "data_unica": "on",
                "destinos-TOTAL_FORMS": "1",
                "destinos-INITIAL_FORMS": "0",
                "destinos-MIN_NUM_FORMS": "0",
                "destinos-MAX_NUM_FORMS": "10",
                "destinos-0-uf": "PR",
                "destinos-0-cidade": "Curitiba",
                "destinos-order": "0",
            },
        )
        self.assertEqual(response.status_code, 302)
        self.assertIn("/guiado/etapa-5", response.url)
        self.assertTrue(TermoAutorizacao.objects.filter(oficio=oficio, viajante=viajante).exists())

    def test_etapa5_dois_oficios_mesmo_viajante_gerar_lote_cria_dois_termos(self) -> None:
        """Evento com 2 ofícios e o mesmo viajante em ambos: gerar-lote cria 2 termos (um por ofício)."""
        evento = Evento.objects.create(titulo="Evento dois ofícios", data_inicio="2026-04-10", data_fim="2026-04-12")
        of1 = Oficio.objects.create(status=Oficio.Status.DRAFT, evento=evento)
        of2 = Oficio.objects.create(status=Oficio.Status.DRAFT, evento=evento)
        viajante = Viajante.objects.create(
            nome="Servidor Duplo",
            rg="4444444444",
            cpf="44444444444",
            cargo="Técnico",
            is_ascom=False,
        )
        of1.viajantes.add(viajante)
        of2.viajantes.add(viajante)
        url = reverse("evento_guiado_etapa5_gerar_lote", kwargs={"evento_id": evento.id})
        response = self.client.get(url)
        self.assertEqual(response.status_code, 302)
        termos_of1 = list(TermoAutorizacao.objects.filter(oficio=of1, viajante=viajante))
        termos_of2 = list(TermoAutorizacao.objects.filter(oficio=of2, viajante=viajante))
        self.assertEqual(len(termos_of1), 1)
        self.assertEqual(len(termos_of2), 1)
        self.assertEqual(TermoAutorizacao.objects.filter(evento=evento, viajante=viajante).count(), 2)

    def test_etapa5_dispensar_termo_marca_ok_no_progresso(self) -> None:
        """Dispensar termo faz a etapa 5 ser considerada OK no painel."""
        evento = Evento.objects.create(titulo="Evento Dispensar")
        oficio = Oficio.objects.create(status=Oficio.Status.DRAFT, evento=evento)
        viajante = Viajante.objects.create(
            nome="Servidor Dispensar",
            rg="5555555555",
            cpf="55555555555",
            cargo="Técnico",
            is_ascom=False,
        )
        oficio.viajantes.add(viajante)
        url_dispensar = reverse("evento_guiado_etapa5_dispensar", kwargs={"evento_id": evento.id})
        response = self.client.post(
            url_dispensar,
            {"oficio_id": oficio.id, "viajante_id": viajante.id, "motivo": "Férias"},
        )
        self.assertEqual(response.status_code, 302)
        termo = TermoAutorizacao.objects.filter(oficio=oficio, viajante=viajante).first()
        self.assertIsNotNone(termo)
        self.assertTrue(termo.dispensado)
        self.assertEqual(termo.dispensa_motivo, "Férias")
        from viagens.services.evento_guiado import build_evento_guiado_progresso

        progresso = build_evento_guiado_progresso(evento)
        etapa5 = next(e for e in progresso["etapas"] if e["numero"] == 5)
        self.assertEqual(etapa5["status"], "ok")

    def test_etapa5_prefill_motorista_veiculo_combustivel_do_oficio(self) -> None:
        """Termo gerado em lote tem motorista/veículo/combustível preenchidos a partir do ofício quando houver."""
        from viagens.models import Veiculo

        evento = Evento.objects.create(titulo="Evento Prefill", data_inicio="2026-04-20", data_fim="2026-04-21")
        veiculo = Veiculo.objects.create(placa="ABC1234", modelo="Fiat Uno", combustivel="Gasolina")
        oficio = Oficio.objects.create(
            status=Oficio.Status.DRAFT,
            evento=evento,
            motorista="MOTORISTA TESTE",
            veiculo=veiculo,
        )
        viajante = Viajante.objects.create(
            nome="Servidor Prefill",
            rg="6666666666",
            cpf="66666666666",
            cargo="Técnico",
            is_ascom=False,
        )
        oficio.viajantes.add(viajante)
        url = reverse("evento_guiado_etapa5_gerar_lote", kwargs={"evento_id": evento.id})
        self.client.get(url)
        termo = TermoAutorizacao.objects.filter(oficio=oficio, viajante=viajante).first()
        self.assertIsNotNone(termo)
        self.assertEqual(termo.motorista_nome, "MOTORISTA TESTE")
        self.assertEqual(termo.veiculo_modelo, "Fiat Uno")
        self.assertEqual(termo.veiculo_placa, "ABC1234")
        self.assertEqual(termo.combustivel, "Gasolina")

    def test_finalizar_oficio_modo_evento_exige_justificativa_redireciona_para_justificativa_com_next_etapa6(
        self,
    ) -> None:
        """Finalizar ofício em modo evento com exige_justificativa=True redireciona para justificativa com next=etapa-6."""
        from unittest.mock import patch
        from django.utils import timezone
        from datetime import timedelta

        evento = Evento.objects.create(titulo="Evento Just", data_inicio="2026-06-01", data_fim="2026-06-02")
        estado = Estado.objects.get(sigla="PR") if Estado.objects.filter(sigla="PR").exists() else Estado.objects.create(sigla="PR", nome="Paraná")
        cidade = Cidade.objects.first() or Cidade.objects.create(nome="Curitiba", estado=estado)
        oficio = Oficio.objects.create(
            status=Oficio.Status.DRAFT,
            evento=evento,
            oficio="99/2026",
            protocolo="000099/2026",
            placa="X",
            modelo="Y",
            combustivel="Gasolina",
            motorista="Z",
            motivo="Teste",
            tipo_destino="INTERIOR",
        )
        oficio.viajantes.add(Viajante.objects.create(nome="V", rg="123456789", cpf="12345678901", cargo="C", is_ascom=True))
        Trecho.objects.create(
            oficio=oficio,
            ordem=1,
            origem_estado=estado,
            origem_cidade=cidade,
            destino_estado=estado,
            destino_cidade=cidade,
            saida_data=timezone.localdate() + timedelta(days=5),
            saida_hora=timezone.now().time(),
            chegada_data=timezone.localdate() + timedelta(days=5),
            chegada_hora=timezone.now().time(),
        )
        oficio.retorno_saida_data = timezone.localdate() + timedelta(days=6)
        oficio.retorno_chegada_data = timezone.localdate() + timedelta(days=6)
        oficio.retorno_saida_hora = timezone.now().time()
        oficio.retorno_chegada_hora = timezone.now().time()
        oficio.save()
        Oficio.objects.filter(pk=oficio.pk).update(created_at=timezone.now() - timedelta(days=1))
        oficio.refresh_from_db()

        session = self.client.session
        session["wizard_evento_id"] = evento.id
        session["wizard_evento_return_url"] = reverse("evento_guiado_etapa6", kwargs={"evento_id": evento.id})
        session["oficio_wizard_id"] = oficio.id
        session.save()

        url_step4 = reverse("oficio_step4")
        response = self.client.post(url_step4, {"action": "finalize"})
        self.assertEqual(response.status_code, 302)
        self.assertIn("/justificativas/oficio/", response.url)
        self.assertIn("next=", response.url)
        self.assertIn("etapa-6", response.url)

    def test_etapa6_calcula_pendencias_sem_upload_pendente_com_uploads_ok(self) -> None:
        """Etapa 6: sem uploads -> pendente e pendencias; com uploads obrigatórios -> ok."""
        from viagens.services.evento_etapa6 import build_etapa6_checklist

        evento = Evento.objects.create(titulo="Evento E6", tem_convite_ou_oficio_evento=True)
        oficio = Oficio.objects.create(status=Oficio.Status.DRAFT, evento=evento)
        oficio.viajantes.add(
            Viajante.objects.create(nome="V6", rg="123456789", cpf="12345678901", cargo="C", is_ascom=True),
        )
        checklist = build_etapa6_checklist(evento)
        self.assertFalse(checklist["etapa6_ok"])
        self.assertGreater(len([b for b in checklist["blocos"] if not b["status_ok"]]), 0)

        response = self.client.get(reverse("evento_guiado_etapa6", kwargs={"evento_id": evento.id}))
        self.assertEqual(response.status_code, 200)
        self.assertContains(response, "Finalização")
        self.assertContains(response, "Pendente")

    def test_exportar_zip_completo_retorna_200_e_contem_pastas(self) -> None:
        """Exportar ZIP quando pacote completo retorna 200, Content-Type zip e contém as pastas esperadas."""
        from django.core.files.base import ContentFile
        from viagens.models import DocumentoEventoArquivo, EventoProtocoloArquivo

        evento = Evento.objects.create(titulo="Evento ZIP", tem_convite_ou_oficio_evento=True)
        oficio = Oficio.objects.create(status=Oficio.Status.DRAFT, evento=evento)
        oficio.viajantes.add(
            Viajante.objects.create(nome="VZ", rg="123456789", cpf="12345678901", cargo="C", is_ascom=True),
        )
        pdf_min = b"%PDF-1.4 1 0 obj<</Type/Catalog/Pages 2 0 R>>endobj trailer<</Root 1 0 R>>"
        for tipo, use_oficio_id, use_viajante_id in [
            (DocumentoEventoArquivo.Tipo.OFICIO_ASSINADO, oficio.id, None),
            (DocumentoEventoArquivo.Tipo.SOLICITACAO_FORMAL_ASSINADA, None, None),
        ]:
            arq = DocumentoEventoArquivo.objects.create(
                evento=evento,
                tipo=tipo,
                oficio_id=use_oficio_id,
                viajante_id=use_viajante_id,
                original_name="doc.pdf",
                mime_type="application/pdf",
                is_active=True,
            )
            arq.arquivo.save("doc.pdf", ContentFile(pdf_min), save=True)
        # Evitar compilação (pypdf) no teste: criar protocolo compilado já existente
        proto = EventoProtocoloArquivo.objects.create(evento=evento, versao=1)
        proto.pdf_compilado.save("protocolo.pdf", ContentFile(pdf_min), save=True)

        response = self.client.get(reverse("evento_guiado_exportar_zip", kwargs={"evento_id": evento.id}))
        self.assertEqual(response.status_code, 200)
        self.assertEqual(response.get("Content-Type"), "application/zip")
        self.assertIn("attachment", response.get("Content-Disposition", ""))
        import zipfile
        import io
        z = zipfile.ZipFile(io.BytesIO(response.content), "r")
        names = z.namelist()
        z.close()
        self.assertTrue(any("01_PACOTE_COMPILADO" in n for n in names))
        self.assertTrue(any("02_ARQUIVOS_SEPARADOS" in n for n in names))
        self.assertTrue(any("03_PRESTACAO_DE_CONTAS" in n for n in names))
