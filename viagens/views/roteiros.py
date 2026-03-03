from __future__ import annotations

from datetime import datetime, time as dt_time, timedelta
import json

from django.contrib.auth.decorators import login_required
from django.contrib import messages
from django.core.paginator import Paginator
from django.db import transaction
from django.db.models import Q
from django.http import JsonResponse
from django.urls import reverse
from django.shortcuts import get_object_or_404, redirect, render
from django.utils.dateparse import parse_date, parse_time
from django.views.decorators.http import require_GET, require_http_methods, require_POST

from ..forms import OficioVincularRoteiroForm, RoteiroForm, TrechoRoteiroFormSet
from ..models import Cidade, Estado, Oficio, OficioRoteiro, Roteiro, TrechoRoteiro


FORMSET_PREFIX = "trechos"


def _request_payload(request):
    if "application/json" in (request.content_type or ""):
        try:
            raw_body = request.body.decode("utf-8") if request.body else "{}"
            data = json.loads(raw_body or "{}")
            if isinstance(data, dict):
                return data
        except (UnicodeDecodeError, json.JSONDecodeError):
            return None
    return request.POST


def _resolve_estado(value):
    if not value:
        return None
    try:
        return Estado.objects.filter(pk=int(value)).first()
    except (TypeError, ValueError):
        return Estado.objects.filter(sigla__iexact=str(value).strip()).first()


def _resolve_cidade(cidade_id):
    if not cidade_id:
        return None
    try:
        return Cidade.objects.filter(pk=int(cidade_id)).select_related("estado").first()
    except (TypeError, ValueError):
        return None


def _resolve_cidade_from_label(nome: str, uf_sigla: str):
    nome = (nome or "").strip()
    uf_sigla = (uf_sigla or "").strip()
    if not nome:
        return None
    qs = Cidade.objects.filter(nome__iexact=nome)
    if uf_sigla:
        qs = qs.filter(estado__sigla__iexact=uf_sigla)
    return qs.select_related("estado").first()


def _resolve_cidade_by_name(nome, estado):
    if not nome or not estado:
        return None
    try:
        return Cidade.objects.select_related("estado").get(
            nome__iexact=str(nome).strip(),
            estado=estado,
        )
    except Cidade.DoesNotExist:
        return None


def _normalize_tipo_deslocamento(raw_value: str):
    normalized = (raw_value or "").strip().upper()
    if normalized in {
        Roteiro.TipoDeslocamentoChoices.INTERIOR,
        Roteiro.TipoDeslocamentoChoices.CAPITAL,
    }:
        return normalized
    if normalized == "BRASILIA":
        return Roteiro.TipoDeslocamentoChoices.OUTRO
    return Roteiro.TipoDeslocamentoChoices.OUTRO


def _serialize_roteiro_search_item(roteiro: Roteiro):
    return {
        "id": roteiro.pk,
        "nome": roteiro.nome,
        "destinos": roteiro.get_destinos_display(),
        "text": roteiro.nome,
        "origem": f"{roteiro.cidade_origem}/{roteiro.uf_origem}",
        "destino": f"{roteiro.cidade_destino}/{roteiro.uf_destino}",
        "tipo": roteiro.get_tipo_deslocamento_display(),
        "distancia_km": str(roteiro.get_distancia_total())
        if roteiro.get_distancia_total() is not None
        else None,
        "uf_sede_id": roteiro.uf_sede_id,
        "cidade_sede_id": roteiro.cidade_sede_id,
        "cidade_sede_nome": roteiro.cidade_sede_nome,
        "uf_sede_sigla": roteiro.uf_origem,
    }


def _serialize_roteiro_cards_payload(roteiro: Roteiro):
    return {
        "id": roteiro.pk,
        "nome": roteiro.nome,
        "uf_sede": roteiro.uf_origem,
        "cidade_sede": roteiro.cidade_origem,
        "tempo_viagem": _serialize_duration_time(_resolve_roteiro_tempo_viagem(roteiro)),
        "cards": roteiro.get_cards_gerados(),
    }


def _serialize_duration_time(value):
    if not value:
        return ""
    return value.strftime("%H:%M")


def _resolve_roteiro_tempo_viagem(roteiro: Roteiro | None):
    if not roteiro:
        return None
    if roteiro.tempo_viagem:
        return roteiro.tempo_viagem.replace(second=0, microsecond=0)

    primeiro_trecho = roteiro.trechos.order_by("ordem", "id").first()
    if primeiro_trecho and primeiro_trecho.tempo_viagem_minutos is not None:
        return _parse_tempo_viagem_time(primeiro_trecho.tempo_viagem_minutos)

    retorno_minutos = _retorno_duration_minutes(
        roteiro.retorno_saida_data,
        roteiro.retorno_saida_hora,
        roteiro.retorno_chegada_data,
        roteiro.retorno_chegada_hora,
    )
    return _parse_tempo_viagem_time(retorno_minutos)


def _build_roteiro_list_context(request):
    roteiros_qs = (
        Roteiro.objects.filter(ativo=True)
        .prefetch_related("trechos")
        .order_by("-criado_em", "-id")
    )
    query = (request.GET.get("q") or "").strip()
    if query:
        roteiros_qs = roteiros_qs.filter(
            Q(nome__icontains=query)
            | Q(cidade_origem__icontains=query)
            | Q(cidade_destino__icontains=query)
            | Q(uf_origem__icontains=query)
            | Q(uf_destino__icontains=query)
            | Q(trechos__cidade_destino__icontains=query)
        ).distinct()

    paginator = Paginator(roteiros_qs, 20)
    page_obj = paginator.get_page(request.GET.get("page"))
    return {
        "roteiros": page_obj.object_list,
        "page_obj": page_obj,
        "is_paginated": page_obj.has_other_pages(),
        "query": query,
        "total_roteiros": paginator.count,
    }


def _build_roteiro_form_context():
    estados = list(Estado.objects.order_by("sigla"))
    estado_padrao = next((estado for estado in estados if estado.sigla == "PR"), None)
    if estado_padrao is None and estados:
        estado_padrao = estados[0]

    cidades_iniciais = (
        list(Cidade.objects.filter(estado=estado_padrao).order_by("nome"))
        if estado_padrao
        else []
    )
    cidade_sede_padrao = next(
        (
            cidade
            for cidade in cidades_iniciais
            if cidade.nome.strip().lower() == "curitiba"
        ),
        None,
    )
    if cidade_sede_padrao is None and cidades_iniciais:
        cidade_sede_padrao = cidades_iniciais[0]

    return {
        "estados": estados,
        "estado_padrao": estado_padrao.sigla if estado_padrao else "PR",
        "cidades_iniciais": cidades_iniciais,
        "cidade_sede_padrao_id": cidade_sede_padrao.id if cidade_sede_padrao else None,
    }


def _build_save_error(message: str, status: int = 400, key: str = "erro"):
    payload = {
        "ok": False,
        "sucesso": False,
        key: message,
    }
    alt_key = "error" if key == "erro" else "erro"
    payload[alt_key] = message
    return JsonResponse(payload, status=status)


def _parse_optional_date(raw_value, label: str):
    value = str(raw_value or "").strip()
    if not value:
        return None
    try:
        return datetime.strptime(value, "%Y-%m-%d").date()
    except ValueError as exc:
        raise ValueError(f"{label} invalida.") from exc


def _parse_optional_time(raw_value, label: str):
    value = str(raw_value or "").strip()
    if not value:
        return None
    try:
        return datetime.strptime(value, "%H:%M").time()
    except ValueError as exc:
        raise ValueError(f"{label} invalida.") from exc


def _resolve_cidade_payload(value, uf_sigla: str = ""):
    cidade = _resolve_cidade(value)
    if cidade:
        return cidade
    value_str = str(value or "").strip()
    if not value_str:
        return None
    return _resolve_cidade_from_label(value_str, uf_sigla)


def _payload_as_dict(request):
    try:
        raw_body = request.body.decode("utf-8") if request.body else "{}"
        payload = json.loads(raw_body or "{}")
        if isinstance(payload, dict):
            return payload, None
    except (UnicodeDecodeError, json.JSONDecodeError, ValueError) as exc:
        if request.POST:
            return request.POST, None
        return None, f"Erro ao processar requisicao: {exc}"
    return None, "Payload JSON invalido."


def _create_roteiro_from_payload(data: dict) -> tuple[Roteiro | None, str | None]:
    sede_uf = (data.get("sede_uf") or data.get("uf_sede") or "PR").strip().upper() or "PR"
    sede_cidade_value = data.get("sede_cidade") or data.get("cidade_sede_id") or data.get("cidade_sede")
    destinos_data = data.get("destinos") or []
    trechos_data = data.get("trechos") or []
    retorno_data = data.get("retorno") or {}
    nome_override = (data.get("nome") or "").strip()

    if not isinstance(destinos_data, list) or not destinos_data:
        return None, "Informe ao menos um destino."

    sede_cidade = _resolve_cidade_payload(sede_cidade_value, sede_uf)
    if not sede_cidade:
        return None, "Informe a cidade sede."

    destinos_normalizados = []
    for index, destino_item in enumerate(destinos_data, start=1):
        if not isinstance(destino_item, dict):
            return None, f"Destino {index} invalido."

        uf_destino = (destino_item.get("uf") or "PR").strip().upper() or "PR"
        cidade_value = (
            destino_item.get("cidade")
            or destino_item.get("cidade_id")
            or destino_item.get("cidade_nome")
        )
        cidade_destino = _resolve_cidade_payload(cidade_value, uf_destino)
        if not cidade_destino:
            return None, f"Informe uma cidade valida para o destino {index}."

        trecho_item = trechos_data[index - 1] if index - 1 < len(trechos_data) else {}
        if not isinstance(trecho_item, dict):
            trecho_item = {}

        try:
            saida_data = _parse_optional_date(
                trecho_item.get("saida_data"),
                f"Data de saida do trecho {index}",
            )
            saida_hora = _parse_optional_time(
                trecho_item.get("saida_hora"),
                f"Hora de saida do trecho {index}",
            )
            chegada_data = _parse_optional_date(
                trecho_item.get("chegada_data"),
                f"Data de chegada do trecho {index}",
            )
            chegada_hora = _parse_optional_time(
                trecho_item.get("chegada_hora"),
                f"Hora de chegada do trecho {index}",
            )
        except ValueError as exc:
            return None, str(exc)

        destinos_normalizados.append(
            {
                "ordem": index,
                "cidade": cidade_destino,
                "uf": uf_destino,
                "modal": trecho_item.get("modal")
                or destino_item.get("modal")
                or TrechoRoteiro.ModalChoices.VEICULO_PROPRIO,
                "distancia_km": trecho_item.get("distancia_km")
                or destino_item.get("distancia_km")
                or None,
                "observacao": (
                    trecho_item.get("observacao")
                    or destino_item.get("observacao")
                    or ""
                ).strip(),
                "saida_data": saida_data,
                "saida_hora": saida_hora,
                "chegada_data": chegada_data,
                "chegada_hora": chegada_hora,
            }
        )

    if not nome_override:
        referencia_datahora = datetime.now()
        primeiro_destino = destinos_normalizados[0]
        if primeiro_destino["saida_data"] and primeiro_destino["saida_hora"]:
            referencia_datahora = datetime.combine(
                primeiro_destino["saida_data"],
                primeiro_destino["saida_hora"],
            )
        nome_override = (
            f"{sede_cidade.nome} > {primeiro_destino['cidade'].nome} "
            f"{referencia_datahora.strftime('%d/%m/%Y %H:%M')}"
        )

    if len(nome_override) < 3:
        return None, "Nome do roteiro e obrigatorio."

    try:
        retorno_saida_data = _parse_optional_date(
            retorno_data.get("saida_data"),
            "Data de saida do retorno",
        )
        retorno_saida_hora = _parse_optional_time(
            retorno_data.get("saida_hora"),
            "Hora de saida do retorno",
        )
        retorno_chegada_data = _parse_optional_date(
            retorno_data.get("chegada_data"),
            "Data de chegada do retorno",
        )
        retorno_chegada_hora = _parse_optional_time(
            retorno_data.get("chegada_hora"),
            "Hora de chegada do retorno",
        )
    except ValueError as exc:
        return None, str(exc)

    ultimo_destino = destinos_normalizados[-1]
    with transaction.atomic():
        roteiro = Roteiro.objects.create(
            nome=nome_override,
            descricao=(data.get("descricao") or "").strip(),
            estado_sede=sede_cidade.estado,
            cidade_sede=sede_cidade,
            uf_origem=sede_uf,
            cidade_origem=sede_cidade.nome,
            uf_destino=ultimo_destino["uf"],
            cidade_destino=ultimo_destino["cidade"].nome,
            retorno_saida_cidade=ultimo_destino["cidade"].nome,
            tipo_deslocamento=_normalize_tipo_deslocamento(
                data.get("tipo_destino") or Roteiro.TipoDeslocamentoChoices.INTERIOR
            ),
            retorno_saida_data=retorno_saida_data,
            retorno_saida_hora=retorno_saida_hora,
            retorno_chegada_cidade=sede_cidade.nome,
            retorno_chegada_data=retorno_chegada_data,
            retorno_chegada_hora=retorno_chegada_hora,
            ativo=True,
        )

        origem_uf = sede_uf
        origem_cidade = sede_cidade.nome
        origem_estado_obj = sede_cidade.estado
        origem_cidade_obj = sede_cidade
        for destino in destinos_normalizados:
            TrechoRoteiro.objects.create(
                roteiro=roteiro,
                ordem=destino["ordem"],
                origem_estado=origem_estado_obj,
                origem_cidade=origem_cidade_obj,
                uf_origem=origem_uf,
                cidade_origem=origem_cidade,
                destino_estado=destino["cidade"].estado,
                destino_cidade=destino["cidade"],
                uf_destino=destino["uf"],
                cidade_destino=destino["cidade"].nome,
                distancia_km=destino["distancia_km"],
                modal=destino["modal"],
                observacao=destino["observacao"],
                saida_data=destino["saida_data"],
                saida_hora=destino["saida_hora"],
                chegada_data=destino["chegada_data"],
                chegada_hora=destino["chegada_hora"],
            )
            origem_uf = destino["uf"]
            origem_cidade = destino["cidade"].nome
            origem_estado_obj = destino["cidade"].estado
            origem_cidade_obj = destino["cidade"]

    return roteiro, None


def _get_sede_cidade_default():
    return (
        Cidade.objects.select_related("estado")
        .filter(nome__iexact="Curitiba", estado__sigla__iexact="PR")
        .first()
    )


def _get_sede_cidade_default_id():
    cidade = _get_sede_cidade_default()
    return cidade.id if cidade else None


def _serialize_sede_destinos_roteiro(post_data):
    sede_uf = (post_data.get("sede_uf") or "").strip().upper()
    sede_cidade_id = (post_data.get("sede_cidade") or "").strip()

    try:
        total_forms = int(post_data.get("destinos-TOTAL_FORMS") or 0)
    except (TypeError, ValueError):
        total_forms = 0

    destinos = []
    for index in range(total_forms):
        uf = (post_data.get(f"destinos-{index}-uf") or "").strip().upper()
        cidade_id = (post_data.get(f"destinos-{index}-cidade") or "").strip()
        if cidade_id:
            destinos.append({"uf": uf, "cidade_id": cidade_id})

    return sede_uf, sede_cidade_id, destinos


def _resolve_cidade_from_post_id(raw_value):
    try:
        return Cidade.objects.select_related("estado").get(pk=int(raw_value))
    except (Cidade.DoesNotExist, TypeError, ValueError):
        return None


def _resolve_estado_from_post_sigla(sigla):
    if not sigla:
        return None
    return Estado.objects.filter(sigla__iexact=str(sigla).strip()).first()


def _build_roteiro_form_template_context(roteiro=None):
    estados = Estado.objects.order_by("sigla")
    sede_padrao = _get_sede_cidade_default()

    if roteiro is None:
        roteiro = Roteiro()
        sede_cidade = sede_padrao
        sede_uf = sede_cidade.estado.sigla if sede_cidade and sede_cidade.estado else "PR"
        destinos = []
        formset = TrechoRoteiroFormSet(instance=roteiro, prefix=FORMSET_PREFIX)
        retorno_saida_data = ""
        retorno_saida_hora = ""
        retorno_chegada_data = ""
        retorno_chegada_hora = ""
        retorno_saida_cidade = ""
        retorno_chegada_cidade = ""
    else:
        roteiro = (
            Roteiro.objects.prefetch_related("trechos")
            .select_related("estado_sede", "cidade_sede")
            .get(pk=roteiro.pk)
        )
        sede_cidade = roteiro.cidade_sede_obj
        sede_uf = roteiro.uf_origem or "PR"
        formset = TrechoRoteiroFormSet(instance=roteiro, prefix=FORMSET_PREFIX)
        destinos = []
        for trecho in roteiro.trechos.order_by("ordem"):
            cidade_destino = trecho.destino_cidade or trecho.cidade_destino_obj
            if not cidade_destino:
                continue
            destinos.append(
                {
                    "uf": trecho.destino_estado_sigla,
                    "cidade": str(cidade_destino.id),
                    "cidade_label": f"{cidade_destino.nome}/{cidade_destino.estado.sigla}",
                }
            )
        retorno_saida_data = roteiro.retorno_saida_data
        retorno_saida_hora = roteiro.retorno_saida_hora
        retorno_chegada_data = roteiro.retorno_chegada_data
        retorno_chegada_hora = roteiro.retorno_chegada_hora
        retorno_saida_cidade = roteiro.retorno_saida_cidade or ""
        retorno_chegada_cidade = roteiro.retorno_chegada_cidade or ""

    return {
        "roteiro": roteiro if roteiro.pk else None,
        "estados": estados,
        "formset": formset,
        "destinos": destinos,
        "destinos_total_forms": len(destinos),
        "destinos_order": ",".join(str(index) for index in range(len(destinos))),
        "sede_uf": sede_uf,
        "sede_label": f"{sede_cidade.nome}/{sede_cidade.estado.sigla}" if sede_cidade and sede_cidade.estado else "",
        "sede_cidade": str(sede_cidade.id) if sede_cidade else str(_get_sede_cidade_default_id() or ""),
        "retorno_saida_cidade": retorno_saida_cidade,
        "retorno_saida_data": retorno_saida_data,
        "retorno_saida_hora": retorno_saida_hora,
        "retorno_chegada_cidade": retorno_chegada_cidade,
        "retorno_chegada_data": retorno_chegada_data,
        "retorno_chegada_hora": retorno_chegada_hora,
        "is_edit": bool(roteiro.pk),
    }


def _parse_date_post_value(post_data, key):
    value = (post_data.get(key) or "").strip()
    return parse_date(value) if value else None


def _parse_time_post_value(post_data, key):
    value = (post_data.get(key) or "").strip()
    return parse_time(value) if value else None


def _save_roteiro_from_post(request, roteiro=None):
    is_ajax = (
        request.headers.get("X-Requested-With") == "XMLHttpRequest"
        or "application/json" in (request.headers.get("Accept") or "")
    )

    try:
        with transaction.atomic():
            sede_uf_sigla, sede_cidade_id, destinos_data = _serialize_sede_destinos_roteiro(
                request.POST
            )

            if not sede_uf_sigla:
                raise ValueError("Informe a UF da sede.")

            estado_sede = _resolve_estado_from_post_sigla(sede_uf_sigla)
            cidade_sede = _resolve_cidade_from_post_id(sede_cidade_id)
            if not estado_sede or not cidade_sede:
                raise ValueError("Informe uma cidade sede valida.")

            if not destinos_data:
                raise ValueError("Adicione ao menos um destino.")

            destinos_resolvidos = []
            for index, destino_item in enumerate(destinos_data, start=1):
                cidade_destino = _resolve_cidade_from_post_id(destino_item["cidade_id"])
                if not cidade_destino:
                    raise ValueError(f"Destino {index} invalido.")
                destinos_resolvidos.append(cidade_destino)

            if roteiro is None:
                roteiro = Roteiro()

            tempo_viagem_source = request.POST.get("tempo_viagem")
            if tempo_viagem_source in (None, ""):
                tempo_viagem_source = request.POST.get("retorno_tempo_viagem_minutos")
            if tempo_viagem_source in (None, ""):
                try:
                    total_trechos = int(request.POST.get(f"{FORMSET_PREFIX}-TOTAL_FORMS") or 0)
                except (TypeError, ValueError):
                    total_trechos = 0
                for index in range(total_trechos):
                    tempo_viagem_source = request.POST.get(
                        f"{FORMSET_PREFIX}-{index}-tempo_viagem_minutos"
                    )
                    if tempo_viagem_source not in (None, ""):
                        break
            tempo_viagem = _parse_tempo_viagem_time(tempo_viagem_source)
            tempo_viagem_minutos = _parse_tempo_viagem(tempo_viagem)
            roteiro.estado_sede = estado_sede
            roteiro.cidade_sede = cidade_sede
            roteiro.ativo = True
            if not roteiro.tipo_deslocamento:
                roteiro.tipo_deslocamento = Roteiro.TipoDeslocamentoChoices.INTERIOR
            roteiro.retorno_saida_data = _parse_date_post_value(request.POST, "retorno_saida_data")
            roteiro.retorno_saida_hora = _parse_time_post_value(request.POST, "retorno_saida_hora")
            roteiro.retorno_chegada_data = _parse_date_post_value(request.POST, "retorno_chegada_data")
            roteiro.retorno_chegada_hora = _parse_time_post_value(request.POST, "retorno_chegada_hora")
            (
                roteiro.retorno_chegada_data,
                roteiro.retorno_chegada_hora,
            ) = _calculate_arrival_from_duration(
                roteiro.retorno_saida_data,
                roteiro.retorno_saida_hora,
                roteiro.retorno_chegada_data,
                roteiro.retorno_chegada_hora,
                tempo_viagem_minutos,
            )
            roteiro.retorno_saida_cidade = request.POST.get("retorno_saida_cidade", "").strip()
            roteiro.retorno_chegada_cidade = request.POST.get("retorno_chegada_cidade", "").strip()
            roteiro.uf_destino = destinos_resolvidos[-1].estado.sigla
            roteiro.cidade_destino = destinos_resolvidos[-1].nome
            roteiro.tempo_viagem = tempo_viagem
            roteiro.save()

            roteiro.trechos.all().delete()

            try:
                total_trechos = int(request.POST.get(f"{FORMSET_PREFIX}-TOTAL_FORMS") or 0)
            except (TypeError, ValueError):
                total_trechos = 0

            ultimo_trecho_obj = None
            for index in range(total_trechos):
                origem_estado = _resolve_estado(request.POST.get(f"{FORMSET_PREFIX}-{index}-origem_estado"))
                origem_cidade = _resolve_cidade(request.POST.get(f"{FORMSET_PREFIX}-{index}-origem_cidade"))
                destino_estado = _resolve_estado(request.POST.get(f"{FORMSET_PREFIX}-{index}-destino_estado"))
                destino_cidade = _resolve_cidade(request.POST.get(f"{FORMSET_PREFIX}-{index}-destino_cidade"))

                if not origem_cidade or not destino_cidade:
                    continue

                if not origem_estado:
                    origem_estado = origem_cidade.estado
                if not destino_estado:
                    destino_estado = destino_cidade.estado

                saida_data = _parse_date_post_value(
                    request.POST,
                    f"{FORMSET_PREFIX}-{index}-saida_data",
                )
                saida_hora = _parse_time_post_value(
                    request.POST,
                    f"{FORMSET_PREFIX}-{index}-saida_hora",
                )
                chegada_data = _parse_date_post_value(
                    request.POST,
                    f"{FORMSET_PREFIX}-{index}-chegada_data",
                )
                chegada_hora = _parse_time_post_value(
                    request.POST,
                    f"{FORMSET_PREFIX}-{index}-chegada_hora",
                )
                chegada_data, chegada_hora = _calculate_arrival_from_duration(
                    saida_data,
                    saida_hora,
                    chegada_data,
                    chegada_hora,
                    tempo_viagem_minutos,
                )

                ultimo_trecho_obj = TrechoRoteiro.objects.create(
                    roteiro=roteiro,
                    ordem=index + 1,
                    origem_estado=origem_estado,
                    origem_cidade=origem_cidade,
                    destino_estado=destino_estado,
                    destino_cidade=destino_cidade,
                    saida_data=saida_data,
                    saida_hora=saida_hora,
                    chegada_data=chegada_data,
                    chegada_hora=chegada_hora,
                    tempo_viagem_minutos=tempo_viagem_minutos,
                )

            if ultimo_trecho_obj:
                ultimo_trecho_obj.retorno_saida_data = roteiro.retorno_saida_data
                ultimo_trecho_obj.retorno_saida_hora = roteiro.retorno_saida_hora
                ultimo_trecho_obj.retorno_chegada_data = roteiro.retorno_chegada_data
                ultimo_trecho_obj.retorno_chegada_hora = roteiro.retorno_chegada_hora
                ultimo_trecho_obj.save(
                    update_fields=[
                        "retorno_saida_data",
                        "retorno_saida_hora",
                        "retorno_chegada_data",
                        "retorno_chegada_hora",
                    ]
                )

            if not roteiro.trechos.exists():
                raise ValueError("Nenhum trecho valido foi gerado.")

            ultimo_trecho = roteiro.trechos.order_by("ordem").last()
            if ultimo_trecho:
                roteiro.uf_destino = ultimo_trecho.destino_estado_sigla or roteiro.uf_destino
                roteiro.cidade_destino = ultimo_trecho.destino_cidade_nome or roteiro.cidade_destino
                roteiro.retorno_saida_cidade = (
                    roteiro.retorno_saida_cidade or ultimo_trecho.destino_cidade_nome
                )
            roteiro.retorno_chegada_cidade = (
                roteiro.retorno_chegada_cidade or cidade_sede.nome
            )

            nome_gerado = roteiro.gerar_nome()
            Roteiro.objects.filter(pk=roteiro.pk).update(
                nome=nome_gerado,
                uf_destino=roteiro.uf_destino,
                cidade_destino=roteiro.cidade_destino,
                retorno_saida_cidade=roteiro.retorno_saida_cidade,
                retorno_chegada_cidade=roteiro.retorno_chegada_cidade,
            )
            roteiro.nome = nome_gerado

        if is_ajax:
            return JsonResponse(
                {
                    "success": True,
                    "roteiro_id": roteiro.pk,
                    "nome": roteiro.nome,
                    "redirect_url": reverse("roteiro_lista"),
                }
            )

        messages.success(request, f'Roteiro "{roteiro.nome}" salvo com sucesso!')
        return redirect("roteiro_lista")
    except Exception as exc:
        if is_ajax:
            return JsonResponse({"success": False, "erro": str(exc)}, status=400)
        messages.error(request, f"Erro ao salvar roteiro: {exc}")
        target = "roteiro_editar" if roteiro and roteiro.pk else "roteiro_novo"
        if target == "roteiro_editar":
            return redirect(target, pk=roteiro.pk)
        return redirect(target)


def _coerce_json_value(value, expected_type, fallback):
    if isinstance(value, expected_type):
        return value
    if value in (None, ""):
        return fallback
    if isinstance(value, str):
        try:
            parsed = json.loads(value)
        except (TypeError, ValueError, json.JSONDecodeError):
            return fallback
        if isinstance(parsed, expected_type):
            return parsed
    return fallback


def _get_roteiro_request_data(request):
    try:
        raw_body = request.body.decode("utf-8") if request.body else "{}"
        data = json.loads(raw_body or "{}")
        if isinstance(data, dict):
            return data
    except (UnicodeDecodeError, json.JSONDecodeError, ValueError):
        pass

    data = request.POST.dict()
    data["destinos"] = _coerce_json_value(request.POST.get("destinos"), list, [])
    data["trechos"] = _coerce_json_value(request.POST.get("trechos"), list, [])
    data["retorno"] = _coerce_json_value(request.POST.get("retorno"), dict, {})
    if not data["retorno"]:
        data["retorno"] = {
            "saida_data": request.POST.get("retorno_saida_data", ""),
            "saida_hora": request.POST.get("retorno_saida_hora", ""),
            "chegada_data": request.POST.get("retorno_chegada_data", ""),
            "chegada_hora": request.POST.get("retorno_chegada_hora", ""),
        }
    return data


def _parse_tempo_viagem(value):
    if value in (None, ""):
        return None
    if isinstance(value, dt_time):
        return (value.hour * 60) + value.minute
    raw_value = str(value).strip()
    if ":" in raw_value:
        parsed_time = parse_time(raw_value)
        if parsed_time is None:
            return None
        return (parsed_time.hour * 60) + parsed_time.minute
    try:
        parsed = int(raw_value)
    except (TypeError, ValueError):
        return None
    return parsed if parsed >= 0 else None


def _parse_tempo_viagem_time(value):
    minutos = _parse_tempo_viagem(value)
    if minutos is None or minutos >= (24 * 60):
        return None
    horas, minutos_restantes = divmod(minutos, 60)
    return dt_time(hour=horas, minute=minutos_restantes)


def _calculate_arrival_from_duration(
    saida_data,
    saida_hora,
    chegada_data,
    chegada_hora,
    tempo_viagem_minutos,
):
    if (
        saida_data
        and saida_hora
        and tempo_viagem_minutos is not None
        and (not chegada_data or not chegada_hora)
    ):
        chegada_dt = datetime.combine(saida_data, saida_hora) + timedelta(
            minutes=tempo_viagem_minutos
        )
        return chegada_dt.date(), chegada_dt.time().replace(second=0, microsecond=0)
    return chegada_data, chegada_hora


def _retorno_duration_minutes(saida_data, saida_hora, chegada_data, chegada_hora):
    if not (saida_data and saida_hora and chegada_data and chegada_hora):
        return ""
    inicio = datetime.combine(saida_data, saida_hora)
    fim = datetime.combine(chegada_data, chegada_hora)
    delta = fim - inicio
    if delta.total_seconds() < 0:
        return ""
    return int(delta.total_seconds() // 60)


def _get_retorno_source_trecho(roteiro: Roteiro):
    return roteiro.trechos.order_by("ordem", "id").last()


def _serialize_trecho_editor(trecho: TrechoRoteiro):
    return {
        "ordem": trecho.ordem,
        "origem_estado": trecho.origem_estado_sigla or "PR",
        "origem_cidade": trecho.origem_cidade_nome or "",
        "destino_estado": trecho.destino_estado_sigla or "PR",
        "destino_cidade": trecho.destino_cidade_nome or "",
        "saida_data": trecho.saida_data.strftime("%Y-%m-%d") if trecho.saida_data else "",
        "saida_hora": trecho.saida_hora.strftime("%H:%M") if trecho.saida_hora else "",
        "chegada_data": trecho.chegada_data.strftime("%Y-%m-%d") if trecho.chegada_data else "",
        "chegada_hora": trecho.chegada_hora.strftime("%H:%M") if trecho.chegada_hora else "",
        "tempo_viagem_minutos": ""
        if trecho.tempo_viagem_minutos is None
        else trecho.tempo_viagem_minutos,
    }


def _serialize_retorno_editor(roteiro: Roteiro):
    retorno_trecho = _get_retorno_source_trecho(roteiro)
    saida_data = (
        retorno_trecho.retorno_saida_data
        if retorno_trecho and retorno_trecho.retorno_saida_data
        else roteiro.retorno_saida_data
    )
    saida_hora = (
        retorno_trecho.retorno_saida_hora
        if retorno_trecho and retorno_trecho.retorno_saida_hora
        else roteiro.retorno_saida_hora
    )
    chegada_data = (
        retorno_trecho.retorno_chegada_data
        if retorno_trecho and retorno_trecho.retorno_chegada_data
        else roteiro.retorno_chegada_data
    )
    chegada_hora = (
        retorno_trecho.retorno_chegada_hora
        if retorno_trecho and retorno_trecho.retorno_chegada_hora
        else roteiro.retorno_chegada_hora
    )
    return {
        "saida_data": saida_data.strftime("%Y-%m-%d")
        if saida_data
        else "",
        "saida_hora": saida_hora.strftime("%H:%M")
        if saida_hora
        else "",
        "chegada_data": chegada_data.strftime("%Y-%m-%d")
        if chegada_data
        else "",
        "chegada_hora": chegada_hora.strftime("%H:%M")
        if chegada_hora
        else "",
        "tempo_viagem": _serialize_duration_time(_resolve_roteiro_tempo_viagem(roteiro)),
    }


def _save_roteiro_payload(data, roteiro=None):
    if not isinstance(data, dict):
        raise ValueError("Payload invalido.")

    sede_uf = str(data.get("sede_uf") or "PR").strip().upper() or "PR"
    sede_cidade_nome = str(data.get("sede_cidade") or "").strip()
    destinos_raw = _coerce_json_value(data.get("destinos"), list, [])
    trechos_raw = _coerce_json_value(data.get("trechos"), list, [])
    retorno_raw = _coerce_json_value(data.get("retorno"), dict, {})
    tempo_viagem_source = data.get("tempo_viagem")
    if tempo_viagem_source in (None, ""):
        for trecho_data in trechos_raw:
            if not isinstance(trecho_data, dict):
                continue
            tempo_viagem_source = trecho_data.get("tempo_viagem")
            if tempo_viagem_source in (None, ""):
                tempo_viagem_source = trecho_data.get("tempo_viagem_minutos")
            if tempo_viagem_source not in (None, ""):
                break
        else:
            tempo_viagem_source = (
                retorno_raw.get("tempo_viagem")
                or retorno_raw.get("tempo_viagem_minutos")
            )

    tempo_viagem = _parse_tempo_viagem_time(tempo_viagem_source)
    tempo_viagem_minutos = _parse_tempo_viagem(tempo_viagem)

    estado_sede = _resolve_estado(sede_uf)
    cidade_sede = _resolve_cidade_by_name(sede_cidade_nome, estado_sede) or _resolve_cidade_payload(
        sede_cidade_nome,
        sede_uf,
    )
    if not estado_sede or not cidade_sede:
        raise ValueError("Informe uma cidade sede valida.")

    if not isinstance(destinos_raw, list) or not destinos_raw:
        raise ValueError("Informe ao menos um destino.")
    if not isinstance(trechos_raw, list):
        trechos_raw = []

    total_trechos = max(len(destinos_raw), len(trechos_raw))
    if total_trechos == 0:
        raise ValueError("Informe ao menos um destino.")

    trechos_resolvidos = []
    origem_estado_default = estado_sede
    origem_cidade_default = cidade_sede
    for index in range(total_trechos):
        trecho_data = trechos_raw[index] if index < len(trechos_raw) else {}
        destino_data = destinos_raw[index] if index < len(destinos_raw) else {}
        if not isinstance(trecho_data, dict):
            trecho_data = {}
        if not isinstance(destino_data, dict):
            destino_data = {}

        origem_estado = _resolve_estado(trecho_data.get("origem_estado")) or origem_estado_default
        origem_cidade_valor = trecho_data.get("origem_cidade") or origem_cidade_default.nome
        origem_cidade = _resolve_cidade_by_name(origem_cidade_valor, origem_estado) or _resolve_cidade_payload(
            origem_cidade_valor,
            origem_estado.sigla if origem_estado else "",
        )
        if not origem_estado or not origem_cidade:
            raise ValueError(f"Origem invalida para o trecho {index + 1}.")

        destino_estado = _resolve_estado(
            trecho_data.get("destino_estado") or destino_data.get("uf")
        ) or origem_estado
        destino_cidade_valor = (
            trecho_data.get("destino_cidade")
            or destino_data.get("cidade")
            or destino_data.get("cidade_id")
            or destino_data.get("cidade_nome")
        )
        destino_cidade = _resolve_cidade_by_name(destino_cidade_valor, destino_estado) or _resolve_cidade_payload(
            destino_cidade_valor,
            destino_estado.sigla if destino_estado else "",
        )
        if not destino_estado or not destino_cidade:
            raise ValueError(f"Destino invalido para o trecho {index + 1}.")

        saida_data = parse_date(trecho_data.get("saida_data") or "")
        saida_hora = parse_time(trecho_data.get("saida_hora") or "")
        chegada_data = parse_date(trecho_data.get("chegada_data") or "")
        chegada_hora = parse_time(trecho_data.get("chegada_hora") or "")
        chegada_data, chegada_hora = _calculate_arrival_from_duration(
            saida_data,
            saida_hora,
            chegada_data,
            chegada_hora,
            tempo_viagem_minutos,
        )

        trechos_resolvidos.append(
            {
                "ordem": index + 1,
                "origem_estado": origem_estado,
                "origem_cidade": origem_cidade,
                "destino_estado": destino_estado,
                "destino_cidade": destino_cidade,
                "saida_data": saida_data,
                "saida_hora": saida_hora,
                "chegada_data": chegada_data,
                "chegada_hora": chegada_hora,
                "tempo_viagem_minutos": tempo_viagem_minutos,
            }
        )
        origem_estado_default = destino_estado
        origem_cidade_default = destino_cidade

    ultimo_trecho = trechos_resolvidos[-1]
    retorno_saida_data = parse_date(
        retorno_raw.get("saida_data") or data.get("retorno_saida_data") or ""
    )
    retorno_saida_hora = parse_time(
        retorno_raw.get("saida_hora") or data.get("retorno_saida_hora") or ""
    )
    retorno_chegada_data = parse_date(
        retorno_raw.get("chegada_data") or data.get("retorno_chegada_data") or ""
    )
    retorno_chegada_hora = parse_time(
        retorno_raw.get("chegada_hora") or data.get("retorno_chegada_hora") or ""
    )
    retorno_chegada_data, retorno_chegada_hora = _calculate_arrival_from_duration(
        retorno_saida_data,
        retorno_saida_hora,
        retorno_chegada_data,
        retorno_chegada_hora,
        tempo_viagem_minutos,
    )

    with transaction.atomic():
        if roteiro is None:
            roteiro = Roteiro()

        roteiro.estado_sede = estado_sede
        roteiro.cidade_sede = cidade_sede
        roteiro.ativo = True
        roteiro.uf_origem = estado_sede.sigla
        roteiro.cidade_origem = cidade_sede.nome
        roteiro.uf_destino = ultimo_trecho["destino_estado"].sigla
        roteiro.cidade_destino = ultimo_trecho["destino_cidade"].nome
        roteiro.retorno_saida_cidade = ultimo_trecho["destino_cidade"].nome
        roteiro.retorno_saida_data = retorno_saida_data
        roteiro.retorno_saida_hora = retorno_saida_hora
        roteiro.retorno_chegada_cidade = cidade_sede.nome
        roteiro.retorno_chegada_data = retorno_chegada_data
        roteiro.retorno_chegada_hora = retorno_chegada_hora
        roteiro.tempo_viagem = tempo_viagem
        if not roteiro.tipo_deslocamento:
            roteiro.tipo_deslocamento = Roteiro.TipoDeslocamentoChoices.INTERIOR
        if not roteiro.nome:
            roteiro.nome = "(gerando...)"
        roteiro.save()

        roteiro.trechos.all().delete()
        ultimo_trecho_obj = None
        for trecho in trechos_resolvidos:
            ultimo_trecho_obj = TrechoRoteiro.objects.create(
                roteiro=roteiro,
                ordem=trecho["ordem"],
                origem_estado=trecho["origem_estado"],
                origem_cidade=trecho["origem_cidade"],
                destino_estado=trecho["destino_estado"],
                destino_cidade=trecho["destino_cidade"],
                saida_data=trecho["saida_data"],
                saida_hora=trecho["saida_hora"],
                chegada_data=trecho["chegada_data"],
                chegada_hora=trecho["chegada_hora"],
                tempo_viagem_minutos=tempo_viagem_minutos,
            )

        if ultimo_trecho_obj:
            ultimo_trecho_obj.retorno_saida_data = retorno_saida_data
            ultimo_trecho_obj.retorno_saida_hora = retorno_saida_hora
            ultimo_trecho_obj.retorno_chegada_data = retorno_chegada_data
            ultimo_trecho_obj.retorno_chegada_hora = retorno_chegada_hora
            ultimo_trecho_obj.save(
                update_fields=[
                    "retorno_saida_data",
                    "retorno_saida_hora",
                    "retorno_chegada_data",
                    "retorno_chegada_hora",
                ]
            )

        roteiro.nome = roteiro.gerar_nome()
        roteiro.save()

    return roteiro


@login_required
@require_http_methods(["GET"])
def roteiro_lista(request):
    q = (request.GET.get("q") or "").strip()
    roteiros = Roteiro.objects.filter(ativo=True).prefetch_related("trechos")
    if q:
        roteiros = roteiros.filter(
            Q(nome__icontains=q)
            | Q(cidade_origem__icontains=q)
            | Q(cidade_destino__icontains=q)
            | Q(trechos__cidade_destino__icontains=q)
        ).distinct()
    roteiros = roteiros.order_by("-criado_em", "-id")
    paginator = Paginator(roteiros, 12)
    page_obj = paginator.get_page(request.GET.get("page"))
    return render(
        request,
        "viagens/roteiro_lista.html",
        {
            "roteiros": page_obj,
            "page_obj": page_obj,
            "q": q,
        },
    )


@login_required
@require_http_methods(["GET", "POST"])
def roteiro_novo(request):
    if request.method == "POST":
        return _save_roteiro_from_post(request, roteiro=None)

    estados = Estado.objects.order_by("sigla")
    estado_padrao = Estado.objects.filter(sigla="PR").first() or estados.first()
    cidades_sede = (
        Cidade.objects.filter(estado=estado_padrao).order_by("nome")
        if estado_padrao
        else Cidade.objects.none()
    )
    cidade_padrao = _get_sede_cidade_default()
    return render(
        request,
        "viagens/roteiro_novo.html",
        {
            "estados": estados,
            "estado_sede_default": estado_padrao,
            "cidade_sede_default": cidade_padrao,
            "cidades_sede": cidades_sede,
        },
    )


@login_required
@require_http_methods(["GET", "POST"])
def roteiro_editar(request, pk: int):
    roteiro = get_object_or_404(Roteiro, pk=pk)
    estados = Estado.objects.order_by("sigla")

    if request.method == "GET":
        estado_atual = roteiro.estado_sede or Estado.objects.filter(sigla="PR").first() or estados.first()
        cidades_sede = (
            Cidade.objects.filter(estado=estado_atual).order_by("nome")
            if estado_atual
            else Cidade.objects.none()
        )
        trechos_json = [
            _serialize_trecho_editor(trecho)
            for trecho in roteiro.trechos.order_by("ordem", "id")
        ]
        return render(
            request,
            "viagens/roteiro_editar.html",
            {
                "roteiro": roteiro,
                "estados": estados,
                "estado_sede_default": estado_atual,
                "cidades_sede": cidades_sede,
                "trechos_json": json.dumps(trechos_json),
                "retorno_json": json.dumps(_serialize_retorno_editor(roteiro)),
                "tempo_viagem": _serialize_duration_time(_resolve_roteiro_tempo_viagem(roteiro)),
                "sede_uf": roteiro.estado_sede.sigla if roteiro.estado_sede else "PR",
                "sede_cidade": roteiro.cidade_sede.nome if roteiro.cidade_sede else "",
            },
        )

    if "application/json" not in (request.content_type or "") and request.POST:
        return _save_roteiro_from_post(request, roteiro=roteiro)

    try:
        roteiro = _save_roteiro_payload(_get_roteiro_request_data(request), roteiro=roteiro)
    except ValueError as exc:
        return JsonResponse({"ok": False, "erro": str(exc), "error": str(exc)}, status=400)
    except Exception as exc:  # pragma: no cover - defesa
        return JsonResponse(
            {"ok": False, "erro": f"Erro ao atualizar roteiro: {exc}", "error": f"Erro ao atualizar roteiro: {exc}"},
            status=500,
        )

    return JsonResponse(
        {
            "ok": True,
            "roteiro_id": roteiro.pk,
            "roteiro_nome": roteiro.nome,
            "redirect_url": reverse("roteiro_detalhe", args=[roteiro.pk]),
        }
    )


@login_required
@require_http_methods(["POST"])
def roteiro_excluir(request, pk: int):
    roteiro = get_object_or_404(Roteiro, pk=pk)
    roteiro.ativo = False
    roteiro.save(update_fields=["ativo", "atualizado_em"])
    is_ajax = request.headers.get("X-Requested-With") == "XMLHttpRequest"
    if is_ajax:
        return JsonResponse({"ok": True})
    messages.success(request, "Roteiro removido.")
    return redirect("roteiro_lista")


@login_required
@require_http_methods(["GET"])
def roteiro_detalhe(request, pk: int):
    roteiro = get_object_or_404(
        Roteiro.objects.prefetch_related("trechos").select_related("estado_sede", "cidade_sede"),
        pk=pk,
        ativo=True,
    )
    return render(
        request,
        "viagens/roteiro_detalhe.html",
        {"roteiro": roteiro},
    )


@login_required
@require_GET
def api_roteiros_buscar(request):
    """Busca roteiros ativos para autocomplete."""
    query = (request.GET.get("q") or "").strip()
    try:
        limit = min(max(int(request.GET.get("limit", 10)), 1), 50)
    except (TypeError, ValueError):
        limit = 10
    if len(query) < 2:
        return JsonResponse({"roteiros": []})

    roteiros = (
        Roteiro.objects.filter(ativo=True)
        .filter(
            Q(nome__icontains=query)
            | Q(cidade_origem__icontains=query)
            | Q(cidade_destino__icontains=query)
            | Q(uf_origem__icontains=query)
            | Q(uf_destino__icontains=query)
            | Q(trechos__cidade_destino__icontains=query)
        )
        .distinct()
        .order_by("-criado_em", "-id")[:limit]
    )

    results = [_serialize_roteiro_search_item(roteiro) for roteiro in roteiros]
    return JsonResponse({"roteiros": results})


def roteiros_buscar_api(request):
    return api_roteiros_buscar(request)


@login_required
@require_GET
def api_roteiro_json(request, pk: int):
    """Retorna um roteiro ativo com seus trechos."""
    roteiro = (
        Roteiro.objects.prefetch_related("trechos")
        .filter(pk=pk, ativo=True)
        .first()
    )
    if not roteiro:
        return JsonResponse({"erro": "Roteiro nao encontrado."}, status=404)
    retorno_payload = _serialize_retorno_editor(roteiro)
    trechos = [
        {
            "ordem": trecho.ordem,
            "uf_origem": trecho.uf_origem,
            "cidade_origem": trecho.cidade_origem,
            "uf_destino": trecho.uf_destino,
            "cidade_destino": trecho.cidade_destino,
            "cidade_destino_id": trecho.cidade_destino_obj.id
            if trecho.cidade_destino_obj
            else None,
            "cidade_destino_nome": trecho.cidade_destino,
            "modal": trecho.modal,
            "distancia_km": str(trecho.distancia_km) if trecho.distancia_km is not None else None,
            "saida_data": trecho.saida_data.isoformat() if trecho.saida_data else "",
            "saida_hora": trecho.saida_hora.strftime("%H:%M") if trecho.saida_hora else "",
            "chegada_data": trecho.chegada_data.isoformat() if trecho.chegada_data else "",
            "chegada_hora": trecho.chegada_hora.strftime("%H:%M") if trecho.chegada_hora else "",
            "tempo_viagem_minutos": ""
            if trecho.tempo_viagem_minutos is None
            else trecho.tempo_viagem_minutos,
        }
        for trecho in roteiro.trechos.all()
    ]

    data = {
        "id": roteiro.pk,
        "nome": roteiro.nome,
        "descricao": roteiro.descricao,
        "uf_origem": roteiro.uf_origem,
        "cidade_origem": roteiro.cidade_origem,
        "uf_destino": roteiro.uf_destino,
        "cidade_destino": roteiro.cidade_destino,
        "uf_sede_id": roteiro.uf_sede_id,
        "cidade_sede_id": roteiro.cidade_sede_id,
        "cidade_sede_nome": roteiro.cidade_sede_nome,
        "uf_sede_sigla": roteiro.uf_origem,
        "destinos": roteiro.get_destinos_display(),
        "distancia_km": str(roteiro.distancia_km) if roteiro.distancia_km is not None else None,
        "tipo_deslocamento": roteiro.tipo_deslocamento,
        "tipo_deslocamento_display": roteiro.get_tipo_deslocamento_display(),
        "tempo_viagem": _serialize_duration_time(_resolve_roteiro_tempo_viagem(roteiro)),
        "retorno": retorno_payload,
        "trechos": trechos,
    }
    return JsonResponse(data)


def api_roteiro_detalhe_json(request, pk: int):
    return api_roteiro_json(request, pk)


@login_required
@require_POST
def api_roteiro_criar_inline(request):
    """Cria um roteiro a partir dos dados atuais da etapa 3 do oficio."""
    try:
        data = json.loads(request.body.decode("utf-8") or "{}")
    except (UnicodeDecodeError, json.JSONDecodeError, ValueError):
        return JsonResponse({"erro": "JSON invalido."}, status=400)

    nome = (data.get("nome") or "").strip()
    if not nome:
        return JsonResponse({"erro": "Nome do roteiro e obrigatorio."}, status=400)

    estado_sede = _resolve_estado(data.get("uf_sede_id"))
    cidade_sede = _resolve_cidade(data.get("cidade_sede_id"))
    if not estado_sede or not cidade_sede:
        return JsonResponse(
            {"erro": "UF e Cidade da sede sao obrigatorios."},
            status=400,
        )

    trechos_data = data.get("trechos") or []
    if not isinstance(trechos_data, list) or not trechos_data:
        return JsonResponse(
            {"erro": "Adicione ao menos um destino antes de salvar o roteiro."},
            status=400,
        )

    normalized_tipo = _normalize_tipo_deslocamento(data.get("tipo_destino"))

    destinos_resolvidos = []
    for index, trecho_data in enumerate(trechos_data, start=1):
        cidade_destino = _resolve_cidade(trecho_data.get("cidade_destino_id"))
        if not cidade_destino:
            return JsonResponse(
                {"erro": f"Cidade de destino invalida para o trecho {index}."},
                status=400,
            )
        uf_destino = (
            (trecho_data.get("uf_destino") or cidade_destino.estado.sigla or "").strip().upper()
        )
        destinos_resolvidos.append(
            {
                "ordem": index,
                "cidade": cidade_destino,
                "uf": uf_destino,
                "modal": trecho_data.get("modal")
                or TrechoRoteiro.ModalChoices.VEICULO_PROPRIO,
                "distancia_km": trecho_data.get("distancia_km") or None,
                "observacao": (trecho_data.get("observacao") or "").strip(),
            }
        )

    ultimo_destino = destinos_resolvidos[-1]
    with transaction.atomic():
        roteiro = Roteiro.objects.create(
            nome=nome,
            descricao=(data.get("descricao") or "").strip(),
            estado_sede=estado_sede,
            cidade_sede=cidade_sede,
            uf_origem=estado_sede.sigla,
            cidade_origem=cidade_sede.nome,
            uf_destino=ultimo_destino["uf"],
            cidade_destino=ultimo_destino["cidade"].nome,
            retorno_saida_cidade=ultimo_destino["cidade"].nome,
            retorno_chegada_cidade=cidade_sede.nome,
            tipo_deslocamento=normalized_tipo,
            distancia_km=None,
        )

        origem_estado = estado_sede.sigla
        origem_cidade = cidade_sede.nome
        origem_estado_obj = estado_sede
        origem_cidade_obj = cidade_sede
        for destino in destinos_resolvidos:
            TrechoRoteiro.objects.create(
                roteiro=roteiro,
                ordem=destino["ordem"],
                origem_estado=origem_estado_obj,
                origem_cidade=origem_cidade_obj,
                uf_origem=origem_estado,
                cidade_origem=origem_cidade,
                destino_estado=destino["cidade"].estado,
                destino_cidade=destino["cidade"],
                uf_destino=destino["uf"],
                cidade_destino=destino["cidade"].nome,
                distancia_km=destino["distancia_km"],
                modal=destino["modal"],
                observacao=destino["observacao"],
            )
            origem_estado = destino["uf"]
            origem_cidade = destino["cidade"].nome
            origem_estado_obj = destino["cidade"].estado
            origem_cidade_obj = destino["cidade"]

    return JsonResponse(
        {
            "sucesso": True,
            "roteiro_id": roteiro.pk,
            "nome": roteiro.nome,
            "mensagem": f'Roteiro "{roteiro.nome}" salvo na biblioteca com sucesso.',
        },
        status=201,
    )


@login_required
@require_GET
def roteiros_cards_api(request, pk: int):
    roteiro = Roteiro.objects.filter(pk=pk, ativo=True).first()
    if not roteiro:
        return JsonResponse({"error": "Roteiro nao encontrado."}, status=404)
    return JsonResponse(_serialize_roteiro_cards_payload(roteiro))


@login_required
@require_http_methods(["POST"])
def api_roteiro_salvar(request):
    try:
        roteiro = _save_roteiro_payload(_get_roteiro_request_data(request))
    except ValueError as exc:
        return JsonResponse(
            {"ok": False, "erro": str(exc), "error": str(exc), "sucesso": False},
            status=400,
        )
    except Exception as exc:  # pragma: no cover - defesa de ultima linha
        message = f"Erro ao salvar roteiro: {exc}"
        return JsonResponse(
            {"ok": False, "erro": message, "error": message, "sucesso": False},
            status=500,
        )

    payload = _serialize_roteiro_cards_payload(roteiro)
    payload.update(
        {
            "ok": True,
            "sucesso": True,
            "id": roteiro.pk,
            "roteiro_id": roteiro.pk,
            "nome": roteiro.nome,
            "roteiro_nome": roteiro.nome,
            "message": f'Roteiro "{roteiro.nome}" salvo com sucesso!',
            "mensagem": f'Roteiro "{roteiro.nome}" salvo com sucesso!',
            "redirect_url": reverse("roteiro_detalhe", args=[roteiro.pk]),
        }
    )
    return JsonResponse(payload, status=201)


def roteiros_salvar_api(request):
    return api_roteiro_salvar(request)


def salvar_roteiro(request):
    return api_roteiro_salvar(request)


@login_required
@require_GET
def roteiros_listar_api(request):
    query = (request.GET.get("q") or "").strip()
    roteiros_qs = Roteiro.objects.filter(ativo=True)
    if query:
        roteiros_qs = roteiros_qs.filter(
            Q(nome__icontains=query)
            | Q(cidade_origem__icontains=query)
            | Q(cidade_destino__icontains=query)
            | Q(trechos__cidade_destino__icontains=query)
        ).distinct()

    roteiros = []
    for roteiro in roteiros_qs.order_by("-criado_em", "-id"):
        cidades = [
            trecho.cidade_destino
            for trecho in roteiro.cidades_destino
            if trecho.cidade_destino
        ]
        roteiros.append(
            {
                "id": roteiro.pk,
                "nome": roteiro.nome,
                "uf_sede": roteiro.uf_origem,
                "cidade_sede": roteiro.cidade_origem,
                "cidades_destino": " -> ".join(cidades),
                "total_cidades": roteiro.total_cidades,
                "criado_em": roteiro.criado_em.strftime("%d/%m/%Y %H:%M"),
                "cards": roteiro.get_cards_gerados(),
            }
        )
    return JsonResponse(roteiros, safe=False)


@require_POST
def oficio_vincular_roteiro(request, oficio_id: int):
    """Vincula um roteiro existente a um oficio."""
    oficio = get_object_or_404(Oficio, pk=oficio_id)
    payload = _request_payload(request)
    if payload is None:
        return JsonResponse(
            {"success": False, "message": "Corpo da requisicao invalido."},
            status=400,
        )

    form = OficioVincularRoteiroForm(payload)
    if not form.is_valid():
        error_message = "Dados invalidos para vincular o roteiro."
        if form.errors:
            error_message = " ".join(form.errors.get("__all__", [])) or " ".join(
                str(errors[0]) for errors in form.errors.values()
            )
        return JsonResponse({"success": False, "message": error_message}, status=400)

    roteiro = form.cleaned_data["roteiro"]
    observacao = form.cleaned_data.get("observacao") or ""

    if OficioRoteiro.objects.filter(oficio=oficio, roteiro=roteiro).exists():
        return JsonResponse(
            {
                "success": False,
                "message": "Este roteiro ja esta vinculado a este oficio.",
            },
            status=400,
        )

    vinculo = OficioRoteiro.objects.create(
        oficio=oficio,
        roteiro=roteiro,
        observacao=observacao,
    )
    return JsonResponse(
        {
            "success": True,
            "message": f"Roteiro '{roteiro.nome}' vinculado com sucesso.",
            "vinculo_id": vinculo.id,
            "roteiro": {
                "id": roteiro.id,
                "nome": roteiro.nome,
                "origem": f"{roteiro.cidade_origem}/{roteiro.uf_origem}",
                "destino": f"{roteiro.cidade_destino}/{roteiro.uf_destino}",
            },
        }
    )


@require_POST
def oficio_desvincular_roteiro(request, oficio_id: int, roteiro_id: int):
    """Remove o vinculo entre um oficio e um roteiro."""
    oficio_roteiro = get_object_or_404(
        OficioRoteiro,
        oficio_id=oficio_id,
        roteiro_id=roteiro_id,
    )
    roteiro_nome = oficio_roteiro.roteiro.nome
    oficio_roteiro.delete()
    return JsonResponse(
        {
            "success": True,
            "message": f"Roteiro '{roteiro_nome}' desvinculado com sucesso.",
        }
    )


def roteiros_lista(request):
    return roteiro_lista(request)


def roteiro_create(request):
    return roteiro_novo(request)


def roteiro_edit(request, pk: int):
    return roteiro_editar(request, pk)


def roteiro_delete(request, pk: int):
    return roteiro_excluir(request, pk)


def api_roteiro_cards(request, pk: int):
    return roteiros_cards_api(request, pk)


def roteiro_salvar(request):
    return api_roteiro_salvar(request)


__all__ = [
    "api_roteiro_cards",
    "api_roteiro_detalhe_json",
    "api_roteiro_criar_inline",
    "api_roteiro_salvar",
    "api_roteiro_json",
    "api_roteiros_buscar",
    "oficio_desvincular_roteiro",
    "oficio_vincular_roteiro",
    "roteiros_cards_api",
    "roteiros_buscar_api",
    "roteiros_listar_api",
    "roteiros_salvar_api",
    "salvar_roteiro",
    "roteiro_create",
    "roteiro_delete",
    "roteiro_detalhe",
    "roteiro_edit",
    "roteiro_lista",
    "roteiro_editar",
    "roteiro_excluir",
    "roteiro_novo",
    "roteiro_salvar",
    "roteiros_lista",
]
