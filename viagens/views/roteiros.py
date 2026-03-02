from __future__ import annotations

import json

from django.contrib.auth.decorators import login_required
from django.contrib import messages
from django.core.paginator import Paginator
from django.db import transaction
from django.db.models import Q
from django.http import JsonResponse
from django.shortcuts import get_object_or_404, redirect, render
from django.views.decorators.http import require_GET, require_POST

from ..forms import OficioVincularRoteiroForm, RoteiroForm, TrechoRoteiroFormSet
from ..models import Cidade, Estado, Oficio, OficioRoteiro, Roteiro, TrechoRoteiro


FORMSET_PREFIX = "trechos_roteiro"


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


def _resolve_estado(estado_id):
    if not estado_id:
        return None
    try:
        return Estado.objects.filter(pk=int(estado_id)).first()
    except (TypeError, ValueError):
        return None


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
        "cards": roteiro.get_cards_gerados(),
    }


def roteiro_lista(request):
    """Lista roteiros ativos com busca e paginacao."""
    roteiros_qs = Roteiro.objects.filter(ativo=True).order_by("-criado_em", "-id")
    query = (request.GET.get("q") or "").strip()
    if query:
        roteiros_qs = roteiros_qs.filter(
            Q(nome__icontains=query)
            | Q(cidade_origem__icontains=query)
            | Q(cidade_destino__icontains=query)
            | Q(uf_origem__icontains=query)
            | Q(uf_destino__icontains=query)
        )

    paginator = Paginator(roteiros_qs, 20)
    page_obj = paginator.get_page(request.GET.get("page"))
    context = {
        "page_obj": page_obj,
        "query": query,
        "total_roteiros": paginator.count,
    }
    return render(request, "viagens/roteiros/lista.html", context)


def roteiro_create(request):
    """Cria um roteiro reutilizavel e seus trechos."""
    roteiro = Roteiro()
    if request.method == "POST":
        form = RoteiroForm(request.POST, instance=roteiro)
        formset = TrechoRoteiroFormSet(
            request.POST,
            instance=roteiro,
            prefix=FORMSET_PREFIX,
        )
        if form.is_valid() and formset.is_valid():
            with transaction.atomic():
                roteiro = form.save()
                formset.instance = roteiro
                formset.save()
            messages.success(request, f"Roteiro '{roteiro.nome}' criado com sucesso.")
            return redirect("roteiro_detalhe", pk=roteiro.pk)
        messages.error(request, "Houve um erro ao criar o roteiro. Revise os campos.")
    else:
        form = RoteiroForm(instance=roteiro)
        formset = TrechoRoteiroFormSet(instance=roteiro, prefix=FORMSET_PREFIX)

    context = {
        "form": form,
        "formset": formset,
        "formset_prefix": FORMSET_PREFIX,
        "is_edit": False,
    }
    return render(request, "viagens/roteiros/form.html", context)


def roteiro_edit(request, pk: int):
    """Edita um roteiro e seus trechos."""
    roteiro = get_object_or_404(Roteiro, pk=pk)
    if request.method == "POST":
        form = RoteiroForm(request.POST, instance=roteiro)
        formset = TrechoRoteiroFormSet(
            request.POST,
            instance=roteiro,
            prefix=FORMSET_PREFIX,
        )
        if form.is_valid() and formset.is_valid():
            with transaction.atomic():
                roteiro = form.save()
                formset.save()
            messages.success(request, f"Roteiro '{roteiro.nome}' atualizado com sucesso.")
            return redirect("roteiro_detalhe", pk=roteiro.pk)
        messages.error(request, "Houve um erro ao atualizar o roteiro. Revise os campos.")
    else:
        form = RoteiroForm(instance=roteiro)
        formset = TrechoRoteiroFormSet(instance=roteiro, prefix=FORMSET_PREFIX)

    context = {
        "form": form,
        "formset": formset,
        "formset_prefix": FORMSET_PREFIX,
        "is_edit": True,
        "roteiro": roteiro,
    }
    return render(request, "viagens/roteiros/form.html", context)


@require_POST
def roteiro_delete(request, pk: int):
    """Desativa um roteiro. Roteiros vinculados nao podem ser desativados."""
    roteiro = get_object_or_404(Roteiro, pk=pk)
    if roteiro.oficios_vinculados.exists():
        messages.error(
            request,
            "Nao e possivel excluir o roteiro pois ele esta vinculado a um ou mais oficios.",
        )
        return redirect("roteiro_detalhe", pk=roteiro.pk)

    roteiro.ativo = False
    roteiro.save(update_fields=["ativo", "atualizado_em"])
    messages.success(request, f"Roteiro '{roteiro.nome}' desativado com sucesso.")
    return redirect("roteiro_lista")


def roteiro_detalhe(request, pk: int):
    """Exibe um roteiro e seus vinculos com oficios."""
    roteiro = get_object_or_404(
        Roteiro.objects.prefetch_related("trechos", "oficios_vinculados__oficio"),
        pk=pk,
    )
    oficios_vinculados = roteiro.oficios_vinculados.select_related("oficio").order_by(
        "-vinculado_em",
        "-id",
    )
    context = {
        "roteiro": roteiro,
        "oficios_vinculados": oficios_vinculados,
    }
    return render(request, "viagens/roteiros/detalhe.html", context)


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
            uf_origem=estado_sede.sigla,
            cidade_origem=cidade_sede.nome,
            uf_destino=ultimo_destino["uf"],
            cidade_destino=ultimo_destino["cidade"].nome,
            tipo_deslocamento=normalized_tipo,
            distancia_km=None,
        )

        origem_estado = estado_sede.sigla
        origem_cidade = cidade_sede.nome
        for destino in destinos_resolvidos:
            TrechoRoteiro.objects.create(
                roteiro=roteiro,
                ordem=destino["ordem"],
                uf_origem=origem_estado,
                cidade_origem=origem_cidade,
                uf_destino=destino["uf"],
                cidade_destino=destino["cidade"].nome,
                distancia_km=destino["distancia_km"],
                modal=destino["modal"],
                observacao=destino["observacao"],
            )
            origem_estado = destino["uf"]
            origem_cidade = destino["cidade"].nome

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
@require_POST
def roteiros_salvar_api(request):
    try:
        data = json.loads(request.body.decode("utf-8") or "{}")
    except (UnicodeDecodeError, json.JSONDecodeError, ValueError):
        return JsonResponse({"error": "JSON invalido."}, status=400)

    nome = (data.get("nome") or "").strip()
    uf_sede = (data.get("uf_sede") or "PR").strip().upper()
    cidade_sede = (data.get("cidade_sede") or "Curitiba").strip()
    destinos_data = data.get("destinos") or []

    if not nome:
        return JsonResponse({"error": "O nome do roteiro e obrigatorio."}, status=400)
    if not isinstance(destinos_data, list) or not destinos_data:
        return JsonResponse(
            {"error": "O roteiro deve ter pelo menos um destino."},
            status=400,
        )

    destinos_normalizados = []
    for index, destino in enumerate(destinos_data, start=1):
        uf_destino = (destino.get("uf") or "PR").strip().upper()
        cidade_destino = (
            destino.get("cidade")
            or destino.get("cidade_destino")
            or destino.get("cidade_nome")
            or ""
        ).strip()
        if not cidade_destino:
            return JsonResponse(
                {"error": f"Cidade invalida para o destino {index}."},
                status=400,
            )
        destinos_normalizados.append(
            {
                "uf": uf_destino or "PR",
                "cidade": cidade_destino,
            }
        )

    ultimo_destino = destinos_normalizados[-1]
    with transaction.atomic():
        roteiro = Roteiro.objects.create(
            nome=nome,
            descricao=(data.get("descricao") or "").strip(),
            uf_origem=uf_sede or "PR",
            cidade_origem=cidade_sede or "Curitiba",
            uf_destino=ultimo_destino["uf"],
            cidade_destino=ultimo_destino["cidade"],
            distancia_km=None,
            tipo_deslocamento=_normalize_tipo_deslocamento(
                data.get("tipo_destino") or Roteiro.TipoDeslocamentoChoices.INTERIOR
            ),
            ativo=True,
        )

        origem_uf = roteiro.uf_origem
        origem_cidade = roteiro.cidade_origem
        for ordem, destino in enumerate(destinos_normalizados, start=1):
            TrechoRoteiro.objects.create(
                roteiro=roteiro,
                ordem=ordem,
                uf_origem=origem_uf,
                cidade_origem=origem_cidade,
                uf_destino=destino["uf"],
                cidade_destino=destino["cidade"],
                modal=TrechoRoteiro.ModalChoices.VEICULO_PROPRIO,
            )
            origem_uf = destino["uf"]
            origem_cidade = destino["cidade"]

    payload = _serialize_roteiro_cards_payload(roteiro)
    return JsonResponse(payload, status=201)


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


__all__ = [
    "api_roteiro_detalhe_json",
    "api_roteiro_criar_inline",
    "api_roteiro_json",
    "api_roteiros_buscar",
    "oficio_desvincular_roteiro",
    "oficio_vincular_roteiro",
    "roteiros_cards_api",
    "roteiros_buscar_api",
    "roteiros_listar_api",
    "roteiros_salvar_api",
    "roteiro_create",
    "roteiro_delete",
    "roteiro_detalhe",
    "roteiro_edit",
    "roteiro_lista",
]
