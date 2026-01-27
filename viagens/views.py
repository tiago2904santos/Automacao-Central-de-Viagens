from __future__ import annotations

from datetime import date, datetime, time, timedelta
from typing import Iterable

from django.conf import settings
from django.forms import inlineformset_factory
from django.http import JsonResponse
from django.shortcuts import get_object_or_404, redirect, render
from django.views.decorators.http import require_GET, require_http_methods

from django.db import models
from django.db.models import Count, Q
from django.db.models.functions import TruncDate
from django.utils import timezone
from .forms import TrechoForm
from .models import Cidade, Estado, Oficio, Trecho, Viajante, Veiculo


def _normalizar_placa(placa: str) -> str:
    return placa.replace(" ", "").replace("-", "").upper()


def _viajantes_payload(viajantes: Iterable[Viajante]) -> list[dict]:
    return [
        {
            "id": viajante.id, # type: ignore
            "nome": viajante.nome,
            "rg": viajante.rg,
            "cpf": viajante.cpf,
            "cargo": viajante.cargo,
            "telefone": viajante.telefone,
        }
        for viajante in viajantes
    ]


def _format_date(date_value: str | date | None) -> str:
    if not date_value:
        return ""
    if isinstance(date_value, date):
        return date_value.strftime("%d/%m/%Y")
    try:
        return datetime.strptime(date_value, "%Y-%m-%d").strftime("%d/%m/%Y")
    except ValueError:
        return str(date_value)


def _format_time(time_value: str | time | None) -> str:
    if not time_value:
        return ""
    if isinstance(time_value, time):
        return time_value.strftime("%H:%M")
    return str(time_value).strip()


def _format_date_time(
    date_value: str | date | None, time_value: str | time | None
) -> str:
    date_part = _format_date(date_value)
    time_part = _format_time(time_value)
    if date_part and time_part:
        return f"{date_part} - {time_part}"
    return date_part or time_part


def _format_trecho_local(cidade: Cidade | None, estado: Estado | None) -> str:
    if cidade and estado:
        return f"{cidade.nome}/{estado.sigla}"
    if cidade:
        return cidade.nome
    if estado:
        return estado.sigla
    return ""


DEFAULT_CARGO_CHOICES = [
    "Agente de Policia Judiciaria",
    "Assessora",
    "Delegado de Policia",
    "Papiloscopista",
    "Motorista",
]

DEFAULT_COMBUSTIVEL_CHOICES = [
    "Gasolina",
    "Diesel",
    "Flex",
    "Etanol",
    "GNV",
    "Eletrico",
    "Hibrido",
]


def _get_cargo_choices() -> list[str]:
    custom = getattr(settings, "CARGO_CHOICES", None)
    if custom:
        return list(custom)
    cargos = (
        Viajante.objects.exclude(cargo="")
        .order_by("cargo")
        .values_list("cargo", flat=True)
        .distinct()
    )
    return list(cargos) or DEFAULT_CARGO_CHOICES


def _get_combustivel_choices() -> list[str]:
    custom = getattr(settings, "COMBUSTIVEL_CHOICES", None)
    if custom:
        return list(custom)
    combustiveis = (
        Veiculo.objects.exclude(combustivel="")
        .order_by("combustivel")
        .values_list("combustivel", flat=True)
        .distinct()
    )
    return list(combustiveis) or DEFAULT_COMBUSTIVEL_CHOICES


def _prune_trailing_trechos_post(data, prefix: str):
    try:
        total = int(data.get(f"{prefix}-TOTAL_FORMS", 0))
    except (TypeError, ValueError):
        return data
    if total <= 1:
        return data

    mutable = data.copy()

    def has_user_data(index: int) -> bool:
        for field in (
            "destino_estado",
            "destino_cidade",
            "saida_data",
            "saida_hora",
            "chegada_data",
            "chegada_hora",
        ):
            value = mutable.get(f"{prefix}-{index}-{field}", "")
            if str(value).strip():
                return True
        return False

    while total > 1 and not has_user_data(total - 1):
        total -= 1

    mutable[f"{prefix}-TOTAL_FORMS"] = str(total)
    return mutable


def _parse_period(request) -> int:
    raw = (request.GET.get("periodo") or "").strip()
    try:
        periodo = int(raw)
    except ValueError:
        periodo = 30
    if periodo not in {7, 30, 90}:
        periodo = 30
    return periodo


def _build_daily_series(queryset, dt_field: str, days: int) -> list[dict]:
    hoje = timezone.localdate()
    inicio = hoje - timedelta(days=days - 1)
    agregados = (
        queryset.filter(**{f"{dt_field}__date__gte": inicio})
        .annotate(dia=TruncDate(dt_field))
        .values("dia")
        .annotate(total=Count("id"))
        .order_by("dia")
    )
    mapa = {item["dia"]: item["total"] for item in agregados}
    serie = []
    for offset in range(days):
        dia = inicio + timedelta(days=offset)
        serie.append({"dia": dia.strftime("%d/%m"), "total": mapa.get(dia, 0)})
    return serie


def _dashboard_payload(periodo: int) -> dict:
    hoje = timezone.localdate()
    inicio = hoje - timedelta(days=periodo - 1)

    oficios_qs = Oficio.objects.all()
    veiculos_qs = Veiculo.objects.all()
    viajantes_qs = Viajante.objects.all()
    trechos_qs = Trecho.objects.select_related("oficio")

    oficios_total = oficios_qs.count()
    oficios_periodo = oficios_qs.filter(created_at__date__gte=inicio).count()

    trechos_periodo = trechos_qs.filter(oficio__created_at__date__gte=inicio).count()

    serie_oficios = _build_daily_series(oficios_qs, "created_at", periodo)
    serie_trechos = _build_daily_series(
        trechos_qs.filter(oficio__created_at__date__gte=inicio), "oficio__created_at", periodo
    )

    return {
        "periodo": periodo,
        "kpis": {
            "oficios": {
                "total": oficios_total,
                "periodo": oficios_periodo,
                "rotulo_periodo": f"ultimos {periodo} dias",
            },
            "veiculos": {
                "total": veiculos_qs.count(),
                "periodo": veiculos_qs.count(),
                "rotulo_periodo": "cadastros",
            },
            "viajantes": {
                "total": viajantes_qs.count(),
                "periodo": viajantes_qs.count(),
                "rotulo_periodo": "cadastros",
            },
            "trechos": {
                "total": trechos_qs.count(),
                "periodo": trechos_periodo,
                "rotulo_periodo": f"trechos em {periodo} dias",
            },
        },
        "series": {
            "oficios": serie_oficios,
            "trechos": serie_trechos,
        },
        "recentes": list(
            oficios_qs.order_by("-created_at").values(
                "id", "oficio", "protocolo", "destino", "created_at"
            )[:5]
        ),
    }


def _get_wizard_data(request) -> dict:
    return request.session.get("oficio_wizard", {})


def _update_wizard_data(request, new_data: dict) -> dict:
    data = _get_wizard_data(request)
    data.update(new_data)
    request.session["oficio_wizard"] = data
    request.session.modified = True
    return data


def _clear_wizard_data(request) -> None:
    request.session.pop("oficio_wizard", None)
    request.session.modified = True


def _document_context(oficio: Oficio, viajantes: Iterable[Viajante]) -> dict:
    trechos = list(
        oficio.trechos.select_related(
            "origem_estado",
            "origem_cidade",
            "destino_estado",
            "destino_cidade",
        ).order_by("ordem")
    )
    trechos_payload = [
        {
            "ordem": trecho.ordem,
            "origem": _format_trecho_local(trecho.origem_cidade, trecho.origem_estado),
            "destino": _format_trecho_local(
                trecho.destino_cidade, trecho.destino_estado
            ),
            "saida": _format_date_time(trecho.saida_data, trecho.saida_hora),
            "chegada": _format_date_time(trecho.chegada_data, trecho.chegada_hora),
        }
        for trecho in trechos
    ]

    return {
        "oficio": oficio.oficio,
        "protocolo": oficio.protocolo,
        "data": oficio.data,
        "destino": oficio.destino,
        "assunto": oficio.assunto,
        "viajantes": viajantes,
        "trechos": trechos_payload,
        "roteiro_ida_saida_local": oficio.roteiro_ida_saida_local,
        "roteiro_ida_saida_datahora": oficio.roteiro_ida_saida_datahora,
        "roteiro_ida_chegada_local": oficio.roteiro_ida_chegada_local,
        "roteiro_ida_chegada_datahora": oficio.roteiro_ida_chegada_datahora,
        "roteiro_volta_saida_local": oficio.roteiro_volta_saida_local,
        "roteiro_volta_saida_datahora": oficio.roteiro_volta_saida_datahora,
        "roteiro_volta_chegada_local": oficio.roteiro_volta_chegada_local,
        "roteiro_volta_chegada_datahora": oficio.roteiro_volta_chegada_datahora,
        "placa": oficio.placa,
        "modelo": oficio.modelo,
        "combustivel": oficio.combustivel,
        "motorista": oficio.motorista,
        "motivo": oficio.motivo,
        "oficio_id": oficio.id,  # type: ignore
    }


def _serialize_recentes(recentes: list[dict]) -> list[dict]:
    payload = []
    for item in recentes:
        criado = item.get("created_at")
        if criado:
            criado_fmt = timezone.localtime(criado).strftime("%d/%m/%Y %H:%M")
        else:
            criado_fmt = ""
        payload.append(
            {
                "id": item.get("id"),
                "oficio": item.get("oficio"),
                "protocolo": item.get("protocolo"),
                "destino": item.get("destino"),
                "created_at": criado_fmt,
            }
        )
    return payload


@require_GET
def dashboard_home(request):
    periodo = _parse_period(request)
    payload = _dashboard_payload(periodo)
    payload["recentes"] = _serialize_recentes(payload["recentes"])
    payload["initial_payload"] = {
        "periodo": payload["periodo"],
        "kpis": payload["kpis"],
        "series": payload["series"],
        "recentes": payload["recentes"],
    }
    return render(request, "viagens/dashboard.html", payload)


@require_GET
def dashboard_data_api(request):
    periodo = _parse_period(request)
    payload = _dashboard_payload(periodo)
    payload["recentes"] = _serialize_recentes(payload["recentes"])
    return JsonResponse(payload)


@require_http_methods(["GET", "POST"])
def formulario(request):
    data = _get_wizard_data(request)
    viajantes = Viajante.objects.order_by("nome")
    erro = ""

    if request.method == "POST":
        goto_step = (request.POST.get("goto_step") or "").strip()
        oficio_val = request.POST.get("oficio", "").strip()
        protocolo_val = request.POST.get("protocolo", "").strip()
        data_val = request.POST.get("data", "").strip()
        assunto_val = request.POST.get("assunto", "").strip()
        viajantes_ids = request.POST.getlist("viajantes_ids")

        if not oficio_val or not protocolo_val or not viajantes_ids:
            erro = "Preencha oficio, protocolo e selecione ao menos um viajante."
        else:
            _update_wizard_data(
                request,
                {
                    "oficio": oficio_val,
                    "protocolo": protocolo_val,
                    "data": data_val,
                    "assunto": assunto_val,
                    "viajantes_ids": viajantes_ids,
                },
            )
            if goto_step == "1":
                return redirect("formulario")
            return redirect("oficio_step2")

        data = {
            "oficio": oficio_val,
            "protocolo": protocolo_val,
            "data": data_val,
            "assunto": assunto_val,
            "viajantes_ids": viajantes_ids,
        }

    selected_ids = [str(item) for item in data.get("viajantes_ids", [])]
    return render(
        request,
        "viagens/form.html",
        {
            "viajantes": viajantes,
            "erro": erro,
            "oficio": data.get("oficio", ""),
            "protocolo": data.get("protocolo", ""),
            "data": data.get("data", ""),
            "assunto": data.get("assunto", ""),
            "selected_ids": selected_ids,
        },
    )


@require_http_methods(["GET", "POST"])
def oficio_step2(request):
    data = _get_wizard_data(request)
    if not data.get("oficio"):
        return redirect("formulario")

    erro = ""
    viajantes_ids = data.get("viajantes_ids", [])
    viajantes_sel = list(Viajante.objects.filter(id__in=viajantes_ids).order_by("nome"))
    preview_viajantes = _viajantes_payload(viajantes_sel)

    if request.method == "POST":
        goto_step = (request.POST.get("goto_step") or "").strip()
        placa_val = request.POST.get("placa", "").strip()
        modelo_val = request.POST.get("modelo", "").strip()
        combustivel_val = request.POST.get("combustivel", "").strip()
        motorista_id = request.POST.get("motorista_id", "").strip()
        motorista_nome = request.POST.get("motorista_nome", "").strip()

        placa_norm = _normalizar_placa(placa_val) if placa_val else ""
        if placa_norm and (not modelo_val or not combustivel_val):
            veiculo = Veiculo.objects.filter(placa__iexact=placa_norm).first()
            if veiculo:
                modelo_val = modelo_val or veiculo.modelo
                combustivel_val = combustivel_val or veiculo.combustivel

        if not placa_val or not modelo_val or not combustivel_val:
            erro = "Preencha placa, modelo e combustivel."
        else:
            _update_wizard_data(
                request,
                {
                    "placa": placa_norm or placa_val,
                    "modelo": modelo_val,
                    "combustivel": combustivel_val,
                    "motorista_id": motorista_id,
                    "motorista_nome": motorista_nome,
                },
            )
            if goto_step == "1":
                return redirect("formulario")
            if goto_step == "2":
                return redirect("oficio_step2")
            return redirect("oficio_step3")

        data = {
            **data,
            "placa": placa_val,
            "modelo": modelo_val,
            "combustivel": combustivel_val,
            "motorista_id": motorista_id,
            "motorista_nome": motorista_nome,
        }
        viajantes_ids = data.get("viajantes_ids", [])
        viajantes_sel = list(
            Viajante.objects.filter(id__in=viajantes_ids).order_by("nome")
        )
        preview_viajantes = _viajantes_payload(viajantes_sel)

    viajantes = Viajante.objects.order_by("nome")
    motorista_preview = None
    motorista_id_val = data.get("motorista_id") or ""
    if motorista_id_val and str(motorista_id_val).isdigit():
        motorista_obj = Viajante.objects.filter(id=motorista_id_val).first()
        if motorista_obj:
            motorista_preview = _servidor_payload(motorista_obj)

    return render(
        request,
        "viagens/oficio_step2.html",
        {
            "viajantes": viajantes,
            "erro": erro,
            "placa": data.get("placa", ""),
            "modelo": data.get("modelo", ""),
            "combustivel": data.get("combustivel", ""),
            "motorista_id": data.get("motorista_id", ""),
            "motorista_nome": data.get("motorista_nome", ""),
            "oficio": data.get("oficio", ""),
            "protocolo": data.get("protocolo", ""),
            "data_oficio": data.get("data", ""),
            "assunto": data.get("assunto", ""),
            "preview_viajantes": preview_viajantes,
            "motorista_preview": motorista_preview,
        },
    )


@require_http_methods(["GET", "POST"])
def oficio_step3(request):
    data = _get_wizard_data(request)
    if not data.get("oficio"):
        return redirect("formulario")
    if not data.get("placa"):
        return redirect("oficio_step2")

    erro = ""
    motivo_val = ""

    TrechoFormSet = inlineformset_factory(
        Oficio,
        Trecho,
        form=TrechoForm,
        extra=1,
        can_delete=False,
    )

    if request.method == "POST":
        motivo_val = request.POST.get("motivo", "").strip()
        post_data = _prune_trailing_trechos_post(request.POST, "trechos")
        formset = TrechoFormSet(post_data, instance=Oficio(), prefix="trechos")

        if formset.is_valid():
            forms_validas = [form for form in formset.forms if form.cleaned_data]
            if not forms_validas:
                erro = "Adicione ao menos um trecho para o roteiro."
            else:
                wizard = _get_wizard_data(request)
                viajantes_ids = wizard.get("viajantes_ids", [])
                viajantes = list(
                    Viajante.objects.filter(id__in=viajantes_ids).order_by("nome")
                )

                placa = wizard.get("placa", "").strip()
                placa_norm = _normalizar_placa(placa) if placa else ""
                veiculo = None
                if placa_norm:
                    veiculo = Veiculo.objects.filter(placa__iexact=placa_norm).first()

                modelo = wizard.get("modelo", "").strip()
                combustivel = wizard.get("combustivel", "").strip()
                if veiculo:
                    modelo = modelo or veiculo.modelo
                    combustivel = combustivel or veiculo.combustivel

                motorista_id = wizard.get("motorista_id") or ""
                motorista_nome = wizard.get("motorista_nome", "").strip()
                motorista_obj = None
                if motorista_id and motorista_id.isdigit():
                    motorista_obj = Viajante.objects.filter(id=motorista_id).first()
                    if motorista_obj:
                        motorista_nome = motorista_obj.nome

                primeiro = forms_validas[0].cleaned_data
                ultimo = forms_validas[-1].cleaned_data
                sede_estado = primeiro.get("origem_estado")
                sede_cidade = primeiro.get("origem_cidade")
                destino_estado = ultimo.get("destino_estado")
                destino_cidade = ultimo.get("destino_cidade")

                destino_texto = ""
                if destino_cidade and destino_estado:
                    destino_texto = f"{destino_cidade.nome} / {destino_estado.sigla}"

                oficio_obj = Oficio.objects.create(
                    oficio=wizard.get("oficio", ""),
                    protocolo=wizard.get("protocolo", ""),
                    data=wizard.get("data", ""),
                    destino=destino_texto or wizard.get("destino", ""),
                    assunto=wizard.get("assunto", ""),
                    estado_sede=sede_estado,
                    cidade_sede=sede_cidade,
                    estado_destino=destino_estado,
                    cidade_destino=destino_cidade,
                    placa=placa_norm or placa,
                    modelo=modelo,
                    combustivel=combustivel,
                    motorista=motorista_nome,
                    motorista_viajante=motorista_obj,
                    motivo=motivo_val,
                    veiculo=veiculo,
                )

                trechos_salvar = []
                for idx, form in enumerate(forms_validas):
                    trecho = form.save(commit=False)
                    trecho.oficio = oficio_obj
                    trecho.ordem = idx + 1
                    if idx == 0:
                        trecho.origem_estado = sede_estado
                        trecho.origem_cidade = sede_cidade
                    else:
                        trecho.origem_estado = trechos_salvar[idx - 1].destino_estado
                        trecho.origem_cidade = trechos_salvar[idx - 1].destino_cidade
                    trechos_salvar.append(trecho)

                Trecho.objects.bulk_create(trechos_salvar)

                if viajantes:
                    oficio_obj.viajantes.set(viajantes)

                _clear_wizard_data(request)

                context = _document_context(oficio_obj, viajantes)
                return render(request, "viagens/document.html", context)
        else:
            erro = "Revise os campos obrigatorios do roteiro."
    else:
        formset = TrechoFormSet(prefix="trechos", instance=Oficio())

    return render(
        request,
        "viagens/oficio_step3.html",
        {
            "erro": erro,
            "formset": formset,
            "motivo": motivo_val,
        },
    )


@require_GET
def documento_oficio(request, oficio_id: int):
    oficio = get_object_or_404(Oficio, id=oficio_id)
    viajantes = oficio.viajantes.order_by("nome")
    context = _document_context(oficio, viajantes)
    return render(request, "viagens/document.html", context)


@require_http_methods(["GET", "POST"])
def viajante_cadastro(request):
    if request.method == "POST":
        nome = request.POST.get("nome", "").strip()
        rg = request.POST.get("rg", "").strip()
        cpf = request.POST.get("cpf", "").strip()
        cargo = request.POST.get("cargo", "").strip()
        telefone = request.POST.get("telefone", "").strip()
        if nome and rg and cpf and cargo:
            Viajante.objects.create(
                nome=nome,
                rg=rg,
                cpf=cpf,
                cargo=cargo,
                telefone=telefone,
            )
            return redirect("viajantes_lista")
        return render(
            request,
            "viagens/viajante_form.html",
            {
                "erro": "Preencha nome, RG, CPF e cargo.",
                "cargo_choices": _get_cargo_choices(),
            },
        )
    return render(
        request,
        "viagens/viajante_form.html",
        {"cargo_choices": _get_cargo_choices()},
    )


@require_http_methods(["GET"])
def viajantes_lista(request):
    q = request.GET.get("q", "").strip()
    cargo = request.GET.get("cargo", "").strip()

    viajantes = Viajante.objects.all()
    if q:
        viajantes = viajantes.filter(
            models.Q(nome__icontains=q)
            | models.Q(rg__icontains=q)
            | models.Q(cpf__icontains=q)
            | models.Q(cargo__icontains=q)
            | models.Q(telefone__icontains=q)
        )
    if cargo:
        viajantes = viajantes.filter(cargo__iexact=cargo)

    viajantes = viajantes.order_by("nome")
    cargo_choices = _get_cargo_choices()
    return render(
        request,
        "viagens/viajantes_list.html",
        {
            "viajantes": viajantes,
            "q": q,
            "cargo_choices": cargo_choices,
            "cargo_selecionado": cargo,
        },
    )


@require_http_methods(["GET", "POST"])
def veiculo_cadastro(request):
    if request.method == "POST":
        placa = request.POST.get("placa", "").strip()
        placa_norm = _normalizar_placa(placa) if placa else ""
        modelo = request.POST.get("modelo", "").strip()
        combustivel = request.POST.get("combustivel", "").strip()
        if placa_norm and modelo and combustivel:
            veiculo, created = Veiculo.objects.get_or_create(
                placa=placa_norm,
                defaults={"modelo": modelo, "combustivel": combustivel},
            )
            if not created:
                veiculo.modelo = modelo
                veiculo.combustivel = combustivel
                veiculo.save(update_fields=["modelo", "combustivel"])
            return redirect("veiculos_lista")
        return render(
            request,
            "viagens/veiculo_form.html",
            {
                "erro": "Preencha todos os campos.",
                "combustivel_choices": _get_combustivel_choices(),
            },
        )
    return render(
        request,
        "viagens/veiculo_form.html",
        {"combustivel_choices": _get_combustivel_choices()},
    )


@require_http_methods(["GET"])
def veiculos_lista(request):
    q = request.GET.get("q", "").strip()
    combustivel = request.GET.get("combustivel", "").strip()

    veiculos = Veiculo.objects.all()
    if q:
        veiculos = veiculos.filter(
            models.Q(placa__icontains=q) | models.Q(modelo__icontains=q)
        )
    if combustivel:
        veiculos = veiculos.filter(combustivel__iexact=combustivel)

    veiculos = veiculos.order_by("placa")
    combustivel_choices = _get_combustivel_choices()
    return render(
        request,
        "viagens/veiculos_list.html",
        {
            "veiculos": veiculos,
            "q": q,
            "combustivel_choices": combustivel_choices,
            "combustivel_selecionado": combustivel,
        },
    )


@require_http_methods(["GET"])
def oficios_lista(request):
    q = request.GET.get("q", "").strip()
    com_veiculo = request.GET.get("com_veiculo", "") == "1"

    oficios = Oficio.objects.select_related(
        "veiculo",
        "cidade_destino",
        "estado_destino",
        "cidade_sede",
        "estado_sede",
    ).prefetch_related("viajantes")
    if q:
        oficios = oficios.filter(
            models.Q(oficio__icontains=q)
            | models.Q(protocolo__icontains=q)
            | models.Q(destino__icontains=q)
            | models.Q(assunto__icontains=q)
            | models.Q(placa__icontains=q)
            | models.Q(motorista__icontains=q)
            | models.Q(viajantes__nome__icontains=q)
            | models.Q(cidade_destino__nome__icontains=q)
            | models.Q(cidade_sede__nome__icontains=q)
            | models.Q(estado_destino__sigla__icontains=q)
            | models.Q(estado_sede__sigla__icontains=q)
        ).distinct()
    if com_veiculo:
        oficios = oficios.filter(veiculo__isnull=False)

    oficios = oficios.order_by("-created_at")
    return render(
        request,
        "viagens/oficios_list.html",
        {"oficios": oficios, "q": q, "com_veiculo": com_veiculo},
    )


@require_GET
def viajantes_api(request):
    ids = request.GET.get("ids", "")
    if not ids:
        return JsonResponse({"viajantes": []})
    raw_ids = [item.strip() for item in ids.split(",") if item.strip().isdigit()]
    viajantes = Viajante.objects.filter(id__in=raw_ids).order_by("nome")
    return JsonResponse({"viajantes": _viajantes_payload(viajantes)})


@require_GET
def veiculo_api(request):
    plate = request.GET.get("plate", "").strip()
    plate_norm = _normalizar_placa(plate) if plate else ""
    if not plate_norm:
        return JsonResponse({"found": False})
    veiculo = Veiculo.objects.filter(placa__iexact=plate_norm).first()
    if not veiculo:
        return JsonResponse({"found": False, "plate": plate_norm})
    return JsonResponse(
        {
            "found": True,
            "id": veiculo.id,
            "plate": plate_norm,
            "modelo": veiculo.modelo,
            "combustivel": veiculo.combustivel,
            "label": f"{veiculo.placa} - {veiculo.modelo}",
        }
    )


@require_GET
def cidades_api(request):
    estado = (
        request.GET.get("estado", "") or request.GET.get("uf", "")
    ).strip().upper()
    if not estado:
        return JsonResponse({"cidades": []})
    cidades = Cidade.objects.filter(estado__sigla=estado).order_by("nome")
    payload = [{"id": cidade.id, "nome": cidade.nome} for cidade in cidades]
    return JsonResponse({"cidades": payload})


@require_GET
def ufs_api(request):
    q = request.GET.get("q", "").strip()
    estados = Estado.objects.all()
    if q:
        estados = estados.filter(Q(sigla__icontains=q) | Q(nome__icontains=q))
    estados = estados.order_by("nome")[:30]
    payload = [
        {
            "id": estado.sigla,
            "sigla": estado.sigla,
            "nome": estado.nome,
            "label": f"{estado.nome} ({estado.sigla})",
        }
        for estado in estados
    ]
    return JsonResponse({"results": payload})


@require_GET
def cidades_busca_api(request):
    uf = request.GET.get("uf", "").strip().upper()
    q = request.GET.get("q", "").strip()
    if not uf:
        return JsonResponse({"results": []})
    cidades = Cidade.objects.filter(estado__sigla=uf)
    if q:
        cidades = cidades.filter(nome__icontains=q)
    cidades = cidades.select_related("estado").order_by("nome")[:50]
    payload = [
        {
            "id": cidade.id,
            "nome": cidade.nome,
            "uf": cidade.estado.sigla if cidade.estado else uf,
            "label": f"{cidade.nome}/{cidade.estado.sigla}",
        }
        for cidade in cidades
    ]
    return JsonResponse({"results": payload})


def _servidor_payload(viajante: Viajante) -> dict:
    return {
        "id": viajante.id,
        "nome": viajante.nome,
        "rg": viajante.rg,
        "cpf": viajante.cpf,
        "cargo": viajante.cargo,
        "telefone": viajante.telefone,
        "label": viajante.nome,
    }


@require_GET
def servidores_api(request):
    q = request.GET.get("q", "").strip()
    viajantes = Viajante.objects.all()
    if q:
        viajantes = viajantes.filter(
            Q(nome__icontains=q)
            | Q(rg__icontains=q)
            | Q(cpf__icontains=q)
            | Q(cargo__icontains=q)
        )
    viajantes = viajantes.order_by("nome")[:40]
    return JsonResponse({"results": [_servidor_payload(v) for v in viajantes]})


@require_GET
def motoristas_api(request):
    return servidores_api(request)


@require_GET
def servidor_detail_api(request, servidor_id: int):
    viajante = get_object_or_404(Viajante, id=servidor_id)
    return JsonResponse(_servidor_payload(viajante))


@require_GET
def motorista_detail_api(request, motorista_id: int):
    viajante = get_object_or_404(Viajante, id=motorista_id)
    return JsonResponse(_servidor_payload(viajante))


@require_GET
def veiculo_detail_api(request, veiculo_id: int):
    veiculo = get_object_or_404(Veiculo, id=veiculo_id)
    return JsonResponse(
        {
            "id": veiculo.id,
            "placa": veiculo.placa,
            "modelo": veiculo.modelo,
            "combustivel": veiculo.combustivel,
            "label": f"{veiculo.placa} - {veiculo.modelo}",
        }
    )


@require_GET
def veiculos_busca_api(request):
    placa = request.GET.get("placa", "").strip()
    q = request.GET.get("q", "").strip()
    termo = placa or q
    veiculos = Veiculo.objects.all()
    if termo:
        termo_norm = _normalizar_placa(termo)
        veiculos = veiculos.filter(
            Q(placa__icontains=termo_norm) | Q(modelo__icontains=termo)
        )
    veiculos = veiculos.order_by("placa")[:40]
    payload = [
        {
            "id": veiculo.id,
            "placa": veiculo.placa,
            "modelo": veiculo.modelo,
            "combustivel": veiculo.combustivel,
            "label": f"{veiculo.placa} - {veiculo.modelo}",
        }
        for veiculo in veiculos
    ]
    return JsonResponse({"results": payload})


@require_http_methods(["GET", "POST"])
def modal_viajante_form(request):
    if request.method == "POST":
        nome = request.POST.get("nome", "").strip()
        rg = request.POST.get("rg", "").strip()
        cpf = request.POST.get("cpf", "").strip()
        cargo = request.POST.get("cargo", "").strip()
        telefone = request.POST.get("telefone", "").strip()

        erros = {}
        if not nome:
            erros["nome"] = "Informe o nome."
        if not rg:
            erros["rg"] = "Informe o RG."
        if not cpf:
            erros["cpf"] = "Informe o CPF."
        if not cargo:
            erros["cargo"] = "Informe o cargo."

        if not erros:
            viajante = Viajante.objects.create(
                nome=nome,
                rg=rg,
                cpf=cpf,
                cargo=cargo,
                telefone=telefone,
            )
            return JsonResponse({"success": True, "item": _servidor_payload(viajante)})

        return render(
            request,
            "viagens/partials/modal_viajante.html",
            {
                "erros": erros,
                "values": request.POST,
                "cargo_choices": _get_cargo_choices(),
            },
            status=400,
        )

    return render(
        request,
        "viagens/partials/modal_viajante.html",
        {"cargo_choices": _get_cargo_choices()},
    )


@require_http_methods(["GET", "POST"])
def modal_veiculo_form(request):
    if request.method == "POST":
        placa = request.POST.get("placa", "").strip()
        modelo = request.POST.get("modelo", "").strip()
        combustivel = request.POST.get("combustivel", "").strip()

        erros = {}
        if not placa:
            erros["placa"] = "Informe a placa."
        if not modelo:
            erros["modelo"] = "Informe o modelo."
        if not combustivel:
            erros["combustivel"] = "Informe o combustivel."

        placa_norm = _normalizar_placa(placa) if placa else ""
        if placa_norm and Veiculo.objects.filter(placa__iexact=placa_norm).exists():
            erros["placa"] = "Ja existe um veiculo com esta placa."

        if not erros:
            veiculo = Veiculo.objects.create(
                placa=placa_norm,
                modelo=modelo,
                combustivel=combustivel,
            )
            return JsonResponse(
                {
                    "success": True,
                    "item": {
                        "id": veiculo.id,
                        "placa": veiculo.placa,
                        "modelo": veiculo.modelo,
                        "combustivel": veiculo.combustivel,
                        "label": f"{veiculo.placa} - {veiculo.modelo}",
                    },
                }
            )

        return render(
            request,
            "viagens/partials/modal_veiculo.html",
            {
                "erros": erros,
                "values": request.POST,
                "combustivel_choices": _get_combustivel_choices(),
            },
            status=400,
        )

    return render(
        request,
        "viagens/partials/modal_veiculo.html",
        {"combustivel_choices": _get_combustivel_choices()},
    )
