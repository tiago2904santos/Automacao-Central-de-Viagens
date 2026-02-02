from __future__ import annotations

import logging
import re
from datetime import date, datetime, time, timedelta
from typing import Iterable

from django.conf import settings
from django.core.paginator import Paginator
from django.db import models, transaction
from django.db.models import Count, Q
from django.db.models.functions import TruncDate
from django.forms import inlineformset_factory
from django.forms.models import BaseInlineFormSet
from django.http import JsonResponse
from django.shortcuts import get_object_or_404, redirect, render
from django.urls import reverse
from django.utils import timezone
from django.utils.dateparse import parse_date, parse_time
from django.views.decorators.http import require_GET, require_http_methods
from .forms import MotoristaSelectForm, ServidoresSelectForm, TrechoForm
from .models import Cidade, Estado, Oficio, Trecho, Viajante, Veiculo
from .services.diarias import calcular_diarias, formatar_valor_diarias
from django.http import HttpResponse
from django.shortcuts import get_object_or_404
from django.views.decorators.http import require_GET
from .documents.document import build_oficio_docx_bytes
from .documents.document import build_oficio_docx_and_pdf


logger = logging.getLogger(__name__)


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


def _combine_date_time(date_value: date | None, time_value: time | None) -> datetime | None:
    if not date_value:
        return None
    return datetime.combine(date_value, time_value or time.min)


def _format_trecho_local(cidade: Cidade | None, estado: Estado | None) -> str:
    if cidade and estado:
        return f"{cidade.nome}/{estado.sigla}"
    if cidade:
        return cidade.nome
    if estado:
        return estado.sigla
    return ""


class OrderedTrechoInlineFormSet(BaseInlineFormSet):
    def get_queryset(self):
        return super().get_queryset().order_by("ordem", "id")

DEFAULT_CARGO_CHOICES = [
    "Agente de Policia Judiciaria",
    "Delegado",
    "Administrativo",
    "Assessor",
    "Papiloscopista",
]

DEFAULT_COMBUSTIVEL_CHOICES = [
    "Gasolina",
    "Etanol",
    "Diesel",
]


def _get_cargo_choices() -> list[str]:
    custom = getattr(settings, "CARGO_CHOICES", None)
    if custom:
        return list(custom)
    return list(DEFAULT_CARGO_CHOICES)


def _get_combustivel_choices() -> list[str]:
    custom = getattr(settings, "COMBUSTIVEL_CHOICES", None)
    if custom:
        return list(custom)
    return list(DEFAULT_COMBUSTIVEL_CHOICES)


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


TRECHO_FIELDS = (
    "origem_estado",
    "origem_cidade",
    "destino_estado",
    "destino_cidade",
    "saida_data",
    "saida_hora",
    "chegada_data",
    "chegada_hora",
)


def _serialize_trechos_from_post(post_data) -> list[dict[str, str | int]]:
    prefix = "trechos"

    def _as_int_or_str(raw_value: str | None) -> str | int:
        value = (raw_value or "").strip()
        if not value:
            return ""
        if value.isdigit():
            return int(value)
        return value

    total_raw = post_data.get(f"{prefix}-TOTAL_FORMS", 0) or 0
    try:
        total = int(total_raw)
    except (TypeError, ValueError):
        total = 0

    trechos: list[dict[str, str | int]] = []
    for index in range(total):
        oe_raw = post_data.get(f"{prefix}-{index}-origem_estado", "")
        oc_raw = post_data.get(f"{prefix}-{index}-origem_cidade", "")
        de_raw = post_data.get(f"{prefix}-{index}-destino_estado", "")
        dc_raw = post_data.get(f"{prefix}-{index}-destino_cidade", "")

        trecho = {
            "origem_estado": _as_int_or_str(oe_raw),
            "origem_cidade": _as_int_or_str(oc_raw),
            "destino_estado": _as_int_or_str(de_raw),
            "destino_cidade": _as_int_or_str(dc_raw),
            "saida_data": (post_data.get(f"{prefix}-{index}-saida_data") or "").strip(),
            "saida_hora": (post_data.get(f"{prefix}-{index}-saida_hora") or "").strip(),
            "chegada_data": (post_data.get(f"{prefix}-{index}-chegada_data") or "").strip(),
            "chegada_hora": (post_data.get(f"{prefix}-{index}-chegada_hora") or "").strip(),
        }

        if any(str(value).strip() for value in trecho.values()):
            trechos.append(trecho)

    return trechos or [{}]


def _normalize_trechos_initial(trechos_data) -> list[dict[str, str | int]]:
    if not trechos_data:
        return [{}]
    normalized: list[dict[str, str | int]] = []
    for trecho in trechos_data:
        entry = {field: trecho.get(field, "") for field in TRECHO_FIELDS}
        normalized.append(entry)
    return normalized or [{}]


def _build_trecho_formset(extra: int):
    return inlineformset_factory(
        Oficio,
        Trecho,
        form=TrechoForm,
        extra=extra,
        can_delete=False,
    )


def _get_retornodata(wizard_data: dict) -> dict[str, str]:
    retorno = wizard_data.get("retorno") or {}
    return {
        "retorno_saida_data": retorno.get("retorno_saida_data", ""),
        "retorno_saida_hora": retorno.get("retorno_saida_hora", ""),
        "retorno_chegada_data": retorno.get("retorno_chegada_data", ""),
        "retorno_chegada_hora": retorno.get("retorno_chegada_hora", ""),
    }


def _build_trechos_summary(trechos_data: list[dict]) -> list[dict[str, str]]:
    summary: list[dict[str, str]] = []
    for trecho in trechos_data:
        origem_estado = _resolve_estado(trecho.get("origem_estado"))
        origem_cidade = _resolve_cidade(trecho.get("origem_cidade"))
        destino_estado = _resolve_estado(trecho.get("destino_estado"))
        destino_cidade = _resolve_cidade(trecho.get("destino_cidade"))
        summary.append(
            {
                "origem": _format_trecho_local(origem_cidade, origem_estado),
                "destino": _format_trecho_local(destino_cidade, destino_estado),
                "saida": _format_date_time(
                    trecho.get("saida_data"), trecho.get("saida_hora")
                ),
                "chegada": _format_date_time(
                    trecho.get("chegada_data"), trecho.get("chegada_hora")
                ),
            }
        )
    return summary


def _wizard_has_step4_data(wizard_data: dict) -> bool:
    if not wizard_data:
        return False
    trechos = wizard_data.get("trechos") or []
    if not trechos:
        return False
    retorno_payload = _get_retornodata(wizard_data)
    if not retorno_payload["retorno_saida_data"] or not retorno_payload["retorno_chegada_data"]:
        return False
    if not wizard_data.get("viajantes_ids"):
        return False
    if not wizard_data.get("tipo_destino"):
        return False
    return True


def _build_step4_context(wizard_data: dict) -> dict:
    viajantes_ids = wizard_data.get("viajantes_ids", [])
    viajantes = list(Viajante.objects.filter(id__in=viajantes_ids).order_by("nome"))
    trechos_summary = _build_trechos_summary(wizard_data.get("trechos") or [])
    retorno_payload = _get_retornodata(wizard_data)

    motorista_id = wizard_data.get("motorista_id") or ""
    motorista_obj = (
        Viajante.objects.filter(id=motorista_id).first()
        if motorista_id and str(motorista_id).isdigit()
        else None
    )
    motorista_nome = (
        motorista_obj.nome if motorista_obj else wizard_data.get("motorista_nome", "")
    )
    motorista_preview = _servidor_payload(motorista_obj) if motorista_obj else None

    return {
        "oficio": wizard_data.get("oficio", ""),
        "protocolo": wizard_data.get("protocolo", ""),
        "assunto": wizard_data.get("assunto", ""),
        "destino": wizard_data.get("destino", ""),
        "placa": wizard_data.get("placa", ""),
        "modelo": wizard_data.get("modelo", ""),
        "combustivel": wizard_data.get("combustivel", ""),
        "motorista_nome": motorista_nome,
        "motorista_preview": motorista_preview,
        "motorista_carona": wizard_data.get("motorista_carona", False),
        "motorista_oficio": wizard_data.get("motorista_oficio", ""),
        "motorista_protocolo": wizard_data.get("motorista_protocolo", ""),
        "tipo_destino": wizard_data.get("tipo_destino", ""),
        "trechos_summary": trechos_summary,
        "trechos_count": len(trechos_summary),
        "retorno_saida_data": retorno_payload["retorno_saida_data"],
        "retorno_saida_hora": retorno_payload["retorno_saida_hora"],
        "retorno_chegada_data": retorno_payload["retorno_chegada_data"],
        "retorno_chegada_hora": retorno_payload["retorno_chegada_hora"],
        "retorno_saida_cidade": wizard_data.get("retorno_saida_cidade", ""),
        "retorno_chegada_cidade": wizard_data.get("retorno_chegada_cidade", ""),
        "quantidade_diarias": wizard_data.get("quantidade_diarias", ""),
        "valor_diarias": wizard_data.get("valor_diarias", ""),
        "valor_diarias_extenso": wizard_data.get("valor_diarias_extenso", ""),
        "motivo": wizard_data.get("motivo", ""),
        "preview_viajantes": _viajantes_payload(viajantes),
        "quantidade_servidores": len(viajantes),
        "voltar_url": reverse("oficio_step3"),
    }


def _finalize_oficio_from_wizard(wizard_data: dict) -> tuple[Oficio, list[Viajante]]:
    viajantes_ids = wizard_data.get("viajantes_ids", [])
    viajantes = list(Viajante.objects.filter(id__in=viajantes_ids).order_by("nome"))
    placa = wizard_data.get("placa", "").strip()
    placa_norm = _normalizar_placa(placa) if placa else ""
    veiculo = (
        Veiculo.objects.filter(placa__iexact=placa_norm).first() if placa_norm else None
    )
    modelo = wizard_data.get("modelo", "").strip()
    combustivel = wizard_data.get("combustivel", "").strip()
    if veiculo:
        modelo = modelo or veiculo.modelo
        combustivel = combustivel or veiculo.combustivel

    motorista_id = wizard_data.get("motorista_id") or ""
    motorista_nome = wizard_data.get("motorista_nome", "").strip()
    motorista_obj = None
    if motorista_id and str(motorista_id).isdigit():
        motorista_obj = Viajante.objects.filter(id=motorista_id).first()
        if motorista_obj:
            motorista_nome = motorista_obj.nome
    motorista_carona = False
    if motorista_id:
        motorista_carona = str(motorista_id) not in [str(item) for item in viajantes_ids]
    elif motorista_nome:
        motorista_carona = True

    tipo_destino = wizard_data.get("tipo_destino", "")
    retorno_payload = _get_retornodata(wizard_data)
    retorno_saida_data = (
        parse_date(retorno_payload["retorno_saida_data"])
        if retorno_payload["retorno_saida_data"]
        else None
    )
    retorno_saida_hora = (
        parse_time(retorno_payload["retorno_saida_hora"])
        if retorno_payload["retorno_saida_hora"]
        else None
    )
    retorno_chegada_data = (
        parse_date(retorno_payload["retorno_chegada_data"])
        if retorno_payload["retorno_chegada_data"]
        else None
    )
    retorno_chegada_hora = (
        parse_time(retorno_payload["retorno_chegada_hora"])
        if retorno_payload["retorno_chegada_hora"]
        else None
    )

    trechos_data = wizard_data.get("trechos") or []
    if not trechos_data:
        raise ValueError("Trechos obrigatorios ausentes.")
    primeiro = trechos_data[0]
    ultimo = trechos_data[-1]
    sede_estado = _resolve_estado(primeiro.get("origem_estado"))
    sede_cidade = _resolve_cidade(primeiro.get("origem_cidade"))
    destino_estado = _resolve_estado(ultimo.get("destino_estado"))
    destino_cidade = _resolve_cidade(ultimo.get("destino_cidade"))

    saida_sede_dt = _combine_date_time(
        parse_date(primeiro.get("saida_data")), parse_time(primeiro.get("saida_hora"))
    )
    retorno_chegada_dt = _combine_date_time(retorno_chegada_data, retorno_chegada_hora)

    resultado_diarias = calcular_diarias(
        tipo_destino=tipo_destino,
        saida_sede=saida_sede_dt,
        chegada_sede=retorno_chegada_dt,
        quantidade_servidores=len(viajantes),
    )

    quantidade_diarias = resultado_diarias.quantidade_diarias_str
    valor_diarias = formatar_valor_diarias(resultado_diarias.valor_total_oficio)

    destino_texto = ""
    if destino_cidade and destino_estado:
        destino_texto = f"{destino_cidade.nome} / {destino_estado.sigla}"

    retorno_saida_cidade = _format_trecho_local(destino_cidade, destino_estado)
    retorno_chegada_cidade = _format_trecho_local(sede_cidade, sede_estado)

    with transaction.atomic():
        oficio_obj = Oficio.objects.create(
            oficio=wizard_data.get("oficio", ""),
            protocolo=wizard_data.get("protocolo", ""),
            destino=destino_texto or wizard_data.get("destino", ""),
            assunto=wizard_data.get("assunto", ""),
            tipo_destino=tipo_destino,
            estado_sede=sede_estado,
            cidade_sede=sede_cidade,
            estado_destino=destino_estado,
            cidade_destino=destino_cidade,
            retorno_saida_cidade=retorno_saida_cidade,
            retorno_saida_data=retorno_saida_data,
            retorno_saida_hora=retorno_saida_hora,
            retorno_chegada_cidade=retorno_chegada_cidade,
            retorno_chegada_data=retorno_chegada_data,
            retorno_chegada_hora=retorno_chegada_hora,
            quantidade_diarias=quantidade_diarias,
            valor_diarias=valor_diarias,
            valor_diarias_extenso=wizard_data.get("valor_diarias_extenso", ""),
            placa=placa_norm or placa,
            modelo=modelo,
            combustivel=combustivel,
            motorista=motorista_nome,
            motorista_oficio=wizard_data.get("motorista_oficio", ""),
            motorista_protocolo=wizard_data.get("motorista_protocolo", ""),
            motorista_carona=motorista_carona,
            motorista_viajante=motorista_obj,
            motivo=wizard_data.get("motivo", ""),
            veiculo=veiculo,
        )

        trechos_instances = []
        for idx, trecho in enumerate(trechos_data):
            origem_estado_obj = _resolve_estado(trecho.get("origem_estado"))
            origem_cidade_obj = _resolve_cidade(trecho.get("origem_cidade"))
            destino_estado_obj = _resolve_estado(trecho.get("destino_estado"))
            destino_cidade_obj = _resolve_cidade(trecho.get("destino_cidade"))
            trecho_obj = Trecho(
                oficio=oficio_obj,
                ordem=idx + 1,
                origem_estado=origem_estado_obj,
                origem_cidade=origem_cidade_obj,
                destino_estado=destino_estado_obj,
                destino_cidade=destino_cidade_obj,
                saida_data=parse_date(trecho.get("saida_data"))
                if trecho.get("saida_data")
                else None,
                saida_hora=parse_time(trecho.get("saida_hora"))
                if trecho.get("saida_hora")
                else None,
                chegada_data=parse_date(trecho.get("chegada_data"))
                if trecho.get("chegada_data")
                else None,
                chegada_hora=parse_time(trecho.get("chegada_hora"))
                if trecho.get("chegada_hora")
                else None,
            )
            trechos_instances.append(trecho_obj)

        Trecho.objects.bulk_create(trechos_instances)
        if viajantes:
            oficio_obj.viajantes.set(viajantes)

    return oficio_obj, viajantes


def _resolve_estado(value: str | int | None) -> Estado | None:
    if value is None or value == "":
        return None
    if isinstance(value, Estado):
        return value
    if isinstance(value, int) or (isinstance(value, str) and value.isdigit()):
        try:
            estado_id = int(value)
        except (TypeError, ValueError):
            estado_id = None
        if estado_id is not None:
            estado = Estado.objects.filter(id=estado_id).first()
            if estado:
                return estado
    sigla = str(value).strip()
    return Estado.objects.filter(sigla__iexact=sigla).first()


def _resolve_cidade(value: str | int | None, estado: Estado | int | str | None = None) -> Cidade | None:
    if value is None or value == "":
        return None
    raw = str(value).strip()
    if not raw:
        return None
    if raw.isdigit():
        return Cidade.objects.filter(id=int(raw)).first()

    qs = Cidade.objects.all()
    if isinstance(estado, Estado):
        qs = qs.filter(estado=estado)
    elif isinstance(estado, int):
        qs = qs.filter(estado_id=estado)
    elif isinstance(estado, str):
        resolved_estado = _resolve_estado(estado)
        if resolved_estado:
            qs = qs.filter(estado=resolved_estado)

    cidade = qs.filter(nome__iexact=raw).first()
    if cidade:
        return cidade
    return qs.filter(nome__icontains=raw).first()


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


def _sync_trechos_origens_post(data, prefix: str):
    try:
        total = int(data.get(f"{prefix}-TOTAL_FORMS", 0))
    except (TypeError, ValueError):
        return data

    if total <= 1:
        return data

    mutable = data.copy()
    prev_estado = mutable.get(f"{prefix}-0-destino_estado", "").strip()
    prev_cidade = mutable.get(f"{prefix}-0-destino_cidade", "").strip()

    for index in range(1, total):
        origem_estado_key = f"{prefix}-{index}-origem_estado"
        origem_cidade_key = f"{prefix}-{index}-origem_cidade"
        if not str(mutable.get(origem_estado_key, "")).strip() and prev_estado:
            mutable[origem_estado_key] = prev_estado
        if not str(mutable.get(origem_cidade_key, "")).strip() and prev_cidade:
            mutable[origem_cidade_key] = prev_cidade

        destino_estado = mutable.get(f"{prefix}-{index}-destino_estado", "").strip()
        destino_cidade = mutable.get(f"{prefix}-{index}-destino_cidade", "").strip()
        if destino_estado and destino_cidade:
            prev_estado = destino_estado
            prev_cidade = destino_cidade

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
    if request.method == "GET" and not request.GET.get("resume"):
        _clear_wizard_data(request)
        data = {}
    viajantes = Viajante.objects.order_by("nome")
    erro = ""
    servidores_form = ServidoresSelectForm(
        initial={"servidores": data.get("viajantes_ids", [])}
    )

    if request.method == "POST":
        goto_step = (request.POST.get("goto_step") or "").strip()
        oficio_val = request.POST.get("oficio", "").strip()
        protocolo_val = request.POST.get("protocolo", "").strip()
        assunto_val = request.POST.get("assunto", "").strip()
        servidores_form = ServidoresSelectForm(request.POST)
        if servidores_form.is_valid():
            viajantes_ids = [
                str(item.id)
                for item in servidores_form.cleaned_data.get("servidores", [])
            ]
        else:
            viajantes_ids = []

        if not oficio_val or not protocolo_val or not viajantes_ids:
            erro = "Preencha oficio, protocolo e selecione ao menos um viajante."
        else:
            _update_wizard_data(
                request,
                {
                    "oficio": oficio_val,
                    "protocolo": protocolo_val,
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
            "assunto": assunto_val,
            "viajantes_ids": viajantes_ids,
        }

    selected_ids = [str(item) for item in data.get("viajantes_ids", [])]
    selected_viajantes = list(
        Viajante.objects.filter(id__in=selected_ids).order_by("nome")
    )
    return render(
        request,
        "viagens/form.html",
        {
            "viajantes": viajantes,
            "erro": erro,
            "oficio": data.get("oficio", ""),
            "protocolo": data.get("protocolo", ""),
            "assunto": data.get("assunto", ""),
            "data_criacao": "",
            "selected_ids": selected_ids,
            "selected_viajantes": selected_viajantes,
            "servidores_form": servidores_form,
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
    motorista_form = MotoristaSelectForm(
        initial={"motorista": data.get("motorista_id", "")}
    )

    if request.method == "POST":
        goto_step = (request.POST.get("goto_step") or "").strip()
        placa_val = request.POST.get("placa", "").strip()
        modelo_val = request.POST.get("modelo", "").strip()
        combustivel_val = request.POST.get("combustivel", "").strip()
        motorista_form = MotoristaSelectForm(request.POST)
        motorista_id = ""
        if motorista_form.is_valid():
            motorista_obj = motorista_form.cleaned_data.get("motorista")
            if motorista_obj:
                motorista_id = str(motorista_obj.id)
        motorista_nome = request.POST.get("motorista_nome", "").strip()
        motorista_oficio = request.POST.get("motorista_oficio", "").strip()
        motorista_protocolo = request.POST.get("motorista_protocolo", "").strip()

        placa_norm = _normalizar_placa(placa_val) if placa_val else ""
        if placa_norm and (not modelo_val or not combustivel_val):
            veiculo = Veiculo.objects.filter(placa__iexact=placa_norm).first()
            if veiculo:
                modelo_val = modelo_val or veiculo.modelo
                combustivel_val = combustivel_val or veiculo.combustivel

        motorista_carona = False
        if motorista_id:
            motorista_carona = motorista_id not in [str(item) for item in viajantes_ids]
        elif motorista_nome:
            motorista_carona = True

        if not placa_val or not modelo_val or not combustivel_val:
            erro = "Preencha placa, modelo e combustivel."
        elif motorista_carona and (not motorista_oficio or not motorista_protocolo):
            erro = "Informe oficio e protocolo do motorista (carona)."
        else:
            _update_wizard_data(
                request,
                {
                    "placa": placa_norm or placa_val,
                    "modelo": modelo_val,
                    "combustivel": combustivel_val,
                    "motorista_id": motorista_id,
                    "motorista_nome": motorista_nome,
                    "motorista_oficio": motorista_oficio,
                    "motorista_protocolo": motorista_protocolo,
                    "motorista_carona": motorista_carona,
                },
            )
            if goto_step == "1":
                return redirect(f"{reverse('formulario')}?resume=1")
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
            "motorista_oficio": motorista_oficio,
            "motorista_protocolo": motorista_protocolo,
            "motorista_carona": motorista_carona,
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
    motorista_nome_val = data.get("motorista_nome", "")
    motorista_carona = False
    if motorista_id_val:
        motorista_carona = str(motorista_id_val) not in [str(item) for item in viajantes_ids]
    elif motorista_nome_val:
        motorista_carona = True

    return render(
        request,
        "viagens/oficio_step2.html",
        {
            "viajantes": viajantes,
            "erro": erro,
            "placa": data.get("placa", ""),
            "modelo": data.get("modelo", ""),
            "combustivel": data.get("combustivel", ""),
            "combustivel_choices": _get_combustivel_choices(),
            "motorista_id": data.get("motorista_id", ""),
            "motorista_nome": motorista_nome_val,
            "motorista_oficio": data.get("motorista_oficio", ""),
            "motorista_protocolo": data.get("motorista_protocolo", ""),
            "motorista_carona": motorista_carona,
            "oficio": data.get("oficio", ""),
            "protocolo": data.get("protocolo", ""),
            "assunto": data.get("assunto", ""),
            "viajantes_ids": viajantes_ids,
            "preview_viajantes": preview_viajantes,
            "motorista_preview": motorista_preview,
            "motorista_form": motorista_form,
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
    motivo_val = data.get("motivo", "")
    tipo_destino = data.get("tipo_destino", "")
    valor_diarias_extenso = data.get("valor_diarias_extenso", "")
    retorno_payload = _get_retornodata(data)
    retorno_saida_data_raw = retorno_payload["retorno_saida_data"]
    retorno_saida_hora_raw = retorno_payload["retorno_saida_hora"]
    retorno_chegada_data_raw = retorno_payload["retorno_chegada_data"]
    retorno_chegada_hora_raw = retorno_payload["retorno_chegada_hora"]

    trechos_initial = _normalize_trechos_initial(data.get("trechos"))
    formset_extra = max(1, len(trechos_initial))
    TrechoFormSet = _build_trecho_formset(formset_extra)

    dummy_oficio = Oficio()
    if request.method == "POST":
        logger.warning(
            "STEP3 trechos-0 origem_estado=%r origem_cidade=%r destino_estado=%r destino_cidade=%r total=%r initial=%r",
            request.POST.get("trechos-0-origem_estado"),
            request.POST.get("trechos-0-origem_cidade"),
            request.POST.get("trechos-0-destino_estado"),
            request.POST.get("trechos-0-destino_cidade"),
            request.POST.get("trechos-TOTAL_FORMS"),
            request.POST.get("trechos-INITIAL_FORMS"),
        )
        motivo_val = request.POST.get("motivo", "").strip()
        tipo_destino = (request.POST.get("tipo_destino") or "").strip().upper()
        valor_diarias_extenso = request.POST.get("valor_diarias_extenso", "").strip()
        retorno_saida_data_raw = request.POST.get("retorno_saida_data", "").strip()
        retorno_saida_hora_raw = request.POST.get("retorno_saida_hora", "").strip()
        retorno_chegada_data_raw = request.POST.get("retorno_chegada_data", "").strip()
        retorno_chegada_hora_raw = request.POST.get("retorno_chegada_hora", "").strip()
        retorno_saida_data = (
            parse_date(retorno_saida_data_raw) if retorno_saida_data_raw else None
        )
        retorno_saida_hora = (
            parse_time(retorno_saida_hora_raw) if retorno_saida_hora_raw else None
        )
        retorno_chegada_data = (
            parse_date(retorno_chegada_data_raw) if retorno_chegada_data_raw else None
        )
        retorno_chegada_hora = (
            parse_time(retorno_chegada_hora_raw) if retorno_chegada_hora_raw else None
        )
        post_data = _prune_trailing_trechos_post(request.POST, "trechos")
        # TEMP: desabilitado para evitar sobrescrever dados originais
        # post_data = _sync_trechos_origens_post(post_data, "trechos")
        if settings.DEBUG:
            origin = post_data.get("trechos-0-origem_estado", "")
            origin_city = post_data.get("trechos-0-origem_cidade", "")
            logger.debug(
                "trecho 0 origem_estado=%s origem_cidade=%s",
                origin,
                origin_city,
            )
        formset = TrechoFormSet(post_data, instance=dummy_oficio, prefix="trechos")
        trechos_serialized = _serialize_trechos_from_post(post_data)
        _update_wizard_data(
            request,
            {
                "trechos": trechos_serialized,
                "tipo_destino": tipo_destino,
                "motivo": motivo_val,
                "valor_diarias_extenso": valor_diarias_extenso,
                "retorno": {
                    "retorno_saida_data": retorno_saida_data_raw,
                    "retorno_saida_hora": retorno_saida_hora_raw,
                    "retorno_chegada_data": retorno_chegada_data_raw,
                    "retorno_chegada_hora": retorno_chegada_hora_raw,
                },
            },
        )

        if formset.is_valid():
            forms_validas = [form.cleaned_data for form in formset.forms if form.cleaned_data]
            if not forms_validas:
                erro = "Adicione ao menos um trecho para o roteiro."
            else:
                if not tipo_destino:
                    erro = "Selecione o tipo de destino para calcular as diarias."
                if not retorno_saida_data or not retorno_chegada_data:
                    erro = erro or "Informe as datas de saida e chegada do retorno."
                primeiro = forms_validas[0]
                ultimo = forms_validas[-1]
                sede_estado = primeiro.get("origem_estado")
                sede_cidade = primeiro.get("origem_cidade")
                destino_estado = ultimo.get("destino_estado")
                destino_cidade = ultimo.get("destino_cidade")
                saida_sede_dt = _combine_date_time(
                    primeiro.get("saida_data"), primeiro.get("saida_hora")
                )
                retorno_saida_dt = _combine_date_time(
                    retorno_saida_data, retorno_saida_hora
                )
                retorno_chegada_dt = _combine_date_time(
                    retorno_chegada_data, retorno_chegada_hora
                )
                if not erro and not saida_sede_dt:
                    erro = "Informe a data e hora de saida da sede no primeiro trecho."
                if not erro and not retorno_chegada_dt:
                    erro = "Informe a data e hora de chegada na sede no retorno."
                if (
                    not erro
                    and retorno_saida_dt
                    and retorno_chegada_dt
                    and retorno_chegada_dt < retorno_saida_dt
                ):
                    erro = "A chegada do retorno deve ocorrer apos a saida."
                if (
                    not erro
                    and saida_sede_dt
                    and retorno_chegada_dt
                    and retorno_chegada_dt < saida_sede_dt
                ):
                    erro = "A chegada na sede deve ocorrer apos a saida da sede."
                if not erro:
                    resultado_diarias = calcular_diarias(
                        tipo_destino=tipo_destino,
                        saida_sede=saida_sede_dt,
                        chegada_sede=retorno_chegada_dt,
                        quantidade_servidores=len(data.get("viajantes_ids", [])),
                    )
                    quantidade_diarias = resultado_diarias.quantidade_diarias_str
                    valor_diarias = formatar_valor_diarias(
                        resultado_diarias.valor_total_oficio
                    )
                    destino_texto = ""
                    if destino_cidade and destino_estado:
                        destino_texto = f"{destino_cidade.nome} / {destino_estado.sigla}"
                    retorno_saida_cidade = _format_trecho_local(
                        destino_cidade, destino_estado
                    )
                    retorno_chegada_cidade = _format_trecho_local(
                        sede_cidade, sede_estado
                    )
                    _update_wizard_data(
                        request,
                        {
                            "destino": destino_texto or data.get("destino", ""),
                            "retorno_saida_cidade": retorno_saida_cidade,
                            "retorno_chegada_cidade": retorno_chegada_cidade,
                            "quantidade_diarias": quantidade_diarias,
                            "valor_diarias": valor_diarias,
                        },
                    )
                    return redirect("oficio_step4")
        else:
            erro = "Revise os campos obrigatorios do roteiro."
    else:
        formset = TrechoFormSet(
            prefix="trechos", instance=dummy_oficio, initial=trechos_initial
        )

    return render(
        request,
        "viagens/oficio_step3.html",
        {
            "erro": erro,
            "formset": formset,
            "motivo": motivo_val,
            "tipo_destino": tipo_destino,
            "retorno_saida_data": retorno_saida_data_raw,
            "retorno_saida_hora": retorno_saida_hora_raw,
            "retorno_chegada_data": retorno_chegada_data_raw,
            "retorno_chegada_hora": retorno_chegada_hora_raw,
            "valor_diarias_extenso": valor_diarias_extenso,
            "quantidade_servidores": len(data.get("viajantes_ids", [])),
        },
    )


@require_http_methods(["GET", "POST"])
def oficio_step4(request):
    wizard_data = _get_wizard_data(request)
    if not _wizard_has_step4_data(wizard_data):
        return redirect("oficio_step3")

    if request.method == "POST":
        oficio_obj, _ = _finalize_oficio_from_wizard(wizard_data)
        _clear_wizard_data(request)
        return redirect("oficio_editar", oficio_id=oficio_obj.id)

    context = _build_step4_context(wizard_data)
    return render(request, "viagens/oficio_step4.html", context)


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
                "values": request.POST,
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
    paginator = Paginator(viajantes, 12)
    page_obj = paginator.get_page(request.GET.get("page"))
    params = request.GET.copy()
    params.pop("page", None)
    querystring = params.urlencode()
    return render(
        request,
        "viagens/viajantes_list.html",
        {
            "viajantes": page_obj,
            "page_obj": page_obj,
            "querystring": querystring,
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
                "values": request.POST,
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
    paginator = Paginator(veiculos, 12)
    page_obj = paginator.get_page(request.GET.get("page"))
    params = request.GET.copy()
    params.pop("page", None)
    querystring = params.urlencode()
    return render(
        request,
        "viagens/veiculos_list.html",
        {
            "veiculos": page_obj,
            "page_obj": page_obj,
            "querystring": querystring,
            "q": q,
            "combustivel_choices": combustivel_choices,
            "combustivel_selecionado": combustivel,
        },
    )


@require_http_methods(["GET"])
def oficios_lista(request):
    q = request.GET.get("q", "").strip()

    oficios = Oficio.objects.select_related(
        "veiculo",
        "motorista_viajante",
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
            | models.Q(motorista_viajante__nome__icontains=q)
            | models.Q(veiculo__placa__icontains=q)
            | models.Q(veiculo__modelo__icontains=q)
            | models.Q(viajantes__nome__icontains=q)
            | models.Q(cidade_destino__nome__icontains=q)
            | models.Q(cidade_sede__nome__icontains=q)
            | models.Q(estado_destino__sigla__icontains=q)
            | models.Q(estado_sede__sigla__icontains=q)
        ).distinct()

    oficios = oficios.order_by("-created_at")
    paginator = Paginator(oficios, 10)
    page_obj = paginator.get_page(request.GET.get("page"))
    params = request.GET.copy()
    params.pop("page", None)
    querystring = params.urlencode()
    return render(
        request,
        "viagens/oficios_list.html",
        {
            "oficios": page_obj,
            "page_obj": page_obj,
            "querystring": querystring,
            "q": q,
        },
    )


@require_http_methods(["GET", "POST"])
def viajante_editar(request, viajante_id: int):
    viajante = get_object_or_404(Viajante, id=viajante_id)
    erros = {}

    if request.method == "POST":
        if request.POST.get("action") == "delete":
            viajante.delete()
            return redirect("viajantes_lista")

        nome = request.POST.get("nome", "").strip()
        rg = request.POST.get("rg", "").strip()
        cpf = request.POST.get("cpf", "").strip()
        cargo = request.POST.get("cargo", "").strip()
        telefone = request.POST.get("telefone", "").strip()

        if not nome:
            erros["nome"] = "Informe o nome."
        if not rg:
            erros["rg"] = "Informe o RG."
        if not cpf:
            erros["cpf"] = "Informe o CPF."
        if not cargo:
            erros["cargo"] = "Informe o cargo."

        if not erros:
            viajante.nome = nome
            viajante.rg = rg
            viajante.cpf = cpf
            viajante.cargo = cargo
            viajante.telefone = telefone
            viajante.save()
            return redirect("viajantes_lista")

    return render(
        request,
        "viagens/viajante_edit.html",
        {"viajante": viajante, "erros": erros, "cargo_choices": _get_cargo_choices()},
    )


@require_http_methods(["GET", "POST"])
def veiculo_editar(request, veiculo_id: int):
    veiculo = get_object_or_404(Veiculo, id=veiculo_id)
    erros = {}

    if request.method == "POST":
        if request.POST.get("action") == "delete":
            veiculo.delete()
            return redirect("veiculos_lista")

        placa = request.POST.get("placa", "").strip()
        placa_norm = _normalizar_placa(placa) if placa else ""
        modelo = request.POST.get("modelo", "").strip()
        combustivel = request.POST.get("combustivel", "").strip()

        if not placa_norm:
            erros["placa"] = "Informe a placa."
        if not modelo:
            erros["modelo"] = "Informe o modelo."
        if not combustivel:
            erros["combustivel"] = "Informe o combustivel."

        if placa_norm and Veiculo.objects.filter(placa__iexact=placa_norm).exclude(
            id=veiculo.id
        ).exists():
            erros["placa"] = "Ja existe um veiculo com esta placa."

        if not erros:
            veiculo.placa = placa_norm
            veiculo.modelo = modelo
            veiculo.combustivel = combustivel
            veiculo.save()
            return redirect("veiculos_lista")

    return render(
        request,
        "viagens/veiculo_edit.html",
        {
            "veiculo": veiculo,
            "erros": erros,
            "combustivel_choices": _get_combustivel_choices(),
        },
    )


@require_http_methods(["GET", "POST"])
def oficio_editar(request, oficio_id: int):
    oficio = get_object_or_404(
        Oficio.objects.prefetch_related("trechos"),
        id=oficio_id,
    )
    erros = {}
    servidores_form = ServidoresSelectForm(
        initial={"servidores": oficio.viajantes.all()}
    )
    motorista_form = MotoristaSelectForm(
        initial={"motorista": oficio.motorista_viajante_id or ""}
    )
    motorista_nome_manual = ""
    data_criacao = ""
    if oficio.created_at:
        data_criacao = timezone.localtime(oficio.created_at).strftime("%d/%m/%Y")
    motorista_carona = oficio.motorista_carona
    motorista_preview = oficio.motorista_viajante
    TrechoFormSet = inlineformset_factory(
        Oficio,
        Trecho,
        form=TrechoForm,
        extra=0,
        formset=OrderedTrechoInlineFormSet,
        can_delete=False,
    )
    formset = TrechoFormSet(instance=oficio, prefix="trechos")
    tipo_destino_val = oficio.tipo_destino or ""
    retorno_saida_data_val = (
        oficio.retorno_saida_data.isoformat() if oficio.retorno_saida_data else ""
    )
    retorno_saida_hora_val = (
        oficio.retorno_saida_hora.strftime("%H:%M") if oficio.retorno_saida_hora else ""
    )
    retorno_chegada_data_val = (
        oficio.retorno_chegada_data.isoformat() if oficio.retorno_chegada_data else ""
    )
    retorno_chegada_hora_val = (
        oficio.retorno_chegada_hora.strftime("%H:%M") if oficio.retorno_chegada_hora else ""
    )
    valor_diarias_extenso_val = oficio.valor_diarias_extenso or ""

    if request.method == "POST":
        if request.POST.get("action") == "delete":
            oficio.delete()
            return redirect("oficios_lista")

        servidores_form = ServidoresSelectForm(request.POST)
        motorista_form = MotoristaSelectForm(request.POST)
        post_data = _prune_trailing_trechos_post(request.POST, "trechos")
        post_data = _sync_trechos_origens_post(post_data, "trechos")
        formset = TrechoFormSet(post_data, instance=oficio, prefix="trechos")

        oficio_val = request.POST.get("oficio", "").strip()
        protocolo = request.POST.get("protocolo", "").strip()
        assunto = request.POST.get("assunto", "").strip()
        placa = request.POST.get("placa", "").strip()
        placa_norm = _normalizar_placa(placa) if placa else ""
        modelo = request.POST.get("modelo", "").strip()
        combustivel = request.POST.get("combustivel", "").strip()
        motorista_nome_manual = request.POST.get("motorista_nome", "").strip()
        motorista_obj = None
        if motorista_form.is_valid():
            motorista_obj = motorista_form.cleaned_data.get("motorista")
            motorista_preview = motorista_obj or motorista_preview
        motorista_oficio = request.POST.get("motorista_oficio", "").strip()
        motorista_protocolo = request.POST.get("motorista_protocolo", "").strip()
        motivo = request.POST.get("motivo", "").strip()
        tipo_destino_val = (request.POST.get("tipo_destino") or "").strip().upper()
        retorno_saida_data_val = request.POST.get("retorno_saida_data", "").strip()
        retorno_saida_hora_val = request.POST.get("retorno_saida_hora", "").strip()
        retorno_chegada_data_val = request.POST.get("retorno_chegada_data", "").strip()
        retorno_chegada_hora_val = request.POST.get("retorno_chegada_hora", "").strip()
        valor_diarias_extenso_val = request.POST.get("valor_diarias_extenso", "").strip()
        servidores_ids = []
        if servidores_form.is_valid():
            servidores_ids = [
                str(item.id) for item in servidores_form.cleaned_data.get("servidores", [])
            ]

        motorista_carona = False
        if motorista_obj:
            motorista_carona = str(motorista_obj.id) not in servidores_ids
        elif motorista_nome_manual:
            motorista_carona = True

        if not oficio_val:
            erros["oficio"] = "Informe o numero do oficio."
        if not protocolo:
            erros["protocolo"] = "Informe o protocolo."
        if motorista_carona and (not motorista_oficio or not motorista_protocolo):
            erros["motorista_oficio"] = "Informe oficio e protocolo do motorista."
        if not formset.is_valid():
            erros["trechos"] = "Revise os trechos do roteiro."
        if not tipo_destino_val:
            erros["tipo_destino"] = "Selecione o tipo de destino."

        if not erros:
            forms_validas = [form for form in formset.forms if form.cleaned_data]
            if not forms_validas:
                erros["trechos"] = "Adicione ao menos um trecho para o roteiro."
            retorno_saida_data = (
                parse_date(retorno_saida_data_val) if retorno_saida_data_val else None
            )
            retorno_saida_hora = (
                parse_time(retorno_saida_hora_val) if retorno_saida_hora_val else None
            )
            retorno_chegada_data = (
                parse_date(retorno_chegada_data_val)
                if retorno_chegada_data_val
                else None
            )
            retorno_chegada_hora = (
                parse_time(retorno_chegada_hora_val)
                if retorno_chegada_hora_val
                else None
            )
            retorno_saida_dt = _combine_date_time(retorno_saida_data, retorno_saida_hora)
            retorno_chegada_dt = _combine_date_time(
                retorno_chegada_data, retorno_chegada_hora
            )
            primeiro = forms_validas[0].cleaned_data if forms_validas else {}
            ultimo = forms_validas[-1].cleaned_data if forms_validas else {}
            sede_estado = primeiro.get("origem_estado")
            sede_cidade = primeiro.get("origem_cidade")
            destino_estado = ultimo.get("destino_estado")
            destino_cidade = ultimo.get("destino_cidade")
            saida_sede_dt = _combine_date_time(
                primeiro.get("saida_data"), primeiro.get("saida_hora")
            )

            if not retorno_saida_data or not retorno_chegada_data:
                erros["retorno"] = "Informe as datas de saida e chegada do retorno."
            if not saida_sede_dt:
                erros["trechos"] = "Informe a data de saida no primeiro trecho."
            if not retorno_chegada_dt:
                erros["retorno"] = "Informe a data e hora de chegada do retorno."
            if (
                retorno_saida_dt
                and retorno_chegada_dt
                and retorno_chegada_dt < retorno_saida_dt
            ):
                erros["retorno"] = "A chegada do retorno deve ocorrer apos a saida."
            if (
                saida_sede_dt
                and retorno_chegada_dt
                and retorno_chegada_dt < saida_sede_dt
            ):
                erros["retorno"] = "A chegada na sede deve ocorrer apos a saida."

        if not erros:
            motorista_nome = (
                motorista_obj.nome if motorista_obj else motorista_nome_manual
            )
            resultado_diarias = calcular_diarias(
                tipo_destino=tipo_destino_val,
                saida_sede=saida_sede_dt,
                chegada_sede=retorno_chegada_dt,
                quantidade_servidores=len(servidores_ids),
            )
            destino_texto = ""
            if destino_cidade and destino_estado:
                destino_texto = f"{destino_cidade.nome} / {destino_estado.sigla}"
            retorno_saida_cidade = _format_trecho_local(destino_cidade, destino_estado)
            retorno_chegada_cidade = _format_trecho_local(sede_cidade, sede_estado)
            oficio.oficio = oficio_val
            oficio.protocolo = protocolo
            oficio.assunto = assunto
            oficio.tipo_destino = tipo_destino_val
            oficio.destino = destino_texto or oficio.destino
            oficio.estado_sede = sede_estado
            oficio.cidade_sede = sede_cidade
            oficio.estado_destino = destino_estado
            oficio.cidade_destino = destino_cidade
            oficio.retorno_saida_cidade = retorno_saida_cidade
            oficio.retorno_saida_data = retorno_saida_data
            oficio.retorno_saida_hora = retorno_saida_hora
            oficio.retorno_chegada_cidade = retorno_chegada_cidade
            oficio.retorno_chegada_data = retorno_chegada_data
            oficio.retorno_chegada_hora = retorno_chegada_hora
            oficio.quantidade_diarias = resultado_diarias.quantidade_diarias_str
            oficio.valor_diarias = formatar_valor_diarias(
                resultado_diarias.valor_total_oficio
            )
            # TODO: gerar valor_diarias_extenso automaticamente quando houver conversor.
            oficio.valor_diarias_extenso = valor_diarias_extenso_val
            oficio.placa = placa_norm or placa
            oficio.modelo = modelo
            oficio.combustivel = combustivel
            oficio.motorista = motorista_nome
            oficio.motorista_oficio = motorista_oficio
            oficio.motorista_protocolo = motorista_protocolo
            oficio.motorista_carona = motorista_carona
            oficio.motorista_viajante = motorista_obj
            oficio.motivo = motivo
            with transaction.atomic():
                oficio.save()
                if servidores_form.is_valid():
                    oficio.viajantes.set(servidores_form.cleaned_data["servidores"])

                existentes = list(oficio.trechos.order_by("ordem"))
                for idx, form in enumerate(forms_validas):
                    origem_estado = form.cleaned_data.get("origem_estado")
                    origem_cidade = form.cleaned_data.get("origem_cidade")
                    destino_estado = form.cleaned_data.get("destino_estado")
                    destino_cidade = form.cleaned_data.get("destino_cidade")
                    trecho_data = form.save(commit=False)
                    if idx < len(existentes):
                        trecho = existentes[idx]
                    else:
                        trecho = Trecho(oficio=oficio)
                    trecho.ordem = idx + 1
                    trecho.origem_estado = origem_estado
                    trecho.origem_cidade = origem_cidade
                    trecho.destino_estado = destino_estado
                    trecho.destino_cidade = destino_cidade
                    trecho.saida_data = trecho_data.saida_data
                    trecho.saida_hora = trecho_data.saida_hora
                    trecho.chegada_data = trecho_data.chegada_data
                    trecho.chegada_hora = trecho_data.chegada_hora
                    trecho.save()

                if len(existentes) > len(forms_validas):
                    extras = existentes[len(forms_validas) :]
                    Trecho.objects.filter(id__in=[item.id for item in extras]).delete()

            return redirect("oficios_lista")

    if servidores_form.is_valid():
        selected_qs = servidores_form.cleaned_data.get("servidores")
        selected_viajantes = list(selected_qs.order_by("nome")) if selected_qs else []
    else:
        selected_viajantes = list(oficio.viajantes.all().order_by("nome"))

    return render(
        request,
        "viagens/oficio_edit.html",
        {
            "oficio": oficio,
            "erros": erros,
            "formset": formset,
            "servidores_form": servidores_form,
            "motorista_form": motorista_form,
            "motorista_nome": motorista_nome_manual
            or ("" if oficio.motorista_viajante_id else oficio.motorista),
            "motorista_preview": motorista_preview,
            "motorista_carona": motorista_carona,
            "preview_viajantes": _viajantes_payload(selected_viajantes),
            "selected_viajantes": selected_viajantes,
            "data_criacao": data_criacao,
            "combustivel_choices": _get_combustivel_choices(),
            "tipo_destino": tipo_destino_val,
            "retorno_saida_data": retorno_saida_data_val,
            "retorno_saida_hora": retorno_saida_hora_val,
            "retorno_chegada_data": retorno_chegada_data_val,
            "retorno_chegada_hora": retorno_chegada_hora_val,
            "valor_diarias_extenso": valor_diarias_extenso_val,
            "quantidade_servidores": len(selected_viajantes),
        },
    )

# viagens/views.py (adicione perto das outras views)

@require_GET
def oficio_download_docx(request, oficio_id: int):
    oficio = get_object_or_404(
        Oficio.objects.prefetch_related("viajantes", "trechos"),
        id=oficio_id,
    )

    buf = build_oficio_docx_bytes(oficio)

    filename = f"oficio_{oficio.oficio or oficio.id}.docx"
    resp = HttpResponse(
        buf.getvalue(),
        content_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )
    resp["Content-Disposition"] = f'attachment; filename="{filename}"'
    return resp


from django.http import HttpResponse
from django.shortcuts import get_object_or_404
from .models import Oficio
from .documents.document import build_oficio_docx_and_pdf_bytes

def oficio_download_pdf(request, oficio_id):
    oficio = get_object_or_404(Oficio, id=oficio_id)

    docx_bytes, pdf_bytes = build_oficio_docx_and_pdf_bytes(oficio)

    response = HttpResponse(pdf_bytes, content_type="application/pdf")
    response["Content-Disposition"] = f'attachment; filename="oficio_{oficio.oficio or oficio.id}.pdf"'
    return response



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


def _autocomplete_viajante_payload(viajante: Viajante) -> dict:
    text = viajante.nome
    if viajante.cpf:
        text = f"{text} - {viajante.cpf}"
    return {
        "id": viajante.id,
        "text": text,
        "label": text,
        "nome": viajante.nome,
        "cpf": viajante.cpf,
        "rg": viajante.rg,
        "cargo": viajante.cargo,
    }


@require_GET
def servidores_api(request):
    q = request.GET.get("q", "").strip()
    queryset = Viajante.objects.all()
    if q:
        viajantes = list(queryset.filter(nome__icontains=q).order_by("nome")[:20])
    else:
        viajantes = list(queryset.order_by("-id")[:20])
        viajantes.sort(key=lambda item: item.nome or "")
    return JsonResponse({"results": [_autocomplete_viajante_payload(v) for v in viajantes]})


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


import json
from django.http import JsonResponse
from django.views.decorators.csrf import csrf_exempt

@csrf_exempt
def validacao_resultado(request):
    if request.method != "POST":
        return JsonResponse({"error": "method not allowed"}, status=405)

    data = json.loads(request.body.decode("utf-8"))
    # aqui você salva no banco ou atualiza a validação
    # exemplo: oficio_id = data["oficio_id"]; status = data["status"]

    return JsonResponse({"ok": True})
