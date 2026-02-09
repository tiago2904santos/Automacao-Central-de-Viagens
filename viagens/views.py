from __future__ import annotations

import re
from datetime import date, datetime, time, timedelta
import unicodedata
from typing import Iterable

from django.conf import settings
from django.contrib import messages
from django.core.paginator import Paginator
from django.db import models, transaction
from django.db.utils import OperationalError, ProgrammingError
from django.db.models import Count, Q
from django.db.models.functions import ExtractMonth, TruncDate
from django.forms import inlineformset_factory
from django.forms.models import BaseInlineFormSet
from django.http import JsonResponse
from django.shortcuts import get_object_or_404, redirect, render
from django.urls import reverse
from django.utils import timezone
from django.utils.dateparse import parse_date, parse_time
from django.views.decorators.http import require_GET, require_http_methods
from .forms import MotoristaSelectForm, ServidoresSelectForm, TrechoForm
from .models import Cargo, Cidade, Estado, Oficio, Trecho, Viajante, Veiculo
from .services.diarias import calcular_diarias, formatar_valor_diarias
from .services.oficio_helpers import build_assunto, infer_tipo_destino, valor_por_extenso_ptbr
from .forms_oficio_config import OficioConfigForm
from .services.oficio_config import get_oficio_config
from django.http import HttpResponse
from django.shortcuts import get_object_or_404
from django.views.decorators.http import require_GET
from .documents.document import build_oficio_docx_bytes
from .documents.document import build_oficio_docx_and_pdf


def _normalizar_placa(placa: str) -> str:
    return placa.replace(" ", "").replace("-", "").upper()


def _normalizar_cargo_key(nome: str) -> str:
    raw = " ".join((nome or "").strip().split())
    if not raw:
        return ""
    raw = raw.casefold()
    raw = unicodedata.normalize("NFKD", raw)
    return "".join(ch for ch in raw if not unicodedata.combining(ch))


def _normalize_destino_choice(value: str | None) -> str:
    raw = (value or "").strip().upper()
    if not raw:
        return Oficio.DestinoChoices.GAB
    if "SESP" in raw:
        return Oficio.DestinoChoices.SESP
    if raw == Oficio.DestinoChoices.SESP:
        return Oficio.DestinoChoices.SESP
    return Oficio.DestinoChoices.GAB


def _normalize_custeio_choice(value: str | None) -> str:
    raw = (value or "").strip().upper()
    if raw == "SEM_ONUS":
        raw = "ONUS_LIMITADOS"
    valid = {
        Oficio.CusteioTipoChoices.UNIDADE,
        Oficio.CusteioTipoChoices.OUTRA_INSTITUICAO,
        Oficio.CusteioTipoChoices.ONUS_LIMITADOS,
    }
    if raw in valid:
        return raw
    return Oficio.CusteioTipoChoices.UNIDADE


def _destino_label_from_code(code: str | None) -> str:
    try:
        return Oficio.DestinoChoices(code or Oficio.DestinoChoices.GAB).label
    except ValueError:
        return Oficio.DestinoChoices.GAB.label

def _placa_valida(placa: str) -> bool:
    if not placa:
        return False
    placa_norm = _normalizar_placa(placa)
    return bool(
        re.fullmatch(r"[A-Z]{3}\d{4}", placa_norm)
        or re.fullmatch(r"[A-Z]{3}\d[A-Z]\d{2}", placa_norm)
    )


def _somente_digitos(valor: str) -> str:
    return re.sub(r"\D", "", valor or "")


def _formatar_telefone(valor: str) -> str:
    digits = _somente_digitos(valor)
    if len(digits) == 11:
        return f"({digits[:2]}) {digits[2:7]}-{digits[7:]}"
    if len(digits) == 10:
        return f"({digits[:2]}) {digits[2:6]}-{digits[6:]}"
    return digits


def _nome_completo(nome: str) -> bool:
    partes = [item for item in (nome or "").strip().split() if item]
    return len(partes) >= 2


def _formatar_oficio_numero(valor: str) -> str:
    if not valor:
        return ""
    if "/" in valor:
        return valor
    ano = timezone.localdate().year
    return f"{valor}/{ano}"


def _formatar_protocolo(valor: str) -> str:
    digits = _somente_digitos(valor)
    if len(digits) < 9:
        return valor
    digits = digits[:9]
    return f"{digits[:2]}.{digits[2:5]}.{digits[5:8]}-{digits[8:]}"


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


def _format_estado_label(estado: Estado | None, fallback: str | None = "") -> str:
    if estado:
        return f"{estado.nome} ({estado.sigla})"
    return (fallback or "").strip()




def _calcular_destino_automatico_from_trechos(trechos):
    for trecho in trechos:
        sigla = None
        if hasattr(trecho, 'destino_estado'):
            estado = trecho.destino_estado
            cidade = trecho.destino_cidade
            if estado:
                sigla = estado.sigla
            elif cidade and cidade.estado:
                sigla = cidade.estado.sigla
        else:
            estado = _resolve_estado(trecho.get('destino_estado'))
            if estado:
                sigla = estado.sigla
            else:
                cidade = _resolve_cidade(trecho.get('destino_cidade'))
                if cidade and cidade.estado:
                    sigla = cidade.estado.sigla
        if sigla and sigla.upper() != 'PR':
            return Oficio.DestinoChoices.SESP
    return Oficio.DestinoChoices.GAB


def _recalcular_destino_no_dado(data):
    trechos = data.get('trechos') or []
    data['destino'] = _calcular_destino_automatico_from_trechos(trechos)
class OrderedTrechoInlineFormSet(BaseInlineFormSet):
    def get_queryset(self):
        return super().get_queryset().order_by("ordem", "id")

DEFAULT_CARGO_CHOICES = [
    "Agente de Polícia Judiciária",
    "Delegado",
    "Administrativo",
    "Assessor",
    "Assessor de Comunicação Social",
    "Papiloscopista",
]

DEFAULT_COMBUSTIVEL_CHOICES = [
    "Gasolina",
    "Etanol",
    "Diesel",
]


def _get_cargo_choices() -> list[str]:
    raw_values: list[str] = []
    try:
        viajante_cargos = (
            Viajante.objects.exclude(cargo__isnull=True)
            .exclude(cargo="")
            .values_list("cargo", flat=True)
            .distinct()
            .order_by("cargo")
        )
        raw_values.extend(viajante_cargos)
    except (OperationalError, ProgrammingError):
        pass
    try:
        cargo_names = (
            Cargo.objects.exclude(nome__isnull=True)
            .exclude(nome="")
            .values_list("nome", flat=True)
            .order_by("nome")
        )
        raw_values.extend(cargo_names)
    except (OperationalError, ProgrammingError):
        pass

    seen: dict[str, str] = {}
    for raw in raw_values:
        cleaned = " ".join((raw or "").strip().split())
        if not cleaned:
            continue
        key = _normalizar_cargo_key(cleaned)
        if not key:
            continue
        if key not in seen:
            seen[key] = cleaned

    ordered = sorted(seen.values(), key=lambda value: value.casefold())
    return ordered


def _buscar_cargo_por_key(nome: str) -> Cargo | None:
    key = _normalizar_cargo_key(nome)
    if not key:
        return None
    try:
        for cargo in Cargo.objects.only("id", "nome"):
            if _normalizar_cargo_key(cargo.nome) == key:
                return cargo
    except (OperationalError, ProgrammingError):
        return None
    return None


def _resolver_cargo_nome(nome: str) -> str:
    raw = " ".join((nome or "").strip().split())
    if not raw:
        return ""
    cargo = _buscar_cargo_por_key(raw)
    return cargo.nome if cargo else raw


def _ensure_cargo_exists(nome: str) -> None:
    raw = (nome or "").strip()
    if not raw:
        return
    try:
        if _buscar_cargo_por_key(raw):
            return
        Cargo.objects.get_or_create(nome=raw)
    except (OperationalError, ProgrammingError):
        return




def _get_carona_oficios_referencia(limit: int = 200, exclude_id: int | None = None) -> list[Oficio]:
    qs = Oficio.objects.order_by("-created_at")
    if exclude_id:
        qs = qs.exclude(id=exclude_id)
    return list(qs[0:limit])

def _get_combustivel_choices() -> list[str]:
    custom = getattr(settings, "COMBUSTIVEL_CHOICES", None)
    if custom:
        return list(custom)
    return list(DEFAULT_COMBUSTIVEL_CHOICES)


def _get_wizard_oficio_id(request) -> int | None:
    raw_id = request.session.get("oficio_wizard_id")
    if isinstance(raw_id, int):
        return raw_id
    if isinstance(raw_id, str) and raw_id.isdigit():
        return int(raw_id)
    return None


def _set_wizard_oficio_id(request, oficio_id: int) -> None:
    request.session["oficio_wizard_id"] = oficio_id
    request.session.modified = True


def _get_wizard_oficio(request, create: bool = False) -> Oficio | None:
    oficio_id = _get_wizard_oficio_id(request)
    oficio = Oficio.objects.filter(id=oficio_id).first() if oficio_id else None
    if oficio:
        return oficio
    if not create:
        return None
    oficio = Oficio.objects.create(status=Oficio.Status.DRAFT)
    _set_wizard_oficio_id(request, oficio.id)
    return oficio


def _hydrate_wizard_data_from_db(oficio: Oficio) -> dict:
    data = _hydrate_edit_session_from_db(oficio)
    data["status"] = oficio.status
    data["oficio_id"] = oficio.id
    return data


def _ensure_wizard_session(request) -> dict:
    oficio = _get_wizard_oficio(request, create=False)
    if not oficio:
        return _get_wizard_data(request)
    data = _hydrate_wizard_data_from_db(oficio)
    request.session["oficio_wizard"] = data
    request.session.modified = True
    return data


def _get_wizard_data(request) -> dict:
    data = request.session.get("oficio_wizard", {})
    oficio = _get_wizard_oficio(request, create=False)
    if oficio and not data:
        data = _hydrate_wizard_data_from_db(oficio)
        request.session["oficio_wizard"] = data
        request.session.modified = True
    if not data.get("destino"):
        data["destino"] = Oficio.DestinoChoices.GAB
        request.session["oficio_wizard"] = data
        request.session.modified = True
    updated = False
    if "custeio_tipo" not in data:
        legacy = data.get("custos")
        data["custeio_tipo"] = _normalize_custeio_choice(legacy)
        updated = True
    if "nome_instituicao_custeio" not in data:
        data["nome_instituicao_custeio"] = ""
        updated = True

    if "carona_oficio_referencia_id" not in data:
        data["carona_oficio_referencia_id"] = ""
        updated = True
    if updated:
        request.session["oficio_wizard"] = data
        request.session.modified = True
    _recalcular_destino_no_dado(data)
    return data


def _update_wizard_data(request, new_data: dict) -> dict:
    data = _get_wizard_data(request)
    data.update(new_data)
    _recalcular_destino_no_dado(data)
    request.session["oficio_wizard"] = data
    request.session.modified = True
    return data


def _clear_wizard_data(request) -> None:
    request.session.pop("oficio_wizard", None)
    request.session.pop("oficio_wizard_id", None)
    request.session.modified = True


def _edit_key(oficio_id: int) -> str:
    return f"oficio_edit_wizard:{oficio_id}"


def _get_edit_data(request, oficio_id: int) -> dict:
    return request.session.get(_edit_key(oficio_id), {})


def _set_edit_data(request, oficio_id: int, data: dict) -> None:
    request.session[_edit_key(oficio_id)] = data
    request.session.modified = True


def _update_edit_data(request, oficio_id: int, new_data: dict) -> dict:
    data = _get_edit_data(request, oficio_id)
    data.update(new_data)
    _recalcular_destino_no_dado(data)
    _set_edit_data(request, oficio_id, data)
    return data


def _clear_edit_data(request, oficio_id: int) -> None:
    request.session.pop(_edit_key(oficio_id), None)
    request.session.modified = True

STEP_VIEW_MAP = {
    "step1": "oficio_edit_step1",
    "step2": "oficio_edit_step2",
    "step3": "oficio_edit_step3",
    "step4": "oficio_edit_step4",
}


def _redirect_to_edit_step(request, oficio_id: int, default_view: str):
    goto_raw = (
        (request.POST.get("goto") or request.POST.get("goto_step") or "")
        .strip()
        .lower()
    )
    target_view = STEP_VIEW_MAP.get(goto_raw) or default_view
    return redirect(target_view, oficio_id=oficio_id)


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


def _reorder_list_by_csv(items: list[dict], order_csv: str | None) -> list[dict]:
    if not items or not order_csv:
        return items
    parts = [part.strip() for part in order_csv.split(",") if part.strip().isdigit()]
    if len(parts) != len(items):
        return items
    try:
        order = [int(part) for part in parts]
    except ValueError:
        return items
    if len(set(order)) != len(order):
        return items
    if any(idx < 0 or idx >= len(items) for idx in order):
        return items
    try:
        return [items[idx] for idx in order]
    except Exception:
        return items


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


def _serialize_sede_destinos_from_post(post_data) -> tuple[str, str, list[dict[str, str]]]:
    sede_uf = (post_data.get("sede_uf") or "").strip().upper()
    sede_cidade = (post_data.get("sede_cidade") or "").strip()

    total_raw = post_data.get("destinos-TOTAL_FORMS")
    try:
        total_forms = int(total_raw or 0)
    except (TypeError, ValueError):
        total_forms = 0

    order_raw = (post_data.get("destinos-order") or "").strip()
    destinos: list[dict[str, str]] = []
    for index in range(total_forms):
        uf = (post_data.get(f"destinos-{index}-uf") or "").strip().upper()
        cidade = (post_data.get(f"destinos-{index}-cidade") or "").strip()
        destinos.append({"uf": uf, "cidade": cidade})
    destinos = _reorder_list_by_csv(destinos, order_raw)

    return sede_uf, sede_cidade, destinos


def _normalize_destinos_for_wizard(destinos_data) -> list[dict[str, str]]:
    if not destinos_data:
        destinos_data = [{}]
    normalized = []
    for destino in destinos_data:
        normalized.append(
            {
                "uf": (destino.get("uf") or "PR").strip().upper(),
                "cidade": (destino.get("cidade") or "").strip(),
            }
        )
    return normalized


def _build_destinos_display(destinos_data) -> list[dict[str, str]]:
    display = []
    for destino in destinos_data:
        uf = destino.get("uf", "").strip().upper()
        cidade = destino.get("cidade", "").strip()
        estado_obj = _resolve_estado(uf)
        cidade_obj = _resolve_cidade(cidade, estado=estado_obj)
        display.append(
            {
                "uf": uf,
                "cidade": cidade,
                "uf_label": _format_estado_label(estado_obj, fallback=uf),
                "cidade_label": _format_trecho_local(cidade_obj, estado_obj),
            }
        )
    return display


def _build_trechos_from_sede_destinos(
    sede_uf: str, sede_cidade: str, destinos_list: list[dict[str, str]]
) -> list[dict[str, str]]:
    trechos: list[dict[str, str]] = []
    origem_estado = (sede_uf or "").strip()
    origem_cidade = (sede_cidade or "").strip()
    if not origem_estado and not origem_cidade:
        return trechos

    current = {"uf": origem_estado, "cidade": origem_cidade}
    for destino in destinos_list:
        destino_estado = (destino.get("uf") or "").strip()
        destino_cidade = (destino.get("cidade") or "").strip()
        if not destino_estado and not destino_cidade:
            continue
        if not current["uf"] and not current["cidade"]:
            break
        trechos.append(
            {
                "origem_estado": current["uf"],
                "origem_cidade": current["cidade"],
                "destino_estado": destino_estado,
                "destino_cidade": destino_cidade,
                "saida_data": "",
                "saida_hora": "",
                "chegada_data": "",
                "chegada_hora": "",
            }
        )
        current = {"uf": destino_estado, "cidade": destino_cidade}

    return trechos


def _merge_datas_horas(
    old_trechos: list[dict[str, str | int]], new_trechos: list[dict[str, str]]
) -> list[dict[str, str | int]]:
    if not new_trechos:
        return []
    def _trecho_key(candidate: dict[str, str | int]) -> tuple[str, str, str, str]:
        return tuple(
            str(candidate.get(field, "") or "").strip()
            for field in ("origem_estado", "origem_cidade", "destino_estado", "destino_cidade")
        )

    merged: list[dict[str, str | int]] = []
    preserved_map: dict[tuple[str, str, str, str], dict[str, str | int]] = {}
    for trecho in old_trechos:
        key = _trecho_key(trecho)
        if any(key):
            preserved_map.setdefault(key, trecho)

    # Preserve dates/horas by matching origem/destino, falling back to positional matching if keys are missing.
    for index, trecho in enumerate(new_trechos):
        entry = {**trecho}
        key = _trecho_key(entry)
        previous = preserved_map.get(key)
        if not previous and index < len(old_trechos):
            previous = old_trechos[index]
        if previous:
            for key_field in ("saida_data", "saida_hora", "chegada_data", "chegada_hora"):
                entry[key_field] = previous.get(key_field, entry.get(key_field, ""))
        merged.append(entry)
    return merged


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


def _hydrate_edit_session_from_db(oficio: Oficio) -> dict:
    trechos = list(
        oficio.trechos.select_related(
            "origem_estado",
            "origem_cidade",
            "destino_estado",
            "destino_cidade",
        ).order_by("ordem", "id")
    )
    viajantes_ids = [str(viajante.id) for viajante in oficio.viajantes.all()]
    trechos_payload: list[dict[str, str]] = []
    for trecho in trechos:
        trechos_payload.append(
            {
                "origem_estado": trecho.origem_estado.sigla if trecho.origem_estado else "",
                "origem_cidade": str(trecho.origem_cidade_id or ""),
                "destino_estado": trecho.destino_estado.sigla if trecho.destino_estado else "",
                "destino_cidade": str(trecho.destino_cidade_id or ""),
                "saida_data": trecho.saida_data.isoformat() if trecho.saida_data else "",
                "saida_hora": trecho.saida_hora.strftime("%H:%M") if trecho.saida_hora else "",
                "chegada_data": trecho.chegada_data.isoformat() if trecho.chegada_data else "",
                "chegada_hora": trecho.chegada_hora.strftime("%H:%M") if trecho.chegada_hora else "",
            }
        )

    sede_uf = ""
    sede_cidade = ""
    if trechos:
        sede_uf = trechos[0].origem_estado.sigla if trechos[0].origem_estado else ""
        sede_cidade = str(trechos[0].origem_cidade_id or "")

    destinos: list[dict[str, str]] = []
    prev_destino: tuple[str, str] | None = None
    for trecho in trechos:
        uf = trecho.destino_estado.sigla if trecho.destino_estado else ""
        cidade = str(trecho.destino_cidade_id or "")
        if not uf and not cidade:
            continue
        current = (uf, cidade)
        if current == prev_destino:
            continue
        destinos.append({"uf": uf, "cidade": cidade})
        prev_destino = current

    motorista_nome = ""
    if not oficio.motorista_viajante_id:
        motorista_nome = oficio.motorista or ""

    retorno = {
        "retorno_saida_data": oficio.retorno_saida_data.isoformat()
        if oficio.retorno_saida_data
        else "",
        "retorno_saida_hora": oficio.retorno_saida_hora.strftime("%H:%M")
        if oficio.retorno_saida_hora
        else "",
        "retorno_chegada_data": oficio.retorno_chegada_data.isoformat()
        if oficio.retorno_chegada_data
        else "",
        "retorno_chegada_hora": oficio.retorno_chegada_hora.strftime("%H:%M")
        if oficio.retorno_chegada_hora
        else "",
    }

    return {
        "oficio": oficio.oficio,
        "protocolo": oficio.protocolo,
        "assunto": oficio.assunto,
        "viajantes_ids": viajantes_ids,
        "placa": oficio.placa,
        "modelo": oficio.modelo,
        "combustivel": oficio.combustivel,
        "tipo_viatura": oficio.tipo_viatura,
        "motorista_id": str(oficio.motorista_viajante_id or ""),
        "motorista_nome": motorista_nome,
        "motorista_oficio": oficio.motorista_oficio,
        "motorista_protocolo": oficio.motorista_protocolo,
        "carona_oficio_referencia_id": str(oficio.carona_oficio_referencia_id or ""),
        "motorista_carona": oficio.motorista_carona,
        "sede_uf": sede_uf,
        "sede_cidade": sede_cidade,
        "destinos": destinos or [{}],
        "trechos": trechos_payload or [{}],
        "retorno": retorno,
        "retorno_saida_cidade": oficio.retorno_saida_cidade,
        "retorno_chegada_cidade": oficio.retorno_chegada_cidade,
        "tipo_destino": oficio.tipo_destino,
        "motivo": oficio.motivo,
        "valor_diarias_extenso": oficio.valor_diarias_extenso,
        "quantidade_diarias": oficio.quantidade_diarias,
        "valor_diarias": oficio.valor_diarias,
        "destino": _calcular_destino_automatico_from_trechos(trechos),
        "erros": {},
        "custeio_tipo": _normalize_custeio_choice(oficio.custeio_tipo or oficio.custos),
        "nome_instituicao_custeio": oficio.nome_instituicao_custeio,
    }


def _ensure_edit_session(request, oficio_id: int) -> dict:
    data = _get_edit_data(request, oficio_id)
    if data:
        return data
    oficio = get_object_or_404(
        Oficio.objects.select_related("motorista_viajante")
        .prefetch_related("viajantes")
        .prefetch_related(
            "trechos__origem_estado",
            "trechos__origem_cidade",
            "trechos__destino_estado",
            "trechos__destino_cidade",
        ),
        id=oficio_id,
    )
    data = _hydrate_edit_session_from_db(oficio)
    _set_edit_data(request, oficio_id, data)
    return data


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
    trechos_obj: list[Trecho] = []
    for trecho in wizard_data.get("trechos") or []:
        trechos_obj.append(
            Trecho(
                destino_estado=_resolve_estado(trecho.get("destino_estado")),
                destino_cidade=_resolve_cidade(trecho.get("destino_cidade")),
                saida_data=parse_date(trecho.get("saida_data"))
                if trecho.get("saida_data")
                else None,
            )
        )
    assunto_payload = build_assunto(Oficio(created_at=timezone.now()), trechos_obj)

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

    destino_code = _calcular_destino_automatico_from_trechos(wizard_data.get("trechos") or [])
    custeio_code = _normalize_custeio_choice(wizard_data.get("custeio_tipo"))
    try:
        custeio_label = Oficio.CusteioTipoChoices(custeio_code).label
    except ValueError:
        custeio_label = Oficio.CusteioTipoChoices.UNIDADE.label
    nome_instituicao_custeio = (wizard_data.get("nome_instituicao_custeio") or "").strip()
    if custeio_code == Oficio.CusteioTipoChoices.OUTRA_INSTITUICAO and nome_instituicao_custeio:
        custeio_label = f"{custeio_label} ? {nome_instituicao_custeio}"
    try:
        destino_display = Oficio.DestinoChoices(destino_code).label
    except ValueError:
        destino_display = Oficio.DestinoChoices.GAB.label
    return {
        "oficio": wizard_data.get("oficio", ""),
        "protocolo": wizard_data.get("protocolo", ""),
        "assunto": assunto_payload["assunto"],
        "destino": destino_code,
        "destino_display": destino_display,
        "placa": wizard_data.get("placa", ""),
        "modelo": wizard_data.get("modelo", ""),
        "combustivel": wizard_data.get("combustivel", ""),
        "custos": custeio_label,
        "nome_instituicao_custeio": nome_instituicao_custeio,
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


def _wizard_status_context(oficio: Oficio | None) -> dict[str, str]:
    if oficio and oficio.status == Oficio.Status.FINAL:
        return {"status_label": "FINALIZADO", "status_class": "badge"}
    return {"status_label": "RASCUNHO", "status_class": "badge badge--danger"}


def _validate_edit_wizard_data(draft: dict) -> dict[str, str]:
    erros: dict[str, str] = {}
    oficio_val = (draft.get("oficio") or "").strip()
    protocolo_val = (draft.get("protocolo") or "").strip()
    viajantes_ids = [str(item) for item in draft.get("viajantes_ids", []) if str(item)]
    placa_val = (draft.get("placa") or "").strip()
    modelo_val = (draft.get("modelo") or "").strip()
    combustivel_val = (draft.get("combustivel") or "").strip()

    if not oficio_val:
        erros["oficio"] = "Informe o numero do oficio."
    if not protocolo_val:
        erros["protocolo"] = "Informe o protocolo."
    if not viajantes_ids:
        erros["viajantes"] = "Selecione ao menos um viajante."
    if not placa_val or not modelo_val or not combustivel_val:
        erros["veiculo"] = "Preencha placa, modelo e combustivel."

    motorista_carona = bool(draft.get("motorista_carona"))
    if motorista_carona:
        motorista_oficio = (draft.get("motorista_oficio") or "").strip()
        motorista_protocolo = (draft.get("motorista_protocolo") or "").strip()
        if not motorista_oficio or not motorista_protocolo:
            erros["motorista_oficio"] = "Informe oficio e protocolo do motorista."

        if not (draft.get("carona_oficio_referencia_id") or "").strip():
            erros["carona_oficio_referencia"] = "Informe o oficio de referencia da carona."

    trechos_data = draft.get("trechos") or []
    if not trechos_data:
        erros["trechos"] = "Adicione ao menos um trecho para o roteiro."
    else:
        for trecho in trechos_data:
            origem_estado = _resolve_estado(trecho.get("origem_estado"))
            origem_cidade = _resolve_cidade(trecho.get("origem_cidade"), estado=origem_estado)
            destino_estado = _resolve_estado(trecho.get("destino_estado"))
            destino_cidade = _resolve_cidade(trecho.get("destino_cidade"), estado=destino_estado)
            if not origem_estado or not origem_cidade or not destino_estado or not destino_cidade:
                erros["trechos"] = "Preencha origem e destino de todos os trechos."
                break
            saida_data = parse_date(trecho.get("saida_data")) if trecho.get("saida_data") else None
            chegada_data = (
                parse_date(trecho.get("chegada_data")) if trecho.get("chegada_data") else None
            )
            saida_hora = parse_time(trecho.get("saida_hora")) if trecho.get("saida_hora") else None
            chegada_hora = (
                parse_time(trecho.get("chegada_hora")) if trecho.get("chegada_hora") else None
            )
            if saida_data and chegada_data:
                saida_dt = _combine_date_time(saida_data, saida_hora)
                chegada_dt = _combine_date_time(chegada_data, chegada_hora)
                if saida_dt and chegada_dt and chegada_dt < saida_dt:
                    erros["trechos"] = "A chegada deve ocorrer no mesmo momento ou apos a saida."
                    break

    tipo_destino = (draft.get("tipo_destino") or "").strip()
    if not tipo_destino:
        erros["tipo_destino"] = "Selecione o tipo de destino."

    retorno = _get_retornodata(draft)
    retorno_saida_data = (
        parse_date(retorno.get("retorno_saida_data")) if retorno.get("retorno_saida_data") else None
    )
    retorno_chegada_data = (
        parse_date(retorno.get("retorno_chegada_data"))
        if retorno.get("retorno_chegada_data")
        else None
    )
    retorno_saida_hora = (
        parse_time(retorno.get("retorno_saida_hora")) if retorno.get("retorno_saida_hora") else None
    )
    retorno_chegada_hora = (
        parse_time(retorno.get("retorno_chegada_hora"))
        if retorno.get("retorno_chegada_hora")
        else None
    )
    if not retorno_saida_data or not retorno_chegada_data:
        erros["retorno"] = "Informe as datas de saida e chegada do retorno."
    else:
        retorno_saida_dt = _combine_date_time(retorno_saida_data, retorno_saida_hora)
        retorno_chegada_dt = _combine_date_time(retorno_chegada_data, retorno_chegada_hora)
        if retorno_saida_dt and retorno_chegada_dt and retorno_chegada_dt < retorno_saida_dt:
            erros["retorno"] = "A chegada do retorno deve ocorrer apos a saida."
        if trechos_data and not erros.get("retorno"):
            primeiro = trechos_data[0]
            saida_sede_dt = _combine_date_time(
                parse_date(primeiro.get("saida_data")) if primeiro.get("saida_data") else None,
                parse_time(primeiro.get("saida_hora")) if primeiro.get("saida_hora") else None,
            )
            if saida_sede_dt and retorno_chegada_dt and retorno_chegada_dt < saida_sede_dt:
                erros["retorno"] = "A chegada na sede deve ocorrer apos a saida da sede."

    return erros


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

    temp_trechos: list[Trecho] = []
    for trecho in trechos_data:
        destino_cidade_tmp = _resolve_cidade(trecho.get("destino_cidade"))
        temp_trechos.append(Trecho(destino_cidade=destino_cidade_tmp))
    tipo_destino_final = infer_tipo_destino(temp_trechos) if temp_trechos else tipo_destino

    resultado_diarias = calcular_diarias(
        tipo_destino=tipo_destino_final,
        saida_sede=saida_sede_dt,
        chegada_sede=retorno_chegada_dt,
        quantidade_servidores=len(viajantes),
    )

    quantidade_diarias = resultado_diarias.quantidade_diarias_str
    valor_diarias = formatar_valor_diarias(resultado_diarias.valor_total_oficio)

    retorno_saida_cidade = _format_trecho_local(destino_cidade, destino_estado)
    retorno_chegada_cidade = _format_trecho_local(sede_cidade, sede_estado)

    destino_code = _calcular_destino_automatico_from_trechos(trechos_serialized)

    with transaction.atomic():
        oficio_obj = Oficio.objects.create(
            oficio=wizard_data.get("oficio", ""),
            protocolo=wizard_data.get("protocolo", ""),
            destino=destino_code,
            assunto=wizard_data.get("assunto", ""),
            tipo_destino=tipo_destino_final,
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
            valor_diarias_extenso=valor_por_extenso_ptbr(valor_diarias),
            placa=placa_norm or placa,
            modelo=modelo,
            combustivel=combustivel,
            motorista=motorista_nome,
            motorista_oficio=wizard_data.get("motorista_oficio", ""),
            motorista_protocolo=wizard_data.get("motorista_protocolo", ""),
            motorista_carona=motorista_carona,
            motorista_viajante=motorista_obj,
            carona_oficio_referencia=_resolve_oficio_by_id(
                wizard_data.get("carona_oficio_referencia_id")
            ),
            motivo=wizard_data.get("motivo", ""),
            custeio_tipo=_normalize_custeio_choice(wizard_data.get("custeio_tipo")),
            nome_instituicao_custeio=(wizard_data.get("nome_instituicao_custeio") or "").strip(),
            veiculo=veiculo,
            tipo_viatura=veiculo.tipo_viatura if veiculo else "",
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
            trecho_obj.save()
            trechos_instances.append(trecho_obj)
        if viajantes:
            oficio_obj.viajantes.set(viajantes)

    assunto_payload = build_assunto(oficio_obj, list(oficio_obj.trechos.all()))
    oficio_obj.assunto = assunto_payload["assunto"]
    oficio_obj.save(update_fields=["assunto"])

    return oficio_obj, viajantes


def _validate_oficio_for_finalize(oficio: Oficio) -> dict[str, str]:
    erros: dict[str, str] = {}
    if not oficio.oficio.strip():
        erros["oficio"] = "Informe o numero do oficio."
    if not oficio.protocolo.strip():
        erros["protocolo"] = "Informe o protocolo."
    if oficio.viajantes.count() == 0:
        erros["viajantes"] = "Selecione ao menos um viajante."
    if not oficio.placa or not oficio.modelo or not oficio.combustivel:
        erros["veiculo"] = "Preencha placa, modelo e combustivel."

    if oficio.motorista_carona and (
        not oficio.motorista_oficio.strip() or not oficio.motorista_protocolo.strip()
    ):
        erros["motorista_oficio"] = "Informe oficio e protocolo do motorista."

    if oficio.motorista_carona and not oficio.carona_oficio_referencia_id:
        erros["carona_oficio_referencia"] = "Informe o oficio de referencia da carona."

    trechos = list(oficio.trechos.select_related("origem_estado", "origem_cidade", "destino_estado", "destino_cidade").order_by("ordem", "id"))
    if not trechos:
        erros["trechos"] = "Adicione ao menos um trecho para o roteiro."
    else:
        for trecho in trechos:
            if not trecho.origem_estado or not trecho.origem_cidade or not trecho.destino_estado or not trecho.destino_cidade:
                erros["trechos"] = "Preencha origem e destino de todos os trechos."
                break
            if trecho.saida_data and trecho.chegada_data:
                saida_dt = _combine_date_time(trecho.saida_data, trecho.saida_hora)
                chegada_dt = _combine_date_time(trecho.chegada_data, trecho.chegada_hora)
                if saida_dt and chegada_dt and chegada_dt < saida_dt:
                    erros["trechos"] = "A chegada deve ocorrer no mesmo momento ou apos a saida."
                    break

    if not oficio.tipo_destino and trechos:
        oficio.tipo_destino = infer_tipo_destino(trechos)
    if not oficio.tipo_destino:
        erros["tipo_destino"] = "Selecione o tipo de destino."

    retorno_saida_data = oficio.retorno_saida_data
    retorno_chegada_data = oficio.retorno_chegada_data
    retorno_saida_hora = oficio.retorno_saida_hora
    retorno_chegada_hora = oficio.retorno_chegada_hora
    if not retorno_saida_data or not retorno_chegada_data:
        erros["retorno"] = "Informe as datas de saida e chegada do retorno."
    else:
        retorno_saida_dt = _combine_date_time(retorno_saida_data, retorno_saida_hora)
        retorno_chegada_dt = _combine_date_time(retorno_chegada_data, retorno_chegada_hora)
        if retorno_saida_dt and retorno_chegada_dt and retorno_chegada_dt < retorno_saida_dt:
            erros["retorno"] = "A chegada do retorno deve ocorrer apos a saida."
    if trechos and not erros.get("retorno"):
        primeiro = trechos[0]
        saida_sede_dt = _combine_date_time(primeiro.saida_data, primeiro.saida_hora)
        if saida_sede_dt and retorno_chegada_dt and retorno_chegada_dt < saida_sede_dt:
            erros["retorno"] = "A chegada na sede deve ocorrer apos a saida da sede."

    if (
        oficio.custeio_tipo == Oficio.CusteioTipoChoices.OUTRA_INSTITUICAO
        and not (oficio.nome_instituicao_custeio or "").strip()
    ):
        erros["nome_instituicao_custeio"] = "Informe a instituição de custeio."

    return erros


def _finalize_oficio_draft(oficio: Oficio) -> tuple[Oficio, list[Viajante]]:
    trechos = list(
        oficio.trechos.select_related(
            "origem_estado",
            "origem_cidade",
            "destino_estado",
            "destino_cidade",
        ).order_by("ordem", "id")
    )
    primeiro = trechos[0]
    ultimo = trechos[-1]
    sede_estado = primeiro.origem_estado
    sede_cidade = primeiro.origem_cidade
    destino_estado = ultimo.destino_estado
    destino_cidade = ultimo.destino_cidade

    saida_sede_dt = _combine_date_time(primeiro.saida_data, primeiro.saida_hora)
    retorno_chegada_dt = _combine_date_time(oficio.retorno_chegada_data, oficio.retorno_chegada_hora)

    oficio.tipo_destino = infer_tipo_destino(trechos)
    resultado_diarias = calcular_diarias(
        tipo_destino=oficio.tipo_destino,
        saida_sede=saida_sede_dt,
        chegada_sede=retorno_chegada_dt,
        quantidade_servidores=oficio.viajantes.count(),
    )

    oficio.estado_sede = sede_estado
    oficio.cidade_sede = sede_cidade
    oficio.estado_destino = destino_estado
    oficio.cidade_destino = destino_cidade
    oficio.retorno_saida_cidade = _format_trecho_local(destino_cidade, destino_estado)
    oficio.retorno_chegada_cidade = _format_trecho_local(sede_cidade, sede_estado)
    oficio.quantidade_diarias = resultado_diarias.quantidade_diarias_str
    oficio.valor_diarias = formatar_valor_diarias(resultado_diarias.valor_total_oficio)
    oficio.valor_diarias_extenso = valor_por_extenso_ptbr(oficio.valor_diarias)
    oficio.status = Oficio.Status.FINAL
    assunto_payload = build_assunto(oficio, trechos)
    oficio.assunto = assunto_payload["assunto"]
    oficio.save()

    return oficio, list(oficio.viajantes.all())






def _normalize_int(value: str | int | None) -> int | None:
    if value is None:
        return None
    if isinstance(value, int):
        return value
    raw = str(value).strip()
    if raw.isdigit():
        return int(raw)
    return None

def _resolve_oficio_by_id(value: str | None) -> Oficio | None:
    if not value:
        return None
    raw = str(value).strip()
    if not raw or not raw.isdigit():
        return None
    return Oficio.objects.filter(id=int(raw)).first()

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


def _sync_trechos_from_serialized(
    oficio: Oficio, trechos_serialized: list[dict[str, str | int]]
) -> list[Trecho]:
    Trecho.objects.filter(oficio=oficio).delete()
    trechos_instances: list[Trecho] = []
    for idx, trecho in enumerate(trechos_serialized or []):
        origem_estado = _resolve_estado(trecho.get("origem_estado"))
        origem_cidade = _resolve_cidade(trecho.get("origem_cidade"), estado=origem_estado)
        destino_estado = _resolve_estado(trecho.get("destino_estado"))
        destino_cidade = _resolve_cidade(trecho.get("destino_cidade"), estado=destino_estado)
        trecho_obj = Trecho(
            oficio=oficio,
            ordem=idx + 1,
            origem_estado=origem_estado,
            origem_cidade=origem_cidade,
            destino_estado=destino_estado,
            destino_cidade=destino_cidade,
            saida_data=parse_date(trecho.get("saida_data")) if trecho.get("saida_data") else None,
            saida_hora=parse_time(trecho.get("saida_hora")) if trecho.get("saida_hora") else None,
            chegada_data=parse_date(trecho.get("chegada_data")) if trecho.get("chegada_data") else None,
            chegada_hora=parse_time(trecho.get("chegada_hora")) if trecho.get("chegada_hora") else None,
        )
        trecho_obj.save()
        trechos_instances.append(trecho_obj)
    return trechos_instances


def _apply_step1_to_oficio(oficio: Oficio, payload: dict) -> None:
    oficio.oficio = payload.get("oficio", "").strip()
    oficio.protocolo = payload.get("protocolo", "").strip()
    oficio.assunto = payload.get("assunto", "").strip()
    oficio.motivo = payload.get("motivo", "").strip()
    oficio.custeio_tipo = _normalize_custeio_choice(payload.get("custeio_tipo") or payload.get("custos"))
    oficio.nome_instituicao_custeio = (payload.get("nome_instituicao_custeio") or "").strip()
    oficio.save()
    viajantes_ids = payload.get("viajantes_ids", [])
    if viajantes_ids:
        viajantes = list(Viajante.objects.filter(id__in=viajantes_ids))
        oficio.viajantes.set(viajantes)
    else:
        oficio.viajantes.clear()


def _apply_step2_to_oficio(oficio: Oficio, payload: dict) -> None:
    placa_val = payload.get("placa", "").strip()
    placa_norm = _normalizar_placa(placa_val) if placa_val else ""
    modelo_val = payload.get("modelo", "").strip()
    combustivel_val = payload.get("combustivel", "").strip()
    veiculo = None
    if placa_norm:
        veiculo = Veiculo.objects.filter(placa__iexact=placa_norm).first()
        if veiculo:
            modelo_val = modelo_val or veiculo.modelo
            combustivel_val = combustivel_val or veiculo.combustivel

    motorista_id = payload.get("motorista_id") or ""
    motorista_nome = payload.get("motorista_nome", "").strip()
    motorista_obj = None
    if motorista_id and str(motorista_id).isdigit():
        motorista_obj = Viajante.objects.filter(id=motorista_id).first()
        if motorista_obj:
            motorista_nome = motorista_obj.nome
    motorista_carona = False
    if motorista_id:
        motorista_carona = str(motorista_id) not in [str(item.id) for item in oficio.viajantes.all()]
    elif motorista_nome:
        motorista_carona = True

    oficio.placa = placa_norm or placa_val
    oficio.modelo = modelo_val
    oficio.combustivel = combustivel_val
    oficio.motorista = motorista_nome
    oficio.motorista_oficio = payload.get("motorista_oficio", "").strip()
    oficio.motorista_protocolo = payload.get("motorista_protocolo", "").strip()
    oficio.motorista_carona = motorista_carona
    oficio.motorista_viajante = motorista_obj
    oficio.carona_oficio_referencia = _resolve_oficio_by_id(payload.get("carona_oficio_referencia_id"))
    oficio.veiculo = veiculo
    if veiculo and veiculo.tipo_viatura:
        oficio.tipo_viatura = veiculo.tipo_viatura
    elif payload.get("tipo_viatura"):
        oficio.tipo_viatura = payload.get("tipo_viatura")
    oficio.save()


def _apply_step3_to_oficio(oficio: Oficio, payload: dict) -> None:
    oficio.motivo = payload.get("motivo", "").strip()
    oficio.tipo_destino = (payload.get("tipo_destino") or "").strip().upper()
    oficio.retorno_saida_data = payload.get("retorno_saida_data")
    oficio.retorno_saida_hora = payload.get("retorno_saida_hora")
    oficio.retorno_chegada_data = payload.get("retorno_chegada_data")
    oficio.retorno_chegada_hora = payload.get("retorno_chegada_hora")

    trechos_serialized = payload.get("trechos") or []
    trechos_instances = _sync_trechos_from_serialized(oficio, trechos_serialized)
    if trechos_instances:
        oficio.tipo_destino = infer_tipo_destino(trechos_instances)

    sede_estado = None
    sede_cidade = None
    destino_estado = None
    destino_cidade = None
    if trechos_instances:
        primeiro = trechos_instances[0]
        ultimo = trechos_instances[-1]
        sede_estado = primeiro.origem_estado
        sede_cidade = primeiro.origem_cidade
        destino_estado = ultimo.destino_estado
        destino_cidade = ultimo.destino_cidade

    oficio.estado_sede = sede_estado
    oficio.cidade_sede = sede_cidade
    oficio.estado_destino = destino_estado
    oficio.cidade_destino = destino_cidade

    retorno_saida_cidade = _format_trecho_local(destino_cidade, destino_estado)
    retorno_chegada_cidade = _format_trecho_local(sede_cidade, sede_estado)
    oficio.retorno_saida_cidade = retorno_saida_cidade
    oficio.retorno_chegada_cidade = retorno_chegada_cidade

    tipo_destino = oficio.tipo_destino
    retorno_saida_data = oficio.retorno_saida_data
    retorno_chegada_data = oficio.retorno_chegada_data
    retorno_saida_hora = oficio.retorno_saida_hora
    retorno_chegada_hora = oficio.retorno_chegada_hora
    if trechos_instances and tipo_destino and retorno_saida_data and retorno_chegada_data:
        saida_sede_dt = _combine_date_time(
            trechos_instances[0].saida_data, trechos_instances[0].saida_hora
        )
        retorno_chegada_dt = _combine_date_time(retorno_chegada_data, retorno_chegada_hora)
        if saida_sede_dt and retorno_chegada_dt:
            resultado = calcular_diarias(
                tipo_destino=tipo_destino,
                saida_sede=saida_sede_dt,
                chegada_sede=retorno_chegada_dt,
                quantidade_servidores=oficio.viajantes.count(),
            )
            oficio.quantidade_diarias = resultado.quantidade_diarias_str
            oficio.valor_diarias = formatar_valor_diarias(resultado.valor_total_oficio)
            oficio.valor_diarias_extenso = valor_por_extenso_ptbr(oficio.valor_diarias)

    oficio.save()


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


def _build_monthly_series(queryset, year: int) -> list[dict]:
    agregados = (
        queryset.filter(created_at__year=year)
        .annotate(mes=ExtractMonth("created_at"))
        .values("mes")
        .annotate(total=Count("id"))
        .order_by("mes")
    )
    mapa = {item["mes"]: item["total"] for item in agregados}
    labels = [
        "Jan",
        "Fev",
        "Mar",
        "Abr",
        "Mai",
        "Jun",
        "Jul",
        "Ago",
        "Set",
        "Out",
        "Nov",
        "Dez",
    ]
    return [
        {"label": labels[idx], "total": mapa.get(idx + 1, 0)}
        for idx in range(12)
    ]


def _build_ranking_series(year: int) -> list[dict]:
    ranking = (
        Viajante.objects.annotate(
            total=Count("oficios", filter=Q(oficios__created_at__year=year), distinct=True)
        )
        .filter(total__gt=0)
        .order_by("-total", "nome")[:8]
    )
    return [{"label": item.nome, "total": item.total} for item in ranking]


def _dashboard_payload(year: int) -> dict:
    oficios_qs = Oficio.objects.all()
    veiculos_qs = Veiculo.objects.all()
    viajantes_qs = Viajante.objects.all()

    oficios_ano = oficios_qs.filter(created_at__year=year).count()

    return {
        "year": year,
        "kpis": {
            "oficios": {
                "total": oficios_ano,
                "periodo": oficios_ano,
                "rotulo_periodo": f"oficios em {year}",
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
        },
        "series": {
            "oficios": _build_monthly_series(oficios_qs, year),
            "ranking": _build_ranking_series(year),
        },
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
    year = int(request.GET.get("ano", timezone.localdate().year))
    payload = _dashboard_payload(year)
    payload["initial_payload"] = {
        "year": payload["year"],
        "kpis": payload["kpis"],
        "series": payload["series"],
    }
    return render(request, "viagens/dashboard.html", payload)


@require_GET
def dashboard_data_api(request):
    year = int(request.GET.get("ano", timezone.localdate().year))
    payload = _dashboard_payload(year)
    return JsonResponse(payload)


@require_http_methods(["GET", "POST"])
def formulario(request):
    data = _get_wizard_data(request)
    erro = ""
    if request.method == "GET":
        if not request.GET.get("resume"):
            _clear_wizard_data(request)
            data = {}
        else:
            data = _ensure_wizard_session(request)
    viajantes = Viajante.objects.order_by("nome")
    servidores_form = ServidoresSelectForm(
        initial={"servidores": data.get("viajantes_ids", [])}
    )
    oficio_obj = _get_wizard_oficio(request, create=False)

    if request.method == "POST":
        goto_step = (request.POST.get("goto_step") or "").strip()
        oficio_val = _formatar_oficio_numero(request.POST.get("oficio", "").strip())
        protocolo_val = _formatar_protocolo(request.POST.get("protocolo", "").strip())
        motivo_val = request.POST.get("motivo", "").strip()
        custeio_tipo_val = _normalize_custeio_choice(request.POST.get("custeio_tipo") or request.POST.get("custos"))
        nome_instituicao_custeio = request.POST.get("nome_instituicao_custeio", "").strip()
        servidores_form = ServidoresSelectForm(request.POST)
        if servidores_form.is_valid():
            viajantes_ids = [
                str(item.id)
                for item in servidores_form.cleaned_data.get("servidores", [])
            ]
        else:
            viajantes_ids = []

        oficio_obj = _get_wizard_oficio(request, create=True)
        payload = {
            "oficio": oficio_val,
            "protocolo": protocolo_val,
            "motivo": motivo_val,
            "viajantes_ids": viajantes_ids,
            "custeio_tipo": custeio_tipo_val,
            "nome_instituicao_custeio": nome_instituicao_custeio,
        }
        _apply_step1_to_oficio(oficio_obj, payload)
        data = _update_wizard_data(request, payload)
        if custeio_tipo_val == Oficio.CusteioTipoChoices.OUTRA_INSTITUICAO and not nome_instituicao_custeio:
            erro = "Informe a instituição de custeio."
        if not erro:
            if goto_step == "1":
                return redirect("formulario")
            if goto_step == "2":
                return redirect("oficio_step2")
            if goto_step == "3":
                return redirect("oficio_step3")
            if goto_step == "4":
                return redirect("oficio_step4")
            return redirect("oficio_step2")

    selected_ids = [str(item) for item in data.get("viajantes_ids", [])]
    selected_viajantes = list(
        Viajante.objects.filter(id__in=selected_ids).order_by("nome")
    )
    data_criacao = ""
    if oficio_obj and oficio_obj.created_at:
        data_criacao = timezone.localtime(oficio_obj.created_at).strftime("%d/%m/%Y")
    status_context = _wizard_status_context(oficio_obj)
    return render(
        request,
        "viagens/form.html",
        {
            "viajantes": viajantes,
            "oficio": data.get("oficio", ""),
            "protocolo": data.get("protocolo", ""),
            "motivo": data.get("motivo", ""),
            "custeio_tipo": data.get("custeio_tipo", Oficio.CusteioTipoChoices.UNIDADE),
            "nome_instituicao_custeio": data.get("nome_instituicao_custeio", ""),
            "custeio_tipo_choices": Oficio.CusteioTipoChoices.choices,
            "data_criacao": data_criacao,
            "selected_ids": selected_ids,
            "selected_viajantes": selected_viajantes,
            "servidores_form": servidores_form,
            "status_label": status_context["status_label"],
            "status_class": status_context["status_class"],
            "erro": erro,
        },
    )


@require_http_methods(["GET", "POST"])
def oficio_step2(request):
    data = _ensure_wizard_session(request)
    oficio_obj = _get_wizard_oficio(request, create=False)

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
        tipo_viatura_val = request.POST.get("tipo_viatura", "").strip()
        motorista_form = MotoristaSelectForm(request.POST)
        motorista_id = ""
        motorista_obj = None
        if motorista_form.is_valid():
            motorista_obj = motorista_form.cleaned_data.get("motorista")
            if motorista_obj:
                motorista_id = str(motorista_obj.id)
        motorista_nome = request.POST.get("motorista_nome", "").strip()
        motorista_oficio = request.POST.get("motorista_oficio", "").strip()
        motorista_protocolo = request.POST.get("motorista_protocolo", "").strip()
        carona_oficio_referencia_id = (request.POST.get("carona_oficio_referencia") or "").strip()
        carona_oficio_referencia = _resolve_oficio_by_id(carona_oficio_referencia_id)

        placa_norm = _normalizar_placa(placa_val) if placa_val else ""
        if placa_norm and (not modelo_val or not combustivel_val):
            veiculo = Veiculo.objects.filter(placa__iexact=placa_norm).first()
            if veiculo:
                modelo_val = modelo_val or veiculo.modelo
                combustivel_val = combustivel_val or veiculo.combustivel
                tipo_viatura_val = tipo_viatura_val or veiculo.tipo_viatura

        motorista_carona = False
        if motorista_id:
            motorista_carona = motorista_id not in [str(item) for item in viajantes_ids]
        elif motorista_nome:
            motorista_carona = True

        erro_carona_ref = ""
        if motorista_carona and not carona_oficio_referencia:
            erro_carona_ref = "Informe o oficio de referencia da carona."

        if not motorista_carona:
            carona_oficio_referencia = None
            carona_oficio_referencia_id = ""

        oficio_obj = _get_wizard_oficio(request, create=True)
        payload = {
            "placa": placa_norm or placa_val,
            "modelo": modelo_val,
            "combustivel": combustivel_val,
            "tipo_viatura": tipo_viatura_val,
            "motorista_id": motorista_id,
            "motorista_nome": motorista_nome,
            "motorista_oficio": motorista_oficio,
            "motorista_protocolo": motorista_protocolo,
            "motorista_carona": motorista_carona,
            "carona_oficio_referencia_id": str(carona_oficio_referencia.id) if carona_oficio_referencia else "",
        }
        if erro_carona_ref:
            status_context = _wizard_status_context(oficio_obj)
            return render(
                request,
                "viagens/oficio_step2.html",
                {
                    "viajantes": Viajante.objects.order_by("nome"),
                    "placa": payload.get("placa", ""),
                    "modelo": payload.get("modelo", ""),
                    "combustivel": payload.get("combustivel", ""),
                    "tipo_viatura": payload.get("tipo_viatura", ""),
                    "combustivel_choices": _get_combustivel_choices(),
                    "motorista_id": payload.get("motorista_id", ""),
                    "motorista_nome": payload.get("motorista_nome", ""),
                    "motorista_oficio": payload.get("motorista_oficio", ""),
                    "motorista_protocolo": payload.get("motorista_protocolo", ""),
                    "motorista_carona": motorista_carona,
                    "viajantes_ids": viajantes_ids,
                    "preview_viajantes": preview_viajantes,
                    "motorista_preview": motorista_obj if motorista_id else None,
                    "motorista_form": motorista_form,
                    "status_label": status_context["status_label"],
                    "status_class": status_context["status_class"],
                    "oficios_referencia": _get_carona_oficios_referencia(),
                    "carona_oficio_referencia_id": _normalize_int(carona_oficio_referencia.id if carona_oficio_referencia else None),
                    "erro_carona_ref": erro_carona_ref,
                },
            )
        _apply_step2_to_oficio(oficio_obj, payload)
        data = _update_wizard_data(request, payload)
        viajantes_ids = data.get("viajantes_ids", [])
        preview_viajantes = _viajantes_payload(
            list(Viajante.objects.filter(id__in=viajantes_ids).order_by("nome"))
        )

        if goto_step == "1":
            return redirect(f"{reverse('formulario')}?resume=1")
        if goto_step == "2":
            return redirect("oficio_step2")
        if goto_step == "3":
            return redirect("oficio_step3")
        if goto_step == "4":
            return redirect("oficio_step4")
        return redirect("oficio_step3")

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

    status_context = _wizard_status_context(oficio_obj)
    return render(
        request,
        "viagens/oficio_step2.html",
        {
            "viajantes": viajantes,
            "placa": data.get("placa", ""),
            "modelo": data.get("modelo", ""),
            "combustivel": data.get("combustivel", ""),
            "tipo_viatura": data.get("tipo_viatura", ""),
            "combustivel_choices": _get_combustivel_choices(),
            "motorista_id": data.get("motorista_id", ""),
            "motorista_nome": motorista_nome_val,
            "motorista_oficio": data.get("motorista_oficio", ""),
            "motorista_protocolo": data.get("motorista_protocolo", ""),
            "motorista_carona": motorista_carona,
            "oficio": data.get("oficio", ""),
            "protocolo": data.get("protocolo", ""),
            "motivo": data.get("motivo", ""),
            "viajantes_ids": viajantes_ids,
            "preview_viajantes": preview_viajantes,
            "motorista_preview": motorista_preview,
            "motorista_form": motorista_form,
            "status_label": status_context["status_label"],
            "status_class": status_context["status_class"],
            "oficios_referencia": _get_carona_oficios_referencia(),
            "carona_oficio_referencia_id": _normalize_int(data.get("carona_oficio_referencia_id")),
            "erro_carona_ref": "",
        },
    )


@require_http_methods(["GET", "POST"])
def oficio_step3(request):
    data = _ensure_wizard_session(request)
    oficio_obj = _get_wizard_oficio(request, create=False)

    motivo_val = data.get("motivo", "")
    tipo_destino = data.get("tipo_destino", "")
    valor_diarias_extenso = data.get("valor_diarias_extenso", "")
    retorno_payload = _get_retornodata(data)
    retorno_saida_data_raw = retorno_payload["retorno_saida_data"]
    retorno_saida_hora_raw = retorno_payload["retorno_saida_hora"]
    retorno_chegada_data_raw = retorno_payload["retorno_chegada_data"]
    retorno_chegada_hora_raw = retorno_payload["retorno_chegada_hora"]

    estados = Estado.objects.order_by("nome")
    sede_uf_raw = (data.get("sede_uf") or "").strip().upper()
    sede_cidade_raw = (data.get("sede_cidade") or "").strip()
    defaults = {}
    if not sede_uf_raw:
        defaults["sede_uf"] = "PR"
    if not sede_cidade_raw:
        curitiba = (
            Cidade.objects.filter(nome__iexact="Curitiba", estado__sigla="PR").first()
        )
        if curitiba:
            defaults["sede_cidade"] = str(curitiba.id)
    if defaults:
        data = _update_wizard_data(request, defaults)
    sede_uf = (data.get("sede_uf") or "").strip().upper() or "PR"
    sede_cidade = (data.get("sede_cidade") or "").strip()
    destinos_session = _normalize_destinos_for_wizard(data.get("destinos"))

    trechos_session = data.get("trechos") or []
    if not trechos_session:
        valid_destinos = [
            destino for destino in destinos_session if destino.get("uf") and destino.get("cidade")
        ]
        trechos_session = _build_trechos_from_sede_destinos(
            sede_uf, sede_cidade, valid_destinos
        )
        _update_wizard_data(request, {"trechos": trechos_session})
    trechos_initial = _normalize_trechos_initial(trechos_session)
    formset_extra = max(1, len(trechos_initial))
    TrechoFormSet = _build_trecho_formset(formset_extra)

    dummy_oficio = Oficio()
    if request.method == "POST":
        goto_step = (request.POST.get("goto_step") or "").strip()
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
        formset = TrechoFormSet(post_data, instance=dummy_oficio, prefix="trechos")
        trechos_serialized = _serialize_trechos_from_post(post_data)
        (
            sede_uf_post,
            sede_cidade_post,
            destinos_raw,
        ) = _serialize_sede_destinos_from_post(post_data)
        valid_destinos = [
            destino for destino in destinos_raw if destino.get("uf") and destino.get("cidade")
        ]
        base_trechos = _build_trechos_from_sede_destinos(
            sede_uf_post, sede_cidade_post, valid_destinos
        )
        if base_trechos:
            trechos_serialized = _merge_datas_horas(trechos_serialized, base_trechos)
        if not trechos_serialized:
            trechos_serialized = [
                {
                    "origem_estado": sede_uf_post,
                    "origem_cidade": sede_cidade_post,
                    "destino_estado": "",
                    "destino_cidade": "",
                    "saida_data": "",
                    "saida_hora": "",
                    "chegada_data": "",
                    "chegada_hora": "",
                }
            ]
        data = _update_wizard_data(
            request,
            {
                "sede_uf": sede_uf_post,
                "sede_cidade": sede_cidade_post,
                "destinos": destinos_raw,
                "trechos": trechos_serialized,
                "tipo_destino": tipo_destino,
                "motivo": data.get("motivo", ""),
                "valor_diarias_extenso": valor_diarias_extenso,
                "retorno": {
                    "retorno_saida_data": retorno_saida_data_raw,
                    "retorno_saida_hora": retorno_saida_hora_raw,
                    "retorno_chegada_data": retorno_chegada_data_raw,
                    "retorno_chegada_hora": retorno_chegada_hora_raw,
                },
            },
        )
        oficio_obj = _get_wizard_oficio(request, create=True)
        _apply_step1_to_oficio(
            oficio_obj,
            {
                "oficio": data.get("oficio", ""),
                "protocolo": data.get("protocolo", ""),
                "motivo": data.get("motivo", ""),
                "viajantes_ids": data.get("viajantes_ids", []),
                "custeio_tipo": data.get("custeio_tipo", Oficio.CusteioTipoChoices.UNIDADE),
                "nome_instituicao_custeio": data.get("nome_instituicao_custeio", ""),
            },
        )
        _apply_step2_to_oficio(
            oficio_obj,
            {
                "placa": data.get("placa", ""),
                "modelo": data.get("modelo", ""),
                "combustivel": data.get("combustivel", ""),
                "tipo_viatura": data.get("tipo_viatura", ""),
                "motorista_id": data.get("motorista_id", ""),
                "motorista_nome": data.get("motorista_nome", ""),
                "motorista_oficio": data.get("motorista_oficio", ""),
                "motorista_protocolo": data.get("motorista_protocolo", ""),
                # assume driver carona determined from data
            },
        )
        _apply_step3_to_oficio(
            oficio_obj,
            {
                "motivo": motivo_val,
                "tipo_destino": tipo_destino,
                "valor_diarias_extenso": valor_diarias_extenso,
                "retorno_saida_data": retorno_saida_data,
                "retorno_saida_hora": retorno_saida_hora,
                "retorno_chegada_data": retorno_chegada_data,
                "retorno_chegada_hora": retorno_chegada_hora,
                "trechos": trechos_serialized,
            },
        )
        if goto_step == "1":
            return redirect(f"{reverse('formulario')}?resume=1")
        if goto_step == "2":
            return redirect("oficio_step2")
        if goto_step == "3":
            return redirect("oficio_step3")
        if goto_step == "4":
            return redirect("oficio_step4")
        return redirect("oficio_step4")
    else:
        formset = TrechoFormSet(
            prefix="trechos", instance=dummy_oficio, initial=trechos_initial
        )

    sede_uf = (data.get("sede_uf") or "").strip().upper() or "PR"
    sede_cidade = (data.get("sede_cidade") or "").strip()
    if not sede_cidade:
        curitiba = (
            Cidade.objects.filter(nome__iexact="Curitiba", estado__sigla="PR").first()
        )
        if curitiba:
            sede_cidade = str(curitiba.id)
    sede_estado = _resolve_estado(sede_uf)
    sede_cidade_obj = _resolve_cidade(sede_cidade, estado=sede_estado)
    sede_label = _format_trecho_local(sede_cidade_obj, sede_estado)
    sede_uf_label = _format_estado_label(sede_estado, fallback=sede_uf)

    destinos_session = _normalize_destinos_for_wizard(data.get("destinos"))
    destinos_display = _build_destinos_display(destinos_session)

    trechos_session = data.get("trechos") or []
    trechos_initial = _normalize_trechos_initial(trechos_session)
    destinos_order = ",".join(str(idx) for idx in range(len(destinos_display)))

    status_context = _wizard_status_context(oficio_obj)
    return render(
        request,
        "viagens/oficio_step3.html",
        {
            "formset": formset,
            "motivo": motivo_val,
            "tipo_destino": tipo_destino,
            "retorno_saida_data": retorno_saida_data_raw,
            "retorno_saida_hora": retorno_saida_hora_raw,
            "retorno_chegada_data": retorno_chegada_data_raw,
            "retorno_chegada_hora": retorno_chegada_hora_raw,
            "valor_diarias_extenso": valor_diarias_extenso,
            "quantidade_servidores": len(data.get("viajantes_ids", [])),
            "estados": estados,
            "sede_uf": sede_uf,
            "sede_uf_label": sede_uf_label,
            "sede_cidade": sede_cidade,
            "sede_label": sede_label,
            "destinos": destinos_display,
            "destinos_total_forms": len(destinos_display),
            "status_label": status_context["status_label"],
            "status_class": status_context["status_class"],
            "destinos_order": destinos_order,
            "destino_display": _destino_label_from_code(
                _calcular_destino_automatico_from_trechos(trechos_session)
            ),
        },
    )


@require_http_methods(["GET", "POST"])
def oficio_step4(request):
    wizard_data = _ensure_wizard_session(request)
    oficio_obj = _get_wizard_oficio(request, create=False)
    erros: dict[str, str] = {}

    if request.method == "POST":
        action = (request.POST.get("action") or "").strip()
        if action == "prev":
            return redirect("oficio_step3")
        if not oficio_obj:
            return redirect("formulario")
        erros = _validate_oficio_for_finalize(oficio_obj)
        if not erros:
            oficio_obj, _ = _finalize_oficio_draft(oficio_obj)
            _clear_wizard_data(request)
            return redirect("oficios_lista")

    context = _build_step4_context(wizard_data)
    status_context = _wizard_status_context(oficio_obj)
    context.update(
        {
            "erros": erros,
            "status_label": status_context["status_label"],
            "status_class": status_context["status_class"],
        }
    )
    return render(request, "viagens/oficio_step4.html", context)


@require_http_methods(["GET", "POST"])
def oficio_edit_step1(request, oficio_id: int):
    data = _ensure_edit_session(request, oficio_id)
    oficio_obj = get_object_or_404(Oficio, id=oficio_id)

    if request.method == "POST":
        oficio_val = _formatar_oficio_numero(request.POST.get("oficio", "").strip())
        protocolo_val = _formatar_protocolo(request.POST.get("protocolo", "").strip())
        motivo_val = request.POST.get("motivo", "").strip()
        custeio_tipo_val = _normalize_custeio_choice(request.POST.get("custeio_tipo") or request.POST.get("custos"))
        nome_instituicao_custeio = request.POST.get("nome_instituicao_custeio", "").strip()
        servidores_form = ServidoresSelectForm(request.POST)
        if servidores_form.is_valid():
            viajantes_ids = [
                str(item.id)
                for item in servidores_form.cleaned_data.get("servidores", [])
            ]
        else:
            viajantes_ids = []

        data = _update_edit_data(
            request,
            oficio_id,
            {
                "oficio": oficio_val,
                "protocolo": protocolo_val,
                "motivo": motivo_val,
                "viajantes_ids": viajantes_ids,
                "custeio_tipo": custeio_tipo_val,
                "nome_instituicao_custeio": nome_instituicao_custeio,
                "erros": {},
            },
        )
        if request.POST.get("action") == "save":
            return oficio_edit_save(request, oficio_id=oficio_id)
        return _redirect_to_edit_step(
            request, oficio_id=oficio_id, default_view="oficio_edit_step2"
        )

    servidores_form = ServidoresSelectForm(
        initial={"servidores": data.get("viajantes_ids", [])}
    )
    selected_ids = [str(item) for item in data.get("viajantes_ids", [])]
    selected_viajantes = list(
        Viajante.objects.filter(id__in=selected_ids).order_by("nome")
    )
    data_criacao = ""
    if oficio_obj.created_at:
        data_criacao = timezone.localtime(oficio_obj.created_at).strftime("%d/%m/%Y")

    return render(
        request,
        "viagens/oficio_edit_step1.html",
        {
            "oficio_id": oficio_id,
            "oficio": data.get("oficio", ""),
            "protocolo": data.get("protocolo", ""),
            "motivo": data.get("motivo", ""),
            "data_criacao": data_criacao,
            "servidores_form": servidores_form,
            "selected_viajantes": selected_viajantes,
            "custeio_tipo": data.get("custeio_tipo", Oficio.CusteioTipoChoices.UNIDADE),
            "nome_instituicao_custeio": data.get("nome_instituicao_custeio", ""),
            "custeio_tipo_choices": Oficio.CusteioTipoChoices.choices,
        },
    )


@require_http_methods(["GET", "POST"])
def oficio_edit_step2(request, oficio_id: int):
    data = _ensure_edit_session(request, oficio_id)

    viajantes_ids = data.get("viajantes_ids", [])
    viajantes_sel = list(Viajante.objects.filter(id__in=viajantes_ids).order_by("nome"))
    preview_viajantes = _viajantes_payload(viajantes_sel)
    motorista_form = MotoristaSelectForm(
        initial={"motorista": data.get("motorista_id", "")}
    )

    if request.method == "POST":
        placa_val = request.POST.get("placa", "").strip()
        modelo_val = request.POST.get("modelo", "").strip()
        combustivel_val = request.POST.get("combustivel", "").strip()
        tipo_viatura_val = request.POST.get("tipo_viatura", "").strip()
        motorista_form = MotoristaSelectForm(request.POST)
        motorista_id = ""
        motorista_obj = None
        if motorista_form.is_valid():
            motorista_obj = motorista_form.cleaned_data.get("motorista")
            if motorista_obj:
                motorista_id = str(motorista_obj.id)
        motorista_nome = request.POST.get("motorista_nome", "").strip()
        motorista_oficio = request.POST.get("motorista_oficio", "").strip()
        motorista_protocolo = request.POST.get("motorista_protocolo", "").strip()
        carona_oficio_referencia_id = (request.POST.get("carona_oficio_referencia") or "").strip()
        carona_oficio_referencia = _resolve_oficio_by_id(carona_oficio_referencia_id)

        placa_norm = _normalizar_placa(placa_val) if placa_val else ""
        if placa_norm and (not modelo_val or not combustivel_val):
            veiculo = Veiculo.objects.filter(placa__iexact=placa_norm).first()
            if veiculo:
                modelo_val = modelo_val or veiculo.modelo
                combustivel_val = combustivel_val or veiculo.combustivel
                tipo_viatura_val = tipo_viatura_val or veiculo.tipo_viatura

        motorista_carona = False
        if motorista_id:
            motorista_carona = motorista_id not in [str(item) for item in viajantes_ids]
        elif motorista_nome:
            motorista_carona = True

        erro_carona_ref = ""
        if motorista_carona and not carona_oficio_referencia:
            erro_carona_ref = "Informe o oficio de referencia da carona."

        if not motorista_carona:
            carona_oficio_referencia = None
            carona_oficio_referencia_id = ""

        if erro_carona_ref:
            return render(
                request,
                "viagens/oficio_edit_step2.html",
                {
                    "oficio_id": oficio_id,
                    "placa": placa_norm or placa_val,
                    "modelo": modelo_val,
                    "combustivel": combustivel_val,
                    "tipo_viatura": tipo_viatura_val,
                    "combustivel_choices": _get_combustivel_choices(),
                    "motorista_id": motorista_id,
                    "motorista_nome": motorista_nome,
                    "motorista_oficio": motorista_oficio,
                    "motorista_protocolo": motorista_protocolo,
                    "motorista_carona": motorista_carona,
                    "oficio": data.get("oficio", ""),
                    "protocolo": data.get("protocolo", ""),
                    "motivo": data.get("motivo", ""),
                    "viajantes_ids": viajantes_ids,
                    "preview_viajantes": preview_viajantes,
                    "motorista_preview": motorista_obj if motorista_id else None,
                    "motorista_form": motorista_form,
                    "oficios_referencia": _get_carona_oficios_referencia(exclude_id=oficio_id),
                    "carona_oficio_referencia_id": _normalize_int(carona_oficio_referencia.id if carona_oficio_referencia else None),
                    "erro_carona_ref": erro_carona_ref,
                },
            )
        data = _update_edit_data(
            request,
            oficio_id,
            {
                "placa": placa_norm or placa_val,
                "modelo": modelo_val,
                "combustivel": combustivel_val,
                "tipo_viatura": tipo_viatura_val,
                "motorista_id": motorista_id,
                "motorista_nome": motorista_nome,
                "motorista_oficio": motorista_oficio,
                "motorista_protocolo": motorista_protocolo,
                "motorista_carona": motorista_carona,
                "carona_oficio_referencia_id": str(carona_oficio_referencia.id) if carona_oficio_referencia else "",
                "erros": {},
            },
        )
        if request.POST.get("action") == "save":
            return oficio_edit_save(request, oficio_id=oficio_id)
        return _redirect_to_edit_step(
            request, oficio_id=oficio_id, default_view="oficio_edit_step3"
        )

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
        "viagens/oficio_edit_step2.html",
        {
            "oficio_id": oficio_id,
            "placa": data.get("placa", ""),
            "modelo": data.get("modelo", ""),
            "combustivel": data.get("combustivel", ""),
            "tipo_viatura": data.get("tipo_viatura", ""),
            "combustivel_choices": _get_combustivel_choices(),
            "motorista_id": data.get("motorista_id", ""),
            "motorista_nome": motorista_nome_val,
            "motorista_oficio": data.get("motorista_oficio", ""),
            "motorista_protocolo": data.get("motorista_protocolo", ""),
            "motorista_carona": motorista_carona,
            "oficio": data.get("oficio", ""),
            "protocolo": data.get("protocolo", ""),
            "motivo": data.get("motivo", ""),
            "viajantes_ids": viajantes_ids,
            "preview_viajantes": preview_viajantes,
            "motorista_preview": motorista_preview,
            "motorista_form": motorista_form,
            "oficios_referencia": _get_carona_oficios_referencia(exclude_id=oficio_id),
            "carona_oficio_referencia_id": _normalize_int(data.get("carona_oficio_referencia_id")),
            "erro_carona_ref": "",
        },
    )


@require_http_methods(["GET", "POST"])
def oficio_edit_step3(request, oficio_id: int):
    data = _ensure_edit_session(request, oficio_id)

    motivo_val = data.get("motivo", "")
    tipo_destino = data.get("tipo_destino", "")
    valor_diarias_extenso = data.get("valor_diarias_extenso", "")
    retorno_payload = _get_retornodata(data)
    retorno_saida_data_raw = retorno_payload["retorno_saida_data"]
    retorno_saida_hora_raw = retorno_payload["retorno_saida_hora"]
    retorno_chegada_data_raw = retorno_payload["retorno_chegada_data"]
    retorno_chegada_hora_raw = retorno_payload["retorno_chegada_hora"]

    estados = Estado.objects.order_by("nome")
    sede_uf_raw = (data.get("sede_uf") or "").strip().upper()
    sede_cidade_raw = (data.get("sede_cidade") or "").strip()
    defaults = {}
    if not sede_uf_raw:
        defaults["sede_uf"] = "PR"
    if not sede_cidade_raw:
        curitiba = (
            Cidade.objects.filter(nome__iexact="Curitiba", estado__sigla="PR").first()
        )
        if curitiba:
            defaults["sede_cidade"] = str(curitiba.id)
    if defaults:
        data = _update_edit_data(request, oficio_id, defaults)
    sede_uf = (data.get("sede_uf") or "").strip().upper() or "PR"
    sede_cidade = (data.get("sede_cidade") or "").strip()
    destinos_session = _normalize_destinos_for_wizard(data.get("destinos"))

    trechos_session = data.get("trechos") or []
    if not trechos_session:
        valid_destinos = [
            destino for destino in destinos_session if destino.get("uf") and destino.get("cidade")
        ]
        trechos_session = _build_trechos_from_sede_destinos(
            sede_uf, sede_cidade, valid_destinos
        )
        _update_edit_data(request, oficio_id, {"trechos": trechos_session})
    trechos_initial = _normalize_trechos_initial(trechos_session)
    formset_extra = max(1, len(trechos_initial))
    TrechoFormSet = _build_trecho_formset(formset_extra)

    dummy_oficio = Oficio()
    if request.method == "POST":
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
        formset = TrechoFormSet(post_data, instance=dummy_oficio, prefix="trechos")
        trechos_serialized = _serialize_trechos_from_post(post_data)
        sede_uf_post, sede_cidade_post, destinos_raw = _serialize_sede_destinos_from_post(
            post_data
        )
        valid_destinos = [
            destino for destino in destinos_raw if destino.get("uf") and destino.get("cidade")
        ]
        base_trechos = _build_trechos_from_sede_destinos(
            sede_uf_post, sede_cidade_post, valid_destinos
        )
        if base_trechos:
            trechos_serialized = _merge_datas_horas(trechos_serialized, base_trechos)
        if not trechos_serialized:
            trechos_serialized = [
                {
                    "origem_estado": sede_uf_post,
                    "origem_cidade": sede_cidade_post,
                    "destino_estado": "",
                    "destino_cidade": "",
                    "saida_data": "",
                    "saida_hora": "",
                    "chegada_data": "",
                    "chegada_hora": "",
                }
            ]
        data = _update_edit_data(
            request,
            oficio_id,
            {
                "sede_uf": sede_uf_post,
                "sede_cidade": sede_cidade_post,
                "destinos": destinos_raw,
                "trechos": trechos_serialized,
                "tipo_destino": tipo_destino,
                "motivo": data.get("motivo", ""),
                "valor_diarias_extenso": valor_diarias_extenso,
                "retorno": {
                    "retorno_saida_data": retorno_saida_data_raw,
                    "retorno_saida_hora": retorno_saida_hora_raw,
                    "retorno_chegada_data": retorno_chegada_data_raw,
                    "retorno_chegada_hora": retorno_chegada_hora_raw,
                },
                "erros": {},
            },
        )

        if trechos_serialized:
            primeiro = trechos_serialized[0]
            ultimo = trechos_serialized[-1]
            sede_estado = _resolve_estado(primeiro.get("origem_estado"))
            sede_cidade_obj = _resolve_cidade(primeiro.get("origem_cidade"))
            destino_estado = _resolve_estado(ultimo.get("destino_estado"))
            destino_cidade = _resolve_cidade(ultimo.get("destino_cidade"))
            retorno_saida_cidade = _format_trecho_local(destino_cidade, destino_estado)
            retorno_chegada_cidade = _format_trecho_local(sede_cidade_obj, sede_estado)
            _update_edit_data(
                request,
                oficio_id,
                {
                    "retorno_saida_cidade": retorno_saida_cidade,
                    "retorno_chegada_cidade": retorno_chegada_cidade,
                },
            )
            if tipo_destino and retorno_chegada_data and retorno_saida_data:
                saida_sede_dt = _combine_date_time(
                    parse_date(primeiro.get("saida_data"))
                    if primeiro.get("saida_data")
                    else None,
                    parse_time(primeiro.get("saida_hora"))
                    if primeiro.get("saida_hora")
                    else None,
                )
                retorno_chegada_dt = _combine_date_time(
                    retorno_chegada_data, retorno_chegada_hora
                )
                if saida_sede_dt and retorno_chegada_dt:
                    resultado_diarias = calcular_diarias(
                        tipo_destino=tipo_destino,
                        saida_sede=saida_sede_dt,
                        chegada_sede=retorno_chegada_dt,
                        quantidade_servidores=len(data.get("viajantes_ids", [])),
                    )
                    _update_edit_data(
                        request,
                        oficio_id,
                        {
                            "quantidade_diarias": resultado_diarias.quantidade_diarias_str,
                            "valor_diarias": formatar_valor_diarias(
                                resultado_diarias.valor_total_oficio
                            ),
                        },
                    )

        if request.POST.get("action") == "save":
            return oficio_edit_save(request, oficio_id=oficio_id)
        return _redirect_to_edit_step(
            request, oficio_id=oficio_id, default_view="oficio_edit_step4"
        )

    formset = TrechoFormSet(prefix="trechos", instance=dummy_oficio, initial=trechos_initial)

    sede_uf = (data.get("sede_uf") or "").strip().upper() or "PR"
    sede_cidade = (data.get("sede_cidade") or "").strip()
    if not sede_cidade:
        curitiba = (
            Cidade.objects.filter(nome__iexact="Curitiba", estado__sigla="PR").first()
        )
        if curitiba:
            sede_cidade = str(curitiba.id)
    sede_estado = _resolve_estado(sede_uf)
    sede_cidade_obj = _resolve_cidade(sede_cidade, estado=sede_estado)
    sede_label = _format_trecho_local(sede_cidade_obj, sede_estado)
    sede_uf_label = _format_estado_label(sede_estado, fallback=sede_uf)

    destinos_session = _normalize_destinos_for_wizard(data.get("destinos"))
    destinos_display = _build_destinos_display(destinos_session)

    trechos_session = data.get("trechos") or []
    trechos_initial = _normalize_trechos_initial(trechos_session)
    destinos_order = ",".join(str(idx) for idx in range(len(destinos_display)))

    return render(
        request,
        "viagens/oficio_edit_step3.html",
        {
            "oficio_id": oficio_id,
            "formset": formset,
            "motivo": motivo_val,
            "tipo_destino": tipo_destino,
            "retorno_saida_data": retorno_saida_data_raw,
            "retorno_saida_hora": retorno_saida_hora_raw,
            "retorno_chegada_data": retorno_chegada_data_raw,
            "retorno_chegada_hora": retorno_chegada_hora_raw,
            "retorno_saida_cidade": data.get("retorno_saida_cidade", ""),
            "retorno_chegada_cidade": data.get("retorno_chegada_cidade", ""),
            "valor_diarias_extenso": valor_diarias_extenso,
            "quantidade_servidores": len(data.get("viajantes_ids", [])),
            "estados": estados,
            "sede_uf": sede_uf,
            "sede_uf_label": sede_uf_label,
            "sede_cidade": sede_cidade,
            "sede_label": sede_label,
            "destinos": destinos_display,
            "destinos_total_forms": len(destinos_display),
            "destinos_order": destinos_order,
        },
    )


@require_http_methods(["GET", "POST"])
def oficio_edit_step4(request, oficio_id: int):
    data = _ensure_edit_session(request, oficio_id)
    if request.method == "POST":
        return _redirect_to_edit_step(
            request, oficio_id=oficio_id, default_view="oficio_edit_step4"
        )
    context = _build_step4_context(data)
    context.update(
        {
            "oficio_id": oficio_id,
            "erros": data.get("erros", {}),
            "salvo": bool(request.GET.get("salvo")),
        }
    )
    return render(request, "viagens/oficio_edit_step4.html", context)


@require_http_methods(["POST"])
def oficio_edit_save(request, oficio_id: int):
    draft = _ensure_edit_session(request, oficio_id)
    erros = _validate_edit_wizard_data(draft)
    if erros:
        _update_edit_data(request, oficio_id, {"erros": erros})
        return redirect("oficio_edit_step4", oficio_id=oficio_id)

    trechos_data = draft.get("trechos") or []
    primeiro = trechos_data[0]
    ultimo = trechos_data[-1]
    sede_estado = _resolve_estado(primeiro.get("origem_estado"))
    sede_cidade = _resolve_cidade(primeiro.get("origem_cidade"), estado=sede_estado)
    destino_estado = _resolve_estado(ultimo.get("destino_estado"))
    destino_cidade = _resolve_cidade(ultimo.get("destino_cidade"), estado=destino_estado)

    retorno_payload = _get_retornodata(draft)
    retorno_saida_data = (
        parse_date(retorno_payload["retorno_saida_data"])
        if retorno_payload.get("retorno_saida_data")
        else None
    )
    retorno_saida_hora = (
        parse_time(retorno_payload["retorno_saida_hora"])
        if retorno_payload.get("retorno_saida_hora")
        else None
    )
    retorno_chegada_data = (
        parse_date(retorno_payload["retorno_chegada_data"])
        if retorno_payload.get("retorno_chegada_data")
        else None
    )
    retorno_chegada_hora = (
        parse_time(retorno_payload["retorno_chegada_hora"])
        if retorno_payload.get("retorno_chegada_hora")
        else None
    )

    saida_sede_dt = _combine_date_time(
        parse_date(primeiro.get("saida_data")) if primeiro.get("saida_data") else None,
        parse_time(primeiro.get("saida_hora")) if primeiro.get("saida_hora") else None,
    )
    retorno_chegada_dt = _combine_date_time(retorno_chegada_data, retorno_chegada_hora)
    temp_trechos: list[Trecho] = []
    for trecho in trechos_data:
        temp_trechos.append(
            Trecho(
                destino_estado=_resolve_estado(trecho.get("destino_estado")),
                destino_cidade=_resolve_cidade(trecho.get("destino_cidade")),
                saida_data=parse_date(trecho.get("saida_data"))
                if trecho.get("saida_data")
                else None,
            )
        )
    tipo_destino_val = infer_tipo_destino(temp_trechos) if temp_trechos else ""
    resultado_diarias = calcular_diarias(
        tipo_destino=tipo_destino_val,
        saida_sede=saida_sede_dt,
        chegada_sede=retorno_chegada_dt,
        quantidade_servidores=len(draft.get("viajantes_ids", [])),
    )

    quantidade_diarias = resultado_diarias.quantidade_diarias_str
    valor_diarias = formatar_valor_diarias(resultado_diarias.valor_total_oficio)

    retorno_saida_cidade = _format_trecho_local(destino_cidade, destino_estado)
    retorno_chegada_cidade = _format_trecho_local(sede_cidade, sede_estado)

    viajantes_ids = draft.get("viajantes_ids", [])
    viajantes = list(Viajante.objects.filter(id__in=viajantes_ids).order_by("nome"))

    placa = (draft.get("placa") or "").strip()
    placa_norm = _normalizar_placa(placa) if placa else ""
    veiculo = (
        Veiculo.objects.filter(placa__iexact=placa_norm).first() if placa_norm else None
    )
    modelo = (draft.get("modelo") or "").strip()
    combustivel = (draft.get("combustivel") or "").strip()

    motorista_id = draft.get("motorista_id") or ""
    motorista_obj = (
        Viajante.objects.filter(id=motorista_id).first()
        if motorista_id and str(motorista_id).isdigit()
        else None
    )
    motorista_nome = motorista_obj.nome if motorista_obj else draft.get("motorista_nome", "")
    motorista_carona = False
    if motorista_obj:
        motorista_carona = str(motorista_obj.id) not in [str(item) for item in viajantes_ids]
    elif motorista_nome:
        motorista_carona = True

    with transaction.atomic():
        oficio_obj = get_object_or_404(Oficio, id=oficio_id)
        oficio_obj.oficio = draft.get("oficio", "")
        oficio_obj.protocolo = draft.get("protocolo", "")
        oficio_obj.assunto = build_assunto(oficio_obj, temp_trechos)["assunto"]
        oficio_obj.tipo_destino = tipo_destino_val
        oficio_obj.estado_sede = sede_estado
        oficio_obj.cidade_sede = sede_cidade
        oficio_obj.estado_destino = destino_estado
        oficio_obj.cidade_destino = destino_cidade
        oficio_obj.retorno_saida_cidade = retorno_saida_cidade
        oficio_obj.retorno_saida_data = retorno_saida_data
        oficio_obj.retorno_saida_hora = retorno_saida_hora
        oficio_obj.retorno_chegada_cidade = retorno_chegada_cidade
        oficio_obj.retorno_chegada_data = retorno_chegada_data
        oficio_obj.retorno_chegada_hora = retorno_chegada_hora
        oficio_obj.quantidade_diarias = quantidade_diarias
        oficio_obj.valor_diarias = valor_diarias
        oficio_obj.valor_diarias_extenso = valor_por_extenso_ptbr(valor_diarias)
        oficio_obj.placa = placa_norm or placa
        oficio_obj.modelo = modelo
        oficio_obj.combustivel = combustivel
        oficio_obj.tipo_viatura = (veiculo.tipo_viatura if veiculo else "") or draft.get(
            "tipo_viatura", ""
        )
        oficio_obj.motorista = motorista_nome
        oficio_obj.motorista_oficio = draft.get("motorista_oficio", "")
        oficio_obj.motorista_protocolo = draft.get("motorista_protocolo", "")
        oficio_obj.motorista_carona = motorista_carona
        oficio_obj.motorista_viajante = motorista_obj
        oficio_obj.carona_oficio_referencia = _resolve_oficio_by_id(
            draft.get("carona_oficio_referencia_id")
        )
        oficio_obj.motivo = draft.get("motivo", "")
        oficio_obj.veiculo = veiculo
        oficio_obj.save()

        oficio_obj.viajantes.set(viajantes)
        oficio_obj.trechos.all().delete()
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
                saida_data=parse_date(trecho.get("saida_data")) if trecho.get("saida_data") else None,
                saida_hora=parse_time(trecho.get("saida_hora")) if trecho.get("saida_hora") else None,
                chegada_data=parse_date(trecho.get("chegada_data")) if trecho.get("chegada_data") else None,
                chegada_hora=parse_time(trecho.get("chegada_hora")) if trecho.get("chegada_hora") else None,
            )
            trecho_obj.save()
            trechos_instances.append(trecho_obj)

    _clear_edit_data(request, oficio_id)
    return redirect(f"{reverse('oficio_edit_step4', args=[oficio_id])}?salvo=1")


@require_http_methods(["GET"])
def oficio_edit_cancel(request, oficio_id: int):
    _clear_edit_data(request, oficio_id)
    return redirect("oficios_lista")


@require_http_methods(["GET", "POST"])
def viajante_cadastro(request):
    if request.method == "POST":
        nome = request.POST.get("nome", "").strip().upper()
        rg = _somente_digitos(request.POST.get("rg", ""))
        cpf = _somente_digitos(request.POST.get("cpf", ""))
        cargo = (request.POST.get("cargo", "") or "").strip()
        cargo_novo = request.POST.get("cargo_novo", "").strip()
        telefone = _formatar_telefone(request.POST.get("telefone", ""))
        if cargo_novo:
            cargo = cargo_novo
        cargo = _resolver_cargo_nome(cargo)
        if nome and rg and cpf and cargo and _nome_completo(nome):
            _ensure_cargo_exists(cargo)
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
                "erro": "Preencha nome completo, RG, CPF e cargo.",
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
        tipo_viatura = (
            request.POST.get("tipo_viatura", "") or "DESCARACTERIZADA"
        ).strip().upper()
        if not _placa_valida(placa):
            return render(
                request,
                "viagens/veiculo_form.html",
                {
                    "erro": "Informe uma placa valida (AAA1234 ou AAA1A23).",
                    "combustivel_choices": _get_combustivel_choices(),
                    "values": request.POST,
                },
            )
        if placa_norm and modelo and combustivel:
            veiculo, created = Veiculo.objects.get_or_create(
                placa=placa_norm,
                defaults={
                    "modelo": modelo,
                    "combustivel": combustivel,
                    "tipo_viatura": tipo_viatura,
                },
            )
            if not created:
                veiculo.modelo = modelo
                veiculo.combustivel = combustivel
                veiculo.tipo_viatura = tipo_viatura
                veiculo.save(update_fields=["modelo", "combustivel", "tipo_viatura"])
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
    params = request.GET.copy()
    params.pop("page", None)
    querystring = params.urlencode()
    dev_safety = False

    try:
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
    except OperationalError:
        # DEV SAFETY: evita crash quando o banco ainda não criou o campo `status`.
        dev_safety = True
        paginator = Paginator(Oficio.objects.none(), 10)

    page_obj = paginator.get_page(request.GET.get("page"))
    return render(
        request,
        "viagens/oficios_list.html",
        {
            "oficios": page_obj,
            "page_obj": page_obj,
            "querystring": querystring,
            "q": q,
            "dev_safety": dev_safety,
        },
    )


@require_http_methods(["GET"])
def oficio_draft_resume(request, oficio_id: int):
    oficio = get_object_or_404(Oficio, id=oficio_id, status=Oficio.Status.DRAFT)
    _set_wizard_oficio_id(request, oficio.id)
    data = _hydrate_wizard_data_from_db(oficio)
    request.session["oficio_wizard"] = data
    request.session.modified = True
    return redirect("oficio_step4")


@require_http_methods(["GET", "POST"])
def viajante_editar(request, viajante_id: int):
    viajante = get_object_or_404(Viajante, id=viajante_id)
    erros = {}

    if request.method == "POST":
        if request.POST.get("action") == "delete":
            viajante.delete()
            return redirect("viajantes_lista")

        nome = request.POST.get("nome", "").strip().upper()
        rg = _somente_digitos(request.POST.get("rg", ""))
        cpf = _somente_digitos(request.POST.get("cpf", ""))
        cargo = (request.POST.get("cargo", "") or "").strip()
        cargo_novo = request.POST.get("cargo_novo", "").strip()
        telefone = _formatar_telefone(request.POST.get("telefone", ""))
        if cargo_novo:
            cargo = cargo_novo
        cargo = _resolver_cargo_nome(cargo)

        if not nome or not _nome_completo(nome):
            erros["nome"] = "Informe nome e sobrenome."
        if not rg:
            erros["rg"] = "Informe o RG."
        if not cpf:
            erros["cpf"] = "Informe o CPF."
        if not cargo:
            erros["cargo"] = "Informe o cargo."

        if not erros:
            _ensure_cargo_exists(cargo)
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
        tipo_viatura = (
            request.POST.get("tipo_viatura", "") or "DESCARACTERIZADA"
        ).strip().upper()

        if not _placa_valida(placa):
            erros["placa"] = "Informe uma placa valida (AAA1234 ou AAA1A23)."
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
            veiculo.tipo_viatura = tipo_viatura
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
    custeio_tipo_val = _normalize_custeio_choice(oficio.custeio_tipo or oficio.custos)
    nome_instituicao_custeio = oficio.nome_instituicao_custeio
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
        carona_oficio_referencia_id = (request.POST.get("carona_oficio_referencia") or "").strip()
        carona_oficio_referencia = _resolve_oficio_by_id(carona_oficio_referencia_id)
        motivo = request.POST.get("motivo", "").strip()
        tipo_destino_val = (request.POST.get("tipo_destino") or "").strip().upper()
        retorno_saida_data_val = request.POST.get("retorno_saida_data", "").strip()
        retorno_saida_hora_val = request.POST.get("retorno_saida_hora", "").strip()
        retorno_chegada_data_val = request.POST.get("retorno_chegada_data", "").strip()
        retorno_chegada_hora_val = request.POST.get("retorno_chegada_hora", "").strip()
        valor_diarias_extenso_val = request.POST.get("valor_diarias_extenso", "").strip()
        custeio_tipo_val = _normalize_custeio_choice(request.POST.get("custeio_tipo") or request.POST.get("custos"))
        nome_instituicao_custeio = request.POST.get("nome_instituicao_custeio", "").strip()
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

        if not motorista_carona:
            carona_oficio_referencia = None

        if not oficio_val:
            erros["oficio"] = "Informe o numero do oficio."
        if not protocolo:
            erros["protocolo"] = "Informe o protocolo."
        if motorista_carona and (not motorista_oficio or not motorista_protocolo):
            erros["motorista_oficio"] = "Informe oficio e protocolo do motorista."

        if motorista_carona and not carona_oficio_referencia:
            erros["carona_oficio_referencia"] = "Informe o oficio de referencia da carona."
        if not formset.is_valid():
            erros["trechos"] = "Revise os trechos do roteiro."
        if not tipo_destino_val:
            erros["tipo_destino"] = "Selecione o tipo de destino."

        if custeio_tipo_val == Oficio.CusteioTipoChoices.OUTRA_INSTITUICAO and not nome_instituicao_custeio:
            erros["nome_instituicao_custeio"] = "Informe a instituição de custeio."

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
            retorno_saida_cidade = _format_trecho_local(destino_cidade, destino_estado)
            retorno_chegada_cidade = _format_trecho_local(sede_cidade, sede_estado)
            oficio.oficio = oficio_val
            oficio.protocolo = protocolo
            oficio.assunto = assunto
            oficio.tipo_destino = tipo_destino_val
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
            oficio.carona_oficio_referencia = carona_oficio_referencia
            oficio.motivo = motivo
            oficio.custeio_tipo = custeio_tipo_val
            oficio.nome_instituicao_custeio = nome_instituicao_custeio
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
            "custeio_tipo": custeio_tipo_val,
            "nome_instituicao_custeio": nome_instituicao_custeio,
            "custeio_tipo_choices": Oficio.CusteioTipoChoices.choices,
            "retorno_saida_data": retorno_saida_data_val,
            "retorno_saida_hora": retorno_saida_hora_val,
            "retorno_chegada_data": retorno_chegada_data_val,
            "retorno_chegada_hora": retorno_chegada_hora_val,
            "valor_diarias_extenso": valor_diarias_extenso_val,
            "quantidade_servidores": len(selected_viajantes),
            "oficios_referencia": _get_carona_oficios_referencia(exclude_id=oficio.id),
            "carona_oficio_referencia_id": _normalize_int(oficio.carona_oficio_referencia_id),
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
            "tipo_viatura": veiculo.tipo_viatura,
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


@require_http_methods(["POST"])
def cargo_criar(request):
    try:
        payload = json.loads(request.body.decode("utf-8"))
    except Exception:
        payload = request.POST

    nome = " ".join((payload.get("nome") or "").strip().split())
    if not nome:
        return JsonResponse({"error": "Informe o nome do cargo."}, status=400)

    cargo = _buscar_cargo_por_key(nome)
    if cargo:
        return JsonResponse({"id": cargo.id, "nome": cargo.nome})

    cargo = Cargo.objects.create(nome=nome)
    return JsonResponse({"id": cargo.id, "nome": cargo.nome})


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
            "tipo_viatura": veiculo.tipo_viatura,
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
            "tipo_viatura": veiculo.tipo_viatura,
            "label": f"{veiculo.placa} - {veiculo.modelo}",
        }
        for veiculo in veiculos
    ]
    return JsonResponse({"results": payload})


@require_http_methods(["GET", "POST"])
def modal_viajante_form(request):
    if request.method == "POST":
        nome = request.POST.get("nome", "").strip().upper()
        rg = _somente_digitos(request.POST.get("rg", ""))
        cpf = _somente_digitos(request.POST.get("cpf", ""))
        cargo = (request.POST.get("cargo", "") or "").strip()
        cargo_novo = request.POST.get("cargo_novo", "").strip()
        telefone = _formatar_telefone(request.POST.get("telefone", ""))
        if cargo_novo:
            cargo = cargo_novo
        cargo = _resolver_cargo_nome(cargo)

        erros = {}
        if not nome or not _nome_completo(nome):
            erros["nome"] = "Informe nome e sobrenome."
        if not rg:
            erros["rg"] = "Informe o RG."
        if not cpf:
            erros["cpf"] = "Informe o CPF."
        if not cargo:
            erros["cargo"] = "Informe o cargo."

        if not erros:
            _ensure_cargo_exists(cargo)
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
        tipo_viatura = request.POST.get("tipo_viatura", "").strip().upper()

        erros = {}
        if not _placa_valida(placa):
            erros["placa"] = "Informe uma placa valida (AAA1234 ou AAA1A23)."
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
                tipo_viatura=tipo_viatura,
            )
            return JsonResponse(
                {
                    "success": True,
                    "item": {
                        "id": veiculo.id,
                        "placa": veiculo.placa,
                        "modelo": veiculo.modelo,
                        "combustivel": veiculo.combustivel,
                        "tipo_viatura": veiculo.tipo_viatura,
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


@require_http_methods(["GET", "POST"])
def configuracoes_oficio(request):
    config = get_oficio_config()

    if request.method == "POST":
        form = OficioConfigForm(request.POST, request.FILES, instance=config)
        if form.is_valid():
            form.save()
            messages.success(request, "Configuracoes do oficio atualizadas.")
            return redirect("config_oficio")
    else:
        form = OficioConfigForm(instance=config)

    return render(
        request,
        "viagens/configuracoes_oficio.html",
        {
            "form": form,
            "config": config,
        },
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
