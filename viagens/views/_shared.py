from __future__ import annotations

import logging
import os
import re
from datetime import date, datetime, time, timedelta
import json
from urllib.parse import quote, urlencode
from urllib.error import HTTPError, URLError
from urllib.request import urlopen
import unicodedata
from typing import Iterable

from django.conf import settings
from django.contrib import messages
from django.contrib.auth.decorators import login_required
from django.core.exceptions import ValidationError
from django.core.paginator import Paginator
from django.db import models, transaction
from django.db.utils import OperationalError, ProgrammingError
from django.db.models import Count, Q
from django.db.models.functions import ExtractMonth, TruncDate
from django.forms import formset_factory, inlineformset_factory
from django.forms.models import BaseInlineFormSet
from django.http import HttpResponseBadRequest, JsonResponse
from django.shortcuts import get_object_or_404, redirect, render
from django.urls import reverse
from django.utils import timezone
from django.utils.dateparse import parse_date, parse_time
from django.utils.http import url_has_allowed_host_and_scheme
from django.views.decorators.http import require_GET, require_http_methods, require_POST
from django.views.decorators.csrf import csrf_exempt
from ..forms import (
    DiariasSimplesForm,
    EfetivoLinhaForm,
    MotoristaSelectForm,
    MotoristaTransporteForm,
    NovoCargoForm,
    OficioNumeracaoForm,
    OrdemServicoForm,
    PlanoTrabalhoStep1Form,
    PlanoTrabalhoStep2Form,
    PlanoTrabalhoStep3Form,
    ServidoresSelectForm,
    TrechoForm,
    ViajanteNormalizeForm,
)
from ..models import (
    Cargo,
    Cidade,
    CoordenadorMunicipal,
    DocumentoEventoArquivo,
    Efetivo,
    Estado,
    Evento,
    EventoProtocoloArquivo,
    ModeloJustificativa,
    Oficio,
    OficioCounter,
    OrdemServico,
    PlanoTrabalho,
    PlanoTrabalhoAtividade,
    PlanoTrabalhoLocalAtuacao,
    PlanoTrabalhoMeta,
    PlanoTrabalhoRecurso,
    Roteiro,
    TrechoRoteiro,
    get_next_ordem_num,
    get_next_plano_num,
    TermoAutorizacao,
    Trecho,
    Viajante,
    Veiculo,
)
from ..diarias import PeriodMarker, calculate_simple_diarias_total
from ..services.diarias_unified import (
    calculate_diarias_from_markers,
    calculate_diarias_from_periods_payload,
    derive_financeiro_diarias,
    normalize_diarias_resultado,
    parse_decimal_br,
)
from ..services.oficio_helpers import build_assunto, infer_tipo_destino, valor_por_extenso_ptbr
from ..services.plano_trabalho import (
    DEFAULT_COORDENADOR_PLANO_CARGO,
    DEFAULT_COORDENADOR_PLANO_NOME,
    DEFAULT_UNIDADE_MOVEL_TEXTO,
    build_coordenacao_formatada,
    destinos_labels,
    efetivo_total_servidores,
    format_data_extenso_br,
    format_lista_portugues,
    format_periodo_evento_extenso,
    formatar_horario_intervalo,
    format_total_servidores,
    formatar_solicitante_exibicao,
    normalize_destinos_payload,
    normalize_efetivo_payload,
    metas_from_atividades,
    normalize_atividades_selecionadas,
    normalize_solicitantes,
    parse_horario_atendimento_intervalo,
    has_coordenador_municipal,
)
from ..forms_oficio_config import OficioConfigForm
from ..services.oficio_config import get_oficio_config
from ..utils.normalize import (
    format_oficio_num,
    format_cpf,
    format_phone,
    format_protocolo_num,
    format_rg,
    normalize_digits,
    normalize_oficio_num,
    normalize_protocolo_num,
    normalize_rg,
    normalize_upper_text,
    split_oficio_num,
)
from django.http import HttpResponse
from ..documents.document import (
    AssinaturaObrigatoriaError,
    DocxPdfConversionError,
    MotoristaCaronaValidationError,
    build_oficio_docx_and_pdf_bytes,
    build_oficio_docx_bytes,
    build_termo_autorizacao_payload_docx_bytes,
    build_termo_autorizacao_docx_bytes,
    docx_bytes_to_pdf_bytes,
)
from ..documents.generator import generate_all_documents
from ..documents.justificativa import (
    build_justificativa_context_from_config,
    build_justificativa_docx_bytes,
    get_justificativa_placeholders,
)
from ..documents.ordem_servico import build_ordem_servico_docx_bytes
from ..documents.plano_trabalho import build_plano_trabalho_docx_bytes


logger = logging.getLogger(__name__)
OFICIO_RESERVED_SESSION_KEY = "oficio_reserved"
VALIDACAO_TOKEN = os.environ.get("VALIDACAO_API_TOKEN", "")


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

def _normalize_nome_instituicao_custeio(custeio_tipo: str, nome: str | None) -> str:
    if custeio_tipo != Oficio.CusteioTipoChoices.OUTRA_INSTITUICAO:
        return ""
    return (nome or "").strip()



def _get_sede_cidade_default_id() -> str:
    config = get_oficio_config()
    default_id = getattr(config, "sede_cidade_default_id", None)
    if default_id:
        return str(default_id)
    curitiba = Cidade.objects.filter(nome__iexact="Curitiba", estado__sigla="PR").first()
    return str(curitiba.id) if curitiba else ""


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
    return normalize_digits(valor)


def _formatar_telefone(valor: str) -> str:
    return normalize_digits(valor)


def _nome_completo(nome: str) -> bool:
    partes = [item for item in (nome or "").strip().split() if item]
    return len(partes) >= 2


def _formatar_oficio_numero(valor: str) -> str:
    return normalize_oficio_num(valor)


def _formatar_protocolo(valor: str) -> str:
    protocolo = normalize_protocolo_num(valor)
    return protocolo if len(protocolo) == 9 else ""


def _parse_oficio_parts(
    oficio_raw: str | None,
    *,
    numero_raw: str | None = None,
    ano_raw: str | None = None,
    default_year: int | None = None,
) -> tuple[int | None, int | None, str]:
    numero_digits = normalize_digits(numero_raw or "")
    ano_digits = normalize_digits(ano_raw or "")

    if oficio_raw and not numero_digits:
        parsed_numero, parsed_ano = split_oficio_num(oficio_raw)
        if parsed_numero is not None:
            numero_digits = str(parsed_numero)
        if parsed_ano is not None and not ano_digits:
            ano_digits = str(parsed_ano)

    if numero_digits and not ano_digits and default_year:
        ano_digits = str(default_year)
    if not numero_digits:
        ano_digits = ""

    numero_int = int(numero_digits) if numero_digits else None
    ano_int = int(ano_digits[-4:]) if ano_digits else None
    return (numero_int, ano_int, format_oficio_num(numero_int, ano_int))


def _viajantes_payload(viajantes: Iterable[Viajante]) -> list[dict]:
    return [
        {
            "id": viajante.id, # type: ignore
            "nome": viajante.nome,
            "rg": format_rg(viajante.rg),
            "cpf": format_cpf(viajante.cpf),
            "cargo": viajante.cargo,
            "telefone": format_phone(viajante.telefone),
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


def _is_truthy_post(value: str | None) -> bool:
    return (value or "").strip().lower() in {"1", "true", "on", "yes", "sim"}


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


def _coerce_date_value(value: date | str | None) -> date | None:
    if isinstance(value, date):
        return value
    if not value:
        return None
    return parse_date(str(value))


def _coerce_time_value(value: time | str | None) -> time | None:
    if isinstance(value, time):
        return value
    if not value:
        return None
    return parse_time(str(value))


def _build_period_markers_from_serialized_trechos(
    trechos_data: list[dict[str, str | int]],
) -> list[PeriodMarker]:
    markers: list[PeriodMarker] = []
    for trecho in trechos_data or []:
        if not isinstance(trecho, dict):
            continue
        if not any(str(trecho.get(field, "")).strip() for field in TRECHO_FIELDS):
            continue

        saida_data = _coerce_date_value(trecho.get("saida_data"))
        saida_hora = _coerce_time_value(trecho.get("saida_hora"))
        saida_dt = _combine_date_time(saida_data, saida_hora)
        if not saida_dt:
            continue

        destino_estado_obj = _resolve_estado(trecho.get("destino_estado"))
        destino_cidade_obj = _resolve_cidade(
            trecho.get("destino_cidade"),
            estado=destino_estado_obj,
        )
        destino_uf = ""
        if destino_estado_obj:
            destino_uf = destino_estado_obj.sigla
        elif destino_cidade_obj and destino_cidade_obj.estado:
            destino_uf = destino_cidade_obj.estado.sigla

        markers.append(
            PeriodMarker(
                saida=saida_dt,
                destino_cidade=destino_cidade_obj.nome if destino_cidade_obj else "",
                destino_uf=destino_uf,
            )
        )

    if not markers:
        raise ValueError("Preencha datas e horas para calcular.")
    return markers


def _build_period_markers_from_trechos(trechos: list[Trecho]) -> list[PeriodMarker]:
    markers: list[PeriodMarker] = []
    for trecho in trechos:
        saida_dt = _combine_date_time(trecho.saida_data, trecho.saida_hora)
        if not saida_dt:
            continue
        destino_estado = trecho.destino_estado
        destino_cidade = trecho.destino_cidade
        destino_uf = ""
        if destino_estado:
            destino_uf = destino_estado.sigla
        elif destino_cidade and destino_cidade.estado:
            destino_uf = destino_cidade.estado.sigla

        markers.append(
            PeriodMarker(
                saida=saida_dt,
                destino_cidade=destino_cidade.nome if destino_cidade else "",
                destino_uf=destino_uf,
            )
        )

    if not markers:
        raise ValueError("Preencha datas e horas para calcular.")
    return markers


def _calculate_periodized_diarias_for_serialized_trechos(
    trechos_data: list[dict[str, str | int]],
    retorno_chegada_data: date | str | None,
    retorno_chegada_hora: time | str | None,
    *,
    quantidade_servidores: int,
) -> dict:
    chegada_data = _coerce_date_value(retorno_chegada_data)
    chegada_hora = _coerce_time_value(retorno_chegada_hora)
    chegada_final = _combine_date_time(chegada_data, chegada_hora)
    if not chegada_final:
        raise ValueError("Preencha datas e horas para calcular.")

    markers = _build_period_markers_from_serialized_trechos(trechos_data)
    servidores = int(quantidade_servidores or 0)
    return calculate_diarias_from_markers(
        markers=markers,
        chegada_final_sede=chegada_final,
        total_servidores=servidores,
    )


def _calculate_periodized_diarias_for_trechos(
    trechos: list[Trecho],
    retorno_chegada_data: date | str | None,
    retorno_chegada_hora: time | str | None,
    *,
    quantidade_servidores: int,
) -> dict:
    chegada_data = _coerce_date_value(retorno_chegada_data)
    chegada_hora = _coerce_time_value(retorno_chegada_hora)
    chegada_final = _combine_date_time(chegada_data, chegada_hora)
    if not chegada_final:
        raise ValueError("Preencha datas e horas para calcular.")

    markers = _build_period_markers_from_trechos(trechos)
    servidores = int(quantidade_servidores or 0)
    return calculate_diarias_from_markers(
        markers=markers,
        chegada_final_sede=chegada_final,
        total_servidores=servidores,
    )


def _diarias_totais_fields(resultado: dict) -> tuple[str, str, str]:
    totais = normalize_diarias_resultado(resultado).get("totais", {})
    quantidade_diarias = str(totais.get("total_diarias", "") or "")
    valor_diarias = str(totais.get("total_valor", "") or "")
    valor_extenso = str(totais.get("valor_extenso", "") or "")
    if valor_diarias and not valor_extenso:
        valor_extenso = valor_por_extenso_ptbr(valor_diarias)
    return quantidade_diarias, valor_diarias, valor_extenso


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


def _cargo_options_for_plano_step2() -> list[dict[str, object]]:
    options: list[dict[str, object]] = []
    try:
        cargos = list(Cargo.objects.filter(ativo=True).exclude(nome="").order_by("ordem", "nome"))
    except (OperationalError, ProgrammingError):
        cargos = []
    for cargo in cargos:
        options.append({"id": cargo.id, "nome": cargo.nome})
    return options


def _default_efetivo_rows_from_oficio(oficio: Oficio) -> list[dict[str, object]]:
    counts: dict[int, dict[str, object]] = {}
    for viajante in oficio.viajantes.all():
        cargo_nome = _resolver_cargo_nome(viajante.cargo or "")
        if not cargo_nome:
            continue
        _ensure_cargo_exists(cargo_nome)
        cargo = _buscar_cargo_por_key(cargo_nome)
        if not cargo:
            continue
        if cargo.id not in counts:
            counts[cargo.id] = {"cargo_id": cargo.id, "cargo": cargo.nome, "quantidade": 0}
        counts[cargo.id]["quantidade"] = int(counts[cargo.id]["quantidade"]) + 1
    return sorted(counts.values(), key=lambda item: str(item.get("cargo", "")).casefold())


def _get_combustivel_choices() -> list[str]:
    custom = getattr(settings, "COMBUSTIVEL_CHOICES", None)
    if custom:
        return list(custom)
    return list(DEFAULT_COMBUSTIVEL_CHOICES)


def reserve_next_oficio_number(ano: int) -> int:
    with transaction.atomic():
        counter, _ = OficioCounter.objects.select_for_update().get_or_create(
            ano=ano,
            defaults={"last_numero": 0},
        )
        counter.last_numero += 1
        counter.save(update_fields=["last_numero", "updated_at"])
        return counter.last_numero


def _get_reserved_oficio(request) -> dict | None:
    raw = request.session.get(OFICIO_RESERVED_SESSION_KEY)
    if not isinstance(raw, dict):
        return None
    ano_raw = normalize_digits(str(raw.get("ano", "")))
    numero_raw = normalize_digits(str(raw.get("numero", "")))
    if not ano_raw or not numero_raw:
        return None
    return {"ano": int(ano_raw), "numero": int(numero_raw)}


def _set_reserved_oficio(request, *, ano: int, numero: int) -> None:
    request.session[OFICIO_RESERVED_SESSION_KEY] = {"ano": int(ano), "numero": int(numero)}
    request.session.modified = True


def _ensure_reserved_oficio_for_year(request, ano: int) -> dict:
    reserved = _get_reserved_oficio(request)
    if reserved and reserved["ano"] == ano and reserved["numero"] > 0:
        return reserved
    numero = reserve_next_oficio_number(ano)
    reserved = {"ano": int(ano), "numero": int(numero)}
    _set_reserved_oficio(request, **reserved)
    return reserved


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
    create_kwargs: dict[str, object] = {"status": Oficio.Status.DRAFT}
    reserved = _get_reserved_oficio(request)
    if reserved:
        create_kwargs["numero"] = reserved["numero"]
        create_kwargs["ano"] = reserved["ano"]
        create_kwargs["oficio"] = format_oficio_num(reserved["numero"], reserved["ano"])
    oficio = Oficio.objects.create(**create_kwargs)
    _set_wizard_oficio_id(request, oficio.id)
    evento_id = request.session.get("wizard_evento_id")
    if evento_id:
        oficio.evento_id = int(evento_id)
        oficio.save(update_fields=["evento_id"])
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
    request.session.pop(OFICIO_RESERVED_SESSION_KEY, None)
    request.session.pop("wizard_evento_id", None)
    request.session.pop("wizard_evento_return_url", None)
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
    normalized: list[dict[str, str]] = []
    for trecho in trechos_data:
        if not isinstance(trecho, dict):
            continue
        normalized.append(
            {
                "origem_estado": str(trecho.get("origem_estado") or ""),
                "origem_cidade": str(trecho.get("origem_cidade") or ""),
                "destino_estado": str(trecho.get("destino_estado") or ""),
                "destino_cidade": str(trecho.get("destino_cidade") or ""),
                "saida_data": str(trecho.get("saida_data") or ""),
                "saida_hora": str(trecho.get("saida_hora") or ""),
                "chegada_data": str(trecho.get("chegada_data") or ""),
                "chegada_hora": str(trecho.get("chegada_hora") or ""),
            }
        )
    return normalized or [{}]


def _oficio_step3_trechos_from_db(oficio) -> list[dict[str, str | int]]:
    if not oficio:
        return []

    trechos = (
        oficio.trechos.select_related(
            "origem_estado",
            "origem_cidade",
            "destino_estado",
            "destino_cidade",
        )
        .order_by("ordem", "id")
    )
    result: list[dict[str, str | int]] = []
    for trecho in trechos:
        result.append(
            {
                "origem_estado": trecho.origem_estado.sigla if trecho.origem_estado else "",
                "origem_cidade": str(trecho.origem_cidade_id or trecho.origem_cidade or ""),
                "destino_estado": trecho.destino_estado.sigla if trecho.destino_estado else "",
                "destino_cidade": str(trecho.destino_cidade_id or trecho.destino_cidade or ""),
                "saida_data": trecho.saida_data.isoformat() if trecho.saida_data else "",
                "saida_hora": trecho.saida_hora.strftime("%H:%M") if trecho.saida_hora else "",
                "chegada_data": trecho.chegada_data.isoformat() if trecho.chegada_data else "",
                "chegada_hora": trecho.chegada_hora.strftime("%H:%M") if trecho.chegada_hora else "",
                "tempo_viagem_minutos": (
                    trecho.tempo_viagem_minutos if hasattr(trecho, "tempo_viagem_minutos") else ""
                ),
            }
        )
    return result


def _seed_trechos_from_evento_roteiros(evento_id: int) -> list[dict[str, str | int]]:
    """
    Converte roteiros do evento (TrechoRoteiro) no formato do wizard de ofício (trechos session).
    Usado em oficio_step3 quando wizard_evento_id está na sessão e ainda não há trechos.
    """
    roteiros = list(
        Roteiro.objects.filter(evento_id=evento_id, ativo=True)
        .order_by("id")
        .prefetch_related(
            "trechos__origem_estado",
            "trechos__origem_cidade",
            "trechos__destino_estado",
            "trechos__destino_cidade",
        )
    )
    result: list[dict[str, str | int]] = []
    for roteiro in roteiros:
        for tr in roteiro.trechos.order_by("ordem"):
            result.append(
                {
                    "origem_estado": tr.origem_estado.sigla if tr.origem_estado else "",
                    "origem_cidade": str(tr.origem_cidade_id or tr.cidade_origem or ""),
                    "destino_estado": tr.destino_estado.sigla if tr.destino_estado else "",
                    "destino_cidade": str(tr.destino_cidade_id or tr.cidade_destino or ""),
                    "saida_data": tr.saida_data.isoformat() if tr.saida_data else "",
                    "saida_hora": tr.saida_hora.strftime("%H:%M") if tr.saida_hora else "",
                    "chegada_data": tr.chegada_data.isoformat() if tr.chegada_data else "",
                    "chegada_hora": tr.chegada_hora.strftime("%H:%M") if tr.chegada_hora else "",
                }
            )
    return result


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


def _resolve_termo_destinos_labels(destinos_data) -> list[str]:
    labels: list[str] = []
    seen: set[str] = set()
    for destino in destinos_data or []:
        uf = (destino.get("uf") or "").strip().upper()
        cidade = (destino.get("cidade") or "").strip()
        estado_obj = _resolve_estado(uf)
        cidade_obj = _resolve_cidade(cidade, estado=estado_obj)
        label = _format_trecho_local(cidade_obj, estado_obj)
        if not label:
            continue
        key = label.casefold()
        if key in seen:
            continue
        seen.add(key)
        labels.append(label)
    return labels


def _sanitize_termo_destino_label(value: str) -> str:
    cleaned = (value or "").replace("/", " - ")
    cleaned = re.sub(r'[\\:*?"<>|]+', " ", cleaned)
    cleaned = re.sub(r"\s+", " ", cleaned).strip()
    return cleaned


def _build_termo_nome(destinos_labels: list[str]) -> str:
    destino = _sanitize_termo_destino_label(destinos_labels[0] if destinos_labels else "")
    if not destino:
        destino = "sem destino"
    return f"termo de autorização {destino}"


def _format_periodo_termo(data_inicio: date | None, data_fim: date | None, data_unica: bool) -> str:
    if not data_inicio:
        return "-"
    if data_unica or not data_fim or data_fim == data_inicio:
        return _format_date(data_inicio)
    return f"{_format_date(data_inicio)} a {_format_date(data_fim)}"


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
                "origem_estado": str(current["uf"] or ""),
                "origem_cidade": str(current["cidade"] or ""),
                "destino_estado": str(destino_estado or ""),
                "destino_cidade": str(destino_cidade or ""),
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


def _infer_tipo_destino_from_serialized_trechos(
    trechos_data: list[dict[str, str | int]],
) -> str:
    temp_trechos: list[Trecho] = []
    for trecho in trechos_data or []:
        temp_trechos.append(
            Trecho(
                destino_estado=_resolve_estado(trecho.get("destino_estado")),
                destino_cidade=_resolve_cidade(trecho.get("destino_cidade")),
            )
        )
    if not temp_trechos:
        return ""
    return infer_tipo_destino(temp_trechos)


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
    trechos_payload = _oficio_step3_trechos_from_db(oficio)

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
        "oficio": oficio.numero_formatado or oficio.oficio,
        "protocolo": oficio.protocolo,
        "assunto": oficio.assunto,
        "roteiro_id": str(oficio.roteiro_id or ""),
        "viajantes_ids": viajantes_ids,
        "placa": oficio.placa,
        "modelo": oficio.modelo,
        "combustivel": oficio.combustivel,
        "tipo_viatura": oficio.tipo_viatura,
        "motorista_id": str(oficio.motorista_viajante_id or ""),
        "motorista_nome": motorista_nome,
        "motorista_oficio": oficio.motorista_oficio_formatado or oficio.motorista_oficio,
        "motorista_oficio_numero": str(oficio.motorista_oficio_numero or ""),
        "motorista_oficio_ano": str(oficio.motorista_oficio_ano or ""),
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
    tipo_destino_val = (wizard_data.get("tipo_destino") or "").strip().upper()
    if not tipo_destino_val:
        tipo_destino_val = _infer_tipo_destino_from_serialized_trechos(
            wizard_data.get("trechos") or []
        )
    custeio_code = _normalize_custeio_choice(wizard_data.get("custeio_tipo"))
    try:
        custeio_label = Oficio.CusteioTipoChoices(custeio_code).label
    except ValueError:
        custeio_label = Oficio.CusteioTipoChoices.UNIDADE.label
    nome_instituicao_custeio = (wizard_data.get("nome_instituicao_custeio") or "").strip()
    if custeio_code == Oficio.CusteioTipoChoices.OUTRA_INSTITUICAO and nome_instituicao_custeio:
        custeio_label = f"{custeio_label} ? {nome_instituicao_custeio}"
    _, _, motorista_oficio_fmt = _parse_oficio_parts(
        wizard_data.get("motorista_oficio", ""),
        numero_raw=wizard_data.get("motorista_oficio_numero", ""),
        ano_raw=wizard_data.get("motorista_oficio_ano", ""),
        default_year=timezone.localdate().year
        if wizard_data.get("motorista_carona")
        else None,
    )
    try:
        destino_display = Oficio.DestinoChoices(destino_code).label
    except ValueError:
        destino_display = Oficio.DestinoChoices.GAB.label
    return {
        "oficio": wizard_data.get("oficio", ""),
        "protocolo": format_protocolo_num(wizard_data.get("protocolo", "")),
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
        "motorista_oficio": motorista_oficio_fmt,
        "motorista_protocolo": format_protocolo_num(
            wizard_data.get("motorista_protocolo", "")
        ),
        "tipo_destino": tipo_destino_val,
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


def _oficio_display_label(oficio: Oficio | None) -> str:
    if not oficio:
        return "(rascunho)"
    return oficio.numero_formatado or "(rascunho)"


def _validate_edit_wizard_data(draft: dict) -> dict[str, str]:
    erros: dict[str, str] = {}
    protocolo_val = (draft.get("protocolo") or "").strip()
    viajantes_ids = [str(item) for item in draft.get("viajantes_ids", []) if str(item)]
    placa_val = (draft.get("placa") or "").strip()
    modelo_val = (draft.get("modelo") or "").strip()
    combustivel_val = (draft.get("combustivel") or "").strip()

    if not protocolo_val:
        erros["protocolo"] = "Informe o protocolo."
    if not viajantes_ids:
        erros["viajantes"] = "Selecione ao menos um viajante."
    if not placa_val or not modelo_val or not combustivel_val:
        erros["veiculo"] = "Preencha placa, modelo e combustivel."

    motorista_carona = bool(draft.get("motorista_carona"))
    if motorista_carona:
        (
            motorista_oficio_numero_int,
            motorista_oficio_ano_int,
            motorista_oficio_fmt,
        ) = _parse_oficio_parts(
            draft.get("motorista_oficio", ""),
            numero_raw=draft.get("motorista_oficio_numero", ""),
            ano_raw=draft.get("motorista_oficio_ano", ""),
            default_year=timezone.localdate().year,
        )
        motorista_oficio_numero = str(motorista_oficio_numero_int or "")
        motorista_oficio_ano = str(motorista_oficio_ano_int or "")
        motorista_protocolo = (draft.get("motorista_protocolo") or "").strip()
        if not motorista_oficio_numero:
            erros["motorista_oficio"] = "Informe o numero do oficio do motorista."
        if not motorista_oficio_ano and motorista_oficio_numero:
            erros["motorista_oficio"] = "Informe o ano do oficio do motorista."
        if not motorista_protocolo:
            erros["motorista_protocolo"] = "Informe o protocolo do motorista."
        draft["motorista_oficio_numero"] = motorista_oficio_numero
        draft["motorista_oficio_ano"] = motorista_oficio_ano
        draft["motorista_oficio"] = motorista_oficio_fmt

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

    tipo_destino = _infer_tipo_destino_from_serialized_trechos(trechos_data)
    if tipo_destino:
        draft["tipo_destino"] = tipo_destino

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

    temp_trechos: list[Trecho] = []
    for trecho in trechos_data:
        destino_cidade_tmp = _resolve_cidade(trecho.get("destino_cidade"))
        temp_trechos.append(Trecho(destino_cidade=destino_cidade_tmp))
    tipo_destino_final = infer_tipo_destino(temp_trechos) if temp_trechos else tipo_destino

    try:
        resultado_diarias = _calculate_periodized_diarias_for_serialized_trechos(
            trechos_data,
            retorno_chegada_data,
            retorno_chegada_hora,
            quantidade_servidores=len(viajantes),
        )
        quantidade_diarias, valor_diarias, valor_diarias_extenso = _diarias_totais_fields(
            resultado_diarias
        )
    except ValueError:
        quantidade_diarias = wizard_data.get("quantidade_diarias", "")
        valor_diarias = wizard_data.get("valor_diarias", "")
        valor_diarias_extenso = wizard_data.get("valor_diarias_extenso", "")

    retorno_saida_cidade = _format_trecho_local(destino_cidade, destino_estado)
    retorno_chegada_cidade = _format_trecho_local(sede_cidade, sede_estado)

    destino_code = _calcular_destino_automatico_from_trechos(trechos_data)
    oficio_numero, oficio_ano, oficio_formatado = _parse_oficio_parts(
        wizard_data.get("oficio", ""),
    )
    (
        motorista_oficio_numero,
        motorista_oficio_ano,
        motorista_oficio_formatado,
    ) = _parse_oficio_parts(
        wizard_data.get("motorista_oficio", ""),
        numero_raw=wizard_data.get("motorista_oficio_numero", ""),
        ano_raw=wizard_data.get("motorista_oficio_ano", ""),
        default_year=timezone.localdate().year if motorista_carona else None,
    )

    with transaction.atomic():
        oficio_obj = Oficio.objects.create(
            oficio=oficio_formatado,
            numero=oficio_numero,
            ano=oficio_ano,
            protocolo=wizard_data.get("protocolo", ""),
            roteiro=_resolve_roteiro_by_id(wizard_data.get("roteiro_id")),
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
            valor_diarias_extenso=valor_diarias_extenso,
            placa=placa_norm or placa,
            modelo=modelo,
            combustivel=combustivel,
            motorista=motorista_nome,
            motorista_oficio=motorista_oficio_formatado,
            motorista_oficio_numero=motorista_oficio_numero,
            motorista_oficio_ano=motorista_oficio_ano,
            motorista_protocolo=wizard_data.get("motorista_protocolo", ""),
            motorista_carona=motorista_carona,
            motorista_viajante=motorista_obj,
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
    if not (oficio.numero_formatado or oficio.oficio.strip()):
        erros["oficio"] = "Informe o numero do oficio."
    if not oficio.protocolo.strip():
        erros["protocolo"] = "Informe o protocolo."
    if oficio.viajantes.count() == 0:
        erros["viajantes"] = "Selecione ao menos um viajante."
    if not oficio.placa or not oficio.modelo or not oficio.combustivel:
        erros["veiculo"] = "Preencha placa, modelo e combustivel."

    if oficio.motorista_carona:
        if not (oficio.motorista_oficio_numero or oficio.motorista_oficio.strip()):
            erros["motorista_oficio"] = "Informe o numero do oficio do motorista."
        if not oficio.motorista_oficio_ano:
            erros["motorista_oficio"] = "Informe o ano do oficio do motorista."
        if not oficio.motorista_protocolo.strip():
            erros["motorista_protocolo"] = "Informe o protocolo do motorista."

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
        erros["tipo_destino"] = "Nao foi possivel determinar o tipo de destino pelo roteiro."

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

    oficio.tipo_destino = infer_tipo_destino(trechos)
    try:
        resultado_diarias = _calculate_periodized_diarias_for_trechos(
            trechos,
            oficio.retorno_chegada_data,
            oficio.retorno_chegada_hora,
            quantidade_servidores=oficio.viajantes.count(),
        )
        quantidade_diarias, valor_diarias, valor_diarias_extenso = _diarias_totais_fields(
            resultado_diarias
        )
    except ValueError:
        quantidade_diarias = oficio.quantidade_diarias or ""
        valor_diarias = oficio.valor_diarias or ""
        valor_diarias_extenso = oficio.valor_diarias_extenso or ""

    oficio.estado_sede = sede_estado
    oficio.cidade_sede = sede_cidade
    oficio.estado_destino = destino_estado
    oficio.cidade_destino = destino_cidade
    oficio.retorno_saida_cidade = _format_trecho_local(destino_cidade, destino_estado)
    oficio.retorno_chegada_cidade = _format_trecho_local(sede_cidade, sede_estado)
    oficio.quantidade_diarias = quantidade_diarias
    oficio.valor_diarias = valor_diarias
    oficio.valor_diarias_extenso = valor_diarias_extenso
    oficio.status = Oficio.Status.FINAL
    assunto_payload = build_assunto(oficio, trechos)
    oficio.assunto = assunto_payload["assunto"]
    oficio.save()

    return oficio, list(oficio.viajantes.all())









@login_required
@require_GET
def oficio_motorista_referencia(request):
    numero = normalize_oficio_num(request.GET.get("numero_oficio_ref"))
    protocolo = normalize_protocolo_num(request.GET.get("protocolo_ref"))
    if not numero or not protocolo:
        return JsonResponse({"error": "Informe numero e protocolo."}, status=400)

    numero_int, ano_int, _ = _parse_oficio_parts(numero)
    oficio_qs = Oficio.objects.select_related("motorista_viajante", "veiculo").filter(
        protocolo__iexact=protocolo
    )
    if numero_int and ano_int:
        oficio_qs = oficio_qs.filter(numero=numero_int, ano=ano_int)
    else:
        oficio_qs = oficio_qs.filter(oficio__iexact=numero)
    oficio = oficio_qs.first()
    if not oficio:
        return JsonResponse({"error": "Oficio nao encontrado."}, status=404)

    motorista_nome = (
        (oficio.motorista_viajante.nome if oficio.motorista_viajante else None)
        or oficio.motorista
        or ""
    )
    veiculo_obj = getattr(oficio, "veiculo", None)
    placa = oficio.placa or (veiculo_obj.placa if veiculo_obj else "") or ""
    viatura = oficio.modelo or (veiculo_obj.modelo if veiculo_obj else "") or ""
    combustivel = oficio.combustivel or (veiculo_obj.combustivel if veiculo_obj else "") or ""
    tipo_viatura = oficio.tipo_viatura or (veiculo_obj.tipo_viatura if veiculo_obj else "") or ""
    meio_transporte = tipo_viatura

    return JsonResponse(
        {
            "motorista_nome": motorista_nome,
            "placa": placa,
            "viatura": viatura,
            "combustivel": combustivel,
            "tipo_viatura": tipo_viatura,
            "meio_transporte": meio_transporte,
            "motorista_oficio": oficio.motorista_oficio_formatado or oficio.motorista_oficio or "",
            "motorista_protocolo": format_protocolo_num(oficio.motorista_protocolo or ""),
        }
    )


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


def _resolve_roteiro_by_id(value: str | int | None) -> Roteiro | None:
    if value is None or value == "":
        return None
    raw = str(value).strip()
    if not raw or not raw.isdigit():
        return None
    return Roteiro.objects.filter(id=int(raw), ativo=True).first()

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


def _normalize_local_reference(
    *,
    cidade_id=None,
    cidade_obj=None,
    cidade_nome: str = "",
    estado_obj=None,
    estado_sigla: str = "",
) -> tuple[str, str]:
    state_value = (
        getattr(estado_obj, "sigla", None)
        or estado_sigla
        or ""
    )
    state_token = str(state_value).strip().upper()

    if cidade_id:
        return state_token, f"id:{cidade_id}"

    if getattr(cidade_obj, "pk", None):
        return state_token, f"id:{cidade_obj.pk}"

    city_value = str(cidade_nome or cidade_obj or "").strip()
    return state_token, city_value.casefold()


def _build_oficio_trecho_signature(trecho: Trecho) -> tuple:
    origem = _normalize_local_reference(
        cidade_id=trecho.origem_cidade_id,
        cidade_obj=trecho.origem_cidade,
        estado_obj=trecho.origem_estado,
    )
    destino = _normalize_local_reference(
        cidade_id=trecho.destino_cidade_id,
        cidade_obj=trecho.destino_cidade,
        estado_obj=trecho.destino_estado,
    )
    return (
        origem,
        destino,
        trecho.saida_data,
        trecho.saida_hora,
        trecho.chegada_data,
        trecho.chegada_hora,
    )


def _build_roteiro_trecho_signature(trecho: TrechoRoteiro) -> tuple:
    origem = _normalize_local_reference(
        cidade_id=trecho.origem_cidade_id,
        cidade_obj=trecho.origem_cidade,
        cidade_nome=trecho.origem_cidade_nome,
        estado_obj=trecho.origem_estado,
        estado_sigla=trecho.origem_estado_sigla,
    )
    destino = _normalize_local_reference(
        cidade_id=trecho.destino_cidade_id,
        cidade_obj=trecho.destino_cidade,
        cidade_nome=trecho.destino_cidade_nome,
        estado_obj=trecho.destino_estado,
        estado_sigla=trecho.destino_estado_sigla,
    )
    return (
        origem,
        destino,
        trecho.saida_data,
        trecho.saida_hora,
        trecho.chegada_data,
        trecho.chegada_hora,
    )


def _build_retorno_signature(
    saida_data,
    saida_hora,
    chegada_data,
    chegada_hora,
) -> tuple:
    return (
        saida_data,
        saida_hora,
        chegada_data,
        chegada_hora,
    )


def _trechos_foram_modificados(
    trechos_oficio: list[Trecho],
    roteiro_original: Roteiro,
    oficio_obj: Oficio | None = None,
) -> bool:
    trechos_originais = list(
        roteiro_original.trechos.select_related(
            "origem_estado",
            "origem_cidade",
            "destino_estado",
            "destino_cidade",
        ).order_by("ordem", "id")
    )
    if len(trechos_oficio) != len(trechos_originais):
        return True

    for trecho_oficio, trecho_original in zip(trechos_oficio, trechos_originais):
        if _build_oficio_trecho_signature(trecho_oficio) != _build_roteiro_trecho_signature(
            trecho_original
        ):
            return True

    if oficio_obj is None:
        return False

    return _build_retorno_signature(
        oficio_obj.retorno_saida_data,
        oficio_obj.retorno_saida_hora,
        oficio_obj.retorno_chegada_data,
        oficio_obj.retorno_chegada_hora,
    ) != _build_retorno_signature(
        roteiro_original.retorno_saida_data,
        roteiro_original.retorno_saida_hora,
        roteiro_original.retorno_chegada_data,
        roteiro_original.retorno_chegada_hora,
    )


def _duration_minutes_from_bounds(
    saida_data,
    saida_hora,
    chegada_data,
    chegada_hora,
) -> int | None:
    inicio = _combine_date_time(saida_data, saida_hora)
    fim = _combine_date_time(chegada_data, chegada_hora)
    if not inicio or not fim or fim < inicio:
        return None
    return int((fim - inicio).total_seconds() // 60)


def _clonar_roteiro_com_novos_dados(
    roteiro_original: Roteiro,
    trechos_oficio: list[Trecho],
    oficio_obj: Oficio,
) -> Roteiro:
    trechos_oficio = list(trechos_oficio)
    if not trechos_oficio:
        raise ValueError("Nao ha trechos para clonar o roteiro.")

    primeiro = trechos_oficio[0]
    ultimo = trechos_oficio[-1]
    estado_sede = roteiro_original.estado_sede or primeiro.origem_estado
    cidade_sede = roteiro_original.cidade_sede or primeiro.origem_cidade
    rota_origem_sigla = (
        estado_sede.sigla if estado_sede else (primeiro.origem_estado.sigla if primeiro.origem_estado else "")
    )
    rota_origem_nome = (
        cidade_sede.nome if cidade_sede else (primeiro.origem_cidade.nome if primeiro.origem_cidade else "")
    )
    rota_destino_sigla = ultimo.destino_estado.sigla if ultimo.destino_estado else ""
    rota_destino_nome = ultimo.destino_cidade.nome if ultimo.destino_cidade else ""

    novo = Roteiro.objects.create(
        nome="(gerando...)",
        descricao=roteiro_original.descricao,
        estado_sede=estado_sede,
        cidade_sede=cidade_sede,
        uf_origem=rota_origem_sigla or roteiro_original.uf_origem or "PR",
        cidade_origem=rota_origem_nome or roteiro_original.cidade_origem or "",
        uf_destino=rota_destino_sigla or roteiro_original.uf_destino or "",
        cidade_destino=rota_destino_nome or roteiro_original.cidade_destino or "",
        retorno_saida_cidade=oficio_obj.retorno_saida_cidade or _format_trecho_local(
            ultimo.destino_cidade,
            ultimo.destino_estado,
        ),
        retorno_saida_data=oficio_obj.retorno_saida_data,
        retorno_saida_hora=oficio_obj.retorno_saida_hora,
        retorno_chegada_cidade=oficio_obj.retorno_chegada_cidade or _format_trecho_local(
            cidade_sede,
            estado_sede,
        ),
        retorno_chegada_data=oficio_obj.retorno_chegada_data,
        retorno_chegada_hora=oficio_obj.retorno_chegada_hora,
        tempo_viagem=roteiro_original.tempo_viagem,
        criado_automaticamente=True,
        distancia_km=roteiro_original.distancia_km,
        tipo_deslocamento=roteiro_original.tipo_deslocamento
        or Roteiro.TipoDeslocamentoChoices.INTERIOR,
        ativo=True,
    )

    trechos_originais = list(roteiro_original.trechos.order_by("ordem", "id"))
    ultimo_trecho_novo = None
    for index, trecho_oficio in enumerate(trechos_oficio, start=1):
        trecho_original = (
            trechos_originais[index - 1] if index - 1 < len(trechos_originais) else None
        )
        tempo_viagem_minutos = _duration_minutes_from_bounds(
            trecho_oficio.saida_data,
            trecho_oficio.saida_hora,
            trecho_oficio.chegada_data,
            trecho_oficio.chegada_hora,
        )
        if tempo_viagem_minutos is None and trecho_original:
            tempo_viagem_minutos = trecho_original.tempo_viagem_minutos

        ultimo_trecho_novo = TrechoRoteiro.objects.create(
            roteiro=novo,
            ordem=index,
            origem_estado=trecho_oficio.origem_estado,
            origem_cidade=trecho_oficio.origem_cidade,
            destino_estado=trecho_oficio.destino_estado,
            destino_cidade=trecho_oficio.destino_cidade,
            saida_data=trecho_oficio.saida_data,
            saida_hora=trecho_oficio.saida_hora,
            chegada_data=trecho_oficio.chegada_data,
            chegada_hora=trecho_oficio.chegada_hora,
            tempo_viagem_minutos=tempo_viagem_minutos,
            distancia_km=trecho_original.distancia_km if trecho_original else None,
            modal=trecho_original.modal
            if trecho_original
            else TrechoRoteiro.ModalChoices.VEICULO_PROPRIO,
            observacao=trecho_original.observacao if trecho_original else "",
        )

    if ultimo_trecho_novo:
        ultimo_trecho_novo.retorno_saida_data = novo.retorno_saida_data
        ultimo_trecho_novo.retorno_saida_hora = novo.retorno_saida_hora
        ultimo_trecho_novo.retorno_chegada_data = novo.retorno_chegada_data
        ultimo_trecho_novo.retorno_chegada_hora = novo.retorno_chegada_hora
        ultimo_trecho_novo.save(
            update_fields=[
                "retorno_saida_data",
                "retorno_saida_hora",
                "retorno_chegada_data",
                "retorno_chegada_hora",
            ]
        )

    novo.nome = novo.gerar_nome()
    novo.save(update_fields=["nome"])
    return novo


def _clonar_roteiro_com_trechos(
    original: Roteiro,
    trechos_oficio: list[Trecho],
) -> Roteiro:
    trechos_oficio = list(trechos_oficio)
    if not trechos_oficio:
        raise ValueError("Nao ha trechos para clonar o roteiro.")

    primeiro = trechos_oficio[0]
    ultimo = trechos_oficio[-1]
    oficio_shadow = Oficio(
        retorno_saida_cidade=_format_trecho_local(
            ultimo.destino_cidade,
            ultimo.destino_estado,
        ),
        retorno_saida_data=original.retorno_saida_data,
        retorno_saida_hora=original.retorno_saida_hora,
        retorno_chegada_cidade=_format_trecho_local(
            original.cidade_sede or primeiro.origem_cidade,
            original.estado_sede or primeiro.origem_estado,
        ),
        retorno_chegada_data=original.retorno_chegada_data,
        retorno_chegada_hora=original.retorno_chegada_hora,
    )
    return _clonar_roteiro_com_novos_dados(
        original,
        trechos_oficio,
        oficio_shadow,
    )


def _persist_modified_roteiro_snapshot(
    oficio_obj: Oficio,
    roteiro_origem_id,
) -> Roteiro | None:
    roteiro_origem = _resolve_roteiro_by_id(roteiro_origem_id)
    if not roteiro_origem:
        return None

    trechos_oficio = list(
        oficio_obj.trechos.select_related(
            "origem_estado",
            "origem_cidade",
            "destino_estado",
            "destino_cidade",
        ).order_by("ordem", "id")
    )
    if not trechos_oficio:
        return None

    if not _trechos_foram_modificados(
        trechos_oficio,
        roteiro_origem,
        oficio_obj=oficio_obj,
    ):
        return roteiro_origem

    novo_roteiro = _clonar_roteiro_com_novos_dados(
        roteiro_origem,
        trechos_oficio,
        oficio_obj,
    )
    oficio_obj.roteiro = novo_roteiro
    oficio_obj.save(update_fields=["roteiro"])
    return novo_roteiro


def _apply_step1_to_oficio(oficio: Oficio, payload: dict) -> None:
    numeracao_form = OficioNumeracaoForm(
        {
            "oficio": payload.get("oficio", ""),
            "protocolo": payload.get("protocolo", ""),
        }
    )
    oficio_numero = None
    oficio_ano = None
    oficio_formatado = ""
    if numeracao_form.is_valid():
        oficio_numero, oficio_ano, oficio_formatado = _parse_oficio_parts(
            numeracao_form.cleaned_data["oficio"]
        )
        oficio.protocolo = numeracao_form.cleaned_data["protocolo"]
    else:
        oficio_numero, oficio_ano, oficio_formatado = _parse_oficio_parts(
            payload.get("oficio", "")
        )
        oficio.protocolo = _formatar_protocolo(payload.get("protocolo", ""))
    if oficio.pk and oficio_numero is None:
        oficio_numero = oficio.numero
    if oficio.pk and oficio_ano is None:
        oficio_ano = oficio.ano
    if not oficio_formatado and oficio_numero and oficio_ano:
        oficio_formatado = format_oficio_num(oficio_numero, oficio_ano)
    oficio.numero = oficio_numero
    oficio.ano = oficio_ano
    oficio.oficio = oficio_formatado
    oficio.assunto = payload.get("assunto", "").strip()
    oficio.motivo = payload.get("motivo", "").strip()
    oficio.custeio_tipo = _normalize_custeio_choice(payload.get("custeio_tipo") or payload.get("custos"))
    oficio.nome_instituicao_custeio = _normalize_nome_instituicao_custeio(
        oficio.custeio_tipo, payload.get("nome_instituicao_custeio")
    )
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

    motorista_fields_form = MotoristaTransporteForm(
        {
            "motorista_nome": payload.get("motorista_nome", ""),
            "motorista_oficio": payload.get("motorista_oficio", ""),
            "motorista_oficio_numero": payload.get("motorista_oficio_numero", ""),
            "motorista_oficio_ano": payload.get("motorista_oficio_ano", ""),
            "motorista_protocolo": payload.get("motorista_protocolo", ""),
        }
    )
    motorista_fields_form.is_valid()
    motorista_id = payload.get("motorista_id") or ""
    motorista_nome = motorista_fields_form.cleaned_data.get(
        "motorista_nome",
        normalize_upper_text(payload.get("motorista_nome", "")),
    )
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
    (
        motorista_oficio_numero,
        motorista_oficio_ano,
        motorista_oficio_formatado,
    ) = _parse_oficio_parts(
        motorista_fields_form.cleaned_data.get("motorista_oficio", ""),
        numero_raw=motorista_fields_form.cleaned_data.get("motorista_oficio_numero", ""),
        ano_raw=motorista_fields_form.cleaned_data.get("motorista_oficio_ano", ""),
        default_year=timezone.localdate().year if motorista_carona else None,
    )
    oficio.motorista_oficio = motorista_oficio_formatado
    oficio.motorista_oficio_numero = motorista_oficio_numero
    oficio.motorista_oficio_ano = motorista_oficio_ano
    oficio.motorista_protocolo = motorista_fields_form.cleaned_data.get(
        "motorista_protocolo",
        _formatar_protocolo(payload.get("motorista_protocolo", "")),
    )
    oficio.motorista_carona = motorista_carona
    oficio.motorista_viajante = motorista_obj
    if motorista_carona:
        oficio.carona_oficio_referencia = _resolve_oficio_by_id(
            payload.get("carona_oficio_referencia_id")
        )
    else:
        oficio.carona_oficio_referencia = None
    oficio.veiculo = veiculo
    if veiculo and veiculo.tipo_viatura:
        oficio.tipo_viatura = veiculo.tipo_viatura
    elif payload.get("tipo_viatura"):
        oficio.tipo_viatura = payload.get("tipo_viatura")
    oficio.save()


def _apply_step3_to_oficio(oficio: Oficio, payload: dict) -> None:
    oficio.motivo = payload.get("motivo", "").strip()
    oficio.tipo_destino = (payload.get("tipo_destino") or "").strip().upper()
    oficio.roteiro = _resolve_roteiro_by_id(payload.get("roteiro_id"))
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

    retorno_chegada_data = oficio.retorno_chegada_data
    retorno_chegada_hora = oficio.retorno_chegada_hora
    if trechos_instances and retorno_chegada_data:
        try:
            resultado = _calculate_periodized_diarias_for_trechos(
                trechos_instances,
                retorno_chegada_data,
                retorno_chegada_hora,
                quantidade_servidores=oficio.viajantes.count(),
            )
        except ValueError:
            resultado = None
        if resultado:
            quantidade_diarias, valor_diarias, valor_diarias_extenso = _diarias_totais_fields(
                resultado
            )
            oficio.quantidade_diarias = quantidade_diarias
            oficio.valor_diarias = valor_diarias
            oficio.valor_diarias_extenso = valor_diarias_extenso

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


def _efetivo_resumo_context() -> dict[str, object]:
    efetivos = list(
        Efetivo.objects.select_related("cargo")
        .filter(cargo__ativo=True)
        .order_by("cargo__ordem", "cargo__nome")
    )
    total_coordenadores = sum(
        int(item.quantidade or 0) for item in efetivos if bool(item.cargo.is_coordenador)
    )
    total_servidores = sum(
        int(item.quantidade or 0) for item in efetivos if not bool(item.cargo.is_coordenador)
    )
    total_geral = total_servidores + total_coordenadores
    return {
        "efetivos": efetivos,
        "total_servidores": total_servidores,
        "total_coordenadores": total_coordenadores,
        "total_geral": total_geral,
        "resumo_servidores": format_total_servidores(total_servidores),
        "resumo_coordenadores": (
            f"{total_coordenadores} coordenador"
            if total_coordenadores == 1
            else f"{total_coordenadores} coordenadores"
        ),
        "resumo_total": f"{total_geral} pessoa" if total_geral == 1 else f"{total_geral} pessoas",
    }


@require_http_methods(["GET", "POST"])
def efetivo(request):
    for cargo in Cargo.objects.filter(ativo=True).order_by("ordem", "nome"):
        Efetivo.objects.get_or_create(cargo=cargo)

    resumo = _efetivo_resumo_context()
    EfetivoFormSet = formset_factory(EfetivoLinhaForm, extra=0)
    initial = [
        {
            "cargo_id": item.cargo_id,
            "cargo_nome": item.cargo.nome,
            "is_coordenador": "True" if item.cargo.is_coordenador else "",
            "quantidade": int(item.quantidade or 0),
        }
        for item in resumo["efetivos"]
    ]

    if request.method == "POST":
        formset = EfetivoFormSet(request.POST, prefix="efetivo")
        cargo_form = NovoCargoForm(request.POST, prefix="novo")
        if formset.is_valid() and cargo_form.is_valid():
            for form in formset:
                cargo_id = int(form.cleaned_data.get("cargo_id") or 0)
                quantidade = int(form.cleaned_data.get("quantidade") or 0)
                if cargo_id:
                    Efetivo.objects.filter(cargo_id=cargo_id).update(quantidade=quantidade)

            nome_cargo = cargo_form.cleaned_data.get("nome_cargo", "")
            if nome_cargo:
                cargo, _created = Cargo.objects.get_or_create(
                    nome=nome_cargo,
                    defaults={
                        "ordem": (Cargo.objects.aggregate(max_ordem=models.Max("ordem")).get("max_ordem") or 0)
                        + 1,
                        "ativo": True,
                        "is_coordenador": bool(cargo_form.cleaned_data.get("is_coordenador")),
                    },
                )
                desired_is_coordenador = bool(cargo_form.cleaned_data.get("is_coordenador"))
                if not cargo.ativo or cargo.is_coordenador != desired_is_coordenador:
                    cargo.ativo = True
                    cargo.is_coordenador = desired_is_coordenador
                    cargo.save(update_fields=["ativo", "is_coordenador"])
                Efetivo.objects.get_or_create(cargo=cargo)
            return redirect("efetivo")
    else:
        formset = EfetivoFormSet(initial=initial, prefix="efetivo")
        cargo_form = NovoCargoForm(prefix="novo")

    resumo = _efetivo_resumo_context()
    linhas = []
    for item, form in zip(resumo["efetivos"], formset.forms):
        linhas.append({"cargo": item.cargo, "form": form})

    return render(
        request,
        "viagens/legacy_efetivo.html",
        {
            "formset": formset,
            "cargo_form": cargo_form,
            "linhas": linhas,
            **resumo,
        },
    )


@require_http_methods(["GET", "POST"])
def diarias(request):
    resultado = None
    if request.method == "POST":
        form = DiariasSimplesForm(request.POST)
        if form.is_valid():
            dias, total = calculate_simple_diarias_total(
                form.cleaned_data["data_saida"],
                form.cleaned_data["data_retorno"],
                form.cleaned_data["valor_diaria"],
                meia_diaria=bool(form.cleaned_data.get("meia_diaria")),
            )
            resultado = {
                "dias": dias,
                "total": total,
                "total_formatado": f"{total:.2f}".replace(".", ","),
            }
    else:
        form = DiariasSimplesForm()

    return render(
        request,
        "viagens/legacy_diarias.html",
        {
            "form": form,
            "resultado": resultado,
        },
    )


@require_GET
def simulacao_diarias(request):
    simulacao_last = dict(request.session.get("simulacao_diarias_last") or {})
    periods = simulacao_last.get("periods")
    if not isinstance(periods, list):
        periods = []
    if not periods:
        periods = [
            {
                "tipo": "INTERIOR",
                "start_date": "",
                "start_time": "",
                "end_date": "",
                "end_time": "",
            }
        ]
    simulacao_last["periods"] = periods
    return render(
        request,
        "viagens/simulacao_diarias.html",
        {
            "simulacao_last": simulacao_last,
            "simulacao_initial_periods_json": json.dumps(periods),
            "simulacao_calcular_url": reverse("simulacao_diarias_calcular"),
        },
    )


@require_http_methods(["POST"])
def simulacao_diarias_calcular(request):
    payload_json: dict | None = None
    if request.content_type and "application/json" in request.content_type:
        try:
            payload_json = json.loads(request.body.decode("utf-8"))
        except (TypeError, ValueError, json.JSONDecodeError):
            payload_json = None
    try:
        quantidade_servidores = int(
            (
                (payload_json or {}).get("quantidade_servidores")
                if payload_json is not None
                else request.POST.get("quantidade_servidores", "1")
            )
            or "1"
        )
    except (TypeError, ValueError):
        quantidade_servidores = 1
    quantidade_servidores = max(1, quantidade_servidores)

    periods_payload: list[dict]
    if payload_json is not None:
        periods_payload = payload_json.get("periods") or []
    else:
        periods_raw = (request.POST.get("periods_payload") or request.POST.get("periods") or "").strip()
        try:
            periods_payload = json.loads(periods_raw) if periods_raw else []
        except json.JSONDecodeError:
            periods_payload = []
    if not isinstance(periods_payload, list):
        periods_payload = []

    try:
        resultado = calculate_diarias_from_periods_payload(
            periods_payload=periods_payload,
            total_servidores=quantidade_servidores,
        )
    except ValueError as exc:
        return JsonResponse({"error": str(exc)}, status=400)

    request.session["simulacao_diarias_last"] = {
        "periods": periods_payload,
        "quantidade_servidores": quantidade_servidores,
        "resultado": resultado,
    }
    request.session.modified = True
    return JsonResponse(resultado)


@require_http_methods(["GET", "POST"])
def formulario(request):
    data = _get_wizard_data(request)
    erro = ""
    if request.method == "GET":
        if not request.GET.get("resume"):
            _clear_wizard_data(request)
            data = {}
            evento_id_raw = request.GET.get("evento_id")
            if evento_id_raw:
                try:
                    eid = int(evento_id_raw)
                    request.session["wizard_evento_id"] = eid
                    request.session["wizard_evento_return_url"] = reverse(
                        "evento_guiado_etapa6", kwargs={"evento_id": eid}
                    )
                    request.session.modified = True
                except (ValueError, TypeError):
                    pass
        else:
            data = _ensure_wizard_session(request)
        oficio_obj_get = _get_wizard_oficio(request, create=False)
        if oficio_obj_get:
            oficio_formatado = oficio_obj_get.numero_formatado or oficio_obj_get.oficio
            if oficio_formatado and data.get("oficio") != oficio_formatado:
                data = _update_wizard_data(request, {"oficio": oficio_formatado})
        else:
            ano_atual = timezone.localdate().year
            reserved = _ensure_reserved_oficio_for_year(request, ano_atual)
            oficio_formatado = format_oficio_num(reserved["numero"], reserved["ano"])
            if oficio_formatado and data.get("oficio") != oficio_formatado:
                data = _update_wizard_data(request, {"oficio": oficio_formatado})
    viajantes = Viajante.objects.order_by("nome")
    servidores_form = ServidoresSelectForm(
        initial={"servidores": data.get("viajantes_ids", [])}
    )
    oficio_obj = _get_wizard_oficio(request, create=False)

    if request.method == "POST":
        goto_step = (request.POST.get("goto_step") or "").strip()
        numeracao_form = OficioNumeracaoForm(
            {
                "oficio": request.POST.get("oficio", "").strip(),
                "protocolo": request.POST.get("protocolo", "").strip(),
            }
        )
        numeracao_form.is_valid()
        protocolo_val = numeracao_form.cleaned_data.get(
            "protocolo", _formatar_protocolo(request.POST.get("protocolo", "").strip())
        )
        motivo_val = request.POST.get("motivo", "").strip()
        custeio_tipo_val = _normalize_custeio_choice(request.POST.get("custeio_tipo") or request.POST.get("custos"))
        nome_instituicao_custeio = _normalize_nome_instituicao_custeio(
            custeio_tipo_val, request.POST.get("nome_instituicao_custeio", "")
        )
        servidores_form = ServidoresSelectForm(request.POST)
        if servidores_form.is_valid():
            viajantes_ids = [
                str(item.id)
                for item in servidores_form.cleaned_data.get("servidores", [])
            ]
        else:
            viajantes_ids = []

        if not oficio_obj:
            _ensure_reserved_oficio_for_year(request, timezone.localdate().year)
        oficio_obj = _get_wizard_oficio(request, create=True)
        oficio_val = oficio_obj.numero_formatado or oficio_obj.oficio
        payload = {
            "oficio": oficio_val,
            "protocolo": protocolo_val,
            "motivo": motivo_val,
            "viajantes_ids": viajantes_ids,
            "custeio_tipo": custeio_tipo_val,
            "nome_instituicao_custeio": nome_instituicao_custeio,
        }
        _apply_step1_to_oficio(oficio_obj, payload)
        payload["oficio"] = oficio_obj.numero_formatado or oficio_obj.oficio
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
            "protocolo": format_protocolo_num(data.get("protocolo", "")),
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
        motorista_fields_form = MotoristaTransporteForm(
            {
                "motorista_nome": request.POST.get("motorista_nome", ""),
                "motorista_oficio": request.POST.get("motorista_oficio", ""),
                "motorista_oficio_numero": request.POST.get("motorista_oficio_numero", ""),
                "motorista_oficio_ano": request.POST.get("motorista_oficio_ano", ""),
                "motorista_protocolo": request.POST.get("motorista_protocolo", ""),
            }
        )
        motorista_fields_form.is_valid()
        motorista_nome = motorista_fields_form.cleaned_data.get(
            "motorista_nome",
            normalize_upper_text(request.POST.get("motorista_nome", "").strip()),
        )
        motorista_oficio = motorista_fields_form.cleaned_data.get(
            "motorista_oficio",
            _formatar_oficio_numero(request.POST.get("motorista_oficio", "").strip()),
        )
        motorista_oficio_numero = motorista_fields_form.cleaned_data.get(
            "motorista_oficio_numero",
            normalize_digits(request.POST.get("motorista_oficio_numero", "")),
        )
        motorista_oficio_ano = motorista_fields_form.cleaned_data.get(
            "motorista_oficio_ano",
            normalize_digits(request.POST.get("motorista_oficio_ano", "")),
        )
        motorista_protocolo = motorista_fields_form.cleaned_data.get(
            "motorista_protocolo",
            _formatar_protocolo(request.POST.get("motorista_protocolo", "").strip()),
        )

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

        oficio_obj = _get_wizard_oficio(request, create=True)
        payload = {
            "placa": placa_norm or placa_val,
            "modelo": modelo_val,
            "combustivel": combustivel_val,
            "tipo_viatura": tipo_viatura_val,
            "motorista_id": motorista_id,
            "motorista_nome": motorista_nome,
            "motorista_oficio": motorista_oficio,
            "motorista_oficio_numero": motorista_oficio_numero,
            "motorista_oficio_ano": motorista_oficio_ano
            or (str(timezone.localdate().year) if motorista_oficio_numero else ""),
            "motorista_protocolo": motorista_protocolo,
            "motorista_carona": motorista_carona,
        }
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
            "motorista_oficio_numero": data.get("motorista_oficio_numero", ""),
            "motorista_oficio_ano": data.get("motorista_oficio_ano")
            or str(timezone.localdate().year),
            "motorista_protocolo": format_protocolo_num(
                data.get("motorista_protocolo", "")
            ),
            "motorista_carona": motorista_carona,
            "oficio": data.get("oficio", ""),
            "protocolo": format_protocolo_num(data.get("protocolo", "")),
            "motivo": data.get("motivo", ""),
            "viajantes_ids": viajantes_ids,
            "preview_viajantes": preview_viajantes,
            "motorista_preview": motorista_preview,
            "motorista_form": motorista_form,
            "status_label": status_context["status_label"],
            "status_class": status_context["status_class"],
        },
    )


@require_http_methods(["GET", "POST"])
def oficio_step3(request):
    data = _ensure_wizard_session(request)
    oficio_obj = _get_wizard_oficio(request, create=False)

    motivo_val = data.get("motivo", "")
    tipo_destino = (data.get("tipo_destino") or "").strip().upper()
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
        sede_default = _get_sede_cidade_default_id()
        if sede_default:
            defaults["sede_cidade"] = sede_default
    if defaults:
        data = _update_wizard_data(request, defaults)
    sede_uf = (data.get("sede_uf") or "").strip().upper() or "PR"
    sede_cidade = (data.get("sede_cidade") or "").strip()
    destinos_session = _normalize_destinos_for_wizard(data.get("destinos"))

    trechos_session = data.get("trechos") or []
    if not trechos_session:
        trechos_banco = _oficio_step3_trechos_from_db(oficio_obj)
        if trechos_banco:
            trechos_session = trechos_banco
            data = _update_wizard_data(request, {"trechos": trechos_session})
    if not trechos_session and request.session.get("wizard_evento_id"):
        evento_id = request.session["wizard_evento_id"]
        try:
            trechos_seed = _seed_trechos_from_evento_roteiros(int(evento_id))
            if trechos_seed:
                payload = {"trechos": trechos_seed}
                if trechos_seed and trechos_seed[0].get("origem_estado"):
                    payload["sede_uf"] = str(trechos_seed[0].get("origem_estado") or "PR")
                if trechos_seed and trechos_seed[0].get("origem_cidade"):
                    payload["sede_cidade"] = str(trechos_seed[0].get("origem_cidade") or "")
                data = _update_wizard_data(request, payload)
                trechos_session = trechos_seed
        except (ValueError, TypeError):
            pass
    if not trechos_session:
        valid_destinos = [
            destino for destino in destinos_session if destino.get("uf") and destino.get("cidade")
        ]
        trechos_session = _build_trechos_from_sede_destinos(
            sede_uf, sede_cidade, valid_destinos
        )
        _update_wizard_data(request, {"trechos": trechos_session})
    tipo_destino_inferido = _infer_tipo_destino_from_serialized_trechos(trechos_session)
    if tipo_destino_inferido and tipo_destino_inferido != tipo_destino:
        data = _update_wizard_data(request, {"tipo_destino": tipo_destino_inferido})
        tipo_destino = tipo_destino_inferido
    trechos_initial = _normalize_trechos_initial(trechos_session)
    formset_extra = max(1, len(trechos_initial))
    TrechoFormSet = _build_trecho_formset(formset_extra)

    dummy_oficio = Oficio()
    if request.method == "POST":
        goto_step = (request.POST.get("goto_step") or "").strip()
        roteiro_id = (request.POST.get("roteiro_id") or "").strip()
        roteiro_origem_id = (
            request.POST.get("roteiro_origem_id")
            or roteiro_id
            or data.get("roteiro_origem_id")
            or ""
        ).strip()
        if roteiro_id and not _resolve_roteiro_by_id(roteiro_id):
            messages.error(
                request,
                "O roteiro selecionado nao foi encontrado ou esta inativo.",
            )
            roteiro_id = ""
        if not roteiro_id:
            roteiro_origem_id = ""
        elif roteiro_origem_id and not _resolve_roteiro_by_id(roteiro_origem_id):
            roteiro_origem_id = roteiro_id
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
        tipo_destino = _infer_tipo_destino_from_serialized_trechos(trechos_serialized)
        data = _update_wizard_data(
            request,
            {
                "sede_uf": sede_uf_post,
                "sede_cidade": sede_cidade_post,
                "roteiro_id": roteiro_id,
                "roteiro_origem_id": roteiro_origem_id,
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
                "roteiro_id": roteiro_id,
                "valor_diarias_extenso": valor_diarias_extenso,
                "retorno_saida_data": retorno_saida_data,
                "retorno_saida_hora": retorno_saida_hora,
                "retorno_chegada_data": retorno_chegada_data,
                "retorno_chegada_hora": retorno_chegada_hora,
                "trechos": trechos_serialized,
            },
        )
        if roteiro_origem_id:
            roteiro_atualizado = _persist_modified_roteiro_snapshot(
                oficio_obj,
                roteiro_origem_id,
            )
            if roteiro_atualizado:
                roteiro_id = str(roteiro_atualizado.pk)
                roteiro_origem_id = str(roteiro_atualizado.pk)
                data = _update_wizard_data(
                    request,
                    {
                        "roteiro_id": roteiro_id,
                        "roteiro_origem_id": roteiro_origem_id,
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
        sede_cidade = _get_sede_cidade_default_id()
    sede_estado = _resolve_estado(sede_uf)
    sede_cidade_obj = _resolve_cidade(sede_cidade, estado=sede_estado)
    sede_label = _format_trecho_local(sede_cidade_obj, sede_estado)
    sede_uf_label = _format_estado_label(sede_estado, fallback=sede_uf)

    destinos_session = _normalize_destinos_for_wizard(data.get("destinos"))
    destinos_display = _build_destinos_display(destinos_session)
    roteiro_id = str(
        data.get("roteiro_id") or getattr(oficio_obj, "roteiro_id", "") or ""
    )
    roteiro_origem_id = str(data.get("roteiro_origem_id") or roteiro_id or "")
    roteiro_atual = _resolve_roteiro_by_id(roteiro_id)

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
            "quantidade_diarias": data.get("quantidade_diarias", ""),
            "valor_diarias": data.get("valor_diarias", ""),
            "valor_diarias_extenso": valor_diarias_extenso,
            "quantidade_servidores": len(data.get("viajantes_ids", [])),
            "estados": estados,
            "sede_uf": sede_uf,
            "sede_uf_label": sede_uf_label,
            "sede_cidade": sede_cidade,
            "sede_label": sede_label,
            "roteiro_id": roteiro_id,
            "roteiro_origem_id": roteiro_origem_id,
            "roteiro_nome": roteiro_atual.nome if roteiro_atual else "",
            "roteiro_destinos": roteiro_atual.get_destinos_display()
            if roteiro_atual
            else "",
            "destinos": destinos_display,
            "destinos_total_forms": len(destinos_display),
            "status_label": status_context["status_label"],
            "status_class": status_context["status_class"],
            "destinos_order": destinos_order,
            "calcular_diarias_url": reverse("oficio_calcular_diarias"),
            "destino_display": _destino_label_from_code(
                _calcular_destino_automatico_from_trechos(trechos_session)
            ),
        },
    )


@require_http_methods(["POST"])
def oficio_calcular_diarias(request, oficio_id: int | None = None):
    if oficio_id:
        session_data = _ensure_edit_session(request, oficio_id)
    else:
        session_data = _ensure_wizard_session(request)

    post_data = _prune_trailing_trechos_post(request.POST, "trechos")
    trechos_serialized = _serialize_trechos_from_post(post_data)
    sede_uf_post, sede_cidade_post, destinos_raw = _serialize_sede_destinos_from_post(
        post_data
    )
    valid_destinos = [
        destino for destino in destinos_raw if destino.get("uf") and destino.get("cidade")
    ]
    base_trechos = _build_trechos_from_sede_destinos(
        sede_uf_post,
        sede_cidade_post,
        valid_destinos,
    )
    if base_trechos:
        trechos_serialized = _merge_datas_horas(trechos_serialized, base_trechos)

    retorno_chegada_data_raw = (request.POST.get("retorno_chegada_data") or "").strip()
    retorno_chegada_hora_raw = (request.POST.get("retorno_chegada_hora") or "").strip()
    retorno_chegada_data = parse_date(retorno_chegada_data_raw) if retorno_chegada_data_raw else None
    retorno_chegada_hora = parse_time(retorno_chegada_hora_raw) if retorno_chegada_hora_raw else None

    quantidade_servidores = len(session_data.get("viajantes_ids", []))
    if quantidade_servidores <= 0:
        try:
            quantidade_servidores = int(request.POST.get("quantidade_servidores", "0"))
        except (TypeError, ValueError):
            quantidade_servidores = 0

    try:
        resultado = _calculate_periodized_diarias_for_serialized_trechos(
            trechos_serialized,
            retorno_chegada_data,
            retorno_chegada_hora,
            quantidade_servidores=quantidade_servidores,
        )
    except ValueError as exc:
        return JsonResponse(
            {"error": str(exc) or "Preencha datas e horas para calcular."},
            status=400,
        )

    resultado["tipo_destino"] = _infer_tipo_destino_from_serialized_trechos(trechos_serialized)
    return JsonResponse(resultado)


@require_http_methods(["GET", "POST"])
def oficio_step4(request):
    wizard_data = _ensure_wizard_session(request)
    oficio_obj = _get_wizard_oficio(request, create=False)
    erros: dict[str, str] = {}

    if request.method == "GET" and oficio_obj:
        from viagens.services.justificativa_helpers import exige_justificativa, justificativa_preenchida
        if exige_justificativa(oficio_obj) and not justificativa_preenchida(oficio_obj):
            justificativa_url = reverse("justificativa_oficio", args=[oficio_obj.id])
            return redirect(f"{justificativa_url}?next=oficio_step4")

    if request.method == "POST":
        action = (request.POST.get("action") or "").strip()
        if action == "prev":
            return redirect("oficio_step3")
        if not oficio_obj:
            return redirect("formulario")
        erros = _validate_oficio_for_finalize(oficio_obj)
        if not erros:
            oficio_obj, _ = _finalize_oficio_draft(oficio_obj)
            return_url = request.session.get("wizard_evento_return_url")
            evento_id = request.session.get("wizard_evento_id")
            # Modo evento: se exige justificativa e ainda vazia, redirecionar para justificativa com next=etapa6
            if evento_id and return_url:
                from viagens.services.justificativa_helpers import exige_justificativa, justificativa_preenchida
                if exige_justificativa(oficio_obj) and not justificativa_preenchida(oficio_obj):
                    justificativa_url = reverse("justificativa_oficio", args=[oficio_obj.id])
                    from urllib.parse import quote
                    next_etapa6 = return_url
                    return redirect(f"{justificativa_url}?next={quote(next_etapa6)}")
            _clear_wizard_data(request)
            if return_url:
                return redirect(return_url)
            return redirect("oficios_lista")

    context = _build_step4_context(wizard_data)
    status_context = _wizard_status_context(oficio_obj)
    context.update(
        {
            "erros": erros,
            "status_label": status_context["status_label"],
            "status_class": status_context["status_class"],
            "oficio_id": oficio_obj.id if oficio_obj else None,
            "wizard_evento_return_url": request.session.get("wizard_evento_return_url"),
        }
    )
    return render(request, "viagens/oficio_step4.html", context)


@require_http_methods(["GET", "POST"])
def oficio_edit_step1(request, oficio_id: int):
    data = _ensure_edit_session(request, oficio_id)
    oficio_obj = get_object_or_404(Oficio, id=oficio_id)

    if request.method == "POST":
        numeracao_form = OficioNumeracaoForm(
            {
                "oficio": request.POST.get("oficio", "").strip(),
                "protocolo": request.POST.get("protocolo", "").strip(),
            }
        )
        numeracao_form.is_valid()
        _, _, oficio_val = _parse_oficio_parts(
            numeracao_form.cleaned_data.get(
                "oficio", _formatar_oficio_numero(request.POST.get("oficio", "").strip())
            )
        )
        protocolo_val = numeracao_form.cleaned_data.get(
            "protocolo", _formatar_protocolo(request.POST.get("protocolo", "").strip())
        )
        motivo_val = request.POST.get("motivo", "").strip()
        custeio_tipo_val = _normalize_custeio_choice(request.POST.get("custeio_tipo") or request.POST.get("custos"))
        nome_instituicao_custeio = _normalize_nome_instituicao_custeio(
            custeio_tipo_val, request.POST.get("nome_instituicao_custeio", "")
        )
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
            "oficio_display": _oficio_display_label(oficio_obj),
            "oficio": data.get("oficio", ""),
            "protocolo": format_protocolo_num(data.get("protocolo", "")),
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
        motorista_fields_form = MotoristaTransporteForm(
            {
                "motorista_nome": request.POST.get("motorista_nome", ""),
                "motorista_oficio": request.POST.get("motorista_oficio", ""),
                "motorista_oficio_numero": request.POST.get("motorista_oficio_numero", ""),
                "motorista_oficio_ano": request.POST.get("motorista_oficio_ano", ""),
                "motorista_protocolo": request.POST.get("motorista_protocolo", ""),
            }
        )
        motorista_fields_form.is_valid()
        motorista_nome = motorista_fields_form.cleaned_data.get(
            "motorista_nome",
            normalize_upper_text(request.POST.get("motorista_nome", "")),
        )
        motorista_oficio = motorista_fields_form.cleaned_data.get(
            "motorista_oficio",
            _formatar_oficio_numero(request.POST.get("motorista_oficio", "")),
        )
        motorista_oficio_numero = motorista_fields_form.cleaned_data.get(
            "motorista_oficio_numero",
            normalize_digits(request.POST.get("motorista_oficio_numero", "")),
        )
        motorista_oficio_ano = motorista_fields_form.cleaned_data.get(
            "motorista_oficio_ano",
            normalize_digits(request.POST.get("motorista_oficio_ano", "")),
        )
        motorista_protocolo = motorista_fields_form.cleaned_data.get(
            "motorista_protocolo",
            _formatar_protocolo(request.POST.get("motorista_protocolo", "")),
        )
        carona_oficio_referencia_id = (
            request.POST.get("carona_oficio_referencia", "").strip()
        )

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
        if not motorista_carona:
            carona_oficio_referencia_id = ""

        logger.debug(
            "[edit-step2] before session save oficio_id=%s motorista_id=%s motorista_nome=%s",
            oficio_id,
            motorista_id,
            motorista_nome,
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
                "motorista_oficio_numero": motorista_oficio_numero,
                "motorista_oficio_ano": motorista_oficio_ano
                or (str(timezone.localdate().year) if motorista_oficio_numero else ""),
                "motorista_protocolo": motorista_protocolo,
                "carona_oficio_referencia_id": carona_oficio_referencia_id,
                "motorista_carona": motorista_carona,
                "erros": {},
            },
        )
        logger.debug(
            "[edit-step2] after session save oficio_id=%s motorista_oficio=%s motorista_protocolo=%s carona_ref=%s",
            oficio_id,
            data.get("motorista_oficio", ""),
            data.get("motorista_protocolo", ""),
            data.get("carona_oficio_referencia_id", ""),
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

    oficio_obj = get_object_or_404(Oficio, id=oficio_id)
    status_context = _wizard_status_context(oficio_obj)
    oficios_referencia = list(
        Oficio.objects.exclude(id=oficio_id).order_by("-created_at", "-id")[:100]
    )
    carona_oficio_referencia_id = _normalize_int(data.get("carona_oficio_referencia_id"))

    return render(
        request,
        "viagens/oficio_edit_step2.html",
        {
            "oficio_id": oficio_id,
            "oficio_display": _oficio_display_label(oficio_obj),
            "placa": data.get("placa", ""),
            "modelo": data.get("modelo", ""),
            "combustivel": data.get("combustivel", ""),
            "tipo_viatura": data.get("tipo_viatura", ""),
            "combustivel_choices": _get_combustivel_choices(),
            "motorista_id": data.get("motorista_id", ""),
            "motorista_nome": motorista_nome_val,
            "motorista_oficio": data.get("motorista_oficio", ""),
            "motorista_oficio_numero": data.get("motorista_oficio_numero", ""),
            "motorista_oficio_ano": data.get("motorista_oficio_ano")
            or str(timezone.localdate().year),
            "motorista_protocolo": format_protocolo_num(
                data.get("motorista_protocolo", "")
            ),
            "motorista_carona": motorista_carona,
            "oficio": data.get("oficio", ""),
            "protocolo": format_protocolo_num(data.get("protocolo", "")),
            "motivo": data.get("motivo", ""),
            "viajantes_ids": viajantes_ids,
            "preview_viajantes": preview_viajantes,
            "motorista_preview": motorista_preview,
            "motorista_form": motorista_form,
            "oficios_referencia": oficios_referencia,
            "carona_oficio_referencia_id": carona_oficio_referencia_id,
            "status_label": status_context["status_label"],
            "status_class": status_context["status_class"],
        },
    )


@require_http_methods(["GET", "POST"])
def oficio_edit_step3(request, oficio_id: int):
    data = _ensure_edit_session(request, oficio_id)
    oficio_obj = get_object_or_404(Oficio, id=oficio_id)

    if request.method == "GET":
        trechos_banco = _oficio_step3_trechos_from_db(oficio_obj)
        data = _update_edit_data(
            request,
            oficio_id,
            {
                "trechos": trechos_banco or [],
                "retorno": {
                    "retorno_saida_data": oficio_obj.retorno_saida_data.isoformat()
                    if oficio_obj.retorno_saida_data
                    else "",
                    "retorno_saida_hora": oficio_obj.retorno_saida_hora.strftime("%H:%M")
                    if oficio_obj.retorno_saida_hora
                    else "",
                    "retorno_chegada_data": oficio_obj.retorno_chegada_data.isoformat()
                    if oficio_obj.retorno_chegada_data
                    else "",
                    "retorno_chegada_hora": oficio_obj.retorno_chegada_hora.strftime("%H:%M")
                    if oficio_obj.retorno_chegada_hora
                    else "",
                },
                "retorno_saida_cidade": oficio_obj.retorno_saida_cidade or "",
                "retorno_chegada_cidade": oficio_obj.retorno_chegada_cidade or "",
                "roteiro_origem_id": str(
                    data.get("roteiro_origem_id") or oficio_obj.roteiro_id or ""
                ),
            },
        )

    motivo_val = data.get("motivo", "")
    tipo_destino = (data.get("tipo_destino") or "").strip().upper()
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
        sede_default = _get_sede_cidade_default_id()
        if sede_default:
            defaults["sede_cidade"] = sede_default
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
    tipo_destino_inferido = _infer_tipo_destino_from_serialized_trechos(trechos_session)
    if tipo_destino_inferido and tipo_destino_inferido != tipo_destino:
        data = _update_edit_data(request, oficio_id, {"tipo_destino": tipo_destino_inferido})
        tipo_destino = tipo_destino_inferido
    trechos_initial = _normalize_trechos_initial(trechos_session)
    formset_extra = max(1, len(trechos_initial))
    TrechoFormSet = _build_trecho_formset(formset_extra)

    dummy_oficio = Oficio()
    if request.method == "POST":
        roteiro_id = (request.POST.get("roteiro_id") or "").strip()
        roteiro_origem_id = (
            request.POST.get("roteiro_origem_id")
            or roteiro_id
            or data.get("roteiro_origem_id")
            or ""
        ).strip()
        if roteiro_id and not _resolve_roteiro_by_id(roteiro_id):
            messages.error(
                request,
                "O roteiro selecionado nao foi encontrado ou esta inativo.",
            )
            roteiro_id = ""
        if not roteiro_id:
            roteiro_origem_id = ""
        elif roteiro_origem_id and not _resolve_roteiro_by_id(roteiro_origem_id):
            roteiro_origem_id = roteiro_id
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
        tipo_destino = _infer_tipo_destino_from_serialized_trechos(trechos_serialized)
        data = _update_edit_data(
            request,
            oficio_id,
            {
                "sede_uf": sede_uf_post,
                "sede_cidade": sede_cidade_post,
                "roteiro_id": roteiro_id,
                "roteiro_origem_id": roteiro_origem_id,
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
            if retorno_chegada_data:
                try:
                    resultado_diarias = _calculate_periodized_diarias_for_serialized_trechos(
                        trechos_serialized,
                        retorno_chegada_data,
                        retorno_chegada_hora,
                        quantidade_servidores=len(data.get("viajantes_ids", [])),
                    )
                except ValueError:
                    resultado_diarias = None
                if resultado_diarias:
                    quantidade_diarias, valor_diarias, valor_diarias_extenso_auto = (
                        _diarias_totais_fields(resultado_diarias)
                    )
                    _update_edit_data(
                        request,
                        oficio_id,
                        {
                            "quantidade_diarias": quantidade_diarias,
                            "valor_diarias": valor_diarias,
                            "valor_diarias_extenso": valor_diarias_extenso_auto
                            or valor_diarias_extenso,
                        },
                    )

        if roteiro_origem_id:
            roteiro_atualizado = _persist_modified_roteiro_snapshot(
                oficio_obj,
                roteiro_origem_id,
            )
            if roteiro_atualizado:
                roteiro_id = str(roteiro_atualizado.pk)
                roteiro_origem_id = str(roteiro_atualizado.pk)
                data = _update_edit_data(
                    request,
                    oficio_id,
                    {
                        "roteiro_id": roteiro_id,
                        "roteiro_origem_id": roteiro_origem_id,
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
        sede_cidade = _get_sede_cidade_default_id()
    sede_estado = _resolve_estado(sede_uf)
    sede_cidade_obj = _resolve_cidade(sede_cidade, estado=sede_estado)
    sede_label = _format_trecho_local(sede_cidade_obj, sede_estado)
    sede_uf_label = _format_estado_label(sede_estado, fallback=sede_uf)

    destinos_session = _normalize_destinos_for_wizard(data.get("destinos"))
    destinos_display = _build_destinos_display(destinos_session)
    roteiro_id = str(data.get("roteiro_id") or oficio_obj.roteiro_id or "")
    roteiro_origem_id = str(data.get("roteiro_origem_id") or roteiro_id or "")
    roteiro_atual = _resolve_roteiro_by_id(roteiro_id)

    trechos_session = data.get("trechos") or []
    trechos_initial = _normalize_trechos_initial(trechos_session)
    destinos_order = ",".join(str(idx) for idx in range(len(destinos_display)))

    return render(
        request,
        "viagens/oficio_edit_step3.html",
        {
            "oficio_id": oficio_id,
            "oficio_display": _oficio_display_label(oficio_obj),
            "formset": formset,
            "motivo": motivo_val,
            "tipo_destino": tipo_destino,
            "retorno_saida_data": retorno_saida_data_raw,
            "retorno_saida_hora": retorno_saida_hora_raw,
            "retorno_chegada_data": retorno_chegada_data_raw,
            "retorno_chegada_hora": retorno_chegada_hora_raw,
            "retorno_saida_cidade": data.get("retorno_saida_cidade", ""),
            "retorno_chegada_cidade": data.get("retorno_chegada_cidade", ""),
            "quantidade_diarias": data.get("quantidade_diarias", ""),
            "valor_diarias": data.get("valor_diarias", ""),
            "valor_diarias_extenso": valor_diarias_extenso,
            "quantidade_servidores": len(data.get("viajantes_ids", [])),
            "estados": estados,
            "sede_uf": sede_uf,
            "sede_uf_label": sede_uf_label,
            "sede_cidade": sede_cidade,
            "sede_label": sede_label,
            "roteiro_id": roteiro_id,
            "roteiro_origem_id": roteiro_origem_id,
            "roteiro_nome": roteiro_atual.nome if roteiro_atual else "",
            "roteiro_destinos": roteiro_atual.get_destinos_display()
            if roteiro_atual
            else "",
            "destinos": destinos_display,
            "destinos_total_forms": len(destinos_display),
            "destinos_order": destinos_order,
            "calcular_diarias_url": reverse(
                "oficio_calcular_diarias_oficio", args=[oficio_id]
            ),
        },
    )


@require_http_methods(["GET", "POST"])
def oficio_edit_step4(request, oficio_id: int):
    data = _ensure_edit_session(request, oficio_id)
    oficio_obj = get_object_or_404(Oficio, id=oficio_id)
    if request.method == "POST":
        return _redirect_to_edit_step(
            request, oficio_id=oficio_id, default_view="oficio_edit_step4"
        )
    context = _build_step4_context(data)
    context.update(
        {
            "oficio_id": oficio_id,
            "oficio_display": _oficio_display_label(oficio_obj),
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

    temp_trechos: list[Trecho] = []
    for trecho in trechos_data:
        temp_trechos.append(
            Trecho(
                destino_estado=_resolve_estado(trecho.get("destino_estado")),
                destino_cidade=_resolve_cidade(trecho.get("destino_cidade")),
            )
        )
    tipo_destino_val = infer_tipo_destino(temp_trechos) if temp_trechos else ""
    try:
        resultado_diarias = _calculate_periodized_diarias_for_serialized_trechos(
            trechos_data,
            retorno_chegada_data,
            retorno_chegada_hora,
            quantidade_servidores=len(draft.get("viajantes_ids", [])),
        )
        quantidade_diarias, valor_diarias, valor_diarias_extenso = _diarias_totais_fields(
            resultado_diarias
        )
    except ValueError:
        quantidade_diarias = draft.get("quantidade_diarias", "")
        valor_diarias = draft.get("valor_diarias", "")
        valor_diarias_extenso = draft.get("valor_diarias_extenso", "")

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
    carona_oficio_referencia = (
        _resolve_oficio_by_id(draft.get("carona_oficio_referencia_id"))
        if motorista_carona
        else None
    )
    custeio_tipo = _normalize_custeio_choice(
        draft.get("custeio_tipo") or draft.get("custos")
    )
    nome_instituicao_custeio = _normalize_nome_instituicao_custeio(
        custeio_tipo, draft.get("nome_instituicao_custeio", "")
    )
    oficio_numero, oficio_ano, oficio_formatado = _parse_oficio_parts(
        draft.get("oficio", "")
    )
    (
        motorista_oficio_numero,
        motorista_oficio_ano,
        motorista_oficio_formatado,
    ) = _parse_oficio_parts(
        draft.get("motorista_oficio", ""),
        numero_raw=draft.get("motorista_oficio_numero", ""),
        ano_raw=draft.get("motorista_oficio_ano", ""),
        default_year=timezone.localdate().year if motorista_carona else None,
    )

    logger.debug(
        "[edit-save] before db save oficio_id=%s custeio_tipo=%s motorista_oficio=%s motorista_protocolo=%s",
        oficio_id,
        custeio_tipo,
        draft.get("motorista_oficio", ""),
        draft.get("motorista_protocolo", ""),
    )

    with transaction.atomic():
        oficio_obj = get_object_or_404(Oficio, id=oficio_id)
        if oficio_numero is None:
            oficio_numero = oficio_obj.numero
        if oficio_ano is None:
            oficio_ano = oficio_obj.ano
        if not oficio_formatado:
            oficio_formatado = format_oficio_num(oficio_numero, oficio_ano)
        oficio_obj.oficio = oficio_formatado
        oficio_obj.numero = oficio_numero
        oficio_obj.ano = oficio_ano
        oficio_obj.protocolo = draft.get("protocolo", "")
        oficio_obj.roteiro = _resolve_roteiro_by_id(draft.get("roteiro_id"))
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
        oficio_obj.valor_diarias_extenso = valor_diarias_extenso
        oficio_obj.placa = placa_norm or placa
        oficio_obj.modelo = modelo
        oficio_obj.combustivel = combustivel
        oficio_obj.tipo_viatura = (veiculo.tipo_viatura if veiculo else "") or draft.get(
            "tipo_viatura", ""
        )
        oficio_obj.motorista = motorista_nome
        oficio_obj.motorista_oficio = motorista_oficio_formatado
        oficio_obj.motorista_oficio_numero = motorista_oficio_numero
        oficio_obj.motorista_oficio_ano = motorista_oficio_ano
        oficio_obj.motorista_protocolo = draft.get("motorista_protocolo", "")
        oficio_obj.motorista_carona = motorista_carona
        oficio_obj.carona_oficio_referencia = carona_oficio_referencia
        oficio_obj.motorista_viajante = motorista_obj
        oficio_obj.motivo = draft.get("motivo", "")
        oficio_obj.custeio_tipo = custeio_tipo
        oficio_obj.nome_instituicao_custeio = nome_instituicao_custeio
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

    logger.debug(
        "[edit-save] after db save oficio_id=%s custeio_tipo=%s nome_instituicao=%s carona_ref=%s",
        oficio_id,
        oficio_obj.custeio_tipo,
        oficio_obj.nome_instituicao_custeio,
        getattr(oficio_obj, "carona_oficio_referencia_id", None),
    )
    _clear_edit_data(request, oficio_id)
    return redirect(f"{reverse('oficio_edit_step4', args=[oficio_id])}?salvo=1")


@require_http_methods(["GET"])
def oficio_edit_cancel(request, oficio_id: int):
    _clear_edit_data(request, oficio_id)
    return redirect("oficios_lista")


@require_http_methods(["GET", "POST"])
def viajante_cadastro(request):
    if request.method == "POST":
        form = ViajanteNormalizeForm(request.POST)
        if form.is_valid():
            cargo = _resolver_cargo_nome(form.cleaned_data.get("cargo", ""))
            if cargo:
                _ensure_cargo_exists(cargo)
                viajante = form.save(commit=False)
                viajante.cargo = cargo
                try:
                    viajante.save()
                except ValidationError as exc:
                    erro = "; ".join(exc.messages) or "Dados invalidos."
                    return render(
                        request,
                        "viagens/viajante_form.html",
                        {
                            "erro": erro,
                            "cargo_choices": _get_cargo_choices(),
                            "values": request.POST,
                        },
                    )
                return redirect("viajantes_lista")

        erro = "Preencha nome completo, RG, CPF e cargo."
        if form.errors:
            primeiro_erro = next(iter(form.errors.values()))
            if primeiro_erro:
                erro = primeiro_erro[0]
        return render(
            request,
            "viagens/viajante_form.html",
            {
                "erro": erro,
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
        q_digits = normalize_digits(q)
        q_rg = normalize_rg(q)
        query = models.Q(nome__icontains=q) | models.Q(cargo__icontains=q)
        if q_rg:
            query |= models.Q(rg__icontains=q_rg)
        if q_digits:
            query |= models.Q(cpf__icontains=q_digits) | models.Q(telefone__icontains=q_digits)
        else:
            query |= models.Q(telefone__icontains=q)
        viajantes = viajantes.filter(query)
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
            q_oficio = normalize_oficio_num(q)
            q_protocolo = normalize_protocolo_num(q)
            query = (
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
            )
            if q_oficio:
                query |= models.Q(oficio__icontains=q_oficio)
            if q_protocolo:
                query |= models.Q(protocolo__icontains=q_protocolo)
            oficios = oficios.filter(query).distinct()
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


def ensure_oficio_has_evento(oficio: Oficio) -> Evento:
    """Se o ofício não tem evento, cria um e associa (migrate-on-access)."""
    if getattr(oficio, "evento_id", None):
        return oficio.evento
    titulo = f"Pacote — {getattr(oficio, 'numero_formatado', None) or getattr(oficio, 'oficio', None) or oficio.assunto or f'Ofício {oficio.id}'}"
    evento = Evento.objects.create(titulo=titulo[:255], tem_convite_ou_oficio_evento=True)
    Oficio.objects.filter(pk=oficio.pk).update(evento=evento)
    oficio.evento = evento
    return evento


@require_GET
def eventos_lista(request):
    """Listagem de eventos (pacotes) com status do checklist."""
    from ..services.evento_checklist import build_evento_checklist

    eventos = (
        Evento.objects.prefetch_related("oficios__trechos", "oficios__viajantes")
        .select_related("cidade_base")
        .order_by("-created_at")
    )
    rows = []
    for ev in eventos:
        checklist = build_evento_checklist(ev)
        oficios_count = ev.oficios.count()
        rows.append({
            "evento": ev,
            "checklist": checklist,
            "oficios_count": oficios_count,
        })
    return render(
        request,
        "viagens/eventos_lista.html",
        {"rows": rows},
    )


@require_GET
def evento_pacote(request, evento_id: int):
    """Pacote do evento: central única com resumo, ofícios, termos, plano/ordem, justificativas e uploads assinados."""
    from ..services.documentos_manager import build_documentos_status
    from ..services.evento_checklist import build_evento_checklist
    from ..services.evento_assinados import (
        get_status_assinados_evento,
        is_evento_pronto_para_compilar,
        listar_pendencias_compilacao,
    )

    evento = get_object_or_404(
        Evento.objects.prefetch_related(
            "oficios__trechos",
            "oficios__viajantes",
        ).select_related("cidade_base"),
        id=evento_id,
    )
    checklist = build_evento_checklist(evento)
    status_assinados = get_status_assinados_evento(evento)
    pronto_para_compilar = is_evento_pronto_para_compilar(evento)
    pendencias_compilacao = listar_pendencias_compilacao(evento)
    ultimo_protocolo = evento.protocolos_compilados.order_by("-compilado_em").first()
    oficios = list(evento.oficios.all().order_by("id"))
    oficios_assinados_map = {s["oficio_id"]: s for s in status_assinados["oficios"]}
    oficios_com_assinado = [
        (o, oficios_assinados_map.get(o.id, {"assinado": False, "arquivo": None, "numero_display": getattr(o, "numero_formatado", None) or getattr(o, "oficio", None) or str(o.id)}))
        for o in oficios
    ]
    documentos_status_por_oficio = {o.id: build_documentos_status(o) for o in oficios}
    tab = (request.GET.get("tab") or "resumo").strip().lower()
    allowed = {"resumo", "oficios", "termos", "plano", "justificativas"}
    if tab not in allowed:
        tab = "resumo"
    from ..services.documentos_manager import _get_ordem, _get_plano
    plano_oficio = None
    ordem_oficio = None
    for o in oficios:
        if _get_plano(o) is not None and plano_oficio is None:
            plano_oficio = o
        if _get_ordem(o) is not None and ordem_oficio is None:
            ordem_oficio = o
    return render(
        request,
        "viagens/evento_pacote.html",
        {
            "evento": evento,
            "checklist": checklist,
            "status_assinados": status_assinados,
            "oficios_com_assinado": oficios_com_assinado,
            "pronto_para_compilar": pronto_para_compilar,
            "pendencias_compilacao": pendencias_compilacao,
            "ultimo_protocolo": ultimo_protocolo,
            "oficios": oficios,
            "documentos_status_por_oficio": documentos_status_por_oficio,
            "active_tab": tab,
            "plano_oficio": plano_oficio,
            "ordem_oficio": ordem_oficio,
        },
    )


def evento_redirect_from_oficio(request, oficio_id: int):
    """Redireciona para o pacote do evento; cria evento se o ofício ainda não tiver."""
    oficio = get_object_or_404(Oficio, id=oficio_id)
    evento = ensure_oficio_has_evento(oficio)
    return redirect("evento_pacote", evento_id=evento.id)


# --- Fluxo guiado por evento (wizard) ---


@login_required
@require_http_methods(["GET", "POST"])
def evento_novo_guiado(request):
    """Cria evento rascunho e redireciona para etapa 1 do fluxo guiado."""
    evento = Evento.objects.create(titulo="Rascunho")
    messages.success(request, "Novo evento criado. Preencha os dados na Etapa 1.")
    return redirect("evento_guiado_etapa1", evento_id=evento.id)


@login_required
@require_http_methods(["GET", "POST"])
def evento_guiado_etapa1(request, evento_id: int):
    """Etapa 1: formulário de cadastro do evento (título, cidade base, datas, tem_convite, tipo_demanda)."""
    evento = get_object_or_404(Evento, id=evento_id)
    from ..forms_evento_guiado import EventoGuiadoStep1Form

    initial = {}
    try:
        default_cidade_id = _get_sede_cidade_default_id()
        if default_cidade_id:
            initial["cidade_base"] = int(default_cidade_id)
    except (ValueError, TypeError):
        pass

    if request.method == "POST":
        form = EventoGuiadoStep1Form(request.POST, instance=evento)
        action = (request.POST.get("action") or "").strip()

        if action == "pular":
            messages.info(request, "Você pulou a etapa. Complete depois pelo painel.")
            return redirect("evento_guiado_painel", evento_id=evento.id)

        if form.is_valid():
            form.save()
            if action == "save_continue":
                messages.success(request, "Evento salvo. Avançando para o painel.")
                return redirect("evento_guiado_painel", evento_id=evento.id)
            messages.success(request, "Evento salvo.")
            return redirect("evento_guiado_etapa1", evento_id=evento.id)
        # form invalid
        return render(
            request,
            "viagens/evento_guiado_etapa1.html",
            {"form": form, "evento": evento},
        )

    form = EventoGuiadoStep1Form(instance=evento, initial=initial)
    return render(
        request,
        "viagens/evento_guiado_etapa1.html",
        {"form": form, "evento": evento},
    )


@login_required
@require_GET
def evento_guiado_painel(request, evento_id: int):
    """Painel do fluxo guiado: progresso e pendências, com link para próximo passo."""
    evento = get_object_or_404(Evento.objects.select_related("cidade_base"), id=evento_id)
    from ..services.evento_guiado import build_evento_guiado_progresso

    progresso = build_evento_guiado_progresso(evento)
    return render(
        request,
        "viagens/evento_guiado_painel.html",
        {"evento": evento, "progresso": progresso},
    )


@login_required
@require_http_methods(["GET", "POST"])
def evento_guiado_etapa2(request, evento_id: int):
    """Etapa 2: roteiros do evento (ida e volta) com cálculo automático de chegada."""
    evento = get_object_or_404(
        Evento.objects.select_related("cidade_base"),
        id=evento_id,
    )
    from ..forms_evento_guiado import EventoRoteiroStep2Form
    from ..services.evento_roteiro_calculo import calcular_chegada, parse_duracao_minutos

    roteiros = list(evento.roteiros.filter(ativo=True).prefetch_related("trechos").order_by("id"))
    initial = {}
    editar_roteiro_id = request.GET.get("editar") or (request.POST.get("roteiro_id") if request.method == "POST" else None)
    if editar_roteiro_id:
        try:
            roteiro_editar = evento.roteiros.prefetch_related("trechos").get(id=int(editar_roteiro_id))
            trecho_ida = roteiro_editar.trechos.order_by("ordem").first()
            if trecho_ida:
                initial["origem_cidade"] = trecho_ida.origem_cidade_id or roteiro_editar.cidade_sede_id
                initial["destino_estado"] = trecho_ida.destino_estado_id
                initial["destino_cidade"] = trecho_ida.destino_cidade_id
                initial["saida_data"] = trecho_ida.saida_data
                initial["saida_hora"] = trecho_ida.saida_hora
                initial["duracao_ida"] = (
                    f"{trecho_ida.tempo_viagem_minutos // 60}:{trecho_ida.tempo_viagem_minutos % 60:02d}"
                    if trecho_ida.tempo_viagem_minutos is not None
                    else ""
                )
            initial["retorno_saida_data"] = roteiro_editar.retorno_saida_data
            initial["retorno_saida_hora"] = roteiro_editar.retorno_saida_hora
            if all([
                roteiro_editar.retorno_saida_data,
                roteiro_editar.retorno_saida_hora,
                roteiro_editar.retorno_chegada_data,
                roteiro_editar.retorno_chegada_hora,
            ]):
                from datetime import datetime
                delta = (
                    datetime.combine(roteiro_editar.retorno_chegada_data, roteiro_editar.retorno_chegada_hora)
                    - datetime.combine(roteiro_editar.retorno_saida_data, roteiro_editar.retorno_saida_hora)
                )
                mins = int(delta.total_seconds() // 60)
                if mins >= 0:
                    initial["duracao_retorno"] = f"{mins // 60}:{mins % 60:02d}"
        except (ValueError, Roteiro.DoesNotExist):
            editar_roteiro_id = None
    if not initial and not evento.cidade_base_id:
        try:
            default_cidade_id = _get_sede_cidade_default_id()
            if default_cidade_id:
                initial["origem_cidade"] = int(default_cidade_id)
        except (ValueError, TypeError):
            pass
    form = EventoRoteiroStep2Form(request.POST or None, evento=evento, initial=initial or None)
    if initial.get("destino_estado"):
        form.fields["destino_cidade"].queryset = Cidade.objects.filter(estado_id=initial["destino_estado"]).order_by("nome")

    # Atualizar queryset de destino_cidade se veio destino_estado no POST
    if request.method == "POST" and request.POST.get("destino_estado"):
        try:
            estado_id = int(request.POST.get("destino_estado"))
            form.fields["destino_cidade"].queryset = Cidade.objects.filter(estado_id=estado_id).order_by("nome")
        except (ValueError, TypeError):
            pass

    if request.method == "POST":
        action = (request.POST.get("action") or "").strip()
        if action == "remover":
            roteiro_id = request.POST.get("roteiro_id")
            try:
                roteiro = evento.roteiros.get(id=int(roteiro_id))
                roteiro.evento_id = None
                roteiro.save(update_fields=["evento_id"])
                messages.success(request, "Roteiro removido do evento.")
            except (ValueError, Roteiro.DoesNotExist):
                messages.error(request, "Roteiro não encontrado.")
            return redirect("evento_guiado_etapa2", evento_id=evento.id)

        if action == "salvar" and form.is_valid():
            origem_cidade = form.cleaned_data.get("origem_cidade")
            destino_cidade = form.cleaned_data.get("destino_cidade")
            saida_data = form.cleaned_data.get("saida_data")
            saida_hora = form.cleaned_data.get("saida_hora")
            duracao_ida_str = (form.cleaned_data.get("duracao_ida") or "").strip()
            retorno_saida_data = form.cleaned_data.get("retorno_saida_data")
            retorno_saida_hora = form.cleaned_data.get("retorno_saida_hora")
            duracao_retorno_str = (form.cleaned_data.get("duracao_retorno") or "").strip()

            if not origem_cidade or not destino_cidade:
                messages.error(request, "Informe origem e destino.")
            else:
                minutos_ida = parse_duracao_minutos(duracao_ida_str) if duracao_ida_str else None
                minutos_retorno = parse_duracao_minutos(duracao_retorno_str) if duracao_retorno_str else None
                chegada_data, chegada_hora = calcular_chegada(saida_data, saida_hora, minutos_ida)
                retorno_chegada_data, retorno_chegada_hora = calcular_chegada(
                    retorno_saida_data, retorno_saida_hora, minutos_retorno
                )

                roteiro_id = request.POST.get("roteiro_id")
                try:
                    with transaction.atomic():
                        if roteiro_id:
                            roteiro = evento.roteiros.get(id=int(roteiro_id))
                        else:
                            roteiro = Roteiro(
                                evento=evento,
                                ativo=True,
                                tipo_deslocamento=Roteiro.TipoDeslocamentoChoices.INTERIOR,
                            )
                        estado_origem = origem_cidade.estado
                        estado_destino = destino_cidade.estado
                        roteiro.estado_sede = estado_origem
                        roteiro.cidade_sede = origem_cidade
                        roteiro.uf_origem = estado_origem.sigla if estado_origem else "PR"
                        roteiro.cidade_origem = origem_cidade.nome
                        roteiro.uf_destino = estado_destino.sigla if estado_destino else "PR"
                        roteiro.cidade_destino = destino_cidade.nome
                        roteiro.retorno_saida_cidade = destino_cidade.nome
                        roteiro.retorno_saida_data = retorno_saida_data
                        roteiro.retorno_saida_hora = retorno_saida_hora
                        roteiro.retorno_chegada_cidade = origem_cidade.nome
                        roteiro.retorno_chegada_data = retorno_chegada_data
                        roteiro.retorno_chegada_hora = retorno_chegada_hora
                        roteiro.save()

                        trechos = list(roteiro.trechos.order_by("ordem"))
                        trecho_ida = trechos[0] if trechos else None
                        if not trecho_ida:
                            trecho_ida = TrechoRoteiro(roteiro=roteiro, ordem=1)
                        trecho_ida.origem_estado = estado_origem
                        trecho_ida.origem_cidade = origem_cidade
                        trecho_ida.destino_estado = estado_destino
                        trecho_ida.destino_cidade = destino_cidade
                        trecho_ida.uf_origem = roteiro.uf_origem
                        trecho_ida.cidade_origem = origem_cidade.nome
                        trecho_ida.uf_destino = roteiro.uf_destino
                        trecho_ida.cidade_destino = destino_cidade.nome
                        trecho_ida.saida_data = saida_data
                        trecho_ida.saida_hora = saida_hora
                        trecho_ida.chegada_data = chegada_data
                        trecho_ida.chegada_hora = chegada_hora
                        trecho_ida.tempo_viagem_minutos = minutos_ida
                        trecho_ida.save()
                    messages.success(request, "Roteiro salvo. Chegada calculada automaticamente.")
                    return redirect("evento_guiado_etapa2", evento_id=evento.id)
                except Roteiro.DoesNotExist:
                    messages.error(request, "Roteiro não encontrado.")
                except Exception as e:
                    messages.error(request, f"Erro ao salvar: {e}")

    # GET ou form inválido: garantir destino_cidade com estado PR se vazio
    if not form.fields["destino_cidade"].queryset.exists():
        estado_pr = Estado.objects.filter(sigla="PR").first()
        if estado_pr:
            form.fields["destino_cidade"].queryset = Cidade.objects.filter(estado=estado_pr).order_by("nome")

    return render(
        request,
        "viagens/evento_guiado_etapa2.html",
        {
            "evento": evento,
            "roteiros": roteiros,
            "form": form,
            "editar_roteiro_id": int(editar_roteiro_id) if editar_roteiro_id else None,
        },
    )


@login_required
@require_GET
def evento_guiado_etapa3(request, evento_id: int):
    """Etapa 3: hub de ofícios do evento — lista, criar novo (wizard com contexto evento), editar, central de documentos."""
    evento = get_object_or_404(
        Evento.objects.prefetch_related("oficios__trechos"),
        id=evento_id,
    )
    oficios = list(evento.oficios.order_by("id"))
    return render(
        request,
        "viagens/evento_guiado_etapa3.html",
        {
            "evento": evento,
            "oficios": oficios,
        },
    )


def _evento_etapa4_planos_ordens(evento: Evento) -> tuple[list[tuple[Oficio, PlanoTrabalho | None]], list[tuple[Oficio, OrdemServico | None]]]:
    """Retorna (lista (oficio, plano), lista (oficio, ordem)) para ofícios do evento."""
    oficios = list(
        evento.oficios.prefetch_related("plano_trabalho", "ordem_servico").order_by("id")
    )
    planos: list[tuple[Oficio, PlanoTrabalho | None]] = []
    ordens: list[tuple[Oficio, OrdemServico | None]] = []
    for oficio in oficios:
        try:
            plano = oficio.plano_trabalho
        except PlanoTrabalho.DoesNotExist:
            plano = None
        try:
            ordem = oficio.ordem_servico
        except OrdemServico.DoesNotExist:
            ordem = None
        planos.append((oficio, plano))
        ordens.append((oficio, ordem))
    return planos, ordens


@login_required
@require_GET
def evento_guiado_etapa4(request, evento_id: int):
    """Etapa 4: Base formal (convite/ofício) ou PT/OS. Se tem_convite: dispensado. Senão: exigir criar PT ou OS vinculado ao evento."""
    evento = get_object_or_404(Evento, id=evento_id)
    tem_convite = getattr(evento, "tem_convite_ou_oficio_evento", True)
    oficios = list(evento.oficios.order_by("id"))
    planos_list: list[tuple[Oficio, PlanoTrabalho | None]] = []
    ordens_list: list[tuple[Oficio, OrdemServico | None]] = []
    if not tem_convite:
        planos_list, ordens_list = _evento_etapa4_planos_ordens(evento)
    url_criar_plano = reverse("evento_guiado_etapa4_criar_plano", kwargs={"evento_id": evento.id})
    url_criar_ordem = reverse("evento_guiado_etapa4_criar_ordem", kwargs={"evento_id": evento.id})
    return_path_etapa4 = reverse("evento_guiado_etapa4", kwargs={"evento_id": evento.id})
    return render(
        request,
        "viagens/evento_guiado_etapa4.html",
        {
            "evento": evento,
            "tem_convite": tem_convite,
            "oficios": oficios,
            "planos_list": planos_list,
            "ordens_list": ordens_list,
            "url_criar_plano": url_criar_plano,
            "url_criar_ordem": url_criar_ordem,
            "return_path_etapa4": return_path_etapa4,
        },
    )


@login_required
@require_GET
def evento_guiado_etapa4_criar_plano(request, evento_id: int):
    """Wrapper: criar ou abrir Plano de Trabalho vinculado ao evento. Redireciona para etapa-1 do plano com ?next=etapa-4."""
    evento = get_object_or_404(Evento, id=evento_id)
    return_url = reverse("evento_guiado_etapa4", kwargs={"evento_id": evento.id})
    oficio_id_raw = request.GET.get("oficio_id")
    novo = request.GET.get("novo")

    if novo:
        oficio = Oficio.objects.create(status=Oficio.Status.DRAFT, evento_id=evento.id)
        _ensure_plano_wizard_instance(oficio)
        messages.success(request, "Ofício base e plano de trabalho criados. Preencha o plano.")
        url = reverse("plano_trabalho_step1", kwargs={"oficio_id": oficio.id})
        return redirect(f"{url}?next={quote(return_url)}")
    if oficio_id_raw:
        try:
            oficio = evento.oficios.get(id=int(oficio_id_raw))
        except (ValueError, Oficio.DoesNotExist):
            messages.error(request, "Ofício inválido ou não pertence a este evento.")
            return redirect("evento_guiado_etapa4", evento_id=evento.id)
        if oficio.evento_id != evento.id:
            oficio.evento_id = evento.id
            oficio.save(update_fields=["evento_id"])
        _ensure_plano_wizard_instance(oficio)
        messages.success(request, "Plano de trabalho aberto para edição.")
        url = reverse("plano_trabalho_step1", kwargs={"oficio_id": oficio.id})
        return redirect(f"{url}?next={quote(return_url)}")

    messages.warning(request, "Escolha um ofício base ou opte por criar novo ofício para o plano.")
    return redirect("evento_guiado_etapa4", evento_id=evento.id)


@login_required
@require_GET
def evento_guiado_etapa4_criar_ordem(request, evento_id: int):
    """Wrapper: criar ou abrir Ordem de Serviço vinculada ao evento. Redireciona para edição da ordem com ?next=etapa-4."""
    evento = get_object_or_404(Evento, id=evento_id)
    return_url = reverse("evento_guiado_etapa4", kwargs={"evento_id": evento.id})
    oficio_id_raw = request.GET.get("oficio_id")
    novo = request.GET.get("novo")

    if novo:
        oficio = Oficio.objects.create(status=Oficio.Status.DRAFT, evento_id=evento.id)
        ano = int(oficio.ano or timezone.localdate().year)
        initial = _ordem_initial_data(oficio)
        acao = _ensure_acao_for_oficio(oficio)
        OrdemServico.objects.create(
            oficio=oficio,
            acao=acao,
            numero=get_next_ordem_num(ano),
            ano=ano,
            referencia=str(initial.get("referencia") or "Diligencias"),
            determinante_nome=str(initial.get("determinante_nome") or ""),
            determinante_cargo=str(initial.get("determinante_cargo") or ""),
            finalidade=str(initial.get("finalidade") or ""),
        )
        messages.success(request, "Ofício base e ordem de serviço criados. Preencha a ordem.")
        url = reverse("ordem_servico_editar", kwargs={"oficio_id": oficio.id})
        return redirect(f"{url}?next={quote(return_url)}")
    if oficio_id_raw:
        try:
            oficio = evento.oficios.get(id=int(oficio_id_raw))
        except (ValueError, Oficio.DoesNotExist):
            messages.error(request, "Ofício inválido ou não pertence a este evento.")
            return redirect("evento_guiado_etapa4", evento_id=evento.id)
        if oficio.evento_id != evento.id:
            oficio.evento_id = evento.id
            oficio.save(update_fields=["evento_id"])
        ordem = _get_ordem_servico(oficio)
        if not ordem:
            acao = _ensure_acao_for_oficio(oficio)
            ano = int(oficio.ano or timezone.localdate().year)
            initial = _ordem_initial_data(oficio)
            OrdemServico.objects.create(
                oficio=oficio,
                acao=acao,
                numero=get_next_ordem_num(ano),
                ano=ano,
                referencia=str(initial.get("referencia") or "Diligencias"),
                determinante_nome=str(initial.get("determinante_nome") or ""),
                determinante_cargo=str(initial.get("determinante_cargo") or ""),
                finalidade=str(initial.get("finalidade") or ""),
            )
        messages.success(request, "Ordem de serviço aberta para edição.")
        url = reverse("ordem_servico_editar", kwargs={"oficio_id": oficio.id})
        return redirect(f"{url}?next={quote(return_url)}")

    messages.warning(request, "Escolha um ofício base ou opte por criar novo ofício para a ordem.")
    return redirect("evento_guiado_etapa4", evento_id=evento.id)


# --- Etapa 5: Termos de autorização (por ofício e viajante) ---


@login_required
@require_GET
def evento_guiado_etapa5(request, evento_id: int):
    """Etapa 5: Termos por ofício e viajante. Lista agrupada por ofício, geração em lote, dispensar, link para pacote."""
    from ..services.evento_termos_etapa5 import get_termos_etapa5_por_oficio

    evento = get_object_or_404(Evento, id=evento_id)
    grupos = get_termos_etapa5_por_oficio(evento)
    url_gerar_lote = reverse("evento_guiado_etapa5_gerar_lote", kwargs={"evento_id": evento.id})
    url_dispensar = reverse("evento_guiado_etapa5_dispensar", kwargs={"evento_id": evento.id})
    return_path_etapa5 = reverse("evento_guiado_etapa5", kwargs={"evento_id": evento.id})
    return render(
        request,
        "viagens/evento_guiado_etapa5.html",
        {
            "evento": evento,
            "grupos": grupos,
            "url_gerar_lote": url_gerar_lote,
            "url_dispensar": url_dispensar,
            "return_path_etapa5": return_path_etapa5,
        },
    )


@login_required
@require_http_methods(["GET", "POST"])
def evento_guiado_etapa5_gerar_lote(request, evento_id: int):
    """Cria TermoAutorizacao por (ofício, viajante) para todos os necessários. Pré-preenche com dados do ofício."""
    from ..services.evento_termos_etapa5 import build_termo_prefill_from_oficio, get_termos_etapa5_por_oficio

    evento = get_object_or_404(Evento, id=evento_id)
    grupos = get_termos_etapa5_por_oficio(evento)
    criados = 0
    hoje = timezone.localdate()
    for gr in grupos:
        oficio = gr["oficio"]
        for row in gr["viajantes"]:
            if not row["precisa_termo"] or row["status"] != "pendente":
                continue
            v = row["viajante"]
            if TermoAutorizacao.objects.filter(oficio=oficio, viajante=v).exists():
                continue
            prefill = build_termo_prefill_from_oficio(oficio, v, evento)
            data_inicio = prefill.get("data_inicio") or hoje
            data_fim = prefill.get("data_fim") or data_inicio
            destinos_raw = prefill.get("destinos") or [{"uf": "PR", "cidade": ""}]
            destinos_validos = [
                d for d in destinos_raw
                if (d.get("uf") or "").strip() and (d.get("cidade") or "").strip()
            ]
            if not destinos_validos:
                destinos_validos = [{"uf": "PR", "cidade": ""}]
            TermoAutorizacao.objects.create(
                oficio=oficio,
                evento=evento,
                viajante=v,
                data_inicio=data_inicio,
                data_fim=data_fim,
                data_unica=(data_inicio == data_fim),
                destinos=destinos_validos,
                motorista_nome=prefill.get("motorista_nome") or "",
                veiculo_modelo=prefill.get("veiculo_modelo") or "",
                veiculo_placa=prefill.get("veiculo_placa") or "",
                combustivel=prefill.get("combustivel") or "",
            )
            criados += 1
    if criados:
        messages.success(request, f"{criados} termo(s) criado(s) e pré-preenchidos a partir dos ofícios.")
    else:
        messages.info(request, "Todos os viajantes necessários já possuem termo ou dispensa.")
    return redirect("evento_guiado_etapa5", evento_id=evento.id)


@login_required
@require_POST
def evento_guiado_etapa5_dispensar(request, evento_id: int):
    """Marca termo como dispensado para (oficio_id, viajante_id) com motivo."""
    evento = get_object_or_404(Evento, id=evento_id)
    oficio_id_raw = request.POST.get("oficio_id")
    viajante_id_raw = request.POST.get("viajante_id")
    motivo = (request.POST.get("motivo") or "").strip()[:200]
    if not oficio_id_raw or not viajante_id_raw:
        messages.error(request, "oficio_id e viajante_id são obrigatórios.")
        return redirect("evento_guiado_etapa5", evento_id=evento.id)
    try:
        oficio = evento.oficios.get(id=int(oficio_id_raw))
        viajante = Viajante.objects.get(id=int(viajante_id_raw))
    except (ValueError, Oficio.DoesNotExist, Viajante.DoesNotExist):
        messages.error(request, "Ofício ou viajante inválido.")
        return redirect("evento_guiado_etapa5", evento_id=evento.id)
    if not oficio.viajantes.filter(pk=viajante.id).exists():
        messages.error(request, "Viajante não pertence a este ofício.")
        return redirect("evento_guiado_etapa5", evento_id=evento.id)
    termo = TermoAutorizacao.objects.filter(oficio=oficio, viajante=viajante).first()
    if termo:
        termo.dispensado = True
        termo.dispensa_motivo = motivo
        termo.save(update_fields=["dispensado", "dispensa_motivo", "updated_at"])
    else:
        TermoAutorizacao.objects.create(
            oficio=oficio,
            evento=evento,
            viajante=viajante,
            data_inicio=timezone.localdate(),
            data_fim=timezone.localdate(),
            data_unica=True,
            destinos=[],
            dispensado=True,
            dispensa_motivo=motivo,
        )
    messages.success(request, "Termo dispensado com sucesso.")
    return redirect("evento_guiado_etapa5", evento_id=evento.id)


@login_required
@require_GET
def evento_guiado_etapa6(request, evento_id: int):
    """Etapa 6: Finalização — checklist de uploads em ordem fixa (ofícios, solicitação/PT-OS, justificativas, termos)."""
    from ..services.evento_etapa6 import build_etapa6_checklist, listar_pendencias_etapa6

    evento = get_object_or_404(Evento, id=evento_id)
    checklist = build_etapa6_checklist(evento)
    pendencias = listar_pendencias_etapa6(evento)
    url_upload = reverse("evento_upload_assinado", kwargs={"evento_id": evento.id})
    return_path_etapa6 = reverse("evento_guiado_etapa6", kwargs={"evento_id": evento.id})
    url_exportar_zip = reverse("evento_guiado_exportar_zip", kwargs={"evento_id": evento.id})
    return render(
        request,
        "viagens/evento_guiado_etapa6.html",
        {
            "evento": evento,
            "blocos": checklist["blocos"],
            "etapa6_ok": checklist["etapa6_ok"],
            "pendencias": pendencias,
            "url_upload": url_upload,
            "return_path_etapa6": return_path_etapa6,
            "url_exportar_zip": url_exportar_zip,
        },
    )


@login_required
@require_GET
def evento_guiado_exportar_zip(request, evento_id: int):
    """Exporta o pacote do evento em ZIP. Bloqueia se faltarem uploads obrigatórios."""
    from ..services.evento_export_zip import build_evento_zip
    from ..services.evento_assinados import is_evento_pronto_para_compilar, listar_pendencias_compilacao
    from django.http import HttpResponse

    evento = get_object_or_404(Evento, id=evento_id)
    if not is_evento_pronto_para_compilar(evento):
        pendencias = listar_pendencias_compilacao(evento)
        messages.error(
            request,
            f"Não é possível exportar: faltam {len(pendencias)} documento(s). Conclua os uploads na Etapa 6.",
        )
        for p in pendencias[:5]:
            messages.warning(request, p)
        return redirect("evento_guiado_etapa6", evento_id=evento.id)
    result = build_evento_zip(evento, compilado_por_id=getattr(request.user, "id", None))
    if not result:
        messages.error(request, "Falha ao montar o ZIP.")
        return redirect("evento_guiado_etapa6", evento_id=evento.id)
    zip_bytes, nome_zip = result
    resp = HttpResponse(zip_bytes, content_type="application/zip")
    resp["Content-Disposition"] = f'attachment; filename="{nome_zip}"'
    return resp


@login_required
@require_GET
def evento_guiado_etapa_placeholder(request, evento_id: int, etapa_num: int):
    """Placeholder para etapas 2–5: página 'Em construção'."""
    evento = get_object_or_404(Evento, id=evento_id)
    return render(
        request,
        "viagens/evento_guiado_placeholder.html",
        {"evento": evento, "etapa_num": etapa_num},
    )


ALLOWED_UPLOAD_MIMES = {"application/pdf", "image/jpeg", "image/png", "image/jpg"}
ALLOWED_UPLOAD_EXTENSIONS = {".pdf", ".jpg", ".jpeg", ".png"}


def _convert_upload_to_pdf_if_image(uploaded_file):
    """Se for imagem (jpg/png), converte para PDF em memória; senão retorna (None, None) para usar o arquivo direto."""
    name = (getattr(uploaded_file, "name", "") or "").lower()
    content_type = getattr(uploaded_file, "content_type", "") or ""
    if content_type.startswith("image/") or name.endswith((".jpg", ".jpeg", ".png")):
        from PIL import Image
        import io as io_module
        uploaded_file.seek(0)
        img = Image.open(uploaded_file)
        if img.mode in ("RGBA", "P"):
            img = img.convert("RGB")
        buf = io_module.BytesIO()
        img.save(buf, "PDF")
        buf.seek(0)
        return buf.getvalue(), ".pdf"
    return None, None


@login_required
@require_POST
def evento_upload_assinado(request, evento_id: int):
    """Upload de documento assinado (PDF ou imagem convertida para PDF) para o pacote do evento."""
    evento = get_object_or_404(Evento, id=evento_id)
    tipo = (request.POST.get("tipo") or "").strip()
    if tipo not in dict(DocumentoEventoArquivo.Tipo.choices):
        messages.error(request, "Tipo de documento inválido.")
        return redirect("evento_pacote", evento_id=evento.id)
    oficio_id = request.POST.get("oficio_id")
    oficio = None
    if oficio_id:
        try:
            oficio = evento.oficios.get(id=int(oficio_id))
        except (ValueError, Oficio.DoesNotExist):
            pass
    viajante_id = request.POST.get("viajante_id")
    viajante = None
    if viajante_id:
        try:
            viajante = Viajante.objects.get(id=int(viajante_id))
        except (ValueError, Viajante.DoesNotExist):
            pass
    arquivo = request.FILES.get("arquivo")
    if not arquivo:
        messages.error(request, "Nenhum arquivo enviado.")
        return redirect("evento_pacote", evento_id=evento.id)
    ext = (arquivo.name or "").lower()
    if not any(ext.endswith(e) for e in ALLOWED_UPLOAD_EXTENSIONS):
        messages.error(request, "Formato não permitido. Use PDF, JPG ou PNG.")
        return redirect("evento_pacote", evento_id=evento.id)
    from django.core.files.base import ContentFile
    pdf_bytes, suffix = _convert_upload_to_pdf_if_image(arquivo)
    if pdf_bytes is not None:
        original_name = (arquivo.name or "documento") + (suffix or "")
        mime = "application/pdf"
        file_to_save = ContentFile(pdf_bytes)
    else:
        arquivo.seek(0)
        original_name = arquivo.name or "documento.pdf"
        mime = getattr(arquivo, "content_type", "") or "application/pdf"
        file_to_save = arquivo
    with transaction.atomic():
        DocumentoEventoArquivo.objects.filter(
            evento=evento,
            tipo=tipo,
            is_active=True,
            oficio=oficio,
            viajante=viajante,
        ).update(is_active=False)
        novo = DocumentoEventoArquivo(
            evento=evento,
            tipo=tipo,
            oficio=oficio,
            viajante=viajante,
            original_name=original_name[:255],
            mime_type=mime[:100],
            uploaded_by_id=getattr(request.user, "id", None),
        )
        novo.arquivo.save(original_name.replace(" ", "_"), file_to_save, save=True)
        novo.save()
    messages.success(request, "Documento assinado enviado com sucesso.")
    next_url = (request.POST.get("next") or request.GET.get("next") or "").strip()
    if next_url and url_has_allowed_host_and_scheme(next_url, allowed_hosts=[request.get_host()]):
        return redirect(next_url)
    return redirect("evento_pacote", evento_id=evento.id)


@login_required
@require_POST
def evento_remover_assinado(request, evento_id: int, arquivo_id: int):
    """Desativa um arquivo assinado (remove da lista ativa)."""
    evento = get_object_or_404(Evento, id=evento_id)
    arq = get_object_or_404(DocumentoEventoArquivo, id=arquivo_id, evento=evento)
    arq.is_active = False
    arq.save(update_fields=["is_active"])
    messages.success(request, "Documento assinado removido.")
    next_url = (request.POST.get("next") or request.GET.get("next") or "").strip()
    if next_url and url_has_allowed_host_and_scheme(next_url, allowed_hosts=[request.get_host()]):
        return redirect(next_url)
    return redirect("evento_pacote", evento_id=evento.id)


@login_required
@require_GET
def evento_download_compilado(request, evento_id: int):
    """Download do PDF compilado do protocolo (última versão)."""
    evento = get_object_or_404(Evento, id=evento_id)
    ultimo = evento.protocolos_compilados.order_by("-compilado_em").first()
    if not ultimo or not ultimo.pdf_compilado:
        messages.warning(request, "Nenhum protocolo compilado disponível. Gere o PDF primeiro.")
        return redirect("evento_pacote", evento_id=evento.id)
    from django.http import HttpResponse
    nome = ultimo.pdf_compilado.name.split("/")[-1] or "protocolo.pdf"
    with ultimo.pdf_compilado.open("rb") as f:
        content = f.read()
    resp = HttpResponse(content, content_type="application/pdf")
    resp["Content-Disposition"] = f'attachment; filename="{nome}"'
    return resp


@login_required
@require_POST
def evento_gerar_pdf_protocolo(request, evento_id: int):
    """Gera o PDF compilado do protocolo e redireciona para o pacote (ou download)."""
    from ..services.evento_compilacao import compilar_pdf_protocolo
    from ..services.evento_assinados import is_evento_pronto_para_compilar
    evento = get_object_or_404(Evento, id=evento_id)
    if not is_evento_pronto_para_compilar(evento):
        messages.error(request, "Pacote incompleto: faça upload de todos os documentos assinados obrigatórios.")
        return redirect("evento_pacote", evento_id=evento.id)
    obj = compilar_pdf_protocolo(evento, compilado_por_id=getattr(request.user, "id", None))
    if obj:
        messages.success(request, "PDF do protocolo gerado com sucesso.")
        if request.POST.get("download") == "1":
            return redirect("evento_download_compilado", evento_id=evento.id)
    else:
        messages.error(request, "Não foi possível gerar o PDF do protocolo.")
    return redirect("evento_pacote", evento_id=evento.id)


def _get_justificativa_config_display() -> dict:
    """Dados da configuração para exibição na página do gerador (somente leitura)."""
    from viagens.documents.justificativa import build_justificativa_context_from_config
    return build_justificativa_context_from_config()


@require_http_methods(["GET", "POST"])
def gerador_justificativas(request):
    """Página dedicada ao gerador de justificativas: formulário com dados da config + assinante + texto."""
    from viagens.documents.justificativa import (
        build_justificativa_context_from_config,
        build_justificativa_docx_bytes,
    )
    from viagens.forms import JustificativaGeradorForm

    config_display = _get_justificativa_config_display()
    _padrao = ModeloJustificativa.objects.filter(ativo=True, padrao=True).first()
    modelo_padrao_id = _padrao.pk if _padrao else None

    if request.method == "POST":
        form = JustificativaGeradorForm(request.POST)
        if form.is_valid():
            viajante = form.cleaned_data["assinante"]
            justificativa_texto = form.cleaned_data["justificativa"]
            oficio_numero_ano = (form.cleaned_data.get("oficio_numero_ano") or "").strip()
            if oficio_numero_ano:
                from viagens.documents.justificativa import replace_oficio_numero_no_texto
                justificativa_texto = replace_oficio_numero_no_texto(justificativa_texto, oficio_numero_ano)
            from viagens.services.text import title_case_pt
            assinante_nome = title_case_pt((viajante.nome or "").strip())
            assinante_cargo = title_case_pt((viajante.cargo or "").strip())

            mapping = build_justificativa_context_from_config(
                assinante_nome=assinante_nome,
                assinante_cargo=assinante_cargo,
                justificativa_texto=justificativa_texto,
            )
            try:
                buf = build_justificativa_docx_bytes(mapping)
            except Exception as exc:
                logger.exception("[justificativa] falha na geracao")
                messages.error(request, f"Falha ao gerar documento. Detalhe: {exc}")
                return render(
                    request,
                    "viagens/gerador_justificativas_pagina.html",
                    {"form": form, "config_display": config_display, "modelo_padrao_id": modelo_padrao_id},
                )

            docx_bytes = buf.getvalue()
            if request.POST.get("gerar_pdf"):
                try:
                    pdf_bytes = docx_bytes_to_pdf_bytes(docx_bytes, oficio_id=None)
                    from django.utils import timezone
                    now = timezone.now()
                    filename = f"justificativa_{now.strftime('%Y%m%d_%H%M')}.pdf"
                    return _pdf_http_response(pdf_bytes, filename)
                except DocxPdfConversionError as exc:
                    messages.error(
                        request,
                        f"PDF indisponível neste ambiente. Baixe o DOCX. Detalhe: {exc}",
                    )
                    return render(
                        request,
                        "viagens/gerador_justificativas_pagina.html",
                        {"form": form, "config_display": config_display, "modelo_padrao_id": modelo_padrao_id},
                    )
                except Exception as exc:
                    logger.exception("[justificativa-pdf] falha")
                    messages.error(request, f"Falha ao gerar PDF. Detalhe: {exc}")
                    return render(
                        request,
                        "viagens/gerador_justificativas_pagina.html",
                        {"form": form, "config_display": config_display, "modelo_padrao_id": modelo_padrao_id},
                    )
            from django.utils import timezone
            now = timezone.now()
            filename = f"justificativa_{now.strftime('%Y%m%d_%H%M')}.docx"
            return _docx_http_response(docx_bytes, filename)
        # form inválido: cai no render abaixo com form
    else:
        initial = {}
        assinante_id = _resolve_assinante_justificativa_id_default()
        if assinante_id is not None:
            initial["assinante"] = assinante_id
        modelo_id = request.GET.get("modelo")
        modelo_padrao = ModeloJustificativa.objects.filter(ativo=True, padrao=True).first()
        if modelo_id:
            modelo = ModeloJustificativa.objects.filter(ativo=True, pk=modelo_id).first()
            if modelo:
                initial["modelo"] = modelo.pk
                initial["justificativa"] = modelo.texto
        elif modelo_padrao:
            initial["modelo"] = modelo_padrao.pk
            initial["justificativa"] = modelo_padrao.texto
        form = JustificativaGeradorForm(initial=initial)

    return render(
        request,
        "viagens/gerador_justificativas_pagina.html",
        {
            "form": form,
            "config_display": config_display,
            "modelo_padrao_id": modelo_padrao_id,
        },
    )


# ---------- Lista de justificativas (pendentes / concluídas) ----------


def _justificativa_lista_querysets():
    """Retorna (pendentes, concluídas): ofícios que exigem justificativa e ainda vazios vs preenchidos."""
    from viagens.services.justificativa_helpers import exige_justificativa, justificativa_preenchida

    oficios = (
        Oficio.objects.prefetch_related("trechos")
        .select_related("cidade_destino", "estado_destino", "cidade_sede", "estado_sede")
        .order_by("-created_at")
    )
    pendentes = []
    concluidas = []
    for o in oficios:
        if exige_justificativa(o):
            if justificativa_preenchida(o):
                concluidas.append(o)
            else:
                pendentes.append(o)
        elif justificativa_preenchida(o):
            concluidas.append(o)
    return pendentes, concluidas


@require_GET
def justificativas_lista(request):
    """Lista de justificativas: abas Pendentes e Concluídas."""
    from viagens.services.justificativa_helpers import (
        get_data_inicio_viagem,
        get_dias_antecedencia,
    )

    tab = (request.GET.get("tab") or "pendentes").strip().lower()
    if tab not in ("pendentes", "concluidas"):
        tab = "pendentes"

    pendentes, concluidas = _justificativa_lista_querysets()

    def _row(o):
        data_saida = get_data_inicio_viagem(o)
        dias = get_dias_antecedencia(o)
        destinos = ""
        if o.cidade_destino and o.estado_destino:
            destinos = f"{o.cidade_destino.nome}/{o.estado_destino.sigla}"
        else:
            destinos = o.get_destino_display() or "-"
        return {
            "oficio": o,
            "data_saida": data_saida,
            "dias_antecedencia": dias,
            "destinos": destinos,
        }

    pendentes_rows = [_row(o) for o in pendentes]
    concluidas_rows = [_row(o) for o in concluidas]

    return render(
        request,
        "viagens/justificativas_lista.html",
        {
            "tab": tab,
            "pendentes": pendentes_rows,
            "concluidas": concluidas_rows,
        },
    )


# ---------- Justificativa vinculada ao ofício ----------


@require_http_methods(["GET", "POST"])
def justificativa_oficio(request, oficio_id: int):
    """Gerador de justificativa vinculado ao ofício: pré-preenche número, Salvar e voltar persiste no Oficio."""
    from viagens.documents.justificativa import (
        build_justificativa_context_from_config,
        build_justificativa_docx_bytes,
        replace_oficio_numero_no_texto,
    )
    from viagens.forms import JustificativaGeradorForm
    from viagens.services.justificativa_helpers import exige_justificativa, justificativa_preenchida

    oficio = get_object_or_404(
        Oficio.objects.prefetch_related("trechos").select_related("cidade_sede", "estado_sede"),
        id=oficio_id,
    )
    next_url = request.GET.get("next", "").strip()
    if not next_url:
        next_url = reverse("oficio_documentos", args=[oficio_id])
    elif not next_url.startswith("/"):
        if next_url == "oficio_step4":
            next_url = reverse("oficio_step4")
        elif next_url == "oficio_documentos":
            next_url = reverse("oficio_documentos", args=[oficio_id])
        elif next_url == "evento_pacote":
            evento = ensure_oficio_has_evento(oficio)
            next_url = reverse("evento_pacote", args=[evento.id])
        else:
            next_url = reverse("oficio_documentos", args=[oficio_id])

    config_display = _get_justificativa_config_display()
    _padrao = ModeloJustificativa.objects.filter(ativo=True, padrao=True).first()
    modelo_padrao_id = _padrao.pk if _padrao else None

    oficio_numero_display = oficio.numero_formatado or (oficio.oficio or "").strip() or f"{oficio.numero or ''}/{oficio.ano or ''}".strip("/") or str(oficio_id)
    if oficio.ano and oficio.numero:
        oficio_numero_display = f"{oficio.numero}/{oficio.ano}"

    if request.method == "POST":
        action = (request.POST.get("action") or "").strip()
        form = JustificativaGeradorForm(request.POST)
        if action == "salvar_voltar":
            if form.is_valid():
                texto = form.cleaned_data["justificativa"]
                oficio_numero_ano = (form.cleaned_data.get("oficio_numero_ano") or "").strip() or oficio_numero_display
                texto = replace_oficio_numero_no_texto(texto, oficio_numero_ano)
                modelo = form.cleaned_data.get("modelo")
                oficio.justificativa_texto = texto
                oficio.justificativa_modelo = (modelo.codigo if modelo else "").strip()
                oficio.save(update_fields=["justificativa_texto", "justificativa_modelo"])
                messages.success(request, "Justificativa salva no ofício.")
                return redirect(next_url)
        elif form.is_valid():
            # Gerar DOCX ou PDF (mesma lógica do gerador livre)
            from viagens.services.text import title_case_pt
            viajante = form.cleaned_data["assinante"]
            justificativa_texto = form.cleaned_data["justificativa"]
            oficio_numero_ano = (form.cleaned_data.get("oficio_numero_ano") or "").strip() or oficio_numero_display
            justificativa_texto = replace_oficio_numero_no_texto(justificativa_texto, oficio_numero_ano)
            assinante_nome = title_case_pt((viajante.nome or "").strip())
            assinante_cargo = title_case_pt((viajante.cargo or "").strip())
            mapping = build_justificativa_context_from_config(
                assinante_nome=assinante_nome,
                assinante_cargo=assinante_cargo,
                justificativa_texto=justificativa_texto,
            )
            try:
                buf = build_justificativa_docx_bytes(mapping)
            except Exception as exc:
                logger.exception("[justificativa] falha na geracao oficio_id=%s", oficio_id)
                messages.error(request, f"Falha ao gerar documento. Detalhe: {exc}")
            else:
                docx_bytes = buf.getvalue()
                if request.POST.get("gerar_pdf"):
                    try:
                        pdf_bytes = docx_bytes_to_pdf_bytes(docx_bytes, oficio_id=oficio_id)
                        from django.utils import timezone
                        now = timezone.now()
                        filename = f"justificativa_{oficio_id}_{now.strftime('%Y%m%d_%H%M')}.pdf"
                        return _pdf_http_response(pdf_bytes, filename)
                    except DocxPdfConversionError as exc:
                        messages.error(request, f"PDF indisponível. Baixe o DOCX. Detalhe: {exc}")
                    except Exception as exc:
                        logger.exception("[justificativa-pdf] falha oficio_id=%s", oficio_id)
                        messages.error(request, f"Falha ao gerar PDF. Detalhe: {exc}")
                else:
                    from django.utils import timezone
                    now = timezone.now()
                    filename = f"justificativa_{oficio_id}_{now.strftime('%Y%m%d_%H%M')}.docx"
                    return _docx_http_response(docx_bytes, filename)
    else:
        initial = {
            "oficio_numero_ano": oficio_numero_display,
            "justificativa": (oficio.justificativa_texto or "").strip(),
        }
        assinante_id = _resolve_assinante_justificativa_id_default()
        if assinante_id is not None:
            initial["assinante"] = assinante_id
        modelo_padrao = ModeloJustificativa.objects.filter(ativo=True, padrao=True).first()
        if not initial["justificativa"] and modelo_padrao:
            initial["modelo"] = modelo_padrao.pk
            initial["justificativa"] = modelo_padrao.texto
        form = JustificativaGeradorForm(initial=initial)

    return render(
        request,
        "viagens/gerador_justificativas_pagina.html",
        {
            "form": form,
            "config_display": config_display,
            "modelo_padrao_id": modelo_padrao_id,
            "oficio": oficio,
            "oficio_id": oficio_id,
            "next_url": next_url,
            "vinculada_oficio": True,
        },
    )


# ---------- CRUD Modelos de Justificativa ----------


@require_GET
def modelos_justificativa_lista(request):
    """Lista modelos de justificativa para gerenciar (adicionar, editar, excluir, definir padrão)."""
    modelos = ModeloJustificativa.objects.all().order_by("ordem", "label")
    return render(
        request,
        "viagens/modelos_justificativa_lista.html",
        {"modelos": modelos},
    )


@require_GET
def modelo_justificativa_texto_api(request, pk: int):
    """API: retorna o texto do modelo (para preencher o textarea ao selecionar no gerador)."""
    modelo = get_object_or_404(ModeloJustificativa.objects.filter(ativo=True), pk=pk)
    return JsonResponse({"texto": modelo.texto or ""})


@require_http_methods(["GET", "POST"])
def modelo_justificativa_criar(request):
    """Criar novo modelo de justificativa."""
    from ..forms import ModeloJustificativaForm
    if request.method == "POST":
        form = ModeloJustificativaForm(request.POST)
        if form.is_valid():
            form.save()
            messages.success(request, "Modelo de justificativa criado.")
            return redirect("modelos_justificativa_lista")
    else:
        form = ModeloJustificativaForm()
    return render(request, "viagens/modelo_justificativa_form.html", {"form": form, "titulo": "Novo modelo"})


@require_http_methods(["GET", "POST"])
def modelo_justificativa_editar(request, pk: int):
    """Editar modelo de justificativa."""
    from ..forms import ModeloJustificativaForm
    obj = get_object_or_404(ModeloJustificativa, pk=pk)
    if request.method == "POST":
        form = ModeloJustificativaForm(request.POST, instance=obj)
        if form.is_valid():
            form.save()
            messages.success(request, "Modelo atualizado.")
            return redirect("modelos_justificativa_lista")
    else:
        form = ModeloJustificativaForm(instance=obj)
    return render(request, "viagens/modelo_justificativa_form.html", {"form": form, "titulo": "Editar modelo", "modelo": obj})


@require_http_methods(["POST"])
def modelo_justificativa_definir_padrao(request, pk: int):
    """Define este modelo como padrão (único)."""
    obj = get_object_or_404(ModeloJustificativa, pk=pk)
    obj.padrao = True
    obj.save(update_fields=["padrao", "updated_at"])
    messages.success(request, f'"{obj.label}" definido como modelo padrão.')
    return redirect("modelos_justificativa_lista")


@require_http_methods(["GET", "POST"])
def modelo_justificativa_salvar_como_novo(request, pk: int):
    """Duplica o modelo com novo código/label (salvar como novo)."""
    from ..forms import ModeloJustificativaForm
    origem = get_object_or_404(ModeloJustificativa, pk=pk)
    if request.method == "POST":
        form = ModeloJustificativaForm(request.POST)
        if form.is_valid():
            obj = form.save(commit=False)
            obj.padrao = False
            obj.save()
            messages.success(request, "Modelo salvo como novo.")
            return redirect("modelos_justificativa_lista")
    else:
        form = ModeloJustificativaForm(
            initial={
                "codigo": f"{origem.codigo}_copia",
                "label": f"{origem.label} (cópia)",
                "texto": origem.texto,
                "ordem": ModeloJustificativa.objects.count(),
                "ativo": True,
                "padrao": False,
            }
        )
    return render(request, "viagens/modelo_justificativa_form.html", {"form": form, "titulo": "Salvar como novo modelo"})


@require_http_methods(["POST"])
def modelo_justificativa_excluir(request, pk: int):
    """Excluir modelo de justificativa."""
    obj = get_object_or_404(ModeloJustificativa, pk=pk)
    label = obj.label
    obj.delete()
    messages.success(request, f'Modelo "{label}" excluído.')
    return redirect("modelos_justificativa_lista")


def _termo_form_context(
    *,
    erros: dict[str, str] | None = None,
    data_unica: bool = False,
    data_inicio_raw: str | None = None,
    data_fim_raw: str | None = None,
    destinos_data: list[dict[str, str]] | None = None,
) -> dict:
    hoje = timezone.localdate()
    estados = Estado.objects.order_by("sigla")
    sede_default = _get_sede_cidade_default_id()

    data_inicio_value = data_inicio_raw if data_inicio_raw is not None else hoje.isoformat()
    data_fim_value = data_fim_raw if data_fim_raw is not None else hoje.isoformat()
    destinos_value = destinos_data if destinos_data is not None else [{"uf": "PR", "cidade": sede_default}]

    destinos_normalizados = _normalize_destinos_for_wizard(destinos_value)
    destinos_display = _build_destinos_display(destinos_normalizados)
    destinos_order = ",".join(str(idx) for idx in range(len(destinos_display)))

    return {
        "erros": erros or {},
        "estados": estados,
        "data_unica": data_unica,
        "data_inicio": data_inicio_value,
        "data_fim": data_fim_value,
        "destinos": destinos_display,
        "destinos_total_forms": len(destinos_display),
        "destinos_order": destinos_order,
    }


@require_http_methods(["GET", "POST"])
def termo_autorizacao_cadastro(request):
    erros: dict[str, str] = {}
    hoje = timezone.localdate()
    data_unica = False
    data_inicio_raw = hoje.isoformat()
    data_fim_raw = hoje.isoformat()
    destinos_data: list[dict[str, str]] = [{"uf": "PR", "cidade": _get_sede_cidade_default_id()}]

    if request.method != "POST":
        return render(
            request,
            "viagens/termos_autorizacao_form.html",
            _termo_form_context(
                erros=erros,
                data_unica=data_unica,
                data_inicio_raw=data_inicio_raw,
                data_fim_raw=data_fim_raw,
                destinos_data=destinos_data,
            ),
        )

    data_unica = _is_truthy_post(request.POST.get("data_unica"))
    data_inicio_raw = (request.POST.get("data_inicio") or "").strip()
    data_fim_raw = (request.POST.get("data_fim") or "").strip()
    _, _, destinos_post = _serialize_sede_destinos_from_post(request.POST)
    destinos_data = destinos_post or [{}]

    data_inicio = parse_date(data_inicio_raw) if data_inicio_raw else None
    data_fim = parse_date(data_fim_raw) if data_fim_raw else None
    if not data_inicio:
        erros["data_inicio"] = "Informe a primeira data."
    if data_unica:
        data_fim = data_inicio
        data_fim_raw = ""
    else:
        if not data_fim:
            erros["data_fim"] = "Informe a segunda data."
        if data_inicio and data_fim and data_fim < data_inicio:
            erros["data_fim"] = "A segunda data deve ser maior ou igual a primeira."

    destinos_validos = [
        destino
        for destino in destinos_post
        if (destino.get("uf") or "").strip() and (destino.get("cidade") or "").strip()
    ]
    destinos_resolvidos = _resolve_termo_destinos_labels(destinos_validos)
    if not destinos_resolvidos:
        erros["destinos"] = "Informe ao menos um destino."

    if erros:
        return render(
            request,
            "viagens/termos_autorizacao_form.html",
            _termo_form_context(
                erros=erros,
                data_unica=data_unica,
                data_inicio_raw=data_inicio_raw,
                data_fim_raw=data_fim_raw,
                destinos_data=destinos_data,
            ),
        )

    termo_nome = _build_termo_nome(destinos_resolvidos)
    TermoAutorizacao.objects.create(
        data_inicio=data_inicio or hoje,
        data_fim=data_fim,
        data_unica=data_unica,
        destinos=destinos_validos,
    )
    messages.success(request, f'Termo salvo como "{termo_nome}".')
    return redirect("termos_autorizacao_lista")


@login_required
@require_http_methods(["GET", "POST"])
def termo_autorizacao_cadastro_contextual(request):
    """Cadastro/edição de termo no contexto evento+ofício+viajante (Etapa 5). Aceita evento_id, oficio_id, viajante_id, next."""
    from ..services.evento_termos_etapa5 import build_termo_prefill_from_oficio

    evento_id_raw = request.GET.get("evento_id") or request.POST.get("evento_id")
    oficio_id_raw = request.GET.get("oficio_id") or request.POST.get("oficio_id")
    viajante_id_raw = request.GET.get("viajante_id") or request.POST.get("viajante_id")
    next_url = (request.GET.get("next") or request.POST.get("next") or "").strip()

    if not evento_id_raw or not oficio_id_raw or not viajante_id_raw:
        messages.error(request, "Parâmetros evento_id, oficio_id e viajante_id são obrigatórios.")
        return redirect("termos_autorizacao_lista")
    try:
        evento = Evento.objects.get(id=int(evento_id_raw))
        oficio = Oficio.objects.get(id=int(oficio_id_raw))
        viajante = Viajante.objects.get(id=int(viajante_id_raw))
    except (ValueError, Evento.DoesNotExist, Oficio.DoesNotExist, Viajante.DoesNotExist):
        messages.error(request, "Evento, ofício ou viajante inválido.")
        return redirect("termos_autorizacao_lista")
    if oficio.evento_id != evento.id:
        messages.error(request, "Ofício não pertence a este evento.")
        return redirect("evento_guiado_etapa5", evento_id=evento.id)
    if not oficio.viajantes.filter(pk=viajante.id).exists():
        messages.error(request, "Viajante não pertence a este ofício.")
        return redirect("evento_guiado_etapa5", evento_id=evento.id)

    termo = TermoAutorizacao.objects.filter(oficio=oficio, viajante=viajante).exclude(dispensado=True).first()
    erros: dict[str, str] = {}
    hoje = timezone.localdate()

    if request.method != "POST":
        if termo:
            data_inicio_raw = termo.data_inicio.isoformat() if termo.data_inicio else hoje.isoformat()
            data_fim_raw = (termo.data_fim or termo.data_inicio or hoje).isoformat()
            data_unica = bool(termo.data_unica or (termo.data_inicio and termo.data_fim and termo.data_inicio == termo.data_fim))
            destinos_data = termo.destinos if isinstance(termo.destinos, list) else [{"uf": "PR", "cidade": ""}]
        else:
            prefill = build_termo_prefill_from_oficio(oficio, viajante, evento)
            data_inicio_raw = (prefill["data_inicio"] or hoje).isoformat()
            data_fim_raw = (prefill["data_fim"] or prefill["data_inicio"] or hoje).isoformat()
            data_unica = prefill.get("data_unica", False)
            destinos_data = prefill.get("destinos") or [{"uf": "PR", "cidade": ""}]
        ctx = _termo_form_context(
            erros=erros,
            data_unica=data_unica,
            data_inicio_raw=data_inicio_raw,
            data_fim_raw=data_fim_raw,
            destinos_data=destinos_data,
        )
        ctx["evento"] = evento
        ctx["oficio"] = oficio
        ctx["viajante"] = viajante
        ctx["termo"] = termo
        ctx["evento_return_url"] = next_url if next_url else None
        ctx["evento_id"] = evento.id
        ctx["oficio_id"] = oficio.id
        ctx["viajante_id"] = viajante.id
        ctx["next_url"] = next_url
        ctx["is_contextual"] = True
        return render(request, "viagens/termos_autorizacao_form.html", ctx)

    data_unica = _is_truthy_post(request.POST.get("data_unica"))
    data_inicio_raw = (request.POST.get("data_inicio") or "").strip()
    data_fim_raw = (request.POST.get("data_fim") or "").strip()
    _, _, destinos_post = _serialize_sede_destinos_from_post(request.POST)
    destinos_data = destinos_post or [{}]

    data_inicio = parse_date(data_inicio_raw) if data_inicio_raw else None
    data_fim = parse_date(data_fim_raw) if data_fim_raw else None
    if not data_inicio:
        erros["data_inicio"] = "Informe a primeira data."
    if data_unica:
        data_fim = data_inicio
        data_fim_raw = ""
    else:
        if not data_fim:
            erros["data_fim"] = "Informe a segunda data."
        if data_inicio and data_fim and data_fim < data_inicio:
            erros["data_fim"] = "A segunda data deve ser maior ou igual a primeira."

    destinos_validos = [
        d for d in destinos_post
        if (d.get("uf") or "").strip() and (d.get("cidade") or "").strip()
    ]
    destinos_resolvidos = _resolve_termo_destinos_labels(destinos_validos)
    if not destinos_resolvidos:
        erros["destinos"] = "Informe ao menos um destino."

    if erros:
        ctx = _termo_form_context(
            erros=erros,
            data_unica=data_unica,
            data_inicio_raw=data_inicio_raw,
            data_fim_raw=data_fim_raw,
            destinos_data=destinos_data,
        )
        ctx["evento"] = evento
        ctx["oficio"] = oficio
        ctx["viajante"] = viajante
        ctx["termo"] = termo
        ctx["evento_return_url"] = next_url if next_url else None
        ctx["evento_id"] = evento.id
        ctx["oficio_id"] = oficio.id
        ctx["viajante_id"] = viajante.id
        ctx["next_url"] = next_url
        ctx["is_contextual"] = True
        return render(request, "viagens/termos_autorizacao_form.html", ctx)

    prefill = build_termo_prefill_from_oficio(oficio, viajante, evento)
    if termo:
        termo.data_inicio = data_inicio or hoje
        termo.data_fim = data_fim
        termo.data_unica = data_unica
        termo.destinos = destinos_validos
        termo.motorista_nome = prefill.get("motorista_nome") or ""
        termo.veiculo_modelo = prefill.get("veiculo_modelo") or ""
        termo.veiculo_placa = prefill.get("veiculo_placa") or ""
        termo.combustivel = prefill.get("combustivel") or ""
        termo.save()
        messages.success(request, "Termo atualizado.")
    else:
        TermoAutorizacao.objects.create(
            oficio=oficio,
            evento=evento,
            viajante=viajante,
            data_inicio=data_inicio or hoje,
            data_fim=data_fim,
            data_unica=data_unica,
            destinos=destinos_validos,
            motorista_nome=prefill.get("motorista_nome") or "",
            veiculo_modelo=prefill.get("veiculo_modelo") or "",
            veiculo_placa=prefill.get("veiculo_placa") or "",
            combustivel=prefill.get("combustivel") or "",
        )
        messages.success(request, "Termo criado e vinculado ao ofício/evento.")

    if next_url and url_has_allowed_host_and_scheme(next_url, allowed_hosts=[request.get_host()]):
        return redirect(next_url)
    return redirect("evento_guiado_etapa5", evento_id=evento.id)


@require_http_methods(["GET"])
def termos_autorizacao_lista(request):
    q = (request.GET.get("q") or "").strip()
    termos_qs = TermoAutorizacao.objects.all()
    if q:
        if q.isdigit():
            termos_qs = termos_qs.filter(id=int(q))
        else:
            termos_qs = termos_qs.none()

    paginator = Paginator(termos_qs, 10)
    page_obj = paginator.get_page(request.GET.get("page"))
    termos_rows: list[dict[str, object]] = []
    for termo in page_obj:
        destinos_labels = _resolve_termo_destinos_labels(termo.destinos or [])
        termos_rows.append(
            {
                "termo": termo,
                "nome": _build_termo_nome(destinos_labels),
                "periodo": _format_periodo_termo(
                    termo.data_inicio,
                    termo.data_fim,
                    termo.data_unica,
                ),
                "destinos": ", ".join(destinos_labels) if destinos_labels else "-",
            }
        )

    return render(
        request,
        "viagens/termos_autorizacao_list.html",
        {
            "q": q,
            "termos_rows": termos_rows,
            "page_obj": page_obj,
        },
    )


@require_GET
def termo_autorizacao_download_docx(request, termo_id: int):
    termo = get_object_or_404(TermoAutorizacao, id=termo_id)
    destinos_labels = _resolve_termo_destinos_labels(termo.destinos or [])
    if not destinos_labels:
        messages.error(request, "Termo sem destinos validos para gerar documento.")
        return redirect("termos_autorizacao_lista")

    try:
        buf = build_termo_autorizacao_payload_docx_bytes(
            data_inicio=termo.data_inicio,
            data_fim=termo.data_fim or termo.data_inicio,
            destinos=destinos_labels,
        )
    except Exception as exc:
        logger.exception("[termo-docx] falha na geracao do termo: termo_id=%s", termo.id)
        messages.error(request, f"Falha ao gerar termo DOCX. Detalhe: {exc}")
        return redirect("termos_autorizacao_lista")

    filename = f"{_build_termo_nome(destinos_labels)}.docx"
    resp = HttpResponse(
        buf.getvalue(),
        content_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )
    resp["Content-Disposition"] = f'attachment; filename="{filename}"'
    return resp


@require_GET
def termo_autorizacao_download_pdf(request, termo_id: int):
    termo = get_object_or_404(TermoAutorizacao, id=termo_id)
    destinos_labels = _resolve_termo_destinos_labels(termo.destinos or [])
    if not destinos_labels:
        messages.error(request, "Termo sem destinos validos para gerar PDF.")
        return redirect("termos_autorizacao_lista")

    try:
        docx_buf = build_termo_autorizacao_payload_docx_bytes(
            data_inicio=termo.data_inicio,
            data_fim=termo.data_fim or termo.data_inicio,
            destinos=destinos_labels,
        )
        pdf_bytes = docx_bytes_to_pdf_bytes(docx_buf.getvalue(), oficio_id=None)
    except DocxPdfConversionError as exc:
        messages.error(request, f"Falha ao gerar PDF. Baixe o DOCX. Detalhe: {exc}")
        return redirect("termo_autorizacao_download_docx", termo_id=termo.id)
    except Exception as exc:
        logger.exception("[termo-pdf] falha inesperada: termo_id=%s", termo.id)
        messages.error(request, f"Falha ao gerar PDF. Baixe o DOCX. Detalhe: {exc}")
        return redirect("termo_autorizacao_download_docx", termo_id=termo.id)

    response = HttpResponse(pdf_bytes, content_type="application/pdf")
    response["Content-Disposition"] = (
        f'attachment; filename="{_build_termo_nome(destinos_labels)}.pdf"'
    )
    return response


@require_http_methods(["POST"])
def oficio_excluir(request, oficio_id: int):
    oficio = get_object_or_404(Oficio, id=oficio_id)
    oficio.delete()
    messages.success(request, "Oficio excluido com sucesso.")
    return redirect("oficios_lista")


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

        form = ViajanteNormalizeForm(request.POST, instance=viajante)
        if form.is_valid():
            cargo = _resolver_cargo_nome(form.cleaned_data.get("cargo", ""))
            if not cargo:
                erros["cargo"] = "Informe o cargo."
            else:
                _ensure_cargo_exists(cargo)
                viajante = form.save(commit=False)
                viajante.cargo = cargo
                try:
                    viajante.save()
                except ValidationError as exc:
                    if hasattr(exc, "message_dict"):
                        for field, messages_list in exc.message_dict.items():
                            if messages_list:
                                erros[field] = messages_list[0]
                    else:
                        erros["cpf"] = "; ".join(exc.messages) or "Dados invalidos."
                    return render(
                        request,
                        "viagens/viajante_edit.html",
                        {
                            "viajante": viajante,
                            "erros": erros,
                            "cargo_choices": _get_cargo_choices(),
                        },
                    )
                return redirect("viajantes_lista")
        else:
            for field, messages_list in form.errors.items():
                if messages_list:
                    erros[field] = messages_list[0]
            viajante.nome = form.data.get("nome", viajante.nome)
            viajante.rg = form.data.get("rg", viajante.rg)
            viajante.cpf = form.data.get("cpf", viajante.cpf)
            viajante.telefone = form.data.get("telefone", viajante.telefone)
            viajante.cargo = form.data.get("cargo", viajante.cargo)

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
    if not tipo_destino_val:
        trechos_iniciais = list(
            oficio.trechos.select_related(
                "destino_estado",
                "destino_cidade",
            ).order_by("ordem", "id")
        )
        if trechos_iniciais:
            tipo_destino_val = infer_tipo_destino(trechos_iniciais)
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

        numeracao_form = OficioNumeracaoForm(
            {
                "oficio": request.POST.get("oficio", ""),
                "protocolo": request.POST.get("protocolo", ""),
            }
        )
        numeracao_form.is_valid()
        oficio_numero, oficio_ano, oficio_val = _parse_oficio_parts(
            numeracao_form.cleaned_data.get(
                "oficio", _formatar_oficio_numero(request.POST.get("oficio", ""))
            )
        )
        if oficio_numero is None:
            oficio_numero = oficio.numero
        if oficio_ano is None:
            oficio_ano = oficio.ano
        if not oficio_val:
            oficio_val = format_oficio_num(oficio_numero, oficio_ano)
        protocolo = numeracao_form.cleaned_data.get(
            "protocolo", _formatar_protocolo(request.POST.get("protocolo", ""))
        )
        assunto = request.POST.get("assunto", "").strip()
        placa = request.POST.get("placa", "").strip()
        placa_norm = _normalizar_placa(placa) if placa else ""
        modelo = request.POST.get("modelo", "").strip()
        combustivel = request.POST.get("combustivel", "").strip()
        motorista_fields_form = MotoristaTransporteForm(
            {
                "motorista_nome": request.POST.get("motorista_nome", ""),
                "motorista_oficio": request.POST.get("motorista_oficio", ""),
                "motorista_oficio_numero": request.POST.get("motorista_oficio_numero", ""),
                "motorista_oficio_ano": request.POST.get("motorista_oficio_ano", ""),
                "motorista_protocolo": request.POST.get("motorista_protocolo", ""),
            }
        )
        motorista_fields_form.is_valid()
        motorista_nome_manual = motorista_fields_form.cleaned_data.get(
            "motorista_nome",
            normalize_upper_text(request.POST.get("motorista_nome", "")),
        )
        motorista_obj = None
        if motorista_form.is_valid():
            motorista_obj = motorista_form.cleaned_data.get("motorista")
            motorista_preview = motorista_obj or motorista_preview
        motorista_oficio = motorista_fields_form.cleaned_data.get(
            "motorista_oficio",
            _formatar_oficio_numero(request.POST.get("motorista_oficio", "")),
        )
        motorista_oficio_numero = motorista_fields_form.cleaned_data.get(
            "motorista_oficio_numero",
            normalize_digits(request.POST.get("motorista_oficio_numero", "")),
        )
        motorista_oficio_ano = motorista_fields_form.cleaned_data.get(
            "motorista_oficio_ano",
            normalize_digits(request.POST.get("motorista_oficio_ano", "")),
        )
        motorista_protocolo = motorista_fields_form.cleaned_data.get(
            "motorista_protocolo",
            _formatar_protocolo(request.POST.get("motorista_protocolo", "")),
        )
        motivo = request.POST.get("motivo", "").strip()
        tipo_destino_val = (request.POST.get("tipo_destino") or "").strip().upper()
        retorno_saida_data_val = request.POST.get("retorno_saida_data", "").strip()
        retorno_saida_hora_val = request.POST.get("retorno_saida_hora", "").strip()
        retorno_chegada_data_val = request.POST.get("retorno_chegada_data", "").strip()
        retorno_chegada_hora_val = request.POST.get("retorno_chegada_hora", "").strip()
        valor_diarias_extenso_val = request.POST.get("valor_diarias_extenso", "").strip()
        custeio_tipo_val = _normalize_custeio_choice(request.POST.get("custeio_tipo") or request.POST.get("custos"))
        nome_instituicao_custeio = _normalize_nome_instituicao_custeio(
            custeio_tipo_val, request.POST.get("nome_instituicao_custeio", "")
        )
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

        if not protocolo:
            erros["protocolo"] = "Informe o protocolo."
        if motorista_carona:
            if not motorista_oficio_numero and not motorista_oficio:
                erros["motorista_oficio"] = "Informe o numero do oficio do motorista."
            if not motorista_oficio_ano and (motorista_oficio_numero or motorista_oficio):
                erros["motorista_oficio"] = "Informe o ano do oficio do motorista."
            if not motorista_protocolo:
                erros["motorista_protocolo"] = "Informe o protocolo do motorista."
        if not formset.is_valid():
            erros["trechos"] = "Revise os trechos do roteiro."

        if custeio_tipo_val == Oficio.CusteioTipoChoices.OUTRA_INSTITUICAO and not nome_instituicao_custeio:
            erros["nome_instituicao_custeio"] = "Informe a instituição de custeio."

        if not erros:
            forms_validas = [form for form in formset.forms if form.cleaned_data]
            if not forms_validas:
                erros["trechos"] = "Adicione ao menos um trecho para o roteiro."
            temp_trechos_edit: list[Trecho] = []
            for form in forms_validas:
                temp_trechos_edit.append(
                    Trecho(
                        destino_estado=form.cleaned_data.get("destino_estado"),
                        destino_cidade=form.cleaned_data.get("destino_cidade"),
                    )
                )
            if temp_trechos_edit:
                tipo_destino_val = infer_tipo_destino(temp_trechos_edit)
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
            trechos_serialized_calc = []
            for form in forms_validas:
                cleaned = form.cleaned_data
                trechos_serialized_calc.append(
                    {
                        "origem_estado": cleaned.get("origem_estado").sigla
                        if cleaned.get("origem_estado")
                        else "",
                        "origem_cidade": str(cleaned.get("origem_cidade").id)
                        if cleaned.get("origem_cidade")
                        else "",
                        "destino_estado": cleaned.get("destino_estado").sigla
                        if cleaned.get("destino_estado")
                        else "",
                        "destino_cidade": str(cleaned.get("destino_cidade").id)
                        if cleaned.get("destino_cidade")
                        else "",
                        "saida_data": cleaned.get("saida_data").isoformat()
                        if cleaned.get("saida_data")
                        else "",
                        "saida_hora": cleaned.get("saida_hora").strftime("%H:%M")
                        if cleaned.get("saida_hora")
                        else "",
                        "chegada_data": cleaned.get("chegada_data").isoformat()
                        if cleaned.get("chegada_data")
                        else "",
                        "chegada_hora": cleaned.get("chegada_hora").strftime("%H:%M")
                        if cleaned.get("chegada_hora")
                        else "",
                    }
                )

            resultado_diarias = _calculate_periodized_diarias_for_serialized_trechos(
                trechos_serialized_calc,
                retorno_chegada_data,
                retorno_chegada_hora,
                quantidade_servidores=len(servidores_ids),
            )

            quantidade_diarias, valor_diarias, valor_diarias_extenso_auto = _diarias_totais_fields(
                resultado_diarias
            )
            retorno_saida_cidade = _format_trecho_local(destino_cidade, destino_estado)
            retorno_chegada_cidade = _format_trecho_local(sede_cidade, sede_estado)
            oficio.oficio = oficio_val
            oficio.numero = oficio_numero
            oficio.ano = oficio_ano
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
            oficio.quantidade_diarias = quantidade_diarias
            oficio.valor_diarias = valor_diarias
            oficio.valor_diarias_extenso = (
                valor_diarias_extenso_val or valor_diarias_extenso_auto
            )
            oficio.placa = placa_norm or placa
            oficio.modelo = modelo
            oficio.combustivel = combustivel
            oficio.motorista = motorista_nome
            oficio.motorista_oficio = motorista_oficio
            oficio.motorista_oficio_numero = (
                int(motorista_oficio_numero) if motorista_oficio_numero else None
            )
            oficio.motorista_oficio_ano = (
                int(motorista_oficio_ano)
                if motorista_oficio_ano
                else (timezone.localdate().year if motorista_oficio_numero else None)
            )
            oficio.motorista_protocolo = motorista_protocolo
            oficio.motorista_carona = motorista_carona
            oficio.motorista_viajante = motorista_obj
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
    motorista_preview_payload = (
        _servidor_payload(motorista_preview)
        if isinstance(motorista_preview, Viajante)
        else motorista_preview
    )

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
            "motorista_preview": motorista_preview_payload,
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
            "oficios_referencia": [],
            "carona_oficio_referencia_id": "",
            "CURRENT_YEAR": timezone.localdate().year,
        },
    )

# viagens/views.py (adicione perto das outras views)


def _docx_http_response(payload: bytes, filename: str) -> HttpResponse:
    response = HttpResponse(
        payload,
        content_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )
    response["Content-Disposition"] = f'attachment; filename="{filename}"'
    return response


def _pdf_http_response(payload: bytes, filename: str) -> HttpResponse:
    response = HttpResponse(payload, content_type="application/pdf")
    response["Content-Disposition"] = f'attachment; filename="{filename}"'
    return response


def _ensure_acao_for_oficio(oficio: Oficio):
    return oficio.ensure_acao()


def _get_plano_trabalho(oficio: Oficio) -> PlanoTrabalho | None:
    acao = _ensure_acao_for_oficio(oficio)
    if acao is not None:
        try:
            return acao.plano_trabalho
        except PlanoTrabalho.DoesNotExist:
            pass
    try:
        return oficio.plano_trabalho
    except PlanoTrabalho.DoesNotExist:
        return None


def _get_ordem_servico(oficio: Oficio) -> OrdemServico | None:
    acao = _ensure_acao_for_oficio(oficio)
    if acao is not None:
        try:
            return acao.ordem_servico
        except OrdemServico.DoesNotExist:
            pass
    try:
        return oficio.ordem_servico
    except OrdemServico.DoesNotExist:
        return None


def _oficio_trechos(oficio: Oficio) -> list[Trecho]:
    return list(
        oficio.trechos.select_related(
            "origem_cidade",
            "origem_estado",
            "destino_cidade",
            "destino_estado",
        ).order_by("ordem", "id")
    )


def _resolve_periodo_oficio(oficio: Oficio) -> tuple[date, date]:
    trechos = _oficio_trechos(oficio)
    inicio_datas = [trecho.saida_data for trecho in trechos if trecho.saida_data]
    fim_datas = [
        trecho.chegada_data or trecho.saida_data
        for trecho in trechos
        if trecho.chegada_data or trecho.saida_data
    ]
    if oficio.retorno_saida_data:
        fim_datas.append(oficio.retorno_saida_data)
    if oficio.retorno_chegada_data:
        fim_datas.append(oficio.retorno_chegada_data)

    today = timezone.localdate()
    data_inicio = min(inicio_datas) if inicio_datas else today
    data_fim = max(fim_datas) if fim_datas else data_inicio
    if data_fim < data_inicio:
        data_fim = data_inicio
    return data_inicio, data_fim


def _resolve_local_oficio(oficio: Oficio) -> str:
    if oficio.cidade_sede and oficio.estado_sede:
        return f"{oficio.cidade_sede.nome}/{oficio.estado_sede.sigla}"
    if oficio.cidade_sede:
        return oficio.cidade_sede.nome
    if oficio.cidade_destino and oficio.estado_destino:
        return f"{oficio.cidade_destino.nome}/{oficio.estado_destino.sigla}"
    if oficio.cidade_destino:
        return oficio.cidade_destino.nome

    cfg = get_oficio_config()
    sede_default = getattr(cfg, "sede_cidade_default", None)
    if sede_default and getattr(sede_default, "estado", None):
        return f"{sede_default.nome}/{sede_default.estado.sigla}"
    return "Curitiba/PR"


def _resolve_destinos_oficio(oficio: Oficio) -> str:
    destinos: list[str] = []
    seen: set[str] = set()
    for trecho in _oficio_trechos(oficio):
        if trecho.destino_cidade and trecho.destino_estado:
            label = f"{trecho.destino_cidade.nome}/{trecho.destino_estado.sigla}"
        elif trecho.destino_cidade:
            label = trecho.destino_cidade.nome
        else:
            label = ""
        if label and label not in seen:
            seen.add(label)
            destinos.append(label)

    if not destinos:
        if oficio.cidade_destino and oficio.estado_destino:
            return f"{oficio.cidade_destino.nome}/{oficio.estado_destino.sigla}"
        if oficio.cidade_destino:
            return oficio.cidade_destino.nome
        return "-"
    if len(destinos) == 1:
        return destinos[0]
    if len(destinos) == 2:
        return f"{destinos[0]} e {destinos[1]}"
    return f"{', '.join(destinos[:-1])} e {destinos[-1]}"


def _resolve_assinante_defaults() -> tuple[str, str]:
    cfg = get_oficio_config()
    assinante = getattr(cfg, "assinante", None)
    if not assinante:
        return "", ""
    return (assinante.nome or "").strip(), (assinante.cargo or "").strip()


def _resolve_assinante_id_default() -> int | None:
    cfg = get_oficio_config()
    assinante = getattr(cfg, "assinante", None)
    if not assinante:
        return None
    return int(assinante.id)


def _resolve_assinante_justificativa_id_default() -> int | None:
    """Assinante padrão apenas para o gerador de justificativas (independente do assinante dos ofícios)."""
    cfg = get_oficio_config()
    assinante = getattr(cfg, "assinante_justificativa", None)
    if not assinante:
        return None
    return int(assinante.id)


def _plano_locais_default(oficio: Oficio, data_inicio: date) -> list[dict[str, str]]:
    locais: list[dict[str, str]] = []
    seen: set[tuple[str, str]] = set()
    for trecho in _oficio_trechos(oficio):
        label = ""
        if trecho.destino_cidade and trecho.destino_estado:
            label = f"{trecho.destino_cidade.nome}/{trecho.destino_estado.sigla}"
        elif trecho.destino_cidade:
            label = trecho.destino_cidade.nome
        elif trecho.destino_estado:
            label = trecho.destino_estado.sigla
        data_str = trecho.saida_data.isoformat() if trecho.saida_data else ""
        if not label:
            continue
        key = (data_str, label)
        if key in seen:
            continue
        seen.add(key)
        locais.append({"data": data_str, "local": label})

    if locais:
        return locais
    fallback_local = _resolve_local_oficio(oficio)
    fallback_data = data_inicio.isoformat() if data_inicio else ""
    return [{"data": fallback_data, "local": fallback_local}]


def _serialize_ordered_text_items(values: list[str]) -> str:
    return json.dumps([{"descricao": value} for value in values], ensure_ascii=False)


def _serialize_locais_items(values: list[dict[str, object]]) -> str:
    payload: list[dict[str, str]] = []
    for item in values:
        data_val = item.get("data")
        if isinstance(data_val, date):
            data_str = data_val.isoformat()
        else:
            data_str = str(data_val or "")
        payload.append(
            {
                "data": data_str,
                "local": str(item.get("local") or "").strip(),
            }
        )
    return json.dumps(payload, ensure_ascii=False)


def _plano_items_default_payload(oficio: Oficio, data_inicio: date) -> dict[str, str]:
    atividades_default_values: list[str] = []
    metas_default_values = metas_from_atividades(atividades_default_values)
    metas_default = _serialize_ordered_text_items(metas_default_values)
    atividades_default = _serialize_ordered_text_items(atividades_default_values)
    recursos_default = _serialize_ordered_text_items(
        [
            "Unidade movel da PCPR.",
        ]
    )
    locais_default = _serialize_locais_items(_plano_locais_default(oficio, data_inicio))
    return {
        "metas_json": metas_default,
        "atividades_json": atividades_default,
        "atividades_selecionadas": atividades_default_values,
        "recursos_json": recursos_default,
        "locais_json": locais_default,
    }


def _sync_plano_ordered_text_items(
    plano: PlanoTrabalho,
    *,
    model_cls,
    values: list[str],
) -> None:
    model_cls.objects.filter(plano=plano).delete()
    bulk = [
        model_cls(plano=plano, ordem=idx + 1, descricao=value)
        for idx, value in enumerate(values)
        if (value or "").strip()
    ]
    if bulk:
        model_cls.objects.bulk_create(bulk)


def _sync_plano_locais(plano: PlanoTrabalho, locais: list[dict[str, object]]) -> None:
    PlanoTrabalhoLocalAtuacao.objects.filter(plano=plano).delete()
    bulk = []
    for idx, item in enumerate(locais):
        local = str(item.get("local") or "").strip()
        if not local:
            continue
        bulk.append(
            PlanoTrabalhoLocalAtuacao(
                plano=plano,
                ordem=idx + 1,
                data=item.get("data") if isinstance(item.get("data"), date) else None,
                local=local,
            )
        )
    if bulk:
        PlanoTrabalhoLocalAtuacao.objects.bulk_create(bulk)


def _plano_items_payload_from_instance(plano: PlanoTrabalho) -> dict[str, str]:
    atividades = normalize_atividades_selecionadas(
        [item.descricao for item in plano.atividades.all().order_by("ordem", "id")]
    )
    metas = metas_from_atividades(atividades)
    recursos = [item.descricao for item in plano.recursos.all().order_by("ordem", "id")]
    locais = [
        {"data": item.data, "local": item.local}
        for item in plano.locais_atuacao.all().order_by("ordem", "id")
    ]
    data_inicio = plano.data_inicio or timezone.localdate()
    defaults = _plano_items_default_payload(plano.oficio, data_inicio)
    return {
        "metas_json": _serialize_ordered_text_items(metas) if metas else defaults["metas_json"],
        "atividades_json": _serialize_ordered_text_items(atividades)
        if atividades
        else defaults["atividades_json"],
        "atividades_selecionadas": atividades if atividades else defaults["atividades_selecionadas"],
        "recursos_json": _serialize_ordered_text_items(recursos)
        if recursos
        else defaults["recursos_json"],
        "locais_json": _serialize_locais_items(locais) if locais else defaults["locais_json"],
    }


def _plano_initial_data(oficio: Oficio) -> dict[str, object]:
    data_inicio, data_fim = _resolve_periodo_oficio(oficio)
    coordenador_nome, coordenador_cargo = _resolve_assinante_defaults()
    items_payload = _plano_items_default_payload(oficio, data_inicio)
    return {
        "ano": int(oficio.ano or timezone.localdate().year),
        "sigla_unidade": "ASCOM",
        "programa_projeto": "PCPR na Comunidade",
        "destino": _resolve_destinos_oficio(oficio),
        "solicitante": "",
        "contexto_solicitacao": "",
        "data_inicio": data_inicio,
        "data_fim": data_fim,
        "horario_atendimento": "das 09h as 17h",
        "efetivo_formatado": f"{int(oficio.viajantes.count() or 0)} servidores.",
        "estrutura_apoio": "",
        "quantidade_servidores": oficio.viajantes.count(),
        "composicao_diarias": "",
        "valor_unitario": "",
        "valor_total_calculado": "",
        "coordenador_plano": _resolve_assinante_id_default(),
        "possui_coordenador_municipal": "nao",
        "coordenador_nome": coordenador_nome,
        "coordenador_cargo": coordenador_cargo,
        "metas_json": items_payload["metas_json"],
        "atividades_json": items_payload["atividades_json"],
        "atividades_selecionadas": items_payload["atividades_selecionadas"],
        "recursos_json": items_payload["recursos_json"],
        "locais_json": items_payload["locais_json"],
    }


def _derive_sigla_unidade_from_config() -> str:
    cfg = get_oficio_config()
    unidade = " ".join((getattr(cfg, "unidade_nome", "") or "").split())
    if not unidade:
        return "ASCOM"
    if unidade.isupper() and len(unidade) <= 10 and " " not in unidade:
        return unidade
    tokens = [
        token
        for token in re.findall(r"[A-Za-z]+", unidade.upper())
        if token not in {"DE", "DA", "DO", "DOS", "DAS", "E"}
    ]
    if not tokens:
        return "ASCOM"
    sigla = "".join(token[0] for token in tokens[:8])
    return sigla or "ASCOM"


def _plano_default_destinos_payload(oficio: Oficio) -> list[dict[str, str]]:
    destinos: list[dict[str, str]] = []
    seen: set[tuple[str, str]] = set()
    for trecho in _oficio_trechos(oficio):
        uf = trecho.destino_estado.sigla if trecho.destino_estado else ""
        cidade = str(trecho.destino_cidade_id or "")
        if not uf and trecho.destino_cidade and trecho.destino_cidade.estado:
            uf = trecho.destino_cidade.estado.sigla
        if not cidade and trecho.destino_cidade:
            cidade = trecho.destino_cidade.nome
        if not uf and not cidade:
            continue
        key = (uf, cidade)
        if key in seen:
            continue
        seen.add(key)
        label = _format_trecho_local(trecho.destino_cidade, trecho.destino_estado)
        destinos.append({"uf": uf, "cidade": cidade, "label": label})

    if destinos:
        return destinos
    if oficio.cidade_destino and oficio.estado_destino:
        return [
            {
                "uf": oficio.estado_destino.sigla,
                "cidade": str(oficio.cidade_destino.id),
                "label": f"{oficio.cidade_destino.nome}/{oficio.estado_destino.sigla}",
            }
        ]
    sede_default_id = _get_sede_cidade_default_id()
    return [{"uf": "PR", "cidade": sede_default_id, "label": ""}]


def _sync_plano_destinos(plano: PlanoTrabalho, destinos_payload: list[dict[str, str]]) -> None:
    normalized = normalize_destinos_payload(destinos_payload)
    labels = destinos_labels(normalized)
    plano.destinos_json = normalized
    plano.destino = format_lista_portugues(labels)
    if labels:
        plano.local = labels[0]
    locais = [{"data": plano.data_inicio, "local": label} for label in labels]
    if locais:
        _sync_plano_locais(plano, locais)


def _ensure_plano_wizard_instance(oficio: Oficio) -> PlanoTrabalho:
    plano = _get_plano_trabalho(oficio)
    if plano:
        return plano

    data_inicio, data_fim = _resolve_periodo_oficio(oficio)
    destinos_payload = _plano_default_destinos_payload(oficio)
    destinos_norm = normalize_destinos_payload(destinos_payload)
    labels = destinos_labels(destinos_norm)
    efetivo_rows_default = _default_efetivo_rows_from_oficio(oficio)
    qtd_servidores = efetivo_total_servidores(efetivo_rows_default) or int(oficio.viajantes.count() or 0)
    coordenador_nome = DEFAULT_COORDENADOR_PLANO_NOME
    coordenador_cargo = DEFAULT_COORDENADOR_PLANO_CARGO

    with transaction.atomic():
        acao = _ensure_acao_for_oficio(oficio)
        plano = PlanoTrabalho.objects.create(
            oficio=oficio,
            acao=acao,
            numero=get_next_plano_num(int(oficio.ano or timezone.localdate().year)),
            ano=int(oficio.ano or timezone.localdate().year),
            sigla_unidade=_derive_sigla_unidade_from_config(),
            programa_projeto="",
            solicitantes_json=[],
            destino=format_lista_portugues(labels) or _resolve_destinos_oficio(oficio),
            destinos_json=destinos_norm,
            solicitante="",
            local=(labels[0] if labels else _resolve_local_oficio(oficio)),
            data_inicio=data_inicio,
            data_fim=data_fim,
            horario_inicio=parse_time("09:00"),
            horario_fim=parse_time("17:00"),
            horario_atendimento="das 09h as 17h",
            efetivo_json=efetivo_rows_default,
            efetivo_formatado=(format_total_servidores(qtd_servidores) if qtd_servidores > 0 else ""),
            efetivo_por_dia=qtd_servidores,
            quantidade_servidores=qtd_servidores,
            unidade_movel=False,
            estrutura_apoio="",
            composicao_diarias=(oficio.quantidade_diarias or "").strip() or "1 x 100%",
            valor_total=(oficio.valor_diarias or "").strip(),
            coordenador_plano=None,
            coordenador_nome=coordenador_nome,
            coordenador_cargo=coordenador_cargo,
            possui_coordenador_municipal=False,
        )
        _sync_plano_destinos(plano, destinos_norm)
        plano.save(update_fields=["destinos_json", "destino", "local", "updated_at"])
        if not plano.recursos.exists():
            PlanoTrabalhoRecurso.objects.create(
                plano=plano,
                ordem=1,
                descricao="Unidade movel da PCPR.",
            )
    return plano


def _plano_destinos_initial(plano: PlanoTrabalho, oficio: Oficio) -> list[dict[str, str]]:
    destinos_payload = (
        plano.destinos_json if isinstance(plano.destinos_json, list) and plano.destinos_json else []
    )
    if not destinos_payload:
        destinos_payload = _plano_default_destinos_payload(oficio)
    normalized = []
    for item in normalize_destinos_payload(destinos_payload):
        normalized.append({"uf": item.get("uf", ""), "cidade": item.get("cidade", "")})
    return _build_destinos_display(_normalize_destinos_for_wizard(normalized))


def _resolve_diarias_chegada_final(
    plano: PlanoTrabalho,
    oficio: Oficio,
    trechos: list[Trecho],
) -> tuple[date, time]:
    fallback_data = oficio.retorno_chegada_data or plano.data_fim or timezone.localdate()
    fallback_hora = oficio.retorno_chegada_hora or parse_time("18:00") or time.min

    candidate_dates = [fallback_data]
    if plano.data_fim:
        candidate_dates.append(plano.data_fim)
    for trecho in trechos:
        if trecho.chegada_data:
            candidate_dates.append(trecho.chegada_data)
        elif trecho.saida_data:
            candidate_dates.append(trecho.saida_data)
    chegada_data = max(candidate_dates)
    chegada_hora = fallback_hora

    last_saida = None
    for trecho in trechos:
        if not trecho.saida_data:
            continue
        saida_hora = trecho.saida_hora or time.min
        saida_dt = datetime.combine(trecho.saida_data, saida_hora)
        if last_saida is None or saida_dt > last_saida:
            last_saida = saida_dt

    chegada_final = datetime.combine(chegada_data, chegada_hora)
    if last_saida and chegada_final <= last_saida:
        chegada_final = last_saida + timedelta(hours=1)
        chegada_data = chegada_final.date()
        chegada_hora = chegada_final.time().replace(microsecond=0)
    return chegada_data, chegada_hora


def _plano_diarias_resultado_or_error(plano: PlanoTrabalho, oficio: Oficio) -> dict:
    trechos = _oficio_trechos(oficio)
    if not trechos:
        raise ValueError("Nao ha trechos validos para calcular as diarias do plano.")
    retorno_data, retorno_hora = _resolve_diarias_chegada_final(plano, oficio, trechos)
    return _calculate_periodized_diarias_for_trechos(
        trechos,
        retorno_data,
        retorno_hora,
        quantidade_servidores=int(plano.quantidade_servidores or 0),
    )


def _plano_diarias_resultado(plano: PlanoTrabalho, oficio: Oficio) -> dict | None:
    try:
        return _plano_diarias_resultado_or_error(plano, oficio)
    except ValueError:
        return None


def _plano_diarias_diagnostico(plano: PlanoTrabalho, oficio: Oficio) -> dict[str, object]:
    missing_requirements: list[str] = []
    actions: list[dict[str, str]] = []
    if not _oficio_trechos(oficio):
        missing_requirements.append("trechos")
        actions.append(
            {
                "label": "Cadastrar trechos no Oficio",
                "url": reverse("oficio_edit_step3", args=[oficio.id]),
            }
        )

    efetivo_rows = normalize_efetivo_payload(
        plano.efetivo_json if isinstance(plano.efetivo_json, list) else []
    )
    total_servidores = efetivo_total_servidores(efetivo_rows) or int(plano.quantidade_servidores or 0)
    if total_servidores < 1:
        missing_requirements.append("efetivo")
        actions.append(
            {
                "label": "Preencher Efetivo na Etapa 2",
                "url": reverse("plano_trabalho_step2", args=[oficio.id]),
            }
        )

    if not missing_requirements:
        message = ""
    elif len(missing_requirements) > 1:
        message = (
            "Para calcular as diarias, cadastre ao menos um trecho no Oficio "
            "e informe o efetivo na Etapa 2."
        )
    elif missing_requirements[0] == "trechos":
        message = "Cadastre ao menos um trecho no Oficio para calcular as diarias."
    else:
        message = "Preencha o efetivo na Etapa 2 para calcular as diarias."

    return {
        "missing_requirements": missing_requirements,
        "actions": actions,
        "message": message,
        "can_calculate": not missing_requirements,
    }


def _resolve_plano_diarias_state(plano: PlanoTrabalho, oficio: Oficio) -> tuple[dict, str, dict[str, object]]:
    diagnostico = _plano_diarias_diagnostico(plano, oficio)
    if not diagnostico["can_calculate"]:
        return normalize_diarias_resultado(None), str(diagnostico["message"] or ""), diagnostico
    try:
        resultado = _plano_diarias_resultado_or_error(plano, oficio)
    except ValueError as exc:
        diagnostico = {
            "missing_requirements": [],
            "actions": [],
            "message": str(exc),
            "can_calculate": False,
        }
        return normalize_diarias_resultado(None), str(exc), diagnostico
    return normalize_diarias_resultado(resultado), "", diagnostico


def _plano_legado_redirect_para_resumo(plano: PlanoTrabalho) -> bool:
    solicitantes = normalize_solicitantes(
        plano.solicitantes_json if isinstance(plano.solicitantes_json, list) else []
    )
    if not solicitantes:
        return False

    if not plano.data_inicio or not plano.data_fim:
        return False

    if not (
        (plano.horario_inicio and plano.horario_fim)
        or " ".join((plano.horario_atendimento or "").split())
    ):
        return False

    destinos = destinos_labels(plano.destinos_json if isinstance(plano.destinos_json, list) else [])
    if not destinos and not " ".join((plano.destino or "").split()):
        return False

    efetivo_rows = normalize_efetivo_payload(
        plano.efetivo_json if isinstance(plano.efetivo_json, list) else []
    )
    total_servidores = efetivo_total_servidores(efetivo_rows) or int(plano.quantidade_servidores or 0)
    if total_servidores <= 0:
        return False

    if not plano.atividades.exists() or not plano.metas.exists():
        return False
    if not plano.recursos.exists():
        return False

    if not plano.get_coordenador_administrativo_nome():
        return False
    if has_coordenador_municipal(solicitantes) and not plano.coordenador_municipal:
        return False

    if not " ".join((plano.composicao_diarias or "").split()):
        return False

    return True


def _plano_resumo_context(plano: PlanoTrabalho, oficio: Oficio) -> dict[str, object]:
    solicitantes = normalize_solicitantes(
        plano.solicitantes_json if isinstance(plano.solicitantes_json, list) else []
    )
    solicitantes_exibicao = formatar_solicitante_exibicao(
        solicitantes,
        nome_pcpr=plano.solicitante or "",
    )
    destinos = destinos_labels(plano.destinos_json if isinstance(plano.destinos_json, list) else [])
    efetivo_rows = normalize_efetivo_payload(
        plano.efetivo_json if isinstance(plano.efetivo_json, list) else []
    )
    efetivo_total = efetivo_total_servidores(efetivo_rows) or int(plano.quantidade_servidores or 0)
    efetivo_rows_preview = [
        {
            "cargo": row.get("cargo", ""),
            "quantidade": int(row.get("quantidade") or 0),
            "quantidade_label": format_total_servidores(int(row.get("quantidade") or 0)),
        }
        for row in efetivo_rows
    ]
    diarias = normalize_diarias_resultado(_plano_diarias_resultado(plano, oficio))
    recursos = [item.descricao for item in plano.recursos.all().order_by("ordem", "id")]
    atividades = [
        " ".join((item.descricao or "").split())
        for item in plano.atividades.all().order_by("ordem", "id")
        if " ".join((item.descricao or "").split())
    ]
    metas = [
        " ".join((item.descricao or "").split())
        for item in plano.metas.all().order_by("ordem", "id")
        if " ".join((item.descricao or "").split())
    ]
    if atividades and not metas:
        metas = metas_from_atividades(atividades)
    return {
        "plano": plano,
        "oficio": oficio,
        "numero_plano_formatado": f"{int(plano.numero or 0):02d}/{int(plano.ano or timezone.localdate().year)}",
        "solicitantes_exibicao": solicitantes_exibicao or "-",
        "destinos_exibicao": format_lista_portugues(destinos) or "-",
        "dias_evento_extenso": format_periodo_evento_extenso(plano.data_inicio, plano.data_fim),
        "horario_atendimento": (
            formatar_horario_intervalo(plano.horario_inicio, plano.horario_fim)
            or plano.horario_atendimento
            or "-"
        ),
        "efetivo_rows": efetivo_rows_preview,
        "efetivo_total": efetivo_total,
        "efetivo_total_label": format_total_servidores(efetivo_total),
        "unidade_movel": plano.unidade_movel,
        "coordenacao_formatada": build_coordenacao_formatada(plano),
        "atividades": atividades,
        "metas": metas,
        "diarias_resultado": diarias,
        "recursos": recursos,
        "data_resumo_extenso": format_data_extenso_br(timezone.localdate()),
    }


def _ordem_initial_data(oficio: Oficio) -> dict[str, object]:
    determinante_nome, determinante_cargo = _resolve_assinante_defaults()
    return {
        "ano": int(oficio.ano or timezone.localdate().year),
        "referencia": "Diligências",
        "determinante_nome": determinante_nome,
        "determinante_cargo": determinante_cargo,
        "finalidade": (oficio.motivo or "").strip(),
    }


def _pdf_unavailable_response(document_label: str, exc: Exception) -> HttpResponse:
    return HttpResponseBadRequest(
        f"PDF indisponivel para {document_label} neste ambiente. Gere o DOCX. Detalhe: {exc}"
    )


def _resolve_documentos_active_tab(
    raw_tab: str | None,
    *,
    tem_plano: bool,
) -> str:
    tab = (raw_tab or "").strip().lower() or "resumo"
    allowed_tabs = {"resumo", "oficio", "termo", "plano", "justificativa"}
    if not tem_plano:
        allowed_tabs.add("ordem")
    if tab not in allowed_tabs:
        return "resumo"
    return tab


def _get_justificativa_placeholders_safe() -> list[dict]:
    """Retorna lista de {key, label} para o formulário (labels amigáveis, sem XML)."""
    try:
        keys = get_justificativa_placeholders()
    except FileNotFoundError:
        return []
    return [
        {"key": k, "label": k.replace("_", " ").strip().title()}
        for k in keys
    ]


@require_GET
def oficio_documentos(request, oficio_id: int):
    from ..services.documentos_manager import build_documentos_status

    oficio = get_object_or_404(
        Oficio.objects.prefetch_related("viajantes", "trechos"),
        id=oficio_id,
    )
    plano_trabalho = _get_plano_trabalho(oficio)
    ordem_servico = _get_ordem_servico(oficio)
    tem_plano = plano_trabalho is not None
    tem_ordem = ordem_servico is not None
    active_tab = _resolve_documentos_active_tab(
        request.GET.get("tab"),
        tem_plano=tem_plano,
    )
    justificativa_placeholders = _get_justificativa_placeholders_safe()
    documentos_status = build_documentos_status(oficio)
    return render(
        request,
        "viagens/oficio_documentos.html",
        {
            "oficio": oficio,
            "plano_trabalho": plano_trabalho,
            "ordem_servico": ordem_servico,
            "active_tab": active_tab,
            "tem_plano": tem_plano,
            "tem_ordem": tem_ordem,
            "justificativa_placeholders": justificativa_placeholders,
            "documentos_status": documentos_status,
        },
    )


@require_http_methods(["POST"])
def oficio_documentos_gerar_todos(request, oficio_id: int):
    oficio = get_object_or_404(
        Oficio.objects.prefetch_related("viajantes", "trechos"),
        id=oficio_id,
    )
    try:
        docs = generate_all_documents(oficio, pdf_if_available=True)
    except MotoristaCaronaValidationError as exc:
        messages.error(request, str(exc))
        return redirect("oficio_edit_step2", oficio_id=oficio.id)
    except AssinaturaObrigatoriaError as exc:
        messages.error(request, str(exc))
        return redirect("config_oficio")
    except Exception as exc:
        logger.exception("[oficio-documentos] falha ao gerar documentos: oficio_id=%s", oficio.id)
        messages.error(request, f"Falha ao gerar documentos. Detalhe: {exc}")
        return redirect("oficio_documentos", oficio_id=oficio.id)

    messages.success(
        request,
        "Documentos gerados com sucesso: " + ", ".join(sorted(docs.keys())),
    )
    return redirect("oficio_documentos", oficio_id=oficio.id)


@require_GET
def oficio_documentos_fragment(request, oficio_id: int):
    """Retorna apenas o HTML dos cards de documentos (para drawer na lista de ofícios)."""
    from ..services.documentos_manager import build_documentos_status

    oficio = get_object_or_404(
        Oficio.objects.prefetch_related("viajantes", "trechos"),
        id=oficio_id,
    )
    documentos_status = build_documentos_status(oficio)
    return render(
        request,
        "viagens/partials/oficio_documentos_cards.html",
        {
            "oficio": oficio,
            "documentos_status": documentos_status,
            "tem_plano": documentos_status["plano"]["status"] == "ok",
            "is_fragment": True,
        },
    )


def _criar_oficio_base_documento() -> Oficio:
    return Oficio.objects.create(status=Oficio.Status.DRAFT)


@require_http_methods(["GET", "POST"])
def plano_trabalho_criar(request):
    oficio = _criar_oficio_base_documento()
    _ensure_plano_wizard_instance(oficio)
    messages.success(request, "Plano de trabalho criado. Continue o preenchimento na etapa 1.")
    return redirect("plano_trabalho_step1", oficio_id=oficio.id)


@require_GET
def planos_trabalho_list(request):
    ano_raw = (request.GET.get("ano") or "").strip()
    oficio_raw = (request.GET.get("oficio") or "").strip()
    protocolo_raw = (request.GET.get("protocolo") or "").strip()
    destino_raw = (request.GET.get("destino") or "").strip()

    params = request.GET.copy()
    params.pop("page", None)
    querystring = params.urlencode()

    oficios = (
        Oficio.objects.select_related(
            "cidade_destino",
            "estado_destino",
            "plano_trabalho",
        )
        .prefetch_related("trechos")
        .order_by("-created_at")
    )

    if ano_raw.isdigit():
        ano_val = int(ano_raw)
        oficios = oficios.filter(
            models.Q(plano_trabalho__ano=ano_val) | models.Q(ano=ano_val)
        )
    if oficio_raw:
        normalized = normalize_oficio_num(oficio_raw)
        oficios = oficios.filter(
            models.Q(oficio__icontains=oficio_raw)
            | models.Q(oficio__icontains=normalized)
            | models.Q(numero__in=[int(x) for x in re.findall(r"\d+", oficio_raw)] if re.findall(r"\d+", oficio_raw) else [])
        )
    if protocolo_raw:
        normalized = normalize_protocolo_num(protocolo_raw)
        oficios = oficios.filter(
            models.Q(protocolo__icontains=protocolo_raw)
            | models.Q(protocolo__icontains=normalized)
        )
    if destino_raw:
        oficios = oficios.filter(
            models.Q(cidade_destino__nome__icontains=destino_raw)
            | models.Q(destino__icontains=destino_raw)
            | models.Q(trechos__destino_cidade__nome__icontains=destino_raw)
        ).distinct()

    paginator = Paginator(oficios, 25)
    page_obj = paginator.get_page(request.GET.get("page"))
    rows: list[dict[str, object]] = []
    for oficio in page_obj:
        plano = _get_plano_trabalho(oficio)
        data_inicio, data_fim = _resolve_periodo_oficio(oficio)
        rows.append(
            {
                "oficio": oficio,
                "plano": plano,
                "status": "Gerado" if plano else "Pendente",
                "data_inicio": data_inicio,
                "data_fim": data_fim,
                "destinos": _resolve_destinos_oficio(oficio),
            }
        )

    return render(
        request,
        "viagens/planos_trabalho_list.html",
        {
            "rows": rows,
            "page_obj": page_obj,
            "querystring": querystring,
            "filters": {
                "ano": ano_raw,
                "oficio": oficio_raw,
                "protocolo": protocolo_raw,
                "destino": destino_raw,
            },
        },
    )


@require_http_methods(["GET", "POST"])
def plano_trabalho_editar(request, oficio_id: int):
    oficio = get_object_or_404(
        Oficio.objects.prefetch_related("trechos", "viajantes"),
        id=oficio_id,
    )
    plano = _get_plano_trabalho(oficio)
    if plano and _plano_legado_redirect_para_resumo(plano):
        return redirect("plano_trabalho_resumo", oficio_id=oficio.id)
    return redirect("plano_trabalho_step1", oficio_id=oficio.id)


def _plano_next_url(request) -> str:
    """Parâmetro next para retorno ao evento guiado (etapa 4)."""
    return (request.GET.get("next") or "").strip()


@require_http_methods(["GET", "POST"])
def plano_trabalho_step1(request, oficio_id: int):
    oficio = get_object_or_404(
        Oficio.objects.prefetch_related("trechos", "viajantes"),
        id=oficio_id,
    )
    plano = _ensure_plano_wizard_instance(oficio)
    next_url = _plano_next_url(request)
    estados = Estado.objects.order_by("sigla")
    destinos_post = _plano_destinos_initial(plano, oficio)
    erros: dict[str, str] = {}

    horario_inicio = plano.horario_inicio
    horario_fim = plano.horario_fim
    if (not horario_inicio or not horario_fim) and plano.horario_atendimento:
        horario_inicio, horario_fim = parse_horario_atendimento_intervalo(plano.horario_atendimento)

    initial = {
        "solicitantes": normalize_solicitantes(
            plano.solicitantes_json if isinstance(plano.solicitantes_json, list) else []
        ),
        "solicitante_pcpr_nome": plano.solicitante or "",
        "data_unica": bool(plano.data_inicio and plano.data_fim and plano.data_inicio == plano.data_fim),
        "data_inicio": plano.data_inicio,
        "data_fim": plano.data_fim,
        "horario_inicio": horario_inicio or parse_time("09:00"),
        "horario_fim": horario_fim or parse_time("17:00"),
    }

    if request.method == "POST":
        form = PlanoTrabalhoStep1Form(request.POST, initial=initial)
        _, _, destinos_raw = _serialize_sede_destinos_from_post(request.POST)
        destinos_validos = [
            destino
            for destino in destinos_raw
            if (destino.get("uf") or "").strip() and (destino.get("cidade") or "").strip()
        ]
        destinos_post = _build_destinos_display(_normalize_destinos_for_wizard(destinos_raw or [{}]))
        if not destinos_validos:
            erros["destinos"] = "Informe ao menos um destino."

        if form.is_valid() and not erros:
            destinos_payload: list[dict[str, str]] = []
            for destino in destinos_validos:
                uf = (destino.get("uf") or "").strip().upper()
                cidade = (destino.get("cidade") or "").strip()
                estado_obj = _resolve_estado(uf)
                cidade_obj = _resolve_cidade(cidade, estado=estado_obj)
                label = _format_trecho_local(cidade_obj, estado_obj)
                if not label:
                    label = f"{cidade}/{uf}" if cidade and uf else cidade or uf
                destinos_payload.append({"uf": uf, "cidade": cidade, "label": label})

            solicitantes = form.cleaned_data["solicitantes"]
            with transaction.atomic():
                plano.solicitantes_json = solicitantes
                plano.solicitante = form.cleaned_data["solicitante_pcpr_nome"]
                plano.data_inicio = form.cleaned_data["data_inicio"]
                plano.data_fim = form.cleaned_data["data_fim"]
                plano.horario_inicio = form.cleaned_data["horario_inicio"]
                plano.horario_fim = form.cleaned_data["horario_fim"]
                plano.horario_atendimento = form.cleaned_data["horario_atendimento"]
                plano.sigla_unidade = _derive_sigla_unidade_from_config()
                plano.possui_coordenador_municipal = False
                if not has_coordenador_municipal(solicitantes):
                    plano.coordenador_municipal = None
                _sync_plano_destinos(plano, destinos_payload)
                plano.save()

            target = reverse("plano_trabalho_step2", kwargs={"oficio_id": oficio.id})
            if next_url:
                target = f"{target}?next={quote(next_url)}"
            return redirect(target)
    else:
        form = PlanoTrabalhoStep1Form(initial=initial)

    destinos_order = ",".join(str(idx) for idx in range(len(destinos_post)))
    return render(
        request,
        "viagens/plano_trabalho_step1.html",
        {
            "oficio": oficio,
            "plano": plano,
            "form": form,
            "erros": erros,
            "estados": estados,
            "destinos": destinos_post,
            "destinos_total_forms": len(destinos_post),
            "destinos_order": destinos_order,
            "numero_plano_formatado": f"{int(plano.numero or 0):02d}/{int(plano.ano or timezone.localdate().year)}",
            "evento_return_url": next_url if next_url else None,
        },
    )


@require_http_methods(["GET", "POST"])
def plano_trabalho_step2(request, oficio_id: int):
    oficio = get_object_or_404(
        Oficio.objects.prefetch_related("trechos", "viajantes"),
        id=oficio_id,
    )
    plano = _ensure_plano_wizard_instance(oficio)
    solicitantes = normalize_solicitantes(
        plano.solicitantes_json if isinstance(plano.solicitantes_json, list) else []
    )
    permite_municipal = has_coordenador_municipal(solicitantes)
    _default_efetivo_rows_from_oficio(oficio)
    cargo_options = _cargo_options_for_plano_step2()
    efetivo_rows = normalize_efetivo_payload(
        plano.efetivo_json if isinstance(plano.efetivo_json, list) else []
    )
    if not efetivo_rows:
        efetivo_rows = _default_efetivo_rows_from_oficio(oficio)
    atividades = [
        " ".join((item.descricao or "").split())
        for item in plano.atividades.all().order_by("ordem", "id")
        if " ".join((item.descricao or "").split())
    ]
    metas = [
        " ".join((item.descricao or "").split())
        for item in plano.metas.all().order_by("ordem", "id")
        if " ".join((item.descricao or "").split())
    ]
    if atividades and not metas:
        metas = metas_from_atividades(atividades)
    efetivo_total_preview = efetivo_total_servidores(efetivo_rows) or int(plano.quantidade_servidores or 0)

    initial = {
        "efetivo_json": json.dumps(efetivo_rows, ensure_ascii=False),
        "atividades_json": _serialize_ordered_text_items(atividades),
        "metas_json": _serialize_ordered_text_items(metas),
        "unidade_movel": "sim" if plano.unidade_movel else "nao",
        "coordenador_plano": plano.coordenador_plano_id or "",
        "coordenador_plano_nome": plano.get_coordenador_administrativo_nome(),
        "coordenador_plano_cargo": plano.get_coordenador_administrativo_cargo(),
        "coordenador_municipal": plano.coordenador_municipal_id or "",
        "coordenador_municipal_nome": "",
        "coordenador_municipal_cargo": "",
        "coordenador_municipal_cidade": "",
    }

    if request.method == "POST":
        form = PlanoTrabalhoStep2Form(
            request.POST,
            permite_municipal=permite_municipal,
            initial=initial,
        )
        if form.is_valid():
            with transaction.atomic():
                efetivo_payload = form.parsed_efetivo
                total_servidores = efetivo_total_servidores(efetivo_payload)

                plano.efetivo_json = efetivo_payload
                plano.quantidade_servidores = total_servidores
                plano.efetivo_por_dia = total_servidores
                plano.efetivo_formatado = (
                    format_total_servidores(total_servidores) if total_servidores > 0 else ""
                )
                plano.unidade_movel = bool(form.cleaned_data["unidade_movel"])
                if plano.unidade_movel:
                    plano.estrutura_apoio = DEFAULT_UNIDADE_MOVEL_TEXTO
                else:
                    plano.estrutura_apoio = ""

                plano.coordenador_plano = form.cleaned_data.get("coordenador_plano")
                plano.coordenador_nome = form.cleaned_data["coordenador_plano_nome"]
                plano.coordenador_cargo = form.cleaned_data["coordenador_plano_cargo"]
                if plano.coordenador_plano:
                    plano.coordenador_nome = plano.coordenador_plano.nome
                    plano.coordenador_cargo = plano.coordenador_plano.cargo

                if permite_municipal:
                    coordenador_municipal = form.cleaned_data.get("coordenador_municipal")
                    if not coordenador_municipal:
                        nome = form.cleaned_data.get("coordenador_municipal_nome", "")
                        cargo = form.cleaned_data.get("coordenador_municipal_cargo", "")
                        cidade = form.cleaned_data.get("coordenador_municipal_cidade", "")
                        if nome and cargo and cidade:
                            coordenador_municipal = CoordenadorMunicipal.objects.filter(
                                nome__iexact=nome,
                                cargo__iexact=cargo,
                                cidade__iexact=cidade,
                            ).first()
                            if not coordenador_municipal:
                                coordenador_municipal = CoordenadorMunicipal.objects.create(
                                    nome=nome,
                                    cargo=cargo,
                                    cidade=cidade,
                                    ativo=True,
                                )
                    plano.coordenador_municipal = coordenador_municipal
                    plano.possui_coordenador_municipal = bool(coordenador_municipal)
                else:
                    plano.coordenador_municipal = None
                    plano.possui_coordenador_municipal = False
                plano.save()
                _sync_plano_ordered_text_items(
                    plano,
                    model_cls=PlanoTrabalhoAtividade,
                    values=form.parsed_atividades,
                )
                _sync_plano_ordered_text_items(
                    plano,
                    model_cls=PlanoTrabalhoMeta,
                    values=form.parsed_metas,
                )
            return redirect("plano_trabalho_step3", oficio_id=oficio.id)
    else:
        form = PlanoTrabalhoStep2Form(
            initial=initial,
            permite_municipal=permite_municipal,
        )

    coordenador_plano_options = [
        {
            "id": item["id"],
            "nome": " ".join((item.get("nome") or "").split()),
            "cargo": " ".join((item.get("cargo") or "").split()),
        }
        for item in form.fields["coordenador_plano"].queryset.values("id", "nome", "cargo")
    ]
    coordenador_municipal_options = [
        {
            "id": item["id"],
            "nome": " ".join((item.get("nome") or "").split()),
            "cargo": " ".join((item.get("cargo") or "").split()),
            "cidade": " ".join((item.get("cidade") or "").split()),
        }
        for item in form.fields["coordenador_municipal"].queryset.values("id", "nome", "cargo", "cidade")
    ]

    return render(
        request,
        "viagens/plano_trabalho_step2.html",
        {
            "oficio": oficio,
            "plano": plano,
            "form": form,
            "permite_municipal": permite_municipal,
            "cargo_options": cargo_options,
            "coordenador_plano_options": coordenador_plano_options,
            "coordenador_municipal_options": coordenador_municipal_options,
            "efetivo_total_preview": format_total_servidores(efetivo_total_preview),
            "numero_plano_formatado": f"{int(plano.numero or 0):02d}/{int(plano.ano or timezone.localdate().year)}",
        },
    )


@require_http_methods(["GET", "POST"])
def plano_trabalho_step3(request, oficio_id: int):
    oficio = get_object_or_404(
        Oficio.objects.prefetch_related("trechos", "viajantes"),
        id=oficio_id,
    )
    plano = _ensure_plano_wizard_instance(oficio)
    trechos_oficio = list(oficio.trechos.order_by("ordem", "id"))
    trechos_display = []
    for trecho in trechos_oficio:
        trechos_display.append(
            {
                "origem": str(trecho.origem_cidade or trecho.origem_estado or "-"),
                "destino": str(trecho.destino_cidade or trecho.destino_estado or "-"),
                "saida_data": trecho.saida_data.strftime("%d/%m/%Y")
                if trecho.saida_data
                else "—",
                "saida_hora": trecho.saida_hora.strftime("%H:%M")
                if trecho.saida_hora
                else "—",
                "chegada_data": trecho.chegada_data.strftime("%d/%m/%Y")
                if trecho.chegada_data
                else "—",
                "chegada_hora": trecho.chegada_hora.strftime("%H:%M")
                if trecho.chegada_hora
                else "—",
            }
        )
    diarias_resultado, diarias_erro_inicial, diarias_diagnostico = _resolve_plano_diarias_state(
        plano,
        oficio,
    )
    financeiro_diarias = derive_financeiro_diarias(diarias_resultado)

    recursos = [item.descricao for item in plano.recursos.all().order_by("ordem", "id")]
    if not recursos:
        recursos = ["Unidade movel da PCPR."]

    initial = {
        "composicao_diarias": plano.composicao_diarias or financeiro_diarias.get("diarias_por_servidor", ""),
        "valor_unitario": (
            f"{plano.valor_unitario:.2f}".replace(".", ",")
            if plano.valor_unitario
            else financeiro_diarias.get("valor_unitario", "")
        ),
        "valor_total_calculado": (
            f"{plano.valor_total_calculado:.2f}".replace(".", ",")
            if plano.valor_total_calculado
            else financeiro_diarias.get("total_geral", "")
        ),
        "recursos_json": json.dumps([{"descricao": item} for item in recursos], ensure_ascii=False),
    }

    if request.method == "POST":
        form = PlanoTrabalhoStep3Form(request.POST, initial=initial)
        if form.is_valid():
            if diarias_diagnostico["can_calculate"]:
                try:
                    diarias_resultado = _plano_diarias_resultado_or_error(plano, oficio)
                except ValueError as exc:
                    diarias_erro_inicial = str(exc)
                    diarias_diagnostico = {
                        "missing_requirements": [],
                        "actions": [],
                        "message": diarias_erro_inicial,
                        "can_calculate": False,
                    }
                    diarias_resultado = normalize_diarias_resultado(None)
                    financeiro_diarias = derive_financeiro_diarias(diarias_resultado)
                    form.add_error(None, diarias_erro_inicial)
                else:
                    diarias_resultado = normalize_diarias_resultado(diarias_resultado)
                    financeiro_diarias = derive_financeiro_diarias(diarias_resultado)
                    valor_unitario_decimal = (
                        parse_decimal_br(financeiro_diarias.get("valor_unitario"))
                        or parse_decimal_br(financeiro_diarias.get("valor_por_servidor"))
                    )
                    total_decimal = parse_decimal_br(financeiro_diarias.get("total_geral"))
                    with transaction.atomic():
                        plano.composicao_diarias = financeiro_diarias.get("diarias_por_servidor", "")
                        plano.valor_unitario = valor_unitario_decimal
                        plano.valor_total_calculado = total_decimal
                        plano.save()
                        _sync_plano_ordered_text_items(
                            plano,
                            model_cls=PlanoTrabalhoRecurso,
                            values=form.parsed_recursos,
                        )
                    return redirect("plano_trabalho_resumo", oficio_id=oficio.id)
            else:
                form.add_error(None, diarias_erro_inicial)
    else:
        form = PlanoTrabalhoStep3Form(initial=initial)

    return render(
        request,
        "viagens/plano_trabalho_step3.html",
        {
            "oficio": oficio,
            "plano": plano,
            "form": form,
            "trechos_oficio": trechos_oficio,
            "trechos_context": trechos_display,
            "trechos_display": trechos_display,
            "diarias_resultado": diarias_resultado,
            "financeiro_diarias": financeiro_diarias,
            "diarias_erro_inicial": diarias_erro_inicial,
            "diarias_acoes_iniciais": diarias_diagnostico.get("actions", []),
            "plano_diarias_calcular_url": reverse(
                "plano_trabalho_calcular_diarias",
                args=[oficio.id],
            ),
            "numero_plano_formatado": f"{int(plano.numero or 0):02d}/{int(plano.ano or timezone.localdate().year)}",
        },
    )


@require_GET
def plano_trabalho_calcular_diarias(request, oficio_id: int):
    oficio = get_object_or_404(
        Oficio.objects.prefetch_related("trechos", "viajantes"),
        id=oficio_id,
    )
    plano = _ensure_plano_wizard_instance(oficio)
    resultado, erro, diagnostico = _resolve_plano_diarias_state(plano, oficio)
    if erro:
        return JsonResponse(
            {
                "error": erro,
                "periodos": resultado.get("periodos", []),
                "totais": resultado.get("totais", {}),
                "financeiro": derive_financeiro_diarias(resultado),
                "actions": diagnostico.get("actions", []),
                "missing_requirements": diagnostico.get("missing_requirements", []),
            },
            status=400,
        )
    return JsonResponse(
        {
            "periodos": resultado.get("periodos", []),
            "totais": resultado.get("totais", {}),
            "financeiro": derive_financeiro_diarias(resultado),
            "actions": diagnostico.get("actions", []),
            "missing_requirements": diagnostico.get("missing_requirements", []),
        }
    )


@require_http_methods(["GET", "POST"])
def plano_trabalho_resumo(request, oficio_id: int):
    oficio = get_object_or_404(
        Oficio.objects.prefetch_related("trechos", "viajantes"),
        id=oficio_id,
    )
    plano = _ensure_plano_wizard_instance(oficio)
    if request.method == "POST":
        action = (request.POST.get("action") or "").strip()
        if action == "voltar":
            return redirect("plano_trabalho_step3", oficio_id=oficio.id)
        if action == "docx":
            return redirect("plano_trabalho_download_docx", oficio_id=oficio.id)
        if action == "pdf":
            return redirect("plano_trabalho_download_pdf", oficio_id=oficio.id)
        return redirect("oficio_documentos", oficio_id=oficio.id)

    context = _plano_resumo_context(plano, oficio)
    context.update(
        {
            "numero_plano_formatado": f"{int(plano.numero or 0):02d}/{int(plano.ano or timezone.localdate().year)}",
        }
    )
    return render(request, "viagens/plano_trabalho_step4.html", context)


@require_GET
def plano_trabalho_download_docx(request, oficio_id: int):
    return oficio_download_plano_trabalho_docx(request, oficio_id)


@require_GET
def plano_trabalho_download_pdf(request, oficio_id: int):
    return oficio_download_plano_trabalho_pdf(request, oficio_id)


@require_GET
def ordens_servico_list(request):
    q = (request.GET.get("q") or "").strip()
    params = request.GET.copy()
    params.pop("page", None)
    querystring = params.urlencode()

    oficios = (
        Oficio.objects.filter(plano_trabalho__isnull=True)
        .select_related("cidade_destino", "estado_destino")
        .prefetch_related("trechos")
        .order_by("-created_at")
    )
    if q:
        q_oficio = normalize_oficio_num(q)
        q_protocolo = normalize_protocolo_num(q)
        query = (
            models.Q(oficio__icontains=q)
            | models.Q(protocolo__icontains=q)
            | models.Q(destino__icontains=q)
            | models.Q(assunto__icontains=q)
            | models.Q(cidade_destino__nome__icontains=q)
            | models.Q(estado_destino__sigla__icontains=q)
        )
        if q_oficio:
            query |= models.Q(oficio__icontains=q_oficio)
        if q_protocolo:
            query |= models.Q(protocolo__icontains=q_protocolo)
        oficios = oficios.filter(query).distinct()

    rows: list[dict[str, object]] = []
    for oficio in oficios:
        ordem = _get_ordem_servico(oficio)
        data_inicio, data_fim = _resolve_periodo_oficio(oficio)
        rows.append(
            {
                "oficio": oficio,
                "ordem": ordem,
                "status": "Completa" if ordem else "Pendente",
                "data_inicio": data_inicio,
                "data_fim": data_fim,
                "destinos": _resolve_destinos_oficio(oficio),
            }
        )
    paginator = Paginator(rows, 25)
    page_obj = paginator.get_page(request.GET.get("page"))
    return render(
        request,
        "viagens/ordens_servico_list.html",
        {
            "rows": page_obj,
            "page_obj": page_obj,
            "q": q,
            "querystring": querystring,
        },
    )


@require_http_methods(["GET", "POST"])
def ordem_servico_criar(request):
    oficio = _criar_oficio_base_documento()
    ano = int(oficio.ano or timezone.localdate().year)
    initial = _ordem_initial_data(oficio)
    acao = _ensure_acao_for_oficio(oficio)
    OrdemServico.objects.create(
        oficio=oficio,
        acao=acao,
        numero=get_next_ordem_num(ano),
        ano=ano,
        referencia=str(initial.get("referencia") or "Diligencias"),
        determinante_nome=str(initial.get("determinante_nome") or ""),
        determinante_cargo=str(initial.get("determinante_cargo") or ""),
        finalidade=str(initial.get("finalidade") or ""),
    )
    messages.success(request, "Ordem de servico criada. Continue o preenchimento.")
    return redirect("ordem_servico_editar", oficio_id=oficio.id)


@require_http_methods(["GET", "POST"])
def ordem_servico_editar(request, oficio_id: int):
    oficio = get_object_or_404(
        Oficio.objects.prefetch_related("trechos", "viajantes"),
        id=oficio_id,
    )
    ordem = _get_ordem_servico(oficio)
    next_url = request.GET.get("next", "").strip()

    if request.method == "POST":
        form = OrdemServicoForm(request.POST, instance=ordem)
        if form.is_valid():
            ordem_obj = form.save(commit=False)
            if not ordem_obj.ano:
                ordem_obj.ano = int(oficio.ano or timezone.localdate().year)
            if not ordem_obj.numero:
                ordem_obj.numero = get_next_ordem_num(int(ordem_obj.ano))
            ordem_obj.oficio = oficio
            ordem_obj.save()
            messages.success(request, "Ordem de servico salva com sucesso.")
            if next_url and url_has_allowed_host_and_scheme(next_url, allowed_hosts=[request.get_host()]):
                return redirect(next_url)
            return redirect("ordens_servico_list")
    else:
        if ordem:
            form = OrdemServicoForm(instance=ordem)
        else:
            form = OrdemServicoForm(initial=_ordem_initial_data(oficio))

    return render(
        request,
        "viagens/ordem_servico_form.html",
        {
            "form": form,
            "oficio": oficio,
            "ordem": ordem,
            "evento_return_url": next_url if next_url else None,
        },
    )


@require_GET
def ordem_servico_download_docx(request, oficio_id: int):
    return oficio_download_ordem_servico_docx(request, oficio_id)


@require_GET
def ordem_servico_download_pdf(request, oficio_id: int):
    return oficio_download_ordem_servico_pdf(request, oficio_id)


@require_GET
def oficio_download_plano_trabalho_docx(request, oficio_id: int):
    oficio = get_object_or_404(
        Oficio.objects.prefetch_related("viajantes", "trechos"),
        id=oficio_id,
    )
    try:
        docx_bytes = build_plano_trabalho_docx_bytes(oficio).getvalue()
    except Exception as exc:
        logger.exception("[plano-docx] falha na geracao: oficio_id=%s", oficio.id)
        messages.error(request, f"Falha ao gerar Plano de Trabalho DOCX. Detalhe: {exc}")
        return redirect("oficio_documentos", oficio_id=oficio.id)

    filename = f"plano_trabalho_{oficio.numero_formatado or oficio.oficio or oficio.id}.docx"
    return _docx_http_response(docx_bytes, filename)


@require_GET
def oficio_download_plano_trabalho_pdf(request, oficio_id: int):
    oficio = get_object_or_404(
        Oficio.objects.prefetch_related("viajantes", "trechos"),
        id=oficio_id,
    )
    try:
        docx_bytes = build_plano_trabalho_docx_bytes(oficio).getvalue()
        pdf_bytes = docx_bytes_to_pdf_bytes(docx_bytes, oficio_id=oficio.id)
    except Exception as exc:
        logger.exception("[plano-pdf] falha na geracao: oficio_id=%s", oficio.id)
        return _pdf_unavailable_response("Plano de Trabalho", exc)

    filename = f"plano_trabalho_{oficio.numero_formatado or oficio.oficio or oficio.id}.pdf"
    return _pdf_http_response(pdf_bytes, filename)


@require_GET
def oficio_download_ordem_servico_docx(request, oficio_id: int):
    oficio = get_object_or_404(
        Oficio.objects.prefetch_related("viajantes", "trechos"),
        id=oficio_id,
    )
    try:
        docx_bytes = build_ordem_servico_docx_bytes(oficio).getvalue()
    except Exception as exc:
        logger.exception("[ordem-docx] falha na geracao: oficio_id=%s", oficio.id)
        messages.error(request, f"Falha ao gerar Ordem de Servico DOCX. Detalhe: {exc}")
        return redirect("oficio_documentos", oficio_id=oficio.id)

    filename = f"ordem_servico_{oficio.numero_formatado or oficio.oficio or oficio.id}.docx"
    return _docx_http_response(docx_bytes, filename)


@require_GET
def oficio_download_ordem_servico_pdf(request, oficio_id: int):
    oficio = get_object_or_404(
        Oficio.objects.prefetch_related("viajantes", "trechos"),
        id=oficio_id,
    )
    try:
        docx_bytes = build_ordem_servico_docx_bytes(oficio).getvalue()
        pdf_bytes = docx_bytes_to_pdf_bytes(docx_bytes, oficio_id=oficio.id)
    except Exception as exc:
        logger.exception("[ordem-pdf] falha na geracao: oficio_id=%s", oficio.id)
        return _pdf_unavailable_response("Ordem de Servico", exc)

    filename = f"ordem_servico_{oficio.numero_formatado or oficio.oficio or oficio.id}.pdf"
    return _pdf_http_response(pdf_bytes, filename)


@require_GET
def oficio_download_docx(request, oficio_id: int):
    oficio = get_object_or_404(
        Oficio.objects.prefetch_related("viajantes", "trechos"),
        id=oficio_id,
    )
    from viagens.services.justificativa_helpers import exige_justificativa, justificativa_preenchida
    if exige_justificativa(oficio) and not justificativa_preenchida(oficio):
        messages.warning(
            request,
            "É necessário preencher a justificativa (prazo menor que 10 dias) antes de gerar o ofício.",
        )
        return redirect("justificativa_oficio", oficio_id=oficio_id)

    try:
        buf = build_oficio_docx_bytes(oficio)
    except MotoristaCaronaValidationError as exc:
        messages.error(request, str(exc))
        return redirect("oficio_edit_step2", oficio_id=oficio.id)
    except AssinaturaObrigatoriaError as exc:
        messages.error(request, str(exc))
        return redirect("config_oficio")

    filename = f"oficio_{oficio.numero_formatado or oficio.oficio or oficio.id}.docx"
    resp = HttpResponse(
        buf.getvalue(),
        content_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )
    resp["Content-Disposition"] = f'attachment; filename="{filename}"'
    return resp


@require_GET
def oficio_download_termo_autorizacao(request, oficio_id: int):
    oficio = get_object_or_404(
        Oficio.objects.prefetch_related("trechos"),
        id=oficio_id,
    )

    try:
        buf = build_termo_autorizacao_docx_bytes(oficio)
    except Exception as exc:
        logger.exception(
            "[oficio-termo] falha na geracao do termo de autorizacao: oficio_id=%s",
            oficio.id,
        )
        messages.error(
            request,
            f"Falha ao gerar termo de autorizacao. Detalhe: {exc}",
        )
        return redirect("oficio_edit_step4", oficio_id=oficio.id)

    filename = f"termo_autorizacao_{oficio.numero_formatado or oficio.oficio or oficio.id}.docx"
    resp = HttpResponse(
        buf.getvalue(),
        content_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )
    resp["Content-Disposition"] = f'attachment; filename="{filename}"'
    return resp


@require_GET
def oficio_download_termo_autorizacao_pdf(request, oficio_id: int):
    oficio = get_object_or_404(
        Oficio.objects.prefetch_related("trechos"),
        id=oficio_id,
    )
    try:
        docx_bytes = build_termo_autorizacao_docx_bytes(oficio).getvalue()
        pdf_bytes = docx_bytes_to_pdf_bytes(docx_bytes, oficio_id=oficio.id)
    except Exception as exc:
        logger.exception(
            "[oficio-termo-pdf] falha na geracao do termo de autorizacao em PDF: oficio_id=%s",
            oficio.id,
        )
        return _pdf_unavailable_response("Termo de Autorizacao", exc)

    filename = f"termo_autorizacao_{oficio.numero_formatado or oficio.oficio or oficio.id}.pdf"
    return _pdf_http_response(pdf_bytes, filename)


JUSTIFICATIVA_FORM_PREFIX = "justificativa_"


def _redirect_oficio_documentos_justificativa(oficio_id: int):
    """Redireciona para a aba Justificativa da Central de Documentos (mantém formulário na aba)."""
    url = reverse("oficio_documentos", args=[oficio_id])
    return redirect(f"{url}?tab=justificativa")


@require_http_methods(["POST"])
def oficio_download_justificativa(request, oficio_id: int):
    get_object_or_404(Oficio, id=oficio_id)
    try:
        placeholders = get_justificativa_placeholders()
    except FileNotFoundError as exc:
        messages.error(request, str(exc))
        return _redirect_oficio_documentos_justificativa(oficio_id)
    prefix_len = len(JUSTIFICATIVA_FORM_PREFIX)
    mapping = {}
    for key in request.POST:
        if key.startswith(JUSTIFICATIVA_FORM_PREFIX):
            placeholder_name = key[prefix_len:]
            mapping[placeholder_name] = request.POST.get(key, "")
    for ph in placeholders:
        if ph not in mapping:
            mapping[ph] = ""
    try:
        buf = build_justificativa_docx_bytes(mapping)
    except Exception as exc:
        logger.exception(
            "[justificativa] falha na geracao: oficio_id=%s",
            oficio_id,
        )
        messages.error(request, f"Falha ao gerar justificativa. Detalhe: {exc}")
        return _redirect_oficio_documentos_justificativa(oficio_id)

    docx_bytes = buf.getvalue()
    gerar_pdf = request.POST.get("gerar_pdf")

    if gerar_pdf:
        try:
            pdf_bytes = docx_bytes_to_pdf_bytes(docx_bytes, oficio_id=oficio_id)
            filename = f"justificativa_{oficio_id}.pdf"
            return _pdf_http_response(pdf_bytes, filename)
        except DocxPdfConversionError as exc:
            messages.error(
                request,
                f"PDF indisponível neste ambiente. Baixe o DOCX. Detalhe: {exc}",
            )
            return _redirect_oficio_documentos_justificativa(oficio_id)
        except Exception as exc:
            logger.exception("[justificativa-pdf] falha: oficio_id=%s", oficio_id)
            messages.error(request, f"Falha ao gerar PDF. Baixe o DOCX. Detalhe: {exc}")
            return _redirect_oficio_documentos_justificativa(oficio_id)

    filename = f"justificativa_{oficio_id}.docx"
    return _docx_http_response(docx_bytes, filename)


def _justificativa_nome_cargo_from_config():
    """Retorna (nome, cargo) do assinante padrão de justificativas."""
    from ..services.text import title_case_pt
    cfg = get_oficio_config()
    assinante = getattr(cfg, "assinante_justificativa", None)
    if not assinante:
        return "", ""
    return (
        title_case_pt((assinante.nome or "").strip()),
        title_case_pt((assinante.cargo or "").strip()),
    )


@require_GET
def oficio_download_justificativa_docx(request, oficio_id: int):
    """Gera DOCX da justificativa a partir do texto salvo no ofício (GET)."""
    oficio = get_object_or_404(Oficio.objects.all(), id=oficio_id)
    texto = (getattr(oficio, "justificativa_texto", "") or "").strip()
    if not texto:
        messages.warning(request, "Preencha a justificativa antes de gerar o documento.")
        return _redirect_oficio_documentos_justificativa(oficio_id)
    nome, cargo = _justificativa_nome_cargo_from_config()
    mapping = build_justificativa_context_from_config(
        assinante_nome=nome,
        assinante_cargo=cargo,
        justificativa_texto=texto,
    )
    try:
        buf = build_justificativa_docx_bytes(mapping)
    except Exception as exc:
        logger.exception("[justificativa-docx] oficio_id=%s", oficio_id)
        messages.error(request, f"Falha ao gerar justificativa. Detalhe: {exc}")
        return _redirect_oficio_documentos_justificativa(oficio_id)
    filename = f"justificativa_{oficio_id}.docx"
    return _docx_http_response(buf.getvalue(), filename)


@require_GET
def oficio_download_justificativa_pdf(request, oficio_id: int):
    """Gera PDF da justificativa a partir do texto salvo no ofício (GET)."""
    oficio = get_object_or_404(Oficio.objects.all(), id=oficio_id)
    texto = (getattr(oficio, "justificativa_texto", "") or "").strip()
    if not texto:
        messages.warning(request, "Preencha a justificativa antes de gerar o documento.")
        return _redirect_oficio_documentos_justificativa(oficio_id)
    nome, cargo = _justificativa_nome_cargo_from_config()
    mapping = build_justificativa_context_from_config(
        assinante_nome=nome,
        assinante_cargo=cargo,
        justificativa_texto=texto,
    )
    try:
        buf = build_justificativa_docx_bytes(mapping)
        docx_bytes = buf.getvalue()
    except Exception as exc:
        logger.exception("[justificativa-pdf] oficio_id=%s", oficio_id)
        messages.error(request, f"Falha ao gerar justificativa. Detalhe: {exc}")
        return _redirect_oficio_documentos_justificativa(oficio_id)
    try:
        pdf_bytes = docx_bytes_to_pdf_bytes(docx_bytes, oficio_id=oficio_id)
        filename = f"justificativa_{oficio_id}.pdf"
        return _pdf_http_response(pdf_bytes, filename)
    except DocxPdfConversionError as exc:
        messages.error(request, f"PDF indisponível. Baixe o DOCX. Detalhe: {exc}")
        return _redirect_oficio_documentos_justificativa(oficio_id)
    except Exception as exc:
        logger.exception("[justificativa-pdf] oficio_id=%s", oficio_id)
        messages.error(request, f"Falha ao gerar PDF. Baixe o DOCX. Detalhe: {exc}")
        return _redirect_oficio_documentos_justificativa(oficio_id)


def oficio_download_pdf(request, oficio_id):
    oficio = get_object_or_404(
        Oficio.objects.prefetch_related("viajantes", "trechos"),
        id=oficio_id,
    )
    from viagens.services.justificativa_helpers import exige_justificativa, justificativa_preenchida
    if exige_justificativa(oficio) and not justificativa_preenchida(oficio):
        messages.warning(
            request,
            "É necessário preencher a justificativa (prazo menor que 10 dias) antes de gerar o ofício.",
        )
        return redirect("justificativa_oficio", oficio_id=oficio_id)

    try:
        _docx_bytes, pdf_bytes = build_oficio_docx_and_pdf_bytes(oficio)
    except MotoristaCaronaValidationError as exc:
        messages.error(request, str(exc))
        return redirect("oficio_edit_step2", oficio_id=oficio.id)
    except (AssinaturaObrigatoriaError, DocxPdfConversionError) as exc:
        messages.error(request, f"Falha ao gerar PDF. Baixe o DOCX. Detalhe: {exc}")
        return redirect("oficio_download_docx", oficio_id=oficio.id)
    except Exception as exc:
        logger.exception("[oficio-pdf] falha inesperada na geracao de PDF: oficio_id=%s", oficio.id)
        messages.error(request, f"Falha ao gerar PDF. Baixe o DOCX. Detalhe: {exc}")
        return redirect("oficio_download_docx", oficio_id=oficio.id)

    response = HttpResponse(pdf_bytes, content_type="application/pdf")
    response["Content-Disposition"] = (
        f'attachment; filename="oficio_{oficio.numero_formatado or oficio.oficio or oficio.id}.pdf"'
    )
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
    cpf_formatado = format_cpf(viajante.cpf)
    return {
        "id": viajante.id,
        "nome": viajante.nome,
        "rg": format_rg(viajante.rg),
        "cpf": cpf_formatado,
        "cargo": viajante.cargo,
        "telefone": format_phone(viajante.telefone),
        "label": viajante.nome,
    }


def _autocomplete_viajante_payload(viajante: Viajante) -> dict:
    cpf_formatado = format_cpf(viajante.cpf)
    text = viajante.nome
    if cpf_formatado:
        text = f"{text} - {cpf_formatado}"
    return {
        "id": viajante.id,
        "text": text,
        "label": text,
        "nome": viajante.nome,
        "cpf": cpf_formatado,
        "rg": format_rg(viajante.rg),
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
def coordenadores_municipais_api(request):
    q = request.GET.get("q", "").strip()
    cidade = request.GET.get("cidade", "").strip()
    queryset = CoordenadorMunicipal.objects.filter(ativo=True)
    if q:
        queryset = queryset.filter(
            Q(nome__icontains=q) | Q(cargo__icontains=q) | Q(cidade__icontains=q)
        )
    if cidade:
        queryset = queryset.filter(cidade__icontains=cidade)
    coordenadores = list(queryset.order_by("nome")[:50])
    payload = [
        {
            "id": item.id,
            "text": f"{item.nome} - {item.cargo} ({item.cidade})",
            "label": f"{item.nome} - {item.cargo} ({item.cidade})",
            "nome": item.nome,
            "cargo": item.cargo,
            "cidade": item.cidade,
        }
        for item in coordenadores
    ]
    return JsonResponse({"results": payload})


@require_GET
def assinantes_api(request):
    q = request.GET.get("q", "").strip()
    queryset = Viajante.objects.all()
    if q:
        viajantes = list(
            queryset.filter(Q(nome__icontains=q) | Q(cargo__icontains=q)).order_by("nome")[:20]
        )
    else:
        viajantes = list(queryset.order_by("-id")[:20])
        viajantes.sort(key=lambda item: item.nome or "")
    return JsonResponse({"results": [_autocomplete_viajante_payload(v) for v in viajantes]})


@require_GET
def cep_api(request, cep: str):
    digits = re.sub(r"\D", "", cep or "")
    if len(digits) != 8:
        return JsonResponse({"error": "CEP invalido."}, status=400)
    url = f"https://viacep.com.br/ws/{digits}/json/"
    try:
        with urlopen(url, timeout=3) as response:
            data = json.load(response)
    except (HTTPError, URLError, ValueError):
        return JsonResponse({"error": "CEP nao encontrado."}, status=404)
    if data.get("erro"):
        return JsonResponse({"error": "CEP nao encontrado."}, status=404)
    return JsonResponse(
        {
            "cep": f"{digits[:5]}-{digits[5:]}",
            "logradouro": data.get("logradouro", ""),
            "bairro": data.get("bairro", ""),
            "cidade": data.get("localidade", ""),
            "uf": data.get("uf", ""),
        }
    )


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
        erros = {}
        form = ViajanteNormalizeForm(request.POST)
        if form.is_valid():
            cargo = _resolver_cargo_nome(form.cleaned_data.get("cargo", ""))
            if cargo:
                _ensure_cargo_exists(cargo)
                viajante = form.save(commit=False)
                viajante.cargo = cargo
                try:
                    viajante.save()
                except ValidationError as exc:
                    if hasattr(exc, "message_dict"):
                        for field, messages_list in exc.message_dict.items():
                            if messages_list:
                                erros[field] = messages_list[0]
                    else:
                        erros["cpf"] = "; ".join(exc.messages) or "Dados invalidos."
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
                return JsonResponse({"success": True, "item": _servidor_payload(viajante)})

        if not erros:
            for field, messages_list in form.errors.items():
                if messages_list:
                    erros[field] = messages_list[0]

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

@csrf_exempt
def validacao_resultado(request):
    if request.method == "POST":
        token = request.headers.get("X-Api-Token", "")
        if not VALIDACAO_TOKEN or token != VALIDACAO_TOKEN:
            return JsonResponse({"erro": "Não autorizado."}, status=401)
    if request.method != "POST":
        return JsonResponse({"error": "method not allowed"}, status=405)

    data = json.loads(request.body.decode("utf-8"))
    # aqui você salva no banco ou atualiza a validação
    # exemplo: oficio_id = data["oficio_id"]; status = data["status"]

    return JsonResponse({"ok": True})
