from __future__ import annotations

from datetime import date
from decimal import Decimal, InvalidOperation
from io import BytesIO
from pathlib import Path

from django.conf import settings
from django.utils import timezone
from docx import Document as DocxFactory

from viagens.models import (
    Oficio,
    PlanoTrabalho,
    PlanoTrabalhoAtividade,
    PlanoTrabalhoLocalAtuacao,
    PlanoTrabalhoMeta,
    PlanoTrabalhoRecurso,
    Trecho,
    get_next_plano_num,
)
from viagens.services.oficio_config import get_oficio_config
from viagens.documents.document import (
    _find_unresolved_placeholders,
    safe_replace_placeholders,
)
from viagens.services.plano_trabalho import (
    ATIVIDADES_ORDEM_FIXA,
    META_POR_ATIVIDADE,
    build_plano_placeholders,
    validate_required_placeholders,
)

PLANO_TEMPLATE_FILENAME = "modelo_plano_de_trabalho.docx"


def _parse_decimal(value) -> Decimal | None:
    if value in (None, ""):
        return None
    try:
        normalized = str(value).strip().replace(".", "").replace(",", ".")
        return Decimal(normalized)
    except (InvalidOperation, TypeError, ValueError):
        return None


def _resolve_periodo(oficio: Oficio, trechos: list[Trecho]) -> tuple[date, date]:
    datas_inicio = [trecho.saida_data for trecho in trechos if trecho.saida_data]
    datas_fim = [
        trecho.chegada_data or trecho.saida_data
        for trecho in trechos
        if trecho.chegada_data or trecho.saida_data
    ]
    if oficio.retorno_saida_data:
        datas_fim.append(oficio.retorno_saida_data)
    if oficio.retorno_chegada_data:
        datas_fim.append(oficio.retorno_chegada_data)

    hoje = timezone.localdate()
    data_inicio = min(datas_inicio) if datas_inicio else hoje
    data_fim = max(datas_fim) if datas_fim else data_inicio
    if data_fim < data_inicio:
        data_fim = data_inicio
    return data_inicio, data_fim


def _resolve_destino(oficio: Oficio, trechos: list[Trecho]) -> str:
    labels: list[str] = []
    seen: set[str] = set()
    for trecho in trechos:
        if trecho.destino_cidade and trecho.destino_estado:
            label = f"{trecho.destino_cidade.nome}/{trecho.destino_estado.sigla}"
        elif trecho.destino_cidade:
            label = trecho.destino_cidade.nome
        elif trecho.destino_estado:
            label = trecho.destino_estado.sigla
        else:
            label = ""
        if label and label not in seen:
            seen.add(label)
            labels.append(label)
    if not labels:
        if oficio.cidade_destino and oficio.estado_destino:
            return f"{oficio.cidade_destino.nome}/{oficio.estado_destino.sigla}"
        if oficio.cidade_destino:
            return oficio.cidade_destino.nome
        return ""
    if len(labels) == 1:
        return labels[0]
    return ", ".join(labels[:-1]) + f" e {labels[-1]}"


def _resolve_local(oficio: Oficio, trechos: list[Trecho]) -> str:
    if oficio.cidade_sede and oficio.estado_sede:
        return f"{oficio.cidade_sede.nome}/{oficio.estado_sede.sigla}"
    if oficio.cidade_sede:
        return oficio.cidade_sede.nome
    for trecho in trechos:
        if trecho.destino_cidade and trecho.destino_estado:
            return f"{trecho.destino_cidade.nome}/{trecho.destino_estado.sigla}"
        if trecho.destino_cidade:
            return trecho.destino_cidade.nome
    cfg = get_oficio_config()
    if cfg.sede_cidade_default and cfg.sede_cidade_default.estado:
        return f"{cfg.sede_cidade_default.nome}/{cfg.sede_cidade_default.estado.sigla}"
    return "Curitiba/PR"


def _ensure_plano_trabalho(oficio: Oficio, trechos: list[Trecho]) -> PlanoTrabalho:
    try:
        return oficio.plano_trabalho
    except PlanoTrabalho.DoesNotExist:
        pass

    cfg = get_oficio_config()
    ano = int(oficio.ano or timezone.localdate().year)
    data_inicio, data_fim = _resolve_periodo(oficio, trechos)
    destino = _resolve_destino(oficio, trechos)
    qtd_servidores = int(oficio.viajantes.count() or 0)
    valor_total_legacy = _parse_decimal(oficio.valor_diarias)

    plano = PlanoTrabalho.objects.create(
        oficio=oficio,
        numero=get_next_plano_num(ano),
        ano=ano,
        sigla_unidade="ASCOM",
        programa_projeto="PCPR na Comunidade",
        destino=destino,
        solicitante="Demanda institucional",
        local=_resolve_local(oficio, trechos),
        data_inicio=data_inicio,
        data_fim=data_fim,
        horario_atendimento="das 09h as 17h",
        efetivo_formatado=f"{qtd_servidores} servidores.",
        efetivo_por_dia=qtd_servidores,
        quantidade_servidores=qtd_servidores,
        composicao_diarias=(oficio.quantidade_diarias or "").strip() or "1 x 100%",
        valor_total_calculado=valor_total_legacy,
        valor_unitario=valor_total_legacy,
        coordenador_plano=getattr(cfg, "assinante", None),
        coordenador_nome=(cfg.assinante.nome if getattr(cfg, "assinante", None) else ""),
        coordenador_cargo=(cfg.assinante.cargo if getattr(cfg, "assinante", None) else ""),
    )
    atividade_padrao = ATIVIDADES_ORDEM_FIXA[0]
    meta_padrao = META_POR_ATIVIDADE[atividade_padrao]
    PlanoTrabalhoMeta.objects.create(
        plano=plano,
        ordem=1,
        descricao=meta_padrao,
    )
    PlanoTrabalhoAtividade.objects.create(
        plano=plano,
        ordem=1,
        descricao=atividade_padrao,
    )
    PlanoTrabalhoRecurso.objects.create(
        plano=plano,
        ordem=1,
        descricao="Unidade movel da PCPR.",
    )
    PlanoTrabalhoLocalAtuacao.objects.create(
        plano=plano,
        ordem=1,
        data=data_inicio,
        local=destino or _resolve_local(oficio, trechos),
    )
    return plano


def _resolve_plano_template_path() -> Path:
    return Path(settings.BASE_DIR) / "viagens" / "documents" / PLANO_TEMPLATE_FILENAME


def build_plano_trabalho_docx_bytes(oficio: Oficio) -> BytesIO:
    trechos = list(
        oficio.trechos.select_related(
            "origem_cidade",
            "origem_estado",
            "destino_cidade",
            "destino_estado",
        ).order_by("ordem", "id")
    )
    plano = _ensure_plano_trabalho(oficio, trechos)
    cfg = get_oficio_config()
    placeholders = build_plano_placeholders(plano, oficio, cfg)
    missing = validate_required_placeholders(placeholders)
    if missing:
        raise ValueError(
            "Plano de trabalho incompleto. Campos obrigatorios ausentes: "
            + ", ".join(missing)
        )
    template_path = _resolve_plano_template_path()
    if not template_path.exists():
        raise FileNotFoundError(
            f"Template do plano nao encontrado: {template_path}"
        )
    doc = DocxFactory(str(template_path))
    safe_replace_placeholders(doc, placeholders)

    buf = BytesIO()
    doc.save(buf)
    leftovers = _find_unresolved_placeholders(buf.getvalue())
    if leftovers:
        raise ValueError(
            "Placeholders nao substituidos no plano DOCX: " + ", ".join(sorted(leftovers))
        )
    buf.seek(0)
    return buf
