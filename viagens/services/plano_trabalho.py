from __future__ import annotations

import re
from datetime import date, time
from decimal import Decimal, InvalidOperation
from typing import Iterable

from django.utils import timezone
from django.utils.dateparse import parse_time

from viagens.models import (
    Oficio,
    OficioConfig,
    PlanoTrabalho,
    PlanoTrabalhoAtividade,
    PlanoTrabalhoLocalAtuacao,
    PlanoTrabalhoMeta,
    PlanoTrabalhoRecurso,
)
from viagens.services.oficio_helpers import valor_por_extenso_ptbr
from viagens.services.text import title_case_pt

MESES_PTBR = {
    1: "janeiro",
    2: "fevereiro",
    3: "março",
    4: "abril",
    5: "maio",
    6: "junho",
    7: "julho",
    8: "agosto",
    9: "setembro",
    10: "outubro",
    11: "novembro",
    12: "dezembro",
}

ATIVIDADE_META_PAIRS: list[tuple[str, str]] = [
    (
        "Confecção da Carteira de Identidade Nacional (CIN)",
        "ampliar o acesso ao documento oficial de identificação civil, garantindo cidadania e inclusão social à população atendida.",
    ),
    (
        "Registro de Boletins de Ocorrência",
        "possibilitar o atendimento imediato de demandas policiais, promovendo orientação e formalização de ocorrências no próprio evento.",
    ),
    (
        "Emissão de Atestado de Antecedentes Criminais",
        "facilitar a obtenção do documento, contribuindo para fins trabalhistas e demais necessidades legais dos cidadãos.",
    ),
    (
        "Palestras e orientações preventivas",
        "desenvolver ações educativas voltadas à prevenção de crimes, conscientização sobre segurança pública e fortalecimento do vínculo comunitário.",
    ),
    (
        "Atividades lúdicas e educativas para crianças",
        "promover aproximação institucional de forma didática, incentivando a cultura de respeito às leis e à cidadania desde a infância.",
    ),
    (
        "Apresentação do trabalho do Núcleo de Operações com Cães (NOC)",
        "demonstrar as atividades operacionais desenvolvidas pela unidade especializada da Polícia Civil do Paraná, evidenciando técnicas e capacidades institucionais.",
    ),
    (
        "Exposição de material tático",
        "apresentar equipamentos utilizados nas atividades policiais, proporcionando transparência e conhecimento sobre os recursos empregados pela instituição.",
    ),
    (
        "Exposição da atividade de perícia papiloscópica",
        "demonstrar os procedimentos técnicos de identificação humana, ressaltando a importância da papiloscopia na investigação criminal e na identificação civil.",
    ),
    (
        "Exposição de viaturas antigas e modernas",
        "apresentar a evolução histórica e tecnológica dos veículos operacionais da instituição.",
    ),
    (
        "Apresentação da banda institucional",
        "fortalecer a integração com a comunidade por meio de atividade cultural representativa da instituição.",
    ),
    (
        "Unidade móvel (ônibus ou caminhão)",
        "viabilizar a prestação descentralizada dos serviços acima descritos, assegurando estrutura adequada para atendimento ao público.",
    ),
]

ATIVIDADES_ORDEM_FIXA: list[str] = [atividade for atividade, _ in ATIVIDADE_META_PAIRS]
META_POR_ATIVIDADE: dict[str, str] = dict(ATIVIDADE_META_PAIRS)

REQUIRED_PLACEHOLDERS = (
    "divisao",
    "unidade",
    "numero_plano",
    "ano_plano",
    "sigla_unidade",
    "dias_evento_extenso",
    "locais_formatado",
    "destino",
    "solicitante",
    "horario_atendimento",
    "quantidade_de_servidores",
    "atividades_formatada",
    "metas_formatadas",
    "valor_total",
    "composicao_diarias",
    "valor_unitario",
    "recursos_formatados",
    "coordenacao_formatada",
    "sede",
    "data_extenso",
    "nome_chefia",
    "cargo_chefia",
)

SOLICITANTE_PCPR = "PCPR na Comunidade"
SOLICITANTE_PARANA_ACAO = "Parana em Acao"
SOLICITANTE_JUSTICA_BAIRRO = "Justica no Bairro"
SOLICITANTES_ORDEM_FIXA = (
    SOLICITANTE_PCPR,
    SOLICITANTE_PARANA_ACAO,
    SOLICITANTE_JUSTICA_BAIRRO,
)
DEFAULT_COORDENADOR_PLANO_NOME = "JULIANA VILLELA DE BARROS"
DEFAULT_COORDENADOR_PLANO_CARGO = "Coordenadora Administrativa"
DEFAULT_UNIDADE_MOVEL_TEXTO = (
    "Unidade movel da PCPR equipada para atendimento e confeccao de documentos."
)


def format_data_extenso_br(value: date | None) -> str:
    if not value:
        return ""
    mes = MESES_PTBR.get(value.month, str(value.month))
    return f"{value.day} de {mes} de {value.year}"


def format_data_curta_br(value: date | None) -> str:
    if not value:
        return ""
    return value.strftime("%d/%m/%Y")


def format_periodo_evento_extenso(data_inicio: date | None, data_fim: date | None) -> str:
    if not data_inicio and not data_fim:
        return ""
    start = data_inicio or data_fim
    end = data_fim or data_inicio
    if not start or not end:
        return ""
    if start == end:
        return format_data_extenso_br(start)
    if start.year == end.year and start.month == end.month:
        mes = MESES_PTBR.get(start.month, str(start.month))
        return f"de {start.day} a {end.day} de {mes} de {start.year}"
    if start.year == end.year:
        mes_inicio = MESES_PTBR.get(start.month, str(start.month))
        mes_fim = MESES_PTBR.get(end.month, str(end.month))
        return f"de {start.day} de {mes_inicio} a {end.day} de {mes_fim} de {start.year}"
    mes_inicio = MESES_PTBR.get(start.month, str(start.month))
    mes_fim = MESES_PTBR.get(end.month, str(end.month))
    return f"de {start.day} de {mes_inicio} de {start.year} a {end.day} de {mes_fim} de {end.year}"


def format_periodo_extenso(data_inicio: date | None, data_fim: date | None) -> str:
    return format_periodo_evento_extenso(data_inicio, data_fim)


def _to_decimal(value) -> Decimal | None:
    if value in (None, ""):
        return None
    if isinstance(value, Decimal):
        return value
    try:
        normalized = str(value).strip().replace("R$", "").replace("r$", "").replace(" ", "")
        if "," in normalized and "." in normalized:
            if normalized.rfind(",") > normalized.rfind("."):
                # 1.234,56
                normalized = normalized.replace(".", "").replace(",", ".")
            else:
                # 1,234.56
                normalized = normalized.replace(",", "")
        elif "," in normalized:
            # 1234,56
            normalized = normalized.replace(".", "").replace(",", ".")
        else:
            # 1234.56
            normalized = normalized.replace(",", "")
        return Decimal(normalized)
    except (InvalidOperation, TypeError, ValueError):
        return None


def format_monetario_br(value) -> str:
    decimal_value = _to_decimal(value)
    if decimal_value is None:
        return ""
    return f"R$ {decimal_value:.2f}".replace(".", ",")


def format_valor_extenso(value) -> str:
    decimal_value = _to_decimal(value)
    if decimal_value is None:
        return ""
    return valor_por_extenso_ptbr(decimal_value)


def _ordered_descriptions(items: Iterable) -> list[str]:
    values: list[str] = []
    for item in items:
        raw = getattr(item, "descricao", "")
        text = " ".join(str(raw or "").split())
        if text:
            values.append(text)
    return values


def format_lista_bullets(values: list[str], *, bullet: str = "-") -> str:
    cleaned = [" ".join(str(value).split()) for value in values if " ".join(str(value).split())]
    if not cleaned:
        return "-"
    return "\n".join(f"{bullet} {value}" for value in cleaned)


def format_lista_portugues(values: list[str]) -> str:
    cleaned = [" ".join(str(value).split()) for value in values if " ".join(str(value).split())]
    if not cleaned:
        return ""
    if len(cleaned) == 1:
        return cleaned[0]
    if len(cleaned) == 2:
        return f"{cleaned[0]} e {cleaned[1]}"
    return f"{', '.join(cleaned[:-1])} e {cleaned[-1]}"


def normalize_solicitantes(values: list[str] | tuple[str, ...] | None) -> list[str]:
    selected = {str(item).strip() for item in (values or []) if str(item).strip()}
    return [item for item in SOLICITANTES_ORDEM_FIXA if item in selected]


def permite_coordenador_municipal(solicitantes: list[str] | tuple[str, ...] | None) -> bool:
    return SOLICITANTE_PCPR in normalize_solicitantes(solicitantes)


def formatar_solicitante_exibicao(
    solicitantes: list[str] | tuple[str, ...] | None,
    *,
    nome_pcpr: str = "",
) -> str:
    selecionados = normalize_solicitantes(solicitantes)
    nome_pcpr_limpo = " ".join((nome_pcpr or "").split())
    valores: list[str] = []
    for item in selecionados:
        if item == SOLICITANTE_PCPR and nome_pcpr_limpo:
            valores.append(f"{item} ({nome_pcpr_limpo})")
        else:
            valores.append(item)
    return format_lista_portugues(valores)


def normalize_destinos_payload(raw_values: list[dict] | None) -> list[dict[str, str]]:
    destinos: list[dict[str, str]] = []
    seen: set[tuple[str, str, str]] = set()
    for item in raw_values or []:
        if not isinstance(item, dict):
            continue
        uf = " ".join(str(item.get("uf", "")).split()).upper()
        cidade = " ".join(str(item.get("cidade", "")).split())
        label = " ".join(str(item.get("label", "")).split())
        if not label:
            if cidade and uf:
                label = f"{cidade}/{uf}"
            else:
                label = cidade or uf
        if not label:
            continue
        key = (uf, cidade.casefold(), label.casefold())
        if key in seen:
            continue
        seen.add(key)
        destinos.append({"uf": uf, "cidade": cidade, "label": label})
    return destinos


def destinos_labels(destinos_payload: list[dict] | None) -> list[str]:
    labels: list[str] = []
    for item in normalize_destinos_payload(destinos_payload):
        label = " ".join((item.get("label") or "").split())
        if label:
            labels.append(label)
    return labels


def parse_horario_atendimento_intervalo(raw: str) -> tuple[time | None, time | None]:
    text = " ".join(str(raw or "").split())
    if not text:
        return (None, None)
    pattern = re.compile(r"(\d{1,2}h(?:\d{2})?)\s*(?:as|às)\s*(\d{1,2}h(?:\d{2})?)", re.I)
    match = pattern.search(text.replace("das ", "").replace("de ", ""))
    if not match:
        return (None, None)
    start_raw = match.group(1).replace("h", ":")
    end_raw = match.group(2).replace("h", ":")
    if ":" not in start_raw:
        start_raw = f"{start_raw}:00"
    if ":" not in end_raw:
        end_raw = f"{end_raw}:00"
    return (parse_time(start_raw), parse_time(end_raw))


def formatar_horario_intervalo(
    horario_inicio: time | None,
    horario_fim: time | None,
) -> str:
    if not horario_inicio or not horario_fim:
        return ""
    if horario_inicio.minute:
        inicio = horario_inicio.strftime("%Hh%M")
    else:
        inicio = horario_inicio.strftime("%Hh")
    if horario_fim.minute:
        fim = horario_fim.strftime("%Hh%M")
    else:
        fim = horario_fim.strftime("%Hh")
    return f"das {inicio} as {fim}"


def normalize_efetivo_payload(raw_values: list[dict] | None) -> list[dict[str, object]]:
    rows: list[dict[str, object]] = []
    for item in raw_values or []:
        if not isinstance(item, dict):
            continue
        cargo = " ".join(str(item.get("cargo", "")).split())
        qtd_raw = str(item.get("quantidade", "")).strip()
        if not cargo:
            continue
        if not qtd_raw.isdigit():
            continue
        quantidade = int(qtd_raw)
        if quantidade <= 0:
            continue
        rows.append({"cargo": cargo, "quantidade": quantidade})
    return rows


def efetivo_total_servidores(raw_values: list[dict] | None) -> int:
    return sum(int(item.get("quantidade", 0) or 0) for item in normalize_efetivo_payload(raw_values))


def formatar_efetivo_resumo(raw_values: list[dict] | None) -> str:
    rows = normalize_efetivo_payload(raw_values)
    if not rows:
        return ""
    partes: list[str] = []
    for row in rows:
        cargo = str(row.get("cargo") or "").strip()
        quantidade = int(row.get("quantidade") or 0)
        if cargo and quantidade > 0:
            partes.append(f"{cargo}: {quantidade}")
    return "; ".join(partes)


def normalize_horario_atendimento(raw_horario: str) -> str:
    text = " ".join(str(raw_horario or "").split()).rstrip(".")
    if not text:
        return ""
    lowered = text.casefold()
    if lowered.startswith("das "):
        return text
    if lowered.startswith("de "):
        return text
    if "às" in lowered or " as " in lowered:
        return f"das {text}"
    return text


def normalize_atividades_selecionadas(atividades: list[str]) -> list[str]:
    selected = {str(item).strip() for item in atividades if str(item).strip()}
    return [item for item in ATIVIDADES_ORDEM_FIXA if item in selected]


def metas_from_atividades(atividades: list[str]) -> list[str]:
    ordered_atividades = normalize_atividades_selecionadas(atividades)
    metas: list[str] = []
    for atividade in ordered_atividades:
        meta = META_POR_ATIVIDADE.get(atividade, "").strip()
        if meta and meta not in metas:
            metas.append(meta)
    return metas


def format_atividades_formatada(atividades: list[str]) -> str:
    ordered = normalize_atividades_selecionadas(atividades)
    return format_lista_bullets(ordered, bullet="\u2022")


def format_metas_formatadas(metas: list[str]) -> str:
    return format_lista_bullets(metas, bullet="\u2022")


def _locais_para_pagina_01(plano: PlanoTrabalho, oficio: Oficio) -> list[str]:
    locais: list[str] = []
    seen: set[str] = set()
    for item in plano.locais_atuacao.all().order_by("ordem", "id"):
        local = " ".join((item.local or "").split())
        if not local or local in seen:
            continue
        seen.add(local)
        locais.append(local)
    if locais:
        return locais
    fallback = " ".join((plano.local or plano.destino or "").split())
    if fallback:
        return [fallback]
    if oficio.cidade_destino and oficio.estado_destino:
        return [f"{oficio.cidade_destino.nome}/{oficio.estado_destino.sigla}"]
    if oficio.cidade_destino:
        return [oficio.cidade_destino.nome]
    return []


def format_locais_atuacao(plano: PlanoTrabalho) -> str:
    locais = list(plano.locais_atuacao.all().order_by("ordem", "id"))
    if not locais:
        fallback = " ".join((plano.local or plano.destino or "").split())
        return fallback or "-"
    if len(locais) == 1 and not locais[0].data:
        return locais[0].local

    rows: list[str] = []
    for item in locais:
        data_label = format_data_curta_br(item.data)
        if data_label:
            rows.append(f"{data_label} - {item.local}")
        else:
            rows.append(item.local)
    return "\n".join(rows)


def build_contextualizacao(plano: PlanoTrabalho) -> str:
    programa = " ".join((plano.programa_projeto or "").split()) or "acao institucional da PCPR"
    destino = " ".join((plano.destino or "").split()) or "municipio definido em planejamento"
    solicitante = formatar_solicitante_exibicao(
        plano.solicitantes_json if isinstance(plano.solicitantes_json, list) else [],
        nome_pcpr=plano.solicitante or "",
    ) or " ".join((plano.solicitante or "").split()) or "demanda institucional registrada"
    periodo = format_periodo_extenso(plano.data_inicio, plano.data_fim) or "periodo informado no cronograma"
    contexto = " ".join((plano.contexto_solicitacao or "").split())

    base = (
        "A Assessoria de Comunicacao Social da Policia Civil do Parana (PCPR), no ambito do "
        f"programa '{programa}', promovera acao itinerante no municipio de {destino}, no periodo de {periodo}. "
        f"A iniciativa visa atender a solicitacao formulada por {solicitante}, levando servicos essenciais de pol\u00edcia judici\u00e1ria a populacao."
    )
    if contexto:
        return f"{base} {contexto}"
    return base


def build_atuacao_formatada(plano: PlanoTrabalho) -> str:
    linhas: list[str] = []
    periodo = format_periodo_extenso(plano.data_inicio, plano.data_fim)
    if periodo:
        linhas.append(f"Datas: {periodo}.")

    locais = list(plano.locais_atuacao.all().order_by("ordem", "id"))
    destinos_json_labels = destinos_labels(
        plano.destinos_json if isinstance(plano.destinos_json, list) else []
    )
    if not locais:
        if len(destinos_json_labels) > 1:
            linhas.append(f"Locais: {format_lista_portugues(destinos_json_labels)}.")
        else:
            local = (
                destinos_json_labels[0]
                if destinos_json_labels
                else " ".join((plano.local or plano.destino or "").split())
            )
            if local:
                linhas.append(f"Local: {local}.")
    elif len(locais) == 1 and not locais[0].data:
        linhas.append(f"Local: {locais[0].local}.")
    else:
        linhas.append("Locais:")
        for item in locais:
            data_label = format_data_curta_br(item.data)
            if data_label:
                linhas.append(f"{data_label} - {item.local}")
            else:
                linhas.append(item.local)

    horario = (
        formatar_horario_intervalo(plano.horario_inicio, plano.horario_fim)
        or " ".join((plano.horario_atendimento or "").split())
    )
    if horario:
        linhas.append(f"Horario de atendimento: {horario}.")

    efetivo = " ".join((plano.efetivo_formatado or "").split())
    if efetivo:
        linhas.append(f"Efetivo: {efetivo}")

    estrutura = " ".join((plano.estrutura_apoio or "").split())
    if not estrutura and plano.unidade_movel:
        estrutura = DEFAULT_UNIDADE_MOVEL_TEXTO
    if estrutura:
        linhas.append(f"Estrutura: {estrutura}")

    linhas.append(
        "O atendimento ao publico sera realizado de forma organizada, com acolhimento, orientacao e prestacao de servicos de pol\u00edcia judici\u00e1ria."
    )
    return "\n".join(linhas).strip()


def build_coordenacao_formatada(plano: PlanoTrabalho) -> str:
    cargo_admin = (
        (plano.coordenador_plano.cargo if plano.coordenador_plano else "")
        or plano.coordenador_cargo
        or DEFAULT_COORDENADOR_PLANO_CARGO
    )
    nome_admin = (
        (plano.coordenador_plano.nome if plano.coordenador_plano else "")
        or plano.coordenador_nome
        or DEFAULT_COORDENADOR_PLANO_NOME
    )
    paragrafo_admin = (
        f"Fica designada como Coordenadora Administrativa do Plano a {cargo_admin} {nome_admin}, a qual "
        "ficara responsavel pelo acompanhamento da execucao administrativa do presente Plano de Trabalho, "
        "organizacao das escalas de servidores, controle de materiais e equipamentos, consolidacao de dados "
        "estatisticos, elaboracao de relatorio final e demais providencias necessarias ao regular cumprimento da acao."
    )
    if not plano.possui_coordenador_municipal or not plano.coordenador_municipal:
        return paragrafo_admin

    coord_municipal = plano.coordenador_municipal
    paragrafo_municipal = (
        "Fica designado(a) como Coordenador(a) Municipal do Evento o(a) "
        f"{coord_municipal.cargo} {coord_municipal.nome}, do municipio de {coord_municipal.cidade}, "
        "que ficara responsavel pela articulacao local da acao, apoio institucional a equipe da Policia Civil do Parana, "
        "organizacao do espaco de atendimento, suporte logistico no ambito municipal e demais providencias necessarias a boa execucao do evento."
    )
    return f"{paragrafo_admin}\n\n{paragrafo_municipal}"


def build_valor_total_bloco(plano: PlanoTrabalho) -> str:
    valor_total_decimal = _to_decimal(plano.valor_total_calculado) or _to_decimal(plano.valor_total)
    valor_unitario_decimal = _to_decimal(plano.valor_unitario) or valor_total_decimal
    valor_total = format_monetario_br(valor_total_decimal)
    valor_total_extenso = format_valor_extenso(valor_total_decimal)
    valor_unitario = format_monetario_br(valor_unitario_decimal)
    valor_unitario_extenso = format_valor_extenso(valor_unitario_decimal)
    composicao = " ".join((plano.composicao_diarias or "").split()) or "1 x 100%"
    qtd_servidores = efetivo_total_servidores(
        plano.efetivo_json if isinstance(plano.efetivo_json, list) else []
    ) or int(plano.quantidade_servidores or 0)

    linhas: list[str] = []
    if valor_total:
        if valor_total_extenso:
            linhas.append(f"Valor total: {valor_total} ({valor_total_extenso}).")
        else:
            linhas.append(f"Valor total: {valor_total}.")
    if composicao and valor_unitario:
        if valor_unitario_extenso:
            linhas.append(
                f"Valor correspondente a {composicao}, por servidor, no valor unitario de {valor_unitario} ({valor_unitario_extenso})."
            )
        else:
            linhas.append(
                f"Valor correspondente a {composicao}, por servidor, no valor unitario de {valor_unitario}."
            )
    if qtd_servidores > 0:
        linhas.append(
            f"Quantidade de servidores considerada para calculo: {qtd_servidores}."
        )
    linhas.append(
        "O custeio contempla deslocamento, logistica e suporte operacional para cumprimento integral da acao."
    )
    return " ".join(linhas).strip()


def _resolve_sede(cfg: OficioConfig, oficio: Oficio) -> str:
    if oficio.cidade_sede and oficio.estado_sede:
        return f"{oficio.cidade_sede.nome}/{oficio.estado_sede.sigla}"
    if oficio.cidade_sede:
        return oficio.cidade_sede.nome
    if cfg.sede_cidade_default and cfg.sede_cidade_default.estado:
        return f"{cfg.sede_cidade_default.nome}/{cfg.sede_cidade_default.estado.sigla}"
    if cfg.cidade and cfg.uf:
        return f"{cfg.cidade}/{cfg.uf}"
    return "Curitiba/PR"


def _resolve_endereco(cfg: OficioConfig) -> str:
    parts: list[str] = []
    logradouro = " ".join((cfg.logradouro or "").split())
    numero = " ".join((cfg.numero or "").split())
    complemento = " ".join((cfg.complemento or "").split())
    bairro = " ".join((cfg.bairro or "").split())
    cidade = " ".join((cfg.cidade or "").split())
    uf = " ".join((cfg.uf or "").split())
    cep = " ".join((cfg.cep or "").split())

    if logradouro:
        base = logradouro
        if numero:
            base = f"{base}, {numero}"
        if complemento:
            base = f"{base} - {complemento}"
        parts.append(base)
    if bairro:
        parts.append(bairro)
    if cidade and uf:
        parts.append(f"{cidade}/{uf}")
    elif cidade:
        parts.append(cidade)
    if cep:
        parts.append(f"CEP {cep}")
    return " - ".join(parts)


def _chefia_nome_cargo(cfg: OficioConfig) -> tuple[str, str]:
    if cfg.assinante:
        nome = " ".join((cfg.assinante.nome or "").split())
        cargo = " ".join((cfg.assinante.cargo or "").split())
        if nome or cargo:
            return nome, cargo
    nome_cfg = " ".join((cfg.unidade_nome or "").split())
    return nome_cfg or "Chefia responsavel", "Cargo da chefia"


def build_plano_placeholders(
    plano: PlanoTrabalho,
    oficio: Oficio,
    cfg: OficioConfig,
) -> dict[str, str]:
    atividades = _ordered_descriptions(
        PlanoTrabalhoAtividade.objects.filter(plano=plano).order_by("ordem", "id")
    )
    metas = _ordered_descriptions(
        PlanoTrabalhoMeta.objects.filter(plano=plano).order_by("ordem", "id")
    )
    if atividades and not metas:
        metas = metas_from_atividades(atividades)
    recursos = _ordered_descriptions(
        PlanoTrabalhoRecurso.objects.filter(plano=plano).order_by("ordem", "id")
    )
    locais = PlanoTrabalhoLocalAtuacao.objects.filter(plano=plano).order_by("ordem", "id")
    sede = _resolve_sede(cfg, oficio)
    nome_chefia, cargo_chefia = _chefia_nome_cargo(cfg)
    divisao = title_case_pt(" ".join((cfg.origem_nome or "").split())) or "Policia Civil do Parana"
    unidade = (
        title_case_pt(" ".join((cfg.unidade_nome or "").split()))
        or "Assessoria de Comunicacao Social"
    )
    composicao_diarias = " ".join((plano.composicao_diarias or "").split()) or "1 x 100%"
    valor_unitario_decimal = _to_decimal(plano.valor_unitario) or _to_decimal(plano.valor_total)
    valor_total_decimal = _to_decimal(plano.valor_total_calculado) or _to_decimal(plano.valor_total)

    destinos_json = plano.destinos_json if isinstance(plano.destinos_json, list) else []
    destinos_json_labels = destinos_labels(destinos_json)
    destino = format_lista_portugues(destinos_json_labels)
    if not destino:
        destino = " ".join((plano.destino or "").split())
    if not destino:
        if oficio.cidade_destino and oficio.estado_destino:
            destino = f"{oficio.cidade_destino.nome}/{oficio.estado_destino.sigla}"
        elif oficio.cidade_destino:
            destino = oficio.cidade_destino.nome

    solicitantes_json = (
        plano.solicitantes_json if isinstance(plano.solicitantes_json, list) else []
    )
    solicitante = formatar_solicitante_exibicao(
        solicitantes_json,
        nome_pcpr=plano.solicitante or "",
    )
    if not solicitante:
        solicitante = " ".join((plano.solicitante or "").split()) or "demanda institucional"

    horario = (
        formatar_horario_intervalo(plano.horario_inicio, plano.horario_fim)
        or normalize_horario_atendimento(plano.horario_atendimento or "")
        or "das 09h as 17h"
    )

    efetivo_rows = plano.efetivo_json if isinstance(plano.efetivo_json, list) else []
    efetivo = formatar_efetivo_resumo(efetivo_rows) or " ".join((plano.efetivo_formatado or "").split())
    quantidade_int = efetivo_total_servidores(efetivo_rows)
    if quantidade_int <= 0:
        quantidade_int = int(plano.quantidade_servidores or plano.efetivo_por_dia or 0)
    if quantidade_int <= 0:
        quantidade_int = int(oficio.viajantes.count())
    if not efetivo and quantidade_int > 0:
        efetivo = f"{quantidade_int} servidores."
    quantidade_servidores = str(quantidade_int)

    locais_pagina_01 = destinos_json_labels or _locais_para_pagina_01(plano, oficio)
    locais_formatado = format_lista_portugues(locais_pagina_01) or destino

    estrutura_formatada = " ".join((plano.estrutura_apoio or "").split())
    if not estrutura_formatada and plano.unidade_movel:
        estrutura_formatada = DEFAULT_UNIDADE_MOVEL_TEXTO
    dias_evento_extenso = format_periodo_evento_extenso(plano.data_inicio, plano.data_fim)
    atividades_formatada = format_atividades_formatada(atividades)
    metas_formatadas = format_metas_formatadas(metas)

    placeholders: dict[str, str] = {
        "divisao": divisao,
        "unidade": unidade,
        "numero_plano": f"{int(plano.numero or 0):02d}",
        "ano_plano": str(int(plano.ano or timezone.localdate().year)),
        "sigla_unidade": " ".join((plano.sigla_unidade or "").split()).upper(),
        "dias_evento_extenso": dias_evento_extenso,
        "locais_formatado": locais_formatado,
        "destino": destino,
        "solicitante": solicitante,
        "contexto_solicitacao": " ".join((plano.contexto_solicitacao or "").split()),
        "metas_formatadas": metas_formatadas,
        "atividades_formatada": atividades_formatada,
        "atividades_formatadas": atividades_formatada,
        "datas_evento_extenso": dias_evento_extenso,
        "locais_formatados": format_locais_atuacao(plano),
        "horario_atendimento": horario,
        "quantidade_de_servidores": quantidade_servidores,
        "efetivo_formatado": efetivo,
        "estrutura_formatada": estrutura_formatada,
        "valor_total": format_monetario_br(valor_total_decimal),
        "valor_total_extenso": format_valor_extenso(valor_total_decimal),
        "composicao_diarias": composicao_diarias,
        "valor_unitario": format_monetario_br(valor_unitario_decimal),
        "valor_unitario_extenso": format_valor_extenso(valor_unitario_decimal),
        "recursos_formatados": format_lista_bullets(recursos),
        "coordenacao_formatada": build_coordenacao_formatada(plano),
        "sede": sede,
        "data_extenso": format_data_extenso_br(timezone.localdate()),
        "nome_chefia": nome_chefia,
        "cargo_chefia": cargo_chefia,
        "unidade_rodape": title_case_pt(" ".join((cfg.unidade_nome or "").split())),
        "endereco": _resolve_endereco(cfg),
        "telefone": " ".join((cfg.telefone or "").split()),
        "email": " ".join((cfg.email or "").split()),
        "contextualizacao_formatada": build_contextualizacao(plano),
        "atuacao_formatada": build_atuacao_formatada(plano),
        "valor_total_bloco": build_valor_total_bloco(plano),
    }

    # Aliases para compatibilidade com placeholders reais do template DOCX.
    placeholders.update(
        {
            "numero_plano_trabalho": f"{int(plano.numero or 0):02d}/{int(plano.ano or timezone.localdate().year)}",
            "metas_formatada": placeholders["metas_formatadas"],
            "recursos_formatado": placeholders["recursos_formatados"],
            "valor_total_por_extenso": placeholders["valor_total_extenso"],
            "valor_unitario_por_extenso": placeholders["valor_unitario_extenso"],
            "diarias_x": placeholders["composicao_diarias"],
            "unidade_movel": placeholders["estrutura_formatada"],
            "coordenação formatada": placeholders["coordenacao_formatada"],
        }
    )

    if not placeholders["locais_formatados"] and locais.exists():
        placeholders["locais_formatados"] = "\n".join(
            f"{format_data_curta_br(item.data)} - {item.local}" if item.data else item.local
            for item in locais
        )

    return placeholders


def validate_required_placeholders(placeholders: dict[str, str]) -> list[str]:
    missing: list[str] = []
    for key in REQUIRED_PLACEHOLDERS:
        value = " ".join(str(placeholders.get(key, "")).split())
        if not value:
            missing.append(key)
    return missing
