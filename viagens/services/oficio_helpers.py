from __future__ import annotations

import unicodedata
from decimal import Decimal, InvalidOperation

from django.conf import settings
from django.utils import timezone

from viagens.models import Oficio, Trecho, Viajante
from viagens.utils.normalize import format_protocolo_num

CAPITAIS = {
    "ARACAJU",
    "BELEM",
    "BELO HORIZONTE",
    "BOA VISTA",
    "BRASILIA",
    "CAMPO GRANDE",
    "CUIABA",
    "CURITIBA",
    "FLORIANOPOLIS",
    "FORTALEZA",
    "GOIANIA",
    "JOAO PESSOA",
    "MACAPA",
    "MACEIO",
    "MANAUS",
    "NATAL",
    "PALMAS",
    "PORTO ALEGRE",
    "PORTO VELHO",
    "RECIFE",
    "RIO BRANCO",
    "RIO DE JANEIRO",
    "SALVADOR",
    "SAO LUIS",
    "SAO PAULO",
    "TERESINA",
    "VITORIA",
}


def is_viagem_fora_pr(oficio: Oficio, trechos: list[Trecho] | None = None) -> bool:
    trechos_list = trechos or list(
        oficio.trechos.select_related("destino_estado", "destino_cidade__estado")
    )
    for trecho in trechos_list:
        estado = trecho.destino_estado or (
            trecho.destino_cidade.estado if trecho.destino_cidade else None
        )
        if estado and estado.sigla.upper() != "PR":
            return True
    estado_oficio = oficio.estado_destino or (
        oficio.cidade_destino.estado if oficio.cidade_destino else None
    )
    return bool(estado_oficio and estado_oficio.sigla.upper() != "PR")


def get_data_saida_viagem(oficio: Oficio, trechos: list[Trecho] | None = None):
    trechos_list = trechos or list(oficio.trechos.order_by("ordem", "id"))
    for trecho in trechos_list:
        if trecho.saida_data:
            return trecho.saida_data
    return None


def build_assunto(oficio: Oficio, trechos: list[Trecho] | None = None) -> dict[str, str]:
    data_oficio = (
        timezone.localdate(oficio.created_at) if oficio.created_at else timezone.localdate()
    )
    data_saida = get_data_saida_viagem(oficio, trechos)
    convalidacao = bool(data_saida and data_oficio >= data_saida)
    assunto_tipo = (getattr(oficio, "assunto_tipo", "") or "").strip().upper()
    if assunto_tipo == Oficio.AssuntoTipo.CONVALIDACAO:
        assunto_termo = "convalidação"
    else:
        assunto_termo = "autorização"
    if convalidacao:
        return {
            "assunto_termo": assunto_termo,
            "assunto": "Solicitação de convalidação e concessão de diárias.",
            "assunto_oficio": "(convalidação)",
        }
    return {
        "assunto_termo": assunto_termo,
        "assunto": "Solicitação de autorização e concessão de diárias.",
        "assunto_oficio": "",
    }


def build_destinos(trechos: list[Trecho]) -> list[str]:
    destinos: list[str] = []
    for trecho in trechos:
        cidade = trecho.destino_cidade
        if cidade and cidade.nome not in destinos:
            destinos.append(cidade.nome)
    return destinos


def infer_tipo_destino(trechos: list[Trecho]) -> str:
    def _normalize_city(value: str) -> str:
        raw = unicodedata.normalize("NFKD", value or "")
        return "".join(ch for ch in raw if not unicodedata.combining(ch)).upper().strip()

    destinos_upper: list[tuple[str, str]] = []
    for trecho in trechos:
        cidade_nome = (trecho.destino_cidade.nome if trecho.destino_cidade else "") or ""
        cidade_upper = _normalize_city(cidade_nome)
        estado_obj = trecho.destino_estado or (
            trecho.destino_cidade.estado if trecho.destino_cidade else None
        )
        estado_sigla = ((estado_obj.sigla if estado_obj else "") or "").upper().strip()
        destinos_upper.append((cidade_upper, estado_sigla))

    if any(cidade == "BRASILIA" and estado == "DF" for cidade, estado in destinos_upper):
        return "BRASILIA"
    if any(cidade in CAPITAIS for cidade, _ in destinos_upper):
        return "CAPITAL"
    return "INTERIOR"


def format_motorista(oficio: Oficio, viajantes: list[Viajante]) -> str:
    motorista_nome = (oficio.motorista or "").strip()
    if oficio.motorista_viajante:
        motorista_nome = oficio.motorista_viajante.nome or motorista_nome
    motorista_nome = motorista_nome or "-"
    viajantes_ids = {str(item.id) for item in viajantes}
    motorista_id = str(oficio.motorista_viajante_id or "")
    motorista_carona = oficio.motorista_carona or (motorista_id and motorista_id not in viajantes_ids)
    if not motorista_carona:
        return motorista_nome
    oficio_motorista = (oficio.motorista_oficio or "-").strip() or "-"
    protocolo_motorista = (
        format_protocolo_num(oficio.motorista_protocolo)
        or (oficio.motorista_protocolo or "-").strip()
        or "-"
    )
    return f"{motorista_nome} (carona) – Ofício {oficio_motorista} – Protocolo {protocolo_motorista}"


def format_armamento(value) -> str:
    if value is None:
        return "Não"
    if isinstance(value, bool):
        return "Sim" if value else "Não"
    if isinstance(value, (int, float, Decimal)):
        return "Sim" if value else "Não"
    raw = str(value).strip().lower()
    if raw in {"s", "sim", "true", "1", "yes", "y"}:
        return "Sim"
    if raw in {"n", "nao", "não", "false", "0", "no"}:
        return "Não"
    return "Sim" if raw else "Não"


def valor_por_extenso_ptbr(valor) -> str:
    if valor is None:
        return "(preencher manualmente)"
    try:
        if isinstance(valor, str):
            raw = valor.replace(".", "").replace(",", ".")
            valor_dec = Decimal(raw)
        else:
            valor_dec = Decimal(valor)
    except (InvalidOperation, TypeError):
        return "(preencher manualmente)"
    try:
        from num2words import num2words  # type: ignore
    except ImportError:
        return "(preencher manualmente)"
    try:
        return num2words(valor_dec, lang="pt_BR", to="currency")
    except Exception:
        return "(preencher manualmente)"


def get_config_oficio() -> dict[str, str]:
    defaults = {"nome_chefia": "Chefia responsável", "cargo_chefia": "Cargo responsável"}
    cfg = getattr(settings, "OFICIO_CHEFIA", {}) or {}
    nome = cfg.get("nome_chefia") or cfg.get("nome") or getattr(
        settings, "OFICIO_CHEFIA_NOME", ""
    )
    cargo = cfg.get("cargo_chefia") or cfg.get("cargo") or getattr(
        settings, "OFICIO_CHEFIA_CARGO", ""
    )
    return {
        "nome_chefia": nome.strip() or defaults["nome_chefia"],
        "cargo_chefia": cargo.strip() or defaults["cargo_chefia"],
    }
