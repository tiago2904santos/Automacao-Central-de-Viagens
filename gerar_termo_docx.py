from __future__ import annotations

from pathlib import Path
from typing import Any

try:
    from docxtpl import DocxTemplate
except ImportError as exc:  # pragma: no cover
    raise ImportError(
        "Biblioteca 'docxtpl' nao encontrada. Instale com: pip install docxtpl"
    ) from exc


MESES_PTBR = {
    1: "janeiro",
    2: "fevereiro",
    3: "marco",
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

PLACEHOLDERS = (
    "divisao",
    "unidade",
    "data_do_evento",
    "destino",
    "unidade_rodape",
    "endereco",
    "telefone",
    "email",
)


def _as_text(value: Any) -> str:
    return str(value).strip() if value not in (None, "") else ""


def _mes_para_texto(mes: Any) -> str:
    if isinstance(mes, int):
        return MESES_PTBR.get(mes, "")

    mes_texto = _as_text(mes).lower()
    if not mes_texto:
        return ""

    # Aceita string numerica de mes, ex.: "2"
    if mes_texto.isdigit():
        return MESES_PTBR.get(int(mes_texto), "")

    return mes_texto


def formatar_data_do_evento(data_do_evento: list[Any] | tuple[Any, ...] | None) -> str:
    """
    Espera formato: [dia_inicio, dia_fim, mes, ano]
    Ex.: [13, 16, "fevereiro", 2026]
    """
    if not isinstance(data_do_evento, (list, tuple)) or len(data_do_evento) != 4:
        return ""

    dia_inicio, dia_fim, mes, ano = data_do_evento

    dia_inicio_txt = _as_text(dia_inicio)
    dia_fim_txt = _as_text(dia_fim)
    mes_txt = _mes_para_texto(mes)
    ano_txt = _as_text(ano)

    if not (dia_inicio_txt and dia_fim_txt and mes_txt and ano_txt):
        return ""

    if dia_inicio_txt == dia_fim_txt:
        return f"dia {dia_inicio_txt} de {mes_txt} de {ano_txt}"

    return f"dia {dia_inicio_txt} a {dia_fim_txt} de {mes_txt} de {ano_txt}"


def formatar_destino(destinos: list[str] | tuple[str, ...] | None) -> str:
    if not isinstance(destinos, (list, tuple)):
        return ""

    itens = [_as_text(item) for item in destinos]
    itens = [item for item in itens if item]

    if not itens:
        return ""
    if len(itens) == 1:
        return itens[0]
    if len(itens) == 2:
        return f"{itens[0]} e {itens[1]}"

    return f"{', '.join(itens[:-1])} e {itens[-1]}"


def montar_contexto(config: dict[str, Any], termo: dict[str, Any]) -> dict[str, str]:
    config = dict(config or {})
    termo = dict(termo or {})

    contexto: dict[str, str] = {chave: "" for chave in PLACEHOLDERS}

    contexto["divisao"] = _as_text(config.get("divisao"))
    contexto["unidade"] = _as_text(config.get("unidade"))
    contexto["unidade_rodape"] = _as_text(config.get("unidade_rodape"))
    contexto["endereco"] = _as_text(config.get("endereco"))
    contexto["telefone"] = _as_text(config.get("telefone"))
    contexto["email"] = _as_text(config.get("email"))

    contexto["data_do_evento"] = formatar_data_do_evento(termo.get("data_do_evento"))
    contexto["destino"] = formatar_destino(termo.get("destino"))

    return contexto


def gerar_termo_docx(
    *,
    config: dict[str, Any],
    termo: dict[str, Any],
    template_path: str | Path = "modelo.docx",
    output_path: str | Path = "termo_autorizacao_gerado.docx",
) -> Path:
    template_path = Path(template_path)
    output_path = Path(output_path)

    if not template_path.exists():
        raise FileNotFoundError(f"Template nao encontrado: {template_path}")

    context = montar_contexto(config, termo)

    doc = DocxTemplate(str(template_path))
    doc.render(context)
    doc.save(str(output_path))

    return output_path


if __name__ == "__main__":
    CONFIG = {
        "divisao": "ASCOM",
        "unidade": "Prefeitura Municipal",
        "unidade_rodape": "Prefeitura Municipal - ASCOM",
        "endereco": "Rua Exemplo, 123",
        "telefone": "(11) 0000-0000",
        "email": "ascom@prefeitura.gov.br",
    }

    TERMO = {
        "data_do_evento": [13, 16, "fevereiro", 2026],
        "destino": ["Curitiba/PR", "Sao Jose dos Pinhais/PR", "Ponta Grossa/PR"],
    }

    arquivo_saida = gerar_termo_docx(
        config=CONFIG,
        termo=TERMO,
        template_path="modelo.docx",
        output_path="termo_autorizacao_gerado.docx",
    )
    print(f"Arquivo gerado: {arquivo_saida}")
