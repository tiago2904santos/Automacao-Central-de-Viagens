# viagens/documents/document.py
from __future__ import annotations

import re
import tempfile
from io import BytesIO
from pathlib import Path
from typing import Iterable, List, Tuple

try:
    import pythoncom
    import win32com.client as win32client  # type: ignore
except ImportError:  # pragma: no cover - optional Windows dependency
    pythoncom = None  # type: ignore[assignment]
    win32client = None  # type: ignore[assignment]

from django.conf import settings
from docx import Document as DocxFactory
from docx.document import Document as DocxDocument
from docx.oxml.ns import qn
from docx.shared import Pt

from viagens.models import Cidade, Estado, Oficio, Trecho, Viajante

PLACEHOLDER_RE = re.compile(r"{{\s*([^}]+?)\s*}}")

# =========================
# AJUSTES DE LAYOUT
# =========================
FONT_NAME = "Times New Roman"
FONT_SIZE_PT = 8

# “largura” aproximada para estimar quebra em tabela
NAME_MAX_VISUAL = 35.0
CARGO_MAX_VISUAL = 35.0

NBSP = "\u00A0"  # não quebra e “segura” linha vazia no Word


# =========================
# FORMATADORES
# =========================
def _fmt_date(value) -> str:
    if not value:
        return ""
    try:
        return value.strftime("%d/%m/%Y")
    except Exception:
        return str(value)


def _fmt_time(value) -> str:
    if not value:
        return ""
    try:
        return value.strftime("%H:%M")
    except Exception:
        return str(value)


def _fmt_local(cidade: Cidade | None, estado: Estado | None) -> str:
    if cidade and estado:
        return f"{cidade.nome}/{estado.sigla}"
    if cidade:
        return cidade.nome
    if estado:
        return estado.sigla
    return ""


def _join(parts: Iterable[str], sep=" - ") -> str:
    return sep.join([p for p in parts if p])


def _title_case(text: str) -> str:
    return " ".join(w.capitalize() for w in (text or "").split())


# =========================
# FONTE (garante Times 8)
# =========================
def _apply_font(run, font_name: str = FONT_NAME, font_size_pt: int = FONT_SIZE_PT):
    run.font.name = font_name
    run.font.size = Pt(font_size_pt)

    # garante Times em todos os “slots” do Word
    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.get_or_add_rFonts()
    rfonts.set(qn("w:ascii"), font_name)
    rfonts.set(qn("w:hAnsi"), font_name)
    rfonts.set(qn("w:cs"), font_name)
    rfonts.set(qn("w:eastAsia"), font_name)


# =========================
# ITERADORES (parágrafos em tabelas + header/footer)
# =========================
def _iter_all_paragraphs(doc: DocxDocument):
    # body
    for p in doc.paragraphs:
        yield p
    # tables (body)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    yield p

    # headers/footers também (pra não “sumir” nada no futuro)
    for section in doc.sections:
        for p in section.header.paragraphs:
            yield p
        for table in section.header.tables:
            for row in table.rows:
                for cell in row.cells:
                    for p in cell.paragraphs:
                        yield p

        for p in section.footer.paragraphs:
            yield p
        for table in section.footer.tables:
            for row in table.rows:
                for cell in row.cells:
                    for p in cell.paragraphs:
                        yield p


# =========================
# REPLACE (preservando formatação)
# - substitui placeholder por "" se não encontrar no mapping
# =========================
def _sub_placeholders(text: str, mapping: dict[str, str]) -> str:
    if "{{" not in text:
        return text

    def repl(m):
        key = m.group(1).strip()
        # se não existir -> vazio (como você pediu)
        return str(mapping.get(key, ""))

    return PLACEHOLDER_RE.sub(repl, text)


def replace_everywhere(doc: DocxDocument, mapping: dict[str, str]):
    """
    Substitui placeholders preservando a formatação do Word.
    (Não limpa o parágrafo; altera run por run.)
    """
    for p in _iter_all_paragraphs(doc):
        # rápido: se não tem placeholder no texto todo, pula
        full_text = "".join(r.text for r in p.runs)
        if "{{" not in full_text:
            continue

        # 1) tenta trocar run-a-run (mantém bold/itálico do run)
        changed_any = False
        for r in p.runs:
            new_text = _sub_placeholders(r.text, mapping)
            if new_text != r.text:
                r.text = new_text
                _apply_font(r)
                changed_any = True

        if changed_any:
            continue

        # 2) fallback: se o Word quebrou o placeholder em runs diferentes,
        # reconstrói o parágrafo inteiro (perde granularidade, mas substitui)
        new_full = _sub_placeholders(full_text, mapping)
        if new_full != full_text:
            p.clear()
            run = p.add_run(new_full)
            _apply_font(run)


# =========================
# MULTILINHA “RICA” (negrito/normal)
# =========================
RunPart = Tuple[str, bool]   # (texto, bold)
RichLine = List[RunPart]     # uma “linha”


def _visual_length(text: str) -> float:
    total = 0.0
    for ch in text:
        total += 1.5 if ch.isupper() else 1.0
    return total


def _estimate_wrapped_lines(text: str, max_visual: float) -> int:
    words = (text or "").split()
    if not words:
        return 1

    lines = 1
    current = 0.0
    for w in words:
        wlen = _visual_length(w)
        add = wlen if current == 0 else wlen + 1.0  # espaço
        if current + add <= max_visual:
            current += add
        else:
            lines += 1
            current = wlen
    return lines


def _blank_lines(n: int) -> list[RichLine]:
    return [[(NBSP, False)] for _ in range(max(0, n))]


def _write_rich_lines(p, lines: list[RichLine]):
    p.clear()
    first_line = True

    for line in lines:
        if not first_line:
            p.add_run().add_break()
        first_line = False

        for text, bold in line:
            run = p.add_run(text)
            run.bold = bold
            _apply_font(run)

    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)


def replace_placeholder_rich(doc: DocxDocument, key: str, lines: list[RichLine]):
    token = f"{{{{{key}}}}}"
    for p in _iter_all_paragraphs(doc):
        full = "".join(r.text for r in p.runs).strip()
        if full == token:
            _write_rich_lines(p, lines)


# =========================
# BLOCOS: SERVIDORES
# =========================
def build_col_nomes(viajantes: list[Viajante]) -> list[RichLine]:
    lines: list[RichLine] = []
    for i, v in enumerate(viajantes):
        nome = _title_case(v.nome or "")
        lines.append([(nome or NBSP, False)])

        if i < len(viajantes) - 1:
            used = _estimate_wrapped_lines(nome, NAME_MAX_VISUAL)
            blanks = 2 if used <= 1 else 1
            lines.extend(_blank_lines(blanks))

    return lines or [[(NBSP, False)]]


def build_col_rgcpf(viajantes: list[Viajante]) -> list[RichLine]:
    lines: list[RichLine] = []
    for i, v in enumerate(viajantes):
        lines.append([("RG: ", True), ((v.rg or "").strip(), False)])
        lines.append([("CPF: ", True), ((v.cpf or "").strip(), False)])

        if i < len(viajantes) - 1:
            lines.extend(_blank_lines(1))

    return lines or [[(NBSP, False)]]


def build_col_cargo(viajantes: list[Viajante]) -> list[RichLine]:
    lines: list[RichLine] = []
    for i, v in enumerate(viajantes):
        cargo = (v.cargo or "").strip()
        lines.append([(cargo or NBSP, True)])  # cargo sempre em negrito

        if i < len(viajantes) - 1:
            used = _estimate_wrapped_lines(cargo, CARGO_MAX_VISUAL)
            blanks = 2 if used <= 1 else 1
            lines.extend(_blank_lines(blanks))

    return lines or [[(NBSP, False)]]


# =========================
# BLOCOS: DESTINOS + ROTEIRO
# =========================
def build_destinos_bloco(trechos: list[Trecho]) -> str:
    destinos: list[str] = []
    seen = set()
    for t in trechos:
        d = _fmt_local(t.destino_cidade, t.destino_estado)
        if d and d not in seen:
            seen.add(d)
            destinos.append(d)
    return ", ".join(destinos)


def build_roteiro_ida(trechos: list[Trecho]) -> tuple[list[RichLine], list[RichLine]]:
    saida_lines: list[RichLine] = []
    chegada_lines: list[RichLine] = []

    for i, t in enumerate(trechos):
        origem = _fmt_local(t.origem_cidade, t.origem_estado)
        destino = _fmt_local(t.destino_cidade, t.destino_estado)

        saida_dt = _join([_fmt_date(t.saida_data), _fmt_time(t.saida_hora)], " ")
        chegada_dt = _join([_fmt_date(t.chegada_data), _fmt_time(t.chegada_hora)], " ")

        # Linha inteira em negrito (como você pediu)
        saida_lines.append([(f"Saída {origem}: {saida_dt}".strip() or NBSP, True)])
        chegada_lines.append([(f"Chegada {destino}: {chegada_dt}".strip() or NBSP, True)])

        if i < len(trechos) - 1:
            saida_lines.extend(_blank_lines(1))
            chegada_lines.extend(_blank_lines(1))

    return (saida_lines or [[(NBSP, False)]], chegada_lines or [[(NBSP, False)]])


# =========================
# ROTEIRO DE RETORNO (mantém negrito mesmo com placeholders)
# =========================
def _patch_roteiro_retorno(doc: DocxDocument, mapping: dict[str, str]):
    """
    No seu modelo o retorno está como:
      "Saída {{destino}}: {{data_hora_saida_destino}}"
      "Chegada {{sede}}: {{data_hora_chegada_sede}}"
    O replace run-a-run pode falhar se o Word quebrar os placeholders em runs.
    Aqui a gente identifica esses parágrafos e reescreve como rich text.
    """
    for p in _iter_all_paragraphs(doc):
        full = "".join(r.text for r in p.runs).strip()
        if "{{data_hora_saida_destino}}" in full:
            # substitui TUDO e deixa a linha inteira em negrito
            line = _sub_placeholders(full, mapping).strip()
            _write_rich_lines(p, [[(line or NBSP, True)]])
        elif "{{data_hora_chegada_sede}}" in full:
            line = _sub_placeholders(full, mapping).strip()
            _write_rich_lines(p, [[(line or NBSP, True)]])


# =========================
# DOCX PRINCIPAL
# =========================
def build_oficio_docx_bytes(oficio: Oficio) -> BytesIO:
    template_path = str(Path(settings.BASE_DIR) / "viagens" / "documents" / "oficio_model.docx")
    doc = DocxFactory(template_path)

    viajantes = list(oficio.viajantes.all().order_by("nome"))
    trechos = list(oficio.trechos.order_by("ordem"))  # type: ignore

    # colunas “ricas”
    replace_placeholder_rich(doc, "col_servidor", build_col_nomes(viajantes))
    replace_placeholder_rich(doc, "col_rgcpf", build_col_rgcpf(viajantes))
    replace_placeholder_rich(doc, "col_cargo", build_col_cargo(viajantes))

    saida_lines, chegada_lines = build_roteiro_ida(trechos)
    replace_placeholder_rich(doc, "col_ida_saida", saida_lines)
    replace_placeholder_rich(doc, "col_ida_chegada", chegada_lines)

    # sede (p/ retorno) - pega da origem do 1º trecho se existir
    sede = ""
    if trechos:
        t0 = trechos[0]
        sede = _fmt_local(t0.origem_cidade, t0.origem_estado)

    # destino principal (último destino do roteiro) – você já tem oficio.destino,
    # mas aqui garantimos um fallback coerente.
    destino_principal = oficio.destino or ""
    if not destino_principal and trechos:
        destino_principal = _fmt_local(trechos[-1].destino_cidade, trechos[-1].destino_estado)

    # campos simples (repare que agora inclui retorno + caracterizada + armamento + sede)
    mapping = {
        "oficio": oficio.oficio or "",
        "ano": str(oficio.created_at.year) if oficio.created_at else "",
        "data_do_oficio": _fmt_date(oficio.created_at.date()) if oficio.created_at else "",
        "protocolo": oficio.protocolo or "",
        "destino": destino_principal,
        "destinos_bloco": build_destinos_bloco(trechos),

        "diarias_x": (oficio.quantidade_diarias or "").strip(),
        "diaria": (oficio.valor_diarias or "").strip(),
        "valor_extenso": (oficio.valor_diarias_extenso or "").strip(),

        "viatura": (oficio.modelo or "").strip(),
        "combustivel": (oficio.combustivel or "").strip(),
        "placa": (oficio.placa or "").strip(),
        "motorista": _title_case((oficio.motorista or "").strip()),

        "caracterizada": "Sim" if oficio.motorista_carona else "Não",
        "armamento": "",  # se você tiver no modelo, pluga aqui

        "sede": sede,

        # retorno
        "data_hora_saida_destino": _join(
            [_fmt_date(oficio.retorno_saida_data), _fmt_time(oficio.retorno_saida_hora)],
            " ",
        ),
        "data_hora_chegada_sede": _join(
            [_fmt_date(oficio.retorno_chegada_data), _fmt_time(oficio.retorno_chegada_hora)],
            " ",
        ),

        "motivo": (oficio.motivo or "").strip(),
    }

    # 1) substitui preservando formatação do modelo
    replace_everywhere(doc, mapping)

    # 2) garante retorno em negrito e substitui mesmo se placeholder estiver “quebrado”
    _patch_roteiro_retorno(doc, mapping)

    buf = BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf


# =========================
# PDF (Word/Windows)
# =========================


def _ensure_pywin32_available() -> None:
    if pythoncom is None or win32client is None:
        raise RuntimeError(
            "A conversão de DOCX para PDF exige pywin32 (pythoncom + win32com). "
            "Instale a dependência no ambiente Windows para usar essa funcionalidade."
        )


def docx_bytes_to_pdf_bytes(docx_bytes: bytes) -> bytes:
    """
    Converte DOCX em PDF usando Microsoft Word (fidelidade alta).
    Requer Windows + Word instalado.
    """
    _ensure_pywin32_available()
    pythoncom.CoInitialize()

    with tempfile.TemporaryDirectory() as tmpdir:
        docx_path = Path(tmpdir) / "oficio.docx"
        pdf_path = Path(tmpdir) / "oficio.pdf"
        docx_path.write_bytes(docx_bytes)

        word = win32client.Dispatch("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0

        try:
            doc = word.Documents.Open(str(docx_path))
            doc.ExportAsFixedFormat(
                OutputFileName=str(pdf_path),
                ExportFormat=17,       # PDF
                OpenAfterExport=False,
                OptimizeFor=0,
                Item=0,
            )
            doc.Close(False)
        finally:
            word.Quit()

        return pdf_path.read_bytes()


def build_oficio_docx_and_pdf_bytes(oficio: Oficio) -> tuple[bytes, bytes]:
    """
    Retorna (docx_bytes, pdf_bytes)
    """
    docx_buf = build_oficio_docx_bytes(oficio)
    docx_bytes = docx_buf.getvalue()
    pdf_bytes = docx_bytes_to_pdf_bytes(docx_bytes)
    return docx_bytes, pdf_bytes


# Alias pra não quebrar import antigo em views.py
def build_oficio_docx_and_pdf(oficio: Oficio):
    return build_oficio_docx_and_pdf_bytes(oficio)
