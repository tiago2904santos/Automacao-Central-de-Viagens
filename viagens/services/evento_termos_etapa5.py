# viagens/services/evento_termos_etapa5.py
"""Termos por ofício e viajante para a Etapa 5 do fluxo guiado. Listagem agrupada por ofício e prefill a partir do ofício."""
from __future__ import annotations

from datetime import date

from viagens.models import Evento, Oficio, TermoAutorizacao, Viajante


def get_termos_etapa5_por_oficio(evento: Evento) -> list[dict]:
    """
    Retorna lista agrupada por ofício para a Etapa 5.
    Cada item: {
        "oficio", "numero_display",
        "viajantes": [ {"viajante", "is_ascom", "precisa_termo", "termo"|None, "status": "existe"|"dispensado"|"pendente"} ]
    }
    """
    oficios = list(
        evento.oficios.prefetch_related("viajantes", "termos_autorizacao").order_by("id")
    )
    result: list[dict] = []
    for oficio in oficios:
        numero_display = (
            getattr(oficio, "numero_formatado", None)
            or getattr(oficio, "oficio", None)
            or f"Ofício {oficio.id}"
        )
        termos_por_viajante = {
            t.viajante_id: t
            for t in oficio.termos_autorizacao.select_related("viajante").all()
            if t.viajante_id
        }
        viajantes_rows: list[dict] = []
        for v in oficio.viajantes.all().order_by("nome"):
            is_ascom = getattr(v, "is_ascom", True)
            precisa_termo = not is_ascom
            termo = termos_por_viajante.get(v.id)
            if termo and termo.dispensado:
                status = "dispensado"
            elif termo:
                status = "existe"
            else:
                status = "pendente"
            viajantes_rows.append({
                "viajante": v,
                "is_ascom": is_ascom,
                "precisa_termo": precisa_termo,
                "termo": termo if (termo and not termo.dispensado) else None,
                "termo_dispensado": termo if (termo and termo.dispensado) else None,
                "status": status,
            })
        result.append({
            "oficio": oficio,
            "numero_display": numero_display,
            "viajantes": viajantes_rows,
        })
    return result


def build_termo_prefill_from_oficio(
    oficio: Oficio, viajante: Viajante | None, evento: Evento | None
) -> dict:
    """
    Monta dados iniciais para TermoAutorizacao a partir do OFÍCIO (trechos + logística).
    - Datas: trechos do ofício (Trecho); fallback evento.data_inicio/fim ou hoje.
    - Destinos: trechos do ofício (origem/destino).
    - Logística: ofício.motorista, ofício.veiculo (modelo, placa), ofício.combustivel.
    """
    hoje = date.today()
    data_inicio = hoje
    data_fim = hoje
    destinos: list[dict[str, str]] = []
    motorista_nome = ""
    veiculo_modelo = ""
    veiculo_placa = ""
    combustivel = ""

    # Datas e destinos dos trechos do ofício (Trecho)
    trechos = list(
        oficio.trechos.select_related(
            "origem_estado", "origem_cidade", "destino_estado", "destino_cidade"
        ).order_by("ordem")
    )
    seen_dest: set[tuple[str, str]] = set()
    for tr in trechos:
        if getattr(tr, "saida_data", None):
            data_inicio = tr.saida_data
            break
    for tr in reversed(trechos):
        if getattr(tr, "chegada_data", None):
            data_fim = tr.chegada_data
            break
    for tr in trechos:
        uf = "PR"
        cidade = ""
        if getattr(tr, "destino_estado", None):
            uf = tr.destino_estado.sigla or uf
        if getattr(tr, "destino_cidade", None):
            cidade = tr.destino_cidade.nome or ""
        if cidade and (uf, cidade) not in seen_dest:
            seen_dest.add((uf, cidade))
            destinos.append({"uf": uf, "cidade": cidade})

    # Fallback datas: evento
    if evento and (data_inicio == hoje or data_fim == hoje):
        if getattr(evento, "data_inicio", None):
            data_inicio = evento.data_inicio
        if getattr(evento, "data_fim", None):
            data_fim = evento.data_fim or data_inicio
    if data_fim is None or data_fim == hoje:
        data_fim = data_inicio

    if not destinos:
        destinos = [{"uf": "PR", "cidade": ""}]

    # Logística do ofício
    motorista_nome = (getattr(oficio, "motorista", None) or "").strip() or ""
    if getattr(oficio, "veiculo", None):
        veiculo = oficio.veiculo
        veiculo_modelo = (getattr(veiculo, "modelo", None) or "").strip() or ""
        veiculo_placa = (getattr(veiculo, "placa", None) or "").strip() or ""
        combustivel = (getattr(veiculo, "combustivel", None) or "").strip() or ""
    combustivel = combustivel or (getattr(oficio, "combustivel", None) or "").strip() or ""

    data_unica = data_inicio == data_fim
    return {
        "data_inicio": data_inicio,
        "data_fim": data_fim,
        "data_unica": data_unica,
        "destinos": destinos,
        "motorista_nome": motorista_nome,
        "veiculo_modelo": veiculo_modelo,
        "veiculo_placa": veiculo_placa,
        "combustivel": combustivel,
    }


def etapa5_necessarios_ok(evento: Evento) -> tuple[bool, bool]:
    """
    Retorna (tem_necessarios, todos_ok).
    tem_necessarios: existe ao menos um viajante não-ASCOM em algum ofício.
    todos_ok: para cada (ofício, viajante) com precisa_termo, existe termo não-dispensado OU termo dispensado.
    """
    grupos = get_termos_etapa5_por_oficio(evento)
    tem_necessarios = False
    todos_ok = True
    for gr in grupos:
        for row in gr["viajantes"]:
            if not row["precisa_termo"]:
                continue
            tem_necessarios = True
            if row["status"] == "pendente":
                todos_ok = False
    if not tem_necessarios:
        return (False, True)  # "não necessário"
    return (True, todos_ok)
