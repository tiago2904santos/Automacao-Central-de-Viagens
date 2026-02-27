from __future__ import annotations

import json
import re
from decimal import Decimal, InvalidOperation
from datetime import datetime, time

from django import forms
from django.utils import timezone
from django.utils.dateparse import parse_date

from .models import (
    Cidade,
    CoordenadorMunicipal,
    Estado,
    Oficio,
    OrdemServico,
    PlanoTrabalho,
    Trecho,
    Viajante,
)
from .services.justificativas import JUSTIFICATIVA_TEMPLATES
from .services.plano_trabalho import (
    ATIVIDADES_ORDEM_FIXA,
    DEFAULT_COORDENADOR_PLANO_CARGO,
    DEFAULT_COORDENADOR_PLANO_NOME,
    SOLICITANTES_ORDEM_FIXA,
    SOLICITANTE_PCPR,
    formatar_horario_intervalo,
    metas_from_atividades,
    normalize_efetivo_payload,
    normalize_atividades_selecionadas,
    normalize_solicitantes,
)
from .utils.normalize import (
    format_oficio_num,
    normalize_digits,
    normalize_oficio_num,
    normalize_protocolo_num,
    normalize_rg,
    normalize_upper_text,
    split_oficio_num,
)


class TrechoForm(forms.ModelForm):
    origem_estado = forms.ModelChoiceField(
        queryset=Estado.objects.order_by("nome"),
        to_field_name="sigla",
        required=False,
    )
    destino_estado = forms.ModelChoiceField(
        queryset=Estado.objects.order_by("nome"),
        to_field_name="sigla",
        required=False,
    )
    origem_cidade = forms.ModelChoiceField(
        queryset=Cidade.objects.none(),
        required=False,
    )
    destino_cidade = forms.ModelChoiceField(
        queryset=Cidade.objects.none(),
        required=False,
    )

    class Meta:
        model = Trecho
        fields = [
            "origem_estado",
            "origem_cidade",
            "destino_estado",
            "destino_cidade",
            "saida_data",
            "saida_hora",
            "chegada_data",
            "chegada_hora",
        ]
        widgets = {
            "saida_data": forms.DateInput(attrs={"type": "date"}, format="%Y-%m-%d"),
            "saida_hora": forms.TimeInput(attrs={"type": "time"}, format="%H:%M"),
            "chegada_data": forms.DateInput(attrs={"type": "date"}, format="%Y-%m-%d"),
            "chegada_hora": forms.TimeInput(attrs={"type": "time"}, format="%H:%M"),
        }

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)

        self._apply_default_uf()
        self._apply_widget_attrs()
        self._set_city_queryset("origem_estado", "origem_cidade")
        self._set_city_queryset("destino_estado", "destino_cidade")
        self._include_existing_city("origem_cidade")
        self._include_existing_city("destino_cidade")
        self._set_state_initial("origem_estado")
        self._set_state_initial("destino_estado")

    def _apply_default_uf(self) -> None:
        if self.is_bound:
            return
        pr = Estado.objects.filter(sigla="PR").first()
        if not pr:
            return
        if not self.initial.get("origem_estado") and not getattr(
            self.instance, "origem_estado_id", None
        ):
            self.initial["origem_estado"] = pr.sigla
        if not self.initial.get("destino_estado") and not getattr(
            self.instance, "destino_estado_id", None
        ):
            self.initial["destino_estado"] = pr.sigla

    def _apply_widget_attrs(self) -> None:
        base_class = "input-field"
        for field_name in self.fields:
            field = self.fields[field_name]
            existing = field.widget.attrs.get("class", "")
            field.widget.attrs["class"] = f"{existing} {base_class}".strip()

        self.fields["origem_estado"].widget.attrs.update({"data-role": "origem-estado"})
        self.fields["origem_cidade"].widget.attrs.update({"data-role": "origem-cidade"})
        self.fields["destino_estado"].widget.attrs.update({"data-role": "destino-estado"})
        self.fields["destino_cidade"].widget.attrs.update({"data-role": "destino-cidade"})
        self.fields["saida_data"].widget.attrs.update({"data-role": "saida-data"})
        self.fields["saida_hora"].widget.attrs.update({"data-role": "saida-hora"})
        self.fields["chegada_data"].widget.attrs.update({"data-role": "chegada-data"})
        self.fields["chegada_hora"].widget.attrs.update({"data-role": "chegada-hora"})
        self.fields["origem_estado"].widget.attrs.update(
            {
                "data-autocomplete-url": "/api/ufs/",
                "data-autocomplete-type": "uf",
            }
        )
        self.fields["destino_estado"].widget.attrs.update(
            {
                "data-autocomplete-url": "/api/ufs/",
                "data-autocomplete-type": "uf",
            }
        )
        self.fields["origem_cidade"].widget.attrs.update(
            {
                "data-autocomplete-url": "/api/cidades-busca/",
                "data-autocomplete-type": "cidade",
            }
        )
        self.fields["destino_cidade"].widget.attrs.update(
            {
                "data-autocomplete-url": "/api/cidades-busca/",
                "data-autocomplete-type": "cidade",
            }
        )

    def _set_city_queryset(self, estado_field: str, cidade_field: str) -> None:
        estado_val = self._get_bound_value(estado_field)
        cidade_val = self._get_bound_value(cidade_field)
        self.fields[cidade_field].widget.attrs["data-selected"] = cidade_val
        if not estado_val:
            self.fields[cidade_field].queryset = Cidade.objects.none()
            return

        self.fields[cidade_field].queryset = Cidade.objects.filter(
            estado__sigla=estado_val
        ).order_by("nome")

    def _get_bound_value(self, field_name: str) -> str:
        key = self.add_prefix(field_name)
        if self.is_bound:
            return self.data.get(key, "")
        if field_name in self.initial:
            return str(self.initial.get(field_name) or "")
        instance_value = getattr(self.instance, field_name, None)
        if instance_value:
            if hasattr(instance_value, "sigla"):
                return str(instance_value.sigla)
            return str(instance_value)
        return ""

    def _set_state_initial(self, field_name: str) -> None:
        field = self.fields[field_name]
        instance_state = getattr(self.instance, field_name, None)
        state_id = getattr(self.instance, f"{field_name}_id", None)
        if state_id and instance_state:
            self.initial[field_name] = instance_state.sigla
            return
        initial_val = self.initial.get(field_name)
        if isinstance(initial_val, Estado):
            self.initial[field_name] = initial_val.sigla
        elif isinstance(initial_val, str) and initial_val.isdigit():
            estado = Estado.objects.filter(id=int(initial_val)).first()
            if estado:
                self.initial[field_name] = estado.sigla

    def _include_existing_city(self, field_name: str) -> None:
        field = self.fields[field_name]
        city = getattr(self.instance, field_name, None)
        if not city:
            city_id = self.initial.get(field_name)
            if city_id:
                try:
                    city = Cidade.objects.get(id=city_id)
                except Cidade.DoesNotExist:
                    city = None
        if not city:
            return

        if not field.queryset.filter(id=city.id).exists():
            city_qs = Cidade.objects.filter(id=city.id)
            field.queryset = field.queryset | city_qs

    def clean(self):
        cleaned_data = super().clean()

        origem_estado = cleaned_data.get("origem_estado")
        origem_cidade = cleaned_data.get("origem_cidade")
        destino_estado = cleaned_data.get("destino_estado")
        destino_cidade = cleaned_data.get("destino_cidade")

        if origem_estado and origem_cidade and origem_cidade.estado_id != origem_estado.id:
            self.add_error("origem_cidade", "A cidade de origem nao pertence a UF selecionada.")
        if destino_estado and destino_cidade and destino_cidade.estado_id != destino_estado.id:
            self.add_error(
                "destino_cidade", "A cidade de destino nao pertence a UF selecionada."
            )

        saida_data = cleaned_data.get("saida_data")
        saida_hora = cleaned_data.get("saida_hora") or time.min
        chegada_data = cleaned_data.get("chegada_data")
        chegada_hora = cleaned_data.get("chegada_hora") or time.min

        if saida_data and chegada_data:
            saida_dt = datetime.combine(saida_data, saida_hora)
            chegada_dt = datetime.combine(chegada_data, chegada_hora)
            if chegada_dt < saida_dt:
                self.add_error(
                    "chegada_data",
                    "A chegada deve ocorrer no mesmo momento ou apos a saida.",
                )

        return cleaned_data


class ServidoresSelectForm(forms.Form):
    servidores = forms.ModelMultipleChoiceField(
        queryset=Viajante.objects.order_by("nome"),
        required=False,
        widget=forms.SelectMultiple(
            attrs={
                "id": "servidoresSelect",
                "class": "input-field",
                "data-autocomplete-url": "/api/servidores/",
                "data-autocomplete-type": "servidor",
            }
        ),
    )


class MotoristaSelectForm(forms.Form):
    motorista = forms.ModelChoiceField(
        queryset=Viajante.objects.order_by("nome"),
        required=False,
        widget=forms.Select(
            attrs={
                "id": "motoristaSelect",
                "class": "input-field",
                "data-autocomplete-url": "/api/motoristas/",
                "data-autocomplete-type": "motorista",
            }
        ),
    )


class ViajanteNormalizeForm(forms.ModelForm):
    cargo_novo = forms.CharField(required=False)

    class Meta:
        model = Viajante
        fields = ["nome", "rg", "cpf", "cargo", "telefone"]

    def clean_nome(self):
        nome = normalize_upper_text(self.cleaned_data.get("nome"))
        partes = [item for item in nome.split(" ") if item]
        if len(partes) < 2:
            raise forms.ValidationError("Informe nome e sobrenome.")
        return nome

    def clean_rg(self):
        rg = normalize_rg(self.cleaned_data.get("rg"))
        if not rg:
            raise forms.ValidationError("Informe o RG.")
        if len(rg) not in {9, 10}:
            raise forms.ValidationError("RG deve conter 9 ou 10 caracteres (digitos + DV).")
        return rg

    def clean_cpf(self):
        cpf = normalize_digits(self.cleaned_data.get("cpf"))
        if len(cpf) != 11:
            raise forms.ValidationError("CPF deve conter 11 digitos.")
        return cpf

    def clean_telefone(self):
        telefone = normalize_digits(self.cleaned_data.get("telefone"))
        if telefone and len(telefone) not in {10, 11}:
            raise forms.ValidationError("Telefone deve conter 10 ou 11 digitos.")
        return telefone

    def clean(self):
        cleaned_data = super().clean()
        cargo_novo = (cleaned_data.get("cargo_novo") or "").strip()
        cargo = (cleaned_data.get("cargo") or "").strip()
        cargo_final = cargo_novo or cargo
        cargo_final = " ".join(cargo_final.split())
        if not cargo_final:
            self.add_error("cargo", "Informe o cargo.")
        else:
            cleaned_data["cargo"] = cargo_final
        return cleaned_data


class OficioNumeracaoForm(forms.Form):
    oficio = forms.CharField(required=False)
    protocolo = forms.CharField(required=False)

    def clean_oficio(self):
        return normalize_oficio_num(self.cleaned_data.get("oficio"))

    def clean_protocolo(self):
        protocolo = normalize_protocolo_num(self.cleaned_data.get("protocolo"))
        if protocolo and len(protocolo) != 9:
            raise forms.ValidationError("Protocolo deve conter 9 digitos.")
        return protocolo


class MotoristaTransporteForm(forms.Form):
    motorista_nome = forms.CharField(required=False)
    motorista_oficio_numero = forms.CharField(required=False)
    motorista_oficio_ano = forms.CharField(required=False)
    motorista_oficio = forms.CharField(required=False)
    motorista_protocolo = forms.CharField(required=False)

    def clean_motorista_nome(self):
        return normalize_upper_text(self.cleaned_data.get("motorista_nome"))

    def clean_motorista_oficio_numero(self):
        digits = normalize_digits(self.cleaned_data.get("motorista_oficio_numero"))
        return str(int(digits)) if digits else ""

    def clean_motorista_oficio_ano(self):
        digits = normalize_digits(self.cleaned_data.get("motorista_oficio_ano"))
        if not digits:
            return ""
        return digits[-4:]

    def clean_motorista_oficio(self):
        return normalize_oficio_num(self.cleaned_data.get("motorista_oficio"))

    def clean_motorista_protocolo(self):
        protocolo = normalize_protocolo_num(self.cleaned_data.get("motorista_protocolo"))
        if protocolo and len(protocolo) != 9:
            raise forms.ValidationError("Protocolo do motorista deve conter 9 digitos.")
        return protocolo

    def clean(self):
        cleaned_data = super().clean()
        oficio_raw = cleaned_data.get("motorista_oficio") or ""
        numero = cleaned_data.get("motorista_oficio_numero") or ""
        ano = cleaned_data.get("motorista_oficio_ano") or ""

        if oficio_raw and not numero:
            parsed_num, parsed_ano = split_oficio_num(oficio_raw)
            if parsed_num is not None:
                numero = str(parsed_num)
            if parsed_ano is not None:
                ano = str(parsed_ano)

        if numero and not ano:
            ano = str(timezone.localdate().year)

        if ano and not numero:
            ano = ""

        cleaned_data["motorista_oficio_numero"] = numero
        cleaned_data["motorista_oficio_ano"] = ano
        cleaned_data["motorista_oficio"] = format_oficio_num(numero, ano)
        return cleaned_data


class PlanoTrabalhoForm(forms.ModelForm):
    valor_unitario = forms.CharField(required=True)
    valor_total_calculado = forms.CharField(required=False)
    possui_coordenador_municipal = forms.TypedChoiceField(
        label="Possui coordenador municipal?",
        choices=(("nao", "Nao"), ("sim", "Sim")),
        coerce=lambda value: str(value).strip().lower() == "sim",
        empty_value=False,
        required=True,
        widget=forms.Select(),
    )
    metas_json = forms.CharField(required=False, widget=forms.HiddenInput())
    atividades_json = forms.CharField(required=False, widget=forms.HiddenInput())
    atividades_selecionadas = forms.MultipleChoiceField(
        required=False,
        choices=(),
        widget=forms.CheckboxSelectMultiple(),
    )
    recursos_json = forms.CharField(required=False, widget=forms.HiddenInput())
    locais_json = forms.CharField(required=False, widget=forms.HiddenInput())
    coordenador_municipal_nome = forms.CharField(required=False)
    coordenador_municipal_cargo = forms.CharField(required=False)
    coordenador_municipal_cidade = forms.CharField(required=False)

    class Meta:
        model = PlanoTrabalho
        fields = [
            "numero",
            "ano",
            "sigla_unidade",
            "programa_projeto",
            "destino",
            "solicitante",
            "contexto_solicitacao",
            "data_inicio",
            "data_fim",
            "horario_atendimento",
            "efetivo_formatado",
            "estrutura_apoio",
            "quantidade_servidores",
            "composicao_diarias",
            "valor_unitario",
            "valor_total_calculado",
            "coordenador_plano",
            "coordenador_municipal",
            "possui_coordenador_municipal",
            "texto_override",
        ]
        widgets = {
            "data_inicio": forms.DateInput(attrs={"type": "date"}, format="%Y-%m-%d"),
            "data_fim": forms.DateInput(attrs={"type": "date"}, format="%Y-%m-%d"),
            "contexto_solicitacao": forms.Textarea(attrs={"rows": 3}),
            "estrutura_apoio": forms.Textarea(attrs={"rows": 3}),
            "texto_override": forms.Textarea(attrs={"rows": 4}),
            "valor_total_calculado": forms.TextInput(attrs={"readonly": "readonly"}),
        }

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.fields["numero"].label = "Numero do plano"
        self.fields["ano"].label = "Ano do plano"
        self.fields["sigla_unidade"].label = "Sigla/unidade"
        self.fields["programa_projeto"].label = "Programa/projeto base"
        self.fields["destino"].label = "Destino/municipio"
        self.fields["solicitante"].label = "Solicitante"
        self.fields["contexto_solicitacao"].label = "Contexto complementar"
        self.fields["data_inicio"].label = "Data inicial"
        self.fields["data_fim"].label = "Data final"
        self.fields["atividades_selecionadas"].label = "Atividades selecionadas"
        self.fields["horario_atendimento"].label = "Horario de atendimento"
        self.fields["efetivo_formatado"].label = "Efetivo"
        self.fields["estrutura_apoio"].label = "Estrutura/unidade movel (opcional)"
        self.fields["quantidade_servidores"].label = "Quantidade de servidores"
        self.fields["composicao_diarias"].label = "Composicao das diarias"
        self.fields["valor_unitario"].label = "Valor unitario"
        self.fields["valor_total_calculado"].label = "Valor total"
        self.fields["coordenador_plano"].label = "Coordenador administrativo do plano"
        self.fields["coordenador_municipal"].label = "Coordenador municipal"
        self.fields["texto_override"].label = "Observacoes internas"
        self.fields["numero"].required = False
        self.fields["ano"].required = False
        self.fields["sigla_unidade"].required = False
        self.fields["coordenador_municipal"].required = False
        self.fields["coordenador_municipal"].queryset = CoordenadorMunicipal.objects.order_by(
            "nome"
        )
        self.fields["atividades_selecionadas"].choices = [
            (atividade, atividade) for atividade in ATIVIDADES_ORDEM_FIXA
        ]
        self.fields["coordenador_municipal"].widget.attrs.update(
            {
                "data-autocomplete-url": "/api/coordenadores-municipais/",
                "data-autocomplete-type": "coordenador-municipal",
            }
        )
        self.fields["coordenador_plano"].required = False
        self.fields["coordenador_plano"].queryset = Viajante.objects.order_by("nome")
        self.fields["coordenador_plano"].widget.attrs.update(
            {
                "data-autocomplete-url": "/api/servidores/",
                "data-autocomplete-type": "servidor",
            }
        )

        possui_val = self.initial.get("possui_coordenador_municipal")
        if isinstance(possui_val, bool):
            self.initial["possui_coordenador_municipal"] = "sim" if possui_val else "nao"
        elif hasattr(self.instance, "pk") and self.instance.pk:
            self.initial["possui_coordenador_municipal"] = (
                "sim" if self.instance.possui_coordenador_municipal else "nao"
            )

        atividades_initial = self.initial.get("atividades_selecionadas")
        if not atividades_initial and hasattr(self.instance, "pk") and self.instance.pk:
            atividades_initial = [
                item.descricao
                for item in self.instance.atividades.all().order_by("ordem", "id")
                if (item.descricao or "").strip()
            ]
        if not atividades_initial:
            atividades_initial = self._parse_ordered_text_items_from_raw(
                self.initial.get("atividades_json")
            )
        atividades_initial = normalize_atividades_selecionadas(atividades_initial or [])
        if atividades_initial:
            self.initial["atividades_selecionadas"] = atividades_initial
            if not (self.initial.get("atividades_json") or "").strip():
                self.initial["atividades_json"] = json.dumps(
                    [{"descricao": value} for value in atividades_initial],
                    ensure_ascii=False,
                )
        metas_initial = self.initial.get("metas_json")
        if not metas_initial and atividades_initial:
            self.initial["metas_json"] = json.dumps(
                [{"descricao": value} for value in metas_from_atividades(atividades_initial)],
                ensure_ascii=False,
            )

        self._parsed_metas: list[str] = []
        self._parsed_atividades: list[str] = []
        self._parsed_recursos: list[str] = []
        self._parsed_locais: list[dict[str, object]] = []

        for field in self.fields.values():
            css = field.widget.attrs.get("class", "")
            field.widget.attrs["class"] = f"{css} input-field".strip()

    @staticmethod
    def _parse_decimal_input(raw_value: str | Decimal | None) -> Decimal | None:
        if raw_value in (None, ""):
            return None
        if isinstance(raw_value, Decimal):
            return raw_value
        try:
            normalized = str(raw_value).strip().replace("R$", "").replace("r$", "").replace(" ", "")
            if "," in normalized and "." in normalized:
                if normalized.rfind(",") > normalized.rfind("."):
                    normalized = normalized.replace(".", "").replace(",", ".")
                else:
                    normalized = normalized.replace(",", "")
            elif "," in normalized:
                normalized = normalized.replace(".", "").replace(",", ".")
            else:
                normalized = normalized.replace(",", "")
            return Decimal(normalized)
        except (InvalidOperation, TypeError, ValueError):
            return None

    @staticmethod
    def _parse_composicao_fator(raw_value: str) -> Decimal:
        raw = (raw_value or "").strip()
        if not raw:
            return Decimal("1")
        pattern = re.compile(
            r"(?P<qtd>\d+(?:[.,]\d+)?)\s*x\s*(?P<pct>\d+(?:[.,]\d+)?)\s*%",
            re.IGNORECASE,
        )
        fator = Decimal("0")
        found = False
        for match in pattern.finditer(raw):
            found = True
            qtd = PlanoTrabalhoForm._parse_decimal_input(match.group("qtd")) or Decimal("0")
            pct = PlanoTrabalhoForm._parse_decimal_input(match.group("pct")) or Decimal("0")
            fator += qtd * (pct / Decimal("100"))
        if found and fator > 0:
            return fator
        fallback = PlanoTrabalhoForm._parse_decimal_input(raw)
        if fallback and fallback > 0:
            return fallback
        return Decimal("1")

    def _parse_ordered_text_items(self, field_name: str, *, label: str) -> list[str]:
        raw_payload = (self.cleaned_data.get(field_name) or "").strip()
        if not raw_payload:
            return []
        try:
            parsed = json.loads(raw_payload)
        except json.JSONDecodeError:
            self.add_error(field_name, f"Nao foi possivel ler a lista de {label}.")
            return []
        if not isinstance(parsed, list):
            self.add_error(field_name, f"Formato invalido para lista de {label}.")
            return []
        return self._normalize_ordered_text_items(parsed)

    def _parse_ordered_text_items_from_raw(
        self,
        raw_payload: str | None,
        *,
        label: str = "itens",
    ) -> list[str]:
        raw_payload = (raw_payload or "").strip()
        if not raw_payload:
            return []
        try:
            parsed = json.loads(raw_payload)
        except json.JSONDecodeError:
            return []
        if not isinstance(parsed, list):
            return []
        return self._normalize_ordered_text_items(parsed)

    @staticmethod
    def _normalize_ordered_text_items(parsed: list[object]) -> list[str]:
        normalized: list[str] = []
        for item in parsed:
            text = ""
            if isinstance(item, str):
                text = item
            elif isinstance(item, dict):
                text = str(item.get("descricao", "") or item.get("texto", ""))
            text = " ".join(str(text).split())
            if text:
                normalized.append(text)
        return normalized

    def _parse_locais(self) -> list[dict[str, object]]:
        raw_payload = (self.cleaned_data.get("locais_json") or "").strip()
        if not raw_payload:
            return []
        try:
            parsed = json.loads(raw_payload)
        except json.JSONDecodeError:
            self.add_error("locais_json", "Nao foi possivel ler a lista de locais.")
            return []
        if not isinstance(parsed, list):
            self.add_error("locais_json", "Formato invalido para a lista de locais.")
            return []

        normalized: list[dict[str, object]] = []
        for item in parsed:
            if isinstance(item, str):
                local_raw = item
                data_raw = ""
            elif isinstance(item, dict):
                local_raw = item.get("local", "")
                data_raw = item.get("data", "")
            else:
                continue
            local = " ".join(str(local_raw or "").split())
            if not local:
                continue
            data_val = parse_date(str(data_raw or "").strip()) if data_raw else None
            normalized.append({"local": local, "data": data_val})
        return normalized

    def clean_valor_unitario(self):
        raw = (self.data.get(self.add_prefix("valor_unitario")) or "").strip()
        value = self._parse_decimal_input(raw)
        if value is None or value <= 0:
            raise forms.ValidationError("Informe um valor unitario valido.")
        return value.quantize(Decimal("0.01"))

    def clean_valor_total_calculado(self):
        raw = (self.data.get(self.add_prefix("valor_total_calculado")) or "").strip()
        if not raw:
            return None
        value = self._parse_decimal_input(raw)
        if value is None or value <= 0:
            raise forms.ValidationError("Informe um valor total valido.")
        return value.quantize(Decimal("0.01"))

    def clean(self):
        cleaned_data = super().clean()

        atividades_selecionadas = list(cleaned_data.get("atividades_selecionadas") or [])
        if not atividades_selecionadas:
            atividades_selecionadas = self._parse_ordered_text_items(
                "atividades_json",
                label="atividades",
            )
        atividades_ordenadas = normalize_atividades_selecionadas(atividades_selecionadas)
        self._parsed_atividades = atividades_ordenadas
        self._parsed_metas = metas_from_atividades(atividades_ordenadas)
        cleaned_data["atividades_selecionadas"] = atividades_ordenadas
        cleaned_data["atividades_json"] = json.dumps(
            [{"descricao": value} for value in atividades_ordenadas],
            ensure_ascii=False,
        )
        cleaned_data["metas_json"] = json.dumps(
            [{"descricao": value} for value in self._parsed_metas],
            ensure_ascii=False,
        )
        self._parsed_recursos = self._parse_ordered_text_items("recursos_json", label="recursos")
        self._parsed_locais = self._parse_locais()

        if not (cleaned_data.get("destino") or "").strip():
            self.add_error("destino", "Informe o destino do plano.")
        if not (cleaned_data.get("solicitante") or "").strip():
            self.add_error("solicitante", "Informe o solicitante.")
        if not cleaned_data.get("data_inicio") or not cleaned_data.get("data_fim"):
            self.add_error("data_inicio", "Informe as datas do evento.")
        if cleaned_data.get("data_inicio") and cleaned_data.get("data_fim"):
            if cleaned_data["data_fim"] < cleaned_data["data_inicio"]:
                self.add_error("data_fim", "A data final deve ser posterior a data inicial.")
        if not (cleaned_data.get("horario_atendimento") or "").strip():
            self.add_error("horario_atendimento", "Informe o horario de atendimento.")
        if not (cleaned_data.get("efetivo_formatado") or "").strip():
            self.add_error("efetivo_formatado", "Informe o efetivo.")
        if not (cleaned_data.get("composicao_diarias") or "").strip():
            self.add_error("composicao_diarias", "Informe a composicao de diarias.")
        if not cleaned_data.get("coordenador_plano"):
            self.add_error("coordenador_plano", "Informe o coordenador administrativo.")

        if not self._parsed_atividades:
            self.add_error(
                "atividades_selecionadas",
                "Selecione ao menos 1 atividade.",
            )
        if not self._parsed_recursos:
            self.add_error("recursos_json", "Informe ao menos 1 recurso.")
        if not self._parsed_locais:
            self.add_error("locais_json", "Informe ao menos 1 local de atuacao.")

        possui_municipal = bool(cleaned_data.get("possui_coordenador_municipal"))
        coordenador_municipal = cleaned_data.get("coordenador_municipal")
        novo_nome = " ".join((cleaned_data.get("coordenador_municipal_nome") or "").split())
        novo_cargo = " ".join((cleaned_data.get("coordenador_municipal_cargo") or "").split())
        novo_cidade = " ".join((cleaned_data.get("coordenador_municipal_cidade") or "").split())

        if possui_municipal:
            if not coordenador_municipal and not (novo_nome and novo_cargo and novo_cidade):
                self.add_error(
                    "coordenador_municipal",
                    "Selecione ou cadastre um coordenador municipal.",
                )
            if bool(novo_nome or novo_cargo or novo_cidade) and not (
                novo_nome and novo_cargo and novo_cidade
            ):
                self.add_error(
                    "coordenador_municipal_nome",
                    "Para cadastrar novo coordenador municipal, preencha nome, cargo e cidade.",
                )
        else:
            cleaned_data["coordenador_municipal"] = None
            cleaned_data["coordenador_municipal_nome"] = ""
            cleaned_data["coordenador_municipal_cargo"] = ""
            cleaned_data["coordenador_municipal_cidade"] = ""

        valor_total = cleaned_data.get("valor_total_calculado")
        valor_unitario = cleaned_data.get("valor_unitario")
        qtd_servidores = int(cleaned_data.get("quantidade_servidores") or 0)
        if not valor_total and valor_unitario and qtd_servidores > 0:
            fator = self._parse_composicao_fator(cleaned_data.get("composicao_diarias") or "")
            cleaned_data["valor_total_calculado"] = (
                valor_unitario * Decimal(qtd_servidores) * fator
            ).quantize(Decimal("0.01"))

        return cleaned_data

    @property
    def parsed_metas(self) -> list[str]:
        return list(self._parsed_metas)

    @property
    def parsed_atividades(self) -> list[str]:
        return list(self._parsed_atividades)

    @property
    def parsed_recursos(self) -> list[str]:
        return list(self._parsed_recursos)

    @property
    def parsed_locais(self) -> list[dict[str, object]]:
        return list(self._parsed_locais)


class PlanoTrabalhoStep1Form(forms.Form):
    solicitantes = forms.MultipleChoiceField(
        required=True,
        choices=[(item, item) for item in SOLICITANTES_ORDEM_FIXA],
        widget=forms.CheckboxSelectMultiple(),
    )
    solicitante_pcpr_nome = forms.CharField(required=False)
    data_unica = forms.BooleanField(required=False)
    data_inicio = forms.DateField(
        required=True,
        widget=forms.DateInput(attrs={"type": "date"}),
    )
    data_fim = forms.DateField(
        required=False,
        widget=forms.DateInput(attrs={"type": "date"}),
    )
    horario_inicio = forms.TimeField(
        required=True,
        widget=forms.TimeInput(attrs={"type": "time"}),
    )
    horario_fim = forms.TimeField(
        required=True,
        widget=forms.TimeInput(attrs={"type": "time"}),
    )

    def clean(self):
        cleaned_data = super().clean()
        solicitantes = normalize_solicitantes(cleaned_data.get("solicitantes") or [])
        if not solicitantes:
            self.add_error("solicitantes", "Selecione ao menos um solicitante.")
        cleaned_data["solicitantes"] = solicitantes

        nome_pcpr = " ".join((cleaned_data.get("solicitante_pcpr_nome") or "").split())
        if SOLICITANTE_PCPR in solicitantes and not nome_pcpr:
            self.add_error(
                "solicitante_pcpr_nome",
                "Informe o nome do solicitante para PCPR na Comunidade.",
            )
        cleaned_data["solicitante_pcpr_nome"] = nome_pcpr

        data_inicio = cleaned_data.get("data_inicio")
        data_fim = cleaned_data.get("data_fim")
        if cleaned_data.get("data_unica"):
            data_fim = data_inicio
            cleaned_data["data_fim"] = data_fim
        elif not data_fim:
            self.add_error("data_fim", "Informe a data final.")
        if data_inicio and data_fim and data_fim < data_inicio:
            self.add_error("data_fim", "A data final deve ser igual ou posterior a inicial.")

        horario_inicio = cleaned_data.get("horario_inicio")
        horario_fim = cleaned_data.get("horario_fim")
        if horario_inicio and horario_fim and horario_fim <= horario_inicio:
            self.add_error("horario_fim", "O horario final deve ser posterior ao inicial.")
        cleaned_data["horario_atendimento"] = formatar_horario_intervalo(
            horario_inicio,
            horario_fim,
        )
        return cleaned_data


class PlanoTrabalhoStep2Form(forms.Form):
    efetivo_json = forms.CharField(required=True, widget=forms.HiddenInput())
    unidade_movel = forms.TypedChoiceField(
        choices=(("nao", "Nao"), ("sim", "Sim")),
        coerce=lambda value: str(value).strip().lower() == "sim",
        required=True,
    )
    coordenador_plano = forms.ModelChoiceField(
        queryset=Viajante.objects.order_by("nome"),
        required=False,
    )
    coordenador_plano_nome = forms.CharField(required=True)
    coordenador_plano_cargo = forms.CharField(required=True)
    coordenador_municipal = forms.ModelChoiceField(
        queryset=CoordenadorMunicipal.objects.order_by("nome"),
        required=False,
    )
    coordenador_municipal_nome = forms.CharField(required=False)
    coordenador_municipal_cargo = forms.CharField(required=False)
    coordenador_municipal_cidade = forms.CharField(required=False)

    def __init__(self, *args, **kwargs):
        self.permite_municipal = bool(kwargs.pop("permite_municipal", False))
        super().__init__(*args, **kwargs)
        self.fields["coordenador_plano"].widget.attrs.update(
            {
                "data-autocomplete-url": "/api/servidores/",
                "data-autocomplete-type": "servidor",
            }
        )
        self.fields["coordenador_municipal"].widget.attrs.update(
            {
                "data-autocomplete-url": "/api/coordenadores-municipais/",
                "data-autocomplete-type": "coordenador-municipal",
            }
        )
        if not self.initial.get("coordenador_plano_nome"):
            self.initial["coordenador_plano_nome"] = DEFAULT_COORDENADOR_PLANO_NOME
        if not self.initial.get("coordenador_plano_cargo"):
            self.initial["coordenador_plano_cargo"] = DEFAULT_COORDENADOR_PLANO_CARGO
        for field in self.fields.values():
            css = field.widget.attrs.get("class", "")
            field.widget.attrs["class"] = f"{css} input-field".strip()

        self._parsed_efetivo: list[dict[str, object]] = []

    def clean(self):
        cleaned_data = super().clean()
        raw_payload = (cleaned_data.get("efetivo_json") or "").strip()
        try:
            parsed_payload = json.loads(raw_payload) if raw_payload else []
        except json.JSONDecodeError:
            self.add_error("efetivo_json", "Nao foi possivel ler o efetivo informado.")
            parsed_payload = []

        efetivo_rows = normalize_efetivo_payload(parsed_payload if isinstance(parsed_payload, list) else [])
        if not efetivo_rows:
            self.add_error("efetivo_json", "Informe ao menos um item de efetivo.")
        self._parsed_efetivo = efetivo_rows
        cleaned_data["efetivo_json"] = json.dumps(efetivo_rows, ensure_ascii=False)

        coordenador_plano = cleaned_data.get("coordenador_plano")
        nome = " ".join((cleaned_data.get("coordenador_plano_nome") or "").split())
        cargo = " ".join((cleaned_data.get("coordenador_plano_cargo") or "").split())
        if coordenador_plano:
            nome = coordenador_plano.nome
            cargo = coordenador_plano.cargo
        if not nome:
            self.add_error("coordenador_plano_nome", "Informe o nome do coordenador do plano.")
        if not cargo:
            self.add_error("coordenador_plano_cargo", "Informe o cargo do coordenador do plano.")
        cleaned_data["coordenador_plano_nome"] = nome or DEFAULT_COORDENADOR_PLANO_NOME
        cleaned_data["coordenador_plano_cargo"] = cargo or DEFAULT_COORDENADOR_PLANO_CARGO

        coord_municipal = cleaned_data.get("coordenador_municipal")
        novo_nome = " ".join((cleaned_data.get("coordenador_municipal_nome") or "").split())
        novo_cargo = " ".join((cleaned_data.get("coordenador_municipal_cargo") or "").split())
        novo_cidade = " ".join((cleaned_data.get("coordenador_municipal_cidade") or "").split())
        if not self.permite_municipal:
            cleaned_data["coordenador_municipal"] = None
            cleaned_data["coordenador_municipal_nome"] = ""
            cleaned_data["coordenador_municipal_cargo"] = ""
            cleaned_data["coordenador_municipal_cidade"] = ""
        else:
            preenchimento_parcial = bool(novo_nome or novo_cargo or novo_cidade)
            if preenchimento_parcial and not (novo_nome and novo_cargo and novo_cidade):
                self.add_error(
                    "coordenador_municipal_nome",
                    "Para cadastrar coordenador municipal, preencha nome, cargo e cidade.",
                )
            if coord_municipal and preenchimento_parcial:
                self.add_error(
                    "coordenador_municipal",
                    "Escolha entre selecionar um coordenador municipal ou cadastrar um novo.",
                )
            cleaned_data["coordenador_municipal_nome"] = novo_nome
            cleaned_data["coordenador_municipal_cargo"] = novo_cargo
            cleaned_data["coordenador_municipal_cidade"] = novo_cidade

        return cleaned_data

    @property
    def parsed_efetivo(self) -> list[dict[str, object]]:
        return list(self._parsed_efetivo)


class PlanoTrabalhoStep3Form(forms.Form):
    composicao_diarias = forms.CharField(required=True)
    valor_unitario = forms.CharField(required=True)
    valor_total_calculado = forms.CharField(required=False)
    recursos_json = forms.CharField(required=True, widget=forms.HiddenInput())

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        for field in self.fields.values():
            css = field.widget.attrs.get("class", "")
            field.widget.attrs["class"] = f"{css} input-field".strip()
        self._parsed_recursos: list[str] = []

    @staticmethod
    def _parse_decimal_input(raw_value: str | Decimal | None) -> Decimal | None:
        if raw_value in (None, ""):
            return None
        if isinstance(raw_value, Decimal):
            return raw_value
        try:
            normalized = str(raw_value).strip().replace("R$", "").replace("r$", "").replace(" ", "")
            if "," in normalized and "." in normalized:
                if normalized.rfind(",") > normalized.rfind("."):
                    normalized = normalized.replace(".", "").replace(",", ".")
                else:
                    normalized = normalized.replace(",", "")
            elif "," in normalized:
                normalized = normalized.replace(".", "").replace(",", ".")
            else:
                normalized = normalized.replace(",", "")
            return Decimal(normalized)
        except (InvalidOperation, TypeError, ValueError):
            return None

    def clean_valor_unitario(self):
        raw = (self.data.get(self.add_prefix("valor_unitario")) or "").strip()
        value = self._parse_decimal_input(raw)
        if value is None or value <= 0:
            raise forms.ValidationError("Informe um valor unitario valido.")
        return value.quantize(Decimal("0.01"))

    def clean(self):
        cleaned_data = super().clean()
        composicao = " ".join((cleaned_data.get("composicao_diarias") or "").split())
        if not composicao:
            self.add_error("composicao_diarias", "Informe a composicao de diarias.")
        cleaned_data["composicao_diarias"] = composicao

        raw_payload = (cleaned_data.get("recursos_json") or "").strip()
        try:
            parsed_payload = json.loads(raw_payload) if raw_payload else []
        except json.JSONDecodeError:
            self.add_error("recursos_json", "Nao foi possivel ler os recursos informados.")
            parsed_payload = []
        recursos: list[str] = []
        if isinstance(parsed_payload, list):
            for item in parsed_payload:
                if isinstance(item, str):
                    text = " ".join(item.split())
                elif isinstance(item, dict):
                    text = " ".join(str(item.get("descricao", "")).split())
                else:
                    text = ""
                if text:
                    recursos.append(text)
        if not recursos:
            self.add_error("recursos_json", "Informe ao menos um recurso necessario.")
        self._parsed_recursos = recursos
        cleaned_data["recursos_json"] = json.dumps(
            [{"descricao": item} for item in recursos],
            ensure_ascii=False,
        )
        return cleaned_data

    @property
    def parsed_recursos(self) -> list[str]:
        return list(self._parsed_recursos)


class OrdemServicoForm(forms.ModelForm):
    class Meta:
        model = OrdemServico
        fields = [
            "numero",
            "ano",
            "referencia",
            "determinante_nome",
            "determinante_cargo",
            "finalidade",
            "texto_override",
        ]
        widgets = {
            "finalidade": forms.Textarea(attrs={"rows": 4}),
            "texto_override": forms.Textarea(attrs={"rows": 6}),
        }

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.fields["numero"].required = False
        self.fields["ano"].required = False
        for field in self.fields.values():
            css = field.widget.attrs.get("class", "")
            field.widget.attrs["class"] = f"{css} input-field".strip()


class JustificativaForm(forms.ModelForm):
    justificativa_modelo = forms.ChoiceField(required=False)

    class Meta:
        model = Oficio
        fields = ["justificativa_modelo", "justificativa_texto"]
        widgets = {
            "justificativa_texto": forms.Textarea(attrs={"rows": 12}),
        }

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.fields["justificativa_modelo"].choices = [
            ("", "Selecione")
        ] + [(key, item["label"]) for key, item in JUSTIFICATIVA_TEMPLATES.items()]
        for field in self.fields.values():
            css = field.widget.attrs.get("class", "")
            field.widget.attrs["class"] = f"{css} input-field".strip()
