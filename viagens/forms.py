from __future__ import annotations

from datetime import datetime, time

from django import forms
from django.utils import timezone

from .models import Cidade, Estado, Trecho, Viajante
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
