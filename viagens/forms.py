from __future__ import annotations

from datetime import datetime, time

from django import forms

from .models import Cidade, Estado, Trecho, Viajante


class TrechoForm(forms.ModelForm):
    origem_estado = forms.ModelChoiceField(
        queryset=Estado.objects.order_by("nome"),
        to_field_name="sigla",
        required=True,
    )
    destino_estado = forms.ModelChoiceField(
        queryset=Estado.objects.order_by("nome"),
        to_field_name="sigla",
        required=True,
    )
    origem_cidade = forms.ModelChoiceField(
        queryset=Cidade.objects.none(),
        required=True,
    )
    destino_cidade = forms.ModelChoiceField(
        queryset=Cidade.objects.none(),
        required=True,
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
            "saida_data": forms.DateInput(attrs={"type": "date"}),
            "saida_hora": forms.TimeInput(attrs={"type": "time"}),
            "chegada_data": forms.DateInput(attrs={"type": "date"}),
            "chegada_hora": forms.TimeInput(attrs={"type": "time"}),
        }

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)

        self._apply_default_uf()
        self._apply_widget_attrs()
        self._set_city_queryset("origem_estado", "origem_cidade")
        self._set_city_queryset("destino_estado", "destino_cidade")

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
