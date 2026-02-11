import re

from django import forms

from viagens.models import Cidade, OficioConfig, Viajante


class OficioConfigForm(forms.ModelForm):
    class Meta:
        model = OficioConfig
        fields = [
            "unidade_nome",
            "origem_nome",
            "plano_divisao",
            "plano_unidade",
            "plano_sede",
            "plano_nome_chefia",
            "plano_cargo_chefia",
            "cep",
            "logradouro",
            "bairro",
            "cidade",
            "uf",
            "numero",
            "complemento",
            "telefone",
            "email",
            "assinante",
            "sede_cidade_default",
        ]

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        base_class = "input-field"
        for field in self.fields.values():
            existing = field.widget.attrs.get("class", "")
            field.widget.attrs["class"] = f"{existing} {base_class}".strip()

        self.fields["assinante"].queryset = Viajante.objects.order_by("nome")
        self.fields["assinante"].empty_label = "Selecione"
        self.fields["assinante"].label_from_instance = (
            lambda obj: f"{obj.nome} - {obj.cargo}".strip(" -")
        )
        self.fields["assinante"].widget.attrs.update(
            {
                "data-autocomplete-url": "/api/assinantes/",
                "data-autocomplete-type": "servidor",
            }
        )
        self.fields["sede_cidade_default"].queryset = Cidade.objects.order_by("nome")
        self.fields["sede_cidade_default"].empty_label = "Selecione"
        self.fields["sede_cidade_default"].widget.attrs.update(
            {
                "data-autocomplete-url": "/api/cidades-busca/",
                "data-autocomplete-type": "cidade",
                "data-role": "sede-cidade",
            }
        )

        readonly_fields = ["logradouro", "bairro", "cidade", "uf"]
        for field_name in readonly_fields:
            self.fields[field_name].widget.attrs.update({"readonly": "readonly"})

        self.fields["cep"].widget.attrs.update(
            {
                "inputmode": "numeric",
                "placeholder": "00000-000",
                "maxlength": "9",
                "data-cep-input": "true",
            }
        )
        self.fields["telefone"].widget.attrs.update(
            {"data-mask": "telefone", "inputmode": "numeric"}
        )

    def clean_unidade_nome(self):
        value = (self.cleaned_data.get("unidade_nome") or "").strip()
        if not value:
            raise forms.ValidationError("Informe o nome da unidade.")
        return value.upper()

    def clean_origem_nome(self):
        value = (self.cleaned_data.get("origem_nome") or "").strip()
        if not value:
            raise forms.ValidationError("Informe o nome da origem.")
        return value.upper()

    def clean_cep(self):
        raw = (self.cleaned_data.get("cep") or "").strip()
        digits = re.sub(r"\D", "", raw)
        if len(digits) != 8:
            raise forms.ValidationError("Informe um CEP valido.")
        return f"{digits[:5]}-{digits[5:]}"

    def clean_numero(self):
        value = (self.cleaned_data.get("numero") or "").strip()
        if not value:
            raise forms.ValidationError("Informe o numero.")
        return value

    def clean_uf(self):
        return (self.cleaned_data.get("uf") or "").strip().upper()

    def clean_sede_cidade_default(self):
        cidade = self.cleaned_data.get("sede_cidade_default")
        if cidade and (cidade.estado.sigla or "").strip().upper() != "PR":
            raise forms.ValidationError("A sede padrao deve ser no estado do PR.")
        return cidade
