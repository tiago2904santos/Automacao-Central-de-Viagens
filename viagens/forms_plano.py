from __future__ import annotations

from django import forms

from .models import PlanoAtividadeOpcao, Viajante


class PlanoStep1Form(forms.Form):
    destino_uf = forms.CharField(max_length=2, required=True, initial="PR")
    destino_cidade = forms.IntegerField(required=True)
    data_evento_inicio = forms.DateField(required=True, widget=forms.DateInput(attrs={"type": "date"}))
    fim_mesmo_dia = forms.BooleanField(required=False)
    data_evento_fim = forms.DateField(required=False, widget=forms.DateInput(attrs={"type": "date"}))
    horario_inicio = forms.TimeField(required=True, widget=forms.TimeInput(attrs={"type": "time"}))
    horario_fim = forms.TimeField(required=True, widget=forms.TimeInput(attrs={"type": "time"}))
    servidor_responsavel = forms.ModelChoiceField(queryset=Viajante.objects.order_by("nome"), required=True)
    metas = forms.CharField(required=False, widget=forms.Textarea(attrs={"rows": 5}))
    incluir_microonibus = forms.BooleanField(required=False)

    def clean_destino_uf(self):
        return (self.cleaned_data.get("destino_uf") or "").strip().upper() or "PR"

    def clean(self):
        cleaned = super().clean()
        inicio = cleaned.get("data_evento_inicio")
        fim = cleaned.get("data_evento_fim")
        mesmo_dia = bool(cleaned.get("fim_mesmo_dia"))
        if mesmo_dia and inicio:
            cleaned["data_evento_fim"] = inicio
        elif not mesmo_dia:
            if not fim:
                self.add_error("data_evento_fim", "Informe a data de fim do evento.")
            elif inicio and fim < inicio:
                self.add_error("data_evento_fim", "A data de fim deve ser maior ou igual à inicial.")
        return cleaned


class PlanoStep2Form(forms.Form):
    atividades = forms.MultipleChoiceField(required=False, widget=forms.CheckboxSelectMultiple)
    atividades_ordenadas = forms.CharField(required=False, widget=forms.HiddenInput())
    quantidade_servidores = forms.IntegerField(min_value=1, required=True, initial=1)
    periodos_json = forms.CharField(required=False, widget=forms.HiddenInput())

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.fields["atividades"].choices = [
            (str(item.id), item.titulo)
            for item in PlanoAtividadeOpcao.objects.filter(ativo=True).order_by("ordem", "titulo")
        ]
