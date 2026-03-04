# Formulários do fluxo guiado por Evento (wizard)
from __future__ import annotations

from django import forms

from .models import Cidade, Estado, Evento


class EventoGuiadoStep1Form(forms.ModelForm):
    """Etapa 1: cadastro do evento (título, cidade base, datas, tem_convite, tipo_demanda)."""

    titulo = forms.CharField(
        max_length=255,
        required=True,
        label="Título / Nome do evento",
        widget=forms.TextInput(attrs={"class": "input-field", "placeholder": "Ex.: Missão técnica região X"}),
    )
    data_inicio = forms.DateField(
        required=False,
        label="Data início",
        widget=forms.DateInput(attrs={"type": "date", "class": "input-field"}),
    )
    data_fim = forms.DateField(
        required=False,
        label="Data fim",
        widget=forms.DateInput(attrs={"type": "date", "class": "input-field"}),
    )

    class Meta:
        model = Evento
        fields = [
            "titulo",
            "cidade_base",
            "data_inicio",
            "data_fim",
            "tem_convite_ou_oficio_evento",
            "tipo_demanda",
        ]
        widgets = {
            "cidade_base": forms.Select(attrs={"class": "input-field"}),
            "tem_convite_ou_oficio_evento": forms.CheckboxInput(attrs={"class": "input-field"}),
            "tipo_demanda": forms.Select(attrs={"class": "input-field"}),
        }

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.fields["cidade_base"].queryset = Cidade.objects.select_related("estado").order_by("nome")
        self.fields["cidade_base"].required = False
        self.fields["cidade_base"].empty_label = "Selecione (opcional)"
        self.fields["tem_convite_ou_oficio_evento"].label = "Tem ofício solicitando ou convite do evento?"
        self.fields["tem_convite_ou_oficio_evento"].required = False
        self.fields["tem_convite_ou_oficio_evento"].initial = True
        self.fields["tipo_demanda"].required = False
        self.fields["tipo_demanda"].label = "Intuito / tipo do evento"
        self.fields["tipo_demanda"].help_text = "Demanda operacional do evento. O padrão é «Outro»."

    def clean(self):
        data = super().clean()
        titulo = (data.get("titulo") or "").strip()
        if not titulo:
            self.add_error("titulo", "O título é obrigatório.")
        data_inicio = data.get("data_inicio")
        data_fim = data.get("data_fim")
        if data_inicio and data_fim and data_fim < data_inicio:
            self.add_error("data_fim", "A data fim deve ser maior ou igual à data início.")
        return data


class EventoRoteiroStep2Form(forms.Form):
    """Etapa 2: um roteiro do evento (ida + retorno). Duração em HH:MM (ex.: 6:30). Chegada calculada no servidor."""

    origem_cidade = forms.ModelChoiceField(
        queryset=Cidade.objects.none(),
        required=False,
        label="Origem (cidade)",
        widget=forms.Select(attrs={"class": "input-field"}),
    )
    destino_estado = forms.ModelChoiceField(
        queryset=Estado.objects.order_by("sigla"),
        required=False,
        label="Destino (UF)",
        widget=forms.Select(attrs={"class": "input-field"}),
    )
    destino_cidade = forms.ModelChoiceField(
        queryset=Cidade.objects.none(),
        required=False,
        label="Destino (cidade)",
        widget=forms.Select(attrs={"class": "input-field"}),
    )
    saida_data = forms.DateField(
        required=False,
        label="Saída (data)",
        widget=forms.DateInput(attrs={"type": "date", "class": "input-field"}),
    )
    saida_hora = forms.TimeField(
        required=False,
        label="Saída (hora)",
        widget=forms.TimeInput(attrs={"type": "time", "class": "input-field"}),
    )
    duracao_ida = forms.CharField(
        required=False,
        max_length=20,
        label="Duração ida (HH:MM)",
        widget=forms.TextInput(attrs={"class": "input-field", "placeholder": "Ex.: 6:30"}),
        help_text="Ex.: 6:30 para 6 horas e 30 minutos.",
    )
    retorno_saida_data = forms.DateField(
        required=False,
        label="Retorno – saída (data)",
        widget=forms.DateInput(attrs={"type": "date", "class": "input-field"}),
    )
    retorno_saida_hora = forms.TimeField(
        required=False,
        label="Retorno – saída (hora)",
        widget=forms.TimeInput(attrs={"type": "time", "class": "input-field"}),
    )
    duracao_retorno = forms.CharField(
        required=False,
        max_length=20,
        label="Duração retorno (HH:MM)",
        widget=forms.TextInput(attrs={"class": "input-field", "placeholder": "Ex.: 6:30"}),
    )

    def __init__(self, *args, evento=None, **kwargs):
        super().__init__(*args, **kwargs)
        self.evento = evento
        # Origen: todas as cidades (ou filtrar por estado padrão)
        self.fields["origem_cidade"].queryset = Cidade.objects.select_related("estado").order_by("nome")
        self.fields["origem_cidade"].empty_label = "Selecione"
        self.fields["destino_estado"].empty_label = "Selecione"
        self.fields["destino_cidade"].queryset = Cidade.objects.none()
        self.fields["destino_cidade"].empty_label = "Selecione"
        if evento and evento.cidade_base_id:
            self.fields["origem_cidade"].initial = evento.cidade_base_id
        estado_pr = Estado.objects.filter(sigla="PR").first()
        if estado_pr:
            self.fields["destino_estado"].initial = estado_pr.id
            self.fields["destino_cidade"].queryset = Cidade.objects.filter(estado=estado_pr).order_by("nome")

    def clean_duracao_ida(self):
        from .services.evento_roteiro_calculo import parse_duracao_minutos

        val = (self.cleaned_data.get("duracao_ida") or "").strip()
        if not val:
            return ""
        minutos = parse_duracao_minutos(val)
        if minutos is None or minutos < 0:
            self.add_error("duracao_ida", "Use o formato HH:MM (ex.: 6:30) ou apenas minutos (≥ 0).")
        return val

    def clean_duracao_retorno(self):
        from .services.evento_roteiro_calculo import parse_duracao_minutos

        val = (self.cleaned_data.get("duracao_retorno") or "").strip()
        if not val:
            return ""
        minutos = parse_duracao_minutos(val)
        if minutos is None or minutos < 0:
            self.add_error("duracao_retorno", "Use o formato HH:MM (ex.: 6:30) ou apenas minutos (≥ 0).")
        return val

    def clean(self):
        data = super().clean()
        # Duração >= 0 já garantido por parse_duracao_minutos
        retorno_data = data.get("retorno_saida_data")
        retorno_hora = data.get("retorno_saida_hora")
        ida_data = data.get("saida_data")
        ida_hora = data.get("saida_hora")
        if retorno_data and retorno_hora and ida_data and ida_hora:
            from datetime import datetime

            dt_ida = datetime.combine(ida_data, ida_hora)
            dt_retorno = datetime.combine(retorno_data, retorno_hora)
            if dt_retorno < dt_ida:
                self.add_error(
                    "retorno_saida_data",
                    "A data/hora de saída do retorno não pode ser anterior à ida.",
                )
        return data
