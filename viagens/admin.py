from django.contrib import admin
from django.utils.html import format_html

from .models import (
    Cidade,
    CoordenadorMunicipal,
    Estado,
    ModeloMotivo,
    Oficio,
    OficioRoteiro,
    OficioViajante,
    PlanoTrabalho,
    PlanoTrabalhoAtividade,
    PlanoTrabalhoLocalAtuacao,
    PlanoTrabalhoMeta,
    PlanoTrabalhoRecurso,
    RoteiroDestino,
    Roteiro,
    Trecho,
    TrechoRoteiro,
    Viajante,
    Veiculo,
)


@admin.register(Viajante)
class ViajanteAdmin(admin.ModelAdmin):
    list_display = ("nome", "rg", "cpf", "cargo", "telefone")
    search_fields = ("nome", "rg", "cpf", "cargo", "telefone")


@admin.register(Veiculo)
class VeiculoAdmin(admin.ModelAdmin):
    list_display = ("placa", "modelo", "combustivel")
    search_fields = ("placa", "modelo", "combustivel")


@admin.register(Estado)
class EstadoAdmin(admin.ModelAdmin):
    list_display = ("sigla", "nome")
    search_fields = ("sigla", "nome")


@admin.register(Cidade)
class CidadeAdmin(admin.ModelAdmin):
    list_display = ("nome", "estado")
    search_fields = ("nome", "estado__sigla", "estado__nome")
    list_filter = ("estado",)


class TrechoInline(admin.TabularInline):
    model = Trecho
    extra = 0
    fields = (
        "ordem",
        "origem_estado",
        "origem_cidade",
        "destino_estado",
        "destino_cidade",
        "saida_data",
        "saida_hora",
        "chegada_data",
        "chegada_hora",
    )


class TrechoRoteiroInline(admin.TabularInline):
    model = TrechoRoteiro
    extra = 0
    fields = (
        "ordem",
        "origem_cidade",
        "destino_cidade",
        "saida_data",
        "saida_hora",
        "chegada_data",
        "chegada_hora",
        "tempo_viagem_minutos",
    )
    ordering = ("ordem", "id")
    raw_id_fields = ("origem_estado", "origem_cidade", "destino_estado", "destino_cidade")


@admin.register(Oficio)
class OficioAdmin(admin.ModelAdmin):
    list_display = ("oficio", "protocolo", "roteiro", "destino_label", "created_at")
    list_filter = ("created_at",)
    search_fields = (
        "oficio",
        "protocolo",
        "roteiro__nome",
        "destino",
        "assunto",
        "placa",
        "motorista",
        "cidade_destino__nome",
        "cidade_sede__nome",
    )
    raw_id_fields = ("roteiro",)
    inlines = (TrechoInline,)

    def destino_label(self, obj):
        return obj.get_destino_display()

    destino_label.short_description = "Destino"


@admin.register(Roteiro)
class RoteiroViagemAdmin(admin.ModelAdmin):
    list_display = (
        "nome",
        "cidade_sede",
        "ativo",
        "created_at",
    )
    list_filter = ("ativo",)
    search_fields = ("nome",)
    inlines = (TrechoRoteiroInline,)
    ordering = ("-criado_em", "-id")
    date_hierarchy = "criado_em"
    readonly_fields = ("nome", "criado_em", "atualizado_em")
    fieldsets = (
        (None, {"fields": ("estado_sede", "cidade_sede", "ativo")}),
        (
            "Dados de retorno",
            {
                "fields": (
                    "retorno_saida_cidade",
                    "retorno_saida_data",
                    "retorno_saida_hora",
                    "retorno_chegada_cidade",
                    "retorno_chegada_data",
                    "retorno_chegada_hora",
                )
            }
        ),
        (
            "Informacoes de sistema",
            {
                "fields": ("nome", "criado_em", "atualizado_em"),
                "classes": ("collapse",),
            },
        ),
    )


@admin.register(OficioRoteiro)
class OficioRoteiroAdmin(admin.ModelAdmin):
    list_display = ("oficio_link", "roteiro_link", "vinculado_em", "observacao")
    list_filter = ("vinculado_em",)
    search_fields = (
        "oficio__oficio",
        "roteiro__nome",
        "observacao",
    )
    raw_id_fields = ("oficio", "roteiro")

    def oficio_link(self, obj):
        return format_html('<a href="{}">{}</a>', obj.oficio.get_admin_url(), obj.oficio)

    oficio_link.short_description = "Oficio"

    def roteiro_link(self, obj):
        return format_html('<a href="{}">{}</a>', obj.roteiro.get_admin_url(), obj.roteiro)

    roteiro_link.short_description = "Roteiro"


@admin.register(CoordenadorMunicipal)
class CoordenadorMunicipalAdmin(admin.ModelAdmin):
    list_display = ("nome", "cargo", "cidade", "ativo", "updated_at")
    search_fields = ("nome", "cargo", "cidade")
    list_filter = ("ativo", "cidade")


class PlanoTrabalhoMetaInline(admin.TabularInline):
    model = PlanoTrabalhoMeta
    extra = 0
    fields = ("ordem", "descricao")
    ordering = ("ordem", "id")


class PlanoTrabalhoAtividadeInline(admin.TabularInline):
    model = PlanoTrabalhoAtividade
    extra = 0
    fields = ("ordem", "descricao")
    ordering = ("ordem", "id")


class PlanoTrabalhoRecursoInline(admin.TabularInline):
    model = PlanoTrabalhoRecurso
    extra = 0
    fields = ("ordem", "descricao")
    ordering = ("ordem", "id")


class PlanoTrabalhoLocalInline(admin.TabularInline):
    model = PlanoTrabalhoLocalAtuacao
    extra = 0
    fields = ("ordem", "data", "local")
    ordering = ("ordem", "id")


@admin.register(PlanoTrabalho)
class PlanoTrabalhoAdmin(admin.ModelAdmin):
    list_display = (
        "numero",
        "ano",
        "sigla_unidade",
        "destino",
        "solicitante",
        "coordenador_plano",
        "possui_coordenador_municipal",
        "updated_at",
    )
    search_fields = ("numero", "ano", "destino", "solicitante")
    list_filter = ("ano", "sigla_unidade", "possui_coordenador_municipal")
    autocomplete_fields = ("coordenador_plano", "coordenador_municipal")
    inlines = (
        PlanoTrabalhoMetaInline,
        PlanoTrabalhoAtividadeInline,
        PlanoTrabalhoRecursoInline,
        PlanoTrabalhoLocalInline,
    )


@admin.register(OficioViajante)
class OficioViajanteAdmin(admin.ModelAdmin):
    list_display = ("oficio", "viajante", "ordem", "nome_snapshot", "cargo_snapshot")
    list_filter = ("oficio__ano",)
    search_fields = ("viajante__nome", "oficio__oficio", "nome_snapshot")
    autocomplete_fields = ("viajante",)


@admin.register(RoteiroDestino)
class RoteiroDestinoAdmin(admin.ModelAdmin):
    list_display = ("roteiro", "nome", "estado", "cidade", "principal", "ordem")
    list_filter = ("principal",)
    search_fields = ("nome", "roteiro__nome")
    autocomplete_fields = ("estado", "cidade")


@admin.register(ModeloMotivo)
class ModeloMotivoAdmin(admin.ModelAdmin):
    list_display = ("codigo", "label", "ordem", "ativo", "padrao")
    list_filter = ("ativo", "padrao")
    search_fields = ("codigo", "label")
