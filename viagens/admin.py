from django.contrib import admin

from .models import (
    Cidade,
    CoordenadorMunicipal,
    Estado,
    Oficio,
    PlanoTrabalho,
    PlanoTrabalhoAtividade,
    PlanoTrabalhoLocalAtuacao,
    PlanoTrabalhoMeta,
    PlanoTrabalhoRecurso,
    Trecho,
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


@admin.register(Oficio)
class OficioAdmin(admin.ModelAdmin):
    list_display = ("oficio", "protocolo", "destino_label", "created_at")
    list_filter = ("created_at",)
    search_fields = (
        "oficio",
        "protocolo",
        "destino",
        "assunto",
        "placa",
        "motorista",
        "cidade_destino__nome",
        "cidade_sede__nome",
    )
    inlines = (TrechoInline,)

    def destino_label(self, obj):
        return obj.get_destino_display()

    destino_label.short_description = "Destino"


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
