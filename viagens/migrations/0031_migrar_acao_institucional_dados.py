from django.db import migrations


def migrar_dados_existentes(apps, schema_editor):
    Oficio = apps.get_model("viagens", "Oficio")
    PlanoTrabalho = apps.get_model("viagens", "PlanoTrabalho")
    OrdemServico = apps.get_model("viagens", "OrdemServico")
    AcaoInstitucional = apps.get_model("viagens", "AcaoInstitucional")

    for oficio in Oficio.objects.all().iterator():
        identificador = ""
        if getattr(oficio, "numero", None) and getattr(oficio, "ano", None):
            identificador = f"{oficio.numero}/{oficio.ano}"
        elif getattr(oficio, "oficio", None):
            identificador = str(oficio.oficio).strip()
        else:
            identificador = str(oficio.pk)

        acao = AcaoInstitucional.objects.create(
            titulo=f"Acao - {identificador}",
            descricao=f"Acao institucional derivada do Oficio {identificador}.",
        )

        Oficio.objects.filter(pk=oficio.pk).update(acao=acao)
        PlanoTrabalho.objects.filter(oficio_id=oficio.pk, acao__isnull=True).update(acao=acao)
        OrdemServico.objects.filter(oficio_id=oficio.pk, acao__isnull=True).update(acao=acao)


class Migration(migrations.Migration):

    dependencies = [
        ("viagens", "0030_acaoinstitucional_oficio_acao_ordemservico_acao_and_more"),
    ]

    operations = [
        migrations.RunPython(migrar_dados_existentes, reverse_code=migrations.RunPython.noop),
    ]
