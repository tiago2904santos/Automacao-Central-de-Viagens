from django.db import migrations, models


def _copy_custos_to_custeio_tipo(apps, schema_editor):
    Oficio = apps.get_model("viagens", "Oficio")
    valid = {"UNIDADE", "OUTRA_INSTITUICAO", "ONUS_LIMITADOS"}
    for oficio in Oficio.objects.all().only("id", "custos", "custeio_tipo"):
        current = (getattr(oficio, "custeio_tipo", "") or "").strip()
        if current == "SEM_ONUS":
            oficio.custeio_tipo = "ONUS_LIMITADOS"
            oficio.save(update_fields=["custeio_tipo"])
            continue
        if current:
            continue
        custos_value = (getattr(oficio, "custos", "") or "").strip()
        if custos_value == "SEM_ONUS":
            custos_value = "ONUS_LIMITADOS"
        if custos_value in valid:
            oficio.custeio_tipo = custos_value
            oficio.save(update_fields=["custeio_tipo"])


def _noop(apps, schema_editor):
    pass


class Migration(migrations.Migration):
    dependencies = [
        ("viagens", "0018_oficio_carona_referencia"),
    ]

    operations = [
        migrations.AddField(
            model_name="oficio",
            name="custeio_tipo",
            field=models.CharField(
                blank=True,
                choices=[
                    ("UNIDADE", "UNIDADE - DPC (diaria e combustivel serao custeados pela DPC)."),
                    ("OUTRA_INSTITUICAO", "OUTRA INSTITUICAO"),
                    ("ONUS_LIMITADOS", "ONUS LIMITADOS AOS PROPRIOS VENCIMENTOS"),
                ],
                default="UNIDADE",
                max_length=30,
            ),
        ),
        migrations.AddField(
            model_name="oficio",
            name="custeio_texto_override",
            field=models.TextField(blank=True, default=""),
        ),
        migrations.RunPython(_copy_custos_to_custeio_tipo, _noop),
    ]
