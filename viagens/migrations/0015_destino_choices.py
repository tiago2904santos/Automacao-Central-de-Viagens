from django.db import migrations, models


def _normalize_destino(apps, schema_editor):
    Oficio = apps.get_model("viagens", "Oficio")
    for oficio in Oficio.objects.all():
        raw = (oficio.destino or "").strip().upper()
        if "SESP" in raw:
            code = "SESP"
        elif "GABINETE" in raw or "DELEGADO GERAL ADJUNTO" in raw:
            code = "GAB"
        else:
            code = "GAB"
        if oficio.destino != code:
            oficio.destino = code
            oficio.save(update_fields=["destino"])


class Migration(migrations.Migration):

    dependencies = [
        ("viagens", "0014_assunto_cargo_viatura_default"),
    ]

    operations = [
        migrations.AlterField(
            model_name="oficio",
            name="destino",
            field=models.CharField(
                choices=[
                    ("GAB", "GABINETE DO DELEGADO GERAL ADJUNTO"),
                    ("SESP", "SESP"),
                ],
                default="GAB",
                max_length=40,
            ),
        ),
        migrations.RunPython(_normalize_destino, migrations.RunPython.noop),
    ]
