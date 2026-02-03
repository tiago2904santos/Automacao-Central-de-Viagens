from django.db import migrations, models
import django.utils.timezone


def set_existing_final(apps, schema_editor):
    Oficio = apps.get_model("viagens", "Oficio")
    Oficio.objects.update(status="FINAL")


class Migration(migrations.Migration):

    dependencies = [
        ("viagens", "0008_oficio_tipo_destino"),
    ]

    operations = [
        migrations.AlterField(
            model_name="oficio",
            name="oficio",
            field=models.CharField(blank=True, default="", max_length=50),
        ),
        migrations.AlterField(
            model_name="oficio",
            name="protocolo",
            field=models.CharField(blank=True, default="", max_length=80),
        ),
        migrations.AlterField(
            model_name="oficio",
            name="destino",
            field=models.CharField(blank=True, default="", max_length=200),
        ),
        migrations.AddField(
            model_name="oficio",
            name="status",
            field=models.CharField(
                choices=[("DRAFT", "Rascunho"), ("FINAL", "Finalizado")],
                db_index=True,
                default="DRAFT",
                max_length=20,
            ),
        ),
        migrations.AddField(
            model_name="oficio",
            name="updated_at",
            field=models.DateTimeField(default=django.utils.timezone.now, auto_now=True),
            preserve_default=False,
        ),
        migrations.RunPython(set_existing_final, migrations.RunPython.noop),
    ]
