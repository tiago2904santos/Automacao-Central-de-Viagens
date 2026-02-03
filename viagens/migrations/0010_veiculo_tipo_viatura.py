from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ("viagens", "0009_oficio_draft_status"),
    ]

    operations = [
        migrations.AddField(
            model_name="veiculo",
            name="tipo_viatura",
            field=models.CharField(
                blank=True,
                choices=[("CARACTERIZADA", "Caracterizada"), ("DESCARACTERIZADA", "Descaracterizada")],
                max_length=20,
            ),
        ),
    ]
