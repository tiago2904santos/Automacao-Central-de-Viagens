from django.db import migrations, models


class Migration(migrations.Migration):
    dependencies = [
        ("viagens", "0007_oficio_campos_documento"),
    ]

    operations = [
        migrations.AddField(
            model_name="oficio",
            name="tipo_destino",
            field=models.CharField(
                blank=True,
                choices=[
                    ("INTERIOR", "Interior"),
                    ("CAPITAL", "Capital"),
                    ("BRASILIA", "Brasilia"),
                ],
                max_length=20,
            ),
        ),
    ]
