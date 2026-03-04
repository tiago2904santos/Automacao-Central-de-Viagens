# Evento.tipo_demanda para fluxo guiado

from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ("viagens", "0051_documento_evento_arquivo_protocolo"),
    ]

    operations = [
        migrations.AddField(
            model_name="evento",
            name="tipo_demanda",
            field=models.CharField(
                blank=True,
                choices=[
                    ("OUTRO", "Outro"),
                    ("INTERIOR", "Interior"),
                    ("CAPITAL", "Capital"),
                    ("BRASILIA", "Brasília"),
                ],
                default="OUTRO",
                max_length=20,
                verbose_name="Tipo de demanda",
            ),
        ),
    ]
