# Assinante padrão para justificativas (separado do assinante dos ofícios)

from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ("viagens", "0048_oficio_justificativa_texto_modelo"),
    ]

    operations = [
        migrations.AddField(
            model_name="oficioconfig",
            name="assinante_justificativa",
            field=models.ForeignKey(
                blank=True,
                null=True,
                on_delete=models.SET_NULL,
                related_name="oficio_config_assinante_justificativa",
                to="viagens.viajante",
            ),
        ),
    ]
