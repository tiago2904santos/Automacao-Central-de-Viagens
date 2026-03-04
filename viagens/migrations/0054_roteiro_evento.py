# Roteiro.evento FK para fluxo guiado (Etapa 2)

from django.db import migrations, models
import django.db.models.deletion


class Migration(migrations.Migration):

    dependencies = [
        ("viagens", "0053_evento_tipo_demanda_choices_data"),
    ]

    operations = [
        migrations.AddField(
            model_name="roteiro",
            name="evento",
            field=models.ForeignKey(
                blank=True,
                null=True,
                on_delete=django.db.models.deletion.SET_NULL,
                related_name="roteiros",
                to="viagens.evento",
                verbose_name="Evento",
            ),
        ),
    ]
