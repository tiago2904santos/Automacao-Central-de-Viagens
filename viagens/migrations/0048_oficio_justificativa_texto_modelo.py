# Generated manually - campos para justificativa vinculada ao ofício

from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ("viagens", "0047_modelos_justificativa_iniciais"),
    ]

    operations = [
        migrations.AddField(
            model_name="oficio",
            name="justificativa_modelo",
            field=models.CharField(
                blank=True,
                default="",
                help_text="Código do modelo de justificativa usado (ex: recebimento_tardio).",
                max_length=80,
            ),
        ),
        migrations.AddField(
            model_name="oficio",
            name="justificativa_texto",
            field=models.TextField(
                blank=True,
                default="",
                help_text="Texto da justificativa preenchido (prazo < 10 dias). Preenchido desbloqueia geração do ofício.",
            ),
        ),
    ]
