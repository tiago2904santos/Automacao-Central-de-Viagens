# Atualiza choices de Evento.tipo_demanda e normaliza valores antigos para OUTRO

from django.db import migrations, models

VALIDOS = {"PCPR_NA_COMUNIDADE", "OPERACAO_POLICIAL", "PARANA_EM_ACAO", "OUTRO"}
MAPEAR_PARA_OUTRO = {"INTERIOR", "CAPITAL", "BRASILIA"}


def normalizar_tipo_demanda_model(Evento_model):
    """Normaliza tipo_demanda: valores antigos ou inválidos viram OUTRO. Reutilizável em testes."""
    for ev in Evento_model.objects.all():
        val = (ev.tipo_demanda or "").strip().upper()
        if val in MAPEAR_PARA_OUTRO or val not in VALIDOS:
            ev.tipo_demanda = "OUTRO"
            ev.save(update_fields=["tipo_demanda"])


def normalizar_tipo_demanda(apps, schema_editor):
    normalizar_tipo_demanda_model(apps.get_model("viagens", "Evento"))


def noop(apps, schema_editor):
    pass


class Migration(migrations.Migration):

    dependencies = [
        ("viagens", "0052_evento_tipo_demanda"),
    ]

    operations = [
        migrations.AlterField(
            model_name="evento",
            name="tipo_demanda",
            field=models.CharField(
                blank=True,
                choices=[
                    ("PCPR_NA_COMUNIDADE", "PCPR na Comunidade"),
                    ("OPERACAO_POLICIAL", "Operação Policial"),
                    ("PARANA_EM_ACAO", "Paraná em Ação"),
                    ("OUTRO", "Outro"),
                ],
                default="OUTRO",
                max_length=32,
                verbose_name="Tipo de demanda",
            ),
        ),
        migrations.RunPython(normalizar_tipo_demanda, noop),
    ]
