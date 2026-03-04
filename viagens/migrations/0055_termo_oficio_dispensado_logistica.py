# Generated migration: Termo por ofício/viajante, dispensado, logística

from django.db import migrations, models
import django.db.models.deletion


class Migration(migrations.Migration):

    dependencies = [
        ("viagens", "0054_roteiro_evento"),
    ]

    operations = [
        migrations.AddField(
            model_name="termoautorizacao",
            name="oficio",
            field=models.ForeignKey(
                blank=True,
                null=True,
                on_delete=django.db.models.deletion.SET_NULL,
                related_name="termos_autorizacao",
                to="viagens.oficio",
                verbose_name="Ofício",
            ),
        ),
        migrations.AddField(
            model_name="termoautorizacao",
            name="dispensado",
            field=models.BooleanField(
                default=False,
                help_text="Se marcado, não exige geração/assinatura do termo (motivo em dispensa_motivo).",
                verbose_name="Termo dispensado",
            ),
        ),
        migrations.AddField(
            model_name="termoautorizacao",
            name="dispensa_motivo",
            field=models.CharField(blank=True, default="", max_length=200),
        ),
        migrations.AddField(
            model_name="termoautorizacao",
            name="motorista_nome",
            field=models.CharField(blank=True, default="", max_length=120),
        ),
        migrations.AddField(
            model_name="termoautorizacao",
            name="veiculo_modelo",
            field=models.CharField(blank=True, default="", max_length=120),
        ),
        migrations.AddField(
            model_name="termoautorizacao",
            name="veiculo_placa",
            field=models.CharField(blank=True, default="", max_length=20),
        ),
        migrations.AddField(
            model_name="termoautorizacao",
            name="combustivel",
            field=models.CharField(blank=True, default="", max_length=60),
        ),
        migrations.AlterField(
            model_name="termoautorizacao",
            name="evento",
            field=models.ForeignKey(
                blank=True,
                null=True,
                on_delete=django.db.models.deletion.SET_NULL,
                related_name="termos",
                to="viagens.evento",
                verbose_name="Evento",
            ),
        ),
        migrations.AlterField(
            model_name="termoautorizacao",
            name="viajante",
            field=models.ForeignKey(
                blank=True,
                null=True,
                on_delete=django.db.models.deletion.SET_NULL,
                related_name="termos_autorizacao",
                to="viagens.viajante",
                verbose_name="Servidor",
            ),
        ),
        migrations.AddConstraint(
            model_name="termoautorizacao",
            constraint=models.UniqueConstraint(
                condition=models.Q(("dispensado", False)) & ~models.Q(oficio=None) & ~models.Q(viajante=None),
                fields=("oficio", "viajante"),
                name="uniq_termo_oficio_viajante_ativo",
            ),
        ),
    ]
