# Gerenciador por Evento (pacote): Evento, Oficio.evento, Viajante.is_ascom, Termo.evento/viajante

from django.db import migrations, models
import django.db.models.deletion


class Migration(migrations.Migration):

    dependencies = [
        ("viagens", "0049_oficio_config_assinante_justificativa"),
    ]

    operations = [
        migrations.CreateModel(
            name="Evento",
            fields=[
                ("id", models.BigAutoField(auto_created=True, primary_key=True, serialize=False, verbose_name="ID")),
                ("titulo", models.CharField(max_length=255, verbose_name="Título / Nome do evento")),
                ("data_inicio", models.DateField(blank=True, null=True, verbose_name="Data início")),
                ("data_fim", models.DateField(blank=True, null=True, verbose_name="Data fim")),
                (
                    "tem_convite_ou_oficio_evento",
                    models.BooleanField(
                        default=True,
                        help_text="Se não, o evento exige Plano de Trabalho ou Ordem de Serviço (1 por evento).",
                        verbose_name="Tem ofício solicitando ou convite do evento?",
                    ),
                ),
                ("created_at", models.DateTimeField(auto_now_add=True)),
                ("updated_at", models.DateTimeField(auto_now=True)),
                (
                    "cidade_base",
                    models.ForeignKey(
                        blank=True,
                        null=True,
                        on_delete=django.db.models.deletion.SET_NULL,
                        related_name="eventos_base",
                        to="viagens.cidade",
                        verbose_name="Cidade base",
                    ),
                ),
            ],
            options={
                "verbose_name": "Evento",
                "verbose_name_plural": "Eventos",
                "ordering": ["-created_at"],
            },
        ),
        migrations.AddField(
            model_name="viajante",
            name="is_ascom",
            field=models.BooleanField(
                default=True,
                help_text="Se sim, não exige Termo de Autorização no pacote do evento.",
                verbose_name="É da ASCOM?",
            ),
        ),
        migrations.AddField(
            model_name="oficio",
            name="evento",
            field=models.ForeignKey(
                blank=True,
                null=True,
                on_delete=django.db.models.deletion.SET_NULL,
                related_name="oficios",
                to="viagens.evento",
                verbose_name="Evento (pacote)",
            ),
        ),
        migrations.AddField(
            model_name="termoautorizacao",
            name="evento",
            field=models.ForeignKey(
                blank=True,
                null=True,
                on_delete=django.db.models.deletion.CASCADE,
                related_name="termos",
                to="viagens.evento",
                verbose_name="Evento",
            ),
        ),
        migrations.AddField(
            model_name="termoautorizacao",
            name="viajante",
            field=models.ForeignKey(
                blank=True,
                null=True,
                on_delete=django.db.models.deletion.CASCADE,
                related_name="termos_autorizacao",
                to="viagens.viajante",
                verbose_name="Servidor",
            ),
        ),
    ]
