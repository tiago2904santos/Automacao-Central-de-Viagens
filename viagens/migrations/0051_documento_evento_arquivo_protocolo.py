# DocumentoEventoArquivo (upload assinados) e EventoProtocoloArquivo (PDF compilado)

from django.db import migrations, models
import django.db.models.deletion


class Migration(migrations.Migration):

    dependencies = [
        ("viagens", "0050_evento_pacote"),
    ]

    operations = [
        migrations.CreateModel(
            name="DocumentoEventoArquivo",
            fields=[
                ("id", models.BigAutoField(auto_created=True, primary_key=True, serialize=False, verbose_name="ID")),
                (
                    "tipo",
                    models.CharField(
                        choices=[
                            ("OFICIO_ASSINADO", "Ofício assinado"),
                            ("PLANO_ASSINADO", "Plano de trabalho assinado"),
                            ("ORDEM_ASSINADO", "Ordem de serviço assinada"),
                            ("JUSTIFICATIVA_ASSINADA", "Justificativa assinada"),
                            ("TERMO_ASSINADO", "Termo de autorização assinado"),
                        ],
                        db_index=True,
                        max_length=32,
                    ),
                ),
                ("arquivo", models.FileField(max_length=500, upload_to="evento_documentos/%Y/%m/", verbose_name="Arquivo")),
                ("original_name", models.CharField(blank=True, max_length=255)),
                ("mime_type", models.CharField(blank=True, max_length=100)),
                (
                    "is_active",
                    models.BooleanField(
                        default=True,
                        help_text="Apenas o ativo por (evento, tipo, ofício/viajante) conta para checklist.",
                    ),
                ),
                ("uploaded_at", models.DateTimeField(auto_now_add=True)),
                ("uploaded_by_id", models.PositiveIntegerField(blank=True, db_index=True, null=True)),
                (
                    "evento",
                    models.ForeignKey(
                        on_delete=django.db.models.deletion.CASCADE,
                        related_name="arquivos_documentos",
                        to="viagens.evento",
                        verbose_name="Evento",
                    ),
                ),
                (
                    "oficio",
                    models.ForeignKey(
                        blank=True,
                        null=True,
                        on_delete=django.db.models.deletion.CASCADE,
                        related_name="arquivos_evento",
                        to="viagens.oficio",
                        verbose_name="Ofício",
                    ),
                ),
                (
                    "viajante",
                    models.ForeignKey(
                        blank=True,
                        null=True,
                        on_delete=django.db.models.deletion.CASCADE,
                        related_name="arquivos_termo_evento",
                        to="viagens.viajante",
                        verbose_name="Servidor (termo)",
                    ),
                ),
            ],
            options={
                "verbose_name": "Arquivo de documento do evento",
                "verbose_name_plural": "Arquivos de documentos do evento",
                "ordering": ["-uploaded_at"],
            },
        ),
        migrations.CreateModel(
            name="EventoProtocoloArquivo",
            fields=[
                ("id", models.BigAutoField(auto_created=True, primary_key=True, serialize=False, verbose_name="ID")),
                (
                    "pdf_compilado",
                    models.FileField(max_length=500, upload_to="evento_protocolo/%Y/%m/", verbose_name="PDF compilado"),
                ),
                ("compilado_em", models.DateTimeField(auto_now_add=True)),
                ("compilado_por_id", models.PositiveIntegerField(blank=True, db_index=True, null=True)),
                ("hash_sha256", models.CharField(blank=True, max_length=64)),
                ("versao", models.PositiveIntegerField(default=1)),
                (
                    "evento",
                    models.ForeignKey(
                        on_delete=django.db.models.deletion.CASCADE,
                        related_name="protocolos_compilados",
                        to="viagens.evento",
                        verbose_name="Evento",
                    ),
                ),
            ],
            options={
                "verbose_name": "Protocolo compilado do evento",
                "verbose_name_plural": "Protocolos compilados",
                "ordering": ["-compilado_em"],
            },
        ),
        migrations.AddIndex(
            model_name="documentoeventoarquivo",
            index=models.Index(fields=["evento", "tipo", "oficio"], name="deae_ev_tipo_of"),
        ),
        migrations.AddIndex(
            model_name="documentoeventoarquivo",
            index=models.Index(fields=["evento", "tipo", "viajante"], name="deae_ev_tipo_via"),
        ),
    ]
