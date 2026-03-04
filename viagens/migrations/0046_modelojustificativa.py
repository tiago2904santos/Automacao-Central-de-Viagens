# Generated manually - Modelos de justificativa pré-prontos

from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ("viagens", "0044_remove_justificativa_and_modelo"),
    ]

    operations = [
        migrations.CreateModel(
            name="ModeloJustificativa",
            fields=[
                ("id", models.BigAutoField(auto_created=True, primary_key=True, serialize=False, verbose_name="ID")),
                ("codigo", models.CharField(help_text="Ex: recebimento_tardio", max_length=80, unique=True)),
                ("label", models.CharField(max_length=200, verbose_name="Nome do modelo")),
                ("texto", models.TextField(verbose_name="Texto da justificativa")),
                ("ordem", models.PositiveIntegerField(default=0)),
                ("ativo", models.BooleanField(default=True)),
                ("padrao", models.BooleanField(default=False, help_text="Se marcado, este modelo será o padrão no gerador.")),
                ("created_at", models.DateTimeField(auto_now_add=True)),
                ("updated_at", models.DateTimeField(auto_now=True)),
            ],
            options={
                "verbose_name": "Modelo de justificativa",
                "verbose_name_plural": "Modelos de justificativa",
                "ordering": ("ordem", "label"),
            },
        ),
    ]
