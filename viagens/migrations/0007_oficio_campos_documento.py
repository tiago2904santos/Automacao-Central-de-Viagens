from django.db import migrations, models


class Migration(migrations.Migration):
    dependencies = [
        ("viagens", "0006_remove_oficio_data"),
    ]

    operations = [
        migrations.AddField(
            model_name="oficio",
            name="quantidade_diarias",
            field=models.CharField(blank=True, max_length=120),
        ),
        migrations.AddField(
            model_name="oficio",
            name="valor_diarias",
            field=models.CharField(blank=True, max_length=120),
        ),
        migrations.AddField(
            model_name="oficio",
            name="valor_diarias_extenso",
            field=models.CharField(blank=True, max_length=200),
        ),
        migrations.AddField(
            model_name="oficio",
            name="tipo_viatura",
            field=models.CharField(
                blank=True,
                choices=[("CARACTERIZADA", "Caracterizada"), ("DESCARACTERIZADA", "Descaracterizada")],
                max_length=20,
            ),
        ),
        migrations.AddField(
            model_name="oficio",
            name="tipo_custeio",
            field=models.CharField(
                blank=True,
                choices=[
                    ("UNIDADE", "Unidade"),
                    ("OUTRA_INSTITUICAO", "Outra instituicao"),
                    ("SEM_ONUS", "Sem onus"),
                ],
                max_length=30,
            ),
        ),
        migrations.AddField(
            model_name="oficio",
            name="retorno_saida_cidade",
            field=models.CharField(blank=True, max_length=120),
        ),
        migrations.AddField(
            model_name="oficio",
            name="retorno_saida_data",
            field=models.DateField(blank=True, null=True),
        ),
        migrations.AddField(
            model_name="oficio",
            name="retorno_saida_hora",
            field=models.TimeField(blank=True, null=True),
        ),
        migrations.AddField(
            model_name="oficio",
            name="retorno_chegada_cidade",
            field=models.CharField(blank=True, max_length=120),
        ),
        migrations.AddField(
            model_name="oficio",
            name="retorno_chegada_data",
            field=models.DateField(blank=True, null=True),
        ),
        migrations.AddField(
            model_name="oficio",
            name="retorno_chegada_hora",
            field=models.TimeField(blank=True, null=True),
        ),
        migrations.AddField(
            model_name="oficio",
            name="google_doc_id",
            field=models.CharField(blank=True, max_length=200),
        ),
        migrations.AddField(
            model_name="oficio",
            name="google_doc_url",
            field=models.URLField(blank=True),
        ),
        migrations.AddField(
            model_name="oficio",
            name="pdf_file_id",
            field=models.CharField(blank=True, max_length=200),
        ),
        migrations.AddField(
            model_name="oficio",
            name="pdf_url",
            field=models.URLField(blank=True),
        ),
    ]
