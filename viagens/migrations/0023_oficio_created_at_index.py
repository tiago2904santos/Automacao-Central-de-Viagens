from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ("viagens", "0022_oficio_numeracao_por_ano"),
    ]

    operations = [
        migrations.AlterField(
            model_name="oficio",
            name="created_at",
            field=models.DateTimeField(auto_now_add=True, db_index=True),
        ),
    ]
