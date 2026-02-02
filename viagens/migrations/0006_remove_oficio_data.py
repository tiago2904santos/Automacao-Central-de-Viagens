from django.db import migrations


class Migration(migrations.Migration):
    dependencies = [
        ("viagens", "0005_oficio_motorista_carona_oficio_motorista_oficio_and_more"),
    ]

    operations = [
        migrations.RemoveField(
            model_name="oficio",
            name="data",
        ),
    ]
