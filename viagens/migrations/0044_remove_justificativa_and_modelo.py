# Generated manually - remove justificativa from Oficio and ModeloJustificativa model

from django.db import migrations


class Migration(migrations.Migration):

    dependencies = [
        ('viagens', '0043_modelo_justificativa'),
    ]

    operations = [
        migrations.RemoveField(
            model_name='oficio',
            name='justificativa_texto',
        ),
        migrations.RemoveField(
            model_name='oficio',
            name='justificativa_modelo',
        ),
        migrations.DeleteModel(
            name='ModeloJustificativa',
        ),
    ]
