# Generated manually - Dados iniciais dos modelos de justificativa

from django.db import migrations


TEXTOS = [
    (
        "recebimento_tardio",
        "Demanda recebida tardiamente",
        1,
        """Em atenção ao prazo de 10 dias estabelecido pelo Decreto nº 6.358/2024 e ao Ofício Circular nº 340/2024-GAF, informamos que o pedido de deslocamento referente ao Ofício nº X/ANO foi encaminhado na data em que a demanda foi formalmente recebida por esta unidade, razão pela qual o presente ofício está sendo enviado com prazo inferior ao estipulado. Esclarecemos que o envio ocorreu imediatamente após o recebimento da solicitação, não havendo possibilidade de cumprimento integral do prazo regulamentar, considerando a data e o horário em que o pedido foi repassado para providências.""",
    ),
    (
        "operacao_policial",
        "Operação policial (aguardar orientações DG)",
        2,
        """Em atenção ao prazo de 10 dias estabelecido pelo Decreto nº 6.358/2024, e Ofício Circular 340/2024-GAF, justificamos que o envio se deu em data próxima ao deslocamento em razão da necessidade de aguardar as orientações do Gabinete do Delegado-Geral acerca da operação policial, imprescindíveis para a definição das diretrizes e correta formalização da demanda.""",
    ),
    (
        "evento",
        "Confirmação tardia de data/local do evento",
        3,
        """O prazo de 10 (dez) dias previsto no Decreto nº 6.358/2024 e no Ofício Circular nº 340/2024-GAF não pôde ser observado, uma vez que a equipe encontrava-se em tratativas para a confirmação da data e definição do local do evento, circunstância que inviabilizou o protocolo antecipado da solicitação de diárias e a adoção das demais providências administrativas pertinentes. Informa-se, ainda, que as servidoras realizarão o deslocamento no dia anterior ao evento, com a finalidade de executar visita técnica e promover os alinhamentos necessários à sua realização. Tais medidas mostram-se indispensáveis para a adequada organização da solenidade, permitindo o reconhecimento prévio do local, a verificação da infraestrutura disponível e o alinhamento das demandas operacionais, a fim de assegurar a qualidade do cerimonial e o regular andamento das atividades planejadas.""",
    ),
    (
        "servidores",
        "Documentos/autorização de servidores de outras unidades",
        4,
        """Em atenção ao prazo de 10 dias estabelecido pelo Decreto nº 6.358/2024, e Ofício Circular 340/2024-GAF, o deslocamento referente ao Ofício X/ANO, ocorreu de forma intempestiva. Justificamos que não foi possível encaminhar o ofício com a devida antecedência, uma vez que contamos com a participação de servidores de outras unidades que compõem a equipe de apoio a esta Assessoria de Comunicação, e essa readequação demanda a espera para recebimento das autorizações das chefias, bem como da documentação desses servidores, que nem sempre ocorrem em tempo hábil.""",
    ),
]


def criar_modelos(apps, schema_editor):
    ModeloJustificativa = apps.get_model("viagens", "ModeloJustificativa")
    for i, (codigo, label, ordem, texto) in enumerate(TEXTOS):
        ModeloJustificativa.objects.get_or_create(
            codigo=codigo,
            defaults={"label": label, "texto": texto.strip(), "ordem": ordem, "ativo": True, "padrao": i == 0},
        )


def reverter(apps, schema_editor):
    ModeloJustificativa = apps.get_model("viagens", "ModeloJustificativa")
    ModeloJustificativa.objects.filter(codigo__in=[t[0] for t in TEXTOS]).delete()


class Migration(migrations.Migration):

    dependencies = [
        ("viagens", "0046_modelojustificativa"),
    ]

    operations = [
        migrations.RunPython(criar_modelos, reverter),
    ]
