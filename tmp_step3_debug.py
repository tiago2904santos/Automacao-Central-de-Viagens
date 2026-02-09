import os
import django
from django.test import Client
from django.urls import reverse

os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'central_viagens.settings')
django.setup()

from viagens.models import Cidade, Estado, Viajante

client = Client(HTTP_HOST="localhost")
source_estado, _ = Estado.objects.get_or_create(sigla="PR", defaults={"nome": "Parana"})
source_cidade_sede, _ = Cidade.objects.get_or_create(nome="Curitiba", estado=source_estado)
source_cidade_inter, _ = Cidade.objects.get_or_create(nome="Maringa", estado=source_estado)
viajante, _ = Viajante.objects.get_or_create(nome="Servidor", rg="123", cpf="000", cargo="Delegado")

session = client.session
session["oficio_wizard"] = {
    "oficio": "123/2024",
    "protocolo": "456/2024",
    "placa": "ABC1234",
    "modelo": "Uno",
    "combustivel": "Gasolina",
    "viajantes_ids": [viajante.id],
}
session.save()

payload = {
    "trechos-TOTAL_FORMS": "2",
    "trechos-INITIAL_FORMS": "0",
    "trechos-MIN_NUM_FORMS": "0",
    "trechos-MAX_NUM_FORMS": "1000",
    "trechos-0-origem_estado": source_estado.sigla,
    "trechos-0-origem_cidade": str(source_cidade_sede.id),
    "trechos-0-destino_estado": source_estado.sigla,
    "trechos-0-destino_cidade": str(source_cidade_inter.id),
    "trechos-0-saida_data": "2024-01-01",
    "trechos-0-saida_hora": "07:00",
    "trechos-1-origem_estado": source_estado.sigla,
    "trechos-1-origem_cidade": str(source_cidade_inter.id),
    "retorno_saida_data": "2024-01-02",
    "retorno_saida_hora": "08:00",
    "retorno_chegada_data": "2024-01-02",
    "retorno_chegada_hora": "18:00",
    "tipo_destino": "INTERIOR",
    "motivo": "Retorno",
}

response = client.post(reverse("oficio_step3"), payload)
print(response.status_code)
print(response.content.decode('utf-8'))
