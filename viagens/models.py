from django.db import models


class Viajante(models.Model):
    nome = models.CharField(max_length=200)
    rg = models.CharField(max_length=50)
    cpf = models.CharField(max_length=50)
    cargo = models.CharField(max_length=120)
    telefone = models.CharField(max_length=30, blank=True)

    def __str__(self) -> str:
        return self.nome


class Veiculo(models.Model):
    placa = models.CharField(max_length=10, unique=True)
    modelo = models.CharField(max_length=120)
    combustivel = models.CharField(max_length=80)

    def __str__(self) -> str:
        return f"{self.placa} - {self.modelo}"


class Estado(models.Model):
    sigla = models.CharField(max_length=2, unique=True)
    nome = models.CharField(max_length=100)

    def __str__(self) -> str:
        return f"{self.nome} ({self.sigla})"


class Cidade(models.Model):
    nome = models.CharField(max_length=120)
    estado = models.ForeignKey(Estado, on_delete=models.CASCADE, related_name="cidades")

    def __str__(self) -> str:
        return f"{self.nome}/{self.estado.sigla}"


class ConfiguracaoOficio(models.Model):
    nome_chefia = models.CharField(
        max_length=120,
        default="Delegado Geral Adjunto",
    )
    cargo_chefia = models.CharField(
        max_length=120,
        default="Gabinete do Delegado Geral Adjunto",
    )
    orgao_origem = models.CharField(
        max_length=200,
        default="Assessoria de Comunicação Social",
    )
    orgao_destino_padrao = models.CharField(
        max_length=200,
        default="Gabinete do Delegado Geral Adjunto",
    )
    rodape = models.TextField(blank=True, default="")
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        verbose_name = "Configuração do Ofício"
        verbose_name_plural = "Configurações do Ofício"

    def __str__(self) -> str:
        return "Configuração do Ofício"

    def save(self, *args, **kwargs):
        self.pk = 1
        super().save(*args, **kwargs)

    @classmethod
    def _default_values(cls) -> dict[str, str]:
        return {
            "nome_chefia": "Delegado Geral Adjunto",
            "cargo_chefia": "Gabinete do Delegado Geral Adjunto",
            "orgao_origem": "Assessoria de Comunicação Social",
            "orgao_destino_padrao": "Gabinete do Delegado Geral Adjunto",
        }

    @classmethod
    def get_solo(cls):
        obj, _ = cls.objects.get_or_create(pk=1, defaults=cls._default_values())
        updated = False
        for field, value in cls._default_values().items():
            if not getattr(obj, field):
                setattr(obj, field, value)
                updated = True
        if updated:
            obj.save()
        return obj


class Oficio(models.Model):
    class Status(models.TextChoices):
        DRAFT = "DRAFT", "Rascunho"
        FINAL = "FINAL", "Finalizado"

    oficio = models.CharField(max_length=50, blank=True, default="")
    protocolo = models.CharField(max_length=80, blank=True, default="")
    destino = models.CharField(max_length=200, blank=True, default="")
    status = models.CharField(
        max_length=20,
        choices=Status.choices,
        default=Status.DRAFT,
        db_index=True,
    )
    assunto = models.CharField(max_length=200, blank=True)
    tipo_destino = models.CharField(
        max_length=20,
        blank=True,
        choices=[
            ("INTERIOR", "Interior"),
            ("CAPITAL", "Capital"),
            ("BRASILIA", "Brasilia"),
        ],
    )
    estado_sede = models.ForeignKey(
        Estado,
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name="oficios_sede",
    )
    cidade_sede = models.ForeignKey(
        Cidade,
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name="oficios_sede",
    )
    estado_destino = models.ForeignKey(
        Estado,
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name="oficios_destino",
    )
    cidade_destino = models.ForeignKey(
        Cidade,
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name="oficios_destino",
    )
    roteiro_ida_saida_local = models.CharField(max_length=200, blank=True)
    roteiro_ida_saida_datahora = models.CharField(max_length=200, blank=True)
    roteiro_ida_chegada_local = models.CharField(max_length=200, blank=True)
    roteiro_ida_chegada_datahora = models.CharField(max_length=200, blank=True)
    roteiro_volta_saida_local = models.CharField(max_length=200, blank=True)
    roteiro_volta_saida_datahora = models.CharField(max_length=200, blank=True)
    roteiro_volta_chegada_local = models.CharField(max_length=200, blank=True)
    roteiro_volta_chegada_datahora = models.CharField(max_length=200, blank=True)
    retorno_saida_cidade = models.CharField(max_length=120, blank=True)
    retorno_saida_data = models.DateField(null=True, blank=True)
    retorno_saida_hora = models.TimeField(null=True, blank=True)
    retorno_chegada_cidade = models.CharField(max_length=120, blank=True)
    retorno_chegada_data = models.DateField(null=True, blank=True)
    retorno_chegada_hora = models.TimeField(null=True, blank=True)
    quantidade_diarias = models.CharField(max_length=120, blank=True)
    valor_diarias = models.CharField(max_length=120, blank=True)
    valor_diarias_extenso = models.CharField(max_length=200, blank=True)
    tipo_viatura = models.CharField(
        max_length=20,
        blank=True,
        choices=[("CARACTERIZADA", "Caracterizada"), ("DESCARACTERIZADA", "Descaracterizada")],
    )
    tipo_custeio = models.CharField(
        max_length=30,
        blank=True,
        choices=[
            ("UNIDADE", "Unidade"),
            ("OUTRA_INSTITUICAO", "Outra instituicao"),
            ("SEM_ONUS", "Sem onus"),
        ],
    )
    google_doc_id = models.CharField(max_length=200, blank=True)
    google_doc_url = models.URLField(blank=True)
    pdf_file_id = models.CharField(max_length=200, blank=True)
    pdf_url = models.URLField(blank=True)
    placa = models.CharField(max_length=10, blank=True)
    modelo = models.CharField(max_length=120, blank=True)
    combustivel = models.CharField(max_length=80, blank=True)
    motorista = models.CharField(max_length=120, blank=True)
    motorista_oficio = models.CharField(max_length=80, blank=True)
    motorista_protocolo = models.CharField(max_length=80, blank=True)
    motorista_carona = models.BooleanField(default=False)
    motorista_viajante = models.ForeignKey(
        Viajante,
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name="oficios_motorista",
    )
    motivo = models.TextField(blank=True)
    veiculo = models.ForeignKey(
        Veiculo,
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name="oficios",
    )
    viajantes = models.ManyToManyField(Viajante, related_name="oficios", blank=True)
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    @property
    def is_draft(self) -> bool:
        return self.status == self.Status.DRAFT

    def __str__(self) -> str:
        destino = self.cidade_destino or self.destino
        return f"Oficio {self.oficio} - {destino}"


class Trecho(models.Model):
    oficio = models.ForeignKey(
        Oficio,
        on_delete=models.CASCADE,
        related_name="trechos",
    )
    ordem = models.PositiveIntegerField(default=1)
    origem_estado = models.ForeignKey(
        Estado,
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name="trechos_origem",
    )
    origem_cidade = models.ForeignKey(
        Cidade,
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name="trechos_origem",
    )
    destino_estado = models.ForeignKey(
        Estado,
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name="trechos_destino",
    )
    destino_cidade = models.ForeignKey(
        Cidade,
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name="trechos_destino",
    )
    saida_data = models.DateField(null=True, blank=True)
    saida_hora = models.TimeField(null=True, blank=True)
    chegada_data = models.DateField(null=True, blank=True)
    chegada_hora = models.TimeField(null=True, blank=True)

    class Meta:
        ordering = ["ordem"]

    def __str__(self) -> str:
        origem = self.origem_cidade or self.origem_estado
        destino = self.destino_cidade or self.destino_estado
        return f"Trecho {self.ordem}: {origem} -> {destino}"
