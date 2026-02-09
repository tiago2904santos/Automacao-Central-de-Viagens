from django.core.exceptions import ValidationError
from django.db import models


class Viajante(models.Model):
    nome = models.CharField(max_length=200)
    rg = models.CharField(max_length=50)
    cpf = models.CharField(max_length=50)
    cargo = models.CharField(max_length=120)
    telefone = models.CharField(max_length=30, blank=True)

    def __str__(self) -> str:
        return self.nome


class Cargo(models.Model):
    nome = models.CharField(max_length=120, unique=True)

    def __str__(self) -> str:
        return self.nome


class Veiculo(models.Model):
    placa = models.CharField(max_length=10, unique=True)
    modelo = models.CharField(max_length=120)
    combustivel = models.CharField(max_length=80)
    tipo_viatura = models.CharField(
        max_length=20,
        blank=True,
        choices=[("CARACTERIZADA", "Caracterizada"), ("DESCARACTERIZADA", "Descaracterizada")],
        default="DESCARACTERIZADA"
    )

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
        default="",
    )
    cargo_chefia = models.CharField(
        max_length=120,
        default="",
    )
    orgao_origem = models.CharField(
        max_length=200,
        default="ASSESSORIA DE COMUNICAÇÃO SOCIAL",
    )
    orgao_destino_padrao = models.CharField(
        max_length=200,
        default="GABINETE DO DELEGADO GERAL ADJUNTO",
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


class OficioConfig(models.Model):
    unidade_nome = models.CharField(max_length=255, default="")
    origem_nome = models.CharField(max_length=255, default="")

    cep = models.CharField(max_length=9, default="")
    logradouro = models.CharField(max_length=255, blank=True, default="")
    bairro = models.CharField(max_length=255, blank=True, default="")
    cidade = models.CharField(max_length=255, blank=True, default="")
    uf = models.CharField(max_length=2, blank=True, default="")
    numero = models.CharField(max_length=30, default="")
    complemento = models.CharField(max_length=120, blank=True, default="")
    telefone = models.CharField(max_length=50, blank=True, default="")
    email = models.EmailField(blank=True, default="")

    assinante = models.ForeignKey(
        Viajante,
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name="oficio_configs",
    )
    sede_cidade_default = models.ForeignKey(
        Cidade,
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name="oficio_config_sede_default",
    )

    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        verbose_name = "Configuracao do Oficio"
        verbose_name_plural = "Configuracoes do Oficio"

    def __str__(self) -> str:
        return "Configuracao do Oficio"

    def save(self, *args, **kwargs):
        if self.unidade_nome:
            self.unidade_nome = self.unidade_nome.upper()
        if self.origem_nome:
            self.origem_nome = self.origem_nome.upper()
        self.pk = 1
        super().save(*args, **kwargs)


class Oficio(models.Model):
    class Status(models.TextChoices):
        DRAFT = "DRAFT", "Rascunho"
        FINAL = "FINAL", "Finalizado"

    class AssuntoTipo(models.TextChoices):
        AUTORIZACAO = "AUTORIZACAO", "Autorizacao"
        CONVALIDACAO = "CONVALIDACAO", "Convalidacao"

    class CustosChoices(models.TextChoices):
        UNIDADE = "UNIDADE", "UNIDADE – DPC (diária e combustível serão custeados pela DPC)."
        OUTRA_INSTITUICAO = "OUTRA_INSTITUICAO", "OUTRA INSTITUIÇÃO"
        SEM_ONUS = "SEM_ONUS", "Com ônus limitados aos próprios vencimentos"

    class CusteioTipoChoices(models.TextChoices):
        UNIDADE = "UNIDADE", "UNIDADE - DPC (diarias e combustivel serao custeados pela DPC)."
        OUTRA_INSTITUICAO = "OUTRA_INSTITUICAO", "OUTRA INSTITUICAO"
        ONUS_LIMITADOS = "ONUS_LIMITADOS", "ONUS LIMITADOS AOS PROPRIOS VENCIMENTOS"


    class DestinoChoices(models.TextChoices):
        GAB = "GAB", "GABINETE DO DELEGADO GERAL ADJUNTO"
        SESP = "SESP", "SESP"

    oficio = models.CharField(max_length=50, blank=True, default="")
    protocolo = models.CharField(max_length=80, blank=True, default="")
    destino = models.CharField(
        max_length=40,
        choices=DestinoChoices.choices,
        default=DestinoChoices.GAB,
        editable=False,
    )
    status = models.CharField(
        max_length=20,
        choices=Status.choices,
        default=Status.DRAFT,
        db_index=True,
    )
    assunto = models.CharField(max_length=200, blank=True)
    assunto_tipo = models.CharField(
        max_length=20,
        choices=AssuntoTipo.choices,
        default=AssuntoTipo.AUTORIZACAO,
    )
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
        default="DESCARACTERIZADA"
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
    custeio_tipo = models.CharField(
        max_length=30,
        blank=True,
        choices=CusteioTipoChoices.choices,
        default=CusteioTipoChoices.UNIDADE,
    )
    custeio_texto_override = models.TextField(blank=True, default="")
    custos = models.CharField(
        max_length=20,
        choices=CustosChoices.choices,
        default=CustosChoices.UNIDADE,
    )
    nome_instituicao_custeio = models.CharField(
        max_length=200,
        blank=True,
        default="",
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
    carona_oficio_referencia = models.ForeignKey(
        "self",
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name="oficios_que_usam_como_carona",
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
        destino = self.cidade_destino or self.get_destino_display() or self.destino
        return f"Oficio {self.oficio} - {destino}"

    def calcular_destino_automatico(self) -> str:
        if not self.pk:
            return self.DestinoChoices.GAB
        trechos = self.trechos.select_related("destino_estado", "destino_cidade__estado")
        for trecho in trechos:
            estado = trecho.destino_estado or (
                trecho.destino_cidade.estado if trecho.destino_cidade else None
            )
            if estado and (estado.sigla or "").strip().upper() != "PR":
                return self.DestinoChoices.SESP
        return self.DestinoChoices.GAB

    def clean(self) -> None:
        super().clean()
        custeio_tipo = (self.custeio_tipo or self.custos or "").strip()
        if custeio_tipo == "SEM_ONUS":
            custeio_tipo = "ONUS_LIMITADOS"
        if (
            custeio_tipo == self.CusteioTipoChoices.OUTRA_INSTITUICAO
            and not (self.nome_instituicao_custeio or "").strip()
        ):
            raise ValidationError(
                {"nome_instituicao_custeio": "Informe a instituicao de custeio."}
            )

    def save(self, *args, **kwargs):
        if not (self.custeio_tipo or "").strip():
            custos_value = (self.custos or "").strip()
            if custos_value == "SEM_ONUS":
                custos_value = "ONUS_LIMITADOS"
            if custos_value:
                self.custeio_tipo = custos_value
        if self.custeio_tipo == "SEM_ONUS":
            self.custeio_tipo = "ONUS_LIMITADOS"
        if self.custeio_tipo != self.CusteioTipoChoices.OUTRA_INSTITUICAO:
            if (self.nome_instituicao_custeio or "").strip():
                self.nome_instituicao_custeio = ""
        self.destino = self.calcular_destino_automatico()
        update_fields = kwargs.get("update_fields")
        if update_fields is not None:
            update_fields = set(update_fields)
            update_fields.add("destino")
            if "nome_instituicao_custeio" not in update_fields and not (
                self.nome_instituicao_custeio or ""
            ):
                update_fields.add("nome_instituicao_custeio")
            kwargs["update_fields"] = list(update_fields)
        super().save(*args, **kwargs)


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
