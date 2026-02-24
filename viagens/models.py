from django.core.exceptions import ValidationError
from django.db import models, transaction
from django.utils import timezone

from .utils.normalize import (
    format_cpf,
    format_oficio_num,
    format_phone,
    format_protocolo_num,
    format_rg,
    normalize_digits,
    normalize_oficio_num,
    normalize_protocolo_num,
    normalize_rg,
    normalize_upper_text,
    split_oficio_num,
)


class Viajante(models.Model):
    nome = models.CharField(max_length=200)
    rg = models.CharField(max_length=50)
    cpf = models.CharField(max_length=50)
    cargo = models.CharField(max_length=120)
    telefone = models.CharField(max_length=30, blank=True)

    def __str__(self) -> str:
        return self.nome

    @property
    def cpf_formatado(self) -> str:
        return format_cpf(self.cpf)

    @property
    def telefone_formatado(self) -> str:
        return format_phone(self.telefone)

    @property
    def rg_formatado(self) -> str:
        return format_rg(self.rg)

    def clean(self) -> None:
        super().clean()
        self.nome = normalize_upper_text(self.nome)
        self.rg = normalize_rg(self.rg)
        self.cpf = normalize_digits(self.cpf)
        self.telefone = normalize_digits(self.telefone)

        errors: dict[str, str] = {}
        if self.rg and len(self.rg) not in {9, 10}:
            errors["rg"] = "RG deve conter 9 ou 10 caracteres (digitos + DV)."
        if self.cpf and len(self.cpf) != 11:
            errors["cpf"] = "CPF deve conter 11 digitos."
        if self.telefone and len(self.telefone) not in {10, 11}:
            errors["telefone"] = "Telefone deve conter 10 ou 11 digitos."
        if errors:
            raise ValidationError(errors)

    def save(self, *args, **kwargs):
        self.nome = normalize_upper_text(self.nome)
        self.rg = normalize_rg(self.rg)
        self.cpf = normalize_digits(self.cpf)
        self.telefone = normalize_digits(self.telefone)
        self.full_clean()
        super().save(*args, **kwargs)


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


class OficioCounter(models.Model):
    ano = models.IntegerField(unique=True)
    last_numero = models.IntegerField(default=0)
    updated_at = models.DateTimeField(auto_now=True)

    def __str__(self) -> str:
        return f"{self.ano}: {self.last_numero}"


class TermoAutorizacao(models.Model):
    data_inicio = models.DateField()
    data_fim = models.DateField(null=True, blank=True)
    data_unica = models.BooleanField(default=False)
    destinos = models.JSONField(default=list, blank=True)
    created_at = models.DateTimeField(auto_now_add=True, db_index=True)
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        verbose_name = "Termo de autorizacao"
        verbose_name_plural = "Termos de autorizacao"
        ordering = ("-created_at", "-id")

    def __str__(self) -> str:
        return f"Termo #{self.id}"


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
    numero = models.PositiveIntegerField(null=True, blank=True, db_index=True)
    ano = models.PositiveIntegerField(null=True, blank=True, db_index=True)
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
    motorista_oficio_numero = models.PositiveIntegerField(null=True, blank=True)
    motorista_oficio_ano = models.PositiveIntegerField(null=True, blank=True)
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
    created_at = models.DateTimeField(auto_now_add=True, db_index=True)
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        constraints = [
            models.UniqueConstraint(
                fields=["ano", "numero"],
                name="uniq_oficio_numero_por_ano",
            )
        ]

    @property
    def is_draft(self) -> bool:
        return self.status == self.Status.DRAFT

    @property
    def numero_formatado(self) -> str:
        return format_oficio_num(self.numero, self.ano)

    @property
    def motorista_oficio_formatado(self) -> str:
        return format_oficio_num(self.motorista_oficio_numero, self.motorista_oficio_ano)

    @property
    def protocolo_formatado(self) -> str:
        return format_protocolo_num(self.protocolo)

    @property
    def motorista_protocolo_formatado(self) -> str:
        return format_protocolo_num(self.motorista_protocolo)

    def __str__(self) -> str:
        destino = self.cidade_destino or self.get_destino_display() or self.destino
        return f"Oficio {self.numero_formatado or self.oficio} - {destino}"

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

    def _sync_numero_from_legacy(self) -> None:
        self.oficio = normalize_oficio_num(self.oficio)
        legacy_numero, legacy_ano = split_oficio_num(self.oficio)
        if self.numero is None and legacy_numero is not None:
            self.numero = legacy_numero
        if self.ano is None and legacy_ano is not None:
            self.ano = legacy_ano
        if self.numero is not None and int(self.numero) <= 0:
            self.numero = None
        if self.ano is not None and int(self.ano) <= 0:
            self.ano = None

    def _sync_motorista_oficio_from_legacy(self) -> None:
        self.motorista_oficio = normalize_oficio_num(self.motorista_oficio)
        legacy_numero, legacy_ano = split_oficio_num(self.motorista_oficio)
        if self.motorista_oficio_numero is None and legacy_numero is not None:
            self.motorista_oficio_numero = legacy_numero
        if self.motorista_oficio_ano is None and legacy_ano is not None:
            self.motorista_oficio_ano = legacy_ano
        if (
            self.motorista_oficio_numero is not None
            and int(self.motorista_oficio_numero) > 0
            and not self.motorista_oficio_ano
        ):
            self.motorista_oficio_ano = timezone.localdate().year
        if not self.motorista_oficio_numero:
            self.motorista_oficio_ano = None
        if self.motorista_oficio_numero is not None and int(self.motorista_oficio_numero) <= 0:
            self.motorista_oficio_numero = None
        if self.motorista_oficio_ano is not None and int(self.motorista_oficio_ano) <= 0:
            self.motorista_oficio_ano = None

    def _sync_legacy_from_parts(self) -> None:
        self.oficio = self.numero_formatado or ""
        self.motorista_oficio = self.motorista_oficio_formatado or ""

    @staticmethod
    def _reserve_next_numero_for_year(ano: int) -> int:
        counter, _ = OficioCounter.objects.select_for_update().get_or_create(
            ano=ano,
            defaults={"last_numero": 0},
        )
        counter.last_numero += 1
        counter.save(update_fields=["last_numero", "updated_at"])
        return counter.last_numero

    @classmethod
    def reserve_next_oficio_number(cls, ano: int) -> int:
        with transaction.atomic():
            return cls._reserve_next_numero_for_year(ano)

    @staticmethod
    def _ensure_counter_floor(ano: int, numero: int) -> None:
        counter, _ = OficioCounter.objects.select_for_update().get_or_create(
            ano=ano,
            defaults={"last_numero": 0},
        )
        if numero > counter.last_numero:
            counter.last_numero = numero
            counter.save(update_fields=["last_numero", "updated_at"])

    def clean(self) -> None:
        super().clean()
        self._sync_numero_from_legacy()
        if self.numero is not None and self.ano is None:
            self.ano = timezone.localdate().year
        self._sync_motorista_oficio_from_legacy()
        self._sync_legacy_from_parts()
        self.protocolo = normalize_protocolo_num(self.protocolo)
        self.motorista_protocolo = normalize_protocolo_num(self.motorista_protocolo)
        self.motorista = normalize_upper_text(self.motorista)
        protocol_errors: dict[str, str] = {}
        if self.protocolo and len(self.protocolo) != 9:
            protocol_errors["protocolo"] = "Protocolo deve conter 9 digitos."
        if self.motorista_protocolo and len(self.motorista_protocolo) != 9:
            protocol_errors["motorista_protocolo"] = (
                "Protocolo do motorista deve conter 9 digitos."
            )
        if protocol_errors:
            raise ValidationError(protocol_errors)
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
        if self.motorista_carona:
            errors: dict[str, str] = {}
            if not self.motorista_oficio_numero:
                errors["motorista_oficio"] = "Informe o numero do oficio do motorista."
            if not self.motorista_oficio_ano:
                errors["motorista_oficio"] = "Informe o ano do oficio do motorista."
            if not self.motorista_protocolo:
                errors["motorista_protocolo"] = "Informe o protocolo do motorista."
            if errors:
                raise ValidationError(errors)

    def save(self, *args, **kwargs):
        self._sync_numero_from_legacy()
        if self.numero is not None and self.ano is None:
            self.ano = timezone.localdate().year
        if self.numero is None:
            self.ano = self.ano or timezone.localdate().year
        self._sync_motorista_oficio_from_legacy()
        self._sync_legacy_from_parts()
        self.protocolo = normalize_protocolo_num(self.protocolo)
        self.motorista_protocolo = normalize_protocolo_num(self.motorista_protocolo)
        self.motorista = normalize_upper_text(self.motorista)
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

        creating = self.pk is None
        if creating and self.numero is None:
            with transaction.atomic():
                self.numero = self._reserve_next_numero_for_year(int(self.ano or timezone.localdate().year))
                self._sync_legacy_from_parts()
                super().save(*args, **kwargs)
            return

        if self.numero is not None and self.ano is not None:
            with transaction.atomic():
                self._ensure_counter_floor(int(self.ano), int(self.numero))
                self._sync_legacy_from_parts()
                update_fields = kwargs.get("update_fields")
                if update_fields is not None:
                    update_fields = set(update_fields)
                    update_fields.update(
                        {
                            "destino",
                            "oficio",
                            "numero",
                            "ano",
                            "protocolo",
                            "motorista",
                            "motorista_oficio",
                            "motorista_oficio_numero",
                            "motorista_oficio_ano",
                            "motorista_protocolo",
                        }
                    )
                    if "nome_instituicao_custeio" not in update_fields and not (
                        self.nome_instituicao_custeio or ""
                    ):
                        update_fields.add("nome_instituicao_custeio")
                    kwargs["update_fields"] = list(update_fields)
                super().save(*args, **kwargs)
            return

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
