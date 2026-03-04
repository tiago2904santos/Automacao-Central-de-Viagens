import re
from datetime import datetime
from decimal import Decimal, InvalidOperation

from django.core.exceptions import ValidationError
from django.db import models, transaction
from django.urls import reverse
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
    is_ascom = models.BooleanField(
        default=True,
        verbose_name="É da ASCOM?",
        help_text="Se sim, não exige Termo de Autorização no pacote do evento.",
    )

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
    ordem = models.PositiveIntegerField(default=0)
    ativo = models.BooleanField(default=True)
    is_coordenador = models.BooleanField(default=False)

    class Meta:
        ordering = ("ordem", "nome")

    def __str__(self) -> str:
        return self.nome


class Efetivo(models.Model):
    cargo = models.OneToOneField(
        Cargo,
        on_delete=models.CASCADE,
        related_name="efetivo",
    )
    quantidade = models.PositiveIntegerField(default=0)
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        ordering = ("cargo__ordem", "cargo__nome")

    def __str__(self) -> str:
        return f"{self.cargo.nome}: {self.quantidade}"


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
    assinante_justificativa = models.ForeignKey(
        Viajante,
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name="oficio_config_assinante_justificativa",
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


class ModeloJustificativa(models.Model):
    """Modelos de texto pré-prontos para justificativas (prazo &lt; 10 dias)."""
    codigo = models.CharField(max_length=80, unique=True, help_text="Ex: recebimento_tardio")
    label = models.CharField(max_length=200, verbose_name="Nome do modelo")
    texto = models.TextField(verbose_name="Texto da justificativa")
    ordem = models.PositiveIntegerField(default=0)
    ativo = models.BooleanField(default=True)
    padrao = models.BooleanField(
        default=False,
        help_text="Se marcado, este modelo será o padrão no gerador.",
    )
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        verbose_name = "Modelo de justificativa"
        verbose_name_plural = "Modelos de justificativa"
        ordering = ("ordem", "label")

    def __str__(self) -> str:
        return self.label

    def save(self, *args, **kwargs):
        if self.padrao:
            ModeloJustificativa.objects.exclude(pk=self.pk).update(padrao=False)
        super().save(*args, **kwargs)


class OficioCounter(models.Model):
    ano = models.IntegerField(unique=True)
    last_numero = models.IntegerField(default=0)
    updated_at = models.DateTimeField(auto_now=True)

    def __str__(self) -> str:
        return f"{self.ano}: {self.last_numero}"


class PlanoTrabalhoCounter(models.Model):
    ano = models.PositiveIntegerField(unique=True)
    last_num = models.PositiveIntegerField(default=0)
    updated_at = models.DateTimeField(auto_now=True)

    def __str__(self) -> str:
        return f"Plano {self.ano}: {self.last_num}"


class OrdemServicoCounter(models.Model):
    ano = models.PositiveIntegerField(unique=True)
    last_num = models.PositiveIntegerField(default=0)
    updated_at = models.DateTimeField(auto_now=True)

    def __str__(self) -> str:
        return f"Ordem {self.ano}: {self.last_num}"


def get_next_plano_num(ano: int) -> int:
    with transaction.atomic():
        counter, _ = PlanoTrabalhoCounter.objects.select_for_update().get_or_create(
            ano=int(ano),
            defaults={"last_num": 0},
        )
        counter.last_num += 1
        counter.save(update_fields=["last_num", "updated_at"])
        return int(counter.last_num)


def get_next_ordem_num(ano: int) -> int:
    with transaction.atomic():
        counter, _ = OrdemServicoCounter.objects.select_for_update().get_or_create(
            ano=int(ano),
            defaults={"last_num": 0},
        )
        counter.last_num += 1
        counter.save(update_fields=["last_num", "updated_at"])
        return int(counter.last_num)


class Evento(models.Model):
    """Pacote do evento: unidade central que agrupa roteiro, ofícios, plano/ordem, termos e justificativas."""

    class TipoDemanda(models.TextChoices):
        PCPR_NA_COMUNIDADE = "PCPR_NA_COMUNIDADE", "PCPR na Comunidade"
        OPERACAO_POLICIAL = "OPERACAO_POLICIAL", "Operação Policial"
        PARANA_EM_ACAO = "PARANA_EM_ACAO", "Paraná em Ação"
        OUTRO = "OUTRO", "Outro"

    titulo = models.CharField(max_length=255, verbose_name="Título / Nome do evento")
    tipo_demanda = models.CharField(
        max_length=32,
        choices=TipoDemanda.choices,
        default=TipoDemanda.OUTRO,
        blank=True,
        verbose_name="Tipo de demanda",
    )
    cidade_base = models.ForeignKey(
        "Cidade",
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name="eventos_base",
        verbose_name="Cidade base",
    )
    data_inicio = models.DateField(null=True, blank=True, verbose_name="Data início")
    data_fim = models.DateField(null=True, blank=True, verbose_name="Data fim")
    tem_convite_ou_oficio_evento = models.BooleanField(
        default=True,
        verbose_name="Tem ofício solicitando ou convite do evento?",
        help_text="Se não, o evento exige Plano de Trabalho ou Ordem de Serviço (1 por evento).",
    )
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        verbose_name = "Evento"
        verbose_name_plural = "Eventos"
        ordering = ["-created_at"]

    def __str__(self) -> str:
        return self.titulo or f"Evento #{self.pk}"


class DocumentoEventoArquivo(models.Model):
    """Arquivo de documento do pacote do evento: gerado pelo sistema ou assinado (upload)."""

    class Tipo(models.TextChoices):
        OFICIO_ASSINADO = "OFICIO_ASSINADO", "Ofício assinado"
        SOLICITACAO_FORMAL_ASSINADA = "SOLICITACAO_FORMAL_ASSINADA", "Solicitação formal (convite/ofício) assinada"
        PLANO_ASSINADO = "PLANO_ASSINADO", "Plano de trabalho assinado"
        ORDEM_ASSINADO = "ORDEM_ASSINADO", "Ordem de serviço assinada"
        JUSTIFICATIVA_ASSINADA = "JUSTIFICATIVA_ASSINADA", "Justificativa assinada"
        TERMO_ASSINADO = "TERMO_ASSINADO", "Termo de autorização assinado"

    evento = models.ForeignKey(
        Evento,
        on_delete=models.CASCADE,
        related_name="arquivos_documentos",
        verbose_name="Evento",
    )
    tipo = models.CharField(max_length=32, choices=Tipo.choices, db_index=True)
    oficio = models.ForeignKey(
        "Oficio",
        on_delete=models.CASCADE,
        null=True,
        blank=True,
        related_name="arquivos_evento",
        verbose_name="Ofício",
    )
    viajante = models.ForeignKey(
        Viajante,
        on_delete=models.CASCADE,
        null=True,
        blank=True,
        related_name="arquivos_termo_evento",
        verbose_name="Servidor (termo)",
    )
    arquivo = models.FileField(
        upload_to="evento_documentos/%Y/%m/",
        max_length=500,
        verbose_name="Arquivo",
    )
    original_name = models.CharField(max_length=255, blank=True)
    mime_type = models.CharField(max_length=100, blank=True)
    is_active = models.BooleanField(
        default=True,
        help_text="Apenas o ativo por (evento, tipo, ofício/viajante) conta para checklist.",
    )
    uploaded_at = models.DateTimeField(auto_now_add=True)
    uploaded_by_id = models.PositiveIntegerField(null=True, blank=True, db_index=True)

    class Meta:
        verbose_name = "Arquivo de documento do evento"
        verbose_name_plural = "Arquivos de documentos do evento"
        ordering = ["-uploaded_at"]
        indexes = [
            models.Index(fields=["evento", "tipo", "oficio"], name="deae_ev_tipo_of"),
            models.Index(fields=["evento", "tipo", "viajante"], name="deae_ev_tipo_via"),
        ]

    def __str__(self) -> str:
        return f"{self.get_tipo_display()} — {self.evento_id}"


class EventoProtocoloArquivo(models.Model):
    """PDF compilado do protocolo (anexo único) do evento."""

    evento = models.ForeignKey(
        Evento,
        on_delete=models.CASCADE,
        related_name="protocolos_compilados",
        verbose_name="Evento",
    )
    pdf_compilado = models.FileField(
        upload_to="evento_protocolo/%Y/%m/",
        max_length=500,
        verbose_name="PDF compilado",
    )
    compilado_em = models.DateTimeField(auto_now_add=True)
    compilado_por_id = models.PositiveIntegerField(null=True, blank=True, db_index=True)
    hash_sha256 = models.CharField(max_length=64, blank=True)
    versao = models.PositiveIntegerField(default=1)

    class Meta:
        verbose_name = "Protocolo compilado do evento"
        verbose_name_plural = "Protocolos compilados"
        ordering = ["-compilado_em"]

    def __str__(self) -> str:
        return f"Protocolo {self.evento_id} v{self.versao}"


class AcaoInstitucionalManager(models.Manager):
    def get_completas(self):
        """Retorna ações com Plano de Trabalho e Ordem de Serviço vinculados."""
        return self.filter(plano_trabalho__isnull=False, ordem_servico__isnull=False)

    def esta_completa(self, pk):
        """Verifica se a ação indicada possui plano e ordem de serviço."""
        return self.filter(
            pk=pk,
            plano_trabalho__isnull=False,
            ordem_servico__isnull=False,
        ).exists()


class AcaoInstitucional(models.Model):
    titulo = models.CharField(max_length=255, verbose_name="Titulo da Acao")
    descricao = models.TextField(blank=True, verbose_name="Descricao Detalhada")
    criado_em = models.DateTimeField(auto_now_add=True, verbose_name="Criado em")
    atualizado_em = models.DateTimeField(auto_now=True, verbose_name="Atualizado em")

    objects = AcaoInstitucionalManager()

    class Meta:
        verbose_name = "Acao Institucional"
        verbose_name_plural = "Acoes Institucionais"
        ordering = ["-criado_em"]

    def __str__(self) -> str:
        return self.titulo

    @property
    def esta_completa_prop(self) -> bool:
        try:
            plano = self.plano_trabalho
        except PlanoTrabalho.DoesNotExist:
            plano = None
        try:
            ordem = self.ordem_servico
        except OrdemServico.DoesNotExist:
            ordem = None
        return plano is not None and ordem is not None


class TermoAutorizacao(models.Model):
    acao = models.ForeignKey(
        AcaoInstitucional,
        on_delete=models.CASCADE,
        related_name="termos",
        null=True,
        blank=True,
        verbose_name="Acao Institucional",
    )
    oficio = models.ForeignKey(
        "Oficio",
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name="termos_autorizacao",
        verbose_name="Ofício",
    )
    evento = models.ForeignKey(
        "Evento",
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name="termos",
        verbose_name="Evento",
    )
    viajante = models.ForeignKey(
        Viajante,
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name="termos_autorizacao",
        verbose_name="Servidor",
    )
    data_inicio = models.DateField()
    data_fim = models.DateField(null=True, blank=True)
    data_unica = models.BooleanField(default=False)
    destinos = models.JSONField(default=list, blank=True)
    dispensado = models.BooleanField(
        default=False,
        verbose_name="Termo dispensado",
        help_text="Se marcado, não exige geração/assinatura do termo (motivo em dispensa_motivo).",
    )
    dispensa_motivo = models.CharField(max_length=200, blank=True, default="")
    motorista_nome = models.CharField(max_length=120, blank=True, default="")
    veiculo_modelo = models.CharField(max_length=120, blank=True, default="")
    veiculo_placa = models.CharField(max_length=20, blank=True, default="")
    combustivel = models.CharField(max_length=60, blank=True, default="")
    created_at = models.DateTimeField(auto_now_add=True, db_index=True)
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        verbose_name = "Termo de autorizacao"
        verbose_name_plural = "Termos de autorizacao"
        ordering = ("-created_at", "-id")
        constraints = [
            models.UniqueConstraint(
                fields=["oficio", "viajante"],
                condition=models.Q(dispensado=False) & models.Q(oficio__isnull=False) & models.Q(viajante__isnull=False),
                name="uniq_termo_oficio_viajante_ativo",
            ),
        ]

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

    acao = models.ForeignKey(
        AcaoInstitucional,
        on_delete=models.SET_NULL,
        related_name="oficios",
        null=True,
        blank=True,
        verbose_name="Acao Institucional",
    )
    evento = models.ForeignKey(
        "Evento",
        on_delete=models.SET_NULL,
        related_name="oficios",
        null=True,
        blank=True,
        verbose_name="Evento (pacote)",
    )
    roteiro = models.ForeignKey(
        "Roteiro",
        on_delete=models.SET_NULL,
        related_name="oficios_como_roteiro",
        null=True,
        blank=True,
        verbose_name="Roteiro de Viagem",
    )
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
    justificativa_modelo = models.CharField(
        max_length=80,
        blank=True,
        default="",
        help_text="Código do modelo de justificativa usado (ex: recebimento_tardio).",
    )
    justificativa_texto = models.TextField(
        blank=True,
        default="",
        help_text="Texto da justificativa preenchido (prazo < 10 dias). Preenchido desbloqueia geração do ofício.",
    )
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

    def get_admin_url(self) -> str:
        return reverse(
            f"admin:{self._meta.app_label}_{self._meta.model_name}_change",
            args=(self.pk,),
        )

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

    def _build_acao_defaults(self) -> dict[str, str]:
        identificador = self.numero_formatado or self.oficio or str(self.pk or "")
        identificador = identificador.strip() or "Sem número"
        return {
            "titulo": f"Ação — {identificador}",
            "descricao": (
                f"Ação institucional derivada do Ofício {identificador}."
                if identificador
                else ""
            ),
        }

    def ensure_acao(self):
        if self.acao_id:
            return self.acao
        if not self.pk:
            return None
        acao = AcaoInstitucional.objects.create(**self._build_acao_defaults())
        type(self).objects.filter(pk=self.pk).update(acao=acao)
        self.acao = acao
        return acao

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
            self.ensure_acao()
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
            self.ensure_acao()
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
        self.ensure_acao()


class Roteiro(models.Model):
    """Roteiro reutilizavel independente de qualquer documento."""

    class TipoDeslocamentoChoices(models.TextChoices):
        INTERIOR = "INTERIOR", "Interior"
        CAPITAL = "CAPITAL", "Capital"
        OUTRO = "OUTRO", "Outro"

    evento = models.ForeignKey(
        "Evento",
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name="roteiros",
        verbose_name="Evento",
    )
    nome = models.CharField(max_length=300, blank=True, verbose_name="Nome do Roteiro")
    descricao = models.TextField(blank=True, verbose_name="Descricao")
    estado_sede = models.ForeignKey(
        Estado,
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name="roteiros_sede_estado",
    )
    cidade_sede = models.ForeignKey(
        Cidade,
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name="roteiros_sede_cidade",
    )
    uf_origem = models.CharField(max_length=2, default="PR", verbose_name="UF Origem")
    cidade_origem = models.CharField(max_length=100, verbose_name="Cidade Origem")
    uf_destino = models.CharField(max_length=2, default="PR", verbose_name="UF Destino")
    cidade_destino = models.CharField(max_length=100, verbose_name="Cidade Destino")
    retorno_saida_cidade = models.CharField(max_length=120, blank=True, default="")
    retorno_saida_data = models.DateField(null=True, blank=True)
    retorno_saida_hora = models.TimeField(null=True, blank=True)
    retorno_chegada_cidade = models.CharField(max_length=120, blank=True, default="")
    retorno_chegada_data = models.DateField(null=True, blank=True)
    retorno_chegada_hora = models.TimeField(null=True, blank=True)
    tempo_viagem = models.TimeField(
        null=True,
        blank=True,
        help_text="Duracao no formato HH:MM",
    )
    criado_automaticamente = models.BooleanField(
        default=False,
        help_text="True quando o roteiro e clonado automaticamente ao ser alterado no contexto de um oficio.",
    )
    distancia_km = models.DecimalField(
        max_digits=8,
        decimal_places=2,
        null=True,
        blank=True,
        verbose_name="Distancia (km)",
    )
    tipo_deslocamento = models.CharField(
        max_length=10,
        choices=TipoDeslocamentoChoices.choices,
        verbose_name="Tipo de Deslocamento",
    )
    criado_em = models.DateTimeField(auto_now_add=True, verbose_name="Criado em")
    atualizado_em = models.DateTimeField(auto_now=True, verbose_name="Atualizado em")
    ativo = models.BooleanField(default=True, verbose_name="Ativo")

    class Meta:
        ordering = ["-criado_em", "-id"]
        verbose_name = "Roteiro"
        verbose_name_plural = "Roteiros"

    def __str__(self) -> str:
        nome = (self.nome or "").strip()
        if nome:
            return nome
        return (
            f"{self.cidade_origem}/{self.uf_origem} -> "
            f"{self.cidade_destino}/{self.uf_destino}"
        )

    @property
    def created_at(self):
        return self.criado_em

    @property
    def updated_at(self):
        return self.atualizado_em

    def get_admin_url(self) -> str:
        return reverse(
            f"admin:{self._meta.app_label}_{self._meta.model_name}_change",
            args=(self.pk,),
        )

    def _sync_legacy_fields(self):
        if self.estado_sede_id and self.estado_sede:
            self.uf_origem = self.estado_sede.sigla
        elif self.uf_origem and not self.estado_sede_id:
            self.estado_sede = self.uf_sede_obj

        if self.cidade_sede_id and self.cidade_sede:
            self.cidade_origem = self.cidade_sede.nome
        elif self.cidade_origem and not self.cidade_sede_id:
            self.cidade_sede = self.cidade_sede_obj

        if self.pk and not self.retorno_saida_cidade and self.trechos.exists():
            ultimo_trecho = self.trechos.order_by("ordem").last()
            if ultimo_trecho:
                self.retorno_saida_cidade = ultimo_trecho.destino_cidade_nome

        if not self.retorno_chegada_cidade and self.cidade_sede:
            self.retorno_chegada_cidade = self.cidade_sede.nome

    def gerar_nome(self):
        sede = self.cidade_sede.nome if self.cidade_sede else "?"
        primeiro_trecho = self.trechos.order_by("ordem").first()

        if primeiro_trecho and primeiro_trecho.destino_cidade:
            destino = primeiro_trecho.destino_cidade.nome
        elif primeiro_trecho and primeiro_trecho.cidade_destino:
            destino = primeiro_trecho.cidade_destino
        else:
            destino = ""

        data_hora = ""
        if primeiro_trecho and primeiro_trecho.saida_data:
            data_hora = primeiro_trecho.saida_data.strftime("%d/%m/%Y")
            if primeiro_trecho.saida_hora:
                data_hora += f" {primeiro_trecho.saida_hora.strftime('%H:%M')}"

        partes = [sede]
        if destino:
            partes.append(destino)
        nome = " > ".join(partes)
        if data_hora:
            nome += f" {data_hora}"
        return nome

    def save(self, *args, **kwargs):
        self._sync_legacy_fields()
        super().save(*args, **kwargs)

        if not self.nome:
            nome_gerado = self.gerar_nome()
            if nome_gerado != self.nome:
                self.nome = nome_gerado
                Roteiro.objects.filter(pk=self.pk).update(nome=nome_gerado)

    def get_distancia_total(self) -> Decimal | None:
        if not self.pk:
            return self.distancia_km
        if self.trechos.exists():
            total = Decimal("0")
            for trecho in self.trechos.all():
                if trecho.distancia_km is not None:
                    total += trecho.distancia_km
            return total
        return self.distancia_km

    def get_destinos_display(self) -> str:
        trechos = self.trechos.order_by("ordem")
        return " -> ".join(
            f"{trecho.destino_cidade_nome}/{trecho.destino_estado_sigla}"
            for trecho in trechos
            if trecho.destino_cidade_nome and trecho.destino_estado_sigla
        )

    def get_trechos_preview(self) -> str:
        trechos = list(self.trechos.order_by("ordem"))
        partes: list[str] = []
        if self.cidade_sede_nome and self.uf_origem:
            partes.append(f"{self.cidade_sede_nome}/{self.uf_origem}")

        for trecho in trechos:
            if trecho.destino_cidade_nome and trecho.destino_estado_sigla:
                partes.append(f"{trecho.destino_cidade_nome}/{trecho.destino_estado_sigla}")

        if self.cidade_sede_nome and self.uf_origem and trechos:
            partes.append(f"{self.cidade_sede_nome}/{self.uf_origem}")

        if not partes:
            return "Nenhum trecho definido"
        return " -> ".join(partes)

    @property
    def cidades_destino(self):
        return self.trechos.order_by("ordem")

    @property
    def total_cidades(self) -> int:
        if not self.pk:
            return 0
        return self.trechos.count()

    def get_cards_gerados(self) -> list[dict[str, str | int]]:
        destinos = [
            trecho
            for trecho in self.trechos.order_by("ordem")
            if trecho.destino_cidade_nome and trecho.destino_estado_sigla
        ]
        if not destinos:
            return []

        sequencia = [{"cidade": self.cidade_sede_nome, "uf": self.uf_origem}]
        sequencia.extend(
            {"cidade": trecho.destino_cidade_nome, "uf": trecho.destino_estado_sigla}
            for trecho in destinos
        )
        sequencia.append({"cidade": self.cidade_sede_nome, "uf": self.uf_origem})

        cards: list[dict[str, str | int]] = []
        last_leg_index = len(sequencia) - 2
        for idx in range(len(sequencia) - 1):
            origem = sequencia[idx]
            destino = sequencia[idx + 1]
            cards.append(
                {
                    "numero": idx + 1,
                    "origem_cidade": origem["cidade"],
                    "origem_uf": origem["uf"],
                    "destino_cidade": destino["cidade"],
                    "destino_uf": destino["uf"],
                    "label": "Retorno"
                    if idx == last_leg_index
                    else f"Trecho {idx + 1} (Ida)",
                }
            )
        return cards

    @property
    def uf_sede_obj(self):
        if self.estado_sede_id:
            return self.estado_sede
        return Estado.objects.filter(sigla__iexact=(self.uf_origem or "").strip()).first()

    @property
    def cidade_sede_obj(self):
        if self.cidade_sede_id:
            return self.cidade_sede
        uf = self.uf_sede_obj
        nome = (self.cidade_origem or "").strip()
        if not nome:
            return None
        qs = Cidade.objects.filter(nome__iexact=nome)
        if uf:
            qs = qs.filter(estado=uf)
        return qs.first()

    @property
    def uf_sede_id(self):
        return self.uf_sede_obj.id if self.uf_sede_obj else None

    @property
    def cidade_sede_id(self):
        return self.cidade_sede_obj.id if self.cidade_sede_obj else None

    @property
    def cidade_sede_nome(self) -> str:
        cidade = self.cidade_sede_obj
        return cidade.nome if cidade else (self.cidade_origem or "")


class TrechoRoteiro(models.Model):
    """Segmentos detalhados dentro de um roteiro reutilizavel."""

    class ModalChoices(models.TextChoices):
        VEICULO_PROPRIO = "veiculo_proprio", "Veiculo Proprio"
        VEICULO_LOCADO = "veiculo_locado", "Veiculo Locado"
        ONIBUS = "onibus", "Onibus"
        AVIAO = "aviao", "Aviao"
        OUTRO = "outro", "Outro"

    roteiro = models.ForeignKey(
        Roteiro,
        on_delete=models.CASCADE,
        related_name="trechos",
        verbose_name="Roteiro",
    )
    ordem = models.PositiveIntegerField(default=1, verbose_name="Ordem")
    origem_estado = models.ForeignKey(
        Estado,
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name="trechos_roteiro_origem_estado",
    )
    origem_cidade = models.ForeignKey(
        Cidade,
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name="trechos_roteiro_origem_cidade",
    )
    destino_estado = models.ForeignKey(
        Estado,
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name="trechos_roteiro_destino_estado",
    )
    destino_cidade = models.ForeignKey(
        Cidade,
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name="trechos_roteiro_destino_cidade",
    )
    uf_origem = models.CharField(max_length=2, default="PR", verbose_name="UF Origem")
    cidade_origem = models.CharField(max_length=100, verbose_name="Cidade Origem")
    uf_destino = models.CharField(max_length=2, default="PR", verbose_name="UF Destino")
    cidade_destino = models.CharField(max_length=100, verbose_name="Cidade Destino")
    saida_data = models.DateField(null=True, blank=True)
    saida_hora = models.TimeField(null=True, blank=True)
    chegada_data = models.DateField(null=True, blank=True)
    chegada_hora = models.TimeField(null=True, blank=True)
    retorno_saida_data = models.DateField(null=True, blank=True)
    retorno_saida_hora = models.TimeField(null=True, blank=True)
    retorno_chegada_data = models.DateField(null=True, blank=True)
    retorno_chegada_hora = models.TimeField(null=True, blank=True)
    tempo_viagem_minutos = models.PositiveIntegerField(
        null=True,
        blank=True,
        help_text="Tempo estimado em minutos",
    )
    distancia_km = models.DecimalField(
        max_digits=8,
        decimal_places=2,
        null=True,
        blank=True,
        verbose_name="Distancia (km)",
    )
    modal = models.CharField(
        max_length=50,
        choices=ModalChoices.choices,
        default=ModalChoices.VEICULO_PROPRIO,
    )
    observacao = models.TextField(blank=True)

    class Meta:
        ordering = ["roteiro", "ordem", "id"]
        unique_together = [("roteiro", "ordem")]
        verbose_name = "Trecho do Roteiro"
        verbose_name_plural = "Trechos do Roteiro"

    def __str__(self) -> str:
        return (
            f"{self.roteiro.nome} - Trecho {self.ordem}: "
            f"{self.origem_cidade_nome} -> {self.destino_cidade_nome}"
        )

    @property
    def origem_cidade_nome(self):
        if self.origem_cidade_id:
            return self.origem_cidade.nome
        return self.cidade_origem

    @property
    def origem_estado_sigla(self):
        if self.origem_estado_id:
            return self.origem_estado.sigla
        return self.uf_origem

    @property
    def destino_cidade_nome(self):
        if self.destino_cidade_id:
            return self.destino_cidade.nome
        return self.cidade_destino

    @property
    def destino_estado_sigla(self):
        if self.destino_estado_id:
            return self.destino_estado.sigla
        return self.uf_destino

    def save(self, *args, **kwargs):
        if self.origem_estado_id and self.origem_estado:
            self.uf_origem = self.origem_estado.sigla
        if self.origem_cidade_id and self.origem_cidade:
            self.cidade_origem = self.origem_cidade.nome
        if self.destino_estado_id and self.destino_estado:
            self.uf_destino = self.destino_estado.sigla
        if self.destino_cidade_id and self.destino_cidade:
            self.cidade_destino = self.destino_cidade.nome
        super().save(*args, **kwargs)

    @property
    def cidade_destino_obj(self):
        if self.destino_cidade_id:
            return self.destino_cidade
        nome = (self.cidade_destino or "").strip()
        if not nome:
            return None
        qs = Cidade.objects.filter(nome__iexact=nome)
        if self.uf_destino:
            qs = qs.filter(estado__sigla__iexact=self.uf_destino)
        return qs.first()


class OficioRoteiro(models.Model):
    """Vinculo entre um oficio e um roteiro reutilizavel."""

    oficio = models.ForeignKey(
        "Oficio",
        on_delete=models.CASCADE,
        related_name="roteiros_vinculados",
        verbose_name="Oficio",
    )
    roteiro = models.ForeignKey(
        Roteiro,
        on_delete=models.CASCADE,
        related_name="oficios_vinculados",
        verbose_name="Roteiro",
    )
    vinculado_em = models.DateTimeField(auto_now_add=True, verbose_name="Vinculado em")
    observacao = models.CharField(max_length=255, blank=True, verbose_name="Observacao")

    class Meta:
        unique_together = [("oficio", "roteiro")]
        verbose_name = "Vinculo Oficio-Roteiro"
        verbose_name_plural = "Vinculos Oficio-Roteiro"

    def __str__(self) -> str:
        identificador = self.oficio.numero_formatado or self.oficio.oficio or str(self.oficio_id)
        return f"Oficio {identificador} vinculado ao roteiro '{self.roteiro.nome}'"


RoteiroViagem = Roteiro
TrechoRoteiroViagem = TrechoRoteiro
TrechoRoteiroDestino = TrechoRoteiro


class PlanoTrabalho(models.Model):
    oficio = models.OneToOneField(
        Oficio,
        on_delete=models.CASCADE,
        related_name="plano_trabalho",
    )
    acao = models.OneToOneField(
        AcaoInstitucional,
        on_delete=models.CASCADE,
        related_name="plano_trabalho",
        null=True,
        blank=True,
        verbose_name="Acao Institucional",
    )
    numero = models.PositiveIntegerField()
    ano = models.PositiveIntegerField()
    sigla_unidade = models.CharField(max_length=30, blank=True, default="ASCOM")
    programa_projeto = models.CharField(max_length=200, blank=True, default="")
    solicitantes_json = models.JSONField(default=list, blank=True)
    destino = models.CharField(max_length=200, blank=True, default="")
    destinos_json = models.JSONField(default=list, blank=True)
    solicitante = models.CharField(max_length=200, blank=True, default="")
    contexto_solicitacao = models.TextField(blank=True, default="")
    local = models.CharField(max_length=120, default="")
    data_inicio = models.DateField()
    data_fim = models.DateField()
    horario_inicio = models.TimeField(null=True, blank=True)
    horario_fim = models.TimeField(null=True, blank=True)
    horario_atendimento = models.CharField(max_length=120, blank=True, default="")
    efetivo_json = models.JSONField(default=list, blank=True)
    efetivo_formatado = models.CharField(max_length=200, blank=True, default="")
    unidade_movel = models.BooleanField(default=False)
    estrutura_apoio = models.TextField(blank=True, default="")
    efetivo_por_dia = models.PositiveIntegerField(default=0)
    quantidade_servidores = models.PositiveIntegerField(default=0)
    composicao_diarias = models.CharField(max_length=200, blank=True, default="")
    valor_unitario = models.DecimalField(
        max_digits=12,
        decimal_places=2,
        null=True,
        blank=True,
    )
    valor_total_calculado = models.DecimalField(
        max_digits=12,
        decimal_places=2,
        null=True,
        blank=True,
    )
    valor_total = models.CharField(max_length=120, blank=True, default="")
    possui_coordenador_municipal = models.BooleanField(default=False)
    coordenador_plano = models.ForeignKey(
        Viajante,
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name="planos_trabalho_coordenados",
    )
    coordenador_municipal = models.ForeignKey(
        "CoordenadorMunicipal",
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name="planos_trabalho",
    )
    coordenador_nome = models.CharField(max_length=200, blank=True, default="")
    coordenador_cargo = models.CharField(max_length=200, blank=True, default="")
    texto_override = models.TextField(blank=True, default="")
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        constraints = [
            models.UniqueConstraint(
                fields=["ano", "numero"],
                name="uniq_plano_trabalho_numero_por_ano",
            )
        ]

    def __str__(self) -> str:
        return f"Plano de Trabalho {self.numero}/{self.ano}"

    @property
    def titulo_formatado(self) -> str:
        sigla = (self.sigla_unidade or "").strip().upper()
        suffix = f"/{sigla}" if sigla else ""
        return f"PLANO DE TRABALHO N\u00ba{int(self.numero or 0):02d}/{int(self.ano or 0)}{suffix}"

    def get_coordenador_administrativo_nome(self) -> str:
        from viagens.services.plano_trabalho import DEFAULT_COORDENADOR_PLANO_NOME

        if self.coordenador_plano and (self.coordenador_plano.nome or "").strip():
            return " ".join((self.coordenador_plano.nome or "").split())
        nome = " ".join((self.coordenador_nome or "").split())
        return nome or DEFAULT_COORDENADOR_PLANO_NOME

    def get_coordenador_administrativo_cargo(self) -> str:
        from viagens.services.plano_trabalho import DEFAULT_COORDENADOR_PLANO_CARGO

        if self.coordenador_plano and (self.coordenador_plano.cargo or "").strip():
            return " ".join((self.coordenador_plano.cargo or "").split())
        cargo = " ".join((self.coordenador_cargo or "").split())
        return cargo or DEFAULT_COORDENADOR_PLANO_CARGO

    def get_coordenador_administrativo_display(self) -> str:
        cargo = self.get_coordenador_administrativo_cargo()
        nome = self.get_coordenador_administrativo_nome()
        return " ".join(part for part in (cargo, nome) if part).strip()

    def get_coordenacao_formatada(self) -> str:
        from viagens.services.plano_trabalho import build_coordenacao_formatada

        return build_coordenacao_formatada(self)

    @staticmethod
    def _parse_decimal(value) -> Decimal | None:
        if value in (None, ""):
            return None

    @staticmethod
    def _format_hora_ptbr(value) -> str:
        if value is None:
            return ""
        if value.minute:
            return value.strftime("%Hh%M")
        return value.strftime("%Hh")
        if isinstance(value, Decimal):
            return value
        try:
            normalized = str(value).strip().replace(".", "").replace(",", ".")
            return Decimal(normalized)
        except (InvalidOperation, TypeError, ValueError):
            return None

    def _composicao_fator(self) -> Decimal:
        raw = (self.composicao_diarias or "").strip()
        if not raw:
            return Decimal("1")
        pattern = re.compile(
            r"(?P<qtd>\d+(?:[.,]\d+)?)\s*x\s*(?P<pct>\d+(?:[.,]\d+)?)\s*%",
            re.IGNORECASE,
        )
        fator = Decimal("0")
        found = False
        for match in pattern.finditer(raw):
            found = True
            qtd = self._parse_decimal(match.group("qtd")) or Decimal("0")
            pct = self._parse_decimal(match.group("pct")) or Decimal("0")
            fator += qtd * (pct / Decimal("100"))
        if found and fator > 0:
            return fator
        fallback = self._parse_decimal(raw)
        if fallback and fallback > 0:
            return fallback
        return Decimal("1")

    def calcular_valor_total(self) -> Decimal | None:
        unitario = self._parse_decimal(self.valor_unitario)
        qtd_servidores = int(self.quantidade_servidores or 0)
        if not unitario or unitario <= 0 or qtd_servidores <= 0:
            return None
        fator = self._composicao_fator()
        total = (unitario * Decimal(qtd_servidores) * fator).quantize(Decimal("0.01"))
        if total <= 0:
            return None
        return total

    def clean(self) -> None:
        super().clean()
        errors: dict[str, str] = {}

        if self.data_inicio and self.data_fim and self.data_fim < self.data_inicio:
            errors["data_fim"] = "A data final deve ser igual ou posterior \u00e0 data inicial."

        if not self.possui_coordenador_municipal:
            self.coordenador_municipal = None

        if self.valor_unitario is not None and self.valor_unitario < 0:
            errors["valor_unitario"] = "Informe um valor unit\u00e1rio v\u00e1lido."
        if self.valor_total_calculado is not None and self.valor_total_calculado < 0:
            errors["valor_total_calculado"] = "Informe um valor total v\u00e1lido."

        if errors:
            raise ValidationError(errors)

    def save(self, *args, **kwargs):
        if not self.acao_id and self.oficio_id:
            self.acao = self.oficio.ensure_acao()
        if not self.ano:
            self.ano = timezone.localdate().year
        if not (self.sigla_unidade or "").strip():
            self.sigla_unidade = "ASCOM"
        if self.horario_inicio and self.horario_fim:
            from .services.plano_trabalho import formatar_horario_intervalo

            horario_formatado = formatar_horario_intervalo(
                self.horario_inicio,
                self.horario_fim,
            )
            if horario_formatado:
                self.horario_atendimento = horario_formatado

        if self.coordenador_plano:
            self.coordenador_nome = " ".join((self.coordenador_plano.nome or "").split())
            self.coordenador_cargo = " ".join((self.coordenador_plano.cargo or "").split())
        else:
            self.coordenador_nome = " ".join((self.coordenador_nome or "").split())
            self.coordenador_cargo = " ".join((self.coordenador_cargo or "").split())

        unitario_decimal = self._parse_decimal(self.valor_unitario)
        if unitario_decimal is not None:
            self.valor_unitario = unitario_decimal.quantize(Decimal("0.01"))

        total = self.calcular_valor_total()
        if total is not None:
            self.valor_total_calculado = total
        total_decimal = self._parse_decimal(self.valor_total_calculado)
        if total_decimal is not None:
            self.valor_total_calculado = total_decimal.quantize(Decimal("0.01"))
            bruto = f"{self.valor_total_calculado:.2f}".replace(".", ",")
            self.valor_total = f"R$ {bruto}"

        if not (self.efetivo_formatado or "").strip() and self.quantidade_servidores:
            from .services.plano_trabalho import format_total_servidores

            self.efetivo_formatado = format_total_servidores(int(self.quantidade_servidores))
        if not self.efetivo_por_dia and self.quantidade_servidores:
            self.efetivo_por_dia = int(self.quantidade_servidores)

        self.full_clean()
        super().save(*args, **kwargs)


class CoordenadorMunicipal(models.Model):
    nome = models.CharField(max_length=200)
    cargo = models.CharField(max_length=200)
    cidade = models.CharField(max_length=120)
    ativo = models.BooleanField(default=True)
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        ordering = ("nome",)

    def __str__(self) -> str:
        return f"{self.nome} - {self.cidade}"

    def clean(self) -> None:
        super().clean()
        self.nome = " ".join((self.nome or "").split())
        self.cargo = " ".join((self.cargo or "").split())
        self.cidade = " ".join((self.cidade or "").split())
        errors: dict[str, str] = {}
        if not self.nome:
            errors["nome"] = "Informe o nome do coordenador municipal."
        if not self.cargo:
            errors["cargo"] = "Informe o cargo do coordenador municipal."
        if not self.cidade:
            errors["cidade"] = "Informe a cidade do coordenador municipal."
        if errors:
            raise ValidationError(errors)

    def save(self, *args, **kwargs):
        self.nome = " ".join((self.nome or "").split())
        self.cargo = " ".join((self.cargo or "").split())
        self.cidade = " ".join((self.cidade or "").split())
        self.full_clean()
        super().save(*args, **kwargs)


class PlanoTrabalhoMeta(models.Model):
    plano = models.ForeignKey(
        PlanoTrabalho,
        on_delete=models.CASCADE,
        related_name="metas",
    )
    ordem = models.PositiveIntegerField(default=1)
    descricao = models.CharField(max_length=350)

    class Meta:
        ordering = ("ordem", "id")

    def __str__(self) -> str:
        return f"Meta {self.ordem} - Plano {self.plano_id}"


class PlanoTrabalhoAtividade(models.Model):
    plano = models.ForeignKey(
        PlanoTrabalho,
        on_delete=models.CASCADE,
        related_name="atividades",
    )
    ordem = models.PositiveIntegerField(default=1)
    descricao = models.CharField(max_length=350)

    class Meta:
        ordering = ("ordem", "id")

    def __str__(self) -> str:
        return f"Atividade {self.ordem} - Plano {self.plano_id}"


class PlanoTrabalhoRecurso(models.Model):
    plano = models.ForeignKey(
        PlanoTrabalho,
        on_delete=models.CASCADE,
        related_name="recursos",
    )
    ordem = models.PositiveIntegerField(default=1)
    descricao = models.CharField(max_length=350)

    class Meta:
        ordering = ("ordem", "id")

    def __str__(self) -> str:
        return f"Recurso {self.ordem} - Plano {self.plano_id}"


class PlanoTrabalhoLocalAtuacao(models.Model):
    plano = models.ForeignKey(
        PlanoTrabalho,
        on_delete=models.CASCADE,
        related_name="locais_atuacao",
    )
    ordem = models.PositiveIntegerField(default=1)
    data = models.DateField(null=True, blank=True)
    local = models.CharField(max_length=240)

    class Meta:
        ordering = ("ordem", "id")

    def __str__(self) -> str:
        return f"Local {self.ordem} - Plano {self.plano_id}"


class OrdemServico(models.Model):
    oficio = models.OneToOneField(
        Oficio,
        on_delete=models.CASCADE,
        related_name="ordem_servico",
    )
    acao = models.OneToOneField(
        AcaoInstitucional,
        on_delete=models.CASCADE,
        related_name="ordem_servico",
        null=True,
        blank=True,
        verbose_name="Acao Institucional",
    )
    numero = models.PositiveIntegerField()
    ano = models.PositiveIntegerField()
    referencia = models.CharField(max_length=200, default="Diligências")
    determinante_nome = models.CharField(max_length=200, blank=True, default="")
    determinante_cargo = models.CharField(max_length=200, blank=True, default="")
    finalidade = models.TextField(blank=True, default="")
    texto_override = models.TextField(blank=True, default="")
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        constraints = [
            models.UniqueConstraint(
                fields=["ano", "numero"],
                name="uniq_ordem_servico_numero_por_ano",
            )
        ]

    def __str__(self) -> str:
        return f"Ordem de Servico {self.numero}/{self.ano}"

    def save(self, *args, **kwargs):
        if not self.acao_id and self.oficio_id:
            self.acao = self.oficio.ensure_acao()
        if not self.ano:
            self.ano = timezone.localdate().year
        self.full_clean()
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
