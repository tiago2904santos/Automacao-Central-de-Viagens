from dataclasses import dataclass
from pathlib import Path

from django.conf import settings
from django.core.management.base import BaseCommand, CommandError
from django.db import connection, models, transaction

from viagens.models import Cargo, Cidade, Estado, Oficio, Trecho, Viajante, Veiculo
from viagens.utils_csv_import import (
    CSVImportLineError,
    CsvRow,
    CsvSource,
    detect_dataset_kind,
    get_first_value,
    is_blank_row,
    is_effectively_empty,
    normalize_key,
    parse_bool,
    parse_date,
    parse_datetime,
    parse_decimal,
    parse_time,
    read_csv_source,
)


@dataclass
class ImportCount:
    created: int = 0
    updated: int = 0
    skipped: int = 0


class Command(BaseCommand):
    help = "Restaura dados do banco a partir de CSVs em um diretorio."

    KIND_TO_MODEL = {
        "estado": Estado,
        "cidade": Cidade,
        "cargo": Cargo,
        "viajante": Viajante,
        "veiculo": Veiculo,
        "oficio": Oficio,
        "trecho": Trecho,
        "oficio_viajante": Oficio.viajantes.through,
    }

    IMPORT_ORDER = [
        "estado",
        "cidade",
        "cargo",
        "viajante",
        "veiculo",
        "oficio",
        "trecho",
        "oficio_viajante",
    ]

    REQUIRED_KINDS = {"estado", "cidade", "viajante"}

    def add_arguments(self, parser):
        parser.add_argument(
            "source_dir",
            help="Diretorio com os CSVs (ex.: data ou viagens/data).",
        )
        parser.add_argument(
            "--truncate",
            action="store_true",
            help="Apaga os dados das tabelas alvo antes de importar.",
        )
        parser.add_argument(
            "--dry-run",
            action="store_true",
            help="Valida e mostra a importacao, mas aplica rollback ao final.",
        )
        parser.add_argument(
            "--encoding",
            default="utf-8-sig",
            help="Encoding preferido (fallback automatico: utf-8-sig, utf-8, latin-1).",
        )
        parser.add_argument(
            "--delimiter",
            default="auto",
            choices=["auto", ";", ","],
            help="Delimitador CSV (auto tenta ';' e ',').",
        )

    def handle(self, *args, **options):
        source_dir = self._resolve_source_dir(options["source_dir"])
        preferred_encoding = options["encoding"]
        delimiter = options["delimiter"]
        truncate = bool(options["truncate"])
        dry_run = bool(options["dry_run"])

        csv_paths = sorted(source_dir.glob("*.csv"))
        if not csv_paths:
            raise CommandError(f"Nenhum CSV encontrado em: {source_dir}")

        detected_sources = self._load_and_detect_sources(
            csv_paths=csv_paths,
            preferred_encoding=preferred_encoding,
            delimiter=delimiter,
        )

        missing_required = sorted(self.REQUIRED_KINDS - set(detected_sources))
        if missing_required:
            missing_label = ", ".join(missing_required)
            raise CommandError(
                f"CSVs obrigatorios nao encontrados para: {missing_label}. "
                f"Arquivos lidos: {[path.name for path in csv_paths]}"
            )

        self.stdout.write(self.style.SUCCESS("Mapeamento de arquivos:"))
        for kind in self.IMPORT_ORDER:
            source = detected_sources.get(kind)
            if not source:
                continue
            self.stdout.write(
                f"- {source.path.name} -> {kind} "
                f"(encoding={source.encoding}, delimiter={source.delimiter})"
            )

        warnings: list[str] = []
        counts_by_kind: dict[str, ImportCount] = {}
        imported_models: list[type[models.Model]] = []

        try:
            with transaction.atomic():
                if truncate:
                    self._truncate_target_tables(set(detected_sources))

                for kind in self.IMPORT_ORDER:
                    source = detected_sources.get(kind)
                    if not source:
                        continue
                    model = self.KIND_TO_MODEL[kind]
                    imported_models.append(model)
                    counts_by_kind[kind] = self._import_kind(kind, source, warnings)

                self._reset_sqlite_sequences(imported_models)

                if dry_run:
                    transaction.set_rollback(True)
        except CSVImportLineError as exc:
            raise CommandError(str(exc)) from exc
        except Exception as exc:
            raise CommandError(f"Falha na restauracao: {exc}") from exc

        self.stdout.write("")
        self.stdout.write(self.style.SUCCESS("Resumo da importacao:"))
        for kind in self.IMPORT_ORDER:
            count = counts_by_kind.get(kind)
            if not count:
                continue
            self.stdout.write(
                f"- {kind}: {count.created} criados, {count.updated} atualizados, {count.skipped} pulados"
            )

        if warnings:
            self.stdout.write("")
            self.stdout.write(self.style.WARNING(f"Warnings ({len(warnings)}):"))
            for warning in warnings:
                self.stdout.write(f"- {warning}")

        if dry_run:
            self.stdout.write(self.style.WARNING("DRY-RUN finalizado com rollback."))
        else:
            self.stdout.write(self.style.SUCCESS("Restauracao concluida com sucesso."))

    def _resolve_source_dir(self, raw_source_dir: str) -> Path:
        raw_path = Path(raw_source_dir)
        candidates: list[Path] = []

        if raw_path.is_absolute():
            candidates.append(raw_path)
        else:
            candidates.append((Path.cwd() / raw_path).resolve())
            candidates.append((Path(settings.BASE_DIR) / raw_path).resolve())
            candidates.append((Path(settings.BASE_DIR) / "viagens" / raw_path).resolve())

        seen: set[Path] = set()
        for candidate in candidates:
            if candidate in seen:
                continue
            seen.add(candidate)
            if candidate.exists() and candidate.is_dir():
                return candidate

        tried = ", ".join(str(path) for path in candidates)
        raise CommandError(f"Diretorio nao encontrado: {raw_source_dir}. Caminhos testados: {tried}")

    def _load_and_detect_sources(
        self,
        *,
        csv_paths: list[Path],
        preferred_encoding: str,
        delimiter: str,
    ) -> dict[str, CsvSource]:
        by_kind: dict[str, CsvSource] = {}

        for csv_path in csv_paths:
            source = read_csv_source(
                csv_path,
                preferred_encoding=preferred_encoding,
                delimiter=delimiter,
            )
            kind = detect_dataset_kind(csv_path, source.header_keys)
            if not kind:
                continue

            if kind in by_kind:
                existing = by_kind[kind]
                raise CommandError(
                    f"Dois arquivos mapeados para '{kind}': "
                    f"{existing.path.name} e {csv_path.name}."
                )

            by_kind[kind] = source

        return by_kind

    def _truncate_target_tables(self, detected_kinds: set[str]) -> None:
        truncate_models: set[type[models.Model]] = set()
        for kind in detected_kinds:
            model = self.KIND_TO_MODEL.get(kind)
            if model:
                truncate_models.add(model)

        if {"oficio", "trecho", "oficio_viajante"} & detected_kinds:
            truncate_models.update({Oficio, Trecho, Oficio.viajantes.through})
        if {"estado", "cidade"} & detected_kinds:
            truncate_models.update({Estado, Cidade})
        if "viajante" in detected_kinds:
            truncate_models.update({Viajante, Cargo})

        delete_order = [
            Oficio.viajantes.through,
            Trecho,
            Oficio,
            Veiculo,
            Viajante,
            Cargo,
            Cidade,
            Estado,
        ]
        for model in delete_order:
            if model in truncate_models:
                model.objects.all().delete()

    def _import_kind(self, kind: str, source: CsvSource, warnings: list[str]) -> ImportCount:
        if kind == "estado":
            return self._import_estados(source, warnings)
        if kind == "cidade":
            return self._import_cidades(source, warnings)
        if kind == "cargo":
            return self._import_cargos(source, warnings)
        if kind == "viajante":
            return self._import_viajantes(source, warnings)
        if kind == "oficio_viajante":
            return self._import_oficio_viajantes(source, warnings)
        if kind == "veiculo":
            return self._import_generic(source, Veiculo, warnings)
        if kind == "oficio":
            return self._import_generic(source, Oficio, warnings)
        if kind == "trecho":
            return self._import_generic(source, Trecho, warnings)
        raise CommandError(f"Tipo de importacao nao suportado: {kind}")

    def _import_estados(self, source: CsvSource, warnings: list[str]) -> ImportCount:
        count = ImportCount()
        for row in source.rows:
            if is_blank_row(row.data):
                count.skipped += 1
                warnings.append(f"{source.path.name}:{row.number} linha vazia pulada.")
                continue

            nome = get_first_value(row.data, ["nome", "estado"])
            sigla = get_first_value(row.data, ["sigla", "uf"]).upper()
            raw_id = get_first_value(row.data, ["id"])

            if not nome:
                raise self._line_error(source, row, "nome", "Informe o nome do estado.")
            if not sigla:
                raise self._line_error(source, row, "sigla", "Informe a sigla do estado.")

            defaults = {"nome": nome, "sigla": sigla}
            pk_value = self._parse_int_optional(raw_id, source, row, "id")

            if pk_value is not None:
                _, created = Estado.objects.update_or_create(pk=pk_value, defaults=defaults)
            else:
                _, created = Estado.objects.update_or_create(sigla=sigla, defaults={"nome": nome})

            if created:
                count.created += 1
            else:
                count.updated += 1
        return count

    def _import_cidades(self, source: CsvSource, warnings: list[str]) -> ImportCount:
        count = ImportCount()
        for row in source.rows:
            if is_blank_row(row.data):
                count.skipped += 1
                warnings.append(f"{source.path.name}:{row.number} linha vazia pulada.")
                continue

            nome = get_first_value(row.data, ["nome", "municipio", "cidade"])
            raw_id = get_first_value(row.data, ["id"])
            raw_estado_id = get_first_value(row.data, ["estado_id"])
            raw_uf = get_first_value(row.data, ["uf", "sigla", "estado_sigla", "estado"])

            if not nome:
                raise self._line_error(source, row, "nome", "Informe o nome da cidade.")

            estado = None
            if not is_effectively_empty(raw_estado_id):
                estado_id = self._parse_int_required(raw_estado_id, source, row, "estado_id")
                estado = Estado.objects.filter(pk=estado_id).first()
            elif not is_effectively_empty(raw_uf):
                uf = raw_uf.strip().upper()
                if len(uf) == 2:
                    estado = Estado.objects.filter(sigla__iexact=uf).first()
                if estado is None:
                    estado = Estado.objects.filter(nome__iexact=raw_uf.strip()).first()

            if estado is None:
                raise self._line_error(
                    source,
                    row,
                    "estado_id/uf",
                    f"Nao foi possivel resolver o estado para cidade '{nome}'.",
                )

            defaults = {"nome": nome, "estado": estado}
            pk_value = self._parse_int_optional(raw_id, source, row, "id")
            if pk_value is not None:
                _, created = Cidade.objects.update_or_create(pk=pk_value, defaults=defaults)
            else:
                _, created = Cidade.objects.update_or_create(nome=nome, estado=estado, defaults={})

            if created:
                count.created += 1
            else:
                count.updated += 1
        return count

    def _import_cargos(self, source: CsvSource, warnings: list[str]) -> ImportCount:
        count = ImportCount()
        for row in source.rows:
            if is_blank_row(row.data):
                count.skipped += 1
                warnings.append(f"{source.path.name}:{row.number} linha vazia pulada.")
                continue

            nome = get_first_value(row.data, ["nome", "cargo"])
            raw_id = get_first_value(row.data, ["id"])
            if not nome:
                raise self._line_error(source, row, "nome", "Informe o nome do cargo.")

            pk_value = self._parse_int_optional(raw_id, source, row, "id")
            if pk_value is not None:
                _, created = Cargo.objects.update_or_create(pk=pk_value, defaults={"nome": nome})
            else:
                _, created = Cargo.objects.update_or_create(nome=nome, defaults={})

            if created:
                count.created += 1
            else:
                count.updated += 1
        return count

    def _import_viajantes(self, source: CsvSource, warnings: list[str]) -> ImportCount:
        count = ImportCount()
        cargo_cache = {
            normalize_key(nome): nome
            for nome in Cargo.objects.values_list("nome", flat=True)
            if nome
        }

        for row in source.rows:
            if is_blank_row(row.data):
                count.skipped += 1
                warnings.append(f"{source.path.name}:{row.number} linha vazia pulada.")
                continue

            raw_id = get_first_value(row.data, ["id"])
            nome = get_first_value(
                row.data,
                ["nome", "servidor_nome_completo", "servidor_nome"],
            )
            rg = get_first_value(row.data, ["rg"])
            cpf = get_first_value(row.data, ["cpf"])
            cargo = get_first_value(row.data, ["cargo"])
            telefone = get_first_value(row.data, ["telefone", "fone", "celular"])

            if not nome:
                raise self._line_error(source, row, "nome", "Informe o nome do viajante.")

            cargo = cargo.strip()
            if cargo:
                cargo_key = normalize_key(cargo)
                if cargo_key not in cargo_cache:
                    Cargo.objects.create(nome=cargo)
                    cargo_cache[cargo_key] = cargo

            defaults = {
                "nome": nome,
                "rg": rg or "",
                "cpf": cpf or "",
                "cargo": cargo or "",
                "telefone": telefone or "",
            }

            pk_value = self._parse_int_optional(raw_id, source, row, "id")
            if pk_value is not None:
                _, created = Viajante.objects.update_or_create(pk=pk_value, defaults=defaults)
            else:
                lookup = {"cpf": cpf} if cpf else {"nome": nome}
                _, created = Viajante.objects.update_or_create(**lookup, defaults=defaults)

            if created:
                count.created += 1
            else:
                count.updated += 1

        return count

    def _import_oficio_viajantes(self, source: CsvSource, warnings: list[str]) -> ImportCount:
        count = ImportCount()
        through_model = Oficio.viajantes.through

        for row in source.rows:
            if is_blank_row(row.data):
                count.skipped += 1
                warnings.append(f"{source.path.name}:{row.number} linha vazia pulada.")
                continue

            raw_id = get_first_value(row.data, ["id"])
            raw_oficio_id = get_first_value(row.data, ["oficio_id"])
            raw_viajante_id = get_first_value(row.data, ["viajante_id"])

            if is_effectively_empty(raw_oficio_id):
                raise self._line_error(source, row, "oficio_id", "Informe oficio_id.")
            if is_effectively_empty(raw_viajante_id):
                raise self._line_error(source, row, "viajante_id", "Informe viajante_id.")

            oficio_id = self._parse_int_required(raw_oficio_id, source, row, "oficio_id")
            viajante_id = self._parse_int_required(raw_viajante_id, source, row, "viajante_id")

            if not Oficio.objects.filter(pk=oficio_id).exists():
                raise self._line_error(
                    source,
                    row,
                    "oficio_id",
                    f"Oficio {oficio_id} nao existe.",
                )
            if not Viajante.objects.filter(pk=viajante_id).exists():
                raise self._line_error(
                    source,
                    row,
                    "viajante_id",
                    f"Viajante {viajante_id} nao existe.",
                )

            pk_value = self._parse_int_optional(raw_id, source, row, "id")
            defaults = {"oficio_id": oficio_id, "viajante_id": viajante_id}
            if pk_value is not None:
                _, created = through_model.objects.update_or_create(pk=pk_value, defaults=defaults)
            else:
                _, created = through_model.objects.update_or_create(**defaults, defaults={})

            if created:
                count.created += 1
            else:
                count.updated += 1

        return count

    def _import_generic(
        self,
        source: CsvSource,
        model: type[models.Model],
        warnings: list[str],
    ) -> ImportCount:
        count = ImportCount()
        concrete_fields = list(model._meta.concrete_fields)

        for row in source.rows:
            if is_blank_row(row.data):
                count.skipped += 1
                warnings.append(f"{source.path.name}:{row.number} linha vazia pulada.")
                continue

            raw_id = row.data.get("id", "")
            pk_value = self._parse_int_optional(raw_id, source, row, "id")
            payload: dict[str, object] = {}

            for field in concrete_fields:
                if field.primary_key:
                    continue

                attname_key = normalize_key(field.attname)
                field_name_key = normalize_key(field.name)
                has_attname = attname_key in row.data
                has_name = field_name_key in row.data
                raw_value = ""

                if isinstance(field, models.ForeignKey):
                    if has_attname:
                        if not is_effectively_empty(row.data.get(attname_key)):
                            raw_value = row.data.get(attname_key, "")
                            payload[field.attname] = self._resolve_fk_id(
                                field, raw_value, source, row
                            )
                        elif field.null:
                            payload[field.attname] = None
                        else:
                            raise self._line_error(
                                source,
                                row,
                                field.attname,
                                "Campo obrigatorio vazio.",
                            )
                    elif has_name:
                        if not is_effectively_empty(row.data.get(field_name_key)):
                            raw_value = row.data.get(field_name_key, "")
                            payload[field.attname] = self._resolve_fk_label(
                                field, raw_value, row.data, source, row
                            )
                        elif field.null:
                            payload[field.attname] = None
                        else:
                            raise self._line_error(
                                source,
                                row,
                                field.name,
                                "Campo obrigatorio vazio.",
                            )
                    continue

                if has_name:
                    raw_value = row.data.get(field_name_key, "")
                elif has_attname:
                    raw_value = row.data.get(attname_key, "")
                else:
                    continue

                converted = self._convert_scalar_field(field, raw_value, source, row)
                if converted is None and not field.null and not isinstance(
                    field, (models.CharField, models.TextField)
                ):
                    continue
                payload[field.name] = converted

            if pk_value is not None:
                _, created = model.objects.update_or_create(pk=pk_value, defaults=payload)
            else:
                lookup = self._build_lookup_for_model(model, row.data, payload)
                if lookup:
                    _, created = model.objects.update_or_create(**lookup, defaults=payload)
                else:
                    model.objects.create(**payload)
                    created = True

            if created:
                count.created += 1
            else:
                count.updated += 1

        return count

    def _build_lookup_for_model(
        self,
        model: type[models.Model],
        row_data: dict[str, str],
        payload: dict[str, object],
    ) -> dict[str, object]:
        if model is Veiculo:
            placa = payload.get("placa") or row_data.get("placa")
            if placa:
                return {"placa": placa}

        if model is Oficio:
            oficio_num = payload.get("oficio") or row_data.get("oficio")
            protocolo = payload.get("protocolo") or row_data.get("protocolo")
            if oficio_num and protocolo:
                return {"oficio": oficio_num, "protocolo": protocolo}
            if oficio_num:
                return {"oficio": oficio_num}

        if model is Trecho:
            oficio_id = payload.get("oficio_id")
            ordem = payload.get("ordem")
            if oficio_id and ordem:
                return {"oficio_id": oficio_id, "ordem": ordem}

        for field in model._meta.fields:
            if field.primary_key or not field.unique:
                continue
            value = payload.get(field.name)
            if value is not None:
                return {field.name: value}
        return {}

    def _convert_scalar_field(
        self,
        field: models.Field,
        raw_value: str,
        source: CsvSource,
        row: CsvRow,
    ):
        if is_effectively_empty(raw_value):
            if field.null:
                return None
            if isinstance(field, (models.CharField, models.TextField)):
                return ""
            if field.has_default():
                return None
            raise ValueError("Campo obrigatorio vazio.")

        try:
            if isinstance(field, models.BooleanField):
                return parse_bool(raw_value)
            if isinstance(field, models.DateTimeField):
                return parse_datetime(raw_value)
            if isinstance(field, models.DateField):
                return parse_date(raw_value)
            if isinstance(field, models.TimeField):
                return parse_time(raw_value)
            if isinstance(field, models.DecimalField):
                return parse_decimal(raw_value)
            if isinstance(
                field,
                (
                    models.AutoField,
                    models.BigAutoField,
                    models.IntegerField,
                    models.BigIntegerField,
                    models.PositiveIntegerField,
                    models.PositiveSmallIntegerField,
                    models.SmallIntegerField,
                ),
            ):
                decimal_value = parse_decimal(raw_value)
                if decimal_value != int(decimal_value):
                    raise ValueError(f"Inteiro invalido: {raw_value!r}")
                return int(decimal_value)
            if isinstance(field, models.FloatField):
                return float(parse_decimal(raw_value))
        except ValueError as exc:
            raise self._line_error(source, row, field.name, str(exc)) from exc

        return raw_value.strip()

    def _resolve_fk_id(
        self,
        field: models.ForeignKey,
        raw_value: str,
        source: CsvSource,
        row: CsvRow,
    ) -> int:
        fk_id = self._parse_int_required(raw_value, source, row, field.attname)
        if not field.related_model.objects.filter(pk=fk_id).exists():
            raise self._line_error(
                source,
                row,
                field.attname,
                f"{field.related_model.__name__} com id={fk_id} nao existe.",
            )
        return fk_id

    def _resolve_fk_label(
        self,
        field: models.ForeignKey,
        raw_value: str,
        row_data: dict[str, str],
        source: CsvSource,
        row: CsvRow,
    ) -> int:
        related_model = field.related_model
        value = raw_value.strip()

        if related_model is Estado:
            estado = Estado.objects.filter(sigla__iexact=value).first()
            if estado is None:
                estado = Estado.objects.filter(nome__iexact=value).first()
            if estado is None:
                raise self._line_error(
                    source,
                    row,
                    field.name,
                    f"Estado '{value}' nao encontrado.",
                )
            return int(estado.pk)

        if related_model is Cidade:
            uf_value = self._extract_city_uf(field.name, row_data)
            query = Cidade.objects.filter(nome__iexact=value)
            if uf_value:
                query = query.filter(estado__sigla__iexact=uf_value)
            matches = list(query[:2])
            if not matches:
                hint = f" e UF='{uf_value}'" if uf_value else ""
                raise self._line_error(
                    source,
                    row,
                    field.name,
                    f"Cidade '{value}'{hint} nao encontrada.",
                )
            if len(matches) > 1:
                raise self._line_error(
                    source,
                    row,
                    field.name,
                    f"Cidade '{value}' ambigua. Informe {field.attname}.",
                )
            return int(matches[0].pk)

        if related_model is Viajante:
            token = "".join(ch for ch in value if ch.isdigit())
            if len(token) >= 11:
                viajante = Viajante.objects.filter(cpf=value).first()
            else:
                viajante = Viajante.objects.filter(nome__iexact=value).first()
            if viajante is None:
                raise self._line_error(
                    source,
                    row,
                    field.name,
                    f"Viajante '{value}' nao encontrado.",
                )
            return int(viajante.pk)

        if related_model is Veiculo:
            veiculo = Veiculo.objects.filter(placa__iexact=value).first()
            if veiculo is None:
                raise self._line_error(
                    source,
                    row,
                    field.name,
                    f"Veiculo '{value}' nao encontrado.",
                )
            return int(veiculo.pk)

        if related_model is Oficio:
            if value.isdigit():
                oficio = Oficio.objects.filter(pk=int(value)).first()
            else:
                oficio = Oficio.objects.filter(oficio=value).first()
            if oficio is None:
                raise self._line_error(
                    source,
                    row,
                    field.name,
                    f"Oficio '{value}' nao encontrado.",
                )
            return int(oficio.pk)

        if value.isdigit() and related_model.objects.filter(pk=int(value)).exists():
            return int(value)

        if any(field_.name == "nome" for field_ in related_model._meta.fields):
            obj = related_model.objects.filter(nome__iexact=value).first()
            if obj:
                return int(obj.pk)

        raise self._line_error(
            source,
            row,
            field.name,
            f"Nao foi possivel resolver FK para {related_model.__name__} com valor '{value}'.",
        )

    def _extract_city_uf(self, field_name: str, row_data: dict[str, str]) -> str:
        candidates = ["uf", "estado_sigla", "sigla"]

        prefixes: list[str] = []
        if field_name.endswith("_cidade"):
            prefixes.append(field_name[: -len("_cidade")])
        if field_name.startswith("cidade_"):
            prefixes.append(field_name[len("cidade_") :])
        if field_name == "cidade":
            prefixes.append("")

        for prefix in prefixes:
            if not prefix:
                continue
            candidates.extend(
                [
                    f"{prefix}_uf",
                    f"{prefix}_estado",
                    f"{prefix}_estado_sigla",
                ]
            )

        for candidate in candidates:
            key = normalize_key(candidate)
            value = (row_data.get(key) or "").strip()
            if not value:
                continue
            if len(value) == 2:
                return value.upper()
            estado = Estado.objects.filter(nome__iexact=value).first()
            if estado:
                return estado.sigla.upper()
        return ""

    def _parse_int_optional(
        self,
        raw_value: str,
        source: CsvSource,
        row: CsvRow,
        column: str,
    ) -> int | None:
        if is_effectively_empty(raw_value):
            return None
        return self._parse_int_required(raw_value, source, row, column)

    def _parse_int_required(
        self,
        raw_value: str,
        source: CsvSource,
        row: CsvRow,
        column: str,
    ) -> int:
        try:
            decimal_value = parse_decimal(raw_value)
        except ValueError as exc:
            raise self._line_error(source, row, column, str(exc)) from exc

        if decimal_value != int(decimal_value):
            raise self._line_error(
                source,
                row,
                column,
                f"Valor inteiro invalido: {raw_value!r}",
            )
        return int(decimal_value)

    def _line_error(
        self,
        source: CsvSource,
        row: CsvRow,
        column: str,
        message: str,
    ) -> CSVImportLineError:
        return CSVImportLineError(
            message,
            file_path=source.path,
            row_number=row.number,
            column=column,
        )

    def _reset_sqlite_sequences(self, imported_models: list[type[models.Model]]) -> None:
        if connection.vendor != "sqlite":
            return

        distinct_models: list[type[models.Model]] = []
        seen: set[type[models.Model]] = set()
        for model in imported_models:
            if model in seen:
                continue
            seen.add(model)
            distinct_models.append(model)

        with connection.cursor() as cursor:
            cursor.execute(
                "SELECT name FROM sqlite_master WHERE type='table' AND name='sqlite_sequence'"
            )
            if not cursor.fetchone():
                return

            for model in distinct_models:
                pk_field = model._meta.pk
                if not isinstance(pk_field, (models.AutoField, models.BigAutoField)):
                    continue

                table = model._meta.db_table
                pk_column = pk_field.column
                cursor.execute(
                    f'SELECT COALESCE(MAX("{pk_column}"), 0) FROM "{table}"'
                )
                max_id = int(cursor.fetchone()[0] or 0)

                cursor.execute("SELECT 1 FROM sqlite_sequence WHERE name = %s", [table])
                if cursor.fetchone():
                    cursor.execute(
                        "UPDATE sqlite_sequence SET seq = %s WHERE name = %s",
                        [max_id, table],
                    )
                else:
                    cursor.execute(
                        "INSERT INTO sqlite_sequence(name, seq) VALUES (%s, %s)",
                        [table, max_id],
                    )
