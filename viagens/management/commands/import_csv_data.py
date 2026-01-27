import csv
from pathlib import Path

from django.core.management.base import BaseCommand

from viagens.models import Cidade, Estado, Viajante


class Command(BaseCommand):
    help = "Importa estados, cidades e viajantes a partir de arquivos CSV."

    def add_arguments(self, parser):
        parser.add_argument(
            "--estados",
            default="viagens/data/estados.csv",
            help="Caminho para estados.csv",
        )
        parser.add_argument(
            "--cidades",
            default="viagens/data/cidades.csv",
            help="Caminho para cidades.csv",
        )
        parser.add_argument(
            "--viajantes",
            default="viagens/data/viajantes.csv",
            help="Caminho para viajantes.csv",
        )
        parser.add_argument(
            "--reset",
            action="store_true",
            help="Apaga estados e cidades antes de importar.",
        )

    def _open_csv(self, path: Path):
        handle = None
        try:
            handle = path.open(mode="r", encoding="utf-8-sig", newline="")
            handle.read(2048)
            handle.seek(0)
            return handle
        except UnicodeDecodeError:
            if handle:
                handle.close()
            return path.open(mode="r", encoding="latin-1", newline="")

    def handle(self, *args, **options):
        estados_path = Path(options["estados"]).resolve()
        cidades_path = Path(options["cidades"]).resolve()
        viajantes_path = Path(options["viajantes"]).resolve()
        reset = options.get("reset", False)

        if reset:
            Cidade.objects.all().delete()
            Estado.objects.all().delete()
            self.stdout.write("Estados e cidades removidos antes da importacao.")

        if not estados_path.exists():
            self.stderr.write(f"Arquivo de estados nao encontrado: {estados_path}")
        else:
            self._import_estados(estados_path)

        if not cidades_path.exists():
            self.stderr.write(f"Arquivo de cidades nao encontrado: {cidades_path}")
        else:
            self._import_cidades(cidades_path)

        if not viajantes_path.exists():
            self.stderr.write(f"Arquivo de viajantes nao encontrado: {viajantes_path}")
        else:
            self._import_viajantes(viajantes_path)

    def _import_estados(self, path: Path):
        created = 0
        updated = 0
        with self._open_csv(path) as handle:
            reader = csv.DictReader(handle)
            for row in reader:
                nome = self._get_value(row, ["nome", "estado"])
                sigla = self._get_value(row, ["sigla", "uf"]).upper()
                if not nome or not sigla:
                    continue
                obj, was_created = Estado.objects.update_or_create(
                    sigla=sigla, defaults={"nome": nome}
                )
                if was_created:
                    created += 1
                else:
                    updated += 1
        self.stdout.write(
            f"Estados importados: {created} criados, {updated} atualizados."
        )

    def _import_cidades(self, path: Path):
        created = 0
        skipped = 0
        with self._open_csv(path) as handle:
            reader = csv.DictReader(handle)
            fieldnames = reader.fieldnames or []
            uf_key = self._pick_field(fieldnames, ["uf", "estado", "sigla"]) or (
                fieldnames[0] if fieldnames else "UF"
            )
            cidade_key = self._pick_field(fieldnames, ["municipio", "cidade"]) or (
                fieldnames[1] if len(fieldnames) > 1 else None
            )
            if not cidade_key:
                self.stderr.write("Nao foi possivel identificar a coluna de cidades.")
                return

            for row in reader:
                uf = str(row.get(uf_key, "")).strip().upper()
                nome = str(row.get(cidade_key, "")).strip()
                if not uf or not nome:
                    continue
                estado = Estado.objects.filter(sigla=uf).first()
                if not estado:
                    skipped += 1
                    continue
                _, was_created = Cidade.objects.get_or_create(
                    nome=nome, estado=estado
                )
                if was_created:
                    created += 1
        self.stdout.write(
            f"Cidades importadas: {created} criadas, {skipped} ignoradas (estado inexistente)."
        )

    def _import_viajantes(self, path: Path):
        created = 0
        updated = 0
        with self._open_csv(path) as handle:
            reader = csv.DictReader(handle)
            for row in reader:
                nome = self._get_value(row, ["nome", "servidor (nome completo)"])
                rg = self._get_value(row, ["rg"]) 
                cpf = self._get_value(row, ["cpf"]) 
                cargo = self._get_value(row, ["cargo"]) 
                telefone = self._get_value(row, ["telefone", "fone", "celular"]) 

                if not nome:
                    continue

                lookup = {"cpf": cpf} if cpf else {"nome": nome}
                obj, was_created = Viajante.objects.update_or_create(
                    **lookup,
                    defaults={
                        "nome": nome,
                        "rg": rg,
                        "cpf": cpf or "",
                        "cargo": cargo or "",
                        "telefone": telefone or "",
                    },
                )
                if was_created:
                    created += 1
                else:
                    updated += 1
        self.stdout.write(
            f"Viajantes importados: {created} criados, {updated} atualizados."
        )

    def _get_value(self, row: dict, keys: list[str]) -> str:
        for key in keys:
            for row_key, value in row.items():
                if row_key.strip().lower() == key.strip().lower():
                    return str(value).strip()
        return ""

    def _pick_field(self, fieldnames: list[str], keys: list[str]) -> str | None:
        for name in fieldnames:
            normalized = name.strip().lower()
            for key in keys:
                if key in normalized:
                    return name
        return None
