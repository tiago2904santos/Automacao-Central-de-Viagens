import csv
import io
import re
import unicodedata
from dataclasses import dataclass
from datetime import date, datetime, time
from decimal import Decimal, InvalidOperation
from pathlib import Path


class CSVImportLineError(ValueError):
    def __init__(self, message: str, *, file_path: Path, row_number: int, column: str):
        super().__init__(message)
        self.file_path = file_path
        self.row_number = row_number
        self.column = column

    def __str__(self) -> str:
        return f"{self.file_path.name}:{self.row_number}:{self.column} - {super().__str__()}"


@dataclass(frozen=True)
class CsvRow:
    number: int
    data: dict[str, str]


@dataclass(frozen=True)
class CsvSource:
    path: Path
    encoding: str
    delimiter: str
    headers: list[str]
    header_keys: set[str]
    rows: list[CsvRow]


def normalize_key(value: str) -> str:
    raw = unicodedata.normalize("NFKD", (value or "").strip())
    raw = "".join(ch for ch in raw if not unicodedata.combining(ch))
    raw = raw.casefold()
    raw = re.sub(r"[^a-z0-9]+", "_", raw)
    return raw.strip("_")


def normalize_token(value: str) -> str:
    return normalize_key(value).replace("_", "")


def unique_preserve_order(values: list[str]) -> list[str]:
    seen: set[str] = set()
    ordered: list[str] = []
    for value in values:
        if value in seen:
            continue
        seen.add(value)
        ordered.append(value)
    return ordered


def _parse_rows(decoded_text: str, delimiter: str) -> tuple[list[str], list[CsvRow]]:
    reader = csv.DictReader(io.StringIO(decoded_text), delimiter=delimiter)
    headers = reader.fieldnames or []
    rows: list[CsvRow] = []
    line_number = 2
    for row in reader:
        normalized: dict[str, str] = {}
        for key, value in (row or {}).items():
            if key is None:
                continue
            normalized[normalize_key(key)] = (value or "").strip()
        rows.append(CsvRow(number=line_number, data=normalized))
        line_number += 1
    return headers, rows


def _choose_delimiter(decoded_text: str, delimiter: str) -> str:
    if delimiter in {";", ","}:
        return delimiter

    sample = decoded_text[:4096]
    try:
        sniffed = csv.Sniffer().sniff(sample, delimiters=";,")
        if sniffed.delimiter in {";", ","}:
            return sniffed.delimiter
    except csv.Error:
        pass

    first_line = decoded_text.splitlines()[0] if decoded_text.splitlines() else ""
    semicolon_count = first_line.count(";")
    comma_count = first_line.count(",")
    return ";" if semicolon_count >= comma_count else ","


def read_csv_source(path: Path, *, preferred_encoding: str = "utf-8-sig", delimiter: str = "auto") -> CsvSource:
    if not path.exists():
        raise FileNotFoundError(path)

    raw_bytes = path.read_bytes()
    encoding_candidates = unique_preserve_order(
        [preferred_encoding or "utf-8-sig", "utf-8-sig", "utf-8", "latin-1"]
    )

    decode_error: Exception | None = None
    for encoding in encoding_candidates:
        try:
            decoded = raw_bytes.decode(encoding)
        except UnicodeDecodeError as exc:
            decode_error = exc
            continue

        selected_delimiter = _choose_delimiter(decoded, delimiter)
        headers, rows = _parse_rows(decoded, selected_delimiter)

        if delimiter == "auto" and len(headers) <= 1:
            fallback = "," if selected_delimiter == ";" else ";"
            alt_headers, alt_rows = _parse_rows(decoded, fallback)
            if len(alt_headers) > len(headers):
                headers, rows = alt_headers, alt_rows
                selected_delimiter = fallback

        return CsvSource(
            path=path,
            encoding=encoding,
            delimiter=selected_delimiter,
            headers=headers,
            header_keys={normalize_key(header) for header in headers},
            rows=rows,
        )

    raise UnicodeDecodeError(
        decode_error.encoding if isinstance(decode_error, UnicodeDecodeError) else "utf-8",
        decode_error.object if isinstance(decode_error, UnicodeDecodeError) else b"",
        decode_error.start if isinstance(decode_error, UnicodeDecodeError) else 0,
        decode_error.end if isinstance(decode_error, UnicodeDecodeError) else 1,
        decode_error.reason if isinstance(decode_error, UnicodeDecodeError) else "Nao foi possivel decodificar o CSV.",
    )


def detect_dataset_kind(path: Path, header_keys: set[str]) -> str | None:
    stem = normalize_key(path.stem)

    if "estado" in stem:
        return "estado"
    if "cidade" in stem or "municipio" in stem:
        return "cidade"
    if "viajante" in stem or "servidor" in stem:
        return "viajante"
    if "veiculo" in stem or "viatura" in stem:
        return "veiculo"
    if "trecho" in stem:
        return "trecho"
    if "oficio_viajante" in stem or "oficio_viajantes" in stem:
        return "oficio_viajante"
    if "oficio" in stem:
        return "oficio"
    if "cargo" in stem:
        return "cargo"

    if {"sigla", "nome"}.issubset(header_keys) and len(header_keys) <= 4:
        return "estado"
    if ("uf" in header_keys or "estado_sigla" in header_keys) and (
        {"municipio"} & header_keys or {"cidade"} & header_keys or {"nome"} & header_keys
    ):
        return "cidade"
    if {"nome", "cpf"} & header_keys and {"rg", "cargo"} & header_keys:
        return "viajante"
    if {"placa", "modelo"} <= header_keys:
        return "veiculo"
    if {"oficio_id", "ordem"} <= header_keys or (
        "oficio_id" in header_keys and ("origem_cidade" in header_keys or "destino_cidade" in header_keys)
    ):
        return "trecho"
    if {"oficio_id", "viajante_id"} <= header_keys:
        return "oficio_viajante"
    if {"protocolo", "assunto"} & header_keys and ("oficio" in header_keys or "oficio_id" in header_keys):
        return "oficio"

    return None


def get_first_value(row_data: dict[str, str], candidates: list[str]) -> str:
    for key in candidates:
        normalized = normalize_key(key)
        if normalized in row_data:
            return (row_data.get(normalized) or "").strip()
    return ""


def is_effectively_empty(value: str | None) -> bool:
    if value is None:
        return True
    raw = value.strip()
    if raw == "":
        return True
    token = normalize_token(raw)
    return token in {"none", "null", "nulo", "nil", "nan"}


def is_blank_row(row_data: dict[str, str]) -> bool:
    return all(is_effectively_empty(value) for value in row_data.values())


def parse_bool(value: str) -> bool:
    token = normalize_token(value)
    if token in {"1", "true", "t", "sim", "s", "yes", "y"}:
        return True
    if token in {"0", "false", "f", "nao", "n", "no"}:
        return False
    raise ValueError(f"Booleano invalido: {value!r}")


def parse_decimal(value: str) -> Decimal:
    raw = value.strip().replace(" ", "")
    if raw == "":
        raise ValueError("Decimal vazio.")

    if "," in raw and "." in raw:
        if raw.rfind(",") > raw.rfind("."):
            raw = raw.replace(".", "").replace(",", ".")
        else:
            raw = raw.replace(",", "")
    elif "," in raw:
        raw = raw.replace(".", "").replace(",", ".")

    try:
        return Decimal(raw)
    except InvalidOperation as exc:
        raise ValueError(f"Decimal invalido: {value!r}") from exc


def parse_date(value: str) -> date:
    raw = value.strip()
    formats = ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y")
    for fmt in formats:
        try:
            return datetime.strptime(raw, fmt).date()
        except ValueError:
            continue
    try:
        return parse_datetime(raw).date()
    except ValueError as exc:
        raise ValueError(f"Data invalida: {value!r}") from exc


def parse_datetime(value: str) -> datetime:
    raw = value.strip()
    formats = (
        "%d/%m/%Y %H:%M",
        "%d/%m/%Y %H:%M:%S",
        "%Y-%m-%d %H:%M",
        "%Y-%m-%d %H:%M:%S",
        "%d-%m-%Y %H:%M",
        "%d-%m-%Y %H:%M:%S",
    )
    for fmt in formats:
        try:
            return datetime.strptime(raw, fmt)
        except ValueError:
            continue
    raise ValueError(f"Data/hora invalida: {value!r}")


def parse_time(value: str) -> time:
    raw = value.strip()
    formats = ("%H:%M", "%H:%M:%S")
    for fmt in formats:
        try:
            return datetime.strptime(raw, fmt).time()
        except ValueError:
            continue
    try:
        return parse_datetime(raw).time()
    except ValueError as exc:
        raise ValueError(f"Hora invalida: {value!r}") from exc
