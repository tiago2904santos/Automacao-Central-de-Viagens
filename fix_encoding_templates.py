from pathlib import Path
from typing import Iterable

TEMPLATE_DIRS = [Path("templates"), Path("viagens") / "templates"]
TARGET_EXT = ".html"


def iter_template_files() -> Iterable[Path]:
    for base in TEMPLATE_DIRS:
        if not base.is_dir():
            continue
        for path in base.rglob(f"*{TARGET_EXT}"):
            yield path


def is_utf8(path: Path) -> bool:
    try:
        path.read_text(encoding="utf-8")
        return True
    except UnicodeDecodeError:
        return False


def convert_to_utf8(path: Path) -> bool:
    for encoding in ("cp1252", "latin-1"):
        try:
            content = path.read_text(encoding=encoding)
            path.write_text(content, encoding="utf-8")
            print(f"Converted {path} from {encoding} to UTF-8")
            return True
        except UnicodeDecodeError:
            continue
    print(f"Failed to convert {path}; unknown encoding")
    return False


if __name__ == "__main__":
    problematic = []
    for template in iter_template_files():
        print(f"Checking {template}")
        if not is_utf8(template):
            problematic.append(template)
    if not problematic:
        print("All templates are already UTF-8")
    else:
        print("\nDetected non-UTF-8 templates:")
        for item in problematic:
            print(f"  {item}")
        print()
        for item in problematic:
            convert_to_utf8(item)
