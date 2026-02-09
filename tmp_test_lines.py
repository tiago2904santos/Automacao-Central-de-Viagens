from pathlib import Path
text = Path('viagens/tests/test_roteiro.py').read_text(encoding='utf-8').splitlines()
for idx in range(40, 140):
    print(f"{idx+1:04d}: {text[idx]}")
