from pathlib import Path
text = Path('viagens/views.py').read_text(encoding='utf-8').splitlines()
for idx in range(1790, 1875):
    print(f"{idx+1:04d}: {text[idx]}")
