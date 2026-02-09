from pathlib import Path
text = Path('viagens/documents/document.py').read_text(encoding='utf-8').splitlines()
start = next(i for i,line in enumerate(text) if 'tipo_custeio' in line)
for idx in range(start, start+40):
    print(f"{idx+1:04d}: {text[idx]}")
