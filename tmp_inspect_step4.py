from pathlib import Path
text = Path('viagens/templates/viagens/oficio_step4.html').read_text().splitlines()
for idx,line in enumerate(text):
    if idx >= 60 and idx <= 90:
        print(f"{idx+1:04d}: {line}")
