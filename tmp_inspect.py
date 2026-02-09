from zipfile import ZipFile
from pathlib import Path
path = Path('viagens/documents/oficio_model.docx')
with ZipFile(path) as z:
    data = z.read('word/document.xml').decode('utf-8')
print('custo' in data)
print(data.count('{{'))
for placeholder in ['{{custo', '{{destino', '{{orgao_destino']:
    print(placeholder, placeholder in data)
