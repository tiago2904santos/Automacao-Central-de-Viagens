from zipfile import ZipFile
from pathlib import Path
path = Path('viagens/documents/oficio_model.docx')
with ZipFile(path) as z:
    data = z.read('word/document.xml').decode('utf-8')
start = data.index('{{custo')
print(data[start-100:start+400])
