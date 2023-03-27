import docx
from docx.shared import Inches, Cm
import json
import docx2pdf

with open('arquivo.json', 'r') as f:
    data = json.load(f)

doc = docx.Document("documento.docx")

def editar_texto(doc, texto_antigo, texto_novo):
    for paragraph in doc.paragraphs:
        if texto_antigo in paragraph.text:
            paragraph.text = paragraph.text.replace(texto_antigo, texto_novo)

palavras_alteraveis = []

for chave, valor in data['client'].items():
    palavra_antiga = f'x{chave}'
    palavra_nova = valor
    palavras_alteraveis.append({'palavra_antiga': palavra_antiga, 'palavra_nova': palavra_nova})

for i, item in enumerate(data['items']):
    chave = i + 1
    palavra_antiga = f'x{chave+2}'
    palavra_nova = str(item['value'])
    palavras_alteraveis.append({'palavra_antiga': palavra_antiga, 'palavra_nova': palavra_nova})

for palavra in palavras_alteraveis:
    editar_texto(doc, palavra['palavra_antiga'], palavra['palavra_nova'])
print((palavras_alteraveis))
arquivo_editado = 'andreigay_editado.docx'

doc.save(arquivo_editado)

arquivo_pdf = 'documento_editado.pdf'

docx2pdf.convert(arquivo_editado, arquivo_pdf)
