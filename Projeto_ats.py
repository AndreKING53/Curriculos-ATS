import json
from docx import Document
import formatacao

doc = Document()
formatacao.configurar_estilo(doc)
formatacao.titulo_estilo(doc)
formatacao.negrito(doc)
arquivo_cliente = input("Cliente: ")

with open(f"clientes/{arquivo_cliente}.json", "r", encoding="utf-8") as arquivo:
    dados = json.load(arquivo)

formatacao.adicionar_titulo(doc, dados)
formatacao.adicionar_contato(doc, dados)
formatacao.adicionar_formacao(doc, dados)
formatacao.adicionar_experiencias(doc, dados)
formatacao.adicionar_competencias(doc, dados)
formatacao.adicionar_idiomas(doc, dados)

doc.save(f"clientes/{arquivo_cliente}.docx")

from docx2pdf import convert

convert(f"clientes/{arquivo_cliente}.docx")