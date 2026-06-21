import fitz
path = '/Users/PabloMan/Hermes/projects/active/obra-hortiguera-710/recursos/Contexto Completo para Agente de Compras — Obra Hortiguera 710.pdf'
doc = fitz.open(path)
print(f'Total pages: {doc.page_count}')
print('=' * 60)
for i, page in enumerate(doc):
    text = page.get_text()
    print(f'\n--- PAGE {i+1} ---\n')
    print(text)
doc.close()
