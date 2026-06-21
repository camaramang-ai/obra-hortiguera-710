import fitz
import os

path = '/Users/PabloMan/Hermes/projects/active/obra-hortiguera-710/recursos/Contexto Completo para Agente de Compras — Obra Hortiguera 710.pdf'
outdir = '/Users/PabloMan/Hermes/projects/active/obra-hortiguera-710/presupuestos'

doc = fitz.open(path)
print(f'Pages: {doc.page_count}')

for page_num in range(doc.page_count):
    page = doc[page_num]
    # Try to get text first
    text = page.get_text()
    if text.strip():
        print(f'\n--- TEXT PAGE {page_num+1} ---')
        print(text[:500])
    
    # Extract images
    images = page.get_images(full=True)
    print(f'\nPage {page_num+1}: {len(images)} embedded image(s)')
    
    for img_idx, img in enumerate(images):
        xref = img[0]
        base_image = doc.extract_image(xref)
        image_bytes = base_image["image"]
        ext = base_image["ext"]
        fname = f'page{page_num+1}_img{img_idx+1}.{ext}'
        fpath = os.path.join(outdir, fname)
        with open(fpath, 'wb') as f:
            f.write(image_bytes)
        print(f'  Saved: {fname} ({len(image_bytes)} bytes, {ext})')

doc.close()
print(f'\nDone. Images saved to {outdir}')
