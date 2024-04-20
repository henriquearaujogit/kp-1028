from docx import Document

def substituir_d01_por_xxxx(docx_file):
    doc = Document(docx_file)
    
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                if '{{h01d01}}' in cell.text:
                    cell.text = cell.text.replace('{{h01d01}}', 'xxxxxxx')

    doc.save('arquivo_modificado.docx')

if __name__ == "__main__":
    substituir_d01_por_xxxx('modelo.docx')
