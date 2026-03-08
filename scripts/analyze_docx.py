import sys; from docx import Document;
def analyze_docx(file_path):
    try:
        doc = Document(file_path); print(f'Analysis of: {file_path}'); print('=' * 30)
        print('[Paragraphs]')
        for i, para in enumerate(doc.paragraphs):
            if para.text.strip(): print(f'{i}: {para.text[:50]}... [Align: {para.alignment}, Bold: {any(run.bold for run in para.runs)}]')
        print('[Tables]')
        for i, table in enumerate(doc.tables):
            print(f'Table {i}: {len(table.rows)} rows, {len(table.columns)} columns')
            for r, row in enumerate(table.rows):
                cells = [cell.text.strip() for cell in row.cells]; print(f'  Row {r}: {cells}')
    except Exception as e: print(f'Error: {e}')
if __name__ == '__main__':
    if len(sys.argv) > 1: analyze_docx(sys.argv[1])
