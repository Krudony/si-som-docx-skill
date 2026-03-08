import sys; from docx import Document;
def analyze(path):
    try:
        doc = Document(path); print(f'Analysis: {path}'); section = doc.sections[0]; print(f'Margins: {section.top_margin}, {section.bottom_margin}, {section.left_margin}, {section.right_margin}')
        for i, p in enumerate(doc.paragraphs):
            if p.text.strip(): print(f'P{i}: {p.text[:50]}... Align={p.alignment}')
        for i, t in enumerate(doc.tables):
            print(f'Table {i}');
            for r, row in enumerate(t.rows):
                cells = [c.text.strip() for c in row.cells]; print(f'  R{r}: {cells}')
    except Exception as e: print(f'Error: {e}')
if __name__ == '__main__':
    analyze(sys.argv[1])
