# Requerimientos: Python3, vscode (editor de codigo), pip install python-docx
# Ejecutar: python3 script.py
import os
from docx import Document
from docx.shared import RGBColor


def process_docx(path):
    doc = Document(path)

    def fix_runs(runs):
        for run in runs:
            run.font.color.rgb = RGBColor(0x00, 0x00, 0x00) # Color de la letra
            run.font.name = "Museo Sans 500" # Colocas el tipo de letra

    for para in doc.paragraphs:
        fix_runs(para.runs)

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    fix_runs(para.runs)

    for section in doc.sections:
        hdr = section.header
        for para in hdr.paragraphs:
            fix_runs(para.runs)
        for table in hdr.tables:
            for row in table.rows:
                for cell in row.cells:
                    for para in cell.paragraphs:
                        fix_runs(para.runs)

        ftr = section.footer
        for para in ftr.paragraphs:
            fix_runs(para.runs)
        for table in ftr.tables:
            for row in table.rows:
                for cell in row.cells:
                    for para in cell.paragraphs:
                        fix_runs(para.runs)

    dir_name, filename = os.path.split(path)
    base, ext = os.path.splitext(filename)
    new_filename = f"{base}_modificado{ext}"
    new_path = os.path.join(dir_name, new_filename)

    doc.save(new_path)
    print(f"Guardado: {new_filename} (original: {filename})")


def main(folder_path):
    for fn in os.listdir(folder_path):
        if fn.lower().endswith(".docx"):
            full_path = os.path.join(folder_path, fn)
            process_docx(full_path)


if __name__ == "__main__":
    carpeta = r"./files"
    main(carpeta)
