import os
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT


def add_code_to_document(doc, file_path, code):
    doc.add_heading(file_path, level=2)
    p = doc.add_paragraph()
    p.add_run(code).font.size = Pt(12)
    p.line_spacing_rule = 1.15
    p.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT


def read_code_files(directory):
    code_files = []
    for root, _, files in os.walk(directory):
        for file in files:
            if file.endswith('.java'):  # you can add other extensions if needed
                file_path = os.path.join(root, file)
                with open(file_path, 'r', encoding='utf-8') as f:
                    code = f.read()
                relative_path = os.path.relpath(file_path, directory)
                code_files.append((relative_path, code))
    return code_files


def create_word_documents(code_files, output_prefix):
    output_dir = 'output'
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    doc = Document()
    current_file_index = 1
    current_page_count = 0

    for file_path, code in code_files:
        add_code_to_document(doc, file_path, code)
        current_page_count += doc.element.xpath('count(//w:sectPr)')

        if current_page_count >= 24:
            output_file = os.path.join(output_dir, f"{output_prefix}_{current_file_index}.docx")
            doc.save(output_file)
            print(f"Saved {output_file}")
            doc = Document()  # Start a new document
            current_file_index += 1
            current_page_count = 0

    # Save remaining pages, if any
    if current_page_count > 0:
        output_file = os.path.join(output_dir, f"{output_prefix}_{current_file_index}.docx")
        doc.save(output_file)
        print(f"Saved {output_file}")


def main():
    input_directory = 'input/app-frw-ng-plus'  # replace with your path
    output_prefix = 'app_code'
    code_files = read_code_files(input_directory)
    create_word_documents(code_files, output_prefix)


if __name__ == '__main__':
    main()
