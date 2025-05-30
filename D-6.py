import os
from pdf2image import convert_from_path
import pytesseract
import spacy
from collections import Counter
import re
from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.shared import Inches

# Load French spaCy model
nlp = spacy.load("fr_core_news_sm")

# Chapter split pattern
CHAPTER_PATTERN = re.compile(r'(Chapitre\s+\d+)', re.IGNORECASE)

def split_text_into_chapters(text):
    splits = CHAPTER_PATTERN.split(text)
    chapters = []
    for i in range(1, len(splits), 2):
        header = splits[i].strip()
        body = splits[i+1].strip() if i+1 < len(splits) else ""
        chapters.append(f"{header}\n{body}")
    return chapters

def extract_words_spacy(text):
    # Normalize curly apostrophes and similar variants to straight apostrophe
    text = text.replace("’", "'").replace("‛", "'").replace("‘", "'")

    doc = nlp(text)
    words = []
    i = 0
    while i < len(doc):
        token = doc[i]
        # Combine contractions like d'eau, j'arrive, qu'on
        if token.text.endswith("'") and i + 1 < len(doc):
            combined = token.text + doc[i + 1].text
            words.append(combined.lower())
            i += 2
        else:
            if token.is_alpha:
                words.append(token.text.lower())
            i += 1
    return words

def save_docx_output(chapters_data, output_docx):
    """
    chapters_data = list of tuples: (chapter_number, [(word, count), ...])
    """
    doc = Document()
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Arial'
    font.size = Pt(11)

    for chapter_num, words_counts in chapters_data:
        doc.add_heading(f'Chapter {chapter_num} Unique Words', level=1)
        table = doc.add_table(rows=1, cols=4)
        table.style = 'Table Grid'  # Ensure table borders are visible
        table.autofit = False
        hdr_cells = table.rows[0].cells
        hdr_cells[0].text = 'S.No'
        hdr_cells[1].text = 'French Word'
        hdr_cells[2].text = 'Occurrences'
        hdr_cells[3].text = 'English Translation'

        # Set column widths (in inches)
        widths = [Inches(0.7), Inches(1.5), Inches(1.1), Inches(3.0)]
        for i, width in enumerate(widths):
            table.columns[i].width = width

        # Helper to set cell padding and row height
        def set_cell_format(cell):
            tc = cell._tc
            tcPr = tc.get_or_add_tcPr()
            # Set cell padding
            for side in ('top', 'bottom', 'left', 'right'):
                node = OxmlElement(f'w:{side}Margin')
                node.set(qn('w:w'), '120')  # 120 twips = ~0.08 inch
                node.set(qn('w:type'), 'dxa')
                tcPr.append(node)

        for i, (word, count) in enumerate(words_counts, 1):
            row_cells = table.add_row().cells
            row_cells[0].text = str(i)
            row_cells[1].text = word
            row_cells[2].text = str(count)
            row_cells[3].text = ''  # blank for translation
            # Set row height (min height)
            tr = table.rows[-1]._tr
            trPr = tr.get_or_add_trPr()
            trHeight = OxmlElement('w:trHeight')
            trHeight.set(qn('w:val'), '800')  # 800 twips = ~0.56 cm
            trHeight.set(qn('w:hRule'), 'atLeast')
            trPr.append(trHeight)
            # Set cell padding for all cells in the row
            for cell in row_cells:
                set_cell_format(cell)

        doc.add_paragraph()  # Space between chapters

    doc.save(output_docx)
    print(f"Saved output DOCX to '{output_docx}'")

def main():
    input_pdf = input("Enter full path to input PDF file (e.g. C:\\Users\\You\\Desktop\\book.pdf): ").strip()
    if not os.path.isfile(input_pdf):
        print("Error: Input PDF file does not exist. Exiting.")
        return

    output_folder = input("Enter full path to output folder where result DOCX will be saved: ").strip()
    if not os.path.isdir(output_folder):
        print("Error: Output folder does not exist. Exiting.")
        return

    try:
        print("Converting PDF pages to images...")
        pages = convert_from_path(input_pdf, dpi=600)
    except Exception as e:
        print(f"Error during PDF to image conversion: {e}")
        return

    print(f"Performing OCR on {len(pages)} pages...")
    full_text = ""
    for i, page in enumerate(pages, 1):
        text = pytesseract.image_to_string(page, lang="fra")
        full_text += text + "\n"

    print("Splitting text into chapters...")
    chapters = split_text_into_chapters(full_text)
    if not chapters:
        print("Warning: No chapters found! Saving entire text as one chapter.")
        chapters = [full_text]

    print(f"Found {len(chapters)} chapters.")

    chapters_data = []
    for idx, chapter_text in enumerate(chapters, 1):
        print(f"Processing Chapter {idx}...")
        words = extract_words_spacy(chapter_text)
        word_counts = Counter(words)
        unique_words = sorted(word_counts.items())  # list of (word, count)
        chapters_data.append((idx, unique_words))

    output_docx = os.path.join(output_folder, "output_unique_words.docx")
    save_docx_output(chapters_data, output_docx)

    print("Done! Check the output DOCX for results.")

if __name__ == "__main__":
    main()
