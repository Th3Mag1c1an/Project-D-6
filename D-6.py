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
from PIL import Image
import fitz  # PyMuPDF
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment

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

def save_french_words_excel(chapters_data, output_xlsx):
    """
    chapters_data = list of tuples: (chapter_number, [(word, count), ...])
    Creates an Excel file with unique French words only.
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "French Words"
    
    # Set up header
    ws['A1'] = "French Word"
    
    # Style the header
    header_font = Font(bold=True)
    ws['A1'].font = header_font
    ws['A1'].alignment = Alignment(horizontal='center')
    
    # Collect all unique words across all chapters
    all_words = set()
    for chapter_num, words_counts in chapters_data:
        for word, _ in words_counts:
            all_words.add(word)
    
    # Add data - one row per unique word
    row = 2
    for word in sorted(all_words):
        ws[f'A{row}'] = word
        row += 1
    
    # Auto-adjust column width
    max_length = 0
    for cell in ws['A']:
        try:
            if len(str(cell.value)) > max_length:
                max_length = len(str(cell.value))
        except:
            pass
    adjusted_width = min(max_length + 2, 50)  # Cap at 50 characters
    ws.column_dimensions['A'].width = adjusted_width
    
    wb.save(output_xlsx)
    print(f"Saved unique French words Excel file to '{output_xlsx}'")

def extract_text_from_pdf(pdf_path):
    """
    Try to extract text directly from a PDF using PyMuPDF (fitz).
    Returns a list of strings, one per page.
    """
    text_pages = []
    with fitz.open(pdf_path) as doc:
        for page in doc:
            text = page.get_text()
            text_pages.append(text)
    return text_pages

def main():
    input_pdf = input("Enter full path to input PDF file (e.g. C:\\Users\\You\\Desktop\\book.pdf): ").strip()
    if not os.path.isfile(input_pdf):
        print("Error: Input PDF file does not exist. Exiting.")
        return

    output_folder = input("Enter full path to output folder where result XLSX will be saved: ").strip()
    if not os.path.isdir(output_folder):
        print("Error: Output folder does not exist. Exiting.")
        return

    try:
        print("Trying to extract text directly from PDF...")
        text_pages = extract_text_from_pdf(input_pdf)
        # Check if most pages have enough text (not just whitespace)
        if sum(len(t.strip()) > 30 for t in text_pages) > 0.7 * len(text_pages):
            print("PDF appears to be text-based. Using direct text extraction.")
            pages = text_pages
            is_text_based = True
        else:
            print("PDF appears to be image-based. Using OCR.")
            pages = convert_from_path(input_pdf, dpi=600)
            is_text_based = False
    except Exception as e:
        print(f"Error during PDF text extraction: {e}\nFalling back to OCR.")
        pages = convert_from_path(input_pdf, dpi=600)
        is_text_based = False

    print(f"Processing {len(pages)} pages...")
    full_text = ""
    for i, page in enumerate(pages, 1):
        if is_text_based:
            text = page
        else:
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

    # Save unique French words to Excel
    output_french_excel = os.path.join(output_folder, "unique_french_words.xlsx")
    save_french_words_excel(chapters_data, output_french_excel)

    print("Done! Check the output Excel file for results.")

if __name__ == "__main__":
    main()
