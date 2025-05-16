import os
from pdf2image import convert_from_path
import pytesseract
import spacy
from collections import Counter
from fpdf import FPDF
import re

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

def main():
    input_pdf = input("Enter full path to input PDF file (e.g. C:\\Users\\You\\Desktop\\book.pdf): ").strip()
    if not os.path.isfile(input_pdf):
        print("Error: Input PDF file does not exist. Exiting.")
        return

    output_folder = input("Enter full path to output folder where result PDF will be saved: ").strip()
    if not os.path.isdir(output_folder):
        print("Error: Output folder does not exist. Exiting.")
        return

    output_pdf = os.path.join(output_folder, "output_unique_words.pdf")

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

    pdf = FPDF()
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.set_font("Arial", size=12)

    for idx, chapter_text in enumerate(chapters, 1):
        print(f"Processing Chapter {idx}...")

        words = extract_words_spacy(chapter_text)
        word_counts = Counter(words)
        unique_words = sorted(word_counts.keys())

        pdf.add_page()
        pdf.set_font("Arial", 'B', 14)
        pdf.cell(0, 10, f"Chapter {idx} Unique Words", ln=True)
        pdf.ln(5)

        pdf.set_font("Arial", 'B', 12)
        pdf.cell(15, 10, "S.No", border=1)
        pdf.cell(60, 10, "French Word", border=1)
        pdf.cell(30, 10, "Occurrences", border=1)
        pdf.cell(80, 10, "English Translation", border=1)
        pdf.ln()

        pdf.set_font("Arial", size=11)
        for i, word in enumerate(unique_words, 1):
            pdf.cell(15, 10, str(i), border=1)
            pdf.cell(60, 10, word, border=1)
            pdf.cell(30, 10, str(word_counts[word]), border=1)
            pdf.cell(80, 10, "", border=1)  # Leave translation blank
            pdf.ln()

    pdf.add_page()

    print(f"Saving output PDF to '{output_pdf}'...")
    pdf.output(output_pdf)

    print("Done! Check the output PDF for results.")

if __name__ == "__main__":
    main()
