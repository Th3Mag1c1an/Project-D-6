# Project D-6
The whole project was made using the help of ChatGPT to answer one of the questions i had when learning french was to increase my vocabulary in order to do you need to write and remember
vocabulary words which used to occur in the book i was frustrated with the process being slow wondered if you could take the photos of the chapters of the book and give it to a program could
it automate the process helping me find all the unique words meaning how can i save time over some menial work. And hence i got to this point.

## Description
The Simple Python Script takes in the 
1. Input -> A PDF of Images 
2. Which then Converts the PDF of Images to the Images
3. OCR Reads the Image
4. We use Spacy here to recognize all of the french words
5. We Find The Unique Words
6. We Then Convert that information to the table format
7. With S.No, Name of the word, Frequency of Occurance, And the last column which remains empty which is the Translation column

## Installation (Windows Only)
I am using Python 3.10.11 to avoid some issues i had with installing spaCy with the latest python so to avoid conflicts i recommend using the same Python 3.10.11 version.

You need the following Python packages and external dependencies:

1. pdf2image (to convert PDF pages to images)
2. pytesseract (OCR engine wrapper)
3. spacy (for French NLP)
4. fpdf (to generate PDF reports)
5. French spaCy model fr_core_news_sm
6. Tesseract OCR engine installed on your system (not a Python package)

Requirements.txt
pdf2image
pytesseract
spacy
fpdf
opencv-python
pillow
pandas
## Install all of the required libs
```pip install -r requirments.txt```

## Download French spaCy model
```python -m spacy download fr_core_news_sm ```

INCOMPLETE RIGHT NOW
