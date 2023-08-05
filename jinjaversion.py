# -*- coding: utf-8 -*-
"""
Created on Sat Aug  5 12:04:07 2023

@author: alexb
"""

import os
#import glob
#import PyPDF2
import fitz
from collections import Counter
from docx import Document
from jinja2 import Environment, FileSystemLoader

from nltk.corpus import stopwords
import re

import pandas as pd


def read_pdf(file_path):
    text = ""
    pdf_document = fitz.open(file_path)
    for page_number in range(pdf_document.page_count):
        page = pdf_document.load_page(page_number)
        text += page.get_text()
    pdf_document.close()
    
    return text
            
def read_docx(file_path):
    text = ""
    doc = Document(file_path)
    for para in doc.paragraphs:
        text += para.text + "\n"
    return text

def removestopwords(string):
    cachedStopWords = stopwords.words("english")
    cleantext = ' '.join([word for word in string.split() if word not in cachedStopWords])
    return cleantext

def analyze_text(text, keywords):
    # Remove filler "stop" words (a, and, for, etc.)
    text = removestopwords(text)
    text = re.sub("[^\w\d'\s]+",'', text)
    
    words = text.lower().split()
    keyword_count = {keyword: words.count(keyword.lower()) for keyword in keywords}
    word_count = Counter(words)
    return keyword_count, word_count

def generate_html_report(results, output_path):
    env = Environment(loader=FileSystemLoader("."))
    template = env.get_template("report_template.html")
    rendered_html = template.render(results=results)
    with open(output_path, "w", encoding="utf-8") as report_file:
        report_file.write(rendered_html)

def main():
    folder_path = "./Input/CVs"
    keywords = ['python', 'R', 'machine learning', 'data analysis', 'data science', 'SQL',
                'MySQL','SQLite','dplyr','pandas', 'numpy', 'ggplot', 'matplotlib', 'matlab', 
                'SAS', 'Stata', 'Alteryx','model']   # Add your defined list of keywords here
    common_words = 15
    
    results = []

    for filename in os.listdir(folder_path):
        if filename.endswith(".pdf"):
            text = read_pdf(os.path.join(folder_path, filename))
        if filename.endswith(".docx"):
            text = read_docx(os.path.join(folder_path, filename))
        keyword_count, word_count = analyze_text(text, keywords)
        results.append({
            "file_name": os.path.basename(filename),
            "keyword_count": keyword_count,
            "most_frequent_words": word_count.most_common(common_words)
        })

    output_path = "./output/report.html"
    generate_html_report(results, output_path)
    print("Report generated successfully.")

if __name__ == "__main__":
    main()




#df = pd.DataFrame(results)
