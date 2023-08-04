import pandas as pd
import re
import os
import fitz  # PyMuPDF
import docx
import os
from collections import Counter


source = r'C:/Users/alexb/OneDrive/Documents/Python Scripts/Input/CVs'

# Identify all PDFs in provided filepath, once found extract text with read_text_from_pdf
def read_text_from_documents(folder_path):
    all_text = []
    for filename in os.listdir(folder_path):
        if filename.endswith(".pdf"):
            pdf_file_path = os.path.join(folder_path, filename)
            #text = read_text_from_pdf(pdf_file_path)
            text = ""
            pdf_document = fitz.open(pdf_file_path)
            for page_number in range(pdf_document.page_count):
                page = pdf_document.load_page(page_number)
                text += page.get_text()
            pdf_document.close()
        if filename.endswith(".docx"):
            doc = docx.Document(os.path.join(folder_path, filename))
            text = []
            for docpara in doc.paragraphs:
                text.append(docpara.text)
            text = '\n'.join([i for i in text[1:]])
        all_text.append([filename,text]) #all_text.append(text)
    return all_text

# Case insensitive search of keywords in text
def analyse_cv(cv_text, keywords):
    keyword_count = {keyword: len(re.findall(rf'\b{keyword}\b', cv_text, re.IGNORECASE)) for keyword in keywords}
    return keyword_count

def main():
    # Sample list of keywords to analyze
    keywords = ['python', 'R', 'machine learning', 'data analysis', 'data science', 'SQL',
                'MySQL','SQLite','dplyr','pandas', 'numpy', 'ggplot', 'matplotlib', 'matlab', 
                'SAS', 'Stata', 'Alteryx','model'] 

    texts_from_pdfs = read_text_from_documents(source) 
    cvs = texts_from_pdfs
    dfcsv = pd.DataFrame(cvs, columns=['file','text'])
    
    # Step 3: Create a function to count the occurrences of words from the list in a given text
    def count_words(text, word_list):
        #word_counts = [text.lower().count(word.lower()) for word in word_list]
        word_counts = {keyword: len(re.findall(rf'\b{keyword}\b', text, re.IGNORECASE)) for keyword in keywords}
        return word_counts
    
    # Apply the function to the DataFrame column and create separate columns
    word_counts_df = dfcsv['text'].apply(lambda text: pd.Series(count_words(text, keywords)))
    # Concatenate the word_counts_df to the original DataFrame
    dfcsv = pd.concat([dfcsv, word_counts_df], axis=1)
    # Rename the columns
    dfcsv.columns = list(dfcsv.columns[:-len(keywords)]) + keywords
    
    # Create a score total column
    dfcsv['sum'] = dfcsv.sum(axis=1)
    
    return dfcsv


if __name__ == "__main__":
    df = main()



# Function to remove common words
def remove_common_words(text):
    filler_words = ['a', 'an', 'the', 'is', 'are', 'and', 'or', 'in', 'on', 'at', 'to', 'of',
                'by', 'name', 'with', 'for', 'that', 'from']
    words = re.findall(r'\b\w+\b', text.lower())
    return ' '.join([word for word in words if word not in filler_words])

# Apply the function to the 'text_column' and create a new column 'cleaned_text'
#df['cleaned_text'] = df['text'].apply(remove_common_words)
cleaned_text = df['text'].apply(remove_common_words)

def most_common_words(list_of_strings, n):
    # Combine all strings into a single string
    combined_string = ' '.join(list_of_strings)
    # Remove punctuation and convert to lowercase
    cleaned_string = re.sub(r'[^\w\s]', '', combined_string.lower())
    # Split the cleaned string into individual words
    words = cleaned_string.split()
    word_counter = Counter(words)
    most_common_words = word_counter.most_common(n)
    return most_common_words

num_most_common = 20
result = pd.DataFrame(most_common_words(cleaned_text, num_most_common), columns=['word','freq'])


with open("a.html", 'w', encoding="utf-8") as _file:
    _file.write(result.to_html()
                + "\n\n\n" + 
                df.drop('text', axis=1).to_html())
    #_file.write(result.to_html())
