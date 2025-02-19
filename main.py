import re
import csv
from datetime import datetime
import pypandoc

def extract_questions_from_docx(file_path, output_csv):
    """
    extracts docs file and make them structured file -> tabular format
    """
    text = pypandoc.convert_file(file_path, 'plain')
    paragraphs = text.split("\n")  # Splitting text into lines to mimic paragraphs
    
    extracted_data = []
    
    current_date = None
    current_name = None
    current_company = None
    yc_batch = None
    question_text = None
    replies = 0
    
    date_pattern = re.compile(r'\b(January|February|March|April|May|June|July|August|September|October|November|December)\s\d{1,2}(st|nd|rd|th)?,?\s\d{4}\b')
    name_pattern = re.compile(r'([A-Z][a-z]+\s[A-Z][a-z]+)(?:\s-\s([A-Za-z0-9]+))?')
    reply_pattern = re.compile(r'(\d+)\srepl(?:y|ies)')
    yc_batch_pattern = re.compile(r'\bYC\s([A-Z]+\d{2})\b')
    
    for text in paragraphs:
        text = text.strip()
        
        if not text:
            continue
        
        # date pattern checking
        if date_pattern.search(text):
            current_date = text
            continue
        
        # check for name and optional company
        name_match = name_pattern.match(text)
        if name_match:
            current_name, current_company = name_match.groups()
            continue
        
        # check for the YC batch
        yc_match = yc_batch_pattern.search(text)
        if yc_match:
            yc_batch = yc_match.group(1)
        
        # check for replied
        reply_match = reply_pattern.search(text)
        if reply_match:
            replies = int(reply_match.group(1))
        
        # capture the question text
        if current_name and current_date:
            question_text = text
            extracted_data.append([current_date, current_name, current_company, yc_batch, question_text, replies])
            
            # reset variables
            current_name = None
            current_company = None
            yc_batch = None
            question_text = None
            replies = 0
    
    # saveing to csv
    with open(output_csv, mode='w', newline='', encoding='utf-8') as file:
        writer = csv.writer(file)
        writer.writerow(["Date", "Name", "Company", "YC Batch", "Question", "Total Replies"])
        writer.writerows(extracted_data)
    
    print(f"Extracted {len(extracted_data)} entries and saved to {output_csv}")

# file reading and function call
input_docx = "/Users/bikashpokharel/Downloads/YC_AWS_Slack_General_Chat.docx"
output_csv = "/Users/bikashpokharel/Downloads/AWS_Slack_Questions_Extracted.csv"
extract_questions_from_docx(input_docx, output_csv)