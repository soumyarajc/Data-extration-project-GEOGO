import docx
import csv

# Function to extract field names from the template file
def extract_fields(template_doc_path):
    doc = docx.Document(template_doc_path)
    fields = []
    for para in doc.paragraphs:
        if '$' in para.text:
            start_idx = para.text.find('$')
            end_idx = para.text.rfind('$')
            if start_idx != -1 and end_idx != -1 and start_idx != end_idx:
                field = para.text[start_idx:end_idx+1]
                fields.append(field)
    return fields

# Function to extract values from the Word document
def extract_values(doc_path, fields):
    doc = docx.Document(doc_path)
    values = {field: None for field in fields}
    for para in doc.paragraphs:
        for field in fields:
            if field in para.text:
                values[field] = para.text.replace(field, "").strip()
    return values

# Paths to the template and the document files
template_doc_path = '/Users/soumyarajchatterjee/MS-Word template file.docx'
doc_path = '/Users/soumyarajchatterjee/MS-Word file.docx'

# Extract fields from the template
fields = extract_fields(template_doc_path)
print(f"Extracted fields: {fields}")

# Extract values from the document
values = extract_values(doc_path, fields)
print(f"Extracted values: {values}")

# Write the results to a CSV file
csv_file_path = 'field_values.csv'
with open(csv_file_path, mode='w', newline='', encoding='utf-8') as file:
    writer = csv.writer(file)
    writer.writerow(['Field', 'Value'])
    for field, value in values.items():
        writer.writerow([field, value])

print(f"CSV file created: {csv_file_path}")






