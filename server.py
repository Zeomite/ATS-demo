from flask import Flask, request, render_template
import textract
import re
import pandas as pd
from io import BytesIO

app = Flask(__name__)

def extract_emails_and_numbers(text):
    phone_numbers=[]
    for match in phonenumbers.PhoneNumberMatcher(text, "IN"):
        phone_numbers.append(str((match.number).national_number))
    emails = re.findall(r"\b[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+\b", text)
    return emails, phone_numbers

@app.route('/upload', methods=['POST'])
def upload_files():
    uploaded_files = request.files.getlist("file[]")
    extracted_data = []
    for file in uploaded_files:
        text = textract.process(file)
        name = re.findall(r'\b[A-Z][a-z]+ [A-Z][a-z]+\b', text.decode())
        email,phone = extract_emails_and_numbers(text)
        extracted_data.append({'Name': name[0] if name else None, 'Email': email, 'Phone': phone})
    
    df = pd.DataFrame(extracted_data)
    
    # Save DataFrame to Excel file
    output = BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    df.to_excel(writer, index=False)
    writer.save()
    output.seek(0)
    
    return send_file(output, attachment_filename='resumes_data.xls', as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)
