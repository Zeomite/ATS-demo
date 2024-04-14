from flask import Flask, request, render_template, send_file
import os
import re
import xlwt
import phonenumbers
import textract

app = Flask(__name__)

UPLOAD_FOLDER = 'uploads'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

def extract_emails_and_numbers(text):
    phone_numbers=[]
    for match in phonenumbers.PhoneNumberMatcher(text, "IN"):
        phone_numbers.append(str((match.number).national_number))
    emails = re.findall(r"\b[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+\b", text)
    return emails, phone_numbers

@app.route('/', methods=['GET', 'POST'])
def upload_file():
    if request.method == 'POST':
        files = request.files.getlist('files[]')
        extracted_data = []
        for file in files:
            filename = file.filename
            file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            file.save(file_path)
            """if filename.endswith('.pdf'):
                text = extract_text_from_pdf(file_path)
            elif filename.endswith('.docx'):
                text = extract_text_from_docx(file_path)
            elif filename.endswith('.doc'):
               """
            text= textract.process(file_path)
            emails, phone_numbers = extract_emails_and_numbers(text)
            extracted_data.append({'filename': filename, 'emails': emails, 'phone_numbers': phone_numbers})
        xls_filename = 'extracted_data.xls'
        workbook = xlwt.Workbook()
        sheet = workbook.add_sheet('Data')
        sheet.write(0, 0, 'Filename')
        sheet.write(0, 1, 'Emails')
        sheet.write(0, 2, 'Phone Numbers')
        row = 1
        for data in extracted_data:
            sheet.write(row, 0, data['filename'])
            sheet.write(row, 1, ', '.join(data['emails']))
            sheet.write(row, 2, ', '.join(data['phone_numbers']))
            row += 1
        xls_filepath = os.path.join(app.config['UPLOAD_FOLDER'], xls_filename)
        workbook.save(xls_filepath)
        return send_file(xls_filepath, as_attachment=True)
    return render_template('upload.html')

if __name__ == '__main__':
    app.run(debug=True)
