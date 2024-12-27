import os
from flask import Flask, render_template, request, redirect, flash
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import openpyxl

app = Flask(__name__)
app.secret_key = 'your_secret_key'

UPLOAD_FOLDER = 'uploads'
EXCEL_FILE = 'submissions.xlsx'
EMAIL_ID = 'pk400711@gmail.com'

# Ensure uploads directory exists
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

# Initialize Excel file
if not os.path.exists(EXCEL_FILE):
    wb = openpyxl.Workbook()
    sheet = wb.active
    sheet.append(['Name', 'Phone Number', 'Email', 'Address', 'Details'])
    wb.save(EXCEL_FILE)

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

def send_email(files, form_data):
    """Send email with uploaded files and form details."""
    sender_email = "your_email@gmail.com"  # Replace with your email
    sender_password = "your_password"     # Replace with your email password
    receiver_email = EMAIL_ID

    # Email setup
    message = MIMEMultipart()
    message['From'] = sender_email
    message['To'] = receiver_email
    message['Subject'] = "New Submission: Shobha Consultancy Services"

    # Email body
    body = f"""
    Name: {form_data['name']}
    Phone Number: {form_data['phone']}
    Email: {form_data['email']}
    Address: {form_data['address']}
    Details: {form_data['details']}
    Aadhar & PAN linked with Phone Number: {form_data['phone']}
    """
    message.attach(MIMEText(body, 'plain'))

    # Attach uploaded files
    for file_path in files:
        attachment = MIMEBase('application', 'octet-stream')
        with open(file_path, 'rb') as f:
            attachment.set_payload(f.read())
        encoders.encode_base64(attachment)
        attachment.add_header(
            'Content-Disposition',
            f'attachment; filename={os.path.basename(file_path)}'
        )
        message.attach(attachment)

    # Send email
    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as server:
        server.login(sender_email, sender_password)
        server.send_message(message)

@app.route('/', methods=['GET', 'POST'])
def contact_form():
    if request.method == 'POST':
        # Collect form data
        name = request.form['name']
        phone = request.form['phone']
        email = request.form['email']
        address = request.form['address']
        details = request.form['details']
        aadhar_file = request.files['aadhar']
        pan_file = request.files['pan']
        bank_statement = request.files['bank_statement']

        # Save uploaded files
        uploaded_files = []
        for file in [aadhar_file, pan_file, bank_statement]:
            if file:
                file_path = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
                file.save(file_path)
                uploaded_files.append(file_path)

        # Append data to Excel
        wb = openpyxl.load_workbook(EXCEL_FILE)
        sheet = wb.active
        sheet.append([name, phone, email, address, details])
        wb.save(EXCEL_FILE)

        # Send email with uploaded files and form data
        try:
            send_email(uploaded_files, {
                'name': name,
                'phone': phone,
                'email': email,
                'address': address,
                'details': details
            })
            flash('Form submitted successfully! Documents mailed to admin.', 'success')
        except Exception as e:
            flash(f'Error sending email: {str(e)}', 'danger')

        return redirect('/')

    return render_template('form.html')

if __name__ == '__main__':
    app.run(debug=True)
