from flask import Flask, render_template, request
import os
import re
import smtplib
from email.message import EmailMessage
# Updated to match your topsis.py function name
from topsis import calculate_topsis 

app = Flask(__name__)

UPLOAD_FOLDER = "uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

def valid_email(email):
    pattern = r'^[\w\.-]+@[\w\.-]+\.\w+$'
    return re.match(pattern, email)

def send_email(receiver, file_path):
    # WARNING: Use a Google App Password, not your login password
    sender = "abc@gmail.com"
    password = "abcd efgh ijkl mnop" 

    msg = EmailMessage()
    msg['Subject'] = "TOPSIS Result"
    msg['From'] = sender
    msg['To'] = receiver
    msg.set_content("Please find attached the TOPSIS result.")

    with open(file_path, 'rb') as f:
        msg.add_attachment(
            f.read(),
            maintype='application',
            # Updated subtype for Excel files
            subtype='vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            filename="result.xlsx"
        )

    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
        smtp.login(sender, password)
        smtp.send_message(msg)

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        try:
            # 1. Collect form data
            file = request.files['file']
            weights = request.form['weights']
            impacts = request.form['impacts']
            email = request.form['email']

            # 2. Basic validation
            if not valid_email(email):
                return "Error: Invalid Email Address", 400

            # 3. Save uploaded file
            filepath = os.path.join(UPLOAD_FOLDER, file.filename)
            file.save(filepath)

            # 4. Define output path (matching your topsis.py .to_excel output)
            output_path = os.path.join(UPLOAD_FOLDER, "result.xlsx")

            # 5. Run your TOPSIS function
            # Note: We pass the raw strings for weights/impacts because 
            # your topsis.py handles the .split() and .strip() logic.
            calculate_topsis(filepath, weights, impacts, output_path)

            # 6. Send result via email
            send_email(email, output_path)

            return "Success: Result sent to email!"

        except Exception as e:
            # Catching the custom 'raise Exception' messages from topsis.py
            return f"An error occurred: {str(e)}", 500

    return render_template('index.html')

if __name__ == '__main__':
    app.run(debug=True)