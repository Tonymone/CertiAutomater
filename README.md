# Certificate Generation Web Application

This is a Flask-based web application that generates certificates from provided Excel files (MS6 and BMS). It includes functionality to generate certificate images, compile them into a Word document, convert the Word document to PDF, and provide a download link for the generated PDF file.

## Features

- Upload MS6 and BMS Excel files to generate certificates.
- Dynamically updates users on the status of certificate generation.
- Downloads generated certificates as a PDF file.

## Requirements

- Python 3.6 or higher
- Flask
- flask-cors
- pandas
- openpyxl
- opencv-python
- docx2pdf
- numpy
- requests

## Usage

1. Clone the repository:

   ```bash
   git clone https://github.com/Tonymone/CertiAutomater.git
   cd CertiAutomater

2. Download packages
   ```bash
   npm install

3. Go to backend
   ```bash
   cd backend
   pip install -r requirements.txt
   python app.py

4. Start the application
   ```bash
   go to project root directory
   npm start

Access the application in your web browser at http://localhost:5000.

Upload MS6.xlsx and BMS.xlsx files containing the necessary data.

Click on "Generate Certificates" to initiate the process.

Monitor the status updates provided on the webpage.

Once certificates are generated, download the certificates.pdf file.
