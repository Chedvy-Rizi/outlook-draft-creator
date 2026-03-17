import os
import json
import time
import threading
import logging
import pythoncom
import win32com.client
from flask import Flask, request, jsonify
from flask_cors import CORS
from werkzeug.utils import secure_filename
import bleach 

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[logging.FileHandler("app_log.txt"), logging.StreamHandler()]
)

app = Flask(__name__)
CORS(app)

UPLOAD_FOLDER = os.path.join(os.getcwd(), 'secure_uploads')
ALLOWED_EXTENSIONS = {'pdf', 'doc', 'docx', 'jpg', 'png'}
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

def is_allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def open_outlook_logic(subject, body, to_list, file_path):
    pythoncom.CoInitialize()
    try:
        logging.info(f"Starting Outlook process for {len(to_list)} recipients.")
        outlook = win32com.client.Dispatch("Outlook.Application")
        
        for recipient in to_list:
            mail = outlook.CreateItem(0)
            mail.Subject = subject
            mail.HTMLBody = body 
            mail.To = recipient.strip()
            
            if file_path and os.path.exists(file_path):
                mail.Attachments.Add(file_path)
            
            mail.Save()
            mail.Display()
        
        if file_path and os.path.exists(file_path):
            time.sleep(10) 
            os.remove(file_path)
            logging.info(f"Temporary file deleted: {file_path}")

    except Exception as e:
        logging.error(f"Error in background process: {str(e)}")
    finally:
        pythoncom.CoUninitialize()

@app.route('/create_draft', methods=['POST'])
def create_draft():
    try:
        subject = bleach.clean(request.form.get('subject', 'No Subject'))
        raw_body = request.form.get('body', '')
        safe_body = bleach.clean(raw_body, tags=['b', 'i', 'u', 'br', 'p'], attributes={})
        
        try:
            to_list = json.loads(request.form.get('to', '[]'))
        except json.JSONDecodeError:
            return jsonify({"status": "error", "message": "Invalid recipients list"}), 400
        
        attachment = request.files.get('attachment')
        file_path = None

        if attachment:
            if is_allowed_file(attachment.filename):
                filename = secure_filename(attachment.filename)
                file_path = os.path.join(UPLOAD_FOLDER, filename)
                attachment.save(file_path)
                logging.info(f"File saved: {filename}")
            else:
                return jsonify({"status": "error", "message": "File type not allowed"}), 400

        thread = threading.Thread(
            target=open_outlook_logic, 
            args=(subject, safe_body, to_list, file_path)
        )
        thread.start()

        return jsonify({
            "status": "success", 
            "message": "The request is being processed. Outlook will open shortly."
        }), 202

    except Exception as e:
        logging.error(f"Critical server error: {str(e)}")
        return jsonify({"status": "error", "message": "Internal Server Error"}), 500

if __name__ == "__main__":
    logging.info("Server started on port 5000")
    app.run(port=5000)