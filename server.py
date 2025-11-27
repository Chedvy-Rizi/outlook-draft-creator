from flask import Flask, request, jsonify
from flask_cors import CORS
import os
import json
import pythoncom
import win32com.client
import tempfile

app = Flask(__name__)
CORS(app)

@app.route('/create_draft', methods=['POST'])
def create_draft():
    try:
        pythoncom.CoInitialize()

        subject = request.form['subject']
        body = request.form['body']
        to_list = json.loads(request.form['to'])

        attachment = request.files.get('attachment')
        file_path = None

        if attachment:
            with tempfile.NamedTemporaryFile(delete=False, suffix=attachment.filename) as tmp:
                attachment.save(tmp.name)
                file_path = tmp.name

        outlook = win32com.client.Dispatch("Outlook.Application")

        for to in to_list:
            mail = outlook.CreateItem(0)
            mail.Subject = subject
            mail.Body = body
            mail.To = to.strip()
            if file_path:
                mail.Attachments.Add(file_path)

            mail.Save()
            mail.Display()
            
        if file_path and os.path.exists(file_path):
            try:
                os.remove(file_path)
            except:
                pass

        return jsonify({"status": "success"})

    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500

    finally:
        pythoncom.CoUninitialize()


if __name__ == "__main__":
    app.run(port=5000)