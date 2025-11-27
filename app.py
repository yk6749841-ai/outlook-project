
from flask import Flask, render_template, request, jsonify
from flask_cors import CORS
import os
import win32com.client as win32
import pythoncom
import sys
import time 

app = Flask(__name__)
CORS(app, resources={r"/*": {"origins": "*"}})

# ----------------------------------------------------------------------
# 1. פונקציה ליצירת טיוטות ב-Outlook
# ----------------------------------------------------------------------
def create_outlook_drafts(subject, recipients, body, attachment=None):
    pythoncom.CoInitialize()
    temp_files = [] 
    
    try:
        print("מנסה להתחבר ל-Outlook...")
        
        try:
            outlook = win32.GetActiveObject('Outlook.Application')
        except Exception:
            outlook = win32.Dispatch('Outlook.Application')

        print("התחברות ל-Outlook הצליחה! ממשיך ליצירת טיוטות...")
        
        recipients_list = [r.strip() for r in recipients.split(",") if r.strip()]

        for recipient in recipients_list:
            print(f"יוצר טיוטה עבור: {recipient}")
            
            mail = outlook.CreateItem(0)
            mail.Subject = subject
            mail.To = recipient
            
            mail.BodyFormat = 2  # HTML
            mail.HTMLBody = body
            
            # קובץ מצורף
            if attachment:
                if not temp_files: 
                    temp_path = os.path.join(os.getcwd(), attachment.filename)
                    attachment.save(temp_path)
                    temp_files.append(temp_path)
                    
                mail.Attachments.Add(temp_files[0])

            # 🔥 החלק החסר — שומר את הטיוטה לתיקיית Drafts
            mail.Save()

            # פותח חלון עריכה למשתמש
            mail.Display(False)
            
            time.sleep(0.5)
            
            print(f"טיוטה נפתחה ונשמרה בהצלחה.")
            
    except Exception as e:
        print("!!! כשל קריטי !!!")
        print("תיאור שגיאה:", e)
        return False, f"שגיאה ב-Outlook COM: {e}"
        
    finally:
        for temp_file in temp_files:
            try:
                os.remove(temp_file)
            except Exception as remove_e:
                print(f"שגיאה בניקוי הקובץ הזמני: {remove_e}.")
        pythoncom.CoUninitialize()
        return True, "הטיוטות נוצרו בהצלחה!"

# ----------------------------------------------------------------------
# ... שאר ה-Routes נשארים זהים ...
# ----------------------------------------------------------------------
@app.route("/")
def index():
    print("Route '/' נקרא!")
    return render_template("index.html")

@app.route("/create_drafts", methods=["POST"])
def create_drafts():
    print("POST /create_drafts נקרא!")
    
    subject = request.form.get("subject")
    recipients = request.form.get("recipients")
    body = request.form.get("body")
    attachment = request.files.get("attachment")

    success, message = create_outlook_drafts(subject, recipients, body, attachment)

    if success:
        return jsonify({"status": "success", "message": message})
    else:
        return jsonify({"status": "error", "message": message}), 500

if __name__ == "__main__":
    app.run(debug=True, threaded=False)