import os
import win32com.client
import pythoncom
import time
from flask import Flask, request, jsonify, send_from_directory
from flask_cors import CORS  # הוספת CORS
import tempfile
from werkzeug.utils import secure_filename
import threading, webbrowser
import re
from email.utils import parseaddr

app = Flask(__name__, static_folder="web_app", static_url_path="")
CORS(app)  # מאפשר לכל הדומיינים להתחבר לשרת


# פונקציות ליצירת טיוטה ב-Outlook
def _create_draft_with_outlook(outlook, recipient, subject, body, attachment_path=None, save_only=False):
    """יוצר טיוטה ב-Outlook. אם save_only=True, שומר בלי להציג חלון."""
    print(f"Creating draft for: {recipient}")
    try:
        mail = outlook.CreateItem(0)  # 0 = olMailItem
        mail.To = recipient
        mail.Subject = subject
        mail.Body = body
        mail.BodyFormat = 2  # טקסט ב־HTML
        if attachment_path and os.path.exists(attachment_path):
            mail.Attachments.Add(attachment_path)
        
        mail.Save()  # שמירה תמידית של הטיוטה
        
        if not save_only:
            mail.Display(False)  # הצגה רק אם לא במצב שמירה בלבד
            print(f"Draft created and displayed for {recipient}")
        else:
            print(f"Draft saved (not displayed) for {recipient}")
            
        time.sleep(1.5)  # השהייה קצרה בין טיוטות
        return True, None
        
    except Exception as e:
        error_msg = f"Error creating draft for {recipient}: {e}"
        print(error_msg)
        return False, error_msg


def create_outlook_draft(recipient, subject, body, attachment_path=None):
    """שמירה על תאימות לאחור: יוצר סשן Outlook זמני לטיוטה בודדת."""
    try:
        pythoncom.CoInitialize()
        outlook = win32com.client.Dispatch("Outlook.Application")
        return _create_draft_with_outlook(outlook, recipient, subject, body, attachment_path)
    except Exception as e:
        print(f"Error occurred: {e}")
        return False, str(e)
    finally:
        pythoncom.CoUninitialize()


@app.route("/")
def home():
    return send_from_directory("web_app", "index.html")


@app.route("/create-drafts", methods=["POST"])
def create_drafts():
    try:
        # קבלת נתונים מהטופס
        recipients = request.form.get("recipients")
        subject = request.form.get("subject")
        body = request.form.get("body")
        file = request.files.get("file")
        save_only = request.form.get("save_only", "false").lower() == "true"  # פרמטר חדש

        # טיפול בשגיאות במידה ויש נתונים חסרים
        if not recipients:
            return jsonify({"ok": False, "error": "Recipient is required"}), 400

        # הדפסת פרטי הבקשה
        print(f"Recipients: {recipients}")
        print(f"Subject: {subject}")
        print(f"Body: {body}")
        if file:
            print(f"Attachment: {file.filename}")

        # שמירת הקובץ המצורף אם קיים
        saved_path = None
        if file:
            tmp_dir = tempfile.gettempdir()
            filename = secure_filename(file.filename)
            saved_path = os.path.join(tmp_dir, filename)
            file.save(saved_path)

        # יצירת טיוטה עבור כל נמען - פרסינג חכם לנמענים
        raw_items = re.split(r'[;,\n]+', recipients or "")
        cleaned = []
        seen = set()
        for item in raw_items:
            s = item.strip()
            if not s:
                continue
            name, addr = parseaddr(s)
            addr = addr.strip()
            if not addr or "@" not in addr:
                continue
            key = addr.lower()
            if key in seen:
                continue
            seen.add(key)
            cleaned.append(addr)

        recipients_list = cleaned
        if not recipients_list:
            return jsonify({"ok": False, "error": "לא נמצאו כתובות מייל תקינות"}), 400
        results = []

        # פתיחת סשן COM יחיד לכול הרשימה לשיפור היציבות
        pythoncom.CoInitialize()
        try:
            # שימוש ב-EnsureDispatch כדי להבטיח יצירת ממשקים חזקים
            outlook = win32com.client.gencache.EnsureDispatch("Outlook.Application")
            # חיבור ל-MAPI ו"חימום" אם Outlook עובר אתחול ברקע
            try:
                session = outlook.GetNamespace("MAPI")
                session.Logon()
            except Exception:
                pass
            time.sleep(2.0)  # זמן חימום ראשוני ל-Outlook במידה והופעל עכשיו

            # חימום ראשוני
            time.sleep(2.0)
            
            # לולאה על הנמענים
            for i, recipient in enumerate(recipients_list, 1):
                try:
                    success, error = _create_draft_with_outlook(
                        outlook, 
                        recipient, 
                        subject, 
                        body, 
                        saved_path,
                        save_only=save_only
                    )
                    # השהייה ארוכה יותר אחרי כל 5 נמענים
                    if i % 5 == 0:
                        time.sleep(3.0)
                except Exception as e:
                    success, error = False, str(e)
                results.append({"recipient": recipient, "success": success, "error": error})
        finally:
            pythoncom.CoUninitialize()

        return jsonify({"ok": True, "results": results})

    except Exception as e:
        print(f"Error occurred: {e}")
        return jsonify({"ok": False, "error": str(e)}), 500


if __name__ == "__main__":
    try:
        print("Starting Flask server...")
        threading.Timer(0.7, lambda: webbrowser.open("http://127.0.0.1:5000?save_only=true")).start()
        app.run(host="127.0.0.1", port=5000, debug=False, use_reloader=False, threaded=False)
    except Exception as e:
        print(f"Error starting server: {e}")
