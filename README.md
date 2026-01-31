# Outlook Drafts Generator

A lightweight Windows utility to automate the creation of multiple Outlook email drafts from a single interface. Perfect for personalized mass outreach without the complexity of a full mail merge.

## 🚀 Overview
This application allows users to input a list of recipients and generate individual Outlook drafts for each one. You can specify a subject line, message body, and even include an attachment.

## ✨ Features
* **Batch Processing:** Create a separate draft for every recipient in one click.
* **Attachment Support:** Automatically attaches a local file to every generated draft.
* **No Installation Required:** Ships as a standalone `.exe` (no Python environment needed).
* **Local & Secure:** Processes data locally on your machine via the Outlook Desktop API.

## 🛠 Requirements
* **OS:** Windows 10 or 11.
* **Software:** Microsoft Outlook Desktop (installed and logged in).
* **Permissions:** Ability to run executable files on your local drive.

## 📖 How to Use

1. **Download & Extract:** Download the latest release ZIP file and extract it to a local folder.
2. **Launch:** Double-click `OutlookDraftsApp.exe`. 
   * *Note: If Windows SmartScreen appears, click "More info" -> "Run anyway".*
3. **Interface:** Your browser will open the control panel at `http://127.0.0.1:5000`.
4. **Input Data:**
   * **Recipients:** Enter emails separated by commas (e.g., `user1@example.com, user2@example.com`).
   * **Content:** Fill in the Subject and Body fields.
   * **Attachment:** (Optional) Select a file to attach.
5. **Generate:** Click **"Open Drafts in Outlook"**.

## 🔍 Troubleshooting

| Issue | Solution |
| :--- | :--- |
| **Drafts don't appear** | Ensure Outlook Desktop is open and active. Try running the app as Administrator. |
| **"Cannot locate Outlook.Application"** | Verify that Microsoft Outlook is installed locally (not just the web version). |
| **Security Alerts** | Allow access if Outlook prompts for "Programmatic Access" or if your Firewall asks for permission. |

## 📝 Notes
* **Network Drives:** It is recommended to run the application from a local disk rather than a shared network drive for stability.
* **Temporary Files:** Attachments are stored in the Windows `%TEMP%` directory only during the draft creation phase and are not permanently stored by the app.

---
*Created for efficient workflow automation.*

מדריך קצר ללקוח להפעלת האפליקציה ליצירת טיוטות ב‑Outlook.

## מה האפליקציה עושה
- פותחת טיוטות חדשות ב‑Outlook (טיוטה אחת לכל נמען), לפי נתונים שתזינו בדף: נמענים, נושא, גוף וקובץ מצורף (לא חובה).

## דרישות
- מחשב Windows עם Outlook Desktop מותקן ומחובר לחשבון.
- אין צורך בהתקנת Python. האפליקציה מגיעה כ‑EXE מוכן.

## איך מפעילים
1) חלצו את הקובץ שקיבלתם (ZIP) לתיקייה מקומית במחשב.
2) פתחו את התיקייה והפעילו בלחיצה כפולה: `OutlookDraftsApp.exe`.
3) הדפדפן ייפתח לכתובת: `http://127.0.0.1:5000`.
   - אם Windows מציג SmartScreen, לחצו "More info" ואז "Run anyway".

## שימוש
1) בשדה "נמענים" הזינו אימיילים מופרדים בפסיקים (למשל: `a@x.com, b@x.com`).
2) הזינו "נושא" ו"גוף ההודעה".
3) הוסיפו קובץ מצורף (אופציונלי).
4) לחצו "פתח טיוטות ב‑Outlook". תיפתח טיוטה נפרדת לכל נמען.

## פתרון תקלות
- לא נפתחות טיוטות: ודאו ש‑Outlook פתוח ומחובר לחשבון. נסו להריץ כמנהל.
- הודעת "Cannot locate Outlook.Application": יש להתקין/לתקן את Outlook Desktop.
- התראות אבטחה: אשרו גישה כאשר מתבקשים (Outlook Programmatic Access / חומת אש).

## הערות
- מומלץ להפעיל מהדיסק המקומי, לא מכונן רשת.
- קובץ מצורף נשמר זמנית בתיקיית Temp של Windows בזמן יצירת הטיוטה בלבד.
.
