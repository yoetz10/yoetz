import asyncio
import os
import time
from datetime import datetime, timedelta
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from email.mime.text import MIMEText
import base64
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from telegram import Update
from telegram.ext import Application, MessageHandler, filters, ContextTypes, ConversationHandler, CommandHandler
import email
import openpyxl

# הגדרות וקבועים
SCOPES = [
    'https://www.googleapis.com/auth/gmail.send',
    'https://www.googleapis.com/auth/gmail.readonly',
    'https://www.googleapis.com/auth/gmail.modify'
]

CREDENTIALS_FILE = 'credentials/credentials.json'  # נתיב יחסי
TOKEN_FILE = 'token.json'
QUESTIONS_FILE = "questions.xlsx"

# פונקציות עזר
def create_message(sender, to, subject, body):
    message = MIMEText(body)
    message['to'] = to
    message['from'] = sender
    message['subject'] = subject
    return {'raw': base64.urlsafe_b64encode(message.as_bytes()).decode()}

def send_message(service, sender, message):
    try:
        message = service.users().messages().send(userId=sender, body=message).execute()
        print(f"Message sent to {sender}")
        return message
    except Exception as error:
        print(f"An error occurred: {error}")
        return None

def authenticate_gmail_api():
    creds = None
    if os.path.exists(TOKEN_FILE):
        creds = Credentials.from_authorized_user_file(TOKEN_FILE, SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS_FILE, SCOPES)
            creds = flow.run_local_server(port=0)
        with open(TOKEN_FILE, 'w') as token:
            token.write(creds.to_json())

    return build('gmail', 'v1', credentials=creds)

# פונקציות Excel
def load_questions_from_excel():
    try:
        workbook = openpyxl.load_workbook(QUESTIONS_FILE)
        sheet = workbook.active
        questions = {}
        for row in sheet.iter_rows(min_row=2):  # דילוג על השורה הראשונה (כותרות)
            question_id = row[0].value
            question_text = row[1].value
            question_title = row[2].value
            similar_questions = [q.value for q in row[3:] if q.value is not None]  # שאלות דומות
            questions[question_id] = {"text": question_text, "title": question_title, "similar": similar_questions}
        return questions
    except FileNotFoundError:
        return {}

def save_questions_to_excel(questions):
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    # הוספת כותרת עבור שאלות ללא דמיון
    sheet.append(["ID", "Question", "Title", "Similar Question 1", "Similar Question 2", "Similar Question 3", "No Similar Question"])
    for question_id, data in questions.items():
        row = [question_id, data["text"], data["title"]] + data["similar"] + [data.get("no_similar", "")]
        sheet.append(row)
    workbook.save(QUESTIONS_FILE)

# פונקציות Telegram

reminders = {}  # Initialize reminders
user_questions = {}  # מילון לשמירת הקשר בין שאלות למשתמשים

def find_similar_question(question_text, questions):  # מימוש הפונקציה - התאמת מילות מפתח
    keywords = question_text.lower().split()  # פיצול השאלה למילות מפתח
    for question_id, data in questions.items():
        if "text" in data:  # Check if the key exists
            existing_keywords = data["text"].lower().split()
            common_keywords = set(keywords) & set(existing_keywords)  # מציאת מילות מפתח משותפות
            if len(common_keywords) > 0:  # אם יש מילות מפתח משותפות
                return data["text"]  # החזר את השאלה הקיימת
    return None  # אם לא נמצאה שאלה דומה

async def get_question_title(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data["question_title"] = update.message.text
    await update.message.reply_text("שאלה התקבלה")
    return QUESTION_TITLE

async def handle_question(update: Update, context: ContextTypes.DEFAULT_TYPE, question_source="telegram"):
    if context.user_data is None:
        context.user_data = {}  # צור מילון ריק אם הוא None

    question_text = update.message.text if question_source == "telegram" else update["body"]
    user_name = update.message.from_user.full_name if question_source == "telegram" else update["sender"]
    title = context.user_data.get("question_title", question_text[:50])  # כותרת אוטומטית אם לא ניתנה

    service = authenticate_gmail_api()

    experts_emails = [  # מלא את זה עם כתובות מייל אמיתיות!
        "yorvada@gmail.com",
        "nomitrapi@gmail.com",
        "phdelazar@gmail.com",
    ]

    question_subject = f"שאלה חדשה מ {user_name} ({question_source})"
    question_body = f"שאלה: {question_text}\nכותרת: {title}"

    # שמירת שאלה באקסל
    questions = load_questions_from_excel()
    new_question_id = max(questions.keys()) + 1 if questions else 1
    questions[new_question_id] = {"text": question_text, "title": title, "similar": []}

    # בדיקה אם יש שאלה דומה
    similar_question = find_similar_question(question_text, questions)
    if similar_question:
        questions[new_question_id]["similar"].append(similar_question)
    else:
        questions[new_question_id]["no_similar"] = question_text

    save_questions_to_excel(questions)

    # שמירת הקשר בין השאלה למשתמש
    user_questions[new_question_id] = {
        "user": user_name,
        "source": question_source,
        "chat_id": update.message.chat_id if question_source == "telegram" else None,
        "email": update["sender"] if question_source == "email" else None
    }

    for expert_email in experts_emails:
        message = create_message("yoetz10bot@gmail.com", expert_email, question_subject, question_body)
        send_message(service, "yoetz10bot@gmail.com", message)

        # שמירת זמן השאלה ותזכורת
        question_time = datetime.now()
        reminders[expert_email] = {"time": question_time, "question": question_text, "user": user_name}

    await update.message.reply_text("השאלה שלך נשלחה למומחים.")
    if question_source == "email":
        # שליחת אישור למייל של השואל
        reply_subject = "השאלה שלך התקבלה"
        reply_body = "השאלה שלך התקבלה ונשלחה למומחים."
        reply_message = create_message("yoetz10bot@gmail.com", update["sender"], reply_subject, reply_body)
        send_message(service, "yoetz10bot@gmail.com", reply_message)

    return ConversationHandler.END  # סיום שיחה בטלגרם

async def check_for_answers(context: ContextTypes.DEFAULT_TYPE):
    try:
        service = authenticate_gmail_api()
        results = service.users().messages().list(userId='me', q="is:unread").execute()
        messages = results.get('messages', [])

        for message in messages:
            msg = service.users().messages().get(userId='me', id=message['id'], format='raw').execute()
            msg_str = base64.urlsafe_b64decode(msg['raw'].encode('ASCII'))
            mime_msg = email.message_from_bytes(msg_str)

            sender_email = mime_msg['From']
            subject = mime_msg['Subject']
            body = mime_msg.get_payload()

            # ניפוי שגיאות: הדפסת פרטי המייל
            print(f"New email received - From: {sender_email}, Subject: {subject}")

            # בדיקה אם המייל הוא תשובה לשאלה
            if "תשובה" in subject or "Re:" in subject:  # ניתן לשפר את הלוגיקה כאן
                # חילוץ השאלה המקורית מהמייל
                original_question = extract_original_question(body)  # פונקציה שתחלץ את השאלה המקורית מהתשובה
                
                if original_question:
                    print(f"Extracted original question: {original_question}")
                    
                    # מציאת המשתמש שהגיש את השאלה
                    user_to_notify = find_user_by_question(original_question)  # פונקציה שתמצא את המשתמש לפי השאלה
                    
                    if user_to_notify:
                        print(f"User to notify: {user_to_notify}")
                        
                        # שליחת התשובה למשתמש
                        if user_to_notify["source"] == "telegram":
                            await context.bot.send_message(chat_id=user_to_notify["chat_id"], text=f"קיבלת תשובה לשאלתך:\n{body}")
                            print("Answer sent to user via Telegram.")
                        else:  # אם המשתמש הוא במייל
                            reply_subject = "תשובה לשאלתך"
                            reply_body = f"קיבלת תשובה לשאלתך:\n{body}"
                            reply_message = create_message("yoetz10bot@gmail.com", user_to_notify["email"], reply_subject, reply_body)
                            send_message(service, "yoetz10bot@gmail.com", reply_message)
                            print("Answer sent to user via email.")
                    else:
                        print("User not found for the question.")
                else:
                    print("Could not extract the original question from the email.")

                # סימון המייל כנקרא
                service.users().messages().modify(userId='me', id=message['id'], body={'removeLabelIds': ['UNREAD']}).execute()
                print("Email marked as read.")
            else:
                print("Email is not a reply to a question.")

    except Exception as e:
        print(f"Error in check_for_answers: {e}")
        import traceback
        traceback.print_exc()

def extract_original_question(body):
    # פונקציה לדוגמה לחילוץ השאלה המקורית מהתשובה
    # ניתן לשפר את הלוגיקה כאן בהתאם לפורמט המייל
    lines = body.split("\n")
    for line in lines:
        if "שאלה:" in line:
            return line.replace("שאלה:", "").strip()
    return None

def find_user_by_question(question_text):
    # מציאת המשתמש לפי השאלה
    for question_id, data in user_questions.items():
        if data["text"] == question_text:
            return data
    return None

async def send_reminders(context: ContextTypes.DEFAULT_TYPE):
    try:
        service = authenticate_gmail_api()
        now = datetime.now()
        for expert, reminder in reminders.items():
            if now - reminder["time"] > timedelta(days=1):  # אם עבר יותר מיום
                reminder_subject = "תזכורת: שאלה ממתינה לתשובה"
                reminder_body = f"שלום {expert},\nתזכורת: שאלה ממתינה לתשובה:\n{reminder['question']}\n\nמשתמש: {reminder['user']}"
                reminder_message = create_message("yoetz10bot@gmail.com", expert, reminder_subject, reminder_body)
                send_message(service, "yoetz10bot@gmail.com", reminder_message)
                del reminders[expert]  # מחיקת התזכורת לאחר שליחה
    except Exception as e:
        print(f"Error in send_reminders: {e}")
        import traceback
        traceback.print_exc()

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("שלום! שלח/י את השאלה שלך.")
    return QUESTION

async def cancel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("הפעולה בוטלה.")
    return ConversationHandler.END

async def main():
    BOT_TOKEN = os.environ.get("BOT_TOKEN")  # קבל טוקן ממשתנה סביבה
    if not BOT_TOKEN:
        raise ValueError("BOT_TOKEN environment variable not set.")

    application = Application.builder().token(BOT_TOKEN).build()

    # ה-Conversation Handler
    conv_handler = ConversationHandler(
        entry_points=[CommandHandler("start", start)],
        states={
            QUESTION: [MessageHandler(filters.TEXT & ~filters.COMMAND, get_question_title)],
            QUESTION_TITLE: [MessageHandler(filters.TEXT & ~filters.COMMAND, handle_question)],
        },
        fallbacks=[CommandHandler("cancel", cancel)],
    )
    application.add_handler(conv_handler)

    # תזמון משימות (check_emails ו-send_reminders)
    application.job_queue.run_repeating(check_for_answers, interval=timedelta(minutes=10))
    application.job_queue.run_repeating(send_reminders, interval=timedelta(hours=24))

    try:
        print("מפעיל את הבוט...")
        # הפעלת הבוט
        await application.initialize()
        await application.start()
        await application.updater.start_polling()  # התחל להאזין להודעות
        while True:  # לולאה אינסופית כדי להשאיר את הבוט פעיל
            await asyncio.sleep(1)  # המתן לשנייה אחת לפני הבדיקה הבאה
    except Exception as e:  # טיפול בשגיאות
        print(f"שגיאה במהלך הפעלת הבוט: {e}")
        import traceback
        traceback.print_exc()  # הדפסת עקבות השגיאה
    finally:
        # סגירה נכונה של הבוט
        await application.updater.stop()  # עצור את ה-Updater
        await application.stop()
        await application.shutdown()
        print("הבוט נסגר בצורה תקינה.")

if __name__ == '__main__':
    QUESTION = range(1)
    QUESTION_TITLE = range(1)

    try:
        # הפעלת הפונקציה הראשית באמצעות asyncio
        asyncio.run(main())
    except KeyboardInterrupt:
        print("\nהבוט נעצר על ידי המשתמש")
    except Exception as e:
        print(f"שגיאה קריטית: {e}")
        import traceback
        traceback.print_exc()