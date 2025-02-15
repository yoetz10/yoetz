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
from dotenv import load_dotenv

# טעינת משתני סביבה
load_dotenv()

# הגדרות וקבועים
SCOPES = [
    'https://www.googleapis.com/auth/gmail.send',
    'https://www.googleapis.com/auth/gmail.readonly',
    'https://www.googleapis.com/auth/gmail.modify'
]

# נתיבים לקבצים
CREDENTIALS_FILE = os.getenv('GOOGLE_CREDENTIALS_PATH', 'credentials/credentials.json')
TOKEN_FILE = 'token.json'
QUESTIONS_FILE = "questions.xlsx"

# קבועים לניהול שיחה
QUESTION, QUESTION_TITLE = range(2)

# מבני נתונים גלובליים
reminders = {}
user_questions = {}

def create_message(sender, to, subject, body):
    message = MIMEText(body, 'plain', 'utf-8')
    message['to'] = to
    message['from'] = sender
    message['subject'] = subject
    return {'raw': base64.urlsafe_b64encode(message.as_bytes()).decode()}

def send_message(service, sender, message):
    try:
        message = service.users().messages().send(userId=sender, body=message).execute()
        print(f"Message sent to {message['to']}")
        return message
    except Exception as error:
        print(f"An error occurred while sending message: {error}")
        return None

def authenticate_gmail_api():
    creds = None
    if os.path.exists(TOKEN_FILE):
        creds = Credentials.from_authorized_user_file(TOKEN_FILE, SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            if not os.path.exists(CREDENTIALS_FILE):
                raise FileNotFoundError(f"Missing credentials file at {CREDENTIALS_FILE}")
            flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS_FILE, SCOPES)
            creds = flow.run_local_server(port=0)
        with open(TOKEN_FILE, 'w') as token:
            token.write(creds.to_json())

    return build('gmail', 'v1', credentials=creds)

def ensure_excel_file_exists():
    if not os.path.exists(QUESTIONS_FILE):
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        sheet.append(["ID", "Question", "Title", "Similar Question 1", "Similar Question 2", "Similar Question 3"])
        workbook.save(QUESTIONS_FILE)

def load_questions_from_excel():
    ensure_excel_file_exists()
    try:
        workbook = openpyxl.load_workbook(QUESTIONS_FILE)
        sheet = workbook.active
        questions = {}
        for row in sheet.iter_rows(min_row=2):
            question_id = row[0].value
            if question_id is not None:
                questions[question_id] = {
                    "text": row[1].value,
                    "title": row[2].value,
                    "similar": [q.value for q in row[3:] if q.value is not None]
                }
        return questions
    except Exception as e:
        print(f"Error loading questions from Excel: {e}")
        return {}

def save_questions_to_excel(questions):
    try:
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        sheet.append(["ID", "Question", "Title", "Similar Question 1", "Similar Question 2", "Similar Question 3"])
        
        for question_id, data in questions.items():
            row = [
                question_id,
                data["text"],
                data["title"]
            ] + data["similar"][:3] + [''] * (3 - len(data["similar"]))
            sheet.append(row)
            
        workbook.save(QUESTIONS_FILE)
    except Exception as e:
        print(f"Error saving questions to Excel: {e}")

def find_similar_question(question_text, questions):
    if not question_text:
        return None
        
    keywords = set(question_text.lower().split())
    best_match = None
    max_common_words = 0
    
    for question_id, data in questions.items():
        if not data.get("text"):
            continue
            
        existing_keywords = set(data["text"].lower().split())
        common_keywords = keywords & existing_keywords
        
        if len(common_keywords) > max_common_words:
            max_common_words = len(common_keywords)
            best_match = data["text"]
            
    return best_match if max_common_words >= 2 else None

async def handle_question(update: Update, context: ContextTypes.DEFAULT_TYPE, question_source="telegram"):
    try:
        if not context.user_data:
            context.user_data = {}

        question_text = update.message.text if question_source == "telegram" else update["body"]
        user_name = update.message.from_user.full_name if question_source == "telegram" else update["sender"]
        title = context.user_data.get("question_title", question_text[:50])

        service = authenticate_gmail_api()

        experts_emails = [
            "yorvada@gmail.com",
            "nomitrapi@gmail.com",
            "phdelazar@gmail.com",
        ]

        # שליחת השאלה למומחים
        question_subject = f"שאלה חדשה מ {user_name}"
        question_body = f"""
שאלה חדשה התקבלה:

שואל: {user_name}
מקור: {question_source}
כותרת: {title}

תוכן השאלה:
{question_text}

אנא השב לשאלה זו בהקדם האפשרי.
"""

        questions = load_questions_from_excel()
        new_question_id = max(questions.keys(), default=0) + 1
        
        questions[new_question_id] = {
            "text": question_text,
            "title": title,
            "similar": []
        }

        similar_question = find_similar_question(question_text, 
            {k: v for k, v in questions.items() if k != new_question_id}
        )
        
        if similar_question:
            questions[new_question_id]["similar"].append(similar_question)

        save_questions_to_excel(questions)

        user_questions[new_question_id] = {
            "user": user_name,
            "source": question_source,
            "chat_id": update.message.chat_id if question_source == "telegram" else None,
            "email": update["sender"] if question_source == "email" else None,
            "text": question_text
        }

        for expert_email in experts_emails:
            message = create_message(
                "yoetz10bot@gmail.com",
                expert_email,
                question_subject,
                question_body
            )
            send_message(service, "yoetz10bot@gmail.com", message)
            reminders[expert_email] = {
                "time": datetime.now(),
                "question": question_text,
                "user": user_name
            }

        await update.message.reply_text("השאלה שלך נשלחה למומחים. תקבל תשובה בקרוב.")
        return ConversationHandler.END

    except Exception as e:
        print(f"Error in handle_question: {e}")
        await update.message.reply_text("אירעה שגיאה בעת שליחת השאלה. אנא נסה שוב מאוחר יותר.")
        return ConversationHandler.END

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
            
            # קבלת תוכן ההודעה
            if mime_msg.is_multipart():
                for part in mime_msg.walk():
                    if part.get_content_type() == 'text/plain':
                        body = part.get_payload(decode=True).decode()
                        break
            else:
                body = mime_msg.get_payload(decode=True).decode()

            if "תשובה" in subject or "Re:" in subject:
                original_question = extract_original_question(body)
                if original_question:
                    user_to_notify = find_user_by_question(original_question)
                    
                    if user_to_notify:
                        # הסרת הציטוט המקורי מהתשובה
                        answer = clean_reply_text(body)
                        
                        if user_to_notify["source"] == "telegram":
                            try:
                                await context.bot.send_message(
                                    chat_id=user_to_notify["chat_id"],
                                    text=f"קיבלת תשובה לשאלתך:\n\n{answer}"
                                )
                            except Exception as e:
                                print(f"Error sending Telegram message: {e}")
                        else:
                            reply_message = create_message(
                                "yoetz10bot@gmail.com",
                                user_to_notify["email"],
                                "תשובה לשאלתך",
                                f"קיבלת תשובה לשאלתך:\n\n{answer}"
                            )
                            send_message(service, "yoetz10bot@gmail.com", reply_message)

            # סימון המייל כנקרא
            service.users().messages().modify(
                userId='me',
                id=message['id'],
                body={'removeLabelIds': ['UNREAD']}
            ).execute()

    except Exception as e:
        print(f"Error in check_for_answers: {e}")

def extract_original_question(body):
    lines = body.split("\n")
    for line in lines:
        if "שאלה:" in line:
            return line.replace("שאלה:", "").strip()
    return None

def clean_reply_text(body):
    # הסרת ציטוטים וחתימות
    lines = body.split("\n")
    clean_lines = []
    for line in lines:
        if not line.startswith(">") and not line.startswith("On") and not line.startswith("From:"):
            clean_lines.append(line)
    return "\n".join(clean_lines).strip()

def find_user_by_question(question_text):
    for question_id, data in user_questions.items():
        if data.get("text") == question_text:
            return data
    return None

async def send_reminders(context: ContextTypes.DEFAULT_TYPE):
    try:
        service = authenticate_gmail_api()
        now = datetime.now()
        reminders_to_delete = []

        for expert, reminder in reminders.items():
            if now - reminder["time"] > timedelta(days=1):
                reminder_subject = "תזכורת: שאלה ממתינה לתשובה"
                reminder_body = f"""
שלום,

זוהי תזכורת אוטומטית לשאלה שטרם נענתה:

שואל: {reminder['user']}
שאלה: {reminder['question']}

אנא השב בהקדם האפשרי.

בברכה,
מערכת יועץ
"""
                message = create_message(
                    "yoetz10bot@gmail.com",
                    expert,
                    reminder_subject,
                    reminder_body
                )
                send_message(service, "yoetz10bot@gmail.com", message)
                reminders_to_delete.append(expert)

        for expert in reminders_to_delete:
            del reminders[expert]

    except Exception as e:
        print(f"Error in send_reminders: {e}")

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "ברוך הבא ליועץ! אנא שלח את השאלה שלך ואעביר אותה למומחים שלנו."
    )
    return QUESTION

async def get_question_title(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data["question_title"] = update.message.text
    await update.message.reply_text("תודה! השאלה התקבלה ותועבר למומחים שלנו.")
    return await handle_question(update, context)

async def cancel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("הפעולה בוטלה. אתה יכול להתחיל מחדש עם /start")
    return ConversationHandler.END

async def main():
    try:
        # טעינת טוקן הבוט ממשתנה סביבה
        BOT_TOKEN = os.getenv("BOT_TOKEN")
        if not BOT_TOKEN:
            raise ValueError("Missing BOT_TOKEN in environment variables")

        # יצירת אפליקציית הבוט
        application = Application.builder().token(BOT_TOKEN).build()

        # הגדרת ה-Conversation Handler
        conv_handler = ConversationHandler(
            entry_points=[CommandHandler("start", start)],
            states={
                QUESTION: [MessageHandler(filters.TEXT & ~filters.COMMAND, get_question_title)],
                QUESTION_TITLE: [MessageHandler(filters.TEXT & ~filters.COMMAND, handle_question)]
            },
            fallbacks=[CommandHandler("cancel", cancel)]
        )
        application.add_handler(conv_handler)

        # הגדרת משימות מתוזמנות
        application.job_queue.run_repeating(check_for_answers, interval=300)  # כל 5 דקות
        application.job_queue.run_repeating(send_reminders, interval=86400)   # כל
        
        print("Starting bot...")
        # איתחול והפעלת הבוט
        await application.initialize()
        await application.start()
        await application.updater.start_polling()

        print("Bot is running...")
        
        try:
            # לולאה אינסופית לשמירת הבוט פעיל
            while True:
                await asyncio.sleep(1)
        except Exception as e:
            print(f"Error in main loop: {e}")
        finally:
            # סגירה מסודרת של הבוט
            print("Shutting down...")
            await application.updater.stop()
            await application.stop()
            await application.shutdown()

if __name__ == '__main__':
    try:
        # הפעלת הפונקציה הראשית
        asyncio.run(main())
    except KeyboardInterrupt:
        print("\nBot stopped by user")
    except Exception as e:
        print(f"Critical error: {e}")
        import traceback
        traceback.print_exc()