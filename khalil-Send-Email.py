import sys, os, smtplib, ssl, base64
from PyQt6.QtWidgets import (
    QApplication, QWidget, QLabel, QLineEdit, QPushButton, QTextEdit,
    QFileDialog, QVBoxLayout, QHBoxLayout, QMessageBox, QFrame, QComboBox
)
from PyQt6.QtGui import QFont
from PyQt6.QtCore import Qt
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders

CRED_FILE = "credentials.txt"

SMTP_SERVERS = {
    "Zoho": ("smtp.zoho.com", 587),
    "Gmail": ("smtp.gmail.com", 587),
    "Outlook": ("smtp.office365.com", 587),
    "Yahoo": ("smtp.mail.yahoo.com", 587),
}

# --- دوال التخزين ---
def save_credentials(email, password):
    data = f"{email}:{password}"
    encoded = base64.b64encode(data.encode("utf-8"))
    with open(CRED_FILE, "wb") as f:
        f.write(encoded)

def load_credentials():
    if os.path.exists(CRED_FILE):
        with open(CRED_FILE, "rb") as f:
            encoded = f.read()
        decoded = base64.b64decode(encoded).decode("utf-8")
        email, password = decoded.split(":", 1)
        return email, password
    return None, None

class ProfessionalEmailApp(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Email Sender - Professional")
        self.setGeometry(100, 100, 750, 600)
        self.setStyleSheet(self.get_stylesheet())
        self.initUI()

    def get_stylesheet(self):
        return """
        QWidget {
            background-color: #2b2b2b;
            color: #ffffff;
            font-family: 'Segoe UI', Arial;
            font-size: 14px;
        }
        QLineEdit, QTextEdit, QComboBox {
            background-color: #3c3f41;
            border: 2px solid #555555;
            border-radius: 8px;
            padding: 6px;
            color: #ffffff;
        }
        QPushButton {
            background-color: #009688;
            border: none;
            border-radius: 10px;
            padding: 10px;
            color: white;
            font-weight: bold;
        }
        QPushButton:hover {
            background-color: #00bfa5;
        }
        QLabel {
            font-weight: bold;
            color: #ffffff;
        }
        QTextEdit#status_text {
            background-color: #1e1e1e;
            border: 2px solid #555555;
            border-radius: 8px;
        }
        QFrame#line {
            background-color: #555555;
            max-height: 2px;
        }
        """

    def initUI(self):
        layout = QVBoxLayout()
        layout.setSpacing(10)

        # مزود الخدمة
        layout.addWidget(QLabel("اختر مزود البريد:"))
        self.provider_box = QComboBox()
        self.provider_box.addItems(SMTP_SERVERS.keys())
        layout.addWidget(self.provider_box)

        # البريد وكلمة السر
        layout.addWidget(QLabel("البريد:"))
        self.email_input = QLineEdit()
        layout.addWidget(self.email_input)

        layout.addWidget(QLabel("App Password / كلمة السر:"))
        self.pass_input = QLineEdit()
        self.pass_input.setEchoMode(QLineEdit.EchoMode.Password)
        layout.addWidget(self.pass_input)

        layout.addWidget(QLabel("الاسم الظاهر:"))
        self.name_input = QLineEdit()
        layout.addWidget(self.name_input)

        # خط فاصل
        line = QFrame()
        line.setFrameShape(QFrame.Shape.HLine)
        line.setObjectName("line")
        layout.addWidget(line)

        # البريد المستلم وموضوع الرسالة
        layout.addWidget(QLabel("البريد المستلم:"))
        self.to_input = QLineEdit()
        layout.addWidget(self.to_input)

        layout.addWidget(QLabel("موضوع البريد:"))
        self.subject_input = QLineEdit()
        layout.addWidget(self.subject_input)

        layout.addWidget(QLabel("محتوى الرسالة (HTML مسموح):"))
        self.body_input = QTextEdit()
        self.body_input.setAcceptRichText(True)
        layout.addWidget(self.body_input)

        # المرفقات
        h_layout_attach = QHBoxLayout()
        self.attach_input = QLineEdit()
        self.attach_btn = QPushButton("اختر ملفات")
        self.attach_btn.clicked.connect(self.select_attachments)
        h_layout_attach.addWidget(self.attach_input)
        h_layout_attach.addWidget(self.attach_btn)
        layout.addWidget(QLabel("مرفقات:"))
        layout.addLayout(h_layout_attach)

        # زر الإرسال
        self.send_btn = QPushButton("إرسال البريد")
        self.send_btn.setFixedHeight(50)
        self.send_btn.setFont(QFont("Arial", 14, QFont.Weight.Bold))
        self.send_btn.clicked.connect(self.send_email)
        layout.addWidget(self.send_btn)

        # نافذة حالة الإرسال
        layout.addWidget(QLabel("حالة الإرسال:"))
        self.status_text = QTextEdit()
        self.status_text.setObjectName("status_text")
        self.status_text.setReadOnly(True)
        layout.addWidget(self.status_text)

        self.setLayout(layout)

        # تحميل البريد وكلمة السر المحفوظة
        saved_email, saved_pass = load_credentials()
        if saved_email:
            self.email_input.setText(saved_email)
        if saved_pass:
            self.pass_input.setText(saved_pass)

    def select_attachments(self):
        paths, _ = QFileDialog.getOpenFileNames(self, "اختر الملفات")
        if paths:
            self.attach_input.setText(", ".join(paths))

    def send_email(self):
        provider = self.provider_box.currentText()
        SMTP_HOST, SMTP_PORT = SMTP_SERVERS[provider]

        sender_email = self.email_input.text().strip()
        password = self.pass_input.text().strip()
        display_name = self.name_input.text().strip()
        receiver = self.to_input.text().strip()
        subject = self.subject_input.text().strip()
        body = self.body_input.toHtml()
        attachments_files = self.attach_input.text().split(", ") if self.attach_input.text() else []

        if not sender_email or not password or not receiver or not body:
            QMessageBox.critical(self, "خطأ", "الرجاء ملء جميع الحقول المطلوبة.")
            return

        save_credentials(sender_email, password)

        attachments = []
        for file in attachments_files:
            if os.path.exists(file):
                with open(file, "rb") as f:
                    attachments.append((os.path.basename(file), f.read()))

        context = ssl.create_default_context()
        try:
            server = smtplib.SMTP(SMTP_HOST, SMTP_PORT)
            server.ehlo()
            server.starttls(context=context)
            server.login(sender_email, password)
        except Exception as e:
            QMessageBox.critical(self, "خطأ", f"فشل الاتصال بخادم البريد: {e}")
            return

        if not body.strip().startswith("<html>"):
            body = f"""
            <html>
              <body style="font-family: Arial; line-height: 1.5;">
                <p>{body.replace('\n', '<br>')}</p>
              </body>
            </html>
            """

        msg = MIMEMultipart()
        msg["From"] = f"{display_name} <{sender_email}>"
        msg["To"] = receiver
        msg["Subject"] = subject if subject else "رسالة جديدة"
        msg.attach(MIMEText(body, "html"))

        for fname, content in attachments:
            part = MIMEBase("application", "octet-stream")
            part.set_payload(content)
            encoders.encode_base64(part)
            part.add_header("Content-Disposition", f"attachment; filename={fname}")
            msg.attach(part)

        try:
            server.sendmail(sender_email, receiver, msg.as_string())
            self.status_text.append(f"<span style='color:#00ff00'>✅ تم الإرسال عبر {provider} إلى: {receiver}</span>")
        except Exception as e:
            self.status_text.append(f"<span style='color:#ff5555'>❌ خطأ عند الإرسال: {e}</span>")
        finally:
            server.quit()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = ProfessionalEmailApp()
    window.show()
    sys.exit(app.exec())
