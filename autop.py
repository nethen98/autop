import pandas as pd # type: ignore
from fpdf import FPDF # type: ignore
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from email.mime.text import MIMEText
import schedule # type: ignore
import time

# Load data from a CSV file
data = pd.read_csv('E:/CSS CODE/autoemail/data.csv')

# Perform data processing (example: calculate summary statistics)
summary = data.describe()

# Create a PDF report
class PDFReport(FPDF):
    def header(self):
        self.set_font('Helvetica', 'B', 12)
        self.cell(0, 10, 'Automated Report', 0, new_x='LMARGIN', new_y='NEXT', align='C')

    def footer(self):
        self.set_y(-15)
        self.set_font('Helvetica', 'I', 8)
        self.cell(0, 10, f'Page {self.page_no()}', 0, new_x='RIGHT', new_y='TOP', align='C')

    def chapter_title(self, title):
        self.set_font('Helvetica', 'B', 12)
        self.cell(0, 10, title, 0, new_x='LMARGIN', new_y='NEXT', align='L')
        self.ln(10)

    def chapter_body(self, body):
        self.set_font('Helvetica', '', 12)
        self.multi_cell(0, 10, body)
        self.ln()

pdf = PDFReport()
pdf.add_page()
pdf.chapter_title('Summary Statistics')
pdf.chapter_body(summary.to_string())
pdf.output('report.pdf')

# Function to send the email
def send_email(subject, body, to, file_path):
    from_email = "starnett11@gmail.com"
    from_password = "1173611739"  # Use your app-specific password here

    msg = MIMEMultipart()
    msg['From'] = from_email
    msg['To'] = to
    msg['Subject'] = subject
    msg.attach(MIMEText(body, 'plain'))

    attachment = open(file_path, "rb")
    part = MIMEBase('application', 'octet-stream')
    part.set_payload((attachment).read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', f"attachment; filename= {file_path}")
    msg.attach(part)

    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.starttls()
    server.login(from_email, from_password)
    text = msg.as_string()
    server.sendmail(from_email, to, text)
    server.quit()

    print("Email sent successfully.")

# Send the report via email
send_email("Automated Report", "Please find the attached report.", "starwar.n3@gmail.com", "report.pdf")

# Schedule the job
def job():
    # Include steps to load data, process data, generate report, and send email
    data = pd.read_csv('E:/CSS CODE/autoemail/data.csv')
    summary = data.describe()
    pdf = PDFReport()
    pdf.add_page()
    pdf.chapter_title('Summary Statistics')
    pdf.chapter_body(summary.to_string())
    pdf.output('report.pdf')
    send_email("Automated Report", "Please find the attached report.", "recipient@example.com", "report.pdf")
    print("Job completed.")

schedule.every().day.at("09:00").do(job)

while True:
    schedule.run_pending()
    time.sleep(1)
