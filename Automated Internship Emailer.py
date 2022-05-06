import smtplib
import csv
from dotenv import load_dotenv
import os

from email.message import EmailMessage
from docx import Document


#initialize .env file
load_dotenv()
#get email and password from .env file
NAME = "Aditya Sen"
EMAIL = os.getenv("EMAIL")
PASSWORD = os.getenv("PASSWORD")


#send email function (takes care of everything to send)
#can change content of cover letter and email body + subject
def send_email_companies(subject_email_address, subject_name) -> None:
    #change subject each specific company TODO Change subject name
    subject = f"Internship Oppurtunity - {subject_name}"
    #INFO OF EMAIL
    #set where credentials of email, TO, FROM, SUBJECT
    message = EmailMessage()
    message["From"] = NAME
    message["To"] = subject_email_address
    message["Subject"] = subject

    #BODY OF EMAIL
    #change content to specificy each company
    #Add email formatting and content
    message.set_content(
    f"""\
    Hello {subject_name},
    another line
    new conent

    ---
    {NAME}
    Oakville, ON
    Grade 11 High School Student @ Iroquois Ridge High School
    {EMAIL}
    """)

    #ATTACHING TO EMAIL
    #attach resume & cover letter to email
    #resume
    with open('Aditya Sen Resume - High School.pdf', 'rb') as content_file:
        content = content_file.read()
        message.add_attachment(content, maintype='application', subtype='pdf', filename='Aditya Sen Resume - High School.pdf')

    #cover letter -- creates new cover letter
    PDF_PATH = f'Custom Cover Letters/{NAME} Cover Letter for {subject_name}.docx'
    new_cover_letter = Document()
    #Add the cover letter material
    new_cover_letter.add_heading('Aditya Sen', 0)

    new_cover_letter.save(PDF_PATH)

    with open(PDF_PATH, 'rb') as content_file:
        content = content_file.read()
        message.add_attachment(content, maintype='application', subtype='docx', filename=f"{NAME}'s Cover Letter for {subject_name}.docx")


    #send the email
    message = message.as_string()
    # Logs in and sends mail
    #NOTE after with statement STMP quits automatically 
    with smtplib.SMTP("smtp.gmail.com", 587) as server:
        server.starttls() #starts TLS (security)

        server.login(EMAIL, PASSWORD) #login to my google gmail
        server.sendmail(EMAIL, subject_email_address, message) #send the message

    return "Successful"


#array to check if all emails when through
email_check = []
#get company name and emails
#iterate through all the rows in the csv file
job_csv = open("Job Email List.csv")
job_csv.readline()
job_csv = csv.reader(job_csv)
#change company email to each specfic company
for data in job_csv:
    subject_name_global = data[0] #company name

    if data[1] != '': #company email
        company_email = data[1]

        #send email & check if it went successfully through
        temp_list = []
        temp_list.append(subject_name_global)
        temp_list.append(send_email_companies(company_email, subject_name_global))
    else:
        #no email address to send to - Unsuccessful
        temp_list = []
        temp_list.append(subject_name_global)
        temp_list.append("Unsuccessful")

    

    email_check.append(temp_list)


#prints out how many emails were successfully sent
for email in email_check:
    print(f'{email[0]:<40}: {email[1]:>30}')