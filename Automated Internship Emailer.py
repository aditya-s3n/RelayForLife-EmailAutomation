import smtplib
import csv
from dotenv import load_dotenv
import os

from email.message import EmailMessage


#initialize .env file
load_dotenv()
#get email and password from .env file
NAME1 = "Aditya Sen"
NAME2 = "Aditya Bhatia"
NAME3 = "Steven Rowe"
NAME4 = "Sebastian Spernac"
NAME5 = 'David "Dodo" Rowe'
TEAM_NAME = "Dodowranglers"
EMAIL = os.getenv("EMAIL")
PASSWORD = os.getenv("PASSWORD")


#send email function (takes care of everything to send)
#can change content of cover letter and email body + subject
def send_email_companies(subject_email_address, subject_name) -> str:
    #change subject each specific company TODO Change subject name
    subject = f"The Canadian Cancer Society Needs your Help"
    #INFO OF EMAIL
    #set where credentials of email, TO, FROM, SUBJECT
    message = EmailMessage()
    message["From"] = f"{NAME1} {NAME2} {NAME3} {NAME4}"
    message["To"] = subject_email_address
    message["Subject"] = subject

    #BODY OF EMAIL
    #change content to specificy each company
    #Add email formatting and content
    message.set_content(
    f"""\
Hello {subject_name},

Cancer doesn't sleep, and neither will we. On June 3rd and 4th, we will run in relay for 12 consecutive hours in solidarity with cancer patients, survivors, and those we have lost.

Please help us reach our team fundraising goal of $1500. No matter how much you donate, or how insignificant is seems to you, every dollar goes a long way. Money donated will go towards cancer research, furthering our efforts to win the battle with cancer.

Together, we can do this. Help us to relay for life!

Donate at:
https://support.cancer.ca/site/TR/RelayForLife/RFLY_NW_odd_?team_id=489955&pg=team&fr_id=28128

---
{NAME1}, {NAME2}, {NAME3}, {NAME4}, {NAME5}
Oakville, ON
Iroquois Ridge High School
{TEAM_NAME}
    """)

    #send the email
    message = message.as_string()
    # Logs in and sends mail
    #NOTE after with statement STMP quits automatically 
    try:
        with smtplib.SMTP("smtp-mail.outlook.com", 587) as server:
            server.starttls() #starts TLS (security)

            server.login(EMAIL, PASSWORD) #login to my google gmail
            server.sendmail(EMAIL, subject_email_address, message) #send the message
    except:
        return "Unsuccessful"

    finally:
        return "Successful"


#array to check if all emails when through
email_check = []
#get company name and emails
#iterate through all the rows in the csv file
job_csv = open("Relay_for_life_-_Sheet2.csv")
job_csv.readline()
job_csv = csv.reader(job_csv)


#change company email to each specfic company
for data in job_csv:
    first_name = data[0] #first name
    last_name = data[1] #last name
    full_name = f"{data[0]} {data[1]}" #make full name
    
    subject_email = data[2] #get person's email

    temp_list = []
    temp_list.append(full_name)
    temp_list.append(send_email_companies(subject_email, full_name))
    

    email_check.append(temp_list)


#prints out how many emails were successfully sent
for email in email_check:
    print(f'{email[0]:<40}: {email[1]:>30}')