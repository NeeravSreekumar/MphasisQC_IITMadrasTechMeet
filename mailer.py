import smtplib, ssl
import os

import pandas as pd

from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email import encoders

smtp_port = 587
smtp_server = 'smtp.gmail.com'

simple_email_context = ssl.create_default_context()

sender_email = "testingdivyam@gmail.com"
#Since the customers dont have a real email ID we are only emailing the sender however it is easy to substitute this for the actual customers email ID
receiver_email = "testingdivyam@gmail.com"
sender_pass = ""
parent_directory = os.path.dirname(os.path.abspath(__file__))
app_pass_path = os.path.join(parent_directory, '..', 'app_pass.txt')

#We will not be sharing the app API key for privacy purposes however it is easy to setup the GMail api for personal use
#Use the google process and replace the sender_email variables and make your own app_pass.txt file containing the key
with open("app_pass.txt") as file:
    read = file.read()
    sender_pass = read

html_msg = ""
with open("mail_template.html") as file:
    read = file.read()
    html_msg = read
def create_email_template(msg,details):
    msg = msg.replace("{name}", details["name"])
    msg = msg.replace("{reloc}", details["reloc"][0])
    msg = msg.replace("{datetime}", str(details["datetime"][0]))
    msg = msg.replace("{class}", str(details["class"][0]))
    msg = msg.replace("{cabin}", str(details["cabin"][0]))

    return msg


sol_filepath = "Solution_file.xlsx"
sol_file = pd.read_excel(sol_filepath, sheet_name=None)
relevant_df = pd.DataFrame(columns = ["id", "name", "email","reloc", "class", "cabin", "datetime", "arr"])

for sheet_name, sheet_data in sol_file.items():
    if "Passenger" in sheet_name:
        print(sheet_name)
        for index, row in sheet_data.iterrows():
            pass_id = row["DOC_ID"]
            reloc, class_, cabin, dt, arr = row[["RECLOC", "COS_CD", "Cabin", "DEP_DTMZ", "ARR_DTMZ"]]
            name = row["FIRST_NAME"] + " " + row["LAST_NAME"]
            email = row["CONTACT_EMAIL"]

            if str(pass_id) in relevant_df["id"].values:
                # Locate the row and update the lists
                idx = relevant_df[relevant_df["id"] == str(pass_id)].index[0]

                if reloc not in relevant_df.at[idx, "reloc"]:
                    # Update the lists
                    relevant_df.at[idx, "reloc"].append(reloc)
                    relevant_df.at[idx, "class"].append(class_)
                    relevant_df.at[idx, "cabin"].append(cabin)
                    relevant_df.at[idx, "datetime"].append(dt)
                    relevant_df.at[idx, "arr"].append(arr)
            else:
                details = {"id": [str(pass_id)],
                           "name": [name],
                           "email": [email],
                           "reloc": [[reloc]],
                           "class": [[class_]],
                           "cabin": [[cabin]],
                           "datetime": [[dt]],
                           "arr": [[arr]]}
                details_df = pd.DataFrame(details)
                relevant_df = pd.concat([relevant_df, details_df], ignore_index=True)
                

server = smtplib.SMTP(smtp_server,smtp_port)
server.starttls(context=simple_email_context) # Secure the connection
server.login(sender_email, sender_pass)

# Use the function to create the email template for each passenger
for index, row in relevant_df.iterrows():
    details = row.to_dict()
    email_template = create_email_template(html_msg, details)

    subject = "Rescheduling of flight"
    msg = MIMEMultipart()

    msg['From'] = sender_email
    msg['To'] = receiver_email
    msg['Subject'] = subject
    msg.attach(MIMEText(email_template, 'html'))

    server.sendmail(sender_email, receiver_email, msg.as_string())



server.quit() 