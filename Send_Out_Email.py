import os
import win32com.client as win32
import base64
import csv
from jinja2 import Environment, FileSystemLoader

def send_email(recipient,addressee,fee_paid,next_meeting):
    # Create Object
    outlook = win32.Dispatch('Outlook.Application')
    
    # Create mail
    mail_item = outlook.CreateItem(0)  

    mail_item.To = recipient
    
    # mail_item.CC = 'meekat@singtel.com;lixinlin@singtel.com'

    # Mail Settings
    mail_item.Subject = f'Welcome to Singtel Toastmasters Club'
    mail_item.BodyFormat = 2  # 2: Html format
    # mail_item.Attachments.Add('WFM_WorkItem_Result.xlsx')

    # Generate Graph String
    with open(f"welcome-email-pic.jpg", "rb") as image_file:
        poster_string = base64.b64encode(image_file.read()).decode('utf-8')

    ENV = Environment(loader=FileSystemLoader('.'))
    template = ENV.get_template("stc-welcome-email-20210327.htm")

    mail_item.HTMLBody = template.render(addressee = addressee,fee_paid = fee_paid, next_meeting = next_meeting, poster_string = poster_string)


    # mail_item.Attachments.Add('path and file')
    mail_item.Send()
    print(f'Email to {addressee} Sent Out')

    # Use the below function to test printing out the html - however it wont work with the base64 encoded jpg string
    # with io.open('final.html', 'w', encoding="utf-8") as f:
    #     f.write(mail_item.HTMLBody)

if __name__=="__main__":

    next_meeting = "7th April, 2021"

    with open('new_mem.txt','r') as f:
            csv_new_mem = csv.reader(f)
            next(csv_new_mem) # Skips the header row

            for row in csv_new_mem:
                send_email(row[1], row[0], row[2], next_meeting)
