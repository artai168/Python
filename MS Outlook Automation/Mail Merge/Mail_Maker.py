# It is a Python project which can send email by batch with different email topics, content and attachments by changing the variables in the below JSON files: 
# mail_template.json
# mail_values.json
# mail_variables.json
# This script is used to generate Outlook emails. Through COM objects, it can automatically fill in the email's recipient, CC, subject, content, and attachments.
# This script requires the pywin32 library to run, which can be installed using pip install pywin32.
# This script requires the Microsoft Outlook client to run
# This script requires the "mail_template.json", "mail_values.json" and "mail_variables.json" files to be prepared in advance. Please look at the appendix of this article for the format and content of these three files.

import json
import os
import win32com.client as win32
import re

class Mail:
    def __init__(self, header, content, mail_to, mail_cc, attachments):
        self.header = header
        self.content = content
        self.mail_to = mail_to
        self.mail_cc = mail_cc
        self.attachments = attachments

def mail_content(in_contents):
    result = ""
    for content in in_contents:
        result += f'''
                    <p class=MsoNormal><span lang=EN-US><o:p>{content}</o:p></span></p>
                    <p class=MsoNormal><span lang=EN-US><o:p>&nbsp;</o:p>
                    '''
    return result

def files_exit(file_paths):
    result = True
    for file_path in file_paths:
        if os.path.exists(file_path):
            result = True
        else:
            result = False
            break
    return result

def create_email_via_com(text, subject, recipient, cc_parties, attachment_path):
    o = win32.Dispatch("Outlook.Application")
    ns = o.GetNamespace("MAPI")

    Msg = o.CreateItem(0)
    Msg.To = ";".join(recipient)
    Msg.CC = ";".join(cc_parties)

    Msg.Subject = subject
    # Add signature
    Msg.GetInspector  # This line is necessary, even if it doesn't seem to do anything. It initializes the body so that the signature can be added to the email.
    Msg.HTMLBody = Msg.HTMLBody.replace("<p class=MsoNormal><span lang=EN-US><o:p>&nbsp;</o:p></span></p><p class=MsoNormal><span lang=EN-US><o:p>&nbsp;</o:p></span></p>",text)   # Add the original HTMLBody (including the signature) to the end of your email body

    # Add attachments
    for attachment in attachment_path:
        if attachment != "":
            Msg.Attachments.Add(attachment)

    # Display the email without sending
      Msg.Display()


#------------- code to operate ----------------------------------
mail_template_object =[]
mail_value_object = []
mail_variable_object = []
str_path = os.path.dirname(os.path.abspath(__file__))

with open(str_path +'\\mail_variables.json', 'r') as json_file:
    mail_variable_object = json.load(json_file)

with open(str_path +'\\mail_values.json', 'r') as json_file:
    mail_value_object = json.load(json_file)

with open(str_path +'\\mail_template.json', 'r') as json_file:
    mail_template_object = json.load(json_file)

#request for code to cerate the mail object
# User is asked to input the number of the lamp post to be generated, separated by ',' or ';'
user_input = input("Please enter the number of the lamp post to be generated, separated by ',' or ';': ")
# The input is split by ',' or ';' and stored in _i_code
_i_code = re.split(',|;', user_input)

for i_code in _i_code:
    mail_value_object["item_code"] = str(i_code)
    mail_obj = []
    temp_header = mail_template_object["header"]
    temp_content = mail_template_object["content"]
    temp_attachments = mail_template_object["attachments"].copy()  # Create a new list for each iteration
        
    for key, value in mail_variable_object.items():
        temp_header = temp_header.replace(value, mail_value_object[key])
        temp_content = temp_content.replace(value, mail_value_object[key])
        # replace the attachment file name
        for i, attachment in enumerate(temp_attachments):
            temp_attachments[i] = attachment.replace(value, mail_value_object[key])
            #print(temp_attachments[i])
    temp_content = mail_content(temp_content.split("\n"))  # Convert the content to HTML format
    mail_obj.append(Mail(temp_header, temp_content, mail_value_object["mail_to"], mail_value_object["mail_cc"], temp_attachments))  # Add the new Mail object to the list

        # check mail object
    for mail in mail_obj:
        if(files_exit(mail.attachments)):
            create_email_via_com(mail.content, mail.header, mail.mail_to, mail.mail_cc, mail.attachments)
            print("Mail create successfully")
            for receiver in mail.mail_to:
                print("Mail to :" + receiver)
            for cc_receiver in mail.mail_cc:
                print("Mail cc :" + cc_receiver)
            print("Header :" + mail.header)
          
