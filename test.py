
import imaplib
import email
import os
import yaml
from datetime import datetime, timedelta

# taking user credentials
user = input("Enter your email: ")
password = input("Enter your password: ")
data = {"user": user, "password": password}

#taking folder location
path = input("Enter folder location where attachments will be saved: ")

#taking file type to downlaod
print("1 : for xlsx type")
print("2 : for csv type")
print("3 : for any type")

choice = input("Select file type of email attachments: ")
selected_type = ''
match choice:
    case '1':
        selected_type = 'xlsx'
    case '2':
        selected_type = 'csv'
    case '3' :
        selected_type = 'any'


def download_attachment_automatically(data, path ,selected_type):
    try:
        # writing to a yaml file the user credential
        with open("credentials.yml", mode='w') as inputfile:
            yaml.dump(data, inputfile, default_flow_style=False)
            
        # reading credentials from yaml for login with IMAP lib to access email inbox
        with open("credentials.yml") as f:
            content = f.read()
        
        my_credentials = yaml.load(content, Loader=yaml.FullLoader)
        user, password = my_credentials["user"], my_credentials["password"]

        imap_url = 'imap.gmail.com'
        mail = imaplib.IMAP4_SSL(imap_url)
        mail.login(user, password)
        mail.select('Inbox')
        
        # date time
        today_date = datetime.now()
        yesterday_date = today_date - timedelta(days=5) 

        # subjects that we need to search in the inbox for attachment download
        SEARCH_SUBJECT = [
            f'SFA & Sunshine P1 P2 Calling Dashboard Till - {yesterday_date.strftime("%d-%b-%Y")}',
            f'SFA & Sunshine Dashboard for__{yesterday_date.strftime("%dth %b %Y")}'
            ]

        modified_path = ''

        for char in path:
            if char == '\\':
                modified_path += '\\\\'  # Append an extra backslash
            else:
                modified_path += char
        
        #saving the as SAVE_FOLDER after modification cause in imap we need to give path address with extra slashes
        # for eg - C:\Users\Ayush Singh\OneDrive\Desktop(original path by user)  --> C:\\Users\\Ayush Singh\\OneDrive\\Desktop (modified path for imap )
        
        SAVE_FOLDER = modified_path

        data = []
        for subject in SEARCH_SUBJECT:
            data.append(mail.uid('search', None, rf'(X-GM-RAW "subject:{subject}")'))

        # print(data) -- > [('OK', [b'29']), ('OK', [b'28'])]
        #data is coming up with result "Ok" and the id for the email that is having a subject we wanted
        
        selected_data = [item[1][0] for item in data]  #only the id of the emails having same subject we want

        email_ids_list = []
        for i in range(len(selected_data)):
            email_ids_list.append(selected_data[i].split())

        
        if email_ids_list:
            for email_ids in email_ids_list:
                if email_ids:
                    
                    for email_id in email_ids:
                        # print(email_id)  --> b'29' id of single email
                        # '(RFC822)': This is a search criteria parameter specified in the IMAP protocol. It indicates that you want to fetch the email in its entirety, including all headers and the body, in the RFC822 format, which is a standard format for representing email messages.
                        result, data = mail.fetch(email_id, '(RFC822)')
                        raw_email = data[0][1]
                        #raw_email is a bytes file
                        msg = email.message_from_bytes(raw_email)
                        
                        for part in msg.walk():
                            if part.get_content_maintype() == 'multipart':
                                continue
                            if part.get('Content-Disposition') is None:
                                continue

                            filename = part.get_filename()
                            
                            #selecting specifile file type to download
                            # print(type(filename))
                            revtype = ''
                            revfilename = filename[::-1]
                            for char in revfilename:
                                if char == '.':
                                    break;
                                else:
                                    revtype += char
                            
                            type = revtype[::-1]
                            
                            # print(type) 
                            
                            
                            #this is the part where if filename is got then it is sotring to the sprecific folder path we gave
                            
                            #downloads specific type of file given by user    
                            if type == selected_type:
                                 if filename:
                                    filepath = os.path.join(SAVE_FOLDER, filename)
                                    with open(filepath, 'wb') as f:
                                        f.write(part.get_payload(decode=True))
                                    print(f"Attachment saved: {filename}")
                            # else:
                            #     print("No emails found with the specified filet")
                                
                            #download any type of file
                            if selected_type == 'any' :
                                 if filename:
                                    filepath = os.path.join(SAVE_FOLDER, filename)
                                    with open(filepath, 'wb') as f:
                                        f.write(part.get_payload(decode=True))
                                    print(f"Attachment saved: {filename}")
                                
                            # if filename:
                            #     filepath = os.path.join(SAVE_FOLDER, filename)
                            #     with open(filepath, 'wb') as f:
                            #         f.write(part.get_payload(decode=True))
                            #     print(f"Attachment saved: {filename}")
                            # else:
                            #     print("No emails found with the specified subject")
                    
        else:
            print("No emails found with the specified subjects.")

        mail.close()
        mail.logout()

    except Exception as e:
        print(f"An error occurred: {e}")

# Call the function with provided parameters
download_attachment_automatically(data, path,selected_type)
