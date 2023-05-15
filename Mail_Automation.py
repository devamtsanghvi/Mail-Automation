from tkinter import *
import tkinter as tk
from tkinter import ttk
import tkinter.font as tkFont
import imaplib
import email
from tkinter import messagebox
from bs4 import BeautifulSoup
import re
import mysql.connector
import pandas as pd
import ver
import requests
import json
import base64
import os
import zipfile
import sys
import subprocess
import io
import shutil
import re

global mail, password, lb_email, var


#AUTO UPDATE THE APPLICATION USING THE RELEASE FEATURE IN GITHUB
#NOTE : PLEASE CHANGE THE RELEASE AFTER UPDATING THE VERSION OF APPLICATION
def check_update():

    # GETTING LATEST VERSION FROM GITHUB
    url = 'https://api.github.com/repos/{owner}/{repo}/contents/{path}'
    owner = 'devamtsanghvi'
    repo = 'auto_update'
    path = 'version_mail_automation.txt'
    params = {'ref': 'main'}

    # Send the GET request to retrieve the file content
    response = requests.get(url.format(owner=owner, repo=repo, path=path), params=params)
    content = json.loads(response.content)
    print(content)
    # Decode the base64-encoded content of the file
    file_content = base64.b64decode(content['content']).decode('utf-8')
    print(file_content)

    if ver.version == file_content:
        print("Latest Version")
    else:
        response = messagebox.askyesno('Update Available', 'An update is available. Do you want to download and install it?')
        if response == tk.YES:
            url1 = 'https://api.github.com/repos/{owner}/{repo}/releases/latest'
            # owner = 'zaladevdeep'
            # repo = 'hello_app'

            # Send the GET request to retrieve the release information
            response = requests.get(url1.format(owner=owner, repo=repo))
            release_info = json.loads(response.content)

            new_dialog = tk.Toplevel()
            new_dialog.geometry("300x100")
            new_dialog.title("Update Downloading")

            # Add a progress bar to the new dialog box
            progress_bar = ttk.Progressbar(new_dialog, orient="horizontal", length=200, mode="determinate")
            progress_bar.pack(pady=20)

            # Add a label to show the progress percentage
            progress_label = tk.Label(new_dialog, text="Downloading update (0%)")
            progress_label.pack()

            # Download the release assets
            r = requests.get(release_info['zipball_url'], stream=True)
            zipfile_bytes = io.BytesIO()
            total_size = int(r.headers.get('Content-Length', sys.maxsize)) # Set an arbitrary large value if Content-Length is not provided
            # chunk_size = 2048 # 2 kB
            chunk_size = 4194304 # 4 MB
            downloaded_size = 0

            # Get the path to the Documents folder on Windows
            documents_folder = os.path.join(os.path.expanduser("~"), "Documents")

            # Create a subfolder in the Documents folder to store the new version
            download_folder = os.path.join(documents_folder, "AGC TEMP")


            if(os.path.exists(download_folder + "\\update_mail.zip")):
                os.remove(download_folder + "\\update_mail.zip")
            with open(download_folder + "\\update_mail.zip", "wb") as f:
                for chunk in r.iter_content(chunk_size=chunk_size):
                    if chunk:
                        zipfile_bytes.write(chunk)
                        downloaded_size += len(chunk)
                        progress_pct = int(downloaded_size / total_size * 100)
                        progress_bar['value'] = progress_pct
                        progress_label.config(text=f"Downloading update ({progress_pct}%)")
                        new_dialog.update()

            #Remove ZIP file
            os.remove(download_folder + "\\update_mail.zip")
            new_dialog.destroy()

            if(os.path.exists(download_folder + "\\update_mail.exe")):
                os.remove(download_folder +"\\update_mail.exe")

            # Extract the exe file from the release assets
            with zipfile.ZipFile(zipfile_bytes) as z:
                for filename in z.namelist():
                    if filename.endswith('mail_automation.exe'):
                        with open(download_folder + "\\update_mail.exe", 'wb') as f:
                            f.write(z.read(filename))
                        break

            with open(download_folder + '\\run_update_mail.bat', 'w') as f:
                f.write('start update_mail.exe')

            

            # os.remove("update.exe")

            os.chdir(download_folder)
            # run batch file
            subprocess.call('run_update_mail.bat', shell=True)

            response = messagebox.askyesno('NOTE', 'Do you want to uninstall previous versions(Recommended)?')
            if response == tk.YES:
                folder_path1 = "D:/Mail Automation"

                if os.path.exists(folder_path1):
                    shutil.rmtree(folder_path1)

                folder_path = "C:/Mail Automation"

                if os.path.exists(folder_path):
                    shutil.rmtree(folder_path)
            
            os.remove("run_update_mail.bat")
            # exit app
            sys.exit(0)

check_update()



def connection():
    conn = mysql.connector.connect(
    host="192.168.190.156",
    user="root",
    password="",
    database="backoffice"
    )
    
    return conn



def func(msg,email_textbox,sender,subject):
    body = ""
    if msg.is_multipart():
        for part in msg.walk():
            ctype = part.get_content_type()
            cdispo = str(part.get('Content-Disposition'))

            if ctype == 'text/plain' and 'attachment' not in cdispo:
                body = part.get_payload(decode=True)
                charset = part.get_content_charset()
                if charset:
                    body = body.decode(charset)
                break
            elif ctype == 'text/html' and 'attachment' not in cdispo:
                body = part.get_payload(decode=True)
                charset = part.get_content_charset()
                if charset:
                    body = body.decode(charset)
                soup = BeautifulSoup(body, 'html.parser')
                body = soup.get_text(separator='\n')  # remove HTML tags
                break
    else:
        body = msg.get_payload(decode=True)
        charset = msg.get_content_charset()
        if charset:
            body = body.decode(charset)
        soup = BeautifulSoup(body, 'html.parser')
        body = soup.get_text(separator='\n')  # remove HTML tags

    email_textbox.insert(tk.END, f"\n{'#'*75}\nFrom: {sender}\nSubject: {subject}\n{'-'*50}\nBody: {body}\n")
# password = 'qvxzwnwygvubgdeh'
# password = 'eoppjzhrwrppnqza'
# password = 'ztvlpxhwogmcgpug'
# password = 'jgvhc6bit40kgoowcs4'


# password = ""

def login(value,password):
    cnt = 0

    email_address = tb_email.get().lower()
    email_address = email_address.replace("\n","")
    email_count = int(tb_count.get())

    # email_window = tk.Toplevel(root)
    # email_window.title("Emails")
    # email_window.geometry("700x500")

    # scrollbar = tk.Scrollbar(email_window)
    # scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

    # email_textbox = tk.Text(email_window, yscrollcommand=scrollbar.set)
    # email_textbox.pack(expand=True, fill=tk.BOTH)
    # scrollbar.config(command=email_textbox.yview)

    frame = tk.Frame(root)
    frame.place(x=0, y=90,width=650, height=310)

    scrollbar = tk.Scrollbar(frame)
    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

    email_textbox = tk.Text(frame, yscrollcommand=scrollbar.set)
    email_textbox.pack(expand=True, fill=tk.BOTH)
    scrollbar.config(command=email_textbox.yview)

    tempStr = email_address.lower().split('@')[1]
    print(tempStr)
    if tempStr == "gmail.com":
        host = 'imap.gmail.com'
        mail = imaplib.IMAP4_SSL(host)
        try:
            mail.login(email_address, password)
        except imaplib.IMAP4.error as e:
            messagebox.showinfo("",e)
        if var.get() == 1:
            res, messages = mail.select('Inbox')
        
        elif var.get() == 2:
            res, messages = mail.select('[Gmail]/Spam')

    elif tempStr == "yahoo.com" or tempStr=="yahoo.co.in" or tempStr=="yahoo.in":
        host = 'imap.mail.yahoo.com'
        mail = imaplib.IMAP4_SSL(host)
        try:
            mail.login(email_address, password)
        except imaplib.IMAP4.error as e:
            messagebox.showinfo("",e)
        if var.get() == 1:
            res, messages = mail.select('Inbox')
        
        elif var.get() == 2:
            res, messages = mail.select('Bulk')


    elif tempStr == "outlook.com":
        host = 'outlook.office365.com'
        mail = imaplib.IMAP4_SSL(host)
        try:
            mail.login(email_address, password)
        except imaplib.IMAP4.error as e:
            messagebox.showinfo("",e)
        if var.get() == 1:
            res, messages = mail.select('Inbox')
        elif var.get() == 2:
            res, messages = mail.select('Junk')


    elif tempStr == "rediffmail.com":
        host = 'mail.rediffmailpro.com'
        mail = imaplib.IMAP4_SSL(host)
        try:
            mail.login(email_address, password)
        except imaplib.IMAP4.error as e:
            messagebox.showinfo("",e)

        if var.get() == 1:
            res, messages = mail.select('Inbox')
        elif var.get() == 2:
            res, messages = mail.select('Spam')

    else:
        messagebox.showerror("Error", "Unsupported email domain")
        return

    # mail = imaplib.IMAP4_SSL(host)
    # time.sleep(2)
    # mail.login(email_address, password)

    total_messages = int(messages[0])
    
    print(total_messages)
    try:
        for i in range(total_messages, 1, -1):
            # RFC822 protocol
            res, msg = mail.fetch(str(i), "(RFC822)")
            for response in msg:
                if isinstance(response, tuple):
                    msg = email.message_from_bytes(response[1])

                    sender = msg["From"]
                    subject = str(msg["Subject"])

                    print("TYPE: ", type(subject))
                    sender_test = re.findall(r'[\w\.-]+@[\w\.-]+', sender)[0]

                    if value == 0:
                        print("All")
                        func(msg,email_textbox,sender,subject)
                        cnt += 1
                    elif value == 1:
                        if sender_test == tb_sender.get():
                            print("Sender matched")
                            func(msg,email_textbox,sender,subject)
                            cnt += 1

                    elif value == 2:
                        if tb_keyword.get().lower() in subject.lower():
                            print("Keyword FOUND")
                            func(msg,email_textbox,sender,subject)
                            cnt += 1

                    elif value == 3:
                        if tb_keyword.get().lower() in subject.lower() and sender_test == tb_sender.get():
                            print("Sender and Keyword FOUND")
                            func(msg,email_textbox,sender,subject)
                            cnt += 1
            if cnt == email_count:
                print("Count reached")
                break
    except Exception as e:
        messagebox.showinfo("Error",e)

        


                # index = body.find("OTP")
                # print(body[index+1])


                # if sender == 'KVB-no.reply@kvbmail.com' and subject == 'On-Demand Tokencode':
                #     match = re.search(r'\b\d{8}\b', body)
                #     if match:
                #         otp = match.group(0)
                #         print("OTP found:", otp)
                #     else:
                #         print("OTP not found")
                

    # mail.close()
    mail.logout()

def submit():
    # conn = connection()
    # cur = conn.cursor()
    # temp = tb_email.get().replace("\n","")
    # query = f"SELECT memail, apppassword FROM backoffice.clientemail where memail='{str(temp)}';"
    # print(query)
    # a = cur.execute(query)
    # b = cur.fetchall()
    # print(b)

    # df = pd.read_sql(query, con=conn)

    # lb_email = df.iloc[0, 0]
    # password = df.iloc[0, 1]
    password = "cmfwtfxmgffqxdxd"
    print(password)
    # getBrow = conbrowser(Customerid, Password, bankid)

    if tb_email.get() == '' :
            messagebox.showinfo("Empty Input Field", "Please enter Email ID!")
            
    elif tb_keyword.get() == '' and tb_sender.get() == '':
        login(0,password)

    elif tb_keyword.get() == '' and tb_sender.get() != '':
        login(1,password)
    
    elif tb_keyword.get() != '' and tb_sender.get() == '':
        login(2,password)
    
    elif tb_keyword.get() != '' and tb_sender.get() != '':
        login(3,password)

root = tk.Tk()

root.title(f"Mail Automation {ver.version}")
root.geometry("600x350")
root['background'] = '#5E2572'
width=650
height=400
screenwidth = root.winfo_screenwidth()
screenheight = root.winfo_screenheight()
alignstr = '%dx%d+%d+%d' % (width, height, (screenwidth - width) / 2, (screenheight - height) / 2)
root.geometry(alignstr)
root.resizable(width=False, height=False)

# Define fonts
ft_label = tkFont.Font(family='Times',size=10, weight='bold')
ft_entry = tkFont.Font(family='Times',size=9)
ft_radio = tkFont.Font(family='Times',size=11, weight='bold')
ft_button = tkFont.Font(family='Times',size=14, weight='bold')

# Define labels
lb_email = tk.Label(root, text="Email ID", font=ft_label, bg="#5E2572", fg="#fff", justify="center")
lb_count = tk.Label(root, text="Number of Emails", font=ft_label, bg="#5E2572", fg="#fff", justify="center")
lb_sender = tk.Label(root, text="Sender", font=ft_label, bg="#5E2572", fg="#fff", justify="center")
lb_keyword = tk.Label(root, text="Keyword", font=ft_label, bg="#5E2572", fg="#fff", justify="center")

# Define entries
tb_email = tk.Entry(root, borderwidth="1px", font=ft_entry, fg="#333333", justify="center")
tb_email.focus_set()
tb_count = tk.Entry(root, borderwidth="1px", font=ft_entry, fg="#333333", justify="center")
tb_count.insert(END, '3')
tb_sender = tk.Entry(root, borderwidth="1px", font=ft_entry, fg="#333333", justify="center")
tb_keyword = tk.Entry(root, borderwidth="1px", font=ft_entry, fg="#333333", justify="center")

# Define radio buttons
var = tk.IntVar()
var.set(1)
inbox_mail = tk.Radiobutton(root, text="INBOX", variable=var, value=1,
                            font=ft_radio,
                            selectcolor="#5E2572",
                            activebackground="#5E2572",
                            activeforeground="#FC7242",
                            bg="#5E2572",
                            fg="#fff",
                            justify="center")
spam_mail = tk.Radiobutton(root, text="SPAM", variable=var, value=2,
                           font=ft_radio,
                           selectcolor="#5E2572",
                           activebackground="#5E2572",
                           activeforeground="#FC7242",
                           bg="#5E2572",
                           fg="#fff",
                           justify="center")

# Define button
btn_submit = tk.Button(root, text="Submit", command=submit,
                       bg="#FC7242",
                       activebackground="#5E2572",
                       activeforeground="#FC7242",
                       font=ft_button,
                       fg="#fff",
                       justify="center")

# Place elements on the grid
# lb_email.grid(row=0,column=0,padx=(50,10),pady=(20,10))
# lb_count.grid(row=1,column=0,padx=(50,10),pady=(10))
# lb_sender.grid(row=2,column=0,padx=(50,10),pady=(10))
# lb_keyword.grid(row=3,column=0,padx=(50,10),pady=(10))
# tb_email.grid(row=0,column=1,padx=(10),pady=(20,10),sticky='ew')
# tb_count.grid(row=1,column=1,padx=(10),pady=(10),sticky='ew')
# tb_sender.grid(row=2,column=1,padx=(10),pady=(10),sticky='ew')
# tb_keyword.grid(row=3,column=1,padx=(10),pady=(10),sticky='ew')
# inbox_mail.grid(row=0,column=2,padx=(50),pady=(20))
# spam_mail.grid(row=1,column=2,padx=(50))
# btn_submit.grid(row=3,column=2,padx=(50),pady=(10))



# root.geometry("600x350")

# Widget: lb_email
lb_email.grid(row=1, column=1, padx=(10, 3), pady=(5, 3))

# Widget: lb_count
lb_count.grid(row=2, column=1, padx=(10, 3), pady=(3, 3))

# Widget: tb_email
tb_email.grid(row=1, column=2, padx=(3), pady=(5, 3), sticky='ew')

# Widget: tb_count
tb_count.grid(row=2, column=2, padx=(3), pady=(3, 3), sticky='ew')

# Widget: lb_sender
lb_sender.grid(row=1, column=3, padx=(10, 3), pady=(5, 3))

# Widget: lb_keyword
lb_keyword.grid(row=2, column=3, padx=(10, 3), pady=(3, 3))

# Widget: tb_sender
tb_sender.grid(row=1, column=4, padx=(3), pady=(5, 3), sticky='ew')

# Widget: tb_keyword
tb_keyword.grid(row=2, column=4, padx=(3), pady=(3, 3), sticky='ew')

# Widget: inbox_mail
inbox_mail.grid(row=1, column=5, padx=(10), pady=(5))

# Widget: spam_mail
spam_mail.grid(row=2, column=5, padx=(10))

# Widget: btn_submit
btn_submit.grid(row=2, column=6, padx=(10), pady=(3))







root.bind("<Return>", lambda event=None: btn_submit.invoke())
# root.attributes("-topmost",True)
root.mainloop()