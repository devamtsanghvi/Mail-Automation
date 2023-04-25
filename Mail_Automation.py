import html
from tkinter import *
import tkinter as tk
from tkinter import ttk
import tkinter.font as tkFont
import imaplib
import email
from tkinter import messagebox
from bs4 import BeautifulSoup
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

global mail, password, lb_email


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
            if(os.path.exists("update.zip")):
                os.remove("update.zip")
            with open("update.zip", "wb") as f:
                for chunk in r.iter_content(chunk_size=chunk_size):
                    if chunk:
                        zipfile_bytes.write(chunk)
                        print("Downloaded size: ", downloaded_size)
                        print("CHUNK SIZE: ", len(chunk))
                        print("TOTAL SIZE: ", total_size)
                        downloaded_size += len(chunk)
                        progress_pct = int(downloaded_size / total_size * 100)
                        progress_bar['value'] = progress_pct
                        progress_label.config(text=f"Downloading update ({progress_pct}%)")
                        print(f"Downloading update ({progress_pct}%")
                        new_dialog.update()

            #Remove ZIP file
            os.remove("update.zip")
            new_dialog.destroy()

            if(os.path.exists("update.exe")):
                os.remove("update.exe")

            # Extract the exe file from the release assets
            with zipfile.ZipFile(zipfile_bytes) as z:
                for filename in z.namelist():
                    if filename.endswith('mail_automation.exe'):
                        with open("update.exe", 'wb') as f:
                            f.write(z.read(filename))
                        break

            with open('run_update.bat', 'w') as f:
                f.write('start update.exe')

            # run batch file
            subprocess.call('run_update.bat', shell=True)

            os.remove("run_update.bat")
            # os.remove("update.exe")

            # exit app
            sys.exit(0)

# check_update()



def connection():
    conn = mysql.connector.connect(
    host="192.168.190.156",
    user="root",
    password="",
    database="backoffice"
    )
    
    return conn


# password = 'qvxzwnwygvubgdeh'
# password = 'eoppjzhrwrppnqza'
# password = 'ztvlpxhwogmcgpug'
# password = 'jgvhc6bit40kgoowcs4'

def submit():
    conn = connection()
    cur = conn.cursor()
    query = f"SELECT memail, apppassword FROM backoffice.clientemail where memail='{str(tb_email.get())}';"
    print(query)
    a = cur.execute(query)
    b = cur.fetchall()
    print(b)

    df = pd.read_sql(query, con=conn)

    lb_email = df.iloc[0, 0]
    password = df.iloc[0, 1]
    # s
    # getBrow = conbrowser(Customerid, Password, bankid)

    if tb_email.get() == '' :
            messagebox.showinfo("Empty Input Field", "Please enter Email ID!")
            
    else:
        email_address = tb_email.get().lower()
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
        frame.place(x=0, y=130,width=650, height=270)

        scrollbar = tk.Scrollbar(frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        email_textbox = tk.Text(frame, yscrollcommand=scrollbar.set)
        email_textbox.pack(expand=True, fill=tk.BOTH)
        scrollbar.config(command=email_textbox.yview)

        tempStr = tb_email.get().lower().split('@')[1]
        # print(tempStr)
        if tempStr == "gmail.com":
            host = 'imap.gmail.com'
        elif tempStr == "yahoo.com":
            host = 'imap.mail.yahoo.com'
        elif tempStr == "outlook.com":
            host = 'outlook.office365.com'
        elif tempStr == "rediffmail.com":
            host = 'mail.rediffmailpro.com'
        else:
            messagebox.showerror("Error", "Unsupported email domain")
            return

        mail = imaplib.IMAP4_SSL(host)
        # time.sleep(2)
        mail.login(email_address, password)
        
        res, messages = mail.select('INBOX')
        total_messages = int(messages[0])

        for i in range(total_messages, total_messages - email_count, -1):
            # RFC822 protocol
            res, msg = mail.fetch(str(i), "(RFC822)")
            for response in msg:
                if isinstance(response, tuple):
                    msg = email.message_from_bytes(response[1])

                    sender = msg["From"]
                    subject = msg["Subject"]

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

        # mail.close()
        mail.logout()


root = tk.Tk()

root.title("Mail Automation v1.2")
root.geometry("600x350")
root['background'] = '#5E2572'
width=650
height=400
screenwidth = root.winfo_screenwidth()
screenheight = root.winfo_screenheight()
alignstr = '%dx%d+%d+%d' % (width, height, (screenwidth - width) / 2, (screenheight - height) / 2)
root.geometry(alignstr)
root.resizable(width=False, height=False)


lb_email = tk.Label(root, text="Email ID")
ft = tkFont.Font(family='Times',size=12, weight='bold')
lb_email["bg"] = "#5E2572"
lb_email["font"] = ft
lb_email["fg"] = "#fff"
lb_email["justify"] = "center"
lb_email.place(x=120,y=20,width=70,height=30)

lb_count = tk.Label(root, text="Number of Emails")
ft = tkFont.Font(family='Times',size=12, weight='bold')
lb_count["bg"] = "#5E2572"
lb_count["font"] = ft
lb_count["fg"] = "#fff"
lb_count["justify"] = "center"
lb_count.place(x=60,y=80,width=139,height=30)

tb_email = tk.Entry(root)
tb_email["borderwidth"] = "1px"
ft = tkFont.Font(family='Times',size=11)
tb_email["font"] = ft
tb_email["fg"] = "#333333"
tb_email["justify"] = "center"
tb_email.place(x=200,y=20,width=275,height=30)

tb_count = tk.Entry(root)
tb_count["borderwidth"] = "1px"
tb_count.insert(END, '3')
ft = tkFont.Font(family='Times',size=12)
tb_count["font"] = ft
tb_count["fg"] = "#333333"
tb_count["justify"] = "center"
tb_count.place(x=200,y=80,width=61,height=30)

btn_submit = tk.Button(root, text="Submit", command=submit)
btn_submit["bg"] = "#FC7242"
btn_submit["activebackground"] = "#5E2572"
btn_submit["activeforeground"] = "#FC7242"
ft = tkFont.Font(family='Times',size=18, weight='bold')
btn_submit["font"] = ft
btn_submit["fg"] = "#fff"
btn_submit["justify"] = "center"
btn_submit.place(x=365,y=80,width=110,height=35)

root.bind("<Return>", lambda event=None: btn_submit.invoke())
root.attributes("-topmost",True)
root.mainloop()