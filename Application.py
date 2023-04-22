from tkinter import *
import tkinter as tk
import tkinter.font as tkFont
import imaplib
import email
from tkinter import messagebox
import time
import mysql.connector
import pandas as pd

global mail, password, lb_email

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
                    temp = msg
                    if temp.is_multipart():
                        for part in temp.walk():
                            ctype = part.get_content_type()
                            cdispo = str(part.get('Content-Disposition'))

                            # skip to text/plain type
                            if ctype == 'text/plain' and 'attachment' not in cdispo:
                                body = part.get_payload()
                                break
                    else:
                        body = temp.get_payload()

                    email_textbox.insert(tk.END, f"\n{'#'*75}\nFrom: {sender}\nSubject: {subject}\n{'-'*50}\nBody: {body}\n")

        # mail.close()
        mail.logout()


root = tk.Tk()

root.title("Mail Automation")
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
root.mainloop()

