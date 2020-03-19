import win32com.client as win32

def mailMessage(send_to, subject, msg):
    try:
        outlook = win32.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)
    except:
        print("help")
        
    mail.To = send_to
    mail.Subject = subject
    mail.Body = msg
    mail.Send()
    
def txtMessage(send_to, subject, msg):
    mailMessage(send_to, subject, msg)
    

#attachment = "Path to the attachment" # (optional)
#mail.Attachments.Add(attachment)

if __name__ == "__main__":
    send_to = 'jbriggs@catiglass.com'
    subject = 'test'
    msg = 'test of the emergency broadcast system'
    mailMessage(send_to, subject, msg)
    send_to = '8159006943@txt.att.net'
    subject = 'test'
    msg = 'text test'
    txtMessage(send_to, subject, msg)