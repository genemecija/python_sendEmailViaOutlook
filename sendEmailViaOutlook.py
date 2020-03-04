def sendEmail(text, subject, recipient):
    import win32com.client as win32   

    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = recipient
    mail.Subject = subject
    mail.HtmlBody = text
    mail.send

if __name__ == "__main__":
    sendEmail('Hello, World!', 'This is my subject.', 'ReceiversEmail@email.com')