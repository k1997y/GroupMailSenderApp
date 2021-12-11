import win32com.client


def send_mail():
    outlook = win32com.client.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)

    mail.to = "k1997y@icloud.com"
    mail.subject = "test"
    mail.bodyFormat = 1
    mail.body = "test"


if __name__ == '__main__':
    send_mail()
