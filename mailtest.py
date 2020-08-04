import win32com.client as win32
outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'marcus.lam@dg-packaging.com'
mail.Subject = 'Python Mail Test'
message = 'THis is the body'
# mail.HTMLBody = '<h2>HTML Message body</h2>' #this field is optional

mail.GetInspector

index = mail.HTMLbody.find('>', mail.HTMLbody.find('<body'))
mail.HTMLbody = mail.HTMLbody[:index + 1] + message + mail.HTMLbody[index + 1:]

#mail.send #uncomment if you want to send instead of displaying

# To attach a file to the email (optional):
attachment = r"C:\Users\DGP-Yoga1\PycharmProjects\DGPInvoicing\dg_merge.ico"
mail.Attachments.Add(attachment)

# mail.Display(True)
mail.Send()


def sendmail(address, subject, body, attpath):
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = address
    mail.Subject = subject
    message = body
    mail.GetInspector
    index = mail.HTMLbody.find('>', mail.HTMLbody.find('<body'))
    mail.HTMLbody = mail.HTMLbody[:index + 1] + message + mail.HTMLbody[index + 1:]
    # To attach a file to the email (optional):
    attachment = attpath
    mail.Attachments.Add(attachment)

    # mail.Display(True)
    mail.Send()

