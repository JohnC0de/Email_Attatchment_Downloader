import win32com.client as win32
from datetime import datetime
import os

outlook = win32.Dispatch("Outlook.application").GetNamespace("MAPI")  # MAPI from Outlook

# Outlook Account Email
print(f"\n[🧾] Your logged Account: {outlook.Accounts[0].DeliveryStore.DisplayName}")

# Messages from a specific Folder Inside Inbox
folderName = 'Index'
inboxMessages = (outlook.GetDefaultFolder(6).Folders[folderName].Items)

# Filter variables Email, Subject, Date
emailSender = "example@gmail.com"
emailSubject = "TITULO"
fromDateRange = datetime.strptime("01/12/2019", "%m/%d/%Y").date()

try:
  for message in list(inboxMessages):
    if (message.SenderEmailAddress == emailSender and message.Subject == emailSubject and
        datetime.date(inboxMessages[0].ReceivedTime) > fromDateRange):
      print(
          f"\n[✅]\tEmail from: {message.SenderEmailAddress}\n\t Email Subject: {message.Subject}\n\t Received At: {datetime.date(inboxMessages[0].ReceivedTime)}\n\t Number of Attachments: {message.Attachments.Count}\n"
      )
      try:
        for attachment in message.Attachments:
          attachment.SaveASFile(os.path.join(os.getcwd(), "output_files", attachment.FileName))
          print(f"[📁] {attachment.FileName} from '{message.Sender}' downloaded with success!")
      except Exception as e:
        print("Error when saving the attachment:" + str(e))
    else:
      print(f"\n[❌] Message by '{message.Sender}' found, but did not match all fiters.")

except Exception as e:

  print("Error when processing emails messages:" + str(e))

print(f"\n[📦] All emails processed with success, downloaded files can be found at: {os.path.join(os.getcwd(), 'output_files')} \n")
