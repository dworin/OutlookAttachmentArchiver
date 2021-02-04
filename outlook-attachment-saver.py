import win32com.client
import win32com

#rootfolder is usually your outlook email address
rootfolder = "YOUREMAIL@YOURDOMAIN.COM"
#downloadfolder is the folder you want to clear/download from - usually 'Inbox'
downloadfolder = "FOLDER-TO-CHECK"
#You need to create this folder or use one that already exists
filepath = "C:\\temp\\attachments\\" 
#If True, you save the attachments before delting them.  Otherwise just delete them
saveattachments=False

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
accounts= win32com.client.Dispatch("Outlook.Application").Session.Accounts

archive  = outlook.Folders[rootfolder].Folders[downloadfolder]
messages = archive.Items
print("Saving Files")

attachmentcount = 0
for m in messages:
    for attachment in m.Attachments:
        try:
            if saveattachments == True:
                filename = f"{m.ReceivedTime} - {attachment.FileName}".replace(":","-")
                fileout = filepath + filename
                print(f"Saving: {fileout}")
                attachment.SaveAsFile(fileout)
            print(f"Deleting {attachment.FileName}")
            attachment.Delete()
            attachmentcount = attachmentcount+1
        except:
            print("Error Saving!")
    m.Save()

print(f"Deleted {attachmentcount} attachments!")
