# OutlookAttachmentArchiver

This script will download and then delete all your outlook attachments to cut down your mailbox size

*Notes Before You Get Started*
-----------------------------
You must install pywin32:
  pip install pywin32

*WARNING: THIS SCRIPT WILL DELETE ALL YOUR ATTACHMENTS.  IT IS IRREVERSABLE.  USE AT YOUR OWN RISK*

Key variables to modify:
* _rootfolder_ is your root folder in outlook (usually your email address)
* _downloadfolder_ is the folder you want to save attachments from (I usually create an archive folder, but in theory it could be Inbox)
* _filepath_ is where you want to save the attachments
* _saveattachments_ When set to True this will save attachments before deleting, otherwise it will just delete all the attachments in the folder

For some reason, not all the attachments get deleted every round.  You may have to run the script a few times
