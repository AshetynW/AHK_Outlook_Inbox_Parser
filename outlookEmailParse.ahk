#SingleInstance Force
;#IfWinActive ahk_exe OUTLOOK.EXE

Loop ;comment to disable actions. Currently, this will loop continuously.
{
display_unread()	; Loop the function below
}
return


display_unread()
{
    oOutlook := ComObjActive("Outlook.Application")	; Initial com
      
    ; Open the inbox folder
    oNameSpace := oOutlook.GetNamespace("MAPI")	; This just has to be here
   ; folder := oNameSpace.GetDefaultFolder(olFolderInbox := 6)
folder := oNameSpace.GetDefaultFolder(6)			;.folders("CHANGEME")	; As it sits, this line will search through all undread inbox emails for Outlook. You can uncomment and change the folder name to search through specific folders if an account has a "rule" set up.

    ; Loop through items in the inbox folder
    ; oldest first
    For item In folder.Items
        If (item.Class = (olMail := 43 )) And (item.UnRead)
	    {
		    item.Display()	; Open the first unread 
				SetTitleMatchMode, 2
		Loop
		{
		ifwinexist, Message (HTML)	; Wait until the message is open to continue
			break
		}
	
Winwait, Message (HTML)		
; below are the outlook object's elements such as the sender's name and the subject lines.
{
objOL := ComObjActive("Outlook.Application").ActiveInspector.CurrentItem
From := objOL.SenderName
To := objOL.To	; From
CC := objOL.CC	; CC'd on it
BCC := objOL.BCC	; Blind CC'd on it
CreationTime := objOL.CreationTime	; Useful for date-specific stuff
Subject := objOL.Subject	; Subject
Body := objOL.Body		;The actual message
Msgbox %FROM% %to% %CC% %BCC% %CreationTime% %Subject% %Body%
}
}
}




