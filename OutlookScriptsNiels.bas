Attribute VB_Name = "NielsOutlookScripts"

Sub Archive()

Set ArchiveFolder = Application.GetNamespace("MAPI").GetDefaultFolder(olFolderInbox).Parent.Folders("Archive")

For Each msg In ActiveExplorer.Selection

msg.Move ArchiveFolder

Next msg

End Sub



Sub Star()

Dim olApp As Outlook.Application
Dim objNS As Outlook.NameSpace
Dim olFolder As Outlook.MAPIFolder
Dim msg As Outlook.MailItem

Set olApp = Outlook.Application
Set objNS = olApp.GetNamespace("MAPI")
Set olFolder = objNS.GetDefaultFolder(olFolderInbox)
Set olFolder = olFolder.Parent.Folders("Inbox")
Set ArchiveFolder = olFolder.Folders("@Star")

For Each msg In ActiveExplorer.Selection

msg.Move ArchiveFolder

Next msg

End Sub



