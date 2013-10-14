'Moves the selected messages to the archive folder
Sub Archive()
    Dim olApp As Outlook.Application
    Dim objNS As Outlook.NameSpace
    Dim olFolder As Outlook.MAPIFolder
    Dim msg As Outlook.MailItem

    Set olApp = Outlook.Application
    Set objNS = olApp.GetNamespace("MAPI")
    Set olFolder = objNS.GetDefaultFolder(olFolderInbox)
    Set ArchiveFolder = olFolder.Parent.Folders("Archive")

    For Each msg In ActiveExplorer.Selection
        msg.Move ArchiveFolder
        logMsg ("Archiving - " & msg)
    Next msg

End Sub

'Moves the selected messages to the inbox \ Star folder
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
        logMsg ("Staring - " & msg)
    Next msg

End Sub
Sub logMsg(logMsg As String)
    Dim sLogFileName As String, nFileNum As Long

    sLogFileName = "C:\temp\outlookVBA.log"
    logMsg = Now & " " & logMsg

    nFileNum = FreeFile                         ' next file number
    Open sLogFileName For Append As #nFileNum   ' create the file if it doesn't exist
    Print #nFileNum, logMsg                ' append information
    Close #nFileNum                             ' close the file
    
End Sub
