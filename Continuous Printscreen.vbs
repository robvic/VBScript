Set Wshshell=CreateObject("Word.Basic")
WshShell.sendkeys"%{prtsc}"
WScript.Sleep 1500

set Wshshell = WScript.CreateObject("WScript.Shell")
Wshshell.Run "mspaint"
WScript.Sleep 5000

WshShell.AppActivate "Paint"
WScript.Sleep 5000

WshShell.sendkeys "^v"
WScript.Sleep 1500

WshShell.sendkeys "^s"
WScript.Sleep 1500

WshShell.sendkeys "~"
WScript.Sleep 1500

WshShell.sendkeys "{LEFT}"
WScript.Sleep 1500

WshShell.sendkeys "~"
WScript.Sleep 1500

Const DestinationFile = "D:\Dropbox\Public\Untitled.png"
Const SourceFile = "C:\Users\Roberto\Desktop\Untitled.png"

Set fso = CreateObject("Scripting.FileSystemObject")
    'Check to see if the file already exists in the destination folder
    If fso.FileExists(DestinationFile) Then
        'Check to see if the file is read-only
        If Not fso.GetFile(DestinationFile).Attributes And 1 Then 
            'The file exists and is not read-only.  Safe to replace the file.
            fso.CopyFile SourceFile, "D:\Dropbox\Public\", True
        Else 
            'The file exists and is read-only.
            'Remove the read-only attribute
            fso.GetFile(DestinationFile).Attributes = fso.GetFile(DestinationFile).Attributes - 1
            'Replace the file
            fso.CopyFile SourceFile, "D:\Dropbox\Public\", True
            'Reapply the read-only attribute
            fso.GetFile(DestinationFile).Attributes = fso.GetFile(DestinationFile).Attributes + 1
        End If
    Else
        'The file does not exist in the destination folder.  Safe to copy file to this folder.
        fso.CopyFile SourceFile, "D:\Dropbox\Public\", True
    End If
Set fso = Nothing




