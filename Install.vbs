dim result
result = MsgBox("Do you want to install the sleep startup programm?", 4, "Installer")

If result=6 Then
BrowseForFile()
End If

Function BrowseForFile()
Dim Shell : Set Shell = CreateObject("Shell.Application")
Dim File : Set File = Shell.BrowseForFolder(0, "Choose a folder:", &H4210)
    If File Is Nothing Then
        BrowseFolder = ""
    Else
        BrowseForFile = file.self.Path
        result = MsgBox("Do you want the files contained in this folder to be executed whenever you wake up your computer?" & vbCrLf & "Folder Path:" & BrowseForFile, 4, "Question Prompt")
        If result=6 Then
            MsgBox("Installing...")

            Dim FSO
            Set FSO = CreateObject("Scripting.FileSystemObject")
            Set OutPutFile = FSO.CreateTextFile("C:\Users\%USERNAME%\Desktop\test.txt")
            OutPutFile.WriteLine("Writing text to a file")
            Set FSO = Nothing

        Else
            BrowseForFile()
        End If
    End If
End Function