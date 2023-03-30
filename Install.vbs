Set ob = CreateObject("Wscript.Shell")

systemfile = "Scripting.FileSystemObject"
systemshell = "Wscript.Shell"
username = "%USERNAME%"
msgtext1 = "Hi "
msgtext2 = "Opening the specified files..."
msgtext3 = "Starting up"
msgtext4 = "Welcome back "
msgtext5 = "!"
msgtext6 = "Opening the specified files..."
msgtext7 = "Pc nap time: " 

Dim Result : Result = MsgBox("Do you want to install the sleep startup programm?", 4, "Installer")

If Result = 6 Then
    BrowseForFile()
    Restart()
End If

Function BrowseForFile()
    Dim Shell : Set Shell = CreateObject("Shell.Application")
    Dim File : Set File = Shell.BrowseForFolder(0, "Choose a folder:", &H4210)
    If File Is Nothing Then
        BrowseFolder = ""
    Else
        BrowseForFile = File.self.Path
        Dim Result : Result = MsgBox("Do you want the files contained in this folder to be executed whenever you wake up your computer?" & vbCrLf & "Folder Path: " & BrowseForFile, 4, "Question Prompt")
        
        If Result=6 Then
            MsgBox("Installing...")

            UsrPrfl = ob.expandenvironmentstrings("%UserProfile%")
            Dim FSO : Set FSO = CreateObject("Scripting.FileSystemObject")
            Set OutPutFile = FSO.CreateTextFile(UsrPrfl & "\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup\StartUp.vbs", True)
            
            OutPutFile.WriteLine("' Folder you want to run" & vbCrLf & _ 
            "UsrFolder = """ & BrowseForFile & """" & vbCrLf & _
            "If UsrFolder = """" Then" & vbCrLf & _
            "Else" & vbCrLf & _
            "Set FSO = CreateObject(""" & systemfile & """)" & vbCrLf & _
            "' Folder you want to run" & vbCrLf & _
            "Set Folder = FSO.GetFolder(""" & BrowseForFile & """)" & vbCrLf & _
            "End If" & vbCrLf & _
            "Set ob = CreateObject(""" & systemshell & """)" & vbCrLf & _
            "UsrNm = ob.expandenvironmentstrings(""" & username & """)" & vbCrLf & _
            "MsgBox """ & msgtext1 & """ & UsrNm & vbCrLf & _ " & vbCrLf & _
            """" & msgtext2 & """, 4096, """ & msgtext3 & """" & vbCrLf & _
            "If UsrFolder = """" Then" & vbCrLf & _
            "Else" & vbCrLf & _
            "For Each File in Folder.Files" & vbCrLf & _
            "ob.Run File.Path" & vbCrLf & _
            "Next" & vbCrLf & _
            "End If" & vbCrLf & _
            "Function GetDiff(Tm_a, Tm_b)" & vbCrLf & _
            "If (isDate(Tm_a) And IsDate(Tm_b)) = False Then" & vbCrLf & _
            "GetDiff = ""00:00:00""" & vbCrLf & _
            "Exit Function" & vbCrLf & _
            "End If" & vbCrLf & _
            "seconds = Abs(DateDiff(""S"", Tm_a, Tm_b))" & vbCrLf & _
            "minutes = seconds \ 60" & vbCrLf & _
            "hours = minutes \ 60" & vbCrLf & _
            "minutes = minutes Mod 60" & vbCrLf & _
            "seconds = seconds Mod 60" & vbCrLf & _
            "If Len(hours) = 1 Then hours = ""0"" & hours" & vbCrLf & _
            "GetDiff = hours & "":"" & Right(""00"" & minutes, 2) & "":"" & Right(""00"" & seconds, 2)" & vbCrLf & _
            "End Function" & vbCrLf & _
            "Do" & vbCrLf & _
            "a = Now" & vbCrLf & _
            "WScript.Sleep 5000" & vbCrLf & _
            "b = Now" & vbCrLf & _
            "If GetDiff(a, b) > ""00:00:15"" Then" & vbCrLf & _
            "' Code Block" & vbCrLf & _
            "UsrNm = ob.expandenvironmentstrings(""" & username & """)" & vbCrLf & _
            "MsgBox """ & msgtext4 & """ & UsrNm & """ & msgtext5 & """ & vbCrLf & _ " & vbCrLf & _
            """" & msgtext6 & """, 4096, """ & msgtext7 & """ & GetDiff(a, b)" & vbCrLf & _
            "If UsrFolder = """" Then" & vbCrLf & _
            "Else" & vbCrLf & _
            "For Each File in Folder.Files" & vbCrLf & _
            "ob.Run File.Path" & vbCrLf & _
            "Next" & vbCrLf & _
            "End If" & vbCrLf & _
            "' Code Block" & vbCrLf & _
            "End If" & vbCrLf & _
            "Loop")
            OutPutFile.Close()

            Set FSO = Nothing
        Else
            BrowseForFile()
        End If
    End If
End Function

Function Restart()
    Set ob = CreateObject("Wscript.Shell")

    dim Result : Result = MsgBox("Do you wish to restart your computer?" & vbCrLf & _ 
                    "This app will not work until you restart your computer",64 + 1,"Restart?")

    If Result = 1 Then
        UsrPrfl = ob.expandenvironmentstrings("%UserProfile%")
        Dim FSO : Set FSO = CreateObject("Scripting.FileSystemObject")
        Set OutPutFile = FSO.CreateTextFile(UsrPrfl & "\Desktop\restartcommand.cmd", True)
        OutPutFile.WriteLine("shutdown -g /t 3")
        OutputFile.Close()

        ob.Run UsrPrfl & "\Desktop\restartcommand.cmd"
        WScript.sleep 1000
        FSO.DeleteFile(UsrPrfl & "\Desktop\restartcommand.cmd")

        Set FSO = Nothing
    End If
End Function