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
            
            OutPutFile.WriteLine("' Folder you want to run")
            OutPutFile.WriteLine("UsrFolder = """ & BrowseForFile & """")
            OutPutFile.WriteLine("If UsrFolder = """" Then")
            OutPutFile.WriteLine("Else")
            OutPutFile.WriteLine("Set FSO = CreateObject(""" & systemfile & """)")
            OutPutFile.WriteLine("' Folder you want to run")
            OutPutFile.WriteLine("Set Folder = FSO.GetFolder(""" & BrowseForFile & """)")
            OutPutFile.WriteLine("End If")
            OutPutFile.WriteLine("Set ob = CreateObject(""" & systemshell & """)")
            OutPutFile.WriteLine("UsrNm = ob.expandenvironmentstrings(""" & username & """)")
            OutPutFile.WriteLine("MsgBox """ & msgtext1 & """ & UsrNm & vbCrLf & _ ")
            OutPutFile.WriteLine("""" & msgtext2 & """, 4096, """ & msgtext3 & """")
            OutPutFile.WriteLine("If UsrFolder = """" Then")
            OutPutFile.WriteLine("Else")
            OutPutFile.WriteLine("For Each File in Folder.Files")
            OutPutFile.WriteLine("ob.Run File.Path")
            OutPutFile.WriteLine("Next")
            OutPutFile.WriteLine("End If")
            OutPutFile.WriteLine("Function GetDiff(Tm_a, Tm_b)")
            OutPutFile.WriteLine("If (isDate(Tm_a) And IsDate(Tm_b)) = False Then")
            OutPutFile.WriteLine("GetDiff = ""00:00:00""")
            OutPutFile.WriteLine("Exit Function")
            OutPutFile.WriteLine("End If")
            OutPutFile.WriteLine("seconds = Abs(DateDiff(""S"", Tm_a, Tm_b))")
            OutPutFile.WriteLine("minutes = seconds \ 60")
            OutPutFile.WriteLine("hours = minutes \ 60")
            OutPutFile.WriteLine("minutes = minutes Mod 60")
            OutPutFile.WriteLine("seconds = seconds Mod 60")
            OutPutFile.WriteLine("If Len(hours) = 1 Then hours = ""0"" & hours")
            OutPutFile.WriteLine("GetDiff = hours & "":"" & Right(""00"" & minutes, 2) & "":"" & Right(""00"" & seconds, 2)")
            OutPutFile.WriteLine("End Function")
            OutPutFile.WriteLine("Do")
            OutPutFile.WriteLine("a = Now")
            OutPutFile.WriteLine("WScript.Sleep 5000")
            OutPutFile.WriteLine("b = Now")
            OutPutFile.WriteLine("If GetDiff(a, b) > ""00:00:15"" Then")
            OutPutFile.WriteLine("' Code Block")
            OutPutFile.WriteLine("UsrNm = ob.expandenvironmentstrings(""" & username & """)")
            OutPutFile.WriteLine("MsgBox """ & msgtext4 & """ & UsrNm & """ & msgtext5 & """ & vbCrLf & _ ")
            OutPutFile.WriteLine("""" & msgtext6 & """, 4096, """ & msgtext7 & """ & GetDiff(a, b)")
            OutPutFile.WriteLine("If UsrFolder = """" Then")
            OutPutFile.WriteLine("Else")
            OutPutFile.WriteLine("For Each File in Folder.Files")
            OutPutFile.WriteLine("ob.Run File.Path")
            OutPutFile.WriteLine("Next")
            OutPutFile.WriteLine("End If")
            OutPutFile.WriteLine("' Code Block")
            OutPutFile.WriteLine("End If")
            OutPutFile.WriteLine("Loop")
            OutPutFile.Close

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