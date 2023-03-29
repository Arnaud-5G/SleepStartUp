' Folder you want to run
UsrFolder = "C:\Users\20200791\Desktop\!StartUpFiles"

If UsrFolder = "" Then 
Else
  Set FSO = CreateObject("Scripting.FileSystemObject")
  ' Folder you want to run
  Set Folder = FSO.GetFolder("C:\Users\20200791\Desktop\!StartUpFiles")
End If

Set ob = CreateObject("Wscript.Shell")

UsrNm = ob.expandenvironmentstrings("%Username%")
MsgBox "Hi " & UsrNm & vbCrLf & _
        "Opening the specified files...", 4096, "Starting up"

If UsrFolder = "" Then 
Else
  For Each File in Folder.Files
  ob.Run File.Path
  Next
End If

If UsrFile = "" Then
Else
ob.Run UsrFile
End If


Function GetDiff(Tm_a, Tm_b)
If (isDate(Tm_a) And IsDate(Tm_b)) = False Then
GetDiff = "00:00:00"
Exit Function
End If
seconds = Abs(DateDiff("S", Tm_a, Tm_b))
minutes = seconds \ 60
hours = minutes \ 60
minutes = minutes Mod 60
seconds = seconds Mod 60
If Len(hours) = 1 Then hours = "0" & hours
GetDiff = hours & ":" & Right("00" & minutes, 2) & ":" & Right("00" & seconds, 2)
End Function
Do
a = Now
WScript.Sleep 5000
b = Now

If GetDiff(a, b) > "00:00:15" Then

' Code Block

UsrNm = ob.expandenvironmentstrings("%Username%")
MsgBox "Welcome back " & UsrNm & "!" & vbCrLf & _
        "Opening the specified files...", 4096, "Pc nap time: " & GetDiff(a, b)

If UsrFolder = "" Then 
Else
  For Each File in Folder.Files
  ob.Run File.Path
  Next
End If

If UsrFile = "" Then
Else
ob.Run UsrFile
End If

' Code Block

End If
Loop