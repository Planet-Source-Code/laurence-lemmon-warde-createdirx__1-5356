<div align="center">

## CreateDirX


</div>

### Description

This code will search the user's C:\ (or any other specified) drive for a given folder and if the folder is not found, it will call the CreateDirX function which in turn calls the API CreateDirectory function to create the specified folder. Once this is created, it will create a new notepad file within the folder and on each subsequent running of the application it will append info to this file. Great for logfile requirements!!
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Laurence Lemmon\-Warde](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/laurence-lemmon-warde.md)
**Level**          |Intermediate
**User Rating**    |4.2 (21 globes from 5 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Windows API Call/ Explanation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-api-call-explanation__1-39.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/laurence-lemmon-warde-createdirx__1-5356/archive/master.zip)

### API Declarations

```
Public Declare Function CreateDirectory Lib "kernel32" Alias _
"CreateDirectoryA" (ByVal lpPathname As String, lpSecurityAttributes _
As SECURITY_ATTRIBUTES) As Long
'Insert into global module
Type SECURITY_ATTRIBUTES
  nLength As Long
  lpSecurityDescriptor As Variant
  bInheritHandle As Boolean
End Type
```


### Source Code

```
'include a common dialog control on your form for this baby to work
Public Sub OpenLog()
Dim LogFile as integer
On Error GoTo exit1
 OpenLog.Flags = cdlOFNHideReadOnly Or cdlOFNExplorer
 OpenLog.CancelError = True
 OpenLog.FileName = "C:\JetLog\JET_LOG.log"  ' or whatever name grabs you by                       ' the nads
 temp = OpenLog.FileName
 Ret = Len(Dir$(temp))
 LogFile = FreeFile
 ' Open the log file.
 Open temp For Binary Access Write As LogFile
 If Err Then
  Exit Sub
 Else
  ' Go to the end of the file so that new data can be appended.
  Seek LogFile, LOF(LogFile) + 1
 End If
 Exit Sub
exit1:  ' Executes if folder is not found
 MsgBox "Application will create new directory 'C:\JetLog' on your hard drive." & vbCrLf & "Replace message with your own text.", vbExclamation, "Message"
 CreateDirX ("C:\JetLog")  'pass the path name you want to create in              ' these brackets
 OpenLog_Click
End Sub
Private Function CreateDirX(lpPathname As String) As Long
 Dim FYL As Long
 Dim DirC As SECURITY_ATTRIBUTES
  FYL = CreateDirectory(lpPathname, DirC)
End Function
```

