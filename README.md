<div align="center">

## McAfee DAT Download \- UPDATED\!


</div>

### Description

This code will automatically download and execute the McAfee SuperDAT. It will delete the old SuperDAT, copy the new SuperDAT from NAI's FTP server, and silently run the update.
 
### More Info
 
The code will run until the file has been downloaded to your computer and run.

Updated with the Function in it... Thanks for pointing that out!


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[JFV](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/jfv.md)
**Level**          |Intermediate
**User Rating**    |4.7 (14 globes from 3 users)
**Compatibility**  |VB 6\.0
**Category**       |[Files/ File Controls/ Input/ Output](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/files-file-controls-input-output__1-3.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/jfv-mcafee-dat-download-updated__1-43489/archive/master.zip)





### Source Code

```
Private Declare Function URLDownloadToFile Lib "urlmon" Alias _
"URLDownloadToFileA" (ByVal pCaller As Long, _
ByVal szURL As String, _
ByVal szFileName As String, _
ByVal dwReserved As Long, _
ByVal lpfnCB As Long) As Long
 Dim i As String
 Dim j As String
 Dim x As Integer
 Dim fso
 Dim fil
 Set fso = CreateObject("Scripting.FileSystemObject")
 Set fil = fso.OpenTextFile(App.Path & "\CopyLog.log", 8, True)
 j = "C:\1 Share\McAfee 4.5.1 Install Tools\SDAT\" 'Download location
 On Error Resume Next
 fso.DeleteFile j & "sdat*.exe", True 'Delete old SuperDAT file
 On Error GoTo ErrorFTP
 x = 4248
 Do While Now = Now 'Run continuously
 i = "ftp://ftp.nai.com/pub/antivirus/superdat/intel/sdat" & x & ".exe" 'Location and file to download from
 DoEvents 'Do Windows Events
 URLDownloadToFile 0, i, j & "sdat" & x & ".exe", 0, 0 'Download the file to your Download Location
 DoEvents
 s = "sdat" & x & ".exe"
 If fso.FileExists(j & s) Then 'If the file exists,
  Shell j & s & " -s" 'then run the SuperDAT in Silent Mode
  Set fso = Nothing 'Destroy Objects
  Set fil = Nothing
  Exit Sub 'Exit Sub or Unload Me
 End If
 x = x + 1 'Add another number and
 Loop 'loop to start again if not terminated.
ErrorFTP:
 fil.WriteLine "FTP Error: " & Err.Number & " " & Err.Description
 Resume Next
```

