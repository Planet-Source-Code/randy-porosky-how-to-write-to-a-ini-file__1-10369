<div align="center">

## How To Write To a INI File


</div>

### Description

This VERY Simple 5 Line Code Writes To a INI File! Enjoy!
 
### More Info
 
Place a Command Button on The Form, Name it Command1. Put The 1 Declaration in a Module. Have Fun!


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Randy Porosky](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/randy-porosky.md)
**Level**          |Beginner
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Files/ File Controls/ Input/ Output](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/files-file-controls-input-output__1-3.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/randy-porosky-how-to-write-to-a-ini-file__1-10369/archive/master.zip)

### API Declarations

```
Declare Function WritePrivateProfileString Lib "kernel32" _
Alias "WritePrivateProfileStringA" (ByVal lpApplicationName _
As String, ByVal lpKeyName As Any, ByVal lpString As Any, _
ByVal lpFileName As String) As Long
```


### Source Code

```
Public Sub iniWrite(sFileName As String, sKey As String, sSection As String, ByVal sValue As String)
Dim iW As String
iW = WritePrivateProfileString(sSection, sKey, sValue, sFileName)
End Sub
Private Sub Command1_Click()
iniWrite "C:\Windows\Desktop\File.ini", "Neat", "Pretty", "Huh?"
End Sub
'Have Fun!
```

