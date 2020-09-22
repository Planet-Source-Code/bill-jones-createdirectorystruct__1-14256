<div align="center">

## CreateDirectoryStruct


</div>

### Description

Creates all non-existing folders in a path. Local or network UNC path.
 
### More Info
 
CreateThisPath as string


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Bill Jones](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/bill-jones.md)
**Level**          |Beginner
**User Rating**    |5.0 (25 globes from 5 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0, VBA MS Access
**Category**       |[Files/ File Controls/ Input/ Output](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/files-file-controls-input-output__1-3.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/bill-jones-createdirectorystruct__1-14256/archive/master.zip)





### Source Code

```
Private Sub CreateDirectoryStruct(CreateThisPath As String)
  'do initial check
  Dim ret As Boolean, temp$, ComputerName As String, IntoItCount As Integer, x%, WakeString As String
  Dim MadeIt As Integer
  If Dir$(CreateThisPath, vbDirectory) <> "" Then Exit Sub
  'is this a network path?
  If Left$(CreateThisPath, 2) = "\\" Then ' this is a UNC NetworkPath
    'must extract the machine name first, then get to the first folder
    IntoItCount = 3
    ComputerName = Mid$(CreateThisPath, IntoItCount, InStr(IntoItCount, CreateThisPath, "\") - IntoItCount)
    IntoItCount = IntoItCount + Len(ComputerName) + 1
    IntoItCount = InStr(IntoItCount, CreateThisPath, "\") + 1
    'temp = Mid$(CreateThisPath, IntoItCount, x)
  Else  ' this is a regular path
    IntoItCount = 4
  End If
  WakeString = Left$(CreateThisPath, IntoItCount - 1)
  'start a loop through the CreateThisPath string
  Do
    x = InStr(IntoItCount, CreateThisPath, "\")
    If x <> 0 Then
      x = x - IntoItCount
      temp = Mid$(CreateThisPath, IntoItCount, x)
    Else
      temp = Mid$(CreateThisPath, IntoItCount)
    End If
    IntoItCount = IntoItCount + Len(temp) + 1
    temp = WakeString + temp
    'Create a directory if it doesn't already exist
    ret = (Dir$(temp, vbDirectory) <> "")
    If Not ret Then
      'ret& = CreateDirectory(temp, Security)
      MkDir temp
    End If
    IntoItCount = IntoItCount  'track where we are in the string
    WakeString = Left$(CreateThisPath, IntoItCount - 1)
  Loop While WakeString <> CreateThisPath
End Sub
```

