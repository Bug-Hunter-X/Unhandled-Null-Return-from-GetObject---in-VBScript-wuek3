```vbscript
On Error Resume Next
Set objExcel = GetObject(, "Excel.Application")
If Err.Number <> 0 Then
    MsgBox "Excel is not running. Please start Excel before continuing.", vbExclamation
    WScript.Quit
End If
On Error GoTo 0

' ... rest of the code that uses objExcel ...

'Example of safe usage:
If Not objExcel Is Nothing Then
  'Access Excel object here
  objExcel.Visible = True
  objExcel.Workbooks.Add
  ' ... rest of the code
Else
  MsgBox "Excel object is null, something went wrong!", vbCritical
End If

Set objExcel = Nothing
```