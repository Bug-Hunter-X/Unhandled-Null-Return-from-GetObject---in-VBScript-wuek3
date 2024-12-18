Function GetObject() is used to retrieve an object from the running object table. However, if the object is not found, it can return a null value which may not be handled properly in VBScript, leading to runtime errors.  For example:

```vbscript
Set objExcel = GetObject(, "Excel.Application")
If Err.Number <> 0 Then
    MsgBox "Excel is not running."
    Exit Sub
End If

' ... rest of the code that uses objExcel ...
```

This code checks for errors, but other code may not.  Also,  some objects might need to be explicitly created using CreateObject rather than retrieved with GetObject.