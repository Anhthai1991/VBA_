# VBA

## Ví dụ chuẩn cho code đẹp
```
Option Explicit
Public Sub SumDataClean()
    On Error GoTo ErrHandler
    Application.ScreenUpdating = False

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Data")

    With ws
        .Range("A1").ClearContents
        .Range("B1:B10").Interior.Color = vbYellow
    End With

CleanExit:
    Application.ScreenUpdating = True
    Exit Sub

ErrHandler:
    MsgBox "Error in SumDataClean: " & Err.Description
    Resume CleanExit
End Sub
```
