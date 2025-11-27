# VBA

## 1. Ví dụ chuẩn cho code đẹp
```
Option Explicit
Public Sub SumDataClean()

    On Error GoTo ErrHandler

' Speed up

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

 ' >>> your main code here <<<

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Data")

    With ws
        .Range("A1").ClearContents
        .Range("B1:B10").Interior.Color = vbYellow
    End With

CleanExit:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Exit Sub

ErrHandler:
    MsgBox "Error: " & Err.Description
    Resume CleanExit

End Sub

```
## 2. Code cho last row
```
Dim lastRow As Long
lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
```

