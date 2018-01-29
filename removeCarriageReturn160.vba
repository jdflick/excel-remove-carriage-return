
Sub RemoveCarriageReturn160()
    Dim MyRange As Range
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
 
    For Each MyRange In ActiveSheet.UsedRange
        If 0 < InStr(MyRange, Chr(160)) Then
            MyRange = Replace(MyRange, Chr(160), "")
        End If
    Next
 
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
End Sub
