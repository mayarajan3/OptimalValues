Sub AngleCheck()
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    Dim x, y, i, z, k As Integer
    Dim count As Integer
    Dim myMode As Variant
    y = 45
    k = 702
    For x = 0 To 1
        Call AngleChange
        myMode = Application.WorksheetFunction.Mode_Mult(Sheet10.Range(Sheet10.Cells(18, 2), Sheet10.Cells(18, 701)))
        Sheet10.Cells(y, 1).Value = CStr(Sheet1.Cells(5, 2).Value)
        Sheet10.Cells(y, 2).Value = myMode
        count = Application.WorksheetFunction.CountIf(Sheet10.Range(Sheet10.Cells(18, 2), Sheet10.Cells(18, 701)), Sheet10.Cells(y, 2).Value)
        Sheet10.Cells(y, 3).Value = count
        y = y + 1
        Sheet1.Cells(4, 2).Value = Sheet1.Cells(4, 2).Value - 0.1
        Sheet1.Cells(5, 2).Value = Sheet1.Cells(5, 2).Value - 0.1
        z = 49
        For i = 0 To 41
            Sheet10.Cells(i + 2, 1).Value = z + i
        Next
    Next
    Sheet1.Cells(4, 2).Value = 3.25
    Sheet1.Cells(5, 2).Value = 3.85

End Sub