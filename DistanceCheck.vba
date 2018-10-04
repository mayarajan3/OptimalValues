Sub DistanceCheck()
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False

    Dim AngleSheet, DistanceSheet As Worksheet
    Set AngleSheet = Application.ActiveWorkbook.Worksheets(10)
    Set DistanceSheet = Application.ActiveWorkbook.Worksheets(11)
    Dim x, y, i, z As Integer
    Dim myMode As Variant
    y = 8

    For x = 0 To 18

        Sheet10.Range(Sheet10.Cells(2, 2), Sheet10.Cells(43, 701)).ClearContents

        Call AngleChange

        Sheet10.Range("ZZ2:ZZ43").Formula = "=COUNTIF(B2:ZY2, "">-1"")"
        Sheet10.Calculate

        Sheet10.Columns(702).Copy
        Sheet11.Columns(y).PasteSpecial (xlPasteValues)

        myMode = Application.WorksheetFunction.Mode_Mult(Sheet10.Range(Sheet10.Cells(2, 2), Sheet10.Cells(43, 701)))
        Sheet11.Range(Sheet11.Cells(45, y), Sheet11.Cells(49, y)).Value = myMode
        Sheet11.Cells(52, y).Value = Application.WorksheetFunction.Max(Sheet11.Range(Sheet11.Cells(2, y), Sheet11.Cells(43, y)))
        Sheet11.Cells(54, y).Value = Application.WorksheetFunction.Sum(Sheet11.Range(Sheet11.Cells(2, y), Sheet11.Cells(43, y)))

        Dim maxCell, ColumnRange As Range
        Set ColumnRange = Sheet11.Range(Sheet11.Cells(2, y), Sheet11.Cells(43, y))
        z = Application.WorksheetFunction.Max(Sheet11.Range(Sheet11.Cells(2, y), Sheet11.Cells(43, y)))
        For Each maxCell In ColumnRange
            If maxCell.Value = z Then
                Sheet11.Cells(51, y).Value = Sheet11.Cells(maxCell.Row, 1).Value
            End If
        Next


        If Sheet11.Cells(45, y).Value = Sheet11.Cells(46, y).Value Then

            Sheet10.Range("AAA2:AAA43").Formula = "=COUNTIF(B2:ZY2, ""=" & Sheet11.Cells(45, y).Value & """)"
            Sheet10.Columns(703).Copy
            Sheet11.Columns(y + 61).PasteSpecial (xlPasteValues)

            Sheet11.Cells(1, y + 61).Value = CStr(Sheet1.Cells(5, 2).Value) & ": " & CStr(Round(Sheet11.Cells(45, y).Value, 1))

            Sheet11.Cells(45, y + 61).Value = Application.WorksheetFunction.Sum(Sheet11.Range(Sheet11.Cells(2, y + 61), Sheet11.Cells(43, y + 61)))
            Sheet11.Cells(48, y + 61).Value = Application.WorksheetFunction.Max(Sheet11.Range(Sheet11.Cells(2, y + 61), Sheet11.Cells(43, y + 61)))

            Dim maxCellMode, ColumnRangeMode As Range
            Set ColumnRangeMode = Sheet11.Range(Sheet11.Cells(2, y + 41), Sheet11.Cells(43, y + 41))

            For Each maxCellMode In ColumnRangeMode
            If maxCellMode.Value = Sheet11.Cells(48, y + 41).Value Then
                Sheet11.Cells(47, y + 61).Value = Sheet11.Cells(maxCellMode.Row, 1).Value
            End If

        Next

        End If

        Sheet11.Cells(1, y).Value = Sheet1.Cells(5, 2).Value

        Sheet1.Cells(4, 2).Value = Sheet1.Cells(4, 2).Value - 0.1
        Sheet1.Cells(5, 2).Value = Sheet1.Cells(5, 2).Value - 0.1

        y = y + 1

    Next

    Sheet1.Cells(4, 2).Value = 3.25
    Sheet1.Cells(5, 2).Value = 3.85

    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic

End Sub