Sub AngleChange()
    'Angle Button Sub called
    'Slight speed optimization
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False

    'Clear out old speeds incase old values yelded more angles
    Sheet10.Range(Sheet10.Cells(2, 2), Sheet10.Range("B2").SpecialCells(xlCellTypeLastCell)).Clear

    'Find end of sheet
    Dim LastRow, n As Integer
    LastRow = Sheet10.Range("A1").SpecialCells(xlCellTypeLastCell).Row

    'Try all the angle values
    For n = 2 To LastRow
        'Sets the speed to whatever is in the first row
        Sheet8.Cells(1, 1).Value = Sheet10.Cells(n, 1).Value
        Sheet7.Cells(1, 1).Value = Sheet10.Cells(n, 1).Value
        'Force Calculation, remember we turned it off
        Sheet7.Calculate
        Sheet8.Calculate
        Sheet9.Calculate
        Call TestTrue(n, 7, 10, 9)

    Next
    'Turn everything back on
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub