Sub SpeedChange()
    'Main Sub called
    'Slight speed optimization
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False

    'Clear out old speeds incase old values yelded more angles
    Sheet5.Range(Sheet5.Cells(2, 2), Sheet5.Range("B2").SpecialCells(xlCellTypeLastCell)).Clear

    'Find end of sheet
    Dim LastRow, n As Integer
    LastRow = Sheet5.Range("A1").SpecialCells(xlCellTypeLastCell).Row

    'Try all the speed values
    For n = 2 To LastRow
        'Sets the speed to whatever is in the first row
        Sheet2.Cells(1, 1).Value = Sheet5.Cells(n, 1).Value
        Sheet3.Cells(1, 1).Value = Sheet5.Cells(n, 1).Value
        'Force Calculation, remember we turned it off
        Sheet2.Calculate
        Sheet3.Calculate
        Sheet4.Calculate
        Call TestTrue(n, 3, 5, 4)
        'Calls our other sub
    Next
    'Turn everything back on
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub