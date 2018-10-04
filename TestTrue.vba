Sub TestTrue(CurrentRow As Integer, YIndex As Integer, AnalysisIndex As Integer, SweetIndex As Integer)

    Dim YSheet, XSheet, ZSheet As Worksheet
    Set YSheet = Application.ActiveWorkbook.Worksheets(YIndex)
    Set AnalysisSheet = Application.ActiveWorkbook.Worksheets(AnalysisIndex)
    Set SweetSheet = Application.ActiveWorkbook.Worksheets(SweetIndex)

    'Secondary sub
    'find and store the values of the last cell in worksheet


    Dim LastRow, LastColumn As Long
    LastRow = SweetSheet.Range("A1").SpecialCells(xlCellTypeLastCell).Row
    LastColumn = SweetSheet.Range("A1").SpecialCells(xlCellTypeLastCell).Column

    Dim TestCell, CheckRange As Range

    'Where are we checking T/F?
    Set CheckRange = SweetSheet.Range(SweetSheet.Cells(3, 2), SweetSheet.Cells(LastRow, LastColumn))

    Dim x, y As Long
    Dim theta As Double

    x = CurrentRow 'If you're looking for this, we fed it from the main sub
    y = 2 'What Column we're starting in
    theta = 0 'Helps intialize


    For Each TestCell In CheckRange
    'Look at all the cells on Sheet4 with t/f values
      If IsNumeric(TestCell.Value) Then
        If TestCell.Value = True Then
            theta = YSheet.Cells(2, TestCell.Column)

                AnalysisSheet.Cells(x, y).Value = theta
                y = y + 1

        End If
      End If
    Next
    Set CheckRange = AnalysisSheet.Range(AnalysisSheet.Cells(x, 2), AnalysisSheet.Cells(x, y))
    If y > 3 Then
        CheckRange.Sort Key1:=CheckRange, Order1:=xlAscending, Orientation:=xlLeftToRight
    End If
End Sub