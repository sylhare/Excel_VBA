Sub CustomInsert()
    'Click on a cell and launch the macro to insert a custom new line,
    'Clearing all sells that has no formula but the x first rows specified with info variable
    Dim wks As Worksheet
    Dim colRange As Range
    Dim LastCol As Long
    Dim info As Integer
    
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    Set wks = ActiveSheet
    
    'The "info" col umns mark the first column to be kept when copied
    info = 7

    'Check last populated column and copy/insert a new one below active cell
    LastCol = wks.Cells(1, wks.Columns.count).End(xlToLeft).Column
    Set colRange = wks.Range(wks.Cells(ActiveCell.row, 1), wks.Cells(ActiveCell.row, LastCol))
    
    If ActiveCell.row > 3 Then
        colRange.Copy
        colRange.Insert Shift:=xlDown
    
        'Clearing all but formulas and info
        For Each cell In colRange
            If cell.HasFormula() = False And cell.Column > info Then
                cell.ClearContents
            End If
        Next cell
    End If
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub