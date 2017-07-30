Sub CustomInsert()
    'Click on a cell and launch the macro to insert a custom new line,
    'Clearing all sells that has no formula but the x first rows specified with info variable
    Dim wks As Worksheet
    Dim colrange As Range
    Dim LastCol As Long
    Dim info As Integer
    
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    
    'The "info" col umns mark the first column to be kept when copied
    info = 7

    'Check last populated column and copy/insert a new one below active cell
    Set wks = ActiveSheet
    LastCol = wks.Cells(1, wks.Columns.Count).End(xlToLeft).Column
    Set colrange = wks.Range(wks.Cells(ActiveCell.row, info), wks.Cells(ActiveCell.row, LastCol))
    
    ActiveCell.Rows("1:1").EntireRow.Select
    Selection.Copy
    Selection.Insert Shift:=xlDown
    Application.CutCopyMode = False
    
    'Clearing all but formulas and info
    For Each cell In colrange
        If cell.HasFormula() = False Then
            cell.ClearContents
        End If
    Next cell
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub