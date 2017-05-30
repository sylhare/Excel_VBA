Sub allColumns()
'Loop through all populated columns
    Dim wks As Worksheet
    Dim colRange As Range
    Dim LastCol As Long
    Dim count As Integer
    Dim msg As String
    
    ScreenUpdating = False
    Set wks = ActiveSheet

    count = 0
    
    LastCol = wks.Cells(1, wks.Columns.count).End(xlToLeft).Column
    'Take the first row and all the columns
    Set colRange = wks.Range(wks.Cells(1, 1), wks.Cells(1, LastCol))
    
    'Example loop to do something to each column
    For Each cell In colRange
        count = count + 1
    Next cell

    msg = count & " " & LastCol
    MsgBox (msg)
    ScreenUpdating = True
End Sub

Sub allRows()
'Loop through all populated rows
    Dim wks As Worksheet
    Dim rowRange As Range
    Dim LastRow As Long
    Dim count As Integer
    Dim msg As String
    
    ScreenUpdating = False
    Set wks = ActiveSheet
    
    LastRow = wks.Cells(wks.Rows.count, "A").End(xlUp).row
    Set rowRange = wks.Range("A1:A" & LastRow)
    
    count = 0
    
    For Each rrow In rowRange
        count = count + 1
    Next rrow
    
    msg = count & " " & LastRow
    MsgBox (msg)
    ScreenUpdating = True
End Sub

Sub allRows_simple()
'Loop through all populated rows
    Dim wks As Worksheet
    Dim row As Range
    Dim count As Integer
    ScreenUpdating = False
    Set sheet = ActiveSheet
    
    count = 0
    
    For Each row In sheet.Rows
        If sheet.Cells(row.row, 1).Value = "" Then
            Exit For
        End If

        count = count + 1

    Next row

    MsgBox (count)
    ScreenUpdating = True
End Sub
