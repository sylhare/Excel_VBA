Sub MigrateCustomerList()

Dim i As Integer, j As Integer, n As Integer
Dim yy As Integer, mm As Integer, d As Date, pass As Boolean

    yy = 2016
    mm = 1
    pass = False

    j = 6 'Column of result
    i = 2   'Row of Customer


Range(Cells(nr, 1), Cells(500, 200)).ClearContents
    With Range(Cells(nr, 1), Cells(500, 200)).Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With


    'For time efficiency
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False

    AscendingSort (c + 1)
    
    For i = 2 To nb
        
        n = nr + 1
        If (IsDate(Cells(i, c + 1).Value) = True) Then
        
            'Migration date
            mm = Month(Cells(i, c + 1).Value)
            yy = Year(Cells(i, c + 1).Value)
            d = DateSerial(yy, mm, 1)
            Cells(n, j).Value = Format(d, "yyyy-mmmm")
            Cells(n, j).Font.Bold = True
       
            
            'List customer by Month of Migration date
            While IsDate(Cells(i, c + 1).Value) = True And pass = False
                If Month(Cells(i, c + 1).Value) = mm Then
                    n = n + 1
                    Fill i:=i, j:=j, c:=c, n:=n
                    i = i + 1
                Else
                    pass = True
                End If
            Wend
            i = i - 1 'Miss one with the while and the for
                
            pass = False
            count j:=j, n:=n
        Else
            List Condition:=Cells(i, c + 1).Value, i:=i, j:=j, c:=c, n:=n
        End If
          
    Next i
    
    'Sum all the customer ready to migrate
    Cells(nr, 5).Value = WorksheetFunction.Sum(Range(Cells(nr, 6), Cells(nr, j - 6)))
         
    AscendingSort (c)
    'ActiveWindow.SmallScroll Down:=264
    
    'To set it back to normal
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
End Sub

Sub AscendingSort(column As Integer)

    ActiveWorkbook.Worksheets("Matrix").AutoFilter.Sort.SortFields. _
        Clear
    ActiveWorkbook.Worksheets("Matrix").AutoFilter.Sort.SortFields. _
        Add Key:=Range(Cells(2, column), Cells(nb, column)), SortOn:=xlSortOnValues, Order:=xlAscending, _
        DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Matrix").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

End Sub

Function List(Condition As String, i As Integer, j As Integer, c As Integer, n As Integer)

    Cells(n, j).Font.Bold = True
    If Condition = "" Then
        Cells(n, j).Value = "Not Scheduled"
    Else
        Cells(n, j).Value = Condition
    End If
        
    While Cells(i, c + 1).Value = Condition
        n = n + 1
        Fill i:=i, j:=j, c:=c, n:=n
        i = i + 1
    Wend
    count j:=j, n:=n
    i = i - 1 'while and for

End Function

Function Fill(i As Integer, j As Integer, c As Integer, n As Integer)
    
    With Cells(n, j)
        .Font.Bold = False
        .Value = Cells(i, c).Value
        .Interior.Color = RGB(22, 22, 22)
        .Interior.TintAndShade = 0.9
    End With
    

End Function

'Add the count of the results
Function count(j As Integer, n As Integer)

    With Cells(nr, j)
        .Font.Bold = True
        .Value = n - (nr + 1) 'Count the number
        .NumberFormat = "General"
        .HorizontalAlignment = xlLeft
    End With
    j = j + 1 'New Column of result
            
End Function

Sub nb_days_month()
   
    'Any date will do for this example
    date_test = CDate("6/2/2012")
   
    'Month / Year of the date
    var_month = Month(date_test)
    var_year = Year(date_test)
   
    'Calculation for the first day of the following month
    date_next_month = DateSerial(var_year, var_month + 1, 1)
   
    'Date of the last day
    last_day_month = date_next_month - 1
   
    'Number for the last day of month (= last day)
    nb_days = Day(last_day_month)
    
End Sub

Sub test()
'
' Macro
'
    With Selection.Font
        .Name = "Arial"
        .Size = 7
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    Range("CJ4").Select
End Sub
