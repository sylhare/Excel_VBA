Sub create_Calendar()

'c and cTemp = column, r = row, m = month, j = day, nbJour = number of days in a month
Dim a As Integer, c As Integer, cTemp As Integer, r As Integer, m As Integer, dYear As Integer, j As Integer, nbJour As Integer
Dim dayOff(1 To 12) As Date, d As Date

'Initialisation
r = 2
Columns("B:B").EntireColumn.Hidden = True
Range("C3:NJ3").ColumnWidth = 1
Range("C3:NJ3").RowHeight = 52
Range("C1:NJ15").Clear
Range("C1:NJ15").UnMerge
Range("C1:NJ15").Interior.ColorIndex = none

With Cells(1, 1)
    .Value = "Année"
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
    .Font.Bold = True
End With

'Cells(2, 1).Value = 2015
dYear = Cells(2, 1).Value
  
  
Call DayOffs(dYear, dayOff())

'Calendar Creation and format
    For m = 1 To 12
             nbJour = Day((DateSerial(dYear, m + 1, 1) - 1))
             c = r + 1
             r = c - 1 + nbJour
             cTemp = c
             
             For j = 1 To nbJour
                    d = DateSerial(dYear, m, j)
                    
                    With Cells(3, cTemp)
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlBottom
                        .Orientation = 90
                        .ReadingOrder = xlContext
                        .Font.Size = 8
                        .Value = Format(d, "dd ddd")
                    End With
                    
                    If Weekday(d, vbMonday) > 5 Then
                    Cells(3, cTemp).Interior.ColorIndex = 16
                    End If
                                                              
                    If (m <> 2 And m <> 3 And m <> 8 And m <> 11 And d = dayOff(m)) Then
                    Cells(3, cTemp).Interior.ColorIndex = 40
                    End If
                                                      
                    cTemp = cTemp + 1
             Next j
             
             'Write & Format Month cells
             With Range(Cells(2, c), Cells(2, r))
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                .Merge
                .Value = Format(d, "mmmm")
                .Font.Bold = True
             End With
                    
    Next m
    
End Sub

Function DayOffs(dYear As Integer, dayOff() As Date)
'Initialisation Off Days
  Dim a As Integer
    
  '1st January
  dayOff(1) = DateSerial(dYear, 1, 1)
  '2
  dayOff(2) = DateSerial(1900, 1, 1)
  '3
  dayOff(3) = DateSerial(1900, 1, 1)
  'Easter Monday until 2099 '"\" means "/" with integer rest
  a = (204 - 11 * (dYear Mod 19)) Mod 30 + 22
  dayOff(4) = DateSerial(dYear, 3, a + 6 + (a > 49) - (dYear + dYear \ 4 + a + (a > 49)) Mod 7) + 1
  'Patriots National Day
  If (((DateSerial(dYear, 5, 7) - (DateSerial(dYear, 7, 1) - 2) Mod 7)) < DateSerial(dYear, 5, 4)) Then
    dayOff(5) = DateSerial(dYear, 5, 28) - (DateSerial(dYear, 5, 7) - 2) Mod 7
  Else
    dayOff(5) = DateSerial(dYear, 5, 21) - (DateSerial(dYear, 5, 7) - 2) Mod 7
  End If
  'Quebec Day
  dayOff(6) = DateSerial(dYear, 6, 24)
  'Canada Day
  If Weekday(DateSerial(dYear, 7, 1), vbMonday) = 7 Then
    dayOff(7) = DateSerial(dYear, 7, 2)
  Else
    dayOff(7) = DateSerial(dYear, 7, 1)
  End If
  '8
  dayOff(8) = DateSerial(1900, 1, 1)
  'Work day
  dayOff(9) = DateSerial(dYear, 9, 7) - (DateSerial(dYear, 9, 7) - 2) Mod 7
  'grace action
  dayOff(10) = DateSerial(dYear, 10, 14) - (DateSerial(dYear, 10, 7) - 2) Mod 7
  '11
  dayOff(11) = DateSerial(1900, 1, 1)
  'Christmas
  dayOff(12) = DateSerial(dYear, 12, 25)

End Function

