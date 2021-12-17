Sub create_BlueCalendar()

'For time efficiency (Hide changes on screen)
Application.Calculation = xlCalculationManual
Application.ScreenUpdating = False

'c and cTemp = column, r = row, m = month, j = day, nbJour = number of days in a month
Dim c As Integer, cTemp As Integer, r As Integer, k As Integer
Dim dYear As Integer, m As Integer, j As Integer, nbJour As Integer, wDays As Integer
Dim dayOff(1 To 12) As Date, d As Date, d1 As Date, d2 As Date

'Entrée de l'année
If (Cells(1, 1).Value <> "") Then
    dYear = Cells(1, 1).Value
Else
    dYear = 2015
End If

'Initialisation - Cleaning
r = 2
cTemp = 3
d1 = DateSerial(dYear, 1, 8) - Weekday(DateSerial(dYear, 1, 8) - 2)
d2 = DateSerial(dYear + 1, 1, 8) - Weekday(DateSerial(dYear + 1, 1, 8) - 2)
wDays = DateDiff("w", d1, d2, vbMonday, vbFirstFullWeek) * 5

initClear
Call formatWeeks(wDays + r)
Call formatMonths(wDays + r)
Call DayOffs(dYear, dayOff())
Cells(1, 1).Value = dYear

'Calendar Creation and format
    For m = 1 To 12
             nbJour = Day((DateSerial(dYear, m + 1, 1) - 1))
             j = 1
             c = cTemp
             
                        While j <= nbJour
                            d = DateSerial(dYear, m, j)
                        
                            'vbMonday -> lundi premier jour de la semaine
                            If Weekday(d, vbMonday) = 1 Then
                            
                                'Write the week days
                                With Range(Cells(3, cTemp), Cells(3, cTemp + 4))
                                    .Merge
                                    .Value = Format(d, "dd") + "-" + Format(d + 4, "dd")
                                    End With
                                    
                                'Black borders of the weeks
                                With Range(Cells(3, cTemp), Cells(100, cTemp + 4))
                                    .Borders(xlEdgeLeft).Weight = xlMedium
                                    .Borders(xlEdgeRight).Weight = xlMedium
                                End With
                                
                                cTemp = cTemp + 5
                                                                                           
                             End If

                            j = j + 1
                               
                         Wend
             
             r = cTemp - 1
             
             'Write Months
             With Range(Cells(2, c), Cells(2, r))
                .Merge
                .Value = Format(d, "mmm. yyyy")
             End With
    Next m
    
    'To set it back to normal
Application.Calculation = xlCalculationAutomatic
Application.ScreenUpdating = True
    
End Sub

Function initClear()
'Clear everything before doing the calendar

Columns("B:B").EntireColumn.Hidden = True
Columns("C:NJ").ColumnWidth = 0.42
Range("C3:NJ150").RowHeight = 10
Cells.ClearContents
Cells.Borders.LineStyle = xlNone
Range("C1:NJ150").UnMerge
Range("C1:NJ150").Interior.ColorIndex = none

End Function

Function formatWeeks(r As Integer)

    'Format Weeks
    With Range(Cells(3, 3), Cells(3, r))
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                .NumberFormat = "@"
                .Font.Size = 7
                .ReadingOrder = xlContext
                .Font.ThemeColor = xlThemeColorDark1
                .Font.Name = "Arial"
                .Interior.Color = 16711680
    End With

End Function

Function formatMonths(r As Integer)

    'Format Months
    With Range(Cells(2, 3), Cells(2, r))
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                .RowHeight = 22.5
                .Interior.Color = 16711680
                .Font.ThemeColor = xlThemeColorDark1
                .Font.Bold = True
                .Font.Name = "Arial"
                .Borders.Weight = xlMedium
    End With

End Function

Sub Browsing()

Dim dYear As Integer, i As Integer, d1 As Date, d2 As Date, m As Integer
Dim nbJour As Integer, k As Integer
Dim workDays(1 To 300) As Date
Dim dayOff(1 To 12) As Date

Rows("4:4").ClearContents

dYear = 2015
k = DayOffs(dYear, dayOff())

'nbJour = Day((DateSerial(dYear, m + 1, 1) - 1))
'd1 commence au premier lundi de l'année
d1 = DateSerial(dYear, 1, 8) - Weekday(DateSerial(dYear, 1, 8) - 2)
d2 = DateSerial(dYear + 1, 1, 8) - Weekday(DateSerial(dYear + 1, 1, 8) - 2)

k = 0
i = 1

'Cells(4, 1) = DateDiff("w", d1, d2, vbMonday, vbFirstFullWeek) * 5

While (d1 + k) < d2

    'Pas un weekend
    If Weekday(d1 + k, vbMonday) <= 5 Then
        workDays(i) = d1 + k
        Cells(4, i + 2).Value = d1 + k
        m = Month(d1 + k)
        If ((d1 + k) = DayOffs(m)) Then
            Cells(4, i + 2).Interior.ColorIndex = 40
        End If
        
        i = i + 1
    End If
    
    k = k + 1
Wend

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

Function firstJanMon(dYear As Integer, d As Date)

'8 -> première répétition. Weekday renvoi le "numéro de série du jour" (Dimanche = 1, lundi = 2, à samedi = 7)
'Pour Obtenir le premier lundi, il faut obtenir le numéro de série du lundi : 2 pour la première semaine. 8-6=2 donc 6 janvier. Pour mardi : 3 -> 8-5=3 donc 5 janvier. Pour Samedi : 7 -> 8-7=1 d'où le 8 janvier.
'weekdays fait le modulo et renvoie entre 1(Dimanche) et 7(Samedi). Le lundi (2) Première apparition du lundi dans les 8 jours
Cells(17, 1).Value = Format(DateSerial(dYear, 1, 8) - Weekday(DateSerial(dYear, 1, 8) - 2), "dd/mm/yy")

'On prend le numéro de série du 7 janvier moins ce numéro de série - 2 (Pour le lundi) modulo 7 pour obtenir le premier de la semaine
'On prend pour point de référence le 7 janvier de la même année – obtenu grâce à la formule date(annee(A2);1; 7) – auquel on ôte 2 (le code du lundi). On prend le modulo par 7 de cette valeur, ce qui représente le nombre de jours qu’il faut ôter au 7 janvier pour tomber sur le premier lundi.
d = DateSerial(dYear, 1, 7) - (DateSerial(dYear, 1, 7) - 2) Mod 7

End Function

