Sub create_BellCalendar()

'c and cTemp = column, r = row, m = month, j = day, nbJour = number of days in a month
Dim a As Integer, c As Integer, cTemp As Integer, r As Integer, m As Integer, dYear As Integer, j As Integer, k As Integer, nbJour As Integer
Dim dayOff(1 To 12) As Date, d As Date, d1 As Date

'Entree de l'annee
If (Cells(1, 1).Value <> "") Then
    dYear = Cells(1, 1).Value
Else
    dYear = 2015
End If

'Initialisation - Nettoyage
r = 2
cTemp = 3
Columns("B:B").EntireColumn.Hidden = True
Range("C3:NJ3").ColumnWidth = 5
Range("C3:NJ3").RowHeight = 10
Cells.ClearContents
Cells.Borders.LineStyle = xlNone
Range("C1:NJ15").UnMerge
Range("C1:NJ15").Interior.ColorIndex = none

k = DayOffs(dYear, dayOff())
Cells(1, 1).Value = dYear

'Calendar Creation and format
    For m = 1 To 12
             nbJour = Day((DateSerial(dYear, m + 1, 1) - 1))
             j = 1
             c = cTemp
             
             While j <= nbJour 'And (month(d + 4) > month(d))
                    d = DateSerial(dYear, m, j)
                        
                            'vbMonday -> lundi premier jour de la semaine
                            If Weekday(d, vbMonday) = 1 Then
                            
                            'For k = cTemp To cTemp + 4
                                                                                                      
                            'Next k
                            
                            With Range(Cells(3, cTemp), Cells(3, cTemp + 4))
                                .HorizontalAlignment = xlCenter
                                .VerticalAlignment = xlCenter
                                .ReadingOrder = xlContext
                                .Merge
                                .Font.Size = 7
                                .NumberFormat = "@"
                                .Value = Format(d, "dd") + "-" + Format(d + 4, "dd")
                                .Font.ThemeColor = xlThemeColorDark1
                                .Font.Name = "Arial"
                                .Interior.Color = 16711680
                                End With
                            With Range(Cells(3, cTemp), Cells(100, cTemp + 4))
                                .Borders(xlEdgeLeft).Weight = xlMedium
                                .Borders(xlEdgeRight).Weight = xlMedium
                            End With
                            Range(Cells(3, cTemp), Cells(3, cTemp + 4)).ColumnWidth = 0.42
                            
                            cTemp = cTemp + 5
                            End If

                        j = j + 1
                        
              Wend
             
             r = cTemp - 1
             
             'Write & Format Month cells
             With Range(Cells(2, c), Cells(2, r))
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                .Merge
                .Value = Format(d, "Mmm. yyyy")
                .RowHeight = 22.5
                .Interior.Color = 16711680
                .Font.ThemeColor = xlThemeColorDark1
                .Font.Bold = True
                .Font.Name = "Arial"
                .Borders.Weight = xlMedium
             End With
                    
    Next m
    
End Sub



Function DayOffs(dYear As Integer, dayOff() As Date) 'As Integer

'Initialisation Off Days
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
  
  'La fonction renvoi 1
  DayOffs = 1

End Function

Function firstJanMon(dYear As Integer, d As Date)

'8 parceque premi鳥 r걩tition des jour. Date serial renvoi le "num곯 de s곩e du jour" (Dimanche = 1, lundi = 2, ࡳamedi = 7)
'Pour Obtenir le premier lundi, il faut obtenir le num곯 de s곩e du lundi : 2 pour la premi鳥 semaine. 8-6=2 donc 6 janvier.
d = Format(DateSerial(dYear, 1, 8) - Weekday(DateSerial(dYear, 1, 6)), "dddd dd/mm/yy")
'On prend le num곯 de s곩e du 7 janvier moins ce num곯 de s곩e - 2 (Pour le lundi) modulo 7 pour obtenir le premier de la semaine
d = DateSerial(dYear, 1, 7) - (DateSerial(dYear, 1, 7) - 2) Mod 7

End Function

Function Calendar()

'Second version do while

 For m = 1 To 12
             nbJour = Day((DateSerial(dYear, m + 1, 1) - 1))
             j = 1
             c = cTemp
             
             Do While j <= nbJour
                    d = DateSerial(dYear, m, j)
                        
                        'If (i + 5 < nbJour) Then
                         '   i = i + 4
                        'End If
                        If Weekday(d, vbMonday) > 5 Then
                        
                        j = j + 1
                        
                        Else
                            Cells(4, cTemp).ColumnWidth = 0.5 '0.42
                        
                            'vbMonday -> lundi premier jour de la semaine
                            If Weekday(d, vbMonday) = 1 Then
                            
                            'For k = cTemp To cTemp + 4
                                                                                                      
                            'Next k
                            
                            With Range(Cells(3, cTemp), Cells(3, cTemp + 4))
                                .HorizontalAlignment = xlCenter
                                .VerticalAlignment = xlBottom
                                .ReadingOrder = xlContext
                                .Merge
                                .Font.Size = 8
                                .NumberFormat = "@"
                                .Value = Format(d, "dd") + " - " + Format(d + 4, "dd")
                            End With
                            
                            If (month(d + 4) > month(d)) Then
                            Exit Do
                            End If
                            
                            End If
                            cTemp = cTemp + 1
                        End If
                        j = j + 1
                        
              Loop
             
             r = cTemp - 1
             
             'Write & Format Month cells
             With Range(Cells(2, c), Cells(2, r))
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                .Merge
                .Value = Format(d, "Mmm. yyyy")
                .RowHeight = 22.5
                .Font.Bold = True
             End With
                    
    Next m
    
    'first version while

    For m = 1 To 12
             nbJour = Day((DateSerial(dYear, m + 1, 1) - 1))
             j = 1
             c = cTemp
             
             While j <= nbJour
                    d = DateSerial(dYear, m, j)
                        
                        'If (i + 5 < nbJour) Then
                         '   i = i + 4
                        'End If
                        If Weekday(d, vbMonday) > 5 Then
                        
                        j = j + 1
                        
                        Else
                            Cells(4, cTemp).ColumnWidth = 0.42
                        
                            If Weekday(d, vbMonday) = 5 Then
                            
                            'For k = cTemp To cTemp + 4
                                                                                                      
                            'Next k
                            
                            With Range(Cells(3, cTemp - 4), Cells(3, cTemp))
                                .HorizontalAlignment = xlCenter
                                .VerticalAlignment = xlBottom
                                .ReadingOrder = xlContext
                                .Merge
                                .Font.Size = 8
                                .NumberFormat = "@"
                                .Value = Format(d - 4, "dd") + " - " + Format(d, "dd")
                            End With
                            
                            End If
                            cTemp = cTemp + 1
                        End If
                        j = j + 1
                        
              Wend
             
             r = cTemp - 1
             
             'Write & Format Month cells
             With Range(Cells(2, c), Cells(2, r))
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                .Merge
                .Value = Format(d, "Mmm. yyyy")
                .RowHeight = 22.5
                .Font.Bold = True
             End With
                    
    Next m
End Function

----------------

Sub create_BellCalendar()

'c and cTemp = column, r = row, m = month, j = day, nbJour = number of days in a month
Dim a As Integer, c As Integer, cTemp As Integer, r As Integer, m As Integer, dYear As Integer, j As Integer, k As Integer, nbJour As Integer
Dim dayOff(1 To 12) As Date, d As Date, d1 As Date

'Entrꥠde l'annꥍ
If (Cells(1, 1).Value <> "") Then
    dYear = Cells(1, 1).Value
Else
    dYear = 2015
End If

'Initialisation - Nettoyage
r = 2
cTemp = 3
Columns("B:B").EntireColumn.Hidden = True
Range("C3:NJ3").ColumnWidth = 5
Range("C3:NJ3").RowHeight = 10
Cells.ClearContents
Cells.Borders.LineStyle = xlNone
Range("C1:NJ15").UnMerge
Range("C1:NJ15").Interior.ColorIndex = none

k = DayOffs(dYear, dayOff())
Cells(5, 1).Value = k
Cells(1, 1).Value = dYear

'Calendar Creation and format
    For m = 1 To 12
             nbJour = Day((DateSerial(dYear, m + 1, 1) - 1))
             j = 1
             c = cTemp
             
             While j <= nbJour 'And (month(d + 4) > month(d))
                    d = DateSerial(dYear, m, j)
                        
                            'vbMonday -> lundi premier jour de la semaine
                            If Weekday(d, vbMonday) = 1 Then
                            
                            'For k = cTemp To cTemp + 4
                                                                                                      
                            'Next k
                            
                            With Range(Cells(3, cTemp), Cells(3, cTemp + 4))
                                .HorizontalAlignment = xlCenter
                                .VerticalAlignment = xlCenter
                                .ReadingOrder = xlContext
                                .Merge
                                .Font.Size = 7
                                .NumberFormat = "@"
                                .Value = Format(d, "dd") + "-" + Format(d + 4, "dd")
                                .Font.ThemeColor = xlThemeColorDark1
                                .Font.Name = "Arial"
                                .Interior.Color = 16711680
                                End With
                            With Range(Cells(3, cTemp), Cells(100, cTemp + 4))
                                .Borders(xlEdgeLeft).Weight = xlMedium
                                .Borders(xlEdgeRight).Weight = xlMedium
                            End With
                            Range(Cells(3, cTemp), Cells(3, cTemp + 4)).ColumnWidth = 0.42
                            
                            cTemp = cTemp + 5
                            End If

                        j = j + 1
                        
              Wend
             
             r = cTemp - 1
             
             'Write & Format Month cells
             With Range(Cells(2, c), Cells(2, r))
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                .Merge
                .Value = Format(d, "Mmm. yyyy")
                .RowHeight = 22.5
                .Interior.Color = 16711680
                .Font.ThemeColor = xlThemeColorDark1
                .Font.Bold = True
                .Font.Name = "Arial"
                .Borders.Weight = xlMedium
             End With
                    
    Next m
    
End Sub



Function DayOffs(dYear As Integer, dayOff() As Date) 'As Integer

'Initialisation Off Days
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
  
  'La fonction renvoi 1
  DayOffs = 1

End Function

Function firstJanMon(dYear As Integer, d As Date)

'8 parceque premiere repetition des jour. Date serial renvoi le "numero de serie du jour" (Dimanche = 1, lundi = 2, samedi = 7)
'Pour Obtenir le premier lundi, il faut obtenir le numero de serie du lundi : 2 pour la premiere semaine. 8-6=2 donc 6 janvier.
d = Format(DateSerial(dYear, 1, 8) - Weekday(DateSerial(dYear, 1, 6)), "dddd dd/mm/yy")
'On prend le numero de serie du 7 janvier moins ce numero de serie - 2 (Pour le lundi) modulo 7 pour obtenir le premier de la semaine
d = DateSerial(dYear, 1, 7) - (DateSerial(dYear, 1, 7) - 2) Mod 7

End Function
