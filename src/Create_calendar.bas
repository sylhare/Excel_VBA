Sub create_Calendar()

'c and cTemp = column, r = row, m = month, j = day, nbJour = number of days in a month
A As interger, c As Integer, cTemp As Integer, r As Integer, m As Integer, dYear As Integer, j As Integer, nbJour As Integer
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
  
  
m = DayOffs(dYear, dayOff())

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

'----------------------------------------------------------------------------

Function DayOffs(dYear As Integer, dayOff() As Date)
'There's nearly one day off per month

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

End Function

'----------------------------------------------------------------------------

Public Function Paques(ByVal an As Integer) As Date
'Calcul de la date du dimanche de P㲵es ࡰartir de l'annꥠ325
'Performance par million d'appel :
'   - Entre 325 et 1582 et entre 1900 et 2099   => 1/4 de seconde
'   - Annꥠsup곩eure ࠱582 hors 1900 - 2099 => 1/2 de seconde
'Philben - v1.0 - Free to use
  Dim a As Integer, b As Integer, c As Integer, d As Integer, e As Integer, f As Integer
   If an < 10000 Then    'Limite sup곩eure des dates sous Access (31 dꤥmbre 9999)
     Select Case an
      Case 1900 To 2099    'Algorithme de Carter
        a = (204 - 11 * (an Mod 19)) Mod 30 + 22
         Paques = DateSerial(an, 3, a + 6 + (a > 49) - (an + an \ 4 + a + (a > 49)) Mod 7)
      Case Is > 1582    'Propos顥n 1876 dans la revue Nature (d곩v顤e l'algorithme de Delambre)
        a = an Mod 19: b = an \ 100: c = an Mod 100
         d = (19 * a + b - b \ 4 - (b - (b + 8) \ 25 + 1) \ 3 + 15) Mod 30
         e = (32 + 2 * (b Mod 4) + 2 * (c \ 4) - d - c Mod 4) Mod 7
         f = d + e - 7 * ((a + 11 * d + 22 * e) \ 451) + 114
         Paques = DateSerial(an, f \ 31, f Mod 31 + 1)
      Case Is > 324    'Algorithme de Oudin pour les dates juliennes < 1583 dꤲit par Claus Tondering
        a = (19 * (an Mod 19) + 15) Mod 30
         Paques = DateSerial(an, 3, 28 + a - (an + an \ 4 + a) Mod 7)
      End Select
   End If
End Function

