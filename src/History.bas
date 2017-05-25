Sub History()

Dim i As Integer

i = 3

If MsgBox("Irreversible Operation:" & vbCrLf & "Are you sure ?", vbQuestion + vbYesNo, "Saving ...") = vbYes Then
While (Cells(i, 3).Value <> "")

If (Cells(i, 5) <> "" And (Cells(i, 7).Value <> "Closed" Or Cells(i, 7).Value <> "TBD")) Then
    Cells(i, 5 + 22).Value = Date & ": " & Cells(i, 5).Value & vbCrLf & Cells(i, 5 + 22).Value
    Cells(i, 5).RowHeight = 13
    Cells(i, 5).Value = ""
End If

i = i + 1
Wend

End If

End Sub
    
-------------------
    
Sub History()

Dim i As Integer, notes As Integer

notes = 9
i = 4

While (Cells(i, 8).Value <> "")
    If (Cells(i, notes) <> "") Then
        Cells(i, notes + 1).Value = Date & ": " & Cells(i, notes).Value & vbCrLf & Cells(i, notes + 1).Value
        Cells(i, notes).RowHeight = 13
        Cells(i, notes).Value = ""
    End If
    
    i = i + 1
Wend

ActiveWorkbook.Save

End Sub
        

---------------
        

Sub History()

    Dim i As Integer

    i = 2

    While (Cells(i, 3).Value <> "")

    If (Cells(i, 5) <> "") Then
        Cells(i, 6).Value = Date & ": " & Cells(i, 5).Value & vbCrLf & Cells(i, 6).Value
        'Cells(i, 5).RowHeight = 13
        Cells(i, 5).Value = ""
    End If

    i = i + 1
    Wend

    'ActiveWorkbook.Save

End Sub

        
        
------------
        
Sub History()

Dim i As Integer

i = 3

If MsgBox("Irreversible Operation:" & vbCrLf & "Are you sure ?", vbQuestion + vbYesNo, "Saving ...") = vbYes Then
    While (Cells(i, 3).Value <> "")

        If (Cells(i, 5) <> "" And (Cells(i, 7).Value <> "Closed" Or Cells(i, 7).Value <> "TBD")) Then
            Cells(i, 5 + 21).Value = Date & ": " & Cells(i, 5).Value & vbCrLf & Cells(i, 5 + 21).Value
            Cells(i, 5).RowHeight = 13
            Cells(i, 5).Value = ""
        End If

        i = i + 1
    Wend

    'Si oui sauvegarde le fichier
    ActiveWorkbook.Save
End If

End Sub
                
---------------

Sub History()

Dim i As Integer

i = 3

While (Cells(i, 3).Value <> "")

    If (Cells(i, 5) <> "" And (Cells(i, 7).Value <> "Closed" Or Cells(i, 7).Value <> "TBD")) Then
        Cells(i, 5 + 21).Value = Date & ": " & Cells(i, 5).Value & vbCrLf & Cells(i, 5 + 21).Value
        Cells(i, 5).RowHeight = 13
        Cells(i, 5).Value = ""
    End If

    i = i + 1
Wend

ActiveWorkbook.Save

End Sub
                    
                    
                    
                    
--------------------------------

Sub Save()
 
Dim path As String
 
    path = "C:\Users\username\Documents\" & Format(Date, "yyyy") & " - " & Format(Date, "mm")
    file = "title " & Format(Date, "mm") & " - " & Format(Date, "dd") & ".xlsx"
 
    Application.DisplayAlerts = False
    ThisWorkbook.SaveAs Filename:=file, FileFormat:=51 'Save
    Application.DisplayAlerts = True
    MsgBox ("Your file has been saved") 'Info window
 
    Exit Sub
 
End Sub


Sub Clean()

    Dim i As Integer

    i = 1

    While (i < 201)
        i = i + 1
        If Cells(i, 1).Value = "" Then
            Rows(i).EntireRow.Delete
        End If

    Wend


End Sub   
                                        