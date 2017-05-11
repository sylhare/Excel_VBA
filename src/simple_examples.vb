Sub EasyFct()

'MsgBox "Hello VBA World!"
'MsgBox "Entered value is " & Range("A1").Value & vbNewLine & "Test is in A1"

'Range, permet de sélectionner la/les case(s), on peut assigner 2 à la case (range(case).value):
Range("B1").Value = "Range"
Range("D1:D2,C4:D5").Value = 10

'Pour nommer une selection de case et leur assigner une valeur
Dim example As Range
Set example = Range("A3:B7")
example.Value = 5

'Compter le nombre de case
MsgBox "Total : " & example.Count & vbNewLine & "Lignes : " & example.Rows.Count & vbNewLine & "Colonnes : " & example.Columns.Count


'Cells permet de selectionner une case avec ses coordonnées
Cells(1, 3).Value = "Cells"
Range(Cells(7, 5), Cells(10, 5)).Value = "RangeCells"

'Select permet de selectionner automatiquement des cellules
Dim test As Range
Set test = Range("A1:F10")
test.Select
'test.Rows(3).Select 'Pour selectionner une ligne
'test.Columns(2).Select 'Pour selectionner une colonne

'Pour faire un copier coller
'Range("A14:A15").Select
'selection.Copy
'Range("E1").Select
'ActiveSheet.Paste
'ou : les valeurs en A14:A15 assignées en E1:E2
Range("E1:E2").Value = Range("A14:A15").Value

'Tout supprimer
'Range("A1:F10").ClearContents
'ou
'Range("A1:F10").Value = ""

'Faire des boucles
Dim c As Integer, i As Integer, j As Integer

For c = 1 To 3
    For i = 10 To 13
        For j = 7 To 8
            Worksheets(c).Cells(i, j).Value = "loop"
        Next j
    Next i
Next c

'Faire un do while
Do While i < 6
    Cells(i, 11).Value = "dow"
    i = i + 1
Loop
End Sub

Sub formatSheet()

Worksheets(2).Range("C3:MV3").ColumnWidth = 3
Worksheets(2).Range("C3:MV3").RowHeight = 52

Dim c As Integer, r As Integer, i As Integer

Dim month As Range

    For i = 1 To 12
             r = 2 + i * 31
             c = 2 + r - 31
             Worksheets(2).Range(Cells(2, c), Cells(2, r)).Merge
             
    Next i
    


End Sub

'Aller dans Mode de création puis propriété pour modifier le nom du bouton et l'associer à la fonction
'Ici au lieu de commandbutton1 j'ai "plus10"
Private Sub plus10_Click()

Dim i As Integer
i = 1

'<> means "not equal to"
'This do while stop when the cells (i,8(=G)) is empty
Do While Cells(i, 7).Value <> ""
    Cells(i, 8).Value = Cells(i, 7).Value + 10
    i = i + 1
Loop

End Sub

