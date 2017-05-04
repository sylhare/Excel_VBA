Option Explicit
 
'Je te conseille de mettre la cible comme une propriete UserForm
' et de la renseigner juste avant de faire TonUserForm.Show sur le clique de la cellule
Public cible As String
Public nomFichier As String

'Pour ouvrir le document (par son lien hypertexte)
Private Sub CommandButton1_Click()
    If Not FichierExiste(cible) Then
        MsgBox "Fichier introuvable" & vbCrLf & cible
    Else
        ThisWorkbook.FollowHyperlink cible
    End If
End Sub

Private Sub Image_Click()

End Sub

' pour reprend le nom du document et afficher la photo correspondante
Private Sub UserForm_Layout()
    'Affichage du nom
    Label1.Caption = nomFichier
    
    'Affichage de la photo
    Dim cheminImage As String
    cheminImage = ThisWorkbook.Path & "\" & nomFichier & ".jpg"
    
    If FichierExiste(cheminImage) Then
        Image.Picture = LoadPicture(cheminImage)
    End If
End Sub

'Teste l'existence d'un fichier
Private Function FichierExiste(filePath As String) As Boolean
    'Dim fso As New FileSystemObject
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    FichierExiste = fso.FileExists(filePath)
End Function


------------------------------------------------------------------

Option Explicit
 
'Sur le changement de selection
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
        If Not Intersect(Target, Me.Range("B2:B20")) Is Nothing Then 'Sur la premier colonne
        If Target.Value <> "" Then 'Si on a un fichier
            Dim cible As String
            
            'Si on n'a pas d'hyperlien, on ne charge pas le userForm
            If Target.Offset(0, 1).Hyperlinks.Count = 0 Then
                MsgBox "Le document sur lequel vous avez cliqu顮'a pas de lien correspondant"
                Exit Sub
            End If
            
            'On attribut la cible du lien au userForm et on l'affiche
            forming.nomFichier = Target.Value
            forming.cible = Target.Offset(0, 1).Hyperlinks(1).Address
            forming.Show
        End If
    End If
End Sub

-------------------------------------------------------------------

Sub lancer_lien_hptxt()
' lancer_lien_hptxt Macro

    Range("C3").Select
    Selection.Hyperlinks(1).Follow NewWindow:=False, AddHistory:=True
End Sub