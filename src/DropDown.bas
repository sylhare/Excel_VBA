Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    'Hide the combobox when not called and place it top of the sheet
    Dim combo As OLEObject
    Dim wks As Worksheet
    Set wks = Application.ActiveSheet
    On Error Resume Next
    Application.EnableEvents = False
    Application.ScreenUpdating = True
    
    'Name of the dropdown (ComboBox) list from the developer tab, it's the default name
    Set combo = wks.OLEObjects("ComboBox1")
    
    With combo
        .Top = 10
        .Left = 10
        .Visible = False
        .Value = ""
    End With
    Application.EnableEvents = True
    
End Sub

Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    'Activate the ComboBox when doubleClicking on a cell that has a data validator (Data >Data validation)
    'It has been customised with J2 the linked cell to make custom search with formulas
    'Otherwise .LinkedCell = Target.Address can be use to modify the clicked cell 
    Dim dropRange As String
    Dim combo As OLEObject
    Dim wks As Worksheet
    Set wks = Application.ActiveSheet
    On Error Resume Next
    Application.EnableEvents = False
    
    Set combo = wks.OLEObjects("ComboBox1")
    
    'Define which search page to look for when using the drop for two different columns in the same page
    'Set the searching cell when doube clicking on column A
    If Not Intersect(Target, Range("A:A")) Is Nothing Then
        With combo
            .LinkedCell = "search!$J$2"
            .Visible = False
        End With
    End If

'Set the searching cell when doube clicking on column H
    If Not Intersect(Target, Range("H:H")) Is Nothing Then
        With combo
            .LinkedCell = "search!$J$5"
            .Visible = False
        End With
    End If

    If Target.Validation.Type = 3 Then
        Cancel = True

        'Define the value of the range, based on data validation of the target cell
        dropRange = Target.Validation.Formula1
        dropRange = Right(dropRange, Len(dropRange) - 1)
        
        'The ComboBox appear when there is a data validation on the cell
        If dropRange <> "" Then
            With combo
                .Visible = True
                .Left = Target.Left - 1
                .Top = Target.Top - 1
                .Width = Target.Width + 15
                .Height = Target.Height + 1
                .ListFillRange = dropRange
            End With
        
            combo.Activate
            Me.ComboBox1.DropDown
        End If
    End If
    
    Application.EnableEvents = True
    
End Sub

Private Sub ComboBox1_KeyDown(ByVal _
        KeyCode As MSForms.ReturnInteger, _
        ByVal Shift As Integer)
    'Define the behaviour of the comboBox named "ComboBox1" when key is touched
    'Modified to work with one "searching" case linked to the ComboBox which value will be copied to the activeCell
    Dim value As String
    
    If Worksheets("search").Range("$J$2").value <> "" Then
        value = Worksheets("search").Range("$J$2").value
    Else
        value = Worksheets("search").Range("$J$5").value
    End If
    
    Select Case KeyCode
        Case 9 'Tab key
            ActiveCell.Value = Worksheets("search").Range("$J$2").Value
            ActiveCell.Offset(0, 1).Activate
        Case 13 'Enter key
            ActiveCell.Value = Worksheets("search").Range("$J$2").Value
            ActiveCell.Offset(1, 0).Activate
        Case 37 'Left Arrow key
            ActiveCell.Offset(0, -1).Activate
        Case 39 'Right arrow key
            ActiveCell.Offset(0, 1).Activate
        Case Else
            'do nothing
    End Select

End Sub
