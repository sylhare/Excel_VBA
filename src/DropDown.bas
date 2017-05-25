'===============================
'Source
'http://blog.contextures.com/archives/2010/09/13/data-validation-combo-box-in-excel-table/
'
'===============================

Option Explicit

Private Sub TempCombo_KeyDown(ByVal _
        KeyCode As MSForms.ReturnInteger, _
        ByVal Shift As Integer)
'Define the behaviour of the dropdown list of the TempCombo activeX object
'The TempCombo is the name was given, it's referred in the Worksheet_SelectionChange as well
        
Dim tb As ListObject
Dim lCols As Long
Dim lCol As Long
Dim lRows As Long
Dim lRow As Long
Dim lColStart As Long
Dim lRowStart As Long

On Error Resume Next
Set tb = ActiveCell.ListObject
lCols = tb.ListColumns.Count
lCol = tb.ListColumns(lCols).Range.Column
lRows = tb.ListRows.Count
lRow = tb.ListRows(lRows).Range.Row
lColStart = tb.ListColumns(1).Range.Column
lRowStart = tb.ListRows(1).Range.Row - 1
        
    'Hide combo box and move to next cell on Enter and Tab
    Select Case KeyCode
        Case 9 'tab
            If ActiveCell.Column = lCol Then
                If ActiveCell.Row = lRow Then
                    tb.Resize Range(Cells(lRowStart, lColStart), Cells(lRows + 2, 3))
                End If
                ActiveCell.Offset(1, -(lCol - lCols)).Activate
            Else
                ActiveCell.Offset(0, 1).Activate
            End If
        Case 13 'enter
            ActiveCell.Offset(1, 0).Activate
        Case Else
            'do nothing
    End Select

End Sub
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
'When the selection change in the Worksheet (detect you click on something else)

Dim str As String
Dim cboTemp As OLEObject
Dim ws As Worksheet
Dim wsList As Worksheet

Set ws = ActiveSheet
Set wsList = Sheets("ValidationLists")
Application.EnableEvents = False
Application.ScreenUpdating = False

'Name of the activeX combox dropdown list from the developer tab
'The name can be changed in design mode, rightclick > properties
Set cboTemp = ws.OLEObjects("TempCombo")
  On Error Resume Next
  With cboTemp
    .Top = 10
    .Left = 10
    .Width = 0
    .ListFillRange = ""
    .LinkedCell = ""
    .Visible = False
    .Value = ""
  End With
  
  
On Error GoTo errHandler
  If Target.Validation.Type = 3 Then
    Application.EnableEvents = False
    str = Target.Validation.Formula1
    str = Right(str, Len(str) - 1)
    With cboTemp
      .Visible = True
      .Left = Target.Left
      .Top = Target.Top
      .Width = Target.Width + 15
      .Height = Target.Height + 5
      .ListFillRange = str
      .LinkedCell = Target.Address
    End With
    cboTemp.Activate
  End If


errHandler:
  Application.ScreenUpdating = True
  Application.EnableEvents = True
  Exit Sub

End Sub



