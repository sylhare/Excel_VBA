'Source: http://www.exceltrick.com/how_to/find-cells-containing-formulas-in-excel/

Sub FindFormulaCells()  
For Each cl In ActiveSheet.UsedRange  
If cl.HasFormula() = True Then  
cl.Interior.ColorIndex = 24  
End If  
Next cl  
End Sub 