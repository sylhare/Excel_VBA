'Source: http://www.exceltrick.com/how_to/find-cells-containing-formulas-in-excel/

Sub FindFormulaCells()  
    For Each cell In ActiveSheet.UsedRange  
        If cell.HasFormula() = True Then  
            cell.Interior.ColorIndex = 24  
        End If  
    Next cell  
End Sub 