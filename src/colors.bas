'Show vb colors in Excel, with a simple loop going through the 56 ColorIndex

	Sub coulors() 
    		Range(« A1 »).Select
    			For i = 1 To 56
        			ActiveCell.Offset(0, 1).Interior.ColorIndex = ActiveCell.Value
        			ActiveCell.Offset(1, 0).Select
    			Next i
    		Range(« A1 »).Select
	End Sub