# Excel VBA

In excel, you have to enable first the macro (and select the developper option). 
Then you can press <kbd>ALT</kbd> + <kbd>F11</kbd> to go into edit macro mode. 
To learn you can start with "recording macro" to see what excel is recording, but it's not the most efficient way.

You can copy / paste the files `.bas` in the `src` folder inside the macro editor (VBS) to execute them.

#### Sources

- Microsoft Excel 2010: [getting started](https://msdn.microsoft.com/fr-fr/library/office/ee814737(v=office.14).aspx)


##### French Tips

- [Excel plus: VBA tips](http://www.excel-plus.fr/category/vba/) 
- [Excel pratique VBA tips](https://www.excel-pratique.com/)
- [Developpez.com outlook vba](http://dolphy35.developpez.com/article/outlook/vba/#LV-A)


##### English Tips

- [Ozgrid excel vb tutorial](http://www.ozgrid.com/Excel/free-training/ExcelVBA1/excel-vba1-index.htm)
- [Excel easy: vba macro](http://www.excel-easy.com/vba.html)
- [Tutorialspoint VBA](https://www.tutorialspoint.com/vba/index.htm)
- [Rondebruin mail with outlook](http://www.rondebruin.nl/win/s1/outlook/mail.htm)
- [Extend Office: AutoComplete](https://www.extendoffice.com/documents/excel/2401-excel-drop-down-list-autocomplete.html)
- [Convert data type vb](http://www.convertdatatypes.com/Language-VB6-VBA.html)


## Excel autofind drop down menu

Find out the detail tutorial on how to do it here:

 - [How to create a dropdown search menu from an excel spreadsheet](https://sylhare.github.io/2015/02/15/Excel-autofind-dropdown-menu.html)


## Other tips

You can find everything on my blog at:

 - [Excel Macro tips](https://sylhare.github.io/2015/04/17/Excel-macro-tips.html)

#### Comment / Uncomment bloc of code

There's a Comment / Uncomment button that can be toggled. For that **right click** on the **menu bar** then click on **edit**, the edit tool bar will appear (you can place it in your quick access bar). There should be a **comment** and **Uncomment** **icon**. This commands will basically add or remove `'` at the beginning of every selected ligns. 

#### Calling a Sub

Here are an example on how to call a subroutine: [here](https://msdn.microsoft.com/en-us/library/office/gg251432.aspx)
It can be tricky.
```vb
Test "N23:Q23", 1
Call Test("N23:Q23", 1)


Sub Test(xRange As Range, val As Integer)
	'some coding
End Sub
```


#### Accelerate Macro

Here are a couple of lines that can greatly improve the efficiency of your VBA macro.

```vb
Sub example()
	'Stop automatic calculation of excel cells
	Application.Calculation = xlCalculationManual
	'Stop screen updating
	Application.ScreenUpdating = False

	'Some code

	'Put it back to "normal"
	Application.Calculation = xlCalculationAutomatic
	Application.ScreenUpdating = True
End Sub
```

#### Hide "0" value of empty cells

Sometime there are some 0 that pops up with the below formulas, so here is a trick to hide them through formating.
Available [here](https://support.office.com/en-us/article/Display-or-hide-zero-values-3ec7a433-46b8-4516-8085-a00e9e476b03):

- Home > Format > Format Cells
- Number > Custom
- type : `0;;;@`

#### Userform
Some example for the Userform

```vb
Userform
    Textbox 
        Multiline : True
        EnterKeyBehavior = True (sinon ctrl + Enter)
```


#### Closing procedure
Procedure to close a file

```vb
Sub arret()
	'stop the current sub
    ActiveWorkbook.Save
    ActiveWorkbook.Close True
End Sub
```

Close the file after 10 seconds

```vb
Private Sub Workbook_Open()
     temp = Now + TimeValue(« 00:00:10 »)
     Application.OnTime temp, « arret »
End Sub
```
