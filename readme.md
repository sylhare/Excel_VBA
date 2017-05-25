# Excel VBA

In excel, you have to enable first the macro (and select the developper option). 
Then you can press `ALT + F11` to go into edit macro mode. 
To learn you can start with "recording macro" to see what excel is recording, but it's not the most efficient way.

#### Sources

- [Microsoft getting started](https://msdn.microsoft.com/fr-fr/library/office/ee814737(v=office.14).aspx)


##### French Tips

- [Excel plus: VBA tips](http://www.excel-plus.fr/category/vba/) 
- [Excel pratique VBA tips](https://www.excel-pratique.com/)
- [Developpez.com outlook vba](http://dolphy35.developpez.com/article/outlook/vba/#LV-A)
- [Excel exercices and solutions](http://users.skynet.be/micdub/vba12.htm)
- [Excel dates](http://boisgontierjacques.free.fr/pages_site/dates.htm#saisieDate)


##### English Tips

- [Ozgrid excel vb tutorial](http://www.ozgrid.com/Excel/free-training/ExcelVBA1/excel-vba1-index.htm)
- [Excel easy: vba macro](http://www.excel-easy.com/vba.html)
- [Tutorialspoint VBA](https://www.tutorialspoint.com/vba/index.htm)
- [Techonthenet vba functions](https://www.techonthenet.com/excel/formulas/index_vba.php)
- [Rondebruin mail with outlook](http://www.rondebruin.nl/win/s1/outlook/mail.htm)
- [Extend Office: AutoComplete](https://www.extendoffice.com/documents/excel/2401-excel-drop-down-list-autocomplete.html)
- [Convert data type vb](http://www.convertdatatypes.com/Language-VB6-VBA.html)

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
'Sub arret() stop the current sub
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
