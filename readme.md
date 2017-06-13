# Excel VBA

In excel, you have to enable first the macro (and select the developper option). 
Then you can press `[ALT] + [F11]` to go into edit macro mode. 
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

## Excel autofind drop down menu

#### Create a Named range

To manage the named ranges, you can go in **Formulas** > **Name Manager**. You can also use defined range by their name in Excel formulas.

##### 1. Static
To [create a named range](https://support.office.com/fr-fr/article/Cr%C3%A9er-une-liste-d%C3%A9roulante-7693307a-59ef-400a-b769-c5402dce407b) in excel, you can select a column of data then **righ click** then select **define name** the name will be the name of the range and how it will be referred to.
The **refers to** is the range itself and is autopopulated with the range of selected cells when clicking on **define name**

##### 2. Dynamic
To get a [dynamic named range](https://trumpexcel.com/named-ranges-in-excel/) you will need to replace the **refers to** of the named range by this kind of formula (for example if values are in the A column):

	=$A$2:INDEX($A$2:$A$100;COUNTIF($A$2:$A$100;"<>"&""))

This formula will start looking at value from `A2` to the index (the coordinates) of the last non empty cell (up to 100 in here).
It will only refers to the populated cells in the dynamic named range.


#### Get other information from the entered item
##### 1. Example:

- You have a range of value with a define name: `list` with all the values to find.
- You have the case where the search value is entered in `C4`. 

Then you can add this formula in the cells next to `C4` to map the cell using what has been entered in `C4`.

```
=INDIRECT("tab_name!"&ADDRESS(MATCH(C4;List;0)+1;COLUMN(List)-1))
```

- The `MATCH` function will match the `ROW` of the entered value (here `C4`) and the `List` value to get the right one.
- the `ADDRESS` function will map the found value and its relative position. (Used with `+1` or `-1` in the `ROW, COLUMN` you can modify the address you get. 
- the `INDIRECT` function print the value of the input coordinates (the `"tab_name!"` where the value is and the address of the found value).

##### 2. Another Example

Or you can use this formula which will look in `List` if it finds the value in `C4`:

```
=VLOOKUP(C4;List;2;FALSE)
```

#### Have a google like search

[Here](https://trumpexcel.com/excel-drop-down-list-with-search-suggestions/) is a sweet example that requires 1 column with the values and 3 helping columns and a cell that will be used to do the google like search:

| **E**. Available values | **F**. criteria matching | **G**. Occurence count | **H**. Found values |
|------------------|-------------------|-----------------|--------------|
| value_one        | 1                 | 1               | value_one    |
| value_two        | 0                 |                 | value_three  |
| value_three      | 1                 | 2               |              |

- Column #1 : **Available Values** you add the values that will be looked at
- Column #2 : **Criteria matching** you add this formula:

```
=--ISNUMBER(IFERROR(SEARCH($B$3,E3,1);""))
```

This formula returns 1 if part of what is in cell `E3` in the **Available values** coulumn is also in cell `B3`, the **search cell**.

- Column #3 : **Occurence count** you add this formula:

```
=IF(F3=1;COUNTIF($F$3:F3,1);"") 
```

This formula starting at `F3`, with `F3` the **criteria matching** look if the **criteria matching** is 1 and count how many there was since first cell (`$F$3`).

- Column #4 : **Found Values** stack all of the criteria matching values with this formula:

	=IFERROR(INDEX($E$3:$E$22,MATCH(ROWS($G$3:G3),$G$3:$G$22,0)),"")

With `G3` in the **Occurence count** column. It works with `MATCH` and `INDEX` looking for occurence. The `IFERROR` will show the corresponding value indexed, or nothing.

You can use this formula to create the dynamic range from the **found values** in `H3`:

```
=$H$3:INDEX($H$3:$H$22;MAX($G$3:$G$22);1)
```

The name will be used for the combobox (dropdown in developper > insert > activeX). Here are the properties to look for:

- AutoWordSelect: False
- LinkedCell: B3
- ListFillRange: name of the created named range
- MatchEntry: 2 – fmMatchEntryNone

The LinkedCell `B3` is the searching cell, it will print the result of the search.
If you haven't change the name of the combobox, the default one should be `ComboBox1` and you can copy paste that into the VBA part of your sheet:

```vb
Private Sub ComboBox1_Change()
'DropDownList is the name of the created Named range
ComboBox1.ListFillRange = "DropDownList"
Me.ComboBox1.DropDown
End Sub
```

This sub ComboBox1_change() overwrites the default attitude of the ComboBox object when changed.

## Other tips

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
