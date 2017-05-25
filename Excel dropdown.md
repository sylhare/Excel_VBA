## Excel autofind drop down menu

#### Create a list
To [create a list](https://support.office.com/fr-fr/article/Cr%C3%A9er-une-liste-d%C3%A9roulante-7693307a-59ef-400a-b769-c5402dce407b) in excel, you can select a column of data then **righ click** then select **define name** the name will be the name of the list and how it will be referred to.

#### Get other information from the entered item
You have a `list` with all the values to find and you have in `C4` the case where the search value is entered. You can had this formula (with `+1` or `-1` in the `ROW, COLUMN` of the `ADDRESS` function. 
The `MATCH` function will match the `ROW` of the entered value and the `list` value.

	=INDIRECT("tab_name!"&ADDRESS(MATCH(C4;List;0)+1;COLUMN(List)-1))
	=INDIRECT("tab_name!"&B4)

Or you can use this formula which will look in the list t

	=VLOOKUP(C4,B8:C14,2,FALSE)