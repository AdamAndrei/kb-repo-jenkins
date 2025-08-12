'Table Function Library
 
'Created By: Juergen Kleinekorte
'Created On: 2017-05-16



Function selectRow(objTable, strSearch, intCol)
intRows = objTable.Object.getRowCount-1
For i = 0 To intRows
strName = objTable.GetCellData (i, intCol)
If strName = strSearch Then
objTable.SelectRow(i)
End If
Next
End Function 




