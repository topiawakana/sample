Attribute VB_Name = "M_ExcelFunction"
Function LastRow(sheetobj As Worksheet, C)
LastRow = sheetobj.Cells(sheetobj.Rows.Count, C).End(xlUp).Row
End Function

Function LastCol(sheetobj As Worksheet, R)
LastCol = sheetobj.Cells(R, sheetobj.Columns.Count).End(xlToLeft).Column
End Function

