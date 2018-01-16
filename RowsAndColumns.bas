Attribute VB_Name = "RowsAndColumns"

' 
' Functions to returen the number of the last used rRow 
'

' First method: 
' UsedRange property to find the last used row number in a worksheet

Function LastUsedRow_1(MySheet As Worksheet) As Long

  Dim lastRow As Long

  lastRow = MySheet.UsedRange.Row - 1 + ActiveSheet.UsedRange.Rows.Count

  LastUsedRow_1 = lastRow

End Function

' Second method: 
' UsedRange property to find the last used row number in a worksheet

Function LastUsedRow_2(MySheet As Worksheet) As Long

  Dim lastRow As Long

  lastRow = MySheet.UsedRange.Rows(ActiveSheet.UsedRange.Rows.Count).Row

  LastUsedRow_2 = lastRow

End Function
