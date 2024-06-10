Attribute VB_Name = "Age"
Public Sub GetAge()
Dim totalNumberOfNames As Integer
Dim doX As Integer
    
    MakeHeader "Age"
    doX = 0
    'get total number of names
    totalNumberOfNames = ActiveWorkbook.Sheets("Names").Cells(3, "F").Value + ActiveWorkbook.Sheets("Names").Cells(4, "F").Value
    
    'get age for each name
    Do Until doX = totalNumberOfNames '= 0
        ActiveWorkbook.Sheets(newSheetName).Cells(doX + 2, colCount).Value = _
            RndBetween(ActiveWorkbook.Sheets("Names").Cells(9, "F").Value, ActiveWorkbook.Sheets("Names").Cells(9, "G").Value)
        'totalNumberOfNames = totalNumberOfNames - 1
        doX = doX + 1
    Loop
    
    
End Sub

