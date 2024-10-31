Attribute VB_Name = "Crisis_Survey"
Public Sub survey_MainMethod()
Dim totalPartcipants As Integer
Dim custNo As Integer
Dim iDate As Date
Dim rowNo As Integer
Dim colNo As Integer

rowNo = 2
colNo = 1
i = 1

    'how many participants
    totalPartcipants = RndBetween(600, 1300)
    
    'for each participate (rowNo)
    Do Until i = totalPartcipants
        ActiveWorkbook.Sheets("Survey 2").Cells(rowNo, colNo).Value = 1776 + i
        colNo = colNo + 1
        'for each col
        'need date
        ActiveWorkbook.Sheets("Survey 2").Cells(rowNo, colNo).Value = getDate
        colNo = colNo + 1
        'random for questions 1-9
        Do Until colNo = 12
            ActiveWorkbook.Sheets("Survey 2").Cells(rowNo, colNo).Value = RndBetween(1, 5)
            colNo = colNo + 1
        Loop
        colNo = 1
        rowNo = rowNo + 1
        i = i + 1
    Loop
    
    
    
    
    
End Sub
Private Function getDate() As String
Dim month As Integer
Dim day As Integer
Dim year As Integer

        'incident occurred on 03/16/18
        
        month = RndBetween(3, 6)
        day = GetDay(month)
        If month = 3 Then
            day = RndBetween(17, day)
        Else
            day = RndBetween(1, day)
        End If
        
        year = 2024
        
        getDate = month & "/" & day & "/" & year
End Function
