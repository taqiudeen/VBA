Attribute VB_Name = "Workbook_Close"
'10/05/18
'09/18/24

Option Explicit
Public Sub CloseDuties()

        'ADD to ThisWorkbook
        'Private Sub Workbook_BeforeClose(Cancel As Boolean)
        '    CloseDuties
        'End Sub
        
    BackUP
    SaveWithoutDisplay
    Application.Quit
End Sub
Public Sub BackUP()
Dim saveDis As String

    saveDis = MsgBox("Do you want to back up?", vbYesNo + vbQuestion, "Back Up?")
    
    If saveDis = vbYes Then
        Application.DisplayAlerts = False
        ActiveWorkbook.SaveCopyAs GetFileName(0)
    End If
    
End Sub
Private Function GetFileName(dayOffset)

Dim folderName As String
Dim fileName As String
    
    folderName = "Flow"
    fileName = "Flow"
    
    EnsureDesktopFolderExists folderName, Environ$("USERPROFILE") & "\Desktop\" & folderName
       
    GetFileName = Environ$("USERPROFILE") & "\Desktop\" & folderName & "\" & fileName & "_" _
         & WorksheetFunction.Text(Month(Date), "00") & "_" & WorksheetFunction.Text(Day(Date), "00") & "_" & Year(Date) & ".xlsm"
    
End Function

Sub testSave()
    ActiveWorkbook.SaveCopyAs GetFileName2(0)
End Sub

Private Sub DashFirst()
    Sheets("Dash").Select
End Sub

Public Sub SaveWithoutDisplay()
    Application.DisplayAlerts = False
    ActiveWorkbook.Save
    Application.DisplayAlerts = True

End Sub

Public Sub EnsureDesktopFolderExists(ByVal folderName As String, path As String)
    
    If Len(Dir(path, vbDirectory)) = 0 Then MkDir path & folderName
End Sub
