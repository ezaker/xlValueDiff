Attribute VB_Name = "OpenFiles"
Public Sub OpenOldFile() 'Lets user browse to the version of the file they want to use as the reference in the comparison
    OldStr = ""
    gOldFilePath = ""
    If Not OpenFile(OldStr, gOldFilePath, GrngOldFileCell, "Select OLD file...") Then Exit Sub
End Sub

Public Sub OpenNewFile() 'Lets user browse to the version of the file they want to compare to the reference file
    NewStr = ""
    gNewFilePath = ""
    If Not OpenFile(NewStr, gNewFilePath, GrngNewFileCell, "Select NEW file...") Then Exit Sub
End Sub

Public Function OpenFile(ByRef strBookName As String, ByRef strBookPath As String, ByRef rngBG As Range, Optional ByVal strMessage As String) As Boolean
    On Error GoTo handleError
    If IsEmpty(strBookName) Or strBookName = "" Then
        On Error Resume Next
        strBookPath = Application.GetOpenFilename("Worksheets (*.xlsx;*.xlsm;*.xlsb;*.xls),*.xlsx;*.xlsm;*.xlsb;*.xls", , strMessage, "Open File")
        Workbooks.Open strBookPath
        Application.ScreenUpdating = False
        If Not Err.number = 0 Then
            rngBG.Interior.Color = COLOR_BAD
            OpenFile = False
            MsgBox "Open failed. Please retry.", vbExclamation, VBA_TITLE
        Else
            OpenFile = True
            strBookName = ActiveWorkbook.Name
            rngBG.Interior.Color = COLOR_GOOD
            For Each sheet In Workbooks(strBookName).Sheets
                If sheet.Visible = False Then
                    sheet.Visible = True
                End If
            Next sheet
        End If
    Else
        If Not IsBookOpen(strBookName) Then
            Workbooks.Open strBookPath
            OpenFile = True
        Else
            OpenFile = True
        End If
    End If
    
    Application.ScreenUpdating = True
    Workbooks(MasterMacroStr).Activate
    
Exit Function
handleError:
    OpenFile = False
    

End Function

Private Function IsBookOpen(ByVal strBookName As String) As Boolean
    On Error GoTo handleError
    
    If Workbooks(strBookName).Name <> "" Then IsBookOpen = True
    
Exit Function
handleError:
    IsBookOpen = False

End Function

