Attribute VB_Name = "zz_Base_mod"
Option Private Module
Option Explicit

'2023-05-26: Fixed some outputs.

    Private Const EXPORT_DIR As String = "C:\Users\Xaker\Downloads\" 'End with \

Public Sub export()
    exportVBA ThisWorkbook
End Sub

'2019-10-14: Rewrote
Private Function exportVBA(ByRef wkb As Workbook) As Boolean

    Const SUB_NAME = "exportVBA()"
    
    Dim strPath As String
    Dim varComp As Variant

Try:
    On Error GoTo Catch

    strPath = EXPORT_DIR & Replace(wkb.Name, ".XLSM", "", , , vbTextCompare) & "\"

    If Not createDir(strPath) Then
        Debug.Print SUB_NAME, "Folder already exists. Export cancelled."
        GoTo Finally
    End If
    
    'Export all Forms and modules...
    For Each varComp In wkb.VBProject.VBComponents
        varComp.export strPath & varComp.Name & fileExt(varComp)
    Next varComp

    exportVBA = True
    Debug.Print SUB_NAME, "Export complete!"
Finally:
    Exit Function

Catch:
    Debug.Print SUB_NAME, Err.Description
    MsgBox Err.Description, vbCritical, SUB_NAME
    Resume Finally
End Function

'2019-10-14: New
Private Function fileExt(ByRef varComp As Variant) As String

    Const SUB_NAME = "fileExt()"

    'Component types
    Const ctModule As Integer = 1
    Const ctClassModule As Integer = 2
    Const ctForm As Integer = 3
    Const ctDocument As Integer = 100

Try:
    On Error GoTo Catch
        
        Select Case varComp.Type
            Case ctModule
                fileExt = ".bas"
            Case ctClassModule
                fileExt = ".cls"
            Case ctForm
                fileExt = ".frm"
            Case ctDocument
                fileExt = ".sht"
                ''' This is a worksheet or workbook object.
                ''' Don't try to export.
            Case Else
                'No extension
        End Select

Finally:
    Exit Function

Catch:
    Debug.Print SUB_NAME, Err.Description
    MsgBox Err.Description, vbCritical, SUB_NAME
    Resume Finally
End Function

Private Function getFilePath(ByRef wkb As Workbook) As String

    Const SUB_NAME = "fileExt()"

Try:
    On Error GoTo Catch

    getFilePath = Left(wkb.FullName, InStrRev(wkb.FullName, ".") - 1) & "\"
    
Finally:
    Exit Function

Catch:
    Debug.Print SUB_NAME, Err.Description
    MsgBox Err.Description, vbCritical, SUB_NAME
    Resume Finally
End Function

Private Function createDir(ByVal Name As String) As Boolean

    Const SUB_NAME = "createDir()"
    
    Dim folder As Variant
    
Try:
    On Error GoTo Catch
    
    folder = Dir(Name, vbDirectory)
    
    If folder = vbNullString Then
        MkDir Name
        createDir = True
    End If

Finally:
    Exit Function
    
Catch:
    Debug.Print SUB_NAME, Err.Description
    Resume Finally
End Function
