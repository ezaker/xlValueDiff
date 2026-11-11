Attribute VB_Name = "Reset"
Option Explicit
'Option Private Module

Public Sub Reset_main()
'2014-06-25: Setup macro workbook: remove highlighting, set scroll area, etc.

'SAVE EXCEL STATE HERE?

    Application.DisplayAlerts = False
    
    If Not IsEmpty(NewStr) Then
            NewStr = ""
    End If
    If Not IsEmpty(OldStr) Then
            OldStr = ""
    End If
        
    With wksMain
        .ScrollArea = "A1:K14"
        'Set status color areas
        Set GrngNewFileCell = .Range("B4:B8")
        Set GrngOldFileCell = .Range("E4:E8")
        Set GrngCompareFilesCell = .Range("E9:E13")
    End With
    
    'Clear status colors
    GrngNewFileCell.Interior.ColorIndex = xlNone
    GrngOldFileCell.Interior.ColorIndex = xlNone
    GrngCompareFilesCell.Interior.ColorIndex = xlNone
    
    Application.DisplayAlerts = True

End Sub

Public Sub GetOptions()

    'Colors
    With wksMain
        GoptHL = .Range("B9").Interior.Color
        GoptIncHL = .Range("B10").Interior.Color
        GoptRemHL = .Range("B11").Interior.Color
    End With
    
    'Checkbox
    GoptCreateLogFile = getBoxValue("Check Box 25")
    GoptAddLinksLog = getBoxValue("Check Box 26")
    GoptCommentNew = getBoxValue("Check Box 33")
    GoptRemoveComm = getBoxValue("Check Box 21")
    
    'Radio
    GoptMarkIncHL = getBoxValue("Option Button 28")
    GoptRemoveHL = getBoxValue("Option Button 29")
    GoptRemAllHL = getBoxValue("Option Button 30")

End Sub

Private Function getBoxValue(ByVal ShapeName As String) As Boolean
    getBoxValue = IIf(wksMain.Shapes(ShapeName).ControlFormat.Value = 1, True, False)
End Function
