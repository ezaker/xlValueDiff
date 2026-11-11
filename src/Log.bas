Attribute VB_Name = "Log"
Option Explicit

Sub createLog()

    If Not GoptCreateLogFile Then Exit Sub

Dim wbkLog As Workbook
    
Set wbkLog = Workbooks.Add

With wbkLog
    .Title = "Diffs"
    .Subject = "Differences"
    .Sheets(1).Range("A1") = "Sheet"
    .Sheets(1).Range("B1") = "Col"
    .Sheets(1).Range("C1") = "Row"
    .Sheets(1).Range("D1") = "Old Value"
    .Sheets(1).Range("E1") = "New Value"
    If GoptAddLinksLog Then .Sheets(1).Range("F1") = "Link"
End With

gLogLine = 1

gLogName = wbkLog.Name

End Sub

Public Sub logDiff(strSheet As String, theCell As Range, ByVal strOld As String, ByVal strNew As String, addLinks As Boolean)
'2014-09-17: ByVal on strOld and strNew for passing array values. -EZ

Dim arrAddress As Variant
    
gLogLine = gLogLine + 1
arrAddress = Split(theCell.Address, "$")

With Workbooks(gLogName).Sheets(1)
    .Range("A" & gLogLine) = strSheet
    .Range("B" & gLogLine) = arrAddress(1)
    .Range("C" & gLogLine) = arrAddress(2)
    .Range("D" & gLogLine) = strOld
    .Range("E" & gLogLine) = strNew
    
    If addLinks Then
        Application.DefaultFilePath = theCell.Parent.Parent.Path
        .Hyperlinks.Add .Range("F" & gLogLine), "\\" & theCell.Parent.Parent.Path & "\" & theCell.Parent.Parent.Name, "'" & theCell.Parent.Name & "'!" & theCell.Address, , theCell.Address
    End If
End With

End Sub


Public Sub FinalizeLog()
    'Auto size columns, etc...
    If GoptCreateLogFile Then
        'Fit column widths on log
        With Workbooks(gLogName).Sheets(1)
            .Columns(1).AutoFit
            .Columns(2).AutoFit
            .Columns(3).AutoFit
            .Columns(4).AutoFit
            .Columns(5).AutoFit
            .Columns(6).AutoFit
        End With
    End If
End Sub
