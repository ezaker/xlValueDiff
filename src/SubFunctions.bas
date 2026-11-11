Attribute VB_Name = "SubFunctions"
Option Explicit
Option Private Module

Public Sub ArrPush(ByRef anArray As Variant, ByVal anElement As Variant)
'Adds an element to the end of an array. 1D array only.
'2014-05-17: Added error support for empty arrays.

    Dim i As Long
    
    On Error Resume Next
    
    i = UBound(anArray)
    If Err.number = 0 Then
        ReDim Preserve anArray(i + 1)
        anArray(UBound(anArray)) = anElement
    Else
        ReDim anArray(0)
        anArray(0) = anElement
    End If

End Sub

Public Function IsArrayEmpty(ByVal anArray As Variant) As Boolean
'2014-06-25: Returns True if array is empty. -EZ
    
    On Error GoTo ErrEmpty
    If UBound(anArray) >= 0 Then IsArrayEmpty = False
    Exit Function

ErrEmpty:
    IsArrayEmpty = True

End Function

Public Function SheetExists(ByVal SheetName As String, ByRef Book As Workbook) As Boolean
'2014-06-25: Returns True if sheet is found

   On Error GoTo ErrEmpty
   If (Book.Sheets(SheetName).Name <> "") Then SheetExists = True
   Exit Function

ErrEmpty:
    SheetExists = False
   
End Function
