VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MainMenu 
   Caption         =   "Compare Open Files"
   ClientHeight    =   5730
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6480
   OleObjectBlob   =   "MainMenu.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "MainMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'2016-08-18: Added -EZ
Private Sub cbCompare_Click()
    If IsNull(lbNew.Value) Then
        MsgBox "Select NEW file.", vbExclamation, "Compare Macro"
        Exit Sub
    Else
        NewStr = lbNew.Value
    End If
    If IsNull(lbOld.Value) Then
        MsgBox "Select OLD file.", vbExclamation, "Compare Macro"
        Exit Sub
    Else
        OldStr = lbOld.Value
    End If
    If NewStr = OldStr Then
        MsgBox "Files can't be the same!"
        Exit Sub
    End If
    'Set globals and pass it to the main compare files function.
    gNewFilePath = Workbooks(lbNew.Value).FullName
    gOldFilePath = Workbooks(lbOld.Value).FullName
    MainMenu.Hide
    CompareFiles.CompareFiles
End Sub

Sub UserForm_Initialize()
    Dim wkBook As Workbook
    For Each wkBook In Workbooks
        lbNew.AddItem wkBook.Name
        lbOld.AddItem wkBook.Name
    Next
End Sub

