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
Option Explicit

Public Sub Init()
    
    Dim wkBook As Workbook
    
    lbNew.Clear
    lbOld.Clear

    For Each wkBook In Workbooks
        If wkBook.Name <> ThisWorkbook.Name Then
            If wkBook.Windows.Count > 0 Then
                If wkBook.Windows(1).Visible Then
                    lbNew.AddItem wkBook.Name
                    lbOld.AddItem wkBook.Name
                End If
            End If
        End If
    Next
    
    Me.Show
    
End Sub

Private Sub cbCompare_Click()

    Dim strNewWkbName As String
    Dim strOldWkbName As String

    If IsNull(lbNew.Value) Then
        MsgBox "Select NEW file.", vbExclamation, "Compare Macro"
        Exit Sub
    Else
        strNewWkbName = lbNew.Value
    End If
    If IsNull(lbOld.Value) Then
        MsgBox "Select OLD file.", vbExclamation, "Compare Macro"
        Exit Sub
    Else
        strOldWkbName = lbOld.Value
    End If
    If strNewWkbName = strOldWkbName Then
        MsgBox "Files can't be the same!"
        Exit Sub
    End If
    
    MainMenu.Hide
    CompareFiles_Main Workbooks(strNewWkbName), Workbooks(strOldWkbName)
End Sub
