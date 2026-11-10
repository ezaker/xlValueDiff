Attribute VB_Name = "Comment"
Public Sub AppendComment(Addr As String, str As String, Optional Color As Long, Optional wkb As Variant, Optional FCMsheet As Variant)
'Adds comment, if there already was a comment it adds to the end
'2014-09-17: Changed color check to use IsHL to account for white. -EZ

    Dim old As String
    On Error Resume Next
    If IsMissing(wkb) Then
        wkb = NewUDSstr
    End If
    If IsMissing(FCMsheet) Then
        appendSht = sht
    Else: appendSht = FCMsheet
    End If
    With Workbooks(wkb).Sheets(appendSht).Range(Addr)
    If Not .Comment Is Nothing Then
        old = .Comment.Text
        If InStr(1, old, str, vbTextCompare) = 0 Then
            old = old & Chr(10) & Chr(10) & str
            .Comment.Delete
            .addComment Text:=old
        End If
        If Not (.Comment) Is Nothing Then
            .Comment.Shape.TextFrame.AutoSize = True
            .Comment.Shape.TextFrame.Characters.Font.Size = 12
        End If
    Else
        .addComment Text:=str
        If Not .Comment Is Nothing Then
            .Comment.Shape.TextFrame.AutoSize = True
            .Comment.Shape.TextFrame.Characters.Font.Size = 12
        End If
    End If
    If Color <> 0 Then
        If Not IsHL(Workbooks(wkb).Sheets(appendSht).Range(Addr)) Then '2014-09-17
            .Interior.Color = Color
        End If
    End If
    End With
End Sub
