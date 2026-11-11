VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clxOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ----- Colors -----
Public Property Get Color_Highlight() As Long
    Color_Highlight = wksMain.Range("B9").Interior.Color
End Property

Public Property Get Color_Incorrect() As Long
    Color_Incorrect = wksMain.Range("B10").Interior.Color
End Property

Public Property Get Color_Remove() As Long
    Color_Remove = wksMain.Range("B11").Interior.Color
End Property

' ----- Checkboxes -----
Public Property Get CommentNew() As Boolean
    CommentNew = (wksMain.Shapes("Check Box 33").ControlFormat.Value = xlOn)
End Property

Public Property Get RemoveComments() As Boolean
    RemoveComments = (wksMain.Shapes("Check Box 21").ControlFormat.Value = xlOn)
End Property

' ----- Radio buttons -----
Public Property Get Mode_Incorrect() As Boolean
    Mode_Incorrect = (wksMain.Shapes("Option Button 28").ControlFormat.Value = xlOn)
End Property

Public Property Get Mode_RemoveHL() As Boolean
    Mode_RemoveHL = (wksMain.Shapes("Option Button 29").ControlFormat.Value = xlOn)
End Property

Public Property Get Mode_RemoveAll() As Boolean
    Mode_RemoveAll = (wksMain.Shapes("Option Button 30").ControlFormat.Value = xlOn)
End Property
