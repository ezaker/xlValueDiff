Attribute VB_Name = "CompareFiles"
Option Explicit
'Option Private Module

'Checks each value in each sheets to same cell in second workbook
'Option Explicit
'2014-09-17: If same value, then logs oldVal, newVal [NO CHANGE]. -EZ
'2016-08-18: Changed max highlights to 100,000.
'   No problem running on test file. Commenting is slow. Froze above 64000; left at 1000.
Public Function CompareFiles_Main(ByRef wkbNew As Workbook, ByRef wkbOld As Workbook) As Boolean

    Const F_NAME As String = "CompareFiles_Main()"

Const TAB_COLOR_NOCHANGE As Long = xlNone
Const TAB_COLOR_ALLMARKED As Long = 49406
Const TAB_COLOR_MISMARKED As Long = 254

'Maxes per sheet
Const MAX_COMMENTS As Long = 1000
Const MAX_HIGHLIGHTS As Long = 100000
Const LOOPS_DOEVENTS As Integer = 1000
Const WARN_NUM_DIFFS As Integer = 500

    Dim Sht As String
    
    Dim Sheets_Added As String
    Dim Sheets_Removed As String


Dim startTime As Date

Dim endTime As Date
Dim runTime As String

'Objects
Dim sheet As Worksheet
Dim cellNew As Range

'Variables
Dim Cols As Long
Dim Rows As Long
Dim TotalCols As Long
Dim TotalRows As Long
Dim arrDiffs() As Variant
Dim arrSames() As Variant


Dim guiContinue As Integer
Dim nLoops As Long 'Loop counter for DoEvents
Dim newVal As String
Dim oldVal As String
Dim newVal2 As String
Dim oldVal2 As String
Dim rng As Range

Dim x As Long

Dim RngAddress As String

Dim arrChn() As Variant
Dim xRow As Long
Dim xCol As Long
Dim lngFalseHL As Long 'False highlights: cell is colored but value is the same.

'Counts
    'Diffs: cell changes
    'UnHL: changes not highlighted
    'FalseHL: cells highlighted without changes
Dim Book_Diffs As Long
Dim Book_UnHLs As Long
Dim Book_FalseHLs As Long
Dim Sheet_Diffs As Long
Dim Sheet_UnHLs As Long
Dim Sheet_FalseHLs As Long

Dim arrNewFormats() As Variant




Dim strNewVal As String
Dim strOldVal As String

Dim arrNewVals() As Variant
Dim arrOldVals() As Variant


    On Error GoTo Catch
    
Try:
    startTime = Time

'Macro safe states
    Application.StatusBar = "Initializing"
    Application.Cursor = xlWait
    Application.EnableEvents = False
    Application.ScreenUpdating = False


'Loop through all sheets in New workbook
For Each sheet In wkbNew.Sheets
    
    'Clear filters, unhide columns and rows.
    On Error Resume Next
    sheet.ShowAllData
    sheet.Cells.EntireColumn.Hidden = False
    sheet.Cells.EntireRow.Hidden = False
    
On Error GoTo Catch
    
'Why not reference sheet instead of by name?
    Sht = sheet.Name
    Application.StatusBar = "Comparing " & Sht & "..."
        
    'See if sheet exists in old file.
    If Not SheetExists(Sht, wkbOld) Then
        'See if sheet is marked
        If (sheet.Tab.ColorIndex = xlNone) Then
            sheet.Tab.Color = g_Options.Color_Highlight
        End If
'PUSH TO ARRAY!!!
        If Sheets_Added <> "" Then
            Sheets_Added = Sheets_Added & Chr(10)
        End If
        Sheets_Added = Sheets_Added & Sht
        
    Else 'Sheet also exists in old workbook, check cells...
    
        'Reset variables
        Erase arrDiffs
        Erase arrSames
        Sheet_Diffs = 0
        Sheet_UnHLs = 0
        Sheet_FalseHLs = 0
        
        'Remove highlight colors
        If (g_Options.Mode_RemoveAll) Then
            sheet.Tab.ColorIndex = xlNone 'Remove all highlight colors
        ElseIf (g_Options.Mode_RemoveHL) Then 'Remove highlight color
            If sheet.Tab.Color = g_Options.Color_Highlight Then sheet.Tab.ColorIndex = xlNone
        End If
            
'       Union old and new sheet ranges (area only, not object)
        RngAddress = unionUsedRanges(sheet.UsedRange, wkbOld.Sheets(Sht).UsedRange)
        Set rng = sheet.Range(RngAddress)
        'Set rng = Union(sheet.Range("A1"), sheet.UsedRange, sheet.Range(Workbooks(OldStr).Sheets(sht).UsedRange.Address))
        
        
        TotalCols = rng.Columns.Count
        TotalRows = rng.Rows.Count
                    
'           Pull data into arrays

        If rng.Count > 1 Then
            arrNewVals = rng.Value
            arrOldVals = wkbOld.Sheets(Sht).Range(rng.Address).Value
        Else 'Only one cell
            ReDim arrNewVals(1, 1)
            ReDim arrOldVals(1, 1)
            arrNewVals(1, 1) = rng.Value
            arrOldVals(1, 1) = wkbOld.Sheets(Sht).Range(rng.Address).Value
        End If
'Pull array of highlights?
        
'COMPARE    Compare all elements, populate differences array
        For Rows = 1 To TotalRows
            For Cols = 1 To TotalCols
                strNewVal = getVal(arrNewVals(Rows, Cols))
                strOldVal = getVal(arrOldVals(Rows, Cols))
                If strNewVal <> strOldVal Then 'DIFFERENCE
                    ArrPush arrDiffs, Array(Rows, Cols, strOldVal, strNewVal)   'List of differences.
                Else 'NO DIFFERENCE
                    ArrPush arrSames, Array(Rows, Cols) 'Cell cell with matching values
                End If
            Next Cols
        Next Rows
        
'DIFFERENCES>
        'If no differences...
        If Not IsArrayEmpty(arrDiffs) Then
        
            Sheet_Diffs = UBound(arrDiffs) + 1
            
            Application.StatusBar = "Commenting/highlighting differences..."
        
            If Sheet_Diffs >= WARN_NUM_DIFFS Then
                guiContinue = MsgBox("There are " & Sheet_Diffs & " differences. Commenting/highlighting changes can take a while and lock up Excel. Continue and mark all entries?", vbExclamation + vbYesNo, "Compare Files Macro")
                'If guiContinue = vbNo Then Exit For
            End If
            
            
            If guiContinue <> vbNo Then
                'Go through differences
                For x = 0 To UBound(arrDiffs)
                    xRow = arrDiffs(x)(0)
                    xCol = arrDiffs(x)(1)
                    strOldVal = CStr(arrDiffs(x)(2))
                    strNewVal = CStr(arrDiffs(x)(3))
                    
                    'ROWS & COLS in the ARRAY ARE NOT THE SAME IN THE SHEET. rng might be offset!!!!!
                    'is union broken too? web site claims it duplicates.
                    
                    If Not IsHL(rng.Cells(xRow, xCol)) Then
                        Sheet_UnHLs = Sheet_UnHLs + 1
'                        sheet.Tab.Color = 254 ' WHAT WAS THIS FOR AGAIN?
                    End If
                    
                    'Comment cell
                    If g_Options.CommentNew Then
                        If x <= MAX_COMMENTS Then
'Does this highlight over the original color?
                            AppendComment rng.Cells(xRow, xCol).Address, strOldVal, g_Options.Color_Highlight, wkbNew, Sht
                            
                        ElseIf x = (MAX_COMMENTS + 1) Then
                            MsgBox "The max amount of comments (" & MAX_COMMENTS & ") has been reached. " & _
                                "Further cells will not be commented, but all differences are logged.", _
                                vbInformation, VBA_TITLE
                        End If
                    End If
                    
                    If (Not (g_Options.CommentNew) Or (x > MAX_COMMENTS)) And (x < MAX_HIGHLIGHTS) Then
                            If Not IsHL(rng.Cells(xRow, xCol)) Then
                                rng.Cells(xRow, xCol).Interior.Color = g_Options.Color_Highlight
                            End If
                    ElseIf x = MAX_HIGHLIGHTS Then
                        MsgBox "The max amount of highlights (" & MAX_HIGHLIGHTS & ") has been reached. " & _
                            "Further cells will not be highlighted, but all differences are logged.", _
                            vbInformation, VBA_TITLE
                    End If
                    
                    'Log Differences
'                    If GoptCreateLogFile Then
'                        logDiff Sht, rng.Cells(xRow, xCol), strOldVal, strNewVal, GoptAddLinksLog
'                    End If
                    
                    'DoEvents after certain amount of loops; keeps Excel responding.
                    If nLoops = LOOPS_DOEVENTS Then
                        'DoEvents
                        nLoops = 0
                    End If
                    
                    nLoops = nLoops + 1
                Next x
            End If
        End If
'<DIFFERENCES
        
'SAME
            'Only run if options call for it. Detect for false positive?
            If Not IsArrayEmpty(arrSames) Then
            
                Application.StatusBar = "Checking similarities..."

                For x = 0 To UBound(arrSames)
                    xRow = arrSames(x)(0)
                    xCol = arrSames(x)(1)
                    
                    'If cell is highlighted, it is marked incorrectly.
                    If IsHL(rng.Cells(xRow, xCol)) Then
                        Sheet_FalseHLs = Sheet_FalseHLs + 1
                        If g_Options.Mode_Incorrect Then rng.Cells(xRow, xCol).Interior.Color = g_Options.Color_Incorrect

                        'Log bad highlights
'                        If GoptCreateLogFile Then
'                            logDiff Sht, rng.Cells(xRow, xCol), arrOldVals(xRow, xCol), arrNewVals(xRow, xCol) & " [NO CHANGE]", GoptAddLinksLog '2014-09-17
'                        End If
                    End If
                        
                    'Remove all highlight colors?
                    If (g_Options.Mode_RemoveAll) Then
                        rng.Cells(xRow, xCol).Interior.ColorIndex = xlNone
                    ElseIf (g_Options.Mode_RemoveHL) Then 'Remove selected highlight color.
                        If rng.Cells(xRow, xCol).Interior.Color = g_Options.Color_Remove Then rng.Cells(xRow, xCol).Interior.ColorIndex = xlNone
                    End If
                    
                    'Remove comments?
                    If (g_Options.RemoveComments) And Not (rng.Cells(xRow, xCol).Comment Is Nothing) Then
                        rng.Cells(xRow, xCol).Comment.Delete
                    End If

                    'DoEvents after certain amount of loops; keeps Excel responding.
                    If nLoops = LOOPS_DOEVENTS Then
                        'DoEvents
                        nLoops = 0
                    End If
                    
                    nLoops = nLoops + 1
                Next x
            End If
        
        'Color tabs
        If Sheet_Diffs = 0 Then
            sheet.Tab.Color = TAB_COLOR_NOCHANGE
        Else
            If (Sheet_UnHLs > 0) Or (Sheet_FalseHLs > 0) Then
                sheet.Tab.Color = TAB_COLOR_MISMARKED
            Else
                sheet.Tab.Color = TAB_COLOR_ALLMARKED
            End If
        End If
        
        'Update totals
        Book_Diffs = Book_Diffs + Sheet_Diffs
        Book_UnHLs = Book_UnHLs + Sheet_UnHLs
        Book_FalseHLs = Book_FalseHLs + Sheet_FalseHLs
        
    End If
    
Next sheet



'Find sheets removed from Old workbook
On Error Resume Next
    For Each sheet In wkbOld.Sheets
        If (Trim(sheet.Name) <> Trim(wkbNew.Sheets(sheet.Name).Name)) Then
            If Sheets_Removed <> "" Then Sheets_Removed = Sheets_Removed & Chr(10)
            Sheets_Removed = Sheets_Removed & sheet.Name
        End If
    Next sheet

On Error GoTo Catch

makeSameLen Sheets_Added, Sheets_Removed 'Make lists same length to display correctly


'Arrange windows
wkbXLvalueDiff.Activate
wkbOld.Activate
wkbNew.Activate

DoEvents

'Calculate run time
endTime = Time
runTime = (endTime - startTime) * 24 * 3600
runTime = Format(runTime, "0.00")

With Report
    .ndeltas = Book_Diffs
    .ndif = Book_UnHLs
    .tbIncHL = Book_FalseHLs
    .NewSheets = Sheets_Added
    .OldSheets = Sheets_Removed
    .NewSheets.Width = 160
    .OldSheets.Width = 160
    If .NewSheets.Height >= .OldSheets.Height Then
        .Height = .NewSheets.Height + 100
    Else
        .Height = .OldSheets.Height + 100
    End If
    .rantime.Caption = "Runtime: " & runTime
    .rantime.Left = .Width - .rantime.Width - 10
    .rantime.Top = .Height - .rantime.Height - 25
    .Caption = "Report"
    .StartUpPosition = 2
    .Show vbModeless
End With

Finally:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Cursor = xlDefault
    Application.StatusBar = False
    
    Exit Function

Catch:
    MsgBox "Any error. Close, don't save and retry. " & vbLf & vbLf & Err.Description, vbCritical, F_NAME
    Resume Finally

End Function

Private Function getVal(strValue As Variant) As String

If IsError(strValue) = True Then
    Select Case strValue
        Case CVErr(xlErrDiv0)
            getVal = "#DIV/0!"
        Case CVErr(xlErrNA)
            getVal = "#N/A"
        Case CVErr(xlErrName)
            getVal = "#NAME?"
        Case CVErr(xlErrNull)
            getVal = "#NULL!"
        Case CVErr(xlErrNum)
            getVal = "#NUM!"
        Case CVErr(xlErrRef)
            getVal = "#REF!"
        Case CVErr(xlErrValue)
            getVal = "#VALUE!"
        Case Else
            getVal = "#ERROR!"
    End Select
Else
    getVal = strValue
End If

End Function

Public Function IsHL(ByRef aRange As Range) As Boolean
    On Error GoTo err_IsHL
    
    Const COLOR_WHITE As Integer = 2
    
    If (aRange.Interior.ColorIndex = xlNone) Or (aRange.Interior.ColorIndex = COLOR_WHITE) Then
        IsHL = False
    Else
        IsHL = True
    End If
            
'DO WE NEED TO CHECK COLOR AND NOT INDEX? SHOULD WE PICK NOT SO EASY TO SEE COLORS?
    
Exit Function
    
err_IsHL:
    IsHL = False

End Function

Private Function unionUsedRanges(ByRef Range1 As Range, ByRef Range2 As Range) As String
'Takes new range and old range. Combines them into a simple range (A1:LastCol,LastRow)
    
    Dim Address1 As String
    Dim Address2 As String
    Dim AddressList() As Variant
    Dim Addr As Variant
    Dim CellList As Variant
    Dim CellAddr As Variant
    Dim CellAddrParts As Variant
        
    Dim RowNum As Variant
    Dim RowNums() As Variant
    Dim ColNum As Variant
    Dim ColNums() As Variant
    
    Dim LastRow As Long
    Dim LastCol As Long
    
    Dim LastCellAddr As String
    Dim RangeAddress As String
    
    Address1 = Range1.Address(ReferenceStyle:=xlR1C1)
    Address2 = Range2.Address(ReferenceStyle:=xlR1C1)
    AddressList = Array(Address1, Address2)

    For Each Addr In AddressList
        CellList = Split(Addr, ":")
        For Each CellAddr In CellList
            CellAddrParts = Split(CellAddr, "C")
            ArrPush ColNums, CellAddrParts(1)
            ArrPush RowNums, Replace(CellAddrParts(0), "R", "")
        Next CellAddr
    Next Addr
    
    'Find last row
    For Each RowNum In RowNums
        If CLng(RowNum) > LastRow Then LastRow = CLng(RowNum)
    Next RowNum
    
    'Find last column
    For Each ColNum In ColNums
        If CLng(ColNum) > LastCol Then LastCol = CLng(ColNum)
    Next ColNum
    
    LastCellAddr = Range1.Parent.Cells(LastRow, LastCol).Address
    unionUsedRanges = "$A$1:" & LastCellAddr

End Function

Private Sub makeSameLen(ByRef List1 As String, ByRef List2 As String)

    Dim arrLines1 As Variant
    Dim arrLines2 As Variant
    Dim extraLines As Long
    Dim x As Long

    arrLines1 = Split(List1, Chr(10))
    arrLines2 = Split(List2, Chr(10))
    
    extraLines = UBound(arrLines1) - UBound(arrLines2)
        
    For x = 1 To Abs(extraLines)
        If (extraLines > 0) Then
            List2 = List2 & Chr(10)
        Else
            List1 = List1 & Chr(10)
        End If
    Next

End Sub
