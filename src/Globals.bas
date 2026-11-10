Attribute VB_Name = "Globals"
'********************************************************************************************************************************
'REVISION LOG
'********************************************************************************************************************************
'2014-09-17: AppendComment, CompareFiles, logDiff

'********************************************************************************************************************************
'GLOBAL CONSTANTS
'********************************************************************************************************************************
Public Const VBA_TITLE As String = "Compare Macro"
Public Const COLOR_BAD As Long = 459000
Public Const COLOR_GOOD As Long = 53000

'********************************************************************************************************************************
'DECLARE GLOBAL VARIABLES
'********************************************************************************************************************************
'Excel Application Settings
Public appCalculation As Integer
Public appDisplayPageBreaks As Boolean
Public appDisplayStatusBar As Boolean
Public appEnableEvents As Boolean
Public appScreenUpdating As Boolean

'Macro control
Public MasterMacroStr As String 'Macro's workbook name
Public MacroBook As Workbook
Public ControlPanel As Worksheet

Public GrngNewFileCell As Range 'Background of Open New File button
Public GrngOldFileCell As Range 'Background of Open Old File button
Public GrngCompareFilesCell As Range 'Background of Compare Files button

'Options
    'Colors
    Public GoptHL As Long 'Highlight color set by user.
    Public GoptIncHL As Long 'Color to highlight incorrectly marked cells.
    Public GoptRemHL As Long 'Color to remove
    
    'Checkbox
    Public GoptCreateLogFile As Boolean
    Public GoptAddLinksLog As Boolean
    Public GoptCommentNew As Boolean
    Public GoptRemoveComm As Boolean
    
    'Radio
    Public GoptMarkIncHL As Boolean
    Public GoptRemoveHL As Boolean
    Public GoptRemAllHL As Boolean


'New and old files
Public NewStr As String 'New workbook name
Public OldStr As String 'Old workbook name
Public gNewFilePath As String 'New file path
Public gOldFilePath As String 'Old file path

'Log file
'Public LogBook As Workbook
Public gLogName As String
Public gLogLine As Long 'Current line of log file.

Public sht As String
'********************************************************************************************************************************
