Attribute VB_Name = "RunnerSheet"
Option Explicit

Private Const CRUNNER_SHEET_NAME_PREFIX = "RUNNER("
Private Const CRUNNER_SHEET_NAME_SUFFIX = ")"

Private Const CCOL_OFFSET_PK As Long = 0        ' <Primary Key> = <Tag> + ">>>" + <Bib>
Private Const CCOL_OFFSET_BIB As Long = 1       ' Bib
Private Const CCOL_OFFSET_TAG As Long = 2       ' Tag
Private Const CCOL_OFFSET_LOCKED As Long = 3    ' Locked
Private Const CCOL_OFFSET_NAME As Long = 4      ' Name
Private Const CCOL_OFFSET_TEAM As Long = 5      ' Team
Private Const CCOL_OFFSET_REMARKS As Long = 6   ' Remarks

Private Function FormatRunnerSheetName(aStartListName As String) As String
    FormatRunnerSheetName = CRUNNER_SHEET_NAME_PREFIX & aStartListName & CRUNNER_SHEET_NAME_SUFFIX
End Function

Public Function ParseRunnerSheetName(aSheetName As String) As String
    Dim nPrefix As Long
    Dim nSuffix As Long
    nPrefix = Len(CRUNNER_SHEET_NAME_PREFIX)
    nSuffix = Len(CRUNNER_SHEET_NAME_SUFFIX)
    If Left(aSheetName, nPrefix) = CRUNNER_SHEET_NAME_PREFIX _
    And Right(aSheetName, nSuffix) = CRUNNER_SHEET_NAME_SUFFIX Then
        ParseRunnerSheetName = Mid(aSheetName, nPrefix + 1, Len(aSheetName) - nPrefix - nSuffix)
    End If
End Function

Private Function IsRunnerSheetName(aSheetName As String) As Boolean
    IsRunnerSheetName = ParseRunnerSheetName(aSheetName) <> ""
End Function

Public Function GetRunnerSheet(aHeatName As String) As Worksheet
On Error Resume Next
    Set GetRunnerSheet = ThisWorkbook.Sheets(FormatRunnerSheetName(aHeatName))
End Function

Public Function ExistsRunnerSheet(aHeatName As String) As Boolean
    ExistsRunnerSheet = Not GetRunnerSheet(aHeatName) Is Nothing
End Function

Public Function CreateNewRunnerSheet(aHeatName As String) As Worksheet
On Error GoTo ErrHandler
    Dim sht As Worksheet
    Dim rng As Range
    Set sht = ThisWorkbook.Sheets.Add()
    sht.Name = FormatRunnerSheetName(aHeatName)
    
    Set rng = sht.Range("A1")
    
    With rng
        .Offset(0, CCOL_OFFSET_PK) = "#"
        .Offset(0, CCOL_OFFSET_BIB) = "Bib"
        .Offset(0, CCOL_OFFSET_TAG) = "Tag"
        .Offset(0, CCOL_OFFSET_LOCKED) = "Locked"
        .Offset(0, CCOL_OFFSET_NAME) = "Name"
        .Offset(0, CCOL_OFFSET_TEAM) = "Team"
        .Offset(0, CCOL_OFFSET_REMARKS) = "Remarks"
    End With
    
    ' FreezePanes - ROW:1
    sht.Parent.Activate
    sht.Activate
    rng.Offset(1, 0).Select
    ActiveWindow.FreezePanes = True
    
    Set CreateNewRunnerSheet = sht
    
ExitProc:
    Exit Function
ErrHandler:
    If Not sht Is Nothing Then
        Application.DisplayAlerts = False
        sht.Delete
        Application.DisplayAlerts = True
    End If
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Function

Public Function GetRunners(aHeatName As String) As Collection
    Dim sht As Worksheet
    Dim rng As Range
    Dim cList As New Collection
    Dim dItem As Dictionary
    Dim rowIndex As Long
    
    Set sht = GetRunnerSheet(aHeatName)
    
    Set rng = sht.Range("A2")
    rowIndex = 0
    Do
        With rng.Offset(rowIndex, 0)
            If .Offset(0, CCOL_OFFSET_BIB).Text = "" Then
                .Offset(0, CCOL_OFFSET_PK) = ""
                GoTo ExitProc
            End If
            .Offset(0, CCOL_OFFSET_PK) = .Offset(0, CCOL_OFFSET_BIB).Text & ">>>" & .Offset(0, CCOL_OFFSET_TAG).Text
            Set dItem = New Dictionary
            dItem.Add "row", rowIndex
            dItem.Add "bib", .Offset(0, CCOL_OFFSET_BIB).Text
            dItem.Add "tag", .Offset(0, CCOL_OFFSET_TAG).Text
            dItem.Add "locked", .Offset(0, CCOL_OFFSET_LOCKED).Text
            cList.Add dItem
        End With
        rowIndex = rowIndex + 1
    Loop
ExitProc:
    Set GetRunners = cList
End Function
