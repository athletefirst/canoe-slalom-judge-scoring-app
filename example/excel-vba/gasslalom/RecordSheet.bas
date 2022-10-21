Attribute VB_Name = "RecordSheet"
Option Explicit

Private Const CRECORD_SHEET_NAME_PREFIX = "RECORD("
Private Const CRECORD_SHEET_NAME_SUFFIX = ")"

Private Const CPK_DELIMITER = ">>>"

Private Const CCOL_OFFSET_PK As Long = 0        ' <Primary Key> = <Tag> + ">>>" + <Bib>
Private Const CCOL_OFFSET_BIB As Long = 1       ' Bib
Private Const CCOL_OFFSET_TAG As Long = 2       ' Tag
Private Const CCOL_OFFSET_TIME As Long = 3      ' Time
Private Const CCOL_OFFSET_PENALTY As Long = 4   ' Penalty
Private Const CCOL_OFFSET_POINT As Long = 5     ' Point
Private Const CCOL_OFFSET_START As Long = 6     ' Start
Private Const CCOL_OFFSET_START2 As Long = 7    ' Start
Private Const CCOL_OFFSET_FINISH As Long = 8    ' Finish
Private Const CCOL_OFFSET_FINISH2 As Long = 9   ' Finish
Private Const CCOL_OFFSET_F50 As Long = 10      ' Team Penalty
Private Const CCOL_OFFSET_GATE As Long = 11     ' Gate

Private Function FormatRecordSheetName(aStartListName As String) As String
    FormatRecordSheetName = CRECORD_SHEET_NAME_PREFIX & aStartListName & CRECORD_SHEET_NAME_SUFFIX
End Function

Public Function ParseRecordSheetName(aSheetName As String) As String
    Dim nPrefix As Long
    Dim nSuffix As Long
    nPrefix = Len(CRECORD_SHEET_NAME_PREFIX)
    nSuffix = Len(CRECORD_SHEET_NAME_SUFFIX)
    If Left(aSheetName, nPrefix) = CRECORD_SHEET_NAME_PREFIX _
    And Right(aSheetName, nSuffix) = CRECORD_SHEET_NAME_SUFFIX Then
        ParseRecordSheetName = Mid(aSheetName, nPrefix + 1, Len(aSheetName) - nPrefix - nSuffix)
    End If
End Function

Private Function IsRecordSheetName(aSheetName As String) As Boolean
    IsRecordSheetName = ParseRecordSheetName(aSheetName) <> ""
End Function

Public Function GetRecordSheet(aHeatName As String) As Worksheet
On Error Resume Next
    Set GetRecordSheet = ThisWorkbook.Sheets(FormatRecordSheetName(aHeatName))
End Function

Public Function ExistsRecordSheet(aHeatName As String) As Boolean
    ExistsRecordSheet = Not GetRecordSheet(aHeatName) Is Nothing
End Function

Public Function CreateNewRecordSheet(aHeatName As String) As Worksheet
On Error GoTo ErrHandler
    Dim sht As Worksheet
    Dim rng As Range
    Dim i As Long
    
    Set sht = ThisWorkbook.Sheets.Add()
    sht.Name = FormatRecordSheetName(aHeatName)
    
    Set rng = sht.Range("A1")
    
    With rng
        .Offset(0, CCOL_OFFSET_PK) = "#"
        .Offset(0, CCOL_OFFSET_BIB) = "Bib"
        .Offset(0, CCOL_OFFSET_TAG) = "Tag"
        .Offset(0, CCOL_OFFSET_TIME) = "Time"
        sht.Columns(CCOL_OFFSET_TIME + 1).NumberFormatLocal = "#,##0.000_ "
        .Offset(0, CCOL_OFFSET_PENALTY) = "Penalty"
        .Offset(0, CCOL_OFFSET_POINT) = "Point"
        sht.Columns(CCOL_OFFSET_POINT + 1).NumberFormatLocal = "#,##0.000_ "
        .Offset(0, CCOL_OFFSET_F50) = "Team Penalty"
        .Offset(0, CCOL_OFFSET_START) = "Started"
        .Offset(0, CCOL_OFFSET_START2) = "Started Time"
        sht.Columns(CCOL_OFFSET_START2 + 1).NumberFormatLocal = "#,##0.000_ "
        .Offset(0, CCOL_OFFSET_FINISH) = "Finished"
        .Offset(0, CCOL_OFFSET_FINISH2) = "Finished Time"
        sht.Columns(CCOL_OFFSET_FINISH2 + 1).NumberFormatLocal = "#,##0.000_ "
        For i = 1 To 30
            .Offset(0, CCOL_OFFSET_GATE + i - 1) = "G" & Format(i, "00")
        Next i
    End With
    
    ' FreezePanes - from "#" to "Point"
    sht.Parent.Activate
    sht.Activate
    rng.Offset(1, CCOL_OFFSET_START).Select
    ActiveWindow.FreezePanes = True
    
    Set CreateNewRecordSheet = sht
    
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

Public Sub PutRecords(aHeatName As String, aList As Collection)
    Dim sht As Worksheet
    Dim rng As Range
    Dim rngRow As Range
    Dim v As Variant
    Dim g As Variant
    Dim rowIndex As Long
    Dim nGate As Long
    
    Dim sJudgeS As String
    Dim sJudgeF As String
    Dim sJudgeG As String
    Dim nStartedTime As Variant
    Dim nFinishedTime As Variant
    Dim nPenalty As Long
    Dim vTime As Variant
    Dim vJudge As Variant
    
    Set sht = GetRecordSheet(aHeatName)
    
    Set rng = sht.Range("A2")
    For Each v In aList
        sJudgeS = ""
        sJudgeF = ""
        sJudgeG = ""
        nStartedTime = CDec(0)
        nFinishedTime = CDec(0)
        nPenalty = 0
        vTime = ""
        vJudge = ""
        
        rowIndex = v.Item("runner").Item("row")
        Set rngRow = rng.Offset(rowIndex, 0)
        With v
            ' runner
            With .Item("runner")
                rngRow.Offset(0, CCOL_OFFSET_PK) = .Item("bib") & ">>>" & .Item("tag")
                rngRow.Offset(0, CCOL_OFFSET_BIB) = .Item("bib")
                rngRow.Offset(0, CCOL_OFFSET_TAG) = .Item("tag")
            End With
            
            ' started
            With .Item("started")
                sJudgeS = .Item("judge")
                If IsNumeric(.Item("time")) Then
                    nStartedTime = CDec(.Item("time"))
                End If
                rngRow.Offset(0, CCOL_OFFSET_START) = sJudgeS
                rngRow.Offset(0, CCOL_OFFSET_START2) = IIf(sJudgeS = CJUDGE_STARTED, nStartedTime, "")
            End With
            
            ' finished
            With .Item("finished")
                sJudgeF = .Item("judge")
                If IsNumeric(.Item("time")) Then
                    nFinishedTime = CDec(.Item("time"))
                End If
                rngRow.Offset(0, CCOL_OFFSET_FINISH) = sJudgeF
                rngRow.Offset(0, CCOL_OFFSET_FINISH2) = IIf(sJudgeF = CJUDGE_FINISHED Or sJudgeF = CJUDGE_FINISHED_50, nFinishedTime, "")
                rngRow.Offset(0, CCOL_OFFSET_F50) = IIf(sJudgeF = CJUDGE_FINISHED_50, 50, "")
            End With
            
            ' gates
            For Each g In .Item("gates")
                With g
                    nGate = .Item("num")
                    If IsNumeric(.Item("judge")) Then
                        nPenalty = nPenalty + CLng(.Item("judge"))
                    Else
                        If .Item("judge") = CJUDGE_DSQ Then
                            sJudgeG = CJUDGE_DSQ
                        End If
                    End If
                    rngRow.Offset(0, CCOL_OFFSET_GATE + nGate - 1) = .Item("judge")
                End With
            Next g
                        
            If sJudgeS = "" And sJudgeG = "" And sJudgeF = "" Then
                vJudge = ""
            ElseIf sJudgeS = CJUDGE_DNS Then
                vJudge = CJUDGE_DNS
            ElseIf sJudgeS = CJUDGE_DSQ Or sJudgeG = CJUDGE_DSQ Or sJudgeF = CJUDGE_DSQ Then
                vJudge = CJUDGE_DSQ
            ElseIf sJudgeF = CJUDGE_DNF Then
                vJudge = CJUDGE_DNF
            Else
                ' �����_�ȉ��Q�����L��
                vTime = Fix((nFinishedTime - nStartedTime) * 100) / 100
                vJudge = vTime + nPenalty
            End If
            
            rngRow.Offset(0, CCOL_OFFSET_TIME) = vTime
            rngRow.Offset(0, CCOL_OFFSET_PENALTY) = IIf(nPenalty = 0, "", nPenalty)
            rngRow.Offset(0, CCOL_OFFSET_POINT) = vJudge
            
        End With
    Next v
    
ExitProc:
End Sub
