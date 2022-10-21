Attribute VB_Name = "TestGasSlalom"
'
' MIT License
'
' Copyright (c) 2022 AthleteFirst TOKYO
'
' Permission is hereby granted, free of charge, to any person obtaining a copy
' of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights
' to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is
' furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all
' copies or substantial portions of the Software.
'
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
' OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
' SOFTWARE.
'

Option Explicit

Private mUrl As String

Public Sub TestSuite(aDummy As Variant)
    Dim ret As Variant
    Dim v As Variant
    Dim g As Variant
    Dim s As String
    
    Debug.Print vbCrLf & "--- TestSuite ---"
    
    ' API URL
    s = InputBox("Your Web App URL:", "TestSuite()", mUrl)
    If s = "" Then
        Debug.Print "canceled."
        Exit Sub
    End If
    mUrl = s
    Debug.Print mUrl
    InitUrl mUrl
    
    ' ExistsHeat()
    Debug.Print "ExistsHeat()"
    ret = ExistsHeat("Test Run")
    Debug.Print "ExistsHeat: "; ret
    
    ' AddHeat()
    Debug.Print "AddHeat()"
    ret = AddHeat("Test Run", True)
    If GetLastErrNumber() <> 0 Then
        Debug.Print GetLastErrDescription()
        Debug.Print GetLastErrSource()
        Debug.Assert False
    End If
    
    ' GetHeatsAll()
    Debug.Print "GetHeatsAll()"
    Set ret = GetHeatsAll()
    Debug.Assert Not ret Is Nothing
    Debug.Print JsonConverter.ConvertToJson(ret)
    For Each v In ret
        Debug.Print v("heatName")
    Next v
    
    ' PutGateSetting()/PutGateSettings()
    Debug.Print "PutGateSetting()"
    ret = PutGateSetting("Test Run", 1, "UP")
    ret = PutGateSetting("Test Run", 2, "DOWN")
    ret = PutGateSetting("Test Run", 30, "FREE")
    
    ' PutRunner()/PutRunners()
    Debug.Print "PutRunner()"
    ret = PutRunner("Test Run", 0, "bib001", "K1", "")  ' Row >= 0
    ret = PutRunner("Test Run", 1, "bib002", "K1", "")
    ret = PutRunner("Test Run", 2, "bib003", "K1", "")
    ret = PutRunner("Test Run", 3, "bib004", "K1", "")
    
    
    ' getRecords()
    Debug.Print "getRecords()"
    Set ret = GetRecords("Test Run")
    Debug.Print ret.Count
    For Each v In ret
        With v
            ' runner
            s = "runner: {"
            With .Item("runner")
                s = s & "row: " & .Item("row") & ", "
                s = s & "bib: """ & .Item("bib") & """, "
                s = s & "tag: """ & .Item("tag") & """, "
                s = s & "locked: """ & .Item("locked") & """, "
            End With
            s = s & "}"
            Debug.Print s
            
            ' started
            s = "started: {"
            With .Item("started")
                s = s & "judge: """ & .Item("judge") & """, "
                s = s & "time: """ & .Item("time") & """, "
            End With
            s = s & "}"
            Debug.Print s
            
            ' finished
            s = "finished: {"
            With .Item("finished")
                s = s & "judge: """ & .Item("judge") & """, "
                s = s & "time: """ & .Item("time") & """, "
            End With
            s = s & "}"
            Debug.Print s
            
            ' gates
            s = "gates: {"
            For Each g In .Item("gates")
                With g
                    s = s & "num: " & .Item("num") & ", "
                    s = s & "judge: """ & .Item("judge") & """, "
                End With
            Next g
            s = s & "}"
            Debug.Print s
            Debug.Print "---"
            
        End With
    Next v
    
    ' getRecord()
    Debug.Print "GetRecord()"
    Set ret = GetRecord("Test Run", 1)
    With ret
        ' runner
        s = "runner: {"
        With .Item("runner")
            s = s & "row: " & .Item("row") & ", "
            s = s & "bib: """ & .Item("bib") & """, "
            s = s & "tag: """ & .Item("tag") & """, "
            s = s & "locked: """ & .Item("locked") & """, "
        End With
        s = s & "}"
        Debug.Print s
        
        ' started
        s = "started: {"
        With .Item("started")
            s = s & "judge: """ & .Item("judge") & """, "
            s = s & "time: """ & .Item("time") & """, "
        End With
        s = s & "}"
        Debug.Print s
        
        ' finished
        s = "finished: {"
        With .Item("finished")
            s = s & "judge: """ & .Item("judge") & """, "
            s = s & "time: """ & .Item("time") & """, "
        End With
        s = s & "}"
        Debug.Print s
        
        ' gates
        s = "gates: {"
        For Each g In .Item("gates")
            With g
                s = s & "num: " & .Item("num") & ", "
                s = s & "judge: """ & .Item("judge") & """, "
            End With
        Next g
        s = s & "}"
        Debug.Print s
        Debug.Print "---"
        
    End With
    
End Sub

Private Function GetConfigFilename(Optional encodingTag As String = "utf16le") As String
    GetConfigFilename = ThisWorkbook.FullName & IIf(encodingTag = "", "", "." & encodingTag) & ".json"
End Function

Private Function LoadConfig() As Boolean
On Error GoTo ErrHandler
    Dim s As String
    Dim dJSON As Dictionary
    With CreateObject("Scripting.FileSystemObject").GetFile(GetConfigFilename()).OpenAsTextStream(1, -1)    ' unicode
        s = .ReadAll()
        .Close
    End With
    Set dJSON = JsonConverter.ParseJson(s)
    With dJSON.Item("app").Item("TestGasSlalom")
        mUrl = .Item("WebAppURL")
    End With
    LoadConfig = True
ExitProc:
    Exit Function
ErrHandler:
    Resume ExitProc
End Function

Private Sub SaveConfig()
    Dim s As String
    Dim dJSON As Dictionary
    Dim dApp As Dictionary
    Dim dTestGasSlalom As Dictionary
    
    Set dJSON = New Dictionary
    Set dApp = New Dictionary
    Set dTestGasSlalom = New Dictionary
    
    dJSON.Add "app", dApp
    dApp.Add "TestGasSlalom", dTestGasSlalom
    dTestGasSlalom.Add "WebAppURL", mUrl
    
    s = JsonConverter.ConvertToJson(dJSON, 4)
    With CreateObject("Scripting.FileSystemObject").CreateTextFile(GetConfigFilename(), True, True)    ' unicode
        .WriteLine s
        .Close
    End With
    
End Sub

Public Sub ResetWebAppUrl()
    InitWebAppUrl
End Sub

Private Function InitWebAppUrl() As Boolean
On Error GoTo ErrHandler
    Dim s As String
    
    ' API URL
    s = InputBox("[ Google Apps Script ]" & vbCrLf & "Your Web App URL:", "InitWebAppUrl()", mUrl)
    If s = "" Then
        Debug.Print "canceled."
        GoTo ExitProc
    End If
    
    mUrl = s
    SaveConfig
    InitWebAppUrl = True
    
ExitProc:
    Exit Function
ErrHandler:
    MsgBox Err.Description, vbCritical, "Err (" & Err.Number & ")"
    Resume ExitProc
End Function

Public Sub NewHeatSheet()
On Error GoTo ErrHandler
    Dim s As String
    Dim sht As Worksheet
    
    s = InputBox("New Heat Name", "NewHeatSheet()")
    If s = "" Then
        MsgBox "Canceled", vbExclamation, "NewHeatSheet()"
        GoTo ExitProc
    End If
    
    If Not ExistsRecordSheet(s) Then
        Set sht = CreateNewRecordSheet(s)
    End If
    
    Set sht = CreateNewRunnerSheet(s)
    sht.Parent.Activate
    sht.Activate
    
ExitProc:
    Exit Sub
ErrHandler:
    MsgBox Err.Description, vbCritical, "Err (" & Err.Number & ")"
    Resume ExitProc
End Sub

Public Sub AppSync()
On Error GoTo ErrHandler
    Dim aHeatName As String
    Dim sht As Worksheet
    Dim cList As Collection
    
    LoadConfig
    
    If mUrl = "" Then
        If Not InitWebAppUrl() Then
            GoTo ExitProc
        End If
    End If
    
    InitUrl mUrl
    
    Set sht = ThisWorkbook.ActiveSheet
    
    aHeatName = ParseRunnerSheetName(sht.Name)
    If aHeatName <> "" Then
        ' UPLOAD RUNNERS
        Set cList = GetRunners(aHeatName)
        If vbOK <> MsgBox("Upload runners.�@[count: " & cList.Count & " ]" & vbCrLf & sht.Name, vbQuestion + vbOKCancel, "UPLOAD RUNNERS") Then
            MsgBox "Canceled", vbExclamation, "AppSync()"
            GoTo ExitProc
        End If
        If Not ExistsHeat(aHeatName) Then
            AddHeat aHeatName
        End If
        PutRunners aHeatName, cList
        MsgBox "Uploaded.", vbInformation, "UPLOAD RUNNERS"
        GoTo ExitProc
    End If
    
    aHeatName = ParseRecordSheetName(sht.Name)
    If aHeatName <> "" Then
        ' DOWNLOAD RECORDS
        If vbOK <> MsgBox("Download records" & vbCrLf & sht.Name, vbQuestion + vbOKCancel, "DOWNLOAD RECORDS") Then
            MsgBox "Canceled", vbExclamation, "AppSync()"
            GoTo ExitProc
        End If
        Set cList = GetRecords(aHeatName)
        PutRecords aHeatName, cList
        MsgBox "Downloaded.", vbInformation, "DOWNLOAD RECORDS"
        GoTo ExitProc
    End If
    
    MsgBox "Select 'RUNNER' or 'RECORD' sheet.", vbExclamation, "AppSync()"
    
ExitProc:
    Exit Sub
ErrHandler:
    MsgBox Err.Description, vbCritical, "Err (" & Err.Number & ")"
    Resume ExitProc
End Sub
