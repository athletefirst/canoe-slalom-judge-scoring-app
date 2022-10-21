Attribute VB_Name = "GasSlalom"
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

Public Const CGATE_UP As String = "UP"
Public Const CGATE_DOWN As String = "DOWN"
Public Const CGATE_FREE As String = "FREE"

Public Const CJUDGE_STARTED As String = "STARTED"
Public Const CJUDGE_FINISHED As String = "FINISHED"
Public Const CJUDGE_FINISHED_50 As String = "FINISHED_50"
Public Const CJUDGE_DSQ As String = "DSQ"
Public Const CJUDGE_DNS As String = "DNS"
Public Const CJUDGE_DNF As String = "DNF"

Public Const CERR_STATUSCODE   As Long = 2001 + vbObjectError   ' web response error
Public Const CERR_RESPONSE     As Long = 2002 + vbObjectError   ' rpc response error
Public Const CERR_CONTENT      As Long = 2003 + vbObjectError   ' content error - parse error, JSON
Public Const CERR_RESULT_OK    As Long = 2004 + vbObjectError   ' result message error

Private Const CRESULT_STR_OK As String = "ok"

Private Const CHTMLBODYDIV_CLOSING As String = "</div></body></html>"
Private Const CTEXT_EXCEPTION As String = ">Exception"

Private mWebClient As New WebClient
Private mUrl As String

Private mLastErrNum As Long
Private mLastErrDsc As String
Private mLastErrSrc As String

Private Sub SetLastError(Optional aErr As ErrObject = Nothing)
    mLastErrNum = 0
    mLastErrDsc = "(no error)"
    mLastErrSrc = "(no error)"
    If Not aErr Is Nothing Then
        If aErr.Number <> 0 Then
            mLastErrNum = Err.Number
            mLastErrDsc = Err.Description
            mLastErrSrc = Err.Source
        End If
    End If
End Sub

Public Function GetLastErrNumber() As Long
    GetLastErrNumber = mLastErrNum
End Function

Public Function GetLastErrDescription() As String
    GetLastErrDescription = mLastErrDsc
End Function

Public Function GetLastErrSource() As String
    GetLastErrSource = mLastErrSrc
End Function

Private Function ParseJsonContent(aContent As String, aOperationId As String) As Dictionary
On Error GoTo ErrHandler
    Dim s As String
    Dim n As Long
    Dim n2 As Long
    Dim errSrc As String
    Dim errDsc As String
    
    Set ParseJsonContent = JsonConverter.ParseJson(aContent)
ExitProc:
    Exit Function
ErrHandler:
    errDsc = "content error ( " & aOperationId & " )" & vbCrLf & aContent
    If LCase(Right(aContent, Len(CHTMLBODYDIV_CLOSING))) = LCase(CHTMLBODYDIV_CLOSING) Then
        n = InStrRev(aContent, CTEXT_EXCEPTION, -1, vbTextCompare)
        If n > 0 Then
            n2 = Len(aContent) - n - Len(CHTMLBODYDIV_CLOSING)
            errDsc = "content error ( " & aOperationId & " )" & vbCrLf & Mid(aContent, n + 1, n2)
        End If
    End If
    errSrc = aOperationId
    Err.Raise CERR_CONTENT, errSrc, errDsc
End Function

Private Function ExecuteRPC(aOperationId As String, aOperationData As Dictionary, Optional aSilentError As Boolean = False, Optional aURL As String = "") As Variant
On Error GoTo ErrHandler
    Dim wr As WebResponse
    Dim sUrl As String
    Dim dBody As Dictionary
    Dim errSrc As String
    Dim errDsc As String
    Dim sContent As String
    Dim ret As Dictionary
    
    If aURL = "" Then
        sUrl = mUrl
    Else
        sUrl = aURL
    End If

    Set dBody = New Dictionary
    dBody.Add "operationId", aOperationId
    dBody.Add "operationData", aOperationData
    Set wr = mWebClient.PostJson(sUrl, dBody)
    
    If wr.StatusCode <> WebStatusCode.Ok Then
        errSrc = aOperationId
        errDsc = "Web Response Error (status code: " & wr.StatusCode & " )" & vbCrLf & aOperationId
        Err.Raise CERR_STATUSCODE, errSrc, errDsc
    End If
    
    sContent = wr.Content
    Set ret = ParseJsonContent(sContent, aOperationId)
    If Not ret.Exists("result") Then
        errSrc = aOperationId
        errDsc = "Missing result." & vbCrLf & aOperationId
        Err.Raise CERR_RESPONSE, errSrc, errDsc
    End If
    
    If IsObject(ret("result")) Then
        Set ExecuteRPC = ret("result")
    Else
        ExecuteRPC = ret("result")
    End If

ExitProc:
    Exit Function
    
ErrHandler:
    SetLastError Err
    If Not aSilentError Then
        Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
    End If
    Resume ExitProc
        
End Function

Private Function ExecuteRpcAndResultOk(aOperationId As String, aOperationData As Dictionary, Optional aSilentError As Boolean = False, Optional aURL As String = "") As Boolean
On Error GoTo ErrHandler
    Dim ret As String
    Dim errSrc As String
    Dim errDsc As String
    
    ret = ExecuteRPC(aOperationId, aOperationData, False, aURL) ' Handling Err by self.
    If ret = CRESULT_STR_OK Then
        ExecuteRpcAndResultOk = True
    Else
        ExecuteRpcAndResultOk = False
        errSrc = aOperationId
        errDsc = "Unknown result ( '" & ret & "' )" & vbCrLf & aOperationId
        Err.Raise CERR_RESULT_OK, errSrc, errDsc
    End If
    
ExitProc:
    Exit Function
    
ErrHandler:
    SetLastError Err
    If Not aSilentError Then
        Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
    End If
    Resume ExitProc
    
End Function

Public Sub InitUrl(aURL As String)
    SetLastError
    mUrl = aURL
End Sub

Public Function ExistsHeat(aHeatName As String, Optional aSilentError As Boolean = False, Optional aURL As String = "") As Boolean
On Error GoTo ErrHandler
    Dim c As Collection
    Dim v As Variant
    
    Set c = GetHeatsAll(False, aURL)
    For Each v In c
        If v("heatName") = aHeatName Then
            ExistsHeat = True
            GoTo ExitProc
        End If
    Next v
    
ExitProc:
    Exit Function
    
ErrHandler:
    SetLastError Err
    If Not aSilentError Then
        Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
    End If
    Resume ExitProc
End Function

Public Function AddHeat(aHeatName As String, Optional aSilentError As Boolean = False, Optional aURL As String = "") As Boolean
    SetLastError
    Dim d As New Dictionary
    d.Add "heatName", aHeatName
    AddHeat = ExecuteRpcAndResultOk("addHeat", d, aSilentError, aURL)
End Function

Public Function GetHeatsAll(Optional aSilentError As Boolean = False, Optional aURL As String = "") As Collection
    SetLastError
    Set GetHeatsAll = ExecuteRPC("getHeatsAll", New Dictionary, aSilentError, aURL)
End Function

Public Function PutGateSettings(aHeatName As String, aGateSettings As Collection, Optional aSilentError As Boolean = False, Optional aURL As String = "") As Boolean
    SetLastError
    Dim d As New Dictionary
    d.Add "heatName", aHeatName
    d.Add "gateSettings", aGateSettings
    PutGateSettings = ExecuteRpcAndResultOk("putGateSettings", d, aSilentError, aURL)
End Function

Public Function PutGateSetting(aHeatName As String, aNum As Long, aDirection As String, Optional aSilentError As Boolean = False, Optional aURL As String = "") As Boolean
    SetLastError
    Dim cList As New Collection
    Dim dItem As New Dictionary
    dItem.Add "num", aNum
    dItem.Add "direction", aDirection
    cList.Add dItem
    PutGateSetting = PutGateSettings(aHeatName, cList, aSilentError, aURL)
End Function

Public Function PutRunners(aHeatName As String, aRunners As Collection, Optional aSilentError As Boolean = False, Optional aURL As String = "") As Boolean
    SetLastError
    Dim d As New Dictionary
    d.Add "heatName", aHeatName
    d.Add "runners", aRunners
    PutRunners = ExecuteRpcAndResultOk("putRunners", d, aSilentError, aURL)
End Function

Public Function PutRunner(aHeatName As String, aRow As Long, aBib As String, aTag As String, aLocked As String, Optional aSilentError As Boolean = False, Optional aURL As String = "") As Boolean
    SetLastError
    Dim cList As New Collection
    Dim dItem As New Dictionary
    dItem.Add "row", aRow
    dItem.Add "bib", aBib
    dItem.Add "tag", aTag
    dItem.Add "locked", aLocked
    cList.Add dItem
    PutRunner = PutRunners(aHeatName, cList, aSilentError, aURL)
End Function

Public Function GetRecords(aHeatName As String, Optional aRow1 As Long = 0, Optional aRow2 As Long = -1, Optional aSilentError As Boolean = False, Optional aURL As String = "") As Collection
    SetLastError
    Dim d As New Dictionary
    d.Add "heatName", aHeatName
    If aRow1 <> 0 Then
        d.Add "row1", aRow1
    End If
    If aRow2 <> -1 Then
        d.Add "row2", aRow2
    End If
    Set GetRecords = ExecuteRPC("getRecords", d, aSilentError, aURL)
End Function

Public Function GetRecord(aHeatName As String, aRow As Long, Optional aSilentError As Boolean = False, Optional aURL As String = "") As Dictionary
On Error GoTo ErrHandler
    Dim ret As Collection
    Dim errSrc As String
    Dim errDsc As String
    
    Set ret = GetRecords(aHeatName, aRow, aRow, False, aURL)
    If ret.Count <> 1 Then
        errSrc = "GetRecord"
        errDsc = "missing record (row: " & aRow & " )" & vbCrLf & "GetRecord"
        Err.Raise CERR_RESULT_OK, errSrc, errDsc
    End If
    
    Set GetRecord = ret(1)
    
ExitProc:
    Exit Function
    
ErrHandler:
    SetLastError Err
    If Not aSilentError Then
        Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
    End If
    Resume ExitProc
End Function
