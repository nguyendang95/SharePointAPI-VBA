VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BatchRequest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Requests As New Collection

Public Property Get Requests() As Collection
    Set Requests = m_Requests
End Property

Public Property Set Requests(Value As Collection)
    Set m_Requests = Value
End Property

Public Function ToJson() As Scripting.Dictionary
    Dim objJson As Scripting.Dictionary
    Set objJson = New Scripting.Dictionary
    With objJson
        If Requests.Count > 0 Then .Add "requests", Requests
    End With
    Set ToJson = objJson
End Function

Public Function Execute() As BatchRequestResponse
    Dim objWinHttp As WinHttp.WinHttpRequest
    Set objWinHttp = New WinHttp.WinHttpRequest
    With objWinHttp
        .Open "POST", "https://graph.microsoft.com/v1.0/$batch", True
        .SetRequestHeader "Accept", "application/json"
        .SetRequestHeader "Content-Type", "application/json"
        .Send JsonConverter.ConvertToJson(ToJson)
        .WaitForResponse
        If .Status = 200 Then
            Dim objJson As Scripting.Dictionary
            Set objJson = JsonConverter.ParseJson(.ResponseText)
            Dim objBatchRequestResponse As BatchRequestResponse
            Set objBatchRequestResponse = New BatchRequestResponse
            Dim i As Long
            If objJson.Item("responses").Count > 0 Then
                For i = 1 To objJson.Item("responses").Count
                    objBatchRequestResponse.Responses.Add FetchResponse(objJson.Item("responses")(i))
                Next
                Set Execute = objBatchRequestResponse
            End If
        Else: Call HandleError(.ResponseText)
        End If
    End With
End Function

Private Function FetchResponse(ByVal ResponseText As Variant) As Response
    Dim objJson As Scripting.Dictionary
    If TypeName(ResponseText) = "String" Then
        Set objJson = JsonConverter.ParseJson(ResponseText)
    ElseIf TypeName(ResponseText) = "Dictionary" Then
        Set objJson = ResponseText
    Else: Exit Function
    End If
    Dim objResponse As Response
    Set objResponse = New Response
    With objResponse
        .Id = objJson.Item("id")
        .Status = objJson.Item("status")
        If objJson.Exists("headers") Then Set .Headers = objJson.Item("headers")
        If objJson.Exists("body") Then Set .Body = objJson.Item("body")
    End With
    Set FetchResponse = objResponse
End Function

Private Sub HandleError(ByVal Source As String)
    Dim objJson As Scripting.Dictionary
    Set objJson = JsonConverter.ParseJson(Source)
    Err.Raise vbObjectError, objJson.Item("error")("code"), objJson.Item("error")("message")
End Sub
