VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OtherResource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum RichLongRunningOperationLocation
    Site
    List
End Enum

Private m_AccessToken As String

Friend Property Get AccessToken() As String
   AccessToken = m_AccessToken
End Property

Friend Property Let AccessToken(Value As String)
   m_AccessToken = Value
End Property

Public Function GetRichLongRunningOperationOperation(ByVal SiteId As String, Location As RichLongRunningOperationLocation, Optional ByVal ListId As String, Optional ByVal ODataQuery As String = vbNullString) As RichLongRunningOperation
    Dim strUrl As String
    Select Case Location
        Case List
            strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/lists/" & ListId & "/operations"
        Case Site
            strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/operations"
    End Select
    If ODataQuery <> vbNullString Then
        strUrl = strUrl & "?" & ODataQuery
    End If
    Dim objWinHttp As WinHttp.WinHttpRequest
    Set objWinHttp = New WinHttp.WinHttpRequest
    With objWinHttp
        .Open "GET", strUrl, True
        .SetRequestHeader "Authorization", "Bearer " & AccessToken
        .SetRequestHeader "Accept", "application/json"
        .Send
        .WaitForResponse
        If .Status = 200 Then
            Dim objJson As Scripting.Dictionary
            Set objJson = JsonConverter.ParseJson(.ResponseText)
            Dim objRichLongRunningOperation As RichLongRunningOperation
            Set objRichLongRunningOperation = New RichLongRunningOperation
            With objRichLongRunningOperation
                .Id = objJson.Item("id")
                .CreatedDateTime = objJson.Item("createdDateTime")
                .ResourceId = objJson.Item("resourceId")
                .ResourceLocation = objJson.Item("resourceLocation")
                .Status = objJson.Item("status")
                .RichLongRunningOperationType = objJson.Item("type")
            End With
        Set GetRichLongRunningOperationOperation = objRichLongRunningOperation
        Else: Call HandleError(.ResponseText)
        End If
    End With
End Function

Private Sub HandleError(ByVal Source As String)
    Dim objJson As Scripting.Dictionary
    Set objJson = JsonConverter.ParseJson(Source)
    Err.Raise vbObjectError, objJson.Item("error")("code"), objJson.Item("error")("message")
End Sub
