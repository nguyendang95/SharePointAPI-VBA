VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DocumentSetVersionResource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_AccessToken As String

Friend Property Get AccessToken() As String
   AccessToken = m_AccessToken
End Property

Friend Property Let AccessToken(Value As String)
   m_AccessToken = Value
End Property

Public Function List(ByVal SiteId As String, ByVal ListId As String, ByVal ItemId As String, Optional ByVal ODataQuery As String) As Collection
    Dim strUrl As String
    strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/lists/" & ListId & "/items/" & ItemId & "/documentSetVersions"
    If ODataQuery <> vbNullString Then strUrl = strUrl & "?" & ODataQuery
    Dim objWinHttp As WinHttp.WinHttpRequest
    Set objWinHttp = New WinHttp.WinHttpRequest
    With objWinHttp
        .Open "GET", strUrl, True
        .SetRequestHeader "Accept", "application/json"
        .SetRequestHeader "Authorization", "Bearer " & AccessToken
        .Send
        .WaitForResponse
        Dim colDocumentSetVersions As New Collection
        If .Status = 200 Then
            Dim objJson As Scripting.Dictionary
            Set objJson = JsonConverter.ParseJson(.ResponseText)
            Dim objCreatedBy As IdentitySet
            Dim objLastModifiedBy As IdentitySet
            Dim objIdentity As Identity
            Dim i As Long
            Dim objDocumentSetVersion As DocumentSetVersion
            Dim objItemReference As ItemReference
            If objJson.Item("value").Count > 0 Then
                For i = 1 To objJson.Item("value").Count
                    Set objDocumentSetVersion = New DocumentSetVersion
                    With objDocumentSetVersion
                        .Id = objJson.Item("value")(i)("id")
                        .CreatedDateTime = objJson.Item("value")(i)("createdDateTime")
                        .Comment = objJson.Item("value")(i)("comment")
                        .ShouldCaptureMinorVersion = CBool(objJson.Item("value")(i)("shouldCaptureMinorVersion"))
                        Set objCreatedBy = New IdentitySet
                        Set objIdentity = New Identity
                        If objJson.Item("value")(i).Exists("createdBy") Then
                            If objJson.Item("value")(i)("createdBy").Exists("user") And IsEmpty(objJson.Item("value")(i)("createdBy")("user")) = False Then
                                objIdentity.DisplayName = objJson.Item("value")(i)("createdBy")("user")("displayName")
                                objIdentity.Id = objJson.Item("value")(i)("createdBy")("user")("id")
                                Set objCreatedBy.User = objIdentity
                            End If
                            If objJson.Item("value")(i)("createdBy").Exists("application") And IsEmpty(objJson.Item("value")(i)("createdBy")("application")) = False Then
                                objIdentity.DisplayName = objJson.Item("value")(i)("createdBy")("application")("displayName")
                                objIdentity.Id = objJson.Item("value")(i)("createdBy")("application")("id")
                                Set objCreatedBy.Application = objIdentity
                            End If
                            If objJson.Item("value")(i)("createdBy").Exists("device") And IsEmpty(objJson.Item("value")(i)("createdBy")("device")) = False Then
                                objIdentity.DisplayName = objJson.Item("value")(i)("createdBy")("device")("displayName")
                                objIdentity.Id = objJson.Item("value")(i)("createdBy")("device")("id")
                                Set objCreatedBy.Device = objIdentity
                            End If
                        End If
                        Set .CreatedBy = objCreatedBy
                        .LastModifiedDateTime = objJson.Item("value")(i)("lastModifiedDateTime")
                        If objJson.Item("value")(i).Exists("lastModifiedBy") Then
                            Set objLastModifiedBy = New IdentitySet
                            Set objIdentity = New Identity
                            If objJson.Item("value")(i)("lastModifiedBy").Exists("user") And IsEmpty(objJson.Item("value")(i)("lastModifiedBy")("user")) = False Then
                                objIdentity.DisplayName = objJson.Item("value")(i)("lastModifiedBy")("user")("displayName")
                                objIdentity.Id = objJson.Item("value")(i)("lastModifiedBy")("user")("id")
                                Set objLastModifiedBy.User = objIdentity
                            End If
                            If objJson.Item("value")(i)("lastModifiedBy").Exists("application") And IsEmpty(objJson.Item("value")(i)("lastModifiedBy")("application")) = False Then
                                objIdentity.DisplayName = objJson.Item("value")(i)("lastModifiedBy")("application")("displayName")
                                objIdentity.Id = objJson.Item("value")(i)("lastModifiedBy")("application")("id")
                                Set objLastModifiedBy.Application = objIdentity
                            End If
                            If objJson.Item("value")(i)("lastModifiedBy").Exists("device") And IsEmpty(objJson.Item("value")(i)("lastModifiedBy")("device")) = False Then
                                objIdentity.DisplayName = objJson.Item("value")(i)("lastModifiedBy")("device")("displayName")
                                objIdentity.Id = objJson.Item("value")(i)("lastModifiedBy")("device")("id")
                                Set objLastModifiedBy.Device = objIdentity
                            End If
                            Set .LastModifiedBy = objLastModifiedBy
                        End If
                        colDocumentSetVersions.Add objDocumentSetVersion
                    End With
                Next
            End If
            Set List = colDocumentSetVersions
        Else: Call HandleError(.ResponseText)
        End If
    End With
End Function

Public Function GetDocumentSetVersion(ByVal SiteId As String, ByVal ListId As String, ByVal ItemId As String, Optional ByVal ODataQuery As String) As DocumentSetVersion
    Dim strUrl As String
    strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/lists/" & ListId & "/items/" & ItemId & "/documentSetVersions"
    If ODataQuery <> vbNullString Then strUrl = strUrl & "?" & ODataQuery
    Dim objWinHttp As WinHttp.WinHttpRequest
    Set objWinHttp = New WinHttp.WinHttpRequest
    With objWinHttp
        .Open "GET", strUrl, True
        .SetRequestHeader "Accept", "application/json"
        .SetRequestHeader "Authorization", "Bearer " & AccessToken
        .Send
        .WaitForResponse
        If .Status = 200 Then
            Dim objJson As Scripting.Dictionary
            Set objJson = JsonConverter.ParseJson(.ResponseText)
            Dim objCreatedBy As IdentitySet
            Dim objLastModifiedBy As IdentitySet
            Dim objIdentity As Identity
            Dim objDocumentSetVersion As DocumentSetVersion
            Dim objItemReference As ItemReference
            Set objDocumentSetVersion = New DocumentSetVersion
            With objDocumentSetVersion
                .Id = objJson.Item("id")
                .CreatedDateTime = objJson.Item("createdDateTime")
                .Comment = objJson.Item("comment")
                .ShouldCaptureMinorVersion = CBool(objJson.Item("shouldCaptureMinorVersion"))
                Set objCreatedBy = New IdentitySet
                Set objIdentity = New Identity
                If objJson.Item("createdBy").Exists("user") And IsEmpty(objJson.Item("createdBy")("user")) = False Then
                    objIdentity.DisplayName = objJson.Item("createdBy")("user")("displayName")
                    objIdentity.Id = objJson.Item("createdBy")("user")("id")
                    Set objCreatedBy.User = objIdentity
                End If
                If objJson.Item("createdBy").Exists("application") And IsEmpty(objJson.Item("createdBy")("application")) = False Then
                    objIdentity.DisplayName = objJson.Item("createdBy")("application")("displayName")
                    objIdentity.Id = objJson.Item("createdBy")("application")("id")
                    Set objCreatedBy.Application = objIdentity
                End If
                If objJson.Item("createdBy").Exists("device") And IsEmpty(objJson.Item("createdBy")("device")) = False Then
                    objIdentity.DisplayName = objJson.Item("createdBy")("device")("displayName")
                    objIdentity.Id = objJson.Item("createdBy")("device")("id")
                    Set objCreatedBy.Device = objIdentity
                End If
                Set .CreatedBy = objCreatedBy
                .LastModifiedDateTime = objJson.Item("lastModifiedDateTime")
                If objJson.Exists("lastModifiedBy") Then
                    Set objLastModifiedBy = New IdentitySet
                    Set objIdentity = New Identity
                    If objJson.Item("lastModifiedBy").Exists("user") And IsEmpty(objJson.Item("lastModifiedBy")("user")) = False Then
                        objIdentity.DisplayName = objJson.Item("lastModifiedBy")("user")("displayName")
                        objIdentity.Id = objJson.Item("lastModifiedBy")("user")("id")
                        Set objLastModifiedBy.User = objIdentity
                    End If
                    If objJson.Item("lastModifiedBy").Exists("application") And IsEmpty(objJson.Item("lastModifiedBy")("application")) = False Then
                        objIdentity.DisplayName = objJson.Item("lastModifiedBy")("application")("displayName")
                        objIdentity.Id = objJson.Item("lastModifiedBy")("application")("id")
                        Set objLastModifiedBy.Application = objIdentity
                    End If
                    If objJson.Item("lastModifiedBy").Exists("device") And IsEmpty(objJson.Item("lastModifiedBy")("device")) = False Then
                        objIdentity.DisplayName = objJson.Item("lastModifiedBy")("device")("displayName")
                        objIdentity.Id = objJson.Item("lastModifiedBy")("device")("id")
                        Set objLastModifiedBy.Device = objIdentity
                    End If
                    Set .LastModifiedBy = objLastModifiedBy
                End If
                Set GetDocumentSetVersion = objDocumentSetVersion
            End With
        Else: Call HandleError(.ResponseText)
        End If
    End With
End Function

Public Sub Delete(ByVal SiteId As String, ByVal ListId As String, ByVal ItemId As String, ByVal DocumentSetVersionId As String)
    Dim strUrl As String
    strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/lists/" & ListId & "/items/" & ItemId & "/documentSetVersions/" & DocumentSetVersionId
    Dim objWinHttp As WinHttp.WinHttpRequest
    Set objWinHttp = New WinHttp.WinHttpRequest
    With objWinHttp
        .Open "DELETE", strUrl, True
        .SetRequestHeader "Accept", "application/json"
        .SetRequestHeader "Authorization", "Bearer " & AccessToken
        .Send
        .WaitForResponse
        If .Status <> 204 Then Call HandleError(.ResponseText)
    End With
End Sub

Public Sub Restore(ByVal SiteId As String, ByVal ListId As String, ByVal ItemId As String, ByVal DocumentSetVersionId As String)
    Dim strUrl As String
    strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/lists/" & ListId & "/items/" & ItemId & "/documentSetVersions/" & DocumentSetVersionId & "/restore"
    Dim objWinHttp As WinHttp.WinHttpRequest
    Set objWinHttp = New WinHttp.WinHttpRequest
    With objWinHttp
        .Open "POST", strUrl, True
        .SetRequestHeader "Accept", "application/json"
        .SetRequestHeader "Authorization", "Bearer " & AccessToken
        .Send
        .WaitForResponse
        If .Status <> 204 Then Call HandleError(.ResponseText)
    End With
End Sub

Public Function Create(ByVal SiteId As String, ByVal ListId As String, ByVal ItemId As String, ByVal DocumentSetVersion As DocumentSetVersion) As DocumentSetVersion
    Dim strUrl As String
    strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/lists/" & ListId & "/items/" & ItemId & "/documentSetVersions"
    Dim objWinHttp As WinHttp.WinHttpRequest
    Set objWinHttp = New WinHttp.WinHttpRequest
    With objWinHttp
        .Open "POST", strUrl, True
        .SetRequestHeader "Accept", "application/json"
        .SetRequestHeader "Authorization", "Bearer " & AccessToken
        .Send JsonConverter.ConvertToJson(DocumentSetVersion.ToJson)
        .WaitForResponse
        If .Status = 201 Then
            Dim objJson As Scripting.Dictionary
            Set objJson = JsonConverter.ParseJson(.ResponseText)
            Dim objCreatedBy As IdentitySet
            Dim objLastModifiedBy As IdentitySet
            Dim objIdentity As Identity
            Dim objDocumentSetVersion As DocumentSetVersion
            Dim objItemReference As ItemReference
            Set objDocumentSetVersion = New DocumentSetVersion
            With objDocumentSetVersion
                .Id = objJson.Item("id")
                .CreatedDateTime = objJson.Item("createdDateTime")
                .Comment = objJson.Item("comment")
                .ShouldCaptureMinorVersion = CBool(objJson.Item("shouldCaptureMinorVersion"))
                Set objCreatedBy = New IdentitySet
                Set objIdentity = New Identity
                If objJson.Item("createdBy").Exists("user") And IsEmpty(objJson.Item("createdBy")("user")) = False Then
                    objIdentity.DisplayName = objJson.Item("createdBy")("user")("displayName")
                    objIdentity.Id = objJson.Item("createdBy")("user")("id")
                    Set objCreatedBy.User = objIdentity
                End If
                If objJson.Item("createdBy").Exists("application") And IsEmpty(objJson.Item("createdBy")("application")) = False Then
                    objIdentity.DisplayName = objJson.Item("createdBy")("application")("displayName")
                    objIdentity.Id = objJson.Item("createdBy")("application")("id")
                    Set objCreatedBy.Application = objIdentity
                End If
                If objJson.Item("createdBy").Exists("device") And IsEmpty(objJson.Item("createdBy")("device")) = False Then
                    objIdentity.DisplayName = objJson.Item("createdBy")("device")("displayName")
                    objIdentity.Id = objJson.Item("createdBy")("device")("id")
                    Set objCreatedBy.Device = objIdentity
                End If
                Set .CreatedBy = objCreatedBy
                .LastModifiedDateTime = objJson.Item("lastModifiedDateTime")
                If objJson.Exists("lastModifiedBy") Then
                    Set objLastModifiedBy = New IdentitySet
                    Set objIdentity = New Identity
                    If objJson.Item("lastModifiedBy").Exists("user") And IsEmpty(objJson.Item("lastModifiedBy")("user")) = False Then
                        objIdentity.DisplayName = objJson.Item("lastModifiedBy")("user")("displayName")
                        objIdentity.Id = objJson.Item("lastModifiedBy")("user")("id")
                        Set objLastModifiedBy.User = objIdentity
                    End If
                    If objJson.Item("lastModifiedBy").Exists("application") And IsEmpty(objJson.Item("lastModifiedBy")("application")) = False Then
                        objIdentity.DisplayName = objJson.Item("lastModifiedBy")("application")("displayName")
                        objIdentity.Id = objJson.Item("lastModifiedBy")("application")("id")
                        Set objLastModifiedBy.Application = objIdentity
                    End If
                    If objJson.Item("lastModifiedBy").Exists("device") And IsEmpty(objJson.Item("lastModifiedBy")("device")) = False Then
                        objIdentity.DisplayName = objJson.Item("lastModifiedBy")("device")("displayName")
                        objIdentity.Id = objJson.Item("lastModifiedBy")("device")("id")
                        Set objLastModifiedBy.Device = objIdentity
                    End If
                    Set .LastModifiedBy = objLastModifiedBy
                End If
                Set Create = objDocumentSetVersion
            End With
        Else: Call HandleError(.ResponseText)
        End If
    End With
End Function

Private Sub HandleError(ByVal Source As String)
    Dim objJson As Scripting.Dictionary
    Set objJson = JsonConverter.ParseJson(Source)
    Err.Raise vbObjectError, objJson.Item("error")("code"), objJson.Item("error")("message")
End Sub
