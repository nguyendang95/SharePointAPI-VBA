VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ContentTypeResource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum ContentTypeLocation
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

Public Function GetContentType(ByVal SiteId As String, Location As ContentTypeLocation, ByVal ContentTypeId As String, Optional ByVal ListId As String, Optional ByVal ODataQuery As String) As ContentType
    Dim strUrl As String
    Select Case Location
        Case List
            strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/lists/" & ListId & "/contentTypes/" & ContentTypeId
        Case Site
            strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/contentTypes/" & ContentTypeId
    End Select
    If ODataQuery <> vbNullString Then strUrl = strUrl & "?" & ODataQuery
    Dim objWinHttp As WinHttp.WinHttpRequest
    Set objWinHttp = New WinHttp.WinHttpRequest
    With objWinHttp
        .Open "POST", strUrl, True
        .SetRequestHeader "Authorization", "Bearer " & AccessToken
        .SetRequestHeader "Accept", "application/json"
        .Send
        .WaitForResponse
        If .Status = 200 Then
            Dim objJson As Scripting.Dictionary
            Set objJson = JsonConverter.ParseJson(.ResponseText)
            Dim objContentType As ContentType
            Set objContentType = New ContentType
            With objContentType
                .Id = objJson.Item("id")
                .IsBuiltIn = CBool(objJson.Item("isBuiltIn"))
                .Description = objJson.Item("description")
                .Group = objJson.Item("group")
                .Hidden = CBool(objJson.Item("hidden"))
                .Name = objJson.Item("name")
                .ParentId = objJson.Item("parentId")
                .ReadOnly = CBool(objJson.Item("readOnly"))
                .Sealed = CBool(objJson.Item("sealed"))
                Dim objBase As ContentType
                Set objBase = New ContentType
                With objBase
                    .Id = objJson.Item("base")("id")
                    .IsBuiltIn = CBool(objJson.Item("base")("isBuiltIn"))
                    .Description = objJson.Item("base")("description")
                    .Group = objJson.Item("base")("group")
                    .Hidden = CBool(objJson.Item("base")("hidden"))
                    .Name = objJson.Item("base")("name")
                    .ParentId = objJson.Item("base")("parentId")
                    .ReadOnly = CBool(objJson.Item("base")("readOnly"))
                    .Sealed = CBool(objJson.Item("base")("sealed"))
                End With
                Set .Base = objBase
            End With
            Set GetContentType = objContentType
        Else: Call HandleError(.ResponseText)
        End If
    End With
End Function

Public Function Update(ByVal SiteId As String, Location As ContentTypeLocation, ByVal ContentTypeId As String, ByVal ContentType As ContentType, Optional ByVal ListId As String) As ContentType
    Dim strUrl As String
    Select Case Location
        Case List
            strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/lists/" & ListId & "/contentTypes/" & ContentTypeId
        Case Site
            strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/contentTypes/" & ContentTypeId
    End Select
    Dim objWinHttp As WinHttp.WinHttpRequest
    Set objWinHttp = New WinHttp.WinHttpRequest
    With objWinHttp
        .Open "PATCH", strUrl, True
        .SetRequestHeader "Authorization", "Bearer " & AccessToken
        .SetRequestHeader "Accept", "application/json"
        .SetRequestHeader "Content-Type", "application/json"
        .Send JsonConverter.ConvertToJson(ContentType.ToJson)
        .WaitForResponse
        If .Status = 200 Then
            Dim objJson As Scripting.Dictionary
            Set objJson = JsonConverter.ParseJson(.ResponseText)
            Dim objContentType As ContentType
            Set objContentType = New ContentType
            With objContentType
                .Id = objJson.Item("id")
                .IsBuiltIn = CBool(objJson.Item("isBuiltIn"))
                .Description = objJson.Item("description")
                .Group = objJson.Item("group")
                .Hidden = CBool(objJson.Item("hidden"))
                .Name = objJson.Item("name")
                .ParentId = objJson.Item("parentId")
                .ReadOnly = CBool(objJson.Item("readOnly"))
                .Sealed = CBool(objJson.Item("sealed"))
                Dim objBase As ContentType
                Set objBase = New ContentType
                With objBase
                    .Id = objJson.Item("base")("id")
                    .IsBuiltIn = CBool(objJson.Item("base")("isBuiltIn"))
                    .Description = objJson.Item("base")("description")
                    .Group = objJson.Item("base")("group")
                    .Hidden = CBool(objJson.Item("base")("hidden"))
                    .Name = objJson.Item("base")("name")
                    .ParentId = objJson.Item("base")("parentId")
                    .ReadOnly = CBool(objJson.Item("base")("readOnly"))
                    .Sealed = CBool(objJson.Item("base")("sealed"))
                End With
                Set .Base = objBase
            End With
            Set Update = objContentType
        Else: Call HandleError(.ResponseText)
        End If
    End With
End Function

Public Sub Delete(ByVal SiteId As String, Location As ContentTypeLocation, ByVal ContentTypeId As String, Optional ByVal ListId As String)
    Dim strUrl As String
    Select Case Location
        Case List
            strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/lists/" & ListId & "/contentTypes/" & ContentTypeId
        Case Site
            strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/contentTypes/" & ContentTypeId
    End Select
    Dim objWinHttp As WinHttp.WinHttpRequest
    Set objWinHttp = New WinHttp.WinHttpRequest
    With objWinHttp
        .Open "DELETE", strUrl, True
        .SetRequestHeader "Authorization", "Bearer " & AccessToken
        .SetRequestHeader "Accept", "application/json"
        .Send
        .WaitForResponse
        If .Status <> 204 Then Call HandleError(.ResponseText)
    End With
End Sub

Public Function IsPublished(ByVal SiteId As String, Location As ContentTypeLocation, ByVal ContentTypeId As String, Optional ByVal ListId As String) As Boolean
    Dim strUrl As String
    Select Case Location
        Case List
            strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/lists/" & ListId & "/contentTypes/" & ContentTypeId & "/isPublished"""
        Case Site
            strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/contentTypes/" & ContentTypeId & "/isPublished"
    End Select
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
            Set objJson = New Scripting.Dictionary
            IsPublished = CBool(objJson.Item("value"))
        Else: Call HandleError(.ResponseText)
        End If
    End With
End Function

Public Sub Publish(ByVal SiteId As String, Location As ContentTypeLocation, ByVal ContentTypeId As String, Optional ByVal ListId As String)
    Dim strUrl As String
    Select Case Location
        Case List
            strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/lists/" & ListId & "/contentTypes/" & ContentTypeId & "/publish"
        Case Site
            strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/contentTypes/" & ContentTypeId & "/publish"
    End Select
    Dim objWinHttp As WinHttp.WinHttpRequest
    Set objWinHttp = New WinHttp.WinHttpRequest
    With objWinHttp
        .Open "POST", strUrl, True
        .SetRequestHeader "Authorization", "Bearer " & AccessToken
        .SetRequestHeader "Accept", "application/json"
        .Send
        .WaitForResponse
        If .Status <> 204 Then Call HandleError(.ResponseText)
    End With
End Sub

Public Function AddCopy(ByVal SiteId As String, Location As ContentTypeLocation, ByVal ContentTypeId As String, ByVal ContentTypeAddCopy As ContentTypeAddCopy, Optional ByVal ListId As String) As ContentType
    Dim strUrl As String
    Select Case Location
        Case List
            strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/lists/" & ListId & "/contentTypes/" & ContentTypeId & "/addCopy"
        Case Site
            strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/contentTypes/" & ContentTypeId & "/addCopy"
    End Select
    Dim objWinHttp As WinHttp.WinHttpRequest
    Set objWinHttp = New WinHttp.WinHttpRequest
    With objWinHttp
        .Open "POST", strUrl, True
        .SetRequestHeader "Authorization", "Bearer " & AccessToken
        .SetRequestHeader "Accept", "application/json"
        .SetRequestHeader "Content-Type", "application/json"
        .Send JsonConverter.ConvertToJson(ContentTypeAddCopy.ToJson)
        .WaitForResponse
        If .Status = 201 Then
            Dim objJson As Scripting.Dictionary
            Set objJson = JsonConverter.ParseJson(.ResponseText)
            Dim objContentType As ContentType
            Set objContentType = New ContentType
            With objContentType
                .Id = objJson.Item("id")
                .IsBuiltIn = CBool(objJson.Item("isBuiltIn"))
                .Description = objJson.Item("description")
                .Group = objJson.Item("group")
                .Hidden = CBool(objJson.Item("hidden"))
                .Name = objJson.Item("name")
                .ParentId = objJson.Item("parentId")
                .ReadOnly = CBool(objJson.Item("readOnly"))
                .Sealed = CBool(objJson.Item("sealed"))
                Dim objBase As ContentType
                Set objBase = New ContentType
                With objBase
                    .Id = objJson.Item("base")("id")
                    .IsBuiltIn = CBool(objJson.Item("base")("isBuiltIn"))
                    .Description = objJson.Item("base")("description")
                    .Group = objJson.Item("base")("group")
                    .Hidden = CBool(objJson.Item("base")("hidden"))
                    .Name = objJson.Item("base")("name")
                    .ParentId = objJson.Item("base")("parentId")
                    .ReadOnly = CBool(objJson.Item("base")("readOnly"))
                    .Sealed = CBool(objJson.Item("base")("sealed"))
                End With
                Set .Base = objBase
                Set AddCopy = objContentType
            End With
        Else: Call HandleError(.ResponseText)
        End If
    End With
End Function

Public Function AddCopyFromContentTypeHub(ByVal SiteId As String, Location As ContentTypeLocation, ByVal ContentTypeId As String, ByVal TargetContentTypeId As String, Optional ByVal ListId As String) As Variant
    Dim strUrl As String
    Select Case Location
        Case List
            strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/lists/" & ListId & "/contentTypes/" & ContentTypeId & "/addCopyFromContentTypeHub"
        Case Site
            strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/contentTypes/" & ContentTypeId & "/addCopyFromContentTypeHub"
    End Select
    Dim objWinHttp As WinHttp.WinHttpRequest
    Set objWinHttp = New WinHttp.WinHttpRequest
    Dim objJson As Scripting.Dictionary
    Set objJson = New Scripting.Dictionary
    objJson.Add "contentTypeId", TargetContentTypeId
    With objWinHttp
        .Open "POST", strUrl, True
        .SetRequestHeader "Authorization", "Bearer " & AccessToken
        .SetRequestHeader "Accept", "application/json"
        .SetRequestHeader "Content-Type", "application/json"
        .Send JsonConverter.ConvertToJson(objJson)
        .WaitForResponse
        If .Status = 200 Then
            Set objJson = JsonConverter.ParseJson(.ResponseText)
            Dim objContentType As ContentType
            Set objContentType = New ContentType
            With objContentType
                .Id = objJson.Item("id")
                .IsBuiltIn = CBool(objJson.Item("isBuiltIn"))
                .Description = objJson.Item("description")
                .Group = objJson.Item("group")
                .Hidden = CBool(objJson.Item("hidden"))
                .Name = objJson.Item("name")
                .ParentId = objJson.Item("parentId")
                .ReadOnly = CBool(objJson.Item("readOnly"))
                .Sealed = CBool(objJson.Item("sealed"))
                Dim objBase As ContentType
                Set objBase = New ContentType
                With objBase
                    .Id = objJson.Item("base")("id")
                    .IsBuiltIn = CBool(objJson.Item("base")("isBuiltIn"))
                    .Description = objJson.Item("base")("description")
                    .Group = objJson.Item("base")("group")
                    .Hidden = CBool(objJson.Item("base")("hidden"))
                    .Name = objJson.Item("base")("name")
                    .ParentId = objJson.Item("base")("parentId")
                    .ReadOnly = CBool(objJson.Item("base")("readOnly"))
                    .Sealed = CBool(objJson.Item("base")("sealed"))
                End With
                Set .Base = objBase
                Set AddCopyFromContentTypeHub = objContentType
            End With
        ElseIf .Status = 202 Then
            Set objJson = JsonConverter.ParseJson(.ResponseText)
            AddCopyFromContentTypeHub = objJson.Item("location")
        Else: Call HandleError(.ResponseText)
        End If
    End With
End Function

Public Sub AssociateWithHubSites(ByVal SiteId As String, ByVal ContentTypeId As String, ByVal HubSite As HubSite)
    Dim strUrl As String
    strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/contentTypes/" & ContentTypeId & "/associateWithHubSites"
    Dim objWinHttp As WinHttp.WinHttpRequest
    Set objWinHttp = New WinHttp.WinHttpRequest
    With objWinHttp
        .Open "POST", strUrl, True
        .SetRequestHeader "Authorization", "Bearer " & AccessToken
        .SetRequestHeader "Accept", "application/json"
        .SetRequestHeader "Content-Type", "application/json"
        .Send JsonConverter.ConvertToJson(HubSite.ToJson)
        .WaitForResponse
        If .Status <> 204 Then Call HandleError(.ResponseText)
    End With
End Sub

Public Sub CopyToDefaultContentLocation(ByVal SiteId As String, ByVal ContentTypeId As String, ByVal SourceFile As SourceFile)
    Dim strUrl As String
    strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/contentTypes/" & ContentTypeId & "/copyToDefaultContentLocation"
    Dim objWinHttp As WinHttp.WinHttpRequest
    Set objWinHttp = New WinHttp.WinHttpRequest
    With objWinHttp
        .Open "POST", strUrl, True
        .SetRequestHeader "Authorization", "Bearer " & AccessToken
        .SetRequestHeader "Accept", "application/json"
        .SetRequestHeader "Content-Type", "application/json"
        .Send JsonConverter.ConvertToJson(SourceFile.ToJson)
        .WaitForResponse
        If .Status <> 204 Then Call HandleError(.ResponseText)
    End With
End Sub

Public Function GetCompatibleHubContentTypes(ByVal SiteId As String, Location As ContentTypeLocation, Optional ByVal ListId As String) As Collection
    Dim strUrl As String
    Select Case Location
        Case List
            strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/lists/" & ListId & "/contentTypes/getCompatibleHubContentTypes"
        Case Site
            strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/contentTypes/getCompatibleHubContentTypes"
    End Select
    Dim objWinHttp As WinHttp.WinHttpRequest
    Set objWinHttp = New WinHttp.WinHttpRequest
    With objWinHttp
        .Open "GET", strUrl, True
        .SetRequestHeader "Authorization", "Bearer " & AccessToken
        .SetRequestHeader "Accept", "application/json"
        .Send
        .WaitForResponse
        Dim colContentTypes As New Collection
        If .Status = 200 Then
            Dim objJson As Scripting.Dictionary
            Set objJson = JsonConverter.ParseJson(.ResponseText)
            Dim i As Long
            Dim objContentType As ContentType, objBase As ContentType
            If objJson.Item("value").Count > 0 Then
                For i = 1 To objJson.Item("value").Count
                    Set objContentType = New ContentType
                    With objContentType
                        .Id = objJson.Item("value")(i)("id")
                        .IsBuiltIn = CBool(objJson.Item("value")(i)("isBuiltIn"))
                        .Description = objJson.Item("value")(i)("description")
                        .Group = objJson.Item("value")(i)("group")
                        .Hidden = CBool(objJson.Item("value")(i)("hidden"))
                        .Name = objJson.Item("value")(i)("name")
                        .ParentId = objJson.Item("value")(i)("parentId")
                        .ReadOnly = CBool(objJson.Item("value")(i)("readOnly"))
                        .Sealed = CBool(objJson.Item("value")(i)("sealed"))
                        If objJson.Item("value")(i).Exists("base") Then
                            Set objBase = New ContentType
                            With objBase
                                .Id = objJson.Item("value")(i)("base")("id")
                                .IsBuiltIn = CBool(objJson.Item("value")(i)("base")("isBuiltIn"))
                                .Description = objJson.Item("value")(i)("base")("description")
                                .Group = objJson.Item("value")(i)("base")("group")
                                .Hidden = CBool(objJson.Item("value")(i)("base")("hidden"))
                                .Name = objJson.Item("value")(i)("base")("name")
                                .ParentId = objJson.Item("value")(i)("base")("parentId")
                                .ReadOnly = CBool(objJson.Item("value")(i)("base")("readOnly"))
                                .Sealed = CBool(objJson.Item("value")(i)("base")("sealed"))
                            End With
                            Set .Base = objBase
                        End If
                    End With
                    colContentTypes.Add objContentType
                Next
                Set GetCompatibleHubContentTypes = colContentTypes
            End If
        Else: Call HandleError(.ResponseText)
        End If
    End With
End Function

Private Sub HandleError(ByVal Source As String)
    Dim objJson As Scripting.Dictionary
    Set objJson = JsonConverter.ParseJson(Source)
    Err.Raise vbObjectError, objJson.Item("error")("code"), objJson.Item("error")("message")
End Sub
