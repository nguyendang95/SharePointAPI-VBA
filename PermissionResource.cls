VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PermissionResource"
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

Public Function GetSitePermission(ByVal SiteId As String, ByVal PermissionId As String, Optional ODataQuery As String) As Permission
    Dim strUrl As String
    strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/permissions/" & PermissionId
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
            Set GetSitePermission = FetchSitePermission(.ResponseText)
        End If
    End With
End Function

Private Function FetchSitePermission(ByVal ResponseText As String) As Permission
    Dim i As Long, j As Long
    Dim objJson As Scripting.Dictionary
    Set objJson = JsonConverter.ParseJson(ResponseText)
    Dim objGrantedToIdentitiesV2 As SharePointIdentitySet
    Dim objThumbnailSet As ThumbnailSet
    Dim colThumbnailSets As New Collection
    Dim objSharePointIdentity As SharePointIdentity
    Dim objIdentity As Identity
    Dim objThumbnail As Thumbnail
    Dim colGrantedToIdentitiesV2 As New Collection
    Dim objPermission As Permission
    Set objPermission = New Permission
    With objPermission
        .Id = objJson.Item("id")
        If objJson.Exists("roles") Then Set .Roles = objJson.Item("roles")
        If objJson.Exists("grantedToIdentitiesV2") Then
            Set objGrantedToIdentitiesV2 = New SharePointIdentitySet
            With objGrantedToIdentitiesV2
                If objJson.Item("grantedToIdentitiesV2").Count > 0 Then
                    For i = 1 To objJson.Item("grantedToIdentitiesV2").Count
                        If objJson.Item("grantedToIdentitiesV2")(i).Exists("application") Then
                            Set objIdentity = New Identity
                            objIdentity.DisplayName = objJson.Item("grantedToIdentitiesV2")(i)("application")("displayName")
                            objIdentity.Id = objJson.Item("grantedToIdentitiesV2")(i)("application")("id")
                            Set .Application = objIdentity
                        ElseIf objJson.Item("grantedToIdentitiesV2")(i).Exists("device") Then
                            Set objIdentity = New Identity
                            objIdentity.DisplayName = objJson.Item("grantedToIdentitiesV2")(i)("device")("displayName")
                            objIdentity.Id = objJson.Item("grantedToIdentitiesV2")(i)("device")("id")
                            Set .Device = objIdentity
                        ElseIf objJson.Item("grantedToIdentitiesV2")(i).Exists("group") Then
                            Set objIdentity = New Identity
                            objIdentity.DisplayName = objJson.Item("grantedToIdentitiesV2")(i)("group")("displayName")
                            objIdentity.Id = objJson.Item("grantedToIdentitiesV2")(i)("group")("id")
                            Set .Group = objIdentity
                        ElseIf objJson.Item("grantedToIdentitiesV2")(i).Exists("user") Then
                            Set objIdentity = New Identity
                            objIdentity.DisplayName = objJson.Item("grantedToIdentitiesV2")(i)("user")("displayName")
                            objIdentity.Id = objJson.Item("grantedToIdentitiesV2")(i)("user")("id")
                            Set .User = objIdentity
                        ElseIf objJson.Item("grantedToIdentitiesV2")(i).Exists("siteGroup") Then
                            Set objSharePointIdentity = New Identity
                            objSharePointIdentity.DisplayName = objJson.Item("grantedToIdentitiesV2")(i)("siteGroup")("displayName")
                            objSharePointIdentity.Id = objJson.Item("grantedToIdentitiesV2")(i)("siteGroup")("id")
                            objSharePointIdentity.LoginName = objJson.Item("grantedToIdentitiesV2")(i)("siteGroup")("loginName")
                            If objJson.Item("grantedToIdentitiesV2")(i)("siteGroup")("thumbnails").Count > 0 Then
                                For j = 1 To objJson.Item("grantedToIdentitiesV2")(i)("siteGroup")("thumbnails").Count
                                    Set objThumbnailSet = New ThumbnailSet
                                    With objThumbnailSet
                                        .Id = objJson.Item("grantedToIdentitiesV2")(i)("siteGroup")("thumbnails")(j)("id")
                                        Set objThumbnail = New Thumbnail
                                        If objJson.Item("grantedToIdentitiesV2")(i)("siteGroup")("thumbnails")(j).Exists("large") Then
                                            With objThumbnail
                                                .Height = objJson.Item("grantedToIdentitiesV2")(i)("siteGroup")("thumbnails")(j)("large")("height")
                                                .SourceItemId = objJson.Item("grantedToIdentitiesV2")(i)("siteGroup")("thumbnails")(j)("large")("sourceItemId")
                                                .Url = objJson.Item("grantedToIdentitiesV2")(i)("siteGroup")("thumbnails")(j)("large")("url")
                                                .Width = objJson.Item("grantedToIdentitiesV2")(i)("siteGroup")("thumbnails")(j)("large")("width")
                                            End With
                                            Set .Large = objThumbnail
                                        ElseIf objJson.Item("grantedToIdentitiesV2")(i)("siteGroup")("thumbnails")(j).Exists("medium") Then
                                            With objThumbnail
                                                .Height = objJson.Item("grantedToIdentitiesV2")(i)("siteGroup")("thumbnails")(j)("medium")("height")
                                                .SourceItemId = objJson.Item("grantedToIdentitiesV2")(i)("siteGroup")("thumbnails")(j)("medium")("sourceItemId")
                                                .Url = objJson.Item("grantedToIdentitiesV2")(i)("siteGroup")("thumbnails")(j)("medium")("url")
                                                .Width = objJson.Item("grantedToIdentitiesV2")(i)("siteGroup")("thumbnails")(j)("medium")("width")
                                            End With
                                            Set .Medium = objThumbnail
                                        ElseIf objJson.Item("grantedToIdentitiesV2")(i)("siteGroup")("thumbnails")(j).Exists("small") Then
                                            With objThumbnail
                                                .Height = objJson.Item("grantedToIdentitiesV2")(i)("siteGroup")("thumbnails")(j)("small")("height")
                                                .SourceItemId = objJson.Item("grantedToIdentitiesV2")(i)("siteGroup")("thumbnails")(j)("small")("sourceItemId")
                                                .Url = objJson.Item("grantedToIdentitiesV2")(i)("siteGroup")("thumbnails")(j)("small")("url")
                                                .Width = objJson.Item("grantedToIdentitiesV2")(i)("siteGroup")("thumbnails")(j)("small")("width")
                                            End With
                                            Set .Small = objThumbnail
                                        ElseIf objJson.Item("grantedToIdentitiesV2")(i)("siteGroup")("thumbnails")(j).Exists("source") Then
                                            With objThumbnail
                                                .Height = objJson.Item("grantedToIdentitiesV2")(i)("siteGroup")("thumbnails")(j)("source")("height")
                                                .SourceItemId = objJson.Item("grantedToIdentitiesV2")(i)("siteGroup")("thumbnails")(j)("source")("sourceItemId")
                                                .Url = objJson.Item("grantedToIdentitiesV2")(i)("siteGroup")("thumbnails")(j)("source")("url")
                                                .Width = objJson.Item("grantedToIdentitiesV2")(i)("siteGroup")("thumbnails")(j)("source")("width")
                                            End With
                                            Set .Source = objThumbnail
                                        End If
                                    End With
                                    colThumbnailSets.Add objThumbnail
                                Next
                                Set objSharePointIdentity.Thumbnails = colThumbnailSets
                            End If
                            Set .SiteGroup = objSharePointIdentity
                        ElseIf objJson.Item("grantedToIdentitiesV2")(i).Exists("siteUser") Then
                            Set objSharePointIdentity = New Identity
                            objSharePointIdentity.DisplayName = objJson.Item("grantedToIdentitiesV2")(i)("siteUser")("displayName")
                            objSharePointIdentity.Id = objJson.Item("grantedToIdentitiesV2")(i)("siteUser")("id")
                            objSharePointIdentity.LoginName = objJson.Item("grantedToIdentitiesV2")(i)("siteUser")("loginName")
                            If objJson.Item("grantedToIdentitiesV2")(i)("siteUser").Exists("thumbnails") And IsEmpty(objJson.Item("grantedToIdentitiesV2")(i)("siteUser")("thumbnails")) = False Then
                                    If objJson.Item("grantedToIdentitiesV2")(i)("siteUser")("thumbnails").Count > 0 Then
                                        For j = 1 To objJson.Item("grantedToIdentitiesV2")(i)("siteUser")("thumbnails").Count
                                            Set objThumbnailSet = New ThumbnailSet
                                            With objThumbnailSet
                                                .Id = objJson.Item("grantedToIdentitiesV2")(i)("siteUser")("thumbnails")(j)("id")
                                                Set objThumbnail = New Thumbnail
                                                If objJson.Item("grantedToIdentitiesV2")(i)("siteUser")("thumbnails")(j).Exists("large") Then
                                                    With objThumbnail
                                                        .Height = objJson.Item("grantedToIdentitiesV2")(i)("siteUser")("thumbnails")(j)("large")("height")
                                                        .SourceItemId = objJson.Item("grantedToIdentitiesV2")(i)("siteUser")("thumbnails")(j)("large")("sourceItemId")
                                                        .Url = objJson.Item("grantedToIdentitiesV2")(i)("siteUser")("thumbnails")(j)("large")("url")
                                                        .Width = objJson.Item("grantedToIdentitiesV2")(i)("siteUser")("thumbnails")(j)("large")("width")
                                                    End With
                                                    Set .Large = objThumbnail
                                                ElseIf objJson.Item("grantedToIdentitiesV2")(i)("siteUser")("thumbnails")(j).Exists("medium") Then
                                                    With objThumbnail
                                                        .Height = objJson.Item("grantedToIdentitiesV2")(i)("siteUser")("thumbnails")(j)("medium")("height")
                                                        .SourceItemId = objJson.Item("grantedToIdentitiesV2")(i)("siteUser")("thumbnails")(j)("medium")("sourceItemId")
                                                        .Url = objJson.Item("grantedToIdentitiesV2")(i)("siteUser")("thumbnails")(j)("medium")("url")
                                                        .Width = objJson.Item("grantedToIdentitiesV2")(i)("siteUser")("thumbnails")(j)("medium")("width")
                                                    End With
                                                    Set .Medium = objThumbnail
                                                ElseIf objJson.Item("grantedToIdentitiesV2")(i)("siteUser")("thumbnails")(j).Exists("small") Then
                                                    With objThumbnail
                                                        .Height = objJson.Item("grantedToIdentitiesV2")(i)("siteUser")("thumbnails")(j)("small")("height")
                                                        .SourceItemId = objJson.Item("grantedToIdentitiesV2")(i)("siteUser")("thumbnails")(j)("small")("sourceItemId")
                                                        .Url = objJson.Item("grantedToIdentitiesV2")(i)("siteUser")("thumbnails")(j)("small")("url")
                                                        .Width = objJson.Item("grantedToIdentitiesV2")(i)("siteUser")("thumbnails")(j)("small")("width")
                                                    End With
                                                    Set .Small = objThumbnail
                                                ElseIf objJson.Item("grantedToIdentitiesV2")(i)("siteUser")("thumbnails")(j).Exists("source") Then
                                                    With objThumbnail
                                                        .Height = objJson.Item("grantedToIdentitiesV2")(i)("siteUser")("thumbnails")(j)("source")("height")
                                                        .SourceItemId = objJson.Item("grantedToIdentitiesV2")(i)("siteUser")("thumbnails")(j)("source")("sourceItemId")
                                                        .Url = objJson.Item("grantedToIdentitiesV2")(i)("siteUser")("thumbnails")(j)("source")("url")
                                                        .Width = objJson.Item("grantedToIdentitiesV2")(i)("siteUser")("thumbnails")(j)("source")("width")
                                                    End With
                                                    Set .Source = objThumbnail
                                                End If
                                            End With
                                            colThumbnailSets.Add objThumbnail
                                        Next
                                    Set objSharePointIdentity.Thumbnails = colThumbnailSets
                                End If
                            End If
                            Set .SiteUser = objSharePointIdentity
                        End If
                    Next
                End If
            End With
        End If
    End With
    Set FetchSitePermission = objPermission
End Function

Public Function ListSitePermissions(ByVal SiteId As String, ByVal PermissionId As String, Optional ODataQuery As String) As Collection
    Dim strUrl As String
    strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/permissions/" & PermissionId
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
            Set ListSitePermissions = FetchPermissions(.ResponseText)
        End If
    End With
End Function

Private Function FetchSitePermissions(ByVal ResponseText) As Collection
    Dim i As Long, j As Long, k As Long
    Dim objJson As Scripting.Dictionary
    Set objJson = JsonConverter.ParseJson(ResponseText)
    Dim objGrantedToIdentitiesV2 As SharePointIdentitySet
    Dim objThumbnailSet As ThumbnailSet
    Dim colThumbnailSets As New Collection
    Dim objSharePointIdentity As SharePointIdentity
    Dim objIdentity As Identity
    Dim objThumbnail As Thumbnail
    Dim colGrantedToIdentitiesV2 As New Collection
    Dim colPermissions As New Collection
    Dim objPermission As Permission
    If objJson.Item("value").Count > 0 Then
        For k = 1 To objJson.Item("value").Count
            Set objPermission = New Permission
            With objPermission
                .Id = objJson.Item("value")(k)("id")
                If objJson.Item("value")(k).Exists("roles") Then Set .Roles = objJson.Item("value")(k)("roles")
                If objJson.Item("value")(k).Exists("grantedToIdentitiesV2") Then
                    Set objGrantedToIdentitiesV2 = New SharePointIdentitySet
                    With objGrantedToIdentitiesV2
                        If objJson.Item("value")(k)("grantedToIdentitiesV2").Count > 0 Then
                            For i = 1 To objJson.Item("value")(k)("grantedToIdentitiesV2").Count
                                If objJson.Item("value")(k)("grantedToIdentitiesV2")(i).Exists("application") Then
                                    Set objIdentity = New Identity
                                    objIdentity.DisplayName = objJson.Item("value")(k)("grantedToIdentitiesV2")(i)("application")("displayName")
                                    objIdentity.Id = objJson.Item("value")(k)("grantedToIdentitiesV2")(i)("application")("id")
                                    Set .Application = objIdentity
                                ElseIf objJson.Item("value")(k)("grantedToIdentitiesV2")(i).Exists("device") Then
                                    Set objIdentity = New Identity
                                    objIdentity.DisplayName = objJson.Item("value")(k)("grantedToIdentitiesV2")(i)("device")("displayName")
                                    objIdentity.Id = objJson.Item("value")(k)("grantedToIdentitiesV2")(i)("device")("id")
                                    Set .Device = objIdentity
                                ElseIf objJson.Item("value")(k)("grantedToIdentitiesV2")(i).Exists("group") Then
                                    Set objIdentity = New Identity
                                    objIdentity.DisplayName = objJson.Item("value")(k)("grantedToIdentitiesV2")(i)("group")("displayName")
                                    objIdentity.Id = objJson.Item("value")(k)("grantedToIdentitiesV2")(i)("group")("id")
                                    Set .Group = objIdentity
                                ElseIf objJson.Item("value")(k)("grantedToIdentitiesV2")(i).Exists("user") Then
                                    Set objIdentity = New Identity
                                    objIdentity.DisplayName = objJson.Item("value")(k)("grantedToIdentitiesV2")(i)("user")("displayName")
                                    objIdentity.Id = objJson.Item("value")(k)("grantedToIdentitiesV2")(i)("user")("id")
                                    Set .User = objIdentity
                                ElseIf objJson.Item("grantedToIdentitiesV2")(i).Exists("siteGroup") Then
                                    Set objSharePointIdentity = New Identity
                                    objSharePointIdentity.DisplayName = objJson.Item("value")(k)("grantedToIdentitiesV2")(i)("siteGroup")("displayName")
                                    objSharePointIdentity.Id = objJson.Item("value")(k)("grantedToIdentitiesV2")(i)("siteGroup")("id")
                                    objSharePointIdentity.LoginName = objJson.Item("value")(k)("grantedToIdentitiesV2")(i)("siteGroup")("loginName")
                                    If objJson.Item("value")(k)("grantedToIdentitiesV2")(i)("siteGroup").Exists("thumbnails") And IsEmpty(objJson.Item("value")(k)("grantedToIdentitiesV2")(i)("siteGroup")("thumbnails")) = False Then
                                        If objJson.Item("value")(k)("grantedToIdentitiesV2")(i)("siteGroup")("thumbnails").Count > 0 Then
                                            For j = 1 To objJson.Item("value")(k)("grantedToIdentitiesV2")(i)("siteGroup")("thumbnails").Count
                                                Set objThumbnailSet = New ThumbnailSet
                                                With objThumbnailSet
                                                    .Id = objJson.Item("value")(k)("grantedToIdentitiesV2")(i)("siteGroup")("thumbnails")(j)("id")
                                                    Set objThumbnail = New Thumbnail
                                                    If objJson.Item("value")(k)("grantedToIdentitiesV2")(i)("siteGroup")("thumbnails")(j).Exists("large") Then
                                                        With objThumbnail
                                                            .Height = objJson.Item("value")(k)("grantedToIdentitiesV2")(i)("siteGroup")("thumbnails")(j)("large")("height")
                                                            .SourceItemId = objJson.Item("value")(k)("grantedToIdentitiesV2")(i)("siteGroup")("thumbnails")(j)("large")("sourceItemId")
                                                            .Url = objJson.Item("value")(k)("grantedToIdentitiesV2")(i)("siteGroup")("thumbnails")(j)("large")("url")
                                                            .Width = objJson.Item("value")(k)("grantedToIdentitiesV2")(i)("siteGroup")("thumbnails")(j)("large")("width")
                                                        End With
                                                        Set .Large = objThumbnail
                                                    ElseIf objJson.Item("value")(k)("grantedToIdentitiesV2")(i)("siteGroup")("thumbnails")(j).Exists("medium") Then
                                                        With objThumbnail
                                                            .Height = objJson.Item("value")(k)("grantedToIdentitiesV2")(i)("siteGroup")("thumbnails")(j)("medium")("height")
                                                            .SourceItemId = objJson.Item("value")(k)("grantedToIdentitiesV2")(i)("siteGroup")("thumbnails")(j)("medium")("sourceItemId")
                                                            .Url = objJson.Item("value")(k)("grantedToIdentitiesV2")(i)("siteGroup")("thumbnails")(j)("medium")("url")
                                                            .Width = objJson.Item("value")(k)("grantedToIdentitiesV2")(i)("siteGroup")("thumbnails")(j)("medium")("width")
                                                        End With
                                                        Set .Medium = objThumbnail
                                                    ElseIf objJson.Item("value")(k)("grantedToIdentitiesV2")(i)("siteGroup")("thumbnails")(j).Exists("small") Then
                                                        With objThumbnail
                                                            .Height = objJson.Item("value")(k)("grantedToIdentitiesV2")(i)("siteGroup")("thumbnails")(j)("small")("height")
                                                            .SourceItemId = objJson.Item("value")(k)("grantedToIdentitiesV2")(i)("siteGroup")("thumbnails")(j)("small")("sourceItemId")
                                                            .Url = objJson.Item("value")(k)("grantedToIdentitiesV2")(i)("siteGroup")("thumbnails")(j)("small")("url")
                                                            .Width = objJson.Item("value")(k)("grantedToIdentitiesV2")(i)("siteGroup")("thumbnails")(j)("small")("width")
                                                        End With
                                                        Set .Small = objThumbnail
                                                    ElseIf objJson.Item("value")(k)("grantedToIdentitiesV2")(i)("siteGroup")("thumbnails")(j).Exists("source") Then
                                                        With objThumbnail
                                                            .Height = objJson.Item("value")(k)("grantedToIdentitiesV2")(i)("siteGroup")("thumbnails")(j)("source")("height")
                                                            .SourceItemId = objJson.Item("value")(k)("grantedToIdentitiesV2")(i)("siteGroup")("thumbnails")(j)("source")("sourceItemId")
                                                            .Url = objJson.Item("value")(k)("grantedToIdentitiesV2")(i)("siteGroup")("thumbnails")(j)("source")("url")
                                                            .Width = objJson.Item("value")(k)("grantedToIdentitiesV2")(i)("siteGroup")("thumbnails")(j)("source")("width")
                                                        End With
                                                        Set .Source = objThumbnail
                                                    End If
                                                End With
                                                colThumbnailSets.Add objThumbnail
                                            Next
                                        End If
                                        Set objSharePointIdentity.Thumbnails = colThumbnailSets
                                    End If
                                    Set .SiteGroup = objSharePointIdentity
                                ElseIf objJson.Item("value")(k)("grantedToIdentitiesV2")(i).Exists("siteUser") Then
                                    Set objSharePointIdentity = New Identity
                                    objSharePointIdentity.DisplayName = objJson.Item("value")(k)("grantedToIdentitiesV2")(i)("siteUser")("displayName")
                                    objSharePointIdentity.Id = objJson.Item("value")(k)("grantedToIdentitiesV2")(i)("siteUser")("id")
                                    objSharePointIdentity.LoginName = objJson.Item("value")(k)("grantedToIdentitiesV2")(i)("siteUser")("loginName")
                                    If objJson.Item("value")(k)("grantedToIdentitiesV2")(i)("siteUser").Exists("thumbnails") And IsEmpty(objJson.Item("value")(k)("grantedToIdentitiesV2")(i)("siteUser")("thumbnails")) = False Then
                                        If objJson.Item("value")(k)("grantedToIdentitiesV2")(i)("siteUser")("thumbnails").Count > 0 Then
                                            For j = 1 To objJson.Item("value")(k)("grantedToIdentitiesV2")(i)("siteUser")("thumbnails").Count
                                                Set objThumbnailSet = New ThumbnailSet
                                                With objThumbnailSet
                                                    .Id = objJson.Item("value")(k)("grantedToIdentitiesV2")(i)("siteUser")("thumbnails")(j)("id")
                                                    Set objThumbnail = New Thumbnail
                                                    If objJson.Item("value")(k)("grantedToIdentitiesV2")(i)("siteUser")("thumbnails")(j).Exists("large") Then
                                                        With objThumbnail
                                                            .Height = objJson.Item("value")(k)("grantedToIdentitiesV2")(i)("siteUser")("thumbnails")(j)("large")("height")
                                                            .SourceItemId = objJson.Item("value")(k)("grantedToIdentitiesV2")(i)("siteUser")("thumbnails")(j)("large")("sourceItemId")
                                                            .Url = objJson.Item("value")(k)("grantedToIdentitiesV2")(i)("siteUser")("thumbnails")(j)("large")("url")
                                                            .Width = objJson.Item("value")(k)("grantedToIdentitiesV2")(i)("siteUser")("thumbnails")(j)("large")("width")
                                                        End With
                                                        Set .Large = objThumbnail
                                                    ElseIf objJson.Item("value")(k)("grantedToIdentitiesV2")(i)("siteUser")("thumbnails")(j).Exists("medium") Then
                                                        With objThumbnail
                                                            .Height = objJson.Item("value")(k)("grantedToIdentitiesV2")(i)("siteUser")("thumbnails")(j)("medium")("height")
                                                            .SourceItemId = objJson.Item("value")(k)("grantedToIdentitiesV2")(i)("siteUser")("thumbnails")(j)("medium")("sourceItemId")
                                                            .Url = objJson.Item("value")(k)("grantedToIdentitiesV2")(i)("siteUser")("thumbnails")(j)("medium")("url")
                                                            .Width = objJson.Item("value")(k)("grantedToIdentitiesV2")(i)("siteUser")("thumbnails")(j)("medium")("width")
                                                        End With
                                                        Set .Medium = objThumbnail
                                                    ElseIf objJson.Item("value")(k)("grantedToIdentitiesV2")(i)("siteUser")("thumbnails")(j).Exists("small") Then
                                                        With objThumbnail
                                                            .Height = objJson.Item("value")(k)("grantedToIdentitiesV2")(i)("siteUser")("thumbnails")(j)("small")("height")
                                                            .SourceItemId = objJson.Item("value")(k)("grantedToIdentitiesV2")(i)("siteUser")("thumbnails")(j)("small")("sourceItemId")
                                                            .Url = objJson.Item("value")(k)("grantedToIdentitiesV2")(i)("siteUser")("thumbnails")(j)("small")("url")
                                                            .Width = objJson.Item("value")(k)("grantedToIdentitiesV2")(i)("siteUser")("thumbnails")(j)("small")("width")
                                                        End With
                                                        Set .Small = objThumbnail
                                                    ElseIf objJson.Item("value")(k)("grantedToIdentitiesV2")(i)("siteUser")("thumbnails")(j).Exists("source") Then
                                                        With objThumbnail
                                                            .Height = objJson.Item("value")(k)("grantedToIdentitiesV2")(i)("siteUser")("thumbnails")(j)("source")("height")
                                                            .SourceItemId = objJson.Item("value")(k)("grantedToIdentitiesV2")(i)("siteUser")("thumbnails")(j)("source")("sourceItemId")
                                                            .Url = objJson.Item("value")(k)("grantedToIdentitiesV2")(i)("siteUser")("thumbnails")(j)("source")("url")
                                                            .Width = objJson.Item("value")(k)("grantedToIdentitiesV2")(i)("siteUser")("thumbnails")(j)("source")("width")
                                                        End With
                                                        Set .Source = objThumbnail
                                                    End If
                                                End With
                                                colThumbnailSets.Add objThumbnail
                                            Next
                                            Set objSharePointIdentity.Thumbnails = colThumbnailSets
                                        End If
                                    End If
                                    Set .SiteUser = objSharePointIdentity
                                End If
                            Next
                        End If
                    End With
                End If
            End With
            colPermissions.Add objPermission
        Next
    End If
    Set FetchSitePermissions = colPermissions
End Function

Public Function List(Destination As Destination, Optional ItemId As String, Optional ItemPath As String, Optional SiteId As String, Optional DriveId As String, Optional UserId As String, Optional GroupId As String, Optional ODataQuery As String) As Collection
    Dim objWebUtilities As WebUtilities
    Set objWebUtilities = New WebUtilities
    Dim strUrl As String
    Select Case Destination
        Case DestinationDrive
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/permissions"
            Else: strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/items/" & ItemId & "/permissions"
            End If
        Case DestinationGroup
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/permissions"
            Else: strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/items/" & ItemId & "/permissions"
            End If
        Case DestinationOtherUser
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/permissions"
            Else: strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/items/" & ItemId & "/permissions"
            End If
        Case DestinationSignedInUser
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/me/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/permissions"
            Else: strUrl = "https://graph.microsoft.com/v1.0/me/drive/items/" & ItemId & "/permissions"
            End If
        Case DestinationSite
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/permissions"
            Else: strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/items/" & ItemId & "/permissions"
            End If
    End Select
    If ODataQuery <> vbNullString Then strUrl = strUrl & "?" & ODataQuery
    Dim objWinHttp As WinHttp.WinHttpRequest
    Set objWinHttp = New WinHttp.WinHttpRequest
    With objWinHttp
        .Open "GET", strUrl, True
        .SetRequestHeader "Authorization", "Bearer " & AccessToken
        .SetRequestHeader "Accept", "application/json"
        .Send
        .WaitForResponse
        If .Status = 200 Then
            Set List = FetchPermissions(.ResponseText)
        Else: Call HandleError(.ResponseText)
        End If
    End With
End Function

Public Function GetPermission(Destination As Destination, ByVal PermissionId As String, Optional ItemId As String, Optional ItemPath As String, Optional SiteId As String, Optional DriveId As String, Optional UserId As String, Optional GroupId As String, Optional ODataQuery As String) As Permission
    Dim objWebUtilities As WebUtilities
    Set objWebUtilities = New WebUtilities
    Dim strUrl As String
    Select Case Destination
        Case DestinationDrive
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/permissions/" & PermissionId
            Else: strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/items/" & ItemId & "/permissions/" & PermissionId
            End If
        Case DestinationGroup
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/permissions/" & PermissionId
            Else: strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/items/" & ItemId & "/permissions/" & PermissionId
            End If
        Case DestinationOtherUser
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/permissions/" & PermissionId
            Else: strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/items/" & ItemId & "/permissions/" & PermissionId
            End If
        Case DestinationSignedInUser
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/me/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/permissions/" & PermissionId
            Else: strUrl = "https://graph.microsoft.com/v1.0/me/drive/items/" & ItemId & "/permissions/" & PermissionId
            End If
        Case DestinationSite
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/permissions/" & PermissionId
            Else: strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/items/" & ItemId & "/permissions/" & PermissionId
            End If
    End Select
    If ODataQuery <> vbNullString Then strUrl = strUrl & "?" & ODataQuery
    Dim objWinHttp As WinHttp.WinHttpRequest
    Set objWinHttp = New WinHttp.WinHttpRequest
    With objWinHttp
        .Open "GET", strUrl, True
        .SetRequestHeader "Authorization", "Bearer " & AccessToken
        .SetRequestHeader "Accept", "application/json"
        .Send
        .WaitForResponse
        If .Status = 200 Then
            Set GetPermission = FetchPermission(.ResponseText)
        Else: Call HandleError(.ResponseText)
        End If
    End With
End Function

Private Function FetchPermission(ResponseText) As Permission
    Dim j As Long
    Dim objJson As Scripting.Dictionary
    Set objJson = JsonConverter.ParseJson(ResponseText)
    Dim objgrantedToV2 As SharePointIdentitySet
    Dim objThumbnailSet As ThumbnailSet
    Dim colThumbnailSets As New Collection
    Dim objSharePointIdentity As SharePointIdentity
    Dim objIdentity As Identity
    Dim objThumbnail As Thumbnail
    Dim colgrantedToV2 As New Collection
    Dim objPermission As Permission
    Set objPermission = New Permission
    With objPermission
        .Id = objJson.Item("id")
        If objJson.Exists("roles") Then Set .Roles = objJson.Item("roles")
        If objJson.Exists("grantedToV2") Then
            Set objgrantedToV2 = New SharePointIdentitySet
            With objgrantedToV2
                If objJson.Item("grantedToV2").Exists("application") Then
                    Set objIdentity = New Identity
                    objIdentity.DisplayName = objJson.Item("grantedToV2")("application")("displayName")
                    objIdentity.Id = objJson.Item("grantedToV2")("application")("id")
                    Set .Application = objIdentity
                ElseIf objJson.Item("grantedToV2").Exists("device") Then
                    Set objIdentity = New Identity
                    objIdentity.DisplayName = objJson.Item("grantedToV2")("device")("displayName")
                    objIdentity.Id = objJson.Item("grantedToV2")("device")("id")
                    Set .Device = objIdentity
                ElseIf objJson.Item("grantedToV2").Exists("group") Then
                    Set objIdentity = New Identity
                    objIdentity.DisplayName = objJson.Item("grantedToV2")("group")("displayName")
                    objIdentity.Id = objJson.Item("grantedToV2")("group")("id")
                    Set .Group = objIdentity
                ElseIf objJson.Item("grantedToV2").Exists("user") Then
                    Set objIdentity = New Identity
                    objIdentity.DisplayName = objJson.Item("grantedToV2")("user")("displayName")
                    objIdentity.Id = objJson.Item("grantedToV2")("user")("id")
                    Set .User = objIdentity
                ElseIf objJson.Item("grantedToV2").Exists("siteGroup") Then
                    Set objSharePointIdentity = New SharePointIdentity
                    objSharePointIdentity.DisplayName = objJson.Item("grantedToV2")("siteGroup")("displayName")
                    objSharePointIdentity.Id = objJson.Item("grantedToV2")("siteGroup")("id")
                    objSharePointIdentity.LoginName = objJson.Item("grantedToV2")("siteGroup")("loginName")
                    If objJson.Item("grantedToV2")("siteGroup").Exists("thumbnails") And IsEmpty(objJson.Item("grantedToV2")("siteGroup")("thumbnails")) = False Then
                        If objJson.Item("grantedToV2")("siteGroup")("thumbnails").Count > 0 Then
                            For j = 1 To objJson.Item("grantedToV2")("siteGroup")("thumbnails").Count
                                Set objThumbnailSet = New ThumbnailSet
                                With objThumbnailSet
                                    .Id = objJson.Item("grantedToV2")("siteGroup")("thumbnails")("id")
                                    Set objThumbnail = New Thumbnail
                                    If objJson.Item("grantedToV2")("siteGroup")("thumbnails").Exists("large") Then
                                        With objThumbnail
                                            .Height = objJson.Item("grantedToV2")("siteGroup")("thumbnails")(j)("large")("height")
                                            .SourceItemId = objJson.Item("grantedToV2")("siteGroup")("thumbnails")(j)("large")("sourceItemId")
                                            .Url = objJson.Item("grantedToV2")("siteGroup")("thumbnails")(j)("large")("url")
                                            .Width = objJson.Item("grantedToV2")("siteGroup")("thumbnails")(j)("large")("width")
                                        End With
                                        Set .Large = objThumbnail
                                    ElseIf objJson.Item("grantedToV2")("siteGroup")("thumbnails")(j).Exists("medium") Then
                                        With objThumbnail
                                            .Height = objJson.Item("grantedToV2")("siteGroup")("thumbnails")(j)("medium")("height")
                                            .SourceItemId = objJson.Item("grantedToV2")("siteGroup")("thumbnails")(j)("medium")("sourceItemId")
                                            .Url = objJson.Item("grantedToV2")("siteGroup")("thumbnails")(j)("medium")("url")
                                            .Width = objJson.Item("grantedToV2")("siteGroup")("thumbnails")(j)("medium")("width")
                                        End With
                                        Set .Medium = objThumbnail
                                    ElseIf objJson.Item("grantedToV2")("siteGroup")("thumbnails")(j).Exists("small") Then
                                        With objThumbnail
                                            .Height = objJson.Item("grantedToV2")("siteGroup")("thumbnails")(j)("small")("height")
                                            .SourceItemId = objJson.Item("grantedToV2")("siteGroup")("thumbnails")(j)("small")("sourceItemId")
                                            .Url = objJson.Item("grantedToV2")("siteGroup")("thumbnails")(j)("small")("url")
                                            .Width = objJson.Item("grantedToV2")("siteGroup")("thumbnails")(j)("small")("width")
                                        End With
                                        Set .Small = objThumbnail
                                    ElseIf objJson.Item("grantedToV2")("siteGroup")("thumbnails")(j).Exists("source") Then
                                        With objThumbnail
                                            .Height = objJson.Item("grantedToV2")("siteGroup")("thumbnails")(j)("source")("height")
                                            .SourceItemId = objJson.Item("grantedToV2")("siteGroup")("thumbnails")(j)("source")("sourceItemId")
                                            .Url = objJson.Item("grantedToV2")("siteGroup")("thumbnails")(j)("source")("url")
                                            .Width = objJson.Item("grantedToV2")("siteGroup")("thumbnails")(j)("source")("width")
                                        End With
                                        Set .Source = objThumbnail
                                    End If
                                End With
                                colThumbnailSets.Add objThumbnail
                            Next
                        End If
                        Set objSharePointIdentity.Thumbnails = colThumbnailSets
                    End If
                    Set .SiteGroup = objSharePointIdentity
                ElseIf objJson.Item("grantedToV2").Exists("siteUser") Then
                    Set objSharePointIdentity = New Identity
                    objSharePointIdentity.DisplayName = objJson.Item("grantedToV2")("siteUser")("displayName")
                    objSharePointIdentity.Id = objJson.Item("grantedToV2")("siteUser")("id")
                    objSharePointIdentity.LoginName = objJson.Item("grantedToV2")("siteUser")("loginName")
                    If objJson.Item("grantedToV2")("siteUser").Exists("thumbnails") And IsEmpty(objJson.Item("grantedToV2")("siteUser")("thumbnails")) = False Then
                        If objJson.Item("grantedToV2")("siteUser")("thumbnails").Count > 0 Then
                            For j = 1 To objJson.Item("grantedToV2")("siteUser")("thumbnails").Count
                                Set objThumbnailSet = New ThumbnailSet
                                With objThumbnailSet
                                    .Id = objJson.Item("grantedToV2")("siteUser")("thumbnails")(j)("id")
                                    Set objThumbnail = New Thumbnail
                                    If objJson.Item("grantedToV2")("siteUser")("thumbnails")(j).Exists("large") Then
                                        With objThumbnail
                                            .Height = objJson.Item("grantedToV2")("siteUser")("thumbnails")(j)("large")("height")
                                            .SourceItemId = objJson.Item("grantedToV2")("siteUser")("thumbnails")(j)("large")("sourceItemId")
                                            .Url = objJson.Item("grantedToV2")("siteUser")("thumbnails")(j)("large")("url")
                                            .Width = objJson.Item("grantedToV2")("siteUser")("thumbnails")(j)("large")("width")
                                        End With
                                        Set .Large = objThumbnail
                                    ElseIf objJson.Item("grantedToV2")("siteUser")("thumbnails")(j).Exists("medium") Then
                                        With objThumbnail
                                            .Height = objJson.Item("grantedToV2")("siteUser")("thumbnails")(j)("medium")("height")
                                            .SourceItemId = objJson.Item("grantedToV2")("siteUser")("thumbnails")(j)("medium")("sourceItemId")
                                            .Url = objJson.Item("grantedToV2")("siteUser")("thumbnails")(j)("medium")("url")
                                            .Width = objJson.Item("grantedToV2")("siteUser")("thumbnails")(j)("medium")("width")
                                        End With
                                        Set .Medium = objThumbnail
                                    ElseIf objJson.Item("grantedToV2")("siteUser")("thumbnails")(j).Exists("small") Then
                                        With objThumbnail
                                            .Height = objJson.Item("grantedToV2")("siteUser")("thumbnails")(j)("small")("height")
                                            .SourceItemId = objJson.Item("grantedToV2")("siteUser")("thumbnails")(j)("small")("sourceItemId")
                                            .Url = objJson.Item("grantedToV2")("siteUser")("thumbnails")(j)("small")("url")
                                            .Width = objJson.Item("grantedToV2")("siteUser")("thumbnails")(j)("small")("width")
                                        End With
                                        Set .Small = objThumbnail
                                    ElseIf objJson.Item("grantedToV2")("siteUser")("thumbnails")(j).Exists("source") Then
                                        With objThumbnail
                                            .Height = objJson.Item("grantedToV2")("siteUser")("thumbnails")(j)("source")("height")
                                            .SourceItemId = objJson.Item("grantedToV2")("siteUser")("thumbnails")(j)("source")("sourceItemId")
                                            .Url = objJson.Item("grantedToV2")("siteUser")("thumbnails")(j)("source")("url")
                                            .Width = objJson.Item("grantedToV2")("siteUser")("thumbnails")(j)("source")("width")
                                        End With
                                        Set .Source = objThumbnail
                                    End If
                                End With
                                colThumbnailSets.Add objThumbnail
                            Next
                            Set objSharePointIdentity.Thumbnails = colThumbnailSets
                        End If
                    End If
                    Set .SiteUser = objSharePointIdentity
                End If
            End With
        End If
    End With
    Set FetchPermission = objPermission
End Function

Private Function FetchPermissions(ByVal ResponseText) As Collection
    Dim j As Long, k As Long
    Dim objJson As Scripting.Dictionary
    Set objJson = JsonConverter.ParseJson(ResponseText)
    Dim objgrantedToV2 As SharePointIdentitySet
    Dim objThumbnailSet As ThumbnailSet
    Dim colThumbnailSets As New Collection
    Dim objSharePointIdentity As SharePointIdentity
    Dim objIdentity As Identity
    Dim objThumbnail As Thumbnail
    Dim colgrantedToV2 As New Collection
    Dim colPermissions As New Collection
    Dim objPermission As Permission
    If objJson.Item("value").Count > 0 Then
        For k = 1 To objJson.Item("value").Count
            Set objPermission = New Permission
            With objPermission
                .Id = objJson.Item("value")(k)("id")
                If objJson.Item("value")(k).Exists("roles") Then Set .Roles = objJson.Item("value")(k)("roles")
                If objJson.Item("value")(k).Exists("grantedToV2") Then
                    Set objgrantedToV2 = New SharePointIdentitySet
                    With objgrantedToV2
                        If objJson.Item("value")(k)("grantedToV2").Exists("application") Then
                            Set objIdentity = New Identity
                            objIdentity.DisplayName = objJson.Item("value")(k)("grantedToV2")("application")("displayName")
                            objIdentity.Id = objJson.Item("value")(k)("grantedToV2")("application")("id")
                            Set .Application = objIdentity
                        ElseIf objJson.Item("value")(k)("grantedToV2").Exists("device") Then
                            Set objIdentity = New Identity
                            objIdentity.DisplayName = objJson.Item("value")(k)("grantedToV2")("device")("displayName")
                            objIdentity.Id = objJson.Item("value")(k)("grantedToV2")("device")("id")
                            Set .Device = objIdentity
                        ElseIf objJson.Item("value")(k)("grantedToV2").Exists("group") Then
                            Set objIdentity = New Identity
                            objIdentity.DisplayName = objJson.Item("value")(k)("grantedToV2")("group")("displayName")
                            objIdentity.Id = objJson.Item("value")(k)("grantedToV2")("group")("id")
                            Set .Group = objIdentity
                        ElseIf objJson.Item("value")(k)("grantedToV2").Exists("user") Then
                            Set objIdentity = New Identity
                            objIdentity.DisplayName = objJson.Item("value")(k)("grantedToV2")("user")("displayName")
                            objIdentity.Id = objJson.Item("value")(k)("grantedToV2")("user")("id")
                            Set .User = objIdentity
                        ElseIf objJson.Item("value")(k)("grantedToV2").Exists("siteGroup") Then
                            Set objSharePointIdentity = New SharePointIdentity
                            objSharePointIdentity.DisplayName = objJson.Item("value")(k)("grantedToV2")("siteGroup")("displayName")
                            objSharePointIdentity.Id = objJson.Item("value")(k)("grantedToV2")("siteGroup")("id")
                            objSharePointIdentity.LoginName = objJson.Item("value")(k)("grantedToV2")("siteGroup")("loginName")
                            If objJson.Item("value")(k)("grantedToV2")("siteGroup").Exists("thumbnails") And IsEmpty(objJson.Item("value")(k)("grantedToV2")("siteGroup")("thumbnails")) = False Then
                                If objJson.Item("value")(k)("grantedToV2")("siteGroup")("thumbnails").Count > 0 Then
                                    For j = 1 To objJson.Item("value")(k)("grantedToV2")("siteGroup")("thumbnails").Count
                                        Set objThumbnailSet = New ThumbnailSet
                                        With objThumbnailSet
                                            .Id = objJson.Item("value")(k)("grantedToV2")("siteGroup")("thumbnails")("id")
                                            Set objThumbnail = New Thumbnail
                                            If objJson.Item("value")(k)("grantedToV2")("siteGroup")("thumbnails").Exists("large") Then
                                                With objThumbnail
                                                    .Height = objJson.Item("value")(k)("grantedToV2")("siteGroup")("thumbnails")(j)("large")("height")
                                                    .SourceItemId = objJson.Item("value")(k)("grantedToV2")("siteGroup")("thumbnails")(j)("large")("sourceItemId")
                                                    .Url = objJson.Item("value")(k)("grantedToV2")("siteGroup")("thumbnails")(j)("large")("url")
                                                    .Width = objJson.Item("value")(k)("grantedToV2")("siteGroup")("thumbnails")(j)("large")("width")
                                                End With
                                                Set .Large = objThumbnail
                                            ElseIf objJson.Item("value")(k)("grantedToV2")("siteGroup")("thumbnails")(j).Exists("medium") Then
                                                With objThumbnail
                                                    .Height = objJson.Item("value")(k)("grantedToV2")("siteGroup")("thumbnails")(j)("medium")("height")
                                                    .SourceItemId = objJson.Item("value")(k)("grantedToV2")("siteGroup")("thumbnails")(j)("medium")("sourceItemId")
                                                    .Url = objJson.Item("value")(k)("grantedToV2")("siteGroup")("thumbnails")(j)("medium")("url")
                                                    .Width = objJson.Item("value")(k)("grantedToV2")("siteGroup")("thumbnails")(j)("medium")("width")
                                                End With
                                                Set .Medium = objThumbnail
                                            ElseIf objJson.Item("value")(k)("grantedToV2")("siteGroup")("thumbnails")(j).Exists("small") Then
                                                With objThumbnail
                                                    .Height = objJson.Item("value")(k)("grantedToV2")("siteGroup")("thumbnails")(j)("small")("height")
                                                    .SourceItemId = objJson.Item("value")(k)("grantedToV2")("siteGroup")("thumbnails")(j)("small")("sourceItemId")
                                                    .Url = objJson.Item("value")(k)("grantedToV2")("siteGroup")("thumbnails")(j)("small")("url")
                                                    .Width = objJson.Item("value")(k)("grantedToV2")("siteGroup")("thumbnails")(j)("small")("width")
                                                End With
                                                Set .Small = objThumbnail
                                            ElseIf objJson.Item("value")(k)("grantedToV2")("siteGroup")("thumbnails")(j).Exists("source") Then
                                                With objThumbnail
                                                    .Height = objJson.Item("value")(k)("grantedToV2")("siteGroup")("thumbnails")(j)("source")("height")
                                                    .SourceItemId = objJson.Item("value")(k)("grantedToV2")("siteGroup")("thumbnails")(j)("source")("sourceItemId")
                                                    .Url = objJson.Item("value")(k)("grantedToV2")("siteGroup")("thumbnails")(j)("source")("url")
                                                    .Width = objJson.Item("value")(k)("grantedToV2")("siteGroup")("thumbnails")(j)("source")("width")
                                                End With
                                                Set .Source = objThumbnail
                                            End If
                                        End With
                                        colThumbnailSets.Add objThumbnail
                                    Next
                                End If
                                Set objSharePointIdentity.Thumbnails = colThumbnailSets
                            End If
                            Set .SiteGroup = objSharePointIdentity
                        ElseIf objJson.Item("value")(k)("grantedToV2").Exists("siteUser") Then
                            Set objSharePointIdentity = New Identity
                            objSharePointIdentity.DisplayName = objJson.Item("value")(k)("grantedToV2")("siteUser")("displayName")
                            objSharePointIdentity.Id = objJson.Item("value")(k)("grantedToV2")("siteUser")("id")
                            objSharePointIdentity.LoginName = objJson.Item("value")(k)("grantedToV2")("siteUser")("loginName")
                            If objJson.Item("value")(k)("grantedToV2")("siteUser").Exists("thumbnails") And IsEmpty(objJson.Item("value")(k)("grantedToV2")("siteUser")("thumbnails")) = False Then
                                If objJson.Item("value")(k)("grantedToV2")("siteUser")("thumbnails").Count > 0 Then
                                    For j = 1 To objJson.Item("value")(k)("grantedToV2")("siteUser")("thumbnails").Count
                                        Set objThumbnailSet = New ThumbnailSet
                                        With objThumbnailSet
                                            .Id = objJson.Item("value")(k)("grantedToV2")("siteUser")("thumbnails")(j)("id")
                                            Set objThumbnail = New Thumbnail
                                            If objJson.Item("value")(k)("grantedToV2")("siteUser")("thumbnails")(j).Exists("large") Then
                                                With objThumbnail
                                                    .Height = objJson.Item("value")(k)("grantedToV2")("siteUser")("thumbnails")(j)("large")("height")
                                                    .SourceItemId = objJson.Item("value")(k)("grantedToV2")("siteUser")("thumbnails")(j)("large")("sourceItemId")
                                                    .Url = objJson.Item("value")(k)("grantedToV2")("siteUser")("thumbnails")(j)("large")("url")
                                                    .Width = objJson.Item("value")(k)("grantedToV2")("siteUser")("thumbnails")(j)("large")("width")
                                                End With
                                                Set .Large = objThumbnail
                                            ElseIf objJson.Item("value")(k)("grantedToV2")("siteUser")("thumbnails")(j).Exists("medium") Then
                                                With objThumbnail
                                                    .Height = objJson.Item("value")(k)("grantedToV2")("siteUser")("thumbnails")(j)("medium")("height")
                                                    .SourceItemId = objJson.Item("value")(k)("grantedToV2")("siteUser")("thumbnails")(j)("medium")("sourceItemId")
                                                    .Url = objJson.Item("value")(k)("grantedToV2")("siteUser")("thumbnails")(j)("medium")("url")
                                                    .Width = objJson.Item("value")(k)("grantedToV2")("siteUser")("thumbnails")(j)("medium")("width")
                                                End With
                                                Set .Medium = objThumbnail
                                            ElseIf objJson.Item("value")(k)("grantedToV2")("siteUser")("thumbnails")(j).Exists("small") Then
                                                With objThumbnail
                                                    .Height = objJson.Item("value")(k)("grantedToV2")("siteUser")("thumbnails")(j)("small")("height")
                                                    .SourceItemId = objJson.Item("value")(k)("grantedToV2")("siteUser")("thumbnails")(j)("small")("sourceItemId")
                                                    .Url = objJson.Item("value")(k)("grantedToV2")("siteUser")("thumbnails")(j)("small")("url")
                                                    .Width = objJson.Item("value")(k)("grantedToV2")("siteUser")("thumbnails")(j)("small")("width")
                                                End With
                                                Set .Small = objThumbnail
                                            ElseIf objJson.Item("value")(k)("grantedToV2")("siteUser")("thumbnails")(j).Exists("source") Then
                                                With objThumbnail
                                                    .Height = objJson.Item("value")(k)("grantedToV2")("siteUser")("thumbnails")(j)("source")("height")
                                                    .SourceItemId = objJson.Item("value")(k)("grantedToV2")("siteUser")("thumbnails")(j)("source")("sourceItemId")
                                                    .Url = objJson.Item("value")(k)("grantedToV2")("siteUser")("thumbnails")(j)("source")("url")
                                                    .Width = objJson.Item("value")(k)("grantedToV2")("siteUser")("thumbnails")(j)("source")("width")
                                                End With
                                                Set .Source = objThumbnail
                                            End If
                                        End With
                                        colThumbnailSets.Add objThumbnail
                                    Next
                                    Set objSharePointIdentity.Thumbnails = colThumbnailSets
                                End If
                            End If
                            Set .SiteUser = objSharePointIdentity
                        End If
                    End With
                End If
            End With
            colPermissions.Add objPermission
        Next
    End If
    Set FetchPermissions = colPermissions
End Function

Public Function Add(Destination As Destination, ByVal Permission As Permission, Optional ItemId As String, Optional ItemPath As String, Optional SiteId As String, Optional DriveId As String, Optional UserId As String, Optional GroupId As String) As Permission
    Dim objWebUtilities As WebUtilities
    Set objWebUtilities = New WebUtilities
    Dim strUrl As String
    Select Case Destination
        Case DestinationDrive
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/invite/"
            Else: strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/items/" & ItemId & "/invite/"
            End If
        Case DestinationGroup
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/invite/"
            Else: strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/items/" & ItemId & "/invite/"
            End If
        Case DestinationOtherUser
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/invite/"
            Else: strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/items/" & ItemId & "/invite/"
            End If
        Case DestinationSignedInUser
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/me/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/invite/"
            Else: strUrl = "https://graph.microsoft.com/v1.0/me/drive/items/" & ItemId & "/invite/"
            End If
        Case DestinationSite
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/invite/"
            Else: strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/items/" & ItemId & "/invite/"
            End If
    End Select
    Dim objWinHttp As WinHttp.WinHttpRequest
    Set objWinHttp = New WinHttp.WinHttpRequest
    With objWinHttp
        .Open "POST", strUrl, True
        .SetRequestHeader "Authorization", "Bearer " & AccessToken
        .SetRequestHeader "Accept", "application/json"
        .SetRequestHeader "Content-Type", "application/json"
        .Send JsonConverter.ConvertToJson(Permission.ToJson)
        .WaitForResponse
        If .Status = 200 Then
            Set Add = FetchPermission(.ResponseText)
        Else: Call HandleError(.ResponseText)
        End If
    End With
End Function

Public Function Update(Destination As Destination, ByVal PermissionId As String, ByVal Permission As Permission, Optional ItemId As String, Optional ItemPath As String, Optional SiteId As String, Optional DriveId As String, Optional UserId As String, Optional GroupId As String, Optional IfMatch As String) As Permission
    Dim objWebUtilities As WebUtilities
    Set objWebUtilities = New WebUtilities
    Dim strUrl As String
    Select Case Destination
        Case DestinationDrive
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/permissions/" & PermissionId
            Else: strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/items/" & ItemId & "/permissions/" & PermissionId
            End If
        Case DestinationGroup
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/permissions/" & PermissionId
            Else: strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/items/" & ItemId & "/permissions/" & PermissionId
            End If
        Case DestinationOtherUser
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/permissions/" & PermissionId
            Else: strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/items/" & ItemId & "/permissions/" & PermissionId
            End If
        Case DestinationSignedInUser
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/me/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/permissions/" & PermissionId
            Else: strUrl = "https://graph.microsoft.com/v1.0/me/drive/items/" & ItemId & "/permissions/" & PermissionId
            End If
        Case DestinationSite
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/permissions/" & PermissionId
            Else: strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/items/" & ItemId & "/permissions/" & PermissionId
            End If
    End Select
    Dim objWinHttp As WinHttp.WinHttpRequest
    Set objWinHttp = New WinHttp.WinHttpRequest
    With objWinHttp
        .Open "PATCH", strUrl, True
        .SetRequestHeader "Authorization", "Bearer " & AccessToken
        .SetRequestHeader "Accept", "application/json"
        .SetRequestHeader "Content-Type", "application/json"
        If IfMatch <> vbNullString Then .SetRequestHeader "if-match", IfMatch
        .Send JsonConverter.ConvertToJson(Permission.ToJson)
        .WaitForResponse
        If .Status = 412 Then Call HandleError(.ResponseText)
        If .Status = 200 Then
            Set Update = FetchPermission(.ResponseText)
        Else: Call HandleError(.ResponseText)
        End If
    End With
End Function

Public Sub Delete(Destination As Destination, ByVal PermissionId As String, Optional ItemId As String, Optional ItemPath As String, Optional SiteId As String, Optional DriveId As String, Optional UserId As String, Optional GroupId As String, Optional IfMatch As String)
    Dim objWebUtilities As WebUtilities
    Set objWebUtilities = New WebUtilities
    Dim strUrl As String
    Select Case Destination
        Case DestinationDrive
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/permissions/" & PermissionId
            Else: strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/items/" & ItemId & "/permissions/" & PermissionId
            End If
        Case DestinationGroup
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/permissions/" & PermissionId
            Else: strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/items/" & ItemId & "/permissions/" & PermissionId
            End If
        Case DestinationOtherUser
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/permissions/" & PermissionId
            Else: strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/items/" & ItemId & "/permissions/" & PermissionId
            End If
        Case DestinationSignedInUser
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/me/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/permissions/" & PermissionId
            Else: strUrl = "https://graph.microsoft.com/v1.0/me/drive/items/" & ItemId & "/permissions/" & PermissionId
            End If
        Case DestinationSite
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/permissions/" & PermissionId
            Else: strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/items/" & ItemId & "/permissions/" & PermissionId
            End If
    End Select
    Dim objWinHttp As WinHttp.WinHttpRequest
    Set objWinHttp = New WinHttp.WinHttpRequest
    With objWinHttp
        .Open "DELETE", strUrl, True
        .SetRequestHeader "Authorization", "Bearer " & AccessToken
        .SetRequestHeader "Accept", "application/json"
        .SetRequestHeader "Content-Type", "application/json"
        If IfMatch <> vbNullString Then .SetRequestHeader "if-match", IfMatch
        .Send
        .WaitForResponse
        If .Status = 412 Then Call HandleError(.ResponseText)
        If .Status <> 204 Then Call HandleError(.ResponseText)
    End With
End Sub

Public Function GrantAccessToSharingLink(ByVal EncodedSharingUrl As String, Permission As Permission) As Permission
    Dim objWinHttp As WinHttp.WinHttpRequest
    Set objWinHttp = New WinHttp.WinHttpRequest
    With objWinHttp
        .Open "POST", "https://graph.microsoft.com/v1.0/shares/" & EncodedSharingUrl & "/permission/grant", True
        .SetRequestHeader "Authorization", "Bearer " & AccessToken
        .SetRequestHeader "Accept", "application/json"
        .SetRequestHeader "Content-Type", "application/json"
        .Send JsonConverter.ConvertToJson(Permission.ToJson)
        .WaitForResponse
        If .Status = 200 Then
            Set GrantAccessToSharingLink = FetchPermission(.ResponseText)
        Else: Call HandleError(.ResponseText)
        End If
    End With
End Function

Public Function UseSharingLinks(ByVal ShareIdOrEncodedSharingUrl As String, Optional Prefer As String) As SharedDriveItem
    Dim objWinHttp As WinHttp.WinHttpRequest
    Set objWinHttp = New WinHttp.WinHttpRequest
    With objWinHttp
        .Open "GET", "https://graph.microsoft.com/v1.0/shares/" & ShareIdOrEncodedSharingUrl
        .SetRequestHeader "Authorization", "Bearer " & AccessToken
        .SetRequestHeader "Accept", "application/json"
        .SetRequestHeader "Content-Type", "application/json"
        If Prefer <> vbNullString Then .SetRequestHeader "Prefer", Prefer
        .Send
        .WaitForResponse
        If .Status = 200 Then
            Set UseSharingLinks = FetchSharedDriveItem(.ResponseText)
        Else: Call HandleError(.ResponseText)
        End If
    End With
End Function

Private Function FetchSharedDriveItem(ByVal ResponseText As String) As SharedDriveItem
    Dim objJson As Scripting.Dictionary
    Set objJson = JsonConverter.ParseJson(ResponseText)
    Dim objSharedDriveItem As SharedDriveItem
    Dim objOwner As IdentitySet
    Dim objIdentity As Identity
    Set objSharedDriveItem = New SharedDriveItem
    With objSharedDriveItem
        If objJson.Exists("owner") Then
            Set objOwner = New IdentitySet
            Set objIdentity = New Identity
            If objJson.Item("owner").Exists("user") And IsEmpty(objJson.Item("owner")("user")) = False Then
                objIdentity.DisplayName = objJson.Item("owner")("user")("displayName")
                objIdentity.Id = objJson.Item("value")("owner")("user")("id")
                Set objOwner.User = objIdentity
            End If
            If objJson.Item("owner").Exists("application") And IsEmpty(objJson.Item("owner")("application")) = False Then
                objIdentity.DisplayName = objJson.Item("owner")("application")("displayName")
                objIdentity.Id = objJson.Item("owner")("application")("id")
                Set objOwner.Application = objIdentity
            End If
            If objJson.Item("owner").Exists("device") And IsEmpty(objJson.Item("lastModifiedBy")("device")) = False Then
                objIdentity.DisplayName = objJson.Item("owner")("device")("displayName")
                objIdentity.Id = objJson.Item("owner")("device")("id")
                Set objOwner.Device = objIdentity
            End If
        End If
        Set .Owner = objOwner
        Set FetchSharedDriveItem = objSharedDriveItem
    End With
End Function

Public Function CreateSharingLink(Destination As Destination, ByVal Permission As Permission, Optional ItemId As String, Optional ItemPath As String, Optional SiteId As String, Optional DriveId As String, Optional UserId As String, Optional GroupId As String) As Permission
    Dim objWebUtilities As WebUtilities
    Set objWebUtilities = New WebUtilities
    Dim strUrl As String
    Select Case Destination
        Case DestinationDrive
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/createLink/"
            Else: strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/items/" & ItemId & "/createLink/"
            End If
        Case DestinationGroup
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/createLink/"
            Else: strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/items/" & ItemId & "/createLink/"
            End If
        Case DestinationOtherUser
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/createLink/"
            Else: strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/items/" & ItemId & "/createLink/"
            End If
        Case DestinationSignedInUser
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/me/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/createLink/"
            Else: strUrl = "https://graph.microsoft.com/v1.0/me/drive/items/" & ItemId & "/createLink/"
            End If
        Case DestinationSite
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/createLink/"
            Else: strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/items/" & ItemId & "/createLink/"
            End If
    End Select
    Dim objWinHttp As WinHttp.WinHttpRequest
    Set objWinHttp = New WinHttp.WinHttpRequest
    With objWinHttp
        .Open "POST", strUrl, True
        .SetRequestHeader "Authorization", "Bearer " & AccessToken
        .SetRequestHeader "Accept", "application/json"
        .SetRequestHeader "Content-Type", "application/json"
        .Send JsonConverter.ConvertToJson(Permission.ToJson)
        .WaitForResponse
        If .Status = 201 Or .Status = 200 Then
            Set CreateSharingLink = FetchPermission(.ResponseText)
        Else: Call HandleError(.ResponseText)
        End If
    End With
End Function

Private Sub HandleError(ByVal Source As String)
    Dim objJson As Scripting.Dictionary
    Set objJson = JsonConverter.ParseJson(Source)
    Err.Raise vbObjectError, objJson.Item("error")("code"), objJson.Item("error")("message")
End Sub
