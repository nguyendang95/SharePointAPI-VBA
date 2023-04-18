VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ListItemResource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum ListItemLocation
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

Public Function GetListItem(ByVal SiteId As String, ByVal ListId As String, Optional ByVal ODataQuery As String) As Collection
    Dim strUrl As String
    strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/lists/" & ListId & "/items"
    If ODataQuery <> vbNullString Then strUrl = strUrl & "?" & ODataQuery
    Dim objWinHttp As WinHttp.WinHttpRequest
    Set objWinHttp = New WinHttp.WinHttpRequest
    With objWinHttp
        .Open "GET", strUrl, True
        .SetRequestHeader "Accept", "application/json"
        .SetRequestHeader "Authorization", "Bearer " & AccessToken
        .Send
        .WaitForResponse
        Dim colListItems As New Collection
        If .Status = 200 Then
            Dim objJson As Scripting.Dictionary
            Set objJson = JsonConverter.ParseJson(.ResponseText)
            Dim objCreatedBy As IdentitySet
            Dim objLastModifiedBy As IdentitySet
            Dim objIdentity As Identity
            Dim objContentType As ContentTypeInfo
            Dim i As Long
            Dim objListItem As ListItem
            Dim objItemReference As ItemReference
            If objJson.Item("value").Count > 0 Then
                For i = 1 To objJson.Item("value").Count
                    Set objListItem = New ListItem
                    With objListItem
                        .Id = objJson.Item("value")(i)("id")
                        .CreatedDateTime = objJson.Item("value")(i)("createdDateTime")
                        .WebUrl = objJson.Item("value")(i)("webUrl")
                        .ETag = objJson.Item("value")(i)("eTag")
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
                        If objJson.Item("value")(i).Exists("parentReference") Then
                            Set objItemReference = New ItemReference
                            With objItemReference
                                .SiteId = objJson.Item("value")(i)("parentReference")("siteId")
                            End With
                            Set .ParentReference = objItemReference
                        End If
                        If objJson.Item("value")(i).Exists("contentType") Then
                            Set objContentType = New ContentTypeInfo
                            With objContentType
                                .Id = objJson.Item("value")(i)("contentType")("id")
                                .Name = objJson.Item("value")(i)("contentType")("name")
                            End With
                            Set .ContentType = objContentType
                        End If
                        If objJson.Item("value")(i).Exists("fields") Then Set .Fields = objJson.Item("value")(i)("fields")
                        colListItems.Add objListItem
                    End With
                Next
            End If
        Else: Call HandleError(.ResponseText)
        End If
    End With
    Set GetListItem = colListItems
End Function

Public Function Create(ByVal SiteId As String, ByVal ListId As String, ByVal ListItem As ListItem, Optional ByVal IfMatch As String) As ListItem
    Dim strUrl As String
    strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/lists/" & ListId & "/items"
    Dim objWinHttp As WinHttp.WinHttpRequest
    Set objWinHttp = New WinHttp.WinHttpRequest
    With objWinHttp
        .Open "POST", strUrl, True
        .SetRequestHeader "Accept", "application/json"
        .SetRequestHeader "Authorization", "Bearer " & AccessToken
        .SetRequestHeader "Content-Type", "application/json"
        If IfMatch <> vbNullString Then .SetRequestHeader "if-match", IfMatch
        .Send JsonConverter.ConvertToJson(ListItem.ToJson)
        .WaitForResponse
        If .Status = 200 Then
            Dim objJson As Scripting.Dictionary
            Set objJson = JsonConverter.ParseJson(.ResponseText)
            Dim objCreatedBy As IdentitySet
            Dim objLastModifiedBy As IdentitySet
            Dim objIdentity As Identity
            Dim objContentType As ContentTypeInfo
            Dim i As Long
            Dim objListItem As ListItem
            Dim objItemReference As ItemReference
            Set objListItem = New ListItem
            With objListItem
                .Id = objJson.Item("id")
                .CreatedDateTime = objJson.Item("createdDateTime")
                .WebUrl = objJson.Item("webUrl")
                .ETag = objJson.Item("eTag")
                Set objCreatedBy = New IdentitySet
                Set objIdentity = New Identity
                If objJson.Exists("createdBy") Then
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
                If objJson.Exists("parentReference") Then
                    Set objItemReference = New ItemReference
                    With objItemReference
                        .SiteId = objJson.Item("parentReference")("siteId")
                    End With
                    Set .ParentReference = objItemReference
                End If
                If objJson.Exists("contentType") Then
                    Set objContentType = New ContentTypeInfo
                    With objContentType
                        .Id = objJson.Item("contentType")("id")
                        .Name = objJson.Item("contentType")("name")
                    End With
                    Set .ContentType = objContentType
                End If
                If objJson.Exists("fields") Then Set .Fields = objJson.Item("fields")
            End With
        Else: Call HandleError(.ResponseText)
        End If
    End With
    Set Create = objListItem
End Function

Public Sub Delete(ByVal SiteId As String, ByVal ListId As String, ByVal ItemId As String, Optional ByVal IfMatch As String)
    Dim strUrl As String
    strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/lists/" & ListId & "/items/" & ItemId
    Dim objWinHttp As WinHttp.WinHttpRequest
    Set objWinHttp = New WinHttp.WinHttpRequest
    With objWinHttp
        .Open "DELETE", strUrl, True
        .SetRequestHeader "Accept", "application/json"
        .SetRequestHeader "Authorization", "Bearer " & AccessToken
        If IfMatch <> vbNullString Then .SetRequestHeader "if-match", IfMatch
        .Send
        .WaitForResponse
        If .Status <> 204 Then
            Call HandleError(.ResponseText)
        End If
    End With
End Sub

Public Function ListVersions(ByVal SiteId As String, ListItemLocation As ListItemLocation, Optional ByVal ListId As String, Optional ByVal ItemId As String) As Collection
    Dim strUrl As String
    Select Case ListItemLocation
        Case Site
            strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/items/" & ItemId & "/versions"
        Case List
            strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/lists/" & ListId & "/items/" & ItemId * "/versions"
    End Select
    Dim objWinHttp As WinHttp.WinHttpRequest
    Set objWinHttp = New WinHttp.WinHttpRequest
    With objWinHttp
        .Open "GET", strUrl, True
        .SetRequestHeader "Authorization", "Bearer " & AccessToken
        .SetRequestHeader "Accept", "application/json"
        .Send
        .WaitForResponse
        Dim colListItemVersions As New Collection
        If .Status = 200 Then
            Dim objJson As Scripting.Dictionary
            Set objJson = JsonConverter.ParseJson(.ResponseText)
            Dim objListItemVersion As ListItemVersion
            Dim objLastModifiedBy As IdentitySet
            Dim objIdentity As Identity
            Dim i As Long
            If objJson.Item("value").Count > 0 Then
                For i = 1 To objJson.Item("value").Count
                    Set objListItemVersion = New ListItemVersion
                    With objListItemVersion
                        .Id = objJson.Item("value")(i)("id")
                        .LastModifiedDateTime = objJson.Item("value")(i)("lastModifiedDateTime")
                        If objJson("value")(i).Exists("lastModifiedBy") Then
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
                    End With
                    colListItemVersions.Add objListItemVersion
                Next
                Set ListVersions = colListItemVersions
            End If
        End If
    End With
End Function

Public Function GetItemAnalyticsByListOnSite(ByVal SiteId As String, ByVal ListId As String, ByVal ItemId As String, TimeRange As PredefinedTimeRange, Optional ByVal ODataQuery As String) As ItemAnalytics
    Dim objWinHttp As WinHttp.WinHttpRequest
    Set objWinHttp = New WinHttp.WinHttpRequest
    Dim strUrl As String
    strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/lists/" & ListId & "/items/" & ItemId & "/analytics/"
    With objWinHttp
        If TimeRange = AllTime Then
            strUrl = strUrl & "allTime"
        Else: strUrl = strUrl & "lastSevenDays"
        End If
        If ODataQuery <> vbNullString Then
            strUrl = strUrl & "?" & ODataQuery
        End If
        .Open "GET", strUrl, True
        .SetRequestHeader "Authorization", "Bearer " & AccessToken
        .SetRequestHeader "Accept", "application/json"
        .Send
        .WaitForResponse
        If .Status = 200 Then
            Dim objJson As Scripting.Dictionary
            Set objJson = JsonConverter.ParseJson(.ResponseText)
            Dim objIncompletedData As IncompleteData
            Dim objItemActionStat As ItemActionStat
            Dim objItemActivityStat As ItemActivityStat
            Dim objItemAnalytics As ItemAnalytics
            If objJson.Exists("incompleteData") Then
                Set objIncompletedData = New IncompleteData
                objIncompletedData.WasThrottled = CBool(objJson.Item("incompleteData")("wasThrottled"))
            End If
            If objJson.Exists("access") Then
            Set objItemActionStat = New ItemActionStat
                With objItemActionStat
                    .ActionCount = objJson.Item("access")("actionCount")
                    .ActorCount = objJson.Item("access")("actorCount")
                End With
            End If
            Set objItemActivityStat = New ItemActivityStat
            With objItemActivityStat
                .StartDateTime = objJson.Item("startDateTime")
                .EndDateTime = objJson.Item("endDateTime")
                .IsTrending = CBool(objJson.Item("isTrending"))
                If Not objIncompletedData Is Nothing Then Set .IncompletedData = objIncompletedData
                If Not objItemActionStat Is Nothing Then Set .Access = objItemActionStat
            End With
            Set objItemAnalytics = New ItemAnalytics
            With objItemAnalytics
                If TimeRange = AllTime Then
                    Set .AllTime = objItemActivityStat
                Else: Set .LastSevenDays = objItemActivityStat
                End If
            End With
            Set GetItemAnalyticsByListOnSite = objItemAnalytics
        Else: Call HandleError(.ResponseText)
        End If
    End With
End Function

Public Function GetItemAnalyticsOnSite(ByVal SiteId As String, TimeRange As PredefinedTimeRange, Optional ByVal ODataQuery As String) As ItemAnalytics
    Dim objWinHttp As WinHttp.WinHttpRequest
    Set objWinHttp = New WinHttp.WinHttpRequest
    Dim strUrl As String
    strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/analytics/"
    With objWinHttp
        If TimeRange = AllTime Then
            strUrl = strUrl & "allTime"
        Else: strUrl = strUrl & "lastSevenDays"
        End If
        If ODataQuery <> vbNullString Then
            strUrl = strUrl & "?" & ODataQuery
        End If
        .Open "GET", strUrl, True
        .SetRequestHeader "Authorization", "Bearer " & AccessToken
        .SetRequestHeader "Accept", "application/json"
        .Send
        .WaitForResponse
        If .Status = 200 Then
            Dim objJson As Scripting.Dictionary
            Set objJson = JsonConverter.ParseJson(.ResponseText)
            Dim objIncompletedData As IncompleteData
            Dim objItemActionStat As ItemActionStat
            Dim objItemActivityStat As ItemActivityStat
            Dim objItemAnalytics As ItemAnalytics
            If objJson.Exists("incompleteData") Then
                Set objIncompletedData = New IncompleteData
                With objIncompletedData
                    .WasThrottled = CBool(objJson.Item("incompleteData")("wasThrottled"))
                End With
            End If
            If objJson.Exists("access") Then
                Set objItemActionStat = New ItemActionStat
                With objItemActionStat
                    .ActionCount = objJson.Item("access")("actionCount")
                    .ActorCount = objJson.Item("access")("actorCount")
                End With
            End If
            Set objItemActivityStat = New ItemActivityStat
            With objItemActivityStat
                .StartDateTime = objJson.Item("startDateTime")
                .EndDateTime = objJson.Item("endDateTime")
                .IsTrending = CBool(objJson.Item("isTrending"))
                If Not objIncompletedData Is Nothing Then Set .IncompletedData = objIncompletedData
                If Not objItemActionStat Is Nothing Then Set .Access = objItemActionStat
            End With
            Set objItemAnalytics = New ItemAnalytics
            With objItemAnalytics
                If TimeRange = AllTime Then
                    Set .AllTime = objItemActivityStat
                Else: Set .LastSevenDays = objItemActivityStat
                End If
            End With
            Set GetItemAnalyticsOnSite = objItemAnalytics
        Else: Call HandleError(.ResponseText)
        End If
    End With
End Function

Public Function GetItemAnalyticsOnDrive(ByVal DriveId As String, ByVal ItemId As String, TimeRange As PredefinedTimeRange, Optional ByVal ODataQuery As String) As ItemAnalytics
    Dim objWinHttp As WinHttp.WinHttpRequest
    Set objWinHttp = New WinHttp.WinHttpRequest
    Dim strUrl As String
    strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "items/" & ItemId & "/analytics/"
    With objWinHttp
        If TimeRange = AllTime Then
            strUrl = strUrl & "allTime"
        Else: strUrl = strUrl & "lastSevenDays"
        End If
        .Open "GET", strUrl, True
        .SetRequestHeader "Authorization", "Bearer " & AccessToken
        .SetRequestHeader "Accept", "application/json"
        .Send
        .WaitForResponse
        If .Status = 200 Then
            Dim objJson As Scripting.Dictionary
            Set objJson = JsonConverter.ParseJson(.ResponseText)
            Dim objIncompletedData As IncompleteData
            Dim objItemActionStat As ItemActionStat
            Dim objItemActivityStat As ItemActivityStat
            Dim objItemAnalytics As ItemAnalytics
            If objJson.Exists("incompleteData") Then
                Set objIncompletedData = New IncompleteData
                With objIncompletedData
                    .WasThrottled = CBool(objJson.Item("incompleteData")("wasThrottled"))
                End With
            End If
            If objJson.Exists("access") Then
                Set objItemActionStat = New ItemActionStat
                With objItemActionStat
                    .ActionCount = objJson.Item("access")("actionCount")
                    .ActorCount = objJson.Item("access")("actorCount")
                End With
            End If
            Set objItemActivityStat = New ItemActivityStat
            With objItemActivityStat
                .StartDateTime = objJson.Item("startDateTime")
                .EndDateTime = objJson.Item("endDateTime")
                .IsTrending = CBool(objJson.Item("isTrending"))
                If Not objIncompletedData Is Nothing Then Set .IncompletedData = objIncompletedData
                If Not objItemActionStat Is Nothing Then Set .Access = objItemActionStat
            End With
            Set objItemAnalytics = New ItemAnalytics
            With objItemAnalytics
                If TimeRange = AllTime Then
                    Set .AllTime = objItemActivityStat
                Else: Set .LastSevenDays = objItemActivityStat
                End If
            End With
            Set GetItemAnalyticsOnDrive = objItemAnalytics
        Else: Call HandleError(.ResponseText)
        End If
    End With
End Function

Private Sub HandleError(ByVal Source As String)
    Dim objJson As Scripting.Dictionary
    Set objJson = JsonConverter.ParseJson(Source)
    Err.Raise vbObjectError, objJson.Item("error")("code"), objJson.Item("error")("message")
End Sub
