VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TaxonomyResource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum RelationLocation
    Term
    SetItem
End Enum

Private m_AccessToken As String

Friend Property Get AccessToken() As String
   AccessToken = m_AccessToken
End Property

Friend Property Let AccessToken(Value As String)
   m_AccessToken = Value
End Property

Public Function GetStore(ByVal SiteId As String) As Store
    Dim objWinHttp As WinHttp.WinHttpRequest
    Set objWinHttp = New WinHttp.WinHttpRequest
    With objWinHttp
        .Open "GET", "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/termStore", True
        .SetRequestHeader "Authorization", "Bearer " & AccessToken
        .SetRequestHeader "Accept", "application/json"
        .Send
        .WaitForResponse
        If .Status = 200 Then
            Dim objJson As Scripting.Dictionary
            Set objJson = JsonConverter.ParseJson(.ResponseText)
            Dim objStore As Store
            Set objStore = New Store
            With objStore
                .Id = objJson.Item("id")
                .DefaultLanguageTag = objJson.Item("defaultLanguageTag")
                Set .LanguageTags = objJson.Item("languageTags")
            End With
            Set GetStore = objStore
        End If
    End With
End Function

Public Function UpdateStore(ByVal SiteId As String, ByVal Store As Store) As Store
    Dim objWinHttp As WinHttp.WinHttpRequest
    Set objWinHttp = New WinHttp.WinHttpRequest
    With objWinHttp
        .Open "PATCH", "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/termStore", True
        .SetRequestHeader "Authorization", "Bearer " & AccessToken
        .SetRequestHeader "Accept", "application/json"
        .SetRequestHeader "Content-Type", "application/json"
        .Send JsonConverter.ConvertToJson(Store.ToJson)
        .WaitForResponse
        If .Status = 200 Then
            Dim objJson As Scripting.Dictionary
            Set objJson = JsonConverter.ParseJson(.ResponseText)
            Dim objStore As Store
            Set objStore = New Store
            With objStore
                .Id = objJson.Item("id")
                .DefaultLanguageTag = objJson.Item("defaultLanguageTag")
                Set .LanguageTags = objJson.Item("languageTags")
            End With
            Set UpdateStore = objStore
        Else: Call HandleError(.ResponseText)
        End If
    End With
End Function

Public Function ListGroups(ByVal SiteId As String, Optional ByVal ODataQuery As String) As Collection
    Dim strUrl As String
    strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/termStore"
    If ODataQuery <> vbNullString Then strUrl = strUrl & "?" & ODataQuery
    Dim objWinHttp As WinHttp.WinHttpRequest
    Set objWinHttp = New WinHttp.WinHttpRequest
    With objWinHttp
        .Open "GET", strUrl, True
        .SetRequestHeader "Authorization", "Bearer " & AccessToken
        .SetRequestHeader "Accept", "application/json"
        .Send
        .WaitForResponse
        Dim colGroups As New Collection
        If .Status = 200 Then
            Dim objJson As Scripting.Dictionary
            Set objJson = JsonConverter.ParseJson(.ResponseText)
            Dim objGroup As Group
            Dim i As Long
            If objJson.Item("value").Count > 0 Then
                For i = 1 To objJson.Item("value").Count
                    Set objGroup = New Group
                    With objGroup
                        .Id = objJson.Item("value")(i)("id")
                        .Description = objJson.Item("value")(i)("description")
                        .DisplayName = objJson.Item("value")(i)("displayName")
                        .CreatedDateTime = objJson.Item("value")(i)("createdDateTime")
                        .Scope = objJson.Item("value")(i)("scope")
                    End With
                    colGroups.Add objGroup
                Next
                Set ListGroups = colGroups
            End If
        Else: Call HandleError(.ResponseText)
        End If
    End With
End Function

Public Function GetGroup(ByVal SiteId As String, ByVal GroupId As String, Optional ByVal ODataQuery As String) As Collection
    Dim strUrl As String
    strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/termStore/groups/" & GroupId
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
            Dim objJson As Scripting.Dictionary
            Set objJson = JsonConverter.ParseJson(.ResponseText)
            Dim objGroup As Group
            Set objGroup = New Group
            With objGroup
                .Id = objJson.Item("id")
                .Description = objJson.Item("description")
                .DisplayName = objJson.Item("displayName")
                .CreatedDateTime = objJson.Item("createdDateTime")
                .Scope = objJson.Item("scope")
            End With
            Set GetGroup = objGroup
        Else: Call HandleError(.ResponseText)
        End If
    End With
End Function

Public Sub DeleteGroup(ByVal SiteId As String, ByVal GroupId As String)
    Dim strUrl As String
    strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/termStore/groups/" & GroupId
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

Public Function CreateGroup(ByVal SiteId As String, ByVal Group As Group) As Group
    Dim strUrl As String
    strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/termStore/groups"
    Dim objWinHttp As WinHttp.WinHttpRequest
    Set objWinHttp = New WinHttp.WinHttpRequest
    With objWinHttp
        .Open "POST", strUrl, True
        .SetRequestHeader "Authorization", "Bearer " & AccessToken
        .SetRequestHeader "Accept", "application/json"
        .SetRequestHeader "Content-Type", "application/json"
        .Send JsonConverter.ConvertToJson(Group.ToJson)
        .WaitForResponse
        If .Status = 201 Then
            Dim objJson As Scripting.Dictionary
            Set objJson = JsonConverter.ParseJson(.ResponseText)
            Dim objGroup As Group
            Set objGroup = New Group
            With objGroup
                .Id = objJson.Item("id")
                .Description = objJson.Item("description")
                .DisplayName = objJson.Item("displayName")
                .CreatedDateTime = objJson.Item("createdDateTime")
                .Scope = objJson.Item("scope")
            End With
            Set CreateGroup = objGroup
        Else: Call HandleError(.ResponseText)
        End If
    End With
End Function

Public Function ListSetItems(ByVal SiteId As String, ByVal GroupId As String, Optional ByVal ODataQuery As String) As Collection
    Dim strUrl As String
    strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/termStore/groups/" & GroupId & "/sets"
    If ODataQuery <> vbNullString Then strUrl = strUrl & "?" & ODataQuery
    Dim objWinHttp As WinHttp.WinHttpRequest
    Set objWinHttp = New WinHttp.WinHttpRequest
    With objWinHttp
        .Open "GET", strUrl, True
        .SetRequestHeader "Authorization", "Bearer " & AccessToken
        .SetRequestHeader "Accept", "application/json"
        .Send
        .WaitForResponse
        Dim colSetItems As New Collection
        If .Status = 200 Then
            Dim objJson As Scripting.Dictionary
            Set objJson = JsonConverter.ParseJson(.ResponseText)
            Dim i As Long, j As Long
            Dim objLocalizedName As LocalizedName
            Dim colLocalizedNames As New Collection
            Dim objSetItem As SetItem
            If objJson.Item("value").Count > 0 Then
                For i = 1 To objJson.Item("value").Count
                    Set objSetItem = New SetItem
                    With objSetItem
                        .Id = objJson.Item("value")(i)("id")
                        .Description = objJson.Item("value")(i)("description")
                        If objJson.Item("value")(i).Exists("localizedNames") And objJson.Item("value")(i)("localizedNames").Count > 0 Then
                            For j = 1 To objJson.Item("value")(i)("localizedNames").Count
                                Set objLocalizedName = New LocalizedName
                                objLocalizedName.LanguageTag = objJson.Item("value")(i)("localizedNames")(j)("languageTag")
                                objLocalizedName.Name = objJson.Item("value")(i)("localizedNames")(j)("name")
                                colLocalizedNames.Add objLocalizedName
                            Next
                            Set .LocalizedNames = colLocalizedNames
                        End If
                    End With
                    colSetItems.Add objSetItem
                Next
            End If
            Set ListSetItems = colSetItems
        End If
    End With
End Function

Public Function ListRelations(ByVal SiteId As String, Location As RelationLocation, ByVal SetId As String, Optional ByVal TermId As String, Optional ByVal ODataQuery As String) As Collection
    Dim strUrl As String
    Select Case Location
        Case SetItem
            strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/termStore/sets/" & SetId & "/relations"
        Case Term
            strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/termStore/sets/" & SetId & "/terms/" & TermId & "/relations"
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
        Dim colRelations As New Collection
        If .Status = 200 Then
            Dim objJson As Scripting.Dictionary
            Set objJson = JsonConverter.ParseJson(.ResponseText)
            Dim i As Long, j As Long
            Dim objRelation As Relation
            If objJson.Item("value").Count > 0 Then
                For i = 1 To objJson.Item("value").Count
                    Set objRelation = New Relation
                    With objRelation
                        .Id = objJson.Item("value")(i)("id")
                        .Relationship = objJson.Item("value")(i)("relationship")
                    End With
                    colRelations.Add objRelation
                Next
            End If
            Set ListRelations = colRelations
        End If
    End With
End Function

Private Sub HandleError(ByVal Source As String)
    Dim objJson As Scripting.Dictionary
    Set objJson = JsonConverter.ParseJson(Source)
    Err.Raise vbObjectError, objJson.Item("error")("code"), objJson.Item("error")("message")
End Sub
