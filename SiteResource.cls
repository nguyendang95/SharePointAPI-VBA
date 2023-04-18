VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SiteResource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum PredefinedTimeRange
    AllTime
    LastSevenDays
End Enum

Public Enum UserType
    UserTypeCurrentUser
    UserTypeOtherUser
End Enum

Private m_AccessToken As String

Friend Property Get AccessToken() As String
   AccessToken = m_AccessToken
End Property

Friend Property Let AccessToken(Value As String)
   m_AccessToken = Value
End Property

Public Function GetApplicableContentTypesForList(ByVal SiteId As String, ByVal ListId As String, Optional ByVal ODataQuery As String) As Collection
    Dim strUrl As String
    strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/getApplicableContentTypesForList(listId='" & ListId & "')"
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
            End If
        Else: Call HandleError(.ResponseText)
        End If
    End With
    Set GetApplicableContentTypesForList = colContentTypes
End Function

Public Function CreateContentType(ByVal SiteId As String, ByVal ContentType As ContentType) As ContentType
    Dim objWinHttp As WinHttp.WinHttpRequest
    Set objWinHttp = New WinHttp.WinHttpRequest
    With objWinHttp
        .Open "POST", "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/contentTypes", True
        .SetRequestHeader "Authorization", "Bearer " & AccessToken
        .SetRequestHeader "Accept", "application/json"
        .SetRequestHeader "Content-Type", "application/json"
        .Send JsonConverter.ConvertToJson(ContentType.ToJson)
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
                Set CreateContentType = objContentType
            End With
        Else: Call HandleError(.ResponseText)
        End If
    End With
End Function

Public Function ListContentTypes(ByVal SiteId As String, Optional ByVal ODataQuery As String) As Collection
    Dim strUrl As String
    strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/contentTypes"
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
            End If
        Else: Call HandleError(.ResponseText)
        End If
    End With
    Set ListContentTypes = colContentTypes
End Function

Public Function CreateColumn(ByVal SiteId As String, ByVal ColumnDefinition As ColumnDefinition) As ColumnDefinition
    Dim objWinHttp As WinHttp.WinHttpRequest
    Set objWinHttp = New WinHttp.WinHttpRequest
    With objWinHttp
        .Open "POST", "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/columns", True
        .SetRequestHeader "Authorization", "Bearer " & AccessToken
        .SetRequestHeader "Accept", "application/json"
        .SetRequestHeader "Content-Type", "application/json"
        .Send JsonConverter.ConvertToJson(ColumnDefinition.ToJson)
        .WaitForResponse
        If .Status = 201 Then
            Dim objJson As Scripting.Dictionary
            Set objJson = JsonConverter.ParseJson(.ResponseText)
            Dim objTextColumn As TextColumn
            Dim objCalculatedColumn As CalculatedColumn
            Dim objChoiceColumn As ChoiceColumn
            Dim objCurrencyColumn As CurrencyColumn
            Dim objDateTimeColumn As DateTimeColumn
            Dim objPersonOrGroup As PersonOrGroupColumn
            Dim objTermColumn As TermColumn
            Dim objDefaultValuesColumns As DefaultColumnValues
            Dim objLookupColumn As LookupColumn
            Dim objColumnValidation As ColumnValidation
            Dim objNumberColumn As NumberColumn
            Dim objHyperlinkOrPictureColumn As HyperlinkOrPictureColumn
            Dim objColumnDefinition As ColumnDefinition
            Set objColumnDefinition = New ColumnDefinition
            With objColumnDefinition
                .ColumnGroup = objJson.Item("columnGroup")
                .EnforceUniqueValues = CBool(objJson.Item("enforceUniqueValues"))
                .Description = objJson.Item("description")
                .DisplayName = objJson.Item("displayName")
                .Hidden = CBool(objJson.Item("hidden"))
                .Id = objJson.Item("id")
                .Indexed = CBool(objJson.Item("indexed"))
                .Name = objJson.Item("name")
                .ReadOnly = CBool(objJson.Item("readOnly"))
                .Required = CBool(objJson.Item("required"))
                If objJson.Exists("text") Then
                    Set objTextColumn = New TextColumn
                    With objTextColumn
                        .AllowMultipleLines = CBool(objJson.Item("text")("allowMultipleLines"))
                        .AppendChangesToExistingText = CBool(objJson.Item("text")("appendChangesToExistingText"))
                        .LinesForEditing = objJson.Item("text")("linesForEditing")
                        .MaxLength = objJson.Item("text")("maxLength")
                        .TextType = objJson.Item("text")("textType")
                    End With
                    Set .Text = objTextColumn
                ElseIf objJson.Exists("calculated") Then
                    Set objCalculatedColumn = New CalculatedColumn
                    With objCalculatedColumn
                        .Format = objJson.Item("calculated")("format")
                        .Formula = objJson.Item("calculated")("formula")
                        .OutputType = objJson.Item("calculated")("outputType")
                    End With
                    Set .Calculated = objCalculatedColumn
                ElseIf objJson.Exists("choice") Then
                    Set objChoiceColumn = New ChoiceColumn
                    With objChoiceColumn
                        .AllowTextEntry = CBool(objJson.Item("choice")("allowTextEntry"))
                        Set .Choices = objJson.Item("choice")("choices")
                        .DisplayAs = objJson.Item("choice")("displayAs")
                    End With
                    Set .Choice = objChoiceColumn
                ElseIf objJson.Exists("currency") Then
                    Set objCurrencyColumn = New CurrencyColumn
                    objCurrencyColumn.Locale = objJson.Item("currency")("locale")
                    Set .CurrencyColumn = objCurrencyColumn
                ElseIf objJson.Exists("dateTime") Then
                    Set objDateTimeColumn = New DateTimeColumn
                    objDateTimeColumn.DisplayAs = objJson.Item("dateTime")("displayAs")
                    objDateTimeColumn.Format = objJson.Item("dateTime")("format")
                    Set .DateTime = objDateTimeColumn
                ElseIf objJson.Exists("defaultValues") Then
                    Set objDefaultValuesColumns = New DefaultColumnValues
                    objDefaultValuesColumns.Formula = objJson.Item("defaultValues")("formula")
                    objDefaultValuesColumns.Value = objJson.Item("defaultValues")
                    Set .DefaultValues = objDefaultValuesColumns
                ElseIf objJson.Exists("hyperlinkOrPicture") Then
                    Set objHyperlinkOrPictureColumn = New HyperlinkOrPictureColumn
                    objHyperlinkOrPictureColumn.IsPicture = CBool(objJson.Item("hyperlinkOrPictureColumn")("isPicture"))
                    Set .HyperlinkOrPicture = objHyperlinkOrPictureColumn
                ElseIf objJson.Exists("lookup") Then
                    Set objLookupColumn = New LookupColumn
                    With objLookupColumn
                        .AllowMultipleValues = CBool(objJson.Item("lookup")("allowMultipleValues"))
                        .AllowUnlimitedLength = CBool(objJson.Item("lookup")("allowUnlimitedLength"))
                        .ColumnName = objJson.Item("lookup")("columnName")
                        .ListId = objJson.Item("lookup")("listId")
                        .PrimaryLookupColumnId = objJson.Item("lookup")("primaryLookupColumnId")
                    End With
                    Set .Lookup = objLookupColumn
                ElseIf objJson.Exists("number") Then
                    Set objNumberColumn = New NumberColumn
                    With objNumberColumn
                        .DecimalPlaces = objJson.Item("number")("decimalPlaces")
                        .DisplayAs = objJson.Item("number")("displayAs")
                        .Maximum = objJson.Item("number")("maximum")
                        .Minimum = objJson.Item("number")("minimum")
                    End With
                    Set .Number = objNumberColumn
                ElseIf objJson.Exists("personOrGroup") Then
                    Set objPersonOrGroup = New PersonOrGroupColumn
                    With objPersonOrGroup
                        .AllowMultipleSelection = CBool(objJson.Item("personOrGroup")("allowMultipleSelection"))
                        .ChooseFromType = objJson.Item("personOrGroup")("chooseFromType")
                        .DisplayAs = objJson.Item("personOrGroup")("displayAs")
                    End With
                    Set .PersonOrGroup = objPersonOrGroup
                ElseIf objJson.Exists("term") Then
                    Set objTermColumn = New TermColumn
                    objTermColumn.AllowMultipleValues = CBool(objJson.Item("term")("allowMultipleValues"))
                    objTermColumn.ShowFullyQualifiedName = CBool(objJson.Item("term")("showFullyQualifiedName"))
                    Set .Term = objTermColumn
                ElseIf objJson.Exists("validation") Then
                    Set objColumnValidation = New ColumnValidation
                    With objColumnValidation
                        .DefaultLanguage = objJson.Item("validation")("defaultLanguage")
                        Set .Descriptions = objJson.Item("validation")("descriptions")
                        .Formula = objJson.Item("validation")("formula")
                    End With
                    Set .Validation = objColumnValidation
                End If
                Set CreateColumn = objColumnDefinition
            End With
        Else: Call HandleError(.ResponseText)
        End If
    End With
End Function

Public Function ListColumns(ByVal SiteId As String, Optional ByVal ODataQuery As String) As Collection
    Dim strUrl As String
    strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/columns"
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
        Dim colColumnDefinitions As New Collection
        If .Status = 200 Then
            Dim objJson As Scripting.Dictionary
            Set objJson = JsonConverter.ParseJson(.ResponseText)
            Dim i As Long
            Dim objColumnDefinition As ColumnDefinition
            Dim objTextColumn As TextColumn
            Dim objCalculatedColumn As CalculatedColumn
            Dim objChoiceColumn As ChoiceColumn
            Dim objCurrencyColumn As CurrencyColumn
            Dim objDateTimeColumn As DateTimeColumn
            Dim objPersonOrGroup As PersonOrGroupColumn
            Dim objTermColumn As TermColumn
            Dim objDefaultValuesColumns As DefaultColumnValues
            Dim objLookupColumn As LookupColumn
            Dim objColumnValidation As ColumnValidation
            Dim objNumberColumn As NumberColumn
            Dim objHyperlinkOrPictureColumn As HyperlinkOrPictureColumn
            If objJson.Item("value").Count > 0 Then
                For i = 1 To objJson.Item("value").Count
                    Set objColumnDefinition = New ColumnDefinition
                    With objColumnDefinition
                        .ColumnGroup = objJson.Item("value")(i)("columnGroup")
                        .EnforceUniqueValues = CBool(objJson.Item("value")(i)("enforceUniqueValues"))
                        .Description = objJson.Item("value")(i)("description")
                        .DisplayName = objJson.Item("value")(i)("displayName")
                        .Hidden = CBool(objJson.Item("value")(i)("hidden"))
                        .Id = objJson.Item("value")(i)("id")
                        .Indexed = CBool(objJson.Item("value")(i)("indexed"))
                        .Name = objJson.Item("value")(i)("name")
                        .ReadOnly = CBool(objJson.Item("value")(i)("readOnly"))
                        .Required = CBool(objJson.Item("value")(i)("required"))
                        If objJson.Item("value")(i).Exists("text") Then
                            Set objTextColumn = New TextColumn
                            With objTextColumn
                                .AllowMultipleLines = CBool(objJson.Item("value")(i)("text")("allowMultipleLines"))
                                .AppendChangesToExistingText = CBool(objJson.Item("value")(i)("text")("appendChangesToExistingText"))
                                .LinesForEditing = objJson.Item("value")(i)("text")("linesForEditing")
                                .MaxLength = objJson.Item("value")(i)("text")("maxLength")
                                .TextType = objJson.Item("value")(i)("text")("textType")
                            End With
                            Set .Text = objTextColumn
                        ElseIf objJson.Item("value")(i).Exists("calculated") Then
                            Set objCalculatedColumn = New CalculatedColumn
                            With objCalculatedColumn
                                .Format = objJson.Item("value")(i)("calculated")("format")
                                .Formula = objJson.Item("value")(i)("calculated")("formula")
                                .OutputType = objJson.Item("value")(i)("calculated")("outputType")
                            End With
                            Set .Calculated = objCalculatedColumn
                        ElseIf objJson.Item("value")(i).Exists("choice") Then
                            Set objChoiceColumn = New ChoiceColumn
                            With objChoiceColumn
                                .AllowTextEntry = CBool(objJson.Item("value")(i)("choice")("allowTextEntry"))
                                Set .Choices = objJson.Item("value")(i)("choice")("choices")
                                .DisplayAs = objJson.Item("value")(i)("choice")("displayAs")
                            End With
                            Set .Choice = objChoiceColumn
                        ElseIf objJson.Item("value")(i).Exists("currency") Then
                            Set objCurrencyColumn = New CurrencyColumn
                            objCurrencyColumn.Locale = objJson.Item("value")(i)("currency")("locale")
                            Set .CurrencyColumn = objCurrencyColumn
                        ElseIf objJson.Item("value")(i).Exists("dateTime") Then
                            Set objDateTimeColumn = New DateTimeColumn
                            objDateTimeColumn.DisplayAs = objJson.Item("value")(i)("dateTime")("displayAs")
                            objDateTimeColumn.Format = objJson.Item("value")(i)("dateTime")("format")
                            Set .DateTime = objDateTimeColumn
                        ElseIf objJson.Item("value")(i).Exists("defaultValues") Then
                            Set objDefaultValuesColumns = New DefaultColumnValues
                            objDefaultValuesColumns.Formula = objJson.Item("value")(i)("defaultValues")("formula")
                            objDefaultValuesColumns.Value = objJson.Item("value")(i)("defaultValues")("value")
                            Set .DefaultValues = objDefaultValuesColumns
                        ElseIf objJson.Item("value")(i).Exists("hyperlinkOrPicture") Then
                            Set objHyperlinkOrPictureColumn = New HyperlinkOrPictureColumn
                            objHyperlinkOrPictureColumn.IsPicture = CBool(objJson.Item("value")(i)("hyperlinkOrPictureColumn")("isPicture"))
                            Set .HyperlinkOrPicture = objHyperlinkOrPictureColumn
                        ElseIf objJson.Item("value")(i).Exists("lookup") Then
                            Set objLookupColumn = New LookupColumn
                            With objLookupColumn
                                .AllowMultipleValues = CBool(objJson.Item("value")(i)("lookup")("allowMultipleValues"))
                                .AllowUnlimitedLength = CBool(objJson.Item("value")(i)("lookup")("allowUnlimitedLength"))
                                .ColumnName = objJson.Item("value")(i)("lookup")("columnName")
                                .ListId = objJson.Item("value")(i)("lookup")("listId")
                                .PrimaryLookupColumnId = objJson.Item("value")(i)("lookup")("primaryLookupColumnId")
                            End With
                            Set .Lookup = objLookupColumn
                        ElseIf objJson.Item("value")(i).Exists("number") Then
                            Set objNumberColumn = New NumberColumn
                            With objNumberColumn
                                .DecimalPlaces = objJson.Item("value")(i)("number")("decimalPlaces")
                                .DisplayAs = objJson.Item("value")(i)("number")("displayAs")
                                .Maximum = objJson.Item("value")(i)("number")("maximum")
                                .Minimum = objJson.Item("value")(i)("number")("minimum")
                            End With
                            Set .Number = objNumberColumn
                        ElseIf objJson.Item("value")(i).Exists("personOrGroup") Then
                            Set objPersonOrGroup = New PersonOrGroupColumn
                            With objPersonOrGroup
                                .AllowMultipleSelection = CBool(objJson.Item("value")(i)("personOrGroup")("allowMultipleSelection"))
                                .ChooseFromType = objJson.Item("value")(i)("personOrGroup")("chooseFromType")
                                .DisplayAs = objJson.Item("value")(i)("personOrGroup")("displayAs")
                            End With
                            Set .PersonOrGroup = objPersonOrGroup
                        ElseIf objJson.Item("value")(i).Exists("term") Then
                            Set objTermColumn = New TermColumn
                            objTermColumn.AllowMultipleValues = CBool(objJson.Item("value")(i)("term")("allowMultipleValues"))
                            objTermColumn.ShowFullyQualifiedName = CBool(objJson.Item("value")(i)("term")("showFullyQualifiedName"))
                            Set .Term = objTermColumn
                        ElseIf objJson.Item("value")(i).Exists("validation") Then
                            Set objColumnValidation = New ColumnValidation
                            With objColumnValidation
                                .DefaultLanguage = objJson.Item("value")(i)("validation")("defaultLanguage")
                                Set .Descriptions = objJson.Item("value")(i)("validation")("descriptions")
                                .Formula = objJson.Item("value")(i)("validation")("formula")
                            End With
                            Set .Validation = objColumnValidation
                        End If
                    End With
                    colColumnDefinitions.Add objColumnDefinition
                Next
            End If
        Else: Call HandleError(.ResponseText)
        End If
    End With
    Set ListColumns = colColumnDefinitions
End Function

Public Function FollowSites(ByVal UserId As String, ByVal SiteIds As ArrayOfSiteIds) As Collection
    Dim objWinHttp As WinHttp.WinHttpRequest
    Set objWinHttp = New WinHttp.WinHttpRequest
    With objWinHttp
        .Open "POST", "https://graph.microsoft.com/v1.0/users/" & UserId & "/followedSites/add", True
        .SetRequestHeader "Authorization", "Bearer " & AccessToken
        .SetRequestHeader "Accept", "application/json"
        .SetRequestHeader "Content-Type", "application/json"
        .Send JsonConverter.ConvertToJson(SiteIds.ToJson)
        .WaitForResponse
        Dim objJson As Scripting.Dictionary
        Dim i As Long
        Dim colSites As New Collection
        If .Status = 200 Or .Status = 207 Then
            Set objJson = JsonConverter.ParseJson(.ResponseText)
            If objJson.Item("value").Count > 0 Then
                Dim objSiteCollection As SiteCollection
                Dim objSharepointIds As SharepointIds
                Dim objErrorObject As ErrorObject
                Dim objInnerError As InnerError
                For i = 1 To objJson.Item("value").Count
                    If objJson.Item("value")(i).Exists("sharepointIds") Then
                        Set objSharepointIds = New SharepointIds
                        With objSharepointIds
                            .ListItemId = objJson.Item("value")(i)("sharepointIds")("listItemId")
                            .SiteId = objJson.Item("value")(i)("sharepointIds")("siteId")
                            .WebId = objJson.Item("value")(i)("sharepointIds")("webId")
                        End With
                        Set objSiteCollection = New SiteCollection
                        objSiteCollection.Hostname = objJson.Item("value")(i)("siteCollection")("hostname")
                        Dim objSite As Site
                        Set objSite = New Site
                        With objSite
                            .Id = objJson.Item("value")(i)("id")
                            .CreatedDateTime = objJson.Item("value")(i)("createdDateTime")
                            .DisplayName = objJson.Item("value")(i)("displayName")
                            .LastModifiedDateTime = objJson.Item("value")(i)("lastModifiedDateTime")
                            .Name = objJson.Item("value")(i)("name")
                            .WebUrl = objJson.Item("value")(i)("webUrl")
                            Set .SiteCollection = objSiteCollection
                            Set .SharepointIds = objSharepointIds
                        End With
                        colSites.Add objSite
                    ElseIf objJson.Item("value")(i).Exists("error") Then
                        Set objErrorObject = New ErrorObject
                        With objErrorObject
                            .Id = objJson.Item("value")(i)("id")
                            .Code = objJson.Item("value")(i)("error")("code")
                            .Message = objJson.Item("value")(i)("error")("message")
                            If objJson.Item("value")(i)("error").Exists("innerError") Then
                                Set objInnerError = New InnerError
                                With objInnerError
                                    .Code = objJson.Item("value")(i)("error")("innerError")("code")
                                    .ErrorType = objJson.Item("value")(i)("error")("innerError")("errorType")
                                    .Message = objJson.Item("value")(i)("error")("innerError")("message")
                                    .ThrowSite = objJson.Item("value")(i)("error")("innerError")("throwSite")
                                End With
                                Set .InnerError = objInnerError
                            End If
                        End With
                        colSites.Add objErrorObject
                    End If
                Next
            End If
        Else: Call HandleError(.ResponseText)
        End If
    End With
    Set FollowSites = colSites
End Function

Public Function UnfollowSites(ByVal UserId As String, ByVal SiteIds As ArrayOfSiteIds) As Collection
    Dim objWinHttp As WinHttp.WinHttpRequest
    Set objWinHttp = New WinHttp.WinHttpRequest
    With objWinHttp
        .Open "POST", "https://graph.microsoft.com/v1.0/users/" & UserId & "/followedSites/remove", True
        .SetRequestHeader "Authorization", "Bearer " & AccessToken
        .SetRequestHeader "Accept", "application/json"
        .SetRequestHeader "Content-Type", "application/json"
        .Send JsonConverter.ConvertToJson(SiteIds.ToJson)
        .WaitForResponse
        Dim i As Long
        Dim colSites As New Collection
        If .Status = 200 Or .Status = 207 Then
            Dim objJson As Scripting.Dictionary
            Set objJson = JsonConverter.ParseJson(.ResponseText)
            If objJson.Item("value").Count > 0 Then
                Dim objSiteCollection As SiteCollection
                Dim objSharepointIds As SharepointIds
                Dim objErrorObject As ErrorObject
                Dim objInnerError As InnerError
                For i = 1 To objJson.Item("value").Count
                    If objJson.Item("value")(i).Exists("sharepointIds") Then
                        Set objSharepointIds = New SharepointIds
                        With objSharepointIds
                            .ListItemId = objJson.Item("value")(i)("sharepointIds")("listItemId")
                            .SiteId = objJson.Item("value")(i)("sharepointIds")("siteId")
                            .WebId = objJson.Item("value")(i)("sharepointIds")("webId")
                        End With
                        Set objSiteCollection = New SiteCollection
                        objSiteCollection.Hostname = objJson.Item("value")(i)("siteCollection")("hostname")
                        Dim objSite As Site
                        Set objSite = New Site
                        With objSite
                            .Id = objJson.Item("value")(i)("id")
                            .CreatedDateTime = objJson.Item("value")(i)("createdDateTime")
                            .DisplayName = objJson.Item("value")(i)("displayName")
                            .LastModifiedDateTime = objJson.Item("value")(i)("lastModifiedDateTime")
                            .Name = objJson.Item("value")(i)("name")
                            .WebUrl = objJson.Item("value")(i)("webUrl")
                            Set .SiteCollection = objSiteCollection
                            Set .SharepointIds = objSharepointIds
                        End With
                        colSites.Add objSite
                    ElseIf objJson.Item("value")(i).Exists("error") Then
                        Set objErrorObject = New ErrorObject
                        With objErrorObject
                            .Id = objJson.Item("value")(i)("id")
                            .Code = objJson.Item("value")(i)("error")("code")
                            .Message = objJson.Item("value")(i)("error")("message")
                            If objJson.Item("value")(i)("error").Exists("innerError") Then
                                Set objInnerError = New InnerError
                                With objInnerError
                                    .Code = objJson.Item("value")(i)("error")("innerError")("code")
                                    .ErrorType = objJson.Item("value")(i)("error")("innerError")("errorType")
                                    .Message = objJson.Item("value")(i)("error")("innerError")("message")
                                    .ThrowSite = objJson.Item("value")(i)("error")("innerError")("throwSite")
                                End With
                                Set .InnerError = objInnerError
                            End If
                        End With
                        colSites.Add objErrorObject
                    End If
                Next
            End If
        Else: Call HandleError(.ResponseText)
        End If
    End With
    Set UnfollowSites = colSites
End Function

Public Function ListFollowedSites(Optional ByVal ODataQuery As String = vbNullString) As Collection
    Dim strUrl As String
    strUrl = "https://graph.microsoft.com/v1.0/me/followedSites"
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
        Dim colSites As New Collection
        If .Status = 200 Then
            Dim objJson As Scripting.Dictionary
            Set objJson = JsonConverter.ParseJson(.ResponseText)
            If objJson.Item("value").Count > 0 Then
                Dim i As Long
                Dim objSiteCollection As SiteCollection
                Dim objSharepointIds As SharepointIds
                For i = 1 To objJson.Item("value").Count
                    If objJson.Item("value")(i).Exists("sharepointIds") Then
                        Set objSharepointIds = New SharepointIds
                        With objSharepointIds
                            .ListItemId = objJson.Item("value")(i)("sharepointIds")("listItemId")
                            .SiteId = objJson.Item("value")(i)("sharepointIds")("siteId")
                            .WebId = objJson.Item("value")(i)("sharepointIds")("webId")
                        End With
                    End If
                    If objJson.Item("value")(i).Exists("siteCollection") Then
                        Set objSiteCollection = New SiteCollection
                        objSiteCollection.Hostname = objJson.Item("value")(i)("siteCollection")("hostname")
                    End If
                    Dim objSite As Site
                    Set objSite = New Site
                    With objSite
                        .Id = objJson.Item("value")(i)("id")
                        .CreatedDateTime = objJson.Item("value")(i)("createdDateTime")
                        .DisplayName = objJson.Item("value")(i)("displayName")
                        .LastModifiedDateTime = objJson.Item("value")(i)("lastModifiedDateTime")
                        .Name = objJson.Item("value")(i)("name")
                        .WebUrl = objJson.Item("value")(i)("webUrl")
                        If Not objSiteCollection Is Nothing Then Set .SiteCollection = objSiteCollection
                        If Not objSharepointIds Is Nothing Then Set .SharepointIds = objSharepointIds
                    End With
                    colSites.Add objSite
                Next
            End If
        Else: Call HandleError(.ResponseText)
        End If
    End With
    Set ListFollowedSites = colSites
End Function

Public Function GetUser(GetUserFrom As UserType, Optional ByVal IdOrUserPrincipalName As String, Optional ByVal ODataQuery As String) As User
    Dim strUrl As String
    Dim objWebUtilities As WebUtilities
    Set objWebUtilities = New WebUtilities
    strUrl = "https://graph.microsoft.com/v1.0/"
    If GetUserFrom = UserTypeCurrentUser Then
        strUrl = strUrl & "me"
    Else
        If Left(IdOrUserPrincipalName, 1) = "$" Then
            strUrl = strUrl & "/users('" & objWebUtilities.URLEncode(IdOrUserPrincipalName) & "')"
        Else: strUrl = strUrl & "/users/" & objWebUtilities.URLEncode(IdOrUserPrincipalName)
        End If
    End If
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
            Dim objUser As User
            Set objUser = New User
            With objUser
                If objJson.Exists("businessPhones") Then
                    If objJson.Item("businessPhones").Count > 0 Then
                        Dim colBusinessPhones As New Collection
                        Dim i As Long
                        For i = 1 To objJson.Item("businessPhones").Count
                            colBusinessPhones.Add objJson.Item("businessPhones")(i)
                        Next
                        Set .BusinessPhones = colBusinessPhones
                    End If
                End If
                If IsNull(objJson.Item("displayName")) = False Then .DisplayName = objJson.Item("displayName")
                If IsNull(objJson.Item("givenName")) = False Then .GivenName = objJson.Item("givenName")
                If IsNull(objJson.Item("jobTitle")) = False Then .JobTitle = objJson.Item("jobTitle")
                If IsNull(objJson.Item("mail")) = False Then .Mail = objJson.Item("mail")
                If IsNull(objJson.Item("mobilePhone")) = False Then .MobilePhone = objJson.Item("mobilePhone")
                If IsNull(objJson.Item("officeLocation")) = False Then .OfficeLocation = objJson.Item("officeLocation")
                If IsNull(objJson.Item("preferredLanguage")) = False Then .PreferredLanguage = objJson.Item("preferredLanguage")
                If IsNull(objJson.Item("surName")) = False Then .Surname = objJson.Item("surName")
                If IsNull(objJson.Item("userPrincipal")) = False Then .UserPrincipalName = objJson.Item("userPrincipalName")
                If IsNull(objJson.Item("id")) = False Then .Id = objJson.Item("id")
            End With
            Set GetUser = objUser
        Else: Call HandleError(.ResponseText)
        End If
    End With
End Function

Public Function GetSiteByServerRelativePath(ByVal Hostname As String, ByVal ServerRelativePath As String, Optional ByVal ODataQuery As String = vbNullString) As Site
    Dim objWinHttp As WinHttp.WinHttpRequest
    Dim objJson As Scripting.Dictionary
    Set objWinHttp = New WinHttp.WinHttpRequest
    With objWinHttp
        If ODataQuery <> vbNullString Then
            .Open "GET", "https://graph.microsoft.com/v1.0/sites/" & Hostname & ":/" & ServerRelativePath & "?" & ODataQuery, True
        Else: .Open "GET", "https://graph.microsoft.com/v1.0/sites/" & Hostname & ":/" & ServerRelativePath, True
        End If
        .SetRequestHeader "Accept", "application/json"
        .SetRequestHeader "Authorization", "Bearer " & AccessToken
        .Send
        .WaitForResponse
        If .Status = 200 Then
            Set objJson = JsonConverter.ParseJson(.ResponseText)
            Dim objSite As Site
            Set objSite = New Site
            With objSite
                .Id = objJson.Item("id")
                .CreatedDateTime = objJson.Item("createdDateTime")
                .DisplayName = objJson.Item("displayName")
                .LastModifiedDateTime = objJson.Item("lastModifiedDateTime")
                .Name = objJson.Item("name")
                .WebUrl = objJson.Item("webUrl")
            End With
            Set GetSiteByServerRelativePath = objSite
        End If
        If .Status <> 200 Then Call HandleError(.ResponseText)
    End With
End Function

Public Function GetGroupTeamSite(ByVal GroupId As String, Optional ByVal ODataQuery As String = vbNullString) As Site
    Dim objWinHttp As WinHttp.WinHttpRequest
    Dim objJson As Scripting.Dictionary
    Set objWinHttp = New WinHttp.WinHttpRequest
    With objWinHttp
        If ODataQuery <> vbNullString Then
            .Open "GET", "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/sites/root?" & ODataQuery, True
        Else: .Open "GET", "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/sites/root", True
        End If
        .SetRequestHeader "Accept", "application/json"
        .SetRequestHeader "Authorization", "Bearer " & AccessToken
        .Send
        .WaitForResponse
        If .Status = 200 Then
            Set objJson = JsonConverter.ParseJson(.ResponseText)
            Dim objSite As Site
            Set objSite = New Site
            With objSite
                .Id = objJson.Item("id")
                .CreatedDateTime = objJson.Item("createdDateTime")
                .DisplayName = objJson.Item("displayName")
                .LastModifiedDateTime = objJson.Item("lastModifiedDateTime")
                .Name = objJson.Item("name")
                .WebUrl = objJson.Item("webUrl")
            End With
            Set GetGroupTeamSite = objSite
        End If
        If .Status <> 200 Then Call HandleError(.ResponseText)
    End With
End Function

Public Function GetRootSite(ByVal SiteId As String, Optional ByVal ODataQuery As String = vbNullString) As Site
    Dim objWinHttp As WinHttp.WinHttpRequest
    Dim objJson As Scripting.Dictionary
    Set objWinHttp = New WinHttp.WinHttpRequest
    With objWinHttp
        If ODataQuery <> vbNullString Then
            .Open "GET", "https://graph.microsoft.com/v1.0/sites/" & SiteId & "?" & ODataQuery, True
        Else
            .Open "GET", "https://graph.microsoft.com/v1.0/sites/" & SiteId, True
        End If
        .SetRequestHeader "Accept", "application/json"
        .SetRequestHeader "Authorization", "Bearer " & AccessToken
        .Send
        .WaitForResponse
        If .Status = 200 Then
            Set objJson = JsonConverter.ParseJson(.ResponseText)
            Dim objSiteCollection As SiteCollection
            Dim objSite As Site
            Set objSite = New Site
            With objSite
                .Id = objJson.Item("id")
                .CreatedDateTime = objJson.Item("createdDateTime")
                .DisplayName = objJson.Item("displayName")
                .LastModifiedDateTime = objJson.Item("lastModifiedDateTime")
                .Name = objJson.Item("name")
                .WebUrl = objJson.Item("webUrl")
                If objJson.Exists("siteCollection") Then
                    Set objSiteCollection = New SiteCollection
                    objSiteCollection.Hostname = objJson.Item("siteCollection")("hostname")
                    Set .SiteCollection = objSiteCollection
                End If
            End With
            Set GetRootSite = objSite
        End If
        If .Status <> 200 Then Call HandleError(.ResponseText)
    End With
End Function

Private Sub HandleError(ByVal Source As String)
    Dim objJson As Scripting.Dictionary
    Set objJson = JsonConverter.ParseJson(Source)
    Err.Raise vbObjectError, objJson.Item("error")("code"), objJson.Item("error")("message")
End Sub

Public Function ListSubsitesOfRootSite(Optional ByVal ODataQuery As String = vbNullString) As Collection
    Dim objWebUtilities As WebUtilities
    Set objWebUtilities = New WebUtilities
    Dim objWinHttp As WinHttp.WinHttpRequest
    Set objWinHttp = New WinHttp.WinHttpRequest
    With objWinHttp
        If ODataQuery <> vbNullString Then
            .Open "GET", "https://graph.microsoft.com/v1.0/sites/root/sites?" & ODataQuery, True
        Else: .Open "GET", "https://graph.microsoft.com/v1.0/sites", True
        End If
        .SetRequestHeader "Authorization", "Bearer " & AccessToken
        .SetRequestHeader "Accept", "application/json"
        .Send
        .WaitForResponse
        Dim colSites As New Collection
        If .Status = 200 Then
            Dim objJson As Scripting.Dictionary
            Set objJson = JsonConverter.ParseJson(.ResponseText)
            Dim i As Long
            Dim objSite As Site
            Dim objSiteCollection As SiteCollection
            Dim objRoot As Root
            If objJson.Item("value").Count > 0 Then
                For i = 1 To objJson.Item("value").Count
                    Set objSite = New Site
                    With objSite
                        .Id = objJson.Item("value")(i)("id")
                        .Name = objJson.Item("value")(i)("name")
                        Set objRoot = New Root
                        Set .Root = objRoot
                        If objJson.Item("value")(i).Exists("siteCollection") Then
                            Set objSiteCollection = New SiteCollection
                            With objSiteCollection
                                .Hostname = objJson.Item("value")(i)("siteCollection")("hostname")
                                .DataLocationCode = objJson.Item("value")(i)("siteCollection")("dataLocationCode")
                                Set .Root = objRoot
                            End With
                        End If
                        .WebUrl = objJson.Item("value")(i)("webUrl")
                    End With
                    colSites.Add objSite
                Next
            End If
        Else: Call HandleError(.ResponseText)
        End If
    End With
    Set ListSubsitesOfRootSite = colSites
End Function

Public Function SearchForSites(ByVal SearchQuery As String) As Collection
    Dim objWinHttp As WinHttp.WinHttpRequest
    Set objWinHttp = New WinHttp.WinHttpRequest
    With objWinHttp
        .Open "GET", "https://graph.microsoft.com/v1.0/sites?search=" & SearchQuery
        .SetRequestHeader "Accept", "application/json"
        .Send
        .WaitForResponse
        Dim colSites As New Collection
        If .Status = 200 Then
            Dim objJson As Scripting.Dictionary
            Set objJson = JsonConverter.ParseJson(.ResponseText)
            Dim i As Long
            Dim objSite As Site
            Dim objSiteCollection As SiteCollection
            Dim objRoot As Root
            If objJson.Item("value").Count > 0 Then
                For i = 1 To objJson.Item("value").Count
                    Set objSite = New Site
                    With objSite
                        .Id = objJson.Item("value")(i)("id")
                        .Name = objJson.Item("value")(i)("name")
                        Set objRoot = New Root
                        Set .Root = objRoot
                        If objJson.Item("value")(i).Exists("siteCollection") Then
                            Set objSiteCollection = New SiteCollection
                            With objSiteCollection
                                .Hostname = objJson.Item("value")(i)("siteCollection")("hostname")
                                .DataLocationCode = objJson.Item("value")(i)("siteCollection")("dataLocationCode")
                                Set .Root = objRoot
                            End With
                        End If
                        .WebUrl = objJson.Item("value")(i)("webUrl")
                    End With
                    colSites.Add objSite
                Next
            End If
        Else: Call HandleError(.ResponseText)
        End If
    End With
    Set SearchForSites = colSites
End Function

Public Function ListOperations(ByVal SiteId As String, Optional ByVal ODataQuery As String) As Collection
    Dim strUrl As String
    strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/operations"
    If ODataQuery <> vbNullString Then strUrl = strUrl & "?" & ODataQuery
    Dim objWinHttp As WinHttp.WinHttpRequest
    Set objWinHttp = New WinHttp.WinHttpRequest
    With objWinHttp
        .Open "GET", strUrl, True
        .SetRequestHeader "Authorization", "Bearer " & AccessToken
        .SetRequestHeader "Accept", "application/json"
        .Send
        .WaitForResponse
        Dim colRichLongRunningOperations As New Collection
        If .Status = 200 Then
            Dim objJson As Scripting.Dictionary
            Set objJson = JsonConverter.ParseJson(.ResponseText)
            Dim i As Long
            Dim objRichLongRunningOperation As RichLongRunningOperation
            If objJson.Item("value").Count > 0 Then
                For i = 1 To objJson.Item("value").Count
                    Set objRichLongRunningOperation = New RichLongRunningOperation
                    With objRichLongRunningOperation
                        .Id = objJson.Item("value")(i)("id")
                        .CreatedDateTime = objJson.Item("value")(i)("createdDateTime")
                        .ResourceId = objJson.Item("value")(i)("resourceId")
                        .ResourceLocation = objJson.Item("value")(i)("resourceLocation")
                        .Status = objJson.Item("value")(i)("status")
                        .RichLongRunningOperationType = objJson.Item("value")(i)("type")
                    End With
                    colRichLongRunningOperations.Add objRichLongRunningOperation
                Next
            End If
        Else: Call HandleError(.ResponseText)
        End If
    End With
    Set ListOperations = colRichLongRunningOperations
End Function
