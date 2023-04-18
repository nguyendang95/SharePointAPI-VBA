VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ListResource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum GetListType
    ById
    ByTitle
    ByFields
End Enum

Private m_AccessToken As String

Friend Property Get AccessToken() As String
    AccessToken = m_AccessToken
End Property

Friend Property Let AccessToken(Value As String)
    m_AccessToken = Value
End Property

Public Function Create(ByVal SiteId As String, List As List) As List
    Dim objWinHttp As WinHttp.WinHttpRequest
    Set objWinHttp = New WinHttp.WinHttpRequest
    With objWinHttp
        .Open "POST", "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/lists", True
        .SetRequestHeader "Accept", "application/json"
        .SetRequestHeader "Authorization", "Bearer " & AccessToken
        .SetRequestHeader "Content-Type", "application/json"
        .Send JsonConverter.ConvertToJson(List.ToJson)
        .WaitForResponse
        If .Status = 201 Then
            Dim objJson As Scripting.Dictionary
            Set objJson = JsonConverter.ParseJson(.ResponseText)
            Dim objCreatedBy As IdentitySet
            Dim objIdentity As Identity
            Dim objItemReference As ItemReference
            Dim objListInfo As ListInfo
            Dim objList As List
            Set objList = New List
            With objList
                .Id = objJson.Item("id")
                .CreatedDateTime = objJson.Item("createdDateTime")
                .WebUrl = objJson.Item("webUrl")
                .ETag = objJson.Item("eTag")
                .Name = objJson.Item("name")
                .DisplayName = objJson.Item("displayName")
                Set objCreatedBy = New IdentitySet
                Set objIdentity = New Identity
                If objJson.Exists("createdBy") Then
                    If objJson.Item("createdBy").Exists("user") And IsEmpty(objJson.Item("createdBy").Exists("user")) = False Then
                        objIdentity.DisplayName = objJson.Item("createdBy")("user")("displayName")
                        objIdentity.Id = objJson.Item("createdBy")("user")("id")
                        Set objCreatedBy.User = objIdentity
                    End If
                    If objJson.Item("createdBy").Exists("application") And IsEmpty(objJson.Item("createdBy").Exists("application")) = False Then
                        objIdentity.DisplayName = objJson.Item("createdBy")("application")("displayName")
                        objIdentity.Id = objJson.Item("createdBy")("application")("id")
                        Set objCreatedBy.Application = objIdentity
                    End If
                    If objJson.Item("createdBy").Exists("device") And IsEmpty(objJson.Item("createdBy").Exists("device")) = False Then
                        objIdentity.DisplayName = objJson.Item("createdBy")("device")("displayName")
                        objIdentity.Id = objJson.Item("createdBy")("device")("id")
                        Set objCreatedBy.Device = objIdentity
                    End If
                End If
                Set .CreatedBy = objCreatedBy
                .LastModifiedDateTime = objJson.Item("lastModifiedDateTime")
                If objJson.Exists("parentReference") Then
                    Set objItemReference = New ItemReference
                    With objItemReference
                        .SiteId = objJson.Item("parentReference")("siteId")
                    End With
                    Set .ParentReference = objItemReference
                End If
                If objJson.Exists("list") Then
                    Set objListInfo = New ListInfo
                    With objListInfo
                        .ContentTypesEnabled = CBool(objJson.Item("list")("contentTypesEnabled"))
                        .Hidden = CBool(objJson.Item("list")("hidden"))
                        .Template = objJson.Item("list")("template")
                    End With
                    Set .List = objListInfo
                End If
            End With
            Set Create = objList
        Else: Call HandleError(.ResponseText)
        End If
    End With
End Function

Public Function GetList(ByVal SiteId As String, Optional ByVal ListType As GetListType = ByTitle, Optional ByVal ListId As String, Optional ByVal ListTitle As String, Optional ByVal ODataQuery As String) As List
    Dim strUrl As String
    strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/lists/"
    Dim objWebUtilities As WebUtilities
    Set objWebUtilities = New WebUtilities
    Select Case ListType
        Case ById
            strUrl = strUrl & ListId
        Case ByTitle
            strUrl = strUrl & objWebUtilities.URLEncode(ListTitle)
    End Select
    If ODataQuery <> vbNullString And ListType = ById Then strUrl = strUrl & "?" & ODataQuery
    Dim objWinHttp As WinHttp.WinHttpRequest
    Set objWinHttp = New WinHttp.WinHttpRequest
    With objWinHttp
        .Open "GET", strUrl, True
        .SetRequestHeader "Accept", "application/json"
        .SetRequestHeader "Authorization", "Bearer " & AccessToken
        .Send
        .WaitForResponse
        Dim colLists As New Collection
        If .Status = 200 Then
            Dim objJson As Scripting.Dictionary
            Set objJson = JsonConverter.ParseJson(.ResponseText)
            Dim objListInfo As ListInfo
            Dim objCreatedBy As IdentitySet
            Dim objIdentity As Identity
            Dim objList As List
            Set objList = New List
            With objList
                .Id = objJson.Item("id")
                .CreatedDateTime = objJson.Item("createdDateTime")
                .WebUrl = objJson.Item("webUrl")
                .ETag = objJson.Item("eTag")
                .Name = objJson.Item("name")
                .DisplayName = objJson.Item("displayName")
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
                If objJson.Exists("parentReference") Then
                    Dim objItemReference As ItemReference
                    Set objItemReference = New ItemReference
                    With objItemReference
                        .SiteId = objJson.Item("parentReference")("siteId")
                    End With
                    Set .ParentReference = objItemReference
                End If
                If objJson.Exists("list") Then
                    Set objListInfo = New ListInfo
                    With objListInfo
                        .ContentTypesEnabled = CBool(objJson.Item("list")("contentTypesEnabled"))
                        .Hidden = CBool(objJson.Item("list")("hidden"))
                        .Template = objJson.Item("list")("template")
                    End With
                    Set .List = objListInfo
                End If
                colLists.Add objList
            End With
        Else: Call HandleError(.ResponseText)
        End If
    End With
    Set GetList = objList
End Function

Public Function List(ByVal SiteId As String, Optional ByVal ODataQuery As String) As Collection
    Dim strUrl As String
    strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/lists"
    If ODataQuery <> vbNullString Then strUrl = strUrl & "?" & ODataQuery
    Dim objWinHttp As WinHttp.WinHttpRequest
    Set objWinHttp = New WinHttp.WinHttpRequest
    With objWinHttp
        .Open "GET", strUrl, True
        .SetRequestHeader "Authorization", "Bearer " & AccessToken
        .SetRequestHeader "Accept", "application/json"
        .Send
        .WaitForResponse
        Dim colLists As New Collection
        If .Status = 200 Then
            Dim objJson As Scripting.Dictionary
            Set objJson = JsonConverter.ParseJson(.ResponseText)
            Dim i As Long
            Dim objCreatedBy As IdentitySet
            Dim objIdentity As Identity
            Dim objItemReference As ItemReference
            Dim objListInfo As ListInfo
            Dim objList As List
            If objJson.Item("value").Count > 0 Then
                For i = 1 To objJson.Item("value").Count
                    Set objList = New List
                    With objList
                        .Id = objJson.Item("value")(i)("id")
                        .CreatedDateTime = objJson.Item("value")(i)("createdDateTime")
                        .WebUrl = objJson.Item("value")(i)("webUrl")
                        .ETag = objJson.Item("value")(i)("eTag")
                        .Name = objJson.Item("value")(i)("name")
                        .DisplayName = objJson.Item("value")(i)("displayName")
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
                        If objJson.Item("value")(i).Exists("parentReference") Then
                            Set objItemReference = New ItemReference
                            With objItemReference
                                .SiteId = objJson.Item("value")(i)("parentReference")("siteId")
                            End With
                            Set .ParentReference = objItemReference
                        End If
                        If objJson.Item("value")(i)("list") Then
                            Set objListInfo = New ListInfo
                            With objListInfo
                                .ContentTypesEnabled = CBool(objJson.Item("value")(i)("list")("contentTypesEnabled"))
                                .Hidden = CBool(objJson.Item("value")(i)("list")("hidden"))
                                .Template = objJson.Item("value")(i)("list")("template")
                            End With
                            Set .List = objListInfo
                        End If
                        colLists.Add objList
                    End With
                Next
            End If
        Else: Call HandleError(.ResponseText)
        End If
    End With
    Set List = colLists
End Function

Public Function ListColumns(ByVal SiteId As String, ByVal ListId As String, Optional ByVal ODataQuery As String) As Collection
    Dim strUrl As String
    strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/lists/" & ListId & "/columns"
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

Public Function CreateColumn(ByVal SiteId As String, ByVal ListId As String, ByVal ColumnDefinition As ColumnDefinition) As ColumnDefinition
    Dim objWinHttp As WinHttp.WinHttpRequest
    Set objWinHttp = New WinHttp.WinHttpRequest
    With objWinHttp
        .Open "POST", "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/lists/" & ListId & "/columns", True
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
            End With
            Set CreateColumn = objColumnDefinition
        Else: Call HandleError(.ResponseText)
        End If
    End With
End Function

Public Function ListContentTypes(ByVal SiteId As String, ByVal ListId As String, Optional ByVal ODataQuery As String = vbNullString) As Collection
    Dim strUrl As String
    strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/lists/" & ListId & "/contentTypes"
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

Public Function ListOperations(ByVal SiteId As String, ByVal ListId As String, Optional ByVal ODataQuery As String = vbNullString) As Collection
    Dim strUrl As String
    strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/lists/" & ListId & "/operations"
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

Private Sub HandleError(ByVal Source As String)
    Dim objJson As Scripting.Dictionary
    Set objJson = JsonConverter.ParseJson(Source)
    Err.Raise vbObjectError, objJson.Item("error")("code"), objJson.Item("error")("message")
End Sub
