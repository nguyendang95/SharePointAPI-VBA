VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ColumnDefinitionResource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum ColumnDefinitionLocation
    SiteColumn
    ListColumn
    ContentTypeColumn
    ListContentTypeColumn
End Enum

Private m_AccessToken As String

Friend Property Get AccessToken() As String
   AccessToken = m_AccessToken
End Property

Friend Property Let AccessToken(Value As String)
   m_AccessToken = Value
End Property

Public Function List(ByVal SiteId As String, ColumnType As ColumnDefinitionLocation, Optional ByVal ListId As String, Optional ByVal ContentTypeId As String, Optional ByVal ColumnId As String, Optional ByVal ODataQuery As String) As Collection
    Dim strUrl As String
    Select Case ColumnType
        Case SiteColumn
            strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/columns/" & ColumnId
        Case ListColumn
            strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/lists/" & ListId & "/columns/" & ColumnId
        Case ContentTypeColumn
            strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/contentTypes/" & ContentTypeId & "/columns/" & ColumnId
        Case ListContentTypeColumn
            strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/lists/" & ListId & "/contentTypes/" & ContentTypeId & "/columns/" & ColumnId
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
            Set List = colColumnDefinitions
        Else: Call HandleError(.ResponseText)
        End If
    End With
    
End Function

Public Function Update(ByVal SiteId As String, ColumnDefinition As ColumnDefinition, ColumnType As ColumnDefinitionLocation, Optional ByVal ListId As String, Optional ByVal ContentTypeId As String, Optional ByVal ColumnId As String, Optional ByVal ODataQuery As String = vbNullString) As ColumnDefinition
    Dim strUrl As String
    Select Case ColumnType
        Case SiteColumn
            strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/columns/" & ColumnId
        Case ListColumn
            strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/lists/" & ListId & "/columns/" & ColumnId
        Case ContentTypeColumn
            strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/contentTypes/" & ContentTypeId & "/columns/" & ColumnId
        Case ListContentTypeColumn
            strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/lists/" & ListId & "/contentTypes/" & ContentTypeId & "/columns/" & ColumnId
    End Select
    Dim objWinHttp As WinHttp.WinHttpRequest
    Set objWinHttp = New WinHttp.WinHttpRequest
    With objWinHttp
        .Open "PATCH", strUrl, True
        .SetRequestHeader "Authorization", "Bearer " & AccessToken
        .SetRequestHeader "Accept", "application/json"
        .SetRequestHeader "Content-Type", "application/json"
        .Send JsonConverter.ConvertToJson(ColumnDefinition.ToJson)
        .WaitForResponse
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
                    objDefaultValuesColumns.Value = objJson.Item("defaultValues")("value")
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
            Set Update = objColumnDefinition
        Else: Call HandleError(.ResponseText)
        End If
    End With
End Function

Public Sub Delete(ByVal SiteId As String, ColumnType As ColumnDefinitionLocation, Optional ByVal ListId As String, Optional ByVal ContentTypeId As String, Optional ByVal ColumnId As String)
    Dim strUrl As String
    Select Case ColumnType
        Case SiteColumn
            strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/columns/" & ColumnId
        Case ListColumn
            strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/lists/" & ListId & "/columns/" & ColumnId
        Case ContentTypeColumn
            strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/contentTypes/" & ContentTypeId & "/columns/" & ColumnId
        Case ListContentTypeColumn
            strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/lists/" & ListId & "/contentTypes/" & ContentTypeId & "/columns/" & ColumnId
    End Select
    Dim objWinHttp As WinHttp.WinHttpRequest
    Set objWinHttp = New WinHttp.WinHttpRequest
    With objWinHttp
        .Open "DELETE", strUrl, True
        .SetRequestHeader "Authorization", "Bearer " & AccessToken
        .SetRequestHeader "Accept", "application/json"
        .Send
        .WaitForResponse
        If .Status <> 204 Then
            Call HandleError(.ResponseText)
        End If
    End With
End Sub

Private Sub HandleError(ByVal Source As String)
    Dim objJson As Scripting.Dictionary
    Set objJson = JsonConverter.ParseJson(Source)
    Err.Raise vbObjectError, objJson.Item("error")("code"), objJson.Item("error")("message")
End Sub
