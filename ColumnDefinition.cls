VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ColumnDefinition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_BooleanColumn As BooleanColumn
Private m_Calculated As CalculatedColumn
Private m_Choice As ChoiceColumn
Private m_ColumnGroup As String
Private m_ContentApprovalStatus As ContentApprovalStatusColumn
Private m_CurrencyColumn As CurrencyColumn
Private m_DateTime As DateTimeColumn
Private m_DefaultValues As DefaultColumnValues
Private m_Description As String
Private m_DisplayName As String
Private m_EnforceUniqueValues As Variant
Private m_Geolocation As GeoLocationColumn
Private m_Hidden As Variant
Private m_HyperlinkOrPicture As HyperlinkOrPictureColumn
Private m_IsDeletable As Variant
Private m_IsReOrderable As Variant
Private m_Id As String
Private m_Indexed As Variant
Private m_IsSealed As Variant
Private m_Lookup As LookupColumn
Private m_Name As String
Private m_Number As NumberColumn
Private m_PersonOrGroup As PersonOrGroupColumn
Private m_PropagateChanges As Variant
Private m_ReadOnly As Variant
Private m_Required As Variant
Private m_SourceContentType As ContentTypeInfo
Private m_Term As TermColumn
Private m_Text As TextColumn
Private m_Thumbnail As ThumbnailColumn
Private m_Validation As ColumnValidation

Public Property Get BooleanColumn() As BooleanColumn
  Set BooleanColumn = m_BooleanColumn
End Property

Public Property Set BooleanColumn(Value As BooleanColumn)
  Set m_BooleanColumn = Value
End Property

Public Property Get Calculated() As CalculatedColumn
  Set Calculated = m_Calculated
End Property

Public Property Set Calculated(Value As CalculatedColumn)
  Set m_Calculated = Value
End Property

Public Property Get Choice() As ChoiceColumn
  Set Choice = m_Choice
End Property

Public Property Set Choice(Value As ChoiceColumn)
  Set m_Choice = Value
End Property

Public Property Get ColumnGroup() As String
   ColumnGroup = m_ColumnGroup
End Property

Public Property Let ColumnGroup(Value As String)
   m_ColumnGroup = Value
End Property

Public Property Get ContentApprovalStatus() As ContentApprovalStatusColumn
  Set ContentApprovalStatus = m_ContentApprovalStatus
End Property

Public Property Set ContentApprovalStatus(Value As ContentApprovalStatusColumn)
  Set m_ContentApprovalStatus = Value
End Property

Public Property Get CurrencyColumn() As CurrencyColumn
  Set CurrencyColumn = m_CurrencyColumn
End Property

Public Property Set CurrencyColumn(Value As CurrencyColumn)
  Set m_CurrencyColumn = Value
End Property

Public Property Get DateTime() As DateTimeColumn
  Set DateTime = m_DateTime
End Property

Public Property Set DateTime(Value As DateTimeColumn)
  Set m_DateTime = Value
End Property

Public Property Get DefaultValues() As DefaultColumnValues
  Set DefaultValues = m_DefaultValues
End Property

Public Property Set DefaultValues(Value As DefaultColumnValues)
  Set m_DefaultValues = Value
End Property

Public Property Get Description() As String
   Description = m_Description
End Property

Public Property Let Description(Value As String)
   m_Description = Value
End Property

Public Property Get DisplayName() As String
   DisplayName = m_DisplayName
End Property

Public Property Let DisplayName(Value As String)
   m_DisplayName = Value
End Property

Public Property Get EnforceUniqueValues() As Variant
   EnforceUniqueValues = m_EnforceUniqueValues
End Property

Public Property Let EnforceUniqueValues(Value As Variant)
   m_EnforceUniqueValues = Value
End Property

Public Property Get Geolocation() As GeoLocationColumn
  Set Geolocation = m_Geolocation
End Property

Public Property Set Geolocation(Value As GeoLocationColumn)
  Set m_Geolocation = Value
End Property

Public Property Get Hidden() As Variant
   Hidden = m_Hidden
End Property

Public Property Let Hidden(Value As Variant)
   m_Hidden = Value
End Property

Public Property Get HyperlinkOrPicture() As HyperlinkOrPictureColumn
  Set HyperlinkOrPicture = m_HyperlinkOrPicture
End Property

Public Property Set HyperlinkOrPicture(Value As HyperlinkOrPictureColumn)
  Set m_HyperlinkOrPicture = Value
End Property

Public Property Get IsDeletable() As Variant
   IsDeletable = m_IsDeletable
End Property

Public Property Let IsDeletable(Value As Variant)
   m_IsDeletable = Value
End Property

Public Property Get IsReorderable() As Variant
   IsReorderable = m_IsReOrderable
End Property

Public Property Let IsReorderable(Value As Variant)
   m_IsReOrderable = Value
End Property

Public Property Get Id() As String
   Id = m_Id
End Property

Public Property Let Id(Value As String)
   m_Id = Value
End Property

Public Property Get Indexed() As Variant
   Indexed = m_Indexed
End Property

Public Property Let Indexed(Value As Variant)
   m_Indexed = Value
End Property

Public Property Get IsSealed() As Variant
   IsSealed = m_IsSealed
End Property

Public Property Let IsSealed(Value As Variant)
   m_IsSealed = Value
End Property

Public Property Get Lookup() As LookupColumn
  Set Lookup = m_Lookup
End Property

Public Property Set Lookup(Value As LookupColumn)
  Set m_Lookup = Value
End Property

Public Property Get Name() As String
   Name = m_Name
End Property

Public Property Let Name(Value As String)
   m_Name = Value
End Property

Public Property Get Number() As NumberColumn
  Set Number = m_Number
End Property

Public Property Set Number(Value As NumberColumn)
  Set m_Number = Value
End Property

Public Property Get PersonOrGroup() As PersonOrGroupColumn
  Set PersonOrGroup = m_PersonOrGroup
End Property

Public Property Set PersonOrGroup(Value As PersonOrGroupColumn)
  Set m_PersonOrGroup = Value
End Property

Public Property Get PropagateChanges() As Variant
   PropagateChanges = m_PropagateChanges
End Property

Public Property Let PropagateChanges(Value As Variant)
   m_PropagateChanges = Value
End Property

Public Property Get ReadOnly() As Variant
   ReadOnly = m_ReadOnly
End Property

Public Property Let ReadOnly(Value As Variant)
   m_ReadOnly = Value
End Property

Public Property Get Required() As Variant
   Required = m_Required
End Property

Public Property Let Required(Value As Variant)
   m_Required = Value
End Property

Public Property Get SourceContentType() As ContentTypeInfo
  Set SourceContentType = m_SourceContentType
End Property

Public Property Set SourceContentType(Value As ContentTypeInfo)
  Set m_SourceContentType = Value
End Property

Public Property Get Term() As TermColumn
  Set Term = m_Term
End Property

Public Property Set Term(Value As TermColumn)
  Set m_Term = Value
End Property

Public Property Get Text() As TextColumn
  Set Text = m_Text
End Property

Public Property Set Text(Value As TextColumn)
  Set m_Text = Value
End Property

Public Property Get Thumbnail() As ThumbnailColumn
  Set Thumbnail = m_Thumbnail
End Property

Public Property Set Thumbnail(Value As ThumbnailColumn)
  Set m_Thumbnail = Value
End Property

Public Property Get Validation() As ColumnValidation
  Set Validation = m_Validation
End Property

Public Property Set Validation(Value As ColumnValidation)
  Set m_Validation = Value
End Property

Public Function ToJson() As Scripting.Dictionary
    Dim objJson As Scripting.Dictionary
    Set objJson = New Scripting.Dictionary
    With objJson
        If Not BooleanColumn Is Nothing Then .Add "booleanColumn", BooleanColumn.ToJson
        If Not Calculated Is Nothing Then .Add "calculated", Calculated.ToJson
        If Not Choice Is Nothing Then .Add "choice", Choice.ToJson
        If ColumnGroup <> vbNullString Then .Add "columnGroup", ColumnGroup
        If Not ContentApprovalStatus Is Nothing Then .Add "contentApprovalStatus", ContentApprovalStatus.ToJson
        If Not CurrencyColumn Is Nothing Then .Add "currencyColumn", CurrencyColumn.ToJson
        If Not DateTime Is Nothing Then .Add "dateTime", DateTime.ToJson
        If Not DefaultValues Is Nothing Then .Add "defaultValue", DefaultValues.ToJson
        If Description <> vbNullString Then .Add "description", Description
        If DisplayName <> vbNullString Then .Add "displayName", DisplayName
        If IsEmpty(EnforceUniqueValues) = False Then .Add "enforceUniqueValues", EnforceUniqueValues
        If Not Geolocation Is Nothing Then .Add "geolocation", Geolocation.ToJson
        If IsEmpty(Hidden) = False Then .Add "hidden", Hidden
        If Not HyperlinkOrPicture Is Nothing Then .Add "hyperlinkOrPicture", HyperlinkOrPicture.ToJson
        If IsEmpty(Indexed) = False Then .Add "indexed", Indexed
        If Not Lookup Is Nothing Then .Add "lookup", Lookup.ToJson
        If Name <> vbNullString Then .Add "name", Name
        If Not Number Is Nothing Then .Add "number", Number.ToJson
        If Not PersonOrGroup Is Nothing Then .Add "personOrGroup", PersonOrGroup.ToJson
        If Not Term Is Nothing Then .Add "term", Term.ToJson
        If Not Text Is Nothing Then .Add "text", Text.ToJson
        If Not Thumbnail Is Nothing Then .Add "thumbnail", Thumbnail.ToJson
        If Not Validation Is Nothing Then .Add "validation", Validation.ToJson
    End With
    Set ToJson = objJson
End Function

