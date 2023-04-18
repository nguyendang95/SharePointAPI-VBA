VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ContentType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_AssociatedHubsUrls As New Collection
Private m_Base As ContentType
Private m_BaseTypes As New Collection
Private m_Columns As New Collection
Private m_ColumnLinks As New Collection
Private m_ColumnPositions As New Collection
Private m_Description As String
Private m_DocumentSet As DocumentSet
Private m_DocumentTemplate As DocumentSetContent
Private m_Group As String
Private m_Hidden As Variant
Private m_Id As String
Private m_InheritedFrom As ItemReference
Private m_IsBuiltIn As Variant
Private m_Name As String
Private m_Order As ContentTypeOrder
Private m_ParentId As String
Private m_PropagateChanges As Variant
Private m_ReadOnly As Variant
Private m_Sealed As Variant


Public Property Get AssociatedHubsUrls() As Collection
    Set AssociatedHubsUrls = m_AssociatedHubsUrls
End Property

Public Property Set AssociatedHubsUrls(Value As Collection)
    Set m_AssociatedHubsUrls = Value
End Property

Public Property Get Base() As ContentType
    Set Base = m_Base
End Property

Public Property Set Base(Value As ContentType)
    Set m_Base = Value
End Property

Public Property Get BaseTypes() As Collection
    Set BaseTypes = m_BaseTypes
End Property

Public Property Set BaseTypes(Value As Collection)
    Set m_BaseTypes = Value
End Property

Public Property Get Columns() As Collection
    Set Columns = m_Columns
End Property

Public Property Set Columns(Value As Collection)
    Set m_Columns = Value
End Property

Public Property Get ColumnLinks() As Collection
    Set ColumnLinks = m_ColumnLinks
End Property

Public Property Set ColumnLinks(Value As Collection)
    Set m_ColumnLinks = Value
End Property

Public Property Get ColumnPositions() As Collection
    Set ColumnPositions = m_ColumnPositions
End Property

Public Property Set ColumnPositions(Value As Collection)
    Set m_ColumnPositions = Value
End Property

Public Property Get Description() As String
    Description = m_Description
End Property

Public Property Let Description(Value As String)
    m_Description = Value
End Property

Public Property Get DocumentSet() As DocumentSet
    Set DocumentSet = m_DocumentSet
End Property

Public Property Set DocumentSet(Value As DocumentSet)
    Set m_DocumentSet = Value
End Property

Public Property Get DocumentTemplate() As DocumentSetContent
    Set DocumentTemplate = m_DocumentTemplate
End Property

Public Property Set DocumentTemplate(Value As DocumentSetContent)
    Set m_DocumentTemplate = Value
End Property

Public Property Get Group() As String
    Group = m_Group
End Property

Public Property Let Group(Value As String)
    m_Group = Value
End Property

Public Property Get Hidden() As Variant
    Hidden = m_Hidden
End Property

Public Property Let Hidden(Value As Variant)
    m_Hidden = Value
End Property

Public Property Get Id() As String
    Id = m_Id
End Property

Public Property Let Id(Value As String)
    m_Id = Value
End Property

Public Property Get InheritedFrom() As ItemReference
    Set InheritedFrom = m_InheritedFrom
End Property

Public Property Set InheritedFrom(Value As ItemReference)
    Set m_InheritedFrom = Value
End Property

Public Property Get IsBuiltIn() As Variant
    IsBuiltIn = m_IsBuiltIn
End Property

Public Property Let IsBuiltIn(Value As Variant)
    m_IsBuiltIn = Value
End Property

Public Property Get Name() As String
    Name = m_Name
End Property

Public Property Let Name(Value As String)
    m_Name = Value
End Property

Public Property Get Order() As ContentTypeOrder
    Set Order = m_Order
End Property

Public Property Set Order(Value As ContentTypeOrder)
    Set m_Order = Value
End Property

Public Property Get ParentId() As String
    ParentId = m_ParentId
End Property

Public Property Let ParentId(Value As String)
    m_ParentId = Value
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

Public Property Get Sealed() As Variant
    Sealed = m_Sealed
End Property

Public Property Let Sealed(Value As Variant)
    m_Sealed = Value
End Property

Public Function ToJson() As Scripting.Dictionary
    Dim objJson As Scripting.Dictionary
    Set objJson = New Scripting.Dictionary
    With objJson
        If AssociatedHubsUrls.Count > 0 Then .Add "associatedHubsUrls", AssociatedHubsUrls
        If Not Base Is Nothing Then .Add "base", Base.ToJson
        If BaseTypes.Count > 0 Then .Add "baseTypes", BaseTypes
        If Columns.Count > 0 Then .Add "columns", Columns
        If ColumnLinks.Count > 0 Then .Add "columnLinks", ColumnLinks
        If ColumnPositions.Count > 0 Then .Add "columnPositions", ColumnPositions
        If Description <> vbNullString Then .Add "description", Description
        If Not DocumentSet Is Nothing Then .Add "documentSet", DocumentSet.ToJson
        If Not DocumentTemplate Is Nothing Then .Add "documentTemplate", DocumentTemplate.ToJson
        If Group <> vbNullString Then .Add "group", Group
        If Id <> vbNullString Then .Add "id", Id
        If Not InheritedFrom Is Nothing Then .Add "inheritedFrom", InheritedFrom.ToJson
        If Name <> vbNullString Then .Add "name", Name
        If IsEmpty(Hidden) = False Then .Add "hidden", Hidden
        If Not Order Is Nothing Then .Add "order", Order.ToJson
        If ParentId <> vbNullString Then .Add "parentId", ParentId
    End With
    Set ToJson = objJson
End Function




