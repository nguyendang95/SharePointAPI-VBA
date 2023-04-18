VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "List"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_DisplayName As String
Private m_List As ListInfo
Private m_System As SystemFacet
Private m_Id As String
Private m_Name As String
Private m_CreatedDateTime As String
Private m_LastModifiedDateTime As String
Private m_WebUrl As String
Private m_ETag As String
Private m_SharepointIds As SharepointIds
Private m_CreatedBy As IdentitySet
Private m_LastModifiedBy As IdentitySet
Private m_ParentReference As ItemReference
Private m_Columns As New Collection
Private m_ContentTypes As New Collection
Private m_Drive As Drive
Private m_Items As New Collection
Private m_Operations As New Collection
Private m_Subscriptions As New Collection
Private m_Description As String

Public Property Get CreatedBy() As IdentitySet
    Set CreatedBy = m_CreatedBy
End Property

Public Property Set CreatedBy(Value As IdentitySet)
    Set m_CreatedBy = Value
End Property

Public Property Get System() As SystemFacet
    Set System = m_System
End Property

Public Property Set System(Value As SystemFacet)
    Set m_System = Value
End Property

Public Property Get List() As ListInfo
    Set List = m_List
End Property

Public Property Set List(Value As ListInfo)
    Set m_List = Value
End Property

Public Property Get Columns() As Collection
    Set Columns = m_Columns
End Property

Public Property Set Columns(Value As Collection)
    Set m_Columns = Value
End Property

Public Property Get ContentTypes() As Collection
    Set ContentTypes = m_ContentTypes
End Property

Public Property Set ContentTypes(Value As Collection)
    Set m_ContentTypes = Value
End Property

Public Property Get Drive() As Drive
    Set Drive = m_Drive
End Property

Public Property Set Drive(Value As Drive)
    Set m_Drive = Value
End Property

Public Property Get Items() As Collection
    Set Items = m_Items
End Property

Public Property Set Items(Value As Collection)
    Set m_Items = Value
End Property

Public Property Get Operations() As Collection
    Set Operations = m_Operations
End Property

Public Property Set Operations(Value As Collection)
    Set m_Operations = Value
End Property

Public Property Get Subscriptions() As Collection
    Set Subscriptions = m_Subscriptions
End Property

Public Property Set Subscriptions(Value As Collection)
    Set m_Subscriptions = Value
End Property

Public Property Get ParentReference() As ItemReference
  Set ParentReference = m_ParentReference
End Property

Public Property Set ParentReference(Value As ItemReference)
  Set m_ParentReference = Value
End Property

Public Property Get LastModifiedBy() As IdentitySet
  Set LastModifiedBy = m_LastModifiedBy
End Property

Public Property Set LastModifiedBy(Value As IdentitySet)
  Set m_LastModifiedBy = Value
End Property

Public Property Get SharepointIds() As SharepointIds
  Set SharepointIds = m_SharepointIds
End Property

Public Property Set SharepointIds(Value As SharepointIds)
  Set m_SharepointIds = Value
End Property

Public Property Get ETag() As String
   ETag = m_ETag
End Property

Public Property Let ETag(Value As String)
   m_ETag = Value
End Property

Public Property Get Id() As String
   Id = m_Id
End Property

Public Property Let Id(Value As String)
   m_Id = Value
End Property

Public Property Get DisplayName() As String
   DisplayName = m_DisplayName
End Property

Public Property Let DisplayName(Value As String)
   m_DisplayName = Value
End Property

Public Property Get Name() As String
   Name = m_Name
End Property

Public Property Let Name(Value As String)
   m_Name = Value
End Property

Public Property Get CreatedDateTime() As String
   CreatedDateTime = m_CreatedDateTime
End Property

Public Property Let CreatedDateTime(Value As String)
   m_CreatedDateTime = Value
End Property

Public Property Get LastModifiedDateTime() As String
   LastModifiedDateTime = m_LastModifiedDateTime
End Property

Public Property Let LastModifiedDateTime(Value As String)
   m_LastModifiedDateTime = Value
End Property

Public Property Get WebUrl() As String
   WebUrl = m_WebUrl
End Property

Public Property Let WebUrl(Value As String)
   m_WebUrl = Value
End Property

Public Property Get Description() As String
   Description = m_Description
End Property

Public Property Let Description(Value As String)
   m_Description = Value
End Property

Public Function ToJson() As Scripting.Dictionary
    Dim objJson As Scripting.Dictionary
    Set objJson = New Scripting.Dictionary
    With objJson
        If DisplayName <> vbNullString Then .Add "displayName", DisplayName
        If Not List Is Nothing Then .Add "list", List.ToJson
        If Not System Is Nothing Then .Add "system", System.ToJson
        If Id <> vbNullString Then .Add "id", Id
        If Name <> vbNullString Then .Add "name", Name
        If CreatedDateTime <> vbNullString Then .Add "createdDateTime", CreatedDateTime
        If LastModifiedDateTime <> vbNullString Then .Add "lastModifiedDateTime", LastModifiedDateTime
        If WebUrl <> vbNullString Then .Add "webUrl", WebUrl
        If ETag <> vbNullString Then .Add "eTag", ETag
        If Not SharepointIds Is Nothing Then .Add "sharepointIds", SharepointIds.ToJson
        If Not CreatedBy Is Nothing Then .Add "createdBy", CreatedBy.ToJson
        If Not LastModifiedBy Is Nothing Then .Add "lastModifiedBy", LastModifiedBy.ToJson
        If Not ParentReference Is Nothing Then .Add "parentReference", ParentReference.ToJson
        If Columns.Count > 0 Then .Add "columns", Columns
        If ContentTypes.Count > 0 Then .Add "contentTypes", ContentTypes
        If Not Drive Is Nothing Then .Add "drive", Drive.ToJson
        If Items.Count > 0 Then .Add "items", Items
        If Operations.Count > 0 Then .Add "operations", Operations
        If Subscriptions.Count > 0 Then .Add "subscriptions", Subscriptions
        If Description <> vbNullString Then .Add "description", Description
    End With
    Set ToJson = objJson
End Function
