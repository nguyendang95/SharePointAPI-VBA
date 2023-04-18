VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ListItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_ContentType As ContentTypeInfo
Private m_Fields As Scripting.Dictionary
Private m_SharepointIds As SharepointIds
Private m_Activities As New Collection
Private m_Analytics As ItemAnalytics
Private m_DocumentSetVersions As New Collection
Private m_DriveItem As DriveItem
Private m_Versions As New Collection
Private m_Id As String
Private m_Name As String
Private m_CreatedBy As IdentitySet
Private m_CreatedDateTime As String
Private m_Description As String
Private m_ETag As String
Private m_LastModifiedBy As IdentitySet
Private m_LastModifiedDateTime As String
Private m_ParentReference As ItemReference
Private m_WebUrl As String

Public Property Get ContentType() As ContentTypeInfo
    Set ContentType = m_ContentType
End Property

Public Property Set ContentType(Value As ContentTypeInfo)
    Set m_ContentType = Value
End Property

Public Property Get Fields() As Scripting.Dictionary
    Set Fields = m_Fields
End Property

Public Property Set Fields(Value As Scripting.Dictionary)
    Set m_Fields = Value
End Property

Public Property Get SharepointIds() As SharepointIds
    Set SharepointIds = m_SharepointIds
End Property

Public Property Set SharepointIds(Value As SharepointIds)
    Set m_SharepointIds = Value
End Property

Public Property Get Activities() As Collection
    Set Activities = m_Activities
End Property

Public Property Set Activities(Value As Collection)
    Set m_Activities = Value
End Property

Public Property Get Analytics() As ItemAnalytics
    Set Analytics = m_Analytics
End Property

Public Property Set Analytics(Value As ItemAnalytics)
    Set m_Analytics = Value
End Property

Public Property Get DocumentSetVersions() As Collection
    Set DocumentSetVersions = m_DocumentSetVersions
End Property

Public Property Set DocumentSetVersions(Value As Collection)
    Set m_DocumentSetVersions = Value
End Property

Public Property Get DriveItem() As DriveItem
    Set DriveItem = m_DriveItem
End Property

Public Property Set DriveItem(Value As DriveItem)
    Set m_DriveItem = Value
End Property

Public Property Get Versions() As Collection
    Set Versions = m_Versions
End Property

Public Property Set Versions(Value As Collection)
    Set m_Versions = Value
End Property

Public Property Get Id() As String
    Id = m_Id
End Property

Public Property Let Id(Value As String)
    m_Id = Value
End Property

Public Property Get Name() As String
    Name = m_Name
End Property

Public Property Let Name(Value As String)
    m_Name = Value
End Property

Public Property Get CreatedBy() As IdentitySet
    Set CreatedBy = m_CreatedBy
End Property

Public Property Set CreatedBy(Value As IdentitySet)
    Set m_CreatedBy = Value
End Property

Public Property Get CreatedDateTime() As String
    CreatedDateTime = m_CreatedDateTime
End Property

Public Property Let CreatedDateTime(Value As String)
    m_CreatedDateTime = Value
End Property

Public Property Get Description() As String
    Description = m_Description
End Property

Public Property Let Description(Value As String)
    m_Description = Value
End Property

Public Property Get ETag() As String
    ETag = m_ETag
End Property

Public Property Let ETag(Value As String)
    m_ETag = Value
End Property

Public Property Get LastModifiedBy() As IdentitySet
    Set LastModifiedBy = m_LastModifiedBy
End Property

Public Property Set LastModifiedBy(Value As IdentitySet)
    Set m_LastModifiedBy = Value
End Property

Public Property Get LastModifiedDateTime() As String
    LastModifiedDateTime = m_LastModifiedDateTime
End Property

Public Property Let LastModifiedDateTime(Value As String)
    m_LastModifiedDateTime = Value
End Property

Public Property Get ParentReference() As ItemReference
    Set ParentReference = m_ParentReference
End Property

Public Property Set ParentReference(Value As ItemReference)
    Set m_ParentReference = Value
End Property

Public Property Get WebUrl() As String
    WebUrl = m_WebUrl
End Property

Public Property Let WebUrl(Value As String)
    m_WebUrl = Value
End Property

Public Function ToJson() As Scripting.Dictionary
    Dim objJson As Scripting.Dictionary
    Set objJson = New Scripting.Dictionary
    With objJson
        If Not ContentType Is Nothing Then .Add "contentType", ContentType.ToJson
        If Not Fields Is Nothing Then .Add "fields", Fields
        If Not SharepointIds Is Nothing Then .Add "sharepointIds", SharepointIds.ToJson
        If Activities.Count > 0 Then .Add "activities", Activities
        If Not Analytics Is Nothing Then .Add "analytics", Analytics.ToJson
        If DocumentSetVersions.Count > 0 Then .Add "documentSetVersions", DocumentSetVersions
        If Not DriveItem Is Nothing Then .Add "driveItem", DriveItem.ToJson
        If Versions.Count > 0 Then .Add "versions", Versions
        If Id <> vbNullString Then .Add "id", Id
        If Name <> vbNullString Then .Add "name", Name
        If Not CreatedBy Is Nothing Then .Add "createdBy", CreatedBy.ToJson
        If CreatedDateTime <> vbNullString Then .Add "createdDateTime", CreatedDateTime
        If Description <> vbNullString Then .Add "description", Description
        If ETag <> vbNullString Then .Add "eTag", ETag
        If Not LastModifiedBy Is Nothing Then .Add "lastModifiedBy", LastModifiedBy.ToJson
        If LastModifiedDateTime <> vbNullString Then .Add "lastModifiedDateTime", LastModifiedDateTime
        If Not ParentReference Is Nothing Then .Add "parentReference", ParentReference.ToJson
        If WebUrl <> vbNullString Then .Add "webUrl", WebUrl
    End With
    Set ToJson = objJson
End Function

