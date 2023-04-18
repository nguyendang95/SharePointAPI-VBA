VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Drive"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_CreatedBy As IdentitySet
Private m_CreatedDateTime As String
Private m_Description As String
Private m_DriveType As String
Private m_Id As String
Private m_LastModifiedBy As IdentitySet
Private m_LastModifiedDateTime As String
Private m_Name As String
Private m_Owner As IdentitySet
Private m_Quota As Quota
Private m_SharepointIds As SharepointIds
Private m_System As SystemFacet
Private m_WebUrl As String

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

Public Property Get DriveType() As String
    DriveType = m_DriveType
End Property

Public Property Let DriveType(Value As String)
    m_DriveType = Value
End Property

Public Property Get Id() As String
    Id = m_Id
End Property

Public Property Let Id(Value As String)
    m_Id = Value
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

Public Property Get Name() As String
    Name = m_Name
End Property

Public Property Let Name(Value As String)
    m_Name = Value
End Property

Public Property Get Owner() As IdentitySet
    Set Owner = m_Owner
End Property

Public Property Set Owner(Value As IdentitySet)
    Set m_Owner = Value
End Property

Public Property Get Quota() As Quota
    Set Quota = m_Quota
End Property

Public Property Set Quota(Value As Quota)
    Set m_Quota = Value
End Property

Public Property Get SharepointIds() As SharepointIds
    Set SharepointIds = m_SharepointIds
End Property

Public Property Set SharepointIds(Value As SharepointIds)
    Set m_SharepointIds = Value
End Property

Public Property Get System() As SystemFacet
    Set System = m_System
End Property

Public Property Set System(Value As SystemFacet)
    Set m_System = Value
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
        If Not CreatedBy Is Nothing Then .Add "createdBy", CreatedBy.ToJson
        If CreatedDateTime <> vbNullString Then .Add "createdDateTime", CreatedDateTime
        If Description <> vbNullString Then .Add "description", Description
        If DriveType <> vbNullString Then .Add "driveType", DriveType
        If Id <> vbNullString Then .Add "id", Id
        If Not LastModifiedBy Is Nothing Then .Add "lastModifiedBy", LastModifiedBy.ToJson
        If LastModifiedDateTime <> vbNullString Then .Add "lastModifiedDateTime", LastModifiedDateTime
        If Name <> vbNullString Then .Add "name", Name
        If Not Owner Is Nothing Then .Add "owner", Owner.ToJson
        If Not Quota Is Nothing Then .Add "quota", Quota.ToJson
        If Not SharepointIds Is Nothing Then .Add "sharepointIds", SharepointIds.ToJson
        If Not System Is Nothing Then .Add "system", System.ToJson
        If WebUrl <> vbNullString Then .Add "webUrl", WebUrl
    End With
    Set ToJson = objJson
End Function
