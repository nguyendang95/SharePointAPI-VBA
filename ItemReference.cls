VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ItemReference"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_DriveId As String
Private m_DriveType As String
Private m_Id As String
Private m_Name As String
Private m_Path As String
Private m_ShareId As String
Private m_SharepointIds As SharepointIds
Private m_SiteId As String

Public Property Get Name() As String
   Name = m_Name
End Property

Public Property Let Name(Value As String)
   m_Name = Value
End Property

Public Property Get DriveId() As String
   DriveId = m_DriveId
End Property

Public Property Let DriveId(Value As String)
   m_DriveId = Value
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

Public Property Get Path() As String
   Path = m_Path
End Property

Public Property Let Path(Value As String)
   m_Path = Value
End Property

Public Property Get ShareId() As String
   ShareId = m_ShareId
End Property

Public Property Let ShareId(Value As String)
   m_ShareId = Value
End Property

Public Property Get SharepointIds() As SharepointIds
  Set SharepointIds = m_SharepointIds
End Property

Public Property Set SharepointIds(Value As SharepointIds)
  Set m_SharepointIds = Value
End Property

Public Property Get SiteId() As String
   SiteId = m_SiteId
End Property

Public Property Let SiteId(Value As String)
   m_SiteId = Value
End Property

Public Function ToJson() As Scripting.Dictionary
    Dim objJson As Scripting.Dictionary
    Set objJson = New Scripting.Dictionary
    With objJson
        If DriveId <> vbNullString Then .Add "driveId", DriveId
        If DriveType <> vbNullString Then .Add "driveType", DriveType
        If Id <> vbNullString Then .Add "id", Id
        If Name <> vbNullString Then .Add "name", Name
        If Path <> vbNullString Then .Add "path", Path
        If ShareId <> vbNullString Then .Add "shareId", ShareId
        If Not SharepointIds Is Nothing Then .Add "sharepointIds", SharepointIds.ToJson
        If SiteId <> vbNullString Then .Add "siteId", SiteId
    End With
    Set ToJson = objJson
End Function

