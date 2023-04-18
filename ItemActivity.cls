VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ItemActivity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Id As String
Private m_Access As AccessAction
Private m_Actor As IdentitySet
Private m_DriveItem As DriveItem
Private m_ListItem As ListItem
Private m_ActivityDateTime As String

Public Property Get Id() As String
    Id = m_Id
End Property

Public Property Let Id(Value As String)
    m_Id = Value
End Property

Public Property Get Access() As AccessAction
    Set Access = m_Access
End Property

Public Property Set Access(Value As AccessAction)
    Set m_Access = Value
End Property

Public Property Get Actor() As IdentitySet
    Set Actor = m_Actor
End Property

Public Property Set Actor(Value As IdentitySet)
    Set m_Actor = Value
End Property

Public Property Get DriveItem() As DriveItem
    Set DriveItem = m_DriveItem
End Property

Public Property Set DriveItem(Value As DriveItem)
    Set m_DriveItem = Value
End Property

Public Property Get ListItem() As ListItem
    Set ListItem = m_ListItem
End Property

Public Property Set ListItem(Value As ListItem)
    Set m_ListItem = Value
End Property

Public Property Get ActivityDateTime() As String
    ActivityDateTime = m_ActivityDateTime
End Property

Public Property Let ActivityDateTime(Value As String)
    m_ActivityDateTime = Value
End Property

Public Function ToJson() As Scripting.Dictionary
    Dim objJson As Scripting.Dictionary
    Set objJson = New Scripting.Dictionary
    With objJson
        If Id <> vbNullString Then .Add "id", Id
        If Not Access Is Nothing Then .Add "access", Access.ToJson
        If Not Actor Is Nothing Then .Add "actor", Actor.ToJson
        If Not DriveItem Is Nothing Then .Add "driveItem", DriveItem.ToJson
        If Not ListItem Is Nothing Then .Add "listItem", ListItem.ToJson
        If ActivityDateTime <> vbNullString Then .Add "activityDateTime", ActivityDateTime
    End With
    Set ToJson = objJson
End Function



