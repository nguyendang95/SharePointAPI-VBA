VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SharePointIdentitySet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private m_Application As Identity
Private m_Device As Identity
Private m_Group As Identity
Private m_User As Identity
Private m_SiteGroup As SharePointIdentity
Private m_SiteUser As SharePointIdentity

Public Property Get Application() As Identity
    Set Application = m_Application
End Property

Public Property Set Application(Value As Identity)
    Set m_Application = Value
End Property

Public Property Get Device() As Identity
    Set Device = m_Device
End Property

Public Property Set Device(Value As Identity)
    Set m_Device = Value
End Property

Public Property Get Group() As Identity
    Set Group = m_Group
End Property

Public Property Set Group(Value As Identity)
    Set m_Group = Value
End Property

Public Property Get User() As Identity
    Set User = m_User
End Property

Public Property Set User(Value As Identity)
    Set m_User = Value
End Property

Public Property Get SiteGroup() As SharePointIdentity
    Set SiteGroup = m_SiteGroup
End Property

Public Property Set SiteGroup(Value As SharePointIdentity)
    Set m_SiteGroup = Value
End Property

Public Property Get SiteUser() As SharePointIdentity
    Set SiteUser = m_SiteUser
End Property

Public Property Set SiteUser(Value As SharePointIdentity)
    Set m_SiteUser = Value
End Property

Public Function ToJson() As Scripting.Dictionary
    Dim objJson As Scripting.Dictionary
    Set objJson = New Scripting.Dictionary
    With objJson
        If Not Application Is Nothing Then .Add "application", Application.ToJson
        If Not Device Is Nothing Then .Add "device", Device.ToJson
        If Not Group Is Nothing Then .Add "group", Group.ToJson
        If Not User Is Nothing Then .Add "user", User.ToJson
        If Not SiteGroup Is Nothing Then .Add "siteGroup", SiteGroup.ToJson
        If Not SiteUser Is Nothing Then .Add "siteUser", SiteUser.ToJson
    End With
    Set ToJson = objJson
End Function

