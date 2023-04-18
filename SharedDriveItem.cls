VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SharedDriveItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Id As String
Private m_Name As String
Private m_Owner As IdentitySet
Private m_DriveItem As DriveItem
Private m_Items As DriveItem
Private m_List As List
Private m_ListItem As ListItem
Private m_Root As DriveItem
Private m_Site As Site

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

Public Property Get Owner() As IdentitySet
    Set Owner = m_Owner
End Property

Public Property Set Owner(Value As IdentitySet)
    Set m_Owner = Value
End Property

Public Property Get DriveItem() As DriveItem
    Set DriveItem = m_DriveItem
End Property

Public Property Set DriveItem(Value As DriveItem)
    Set m_DriveItem = Value
End Property

Public Property Get Items() As DriveItem
    Set Items = m_Items
End Property

Public Property Set Items(Value As DriveItem)
    Set m_Items = Value
End Property

Public Property Get List() As List
    Set List = m_List
End Property

Public Property Set List(Value As List)
    Set m_List = Value
End Property

Public Property Get ListItem() As ListItem
    Set ListItem = m_ListItem
End Property

Public Property Set ListItem(Value As ListItem)
    Set m_ListItem = Value
End Property

Public Property Get Root() As DriveItem
    Set Root = m_Root
End Property

Public Property Set Root(Value As DriveItem)
    Set m_Root = Value
End Property

Public Property Get Site() As Site
    Set Site = m_Site
End Property

Public Property Set Site(Value As Site)
    Set m_Site = Value
End Property

Public Function ToJson() As Scripting.Dictionary
    Dim objJson As Scripting.Dictionary
    Set objJson = New Scripting.Dictionary
    With objJson
        If Id <> vbNullString Then .Add "id", Id
        If Name <> vbNullString Then .Add "name", Name
        If Not Owner Is Nothing Then .Add "owner", Owner.ToJson
        If Not DriveItem Is Nothing Then .Add "driveItem", DriveItem.ToJson
        If Not Items Is Nothing Then .Add "items", Items.ToJson
        If Not List Is Nothing Then .Add "list", List.ToJson
        If Not ListItem Is Nothing Then .Add "listItem", ListItem.ToJson
        If Not Root Is Nothing Then .Add "root", Root.ToJson
        If Not Site Is Nothing Then .Add "site", Site.ToJson
    End With
    Set ToJson = objJson
End Function


