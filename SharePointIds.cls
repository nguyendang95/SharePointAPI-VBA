VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SharePointIds"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_ListId As String
Private m_ListItemId As String
Private m_ListItemUniqueId As String
Private m_SiteId As String
Private m_SiteUrl As String
Private m_TenantId As String
Private m_WebId As String

Public Property Get ListId() As String
   ListId = m_ListId
End Property

Public Property Let ListId(Value As String)
   m_ListId = Value
End Property

Public Property Get ListItemId() As String
   ListItemId = m_ListItemId
End Property

Public Property Let ListItemId(Value As String)
   m_ListItemId = Value
End Property

Public Property Get ListItemUniqueId() As String
   ListItemUniqueId = m_ListItemUniqueId
End Property

Public Property Let ListItemUniqueId(Value As String)
   m_ListItemUniqueId = Value
End Property

Public Property Get SiteId() As String
   SiteId = m_SiteId
End Property

Public Property Let SiteId(Value As String)
   m_SiteId = Value
End Property

Public Property Get SiteUrl() As String
   SiteUrl = m_SiteUrl
End Property

Public Property Let SiteUrl(Value As String)
   m_SiteUrl = Value
End Property

Public Property Get TenantId() As String
   TenantId = m_TenantId
End Property

Public Property Let TenantId(Value As String)
   m_TenantId = Value
End Property

Public Property Get WebId() As String
   WebId = m_WebId
End Property

Public Property Let WebId(Value As String)
   m_WebId = Value
End Property

Public Function ToJson() As Scripting.Dictionary
    Dim objJson As Scripting.Dictionary
    Set objJson = New Scripting.Dictionary
    With objJson
        If ListId <> vbNullString Then .Add "listId", ListId
        If ListItemId <> vbNullString Then .Add "listItemId", ListItemId
        If SiteId <> vbNullString Then .Add "siteId", SiteId
        If SiteUrl <> vbNullString Then .Add "siteUrl", SiteUrl
        If TenantId <> vbNullString Then .Add "tenantId", TenantId
        If WebId <> vbNullString Then .Add "webId", WebId
    End With
    Set ToJson = objJson
End Function

