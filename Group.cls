VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Group"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_CreatedDateTime As String
Private m_Description As String
Private m_DisplayName As String
Private m_Id As String
Private m_ParentSiteId As String
Private m_Scope As String

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

Public Property Get DisplayName() As String
    DisplayName = m_DisplayName
End Property

Public Property Let DisplayName(Value As String)
    m_DisplayName = Value
End Property

Public Property Get Id() As String
    Id = m_Id
End Property

Public Property Let Id(Value As String)
    m_Id = Value
End Property

Public Property Get ParentSiteId() As String
    ParentSiteId = m_ParentSiteId
End Property

Public Property Let ParentSiteId(Value As String)
    m_ParentSiteId = Value
End Property

Public Property Get Scope() As String
    Scope = m_Scope
End Property

Public Property Let Scope(Value As String)
    m_Scope = Value
End Property




Public Function ToJson() As Scripting.Dictionary
    Dim objJson As Scripting.Dictionary
    Set objJson = New Scripting.Dictionary
    With objJson
        If CreatedDateTime <> vbNullString Then .Add "createdDateTime", CreatedDateTime
        If Description <> vbNullString Then .Add "description", Description
        If DisplayName <> vbNullString Then .Add "displayName", DisplayName
        If Id <> vbNullString Then .Add "id", Id
        If ParentSiteId <> vbNullString Then .Add "parentSiteId", ParentSiteId
        If Scope <> vbNullString Then .Add "scope", Scope
    End With
    Set ToJson = objJson
End Function

