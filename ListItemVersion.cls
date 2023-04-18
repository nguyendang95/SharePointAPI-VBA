VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ListItemVersion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Fields As Scripting.Dictionary
Private m_Id As String
Private m_LastModifiedBy As IdentitySet
Private m_LastModifiedDateTime As String
Private m_Published As PublicationFacet

Public Property Get Fields() As Scripting.Dictionary
    Set Fields = m_Fields
End Property

Public Property Set Fields(Value As Scripting.Dictionary)
    Set m_Fields = Value
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

Public Property Get Published() As PublicationFacet
    Set Published = m_Published
End Property

Public Property Set Published(Value As PublicationFacet)
    Set m_Published = Value
End Property

Public Function ToJson() As Scripting.Dictionary
    Dim objJson As Scripting.Dictionary
    Set objJson = New Scripting.Dictionary
    With objJson
        If Not Fields Is Nothing Then .Add "fields", Fields
        If Id <> vbNullString Then .Add "id", Id
        If Not LastModifiedBy Is Nothing Then .Add "lastModifiedBy", LastModifiedBy.ToJson
        If LastModifiedDateTime <> vbNullString Then .Add "lastModifiedDateTime", LastModifiedDateTime
        If Not Published Is Nothing Then .Add "published", Published.ToJson
    End With
    Set ToJson = objJson
End Function







