VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DriveItemVersion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Id As String
Private m_LastModifiedBy As IdentitySet
Private m_LastModifiedDateTime As String
Private m_Publication As PublicationFacet
Private m_Size As Variant
Private m_DownloadUrl As String

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

Public Property Get Publication() As PublicationFacet
    Set Publication = m_Publication
End Property

Public Property Set Publication(Value As PublicationFacet)
    Set m_Publication = Value
End Property

Public Property Get Size() As Variant
    Size = m_Size
End Property

Public Property Let Size(Value As Variant)
    m_Size = Value
End Property

Public Property Get DownloadUrl() As String
    DownloadUrl = m_DownloadUrl
End Property

Public Property Let DownloadUrl(Value As String)
    m_DownloadUrl = Value
End Property

Public Function ToJson() As Scripting.Dictionary
    Dim objJson As Scripting.Dictionary
    Set objJson = New Scripting.Dictionary
    With objJson
        If Id <> vbNullString Then .Add "id", Id
        If Not LastModifiedBy Is Nothing Then .Add "lastModifiedBy", LastModifiedBy.ToJson
        If LastModifiedDateTime <> vbNullString Then .Add "lastModifiedDateTime", LastModifiedDateTime
        If Not Publication Is Nothing Then .Add "publication", Publication.ToJson
        If IsEmpty(Size) = False Then .Add "size", Size
    End With
    Set ToJson = objJson
End Function




