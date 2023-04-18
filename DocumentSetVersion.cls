VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DocumentSetVersion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Comment As String
Private m_CreatedDateTime As String
Private m_CreatedBy As IdentitySet
Private m_Id As String
Private m_Items As New Collection
Private m_LastModifiedBy As IdentitySet
Private m_LastModifiedDateTime As String
Private m_Publication As PublicationFacet
Private m_ShouldCaptureMinorVersion As Variant

Public Property Get Comment() As String
    Comment = m_Comment
End Property

Public Property Let Comment(Value As String)
    m_Comment = Value
End Property

Public Property Get CreatedDateTime() As String
    CreatedDateTime = m_CreatedDateTime
End Property

Public Property Let CreatedDateTime(Value As String)
    m_CreatedDateTime = Value
End Property

Public Property Get CreatedBy() As IdentitySet
    Set CreatedBy = m_CreatedBy
End Property

Public Property Set CreatedBy(Value As IdentitySet)
    Set m_CreatedBy = Value
End Property

Public Property Get Id() As String
    Id = m_Id
End Property

Public Property Let Id(Value As String)
    m_Id = Value
End Property

Public Property Get Items() As Collection
    Set Items = m_Items
End Property

Public Property Set Items(Value As Collection)
    Set m_Items = Value
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

Public Property Get ShouldCaptureMinorVersion() As Variant
    ShouldCaptureMinorVersion = m_ShouldCaptureMinorVersion
End Property

Public Property Let ShouldCaptureMinorVersion(Value As Variant)
    m_ShouldCaptureMinorVersion = Value
End Property

Public Function ToJson() As Scripting.Dictionary
    Dim objJson As Scripting.Dictionary
    Set objJson = New Scripting.Dictionary
    With objJson
        If Comment <> vbNullString Then .Add "comment", Comment
        If CreatedDateTime <> vbNullString Then .Add "createdDateTime", CreatedDateTime
        If Not CreatedBy Is Nothing Then .Add "createdBy", CreatedBy.ToJson
        If Id <> vbNullString Then .Add "id", Id
        If Items.Count > 0 Then .Add "items", Items
        If Not LastModifiedBy Is Nothing Then .Add "lastModifiedBy", LastModifiedBy.ToJson
        If LastModifiedDateTime <> vbNullString Then .Add "lastModifiedDateTime", LastModifiedDateTime
        If Not Publication Is Nothing Then .Add "publication", Publication.ToJson
        If IsEmpty(ShouldCaptureMinorVersion) = False Then .Add "ShouldCaptureMinorVersion", ShouldCaptureMinorVersion
    End With
    Set ToJson = objJson
End Function

