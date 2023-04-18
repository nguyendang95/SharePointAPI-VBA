VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Site"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Id As String
Private m_DisplayName As String
Private m_Name As String
Private m_CreatedDateTime As String
Private m_LastModifiedDateTime As String
Private m_WebUrl As String
Private m_ETag As String
Private m_SiteCollection As SiteCollection
Private m_Root As Root
Private m_SharepointIds As SharepointIds
Private m_Error As ErrorObject

Public Property Get Error() As ErrorObject
  Set Error = m_Error
End Property

Public Property Set Error(Value As ErrorObject)
  Set m_Error = Value
End Property

Public Property Get SharepointIds() As SharepointIds
  Set SharepointIds = m_SharepointIds
End Property

Public Property Set SharepointIds(Value As SharepointIds)
  Set m_SharepointIds = Value
End Property

Public Property Get Root() As Root
  Set Root = m_Root
End Property

Public Property Set Root(Value As Root)
  Set m_Root = Value
End Property

Public Property Get SiteCollection() As SiteCollection
  Set SiteCollection = m_SiteCollection
End Property

Public Property Set SiteCollection(Value As SiteCollection)
  Set m_SiteCollection = Value
End Property

Public Property Get ETag() As String
   ETag = m_ETag
End Property

Public Property Let ETag(Value As String)
   m_ETag = Value
End Property

Public Property Get Id() As String
   Id = m_Id
End Property

Public Property Let Id(Value As String)
   m_Id = Value
End Property

Public Property Get DisplayName() As String
   DisplayName = m_DisplayName
End Property

Public Property Let DisplayName(Value As String)
   m_DisplayName = Value
End Property

Public Property Get Name() As String
   Name = m_Name
End Property

Public Property Let Name(Value As String)
   m_Name = Value
End Property

Public Property Get CreatedDateTime() As String
   CreatedDateTime = m_CreatedDateTime
End Property

Public Property Let CreatedDateTime(Value As String)
   m_CreatedDateTime = Value
End Property

Public Property Get LastModifiedDateTime() As String
   LastModifiedDateTime = m_LastModifiedDateTime
End Property

Public Property Let LastModifiedDateTime(Value As String)
   m_LastModifiedDateTime = Value
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
        If Id <> vbNullString Then .Add "id", Id
        If DisplayName <> vbNullString Then .Add "displayName", DisplayName
        If Name <> vbNullString Then .Add "name", Name
        If CreatedDateTime <> vbNullString Then .Add "createdDateTime", CreatedDateTime
        If LastModifiedDateTime <> vbNullString Then .Add "lastModifiedDateTime", LastModifiedDateTime
        If WebUrl <> vbNullString Then .Add "webUrl", WebUrl
        If ETag <> vbNullString Then .Add "eTag", ETag
        If Not SiteCollection Is Nothing Then .Add "siteCollection", SiteCollection.ToJson
        If Not Root Is Nothing Then .Add "root", Root.ToJson
        If Not SharepointIds Is Nothing Then .Add "sharepointIds", SharepointIds.ToJson
        If Not Error Is Nothing Then .Add "error", Error.ToJson
    End With
    Set ToJson = objJson
End Function

