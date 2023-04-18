VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DocumentSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_AllowedContentTypes As New Collection
Private m_DefaultContents As DocumentSetContent
Private m_PropagateWelcomePageChanges As Variant
Private m_ShouldPrefixNameToFile As Variant
Private m_WelcomePageUrl As String

Public Property Get AllowedContentTypes() As Collection
    Set AllowedContentTypes = m_AllowedContentTypes
End Property

Public Property Set AllowedContentTypes(Value As Collection)
    Set m_AllowedContentTypes = Value
End Property

Public Property Get DefaultContents() As DocumentSetContent
    Set DefaultContents = m_DefaultContents
End Property

Public Property Set DefaultContents(Value As DocumentSetContent)
    Set m_DefaultContents = Value
End Property

Public Property Get PropagateWelcomePageChanges() As Variant
    PropagateWelcomePageChanges = m_PropagateWelcomePageChanges
End Property

Public Property Let PropagateWelcomePageChanges(Value As Variant)
    m_PropagateWelcomePageChanges = Value
End Property

Public Property Get ShouldPrefixNameToFile() As Variant
    ShouldPrefixNameToFile = m_ShouldPrefixNameToFile
End Property

Public Property Let ShouldPrefixNameToFile(Value As Variant)
    m_ShouldPrefixNameToFile = Value
End Property

Public Property Get WelcomePageUrl() As String
    WelcomePageUrl = m_WelcomePageUrl
End Property

Public Property Let WelcomePageUrl(Value As String)
    m_WelcomePageUrl = Value
End Property

Public Function ToJson() As Scripting.Dictionary
    Dim objJson As Scripting.Dictionary
    Set objJson = New Scripting.Dictionary
    With objJson
        If AllowedContentTypes.Count > 0 Then .Add "allowedContentTypes", AllowedContentTypes
        If Not DefaultContents Is Nothing Then .Add "defaultContents", DefaultContents.ToJson
        If IsEmpty(PropagateWelcomePageChanges) = False Then .Add "PropagateWelcomePageChanges", PropagateWelcomePageChanges
        If IsEmpty(ShouldPrefixNameToFile) = False Then .Add "ShouldPrefixNameToFile", ShouldPrefixNameToFile
        If WelcomePageUrl <> vbNullString Then .Add "welcomePageUrl", WelcomePageUrl
    End With
    Set ToJson = objJson
End Function


