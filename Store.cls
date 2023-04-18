VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Store"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Id As String
Private m_DefaultLanguageTag As String
Private m_LanguageTags As New Collection

Public Property Get Id() As String
    Id = m_Id
End Property

Public Property Let Id(Value As String)
    m_Id = Value
End Property

Public Property Get DefaultLanguageTag() As String
    DefaultLanguageTag = m_DefaultLanguageTag
End Property

Public Property Let DefaultLanguageTag(Value As String)
    m_DefaultLanguageTag = Value
End Property

Public Property Get LanguageTags() As Collection
    Set LanguageTags = m_LanguageTags
End Property

Public Property Set LanguageTags(Value As Collection)
    Set m_LanguageTags = Value
End Property

Public Function ToJson() As Scripting.Dictionary
    Dim objJson As Scripting.Dictionary
    Set objJson = New Scripting.Dictionary
    With objJson
        If Id <> vbNullString Then .Add "id", Id
        If DefaultLanguageTag <> vbNullString Then .Add "defaultLanguageTag", DefaultLanguageTag
        If LanguageTags.Count > 0 Then .Add "languageTags", LanguageTags
    End With
    Set ToJson = objJson
End Function

