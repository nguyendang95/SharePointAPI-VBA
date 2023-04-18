VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LocalizedName"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_LanguageTag As String
Private m_Name As String

Public Property Get LanguageTag() As String
    LanguageTag = m_LanguageTag
End Property

Public Property Let LanguageTag(Value As String)
    m_LanguageTag = Value
End Property

Public Property Get Name() As String
    Name = m_Name
End Property

Public Property Let Name(Value As String)
    m_Name = Value
End Property

Public Function ToJson() As Scripting.Dictionary
    Dim objJson As Scripting.Dictionary
    Set objJson = New Scripting.Dictionary
    With objJson
        If LanguageTag <> vbNullString Then .Add "languageTag", LanguageTag
        If Name <> vbNullString Then .Add "name", Name
    End With
    Set ToJson = objJson
End Function


