VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ListInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_ContentTypesEnabled As Variant
Private m_Hidden As Variant
Private m_Template As String

Public Property Get ContentTypesEnabled() As Variant
   ContentTypesEnabled = m_ContentTypesEnabled
End Property

Public Property Let ContentTypesEnabled(Value As Variant)
   m_ContentTypesEnabled = Value
End Property

Public Property Get Hidden() As Variant
   Hidden = m_Hidden
End Property

Public Property Let Hidden(Value As Variant)
   m_Hidden = Value
End Property

Public Property Get Template() As String
   Template = m_Template
End Property

Public Property Let Template(Value As String)
   m_Template = Value
End Property

Public Function ToJson() As Scripting.Dictionary
    Dim objJson As Scripting.Dictionary
    Set objJson = New Scripting.Dictionary
    With objJson
        If IsEmpty(ContentTypesEnabled) = False Then .Add "ContentTypeEnabled", ContentTypesEnabled
        If IsEmpty(Hidden) = False Then .Add "Hidden", Hidden
        If Template <> vbNullString Then .Add "template", Template
    End With
    Set ToJson = objJson
End Function

