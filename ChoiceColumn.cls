VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ChoiceColumn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_AllowTextEntry As Variant
Private m_Choices As New Collection
Private m_DisplayAs As String

Public Property Get AllowTextEntry() As Variant
   AllowTextEntry = m_AllowTextEntry
End Property

Public Property Let AllowTextEntry(Value As Variant)
   m_AllowTextEntry = Value
End Property

Public Property Get Choices() As Collection
  Set Choices = m_Choices
End Property

Public Property Set Choices(Value As Collection)
  Set m_Choices = Value
End Property

Public Property Get DisplayAs() As String
   DisplayAs = m_DisplayAs
End Property

Public Property Let DisplayAs(Value As String)
   m_DisplayAs = Value
End Property

Public Function ToJson() As Scripting.Dictionary
    Dim objJson As Scripting.Dictionary
    Set objJson = New Scripting.Dictionary
    With objJson
        If IsEmpty(AllowTextEntry) = False Then .Add "allowTextEntry", AllowTextEntry
        If Choices.Count > 0 Then .Add "choices", Choices
        If DisplayAs <> vbNullString Then .Add "displayAs", DisplayAs
    End With
    Set ToJson = objJson
End Function

