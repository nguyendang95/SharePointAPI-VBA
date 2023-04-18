VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TermColumn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_AllowMultipleValues As Variant
Private m_ShowFullyQualifiedName As Variant

Public Property Get AllowMultipleValues() As Variant
   AllowMultipleValues = m_AllowMultipleValues
End Property

Public Property Let AllowMultipleValues(Value As Variant)
   m_AllowMultipleValues = Value
End Property

Public Property Get ShowFullyQualifiedName() As Variant
   ShowFullyQualifiedName = m_ShowFullyQualifiedName
End Property

Public Property Let ShowFullyQualifiedName(Value As Variant)
   m_ShowFullyQualifiedName = Value
End Property

Public Function ToJson() As Scripting.Dictionary
    Dim objJson As Scripting.Dictionary
    Set objJson = New Scripting.Dictionary
    With objJson
        If IsEmpty(AllowMultipleValues) = False Then .Add "AllowMultipleValues", AllowMultipleValues
        If IsEmpty(ShowFullyQualifiedName) = False Then .Add "ShowFullyQualifiedName", ShowFullyQualifiedName
    End With
    Set ToJson = objJson
End Function

