VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PersonOrGroupColumn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_AllowMultipleSelection As Variant
Private m_DisplayAs As String
Private m_ChooseFromType As String

Public Property Get AllowMultipleSelection() As Variant
   AllowMultipleSelection = m_AllowMultipleSelection
End Property

Public Property Let AllowMultipleSelection(Value As Variant)
   m_AllowMultipleSelection = Value
End Property

Public Property Get DisplayAs() As String
   DisplayAs = m_DisplayAs
End Property

Public Property Let DisplayAs(Value As String)
   m_DisplayAs = Value
End Property

Public Property Get ChooseFromType() As String
   ChooseFromType = m_ChooseFromType
End Property

Public Property Let ChooseFromType(Value As String)
   m_ChooseFromType = Value
End Property

Public Function ToJson() As Scripting.Dictionary
    Dim objJson As Scripting.Dictionary
    Set objJson = New Scripting.Dictionary
    With objJson
        If IsEmpty(AllowMultipleSelection) = False Then .Add "AllowMultipleSelection", AllowMultipleSelection
        If DisplayAs <> vbNullString Then .Add "displayAs", DisplayAs
        If ChooseFromType <> vbNullString Then .Add "chooseFromType", ChooseFromType
    End With
    Set ToJson = objJson
End Function

