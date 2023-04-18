VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "NumberColumn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_DecimalPlaces As String
Private m_DisplayAs As String
Private m_Maximum As Variant
Private m_Minimum As Variant

Public Property Get DecimalPlaces() As String
   DecimalPlaces = m_DecimalPlaces
End Property

Public Property Let DecimalPlaces(Value As String)
   m_DecimalPlaces = Value
End Property

Public Property Get DisplayAs() As String
   DisplayAs = m_DisplayAs
End Property

Public Property Let DisplayAs(Value As String)
   m_DisplayAs = Value
End Property

Public Property Get Maximum() As Variant
   Maximum = m_Maximum
End Property

Public Property Let Maximum(Value As Variant)
   m_Maximum = Value
End Property

Public Property Get Minimum() As Variant
   Minimum = m_Minimum
End Property

Public Property Let Minimum(Value As Variant)
   m_Minimum = Value
End Property

Public Function ToJson() As Scripting.Dictionary
    Dim objJson As Scripting.Dictionary
    Set objJson = New Scripting.Dictionary
    With objJson
        If DecimalPlaces <> vbNullString Then .Add "decimalPlaces", DecimalPlaces
        If DisplayAs <> vbNullString Then .Add "displayAs", DisplayAs
        If IsEmpty(Maximum) = False Then .Add "Maximum", Maximum
        If IsEmpty(Minimum) = False Then .Add "Minimum", Minimum
    End With
    Set ToJson = objJson
End Function

