VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ArrayOfSiteIds"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Value As New Collection

Public Property Get Value() As Collection
    Set Value = m_Value
End Property

Public Property Set Value(Value As Collection)
    Set m_Value = Value
End Property

Public Function ToJson() As Scripting.Dictionary
    Dim objJson As Scripting.Dictionary
    Set objJson = New Scripting.Dictionary
    With objJson
        If Value.Count > 0 Then .Add "value", Value
    End With
    Set ToJson = objJson
End Function
