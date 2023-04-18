VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DefaultColumnValues"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Formula As String
Private m_Value As String

Public Property Get Formula() As String
   Formula = m_Formula
End Property

Public Property Let Formula(Value As String)
   m_Formula = Value
End Property

Public Property Get Value() As String
   Value = m_Value
End Property

Public Property Let Value(Value As String)
   m_Value = Value
End Property

Public Function ToJson() As Scripting.Dictionary
    Dim objJson As Scripting.Dictionary
    Set objJson = New Scripting.Dictionary
    With objJson
        If Formula <> vbNullString Then .Add "formula", Formula
        If Value <> vbNullString Then .Add "value", Value
    End With
    Set ToJson = objJson
End Function

