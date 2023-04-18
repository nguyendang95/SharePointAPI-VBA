VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "KeyValue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Key As String
Private m_Value As String

Public Property Get Key() As String
    Key = m_Key
End Property

Public Property Let Key(Value As String)
    m_Key = Value
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
        If Key <> vbNullString Then .Add "key", Key
        If Value <> vbNullString Then .Add "value", Value
    End With
    Set ToJson = objJson
End Function


