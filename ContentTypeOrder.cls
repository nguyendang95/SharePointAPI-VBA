VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ContentTypeOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Default As Variant
Private m_Position As Variant

Public Property Get Default() As Variant
    Default = m_Default
End Property

Public Property Let Default(Value As Variant)
    m_Default = Value
End Property

Public Property Get Position() As Variant
    Position = m_Position
End Property

Public Property Let Position(Value As Variant)
    m_Position = Value
End Property

Public Function ToJson() As Scripting.Dictionary
    Dim objJson As Scripting.Dictionary
    Set objJson = New Scripting.Dictionary
    With objJson
        If IsEmpty(Default) = False Then .Add "Default", Default
        If IsEmpty(Position) = False Then .Add "Position", Position
    End With
    Set ToJson = objJson
End Function

