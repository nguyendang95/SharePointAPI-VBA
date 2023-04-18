VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ItemActionStat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_ActionCount As Variant
Private m_ActorCount As Variant

Public Property Get ActionCount() As Variant
    ActionCount = m_ActionCount
End Property

Public Property Let ActionCount(Value As Variant)
    m_ActionCount = Value
End Property

Public Property Get ActorCount() As Variant
    ActorCount = m_ActorCount
End Property

Public Property Let ActorCount(Value As Variant)
    m_ActorCount = Value
End Property

Public Function ToJson() As Scripting.Dictionary
    Dim objJson As Scripting.Dictionary
    Set objJson = New Scripting.Dictionary
    With objJson
        If IsEmpty(ActionCount) = False Then .Add "ActionCount", ActionCount
        If IsEmpty(ActorCount) = False Then .Add "ActorCount", ActorCount
    End With
    Set ToJson = objJson
End Function

