VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UploadableParameters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Item As DriveItemUploadableProperties
Private m_DeferCommit As Variant

Public Property Get Item() As DriveItemUploadableProperties
    Set Item = m_Item
End Property

Public Property Set Item(Value As DriveItemUploadableProperties)
    Set m_Item = Value
End Property

Public Property Get DeferCommit() As Variant
    DeferCommit = m_DeferCommit
End Property

Public Property Let DeferCommit(Value As Variant)
    m_DeferCommit = Value
End Property

Public Function ToJson() As Scripting.Dictionary
    Dim objJson As Scripting.Dictionary
    Set objJson = New Scripting.Dictionary
    With objJson
        If Not Item Is Nothing Then .Add "item", Item.ToJson
        If IsEmpty(DeferCommit) = False Then .Add "deferCommit", DeferCommit
    End With
    Set ToJson = objJson
End Function


